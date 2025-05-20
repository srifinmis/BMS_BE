const moment = require('moment');
const jwt = require('jsonwebtoken');
const { writeToStream } = require('fast-csv');
const ExcelJS = require('exceljs');
const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const models = initModels(sequelize);
const { payment_details, tranche_details, repayment_schedule } = models;

exports.getPaymentDetails = async (req, res) => {
  const { lender_code, sanction_id, tranche_id } = req.body;

  if (!lender_code || !sanction_id || !tranche_id) {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  try {
    // Get all payments for the tranche
    // Step 1: Get due_dates from repayment_schedule
    const schedules = await repayment_schedule.findAll({
      attributes: ['due_date'],
      where: {
        lender_code,
        sanction_id,
        tranche_id
      }
    });

    const dueDates = schedules.map(s => s.due_date);

    // Step 2: Use those due_dates to filter payment_details
    const results = await payment_details.findAll({
      where: {
        lender_code,
        sanction_id,
        tranche_id,
        due_date: dueDates  // Sequelize will convert this to a WHERE IN clause
      }
    });


    // Get tranche amount (if any)
    const tranche = await tranche_details.findOne({
      where: { lender_code, sanction_id, tranche_id },
      attributes: ['tranche_amount']
    });

    if (!tranche || tranche.tranche_amount == null) {
      return res.status(404).json({ error: 'No tranche found or tranche amount missing' });
    }

    // Safely calculate total paid
    const totalPaid = results.reduce((sum, row) => {
      const amt = parseFloat(row.pricipal_coll);
      return sum + (isNaN(amt) ? 0 : amt);
    }, 0);
    console.log("tranche before subtract: ", tranche.tranche_amount)
    console.log("totalPaid: ", totalPaid)

    const trancheAmount = parseFloat(tranche.tranche_amount);
    const outstanding = isNaN(trancheAmount) ? 0 : Math.max(trancheAmount - totalPaid, 0);
    console.log("Outstanding: ", outstanding)


    // Fetch next repayment based on number of payments already made
    const offsetIndex = Math.max(0, results.length);
    const adjustedOffset = offsetIndex > 0 ? offsetIndex - 1 : 0;
    const nextRepayment = await repayment_schedule.findOne({
      where: { lender_code, sanction_id, tranche_id },
      order: [['from_date', 'ASC']],
      offset: adjustedOffset,
      limit: 1
    });

    if (!nextRepayment || !nextRepayment.due_date) {
      return res.status(404).json({ error: 'No further repayment schedule found' });
    }
    const revisedTrancheDate = (results.length === 0)
      ? nextRepayment.from_date
      : nextRepayment.due_date;

    return res.status(200).json({
      count: results.length,
      outstanding_amount: outstanding.toFixed(2),
      revised_tranche_date: revisedTrancheDate,
      data: results
    });

  } catch (err) {
    console.error('Error fetching payment details:', err);
    return res.status(500).json({ error: 'Internal Server Error' });
  }
};


exports.getTenureMonths = async (req, res) => {
  const { lender_code, sanction_id, tranche_id } = req.body;

  if (!lender_code || !sanction_id || !tranche_id) {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  try {
    const result = await tranche_details.findOne({
      where: {
        lender_code,
        sanction_id,
        tranche_id
      },
      attributes: ['tenure_months', 'interest_rate']
    });

    if (!result) {
      return res.status(404).json({ error: 'Record not found' });
    }

    return res.status(200).json({
      tenure_months: result.tenure_months,
      interest_rate: result.interest_rate
    });
  } catch (err) {
    console.error('Error fetching tenure_months:', err);
    return res.status(500).json({ error: 'Internal Server Error' });
  }
};