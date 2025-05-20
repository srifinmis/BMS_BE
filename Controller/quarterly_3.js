const moment = require('moment');
const jwt = require('jsonwebtoken');
const { writeToStream } = require('fast-csv');
const ExcelJS = require('exceljs');
const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const models = initModels(sequelize);
const { repayment_schedule_temp } = models;

exports.generatequarterly_3 = async (req, res) => {
  try {
    const {
      tranche_amount,
      tenure_months,
      interest_rate,
      moratorium_start_date,
      moratorium_end_date,
      input_file_format,
      interest_calculation_days,
      principal_payment_frequency,
      interest_payment_frequency,
      sanction_id,
      tranche_id,
      lender_code,
      loan_type,
      created_by
    } = req.body;

    const format = req.query.format || 'csv';

    if (!tranche_amount || !tenure_months || !interest_rate || !moratorium_start_date || !moratorium_end_date) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    // const JWT_SECRET = process.env.JWT_SECRET;

    // if (!created_by) {
    //   return res.status(401).json({ message: "JWT token not provided" });
    // }

    // let decoded;
    // try {
    //   decoded = jwt.verify(created_by, JWT_SECRET);
    // } catch (error) {
    //   return res.status(401).json({ message: "Invalid token", error: error.message });
    // }

    // await repayment_schedule_temp.destroy({
    //   where: {
    //     created_by: decoded.id
    //   }
    // });

    // --- SCHEDULE GENERATION LOGIC ---
    const schedule = [];
    let openingBalance = parseFloat(tranche_amount);
    const ratePerDay = interest_rate / 100 / interest_calculation_days;
    const fromDateStart = moment(moratorium_start_date, 'DD-MMM-YY');
    const toDateEnd = moment(moratorium_end_date, 'DD-MMM-YY');
    const months = tenure_months;
    const principalPerQuarter = openingBalance / Math.ceil(months / 3);

    for (let emiCount = 1; emiCount <= months; emiCount++) {
      const fromDate = moment(fromDateStart).add(emiCount - 1, 'months');
      const dueDate = moment(fromDateStart).add(emiCount, 'months').subtract(1, 'day');
      const noOfDays = dueDate.diff(fromDate, 'days') + 1;

      const interest = parseFloat((openingBalance * ratePerDay * noOfDays).toFixed(2));
      const principalRepayment = emiCount % 3 === 0 ? parseFloat(principalPerQuarter.toFixed(2)) : 0;
      const totalPayment = interest + principalRepayment;
      const closingBalance = parseFloat((openingBalance - principalRepayment).toFixed(2));

      schedule.push({
        monthNumber: emiCount,
        openingBalance,
        fromDate: fromDate.format('DD-MMM-YY'),
        dueDate: dueDate.format('DD-MMM-YY'),
        numDays: noOfDays,
        interestRate: (ratePerDay * 100).toFixed(2) + '%',
        interest,
        principalRepayment,
        totalPayment,
        closingBalance,
        emi: ''
      });

      openingBalance = closingBalance;
    }
    console.log("Generated Repayment Schedule:\n", schedule);


    // --- DB INSERTION ---
    const createdAt = new Date();
    const repaymentRows = schedule.map(row => ({
      sanction_id,
      tranche_id,
      due_date: row.dueDate,
      principal_due: row.principalRepayment,
      interest_due: row.interest,
      total_due: row.totalPayment,
      createdat: createdAt,
      lender_code,
      opening_balance: cleanNumber(row.openingBalance),
      closing_balance: cleanNumber(row.closingBalance),
      from_date: row.fromDate,
      interest_days: row.numDays,
      interest_rate: parseFloat(row.interestRate.replace('%', '')),
      emi_sequence: row.monthNumber,
      repayment_type: loan_type
      // created_by: decoded.id
    }));

    await repayment_schedule_temp.bulkCreate(repaymentRows);

    // --- TOTALS ---
    const totalInterest = schedule.reduce((sum, row) => sum + row.interest, 0);
    const totalPrincipalRepayment = schedule.reduce((sum, row) => sum + row.principalRepayment, 0);
    const totalPayment = schedule.reduce((sum, row) => sum + row.totalPayment, 0);
    const totalDays = schedule.reduce((sum, row) => sum + row.numDays, 0);

    const totalsRow = {
      monthNumber: 'Total',
      openingBalance: '',
      fromDate: '',
      dueDate: '',
      numDays: totalDays,
      interestRate: '',
      interest: totalInterest,
      principalRepayment: totalPrincipalRepayment,
      totalPayment: totalPayment,
      closingBalance: ''
    };

    // --- EXPORT ---
    if (input_file_format === 'excel') {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Loan Schedule');

      sheet.columns = [
        { header: 'No. of Months', key: 'monthNumber' },
        { header: 'Opening Balance', key: 'openingBalance' },
        { header: 'From Date', key: 'fromDate' },
        { header: 'Due Date', key: 'dueDate' },
        { header: 'No. of Days', key: 'numDays' },
        { header: 'Interest Rate', key: 'interestRate' },
        { header: 'Servicing of Interest', key: 'interest' },
        { header: 'Principal Repayment', key: 'principalRepayment' },
        { header: 'Total Payment', key: 'totalPayment' },
        { header: 'Closing Balance', key: 'closingBalance' },
      ];

      schedule.forEach(row => {
        sheet.addRow({
          ...row,
          openingBalance: formatNumber(row.openingBalance),
          interest: formatNumber(row.interest),
          principalRepayment: formatNumber(row.principalRepayment),
          totalPayment: formatNumber(row.totalPayment),
          closingBalance: formatNumber(row.closingBalance),
        });
      });

      sheet.addRow({
        ...totalsRow,
        interest: formatNumber(totalsRow.interest),
        principalRepayment: formatNumber(totalsRow.principalRepayment),
        totalPayment: formatNumber(totalsRow.totalPayment),
      });

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=loan_schedule.xlsx');
      await workbook.xlsx.write(res);
      return res.end();
    } else {
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', 'attachment; filename=loan_schedule.csv');

      return writeToStream(res, [...schedule, totalsRow], {
        headers: [
          'No. of Months',
          'Opening Balance',
          'From Date',
          'Due Date',
          'No. of Days',
          'Interest Rate',
          'Servicing of Interest',
          'Principal Repayment',
          'Total Payment',
          'Closing Balance'
        ],
        writeHeaders: true,
        transform: (row) => ({
          'No. of Months': row.monthNumber,
          'Opening Balance': formatNumber(row.openingBalance || ''),
          'From Date': row.fromDate || '',
          'Due Date': row.dueDate || '',
          'No. of Days': row.numDays || '',
          'Interest Rate': row.interestRate || '',
          'Servicing of Interest': formatNumber(row.interest || 0),
          'Principal Repayment': formatNumber(row.principalRepayment || 0),
          'Total Payment': formatNumber(row.totalPayment || 0),
          'Closing Balance': formatNumber(row.closingBalance || ''),
        })
      });
    }
  } catch (error) {
    console.error('Loan report error:', error);
    return res.status(500).json({ error: 'Internal server error', details: error.message });
  }
};

// Helper functions
function formatNumber(number) {
  if (typeof number === 'number') {
    return number.toLocaleString();
  }
  return number || '';
}

function cleanNumber(number) {
  return Number((number || 0).toFixed(2));
}
