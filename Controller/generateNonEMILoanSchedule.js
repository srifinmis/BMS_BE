const ExcelJS = require('exceljs');
const moment = require('moment');
const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const models = initModels(sequelize);
const { repayment_schedule_temp,repayment_schedule,payment_details } = models;
const jwt = require("jsonwebtoken");

function generateReducingBalanceScheduleWithoutEMI({
  tranche_amount,
  tenure_months,
  interest_rate,
  tranche_date,
  input_file_format,
  interest_calculation_days,
  principal_payment_frequency,
  interest_payment_frequency
}) {
  const interestDaysBase = Array.isArray(interest_calculation_days)
  ? interest_calculation_days.includes("365") || interest_calculation_days.includes(365)
    ? 365
    : 360
  : interest_calculation_days === 365
    ? 365
    : 360;

  const schedule = [];

  const fixedPrincipal = Math.round(tranche_amount / tenure_months);
  let openingBalance = tranche_amount;
  let currentDate = moment(tranche_date, "DD-MMM-YY");
  const ratePerAnnum = interest_rate / 100;

  for (let i = 0; i < tenure_months; i++) {
    const fromDate = moment(currentDate);
    let dueDate;
        if (fromDate.date() > 1) {
            dueDate = moment(fromDate).add(1, 'month').startOf('month');
        } else {
            // Otherwise, just add one month normally
            dueDate = moment(fromDate).add(1, 'month');
        }
    const numDays = dueDate.diff(fromDate, 'days');
    const effectiveRate = (numDays / interestDaysBase) * ratePerAnnum;

    const interest = Math.round(openingBalance * effectiveRate);
    const principalRepayment = fixedPrincipal;
    const totalPayment = interest + principalRepayment;

    const closingBalance = openingBalance - principalRepayment;

    schedule.push({
      month: i + 1,
      openingBalance: Math.round(openingBalance),
      fromDate: fromDate.format("DD-MMM-YY"),
      dueDate: dueDate.format("DD-MMM-YY"),
      numDays,
      interestRate: (ratePerAnnum/365 * 100).toFixed(2) + '%',
      interest,
      principalRepayment,
      totalPayment,
      closingBalance: Math.round(closingBalance),
    });

    openingBalance = closingBalance;
    currentDate = dueDate;
  }

  return { schedule };
}

exports.generateNonEMILoanScheduleFile = async (req, res) => {
  try {
    const {
      tranche_amount,
      tenure_months,
      interest_rate,
      tranche_date,
      input_file_format,
      interest_calculation_days,
      principal_payment_frequency,
      interest_payment_frequency,
      sanction_id,
      tranche_id,
      lender_code,
      loan_type,
      created_by,
      flag
    } = req.body;
  const JWT_SECRET = process.env.JWT_SECRET;
  
      if (!created_by) {
            return res.status(401).json({ message: "JWT token not provided" });
          }
      
          let decoded;
          try {
            decoded = jwt.verify(created_by, JWT_SECRET);
          } catch (error) {
            return res.status(401).json({ message: "Invalid token", error: error.message });
          }
      console.log("decoded "+JSON.stringify(decoded));
      await repayment_schedule_temp.destroy({
        where: {
          // tranche_id: tranche_id,
          // repayment_type: loan_type,
          created_by: decoded.id
        }
      });
    const { schedule } = generateReducingBalanceScheduleWithoutEMI({
      tranche_amount,
      tenure_months,
      interest_rate,
      tranche_date,
      input_file_format,
      interest_calculation_days,
      principal_payment_frequency,
      interest_payment_frequency
    });
    const createdAt = new Date(); // current time
    // Prepare data for bulk insert
const repaymentRows = schedule.map(row => ({
  sanction_id: sanction_id,
  tranche_id: tranche_id,
  due_date: row.dueDate,
  principal_due: row.principalRepayment,
  interest_due: row.interest,
  total_due: row.totalPayment,
  createdat: createdAt,
  lender_code: lender_code,
  opening_balance: row.openingBalance,
  closing_balance: row.closingBalance,
  from_date: row.fromDate,
  interest_days: row.numDays,
  interest_rate: parseFloat(row.interestRate.replace('%', '')),
  emi_sequence: row.month,
  repayment_type: loan_type,
  created_by:decoded.id
}));
console.log('Repayment Rows:\n', JSON.stringify(repaymentRows, null, 2));
// Insert into DB using Sequelize
await repayment_schedule_temp.bulkCreate(repaymentRows);


    const headers = [
      'No. of Months',
      'Opening Balance',
      'From Date',
      'Due Date',
      'No. of Days',
      'Interest Rate',
      'Interest Amount',
      'Principal Repayment',
      'Total Payment',
      'Closing Balance',
    ];

    let previousScheduleRows = [];
    if (flag) {
      const paymentDetails = await payment_details.findAll({
        where: { lender_code, sanction_id, tranche_id }
      });
    
      for (const payment of paymentDetails) {
        const existingRows = await repayment_schedule.findAll({
          where: {
            lender_code: payment.lender_code,
            sanction_id: payment.sanction_id,
            tranche_id: payment.tranche_id,
            due_date: payment.due_date
          }
        });
    
        previousScheduleRows.push(...existingRows.map(row => ({
          month: row.emi_sequence,
          openingBalance: row.opening_balance,
          fromDate: row.from_date,
          dueDate: row.due_date,
          numDays: row.interest_days,
          interestRate: `${row.interest_rate}%`,
          interest: row.interest_due,
          principalRepayment: row.principal_due,
          totalPayment: row.total_due,
          closingBalance: row.closing_balance,
          emi: '' // if not present, leave empty
        })));
      }
    }
    
    let totalDays, totalInterest, totalPrincipalRepayment, totalPayment;

    if (flag) {
      const fullSchedule = [...previousScheduleRows, ...schedule];
      totalDays = fullSchedule.reduce((sum, row) => sum + (Number(row.numDays) || 0), 0);
      totalInterest = fullSchedule.reduce((sum, row) => sum + (Number(row.interest) || 0), 0);
      totalPrincipalRepayment = fullSchedule.reduce((sum, row) => sum + (Number(row.principalRepayment) || 0), 0);
      totalPayment = fullSchedule.reduce((sum, row) => sum + (Number(row.totalPayment) || 0), 0);
    } else {
      totalDays = schedule.reduce((sum, row) => sum + row.numDays, 0);
      totalInterest = schedule.reduce((sum, row) => sum + row.interest, 0);
      totalPrincipalRepayment = schedule.reduce((sum, row) => sum + row.principalRepayment, 0);
      totalPayment = schedule.reduce((sum, row) => sum + row.totalPayment, 0);
    }
    const totalsRow = [
      'Total', '', '', '', totalDays, '', totalInterest,
      totalPrincipalRepayment, totalPayment, ''
    ];

    if (input_file_format === '.csv') {
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', 'attachment; filename=ReducingBalanceSchedule.csv');
      res.write(headers.join(',') + '\n');
      if(flag){
        // Write previous schedule rows
       for (const row of previousScheduleRows) {
         res.write([
           row.month,
           row.openingBalance,
           new Date(row.fromDate).toLocaleDateString('en-GB', {
            day: '2-digit',
            month: 'short',
            year: '2-digit'
          }).replace(/ /g, '-'),
          new Date(row.dueDate).toLocaleDateString('en-GB', {
            day: '2-digit',
            month: 'short',
            year: '2-digit'
          }).replace(/ /g, '-'),
           row.numDays,
           row.interestRate,
           row.interest,
           row.principalRepayment,
           row.totalPayment,
           row.closingBalance,
           row.emi,
         ].join(',') + '\n');
       }
     }
     // Determine the starting month
     let startMonth = 1;
     if (flag && previousScheduleRows.length > 0) {
       const lastPreviousRow = previousScheduleRows[previousScheduleRows.length - 1];
       startMonth = lastPreviousRow.month + 1;
     }
      schedule.forEach((row,i) => {
         row = schedule[i];
        row.month = startMonth+i;
        res.write([
          row.month,
          row.openingBalance,
          row.fromDate,
          row.dueDate,
          row.numDays,
          row.interestRate,
          row.interest,
          row.principalRepayment,
          row.totalPayment,
          row.closingBalance
        ].join(',') + '\n');
      });

      res.write(totalsRow.join(',') + '\n');
      res.end();
    } else {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Loan Schedule - Reducing');

      const headerRow = sheet.addRow(headers);
      headerRow.font = { bold: true };
      let startMonth = 1;
      if(flag){
        // sheet.addRow(['Previous Schedule']).font = { bold: true };
          // Previous rows
        previousScheduleRows.forEach(row => {
          sheet.addRow([
            row.month,
            row.openingBalance,
            new Date(row.fromDate).toLocaleDateString('en-GB', {
              day: '2-digit',
              month: 'short',
              year: '2-digit'
            }).replace(/ /g, '-'),
            new Date(row.dueDate).toLocaleDateString('en-GB', {
              day: '2-digit',
              month: 'short',
              year: '2-digit'
            }).replace(/ /g, '-'),
            row.numDays,
            row.interestRate,
            row.interest,
            row.principalRepayment,
            row.totalPayment,
            row.closingBalance,
            row.emi
          ]);
        });
        startMonth = previousScheduleRows[previousScheduleRows.length-1].month+1;
      }
      
      schedule.forEach((row,index) => {
        row.month = startMonth+index;
        sheet.addRow([
          row.month,
          row.openingBalance,
          row.fromDate,
          row.dueDate,
          row.numDays,
          row.interestRate,
          row.interest,
          row.principalRepayment,
          row.totalPayment,
          row.closingBalance
        ]);
      });

      const totalRow = sheet.addRow(totalsRow);
      totalRow.font = { bold: true };

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=ReducingBalanceSchedule.xlsx');

      await workbook.xlsx.write(res);
      res.end();
    }
  } catch (err) {
    console.error('Error generating schedule:', err);
    res.status(500).json({ error: 'Failed to generate file' });
  }
};
