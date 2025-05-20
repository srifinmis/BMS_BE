const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const moment = require('moment');
const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const models = initModels(sequelize);
const { repayment_schedule_temp,repayment_schedule,payment_details } = models;
const jwt = require("jsonwebtoken");

async function exportBulletLoanSchedule(req, res) {
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

     console.log("hellooo "+JSON.stringify(req.body));
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

    const format = input_file_format?.toLowerCase(); // e.g. '.csv'
    const principalAmount = Number(tranche_amount);
    const tenureMonths = Number(tenure_months);
    const annualInterestRate = Number(interest_rate) / 100;
    const drawdownDate = dayjs(tranche_date);
    const daysInYear = Number(interest_calculation_days);

    let totalInterest = 0, totalPrincipal = 0, totalPayment = 0,totalDays=0;

    const headers = [
      'No. of Months', 'Opening Balance', 'From Date', 'Due Date', 'No. of Days',
      'Interest Rate', 'Servicing of Interest', 'Principal Repayment',
      'Total Payment', 'Closing Balance', 'EMI'
    ];

    const schedule = [];
    let currentDate = moment(tranche_date, 'DD-MMM-YY');

    for (let month = 0; month < tenureMonths; month++) {
      const fromDate = moment(currentDate);
      const dueDate = fromDate.date() > 1
        ? moment(fromDate).add(1, 'month').startOf('month')
        : moment(fromDate).add(1, 'month');

      const numDays = dueDate.diff(fromDate, 'day');
      const effectiveRate = (numDays / daysInYear) * annualInterestRate;
      const interest = Math.round(principalAmount * effectiveRate);
      const principalRepayment = month === tenureMonths - 1 ? principalAmount : 0;
      const total = interest + principalRepayment;
      const closingBalance = month === tenureMonths - 1 ? 0 : principalAmount;

      totalInterest += interest;
      totalPrincipal += principalRepayment;
      totalPayment += total;
      totalDays+=numDays;
      schedule.push({
        month: month + 1,
        openingBalance: principalAmount,
        fromDate: fromDate.format('DD-MMM-YY'),
        dueDate: dueDate.format('DD-MMM-YY'),
        numDays,
        interestRate: `${(annualInterestRate / daysInYear * 100).toFixed(2)}%`,
        interest,
        principalRepayment,
        totalPayment: total,
        closingBalance,
        emi: ''
      });

      currentDate = dueDate;
    }

    // Final total row as an object for Excel + CSV
    const totalRow = {
      month: 'Total',
      openingBalance: '',
      fromDate: '',
      dueDate: '',
      numDays: totalDays,
      interestRate: '',
      interest: totalInterest,
      principalRepayment: totalPrincipal,
      totalPayment: totalPayment,
      closingBalance: '',
      emi: ''
    };


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


    // ===================== CSV Output =====================
    if (format === '.csv') {
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', 'attachment; filename=LoanSchedule.csv');
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
      for (const row of schedule) {
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
          row.closingBalance,
          row.emi
        ].join(',') + '\n');
      }

      // Write total row
      res.write([
        totalRow.month,
        totalRow.openingBalance,
        totalRow.fromDate,
        totalRow.dueDate,
        totalRow.numDays,
        totalRow.interestRate,
        totalRow.interest,
        totalRow.principalRepayment,
        totalRow.totalPayment,
        totalRow.closingBalance,
        totalRow.emi
      ].join(',') + '\n');

      return res.end(); // Important
    }

    // ===================== Excel Output =====================
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Loan Schedule');

    // Set up headers with keys
    sheet.columns = [
      { header: 'No. of Months', key: 'month' },
      { header: 'Opening Balance', key: 'openingBalance' },
      { header: 'From Date', key: 'fromDate' },
      { header: 'Due Date', key: 'dueDate' },
      { header: 'No. of Days', key: 'numDays' },
      { header: 'Interest Rate', key: 'interestRate' },
      { header: 'Servicing of Interest', key: 'interest' },
      { header: 'Principal Repayment', key: 'principalRepayment' },
      { header: 'Total Payment', key: 'totalPayment' },
      { header: 'Closing Balance', key: 'closingBalance' },
      { header: 'EMI', key: 'emi' }
    ];
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
    }
    // Add data rows
    schedule.forEach(row => {
      sheet.addRow(row);
    });

    // Add total row
    const totalRowExcel = sheet.addRow(totalRow);
    totalRowExcel.font = { bold: true };

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader('Content-Disposition', 'attachment; filename=LoanSchedule.xlsx');

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Error generating schedule file:', err);
    res.status(500).json({ error: 'Failed to generate file' });
  }
}

module.exports = { exportBulletLoanSchedule };
