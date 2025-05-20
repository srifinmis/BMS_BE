const moment = require('moment');
const { writeToStream } = require('fast-csv');
const ExcelJS = require('exceljs');
const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const models = initModels(sequelize);
const { repayment_schedule_temp,repayment_schedule,payment_details } = models;
const jwt = require("jsonwebtoken");

function cleanNumber(value) {
  return Number(String(value || '0').replace(/,/g, ''));
}
exports.generateLoanReport = async (req, res) => {
  try {
    const {tranche_amount,
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
    const format = req.query.format || 'csv';
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
    if (!tranche_amount || !tenure_months || !interest_rate || !tranche_date) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const schedule = [];
    let openingBalance = parseFloat(tranche_amount);
    const totalQuarterlyInstallments = Math.ceil(tenure_months / 3);
    const quarterlyPrincipal = +(openingBalance / totalQuarterlyInstallments).toFixed(0);
    const ratePerDay = interest_rate / 100 / 365;
    let currentDate = moment(tranche_date, 'DD-MMM-YY');

    for (let month = 1; month <= tenure_months; month++) {
      const fromDate = moment(currentDate);
      const dueDate = fromDate.date() > 1
        ? moment(fromDate).add(1, 'month').startOf('month')
        : moment(fromDate).add(1, 'month');

      const noOfDays = dueDate.diff(fromDate, 'days');
      const interest = +(openingBalance * ratePerDay * noOfDays).toFixed(0);

      let principalRepayment = 0;
      if (month % 3 === 0 && openingBalance > 0) {
        principalRepayment = Math.min(openingBalance, quarterlyPrincipal);
      }

      const totalPayment = interest + principalRepayment;
      const closingBalance = +(openingBalance - principalRepayment).toFixed(0);

      schedule.push({
        monthNumber: month,
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
      currentDate = dueDate;

      if (openingBalance <= 0) break;
    }
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
  opening_balance: cleanNumber(row.openingBalance),
  closing_balance: cleanNumber(row.closingBalance),
  from_date: row.fromDate,
  interest_days: row.numDays,
  interest_rate: parseFloat(row.interestRate.replace('%', '')),
  emi_sequence: row.monthNumber,
  repayment_type: loan_type,
  created_by:decoded.id
}));
console.log('Repayment Rows:\n', JSON.stringify(repaymentRows, null, 2));
// Insert into DB using Sequelize
await repayment_schedule_temp.bulkCreate(repaymentRows);

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

    if (input_file_format === '.xlsx') {
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

      // Add totals row
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
    
      // Combine previous and new schedule
      let combinedSchedule = flag ? [...previousScheduleRows, ...schedule] : [...schedule];
    
      // Append totals row as a plain object with matching structure
      combinedSchedule.push({
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
      });
    
      return writeToStream(res, combinedSchedule, {
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


        // new Date(row.fromDate).toLocaleDateString('en-GB', {
        //   day: '2-digit',
        //   month: 'short',
        //   year: '2-digit'
        // }).replace(/ /g, '-'),
        // new Date(row.dueDate).toLocaleDateString('en-GB', {
        //   day: '2-digit',
        //   month: 'short',
        //   year: '2-digit'
        // }).replace(/ /g, '-'),
        writeHeaders: true,
        transform: (row) => ({
          'No. of Months': row.month || row.monthNumber || '',
          'Opening Balance': formatNumber(row.openingBalance || ''),
          'From Date': new Date(row.fromDate).toLocaleDateString('en-GB', {
            day: '2-digit',
            month: 'short',
            year: '2-digit'
          }).replace(/ /g, '-') || '',
          'Due Date': new Date(row.dueDate).toLocaleDateString('en-GB', {
            day: '2-digit',
            month: 'short',
            year: '2-digit'
          }).replace(/ /g, '-') || '',
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
    return res.status(500).json({ error: 'Internal server error' });
  }
};

// Format numbers with commas
function formatNumber(number) {
  if (typeof number === 'number') {
    return number.toLocaleString();
  }
  return number || '';
}
