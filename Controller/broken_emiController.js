const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const moment = require('moment');
const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const models = initModels(sequelize);
const { repayment_schedule_temp,repayment_schedule,payment_details } = models;
const jwt = require("jsonwebtoken");

// Format amount with commas
function formatAmount(amount) {
    return Math.round(amount).toLocaleString('en-IN');
}

// Generate loan repayment schedule
function generateLoanSchedule({ 
    tranche_amount, 
    interest_rate, 
    interest_calculation_days,
    moratorium_start_date, 
    moratorium_end_date, 
    tenure_months 
}) {
    let schedule = [];
    let openingBalance = tranche_amount;
    // let currentDate = moment(tranche_date, 'DD-MMM-YY');
    let fromDate = moment(moratorium_start_date, 'DD-MMM-YY');
    let dueDate = moment(moratorium_end_date, 'DD-MMM-YY');
    let interestRatePerDay = Number(interest_rate)/interest_calculation_days/100;
    // Moratorium Interest Only
    let noOfDays = dueDate.diff(fromDate, 'day');
    let servicingOfInterest = openingBalance * interestRatePerDay * noOfDays;
    const annualInterestRate = Number(interest_rate) / 100;
    const ratePerAnnum = interest_rate / 100;
    const monthlyRate = ratePerAnnum / 12;

    schedule.push({
        month: 1,
        openingBalance: formatAmount(openingBalance),
        fromDate: fromDate.format('DD-MMM-YY'),
        dueDate: dueDate.format('DD-MMM-YY'),
        noOfDays: noOfDays,
        interestRate: `${(annualInterestRate / interest_calculation_days * 100).toFixed(2)}%`,
        servicingOfInterest: formatAmount(servicingOfInterest),
        principalRepayment: '',
        totalPayment: formatAmount(servicingOfInterest),
        closingBalance: formatAmount(openingBalance),
        emi: ''
    });

    fromDate = dueDate;
    const emiAmount = Math.round(
        (tranche_amount * monthlyRate * Math.pow(1 + monthlyRate, tenure_months)) /
        (Math.pow(1 + monthlyRate, tenure_months) - 1)
      );
    // EMI payments
    for (let month = 2; month <= tenure_months; month++) {
        let nextDueDate = moment(fromDate).add(1, 'month');
        let days = nextDueDate.diff(fromDate, 'day');

        let interest = openingBalance * interestRatePerDay * days;
        let principalRepayment = emiAmount - interest;

        if (principalRepayment > openingBalance) {
            principalRepayment = openingBalance;
        }

        let closingBalance = openingBalance - principalRepayment;
        if (closingBalance < 0) closingBalance = 0;

        schedule.push({
            month: month,
            openingBalance: formatAmount(openingBalance),
            fromDate: fromDate.format('DD-MMM-YY'),
            dueDate: nextDueDate.format('DD-MMM-YY'),
            noOfDays: days,
            interestRate: (interestRatePerDay * 100).toFixed(2) + '%',
            servicingOfInterest: formatAmount(interest),
            principalRepayment: formatAmount(principalRepayment),
            totalPayment: formatAmount(emiAmount),
            closingBalance: formatAmount(closingBalance),
            emi: formatAmount(emiAmount)
        });

        openingBalance = closingBalance;
        fromDate = nextDueDate;
    }

    return schedule;
}


exports.generateBrokenEMI = async (req, res) => {
    try {
      const {
        tranche_amount,
        interest_rate,
        moratorium_start_date,
        moratorium_end_date,
        tenure_months,
        interest_calculation_days,
        interest_payment_frequency,
        principal_payment_frequency,
        input_file_format,
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
    
  
      // Generate schedule
      const schedule = generateLoanSchedule({
        tranche_amount,
        interest_rate,
        moratorium_start_date,
        moratorium_end_date,
        interest_calculation_days,
        tenure_months
      });
  
      const createdAt = new Date();
      console.log("moritoriam dates: ",schedule.moratorium_start_date," -- ",schedule.moratorium_end_date)
      // Prepare repayment rows for DB
      const repaymentRows = schedule.map(row => ({
        sanction_id: sanction_id,
        tranche_id: tranche_id,
        due_date: row.dueDate,
        principal_due: parseFloat(row.principalRepayment.replace(/,/g, '')) || 0,
        interest_due: parseFloat(row.servicingOfInterest.replace(/,/g, '')) || 0,
        total_due: parseFloat(row.totalPayment.replace(/,/g, '')) || 0,
        createdat: createdAt,
        lender_code: lender_code,
        opening_balance: parseFloat(row.openingBalance.replace(/,/g, '')) || 0,
        closing_balance: parseFloat(row.closingBalance.replace(/,/g, '')) || 0,
        from_date: row.fromDate,
        interest_days: row.noOfDays,
        interest_rate: parseFloat(row.interestRate.replace('%', '')) || 0,
        emi_sequence: row.month,
        repayment_type: loan_type,
        created_by:decoded.id,
        moratorium_start_date: moratorium_start_date,
        moratorium_end_date: moratorium_end_date
      }));
  
      console.log('Repayment Rows:\n', JSON.stringify(repaymentRows, null, 2));
  
      // Insert into DB
      await repayment_schedule_temp.bulkCreate(repaymentRows);
  
      // Headers for file
      const headers = [
        'No. of Months',
        'Opening Balance',
        'From Date',
        'Due Date',
        'No. of Days',
        'Interest Rate',
        'Servicing of Interest',
        'Principal Repayment',
        'Total Payment',
        'Closing Balance',
        'EMI'
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
        'Total', '', '', '',
        totalDays,
        '',
        Math.round(totalInterest),
        Math.round(totalPrincipalRepayment),
        Math.round(totalPayment),
        '', ''
      ];
  
      if (input_file_format === '.csv') {
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
            row.openingBalance.replace(/,/g, ''),  // Remove commas for numeric fields
            row.fromDate,
            row.dueDate,
            row.noOfDays,
            row.interestRate,
            row.servicingOfInterest.replace(/,/g, ''),
            row.principalRepayment.replace(/,/g, ''),  // Remove commas for numeric fields
            row.totalPayment.replace(/,/g, ''),  // Remove commas for numeric fields
            row.closingBalance.replace(/,/g, ''),
            row.emi.replace(/,/g, '') || ''
          ].join(',') + '\n');
        }
  
        // Write the totals row to the CSV
        res.write(totalsRow.join(',') + '\n');
        return res.end();
      } else {
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Loan Schedule');
  
        const headerRow = sheet.addRow(headers);
        headerRow.font = { bold: true };
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
          sheet.addRow([
            row.month,
            row.openingBalance,
            row.fromDate,
            row.dueDate,
            row.noOfDays,
            row.interestRate,
            row.servicingOfInterest,
            row.principalRepayment || '',
            row.totalPayment,
            row.closingBalance,
            row.emi || ''
          ]);
        });
  
        const totalRow = sheet.addRow(totalsRow);
        totalRow.font = { bold: true };
  
        res.setHeader(
          'Content-Type',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader('Content-Disposition', 'attachment; filename=LoanSchedule.xlsx');
  
        await workbook.xlsx.write(res);
        res.end();
      }
  
    } catch (err) {
      console.error('Error generating schedule file:', err);
      res.status(500).json({ error: 'Failed to generate schedule file' });
    }
  };
  

