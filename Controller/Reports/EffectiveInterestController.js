const express = require('express');
const ExcelJS = require('exceljs');
const moment = require('moment');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType } = require('docx');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

require('dotenv').config();
const models = initModels(sequelize);
const { lender_master, payment_details, repayment_schedule, tranche_details, sanction_details } = models;

exports.generateEffectiveInterestofReport = async (req, res) => {

    // try {
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest');

    //     const loanTypes = new Set();
    //     const grouped = {};

    //     const sanctionDates = sanctions.map(entry => moment(entry.sanction_date).startOf('month'));
    //     const firstMonth = moment.min(sanctionDates);
    //     const lastMonth = moment().startOf('month');

    //     const allMonths = [];
    //     let cursor = firstMonth.clone();
    //     while (cursor.isSameOrBefore(lastMonth)) {
    //         allMonths.push(cursor.format('MMM-YYYY'));
    //         cursor.add(1, 'month');
    //     }
    //     allMonths.reverse();

    //     // Initialize grouped structure
    //     allMonths.forEach(month => {
    //         grouped[month] = {};
    //         sanctions.forEach(entry => {
    //             const loanType = entry.loan_type || 'N/A';
    //             loanTypes.add(loanType);
    //             grouped[month][loanType] = {
    //                 total_amount: 0,
    //                 total_weighted_interest: 0,
    //                 total_weighted_proc_fee: 0,
    //                 total_payment: 0,
    //                 interest_numerator: 0,
    //                 interest_denominator: 0,
    //                 total_tranche_amount: 0, // Added to calculate avg processing fee
    //                 total_processing_fee_annually: 0, // Added for fee calculations
    //             };
    //         });
    //     });

    //     // Populate grouped structure
    //     allMonths.forEach(month => {
    //         const monthMoment = moment(month, 'MMM-YYYY');

    //         // Sanctions
    //         sanctions.forEach(entry => {
    //             const sanctionMonth = moment(entry.sanction_date).startOf('month');
    //             if (sanctionMonth.isAfter(monthMoment)) return;

    //             const loanType = entry.loan_type || 'N/A';
    //             const data = grouped[month][loanType];
    //             const sanctionAmount = parseFloat(entry.sanction_amount || 0);
    //             const interestRate = parseFloat(entry.interest_rate || 0);
    //             const procFeeRate = parseFloat(entry.processing_fee_rate || 0);
    //             const tenureMonths = parseFloat(entry.tenure_months || 0);

    //             data.total_amount += sanctionAmount;
    //             data.total_weighted_interest += interestRate * sanctionAmount;
    //             data.total_weighted_proc_fee += procFeeRate * sanctionAmount;

    //             // Calculate annual processing fee
    //             const processingFeeAnnually = (sanctionAmount * procFeeRate) / tenureMonths * 12;
    //             data.total_processing_fee_annually += processingFeeAnnually;

    //             // Track total tranche amount for avg processing fee calculations
    //             data.total_tranche_amount += sanctionAmount;
    //         });

    //         // Payments
    //         payments.forEach(payment => {
    //             const paymentDate = moment(payment.payment_date).startOf('month');
    //             if (paymentDate.isAfter(monthMoment)) return;

    //             const sanction = sanctions.find(s => s.sanction_id === payment.sanction_id);
    //             if (!sanction) return;

    //             const loanType = sanction.loan_type || 'N/A';
    //             const paymentAmount = parseFloat(payment.payment_amount || 0);
    //             const data = grouped[month][loanType];

    //             data.total_payment += paymentAmount;

    //             const tranche = tranches.find(t =>
    //                 t.sanction_id === payment.sanction_id &&
    //                 t.tranche_id === payment.tranche_id
    //             );

    //             if (tranche) {
    //                 const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                 const diff = trancheAmount - paymentAmount;

    //                 if (diff > 0) {
    //                     const interestRate = parseFloat(tranche.interest_rate || 0) / 100;
    //                     data.interest_numerator += interestRate * diff;
    //                     data.interest_denominator += diff;
    //                 }
    //             }
    //         });
    //     });

    //     const sortedLoanTypes = Array.from(loanTypes);
    //     const applyFullBorder = cell => {
    //         cell.border = {
    //             top: { style: 'thin' },
    //             left: { style: 'thin' },
    //             bottom: { style: 'thin' },
    //             right: { style: 'thin' }
    //         };
    //     };

    //     // Header Row
    //     sheet.getCell('A3').value = 'As on';
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getCell('A3').font = { bold: true, size: 12 };
    //     applyFullBorder(sheet.getCell('A3'));

    //     let startCol = 2;
    //     allMonths.forEach(month => {
    //         sheet.mergeCells(3, startCol, 3, startCol + 3);
    //         const cell = sheet.getRow(3).getCell(startCol);
    //         cell.value = month;
    //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //         cell.font = { bold: true, size: 12 };

    //         for (let i = 0; i < 4; i++) {
    //             applyFullBorder(sheet.getRow(3).getCell(startCol + i));
    //         }

    //         startCol += 4;
    //     });

    //     // Sub-Headers
    //     const subHeaderRow = sheet.getRow(4);
    //     subHeaderRow.height = 30;
    //     subHeaderRow.getCell(1).value = 'Loan Type';
    //     subHeaderRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    //     subHeaderRow.getCell(1).font = { bold: true, size: 12 };
    //     applyFullBorder(subHeaderRow.getCell(1));

    //     let subCol = 2;
    //     allMonths.forEach(() => {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = subHeaderRow.getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true, size: 12 };
    //             applyFullBorder(cell);
    //         });
    //     });

    //     const formatRate = val => `${parseFloat(val).toFixed(2)}%`;

    //     // Data Rows
    //     sortedLoanTypes.forEach(loanType => {
    //         const row = [loanType];
    //         allMonths.forEach(month => {
    //             const data = grouped[month]?.[loanType];
    //             if (data && data.total_amount > 0) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const weightedInterest = data.interest_denominator > 0
    //                     ? data.interest_numerator / data.interest_denominator
    //                     : 0;
    //                 const weightedProcFee = data.total_weighted_proc_fee / data.total_amount;
    //                 const totalRate = weightedInterest + weightedProcFee;

    //                 // Calculate Avg. Processing Fee Rate
    //                 const avgProcessingFeeRate = data.total_tranche_amount > 0
    //                     ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100
    //                     : 0;

    //                 row.push(
    //                     outstanding.toFixed(2),
    //                     formatRate(weightedInterest),
    //                     formatRate(avgProcessingFeeRate),
    //                     formatRate(totalRate)
    //                 );
    //             } else {
    //                 row.push('', '', '', '');
    //             }
    //         });

    //         const dataRow = sheet.addRow(row);
    //         dataRow.eachCell(cell => {
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             applyFullBorder(cell);
    //         });
    //     });

    //     // Total Row
    //     const totalRow = ['Total'];
    //     allMonths.forEach(month => {
    //         let totalAmount = 0, totalPayment = 0, interestNumerator = 0, interestDenominator = 0, totalProcFee = 0, totalProcessingFeeAnnually = 0;

    //         sortedLoanTypes.forEach(loanType => {
    //             const data = grouped[month]?.[loanType];
    //             if (data && data.total_amount > 0) {
    //                 totalAmount += data.total_amount;
    //                 totalPayment += data.total_payment;
    //                 interestNumerator += data.interest_numerator;
    //                 interestDenominator += data.interest_denominator;
    //                 totalProcFee += data.total_weighted_proc_fee;
    //                 totalProcessingFeeAnnually += data.total_processing_fee_annually;
    //             }
    //         });

    //         if (totalAmount > 0) {
    //             const outstanding = (totalAmount - totalPayment) / 1e7;
    //             const weightedInterest = interestDenominator > 0 ? interestNumerator / interestDenominator : 0;
    //             const weightedProcFee = totalProcFee / totalAmount;
    //             const totalRate = weightedInterest + weightedProcFee;

    //             // Calculate Avg. Processing Fee Rate
    //             const avgProcessingFeeRate = totalAmount > 0
    //                 ? (totalProcessingFeeAnnually / totalAmount) * 100
    //                 : 0;

    //             totalRow.push(
    //                 outstanding.toFixed(2),
    //                 formatRate(weightedInterest),
    //                 formatRate(avgProcessingFeeRate),
    //                 formatRate(totalRate)
    //             );
    //         } else {
    //             totalRow.push('', '', '', '');
    //         }
    //     });

    //     const tRow = sheet.addRow(totalRow);
    //     tRow.eachCell(cell => {
    //         cell.font = { bold: true };
    //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyFullBorder(cell);
    //     });

    //     // Auto-size columns
    //     sheet.columns.forEach(col => {
    //         let maxLength = 10;
    //         col.eachCell({ includeEmpty: true }, cell => {
    //             const val = cell.value ? cell.value.toString() : '';
    //             maxLength = Math.max(maxLength, val.length);
    //         });
    //         col.width = maxLength + 2;
    //     });

    //     sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=Effective_Interest_Report.xlsx');
    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error('Error generating Excel:', error);
    //     res.status(500).send('Failed to generate summary report');
    // }

    //24 apr

    // try {
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest');

    //     const loanTypes = new Set();
    //     const grouped = {};

    //     const sanctionDates = sanctions.map(entry => moment(entry.sanction_date).startOf('month'));
    //     const firstMonth = moment.min(sanctionDates);
    //     const lastMonth = moment().startOf('month');

    //     const allMonths = [];
    //     let cursor = firstMonth.clone();
    //     while (cursor.isSameOrBefore(lastMonth)) {
    //         allMonths.push(cursor.format('MMM-YYYY'));
    //         cursor.add(1, 'month');
    //     }
    //     allMonths.reverse();

    //     // Initialize grouped structure
    //     allMonths.forEach(month => {
    //         grouped[month] = {};
    //         sanctions.forEach(entry => {
    //             const loanType = entry.loan_type || 'N/A';
    //             loanTypes.add(loanType);
    //             grouped[month][loanType] = {
    //                 total_amount: 0,
    //                 total_weighted_interest: 0,
    //                 total_weighted_proc_fee: 0,
    //                 total_payment: 0,
    //                 interest_numerator: 0,
    //                 interest_denominator: 0,
    //                 total_tranche_amount: 0,
    //                 total_processing_fee_annually: 0,
    //             };
    //         });
    //     });

    //     // Populate grouped structure
    //     allMonths.forEach(month => {
    //         const monthMoment = moment(month, 'MMM-YYYY');

    //         sanctions.forEach(entry => {
    //             const sanctionMonth = moment(entry.sanction_date).startOf('month');
    //             if (sanctionMonth.isAfter(monthMoment)) return;

    //             const loanType = entry.loan_type || 'N/A';
    //             const data = grouped[month][loanType];
    //             const sanctionAmount = parseFloat(entry.sanction_amount || 0);

    //             // Filter payments for current sanction up to current month
    //             const currentMonthPayments = payments
    //                 .filter(p => p.sanction_id === entry.sanction_id)
    //                 .filter(p => moment(p.payment_date).startOf('month').isSameOrBefore(monthMoment));

    //             const totalPaymentAmount = currentMonthPayments.reduce((sum, p) => {
    //                 return sum + parseFloat(p.payment_amount || 0);
    //             }, 0);

    //             const interestRate = parseFloat(entry.interest_rate || 0);
    //             const procFeeRate = parseFloat(entry.processing_fee_rate || 0);
    //             const tenureMonths = parseFloat(entry.tenure_months || 0);

    //             data.total_amount += sanctionAmount;
    //             data.total_payment += totalPaymentAmount;
    //             data.total_weighted_interest += interestRate * sanctionAmount;
    //             data.total_weighted_proc_fee += procFeeRate * sanctionAmount;

    //             const processingFeeAnnually = (sanctionAmount * procFeeRate) / tenureMonths * 12;
    //             data.total_processing_fee_annually += processingFeeAnnually;
    //             data.total_tranche_amount += sanctionAmount;
    //         });

    //         payments.forEach(payment => {
    //             const paymentDate = moment(payment.payment_date).startOf('month');
    //             if (paymentDate.isAfter(monthMoment)) return;

    //             const sanction = sanctions.find(s => s.sanction_id === payment.sanction_id);
    //             if (!sanction) return;

    //             const loanType = sanction.loan_type || 'N/A';
    //             const data = grouped[month][loanType];
    //             const paymentAmount = parseFloat(payment.payment_amount || 0);

    //             const tranche = tranches.find(t =>
    //                 t.sanction_id === payment.sanction_id &&
    //                 t.tranche_id === payment.tranche_id
    //             );

    //             if (tranche) {
    //                 const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                 const diff = trancheAmount - paymentAmount;

    //                 if (diff > 0) {
    //                     const interestRate = parseFloat(tranche.interest_rate || 0) / 100;
    //                     data.interest_numerator += interestRate * diff;
    //                     data.interest_denominator += diff;
    //                 }
    //             }
    //         });
    //     });

    //     const sortedLoanTypes = Array.from(loanTypes);
    //     const applyFullBorder = cell => {
    //         cell.border = {
    //             top: { style: 'thin' },
    //             left: { style: 'thin' },
    //             bottom: { style: 'thin' },
    //             right: { style: 'thin' }
    //         };
    //     };

    //     // Header Row
    //     sheet.getCell('A3').value = 'As on';
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getCell('A3').font = { bold: true, size: 12 };
    //     applyFullBorder(sheet.getCell('A3'));

    //     let startCol = 2;
    //     allMonths.forEach(month => {
    //         sheet.mergeCells(3, startCol, 3, startCol + 3);
    //         const cell = sheet.getRow(3).getCell(startCol);
    //         cell.value = month;
    //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //         cell.font = { bold: true, size: 12 };

    //         for (let i = 0; i < 4; i++) {
    //             applyFullBorder(sheet.getRow(3).getCell(startCol + i));
    //         }

    //         startCol += 4;
    //     });

    //     // Sub-Headers
    //     const subHeaderRow = sheet.getRow(4);
    //     subHeaderRow.height = 30;
    //     subHeaderRow.getCell(1).value = 'Loan Type';
    //     subHeaderRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    //     subHeaderRow.getCell(1).font = { bold: true, size: 12 };
    //     applyFullBorder(subHeaderRow.getCell(1));

    //     let subCol = 2;
    //     allMonths.forEach(() => {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = subHeaderRow.getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true, size: 12 };
    //             applyFullBorder(cell);
    //         });
    //     });

    //     const formatRate = val => `${parseFloat(val).toFixed(2)}%`;

    //     // Data Rows
    //     sortedLoanTypes.forEach(loanType => {
    //         const row = [loanType];
    //         allMonths.forEach(month => {
    //             const data = grouped[month]?.[loanType];
    //             if (data && data.total_amount > 0) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const weightedInterest = data.interest_denominator > 0
    //                     ? data.interest_numerator / data.interest_denominator
    //                     : 0;
    //                 const avgProcessingFeeRate = data.total_tranche_amount > 0
    //                     ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100
    //                     : 0;
    //                 const totalRate = weightedInterest + avgProcessingFeeRate / 100;

    //                 row.push(
    //                     outstanding.toFixed(2),
    //                     formatRate(weightedInterest),
    //                     formatRate(avgProcessingFeeRate),
    //                     formatRate(totalRate)
    //                 );
    //             } else {
    //                 row.push('', '', '', '');
    //             }
    //         });

    //         const dataRow = sheet.addRow(row);
    //         dataRow.eachCell(cell => {
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             applyFullBorder(cell);
    //         });
    //     });

    //     // Total Row
    //     const totalRow = ['Total'];
    //     allMonths.forEach(month => {
    //         let totalAmount = 0, totalPayment = 0, interestNumerator = 0, interestDenominator = 0, totalProcessingFeeAnnually = 0;

    //         sortedLoanTypes.forEach(loanType => {
    //             const data = grouped[month]?.[loanType];
    //             if (data && data.total_amount > 0) {
    //                 totalAmount += data.total_amount;
    //                 totalPayment += data.total_payment;
    //                 interestNumerator += data.interest_numerator;
    //                 interestDenominator += data.interest_denominator;
    //                 totalProcessingFeeAnnually += data.total_processing_fee_annually;
    //             }
    //         });

    //         if (totalAmount > 0) {
    //             const outstanding = (totalAmount - totalPayment) / 1e7;
    //             const weightedInterest = interestDenominator > 0 ? interestNumerator / interestDenominator : 0;
    //             const avgProcessingFeeRate = totalAmount > 0
    //                 ? (totalProcessingFeeAnnually / totalAmount) * 100
    //                 : 0;
    //             const totalRate = weightedInterest + avgProcessingFeeRate / 100;

    //             totalRow.push(
    //                 outstanding.toFixed(2),
    //                 formatRate(weightedInterest),
    //                 formatRate(avgProcessingFeeRate),
    //                 formatRate(totalRate)
    //             );
    //         } else {
    //             totalRow.push('', '', '', '');
    //         }
    //     });

    //     const tRow = sheet.addRow(totalRow);
    //     tRow.eachCell(cell => {
    //         cell.font = { bold: true };
    //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyFullBorder(cell);
    //     });

    //     sheet.columns.forEach(col => {
    //         let maxLength = 10;
    //         col.eachCell({ includeEmpty: true }, cell => {
    //             const val = cell.value ? cell.value.toString() : '';
    //             maxLength = Math.max(maxLength, val.length);
    //         });
    //         col.width = maxLength + 2;
    //     });

    //     sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=Effective_Interest_Report.xlsx');
    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error('Error generating Excel:', error);
    //     res.status(500).send('Failed to generate summary report');
    // }

    // above done outstanding


    // try {
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest');

    //     const loanTypes = new Set();
    //     const grouped = {};

    //     const sanctionDates = sanctions.map(entry => moment(entry.sanction_date).startOf('month'));
    //     const firstMonth = moment.min(sanctionDates);
    //     const lastMonth = moment().startOf('month');

    //     const allMonths = [];
    //     let cursor = firstMonth.clone();
    //     while (cursor.isSameOrBefore(lastMonth)) {
    //         allMonths.push(cursor.format('MMM-YYYY'));
    //         cursor.add(1, 'month');
    //     }
    //     allMonths.reverse();

    //     allMonths.forEach(month => {
    //         grouped[month] = {};
    //         sanctions.forEach(entry => {
    //             const loanType = entry.loan_type || 'N/A';
    //             loanTypes.add(loanType);
    //             grouped[month][loanType] = {
    //                 total_amount: 0,
    //                 total_weighted_interest: 0,
    //                 total_weighted_proc_fee: 0,
    //                 total_payment: 0,
    //                 interest_numerator: 0,
    //                 interest_denominator: 0,
    //                 total_tranche_amount: 0,
    //                 total_processing_fee_annually: 0,
    //             };
    //         });
    //     });

    //     // Process sanctions and payments
    //     allMonths.forEach(month => {
    //         const monthMoment = moment(month, 'MMM-YYYY');

    //         sanctions.forEach(entry => {
    //             const sanctionMonth = moment(entry.sanction_date).startOf('month');
    //             if (sanctionMonth.isAfter(monthMoment)) return;

    //             const loanType = entry.loan_type || 'N/A';
    //             const data = grouped[month][loanType];
    //             const sanctionAmount = parseFloat(entry.sanction_amount || 0);

    //             const currentMonthPayments = payments
    //                 .filter(p => p.sanction_id === entry.sanction_id)
    //                 .filter(p => moment(p.payment_date).startOf('month').isSameOrBefore(monthMoment));

    //             const totalPaymentAmount = currentMonthPayments.reduce((sum, p) => {
    //                 return sum + parseFloat(p.payment_amount || 0);
    //             }, 0);

    //             const procFeeRate = parseFloat(entry.processing_fee_rate || 0);
    //             const tenureMonths = parseFloat(entry.tenure_months || 0);

    //             data.total_amount += sanctionAmount;
    //             data.total_payment += totalPaymentAmount;
    //             data.total_weighted_proc_fee += procFeeRate * sanctionAmount;

    //             const processingFeeAnnually = (sanctionAmount * procFeeRate) / tenureMonths * 12;
    //             data.total_processing_fee_annually += processingFeeAnnually;
    //             data.total_tranche_amount += sanctionAmount;
    //         });
    //     });

    //     // New logic for Weighted Avg. Interest Rate
    //     tranches.forEach(tranche => {
    //         const trancheDate = moment(tranche.tranche_date).startOf('month');

    //         allMonths.forEach(month => {
    //             const monthMoment = moment(month, 'MMM-YYYY');
    //             if (trancheDate.isAfter(monthMoment)) return;

    //             const sanction = sanctions.find(s => s.sanction_id === tranche.sanction_id);
    //             if (!sanction) return;

    //             const loanType = sanction.loan_type || 'N/A';
    //             const data = grouped[month][loanType];

    //             const trancheAmount = parseFloat(tranche.tranche_amount || 0);

    //             const paymentsForTranche = payments.filter(p =>
    //                 p.sanction_id === tranche.sanction_id &&
    //                 p.tranche_id === tranche.tranche_id &&
    //                 moment(p.payment_date).startOf('month').isSameOrBefore(monthMoment)
    //             );

    //             const totalPaid = paymentsForTranche.reduce((sum, p) => {
    //                 return sum + parseFloat(p.payment_amount || 0);
    //             }, 0);

    //             const trancheOutstanding = trancheAmount - totalPaid;
    //             const interestRate = parseFloat(tranche.interest_rate || 0) / 100;

    //             if (trancheOutstanding > 0) {
    //                 data.interest_numerator += interestRate * trancheOutstanding;
    //                 data.interest_denominator += trancheOutstanding;
    //             }
    //         });
    //     });

    //     const sortedLoanTypes = Array.from(loanTypes);
    //     const applyFullBorder = cell => {
    //         cell.border = {
    //             top: { style: 'thin' },
    //             left: { style: 'thin' },
    //             bottom: { style: 'thin' },
    //             right: { style: 'thin' }
    //         };
    //     };

    //     sheet.getCell('A3').value = 'As on';
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getCell('A3').font = { bold: true, size: 12 };
    //     applyFullBorder(sheet.getCell('A3'));

    //     let startCol = 2;
    //     allMonths.forEach(month => {
    //         sheet.mergeCells(3, startCol, 3, startCol + 3);
    //         const cell = sheet.getRow(3).getCell(startCol);
    //         cell.value = month;
    //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //         cell.font = { bold: true, size: 12 };

    //         for (let i = 0; i < 4; i++) {
    //             applyFullBorder(sheet.getRow(3).getCell(startCol + i));
    //         }

    //         startCol += 4;
    //     });

    //     const subHeaderRow = sheet.getRow(4);
    //     subHeaderRow.height = 30;
    //     subHeaderRow.getCell(1).value = 'Loan Type';
    //     subHeaderRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    //     subHeaderRow.getCell(1).font = { bold: true, size: 12 };
    //     applyFullBorder(subHeaderRow.getCell(1));

    //     let subCol = 2;
    //     allMonths.forEach(() => {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = subHeaderRow.getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true, size: 12 };
    //             applyFullBorder(cell);
    //         });
    //     });

    //     const formatRate = val => `${parseFloat(val).toFixed(2)}%`;

    //     sortedLoanTypes.forEach(loanType => {
    //         const row = [loanType];
    //         allMonths.forEach(month => {
    //             const data = grouped[month]?.[loanType];
    //             if (data && data.total_amount > 0) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const weightedInterest = data.interest_denominator > 0
    //                     ? data.interest_numerator / data.interest_denominator
    //                     : 0;
    //                 const avgProcessingFeeRate = data.total_tranche_amount > 0
    //                     ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100
    //                     : 0;
    //                 const totalRate = weightedInterest + avgProcessingFeeRate / 100;

    //                 row.push(
    //                     outstanding.toFixed(2),
    //                     formatRate(weightedInterest),
    //                     formatRate(avgProcessingFeeRate),
    //                     formatRate(totalRate)
    //                 );
    //             } else {
    //                 row.push('', '', '', '');
    //             }
    //         });

    //         const dataRow = sheet.addRow(row);
    //         dataRow.eachCell(cell => {
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             applyFullBorder(cell);
    //         });
    //     });

    //     const totalRow = ['Total'];
    //     allMonths.forEach(month => {
    //         let totalAmount = 0, totalPayment = 0, interestNumerator = 0, interestDenominator = 0, totalProcessingFeeAnnually = 0;

    //         sortedLoanTypes.forEach(loanType => {
    //             const data = grouped[month]?.[loanType];
    //             if (data && data.total_amount > 0) {
    //                 totalAmount += data.total_amount;
    //                 totalPayment += data.total_payment;
    //                 interestNumerator += data.interest_numerator;
    //                 interestDenominator += data.interest_denominator;
    //                 totalProcessingFeeAnnually += data.total_processing_fee_annually;
    //             }
    //         });

    //         if (totalAmount > 0) {
    //             const outstanding = (totalAmount - totalPayment) / 1e7;
    //             const weightedInterest = interestDenominator > 0 ? interestNumerator / interestDenominator : 0;
    //             const avgProcessingFeeRate = totalAmount > 0
    //                 ? (totalProcessingFeeAnnually / totalAmount) * 100
    //                 : 0;
    //             const totalRate = weightedInterest + avgProcessingFeeRate / 100;

    //             totalRow.push(
    //                 outstanding.toFixed(2),
    //                 formatRate(weightedInterest),
    //                 formatRate(avgProcessingFeeRate),
    //                 formatRate(totalRate)
    //             );
    //         } else {
    //             totalRow.push('', '', '', '');
    //         }
    //     });

    //     const tRow = sheet.addRow(totalRow);
    //     tRow.eachCell(cell => {
    //         cell.font = { bold: true };
    //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyFullBorder(cell);
    //     });

    //     sheet.columns.forEach(col => {
    //         let maxLength = 10;
    //         col.eachCell({ includeEmpty: true }, cell => {
    //             const val = cell.value ? cell.value.toString() : '';
    //             maxLength = Math.max(maxLength, val.length);
    //         });
    //         col.width = maxLength + 2;
    //     });

    //     sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=Effective_Interest_Report.xlsx');
    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error('Error generating Excel:', error);
    //     res.status(500).send('Failed to generate summary report');
    // }

    // test for weighted avg 
    // try {
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest');

    //     const loanTypes = new Set();
    //     const grouped = {};

    //     const sanctionDates = sanctions.map(entry => moment(entry.sanction_date).startOf('month'));
    //     const firstMonth = moment.min(sanctionDates);
    //     const lastMonth = moment().startOf('month');

    //     const allMonths = [];
    //     let cursor = firstMonth.clone();
    //     while (cursor.isSameOrBefore(lastMonth)) {
    //         allMonths.push(cursor.format('MMM-YYYY'));
    //         cursor.add(1, 'month');
    //     }
    //     allMonths.reverse();

    //     // Initialize grouped structure
    //     allMonths.forEach(month => {
    //         grouped[month] = {};
    //         sanctions.forEach(entry => {
    //             const loanType = entry.loan_type || 'N/A';
    //             loanTypes.add(loanType);
    //             grouped[month][loanType] = {
    //                 total_amount: 0,
    //                 total_weighted_interest: 0,
    //                 total_weighted_proc_fee: 0,
    //                 total_payment: 0,
    //                 interest_numerator: 0,
    //                 interest_denominator: 0,
    //                 total_tranche_amount: 0,
    //                 total_processing_fee_annually: 0,
    //                 total_interest: 0, // For total interest calculation
    //             };
    //         });
    //     });

    //     // Populate grouped structure
    //     allMonths.forEach(month => {
    //         const monthMoment = moment(month, 'MMM-YYYY');

    //         sanctions.forEach(entry => {
    //             const sanctionMonth = moment(entry.sanction_date).startOf('month');
    //             if (sanctionMonth.isAfter(monthMoment)) return;

    //             const loanType = entry.loan_type || 'N/A';
    //             const data = grouped[month][loanType];
    //             const sanctionAmount = parseFloat(entry.sanction_amount || 0);

    //             // Filter payments for current sanction up to current month
    //             const currentMonthPayments = payments
    //                 .filter(p => p.sanction_id === entry.sanction_id)
    //                 .filter(p => moment(p.payment_date).startOf('month').isSameOrBefore(monthMoment));

    //             const totalPaymentAmount = currentMonthPayments.reduce((sum, p) => {
    //                 return sum + parseFloat(p.payment_amount || 0);
    //             }, 0);

    //             const interestRate = parseFloat(entry.interest_rate || 0);
    //             const procFeeRate = parseFloat(entry.processing_fee_rate || 0);
    //             const tenureMonths = parseFloat(entry.tenure_months || 0);

    //             data.total_amount += sanctionAmount;
    //             data.total_payment += totalPaymentAmount;
    //             data.total_weighted_interest += interestRate * sanctionAmount;
    //             data.total_weighted_proc_fee += procFeeRate * sanctionAmount;

    //             const processingFeeAnnually = (sanctionAmount * procFeeRate) / tenureMonths * 12;
    //             data.total_processing_fee_annually += processingFeeAnnually;
    //             data.total_tranche_amount += sanctionAmount;
    //         });

    //         payments.forEach(payment => {
    //             const paymentDate = moment(payment.payment_date).startOf('month');
    //             if (paymentDate.isAfter(monthMoment)) return;

    //             const sanction = sanctions.find(s => s.sanction_id === payment.sanction_id);
    //             if (!sanction) return;

    //             const loanType = sanction.loan_type || 'N/A';
    //             const data = grouped[month][loanType];
    //             const paymentAmount = parseFloat(payment.payment_amount || 0);

    //             const tranche = tranches.find(t =>
    //                 t.sanction_id === payment.sanction_id &&
    //                 t.tranche_id === payment.tranche_id
    //             );

    //             if (tranche) {
    //                 const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                 const diff = trancheAmount - paymentAmount;

    //                 if (diff > 0) {
    //                     const interestRate = parseFloat(tranche.interest_rate || 0) / 100;
    //                     data.interest_numerator += interestRate * diff;
    //                     data.interest_denominator += diff;

    //                     // Calculate the Weighted Avg. Interest Rate
    //                     data.total_interest += diff;
    //                 }
    //             }
    //         });
    //     });

    //     const sortedLoanTypes = Array.from(loanTypes);
    //     const applyFullBorder = cell => {
    //         cell.border = {
    //             top: { style: 'thin' },
    //             left: { style: 'thin' },
    //             bottom: { style: 'thin' },
    //             right: { style: 'thin' }
    //         };
    //     };

    //     // Header Row
    //     sheet.getCell('A3').value = 'As on';
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getCell('A3').font = { bold: true, size: 12 };
    //     applyFullBorder(sheet.getCell('A3'));

    //     let startCol = 2;
    //     allMonths.forEach(month => {
    //         sheet.mergeCells(3, startCol, 3, startCol + 3);
    //         const cell = sheet.getRow(3).getCell(startCol);
    //         cell.value = month;
    //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //         cell.font = { bold: true, size: 12 };

    //         for (let i = 0; i < 4; i++) {
    //             applyFullBorder(sheet.getRow(3).getCell(startCol + i));
    //         }

    //         startCol += 4;
    //     });

    //     // Sub-Headers
    //     const subHeaderRow = sheet.getRow(4);
    //     subHeaderRow.height = 30;
    //     subHeaderRow.getCell(1).value = 'Loan Type';
    //     subHeaderRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    //     subHeaderRow.getCell(1).font = { bold: true, size: 12 };
    //     applyFullBorder(subHeaderRow.getCell(1));

    //     let subCol = 2;
    //     allMonths.forEach(() => {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = subHeaderRow.getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true, size: 12 };
    //             applyFullBorder(cell);
    //         });
    //     });

    //     const formatRate = val => `${parseFloat(val).toFixed(2)}%`;

    //     // Data Rows
    //     sortedLoanTypes.forEach(loanType => {
    //         const row = [loanType];
    //         allMonths.forEach(month => {
    //             const data = grouped[month]?.[loanType];
    //             if (data && data.total_amount > 0) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const weightedInterest = data.interest_denominator > 0
    //                     ? data.interest_numerator / data.interest_denominator
    //                     : 0;
    //                 const avgProcessingFeeRate = data.total_tranche_amount > 0
    //                     ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100
    //                     : 0;
    //                 const totalRate = weightedInterest + avgProcessingFeeRate / 100;

    //                 const weightedAvgInterestRate = data.total_interest > 0
    //                     ? (data.interest_numerator / data.total_interest)
    //                     : 0;

    //                 row.push(
    //                     outstanding.toFixed(2),
    //                     formatRate(weightedAvgInterestRate),
    //                     formatRate(avgProcessingFeeRate),
    //                     formatRate(totalRate)
    //                 );
    //             } else {
    //                 row.push('', '', '', '');
    //             }
    //         });

    //         const dataRow = sheet.addRow(row);
    //         dataRow.eachCell(cell => {
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             applyFullBorder(cell);
    //         });
    //     });

    //     // Total Row
    //     const totalRow = ['Total'];
    //     allMonths.forEach(month => {
    //         let totalAmount = 0, totalPayment = 0, interestNumerator = 0, interestDenominator = 0, totalProcessingFeeAnnually = 0;

    //         sortedLoanTypes.forEach(loanType => {
    //             const data = grouped[month]?.[loanType];
    //             if (data && data.total_amount > 0) {
    //                 totalAmount += data.total_amount;
    //                 totalPayment += data.total_payment;
    //                 interestNumerator += data.interest_numerator;
    //                 interestDenominator += data.interest_denominator;
    //                 totalProcessingFeeAnnually += data.total_processing_fee_annually;
    //             }
    //         });

    //         if (totalAmount > 0) {
    //             const outstanding = (totalAmount - totalPayment) / 1e7;
    //             const weightedInterest = interestDenominator > 0 ? interestNumerator / interestDenominator : 0;
    //             const avgProcessingFeeRate = totalAmount > 0
    //                 ? (totalProcessingFeeAnnually / totalAmount) * 100
    //                 : 0;
    //             const totalRate = weightedInterest + avgProcessingFeeRate / 100;

    //             totalRow.push(
    //                 outstanding.toFixed(2),
    //                 formatRate(weightedInterest),
    //                 formatRate(avgProcessingFeeRate),
    //                 formatRate(totalRate)
    //             );
    //         } else {
    //             totalRow.push('', '', '', '');
    //         }
    //     });

    //     const tRow = sheet.addRow(totalRow);
    //     tRow.eachCell(cell => {
    //         cell.font = { bold: true };
    //         cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyFullBorder(cell);
    //     });

    //     sheet.columns.forEach(col => {
    //         let maxLength = 10;
    //         col.eachCell({ includeEmpty: true }, cell => {
    //             const val = cell.value ? cell.value.toString() : '';
    //             maxLength = Math.max(maxLength, val.length);
    //         });
    //         col.width = maxLength + 2;
    //     });

    //     sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 4 }];

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename="effective_interest.xlsx"');
    //     await workbook.xlsx.write(res);
    //     res.end();
    // } catch (err) {
    //     console.log(err);
    //     res.status(500).send('Internal server error');
    // }

    //  todays code // 12:47

    // above caluculation mistake of weighted avg

    // try {
    //     // Fetch all data
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });

    //     // Initialize Excel workbook
    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest Report');

    //     // Find months range
    //     const sanctionDates = sanctions.map(e => moment(e.sanction_date));
    //     const firstMonth = moment.min(sanctionDates).startOf('month');
    //     const lastMonth = moment().startOf('month');
    //     const allMonths = [];

    //     let tempMonth = lastMonth.clone();
    //     while (tempMonth.isSameOrAfter(firstMonth)) {
    //         allMonths.push(tempMonth.format('MMM-YYYY'));
    //         tempMonth.subtract(1, 'month');
    //     }

    //     // Prepare structure
    //     const groupedData = {};
    //     const loanTypes = new Set();

    //     allMonths.forEach(month => {
    //         groupedData[month] = {};
    //     });

    //     // Populate sanctions data
    //     for (const sanction of sanctions) {
    //         const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //         const loanType = sanction.loan_type || 'Other';
    //         loanTypes.add(loanType);

    //         let sanctioned = false;
    //         for (const month of allMonths) {
    //             if (moment(month, 'MMM-YYYY').isSameOrAfter(moment(sanctionMonth, 'MMM-YYYY'))) {
    //                 if (!groupedData[month][loanType]) {
    //                     groupedData[month][loanType] = {
    //                         total_amount: 0,
    //                         total_payment: 0,
    //                         interest_numerator: 0,
    //                         interest_denominator: 0,
    //                         total_processing_fee_annually: 0,
    //                         total_tranche_amount: 0,
    //                     };
    //                 }
    //                 const data = groupedData[month][loanType];
    //                 const sanctionAmount = parseFloat(sanction.sanction_amount || 0);
    //                 const procFeeRate = parseFloat(sanction.processing_fee_rate || 0);
    //                 const tenureMonths = parseFloat(sanction.tenure_months || 1);

    //                 if (!sanctioned) {
    //                     const annualProcessingFee = (sanctionAmount * procFeeRate) / tenureMonths * 12;
    //                     data.total_processing_fee_annually += annualProcessingFee;
    //                     data.total_tranche_amount += sanctionAmount;
    //                     sanctioned = true;
    //                 }
    //                 data.total_amount += sanctionAmount;
    //             }
    //         }
    //     }

    //     // Process tranches and payments based on tranche_date
    //     for (const tranche of tranches) {
    //         const trancheDate = moment(tranche.tranche_date); // Get the tranche date
    //         const loanType = tranche.loan_type || 'Other';
    //         const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //         const interestRate = parseFloat(tranche.interest_rate || 0);

    //         // Find the month that corresponds to the tranche_date
    //         const trancheMonth = trancheDate.format('MMM-YYYY');

    //         // If the tranche month exists in the range of months, update the corresponding data
    //         if (groupedData[trancheMonth] && groupedData[trancheMonth][loanType]) {
    //             const data = groupedData[trancheMonth][loanType];
    //             const outstanding = data.total_amount - data.total_payment; // Remaining amount

    //             // Add tranche to the interest numerator and denominator
    //             if (outstanding > 0) {
    //                 data.interest_numerator += interestRate * outstanding;
    //                 data.interest_denominator += outstanding;
    //             }

    //             // Update total tranche amount
    //             data.total_tranche_amount += trancheAmount;
    //         }
    //     }

    //     // Process payments and calculate Weighted Average Interest Rate (WAIR)
    //     for (const payment of payments) {
    //         const paymentMonth = moment(payment.payment_date).format('MMM-YYYY');
    //         const sanction = sanctions.find(s => s.sanction_id === payment.sanction_id);
    //         if (!sanction) continue;

    //         const loanType = sanction.loan_type || 'Other';

    //         for (const month of allMonths) {
    //             if (moment(month, 'MMM-YYYY').isSameOrAfter(moment(paymentMonth, 'MMM-YYYY'))) {
    //                 const data = groupedData[month][loanType];
    //                 if (!data) continue;

    //                 const paymentAmount = parseFloat(payment.payment_amount || 0);
    //                 data.total_payment += paymentAmount;

    //                 const tranche = tranches.find(t =>
    //                     t.sanction_id === payment.sanction_id && t.tranche_id === payment.tranche_id
    //                 );

    //                 if (tranche) {
    //                     const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                     const interestRate = parseFloat(tranche.interest_rate || 0);
    //                     const trancheOutstanding = trancheAmount - paymentAmount;

    //                     if (trancheOutstanding > 0) {
    //                         data.interest_numerator += interestRate * trancheOutstanding;
    //                         data.interest_denominator += trancheOutstanding;
    //                     }
    //                 }
    //             }
    //         }
    //     }

    //     // Excel Headers
    //     sheet.getCell('A3').value = 'Loan Type';
    //     sheet.getCell('A3').font = { bold: true };
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getRow(3).height = 30;

    //     let colIndex = 2;
    //     for (const month of allMonths) {
    //         sheet.mergeCells(3, colIndex, 3, colIndex + 3);
    //         sheet.getCell(3, colIndex).value = month;
    //         sheet.getCell(3, colIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    //         sheet.getCell(3, colIndex).font = { bold: true };
    //         colIndex += 4;
    //     }

    //     // Sub-Headers
    //     sheet.getRow(4).getCell(1).value = 'Loan Type';
    //     let subCol = 2;
    //     for (const month of allMonths) {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = sheet.getRow(4).getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true };
    //         });
    //     }

    //     // Thin Border Function
    //     const applyThinBorder = (row) => {
    //         row.eachCell({ includeEmpty: true }, (cell) => {
    //             cell.border = {
    //                 top: { style: 'thin' },
    //                 left: { style: 'thin' },
    //                 bottom: { style: 'thin' },
    //                 right: { style: 'thin' }
    //             };
    //         });
    //     };

    //     // Format function
    //     const formatPercent = value => (parseFloat(value) || 0).toFixed(2) + '%';

    //     // Data Rows
    //     for (const loanType of Array.from(loanTypes)) {
    //         const rowData = [loanType];

    //         for (const month of allMonths) {
    //             const data = groupedData[month][loanType];

    //             if (data) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const wair = data.interest_denominator !== 0 ? (data.interest_numerator / data.interest_denominator) : 0;
    //                 const processingFeeRate = data.total_tranche_amount !== 0 ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100 : 0;
    //                 const totalRate = wair + (processingFeeRate / 100);

    //                 rowData.push(
    //                     outstanding.toFixed(2),
    //                     formatPercent(wair),
    //                     formatPercent(processingFeeRate),
    //                     formatPercent(totalRate)
    //                 );
    //             } else {
    //                 rowData.push('', '', '', '');
    //             }
    //         }

    //         const newRow = sheet.addRow(rowData);
    //         newRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyThinBorder(newRow);
    //     }

    //     // Total Row
    //     const totalRowData = ['Total'];

    //     for (const month of allMonths) {
    //         let total_amount = 0, total_payment = 0, total_interest_numerator = 0, total_interest_denominator = 0, total_processing_fee_annually = 0, total_tranche_amount = 0;

    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (data) {
    //                 total_amount += data.total_amount;
    //                 total_payment += data.total_payment;
    //                 total_interest_numerator += data.interest_numerator;
    //                 total_interest_denominator += data.interest_denominator;
    //                 total_processing_fee_annually += data.total_processing_fee_annually;
    //                 total_tranche_amount += data.total_tranche_amount;
    //             }
    //         }

    //         const outstanding = (total_amount - total_payment) / 1e7;
    //         const wair = total_interest_denominator !== 0 ? (total_interest_numerator / total_interest_denominator) : 0;
    //         const processingFeeRate = total_tranche_amount !== 0 ? (total_processing_fee_annually / total_tranche_amount) * 100 : 0;
    //         const totalRate = wair + (processingFeeRate / 100);

    //         totalRowData.push(
    //             outstanding.toFixed(2),
    //             formatPercent(wair),
    //             formatPercent(processingFeeRate),
    //             formatPercent(totalRate)
    //         );
    //     }

    //     const totalRow = sheet.addRow(totalRowData);
    //     totalRow.font = { bold: true };
    //     totalRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //     applyThinBorder(totalRow);

    //     // Format header rows
    //     applyThinBorder(sheet.getRow(3));
    //     applyThinBorder(sheet.getRow(4));

    //     // Column Width
    //     sheet.columns.forEach(column => {
    //         column.width = 22;
    //     });

    //     // Send Excel File
    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=EffectiveInterestReport.xlsx');

    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error('Error generating WAIR Report:', error);
    //     res.status(500).send('Error generating report');
    // }

    // above working for 1 tranche_id

    // try {
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });

    //     const sanctionMap = new Map();
    //     sanctions.forEach(s => sanctionMap.set(s.sanction_id, s));

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest Report');

    //     const sanctionDates = sanctions.map(e => moment(e.sanction_date));
    //     const firstMonth = moment.min(sanctionDates).startOf('month');
    //     const lastMonth = moment().startOf('month');
    //     const allMonths = [];

    //     let tempMonth = lastMonth.clone();
    //     while (tempMonth.isSameOrAfter(firstMonth)) {
    //         allMonths.push(tempMonth.format('MMM-YYYY'));
    //         tempMonth.subtract(1, 'month');
    //     }

    //     const groupedData = {};
    //     const loanTypes = new Set();

    //     allMonths.forEach(month => {
    //         groupedData[month] = {};
    //     });

    //     for (const sanction of sanctions) {
    //         const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //         const loanType = sanction.loan_type || 'Other';
    //         loanTypes.add(loanType);

    //         let sanctioned = false;
    //         for (const month of allMonths) {
    //             if (moment(month, 'MMM-YYYY').isSameOrAfter(moment(sanctionMonth, 'MMM-YYYY'))) {
    //                 if (!groupedData[month][loanType]) {
    //                     groupedData[month][loanType] = {
    //                         total_amount: 0,
    //                         total_payment: 0,
    //                         interest_numerator: 0,
    //                         interest_denominator: 0,
    //                         total_processing_fee_annually: 0,
    //                         total_tranche_amount: 0,
    //                     };
    //                 }
    //                 const data = groupedData[month][loanType];
    //                 const sanctionAmount = parseFloat(sanction.sanction_amount || 0);
    //                 const procFeeRate = parseFloat(sanction.processing_fee_rate || 0);
    //                 const tenureMonths = parseFloat(sanction.tenure_months || 1);

    //                 if (!sanctioned) {
    //                     const annualProcessingFee = (sanctionAmount * procFeeRate) / tenureMonths * 12;
    //                     data.total_processing_fee_annually += annualProcessingFee;
    //                     data.total_tranche_amount += sanctionAmount;
    //                     sanctioned = true;
    //                 }
    //                 data.total_amount += sanctionAmount;
    //             }
    //         }
    //     }

    //     // Prepare payments map for faster lookup
    //     const paymentMap = new Map();
    //     for (const payment of payments) {
    //         const key = `${payment.sanction_id}_${moment(payment.payment_date).format('MMM-YYYY')}`;
    //         if (!paymentMap.has(key)) paymentMap.set(key, 0);
    //         paymentMap.set(key, paymentMap.get(key) + parseFloat(payment.payment_amount || 0));
    //     }

    //     // Process month by month
    //     for (const month of allMonths) {
    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (!data) continue;

    //             for (const sanction of sanctions.filter(s => (s.loan_type || 'Other') === loanType)) {
    //                 const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //                 if (moment(month, 'MMM-YYYY').isBefore(moment(sanctionMonth, 'MMM-YYYY'))) continue;

    //                 // For this sanction in this month, get payment
    //                 const paymentKey = `${sanction.sanction_id}_${month}`;
    //                 const paymentAmount = paymentMap.get(paymentKey) || 0;

    //                 data.total_payment += paymentAmount;
    //             }

    //             // Calculate WAIR now
    //             const outstanding = data.total_amount - data.total_payment;

    //             if (outstanding > 0) {
    //                 // Find tranches linked to sanctions
    //                 for (const tranche of tranches) {
    //                     const sanction = sanctionMap.get(tranche.sanction_id);
    //                     if (!sanction) continue;
    //                     const trancheLoanType = sanction.loan_type || 'Other';
    //                     if (trancheLoanType !== loanType) continue;

    //                     const trancheDate = moment(tranche.tranche_date).format('MMM-YYYY');
    //                     if (moment(trancheDate, 'MMM-YYYY').isAfter(moment(month, 'MMM-YYYY'))) continue;

    //                     const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                     const interestRate = parseFloat(tranche.interest_rate || 0);

    //                     data.interest_numerator += interestRate * trancheAmount;
    //                     data.interest_denominator += trancheAmount;
    //                 }
    //             }
    //         }
    //     }

    //     // Excel Headers
    //     sheet.getCell('A3').value = 'Loan Type';
    //     sheet.getCell('A3').font = { bold: true };
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getRow(3).height = 30;

    //     let colIndex = 2;
    //     for (const month of allMonths) {
    //         sheet.mergeCells(3, colIndex, 3, colIndex + 3);
    //         sheet.getCell(3, colIndex).value = month;
    //         sheet.getCell(3, colIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    //         sheet.getCell(3, colIndex).font = { bold: true };
    //         colIndex += 4;
    //     }

    //     sheet.getRow(4).getCell(1).value = 'Loan Type';
    //     let subCol = 2;
    //     for (const month of allMonths) {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = sheet.getRow(4).getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true };
    //         });
    //     }

    //     const applyThinBorder = (row) => {
    //         row.eachCell({ includeEmpty: true }, (cell) => {
    //             cell.border = {
    //                 top: { style: 'thin' },
    //                 left: { style: 'thin' },
    //                 bottom: { style: 'thin' },
    //                 right: { style: 'thin' }
    //             };
    //         });
    //     };

    //     const formatPercent = value => (parseFloat(value) || 0).toFixed(2) + '%';

    //     for (const loanType of Array.from(loanTypes)) {
    //         const rowData = [loanType];

    //         for (const month of allMonths) {
    //             const data = groupedData[month][loanType];

    //             if (data) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const wair = data.interest_denominator !== 0 ? (data.interest_numerator / data.interest_denominator) : 0;
    //                 const processingFeeRate = data.total_tranche_amount !== 0 ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100 : 0;
    //                 const totalRate = wair + (processingFeeRate / 100);

    //                 rowData.push(
    //                     outstanding.toFixed(2),
    //                     formatPercent(wair),
    //                     formatPercent(processingFeeRate),
    //                     formatPercent(totalRate)
    //                 );
    //             } else {
    //                 rowData.push('', '', '', '');
    //             }
    //         }

    //         const newRow = sheet.addRow(rowData);
    //         newRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyThinBorder(newRow);
    //     }

    //     const totalRowData = ['Total'];

    //     for (const month of allMonths) {
    //         let total_amount = 0, total_payment = 0, total_interest_numerator = 0, total_interest_denominator = 0, total_processing_fee_annually = 0, total_tranche_amount = 0;

    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (data) {
    //                 total_amount += data.total_amount;
    //                 total_payment += data.total_payment;
    //                 total_interest_numerator += data.interest_numerator;
    //                 total_interest_denominator += data.interest_denominator;
    //                 total_processing_fee_annually += data.total_processing_fee_annually;
    //                 total_tranche_amount += data.total_tranche_amount;
    //             }
    //         }

    //         const outstanding = (total_amount - total_payment) / 1e7;
    //         const wair = total_interest_denominator !== 0 ? (total_interest_numerator / total_interest_denominator) : 0;
    //         const processingFeeRate = total_tranche_amount !== 0 ? (total_processing_fee_annually / total_tranche_amount) * 100 : 0;
    //         const totalRate = wair + (processingFeeRate / 100);

    //         totalRowData.push(
    //             outstanding.toFixed(2),
    //             formatPercent(wair),
    //             formatPercent(processingFeeRate),
    //             formatPercent(totalRate)
    //         );
    //     }

    //     const totalRow = sheet.addRow(totalRowData);
    //     totalRow.font = { bold: true };
    //     totalRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //     applyThinBorder(totalRow);

    //     applyThinBorder(sheet.getRow(3));
    //     applyThinBorder(sheet.getRow(4));

    //     sheet.columns.forEach(column => {
    //         column.width = 22;
    //     });

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=EffectiveInterestReport.xlsx');

    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error('Error generating WAIR Report:', error);
    //     res.status(500).send('Error generating report');
    // }

    //below is 28th apr code

    // try {
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });

    //     const sanctionMap = new Map();
    //     sanctions.forEach(s => sanctionMap.set(s.sanction_id, s));

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest Report');

    //     const sanctionDates = sanctions.map(e => moment(e.sanction_date));
    //     const firstMonth = moment.min(sanctionDates).startOf('month');
    //     const lastMonth = moment().startOf('month');
    //     const allMonths = [];

    //     let tempMonth = lastMonth.clone();
    //     while (tempMonth.isSameOrAfter(firstMonth)) {
    //         allMonths.push(tempMonth.format('MMM-YYYY'));
    //         tempMonth.subtract(1, 'month');
    //     }

    //     const groupedData = {};
    //     const loanTypes = new Set();

    //     allMonths.forEach(month => {
    //         groupedData[month] = {};
    //     });

    //     for (const sanction of sanctions) {
    //         const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //         const loanType = sanction.loan_type || 'Other';
    //         loanTypes.add(loanType);

    //         let sanctioned = false;
    //         for (const month of allMonths) {
    //             if (moment(month, 'MMM-YYYY').isSameOrAfter(moment(sanctionMonth, 'MMM-YYYY'))) {
    //                 if (!groupedData[month][loanType]) {
    //                     groupedData[month][loanType] = {
    //                         total_amount: 0,
    //                         total_payment: 0,
    //                         interest_numerator: 0,
    //                         interest_denominator: 0,
    //                         total_processing_fee_annually: 0,
    //                         total_tranche_amount: 0,
    //                     };
    //                 }
    //                 const data = groupedData[month][loanType];
    //                 const sanctionAmount = parseFloat(sanction.sanction_amount || 0);
    //                 const procFeeRate = parseFloat(sanction.processing_fee_rate || 0);
    //                 const tenureMonths = parseFloat(sanction.tenure_months || 1);

    //                 if (!sanctioned) {
    //                     const annualProcessingFee = (sanctionAmount * procFeeRate) / tenureMonths * 12;
    //                     data.total_processing_fee_annually += annualProcessingFee;
    //                     data.total_tranche_amount += sanctionAmount;
    //                     sanctioned = true;
    //                 }
    //                 data.total_amount += sanctionAmount;
    //             }
    //         }
    //     }

    //     // Prepare payments map for faster lookup
    //     const paymentMap = new Map();
    //     for (const payment of payments) {
    //         const key = `${payment.sanction_id}_${moment(payment.payment_date).format('MMM-YYYY')}`;
    //         if (!paymentMap.has(key)) paymentMap.set(key, 0);
    //         paymentMap.set(key, paymentMap.get(key) + parseFloat(payment.payment_amount || 0));
    //     }

    //     // Process month by month
    //     for (const month of allMonths) {
    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (!data) continue;

    //             for (const sanction of sanctions.filter(s => (s.loan_type || 'Other') === loanType)) {
    //                 const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //                 if (moment(month, 'MMM-YYYY').isBefore(moment(sanctionMonth, 'MMM-YYYY'))) continue;

    //                 // For this sanction in this month, get payment
    //                 const paymentKey = `${sanction.sanction_id}_${month}`;
    //                 const paymentAmount = paymentMap.get(paymentKey) || 0;

    //                 data.total_payment += paymentAmount;
    //             }

    //             // Calculate WAIR now
    //             const outstanding = data.total_amount - data.total_payment;

    //             if (outstanding > 0) {
    //                 // Find tranches linked to sanctions
    //                 for (const tranche of tranches) {
    //                     const sanction = sanctionMap.get(tranche.sanction_id);
    //                     if (!sanction) continue;
    //                     const trancheLoanType = sanction.loan_type || 'Other';
    //                     if (trancheLoanType !== loanType) continue;

    //                     const trancheDate = moment(tranche.tranche_date).format('MMM-YYYY');
    //                     if (moment(trancheDate, 'MMM-YYYY').isAfter(moment(month, 'MMM-YYYY'))) continue;

    //                     const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                     const interestRate = parseFloat(tranche.interest_rate || 0);

    //                     // Calculate the difference between tranche_amount and payment_amount
    //                     const paymentKey = `${sanction.sanction_id}_${month}`;
    //                     const totalPaid = paymentMap.get(paymentKey) || 0;
    //                     const adjustedTrancheAmount = trancheAmount - totalPaid;

    //                     // Now calculate the numerator and denominator for the interest rate
    //                     data.interest_numerator += interestRate * adjustedTrancheAmount;
    //                     data.interest_denominator += adjustedTrancheAmount;
    //                 }
    //             }
    //         }
    //     }

    //     // Excel Headers
    //     sheet.getCell('A3').value = 'Loan Type';
    //     sheet.getCell('A3').font = { bold: true };
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getRow(3).height = 30;

    //     let colIndex = 2;
    //     for (const month of allMonths) {
    //         sheet.mergeCells(3, colIndex, 3, colIndex + 3);
    //         sheet.getCell(3, colIndex).value = month;
    //         sheet.getCell(3, colIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    //         sheet.getCell(3, colIndex).font = { bold: true };
    //         colIndex += 4;
    //     }

    //     sheet.getRow(4).getCell(1).value = 'Loan Type';
    //     let subCol = 2;
    //     for (const month of allMonths) {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = sheet.getRow(4).getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true };
    //         });
    //     }

    //     const applyThinBorder = (row) => {
    //         row.eachCell({ includeEmpty: true }, (cell) => {
    //             cell.border = {
    //                 top: { style: 'thin' },
    //                 left: { style: 'thin' },
    //                 bottom: { style: 'thin' },
    //                 right: { style: 'thin' }
    //             };
    //         });
    //     };

    //     const formatPercent = value => (parseFloat(value) || 0).toFixed(2) + '%';

    //     for (const loanType of Array.from(loanTypes)) {
    //         const rowData = [loanType];

    //         for (const month of allMonths) {
    //             const data = groupedData[month][loanType];

    //             if (data) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const wair = data.interest_denominator !== 0 ? (data.interest_numerator / data.interest_denominator) : 0;
    //                 const processingFeeRate = data.total_tranche_amount !== 0 ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100 : 0;
    //                 const totalRate = wair + (processingFeeRate / 100);

    //                 rowData.push(
    //                     outstanding.toFixed(2),
    //                     formatPercent(wair),
    //                     formatPercent(processingFeeRate),
    //                     formatPercent(totalRate)
    //                 );
    //             } else {
    //                 rowData.push('', '', '', '');
    //             }
    //         }

    //         const newRow = sheet.addRow(rowData);
    //         newRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyThinBorder(newRow);
    //     }

    //     const totalRowData = ['Total'];

    //     for (const month of allMonths) {
    //         let total_amount = 0, total_payment = 0, total_interest_numerator = 0, total_interest_denominator = 0, total_processing_fee_annually = 0, total_tranche_amount = 0;

    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (data) {
    //                 total_amount += data.total_amount;
    //                 total_payment += data.total_payment;
    //                 total_interest_numerator += data.interest_numerator;
    //                 total_interest_denominator += data.interest_denominator;
    //                 total_processing_fee_annually += data.total_processing_fee_annually;
    //                 total_tranche_amount += data.total_tranche_amount;
    //             }
    //         }

    //         const outstanding = (total_amount - total_payment) / 1e7;
    //         const wair = total_interest_denominator !== 0 ? (total_interest_numerator / total_interest_denominator) : 0;
    //         const processingFeeRate = total_tranche_amount !== 0 ? (total_processing_fee_annually / total_tranche_amount) * 100 : 0;
    //         const totalRate = wair + (processingFeeRate / 100);

    //         totalRowData.push(
    //             outstanding.toFixed(2),
    //             formatPercent(wair),
    //             formatPercent(processingFeeRate),
    //             formatPercent(totalRate)
    //         );
    //     }

    //     const totalRow = sheet.addRow(totalRowData);
    //     totalRow.font = { bold: true };
    //     totalRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //     applyThinBorder(totalRow);

    //     applyThinBorder(sheet.getRow(3));
    //     applyThinBorder(sheet.getRow(4));

    //     sheet.columns.forEach(column => {
    //         column.width = 22;
    //     });

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=EffectiveInterestReport.xlsx');

    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error('Error generating WAIR Report:', error);
    //     res.status(500).send('Error generating report');
    // }

    // code 11:08

    // try {
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });

    //     const sanctionMap = new Map();
    //     sanctions.forEach(s => sanctionMap.set(s.sanction_id, s));

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest Report');

    //     const sanctionDates = sanctions.map(e => moment(e.sanction_date));
    //     const firstMonth = moment.min(sanctionDates).startOf('month');
    //     const lastMonth = moment().startOf('month');
    //     const allMonths = [];

    //     let tempMonth = lastMonth.clone();
    //     while (tempMonth.isSameOrAfter(firstMonth)) {
    //         allMonths.push(tempMonth.format('MMM-YYYY'));
    //         tempMonth.subtract(1, 'month');
    //     }

    //     const groupedData = {};
    //     const loanTypes = new Set();

    //     allMonths.forEach(month => {
    //         groupedData[month] = {};
    //     });

    //     for (const sanction of sanctions) {
    //         const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //         const loanType = sanction.loan_type || 'Other';
    //         loanTypes.add(loanType);

    //         let sanctioned = false;
    //         for (const month of allMonths) {
    //             if (moment(month, 'MMM-YYYY').isSameOrAfter(moment(sanctionMonth, 'MMM-YYYY'))) {
    //                 if (!groupedData[month][loanType]) {
    //                     groupedData[month][loanType] = {
    //                         total_amount: 0,
    //                         total_payment: 0,
    //                         interest_numerator: 0,
    //                         interest_denominator: 0,
    //                         total_processing_fee_annually: 0,
    //                         total_tranche_amount: 0,
    //                     };
    //                 }
    //                 const data = groupedData[month][loanType];
    //                 const sanctionAmount = parseFloat(sanction.sanction_amount || 0);
    //                 const procFeeRate = parseFloat(sanction.processing_fee_rate || 0);
    //                 const tenureMonths = parseFloat(sanction.tenure_months || 1);

    //                 if (!sanctioned) {
    //                     const annualProcessingFee = (sanctionAmount * procFeeRate) / tenureMonths * 12;
    //                     data.total_processing_fee_annually += annualProcessingFee;
    //                     data.total_tranche_amount += sanctionAmount;
    //                     sanctioned = true;
    //                 }
    //                 data.total_amount += sanctionAmount;
    //             }
    //         }
    //     }

    //     // Prepare payments map for faster lookup
    //     const paymentMap = new Map();
    //     for (const payment of payments) {
    //         const key = `${payment.sanction_id}_${moment(payment.payment_date).format('MMM-YYYY')}`;
    //         if (!paymentMap.has(key)) paymentMap.set(key, 0);
    //         paymentMap.set(key, paymentMap.get(key) + parseFloat(payment.payment_amount || 0));
    //     }

    //     // Process month by month
    //     for (const month of allMonths) {
    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (!data) continue;

    //             for (const sanction of sanctions.filter(s => (s.loan_type || 'Other') === loanType)) {
    //                 const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //                 if (moment(month, 'MMM-YYYY').isBefore(moment(sanctionMonth, 'MMM-YYYY'))) continue;

    //                 // For this sanction in this month, get payment
    //                 const paymentKey = `${sanction.sanction_id}_${month}`;
    //                 const paymentAmount = paymentMap.get(paymentKey) || 0;

    //                 data.total_payment += paymentAmount;
    //             }

    //             // Calculate WAIR now
    //             const outstanding = data.total_amount - data.total_payment;

    //             if (outstanding > 0) {
    //                 // Find tranches linked to sanctions
    //                 for (const tranche of tranches) {
    //                     const sanction = sanctionMap.get(tranche.sanction_id);
    //                     if (!sanction) continue;
    //                     const trancheLoanType = sanction.loan_type || 'Other';
    //                     if (trancheLoanType !== loanType) continue;

    //                     const trancheDate = moment(tranche.tranche_date).format('MMM-YYYY');
    //                     if (moment(trancheDate, 'MMM-YYYY').isAfter(moment(month, 'MMM-YYYY'))) continue;

    //                     const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                     const interestRate = parseFloat(tranche.interest_rate || 0);

    //                     // Calculate the difference between tranche_amount and payment_amount
    //                     const paymentKey = `${sanction.sanction_id}_${month}`;
    //                     const totalPaid = paymentMap.get(paymentKey) || 0;
    //                     const adjustedTrancheAmount = trancheAmount - totalPaid;

    //                     // Now calculate the numerator and denominator for the interest rate
    //                     data.interest_numerator += interestRate * adjustedTrancheAmount;
    //                     data.interest_denominator += adjustedTrancheAmount;
    //                 }
    //             }
    //         }
    //     }

    //     // Excel Headers
    //     sheet.getCell('A3').value = 'Loan Type';
    //     sheet.getCell('A3').font = { bold: true };
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getRow(3).height = 30;

    //     let colIndex = 2;
    //     for (const month of allMonths) {
    //         sheet.mergeCells(3, colIndex, 3, colIndex + 3);
    //         sheet.getCell(3, colIndex).value = month;
    //         sheet.getCell(3, colIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    //         sheet.getCell(3, colIndex).font = { bold: true };
    //         colIndex += 4;
    //     }

    //     sheet.getRow(4).getCell(1).value = 'Loan Type';
    //     let subCol = 2;
    //     for (const month of allMonths) {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = sheet.getRow(4).getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true };
    //         });
    //     }

    //     const applyThinBorder = (row) => {
    //         row.eachCell({ includeEmpty: true }, (cell) => {
    //             cell.border = {
    //                 top: { style: 'thin' },
    //                 left: { style: 'thin' },
    //                 bottom: { style: 'thin' },
    //                 right: { style: 'thin' }
    //             };
    //         });
    //     };

    //     const formatPercent = value => (parseFloat(value) || 0).toFixed(2) + '%';

    //     // Add rows for each loan type
    //     for (const loanType of Array.from(loanTypes)) {
    //         const rowData = [loanType];

    //         for (const month of allMonths) {
    //             const data = groupedData[month][loanType];

    //             if (data) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const wair = data.interest_denominator !== 0 ? (data.interest_numerator / data.interest_denominator) : 0;
    //                 const processingFeeRate = data.total_tranche_amount !== 0 ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100 : 0;
    //                 const totalRate = wair + (processingFeeRate / 100);

    //                 rowData.push(
    //                     outstanding.toFixed(2),
    //                     formatPercent(wair),
    //                     formatPercent(processingFeeRate),
    //                     formatPercent(totalRate)
    //                 );
    //             } else {
    //                 rowData.push('', '', '', '');
    //             }
    //         }

    //         const newRow = sheet.addRow(rowData);
    //         newRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyThinBorder(newRow);
    //     }

    //     // Total Row Data
    //     const totalRowData = ['Total'];

    //     for (const month of allMonths) {
    //         let total_amount = 0, total_payment = 0, total_interest_numerator = 0, total_interest_denominator = 0, total_processing_fee_annually = 0, total_tranche_amount = 0;

    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (data) {
    //                 total_amount += data.total_amount;
    //                 total_payment += data.total_payment;
    //                 total_interest_numerator += data.interest_numerator;
    //                 total_interest_denominator += data.interest_denominator;
    //                 total_processing_fee_annually += data.total_processing_fee_annually;
    //                 total_tranche_amount += data.total_tranche_amount;
    //             }
    //         }

    //         const outstanding = (total_amount - total_payment) / 1e7;
    //         const wair = total_interest_denominator !== 0 ? (total_interest_numerator / total_interest_denominator) : 0;
    //         const processingFeeRate = total_tranche_amount !== 0 ? (total_processing_fee_annually / total_tranche_amount) * 100 : 0;
    //         const totalRate = wair + (processingFeeRate / 100);

    //         totalRowData.push(
    //             outstanding.toFixed(2),
    //             formatPercent(wair),
    //             formatPercent(processingFeeRate),
    //             formatPercent(totalRate)
    //         );
    //     }

    //     const totalRow = sheet.addRow(totalRowData);
    //     totalRow.font = { bold: true };
    //     totalRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //     applyThinBorder(totalRow);

    //     applyThinBorder(sheet.getRow(3));
    //     applyThinBorder(sheet.getRow(4));

    //     sheet.columns.forEach(column => {
    //         column.width = 22;
    //     });

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=EffectiveInterestReport.xlsx');

    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error('Error generating WAIR Report:', error);
    //     res.status(500).send('Error generating report');
    // }

    //code 2 good above

    // try {
    //     const sanctions = await sanction_details.findAll({ raw: true });
    //     const tranches = await tranche_details.findAll({ raw: true });
    //     const payments = await payment_details.findAll({ raw: true });

    //     const sanctionMap = new Map();
    //     sanctions.forEach(s => sanctionMap.set(s.sanction_id, s));

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest Report');

    //     const sanctionDates = sanctions.map(e => moment(e.sanction_date));
    //     const firstMonth = moment.min(sanctionDates).startOf('month');
    //     const lastMonth = moment().startOf('month');
    //     const allMonths = [];

    //     let tempMonth = lastMonth.clone();
    //     while (tempMonth.isSameOrAfter(firstMonth)) {
    //         allMonths.push(tempMonth.format('MMM-YYYY'));
    //         tempMonth.subtract(1, 'month');
    //     }

    //     const groupedData = {};
    //     const loanTypes = new Set();

    //     allMonths.forEach(month => {
    //         groupedData[month] = {};
    //     });

    //     for (const sanction of sanctions) {
    //         const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //         const loanType = sanction.loan_type || 'Other';
    //         loanTypes.add(loanType);

    //         let sanctioned = false;
    //         for (const month of allMonths) {
    //             if (moment(month, 'MMM-YYYY').isSameOrAfter(moment(sanctionMonth, 'MMM-YYYY'))) {
    //                 if (!groupedData[month][loanType]) {
    //                     groupedData[month][loanType] = {
    //                         total_amount: 0,
    //                         total_payment: 0,
    //                         interest_numerator: 0,
    //                         interest_denominator: 0,
    //                         total_processing_fee_annually: 0,
    //                         total_tranche_amount: 0,
    //                     };
    //                 }
    //                 const data = groupedData[month][loanType];
    //                 const sanctionAmount = parseFloat(sanction.sanction_amount || 0);
    //                 const trancheAmount = parseFloat(tranches.tranche_amount || 0);
    //                 const procFeeRate = parseFloat(sanction.processing_fee || 0) / 100;
    //                 const tenureMonths = parseFloat(sanction.tenure_months || 1);

    //                 if (!sanctioned) {
    //                     // Calculating the annual processing fee with the correct formula
    //                     const annualProcessingFee = (trancheAmount * procFeeRate) / tenureMonths * 12;
    //                     const annualProcessingFeeNumber = annualProcessingFee * 100;
    //                     console.log("annual: ", annualProcessingFeeNumber)
    //                     data.total_processing_fee_annually += annualProcessingFeeNumber;
    //                     data.total_tranche_amount += sanctionAmount;
    //                     sanctioned = true;
    //                 }
    //                 data.total_amount += sanctionAmount;
    //             }
    //         }
    //     }

    //     // Prepare payments map for faster lookup
    //     const paymentMap = new Map();
    //     for (const payment of payments) {
    //         const key = `${payment.sanction_id}_${moment(payment.payment_date).format('MMM-YYYY')}`;
    //         if (!paymentMap.has(key)) paymentMap.set(key, 0);
    //         paymentMap.set(key, paymentMap.get(key) + parseFloat(payment.payment_amount || 0));
    //     }

    //     // Process month by month
    //     for (const month of allMonths) {
    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (!data) continue;

    //             for (const sanction of sanctions.filter(s => (s.loan_type || 'Other') === loanType)) {
    //                 const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
    //                 if (moment(month, 'MMM-YYYY').isBefore(moment(sanctionMonth, 'MMM-YYYY'))) continue;

    //                 // For this sanction in this month, get payment
    //                 const paymentKey = `${sanction.sanction_id}_${month}`;
    //                 const paymentAmount = paymentMap.get(paymentKey) || 0;

    //                 data.total_payment += paymentAmount;
    //             }

    //             // Calculate WAIR now
    //             const outstanding = data.total_amount - data.total_payment;

    //             if (outstanding > 0) {
    //                 // Find tranches linked to sanctions
    //                 for (const tranche of tranches) {
    //                     const sanction = sanctionMap.get(tranche.sanction_id);
    //                     if (!sanction) continue;
    //                     const trancheLoanType = sanction.loan_type || 'Other';
    //                     if (trancheLoanType !== loanType) continue;

    //                     const trancheDate = moment(tranche.tranche_date).format('MMM-YYYY');
    //                     if (moment(trancheDate, 'MMM-YYYY').isAfter(moment(month, 'MMM-YYYY'))) continue;

    //                     const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                     console.log("in tranches: ", trancheAmount)
    //                     const interestRate = parseFloat(tranche.interest_rate || 0);

    //                     // Calculate the difference between tranche_amount and payment_amount
    //                     const paymentKey = `${sanction.sanction_id}_${month}`;
    //                     const totalPaid = paymentMap.get(paymentKey) || 0;
    //                     const adjustedTrancheAmount = trancheAmount - totalPaid;

    //                     // Now calculate the numerator and denominator for the interest rate
    //                     data.interest_numerator += interestRate * adjustedTrancheAmount;
    //                     data.interest_denominator += adjustedTrancheAmount;
    //                 }
    //             }
    //         }
    //     }

    //     // Excel Headers
    //     sheet.getCell('A3').value = 'Loan Type';
    //     sheet.getCell('A3').font = { bold: true };
    //     sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
    //     sheet.getRow(3).height = 30;

    //     let colIndex = 2;
    //     for (const month of allMonths) {
    //         sheet.mergeCells(3, colIndex, 3, colIndex + 3);
    //         sheet.getCell(3, colIndex).value = month;
    //         sheet.getCell(3, colIndex).alignment = { vertical: 'middle', horizontal: 'center' };
    //         sheet.getCell(3, colIndex).font = { bold: true };
    //         colIndex += 4;
    //     }

    //     sheet.getRow(4).getCell(1).value = 'Loan Type';
    //     let subCol = 2;
    //     for (const month of allMonths) {
    //         ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
    //             const cell = sheet.getRow(4).getCell(subCol++);
    //             cell.value = label;
    //             cell.alignment = { vertical: 'middle', horizontal: 'center' };
    //             cell.font = { bold: true };
    //         });
    //     }

    //     const applyThinBorder = (row) => {
    //         row.eachCell({ includeEmpty: true }, (cell) => {
    //             cell.border = {
    //                 top: { style: 'thin' },
    //                 left: { style: 'thin' },
    //                 bottom: { style: 'thin' },
    //                 right: { style: 'thin' }
    //             };
    //         });
    //     };

    //     const formatPercent = value => (parseFloat(value) || 0).toFixed(2) + '%';

    //     // Add rows for each loan type
    //     for (const loanType of Array.from(loanTypes)) {
    //         const rowData = [loanType];

    //         for (const month of allMonths) {
    //             const data = groupedData[month][loanType];

    //             if (data) {
    //                 const outstanding = (data.total_amount - data.total_payment) / 1e7;
    //                 const wair = data.interest_denominator !== 0 ? (data.interest_numerator / data.interest_denominator) : 0;
    //                 const processingFeeRate = data.total_tranche_amount !== 0 ? (data.total_processing_fee_annually / data.total_tranche_amount) * 100 : 0;
    //                 console.log("annual: ", processingFeeRate)
    //                 const totalRate = wair + (processingFeeRate / 100);

    //                 rowData.push(
    //                     outstanding.toFixed(2),
    //                     formatPercent(wair),
    //                     formatPercent(processingFeeRate),
    //                     formatPercent(totalRate)
    //                 );
    //             } else {
    //                 rowData.push('', '', '', '');
    //             }
    //         }

    //         const newRow = sheet.addRow(rowData);
    //         newRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //         applyThinBorder(newRow);
    //     }

    //     // Total Row Data
    //     const totalRowData = ['Total'];

    //     for (const month of allMonths) {
    //         let total_amount = 0, total_payment = 0, total_interest_numerator = 0, total_interest_denominator = 0, total_processing_fee_annually = 0, total_tranche_amount = 0;

    //         for (const loanType of loanTypes) {
    //             const data = groupedData[month][loanType];
    //             if (data) {
    //                 total_amount += data.total_amount;
    //                 total_payment += data.total_payment;
    //                 total_interest_numerator += data.interest_numerator;
    //                 total_interest_denominator += data.interest_denominator;
    //                 total_processing_fee_annually += data.total_processing_fee_annually;
    //                 total_tranche_amount += data.total_tranche_amount;
    //             }
    //         }

    //         const outstanding = (total_amount - total_payment) / 1e7;
    //         const wair = total_interest_denominator !== 0 ? (total_interest_numerator / total_interest_denominator) : 0;
    //         const processingFeeRate = total_tranche_amount !== 0 ? (total_processing_fee_annually / total_tranche_amount) * 100 : 0;
    //         const totalRate = wair + (processingFeeRate / 100);

    //         totalRowData.push(
    //             outstanding.toFixed(2),
    //             formatPercent(wair),
    //             formatPercent(processingFeeRate),
    //             formatPercent(totalRate)
    //         );
    //     }

    //     const totalRow = sheet.addRow(totalRowData);
    //     totalRow.font = { bold: true };
    //     totalRow.alignment = { vertical: 'middle', horizontal: 'center' };
    //     applyThinBorder(totalRow);

    //     applyThinBorder(sheet.getRow(3));
    //     applyThinBorder(sheet.getRow(4));

    //     sheet.columns.forEach(column => {
    //         column.width = 22;
    //     });

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=effective_interest_report.xlsx');

    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (err) {
    //     console.error(err);
    //     res.status(500).send('Error generating report');
    // }

    // code 150pm

    try {
        // Fetch data
        const sanctions = await sanction_details.findAll({ raw: true });
        const tranches = await tranche_details.findAll({ raw: true });
        const payments = await payment_details.findAll({ raw: true });

        // Validate
        const validateColumn = (data, requiredColumns) => {
            for (const col of requiredColumns) {
                if (!data.hasOwnProperty(col)) {
                    throw new Error(`Missing column: ${col}`);
                }
            }
        };

        const sanctionColumns = ['loan_type', 'processing_fee'];
        const trancheColumns = ['tranche_amount', 'tenure_months'];

        sanctions.forEach(sanction => validateColumn(sanction, sanctionColumns));
        tranches.forEach(tranche => validateColumn(tranche, trancheColumns));

        const sanctionMap = new Map();
        sanctions.forEach(s => sanctionMap.set(s.sanction_id, s));

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Effective Interest Report');

        const sanctionDates = sanctions.map(e => moment(e.sanction_date));
        const firstMonth = moment.min(sanctionDates).startOf('month');
        const lastMonth = moment().startOf('month');
        const allMonths = [];

        let tempMonth = lastMonth.clone();
        while (tempMonth.isSameOrAfter(firstMonth)) {
            allMonths.push(tempMonth.format('MMM-YYYY'));
            tempMonth.subtract(1, 'month');
        }

        const groupedData = {};
        const loanTypes = new Set();

        allMonths.forEach(month => {
            groupedData[month] = {};
        });

        // Calculate yearly processing fee and all tranche amount per loan_type
        const loanTypeProcessingData = {};

        for (const tranche of tranches) {
            const sanction = sanctionMap.get(tranche.sanction_id);
            if (!sanction) continue;
            const loanType = sanction.loan_type || 'Other';
            loanTypes.add(loanType);

            const trancheAmount = parseFloat(tranche.tranche_amount || 0);
            const tenureMonths = parseFloat(tranche.tenure_months || 1);
            const processingFee = parseFloat(sanction.processing_fee || 0) / 100;

            const yearlyProcessingFee = (trancheAmount * processingFee / tenureMonths) * 12;

            if (!loanTypeProcessingData[loanType]) {
                loanTypeProcessingData[loanType] = {
                    totalYearlyProcessingFee: 0,
                    totalTrancheAmount: 0,
                };
            }
            loanTypeProcessingData[loanType].totalYearlyProcessingFee += yearlyProcessingFee;
            loanTypeProcessingData[loanType].totalTrancheAmount += trancheAmount;
        }

        for (const sanction of sanctions) {
            const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
            const loanType = sanction.loan_type || 'Other';

            for (const month of allMonths) {
                if (moment(month, 'MMM-YYYY').isSameOrAfter(moment(sanctionMonth, 'MMM-YYYY'))) {
                    if (!groupedData[month][loanType]) {
                        groupedData[month][loanType] = {
                            total_amount: 0,
                            total_payment: 0,
                            interest_numerator: 0,
                            interest_denominator: 0,
                        };
                    }
                    const data = groupedData[month][loanType];
                    const sanctionAmount = parseFloat(sanction.sanction_amount || 0);

                    data.total_amount += sanctionAmount;
                }
            }
        }

        const paymentMap = new Map();
        for (const payment of payments) {
            const key = `${payment.sanction_id}_${moment(payment.payment_date).format('MMM-YYYY')}`;
            if (!paymentMap.has(key)) paymentMap.set(key, 0);
            paymentMap.set(key, paymentMap.get(key) + parseFloat(payment.payment_amount || 0));
        }

        for (const month of allMonths) {
            for (const loanType of loanTypes) {
                const data = groupedData[month][loanType];
                if (!data) continue;

                for (const sanction of sanctions.filter(s => (s.loan_type || 'Other') === loanType)) {
                    const sanctionMonth = moment(sanction.sanction_date).format('MMM-YYYY');
                    if (moment(month, 'MMM-YYYY').isBefore(moment(sanctionMonth, 'MMM-YYYY'))) continue;

                    const paymentKey = `${sanction.sanction_id}_${month}`;
                    const paymentAmount = paymentMap.get(paymentKey) || 0;

                    data.total_payment += paymentAmount;
                }

                const outstanding = data.total_amount - data.total_payment;

                if (outstanding > 0) {
                    for (const tranche of tranches) {
                        const sanction = sanctionMap.get(tranche.sanction_id);
                        if (!sanction) continue;
                        const trancheLoanType = sanction.loan_type || 'Other';
                        if (trancheLoanType !== loanType) continue;

                        const trancheDate = moment(tranche.tranche_date).format('MMM-YYYY');
                        if (moment(trancheDate, 'MMM-YYYY').isAfter(moment(month, 'MMM-YYYY'))) continue;

                        const trancheAmount = parseFloat(tranche.tranche_amount || 0);
                        const interestRate = parseFloat(tranche.interest_rate || 0);

                        const paymentKey = `${sanction.sanction_id}_${month}`;
                        const totalPaid = paymentMap.get(paymentKey) || 0;
                        const adjustedTrancheAmount = trancheAmount - totalPaid;

                        data.interest_numerator += interestRate * adjustedTrancheAmount;
                        data.interest_denominator += adjustedTrancheAmount;
                    }
                }
            }
        }

        // Prepare Excel
        sheet.getCell('A3').value = 'Loan Type';
        sheet.getCell('A3').font = { bold: true };
        sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
        sheet.getRow(3).height = 30;

        let colIndex = 2;
        for (const month of allMonths) {
            sheet.mergeCells(3, colIndex, 3, colIndex + 3);
            sheet.getCell(3, colIndex).value = month;
            sheet.getCell(3, colIndex).alignment = { vertical: 'middle', horizontal: 'center' };
            sheet.getCell(3, colIndex).font = { bold: true };
            colIndex += 4;
        }

        sheet.getRow(4).getCell(1).value = 'Loan Type';
        let subCol = 2;
        for (const month of allMonths) {
            ['Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'].forEach(label => {
                const cell = sheet.getRow(4).getCell(subCol++);
                cell.value = label;
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.font = { bold: true };
            });
        }

        const applyThinBorder = (row) => {
            row.eachCell({ includeEmpty: true }, (cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        };

        const formatPercent = value => (parseFloat(value) || 0).toFixed(2) + '%';

        for (const loanType of Array.from(loanTypes)) {
            const rowData = [loanType];

            const procData = loanTypeProcessingData[loanType] || { totalYearlyProcessingFee: 0, totalTrancheAmount: 0 };
            const avgProcessingFeeRate = procData.totalTrancheAmount !== 0
                ? (procData.totalYearlyProcessingFee / procData.totalTrancheAmount) 
                : 0;
            console.log("avg fee: ", avgProcessingFeeRate)

            for (const month of allMonths) {
                const data = groupedData[month][loanType];

                if (data) {
                    const outstanding = (data.total_amount - data.total_payment) / 1e7;
                    const wair = data.interest_denominator !== 0 ? (data.interest_numerator / data.interest_denominator) : 0;
                    const totalRate = wair + (avgProcessingFeeRate);

                    rowData.push(
                        outstanding.toFixed(2),
                        formatPercent(wair),
                        formatPercent(avgProcessingFeeRate),
                        formatPercent(totalRate)
                    );
                } else {
                    rowData.push('', '', '', '');
                }
            }

            const newRow = sheet.addRow(rowData);
            newRow.alignment = { vertical: 'middle', horizontal: 'center' };
            applyThinBorder(newRow);
        }

        const totalRowData = ['Total'];

        for (const month of allMonths) {
            let total_amount = 0, total_payment = 0, total_interest_numerator = 0, total_interest_denominator = 0;
            let totalYearlyProcessingFee = 0, totalTrancheAmount = 0;

            for (const loanType of loanTypes) {
                const data = groupedData[month][loanType];
                const procData = loanTypeProcessingData[loanType];

                if (data) {
                    total_amount += data.total_amount;
                    total_payment += data.total_payment;
                    total_interest_numerator += data.interest_numerator;
                    total_interest_denominator += data.interest_denominator;
                }
                if (procData) {
                    totalYearlyProcessingFee += procData.totalYearlyProcessingFee;
                    totalTrancheAmount += procData.totalTrancheAmount;
                }
            }

            const outstanding = (total_amount - total_payment) / 1e7;
            const wair = total_interest_denominator !== 0 ? (total_interest_numerator / total_interest_denominator) : 0;
            const avgProcessingFeeRate = totalTrancheAmount !== 0
                ? (totalYearlyProcessingFee / totalTrancheAmount)
                : 0;
            const totalRate = wair + (avgProcessingFeeRate );

            totalRowData.push(
                outstanding.toFixed(2),
                formatPercent(wair),
                formatPercent(avgProcessingFeeRate),
                formatPercent(totalRate)
            );
        }

        const totalRow = sheet.addRow(totalRowData);
        totalRow.font = { bold: true };
        totalRow.alignment = { vertical: 'middle', horizontal: 'center' };
        applyThinBorder(totalRow);

        applyThinBorder(sheet.getRow(3));
        applyThinBorder(sheet.getRow(4));

        sheet.columns.forEach(column => {
            column.width = 22;
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=effective_interest_report.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).send('Error generating report');
    }



}    