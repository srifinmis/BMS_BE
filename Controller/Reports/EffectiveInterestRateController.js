const moment = require('moment');
const express = require('express');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType } = require('docx');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

require('dotenv').config();
const models = initModels(sequelize);
const { lender_master, tranche_details, repayment_schedule, sanction_details, payment_details } = models;


exports.generateEffectiveInterestRateReport = async (req, res) => {
    const { fromDate } = req.body;
    console.log("Daily effective backend: ", fromDate)

    // try {

    //     const selectedDate = new Date(fromDate);
    //     const startOfDay = new Date(selectedDate);
    //     startOfDay.setHours(0, 0, 0, 0);

    //     const endOfDay = new Date(selectedDate);
    //     endOfDay.setHours(23, 59, 59, 999);

    //     if (isNaN(selectedDate)) {
    //         return res.status(400).json({ error: 'Invalid date range provided' });
    //     }

    //     const whereClause = {
    //         createdat: { [Op.between]: [startOfDay, endOfDay] }
    //     };
    //     const data = await repayment_schedule.findAll({
    //         where: whereClause,
    //         include: [
    //             {
    //                 model: sanction_details,
    //                 as: 'sanction',
    //                 attributes: ['sanction_amount', 'sanction_date'],
    //                 include: [
    //                     {
    //                         model: lender_master,
    //                         as: 'lender_code_lender_master',
    //                         attributes: ['lender_code']
    //                     }
    //                 ]
    //             },
    //             {
    //                 model: tranche_details,
    //                 as: 'tranche',
    //                 attributes: [
    //                     'tranche_date',
    //                     'tranche_amount',
    //                 ]
    //             }
    //         ],
    //         // ...(orderClause && { order: orderClause }),
    //         raw: true
    //     });


    //     console.log("daily fetching: ", data)

    //     if (!data || data.length === 0) {
    //         return res.status(404).json({ message: 'No records found for the selected filters.' });
    //     }
    //     // console.log("roc data backend: ", data)

    //     const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
    //     const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
    //     const today = new Date().toLocaleDateString('en-GB');
    //     const REPORT_TITLE = process.env.REPORT_TITLE || `Report: Effective Interest Date As on ${today}`;

    //     const headerInfo = [
    //         ORG_NAME,
    //         '',
    //         ORG_ADDRESS,
    //         '',
    //         REPORT_TITLE,
    //         ''
    //     ];

    //     const columns = [
    //         { header: 'Loan Type', key: 'loan', width: 20 },
    //         { header: 'Out-standing Amount (In ₹)', key: 'out', width: 20 },
    //         { header: 'Weighted Avg. Interest Rate', key: '', width: 25 },
    //         { header: 'Avg. Processing Fee Rate', key: '', width: 20 },
    //         { header: 'Total Rate', key: '' },
    //     ];

    //     // === EXCEL FORMAT ===
    //     // if (format === 'excel') {
    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest Rate Report');

    //     const totalCols = columns.length;

    //     // === Add Organization Name ===
    //     const orgNameRow = sheet.addRow([ORG_NAME]);
    //     sheet.mergeCells(`A${orgNameRow.number}:` + String.fromCharCode(65 + totalCols - 1) + `${orgNameRow.number}`);
    //     orgNameRow.font = { bold: true, size: 14 };
    //     orgNameRow.alignment = { vertical: 'middle', horizontal: 'center' };

    //     // === Spacer Row === (After Organization Name)
    //     sheet.addRow([]);

    //     // === Add Address ===
    //     const addressRow = sheet.addRow([ORG_ADDRESS]);
    //     sheet.mergeCells(`A${addressRow.number}:` + String.fromCharCode(65 + totalCols - 1) + `${addressRow.number}`);
    //     addressRow.font = { bold: true, size: 12 };
    //     addressRow.alignment = { vertical: 'middle', horizontal: 'center' };

    //     // === Spacer Row === (After Address)
    //     sheet.addRow([]);

    //     // === Add Report Date ===
    //     const dateRow = sheet.addRow([REPORT_TITLE]);
    //     sheet.mergeCells(`A${dateRow.number}:` + String.fromCharCode(65 + totalCols - 1) + `${dateRow.number}`);
    //     dateRow.font = { bold: true, size: 12 };
    //     dateRow.alignment = { vertical: 'middle', horizontal: 'center' };

    //     // === Spacer Row === (After Report Date)
    //     sheet.addRow([]);

    //     // === Add Table Header Row ===
    //     const headerRow = sheet.addRow(columns.map(col => col.header));
    //     headerRow.font = { bold: true };
    //     headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    //     // === Add Borders to Header Row ===
    //     columns.forEach((col, index) => {
    //         const cell = sheet.getCell(`${String.fromCharCode(65 + index)}${headerRow.number}`);
    //         cell.border = {
    //             top: { style: 'thin' },
    //             left: { style: 'thin' },
    //             bottom: { style: 'thin' },
    //             right: { style: 'thin' }
    //         };
    //     });

    //     // === Add Table Data Rows ===
    //     data.forEach(row => {
    //         const rowValues = columns.map(col => {
    //             const keys = col.key.split('.');
    //             return keys.length === 2
    //                 ? row[`${keys[0]}.${keys[1]}`] || row[keys.join('.')]
    //                 : row[col.key];
    //         });

    //         const dataRow = sheet.addRow(rowValues);
    //         // console.log("excel data: ", dataRow)

    //         // === Add Borders to Data Rows ===
    //         rowValues.forEach((_, index) => {
    //             const cell = sheet.getCell(`${String.fromCharCode(65 + index)}${dataRow.number}`);
    //             cell.border = {
    //                 top: { style: 'thin' },
    //                 left: { style: 'thin' },
    //                 bottom: { style: 'thin' },
    //                 right: { style: 'thin' }
    //             };
    //         });
    //     });

    //     // === Set Column Widths ===
    //     columns.forEach((col, i) => {
    //         sheet.getColumn(i + 1).width = col.width || 20;
    //     });

    //     // === Finalize and Send ===
    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=Effective_Interest_Rate_Report.xlsx');
    //     await workbook.xlsx.write(res);
    //     res.end();
    //     // }
    //     // === INVALID FORMAT ===
    //     // else {
    //     //     res.status(400).json({ error: 'Invalid format selected' });
    //     // }
    // } catch (error) {
    //     console.error('Error generating report:', error);
    //     res.status(500).send('Server Error');
    // }
    // try {
    //     const selectedDate = new Date(fromDate);
    //     if (isNaN(selectedDate)) {
    //         return res.status(400).json({ error: 'Invalid date provided' });
    //     }

    //     const startOfMonth = new Date(selectedDate.getFullYear(), selectedDate.getMonth(), 1);
    //     const endOfMonth = new Date(selectedDate.getFullYear(), selectedDate.getMonth() + 1, 0, 23, 59, 59, 999);

    //     const whereClause = {
    //         '$sanction.sanction_date$': { [Op.between]: [startOfMonth, endOfMonth] }
    //     };
    //     console.log("where clause: ", whereClause)

    //     const data = await repayment_schedule.findAll({
    //         where: whereClause,
    //         include: [
    //             {
    //                 model: sanction_details,
    //                 as: 'sanction',
    //                 attributes: ['sanction_amount', 'sanction_date', 'loan_type', 'processing_fee'],
    //                 include: [
    //                     {
    //                         model: lender_master,
    //                         as: 'lender_code_lender_master',
    //                         attributes: ['lender_code']
    //                     }
    //                 ]
    //             },
    //             {
    //                 model: tranche_details,
    //                 as: 'tranche',
    //                 attributes: ['tranche_date', 'tranche_amount', 'interest_rate']
    //             }
    //         ],
    //         raw: true
    //     });

    //     console.log("Fetched repayment_schedule data:", data);

    //     if (!data || data.length === 0) {
    //         return res.status(404).json({ message: 'No records found for the selected date.' });
    //     }

    //     const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
    //     const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Hyderabad.';
    //     const today = new Date().toLocaleDateString('en-GB');
    //     const REPORT_TITLE = process.env.REPORT_TITLE || `Report: Effective Interest Rate as on ${today}`;

    //     const headerInfo = [ORG_NAME, '', ORG_ADDRESS, '', REPORT_TITLE, ''];

    //     const columns = [
    //         { header: 'Loan Type', key: 'loan', width: 20 },
    //         { header: 'Out-standing Amount (In ₹)', key: 'out', width: 25 },
    //         { header: 'Weighted Avg. Interest Rate', key: 'wair', width: 25 },
    //         { header: 'Avg. Processing Fee Rate', key: 'processingFee', width: 25 },
    //         { header: 'Total Rate', key: 'totalRate', width: 20 },
    //     ];

    //     // ==== Group and Calculate ====
    //     const groupedData = {};

    //     for (const row of data) {
    //         const loanType = row['sanction.loan_type'] || 'Other';
    //         const sanctionAmount = parseFloat(row['sanction.sanction_amount'] || 0);
    //         const trancheAmount = parseFloat(row['tranche.tranche_amount'] || 0);
    //         const interestRate = parseFloat(row['tranche.interest_rate'] || 0) / 100; // assuming interest_rate in %.
    //         const processingFee = parseFloat(row['sanction.processing_fee'] || 0) / 100; // assuming processing fee in %.

    //         if (!groupedData[loanType]) {
    //             groupedData[loanType] = {
    //                 totalSanctionAmount: 0,
    //                 totalOutstanding: 0,
    //                 interestNumerator: 0,
    //                 interestDenominator: 0,
    //                 totalProcessingFeeAmount: 0,
    //                 totalTrancheAmount: 0
    //             };
    //         }

    //         groupedData[loanType].totalSanctionAmount += sanctionAmount;
    //         groupedData[loanType].totalOutstanding += trancheAmount; // outstanding = tranche amounts
    //         groupedData[loanType].interestNumerator += (trancheAmount * interestRate);
    //         groupedData[loanType].interestDenominator += trancheAmount;
    //         groupedData[loanType].totalProcessingFeeAmount += (trancheAmount * processingFee);
    //         groupedData[loanType].totalTrancheAmount += trancheAmount;
    //     }

    //     // === EXCEL EXPORT ===
    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest Rate Report');

    //     const totalCols = columns.length;

    //     // Add organization header
    //     for (const info of headerInfo) {
    //         const row = sheet.addRow([info]);
    //         sheet.mergeCells(`A${row.number}:${String.fromCharCode(65 + totalCols - 1)}${row.number}`);
    //         row.font = { bold: true, size: info ? 14 : 12 };
    //         row.alignment = { vertical: 'middle', horizontal: 'center' };
    //     }

    //     sheet.addRow([]);

    //     const headerRow = sheet.addRow(columns.map(c => c.header));
    //     headerRow.font = { bold: true };
    //     headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    //     columns.forEach((col, idx) => {
    //         sheet.getColumn(idx + 1).width = col.width;
    //         const cell = sheet.getCell(`${String.fromCharCode(65 + idx)}${headerRow.number}`);
    //         cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    //     });

    //     // === Add Data Rows ===
    //     for (const [loanType, info] of Object.entries(groupedData)) {
    //         const wair = info.interestDenominator ? (info.interestNumerator / info.interestDenominator) : 0;
    //         const avgProcessingFeeRate = info.totalTrancheAmount ? (info.totalProcessingFeeAmount / info.totalTrancheAmount) : 0;
    //         const totalRate = wair + avgProcessingFeeRate;

    //         const row = [
    //             loanType,
    //             info.totalOutstanding.toFixed(2),
    //             (wair * 100).toFixed(2) + '%',
    //             (avgProcessingFeeRate * 100).toFixed(2) + '%',
    //             (totalRate * 100).toFixed(2) + '%'
    //         ];
    //         const dataRow = sheet.addRow(row);

    //         row.forEach((_, idx) => {
    //             const cell = sheet.getCell(`${String.fromCharCode(65 + idx)}${dataRow.number}`);
    //             cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    //         });
    //     }

    //     // Send file
    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=Effective_Interest_Rate_Report.xlsx');
    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error('Error generating report:', error);
    //     res.status(500).send('Server Error');
    // }
    //////////////////////////////////////////

    // try {
    //     const { fromDate } = req.body;
    //     if (!fromDate) {
    //         return res.status(400).json({ error: 'fromDate is required' });
    //     }

    //     const selectedMonth = moment(fromDate, 'YYYY-MM-DD');
    //     if (!selectedMonth.isValid()) {
    //         return res.status(400).json({ error: 'Invalid date provided' });
    //     }

    //     const startOfMonth = selectedMonth.startOf('month').toDate();
    //     const endOfMonth = selectedMonth.endOf('month').toDate();
    //     const selectedMonthStr = selectedMonth.format('MMM-YYYY');

    //     // Fetch necessary data
    //     const sanctions = await sanction_details.findAll({
    //         where: {
    //             sanction_date: { [Op.between]: [startOfMonth, endOfMonth] }
    //         },
    //         raw: true
    //     });

    //     const sanctionIds = sanctions.map(s => s.sanction_id);

    //     const tranches = await tranche_details.findAll({
    //         where: {
    //             sanction_id: sanctionIds.length > 0 ? { [Op.in]: sanctionIds } : undefined
    //         },
    //         raw: true
    //     });

    //     const payments = await payment_details.findAll({
    //         where: {
    //             sanction_id: sanctionIds.length > 0 ? { [Op.in]: sanctionIds } : undefined,
    //             payment_date: { [Op.between]: [startOfMonth, endOfMonth] }
    //         },
    //         raw: true
    //     });

    //     if (sanctions.length === 0) {
    //         return res.status(404).json({ message: 'No sanction records found for the selected month.' });
    //     }

    //     // Validate
    //     const validateColumn = (data, requiredColumns) => {
    //         for (const col of requiredColumns) {
    //             if (!data.hasOwnProperty(col)) {
    //                 throw new Error(`Missing column: ${col}`);
    //             }
    //         }
    //     };

    //     const sanctionColumns = ['loan_type', 'processing_fee'];
    //     const trancheColumns = ['tranche_amount', 'tenure_months'];

    //     sanctions.forEach(sanction => validateColumn(sanction, sanctionColumns));
    //     tranches.forEach(tranche => validateColumn(tranche, trancheColumns));

    //     const sanctionMap = new Map();
    //     sanctions.forEach(s => sanctionMap.set(s.sanction_id, s));

    //     const workbook = new ExcelJS.Workbook();
    //     const sheet = workbook.addWorksheet('Effective Interest Report');

    //     const groupedData = {};
    //     const loanTypes = new Set();

    //     // Calculate yearly processing fee and tranche amount per loan_type
    //     const loanTypeProcessingData = {};

    //     for (const tranche of tranches) {
    //         const sanction = sanctionMap.get(tranche.sanction_id);
    //         if (!sanction) continue;
    //         const loanType = sanction.loan_type || 'Other';
    //         loanTypes.add(loanType);

    //         const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //         const tenureMonths = parseFloat(tranche.tenure_months || 1);
    //         const processingFee = parseFloat(sanction.processing_fee || 0) / 100;

    //         const yearlyProcessingFee = (trancheAmount * processingFee / tenureMonths) * 12;

    //         if (!loanTypeProcessingData[loanType]) {
    //             loanTypeProcessingData[loanType] = {
    //                 totalYearlyProcessingFee: 0,
    //                 totalTrancheAmount: 0,
    //             };
    //         }
    //         loanTypeProcessingData[loanType].totalYearlyProcessingFee += yearlyProcessingFee;
    //         loanTypeProcessingData[loanType].totalTrancheAmount += trancheAmount;
    //     }

    //     // Grouped outstanding and interest calculation
    //     for (const sanction of sanctions) {
    //         const loanType = sanction.loan_type || 'Other';
    //         loanTypes.add(loanType);

    //         if (!groupedData[loanType]) {
    //             groupedData[loanType] = {
    //                 total_amount: 0,
    //                 total_payment: 0,
    //                 interest_numerator: 0,
    //                 interest_denominator: 0,
    //             };
    //         }
    //         const data = groupedData[loanType];
    //         data.total_amount += parseFloat(sanction.sanction_amount || 0);
    //     }

    //     const paymentMap = new Map();
    //     for (const payment of payments) {
    //         const key = payment.sanction_id;
    //         if (!paymentMap.has(key)) paymentMap.set(key, 0);
    //         paymentMap.set(key, paymentMap.get(key) + parseFloat(payment.payment_amount || 0));
    //     }

    //     for (const loanType of loanTypes) {
    //         const data = groupedData[loanType];
    //         if (!data) continue;

    //         for (const sanction of sanctions.filter(s => (s.loan_type || 'Other') === loanType)) {
    //             const sanctionId = sanction.sanction_id;
    //             const paymentAmount = paymentMap.get(sanctionId) || 0;

    //             data.total_payment += paymentAmount;
    //         }

    //         const outstanding = data.total_amount - data.total_payment;

    //         if (outstanding > 0) {
    //             for (const tranche of tranches) {
    //                 const sanction = sanctionMap.get(tranche.sanction_id);
    //                 if (!sanction) continue;
    //                 const trancheLoanType = sanction.loan_type || 'Other';
    //                 if (trancheLoanType !== loanType) continue;

    //                 const trancheDate = moment(tranche.tranche_date).format('MMM-YYYY');
    //                 if (trancheDate !== selectedMonthStr) continue; // Only current selected month tranches

    //                 const trancheAmount = parseFloat(tranche.tranche_amount || 0);
    //                 const interestRate = parseFloat(tranche.interest_rate || 0);

    //                 const paymentAmount = paymentMap.get(sanction.sanction_id) || 0;
    //                 const adjustedTrancheAmount = trancheAmount - paymentAmount;

    //                 data.interest_numerator += interestRate * adjustedTrancheAmount;
    //                 data.interest_denominator += adjustedTrancheAmount;
    //             }
    //         }
    //     }

    //     // Excel Header
    //     const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
    //     const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Hyderabad.';
    //     const today = new Date().toLocaleDateString('en-GB');
    //     const REPORT_TITLE = process.env.REPORT_TITLE || `Report: Effective Interest Rate as on ${fromDate}`;

    //     const headerInfo = [ORG_NAME, '', ORG_ADDRESS, '', REPORT_TITLE, ''];

    //     let headerRow = 1;
    //     headerInfo.forEach((line) => {
    //         sheet.getRow(headerRow).getCell(1).value = line;
    //         sheet.mergeCells(`A${headerRow}:E${headerRow}`);
    //         sheet.getRow(headerRow).getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    //         headerRow++;
    //     });

    //     const columns = [
    //         { header: 'Loan Type', key: 'loan', width: 30 },
    //         { header: 'Outstanding Amount (In ₹)', key: 'out', width: 25 },
    //         { header: 'Weighted Avg. Interest Rate', key: 'wair', width: 25 },
    //         { header: 'Avg. Processing Fee Rate', key: 'processingFee', width: 25 },
    //         { header: 'Total Rate', key: 'totalRate', width: 20 },
    //     ];
    //     sheet.columns = columns;

    //     // Thin border
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
    //         const data = groupedData[loanType];
    //         if (!data) continue;

    //         const procData = loanTypeProcessingData[loanType] || { totalYearlyProcessingFee: 0, totalTrancheAmount: 0 };

    //         const outstanding = (data.total_amount - data.total_payment).toFixed(2);
    //         const wair = data.interest_denominator !== 0 ? (data.interest_numerator / data.interest_denominator) : 0;
    //         const avgProcessingFeeRate = procData.totalTrancheAmount !== 0
    //             ? (procData.totalYearlyProcessingFee / procData.totalTrancheAmount)
    //             : 0;
    //         const totalRate = wair + avgProcessingFeeRate;

    //         const row = sheet.addRow([
    //             loanType,
    //             outstanding,
    //             formatPercent(wair),
    //             formatPercent(avgProcessingFeeRate),
    //             formatPercent(totalRate)
    //         ]);
    //         applyThinBorder(row);
    //         row.alignment = { vertical: 'middle', horizontal: 'center' };
    //     }

    //     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    //     res.setHeader('Content-Disposition', 'attachment; filename=effective_interest_report.xlsx');
    //     await workbook.xlsx.write(res);
    //     res.end();

    // } catch (error) {
    //     console.error(error);
    //     res.status(500).send('Error generating report');
    // }

    // above formate missing

    try {
        const { fromDate } = req.body;
        if (!fromDate) {
            return res.status(400).json({ error: 'fromDate is required' });
        }

        const selectedMonth = moment(fromDate, 'YYYY-MM-DD');
        if (!selectedMonth.isValid()) {
            return res.status(400).json({ error: 'Invalid date provided' });
        }

        const startOfMonth = selectedMonth.startOf('month').toDate();
        const endOfMonth = selectedMonth.endOf('month').toDate();
        const selectedMonthStr = selectedMonth.format('MMM-YYYY');

        // Fetch necessary data
        const sanctions = await sanction_details.findAll({
            where: {
                sanction_date: { [Op.between]: [startOfMonth, endOfMonth] }
            },
            raw: true
        });

        const sanctionIds = sanctions.map(s => s.sanction_id);

        const tranches = await tranche_details.findAll({
            where: {
                sanction_id: sanctionIds.length > 0 ? { [Op.in]: sanctionIds } : undefined
            },
            raw: true
        });

        const payments = await payment_details.findAll({
            where: {
                sanction_id: sanctionIds.length > 0 ? { [Op.in]: sanctionIds } : undefined,
                payment_date: { [Op.between]: [startOfMonth, endOfMonth] }
            },
            raw: true
        });

        if (sanctions.length === 0) {
            return res.status(404).json({ message: 'No sanction records found for the selected month.' });
        }

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

        const groupedData = {};
        const loanTypes = new Set();

        // Calculate yearly processing fee and tranche amount per loan_type
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

        // Grouped outstanding and interest calculation
        for (const sanction of sanctions) {
            const loanType = sanction.loan_type || 'Other';
            loanTypes.add(loanType);

            if (!groupedData[loanType]) {
                groupedData[loanType] = {
                    total_amount: 0,
                    total_payment: 0,
                    interest_numerator: 0,
                    interest_denominator: 0,
                };
            }
            const data = groupedData[loanType];
            data.total_amount += parseFloat(sanction.sanction_amount || 0);
        }

        const paymentMap = new Map();
        for (const payment of payments) {
            const key = payment.sanction_id;
            if (!paymentMap.has(key)) paymentMap.set(key, 0);
            paymentMap.set(key, paymentMap.get(key) + parseFloat(payment.payment_amount || 0));
        }

        for (const loanType of loanTypes) {
            const data = groupedData[loanType];
            if (!data) continue;

            for (const sanction of sanctions.filter(s => (s.loan_type || 'Other') === loanType)) {
                const sanctionId = sanction.sanction_id;
                const paymentAmount = paymentMap.get(sanctionId) || 0;

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
                    if (trancheDate !== selectedMonthStr) continue; // Only current selected month tranches

                    const trancheAmount = parseFloat(tranche.tranche_amount || 0);
                    const interestRate = parseFloat(tranche.interest_rate || 0);

                    const paymentAmount = paymentMap.get(sanction.sanction_id) || 0;
                    const adjustedTrancheAmount = trancheAmount - paymentAmount;

                    data.interest_numerator += interestRate * adjustedTrancheAmount;
                    data.interest_denominator += adjustedTrancheAmount;
                }
            }
        }

        // Excel Header
        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Hyderabad.';
        const today = new Date().toLocaleDateString('en-GB');
        const REPORT_TITLE = process.env.REPORT_TITLE || `Report: Effective Interest Rate as on ${fromDate}`;

        const headerInfo = [ORG_NAME, '', ORG_ADDRESS, '', REPORT_TITLE, ''];

        let headerRow = 1;
        headerInfo.forEach((line) => {
            sheet.getRow(headerRow).getCell(1).value = line;
            sheet.mergeCells(`A${headerRow}:E${headerRow}`);
            sheet.getRow(headerRow).getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
            headerRow++;
        });

        // Add new row for column headers
        const headers = ['Loan Type', 'Outstanding Amount (in Crs.)', 'Weighted Avg. Interest Rate', 'Avg. Processing Fee Rate', 'Total Rate'];
        let rowHeaders = sheet.getRow(7);
        rowHeaders.values = headers;
        rowHeaders.eachCell((cell, colNumber) => {
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            sheet.getColumn(colNumber).width = 25;
        });

        // Apply thin border to all cells in a row
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

        // Loop through and add the data rows below the headers (starting from row 8)
        let rowNum = 8;
        let totalOutstanding = 0;
        let totalInterestRate = 0;
        let totalProcessingFeeRate = 0;
        let totalRate = 0;

        for (const loanType of Array.from(loanTypes)) {
            const data = groupedData[loanType];
            if (!data) continue;

            const procData = loanTypeProcessingData[loanType] || { totalYearlyProcessingFee: 0, totalTrancheAmount: 0 };

            const outstanding = (data.total_amount - data.total_payment).toFixed(2);
            const wair = data.interest_denominator !== 0 ? (data.interest_numerator / data.interest_denominator) : 0;
            const avgProcessingFeeRate = procData.totalTrancheAmount !== 0
                ? (procData.totalYearlyProcessingFee / procData.totalTrancheAmount)
                : 0;
            const totalRow = wair + avgProcessingFeeRate;

            totalOutstanding += parseFloat(outstanding);
            totalInterestRate += wair;
            totalProcessingFeeRate += avgProcessingFeeRate;
            totalRate += totalRow;

            const row = sheet.addRow([
                loanType,
                outstanding,
                formatPercent(wair),
                formatPercent(avgProcessingFeeRate),
                formatPercent(totalRow)
            ]);
            applyThinBorder(row); // Apply thin border to this row
            row.alignment = { vertical: 'middle', horizontal: 'center' };
        }

        // Add the total row
        const totalRow = sheet.addRow([
            'Total',
            totalOutstanding.toFixed(2),
            formatPercent(totalInterestRate),
            formatPercent(totalProcessingFeeRate),
            formatPercent(totalRate)
        ]);
        totalRow.font = { bold: true };
        applyThinBorder(totalRow); // Apply thin border to the total row
        totalRow.alignment = { vertical: 'middle', horizontal: 'center' };

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=effective_interest_report.xlsx');
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error(error);
        res.status(500).send('Error generating report');
    }



};