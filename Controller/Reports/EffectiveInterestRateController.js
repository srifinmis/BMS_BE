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