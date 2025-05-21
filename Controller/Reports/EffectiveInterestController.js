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