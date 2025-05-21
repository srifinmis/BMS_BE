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

exports.generateRundownReport = async (req, res) => {
    try {
        const data = await repayment_schedule.findAll({
            include: [
                {
                    model: lender_master,
                    as: 'lender_code_lender_master',
                    attributes: ['lender_name']
                },
                {
                    model: tranche_details,
                    as: 'tranche',
                    attributes: ['principal_start_date', 'interest_start_date']
                },
                {
                    model: sanction_details,
                    as: 'sanction',
                    attributes: ['loan_type', 'sanction_amount']
                }
            ]
        });

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Rundown Report');

        // Collect unique months
        const monthSet = new Set();
        data.forEach(entry => {
            const month = moment(entry.due_date).format('MMM-YY');
            monthSet.add(month);
        });

        const sortedMonths = Array.from(monthSet).sort((a, b) =>
            moment(a, 'MMM-YY') - moment(b, 'MMM-YY')
        );

        // Header row
        const header = ['Name of the Lender', 'Facility Type', 'Amount in Crs', 'Type', ...sortedMonths, 'Total'];
        const headerRow = sheet.addRow(header);
        headerRow.height = 45;
        headerRow.eachCell(cell => {
            cell.font = { bold: true, size: 12 };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // Group data
        const grouped = {};

        data.forEach(entry => {
            const lender = entry.lender_code_lender_master?.lender_name || 'N/A';
            const facility = entry.sanction?.loan_type || 'N/A';
            const amount = (parseFloat(entry.sanction?.sanction_amount || 0) / 1e7).toFixed(2);
            const month = moment(entry.due_date).format('MMM-YY');
            const key = `${lender}_${facility}`;

            if (!grouped[key]) {
                grouped[key] = {
                    lender,
                    facility,
                    amount,
                    principal: {},
                    interest: {}
                };
            }

            grouped[key].principal[month] = (grouped[key].principal[month] || 0) + parseFloat(entry.principal_due || 0);
            grouped[key].interest[month] = (grouped[key].interest[month] || 0) + parseFloat(entry.interest_due || 0);
        });

        const principalRows = [];
        const interestRows = [];
        const totalRows = [];

        const buildRow = (lender, facility, amount, type, valuesMap, excludeAmount = false) => {
            const row = excludeAmount ? [lender, facility, '', type] : [lender, facility, amount, type];
            let total = 0;
            sortedMonths.forEach(month => {
                const val = valuesMap[month] || 0;
                row.push(val);
                total += val;
            });
            row.push(total);
            return row;
        };

        const principalMonthTotals = {};
        const interestMonthTotals = {};
        const overallMonthTotals = {};
        let principalGrandTotal = 0;
        let interestGrandTotal = 0;
        let overallGrandTotal = 0;

        for (const key in grouped) {
            const { lender, facility, amount, principal, interest } = grouped[key];

            // Principal row
            const principalRow = buildRow(lender, facility, amount, 'Principal', principal);
            principalRows.push(principalRow);
            sortedMonths.forEach(month => {
                principalMonthTotals[month] = (principalMonthTotals[month] || 0) + (principal[month] || 0);
            });
            principalGrandTotal += Object.values(principal).reduce((a, b) => a + b, 0);

            // Interest row
            const interestRow = buildRow(lender, facility, amount, 'Interest', interest);
            interestRows.push(interestRow);
            sortedMonths.forEach(month => {
                interestMonthTotals[month] = (interestMonthTotals[month] || 0) + (interest[month] || 0);
            });
            interestGrandTotal += Object.values(interest).reduce((a, b) => a + b, 0);

            // Combined row (for Overall)
            const combined = {};
            sortedMonths.forEach(month => {
                combined[month] = (principal[month] || 0) + (interest[month] || 0);
            });
            const totalRow = buildRow(lender, facility, amount, 'Total', combined, true);
            totalRows.push(totalRow);
        }

        // Calculate overall month totals
        sortedMonths.forEach(month => {
            const principal = principalMonthTotals[month] || 0;
            const interest = interestMonthTotals[month] || 0;
            overallMonthTotals[month] = principal + interest;
            overallGrandTotal += overallMonthTotals[month];
        });

        // Add Principal Rows
        // sheet.addRow(['', '', '', 'PRINCIPAL TOTALS']).font = { bold: true };
        principalRows.forEach(row => sheet.addRow(row));

        // Add Principal Monthly Total Row
        const totalPrincipalRow = ['', '', '', 'Total'];
        sortedMonths.forEach(month => {
            totalPrincipalRow.push(principalMonthTotals[month] || 0);
        });
        totalPrincipalRow.push(principalGrandTotal);
        const totalRow = sheet.addRow(totalPrincipalRow);
        totalRow.font = { bold: true };

        sheet.addRow([]);
        sheet.addRow([]);

        // Add Interest Rows
        // sheet.addRow(['', '', '', 'INTEREST TOTALS']).font = { bold: true };
        interestRows.forEach(row => sheet.addRow(row));

        // Add Interest Monthly Total Row
        const totalInterestRow = ['', '', '', 'Total'];
        sortedMonths.forEach(month => {
            totalInterestRow.push(interestMonthTotals[month] || 0);
        });
        totalInterestRow.push(interestGrandTotal);
        const interestTotalRow = sheet.addRow(totalInterestRow);
        interestTotalRow.font = { bold: true };

        sheet.addRow([]);
        sheet.addRow([]);

        // Add Total Rows
        // sheet.addRow(['', '', '', 'OVERALL TOTALS']).font = { bold: true };
        totalRows.forEach(rowData => {
            const row = sheet.addRow(rowData);
            row.eachCell(cell => {
                cell.font = { bold: true };
            });
        });

        // Add Overall Monthly Total Row
        const overallTotalRow = ['', '', '', 'Total'];
        sortedMonths.forEach(month => {
            overallTotalRow.push(overallMonthTotals[month]);
        });
        overallTotalRow.push(overallGrandTotal);
        const overallRow = sheet.addRow(overallTotalRow);
        overallRow.font = { bold: true };

        // Auto column widths
        sheet.columns.forEach(col => {
            let maxLength = 10;
            col.eachCell({ includeEmpty: true }, cell => {
                const valLength = cell.value ? cell.value.toString().length : 0;
                maxLength = Math.max(maxLength, valLength);
            });
            col.width = maxLength + 2;
        });

        // Send Excel
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Rundown_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Failed to generate report');
    }
};
