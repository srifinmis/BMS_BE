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
    // const { fromDate, toDate } = req.body;
    // console.log("Daily backend: ", fromDate, toDate)
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

        // Apply bold and border to header
        headerRow.height = 45; // Simulate padding with increased row height

        headerRow.eachCell(cell => {
            cell.font = { bold: true, size: 12 };
            cell.alignment = { vertical: 'middle', horizontal: 'center' }; // Center align
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

        for (const key in grouped) {
            const { lender, facility, amount, principal, interest } = grouped[key];

            const addTypeRow = (type, valuesMap, excludeAmount = false) => {
                const row = excludeAmount ? [lender, facility, '', type] : [lender, facility, amount, type];
                let total = 0;
                sortedMonths.forEach(month => {
                    const val = valuesMap[month] || 0;
                    row.push(val);
                    total += val;
                });
                row.push(total);
                sheet.addRow(row);
            };

            const addTotalRow = (valuesMap) => {
                const row = ['', '', '', 'Total'];
                let total = 0;
                sortedMonths.forEach(month => {
                    const val = valuesMap[month] || 0;
                    row.push(val);
                    total += val;
                });
                row.push(total);
                const totalRow = sheet.addRow(row);

                // Apply bold font to all total row cells
                totalRow.eachCell(cell => {
                    cell.font = { bold: true };
                });
            };

            // Principal row
            addTypeRow('Principal', principal);

            // 3 empty rows
            sheet.addRow([]);
            sheet.addRow([]);
            sheet.addRow([]);

            // Principal total row (bold)
            addTotalRow(principal);

            // 3 empty rows
            sheet.addRow([]);
            sheet.addRow([]);

            // Interest row
            addTypeRow('Interest', interest);

            // 3 empty rows
            sheet.addRow([]);
            sheet.addRow([]);
            sheet.addRow([]);

            // Interest total row (bold)
            addTotalRow(interest);

            // 3 empty rows
            sheet.addRow([]);
            sheet.addRow([]);

            // Combined total (excluding Amount in Crs)
            const combined = {};
            sortedMonths.forEach(month => {
                combined[month] = (principal[month] || 0) + (interest[month] || 0);
            });

            addTypeRow('Total', combined, true); // Exclude amount

            // 1 empty row below each group
            sheet.addRow([]);
        }

        // Auto column widths
        sheet.columns.forEach(col => {
            let maxLength = 10;
            col.eachCell({ includeEmpty: true }, cell => {
                const valLength = cell.value ? cell.value.toString().length : 0;
                maxLength = Math.max(maxLength, valLength);
            });
            col.width = maxLength + 2;
        });

        // Send file
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Rundown_Report.xlsx');
        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Failed to generate report');
    }
}    