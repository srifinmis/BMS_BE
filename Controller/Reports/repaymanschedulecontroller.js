const express = require('express');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType } = require('docx');
const { Op, fn, col } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

require('dotenv').config();
const models = initModels(sequelize);
const { lender_master, sanction_details, tranche_details, repayment_schedule, payment_details } = models;

exports.generateRepaymentScheduleReport = async (req, res) => {
    const { banks = "all", sanctions = "all", tranches = "all", format = "excel" } = req.body;
    console.log("Repayment Schedule backend: ", banks, sanctions, tranches, format)

    try {
        const whereClause = {};
        if (banks !== 'all') {
            whereClause.lender_code = { [Op.in]: banks };
        }
        if (sanctions !== 'all') {
            whereClause['$sanction.sanction_id$'] = { [Op.in]: sanctions };
        }
        if (tranches !== 'all') {
            whereClause.tranche_id = { [Op.in]: tranches };
        }

        const data = await repayment_schedule.findAll({
            where: whereClause,
            include: [
                {
                    model: tranche_details,
                    as: 'tranche',
                    attributes: ['tranche_id', "tranche_amount"],
                },
                {
                    model: sanction_details,
                    as: 'sanction',
                    attributes: ['sanction_id'],
                    include: [
                        {
                            model: lender_master,
                            as: 'lender_code_lender_master',
                            attributes: ['lender_code', 'lender_name']
                        }
                    ]
                }
            ],
            raw: true
        });
        // Fetch all relevant payments for the due dates
        const trancheIds = [...new Set(data.map(d => d.tranche_id))];
        const dueDates = [...new Set(
            data
                .filter(d => d.due_date)
                .map(d => {
                    const dt = new Date(d.due_date);
                    return isNaN(dt) ? null : dt.toISOString().split('T')[0];
                })
                .filter(Boolean) // remove nulls
            )];


        const payments = await payment_details.findAll({
            where: {
                tranche_id: { [Op.in]: trancheIds },
                payment_date: { [Op.in]: dueDates }
            },
            attributes: [
                'tranche_id',
                'payment_date',
                [fn('SUM', col('payment_amount')), 'total_payment']
            ],
            group: ['tranche_id', 'payment_date'],
            raw: true
        });

        // Create a lookup map
        const paymentMap = {};
        payments.forEach(p => {
            const paymentDate = new Date(p.payment_date);
            if (!isNaN(paymentDate)) {
                const key = `${p.tranche_id}_${new Date(paymentDate).toISOString().split('T')[0]}`;
                paymentMap[key] = parseFloat(p.total_payment || 0);
            }
        });

        // Append Loan Outstanding to each row
        const result = data.map(row => {
            const dueDateStr = row.due_date ? new Date(row.due_date).toISOString().split('T')[0] : '';
            const key = `${row.tranche_id}_${dueDateStr}`;
            const paymentMade = paymentMap[key] || 0;
            const loanOutstanding = parseFloat(row['tranche.tranche_amount'] || 0) - paymentMade;

            return {
                ...row,
                loan_outstanding: loanOutstanding
            };
        });



        console.log("Repayment Schedule fetching: ", result)

        if (!result || result.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }
        // console.log("roc data backend: ", data)

        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
        const REPORT_TITLE = process.env.REPORT_TITLE || 'Repayment Schedule';
        const today = new Date().toLocaleDateString('en-GB');

        const headerInfo = [
            ORG_NAME,
            '',
            ORG_ADDRESS,
            '',
            REPORT_TITLE,
            ''
        ];

        const columns = [
            { header: 'Due Date', key: 'due_date', width: 20 },
            { header: 'Loan Outstanding (In Rs)', key: 'loan_outstanding', width: 25 },
            { header: 'Principal Due (In Rs)', key: 'principal_due', width: 25 },
            { header: 'No. of Days', key: 'interest_days', width: 20 },
            { header: 'Rate Of Interest', key: 'interest_rate', width: 20 },
            { header: 'Interest Amount (In Rs)', key: 'interest_due', width: 25 },
            { header: 'Total Due (In Rs)', key: 'total_due', width: 25 },
        ];

        // === EXCEL FORMAT ===
        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Repayment Schedule');

            const totalCols = columns.length;

            // === Add Organization Name ===
            const orgNameRow = sheet.addRow([ORG_NAME]);
            sheet.mergeCells(`A${orgNameRow.number}:` + String.fromCharCode(65 + totalCols - 1) + `${orgNameRow.number}`);
            orgNameRow.font = { bold: true, size: 14 };
            orgNameRow.alignment = { vertical: 'middle', horizontal: 'center' };

            // === Spacer Row === (After Organization Name)
            sheet.addRow([]);

            // === Add Address ===
            const addressRow = sheet.addRow([ORG_ADDRESS]);
            sheet.mergeCells(`A${addressRow.number}:` + String.fromCharCode(65 + totalCols - 1) + `${addressRow.number}`);
            addressRow.font = { bold: true, size: 12 };
            addressRow.alignment = { vertical: 'middle', horizontal: 'center' };

            // === Spacer Row === (After Address)
            sheet.addRow([]);

            // === Add Report Date ===
            const dateRow = sheet.addRow([REPORT_TITLE]);
            sheet.mergeCells(`A${dateRow.number}:` + String.fromCharCode(65 + totalCols - 1) + `${dateRow.number}`);
            dateRow.font = { bold: true, size: 12 };
            dateRow.alignment = { vertical: 'middle', horizontal: 'center' };

            // === Spacer Row === (After Report Date)
            sheet.addRow([]);

            // === Add Table Header Row ===
            const headerRow = sheet.addRow(columns.map(col => col.header));
            headerRow.font = { bold: true };
            headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

            // === Add Borders to Header Row ===
            columns.forEach((col, index) => {
                const cell = sheet.getCell(`${String.fromCharCode(65 + index)}${headerRow.number}`);
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // === Add Table Data Rows ===
            result.forEach(row => {
                const rowValues = columns.map(col => {
                    const keys = col.key.split('.');
                    return keys.length === 2
                        ? row[`${keys[0]}.${keys[1]}`] || row[keys.join('.')]
                        : row[col.key];
                });

                const dataRow = sheet.addRow(rowValues);
                // console.log("excel data: ", dataRow)

                // === Add Borders to Data Rows ===
                rowValues.forEach((_, index) => {
                    const cell = sheet.getCell(`${String.fromCharCode(65 + index)}${dataRow.number}`);
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            });

            // === Set Column Widths ===
            columns.forEach((col, i) => {
                sheet.getColumn(i + 1).width = col.width || 20;
            });

            // === Finalize and Send ===
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', 'attachment; filename=Repayment_Schedule_Report.xlsx');
            await workbook.xlsx.write(res);
            res.end();
        }
        // === PDF FORMAT ===
        else if (format === 'pdf') {
            const PDFDocument = require('pdfkit');
            const doc = new PDFDocument({ margin: 20, size: 'A4', layout: 'landscape' });

            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename=Repayment_Schedule_Report.pdf');
            doc.pipe(res);

            const headerWidth = 500;
            const pageCenter = doc.page.width / 2;
            const headerX = pageCenter - headerWidth / 2;

            // Centered header info
            headerInfo.forEach(line => {
                doc.fontSize(12).text(line, headerX, doc.y, {
                    width: headerWidth,
                    align: 'center'
                });
                doc.moveDown(0.5);
            });
            doc.moveDown(2);

            const pageWidth = doc.page.width;
            const pageMargins = doc.page.margins.left + doc.page.margins.right;
            const availableWidth = pageWidth - pageMargins;
            const padding = 2;
            const rowHeight = 40;
            const charWidth = 6;

            // Step 1: Calculate natural column widths
            let naturalWidths = columns.map(col => {
                const headerLen = col.header.length;
                const maxDataLen = Math.max(...data.map(row => {
                    const keys = col.key.split('.');
                    const value = keys.length === 2 ? row[`${keys[0]}.${keys[1]}`] || row[keys.join('.')] : row[col.key];
                    return `${value ?? ''}`.length;
                }));
                const maxLen = Math.max(headerLen, maxDataLen);
                return maxLen * charWidth + padding * 2;
            });

            // Step 2: Scale if total width exceeds available space
            const totalNaturalWidth = naturalWidths.reduce((sum, w) => sum + w, 0);
            let columnWidths = [...naturalWidths];

            if (totalNaturalWidth > availableWidth) {
                const scale = availableWidth / totalNaturalWidth;
                columnWidths = naturalWidths.map(w => w * scale);
            }

            const startX = doc.page.margins.left;
            let y = doc.y;

            // Step 3: Draw table headers
            let x = startX;
            columns.forEach((col, i) => {
                doc.rect(x, y, columnWidths[i], rowHeight).stroke();
                doc.font('Helvetica-Bold').fontSize(10).text(col.header, x + padding, y + 6, {
                    width: columnWidths[i] - padding * 2,
                    align: 'center'
                });
                x += columnWidths[i];
            });

            y += rowHeight;

            // Step 4: Draw data rows
            result.forEach(row => {
                x = startX;
                columns.forEach((col, i) => {
                    const keys = col.key.split('.');
                    const value = keys.length === 2 ? row[`${keys[0]}.${keys[1]}`] || row[keys.join('.')] : row[col.key];
                    const cellText = `${value ?? ''}`;
                    doc.rect(x, y, columnWidths[i], rowHeight).stroke();
                    doc.font('Helvetica').fontSize(7).text(cellText, x + padding, y + 6, {
                        width: columnWidths[i] - padding * 2,
                        align: 'center'
                    });
                    x += columnWidths[i];
                });

                y += rowHeight;
                if (y + rowHeight > doc.page.height - doc.page.margins.bottom) {
                    doc.addPage();
                    y = doc.y;
                }
            });

            doc.end();
        }

        // === WORD FORMAT ===
        else if (format === 'word') {
            // Helper function to get value from nested keys
            function getValue(obj, path) {
                if (!path) return ''; // Return empty string if no path is provided
                if (obj[path] !== undefined) return obj[path]; // Handle direct keys
                return path.split('.').reduce((acc, part) => acc && acc[part], obj) ?? ''; // Handle nested keys
            }

            // Table header row (static, just column names)
            const tableRows = [
                new TableRow({
                    children: columns.map(col =>
                        new TableCell({
                            children: [new Paragraph({
                                text: col.header,
                                bold: true
                            })],
                            width: { size: 100 / columns.length, type: WidthType.PERCENTAGE }  // Width as percentage of total document width
                        })
                    )
                })
            ];

            // Add data rows
            result.forEach((row, index) => {
                // Debugging: Check the row object and the columns keys

                const cells = columns.map(col => {
                    const value = getValue(row, col.key);

                    // Ensure that the value is converted to a string if it's a number or other types
                    return new TableCell({
                        children: [new Paragraph({
                            text: String(value ?? '')  // Ensure null/undefined are converted to an empty string
                        })],
                        width: { size: 100 / columns.length, type: WidthType.PERCENTAGE } // Width as percentage of total document width
                    });
                });

                // Push the row with the constructed cells
                tableRows.push(new TableRow({ children: cells }));
            });

            // Create Word document
            const doc = new Document({
                sections: [{
                    children: [
                        // Add header information (organization name, address, etc.)
                        ...headerInfo.map(line => new Paragraph({ text: line, alignment: AlignmentType.CENTER })),

                        // Add a spacer paragraph
                        new Paragraph({ text: '' }),

                        // Add the table to the document
                        new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } })  // Ensure table spans full width
                    ]
                }]
            });

            // Convert to buffer and send response
            const buffer = await Packer.toBuffer(doc);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', 'attachment; filename=Repayment_Schedule_Report.docx');
            res.send(buffer);
        }


        // === INVALID FORMAT ===
        else {
            res.status(400).json({ error: 'Invalid format selected' });
        }

    } catch (error) {
        console.error('Error generating report:', error.message, error.stack);
        res.status(500).json({ error: error.message });
    }
};
