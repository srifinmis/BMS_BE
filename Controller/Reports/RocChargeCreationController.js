const express = require('express');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType } = require('docx');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

require('dotenv').config();
const models = initModels(sequelize);
const { roc_forms, sanction_details } = models;

exports.generateRocChargeCreationReport = async (req, res) => {
    const { fromDate, toDate, lenders, format, sortBy } = req.body;

    try {
        const start = new Date(fromDate);
        const end = new Date(toDate);
        end.setHours(23, 59, 59, 999);

        if (isNaN(start) || isNaN(end)) {
            return res.status(400).json({ error: 'Invalid date range provided' });
        }

        const whereClause = {
            createdat: { [Op.between]: [start, end] }
        };

        if (lenders !== 'all') {
            whereClause.lender_code = { [Op.in]: lenders };
        }

        const validSortFields = ['lender_code', 'sanction_id', 'createdat'];
        const sortColumn = validSortFields.includes(sortBy) ? sortBy : 'createdat';

        const data = await roc_forms.findAll({
            where: whereClause,
            include: [
                {
                    model: sanction_details,
                    as: 'sanction',
                    attributes: ['sanction_amount', 'sanction_date']
                }
            ],
            order: [[sortColumn, 'ASC']],
            raw: true
        });
        if (!data || data.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }
        // console.log("roc data backend: ", data)

        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
        const REPORT_TITLE = process.env.REPORT_TITLE || 'ROC Charge Creation';
        const today = new Date().toLocaleDateString('en-GB');

        const headerInfo = [
            ORG_NAME,
            '',
            ORG_ADDRESS,
            '',
            `Effective Interest Date As on: ${today}`,
            ''
        ];

        const columns = [
            { header: 'Lender Code', key: 'lender_code', width: 20 },
            { header: 'Sanction No.', key: 'sanction_id', width: 20 },
            { header: 'Sanction Amount (In RS)', key: 'sanction.sanction_amount', width: 25 },
            { header: 'Sanction Date', key: 'sanction.sanction_date', width: 20 },
            { header: 'Document Execution Date', key: 'document_executed_date' },
            { header: 'Due Date of Filing', key: 'due_date_charge_creation', width: 25 },
            { header: 'Actual Date of Filing', key: 'due_date_satisfaction', width: 25 }
        ];

        // === EXCEL FORMAT ===
        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('ROC Charge Creation');

            // === Add Organization Name ===
            const orgNameRow = sheet.addRow([ORG_NAME]);
            sheet.mergeCells(`A${orgNameRow.number}:G${orgNameRow.number}`);
            orgNameRow.font = { bold: true, size: 14 };
            orgNameRow.alignment = { vertical: 'middle', horizontal: 'center' };

            // === Spacer Row === (After Organization Name)
            sheet.addRow([]);

            // === Add Address ===
            const addressRow = sheet.addRow([ORG_ADDRESS]);
            sheet.mergeCells(`A${addressRow.number}:G${addressRow.number}`);
            addressRow.font = { bold: true, size: 12 };
            addressRow.alignment = { vertical: 'middle', horizontal: 'center' };

            // === Spacer Row === (After Address)
            sheet.addRow([]);

            // === Add Report Date ===
            const dateRow = sheet.addRow([`Effective Interest Date As on: ${today}`]);
            sheet.mergeCells(`A${dateRow.number}:G${dateRow.number}`);
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
            data.forEach(row => {
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
            res.setHeader('Content-Disposition', 'attachment; filename=ROC_Charge_Creation_Report.xlsx');
            await workbook.xlsx.write(res);
            res.end();
        }
        // === PDF FORMAT ===
        else if (format === 'pdf') {
            const doc = new PDFDocument({ margin: 20, size: 'A4' });
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename=ROC_Charge_Creation_Report.pdf');
            doc.pipe(res);

            const headerWidth = 350;
            const pageCenter = doc.page.width / 2;
            const headerX = pageCenter - headerWidth / 2;

            headerInfo.forEach(line => {
                doc.fontSize(10).text(line, headerX, doc.y, {
                    width: headerWidth,
                    align: 'center'
                });
                doc.moveDown(0.5); // Optional: adjust vertical gap
            });
            doc.moveDown().fontSize(10).text(`Report: ${REPORT_TITLE}`, headerX, doc.y, {
                width: headerWidth,
                align: 'center'
            });
            doc.moveDown(2);

            const pageWidth = doc.page.width;
            const pageMargins = doc.page.margins.left + doc.page.margins.right;
            const availableWidth = pageWidth - pageMargins;
            const padding = 4;
            const rowHeight = 30;
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

            // Step 2: Scale if total width exceeds A4
            const totalNaturalWidth = naturalWidths.reduce((sum, w) => sum + w, 0);
            let columnWidths = [...naturalWidths];

            if (totalNaturalWidth > availableWidth) {
                const scale = availableWidth / totalNaturalWidth;
                columnWidths = naturalWidths.map(w => w * scale);
            }

            const startX = doc.page.margins.left;
            let y = doc.y;

            // Step 3: Draw Header
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

            // Step 4: Draw Data Rows
            data.forEach(row => {
                x = startX;
                columns.forEach((col, i) => {
                    const keys = col.key.split('.');
                    const value = keys.length === 2 ? row[`${keys[0]}.${keys[1]}`] || row[keys.join('.')] : row[col.key];
                    const cellText = `${value ?? ''}`;
                    doc.rect(x, y, columnWidths[i], rowHeight).stroke();
                    doc.font('Helvetica').fontSize(10).text(cellText, x + padding, y + 6, {
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
            // console.log("Generating Word report...");
            // console.log("Sample data row:", data[0]); // Check the structure of the first data row

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
            data.forEach((row, index) => {
                // Debugging: Check the row object and the columns keys
                // console.log(`Processing Row ${index + 1}:`, row);

                const cells = columns.map(col => {
                    const value = getValue(row, col.key);

                    // Debugging: Output the value for each cell
                    // console.log(`Row ${index + 1}, Column: ${col.header}, Value: ${value}`);

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
            res.setHeader('Content-Disposition', 'attachment; filename=ROC_Charge_Creation_Report.docx');
            res.send(buffer);
        }


        // === INVALID FORMAT ===
        else {
            res.status(400).json({ error: 'Invalid format selected' });
        }

    } catch (error) {
        console.error('Error generating report:', error);
        res.status(500).send('Server Error');
    }
};

exports.getroclenders = async (req, res) => {
    try {
        const roc = await roc_forms.findAll({
            attributes: [
                "lender_code", "sanction_id"
            ]
        });

        return res.status(201).json({ success: true, data: roc });
    } catch (error) {
        console.error("Error fetching ROC:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
}



exports.generateRocSatisfactionChargeReport = async (req, res) => {
    const { fromDate, toDate, lenders, format, sortBy } = req.body;

    try {
        const start = new Date(fromDate);
        const end = new Date(toDate);
        end.setHours(23, 59, 59, 999);

        if (isNaN(start) || isNaN(end)) {
            return res.status(400).json({ error: 'Invalid date range provided' });
        }

        const whereClause = {
            createdat: { [Op.between]: [start, end] }
        };

        if (lenders !== 'all') {
            whereClause.lender_code = { [Op.in]: lenders };
        }

        const validSortFields = ['lender_code', 'sanction_id', 'createdat'];
        const sortColumn = validSortFields.includes(sortBy) ? sortBy : 'createdat';

        const data = await roc_forms.findAll({
            where: whereClause,
            include: [
                {
                    model: sanction_details,
                    as: 'sanction',
                    attributes: ['sanction_amount', 'sanction_date']
                }
            ],
            order: [[sortColumn, 'ASC']],
            raw: true
        });
        if (!data || data.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }
        // console.log("roc data backend: ", data)

        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
        const REPORT_TITLE = process.env.REPORT_TITLE || 'ROC Satisfaction of Charge';
        const today = new Date().toLocaleDateString('en-GB');

        const headerInfo = [
            'SRIFIN CREDIT PRIVATE LIMITED',
            '',
            ORG_ADDRESS,
            '',
            `Effective Interest Date As on: ${today}`,
            ''
        ];

        const columns = [
            { header: 'Lender Code', key: 'lender_code', width: 20 },
            { header: 'Sanction No.', key: 'sanction_id', width: 20 },
            { header: 'Sanction Amount (In RS)', key: 'sanction.sanction_amount', width: 25 },
            { header: 'Sanction Date', key: 'sanction.sanction_date', width: 20 },
            { header: 'Loan Closure Date', key: 'loan_date' },
            { header: 'Due Date of Filing', key: 'due_date_charge_creation', width: 25 },
            { header: 'Actual Date of Filing', key: 'due_date', width: 25 }
        ];

        // === EXCEL FORMAT ===
        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('ROC Satisfaction of Charge');

            // === Add Organization Name ===
            const orgNameRow = sheet.addRow([ORG_NAME]);
            sheet.mergeCells(`A${orgNameRow.number}:G${orgNameRow.number}`);
            orgNameRow.font = { bold: true, size: 14 };
            orgNameRow.alignment = { vertical: 'middle', horizontal: 'center' };

            // === Spacer Row === (After Organization Name)
            sheet.addRow([]);

            // === Add Address ===
            const addressRow = sheet.addRow([ORG_ADDRESS]);
            sheet.mergeCells(`A${addressRow.number}:G${addressRow.number}`);
            addressRow.font = { bold: true, size: 12 };
            addressRow.alignment = { vertical: 'middle', horizontal: 'center' };

            // === Spacer Row === (After Address) 
            sheet.addRow([]);

            // === Add Report Date ===
            const dateRow = sheet.addRow([`Effective Interest Date As on: ${today}`]);
            sheet.mergeCells(`A${dateRow.number}:G${dateRow.number}`);
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
            data.forEach(row => {
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
            res.setHeader('Content-Disposition', 'attachment; filename=Roc_Satisfaction_of_Charge_Report.xlsx');
            await workbook.xlsx.write(res);
            res.end();
        }
        // === PDF FORMAT ===
        else if (format === 'pdf') {
            const doc = new PDFDocument({ margin: 20, size: 'A4' });
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename=Roc_Satisfaction_of_Charge_Report.pdf');
            doc.pipe(res);

            const headerWidth = 350;
            const pageCenter = doc.page.width / 2;
            const headerX = pageCenter - headerWidth / 2;

            headerInfo.forEach(line => {
                doc.fontSize(10).text(line, headerX, doc.y, {
                    bold: true,
                    width: headerWidth,
                    align: 'center'
                });
                doc.moveDown(0.5); // Optional: adjust vertical gap
            });
            doc.moveDown().fontSize(10).text(`Report: ${REPORT_TITLE}`, headerX, doc.y, {
                bold: true,
                width: headerWidth,
                align: 'center'
            });
            doc.moveDown(2);

            const pageWidth = doc.page.width;
            const pageMargins = doc.page.margins.left + doc.page.margins.right;
            const availableWidth = pageWidth - pageMargins;
            const padding = 4;
            const rowHeight = 30;
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

            // Step 2: Scale if total width exceeds A4
            const totalNaturalWidth = naturalWidths.reduce((sum, w) => sum + w, 0);
            let columnWidths = [...naturalWidths];

            if (totalNaturalWidth > availableWidth) {
                const scale = availableWidth / totalNaturalWidth;
                columnWidths = naturalWidths.map(w => w * scale);
            }

            const startX = doc.page.margins.left;
            let y = doc.y;

            // Step 3: Draw Header
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

            // Step 4: Draw Data Rows
            data.forEach(row => {
                x = startX;
                columns.forEach((col, i) => {
                    const keys = col.key.split('.');
                    const value = keys.length === 2 ? row[`${keys[0]}.${keys[1]}`] || row[keys.join('.')] : row[col.key];
                    const cellText = `${value ?? ''}`;
                    doc.rect(x, y, columnWidths[i], rowHeight).stroke();
                    doc.font('Helvetica').fontSize(10).text(cellText, x + padding, y + 6, {
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
            // console.log("Generating Word report...");
            // console.log("Sample data row:", data[0]); // Check the structure of the first data row

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
            data.forEach((row, index) => {
                // Debugging: Check the row object and the columns keys
                // console.log(`Processing Row ${index + 1}:`, row);

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
            res.setHeader('Content-Disposition', 'attachment; filename=ROC_Satisfaction_of_Charge_Report.docx');
            res.send(buffer);
        }

        // === INVALID FORMAT ===
        else {
            res.status(400).json({ error: 'Invalid format selected' });
        }

    } catch (error) {
        console.error('Error generating report:', error);
        res.status(500).send('Server Error');
    }
};