const express = require('express');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType, PageOrientation } = require('docx');
// const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType } = require('docx');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

require('dotenv').config();
const models = initModels(sequelize);
const { lender_master, tranche_details, repayment_schedule, sanction_details } = models;

exports.generateDatewiseRepaymentReport = async (req, res) => {
    const { fromDate, toDate, lenders, format, sortBy } = req.body;
    console.log("Datewise backend: ", fromDate, toDate, lenders, format, sortBy)

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

        const sortBy = (req.query.sortBy || '').toLowerCase().trim();

        const validSortFields = ['lender_code', 'due_date'];
        const orderClause = validSortFields.includes(sortBy) ? [[sortBy, 'ASC']] : undefined;
        const data = await repayment_schedule.findAll({
            where: whereClause,
            include: [
                {
                    model: sanction_details,
                    as: 'sanction',
                    attributes: ['sanction_amount'],
                    include: [
                        {
                            model: lender_master,
                            as: 'lender_code_lender_master',
                            attributes: ['lender_name']
                        }
                    ]
                },
                {
                    model: tranche_details,
                    as: 'tranche',
                    attributes: [
                        'tranche_number',
                        'tranche_amount',
                        'current_ac_no',
                        'bank_name',
                        'bank_branch',
                        'ifsc_code'
                    ]
                }
            ],
            ...(orderClause && { order: orderClause }),
            raw: true // remove this if you want nested structure
        });


        console.log("datewise fetching: ", data)

        if (!data || data.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }
        // console.log("roc data backend: ", data)

        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
        const REPORT_TITLE = process.env.REPORT_TITLE || 'Date Wise Repayment Statement';
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
            { header: 'Lender Code', key: 'lender_code', width: 20 },
            { header: 'Lender Name', key: 'sanction.lender_code_lender_master.lender_name', width: 20 },
            { header: 'Due Date', key: 'due_date', width: 20 },
            { header: 'Sanction Amount (In Rs)', key: 'sanction.sanction_amount', width: 25 },
            { header: 'Tranche Number', key: 'tranche.tranche_number', width: 20 },
            { header: 'Tranche Amount (In Rs)', key: 'tranche.tranche_amount' },
            { header: 'Principal Amount (In Rs)', key: 'principal_due', width: 25 },
            { header: 'Interest Amount (In Rs)', key: 'interest_due', width: 25 },
            { header: 'Total Amount (In Rs)', key: 'total_due', width: 25 },
            { header: 'Current Account Number', key: 'tranche.current_ac_no', width: 25 },
            { header: 'Bank Name', key: 'tranche.bank_name', width: 25 },
            { header: 'Bank Branch Name', key: 'tranche.bank_branch', width: 25 },
            { header: 'IFSC Code', key: 'tranche.ifsc_code', width: 25 }

        ];

        // === EXCEL FORMAT ===
        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('DateWise Repayment Statement Report');

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
            res.setHeader('Content-Disposition', 'attachment; filename=DateWise_Repayment_Statement_Report.xlsx');
            await workbook.xlsx.write(res);
            res.end();
        }
        // === PDF FORMAT ===
        else if (format === 'pdf') {
            const doc = new PDFDocument({ margin: 20, size: 'A4', layout: 'landscape' });
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename=DateWise_Repayment_Statement_Report.pdf');
            doc.pipe(res);

            const headerWidth = 500;
            const pageCenter = doc.page.width / 2;
            const headerX = pageCenter - headerWidth / 2;

            headerInfo.forEach(line => {
                doc.fontSize(12).text(line, headerX, doc.y, {
                    width: headerWidth,
                    align: 'center'
                });
                doc.moveDown(0.5);
            });

            doc.moveDown(2);

            // Layout constants
            const padding = 2;
            const rowHeight = 40;
            const startX = doc.page.margins.left;
            const availableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;

            // Step 1: Calculate natural column width proportion
            const charWidth = 6;

            // Estimate content length (max(header, values.length)) for proportional layout
            const contentLengths = columns.map(col => {
                const headerLen = col.header.length;
                const maxDataLen = Math.max(...data.map(row => {
                    const keys = col.key.split('.');
                    const value = keys.length === 2 ? row[`${keys[0]}.${keys[1]}`] || row[keys.join('.')] : row[col.key];
                    return `${value ?? ''}`.length;
                }));
                return Math.max(headerLen, maxDataLen);
            });

            const totalContentLength = contentLengths.reduce((a, b) => a + b, 0);

            // Step 2: Scale to availableWidth
            const columnWidths = contentLengths.map(len => (len / totalContentLength) * availableWidth);

            // Step 3: Draw Table Header
            let x = startX;
            let y = doc.y;

            columns.forEach((col, i) => {
                doc.rect(x, y, columnWidths[i], rowHeight).stroke();
                doc.font('Helvetica-Bold').fontSize(10).text(col.header, x + padding, y + 6, {
                    width: columnWidths[i] - padding * 2,
                    align: 'center'
                });
                x += columnWidths[i];
            });

            y += rowHeight;

            // Step 4: Draw Table Rows
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

                // Check for page break
                if (y + rowHeight > doc.page.height - doc.page.margins.bottom) {
                    doc.addPage({ layout: 'landscape', size: 'A4' });
                    y = doc.page.margins.top;

                    // Repeat Header on New Page
                    x = startX;
                    columns.forEach((col, i) => {
                        doc.rect(x, y, columnWidths[i], rowHeight).stroke();
                        doc.font('Helvetica-Bold').fontSize(10).text(col.header, x + padding, y + 6, {
                            width: columnWidths[i] - padding * 2,
                            align: 'center'
                        });
                        x += columnWidths[i];
                    });

                    y += rowHeight;
                }
            });

            doc.end();
        }

        // === WORD FORMAT ===

        else if (format === 'word') {
            const tableRows = [];

            // Add header row with styling
            const headerCells = columns.map(col =>
                new TableCell({
                    children: [new Paragraph({ text: col.header, bold: true, alignment: AlignmentType.CENTER })],
                    shading: { fill: 'D9D9D9' }, // light gray background for header
                    verticalAlign: 'center',
                    width: { size: 1000, type: WidthType.DXA }
                })
            );

            tableRows.push(new TableRow({ children: headerCells }));

            // Add data rows
            data.forEach(row => {
                const dataCells = columns.map(col => {
                    const keys = col.key.split('.');
                    const text = keys.length === 3 ? row[`${keys[0]}.${keys[1]}.${keys[2]}`]
                        : keys.length === 2 ? row[`${keys[0]}.${keys[1]}`]
                            : row[col.key];
                    return new TableCell({
                        children: [new Paragraph(String(text || ''))],
                        verticalAlign: 'center',
                        width: { size: 1000, type: WidthType.DXA }
                    });
                });
                tableRows.push(new TableRow({ children: dataCells }));
            });

            // Create the document
            const doc = new Document({
                sections: [{
                    properties: { page: { margin: { top: 700, right: 700, bottom: 700, left: 700 }, size: { orientation: PageOrientation.LANDSCAPE } } },
                    children: [
                        new Paragraph({
                            text: ORG_NAME,
                            heading: "Title",
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 200 }
                        }),
                        new Paragraph({
                            text: ORG_ADDRESS,
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 200 }
                        }),
                        new Paragraph({
                            text: REPORT_TITLE,
                            heading: "Heading1",
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 400 }
                        }),
                        new Table({
                            rows: tableRows,
                            width: {
                                size: 100,
                                type: WidthType.PERCENTAGE,
                            },
                            borders: {
                                top: { style: 'single', size: 1 },
                                bottom: { style: 'single', size: 1 },
                                left: { style: 'single', size: 1 },
                                right: { style: 'single', size: 1 },
                                insideHorizontal: { style: 'single' },
                                insideVertical: { style: 'single', size: 1 },
                            },
                        }),
                    ],
                }],
            });

            const buffer = await Packer.toBuffer(doc);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', 'attachment; filename=DateWise_Repayment_Statement_Report.docx');
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


// /////////////////////////////daily repayment statment report code /////////////


exports.generateDailyRepaymentReport = async (req, res) => {
    const { fromDate, toDate, date, lenders, format, sortBy } = req.body;
    console.log("Daily backend: ", fromDate, toDate, date, lenders, format, sortBy)

    try {
        const selectedDate = new Date(date);
        const startOfDay = new Date(selectedDate);
        startOfDay.setHours(0, 0, 0, 0);

        const endOfDay = new Date(selectedDate);
        endOfDay.setHours(23, 59, 59, 999);

        if (isNaN(selectedDate)) {
            return res.status(400).json({ error: 'Invalid date range provided' });
        }

        const whereClause = {
            createdat: { [Op.between]: [startOfDay, endOfDay] }
        };

        if (lenders !== 'all') {
            whereClause.lender_code = { [Op.in]: lenders };
        }

        const sortBy = "lender_code";

        const validSortFields = ['lender_code', 'due_date'];
        const orderClause = validSortFields.includes(sortBy) ? [[sortBy, 'ASC']] : undefined;
        const data = await repayment_schedule.findAll({
            where: whereClause,
            include: [
                {
                    model: sanction_details,
                    as: 'sanction',
                    attributes: ['sanction_amount', 'sanction_date'],
                    include: [
                        {
                            model: lender_master,
                            as: 'lender_code_lender_master',
                            attributes: ['lender_code']
                        }
                    ]
                },
                {
                    model: tranche_details,
                    as: 'tranche',
                    attributes: [
                        'tranche_date',
                        'tranche_amount',
                    ]
                }
            ],
            ...(orderClause && { order: orderClause }),
            raw: true // remove this if you want nested structure
        });


        console.log("daily fetching: ", data)

        if (!data || data.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }
        // console.log("roc data backend: ", data)

        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
        const REPORT_TITLE = process.env.REPORT_TITLE || 'Report: Daily Repayment Statement';
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
            { header: 'Lender Code', key: 'lender_code', width: 20 },
            { header: 'Sanction Date', key: 'sanction.sanction_date', width: 20 },
            { header: 'Sanction Amount (In Rs)', key: 'sanction.sanction_amount', width: 25 },
            { header: 'Tranche Drawdown Date', key: 'tranche.tranche_date', width: 20 },
            { header: 'Tranche Amount (In Rs)', key: 'tranche.tranche_amount' },
            { header: 'Due Date', key: 'due_date', width: 20 },
            { header: 'Principal Amount (In Rs)', key: 'principal_due', width: 25 },
            { header: 'Interest Amount (In Rs)', key: 'interest_due', width: 25 },
            { header: 'Total Amount (In Rs)', key: 'total_due', width: 25 },
        ];

        // === EXCEL FORMAT ===
        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Daily Repayment Statement Report');

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
            res.setHeader('Content-Disposition', 'attachment; filename=Daily_Repayment_Statement_Report.xlsx');
            await workbook.xlsx.write(res);
            res.end();
        }
        // === PDF FORMAT ===
        else if (format === 'pdf') {
            const PDFDocument = require('pdfkit');
            const doc = new PDFDocument({
                margin: 20,
                size: 'A4',
                layout: 'landscape' // 👈 Enable landscape orientation
            });

            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename=Daily_Repayment_Statement_Report.pdf');
            doc.pipe(res);

            const headerWidth = 500; // Adjusted for landscape width
            const pageCenter = doc.page.width / 2;
            const headerX = pageCenter - headerWidth / 2;

            // Render header
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

            // Step 2: Scale if total width exceeds page
            const totalNaturalWidth = naturalWidths.reduce((sum, w) => sum + w, 0);
            let columnWidths = [...naturalWidths];

            if (totalNaturalWidth > availableWidth) {
                const scale = availableWidth / totalNaturalWidth;
                columnWidths = naturalWidths.map(w => w * scale);
            }

            const startX = doc.page.margins.left;
            let y = doc.y;

            // Step 3: Draw Header Row
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
                    y = doc.page.margins.top; // Reset y for new page
                }
            });

            doc.end();
        }
        // === WORD FORMAT ===
        else if (format === 'word') {
            const { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType, PageOrientation } = require("docx");

            // Helper function to get value from nested keys
            function getValue(obj, path) {
                if (!path) return '';
                if (obj[path] !== undefined) return obj[path];
                return path.split('.').reduce((acc, part) => acc && acc[part], obj) ?? '';
            }

            // Table header row
            const tableRows = [
                new TableRow({
                    children: columns.map(col =>
                        new TableCell({
                            children: [new Paragraph({
                                text: col.header,
                                bold: true
                            })],
                            width: { size: 100 / columns.length, type: WidthType.PERCENTAGE }
                        })
                    )
                })
            ];

            // Add data rows
            data.forEach((row) => {
                const cells = columns.map(col => {
                    const value = getValue(row, col.key);
                    return new TableCell({
                        children: [new Paragraph({
                            text: String(value ?? '')
                        })],
                        width: { size: 100 / columns.length, type: WidthType.PERCENTAGE }
                    });
                });

                tableRows.push(new TableRow({ children: cells }));
            });

            // Create the document in landscape orientation
            const doc = new Document({
                sections: [{
                    properties: {
                        page: {
                            size: {
                                orientation: PageOrientation.LANDSCAPE, // Landscape mode
                            }
                        }
                    },
                    children: [
                        ...headerInfo.map(line => new Paragraph({ text: line, alignment: AlignmentType.CENTER })),
                        new Paragraph({ text: '' }),
                        new Table({
                            rows: tableRows,
                            width: { size: 100, type: WidthType.PERCENTAGE }
                        })
                    ]
                }]
            });

            // Convert to buffer and send
            const buffer = await Packer.toBuffer(doc);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', 'attachment; filename=Daily_Repayment_Statement_Report.docx');
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