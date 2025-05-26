const express = require('express');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, PageOrientation } = require('docx');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

require('dotenv').config();
const models = initModels(sequelize);
const { lender_master, tranche_details, sanction_details } = models;

exports.generateLoanTrancheDetailsReport = async (req, res) => {
    const { fromDate, toDate, lenders, format, sortBy } = req.body;
    console.log("Loan Tranche backend: ", fromDate, toDate, lenders, format, sortBy)

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
        const data = await tranche_details.findAll({
            where: whereClause,
            include: [
                {
                    model: sanction_details,
                    as: 'sanction',
                    attributes: ['sanction_id', "loan_type"],
                    include: [
                        {
                            model: lender_master,
                            as: 'lender_code_lender_master',
                            attributes: ['lender_code', 'lender_name']
                        }
                    ]
                }
            ],
            ...(orderClause && { order: orderClause }),
            raw: true // remove this if you want nested structure
        });

        console.log("Loan Tranche fetching: ", data)

        if (!data || data.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }
        // console.log("roc data backend: ", data)

        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
        const REPORT_TITLE = process.env.REPORT_TITLE || 'Tranche Details';
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
            { header: 'Tranche ID', key: 'tranche_id', width: 20 },
            { header: 'Drawdown Date', key: 'tranche_date', width: 20 },
            { header: 'Lender Code', key: 'lender_code', width: 20 },
            { header: 'Lender Name', key: 'sanction.lender_code_lender_master.lender_name', width: 20 },
            { header: 'Sanction ID', key: 'sanction.sanction_id', width: 25 },
            { header: 'Tranche Number', key: 'tranche_number', width: 20 },
            { header: 'Loan Type', key: 'sanction.loan_type', width: 20 },
            { header: 'Tranche Amount (In Rs)', key: 'tranche_amount' },
            { header: 'Interest Rate', key: 'interest_rate', width: 25 },
            { header: 'Principal Payment Frequency', key: 'principal_payment_frequency', width: 25 },
            { header: 'Principal Start Date', key: 'principal_start_date', width: 25 },
            { header: 'Interest Payment Frequency', key: 'interest_payment_frequency', width: 25 },
            { header: 'Interest Start Date', key: 'interest_start_date', width: 25 },
            { header: 'Tenure', key: 'tenure_months', width: 25 },
            { header: 'Interest Type', key: 'interest_type', width: 25 },
        ];

        // === EXCEL FORMAT ===
        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Loan Tranche Details');

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
            res.setHeader('Content-Disposition', 'attachment; filename=Loan_Tranche_Details_Report.xlsx');
            await workbook.xlsx.write(res);
            res.end();
        }
        // === PDF FORMAT ===
        else if (format === 'pdf') {
            const PDFDocument = require('pdfkit');
            const doc = new PDFDocument({ margin: 20, size: 'A4', layout: 'landscape' });

            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename=Loan_Tranche_Details_Report.pdf');
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
            data.forEach(row => {
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
            function getValue(obj, path) {
                if (!path) return '';
                if (obj[path] !== undefined) return obj[path];
                return path.split('.').reduce((acc, part) => acc && acc[part], obj) ?? '';
            }

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

            const doc = new Document({
                sections: [{
                    properties: {
                        page: {
                            size: {
                                orientation: PageOrientation.LANDSCAPE
                            },
                            margin: {
                                top: 720, bottom: 720, left: 720, right: 720
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

            const buffer = await Packer.toBuffer(doc);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', 'attachment; filename=Loan_Tranche_Details_Report.docx');
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
