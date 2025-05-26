
const express = require('express');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, PageOrientation } = require('docx');
// const { Document, Packer, Paragraph, Table, TableRow, TableCell, AlignmentType, WidthType } = require('docx');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

require('dotenv').config();
const models = initModels(sequelize);
const { sanction_details, lender_master } = models;


exports.generateSanctionDetailsReport = async (req, res) => {
    const { fromDate, toDate, lenders, format, sortBy } = req.body;

    try {
        const start = new Date(fromDate);
        const end = new Date(toDate);
        end.setHours(23, 59, 59, 999);

        if (isNaN(start) || isNaN(end)) {
            return res.status(400).json({ error: 'Invalid date range provided' });
        }

        const whereClause = {
            sanction_date: { [Op.between]: [start, end] }
        };

        if (lenders !== 'all') {
            whereClause.lender_code = { [Op.in]: lenders };
        }

        const validSortFields = ['lender_code', 'sanction_id', 'sanction_date'];
        const sortColumn = validSortFields.includes(sortBy) ? sortBy : 'sanction_date';

        const data = await sanction_details.findAll({
            where: whereClause,
            order: [[sortColumn, 'ASC']],
            include: [{
                model: lender_master,
                as: 'lender_code_lender_master',
                attributes: ['lender_name', 'status']
            }],
            raw: true
        });
        if (!data || data.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }

        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
        const REPORT_TITLE = process.env.REPORT_TITLE || 'Sanction Details';

        const columns = [
            { header: 'Sanction ID', key: 'sanction_id' },
            { header: 'Sanction Date', key: 'sanction_date' },
            { header: 'Lender Code', key: 'lender_code' },
            { header: 'Lender Name', key: 'lender_code_lender_master.lender_name' },
            { header: 'Facility Type', key: 'loan_type' },
            { header: 'Purpose Of Loan', key: 'purpose_of_loan' },
            { header: 'Interest Type', key: 'interest_type' },
            { header: 'Interest Rate (%)', key: 'interest_rate_fixed' },
            { header: 'Loan Tenure', key: 'tenure_months' },
            { header: 'Sanctioned Amount (in ₹)', key: 'sanction_amount' },
            { header: 'Processing Fee', key: 'processing_fee' },
            { header: 'Management Fee', key: 'syndication_fee' },
            { header: 'Other Expenses', key: 'other_expenses' },
            { header: 'Loan Status', key: 'lender_code_lender_master.status' },
            { header: 'Closure Date', key: 'sanction_validity' }
        ];

        // === Excel Format ===
        if (format === 'excel') {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Sanction Details');

            const orgNameRow = sheet.addRow([ORG_NAME]);
            sheet.mergeCells(`A${orgNameRow.number}:O${orgNameRow.number}`);
            orgNameRow.font = { bold: true, size: 14 };
            orgNameRow.alignment = { horizontal: 'center' };

            sheet.addRow([]);
            const addressRow = sheet.addRow([ORG_ADDRESS]);
            sheet.mergeCells(`A${addressRow.number}:O${addressRow.number}`);
            addressRow.font = { size: 12 };
            addressRow.alignment = { horizontal: 'center' };
            sheet.addRow([]);

            const headerRow = sheet.addRow(columns.map(col => col.header));
            headerRow.font = { bold: true };
            headerRow.alignment = { horizontal: 'center' };
            columns.forEach((col, idx) => {
                headerRow.getCell(idx + 1).border = {
                    top: { style: 'thin' }, left: { style: 'thin' },
                    bottom: { style: 'thin' }, right: { style: 'thin' }
                };
            });

            data.forEach(row => {
                const rowData = columns.map(col => row[col.key] || '');
                const dataRow = sheet.addRow(rowData);
                dataRow.eachCell(cell => {
                    cell.border = {
                        top: { style: 'thin' }, left: { style: 'thin' },
                        bottom: { style: 'thin' }, right: { style: 'thin' }
                    };
                });
            });

            columns.forEach((col, i) => sheet.getColumn(i + 1).width = 25);

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', 'attachment; filename=Sanction_Details_Report.xlsx');
            await workbook.xlsx.write(res);
            return res.end();
        }

        // === PDF Format ===
        else if (format === 'pdf') {
            const doc = new PDFDocument({ margin: 20, size: 'A3', layout: 'landscape' });
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename=Sanction_Details_Report.pdf');
            doc.pipe(res);

            doc.fontSize(12).text(ORG_NAME, { align: 'center' });
            doc.moveDown();
            doc.fontSize(10).text(ORG_ADDRESS, { align: 'center' });
            doc.moveDown().fontSize(10).text(`Report: ${REPORT_TITLE}`, { align: 'center' });
            doc.moveDown(2);

            const columnWidth = (doc.page.width - doc.page.margins.left - doc.page.margins.right) / columns.length;
            let y = doc.y;

            columns.forEach((col, i) => {
                doc.font('Helvetica-Bold').fontSize(9).text(col.header, doc.page.margins.left + i * columnWidth, y, {
                    width: columnWidth,
                    align: 'center'
                });
            });

            y += 20;

            data.forEach(row => {
                columns.forEach((col, i) => {
                    const value = row[col.key] ?? '';
                    doc.font('Helvetica').fontSize(9).text(`${value}`, doc.page.margins.left + i * columnWidth, y, {
                        width: columnWidth,
                        align: 'center'
                    });
                });
                y += 20;

                if (y + 20 > doc.page.height - doc.page.margins.bottom) {
                    doc.addPage();
                    y = doc.y;
                }
            });

            doc.end();
        }

        // === Word Format ===
        else if (format === 'word') {
            const headerParagraphs = [
                new Paragraph({
                    text: ORG_NAME,
                    alignment: AlignmentType.CENTER,
                    heading: 'Heading1',
                }),
                new Paragraph({
                    text: ORG_ADDRESS,
                    alignment: AlignmentType.CENTER,
                }),
                new Paragraph({
                    text: `Report: ${REPORT_TITLE}`,
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 },
                }),
            ];

            const tableRows = [];

            // Header row
            const headerCells = columns.map(col =>
                new TableCell({
                    children: [
                        new Paragraph({
                            text: col.header,
                            alignment: AlignmentType.CENTER,
                            bold: true,
                        }),
                    ],
                    width: { size: 100 / columns.length, type: WidthType.PERCENTAGE },
                    margins: { top: 100, bottom: 100, left: 100, right: 100 },
                })
            );
            tableRows.push(new TableRow({ children: headerCells }));

            // Data rows
            data.forEach(row => {
                const cells = columns.map(col =>
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: `${row[col.key] ?? ''}`,
                                alignment: AlignmentType.LEFT,
                            }),
                        ],
                        width: { size: 100 / columns.length, type: WidthType.PERCENTAGE },
                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                    })
                );
                tableRows.push(new TableRow({ children: cells }));
            });

            const table = new Table({
                rows: tableRows,
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE,
                },
                alignment: AlignmentType.CENTER,
            });

            // Generate document in landscape mode
            const doc = new Document({
                sections: [{
                    properties: {
                        page: {
                            size: {
                                orientation: PageOrientation.LANDSCAPE,
                            },
                        },
                    },
                    children: [
                        ...headerParagraphs,
                        table,
                    ],
                }],
            });

            const buffer = await Packer.toBuffer(doc);

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', 'attachment; filename=Sanction_Details_Report.docx');
            return res.send(buffer);
        }
        else {
            return res.status(400).json({ error: 'Invalid format selected' });
        }

    } catch (err) {
        console.error('Error generating Sanction Details Report:', err);
        return res.status(500).send('Server Error');
    }
};