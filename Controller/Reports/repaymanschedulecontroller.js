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
    const { banks = [], sanctions = [], tranches = [], format } = req.body;
    console.log("Repayment Schedule backend: ", banks, sanctions, tranches, format);

    try {
        const whereClause = {};
        if (banks !== 'all' && banks.length > 0) {
            whereClause['$tranche.bank_name$'] = { [Op.in]: banks };
        }
        if (sanctions !== 'all' && sanctions.length > 0) {
            whereClause['sanction_id'] = { [Op.in]: sanctions };
        }
        if (tranches !== 'all' && tranches.length > 0) {
            whereClause['tranche_id'] = { [Op.in]: tranches };
        }

        const data = await repayment_schedule.findAll({
            where: whereClause,
            include: [
                {
                    model: tranche_details,
                    as: 'tranche',
                    attributes: ['tranche_id', "tranche_amount"],
                    include: [
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
                }
            ]
        });

        if (!data || data.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }

        const trancheIds = [...new Set(data.map(d => d.tranche_id))];
        const dueDates = [...new Set(
            data.map(d => d.due_date ? new Date(d.due_date).toISOString().split('T')[0] : null).filter(Boolean)
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

        const paymentMap = {};
        payments.forEach(p => {
            const dateKey = new Date(p.payment_date).toISOString().split('T')[0];
            paymentMap[`${p.tranche_id}_${dateKey}`] = parseFloat(p.total_payment || 0);
        });

        const result = data.map(row => {
            const dueDate = row.due_date ? new Date(row.due_date).toISOString().split('T')[0] : '';
            const paymentMade = paymentMap[`${row.tranche_id}_${dueDate}`] || 0;
            const loanOutstanding = parseFloat(row.tranche?.tranche_amount || 0) - paymentMade;
            return {
                ...row.toJSON(),
                loan_outstanding: loanOutstanding
            };
        });

        console.log("Repayment Schedule fetching: ", data)

        if (!result || result.length === 0) {
            return res.status(404).json({ message: 'No records found for the selected filters.' });
        }

        const ORG_NAME = process.env.LENDER_HEADER_LINE1 || 'SRIFIN CREDIT PRIVATE LIMITED';
        const ORG_ADDRESS = process.env.ORG_ADDRESS || 'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),Patrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
        const REPORT_TITLE = process.env.REPORT_TITLE || 'Repayment Schedule';

        const headerInfo = [ORG_NAME, '', ORG_ADDRESS, '', REPORT_TITLE, ''];
        const today = new Date().toLocaleDateString('en-GB');

        const columns = [
            { header: 'Due Date', key: 'due_date', width: 20 },
            { header: 'Loan Outstanding (In Rs)', key: 'loan_outstanding', width: 25 },
            { header: 'Principal Due (In Rs)', key: 'principal_due', width: 25 },
            { header: 'No. of Days', key: 'interest_days', width: 20 },
            { header: 'Rate Of Interest', key: 'interest_rate', width: 20 },
            { header: 'Interest Amount (In Rs)', key: 'interest_due', width: 25 },
            { header: 'Total Due (In Rs)', key: 'total_due', width: 25 },
        ];

        if (format === 'excel') {
            const ExcelJS = require('exceljs');
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Repayment Schedule');

            const totalCols = columns.length;
            const mergeRange = `A1:${String.fromCharCode(64 + totalCols)}1`;
            sheet.mergeCells(mergeRange);

            headerInfo.forEach((line, index) => {
                const row = sheet.addRow([line]);
                sheet.mergeCells(`A${row.number}:${String.fromCharCode(64 + totalCols)}${row.number}`);
                row.font = { bold: true, size: index === 0 ? 14 : 12 };
                row.alignment = { vertical: 'middle', horizontal: 'center' };
                if (line === '') sheet.addRow([]);
            });

            const headerRow = sheet.addRow(columns.map(col => col.header));
            headerRow.font = { bold: true };
            headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

            columns.forEach((col, i) => sheet.getColumn(i + 1).width = col.width);

            result.forEach(row => {
                const rowValues = columns.map(col => row[col.key]);
                sheet.addRow(rowValues);
            });

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', 'attachment; filename=RepaymentSchedule_Report.xlsx');
            await workbook.xlsx.write(res);
            res.end();

        } else if (format === 'pdf') {
            const PDFDocument = require('pdfkit');
            const doc = new PDFDocument({ margin: 40, size: 'A4', layout: 'landscape' });

            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', 'attachment; filename="RepaymentSchedule_Report.pdf"');
            doc.pipe(res);

            const pageWidth = doc.page.width;
            const headerWidth = 500;
            const headerX = (pageWidth - headerWidth) / 2;

            headerInfo.forEach(line => {
                doc.fontSize(12).text(line, headerX, doc.y, { width: headerWidth, align: 'center' });
                doc.moveDown(0.5);
            });

            doc.moveDown(1.5); 

            const tableTop = doc.y;
            const rowHeight = 30;
            const colWidths = columns.map(() => 100); 
            const tableLeft = doc.page.margins.left;

            let x = tableLeft;
            columns.forEach((col, i) => {
                doc.rect(x, tableTop, colWidths[i], rowHeight).fillAndStroke('#f0f0f0', 'black');
                doc.fillColor('black')
                    .fontSize(10)
                    .text(col.header, x + 7, tableTop + 7, { width: colWidths[i] - 10, align: 'left' });
                    x += colWidths[i];
            });

            let y = tableTop + rowHeight;
            result.forEach((row, rowIndex) => {
                x = tableLeft;

                if (y + rowHeight > doc.page.height - doc.page.margins.bottom) {
                    doc.addPage();
                    y = doc.page.margins.top;
                    x = tableLeft;
                    columns.forEach((col, i) => {
                        doc.rect(x, y, colWidths[i], rowHeight).fillAndStroke('#f0f0f0', 'black');
                        doc.fillColor('black')
                            .fontSize(10)
                            .text(col.header, x + 5, y + 7, { width: colWidths[i] - 10, align: 'left' });
                        x += colWidths[i];
                    });
                        y += rowHeight;
                }

                x = tableLeft;
                columns.forEach((col, i) => {
                    const cellText = row[col.key] != null ? String(row[col.key]) : '';
                    doc.rect(x, y, colWidths[i], rowHeight).stroke();
                    doc.fillColor('black')
                        .fontSize(10)
                        .text(cellText, x + 5, y + 7, { width: colWidths[i] - 10, align: 'left' });
                    x += colWidths[i];
                });
                    y += rowHeight;
            });
            doc.end();
        }

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
                            width: { size: 100 / columns.length, type: WidthType.PERCENTAGE }  // Width as percentage of total document width
                        })
                    )
                })
            ];

            result.forEach((row, index) => {
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
                    children: [
                        ...headerInfo.map(line => new Paragraph({ text: line, alignment: AlignmentType.CENTER })),
                        new Paragraph({ text: '' }),
                        new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } })  
                    ]
                }]
            });

            const buffer = await Packer.toBuffer(doc);
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', 'attachment; filename=RepaymentSchedule_Report.docx');
            res.send(buffer);

        } else {
            res.status(400).json({ message: 'Invalid format selected.' });
        }

    } catch (error) {
        console.error('Error generating repayment schedule report:', error);
        res.status(500).send('Server Error');
    }
};