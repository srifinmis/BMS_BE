const express = require('express');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType } = require('docx');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

require('dotenv').config();
const models = initModels(sequelize);
const { lender_master, payment_details, repayment_schedule, tranche_details, sanction_details } = models;

exports.generateFundingMixReport = async (req, res) => {
    const { fromDate, toDate } = req.body;
    console.log("Daily backend: ", fromDate, toDate)

    try {
        const { start_date, end_date, lender_codes, sort_by } = req.query;

        const whereClause = {};
        if (start_date && end_date) {
            whereClause.tranche_date = {
                [Op.between]: [start_date, end_date]
            };
        }

        const paymentData = await payment_details.findAll({
            where: whereClause,
            include: [
                {
                    model: sanction_details,
                    as: 'sanction',
                    include: [
                        {
                            model: lender_master,
                            as: 'lender_code_lender_master',
                        }
                    ]
                },
                {
                    model: tranche_details,
                    as: 'tranche',
                }
            ],
            raw: true,
            nest: true
        });

        // Group payment amounts by sanction_id
        const paymentsBySanction = {};
        for (const d of paymentData) {
            const sid = d.sanction_id;
            if (!paymentsBySanction[sid]) {
                paymentsBySanction[sid] = 0;
            }
            paymentsBySanction[sid] += Number(d.payment_amount) || 0;
        }

        // Fetch all sanction details used in the paymentData
        const sanctionIds = Object.keys(paymentsBySanction);

        const sanctions = await sanction_details.findAll({
            where: {
                sanction_id: { [Op.in]: sanctionIds }
            },
            include: [{ model: lender_master, as: 'lender_code_lender_master' }],
            raw: true,
            nest: true
        });

        const sanctionMap = {};
        for (const s of sanctions) {
            sanctionMap[s.sanction_id] = s;
        }

        // Start ExcelJS logic
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Funding Mix');

        const ORG_NAME = `${process.env.LENDER_HEADER_LINE1}`;
        const ORG_CODE = `${process.env.ADDRESS_HEADER_TITLE}`;
        const today = new Date().toLocaleDateString('en-GB');
        const REPORT_TITLE = `Outstanding as on ${today}`;

        const orgNameRow = sheet.addRow([ORG_NAME]);
        sheet.mergeCells(`A${orgNameRow.number}:O${orgNameRow.number}`);
        orgNameRow.font = { bold: true, size: 14 };
        orgNameRow.alignment = { vertical: 'middle', horizontal: 'center' };

        const orgCodeRow = sheet.addRow([ORG_CODE]);
        sheet.mergeCells(`A${orgCodeRow.number}:O${orgCodeRow.number}`);
        orgCodeRow.font = { bold: true, size: 12 };
        orgCodeRow.alignment = { vertical: 'middle', horizontal: 'center' };

        const titleRow = sheet.addRow([REPORT_TITLE]);
        sheet.mergeCells(`A${titleRow.number}:O${titleRow.number}`);
        titleRow.font = { bold: true, size: 12 };
        titleRow.alignment = { vertical: 'middle', horizontal: 'center' };

        sheet.addRow([]);

        // Headers
        sheet.addRow([
            'Sr. No.',
            'Name of the Lender',
            'Lender Type',
            'Term Loan', '',
            'Commercial Paper', '',
            'Cash Credit', '',
            'NCD', '',
            'Securitisation / Asset Assignment', '',
            'Grand Total', ''
        ]);
        sheet.addRow([
            '', '', '',
            'Sanction Amount', 'Outstanding Amount (In ₹)',
            'Sanction Amount', 'Outstanding Amount (In ₹)',
            'Sanction Amount', 'Outstanding Amount (In ₹)',
            'Sanction Amount', 'Outstanding Amount (In ₹)',
            'Sanction', 'Outstanding Amount (In ₹)',
            'Sanction Amount', 'Outstanding Amount (In ₹)'
        ]);
        [5, 6].forEach(rowNum => {
            const row = sheet.getRow(rowNum);
            row.font = { bold: true };
            row.alignment = { vertical: 'middle', horizontal: 'center' };
            row.eachCell(cell => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });
        const headerMerges = [
            ['A5', 'A6'], ['B5', 'B6'], ['C5', 'C6'],
            ['D5', 'E5'], ['F5', 'G5'], ['H5', 'I5'],
            ['J5', 'K5'], ['L5', 'M5'], ['N5', 'O5']
        ];
        headerMerges.forEach(([start, end]) => sheet.mergeCells(`${start}:${end}`));

        let rowNumber = 7;
        let count = 1;

        const total = {
            termSanction: 0,
            termOutstanding: 0,
            ncdOutstanding: 0,
            securitisationOutstanding: 0,
            grandSanction: 0,
            grandOutstanding: 0
        };

        for (const sid of sanctionIds) {
            const s = sanctionMap[sid];
            const paymentAmt = paymentsBySanction[sid];
            const termSanction = Number(s.sanction_amount) || 0;
            const outstanding = termSanction - paymentAmt;

            const ncdSanction = Number(s.ncd_sanction_amount) || 0;
            const ncdOutstanding = Number(s.ncd_outstanding_amount) || 0;
            const securitisation = Number(s.securitisation_outstanding_amount) || 0;

            const grandSanction = termSanction + ncdSanction;
            const grandOutstanding = outstanding + ncdOutstanding + securitisation;

            total.termSanction += termSanction;
            total.termOutstanding += outstanding;
            total.ncdOutstanding += ncdOutstanding;
            total.securitisationOutstanding += securitisation;
            total.grandSanction += grandSanction;
            total.grandOutstanding += grandOutstanding;

            const row = sheet.addRow([
                count++,
                s.lender_code_lender_master.lender_name || '',
                s.lender_code_lender_master.lender_type || '',
                termSanction,
                outstanding,
                '', '',
                '', '',
                '', ncdOutstanding,
                '', securitisation,
                grandSanction,
                grandOutstanding
            ]);

            row.eachCell(cell => {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        }

        const totalRow = sheet.addRow([
            '', 'TOTAL', '',
            total.termSanction,
            total.termOutstanding,
            '', '',
            '', '',
            '', total.ncdOutstanding,
            '', total.securitisationOutstanding,
            total.grandSanction,
            total.grandOutstanding
        ]);
        totalRow.font = { bold: true };
        totalRow.eachCell(cell => {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        const buffer = await workbook.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=FundingMix_Report.xlsx');
        res.send(buffer);

    } catch (error) {
        console.error('Error generating funding mix Excel:', error);
        res.status(500).json({ error: 'Internal Server Error' });
    }
}