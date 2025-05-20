require('dotenv').config();
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

const models = initModels(sequelize);
const { lender_master } = models;

// Environment variables for header
const headerLine1 = process.env.LENDER_HEADER_LINE1;
const headerLine2 = process.env.LENDER_HEADER_LINE2;
const headerLine3 = process.env.LENDER_HEADER_LINE3;
const headerTitle = process.env.LENDER_HEADER_TITLE;

exports.generateLenderMasterReport = async (req, res) => {
  const { fromDate, toDate, lenders, format } = req.body;

  try {
    const start = new Date(fromDate);
    const end = new Date(toDate);

    // Validate dates
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      return res.status(400).json({ error: 'Invalid date range provided' });
    }

    const whereClause = {
      createdat: {
        [Op.between]: [start, end],
      },
    };

    if (lenders && lenders.length > 0 && !lenders.includes('all')) {
      whereClause.lender_code = {
        [Op.in]: lenders,
      };
    }

    const data = await lender_master.findAll({
      where: whereClause,
      raw: true,
    });

    // Excel Report
    if (format === 'excel') {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Lender Master');

      // Add Header lines (Static)
      sheet.mergeCells('A1:E1');
      sheet.getCell('A1').value = headerLine1;
      sheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
      sheet.getRow(1).height = 30;

      sheet.mergeCells('A2:E2');
      sheet.getCell('A2').value = headerLine2;
      sheet.getCell('A2').alignment = { vertical: 'middle', horizontal: 'center' };
      sheet.getRow(2).height = 20;

      sheet.mergeCells('A3:E3');
      sheet.getCell('A3').value = headerLine3;
      sheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'center' };
      sheet.getRow(3).height = 20;

      sheet.mergeCells('A4:E4');
      sheet.getCell('A4').value = headerTitle;
      sheet.getCell('A4').alignment = { vertical: 'middle', horizontal: 'center' };
      sheet.getRow(4).height = 20;

      // Add Column Headers
      sheet.columns = [
        { header: 'Sl. No.', key: 'slNo', width: 10 },
        { header: "Lender's Code", key: 'lender_code', width: 20 },
        { header: "Lender's Name", key: 'lender_name', width: 30 },
        { header: 'Address - 1', key: 'lender_address_1', width: 40 },
        { header: 'Address - 2', key: 'lender_address_2', width: 40 },
        { header: 'Address - 3', key: 'lender_address_3', width: 40 },
      ];

      // Add Data Rows
      data.forEach((row, index) => {
        sheet.addRow({
          slNo: index + 1,
          lender_code: row.lender_code,
          lender_name: row.lender_name,
          lender_address_1: row.lender_address_1,
          lender_address_2: row.lender_address_2,
          lender_address_3: row.lender_address_3,
        });
      });

      // Set formatting for the header row
      sheet.getRow(4).font = { bold: true };

      // Return Excel file
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="LenderMasterReport.xlsx"');
      await workbook.xlsx.write(res);
      res.end();
    }

    // PDF Report
    else if (format === 'pdf') {
      const doc = new PDFDocument({ margin: 30 });
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="LenderMasterReport.pdf"');
      doc.pipe(res);

      // Add the header from .env
      doc.fontSize(16).text(headerLine1, { align: 'center' }).moveDown(0.5);
      doc.fontSize(12).text(headerLine2, { align: 'center' }).moveDown(0.5);
      doc.text(headerLine3, { align: 'center' }).moveDown(0.5);
      doc.text(headerTitle, { align: 'center' }).moveDown(1);

      // Add data rows
      data.forEach((row, index) => {
        doc.fontSize(12).text(`${index + 1}. ${row.lender_name} (${row.lender_code})`);
        doc.text(`Address: ${row.lender_address_1 || ''}, ${row.lender_address_2 || ''}, ${row.lender_address_3 || ''}`);
        doc.text(`Contact: ${row.lender_spoc_name || ''}, ${row.lender_spoc_contact || ''}, ${row.lender_spoc_email || ''}`);
        doc.moveDown();
      });

      doc.end();
    }

    // Word Report
    else if (format === 'word') {
      const doc = new Document({
        sections: [{
          properties: {},
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: headerLine1,
                  bold: true,
                  size: 28,
                }),
              ],
              alignment: "center",
            }),
            new Paragraph({
              children: [
                new TextRun(headerLine2),
              ],
              alignment: "center",
            }),
            new Paragraph({
              children: [
                new TextRun(headerLine3),
              ],
              alignment: "center",
            }),
            new Paragraph({
              children: [
                new TextRun(headerTitle),
                new TextRun("\n\n"),
              ],
              alignment: "center",
            }),
            ...data.flatMap((row, index) => ([ // Add each lender's info as a paragraph
              new Paragraph({
                children: [
                  new TextRun(`${index + 1}. ${row.lender_name} (${row.lender_code})`),
                ],
              }),
              new Paragraph(`Address: ${row.lender_address_1 || ''}, ${row.lender_address_2 || ''}, ${row.lender_address_3 || ''}`),
              new Paragraph(`Contact: ${row.lender_spoc_name || ''}, ${row.lender_spoc_contact || ''}, ${row.lender_spoc_email || ''}`),
              new Paragraph(""),
            ])),
          ],
        }],
      });

      const buffer = await Packer.toBuffer(doc);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', 'attachment; filename="LenderMasterReport.docx"');
      res.send(buffer);
    } else {
      res.status(400).json({ error: 'Invalid format selected' });
    }

  } catch (error) {
    console.error('Error generating report:', error);
    res.status(500).send('Server Error');
  }
};
