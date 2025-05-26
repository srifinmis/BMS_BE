require('dotenv').config();
const ExcelJS = require('exceljs');
// const { Table } = require('pdfkit-table');
const PDFDocument = require('pdfkit');
const { table } = require('pdfkit-table');
const { Document, Packer,Table, Paragraph, TableCell, TableRow, TextRun, WidthType } = require('docx');
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
    if (!data || data.length === 0) {
      return res.status(404).json({ message: 'No records found for the selected filters.' });
    }
    // Excel Report
    if (format === 'excel') {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Lender Master');
      // Static Header Line 1
      sheet.mergeCells('A1:F1');
      const cell1 = sheet.getCell('A1');
      cell1.value = headerLine1;
      cell1.alignment = { vertical: 'middle', horizontal: 'center' };
      cell1.font = { size: 14, bold: true };
      cell1.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEFEFEF' },
      };
      sheet.getRow(1).height = 30;

      // Static Header Line 2
      sheet.mergeCells('A3:F3');
      const cell3 = sheet.getCell('A3');
      cell3.value = headerLine2;
      cell3.alignment = { vertical: 'middle', horizontal: 'center' };
      cell3.font = { size: 14, bold: true };
      cell3.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEFEFEF' },
      };
      sheet.getRow(3).height = 30;

      // Static Header Line 3
      sheet.mergeCells('A4:F4');
      const cell4 = sheet.getCell('A4');
      cell4.value = headerLine3;
      cell4.alignment = { vertical: 'middle', horizontal: 'center' };
      cell4.font = { size: 14, bold: true };
      cell4.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEFEFEF' },
      };
      sheet.getRow(4).height = 30;

      // Static Header Line 3
      sheet.mergeCells('A5:F5');
      const cell5 = sheet.getCell('A5');
      cell5.value = headerTitle;
      cell5.alignment = { vertical: 'middle', horizontal: 'center' };
      cell5.font = { size: 14, bold: true };
      cell5.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEFEFEF' },
      };
      sheet.getRow(5).height = 30;
      // --- Column Headers at Row 7 ---
      const headerRowValues = [
        'Sl. No.',
        "Lender's Code",
        "Lender's Name",
        'Address - 1',
        'Address - 2',
        'Address - 3',
      ];
      sheet.getRow(7).values = headerRowValues;
      sheet.getRow(7).font = { bold: true };
      sheet.getRow(7).alignment = { vertical: 'middle', horizontal: 'center' };
      sheet.getRow(7).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEFEFEF' },
      };
      sheet.getRow(7).height = 20;

      // Apply border to each header cell
      sheet.getRow(7).eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
      // Optional: Set column widths
      sheet.columns = [
        { key: 'slNo', width: 10 },
        { key: 'lender_code', width: 20 },
        { key: 'lender_name', width: 30 },
        { key: 'lender_address_1', width: 40 },
        { key: 'lender_address_2', width: 40 },
        { key: 'lender_address_3', width: 40 },
      ];

      // --- Data Rows start at row 8 ---
      data.forEach((row, index) => {
        const newRow = sheet.addRow({
          slNo: index + 1,
          lender_code: row.lender_code,
          lender_name: row.lender_name,
          lender_address_1: [row.addr1_line1, row.addr1_contact1, row.addr1_email1, row.addr1_spoc_name].filter(Boolean).join(' '),
          lender_address_2: [row.addr2_line1, row.addr2_contact1, row.addr2_email1, row.addr2_spoc_name].filter(Boolean).join(' '),
          lender_address_3: [row.addr3_line1, row.addr3_contact1, row.addr3_email1, row.addr3_spoc_name].filter(Boolean).join(' '),

        });
        // Apply borders to each cell in the row
        // Apply borders to each cell from column 1 to 6 (A to F), even if null
        for (let col = 1; col <= 6; col++) {
          const cell = newRow.getCell(col);
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        }
      });

      // Return Excel file
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="LenderMasterReport.xlsx"');
      await workbook.xlsx.write(res);
      res.end();
    }

    // PDF Report
    // const PDFDocument = require('pdfkit');

    else if (format === 'pdf') {
      const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 40 });

      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="LenderMasterReport.pdf"');

      doc.pipe(res);

      // Header Lines
      doc.fontSize(20).font('Helvetica-Bold').text(headerLine1, { align: 'center' });
      doc.moveDown(0.5);
      doc.fontSize(14).font('Helvetica').text(headerLine2, { align: 'center' });
      doc.text(headerLine3, { align: 'center' });
      doc.text(headerTitle, { align: 'center' });
      doc.moveDown(1);

      // Table setup
      const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;

      const columnWidths = {
        slNo: 0.05 * pageWidth,
        lenderCode: 0.15 * pageWidth,
        lenderName: 0.20 * pageWidth,
        addr1: 0.20 * pageWidth,
        addr2: 0.20 * pageWidth,
        addr3: 0.20 * pageWidth,
      };

      const headers = [
        { text: 'Sl. No.', width: columnWidths.slNo },
        { text: "Lender's Code", width: columnWidths.lenderCode },
        { text: "Lender's Name", width: columnWidths.lenderName },
        { text: 'Address - 1', width: columnWidths.addr1 },
        { text: 'Address - 2', width: columnWidths.addr2 },
        { text: 'Address - 3', width: columnWidths.addr3 },
      ];

      const startX = doc.page.margins.left;
      let y = doc.y;
      const padding = 5; // padding inside each cell

      // Function to draw table header
      function drawTableHeader(yPos) {
        let x = startX;
        doc.font('Helvetica-Bold').fontSize(10);
        headers.forEach(header => {
          doc.rect(x, yPos, header.width, 20).stroke();
          doc.text(header.text, x + padding, yPos + padding, { width: header.width - 2 * padding, align: 'left' });
          x += header.width;
        });
        return yPos + 20;
      }

      // Draw first header
      y = drawTableHeader(y);

      doc.font('Helvetica').fontSize(9);

      data.forEach((row, index) => {
        const rowValues = [
          String(index + 1),
          row.lender_code || '',
          row.lender_name || '',
          [row.addr1_line1, row.addr1_contact1, row.addr1_email1, row.addr1_spoc_name].filter(Boolean).join('\n'),
          [row.addr2_line1, row.addr2_contact1, row.addr2_email2, row.addr2_spoc_name].filter(Boolean).join('\n'),
          [row.addr3_line1, row.addr3_contact1, row.addr3_email3, row.addr3_spoc_name].filter(Boolean).join('\n'),
        ];

        // Calculate max height needed by wrapped text in this row (add padding)
        const heights = rowValues.map((text, i) => {
          return doc.heightOfString(text, { width: Object.values(columnWidths)[i] - 2 * padding, align: 'left' }) + 2 * padding;
        });

        const rowHeight = Math.max(...heights, 20); // minimum row height

        let x = startX;

        // Check for page break: if row exceeds page height, add new page and redraw header
        if (y + rowHeight > doc.page.height - doc.page.margins.bottom) {
          doc.addPage({ size: 'A4', layout: 'landscape', margin: 40 });
          y = doc.y;
          y = drawTableHeader(y);
          doc.font('Helvetica').fontSize(9);
        }

        // Draw each cell: border + text with padding
        rowValues.forEach((text, i) => {
          const width = Object.values(columnWidths)[i];
          doc.rect(x, y, width, rowHeight).stroke();
          doc.text(text, x + padding, y + padding, {
            width: width - 2 * padding,
            align: 'left',
          });
          x += width;
        });

        y += rowHeight;
      });

      doc.end();
    }
    // --- Word ---
    else if (format === 'word') {
      const tableRows = [
        new TableRow({
          children: [
            { text: 'Sl. No.', width: 5 },
            { text: "Lender's Code", width: 15 },
            { text: "Lender's Name", width: 20 },
            { text: 'Address - 1', width: 20 },
            { text: 'Address - 2', width: 20 },
            { text: 'Address - 3', width: 20 },
          ].map(col =>
            new TableCell({
              width: {
                size: col.width,
                type: "pct",
              },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: col.text, bold: true })],
                }),
              ],
            })
          ),
        }),
        ...data.map((row, index) =>
          new TableRow({
            children: [
              { text: String(index + 1), width: 5 },
              { text: row.lender_code || '', width: 15 },
              { text: row.lender_name || '', width: 20 },
              {
                text: [row.addr1_line1, row.addr1_contact1, row.addr1_email1, row.addr1_spoc_name].filter(Boolean).join(' '),
                width: 20,
              },
              {
                text: [row.addr2_line1, row.addr2_contact1, row.addr2_email1, row.addr2_spoc_name].filter(Boolean).join(' '),
                width: 20,
              },
              {
                text: [row.addr3_line1, row.addr3_contact1, row.addr3_email1, row.addr3_spoc_name].filter(Boolean).join(' '),
                width: 20,
              },
            ].map(col =>
              new TableCell({
                width: {
                  size: col.width,
                  type: "pct",
                },
                children: [new Paragraph({ text: col.text })],
              })
            ),
          })
        ),
      ];

      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                size: {
                  orientation: 'landscape',
                },
              },
            },
            children: [
              new Paragraph({
                children: [new TextRun({ text: headerLine1, bold: true, size: 28 })],
                alignment: "center",
              }),
              new Paragraph({ children: [new TextRun(headerLine2)], alignment: "center" }),
              new Paragraph({ children: [new TextRun(headerLine3)], alignment: "center" }),
              new Paragraph({ children: [new TextRun(headerTitle)], alignment: "center" }),
              new Paragraph({ text: "" }), // Spacer
              new Table({
                rows: tableRows,
                width: {
                  size: 100,
                  type: "pct",
                },
              }),
            ],
          },
        ],
      });

      const buffer = await Packer.toBuffer(doc);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', 'attachment; filename="LenderMasterReport.docx"');
      res.send(buffer);
    }

    else {
      res.status(400).json({ error: 'Invalid format selected' });
    }

  } catch (error) {
    console.error('Error generating report:', error);
    res.status(500).send('Server Error');
  }
};