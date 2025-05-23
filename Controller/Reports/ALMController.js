const express = require('express');
const ExcelJS = require('exceljs');
const moment = require('moment');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');
require('dotenv').config();
const models = initModels(sequelize);
const { repayment_schedule } = models;

exports.almreport = async (req, res) => {
  console.log("Incoming request body:", req.body);
  const { date } = req.body;

  if (!date || !moment(date, 'YYYY-MM-DD', true).isValid()) {
    return res.status(400).json({
      error: 'Invalid or missing date. Expected format: YYYY-MM-DD',
    });
  }

  try {
    const result = await sequelize.query(
      `SELECT
        TO_CHAR(due_date, 'Mon-YY') AS period,
        EXTRACT(MONTH FROM due_date) AS month,
        EXTRACT(YEAR FROM due_date) AS year,
        SUM(principal_due) AS principal_due
      FROM repayment_schedule
      WHERE due_date > :date
      GROUP BY period, month, year
      ORDER BY year, month`,
      {
        replacements: { date },
        type: sequelize.QueryTypes.SELECT,
      }
    );
    
    if (!result || result.length === 0) {
      return res.status(404).json({ message: 'No records found for the selected filters.' });
    }
    const rows = result;
    const totalPrincipal = rows.reduce((sum, row) => sum + Number(row.principal_due), 0);
    let weightedMonthSum = 0;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('ALM Report');

    sheet.mergeCells('A1:E1');
    const titleCell = sheet.getCell('A1');
    titleCell.value = 'SRIFIN CREDIT PRIVATE LIMITED';
    titleCell.font = { bold: true, name: 'Arial', size: 11 };
    titleCell.alignment = { horizontal: 'center' };
    sheet.addRow([]);
    sheet.mergeCells('A3:E3');
    const addressCell = sheet.getCell('A3');
    addressCell.value =
      'Unit No. 509, 5th Floor, Gowra Fountainhead, Sy. No. 83(P) & 84(P),\nPatrika Nagar, Madhapur, Hitech City, Hyderabad - 500081, Telangana.';
    addressCell.font = { bold: true, name: 'Arial', size: 11 };
    addressCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true, };
    sheet.addRow([]);
    sheet.mergeCells('A5:E5');
    const titleRow = sheet.getCell('A5');
    titleRow.value = `Outstanding as on ${moment(date).format('DD-MMM-YY')} `;
    titleRow.alignment = { horizontal: 'center' };
    titleRow.font = { bold: true, name: 'Arial' };
    sheet.addRow([]);
    // Column Headers
    const headerRow = sheet.addRow(['Period', 'Month', 'Principal Due (In ₹)', 'Weight', 'Weighted Avg. Month',]);
    headerRow.font = { bold: true, name: 'Arial', size: 11 };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    headerRow.eachCell((cell) => {
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, };
    });

    // Data Rows
    rows.forEach((row, index) => {
      const weight = Number(row.principal_due) / totalPrincipal;
      const weightedAvg = weight * (index + 1);
      weightedMonthSum += weightedAvg;

      const newRow = sheet.addRow([
        row.period,
        index + 1,
        Number(row.principal_due),
        weight.toFixed(2),
        weightedAvg.toFixed(2),
      ]);
      newRow.font = { name: 'Arial', size: 10 };
      newRow.alignment = { horizontal: 'center', vertical: 'middle' };
      newRow.eachCell((cell) => {
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, };
      });
    });

    // Totals Row
    const totalRow = sheet.addRow(['Total', '', totalPrincipal.toFixed(2), '', weightedMonthSum.toFixed(2),]);
    totalRow.font = { bold: true, name: 'Arial', size: 10 };
    totalRow.alignment = { horizontal: 'center', vertical: 'middle' };
    totalRow.eachCell((cell) => {
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, };
    });

    sheet.columns = [{ width: 12 }, { width: 10 }, { width: 22 }, { width: 10 }, { width: 22 },];

    // Set response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=alm-report-${moment(date).format('YYYY-MM-DD')}.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Error generating ALM Report Excel:', err);
    res.status(500).send('Error generating report');
  }
};
