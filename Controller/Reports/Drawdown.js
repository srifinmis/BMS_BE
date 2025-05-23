const ExcelJS = require('exceljs');
const moment = require('moment');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');
const models = initModels(sequelize);
const { tranche_details, sanction_details, lender_master } = models;

// Define off-balance-sheet facility types here
const OFF_BS_FACILITY_TYPES = ['Securitization', 'Factoring']; // Add more off-BS facility types if needed

function getQuarter(date) {
  const month = moment(date).month(); // 0-indexed
  if (month >= 3 && month <= 5) return 'Q1';
  if (month >= 6 && month <= 8) return 'Q2';
  if (month >= 9 && month <= 11) return 'Q3';
  return 'Q4';
}

function getFinancialYear(date) {
  const m = moment(date);
  const year = m.year();
  const month = m.month(); // 0-indexed
  const startYear = month < 3 ? year - 1 : year;
  const endYear = startYear + 1;
  return `FY${startYear % 100}-${endYear % 100}`;
}

function formatAmount(value) {
  return Number(value).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Helper to decide if a tranche is off-balance-sheet
function isOffBalanceSheet(facilityType, off_bs_flag) {
  if ((off_bs_flag || '').toUpperCase() === 'Y') return true;
  if (!facilityType) return false;
  return OFF_BS_FACILITY_TYPES.some(type =>
    facilityType.toLowerCase() === type.toLowerCase()
  );
}

exports.generateDrawdownReport = async (req, res) => {
  const { year, format } = req.body;

  if (!year || !format) {
    return res.status(400).json({ error: 'Missing required fields: year and format' });
  }

  if (format.toLowerCase() !== 'excel') {
    return res.status(400).json({ error: 'Currently only Excel format is supported' });
  }

  try {
    let drawdowns;
    if (year.toLowerCase() === 'consolidated') {
      drawdowns = await tranche_details.findAll({
        include: [
          {
            model: sanction_details,
            as: 'sanction',
            include: [{ model: lender_master, as: 'lender_code_lender_master' }]
          }
        ]
      });
    } else {
      const fyMatch = year.match(/^FY(\d{2,4})(?:[-/]?(\d{2,4}))?$/i);
      if (!fyMatch) {
        return res.status(400).json({ error: 'Year format invalid. Use FY25 or FY2023-FY2024.' });
      }

      const startYearStr = fyMatch[1];
      const endYearStr = fyMatch[2];
      const startYear = startYearStr.length === 2 ? parseInt('20' + startYearStr) : parseInt(startYearStr);
      const endYear = endYearStr
        ? (endYearStr.length === 2 ? parseInt('20' + endYearStr) : parseInt(endYearStr))
        : startYear + 1;

      const startDate = new Date(`${startYear}-04-01`);
      const endDate = new Date(`${endYear}-03-31`);

      drawdowns = await tranche_details.findAll({
        where: { tranche_date: { [Op.between]: [startDate, endDate] } },
        include: [
          {
            model: sanction_details,
            as: 'sanction',
            include: [{ model: lender_master, as: 'lender_code_lender_master' }]
          }
        ]
      });
    }

    if (!drawdowns || drawdowns.length === 0) {
      return res.status(404).json({ message: 'No records found for the selected filters.' });
    }

    // Group drawdowns by quarter and FY
    const grouped = {};
    drawdowns.forEach(tranche => {
      const date = tranche.tranche_date;
      const fy = getFinancialYear(date);
      const qtr = getQuarter(date);
      const key = `${qtr} : ${fy}`;

      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(tranche);
    });

    // Prepare summary totals by quarter for summary table at top
    const summaryTotals = {};
    for (const [key, tranches] of Object.entries(grouped)) {
      let onBS = 0;
      let offBS = 0;
      tranches.forEach(tranche => {
        const amount = parseFloat(tranche.tranche_amount || 0);
        const facilityType = tranche.sanction?.loan_type;
        const off_bs_flag = tranche.sanction?.off_bs_flag;
        const isOffBS = isOffBalanceSheet(facilityType, off_bs_flag);

        if (isOffBS) offBS += parseFloat(amount);
        else onBS += parseFloat(amount);
      });
      summaryTotals[key] = { onBS, offBS, total: onBS + offBS };
    }

    // Calculate FY total row (sum all quarters)
    const fyTotals = {};
    Object.entries(summaryTotals).forEach(([key, val]) => {
      console.log("fy "+key);
      const fy = key.split(':')[1].trim();
      if (!fyTotals[fy]) fyTotals[fy] = { onBS: 0, offBS: 0, total: 0 };
      fyTotals[fy].onBS += val.onBS;
      fyTotals[fy].offBS += val.offBS;
      fyTotals[fy].total += val.total;
    });

    // Create Excel workbook and sheet
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Drawdowns Report');

    // Title
    sheet.mergeCells('A1', 'G1');
    sheet.getCell('A1').value = `Drawdowns Report : as on ${moment().format('DD-MMM-YY')}`;
    sheet.getCell('A1').font = { size: 14, bold: true };
    sheet.addRow([]);

    // Summary Table Headers
    sheet.addRow(['', 'Drawdowns', 'On BS', 'Off BS', '', 'Total', '']);
    const headerRow = sheet.addRow(['', '', '', '', '', '', '']);
    headerRow.font = { bold: true };

    // Summary data rows by quarter
    Object.keys(summaryTotals).sort().forEach(key => {
      const val = summaryTotals[key];
      sheet.addRow([
        key,
        formatAmount(val.total),
        formatAmount(val.onBS),
        formatAmount(val.offBS),
        '',
        formatAmount(val.total),
        ''
      ]);
    });

    // FY Total row
    for (const fy in fyTotals) {
      const val = fyTotals[fy];
      const fyLabel = `FY${fy}`;
      sheet.addRow([
        fyLabel,
        formatAmount(val.onBS + val.offBS),
        formatAmount(val.onBS),
        formatAmount(val.offBS),
        '',
        formatAmount(val.total),
        ''
      ]);
    }

    sheet.addRow([]);

    // Detailed Table Header
    const detailHeader = [
      'Drawdown Date',
      'Bank / FI',
      'Facility Type',
      'Drawdown Amount (Rs. Crs.)',
      '',
      'Processing Fee (% on Sanction/ Drawdown)',
      'Processing Fee (Rs. Crs.)'
    ];
    const header = sheet.addRow(detailHeader);
    header.font = { bold: true };

    // Initialize grand totals
    let grandOnBS = 0, grandOffBS = 0, grandPF = 0;

    // For each quarter group, add detail rows and subtotals
    for (const key of Object.keys(grouped).sort()) {
      // Quarter title row
      const titleRow = sheet.addRow([key]);
      titleRow.font = { bold: true };

      let subOnBS = 0, subOffBS = 0, subPF = 0;

      grouped[key].forEach(tranche => {
        const dateStr = moment(tranche.tranche_date).format('DD-MMM-YY');
        const lender = tranche.sanction?.lender_code_lender_master?.lender_name || 'N/A';
        const facility = tranche.sanction?.loan_type || 'N/A';
        const amount = parseFloat(tranche.tranche_amount || 0);
        const facilityType = tranche.sanction?.loan_type;
        const off_bs_flag = tranche.sanction?.off_bs_flag;
        const isOffBS = isOffBalanceSheet(facilityType, off_bs_flag);
        const pfPercent = tranche.sanction?.processing_fee || 0;
        const pfAmount = (amount * pfPercent) / 100;

        if (isOffBS) subOffBS += parseFloat(amount);
        else subOnBS += parseFloat(amount);
        subPF += parseFloat(pfAmount);

        sheet.addRow([
          dateStr,
          lender,
          facility,
          isOffBS ? '' : formatAmount(amount),
          isOffBS ? formatAmount(amount) : '',
          pfPercent ? pfPercent + '%' : '',
          pfAmount ? formatAmount(pfAmount) : ''
        ]);
      });

      // Subtotal row
      const subtotalRow = sheet.addRow([
        '',
        '',
        'Subtotal',
        formatAmount(subOnBS),
        formatAmount(subOffBS),
        '',
        formatAmount(subPF)
      ]);
      subtotalRow.font = { bold: true };

      grandOnBS += subOnBS;
      grandOffBS += subOffBS;
      grandPF += subPF;
    }

    // Grand total row
    const grandTotalRow = sheet.addRow([
      '',
      '',
      'Grand Total',
      formatAmount(grandOnBS),
      formatAmount(grandOffBS),
      '',
      formatAmount(grandPF)
    ]);
    grandTotalRow.font = { bold: true };

    // Adjust column widths
    sheet.columns = [
      { width: 15 },
      { width: 25 },
      { width: 20 },
      { width: 15 },
      { width: 15 },
      { width: 20 },
      { width: 20 }
    ];

    // Send Excel file response
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=drawdowns_report_${year}_${moment().format('YYYYMMDD_HHmmss')}.xlsx`
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error generating drawdown report:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
};
