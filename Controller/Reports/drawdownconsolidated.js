// const ExcelJS = require('exceljs');
// const moment = require('moment');
// const { Op } = require('sequelize');
// const { sequelize } = require('../../config/db');
// const initModels = require('../../models/init-models');
// const models = initModels(sequelize);
// const { tranche_details, sanction_details, lender_master } = models;

// const OFF_BS_FACILITY_TYPES = ['Securitization', 'Factoring']; // Add more as needed

// function getFinancialYear(date) {
//   const m = moment(date);
//   const year = m.year();
//   const month = m.month(); // 0-indexed
//   const startYear = month < 3 ? year - 1 : year;
//   const endYear = startYear + 1;
//   return `FY${startYear % 100}-${endYear % 100}`;
// }

// function formatAmount(value) {
//   if (!value) return '-';
//   return Number(value).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
// }

// function isOffBalanceSheet(facilityType, off_bs_flag) {
//   if ((off_bs_flag || '').toUpperCase() === 'Y') return true;
//   if (!facilityType) return false;
//   return OFF_BS_FACILITY_TYPES.some(type =>
//     facilityType.toLowerCase() === type.toLowerCase()
//   );
// }

// exports.generateDrawdownReportconsolidated = async (req, res) => {
//   try {
//     const { year, format } = req.body;

//     if (!year || !format) {
//       return res.status(400).json({ error: 'Missing required fields: year and format' });
//     }
//     if (format.toLowerCase() !== 'excel') {
//       return res.status(400).json({ error: 'Only Excel format is supported currently' });
//     }

//     let drawdowns;
//     if (year.toLowerCase() === 'consolidated') {
//       drawdowns = await tranche_details.findAll({
//         include: [
//           {
//             model: sanction_details,
//             as: 'sanction',
//             include: [{ model: lender_master, as: 'lender_code_lender_master' }]
//           }
//         ],
//         order: [['tranche_date', 'ASC']]
//       });
//     } else {
//       // Parse FY year range
//       const fyMatch = year.match(/^FY(\d{2,4})(?:[-/]?(\d{2,4}))?$/i);
//       if (!fyMatch) {
//         return res.status(400).json({ error: 'Year format invalid. Use FY25 or FY2023-FY2024.' });
//       }

//       const startYearStr = fyMatch[1];
//       const endYearStr = fyMatch[2];
//       const startYear = startYearStr.length === 2 ? parseInt('20' + startYearStr) : parseInt(startYearStr);
//       const endYear = endYearStr
//         ? (endYearStr.length === 2 ? parseInt('20' + endYearStr) : parseInt(endYearStr))
//         : startYear + 1;

//       const startDate = new Date(`${startYear}-04-01`);
//       const endDate = new Date(`${endYear}-03-31`);

//       drawdowns = await tranche_details.findAll({
//         where: { tranche_date: { [Op.between]: [startDate, endDate] } },
//         include: [
//           {
//             model: sanction_details,
//             as: 'sanction',
//             include: [{ model: lender_master, as: 'lender_code_lender_master' }]
//           }
//         ],
//         order: [['tranche_date', 'ASC']]
//       });
//     }

//     // Group by FY
//     const groupedByFY = {};
//     drawdowns.forEach(d => {
//       const fy = getFinancialYear(d.tranche_date);
//       if (!groupedByFY[fy]) groupedByFY[fy] = [];
//       groupedByFY[fy].push(d);
//     });

//     // Calculate summary totals for FY
//     const summaryTotals = {};
//     Object.entries(groupedByFY).forEach(([fy, tranches]) => {
//       let onBS = 0, offBS = 0;
//       tranches.forEach(t => {
//         const amount = t.tranche_amount || 0;
//         const facilityType = t.sanction?.loan_type;
//         const off_bs_flag = t.sanction?.off_bs_flag;
//         if (isOffBalanceSheet(facilityType, off_bs_flag)) {
//           offBS += amount;
//         } else {
//           onBS += amount;
//         }
//       });
//       summaryTotals[fy] = { onBS, offBS, total: onBS + offBS };
//     });

//     // Create workbook & worksheet
//     const workbook = new ExcelJS.Workbook();
//     const sheet = workbook.addWorksheet('Drawdowns Report');

//     // Title
//     sheet.mergeCells('A1', 'G1');
//     sheet.getCell('A1').value = `Drawdowns Report : as on ${moment().format('DD-MMM-YY')}`;
//     sheet.getCell('A1').font = { size: 14, bold: true };
//     sheet.addRow([]);

//     // Summary table header
//     sheet.addRow(['', 'Drawdowns', 'On BS', 'Off BS', '', 'Total', '']);
//     const summaryHeader = sheet.addRow(['', '', '', '', '', '', '']);
//     summaryHeader.font = { bold: true };

//     // Summary data rows
//     for (const fy of Object.keys(summaryTotals).sort()) {
//       const val = summaryTotals[fy];
//       sheet.addRow([
//         fy,
//         formatAmount(val.total),
//         formatAmount(val.onBS),
//         formatAmount(val.offBS),
//         '',
//         formatAmount(val.total),
//         ''
//       ]);
//     }
//     sheet.addRow([]);

//     // Detailed table header
//     const detailHeader = [
//       'Drawdown Date',
//       'Bank / FI',
//       'Facility Type',
//       'On BS (₹ Cr)',
//       'Off BS (₹ Cr)',
//       'Processing Fee (%)',
//       'Processing Fee (₹ Cr)',
//       'Remarks'
//     ];
//     const headerRow = sheet.addRow(detailHeader);
//     headerRow.font = { bold: true };

//     // Detailed rows grouped by FY with subtotals
//     for (const fy of Object.keys(groupedByFY).sort()) {
//       const tranches = groupedByFY[fy];

//       // FY header row
//       const fyHeaderRow = sheet.addRow([fy]);
//       fyHeaderRow.font = { bold: true };

//       let subOnBS = 0, subOffBS = 0, subPF = 0;

//       for (const t of tranches) {
//         const dateStr = moment(t.tranche_date).format('D-MMM-YY');
//         const lender = t.sanction?.lender_code_lender_master?.lender_name || 'N/A';
//         const facility = t.sanction?.loan_type || 'N/A';
//         const amount = t.tranche_amount || 0;
//         const pfPercent = t.processing_fee_percentage || 0;  // Assume field exists
//         const pfAmount = t.processing_fee_amount || 0;       // Assume field exists
//         const remarks = t.remarks || (t.processing_fee_on === 'Drawdown' ? 'On Drawdown' : 'On Sanction') || '';

//         const isOffBS = isOffBalanceSheet(facility, t.sanction?.off_bs_flag);

//         const onBSVal = isOffBS ? '' : formatAmount(amount);
//         const offBSVal = isOffBS ? formatAmount(amount) : '';

//         sheet.addRow([
//           dateStr,
//           lender,
//           facility,
//           onBSVal,
//           offBSVal,
//           pfPercent ? `${pfPercent}%` : '-',
//           pfAmount ? formatAmount(pfAmount) : '-',
//           remarks
//         ]);

//         if (isOffBS) subOffBS += amount;
//         else subOnBS += amount;
//         subPF += pfAmount || 0;
//       }

//       // Subtotal row per FY
//       const subtotalRow = sheet.addRow([
//         'Grand Total',
//         '',
//         '',
//         formatAmount(subOnBS),
//         formatAmount(subOffBS),
//         '',
//         formatAmount(subPF),
//         ''
//       ]);
//       subtotalRow.font = { bold: true };
//       sheet.addRow([]);
//     }

//     // Set column widths for readability
//     sheet.columns = [
//       { width: 15 },
//       { width: 40 },
//       { width: 20 },
//       { width: 15 },
//       { width: 15 },
//       { width: 18 },
//       { width: 20 },
//       { width: 20 }
//     ];

//     // Send workbook as Excel file response
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     res.setHeader('Content-Disposition', `attachment; filename=Drawdowns_Report_${moment().format('YYYYMMDD')}.xlsx`);
//     await workbook.xlsx.write(res);
//     res.end();

//   } catch (err) {
//     console.error('Error generating drawdown report:', err);
//     res.status(500).json({ error: 'Internal Server Error' });
//   }
// };
const ExcelJS = require('exceljs');
const moment = require('moment');
const { Op } = require('sequelize');
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');
const models = initModels(sequelize);
const { tranche_details, sanction_details, lender_master } = models;

const OFF_BS_FACILITY_TYPES = ['Securitization', 'Factoring']; // Add more as needed

function getFinancialYear(date) {
  const m = moment(date);
  const year = m.year();
  const month = m.month(); // 0-indexed
  const startYear = month < 3 ? year - 1 : year;
  const endYear = startYear + 1;
  return `FY${startYear % 100}-${endYear % 100}`;
}

function formatAmount(value) {
  if (!value) return '-';
  return Number(value).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function isOffBalanceSheet(facilityType, off_bs_flag) {
  if ((off_bs_flag || '').toUpperCase() === 'Y') return true;
  if (!facilityType) return false;
  return OFF_BS_FACILITY_TYPES.some(type =>
    facilityType.toLowerCase() === type.toLowerCase()
  );
}

exports.generateDrawdownReportconsolidated = async (req, res) => {
  try {
    const { year, format } = req.body;
    // year = "consolidated"

    if (!year || !format) {
      return res.status(400).json({ error: 'Missing required fields: year and format' });
    }
    if (format.toLowerCase() !== 'excel') {
      return res.status(400).json({ error: 'Only Excel format is supported currently' });
    }

    let drawdowns;
    if (year.toLowerCase() === 'consolidated') {
      drawdowns = await tranche_details.findAll({
        include: [
          {
            model: sanction_details,
            as: 'sanction',
            include: [{ model: lender_master, as: 'lender_code_lender_master' }]
          }
        ],
        order: [['tranche_date', 'ASC']]
      });
    } else {
      // Parse FY year range
      const fyMatch = year.match(/^FY(\d{2,4})(?:[-/]?(\d{2,4}))?$/i);
      // console.log("fymatch "+fyMatch);
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
      // console.log("start "+startDate+" "+endDate);
      drawdowns = await tranche_details.findAll({
        where: { tranche_date: { [Op.between]: [startDate, endDate] } },
        include: [
          {
            model: sanction_details,
            as: 'sanction',
            include: [{ model: lender_master, as: 'lender_code_lender_master' }]
          }
        ],
        order: [['tranche_date', 'ASC']]
      });
    }
    // console.log("drwa "+JSON.stringify(drawdowns));
    if (!drawdowns || drawdowns.length === 0) {
      return res.status(404).json({ message: 'No records found for the selected filters.' });
    }
    // console.log('Raw Drawdowns:', JSON.stringify(drawdowns, null, 2));

    // Group by FY
    const groupedByFY = {};
    drawdowns.forEach(d => {
      const fy = getFinancialYear(d.tranche_date);
      if (!groupedByFY[fy]) groupedByFY[fy] = [];
      groupedByFY[fy].push(d);
    });

    
// console.log('Grouped by Financial Year:', JSON.stringify(groupedByFY, null, 2));

    // Calculate summary totals for FY
    const summaryTotals = {};
    Object.entries(groupedByFY).forEach(([fy, tranches]) => {
      let onBS = 0, offBS = 0;
      console.log("fy "+fy);
      tranches.forEach(t => {
        const amount = parseFloat(t.tranche_amount || 0);
        // console.log("Tranche Amount:", t.tranche_amount);
        const facilityType = t.sanction?.loan_type;
        const off_bs_flag = t.sanction?.off_bs_flag;
        if (isOffBalanceSheet(facilityType, off_bs_flag)) {
          offBS += amount;
        } else {
          onBS += amount;
        }
      });
      summaryTotals[fy] = { onBS, offBS, total: onBS + offBS };
    });
    // console.log('Summary Totals by FY:', summaryTotals);

    // Create workbook & worksheet
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Drawdowns Report');

    // Title
    sheet.mergeCells('A1');
    sheet.getCell('A1').value = `Drawdowns Report : as on ${moment().format('DD-MMM-YY')}`;
    sheet.getCell('A1').font = { size: 14, bold: true };
    sheet.addRow([]);

    // Summary table header
    sheet.addRow(['Drawdowns', 'On BS', 'Off BS', '', 'Total']); // A4 to E4

    // Optional: Style the header
    const headerRowTop = sheet.getRow(4);
    headerRowTop.font = { bold: true };
    headerRowTop.alignment = { vertical: 'middle', horizontal: 'center' };

    // Optionally set column widths for better appearance
    sheet.columns = [
      { key: 'drawdowns', width: 20 },
      { key: 'on_bs', width: 15 },
      { key: 'off_bs', width: 15 },
      { key: 'spacer', width: 5 }, // column D left blank
      { key: 'total', width: 15 }
    ];

    // Summary data rows
    for (const fy of Object.keys(summaryTotals).sort()) {
      const val = summaryTotals[fy];
      const row = [
          fy,
          formatAmount(val.total),
          formatAmount(val.onBS),
          formatAmount(val.offBS),
          formatAmount(val.total),
          ''
        ];
        sheet.addRow(row);
        // console.log(`Added Excel row for FY ${fy}:`, row);
    }
    sheet.addRow([]);

    // Detailed table header
    const detailHeader = [
      'Drawdown Date',
      'Bank / FI',
      'Facility Type',
      'Drawdown Amount (₹ Cr)',
      '',
      'Processing Fee (%)',
      'Processing Fee (₹ Cr)',
      'Remarks'
    ];
    const headerRow = sheet.addRow(detailHeader);
    headerRow.font = { bold: true };

    // Detailed rows grouped by FY with subtotals
    for (const fy of Object.keys(groupedByFY).sort()) {
      const tranches = groupedByFY[fy];

      // FY header row
      const fyHeaderRow = sheet.addRow([fy]);
      fyHeaderRow.font = { bold: true };

      let subOnBS = 0, subOffBS = 0, subPF = 0;

      for (const t of tranches) {
        const dateStr = moment(t.tranche_date).format('D-MMM-YY');
        const lender = t.sanction?.lender_code_lender_master?.lender_name || 'N/A';
        const facility = t.sanction?.loan_type || 'N/A';
        const amount = t.tranche_amount || 0;
        const pfPercent = t.sanction?.processing_fee || 0;
        const pfAmount = (amount * pfPercent) / 100;
        // const pfPercent = t.processing_fee_percentage || 0;  // Assume field exists
        // const pfAmount = t.processing_fee_amount || 0;       // Assume field exists
        const remarks = t.remarks || (t.processing_fee_on === 'Drawdown' ? 'On Drawdown' : 'On Sanction') || '';

        const isOffBS = isOffBalanceSheet(facility, t.sanction?.off_bs_flag);

        const onBSVal = isOffBS ? '' : formatAmount(amount);
        const offBSVal = isOffBS ? formatAmount(amount) : '';

        sheet.addRow([
          dateStr,
          lender,
          facility,
          onBSVal,
          offBSVal,
          pfPercent ? `${pfPercent}%` : '-',
          pfAmount ? formatAmount(pfAmount) : '-',
          remarks
        ]);

        if (isOffBS) subOffBS += parseFloat(amount);
        else subOnBS += parseFloat(amount);
        subPF += pfAmount || 0;
      }

      // Subtotal row per FY
      const subtotalRow = sheet.addRow([
        'Grand Total',
        '',
        '',
        formatAmount(subOnBS),
        formatAmount(subOffBS),
        '',
        formatAmount(subPF),
        ''
      ]);
      subtotalRow.font = { bold: true };
      sheet.addRow([]);
    }

    // Set column widths for readability
    sheet.columns = [
      { width: 15 },
      { width: 40 },
      { width: 20 },
      { width: 15 },
      { width: 15 },
      { width: 18 },
      { width: 20 },
      { width: 20 }
    ];

    // Send workbook as Excel file response
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Drawdowns_Report_${moment().format('YYYYMMDD')}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('Error generating drawdown report:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
};
