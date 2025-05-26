const { sequelize } = require('../../config/db');
const ExcelJS = require('exceljs');
const moment = require('moment');

function getOrdinal(n) {
  const s = ["th", "st", "nd", "rd"];
  const v = n % 100;
  return s[(v - 20) % 10] || s[v] || s[0];
}

const getMasterReport = async (req, res) => {
  try {
    const query = `
      SELECT 
        lm.lender_name AS "NameoftheLender",
        sd.loan_type AS "FacilityType",
        sd.sanction_id AS "SanctionNo",
        sd.sanction_date AS "SanctionDate",
        sd.sanction_amount AS "SanctionAmountRs",
        td.tranche_id AS "TrenchNo",
        td.tranche_date AS "DateofAvailment",
        td.tranche_amount AS "DrawdownAmountRs",
        td.interest_rate AS "RateofInterestPA",
        td.interest_type AS "InterestType",
        sd.spread_floating AS "IfFloatingInterestCondition",
        td.moratorium_start_date,
        td.moratorium_end_date,
        rs.due_date AS "RepaymentDate",
        td.principal_payment_frequency AS "PrincipalRepayment",
        td.interest_payment_frequency AS "InterestPayment",
        sd.book_debt_margin AS "BookDebtsMargin",
        sd.cash_margin AS "CashDepositMargin",
        td.tenure_months AS "TenureInMonths",
        lm.status AS "LoanStatus",
        sd.processing_fee AS "ProcessingFee"
      FROM lender_master lm
      INNER JOIN sanction_details sd ON lm.lender_code = sd.lender_code
      INNER JOIN tranche_details td ON sd.sanction_id = td.sanction_id
      LEFT JOIN repayment_schedule rs ON td.tranche_id = rs.tranche_id
      ORDER BY TRIM(lm.lender_name), TRIM(sd.sanction_id), TRIM(td.tranche_id);
    `;
    // Execute the query
    if (!query || query.length === 0) {
      return res.status(404).json({ message: 'No records found for the selected filters.' });
    }

    const [results] = await sequelize.query(query);
    if (!results.length) return res.status(404).send('No data found for the report');

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Master Report');

    const startMonth = moment('2025-04-01');
    const endMonth = moment('2027-11-01');
    const monthHeaders = [];
    while (startMonth.isSameOrBefore(endMonth)) {
      monthHeaders.push(startMonth.format('MMM-YY'));
      startMonth.add(1, 'month');
    }

    const headerIndex = (title) => staticHeaders.findIndex(h => h === title);
    const staticHeaders = [
      'Name of the Lender', 'Facility\nType', 'Sanction\nNo.', 'Sanction\nDate', 'Sanction\nAmount\n(Rs.)', 'Trench\nNo.', 'Date of\nAvailment', 'Drawdown\nAmount\n(Rs.)',
      'Un-Drawdown\nAmount\n(Rs.)', 'Tranche\nOutstanding\n31-Mar-25 (Rs.)', 'Rate of\nInterest\nP. A', 'Interest\nType', 'If Floating, Interest Condition', 'MCLR /\nInterest\nReset',
      'Repayment\nTerms', 'Moratorium', 'Repayment Date', 'Loan\nClosure\nDate', 'Principal\nRepayment', 'Interest\nPayment', 'Book\nDebts\nFrequency', 'Book\nDebts\nMargin', 'Cash\nDeposit\nMargin',
      'Tenure\n(In Months)', 'No. of\nInstallments\nLeft', 'Loan\nStatus', ...monthHeaders, 'Total', '', 'Processing\nFee %', 'Processing\nFee\nAmount',
      'Processing\nFee per\nMonth', 'Processing\nFee\nAnnulised', '', 'Tranche\nOutstanding\n31-Dec-24 (Rs.)', 'Marging\nAmount', 'Book Debts\nNeed to Submit'
    ];
    const customColumnWidths = [
      40, 14, 9, 11, 15, 7, 11, 15, 16, 18, 9, 11, 28, 12, 15, 12, 22, 12, 12, 11, 11, 9, 9, 10, 12, 9,
      ...Array(monthHeaders.length).fill(12), 14, 12, 12, 14, 14, 14, 12, 18, 18, 18
    ];
    worksheet.columns = customColumnWidths.map(w => ({
      width: w,
      style: { alignment: { wrapText: true, horizontal: 'center', vertical: 'middle' } }
    }));

    worksheet.addRow([]);
    const headerRow = worksheet.addRow(staticHeaders);
    worksheet.views = [{ state: 'frozen', xSplit: 8, ySplit: 2 }];
    headerRow.height = 50;
    headerRow.eachCell((cell, colNumber) => {
      cell.value = staticHeaders[colNumber - 1];
      cell.font = { bold: true, name: 'Arial', size: 10 };
      cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }
      };
    });

    const groupedByTranche = {};
    results.forEach(row => {
      const trancheKey = row.TrenchNo.trim();
      groupedByTranche[trancheKey] = row;
    });

    const sanctionCounterMap = new Map();
    const trancheCounterMap = new Map();
    const facilityGroups = {};
    for (const row of Object.values(groupedByTranche)) {
      const facilityType = row.FacilityType?.trim() || 'Unknown';
      if (!facilityGroups[facilityType]) facilityGroups[facilityType] = [];
      facilityGroups[facilityType].push(row);
    }

    const subtotalRowsForGrandTotal = []; // Store Excel row objects for subtotals
    const allInterestRates = [];
    for (const [facilityType, rows] of Object.entries(facilityGroups)) {
      let lastLender = null;
      let lastSanctionKey = null;
      const uniqueTranches = new Set();
      const rowStartIndex = worksheet.rowCount + 1;

      for (const row of rows) {
        const trancheId = row.TrenchNo.trim();
        const lender = row.NameoftheLender.trim();
        const sanctionId = row.SanctionNo.trim();
        const originalSanctionId = row.SanctionNo.trim();
        const originalTrancheId = row.TrenchNo.trim();
        const facility = row.FacilityType?.trim() || 'Unknown';
        const lenderFacilityKey = `${lender}-${facility}`;

        if (!sanctionCounterMap.has(lenderFacilityKey)) {
          sanctionCounterMap.set(lenderFacilityKey, new Map());
        }
        const sanctionMap = sanctionCounterMap.get(lenderFacilityKey);
        if (!sanctionMap.has(originalSanctionId)) {
          sanctionMap.set(originalSanctionId, sanctionMap.size + 1);
        }
        const incrementalSanctionNo = sanctionMap.get(originalSanctionId);

        if (!trancheCounterMap.has(originalSanctionId)) {
          trancheCounterMap.set(originalSanctionId, new Map());
        }
        const trancheMap = trancheCounterMap.get(originalSanctionId);
        if (!trancheMap.has(originalTrancheId)) {
          trancheMap.set(originalTrancheId, trancheMap.size + 1);
        }
        const incrementalTrenchNo = trancheMap.get(originalTrancheId);

        const sanctionKey = `${lender}-${sanctionId}`;
        const isSameBlock = lastLender === lender;
        const isSameSanction = lastSanctionKey === sanctionKey;

        const drawdownAmount = Number(row.DrawdownAmountRs || 0);
        const tenure = Number(row.TenureInMonths || 0);
        const rate = Number(row.RateofInterestPA || 0);
        if (!isNaN(rate) && rate > 0) {
          allInterestRates.push(rate);
        }
        const procFee = Number(row.ProcessingFee || 0);
        const availDate = moment(row.DateofAvailment);

        const repayments = await sequelize.query(
          `SELECT principal_due, total_due, due_date, repayment_type FROM repayment_schedule WHERE tranche_id = :trancheId`,
          {
            replacements: { trancheId: row.TrenchNo },
            type: sequelize.QueryTypes.SELECT
          }
        );

        const principalPaidTillMar25 = repayments
          .filter(r => moment(r.due_date).isSameOrBefore('2025-03-31'))
          .reduce((sum, r) => sum + Number(r.principal_due || 0), 0);
        const trancheOut31Mar25 = Math.max(0, +(drawdownAmount - principalPaidTillMar25).toFixed(2));
        const principalPaidTillDec24 = repayments
          .filter(r => moment(r.due_date).isSameOrBefore('2024-12-31'))
          .reduce((sum, r) => sum + Number(r.principal_due || 0), 0);
        const trancheOut31Dec24 = Math.max(0, +(drawdownAmount - principalPaidTillDec24).toFixed(2));

        const closureDate = availDate.clone().add(tenure, 'months');
        const processingFeeAmount = +(drawdownAmount * procFee / 100).toFixed(2);
        const procFeeMonth = +(processingFeeAmount / tenure).toFixed(2);
        const procFeeAnn = +(procFeeMonth * 12).toFixed(2);
        const marginAmount = +(trancheOut31Dec24 * (row.BookDebtsMargin / 100)).toFixed(2);
        const bookDebtsToSubmit = +(trancheOut31Dec24 + marginAmount).toFixed(2);

        const totalPerMonth = {};
        repayments.forEach(r => {
          const dueMonth = moment(r.due_date).format('MMM-YY');
          if (moment(r.due_date).isAfter('2025-03-31')) {
            totalPerMonth[dueMonth] = (totalPerMonth[dueMonth] || 0) + Number(r.total_due || 0);
          }
        });
        const monthlyValues = monthHeaders.map(label => +(totalPerMonth[label] || 0).toFixed(2));
        const monthlyTotal = monthlyValues.reduce((sum, val) => sum + val, 0);
        const noOfInstLeft = Object.keys(totalPerMonth).length;

        const repaymentTermsSet = new Set(repayments.map(r => r.repayment_type).filter(Boolean));
        const repaymentTypeMap = { emi: 'Monthly', ewi: 'Weekly', eqi: 'Quarterly', ayi: 'Annual' };
        const repaymentTerms = Array.from(repaymentTermsSet)
          .map(type => repaymentTypeMap[type?.toLowerCase()] || type || 'N/A')
          .join(', ') || 'N/A';
        let repaymentDay = 'N/A';
        if (repayments.length) {
          const firstDate = moment(repayments[0].due_date);
          const day = firstDate.date();
          repaymentDay = `${day}${getOrdinal(day)} day of every month` || 'Different Dates';
        }

        const sanctionRows = results.filter(r => r.SanctionNo === sanctionId);
        const sanctionDrawdowns = new Set();
        const totalDrawdown = sanctionRows.reduce((sum, t) => {
          if (!sanctionDrawdowns.has(t.TrenchNo)) {
            sanctionDrawdowns.add(t.TrenchNo);
            return sum + Number(t.DrawdownAmountRs || 0);
          }
          return sum;
        }, 0);
        let undrawdown = +(row.SanctionAmountRs - totalDrawdown);
        if (undrawdown < 0) undrawdown = 0;

        let moratoriumDuration = 'N/A';
        if (row.moratorium_start_date && row.moratorium_end_date) {
          const start = moment(row.moratorium_start_date);
          const end = moment(row.moratorium_end_date);
          if (end.isAfter(start)) {
            moratoriumDuration = end.diff(start, 'Days', true) + 'Days';
          }
        }

        const rowData = [
          isSameBlock ? '' : lender,
          isSameBlock ? '' : facilityType,
          isSameSanction ? '' : incrementalSanctionNo,
          isSameSanction ? '' : moment(row.SanctionDate).format('DD-MMM-YY'),
          isSameSanction ? '' : row.SanctionAmountRs,
          incrementalTrenchNo,
          moment(row.DateofAvailment).format('DD-MMM-YY'),
          drawdownAmount,
          isSameSanction ? '' : undrawdown,
          trancheOut31Mar25,
          rate,
          row.InterestType,
          (row.InterestType?.toLowerCase() === 'fixed' ? 'N/A' : row.IfFloatingInterestCondition || 'N/A'),
          'N/A',
          repaymentTerms,
          moratoriumDuration,
          repaymentDay,
          closureDate.format('DD-MMM-YY'),
          row.PrincipalRepayment,
          row.InterestPayment,
          (row.facilityType?.toLowerCase() === 'Securitisation' ? 'N/A' : 'Monthly' || 'N/A'),
          row.BookDebtsMargin,
          row.CashDepositMargin,
          tenure,
          noOfInstLeft,
          row.LoanStatus,
          ...monthlyValues,
          monthlyTotal,
          '',
          procFee,
          processingFeeAmount,
          procFeeMonth,
          procFeeAnn,
          '',
          trancheOut31Dec24,
          marginAmount,
          bookDebtsToSubmit
        ];

        const dataRow = worksheet.addRow(rowData);
        dataRow.height = 15;
        dataRow.font = { name: 'Arial', size: 10 };
        dataRow.alignment = { horizontal: 'center', vertical: 'middle' };
        dataRow.eachCell(cell => {
          cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
        });

        lastLender = lender;
        lastSanctionKey = sanctionKey;
        uniqueTranches.add(trancheId);
      }

      // Add subtotal for this facility group
      const lastRows = worksheet.getRows(rowStartIndex, uniqueTranches.size);
      const subtotalRow = [];
      const sum = (colIdx) => {
        if (colIdx <= 0) return 0;
        return lastRows.reduce((acc, row) => acc + (Number(row.getCell(colIdx).value) || 0), 0);
      };
      const avg = (colIdx) => {
        if (colIdx <= 0) return '';
        const valid = lastRows
          .map(r => r.getCell(colIdx).value)
          .filter(v => typeof v === 'number' && !isNaN(v));
        const total = valid.reduce((acc, val) => acc + val, 0);
        return valid.length ? +(total / valid.length).toFixed(2) : '';
      };
      subtotalRow[0] = `Sub-Total : ${facilityType}`;
      subtotalRow[headerIndex('Sanction\nAmount\n(Rs.)')] = sum(headerIndex('Sanction\nAmount\n(Rs.)') + 1);
      subtotalRow[headerIndex('Drawdown\nAmount\n(Rs.)')] = sum(headerIndex('Drawdown\nAmount\n(Rs.)') + 1);
      subtotalRow[headerIndex('Un-Drawdown\nAmount\n(Rs.)')] = sum(headerIndex('Un-Drawdown\nAmount\n(Rs.)') + 1);
      subtotalRow[headerIndex('Tranche\nOutstanding\n31-Mar-25 (Rs.)')] = sum(headerIndex('Tranche\nOutstanding\n31-Mar-25 (Rs.)') + 1);
      subtotalRow[headerIndex('Rate of\nInterest\nP. A')] = avg(headerIndex('Rate of\nInterest\nP. A') + 1);
      subtotalRow[headerIndex('Total')] = sum(headerIndex('Total') + 1);
      subtotalRow[headerIndex('Processing\nFee\nAmount')] = sum(headerIndex('Processing\nFee\nAmount') + 1);
      subtotalRow[headerIndex('Processing\nFee per\nMonth')] = sum(headerIndex('Processing\nFee per\nMonth') + 1);
      subtotalRow[headerIndex('Processing\nFee\nAnnulised')] = sum(headerIndex('Processing\nFee\nAnnulised') + 1);
      subtotalRow[headerIndex('Tranche\nOutstanding\n31-Dec-24 (Rs.)')] = sum(headerIndex('Tranche\nOutstanding\n31-Dec-24 (Rs.)') + 1);
      subtotalRow[headerIndex('Marging\nAmount')] = sum(headerIndex('Marging\nAmount') + 1);
      subtotalRow[headerIndex('Book Debts\nNeed to Submit')] = sum(headerIndex('Book Debts\nNeed to Submit') + 1);
      monthHeaders.forEach((label) => {
        const colIdx = headerIndex(label);
        if (colIdx >= 0) {
          subtotalRow[colIdx] = sum(colIdx + 1); // ExcelJS uses 1-based indexing
        }
      });

      const rowObj = worksheet.addRow(subtotalRow);
      subtotalRowsForGrandTotal.push(rowObj); // save for grand total
      rowObj.font = { bold: true, name: 'Arial', size: 10 };
      rowObj.alignment = { horizontal: 'center', vertical: 'middle' };
      for (let colIdx = 1; colIdx <= staticHeaders.length; colIdx++) {
        const cell = rowObj.getCell(colIdx);
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
      };
      worksheet.addRow([]); // Spacer before next group
    }

    const grandTotalRow = Array(staticHeaders.length).fill('');
    grandTotalRow[0] = 'Grand Total';
    const sumAcrossSubtotals = (colIdx) => {
      if (colIdx <= 0) return '';
      const values = subtotalRowsForGrandTotal
        .map(row => {
          const val = row.getCell(colIdx).value;
          if (typeof val === 'string') {
            return parseFloat(val.replace(/,/g, ''));
          }
          return typeof val === 'number' ? val : NaN;
        })
        .filter(v => !isNaN(v));
      return values.length ? (values.reduce((a, b) => a + b, 0)) : '';
    };
    const avgAcrossSubtotals = (colIdx) => {
      if (colIdx <= 0) return '';
      const values = subtotalRowsForGrandTotal
        .map(row => {
          const val = row.getCell(colIdx).value;
          if (typeof val === 'string') {
            return parseFloat(val.replace(/,/g, ''));
          }
          return typeof val === 'number' ? val : NaN;
        })
        .filter(v => !isNaN(v));
      return values.length ? (values.reduce((a, b) => a + b, 0) / values.length) : '';
    };
    // Columns to sum
    [
      'Sanction\nAmount\n(Rs.)', 'Drawdown\nAmount\n(Rs.)', 'Un-Drawdown\nAmount\n(Rs.)', 'Tranche\nOutstanding\n31-Mar-25 (Rs.)', 'Total', 'Processing\nFee\nAmount',
      'Processing\nFee per\nMonth', 'Processing\nFee\nAnnulised', 'Tranche\nOutstanding\n31-Dec-24 (Rs.)', 'Marging\nAmount', 'Book Debts\nNeed to Submit'
    ].forEach(title => {
      const idx = headerIndex(title) + 1;
      grandTotalRow[idx - 1] = sumAcrossSubtotals(idx);
    });
    // Average for interest
    const interestIdx = headerIndex('Rate of\nInterest\nP. A') + 1;
    grandTotalRow[interestIdx - 1] = avgAcrossSubtotals(interestIdx);
    // Monthly headers
    monthHeaders.forEach(label => {
      const colIdx = headerIndex(label);
      if (colIdx >= 0) {
        grandTotalRow[colIdx] = sumAcrossSubtotals(colIdx + 1);
      }
    });
    // Add the Grand Total row to worksheet
    const grandRow = worksheet.addRow(grandTotalRow);
    grandRow.font = { bold: true, name: 'Arial', size: 10 };
    grandRow.alignment = { horizontal: 'center', vertical: 'middle' };
    // Add borders
    for (let colIdx = 1; colIdx <= staticHeaders.length; colIdx++) {
      const cell = grandRow.getCell(colIdx);
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    }

    worksheet.columns.forEach((col, index) => {
      const header = staticHeaders[index] || '';
      const currencyHeaders = ['Sanction\nAmount\n(Rs.)', 'Drawdown\nAmount\n(Rs.)', 'Un-Drawdown\nAmount\n(Rs.)', 'Tranche\nOutstanding\n31-Mar-25 (Rs.)', 'Processing\nFee\nAmount',
        'Processing\nFee per\nMonth', 'Processing\nFee\nAnnulised', 'Tranche\nOutstanding\n31-Dec-24 (Rs.)', 'Marging\nAmount', 'Book Debts\nNeed to Submit', 'Total'];
      const isMonthlyHeader = /^[A-Za-z]{3}-\d{2}$/i.test(header);
      const isCurrency = header.includes('(Rs.)') || header.includes('Amount') || header.includes('Fee') || currencyHeaders.includes(header) || isMonthlyHeader;
      if (isCurrency) {
        col.numFmt = '#,##0';
      }
    });

    // Save file and send
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Master_Report.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error generating master report:', error);
    res.status(500).send('Internal Server Error');
  }
};

module.exports = { getMasterReport };