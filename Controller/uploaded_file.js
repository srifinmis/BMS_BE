const moment = require("moment");
const jwt = require("jsonwebtoken");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const models = initModels(sequelize);
const { repayment_schedule_staging, payment_details_staging, payment_details } = models;

// Configure multer storage
const upload = multer({ dest: "uploads/" }).single("file");

function parseNumber(value) {
  console.log('Original value:', value);

  // Handle null, undefined, or empty string as 0, but only if there's no valid value
  if (value === undefined || value === null) {
    console.log('Returning 0 due to undefined or null');
    return 0;
  }

  // Handle empty string case explicitly
  if (value === "") {
    console.log('Returning 0 due to empty string');
    return 0;
  }

  // If value is already a number, return it directly
  if (typeof value === 'number') {
    console.log('Returning number:', value);
    return value;
  }

  // If value is a string, remove commas and try to parse it as a number
  if (typeof value === 'string') {
    const cleanedValue = value.replace(/,/g, '').trim(); // Remove commas and extra spaces
    const parsedValue = parseFloat(cleanedValue);

    console.log('Cleaned value:', cleanedValue);
    console.log('Parsed value:', parsedValue);

    // Check if the parsed value is a valid number
    if (isNaN(parsedValue)) {
      console.log('Returning 0 due to invalid number');
      return 0; // Default value when invalid input is encountered
    }

    return parsedValue;
  }

  // Return 0 for any other data type (just as a safety net)
  console.log('Returning 0 due to invalid type');
  return 0;
}

// exports.UTRuploadedFile = (req, res) => {
//   upload(req, res, async (err) => {
//     if (err) {
//       return res.status(400).json({ message: "File upload failed", error: err.message });
//     }

//     const { sanction_id, tranche_id, lender_code, due_date, loan_type, created_by } = req.body;
//     const JWT_SECRET = process.env.JWT_SECRET;

//     if (!created_by) {
//       return res.status(401).json({ message: "JWT token not provided" });
//     }

//     let decoded;
//     try {
//       decoded = jwt.verify(created_by, JWT_SECRET);
//     } catch (error) {
//       return res.status(401).json({ message: "Invalid token", error: error.message });
//     }

//     if (!req.file) {
//       return res.status(400).json({ message: "No file provided" });
//     }

//     const filePath = path.join(__dirname, "../", req.file.path);
//     const ext = path.extname(req.file.originalname).toLowerCase();

//     try {
//       let records = [];

//       if (ext === ".xlsx" || ext === ".xls") {
//         const workbook = XLSX.readFile(filePath);
//         const sheetName = workbook.SheetNames[0];
//         records = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
//       } else if (ext === ".csv") {
//         const csv = require("csv-parser");
//         const stream = fs.createReadStream(filePath).pipe(csv());
//         for await (const row of stream) {
//           records.push(row);
//         }
//       } else {
//         return res.status(400).json({ message: "Unsupported file type" });
//       }

//       if (!records.length) {
//         return res.status(400).json({ message: "No data found in file" });
//       }
//       // row.openingBalance.replace(/,/g, '')
//       const enrichedRecords = records
//         .filter((record) => record["No. of Months"] !== "Total") // Exclude rows with 'Total'
//         .map((record) => ({
//           payment_date: moment(record["Payment Date"], "DD-MMM-YY").format("YYYY-MM-DD"),
//           due_date: moment(record["Due Date"], "DD-MMM-YY").format("YYYY-MM-DD"),
//           utr_no: parseNumber(record["Utr No"]),
//           due_amt: parseNumber(record["Due Amount"]),
//           payment_amount: parseNumber(record["Payment Amount"]),
//           pricipal_coll: parseNumber(record["Pricipal"]),
//           interest_coll: parseNumber(record["Interest"]),
//           lender_code: parseNumber(record["Lender Code"]),
//           sanction_id: parseNumber(record["Sanction Id"]),
//           tranche_id: parseNumber(record["Tranche Id"]),
//           approval_status: "Approval Pending",
//           // sanction_id,
//           // tranche_id,
//           // lender_code,
//           createdby: decoded.id || "SFTADM",
//         }));

//       // await repayment_schedule_staging.bulkCreate(enrichedRecords);
//       for (const record of enrichedRecords) {
//         const existing = await payment_details_staging.findOne({
//           where: {
//             lender_code: lender_code,
//             sanction_id: sanction_id,
//             tranche_id: tranche_id,
//             due_date: due_date
//           }
//         });

//         if (existing) {
//           await existing.update(record); // Update existing row
//         } else {
//           await payment_details_staging.create(record); // Insert new row
//         }
//       }

//       fs.unlinkSync(filePath); // Cleanup

//       res.json({ message: "UTR Upload Data uploaded to DB successfully" });
//     } catch (error) {
//       console.error("Error processing UTR Upload file:", error);
//       res.status(500).json({ message: "Failed to process file", error: error.message });
//     }
//   });
// };

exports.UTRuploadedFile = (req, res) => {
  upload(req, res, async (err) => {
    if (err) {
      return res.status(400).json({ message: "File upload failed", error: err.message });
    }

    const JWT_SECRET = process.env.JWT_SECRET;
    const { created_by } = req.body;

    if (!created_by) {
      return res.status(401).json({ message: "JWT token not provided" });
    }

    let decoded;
    try {
      decoded = jwt.verify(created_by, JWT_SECRET);
    } catch (error) {
      return res.status(401).json({ message: "Invalid token", error: error.message });
    }

    if (!req.file) {
      return res.status(400).json({ message: "No file provided" });
    }

    const filePath = path.join(__dirname, "../", req.file.path);
    const ext = path.extname(req.file.originalname).toLowerCase();

    try {
      let records = [];

      if (ext === ".xlsx" || ext === ".xls") {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        records = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
      } else if (ext === ".csv") {
        const csv = require("csv-parser");
        const stream = fs.createReadStream(filePath).pipe(csv());
        for await (const row of stream) {
          records.push(row);
        }
      } else {
        return res.status(400).json({ message: "Unsupported file type" });
      }

      if (!records.length) {
        return res.status(400).json({ message: "No data found in file" });
      }

      const enrichedRecords = records
        .filter((record) => record["No. of Months"] !== "Total")
        .map((record) => ({
          payment_date: moment(record["Payment Date"], "DD-MMM-YY").format("YYYY-MM-DD"),
          due_date: moment(record["Due Date"], "DD-MMM-YY").format("YYYY-MM-DD"),
          utr_no: record["Utr No"] ? record["Utr No"].toString().trim() : null,
          due_amt: parseNumber((record["Due Amount"] ? record["Due Amount"].toString().trim() : '0')),
          payment_amount: parseNumber((record["Payment Amount"] ? record["Payment Amount"].toString().trim() : '0')),
          pricipal_coll: parseNumber((record["Pricipal"] ? record["Pricipal"].toString().trim() : '0')),
          interest_coll: parseNumber((record["Interest"] ? record["Interest"].toString().trim() : '0')),
          lender_code: record["Lender Code"] ? record["Lender Code"].toString().trim() : null,
          sanction_id: record["Sanction Id"] ? record["Sanction Id"].toString().trim() : null,
          tranche_id: record["Tranche Id"] ? record["Tranche Id"].toString().trim() : null,
          approval_status: "Approval Pending",
          createdat: new Date(),
          createdby: decoded.id || "SFTADM",
        }));

      for (const record of enrichedRecords) {
        const existing = await payment_details_staging.findOne({
          where: {
            lender_code: record.lender_code,
            sanction_id: record.sanction_id,
            tranche_id: record.tranche_id,
            due_date: record.due_date,
          },
        });

        if (existing) {
          await existing.update(record);
        } else {
          await payment_details_staging.create(record);
        }
      }

      fs.unlinkSync(filePath);

      res.json({ message: "UTR Upload Data uploaded to DB successfully" });
    } catch (error) {
      console.error("Error processing UTR Upload file:", error);
      res.status(500).json({ message: "Failed to process file", error: error.message });
    }
  });
};


exports.uploadedFile = (req, res) => {
  upload(req, res, async (err) => {
    if (err) {
      return res.status(400).json({ message: "File upload failed", error: err.message });
    }

    const { sanction_id, tranche_id, lender_code, due_date, loan_type, created_by } = req.body;
    const JWT_SECRET = process.env.JWT_SECRET;

    if (!created_by) {
      return res.status(401).json({ message: "JWT token not provided" });
    }

    let decoded;
    try {
      decoded = jwt.verify(created_by, JWT_SECRET);
    } catch (error) {
      return res.status(401).json({ message: "Invalid token", error: error.message });
    }

    if (!req.file) {
      return res.status(400).json({ message: "No file provided" });
    }

    const filePath = path.join(__dirname, "../", req.file.path);
    const ext = path.extname(req.file.originalname).toLowerCase();

    try {
      let records = [];

      if (ext === ".xlsx" || ext === ".xls") {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        records = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
      } else if (ext === ".csv") {
        const csv = require("csv-parser");
        const stream = fs.createReadStream(filePath).pipe(csv());
        for await (const row of stream) {
          records.push(row);
        }
      } else {
        return res.status(400).json({ message: "Unsupported file type" });
      }

      if (!records.length) {
        return res.status(400).json({ message: "No data found in file" });
      }
      // row.openingBalance.replace(/,/g, '')
      const enrichedRecords = records
        .filter((record) => record["No. of Months"] !== "Total") // Exclude rows with 'Total'
        .map((record) => ({
          principal_due: parseNumber(record["Principal Repayment"]),
          interest_due: parseNumber(record["Servicing of Interest"]),
          total_due: parseNumber(record["Total Payment"]),
          opening_balance: parseNumber(record["Opening Balance"]),
          closing_balance: parseNumber(record["Closing Balance"]),
          interest_days: record["No. of Days"] || null,
          interest_rate: record["Interest Rate"]
            ? parseFloat(record["Interest Rate"].toString().replace('%', '').trim())
            : null,
          emi_sequence: record["No. of Months"],
          repayment_type: loan_type,
          approval_status: "Approval Pending",
          sanction_id,
          tranche_id,
          lender_code,
          createdby: decoded.id || "SFTADM",
          from_date: moment(record["From Date"], "DD-MMM-YY").format("YYYY-MM-DD"),
          due_date: moment(record["Due Date"], "DD-MMM-YY").format("YYYY-MM-DD"),
        }));

      // await repayment_schedule_staging.bulkCreate(enrichedRecords);
      for (const record of enrichedRecords) {
        const existing = await repayment_schedule_staging.findOne({
          where: {
            lender_code: lender_code,
            sanction_id: sanction_id,
            tranche_id: tranche_id,
            due_date: record.due_date
          }
        });

        if (existing) {
          await existing.update(record); // Update existing row
        } else {
          await repayment_schedule_staging.create(record); // Insert new row
        }
      }

      fs.unlinkSync(filePath); // Cleanup

      res.json({ message: "Data uploaded to DB successfully" });
    } catch (error) {
      console.error("Error processing file:", error);
      res.status(500).json({ message: "Failed to process file", error: error.message });
    }
  });
};
