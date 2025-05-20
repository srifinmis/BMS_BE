// fileProcessor.js
const { sequelize } = require('../config/db');
const fs = require("fs");
const csv = require("csv-parser");
const xlsx = require("xlsx");
const initModels = require('../models/init-models');
// const { repayment_schedule_staging } = require('../models/repayment_schedule_staging');
// uploadCSV: Main function to handle the CSV/Excel upload

const models = initModels(sequelize);
const { repayment_schedule_staging } = models;



const uploadCSV = async (req, res) => {
  try {
    const { originalname, path } = req.file;
    let data;

    // Check file extension and process accordingly
    if (originalname.endsWith(".csv")) {
      data = await parseCSV(path);  // Parse CSV if file is a CSV
    } else if (originalname.endsWith(".xlsx") || originalname.endsWith(".xls")) {
      data = await parseExcel(path);  // Parse Excel if file is an Excel file
    } else {
      return res.status(400).send("Invalid file type. Please upload CSV or Excel.");
    }

    // Insert data into the database
    await insertDataToDatabase(data);

    // Send success response
    res.status(200).send("Data inserted successfully.");
  } catch (error) {
    console.error(error);
    res.status(500).send("Error processing file.");
  }
};

// CSV Parsing function
const parseCSV = (filePath) => {
  return new Promise((resolve, reject) => {
    const data = [];
    fs.createReadStream(filePath)
      .pipe(csv())
      .on("data", (row) => {
        data.push(row); // Push each row of data into the array
      })
      .on("end", () => {
        resolve(data);  // Resolve with the parsed data
      })
      .on("error", (error) => {
        reject(error);  // Reject on error
      });
  });
};

// Excel Parsing function
const parseExcel = (filePath) => {
  return new Promise((resolve, reject) => {
    try {
      const workbook = xlsx.readFile(filePath);
      const sheetNames = workbook.SheetNames;
      const sheet = workbook.Sheets[sheetNames[0]];  // Assume data is in the first sheet
      const data = xlsx.utils.sheet_to_json(sheet);  // Convert Excel sheet to JSON
      resolve(data);
    } catch (error) {
      reject(error);  // Reject on error
    }
  });
};

// database.js
const insertDataToDatabase = async (data) => {
  try {
    // Insert the parsed data into the database
    // await repayment_schedule_staging.bulkCreate(data);
    // console.log("Successfully inserted data into repayment_schedule_staging.");

    // Optionally: fetch and print the first 5 records
    const preview = await repayment_schedule_staging.findAll({ limit: 5 });
    console.log("Preview of repayment_schedule_staging:", preview);
  } catch (error) {
    console.error("Error inserting data into database:", error);
    throw error;  // Rethrow error to be caught in the main flow
  }
};

  

module.exports = { uploadCSV, parseCSV, parseExcel };
