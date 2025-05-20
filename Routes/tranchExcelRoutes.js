const express = require("express");
const router = express.Router();
const multer = require("multer");

const { generateScheduleFile } = require('../Controller/tranchExcelController');
const { generateNonEMILoanScheduleFile } = require('../Controller/generateNonEMILoanSchedule');
const { generateEMILoanScheduleFile } = require('../Controller/quarterly_1');
const { generateLoanReport } = require('../Controller/quarterlyPrincipalMonthlyInterest');
const { exportBulletLoanSchedule } = require('../Controller/bulletLoanSchedule');
const { uploadCSV } = require('../Controller/fileProcessor');
const upload = multer({ dest: "../assets" }); // Set the upload directory
const { uploadScheduleFile } = require('../Controller/upload_schedule')
const { uploadedFile, UTRuploadedFile } = require('../Controller/uploaded_file');
const { generateBrokenEMI } = require('../Controller/broken_emiController');
const { generatequarterly_3 } = require('../Controller/quarterly_3');
const { tranche_sum } = require('../Controller/tranche_sum');
const { getPaymentDetails, getTenureMonths } = require('../Controller/paid_schedules');

router.post('/generate-emi-schedule', generateScheduleFile);
router.post('/generate-nonemi-schedule', generateNonEMILoanScheduleFile);
router.post('/generate-quarterly-report', generateEMILoanScheduleFile);
router.post('/generate-princ_monthly-report', generateLoanReport);
router.post('/loan-schedule/bullet', exportBulletLoanSchedule);
router.post('/generate-broken-emi', generateBrokenEMI);
router.post("/upload-csv-excel", upload.single("file"), uploadCSV);
router.post("/upload_schedule", uploadScheduleFile);

router.post("/upload/file", uploadedFile);
router.post("/UTR/upload/file", UTRuploadedFile);

router.post("/quarterly_3", generatequarterly_3);
router.post("/tranche_sum", tranche_sum);
router.post("/get_tranche_payments", getPaymentDetails);
router.post("/getTenureMonths", getTenureMonths);


// do tenure_months = (/getTenureMonths) - (/get_tranche_payments.count)

module.exports = router;
