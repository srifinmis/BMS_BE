const express = require("express");
const router = express.Router();
// const multer = require('multer');
// const storage = multer.memoryStorage();
// const upload = multer({ storage: storage });
const { login } = require("../Controller/loginController");
const ForgotRoutes = require("../Controller/ForgotPassword/forgotpassword");
const ResetRoutes = require("../Controller/ForgotPassword/Reset_Pass");

const { lenderCreate,
    getLenders,
    getLendesAll,
    lenderApprove,
    lenderreject,
    lenderPending,
    lenderupdate,
    lenderCheck } = require("../Controller/LenderApprovalController");
const { getLoadData } = require("../Controller/borrowingLoadController");
const { cronalert } = require("../Controller/alertTriggerController");
const { generateCronExpression } = require("../Controller/cronExpressionController");
const { alertData } = require("../Controller/alertApiController");

// Sanction
const { sanctionApprove, sanctionCheck } = require("../Controller/sanctionController");
const { lenderCodes } = require("../Controller/sanctionController");
const { sanctionUpdate } = require("../Controller/sanctionController");
const { sanctionView } = require("../Controller/sanctionController");
const { sanctionReject } = require("../Controller/sanctionController");
const { sanctionFetch } = require("../Controller/sanctionController");
const { sanctionPending } = require("../Controller/sanctionController");
const { sanctionCreate } = require("../Controller/sanctionController");

// ROC Form
const { sanctionId, rocValidate } = require("../Controller/rocformController")
const { rocCreate } = require("../Controller/rocformController")
const { rocFetch } = require("../Controller/rocformController")
const { rocView } = require("../Controller/rocformController")
const { rocUpdate } = require("../Controller/rocformController")
const { rocApprove } = require("../Controller/rocformController")
const { rocReject } = require("../Controller/rocformController")
const { rocPending } = require("../Controller/rocformController")

// Bank Account Repayment 
// const { bankCreate,
//     bankFetch,
//     bankView,
//     bankUpdate,
//     bankApprove,
//     bankReject,
//     bankPending } = require("../Controller/repaymentController")


// Executed Documents 
const { executedCreate,
    executedFetch,
    executedView,
    executedUpdate,
    executedApprove,
    executedReject,
    executedPending,
    viewdocument,
    executedValidate } = require("../Controller/executedDocxController");

// Tranche
const { trancheCreate,
    trancheFetch,
    trancheView,
    trancheUpdate,
    trancheApprove,
    trancheReject,
    tranchePending,
    trancheTwo,
    trancheCheck } = require("../Controller/trancheController");

//upload
const { uploadFileToS3 } = require("../s3_Bucket/S3service")
const { uploadFileToLocal } = require("../s3_Bucket/localstorage")

//repayment
const { uploadRepaymentSchedule,
    RSfetchAll,
    scheduleApprove,
    scheduleReject,
    repaymentLenders } = require("../Controller/repaymentscheduleuploadcontroller");

const { intrestCreate,
    intrestFetch,
    intrestView,
    intrestUpdate,
    intrestPending,
    intrestApprove,
    intrestReject,
    interestThree } = require("../Controller/IntrestRateController");


const { generateDrawdownReport } = require("../Controller/Reports/Drawdown")
const { generateDrawdownReportconsolidated } = require("../Controller/Reports/drawdownconsolidated")
//reports
const { generateLenderMasterReport } = require('../Controller/Reports/Reportscontroller');
const { generateSanctionDetailsReport } = require('../Controller/Reports/sancctiondetailsreportcontroller');

const { fetchEmployee, roleChange, fetchRoles } = require("../Controller/RoleBase/RoleChangeController");
const { generateRocChargeCreationReport, getroclenders, generateRocSatisfactionChargeReport } = require("../Controller/Reports/RocChargeCreationController");
const { generateDatewiseRepaymentReport, generateDailyRepaymentReport } = require("../Controller/Reports/DateWiseRepaymentReportController");
const { generateLoanTrancheDetailsReport } = require("../Controller/Reports/TrancheLoanDetailsReportController");
const { generateFundingMixReport } = require("../Controller/Reports/FundingMixReportController");
const { generateEffectiveInterestRateReport } = require("../Controller/Reports/EffectiveInterestRateController");
const { generateRundownReport } = require("../Controller/Reports/RundownController");
const { generateEffectiveInterestofReport } = require("../Controller/Reports/EffectiveInterestController");
const { UTRuploadFileToLocal } = require("../s3_Bucket/utrstorage");
const { UTRThree, UTRCreate, UTRFetch, UTRPending, UTRView, UTRUpdate, UTRApprove, UTRReject } = require("../Controller/UTRController");
const { generateRepaymentScheduleReport } = require("../Controller/Reports/repaymanschedulecontroller");
const { appendXML } = require("pdfkit");
const { almreport } = require("../Controller/Reports/ALMController");
const { getMasterReport } = require("../Controller/Reports/MasterController");


router.post("/Login", login);
router.use("/forgot", ForgotRoutes);
router.use("/reset", ResetRoutes);

// Role Change 
router.get("/employee/idname", fetchEmployee)
router.get("/role/fetchRole", fetchRoles)
router.patch("/role/update", roleChange)

// Post Main Table API
router.post("/Lender/create", lenderCreate);
router.get("/lender/lenderCodecheck", lenderCheck)
router.get("/lender/list", getLenders);
router.post("/lender/Approve", lenderApprove);
router.post("/lender/reject", lenderreject);
router.get("/lender/details", getLendesAll);
router.patch("/lender/update/:lender_code", lenderupdate);
router.get("/lender/pendingData", lenderPending);
router.get("/Load/AllData", getLoadData);
// router.get("/lender/fetchlendercodes", getLendercodes)

// Sanctiondetails API's
router.get("/sanction/lendercodes", lenderCodes);
router.get("/sanction/sanctionidcheck", sanctionCheck)
router.post("/sanction/create", sanctionCreate);
router.get("/sanction/fetchAll", sanctionFetch);
router.get("/sanction/details", sanctionView);
router.patch("/sanction/update/:sanction_id", sanctionUpdate);
router.post("/sanction/Approve", sanctionApprove);
router.post("/sanction/reject", sanctionReject)
router.get("/sanction/pendingData", sanctionPending);

// ROC Form
// use all sanctionid requests also gets lender codes&lender name  
router.get("/roc/sanctionid", sanctionId);

router.post("/roc/create", rocCreate);
router.get("/roc/fetchAll", rocFetch);
router.get("/roc/details", rocView);
router.patch("/roc/update/:sanction_id", rocUpdate);
router.post("/roc/Approve", rocApprove);
router.post("/roc/reject", rocReject);
router.get("/roc/pendingData", rocPending);
router.get("/roc/validate", rocValidate)

// Bank Account Repayment
// router.get("/bankrepayment/sanctionid", sanctionId);
// router.post("/bankrepayment/create", bankCreate);
// router.get("/bankrepayment/fetchAll", bankFetch);
// router.get("/bankrepayment/details", bankView);
// router.patch("/bankrepayment/update/:sanction_id", bankUpdate);
// router.post("/bankrepayment/Approve", bankApprove);
// router.post("/bankrepayment/reject", bankReject);
// router.get("/bankrepayment/pendingData", bankPending);

// Executed Documents API's
router.post("/executed/create", executedCreate);
router.get("/executed/fetchAll", executedFetch);
router.get("/executed/details/:sanction_id", executedView);
router.patch("/executed/update/:sanction_id", executedUpdate);
router.post("/executed/Approve", executedApprove);
router.post("/executed/reject", executedReject);
router.get("/executed/pendingData", executedPending);
router.get("/executed/document/:sanction_id", viewdocument);
router.get("/executed/validate", executedValidate);

// Tranche
router.post("/tranche/create", trancheCreate)
router.get("/tranche/fetchAll", trancheFetch);
router.get("/tranche/details", trancheView);
router.patch("/tranche/update/:tranche_id", trancheUpdate);
router.post("/tranche/Approve", trancheApprove);
router.post("/tranche/reject", trancheReject);
router.get("/tranche/pendingData", tranchePending);
// Tranche id,sanction_id api,lender_code
router.get("/tranche/checkValidity", trancheCheck);
router.get("/tranche/findTwo", trancheTwo);

// Intrest Rate
router.post("/interest/create", intrestCreate)
router.get("/interest/fetchAll", intrestFetch)
router.get("/interest/details", intrestView);
router.patch("/interest/update/:sanction_id", intrestUpdate);
router.post("/interest/Approve", intrestApprove);
router.post("/interest/reject", intrestReject);
router.get("/interest/pendingData", intrestPending);
router.get("/interest/findthree", interestThree);

// Upload Repayment 
// router.post("/uploadrepayment/create",uploadRepayment);

// cron API's
router.post("/crons/alert", cronalert);
router.post("/cron/create", generateCronExpression)

//repayment 
router.get("/alert/findall", alertData);

//upload
router.post("/upload-s3", uploadFileToS3);
router.post("/upload-local", uploadFileToLocal);

// UTR uploads
router.post("/UTR/create", UTRCreate)
router.post("/UTR-upload-local", UTRuploadFileToLocal);

// UTR 
router.get("/UTR/findthree", UTRThree);
router.get("/UTR/fetchAll", UTRFetch)
router.get("/UTR/details", UTRView);
router.patch("/UTR/update/:sanction_id", UTRUpdate);
router.post("/UTR/Approve", UTRApprove);
router.post("/UTR/reject", UTRReject);
router.get("/UTR/pendingData", UTRPending);



// Repayment Schedule
router.post("/upload-repayment-schedule", uploadRepaymentSchedule);
router.get("/repayment/schedule/fetchAll", RSfetchAll);
router.post("/schedule/Approve", scheduleApprove);
router.post("/schedule/Reject", scheduleReject);
router.get("/repaymentschedule/lenders", repaymentLenders)


//reports
router.post('/generate-lender-master', generateLenderMasterReport);
router.post('/generate-sanction-master', generateSanctionDetailsReport);
router.post("/generate-Rocchangecreation", generateRocChargeCreationReport);
router.post("/generate-Rocsatisfactionchange", generateRocSatisfactionChargeReport);
router.post("/generate-DatewiseRepaymentReport", generateDatewiseRepaymentReport);
router.post("/generate-DailyRepaymentReport", generateDailyRepaymentReport);
router.post("/generate-LoanTrancheDetailsReport", generateLoanTrancheDetailsReport);
router.post("/generate-FundingMixReport", generateFundingMixReport);
router.post("/generate-EffectiveInterestRateReport", generateEffectiveInterestRateReport);
router.post("/generate-RundownReport", generateRundownReport);
router.post("/generate-EffectiveInterestofReport", generateEffectiveInterestofReport);
router.post("/generate-RepaymentScheduleReport", generateRepaymentScheduleReport);
router.post("/generate-ALMReport", almreport);
router.post("/generate-MasterReport", getMasterReport);
// router.post("/drawdowns", generateDrawdownReport);
// router.post("/drawdowns-consolidated", generateDrawdownReportconsolidated);

// report get Lendercodes apis
router.get("/generate/roc/lendercodes", getroclenders)

module.exports = router;