// const { sequelize } = require('../config/db');
// const initModels = require('../models/init-models');
// const { Op } = require("sequelize");
// const jwt = require("jsonwebtoken");

// const models = initModels(sequelize);
// const { sanction_details, bank_account_details_staging, bank_account_details } = models;

// //sending Create Bank Account 
// exports.bankCreate = async (req, res) => {
//     const data = req.body;
//     let temp = data.createdby;
//     const JWT_SECRET = process.env.JWT_SECRET;
//     const decoded = jwt.verify(temp, JWT_SECRET);
//     const bankData = {
//         sanction_id: data.sanctionId,
//         current_ac_no: data.currentAccountNo,
//         bank_name: data.bankName,
//         bank_branch: data.bankBranch,
//         location: data.location,
//         ifsc_code: data.ifscCode,
//         conf_acc_no: data.conf_acc_no,
//         createdat: data.createdat,
//         updatedat: data.updatedat,
//         createdby: decoded.id,
//         updatedby: decoded.id,
//         remarks: data.remarks || null,
//         approval_status: data.approval_status || "Approval Pending"
//     };

//     // console.log('Data from FD:  ', datagot);
//     try {

//         const newRoc = await bank_account_details_staging.create(bankData);
//         res.status(201).json({ message: "Bank Account added successfully", data: newRoc });
//     } catch (error) {
//         console.error("Login Error:", error);
//         res.status(500).json({ message: "Internal server error", error: error.message });
//     }
// };

// // repayment Staging from maintable get approved API all data
// exports.bankFetch = async (req, res) => {
//     const datagot = req.body;
//     try {
//         const roc = await bank_account_details_staging.findAll({
//             attributes: [
//                 "sanction_id", "current_ac_no", "bank_name", "bank_branch", "location", "ifsc_code", "approval_status"
//             ], where: {
//                 approval_status: { [Op.or]: ["Approval Pending", "Rejected"] }
//             }
//         });
//         const rocmain = await bank_account_details.findAll({
//             attributes: [
//                 "sanction_id", "current_ac_no", "bank_name", "bank_branch", "location", "ifsc_code", "approval_status"
//             ], where: { approval_status: "Approved" }
//         });

//         return res.status(201).json({ success: true, data: roc, mainData: rocmain });
//     } catch (error) {
//         console.error("Error fetching lenders:", error);
//         return res.status(500).json({ success: false, message: "Server Error", error: error.message });
//     }
// };



// // bank get by sanction_id Page API
// exports.bankView = async (req, res) => {
//     const { sanction_id, approval_status } = req.query;
//     console.log("backend 2 getall: ", sanction_id, approval_status);
//     try {
//         if (approval_status === 'Approved') {
//             const sanction = await bank_account_details.findOne({
//                 where: { sanction_id, approval_status: 'Approved' }
//             });
//             if (sanction) {
//                 return res.status(200).json({ sanction });
//             } else {
//                 return res.status(404).json({ message: "Approved Bank Repayment not found" });
//             }
//         } else if (approval_status === 'Approval Pending') {
//             const sanction = await bank_account_details_staging.findOne({
//                 where: { sanction_id, approval_status: 'Approval Pending' }
//             });
//             if (sanction) {
//                 return res.status(200).json({ sanction });
//             } else {
//                 return res.status(404).json({ message: "Approval Pending Bank Repayment not found" });
//             }
//         }
//         else if (approval_status === 'Rejected') {
//             const sanction = await bank_account_details_staging.findOne({
//                 where: { sanction_id, approval_status: 'Rejected' }
//             });
//             if (sanction) {
//                 return res.status(200).json({ sanction });
//             } else {
//                 return res.status(404).json({ message: "Rejected Bank Repayment not found" });
//             }
//         } else {
//             return res.status(400).json({ message: "Invalid approval status" });
//         }

//     } catch (error) {
//         console.error("Error fetching Sanction:", error);
//         res.status(500).json({ message: "Internal server error", error: error.message });
//     }
// };

// // ROCUpdate Data Staging and Main tables
// exports.bankUpdate = async (req, res) => {
//     const { sanction_id } = req.body;
//     try {
//         let temp = req.body.createdby;
//         const JWT_SECRET = process.env.JWT_SECRET;
//         const decoded = jwt.verify(temp, JWT_SECRET);
//         // Check if the sanction exists in sanction_details with approval_status: "Approved"
//         const existingSanction = await bank_account_details.findOne({
//             where: {
//                 sanction_id: sanction_id,
//                 approval_status: "Approved"
//             }
//         });

//         if (existingSanction) {
//             // Update lender in lender_master to "Approval Pending"
//             await bank_account_details.update(
//                 { approval_status: "Approved" },
//                 { where: { sanction_id: sanction_id } }
//             );

//             // Check if the lender exists in lender_master_staging with approval_status: "Approval Pending"
//             const existingStagingSanction = await bank_account_details_staging.findOne({
//                 where: {
//                     sanction_id: sanction_id,
//                     approval_status: "Approval Pending"
//                 }
//             });

//             if (existingStagingSanction) {
//                 // If the record already exists with "Approval Pending", do not create a new one
//                 return res.status(400).json({
//                     status: "error",
//                     message: "There is a already a record in progess for edit , no futher updates allowed!"
//                 });
//             }

//             // Create a new record in lender_master_staging with approval status "Approval Pending"
//             const recordWithPendingApproval = {
//                 ...req.body,
//                 approval_status: "Approval Pending",
//                 createdby: decoded.id,  // Add the user who is creating this record
//                 updatedby: decoded.id   // Add the user who is updating this record
//             };

//             const newStagingRecord = await bank_account_details_staging.create(recordWithPendingApproval);

//             return res.status(201).json({
//                 status: "success",
//                 message: "Bank Repayment already exists ,No further updates allowed until approved.",
//                 NewStagingRecord: newStagingRecord
//             });
//         }

//         // If lender does not exist in lender_master or is not "Approved", check if it's in lender_master_staging with "Approval Pending"
//         const existingStagingSanction = await bank_account_details_staging.findOne({
//             where: {
//                 sanction_id: sanction_id,
//                 approval_status: "Approval Pending"
//             }
//         });

//         if (existingStagingSanction) {
//             // If the record exists in lender_master_staging and is pending, do not create another record
//             return res.status(500).json({
//                 status: "error",
//                 message: "Bank Repayment already exists ,No further updates allowed until approved.",
//             });
//         }

//         // If lender is not yet updated in lender_master_staging or is not pending, proceed to update it
//         const updatedStagingSanction = await bank_account_details.update(req.body, {
//             where: { sanction_id: sanction_id },
//             returning: true  // Include returning to get the updated record
//         });

//         return res.status(200).json({
//             status: "success",
//             message: "Bank updated successfully.",
//             UpdatedLender: updatedStagingSanction[1][0]  // Retrieve the updated record from the response
//         });

//     } catch (error) {
//         // console.error("Update Error:", error);
//         res.status(500).json({ status: "error", message: "Internal server error", error: error.message });
//     }
// }




// // bank Details Approval API
// exports.bankApprove = async (req, res) => {
//     try {
//         // console.log("Received Lender Checkbox Data:", req.body);
//         // Ensure req.body is an array for bulk insert
//         if (!Array.isArray(req.body)) {
//             return res.status(400).json({
//                 message: "Invalid data format, expected an array of Bank"
//             });
//         }
//         // 1. Upsert lenders in lender_master (update if exists, create if new)
//         const newSanction = await Promise.all(
//             req.body.map(async (sanction) => {
//                 const existingLender = await bank_account_details.findOne({
//                     where: { sanction_id: sanction.sanction_id }
//                 });

//                 if (existingLender) {
//                     // Lender exists, update the record
//                     await bank_account_details.update(
//                         {
//                             ...sanction,
//                             remarks: sanction.remarks,
//                             approval_status: "Approved",
//                         },
//                         { where: { sanction_id: sanction.sanction_id } }
//                     );

//                     return existingLender;
//                 } else {
//                     // Lender does not exist, insert new record
//                     const newLenderRecord = await bank_account_details.create({
//                         ...sanction,
//                         sanction_id: sanction.sanction_id,
//                         remarks: sanction.remarks,
//                         approval_status: "Approved",

//                     });

//                     return newLenderRecord;
//                 }
//             })
//         );

//         // 2. Update lender_master_staging
//         const updatedLenders = await Promise.all(
//             req.body.map(async (sanction) => {
//                 return await bank_account_details_staging.update(
//                     {
//                         approval_status: "Approved",
//                         remarks: sanction.remarks  // Set remarks for each lender
//                     },
//                     {
//                         where: {
//                             sanction_id: sanction.sanction_id,
//                             approval_status: "Approval Pending"  // Only update those with "Approval Pending" status
//                         }
//                     }
//                 );
//             })
//         );
//         console.log("Update staging: ", updatedLenders)

//         res.status(201).json({
//             message: "Sanction added successfully",
//             datatoMain: newSanction,
//             StagingUpdate: updatedLenders
//         });
//     } catch (error) {
//         console.error("Error:", error);
//         res.status(500).json({
//             message: "Internal server error",
//             error: error.message
//         });
//     }
// };



// // ROC Form reject API
// exports.bankReject = async (req, res) => {
//     try {
//         // console.log("Received Lender Data:", req.body);

//         // Ensure req.body is an array for bulk insert
//         if (!Array.isArray(req.body)) {
//             return res.status(400).json({
//                 message: "Invalid data format, expected an array of bank"
//             });
//         }

//         const updatePromises = req.body.map(sanction => {
//             return bank_account_details_staging.update(
//                 { approval_status: "Rejected", remarks: sanction.remarks },
//                 { where: { sanction_id: sanction.sanction_id, approval_status: "Approval Pending" } }
//             );
//         });

//         // Wait for all updates to finish
//         await Promise.all(updatePromises);

//         res.status(201).json({
//             message: "Bank Repayment Rejected successfully"
//         });
//     } catch (error) {
//         console.error("Error:", error);
//         res.status(500).json({
//             message: "Internal server error",
//             error: error.message
//         });
//     }
// }


// //bank repayment apprivalpending api
// exports.bankPending = async (req, res) => {
//     try {
//         // const { lender_code } = req.params;
//         const sanction = await bank_account_details_staging.findAll({
//             where: { approval_status: "Approval Pending" }
//         });

//         if (!sanction || sanction.length === 0) {
//             return res.status(404).json({ message: "Sanction Details not found" });
//         }

//         res.status(201).json({ success: true, data: sanction });
//     } catch (error) {
//         console.error("Error fetching lender:", error);
//         res.status(500).json({ message: "Internal server error", error: error.message });
//     }
// }