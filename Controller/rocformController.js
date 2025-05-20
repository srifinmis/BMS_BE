const { sequelize } = require('../config/db');
const initModels = require('../models/init-models');
const { Op } = require("sequelize");
const jwt = require("jsonwebtoken");

const models = initModels(sequelize);
const { sanction_details, roc_forms_staging, roc_forms, tranche_details } = models;

//Sanction ID's Fetch  API
exports.sanctionId = async (req, res) => {
    const datagot = req.body;
    // console.log('Data from sanction id FD:  ', datagot);
    try {

        const sanctionget = await sanction_details.findAll({
            attributes: ["sanction_id", "lender_code", "sanction_amount", "sanction_date"],
        });
        // console.log("roc: ", sanctionget)


        // const Lenderget = await lender_master
        res.status(201).json({ message: "Roc Fetch successfully", data: sanctionget });
    } catch (error) {
        console.error("Login Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};

//sending Create ROC_Form 
exports.rocCreate = async (req, res) => {
    const data = req.body;
    let temp = data.createdby;
    const JWT_SECRET = process.env.JWT_SECRET;
    const decoded = jwt.verify(temp, JWT_SECRET);
    // console.log('Data from FD:  ', datagot);
    const rocData = {
        sanction_id: data.sanction_id,
        approved_by: data.approved_by,
        date_of_approval: data.date_of_approval,
        document_executed_date: data.document_executed_date,
        due_date_charge_creation: data.due_date_charge_creation,
        date_of_form_filed_creation: data.date_of_form_filed_creation || null,
        due_date_satisfaction: data.due_date_satisfaction || null,
        date_of_filing_satisfaction: data.date_of_filing_satisfaction || null,
        // createdat: new Date(),
        // updatedat: new Date(),
        createdby: decoded.id,
        updatedby: decoded.id,
        remarks: data.remarks || null,
        lender_code: data.lender_code,
        user_type: "N",
        approval_status: data.approval_status || "Approval Pending"

    }
    // console.log('Data from FD:  ', datagot);
    try {

        const newRoc = await roc_forms_staging.create(rocData);
        res.status(201).json({ message: "Roc added successfully", data: newRoc });
    } catch (error) {
        console.error("Login Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });

        // res.status(500).json({ message: "Server Error", error: error.message });
    }
};


// rocform  Staging from maintable get approved API all data
exports.rocFetch = async (req, res) => {
    const datagot = req.body;
    try {
        const roc = await roc_forms_staging.findAll({
            attributes: [
                "lender_code", "sanction_id", "approved_by", "date_of_approval", "document_executed_date", "due_date_charge_creation", "approval_status", "createdat", "updatedat"
            ], where: {
                approval_status: { [Op.or]: ["Approval Pending", "Rejected"] }
            }
        });
        const rocmain = await roc_forms.findAll({
            attributes: [
                "lender_code", "sanction_id", "approved_by", "date_of_approval", "document_executed_date", "due_date_charge_creation", "approval_status", "createdat", "updatedat"
            ], where: { approval_status: "Approved" }
        });

        return res.status(201).json({ success: true, data: roc, mainData: rocmain });
    } catch (error) {
        console.error("Error fetching lenders:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
};
exports.rocValidate = async (req, res) => {
    try {
        const { sanction_id, lender_code } = req.query;
        console.log("backend check rocforms:", sanction_id, lender_code)
        const sanction = await roc_forms_staging.findAll({
            where: { sanction_id, lender_code }
        });

        if (!sanction || sanction.length === 0) {
            return res.status(404).json({ message: "ROC FORMS not found" });
        }

        res.status(201).json({ success: true, data: sanction });
    } catch (error) {
        console.error("Error fetching ROC FORMS:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}


// roc get by sanction_id Page API
exports.rocView = async (req, res) => {
    const { sanction_id, lender_code, approval_status, updatedat } = req.query;
    // const originalDate = new Date(updatedat);
    // const updatedDate = new Date(originalDate.getTime() + (5 * 60 + 30) * 60 * 1000);
    // console.log("backend roc 4 getall: ", lender_code, sanction_id, approval_status, updatedDate);
    try {
        if (approval_status === 'Approved') {
            const roc = await roc_forms.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, approval_status: approval_status
                    // , updatedat: updatedDate 
                }
            });
            if (roc) {
                return res.status(200).json({ roc });
            } else {
                return res.status(404).json({ message: "Approved roc not found" });
            }
        } else if (approval_status === 'Approval Pending') {
            const roc = await roc_forms_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, approval_status: approval_status
                    // , updatedat: updatedDate
                }
            });
            if (roc) {
                return res.status(200).json({ roc });
            } else {
                return res.status(404).json({ message: "Approval Pending roc not found" });
            }
        }
        else if (approval_status === 'Rejected') {
            const roc = await roc_forms_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, approval_status: approval_status
                    // , updatedat: updatedDate 
                }
            });
            if (roc) {
                return res.status(200).json({ roc });
            } else {
                return res.status(404).json({ message: "Rejected roc not found" });
            }
        } else {
            return res.status(400).json({ message: "Invalid approval status" });
        }

    } catch (error) {
        console.error("Error fetching roc:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};


// ROCUpdate Data Staging and Main tables
exports.rocUpdate = async (req, res) => {
    const { sanction_id, lender_code, user_type } = req.body;
    const data = req.body;
    data.id = null;
    console.log("data: ", data);
    const newData = data;
    console.log("new object data: ", newData);

    try {
        const JWT_SECRET = process.env.JWT_SECRET;
        const decoded = jwt.verify(data.createdby, JWT_SECRET);

        // 🔐 Global check for any pending approval record in staging
        const existingStagingLender = await roc_forms_staging.findOne({
            where: {
                lender_code,
                sanction_id,
                approval_status: "Approval Pending"
            }
        });

        if (existingStagingLender) {
            return res.status(400).json({
                status: "error",
                message: "There is already a record in progress. No further updates allowed until approved or rejected."
            });
        }

        // Check for Approved record in lender_master
        const existingLender = await roc_forms.findOne({
            where: {
                lender_code,
                sanction_id,
                approval_status: "Approved"
            }
        });

        // Check for Rejected records in staging
        const rejectedStagingLenders = await roc_forms_staging.findAll({
            where: {
                lender_code,
                sanction_id,
                approval_status: "Rejected"
            }
        });

        // Case 1: user_type === "N" (New record)
        if (user_type === "N") {
            const existsInMaster = await roc_forms.findOne({
                where: {
                    lender_code,
                    sanction_id,
                }
            });

            if (existsInMaster) {
                return res.status(400).json({
                    status: "error",
                    message: "This lender_code already exists in master. Cannot create new record."
                });
            }

            let updatedFields = [];
            if (rejectedStagingLenders.length > 0) {
                const lastRejected = rejectedStagingLenders[rejectedStagingLenders.length - 1];
                Object.keys(data).forEach((key) => {
                    if (data[key] !== lastRejected[key]) {
                        updatedFields.push(key);
                    }
                });
            }

            const newRecord = {
                ...data,
                createdat: new Date(),
                updatedat: new Date(),
                approval_status: "Approval Pending",
                createdby: decoded.id,
                updatedby: decoded.id,
                updated_fields: updatedFields,
                id: null
            };

            const newStagingRecord = await roc_forms_staging.create(newRecord);

            return res.status(201).json({
                status: "success",
                message: "New ROC created .",
                NewStagingRecord: newStagingRecord,
                updatedFields
            });
        }

        // Case 2: user_type === "U" (Create update request directly)
        if (user_type === "U") {
            const newRecord = {
                ...data,
                createdat: new Date(),
                updatedat: new Date(),
                approval_status: "Approval Pending",
                createdby: decoded.id,
                updatedby: decoded.id,
                id: null
            };

            const newStagingRecord = await roc_forms_staging.create(newRecord);

            return res.status(201).json({
                status: "success",
                message: "Record inserted.",
                NewStagingRecord: newStagingRecord
            });
        }

        // Case 3: If record exists in master, compare and stage updates
        let updatedFields = [];
        if (existingLender) {
            Object.keys(newData).forEach((key) => {
                if (newData[key] !== existingLender[key]) {
                    updatedFields.push(key);
                }
            });

            const recordWithPendingApproval = {
                ...data,
                createdat: new Date(),
                updatedat: new Date(),
                updated_fields: updatedFields,
                approval_status: "Approval Pending",
                id: null,
                createdby: decoded.id,
                updatedby: decoded.id
            };

            const newStagingRecord = await roc_forms_staging.create(recordWithPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "ROC update request is in progress. No further updates allowed until approved.",
                NewStagingRecord: newStagingRecord,
                updatedFields
            });
        }

        // Case 4: If no approved record but previously rejected exists
        if (!existingLender && rejectedStagingLenders.length > 0) {
            const lastRejected = rejectedStagingLenders[rejectedStagingLenders.length - 1];

            updatedFields = [];
            Object.keys(newData).forEach((key) => {
                if (newData[key] !== lastRejected[key]) {
                    updatedFields.push(key);
                }
            });

            const recordPendingApproval = {
                ...data,
                createdat: new Date(),
                updatedat: new Date(),
                approval_status: "Approval Pending",
                createdby: decoded.id,
                updatedby: decoded.id,
                user_type: "U",
                updated_fields: updatedFields,
                id: null
            };

            const newStagingRecord = await roc_forms_staging.create(recordPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "New record created for previously rejected lender_code.",
                NewStagingRecord: newStagingRecord,
                updatedFields: updatedFields
            });
        }

        // Final fallback: attempt to update existing staging record
        const [updateCount, updatedRecords] = await roc_forms_staging.update(data, {
            where: {
                lender_code,
                sanction_id,
            },
            returning: true
        });

        if (updateCount === 0) {
            return res.status(404).json({
                status: "error",
                message: "Sanction not found or no changes detected."
            });
        }

        return res.status(200).json({
            status: "success",
            message: "ROC updated successfully.",
            updatedFields: updatedFields,
            UpdatedLender: updatedRecords ? updatedRecords[0] : null
        });

    } catch (error) {
        console.error("Update Error:", error);
        res.status(500).json({ status: "error", message: "Internal server error", error: error.message });
    }
};


// Sanction Details Approval API
exports.rocApprove = async (req, res) => {
    try {
        console.log("approve Roc Form backend:", req.body)
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Roc Form"
            });
        }
        // 1. Upsert lenders in lender_master (update if exists, create if new)
        const newSanction = await Promise.all(
            req.body.map(async (sanction) => {
                const existingSanction = await roc_forms.findOne({
                    where: { sanction_id: sanction.sanction_id, lender_code: sanction.lender_code }
                });

                if (existingSanction) {
                    // Lender exists, update the record
                    await roc_forms.update(
                        {
                            ...sanction,
                            // createdat: new Date(),
                            updatedat: new Date(),
                            remarks: sanction.remarks,
                            approval_status: "Approved",
                        },
                        { where: { sanction_id: sanction.sanction_id, lender_code: sanction.lender_code } }
                    );

                    return existingSanction;
                } else {
                    // Lender does not exist, insert new record
                    const newLenderRecord = await roc_forms.create({
                        ...sanction,
                        sanction_id: sanction.sanction_id,
                        createdat: new Date(),
                        updatedat: new Date(),
                        remarks: sanction.remarks,
                        approval_status: "Approved",

                    });

                    return newLenderRecord;
                }
            })
        );

        // 2. Update lender_master_staging
        const updatedLenders = await Promise.all(
            req.body.map(async (sanction) => {
                return await roc_forms_staging.update(
                    {
                        approval_status: "Approved",
                        user_type: "U",
                        remarks: sanction.remarks  // Set remarks for each lender
                    },
                    {
                        where: {
                            sanction_id: sanction.sanction_id,
                            lender_code: sanction.lender_code,
                            id: sanction.id,
                            approval_status: "Approval Pending"  // Only update those with "Approval Pending" status
                        }
                    }
                );
            })
        );
        console.log("Update staging: ", updatedLenders)

        res.status(201).json({
            message: "Roc Form added successfully",
            datatoMain: newSanction,
            StagingUpdate: updatedLenders
        });
    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({
            message: "Internal server error",
            error: error.message
        });
    }
};



// ROC Form reject API
exports.rocReject = async (req, res) => {
    try {
        console.log("Received Roc Form rejection Data:", req.body);

        // Ensure req.body is an array for bulk insert
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of ROC Form"
            });
        }

        const updatePromises = req.body.map(sanction => {
            return roc_forms_staging.update(
                {
                    approval_status: "Rejected", user_type: sanction.user_type, remarks: sanction.remarks
                },
                {
                    where: {
                        sanction_id: sanction.sanction_id, id: sanction.id, lender_code: sanction.lender_code, approval_status: "Approval Pending"
                    }
                }
            );
        });

        // Wait for all updates to finish
        await Promise.all(updatePromises);

        res.status(201).json({
            message: "Roc Form Rejected successfully"
        });
    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({
            message: "Internal server error",
            error: error.message
        });
    }
}

//ROC apprivalpending api
exports.rocPending = async (req, res) => {
    try {
        // const { lender_code } = req.params;
        const roc = await roc_forms_staging.findAll({
            where: { approval_status: "Approval Pending" }
        });

        if (!roc || roc.length === 0) {
            return res.status(404).json({ message: "Sanction Details not found" });
        }

        res.status(201).json({ success: true, data: roc });
    } catch (error) {
        console.error("Error fetching roc:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}