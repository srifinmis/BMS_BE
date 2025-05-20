const { sequelize } = require('../config/db');
const initModels = require('../models/init-models');
const { Op } = require("sequelize");
const jwt = require("jsonwebtoken");
const path = require("path");
const fs = require("fs");

const models = initModels(sequelize);
const { sanction_details, tranche_details, executed_documents_staging, executed_documents, payment_details, payment_details_staging } = models;

//sending Create Executed Documents  
exports.UTRCreate = async (req, res) => {

    const data = req.body;
    let temp = data.createdby;
    const JWT_SECRET = process.env.JWT_SECRET;
    const decoded = jwt.verify(temp, JWT_SECRET);

    const UTRData = {
        lender_code: data.lender_code,
        sanction_id: data.sanction_id,
        tranche_id: data.tranche_id,
        payment_date: data.payment_date,
        utr_no: data.utr_no,
        due_date: data.due_date,
        payment_amount: data.payment_amount,
        pricipal_coll: data.pricipal_coll,
        interest_coll: data.interest_coll,
        due_amt: data.due_amt,
        // document_type: data.document_type,
        // file_name: data.fileName,
        // uploaded_date: data.uploadedDate,
        // document_url: data.fileUrl || null, // Convert "" to null
        createdat: new Date(),
        approval_status: data.approval_status || "Approval Pending",
        user_type: "N",
        createdby: decoded.id
    };
    console.log('Data from FD:  ', data);
    try {

        const newutr = await payment_details_staging.create(UTRData);
        res.status(201).json({ message: "UTR Uploaded successfully", data: newutr });
    } catch (error) {
        console.error("upload Failed Error:", error);
        res.status(500).json({ message: "Upload Failed error", error: error.message });
    }
};

exports.UTRThree = async (req, res) => {
    try {
        const tranche = await payment_details.findAll({
            attributes: [
                'lender_code', 'sanction_id', 'tranche_id'
            ]
        });

        return res.status(201).json({ success: true, data: tranche });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }

}

exports.UTRFetch = async (req, res) => {
    const datagot = req.body;
    try {
        const tranche = await payment_details_staging.findAll({
            attributes: [
                "lender_code", "sanction_id", "tranche_id", "payment_date", "utr_no", "approval_status", "createdat"
            ], where: {
                approval_status: { [Op.or]: ["Approval Pending", "Rejected"] }
            }
        });
        const tranchemain = await payment_details.findAll({
            attributes: [
                "lender_code", "sanction_id", "tranche_id", "payment_date", "utr_no", "approval_status", "createdat"
            ], where: { approval_status: "Approved" }
        });

        return res.status(201).json({ success: true, data: tranche, mainData: tranchemain });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
}

exports.UTRView = async (req, res) => {
    const { lender_code, sanction_id, tranche_id, approval_status, createdat } = req.query;
    try {
        if (approval_status === 'Approved') {
            const interest = await payment_details.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , createdat: createDate 
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "Approved UTR Upload not found" });
            }
        } else if (approval_status === 'Approval Pending') {
            const interest = await payment_details_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , createdat: createDate
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "Approval Pending UTR Upload not found" });
            }
        }
        else if (approval_status === 'Rejected') {
            const interest = await payment_details_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , createdat: createDate 
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "UTR Upload not found" });
            }
        } else {
            return res.status(400).json({ message: "Invalid approval status" });
        }


    } catch (error) {
        console.error("Error fetching UTR Upload:", error);
        res.status(500).json({ message: "Internal server error UTR Upload", error: error.message });
    }
}


exports.UTRUpdate = async (req, res) => {
    const { sanction_id, tranche_id, due_date, lender_code, user_type, approval_status, updatedat } = req.body;
    const data = req.body;
    data.payment_id = null;
    console.log("utr data: ", data);
    const newData = data;
    console.log("new utr data: ", newData);
    try {
        const JWT_SECRET = process.env.JWT_SECRET;
        // const decoded = jwt.verify(data.updatedFormData.createdby, JWT_SECRET);
        const decoded = jwt.verify(data.createdby, JWT_SECRET);
        // 🔐 Global check for any pending approval record in staging
        const existingStagingLender = await payment_details_staging.findOne({
            where: {
                lender_code,
                sanction_id,
                tranche_id,
                due_date,
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
        const existingLender = await payment_details.findOne({
            where: {
                lender_code,
                sanction_id,
                tranche_id,
                due_date,
                approval_status: "Approved"
            }
        });

        // Check for Rejected records in staging
        const rejectedStagingLenders = await payment_details_staging.findAll({
            where: {
                lender_code,
                sanction_id,
                tranche_id,
                due_date,
                approval_status: "Rejected"
            }
        });

        // Case 1: user_type === "N" (New record)
        if (user_type === "N") {
            const existsInMaster = await payment_details.findOne({
                where: {
                    lender_code,
                    sanction_id,
                    tranche_id,
                    due_date
                }
            });

            if (existsInMaster) {
                return res.status(400).json({
                    status: "error",
                    message: "This lenderCode,sanctionID,TrancheID already exists in master. Cannot create new record."
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
                // updatedat: new Date(),
                approval_status: "Approval Pending",
                createdby: decoded.id,
                // updatedby: data.updatedby,
                updated_fields: updatedFields,
                payment_id: null
            };
            console.log("new update: ", newRecord)
            const newStagingRecord = await payment_details_staging.create(newRecord);

            return res.status(201).json({
                status: "success",
                message: "New UTR created .",
                NewStagingRecord: newStagingRecord,
                updatedFields
            });
        }

        // Case 2: user_type === "U" (Create update request directly)
        if (user_type === "U") {
            const newRecord = {
                ...data,
                createdat: new Date(),
                // updatedat: new Date(),
                approval_status: "Approval Pending",
                createdby: decoded.id,
                // updatedby: data.updatedby,
                payment_id: null
            };

            const newStagingRecord = await payment_details_staging.create(newRecord);

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
                // updatedat: new Date(),
                updated_fields: updatedFields,
                approval_status: "Approval Pending",
                payment_id: null,
                createdby: decoded.id,
                // updatedby: data.updatedby
            };

            const newStagingRecord = await payment_details_staging.create(recordWithPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "UTR update request is in progress. No further updates allowed until approved.",
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
                // updatedat: new Date(),
                approval_status: "Approval Pending",
                createdby: decoded.id,
                // updatedby: data.updatedby,
                user_type: "U",
                updated_fields: updatedFields,
                payment_id: null
            };

            const newStagingRecord = await payment_details_staging.create(recordPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "New record created for previously rejected lenderCode,sanctionID,TrancheID.",
                NewStagingRecord: newStagingRecord,
                updatedFields: updatedFields
            });
        }

        // Final fallback: attempt to update existing staging record
        const [updateCount, updatedRecords] = await payment_details_staging.update(data, {
            where: {
                lender_code,
                sanction_id,
                tranche_id,
                due_date
            },
            returning: true
        });

        if (updateCount === 0) {
            return res.status(404).json({
                status: "error",
                message: "UTR not found or no changes detected."
            });
        }

        return res.status(200).json({
            status: "success",
            message: "UTR Upload updated successfully.",
            updatedFields: updatedFields,
            UpdatedLender: updatedRecords ? updatedRecords[0] : null
        });

    } catch (error) {
        console.error("Update Error:", error);
        res.status(500).json({ status: "error", message: "Internal server error", error: error.message });
    }
};

exports.UTRApprove = async (req, res) => {
    console.log("Approve UTR Schedule:", req.body);

    try {
        const schedules = req.body;

        if (!Array.isArray(schedules) || schedules.length === 0) {
            return res.status(400).json({
                message: "Invalid or empty request body. Expected a non-empty array of UTR uploads."
            });
        }

        const { sanction_id, tranche_id } = schedules[0];

        if (!sanction_id || !tranche_id) {
            return res.status(400).json({
                message: "Missing 'sanction_id' or 'tranche_id' in UTR data."
            });
        }

        // Check if sanction and tranche exist (optional validation)
        const [sanctionExists, trancheExists] = await Promise.all([
            sanction_details.findOne({ where: { sanction_id } }),
            tranche_details.findOne({ where: { sanction_id, tranche_id } })
        ]);

        if (!sanctionExists) {
            return res.status(404).json({ message: `Sanction not found for Sanction ID: ${sanction_id}` });
        }
        if (!trancheExists) {
            return res.status(404).json({ message: `Tranche not found for Sanction ID: ${sanction_id}, Tranche ID: ${tranche_id}` });
        }

        // Process each schedule item
        const approvalResults = await Promise.all(
            schedules.map(async (schedule) => {
                const { lender_code, due_date } = schedule;

                const existingPayment = await payment_details.findOne({
                    where: { lender_code, sanction_id, tranche_id, due_date }
                });

                if (existingPayment) {
                    // Optional: update main table if required
                    return payment_details.update(
                        {
                            ...schedule,
                            approval_status: "Approved",
                            remarks: schedule.remarks || null
                        },
                        {
                            where: { lender_code, sanction_id, tranche_id, due_date }
                        }
                    );
                } else {
                    return payment_details.create({
                        ...schedule,
                        approval_status: "Approved",
                        remarks: schedule.remarks || null
                    });
                }
            })
        );

        // Update staging table approval status
        const stagingUpdateResults = await Promise.all(
            schedules.map((schedule) => {
                const { lender_code, due_date } = schedule;

                return payment_details_staging.update(
                    {
                        approval_status: "Approved",
                        remarks: schedule.remarks || null
                    },
                    {
                        where: {
                            lender_code,
                            sanction_id,
                            tranche_id,
                            due_date,
                            approval_status: "Approval Pending"
                        }
                    }
                );
            })
        );

        res.status(201).json({
            message: "UTR Upload(s) approved successfully.",
            approved: approvalResults.length,
            updatedStaging: stagingUpdateResults.length
        });

    } catch (error) {
        console.error("UTR Approve Error:", error);

        if (error.name === 'SequelizeForeignKeyConstraintError') {
            return res.status(400).json({
                message: "Tranche approval is pending — please approve the tranche before approving the UTR Upload.",
                detail: error.parent?.detail || error.message
            });
        }

        res.status(500).json({
            message: "Internal server error while approving UTR Upload.",
            error: error.message
        });
    }
};



exports.UTRReject = async (req, res) => {
    console.log("Received Rejected Data:", req.body);

    try {
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of UTR"
            });
        }

        // Validate required fields in the first schedule item
        const sanction_id = req.body[0]?.sanction_id;
        const tranche_id = req.body[0]?.tranche_id;

        if (!sanction_id || !tranche_id) {
            return res.status(400).json({
                message: "Missing sanction_id or tranche_id in schedule data"
            });
        }

        // Optional: Check if the tranche exists
        const checkExistingTranche = await tranche_details.findAll({
            where: { sanction_id, tranche_id }
        });
        console.log("Existing tranche_details for rejection:", checkExistingTranche);

        const updateResults = await Promise.all(
            req.body.map(async (schedule) => {
                return await payment_details_staging.update(
                    {
                        approval_status: "Rejected",
                        createdat: new Date(),
                        remarks: schedule.remarks
                    },
                    {
                        where: {
                            sanction_id: schedule.sanction_id,
                            tranche_id: schedule.tranche_id,
                            approval_status: "Approval Pending"
                        }
                    }
                );
            })
        );

        res.status(201).json({
            message: "UTR Upload rejected successfully",
            updates: updateResults
        });

    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({
            message: "Internal server error",
            error: error.message
        });
    }
};


exports.UTRPending = async (req, res) => {
    try {
        const intrestPending = await payment_details_staging.findAll({
            where: { approval_status: "Approval Pending" }
        });

        if (!intrestPending || intrestPending.length === 0) {
            return res.status(404).json({ message: "No Pending UTR Upload found" });
        }

        res.status(201).json({ success: true, data: intrestPending });
    } catch (error) {
        console.error("Error fetching UTR Upload:", error);
        res.status(500).json({ message: "Error fetching UTR Upload", error: error.message });
    }
}