const { sequelize } = require('../config/db');
const initModels = require('../models/init-models');
const { Op } = require("sequelize");
const jwt = require("jsonwebtoken");

const models = initModels(sequelize);
const { interest_rate_changes, interest_rate_changes_staging, tranche_details } = models;

// Sending Create Executed Documents  
exports.intrestCreate = async (req, res) => {
    const data = req.body;
    console.log("intrest backend data: ", data)
    let temp = data.createdby;
    const JWT_SECRET = process.env.JWT_SECRET;
    const decoded = jwt.verify(temp, JWT_SECRET);
    const intrestData = {
        // change_id: data.change_id,
        new_interest_rate: data.new_interest_rate,
        effective_date: data.effective_date,
        tranche_id: data.tranche_id,
        sanction_id: data.sanction_id,
        lender_code: data.lender_code,
        user_type: "N",
        // createdat: new Date(),
        updatedat: data.updatedat,
        createdby: decoded.id || "vishal",
        updatedby: data.updatedby || "vishal",
        remarks: data.remarks || null,
        approval_status: data.approval_status || "Approval Pending"
    };
    // console.log('Data from FD:  ', datagot);
    try {

        const newRoc = await interest_rate_changes_staging.create(intrestData);
        res.status(201).json({ message: "Intrest Rate added successfully", data: newRoc });
    } catch (error) {
        console.error("Intrest Rate Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }

}

exports.interestThree = async (req, res) => {
    try {
        const tranche = await interest_rate_changes.findAll({
            attributes: [
                'lender_code', 'sanction_id', 'tranche_id'
            ],
            include: [{
                model: tranche_details,
                as: 'tranche',
                attributes: ['interest_start_date']
            }]
        });


        return res.status(201).json({ success: true, data: tranche });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }

}
exports.intrestFetch = async (req, res) => {
    const datagot = req.body;
    try {
        const tranche = await interest_rate_changes_staging.findAll({
            attributes: [
                "lender_code", "sanction_id", "tranche_id", "new_interest_rate", "effective_date", "approval_status", "createdat"
            ], where: {
                approval_status: { [Op.or]: ["Approval Pending", "Rejected"] }
            }
        });
        const tranchemain = await interest_rate_changes.findAll({
            attributes: [
                "lender_code", "sanction_id", "tranche_id", "new_interest_rate", "effective_date", "approval_status", "createdat"
            ], where: { approval_status: "Approved" }
        });

        return res.status(201).json({ success: true, data: tranche, mainData: tranchemain });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
}

exports.intrestView = async (req, res) => {
    const { lender_code, sanction_id, tranche_id, approval_status, createdat } = req.query;
    // const originalDate = new Date(createdat);
    // const createDate = new Date(originalDate.getTime() + (5 * 60 + 30) * 60 * 1000);
    // console.log("backend Interest 5 getall: ", lender_code, sanction_id, tranche_id, approval_status, createDate);

    // const { sanction_id, approval_status } = req.query;
    // console.log("backend 2 getall: ", tranche_id, approval_status);
    try {
        if (approval_status === 'Approved') {
            const interest = await interest_rate_changes.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , createdat: createDate 
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "Approved Interest not found" });
            }
        } else if (approval_status === 'Approval Pending') {
            const interest = await interest_rate_changes_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , createdat: createDate
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "Approval Pending Interest not found" });
            }
        }
        else if (approval_status === 'Rejected') {
            const interest = await interest_rate_changes_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , createdat: createDate 
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "Rejected Interest not found" });
            }
        } else {
            return res.status(400).json({ message: "Invalid approval status" });
        }


    } catch (error) {
        console.error("Error fetching Interest:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}

exports.intrestUpdate = async (req, res) => {
    const { sanction_id, tranche_id, lender_code, user_type, approval_status, updatedat } = req.body;
    const data = req.body;
    data.change_id = null;
    console.log("data: ", data);
    const newData = data;
    console.log("new object data: ", newData);
    try {
        const JWT_SECRET = process.env.JWT_SECRET;
        // const decoded = jwt.verify(data.updatedFormData.createdby, JWT_SECRET);
        const decoded = jwt.verify(data.createdby, JWT_SECRET);
        // 🔐 Global check for any pending approval record in staging
        const existingStagingLender = await interest_rate_changes_staging.findOne({
            where: {
                lender_code,
                sanction_id,
                tranche_id,
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
        const existingLender = await interest_rate_changes.findOne({
            where: {
                lender_code,
                sanction_id,
                tranche_id,
                approval_status: "Approved"
            }
        });

        // Check for Rejected records in staging
        const rejectedStagingLenders = await interest_rate_changes_staging.findAll({
            where: {
                lender_code,
                sanction_id,
                tranche_id,
                approval_status: "Rejected"
            }
        });

        // Case 1: user_type === "N" (New record)
        if (user_type === "N") {
            const existsInMaster = await interest_rate_changes.findOne({
                where: {
                    lender_code,
                    sanction_id,
                    tranche_id
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
                updatedat: new Date(),
                approval_status: "Approval Pending",
                createdby: decoded.id,
                updatedby: data.updatedby,
                updated_fields: updatedFields,
                id: null
            };
            console.log("new update: ", newRecord)
            const newStagingRecord = await interest_rate_changes_staging.create(newRecord);

            return res.status(201).json({
                status: "success",
                message: "New Interest created .",
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
                updatedby: data.updatedby,
                id: null
            };

            const newStagingRecord = await interest_rate_changes_staging.create(newRecord);

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
                updatedby: data.updatedby
            };

            const newStagingRecord = await interest_rate_changes_staging.create(recordWithPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "Interest update request is in progress. No further updates allowed until approved.",
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
                updatedby: data.updatedby,
                user_type: "U",
                updated_fields: updatedFields,
                id: null
            };

            const newStagingRecord = await interest_rate_changes_staging.create(recordPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "New record created for previously rejected lenderCode,sanctionID,TrancheID.",
                NewStagingRecord: newStagingRecord,
                updatedFields: updatedFields
            });
        }

        // Final fallback: attempt to update existing staging record
        const [updateCount, updatedRecords] = await interest_rate_changes_staging.update(data, {
            where: {
                lender_code,
                sanction_id,
                tranche_id
            },
            returning: true
        });

        if (updateCount === 0) {
            return res.status(404).json({
                status: "error",
                message: "Interest not found or no changes detected."
            });
        }

        return res.status(200).json({
            status: "success",
            message: "Interest updated successfully.",
            updatedFields: updatedFields,
            UpdatedLender: updatedRecords ? updatedRecords[0] : null
        });

    } catch (error) {
        console.error("Update Error:", error);
        res.status(500).json({ status: "error", message: "Internal server error", error: error.message });
    }
};



exports.intrestApprove = async (req, res) => {
    try {
        console.log("approve Interest Rate backend:", req.body)
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Interest Rate"
            });
        }
        // 1. Upsert lenders in lender_master (update if exists, create if new)
        const newSanction = await Promise.all(
            req.body.map(async (sanction) => {
                const existingSanction = await interest_rate_changes.findOne({
                    where: {
                        sanction_id: sanction.sanction_id,
                        lender_code: sanction.lender_code,
                        tranche_id: sanction.tranche_id
                    }
                });

                if (existingSanction) {
                    // Lender exists, update the record
                    await interest_rate_changes.update(
                        {
                            ...sanction,
                            remarks: sanction.remarks,
                            approval_status: "Approved",
                        },
                        {
                            where: {
                                sanction_id: sanction.sanction_id,
                                lender_code: sanction.lender_code,
                                tranche_id: sanction.tranche_id
                            }
                        }
                    );

                    return existingSanction;
                } else {
                    // Lender does not exist, insert new record
                    const newLenderRecord = await interest_rate_changes.create({
                        ...sanction,
                        sanction_id: sanction.sanction_id,
                        createdat: new Date(),
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
                return await interest_rate_changes_staging.update(
                    {
                        approval_status: "Approved",
                        user_type: "U",
                        remarks: sanction.remarks  // Set remarks for each lender
                    },
                    {
                        where: {
                            sanction_id: sanction.sanction_id,
                            lender_code: sanction.lender_code,
                            tranche_id: sanction.tranche_id,
                            change_id: sanction.change_id,
                            approval_status: "Approval Pending"  // Only update those with "Approval Pending" status
                        }
                    }
                );
            })
        );
        console.log("Update staging: ", updatedLenders)

        res.status(201).json({
            message: "Interest Rate added successfully",
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

exports.intrestReject = async (req, res) => {
    try {
        console.log("Received Interest Rate rejection Data:", req.body);

        // Ensure req.body is an array for bulk insert
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Interest Rate"
            });
        }

        const updatePromises = req.body.map(sanction => {
            return interest_rate_changes_staging.update(
                {
                    approval_status: "Rejected",
                    user_type: sanction.user_type,
                    remarks: sanction.remarks
                },
                {
                    where: {
                        sanction_id: sanction.sanction_id,
                        change_id: sanction.change_id,
                        lender_code: sanction.lender_code,
                        tranche_id: sanction.tranche_id,
                        approval_status: "Approval Pending"
                    }
                }
            );
        });

        // Wait for all updates to finish
        await Promise.all(updatePromises);

        res.status(201).json({
            message: "Interest Rate Rejected successfully"
        });
    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({
            message: "Internal server error",
            error: error.message
        });
    }
}


exports.intrestPending = async (req, res) => {
    try {
        const intrestPending = await interest_rate_changes_staging.findAll({
            where: { approval_status: "Approval Pending" }
        });

        if (!intrestPending || intrestPending.length === 0) {
            return res.status(404).json({ message: "No Pending interest rate found" });
        }

        res.status(201).json({ success: true, data: intrestPending });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}