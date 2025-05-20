const { sequelize } = require('../config/db');
const initModels = require('../models/init-models');
const { Op } = require("sequelize");
const jwt = require("jsonwebtoken");

// const tranche_details = require('../models/tranche_details');

const models = initModels(sequelize);
const { tranche_details_staging, tranche_details, repayment_schedule, repayment_schedule_staging } = models;

// Sending Create Executed Documents  
exports.trancheCreate = async (req, res) => {
    const data = req.body.finalFormData;
    console.log("data tranche create: ", data)
    let temp = data.createdby;
    console.log("temp: ", temp)
    const JWT_SECRET = process.env.JWT_SECRET;
    console.log("secret: ", JWT_SECRET)
    const decoded = jwt.verify(temp, JWT_SECRET);
    console.log("decode: ", decoded)

    const trancheData = {
        tranche_id: data.tranche_id,
        sanction_id: data.sanction_id,
        tranche_date: data.tranche_date,
        tranche_number: data.tranche_number,
        tranche_amount: data.tranche_amount || null,
        interest_type: data.interest_type,
        interest_rate: data.interest_rate || null,
        tenure_months: data.tenure_months || null,
        principal_start_date: data.principal_start_date,
        interest_start_date: data.interest_start_date,
        principal_payment_frequency: data.principal_payment_frequency,
        interest_payment_frequency: data.interest_payment_frequency,
        applicable_of_leap_year: data.applicable_of_leap_year,
        interest_calculation_days: data.interest_calculation_days,
        moratorium_start_date: data.moratorium_start_date,
        moratorium_end_date: data.moratorium_end_date,
        // createdat: data.createdat || new Date(),
        // updatedat: data.updatedat || new Date(),
        remarks: data.remarks || null,
        approval_status: data.approval_status || "Approval Pending",
        createdby: decoded.id,
        updatedby: decoded.id,
        user_type: "N",
        lender_code: data.lender_code,
        current_ac_no: data.current_ac_no,
        bank_name: data.bank_name,
        bank_branch: data.bank_branch,
        location: data.location,
        ifsc_code: data.ifsc_code,
        conf_acc_no: data.conf_acc_no
    };

    try {
        const newTranche = await tranche_details_staging.create(trancheData);
        res.status(201).json({ message: "Tranche added successfully", data: newTranche });
    } catch (error) {
        console.error("Tranche Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};

// Tranche Staging from main table get approved API all data
exports.trancheFetch = async (req, res) => {
    const datagot = req.body;
    try {
        const tranche = await tranche_details_staging.findAll({
            attributes: [
                "lender_code", "tranche_id", "sanction_id", "tranche_amount", "interest_type", "interest_rate", "approval_status", "createdat", "updatedat", "id"
            ], where: {
                approval_status: { [Op.or]: ["Approval Pending", "Rejected"] }
            }
        });
        const tranchemain = await tranche_details.findAll({
            attributes: [
                "lender_code", "tranche_id", "sanction_id", "tranche_amount", "interest_type", "interest_rate", "approval_status", "createdat", "updatedat", "id"
            ], where: { approval_status: "Approved" }
        });

        return res.status(201).json({ success: true, data: tranche, mainData: tranchemain });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
};

// exports.trancheTwo = async (req, res) => {

//     const datagot = req.body;
//     try {
//         // const tranchemain = await tranche_details.findAll({
//         //     attributes: [
//         //         "tranche_id", "sanction_id", "lender_code"
//         //     ], where: { approval_status: "Approved" }
//         // });
//         const tranchemain = await tranche_details.findAll({
//             attributes: [
//                 "tranche_id", "sanction_id", "lender_code", "interest_start_date"
//             ]
//             // , where: { approval_status: "Approved" }
//         });

//         return res.status(201).json({ success: true, data: tranchemain });
//     } catch (error) {
//         console.error("Error fetching tranche:", error);
//         return res.status(500).json({ success: false, message: "Server Error", error: error.message });
//     }
// }



exports.trancheTwo = async (req, res) => {
    const { flag } = req.body; // Or use req.query if sent via query param

    try {
        // Define base attributes
        const attributes = ["tranche_id", "sanction_id", "lender_code", "interest_start_date"];

        // Add "bank_name" if flag == 1
        if (flag == 1) {
            attributes.push("bank_name");
        }

        const tranchemain = await tranche_details.findAll({
            attributes: attributes
            // , where: { approval_status: "Approved" }
        });

        return res.status(201).json({ success: true, data: tranchemain });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
}
  
exports.trancheCheck = async (req, res) => {

    const datagot = req.body;
    try {
        // const tranchemain = await tranche_details.findAll({
        //     attributes: [
        //         "tranche_id", "sanction_id", "lender_code"
        //     ], where: { approval_status: "Approved" }
        // });
        const tranchemain = await tranche_details_staging.findAll({
            attributes: [
                "tranche_id", "sanction_id", "lender_code"
            ]
            // , where: { approval_status: "Approved" }
        });

        return res.status(201).json({ success: true, data: tranchemain });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
}

// Tranche View get by tranche_id Page API
exports.trancheView = async (req, res) => {
    const { lender_code, id, sanction_id, tranche_id, approval_status, updatedat } = req.query;
    // const originalDate = new Date(updatedat);
    // const updatedDate = new Date(originalDate.getTime() + (5 * 60 + 30) * 60 * 1000);
    // console.log("backend tranche 5 getall: ", lender_code, sanction_id, tranche_id, approval_status, updatedDate);
    // console.log("backend 5 getall: ", lender_code, id, sanction_id, tranche_id, approval_status);
    try {
        if (approval_status === 'Approved') {
            const tranche = await tranche_details.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , updatedat: updatedDate 
                }
            });
            if (tranche) {
                // Now get the repayment_type from repayment_schedule
                const repayment = await repayment_schedule_staging.findOne({
                    where: {
                        lender_code: lender_code,
                        sanction_id: sanction_id,
                        tranche_id: tranche_id,
                        // due_date: tranche.due_date
                    },
                    attributes: ['repayment_type']
                });

                if (repayment) {
                    return res.status(200).json({
                        tranche,
                        repayment_type: repayment.repayment_type
                    });
                } else {
                    // return res.status(404).json({ message: "Repayment schedule not found" });
                    return res.status(200).json({
                        tranche,
                    });
                }
            } else {
                return res.status(404).json({ message: "Approved Tranche not found" });
            }
        } else if (approval_status === 'Approval Pending') {
            const tranche = await tranche_details_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , updatedat: updatedDate 
                }
            });
            if (tranche) {
                const repayment = await repayment_schedule_staging.findOne({
                    where: {
                        lender_code: lender_code,
                        sanction_id: sanction_id,
                        tranche_id: tranche_id,
                        // due_date: tranche.due_date
                    },
                    attributes: ['repayment_type']
                });

                if (repayment) {
                    return res.status(200).json({
                        tranche,
                        repayment_type: repayment.repayment_type
                    });
                } else {
                    return res.status(404).json({ message: "Repayment schedule not found" });
                }
            } else {
                return res.status(404).json({ message: "Approved Tranche not found" });
            }
        }
        else if (approval_status === 'Rejected') {
            const tranche = await tranche_details_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , updatedat: updatedDate
                }
            });
            if (tranche) {
                const repayment = await repayment_schedule_staging.findOne({
                    where: {
                        lender_code: lender_code,
                        sanction_id: sanction_id,
                        tranche_id: tranche_id,
                        // due_date: tranche.due_date
                    },
                    attributes: ['repayment_type']
                });

                if (repayment) {
                    return res.status(200).json({
                        tranche,
                        repayment_type: repayment.repayment_type
                    });
                } else {
                    return res.status(404).json({ message: "Repayment schedule not found" });
                }
            } else {
                return res.status(404).json({ message: "Approved Tranche not found" });
            }
        } else {
            return res.status(400).json({ message: "Invalid approval status" });
        }

    } catch (error) {
        console.error("Error fetching Sanction:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};

// Tranche Update Data Staging and Main tables
exports.trancheUpdate = async (req, res) => {
    const { sanction_id, tranche_id, lender_code, user_type } = req.body;
    const data = req.body;
    data.id = null;
    // console.log("data: ", data);
    const newData = data;
    // console.log("new object data: ", newData);
    try {
        const JWT_SECRET = process.env.JWT_SECRET;
        const decoded = jwt.verify(data.createdby, JWT_SECRET);

        // 🔐 Global check for any pending approval record in staging
        const existingStagingLender = await tranche_details_staging.findOne({
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
        const existingLender = await tranche_details.findOne({
            where: {
                lender_code,
                sanction_id,
                tranche_id,
                approval_status: "Approved"
            }
        });

        // Check for Rejected records in staging
        const rejectedStagingLenders = await tranche_details_staging.findAll({
            where: {
                lender_code,
                sanction_id,
                tranche_id,
                approval_status: "Rejected"
            }
        });

        // Case 1: user_type === "N" (New record)
        if (user_type === "N") {
            const existsInMaster = await tranche_details.findOne({
                where: {
                    lender_code,
                    sanction_id,
                    tranche_id
                }
            });

            if (existsInMaster) {

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

            const newStagingRecord = await tranche_details_staging.create(newRecord);

            return res.status(201).json({
                status: "success",
                message: "New Tranche created .",
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

            const newStagingRecord = await tranche_details_staging.create(newRecord);

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

            const newStagingRecord = await tranche_details_staging.create(recordWithPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "Tranche update request is in progress. No further updates allowed until approved.",
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

            const newStagingRecord = await tranche_details_staging.create(recordPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "New record created for previously rejected lenderCode,sanctionID,TrancheID.",
                NewStagingRecord: newStagingRecord,
                updatedFields: updatedFields
            });
        }

        // Final fallback: attempt to update existing staging record
        const [updateCount, updatedRecords] = await roc_forms_staging.update(data, {
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
                message: "Tranche not found or no changes detected."
            });
        }

        return res.status(200).json({
            status: "success",
            message: "Tranche updated successfully.",
            updatedFields: updatedFields,
            UpdatedLender: updatedRecords ? updatedRecords[0] : null
        });

    } catch (error) {
        console.error("Update Error:", error);
        res.status(500).json({ status: "error", message: "Internal server error", error: error.message });
    }
};



// Tranche Details Approval API
exports.trancheApprove = async (req, res) => {
    try {
        console.log("approve Tranche backend:", req.body)
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Tranche"
            });
        }
        // 1. Upsert lenders in lender_master (update if exists, create if new)
        const newSanction = await Promise.all(
            req.body.map(async (sanction) => {
                const existingSanction = await tranche_details.findOne({
                    where: {
                        sanction_id: sanction.sanction_id,
                        lender_code: sanction.lender_code,
                        tranche_id: sanction.tranche_id
                    }
                });

                if (existingSanction) {
                    // Lender exists, update the record
                    await tranche_details.update(
                        {
                            ...sanction,
                            updatedat: new Date(),
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
                    const newLenderRecord = await tranche_details.create({
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
                return await tranche_details_staging.update(
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
                            id: sanction.id,
                            approval_status: "Approval Pending"  // Only update those with "Approval Pending" status
                        }
                    }
                );
            })
        );
        console.log("Update staging: ", updatedLenders)

        res.status(201).json({
            message: "Tranche added successfully",
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
// Tranche reject API
exports.trancheReject = async (req, res) => {
    try {
        console.log("Received Tranche rejection Data:", req.body);

        // Ensure req.body is an array for bulk insert
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Tranche"
            });
        }

        const updatePromises = req.body.map(sanction => {
            return tranche_details_staging.update(
                {
                    approval_status: "Rejected",
                    user_type: sanction.user_type,
                    remarks: sanction.remarks
                },
                {
                    where: {
                        sanction_id: sanction.sanction_id,
                        id: sanction.id,
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


// Tranche approval pending API
exports.tranchePending = async (req, res) => {
    try {
        const tranchePending = await tranche_details_staging.findAll({
            where: { approval_status: "Approval Pending" }
        });

        if (!tranchePending || tranchePending.length === 0) {
            return res.status(404).json({ message: "No Pending Tranches found" });
        }

        res.status(201).json({ success: true, data: tranchePending });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};