const { sequelize } = require("../config/db");
require("dotenv").config();
const { Op } = require("sequelize");
const initModels = require("../models/init-models");

const models = initModels(sequelize);
const { repayment_schedule, alert_management, alert_management_staging } = models;

exports.alertData = async (req, res) => {

    try {
        const datafetch = await alert_management.findAll({
            // attributes: [
            //     "sanction_id", "tranche_id", "due_date", "total_due"
            // ],
        });

        return res.status(201).json({ success: true, data: datafetch });
    } catch (error) {
        console.error("Error fetching Alerts:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
}


exports.AlertFetch = async (req, res) => {
    const datagot = req.body;
    try {
        const tranche = await alert_management_staging.findAll({
            attributes: [
                "lender_code", "sanction_id", "tranche_id", "alert_time", "alert_start_date", "alert_end_date", "alert_frequency", "to_addr", "cc_addr", "alert_trigger", "approval_status", "createdat"
            ], where: {
                approval_status: { [Op.or]: ["Approval Pending", "Rejected"] }
            }
        });
        const tranchemain = await alert_management.findAll({
            attributes: [
                "lender_code", "sanction_id", "tranche_id", "alert_time", "alert_start_date", "alert_end_date", "alert_frequency", "to_addr", "cc_addr", "alert_trigger", "approval_status", "createdat"
            ], where: { approval_status: "Approved" }
        });

        return res.status(201).json({ success: true, data: tranche, mainData: tranchemain });
    } catch (error) {
        console.error("Error fetching tranche:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
}

exports.AlertView = async (req, res) => {
    const { lender_code, sanction_id, tranche_id, approval_status, createdat } = req.query;
    console.log("Received Alert View Data:", req.query);
    try {
        if (approval_status === 'Approved') {
            const interest = await alert_management.findOne({
                where: {
                    lender_code: lender_code,
                    sanction_id: sanction_id,
                    tranche_id: tranche_id,
                    approval_status: approval_status
                    // , createdat: createDate 
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "Approved Alert Details not found" });
            }
        } else if (approval_status === 'Approval Pending') {
            const interest = await alert_management_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , createdat: createDate
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "Approval Pending Alert Details not found" });
            }
        }
        else if (approval_status === 'Rejected') {
            const interest = await alert_management_staging.findOne({
                where: {
                    lender_code: lender_code, sanction_id: sanction_id, tranche_id: tranche_id, approval_status: approval_status
                    // , createdat: createDate 
                }
            });
            if (interest) {
                // console.log("Data View: ", interest)
                return res.status(200).json({ interest });
            } else {
                return res.status(404).json({ message: "Alert Details not found" });
            }
        } else {
            return res.status(400).json({ message: "Invalid Approval status" });
        }


    } catch (error) {
        console.error("Error fetching Alert Details:", error);
        res.status(500).json({ message: "Internal server error Alert Details", error: error.message });
    }
}


exports.AlertApprove = async (req, res) => {
    console.log("Approve Alert Details Approve Backend:", req.body);

    try {
        const schedules = req.body;

        if (!Array.isArray(schedules) || schedules.length === 0) {
            return res.status(400).json({
                message: "Invalid or empty request body. Expected a non-empty array of Alert Details."
            });
        }

        const { sanction_id, tranche_id } = schedules[0];

        if (!sanction_id || !tranche_id) {
            return res.status(400).json({
                message: "Missing 'sanction_id' or 'tranche_id' in Alert Details data."
            });
        }

        // Check if sanction and tranche exist (optional validation)
        // const [sanctionExists, trancheExists] = await Promise.all([
        //     sanction_details.findOne({ where: { sanction_id } }),
        //     tranche_details.findOne({ where: { sanction_id, tranche_id } })
        // ]);

        // if (!sanctionExists) {
        //     return res.status(404).json({ message: `Sanction not found for Sanction ID: ${sanction_id}` });
        // }
        // if (!trancheExists) {
        //     return res.status(404).json({ message: `Tranche not found for Sanction ID: ${sanction_id}, Tranche ID: ${tranche_id}` });
        // }

        // Process each schedule item
        const approvalResults = await Promise.all(
            schedules.map(async (schedule) => {
                const { lender_code, alert_end_date } = schedule;

                const existingPayment = await alert_management.findOne({
                    where: { lender_code, sanction_id, tranche_id, alert_end_date }
                });

                if (existingPayment) {
                    // Optional: update main table if required
                    return alert_management.update(
                        {
                            ...schedule,
                            approval_status: "Approved",
                            remarks: schedule.remarks || null
                        },
                        {
                            where: { lender_code, sanction_id, tranche_id, alert_end_date }
                        }
                    );
                } else {
                    return alert_management.create({
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
                const { lender_code, alert_end_date } = schedule;

                return alert_management_staging.update(
                    {
                        approval_status: "Approved",
                        remarks: schedule.remarks || null
                    },
                    {
                        where: {
                            lender_code,
                            sanction_id,
                            tranche_id,
                            alert_end_date: schedule.alert_end_date,
                            approval_status: "Approval Pending"
                        }
                    }
                );
            })
        );

        res.status(201).json({
            message: "Alert Detail's approved successfully.",
            approved: approvalResults.length,
            updatedStaging: stagingUpdateResults.length
        });

    } catch (error) {
        console.error("Alert Detail's Approve Error:", error);

        if (error.name === 'SequelizeForeignKeyConstraintError') {
            return res.status(400).json({
                message: "Tranche approval is pending — please approve the tranche before approving the Alert Detail's.",
                detail: error.parent?.detail || error.message
            });
        }

        res.status(500).json({
            message: "Internal server error while approving Alert Detail's.",
            error: error.message
        });
    }
};


exports.AlertReject = async (req, res) => {
    console.log("Received Alert Rejected Data:", req.body);

    try {
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Alert Details."
            });
        }

        // Validate required fields in the first schedule item
        const sanction_id = req.body[0]?.sanction_id;
        const tranche_id = req.body[0]?.tranche_id;

        if (!sanction_id || !tranche_id) {
            return res.status(400).json({
                message: "Missing sanction_id or tranche_id in Alert Details data"
            });
        }

        // Optional: Check if the tranche exists
        // const checkExistingTranche = await tranche_details.findAll({
        //     where: { sanction_id, tranche_id }
        // });
        // console.log("Existing tranche_details for rejection:", checkExistingTranche);

        const updateResults = await Promise.all(
            req.body.map(async (schedule) => {
                return await alert_management_staging.update(
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
            message: "Alert Details Rejected successfully",
            updates: updateResults
        });

    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({
            message: "Internal server Error: Alert Details Rejected",
            error: error.message
        });
    }
};

exports.AlertPending = async (req, res) => {
    try {
        const alertPending = await alert_management_staging.findAll({
            where: { approval_status: "Approval Pending" }
        });

        if (!alertPending || alertPending.length === 0) {
            return res.status(404).json({ message: "No Pending Alert Details found" });
        }

        res.status(201).json({ success: true, data: alertPending });
    } catch (error) {
        console.error("Error fetching Alert Details:", error);
        res.status(500).json({ message: "Error fetching Alert Details", error: error.message });
    }
}