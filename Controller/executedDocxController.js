const { sequelize } = require('../config/db');
const initModels = require('../models/init-models');
const { Op } = require("sequelize");
const jwt = require("jsonwebtoken");
const path = require("path");
const fs = require("fs");

const models = initModels(sequelize);
const { sanction_details, executed_documents_staging, executed_documents } = models;

//sending Create Executed Documents  
exports.executedCreate = async (req, res) => {

    const data = req.body;
    let temp = data.createdby;
    const JWT_SECRET = process.env.JWT_SECRET;
    const decoded = jwt.verify(temp, JWT_SECRET);

    const executedData = {
        sanction_id: data.sanctionId,
        document_type: data.documentType,
        file_name: data.fileName,
        uploaded_date: data.uploadedDate,
        document_url: data.fileUrl || null, // Convert "" to null
        approval_status: data.approval_status || "Approval Pending",
        // createdat: new Date(),
        // updatedat: new Date(),
        lender_code: data.lender_code,
        user_type: "N",
        updatedby: decoded.id,
        createdby: decoded.id
    };
    // console.log('Data from FD:  ', data);
    try {

        const newRoc = await executed_documents_staging.create(executedData);
        res.status(201).json({ message: "Executed Documents added successfully", data: newRoc });
    } catch (error) {
        console.error("Login Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};


// Sanction Staging from maintable get approved API all data
exports.executedFetch = async (req, res) => {
    const datagot = req.body;
    try {
        const executed = await executed_documents_staging.findAll({
            attributes: [
                "document_id", "lender_code", "sanction_id", "document_type", "uploaded_date", "approval_status", "createdat", "updatedat"
            ], where: {
                approval_status: { [Op.or]: ["Approval Pending", "Rejected"] }
            }
        });
        const executedmain = await executed_documents.findAll({
            attributes: [
                "document_id", "lender_code", "sanction_id", "document_type", "uploaded_date", "approval_status", "createdat", "updatedat"
            ], where: { approval_status: "Approved" }
        });

        return res.status(201).json({ success: true, data: executed, mainData: executedmain });
    } catch (error) {
        console.error("Error fetching sanction:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
};

// Executed get by sanction_id Page API
exports.executedView = async (req, res) => {
    try {
        const { sanction_id } = req.params;
        const sanction = await executed_documents_staging.findAll({
            where: { sanction_id }
        });

        if (!sanction || sanction.length === 0) {
            return res.status(404).json({ message: "Executed not found" });
        }

        res.status(201).json({ success: true, data: sanction });
    } catch (error) {
        console.error("Error fetching Executed:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};
exports.executedValidate = async (req, res) => {
    try {
        const { sanction_id, lender_code } = req.query;
        console.log("backend check executed:", sanction_id, lender_code)
        const sanction = await executed_documents_staging.findAll({
            where: { sanction_id, lender_code }
        });

        if (!sanction || sanction.length === 0) {
            return res.status(404).json({ message: "Executed not found" });
        }

        res.status(201).json({ success: true, data: sanction });
    } catch (error) {
        console.error("Error fetching Executed:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }

}


// Executed Data Staging and Main tables
exports.executedUpdate = async (req, res) => {
    try {
        const updatedexecuted = await executed_documents_staging.update(req.body, { where: { sanction_id: req.body.sanction_id } });
        const updatemainexecuted = await executed_documents.update(req.body, { where: { sanction_id: req.body.sanction_id, approval_status: "Approved" } })
        // console.log("Main backend data: ", updatemainlender)

        res.status(201).json({ message: "Executed Updated Successfully", Update: updatedexecuted, Mainupdate: updatemainexecuted });
    } catch (error) {
        console.error("Executed Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}

// Sanction Details Approval API
exports.executedApprove = async (req, res) => {
    try {
        console.log("approve Executed backend:", req.body)
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Executed"
            });
        }
        // 1. Upsert lenders in lender_master (update if exists, create if new)
        const newSanction = await Promise.all(
            req.body.map(async (sanction) => {
                const existingSanction = await executed_documents.findOne({
                    where: {
                        sanction_id: sanction.sanction_id,
                        lender_code: sanction.lender_code,
                    }
                });

                if (existingSanction) {
                    // Lender exists, update the record
                    await executed_documents.update(
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
                            }
                        }
                    );

                    return existingSanction;
                } else {
                    // Lender does not exist, insert new record
                    const newLenderRecord = await executed_documents.create({
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
                return await executed_documents_staging.update(
                    {
                        approval_status: "Approved",
                        user_type: "U",
                        remarks: sanction.remarks  // Set remarks for each lender
                    },
                    {
                        where: {
                            sanction_id: sanction.sanction_id,
                            lender_code: sanction.lender_code,
                            document_id: sanction.document_id,
                            approval_status: "Approval Pending"  // Only update those with "Approval Pending" status
                        }
                    }
                );
            })
        );
        console.log("Update staging: ", updatedLenders)

        res.status(201).json({
            message: "Executed added successfully",
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



// Executed reject API
exports.executedReject = async (req, res) => {
    try {
        console.log("Received Executed rejection Data:", req.body);

        // Ensure req.body is an array for bulk insert
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Executed"
            });
        }

        const updatePromises = req.body.map(sanction => {
            return executed_documents_staging.update(
                {
                    approval_status: "Rejected",
                    user_type: sanction.user_type,
                    remarks: sanction.remarks
                },
                {
                    where: {
                        sanction_id: sanction.sanction_id,
                        document_id: sanction.document_id,
                        lender_code: sanction.lender_code,
                        approval_status: "Approval Pending"
                    }
                }
            );
        });

        // Wait for all updates to finish
        await Promise.all(updatePromises);

        res.status(201).json({
            message: "Executed Rejected successfully"
        });
    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({
            message: "Internal server error",
            error: error.message
        });
    }
}


//Executed apprivalpending api

exports.executedPending = async (req, res) => {
    try {
        // const { lender_code } = req.params;
        const sanction = await executed_documents_staging.findAll({
            where: { approval_status: "Approval Pending" }
        });

        if (!sanction || sanction.length === 0) {
            return res.status(404).json({ message: "Executed Details not found" });
        }

        res.status(201).json({ success: true, data: sanction });
    } catch (error) {
        console.error("Error fetching lender:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}

//app.get("/api/executed/document/:sanction_id",
exports.viewdocument = async (req, res) => {
    const { sanction_id, lender_code } = req.query;
    console.log("backend params view executed: ", sanction_id, lender_code)

    try {
        // Fetch document details from DB
        const sanction = await executed_documents_staging.findOne({
            where: { sanction_id, lender_code }
        });
        console.log("sanction data view : ", sanction)

        if (!sanction || !sanction.document_url) {
            return res.status(404).json({ success: false, message: "Executed document not found" });
        }

        // Define absolute file path
        const filePath = sanction.document_url;

        // Check if file exists
        if (!fs.existsSync(filePath)) {
            return res.status(404).json({
                status: "error",
                message: "File not found on server,Please Reupload File"
            });
        }

        // Set response headers to display PDF
        res.setHeader("Content-Type", "application/pdf");
        res.sendFile(filePath);
    } catch (error) {
        console.error("Error fetching Executed document:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};

