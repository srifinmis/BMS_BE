const { sequelize } = require('../config/db');
const initModels = require('../models/init-models');
const { Op } = require("sequelize");
const jwt = require("jsonwebtoken");

const models = initModels(sequelize);
const { lender_master, lender_master_staging, sanction_details_staging, sanction_details } = models;


//sending Sanction_Details 
exports.sanctionCreate = async (req, res) => {
    const data = req.body;
    let temp = data.createdby;
    const JWT_SECRET = process.env.JWT_SECRET;
    const decoded = jwt.verify(temp, JWT_SECRET);
    // console.log('Data from FD:  ', datagot);
    const sanitizedData = {
        id: data.id,
        sanction_id: data.sanction_id,
        lender_code: data.lender_code,
        loan_type: data.loan_type,
        purpose_of_loan: data.purpose_of_loan,
        interest_type: data.interest_type,
        interest_rate_fixed: data.interest_rate_fixed || null, // Convert "" to null
        benchmark_rate: data.benchmark_rate || null,
        spread_floating: data.spread_floating || null,
        tenure_months: data.tenure_months || null,
        principal_repayment_term: data.principal_repayment_term,
        interest_payment_term: data.interest_payment_term,
        sanction_validity: data.sanction_validity,
        sanction_amount: data.sanction_amount || null,
        processing_fee: data.processing_fee || null,
        other_expenses: data.other_expenses || null, // Added new field
        book_debt_margin: data.book_debt_margin || null,
        cash_margin: data.cash_margin || null,
        prepayment_charges: data.prepayment_charges || null,
        corporate_guarantee: data.corporate_guarantee || null,
        penal_charges: data.penal_charges || null,
        syndication_fee: data.syndication_fee || null,
        syndicated_by: data.syndicated_by,
        sanction_date: data.sanction_date,
        createdat: new Date(),
        updatedat: new Date(),
        createdby: decoded.id,
        updatedby: decoded.id,
        remarks: data.remarks,
        user_type: "N",
        approval_status: data.approval_status || "Approval Pending"
    };

    try {

        const newSanction = await sanction_details_staging.create(sanitizedData);
        res.status(201).json({ message: "Sanction added successfully", data: newSanction });
    } catch (error) {
        console.error("Login Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};

//Lender Codes Fetch  API
exports.lenderCodes = async (req, res) => {
    const datagot = req.body;
    // console.log('Data from FD:  ', datagot);
    try {

        const Lenderget = await lender_master.findAll({
            attributes: ["lender_code", "lender_name"],
        });
        // const Lenderget = await lender_master_staging.findAll({
        //     attributes: ["lender_code", "lender_name"],
        // });
        if (!Lenderget || Lenderget.length === 0) {
            return res.status(404).json({ message: "No Pending Lender found" });
        }

        // const Lenderget = await lender_master
        res.status(201).json({ message: "Lender Fetch successfully", data: Lenderget });
    } catch (error) {
        console.error("Login Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};

exports.sanctionCheck = async (req, res) => {
    const datagot = req.body;
    // console.log('Data from sanction id FD:  ', datagot);
    try {

        const sanctionget = await sanction_details_staging.findAll({
            attributes: ["sanction_id", "lender_code"],
        });
        // console.log("roc: ", sanctionget)
        // const Lenderget = await lender_master
        res.status(201).json({ message: "Santion ID Fetch successfully", data: sanctionget });
    } catch (error) {
        console.error("Login Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}


// 
// Sanction Staging from maintable get approved API all data
exports.sanctionFetch = async (req, res) => {
    const { token } = req.query;
    // console.log("datagot fetch sanction: ", { token });
    const JWT_SECRET = process.env.JWT_SECRET;
    const temptoken = token;

    try {

        const decoded = jwt.verify(temptoken, JWT_SECRET);
        // console.log("Decoded Token:", decoded);

        const emp_id = decoded.id;
        const role = decoded.Role;

        const isAdmin = Array.isArray(role) && role.includes("ADMIN");

        const stagingWhereCondition = {
            approval_status: { [Op.or]: ["Approval Pending", "Rejected"] }
        };

        if (!isAdmin) {
            stagingWhereCondition[Op.or] = [
                { createdby: emp_id },
                { updatedby: emp_id }
            ];
        }
        const sanctions = await sanction_details_staging.findAll({
            attributes: [
                "lender_code",
                "sanction_id",
                "sanction_amount",
                "sanction_date",
                "createdat",
                "updatedat",
                "approval_status"
            ],
            where: stagingWhereCondition
        });

        // Step 2: Get distinct lender_codes from sanctions
        const lenderCodes = [...new Set(sanctions.map(s => s.lender_code))];

        // Step 3: Fetch lender names from lender_master_staging
        const lenders = await lender_master_staging.findAll({
            where: { lender_code: lenderCodes },
            attributes: ["lender_code", "lender_name"]
        });

        // Step 4: Convert to map for easy lookup
        const lenderMap = {};
        lenders.forEach(l => {
            lenderMap[l.lender_code] = l.lender_name;
        });

        // Step 5: Attach lender_name to each sanction record
        const sanction = sanctions.map(s => {
            return {
                ...s.get({ plain: true }),
                lender_name: lenderMap[s.lender_code] || null
            };
        });

        // console.log(sanction);

        const mainWhereCondition = {
            approval_status: "Approved"
        };

        if (!isAdmin) {
            mainWhereCondition[Op.or] = [
                { createdby: emp_id },
                { updatedby: emp_id }
            ];
        }

        const sanctionmain = await sanction_details.findAll({
            attributes: [
                "lender_code", "sanction_id", "sanction_amount", "sanction_date", "createdat", "updatedat", "approval_status"
            ],
            // where: { approval_status: "Approved" },
            where: mainWhereCondition,
            include: [
                {
                    model: lender_master,
                    as: 'lender_code_lender_master',
                    attributes: ["lender_name"]
                }
            ]
        });
        console.log("sanction fetch all : ", sanction, " : ", sanctionmain)

        return res.status(201).json({ success: true, data: sanction, mainData: sanctionmain });
    } catch (error) {
        console.error("Error fetching lenders:", error);
        return res.status(500).json({ success: false, message: "Server Error", error: error.message });
    }
};

// Sanctiondeetails get by sanction_id Page API
exports.sanctionView = async (req, res) => {
    const { sanction_id, lender_code, approval_status, updatedat } = req.query;
    // const originalDate = new Date(updatedat);
    // const updatedDate = new Date(originalDate.getTime() + (5 * 60 + 30) * 60 * 1000);
    // console.log("sanction backend getall: ", sanction_id, lender_code, approval_status, updatedDate)
    // console.log("sanction backend 2 getall: ", sanction_id, lender_code, approval_status, updatedat);
    try {
        if (approval_status === 'Approved') {
            const sanction = await sanction_details.findOne({
                where: {
                    sanction_id: sanction_id, lender_code: lender_code, approval_status: approval_status
                    // , updatedat: updatedDate 
                }
            });
            if (sanction) {
                return res.status(200).json({ sanction });
            } else {
                return res.status(404).json({ message: "Approved Sanction not found" });
            }
        } else if (approval_status === 'Approval Pending') {
            const sanction = await sanction_details_staging.findOne({
                where: {
                    sanction_id: sanction_id, lender_code: lender_code, approval_status: approval_status
                    // , updatedat: updatedDate 
                },
                // include: [
                //     {
                //         model: lender_master,
                //         as: 'lender_code_lender_master',
                //         attributes: ["lender_name"]
                //     }
                // ]
            });
            if (sanction) {
                return res.status(200).json({ sanction });
            } else {
                return res.status(404).json({ message: "Approval Pending Sanction not found" });
            }
        }
        else if (approval_status === 'Rejected') {
            const sanction = await sanction_details_staging.findOne({
                where: {
                    sanction_id: sanction_id, lender_code: lender_code, approval_status: approval_status
                    // , updatedat: updatedDate 
                }, include: [
                    {
                        model: lender_master,
                        as: 'lender_code_lender_master',
                        attributes: ["lender_name"]
                    }
                ]
            });
            if (sanction) {
                return res.status(200).json({ sanction });
            } else {
                return res.status(404).json({ message: "Rejected Sanction not found" });
            }
        } else {
            return res.status(400).json({ message: "Invalid approval status" });
        }

    } catch (error) {
        console.error("Error fetching Sanction:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};


// SanctionUpdate Data Staging and Main tables
exports.sanctionUpdate = async (req, res) => {
    const { sanction_id, lender_code, user_type } = req.body;
    const data = req.body;
    data.id = null;
    // console.log("data: ", data);
    const newData = data;
    // console.log("new object data: ", newData);

    try {
        const JWT_SECRET = process.env.JWT_SECRET;
        const decoded = jwt.verify(data.createdby, JWT_SECRET);

        // 🔐 Global check for any pending approval record in staging
        const existingStagingLender = await sanction_details_staging.findOne({
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
        const existingLender = await sanction_details.findOne({
            where: {
                lender_code,
                sanction_id,
                approval_status: "Approved"
            }
        });

        // Check for Rejected records in staging
        const rejectedStagingLenders = await sanction_details_staging.findAll({
            where: {
                lender_code,
                sanction_id,
                approval_status: "Rejected"
            }
        });

        // Case 1: user_type === "N" (New record)
        if (user_type === "N") {
            const existsInMaster = await sanction_details.findOne({
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

            const newStagingRecord = await sanction_details_staging.create(newRecord);

            return res.status(201).json({
                status: "success",
                message: "New Sanction created .",
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

            const newStagingRecord = await sanction_details_staging.create(newRecord);

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

            const newStagingRecord = await sanction_details_staging.create(recordWithPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "Sanction update request is in progress. No further updates allowed until approved.",
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

            const newStagingRecord = await sanction_details_staging.create(recordPendingApproval);

            return res.status(201).json({
                status: "success",
                message: "New record created for previously rejected lender_code.",
                NewStagingRecord: newStagingRecord,
                updatedFields: updatedFields
            });
        }

        // Final fallback: attempt to update existing staging record
        const [updateCount, updatedRecords] = await sanction_details_staging.update(data, {
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
            message: "Sanction updated successfully.",
            updatedFields: updatedFields,
            UpdatedLender: updatedRecords ? updatedRecords[0] : null
        });

    } catch (error) {
        console.error("Update Error:", error);
        res.status(500).json({ status: "error", message: "Internal server error", error: error.message });
    }
};


// Sanction Details Approval API
exports.sanctionApprove = async (req, res) => {
    try {
        // console.log("approve sanction backend:", req.body)
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of Sanction"
            });
        }
        // 1. Upsert lenders in lender_master (update if exists, create if new)
        const newSanction = await Promise.all(
            req.body.map(async (sanction) => {
                const existingSanction = await sanction_details.findOne({
                    where: { sanction_id: sanction.sanction_id, lender_code: sanction.lender_code }
                });

                if (existingSanction) {
                    // Lender exists, update the record
                    await sanction_details.update(
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
                    const newLenderRecord = await sanction_details.create({
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
                return await sanction_details_staging.update(
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
        // console.log("Update staging: ", updatedLenders)

        res.status(201).json({
            message: "Sanction added successfully",
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


// SanctionDetails reject API
exports.sanctionReject = async (req, res) => {
    try {
        // console.log("Received sanction Data:", req.body);

        // Ensure req.body is an array for bulk insert
        if (!Array.isArray(req.body)) {
            return res.status(400).json({
                message: "Invalid data format, expected an array of lenders"
            });
        }

        const updatePromises = req.body.map(sanction => {
            return sanction_details_staging.update(
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
            message: "Sanction Rejected successfully"
        });
    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({
            message: "Internal server error",
            error: error.message
        });
    }
}



//Sanction apprivalpending api
exports.sanctionPending = async (req, res) => {
    try {
        // const { lender_code } = req.params;
        const sanction = await sanction_details_staging.findAll({
            where: { approval_status: "Approval Pending" }
        });

        if (!sanction || sanction.length === 0) {
            return res.status(404).json({ message: "Sanction Details not found" });
        }

        res.status(201).json({ success: true, data: sanction });
    } catch (error) {
        console.error("Error fetching sanction:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}