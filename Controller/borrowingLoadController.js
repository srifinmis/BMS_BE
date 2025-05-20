const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");

const models = initModels(sequelize);
const { lender_master_staging, lender_master } = models;

exports.getLoadData = async (req, res) => {
    try {
        const lender1countapproved = await lender_master.count({ where: { approval_status: "Approved" } });
        const lender2countpending = await lender_master_staging.count({ where: { approval_status: "Approval Pending" } });
        const lender2countreject = await lender_master_staging.count({ where: { approval_status: "Rejected" } });
        const totalcount = lender1countapproved + lender2countpending + lender2countreject;

        res.status(200).json({ Approved: lender1countapproved, Rejected: lender2countreject, ApprovalPending: lender2countpending, Total: totalcount });
    } catch (error) {
        console.error("Error fetching lender:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}