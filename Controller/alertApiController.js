const { sequelize } = require("../config/db");
// const alert_management = require("../models/alert_management");
require("dotenv").config();

const initModels = require("../models/init-models");

const models = initModels(sequelize);
const { repayment_schedule, alert_management } = models;

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