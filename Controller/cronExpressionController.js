const moment = require("moment");
const dayjs = require("dayjs");

const { sequelize } = require('../config/db');
const initModels = require('../models/init-models');

const models = initModels(sequelize);
const { alert_management } = models;


exports.generateCronExpression = async (req, res) => {
    const datagot = req.body;
    console.log("data: ", datagot)
    const startDate = dayjs(datagot.alert_start_date);
    const endDate = dayjs(datagot.alert_end_date);
    // const data = req.body;
    try {
        // const newLender = await alert_management.create(req.body);
        if (!datagot.sanction_id ||
            !datagot.tranche_id ||
            !datagot.alert_time ||
            !datagot.alert_start_date || !datagot.alert_end_date || !datagot.alert_frequency ||
            !datagot.to_addr || !datagot.cc_addr || datagot.tranche_days === undefined
        ) {
            return res.status(400).json({ message: "Missing required fields" });
        }

        // Extract hours and minutes from the time input
        const [hours, minutes] = datagot.alert_time.split(":");

        let cronExpression = "";

        if (startDate.isAfter(endDate)) {
            console.log("Expired Due Date: ", startDate.format(), endDate.format(), datagot.sanction_id);
            return res.status(400).json({ message: "Due date has already passed" });
        }

        console.log("Running Date: ", startDate, endDate, datagot.sanction_id);

        if (datagot.alert_frequency === "Daily") {
            cronExpression = `${minutes} ${hours} * * *`;
            console.log("Cron Expression for Daily:", cronExpression);
        } else if (datagot.alert_frequency === "Weekly") {
            // Extract the day of the week (0-6) from start_date
            const dayOfWeek = moment(startDate).day();
            cronExpression = `${minutes} ${hours} * * ${dayOfWeek}`;
            console.log("Cron Expression for Weekly:", cronExpression);
        } else {
            return res.status(400).json({ message: "Invalid frequency. Choose 'Daily' or 'Weekly'." });
        }
        // Insert data into the alert_management table
        const newAlert = await alert_management.create({
            alert_id:datagot.tranche_id,
            sanction_id: datagot.sanction_id,
            tranche_id: datagot.tranche_id,
            alert_time: datagot.alert_time,
            alert_start_date: datagot.alert_start_date,
            alert_end_date: datagot.alert_end_date,
            alert_frequency: datagot.alert_frequency,
            // tranche_days: datagot.tranche_days,  // Include tranche_days
            cron_expression: cronExpression, // Save the generated cron expression
            to_addr: datagot.to_addr,
            cc_addr: datagot.cc_addr,
        });

        console.log("Generated Cron Expression:", cronExpression);
        res.status(200).json({ message: "Successfully Inserted Record", cronExpression });

    } catch (error) {
        console.error("Error Generating Cron Expression:", error.message);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
};





// exports.generateCronExpression = async (req, res) => {
//     const datagot = req.body;
//     console.log("data: ", datagot)
//     // const data = req.body;
//     try {
//         // Extract hours and minutes from the time input
//         const [hours, minutes] = datagot.time.split(":");

//         let cronExpression = "";

//         if (datagot.frequency === "daily") {
//             // Every day at a specific time
//             cronExpression = `${minutes} ${hours} * * *`;
//             console.log("Cron Expression for Daily:", cronExpression);
//         } else if (datagot.frequency === "weekly") {
//             // Extract the day of the week (0-6) from start_date
//             const dayOfWeek = moment(datagot.start_date).day();
//             cronExpression = `${minutes} ${hours} * * ${dayOfWeek}`;
//             console.log("Cron Expression for Weekly:", cronExpression);
//         } else {
//             throw new Error("Invalid frequency. Choose 'daily' or 'weekly'.");
//         }
//         res.status(201).json({ message: "Cron data", Expression: cronExpression, data: datagot });
//     } catch (error) {
//         console.error("Error Generating Cron Expression:", error.message);
//         res.status(500).json({ message: "Internal server error", error: error.message });
//     }
// };
