const { sequelize } = require("../config/db");
require("dotenv").config();
const cron = require('node-cron');
const nodemailer = require("nodemailer");
const dayjs = require("dayjs");


const initModels = require("../models/init-models");

const models = initModels(sequelize);
const { alert_management } = models;
// 
let activeJobs = {};

// Create transporter
const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
    },
});
// console.log("Email:", process.env.EMAIL_USER, process.env.EMAIL_PASS)

// Function to send email
const sendMail = async (to, cc, subject, text) => {
    try {
        const info = await transporter.sendMail({
            from: process.env.EMAIL_USER,
            to,
            cc,
            subject,
            text,
        });

        console.log("Email sent: " + info.response);
    } catch (error) {
        console.error("Error sending email:", error);
    }
};

// Function to start a cron job dynamically
// working code

// exports.startCronJob = async (alert_id, sanction_id, tranche_id, alert_time, alert_start_date, alert_end_date, alert_frequency, cron_expression, to_addr, cc_addr) => {
//     console.log("Data getting: ", alert_id, sanction_id, tranche_id, alert_time, alert_start_date, alert_end_date, alert_frequency, cron_expression, to_addr, cc_addr);
//     const now = new Date();


//     //

//     // Convert email addresses into arrays, trimming spaces
//     const toAddresses = to_addr ? to_addr.split(",").map(email => email.trim()) : [];
//     const ccAddresses = cc_addr ? cc_addr.split(",").map(email => email.trim()) : [];

//     const startDate = dayjs(alert_start_date);
//     const endDate = dayjs(alert_end_date);
//     const currentDate = dayjs();

//     if (startDate.isAfter(endDate)) {
//         console.log(`Skipping job for alert ID: ${alert_id}, as it has expired.`);

//         // Update alert_management table to mark as expired
//         await alert_management.update(
//             { cron_expression: null, alert_frequency: 'expired' },
//             { where: { alert_id } }
//         );

//         // Stop and remove the active job if it exists
//         if (activeJobs[alert_id]) {
//             activeJobs[alert_id].stop();
//             delete activeJobs[alert_id];
//         }

//         return;
//     }


//     if (!cron_expression || typeof cron_expression !== "string") {
//         console.error("Invalid cron expression:", cron_expression);
//         return;
//     }
//     console.log("cron :exp ", cron_expression)
//     try {
//         cron.schedule(cron_expression, async () => {
//             try {
//                 console.log("🚀 Running Scheduled Cron Job");

//                 await sendMail(
//                     toAddresses.join(","),
//                     ccAddresses.length > 0 ? ccAddresses.join(",") : "",
//                     "Test Mail from BR",
//                     `Hello Vishal, This is a test email.. Your Sanction ID: ${sanction_id}, Tranche ID: ${tranche_id} , End date of Loan : ${endDate}`
//                 );

//                 console.log("📧 Email sent successfully.");
//             } catch (error) {
//                 console.error("❌ Error in cron job execution:", error.message);
//             }
//         }, {
//             timezone: "Asia/Kolkata", // Ensure correct timezone
//         });

//         console.log("✅ Cron job scheduled successfully.");
//     } catch (error) {
//         console.error("❌ Error scheduling cron job:", error.message);
//     }

//     // Schedule job only if alert is active
//     if (cron_expression != null) {
//         const job = cron.schedule(cron_expression, async () => {
//             if (new Date() >= new Date(alert_end_date)) {
//                 // console.log(`Stopping job for alert ID: ${alert_id}, as it reached end date.`);
//                 job.stop();
//                 await alert_management.update({ cron_expression: null, alert_frequency: 'expired' }, { where: { alert_id } });
//                 delete activeJobs[alert_id];
//                 return;
//             }
//             console.log(`Executing alert ID: ${alert_id}, Frequency: ${alert_frequency}`);
//         });

//         activeJobs[alert_id] = job;
//     }


//     console.log(`Cron job started for alert ID: ${alert_id} with expression: ${cron_expression}`);
// }
// 



exports.cronalert = async (req, res) => {

    try {
        console.log("Received Alert Data:", req.body);
        const { sanction_id } = req.body;
        const alert = await alert_management.findAll({ where: { sanction_id } });
        console.log("Alert Management Data: ", alert)

        const task = cron.schedule('* * * * *', () => {
            const now = new Date();
            console.log(`Cron running at: ${now}`);

            // Stop the cron job on 21st March 2025 at 6:42 PM
            const stopDate = new Date('2025-03-21T18:42:00');

            if (now >= stopDate) {
                console.log('Stopping the cron job at 6:42 PM on 21st March 2025...');
                task.stop();
            }
        });

        task.start();
        console.log('Cron job started...');

        res.status(201).json({ message: " Alert Management Data Fetched successfully", gotdata: alert });
    } catch (error) {
        console.error("Login Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }

}

exports.startCronJob = async (alert_id, sanction_id, tranche_id, alert_time, alert_start_date, alert_end_date, alert_frequency, cron_expression, to_addr, cc_addr) => {
    console.log("Data getting: ", alert_id, sanction_id, tranche_id, alert_time, alert_start_date, alert_end_date, alert_frequency, cron_expression, to_addr, cc_addr);

    const startDate = dayjs(alert_start_date);
    const endDate = dayjs(alert_end_date);
    const currentDate = dayjs();

    // Convert email addresses into arrays, trimming spaces
    const toAddresses = to_addr ? to_addr.split(",").map(email => email.trim()) : [];
    const ccAddresses = cc_addr ? cc_addr.split(",").map(email => email.trim()) : [];

    if (startDate.isAfter(endDate)) {
        console.log(`Skipping job for alert ID: ${alert_id}, as it has expired.`);

        // Update alert_management table to mark as expired
        await alert_management.update(
            { cron_expression: null, alert_frequency: 'expired' },
            { where: { alert_id } }
        );

        // Stop and remove the active job if it exists
        if (activeJobs[alert_id]) {
            activeJobs[alert_id].stop();
            delete activeJobs[alert_id];
        }

        return;
    }

    if (!cron_expression || typeof cron_expression !== "string") {
        console.error("Invalid cron expression:", cron_expression);
        return;
    }

    try {
        // Ensure only one cron job is scheduled
        if (!activeJobs[alert_id]) {
            const job = cron.schedule(cron_expression, async () => {
                // Check if the current date exceeds the end date, if so, stop the job
                if (new Date() >= new Date(alert_end_date)) {
                    console.log(`Stopping job for alert ID: ${alert_id}, as it reached end date.`);
                    job.stop();
                    await alert_management.update(
                        { cron_expression: null, alert_frequency: 'expired' },
                        { where: { alert_id } }
                    );
                    delete activeJobs[alert_id];
                    return;
                }

                console.log(`Executing alert ID: ${alert_id}, Frequency: ${alert_frequency}`);

                try {
                    // Send an email if the cron job executes
                    await sendMail(
                        toAddresses.join(","),
                        ccAddresses.length > 0 ? ccAddresses.join(",") : "",
                        "Test Mail from BR",
                        `Hello Vishal, This is a test email.. Your Sanction ID: ${sanction_id}, Tranche ID: ${tranche_id} , End date of Loan : ${endDate}`
                    );
                    console.log("📧 Email sent successfully.");
                } catch (error) {
                    console.error("❌ Error in sending email:", error.message);
                }
            }, {
                timezone: "Asia/Kolkata", // Ensure correct timezone
            });

            // Store the job in activeJobs to manage it
            activeJobs[alert_id] = job;
            console.log("✅ Cron job scheduled successfully.");
        } else {
            console.log(`Cron job for alert ID: ${alert_id} is already active.`);
        }
    } catch (error) {
        console.error("❌ Error scheduling cron job:", error.message);
    }

    console.log(`Cron job started for alert ID: ${alert_id} with expression: ${cron_expression}`);
};
