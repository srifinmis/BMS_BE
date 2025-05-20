const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const helmet = require('helmet');
const xss = require('xss-clean');
require("dotenv").config();

const { sequelize } = require("./config/db");
const cron = require('node-cron');
const initModels = require("./models/init-models");

const models = initModels(sequelize);
const { alert_management } = models;
const { startCronJob } = require("./Controller/alertTriggerController");

// const axios = require('axios');
const Router = require("./Routes/Router")
const excelRoutes = require('./Routes/tranchExcelRoutes');

const app = express();

app.use(helmet()); //Secure HTTP headers
app.use(xss());  //Prevent XSS attack
app.use(cors());
app.use(express.json());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

const API_URL = process.env.REACT_APP_API_URL;
const PORT = process.env.PORT;
app.use("/api", Router);
app.use('/api/excel', excelRoutes);
// **Auto-Start Jobs on Server Restart**
const loadCronJobs = async () => {
    const rows = await alert_management.findAll({
        attributes: [
            "alert_id", "sanction_id", "tranche_id", "alert_time", "alert_start_date", "alert_end_date", "alert_frequency", "cron_expression", "to_addr", "cc_addr"
        ],
    });
    // console.log('loading rows: ', rows);
    rows.forEach(({ alert_id, sanction_id, tranche_id, alert_time, alert_start_date, alert_end_date, alert_frequency, cron_expression, to_addr, cc_addr }) => {
        // console.log('loading Data: ', alert_id, alert_end_date, cron_expression);
        startCronJob(alert_id, sanction_id, tranche_id, alert_time, alert_start_date, alert_end_date, alert_frequency, cron_expression, to_addr, cc_addr);
    });
};

app.listen(PORT, async () => {
    console.log(`Server running on http://localhost:${PORT}`);
    // await loadCronJobs();
});

// **Schedule daily database load at midnight**
// cron.schedule('0 * * * * *', async () => {
//     console.log("Running daily database load at midnight...");
//     await loadCronJobs();
// });