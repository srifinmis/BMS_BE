const jwt = require("jsonwebtoken");
const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const { configDotenv } = require("dotenv");
const models = initModels(sequelize);
const { repayment_schedule_temp, repayment_schedule_staging } = models;

exports.uploadScheduleFile = async (req, res) => {
  try {
    console.log("uploadScheduleFile : ", req.body);

    const JWT_SECRET = process.env.JWT_SECRET;
    if (!JWT_SECRET) {
      return res.status(500).json({ message: "JWT secret is not configured" });
    }

    const token = req.body.finalFormData?.createdby;
    const decoded = jwt.verify(token, JWT_SECRET);

    if (!decoded?.id) {
      return res.status(401).json({ message: "Invalid JWT token" });
    }

    await sequelize.transaction(async (t) => {
      const { tranche_id, loan_type } = req.body.finalFormData;

      const tempRecords = await repayment_schedule_temp.findAll({
        where: { tranche_id, repayment_type: loan_type },
        transaction: t,
      });

      if (tempRecords.length === 0) {
        return res.status(400).json({ message: "No data found in temp table" });
      }

      const plainRecords = tempRecords.map((record) => ({
        ...record.get({ plain: true }),
        approval_status: "Approval Pending",
        user_type: "N",
        createdby: decoded.id || "SFTADM",
        createdAt: new Date(),
        updatedAt: new Date(),
      }));
      console.log("heloooooooooooooo"+ JSON.stringify(req.body.finalFormData.flag));
      if (req.body.finalFormData.flag) {
        
        for (const record of plainRecords) {
          const [existingRecord, created] = await repayment_schedule_staging.findOrCreate({
            where: {
              lender_code: record.lender_code,
              sanction_id: record.sanction_id,
              tranche_id: record.tranche_id,
              due_date: record.due_date
            },
            defaults: record,
            transaction: t,
          });

          if (!created) {
            await existingRecord.update(record, { transaction: t });
          }
        }
      } else {
        await repayment_schedule_staging.bulkCreate(plainRecords, { transaction: t });
      }

      await repayment_schedule_temp.destroy({
        where: { tranche_id, repayment_type: loan_type },
        transaction: t,
      });
    });

    res.json({ message: "Data moved successfully" });

  } catch (error) {
    console.error("Error moving data:", error);
    res.status(500).json({ message: "Failed to move data", error: error.message });
  }
};
