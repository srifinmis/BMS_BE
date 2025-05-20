const { sequelize } = require("../config/db");
const initModels = require("../models/init-models");
const models = initModels(sequelize);
const { repayment_schedule_temp, repayment_schedule_staging, tranche_details } = models;

const { Op } = require("sequelize");

exports.tranche_sum = async (req, res) => {
  try {
    const { sanction_id, lender_code, tranche_id, flag } = req.body;

    if (!sanction_id || !lender_code) {
      return res.status(400).json({ error: "Missing sanction_id or lender_code" });
    }

    let whereClause = {
      sanction_id,
      lender_code
    };

    if (flag && tranche_id) {
      // Exclude the given tranche_id
      whereClause.tranche_id = { [Op.ne]: tranche_id };
    }

    const result = await tranche_details.findOne({
      attributes: [
        [sequelize.fn("SUM", sequelize.col("tranche_amount")), "total_tranche_amount"]
      ],
      where: whereClause,
      raw: true
    });

    return res.status(200).json({
      success: true,
      sanction_id,
      lender_code,
      total_tranche_amount: result.total_tranche_amount || 0
    });

  } catch (e) {
    console.error("Error in tranche_sum:", e);
    return res.status(500).json({ error: "Internal server error" });
  }
};
