const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('alert_management_staging', {
    alert_id: {
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      references: {
        model: 'sanction_details',
        key: 'sanction_id'
      }
    },
    tranche_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      references: {
        model: 'tranche_details',
        key: 'tranche_id'
      }
    },
    alert_time: {
      type: DataTypes.TIME,
      allowNull: true
    },
    alert_start_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    alert_end_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    alert_frequency: {
      type: DataTypes.STRING(30),
      allowNull: true
    },
    cron_expression: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false,
      defaultValue: "Approval Pending"
    },
    to_addr: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    cc_addr: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    lender_code: {
      type: DataTypes.STRING(10),
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'alert_management_staging',
    schema: 'staging',
    timestamps: false,
    indexes: [
      {
        name: "alert_management_staging_pkey",
        unique: true,
        fields: [
          { name: "alert_id" },
        ]
      },
      {
        name: "s_alert_management_composite_idx",
        fields: [
          { name: "sanction_id" },
          { name: "tranche_id" },
          { name: "alert_start_date" },
          { name: "alert_end_date" },
        ]
      },
    ]
  });
};
