const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('alert_management', {
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
    to_addr: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    cc_addr: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    alert_id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    lender_code: {
      type: DataTypes.STRING(10),
      allowNull: true,
      references: {
        model: 'lender_master',
        key: 'lender_code'
      }
    }
  }, {
    sequelize,
    tableName: 'alert_management',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "alert_management_composite_idx",
        fields: [
          { name: "sanction_id" },
          { name: "tranche_id" },
          { name: "alert_start_date" },
          { name: "alert_end_date" },
        ]
      },
      {
        name: "alert_management_pkey",
        unique: true,
        fields: [
          { name: "alert_id" },
        ]
      },
    ]
  });
};
