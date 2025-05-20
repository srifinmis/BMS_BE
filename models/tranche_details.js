const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('tranche_details', {
    tranche_id: {
      type: DataTypes.STRING(20),
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
    tranche_number: {
      type: DataTypes.INTEGER,
      allowNull: false
    },
    tranche_amount: {
      type: DataTypes.DECIMAL,
      allowNull: false
    },
    interest_type: {
      type: DataTypes.STRING(20),
      allowNull: false
    },
    interest_rate: {
      type: DataTypes.DECIMAL,
      allowNull: false
    },
    tenure_months: {
      type: DataTypes.INTEGER,
      allowNull: false
    },
    principal_start_date: {
      type: DataTypes.DATEONLY,
      allowNull: false
    },
    interest_start_date: {
      type: DataTypes.DATEONLY,
      allowNull: false
    },
    principal_payment_frequency: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    interest_payment_frequency: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    applicable_of_leap_year: {
      type: DataTypes.STRING(5),
      allowNull: false
    },
    interest_calculation_days: {
      type: DataTypes.INTEGER,
      allowNull: false
    },
    createdat: {
      type: DataTypes.DATE,
      allowNull: true,
      defaultValue: Sequelize.Sequelize.literal('CURRENT_TIMESTAMP')
    },
    updatedat: {
      type: DataTypes.DATE,
      allowNull: true,
      defaultValue: Sequelize.Sequelize.literal('CURRENT_TIMESTAMP')
    },
    remarks: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    createdby: {
      type: DataTypes.STRING(10),
      allowNull: false
    },
    updatedby: {
      type: DataTypes.STRING(10),
      allowNull: false
    },
    lender_code: {
      type: DataTypes.STRING(10),
      allowNull: true,
      references: {
        model: 'lender_master',
        key: 'lender_code'
      }
    },
    current_ac_no: {
      type: DataTypes.STRING(50),
      allowNull: false
    },
    bank_name: {
      type: DataTypes.TEXT,
      allowNull: false
    },
    bank_branch: {
      type: DataTypes.TEXT,
      allowNull: false
    },
    location: {
      type: DataTypes.TEXT,
      allowNull: false
    },
    ifsc_code: {
      type: DataTypes.STRING(20),
      allowNull: false
    },
    conf_acc_no: {
      type: DataTypes.STRING(50),
      allowNull: false
    },
    tranche_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false
    },
    moratorium_start_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    moratorium_end_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'tranche_details',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "tranche_details_composite_idx",
        fields: [
          { name: "tranche_id" },
          { name: "sanction_id" },
          { name: "approval_status" },
          { name: "tranche_number" },
        ]
      },
      {
        name: "tranche_details_pkey",
        unique: true,
        fields: [
          { name: "tranche_id" },
        ]
      },
    ]
  });
};
