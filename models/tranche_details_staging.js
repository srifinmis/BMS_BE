const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('tranche_details_staging', {
    tranche_id: {
      type: DataTypes.STRING(20),
      allowNull: false
    },
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false
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
    createdby: {
      type: DataTypes.STRING(20),
      allowNull: false
    },
    updatedby: {
      type: DataTypes.STRING(20),
      allowNull: false
    },
    remarks: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    lender_code: {
      type: DataTypes.STRING(10),
      allowNull: true
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
    user_type: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    updated_fields: {
      type: DataTypes.JSONB,
      allowNull: true
    },
    tranche_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
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
    tableName: 'tranche_details_staging',
    schema: 'staging',
    timestamps: false,
    indexes: [
      {
        name: "tranche_details_staging_pkey",
        unique: true,
        fields: [
          { name: "id" },
        ]
      },
    ]
  });
};
