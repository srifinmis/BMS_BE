const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('repayment_schedule_staging', {
    repayment_id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      unique: "unique_sanction_tranche"
    },
    tranche_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      unique: "unique_sanction_tranche"
    },
    due_date: {
      type: DataTypes.DATEONLY,
      allowNull: true,
      unique: "unique_sanction_tranche"
    },
    principal_due: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    interest_due: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    total_due: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    createdby: {
      type: DataTypes.STRING(50),
      allowNull: false
    },
    createdat: {
      type: DataTypes.DATE,
      allowNull: true,
      defaultValue: Sequelize.Sequelize.literal('CURRENT_TIMESTAMP')
    },
    remarks: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    lender_code: {
      type: DataTypes.STRING(10),
      allowNull: true
    },
    updated_fields: {
      type: DataTypes.JSONB,
      allowNull: true
    },
    user_type: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    opening_balance: {
      type: DataTypes.BIGINT,
      allowNull: true
    },
    closing_balance: {
      type: DataTypes.BIGINT,
      allowNull: true
    },
    from_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    interest_days: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    interest_rate: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    emi_sequence: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    repayment_type: {
      type: DataTypes.STRING(200),
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'repayment_schedule_staging',
    schema: 'staging',
    timestamps: false,
    indexes: [
      {
        name: "repayment_schedule_staging_pkey",
        unique: true,
        fields: [
          { name: "repayment_id" },
        ]
      },
      {
        name: "unique_sanction_tranche",
        unique: true,
        fields: [
          { name: "sanction_id" },
          { name: "tranche_id" },
          { name: "due_date" },
        ]
      },
    ]
  });
};
