const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('repayment_schedule_temp', {
    repayment_id: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: true
    },
    tranche_id: {
      type: DataTypes.STRING(20),
      allowNull: true
    },
    due_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
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
    createdat: {
      type: DataTypes.DATE,
      allowNull: true
    },
    lender_code: {
      type: DataTypes.STRING(10),
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
    },
    id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    created_by: {
      type: DataTypes.STRING(10),
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
    tableName: 'repayment_schedule_temp',
    schema: 'temp',
    timestamps: false,
    indexes: [
      {
        name: "repayment_schedule_temp_pkey",
        unique: true,
        fields: [
          { name: "id" },
        ]
      },
    ]
  });
};
