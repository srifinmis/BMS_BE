const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('sanction_details', {
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      primaryKey: true
    },
    lender_code: {
      type: DataTypes.STRING(10),
      allowNull: false,
      references: {
        model: 'lender_master',
        key: 'lender_code'
      }
    },
    loan_type: {
      type: DataTypes.STRING(200),
      allowNull: false
    },
    purpose_of_loan: {
      type: DataTypes.TEXT,
      allowNull: false
    },
    interest_type: {
      type: DataTypes.STRING(100),
      allowNull: false
    },
    interest_rate_fixed: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    benchmark_rate: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    spread_floating: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    tenure_months: {
      type: DataTypes.INTEGER,
      allowNull: false
    },
    principal_repayment_term: {
      type: DataTypes.STRING(100),
      allowNull: false
    },
    interest_payment_term: {
      type: DataTypes.STRING(100),
      allowNull: false
    },
    sanction_validity: {
      type: DataTypes.DATEONLY,
      allowNull: false
    },
    sanction_amount: {
      type: DataTypes.DECIMAL,
      allowNull: false
    },
    processing_fee: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    other_expenses: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    book_debt_margin: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    cash_margin: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    prepayment_charges: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    corporate_guarantee: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    penal_charges: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    syndication_fee: {
      type: DataTypes.DECIMAL,
      allowNull: true
    },
    syndicated_by: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    sanction_date: {
      type: DataTypes.DATEONLY,
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
    }
  }, {
    sequelize,
    tableName: 'sanction_details',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "sanction_details_composite_idx",
        fields: [
          { name: "sanction_id" },
          { name: "lender_code" },
          { name: "approval_status" },
        ]
      },
      {
        name: "sanction_details_pkey",
        unique: true,
        fields: [
          { name: "sanction_id" },
        ]
      },
    ]
  });
};
