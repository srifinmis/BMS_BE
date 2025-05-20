const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('roc_forms_staging', {
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false
    },
    approved_by: {
      type: DataTypes.STRING(50),
      allowNull: false
    },
    date_of_approval: {
      type: DataTypes.DATEONLY,
      allowNull: false
    },
    document_executed_date: {
      type: DataTypes.DATEONLY,
      allowNull: false
    },
    due_date_charge_creation: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    date_of_form_filed_creation: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    due_date_satisfaction: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    date_of_filing_satisfaction: {
      type: DataTypes.DATEONLY,
      allowNull: true
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
      type: DataTypes.STRING(10),
      allowNull: false
    },
    updatedby: {
      type: DataTypes.STRING(10),
      allowNull: false
    },
    remarks: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false,
      defaultValue: "Approval Pending"
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
    updated_fields: {
      type: DataTypes.JSONB,
      allowNull: true
    },
    user_type: {
      type: DataTypes.TEXT,
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'roc_forms_staging',
    schema: 'staging',
    timestamps: false,
    indexes: [
      {
        name: "roc_forms_staging_pkey",
        unique: true,
        fields: [
          { name: "id" },
        ]
      },
    ]
  });
};
