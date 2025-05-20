const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('roc_forms', {
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      primaryKey: true,
      references: {
        model: 'sanction_details',
        key: 'sanction_id'
      }
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
    }
  }, {
    sequelize,
    tableName: 'roc_forms',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "roc_forms_pkey",
        unique: true,
        fields: [
          { name: "sanction_id" },
        ]
      },
    ]
  });
};
