const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('bank_account_details', {
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      primaryKey: true,
      references: {
        model: 'sanction_details',
        key: 'sanction_id'
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
    conf_acc_no: {
      type: DataTypes.STRING(50),
      allowNull: false
    },
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false,
      defaultValue: "Approval Pending"
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
    tableName: 'bank_account_details',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "bank_account_details_pkey",
        unique: true,
        fields: [
          { name: "sanction_id" },
        ]
      },
    ]
  });
};
