const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('interest_rate_changes', {
    change_id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    new_interest_rate: {
      type: DataTypes.DECIMAL,
      allowNull: false
    },
    effective_date: {
      type: DataTypes.DATEONLY,
      allowNull: false
    },
    createdat: {
      type: DataTypes.DATE,
      allowNull: true,
      defaultValue: Sequelize.Sequelize.literal('CURRENT_TIMESTAMP')
    },
    updatedat: {
      type: DataTypes.DATEONLY,
      allowNull: true,
      defaultValue: Sequelize.Sequelize.literal('CURRENT_TIMESTAMP')
    },
    remarks: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    tranche_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      references: {
        model: 'tranche_details',
        key: 'tranche_id'
      }
    },
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      references: {
        model: 'sanction_details',
        key: 'sanction_id'
      }
    },
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    createdby: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    updatedby: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    lender_code: {
      type: DataTypes.STRING(10),
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'interest_rate_changes',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "interest_rate_changes_pkey",
        unique: true,
        fields: [
          { name: "change_id" },
        ]
      },
    ]
  });
};
