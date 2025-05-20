const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('interest_rate_changes_staging', {
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
    tranche_id: {
      type: DataTypes.STRING(20),
      allowNull: false
    },
    sanction_id: {
      type: DataTypes.STRING(20),
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
    createdby: {
      type: DataTypes.STRING(30),
      allowNull: false
    },
    updatedby: {
      type: DataTypes.STRING(30),
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
    updated_fields: {
      type: DataTypes.JSONB,
      allowNull: true
    },
    user_type: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    lender_code: {
      type: DataTypes.STRING(10),
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'interest_rate_changes_staging',
    schema: 'staging',
    timestamps: false,
    indexes: [
      {
        name: "interest_rate_changes_staging_pkey",
        unique: true,
        fields: [
          { name: "change_id" },
        ]
      },
    ]
  });
};
