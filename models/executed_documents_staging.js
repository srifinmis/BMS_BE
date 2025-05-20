const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('executed_documents_staging', {
    document_id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false
    },
    document_type: {
      type: DataTypes.TEXT,
      allowNull: false
    },
    file_name: {
      type: DataTypes.TEXT,
      allowNull: false
    },
    uploaded_date: {
      type: DataTypes.DATE,
      allowNull: false
    },
    document_url: {
      type: DataTypes.TEXT,
      allowNull: false
    },
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false,
      defaultValue: "Approval Pending"
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
    tableName: 'executed_documents_staging',
    schema: 'staging',
    timestamps: false,
    indexes: [
      {
        name: "executed_documents_staging_pkey",
        unique: true,
        fields: [
          { name: "document_id" },
        ]
      },
    ]
  });
};
