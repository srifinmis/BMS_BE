const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('executed_documents', {
    document_id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      references: {
        model: 'sanction_details',
        key: 'sanction_id'
      }
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
      type: DataTypes.STRING(50),
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
    tableName: 'executed_documents',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "executed_documents_pkey",
        unique: true,
        fields: [
          { name: "document_id" },
        ]
      },
    ]
  });
};
