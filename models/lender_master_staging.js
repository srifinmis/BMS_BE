const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('lender_master_staging', {
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
    lender_name: {
      type: DataTypes.STRING(100),
      allowNull: false
    },
    lender_type: {
      type: DataTypes.STRING(100),
      allowNull: false
    },
    lender_escalation_name: {
      type: DataTypes.STRING(255),
      allowNull: true
    },
    lender_escalation_contact: {
      type: DataTypes.STRING(20),
      allowNull: true
    },
    lender_escalation_email: {
      type: DataTypes.STRING(200),
      allowNull: true
    },
    status: {
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
    approval_status: {
      type: DataTypes.STRING(30),
      allowNull: false,
      defaultValue: "Approval Pending"
    },
    updated_fields: {
      type: DataTypes.JSONB,
      allowNull: true
    },
    user_type: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    addr1_line1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_line2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_line3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_contact1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_contact2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_contact3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_email1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_email2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_email3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_spoc_name: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_spoc_contact: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr1_spoc_email: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_line1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_line2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_line3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_contact1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_contact2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_contact3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_email1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_email2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_email3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_spoc_name: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_spoc_contact: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr2_spoc_email: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_line1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_line2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_line3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_contact1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_contact2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_contact3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_email1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_email2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_email3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_spoc_name: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_spoc_contact: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr3_spoc_email: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_line1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_line2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_line3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_contact1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_contact2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_contact3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_email1: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_email2: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_email3: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_spoc_name: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_spoc_contact: {
      type: DataTypes.STRING(100),
      allowNull: true
    },
    addr4_spoc_email: {
      type: DataTypes.STRING(100),
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'lender_master_staging',
    schema: 'staging',
    timestamps: false,
    indexes: [
      {
        name: "lms_pk",
        unique: true,
        fields: [
          { name: "id" },
        ]
      },
    ]
  });
};
