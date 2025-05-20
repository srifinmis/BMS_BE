const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('payment_details', {
    payment_id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
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
    sanction_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      references: {
        model: 'sanction_details',
        key: 'sanction_id'
      }
    },
    tranche_id: {
      type: DataTypes.STRING(20),
      allowNull: false,
      references: {
        model: 'tranche_details',
        key: 'tranche_id'
      }
    },
    payment_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    utr_no: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    due_date: {
      type: DataTypes.DATEONLY,
      allowNull: true
    },
    due_amt: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    payment_upload_url: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    payment_amount: {
      type: DataTypes.BIGINT,
      allowNull: true
    },
    createdby: {
      type: DataTypes.STRING(10),
      allowNull: true
    },
    createdat: {
      type: DataTypes.DATE,
      allowNull: true
    },
    pricipal_coll: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    interest_coll: {
      type: DataTypes.INTEGER,
      allowNull: true
    },
    approval_status: {
      type: DataTypes.STRING(20),
      allowNull: true
    },
    remarks: {
      type: DataTypes.STRING(20),
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'payment_details',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "payment_details_pkey",
        unique: true,
        fields: [
          { name: "payment_id" },
        ]
      },
    ]
  });
};
