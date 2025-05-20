const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('employee_master', {
    emp_id: {
      type: DataTypes.STRING(10),
      allowNull: false,
      primaryKey: true
    },
    emp_name: {
      type: DataTypes.STRING,
      allowNull: false
    },
    email: {
      type: DataTypes.STRING,
      allowNull: true
    },
    passwd: {
      type: DataTypes.STRING,
      allowNull: true
    },
    emp_status: {
      type: DataTypes.STRING(30),
      allowNull: true
    },
    access_status: {
      type: DataTypes.STRING(30),
      allowNull: true
    },
    designation_name: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    reports_to_id: {
      type: DataTypes.STRING(10),
      allowNull: true
    },
    department_name: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    branchid_name: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    phone: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    system_role: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    created_at: {
      type: DataTypes.DATE,
      allowNull: true,
      defaultValue: "2025-01-01 00:00:00+00"
    },
    updated_at: {
      type: DataTypes.DATE,
      allowNull: true,
      defaultValue: Sequelize.Sequelize.literal('CURRENT_TIMESTAMP')
    },
    role_id: {
      type: DataTypes.INTEGER,
      allowNull: true,
      references: {
        model: 'roles',
        key: 'role_id'
      }
    },
    role_ids_assigned: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    regionid_name: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    system_id: {
      type: DataTypes.INTEGER,
      allowNull: false,
      defaultValue: 0
    },
    areaid_name: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    clusterid_name: {
      type: DataTypes.STRING(50),
      allowNull: true
    },
    state: {
      type: DataTypes.STRING(50),
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'employee_master',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "employee_master_pkey",
        unique: true,
        fields: [
          { name: "emp_id" },
        ]
      },
    ]
  });
};
