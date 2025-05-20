const Sequelize = require('sequelize');
module.exports = function(sequelize, DataTypes) {
  return sequelize.define('projects_test', {
    id: {
      autoIncrement: true,
      type: DataTypes.INTEGER,
      allowNull: false,
      primaryKey: true
    },
    name: {
      type: DataTypes.TEXT,
      allowNull: true
    },
    team_members: {
      type: DataTypes.JSONB,
      allowNull: true
    }
  }, {
    sequelize,
    tableName: 'projects_test',
    schema: 'public',
    timestamps: false,
    indexes: [
      {
        name: "projects_test_pkey",
        unique: true,
        fields: [
          { name: "id" },
        ]
      },
    ]
  });
};
