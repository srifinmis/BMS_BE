const bcrypt = require("bcrypt");
const jwt = require("jsonwebtoken");
const { sequelize } = require('../config/db');
const initModels = require('../models/init-models');

const models = initModels(sequelize);
const { employee_master, roles_modules, roles, modules } = models;

module.exports.login = async (req, res) => {
  const { emp_id, passwd } = req.body;

  try {
    const user = await employee_master.findOne({ where: { emp_id } });
    if (!user) return res.status(400).json({ message: "User not found" });

    if (user.emp_status !== "ACCEPTED" || user.access_status !== "GRANTED")
      return res.status(400).json({ message: "ACCESS DENIED!" });

    const validpasswd = await bcrypt.compare(passwd, user.passwd);
    if (!validpasswd) return res.status(400).json({ message: "Invalid Credentials" });

    // Fetch role_id from employee_master
    const roleId = user.role_ids_assigned;
    if (!roleId) return res.status(404).json({ message: "User role not assigned" });
    const roleIds = user.role_ids_assigned.split(',').map(id => parseInt(id.trim(), 10));

    // Fetch role name
    const rolesData = await roles.findAll({
      where: { role_id: roleIds },
      attributes: ['role_id', 'role_name']
    });
    // console.log("Roles: ", rolesData);


    if (!rolesData || rolesData.length === 0) {
      return res.status(404).json({ message: "Roles not found" });
    }
    const roleNames = rolesData.map(role => role.role_name);

    // Fetch all modules for the role
    const roleModules = await roles_modules.findAll({
      where: { role_id: roleIds },
      include: [{
        model: modules,
        as: "module",
        attributes: ['module_id', 'module_name']
      }],
      attributes: ['role_id', 'module_id']
    });
    // console.log("roleModules: ", roleModules);

    const modulesList = [...new Set(
      roleModules.map(rm => rm.module?.module_name).filter(Boolean)
    )];
    // console.log("ModuleList: ", modulesList);


    const token = jwt.sign(
      { id: user.emp_id, empName: user.emp_name, Role: roleNames, modules: modulesList, },
      process.env.JWT_SECRET,
      { expiresIn: process.env.JWT_EXPIRES_IN }
    );
    console.log("Token: ", token)

    res.status(200).json({
      message: "Login successful",
      token,
      emp_id: user.emp_id,
      role: roleNames,
      modules: modulesList
    });
  } catch (error) {
    console.error("Login Error:", error);
    res.status(500).json({ message: "Server Error", error: error.message });
  }
};
