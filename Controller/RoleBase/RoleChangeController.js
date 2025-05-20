const bcrypt = require("bcrypt");
const jwt = require("jsonwebtoken");
const { Op } = require("sequelize");
const { sequelize } = require('../../config/db');
const initModels = require('../../models/init-models');

const models = initModels(sequelize);
const { employee_master, roles } = models;

module.exports.fetchEmployee = async (req, res) => {
    const datagot = req.body;
    // console.log('Data from FD:  ', datagot);
    try {

        const employeedata = await employee_master.findAll({
            attributes: ["emp_id", "emp_name", "role_ids_assigned"],
            where: {
                [Op.and]: [
                    {
                        passwd: {
                            [Op.ne]: null, // Exclude null values
                        },
                    },
                    {
                        passwd: {
                            [Op.ne]: "", // Exclude empty strings
                        },
                    },
                ],
            },
            include: [
                {
                    model: roles,
                    as: 'role',
                    attributes: ["role_name"],
                },
            ],
        });

        // console.log("employee details: ", employeedata)

        if (!employeedata || employeedata.length === 0) {
            return res.status(404).json({ message: "No Pending Tranches found" });
        }

        // const Lenderget = await lender_master
        res.status(201).json({ message: "Employee Fetch successfully", data: employeedata });
    } catch (error) {
        console.error("Employee Data Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }

}

module.exports.fetchRoles = async (req, res) => {
    try {

        const rolesdata = await roles.findAll({
            attributes: ["role_id", "role_name"],
        });
        // console.log("employee details: ", employeedata)

        if (!rolesdata || rolesdata.length === 0) {
            return res.status(404).json({ message: "No Pending Tranches found" });
        }
        res.status(201).json({ message: "Roles", data: rolesdata });
    } catch (error) {
        console.error("Employee Data Error:", error);
        res.status(500).json({ message: "Internal server error", error: error.message });
    }
}
module.exports.roleChange = async (req, res) => {
    const data = req.body;
    // console.log("data to role Change: ", data);

    // Role to role_id mapping
    const roleMapping = {
        "ADMIN": 1,
        "FINANCE MANAGER": 2,
        "TREASURY OFFICER": 3,
        "AUDITOR": 4,
        "HO USER": 5,
        "IT-ADMIN": 6
    };

    try {
        const employee = await employee_master.findOne({
            where: { emp_id: data.emp_id }
        });

        if (!employee) {
            return res.status(404).json({ message: "Employee not found" });
        }

        // Convert role names to role_ids
        const roleIds = data.roles.map(role => roleMapping[role]).filter(id => id !== undefined);

        if (roleIds.length === 0) {
            return res.status(400).json({ message: "Invalid roles provided" });
        }

        const roleIdsString = roleIds.join(',');

        // Update employee_master with role_ids_assigned as string
        await employee_master.update(
            { role_ids_assigned: roleIdsString },
            { where: { emp_id: data.emp_id } }
        );
        // console.log("sending to db: ", res)

        return res.status(200).json({ message: "User roles updated successfully" });

    } catch (error) {
        console.error("Employee Role Update Error:", error);
        return res.status(500).json({ message: "Internal server error", error: error.message });
    }
}
