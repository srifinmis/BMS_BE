var DataTypes = require("sequelize").DataTypes;
var _alert_management = require("./alert_management");
var _alert_management_staging = require("./alert_management_staging");
var _bank_account_details = require("./bank_account_details");
var _bank_account_details_staging = require("./bank_account_details_staging");
var _employee_master = require("./employee_master");
var _employee_master_temp = require("./employee_master_temp");
var _executed_documents = require("./executed_documents");
var _executed_documents_staging = require("./executed_documents_staging");
var _interest_rate_changes = require("./interest_rate_changes");
var _interest_rate_changes_staging = require("./interest_rate_changes_staging");
var _lender_master = require("./lender_master");
var _lender_master_staging = require("./lender_master_staging");
var _modules = require("./modules");
var _payment_details = require("./payment_details");
var _payment_details_staging = require("./payment_details_staging");
var _projects_test = require("./projects_test");
var _repayment_schedule = require("./repayment_schedule");
var _repayment_schedule_staging = require("./repayment_schedule_staging");
var _repayment_schedule_temp = require("./repayment_schedule_temp");
var _roc_forms = require("./roc_forms");
var _roc_forms_staging = require("./roc_forms_staging");
var _roles = require("./roles");
var _roles_modules = require("./roles_modules");
var _sanction_details = require("./sanction_details");
var _sanction_details_staging = require("./sanction_details_staging");
var _tranche_details = require("./tranche_details");
var _tranche_details_staging = require("./tranche_details_staging");

function initModels(sequelize) {
  var alert_management = _alert_management(sequelize, DataTypes);
  var alert_management_staging = _alert_management_staging(sequelize, DataTypes);
  var bank_account_details = _bank_account_details(sequelize, DataTypes);
  var bank_account_details_staging = _bank_account_details_staging(sequelize, DataTypes);
  var employee_master = _employee_master(sequelize, DataTypes);
  var employee_master_temp = _employee_master_temp(sequelize, DataTypes);
  var executed_documents = _executed_documents(sequelize, DataTypes);
  var executed_documents_staging = _executed_documents_staging(sequelize, DataTypes);
  var interest_rate_changes = _interest_rate_changes(sequelize, DataTypes);
  var interest_rate_changes_staging = _interest_rate_changes_staging(sequelize, DataTypes);
  var lender_master = _lender_master(sequelize, DataTypes);
  var lender_master_staging = _lender_master_staging(sequelize, DataTypes);
  var modules = _modules(sequelize, DataTypes);
  var payment_details = _payment_details(sequelize, DataTypes);
  var payment_details_staging = _payment_details_staging(sequelize, DataTypes);
  var projects_test = _projects_test(sequelize, DataTypes);
  var repayment_schedule = _repayment_schedule(sequelize, DataTypes);
  var repayment_schedule_staging = _repayment_schedule_staging(sequelize, DataTypes);
  var repayment_schedule_temp = _repayment_schedule_temp(sequelize, DataTypes);
  var roc_forms = _roc_forms(sequelize, DataTypes);
  var roc_forms_staging = _roc_forms_staging(sequelize, DataTypes);
  var roles = _roles(sequelize, DataTypes);
  var roles_modules = _roles_modules(sequelize, DataTypes);
  var sanction_details = _sanction_details(sequelize, DataTypes);
  var sanction_details_staging = _sanction_details_staging(sequelize, DataTypes);
  var tranche_details = _tranche_details(sequelize, DataTypes);
  var tranche_details_staging = _tranche_details_staging(sequelize, DataTypes);

  modules.belongsToMany(roles, { as: 'role_id_roles', through: roles_modules, foreignKey: "module_id", otherKey: "role_id" });
  roles.belongsToMany(modules, { as: 'module_id_modules', through: roles_modules, foreignKey: "role_id", otherKey: "module_id" });
  alert_management.belongsTo(lender_master, { as: "lender_code_lender_master", foreignKey: "lender_code"});
  lender_master.hasMany(alert_management, { as: "alert_managements", foreignKey: "lender_code"});
  executed_documents.belongsTo(lender_master, { as: "lender_code_lender_master", foreignKey: "lender_code"});
  lender_master.hasMany(executed_documents, { as: "executed_documents", foreignKey: "lender_code"});
  payment_details.belongsTo(lender_master, { as: "lender_code_lender_master", foreignKey: "lender_code"});
  lender_master.hasMany(payment_details, { as: "payment_details", foreignKey: "lender_code"});
  repayment_schedule.belongsTo(lender_master, { as: "lender_code_lender_master", foreignKey: "lender_code"});
  lender_master.hasMany(repayment_schedule, { as: "repayment_schedules", foreignKey: "lender_code"});
  roc_forms.belongsTo(lender_master, { as: "lender_code_lender_master", foreignKey: "lender_code"});
  lender_master.hasMany(roc_forms, { as: "roc_forms", foreignKey: "lender_code"});
  sanction_details.belongsTo(lender_master, { as: "lender_code_lender_master", foreignKey: "lender_code"});
  lender_master.hasMany(sanction_details, { as: "sanction_details", foreignKey: "lender_code"});
  tranche_details.belongsTo(lender_master, { as: "lender_code_lender_master", foreignKey: "lender_code"});
  lender_master.hasMany(tranche_details, { as: "tranche_details", foreignKey: "lender_code"});
  roles_modules.belongsTo(modules, { as: "module", foreignKey: "module_id"});
  modules.hasMany(roles_modules, { as: "roles_modules", foreignKey: "module_id"});
  employee_master.belongsTo(roles, { as: "role", foreignKey: "role_id"});
  roles.hasMany(employee_master, { as: "employee_masters", foreignKey: "role_id"});
  employee_master_temp.belongsTo(roles, { as: "role", foreignKey: "role_id"});
  roles.hasMany(employee_master_temp, { as: "employee_master_temps", foreignKey: "role_id"});
  roles_modules.belongsTo(roles, { as: "role", foreignKey: "role_id"});
  roles.hasMany(roles_modules, { as: "roles_modules", foreignKey: "role_id"});
  alert_management.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasMany(alert_management, { as: "alert_managements", foreignKey: "sanction_id"});
  bank_account_details.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasOne(bank_account_details, { as: "bank_account_detail", foreignKey: "sanction_id"});
  executed_documents.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasMany(executed_documents, { as: "executed_documents", foreignKey: "sanction_id"});
  interest_rate_changes.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasMany(interest_rate_changes, { as: "interest_rate_changes", foreignKey: "sanction_id"});
  payment_details.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasMany(payment_details, { as: "payment_details", foreignKey: "sanction_id"});
  repayment_schedule.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasMany(repayment_schedule, { as: "repayment_schedules", foreignKey: "sanction_id"});
  roc_forms.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasOne(roc_forms, { as: "roc_form", foreignKey: "sanction_id"});
  tranche_details.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasMany(tranche_details, { as: "tranche_details", foreignKey: "sanction_id"});
  alert_management.belongsTo(tranche_details, { as: "tranche", foreignKey: "tranche_id"});
  tranche_details.hasMany(alert_management, { as: "alert_managements", foreignKey: "tranche_id"});
  interest_rate_changes.belongsTo(tranche_details, { as: "tranche", foreignKey: "tranche_id"});
  tranche_details.hasMany(interest_rate_changes, { as: "interest_rate_changes", foreignKey: "tranche_id"});
  payment_details.belongsTo(tranche_details, { as: "tranche", foreignKey: "tranche_id"});
  tranche_details.hasMany(payment_details, { as: "payment_details", foreignKey: "tranche_id"});
  repayment_schedule.belongsTo(tranche_details, { as: "tranche", foreignKey: "tranche_id"});
  tranche_details.hasMany(repayment_schedule, { as: "repayment_schedules", foreignKey: "tranche_id"});
  alert_management_staging.belongsTo(sanction_details, { as: "sanction", foreignKey: "sanction_id"});
  sanction_details.hasMany(alert_management_staging, { as: "alert_management_stagings", foreignKey: "sanction_id"});
  alert_management_staging.belongsTo(tranche_details, { as: "tranche", foreignKey: "tranche_id"});
  tranche_details.hasMany(alert_management_staging, { as: "alert_management_stagings", foreignKey: "tranche_id"});

  return {
    alert_management,
    alert_management_staging,
    bank_account_details,
    bank_account_details_staging,
    employee_master,
    employee_master_temp,
    executed_documents,
    executed_documents_staging,
    interest_rate_changes,
    interest_rate_changes_staging,
    lender_master,
    lender_master_staging,
    modules,
    payment_details,
    payment_details_staging,
    projects_test,
    repayment_schedule,
    repayment_schedule_staging,
    repayment_schedule_temp,
    roc_forms,
    roc_forms_staging,
    roles,
    roles_modules,
    sanction_details,
    sanction_details_staging,
    tranche_details,
    tranche_details_staging,
  };
}
module.exports = initModels;
module.exports.initModels = initModels;
module.exports.default = initModels;
