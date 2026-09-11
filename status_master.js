"use strict";

const STATUS_MASTER = [
  { code: "P", display_name: "Present", category: "PRESENT", is_half_day: false, hr_definition_required: false },
  { code: "A", display_name: "Absent", category: "ABSENT", is_half_day: false, hr_definition_required: false },
  { code: "WO", display_name: "Week Off", category: "WEEK_OFF", is_half_day: false, hr_definition_required: false },
  { code: "H", display_name: "Holiday", category: "HOLIDAY", is_half_day: false, hr_definition_required: false },
  { code: "CL", display_name: "Casual Leave", category: "LEAVE", is_half_day: false, hr_definition_required: false },
  { code: "EL", display_name: "Earned Leave", category: "LEAVE", is_half_day: false, hr_definition_required: false },
  { code: "WFH", display_name: "Work From Home", category: "WFH", is_half_day: false, hr_definition_required: false },
  { code: "LOP", display_name: "Loss of Pay", category: "ABSENT", is_half_day: false, hr_definition_required: false },
  { code: "AEL", display_name: "AEL", category: "OTHER", is_half_day: false, hr_definition_required: true },
  { code: "CO", display_name: "Comp Off", category: "OTHER", is_half_day: false, hr_definition_required: true },
  { code: "HP", display_name: "HP", category: "OTHER", is_half_day: false, hr_definition_required: true },
  { code: "CLP", display_name: "CLP", category: "LEAVE", is_half_day: false, hr_definition_required: true },
  { code: "WOP", display_name: "WOP", category: "OTHER", is_half_day: false, hr_definition_required: true },
  { code: "ELP", display_name: "ELP", category: "LEAVE", is_half_day: false, hr_definition_required: true },
  { code: "½CL", display_name: "Half Casual Leave", category: "LEAVE", is_half_day: true, hr_definition_required: true },
  { code: "½CLP", display_name: "Half CLP", category: "LEAVE", is_half_day: true, hr_definition_required: true },
  { code: "½EL", display_name: "Half Earned Leave", category: "LEAVE", is_half_day: true, hr_definition_required: true },
  { code: "½ELP", display_name: "Half ELP", category: "LEAVE", is_half_day: true, hr_definition_required: true },
  { code: "½WFHP", display_name: "Half WFH", category: "WFH", is_half_day: true, hr_definition_required: true },
  { code: "½AELP", display_name: "Half AELP", category: "OTHER", is_half_day: true, hr_definition_required: true },
  { code: "WFHP", display_name: "WFHP", category: "WFH", is_half_day: false, hr_definition_required: true },
];

module.exports = { STATUS_MASTER };
