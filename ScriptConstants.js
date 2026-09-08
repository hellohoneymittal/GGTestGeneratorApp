//api constant

const APPLICATION_URL =
  "https://script.google.com/macros/s/AKfycbztRT61BtnaLpEOcE9nghNdMz5rm-l8Oi8xbNhEO50V9TvrZBNtR_JE9h6hJxSqurx-Dw/exec";
const IMAGE_CONSTANT = {
  clickHere: "https://i.postimg.cc/g0LSdBpL/Click-Here.jpg",
  addUserIcon: "https://imghost.net/ib/E5PegaLvH4xfUED_1729512954.png",
  confirmationIcon: "https://i.ibb.co/BsvQsfb/Confirmation-icon.png",
  deleteIcon: "https://i.postimg.cc/cJZRzYzT/delete-Icon.png",
};
const API_TYPE_CONSTANT = {
  TESTING_DOCS: "TESTING_DOCS",
  CHECK_PASSWORD: "CHECK_PASSWORD",
  GET_CLASS_SUBJECT_LIST: "GET_CLASS_SUBJECT_LIST",
  GET_LIST_BY_CLASS_SUBJECT: "GET_LIST_BY_CLASS_SUBJECT",
  INSERT_CHAPTER_DATA: "INSERT_CHAPTER_DATA",
};
const DATE_FORMAT_CONSTANT = {
  grid: "DD MMM YYYY",
  database: "yyyy-MM-dd",
  gridWithDate: "DD MMM YYYY hh:mm A",
};

const PASSWORD_ERROR_STR = "Please enter a correct password";
const DATE_UTC = new Date().toISOString();

const CONTROL_TYPE_CONSTAINT = {
  input: "input",
  button: "button",
  checkbox: "checkbox",
};

//page constant
const PASSWORD_CONTAINER = "passwordContainer";
const HM_CONTANER = "homeContainer";
const DM_CONTAINER = "donorMasterContainer";

const bheeshmUserNameLSKey = "bheeshmUserName";
const bheeshmUserFacilitatorLSKey = "bheeshmUserFacilitator";

const POPUP_CONSTANT = {
  error: "errorPopup",
  success: "successPopup",
};

const ICON_CONSTANT = {
  downloadIcon: "https://cdn-thumbs.imagevenue.com/85/09/8b/ME196HF8_t.png",
};

const ROLE_CONSTANT = {
  admin: "Admin",
  superAdmin: "Super Admin",
};

const ERROR_MESSAGE_CONSTANT = {
  general: "Something Went Wrong",
};

function getFormattedDateForDownload() {
  const today = new Date();
  const day = today.getDate();
  const monthNames = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  const month = monthNames[today.getMonth()];
  return `${day}${daySuffix(day)}${month}`;
}

const ExcelDate = getFormattedDateForDownload();

const VALIDATION_CONSTANT = {
  numberWithDecimal: "^d*.?d*$",
};
