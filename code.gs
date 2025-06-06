// ========================================================================
// HẰNG SỐ CẤU HÌNH
// ========================================================================
const SPREADSHEET_NAME = "GHÉP XE";
const SINGLE_DATA_SHEET_NAME = "Data";
const USER_SHEET_NAME = "Users";
const ACTUAL_INVENTORY_SHEET_NAME = "khoxethucte";

const USER_SHEET_HEADERS = ["Username", "Password", "Role"];
const ACTUAL_VIN_COLUMN_LETTER = "F";

// BẠN CẦN CẬP NHẬT ID CÁC THƯ MỤC GOOGLE DRIVE CỦA BẠN VÀO ĐÂY
const TEMPLATE_FOLDER_ID = "YOUR_GOOGLE_DRIVE_TEMPLATE_FOLDER_ID_HERE"; 
const OUTPUT_FOLDER_ID = "YOUR_GOOGLE_DRIVE_OUTPUT_FOLDER_ID_HERE"; // Tùy chọn

// HIGHLIGHTED CHANGE: Đảm bảo "Ngày ghép xe" là cột cuối cùng (cột Z)
const ALL_HEADERS = [
  "Số Khung", "Tên Xe", "Màu Sắc", "Năm Sản Xuất", "Số Máy", "Giá Nhập", "Trạng Thái Xe",
  "Mã Đơn Hàng", "Mã VSO", "Tên Khách Hàng",
  "SĐT KH", "CCCD/MST KH", "Địa chỉ KH",
  "Tư Vấn BH", "Tên Showroom",
  "Giá Bán", // Giá Niêm Yết
  "Tiền cọc", "Gói phụ kiện", "Giảm giá", "Giá cuối cùng",
  "TT Thanh Toán", "Ngày Xuất HĐ", "Ngày DKGN",
  "Ngày Giao Xe TT",
  "Ghi Chú ĐH",
  "Ngày ghép xe" // Trường mới ở cuối (Cột Z)
];

let USERS = {};
let ROLES = {};

// ========================================================================
// QUẢN LÝ NGƯỜI DÙNG
// ========================================================================
function loadUsersAndRolesFromStore() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet) {
      userSheet = ss.insertSheet(USER_SHEET_NAME);
      userSheet.appendRow(USER_SHEET_HEADERS);
      userSheet.appendRow(["superadmin", "superadminpass", "superadmin"]);
      userSheet.appendRow(["admin", "adminpass", "admin"]);
      userSheet.appendRow(["viewer", "viewerpass", "viewer"]);
      SpreadsheetApp.flush();
    }
    const userValues = userSheet.getDataRange().getValues();
    const tempUsers = {};
    const tempRoles = {};
    if (userValues.length > 1) {
      const headers = userValues[0];
      const usernameCol = headers.indexOf(USER_SHEET_HEADERS[0]);
      const passwordCol = headers.indexOf(USER_SHEET_HEADERS[1]);
      const roleCol = headers.indexOf(USER_SHEET_HEADERS[2]);

      if (usernameCol === -1 || passwordCol === -1 || roleCol === -1) {
        throw new Error("Tiêu đề cột trong sheet 'Users' không đúng.");
      }
      for (let i = 1; i < userValues.length; i++) {
        const row = userValues[i];
        const username = String(row[usernameCol] || '').trim();
        const password = String(row[passwordCol] || '');
        const role = String(row[roleCol] || '').toLowerCase().trim();
        if (username && password && role) {
          tempUsers[username] = password;
          if (role === 'superadmin') tempRoles[username] = ['superadmin', 'admin'];
          else if (role === 'admin') tempRoles[username] = ['admin'];
          else tempRoles[username] = ['viewer'];
        }
      }
    }
    if (!tempUsers.superadmin) {
        tempUsers.superadmin = "superadminpass";
        tempRoles.superadmin = ['superadmin', 'admin'];
        const usernameColIdx = userValues[0].indexOf(USER_SHEET_HEADERS[0]);
        const superadminExists = userValues.slice(1).some(row => row[usernameColIdx] === "superadmin");
        if (!superadminExists) userSheet.appendRow(["superadmin", "superadminpass", "superadmin"]);
    }
    USERS = tempUsers;
    ROLES = tempRoles;
  } catch (e) {
    console.error("Lỗi tải Users/Roles: " + e.toString());
    USERS = { "superadmin": "superadminpass" };
    ROLES = { "superadmin": ["superadmin", 'admin'] };
  }
}

function saveUsersAndRolesToStore() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(USER_SHEET_NAME);
    if (!userSheet) throw new Error(`Sheet "${USER_SHEET_NAME}" không tồn tại.`);
    const dataToSave = [USER_SHEET_HEADERS];
    for (const username in USERS) {
      if (USERS.hasOwnProperty(username)) {
        dataToSave.push([username, USERS[username], ROLES[username] ? ROLES[username][0] : 'viewer']);
      }
    }
    userSheet.clearContents();
    userSheet.getRange(1, 1, dataToSave.length, USER_SHEET_HEADERS.length).setValues(dataToSave);
  } catch (e) {
    console.error("Lỗi lưu Users/Roles: " + e.toString());
  }
}

// ========================================================================
// HÀM CHÍNH WEB APP
// ========================================================================
function doGet() { return HtmlService.createTemplateFromFile("index").evaluate().setTitle("Quản Lý Kho Xe").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
function include(filename) { return HtmlService.createTemplateFromFile(filename).evaluate().getContent(); }

function authenticateUser(username, password) {
  loadUsersAndRolesFromStore();
  const activeUser = Session.getActiveUser() ? Session.getActiveUser().getEmail() : "N/A";
  console.log(`Xác thực cho: ${username} bởi ActiveUser: ${activeUser}`);
  if (USERS[username] && USERS[username] === password) {
    const userRole = ROLES[username] ? ROLES[username][0] : 'viewer';
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('userRole', userRole);
    userProperties.setProperty('username', username);
    return { success: true, role: userRole };
  }
  PropertiesService.getUserProperties().deleteAllProperties();
  return { success: false, message: "Tên đăng nhập hoặc mật khẩu không đúng." };
}

function logoutUserOnServer() {
  PropertiesService.getUserProperties().deleteAllProperties();
  return { success: true };
}

function isSuperAdmin() { return PropertiesService.getUserProperties().getProperty('userRole') === 'superadmin'; }
function isAdminOrSuperAdmin() { const role = PropertiesService.getUserProperties().getProperty('userRole'); return role === 'admin' || role === 'superadmin'; }

// ========================================================================
// HÀM XỬ LÝ DỮ LIỆU
// ========================================================================
function getDataFromSheet_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SINGLE_DATA_SHEET_NAME);
    if (!sheet) {
      const newSheet = ss.insertSheet(SINGLE_DATA_SHEET_NAME);
      newSheet.appendRow(ALL_HEADERS);
      return [];
    }
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    const headers = values[0];
    return values.slice(1).map(row => {
      const item = {};
      row.forEach((cell, index) => {
        if (headers[index]) item[headers[index]] = cell;
      });
      return item;
    });
  } catch (e) {
    console.error("Lỗi trong getDataFromSheet_: " + e.toString());
    return [];
  }
}

function getActualInventoryVinList_() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTUAL_INVENTORY_SHEET_NAME); 
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const vinColumnNumber = sheet.getRange(ACTUAL_VIN_COLUMN_LETTER + "1").getColumn();
    return sheet.getRange(2, vinColumnNumber, lastRow - 1, 1).getValues()
                .map(row => String(row[0] || "").trim().toUpperCase())
                .filter(vin => vin !== ""); 
  } catch (e) {
    console.error(`Lỗi đọc sheet "${ACTUAL_INVENTORY_SHEET_NAME}": ${e.toString()}`);
    return [];
  }
}

function getVehiclesWithOrders() {
  if (!PropertiesService.getUserProperties().getProperty('userRole')) return [];
  try {
    const mainData = getDataFromSheet_();
    const actualVinList = getActualInventoryVinList_();
    return mainData.map(vehicle => {
      const chassisNumber = String(vehicle["Số Khung"] || '').trim().toUpperCase();
      const isInActualInventory = chassisNumber ? actualVinList.includes(chassisNumber) : false;
      const pairingDate = vehicle["Ngày ghép xe"];
      let daysSincePaired = "";
      if (pairingDate && vehicle["Trạng Thái Xe"] === "Đã Ghép") {
        try {
          const today = new Date();
          today.setHours(0, 0, 0, 0);
          const pairedDateObj = new Date(pairingDate);
          if (!isNaN(pairedDateObj.getTime())) {
            pairedDateObj.setHours(0, 0, 0, 0);
            const diffTime = Math.abs(today - pairedDateObj);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            daysSincePaired = diffDays;
          }
        } catch(e) { /* ignore date parsing errors */ }
      }
      return { ...vehicle, isActualInventory: isInActualInventory, daysSincePaired: daysSincePaired };
    });
  } catch (e) {
    console.error("Lỗi trong getVehiclesWithOrders: " + e.toString());
    return [];
  }
}

function writeDataToSheet_(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SINGLE_DATA_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet ${SINGLE_DATA_SHEET_NAME} không tồn tại.`);
    const dataToSave = [ALL_HEADERS, ...data.map(item => ALL_HEADERS.map(header => item[header] !== undefined && item[header] !== null ? item[header] : ''))];
    sheet.clearContents(); 
    sheet.getRange(1, 1, dataToSave.length, ALL_HEADERS.length).setValues(dataToSave);
  } catch (e) {
    console.error("Lỗi trong writeDataToSheet_: " + e.toString());
    throw e;
  }
}

// ========================================================================
// CÁC HÀM CRUD
// ========================================================================
function addVehicle(vehicleData) {
  if (!isAdminOrSuperAdmin()) return { success: false, message: "Không có quyền." };
  try {
    const data = getDataFromSheet_();
    if (data.some(v => v["Số Khung"] === vehicleData["Số Khung"])) {
      return { success: false, message: `Số Khung ${vehicleData["Số Khung"]} đã tồn tại.` };
    }
    const newVehicle = {};
    ALL_HEADERS.forEach(header => newVehicle[header] = vehicleData[header] || "");
    data.push(newVehicle);
    writeDataToSheet_(data);
    return { success: true, message: "Thêm xe thành công." };
  } catch (e) { return { success: false, message: "Lỗi thêm xe: " + e.message }; }
}

function saveOrder(orderData) {
  if (!isAdminOrSuperAdmin()) return { success: false, message: "Không có quyền." };
  if (!orderData["TT Thanh Toán"]) return { success: false, message: "Vui lòng chọn Trạng Thái Thanh Toán." };
  try {
    const data = getDataFromSheet_();
    const index = data.findIndex(v => v["Số Khung"] === orderData["Số Khung"]);
    if (index === -1) return { success: false, message: `Không tìm thấy xe: ${orderData["Số Khung"]}.` };
    
    const currentVehicleData = data[index];
    const updatedOrderInfo = {};
    ALL_HEADERS.forEach(header => {
      if (orderData.hasOwnProperty(header) && !["Số Khung", "Tên Xe", "Màu Sắc", "Năm Sản Xuất", "Số Máy", "Giá Nhập"].includes(header)) {
        updatedOrderInfo[header] = orderData[header] || "";
      }
    });
    updatedOrderInfo["Mã Đơn Hàng"] = orderData["Mã Đơn Hàng"] || currentVehicleData["Mã Đơn Hàng"] || "";
    
    let vehicleStatus = currentVehicleData["Trạng Thái Xe"];
    let pairingDate = currentVehicleData["Ngày ghép xe"] || "";

    if (updatedOrderInfo["Mã Đơn Hàng"]) {
      if (orderData["TT Thanh Toán"] === "Đã thanh toán") {
        vehicleStatus = "Đã bán";
      } else {
        const previousStatus = currentVehicleData["Trạng Thái Xe"];
        vehicleStatus = "Đã Ghép";
        if (previousStatus !== "Đã Ghép" && !pairingDate) {
          pairingDate = new Date();
        }
      }
    } else if (!updatedOrderInfo["Mã Đơn Hàng"] && currentVehicleData["Mã Đơn Hàng"]) {
      vehicleStatus = "Còn hàng";
      pairingDate = "";
    }
    
    data[index] = { ...currentVehicleData, ...updatedOrderInfo, "Trạng Thái Xe": vehicleStatus, "Ngày ghép xe": pairingDate };
    writeDataToSheet_(data);
    return { success: true, message: "Lưu đơn hàng thành công." };
  } catch (e) { return { success: false, message: "Lỗi lưu đơn hàng: " + e.message }; }
}

function clearOrderForVehicle(chassisNumber) {
  if (!isAdminOrSuperAdmin()) return { success: false, message: "Không có quyền." };
  try {
    const data = getDataFromSheet_();
    const index = data.findIndex(v => v["Số Khung"] === chassisNumber);
    if (index === -1) return { success: false, message: `Không tìm thấy xe: ${chassisNumber}.` };
    const vehicleToClear = data[index];
    const vehicleCoreHeaders = ["Số Khung", "Tên Xe", "Màu Sắc", "Năm Sản Xuất", "Số Máy", "Giá Nhập"];
    ALL_HEADERS.forEach(header => { if (!vehicleCoreHeaders.includes(header) && header !== "Trạng Thái Xe") vehicleToClear[header] = ""; });
    vehicleToClear["Trạng Thái Xe"] = "Còn hàng";
    data[index] = vehicleToClear;
    writeDataToSheet_(data);
    return { success: true, message: "Đã xóa thông tin đơn hàng." };
  } catch (e) { return { success: false, message: "Lỗi xóa đơn hàng: " + e.message }; }
}

function updateVehicle(vehicleData) { /* ... giữ nguyên ... */ }
function deleteVehicle(chassisNumber) { /* ... giữ nguyên ... */ }
// ... các hàm quản lý user và tiện ích khác giữ nguyên ...
