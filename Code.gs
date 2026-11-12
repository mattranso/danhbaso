const SHEET_ID = '1vTCBqjZFqxX9LgnGBku0IZTp5VbLaiDh8EAhX5TvmyY'; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Danh bạ - Ủy ban Mặt trận Tổ quốc Việt Nam thành phố Đồng Nai')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HÀM MỚI: TĂNG VÀ LẤY TỔNG SỐ LƯỢT TRUY CẬP
function getAndIncrementPageViews() {
  let props = PropertiesService.getScriptProperties();
  let count = props.getProperty('TOTAL_VIEWS');
  count = count ? parseInt(count) + 1 : 1;
  props.setProperty('TOTAL_VIEWS', count.toString());
  return count.toLocaleString('vi-VN'); // Định dạng số cho đẹp (VD: 1.234)
}

function loginUser(username, password) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan');
  if(!sheet) throw new Error("Chưa có Sheet tên 'TaiKhoan'");
  const data = sheet.getDataRange().getValues();
  let u = String(username).toLowerCase().trim();
  let p = String(password);
  
  if(u === 'admin') {
    let adminPw = 'Dongnai@2026';
    for(let i = 1; i < data.length; i++) { if(String(data[i][0]) === 'admin') { adminPw = String(data[i][1]); break; } }
    if(p === adminPw) return { success: true, role: 'admin', donVi: 'Tất cả', username: 'admin', fullname: 'Quản trị viên' };
    return { success: false, msg: 'Sai tài khoản hoặc mật khẩu!' };
  }

  for(let i = 1; i < data.length; i++) {
    if(String(data[i][0]).toLowerCase().trim() === u && String(data[i][1]) === p) {
      if(String(data[i][3]) !== 'Đã duyệt') return { success: false, msg: 'Tài khoản đang chờ duyệt hoặc bị khóa!' };
      return { success: true, role: 'user', donVi: String(data[i][2]), username: u, fullname: String(data[i][4]) || u };
    }
  }
  return { success: false, msg: 'Sai tài khoản hoặc mật khẩu!' };
}

function registerAccount(username, password, donVi, fullname) {
  let u = String(username).toLowerCase().trim();
  if (!/^[a-z0-9_.]+$/.test(u) || u === 'admin') return { success: false, msg: 'Tên đăng nhập không hợp lệ!' };
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan');
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) { if(String(data[i][0]).toLowerCase() === u) return { success: false, msg: 'Tên đăng nhập đã tồn tại!' }; }
  sheet.appendRow([u, password, donVi, 'Chờ duyệt', fullname]);
  return { success: true, msg: 'Đăng ký thành công! Vui lòng chờ duyệt.' };
}

function updateUserInfo(username, newFullname) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan');
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(String(data[i][0]) === String(username)) {
      sheet.getRange(i + 1, 5).setValue(newFullname);
      return {success: true, msg: 'Cập nhật thông tin thành công!'};
    }
  }
  return {success: false, msg: 'Không tìm thấy tài khoản!'};
}

function changePassword(username, oldPw, newPw) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan');
  const data = sheet.getDataRange().getValues();
  if(String(username) === 'admin') {
    let adminRow = -1; let currentPw = 'Dongnai@2026';
    for(let i = 1; i < data.length; i++) { if(String(data[i][0]) === 'admin') { adminRow = i + 1; currentPw = String(data[i][1]); break; } }
    if(String(oldPw) !== currentPw) return {success: false, msg: 'Mật khẩu hiện tại không chính xác!'};
    if(adminRow !== -1) { sheet.getRange(adminRow, 2).setValue(newPw); } else { sheet.appendRow(['admin', newPw, 'Tất cả', 'Đã duyệt', 'Quản trị viên']); }
    return {success: true, msg: 'Đổi mật khẩu thành công!'};
  }
  for(let i = 1; i < data.length; i++) {
    if(String(data[i][0]) === String(username)) {
       if(String(data[i][1]) !== String(oldPw)) return {success: false, msg: 'Mật khẩu hiện tại không chính xác!'};
       sheet.getRange(i + 1, 2).setValue(newPw); return {success: true, msg: 'Đổi mật khẩu thành công!'};
    }
  }
  return {success: false, msg: 'Không tìm thấy tài khoản!'};
}

function getAccountList() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan');
  const data = sheet.getDataRange().getValues();
  let accs = [];
  for(let i = 1; i < data.length; i++) { 
    if(data[i][0] && String(data[i][0]) !== 'admin') {
      accs.push({ row: i + 1, username: String(data[i][0]), donVi: String(data[i][2]), status: String(data[i][3]), fullname: String(data[i][4]) || 'Chưa cập nhật' }); 
    }
  }
  return accs;
}

function approveAccount(rowNum, newStatus) { 
  if(!newStatus) newStatus = 'Đã duyệt';
  SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan').getRange(rowNum, 4).setValue(newStatus); 
  return true; 
}

function resetAccountPassword(rowNum) { SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan').getRange(rowNum, 2).setValue('Mttq@123456'); return true; }
function deleteAccount(rowNum) { SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan').deleteRow(rowNum); return true; }
function changeUserUnit(rowNum, newUnit) { SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan').getRange(rowNum, 3).setValue(newUnit); return true; }

function getDanhBa(role, donViPhuTrach) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DanhBa');
    if(!sheet) return [{id: 'error', hoten: 'LỖI: Chưa có Sheet tên DanhBa', chucvu: '', donvi: '', sdt: '', email: ''}];
    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return []; 
    let result = [];
    for(let i = 1; i < data.length; i++) {
      let row = data[i];
      if(!row || !row[1] || String(row[1]).trim() === '') continue; 
      let unitName = row[3] ? String(row[3]).trim() : 'Đơn vị Khác';
      let isMatchUser = String(role) === 'user' && (unitName === donViPhuTrach || unitName.startsWith(donViPhuTrach + " - ") || unitName.startsWith(donViPhuTrach + " -"));
      
      let rawSdt = row[5] ? String(row[5]) : '';
      let cleanSdt = rawSdt.replace(/[\s\.]/g, ''); 
      if (cleanSdt !== '' && !cleanSdt.startsWith('0')) {
        cleanSdt = '0' + cleanSdt; 
      }

      if(String(role) === 'admin' || isMatchUser || String(role) === 'guest') {
        result.push({ id: row[0] ? String(row[0]) : '', hoten: row[1] ? String(row[1]) : '', chucvu: row[2] ? String(row[2]) : '', donvi: unitName, hinhanh: row[4] ? String(row[4]) : 'https://cdn-icons-png.flaticon.com/512/149/149071.png', sdt: cleanSdt, email: row[6] ? String(row[6]) : '' });
      }
    }
    return result;
  } catch (error) { return [{id: 'error', hoten: 'LỖI HỆ THỐNG: ' + error.message, chucvu: '', donvi: '', sdt: '', email: ''}]; }
}

function saveContact(id, hoten, chucvu, donvi, sdt, email, hinhanh) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DanhBa');
  const data = sheet.getDataRange().getValues();
  
  let cleanSdt = String(sdt || '').replace(/[\s\.]/g, '');
  if (cleanSdt !== '' && !cleanSdt.startsWith('0')) {
    cleanSdt = '0' + cleanSdt;
  }
  let finalSdt = cleanSdt !== '' ? "'" + cleanSdt : ""; 

  if (id) { 
    for(let i = 1; i < data.length; i++) {
      if(String(data[i][0]) === String(id)) { 
        let finalImg = hinhanh ? hinhanh : data[i][4]; 
        sheet.getRange(i + 1, 2, 1, 6).setValues([[hoten, chucvu, donvi, finalImg, finalSdt, email]]); 
        return true; 
      }
    }
  } else { 
    sheet.appendRow([new Date().getTime(), hoten, chucvu, donvi, hinhanh ? hinhanh : 'https://cdn-icons-png.flaticon.com/512/149/149071.png', finalSdt, email]); 
    return true; 
  }
}

function deleteContact(id) { 
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DanhBa'); 
  const data = sheet.getDataRange().getValues(); 
  for(let i = 1; i < data.length; i++) { 
    if(String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); return true; } 
  } 
}

function deleteMultipleContacts(ids) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DanhBa'); 
    const data = sheet.getDataRange().getValues(); 
    let count = 0;
    for(let i = data.length - 1; i >= 1; i--) { 
      if(ids.includes(String(data[i][0]))) { 
        sheet.deleteRow(i + 1); 
        count++;
      } 
    } 
    return { success: true, count: count };
  } catch(e) {
    return { success: false, msg: e.message };
  }
}

function moveContact(id, direction) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DanhBa'); const data = sheet.getDataRange().getValues();
  if (data.length <= 2) return {success: false, msg: 'Không đủ dữ liệu.'};
  let header = data[0]; let rows = data.slice(1);
  let currentIndex = rows.findIndex(r => String(r[0]) === String(id));
  if (currentIndex === -1) return {success: false, msg: 'Không tìm thấy!'};
  let currentContact = rows[currentIndex]; let donVi = currentContact[3]; let targetIndex = -1;
  if (direction === 'up') { for (let i = currentIndex - 1; i >= 0; i--) { if (rows[i][3] === donVi) { targetIndex = i; break; } } } 
  else if (direction === 'down') { for (let i = currentIndex + 1; i < rows.length; i++) { if (rows[i][3] === donVi) { targetIndex = i; break; } } }
  if (targetIndex !== -1) { let temp = rows[currentIndex]; rows[currentIndex] = rows[targetIndex]; rows[targetIndex] = temp; let newData = [header].concat(rows); sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData); return {success: true}; } 
  else { return {success: false, msg: 'Đã ở vị trí tận cùng.'}; }
}

function importCSVData(csvText, role, userDonVi) {
  try {
    let data = Utilities.parseCsv(csvText); let sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DanhBa'); let count = 0;
    for(let i = 1; i < data.length; i++) {
      let row = data[i]; if(!row || !row[1] || String(row[1]).trim() === '') continue; 
      let targetDonVi = (String(role) === 'admin') ? (row[3] || 'Đơn vị Khác') : userDonVi;
      
      let cleanSdt = String(row[5] || '').replace(/[\s\.]/g, '');
      if (cleanSdt !== '' && !cleanSdt.startsWith('0')) { cleanSdt = '0' + cleanSdt; }
      let finalSdt = cleanSdt !== '' ? "'" + cleanSdt : "";

      sheet.appendRow([new Date().getTime() + count, row[1], row[2], targetDonVi, row[4] || 'https://cdn-icons-png.flaticon.com/512/149/149071.png', finalSdt, row[6]]); count++;
    }
    return {success: true, msg: `Đã nhập thành công ${count} liên hệ!`};
  } catch (e) { return {success: false, msg: "Lỗi cấu trúc file CSV."}; }
}

function getInitData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonVi');
  let units = [];
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) { if(data[i][0]) units.push(String(data[i][0]).trim()); }
  }
  
  let props = PropertiesService.getScriptProperties();
  let blocksStr = props.getProperty('explicitCityBlocks');
  let cityBlocks = [];
  if (blocksStr) {
    cityBlocks = JSON.parse(blocksStr);
  } else {
    cityBlocks = ["Ban Thường trực", "Văn phòng", "Liên đoàn Lao động thành phố", "Hội Liên hiệp Phụ nữ Việt Nam thành phố", "Hội Nông dân thành phố", "Hội Cựu Chiến binh thành phố", "Thành đoàn"];
    props.setProperty('explicitCityBlocks', JSON.stringify(cityBlocks));
  }
  
  return { units: units, cityBlocks: cityBlocks };
}

function saveCityBlocks(blocks) {
  PropertiesService.getScriptProperties().setProperty('explicitCityBlocks', JSON.stringify(blocks));
  return true;
}

function getDonViList() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonVi');
  if(!sheet) return ["Lỗi: Chưa có Sheet DonVi"];
  const data = sheet.getDataRange().getValues(); let units = [];
  for(let i = 1; i < data.length; i++) { if(data[i][0]) units.push(String(data[i][0]).trim()); }
  return units;
}

function getSubUnits(parentUnit) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('DonViCon');
  if(!sheet) {
    sheet = ss.insertSheet('DonViCon');
    sheet.appendRow(['DonViCha', 'TenDonViCon']);
    return [];
  }
  let data = sheet.getDataRange().getValues();
  let subs = [];
  for(let i = 1; i < data.length; i++) {
    let subName = data[i][1] ? String(data[i][1]).trim() : '';
    if(subName === "Thôn 1" || subName === "Thôn 2" || subName.includes("Ban Công tác Mặt trận") || subName.includes("Chi hội") || subName.includes("Chi đoàn")) {
      continue;
    }
    if(String(data[i][0]) === String(parentUnit) && subName) {
      subs.push({ row: i + 1, name: subName });
    }
  }
  return subs; 
}

function addSubUnit(parentUnit, subName, userRole, userUnit) {
  if (String(userRole) !== 'admin' && String(userUnit) !== String(parentUnit)) {
    return { success: false, msg: 'Bạn không có quyền thêm đơn vị con cho xã/phường khác!' };
  }
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('DonViCon');
  if(!sheet) { sheet = ss.insertSheet('DonViCon'); sheet.appendRow(['DonViCha', 'TenDonViCon']); }
  sheet.appendRow([parentUnit, subName]);
  return {success: true};
}

function editSubUnit(rowNum, newName, userRole, userUnit) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonViCon');
  if(sheet && rowNum > 1) {
    let parentUnit = String(sheet.getRange(rowNum, 1).getValue());
    if (String(userRole) !== 'admin' && String(userUnit) !== parentUnit) {
      return { success: false, msg: 'Bạn chỉ được sửa đơn vị con thuộc xã của mình!' };
    }
    sheet.getRange(rowNum, 2).setValue(newName);
    return {success: true};
  }
  return {success: false, msg: 'Không tìm thấy dòng dữ liệu!'};
}

function moveSubUnitToParent(rowNum, newParentUnit, userRole, userUnit) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonViCon');
  if(sheet && rowNum > 1) {
    let currentParent = String(sheet.getRange(rowNum, 1).getValue());
    if (String(userRole) !== 'admin' && (String(userUnit) !== currentParent || String(userUnit) !== String(newParentUnit))) {
      return { success: false, msg: 'Bạn chỉ được di chuyển đơn vị con trong phạm vi xã của mình!' };
    }
    sheet.getRange(rowNum, 1).setValue(newParentUnit);
    return {success: true};
  }
  return {success: false, msg: 'Không thể di chuyển đơn vị con!'};
}

function deleteSubUnit(rowNum, userRole, userUnit) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonViCon');
  if(sheet && rowNum > 1) {
    let parentUnit = String(sheet.getRange(rowNum, 1).getValue());
    if (String(userRole) !== 'admin' && String(userUnit) !== parentUnit) {
      return { success: false, msg: 'Bạn chỉ được xóa đơn vị con thuộc xã của mình!' };
    }
    sheet.deleteRow(rowNum);
    return {success: true};
  }
  return {success: false, msg: 'Không thể xóa!'};
}

function addDonVi(name) { 
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonVi'); 
  const data = sheet.getDataRange().getValues(); 
  for(let i = 1; i < data.length; i++) { 
    if(String(data[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      return { success: false, msg: 'Đơn vị đã tồn tại!' }; 
    } 
  } 
  sheet.appendRow([name]); 
  return { success: true }; 
}

function editDonVi(oldName, newName) {
  const sheetDonVi = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonVi'); const dataDonVi = sheetDonVi.getDataRange().getValues();
  for(let i = 1; i < dataDonVi.length; i++) { if(String(dataDonVi[i][0]) === String(newName)) return { success: false, msg: 'Tên Đơn vị mới đã tồn tại!' }; }
  let updated = false;
  for(let i = 1; i < dataDonVi.length; i++) { if(String(dataDonVi[i][0]) === String(oldName)) { sheetDonVi.getRange(i + 1, 1).setValue(newName); updated = true; break; } }
  if(!updated) return { success: false, msg: 'Lỗi không tìm thấy đơn vị!' };

  const sheetDanhBa = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DanhBa'); const dataDanhBa = sheetDanhBa.getDataRange().getValues();
  for(let i = 1; i < dataDanhBa.length; i++) { if(String(dataDanhBa[i][3]) === String(oldName)) sheetDanhBa.getRange(i + 1, 4).setValue(newName); }
  const sheetTaiKhoan = SpreadsheetApp.openById(SHEET_ID).getSheetByName('TaiKhoan'); const dataTaiKhoan = sheetTaiKhoan.getDataRange().getValues();
  for(let i = 1; i < dataTaiKhoan.length; i++) { if(String(dataTaiKhoan[i][2]) === String(oldName)) sheetTaiKhoan.getRange(i + 1, 3).setValue(newName); }
  return { success: true, msg: 'Đổi tên thành công!' };
}

function deleteDonVi(name) { const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonVi'); const data = sheet.getDataRange().getValues(); for(let i = 1; i < data.length; i++) { if(String(data[i][0]) === String(name)) { sheet.deleteRow(i + 1); return { success: true }; } } }

function moveDonVi(name, direction) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DonVi'); const data = sheet.getDataRange().getValues();
  if (data.length <= 2) return {success: false, msg: 'Không đủ dữ liệu để di chuyển.'};
  let header = data.shift(); let index = data.findIndex(r => String(r[0]) === String(name));
  if (index === -1) return {success: false, msg: 'Không tìm thấy!'};
  let item = data.splice(index, 1)[0]; 
  if (direction === 'up' && index > 0) { data.splice(index - 1, 0, item); } else if (direction === 'down' && index < data.length) { data.splice(index + 1, 0, item); } else if (direction === 'top') { data.unshift(item); } else if (direction === 'bottom') { data.push(item); } else { data.splice(index, 0, item); return {success: false, msg: 'Không thể di chuyển thêm.'}; }
  let newData = [header].concat(data); sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData); return {success: true};
}