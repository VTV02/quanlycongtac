// ============================================
// HỆ THỐNG QUẢN LÝ CÔNG TÁC - THACO AGRI KLH SNUOL
// Version: 2.1.0 - UPDATED
// Author: Trung IT
// 
// CẬP NHẬT:
// - Sửa lỗi tên địa điểm không đồng bộ
// - Thêm cách xưng hô tự động trong email
// - Phân quyền chặt chẽ: chỉ người trong danh sách "Phân quyền" mới có quyền duyệt
// ============================================

// ============================================
// CẤU HÌNH HỆ THỐNG
// ============================================
var CONFIG = {
  sheetNames: {
    dangKy: 'Đăng ký công tác',
    cauHinhEmail: 'Cấu hình Email',
    danhMuc: 'Danh mục',
    phanQuyen: 'Phân quyền',
    lichSuDuyet: 'Lịch sử duyệt',
    thongKe: 'Thống kê'
  },
  columns: {
    timestamp: 0,       // A
    maDangKy: 1,        // B
    hoTen: 2,           // C
    chucVu: 3,          // D
    phongBan: 4,        // E
    thongTin: 5,        // F
    ngayDen: 6,         // G
    ngayDi: 7,          // H
    phuongTien: 8,      // I
    nhaAn: 9,           // J
    diaDiem: 10,        // K
    email: 11,          // L
    trangThai: 12,      // M
    emailDaGui: 13,     // N
    thoiGianGui: 14,    // O
    nguoiDuyet: 15,     // P
    ngayDuyet: 16,      // Q
    lyDoTuChoi: 17,     // R
    lichSu: 18,         // S
    fileDinhKem: 19,    // T
    daXoa: 20           // U
  },
  emailSubject: '[THACO AGRI KLH SNUOL] Thông báo tiếp nhận công tác',
  brandColor: '#00A86B',
  webAppUrl: '' // Sẽ cập nhật sau khi deploy
};

// ============================================
// HÀM TẠO MENU CUSTOM
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(' Quản lý Công tác')
    .addItem('Mở Dashboard', 'openDashboard')
    .addSeparator()
    .addItem('Cài đặt hệ thống', 'setupSystem')
    .addItem('Xem thống kê', 'viewStatistics')
    .addSeparator()
    .addItem('Export Excel', 'exportToExcel')
    .addItem('Dọn dẹp dữ liệu cũ', 'cleanOldData')
    .addToUi();
}

// ============================================
// MỞ DASHBOARD WEB APP
// ============================================
function openDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Quản lý Công tác - THACO AGRI KLH SNUOL')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Quản lý Công tác');
}

// ============================================
// SERVE DASHBOARD KHI TRUY CẬP URL
// ============================================
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Dashboard');
  return template.evaluate()
    .setTitle('Hệ thống Quản lý Công tác - THACO AGRI KLH SNUOL')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================
// INCLUDE CSS/JS FILES
// ============================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// API: LẤY DỮ LIỆU DANH MỤC
// ============================================
function getDanhMuc() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.danhMuc);
    var data = sheet.getDataRange().getValues();
    
    return {
      success: true,
      data: {
        chucVu: getColumnData(data, 0),
        phongBan: getColumnData(data, 2),
        phuongTien: getColumnData(data, 4),
        nhaAn: getColumnData(data, 6),
        diaDiem: getColumnData(data, 8),
        trangThai: getColumnData(data, 10)
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getColumnData(data, colIndex) {
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][colIndex]) {
      result.push(data[i][colIndex]);
    }
  }
  return result;
}

// ============================================
// API: LẤY DỮ LIỆU ĐĂNG KÝ (CÓ PHÂN TRANG & LỌC)
// ============================================
function getDangKyList(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var data = sheet.getDataRange().getValues();
    
    var result = [];
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Bỏ qua header
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Bỏ qua dòng đã xóa
      if (row[CONFIG.columns.daXoa] === true) continue;
      
      // Kiểm tra quyền xem
      if (!canViewRecord(userRole, userEmail, row)) continue;
      
      // Áp dụng filters
      if (filters) {
        if (filters.search && !matchSearch(row, filters.search)) continue;
        if (filters.trangThai && row[CONFIG.columns.trangThai] !== filters.trangThai) continue;
        if (filters.diaDiem && row[CONFIG.columns.diaDiem] !== filters.diaDiem) continue;
        if (filters.fromDate && new Date(row[CONFIG.columns.ngayDen]) < new Date(filters.fromDate)) continue;
        if (filters.toDate && new Date(row[CONFIG.columns.ngayDen]) > new Date(filters.toDate)) continue;
      }
      
      result.push({
        rowIndex: i + 1,
        timestamp: formatDateTime(row[CONFIG.columns.timestamp]),
        maDangKy: row[CONFIG.columns.maDangKy],
        hoTen: row[CONFIG.columns.hoTen],
        chucVu: row[CONFIG.columns.chucVu],
        phongBan: row[CONFIG.columns.phongBan],
        thongTin: row[CONFIG.columns.thongTin],
        ngayDen: formatDate(row[CONFIG.columns.ngayDen]),
        ngayDi: formatDate(row[CONFIG.columns.ngayDi]),
        phuongTien: row[CONFIG.columns.phuongTien],
        nhaAn: row[CONFIG.columns.nhaAn],
        diaDiem: row[CONFIG.columns.diaDiem],
        email: row[CONFIG.columns.email],
        trangThai: row[CONFIG.columns.trangThai],
        nguoiDuyet: row[CONFIG.columns.nguoiDuyet],
        ngayDuyet: row[CONFIG.columns.ngayDuyet] ? formatDateTime(row[CONFIG.columns.ngayDuyet]) : '',
        lyDoTuChoi: row[CONFIG.columns.lyDoTuChoi]
      });
    }
    
    return {
      success: true,
      data: result,
      total: result.length
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: THÊM ĐĂNG KÝ MỚI
// ============================================
function submitDangKy(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    
    // Validate dữ liệu
    var validation = validateFormData(formData);
    if (!validation.valid) {
      return {
        success: false,
        error: validation.error
      };
    }
    
    // Tạo mã đăng ký tự động
    var maDangKy = generateMaDangKy();
    
    // Tạo timestamp
    var now = new Date();
    
    // Tạo log lịch sử
    var lichSu = JSON.stringify([{
      time: formatDateTime(now),
      action: 'Tạo mới',
      user: userEmail
    }]);
    
    // Thêm dòng mới
    var newRow = [
      now,                                    // A: Timestamp
      maDangKy,                               // B: Mã đăng ký
      formData.hoTen,                         // C: Họ tên
      formData.chucVu,                        // D: Chức vụ
      formData.phongBan,                      // E: Loại phòng
      formData.thongTin || '',                // F: Thông tin
      new Date(formData.ngayDen),             // G: Ngày đến
      new Date(formData.ngayDi),              // H: Ngày đi
      formData.phuongTien,                    // I: Phương tiện
      formData.nhaAn ? formData.nhaAn.join(', ') : '', // J: Nhà ăn
      formData.diaDiem,                       // K: Địa điểm
      formData.email,                         // L: Email
      'Chờ duyệt',                            // M: Trạng thái
      '',                                     // N: Email đã gửi
      '',                                     // O: Thời gian gửi
      '',                                     // P: Người duyệt
      '',                                     // Q: Ngày duyệt
      '',                                     // R: Lý do từ chối
      lichSu,                                 // S: Lịch sử
      '',                                     // T: File đính kèm
      false                                   // U: Đã xóa
    ];
    
    sheet.appendRow(newRow);
    var newRowIndex = sheet.getLastRow();
    
    // Format dòng mới
    formatNewRow(sheet, newRowIndex);
    
    // Gửi email thông báo
    var emailResult = sendEmailThongBao(maDangKy, formData, 'Chờ duyệt');
    
    // Cập nhật thông tin email đã gửi
    if (emailResult.success) {
      sheet.getRange(newRowIndex, CONFIG.columns.emailDaGui + 1).setValue(emailResult.sentTo);
      sheet.getRange(newRowIndex, CONFIG.columns.thoiGianGui + 1).setValue(new Date());
    }
    
    return {
      success: true,
      message: 'Đăng ký thành công! Mã đăng ký: ' + maDangKy,
      maDangKy: maDangKy,
      rowIndex: newRowIndex
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: CẬP NHẬT ĐĂNG KÝ
// ============================================
function updateDangKy(rowIndex, formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Kiểm tra quyền
    var currentData = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    if (!canEditRecord(userRole, userEmail, currentData)) {
      return {
        success: false,
        error: 'Bạn không có quyền chỉnh sửa đăng ký này!'
      };
    }
    
    // Validate
    var validation = validateFormData(formData);
    if (!validation.valid) {
      return {
        success: false,
        error: validation.error
      };
    }
    
    // Lấy lịch sử cũ và thêm log mới
    var oldLichSu = currentData[CONFIG.columns.lichSu];
    var lichSuArray = oldLichSu ? JSON.parse(oldLichSu) : [];
    lichSuArray.push({
      time: formatDateTime(new Date()),
      action: 'Chỉnh sửa',
      user: userEmail
    });
    
    // Cập nhật dữ liệu
    sheet.getRange(rowIndex, CONFIG.columns.hoTen + 1).setValue(formData.hoTen);
    sheet.getRange(rowIndex, CONFIG.columns.chucVu + 1).setValue(formData.chucVu);
    sheet.getRange(rowIndex, CONFIG.columns.phongBan + 1).setValue(formData.phongBan);
    sheet.getRange(rowIndex, CONFIG.columns.thongTin + 1).setValue(formData.thongTin || '');
    sheet.getRange(rowIndex, CONFIG.columns.ngayDen + 1).setValue(new Date(formData.ngayDen));
    sheet.getRange(rowIndex, CONFIG.columns.ngayDi + 1).setValue(new Date(formData.ngayDi));
    sheet.getRange(rowIndex, CONFIG.columns.phuongTien + 1).setValue(formData.phuongTien);
    sheet.getRange(rowIndex, CONFIG.columns.nhaAn + 1).setValue(formData.nhaAn ? formData.nhaAn.join(', ') : '');
    sheet.getRange(rowIndex, CONFIG.columns.diaDiem + 1).setValue(formData.diaDiem);
    sheet.getRange(rowIndex, CONFIG.columns.email + 1).setValue(formData.email);
    sheet.getRange(rowIndex, CONFIG.columns.lichSu + 1).setValue(JSON.stringify(lichSuArray));
    
    return {
      success: true,
      message: 'Cập nhật thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: XÓA ĐĂNG KÝ (XÓA MỀM)
// ============================================
function deleteDangKy(rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Kiểm tra quyền
    var currentData = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    if (!canDeleteRecord(userRole, userEmail, currentData)) {
      return {
        success: false,
        error: 'Bạn không có quyền xóa đăng ký này!'
      };
    }
    
    // Xóa mềm - đánh dấu đã xóa
    sheet.getRange(rowIndex, CONFIG.columns.daXoa + 1).setValue(true);
    
    // Thêm log
    var oldLichSu = currentData[CONFIG.columns.lichSu];
    var lichSuArray = oldLichSu ? JSON.parse(oldLichSu) : [];
    lichSuArray.push({
      time: formatDateTime(new Date()),
      action: 'Xóa',
      user: userEmail
    });
    sheet.getRange(rowIndex, CONFIG.columns.lichSu + 1).setValue(JSON.stringify(lichSuArray));
    
    // Tô màu xám dòng đã xóa
    sheet.getRange(rowIndex, 1, 1, 21).setBackground('#f0f0f0');
    
    return {
      success: true,
      message: 'Xóa thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: DUYỆT ĐĂNG KÝ
// ============================================
function approveDangKy(rowIndex, ghiChu) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Kiểm tra quyền duyệt
    if (userRole !== 'Admin' && userRole !== 'Approver') {
      return {
        success: false,
        error: 'Bạn không có quyền duyệt đăng ký!'
      };
    }
    
    var currentData = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    var diaDiem = currentData[CONFIG.columns.diaDiem];
    
    // Kiểm tra Approver chỉ được duyệt địa điểm của mình
    if (userRole === 'Approver' && !canApproveLocation(userEmail, diaDiem)) {
      return {
        success: false,
        error: 'Bạn chỉ có thể duyệt đăng ký thuộc địa điểm bạn quản lý!'
      };
    }
    
    var now = new Date();
    
    // Cập nhật trạng thái
    sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setValue('Đã duyệt');
    sheet.getRange(rowIndex, CONFIG.columns.nguoiDuyet + 1).setValue(userEmail);
    sheet.getRange(rowIndex, CONFIG.columns.ngayDuyet + 1).setValue(now);
    
    // Tô màu xanh
    sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setBackground('#d9ead3');
    
    // Thêm log lịch sử
    var oldLichSu = currentData[CONFIG.columns.lichSu];
    var lichSuArray = oldLichSu ? JSON.parse(oldLichSu) : [];
    lichSuArray.push({
      time: formatDateTime(now),
      action: 'Duyệt',
      user: userEmail,
      note: ghiChu || ''
    });
    sheet.getRange(rowIndex, CONFIG.columns.lichSu + 1).setValue(JSON.stringify(lichSuArray));
    
    // Lưu vào lịch sử duyệt
    saveApprovalHistory(currentData[CONFIG.columns.maDangKy], userEmail, 'Duyệt', ghiChu);
    
    // Gửi email thông báo
    var formData = rowToFormData(currentData);
    sendEmailThongBao(currentData[CONFIG.columns.maDangKy], formData, 'Đã duyệt', ghiChu);
    
    return {
      success: true,
      message: 'Duyệt thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: TỪ CHỐI ĐĂNG KÝ
// ============================================
function rejectDangKy(rowIndex, lyDo) {
  try {
    if (!lyDo || lyDo.trim() === '') {
      return {
        success: false,
        error: 'Vui lòng nhập lý do từ chối!'
      };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    
    // Kiểm tra quyền
    if (userRole !== 'Admin' && userRole !== 'Approver') {
      return {
        success: false,
        error: 'Bạn không có quyền từ chối đăng ký!'
      };
    }
    
    var currentData = sheet.getRange(rowIndex, 1, 1, 21).getValues()[0];
    var diaDiem = currentData[CONFIG.columns.diaDiem];
    
    if (userRole === 'Approver' && !canApproveLocation(userEmail, diaDiem)) {
      return {
        success: false,
        error: 'Bạn chỉ có thể từ chối đăng ký thuộc địa điểm bạn quản lý!'
      };
    }
    
    var now = new Date();
    
    // Cập nhật trạng thái
    sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setValue('Từ chối');
    sheet.getRange(rowIndex, CONFIG.columns.nguoiDuyet + 1).setValue(userEmail);
    sheet.getRange(rowIndex, CONFIG.columns.ngayDuyet + 1).setValue(now);
    sheet.getRange(rowIndex, CONFIG.columns.lyDoTuChoi + 1).setValue(lyDo);
    
    // Tô màu đỏ
    sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setBackground('#f4cccc');
    
    // Thêm log
    var oldLichSu = currentData[CONFIG.columns.lichSu];
    var lichSuArray = oldLichSu ? JSON.parse(oldLichSu) : [];
    lichSuArray.push({
      time: formatDateTime(now),
      action: 'Từ chối',
      user: userEmail,
      note: lyDo
    });
    sheet.getRange(rowIndex, CONFIG.columns.lichSu + 1).setValue(JSON.stringify(lichSuArray));
    
    // Lưu lịch sử duyệt
    saveApprovalHistory(currentData[CONFIG.columns.maDangKy], userEmail, 'Từ chối', lyDo);
    
    // Gửi email
    var formData = rowToFormData(currentData);
    sendEmailThongBao(currentData[CONFIG.columns.maDangKy], formData, 'Từ chối', lyDo);
    
    return {
      success: true,
      message: 'Từ chối thành công!'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: LẤY THỐNG KÊ
// ============================================
function getThongKe() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetThongKe = ss.getSheetByName(CONFIG.sheetNames.thongKe);
    
    var data = {
      tongSo: sheetThongKe.getRange('B2').getValue(),
      thangNay: sheetThongKe.getRange('B3').getValue(),
      choDuyet: sheetThongKe.getRange('B4').getValue(),
      daDuyet: sheetThongKe.getRange('B5').getValue(),
      tuChoi: sheetThongKe.getRange('B6').getValue(),
      theoDiaDiem: {
        'Văn phòng 55': sheetThongKe.getRange('B9').getValue(),
        'Bình Phước 1': sheetThongKe.getRange('B10').getValue(),
        'Bình Phước 2': sheetThongKe.getRange('B11').getValue(),
        'ERC': sheetThongKe.getRange('B12').getValue(),
        'Xi nghiệp Bò': sheetThongKe.getRange('B13').getValue()
      }
    };
    
    return {
      success: true,
      data: data
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// API: LẤY THÔNG TIN USER
// ============================================
function getUserInfo() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var userRole = getUserRole(userEmail);
    var managedLocations = getManagedLocations(userEmail);
    
    return {
      success: true,
      data: {
        email: userEmail,
        role: userRole,
        managedLocations: managedLocations
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// HELPER: LẤY VAI TRÒ USER
// QUAN TRỌNG: Chỉ người có email trong sheet "Phân quyền" 
// với trạng thái "Active" mới được cấp quyền
// ============================================
function getUserRole(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.phanQuyen);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    // Cột: A=Email, B=Họ tên, C=Vai trò, D=Địa điểm quản lý, E=Trạng thái
    if (data[i][0] === email && data[i][4] === 'Active') {
      return data[i][2]; // Cột C: Vai trò
    }
  }
  
  return 'User'; // Mặc định: người dùng thường (chỉ xem/sửa/xóa đăng ký của mình)
}

// ============================================
// HELPER: LẤY ĐỊA ĐIỂM QUẢN LÝ
// QUAN TRỌNG: Approver chỉ được duyệt đăng ký thuộc địa điểm mình quản lý
// Admin có quyền duyệt tất cả (địa điểm = "Tất cả")
// ============================================
function getManagedLocations(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.phanQuyen);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    // Cột: A=Email, B=Họ tên, C=Vai trò, D=Địa điểm quản lý, E=Trạng thái
    if (data[i][0] === email && data[i][4] === 'Active') {
      var locations = data[i][3]; // Cột D: Địa điểm quản lý
      if (locations === 'Tất cả') {
        return ['Tất cả']; // Admin quản lý tất cả
      }
      // Có thể quản lý nhiều địa điểm, cách nhau bởi dấu phẩy
      return locations.split(',').map(function(loc) { return loc.trim(); });
    }
  }
  
  return []; // Không có quyền quản lý
}

// ============================================
// HELPER: KIỂM TRA QUYỀN XEM
// PHÂN QUYỀN:
// - Admin: Xem tất cả
// - Approver: Xem đăng ký thuộc địa điểm mình quản lý
// - User: Chỉ xem đăng ký của chính mình
// ============================================
function canViewRecord(role, email, rowData) {
  if (role === 'Admin') return true;
  
  if (role === 'Approver') {
    var managedLocations = getManagedLocations(email);
    return managedLocations.indexOf('Tất cả') !== -1 || 
           managedLocations.indexOf(rowData[CONFIG.columns.diaDiem]) !== -1;
  }
  
  // User chỉ xem của mình
  return rowData[CONFIG.columns.email] === email;
}

// ============================================
// HELPER: KIỂM TRA QUYỀN SỬA
// PHÂN QUYỀN:
// - Admin: Sửa tất cả
// - User: Chỉ sửa được đăng ký của mình khi đang "Chờ duyệt"
// - Approver: KHÔNG được sửa, chỉ được duyệt/từ chối
// ============================================
function canEditRecord(role, email, rowData) {
  if (role === 'Admin') return true;
  
  // User/Approver chỉ sửa được khi đang Chờ duyệt và là người tạo
  return rowData[CONFIG.columns.trangThai] === 'Chờ duyệt' && 
         rowData[CONFIG.columns.email] === email;
}

// ============================================
// HELPER: KIỂM TRA QUYỀN XÓA
// PHÂN QUYỀN:
// - Admin: Xóa tất cả
// - User: Chỉ xóa được đăng ký của mình khi đang "Chờ duyệt"
// - Approver: KHÔNG được xóa, chỉ được duyệt/từ chối
// ============================================
function canDeleteRecord(role, email, rowData) {
  if (role === 'Admin') return true;
  
  // User/Approver chỉ xóa được khi đang Chờ duyệt và là người tạo
  return rowData[CONFIG.columns.trangThai] === 'Chờ duyệt' && 
         rowData[CONFIG.columns.email] === email;
}

// ============================================
// HELPER: KIỂM TRA QUYỀN DUYỆT ĐỊA ĐIỂM
// QUAN TRỌNG: Approver CHỈ được duyệt đăng ký thuộc địa điểm mình quản lý
// Ví dụ: Anh Tịnh chỉ duyệt được "Bình Phước 1"
// ============================================
function canApproveLocation(email, diaDiem) {
  var managedLocations = getManagedLocations(email);
  return managedLocations.indexOf('Tất cả') !== -1 || 
         managedLocations.indexOf(diaDiem) !== -1;
}

// ============================================
// HELPER: VALIDATE FORM DATA
// ============================================
function validateFormData(data) {
  if (!data.hoTen || data.hoTen.trim() === '') {
    return { valid: false, error: 'Vui lòng nhập họ tên!' };
  }
  if (!data.chucVu) {
    return { valid: false, error: 'Vui lòng chọn chức vụ!' };
  }
  if (!data.phongBan) {
    return { valid: false, error: 'Vui lòng chọn Loại phòng!' };
  }
  if (!data.ngayDen) {
    return { valid: false, error: 'Vui lòng chọn ngày đến!' };
  }
  if (!data.ngayDi) {
    return { valid: false, error: 'Vui lòng chọn ngày đi!' };
  }
  if (!data.diaDiem) {
    return { valid: false, error: 'Vui lòng chọn địa điểm công tác!' };
  }
  if (!data.email || !isValidEmail(data.email)) {
    return { valid: false, error: 'Email không hợp lệ!' };
  }
  
  // Kiểm tra ngày đi > ngày đến
  var ngayDen = new Date(data.ngayDen);
  var ngayDi = new Date(data.ngayDi);
  if (ngayDi < ngayDen) {
    return { valid: false, error: 'Ngày đi phải sau ngày đến!' };
  }
  
  return { valid: true };
}

// ============================================
// HELPER: VALIDATE EMAIL
// ============================================
function isValidEmail(email) {
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

// ============================================
// HELPER: TẠO MÃ ĐĂNG KÝ TỰ ĐỘNG
// ============================================
function generateMaDangKy() {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'GMT+7', 'yyyyMMdd');
  var prefix = 'DK-' + dateStr + '-';
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
  var data = sheet.getRange('B:B').getValues();
  
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var ma = data[i][0];
    if (ma && ma.toString().indexOf(prefix) === 0) {
      var num = parseInt(ma.toString().split('-')[2]);
      if (num > maxNum) maxNum = num;
    }
  }
  
  var newNum = (maxNum + 1).toString().padStart(3, '0');
  return prefix + newNum;
}

// ============================================
// HELPER: FORMAT DATE
// ============================================
function formatDate(date) {
  if (!date) return '';
  var d = new Date(date);
  return Utilities.formatDate(d, 'GMT+7', 'dd/MM/yyyy');
}

function formatDateTime(date) {
  if (!date) return '';
  var d = new Date(date);
  return Utilities.formatDate(d, 'GMT+7', 'dd/MM/yyyy HH:mm:ss');
}

// ============================================
// HELPER: FORMAT DÒNG MỚI
// ============================================
function formatNewRow(sheet, rowIndex) {
  // Format ngày tháng
  sheet.getRange(rowIndex, CONFIG.columns.timestamp + 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sheet.getRange(rowIndex, CONFIG.columns.ngayDen + 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(rowIndex, CONFIG.columns.ngayDi + 1).setNumberFormat('dd/mm/yyyy');
  
  // Tô màu vàng cho Chờ duyệt
  sheet.getRange(rowIndex, CONFIG.columns.trangThai + 1).setBackground('#fff2cc');
}

// ============================================
// HELPER: MATCH SEARCH
// ============================================
function matchSearch(row, searchTerm) {
  searchTerm = searchTerm.toLowerCase();
  var searchFields = [
    row[CONFIG.columns.maDangKy],
    row[CONFIG.columns.hoTen],
    row[CONFIG.columns.email],
    row[CONFIG.columns.diaDiem],
    row[CONFIG.columns.phongBan]
  ];
  
  for (var i = 0; i < searchFields.length; i++) {
    if (searchFields[i] && searchFields[i].toString().toLowerCase().indexOf(searchTerm) !== -1) {
      return true;
    }
  }
  return false;
}

// ============================================
// HELPER: ROW TO FORM DATA
// ============================================
function rowToFormData(row) {
  return {
    hoTen: row[CONFIG.columns.hoTen],
    chucVu: row[CONFIG.columns.chucVu],
    phongBan: row[CONFIG.columns.phongBan],
    thongTin: row[CONFIG.columns.thongTin],
    ngayDen: row[CONFIG.columns.ngayDen],
    ngayDi: row[CONFIG.columns.ngayDi],
    phuongTien: row[CONFIG.columns.phuongTien],
    nhaAn: row[CONFIG.columns.nhaAn],
    diaDiem: row[CONFIG.columns.diaDiem],
    email: row[CONFIG.columns.email]
  };
}

// ============================================
// HELPER: LƯU LỊCH SỬ DUYỆT
// ============================================
function saveApprovalHistory(maDangKy, nguoiThaoTac, hanhDong, ghiChu) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.lichSuDuyet);
  
  var lastRow = sheet.getLastRow();
  var newId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;
  
  sheet.appendRow([
    newId,
    maDangKy,
    nguoiThaoTac,
    hanhDong,
    ghiChu || '',
    new Date(),
    '' // IP Address (có thể thêm sau)
  ]);
}

// ============================================
// GỬI EMAIL THÔNG BÁO
// ============================================
function sendEmailThongBao(maDangKy, formData, trangThai, ghiChu) {
  try {
    // Lấy thông tin người nhận theo địa điểm
    var thongTinNguoiNhan = layThongTinNguoiNhan(formData.diaDiem);
    
    if (!thongTinNguoiNhan) {
      return { success: false, error: 'Không tìm thấy email người nhận' };
    }
    
    // Tạo subject theo trạng thái
    var subject = CONFIG.emailSubject;
    if (trangThai === 'Đã duyệt') {
      subject = '[THACO AGRI KLH SNUOL]Đăng ký công tác đã được duyệt - ' + maDangKy;
    } else if (trangThai === 'Từ chối') {
      subject = '[THACO AGRI KLH SNUOL] Đăng ký công tác bị từ chối - ' + maDangKy;
    }
    
    // Tạo nội dung email với cách xưng hô phù hợp
    var emailBody = taoNoiDungEmail(maDangKy, formData, trangThai, ghiChu, thongTinNguoiNhan.cachXungHo);
    
    // Gửi email
    var recipients = thongTinNguoiNhan.email;
    var cc = formData.email;
    
    if (trangThai === 'Đã duyệt' || trangThai === 'Từ chối') {
      // Chỉ gửi cho người đăng ký và người duyệt
      recipients = formData.email;
      cc = thongTinNguoiNhan.email;
    }
    
    MailApp.sendEmail({
      to: recipients,
      cc: cc,
      subject: subject,
      htmlBody: emailBody
    });
    
    return {
      success: true,
      sentTo: recipients
    };
    
  } catch (error) {
    Logger.log('Lỗi gửi email: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// LẤY THÔNG TIN NGƯỜI NHẬN THEO ĐỊA ĐIỂM
// ============================================
function layThongTinNguoiNhan(diaDiem) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNames.cauHinhEmail);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    // Cột: A=Địa điểm, B=Email, C=Tên, D=Cách xưng hô, E=Ghi chú, F=Trạng thái
    if (data[i][0] === diaDiem && data[i][5] === 'Active') {
      return {
        email: data[i][1],
        ten: data[i][2],
        cachXungHo: data[i][3]
      };
    }
  }
  
  return null;
}

// ============================================
// LẤY EMAIL NGƯỜI NHẬN (BACKWARD COMPATIBLE)
// ============================================
function layEmailNguoiNhan(diaDiem) {
  var thongTin = layThongTinNguoiNhan(diaDiem);
  return thongTin ? thongTin.email : null;
}

// ============================================
// TẠO NỘI DUNG EMAIL HTML
// ============================================
function taoNoiDungEmail(maDangKy, formData, trangThai, ghiChu, cachXungHo) {
  var statusBadge = '';
  var statusColor = '';
  var statusText = '';
  
  if (trangThai === 'Chờ duyệt') {
    statusBadge = '⏳';
    statusColor = '#ff9800';
    statusText = 'CHỜ DUYỆT';
  } else if (trangThai === 'Đã duyệt') {
    statusBadge = `<svg class="w-6 h-6 text-gray-800 dark:text-white" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 24 24">
  <path fill-rule="evenodd" d="M12 2c-.791 0-1.55.314-2.11.874l-.893.893a.985.985 0 0 1-.696.288H7.04A2.984 2.984 0 0 0 4.055 7.04v1.262a.986.986 0 0 1-.288.696l-.893.893a2.984 2.984 0 0 0 0 4.22l.893.893a.985.985 0 0 1 .288.696v1.262a2.984 2.984 0 0 0 2.984 2.984h1.262c.261 0 .512.104.696.288l.893.893a2.984 2.984 0 0 0 4.22 0l.893-.893a.985.985 0 0 1 .696-.288h1.262a2.984 2.984 0 0 0 2.984-2.984V15.7c0-.261.104-.512.288-.696l.893-.893a2.984 2.984 0 0 0 0-4.22l-.893-.893a.985.985 0 0 1-.288-.696V7.04a2.984 2.984 0 0 0-2.984-2.984h-1.262a.985.985 0 0 1-.696-.288l-.893-.893A2.984 2.984 0 0 0 12 2Zm3.683 7.73a1 1 0 1 0-1.414-1.413l-4.253 4.253-1.277-1.277a1 1 0 0 0-1.415 1.414l1.985 1.984a1 1 0 0 0 1.414 0l4.96-4.96Z" clip-rule="evenodd"/>
</svg>
`;
    statusColor = '#4caf50';
    statusText = 'ĐÃ DUYỆT';
  } else if (trangThai === 'Từ chối') {
    statusBadge = `<svg class="w-6 h-6 text-gray-800 dark:text-white" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="currentColor" viewBox="0 0 24 24">
  <path fill-rule="evenodd" d="M2 12C2 6.477 6.477 2 12 2s10 4.477 10 10-4.477 10-10 10S2 17.523 2 12Zm7.707-3.707a1 1 0 0 0-1.414 1.414L10.586 12l-2.293 2.293a1 1 0 1 0 1.414 1.414L12 13.414l2.293 2.293a1 1 0 0 0 1.414-1.414L13.414 12l2.293-2.293a1 1 0 0 0-1.414-1.414L12 10.586 9.707 8.293Z" clip-rule="evenodd"/>
</svg>
`;
    statusColor = '#f44336';
    statusText = 'TỪ CHỐI';
  }
  
  // Sử dụng cách xưng hô phù hợp, mặc định là "Quý Anh/Chị"
  var loiChao = cachXungHo ? 'Kính gửi ' + cachXungHo + ',' : 'Kính gửi Quý Anh/Chị,';
  
  var html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        line-height: 1.6;
        color: #333;
        margin: 0;
        padding: 0;
        background-color: #f5f5f5;
      }
      .container {
        max-width: 650px;
        margin: 20px auto;
        background: white;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      }
      .header {
        background: linear-gradient(135deg, ${CONFIG.brandColor} 0%, #008556 100%);
        color: white;
        padding: 30px;
        text-align: center;
      }
      .header h1 {
        margin: 0;
        font-size: 24px;
        font-weight: 600;
      }
      .status-badge {
        display: inline-block;
        padding: 8px 20px;
        background: ${statusColor};
        color: white;
        border-radius: 20px;
        font-weight: bold;
        margin-top: 10px;
      }
      .content {
        padding: 30px;
      }
      .info-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
      }
      .info-table td {
        padding: 12px;
        border-bottom: 1px solid #eee;
      }
      .info-table td:first-child {
        font-weight: 600;
        color: ${CONFIG.brandColor};
        width: 40%;
      }
      .highlight-box {
        background: #f0f9f5;
        border-left: 4px solid ${CONFIG.brandColor};
        padding: 15px;
        margin: 20px 0;
        border-radius: 4px;
      }
      .warning-box {
        background: #fff3e0;
        border-left: 4px solid #ff9800;
        padding: 15px;
        margin: 20px 0;
        border-radius: 4px;
      }
      .danger-box {
        background: #ffebee;
        border-left: 4px solid #f44336;
        padding: 15px;
        margin: 20px 0;
        border-radius: 4px;
      }
      .footer {
        background: #f9f9f9;
        padding: 20px 30px;
        text-align: center;
        font-size: 12px;
        color: #666;
      }
      @media only screen and (max-width: 600px) {
        .container {
          margin: 0;
          border-radius: 0;
        }
        .info-table td {
          display: block;
          width: 100% !important;
        }
        .info-table td:first-child {
          padding-bottom: 5px;
        }
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">
        <h1>${statusBadge} THÔNG BÁO CÔNG TÁC</h1>
        <div class="status-badge">${statusText}</div>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">Mã đăng ký: ${maDangKy}</p>
      </div>
      
      <div class="content">
        <p style="font-size: 16px; color: #333;">${loiChao}</p>
        
        <p>Thông tin đăng ký công tác ${trangThai === 'Chờ duyệt' ? 'mới' : ''}:</p>
        
        <table class="info-table">
          <tr>
            <td>👤 Họ tên</td>
            <td><strong>${formData.hoTen}</strong></td>
          </tr>
          <tr>
            <td>💼 Chức vụ</td>
            <td>${formData.chucVu}</td>
          </tr>
          <tr>
            <td>🏢 Loại phòng</td>
            <td>${formData.phongBan}</td>
          </tr>
          <tr>
            <td>📅 Ngày đến</td>
            <td><strong style="color: ${CONFIG.brandColor}">${formatDate(formData.ngayDen)}</strong></td>
          </tr>
          <tr>
            <td>📅 Ngày đi</td>
            <td><strong style="color: ${CONFIG.brandColor}">${formatDate(formData.ngayDi)}</strong></td>
          </tr>
          <tr>
            <td>🚗 Phương tiện</td>
            <td>${formData.phuongTien}</td>
          </tr>
          <tr>
            <td>🍽️ Nhà ăn</td>
            <td>${formData.nhaAn || 'Không'}</td>
          </tr>
          <tr>
            <td>📍 Địa điểm</td>
            <td><strong>${formData.diaDiem}</strong></td>
          </tr>
        </table>
        
        ${formData.thongTin ? `
        <div class="highlight-box">
          <strong style="color: ${CONFIG.brandColor};">📋 Thông tin cần thiết:</strong>
          <p style="margin: 10px 0 0 0;">${formData.thongTin}</p>
        </div>
        ` : ''}
        
        ${trangThai === 'Chờ duyệt' ? `
        <div class="warning-box">
          <strong style="color: #ff9800;">⏳ Đăng ký đang chờ duyệt</strong>
          <p style="margin: 10px 0 0 0;">Vui lòng kiểm tra và duyệt đăng ký này trên hệ thống.</p>
        </div>
        ` : ''}
        
        ${trangThai === 'Đã duyệt' && ghiChu ? `
        <div class="highlight-box">
          <strong style="color: ${CONFIG.brandColor};"> Ghi chú từ người duyệt:</strong>
          <p style="margin: 10px 0 0 0;">${ghiChu}</p>
        </div>
        ` : ''}
        
        ${trangThai === 'Từ chối' ? `
        <div class="danger-box">
          <strong style="color: #f44336;">Lý do từ chối:</strong>
          <p style="margin: 10px 0 0 0;">${ghiChu || 'Không có lý do cụ thể'}</p>
          <p style="margin: 10px 0 0 0;"><em>Bạn có thể đăng ký lại sau khi điều chỉnh thông tin.</em></p>
        </div>
        ` : ''}
        
        <p style="margin-top: 30px; color: #666;">
          ${trangThai === 'Chờ duyệt' ? 'Vui lòng sắp xếp và chuẩn bị đón tiếp theo thông tin trên.' : ''}
          ${trangThai === 'Đã duyệt' ? 'Đăng ký của bạn đã được xác nhận. Chúc bạn có chuyến công tác hiệu quả!' : ''}
        </p>
      </div>
      
      <div class="footer">
        <p style="margin: 0;">Email này được gửi tự động từ <strong>Hệ thống quản lý công tác THACO AGRI KLH SNUOL</strong></p>
        <p style="margin: 5px 0 0 0;">© ${new Date().getFullYear()} THACO AGRI KLH SNUOL. All rights reserved.</p>
      </div>
    </div>
  </body>
  </html>
  `;
  
  return html;
}

// ============================================
// EXPORT EXCEL
// ============================================
function exportToExcelData(filters) {
  var result = getDangKyList(filters);
  if (!result.success) {
    return result;
  }
  
  return {
    success: true,
    data: result.data,
    sheetName: 'Danh sách công tác',
    filename: 'DanhSachCongTac_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss') + '.xlsx'
  };
}

// ============================================
// SETUP HỆ THỐNG LẦN ĐẦU
// ============================================
function setupSystem() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Cài đặt hệ thống',
    'Bạn có muốn thiết lập các sheet mẫu và công thức tính toán không?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      // Setup các sheet nếu chưa có
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Kiểm tra và tạo sheets
      var sheetNames = Object.values(CONFIG.sheetNames);
      for (var i = 0; i < sheetNames.length; i++) {
        if (!ss.getSheetByName(sheetNames[i])) {
          ss.insertSheet(sheetNames[i]);
        }
      }
      
      ui.alert('Cài đặt thành công!\n\nVui lòng thiết lập cấu trúc dữ liệu theo hướng dẫn.');
      
    } catch (error) {
      ui.alert('Lỗi: ' + error.toString());
    }
  }
}

// ============================================
// XEM THỐNG KÊ
// ============================================
function viewStatistics() {
  var result = getThongKe();
  if (result.success) {
    var data = result.data;
    var message = 
      '📊 THỐNG KÊ HỆ THỐNG\n\n' +
      '📝 Tổng số đăng ký: ' + data.tongSo + '\n' +
      '📅 Đăng ký tháng này: ' + data.thangNay + '\n\n' +
      '⏳ Chờ duyệt: ' + data.choDuyet + '\n' +
      '✅ Đã duyệt: ' + data.daDuyet + '\n' +
      '❌ Từ chối: ' + data.tuChoi + '\n\n' +
      '📍 THEO ĐỊA ĐIỂM:\n' +
      '- Văn phòng 55: ' + data.theoDiaDiem['Văn phòng 55'] + '\n' +
      '- Bình Phước 1: ' + data.theoDiaDiem['Bình Phước 1'] + '\n' +
      '- Bình Phước 2: ' + data.theoDiaDiem['Bình Phước 2'] + '\n' +
      '- ERC: ' + data.theoDiaDiem['ERC'] + '\n' +
      '- Xi nghiệp Bò: ' + data.theoDiaDiem['Xi nghiệp Bò'];
    
    SpreadsheetApp.getUi().alert(message);
  }
}

// ============================================
// DỌN DẸP DỮ LIỆU CŨ (>6 THÁNG)
// ============================================
function cleanOldData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Dọn dẹp dữ liệu',
    'Bạn có muốn xóa các đăng ký cũ hơn 6 tháng không?\n(Chỉ xóa dữ liệu đã đánh dấu xóa)',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(CONFIG.sheetNames.dangKy);
      var data = sheet.getDataRange().getValues();
      
      var sixMonthsAgo = new Date();
      sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
      
      var deletedCount = 0;
      
      // Duyệt từ dưới lên để xóa không ảnh hưởng index
      for (var i = data.length - 1; i > 0; i--) {
        var row = data[i];
        var timestamp = new Date(row[CONFIG.columns.timestamp]);
        var daXoa = row[CONFIG.columns.daXoa];
        
        if (daXoa === true && timestamp < sixMonthsAgo) {
          sheet.deleteRow(i + 1);
          deletedCount++;
        }
      }
      
      ui.alert('Đã xóa ' + deletedCount + ' dòng dữ liệu cũ!');
      
    } catch (error) {
      ui.alert('Lỗi: ' + error.toString());
    }
  }
}


function testSystem() {
  try {
    Logger.log('=== TEST BẮT ĐẦU ===');
    
    // Test 1: Kiểm tra sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = ['Đăng ký công tác', 'Cấu hình Email', 'Danh mục', 'Phân quyền', 'Lịch sử duyệt', 'Thống kê'];
    
    Logger.log('Test 1: Checking sheets...');
    sheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) {
        Logger.log('✅ Sheet "' + name + '" OK');
      } else {
        Logger.log('❌ Sheet "' + name + '" MISSING!');
      }
    });
    
    // Test 2: Kiểm tra user
    Logger.log('\nTest 2: Checking user...');
    var userEmail = Session.getActiveUser().getEmail();
    Logger.log('User email: ' + userEmail);
    
    // Test 3: Kiểm tra functions
    Logger.log('\nTest 3: Checking functions...');
    
    var danhMuc = getDanhMuc();
    Logger.log('getDanhMuc: ' + (danhMuc.success ? '✅ OK' : '❌ FAILED'));
    
    var thongKe = getThongKe();
    Logger.log('getThongKe: ' + (thongKe.success ? '✅ OK' : '❌ FAILED'));
    
    var userInfo = getUserInfo();
    Logger.log('getUserInfo: ' + (userInfo.success ? '✅ OK' : '❌ FAILED'));
    
    Logger.log('\n=== TEST HOÀN THÀNH ===');
    
  } catch (error) {
    Logger.log('❌ LỖI: ' + error.toString());
  }
}

// ============================================
// HÀM DEBUG - KIỂM TRA EMAIL & QUYỀN
// ============================================

function debugUserInfo() {
  var userEmail = Session.getActiveUser().getEmail();
  var effectiveEmail = Session.getEffectiveUser().getEmail();
  var userRole = getUserRole(userEmail);
  var managedLocations = getManagedLocations(userEmail);
  
  Logger.log('=== DEBUG USER INFO ===');
  Logger.log('Active User Email: ' + userEmail);
  Logger.log('Effective User Email: ' + effectiveEmail);
  Logger.log('User Role: ' + userRole);
  Logger.log('Managed Locations: ' + JSON.stringify(managedLocations));
  
  // Kiểm tra trong sheet Phân quyền
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Phân quyền');
  var data = sheet.getDataRange().getValues();
  
  Logger.log('\n=== DANH SÁCH PHÂN QUYỀN ===');
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === 'Active') {
      Logger.log('Email: ' + data[i][0] + ' | Vai trò: ' + data[i][2] + ' | Địa điểm: ' + data[i][3]);
    }
  }
  
  return {
    activeUser: userEmail,
    effectiveUser: effectiveEmail,
    role: userRole,
    locations: managedLocations
  };
}

// Hàm test cho Dashboard
function testGetUserInfo() {
  var result = getUserInfo();
  Logger.log('=== TEST getUserInfo() ===');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
// ============================================
// KIỂM TRA TÍNH ĐỒNG BỘ TÊN ĐỊA ĐIỂM
// ============================================

function kiemTraDongBoDiaDiem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Lấy địa điểm từ "Cấu hình Email"
  var sheetEmail = ss.getSheetByName('Cấu hình Email');
  var dataEmail = sheetEmail.getDataRange().getValues();
  var diaDiemEmail = {};
  
  Logger.log('=== ĐỊA ĐIỂM TRONG "CẤU HÌNH EMAIL" ===');
  for (var i = 1; i < dataEmail.length; i++) {
    if (dataEmail[i][0] && dataEmail[i][5] === 'Active') {
      var dd = dataEmail[i][0];
      diaDiemEmail[dd] = true;
      Logger.log('✓ "' + dd + '"');
    }
  }
  
  // Lấy địa điểm từ "Phân quyền"
  var sheetPQ = ss.getSheetByName('Phân quyền');
  var dataPQ = sheetPQ.getDataRange().getValues();
  var diaDiemPQ = {};
  
  Logger.log('\n=== ĐỊA ĐIỂM TRONG "PHÂN QUYỀN" ===');
  for (var i = 1; i < dataPQ.length; i++) {
    if (dataPQ[i][3] && dataPQ[i][3] !== 'Tất cả' && dataPQ[i][4] === 'Active') {
      var dd = dataPQ[i][3];
      diaDiemPQ[dd] = true;
      Logger.log('✓ "' + dd + '"');
    }
  }
  
  // Lấy địa điểm từ "Danh mục"
  var sheetDM = ss.getSheetByName('Danh mục');
  var dataDM = sheetDM.getDataRange().getValues();
  var diaDiemDM = {};
  
  Logger.log('\n=== ĐỊA ĐIỂM TRONG "DANH MỤC" ===');
  for (var i = 1; i < dataDM.length; i++) {
    if (dataDM[i][0]) {
      var dd = dataDM[i][0];
      diaDiemDM[dd] = true;
      Logger.log('✓ "' + dd + '"');
    }
  }
  
  // So sánh
  Logger.log('\n=== KIỂM TRA TÍNH ĐỒNG BỘ ===');
  
  var allDiaDiem = {};
  for (var dd in diaDiemEmail) allDiaDiem[dd] = true;
  for (var dd in diaDiemPQ) allDiaDiem[dd] = true;
  for (var dd in diaDiemDM) allDiaDiem[dd] = true;
  
  var hasError = false;
  
  for (var dd in allDiaDiem) {
    var inEmail = diaDiemEmail[dd] ? '✓' : '✗';
    var inPQ = diaDiemPQ[dd] ? '✓' : '✗';
    var inDM = diaDiemDM[dd] ? '✓' : '✗';
    
    var status = (inEmail === '✓' && inPQ === '✓' && inDM === '✓') ? '✅ OK' : '❌ THIẾU';
    
    Logger.log('"' + dd + '": Email[' + inEmail + '] PQ[' + inPQ + '] DM[' + inDM + '] → ' + status);
    
    if (status.indexOf('❌') !== -1) {
      hasError = true;
    }
  }
  
  Logger.log('\n=== KẾT QUẢ ===');
  if (hasError) {
    Logger.log('CÒN LỖI: Tên địa điểm chưa đồng bộ!');
    Logger.log('→ Hãy sửa cho tất cả địa điểm có dấu ✗');
  } else {
    Logger.log('HOÀN HẢO: Tất cả địa điểm đã đồng bộ!');
  }
}

// ============================================
// TEST ĐỌC DỮ LIỆU ĐĂNG KÝ
// ============================================

function kiemTraDuLieuDangKy() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Đăng ký công tác');
  
  if (!sheet) {
    Logger.log('KHÔNG TÌM THẤY sheet "Đăng ký công tác"!');
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  
  Logger.log('=== KIỂM TRA DỮ LIỆU ĐĂNG KÝ ===');
  Logger.log('Tổng số dòng (bao gồm header): ' + data.length);
  
  if (data.length <= 1) {
    Logger.log('Sheet chỉ có header, CHƯA CÓ ĐĂNG KÝ NÀO!');
    Logger.log('→ Hãy tạo đăng ký thử để test hệ thống');
    return;
  }
  
  Logger.log('\n=== DANH SÁCH ĐĂNG KÝ ===');
  
  var chuaDuyet = 0;
  var daDuyet = 0;
  var tuChoi = 0;
  var daXoa = 0;
  
  for (var i = 1; i < data.length; i++) {
    var maDangKy = data[i][1];
    var hoTen = data[i][2];
    var diaDiem = data[i][10];
    var trangThai = data[i][12];
    var isDeleted = data[i][20];
    
    if (isDeleted === true) {
      daXoa++;
      continue;
    }
    
    Logger.log((i) + '. ' + maDangKy + ' | ' + hoTen + ' | ' + diaDiem + ' | ' + trangThai);
    
    if (trangThai === 'Chờ duyệt') chuaDuyet++;
    else if (trangThai === 'Đã duyệt') daDuyet++;
    else if (trangThai === 'Từ chối') tuChoi++;
  }
  
  Logger.log('\n=== THỐNG KÊ ===');
  Logger.log('Tổng đăng ký: ' + (data.length - 1 - daXoa));
  Logger.log('Chờ duyệt: ' + chuaDuyet);
  Logger.log('Đã duyệt: ' + daDuyet);
  Logger.log('Từ chối: ' + tuChoi);
  Logger.log('Đã xóa: ' + daXoa);
}

// ============================================
// TEST PHÂN QUYỀN CỦA USER
// ============================================

function kiemTraQuyenCuaToi() {
  var userEmail = Session.getActiveUser().getEmail();
  var userRole = getUserRole(userEmail);
  var managedLocations = getManagedLocations(userEmail);
  
  Logger.log('=== THÔNG TIN QUYỀN CỦA BẠN ===');
  Logger.log('Email: ' + userEmail);
  Logger.log('Vai trò: ' + userRole);
  Logger.log('Địa điểm quản lý: ' + JSON.stringify(managedLocations));
  
  Logger.log('\n=== QUYỀN CỤ THỂ ===');
  
  if (userRole === 'Admin') {
    Logger.log('Xem: TẤT CẢ đăng ký');
    Logger.log('Sửa: TẤT CẢ đăng ký');
    Logger.log('Xóa: TẤT CẢ đăng ký');
    Logger.log('Duyệt: TẤT CẢ địa điểm');
  } else if (userRole === 'Approver') {
    Logger.log('Xem: Đăng ký thuộc ' + managedLocations.join(', '));
    Logger.log('Sửa: Chỉ đăng ký của mình (khi Chờ duyệt)');
    Logger.log('Xóa: Chỉ đăng ký của mình (khi Chờ duyệt)');
    Logger.log('Duyệt: Chỉ địa điểm ' + managedLocations.join(', '));
  } else {
    Logger.log('Xem: Chỉ đăng ký của mình');
    Logger.log('Sửa: Chỉ đăng ký của mình (khi Chờ duyệt)');
    Logger.log('Xóa: Chỉ đăng ký của mình (khi Chờ duyệt)');
    Logger.log('Duyệt: KHÔNG có quyền');
  }
  
  // Test xem có thể xem những đăng ký nào
  Logger.log('\n=== TEST ĐỌC DANH SÁCH ===');
  var result = getDangKyList({});
  
  if (result.success) {
    Logger.log('Số đăng ký bạn có thể xem: ' + result.total);
    
    if (result.total > 0) {
      Logger.log('\nDanh sách:');
      for (var i = 0; i < Math.min(5, result.data.length); i++) {
        var item = result.data[i];
        Logger.log((i+1) + '. ' + item.maDangKy + ' | ' + item.hoTen + ' | ' + item.diaDiem + ' | ' + item.trangThai);
      }
      if (result.total > 5) {
        Logger.log('... và ' + (result.total - 5) + ' đăng ký khác');
      }
    } else {
      Logger.log('Không tìm thấy đăng ký nào bạn có quyền xem!');
    }
  } else {
    Logger.log('Lỗi: ' + result.error);
  }
}
