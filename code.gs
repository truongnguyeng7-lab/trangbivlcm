/****************************
 * MENU & DIALOG
 ****************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("VLCM 2026")
    .addItem("Tra Cứu", "showSearchDialog")
    .addSeparator()
    .addItem("Tạo/kiểm tra Sheet yêu cầu VIP", "setupVipRequestSheet")
    .addSeparator()
    .addItem("Duyệt VIP dòng đang chọn", "approveSelectedVipRequest")
    .addToUi();
}

function setupVipRequestSheet() {
  const sheet = getVipRequestSheet_();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert("Đã tạo/kiểm tra sheet VIP_REQUESTS.");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showSearchDialog() {
  const html = HtmlService.createTemplateFromFile("Dialog")
    .evaluate()
    .setWidth(1200)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, " ");
}

function doGet(e) {
  e = e || {};
  const params = e.parameter || {};

  if (params.action === "approveVip") {
    return renderVipApprovePage_(params.token || "");
  }

  return HtmlService.createTemplateFromFile("Dialog")
    .evaluate()
    .setTitle("VLCM 2026");
}


/****************************
 * VIP LOGIN ACCESS – LẤY TÀI KHOẢN TỪ SHEET
 ****************************/


const VIP_ADMIN_EMAIL = "truongnguyen.g7@gmail.com";
const VIP_REQUEST_SHEET_NAME = "VIP_REQUESTS";

function getVipRequestSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(VIP_REQUEST_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(VIP_REQUEST_SHEET_NAME);
    sheet.setTabColor("#ff9800");
  }

  const headers = [
    "Thời gian",
    "Email Google đang dùng",
    "Email liên hệ",
    "Số điện thoại",
    "Trạng thái",
    "Gói đăng ký",
    "Xác nhận thanh toán",
    "Email VIP đã cấp",
    "Mật khẩu VIP đã cấp",
    "Thời gian duyệt",
    "Ghi chú",
    "Mã duyệt điện thoại",
    "Link duyệt điện thoại"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#202020")
    .setFontColor("#ffffff");

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 240);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 260);
  sheet.setColumnWidth(9, 180);
  sheet.setColumnWidth(10, 170);
  sheet.setColumnWidth(11, 260);
  sheet.setColumnWidth(12, 260);
  sheet.setColumnWidth(13, 420);

  return sheet;
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function checkVipEmailExists_(email) {
  const normalizedEmail = String(email || "").trim().toLowerCase();
  const sheet = getVipRequestSheet_();
  const lastRow = sheet.getLastRow();

  if (!normalizedEmail || lastRow < 2) {
    return {
      exists: false,
      status: "",
      row: 0
    };
  }

  const dataRows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

  // Duyệt từ dưới lên để lấy lần đăng ký mới nhất
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];

    const contactEmail = String(row[2] || "").trim().toLowerCase(); // Cột C
    const status = String(row[4] || "").trim(); // Cột E
    const vipEmail = String(row[7] || "").trim().toLowerCase(); // Cột H

    if (contactEmail === normalizedEmail || vipEmail === normalizedEmail) {
      return {
        exists: true,
        status: status,
        row: i + 2
      };
    }
  }

  return {
    exists: false,
    status: "",
    row: 0
  };
}


function submitVipContactRequest(form) {
  form = form || {};

  const googleEmail = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  const contactEmail = String(form.contactEmail || "").trim().toLowerCase();
  const phone = String(form.phone || "").trim();
  const paymentConfirmed = form.paymentConfirmed === true || String(form.paymentConfirmed) === "true";
  const packageName = String(form.packageName || "VIP 50K VĨNH VIỄN").trim();
  const status = paymentConfirmed ? "Khách báo đã thanh toán - Chờ admin kiểm tra" : "Chờ xử lý";

  if (!contactEmail || !contactEmail.includes("@")) {
    throw new Error("Email liên hệ không hợp lệ.");
  }

  if (!paymentConfirmed) {
    throw new Error("Vui lòng xác nhận đã thanh toán 50K.");
  }

  const sheet = getVipRequestSheet_();
  
  const approveToken = generateVipApproveToken_();
  const approveLink = buildVipApproveLink_(approveToken);
  
  const existed = checkVipEmailExists_(contactEmail);

if (existed.exists) {
  return {
    ok: false,
    code: "EMAIL_EXISTS",
    message: "Email này đã đăng ký VIP. Vui lòng đăng nhập lại hoặc dùng chức năng quên mật khẩu."
  };
}


  sheet.appendRow([
  new Date(),
  googleEmail || "Không lấy được email Google",
  contactEmail,
  phone,
  status,
  packageName,
  "Khách đã tick xác nhận",
  "",
  "",
  "",
  "Chưa duyệt",
  approveToken,
  approveLink
]);

  const subject = "[VLCM] Yêu cầu cấp VIP mới - Chờ kiểm tra thanh toán";

  const plainBody =
    "Có yêu cầu cấp VIP mới:\n\n" +
    "Email Google đang dùng: " + (googleEmail || "Không lấy được") + "\n" +
    "Email liên hệ: " + contactEmail + "\n" +
    "Số điện thoại: " + phone + "\n" +
    "Gói đăng ký: " + packageName + "\n" +
    "Xác nhận thanh toán: KHÁCH BÁO ĐÃ THANH TOÁN 50K\n" +
    "Thời gian: " + new Date() + "\n\n" +
    "Sau khi kiểm tra đã nhận tiền, bạn có thể duyệt bằng 1 trong 2 cách:\n\n" +
"Cách 1: Vào sheet VIP_REQUESTS, chọn đúng dòng khách rồi bấm VLCM 2026 > Duyệt VIP dòng đang chọn.\n" +
"Cách 2: Bấm link duyệt trên điện thoại:\n" + approveLink;

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#222">' +
      '<h2 style="margin:0 0 12px;color:#0078d4">Yêu cầu cấp VIP mới</h2>' +
      '<table cellpadding="8" cellspacing="0" style="border-collapse:collapse;border:1px solid #ddd">' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Email Google đang dùng</td><td style="border:1px solid #ddd">' + escapeHtml_(googleEmail || "Không lấy được") + '</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Email liên hệ</td><td style="border:1px solid #ddd">' + escapeHtml_(contactEmail) + '</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Số điện thoại</td><td style="border:1px solid #ddd">' + escapeHtml_(phone) + '</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Gói đăng ký</td><td style="border:1px solid #ddd">' + escapeHtml_(packageName) + '</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Trạng thái</td><td style="border:1px solid #ddd;color:#f57c00;font-weight:bold">Chờ admin kiểm tra thanh toán</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Thời gian</td><td style="border:1px solid #ddd">' + escapeHtml_(new Date()) + '</td></tr>' +
      '</table>' +
      '<p style="margin-top:14px">Sau khi đã nhận tiền, bạn có thể duyệt bằng 1 trong 2 cách:</p>' +
'<p><b>Cách 1:</b> Vào sheet <b>VIP_REQUESTS</b>, chọn dòng khách rồi bấm <b>VLCM 2026 → Duyệt VIP dòng đang chọn</b>.</p>' +
'<p><b>Cách 2:</b> Bấm nút bên dưới để duyệt trực tiếp trên điện thoại.</p>' +
'<p style="margin-top:14px">' +
  '<a href="' + escapeHtml_(approveLink) + '" style="display:inline-block;background:#0078d4;color:#fff;text-decoration:none;padding:12px 18px;border-radius:8px;font-weight:bold">' +
    '✅ Duyệt VIP trên điện thoại' +
  '</a>' +
'</p>' +
    '</div>';

  MailApp.sendEmail({
    to: VIP_ADMIN_EMAIL,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    replyTo: contactEmail,
    name: "VLCM VIP Request"
  });

  return {
    ok: true,
    emailSentTo: VIP_ADMIN_EMAIL,
    message: "Đã gửi yêu cầu đến admin. Sau khi admin kiểm tra thanh toán, tài khoản VIP sẽ được gửi về email của bạn."
  };
}

function generateVipPassword_() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789";
  let password = "VIP-";

  for (let i = 0; i < 8; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }

  return password;
}


function sendVipAccountEmailToCustomer_(vipAccount, packageName) {
  const subject = "[VLCM] Tài khoản VIP của bạn";

  const plainBody =
    "Chào bạn,\n\n" +
    "Thanh toán của bạn đã được admin xác nhận.\n" +
    "Tài khoản VIP của bạn đã được tạo thành công.\n\n" +
    "Email đăng nhập: " + vipAccount.email + "\n" +
    "Mật khẩu: " + vipAccount.password + "\n" +
    "Gói: " + packageName + "\n\n" +
    "Vui lòng vào trang http://vlcm-trangbi.gamer.gd/ và chọn ĐĂNG NHẬP VIP để sử dụng.\n\n" +
    "Cảm ơn bạn đã đăng ký VIP.";

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#222">' +
      '<h2 style="margin:0 0 12px;color:#0078d4">Tài khoản VIP của bạn</h2>' +
      '<p>Chào bạn, thanh toán của bạn đã được admin xác nhận.</p>' +
      '<p>Tài khoản VIP của bạn đã được tạo thành công:</p>' +
      '<table cellpadding="8" cellspacing="0" style="border-collapse:collapse;border:1px solid #ddd">' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Email đăng nhập</td><td style="border:1px solid #ddd;color:#0078d4;font-weight:bold">' + escapeHtml_(vipAccount.email) + '</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Mật khẩu</td><td style="border:1px solid #ddd;color:#d32f2f;font-weight:bold;font-size:16px">' + escapeHtml_(vipAccount.password) + '</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Gói</td><td style="border:1px solid #ddd">' + escapeHtml_(packageName) + '</td></tr>' +
      '</table>' +
      '<p style="margin-top:14px">Vui lòng vào trang http://vlcm-trangbi.gamer.gd/ và chọn <b>ĐĂNG NHẬP VIP</b> để sử dụng.</p>' +
      '<p>Cảm ơn bạn đã đăng ký VIP.</p>' +
    '</div>';

  MailApp.sendEmail({
    to: vipAccount.email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    name: "VLCM VIP"
  });
}

function generateVipApproveToken_() {
  return Utilities.getUuid().replace(/-/g, "") + "_" + Date.now();
}

function buildVipApproveLink_(token) {
  const url = ScriptApp.getService().getUrl();

  if (!url) {
    return "";
  }

  return url + "?action=approveVip&token=" + encodeURIComponent(token);
}

function findVipRequestByToken_(token) {
  token = String(token || "").trim();

  if (!token) {
    return null;
  }

  const sheet = getVipRequestSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const dataRows = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];
    const rowToken = String(row[11] || "").trim(); // Cột L

    if (rowToken === token) {
      return {
        sheet: sheet,
        rowNumber: i + 2,
        rowData: row
      };
    }
  }

  return null;
}

function getVipApproveInfoByToken(token) {
  const found = findVipRequestByToken_(token);

  if (!found) {
    return {
      ok: false,
      message: "Link duyệt không hợp lệ hoặc đã hết hạn."
    };
  }

  const row = found.rowData;

  return {
    ok: true,
    rowNumber: found.rowNumber,
    contactEmail: String(row[2] || "").trim(),
    phone: String(row[3] || "").trim(),
    status: String(row[4] || "").trim(),
    packageName: String(row[5] || "VIP 50K VĨNH VIỄN").trim(),
    paymentStatus: String(row[6] || "").trim()
  };
}

function approveVipRequestByToken(token) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const found = findVipRequestByToken_(token);

    if (!found) {
      return {
        ok: false,
        message: "Link duyệt không hợp lệ hoặc đã được sử dụng."
      };
    }

    const sheet = found.sheet;
    const row = found.rowNumber;
    const rowData = found.rowData;

    const contactEmail = String(rowData[2] || "").trim().toLowerCase();
    const phone = String(rowData[3] || "").trim();
    const currentStatus = String(rowData[4] || "").trim();
    const packageName = String(rowData[5] || "VIP 50K VĨNH VIỄN").trim();
    const oldPassword = String(rowData[8] || "").trim();

    if (!contactEmail || !contactEmail.includes("@")) {
      return {
        ok: false,
        message: "Dòng đăng ký này không có email hợp lệ."
      };
    }

    if (currentStatus === "Đã duyệt VIP - Đã gửi tài khoản") {
      return {
        ok: false,
        message: "Tài khoản này đã được duyệt trước đó."
      };
    }

    const password = oldPassword || generateVipPassword_();

    const vipAccount = {
      email: contactEmail,
      password: password
    };

    sendVipAccountEmailToCustomer_(vipAccount, packageName);

    sheet.getRange(row, 5).setValue("Đã duyệt VIP - Đã gửi tài khoản");
    sheet.getRange(row, 8).setValue(contactEmail);
    sheet.getRange(row, 9).setValue(password);
    sheet.getRange(row, 10).setValue(new Date());
    sheet.getRange(row, 11).setValue(
      "Admin đã duyệt bằng điện thoại - Đã gửi tài khoản VIP cho khách" +
      (phone ? " - SĐT: " + phone : "")
    );

    // Vô hiệu hóa token sau khi duyệt, tránh bấm lại nhiều lần
    sheet.getRange(row, 12).setValue("USED-" + token);

    return {
      ok: true,
      email: contactEmail,
      password: password,
      message: "Đã duyệt gói VIP thành công và đã gửi tài khoản đăng nhập về email khách."
    };

  } catch (err) {
    return {
      ok: false,
      message: "Lỗi duyệt VIP: " + (err && err.message ? err.message : String(err))
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

function renderVipApprovePage_(token) {
  token = String(token || "").trim();

  const safeToken = JSON.stringify(token);

  const html =
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
      '<base target="_top">' +
      '<meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<style>' +
        'body{margin:0;background:#0f0f0f;color:#fff;font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:18px;box-sizing:border-box}' +
        '.card{width:430px;max-width:100%;background:#1f1f1f;border:1px solid #3a3a3a;border-radius:18px;padding:22px;box-shadow:0 18px 50px rgba(0,0,0,.55)}' +
        '.kicker{color:#4cc2ff;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px}' +
        'h1{font-size:22px;margin:0 0 12px}' +
        '.info{background:#121212;border:1px solid #333;border-radius:12px;padding:12px;margin:14px 0;color:#ddd;font-size:14px;line-height:1.7}' +
        '.row b{color:#9fdfff}' +
        '.btn{width:100%;border:0;border-radius:12px;padding:14px 16px;font-weight:900;font-size:15px;cursor:pointer;margin-top:10px}' +
        '.primary{background:#0078d4;color:#fff}' +
        '.primary:disabled{opacity:.55;cursor:not-allowed}' +
        '.secondary{background:#333;color:#ddd}' +
        '.msg{display:none;margin-top:14px;padding:12px;border-radius:12px;font-size:14px;line-height:1.5}' +
        '.ok{background:rgba(129,201,149,.12);border:1px solid rgba(129,201,149,.35);color:#9ee6b2}' +
        '.err{background:rgba(255,107,107,.12);border:1px solid rgba(255,107,107,.35);color:#ff9c9c}' +
      '</style>' +
    '</head>' +
    '<body>' +
      '<div class="card">' +
        '<div class="kicker">VLCM VIP</div>' +
        '<h1>Duyệt VIP bằng điện thoại</h1>' +
        '<div id="info" class="info">Đang tải thông tin đăng ký...</div>' +
        '<button id="approveBtn" class="btn primary" onclick="approveNow()">Xác nhận đã nhận thanh toán & duyệt VIP</button>' +
        '<button class="btn secondary" onclick="window.close()">Đóng</button>' +
        '<div id="msg" class="msg"></div>' +
      '</div>' +

      '<script>' +
        'const TOKEN=' + safeToken + ';' +

        'function showMsg(text, ok){' +
          'const box=document.getElementById("msg");' +
          'box.innerText=text;' +
          'box.className="msg " + (ok ? "ok" : "err");' +
          'box.style.display="block";' +
        '}' +

        'function loadInfo(){' +
          'google.script.run.withSuccessHandler(function(res){' +
            'const info=document.getElementById("info");' +
            'const btn=document.getElementById("approveBtn");' +
            'if(!res || !res.ok){' +
              'info.innerHTML=(res&&res.message)?res.message:"Không tải được thông tin.";' +
              'btn.disabled=true;' +
              'return;' +
            '}' +
            'info.innerHTML=' +
              '"<div class=\\"row\\"><b>Email:</b> " + res.contactEmail + "</div>" +' +
              '"<div class=\\"row\\"><b>SĐT:</b> " + (res.phone || "Không có") + "</div>" +' +
              '"<div class=\\"row\\"><b>Gói:</b> " + res.packageName + "</div>" +' +
              '"<div class=\\"row\\"><b>Trạng thái:</b> " + res.status + "</div>" +' +
              '"<div class=\\"row\\"><b>Thanh toán:</b> " + res.paymentStatus + "</div>";' +
          '}).getVipApproveInfoByToken(TOKEN);' +
        '}' +

        'function approveNow(){' +
          'const btn=document.getElementById("approveBtn");' +
          'btn.disabled=true;' +
          'btn.innerText="Đang duyệt...";' +
          'google.script.run.withSuccessHandler(function(res){' +
            'btn.innerText="Xác nhận đã nhận thanh toán & duyệt VIP";' +
            'if(res && res.ok){' +
              'showMsg(res.message + "\\nEmail: " + res.email + "\\nMật khẩu: " + res.password, true);' +
              'btn.disabled=true;' +
              'loadInfo();' +
              'return;' +
            '}' +
            'btn.disabled=false;' +
            'showMsg((res&&res.message)?res.message:"Duyệt thất bại.", false);' +
          '}).withFailureHandler(function(err){' +
            'btn.disabled=false;' +
            'btn.innerText="Xác nhận đã nhận thanh toán & duyệt VIP";' +
            'showMsg("Lỗi: " + (err.message || err), false);' +
          '}).approveVipRequestByToken(TOKEN);' +
        '}' +

        'loadInfo();' +
      '</script>' +
    '</body>' +
    '</html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle("Duyệt VIP");
}


function requestVipPasswordReset(form) {
  form = form || {};

  const email = String(form.email || "").trim().toLowerCase();

  if (!email || !email.includes("@")) {
    return {
      ok: false,
      code: "INVALID_EMAIL",
      message: "Vui lòng nhập địa chỉ email hợp lệ."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VIP_REQUEST_SHEET_NAME);

  if (!sheet) {
    return {
      ok: false,
      code: "SYSTEM_ERROR",
      message: "Lỗi hệ thống: Không tìm thấy sheet VIP_REQUESTS."
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      ok: false,
      code: "NOT_REGISTERED",
      message: "Tài khoản chưa đăng ký VIP. Vui lòng đăng ký tài khoản trước."
    };
  }

  const dataRows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

  let latestMatchedRow = 0;
  let approvedRow = 0;
  let packageName = "VIP 50K VĨNH VIỄN";
  let latestStatus = "";

  // Duyệt từ dưới lên để lấy yêu cầu mới nhất
  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];

    const contactEmail = String(row[2] || "").trim().toLowerCase();
    const status = String(row[4] || "").trim();
    const currentPackageName = String(row[5] || "VIP 50K VĨNH VIỄN").trim();
    const vipEmail = String(row[7] || "").trim().toLowerCase();

    const loginEmail = vipEmail || contactEmail;

    if (loginEmail !== email && contactEmail !== email) {
      continue;
    }

    if (!latestMatchedRow) {
      latestMatchedRow = i + 2;
      latestStatus = status;
      packageName = currentPackageName;
    }

    if (status === "Đã duyệt VIP - Đã gửi tài khoản") {
      approvedRow = i + 2;
      packageName = currentPackageName;
      break;
    }
  }

  if (!latestMatchedRow) {
    logVipPasswordReset_(email, "TỪ CHỐI", "Email chưa đăng ký VIP", "");
    return {
      ok: false,
      code: "NOT_REGISTERED",
      message: "Tài khoản chưa đăng ký VIP. Vui lòng đăng ký tài khoản trước."
    };
  }

  if (!approvedRow) {
    logVipPasswordReset_(email, "TỪ CHỐI", "Email đã đăng ký nhưng chưa được duyệt: " + latestStatus, latestMatchedRow);
    return {
      ok: false,
      code: "NOT_APPROVED",
      message: "Email này đã đăng ký nhưng chưa được admin duyệt thanh toán. Vui lòng chờ admin kiểm tra."
    };
  }

  const newPassword = generateVipPassword_();

  // Cột H: Email VIP đã cấp
  // Cột I: Mật khẩu VIP đã cấp
  // Cột K: Ghi chú
  sheet.getRange(approvedRow, 8).setValue(email);
  sheet.getRange(approvedRow, 9).setValue(newPassword);
  sheet.getRange(approvedRow, 11).setValue("Khách yêu cầu cấp lại mật khẩu mới lúc " + new Date());

  sendVipPasswordResetEmailToCustomer_({
    email: email,
    password: newPassword
  }, packageName);

  logVipPasswordReset_(email, "THÀNH CÔNG", "Đã cấp lại mật khẩu mới", approvedRow);

  return {
    ok: true,
    code: "RESET_SENT",
    message: "Đã cấp lại mật khẩu mới và gửi về email đăng ký VIP."
  };
}

function sendVipPasswordResetEmailToCustomer_(vipAccount, packageName) {
  const subject = "[VLCM] Cấp lại mật khẩu VIP";

  const plainBody =
    "Chào bạn,\n\n" +
    "Hệ thống đã cấp lại mật khẩu VIP mới cho tài khoản của bạn.\n\n" +
    "Email đăng nhập: " + vipAccount.email + "\n" +
    "Mật khẩu mới: " + vipAccount.password + "\n" +
    "Gói: " + packageName + "\n\n" +
    "Vui lòng vào trang http://vlcm-trangbi.gamer.gd/ và chọn ĐĂNG NHẬP VIP để sử dụng.\n\n" +
    "Nếu bạn không yêu cầu cấp lại mật khẩu, vui lòng liên hệ admin ngay.";

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#222">' +
      '<h2 style="margin:0 0 12px;color:#0078d4">Cấp lại mật khẩu VIP</h2>' +
      '<p>Chào bạn, hệ thống đã cấp lại mật khẩu VIP mới cho tài khoản của bạn.</p>' +
      '<table cellpadding="8" cellspacing="0" style="border-collapse:collapse;border:1px solid #ddd">' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Email đăng nhập</td><td style="border:1px solid #ddd;color:#0078d4;font-weight:bold">' + escapeHtml_(vipAccount.email) + '</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Mật khẩu mới</td><td style="border:1px solid #ddd;color:#d32f2f;font-weight:bold;font-size:16px">' + escapeHtml_(vipAccount.password) + '</td></tr>' +
        '<tr><td style="border:1px solid #ddd;font-weight:bold;background:#f5f5f5">Gói</td><td style="border:1px solid #ddd">' + escapeHtml_(packageName) + '</td></tr>' +
      '</table>' +
      '<p style="margin-top:14px">Vui lòng vào trang http://vlcm-trangbi.gamer.gd/ và chọn <b>ĐĂNG NHẬP VIP</b> để sử dụng.</p>' +
      '<p style="color:#d32f2f"><b>Nếu bạn không yêu cầu cấp lại mật khẩu, vui lòng liên hệ admin ngay.</b></p>' +
    '</div>';

  MailApp.sendEmail({
    to: vipAccount.email,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    name: "VLCM VIP"
  });
}

function logVipPasswordReset_(email, status, note, requestRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("VIP_Reset_Password");

  if (!sheet) {
    sheet = ss.insertSheet("VIP_Reset_Password");
    sheet.appendRow([
      "Thời gian",
      "Email",
      "Trạng thái",
      "Ghi chú",
      "Dòng trong VIP_REQUESTS"
    ]);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    email,
    status,
    note,
    requestRow
  ]);
}


function approveSelectedVipRequest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== VIP_REQUEST_SHEET_NAME) {
    ui.alert("Vui lòng mở sheet VIP_REQUESTS, chọn đúng dòng khách cần duyệt rồi bấm lại.");
    return;
  }

  const row = sheet.getActiveRange().getRow();

  if (row <= 1) {
    ui.alert("Vui lòng chọn dòng dữ liệu của khách, không chọn dòng tiêu đề.");
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, 11).getValues()[0];

  const contactEmail = String(rowData[2] || "").trim().toLowerCase();
  const phone = String(rowData[3] || "").trim();
  const currentStatus = String(rowData[4] || "").trim();
  const packageName = String(rowData[5] || "VIP 50K VĨNH VIỄN").trim();
  const oldPassword = String(rowData[8] || "").trim();

  if (!contactEmail || !contactEmail.includes("@")) {
    ui.alert("Dòng này không có email liên hệ hợp lệ.");
    return;
  }

  if (currentStatus === "Đã duyệt VIP - Đã gửi tài khoản") {
    ui.alert("Dòng này đã được duyệt và đã gửi tài khoản rồi.");
    return;
  }

  const confirm = ui.alert(
    "Xác nhận duyệt VIP",
    "Bạn đã kiểm tra và chắc chắn đã nhận thanh toán của:\n\n" +
    contactEmail + "\n\n" +
    "Bấm OK để tạo tài khoản VIP và gửi email mật khẩu cho khách.",
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) {
    return;
  }

  const password = oldPassword || generateVipPassword_();

  const vipAccount = {
    email: contactEmail,
    password: password
  };

  sendVipAccountEmailToCustomer_(vipAccount, packageName);

  sheet.getRange(row, 5).setValue("Đã duyệt VIP - Đã gửi tài khoản");
  sheet.getRange(row, 8).setValue(contactEmail);
  sheet.getRange(row, 9).setValue(password);
  sheet.getRange(row, 10).setValue(new Date());
  sheet.getRange(row, 11).setValue(
    "Admin đã kiểm tra thanh toán - Đã gửi tài khoản VIP cho khách" +
    (phone ? " - SĐT: " + phone : "")
  );

  ui.alert(
    "Đã duyệt VIP thành công.\n\n" +
    "Email đăng nhập: " + contactEmail + "\n" +
    "Mật khẩu: " + password + "\n\n" +
    "Hệ thống đã gửi tài khoản về email khách."
  );
}

/****************************
 * UTIL
 ****************************/
function removeVietnameseTones(str) {
  return String(str || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase();
}

// escape regex, giữ * làm wildcard
function buildRegexFromKeyword_(kw) {
  let base = removeVietnameseTones(kw.trim());
  base = base.replace(/[-/\\^$+?.()|[\]{}]/g, "\\$&");
  base = base.replace(/\*/g, ".*");
  return base;
}

const STOP_WORDS = [
  "co", "the", "bi", "nhan",
  "trang", "thai",
  "hieu", "qua",
  "moi", "toan", "bo",
  "ngay", "lap", "tuc",
  "phan", "tram"
];

const SEMANTIC_MAP = {
  MIEN_DICH: [
    "mien dich",
    "mien 100% sat thuong",
    "mien dich 100% sat thuong",
    "mien toan bo 100% sat thuong"
  ],

  BAT_LOI: [
    "khong the nhan bat loi",
    "khong the nhan hieu qua bat loi",
    "khong the nhan trang thai bat loi",
    "khong the nhan bat ky bat loi"
  ],

  BUFF_TAY: [
    "khi su dung ky nang",
    "khi thi trien ky nang",
    "khi ban than su dung"
  ],

  NHAY_LAN: [
    "khi nhay(\\s+.*)?(\\s+\\d+)?\\s+lan",
    "khi nhay nhieu lan",
    "khi nhay lien tuc"
  ]
};

function tokenizeKeyword_(text) {
  return removeVietnameseTones(text)
    .split(/\s+/)
    .filter(w => w.length >= 3 && !STOP_WORDS.includes(w));
}

/****************************
 * DATA
 ****************************/
function isSystemSheet_(name) {
  const upper = String(name || "").toUpperCase();

  return (
    upper === "VIP_EMAILS" ||
    upper === "VIP_REQUESTS" ||
    upper === "VIP_ACCESS" ||
    upper === "VIP_FREE_USERS" ||
    upper === "VIP_SEARCH" ||
    upper.startsWith("VIP_") ||
    upper.startsWith("CHAT_") ||
    upper.startsWith("SUPPORT_") ||
    upper.startsWith("SYS_")
  );
}

function getAvailableSheets() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName())
    .filter(name => !isSystemSheet_(name));
}

function getSearchableSheets_() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets()
    .filter(s => !isSystemSheet_(s.getName()));
}

/****************************
 * CACHE HELPERS
 ****************************/
function buildSearchCacheKey_(keywordInput, phraseList, sheetArg, searchType) {
  // Lấy ra mốc thời gian sửa đổi gần nhất, nếu không có thì mặc định là '1'
  const dataVersion = PropertiesService.getScriptProperties().getProperty('DATA_VERSION') || '1';
  
  const raw = JSON.stringify({
    keywordInput: keywordInput || "",
    phraseList: phraseList || [],
    sheetArg: sheetArg || [],
    searchType: searchType || "",
    version: dataVersion // Đưa phiên bản vào đây để bẻ gãy Cache khi dữ liệu mới
  });
  return "SEARCH_" + Utilities.base64EncodeWebSafe(raw).slice(0, 220);
}

/****************************
 * SEARCH
 ****************************/
function searchItems(keyword, phrase, sheets, type, isSearchAll) {
  const cache = CacheService.getScriptCache();
  const cacheKey = buildSearchCacheKey_(keyword, phrase, sheets, type);

  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let targetSheets = [];
  if (Array.isArray(sheets)) {
    targetSheets = sheets.map(n => ss.getSheetByName(n)).filter(Boolean);
  } else if (sheets === "Tất cả") {
    targetSheets = getSearchableSheets_();
  } else if (typeof sheets === "string") {
    const sh = ss.getSheetByName(sheets);
    if (sh) targetSheets = [sh];
  }

  if (!targetSheets.length) return [];

  const isStarSearch = !!isSearchAll || String(keyword || "").trim() === "*";

  const typedKeywords = isStarSearch
    ? []
    : String(keyword || "")
        .split(",")
        .map(k => k.trim())
        .filter(Boolean);

  const wildcardPatterns = [];
  const fuzzyTokens = [];

  for (let i = 0; i < typedKeywords.length; i++) {
    const k = typedKeywords[i];

    if (k.includes("***")) {
      const pattern = removeVietnameseTones(k)
        .replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
        .replace(/\\\*\\\*\\\*/g, ".*")
        .trim();

      wildcardPatterns.push(new RegExp(pattern, "i"));
    }

    const tokens = tokenizeKeyword_(typedKeywords[i]);
    for (let j = 0; j < tokens.length; j++) {
      fuzzyTokens.push(tokens[j]);
    }
  }

  const commonKeywords = typedKeywords
    .filter(k => k && !k.includes("★") && !k.includes("***"))
    .map(k => removeVietnameseTones(k));

  const phraseRegexes = [];
  (phrase || []).forEach(key => {
    if (SEMANTIC_MAP[key]) {
      SEMANTIC_MAP[key].forEach(p => {
        phraseRegexes.push(new RegExp(p, "i"));
      });
    } else {
      phraseRegexes.push(new RegExp(key, "i"));
    }
  });

  if (
    !isStarSearch &&
    !wildcardPatterns.length &&
    !fuzzyTokens.length &&
    !commonKeywords.length &&
    !phraseRegexes.length
  ) {
    return [];
  }

  const START_ROW = 3;
  const START_COL = 3;
  const NUM_COLS = 7;

  let searchCol;
  if (type === "Trang Bị") searchCol = 2;
  else if (type === "Chỉ số") searchCol = 3;
  else searchCol = 4;

  const RETURN_COLS = [1, 2, 3, 4];
  const output = [];

  for (let s = 0; s < targetSheets.length; s++) {
    const sheet = targetSheets[s];
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW) continue;

    const numRows = lastRow - START_ROW + 1;
    const data = sheet.getRange(START_ROW, START_COL, numRows, NUM_COLS).getValues();

    for (let idx = 0; idx < data.length; idx++) {
      const row = data[idx];
      const cellValue = row[searchCol];
      if (!cellValue && !isStarSearch) continue;

      const normalized = removeVietnameseTones(cellValue);
      let matched = isStarSearch;

      if (!matched) {
        if (wildcardPatterns.length) {
          for (let i = 0; i < wildcardPatterns.length; i++) {
            if (wildcardPatterns[i].test(normalized)) {
              matched = true;
              break;
            }
          }
        } else if (phraseRegexes.length) {
          for (let i = 0; i < phraseRegexes.length; i++) {
            if (phraseRegexes[i].test(normalized)) {
              matched = true;
              break;
            }
          }
        } else if (commonKeywords.length) {
          for (let i = 0; i < commonKeywords.length; i++) {
            if (normalized.includes(commonKeywords[i])) {
              matched = true;
              break;
            }
          }
        } else if (fuzzyTokens.length) {
          let matchedCount = 0;
          const need = Math.ceil(fuzzyTokens.length * 0.7);

          for (let i = 0; i < fuzzyTokens.length; i++) {
            if (normalized.includes(fuzzyTokens[i])) {
              matchedCount++;
              if (matchedCount >= need) {
                matched = true;
                break;
              }
            }
          }
        }
      }

      if (!matched) continue;

      const values = [];
      for (let i = 0; i < RETURN_COLS.length; i++) {
        values.push(highlightSentencePro(String(row[RETURN_COLS[i]] || "")));
      }

      output.push({
        sheet: sheet.getName(),
        row: START_ROW + idx,
        values: values
      });
    }
  }

  try {
    const json = JSON.stringify(output);
    if (json.length < 90000) {
      cache.put(cacheKey, json, 120);
    }
  } catch (e) {}

  return output;
}

/****************************
 * JUMP TO CELL
 ****************************/
function jumpTo(sheetName, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  ss.setActiveSheet(sh);
  sh.setActiveRange(sh.getRange(row, 1));
}

/****************************
 * HIGHLIGHT PRO – PHRASE ONLY
 ****************************/
function highlightSentencePro(text) {
  if (!text) return text;

  let result = text;

  // RULE 1: "không thể ... bất lợi"
  result = result.replace(
    /(không thể[^.。,\n]*?bất lợi)/gi,
    '<span class="hl-batloi">$1</span>'
  );

  // RULE 2a: "bỏ qua ... vô địch"
  result = result.replace(
    /(bỏ qua[^.。,\n]*?vô địch)/gi,
    '<span class="hl-vodich">$1</span>'
  );

  // RULE 2b: "miễn dịch 100% sát thương"
  result = result.replace(
    /(miễn dịch\s*100%\s*sát thương)/gi,
    '<span class="hl-vodich">$1</span>'
  );

  // RULE 3: đánh bại / hạ gục
  result = result.replace(
    /(đánh bại trong 1 đòn đánh|đánh bại ngay lập tức|hạ gục|lập tức bị đánh bại)/gi,
    '<span class="hl-danhbai">$1</span>'
  );

  // RULE 4: "kháng 100%"
  result = result.replace(
    /(kháng\s*100%)/gi,
    '<span class="hl-khang">$1</span>'
  );

  // RULE 5: tốc đánh
  result = result.replace(
    /(tốc\s*độ\s*đánh|tốc\s*đánh)/gi,
    '<span class="hl-toc-danh">$1</span>'
  );

  // RULE 6: tốc chạy
  result = result.replace(
    /(tốc\s*độ\s*chạy|tốc\s*chạy)/gi,
    '<span class="hl-toc-chay">$1</span>'
  );

  return result;
}

// TỰ ĐỘNG XÓA CACHE KHI CÓ NGƯỜI SỬA SHEET
function onEdit(e) {
  // Tạo ra một mốc thời gian mới mỗi khi dữ liệu Sheet thay đổi
  PropertiesService.getScriptProperties().setProperty('DATA_VERSION', Date.now().toString());
}




/****************************
 * ONLINE COUNTER – mỗi TAB = 1 user
 ****************************/
function pingOnline(sessionId) {
  if (!sessionId) return { online: 0, total: getTotalVisits() };

  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();

  const raw = cache.get("ONLINE_SESSIONS");
  let sessions = raw ? JSON.parse(raw) : {};
  sessions[sessionId] = now;

  const TTL = 90 * 1000;
  Object.keys(sessions).forEach(id => {
    if (now - sessions[id] > TTL) delete sessions[id];
  });

  cache.put("ONLINE_SESSIONS", JSON.stringify(sessions), 120);

  const visitKey = "VISITED_" + sessionId;
  if (!cache.get(visitKey)) {
    const total = Number(props.getProperty("TOTAL_VISITS") || 0) + 1;
    props.setProperty("TOTAL_VISITS", String(total));
    cache.put(visitKey, "1", 6 * 60 * 60);
  }

  return {
    online: Object.keys(sessions).length,
    total: getTotalVisits()
  };
}

function getTotalVisits() {
  const props = PropertiesService.getScriptProperties();
  return Number(props.getProperty("TOTAL_VISITS") || 0);
}

/****************************
 * CHAT SUPPORT
 ****************************/
function saveChatMessage(sessionId, message) {
  if (!message) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("CHAT_SUPPORT");

  if (!sheet) {
    sheet = ss.insertSheet("CHAT_SUPPORT");
    sheet.appendRow(["Thời gian", "Session", "Nội dung"]);
  }

  sheet.appendRow([
    new Date(),
    sessionId || "unknown",
    message
  ]);
}

// Kiểm tra tài khoản truy cập

function verifyVipLogin(data) {
  data = data || {};

  const email = String(data.email || "").trim().toLowerCase();
  const password = String(data.password || "").trim();

  if (!email || !email.includes("@")) {
    return {
      ok: false,
      email: "",
      message: "Vui lòng nhập địa chỉ email hợp lệ."
    };
  }

  if (!password) {
    return {
      ok: false,
      email: email,
      message: "Vui lòng nhập mật khẩu VIP."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = ss.getSheetByName(VIP_REQUEST_SHEET_NAME);

  let logSheet = ss.getSheetByName("VIP_Access");
  if (!logSheet) {
    logSheet = ss.insertSheet("VIP_Access");
    logSheet.appendRow(["Thời gian", "Email đăng nhập", "Trạng thái", "Ghi chú chi tiết"]);
  }

  if (!requestSheet) {
    return {
      ok: false,
      email: email,
      message: "Lỗi hệ thống: Không tìm thấy sheet VIP_REQUESTS."
    };
  }

  const lastRow = requestSheet.getLastRow();

  if (lastRow < 2) {
    logSheet.appendRow([new Date(), email, "TỪ CHỐI", "VIP_REQUESTS chưa có tài khoản nào"]);
    return {
      ok: false,
      email: email,
      message: "Chưa có tài khoản VIP nào được duyệt."
    };
  }

  const dataRows = requestSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  let foundEmail = false;
  let foundApprovedAccount = false;

  for (let i = dataRows.length - 1; i >= 0; i--) {
    const row = dataRows[i];

    const contactEmail = String(row[2] || "").trim().toLowerCase();
    const status = String(row[4] || "").trim();
    const vipEmail = String(row[7] || "").trim().toLowerCase();
    const vipPassword = String(row[8] || "").trim();

    const loginEmail = vipEmail || contactEmail;

    if (loginEmail !== email) {
      continue;
    }

    foundEmail = true;

    if (status !== "Đã duyệt VIP - Đã gửi tài khoản") {
      continue;
    }

    foundApprovedAccount = true;

    if (vipPassword === password) {
      logSheet.appendRow([new Date(), email, "THÀNH CÔNG", "Đăng nhập hợp lệ!"]);

      return {
        ok: true,
        email: email,
        message: "Đăng nhập VIP thành công."
      };
    }
  }

  if (!foundEmail) {
    logSheet.appendRow([new Date(), email, "TỪ CHỐI", "Email không tồn tại!"]);
    return {
      ok: false,
      email: email,
      message: "Tài khoản không tồn tại trong hệ thống."
    };
  }

  if (!foundApprovedAccount) {
    logSheet.appendRow([new Date(), email, "TỪ CHỐI", "Tài khoản chưa được admin duyệt thanh toán"]);
    return {
      ok: false,
      email: email,
      message: "Tài khoản này chưa được admin duyệt thanh toán."
    };
  }

  logSheet.appendRow([new Date(), email, "TỪ CHỐI", "Sai mật khẩu"]);

  return {
    ok: false,
    email: email,
    message: "Mật khẩu không chính xác."
  };
}


// Kiểm tra tài khoản FREE đăng nhập
function logFreeUserSearch(deviceId, userAgent, searchKeyword, currentUsage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VIP_Free_Users");
  
  // Tự động tạo sheet nếu chưa có
  if (!sheet) {
    sheet = ss.insertSheet("VIP_Free_Users");
    sheet.appendRow(["Thời gian", "Mã thiết bị (Device ID)", "Lượt dùng", "Từ khóa tìm kiếm", "Thông tin thiết bị (User Agent)"]);
  }
  
  // Ghi nhận lịch sử tìm kiếm của người dùng ẩn danh
  sheet.appendRow([new Date(), deviceId, currentUsage + "/3", searchKeyword, userAgent]);
}

//Kiểm tra hành vi tài khoản VIP
// Hàm ghi log cho tài khoản VIP
function logVipUserSearch(email, deviceId, userAgent, searchKeyword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("VIP_Search");
  
  // Tự động tạo sheet nếu chưa có
  if (!sheet) {
    sheet = ss.insertSheet("VIP_Search");
    sheet.appendRow(["Thời gian", "Email VIP", "Mã thiết bị (Device ID)", "Từ khóa tìm kiếm", "Thông tin thiết bị"]);
  }
  
  sheet.appendRow([new Date(), email, deviceId, searchKeyword, userAgent]);
}
