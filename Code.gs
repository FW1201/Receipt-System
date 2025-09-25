// 電子領據系統 - 完整版
let SHEET_ID = "";
let SIGNATURE_FOLDER_ID = "";
let PDF_FOLDER_ID = "";
let TEMPLATE_ID = "";
let BANKBOOK_FOLDER_ID = "";

const SHEET_NAME = "Submissions";

function doGet(e) {
  e = e || {};
  e.parameter = e.parameter || {};
  
  if (e.parameter.page === 'admin') {
    return HtmlService.createTemplateFromFile('Interface').evaluate()
      .setTitle('電子領據管理後台').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (e.parameter.id) {
    const data = findRowByUniqueId_(e.parameter.id);
    if (!data || data.status !== "Sent") {
      return HtmlService.createHtmlOutput("<h1>連結無效或已過期</h1><p>請聯繫管理員重新發送邀請。</p>");
    }
    const template = HtmlService.createTemplateFromFile('Interface');
    template.mode = 'form';
    template.data = data;
    return template.evaluate().setTitle('電子領據填報').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createTemplateFromFile('Interface').evaluate().setTitle('電子領據管理後台');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('電子領據系統')
    .addItem('🚀 首次設定', 'setupSystem')
    .addItem('📊 系統狀態', 'showSystemStatus')
    .addItem('🔄 重置系統', 'resetSystem')
    .addToUi();
}

function setupSystem() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    SHEET_ID = ss.getId();
    const adminEmail = Session.getActiveUser().getEmail();
    
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = ["唯一ID", "狀態", "領據抬頭", "事件名稱", "事由", "領款人姓名", "Email", "身分證字號", "金額", "時數", "撥款單位", "連結發送時間", "提交時間", "填寫身分證", "戶籍地址", "聯絡電話", "服務單位", "銀行代號", "分行代號", "帳號", "存摺影本", "簽名檔案", "PDF領據"];
      sheet.getRange(1, 1, 1, 23).setValues([headers]).setFontWeight("bold").setBackground("#e3f2fd");
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, 23);
    }

    const rootFolder = DriveApp.getRootFolder();
    const systemFolder = rootFolder.createFolder(`電子領據系統_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`);
    const signatureFolder = systemFolder.createFolder("簽名檔案");
    const pdfFolder = systemFolder.createFolder("PDF領據");
    const bankbookFolder = systemFolder.createFolder("存摺影本");
    
    SIGNATURE_FOLDER_ID = signatureFolder.getId();
    PDF_FOLDER_ID = pdfFolder.getId();
    BANKBOOK_FOLDER_ID = bankbookFolder.getId();

    const templateDoc = DocumentApp.create("領據範本");
    const body = templateDoc.getBody();
    body.clear();
    body.appendParagraph("{{receipt_header}}").setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("領據").setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("");
    body.appendParagraph("茲收到 {{recipient_name}} 先生/女士");
    body.appendParagraph("身分證字號：{{submitted_id_number}}");
    body.appendParagraph("戶籍地址：{{submitted_address}}");
    body.appendParagraph("聯絡電話：{{submitted_phone}}");
    body.appendParagraph("服務單位：{{service_unit}}");
    body.appendParagraph("撥款單位：{{payer_unit}}");
    body.appendParagraph("事由：{{event_reason}}");
    body.appendParagraph("金額：新台幣 {{amount}} 元整 ({{amount_text}})");
    body.appendParagraph("時數：{{hours}} 小時");
    body.appendParagraph("銀行資訊：{{bank_code}}-{{branch_code}} 帳號：{{account_number}}");
    body.appendParagraph("");
    body.appendParagraph("日期：{{submission_date}}");
    body.appendParagraph("簽名：{{signature}}");
    templateDoc.saveAndClose();
    systemFolder.addFile(DriveApp.getFileById(templateDoc.getId()));
    TEMPLATE_ID = templateDoc.getId();

    PropertiesService.getScriptProperties().setProperties({
      'SHEET_ID': SHEET_ID,
      'SIGNATURE_FOLDER_ID': SIGNATURE_FOLDER_ID,
      'PDF_FOLDER_ID': PDF_FOLDER_ID,
      'BANKBOOK_FOLDER_ID': BANKBOOK_FOLDER_ID,
      'TEMPLATE_ID': TEMPLATE_ID,
      'ADMIN_EMAIL': adminEmail
    });

    ui.alert('系統設定完成！', `管理員：${adminEmail}\nWeb App：${ScriptApp.getService().getUrl()}?page=admin\n\n請將 Web App URL 加入書籤以便日後使用。`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('設定失敗', `錯誤訊息：${error.message}\n\n請檢查權限設定並重試。`, ui.ButtonSet.OK);
  }
}

function resetSystem() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('確認重置', '此操作將清除所有系統設定，是否繼續？', ui.ButtonSet.YES_NO);
  if (result === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteAll();
    ui.alert('系統已重置', '請重新執行「首次設定」。', ui.ButtonSet.OK);
  }
}

function verifyAdminLogin() {
  try {
    const currentUser = Session.getEffectiveUser().getEmail();
    if (!currentUser) {
      return { success: false, message: "無法取得用戶資訊，請確認 Web App 部署設定" };
    }
    
    loadConfig_();
    let adminEmail = getAdminEmail_();
    
    if (!adminEmail) {
      adminEmail = currentUser;
      PropertiesService.getScriptProperties().setProperty('ADMIN_EMAIL', adminEmail);
    }
    
    if (currentUser !== adminEmail) {
      return { success: false, message: `權限不足。僅限管理員：${adminEmail}` };
    }
    
    return { success: true, user: currentUser };
  } catch (error) {
    return { success: false, message: `驗證失敗：${error.message}` };
  }
}

function sendInvitation(data) {
  const loginCheck = verifyAdminLogin();
  if (!loginCheck.success) return loginCheck.message;
  
  loadConfig_();
  if (!SHEET_ID) return "系統未初始化，請先執行首次設定";
  
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) return "找不到資料表，請重新執行首次設定";
    
    const uniqueId = Utilities.getUuid();
    const link = `${ScriptApp.getService().getUrl()}?id=${uniqueId}`;
    const timestamp = new Date();

    sheet.appendRow([uniqueId, "Pending", data.receiptHeader, data.eventName, data.eventReason, data.recipientName, data.recipientEmail, "", data.amount, data.hours, data.payerUnit, timestamp]);

    MailApp.sendEmail({
      to: data.recipientEmail,
      subject: `【領據簽署】${data.eventName} - ${data.receiptHeader}`,
      htmlBody: createEmailTemplate(data, link),
      name: '電子領據系統'
    });
    
    sheet.getRange(sheet.getLastRow(), 2).setValue('Sent');
    sheet.getRange(sheet.getLastRow(), 12).setValue(timestamp);
    
    return "邀請已發送";
  } catch(e) {
    return `發送失敗：${e.message}`;
  }
}

function createEmailTemplate(data, link) {
  return `
    <div style="font-family:'Microsoft JhengHei',Arial,sans-serif;max-width:600px;margin:0 auto;padding:20px;background:#f8f9fa;">
      <div style="background:#0d6efd;color:white;padding:30px;text-align:center;border-radius:8px 8px 0 0;">
        <h2 style="margin:0;font-size:24px;">電子領據填報通知</h2>
      </div>
      <div style="background:white;padding:30px;border-radius:0 0 8px 8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);">
        <p style="font-size:16px;color:#333;">您好 <strong>${data.recipientName}</strong>，</p>
        <p style="color:#666;">關於「<strong>${data.eventName}</strong>」，請點擊下方連結填寫領據資料：</p>
        
        <div style="background:#f8f9fa;padding:20px;border-radius:6px;margin:20px 0;">
          <p style="margin:5px 0;"><strong>領據抬頭：</strong>${data.receiptHeader}</p>
          <p style="margin:5px 0;"><strong>撥款單位：</strong>${data.payerUnit}</p>
          <p style="margin:5px 0;"><strong>金額：</strong>新台幣 ${data.amount} 元整</p>
          <p style="margin:5px 0;"><strong>時數：</strong>${data.hours} 小時</p>
        </div>
        
        <div style="text-align:center;margin:30px 0;">
          <a href="${link}" style="display:inline-block;padding:15px 30px;background:#0d6efd;color:white;text-decoration:none;border-radius:6px;font-size:16px;font-weight:bold;">點此填寫領據</a>
        </div>
        
        <div style="border-top:1px solid #eee;padding-top:20px;margin-top:30px;">
          <p style="font-size:12px;color:#999;margin:0;">如無法點擊按鈕，請複製以下連結至瀏覽器：</p>
          <p style="font-size:12px;color:#666;word-break:break-all;margin:5px 0 0 0;">${link}</p>
        </div>
      </div>
    </div>
  `;
}

function getRecords() {
  const loginCheck = verifyAdminLogin();
  if (!loginCheck.success) return [];
  
  loadConfig_();
  if (!SHEET_ID) return [];
  
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 23).getDisplayValues();
    const headers = ["uniqueId", "status", "receipt_header", "event_name", "event_reason", "recipient_name", "recipient_email", "recipient_id_number", "amount", "hours", "payer_unit", "link_sent_timestamp", "submission_timestamp", "submitted_id_number", "submitted_address", "submitted_phone", "service_unit", "bank_code", "branch_code", "account_number", "bankbook_image_link", "signature_image_link", "pdf_receipt_link"];
    
    return values.map(row => {
      let obj = {};
      headers.forEach((header, i) => obj[header] = row[i]);
      return obj;
    }).reverse();
  } catch (error) {
    return [];
  }
}

function processFormSubmission(formData) {
  loadConfig_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const data = findRowByUniqueId_(formData.uniqueId);
    if (!data || data.status !== "Sent") return "連結無效或已使用";

    const signatureFolder = DriveApp.getFolderById(SIGNATURE_FOLDER_ID);
    const bankbookFolder = DriveApp.getFolderById(BANKBOOK_FOLDER_ID);
    
    const decodedSignature = Utilities.base64Decode(formData.signature.split(',')[1]);
    const signatureFile = signatureFolder.createFile(Utilities.newBlob(decodedSignature, 'image/png', `signature_${formData.uniqueId}.png`));
    
    let bankbookFileUrl = '';
    if (formData.bankbookImage) {
      const decodedBankbook = Utilities.base64Decode(formData.bankbookImage.split(',')[1]);
      const bankbookFile = bankbookFolder.createFile(Utilities.newBlob(decodedBankbook, 'image/jpeg', `bankbook_${formData.uniqueId}.jpg`));
      bankbookFileUrl = bankbookFile.getUrl();
    }
    
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const targetRow = sheet.getRange(data.row, 1, 1, 23);
    
    targetRow.getCell(1, 13).setValue(new Date());
    targetRow.getCell(1, 14).setValue(formData.idNumber);
    targetRow.getCell(1, 15).setValue(formData.address);
    targetRow.getCell(1, 16).setValue(formData.phone);
    targetRow.getCell(1, 17).setValue(formData.serviceUnit);
    targetRow.getCell(1, 18).setValue(formData.bankCode);
    targetRow.getCell(1, 19).setValue(formData.branchCode);
    targetRow.getCell(1, 20).setValue(formData.accountNumber);
    targetRow.getCell(1, 21).setValue(bankbookFileUrl);
    targetRow.getCell(1, 22).setValue(signatureFile.getUrl());
    targetRow.getCell(1, 2).setValue("Submitted");
    
    return "提交成功！資料已儲存，管理員將為您產生正式領據。";
  } catch (error) {
    return `系統錯誤：${error.message}`;
  } finally {
    lock.releaseLock();
  }
}

function generatePdf(uniqueId) {
  const loginCheck = verifyAdminLogin();
  if (!loginCheck.success) return { success: false, message: loginCheck.message };
  
  loadConfig_();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const data = findRowByUniqueId_(uniqueId, true);
    if (!data || data.status !== "Submitted") return { success: false, message: '資料狀態錯誤，無法產生PDF' };
    if (data.pdf_receipt_link) return { success: true, url: data.pdf_receipt_link, message: 'PDF已存在' };

    const templateFile = DriveApp.getFileById(TEMPLATE_ID);
    const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
    const fileName = `領據_${data.event_name}_${data.recipient_name}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}`;
    const newFile = templateFile.makeCopy(fileName, pdfFolder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    // 替換所有變數
    const replacements = {
      '{{receipt_header}}': data.receipt_header || '',
      '{{recipient_name}}': data.recipient_name || '',
      '{{submitted_id_number}}': data.submitted_id_number || '',
      '{{submitted_address}}': data.submitted_address || '',
      '{{submitted_phone}}': data.submitted_phone || '',
      '{{service_unit}}': data.service_unit || '',
      '{{payer_unit}}': data.payer_unit || '',
      '{{event_reason}}': data.event_reason || '',
      '{{amount}}': data.amount || '',
      '{{amount_text}}': convertToChineseNumber(data.amount),
      '{{hours}}': data.hours || '',
      '{{bank_code}}': data.bank_code || '',
      '{{branch_code}}': data.branch_code || '',
      '{{account_number}}': data.account_number || '',
      '{{submission_date}}': Utilities.formatDate(new Date(data.submission_timestamp), Session.getScriptTimeZone(), 'yyyy年MM月dd日')
    };

    Object.entries(replacements).forEach(([key, value]) => {
      body.replaceText(key, value);
    });

    // 處理簽名
    if (data.signature_image_link) {
      try {
        const signatureId = data.signature_image_link.match(/id=([^&]+)/);
        if (signatureId && signatureId[1]) {
          const signatureImage = DriveApp.getFileById(signatureId[1]).getBlob();
          const placeholder = body.findText('{{signature}}');
          if (placeholder) {
            placeholder.getElement().getParent().asParagraph().clear().insertInlineImage(0, signatureImage).setWidth(120).setHeight(60);
          }
        } else {
          body.replaceText('{{signature}}', '[簽名檔案]');
        }
      } catch (e) {
        body.replaceText('{{signature}}', '[簽名載入失敗]');
      }
    } else {
      body.replaceText('{{signature}}', '[無簽名]');
    }
    
    doc.saveAndClose();
    const pdfFile = pdfFolder.createFile(newFile.getAs('application/pdf')).setName(`${fileName}.pdf`);
    DriveApp.getFileById(newFile.getId()).setTrashed(true);
    
    SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME).getRange(data.row, 23).setValue(pdfFile.getUrl());
    return { success: true, url: pdfFile.getUrl() };
  } catch (e) {
    return { success: false, message: `PDF產生失敗：${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function getAdminEmail_() {
  const storedEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
  if (storedEmail) return storedEmail;
  
  const currentEmail = Session.getEffectiveUser().getEmail();
  if (currentEmail) {
    PropertiesService.getScriptProperties().setProperty('ADMIN_EMAIL', currentEmail);
    return currentEmail;
  }
  
  return null;
}

function getAdminEmail() {
  return getAdminEmail_();
}

function loadConfig_() {
  const props = PropertiesService.getScriptProperties();
  SHEET_ID = props.getProperty('SHEET_ID') || "";
  SIGNATURE_FOLDER_ID = props.getProperty('SIGNATURE_FOLDER_ID') || "";
  PDF_FOLDER_ID = props.getProperty('PDF_FOLDER_ID') || "";
  BANKBOOK_FOLDER_ID = props.getProperty('BANKBOOK_FOLDER_ID') || "";
  TEMPLATE_ID = props.getProperty('TEMPLATE_ID') || "";
}

function findRowByUniqueId_(uniqueId, fullData = false) {
  loadConfig_();
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === uniqueId) {
      const rowNum = i + 2;
      if (fullData) {
        const rowValues = sheet.getRange(rowNum, 1, 1, 23).getValues()[0];
        const headers = ["uniqueId", "status", "receipt_header", "event_name", "event_reason", "recipient_name", "recipient_email", "recipient_id_number", "amount", "hours", "payer_unit", "link_sent_timestamp", "submission_timestamp", "submitted_id_number", "submitted_address", "submitted_phone", "service_unit", "bank_code", "branch_code", "account_number", "bankbook_image_link", "signature_image_link", "pdf_receipt_link"];
        let obj = { row: rowNum };
        headers.forEach((header, index) => obj[header] = rowValues[index]);
        return obj;
      } else {
        const rowValues = sheet.getRange(rowNum, 1, 1, 11).getValues()[0];
        return { row: rowNum, uniqueId: rowValues[0], status: rowValues[1], receiptHeader: rowValues[2], recipientName: rowValues[5], amount: rowValues[8], hours: rowValues[9] };
      }
    }
  }
  return null;
}

function convertToChineseNumber(n) {
  if (isNaN(n) || n === '') return "";
  const fraction = ['角', '分'];
  const digit = ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖'];
  const unit = [['元', '萬', '億'], ['', '拾', '佰', '仟']];
  let head = n < 0 ? '負' : '';
  n = Math.abs(n);
  let s = '';
  for (let i = 0; i < fraction.length; i++) {
    s += (digit[Math.floor(n * 10 * Math.pow(10, i)) % 10] + fraction[i]).replace(/零./, '');
  }
  s = s || '整';
  n = Math.floor(n);
  for (let i = 0; i < unit[0].length && n > 0; i++) {
    let p = '';
    for (let j = 0; j < unit[1].length && n > 0; j++) {
      p = digit[n % 10] + unit[1][j] + p;
      n = Math.floor(n / 10);
    }
    s = p.replace(/(零.)*零$/, '').replace(/^$/, '零') + unit[0][i] + s;
  }
  return head + s.replace(/(零.)*零元/, '元').replace(/(零.)+/g, '零').replace(/^整$/, '零元整');
}

function showSystemStatus() {
  loadConfig_();
  const adminEmail = getAdminEmail_();
  const webAppUrl = ScriptApp.getService().getUrl();
  
  SpreadsheetApp.getUi().alert('系統狀態', 
    `管理員：${adminEmail}\nWeb App：${webAppUrl}?page=admin\n\n配置狀態：\n✅ 試算表：${SHEET_ID ? '已設定' : '❌ 未設定'}\n✅ 簽名資料夾：${SIGNATURE_FOLDER_ID ? '已設定' : '❌ 未設定'}\n✅ PDF資料夾：${PDF_FOLDER_ID ? '已設定' : '❌ 未設定'}\n✅ 存摺資料夾：${BANKBOOK_FOLDER_ID ? '已設定' : '❌ 未設定'}\n✅ 範本文件：${TEMPLATE_ID ? '已設定' : '❌ 未設定'}`, 
    SpreadsheetApp.getUi().ButtonSet.OK);
}