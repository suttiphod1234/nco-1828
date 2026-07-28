/**
 * Google Apps Script Backend for NCO.1828 Reunion Booking System
 * Spreadsheet ID: 1N4E2U1sXdU7HMCDASg5s8hxceXLJ_aRigan-pYsh-p0
 */

const SPREADSHEET_ID = "1N4E2U1sXdU7HMCDASg5s8hxceXLJ_aRigan-pYsh-p0";
const SHEET_NAME = "รายการจอง";

function doGet(e) {
  const action = e.parameter ? e.parameter.action : null;
  const keyword = e.parameter ? (e.parameter.keyword || "").trim().toLowerCase() : "";

  // Handle Search & List API for Checking Registration Status
  if (action === "search" || action === "list") {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", data: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) {
        return ContentService.createTextOutput(JSON.stringify({ status: "success", data: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const results = [];

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const timestamp = row[0] || "";
        const name = String(row[1] || "");
        const email = String(row[2] || "");
        const phone = String(row[3] || "");
        const memberStatus = row[4] || "";
        const packageSelected = row[5] || "";
        const mainShirtSize = row[6] || "";
        const followerCount = row[7] || 0;
        const followerRoom = row[8] || "";
        const extraShirts = row[9] || "ไม่มี";
        const totalAmount = row[10] || 0;
        const paymentStatus = row[11] || "ลงทะเบียนแล้ว (รอตรวจสอบสลิป)";

        // Filter if keyword provided
        if (keyword) {
          const matchName = name.toLowerCase().includes(keyword);
          const matchPhone = phone.toLowerCase().includes(keyword);
          const matchEmail = email.toLowerCase().includes(keyword);
          if (!matchName && !matchPhone && !matchEmail) {
            continue;
          }
        }

        // Mask phone for privacy in public search (e.g., 089-***-1550)
        let maskedPhone = phone;
        if (phone.length >= 9) {
          maskedPhone = phone.substring(0, 3) + "-***-" + phone.substring(phone.length - 4);
        }

        results.push({
          rowNum: i + 1,
          timestamp: timestamp,
          fullName: name,
          email: email,
          phoneMasked: maskedPhone,
          memberStatus: memberStatus,
          packageSelected: packageSelected,
          mainShirtSize: mainShirtSize,
          followerCount: followerCount,
          followerRoom: followerRoom,
          extraShirts: extraShirts,
          totalAmount: totalAmount,
          paymentStatus: paymentStatus
        });
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "success", data: results }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Default Web App HTML output
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ลงทะเบียนงานเลี้ยงรุ่น NCO.1828 เหล่าทหารปืนใหญ่')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "NCO.1828 Booking Web App API is running" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    // Create sheet if not exists
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "วัน-เวลา ลงทะเบียน",
        "ยศ - ชื่อ-นามสกุล",
        "อีเมล",
        "เบอร์โทรศัพท์",
        "สถานะสมาชิก",
        "แพ็กเกจที่เลือก",
        "ไซส์เสื้อหลัก",
        "จำนวนผู้ติดตาม",
        "ผู้ติดตามต้องการห้องพัก",
        "เสื้อที่สั่งเพิ่ม",
        "ยอดรวมทั้งสิ้น (บาท)",
        "สถานะชำระเงิน",
        "หมายเหตุ",
        "สลิปการโอนเงิน (URL/Base64)"
      ]);
      sheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#d4af37").setFontColor("#1a202c");
    }
    
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = e.parameter;
    }
    
    const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    const name = data.fullName || "";
    const email = data.email || "";
    const phone = data.phone || "";
    const memberStatus = data.memberStatus || "";
    const packageSelected = data.packageSelected || "";
    const mainShirtSize = data.mainShirtSize || "";
    const followerCount = data.followerCount || 0;
    const followerRoom = data.followerRoom ? "รับห้องพัก (+200B)" : "ไม่รับห้องพัก";
    const extraShirts = data.extraShirtsDetail || "ไม่มี";
    const totalAmount = data.totalAmount || 0;
    const note = data.note || "";
    const slipData = data.slipImage || "";
    const initialStatus = "รับข้อมูลแล้ว (รอตรวจสอบสลิป)";

    // Save image to Google Drive if folder exists or store link
    let slipUrl = slipData;
    if (slipData && slipData.startsWith("data:image")) {
      try {
        const folderName = "NCO1828_Slips";
        let folder;
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.createFolder(folderName);
        }
        
        const base64Data = slipData.split(",")[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", `Slip_${name}_${Date.now()}.jpg`);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        slipUrl = file.getUrl();
      } catch (err) {
        slipUrl = "แนบสลิปเรียบร้อยแล้ว (อัปโหลดเข้า Drive ติดขัดสิทธิ์)";
      }
    }
    
    sheet.appendRow([
      timestamp,
      name,
      email,
      phone,
      memberStatus,
      packageSelected,
      mainShirtSize,
      followerCount,
      followerRoom,
      extraShirts,
      totalAmount,
      initialStatus,
      note,
      slipUrl
    ]);

    // Send Confirmation Email if Email provided
    if (email && email.includes("@")) {
      try {
        const emailSubject = `[ยืนยันการลงทะเบียน] งานเลี้ยงรุ่น NCO.1828 เหล่าทหารปืนใหญ่ - ${name}`;
        const emailBody = `
          <div style="font-family: 'Sarabun', Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #d4af37; border-radius: 16px; padding: 25px; background-color: #0b1326; color: #f8fafc;">
            <h2 style="color: #d4af37; text-align: center; margin-bottom: 5px;">ยืนยันการลงทะเบียนงานเลี้ยงรุ่น NCO.1828</h2>
            <p style="text-align: center; color: #94a3b8; font-size: 0.95rem; margin-top: 0;">"กลับบ้านที่ลพบุรี แหล่งกำเนิดของพวกเราทุกคน"</p>
            <hr style="border-color: rgba(212,175,55,0.3); margin: 20px 0;">
            
            <p>เรียน <strong>${name}</strong>,</p>
            <p>ระบบได้รับการลงทะเบียนและหลักฐานการโอนเงินเรียบร้อยแล้ว โดยมีรายละเอียดดังนี้:</p>
            
            <table style="width: 100%; border-collapse: collapse; margin: 20px 0; color: #fff; font-size: 0.95rem; background: rgba(255,255,255,0.03); border-radius: 8px; overflow: hidden;">
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);"><td style="padding: 10px; color: #94a3b8;">ยศ - ชื่อ-นามสกุล:</td><td style="padding: 10px; font-weight: bold;">${name}</td></tr>
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);"><td style="padding: 10px; color: #94a3b8;">เบอร์โทรศัพท์:</td><td style="padding: 10px;">${phone}</td></tr>
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);"><td style="padding: 10px; color: #94a3b8;">สถานะสมาชิก:</td><td style="padding: 10px;">${memberStatus}</td></tr>
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);"><td style="padding: 10px; color: #94a3b8;">แพ็กเกจที่เลือก:</td><td style="padding: 10px;">${packageSelected}</td></tr>
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);"><td style="padding: 10px; color: #94a3b8;">ไซส์เสื้อหลัก:</td><td style="padding: 10px; font-weight: bold; color: #d4af37;">${mainShirtSize}</td></tr>
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);"><td style="padding: 10px; color: #94a3b8;">จำนวนผู้ติดตาม:</td><td style="padding: 10px;">${followerCount} คน (${followerRoom})</td></tr>
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);"><td style="padding: 10px; color: #94a3b8;">เสื้อสั่งซื้อเพิ่ม:</td><td style="padding: 10px;">${extraShirts}</td></tr>
              <tr style="border-bottom: 1px solid rgba(255,255,255,0.1);"><td style="padding: 10px; color: #94a3b8;">ยอดชำระเงินรวม:</td><td style="padding: 10px; font-weight: bold; color: #22c55e; font-size: 1.15rem;">${totalAmount} บาท</td></tr>
            </table>

            <div style="background: rgba(212, 175, 55, 0.15); border: 1px solid #d4af37; border-radius: 12px; padding: 15px; text-align: center; margin: 20px 0;">
              <p style="margin: 0; color: #d4af37; font-weight: bold; font-size: 1.05rem;">📅 กำหนดการวันจัดงาน:</p>
              <p style="margin: 6px 0 0 0; color: #fff;">วันศุกร์ที่ 19 กุมภาพันธ์ 2570 ณ โรงแรมโตเกียว จ.ลพบุรี</p>
            </div>

            <p style="font-size: 0.85rem; color: #94a3b8; margin-top: 25px; text-align: center;">สอบถามข้อมูลเพิ่มเติม โทร: 089-208-1550 (ร.ต.วันชัย มีชาญ)</p>
          </div>
        `;

        MailApp.sendEmail({
          to: email,
          subject: emailSubject,
          htmlBody: emailBody
        });
      } catch (mailErr) {
        Logger.log("Email notification error: " + mailErr.toString());
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "ลงทะเบียนเรียบร้อยแล้ว"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
