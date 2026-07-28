/**
 * Google Apps Script Backend for NCO.1828 Reunion Booking System
 * Spreadsheet ID: 1N4E2U1sXdU7HMCDASg5s8hxceXLJ_aRigan-pYsh-p0
 */

const SPREADSHEET_ID = "1N4E2U1sXdU7HMCDASg5s8hxceXLJ_aRigan-pYsh-p0";
const SHEET_NAME = "รายการจอง";

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "NCO.1828 Booking Web App API is running" }))
    .setMimeType(ContentService.MimeType.JSON);
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
        "เบอร์โทรศัพท์",
        "สถานะสมาชิก",
        "แพ็กเกจที่เลือก",
        "ไซส์เสื้อหลัก",
        "จำนวนผู้ติดตาม",
        "ผู้ติดตามต้องการห้องพัก",
        "เสื้อที่สั่งเพิ่ม",
        "ยอดรวมทั้งสิ้น (บาท)",
        "หมายเหตุ",
        "สลิปการโอนเงิน (URL/Base64)"
      ]);
      sheet.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#d4af37").setFontColor("#1a202c");
    }
    
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = e.parameter;
    }
    
    const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    const name = data.fullName || "";
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
        // Fallback to storing indicator if Drive permission issue
        slipUrl = "แนบสลิปเรียบร้อยแล้ว (การอัปโหลดไฟล์ไปยัง Drive ติดขัดสิทธิ์)";
      }
    }
    
    sheet.appendRow([
      timestamp,
      name,
      phone,
      memberStatus,
      packageSelected,
      mainShirtSize,
      followerCount,
      followerRoom,
      extraShirts,
      totalAmount,
      note,
      slipUrl
    ]);
    
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
