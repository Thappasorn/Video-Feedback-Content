/**
 * Video Review → Google Sheet — Apps Script Web App
 *
 * วิธีติดตั้ง (ทำครั้งเดียว)
 *  1. เปิด https://script.google.com/home/projects/create → New project
 *  2. ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้ลงไป → บันทึก (Ctrl/Cmd+S)
 *  3. Deploy → New deployment → Select type: Web app
 *  4. Execute as: Me / Who has access: Anyone → Deploy
 *  5. อนุญาตสิทธิ์ (Advanced → Go to project (unsafe) → Allow)
 *     แล้วคัดลอก Web app URL (ลงท้าย /exec) ไปใส่ในหน้า ตั้งค่า ⚙ ของเว็บ Video Review
 *
 * เว้น Spreadsheet ID ว่าง = สร้างไฟล์ใหม่ทุกครั้ง / ใส่ ID = เพิ่มชีตในไฟล์เดิม
 * ภาพเฟรมอัปโหลดขึ้น Drive โฟลเดอร์ "Video Review Frames" และตั้งแชร์ "ทุกคนที่มีลิงก์"
 * เพื่อให้สูตร =IMAGE() แสดงผลในชีตได้
 */

function doGet(){ return ContentService.createTextOutput('Video Review endpoint OK'); }

function doPost(e){
  try{
    var p = JSON.parse(e.postData.contents);
    var ss = p.sheetId ? SpreadsheetApp.openById(p.sheetId)
                       : SpreadsheetApp.create('Video Review - ' + (p.project||'untitled'));
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd HH:mm');
    var sh = ss.insertSheet((p.project||'log').substring(0,20) + ' ' + stamp);

    sh.getRange(1,1,4,2).setValues([
      ['โปรเจกต์', p.project||''], ['วิดีโอ', p.source||''],
      ['ความยาว', p.duration||''], ['ส่งออกเมื่อ', new Date()]
    ]);
    sh.getRange(1,1,4,1).setFontWeight('bold');

    var head = ['#','Timecode','วินาที','หมวด','คอมเมนต์','ผู้คอมเมนต์','สถานะ','ลิงก์ไปเวลานั้น','ภาพเฟรม'];
    sh.getRange(6,1,1,head.length).setValues([head])
      .setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');

    var folder = null;
    function putImage(sheet,row,col,b64,name){
      if(!folder){
        var it = DriveApp.getFoldersByName('Video Review Frames');
        folder = it.hasNext() ? it.next() : DriveApp.createFolder('Video Review Frames');
      }
      var f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64),'image/jpeg',name));
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheet.getRange(row,col).setFormula('=IMAGE("https://drive.google.com/uc?export=view&id=' + f.getId() + '")');
      sheet.setRowHeight(row,90);
    }

    (p.rows||[]).forEach(function(r,i){
      var row = 7 + i;
      sh.getRange(row,1,1,8).setValues([[r.no,r.tc,r.sec,r.tag,r.text,r.author,r.status,r.link]]);
      if (r.image) putImage(sh,row,9,r.image, r.no + '_' + String(r.tc).replace(/:/g,'-') + '.jpg');
    });
    sh.setColumnWidth(5,420); sh.setColumnWidth(9,190);
    if (sh.getLastRow() > 5) sh.getRange(6,1,sh.getLastRow()-5,9).setVerticalAlignment('top').setWrap(true);
    sh.setFrozenRows(6);

    if (p.stocks && p.stocks.length) {
      var sh2 = ss.insertSheet('stock ' + stamp);
      var h2 = ['#','IN','OUT','ระยะ(วินาที)','ประโยค / คำพูด (VO)','สต็อกที่จะวาง','หมายเหตุ','ลิงก์ไปเวลานั้น','ภาพเฟรม'];
      sh2.getRange(1,1,1,h2.length).setValues([h2])
        .setFontWeight('bold').setBackground('#4c1d95').setFontColor('#ffffff');
      p.stocks.forEach(function(r,i){
        var row = 2 + i;
        sh2.getRange(row,1,1,8).setValues([[r.no,r.in,r.out,r.dur,r.line,r.stock,r.note,r.link]]);
        if (r.image) putImage(sh2,row,9,r.image,'stock_' + r.no + '.jpg');
      });
      sh2.setColumnWidth(5,420); sh2.setColumnWidth(6,220); sh2.setColumnWidth(9,190);
      sh2.getRange(1,1,sh2.getLastRow(),9).setVerticalAlignment('top').setWrap(true);
      sh2.setFrozenRows(1);
    }

    return ContentService.createTextOutput(JSON.stringify({ok:true, url:ss.getUrl()}))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
