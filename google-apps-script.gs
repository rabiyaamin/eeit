/*
 EEIT Admission Form -> Google Sheet

 SETUP (sirf ek baar):
 1. Google Drive mein nayi Google Sheet banayein, naam: "EEIT Admissions".
 2. Sheet mein menu: Extensions > Apps Script kholein.
 3. Wahan jo purana code ho wo hata kar ye poora code paste karein, phir Save karein.
 4. Deploy > New deployment > type: "Web app" chunein.
      - Execute as: Me
      - Who has access: Anyone
    Deploy dabayein aur permissions Allow karein.
 5. Jo "Web app URL" mile usay copy karein.
 6. admission-form.html kholein, SCRIPT_URL = "PASTE_GOOGLE_APPS_SCRIPT_URL_HERE"
    ki jagah wo URL paste karein.
 7. Code mein koi change karein to: Deploy > Manage deployments > Edit > New version.
*/

var SHEET_NAME   = 'Admissions';
var NOTIFY_EMAIL = 'eeinstituteoftechnology@gmail.com';   // naya form aane par email; khali chhor dein to email band
var HEADERS = ['Submitted at','Course','Class mode','Full name','Father / Guardian','Date of birth',
               'CNIC / B-Form','Phone / WhatsApp','Email','Address','Qualification','Preferred timing','Agreed'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var p = e.parameter || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0b1f3a').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }

    var when = Utilities.formatDate(new Date(), 'Asia/Karachi', 'yyyy-MM-dd HH:mm:ss');
    var row = [when, p.course, p.mode, p.name, p.guardian, p.dob, p.cnic, p.phone,
               p.email, p.address, p.qualification, p.timing, p.agree].map(function (v) {
      return (v === undefined || v === null) ? '' : String(v).substring(0, 500);
    });

    var r = sh.getLastRow() + 1;
    // Plain text format: formulas nahi chalte, aur phone/CNIC ke leading zero bache rehte hain
    sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);

    if (NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(NOTIFY_EMAIL, 'New admission: ' + p.course + ' - ' + p.name,
          HEADERS.map(function (h, i) { return h + ': ' + row[i]; }).join('\n'));
      } catch (mailErr) {}
    }
    return ContentService.createTextOutput(JSON.stringify({result: 'success'}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({result: 'error', message: String(err)}))
                         .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('EEIT admission form endpoint is running.');
}
