/**
 * DRISHTI COMPUTER EDUCATION — REGISTRATION + LOGIN BACKEND
 *
 * IMPORTANT FIRST-TIME SETUP:
 * 1) Change INITIAL_ADMIN_PASSWORD below to a strong private password.
 * 2) Save this file.
 * 3) Run setupAdmin() once from the Apps Script editor and approve permissions.
 * 4) Deploy > Manage deployments > Edit > New version > Deploy.
 *
 * Admin login ID: admin (or change INITIAL_ADMIN_USERNAME)
 * Student accounts are created by the admin from the website dashboard.
 */

const SHEET_NAME = 'Student Registrations';
const ACCOUNTS_SHEET = 'Student Accounts';
const INITIAL_ADMIN_USERNAME = 'admin';
const INITIAL_ADMIN_PASSWORD = 'ChangeThis!2026'; // CHANGE BEFORE RUNNING setupAdmin()

function setupAdmin() {
  if (INITIAL_ADMIN_PASSWORD === 'ChangeThis!2026' || INITIAL_ADMIN_PASSWORD.length < 10) {
    throw new Error('Change INITIAL_ADMIN_PASSWORD to a private password of at least 10 characters before running setupAdmin().');
  }
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ADMIN_HASH')) {
    throw new Error('Admin is already initialized. Use the existing admin credentials.');
  }
  const salt = Utilities.getUuid();
  props.setProperties({
    ADMIN_USERNAME: INITIAL_ADMIN_USERNAME,
    ADMIN_SALT: salt,
    ADMIN_HASH: hashPassword_(INITIAL_ADMIN_PASSWORD, salt)
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let accounts = ss.getSheetByName(ACCOUNTS_SHEET);
  if (!accounts) accounts = ss.insertSheet(ACCOUNTS_SHEET);
  if (accounts.getLastRow() === 0) {
    accounts.appendRow(['Login ID', 'Salt', 'Password Hash', 'Registration ID', 'Created At', 'Active']);
    accounts.setFrozenRows(1);
  }
  return 'Admin initialized. Login ID: ' + INITIAL_ADMIN_USERNAME;
}

function doPost(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = String(p.action || 'register');
  try {
    let result;
    if (action === 'register') result = registerStudent_(p);
    else if (action === 'login') result = login_(p);
    else if (action === 'listStudents') result = listStudents_(p);
    else if (action === 'createStudent') result = createStudent_(p);
    else result = {ok:false, action:action, message:'Unknown request.'};
    return response_(Object.assign({type: action === 'register' ? 'drishti-registration' : 'drishti-portal', action:action}, result));
  } catch (err) {
    return response_({
      type: action === 'register' ? 'drishti-registration' : 'drishti-portal',
      action:action, ok:false, message:String(err.message || err)
    });
  }
}

function registerStudent_(p) {
  const required = ['fullName','mobile','course','qualification','address','district','consent'];
  required.forEach(k => { if (!String(p[k] || '').trim()) throw new Error('Please complete all required fields.'); });
  const mobile = String(p.mobile).replace(/\D/g,'');
  if (!/^[6-9]\d{9}$/.test(mobile)) throw new Error('Enter a valid 10-digit mobile number.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  const headers = ['Registration ID','Submitted At','Full Name','Mobile','WhatsApp','Email','Date of Birth','Gender','Course','Preferred Batch','Learning Mode','Qualification','Learning Goal','Address','District','PIN Code','How They Heard','Consent','Admission Status'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < headers.length) {
    sheet.getRange(1, sheet.getLastColumn()+1, 1, headers.length-sheet.getLastColumn()).setValues([headers.slice(sheet.getLastColumn())]);
  }
  const id = 'DRI-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Utilities.getUuid().slice(0,6).toUpperCase();
  sheet.appendRow([
    id, new Date(), clean_(p.fullName), mobile, clean_(p.whatsapp), clean_(p.email),
    clean_(p.dob), clean_(p.gender), clean_(p.course), clean_(p.batch), clean_(p.mode),
    clean_(p.qualification), clean_(p.goal), clean_(p.address), clean_(p.district),
    clean_(p.pincode), clean_(p.source), 'Yes', 'Enquiry received'
  ]);
  return {ok:true, registrationId:id};
}

function login_(p) {
  const username = String(p.username || '').trim();
  const password = String(p.password || '');
  if (!username || !password) throw new Error('Enter your login ID and password.');
  const props = PropertiesService.getScriptProperties();
  const adminName = props.getProperty('ADMIN_USERNAME');
  if (!adminName || !props.getProperty('ADMIN_HASH')) throw new Error('Admin setup is incomplete. Run setupAdmin() in Apps Script first.');

  let role = '', student = null;
  if (String(p.role) === 'admin') {
    const valid = username === adminName &&
      hashPassword_(password, props.getProperty('ADMIN_SALT')) === props.getProperty('ADMIN_HASH');
    if (!valid) throw new Error('Invalid admin login ID or password.');
    role = 'admin';
  } else {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ACCOUNTS_SHEET);
    if (!sh || sh.getLastRow() < 2) throw new Error('Student login has not been created yet. Contact the institute.');
    const rows = sh.getDataRange().getValues();
    for (let i=1; i<rows.length; i++) {
      const r=rows[i];
      if (String(r[0]) === username && String(r[5]).toLowerCase() !== 'false') {
        if (hashPassword_(password, String(r[1])) !== String(r[2])) throw new Error('Invalid student login ID or password.');
        student = studentByRegistration_(String(r[3]));
        role='student'; break;
      }
    }
    if (!student) throw new Error('Invalid student login ID or password.');
  }
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('auth_'+token, role+':'+username, 21600);
  return {ok:true, token:token, role:role, student:student};
}

function listStudents_(p) {
  requireAdmin_(p.token);
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow()<2) return {ok:true, students:[]};
  const rows=sh.getDataRange().getValues();
  const students=rows.slice(1).reverse().slice(0,150).map(r=>({
    registrationId:String(r[0]||''), name:String(r[2]||''), mobile:String(r[3]||''),
    course:String(r[8]||''), status:String(r[18]||'Enquiry received')
  }));
  return {ok:true, students:students};
}

function createStudent_(p) {
  requireAdmin_(p.token);
  const registrationId=String(p.registrationId||'').trim();
  const username=String(p.newUsername||'').trim();
  const password=String(p.newPassword||'');
  if (!registrationId || !username || password.length<8) throw new Error('Enter registration ID, login ID and a password of at least 8 characters.');
  const student=studentByRegistration_(registrationId);
  if (!student) throw new Error('Registration ID was not found in the Student Registrations sheet.');
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(ACCOUNTS_SHEET);
  if (!sh) { sh=ss.insertSheet(ACCOUNTS_SHEET); sh.appendRow(['Login ID','Salt','Password Hash','Registration ID','Created At','Active']); }
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++) if(String(rows[i][0]).toLowerCase()===username.toLowerCase()) throw new Error('That login ID is already in use.');
  const salt=Utilities.getUuid();
  sh.appendRow([username,salt,hashPassword_(password,salt),registrationId,new Date(),true]);
  return {ok:true, message:'Student account created.'};
}

function studentByRegistration_(id) {
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if(!sh || sh.getLastRow()<2) return null;
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++) if(String(rows[i][0])===id) {
    const r=rows[i];
    return {registrationId:String(r[0]),name:String(r[2]),mobile:String(r[3]),course:String(r[8]),status:String(r[18]||'Enquiry received')};
  }
  return null;
}

function requireAdmin_(token) {
  const session=CacheService.getScriptCache().get('auth_'+String(token||''));
  if(!session || !session.startsWith('admin:')) throw new Error('Your admin session expired. Please log in again.');
}

function hashPassword_(password,salt) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt)+'|'+String(password), Utilities.Charset.UTF_8);
  return bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}

function clean_(value) {
  const s=String(value||'').trim().slice(0,1000);
  return /^[=+\-@]/.test(s) ? "'"+s : s;
}

function response_(payload) {
  const json=JSON.stringify(payload).replace(/</g,'\\u003c');
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><body><script>window.parent.postMessage('+json+', "*");</script>Response sent.</body></html>'
  );
}
