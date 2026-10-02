/**
 * Dowebcode Learn & Earn V1
 * Backend: Google Apps Script Web App + Google Sheets
 *
 * NO YouTube Data API.
 * NO external speech-to-text API.
 * YouTube videos are manually configured by admins using Video IDs.
 * Browser voice input is handled by the frontend with Web Speech API where supported.
 *
 * IMPORTANT SETUP:
 * 1) Create a Google Sheet.
 * 2) Put its ID in Script Properties as SPREADSHEET_ID.
 * 3) Run setupDatabase() once.
 * 4) Deploy as Web App: Execute as Me; Who has access: Anyone.
 */

const REQUIRED_SHEETS = {
  Users: [
    'MemberID','Name','Phone','Email','PasswordHash','PasswordSalt','Verified','Status','Coins','CompletedDays',
    'LastLogin','CreatedAt','UpdatedAt','OtpHash','OtpSalt','OtpPurpose','OtpExpiresAt','OtpAttempts','LastOtpSentAt',
    'ResetTokenHash','ResetTokenExpiresAt'
  ],
  Videos: ['VideoID','YouTubeVideoID','Title','Description','DayNumber','RequiredWatchPercent','CoinReward','Active','PublishDate','CreatedAt','UpdatedAt'],
  Attendance: ['AttendanceID','MemberID','VideoID','Date','WatchPercent','LearningSubmitted','Completed','CoinAwarded','CompletedAt','UpdatedAt'],
  LearningLogs: ['LogID','MemberID','VideoID','Date','LearningText','InputMethod','WatchPercent','SubmittedAt','CoinAwarded'],
  Rewards: ['RewardID','Name','Description','RequiredDays','RequiredCoins','RewardMode','SourceCodeURL','DownloadURL','PromptText','LicenseText','Active','CreatedAt','UpdatedAt'],
  Redemptions: ['RedemptionID','MemberID','RewardID','RequiredDays','RequiredCoins','RedeemedAt','Status','DeliveryType'],
  Sessions: ['SessionToken','MemberID','Email','CreatedAt','ExpiresAt','LastSeenAt','Revoked'],
  Settings: ['Key','Value','Description','UpdatedAt'],
  AdminLogs: ['LogID','AdminEmail','Action','TargetType','TargetID','Details','CreatedAt'],
  CoinTransactions: ['TransactionID','MemberID','Amount','Reason','Note','AdminEmail','CreatedAt'],
  WatchSessions: ['WatchSessionID','MemberID','VideoID','Date','DurationSeconds','WatchSeconds','LastServerAt','LastClientTime','Status','CreatedAt','UpdatedAt']
};

const DEFAULTS = {
  APP_NAME: 'Dowebcode Learn & Earn',
  ADMIN_EMAIL: 'adnan55amir44@gmail.com', // Replace with the verified admin account if needed.
  FRONTEND_APP_URL: '', // Optional override; normally the deployed Apps Script Web App URL is used automatically.
  REQUIRED_WATCH_PERCENT: '80',
  MIN_LEARNING_TEXT_LENGTH: '20',
  DEFAULT_DAILY_COIN: '1',
  REWARD_MODE: 'BY_DAYS',
  REQUIRED_DAYS: '7',
  REQUIRED_COINS: '7'
};

const SHEETS = Object.keys(REQUIRED_SHEETS);
const SESSION_DAYS = 7;
const OTP_TTL_MINUTES = 10;
const OTP_MAX_ATTEMPTS = 5;
const OTP_RESEND_SECONDS = 60;
const RESET_TTL_MINUTES = 30;
const PASSWORD_ITERATIONS = 1200;
const WATCH_HEARTBEAT_MAX_SECONDS = 20;
const WATCH_SERVER_GRACE_SECONDS = 12;
const SESSION_CLEANUP_CHANCE = 0.08;

function doGet(e) {
  // The complete frontend is served directly by this Apps Script Web App.
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Dowebcode Learn & Earn')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Frontend calls this function through google.script.run.
// Keeping the API in the same Apps Script project removes the need for Vercel, CORS, or a proxy.
function serverRequest(req) {
  try {
    req = req || {};
    const action = String(req.action || '').trim();
    if (!action) return {success:false,message:'Missing action.'};
    return routeAction(action, req);
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return {success:false,message:safeErrorMessage(err)};
  }
}

// Optional compatibility endpoint. The normal frontend does not use doPost.
function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const req = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return jsonResponse(serverRequest(req));
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return jsonResponse({success:false,message:safeErrorMessage(err)});
  }
}

function routeAction(action, req) {
  switch (action) {
    case 'setupDatabase': return setupDatabaseApi();
    case 'signup': return signupUser(req);
    case 'resendOtp': return resendOtp(req);
    case 'verifyOTP': return verifyOTP(req);
    case 'login': return loginUser(req);
    case 'logout': return logoutUser(req);
    case 'forgotPassword': return forgotPassword(req);
    case 'resetPassword': return resetPassword(req);
    case 'getDashboard': return getDashboard(req);
    case 'recordWatchProgress': return recordWatchProgress(req);
    case 'submitLearning': return submitLearning(req);
    case 'getAttendance': return getAttendance(req);
    case 'getCoinHistory': return getCoinHistory(req);
    case 'getRewards': return getRewards(req);
    case 'getRewardAccess': return getRewardAccess(req);
    case 'updateProfile': return updateProfile(req);
    case 'adminOverview': return adminOverview(req);
    case 'adminUsers': return adminUsers(req);
    case 'adminAttendance': return adminAttendance(req);
    case 'adminLearningLogs': return adminLearningLogs(req);
    case 'adminCoinHistory': return adminCoinHistory(req);
    case 'adminVideos': return adminVideos(req);
    case 'adminRewards': return adminRewards(req);
    case 'adminSettings': return adminSettings(req);
    case 'adminLogs': return adminLogs(req);
    case 'adjustCoins': return adjustCoins(req);
    case 'adjustAttendance': return adjustAttendance(req);
    case 'saveVideo': return saveVideo(req);
    case 'setVideoStatus': return setVideoStatus(req);
    case 'saveReward': return saveReward(req);
    case 'setRewardStatus': return setRewardStatus(req);
    case 'saveAdminSettings': return saveAdminSettings(req);
    default: return {success:false,message:'Unknown action.'};
  }
}

/* =============================
   SETUP
   ============================= */

function setupDatabase() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getSpreadsheet();
    const report = ensureSheets(ss);
    seedSettings();
    seedDefaultReward();
    return report;
  } finally {
    lock.releaseLock();
  }
}

function setupDatabaseApi() {
  // This endpoint intentionally has no public authorization capability.
  // Keep setupDatabase() as the preferred manual Apps Script function.
  return {success:false,message:'Run setupDatabase() directly from the Apps Script editor once.'};
}

function ensureSheets(ss) {
  const created = [];
  const repaired = [];
  SHEETS.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      created.push(name);
    }
    const required = REQUIRED_SHEETS[name];
    const currentLastCol = Math.max(0, sheet.getLastColumn());
    const currentHeaders = currentLastCol ? sheet.getRange(1,1,1,currentLastCol).getValues()[0].map(String) : [];
    let changed = false;
    required.forEach(header => {
      if (currentHeaders.indexOf(header) === -1) {
        sheet.getRange(1, sheet.getLastColumn()+1).setValue(header);
        currentHeaders.push(header);
        changed = true;
      }
    });
    if (changed) repaired.push(name);
    if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
    if (sheet.getLastColumn() > 0) sheet.getRange(1,1,1,sheet.getLastColumn()).setFontWeight('bold');
  });
  return {success:true,message:'Database setup complete.',data:{created,repaired,sheets:SHEETS}};
}

function seedSettings() {
  const sheet = getSheet('Settings');
  const existing = readRows('Settings');
  const map = {};
  existing.forEach(r => map[String(r.Key||'')] = r);
  Object.keys(DEFAULTS).forEach(key => {
    if (!map[key]) appendObject('Settings', {Key:key,Value:DEFAULTS[key],Description:settingDescription(key),UpdatedAt:nowIso()});
  });
  // Script property can override sample admin email without putting it in frontend code.
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_EMAIL') && DEFAULTS.ADMIN_EMAIL !== 'admin@example.com') {
    props.setProperty('ADMIN_EMAIL', DEFAULTS.ADMIN_EMAIL);
  }
  if (Math.random() < SESSION_CLEANUP_CHANCE) cleanExpiredSessions();
}

function seedDefaultReward() {
  const rows = readRows('Rewards');
  if (rows.length) return;
  appendObject('Rewards', {
    RewardID:'REWARD-001',
    Name:'7-Day Website Starter',
    Description:'Complete the configured learning target to unlock the website source-code and prompt resources configured by the admin.',
    RequiredDays:7,
    RequiredCoins:7,
    RewardMode:'BY_DAYS',
    SourceCodeURL:'',
    DownloadURL:'',
    PromptText:'',
    LicenseText:'Admin must configure the usage and license terms for this reward. No ownership or resale rights are implied by this placeholder configuration.',
    Active:false,
    CreatedAt:nowIso(),
    UpdatedAt:nowIso()
  });
}

function settingDescription(key) {
  return ({
    APP_NAME:'Application display name',
    ADMIN_EMAIL:'Primary admin email. Use a real verified account.',
    FRONTEND_APP_URL:'Optional Web App URL override for password-reset links. Leave blank to use the deployed Apps Script URL automatically.',
    REQUIRED_WATCH_PERCENT:'Minimum valid watch percentage required for learning submission.',
    MIN_LEARNING_TEXT_LENGTH:'Minimum characters required in a learning response.',
    DEFAULT_DAILY_COIN:'Fallback daily coin reward when a video does not specify a reward.',
    REWARD_MODE:'Default reward eligibility mode: BY_DAYS, BY_COINS, or BY_DAYS_AND_COINS.',
    REQUIRED_DAYS:'Default reward days threshold.',
    REQUIRED_COINS:'Default reward coin threshold.'
  })[key] || '';
}

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('SPREADSHEET_ID is not configured. Add it in Apps Script → Project Settings → Script Properties.');
}

function getSheet(name) {
  const sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Required sheet is missing: '+name+'. Run setupDatabase().');
  return sheet;
}

function readRows(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const values = sheet.getRange(1,1,lastRow,lastCol).getValues();
  const headers = values[0].map(h=>String(h));
  return values.slice(1).filter(row=>row.some(v=>v!=='' && v!==null)).map(row=>rowToObject(headers,row));
}

function rowToObject(headers,row) {
  const o = {};
  headers.forEach((h,i)=>o[h]=row[i]);
  return o;
}

function objectToRow(sheetName,obj) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);
  return headers.map(h=>Object.prototype.hasOwnProperty.call(obj,h)?obj[h]:'');
}

function appendObject(sheetName,obj) {
  const sheet = getSheet(sheetName);
  const row = objectToRow(sheetName,obj);
  sheet.getRange(sheet.getLastRow()+1,1,1,row.length).setValues([row]);
  return sheet.getLastRow();
}

function findRowIndexBy(sheetName,column,value,normalize) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(String);
  const col = headers.indexOf(column);
  if (col === -1) return -1;
  const values = sheet.getRange(2,col+1,lastRow-1,1).getValues().map(r=>r[0]);
  const target = normalize ? normalize(value) : value;
  for (let i=0;i<values.length;i++) {
    const current = normalize ? normalize(values[i]) : values[i];
    if (current === target) return i+2;
  }
  return -1;
}

function updateRowObject(sheetName,rowIndex,patch) {
  const sheet = getSheet(sheetName);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(String);
  const row = sheet.getRange(rowIndex,1,1,lastCol).getValues()[0];
  headers.forEach((h,i)=>{ if (Object.prototype.hasOwnProperty.call(patch,h)) row[i]=patch[h]; });
  sheet.getRange(rowIndex,1,1,lastCol).setValues([row]);
}

/* =============================
   AUTH / PASSWORDS / SESSION
   ============================= */

function signupUser(req) {
  ensureDatabaseReady();
  const name = cleanText(req.name,80);
  const phone = cleanText(req.phone,30);
  const email = normalizeEmail(req.email);
  const password = String(req.password||'');
  validateRegistration(name,phone,email,password);
  if (findRowIndexBy('Users','Email',email,normalizeEmail) !== -1) return {success:false,message:'An account with this email already exists. Please sign in or use password reset.'};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const salt = randomSalt();
    const hash = hashPassword(password,salt);
    const otp = generateOtp();
    const otpSalt = randomSalt();
    const now = nowIso();
    appendObject('Users',{
      MemberID:'',Name:name,Phone:phone,Email:email,PasswordHash:hash,PasswordSalt:salt,Verified:false,Status:'active',Coins:0,CompletedDays:0,
      LastLogin:'',CreatedAt:now,UpdatedAt:now,OtpHash:hashOtp(otp,otpSalt),OtpSalt:otpSalt,OtpPurpose:'signup',OtpExpiresAt:isoMinutesFromNow(OTP_TTL_MINUTES),OtpAttempts:0,
      LastOtpSentAt:now,ResetTokenHash:'',ResetTokenExpiresAt:''
    });
  } finally { lock.releaseLock(); }
  try { sendVerificationEmail(email,name,otp); } catch (mailErr) { console.error(mailErr); }
  return {success:true,message:'Account created. Check your email for the 6-digit verification code.'};
}

function validateRegistration(name,phone,email,password){
  if (!name || name.length < 2) throw new Error('Please enter your full name.');
  if (!phone || phone.length < 7) throw new Error('Please enter a valid phone number.');
  if (!isValidEmail(email)) throw new Error('Please enter a valid email address.');
  if (password.length < 8) throw new Error('Password must be at least 8 characters.');
}

function resendOtp(req) {
  ensureDatabaseReady();
  const email = normalizeEmail(req.email);
  const row = findRowIndexBy('Users','Email',email,normalizeEmail);
  if (row === -1) return {success:false,message:'No account was found for this email.'};
  const user = rowObjectAt('Users',row);
  if (truthy(user.Verified)) return {success:false,message:'This account is already verified. Please sign in.'};
  const last = user.LastOtpSentAt ? new Date(user.LastOtpSentAt).getTime() : 0;
  if (Date.now()-last < OTP_RESEND_SECONDS*1000) return {success:false,message:'Please wait a moment before requesting another code.'};
  const otp = generateOtp(); const salt=randomSalt();
  updateRowObject('Users',row,{OtpHash:hashOtp(otp,salt),OtpSalt:salt,OtpPurpose:'signup',OtpExpiresAt:isoMinutesFromNow(OTP_TTL_MINUTES),OtpAttempts:0,LastOtpSentAt:nowIso(),UpdatedAt:nowIso()});
  sendVerificationEmail(email,String(user.Name||''),otp);
  return {success:true,message:'A new verification code has been sent.'};
}

function verifyOTP(req) {
  ensureDatabaseReady();
  const email=normalizeEmail(req.email),otp=String(req.otp||'').trim();
  if (!/^\d{6}$/.test(otp)) return {success:false,message:'Enter the 6-digit verification code.'};
  const row=findRowIndexBy('Users','Email',email,normalizeEmail);if(row===-1)return {success:false,message:'Account not found.'};
  const user=rowObjectAt('Users',row);
  if(truthy(user.Verified))return {success:false,message:'This account is already verified.'};
  if(!user.OtpExpiresAt || new Date(user.OtpExpiresAt).getTime()<Date.now())return {success:false,message:'That code has expired. Request a new one.'};
  const attempts=Number(user.OtpAttempts||0);if(attempts>=OTP_MAX_ATTEMPTS)return {success:false,message:'Too many incorrect attempts. Request a new code.'};
  const valid=constantTimeEqual(hashOtp(otp,String(user.OtpSalt||'')),String(user.OtpHash||''));
  if(!valid){updateRowObject('Users',row,{OtpAttempts:attempts+1,UpdatedAt:nowIso()});return {success:false,message:'Incorrect verification code.'};}
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  let memberId;
  try{memberId=generateMemberID();updateRowObject('Users',row,{MemberID:memberId,Verified:true,Status:'active',OtpHash:'',OtpSalt:'',OtpPurpose:'',OtpExpiresAt:'',OtpAttempts:0,UpdatedAt:nowIso()});}finally{lock.releaseLock()}
  const session=createSession(memberId,email);
  return {success:true,message:'Email verified successfully.',data:{sessionToken:session.token,user:publicUser(rowObjectAt('Users',row))}};
}

function loginUser(req){
  ensureDatabaseReady();
  const email=normalizeEmail(req.email),password=String(req.password||'');
  if(!isValidEmail(email)||!password)return {success:false,message:'Enter your email and password.'};
  const row=findRowIndexBy('Users','Email',email,normalizeEmail);if(row===-1)return {success:false,message:'Incorrect email or password.'};
  const user=rowObjectAt('Users',row);if(!truthy(user.Verified))return {success:false,message:'Please verify your email before signing in.'};if(String(user.Status||'active')!=='active')return {success:false,message:'This account is not active.'};
  if(!verifyPassword(password,String(user.PasswordSalt||''),String(user.PasswordHash||'')))return {success:false,message:'Incorrect email or password.'};
  updateRowObject('Users',row,{LastLogin:nowIso(),UpdatedAt:nowIso()});
  const session=createSession(user.MemberID,email);
  return {success:true,message:'Login successful.',data:{sessionToken:session.token,user:publicUser(rowObjectAt('Users',row))}};
}

function logoutUser(req){if(req.sessionToken)revokeSession(req.sessionToken);return {success:true,message:'Logged out.'};}

function forgotPassword(req){
  ensureDatabaseReady();const email=normalizeEmail(req.email);if(!isValidEmail(email))return {success:false,message:'Enter a valid email address.'};
  const row=findRowIndexBy('Users','Email',email,normalizeEmail);
  // Do not reveal whether an email exists.
  if(row===-1)return {success:true,message:'If that email belongs to an account, a secure reset link has been sent.'};
  const user=rowObjectAt('Users',row);
  if(!truthy(user.Verified))return {success:true,message:'If that email belongs to an account, a secure reset link has been sent.'};
  const token=generateToken();const salt=randomSalt();const tokenHash=hashOtp(token,salt);
  // Store token hash and salt in ResetTokenHash as salt:hash. The raw token exists only inside the email link.
  updateRowObject('Users',row,{ResetTokenHash:salt+':'+tokenHash,ResetTokenExpiresAt:isoMinutesFromNow(RESET_TTL_MINUTES),UpdatedAt:nowIso()});
  const configuredUrl=String(getSetting('FRONTEND_APP_URL')||PropertiesService.getScriptProperties().getProperty('FRONTEND_APP_URL')||'').replace(/\/$/,'');
  const appUrl=configuredUrl||String(ScriptApp.getService().getUrl()||'').replace(/\/$/,'');
  if(!appUrl) throw new Error('The Apps Script Web App URL is not available yet. Deploy this project as a Web App first.');
  const link=appUrl+'?reset='+encodeURIComponent(token);
  sendPasswordResetEmail(email,String(user.Name||''),link);
  return {success:true,message:'If that email belongs to an account, a secure reset link has been sent.'};
}

function resetPassword(req){
  ensureDatabaseReady();const token=String(req.token||'');const password=String(req.password||'');if(token.length<20)return {success:false,message:'Reset token is missing or invalid.'};if(password.length<8)return {success:false,message:'Password must be at least 8 characters.'};
  const users=readRows('Users');let match=null;
  for(let i=0;i<users.length;i++){const u=users[i];if(!u.ResetTokenHash||!u.ResetTokenExpiresAt)continue;if(new Date(u.ResetTokenExpiresAt).getTime()<Date.now())continue;const parts=String(u.ResetTokenHash).split(':');if(parts.length!==2)continue;if(constantTimeEqual(hashOtp(token,parts[0]),parts[1])){match={row:i+2,user:u};break;}}
  if(!match)return {success:false,message:'That reset link is invalid or expired.'};
  const salt=randomSalt(),hash=hashPassword(password,salt);
  updateRowObject('Users',match.row,{PasswordSalt:salt,PasswordHash:hash,ResetTokenHash:'',ResetTokenExpiresAt:'',UpdatedAt:nowIso()});
  return {success:true,message:'Password updated. You can now sign in.'};
}

function createSession(memberId,email){
  const token=generateToken();const created=new Date();const expires=new Date(created.getTime()+SESSION_DAYS*86400000);
  appendObject('Sessions',{SessionToken:token,MemberID:memberId,Email:email,CreatedAt:created.toISOString(),ExpiresAt:expires.toISOString(),LastSeenAt:created.toISOString(),Revoked:false});
  return {token,expiresAt:expires.toISOString()};
}

function requireSession(token){
  const t=String(token||'');if(!t)throw new Error('Session required.');
  const row=findRowIndexBy('Sessions','SessionToken',t,String);if(row===-1)throw new Error('Session expired. Please sign in again.');
  const s=rowObjectAt('Sessions',row);if(truthy(s.Revoked)||!s.ExpiresAt||new Date(s.ExpiresAt).getTime()<Date.now()){revokeSession(t);throw new Error('Session expired. Please sign in again.');}
  updateRowObject('Sessions',row,{LastSeenAt:nowIso()});
  const userRow=findRowIndexBy('Users','MemberID',s.MemberID,String);if(userRow===-1)throw new Error('Account not found.');
  const user=rowObjectAt('Users',userRow);if(String(user.Status||'active')!=='active')throw new Error('This account is not active.');
  return {session:s, user, userRow};
}

function requireAdmin(token){
  const auth=requireSession(token);if(!isAdminEmail(auth.user.Email))throw new Error('Admin authorization required.');auth.userRow=auth.userRow;return auth;
}

function isAdminEmail(email){
  const configured=String(PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')||getSetting('ADMIN_EMAIL')||'').toLowerCase().trim();
  return !!configured && normalizeEmail(email)===normalizeEmail(configured);
}

function revokeSession(token){const row=findRowIndexBy('Sessions','SessionToken',String(token),String);if(row!==-1)updateRowObject('Sessions',row,{Revoked:true,LastSeenAt:nowIso()});}

function cleanExpiredSessions(){
  const sheet=getSheet('Sessions'),rows=readRows('Sessions'),now=Date.now();
  rows.forEach((s,idx)=>{if(s.ExpiresAt && new Date(s.ExpiresAt).getTime()<now && !truthy(s.Revoked))updateRowObject('Sessions',idx+2,{Revoked:true,LastSeenAt:nowIso()});});
}

function publicUser(user){return {memberId:String(user.MemberID||''),name:String(user.Name||''),phone:String(user.Phone||''),email:String(user.Email||''),verified:truthy(user.Verified),status:String(user.Status||'active'),coins:Number(user.Coins||0),completedDays:Number(user.CompletedDays||0),createdAt:user.CreatedAt||'',lastLogin:user.LastLogin||'',isAdmin:isAdminEmail(user.Email)};}

/* =============================
   DASHBOARD / VIDEO
   ============================= */

function getDashboard(req){
  const auth=requireSession(req.sessionToken);const user=auth.user;
  const settings=getPublicSettings();
  const todayKey=dateKey(new Date());
  const activeVideos=getActiveVideos();
  const dayNumber=Math.max(1,Number(user.CompletedDays||0)+1);
  let video=activeVideos.find(v=>Number(v.DayNumber)===dayNumber);
  if(!video && activeVideos.length===1 && dayNumber===1)video=activeVideos[0];
  const todayAttendance=video?findAttendance(auth.user.MemberID,video.VideoID,todayKey):null;
  let watch=null;
  if(video && !truthy(todayAttendance?.Completed)) watch=startOrGetWatchSession(auth.user.MemberID,video,todayKey);
  else if(video) watch=getWatchSnapshot(auth.user.MemberID,video.VideoID,todayKey);
  const rewards=getRewardEvaluations(user,activeVideos);
  const featuredReward=rewards.find(r=>r.active)||null;
  return {success:true,message:'Dashboard loaded.',data:{user:publicUser(user),settings,video:video?publicVideo(video):null,todayAttendance:todayAttendance?publicAttendance(todayAttendance):null,watch,rewards,featuredReward}};
}

function getActiveVideos(){
  return readRows('Videos').filter(v=>truthy(v.Active)).sort((a,b)=>Number(a.DayNumber||0)-Number(b.DayNumber||0));
}
function getActiveVideoById(videoId){return readRows('Videos').find(v=>String(v.VideoID)===String(videoId)&&truthy(v.Active));}
function publicVideo(v){return {videoId:String(v.VideoID),youtubeVideoId:String(v.YouTubeVideoID),title:String(v.Title||''),description:String(v.Description||''),dayNumber:Number(v.DayNumber||0),requiredWatchPercent:clamp(Number(v.RequiredWatchPercent||getSettingNumber('REQUIRED_WATCH_PERCENT',80)),1,100),coinReward:Number(v.CoinReward||getSettingNumber('DEFAULT_DAILY_COIN',1))};}

function startOrGetWatchSession(memberId,video,dateKey){
  const rows=readRows('WatchSessions');const existing=rows.find(w=>String(w.MemberID)===String(memberId)&&String(w.VideoID)===String(video.VideoID)&&String(w.Date)===String(dateKey)&&String(w.Status||'active')==='active');
  if(existing)return watchPublic(existing,video);
  const sessionId='WATCH-'+Utilities.getUuid();const row={WatchSessionID:sessionId,MemberID:memberId,VideoID:video.VideoID,Date:dateKey,DurationSeconds:0,WatchSeconds:0,LastServerAt:nowIso(),LastClientTime:0,Status:'active',CreatedAt:nowIso(),UpdatedAt:nowIso()};appendObject('WatchSessions',row);return watchPublic(row,video);
}
function getWatchSnapshot(memberId,videoId,dateKey){const rows=readRows('WatchSessions').filter(w=>String(w.MemberID)===String(memberId)&&String(w.VideoID)===String(videoId)&&String(w.Date)===String(dateKey));if(!rows.length)return null;rows.sort((a,b)=>new Date(b.UpdatedAt||0)-new Date(a.UpdatedAt||0));const video=readRows('Videos').find(v=>String(v.VideoID)===String(videoId));return video?watchPublic(rows[0],video):null;}
function watchPublic(w,video){const duration=Number(w.DurationSeconds||0),seconds=Number(w.WatchSeconds||0),required=clamp(Number(video.RequiredWatchPercent||80),1,100);const pct=duration?Math.min(100,(seconds/duration)*100):0;return {sessionId:String(w.WatchSessionID),watchSeconds:seconds,durationSeconds:duration,watchPercent:pct,requiredWatchPercent:required,eligible:duration>0&&pct>=required};}

function recordWatchProgress(req){
  const auth=requireSession(req.sessionToken);const memberId=auth.user.MemberID;const videoId=String(req.videoId||'');const watchSessionId=String(req.watchSessionId||'');
  if(!videoId||!watchSessionId)throw new Error('Watch session information is missing.');
  const video=getActiveVideoById(videoId);if(!video)throw new Error('This lesson is no longer active.');
  const sessionRow=findRowIndexBy('WatchSessions','WatchSessionID',watchSessionId,String);if(sessionRow===-1)throw new Error('Watch session expired. Reload the lesson.');
  const watch= rowObjectAt('WatchSessions',sessionRow);if(String(watch.MemberID)!==String(memberId)||String(watch.VideoID)!==videoId)throw new Error('Invalid watch session.');
  const today=dateKey(new Date());if(String(watch.Date)!==today)throw new Error('This lesson session is no longer active for today.');
  const duration=Math.max(0,Number(req.duration||0));const clientTime=Math.max(0,Number(req.clientCurrentTime||0));if(duration>0&&clientTime>duration+5)throw new Error('Invalid playback position.');
  const rawDelta=Math.max(0,Number(req.deltaSeconds||0));if(!isFinite(rawDelta)||rawDelta>WATCH_HEARTBEAT_MAX_SECONDS)throw new Error('Invalid watch progress update.');
  const lastServer=new Date(watch.LastServerAt||watch.UpdatedAt||Date.now()).getTime();const elapsed=Math.max(0,(Date.now()-lastServer)/1000);const permitted=Math.min(WATCH_HEARTBEAT_MAX_SECONDS,elapsed+WATCH_SERVER_GRACE_SECONDS);const delta=Math.min(rawDelta,Math.max(0,permitted));
  const oldDuration=Math.max(0,Number(watch.DurationSeconds||0));const finalDuration=Math.max(oldDuration,duration);const oldWatch=Math.max(0,Number(watch.WatchSeconds||0));const newWatch=Math.min(finalDuration||Number.MAX_SAFE_INTEGER,oldWatch+delta);
  updateRowObject('WatchSessions',sessionRow,{DurationSeconds:finalDuration,WatchSeconds:newWatch,LastServerAt:nowIso(),LastClientTime:clientTime,Status:'active',UpdatedAt:nowIso()});
  const required=clamp(Number(video.RequiredWatchPercent||80),1,100);const pct=finalDuration?Math.min(100,(newWatch/finalDuration)*100):0;return {success:true,message:'Watch progress saved.',data:{watchSeconds:newWatch,durationSeconds:finalDuration,watchPercent:pct,eligible:finalDuration>0&&pct>=required}};
}

/* =============================
   LEARNING / ATTENDANCE / COINS
   ============================= */

function submitLearning(req){
  const auth=requireSession(req.sessionToken);const memberId=auth.user.MemberID;const videoId=String(req.videoId||'');const learningText=cleanTextPreserve(req.learningText,10000).trim();const inputMethod=req.inputMethod==='voice'?'voice':'typed';
  if(!videoId||!learningText)throw new Error('Learning response is required.');
  const minLen=getSettingNumber('MIN_LEARNING_TEXT_LENGTH',20);if(learningText.length<minLen)throw new Error('Please write at least '+minLen+' characters in your own words.');
  const video=getActiveVideoById(videoId);if(!video)throw new Error('The assigned lesson is not available.');
  const today=dateKey(new Date());const existing=findAttendance(memberId,videoId,today);if(existing&&truthy(existing.Completed))return {success:true,message:'Today\'s attendance was already completed.',data:{alreadyCompleted:true,user:publicUser(auth.user)}};
  const watch=getWatchSnapshot(memberId,videoId,today);if(!watch||!watch.eligible)throw new Error('Please complete the required watch progress before submitting your learning.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const recheck=findAttendance(memberId,videoId,today);if(recheck&&truthy(recheck.Completed))return {success:true,message:'Today\'s attendance was already completed.',data:{alreadyCompleted:true,user:publicUser(auth.user)}};
    const coinReward=Math.max(0,Number(video.CoinReward||getSettingNumber('DEFAULT_DAILY_COIN',1)));
    const attendanceId='ATT-'+Utilities.getUuid();
    appendObject('LearningLogs',{LogID:'LOG-'+Utilities.getUuid(),MemberID:memberId,VideoID:videoId,Date:today,LearningText:learningText,InputMethod:inputMethod,WatchPercent:watch.watchPercent,SubmittedAt:nowIso(),CoinAwarded:coinReward});
    if(recheck){
      updateRowObject('Attendance',findRowIndexBy('AttendanceID','AttendanceID',recheck.AttendanceID,String),{WatchPercent:watch.watchPercent,LearningSubmitted:true,Completed:true,CoinAwarded:coinReward,CompletedAt:nowIso(),UpdatedAt:nowIso()});
    } else {
      appendObject('Attendance',{AttendanceID:attendanceId,MemberID:memberId,VideoID:videoId,Date:today,WatchPercent:watch.watchPercent,LearningSubmitted:true,Completed:true,CoinAwarded:coinReward,CompletedAt:nowIso(),UpdatedAt:nowIso()});
    }
    if(coinReward!==0) addCoinTransaction(memberId,coinReward,'Daily Lesson','Validated lesson completion','SYSTEM');
    const newCompletedDays=recalculateCompletedDays(memberId);
    const userRow=findRowIndexBy('Users','MemberID',memberId,String);const current= rowObjectAt('Users',userRow);updateRowObject('Users',userRow,{Coins:Number(current.Coins||0)+coinReward,CompletedDays:newCompletedDays,UpdatedAt:nowIso()});
    const updated=rowObjectAt('Users',userRow);return {success:true,message:'Learning submitted. Attendance completed and your coin reward was issued.',data:{user:publicUser(updated),attendance:{videoId,date:today,watchPercent:watch.watchPercent,coinAwarded:coinReward},rewardUnlocked:getRewardEvaluations(updated,getActiveVideos()).some(r=>r.eligible)}};
  }finally{lock.releaseLock()}
}

function findAttendance(memberId,videoId,dateStr){return readRows('Attendance').find(a=>String(a.MemberID)===String(memberId)&&String(a.VideoID)===String(videoId)&&String(a.Date)===String(dateStr))||null;}
function publicAttendance(a){return {attendanceId:String(a.AttendanceID),memberId:String(a.MemberID),videoId:String(a.VideoID),date:String(a.Date),watchPercent:Number(a.WatchPercent||0),learningSubmitted:truthy(a.LearningSubmitted),completed:truthy(a.Completed),coinAwarded:Number(a.CoinAwarded||0),completedAt:a.CompletedAt||''};}
function recalculateCompletedDays(memberId){const done=readRows('Attendance').filter(a=>String(a.MemberID)===String(memberId)&&truthy(a.Completed));const unique={};done.forEach(a=>unique[String(a.VideoID)]=true);return Object.keys(unique).length;}
function addCoinTransaction(memberId,amount,reason,note,adminEmail){appendObject('CoinTransactions',{TransactionID:'CT-'+Utilities.getUuid(),MemberID:memberId,Amount:Number(amount),Reason:String(reason||''),Note:String(note||''),AdminEmail:String(adminEmail||'SYSTEM'),CreatedAt:nowIso()});}
function getAttendance(req){const auth=requireSession(req.sessionToken);const records=readRows('Attendance').filter(a=>String(a.MemberID)===String(auth.user.MemberID)).sort((a,b)=>new Date(b.Date||0)-new Date(a.Date||0));const videos=readRows('Videos');const map={};videos.forEach(v=>map[v.VideoID]=v.Title);return {success:true,message:'Attendance loaded.',data:{records:records.map(a=>({...publicAttendance(a),videoTitle:map[a.VideoID]||a.VideoID}))}};}
function getCoinHistory(req){const auth=requireSession(req.sessionToken);const records=readRows('CoinTransactions').filter(r=>String(r.MemberID)===String(auth.user.MemberID)).sort((a,b)=>new Date(b.CreatedAt||0)-new Date(a.CreatedAt||0));return {success:true,message:'Coin history loaded.',data:{records:records.map(r=>({transactionId:r.TransactionID,memberId:r.MemberID,amount:Number(r.Amount||0),reason:String(r.Reason||''),note:String(r.Note||''),createdAt:r.CreatedAt||''}) )}};}

/* =============================
   REWARDS
   ============================= */

function getRewardEvaluations(user,activeVideos){
  return readRows('Rewards').filter(r=>truthy(r.Active)).map(r=>evaluateReward(r,user));
}
function evaluateReward(reward,user){
  const mode=normalizeMode(reward.RewardMode||getSetting('REWARD_MODE')||'BY_DAYS');const requiredDays=Math.max(0,Number(reward.RequiredDays||getSettingNumber('REQUIRED_DAYS',7)));const requiredCoins=Math.max(0,Number(reward.RequiredCoins||getSettingNumber('REQUIRED_COINS',7)));const days=Number(user.CompletedDays||0);const coins=Number(user.Coins||0);const eligible=mode==='BY_COINS'?coins>=requiredCoins:mode==='BY_DAYS_AND_COINS'?days>=requiredDays&&coins>=requiredCoins:days>=requiredDays;let percent=0,progressText='';if(mode==='BY_COINS'){percent=requiredCoins?Math.min(100,coins/requiredCoins*100):100;progressText=`${coins} / ${requiredCoins} coins`;}else if(mode==='BY_DAYS_AND_COINS'){percent=Math.min(100,requiredDays?days/requiredDays*100:100,requiredCoins?coins/requiredCoins*100:100);progressText=`${days}/${requiredDays} days · ${coins}/${requiredCoins} coins`;}else{percent=requiredDays?Math.min(100,days/requiredDays*100):100;progressText=`${days} / ${requiredDays} days`;}return {rewardId:String(reward.RewardID),name:String(reward.Name||''),description:String(reward.Description||''),rewardMode:mode,requiredDays,requiredCoins,eligible,progressPercent:percent,progressText,lockReason:eligible?'Unlocked':(mode==='BY_COINS'?`${Math.max(0,requiredCoins-coins)} more coins needed`:mode==='BY_DAYS_AND_COINS'?`${Math.max(0,requiredDays-days)} more days and coin target needed`:`${Math.max(0,requiredDays-days)} more days needed`),active:truthy(reward.Active)};
}
function getRewards(req){const auth=requireSession(req.sessionToken);const rewards=getRewardEvaluations(auth.user,getActiveVideos());return {success:true,message:'Rewards loaded.',data:{rewards}};}
function getRewardAccess(req){const auth=requireSession(req.sessionToken);const reward=readRows('Rewards').find(r=>String(r.RewardID)===String(req.rewardId)&&truthy(r.Active));if(!reward)throw new Error('Reward not found.');const evald=evaluateReward(reward,auth.user);if(!evald.eligible)throw new Error('This reward is still locked: '+evald.lockReason+'.');
  const existing=readRows('Redemptions').find(r=>String(r.MemberID)===String(auth.user.MemberID)&&String(r.RewardID)===String(reward.RewardID)&&String(r.Status)==='redeemed');if(!existing)appendObject('Redemptions',{RedemptionID:'RED-'+Utilities.getUuid(),MemberID:auth.user.MemberID,RewardID:reward.RewardID,RequiredDays:evald.requiredDays,RequiredCoins:evald.requiredCoins,RedeemedAt:nowIso(),Status:'redeemed',DeliveryType:reward.SourceCodeURL&&reward.PromptText?'source+prompt':reward.SourceCodeURL?'source-code':reward.PromptText?'prompt':'configured-resource'});
  return {success:true,message:'Reward access approved.',data:{name:String(reward.Name||''),description:String(reward.Description||''),sourceCodeUrl:safeHttpUrl(reward.SourceCodeURL),downloadUrl:safeHttpUrl(reward.DownloadURL),promptText:String(reward.PromptText||''),licenseText:String(reward.LicenseText||''),rewardId:reward.RewardID}};
}

/* =============================
   USER PROFILE
   ============================= */

function updateProfile(req){const auth=requireSession(req.sessionToken);const name=cleanText(req.name,80),phone=cleanText(req.phone,30);if(name.length<2)throw new Error('Please enter your name.');if(phone.length<7)throw new Error('Please enter a valid phone number.');updateRowObject('Users',auth.userRow,{Name:name,Phone:phone,UpdatedAt:nowIso()});return {success:true,message:'Profile updated successfully.',data:{user:publicUser(rowObjectAt('Users',auth.userRow))}};}

/* =============================
   ADMIN
   ============================= */

function adminOverview(req){requireAdmin(req.sessionToken);const users=readRows('Users'),today=dateKey(new Date()),attendance=readRows('Attendance').filter(a=>String(a.Date)===today&&truthy(a.Completed)),videos=readRows('Videos'),rewards=readRows('Rewards').filter(r=>truthy(r.Active));return {success:true,message:'Admin overview loaded.',data:{users:users.length,attendanceToday:attendance.length,videos:videos.length,rewards:rewards.length}};}
function adminUsers(req){requireAdmin(req.sessionToken);const users=readRows('Users').sort((a,b)=>new Date(b.CreatedAt||0)-new Date(a.CreatedAt||0));return {success:true,message:'Users loaded.',data:{users:users.map(publicUser)}};}
function adminAttendance(req){requireAdmin(req.sessionToken);const videos=readRows('Videos'),map={};videos.forEach(v=>map[v.VideoID]=v.Title);const records=readRows('Attendance').sort((a,b)=>new Date(b.Date||0)-new Date(a.Date||0));return {success:true,message:'Attendance loaded.',data:{records:records.map(a=>({...publicAttendance(a),dateKey:String(a.Date||''),videoTitle:map[a.VideoID]||a.VideoID}))}};}
function adminLearningLogs(req){requireAdmin(req.sessionToken);const records=readRows('LearningLogs').sort((a,b)=>new Date(b.SubmittedAt||0)-new Date(a.SubmittedAt||0));return {success:true,message:'Learning logs loaded.',data:{records:records.map(r=>({logId:r.LogID,memberId:r.MemberID,videoId:r.VideoID,date:r.Date,learningText:String(r.LearningText||''),inputMethod:r.InputMethod,watchPercent:Number(r.WatchPercent||0),submittedAt:r.SubmittedAt,coinAwarded:Number(r.CoinAwarded||0)}))}};}
function adminCoinHistory(req){requireAdmin(req.sessionToken);const records=readRows('CoinTransactions').sort((a,b)=>new Date(b.CreatedAt||0)-new Date(a.CreatedAt||0));return {success:true,message:'Coin transactions loaded.',data:{records:records.map(r=>({transactionId:r.TransactionID,memberId:r.MemberID,amount:Number(r.Amount||0),reason:r.Reason,note:r.Note,adminEmail:r.AdminEmail,createdAt:r.CreatedAt}))}};}
function adminVideos(req){requireAdmin(req.sessionToken);return {success:true,message:'Videos loaded.',data:{videos:readRows('Videos').sort((a,b)=>Number(a.DayNumber||0)-Number(b.DayNumber||0)).map(publicAdminVideo)}};}
function publicAdminVideo(v){return {videoId:String(v.VideoID),youtubeVideoId:String(v.YouTubeVideoID),title:String(v.Title||''),description:String(v.Description||''),dayNumber:Number(v.DayNumber||0),requiredWatchPercent:Number(v.RequiredWatchPercent||80),coinReward:Number(v.CoinReward||0),active:truthy(v.Active),publishDate:v.PublishDate||'',createdAt:v.CreatedAt||'',updatedAt:v.UpdatedAt||''};}
function adminRewards(req){requireAdmin(req.sessionToken);return {success:true,message:'Rewards loaded.',data:{rewards:readRows('Rewards').sort((a,b)=>String(a.Name).localeCompare(String(b.Name))).map(publicAdminReward)}};}
function publicAdminReward(r){return {rewardId:String(r.RewardID),name:String(r.Name||''),description:String(r.Description||''),requiredDays:Number(r.RequiredDays||0),requiredCoins:Number(r.RequiredCoins||0),rewardMode:normalizeMode(r.RewardMode),sourceCodeUrl:String(r.SourceCodeURL||''),downloadUrl:String(r.DownloadURL||''),promptText:String(r.PromptText||''),licenseText:String(r.LicenseText||''),active:truthy(r.Active),createdAt:r.CreatedAt||'',updatedAt:r.UpdatedAt||''};}
function adminSettings(req){requireAdmin(req.sessionToken);const rows=readRows('Settings');const o={};rows.forEach(r=>o[String(r.Key||'')]=String(r.Value??''));return {success:true,message:'Settings loaded.',data:{settings:o}};}
function adminLogs(req){requireAdmin(req.sessionToken);const records=readRows('AdminLogs').sort((a,b)=>new Date(b.CreatedAt||0)-new Date(a.CreatedAt||0));return {success:true,message:'Admin logs loaded.',data:{records:records.map(r=>({logId:r.LogID,adminEmail:r.AdminEmail,action:r.Action,targetType:r.TargetType,targetId:r.TargetID,details:r.Details,createdAt:r.CreatedAt}))}};}

function adjustCoins(req){const auth=requireAdmin(req.sessionToken);const memberId=String(req.memberId||'').trim();const amount=Number(req.amount);const reason=cleanText(req.reason,120);const note=cleanText(req.note,500);if(!memberId)throw new Error('Member ID is required.');if(!isFinite(amount)||amount===0)throw new Error('Coin adjustment cannot be zero.');if(!reason)throw new Error('A reason is required.');const row=findRowIndexBy('Users','MemberID',memberId,String);if(row===-1)throw new Error('User not found.');const lock=LockService.getScriptLock();lock.waitLock(30000);try{const u=rowObjectAt('Users',row),current=Number(u.Coins||0),next=Math.max(0,current+amount);if(current+amount<0)throw new Error('Coin balance cannot go below zero.');updateRowObject('Users',row,{Coins:next,UpdatedAt:nowIso()});addCoinTransaction(memberId,amount,reason,note,auth.user.Email);logAdmin(auth.user.Email,'ADJUST_COINS','USER',memberId,`Amount ${amount}; reason: ${reason}; note: ${note}`);return {success:true,message:'Coin balance updated.',data:{coins:next}};}finally{lock.releaseLock()}}

function adjustAttendance(req){const auth=requireAdmin(req.sessionToken);const attendanceId=String(req.attendanceId||'');const completed=!!req.completed;const reason=cleanText(req.reason,500);if(!attendanceId||!reason)throw new Error('Attendance ID and reason are required.');const row=findRowIndexBy('AttendanceID','AttendanceID',attendanceId,String);if(row===-1)throw new Error('Attendance record not found.');const a=rowObjectAt('Attendance',row);updateRowObject('Attendance',row,{Completed:completed,UpdatedAt:nowIso(),CompletedAt:completed?(a.CompletedAt||nowIso()):''});const userRow=findRowIndexBy('MemberID','MemberID',a.MemberID,String);if(userRow!==-1){const user=rowObjectAt('Users',userRow);updateRowObject('Users',userRow,{CompletedDays:recalculateCompletedDays(a.MemberID),UpdatedAt:nowIso()});}
  logAdmin(auth.user.Email,'ADJUST_ATTENDANCE','ATTENDANCE',attendanceId,`Completed=${completed}; reason: ${reason}`);return {success:true,message:'Attendance correction saved.'};}

function saveVideo(req){const auth=requireAdmin(req.sessionToken);const videoId=String(req.videoId||'').trim();const youtubeId=cleanText(req.youtubeVideoId,50);const title=cleanText(req.title,160);const description=cleanText(req.description,1000);const day=Number(req.dayNumber);const watch=Number(req.requiredWatchPercent);const coin=Number(req.coinReward);const active=!!req.active;if(!/^[A-Za-z0-9_-]{6,20}$/.test(youtubeId))throw new Error('Enter a valid YouTube Video ID.');if(!title||!Number.isInteger(day)||day<1)throw new Error('Video title and day number are required.');if(!isFinite(watch)||watch<1||watch>100)throw new Error('Required watch % must be between 1 and 100.');if(!isFinite(coin)||coin<0)throw new Error('Coin reward must be zero or greater.');const duplicate=readRows('Videos').find(v=>String(v.YouTubeVideoID)===youtubeId&&String(v.VideoID)!==videoId&&truthy(v.Active));if(duplicate)throw new Error('That YouTube Video ID is already configured as an active lesson.');const now=nowIso();if(videoId){const row=findRowIndexBy('VideoID','VideoID',videoId,String);if(row===-1)throw new Error('Video not found.');updateRowObject('Videos',row,{YouTubeVideoID:youtubeId,Title:title,Description:description,DayNumber:day,RequiredWatchPercent:watch,CoinReward:coin,Active:active,UpdatedAt:now});logAdmin(auth.user.Email,'UPDATE_VIDEO','VIDEO',videoId,`Day ${day}; active ${active}`);return {success:true,message:'Video updated.'};}const newId='VID-'+Utilities.getUuid().slice(0,8);appendObject('Videos',{VideoID:newId,YouTubeVideoID:youtubeId,Title:title,Description:description,DayNumber:day,RequiredWatchPercent:watch,CoinReward:coin,Active:active,PublishDate:'',CreatedAt:now,UpdatedAt:now});logAdmin(auth.user.Email,'CREATE_VIDEO','VIDEO',newId,`Day ${day}; active ${active}`);return {success:true,message:'Video added.'};}
function setVideoStatus(req){const auth=requireAdmin(req.sessionToken);const id=String(req.videoId||'');const active=!!req.active;const row=findRowIndexBy('VideoID','VideoID',id,String);if(row===-1)throw new Error('Video not found.');updateRowObject('Videos',row,{Active:active,UpdatedAt:nowIso()});logAdmin(auth.user.Email,'SET_VIDEO_STATUS','VIDEO',id,`Active=${active}`);return {success:true,message:active?'Video activated.':'Video deactivated.'};}

function saveReward(req){const auth=requireAdmin(req.sessionToken);const rewardId=String(req.rewardId||'').trim();const name=cleanText(req.name,180),description=cleanText(req.description,1000),mode=normalizeMode(req.rewardMode);const requiredDays=Math.max(0,Number(req.requiredDays||0)),requiredCoins=Math.max(0,Number(req.requiredCoins||0));const source=safeHttpUrl(req.sourceCodeUrl),download=safeHttpUrl(req.downloadUrl);const prompt=cleanTextPreserve(req.promptText,30000),license=cleanTextPreserve(req.licenseText,10000);const active=!!req.active;if(!name)throw new Error('Reward name is required.');if(mode==='BY_DAYS'&&requiredDays<1)throw new Error('Required days must be at least 1 for BY_DAYS.');if(mode==='BY_COINS'&&requiredCoins<1)throw new Error('Required coins must be at least 1 for BY_COINS.');if(mode==='BY_DAYS_AND_COINS'&&(requiredDays<1||requiredCoins<1))throw new Error('Both days and coins thresholds are required.');const now=nowIso();if(rewardId){const row=findRowIndexBy('RewardID','RewardID',rewardId,String);if(row===-1)throw new Error('Reward not found.');updateRowObject('Rewards',row,{Name:name,Description:description,RequiredDays:requiredDays,RequiredCoins:requiredCoins,RewardMode:mode,SourceCodeURL:source,DownloadURL:download,PromptText:prompt,LicenseText:license,Active:active,UpdatedAt:now});logAdmin(auth.user.Email,'UPDATE_REWARD','REWARD',rewardId,`Mode ${mode}; active ${active}`);return {success:true,message:'Reward updated.'};}const id='REWARD-'+Utilities.getUuid().slice(0,8);appendObject('Rewards',{RewardID:id,Name:name,Description:description,RequiredDays:requiredDays,RequiredCoins:requiredCoins,RewardMode:mode,SourceCodeURL:source,DownloadURL:download,PromptText:prompt,LicenseText:license,Active:active,CreatedAt:now,UpdatedAt:now});logAdmin(auth.user.Email,'CREATE_REWARD','REWARD',id,`Mode ${mode}; active ${active}`);return {success:true,message:'Reward added.'};}
function setRewardStatus(req){const auth=requireAdmin(req.sessionToken);const id=String(req.rewardId||'');const active=!!req.active;const row=findRowIndexBy('RewardID','RewardID',id,String);if(row===-1)throw new Error('Reward not found.');updateRowObject('Rewards',row,{Active:active,UpdatedAt:nowIso()});logAdmin(auth.user.Email,'SET_REWARD_STATUS','REWARD',id,`Active=${active}`);return {success:true,message:active?'Reward activated.':'Reward deactivated.'};}

function saveAdminSettings(req){const auth=requireAdmin(req.sessionToken);const incoming=req.settings||{};const allowed=['APP_NAME','ADMIN_EMAIL','FRONTEND_APP_URL','REQUIRED_WATCH_PERCENT','MIN_LEARNING_TEXT_LENGTH','DEFAULT_DAILY_COIN','REWARD_MODE','REQUIRED_DAYS','REQUIRED_COINS'];const lock=LockService.getScriptLock();lock.waitLock(30000);try{allowed.forEach(key=>{if(!(key in incoming))return;let value=cleanTextPreserve(incoming[key],1000).trim();if(key==='ADMIN_EMAIL'&&!isValidEmail(value))throw new Error('ADMIN_EMAIL must be a valid email.');if(key==='FRONTEND_APP_URL'&&value&& !/^https:\/\//i.test(value))throw new Error('FRONTEND_APP_URL must start with https://');if(key==='REQUIRED_WATCH_PERCENT'){const n=Number(value);if(!isFinite(n)||n<1||n>100)throw new Error('REQUIRED_WATCH_PERCENT must be 1–100.');value=String(n);}if(key==='MIN_LEARNING_TEXT_LENGTH'){const n=Number(value);if(!isFinite(n)||n<1||n>10000)throw new Error('MIN_LEARNING_TEXT_LENGTH is invalid.');value=String(Math.floor(n));}if(key==='DEFAULT_DAILY_COIN'){const n=Number(value);if(!isFinite(n)||n<0)throw new Error('DEFAULT_DAILY_COIN is invalid.');value=String(n);}if(key==='REWARD_MODE'){value=normalizeMode(value);}if(key==='REQUIRED_DAYS'||key==='REQUIRED_COINS'){const n=Number(value);if(!isFinite(n)||n<0)throw new Error(key+' is invalid.');value=String(Math.floor(n));}const row=findRowIndexBy('Settings','Key',key,String);if(row===-1)appendObject('Settings',{Key:key,Value:value,Description:settingDescription(key),UpdatedAt:nowIso()});else updateRowObject('Settings',row,{Value:value,Description:settingDescription(key),UpdatedAt:nowIso()});});
    const props=PropertiesService.getScriptProperties();if(incoming.ADMIN_EMAIL)props.setProperty('ADMIN_EMAIL',String(incoming.ADMIN_EMAIL).trim());if(incoming.FRONTEND_APP_URL)props.setProperty('FRONTEND_APP_URL',String(incoming.FRONTEND_APP_URL).trim());
    logAdmin(auth.user.Email,'SAVE_SETTINGS','SETTINGS','GLOBAL','Updated application settings.');return {success:true,message:'Settings saved.'};
  }finally{lock.releaseLock()}}
function logAdmin(adminEmail,action,targetType,targetId,details){appendObject('AdminLogs',{LogID:'ALOG-'+Utilities.getUuid(),AdminEmail:String(adminEmail||''),Action:String(action||''),TargetType:String(targetType||''),TargetID:String(targetId||''),Details:String(details||''),CreatedAt:nowIso()});}

/* =============================
   HELPERS
   ============================= */

function ensureDatabaseReady(){
  const ss=getSpreadsheet();const missing=SHEETS.filter(n=>!ss.getSheetByName(n));if(missing.length)throw new Error('Database is not ready. Run setupDatabase() once. Missing: '+missing.join(', '));
  // Repair headers without deleting data.
  ensureSheets(ss);
}

function getPublicSettings(){return {requiredWatchPercent:getSettingNumber('REQUIRED_WATCH_PERCENT',80),minLearningTextLength:getSettingNumber('MIN_LEARNING_TEXT_LENGTH',20),defaultDailyCoin:getSettingNumber('DEFAULT_DAILY_COIN',1),rewardMode:normalizeMode(getSetting('REWARD_MODE')||'BY_DAYS'),requiredDays:getSettingNumber('REQUIRED_DAYS',7),requiredCoins:getSettingNumber('REQUIRED_COINS',7),appName:String(getSetting('APP_NAME')||'Dowebcode Learn & Earn')}}
function getSetting(key){const row=findRowIndexBy('Settings','Key',key,String);if(row===-1)return DEFAULTS[key]||'';return String(rowObjectAt('Settings',row).Value??'');}
function getSettingNumber(key,fallback){const n=Number(getSetting(key));return isFinite(n)?n:fallback;}
function rowObjectAt(sheetName,rowIndex){const sheet=getSheet(sheetName),lastCol=sheet.getLastColumn();const headers=sheet.getRange(1,1,1,lastCol).getValues()[0].map(String),row=sheet.getRange(rowIndex,1,1,lastCol).getValues()[0];return rowToObject(headers,row);}
function normalizeEmail(email){return String(email||'').trim().toLowerCase();}
function isValidEmail(email){return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i.test(String(email||''));}
function cleanText(value,max){return cleanTextPreserve(value,max).replace(/[<>]/g,'');}
function cleanTextPreserve(value,max){let s=String(value??'');if(s.length>max)s=s.slice(0,max);return s;}
function truthy(v){return v===true||v===1||String(v).toLowerCase()==='true'||String(v).toLowerCase()==='yes';}
function nowIso(){return new Date().toISOString();}
function isoMinutesFromNow(minutes){return new Date(Date.now()+minutes*60000).toISOString();}
function dateKey(date){return Utilities.formatDate(date,Session.getScriptTimeZone()||'Asia/Karachi','yyyy-MM-dd');}
function generateOtp(){return String(Math.floor(100000+Math.random()*900000));}
function generateToken(){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,Utilities.getUuid()+Utilities.getUuid()+Math.random()+Date.now())).replace(/=/g,'')+Utilities.getUuid().replace(/-/g,'');}
function randomSalt(){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,Utilities.getUuid()+Date.now()+Math.random())).replace(/=/g,'');}
function hashOtp(value,salt){return bytesToHex(Utilities.computeHmacSha256Signature(String(value),String(salt)));}
function hashPassword(password,salt){let state=String(password)+'|'+String(salt);for(let i=0;i<PASSWORD_ITERATIONS;i++){state=bytesToHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,state));}return state;}
function verifyPassword(password,salt,hash){return constantTimeEqual(hashPassword(password,salt),hash);}
function constantTimeEqual(a,b){a=String(a);b=String(b);if(a.length!==b.length)return false;let diff=0;for(let i=0;i<a.length;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i);return diff===0;}
function bytesToHex(bytes){return bytes.map(b=>{const n=(b<0?b+256:b);return ('0'+n.toString(16)).slice(-2);}).join('');}
function generateMemberID(){
  const props=PropertiesService.getScriptProperties();const year=new Date().getFullYear();let next=Number(props.getProperty('MEMBER_SEQUENCE')||0);
  if(!next){const users=readRows('Users');users.forEach(u=>{const m=String(u.MemberID||'');const match=m.match(/^DWC-\d{4}-(\d{6})$/);if(match)next=Math.max(next,Number(match[1]));});}
  next+=1;props.setProperty('MEMBER_SEQUENCE',String(next));return 'DWC-'+year+'-'+String(next).padStart(6,'0');
}
function normalizeMode(mode){const m=String(mode||'BY_DAYS').toUpperCase();return ['BY_DAYS','BY_COINS','BY_DAYS_AND_COINS'].indexOf(m)!==-1?m:'BY_DAYS';}
function clamp(n,min,max){return Math.min(max,Math.max(min,n));}
function safeHttpUrl(url){const u=String(url||'').trim();if(!u)return '';if(!/^https:\/\//i.test(u))throw new Error('Reward links must use HTTPS.');return u;}
function safeErrorMessage(err){const m=err&&err.message?String(err.message):'Something went wrong. Please try again.';return m.length>280?m.slice(0,280)+'…':m;}

/* =============================
   EMAILS
   ============================= */

function sendVerificationEmail(to,name,otp){const subject='Verify your Dowebcode Learn & Earn account';const body=`Hello ${name||'Learner'},\n\nYour Dowebcode Learn & Earn verification code is:\n\n${otp}\n\nThis code expires in ${OTP_TTL_MINUTES} minutes and can only be used once.\n\nIf you did not create this account, you can ignore this email.\n\nDowebcode Learn & Earn`;MailApp.sendEmail({to,subject,body});}
function sendPasswordResetEmail(to,name,link){const subject='Reset your Dowebcode Learn & Earn password';const body=`Hello ${name||'Learner'},\n\nA password reset was requested for your Dowebcode Learn & Earn account.\n\nOpen this secure link to create a new password:\n${link}\n\nThis link expires in ${RESET_TTL_MINUTES} minutes and can be used once.\n\nIf you did not request this, you can ignore this email.\n\nDowebcode Learn & Earn`;MailApp.sendEmail({to,subject,body});}
