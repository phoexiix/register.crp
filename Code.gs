/** CRP Recruitment | Google Apps Script | GitHub Pages bridge */
const CRP = Object.freeze({spreadsheetId:'1fAI-KZzfAXJK-AWfUh52SyOM3S5SWUYiKpLt5F6RfMs', sheet:'Applications', sessions:'AdminSessions', hours:6});
const HEADERS=['ApplicationID','SubmittedAt','Prefix','FirstName','LastName','BirthDate','Phone','Email','Province','Education','Motivation','Status','AdminNote','UpdatedAt'];
const STATUSES=['รอตรวจสอบ','อยู่ระหว่างพิจารณา','ผ่านการคัดเลือก','ไม่ผ่านการคัดเลือก'];
function setup(){
 const p=PropertiesService.getScriptProperties();p.setProperty('SPREADSHEET_ID',CRP.spreadsheetId);
 const ss=SpreadsheetApp.openById(CRP.spreadsheetId);
 let a=ss.getSheetByName(CRP.sheet)||ss.insertSheet(CRP.sheet);
 if(!a.getLastRow()){a.appendRow(HEADERS);a.setFrozenRows(1);}else{
  const actual=a.getRange(1,1,1,HEADERS.length).getDisplayValues()[0];
  if(HEADERS.some((h,i)=>actual[i]!==h))throw Error('หัวตาราง Applications ไม่ตรงกับระบบ กรุณาตรวจสอบก่อนใช้งาน');
 }
 let s=ss.getSheetByName(CRP.sessions)||ss.insertSheet(CRP.sessions);
 if(!s.getLastRow()){s.appendRow(['TokenHash','ExpiresAt']);s.setFrozenRows(1);}
 Logger.log(ss.getUrl());
}
function db_(){return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')||CRP.spreadsheetId);}
function response_(ok,data){return ok?{ok:true,data}:{ok:false,error:String(data)};}
function json_(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);}
function doGet(e){
 const q=(e&&e.parameter)||{};
 if(q.bridge==='1'){
  const origin=String(q.origin||''),allowed=PropertiesService.getScriptProperties().getProperty('ALLOWED_ORIGIN')||'';
  if(!/^https:\/\/[^/]+$/.test(origin)||origin!==allowed)return HtmlService.createHtmlOutput('Origin not allowed');
  if(!/^[A-Za-z0-9_-]{12,100}$/.test(String(q.channel||'')))return HtmlService.createHtmlOutput('Invalid channel');
  const t=HtmlService.createTemplateFromFile('Bridge');t.origin=origin;t.channel=String(q.channel);
  return t.evaluate().setTitle('CRP Bridge').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
 }
 return json_(response_(true,{service:'CRP API',version:'2.0'}));
}
function doPost(e){try{return json_(dispatch_(JSON.parse(e.postData.contents||'{}')));}catch(err){return json_(response_(false,err.message));}}
function bridgeDispatch(p){return dispatch_(p||{});}
function dispatch_(p){try{
 const a=String(p.action||'');
 if(a==='publicList')return response_(true,publicList_(p.search));
 if(a==='login')return response_(true,login_(p.username,p.password));
 if(!authorized_(p.token))throw Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
 if(a==='logout'){logout_(p.token);return response_(true,{done:true});}
 if(a==='submit')return response_(true,submit_(p.data||{}));
 if(a==='list')return response_(true,list_(p.search));
 if(a==='update')return response_(true,update_(p.id,p.status,p.note));
 if(a==='stats')return response_(true,stats_());
 throw Error('Unknown action');
 }catch(err){return response_(false,err.message);}}
function clean_(v,max){return String(v==null?'':v).trim().slice(0,max).replace(/[\u0000-\u001f]/g,' ');}
function safe_(x){const s=String(x);return /^[=+\-@]/.test(s)?"'"+s:s;}
function hash_(x){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(x)).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');}
function login_(u,p){
 const cfg=PropertiesService.getScriptProperties(), user=cfg.getProperty('ADMIN_USER'),pass=cfg.getProperty('ADMIN_PASS');
 if(!user||!pass)throw Error('ตั้งค่า ADMIN_USER และ ADMIN_PASS ก่อน');
 if(String(u)!==user||String(p)!==pass)throw Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
 const token=Utilities.getUuid()+Utilities.getUuid(),expiry=Date.now()+CRP.hours*3600000;
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{db_().getSheetByName(CRP.sessions).appendRow([hash_(token),expiry]);}finally{lock.releaseLock();}
 return {token,expiresAt:expiry};
}
function authorized_(token){
 if(!token)return false;const sh=db_().getSheetByName(CRP.sessions),n=sh.getLastRow();if(n<2)return false;
 const h=hash_(token);return sh.getRange(2,1,n-1,2).getValues().some(r=>r[0]===h&&Number(r[1])>Date.now());
}
function logout_(token){if(!token)return;const sh=db_().getSheetByName(CRP.sessions),n=sh.getLastRow();if(n<2)return;const h=hash_(token),a=sh.getRange(2,1,n-1,1).getValues();for(let i=a.length-1;i>=0;i--)if(a[i][0]===h)sh.deleteRow(i+2);}
function rows_(){const sh=db_().getSheetByName(CRP.sheet),n=sh.getLastRow();return n<2?[]:sh.getRange(2,1,n-1,HEADERS.length).getDisplayValues().map(r=>Object.fromEntries(HEADERS.map((h,i)=>[h,r[i]])));}
function publicList_(search){
 const q=clean_(search,100).toLocaleLowerCase('th');
 // Public lookup never exposes contact details, birthdays, motivation, or internal notes.
 return rows_().filter(r=>!q||[r.ApplicationID,r.Prefix,r.FirstName,r.LastName].join(' ').toLocaleLowerCase('th').includes(q)).reverse().slice(0,500).map(r=>({ApplicationID:r.ApplicationID,Prefix:r.Prefix,FirstName:r.FirstName,LastName:r.LastName,Province:r.Province,Education:r.Education,Status:r.Status,SubmittedAt:r.SubmittedAt,UpdatedAt:r.UpdatedAt}));
}
function list_(search){const q=clean_(search,100).toLocaleLowerCase('th');return rows_().filter(r=>!q||[r.ApplicationID,r.FirstName,r.LastName,r.Phone,r.Status].join(' ').toLocaleLowerCase('th').includes(q)).reverse().slice(0,1000);}
function stats_(){const rows=rows_(),byStatus={};STATUSES.forEach(s=>byStatus[s]=0);rows.forEach(r=>byStatus[r.Status]=(byStatus[r.Status]||0)+1);return {total:rows.length,byStatus};}
function submit_(d){
 const f={prefix:clean_(d.prefix,20),firstName:clean_(d.firstName,100),lastName:clean_(d.lastName,100),birthDate:clean_(d.birthDate,20),phone:clean_(d.phone,20),email:clean_(d.email,150),province:clean_(d.province,100),education:clean_(d.education,150),motivation:clean_(d.motivation,2000)};
 if(!f.firstName||!f.lastName||!f.phone||!f.email)throw Error('กรุณากรอกชื่อ นามสกุล โทรศัพท์ และอีเมล');
 if(!/^\+?[0-9\- ]{9,16}$/.test(f.phone))throw Error('หมายเลขโทรศัพท์ไม่ถูกต้อง');
 if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(f.email))throw Error('อีเมลไม่ถูกต้อง');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const id='CRP-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase();
  const now=new Date().toISOString();
  db_().getSheetByName(CRP.sheet).appendRow([id,now,safe_(f.prefix),safe_(f.firstName),safe_(f.lastName),safe_(f.birthDate),safe_(f.phone),safe_(f.email),safe_(f.province),safe_(f.education),safe_(f.motivation),'รอตรวจสอบ','',now]);
  return {applicationId:id,status:'รอตรวจสอบ'};
 }finally{lock.releaseLock();}
}
function update_(id,status,note){
 if(!STATUSES.includes(status))throw Error('สถานะไม่ถูกต้อง');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{const sh=db_().getSheetByName(CRP.sheet),n=sh.getLastRow();if(n<2)throw Error('ไม่พบข้อมูล');const ids=sh.getRange(2,1,n-1,1).getDisplayValues(),i=ids.findIndex(r=>r[0]===id);if(i<0)throw Error('ไม่พบข้อมูล');sh.getRange(i+2,12,1,3).setValues([[status,safe_(clean_(note,2000)),new Date().toISOString()]]);return {id,status};}finally{lock.releaseLock();}
}
