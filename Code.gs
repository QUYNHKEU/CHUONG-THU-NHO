// Google Apps Script: tạo từ Extensions → Apps Script của Google Sheet.
const SHEET_NAME = 'Ket_qua';
function getSheet_(){const ss=SpreadsheetApp.getActiveSpreadsheet();let s=ss.getSheetByName(SHEET_NAME);if(!s){s=ss.insertSheet(SHEET_NAME);s.appendRow(['serverTime','classCode','group','attemptId','kind','correct','wrong','seq','at'])}return s}
function doPost(e){
  try{
    const p=e.parameter||{};
    const classCode=String(p.classCode||'').slice(0,80),group=String(p.group||'').slice(0,60),attemptId=String(p.attemptId||'').slice(0,90),kind=String(p.kind||'').slice(0,25);
    if(!classCode||!group||!attemptId)throw new Error('Thiếu thông tin');
    const correct=Math.min(4,Math.max(0,Number(p.correct)||0)),wrong=Math.max(0,Number(p.wrong)||0),seq=Math.max(0,Number(p.seq)||0);
    const lock=LockService.getScriptLock();lock.waitLock(25000);
    try{getSheet_().appendRow([new Date(),classCode,group,attemptId,kind,correct,wrong,seq,Number(p.at)||Date.now()])}finally{lock.releaseLock()}
    return ContentService.createTextOutput('ok');
  }catch(err){return ContentService.createTextOutput('error: '+err.message)}
}
function doGet(e){
 const p=e.parameter||{};const cb=String(p.callback||'');
 if(!/^gameCB[0-9]+$/.test(cb))return ContentService.createTextOutput('callback không hợp lệ');
 const code=String(p.classCode||'');const data=getSheet_().getDataRange().getValues();const latest=new Map();
 for(let i=1;i<data.length;i++){
   const r=data[i];if(String(r[1])!==code)continue;
   const key=String(r[2]).toLowerCase().trim();const seq=Number(r[7])||0;const at=Number(r[8])||0;
   const old=latest.get(key);
   // Chỉ hiện phiên chơi mới nhất của từng nhóm; sự kiện có thể tới lệch thứ tự.
   if(!old||at>old.updated||(at===old.updated&&seq>old.seq))latest.set(key,{group:String(r[2]),attemptId:String(r[3]),kind:String(r[4]),correct:Number(r[5])||0,wrong:Number(r[6])||0,seq,updated:at});
 }
 const payload=JSON.stringify({rows:Array.from(latest.values())});
 return ContentService.createTextOutput(cb+'('+payload+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
