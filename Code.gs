const SPREADSHEET_ID = '1TmHtDiDTYvOmRCb-esJi8Iqq_IrZfYgyHHlPD9eOyTQ';
const FIELD_IDS = ['tech0_0','tech0_1','tech0_2','tech1_0','tech1_1','tech1_2','tech2_0','tech2_1','tech2_2','tech3_0','tech3_1','tech3_2','company','respect','respectReason','gerty','aiRule','aiReason','hidden1','hidden2','integrity','justice0_0','justice0_1','justice0_2','justice0_3','justice1_0','justice1_1','justice1_2','justice1_3','stance','justiceReason','rule1','ruleReason1','rule2','ruleReason2','reflection'];
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return json_({ok:true,service:'moon-ethics',version:1}); }
function text_(v) { const s=String(v == null ? '' : v).trim(); if(s.length>10000)throw Error('답안 길이를 확인하세요.'); return /^[=+\-@]/.test(s) ? "'"+s : s; }
function hash_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join(''); }
function doPost(e) {
  const lock=LockService.getScriptLock();
  try {
    if(!e || !e.postData || e.postData.contents.length>420000)throw Error('요청 크기를 확인하세요.');
    const p=JSON.parse(e.postData.contents), a=p.answers || {};
    if(!/^[a-f0-9-]{36}$/.test(p.submissionId || '') || !/^[a-f0-9-]{72}$/.test(p.editToken || ''))throw Error('제출 확인값이 올바르지 않습니다.');
    if(!String(a.school||'').trim() || !String(a.name||'').trim() || !['1','2'].includes(String(a.grade)) || !/^[1-9]\d{0,3}$/.test(String(a.class)) || !/^[1-9]\d{0,3}$/.test(String(a.number)))throw Error('학교·학년·반·번호·이름을 확인하세요.');
    if(a.respect && !['존중받아야 한다','존중받을 필요가 없다','판단하기 어렵다'].includes(a.respect))throw Error('입장을 확인하세요.');
    if(a.stance && !['찬성','반대'].includes(a.stance))throw Error('찬반 선택을 확인하세요.');
    lock.waitLock(30000);
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), raw=ss.getSheetByName('제출 원본'), grades=ss.getSheetByName('채점표');
    if(!raw || !grades)throw Error('교사용 시트 설정이 필요합니다.');
    const count=raw.getLastRow(), found=count>1?raw.getRange(2,1,count-1,1).createTextFinder(p.submissionId).matchEntireCell(true).findNext():null;
    const row=found?found.getRow():count+1, tokenHash=hash_(p.editToken), now=new Date().toISOString();
    if(row>5000)throw Error('교사에게 제출 저장공간 확장을 요청하세요.');
    if(found && raw.getRange(row,2).getDisplayValue()!==tokenHash)throw Error('이 답안을 수정할 권한이 없습니다.');
    const first=found?raw.getRange(row,3).getDisplayValue():now;
    const classKey=`${a.school} · ${a.grade}학년 ${a.class}반`;
    const studentKey=`${a.name} · ${classKey} ${a.number}번 · #${p.submissionId.slice(0,8)}`;
    const values=[p.submissionId,tokenHash,first,now,a.school,String(a.grade),String(a.class),String(a.number),a.name,classKey,studentKey,...FIELD_IDS.map(id=>a[id]||'')].map(text_);
    raw.getRange(row,1,1,values.length).setValues([values]);
    const gcount=grades.getLastRow(), gfound=gcount>1?grades.getRange(2,1,gcount-1,1).createTextFinder(p.submissionId).matchEntireCell(true).findNext():null;
    const gr=gfound?gfound.getRow():gcount+1;
    grades.getRange(gr,1,1,2).setValues([[p.submissionId,text_(studentKey)]]);
    grades.getRange(gr,9).setFormula(`=IF(COUNT(C${gr}:H${gr})=0,"",SUM(C${gr}:H${gr}))`);
    grades.getRange(gr,11).setValue((found?'수정됨 · 재채점 확인 · ':'최초 접수 · ')+now);
    grades.getRange(gr,11).setBackground(found?'#fff2cc':'#ffffff');
    if(!found)raw.getRange(row,1,1,values.length).setWrap(true).setVerticalAlignment('top');
    SpreadsheetApp.flush();
    return json_({ok:true,receiptId:p.submissionId,updatedAt:now,revised:!!found});
  } catch(err) { return json_({ok:false,message:String(err.message||err)}); }
  finally { if(lock.hasLock())lock.releaseLock(); }
}
function onEdit(e) {
  if(!e || e.range.getSheet().getName()!=='교사용 학생 답안')return;
  const sheet=e.range.getSheet(), cell=e.range.getA1Notation();
  if(cell==='B2')sheet.getRange('B3').clearContent();
  if(cell==='B3'){SpreadsheetApp.flush();sheet.autoResizeRows(15,42);}
}