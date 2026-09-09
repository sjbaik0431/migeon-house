/**
 * 미건 임대 주택 현황 — Google Apps Script 백엔드
 * --------------------------------------------------
 * 이 파일 하나를 Google Sheets의 [확장 프로그램 → Apps Script]에 붙여넣고
 * setupSheets() → installTriggers() 순서로 실행하면 매월 손 안 대는 자동화가 시작됩니다.
 * 자세한 순서는 같은 폴더의 "사람이_직접_할_일.md"를 참고하세요.
 *
 * 이번 구축 범위(사용자 선택 기준):
 *  - 임차인 알림: 알림톡 없이 SMS만 (USE_ALIMTALK = false)
 *  - 은행 입금: 문자 자동수신 없이 "은행업로드" 시트에 엑셀을 붙여넣어 대사
 *  - Cloudflare Worker 게이트웨이 없음 → 대시보드가 Apps Script 웹앱 URL을 직접 호출
 *  - 배포처: GitHub Pages (주간 리포트도 같은 저장소에 커밋)
 */

/* ============================= 설정값 ============================= */
const TZ = Session.getScriptTimeZone() || "Asia/Seoul";
const USE_ALIMTALK = false; // 사업자 등록 후 알림톡으로 전환 시 true + 템플릿ID 등록

const SHEETS = {
  UNITS: "세대", INVOICES: "청구", DEPOSITS: "입금", BANK_UPLOAD: "은행업로드",
  REPAIRS: "수리", NOTICE: "알림로그"
};
const UNIT_HEADERS = ["호실","동","상태","임차인","전화","비상연락","입금자명","보증금","월세","관리비","납부일","최초입주일","계약시작","계약종료","확정일자","전월세신고"];
const INVOICE_HEADERS = ["청구월","호실","월세","관리비","합계","입금액","납부상태","최종입금일"];
const DEPOSIT_HEADERS = ["입금일","입금자명","금액","매칭호실","매칭상태","비고"];
const BANK_UPLOAD_HEADERS = ["거래일시","적요(입금자명)","입금액","처리완료"];
const REPAIR_HEADERS = ["접수일","호실","내용","상태","비용","업체"];
const NOTICE_HEADERS = ["발송일시","호실","종류","채널","내용"];

// 비밀값은 시트가 아닌 [프로젝트 설정 → 스크립트 속성]에 저장합니다 (사람이_직접_할_일.md 참고)
function prop_(key){ return PropertiesService.getScriptProperties().getProperty(key); }

/* ============================= STEP 1: 시트 세팅 ============================= */
function setupSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEETS.UNITS, UNIT_HEADERS);
  ensureSheet_(ss, SHEETS.INVOICES, INVOICE_HEADERS);
  ensureSheet_(ss, SHEETS.DEPOSITS, DEPOSIT_HEADERS);
  ensureSheet_(ss, SHEETS.BANK_UPLOAD, BANK_UPLOAD_HEADERS);
  ensureSheet_(ss, SHEETS.REPAIRS, REPAIR_HEADERS);
  ensureSheet_(ss, SHEETS.NOTICE, NOTICE_HEADERS);

  const unitSheet = ss.getSheetByName(SHEETS.UNITS);
  setDropdown_(unitSheet, 3, ["입주","공실","퇴거예정"]); // 상태 열
  setDropdown_(unitSheet, 16, ["신고완료","미신고","-"]); // 전월세신고 열
  const invSheet = ss.getSheetByName(SHEETS.INVOICES);
  setDropdown_(invSheet, 7, ["예정","완납","부분","미납"]); // 납부상태 열
  SpreadsheetApp.getUi().alert("시트 세팅 완료. 이제 installTriggers()를 실행하세요.");
}
function ensureSheet_(ss, name, headers){
  let sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); }
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight("bold");
  sh.setFrozenRows(1);
  if(name===SHEETS.BANK_UPLOAD) return sh; // 사용자가 직접 붙여넣는 시트, 폭 조정 생략
  for(let c=1;c<=headers.length;c++) sh.autoResizeColumn(c);
  return sh;
}
function setDropdown_(sheet, col, values){
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values,true).setAllowInvalid(false).build();
  sheet.getRange(2,col,Math.max(sheet.getMaxRows()-1,500),1).setDataValidation(rule);
}

/* ============================= 공통 유틸 ============================= */
function sheetRows_(name){
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values.filter(r=>r.join("")!=="").map(r=>{
    const o={}; headers.forEach((h,i)=>o[h]=r[i]); o._row=null; return o;
  });
}
function fmtDate_(d){ return Utilities.formatDate(new Date(d), TZ, "yyyy-MM-dd"); }
function monthKey_(d){ return Utilities.formatDate(new Date(d), TZ, "yyyy-MM"); }
function todayStr_(){ return fmtDate_(new Date()); }
function addDays_(dateStr,n){ const d=new Date(dateStr+"T00:00:00"); d.setDate(d.getDate()+n); return d; }
function dayDiff_(a,b){ return Math.round((new Date(b+"T00:00:00")-new Date(a+"T00:00:00"))/86400000); }

/* ============================= STEP 2: 매월 청구 생성 ============================= */
function generateMonthlyInvoices(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unitSheet = ss.getSheetByName(SHEETS.UNITS);
  const invSheet = ss.getSheetByName(SHEETS.INVOICES);
  const units = sheetRows_(SHEETS.UNITS).filter(u=>u["상태"]==="입주");
  const existing = sheetRows_(SHEETS.INVOICES);
  const ym = monthKey_(new Date());
  const rows=[];
  units.forEach(u=>{
    const dup = existing.some(inv=>inv["호실"]===u["호실"] && inv["청구월"]===ym);
    if(dup) return;
    const rent=Number(u["월세"])||0, mfee=Number(u["관리비"])||0;
    rows.push([ym, u["호실"], rent, mfee, rent+mfee, 0, "예정", ""]);
  });
  if(rows.length) invSheet.getRange(invSheet.getLastRow()+1,1,rows.length,INVOICE_HEADERS.length).setValues(rows);
  Logger.log(ym+" 청구 "+rows.length+"건 생성");
}

/* ============================= STEP 3: 은행 엑셀 업로드 대사 ============================= */
// 사용법: 은행 홈페이지에서 받은 거래내역 엑셀을 "은행업로드" 시트에 붙여넣고(거래일시/적요/입금액 순서만 맞추면 됨)
//        상단 메뉴 "미건 관리 → 은행 입금 대사 실행"을 누르면 자동 매칭됩니다.
function importBankExcel(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bankSheet = ss.getSheetByName(SHEETS.BANK_UPLOAD);
  const values = bankSheet.getDataRange().getValues();
  const headers = values.shift();
  const units = sheetRows_(SHEETS.UNITS).filter(u=>u["상태"]==="입주");
  const invSheet = ss.getSheetByName(SHEETS.INVOICES);
  const invoices = sheetRows_(SHEETS.INVOICES);
  const depositSheet = ss.getSheetByName(SHEETS.DEPOSITS);
  const ym = monthKey_(new Date());
  let matched=0, pending=0;
  const newDeposits=[];
  values.forEach((row,idx)=>{
    if(row[3]===true || row[3]==="완료") return; // 이미 처리됨
    const name=String(row[1]||"").trim(), amount=Number(row[2])||0;
    if(!name || !amount) return;
    const unit = units.find(u=>{
      const alias=String(u["입금자명"]||u["임차인"]||"");
      return alias && (name.indexOf(alias)>=0 || alias.indexOf(name)>=0);
    });
    let matchedUnit = "", status="확인필요";
    if(unit){
      const inv = invoices.find(i=>i["호실"]===unit["호실"] && i["청구월"]===ym && i["납부상태"]!=="완납");
      if(inv){
        applyPayment_(invSheet, unit["호실"], ym, amount);
        matchedUnit = unit["호실"]; status="자동"; matched++;
      } else { pending++; }
    } else { pending++; }
    newDeposits.push([fmtDate_(new Date()), name, amount, matchedUnit, status, ""]);
    bankSheet.getRange(idx+2,4).setValue(true);
  });
  if(newDeposits.length) depositSheet.getRange(depositSheet.getLastRow()+1,1,newDeposits.length,DEPOSIT_HEADERS.length).setValues(newDeposits);
  SpreadsheetApp.getUi().alert("대사 완료: 자동매칭 "+matched+"건, 확인필요 "+pending+"건");
}
function applyPayment_(invSheet, unitId, ym, amount){
  const data = invSheet.getDataRange().getValues();
  for(let r=1;r<data.length;r++){
    if(data[r][1]===unitId && data[r][0]===ym){
      const total=Number(data[r][4])||0;
      const paid=(Number(data[r][5])||0)+amount;
      const status = paid>=total ? "완납" : "부분";
      invSheet.getRange(r+1,6).setValue(paid);
      invSheet.getRange(r+1,7).setValue(status);
      invSheet.getRange(r+1,8).setValue(fmtDate_(new Date()));
      return;
    }
  }
}

/* ============================= STEP 4: 미납 판정 ============================= */
function markOverdue(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEETS.INVOICES);
  const units = sheetRows_(SHEETS.UNITS);
  const data = invSheet.getDataRange().getValues();
  const today = new Date();
  for(let r=1;r<data.length;r++){
    const status=data[r][6];
    if(status==="완납") continue;
    const unit = units.find(u=>u["호실"]===data[r][1]);
    if(!unit) continue;
    const dueDay = Number(unit["납부일"])||25;
    if(data[r][0]===monthKey_(today) && today.getDate()>=dueDay && status==="예정"){
      invSheet.getRange(r+1,7).setValue("미납");
    }
  }
}

/* ============================= STEP 5: 미납 안내 SMS ============================= */
function sendOverdueNotices(){
  const hour = new Date().getHours();
  if(hour<9 || hour>=20) return; // 야간 발송 금지
  const units = sheetRows_(SHEETS.UNITS);
  const invoices = sheetRows_(SHEETS.INVOICES).filter(i=>i["납부상태"]==="미납"||i["납부상태"]==="부분");
  const notices = sheetRows_(SHEETS.NOTICE);
  const today = todayStr_();
  invoices.forEach(inv=>{
    const unit = units.find(u=>u["호실"]===inv["호실"]); if(!unit || !unit["전화"]) return;
    const sentThisMonth = notices.filter(n=>n["호실"]===inv["호실"] && String(n["발송일시"]).slice(0,7)===monthKey_(new Date())).length;
    if(sentThisMonth>=2) return; // 월 2회 상한
    const alreadyToday = notices.some(n=>n["호실"]===inv["호실"] && String(n["발송일시"]).slice(0,10)===today);
    if(alreadyToday) return;
    const amount=(Number(inv["합계"])||0)-(Number(inv["입금액"])||0);
    const msg = "[미건] "+unit["임차인"]+"님, "+inv["청구월"]+" 임대료 "+amount.toLocaleString("ko-KR")+"원이 미납되었습니다. 확인 부탁드립니다.";
    sendSMS_(unit["전화"], msg);
    logNotice_(inv["호실"], "미납 안내", "SMS", msg);
  });
}
function sendSMS_(to, text){
  const apiKey=prop_("SOLAPI_API_KEY"), apiSecret=prop_("SOLAPI_API_SECRET"), sender=prop_("SMS_SENDER");
  if(!apiKey||!apiSecret||!sender){ Logger.log("Solapi 설정 없음 — 발송 스킵: "+text); return; }
  const date = new Date().toISOString();
  const salt = Utilities.getUuid();
  const signature = Utilities.computeHmacSha256Signature(date+salt, apiSecret)
    .map(b=>("0"+(b&0xFF).toString(16)).slice(-2)).join("");
  const payload = { message: { to: to.replace(/-/g,""), from: sender.replace(/-/g,""), text: text } };
  const options = {
    method:"post", contentType:"application/json",
    headers:{ Authorization: "HMAC-SHA256 apiKey="+apiKey+", date="+date+", salt="+salt+", signature="+signature },
    payload: JSON.stringify(payload), muteHttpExceptions:true
  };
  const res = UrlFetchApp.fetch("https://api.solapi.com/messages/v4/send", options);
  Logger.log(res.getContentText());
}
function logNotice_(unitId, type, channel, content){
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.NOTICE);
  sh.appendRow([Utilities.formatDate(new Date(),TZ,"yyyy-MM-dd HH:mm"), unitId, type, channel, content]);
}

/* ============================= STEP 6: 계약 만기 · 전월세신고 알림 ============================= */
function checkComplianceDates(){
  const units = sheetRows_(SHEETS.UNITS).filter(u=>u["상태"]!=="공실");
  const today = todayStr_();
  const thresholds=[180,120,90,60];
  const lines=[];
  units.forEach(u=>{
    if(u["계약종료"]){
      const end=fmtDate_(u["계약종료"]);
      const d=dayDiff_(today,end);
      if(thresholds.indexOf(d)>=0) lines.push(u["호실"]+"호 "+u["임차인"]+" 계약만기 D-"+d+" ("+end+")");
    }
    if(u["확정일자"] && (u["전월세신고"]==="미신고")){
      const deadline=fmtDate_(addDays_(fmtDate_(u["확정일자"]),30));
      const d=dayDiff_(today,deadline);
      if(d===30||d===7) lines.push(u["호실"]+"호 전월세신고 마감 D-"+d);
    }
  });
  if(lines.length) kakaoSendToMe_("[미건] 계약/신고 알림\n"+lines.join("\n"));
}

/* ============================= STEP 7: 대표 일일 카톡 보고 ============================= */
function sendDailyOwnerReport(){
  const state = buildStatePayload_();
  const occ = state.units.filter(u=>u.status!=="공실");
  const curMonth = state.units[0] ? state.units[0].months[state.units[0].months.length-1] : null;
  const paid = occ.filter(u=>curMonth && u.payments[curMonth]==="완납").length;
  const due = occ.filter(u=>curMonth && ["완납","미납","부분"].indexOf(u.payments[curMonth])>=0).length;
  const rate = due? Math.round(paid/due*100) : null;
  const overdue = state.units.filter(u=>u.arrears.months>0);
  const overdueAmt = overdue.reduce((a,u)=>a+u.arrears.amount,0);
  const vacant = state.units.filter(u=>u.status==="공실").length;
  const msg = "[미건 아침 보고 "+state.asOf+"]\n"
    +"수납률: "+(rate===null?"기록없음":rate+"%")+"\n"
    +"미납: "+overdue.length+"세대 · "+overdueAmt.toLocaleString("ko-KR")+"원\n"
    +"공실: "+vacant+"세대\n"
    +"확인필요 입금: "+state.pending.length+"건";
  const ok = kakaoSendToMe_(msg);
  if(!ok){
    const owner = prop_("OWNER_PHONE");
    if(owner) sendSMS_(owner, msg);
  }
}
function kakaoSendToMe_(text){
  const clientId=prop_("KAKAO_CLIENT_ID"), refreshToken=prop_("KAKAO_REFRESH_TOKEN");
  if(!clientId||!refreshToken) return false;
  try{
    const tokenRes = UrlFetchApp.fetch("https://kauth.kakao.com/oauth/token", {
      method:"post", payload:{grant_type:"refresh_token", client_id:clientId, refresh_token:refreshToken}, muteHttpExceptions:true
    });
    const tokenJson = JSON.parse(tokenRes.getContentText());
    if(!tokenJson.access_token) return false;
    if(tokenJson.refresh_token) PropertiesService.getScriptProperties().setProperty("KAKAO_REFRESH_TOKEN", tokenJson.refresh_token);
    const template = { object_type:"text", text:text, link:{ web_url:"https://"+(prop_("GITHUB_PAGES_HOST")||""), mobile_web_url:"https://"+(prop_("GITHUB_PAGES_HOST")||"") } };
    const sendRes = UrlFetchApp.fetch("https://kapi.kakao.com/v2/api/talk/memo/default/send", {
      method:"post", headers:{Authorization:"Bearer "+tokenJson.access_token},
      payload:{ template_object: JSON.stringify(template) }, muteHttpExceptions:true
    });
    return sendRes.getResponseCode()===200;
  }catch(e){ Logger.log(e); return false; }
}

/* ============================= STEP 8: 주간 리포트 → GitHub 커밋 ============================= */
function buildWeeklyReport(){
  const state = buildStatePayload_();
  const overdue = state.units.filter(u=>u.arrears.months>0).sort((a,b)=>b.arrears.amount-a.arrears.amount);
  const rows = overdue.map(u=>"<tr><td>"+u.building+" "+u.id+"호</td><td>"+u.tenant+"</td><td>"+u.arrears.months+"개월</td><td>"+u.arrears.amount.toLocaleString("ko-KR")+"원</td></tr>").join("");
  const html = "<!doctype html><meta charset='utf-8'><title>미건 주간 리포트 "+state.asOf+"</title>"
    +"<style>body{font-family:sans-serif;padding:24px;color:#1c2231}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:8px}</style>"
    +"<h1>미건 임대 주택 주간 리포트</h1><p>기준일 "+state.asOf+"</p>"
    +"<h2>미납 세대 ("+overdue.length+")</h2><table><tr><th>호실</th><th>임차인</th><th>연체개월</th><th>금액</th></tr>"+rows+"</table>";
  const path = "reports/weekly/"+Utilities.formatDate(new Date(),TZ,"yyyy-'W'ww")+".html";
  const url = commitToGitHub_(path, html, "주간 리포트 "+state.asOf);
  if(url) kakaoSendToMe_("[미건] 주간 리포트가 준비됐습니다\n"+url);
}
function commitToGitHub_(path, content, message){
  const token=prop_("GITHUB_TOKEN"), repo=prop_("GITHUB_REPO");
  if(!token||!repo) { Logger.log("GITHUB_TOKEN/GITHUB_REPO 설정 필요 — 커밋 스킵"); return null; }
  const base = "https://api.github.com/repos/"+repo+"/contents/"+path;
  const headers = { Authorization:"token "+token, "User-Agent":"migeon-house-script" };
  let sha=null;
  const getRes = UrlFetchApp.fetch(base, {headers:headers, muteHttpExceptions:true});
  if(getRes.getResponseCode()===200) sha = JSON.parse(getRes.getContentText()).sha;
  const payload = { message:message, content:Utilities.base64Encode(content, Utilities.Charset.UTF_8) };
  if(sha) payload.sha = sha;
  const putRes = UrlFetchApp.fetch(base, {method:"put", headers:headers, contentType:"application/json", payload:JSON.stringify(payload), muteHttpExceptions:true});
  if(putRes.getResponseCode()>=200 && putRes.getResponseCode()<300){
    const host = prop_("GITHUB_PAGES_HOST"); // 예: sjbaik0431.github.io/migeon-house
    return host ? "https://"+host+"/"+path : null;
  }
  Logger.log(putRes.getContentText()); return null;
}

/* ============================= STEP 9: 대시보드 API (doGet) ============================= */
function doGet(e){
  const token = e && e.parameter && e.parameter.token;
  if(token !== prop_("WEBAPP_TOKEN")){
    return ContentService.createTextOutput(JSON.stringify({error:"unauthorized"})).setMimeType(ContentService.MimeType.JSON);
  }
  const payload = buildStatePayload_();
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
function buildMonthsList_(asOf){
  const m=[]; for(let i=11;i>=0;i--){ const d=new Date(asOf+"T00:00:00"); d.setMonth(d.getMonth()-i); m.push(monthKey_(d)); } return m;
}
function buildStatePayload_(){
  const asOf = todayStr_();
  const months = buildMonthsList_(asOf);
  const unitRows = sheetRows_(SHEETS.UNITS);
  const invoiceRows = sheetRows_(SHEETS.INVOICES);
  const units = unitRows.map(u=>{
    const payments={};
    months.forEach(m=>{
      if(u["상태"]==="공실"){ payments[m]="해당없음"; return; }
      if(u["최초입주일"] && new Date(m+"-01")<new Date(monthKey_(u["최초입주일"])+"-01")){ payments[m]="해당없음"; return; }
      const inv = invoiceRows.find(i=>i["호실"]===u["호실"] && i["청구월"]===m);
      payments[m] = inv ? inv["납부상태"] : "미기록";
    });
    let arrMonths=0, arrAmount=0, firstUnpaid=null;
    for(let k=months.length-1;k>=0;k--){
      const st=payments[months[k]];
      if(st==="미납"||st==="부분"){
        arrMonths++; const inv=invoiceRows.find(i=>i["호실"]===u["호실"]&&i["청구월"]===months[k]);
        arrAmount += inv ? (Number(inv["합계"])-Number(inv["입금액"]||0)) : 0;
        firstUnpaid = months[k];
      } else break;
    }
    const overdueDays = firstUnpaid ? Math.max(0,dayDiff_(firstUnpaid+"-"+String(u["납부일"]||25).padStart(2,"0"), asOf)) : 0;
    return {
      id:String(u["호실"]), building:u["동"], floor:parseInt(u["호실"],10)||0,
      tenant:u["임차인"]||"-", phone:u["전화"]||"-", emergency:u["비상연락"]||"-", depositorAlias:u["입금자명"]||u["임차인"]||"-",
      deposit:Number(u["보증금"])||0, rent:Number(u["월세"])||0, maintenanceFee:Number(u["관리비"])||0,
      dueDay:Number(u["납부일"])||25, status:u["상태"],
      moveInDate:u["최초입주일"]?fmtDate_(u["최초입주일"]):null, renewals:0,
      contractStart:u["계약시작"]?fmtDate_(u["계약시작"]):null, contractEnd:u["계약종료"]?fmtDate_(u["계약종료"]):null,
      confirmedDate:u["확정일자"]?fmtDate_(u["확정일자"]):null, reportStatus:u["전월세신고"]||"-",
      payments:payments, months:months, arrears:{months:arrMonths, amount:arrAmount, overdueDays:overdueDays}
    };
  });
  const repairs = sheetRows_(SHEETS.REPAIRS).map((r,i)=>({id:"R"+i, unit:String(r["호실"]), building:(units.find(u=>u.id===String(r["호실"]))||{}).building||"", tenant:(units.find(u=>u.id===String(r["호실"]))||{}).tenant||"", date:r["접수일"]?fmtDate_(r["접수일"]):"", content:r["내용"], status:r["상태"], cost:Number(r["비용"])||0, vendor:r["업체"]}));
  const pending = sheetRows_(SHEETS.DEPOSITS).filter(d=>d["매칭상태"]==="확인필요").map((d,i)=>({id:"P"+i, amount:Number(d["금액"])||0, rawName:d["입금자명"], date:d["입금일"]?fmtDate_(d["입금일"]):"", note:"입금자명 불일치 — 세대 확인 필요"}));
  const noticeLog = sheetRows_(SHEETS.NOTICE).slice(-30).reverse().map(n=>({date:String(n["발송일시"]).slice(0,10), unit:String(n["호실"]), tenant:(units.find(u=>u.id===String(n["호실"]))||{}).tenant||"", type:n["종류"], channel:n["채널"]}));
  return {asOf:asOf, units:units, repairs:repairs, pending:pending, noticeLog:noticeLog, isSample:false, readOnly:true};
}

/* ============================= STEP 10: 트리거 설치 ============================= */
function installTriggers(){
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("generateMonthlyInvoices").timeBased().onMonthDay(1).atHour(0).nearMinute(5).create();
  ScriptApp.newTrigger("markOverdue").timeBased().everyDays(1).atHour(0).nearMinute(10).create();
  ScriptApp.newTrigger("sendOverdueNotices").timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger("sendDailyOwnerReport").timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger("checkComplianceDates").timeBased().everyDays(1).atHour(8).nearMinute(30).create();
  ScriptApp.newTrigger("buildWeeklyReport").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).nearMinute(30).create();
  SpreadsheetApp.getUi().alert("트리거 6개 설치 완료. 은행 입금 대사는 상단 메뉴에서 수동 실행하세요.");
}

/* ============================= 스프레드시트 메뉴 ============================= */
function onOpen(){
  SpreadsheetApp.getUi().createMenu("미건 관리")
    .addItem("① 시트 세팅 (최초 1회)", "setupSheets")
    .addItem("② 트리거 설치 (최초 1회)", "installTriggers")
    .addSeparator()
    .addItem("이번 달 청구 생성", "generateMonthlyInvoices")
    .addItem("은행 입금 대사 실행", "importBankExcel")
    .addItem("미납 판정 다시 실행", "markOverdue")
    .addItem("미납 안내 문자 지금 보내기", "sendOverdueNotices")
    .addItem("대표 보고 지금 보내기", "sendDailyOwnerReport")
    .addItem("주간 리포트 지금 만들기", "buildWeeklyReport")
    .addToUi();
}
