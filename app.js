(()=>{
const $=id=>document.getElementById(id);
const API=window.OLIVE_API_URL||'';
const SKEY='olive_att_session_v2',EKEY='olive_att_empid';
let S=null,month=null,amonth=null,st=null,shift='日勤',decReq=null,editRec=null;
const store={get(k){try{return localStorage.getItem(k)}catch(e){return null}},set(k,v){try{localStorage.setItem(k,v)}catch(e){}},del(k){try{localStorage.removeItem(k)}catch(e){}}};

const jst=()=>new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Tokyo',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false}).format(new Date());
const hm=s=>s?String(s).slice(5,16).replace('-','/'):'未打刻';
const dur=m=>m===''||m==null?'':Math.floor(m/60)+'時間'+String(m%60).padStart(2,'0')+'分';
const shiftMonth=(m,d)=>{let[y,mo]=m.split('-').map(Number);mo+=d;if(mo<1){mo=12;y--}if(mo>12){mo=1;y++}return y+'-'+String(mo).padStart(2,'0')};
function el(tag,props,...kids){const e=document.createElement(tag);Object.assign(e,props||{});kids.flat().forEach(k=>e.append(k));return e}
function say(id,text,cls){const n=$(id);n.textContent=text||'';n.className='notice'+(cls?' '+cls:'')}
function show(view){['login','main'].forEach(v=>$(v).classList.toggle('hidden',v!==view))}

async function api(action,data){
  if(!API)throw new Error('config.js にAPIのURLが設定されていません');
  let res;
  try{res=await fetch(API,{method:'POST',body:JSON.stringify({action,token:S&&S.token,...data})})}
  catch(e){throw new Error('通信できません。電波の良い場所でもう一度お試しください')}
  let j;try{j=await res.json()}catch(e){throw new Error('サーバーの応答が不正です')}
  if(!j.ok){if(j.code==='AUTH'){logout(j.error)}const err=new Error(j.error);err.handled=j.code==='AUTH';throw err}
  return j;
}
function busy(btns,on){btns.forEach(b=>b.disabled=on)}

/* ---- ログイン ---- */
function logout(msg){S=null;store.del(SKEY);show('login');$('pin').value='';$('empId').value=store.get(EKEY)||'';say('loginMsg',msg||'','err')}
$('loginBtn').onclick=async()=>{
  const b=$('loginBtn');busy([b],true);say('loginMsg','');
  try{
    const empId=$('empId').value.trim();
    const j=await api('login',{empId,pin:$('pin').value});
    S={token:j.token,user:j.user,conf:j.conf};store.set(SKEY,JSON.stringify(S));store.set(EKEY,empId);$('pin').value='';
    start();
  }catch(e){if(!e.handled)say('loginMsg',e.message,'err')}finally{busy([b],false)}
};
$('pin').addEventListener('keydown',e=>{if(e.key==='Enter')$('loginBtn').click()});
$('logout').onclick=()=>{$('menuDlg').dataset.force='';$('menuDlg').close();logout()};

function start(){
  show('main');$('who').textContent=S.user.name;
  $('tabs').classList.toggle('hidden',!S.user.admin);
  $('site').replaceChildren(...S.conf.sites.map(s=>el('option',{value:s,textContent:s})));
  const last=store.get('olive_att_site_'+S.user.id);if(last&&S.conf.sites.includes(last))$('site').value=last;
  $('rqShift').replaceChildren(...S.conf.shifts.map(s=>el('option',{value:s,textContent:s})));
  $('shiftSeg').replaceChildren(...S.conf.shifts.map(s=>el('button',{type:'button',textContent:s,onclick:()=>{shift=s;paintShift()}})));
  paintShift();month=jst().slice(0,7);amonth=month;showTab('punch');
  if(S.user.mustChange)openMenu(true);
}
function paintShift(){[...$('shiftSeg').children].forEach(b=>b.classList.toggle('on',b.textContent===shift))}

/* ---- 打刻 ---- */
async function refresh(){
  try{
    st=await api('status',{month});
    const o=st.open;
    $('state').textContent=o?'勤務中':'勤務前';
    $('inForm').classList.toggle('hidden',!!o);
    $('inBtn').disabled=!!o;$('outBtn').disabled=!o;
    if(o)say('msg',`${o.shift}　${hm(o.inAt)} 出勤　${o.site}`);
    $('monthLabel').textContent=month.replace('-','年')+'月';
    const h=$('history');h.replaceChildren();
    if(!st.records.length)h.textContent='この月の記録はありません';
    for(const r of st.records){
      const tagCls=r.state==='打刻漏れ'?' warn':'';
      const d=el('div',{className:'history'});
      d.append(el('b',{textContent:`${String(r.inAt).slice(5,10).replace('-','/')}　${r.shift}`}),el('span',{className:'tag'+tagCls,textContent:r.state}));
      d.append(document.createTextNode(`\n出勤：${hm(r.inAt)}　退勤：${hm(r.outAt)}\n勤務先：${r.site||'-'}`+(r.outAt?`　休憩：${r.brk}分　実労働：${dur(r.work)}`:'')));
      d.append(el('div',{},el('button',{className:'secondary small',textContent:'修正を申請',onclick:()=>openReq(r)})));
      h.append(d);
    }
    const q=$('myReqs');q.replaceChildren();
    if(!st.requests.length)q.textContent='申請はありません';
    for(const r of st.requests){
      const cls=r.state==='承認'?'':r.state==='却下'?' ng':' warn';
      const d=el('div',{className:'history'});
      d.append(el('b',{textContent:`${hm(r.newIn)} → ${hm(r.newOut)}`}),el('span',{className:'tag'+cls,textContent:r.state}));
      d.append(document.createTextNode(`\n理由：${r.reason}`+(r.comment?`\nコメント：${r.comment}`:'')));
      q.append(d);
    }
  }catch(e){if(!e.handled)say('msg',e.message,'err')}
}
$('inBtn').onclick=()=>doPunch('in');$('outBtn').onclick=()=>doPunch('out');
async function doPunch(type){
  const label=type==='in'?'出勤':'退勤';
  const extra=type==='in'?`\n${shift}／${$('site').value}`:'';
  if(!confirm(label+'を記録しますか？'+extra))return;
  busy([$('inBtn'),$('outBtn')],true);say('msg','送信中…');
  try{
    if(type==='in')store.set('olive_att_site_'+S.user.id,$('site').value);
    const j=await api('punch',{type,shift,site:$('site').value});
    await refresh();say('msg',`${label}を記録しました（${j.at.slice(11,16)}）`,'ok');
  }catch(e){await refresh();if(!e.handled)say('msg',e.message,'err')}
}
$('prev').onclick=()=>{month=shiftMonth(month,-1);refresh()};
$('next').onclick=()=>{month=shiftMonth(month,1);refresh()};

/* ---- 修正申請 ---- */
const toLocal=s=>s?String(s).slice(0,16).replace(' ','T'):'';
function openReq(r){
  editRec=r||null;say('rqMsg','');
  $('reqTarget').textContent=r?`対象：${hm(r.inAt)}　${r.shift}`:'打刻をし忘れた日の勤務を追加します';
  $('rqShift').value=r?r.shift:shift;$('rqIn').value=toLocal(r&&r.inAt);$('rqOut').value=toLocal(r&&r.outAt);$('rqReason').value='';
  $('reqDlg').showModal();
}
$('addReq').onclick=()=>openReq(null);
$('rqCancel').onclick=()=>$('reqDlg').close();
$('reqForm').onsubmit=async ev=>{
  ev.preventDefault();say('rqMsg','送信中…');
  try{
    await api('requestCorrection',{recordId:editRec&&editRec.id,newIn:$('rqIn').value,newOut:$('rqOut').value,shift:$('rqShift').value,reason:$('rqReason').value});
    $('reqDlg').close();await refresh();say('msg','修正を申請しました。岩久保の承認をお待ちください','ok');
  }catch(e){if(!e.handled)say('rqMsg',e.message,'err')}
};

/* ---- 設定(PIN) ---- */
function openMenu(force){
  say('pinMsg',force?'初回ログインです。新しいPINを設定してください':'');
  $('menuClose').classList.toggle('hidden',!!force);$('oldPin').value='';$('newPin').value='';
  $('menuDlg').dataset.force=force?'1':'';
  $('menuDlg').showModal();
}
$('menuBtn').onclick=()=>openMenu(false);$('menuClose').onclick=()=>$('menuDlg').close();
$('menuDlg').addEventListener('cancel',e=>{if($('menuDlg').dataset.force)e.preventDefault()});
$('pinBtn').onclick=async()=>{
  try{
    const j=await api('changePin',{oldPin:$('oldPin').value,newPin:$('newPin').value});
    S.token=j.token;S.user.mustChange=false;store.set(SKEY,JSON.stringify(S));
    $('menuDlg').dataset.force='';$('menuDlg').close();say('msg','PINを変更しました','ok');
  }catch(e){if(!e.handled)say('pinMsg',e.message,'err')}
};

/* ---- 管理者 ---- */
function showTab(t){
  $('punchView').classList.toggle('hidden',t!=='punch');$('adminView').classList.toggle('hidden',t!=='admin');
  $('tabPunch').classList.toggle('on',t==='punch');$('tabAdmin').classList.toggle('on',t==='admin');
  (t==='punch'?refresh:adminRefresh)();
}
$('tabPunch').onclick=()=>showTab('punch');$('tabAdmin').onclick=()=>showTab('admin');
let adminData=null;
async function adminRefresh(){
  if(!S.user.admin)return;
  try{
    const [rq,mo]=await Promise.all([api('adminRequests',{state:'申請中'}),api('adminMonth',{month:amonth})]);
    adminData=mo;
    const p=$('pending');p.replaceChildren();
    if(!rq.requests.length)p.textContent='承認待ちはありません';
    for(const r of rq.requests){
      const d=el('div',{className:'history'});
      d.append(el('b',{textContent:`${r.name}　${r.newShift}`}),document.createTextNode(`\n${r.recordId?`修正前：${hm(r.oldIn)} → ${hm(r.oldOut)}\n`:'新規追加\n'}修正後：${hm(r.newIn)} → ${hm(r.newOut)}\n理由：${r.reason}`));
      d.append(el('div',{},el('button',{className:'secondary small',textContent:'判断する',onclick:()=>openDec(r)})));
      p.append(d);
    }
    $('amonthLabel').textContent=amonth.replace('-','年')+'月';
    const by={};
    for(const r of mo.records){const k=r.empId;(by[k]=by[k]||{name:r.name,days:0,work:0,open:0});by[k].days++;by[k].work+=Number(r.work)||0;if(r.state==='勤務中'||r.state==='打刻漏れ')by[k].open++}
    const tb=el('table');tb.append(el('tr',{},el('th',{textContent:'氏名'}),el('th',{textContent:'回数'}),el('th',{textContent:'実労働'}),el('th',{textContent:'要確認'})));
    Object.values(by).forEach(v=>tb.append(el('tr',{},el('td',{textContent:v.name}),el('td',{textContent:v.days}),el('td',{textContent:dur(v.work)}),el('td',{textContent:v.open?v.open+'件':''}))));
    $('summary').replaceChildren(Object.keys(by).length?tb:'この月の記録はありません');
    $('resetSel').replaceChildren(...mo.employees.filter(e=>e.active).map(e=>el('option',{value:e.id,textContent:`${e.id} ${e.name}${e.pinSet?'':'（PIN未発行）'}`})));
  }catch(e){if(!e.handled)say('amsg',e.message,'err')}
}
$('aprev').onclick=()=>{amonth=shiftMonth(amonth,-1);adminRefresh()};$('anext').onclick=()=>{amonth=shiftMonth(amonth,1);adminRefresh()};
function openDec(r){
  decReq=r;say('decMsg','');$('decComment').value='';
  $('decBody').textContent=`${r.name}　${r.newShift}\n${r.recordId?`修正前：${hm(r.oldIn)} → ${hm(r.oldOut)}\n`:'新規追加\n'}修正後：${hm(r.newIn)} → ${hm(r.newOut)}\n理由：${r.reason}`;
  $('decDlg').showModal();
}
$('decCancel').onclick=()=>$('decDlg').close();
const decide=async decision=>{
  try{
    await api('adminDecide',{requestId:decReq.id,decision,comment:$('decComment').value});
    $('decDlg').close();say('amsg',decision+'しました','ok');adminRefresh();
  }catch(e){if(!e.handled)say('decMsg',e.message,'err')}
};
$('decOk').onclick=()=>decide('承認');$('decNg').onclick=()=>decide('却下');
$('resetBtn').onclick=async()=>{
  const sel=$('resetSel'),name=sel.selectedOptions[0]&&sel.selectedOptions[0].textContent;
  if(!name||!confirm(name+' のPINを再発行します。よろしいですか？'))return;
  try{const j=await api('adminResetPin',{empId:sel.value});alert(`${j.name} さんの新しいPIN：${j.pin}\n（この画面を閉じると再表示できません。本人に直接伝えてください）`);adminRefresh()}
  catch(e){if(!e.handled)say('amsg',e.message,'err')}
};
$('csv').onclick=()=>{
  if(!adminData)return;
  const rows=[['従業員ID','氏名','勤務区分','勤務先','出勤','退勤','休憩分','実労働分','状態'],...adminData.records.map(r=>[r.empId,r.name,r.shift,r.site,r.inAt,r.outAt,r.brk,r.work,r.state])];
  const csv='﻿'+rows.map(r=>r.map(v=>'"'+String(v??'').replaceAll('"','""')+'"').join(',')).join('\r\n');
  const a=el('a',{href:URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'})),download:`打刻_${amonth}.csv`});a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
};

/* ---- 起動 ---- */
setInterval(()=>{$('clock').textContent=jst()},1000);$('clock').textContent=jst();
if(!API)$('config-warn').hidden=false;
try{S=JSON.parse(store.get(SKEY)||'null')}catch(e){S=null}
if(S&&S.token)start();else{show('login');$('empId').value=store.get(EKEY)||''}
if('serviceWorker' in navigator)navigator.serviceWorker.register('./sw.js').catch(()=>{});
})();
