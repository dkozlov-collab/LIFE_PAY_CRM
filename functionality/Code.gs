/** LIFE PAY CRM — backend v7.0 */
var SS = SpreadsheetApp.getActiveSpreadsheet();
var SHEETS = {
  req:'Заявки', org:'Организации', exe:'Исполнители',
  usr:'Пользователи', log:'Журнал', set:'Настройки', ses:'Сессии'
};

/* ─── точки входа ─── */
function doGet(){ return out({ok:true, ping:true, version:'CRM 7.0'}); }
function doPost(e){
  var r={};
  try{ r=JSON.parse(e.postData.contents); }catch(err){ return out({ok:false,error:'Битый JSON'}); }
  try{ return out({ok:true, result: dispatch(r.action, r.payload||{}, r.token||'')}); }
  catch(err){ return out({ok:false, error:String(err && err.message || err)}); }
}
function out(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ─── таблицы ─── */
function sh(name){
  var s=SS.getSheetByName(name);
  if(!s){ s=SS.insertSheet(name); s.appendRow(HEADERS[name]); }
  return s;
}
var HEADERS={
 'Заявки':['id','no','ts','employee','priority','status','executorId','orgId','kktCount','rnm','zn','fn','model','needTrip','tripAddress','esp','kep','wifi','comment','updatedAt'],
 'Организации':['id','name','inn','contact','phone','email','active','createdAt'],
 'Исполнители':['id','name','phone','email','active'],
 'Пользователи':['id','login','passHash','name','role','orgId','execId','active'],
 'Журнал':['ts','login','action','entity','entityId','details'],
 'Настройки':['key','value'],
 'Сессии':['token','login','exp']
};
function readAll(name){
  var s=sh(name), v=s.getDataRange().getValues();
  if(v.length<2) return [];
  var h=v[0], res=[];
  for(var i=1;i<v.length;i++){ var o={}; for(var j=0;j<h.length;j++) o[h[j]]=v[i][j]; o.__row=i+1; res.push(o); }
  return res;
}
function writeObj(name,obj){
  var s=sh(name), h=HEADERS[name];
  s.appendRow(h.map(function(k){ return obj[k]!==undefined?obj[k]:''; }));
  return obj;
}
function updateObj(name,obj){
  var s=sh(name), h=HEADERS[name];
  s.getRange(obj.__row,1,1,h.length).setValues([h.map(function(k){ return obj[k]!==undefined?obj[k]:''; })]);
  return obj;
}
function removeRow(name,row){ sh(name).deleteRow(row); }
function uid(){ return Utilities.getUuid().slice(0,8); }
function sha(s){
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s))
    .map(function(b){ return ('0'+(b&255).toString(16)).slice(-2); }).join('');
}

/* ─── настройки / SLA ─── */
function settings(){
  var o={slaHours:60,notifyOrg:true,notifyExec:true,mailFrom:'',version:'CRM 7.0'};
  readAll(SHEETS.set).forEach(function(r){
    var v=r.value; if(v==='true')v=true; if(v==='false')v=false;
    o[r.key]=v;
  });
  o.slaHours=Number(o.slaHours)||60;
  return o;
}
function isOverdue(r,sla){
  if(r.status==='Выполнено'||r.status==='Отменено') return false;
  return (Date.now()-new Date(r.ts).getTime())/36e5 > sla;
}

/* ─── авторизация ─── */
function session(token){
  if(!token) throw new Error('Не авторизовано');
  var s=readAll(SHEETS.ses).filter(function(x){ return x.token===token; })[0];
  if(!s || Number(s.exp)<Date.now()) throw new Error('Сессия истекла');
  var u=readAll(SHEETS.usr).filter(function(x){ return x.login===s.login; })[0];
  if(!u || u.active===false || u.active==='FALSE') throw new Error('Учётная запись отключена');
  return u;
}
function needAdmin(u){ if(u.role!=='ADMIN') throw new Error('Недостаточно прав'); }
function log(login,action,entity,entityId,details){
  sh(SHEETS.log).appendRow([new Date(),login,action,entity,entityId||'',details||'']);
}

/* ─── диспетчер ─── */
function dispatch(action,p,token){
  if(action==='ping')  return {ok:true,version:'CRM 7.0'};
  if(action==='login') return login(p.login,p.pass);

  var me=session(token);

  switch(action){
    case 'logout':        return logout(token,me);
    case 'bootstrap':     return bootstrap(me);
    case 'saveRequest':   return saveRequest(me,p);
    case 'setStatus':     return setStatus(me,p);
    case 'transfer':      needAdmin(me); return transfer(me,p);
    case 'deleteRequest': needAdmin(me); return delEntity(me,SHEETS.req,p.id,'request');
    case 'saveOrg':       needAdmin(me); return saveSimple(me,SHEETS.org,p,'org',['name','inn','contact','phone','email','active']);
    case 'toggleOrg':     needAdmin(me); return toggle(me,SHEETS.org,p.id,'org');
    case 'deleteOrg':     needAdmin(me);
      if(readAll(SHEETS.req).some(function(r){return r.orgId===p.id;})) throw new Error('Есть связанные заявки — отключите организацию');
      return delEntity(me,SHEETS.org,p.id,'org');
    case 'saveExec':      needAdmin(me); return saveSimple(me,SHEETS.exe,p,'executor',['name','phone','email','active']);
    case 'toggleExec':    needAdmin(me); return toggle(me,SHEETS.exe,p.id,'executor');
    case 'deleteExec':    needAdmin(me);
      if(readAll(SHEETS.req).some(function(r){return r.executorId===p.id;})) throw new Error('Есть назначенные заявки');
      return delEntity(me,SHEETS.exe,p.id,'executor');
    case 'saveUser':      needAdmin(me); return saveUser(me,p);
    case 'toggleUser':    needAdmin(me);
      if(p.id===me.id) throw new Error('Нельзя отключить самого себя');
      return toggle(me,SHEETS.usr,p.id,'user');
    case 'deleteUser':    needAdmin(me);
      if(p.id===me.id) throw new Error('Нельзя удалить самого себя');
      return delEntity(me,SHEETS.usr,p.id,'user');
    case 'saveSettings':  needAdmin(me); return saveSettings(me,p);
    case 'testMail':      needAdmin(me);
      MailApp.sendEmail(p.to,'LIFE PAY CRM — тест','Проверка отправки почты из CRM.');
      log(me.login,'test','mail','',p.to); return {ok:true,message:'Письмо отправлено на '+p.to};
    default: throw new Error('Неизвестное действие: '+action);
  }
}

/* ─── реализация ─── */
function login(lg,pw){
  var u=readAll(SHEETS.usr).filter(function(x){ return String(x.login)===String(lg); })[0];
  if(!u) throw new Error('Неверный логин или пароль');
  if(u.active===false||u.active==='FALSE') throw new Error('Учётная запись отключена');
  if(String(u.passHash)!==sha(pw) && String(u.passHash)!==String(pw)) throw new Error('Неверный логин или пароль');
  var t=Utilities.getUuid();
  sh(SHEETS.ses).appendRow([t,u.login,Date.now()+12*36e5]);
  log(u.login,'login','user',u.id,'');
  return {token:t,user:pubUser(u)};
}
function logout(token,me){
  readAll(SHEETS.ses).filter(function(s){return s.token===token;})
    .forEach(function(s){ removeRow(SHEETS.ses,s.__row); });
  log(me.login,'logout','user',me.id,'');
  return {ok:true};
}
function pubUser(u){
  var o=readAll(SHEETS.org).filter(function(x){return x.id===u.orgId;})[0];
  return {login:u.login,name:u.name,role:u.role,orgId:u.orgId,
          organization:o?o.name:'LIFE PAY'};
}
function bootstrap(me){
  var st=settings(), orgs=readAll(SHEETS.org), execs=readAll(SHEETS.exe), all=readAll(SHEETS.req);
  var oMap={},eMap={};
  orgs.forEach(function(o){oMap[o.id]=o;}); execs.forEach(function(e){eMap[e.id]=e;});

  var scoped = me.role==='ADMIN' ? all
    : me.role==='ORG' ? all.filter(function(r){return r.orgId===me.orgId;})
    : all.filter(function(r){return r.executorId===me.execId;});

  var rows=scoped.map(function(r){
    var o=oMap[r.orgId]||{}, e=eMap[r.executorId]||{};
    return {id:r.id,no:r.no,ts:new Date(r.ts).getTime(),employee:r.employee,priority:r.priority,
      status:r.status,executorId:r.executorId,executor:e.name||'—',orgId:r.orgId,
      organization:o.name||'—',inn:o.inn||'',contact:o.contact||'',phone:o.phone||'',email:o.email||'',
      kktCount:r.kktCount,rnm:r.rnm,zn:r.zn,fn:r.fn,model:r.model,needTrip:r.needTrip,
      tripAddress:r.tripAddress,esp:r.esp,kep:r.kep,wifi:r.wifi,comment:r.comment,
      overdue:isOverdue(r,st.slaHours)};
  }).sort(function(a,b){return b.ts-a.ts;});

  var chart=[], today=new Date(); today.setHours(0,0,0,0);
  for(var i=6;i>=0;i--){
    var d0=today.getTime()-i*864e5, d1=d0+864e5;
    chart.push(rows.filter(function(r){return r.ts>=d0&&r.ts<d1;}).length);
  }
  var cnt=function(s){return rows.filter(function(r){return r.status===s;}).length;};

  return {
    me:pubUser(me),requests:rows,
    stats:{total:rows.length,newCount:cnt('Новая'),inWork:cnt('В работе'),wait:cnt('Ожидание'),
           done:cnt('Выполнено'),overdue:rows.filter(function(r){return r.overdue;}).length},
    chart:chart,
    orgs: me.role==='ADMIN'?orgs.map(clean):[],
    execs: execs.map(clean),
    users: me.role==='ADMIN'?readAll(SHEETS.usr).map(function(u){
      var o=oMap[u.orgId]||{};
      return {id:u.id,login:u.login,name:u.name,role:u.role,orgId:u.orgId,execId:u.execId,
              organization:o.name||'—',active:u.active!==false&&u.active!=='FALSE'};}):[],
    log: me.role==='ADMIN'?readAll(SHEETS.log).slice(-200).reverse().map(function(l){
      return {ts:new Date(l.ts).getTime(),login:l.login,action:l.action,entity:l.entity,details:l.details};}):[],
    settings: st
  };
}
function clean(o){ var c={}; for(var k in o) if(k!=='__row'&&k!=='passHash') c[k]=o[k]; return c; }

function ensureOrg(me,p){
  var orgs=readAll(SHEETS.org);
  var found=orgs.filter(function(o){
    return (p.inn && String(o.inn)===String(p.inn)) ||
           (p.organization && String(o.name).toLowerCase()===String(p.organization).toLowerCase()); })[0];
  if(found) return found;
  var o={id:uid(),name:p.organization||'Без названия',inn:p.inn||'',contact:p.contact||'',
         phone:p.phone||'',email:p.email||'',active:true,createdAt:new Date()};
  writeObj(SHEETS.org,o);
  log(me.login,'autoCreate','org',o.id,'Автосоздание: '+o.name);
  return o;
}
function nextNo(){
  var all=readAll(SHEETS.req), max=1000;
  all.forEach(function(r){ if(Number(r.no)>max) max=Number(r.no); });
  return max+1;
}
function saveRequest(me,p){
  var org=ensureOrg(me,p), st=settings(), all=readAll(SHEETS.req);
  if(p.id){
    var r=all.filter(function(x){return x.id===p.id;})[0];
    if(!r) throw new Error('Заявка не найдена');
    if(me.role==='ORG' && r.orgId!==me.orgId) throw new Error('Чужая заявка');
    if(me.role==='EXECUTOR' && r.executorId!==me.execId) throw new Error('Чужая заявка');
    var before=r.status;
    ['employee','priority','status','kktCount','rnm','zn','fn','model','needTrip','tripAddress','esp','kep','wifi','comment']
      .forEach(function(k){ if(p[k]!==undefined) r[k]=p[k]; });
    if(me.role==='ADMIN'){ r.orgId=org.id; if(p.executorId!==undefined) r.executorId=p.executorId; if(p.no) r.no=Number(p.no); }
    r.updatedAt=new Date();
    updateObj(SHEETS.req,r);
    log(me.login,'update','request',r.id,'No'+r.no+(before!==r.status?(' '+before+' → '+r.status):' поля изменены'));
    return {ok:true};
  }
  var n={id:uid(),no:nextNo(),ts:new Date(),employee:p.employee||me.name,priority:p.priority||'Средний',
    status:'Новая',executorId:p.executorId||'',orgId:org.id,kktCount:p.kktCount||1,rnm:p.rnm,zn:p.zn,fn:p.fn,
    model:p.model,needTrip:p.needTrip,tripAddress:p.tripAddress,esp:p.esp,kep:p.kep,wifi:p.wifi,
    comment:p.comment,updatedAt:new Date()};
  writeObj(SHEETS.req,n);
  log(me.login,'create','request',n.id,'No'+n.no+' · '+org.name);
  notify(n,org,st);
  return {ok:true};
}
function notify(r,org,st){
  try{
    if(st.notifyOrg && org.email)
      MailApp.sendEmail(org.email,'LIFE PAY — заявка No'+r.no+' принята',
        'Заявка No'+r.no+' зарегистрирована.\nКомментарий: '+(r.comment||'—'));
    if(st.notifyExec && r.executorId){
      var e=readAll(SHEETS.exe).filter(function(x){return x.id===r.executorId;})[0];
      if(e && e.email) MailApp.sendEmail(e.email,'LIFE PAY — новая заявка No'+r.no,
        'Организация: '+org.name+'\nПриоритет: '+r.priority+'\nКомментарий: '+(r.comment||'—'));
    }
  }catch(err){ log('system','mailError','request',r.id,String(err)); }
}
function setStatus(me,p){
  var r=readAll(SHEETS.req).filter(function(x){return x.id===p.id;})[0];
  if(!r) throw new Error('Заявка не найдена');
  if(me.role==='ORG' && r.orgId!==me.orgId) throw new Error('Чужая заявка');
  if(me.role==='EXECUTOR' && r.executorId!==me.execId) throw new Error('Чужая заявка');
  var b=r.status; r.status=p.status;
  if(p.comment!==undefined) r.comment=p.comment;
  r.updatedAt=new Date(); updateObj(SHEETS.req,r);
  log(me.login,'status','request',r.id,'No'+r.no+': '+b+' → '+r.status);
  return {ok:true};
}
function transfer(me,p){
  var r=readAll(SHEETS.req).filter(function(x){return x.id===p.id;})[0];
  if(!r) throw new Error('Заявка не найдена');
  var from=r.executorId, fromO=r.orgId;
  if(p.executorId) r.executorId=p.executorId;
  if(p.orgId) r.orgId=p.orgId;
  r.updatedAt=new Date(); updateObj(SHEETS.req,r);
  log(me.login,'transfer','request',r.id,'No'+r.no+': исп. '+from+' → '+r.executorId+'; орг. '+fromO+' → '+r.orgId);
  return {ok:true};
}
function saveSimple(me,sheet,p,entity,fields){
  var all=readAll(sheet);
  if(p.id){
    var it=all.filter(function(x){return x.id===p.id;})[0];
    if(!it) throw new Error('Не найдено');
    fields.forEach(function(k){ if(p[k]!==undefined) it[k]=p[k]; });
    updateObj(sheet,it); log(me.login,'update',entity,it.id,it.name||'');
    return {ok:true};
  }
  var o={id:uid(),active:true,createdAt:new Date()};
  fields.forEach(function(k){ o[k]=p[k]!==undefined?p[k]:''; });
  if(o.active==='') o.active=true;
  writeObj(sheet,o); log(me.login,'create',entity,o.id,o.name||'');
  return {ok:true};
}
function saveUser(me,p){
  var all=readAll(SHEETS.usr);
  if(p.id){
    var u=all.filter(function(x){return x.id===p.id;})[0];
    if(!u) throw new Error('Не найдено');
    ['name','role','orgId','execId','active'].forEach(function(k){ if(p[k]!==undefined) u[k]=p[k]; });
    if(p.pass) u.passHash=sha(p.pass);
    updateObj(SHEETS.usr,u); log(me.login,'update','user',u.id,u.login);
    return {ok:true};
  }
  if(all.some(function(x){return String(x.login)===String(p.login);})) throw new Error('Такой логин уже есть');
  var n={id:uid(),login:p.login,passHash:sha(p.pass||'123456'),name:p.name||p.login,
         role:p.role||'ORG',orgId:p.orgId||'',execId:p.execId||'',active:true};
  writeObj(SHEETS.usr,n); log(me.login,'create','user',n.id,n.login);
  return {ok:true};
}
function toggle(me,sheet,id,entity){
  var it=readAll(sheet).filter(function(x){return x.id===id;})[0];
  if(!it) throw new Error('Не найдено');
  it.active=!(it.active===true||it.active==='TRUE');
  updateObj(sheet,it); log(me.login,'toggle',entity,id,(it.name||it.login)+' → '+(it.active?'вкл':'выкл'));
  return {ok:true};
}
function delEntity(me,sheet,id,entity){
  var it=readAll(sheet).filter(function(x){return x.id===id;})[0];
  if(!it) throw new Error('Не найдено');
  log(me.login,'delete',entity,id,it.name||it.login||it.no||'');
  removeRow(sheet,it.__row);
  return {ok:true};
}
function saveSettings(me,p){
  var s=sh(SHEETS.set), cur=readAll(SHEETS.set), map={};
  cur.forEach(function(r){ map[r.key]=r; });
  Object.keys(p).forEach(function(k){
    if(map[k]) s.getRange(map[k].__row,2).setValue(p[k]);
    else s.appendRow([k,p[k]]);
  });
  log(me.login,'update','settings','',JSON.stringify(p));
  return settings();
}

/* ─── первичная установка: запустить один раз из редактора ─── */
function setup(){
  Object.keys(HEADERS).forEach(function(n){ sh(n); });
  if(!readAll(SHEETS.usr).length){
    writeObj(SHEETS.usr,{id:uid(),login:'admin',passHash:sha('admin'),name:'Администратор',
      role:'ADMIN',orgId:'',execId:'',active:true});
  }
  if(!readAll(SHEETS.set).length){
    [['slaHours',60],['notifyOrg',true],['notifyExec',true],
     ['mailFrom',Session.getActiveUser().getEmail()],['version','CRM 7.0']]
      .forEach(function(r){ sh(SHEETS.set).appendRow(r); });
  }
}
