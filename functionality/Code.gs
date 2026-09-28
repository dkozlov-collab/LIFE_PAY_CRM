function intake(p){
  var org = ensureOrgPublic(p);
  var n = {id:uid(), no:nextNo(), ts:new Date(), source:p.source||'Веб-форма',
    employee:p.employee||'Веб-форма', priority:p.priority||'Средний', status:'Новая',
    executorId:'', orgId:org.id, kktCount:p.kktCount||1, rnm:p.rnm||'', zn:p.zn||'',
    fn:p.fn||'', model:p.model||'', needTrip:p.needTrip||'Нет', tripAddress:p.tripAddress||'',
    esp:p.esp||'Нет', kep:p.kep||'Нет', wifi:p.wifi||'Нет', comment:p.comment||'', updatedAt:new Date()};
  writeObj(SHEETS.req, n);
  log('веб-форма','intake','request',n.id,'№'+n.no+' · '+org.name);
  notify(n, org, settings());
  return {ok:true, no:n.no};
}
function ensureOrgPublic(p){
  var f=readAll(SHEETS.org).filter(function(o){
    return (p.inn && String(o.inn)===String(p.inn)) ||
           (p.organization && String(o.name).toLowerCase()===String(p.organization).toLowerCase());})[0];
  if(f) return f;
  var o={id:uid(),name:p.organization||'Без названия',inn:p.inn||'',contact:p.contact||'',
         phone:p.phone||'',email:p.email||'',active:true,createdAt:new Date()};
  writeObj(SHEETS.org,o);
  log('веб-форма','autoCreate','org',o.id,'Автосоздание: '+o.name);
  return o;
}
