/* Ajustes MELI — normalização de datas e competência */
"use strict";

/* ================================================================
   DATAS — YMD inteiro, imune a timezone
   ================================================================ */
function ymd(y,m,d){ return y*10000+m*100+d; }
function ymdParts(v){ return {y:Math.floor(v/10000), m:Math.floor(v/100)%100, d:v%100}; }
function ymdToUTC(v){ const p=ymdParts(v); return Date.UTC(p.y,p.m-1,p.d); }
function utcToYmd(ms){ const t=new Date(ms); return ymd(t.getUTCFullYear(),t.getUTCMonth()+1,t.getUTCDate()); }
function addDays(v,n){ return utcToYmd(ymdToUTC(v)+n*86400000); }
function diffDaysInclusive(a,b){ return Math.round((ymdToUTC(b)-ymdToUTC(a))/86400000)+1; }
function fmtYmd(v){ if(v==null||Number.isNaN(v)) return ""; const p=ymdParts(v); return String(p.d).padStart(2,"0")+"/"+String(p.m).padStart(2,"0")+"/"+p.y; }
function fmtShort(v){ const p=ymdParts(v); return String(p.d).padStart(2,"0")+"/"+String(p.m).padStart(2,"0"); }
function ymdToExcelDate(v){ const p=ymdParts(v); return new Date(Date.UTC(p.y,p.m-1,p.d,12)); } // meio-dia UTC evita shift
function isValidYmd(v){
  if(v==null) return false;
  const p=ymdParts(v);
  if(p.y<1990||p.y>2100||p.m<1||p.m>12||p.d<1) return false;
  return p.d <= new Date(Date.UTC(p.y,p.m,0)).getUTCDate();
}
function parseExcelDate(v){
  if(v===null||v===undefined||v==="") return null;
  if(v instanceof Date){
    if(isNaN(v.getTime())) return NaN;
    const r=utcToYmd(v.getTime()+43200000-((v.getTime()+43200000)%86400000));
    return isValidYmd(r)?r:NaN;
  }
  if(typeof v==="number"){
    if(!isFinite(v)||v<=0) return NaN;
    const r=utcToYmd(Date.UTC(1899,11,30)+Math.round(v)*86400000);
    return isValidYmd(r)?r:NaN;
  }
  if(typeof v==="string"){
    const s=v.trim(); if(!s) return null;
    let m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if(m){ let y=+m[3]; if(y<100)y+=2000; const r=ymd(y,+m[2],+m[1]); return isValidYmd(r)?r:NaN; }
    m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if(m){ const r=ymd(+m[1],+m[2],+m[3]); return isValidYmd(r)?r:NaN; }
    if(/^\d+(\.\d+)?$/.test(s)) return parseExcelDate(+s);
    return NaN;
  }
  return NaN;
}
function parseRateio(v){
  if(v===null||v===undefined||v==="") return null;
  let n;
  if(typeof v==="number") n=v;
  else{
    const s=String(v).replace("%","").replace(",",".").trim();
    if(!s) return null;
    n=parseFloat(s); if(!isFinite(n)) return NaN;
    if(String(v).includes("%")) n/=100;
  }
  if(!isFinite(n)) return NaN;
  if(n>1&&n<=100) n/=100;
  return n;
}
function buildCompetence(y,m){
  const lastDay=new Date(Date.UTC(y,m,0)).getUTCDate();
  const nm=m===12?{y:y+1,m:1}:{y:y,m:m+1};
  return {y,m, first:ymd(y,m,1), last:ymd(y,m,lastDay), days:lastDay,
    snapshot:ymd(y,m,CUT_DAY), label:MONTHS[m-1]+"/"+y, nextLabel:MONTHS[nm.m-1]+"/"+nm.y,
    fileLabel:MONTHS[m-1]+" "+y};
}
