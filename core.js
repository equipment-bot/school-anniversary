(function(root){
'use strict';
const key=s=>`${s.classroom}-${s.seat}`;
const norm=s=>String(s).normalize('NFKC').replace(/[\s\u200b\ufeff]/gu,'');
function validate(s){
 if(!s||s.schema!==1||![1,2,3].includes(s.grade)||typeof s.station!=='string'||s.station.length>80||!Array.isArray(s.students)||!s.students.length||s.students.length>3000||!s.records||typeof s.records!=='object'||Array.isArray(s.records)||!Array.isArray(s.payments)||!Array.isArray(s.audit))throw Error('備份格式不正確');
 const ids=new Set();for(const x of s.students){if(x.grade!==s.grade||typeof x.name!=='string'||!x.name.trim()||x.name.length>80||!/^\d{1,6}$/.test(x.classroom)||!Number.isInteger(x.seat)||x.seat<1||x.seat>999||ids.has(key(x)))throw Error('學生名冊格式或班級座號重複');ids.add(key(x));}
 for(const [k,r] of Object.entries(s.records)){if(!ids.has(k)||!r||typeof r.present!=='boolean'||(r.noDonation!==undefined&&typeof r.noDonation!=='boolean')||!Number.isInteger(r.guests)||r.guests<0||r.guests>99||typeof r.note!=='string'||r.note.length>2000||typeof r.time!=='string')throw Error('簽到紀錄格式不正確');}
 const paymentIds=new Set();for(const p of s.payments){if(!p||!ids.has(p.student)||typeof p.id!=='string'||paymentIds.has(p.id)||!Number.isSafeInteger(p.amount)||p.amount<1||p.amount>99999999||!['現金','轉帳','其他'].includes(p.method)||typeof p.time!=='string'||typeof p.note!=='string'||p.note.length>2000||(p.void!==null&&(typeof p.void!=='object'||!p.void||typeof p.void.reason!=='string'||typeof p.void.time!=='string')))throw Error('收款紀錄格式不正確');paymentIds.add(p.id);}
 if(Object.entries(s.records).some(([k,r])=>r.noDonation&&s.payments.some(p=>p.student===k&&!p.void)))throw Error('無捐款與有效收款紀錄矛盾');
 if(s.audit.length>100000||s.audit.some(a=>!a||typeof a.time!=='string'||typeof a.action!=='string'))throw Error('異動紀錄格式不正確');return s;
}
function sheets(s){
 validate(s);const active=s.payments.filter(p=>!p.void),sum=ps=>ps.reduce((a,p)=>a+p.amount,0),paid=k=>sum(active.filter(p=>p.student===k)),arrived=Object.values(s.records).filter(r=>r.present).length;
 const parents=[['班級','座號','姓名','家長已到場','家長簽到時間','實收金額','備註','登記台','捐款狀態']];
 const money=[['班級','座號','姓名','有效實收金額']];const classes=new Map();
 for(const x of [...s.students].sort((a,b)=>Number(a.classroom)-Number(b.classroom)||a.seat-b.seat)){const k=key(x),r=s.records[k],amount=paid(k);if(r?.present)parents.push([x.classroom,x.seat,x.name,'是',r.time,amount,r.note,s.station,r.noDonation?'無捐款':amount?'已收款':'未登記']);if(amount)money.push([x.classroom,x.seat,x.name,amount]);if(!classes.has(x.classroom))classes.set(x.classroom,{count:0,present:0,amount:0});const c=classes.get(x.classroom);c.count++;c.present+=r?.present?1:0;c.amount+=amount;}
 const ratio=(a,n)=>(100*a/n).toFixed(1)+'%';const summary=[['班級','名冊學生數（應到家長數）','已到家長','未到家長','家長到場比例','有效實收金額'],...Array.from(classes,([k,v])=>[k,v.count,v.present,v.count-v.present,ratio(v.present,v.count),v.amount]),['合計',s.students.length,arrived,s.students.length-arrived,ratio(arrived,s.students.length),sum(active)]];
 return{'已到家長':parents,'已收款項':money,'統計摘要':summary};
}
// Minimal OOXML writer. Text is always an inline string, never an Excel formula.
function xlsx(s){return xlsxTables(sheets(s));}
function xlsxTables(tables){const enc=new TextEncoder(),xml=v=>String(v).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
 const col=n=>{let r='';for(n++;n;n=Math.floor((n-1)/26))r=String.fromCharCode(65+(n-1)%26)+r;return r;};
 const files=[],add=(name,value)=>files.push([enc.encode(name),enc.encode(value)]);const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 add('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+Object.keys(tables).map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')+'</Types>');
 add('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
 add('xl/workbook.xml',`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`+Object.keys(tables).map((name,i)=>`<sheet name="${xml(name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')+'</sheets></workbook>');
 add('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+Object.keys(tables).map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')+'</Relationships>');
 Object.values(tables).forEach((rows,i)=>add(`xl/worksheets/sheet${i+1}.xml`,`<worksheet xmlns="${ns}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="20" width="22" customWidth="1"/></cols><sheetData>`+rows.map((row,r)=>`<row r="${r+1}">`+row.map((v,c)=>typeof v==='number'?`<c r="${col(c)}${r+1}"><v>${v}</v></c>`:`<c r="${col(c)}${r+1}" t="inlineStr"><is><t xml:space="preserve">${xml(v)}</t></is></c>`).join('')+'</row>').join('')+'</sheetData></worksheet>'));
 const crc=b=>{let c=0xffffffff;for(const v of b){c^=v;for(let i=0;i<8;i++)c=(c>>>1)^((c&1)?0xedb88320:0);}return(c^0xffffffff)>>>0;};const chunks=[],central=[];let offset=0;const header=(size,fn)=>{const b=new Uint8Array(size);fn(new DataView(b.buffer));return b;};
 for(const[n,b]of files){const check=crc(b);const h=header(30,v=>{v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint32(14,check,true);v.setUint32(18,b.length,true);v.setUint32(22,b.length,true);v.setUint16(26,n.length,true);});const c=header(46,v=>{v.setUint32(0,0x02014b50,true);v.setUint16(4,20,true);v.setUint16(6,20,true);v.setUint32(16,check,true);v.setUint32(20,b.length,true);v.setUint32(24,b.length,true);v.setUint16(28,n.length,true);v.setUint32(42,offset,true);});chunks.push(h,n,b);central.push(c,n);offset+=h.length+n.length+b.length;}
 const length=central.reduce((a,b)=>a+b.length,0);chunks.push(...central,header(22,v=>{v.setUint32(0,0x06054b50,true);v.setUint16(8,files.length,true);v.setUint16(10,files.length,true);v.setUint32(12,length,true);v.setUint32(16,offset,true);}));const result=new Uint8Array(offset+length+22);let at=0;for(const b of chunks){result.set(b,at);at+=b.length;}return result;
}
root.RegistryCore={key,norm,validate,sheets,xlsx,xlsxTables};if(typeof module!=='undefined')module.exports=root.RegistryCore;
})(typeof window==='undefined'?globalThis:window);
