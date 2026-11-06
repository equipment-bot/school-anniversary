(function(root){
'use strict';
const C=typeof module!=='undefined'?require('./core.js'):RegistryCore;
const stable=x=>JSON.stringify(x,(k,v)=>v&&typeof v==='object'&&!Array.isArray(v)?Object.fromEntries(Object.keys(v).sort().map(k=>[k,v[k]])):v);
function plan(files){
 if(!files.length)throw Error('請先選取完整 JSON 備份');
 const groups=new Map(),issues=[],identities=new Map(),records=new Map(),payments=new Map();
 const add=(map,key,value,source)=>{if(!map.has(key))map.set(key,[]);const values=map.get(key),same=values.find(v=>stable(v.value)===stable(value));if(same)same.sources.push(source);else values.push({value,sources:[source]});};
 for(const f of files){const s=C.validate(f.data);if(s.demo)throw Error(f.name+' 是示範資料，不能混入正式整合');const source=f.name+'（'+s.station+'）';if(!groups.has(s.grade))groups.set(s.grade,{schema:1,grade:s.grade,station:'全校整合',students:[],records:{},payments:[],audit:[]});for(const student of s.students){const k=s.grade+'/'+C.key(student);add(identities,k,student,source);}for(const[k,r]of Object.entries(s.records)){// Include explicit cancellations as well as positive registrations.
 add(records,s.grade+'/'+k,{...r,guests:0,noDonation:!!r.noDonation},source);
 }for(const p of s.payments)add(payments,s.grade+'/'+p.id,p,source);}
 const issue=(id,title,options)=>{issues.push({id,title,options});return id;};
 for(const[k,v]of identities)if(v.length>1)issue('identity:'+k,k+' 學生資料不同',v);
 for(const[k,v]of records)if(v.length>1)issue('record:'+k,k+' 家長到場／備註不同，請選保留版本',v);
 for(const[k,v]of payments)if(v.length>1)issue('payment:'+k,'同一收款編號有不同版本（可能含作廢）',v);
 // Different IDs can still represent the same real-world payment. Ask; never delete automatically.
 const signatures=new Map();for(const[k,versions]of payments)for(const v of versions){const p=v.value;if(p.void)continue;const sig=k.split('/')[0]+'/'+p.student+'/'+p.amount+'/'+p.method;if(!signatures.has(sig))signatures.set(sig,new Set());signatures.get(sig).add(k);}
 const suspicious=new Set();for(const keys of signatures.values())if(keys.size>1)for(const k of keys)suspicious.add(k);
 for(const k of suspicious)issue('include:'+k,'同學生有多筆同金額同方式收款，請核對是否重複', [{value:true,sources:payments.get(k).flatMap(v=>v.sources)},{value:false,sources:['排除此筆（只影響整合結果）']}]);
 for(const[k,v]of records){const hasMoney=[...payments].some(([pk,pv])=>pk.startsWith(k.split('/')[0]+'/')&&pv.some(x=>x.value.student===k.split('/')[1]&&!x.value.void));if(hasMoney&&v.some(x=>x.value.noDonation))issue('donation:'+k,k+' 有「無捐款」與有效收款，請核對', [{value:'paid',sources:['保留有效收款，取消無捐款狀態']},{value:'none',sources:['保留無捐款，排除此學生全部收款']}]);}
 return{files:files.map(f=>f.name),groups,identities,records,payments,issues};
}
function finish(p,choices={}){
 const get=(id,versions)=>{if(versions.length===1)return versions[0].value;const n=choices[id];if(!Number.isInteger(n)||!versions[n])throw Error('請完成所有差異確認');return versions[n].value;};
 for(const i of p.issues)if(!Number.isInteger(choices[i.id])||!i.options[choices[i.id]])throw Error('請完成所有差異確認');
 const groups=new Map([...p.groups].map(([g,s])=>[g,structuredClone(s)]));
 for(const[k,v]of p.identities)groups.get(Number(k.split('/')[0])).students.push(structuredClone(get('identity:'+k,v)));
 for(const[k,v]of p.records){const r=structuredClone(get('record:'+k,v));if(choices['donation:'+k]!==undefined)r.noDonation=p.issues.find(i=>i.id==='donation:'+k).options[choices['donation:'+k]].value==='none';groups.get(Number(k.split('/')[0])).records[k.split('/')[1]]=r;}
 for(const[k,v]of p.payments){if(choices['include:'+k]!==undefined&&!p.issues.find(i=>i.id==='include:'+k).options[choices['include:'+k]].value)continue;const payment=structuredClone(get('payment:'+k,v)),g=Number(k.split('/')[0]);if(choices['donation:'+g+'/'+payment.student]!==undefined&&p.issues.find(i=>i.id==='donation:'+g+'/'+payment.student).options[choices['donation:'+g+'/'+payment.student]].value==='none')continue;groups.get(g).payments.push(payment);}
 const tables={'已到家長':[['年級','班級','座號','姓名','家長已到場','簽到時間','實收金額','備註','登記台','捐款狀態']],'已收款項':[['年級','班級','座號','姓名','有效實收金額']],'統計摘要':[['年級','班級／統計','應到家長','已到家長','未到家長','家長到場比例','有效實收金額']]};
 let total=0,arrived=0,money=0;for(const[g,s]of [...groups].sort((a,b)=>a[0]-b[0])){C.validate(s);const sheets=C.sheets(s),label=['','高一','高二','高三'][g];for(const name of ['已到家長','已收款項'])for(const row of sheets[name].slice(1))tables[name].push([label,...row]);for(const row of sheets['統計摘要'].slice(1))tables['統計摘要'].push([label,...row]);total+=s.students.length;arrived+=Object.values(s.records).filter(r=>r.present).length;money+=s.payments.filter(p=>!p.void).reduce((a,p)=>a+p.amount,0);}
 tables['統計摘要'].push(['全校（已匯入年級）','合計',total,arrived,total-arrived,(100*arrived/total).toFixed(1)+'%',money]);
 tables['整合來源']=[['匯入備份檔案'],...p.files.map(name=>[name]),[],['差異確認項目','選擇內容'],...p.issues.map(i=>[i.title,stable(i.options[choices[i.id]])])];
 return{tables,groups:[...groups.values()],total,arrived,money};
}
root.RegistryMerge={plan,finish};if(typeof module!=='undefined')module.exports=root.RegistryMerge;
})(typeof window==='undefined'?globalThis:window);
