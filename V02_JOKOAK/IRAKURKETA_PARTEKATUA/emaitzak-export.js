/* IRAKURKETA PARTEKATUA V24 — esportazio lokala, kanpoko liburutegirik gabe */
(function(){'use strict';
const encoder=new TextEncoder();
function clean(x){return String(x==null?'':x).replace(/\u0000/g,'');}
function xml(x){return clean(x).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}
function html(x){return xml(x);}
function norm(x){return clean(x).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim().replace(/\s+/g,' ');}
function dateText(x){if(!x)return '';try{const d=typeof x.toDate==='function'?x.toDate():x.seconds!=null?new Date(x.seconds*1000):new Date(x);return Number.isNaN(+d)?'':d.toLocaleString('eu-ES');}catch{return '';}}
function evaluate(q,a){if(!a)return '';if(!q)return 'EZIN EBALUATU';if(q.type==='LABURRA')return (q.accepted||[]).some(v=>norm(v)===norm(a.answer))?'ZUZEN':'OKER';return Number(a.answer)===Number(q.correct)?'ZUZEN':'OKER';}
function answerLabel(q,a){if(!a)return '';if(q?.type==='LABURRA'||typeof a.answer==='string')return clean(a.answer);return q?.options?.[Number(a.answer)]||String(a.answer??'');}
function correctLabel(q){if(!q)return '';return q.type==='LABURRA'?(q.accepted||[]).join(' / '):q.options?.[q.correct]||'';}
function normalizedData(d){
 const people=d.people||[],answers=d.responseRows||[],summaries=d.summaries||[],qs=d.questions;
 const groups=new Map();const getQ=(p,q)=>qs?.[p]?.[q];
 for(const a of answers){const k=a.paragraphIndex+'|'+a.questionIndex;const g=groups.get(k)||[];g.push(a);groups.set(k,g);}
 const keyList=[...groups.keys()].sort((a,b)=>{const x=a.split('|').map(Number),y=b.split('|').map(Number);return x[0]-y[0]||x[1]-y[1];});
 const pmap=new Map(people.map(p=>[p.uid,p.nick]));
 const answerRows=[];
 for(const k of keyList){const [pi,qi]=k.split('|').map(Number),q=getQ(pi,qi),group=groups.get(k),byUid=new Map(group.map(a=>[a.uid,a]));
  const allUids=new Set([...pmap.keys(),...byUid.keys()]);
  for(const uid of allUids){const a=byUid.get(uid);answerRows.push({nick:pmap.get(uid)||a?.nick||'IKASLEA',paragraph:pi+1,question:qi+1,questionText:q?.text||'',type:q?.type||'',answer:answerLabel(q,a),correctAnswer:correctLabel(q),result:a?evaluate(q,a):'ERANTZUN GABE',submitted:dateText(a?.submittedAt),uid,answered:!!a});}
 }
 // Kontagailuak: soilik bidalitako eta galdera kargatuta duten erantzunak.
 const summaryRows=people.map(p=>{
  const own=answerRows.filter(a=>a.uid===p.uid),valid=own.filter(a=>a.answered&&(a.result==='ZUZEN'||a.result==='OKER'));
  const ok=valid.filter(a=>a.result==='ZUZEN').length,err=valid.length-ok;
  return {nick:p.nick,questions:own.length,answered:own.filter(a=>a.answered).length,ok,err,percent:valid.length?Math.round(ok*10000/valid.length)/100+'%':'—',ungraded:own.filter(a=>a.result==='EZIN EBALUATU').length};
 });
 const unregistered=[...new Set(answers.filter(a=>!pmap.has(a.uid)).map(a=>a.uid))];
 for(const uid of unregistered){const own=answerRows.filter(a=>a.uid===uid),valid=own.filter(a=>a.answered&&(a.result==='ZUZEN'||a.result==='OKER')),ok=valid.filter(a=>a.result==='ZUZEN').length;summaryRows.push({nick:own[0]?.nick||'IKASLEA',questions:own.length,answered:own.filter(a=>a.answered).length,ok,err:valid.length-ok,percent:valid.length?Math.round(ok*10000/valid.length)/100+'%':'—',ungraded:own.filter(a=>a.result==='EZIN EBALUATU').length});}
 const summaryTexts=summaries.map(s=>({nick:pmap.get(s.uid)||s.nick||'IKASLEA',paragraph:Number(s.paragraphIndex)+1,text:s.text||'',submitted:dateText(s.submittedAt)})).sort((a,b)=>a.paragraph-b.paragraph);
 return {title:d.title||'',pin:d.pin,createdAt:d.createdAt,summaryRows,answerRows,summaryTexts,anyMissing:answerRows.some(a=>a.result==='EZIN EBALUATU')};
}
function colName(i){let x='';for(i++;i>0;i=Math.floor((i-1)/26))x=String.fromCharCode(65+(i-1)%26)+x;return x;}
function worksheet(rows,widths){const last=colName(rows[0].length-1);let str=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${last}${rows.length}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="16"/><cols>`;
 widths.forEach((w,i)=>{str+=`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;});
 str+='</cols><sheetData>';
 rows.forEach((row,ri)=>{str+=`<row r="${ri+1}">`;row.forEach((v,ci)=>{const ref=colName(ci)+(ri+1);str+=`<c r="${ref}" s="${ri===0?1:0}" t="inlineStr"><is><t xml:space="preserve">${xml(v)}</t></is></c>`;});str+='</row>';});
 return str+`</sheetData><autoFilter ref="A1:${last}${rows.length}"/></worksheet>`;
}
const crcTable=(()=>{const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?0xedb88320^(c>>>1):c>>>1;t[n]=c>>>0;}return t;})();
function crc32(a){let c=0xffffffff;for(const b of a)c=crcTable[(c^b)&255]^(c>>>8);return (c^0xffffffff)>>>0;}
function zip(entries){const parts=[],dir=[];let offset=0;function u16(n){return [n&255,(n>>>8)&255];}function u32(n){return [n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255];}
 for(const [filename,body] of entries){const name=encoder.encode(filename),data=encoder.encode(body),crc=crc32(data);
  const local=new Uint8Array([0x50,0x4b,0x03,0x04,...u16(20),...u16(0x800),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(data.length),...u32(data.length),...u16(name.length),...u16(0),...name]);parts.push(local,data);
  dir.push(new Uint8Array([0x50,0x4b,0x01,0x02,...u16(20),...u16(20),...u16(0x800),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(data.length),...u32(data.length),...u16(name.length),...u16(0),...u16(0),...u16(0),...u16(0),...u32(0),...u32(offset),...name]));offset+=local.length+data.length;
 }
 const cdSize=dir.reduce((n,x)=>n+x.length,0);const tail=new Uint8Array([0x50,0x4b,0x05,0x06,...u16(0),...u16(0),...u16(entries.length),...u16(entries.length),...u32(cdSize),...u32(offset),...u16(0)]);
 const bytes=new Uint8Array(offset+cdSize+tail.length);let p=0;for(const v of [...parts,...dir,tail]){bytes.set(v,p);p+=v.length;}return bytes;
}
function toXlsx(raw){const d=normalizedData(raw);
 const sheets=[
 ['IKASLEAK', [['IKASLEA','GALDERAK GUZTIRA','BIDALITAKOAK','ZUZEN','OKER','ZUZEN %','EBALUATU GABE'],...d.summaryRows.map(a=>[a.nick,a.questions,a.answered,a.ok,a.err,a.percent,a.ungraded])], [24,23,20,12,12,16,19]],
 ['ERANTZUNAK', [['IKASLEA','PARAGRAFOA','GALDERA','GALDERAREN TESTUA','MOTA','IKASLEAREN ERANTZUNA','ERANTZUN ZUZENA','EMAITZA','BIDALKETA DATA'],...d.answerRows.map(a=>[a.nick,a.paragraph,a.question,a.questionText,a.type,a.answer,a.correctAnswer,a.result,a.submitted])],[24,15,12,65,19,48,48,19,22]],
 ['LABURPENAK', [['IKASLEA','PARAGRAFOA','LABURPENA','BIDALKETA DATA'],...d.summaryTexts.map(a=>[a.nick,a.paragraph,a.text,a.submitted])],[24,15,100,22]]
 ];
 let bookXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>';
 sheets.forEach(([name],i)=>bookXml+=`<sheet name="${xml(name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`);bookXml+='</sheets></workbook>';
 const entries=[
 ['[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')+'</Types>'],
 ['_rels/.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
 ['xl/workbook.xml',bookXml],
 ['xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')+`<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
 ['xl/styles.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF183D50"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>']
 ];
 sheets.forEach(([name,rows,widths],i)=>entries.push([`xl/worksheets/sheet${i+1}.xml`,worksheet(rows,widths)]));
 return zip(entries);
}
function toPrintHtml(raw){const d=normalizedData(raw);
 const table=(headers,rows)=>`<table><thead><tr>${headers.map(h=>`<th>${html(h)}</th>`).join('')}</tr></thead><tbody>${rows.length?rows.map(r=>`<tr>${r.map(c=>`<td>${html(c)}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="${headers.length}">Ez dago daturik</td></tr>`}</tbody></table>`;
 const sum=table(['Ikaslea','Bidali','Zuzen','Oker','Zuzen %'],d.summaryRows.map(a=>[a.nick,a.answered,a.ok,a.err,a.percent]));
 const detail=table(['Ikaslea','P.','G.','Galdera','Erantzuna','Zuzena','Emaitza'],d.answerRows.map(a=>[a.nick,a.paragraph,a.question,a.questionText,a.answer,a.correctAnswer,a.result]));
 const sums=table(['Ikaslea','P.','Laburpena'],d.summaryTexts.map(a=>[a.nick,a.paragraph,a.text]));
 const warning=d.anyMissing?'<p class="warn">Excel kargatu gabe: galderen zuzentasuna ezin da ebaluatu.</p>':'';
 return `<!doctype html><html lang="eu"><head><meta charset="utf-8"><title>IRAKURKETA EMAITZAK ${html(d.pin)}</title><style>body{font:11px/1.45 Arial,sans-serif;color:#142631;margin:30px}h1{font-size:23px;color:#145777}h2{font-size:17px;color:#145777;border-bottom:2px solid #8ab5c1;padding-bottom:5px;margin-top:22px}p{margin:5px 0}table{border-collapse:collapse;width:100%;font-size:10px;margin:12px 0;table-layout:auto}th{background:#183d50;color:white;text-align:left}th,td{border:1px solid #bccbd1;padding:5px;vertical-align:top;overflow-wrap:anywhere}tr:nth-child(even) td{background:#eef5f5}.warn{color:#a30}.no-print{background:#eaf1f4;border:1px solid #99b5c1;padding:12px;font-size:13px}button{padding:8px 15px;background:#174f69;color:white;border:0;cursor:pointer}@page{size:A4 landscape;margin:12mm}@media print{body{margin:0}.no-print{display:none}thead{display:table-header-group}tr{break-inside:avoid}h2{break-after:avoid}}</style></head><body><div class="no-print"><b>PDF fitxategia:</b> sakatu Inprimatu, eta Chrome-n aukeratu <b>Guardar como PDF / Gorde PDF gisa</b>. <button onclick="window.print()">INPRIMATU / GORDE PDF</button></div><h1>IRAKURKETA PARTEKATUA — EMAITZAK</h1><p><b>Saioa (PIN):</b> ${html(d.pin)} &nbsp; <b>Irakurketa:</b> ${html(d.title)} &nbsp; <b>Deskarga:</b> ${html(d.createdAt)}</p>${warning}<h2>IKASLEEN EMAITZAK</h2>${sum}<h2>GALDEREN ERANTZUNAK</h2>${detail}<h2>LABURPENAK</h2>${sums}</body></html>`;
}
window.IPResultsExport={toXlsx,toPrintHtml,normalizedData};
})();
