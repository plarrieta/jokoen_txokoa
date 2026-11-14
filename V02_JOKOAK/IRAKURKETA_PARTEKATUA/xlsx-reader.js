async function unzipEntries(buffer){
 const bytes=new Uint8Array(buffer),v=new DataView(buffer);let eocd=-1;
 for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--){if(v.getUint32(i,true)===0x06054b50){eocd=i;break;}}
 if(eocd<0)throw Error('Fitxategia ez da XLSX/ZIP baliozkoa.');
 const count=v.getUint16(eocd+10,true),offset=v.getUint32(eocd+16,true),decoder=new TextDecoder();
 const out={};let p=offset;
 for(let k=0;k<count;k++){
  if(v.getUint32(p,true)!==0x02014b50)throw Error('XLSX fitxategiaren egitura hondatuta dago.');
  const method=v.getUint16(p+10,true),size=v.getUint32(p+20,true),nlen=v.getUint16(p+28,true),elen=v.getUint16(p+30,true),clen=v.getUint16(p+32,true),lo=v.getUint32(p+42,true);
  const name=decoder.decode(bytes.slice(p+46,p+46+nlen));p+=46+nlen+elen+clen;
  if(!name.startsWith('xl/'))continue;
  if(v.getUint32(lo,true)!==0x04034b50)throw Error('XLSX: fitxategi-sarrera hondatuta.');
  const start=lo+30+v.getUint16(lo+26,true)+v.getUint16(lo+28,true);
  let data=bytes.slice(start,start+size);
  if(method===8){
   if(typeof DecompressionStream==='undefined')throw Error('Nabigatzaile honek ez du XLSX deskonpresioa onartzen. Erabili Chrome eguneratua.');
   const stream=new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
   data=new Uint8Array(await new Response(stream).arrayBuffer());
  }else if(method!==0)throw Error('XLSX fitxategiak konpresio ezezaguna du.');
  out[name]=decoder.decode(data);
 }
 return out;
}
function xmlDoc(x){if(!x)throw Error('XLSXren beharrezko orri bat falta da.');let d=new DOMParser().parseFromString(x,'text/xml');if(d.querySelector('parsererror'))throw Error('XLSXren XML datuak hondatuta daude.');return d;}
function tags(node,name){return Array.from(node.getElementsByTagNameNS('*',name));}
function colNumber(ref){let x=0;for(let c of (ref.match(/^[A-Z]+/i)||[''])[0].toUpperCase()){x=x*26+c.charCodeAt(0)-64;}return x-1;}
function sheetCells(xml,strings){
 const doc=xmlDoc(xml),rows=new Map();
 for(const r of tags(doc,'row')){
  const rn=Number(r.getAttribute('r'));const record=[];
  for(const c of tags(r,'c')){
    const idx=colNumber(c.getAttribute('r')||'A');let value='';
    if(c.getAttribute('t')==='inlineStr')value=tags(c,'t').map(x=>x.textContent).join('');
    else{const v=tags(c,'v')[0];value=v?v.textContent:'';if(c.getAttribute('t')==='s')value=strings[Number(value)]||'';}
    record[idx]=value;
  }
  rows.set(rn,record);
 }
 return rows;
}
async function readWorkbook(file){
 const entries=await unzipEntries(await file.arrayBuffer());
 const wb=xmlDoc(entries['xl/workbook.xml']),rels=xmlDoc(entries['xl/_rels/workbook.xml.rels']);
 const relMap={};for(const r of tags(rels,'Relationship'))relMap[r.getAttribute('Id')]=r.getAttribute('Target');
 const shared=entries['xl/sharedStrings.xml']?tags(xmlDoc(entries['xl/sharedStrings.xml']),'si').map(si=>tags(si,'t').map(t=>t.textContent).join('')):[];
 const sheets={};
 for(const sh of tags(wb,'sheet')){
  const name=sh.getAttribute('name'),rid=sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')||sh.getAttribute('r:id');
  const raw=relMap[rid];if(!raw)continue;
  const clean=raw.startsWith('/')?raw.slice(1):'xl/'+raw.replace(/^\.\//,'');
  if(['IRAKURKETA','PARAGRAFOAK','GALDERAK'].includes(name))sheets[name]=sheetCells(entries[clean],shared);
 }
 for(const name of ['IRAKURKETA','PARAGRAFOAK','GALDERAK'])if(!sheets[name])throw Error(name+' izeneko fitxa falta da.');
 const meta=sheets.IRAKURKETA,title=String((meta.get(4)||[])[1]||'').trim();
 if(!title)throw Error('IRAKURKETA!B4: izenburua falta da.');
 const pmap=new Map();
 for(const [row,c] of sheets.PARAGRAFOAK){if(row<4)continue;
  const t=String(c[1]||'').trim();if(!t)continue;
  const id=Number(c[0]);if(!Number.isInteger(id)||id<1||pmap.has(id))throw Error('PARAGRAFOAK!A'+row+': paragrafo ID desegokia edo errepikatua.');
  pmap.set(id,t);
 }
 if(!pmap.size)throw Error('PARAGRAFOAK fitxan ez dago testurik.');
 const ids=Array.from(pmap.keys()).sort((a,b)=>a-b);
 ids.forEach((id,i)=>{if(id!==i+1)throw Error('PARAGRAFOAK: ID zenbakiak 1, 2, 3... izan behar dira.');});
 const groups=ids.map(()=>[]),questionNums=ids.map(()=>new Set());
 for(const [row,c] of sheets.GALDERAK){if(row<4)continue;
  const text=String(c[3]||'').trim();if(!text)continue;
  const pid=Number(c[0]),num=Number(c[1]),type=String(c[2]||'').trim().toUpperCase();
  if(!pmap.has(pid))throw Error('GALDERAK!A'+row+': paragrafo ID ezezaguna.');
  if(!Number.isInteger(num)||num<1||questionNums[pid-1].has(num))throw Error('GALDERAK!B'+row+': galdera-zenbakia desegokia edo errepikatua.');
  questionNums[pid-1].add(num);
  if(!['AUKERAK','EGIA_GE_ZURRA','LABURRA'].includes(type))throw Error('GALDERAK!C'+row+': mota ezezaguna: '+type);
  let options=[],correct=-1,accepted=[];
  if(type==='AUKERAK'){
   options=c.slice(4,8).map(x=>String(x||'').trim()).filter(Boolean);
   correct='ABCD'.indexOf(String(c[8]||'').trim().toUpperCase());
   if(options.length<2||correct<0||correct>=options.length)throw Error('GALDERAK!'+row+': gutxienez bi aukera eta I zutabean A/B/C/D zuzen bat behar dira.');
  }else if(type==='EGIA_GE_ZURRA'){
   options=['EGIA','GEZURRA'];const ans=String(c[8]||'').trim().toUpperCase();correct=options.indexOf(ans);
   if(correct<0)throw Error('GALDERAK!I'+row+': EGIA edo GEZURRA idatzi.');
  }else{
   accepted=String(c[9]||'').split('|').map(a=>a.trim()).filter(Boolean);
   if(!accepted.length)throw Error('GALDERAK!J'+row+': erantzun laburrerako onartutako erantzunak falta dira.');
  }
  groups[pid-1].push({num,type,text,options,correct,accepted,explanation:String(c[10]||'').trim()});
 }
 groups.forEach((qs,i)=>{
  qs.sort((a,b)=>a.num-b.num);
  if(qs.length<1||qs.length>3)throw Error((i+1)+'. paragrafoak 1-3 galdera behar ditu (orain: '+qs.length+').');
  qs.forEach((q,j)=>{if(q.num!==j+1)throw Error((i+1)+'. paragrafoan galderen zenbakiak 1,2,3... izan behar dira.');});
 });
 return {title,paragraphs:ids.map(id=>pmap.get(id)),questions:groups};
}

window.IP_readWorkbook = readWorkbook;
