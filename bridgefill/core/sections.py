"""共用或重複 table id 時，以標題文字／區段 ID 定位，不固定 tr。"""
SECTION_TABLE = r"""({name,ids,tid,token,headings})=>{
 const norm=s=>(s||'').normalize('NFKC').replace(/\s+/g,'').replace(/[–—]/g,'-').replace(/[（(]\d+[)）]$/,'');
 const visible=e=>e.getClientRects().length && getComputedStyle(e).visibility!=='hidden';
 const tables=Array.from(document.querySelectorAll('table[id]')).filter(t=>t.id===tid && visible(t));
 const nodes=Array.from(document.querySelectorAll('span,h1,h2,h3,h4,legend,caption,div,td'));
 const labels=nodes.filter(e=>visible(e) && headings.some(h=>norm(e.textContent)===norm(h)));
 const owners=nodes.filter(e=>visible(e) && (ids.includes(e.id)||norm(e.textContent)===norm(name)));
 const found=new Set();
 for(const owner of owners){
   const inside=tables.filter(t=>t.contains(owner));
   if(inside.length===1){found.add(inside[0]);continue;}
   const next=tables.find(t=>owner.compareDocumentPosition(t)&Node.DOCUMENT_POSITION_FOLLOWING);
   if(!next)continue;
   const crossed=labels.some(other=>other!==owner && !owner.contains(other) &&
      (owner.compareDocumentPosition(other)&Node.DOCUMENT_POSITION_FOLLOWING) &&
      (other.compareDocumentPosition(next)&Node.DOCUMENT_POSITION_FOLLOWING));
   if(!crossed)found.add(next);
 }
 if(!found.size && tid==='pylon_anchor' && tables.length===1)found.add(tables[0]);
 if(found.size>1)return {error:'同構件標題對應多個可見表格，無法唯一辨識'};
 if(!found.size)return {missing:true};
 const table=Array.from(found)[0];
 table.setAttribute('data-bridgefill-section',token);
 return {selector:'[data-bridgefill-section="'+token+'"]',order:Array.from(document.querySelectorAll('table[id]')).indexOf(table)};
}"""
