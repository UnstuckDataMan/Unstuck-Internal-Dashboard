/* Company reserve alert in Operations, using the same server count as Gmail Set. */
(() => {
  const el=document.getElementById('gs-home-stock');if(!el)return;
  async function refresh(){try{const r=await fetch('/api/gmail-setup/summary',{credentials:'same-origin',cache:'no-store'});if(!r.ok)return;const s=await r.json();el.textContent=s.low_stock?`Order ${s.shortfall} more ${s.order_kind==='Phone'?(s.shortfall===1?'phone':'phones'):(s.shortfall===1?'laptop / PC':'laptops / PCs')}`:'No new devices needed';el.style.color=s.low_stock?'#fca5a5':'#a7f3d0';}catch{/* Keep the tool link available if a summary request fails. */}}
  refresh();setInterval(()=>{if(!document.hidden)refresh();},60000);
})();
