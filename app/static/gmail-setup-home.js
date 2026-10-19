/* Company reserve alert in Operations, using the same server count as Gmail Set. */
(() => {
  const el=document.getElementById('gs-home-stock');if(!el)return;
  let pending=false,queued=false;
  async function refresh(){
    if(pending){queued=true;return;}pending=true;
    const controller=new AbortController(),timeout=setTimeout(()=>controller.abort(),10000);
    try{
      const r=await fetch('/api/gmail-setup/summary',{credentials:'same-origin',cache:'no-store',signal:controller.signal});
      if(!r.ok)throw Error('Summary unavailable');
      const s=await r.json();
      el.textContent=s.low_stock?`Order ${s.shortfall} more ${s.order_kind==='Phone'?(s.shortfall===1?'phone':'phones'):(s.shortfall===1?'laptop / PC':'laptops / PCs')}`:'No new devices needed';
      el.style.color=s.low_stock?'#fca5a5':'#a7f3d0';
    }catch{
      // Do not leave a previous figure looking current after a failed read.
      el.textContent='Device needs unavailable';el.style.color='#c4b5fd';
    }finally{clearTimeout(timeout);pending=false;if(queued){queued=false;refresh();}}
  }
  refresh();setInterval(()=>{if(!document.hidden)refresh();},30000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)refresh();});
  window.addEventListener('pageshow',event=>{if(event.persisted)refresh();});
  window.addEventListener('popstate',()=>{if(location.hash==='#operations')refresh();});
  document.getElementById('btn-operations')?.addEventListener('click',refresh);
})();
