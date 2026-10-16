/* Company reserve alert in Operations, using the same server count as Gmail Set. */
(() => {
  const el=document.getElementById('gs-home-stock');if(!el)return;
  async function refresh(){try{const r=await fetch('/api/gmail-setup/summary',{credentials:'same-origin',cache:'no-store'});if(!r.ok)return;const s=await r.json();el.textContent=`Setup capacity: ${s.available_setups} / ${s.target} required. ${s.todo} to do${s.planned_todo?` (${s.actual_todo} actual + ${s.planned_todo} planned)`:""} + ${s.contingency} contingency setups. ${s.phone_available} from phones + ${s.pc_available} from laptops / PCs${s.purchased_setups?` + ${s.purchased_setups} purchases already covered`:''}. `+(s.low_stock?`${s.setup_shortfall} setups short. Order ${s.shortfall} more ${s.order_kind==='Phone'?'phones':'laptops / PCs'} at ${s.order_capacity} setups each, rounded up.`:'Setup capacity covered.')+` Limits: ${s.phone_setups} setups/phone and ${s.pc_setups} setups/laptop or PC.`;el.style.color=s.low_stock?'#fca5a5':'#a7f3d0';}catch{/* Keep the tool link available if a summary request fails. */}}
  refresh();setInterval(()=>{if(!document.hidden)refresh();},60000);
})();
