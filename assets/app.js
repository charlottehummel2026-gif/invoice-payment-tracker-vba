'use strict';
(() => {
  const E = TrackerEngine;
  const $ = id => document.getElementById(id);
  let state=E.clone(DEMO_SEED), view='invoices', overdue=false, highlighted='', mode='invoice', tourStep=-1, tourId='';
  const cash = value => new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',minimumFractionDigits:0,maximumFractionDigits:value%100?2:0}).format(value/100);
  const date = value => new Date(value+'T00:00:00Z').toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric',timeZone:'UTC'});
  const esc = value => String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const statusClass = value => value.toLowerCase().replace(/ /g,'-');
  const lockTour = locked => ['as-of','new-invoice','new-payment'].forEach(id=>$(id).disabled=locked);
  function message(text) {$('demo-message').textContent=text;}
  function render() {
    const summary=E.summarize(state);
    $('as-of').value=state.asOf;
    for(const [id,key] of [['total-invoiced','invoiced'],['total-received','received'],['total-outstanding','outstanding'],['total-past-due','pastDue']]) $(id).textContent=cash(summary[key]);
    $('open-count').textContent=`${summary.open} open invoice${summary.open===1?'':'s'}`;
    const max=Math.max(1,...summary.aging.map(a=>a.value));
    $('aging-bars').innerHTML=summary.aging.map(a=>`<div class="aging-item" role="listitem"><span>${a.label}</span><strong>${cash(a.value)}</strong><div class="bar-track" aria-hidden="true"><i style="width:${100*a.value/max}%"></i></div></div>`).join('');
    $('invoice-count').textContent=state.invoices.length;
    $('payment-count').textContent=state.payments.length;
    $('tab-invoices').classList.toggle('active',view==='invoices');
    $('tab-payments').classList.toggle('active',view==='payments');
    $('tab-invoices').setAttribute('aria-pressed',view==='invoices');
    $('tab-payments').setAttribute('aria-pressed',view==='payments');
    $('overdue-filter').hidden=view!=='invoices';
    $('overdue-filter').setAttribute('aria-pressed',overdue);
    const query=$('search').value.trim().toLowerCase();
    let rows;
    if(view==='invoices') {
      rows=summary.rows.filter(r=>(!overdue||r.daysLate>0)&&[r.id,r.customer,r.description].join(' ').toLowerCase().includes(query)).reverse();
      $('register-caption').textContent='Invoice register';
      $('register-head').innerHTML='<tr><th scope="col">Invoice</th><th scope="col">Customer</th><th scope="col">Due date</th><th scope="col" class="number">Amount</th><th scope="col" class="number">Received</th><th scope="col" class="number">Outstanding</th><th scope="col">Status</th></tr>';
      $('register-body').innerHTML=rows.map(r=>`<tr class="${r.id===highlighted?'highlight':''}"><td>${esc(r.id)}</td><td>${esc(r.customer)}</td><td>${date(r.due)}</td><td class="number">${cash(r.amount)}</td><td class="number">${cash(r.received)}</td><td class="number">${r.included?cash(r.outstanding):'—'}</td><td><span class="status ${statusClass(r.status)}">${r.status}</span></td></tr>`).join('');
    } else {
      rows=state.payments.map(r=>({...r,customer:state.invoices.find(i=>i.id===r.invoiceId)?.customer||''})).filter(r=>[r.id,r.invoiceId,r.customer,r.reference].join(' ').toLowerCase().includes(query)).reverse();
      $('register-caption').textContent='Payment register';
      $('register-head').innerHTML='<tr><th scope="col">Payment</th><th scope="col">Invoice</th><th scope="col">Customer</th><th scope="col">Received</th><th scope="col">Method</th><th scope="col" class="number">Amount</th><th scope="col">Reference</th></tr>';
      $('register-body').innerHTML=rows.map(r=>`<tr><td>${esc(r.id)}</td><td>${esc(r.invoiceId)}</td><td>${esc(r.customer)}</td><td>${date(r.received)}${r.received>state.asOf?' *':''}</td><td>${esc(r.method)}</td><td class="number">${cash(r.amount)}</td><td>${esc(r.reference)}</td></tr>`).join('');
    }
    if(!rows.length)$('register-body').innerHTML='<tr><td colspan="7" class="empty-row">No matching records. Try a different search or filter.</td></tr>';
    $('row-count').textContent=`${rows.length} ${view==='invoices'?'invoice':'payment'}${rows.length===1?'':'s'} shown`;
    $('overpayment-note').textContent=view==='invoices'?`${cash(summary.overpayment)} in overpayments tracked separately`:'* After reporting date; excluded from totals';
  }
  function reset(announce=true) {
    state=E.clone(DEMO_SEED);view='invoices';overdue=false;highlighted='';tourId='';tourStep=-1;
    $('search').value='';$('tour-panel').hidden=true;lockTour(false);render();
    if(announce)message('Demo reset. The original fictional data is restored.');
  }
  const options=(items,value)=>items.map(item=>`<option value="${esc(typeof item==='string'?item:item.value)}" ${(typeof item==='string'?item:item.value)===value?'selected':''}>${esc(typeof item==='string'?item:item.label)}</option>`).join('');
  function field(name,label,type='text',value='',extra='',full=false) {return `<label${full?' class="full"':''}>${label}<input name="${name}" type="${type}" value="${esc(value)}" ${extra}></label>`;}
  function select(name,label,items,value='',full=false) {return `<label${full?' class="full"':''}>${label}<select name="${name}" required><option value="">Choose…</option>${options(items,value)}</select></label>`;}
  function openForm(kind) {
    mode=kind;
    $('form-error').textContent='';
    $('dialog-title').textContent=kind==='invoice'?'New invoice':'Record payment';
    $('submit-entry').textContent=kind==='invoice'?'Save invoice':'Save payment';
    $('dialog-intro').textContent='Required fields are marked *. An ID is assigned automatically when you save.';
    if(kind==='invoice') {
      const due=new Date(state.asOf+'T00:00:00Z');due.setUTCDate(due.getUTCDate()+30);
      $('form-fields').innerHTML=select('customer','Customer *',state.customers,'',true)+field('description','Description *','text','','required maxlength="200"',true)+field('issued','Issued date *','date',state.asOf,'required min="2000-01-01" max="2100-12-31"')+field('due','Due date *','date',due.toISOString().slice(0,10),'required min="2000-01-01" max="2100-12-31"')+field('amount','Amount (USD) *','number','','required min="0.01" max="999999999.99" step="0.01"')+select('owner','Owner *',state.owners)+field('po','Order / PO')+field('followUp','Follow-up date','date','','min="2000-01-01" max="2100-12-31"');
    } else {
      const invoices=state.invoices.slice().reverse().map(r=>({value:r.id,label:`${r.id} · ${r.customer}`}));
      $('form-fields').innerHTML=select('invoiceId','Invoice *',invoices,highlighted,true)+field('received','Received date *','date',state.asOf,'required min="2000-01-01" max="2100-12-31"')+select('method','Payment method *',E.methods,'ACH')+field('amount','Amount (USD) *','number','','required min="0.01" max="999999999.99" step="0.01"')+field('reference','Receipt reference *','text','','required maxlength="100"')+select('recordedBy','Recorded by *',state.owners,'',true);
    }
    $('form-fields').insertAdjacentHTML('beforeend','<label class="full">Notes<textarea name="notes" maxlength="500"></textarea></label>');
    $('entry-dialog').showModal();
  }
  $('entry-form').addEventListener('submit',event=>{
    event.preventDefault();
    const input=Object.fromEntries(new FormData(event.currentTarget));
    try {
      const record=mode==='invoice'?E.addInvoice(state,input):E.addPayment(state,input);
      highlighted=mode==='invoice'?record.id:record.invoiceId;view='invoices';overdue=false;$('search').value='';
      $('entry-dialog').close();render();
      message(`Saved ${record.id}${mode==='payment'?` against ${record.invoiceId}`:''}. The balances and aging have updated.`);
    } catch(error) {$('form-error').textContent=error.message;}
  });
  $('close-dialog').onclick=$('cancel-entry').onclick=()=>$('entry-dialog').close();
  $('new-invoice').onclick=()=>openForm('invoice');$('new-payment').onclick=()=>openForm('payment');
  $('reset-demo').onclick=()=>reset();
  $('search').oninput=render;
  $('tab-invoices').onclick=()=>{view='invoices';render();};$('tab-payments').onclick=()=>{view='payments';render();};
  $('overdue-filter').onclick=()=>{overdue=!overdue;render();};
  $('as-of').addEventListener('change',()=>{
    if(!E.validDate($('as-of').value)){message('Choose a valid reporting date between 2000 and 2100.');$('as-of').value=state.asOf;return;}
    state.asOf=$('as-of').value;render();message(`Balances now reflect activity through ${date(state.asOf)}.`);
  });
  function updateTour() {
    const steps=[
      ['1 OF 3 · CREATE','Start with a $1,200 invoice.','Add a fictional office-support invoice due October 15. The tracker assigns the next invoice ID and includes the amount in outstanding and past-due balances.','Add sample invoice'],
      ['2 OF 3 · RECORD','Apply a $400 partial payment.',`${tourId} has been added. Record a receipt against it; the original invoice stays in place while the remaining balance falls to $800.`,'Record $400 payment'],
      ['3 OF 3 · REVIEW','Find the remaining $800.',`The payment is linked to ${tourId}. Filter for overdue invoices to see the unpaid balance that still needs follow-up.`,'Show overdue invoices'],
      ['DEMO COMPLETE','One workflow. A traceable balance.','The invoice, receipt, and remaining balance are connected. You can continue experimenting with the forms, or reset the sample data.','Start again']
    ];
    const [step,title,copy,button]=steps[tourStep];
    $('tour-step').textContent=step;$('tour-title').textContent=title;$('tour-copy').textContent=copy;$('tour-next').textContent=button;
  }
  $('start-tour').onclick=()=>{reset(false);tourStep=0;lockTour(true);$('tour-panel').hidden=false;updateTour();message('Guided demo ready. It starts from the original sample data.');};
  $('tour-close').onclick=()=>{$('tour-panel').hidden=true;tourStep=-1;lockTour(false);};
  $('tour-next').onclick=()=>{
    if(tourStep===3){$('start-tour').click();return;}
    if(tourStep===0){const r=E.addInvoice(state,{customer:'Alder Office Co.',description:'Office support — guided demo',issued:'2026-10-01',due:'2026-10-15',amount:1200,owner:'Alex Morgan',po:'DEMO-1001'});tourId=r.id;highlighted=r.id;message(`Created ${r.id}. Outstanding and past-due balances increased by $1,200.`);}
    else if(tourStep===1){const r=E.addPayment(state,{invoiceId:tourId,received:'2026-10-20',method:'ACH',amount:400,reference:'DEMO-RCPT-001',recordedBy:'Alex Morgan'});message(`Recorded ${r.id}. ${tourId} now has an $800 outstanding balance.`);}
    else if(tourStep===2){overdue=true;view='invoices';$('search').value=tourId;message(`Showing ${tourId}: $1,200 invoiced, $400 received, $800 still overdue.`);}
    tourStep++;view='invoices';if(tourStep===3)lockTour(false);render();updateTour();
  };
  reset(false);
})();
