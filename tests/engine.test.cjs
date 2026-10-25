const test=require('node:test');
const assert=require('node:assert/strict');
const E=require('../assets/tracker-engine.js');
const seed=require('../assets/demo-data.js');
const fresh=()=>E.clone(seed);
const invoice={customer:seed.customers[0],description:'Test support',issued:'2026-10-01',due:'2026-10-15',amount:1200,owner:'Alex Morgan'};
const receipt=(invoiceId,amount=400)=>({invoiceId,received:'2026-10-20',method:'ACH',amount,reference:'TEST-001',recordedBy:'Alex Morgan'});
test('published sample reconciles to the Excel dashboard',()=>{
  const s=E.summarize(seed);
  assert.equal(s.invoiced,6490000);assert.equal(s.received,2550000);assert.equal(s.outstanding,3945000);assert.equal(s.pastDue,1870000);assert.equal(s.overpayment,5000);
  assert.equal(s.invoiced-s.received+s.overpayment,s.outstanding);
  assert.equal(s.aging.reduce((n,b)=>n+b.value,0),s.outstanding);
});
test('guided workflow creates a unique ID and changes overdue totals',()=>{
  const state=fresh(),base=E.summarize(state),r=E.addInvoice(state,invoice),s=E.summarize(state);
  assert.equal(r.id,'INV-2026-026');assert.equal(s.outstanding,base.outstanding+120000);assert.equal(s.pastDue,base.pastDue+120000);
});
test('partial payments preserve invoice amount and reduce remaining balance',()=>{
  const state=fresh(),r=E.addInvoice(state,invoice);E.addPayment(state,receipt(r.id));const row=E.summarize(state).rows.find(i=>i.id===r.id);
  assert.equal(row.amount,120000);assert.equal(row.received,40000);assert.equal(row.outstanding,80000);assert.equal(row.status,'Overdue');
});
test('overpayment is tracked separately without a negative balance',()=>{
  const state=fresh(),r=E.addInvoice(state,invoice);E.addPayment(state,receipt(r.id,1250));const row=E.summarize(state).rows.find(i=>i.id===r.id);
  assert.equal(row.outstanding,0);assert.equal(row.overpayment,5000);assert.equal(row.status,'Overpaid');
});
test('full payment marks an invoice paid',()=>{const s=fresh(),r=E.addInvoice(s,invoice);E.addPayment(s,receipt(r.id,1200));assert.equal(E.summarize(s).rows.find(i=>i.id===r.id).status,'Paid');});
test('future receipt is excluded until reporting date reaches it',()=>{
  const s=fresh(),r=E.addInvoice(s,invoice);E.addPayment(s,{...receipt(r.id),received:'2026-11-01'});
  assert.equal(E.summarize(s).rows.find(i=>i.id===r.id).received,0);assert.equal(E.summarize(s,'2026-11-01').rows.find(i=>i.id===r.id).received,40000);
});
test('not-yet-issued invoices do not enter dashboard totals',()=>{const s=fresh(),base=E.summarize(s);E.addInvoice(s,{...invoice,issued:'2026-11-01',due:'2026-12-01'});assert.equal(E.summarize(s).invoiced,base.invoiced);assert.equal(E.summarize(s).rows.at(-1).status,'Not issued');});
test('due today is not overdue; the following day is',()=>{const s=fresh(),r=E.addInvoice(s,{...invoice,due:seed.asOf});assert.equal(E.summarize(s).rows.find(i=>i.id===r.id).status,'Due today');assert.equal(E.summarize(s,'2026-10-26').rows.find(i=>i.id===r.id).daysLate,1);});
test('all aging boundaries use the correct bucket',()=>{
  for(const [days,index]of [[0,0],[1,1],[30,1],[31,2],[60,2],[61,3],[90,3],[91,4]]){const due=new Date(seed.asOf+'T00:00:00Z');due.setUTCDate(due.getUTCDate()-days);const s=fresh();s.invoices=[];s.payments=[];E.addInvoice(s,{...invoice,issued:'2026-01-01',due:due.toISOString().slice(0,10)});assert.equal(E.summarize(s).aging[index].value,120000);}
});
test('invalid amounts never create an invoice',()=>{for(const amount of [0,-1,'',null,NaN,Infinity,'abc',.001,1e9]){const s=fresh();assert.throws(()=>E.addInvoice(s,{...invoice,amount}));assert.equal(s.invoices.length,seed.invoices.length);}});
test('invalid dates and reversed chronology are rejected',()=>{for(const due of ['2026-02-30','bad','2026-09-30','1999-12-31'])assert.throws(()=>E.addInvoice(fresh(),{...invoice,due}));assert.equal(E.validDate('2024-02-29'),true);assert.equal(E.validDate('2026-02-29'),false);});
test('unmatched invoices and early receipts are rejected',()=>{const s=fresh(),r=E.addInvoice(s,invoice),count=s.payments.length;assert.throws(()=>E.addPayment(s,receipt('MISSING')));assert.throws(()=>E.addPayment(s,{...receipt(r.id),received:'2026-09-30'}));assert.equal(s.payments.length,count);});
test('required fields reject whitespace and unknown customers',()=>{assert.throws(()=>E.addInvoice(fresh(),{...invoice,description:' '}));assert.throws(()=>E.addInvoice(fresh(),{...invoice,owner:' '}));assert.throws(()=>E.addInvoice(fresh(),{...invoice,customer:'Unknown'}));});
test('receipt reference and payment method are validated',()=>{const s=fresh(),r=E.addInvoice(s,invoice);assert.throws(()=>E.addPayment(s,{...receipt(r.id),reference:''}));assert.throws(()=>E.addPayment(s,{...receipt(r.id),method:'Other'}));});
test('IDs increase across gaps and years without duplication',()=>{assert.equal(E.nextId([{id:'INV-2026-001'},{id:'INV-2026-009'},{id:'INV-2025-999'}],'INV-2026-'),'INV-2026-010');});
test('cent arithmetic preserves a small decimal payment',()=>{const s=fresh(),r=E.addInvoice(s,{...invoice,amount:.30});E.addPayment(s,receipt(r.id,.10));assert.equal(E.summarize(s).rows.find(i=>i.id===r.id).outstanding,20);});
test('cloning and resetting a demo does not mutate the baseline',()=>{const s=fresh();E.addInvoice(s,invoice);assert.equal(seed.invoices.length,25);assert.equal(fresh().invoices.length,25);});
