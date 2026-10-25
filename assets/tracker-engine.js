/* Business rules for the browser preview. VBA remains in the downloadable workbook. */
(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.TrackerEngine = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  'use strict';
  const DAY = 86400000;
  const methods = ['ACH', 'Bank transfer', 'Card', 'Check', 'Cash'];
  const clone = value => JSON.parse(JSON.stringify(value));
  const dateNumber = value => Date.parse(value + 'T00:00:00Z');
  function validDate(value) {
    return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value) && value >= '2000-01-01' && value <= '2100-12-31' && Number.isFinite(dateNumber(value)) && new Date(dateNumber(value)).toISOString().slice(0,10) === value;
  }
  function cents(value) {
    if (value === '' || value === null || typeof value === 'boolean') throw new Error('Enter an amount greater than zero.');
    const number = Number(value);
    if (!Number.isFinite(number) || number < 0.01 || number > 999999999.99) throw new Error('Enter an amount between $0.01 and $999,999,999.99.');
    return Math.round((number + Number.EPSILON) * 100);
  }
  function nextId(records, prefix) {
    const numbers = records.filter(r => r.id.startsWith(prefix)).map(r => r.id.slice(prefix.length)).filter(s => /^\d+$/.test(s)).map(Number);
    return prefix + String(Math.max(0, ...numbers) + 1).padStart(3,'0');
  }
  function required(value, message) {
    const text = String(value ?? '').trim();
    if (!text) throw new Error(message);
    return text;
  }
  function addInvoice(state, input) {
    const customer = required(input.customer, 'Choose a customer.');
    if (!state.customers.includes(customer)) throw new Error('Choose a customer from the demo list.');
    const description = required(input.description, 'Enter a description.');
    if (!validDate(input.issued) || !validDate(input.due)) throw new Error('Enter valid issued and due dates between 2000 and 2100.');
    if (input.due < input.issued) throw new Error('The due date must be on or after the issued date.');
    if (input.followUp && !validDate(input.followUp)) throw new Error('Enter a valid follow-up date.');
    const amount = cents(input.amount);
    const owner = required(input.owner, 'Enter an owner.');
    const record = {id:nextId(state.invoices, `INV-${input.issued.slice(0,4)}-`), customer, description, issued:input.issued, due:input.due, amount, owner, po:String(input.po || '').trim(), followUp:input.followUp || '', notes:String(input.notes || '').trim()};
    state.invoices.push(record);
    return record;
  }
  function addPayment(state, input) {
    const invoice = state.invoices.find(r => r.id === input.invoiceId);
    if (!invoice) throw new Error('Choose an existing invoice.');
    if (!validDate(input.received)) throw new Error('Enter a valid received date between 2000 and 2100.');
    if (input.received < invoice.issued) throw new Error('The payment date must be on or after the invoice issued date.');
    if (!methods.includes(input.method)) throw new Error('Choose a payment method.');
    const amount = cents(input.amount);
    const reference = required(input.reference, 'Enter a receipt reference.');
    const recordedBy = required(input.recordedBy, 'Enter who recorded the payment.');
    const record = {id:nextId(state.payments,'PAY-'), invoiceId:invoice.id, received:input.received, method:input.method, amount, reference, recordedBy, notes:String(input.notes || '').trim()};
    state.payments.push(record);
    return record;
  }
  function summarize(state, asOf = state.asOf) {
    if (!validDate(asOf)) throw new Error('Choose a valid reporting date.');
    const rows = state.invoices.map(invoice => {
      const included = invoice.issued <= asOf;
      const received = included ? state.payments.filter(p => p.invoiceId === invoice.id && p.received <= asOf).reduce((sum,p) => sum+p.amount,0) : 0;
      const outstanding = included ? Math.max(invoice.amount-received,0) : 0;
      const overpayment = included ? Math.max(received-invoice.amount,0) : 0;
      const daysLate = outstanding ? Math.max(Math.floor((dateNumber(asOf)-dateNumber(invoice.due))/DAY),0) : 0;
      const status = !included ? 'Not issued' : overpayment ? 'Overpaid' : !outstanding ? 'Paid' : daysLate ? 'Overdue' : invoice.due === asOf ? 'Due today' : received ? 'Part paid' : 'Open';
      return {...invoice, included, received, outstanding, overpayment, daysLate, status};
    });
    const aging = [{label:'Not overdue',value:0},{label:'1–30 days',value:0},{label:'31–60 days',value:0},{label:'61–90 days',value:0},{label:'91+ days',value:0}];
    rows.forEach(row => {if (row.outstanding) aging[row.daysLate===0?0:row.daysLate<=30?1:row.daysLate<=60?2:row.daysLate<=90?3:4].value+=row.outstanding;});
    return {rows,aging,invoiced:rows.filter(r=>r.included).reduce((s,r)=>s+r.amount,0),received:rows.reduce((s,r)=>s+r.received,0),outstanding:rows.reduce((s,r)=>s+r.outstanding,0),pastDue:rows.filter(r=>r.daysLate>0).reduce((s,r)=>s+r.outstanding,0),overpayment:rows.reduce((s,r)=>s+r.overpayment,0),open:rows.filter(r=>r.outstanding>0).length};
  }
  return {clone,validDate,cents,nextId,addInvoice,addPayment,summarize,methods};
});
