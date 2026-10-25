"""Run the embedded VBA in pyOpenVBA's offline interpreter; never resave the workbook.
Install the optional pinned dependency: python -m pip install pyopenvba==6.5.0
Run from the repository root: python tests/workbook_smoke.py
This is not native Microsoft Excel execution.
"""
from pathlib import Path
from decimal import Decimal
import json
from pyopenvba import ExcelFile
from pyopenvba.apps.excel import ExcelApplication

root=Path(__file__).resolve().parents[1]
path=root/'downloads/invoice-payment-tracker.xlsm'
app=ExcelApplication.open(path)
checks=[]
def check(name,condition):
    assert condition,(name,app.dialogs[-1:] if app.dialogs else [])
    checks.append(name)
def run(macro):
    before=len(app.dialogs)
    app.run(macro)
    assert not any('could not' in d[1].lower() for d in app.dialogs[before:]),app.dialogs[before:]
def fill(values):
    for cell,value in values.items():app.sheet('Entry').set_value(cell,value)
with ExcelFile(path) as book:
    for module,extension in [('ModTracker','bas'),('ModFormulas','bas'),('ThisWorkbook','cls')]:
        normalize=lambda source:source.replace('\r\n','\n').strip()
        check(module+' source matches workbook',normalize(book.get_module(module))==normalize((root/f'src/vba/{module}.{extension}').read_text()))
run('ThisWorkbook.Workbook_Open');check('Workbook open selects cover',app.workbook.active_sheet.name=='Cover')
run('GoOverview');check('Dashboard navigation',app.workbook.active_sheet.name=='Overview')
before=app.sheet('Overview').value('H7')
run('NewInvoice');check('Invoice form navigation',app.workbook.active_sheet.name=='Entry')
invoice_rows=app.evaluate('ThisWorkbook.Worksheets("Invoices").ListObjects("InvoicesTable").ListRows.Count')
run('SaveInvoice');check('Blank form rejected',app.evaluate('ThisWorkbook.Worksheets("Invoices").ListObjects("InvoicesTable").ListRows.Count')==invoice_rows)
fill({'C10':'Alder Office Co.','C14':'Offline workflow test','C16':'10/01/2026','C18':'10/15/2026','C20':1200,'C22':'Alex Morgan'})
run('SaveInvoice')
check('New invoice ID generated',app.sheet('Invoices').value('A34')=='INV-2026-026')
check('Outstanding updated',app.sheet('Overview').value('H7')==before+1200)
check('Form cleared after success',app.sheet('Entry').value('C10') is None)
fill({'I10':'INV-2026-026','I12':'10/20/2026','I14':'ACH','I16':400,'I18':'OFFLINE-TEST-001','I20':'Alex Morgan'})
run('SavePayment')
check('Payment ID generated',app.sheet('Payments').value('A24')=='PAY-016')
check('Partial payment applied',app.sheet('Invoices').value('I34')==800)
fill({'I10':'INV-2026-026','I12':'10/20/2026','I14':'ACH','I16':900,'I18':'OFFLINE-TEST-002','I20':'Alex Morgan'})
run('SavePayment')
check('Overpayment tracked separately',app.sheet('Invoices').value('J34')==100 and app.sheet('Invoices').value('I34')==0)
app.sheet('Invoices').set_value('I34',999)
run('RefreshTracker');check('Formula restored',app.sheet('Invoices').value('I34')==0)
run('ShowOverdue');run('ShowAllInvoices')
check('Event state restored',app.evaluate('Application.EnableEvents') is True)
print(json.dumps({'engine':'pyOpenVBA 6.5.0 offline interpreter','nativeExcel':False,'passed':len(checks),'checks':checks},indent=2))
