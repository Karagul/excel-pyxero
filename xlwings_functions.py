import xlwings as xw


@xw.func
def pyCreateInvoice(ContactID,cYear,cMonth,cDay,dYear,dMonth,dDay,vbLineItems):
	from xero import Xero
	from xero.auth import PrivateCredentials
	from ast import literal_eval
	import datetime

	with open(r"PATH TO YOUR PRIVATEKEY.PEM REGISTERED ON XERO") as keyfile: 
		rsa_key = keyfile.read()
	credentials = PrivateCredentials("CONSUMER KEY HERE", rsa_key)
	xero = Xero(credentials)  

	evalLineItems = literal_eval(vbLineItems)

	original = {
	'Type': 'ACCREC',
	'Contact': {'ContactID': ContactID},
	'Date': datetime.date(cYear, cMonth, cDay),
	'DueDate': datetime.date(dYear, dMonth, dDay),
	'Status': 'DRAFT',
	'LineAmountTypes': 'Exclusive',
	'LineItems': evalLineItems,

	}

	pyInvoiceID = xero.invoices.put(original)[0]['InvoiceID']
	pyInvoiceNumber = xero.invoices.get(pyInvoiceID)[0]['InvoiceNumber']
	response = [pyInvoiceID,pyInvoiceNumber]
	return response

@xw.func
def pyAttachPDF(fName,fPath,invID):
	from xero import Xero
	from xero.auth import PrivateCredentials
	from ast import literal_eval
	from os import chdir
	import datetime
	

	with open(r"PATH TO YOUR PRIVATEKEY.PEM REGISTERED ON XERO") as keyfile: 
		rsa_key = keyfile.read()
	credentials = PrivateCredentials("CONSUMER KEY HERE", rsa_key)
	xero = Xero(credentials) 
	
	chdir(fPath)
	pyFile = open(fName, 'rb')
	xero.invoices.put_attachment(invID, fName, pyFile, 'application/pdf')
	pyFile.close()
