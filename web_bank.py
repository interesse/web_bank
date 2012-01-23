#!/usr/bin/python
# -*- coding: UTF-8 -*-
'''
Holt Kreditkarten-Umsätze per Web-Scraping vom Webfrontend der DKB 
(Deutsche Kreditbank), die für diese Daten kein HCBI anbietet.
Die Umsätze werden im Quicken Exchange Format (.qif) ausgegeben 
oder gespeichert und können somit in eine Buchhaltungssoftware 
importiert werden.

Geschrieben 2007 von Jens Herrmann <jens.herrmann@qoli.de> http://qoli.de
Benutzung des Programms ohne Gewähr. Speichern Sie nicht ihr Passwort
in der Kommandozeilen-History!

7.9.2008: Anstelle von xml.xpath, das von Ubuntu 8.04 nicht mehr unterstützt
wird, wird jetzt lxml.etree benutzt.

25.10.2008: Kleiner Fix wegen einer HTML-Änderung der DKB (table wird nicht mehr 
gebraucht), dafür wird jetzt exportiertes CSV ausgewertet, das nicht auf eine 
Seite Ausgabe beschränkt ist. Da kein HTML mehr geparsed wird, entfällt auch die
Abhängigkeit zu tidy und xpath.

27.12.2009: Auswahl der Kartennummer

23.11.2011: Fix des Session-Handlings durch eine Änderung der DKB

Benutzung: web_bank.py [OPTIONEN]

 -a, --account=ACCOUNT      Kontonummer des Hauptkontos. Angabe notwendig
 -c, --card=NUMBER          Die letzten 4 Stellen der Kartennummer, falls 
                            mehrere Karten vorhanden sind.
 -p, --password=PASSWORD    Passwort (Benutzung nicht empfohlen, 
                            geben Sie das Passwort ein, wenn Sie danach
                            gefragt werden)
 -f, --from=DD.MM.YYYY      Buchungen ab diesem Datum abfragen
 -t, --till=DD.MM.YYYY      Buchungen bis zu diesem Datum abfragen 
                            Default: Heute
 -o, --outfile=FILE         Dateiname für die Ausgabedatei
                            Default: Standardausgabe (Fenster)
 -v, --verbose              Gibt zusätzliche Debug-Informationen aus
'''

import sys, getopt
from datetime import datetime
from getpass import getpass
import urllib2, urllib, re

def group(lst, n):
	return zip(*[lst[i::n] for i in range(n)])
    
debug=False
def log(msg):
	if debug:
		print msg
    
# Parser for new style banking pages
class NewParser:
	URL = "https://banking.dkb.de"
	BETRAG = 'frmBuchungsbetrag'
	ZWECK = 'frmVerwendungszweck'
	TAG = 'frmBuchungstag'
	PLUSMINUS = 'frmSollHabenKennzeichen'
	MINUS_CHAR='S'
	DATUM = 'frmBelegdatum'
	
	def get_cc_index(self, card, data):
		log('Finde Kreditkartenindex für Karte ***%s...'%card)
		pattern= r'<option value="(.)" id=".*"( selected="selected" )?>.{12}%s / Kreditkarte'%card
		index= re.findall(pattern, data)
		if len(index)>0:
			return index[0][0]
		else:
			return '0'

	def get_cc_csv(self, account, card, password, fromdate, till):
		log('Hole sessionID und Token...')
		# retrieve sessionid and token
		url= self.URL+"/dkb/-?$javascript=disabled"
		page= urllib.urlopen(url,).read()
		session= re.findall(';jsessionid=.*?["\?]',page)[0][:-1]
		token= re.findall('<input type="hidden" name="token" value="(.*)" id=',page)[0]
		log('SessionID: %s Token: %s'%(session,token))
		# login
		url= self.URL+'/dkb/-'+session
		request=urllib2.Request(url, data= urllib.urlencode({
		                                                     '$$event_login.x': '0',
		                                                     '$$event_login.y': '0',
		                                                     'token': token,
		                                                     'j_username': account,
		                                                     'j_password': password,
		                                                     '$part': 'Welcome.login',
		                                                     '$$$event_login': 'login',
		}))
		page=urllib2.urlopen(request).read()

		# new sessionid after login
		session= re.findall(';jsessionid=.*?["\?]',page)[0][:-1]
		referer = url		
		url= self.URL+'/dkb/-'+session

		# init search
		request=urllib2.Request(url+'?$part=DkbTransactionBanking.content.creditcard.CreditcardTransactionSearch&$event=init',
		                        headers={'Referer':urllib.quote_plus(referer)})
		throwaway=urllib2.urlopen(request).read()
		referer = url

		# retrieve data
		request=urllib2.Request(url, data= urllib.urlencode({
		                                                     'slCreditCard': '0',
		                                                     'searchPeriod': '0',
		                                                     'postingDate': fromdate,
		                                                     'toPostingDate': till,
		                                                     '$$event_search': 'Umsätze+anzeigen',
		                                                     '$part': 'DkbTransactionBanking.content.creditcard.CreditcardTransactionSearch',
		                                                     '$$$event_search': 'search',
		}), headers={'Referer':urllib.quote_plus(referer)})
		data= ''.join(urllib2.urlopen(request).readlines())

		# find card index
		if not card=='':
		    	cc_index= self.get_cc_index(card, data)
		    	# again retrieve data for correct card
		    	request=urllib2.Request(url, data= urllib.urlencode({
		    	                                                     'slCreditCard': cc_index,
		    	                                                     'searchPeriod': '0',
		    	                                                     'postingDate': fromdate,
		    	                                                     'toPostingDate': till,
		    	                                                     '$$event_search': 'Umsätze+anzeigen',
		    	                                                     '$part': 'DkbTransactionBanking.content.creditcard.CreditcardTransactionSearch',
		    	                                                     '$$$event_search': 'search',
		    	}), headers={'Referer':urllib.quote_plus(url+"?$part=DkbTransactionBanking.content.banking.FinancialStatus.FinancialStatus&$event=paymentTransaction&row=1&table=cashTable")})
		    	throwaway= ''.join(urllib2.urlopen(request).readlines())



		#CSV abrufen
		request=urllib2.Request(url+'?$part=DkbTransactionBanking.content.creditcard.CreditcardTransactionSearch&$event=csvExport',
		                        headers={'Referer':urllib.quote_plus(url)})
		antwort= urllib2.urlopen(request).read()
		log('Daten empfangen. Länge: %s'%len(antwort))
		return antwort
	   
	def parse_csv(self, cc_csv):
		result=[]
		for line in cc_csv.split('\n')[8:]: # Liste beginnt in Zeile 9 des CSV
			g= line.split(';')
			if len(g)==7: #Jede Zeile hat 7 Elemente
				act={}
				act[self.ZWECK]=g[3][1:-1]
				act[self.TAG]=g[1][1:-1]
				act[self.DATUM]=g[2][1:-1]
				act[self.PLUSMINUS]=''
				act[self.BETRAG]=g[4][1:-1]
				result.append(act)
		return result

CC_NAME= 'VISA'
CC_NUMBER= ''
LOGIN_ACCOUNT=''
LOGIN_PASSWORD=''
PARSER= NewParser()

GUESSES=[
		(PARSER.BETRAG,'-150.0',u'Aktiva:Barvermögen:Bargeld'),
]

def guessCategories(f):
	for g in GUESSES:
		if g[1] in f[g[0]].upper():
			return g[2]

def render_qif(cc_data):
	cc_qif=[]
	cc_qif.append('!Account')
	cc_qif.append('N'+CC_NAME)
	cc_qif.append('^')
	cc_qif.append('!Type:Bank')
	log('Für Ausgabe vorbereiten:')
	for f in cc_data:
		log(str(f))
		if PARSER.TAG in f.keys():
			f[PARSER.BETRAG]= float(f[PARSER.BETRAG].replace('.','').replace(',','.'))
			if PARSER.MINUS_CHAR in f[PARSER.PLUSMINUS]:
				f[PARSER.BETRAG]= -f[PARSER.BETRAG]
			f[PARSER.BETRAG]=str(f[PARSER.BETRAG])
			datum=f[PARSER.DATUM].split('.')
			cc_qif.append('D'+datum[1]+'/'+datum[0]+'/'+datum[2])
			cc_qif.append('T'+f[PARSER.BETRAG])
			if PARSER.ZWECK+"1" in f:
				for n in range(1,8):
					if f[PARSER.ZWECK+str(n)].strip():
						cc_qif.append('M'+f[PARSER.ZWECK+str(n)])
			else:
				cc_qif.append('M'+f[PARSER.ZWECK])
			c= guessCategories(f)
			if c:
				cc_qif.append('L'+c)
			cc_qif.append('^')
	return u'\n'.join(cc_qif)
			
class Usage(Exception):
	def __init__(self, msg):
		self.msg = msg

def main(argv=None):
	account=LOGIN_ACCOUNT
	password= LOGIN_PASSWORD
	card_no= CC_NUMBER
	fromdate=''
	till=datetime.now().strftime('%d.%m.%Y')
	outfile= sys.stdout
	
	if argv is None:
		argv = sys.argv
	try:
		try:
			opts, args = getopt.getopt(argv[1:], "ha:c:p:f:t:o:v", ['help','account=','card=','password=','from=','till=','outfile=','verbose'])
		except getopt.error, msg:
			raise Usage(msg)
		for o, a in opts:
			if o in ("-h", "--help"):
				print __doc__
				return 0
			if o in ('-a','--account'):
				account= a
			if o in ('-c','--card'):
				card_no= a
			if o in ('-p','--password'):
				password= a
			if o in ('-f','--from'):
				fromdate= a
			if o in ('-t','--till'):
				till= a
			if o in ('-o','--outfile'):
				try:
					outfile=open(a,'w')
				except IOError, msg:
					raise Usage(msg)
			if o in ('-v','--verbose'):
				print 'Mit Debug-Ausgaben'
				global debug
				debug=True
		if not account or not fromdate:
			raise Usage('Anfangsdatum und Kontonummer müssen angegeben sein.')
		if not password:
			try:
				password=getpass('Geben Sie das Passwort für das Konto '+account+' ein: ')
			except KeyboardInterrupt:
				raise Usage('Sie müssen ein Passwort eingeben!')
			
		cc_csv = PARSER.get_cc_csv(account, card_no, password, fromdate, till)
		cc_data = PARSER.parse_csv(cc_csv)

		print >>outfile, render_qif(cc_data).encode('utf-8')
	 	
	except Usage, err:
		print >>sys.stderr, __doc__
		print >>sys.stderr, err.msg
		return 2
	
if __name__ == '__main__':
	sys.exit(main())

