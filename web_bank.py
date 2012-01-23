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

Benutzung: web_bank.py [OPTIONEN]

 -a, --account=ACCOUNT      Kontonummer des Hauptkontos. Angabe notwendig
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
import tidy
from xml.dom import minidom
from xml import xpath

def group(lst, n):
    return zip(*[lst[i::n] for i in range(n)])
    
debug=False
def log(msg):
	if debug:
		print msg
    
# Parser for new style banking pages
class NewParser:
	URL = "https://banking.dkb.de/dkb/-"
	BETRAG = 'frmBuchungsbetrag'
	ZWECK = 'frmVerwendungszweck'
	TAG = 'frmBuchungstag'
	PLUSMINUS = 'frmSollHabenKennzeichen'
	MINUS_CHAR='S'
	DATUM = 'frmBelegdatum'

	def get_cc_html(self, account, password, fromdate, till):
	    log('Hole sessionID und Token...')
	    # retrieve sessionid and token
	    url= self.URL+"?$javascript=disabled"
	    page= urllib.urlopen(url,).read()
	    session= re.findall(';jsessionid=.*?["\?]',page)[0][:-1]
	    token= re.findall('<input type="hidden" name="token" value="(.*)" id=',page)[0]
	    log('SessionID: %s Token: %s'%(session,token))
	    # login
	    url= self.URL+session
	    request=urllib2.Request(url, data= urllib.urlencode({
															'$$event_login.x': '0',
															'$$event_login.y': '0',
															'token': token,
															'j_username': account,
															'j_password': password,
															'$part': 'Welcome.login',
															'$$$event_login': 'login',
	    }))
	    throwaway=urllib2.urlopen(request).read()

	    # Call page for chosing data
	    log('Hole Table...')
	    throwaway=urllib.urlopen(url+'?$part=DkbTransactionBanking.index.menu&treeAction=selectNode&node=2.1&tree=menu').read()
	    table= re.findall('<input type="hidden" name="table" value="([^"]*)"',throwaway)[0]
	    log('Table: %s'%table)

		# retrieve data
	    request=urllib2.Request(url, data= urllib.urlencode({
															'slCreditCard': '0',
															'searchPeriod': '0',
															'postingDate': fromdate,
															'toPostingDate': till,
															'$$event_search': 'Umsätze+anzeigen',
															'table': table,
															'$part': 'DkbTransactionBanking.content.creditcard.CreditcardTransactionSearch',
															'$$$Sevent_search': 'search',
	    }), headers={'Referer':urllib.quote_plus(url+"?$part=DkbTransactionBanking.content.banking.FinancialStatus.FinancialStatus&$event=paymentTransaction&row=1&table=cashTable")})
	    
	    antwort= ''.join(urllib2.urlopen(request).readlines())
	    log('Daten empfangen. Länge: %s'%len(antwort))
	    return antwort
	   
	def parse_html(self, cc_html):
		options = dict(output_xhtml=1,add_xml_decl=1,indent=0,tidy_mark=0)
		outstr= tidy.parseString(cc_html, **options)
		outstr= str(outstr).replace("&nbsp;"," ")
		doc= minidom.parseString(outstr)
		n=xpath.Evaluate('//td[@headers]//text()',doc.documentElement)
		groups= group([node.data.strip() for node in n],11)
		result=[]
		log('Daten enthalten %s Einträge'%len(groups))
		for g in groups:
			act={}
			act[self.ZWECK]=g[3].replace("\n"," ")
			act[self.TAG]=g[1]
			act[self.DATUM]=g[2]
			act[self.PLUSMINUS]=g[5][-1]
			act[self.BETRAG]=g[5].replace("\n"," ").split(" ")[0]
			result.append(act)
			act=None
		log('%s Einträge verarbeitet'%len(result))
		return result

CC_NAME= 'VISA'
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
	fromdate=''
	till=datetime.now().strftime('%d.%m.%Y')
	outfile= sys.stdout
	
	if argv is None:
		argv = sys.argv
	try:
		try:
			opts, args = getopt.getopt(argv[1:], "ha:p:f:t:o:v", ['help','account=','password=','from=','till=','outfile=','verbose'])
		except getopt.error, msg:
			raise Usage(msg)
		for o, a in opts:
			if o in ("-h", "--help"):
				print __doc__
				return 0
			if o in ('-a','--account'):
				account= a
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
			
		cc_html = PARSER.get_cc_html(account, password, fromdate, till)
		cc_data = PARSER.parse_html(cc_html)

		print >>outfile, render_qif(cc_data).encode('utf-8')
	 	
	except Usage, err:
	    print >>sys.stderr, __doc__
	    print >>sys.stderr, err.msg
	    return 2
	
if __name__ == '__main__':
    sys.exit(main())

