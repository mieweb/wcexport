#!/usr/bin/env python
import urllib2
import urllib
import base64
import json

URL = 'https://zeus-web.med-web.com/webchart/wctdrummondmast/webchart.cgi'
USERNAME = 'dave'
PASSWORD = 'dave'
COOKIE = None

if __name__ == '__main__':
	print('Initializing session')
	try:
		out = urllib2.urlopen(URL, urllib.urlencode({
			'login_user': USERNAME,
			'login_passwd': PASSWORD
		}))
		COOKIE = out.headers.get('Set-Cookie').split('=')[1].split(';')[0]
	except Exception as e:
		print('Session failed to initialize {0}'.format(e))

	if COOKIE:
		requests = {
			'Last Name LIKE "Hart"': 'GET/db/patients/LIKE_last_name=Hart',
			'Last Name LIKE "Pregnant"': 'GET/db/patients/LIKE_last_name=Pregnan',
		}
		for title, url in requests.iteritems():
			print('\nQuerying for patients: {0}'.format(title))
			js = json.load(
				urllib2.urlopen(URL, urllib.urlencode({
					'f': 'json',
					'session_id': COOKIE,
					'apistring': base64.b64encode(url)
				})))
			print(json.dumps(js))
	
