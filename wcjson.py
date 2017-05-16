#!/usr/bin/env python
import urllib2
import urllib
import base64
import json

if __name__ == '__main__':
	print('Showing patients with last name LIKE "Hart"')
	out = urllib2.urlopen('https://zeus-web.med-web.com/webchart/wctdrummondmast/webchart.cgi',
		urllib.urlencode({
			'f': 'json',
			'login_user': 'dave',
			'login_passwd': 'dave',
			'apistring': base64.b64encode('GET/db/patients/LIKE_last_name=Hart')
		}))
	js = json.load(out)
	print(json.dumps(js))
	
