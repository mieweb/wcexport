#!/usr/bin/env python
import urllib2
import urllib
import base64
import json
import os

URL = 'https://zeus-web.med-web.com/webchart/wctDrumCCDA/webchart.cgi'
USERNAME = 'dave'
PASSWORD = 'dave'
COOKIE = None


# Download a document
def downloadDocument(doc_id, filename):
    if not os.path.exists(filename):
        out = urllib2.urlopen(URL, urllib.urlencode({
            'f': 'stream',
            'doc_id': doc_id,
            'session_id': COOKIE,
            'rawdata': '1'
        }))
        with open(filename, 'wb') as fp:
            fp.write(out.read())

def downloadDocumentMeta(pat_id):
	try:
		api = "GET/db/documents/storage_type=19&LIKE_service_date=2017-05-02%25&pat_id=" + pat_id
		print('\nQuerying for patients: {0}'.format(pat_id))
		docs = json.load(
			urllib2.urlopen(URL, urllib.urlencode({
				'f': 'json',
				'session_id': COOKIE,
				'apistring': base64.b64encode(api)
			})))
		return docs["db"][0]["doc_id"];
	except:
		return ""


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

	print('Getting Patients')
	if COOKIE:
		requests = {
			'Last Name LIKE "Newman"': 'GET/db/patients/LIKE_last_name=Newman',
			'Last Name LIKE "Larson"': 'GET/db/patients/LIKE_last_name=Larson',
			'Last Name LIKE "Bates"': 'GET/db/patients/LIKE_last_name=Bates',
			'Last Name LIKE "Wright"': 'GET/db/patients/LIKE_last_name=Wright',
		}
		for title, url in requests.iteritems():
			print('\nQuerying for patients: {0}'.format(title))
			js = json.load(
				urllib2.urlopen(URL, urllib.urlencode({
					'f': 'json',
					'session_id': COOKIE,
					'apistring': base64.b64encode(url)
				})))
			pat_id = js["db"][0]["pat_id"]
			name = js["db"][0]["last_name"]


			print("Getting Documents for Patient:" + pat_id)
			doc_id = downloadDocumentMeta(pat_id);

			if doc_id != "":
				print("Downloading Document:" + doc_id)
				downloadDocument(doc_id, name + "_" + doc_id + ".xml")
			else:
				print("No documents exist for that patient that meet the criteria.")
#			print(json.dumps(js))
	
