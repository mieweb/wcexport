#!/usr/bin/env python
import sys
import os
import urllib2
import urllib
import base64
import json
import re

USERNAME = 'dave'
PASSWORD = 'dave'
COOKIE = None

DTREG = '\d{4}-\d{2}-\d{2}'
OUTPUT = 'output'
APIS = {
	'Patient Name': 'Patient Name',
	'Sex': 'Gender Code',
	'Date of Birth': 'Birth Date',
	'Race': 'Patient Race',
	'Ethnicity': 'Patient Ethnicity',
	'Preferred Language': 'Patient Language',
	'Smoking Status': 'Smoking Status',
	'Problems': 'Problems',
	'Medications': 'Medications',
	'Medication Allergies': 'Allergies',
	'Lab Values_Result': 'Results',
	'Vital Signs': 'Vital Signs',
	'Procedures': 'Procedures',
	'Immunizations': 'Immunizations',
}

def usage():
	print('Usage: {0} URL [startDate [endDate]] PatientLastName1 PatientLastName2 ...'.format(__file__))
	exit()

if __name__ == '__main__':
	if len(sys.argv) < 3:
		usage()
	URL = sys.argv[1]
	sdate = ''
	edate = ''
	names = sys.argv[2:]
	dtmatches = [x for x in names if re.match(DTREG, x)]
	if dtmatches:
		if len(dtmatches) == 1:
			sdate = dtmatches[0]
		else:
			sdate = dtmatches[0]
			edate = dtmatches[1]
	charts = {}
	print('Initializing session at {0}'.format(URL))
	try:
		out = urllib2.urlopen(URL, urllib.urlencode({
			'login_user': USERNAME,
			'login_passwd': PASSWORD
		}))
		COOKIE = out.headers.get('Set-Cookie').split('=')[1].split(';')[0]
	except Exception as e:
		print('Session failed to initialize {0}'.format(e))

	if COOKIE:
		for name in names:
			js = json.load(urllib2.urlopen(URL, urllib.urlencode({
				'f': 'json',
				'session_id': COOKIE,
				'apistring': base64.b64encode('GET/db/patients/LIKE_last_name={0}'.format(name))
			})))
			if js and js['db']:
				for rec in js['db']:
					charts[rec['pat_id']] = rec
		for cid, chart in charts.iteritems():
			patname = '{0},{1},{2}_{3}'.format(chart['last_name'], chart['first_name'],
				chart['middle_name'], cid)
			if not os.path.exists(os.path.join(OUTPUT, patname)):
				os.makedirs(os.path.join(OUTPUT, patname))
			print('Retrieving data for {0} {1} {2}'.format(patname,
				'after' if not edate and sdate else 'between' if edate and sdate else '',
				sdate if not edate else '{0} and {1}'.format(sdate, edate)))
			for k, v in APIS.iteritems():
				res = urllib2.urlopen(URL, urllib.urlencode({
					'session_id': COOKIE,
					'f': 'layout',
					'module': 'StructDocAPI',
					'XML': '1',
					'name': v,
					'pat_id': cid,
					'sdate': sdate,
					'edate': edate,
				}))
				with open (os.path.join(OUTPUT, patname, '{0}.xml'.format(k)), 'w') as fp:
					fp.write(res.read())
			
