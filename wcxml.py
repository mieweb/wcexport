#!/usr/bin/env python
import sys
import os
import urllib2
import urllib
import base64
import json
import re

URL = 'https://zeus-web.med-web.com/webchart/wctdcarlson_xdr/webchart.cgi'
USERNAME = 'selenium'
PASSWORD = 'Selenium1!'
OUTPUT_FOLDER = 'output'

SDATE = ''
EDATE = ''
NAMES = [ 'Newman' ]


# this will get set as part of the API
SESSION_COOKIE = None

DTREG = '\d{4}-\d{2}-\d{2}'
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
	'Care Team Members': 'Care Team Members',
	'Medical_Equipment': 'Medical_Equipment',
	'Goals': 'Goals',								# Encounter
	'Health Concerns': 'Health Concerns',			# Encounter
	'Assessments': 'Assessments',					# Encounter
	'Health Concerns': 'Health Concerns',			# Encounter
	'Plan': 'Plan'									# Encounter
}


if __name__ == '__main__':
	dtmatches = [x for x in NAMES if re.match(DTREG, x)]

	print('Initializing session at {0}'.format(URL))
	try:
		out = urllib2.urlopen(URL, urllib.urlencode({
			'login_user': USERNAME,
			'login_passwd': PASSWORD
		}))
		SESSION_COOKIE = out.headers.get('Set-Cookie').split('=')[1].split(';')[0]
	except Exception as e:
		print('Session failed to initialize {0}'.format(e))

	if SESSION_COOKIE:
		print('Getting Chart identifiers')
		charts = {}
		for name in NAMES:
			js = json.load(urllib2.urlopen(URL, urllib.urlencode({
				'f': 'json',
				'session_id': SESSION_COOKIE,
				'apistring': base64.b64encode('GET/db/patients/LIKE_last_name={0}'.format(name))
			})))
			if js and js['db']:
				for rec in js['db']:
					charts[rec['pat_id']] = rec


		for cid, chart in charts.iteritems():
			patname = '{0},{1},{2}_{3}'.format(chart['last_name'], chart['first_name'],
				chart['middle_name'], cid)
			if not os.path.exists(os.path.join(OUTPUT_FOLDER, patname)):
				os.makedirs(os.path.join(OUTPUT_FOLDER, patname))
			print('Retrieving data for {0} {1} {2}'.format(patname,
				'after' if not EDATE and SDATE else 'between' if EDATE and SDATE else '',
				SDATE if not EDATE else '{0} and {1}'.format(SDATE, EDATE)))
			for k, v in APIS.iteritems():
				res = urllib2.urlopen(URL, urllib.urlencode({
					'session_id': SESSION_COOKIE,
					'f': 'layout',
					'module': 'StructDocAPI',
					'XML': '1',
					'name': v,
					'pat_id': cid,
#					'encounter_id': 76,  # encounter is needed for some of the calls 
					'SDATE': SDATE,
					'EDATE': EDATE,
				}))
				with open (os.path.join(OUTPUT_FOLDER, patname, '{0}.xml'.format(k)), 'w') as fp:
					fp.write(res.read())
			
