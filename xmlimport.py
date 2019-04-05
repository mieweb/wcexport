#!/usr/bin/env python

import os
import sys
import requests
import getpass
import csv
import time
from StringIO import StringIO
from xml.dom import minidom

EXTENSIONS = ['.xml']
COOKIE = None
STORAGE_TYPE = 19
DOC_TYPE = 'WCCDA'

LOGFILE = 'import.log'

def usage():
    print('Usage: {0} WebChartUrl docpath mrcsvfile wcuser'.format(sys.argv[0]))

def getResponse(url, data={}):
    if data and COOKIE:
        data['session_id'] = COOKIE
    try:
        old = sys.stderr
        sys.stderr = open(os.devnull, 'w')
        res = requests.post(url, data=data, verify=False)
    except Exception as e:
        d = data.copy()
        if 'login_passwd' in d:
            d['login_passwd'] = 'XXXXX'
        raise Warning('Internal error in request [ {0} : {1} ] at [ {2} : {3} ]'.format(
            type(e), str(e), url, d))
    finally:
        sys.stderr = old
    if res.status_code not in [200]:
        raise Warning('Invalid http response code [ {0} ]'.format(res.headers))
    if res.headers.get('X-lg_status').lower() != 'success':
        raise Exception('Login failed [ {0} ]'.format(res.headers.get('X-status_desc')))
    out = res.text
    if [ord(x) for x in out[0:3]] == [239 ,187 ,191]:
        # Strip out utf-8 BOM from webchart CSV output
        out = out[3:]
    return out, res

def log(msg, echo):
    if echo:
        print(msg)
    with open(LOGFILE, 'a') as fp:
        fp.writelines('{0} | {1} \n'.format(time.strftime('%m/%d/%Y %H:%M:%S'), msg))

def getmrnumber(data, key, field):
    """ This method should be turned into a generic map function that can be customized
    per import method by the client
    """
    return data[key][field]

def unlinkLogFile():
    if os.path.exists(LOGFILE):
        idx = 0
        to = '{0}.bak.{1}'.format(LOGFILE, idx)
        while (os.path.exists(to)):
            idx += 1
            to = '{0}.bak.{1}'.format(LOGFILE, idx)
        os.rename(LOGFILE, to)

def loadCSV(path, key):
    """ Load the entire csv file into a single dict using the given key """
    data = {}
    reader = csv.DictReader(open(path, 'rb'))
    for row in reader:
        data[row[key]] = row
    return data

if __name__ == '__main__':
    if len(sys.argv) != 5:
        usage()
        exit(1)
    url = sys.argv[1]
    path = os.path.expanduser(sys.argv[2])
    mrcsvfile = os.path.expanduser(sys.argv[3])
    wcuser = sys.argv[4]
    unlinkLogFile()
    if not os.path.exists(path):
        print('{0} is not a valid path'.format(path))
        exit(1)
    files = [x for x in os.listdir(path) if os.path.isfile(os.path.join(path, x)) and os.path.splitext(x)[-1] in EXTENSIONS]
    if not files:
        print('No valid [ {0} ] files found in [ {1} ]'.format(','.join(EXTENSIONS), path))
        exit(1)
    if not os.path.exists(mrcsvfile):
        mrcsvfile = os.path.join(path, mrcsvfile)
        if not os.path.exists(mrcsvfile):
            print('{0} is not a valid absolute or relative path'.format(mrcsvfile))
            exit(1)
    data = loadCSV(mrcsvfile, 'document_name')
    wcpass = getpass.getpass('Please enter the webchart password for user [ {0} ]: => '.format(wcuser))
    try:
        out, res = getResponse(url, {'login_user': wcuser, 'login_passwd': wcpass})
        COOKIE = res.headers.get('Set-Cookie').split('=')[1].split(';')[0]
    except Exception as e:
        print('Failed to get the webchart cookie', e)
        exit(1)
    skipped = 0
    errors = 0
    completed = 0
    for idx, f in enumerate(files, start=1):
        try:
            mrnumber = 'MR-{0}'.format(getmrnumber(data, f, 'patient_id'))
        except KeyError as e:
            log('Skipped {0} => No patient_id found in csv file'.format(f), True)
            skipped += 1
            continue
        print('Importing file [ {0} / {1} ]  {2} => {3}'.format(idx, len(files), f, mrnumber))
        out, res = getResponse(url, {
            'f': 'chart',
            's': 'upload',
            'storage_type': STORAGE_TYPE,
            'doc_type': DOC_TYPE,
            'file': open(os.path.join(path, f), 'rb').read(),
            'pat_id': 18,
#            'mrnumber': mrnumber,
        })
        try:
            dom = minidom.parse(StringIO(out))
            error = str(dom.getElementsByTagName('ErrorCode')[0].firstChild.nodeValue)
            msg = str(dom.getElementsByTagName('Text')[0].firstChild.nodeValue)
            if error != '2048':    # Success error code from docupload.h
                log('Failed import {0} => {1} with error: {2}'.format(f, mrnumber, msg), True)
                errors += 1
            else:
                log('Success: {0} => {1} => {2}'.format(f, mrnumber, msg), False)
                completed += 1
        except Exception as e:
            print('Failed to parse webchart xml response, stopping the import', e)
            print(out)
            exit(1)
    print('\nDocument import process complete:\nUploaded: {0}\nSkipped: {1}\nErrors: {2}'.format(
        completed, skipped, errors))
    print('\nSee {0} for details'.format(LOGFILE))

