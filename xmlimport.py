#!/usr/bin/env python

import os
import sys
import requests
import getpass

EXTENSIONS = ['.xml']
COOKIE = None

def usage():
    print('Usage: {0} WebChartUrl DocPath wcuser'.format(sys.argv[0]))

def getResponse(url, data={}):
    if data and COOKIE:
        data['session_id'] = COOKIE
    try:
        res = requests.post(url, data=data, verify=False)
    except Exception as e:
        raise Warning('Internal error in request [ {0} : {1} ] at [ {2} : {3} ]'.format(
            type(e), str(e), url, data))
    if res.status_code not in [200]:
        raise Warning('Invalid http response code [ {0} ]'.format(res.headers))
    if res.headers.get('X-lg_status').lower() != 'success':
        raise Exception('Login failed [ {0} ]'.format(res.headers.get('X-status_desc')))
    out = res.text
    if [ord(x) for x in out[0:3]] == [239 ,187 ,191]:
        # Strip out utf-8 BOM from webchart CSV output
        out = out[3:]
    return out, res

if __name__ == '__main__':
    if len(sys.argv) != 4:
        usage()
        exit(1)
    url = sys.argv[1]
    path = sys.argv[2]
    wcuser = sys.argv[3]
    if not os.path.exists(path):
        print('{0} is not a valid path'.format(path))
        exit(1)
    files = [x for x in os.listdir(path) if os.path.isfile(os.path.join(path, x)) and os.path.splitext(x)[-1] in EXTENSIONS]
    if not files:
        print('No valid [ {0} ] files found in [ {1} ]'.format(','.join(EXTENSIONS), path))
        exit(1)
    wcpass = getpass.getpass('Please enter the webchart password for user [ {0} ]: => '.format(wcuser))
    try:
        out, res = getResponse(url, {'login_user': wcuser, 'login_passwd': wcpass})
        COOKIE = res.headers.get('Set-Cookie').split('=')[1].split(';')[0]
    except Exception as e:
        print('Failed to get the webchart cookie', e)
        exit(1)
    for idx, f in enumerate(files, start=1):
        print('Importing {0} / {1} files'.format(idx, len(files)))
        out, res = getResponse(url, {
            'f': 'chart',
            's': 'upload',
            'storage_type': 19,
            'doc_type': 'WCCDA',
            'file': open(os.path.join(path, f), 'rb').read(),
            'Submit File': 1,
            'register_patient': 1,
            'mrnumber': 'MR-12345',
        })
        #print(out)
    print('Document import complete')
