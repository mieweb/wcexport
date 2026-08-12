#!/usr/bin/env python3
"""Headless reproduction of the wcexport login + "Appliance Synchronization" permission check.

Mirrors the request sequence in common.MainWin.validateCredentials() without tkinter, so the
failure can be reproduced and inspected on any platform.

Credentials are read from the environment (never pass them on the command line, where they
would land in shell history and process listings):

    WC_URL   e.g. https://example.webchartnow.com/webchart.cgi
    WC_USER  WebChart username
    WC_PASS  WebChart password  (omit to be prompted)

Usage:
    WC_URL=... WC_USER=... python3 repro_permission.py
    python3 repro_permission.py --category "Appliance Synchronization" --save-body out.raw
"""
import argparse
import getpass
import http.cookies
import os
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
from io import StringIO
from xml.dom import minidom

BODY_PREVIEW_LIMIT = 4000
SENSITIVE_PARAMS = ('login_passwd', 'password', 'passwd', 'pass')


def getSSLContext():
    """Match the (permissive) TLS behaviour of common.getSSLContext()."""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def maskParams(data):
    masked = {}
    for key, value in data.items():
        if key in SENSITIVE_PARAMS:
            masked[key] = '<redacted len={0}>'.format(len(str(value)))
        elif key == 'session_id':
            text = str(value)
            masked[key] = '{0}...(len={1})'.format(text[:4], len(text))
        else:
            masked[key] = value
    return masked


def section(title):
    print('\n' + '=' * 72)
    print(title)
    print('=' * 72)


def request(url, params, label):
    """POST form-encoded params and dump the full exchange. Returns (body_bytes, response)."""
    section(label)
    print('URL:    {0}'.format(url))
    print('Params: {0}'.format(maskParams(params)))
    encoded = urllib.parse.urlencode(params, doseq=True).encode('utf-8')
    try:
        res = urllib.request.urlopen(url, context=getSSLContext(), data=encoded)
    except urllib.error.HTTPError as e:
        print('HTTPError {0}: {1}'.format(e.code, e.reason))
        body = e.read()
        print('Body:\n{0}'.format(body.decode('utf-8', errors='replace')[:BODY_PREVIEW_LIMIT]))
        return body, e
    except Exception as e:
        print('{0}: {1}'.format(type(e).__name__, e))
        return None, None

    body = res.read()
    if body[:3] == b'\xef\xbb\xbf':
        print('Note: UTF-8 BOM present, stripping')
        body = body[3:]

    print('Status: {0} {1}'.format(res.getcode(), res.reason))
    print('Final URL: {0}'.format(res.geturl()))
    print('Response headers:')
    for name, value in res.headers.items():
        if name.lower() == 'set-cookie':
            print('  {0}: <redacted len={1}>'.format(name, len(value)))
        else:
            print('  {0}: {1}'.format(name, value))
    print('Body length: {0} bytes'.format(len(body)))
    print('Body (first {0} chars):'.format(BODY_PREVIEW_LIMIT))
    print(body.decode('utf-8', errors='replace')[:BODY_PREVIEW_LIMIT])
    return body, res


def parseSessionId(setCookieHeader):
    """Compare the app's naive parse against a correct cookie parse."""
    section('Session cookie parsing')
    print('Raw Set-Cookie: {0!r}'.format(setCookieHeader))
    if not setCookieHeader:
        print('No Set-Cookie header -> login did not establish a session.')
        return None

    # How common.py does it today: split on EVERY '=' and take element [1].
    naive = None
    parts = setCookieHeader.split('=')
    if len(parts) > 1:
        naive = parts[1].split(';')[0]
    print('wcexport parse   -> {0!r}'.format(naive))

    # How it should be done.
    jar = http.cookies.SimpleCookie()
    try:
        jar.load(setCookieHeader)
    except http.cookies.CookieError as e:
        print('CookieError: {0}'.format(e))
        return naive
    correct = {name: morsel.value for name, morsel in jar.items()}
    print('Correct parse    -> {0}'.format(correct))
    if naive is not None and naive not in correct.values():
        print('MISMATCH: the naive parse did not produce any real cookie value.')
    return naive


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--category', default='Appliance Synchronization',
                        help='Permission category_name to check')
    parser.add_argument('--module', default='WebChart', help='Permission module')
    parser.add_argument('--save-body', help='Write the raw permission response to this path')
    args = parser.parse_args()

    url = os.environ.get('WC_URL')
    user = os.environ.get('WC_USER')
    password = os.environ.get('WC_PASS')
    if not url or not user:
        print('WC_URL and WC_USER must be set in the environment.', file=sys.stderr)
        return 2
    if not password:
        password = getpass.getpass('WebChart password for {0}: '.format(user))

    body, res = request(url, {'login_user': user, 'login_passwd': password}, 'STEP 1: Login')
    if res is None:
        return 1
    print('\nX-lg_status:   {0!r}'.format(res.headers.get('X-lg_status')))
    print('X-status_desc: {0!r}'.format(res.headers.get('X-status_desc')))

    sessionId = parseSessionId(res.headers.get('Set-Cookie'))
    if not sessionId:
        print('\nCannot continue without a session id.')
        return 1

    params = {
        'f': 'ajaxget',
        's': 'permission',
        'module': args.module,
        'category_name': args.category,
        'session_id': sessionId,
    }
    body, res = request(url, params, 'STEP 2: Permission check')
    if res is None:
        return 1

    section('STEP 3: XML parse (what wcexport does)')
    text = body.decode('utf-8', errors='replace')
    head = text[:8000]
    if '<title>Setup 2FA' in head or 'id="2fa_win"' in head or '2fa_barcode' in head:
        print('DIAGNOSIS: WebChart served the "Setup 2FA" enrollment page instead of XML.')
        print('This user must complete two-factor authentication enrollment before the export')
        print('can run. Note that X-lg_status was still "success", which is why wcexport does')
        print('not detect it and instead fails on the XML parse below.')
    try:
        dom = minidom.parse(StringIO(text))
    except Exception as e:
        print('PARSE FAILED -> this is the "did not return a valid XML response" error.')
        print('{0}: {1}'.format(type(e).__name__, e))
        if args.save_body:
            with open(args.save_body, 'wb') as fp:
                fp.write(body)
            print('Raw body written to {0}'.format(args.save_body))
        return 1

    root = dom.documentElement
    print('Parsed OK. Root element: <{0}>'.format(root.tagName if root is not None else None))
    permissions = dom.getElementsByTagName('permission')
    print('<permission> nodes found: {0}'.format(len(permissions)))
    for node in permissions:
        print('  {0}'.format(node.toxml()))
    if not permissions:
        print('No <permission> nodes -> wcexport would report "did not return any permission nodes".')
        return 1
    try:
        value = int(permissions[0].attributes['value'].value)
    except Exception as e:
        print('Could not read the value attribute: {0}: {1}'.format(type(e).__name__, e))
        return 1
    print('Permission value: {0} ({1})'.format(value, 'DENIED' if value == 0 else 'granted'))
    return 0 if value != 0 else 1


if __name__ == '__main__':
    sys.exit(main())
