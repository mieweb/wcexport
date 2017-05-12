#!/usr/bin/env python
import Tkinter as tk
import ttk
import tkMessageBox as tkm
import urllib
import urllib2
import urlparse
import ssl
import re
import os
import csv
import time
import string
from StringIO import StringIO

title = 'WebChart Export Utility'

def getSSLContext():
    ctx = ssl.create_default_context()
    ctx.check_hostname=False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx

def getOutDir(cls, path):
    if os.path.isabs(path):
        return os.path.expanduser(path)
    else:
        return os.path.join(os.path.expanduser('~'), path)

class MainWin(object):
    def __init__(self, win):
        self.win = win
        win.title(title)

        self.logfp = None

        wcinputs = tk.Frame(win)
        wcinputs.pack()
        tk.Label(wcinputs, text='WebChart URL').grid(row=0)
        tk.Label(wcinputs, text='WebChart Username').grid(row=1)
        tk.Label(wcinputs, text='WebChart Password').grid(row=2)
        tk.Label(wcinputs, text='System Report').grid(row=3)
        tk.Label(wcinputs, text='Print Definition').grid(row=4)
        tk.Label(wcinputs, text='Output Directory').grid(row=5)

        self.url = tk.Entry(wcinputs)
        self.url.grid(row=0, column=1)

        self.username = tk.Entry(wcinputs)
        self.username.grid(row=1, column=1)
        self.password = tk.Entry(wcinputs, show='*')
        self.password.grid(row=2, column=1)

        self.report = tk.Entry(wcinputs)
        self.report.insert(0, 'WebChart Export')
        self.report.grid(row=3, column=1)
        self.printdef = tk.Entry(wcinputs)
        self.printdef.insert(0, 'WebChart Export')
        self.printdef.grid(row=4, column=1)
        self.outstring = tk.StringVar()
        self.outstring.set('wcexport')
        self.outdirE = tk.Entry(wcinputs, textvariable=self.outstring)
        self.outdirE.grid(row=5, column=1)
        self.progressFrame = tk.Frame(win)
        self.progressFrame.pack()

        buttons = tk.Frame(win)
        buttons.pack()
        tk.Button(buttons, text='Begin Export', fg='Green', command=self.export).grid(row=0, column=0)
        tk.Button(buttons, text='Exit', command=win.quit).grid(row=0, column=1)

        self.notes = tk.Frame(win)
        self.notes.pack()
        def genNotes(*args):
            self.outdir = getOutDir(self, self.outstring.get())
            self.notesText = tk.Label(self.notes,
                     text='* System report must be a valid system report in the given webchart system.\n'\
                     '\tIt must contain at least the column \'pat_id\' to indicate which charts are to be '\
                     'exported.\n\tIt may also contain a column for \'filename\' to specify what the resulting '\
                     'pdf file should be named\n'\
					 '\tAny additional columns with values starting with \'?\' will be '\
					 'treated as a relative url to download a separate file (CCD, CCR, etc..)\n'
                     '* All downloads will be placed in: [ {0} ]'.format(self.outdir,
                     justify=tk.LEFT)).grid(row=0)
        genNotes(None)
        self.outstring.trace('w', genNotes)

    def getURLResponse(self, url, data={}):
        if data and hasattr(self, 'session_id'):
            data['session_id'] = self.session_id
        try:
            res = urllib2.urlopen(url, context=getSSLContext(),
                    data=urllib.urlencode(data, doseq=True) if data else None)
        except Exception as e:
            raise Warning('Internal error in urlopen [ {0} : {1} ] at [ {2} : {3} ]'.format(
                type(e), str(e), url, data))
        if res.getcode() not in [200]:
            raise Warning('Invalid http response code [ {0} ]'.format(res.headers.getcode()))
        if res.headers.get('X-lg_status').lower() != 'success':
            self.log('Login failed for {0}: {1}'.format(url, data))
            raise Exception('Login failed [ {0} ]'.format(res.headers.get('X-status_desc')))
        out = res.read()
        return out, res

    def validateInputs(self):
        d = {
            'url': 'URL',
            'report': 'System report',
            'printdef': 'Print definition',
            'username': 'Username',
            'password': 'Password',
        }
        for i, v in d.iteritems():
            if not getattr(self, i).get():
                tkm.showwarning(message='{0} must be given'.format(v))
                return False
        return True
    
    def validateURL(self):
        try:
            res = urllib2.urlopen(self.url.get(), context=getSSLContext())
        except Exception as e:
            tkm.showwarning(message='Url could not be opened: {0}'.format(e))
            return False
        if res.getcode() not in [200]:
            tkm.showwarning(message='{0} does not appear to be a valid url'.format(self.url.get()))
            return False
        if 'X-lg_status' not in res.headers:
            tkm.showwarning(message='{0} does not appear to be a webchart system'.format(self.url.get()))
            return False
        return True

    def validateCredentials(self):
        d = {
            'login_user': self.username.get(),
            'login_passwd': self.password.get(),
        }
        try:
            out, res = self.getURLResponse(self.url.get(), d)
        except Exception as e:
            tkm.showwarning(message='Invalid webchart credentials: {0}'.format(e))
            return False
        cookie = res.headers.get('Set-Cookie').split('=')
        try:
            c = cookie[1].split(';')[0]
        except:
            tkm.showerror(message='WebChart cookie could not be parsed: {0}'.format(cookie))
            return False
        else:
            self.session_id = c
        return True

    def validatePrintDef(self):
        d = {
            'f': 'chart',
            's': 'printdefedit',
            'item': 'def',
            'opp': 'edit',
            'olddefname': self.printdef.get()
        }
        try:
            out, _ = self.getURLResponse(self.url.get(), data=d)
        except Exception as e:
            tkm.showwarning(message='Invalid url for printdef validation: {0}'.format(e))
            return False
        if 'Unable to find Print Definition' in out:
            tkm.showwarning(message='Invalid print definition: {0}'.format(self.printdef.get()))
            return False
        return True

    def validate(self):
        if self.validateInputs() and self.validateURL() and self.validateCredentials()\
            and self.validatePrintDef():
            return True

    def getSystemReport(self):
        data = {
            'f': 'admin',
            's': 'system_report',
            'opp': 'querycsv',
            'report_name': self.report.get(),
            'submit_query': '1'
        }
        try:
            out, res = self.getURLResponse(self.url.get(), data)
        except Exception as e:
            tkm.showerror(message='Invalid system report response: {0}'.format(e))
            return False
        if [ord(x) for x in out[0:3]] == [239 ,187 ,191]:
            # Strip out utf-8 BOM from webchart
            out = out[3:]
        reader = csv.DictReader(StringIO(out), delimiter=',')
        if 'pat_id' not in reader.fieldnames:
            tkm.showerror(message='System report [ {0} ] does not contain the required "pat_id" column'
                .format(self.report.get()))
            return False
        extrafields = [x for x in reader.fieldnames if x not in ['pat_id', 'filename']]
        self.charts = []
        validChars = '_-,.()[] {0}{1}'.format(string.ascii_letters, string.digits)
        for row in reader:
            filename = row['pat_id']
            if 'filename' in row:
                filename = ''.join([x if x in validChars else '_' for x in row['filename']])
            self.charts.append({
                'pat_id': row['pat_id'],
                'filename': '{0}.pdf'.format(os.path.join(self.outdir, filename)),
                'urls': dict(zip(extrafields, [row[x] for x in extrafields]))
            })
        return True 

    def log(self, msg):
        if self.logfp is not None:
            self.logfp.write('{0} {1}\n'.format(time.ctime(), msg))

    def export(self):
        if not self.validate():
            return False
        if not self.getSystemReport():
            return False
        if not os.path.exists(self.outdir):
            os.makedirs(self.outdir)
        if not self.logfp:
            self.logfp = open(os.path.join(self.outdir, 'wcexport.log'), 'a')
            self.log('Beginning Export of {0} with report [ {1} ] and print definition [ {2} ]'.format(
                self.url.get(), self.report.get(), self.printdef.get()))
        data = {
            'f': 'chart',
            's': 'print',
            'print_definition': self.printdef.get(),
            'print_reason': 'System Export',
            'print_printer_name': 'My Computer',
            'print_priority': 2,
            'print_submit_print': '-- Print --'
        }
        self.progress = ttk.Progressbar(self.progressFrame, orient=tk.HORIZONTAL, length=400,
                                        mode='determinate', maximum=len(self.charts))
        self.progress.grid(row=0, column=0)
        self.progressLabel = tk.Label(self.progressFrame, text='0 / {0}'.format(len(self.charts)))
        self.progressLabel.grid(row=1, column=0)
        self.progressCurrent = tk.Label(self.progressFrame, text='Exporting ...')
        self.progressCurrent.grid(row=2, column=0)
        self.win.update()

        def printChart(chart_id):
            d = data.copy()
            d['pat_id'] = chart_id
            url = None
            out, _ = self.getURLResponse(self.url.get(), d)
            m = re.search('input type="hidden" name="job_url" value="(.*?)"', out)
            if m:
                url = m.group(1)
            elif 'You Currently Do Not Have Access to:' in out:
                self.log('Access denied to chart: {0}'.format(chart_id))
            else:
                m = re.search('pjob_id=([\d]+)', out)
                if m:
                    self.log('Print for chart {0} failed due to document errors, '\
                             'but what printed successfully was saved'.format(chart_id))
                    pjob_id = m.group(1)
                    url = '{0}?f=admin&s=printman&v=view_pjob&job_id={1}&session_id={2}'.format(
                        self.url.get(), pjob_id, self.session_id)
                else:
                    self.log('Failed to find job_url input or a pjob_id in the response for chart {0}'.format(chart_id))
            return url

        def downloadPrintJob(url, chart_id, filename):
            if url:
                out, _ = self.getURLResponse(url)
                if out.strip() == 'Print Spool is currently empty.':
                    out = 'Chart print contained no data to be printed'
                    filename = '{0}.txt'.format(os.path.splitext(filename)[0])
                with open(filename, 'wb') as fp:
                    fp.write(out)

        def getChart(chart_id, filename):
            if not os.path.exists(filename):
                downloadPrintJob(printChart(chart_id), chart_id, filename)
            return filename

        def getExternalUrls(baseName, urls):
            for filename, url in urls.iteritems():
                if url.startswith('?'):
                    out, _ = self.getURLResponse(self.url.get(), urlparse.parse_qs(url[1:]))
                    fname = '{0}_{1}'.format(os.path.splitext(baseName)[0], filename)
                    if not os.path.exists(fname):
                        with open(fname, 'w') as fp:
                            fp.write(out)
                    else:
                        self.log('Skipping already present external url file {0}'.format(fname))

        msg = 'Export Complete'
        for idx, chart in enumerate(self.charts):
            self.progressCurrent['text'] = 'Exporting chart [ {0} ]'.format(chart['pat_id'])
            self.win.update()
            try:
                getChart(chart['pat_id'], chart['filename'])
                self.log(str(chart['urls']))
                getExternalUrls(chart['filename'], chart['urls'])
            except Warning as w:
                self.log(w)
                if not tkm.askyesno(title='Warning',
                        message='{0}\n\nDo you want to continue the export?'.format(w)):
                    msg = 'Export Aborted Due To User Request'
                    break
            except Exception as e:
                self.log(e)
                tkm.showerror(title='Fatal Error', message=e)
                msg = 'Export Aborted Due to Error'
                break
            self.progress.step()
            self.progressLabel['text'] = '{0} / {1}'.format(idx + 1, len(self.charts))
            self.win.update()
        tkm.showinfo(message=msg)
        self.log(msg)
        self.logfp.close()
        self.logfp = None

def main():
    win = tk.Tk()
    app = MainWin(win)
    win.mainloop()
    try:
        win.destroy()
    except tk.TclError:
        pass

if __name__ == '__main__':
    main()

