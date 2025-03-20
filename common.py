import tk
from tkinter import ttk
import tkinter
import tkinter.messagebox as tkm
import urllib
from urllib.request import urlopen
from urllib.parse import urlparse
import ssl
import re
import os
import csv
import time
import string
from io import StringIO
from datetime import datetime, timedelta
import time
import math
from xml.dom import minidom
from tkinter import ttk

validChars = '_-,.()[] {0}{1}'.format(string.ascii_letters, string.digits)
def sanitizeFilename(filename):
    return ''.join([x if x in validChars else '_' for x in filename])

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
    def __init__(self, win, fullExport=True):
        self.win = win
        self.fullExport = fullExport
        self.scheduleID = None
        lastrow = 5
        if fullExport:
            win.title('WebChart Export Utility')
        else:
            lastrow = 10
            win.title('WebChart Document Export Utility')

        self.logfp = None

        wcinputs = tkinter.Frame(win)
        wcinputs.pack()
        tkinter.Label(wcinputs, text='WebChart URL').grid(row=0)
        tkinter.Label(wcinputs, text='WebChart Username').grid(row=1)
        tkinter.Label(wcinputs, text='WebChart Password').grid(row=2)

        if fullExport:
            tkinter.Label(wcinputs, text='System Report').grid(row=3)
            tkinter.Label(wcinputs, text='Print Definition').grid(row=4)
        else:
            tkinter.Label(wcinputs, text='Begin Date').grid(row=3)
            tkinter.Label(wcinputs, text='End Date').grid(row=4)
            tkinter.Label(wcinputs, text='Extra CGI').grid(row=5)
            tkinter.Label(wcinputs, text='Document Types').grid(row=6)
            tkinter.Label(wcinputs, text='Single Export').grid(row=7)
            tkinter.Label(wcinputs, text='Schedule At').grid(row=8)
            tkinter.Label(wcinputs, text='Schedule Every').grid(row=9)

        tkinter.Label(wcinputs, text='Output Directory').grid(row=lastrow)

        self.url = tkinter.Entry(wcinputs)
        self.url.grid(row=0, column=1)

        self.username = tkinter.Entry(wcinputs)
        self.username.grid(row=1, column=1)
        self.password = tkinter.Entry(wcinputs, show='*')
        self.password.grid(row=2, column=1)

        if fullExport:
            self.report = tkinter.Entry(wcinputs)
            self.report.insert(0, 'WebChart Export')
            self.report.grid(row=3, column=1)
            self.printdef = tkinter.Entry(wcinputs)
            self.printdef.insert(0, 'WebChart Export')
            self.printdef.grid(row=4, column=1)
        else:
            now = datetime.now()
            self.bd_d = tkinter.Entry(wcinputs, width=2)
            self.bd_d.insert(0, now.day)
            self.bd_d.grid(row=3, column=1)
            self.bd_m = tkinter.Entry(wcinputs, width=2)
            self.bd_m.insert(0, now.month)
            self.bd_m.grid(row=3, column=2)
            self.bd_y = tkinter.Entry(wcinputs, width=4)
            self.bd_y.insert(0, now.year)
            self.bd_y.grid(row=3, column=3)
            self.ed_d = tkinter.Entry(wcinputs, width=2)
            self.ed_d.insert(0, now.day)
            self.ed_d.grid(row=4, column=1)
            self.ed_m = tkinter.Entry(wcinputs, width=2)
            self.ed_m.insert(0, now.month)
            self.ed_m.grid(row=4, column=2)
            self.ed_y = tkinter.Entry(wcinputs, width=4)
            self.ed_y.insert(0, now.year)
            self.ed_y.grid(row=4, column=3)

            self.cgi = tkinter.Entry(wcinputs)
            self.cgi.grid(row=5, column=1)

            self.cda = tkinter.IntVar()
            self.ccr = tkinter.IntVar()
            cda = tkinter.Checkbutton(wcinputs, text='CDA', variable=self.cda)
            ccr = tkinter.Checkbutton(wcinputs, text='CCR', variable=self.ccr)
            cda.grid(row=6, column=1)
            ccr.grid(row=6, column=2)

            self.schedule = tkinter.IntVar()
            tkinter.Radiobutton(wcinputs, text='No Schedule', variable=self.schedule,
                value=0).grid(row=7, column=1)
            tkinter.Radiobutton(wcinputs, text='Date', variable=self.schedule, value=1).grid(
                row=8, column=1)
            self.scheduleOnm = tkinter.Entry(wcinputs, width=2)
            self.scheduleOnd = tkinter.Entry(wcinputs, width=2)
            self.scheduleOny = tkinter.Entry(wcinputs, width=4)
            self.scheduleOnh = tkinter.Entry(wcinputs, width=2)
            self.scheduleOnmm = tkinter.Entry(wcinputs, width=2)
            self.scheduleOnm.grid(row=8, column=2)
            self.scheduleOnd.grid(row=8, column=3)
            self.scheduleOny.grid(row=8, column=4)
            self.scheduleOnh.grid(row=8, column=5)
            self.scheduleOnmm.grid(row=8, column=6)
            tkinter.Radiobutton(wcinputs, text='Every Nth day of the month',
                variable=self.schedule, value=2).grid(row=9, column=1)
            self.scheduleEveryd = tkinter.Entry(wcinputs, width=2)
            self.scheduleEveryd.grid(row=9, column=2)


        self.outstring = tkinter.StringVar()
        self.outstring.set('wcexport')
        self.outdirE = tkinter.Entry(wcinputs, textvariable=self.outstring)
        self.outdirE.grid(row=lastrow, column=1)
        self.progressFrame = tkinter.Frame(win)
        self.progressFrame.pack()

        buttons = tkinter.Frame(win)
        buttons.pack()
        self.exportText = tkinter.StringVar()
        self.exportText.set('Export')
        self.exportButton = tkinter.Button(buttons, textvariable=self.exportText, fg='Green',
            command=self.exportWrapper)
        self.exportButton.grid(row=0, column=0)
        tkinter.Button(buttons, text='Exit', command=win.quit).grid(row=0, column=1)

        self.notes = tkinter.Frame(win)
        self.notes.pack()
        def genNotes(*args):
            self.outdir = getOutDir(self, self.outstring.get())
            text = ''
            if self.fullExport:
                text += '* System report must be a valid system report in the given webchart system.\n'\
                '\tIt must contain at least the column \'pat_id\' to indicate which charts are to be '\
                'exported.\n\tIt may also contain a column for \'filename\' to specify what the resulting '\
                'pdf file should be named\n'\
                '\tAny additional columns with values starting with \'?\' will be '\
                'treated as a relative url to download a separate file (CCD, CCR, etc..)\n'
            text += '* All downloads will be placed in: [ {0} ]'.format(self.outdir)
            self.notesText = tkinter.Label(self.notes, text=text, justify=tkinter.LEFT).grid(row=0)
        genNotes(None)
        self.outstring.trace('w', genNotes)

    def getURLResponse(self, url, data={}):
        if data and hasattr(self, 'session_id'):
            data['session_id'] = self.session_id
        try:
            res = urlopen(url, context=getSSLContext(),
                    data=urllib.parse.urlencode(data, doseq=True).encode("utf-8") if data else None)
        except Exception as e:
            raise Warning('Internal error in urlopen [ {0} : {1} ] at [ {2} : {3} ]'.format(
                type(e), str(e), url, data))
        if res.getcode() not in [200,401]:
            raise Warning('Invalid http response code [ {0} ]'.format(res.headers.getcode()))
        if res.headers.get('X-lg_status').lower() != 'success':
            self.log('Login failed for {0}: {1}'.format(url, data))
            raise Exception('Login failed [ {0} ]'.format(res.headers.get('X-status_desc')))
        out = res.read()
        if out[:3] == [239, 187, 191]:
            # Strip out utf-8 BOM from webchart CSV output
            out = out[3:]
        return out, res

    def validateInputs(self):
        d = {
            'url': 'URL',
            'username': 'Username',
            'password': 'Password',
        }
        if self.fullExport:
            d.update({
                'report': 'System report',
                'printdef': 'Print definition',
            })
        for i, v in d.items():
            if not getattr(self, i).get():
                tkm.showwarning(message='{0} must be given'.format(v))
                return False
        if not self.fullExport:
            if not self.cda.get() and not self.ccr.get() and not self.cgi.get():
                tkm.showwarning(message='At least one document type must be selected or '\
                    'some extra cgi must be specified')
                return False
        return True
    
    def validateURL(self):
        try:
            res = urlopen(self.url.get(), context=getSSLContext())
        except Exception as e:
            tkm.showwarning(message='Url could not be opened: {0}'.format(e))
            return False
        if res.getcode() not in [200,401]:
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
        out, res = self.getURLResponse(self.url.get(), {
            'f': 'ajaxget',
            's': 'permission',
            'module': 'WebChart',
            'category_name': 'Appliance Synchronization'
        })
        try:
            dom = minidom.parse(StringIO(out.decode("utf-8")))
        except Exception as e:
            tkm.showerror(message='WebChart permission check did not return a valid XML response')
            return False
        permissions = dom.getElementsByTagName('permission')
        if not permissions:
            tkm.showerror(message='WebChart permission check did not return any permission nodes')
            return False
        try:
            if int(permissions[0].attributes['value'].value) == 0:
                tkm.showerror(message='You do not have the required "Appliance Synchronization" permission to perform this export')
                return False
        except Exception as e:
            tkm.showerror(message='Your permission to perform this export could not be determined')
            return False
        return True

    def validatePrintDef(self):
        if self.fullExport:
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
            if 'Unable to find Print Definition' in out.decode("utf-8", errors="ignore"):
                tkm.showwarning(message='Invalid print definition: {0}'.format(self.printdef.get()))
                return False
        return True

    def validate(self):
        if self.validateInputs() and self.validateCredentials()\
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
        reader = csv.DictReader(StringIO(out.decode("utf-8-sig", errors="ignore")), delimiter=',')
        if 'pat_id' not in reader.fieldnames:
            print("Raw CSV Data:", out.decode("utf-8", errors="ignore")[:500])  # Print first 500 chars
            print("Detected Headers:", reader.fieldnames)
            tkm.showerror(message='System report [ {0} ] does not contain the required "pat_id" column'
                .format(self.report.get()))
            return False
        extrafields = [x for x in reader.fieldnames if x not in ['pat_id', 'filename']]
        self.charts = []
        for row in reader:
            filename = row['pat_id']
            if 'filename' in row:
                filename = sanitizeFilename(row['filename'])
            self.charts.append({
                'pat_id': row['pat_id'],
                'filename': '{0}.pdf'.format(os.path.join(self.outdir, filename)),
                'urls': dict(zip(extrafields, [row[x] for x in extrafields]))
            })
        return True 

    def log(self, msg):
        if self.logfp is not None:
            self.logfp.write('{0} {1}\n'.format(time.ctime(), msg))

    def cancelSchedule(self):
        self.win.after_cancel(self.scheduleID)
        self.exportText.set('Export')
        self.exportButton.configure(command=self.exportWrapper)

    def exportWrapper(self):
        if self.fullExport or not self.schedule.get():
            self.export()
        else:
            if not self.validate():
                return False
            now = datetime.now()
            if not tkm.askyesno(title='Schedule Export?', message='Confirm scheduled export?'):
                return
            if self.schedule.get() == 1:
                m = self.scheduleOnm.get()
                d = self.scheduleOnd.get()
                y = self.scheduleOny.get()
                h = self.scheduleOnh.get() or 0
                mm = self.scheduleOnmm.get() or 0
                try:
                    date = datetime(int(y), int(m), int(d), int(h), int(mm))
                    mm = int(mm)
                    h = int(d)
                    if date < now or mm < 0 or mm > 60 or h < 0 or h > 24:
                        raise Exception('no')
                except Exception as e:
                    print(e)
                    tkm.showwarning(message='Please enter a valid future date in for the MM/DD/YYYY format')
                    return
            elif self.schedule.get() == 2:
                try:
                    d = int(self.scheduleEveryd.get())
                    if d < 0 or d > 28:
                        raise Exception('lol no')
                except Exception as e:
                    tkm.showwarning(message='Please enter a valid day of the month (1-28)')
                    return
                if now.day < d:
                    date = now + timedelta(days=d - now.day)
                else:
                    if now.month == 12:
                        y = now.year + 1
                        m = 1
                    else:
                        y = now.year
                        m = now.month + 1
                    date = datetime(y, m, d)
            self.exportButton.configure(command=self.cancelSchedule)
            self.exportText.set('Cancel Schedule')
            self.win.wm_state('iconic')
            interval = int(math.ceil((date - now).total_seconds())) * 1000
            if self.schedule.get() == 1:
                self.scheduleID = self.win.after(interval, self.export)
            else:
                if self.scheduleID:
                    self.export()
                self.scheduleID = self.win.after(interval, self.exportWrapper)

    def export(self):
        if not self.validate():
            return False
        if self.fullExport and not self.getSystemReport():
            return False
        if not os.path.exists(self.outdir):
            os.makedirs(self.outdir)
        if not self.logfp:
            self.logfp = open(os.path.join(self.outdir, 'wcexport.log'), 'a')
            if self.fullExport:
                self.log('Beginning Export of {0} with report [ {1} ] and print definition [ {2} ]'.format(
                    self.url.get(), self.report.get(), self.printdef.get()))
                maxprogress = len(self.charts)
                data = {
                    'f': 'chart',
                    's': 'print',
                    'print_definition': self.printdef.get(),
                    'print_reason': 'System Export',
                    'print_printer_name': 'My Computer',
                    'print_priority': 2,
                    'print_submit_print': '-- Print --'
                }
            else:
                self.log('Beginning document export')
                maps = {
                    'cda': '19',
                    'ccr': '21',
                }
                data = {
                    'f': 'chart',
                    's': 'search',
                    'search_method': 'doc',
                    'servicestartdateMONTH': self.bd_m.get(),
                    'servicestartdateDAY': self.bd_d.get(),
                    'servicestartdateYEAR': self.bd_y.get(),
                    'servicestartdateTIME': '00:00',
                    'serviceenddateMONTH': self.ed_m.get(),
                    'serviceenddateDAY': self.ed_d.get(),
                    'serviceenddateYEAR': self.ed_y.get(),
                    'serviceenddateTIME': '23:59',
                    'pat_search': 'Search',
                    'docstg_type': [],
                    'csv': '1',
                }
                if self.cgi.get():
                    data.update(urlparse.parse_qs(self.cgi.get()))
                data['docstg_type'].extend([v for k,v in maps.items() if getattr(self, k).get()])
                try:
                    out, _ = documents = self.getURLResponse(self.url.get(), data)
                except Exception as e:
                    tkm.showerror(message='Failed to get documents csv list {0}'.format(e))
                    return False
                reader = csv.DictReader(StringIO(out.decode("utf-8", errors="ignore")), delimiter=',')
                documents = []
                if 'MR Number' not in reader.fieldnames:
                    tkm.showerror(message='CSV data does not contain MR Number, unable to '\
                        'write unique filenames without this column')
                    return False
                for row in reader:
                    filename = '{MR Number}_{Last}_{First}_{Doc ID}'.format(**row)
                    documents.append({
                        'doc_id': row['Doc ID'],
                        'filename': '{0}'.format(os.path.join(self.outdir,
                            sanitizeFilename(
                                '{MR Number}_{Last}_{First}_{Doc ID}'.format(**row))))
                    })
                maxprogress = len(documents) 

        def printChart(chart_id):
            d = data.copy()
            d['pat_id'] = chart_id
            url = None
            out, _ = self.getURLResponse(self.url.get(), d)
            m = re.search('input type="hidden" name="job_url" value="(.*?)"', out.decode("utf-8", errors="ignore"))
            if m:
                url = m.group(1)
            elif 'You Currently Do Not Have Access to:' in out.decode("utf-8", errors="ignore"):
                self.log('Access denied to chart: {0}'.format(chart_id))
            else:
                m = re.search('pjob_id=([\\d]+)', out.decode("utf-8", errors="ignore"))
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
                if out.decode("utf-8", errors="ignore").strip() == 'Print Spool is currently empty.':
                    out = 'Chart print contained no data to be printed'
                    filename = '{0}.txt'.format(os.path.splitext(filename)[0])
                with open(filename, 'wb') as fp:
                    fp.write(out)

        def getChart(chart_id, filename):
            if not os.path.exists(filename):
                downloadPrintJob(printChart(chart_id), chart_id, filename)
            return filename

        def getExternalUrls(baseName, urls):
            for filename, url in urls.items():
                if url.startswith('?'):
                    out, _ = self.getURLResponse(self.url.get(), urlparse.parse_qs(url[1:]))
                    fname = '{0}_{1}'.format(os.path.splitext(baseName)[0], filename)
                    if not os.path.exists(fname):
                        with open(fname, 'wb') as fp:
                            fp.write(out)
                    else:
                        self.log('Skipping already present external url file {0}'.format(fname))

        def downloadDocument(doc_id, filename):
            if not os.path.exists(filename):
                out, _ = self.getURLResponse(self.url.get(), {
                    'f': 'stream',
                    'doc_id': doc_id,
                    'rawdata': '1'
                })
                with open(filename, 'wb') as fp:
                    fp.write(out)

        self.progress = ttk.Progressbar(self.progressFrame, orient=tkinter.HORIZONTAL, length=400,
                                        mode='determinate', maximum=maxprogress)
        self.progress.grid(row=0, column=0)
        self.progressLabel = tkinter.Label(self.progressFrame, text='0 / {0}'.format(maxprogress))
        self.progressLabel.grid(row=1, column=0)
        self.progressCurrent = tkinter.Label(self.progressFrame, text='Exporting ...')
        self.progressCurrent.grid(row=2, column=0)
        self.win.update()

        msg = 'Export Complete'
        for idx, current in enumerate(self.charts if self.fullExport else documents):
            if self.fullExport:
                self.progressCurrent['text'] = 'Exporting chart [ {0} ]'.format(current['pat_id'])
            else:
                self.progressCurrent['text'] = 'Exporting document [ {0} ]'.format(current['doc_id'])
            self.win.update()
            try:
                if self.fullExport:
                    getChart(current['pat_id'], current['filename'])
                    self.log(str(current['urls']))
                    getExternalUrls(current['filename'], current['urls'])
                else:
                    downloadDocument(current['doc_id'], current['filename'])
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
            self.progressLabel['text'] = '{0} / {1}'.format(idx + 1, maxprogress)
            self.win.update()
        if not self.schedule.get():
            tkm.showinfo(message=msg)
        self.log(msg)
        self.logfp.close()
        self.logfp = None

