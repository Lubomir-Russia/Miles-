import re
import os
import mechanize
import cookielib
import BeautifulSoup
import tempfile
import shutil
import unicodedata
from lxml.html import parse

Period_Year = '2013'
Period_Month = '10'
Report_dir = 'C:\\Users\\lptashin\\Documents\\Time_and_money\\Reports\\'

MILES_file = 'MILESplus_Act_as_Representative.htm'
logined = False
read_from = 'file'
#read_from = 'url'
br = mechanize.Browser()
br.set_handle_robots(False)   # ignore robots
br.set_handle_refresh(False)  # can sometimes hang without this
br.addheaders = [('User-agent', 'Firefox')]    # [('User-agent', 'Firefox')]

url_base = 'https://websso.t-systems.com/milesplus/prod/plsql/'
url_login  = url_base + 'prost.create_main'
url_repres = url_base + 'deputy.show_new'
url_act_as = url_base + 'deputy.beginver'
url_report = url_base + 'pselectreport_gc.monthlyma'


def get_page_from(read_from = read_from):
    if read_from == 'file':
        return get_page_from_file()
    else:
        return get_page_from_url()

def get_page_from_file():
    return parse(MILES_file)

def get_page_from_url():
    get_login()
    response = br.open(url_act_as)
    return parse(response)

def get_username_and_pass(username='lptashin', pass_file='pass_file.txt'):
    try:
        with open(pass_file) as pass_file:
            for line in pass_file:
                match = re.match( r'^\s*#', line)
                if match:
                    continue
                else:
                    match = re.search(r'user\w*\s*=\s*(\w+).*pass\w*\s*=\s*(\S+)', line)
                    if match:
                        username_in_file = match.group(1)
                        password = match.group(2)
                        if username == username_in_file:
                            return (username, password)
    except:
        pass
    print 'Password for user ' + username + ' not found in pass file.'
    import getpass
    password = getpass.getpass("Enter your password:")
    return (username, password)

def get_login(*args):
    full_name = get_login_full_name()
    if full_name : return full_name
    print 'Not Logined yet'
    (username, password) = get_username_and_pass(*args)
    br.open(url_login)
    br.select_form(nr = 0)
    br.form['username'] = username
    br.form['password'] = password
    br.submit()
    full_name = get_login_full_name()
    if full_name: 
        print 'Login successfully as ' + full_name
        return full_name

def get_login_full_name():
    page = br.open(url_repres).read()
    match = re.search( r'Display of The Representatives for The User\s*([^\)].*\))', page, re.M)
    if match:
        full_name = match.group(1)
        return full_name
    else:
        return None

def get_names_from_rows(rows):
    names = list()
    for row in rows:
        name = [c.text for c in row][0].strip()
        #print type(name)
        if type(name) == unicode:
            name = unicodedata.normalize('NFKD', name).encode('ascii','ignore')
            if name == 'Voronkov, lexey (lvoronko)':
                name = 'Voronkov, Alexey (lvoronko)'
            else:
                print "Unkown UNICODE name:", name
        names.append(name)
    return names

def get_ma_ids_from_rows(rows):
    ma_ids = list()
    for row in rows:
        match = re.match( r'submitForm\(\'(\d*)\'', row)
        ma_ids.append(match.group(1))
    return ma_ids

def act_as_representative(ma_id):
    br.open(url_act_as)
    br.select_form(nr = 0)
    br.set_all_readonly(False)
    br.form['i_ma_id'] = ma_id
    #br.form['i_ver_id'] = 'undefined'
    br.form['i_form_action'] = 'ACTIVATE'
    br.submit()

def save_report(ma_name, year = Period_Year, month = Period_Month,
                    report_format = 'pdf', prj_id = None):
    report = generate_report(ma_name, year = year, month = month,
                    report_format = report_format, prj_id = prj_id)
    if report:
        try:
            filename = Report_dir + ma_name + ' - ' + year + '-' + month + '.' + report_format
            with open(filename, 'wb') as f:
                shutil.copyfileobj(report, f)
        except:
            print "Cannot save " + report_format + " report for " + ma_name + ' to ' + filename

def generate_report(ma_name, year = Period_Year, month = Period_Month,
                    report_format = 'pdf', prj_id = None):
    report_formats = {'pdf' : '1', 'html': '2', 'txt' : '3'}
    if prj_id == None : prj_id = ['-99999',]
    try:
        br.open(url_report)
        br.select_form(nr = 1)
        br.form['i_bumo_id'] = [year + month,]
        br.form['i_prj_id'] = prj_id
        br.form['i_activity'] = ['j',]
        br.form['i_leistungsart'] = ['j',]
        br.form['i_ratetype'] = ['j',]
        br.form['i_printcode'] = [report_formats[report_format],]
        return br.submit()
    except:
        print "Cannot generate " + report_format + " report for " + ma_name
        return None

def get_groups_of_ma(file = 'People_and_groups.txt'):
    groups_of_ma = {}
    with open(file) as file:
        for line in file:
            match = re.match( r'^\s*#', line)
            if match:
                continue
            else:
                match = re.search(r'(\w.*\))\s*=\s*(.+)', line)
                if match:
                    username = match.group(1)
                    group_line = match.group(2)
                    groups = re.findall(r'[\w\-]+', group_line)
                    groups_of_ma[username] = groups
    return groups_of_ma

def main(groups = ['All',], report_formats = ['pdf',]):
    page = get_page_from(read_from)

    rows = page.xpath("//*[@id=\"beginver\"]/table/tbody/*")
    ma_names = get_names_from_rows(rows)

    rows = page.xpath("//*[@id=\"beginver\"]//input/@onclick")
    ma_ids = get_ma_ids_from_rows(rows)

    ma = dict(zip(ma_names, ma_ids))
    print ma
    my_name = get_login()
    ma[my_name] = '0'

    groups_of_ma = get_groups_of_ma()
    print groups_of_ma

    print 'Names found on the Act as Representative page:'
    for ma_name in sorted(ma.iterkeys()):
        print ma_name
        #print "%6s: %s" % ma[ma_name], ma_name

    print "Getting reports for:\n"
    for ma_name in sorted(ma.iterkeys()):
        include_ma = False
        for group in groups:
            if ((group.lower() == 'all') or (group.lower() in map(lambda each:each.lower(), groups_of_ma[ma_name]))):
                include_ma = True
                break
        if include_ma:
            print ma_name
            act_as_representative(ma[ma_name])
            for report_format in report_formats:
                save_report(ma_name, report_format = report_format)

if __name__ == '__main__':
    main()
