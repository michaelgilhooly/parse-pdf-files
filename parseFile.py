import os
import csv
import re
import subprocess

# subprocess.call('pdf2txt.py -o datafile.txt -t text pdfs/*.pdf', shell=True)

startflag = False
with open('dataFile.txt','r') as infile:
    with open('costData.txt','w') as outfile:
        writer = csv.writer(outfile)
        writer.writerow( ('RunDate', 'DateFrom', 'DateTo', 'CostCenter', 'WorkRequest', 'Manager', 'ActivityCode', 'Time', 'EmployeeID', 'EmployeeName') )
        for line in infile:
            if 'Total for Billing Period: ' in line:
                if startflag:
                    writer.writerow( (runDate, 'dateFrom', 'dateTo', 'costCenter', workRequest, 'manager', 'activityCode', 'time', 'employeeID', 'employeeName') )
                else:
                    startflag = True
                continue

            if 'Dev Manager :' in line:
                startflag = False

            run_date = re.match('Run Date:     (\d\d.\d\d.\d\d\d\d)', line)
            print run_date
            if run_date:
                runDate = run_date.group(1)

            work_request = re.match('^Work Request: (\w\w\w\w\w\w\w\w\w)', line)
            print work_request
            if work_request:
                workRequest = work_request.group(1)



            # user_name = re.match('^USER: (\w{5})', line)
            # if user_name:
            #     user = user_name.group(1)

            # acc_name = re.findall('HH\d+ \w+,\w+', line)
            # if acc_name:
            #     account, name = acc_name[0].split(' ')

            # date_time = re.findall('(?<=Coll: ).+(?= Recd:)', line)
            # if date_time:
            #     date, time = date_time[0].split('-')

            # source_re = re.findall('(?<=Source: ).+',line)
            # if source_re:
            #     source = source_re[0].strip()

            # anti_spec = re.findall('^ +(?!Source)\w+ *\w+ + \S+', line)
            # if anti_spec:
            #     stripped_list = anti_spec[0].strip().split()
            #     anti = stripped_list[-1]
            #     spec = ' '.join(stripped_list[:-1])