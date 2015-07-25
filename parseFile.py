import os
import csv
import re
import subprocess

# subprocess.call('pdf2txt.py -o outfile.txt -t text pdfs/*.pdf', shell=True)

startflag = False
with open('sample.txt','r') as infile:
    with open('costData.txt','w') as outfile:
        outfile.write('Date,Time,Name,Account,Specimen,Source,Antibiotic,User\n')
        for line in infile:
            if '---------------' in line:
                if startflag:
                    outfile.write(','.join((date, time, name, account, spec, source, anti, user))+'\n')
                else:
                    startflag = True
                continue
            if 'Activity' in line:
                startflag = False

            user_name = re.match('^USER: (\w{5})', line)
            if user_name:
                user = user_name.group(1)

            acc_name = re.findall('HH\d+ \w+,\w+', line)
            if acc_name:
                account, name = acc_name[0].split(' ')

            date_time = re.findall('(?<=Coll: ).+(?= Recd:)', line)
            if date_time:
                date, time = date_time[0].split('-')

            source_re = re.findall('(?<=Source: ).+',line)
            if source_re:
                source = source_re[0].strip()

            anti_spec = re.findall('^ +(?!Source)\w+ *\w+ + \S+', line)
            if anti_spec:
                stripped_list = anti_spec[0].strip().split()
                anti = stripped_list[-1]
                spec = ' '.join(stripped_list[:-1])