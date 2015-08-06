import re
import subprocess
import glob
import csv
import pandas
import xlsxwriter


def list_and_parse_files():
    list_all_pdf_files = ' '.join('"{0}"'.format(pdfName) for pdfName in glob.glob('pdfs/*.pdf'))
    subprocess.call('pdf2txt.py -o datafile.txt -t text ' + list_all_pdf_files, shell=True)


def parse_data():
    global cost_centre_regex, cost_centre_match_groups_count, work_request_regex, work_request_match_groups_count, activity_code_regex, activity_code_match_groups_count, employee_line_regex, employee_line_match_groups_count, billing_period_regex, billing_period_match_groups_count, MyException, parse_line
    cost_centre_regex = re.compile('^Cost Centre:\s+([\d]+)\s+(.*)')
    cost_centre_match_groups_count = 2
    work_request_regex = re.compile('^Work Request:\s+([^\s]+)\s+(.*)')
    work_request_match_groups_count = 2
    activity_code_regex = re.compile('^\s+Total for Activity Code:\s+([^\s]+)\s+([^\d\.])+\s+([\d]+.[\d]+).*')
    activity_code_match_groups_count = 3
    employee_line_regex = re.compile('^\s+([\d]+)\s+([^\d]+)\s+\d+\d+.\d+\s+\d+\.\d+.*')
    employee_line_match_groups_count = 2
    billing_period_regex = re.compile('^([\d]+)\s+\(\s+(\d{2}\.\d{2}\.\d{4})\s+-\s+(\d{2}\.\d{2}\.\d{4})\s+\).*')
    billing_period_match_groups_count = 3

    class MyException(Exception):
        pass

    def parse_line(line, regex, match_groups_count):
        match = regex.search(line)
        if match is None:
            return None
        if len(match.group()) != len(line):
            raise MyException('Matching group [%s] does not match entire line [%s]' % (match.group(), line))
        if len(match.groups()) != match_groups_count:
            raise MyException('Expected %s matching groups, found %s' % (match_groups_count, match.groups()))
        return match.groups()

    def parse_cost_centre_line(line):
        return parse_line(line, cost_centre_regex, cost_centre_match_groups_count)

    def parse_work_request_line(line):
        return parse_line(line, work_request_regex, work_request_match_groups_count)

    def parse_activity_code_line(line):
        return parse_line(line, activity_code_regex, activity_code_match_groups_count)

    def parse_employee_line(line):
        return parse_line(line, employee_line_regex, employee_line_match_groups_count)

    def parse_billing_period_line(line):
        return parse_line(line, billing_period_regex, billing_period_match_groups_count)

    def reformat_date(date):
        date_parts = date.split('.')
        return date_parts[2] + '-' + date_parts[1] + '-' + date_parts[0]

    out_file = open('csvfile.csv', 'w')
    with open('dataFile.txt', 'r') as datafile:
        # TODO: use the csv library
        out_file.write(
            'cost_centre_number,cost_centre_name,work_request_number,work_request_name,employee_number,employee_name,billing_period,billing_period_from_date,billing_period_to_date,activity_code_number,activity_code_total\n ')

        for line in datafile:
            line = line.rstrip('\n')
            cost_centre_groups = parse_cost_centre_line(line)
            if cost_centre_groups:
                cost_centre_number = cost_centre_groups[0].strip()
                cost_centre_name = cost_centre_groups[1].strip()
                continue

            work_request_groups = parse_work_request_line(line)
            if work_request_groups:
                work_request_number = work_request_groups[0].strip()
                work_request_name = work_request_groups[1].strip()
                continue

            employee_line_groups = parse_employee_line(line)
            if employee_line_groups:
                employee_number = employee_line_groups[0].strip()
                employee_name = employee_line_groups[1].strip()
                continue

            billing_period_groups = parse_billing_period_line(line)
            if billing_period_groups:
                billing_period = billing_period_groups[0]
                billing_period_from_date = reformat_date(billing_period_groups[1])
                billing_period_to_date = reformat_date(billing_period_groups[2])
                continue

            activity_code_groups = parse_activity_code_line(line)
            if activity_code_groups:
                activity_code_number = activity_code_groups[0].strip()
                activity_code_total = activity_code_groups[2].strip()

                out_file.write('%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n' % (
                    cost_centre_number, cost_centre_name, work_request_number, work_request_name, employee_number,
                    employee_name, billing_period, billing_period_from_date, billing_period_to_date,
                    activity_code_number,
                    activity_code_total))
                continue
    out_file.close()


def sort_data_for_person(name, data):
    # TODO: Create unique user identifier from employee_number and employee_name. Ex: "Michael-s72381")
    # for row_index, row in data.iterrows():
    #     employee_first_name = re.search('^(.+)\s', row['Name'])
    #     worksheet_name = '{}-{}'.format(employee_first_name.group(1), row['Employee Number'])

    grouped_data = data.groupby(
        ['Employee Number', 'Name', 'WRK Number', 'WRK Name', 'Activity Code'],
        as_index=True).sum()

    print grouped_data

    sheetName = name

    writer = pandas.ExcelWriter('test-workbook-{}.xlsx'.format(name), engine='xlsxwriter')
    grouped_data.to_excel(writer, sheet_name=sheetName)

    workbook = writer.book
    worksheet = writer.sheets[sheetName]

    chart = workbook.add_chart({'type': 'pie'})

    chart.add_series({
        'categories': '={}!$E$3:$E${}'.format(sheetName, 9),
        'values': '={}!$F$3:$F${}'.format(sheetName, 9),
        'data_labels': {'percentage': True},
    })

    worksheet.insert_chart('H{}'.format(2), chart)

    writer.save()


def sort_data():
    data = pandas.DataFrame(pandas.read_csv('csvfileNEW.csv'))

    subset_data = data.iloc[:, [2, 3, 4, 5, 9, 10]]
    subset_data.columns = ['WRK Number', 'WRK Name', 'Employee Number', 'Name', 'Activity Code', 'Total']
    # subset_data["Worksheet"] = "N"
    # print subset_data

    for name in set(subset_data['Name']):
        name_subset_data = subset_data[subset_data['Name'] == name]
        sort_data_for_person(name, name_subset_data)


# list_and_parse_files()
# parse_data()
sort_data()
