import glob
import re
import subprocess
import pandas

def read_pdfs_directory():
    print "*****PROCESS STARTED*****"
    list_of_pdf_files = ' '.join('"{0}"'.format(pdfName) for pdfName in glob.glob('pdfs/*.pdf'))
    print "List of files that are going to be parsed:"
    print list_of_pdf_files
    subprocess.call('pdf2txt.py -o datafile.txt -t text ' + list_of_pdf_files, shell=True)
    print "Completed reading PDF files"
    print "Created datafile.txt from PDFs"


def parse_pdfs():
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
        print "Parsing datafile.txt"
        print "Creating csvfile.csv"
        out_file.write(
            'cost_centre_number,cost_centre_name,work_request_number,work_request_name,employee_number,employee_name,billing_period,billing_period_from_date,billing_period_to_date,activity_code_number,activity_code_total\n')

        for line in datafile:
            line = line.rstrip('\n')
            cost_centre_groups = parse_cost_centre_line(line)
            if cost_centre_groups:
                cost_centre_number = cost_centre_groups[0].strip()
                print "Found Cost Centre Number: {}".format(cost_centre_number)
                cost_centre_name = cost_centre_groups[1].strip()
                print "Found Cost Centre Name: {}".format(cost_centre_name)
                continue

            work_request_groups = parse_work_request_line(line)
            if work_request_groups:
                work_request_number = work_request_groups[0].strip()
                print "Found Work Request Number: {}".format(work_request_number)
                work_request_name = work_request_groups[1].strip()
                print "Found Work Request Name: {}".format(work_request_name)
                continue

            employee_line_groups = parse_employee_line(line)
            if employee_line_groups:
                employee_number = employee_line_groups[0].strip()
                print "Found Employee Number: {}".format(employee_number)
                employee_name = employee_line_groups[1].strip()
                print "Found Employee Name: {}".format(employee_name)
                continue

            billing_period_groups = parse_billing_period_line(line)
            if billing_period_groups:
                billing_period = billing_period_groups[0]
                print "Found Billing Period: {}".format(billing_period)
                billing_period_from_date = reformat_date(billing_period_groups[1])
                print "Found Billing from data: {}".format(billing_period_from_date)
                billing_period_to_date = reformat_date(billing_period_groups[2])
                print "Found Billing to date: {}".format(billing_period_to_date)
                continue

            activity_code_groups = parse_activity_code_line(line)
            if activity_code_groups:
                activity_code_number = activity_code_groups[0].strip()
                print "Found Activity Code Number: {}".format(activity_code_number)
                activity_code_total = activity_code_groups[2].strip()
                print "Found Activity Code Total: {}".format(activity_code_total)

                out_file.write('%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n' % (
                    cost_centre_number, cost_centre_name, work_request_number, work_request_name, employee_number,
                    employee_name, billing_period, billing_period_from_date, billing_period_to_date,
                    activity_code_number,
                    activity_code_total))
                print "Print to file: csvfile.csv"
                continue
    out_file.close()


def sort_data():

    print "Reading the csvfile.csv..."
    data = pandas.DataFrame(pandas.read_csv('csvfile.csv'))

    print "Sorting through csvfile table..."
    subset_data = data.iloc[:, [2, 3, 4, 5, 9, 10]]
    subset_data.columns = ['WRK Number', 'WRK Name', 'Employee Number', 'Name', 'Activity Code', 'Total']

    print "Creating team-utilisation-workbook.xlsx..."
    writer = pandas.ExcelWriter('team-utilisation-workbook.xlsx', engine='xlsxwriter')

    print "Subsetting and looping through data per person..."
    for name in set(subset_data['Name']):
        name_subset_data = subset_data[subset_data['Name'] == name]

        print "Grouping data and summing activity times..."
        grouped_data = name_subset_data.groupby(
            ['Employee Number', 'Name', 'WRK Number', 'WRK Name', 'Activity Code'],
            as_index=True).sum()

        table_length = len(grouped_data.index)

        print "Writing table to file..."
        grouped_data.to_excel(writer, sheet_name=name)

        workbook = writer.book
        worksheet = writer.sheets[name]

        chart = workbook.add_chart({'type': 'pie'})

        chart.set_title({'name': '{}\'s Utilisation chart'.format(name)})

        print "Creating chart..."
        chart.add_series({
            'categories': '={}!$E$3:$E${}'.format(name, 2 + table_length),
            'values': '={}!$F$3:$F${}'.format(name, 2 + table_length),
            'data_labels': {'percentage': True},
        })

        print "Inserting chart..."
        worksheet.insert_chart('H{}'.format(2), chart)

        worksheet.set_column('A:C', 10)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:F', 10)

    print "Saving xlsx file..."
    writer.save()
    print "*****PROCESS COMPLETED*****"

read_pdfs_directory()
parse_pdfs()
sort_data()
