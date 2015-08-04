import glob
import re
import subprocess
import pandas
import xlsxwriter


def read_pdfs_directory():
    list_of_pdf_files = ' '.join('"{0}"'.format(pdfName) for pdfName in glob.glob('pdfs/*.pdf'))
    print "List of files that are going to be parsed:"
    print list_of_pdf_files
    subprocess.call('pdf2txt.py -o datafile.txt -t text ' + list_of_pdf_files, shell=True)

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

        out_file.write('cost_centre_number,cost_centre_name,work_request_number,work_request_name,employee_number,employee_name,billing_period,billing_period_from_date,billing_period_to_date,activity_code_number,activity_code_total\n')

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
                employee_name, billing_period, billing_period_from_date, billing_period_to_date, activity_code_number,
                activity_code_total))
                continue
    out_file.close()

def sort_data():
    data = pandas.DataFrame(pandas.read_csv('csvfile.csv'))

    subset_data = data.iloc[:,[2,3,4,5,9,10]]
    grouped_data = subset_data.groupby(["employee_number", "employee_name", "work_request_number", "work_request_name", "activity_code_number"], as_index=True).sum()

    # grouped_data.to_excel('test.xlsx', index_label='label', merge_cells=False)

    workbook = xlsxwriter.Workbook('chart_pie.xlsx')

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})

    # Add the worksheet data that the charts will refer to.
    headings = ['Category', 'Values']
    data = [
        ['Apple', 'Cherry', 'Pecan'],
        [60, 30, 10],
    ]

    worksheet.write_row('A1', headings, bold)
    worksheet.write_column('A2', data[0])
    worksheet.write_column('B2', data[1])

    #######################################################################
    #
    # Create a new chart object.
    #
    chart1 = workbook.add_chart({'type': 'pie'})

    # Configure the series. Note the use of the list syntax to define ranges:
    chart1.add_series({
        'name':       'Pie sales data',
        'categories': ['Sheet1', 1, 0, 3, 0],
        'values':     ['Sheet1', 1, 1, 3, 1],
    })

    # Add a title.
    chart1.set_title({'name': 'Popular Pie Types'})

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('C2', chart1, {'x_offset': 100, 'y_offset': 10})

    workbook.close()

# read_pdfs_directory()
# parse_pdfs()
sort_data()