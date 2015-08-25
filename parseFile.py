import glob
import re
import pandas
import csv
import os
import pdf2txt


def read_pdfs_directory():
    print "*****PROCESS STARTED*****\n"
    print "Cleaning directory"
    clean_up()
    list_of_pdf_files = glob.glob('pdfs/*.pdf')
    print "List of files that are going to be parsed:"
    print list_of_pdf_files
    with open("datafile.txt", "a") as myfile:
        for individual in list_of_pdf_files:
            print "Reading: {}".format(individual)
            pdf2txt.main(['', '-o', 'individualfile.txt', '-t', 'text', individual])
            individual_file = open('individualfile.txt', 'r')
            individual_content = individual_file.read()
            myfile.write(individual_content)
            print "Finished reading: {}".format(individual)
    print "Completed reading PDF files"
    print "Created datafile.txt from PDFs"


def parse_pdfs():
    cost_centre_regex = re.compile('^Cost Centre:\s+([\d]+)\s+(.*)')
    cost_centre_match_groups_count = 2
    work_request_regex = re.compile('^Work Request:\s+([^\s]+)\s+(.*)')
    work_request_match_groups_count = 2
    activity_code_regex = re.compile('^\s+Total for Activity Code:\s+([^\s]+)\s+([^\d\.])+\s+([\d]+.[\d]+).*')
    activity_code_match_groups_count = 3
    employee_line_regex = re.compile('^\s+(\d+)\s+(\w+\s+\w+)\s+\d+\d+.\d+\s+\d+\.\d+.*')
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

    with open('datafile.txt', 'r') as datafile, open('csvFile.csv', 'wb') as out_file:
        writer = csv.writer(out_file)

        print "Parsing completeDataFile.txt"
        print "Creating csvFile.csv"

        writer.writerow(
            ["cost_centre_number", "cost_centre_name", "work_request_number", "work_request_name", "employee_number",
             "employee_name", "billing_period", "billing_period_from_date", "billing_period_to_date",
             "activity_code_number", "activity_code_total"])

        for line in datafile:
            line = line.rstrip('\n')
            line = line.decode('utf-8').replace(u'\u2013', '-')

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

                csv_data_set = [cost_centre_number, cost_centre_name, work_request_number, work_request_name,
                                employee_number,
                                employee_name, billing_period, billing_period_from_date, billing_period_to_date,
                                activity_code_number,
                                activity_code_total]

                writer.writerow(csv_data_set)
                print "Print to file: csvFile.csv"
                continue


def sort_data():
    print "Reading the csvFile.csv..."
    data = pandas.DataFrame(pandas.read_csv('csvFile.csv'))

    print "Sorting through csvFile table..."
    subset_data = data.iloc[:, [2, 3, 4, 5, 9, 10]]
    subset_data.columns = ['WRK Number', 'WRK Name', 'Employee Number', 'Name', 'Activity Code', 'Total']
    subset_data.insert(5, "Utilisation", "")

    values = set(['LVE_ANNUAL', 'LVE_OTHER', 'LVE_LONG', 'LVE_PUBLIC', 'LVE_SICK', 'CONTR_ABS'])

    print "Creating team-utilisation-workbook.xlsx..."
    writer = pandas.ExcelWriter('team-utilisation-workbook.xlsx', engine='xlsxwriter')

    print "Subsetting and looping through data for each person..."
    for name in set(subset_data['Name']):
        name_subset_data = subset_data[subset_data['Name'] == name]

        for row_index, row in name_subset_data.iterrows():
            if row['Activity Code'] in values:
                name_subset_data.loc[row_index, 'Utilisation'] = 'Available time'
            else:
                name_subset_data.loc[row_index, 'Utilisation'] = 'Billable time'

        print "Grouping activity data and summing times for " + name + "..."
        activity_data = name_subset_data.groupby(
            ['Employee Number', 'Name', 'WRK Number', 'WRK Name', 'Activity Code', 'Utilisation'],
            as_index=True).sum()

        activity_table_length = len(activity_data.index)

        print "Writing activity table for " + name + "..."
        activity_data.to_excel(writer, sheet_name=name)

        print "Grouping utilisation data and summing times for " + name + "..."
        utilisation_data = name_subset_data.groupby(['Utilisation'], as_index=False)['Total'].sum()
        utilisation_table_length = len(utilisation_data.index)

        print "Writing utilisation table for " + name + "..."
        utilisation_data.to_excel(writer, sheet_name=name, startrow=activity_table_length+4)

        # Builds the chart from data
        workbook = writer.book
        worksheet = writer.sheets[name]

        activity_chart = workbook.add_chart({'type': 'pie'})

        activity_chart.set_title({'name': '{}\'s Utilisation chart'.format(name)})

        print "Creating activity chart for " + name + "..."
        activity_chart.add_series({
            'categories': '={}!$E$3:$E${}'.format(name, 2 + activity_table_length),
            'values': '={}!$G$3:$G${}'.format(name, 2 + activity_table_length),
            'data_labels': {'percentage': True},
        })

        print "Inserting activity chart for " + name + "..."
        worksheet.insert_chart('I{}'.format(2), activity_chart)

        utilisation_chart = workbook.add_chart({'type': 'pie'})

        utilisation_chart.set_title({'name': '{}\'s Utilisation chart'.format(name)})

        print "Creating utilisation chart for " + name + "..."
        utilisation_chart.add_series({
            'categories': '={}!$B${}:$B${}'.format(name, 6 + activity_table_length,
                                                   5 + activity_table_length + utilisation_table_length),
            'values': '={}!$C${}:$C${}'.format(name, 6 + activity_table_length,
                                               5 + activity_table_length + utilisation_table_length),
            'data_labels': {'percentage': True},
        })

        print "Inserting utilisation chart for " + name + "..."
        worksheet.insert_chart('I{}'.format(18), utilisation_chart)

        worksheet.set_column('A:C', 15)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:F', 12)

    print "Saving xlsx file..."
    writer.save()
    print "Deleting supporting files"
    # TODO: Re-initiate the cleanup comment
    clean_up()
    print "\n*****PROCESS COMPLETED*****"


def clean_up():
    for project_file in ('datafile.txt', 'individualfile.txt', 'csvFile.csv'):
        try:
            os.remove(project_file)
        except OSError:
            pass


read_pdfs_directory()
parse_pdfs()
sort_data()