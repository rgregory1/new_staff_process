import pygsheets
import pprint
import os
import yagmail
import credentials
import datetime

is_new_staff = False

# setup credentials for sending email
gmail_user = credentials.gmail_user
gmail_password = credentials.gmail_password
yag = yagmail.SMTP( gmail_user, gmail_password)

# get timestamp for log
temp_timestamp = str(datetime.datetime.now())
print(temp_timestamp)

# open up google sheet to see if new staff have been added
gc = pygsheets.authorize(outh_file='client_secret.json')
workbook = gc.open('New Staff Form (Responses)')
wks = workbook.worksheet_by_title("new_staff_list")

# download all data from sheet as cell_matrix
cell_matrix = wks.get_all_values(returnas='matrix')

# gather 'keys' for new dict from 1st row in sheet
dict_key_list = [x for x in cell_matrix[0] if x != '']

# initialize dict for data
worksheet_data = {}

# put cell_matrix list of lists into a dictionary
counter = 1
for row in cell_matrix:
    # for item in row:
    #     print(item)
    line_dict = dict(zip(dict_key_list, row))
    line_dict['counter'] = counter
    worksheet_data[counter] = line_dict
    counter += 1

# grab copy of the base new staff sheet
original_sheet = gc.open('Original New Staff Sheet')
# open new staff process sheet
staff_sheet = gc.open('New Staff Process')


for staff in worksheet_data:
    if worksheet_data[staff]['Sheet Setup'] == '':

        is_new_staff = True

        # create staff_name to use for human readable names for things
        staff_name = worksheet_data[staff]['First Name'] + ' ' + worksheet_data[staff]['Last Name']

        # print to log, staff name and data
        print(staff_name)
        print(worksheet_data[staff])

        # copy new sheet to workbook
        wks = staff_sheet.add_worksheet(staff_name, src_worksheet=original_sheet.worksheet_by_title("Original"), index=2)

        # get new sheet (last in list) and move to second position in list
        wks_list = staff_sheet.worksheets()
        length_of_list = len(wks_list)
        employee_sheet = staff_sheet.worksheet('index', (length_of_list - 1))
        employee_sheet.index = 1

        # update worksheet with staff info
        employee_sheet.update_cell('C2', staff_name)
        employee_sheet.update_cell('C3', worksheet_data[staff]['Hire Date'])
        employee_sheet.update_cell('C4', worksheet_data[staff]['Start Date'])
        employee_sheet.update_cell('C5', worksheet_data[staff]['Position'])
        employee_sheet.update_cell('C6', worksheet_data[staff]['Building Base'])

        # protect the new page
        employee_sheet.create_protected_range(employee_sheet.get_gridrange('B1', 'G74'))

        # move to MasterList to add formulas to check on completion
        master_list = staff_sheet.worksheet_by_title('MasterList')

        # add check to Administrator List
        admin_check = 'D' + str(worksheet_data[staff]['counter'])
        admin_cell_coord = master_list.cell(admin_check)
        admin_cell_coord.formula = "'" + staff_name + "'!D11"

        # add check to Office Manager List
        office_check = 'E' + str(worksheet_data[staff]['counter'])
        office_cell_coord = master_list.cell(office_check)
        office_cell_coord.formula = "'" + staff_name + "'!D24"

        # add check to Administrative Assistan List
        adas_check = 'F' + str(worksheet_data[staff]['counter'])
        adas_cell_coord = master_list.cell(adas_check)
        adas_cell_coord.formula = "'" + staff_name + "'!D41"

        # add check to Tech Support List
        tech_sup_check = 'G' + str(worksheet_data[staff]['counter'])
        tech_sup_cell_coord = master_list.cell(tech_sup_check)
        tech_sup_cell_coord.formula = "'" + staff_name + "'!D49"

        # add check to Tech Integration List
        tech_int_check = 'H' + str(worksheet_data[staff]['counter'])
        tech_int_cell_coord = master_list.cell(tech_int_check)
        tech_int_cell_coord.formula = "'" + staff_name + "'!D66"

        # add master check
        master_check = 'B' + str(worksheet_data[staff]['counter'])
        master_cell_coord = master_list.cell(master_check)
        master_cell_coord.formula = "'" + staff_name + "'!C7"

        # add Name to Master List
        name_in_list = 'A' + str(worksheet_data[staff]['counter'])
        master_list.update_cell(name_in_list, staff_name)


        # add x to new_staff_list so a new sheet is only added once
        xmark = 'J' + str(worksheet_data[staff]['counter'])
        workbook = gc.open('New Staff Form (Responses)')
        wks = workbook.worksheet_by_title("new_staff_list")
        wks.update_cell(xmark, 'x')
        print('added x to new staff form sheet')

        # begin email notifications
        contents = 'A new staff member, ' + staff_name + ', was added to the New Staff Process spreadsheet, go and check it out. \n\n'
        html = '<a href="https://docs.google.com/spreadsheets/d/1qK55DXqbnKpUrsqUMaStCAw48f0r7NqDXTvbLlAj0Qc/edit#gid=0">New Staff Process spreadsheet</a>'
        yag.send(['rgregory@fnwsu.org', 'jjennett@fnwsu.org', 'dstamour@fnwsu.org', 'dtessier@fnwsu.org', 'mellis@fnwsu.org', 'clongway@fnwsu.org'], 'New Employee', [contents, html])
        print('sent main emails')


        # special email for Jon with email groups in it.
        jon_contents = 'Email Gropus: ' + worksheet_data[staff]['Email Groups'] + '\n\n'
        html = '<a href="https://docs.google.com/spreadsheets/d/1qK55DXqbnKpUrsqUMaStCAw48f0r7NqDXTvbLlAj0Qc/edit#gid=0">New Staff Process spreadsheet</a>'
        yag.send('jhavens@fnwsu.org', 'New Employee', [contents, jon_contents, html])
        print('sent jons emails')

if is_new_staff == False:
    print('program comlpete, no new staff')
