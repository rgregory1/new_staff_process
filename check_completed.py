import pygsheets
from pprint import pprint
import os
import yagmail
import credentials

# setup gmail link
gmail_user = credentials.gmail_user
gmail_password = credentials.gmail_password
yag = yagmail.SMTP( gmail_user, gmail_password)


def make_dict_from_matrix(master_list_matrix):
    dict_key_list = [x for x in master_list_matrix[0] if x != '']

    master_list_data = {}

    counter = 1
    for row in master_list_matrix:
        line_dict = dict(zip(dict_key_list, row))
        line_dict['counter'] = counter
        master_list_data[counter] = line_dict
        counter += 1
    return master_list_data

print(10 * '\n')


gc = pygsheets.authorize(outh_file='client_secret.json')
workbook = gc.open('New Staff Process')
MasterList = workbook.worksheet_by_title("MasterList")

master_list_matrix = MasterList.get_all_values(returnas='matrix')


dict_key_list = [x for x in master_list_matrix[0] if x != '']

master_list_data = {}

counter = 1
for row in master_list_matrix:
    line_dict = dict(zip(dict_key_list, row))
    line_dict['counter'] = counter
    master_list_data[counter] = line_dict
    counter += 1

# pprint(master_list_data)

final_admin_todo = ''
final_office_todo = ''
final_admin_ass_todo = ''
final_tech_sup_todo = ''
final_tech_int_todo = ''

for staff in master_list_data:
    if master_list_data[staff]['Status'] == 'Not Complete':
        # print(master_list_data[staff]['Staff Name'])
        this_staff_sheet = workbook.worksheet_by_title(master_list_data[staff]['Staff Name'])
        this_staff_matrix = this_staff_sheet.get_all_values(returnas='matrix')

        counter = 1
        this_staff_data = {}
        new_line = {}

        for line in this_staff_matrix:
            this_line_data = {}
            # this_line_data['row'] = counter
            this_line_data['a'] = line[0]
            this_line_data['b'] = line[1]
            this_staff_data[counter] = this_line_data
            counter = counter + 1
        # pprint(this_staff_data)

        # begin admin email notifications
        admin_list = [12,13,14,15,16,17,18,19,20]
        admin_todo = ''
        for number in admin_list:
            # print(this_staff_data[number])
            if this_staff_data[number]['a'] == '':
                admin_todo = admin_todo + this_staff_data[number]['b'] + '\n'
        if admin_todo != '':
            final_admin_todo = final_admin_todo + master_list_data[staff]['Staff Name'] + '\n \n' + admin_todo + '\n\n'

        # begin office manager notifications
        office_list = [25,26,27,28,29,30,31,32,33,34,35,36,37]
        office_todo = ''
        for number in office_list:
            # print(this_staff_data[number])
            if this_staff_data[number]['a'] == '':
                office_todo = office_todo + this_staff_data[number]['b'] + '\n'
        if office_todo != '':
            final_office_todo = final_office_todo +master_list_data[staff]['Staff Name'] + '\n \n' + office_todo + '\n\n'

        # begin Admin Assistant notifications
        admin_ass_list = [42,43,44,45]
        admin_ass_todo = ''
        for number in admin_ass_list:
            # print(this_staff_data[number])
            if this_staff_data[number]['a'] == '':
                admin_ass_todo = admin_ass_todo + this_staff_data[number]['b'] + '\n'
        if admin_ass_todo != '':
            final_admin_ass_todo = final_admin_ass_todo +master_list_data[staff]['Staff Name'] + '\n \n' + admin_ass_todo + '\n\n'

        # begin tech support notifications
        tech_sup_list = [50,51,52,53,54,55,56,57,58,59,60,61,62]
        tech_sup_todo = ''
        for number in tech_sup_list:
            # print(this_staff_data[number])
            if this_staff_data[number]['a'] == '':
                tech_sup_todo = tech_sup_todo + this_staff_data[number]['b'] + '\n'
        if tech_sup_todo != '':
            final_tech_sup_todo = final_tech_sup_todo +master_list_data[staff]['Staff Name'] + '\n \n' + tech_sup_todo + '\n\n'

        # begin tech int notifications
        tech_int_list = [67,68,69,70]
        tech_int_todo = ''
        for number in tech_int_list:
            # print(this_staff_data[number])
            if this_staff_data[number]['a'] == '':
                tech_int_todo = tech_int_todo + this_staff_data[number]['b'] + '\n'
        if tech_int_todo != '':
            final_tech_int_todo = final_tech_int_todo +master_list_data[staff]['Staff Name'] + '\n \n' + tech_int_todo + '\n\n'


print(final_admin_todo)
print(final_office_todo)

# begin email notifications

contents = 'This is your friendly weekly reminder of things to do for new staff memmbers. \n \n \n'
contents2 = 'Due to your efficiency, there is actually nothing for you to do for new hires!'
html = '<a href="https://docs.google.com/spreadsheets/d/1qK55DXqbnKpUrsqUMaStCAw48f0r7NqDXTvbLlAj0Qc/edit#gid=0">New Staff Process spreadsheet</a>'

# Admin emails
# if final_admin_todo != '':
#     yag.send('jjennett@fnwsu.org', 'New Staff Weekly Reminder', [contents, final_admin_todo, html])
# else:
#     yag.send('jjennett@fnwsu.org', 'New Staff Weekly Reminder', [contents, contents2, html])

# # office manager emails
# if final_office_todo != '':
#     yag.send('add chrissy', 'New Staff Weekly Reminder', [contents, final_office_todo, html])
# else:
#     yag.send('add chrissy', 'New Staff Weekly Reminder', [contents, contents2, html])
#
# # admin assistant emails
# if final_admin_ass_todo != '':
#     yag.send(['add mary', 'add dawn'], 'New Staff Weekly Reminder', [contents, final_admin_ass_todo, html])
# else:
#     yag.send(['add mary', 'add dawn'], 'New Staff Weekly Reminder', [contents, contents2, html])

# tech support emails
if final_tech_sup_todo != '':
    yag.send('jhavens@fnwsu.org', 'New Staff Weekly Reminder', [contents, final_tech_sup_todo, html])
else:
    yag.send('jhavens@fnwsu.org', 'New Staff Weekly Reminder', [contents, contents2, html])

# tech integration emails
if final_tech_int_todo != '':
    yag.send('rgregory@fnwsu.org', 'New Staff Weekly Reminder', [contents, final_tech_int_todo, html])
else:
    yag.send('rgregory@fnwsu.org', 'New Staff Weekly Reminder', [contents, contents2, html])
