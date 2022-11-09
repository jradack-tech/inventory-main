import csv

with open('employee_file.csv', mode='w', encoding='utf-8-sig') as employee_file:
    employee_writer = csv.writer(employee_file)

    employee_writer.writerow(['買主捺印欄', 'Accounting', 'November'])
    employee_writer.writerow(['Erica Meyers', 'IT', 'March'])