from configparser import ConfigParser
import json
from os import path
try:
    import pandas as pd
    import matplotlib.pyplot as plt
except:
    print('Import Unavilable')
    raise ImportError
    
class Project:
    '''
        This class Project consist of the following property for the data read from csv
        1.Employees which will give the all the employees in the project
        2.Tag which will give the all the Tags related to the project
        3.Billing amount in (INR) will have the total billable amount of the project
        4.Billing amount in (USD)
        5.Total hours spent on the project
        And following methods to the project 
        1.To calulate activity summary --> Table of Tag Vs Employees
        2.To Calculate the employee summary --> Table of Employee Vs Hours spent
        3.To write a report into json file
        4.To write a report into HTML file
        5.To display a bar graph for Employee Vs Billing amount
    '''

    def __init__(self, project_name):
        '''
            To initialize and calulate the properties of the class
        '''
        
        self.fileloc = path.dirname(__file__)
        self.project_name = project_name

        try:
            sheet_file = pd.read_csv((self.fileloc+'/data_files/timesheet.csv'))
        except:
            print('File not found')
            raise FileNotFoundError

        self.project_names = sheet_file['Project Name'].to_list()
        self.employee_detail = sheet_file[sheet_file['Project Name'] == self.project_name]
        self.total_hours_spent = sum(self.employee_detail['Hours(For Calculation)'])
        self.tags = []
        self.employees_name = []

        # To get all the tags and employees name
        for i,j in zip(self.employee_detail['Tag'], self.employee_detail['User']):
            if i not in self.tags:
                self.tags.append(i)
            if j not in self.employees_name:
                self.employees_name.append(j)

        self.role_time = {}
        self.employee_billable = self.employee_detail[self.employee_detail['Billing Type'] == 'Billable']
        # To calculate the role and total time into a dictonary
        for (i,j) in zip(self.employee_billable['Role'], self.employee_billable['Hours(For Calculation)']):
            i = str(i).title()
            i = str(i).replace(' ', '_')
            if i not in self.role_time:
                self.role_time[i] = float(j)
            else:
                self.role_time[i] += float(j)
        
        self.config_bill_amount = ConfigParser()

        try:
            self.config_bill_amount.read((self.fileloc+'/config_files/Config_file_billing_amount.ini'))
        except:
            print('File not found')
            raise FileNotFoundError

        self.billing_amount_in_inr = 0

        # To calculate billable amount w.r.t role 
        for i in self.role_time:
            self.billing_amount_in_inr += (self.role_time[i]*float(self.config_bill_amount.get('Amount', i)))     
        self.billing_amount_in_usd = (self.billing_amount_in_inr/(73.70))

    def calculate_activity_summary(self) :
        '''
            To calculate the activity summary by forming a dictonary 
            with employee name as Key and tag used by employee as Value
            and form them into a Dataframe table
        '''

        self.employee_name_tag = {}
        # To calculate name and tag into a dictonary
        for i in self.employee_detail['User']:
            employee_name_detail = self.employee_detail.loc[self.employee_detail['User'] == i, 'Tag']
            for j in employee_name_detail:
                if i not in self.employee_name_tag.keys():
                    self.employee_name_tag[i] = []
                    self.employee_name_tag[i].append(j)
                else:
                    if j not in self.employee_name_tag[i]:
                        self.employee_name_tag[i].append(j)

        # return name vs tag as a table
        return pd.DataFrame.from_dict(self.employee_name_tag, orient='index').transpose()

    def calculate_employee_summary(self):
        '''
            To calculate employee summary by forming a dictonary
            with employee name as Key and sum of time as Value
            and form them into a Dataframe table
        '''

        self.employee_name_time = {}
        # To calculate name and time into a dictonary
        for (i, j) in zip(self.employee_detail.User, self.employee_detail['Hours(For Calculation)']):
            if i not in self.employee_name_time.keys():
                self.employee_name_time[i] = float(j)
            else:
                self.employee_name_time[i] += float(j)
        self.employee_name_time.popitem()

        # return name vs time as a table
        return pd.DataFrame.from_dict(self.employee_name_time, orient='index').transpose()

    def display_bar_chart(self):
        '''
            Will display a bar chart with employee and bill amount
            of the employee on the project and by forming a dictonary 
            with employee name as Key and sum of billing amount as Value
            converting it into a table and ploting the table as a bar graph
        '''

        self.employee_bill_amount = {}
        # To calculate employee and bill amount into a dictonary
        for (i, j, k) in zip(self.employee_billable['User'], self.employee_billable['Hours(For Calculation)'], self.employee_billable['Role']):
            k = str(k).title()
            k = str(k).replace(' ', '_')
            if i not in self.employee_bill_amount.keys():
                self.employee_bill_amount[i] = (j*float(self.config_bill_amount.get('Amount', k)))
            else:
                self.employee_bill_amount[i] += (j*float(self.config_bill_amount.get('Amount', k)))

        self.employee_vs_bill = pd.DataFrame.from_dict(self.employee_bill_amount, orient='index').transpose()
        self.employee_vs_bill = self.employee_vs_bill.set_index(pd.Index([self.project_name]))
        self.employee_vs_bill.plot(kind = 'bar', figsize = (15, 15))
        plt.show()

    def write_report_to_json(self):

        jsonfileloc = self.fileloc + '/generated_reports/JSON Reports/'

        with open((jsonfileloc + 'Project_activity_summary_report.json'), 'w') as jfile:
            report_activity_summary = {}
            report_activity_summary[self.project_name] = self.calculate_activity_summary().to_dict(orient='index')
            json.dump(report_activity_summary, jfile, indent=4)
        with open((jsonfileloc + 'Project_employee_summary_report.json'), 'w') as jfile:
            report_employee_summary = {}
            report_employee_summary[self.project_name] = self.calculate_employee_summary().to_dict(orient='index')
            json.dump(report_employee_summary, jfile, indent=4)

    def write_report_to_html(self) :

        htmlfileloc = self.fileloc + '/generated_reports/HTML Reports/'

        with open((htmlfileloc + 'Project_activity_summary_report.html'), 'w') as htmfile:
            report_activity_summary = self.calculate_activity_summary().to_html()
            htmfile.write(f'<h1>{self.project_name}</h1>')
            htmfile.write(report_activity_summary)
        with open((htmlfileloc + 'Project_employee_summary_report.html'), 'w') as htmfile:
            report_employee_summary = ((self.calculate_employee_summary().set_index(pd.Index(['Hours Spent']))).transpose()).to_html()
            htmfile.write(f'<h1>{self.project_name}</h1>')
            htmfile.write(report_employee_summary)

class Employee:
    '''
        This class Project consist of the following property for the data read from csv
        1.Projects which will give the all the projects the employee is part of
        2.Tag which will give the all the Tags used by the employee
        3.Billing amount in (INR) will have the total billable amount of the project which the employee
        is part.
        4.Billing amount in (USD)
        5.Total hours spent by the employee 
        And following methods to the project 
        1.To calulate activity summary --> Dictonary of {Tag : Hours spent}
        2.To Calculate the employee summary --> Dictonary of {Project : Hours spent}
        3.To write a report into json file
        4.To write a report into HTML file
        5.To display a bar graph for Project Vs Billing amount
    '''

    def __init__(self, employee_name):
        '''
            To initialize and calulate the properties of the class
        '''

        self.fileloc = path.dirname(__file__)
        self.employee_name = employee_name

        try:
            self.sheet_file = pd.read_csv((self.fileloc+'/data_files/timesheet.csv'))
        except:
            print('File not found')
            raise FileNotFoundError

        self.employee_names = self.sheet_file['User'].to_list()
        self.employee_detail = self.sheet_file[self.sheet_file['User'] == self.employee_name]
        self.total_hours_spent = sum(self.employee_detail['Hours(For Calculation)'])
        self.tags = []
        self.projects_name = []

        # To get all the tags and project name
        for i,j in zip(self.employee_detail['Tag'], self.employee_detail['Project Name']):
            if i not in self.tags:
                self.tags.append(i)
            if j not in self.projects_name:
                self.projects_name.append(j)

        self.role_time = {}
        self.employee_billable = self.employee_detail[self.employee_detail['Billing Type'] == 'Billable']

        # To calculate role and time into a dictonary
        for (i,j) in zip(self.employee_billable['Role'], self.employee_billable['Hours(For Calculation)']):
            i = str(i).title()
            i = str(i).replace(' ', '_')
            if i not in self.role_time:
                self.role_time[i] = float(j)
            else:
                self.role_time[i] += float(j)

        self.config_bill_amount = ConfigParser()

        try:
            self.config_bill_amount.read((self.fileloc+'/config_files/Config_file_billing_amount.ini'))
        except:
            print('File not found')
            raise FileNotFoundError

        self.total_billing_in_inr = 0
        self.billing_amount_in_usd = 0

        # To calculate billable amount w.r.t role
        for i in self.role_time:
            self.total_billing_in_inr += ((self.role_time[i])*float(self.config_bill_amount.get('Amount', i)))  
        self.billing_amount_in_usd = (self.total_billing_in_inr/(73.70)) 

    def calculate_activity_summary(self):
        '''
            To calculate the activity summary by forming a dictonary 
            with Tag as Key and Hours spent as Value
        '''

        self.employee_tag_time = {}

        # To calculate employee and tag into a dictonary
        for (i,j) in zip(self.employee_detail['Tag'], self.employee_detail['Hours(For Calculation)']):
            if i not in self.employee_tag_time.keys():
                self.employee_tag_time[i] = float(j)
            else:
                self.employee_tag_time[i] += float(j)

        # To return dictonary of tag and time spent
        return self.employee_tag_time

    def calculate_project_summary(self):
        '''
            To calculate the activity summary by forming a dictonary 
            with Project as Key and Hours spent as Value
        '''

        self.employee_project_time = {}

        # To calculate project and time into a dictonary
        for (i, j) in zip(self.employee_detail['Project Name'], self.employee_detail['Hours(For Calculation)']):
            if i not in self.employee_project_time.keys():
                self.employee_project_time[i] = float(j)
            else:
                self.employee_project_time[i] += float(j)

        # To return dictonary of project and time spent
        return self.employee_project_time

    def display_bar_chart(self):
        '''
            Will display a bar chart with Project and bill amount
            of the employee on the project and by forming a dictonary 
            with project as Key and sum of billing amount as Value
            converting it into a table and ploting the table as a bar graph
        '''

        self.project_bill_amount = {}

        # To calculate project and billing amount into a dictonary
        for (i, j, k) in zip(self.employee_billable['Project Name'], self.employee_billable['Hours(For Calculation)'], self.employee_billable['Role']):
            k = str(k).title()
            k = str(k).replace(' ', '_')
            if i not in self.project_bill_amount.keys():
                self.project_bill_amount[i] = (j*float(self.config_bill_amount.get('Amount', k)))
            else:
                self.project_bill_amount[i] += (j*float(self.config_bill_amount.get('Amount', k)))
        self.employee_vs_bill = pd.DataFrame.from_dict(self.project_bill_amount, orient='index').transpose()
        self.employee_vs_bill = self.employee_vs_bill.set_index(pd.Index([self.employee_name]))
        self.employee_vs_bill.plot(kind = 'bar', figsize = (15, 15))
        plt.show()

    def write_report_to_json(self):

        jsonfileloc = self.fileloc + '/generated_reports/JSON Reports/'

        with open((jsonfileloc + 'Employee_activity_summary_report.json'), 'w') as jfile:
            report_activity_summary = {}
            report_activity_summary[self.employee_name] = self.calculate_activity_summary()
            json.dump(report_activity_summary, jfile, indent=4)
        with open((jsonfileloc + 'Employee_project_summary_report.json'), 'w') as jfile:
            report_project_summary = {}
            report_project_summary[self.employee_name] = self.calculate_project_summary()
            json.dump(report_project_summary, jfile, indent=4)

    def write_report_to_html(self):

        htmlfileloc = self.fileloc + '/generated_reports/HTML Reports/'

        with open((htmlfileloc + 'Employee_activity_summary_report.html'), 'w') as htmfile:
            report_activity_summary = ((((pd.DataFrame.from_dict(self.calculate_activity_summary(), orient='index')).transpose()).set_index(pd.Index(['Hours Spent']))).transpose()).to_html()
            htmfile.write(f'<h1>{self.employee_name}-Tag Vs Hours </h1>')
            htmfile.write(report_activity_summary)
        with open((htmlfileloc + 'Employee_project_summary_report.html'),'w') as htmfile:
            report_project_summary = ((((pd.DataFrame.from_dict(self.calculate_project_summary(), orient='index')).transpose()).set_index(pd.Index(['Hours Spent']))).transpose()).to_html()
            htmfile.write(f'<h1>{self.employee_name}-Project vs Hours </h1>')
            htmfile.write(report_project_summary)

if __name__ == '__main__':
    choice = int(input('''Choose To Generatre Report:\n
                1.Employee Details\n
                2.Project Details\n'''))

    if choice == 1:
        emp_name = str(input('Enter A Name Of Employee:'))
        a = Employee(emp_name)

        if emp_name in a.employee_names:
            choice = int(input('''Choose one\n
                        1.All the Projects employee is part of\n
                        2.All the tags used by the employee\n
                        3.Total billing amount for the project for the employee (in INR)\n
                        4.Total billing amount for the project for the employee (in USD)\n
                        5.Total hours spent\n
                        6.Calculate acitivity summary → Dictonary - (tag:hours_spent)\n
                        7.Calculate project summary → Dictonary - (project:hours_spent)\n
                        8.Write_report to json\n
                        9.Write_report to html\n
                        10.Display bar chart\n'''))
                        
            if choice == 1:
                print(a.projects_name)
            elif choice == 2:
                print(a.tags)
            elif choice == 3:
                print(a.total_billing_in_inr)
            elif choice == 4:
                print(a.billing_amount_in_usd)
            elif choice == 5:
                print(a.total_hours_spent)
            elif choice == 6:
                print(a.calculate_activity_summary())
            elif choice == 7:
                print(a.calculate_project_summary())
            elif choice == 8:
                a.write_report_to_json()
                print('Report generated At: ../generated_reports/JSON Reports')
            elif choice == 9:
                a.write_report_to_html()
                print('Report generated At: ../generated_reports/HTML Reports')
            elif choice == 10:
                a.display_bar_chart()
            else:
                print('Choice Not Avilable')

        else:
            print('Name Not Avilable')

    elif choice == 2:
        pro_name = str(input('Enter A Project Name:'))
        a = Project(pro_name)

        if pro_name in a.project_names:
            choice = int(input('''Choose one\n
                        1.All the employees working in the project\n
                        2.All the tags related to the project\n
                        3.Total billing amount for the project (in INR)\n
                        4.Total billing amount for the project (in USD)\n
                        5.Total hours spent\n
                        6.Calculate acitivity summary → Table with Rows as Tags and Columns as Employees\n
                        7.Calculate employee summary → Table of Employee vs hours spent\n
                        8.Write_report to json\n
                        9.Write_report to html\n
                        10.Display bar chart\n'''))

            if choice == 1:
                print(a.employees_name)
            elif choice == 2:
                print(a.tags)
            elif choice == 3:
                print(a.billing_amount_in_inr)
            elif choice == 4:
                print(a.billing_amount_in_usd)
            elif choice == 5:
                print(a.total_hours_spent)
            elif choice == 6:
                print(a.calculate_activity_summary())
            elif choice == 7:
                print(a.calculate_employee_summary())
            elif choice == 8:
                a.write_report_to_json()
                print('Report generated At: ../generated_reports/JSON Reports')
            elif choice == 9:
                a.write_report_to_html()
                print('Report generated At: ../generated_reports/HTML Reports')
            elif choice == 10:
                a.display_bar_chart()
            else:
                print('Choice Not Avilable')

        else:
            print('Project Not Avilable')

    else:
        print('Choice Not Avilable')