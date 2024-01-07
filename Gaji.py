import PySimpleGUI as sg
import pandas as pd
import os

class CalculateSalary:
    def __init__(self):
        self.base_salary = {
            "Staf": 1000000,
            "Sekretaris": 1200000,
            "Manajer": 1500000,
            "Direktur": 2000000,
            "Keuangan": 1300000,
            "Purchasing": 1100000,
            "Engineering": 1400000
        }
        self.status = ["Married", "Single"]
        self.tax = {"Single": 0.05, "Married": 0.10}
        self.family_allowance_rate = 0.20
        self.child_allowance_rate = 0.10
        self.max_children_allowance = 3
        self.excel_file = 'salary_data.xlsx'

    def calculate_allowance(self, position, status, num_children):
        allowance = 0
        if status == 'Married':
            base = self.base_salary[position]
            self.family_allowance = self.family_allowance_rate * base
            self.child_allowance = self.child_allowance_rate * base * min(num_children, self.max_children_allowance)
            allowance = self.family_allowance + self.child_allowance
        return allowance

    def calculate_tax(self, position, status):
        base = self.base_salary[position]
        tax_rate = self.tax[status]
        return tax_rate * base

    def calculate_total(self, position, status, num_children):
        base = self.base_salary[position]
        allowance = self.calculate_allowance(position, status, num_children)
        tax = self.calculate_tax(position, status)
        total = base + allowance - tax
        return total

    def save_to_excel(self, name, position, status, num_children, total):
        # Prepare data dictionary for DataFrame
        data = {
            'Name': name,
            'Position': position,
            'Marital Status': status,
            'Number of Children': num_children,
            'Basic Salary': self.base_salary[position],
            'Total Allowance': self.calculate_allowance(position, status, num_children),
            'Tax': self.calculate_tax(position, status),
            'Total Salary': total
        }

        df = pd.DataFrame([data])
        
        if os.path.exists(self.excel_file):
            with pd.ExcelWriter(self.excel_file, mode='a', if_sheet_exists='overlay') as writer:
                
                start_row = writer.sheets['Sheet1'].max_row
                df.to_excel(writer, index=False, header=False, startrow=start_row)
        else: 
            df.to_excel(self.excel_file, index=False)
        
        sg.Popup('Success', 'Data has been saved to Excel')

    def run(self):
        layout = [
            [sg.Text('Name', size=(10, 1))],
            [sg.InputText(key='name')],
            [sg.Text('Please enter your Position')],
            [sg.InputCombo(list(self.base_salary.keys()), key='position')],
            [sg.Text('Please enter your Marital Status')],
            [sg.InputCombo(self.status, key='status')],
            [sg.Text('Please enter the Number of Children')],
            [sg.Spin([i for i in range(0, 16)], initial_value=0, key='num_children')],
            [sg.Button('Calculate'), sg.Exit()]
        ]

        window = sg.Window('Salary Calculator', layout)
        while True:
            event, values = window.read()
            if event in (None, 'Exit'):
                break
            if event == 'Calculate':
                try:
                    num_children = int(values['num_children'])
                    position = values['position']
                    status = values['status']
                    name = values['name']
                    total = int(self.calculate_total(position, status, num_children))
                    sg.Popup('Total Salary', f'The total salary for {name} is: {total}')
                    self.save_to_excel(name, position, status, num_children, total)  # Save the data to Excel
                except ValueError as e:
                    sg.Popup('Kesalahan', 'Pastikan semua input sudah benar', str(e))
                except KeyError as e:
                    sg.Popup('Kesalahan', 'Pilihan posisi atau status tidak valid', str(e))

        window.close()

if __name__ == "__main__":
    salary_calculator = CalculateSalary()
    salary_calculator.run()

