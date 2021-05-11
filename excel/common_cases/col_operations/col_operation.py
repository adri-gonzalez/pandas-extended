import pandas as pd
import numpy as np

students_001 = pd.read_excel('./Students.xlsx', sheet_name='Page_001')
students_002 = pd.read_excel('./Students.xlsx', sheet_name='Page_002')
print('\n----Page_001----')
print(students_001)
print('\n----Page_002----')
print(students_002)

students_add_dates = pd.concat([students_001, students_002], axis=1)
print(students_add_dates)

students = pd.concat([students_001, students_002]).reset_index(drop=True)
print(students)

students['Age'] = np.arange(0, len(students))
print(students)

students.drop(columns='Age', inplace=True)
print(students)

students.insert(1, column='Foo', value=np.repeat('foo', len(students)))
print(students)

students.rename(columns={'Foo': 'FOO', 'Name': 'NAME'}, inplace=True)
print(students)

students['ID'] = students['ID'].astype(float)
for i in range(3, 5):
    students['ID'].at[i] = np.nan

students.dropna(inplace=True)
print(students)
