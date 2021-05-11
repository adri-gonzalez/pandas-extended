import pandas as pd

students_001 = pd.read_excel('./Students.xlsx', sheet_name='Page_001', index_col='ID')
students_002 = pd.read_excel('./Students.xlsx', sheet_name='Page_002', index_col='ID')
print('\n----Page_001----')
print(students_001)
print('\n----Page_002----')
print(students_002)

students_add_dates = students_001.append(students_002)
print(students_add_dates)

stu_col1 = pd.Series({'Name': 'Abel', 'Score': 99})
students_add_col = students_add_dates.append(stu_col1, ignore_index=True)
print(students_add_col)

students_001.at[1, 'Name'] = 'Jack'
students_001.at[1, 'Score'] = 100
print(students_001)

stu_col2 = pd.Series({'ID': 1, "Name": 'Chen', 'Score': 110})
students_001.iloc[0] = stu_col2
print(students_001)

stu_col3 = pd.Series({"Name": 'Scort', 'Score': 110})
part1 = students_001[:15]
part2 = students_001[15:]
students_001 = part1.append(stu_col3, ignore_index=True).append(part2, ignore_index=True)
print(students_001)

students_drop_col = students_001.drop(index=[15])
print(students_drop_col)

for i in range(5, 15):
    students_001['Name'].at[i] = ''

missing = students_001.loc[students_001['Name'] == '']
students_001.drop(missing.index, inplace=True)
print(students_001)
