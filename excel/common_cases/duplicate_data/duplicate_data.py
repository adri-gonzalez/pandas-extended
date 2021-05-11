import pandas as pd

students = pd.read_excel('./Students_Duplicates.xlsx')
print(students)
print(students.columns)

dupe = students.duplicated(subset='Name')
print(dupe)

dupe = dupe[dupe]
print(students.iloc[dupe.index])

students.drop_duplicates(subset='Name', inplace=True)
print(students)
