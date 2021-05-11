import pandas as pd


def score_validation(row):
    try:
        assert 0 <= row.Score <= 100
    except:
        print(f'#{row.ID}\tstudent {row.Name} has an invalid score {row.Score}.')


def score_validation2(row):
    if not 0 <= row.Score <= 100:
        print(f'#{row.ID}\tstudent {row.Name} has an invalid score {row.Score}.')


students = pd.read_excel('./Students.xlsx')
print(students)
print(students.columns)

students.apply(score_validation, axis=1)
