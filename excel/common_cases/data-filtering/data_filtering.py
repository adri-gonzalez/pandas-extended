import pandas as pd

Students = pd.read_excel('./Students.xlsx', index_col='ID')

print(Students)


def age_18_to_30(age):
    return 18 <= age < 30


def level_a(score):
    return 85 <= score <= 100


Students = Students.loc[Students['Age'].apply(lambda age: 18 <= age < 30)]
print(Students)

Students = Students.loc[Students.Age.apply(age_18_to_30)] \
    .loc[Students.Score.apply(level_a)]
print(Students)
