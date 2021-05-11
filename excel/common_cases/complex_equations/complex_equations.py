import pandas as pd
import numpy as np


def get_circumscribed_circle_area(lengh, height):
    r = np.sqrt(lengh ** 2 + height ** 2) / 2
    return r ** 2 * np.pi


def wrapper(row):
    return get_circumscribed_circle_area(row['Length'], row['Height'])


rectangles = pd.read_excel('./Rectangles.xlsx', index_col='ID')
print(rectangles)

rectangles['CA'] = rectangles.apply(lambda row: get_circumscribed_circle_area(row['Length'], row['Height']), axis=1)
# rectangles['CA'] = rectangles.apply(wrapper,axis=1)
print(rectangles)
