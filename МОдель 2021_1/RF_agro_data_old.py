# -*- coding: utf-8 -*-
"""
Created on Sun Sep  5 11:15:18 2021

@author: n.klyukin
"""

import pandas as pd
import statsmodels.api as sm
from sklearn.linear_model import LinearRegression
import numpy as np

def pandas_insert(df, idx, row_contents):
	top = df.iloc[:idx]
	bot = df.iloc[idx:]
	inserted = pd.concat([top, row_contents, bot])
	return inserted

def filter_years(year):
	if year <= 2020:
		return True
	else:
		return False

def write_string_to_excel(writer, message, sheet_name, startrow, add_row_pos=0):
	message_str = pd.Series(message)
	message_str.to_excel(writer, sheet_name=sheet_name, header=False, index=False, startrow=add_row_pos + startrow)

def shorten_years_str(years):
	years_str = ''
	for i, year in enumerate(years):
		if i == 0:
			if i == len(years) - 1:
				years_str += str(year)
				return years_str
			elif years[i+1] - year > 1:
				years_str += str(year) + ', '
				years_str += shorten_years_str(years[i+1:])
				return years_str
			else:
				years_str += str(year) + '-'
		elif year - years[i-1] == 1:
			if i == len(years) - 1:
				years_str += str(year)
				return years_str
			continue
		elif year - years[i-1] > 1:
			years_str += str(years[i-1]) + ', '
			years_str += shorten_years_str(years[i:])
			return years_str


data_area = pd.read_excel('RF_Agro_data.xls', sheet_name='Данные', )

#data_area_CFO1 = data_area.loc[(data_area['Территории']=='Ростовская область') & (data_area['Тип хозяйства']=='Хозяйства всех категорий') & (data_area['Культуры']=='Зерновые и зернобобовые культуры')[:]

#data_production = 
#data_udobr = 
#data_meat_prod = 



#areas_CFO = 
#areas = Area_data_land.index.tolist()[0:2]

#Списки данных
# data_area_CFO = data_area['Территории'].value_counts()
# data_area_orgs = data_area['Тип хозяйства'].unique()[2]
# data_area_cultures = data_area['Культуры'].unique()

# data_areas_CFO = data_area_CFO.index.tolist()[1:8]


#data_CFO = data_area.loc[(data_area['Территории'] == data_area_CFO)]
#data_CFO = data_area.transpose()



#Wolrd_Land_data_AgricLand = data_land.loc[(data_land['Area'] =='World') & (data_land['Item']=='Agricultural land')][['Area', 'Item', 'Unit', 'Value']]

