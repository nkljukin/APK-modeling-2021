# -*- coding: utf-8 -*-
"""
Created on Thu Oct 21 10:10:21 2021

@author: n.klyukin
"""

import pandas as pd

meteo_Stations = pd.read_excel('RF_meteo_stations_list.xlsx', sheet_name='Станции_области', header=0, usecols='A:C')

#ОСАДКИ
meteo_R = pd.read_excel('RF_meteo_R.xlsx', sheet_name='R_mes', header=0, usecols='A:R')
meteo_R = meteo_R.merge(meteo_Stations, on=['Номер станции'], how='right')
meteo_R = meteo_R.set_index('Год')




#ТЕМПЕРАТУРА
meteo_T = pd.read_excel('RF_meteo_T.xlsx', sheet_name = 'T_mes', header=0, usecols='A:S')
meteo_T = meteo_T.merge(meteo_Stations, on=['Номер станции'], how='right')
meteo_T = meteo_T.set_index('Год')

rrr = meteo_R.loc[(meteo_R['Область'] == '        Белгородская область')][['Область', 'Янв_R', 'Фев_R', 'Мар_R', 'Апр_R', 'Май_R', 'Июн_R', 'Июл_R', 'Авг_R', 'Сен_R', 'Окт_R', 'Ноя_R', 'Дек_R', 'Year_Sum_R', 'Sum_Mar-May_R', 'Sum_Jun-Jul_R', 'Sum_Sen-Nov_R']]

rrr = pd.pivot_table(rrr,
					 index = 'Год',
					 values = ['Янв_R', 'Фев_R', 'Мар_R', 'Апр_R', 'Май_R', 'Июн_R', 'Июл_R', 'Авг_R', 'Сен_R', 'Окт_R', 'Ноя_R', 'Дек_R', 'Year_Sum_R', 'Sum_Mar-May_R', 'Sum_Jun-Jul_R', 'Sum_Sen-Nov_R'],
					 aggfunc = 'mean')
rrr = rrr.reindex(columns = ['Янв_R', 'Фев_R', 'Мар_R', 'Апр_R', 'Май_R', 'Июн_R', 'Июл_R', 'Авг_R', 'Сен_R', 'Окт_R', 'Ноя_R', 'Дек_R', 'Year_Sum_R', 'Sum_Mar-May_R', 'Sum_Jun-Jul_R', 'Sum_Sen-Nov_R'])
					 