import pandas as pd
import statsmodels.api as sm
from sklearn.linear_model import LinearRegression
import numpy as np

import os

pd.set_option('display.max_columns', 200) #replace n with the number of columns you want to see completely
pd.set_option('display.max_rows', 200) #replace n with the number of rows you want to see completely

def pandas_insert(df, idx, row_contents):
    top = df.iloc[:idx]
    bot = df.iloc[idx:]
    inserted = pd.concat([top, row_contents, bot])
    return inserted

def filter_years(year, thresh=2020, type='lower'):
    if type == 'lower':
        condition = year <= thresh
    elif type == 'higher':
        condition = year >= thresh
        
    if condition:
        return True
    else:
        return False

def write_string_to_excel(writer, message, sheet_name, startrow):
    message_str = pd.Series(message)
    message_str.to_excel(writer, sheet_name=sheet_name, header=False, index=False, startrow=startrow)

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
        
#Площадь земельного фонда в границах территорий Российской Федерации  https://www.fedstat.ru/indicator/38150
data_lands = pd.read_excel('source_data/RF_lands.xlsx', sheet_name='Данные', header=2)
data_lands.columns = ['Территория', 'Тип угодий', 'Единица измер', 'Years', 'Площадь земельного фонда']
data_lands = data_lands.set_index('Years')

#ПЛОЩАДИ СХ КУЛЬТУР https://www.fedstat.ru/indicator/31328
data_area = pd.read_excel('source_data/RF_squares.xls', sheet_name='Данные', header=1)
data_area1 = pd.read_excel('source_data/RF_squares.xls', sheet_name='Данные 2', header=1)
data_area = data_area.append(data_area1).drop(index=0)
data_area.columns = ['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Years', 'Посевные площади']
data_area = data_area.set_index('Years')

#СБОР СХ КУЛЬТУР 
data_sbor = pd.read_excel('source_data/RF_sbor.xls', sheet_name='Данные', header=1)
data_sbor1 = pd.read_excel('source_data/RF_sbor.xls', sheet_name='Данные 2', header=1)
data_sbor2 = pd.read_excel('source_data/RF_sbor.xls', sheet_name='Данные 3', header=1)
data_sbor3 = pd.read_excel('source_data/RF_sbor.xls', sheet_name='Данные 4', header=1)
data_sbor = pd.concat((data_sbor, data_sbor1, data_sbor2, data_sbor3), axis=0)
data_sbor = data_sbor.append(data_sbor1).drop(index=0)
data_sbor.columns = ['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Years', 'Валовый сбор']
data_sbor = data_sbor.set_index('Years')      
    
#УДОБРЕНИЯ НА 1 га https://www.fedstat.ru/indicator/30964
data_ferts = pd.read_excel('source_data/RF_ferts.xls', sheet_name='Данные', header=1)
data_ferts = data_ferts.drop(index=0)
data_ferts.columns = ['Территория', 'Культура', 'Единица измер', 'Years', 'Внесено мин удобр на 1 га']
data_ferts = data_ferts.set_index('Years')
    
#УДОБРЕНИЯ ОРГАНИЧЕСКИЕ НА 1 га https://www.fedstat.ru/indicator/30966
data_org_ferts = pd.read_excel('source_data/RF_org_ferts.xls', sheet_name='Данные', header=1)
data_org_ferts = data_org_ferts.drop(index=0)
data_org_ferts.columns = ['Территория', 'Культура', 'Единица измер', 'Years', 'Внесено орг удобр на 1 га']
data_org_ferts= data_org_ferts.set_index('Years')
    
#ЭНЕРГЕТИЧЕСКИЕ МОЩНОСТИ НА 100 га  https://www.fedstat.ru/indicator/31632
data_power = pd.read_excel('source_data/RF_power_on100ha.xls', sheet_name='Данные', header=1)
data_power = data_power.drop(index=0)
data_power.columns = ['Территория', 'Единица измер', 'Years', 'Мощности на 100 га']
data_power = data_power.set_index('Years')

#НАЛИЧИЕ ТЕХНИКИ по видам  https://www.fedstat.ru/indicator/31632
data_techs = pd.read_excel('source_data/RF_technics.xls', sheet_name='Данные', header=1)
data_techs = data_techs.drop(index=0)
data_techs.columns = ['Территория', 'Вид техники', 'Единица измер', 'Years', 'Единиц техники шт']
data_techs = data_techs.set_index('Years')

#УРОЖАЙНОСТИ https://fedstat.ru/indicator/31533
data_yield = pd.read_excel('source_data/RF_yield.xls', sheet_name='Данные', header=1)
data_yield1 = pd.read_excel('source_data/RF_yield.xls', sheet_name='Данные 2', header=1)
data_yield = pd.concat((data_yield, data_yield1), axis=0)
data_yield = data_yield.append(data_yield1).drop(index=0)
data_yield.columns = ['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Years', 'Урожайность ц/га']
data_yield = data_yield.set_index('Years')

#Метео данные
meteo_Stations = pd.read_excel('source_data/RF_meteo_stations_list.xlsx', sheet_name='Станции_области', header=0, usecols='A:C')

#Осадки%
meteo_R = pd.read_excel('source_data/RF_meteo_R.xlsx', sheet_name='R_mes', header=0, usecols='A:R')
meteo_R = meteo_R.merge(meteo_Stations, on=['Номер станции'], how='right')
meteo_R = meteo_R.set_index('Год')

#Температура
meteo_T = pd.read_excel('source_data/RF_meteo_T.xlsx', sheet_name = 'Лист1', header=0, usecols='A:AH')
meteo_T = meteo_T.merge(meteo_Stations, on=['Номер станции'], how='right')
meteo_T = meteo_T.set_index('Год')

#Списки данных
lands = data_lands['Тип угодий'].unique()  # угодья
orgs = data_area['Тип хозяйств'].unique()
crop_list_land = data_area['Культура'].unique()
crop_list_sbor = data_sbor['Культура'].unique()
crop_list_ferts = data_ferts['Культура'].unique()
techs_list = data_techs['Вид техники'].unique()
yield_list = data_yield['Культура'].unique()
oblasts_T = meteo_T['Область'].unique()
oblasts_R = meteo_R['Область'].unique()


#Выбор регионов для расчёта
areas = data_area['Территория'].unique()[2:19] # Нужно указывать районы цифрами
# areas = list(map(lambda area: ' ' * 8 + area, ['Белгородская область'])) # Нужно указывать районы текстом (без пробелов в начале)

if type(areas) == str:
    areas_names = areas.strip()
    areas = [areas]
else:
    areas_names = ', '.join(areas).replace(' '*8, '')
print('Areas: ' + areas_names)

# Создание папки с выходными данными
output_path = './output'
try:
    os.mkdir(output_path)
    print('Creating the ' + output_path + ' dir and put your data in it.')
except:
    print(output_path + ' dir was found. Your data will be in it.')

# Создание итоговой таблицы
file_path = './output/AGRO_table_CFO.xlsx'
writer = pd.ExcelWriter(file_path)

# Создание файла с ошибками
err_file = open('./output/agro_err.txt', 'w')

# Переменная сводной таблицы по всем регионам
pivot = pd.DataFrame()

for area in areas:
    
    area_name = area.strip()

    sheet_name = ''
    if len(area) < 30:
        sheet_name = area
    else:
        sheet_name = area[31]

    area_name_table = pd.Series(area_name)
    area_name_table.to_excel(writer, sheet_name=sheet_name, header=False, index=False, startrow=0, startcol=1)

    print('\n' + 'Current area: ' + area + '\n')

    #ПЛОЩАДИ ЗЕМЕЛЬ https://www.fedstat.ru/indicator/38150
    land_all = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Общая площадь')]
    [['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_buildings = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Земли застройки')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]

    land_farmland = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Сельскохозяйственные угодья (всего)')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_arable = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Пашня')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_sown_area = data_area.loc[(data_area['Территория'] == area) & (data_area['Тип хозяйств']=='Хозяйства всех категорий')& (data_area['Культура']=='Вся посевная площадь')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Посевные площади']]
    land_arable['Пары'] = land_arable['Площадь земельного фонда']-land_sown_area['Посевные площади']
    land_arable['Доля паров'] = land_arable['Пары']/land_arable['Площадь земельного фонда']

    land_swamps = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Болота')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_forest = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Лесные площади')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_forest_other = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Лесные насаждения, не входящие в лесной фонд')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_depos = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Залежь')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_pastures = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Пастбища')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_hay = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Сенокосы')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_perr_plant = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Многолетние насаждения')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_break = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Нарушенные земли')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    land_waters = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Под водой')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]

    land_area = pd.concat((land_all['Территория'],
                        land_all['Площадь земельного фонда']/1000,
                        land_farmland['Площадь земельного фонда']/1000,
                        land_arable['Площадь земельного фонда']/1000,
                        land_sown_area['Посевные площади']/1000,
                        land_arable['Пары']/1000,
                        land_arable['Доля паров'],
                        land_hay['Площадь земельного фонда']/1000,
                        land_pastures['Площадь земельного фонда']/1000,
                        land_perr_plant['Площадь земельного фонда']/1000,
                        land_buildings['Площадь земельного фонда']/1000,
                        ), axis=1)
    land_area.columns = ['Территория', 'Общая площадь, млн га', 'СХ угодья, млн га', 'Пашня, млн га', 'Посевная площадь, млн. га', 'Пары, млн га', 
                      'Доля паров', 'Сенокосы, млн га', 'Пастбища, млн га', 'Многолетние насаждения, млн га', 'Земли застройки, млн га']

    #Зерновые и зернобобовые
    zern_zernbob_area = data_area.loc[(data_area['Территория'] == area) & (data_area['Тип хозяйств']=='Хозяйства всех категорий') & (data_area['Культура']=='Зерновые и зернобобовые культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Посевные площади']]
    zern_zernbob_prod = data_sbor.loc[(data_sbor['Территория'] == area) & (data_sbor['Тип хозяйств']=='Хозяйства всех категорий') & (data_sbor['Культура']=='Зерновые и зернобобовые культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Валовый сбор']]
    zern_zernbob_yield = zern_zernbob_prod['Валовый сбор']/zern_zernbob_area['Посевные площади']

    #Зернобобовые
    zern_bob_area = (data_area.loc[(data_area['Территория'] == area) & (data_area['Тип хозяйств']=='Хозяйства всех категорий') & (data_area['Культура']=='Зернобобовые культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Посевные площади']])
    zern_bob_prod = data_sbor.loc[(data_sbor['Территория'] == area) & (data_sbor['Тип хозяйств']=='Хозяйства всех категорий') & (data_sbor['Культура']=='Зернобобовые культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Валовый сбор']]
    zern_bob_yield = zern_bob_prod['Валовый сбор']/zern_bob_area['Посевные площади']
    zern_bob_data = pd.concat((zern_bob_area,
                               zern_bob_prod,
                               zern_bob_yield/10
                               ), axis=1)

    #Картофель
    potat_area = data_area.loc[(data_area['Территория'] == area) & (data_area['Тип хозяйств']=='Хозяйства всех категорий')& (data_area['Культура']=='Картофель - всего')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Посевные площади']]
    #Овощи
    veget_area = data_area.loc[(data_area['Территория'] == area) & (data_area['Тип хозяйств']=='Хозяйства всех категорий')& (data_area['Культура']=='Овощи открытого грунта')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Посевные площади']]

    #вычисляем зерновые
    zern_area = pd.concat((land_sown_area.iloc[:,4]/1000,
                        zern_zernbob_area.iloc[:,4]/1000,
                        zern_bob_area.iloc[:,4]/1000,
                        potat_area.iloc[:,4]/1000,
                        veget_area.iloc[:,4]/1000), axis=1)
    zern_area.columns = ['Посевная площадь млн га', 'Зерн и зернобобовые млн га', 'Зернобобовые млн га', 
                      'Картофель млн га', 'Овощи открытого грунта млн га']
    zern_area['Зерновые млн га'] = zern_area['Зерн и зернобобовые млн га']-zern_area['Зернобобовые млн га']
    zern_area['Зернобобовые доля'] = zern_area['Зернобобовые млн га']/zern_area['Зерн и зернобобовые млн га']

    zern_prod =  pd.concat((zern_bob_prod.iloc[:,0],
                        zern_zernbob_prod.iloc[:,4]/10000,
                        zern_bob_prod.iloc[:,4]/10000), axis=1)
    zern_prod.columns = ['Территория', 'Зерн и зернобобовые млн т', 'Зернобобовые млн т']
    zern_prod['Зерновые млн т'] = zern_prod['Зерн и зернобобовые млн т']-zern_prod['Зернобобовые млн т']

    #Категории земель

    #lands_all = data_lands.loc[(data_lands['Территория'] == area) & (data_lands['Тип угодий']=='Общая площадь')][['Территория', 'Тип угодий', 'Единица измер', 'Площадь земельного фонда']]
    #agro_data['Доля паров'] = land_arable['Доля паров']

    agro_data = pd.concat((land_area,
                        zern_area.iloc[:,1:7], 
                        zern_prod.iloc[:,1:4]), axis=1)
    agro_data['Зерновые т/га'] = agro_data['Зерновые млн т']/agro_data['Зерновые млн га']
    agro_data['Зерн и зернобоб т/га'] = zern_zernbob_yield/10
    agro_data['Зернобоб т/га'] = zern_bob_yield/10

    #картофель
    potat_area = data_area.loc[(data_area['Территория'] == area) & (data_area['Тип хозяйств']=='Хозяйства всех категорий') & (data_area['Культура']=='Картофель - всего')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Посевные площади']]
    potat_prod = data_sbor.loc[(data_sbor['Территория'] == area) & (data_sbor['Тип хозяйств']=='Хозяйства всех категорий') & (data_sbor['Культура']=='Картофель - всего')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Валовый сбор']]
    potat_yield = potat_prod['Валовый сбор']/10/potat_area['Посевные площади'] 

    #Масличные
    oilscrops_area = data_area.loc[(data_area['Территория'] == area) & (data_area['Тип хозяйств']=='Хозяйства всех категорий') & (data_area['Культура']=='Масличные культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Посевные площади']]
    oilscrops_prod = data_sbor.loc[(data_sbor['Территория'] == area) & (data_sbor['Тип хозяйств']=='Хозяйства всех категорий') & (data_sbor['Культура']=='Масличные культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Валовый сбор']]
    oilscrops_yield = oilscrops_prod['Валовый сбор']/oilscrops_area['Посевные площади']

    agro_data ['Масличные млн га'] = oilscrops_area['Посевные площади']/1000
    agro_data ['Масличные млн т'] = oilscrops_prod['Валовый сбор']/10000
    agro_data ['Масличные т/га'] = oilscrops_yield

    agro_data['Пашня исп млн га'] = (agro_data['Зерн и зернобобовые млн га']+agro_data['Картофель млн га']+agro_data['Масличные млн га']+agro_data['Овощи открытого грунта млн га'])

    #Овощи
    veget_otkr_zakr_prod = data_sbor.loc[(data_sbor['Территория'] == area) & (data_sbor['Тип хозяйств']=='Хозяйства всех категорий') & (data_sbor['Культура']=='Овощи открытого и закрытого грунта')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Валовый сбор']]
    veget_zakr_prod = data_sbor.loc[(data_sbor['Территория'] == area) & (data_sbor['Тип хозяйств']=='Хозяйства всех категорий') & (data_sbor['Культура']=='Овощи закрытого грунта')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Валовый сбор']]

    agro_data ['Картофель т/га'] = potat_yield
    #	agro_data ['Подсолнечник т/га'] = sunflow_yield
    agro_data ['Овощи млн т'] = veget_otkr_zakr_prod.iloc[:,4]/10000

    #Урожайности из отдельного файла
    zenrbob1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Зернобобовые культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]
    zernzenrbob1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Зерновые и зернобобовые культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]
    zernbkuk1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Зерновые культуры (без кукурузы)')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]
    potat1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Картофель - всего')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]
    corn1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Кукуруза на зерно')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]
    cornsylos1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Кукуруза на силос, зеленый корм и сенаж (вес зеленой массы)')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]
    oils1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Масличные культуры')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]
    sunflow1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Подсолнечник')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]
    vegsopen1_y = data_yield.loc[(data_yield['Территория'] == area) & (data_yield['Тип хозяйств']=='Хозяйства всех категорий') & (data_yield['Культура']=='Овощи открытого грунта')][['Территория', 'Тип хозяйств', 'Культура', 'Единица измер', 'Урожайность ц/га']]

    agro_data['Зернобоб 1 т/га'] = zenrbob1_y.iloc[:,4]/10
    agro_data['Зерн и зернобоб 1 т/га'] = zernzenrbob1_y.iloc[:,4]/10
    #agro_data['Зерн без кукурузы 1 т/га'] = zernbkuk1_y.iloc[:,4]/10
    #agro_data['Картофель 1 т/га'] = zernbkuk1_y.iloc[:,4]/10
    agro_data['Кукуруза на зер 1 т/га'] = corn1_y.iloc[:,4]/10
    #agro_data['Кукуруза на силос 1 т/га'] = cornsylos1_y.iloc[:,4]/10
    agro_data['Масличные 1 т/га'] = oils1_y.iloc[:,4]/10
    agro_data['Подсолнечник 1 т/га'] = sunflow1_y.iloc[:,4]/10
    agro_data['Овощи откр гр 1 т/га'] = vegsopen1_y.iloc[:,4]/10

    #Удобрения
    ferts_all = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='ВСЕГО')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_shall = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Сельскохозяйственные культуры - всего')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]

    ferts_zernzernbob = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Зерновые и зернобобовые культуры (без кукурузы)')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_zern = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Зерновые культуры (без кукурузы)')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_potat = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Картофель')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_veget = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Овощи')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_vegetall = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Овощи - всего')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_vegetallbahch = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Овощи и бахчевые культуры')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_suflow = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Подсолнечник')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_sugarsv = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Сахарная свекла')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]
    ferts_sh_all = data_ferts.loc[(data_ferts['Территория'] == area) & (data_ferts['Культура']=='Сельскохозяйственные культуры - всего')][['Территория', 'Культура', 'Единица измер', 'Внесено мин удобр на 1 га']]

    agro_data['МУ ВСЕГО до 2019'] = ferts_all.iloc[:,3]
    agro_data['МУ ВСЕГО с 2020'] = ferts_shall.iloc[:,3]
    agro_data['МУ ВСЕГО'] = agro_data['МУ ВСЕГО до 2019'].fillna(agro_data['МУ ВСЕГО с 2020'])

    agro_data['МУ Зерновые и зернобобовые'] = ferts_zernzernbob.iloc[:,3]
    agro_data['МУ Зерновые'] = ferts_zern.iloc[:,3]
    agro_data['МУ Зерновые плюс зернобоб'] = agro_data['МУ Зерновые и зернобобовые'].fillna(agro_data['МУ Зерновые'])

    agro_data['МУ Картофель'] = ferts_potat.iloc[:,3]
    agro_data['МУ Овощи с 2020'] = ferts_veget.iloc[:,3]
    agro_data['МУ Овощи до 2019'] = ferts_vegetall.iloc[:,3]
    agro_data['МУ Овощи'] = agro_data['МУ Овощи с 2020'].fillna(agro_data['МУ Овощи до 2019'])

    #agro_data['МУ Овощи и бахчевые'] = ferts_vegetallbahch.iloc[:,3]
    agro_data['МУ Подсолнечник'] = ferts_suflow.iloc[:,3]
    agro_data['МУ Сахарная свекла'] = ferts_sugarsv.iloc[:,3]

    #Техника
    # 	techs_list1 = data_techs['Вид техники'].value_counts()
    # 	
    # 	techs_traks = data_techs.loc[(data_techs['Территория'] == area) & (data_techs['Вид техники']=='Тракторы')][['Территория', 'Вид техники', 'Единица измер', 'Единиц техники шт']]
    # 	techs_trakskols = data_techs.loc[(data_techs['Территория'] == area) & (data_techs['Вид техники']=='Тракторы  колесные общего назначения')][['Территория', 'Вид техники', 'Единица измер', 'Единиц техники шт']]
    # 	techs_oprysk = data_techs.loc[(data_techs['Территория'] == area) & (data_techs['Вид техники']=='Опрыскиватели')][['Территория', 'Вид техники', 'Единица измер', 'Единиц техники шт']]
    # 	techs_kartkop = data_techs.loc[(data_techs['Территория'] == area) & (data_techs['Вид техники']=='Копатели  картофеля')][['Территория', 'Вид техники', 'Единица измер', 'Единиц техники шт']]
    # 	techs_minudobr = data_techs.loc[(data_techs['Территория'] == area) & (data_techs['Вид техники']=='Машины для минеральных удобрений')][['Территория', 'Вид техники', 'Единица измер', 'Единиц техники шт']]
    # 	techs_poskompl = data_techs.loc[(data_techs['Территория'] == area) & (data_techs['Вид техники']=='Посевные комплексы для зерна')][['Территория', 'Вид техники', 'Единица измер', 'Единиц техники шт']]
    # 	techs_seyalki = data_techs.loc[(data_techs['Территория'] == area) & (data_techs['Вид техники']=='Сеялки')][['Территория', 'Вид техники', 'Единица измер', 'Единиц техники шт']]

    # 	agro_data['Тракторы'] = techs_traks.iloc[:,3]
    # 	agro_data['Тракторы колесные'] = techs_trakskols.iloc[:,3]
    # 	agro_data['Опрыскиватели'] = techs_oprysk.iloc[:,3]
    # 	agro_data['Копатели картофеля'] = techs_kartkop.iloc[:,3]
    # 	agro_data['Машины для мин удобр'] = techs_minudobr.iloc[:,3]
    # 	agro_data['Посев компл для зерн'] = techs_poskompl.iloc[:,3]
    # 	agro_data['Сеялки'] = techs_seyalki.iloc[:,3]
    # 	agro_data['Тракторы шт/1000 га'] =techs_traks.iloc[:,3]/(agro_data.iloc[:,1]*1000)
    # 	agro_data['Опрыскиватели шт/1000 га'] =techs_oprysk.iloc[:,3]/(agro_data.iloc[:,1]*1000)
    # 	agro_data['Машины для мин удобр шт/1000 га'] = techs_minudobr.iloc[:,3]/(agro_data.iloc[:,1]*1000)
    # 	agro_data['Посев компл для зерн шт/1000 га'] =  techs_poskompl.iloc[:,3]/(agro_data.iloc[:,2]*1000)
    # 	agro_data['Сеялки шт/1000 га'] = techs_seyalki.iloc[:,3]/(agro_data.iloc[:,2]*1000)

    #Энергетические мощности
    power = data_power.loc[(data_power['Территория'] == area)][['Территория', 'Единица измер', 'Мощности на 100 га']]


    agro_data['Энерг мощ лс/100га'] = power.iloc[:,2]

    #Метеоданные
    rrr = meteo_R.loc[(meteo_R['Область'] == area)][['Янв_R', 'Фев_R', 'Мар_R', 'Апр_R', 
                                                  'Май_R', 'Июн_R', 'Июл_R', 'Авг_R', 
                                                  'Сен_R', 'Окт_R', 'Ноя_R', 'Дек_R', 
                                                  'Year_Sum_R', 'Sum_Mar-May_R', 'Sum_Jun-Jul_R', 'Sum_Sen-Nov_R']]
    rrr = pd.pivot_table(rrr, index = 'Год', aggfunc = 'mean').reindex(columns = ['Янв_R', 'Фев_R', 'Мар_R', 'Апр_R', 
                                                                                    'Май_R', 'Июн_R', 'Июл_R', 'Авг_R', 'Сен_R', 
                                                                                    'Окт_R', 'Ноя_R', 'Дек_R', 'Year_Sum_R', 
                                                                                    'Sum_Mar-May_R', 'Sum_Jun-Jul_R', 'Sum_Sen-Nov_R'])

    rrr_mean = rrr.mean().reset_index()
    rrr_mean.columns = ['Parameters', 'Values']
    rrr_max = rrr.max().reset_index()
    rrr_max.columns = ['Parameters', 'Values']

    temp = meteo_T.loc[(meteo_T['Область'] == area)][['Янв_t', 'Фев_t', 'Мар_t', 'Апр_t', 
                                                    'Май_t', 'Июн_t', 'Июл_t', 'Авг_t', 
                                                    'Сен_t', 'Окт_t', 'Ноя_t', 'Дек_t', 
                                                    'T_год_ср', 'T_год_сумм', 'Mean_Mar-May_T', 
                                                    'Mean_Jun-Jul_T', 'Sum_Mar_Aug_T', 'ГТК']]
    temp = pd.pivot_table(temp, index = 'Год', aggfunc = 'mean').reindex(columns = ['Янв_t', 'Фев_t', 'Мар_t', 'Апр_t', 
                                                                                    'Май_t', 'Июн_t', 'Июл_t', 'Авг_t', 'Сен_t', 
                                                                                    'Окт_t', 'Ноя_t', 'Дек_t', 'T_год_ср', 
                                                                                    'T_год_сумм', 'Mean_Mar-May_T', 'Mean_Jun-Jul_T', 'Sum_Mar_Aug_T', 'ГТК'])

    temp_mean = temp.mean().reset_index()
    temp_mean.columns = ['Parameters', 'Values']
    temp_max = temp.max().reset_index()
    temp_max.columns = ['Parameters', 'Values']

    target_x = pd.read_excel('source_data/Target_values.xlsx', header=0)
    target_x = pd.concat((target_x, rrr_mean, temp_mean), axis=0)
    target_x.set_index('Parameters', inplace=True)

    agro_data = pd.concat((agro_data, temp, rrr), axis=1)

    agro_data.columns.tolist() #Это я смотрю список колонок

    years_list = list(filter(filter_years, agro_data.index.tolist()))
    agro_data= agro_data.loc[years_list]

    X_data_raw_0 = agro_data[['МУ ВСЕГО', 
                        'МУ Зерновые и зернобобовые', 
                        'МУ Зерновые', 
                        'МУ Зерновые плюс зернобоб',
                        'МУ Картофель',
                        'МУ Овощи',
                        'МУ Подсолнечник',
                        'Энерг мощ лс/100га', 'ГТК']] # Предварительный список

    X_data_raw = X_data_raw_0[['МУ ВСЕГО', 
                        'МУ Зерновые плюс зернобоб',
                        'МУ Картофель',
                        'МУ Овощи',
                        'МУ Подсолнечник',
                        'Энерг мощ лс/100га', 'ГТК']]

    crops_list = ['Зерновые т/га',
                'Зерн и зернобоб т/га',
                'Масличные т/га',
                'Масличные 1 т/га', 
                'Зернобоб 1 т/га',
                'Зерн и зернобоб 1 т/га',
				'Картофель т/га',
				'Овощи откр гр 1 т/га']


    #Убираем столбцы, полностью состоящие из Nan (Подсолнечник, например, во многих местах), 
    #и также удаляем столбцы, в которых ячеек с данными (не Nan) больше 10
    X_data_raw = X_data_raw.dropna(axis=1, how='all', thresh=10)

    # Расчитываем среднее по с/х
    area_agriculture = X_data_raw[['МУ ВСЕГО', 
                                    'МУ Зерновые плюс зернобоб',
                                    'МУ Картофель',
                                    'МУ Овощи',
                                    'Энерг мощ лс/100га', 'ГТК']].mean()

    # Считаем ГТК

    # Делаем датафрейм с температурами и осадками за год
    precip_temp = rrr.iloc[:,:-4].join(temp.iloc[:,:-5]) 

    # Берём данные до 1990 года
    precip_temp_years_list = list((item for item in precip_temp.index.tolist() if filter_years(item, 1990, type='higher'))) 
    precip_temp = precip_temp.loc[precip_temp_years_list]

    # Рассчитываем сумму осадков при температуре выше 10, и сумму температур выше 10
    precipitation_sum = 0
    temp_sum = 0
    for month in ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]:
        precipitation_sum += precip_temp.query(month + '_t' + ' > 10')[month + '_R'].sum()
        temp_sum += precip_temp.query(month + '_t' + ' > 10')[month + '_t'].sum()

    # Вычисляем ГТК
    gtk = precipitation_sum / (temp_sum * 3)

    print(f'{area_name}, сумма осадков: {precipitation_sum}')
    print(f'{area_name}, сумма температур: {temp_sum}')
    print(f'{area_name}, ГТК: {gtk}')

    # Делаем сводную таблицу
    area_pivot = pd.DataFrame(area_agriculture)
    area_pivot['Области'] = area_pivot.index
    area_pivot = area_pivot.pivot_table(columns='Области')
    area_pivot['T > 10'] = temp_sum
    area_pivot['R при T > 10'] = precipitation_sum
    area_pivot['ГТК'] = gtk
    area_pivot.index = [area_name]
    
    startrow = 1
    for crop in crops_list:
        print('\n' + 'Current crop: ' + crop)
        model = LinearRegression()
        y = agro_data[[crop]]
        current_crop_data = pd.concat((X_data_raw, y), axis=1)
        current_crop_data = current_crop_data.replace([np.inf, -np.inf], np.nan)
        #current_crop_data = current_crop_data.dropna(axis=1)
        current_crop_data = current_crop_data.dropna(axis=0)
        y = current_crop_data[[crop]]
        counted_years = current_crop_data.index.tolist()

        if len(counted_years) == 0:
            write_string_to_excel(writer, crop, sheet_name, startrow)
            write_string_to_excel(writer, 'There is no data on ' + crop, sheet_name, startrow+1)
            startrow += 2

            err_file.write(area + ', ' + crop + ': ' + 'no information on crop' '\n\n')
            continue

        X_data = current_crop_data[['МУ ВСЕГО', 
                     'МУ Зерновые плюс зернобоб',
                     'МУ Картофель',
                    'МУ Овощи',
                     'Энерг мощ лс/100га']]
#         print(y)
#         print(X_data)

        all_coefs_valid = False
        pvalues_valid = False
        enable_second_condition = False
        statistic_error = False
        incorrect_coefs = []
        iteration = 0
        beta_pv_result = pd.DataFrame()
        statistic_values_result = pd.DataFrame()
        while all_coefs_valid == False or pvalues_valid == False:
            try:
                result = sm.OLS(y, X_data).fit()
                summary_df = result.summary2().tables[1]
            except Exception as e:
                print(e)
                title = pd.Series(crop)
                error = pd.Series(e)
                title.to_excel(writer, sheet_name=sheet_name, header=False, index=False, startrow=1+startrow)
                error.to_excel(writer, sheet_name=sheet_name, header=False, index=False, startrow=2+startrow)
                startrow = startrow - 22
                err_file.write(area + ', ' + crop + ': ' + str(e) + '\n\n')
                statistic_error = True
                break

            beta_pv_df = summary_df[['Coef.', 'P>|t|']]
            beta_pv_df = beta_pv_df.rename(columns={'Coef.': 'beta', 'P>|t|': 'P_value'})
            potencial_x = target_x.iloc[:,0]*beta_pv_df.iloc[:,0]
            potencial = potencial_x.sum()

            if beta_pv_result.empty:
                beta_pv_result = beta_pv_df
            else:
                beta_pv_result = pd.merge(beta_pv_result, beta_pv_df, left_index=True, right_index=True, how='left', 
                                            suffixes=("-Iteration " + str(iteration), "-Iteration " + str(iteration + 1)))

            if beta_pv_df[beta_pv_df['beta'] < 0]['P_value'].empty:
                max_pv_with_negative_beta = False
            else:
                max_pv_with_negative_beta = max(beta_pv_df[beta_pv_df['beta'] < 0]['P_value'])

            if all_coefs_valid:
                enable_second_condition = True    
                
            # Условие один
            for row_number, row_name in enumerate(beta_pv_df.index):
                current_beta = beta_pv_df.loc[row_name, 'beta']
                current_p_value = beta_pv_df.loc[row_name, 'P_value']
                if len(beta_pv_df.index) == 1:
                    all_coefs_valid = True
                elif row_name == 'const':
                    continue
                elif current_beta < 0 or current_p_value > 0.05 and current_p_value == max_pv_with_negative_beta and max_pv_with_negative_beta:
                    incorrect_coefs.append([row_name, row_number])
                    del X_data[row_name]
                    break
                elif row_number == len(beta_pv_df.index)-1:
                    all_coefs_valid = True
                    
            # Условие два
            if enable_second_condition:
                for row_number, row_name in enumerate(beta_pv_df.index):
                    current_beta = beta_pv_df.loc[row_name, 'beta']
                    current_p_value = beta_pv_df.loc[row_name, 'P_value']
                    if len(beta_pv_df.index) == 1:
                        pvalues_valid = True
                    elif row_name == 'const':
                        continue
                    elif current_beta > 0 and current_p_value > 0.05:
                        incorrect_coefs.append([row_name, row_number])
                        del X_data[row_name]
                        break
                    elif row_number == len(beta_pv_df.index)-1:
                        pvalues_valid = True

            statistic_values = pd.DataFrame({
                'Статистика': ["R2", "R2 adj.", "F", "Significance F", "Cond. No.", "Потенциал"],
                "Iteration "+str(iteration+1): [result.rsquared, 
                                 result.rsquared_adj, 
                                 result.fvalue, 
                                 result.f_pvalue, 
                                 result.condition_number, potencial],
                    })

            if statistic_values_result.empty:
                statistic_values_result = statistic_values
            else:
                statistic_values_result = pd.merge(statistic_values_result, statistic_values, 
                                                    left_on='Статистика', right_on='Статистика', how='left',
                                                    suffixes=('-Iteration ' + str(iteration), '-Iteration ' + str(iteration + 1)))
                potenc = statistic_values_result.transpose().reset_index().drop(['index'], axis=1)
                potenc.columns = potenc.iloc[0]
                potenc.drop([0], inplace=True)
            iteration += 1
            
        if statistic_error:
            continue
            
        # Расчёт квантилей для Зерновых и добавление их в сводную таблицу
        if crop == 'Зерновые т/га':
            cereals_potencials= pd.Series(statistic_values_result.iloc[5].values[1:])
            quantiles = cereals_potencials.quantile([0.05, 0.95])
            quant_value = quantiles.mean()
            area_pivot['Агропотенциал ц/га'] = quant_value

        if len(beta_pv_result.columns)/2 % 2:
            beta_pv_result = beta_pv_result.rename(columns={'beta': 'beta-Iteration ' + str(iteration), 
                                                            'P_value': 'P_value-Iteration ' + str(iteration)})

        beta_pv_result.columns = pd.MultiIndex.from_tuples([tuple([c.split("-")[1], c.split("-")[0]]) for c in beta_pv_result.columns])

        for column_num in range(len(statistic_values_result.columns)*2-1):
            if column_num % 2:
                statistic_values_result.insert(column_num, column_num, None)
            elif column_num == 0:
                continue

        variables = ['C' + str(num) for num in range(1, len(beta_pv_result.index) + 1)]
        parameters = beta_pv_result.index
        beta_pv_result.insert(0, 'Variables', variables)
        beta_pv_result.insert(0, 'Parameters', parameters)
        beta_pv_result = beta_pv_result.set_index(['Parameters', 'Variables'])

        print(beta_pv_result)
        print(statistic_values_result)

        write_string_to_excel(writer, crop, sheet_name, startrow)
        startrow += 1
        if len(counted_years) < 54:
            counted_years_str = shorten_years_str(counted_years)
            counted_years_message = 'Calculated on the basis of information in the (' + str(len(counted_years)) + ') ' + 'years: ' + counted_years_str
            write_string_to_excel(writer, counted_years_message, sheet_name, startrow)
            startrow += 1

        beta_pv_result.to_excel(writer, sheet_name=sheet_name, startrow=startrow)
        startrow += beta_pv_result.shape[0] + 4
        statistic_values_result.to_excel(writer, startcol=1, startrow=startrow, index=False, header=False, sheet_name=sheet_name)
        startrow += statistic_values_result.shape[0] + 1
#         target_x.to_excel(writer, sheet_name=sheet_name, startcol=0, startrow=146)
#         cereals_potencials.to_excel(writer, sheet_name=sheet_name, startcol=0, startrow=145, index=False, header=False)
        agro_data.to_excel(writer, sheet_name=sheet_name, startcol=0, startrow=148)

        agro_data_corr = X_data_raw.join(agro_data[crops_list]).dropna(axis=1, thresh=10).dropna(axis=0)
        agro_data_corr.to_csv('./source_data/agro_data_corr.csv')
        agc = pd.read_csv('./source_data/agro_data_corr.csv', index_col=0)
        agro_data_corr = agc.corr()
        agro_data_corr.to_excel(writer, sheet_name=sheet_name, startcol=0, startrow=148 + agro_data.shape[0] + 2)

    
    if pivot.empty:
        pivot = area_pivot
    else:
        pivot = pivot.append(area_pivot)
        
pivot.to_excel('./output/agro_pivot.xlsx')

writer.save()
err_file.close()