import pandas as pd
import os
import openpyxl

# Заводим рабочую папку
os.chdir("C:\\Users\\OMEN\\Desktop\\DF\\HC\\new_test")

# Считываем эксели
df_tableau = pd.read_excel('Tableau_W202021.xlsx')
df_cfo = pd.read_excel('CFO_W202021.xlsx')
df_joblist = pd.read_excel('joblistW202021.xlsx')
df_week = pd.read_excel('week202021.xlsx', sheet_name="ТС5")

# Убираем НаН
df_tableau.dropna(subset=['Должность Полное Наименование'])
df_tableau.fillna(0)

# Удаляем лишние столбцы
df_tableau = df_tableau.drop(['Week number (new)', 'Day of Дата'], axis=1)

# Группируем по признакаом ЦФО и Должность
df_tableau = df_tableau.groupby(['ЦФО ID', 'Должность Полное Наименование'], as_index=False)[['Факт часов СП+АК для комплектности',
                                                                              'План часов Вак',
                                                                              'План часов СП',
                                                                              'Утверждено ШД',
                                                                              'Оформлено ШД']] .sum()

# Заводим новый столбце с ключем
df_tableau['План часов всего'] = (df_tableau['План часов СП'] + df_tableau['План часов Вак'])

# Лефт джоин двух датафреймов
df_tableau = df_tableau.merge(df_joblist, on='Должность Полное Наименование', how='left')

# Заполняем НаН нулями
df_tableau['СП'] = df_tableau['СП'].fillna(0)
df_tableau['АУТ'] = df_tableau['АУТ'].fillna(0)

# Расставляем признаки для фильтрации
df_tab_sp = df_tableau.loc[df_tableau['СП'] == 1]
df_tab_sp.rename(columns={"Факт часов СП+АК для комплектности": "Факт часов СП"})
df_tab_out = df_tableau.loc[df_tableau['АУТ'] == 1]
df_tab_out.rename(columns={'Факт часов СП+АК для комплектности': 'Факт часов АУП'})

# Делаем слияние
df_tab_all = pd.concat([df_tab_sp, df_tab_out])
df_tab_all.fillna(0)

# Заводим ключ
df_tab_all['Ключ'] = df_tab_all['ЦФО ID'] + df_tab_all['Должность Полное Наименование']

# Переименовываем колонки
df_tab_all = df_tab_all.rename(columns={'Должность Полное Наименование': 'Должность'})
df_tab_all = df_tab_all.rename(columns={'ЦФО ID': 'Центр финансовой отчетности'})
df_tab_all = df_tab_all.rename(columns={'Факт часов СП': 'Фактически отработанное время без переработок (часов)'})
df_tab_all = df_tab_all.rename(columns={'Факт часов АУП': 'Фактическое время внешнего персонала (часов)'})
df_tab_all = df_tab_all.rename(columns={'Утверждено ШД': 'Утверждено по штатному расписанию'})
df_tab_all = df_tab_all.rename(columns={'Оформлено ШД': 'Оформлено по штатному расписанию'})
df_tab_all = df_tab_all.rename(columns={'План часов всего': 'Нормативное время с учетом вакансий (часов)'})

# Удаляем лишние столбцы
df_tab_all = df_tab_all.drop(['План часов Вак', 'План часов СП', 'СП', 'АУТ'], axis=1)

# Заводим пустые столбцы для ДФ

df_tab_all['Признак НЕ открытого магазина'] = 0
df_tab_all['Признак НЕ стажера НЕ Итого НЕ Объекты розница'] = 0
df_tab_all['Признак только стажер'] = 0
df_tab_all['Базовые должности розницы (Дивизион)'] = 0
df_tab_all['Признак НЕ Итого НЕ Объекты розница'] = 0

# Лефт джоин новых основной ДФ и справочник ЦФО

df_tab_all = df_tab_all.merge(df_cfo, left_on='Центр финансовой отчетности', right_on='ЦФО ID', how='left')

df_tab_all = df_tab_all.rename(columns={'ЦФО Наименование': 'Магазин',\
                                        'ЦФО Дивизион/гр.кластеров': 'Группа кластеров',\
                                        'ЦФО Кластер': 'Кластер'})

df_week['Ключ']=df_week['Центр финансовой отчетности']+df_week['Должность']

df_result = df_tab_all.merge(df_week, left_on='Ключ', right_on='Ключ', how='left')

df_result.drop_duplicates(subset='Ключ', keep = False, inplace=True)

df_result['Путь'] = 0
#df_result_shape = df_result[[]]

df_result.to_excel("C:\\Users\\OMEN\\Desktop\\DF\\HC\\new_test\\output.xlsx")
