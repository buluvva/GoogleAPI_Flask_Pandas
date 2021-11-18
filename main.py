import pandas as pd
import datetime


def fill_table():
    #  combining name surname and patronymic into full name
    table = pd.read_excel('Новая таблица.xlsx', index_col=0)
    name_series = table['Имя'].tolist()
    surname_series = table['Фамилия'].tolist()
    patronym_series = table['Отчество'].tolist()
    fullname_series = table['ФИО'].tolist()
    for i in range(table.shape[0]):
        fullname_series[i] = surname_series[i] + ' ' + name_series[i][0] + '.' + patronym_series[i][0] + '.'
    table['ФИО'] = fullname_series

    #  calculating year earnings and taxes

    salary = table['Зарплата в месяц'].tolist()
    taxes = table['Зарплата в месяц'].tolist()
    sal_n_tax = table['Годовой доход - НДФЛ'].tolist()
    year_tax = table['НДФЛ в год'].tolist()
    for i in range(table.shape[0]):
        salary[i] *= 12
        taxes[i] *= 0.13
        sal_n_tax[i] = salary[i] - taxes[i]
        year_tax[i] = taxes[i] * 12

    table['Годовой доход'] = salary
    table['Годовой доход - НДФЛ'] = sal_n_tax
    table['НДФЛ в год'] = year_tax

    #  calculating age

    date = table['Дата рождения'].copy()  # making copies to avoid loss of data
    age = table['Возраст'].copy()
    is_30 = table['Младше 30?'].copy()
    today = datetime.date.today()
    for i in range(table.shape[0]):
        age[i] = today.year - date[i].year - ((today.month, today.day) < (date[i].month, date[i].day))
        if age[i] < 30:
            is_30[i] = 'Да'
        else:
            is_30[i] = 'Нет'
        date[i] = date[i].strftime("%m/%d/%Y")
    table['Дата рождения'] = date
    table['Дата рождения'] = table['Дата рождения'].dt.strftime('%m/%d/%Y')
    table['Возраст'] = age
    table['Младше 30?'] = is_30
    table.to_excel('result.xlsx', index=False)  # writing DataFrame to result file


if __name__ == '__main__':
    fill_table()
