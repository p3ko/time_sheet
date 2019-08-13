import xlsxwriter
import calendar
import settings
import datetime
import random
from ics import Calendar
import requests

MONTHS = {'Styczeń': 1, 'Luty': 2, 'Marzec': 3, 'Kwiecień': 4, 'Maj': 5, 'Czerwiec': 6, 'Lipiec': 7, 'Sierpień': 8,
          'Wrzesień': 9, 'Październik': 10, 'Listopad': 11, 'Grudzień': 12}

calendar_url = 'https://www.officeholidays.com/ics-clean/poland'


def my_random(days, min_hours, max_hours):
    working_hours = []
    for x in range(int(days / 8)):
        working_hours.append(random.randint(min_hours, max_hours))

    return working_hours


def holidays(url, year, month):
    holidays_list = []
    c = Calendar(requests.get(url).text)
    e = list(c.timeline)

    for i in e:
        if month < 10:
            if f'{year}-0{month}' in str(i.begin):
                holidays_list.append(str(i.begin)[:10])
        else:
            if f'{year}-{month}' in str(i.begin):
                holidays_list.append(str(i.begin)[:10])

    return holidays_list


def list_sum(my_list):
    suma = 0
    for number in my_list:
        suma += number

    return suma


def main():
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(
        f'{settings.employee} - EWIDENCJA CZASU PRACY - {settings.year}.{MONTHS[settings.month.capitalize()]} {settings.month.capitalize()}.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.set_column('A:F', 12)

    merge_format_bold = workbook.add_format({
        'bold': 1,
        'bottom': 2,
        'top': 2,
        'left': 2,
        'right': 2,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#d9d9d9',
        'font': 'Tahoma',
        'font_size': 10
    })

    merge_format = workbook.add_format({
        'bottom': 2,
        'top': 2,
        'left': 2,
        'right': 2,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#d9d9d9',
        'font': 'Tahoma',
        'font_size': 10})

    weekend_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#d9d9d9',
        'font': 'Tahoma',
        'font_size': 10})

    date_format = workbook.add_format({
        'num_format': 'dd.mm.yyyy',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#d9d9d9',
        'font': 'Tahoma',
        'font_size': 10})

    weekend_date_format = workbook.add_format({
        'num_format': 'dd.mm.yyyy',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#d9d9d9',
        'font': 'Tahoma',
        'font_size': 10})

    worksheet.merge_range('A1:C1', 'EWIDENCJA CZASU PRACY', merge_format_bold)
    worksheet.merge_range('D1:F1', settings.company, merge_format_bold)
    worksheet.merge_range('A2:A5', 'Osoba:', merge_format)
    worksheet.merge_range('B2:C5', settings.employee, merge_format)
    worksheet.merge_range('D2:D3', 'Miesiąc:', merge_format)
    worksheet.merge_range('E2:F3', settings.month.capitalize(), merge_format)
    worksheet.merge_range('D4:D5', 'Rok:', merge_format)
    worksheet.merge_range('E4:F5', settings.year, merge_format)

    center = workbook.add_format({'align': 'center', 'border': 1, 'font': 'Tahoma', 'font_size': 10})

    headers_format = workbook.add_format({
        'align': 'center',
        'bottom': 2,
        'top': 2,
        'left': 2,
        'right': 2,
        'font': 'Tahoma',
        'font_size': 10})
    total_hours_format = workbook.add_format(
        {'align': 'center', 'border': 1, 'fg_color': '#ed7d31', 'font': 'Tahoma', 'font_size': 10})
    hours_remained = workbook.add_format(
        {'align': 'center', 'border': 1, 'fg_color': '#ffff00', 'font': 'Tahoma', 'font_size': 10})

    worksheet.write('A6', 'Data', headers_format)
    worksheet.write('B6', 'Projekt', headers_format)
    worksheet.write('C6', 'Liczba godz.', headers_format)
    worksheet.write('D6', 'Procesy', headers_format)
    worksheet.write('E6', 'Czynności', headers_format)
    worksheet.write('F6', 'Uwagi', headers_format)

    c = calendar.TextCalendar(calendar.MONDAY)
    holidays_list = holidays(calendar_url, settings.year, MONTHS[settings.month.capitalize()])

    suma = 0
    while True:
        if suma != settings.total_working_hours:
            md = my_random(settings.total_working_hours, settings.min_hours,
                           settings.max_hours)
            suma = list_sum(md)
        else:
            break

    working_days = 0
    counter = 7

    for day in c.itermonthdays(settings.year, MONTHS[settings.month.capitalize()]):
        if day != 0:

            weekday = calendar.weekday(settings.year, MONTHS[settings.month.capitalize()], int(day))
            date_time = datetime.datetime.strptime(f'{settings.year}-{MONTHS[settings.month.capitalize()]}-{day}',
                                                   '%Y-%m-%d')

            if weekday == 5 or weekday == 6 or str(date_time)[:10] in holidays_list:
                worksheet.write_datetime(f'A{counter}', date_time, weekend_date_format)
                worksheet.write_blank(f'B{counter}', None, weekend_format)
                worksheet.write_blank(f'C{counter}', None, weekend_format)
                worksheet.write_blank(f'D{counter}', None, weekend_format)
                worksheet.write_blank(f'E{counter}', None, weekend_format)
                worksheet.write_blank(f'F{counter}', None, weekend_format)
                counter += 1
            else:
                worksheet.write_datetime(f'A{counter}', date_time, date_format)
                worksheet.write_string(f'B{counter}', settings.project, center)
                worksheet.write_number(f'C{counter}', md[working_days], center)
                worksheet.write_blank(f'D{counter}', None, center)
                worksheet.write_blank(f'E{counter}', None, center)
                worksheet.write_blank(f'F{counter}', None, center)
                counter += 1
                working_days += 1

    sum_formula = '{=SUM(C7:C' + str(counter - 1) + ')}'
    remain_formula = '{=(C' + str(counter) + '-' + str(settings.total_working_hours) + ')}'

    worksheet.write_blank(f'A{counter}', None, center)
    worksheet.write_blank(f'B{counter}', None, center)
    worksheet.write_formula(f'C{counter}', sum_formula, total_hours_format)
    worksheet.write_formula(f'C{counter + 1}', remain_formula, hours_remained)
    worksheet.write_blank(f'D{counter}', None, center)
    worksheet.write_blank(f'E{counter}', None, center)
    worksheet.write_blank(f'F{counter}', None, center)

    worksheet.write_blank(f'A{counter + 1}', None, center)
    worksheet.write_blank(f'B{counter + 1}', None, center)
    worksheet.write_blank(f'D{counter + 1}', None, center)
    worksheet.write_blank(f'E{counter + 1}', None, center)
    worksheet.write_blank(f'F{counter + 1}', None, center)

    workbook.close()


if __name__ == '__main__':
    main()
