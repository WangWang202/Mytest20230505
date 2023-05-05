import datetime
import holidays
import chinese_calendar as cn_calendar
from cn_calendar import HTMLCalendar


class MyHTMLCalendar(HTMLCalendar):
    def __init__(self, holidays_dict):
        super().__init__()
        self.holidays_dict = holidays_dict

    def formatday(self, day, weekday):
        if day == 0:
            return '<td class="noday">&nbsp;</td>'
        else:
            date_str = f'{self.year}-{self.month:02d}-{day:02d}'
            holiday_info = self.holidays_dict.get(date_str, '')
            if holiday_info:
                holiday_name = holiday_info.get('name', '')
                holiday_des = holiday_info.get('description', '')
                holiday = f'{holiday_name}<br><span class="holiday-des">{holiday_des}</span>'
                return f'<td class="holiday" title="{holiday_des}">{day}<br>{holiday}</td>'
            else:
                return f'<td>{day}</td>'

    def formatmonth(self, theyear, themonth, withyear=True):
        self.year, self.month = theyear, themonth
        return super().formatmonth(theyear, themonth, withyear=withyear)


def get_holidays(start, end):
    # Get holidays in China and other countries
    china_holidays = holidays.CountryHoliday('China')
    us_holidays = holidays.CountryHoliday('US')
    ca_holidays = holidays.CountryHoliday('Canada')
    uk_holidays = holidays.CountryHoliday('UK')
    japan_holidays = holidays.CountryHoliday('Japan')
    india_holidays = holidays.CountryHoliday('India')
    aus_holidays = holidays.CountryHoliday('Australia')
    brazil_holidays = holidays.CountryHoliday('Brazil')

    # Get lunar holidays in China
    lunar_holidays = {}
    for year in range(start.year, end.year + 1):
        for month in range(1, 13):
            date = datetime.date(year, month, 1)
            month_days = cn_calendar.monthrange(year, month)[1]
            for day in range(1, month_days + 1):
                lunar_date = cn_calendar.LunarDate.from_solar_date(year, month, day)
                if lunar_date.is_holiday:
                    holiday_info = {'name': lunar_date.holiday, 'description': '农历节日'}
                    lunar_holidays[lunar_date.__str__()] = holiday_info

    # Merge all holidays into a dictionary
    holidays_dict = {}
    for holiday in [china_holidays, us_holidays, ca_holidays, uk_holidays, japan_holidays,
                    india_holidays, aus_holidays, brazil_holidays, lunar_holidays]:
        for date_str, holiday_name in holiday.items():
            if date_str not in holidays_dict:
                holiday_info = {'name': holiday_name, 'description': ''}
                holidays_dict[date_str] = holiday_info
            else:
                holidays_dict[date_str]['name'] += f',{holiday_name}'

    return holidays_dict


def generate_calendar(start_year, end_year):
    # Create a dictionary with all holidays in the specified range
    start_date = datetime.date(start_year, 1, 1)
    end_date = datetime.date(end_year, 12, 31)
    holidays_dict = get_holidays(start_date, end_date)

    # Create HTML calendars for each year and month
html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>2023年节日日历</title>
<style>
table {
  font-family: arial, sans-serif;
  border-collapse: collapse;
  width: 100%;
}

td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

th {
  background-color: #dddddd;
}
</style>
</head>
<body>

"""
for month in range(1, 13):
    month_days = cn_calendar.monthcalendar(year, month)
    month_name = cn_calendar.month_name[month]
    html += f'<h2>{year}年{month}月</h2>'
    html += '<table>'
    html += '<tr><th>周一</th><th>周二</th><th>周三</th><th>周四</th><th>周五</th><th>周六</th><th>周日</th></tr>'

    for week in month_days:
        html += '<tr>'
        for day in week:
            if day == 0:
                html += '<td></td>'
            else:
                # Add Chinese solar terms
                solar_term = chinese_calendar.get_solar_term(year, month, day)
                if solar_term:
                    html += f'<td>{day}<br>{solar_term}</td>'
                else:
                    html += f'<td>{day}</td>'
                # Add China holidays
                china_holidays = holidays.China(years=year)
                china_holiday = china_holidays.get(datetime.date(year, month, day))
                if china_holiday:
                    html += f'<td>{china_holiday}</td>'
                else:
                    html += '<td></td>'
                # Add world holidays
                world_holidays = holidays.CountryHoliday('WORLD', years=year)
                world_holiday = world_holidays.get(datetime.date(year, month, day))
                if world_holiday:
                    html += f'<td>{world_holiday}</td>'
                else:
                    html += '<td></td>'

        html += '</tr>'
    html += '</table>'

html += """
</body>
</html>
"""

with open('cn_calendar.html', 'w') as f:
    f.write(html)

