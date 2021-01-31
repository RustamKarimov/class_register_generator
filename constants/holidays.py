import datetime

DATE_FORMAT = '%Y-%m-%d'

HOLIDAYS_AS_LIST = [
    '2021-01-01',
    '2021-03-21',
    '2021-03-22',
    '2021-04-02',
    '2021-04-05',
    '2021-04-26',
    '2021-04-27',
    '2021-05-01',
    '2021-06-16',
    '2021-08-09',
    '2021-09-24',
    '2021-12-16',
    '2021-12-25',
    '2021-12-26',
]

HOLIDAYS_AS_DATE = [datetime.datetime.strptime(day, DATE_FORMAT) for day in HOLIDAYS_AS_LIST]