from icalendar import Calendar, Event, vText
from datetime import datetime, timedelta
from pathlib import Path
import os
import pytz
import pandas as pd


def parse_xls(filename):
    df_in = pd.read_excel(filename, index_col=0, header=None)
    start_date = datetime(2023, 4, 16)
    df_out = pd.DataFrame(columns=['DATE', 'TRAINING'])
    tr_list = [item for index, row in df_in.iterrows() for item in row]

    for i in range(1, df_in.size + 1):
        df_out.loc[i, 'DATE'] = start_date - timedelta(days=int(df_in.size) - i)
        df_out.loc[i, 'TRAINING'] = tr_list[i - 1]

    df_out.to_excel('SORTED_TRAINING.xlsx', index=False)
    return df_out


def save_to_disk(cal):
    directory = Path.cwd() / 'MyCalendar'
    try:
        directory.mkdir(parents=True, exist_ok=False)
    except FileExistsError:
        print("Folder already exists")
    else:
        print("Folder was created")

    f = open(os.path.join(directory, 'TRAINING_PLAN.ics'), 'wb')
    f.write(cal.to_ical())
    f.close()


def display(cal):
    return cal.to_ical().decode("utf-8").replace('\r\n', '\n').strip()

def main():
    df = parse_xls('input.xlsx')
    # init the calendar
    cal = Calendar()

    # Some properties are required to be compliant
    cal.add('prodid', '-/Marathon')
    cal.add('version', '2.0')

    # Add subcomponents
    for index, row in df.iterrows():
        event = Event()

        date = row['DATE'].strftime('%Y-%m-%d %X').split(' ')[0].split('-')
        if row['TRAINING'] == 'Rest' or row['TRAINING'] == 'Cross':
            training = row['TRAINING']
        else:
            training = str(row['TRAINING']) + " KM"

        if row['DATE'].strftime("%A") == 'Saturday':
            event.add('dtstart', datetime(int(date[0]), int(date[1]), int(date[2]), 12, 0, 0, tzinfo=pytz.utc))
            event.add('dtend', datetime(int(date[0]), int(date[1]), int(date[2]), 14, 0, 0, tzinfo=pytz.utc))
        else:
            event.add('dtstart', datetime(int(date[0]), int(date[1]), int(date[2]), 19, 0, 0, tzinfo=pytz.utc))
            event.add('dtend', datetime(int(date[0]), int(date[1]), int(date[2]), 20, 0, 0, tzinfo=pytz.utc))
        event.add('summary', training)
        event['location'] = vText('Amsterdam, Netherlands')
        cal.add_component(event)

    save_to_disk(cal)


if __name__ == '__main__':
    main()




