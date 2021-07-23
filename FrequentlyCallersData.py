import openpyxl
import pandas as pd
from datetime import datetime, time
import matplotlib.pyplot as plt


def convert_to_date(string):
    return datetime.strptime(string, '%d-%m-%Y %H:%M').strftime('%d-%m-%Y %H:%M')


def convert_to_hour(string):
    return int(datetime.strptime(string, '%d-%m-%Y %H:%M').strftime('%H'))


def convert_to_day(string):
    return datetime.strptime(string, '%d-%m-%Y').strftime('%d-%a')


def strip_time(string):
    return datetime.strptime(string, '%d-%m-%Y %H:%M').strftime('%d-%m-%Y')


def create_weekday(date: datetime):
    day = date.strftime('%A')
    weekdays = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
    day_count = weekdays.index(day.lower()) + 1
    return date.strftime(f'0{day_count} - {day}')


class FrequentlyCallersData:

    def __init__(self, path_name, month, output_path_name):
        self.df = pd.read_excel(path_name, parse_dates=['Datum', 'Begintijd', 'Eindtijd'])
        self.writer = pd.ExcelWriter(output_path_name, engine='openpyxl', mode='a')
        self.month = month
        self.convert_date()

    def convert_date(self):
        self.df['Datum'] = self.df['Datum'].apply(lambda x: x.date())
        self.df['Begintijd'] = self.df['Begintijd'].apply(lambda x: x.time())
        self.df['Eindtijd'] = self.df['Eindtijd'].apply(lambda x: x.time())

    def get_date_list(self):
        return self.df['Datum'].unique()

    def create_dataframes_by_day(self):
        df_list = []
        for date in self.get_date_list():
            df_to_append = self.df[self.df['Datum'] == date]
            df_list.append(df_to_append)

        return df_list

    def create_frequency_dataframe(self, dataframe, hour=13, minute=0):
        morning_time = time(hour=hour, minute=minute)

        df_data = pd.DataFrame(columns=
                               ['Date', f'Frequency before {morning_time}', f'Frequency after {morning_time}',
                                'Total missed calls'])

        df_data['Date'] = [dataframe.iloc[0]['Datum']]
        df_data[f'Frequency before {morning_time}'] = [len(dataframe[dataframe['Begintijd'] < morning_time])]
        df_data[f'Frequency after {morning_time}'] = [len(dataframe[dataframe['Begintijd'] >= morning_time])]
        df_data['Total missed calls'] = [df_data[f'Frequency before {morning_time}'].sum() +
                                         df_data[f'Frequency after {morning_time}'].sum()]

        return df_data

    def create_overall_frequency_df(self):
        df_list = self.create_dataframes_by_day()
        temp_df_list = [self.create_frequency_dataframe(d) for d in df_list]
        main_df = temp_df_list[0]
        main_df = main_df.append([d for d in temp_df_list[1:]])
        return main_df

    def group_frequency_by_days(self, df: pd.DataFrame):
        df = df.copy(deep=True)
        df['Date'] = df['Date'].apply(lambda x: create_weekday(x))
        df = df.groupby('Date').sum()
        return df

    def create_plot(self, dataframe, sheet_name):
        # dataframe = dataframe.set_index('Date')
        title = f'Amount of missed calls in {self.month}'

        ax = dataframe.plot(kind='bar', rot=20, title=title, figsize=(12, 6))
        for p in ax.patches:
            ax.annotate(str(p.get_height()), (p.get_x() * 1.007, p.get_height() * 1.007))

        image_name = f'{title}.png'
        plt.savefig(image_name,  bbox_inches='tight', pad_inches=0.1)
        plt.clf()

        dataframe.to_excel(self.writer, sheet_name=sheet_name)
        self.write_plot_to_excel(sheet_name, image_name, 'G2')

    def write_plot_to_excel(self, sheet_name, image_name, cell_placement):
        worksheet = self.writer.sheets[sheet_name]
        img = openpyxl.drawing.image.Image(image_name)
        img.anchor = cell_placement
        worksheet.add_image(img)

    def write_data_to_excel(self):
        overall_df = self.create_overall_frequency_df()
        df_to_use = self.group_frequency_by_days(overall_df)
        sheet_name = f'Frequency calls {self.month}'
        self.create_plot(df_to_use, sheet_name)
        self.writer.save()
