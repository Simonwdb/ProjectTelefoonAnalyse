import openpyxl
import pandas as pd
from io import BytesIO
from datetime import datetime, time
import matplotlib.pyplot as plt


def convert_to_date(string: str) -> str:
    return datetime.strptime(string, '%d-%m-%Y %H:%M').strftime('%d-%m-%Y %H:%M')


def convert_to_hour(string: str) -> int:
    return int(datetime.strptime(string, '%d-%m-%Y %H:%M').strftime('%H'))


def convert_to_day(string: str) -> str:
    return datetime.strptime(string, '%d-%m-%Y').strftime('%d-%a')


def strip_time(string: str) -> str:
    return datetime.strptime(string, '%d-%m-%Y %H:%M').strftime('%d-%m-%Y')


def create_day(string: str) -> str:
    return '0' + string + '-' if int(string) < 10 else string + '-'


def create_month(string: str) -> str:
    index = string.index('-')
    first_part = create_day(string[:index])
    second_part = string[index + 1:]
    number = second_part[:second_part.index('-')]
    if int(number) < 10:
        second_part = '0' + second_part
    return first_part + second_part


def create_weekday(string: str) -> str:
    day = datetime.strptime(string, '%d-%m-%Y').strftime('%A')
    weekdays = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
    day_count = weekdays.index(day.lower()) + 1
    return f'0{day_count} - {day}'


def parse_to_datetime(string: str) -> datetime.time:
    return datetime.strptime(string, '%d-%m-%Y %H:%M').time()


class FrequentlyCallersData:

    def __init__(self, path_name, month, output_path_name):
        self.df = pd.read_excel(path_name)
        self.writer = pd.ExcelWriter(output_path_name, engine='openpyxl', mode='a')
        self.month = month
        self.convert_date()

    def convert_date(self):
        self.df['Datum'] = self.df['Datum'].apply(lambda x: create_month(x))
        self.df['Begintijd'] = self.df['Begintijd'].apply(lambda x: create_month(x))
        self.df['Eindtijd'] = self.df['Eindtijd'].apply(lambda x: create_month(x))

        self.df['Datum'] = self.df['Datum'].apply(lambda x: strip_time(x))
        self.df['Begintijd'] = self.df['Begintijd'].apply(lambda x: parse_to_datetime(x))
        self.df['Eindtijd'] = self.df['Eindtijd'].apply(lambda x: parse_to_datetime(x))

    def get_date_list(self) -> list:
        return self.df['Datum'].unique()

    def create_dataframes_by_day(self) -> list[pd.DataFrame]:
        df_list = []
        for date in self.get_date_list():
            df_to_append = self.df[self.df['Datum'] == date]
            df_list.append(df_to_append)

        return df_list

    def create_frequency_dataframe(self, dataframe: pd.DataFrame, hour: int = 13, minute: int = 0) -> pd.DataFrame:
        morning_time = time(hour=hour, minute=minute)

        df_data = pd.DataFrame(columns=
                               ['Day', f'Frequency before {morning_time}', f'Frequency after {morning_time}',
                                'Total missed calls'])

        df_data['Day'] = [dataframe.iloc[0]['Datum']]
        df_data[f'Frequency before {morning_time}'] = [len(dataframe[dataframe['Begintijd'] < morning_time])]
        df_data[f'Frequency after {morning_time}'] = [len(dataframe[dataframe['Begintijd'] >= morning_time])]
        df_data['Total missed calls'] = [df_data[f'Frequency before {morning_time}'].sum() +
                                         df_data[f'Frequency after {morning_time}'].sum()]

        return df_data

    def create_overall_frequency_df(self) -> pd.DataFrame:
        df_list = self.create_dataframes_by_day()
        temp_df_list = [self.create_frequency_dataframe(d) for d in df_list]
        main_df = temp_df_list[0]
        main_df = main_df.append([d for d in temp_df_list[1:]])
        return main_df

    def group_frequency_by_days(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy(deep=True)
        df['Day'] = df['Day'].apply(lambda x: create_weekday(x))
        df = df.groupby('Day').sum()
        return df

    def create_grouped_days_plot(self, dataframe, sheet_name) -> None:
        title = f'Missed calls in {self.month} (grouped by day)'

        ax = dataframe.plot(kind='bar', rot=20, title=title, figsize=(9, 7))
        for p in ax.patches:
            ax.annotate(str(p.get_height()), (p.get_x() + 0.02, p.get_height() + 0.35))

        # trying to change the image_name
        # image_name = f'{title}.png'
        image_name = BytesIO()
        plt.savefig(image_name,  bbox_inches='tight', pad_inches=0.1, format='png')

        # plt.clf()
        # testing if this will help the permission denied error in the main.exe
        plt.close()

        image_name.seek(0)

        dataframe.to_excel(self.writer, sheet_name=sheet_name)
        self.write_plot_to_excel(sheet_name, image_name, 'G2')

    def create_single_days_plot(self, dataframe: pd.DataFrame, sheet_name: str) -> None:
        dataframe = dataframe.sort_values(by='Day', ascending=True)
        dataframe['Day'] = dataframe['Day'].apply(lambda x: convert_to_day(x))
        dataframe = dataframe.set_index('Day')
        title = f'Missed calls in {self.month} (per day)'

        ax = dataframe.plot(kind='bar', rot=20, title=title, figsize=(21, 7))
        for p in ax.patches:
            ax.annotate(str(p.get_height()), (p.get_x(), p.get_height() + 0.15))

        # trying to change the image_name
        # image_name = f'{title}.png'
        image_name = BytesIO()
        plt.savefig(image_name,  bbox_inches='tight', pad_inches=0.1, format='png')

        # plt.clf()
        # testing if this will help the permission denied error in the main.exe
        plt.close()

        image_name.seek(0)

        dataframe.to_excel(self.writer, sheet_name=sheet_name, startrow=9)
        self.write_plot_to_excel(sheet_name, image_name, 'T2')

    def write_plot_to_excel(self, sheet_name, image_name, cell_placement) -> None:
        worksheet = self.writer.sheets[sheet_name]
        img = openpyxl.drawing.image.Image(image_name)
        img.anchor = cell_placement
        worksheet.add_image(img)

    def write_data_to_excel(self) -> None:
        overall_df = self.create_overall_frequency_df()
        df_to_use = self.group_frequency_by_days(overall_df)
        sheet_name = f'Frequency calls {self.month}'
        self.create_grouped_days_plot(df_to_use, sheet_name)
        self.create_single_days_plot(overall_df, sheet_name)
        self.writer.save()
