import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO


def strip_name(string):
    id1 = string.index('-') + 2
    id2 = string.index('total')
    return string[id1:id2]


def strip_label_employee(string):
    id1 = string.index('-') + 2
    return string[id1:]


def strip_label(string):
    id1 = string.index('-') + 5
    return string[id1:]


def convert_decimal(string):
    index = string.index('.')
    num = int(string[:index])
    decimal = float(string[index + 1:]) / 100
    return (num + decimal) * 60 * 60


def convert_to_seconds(string):
    if '1899' in string:
        return 0
    if len(string) == 2:
        amount = int(string)
        return amount
    if len(string) == 5:
        index = string.index(':')
        s = string[:index]
        amount = int(s) * 60
        return amount + convert_to_seconds(string[index + 1:])
    if len(string) >= 8:
        index = string.index(':')
        s = string[:index]
        if '.' in s:
            amount = convert_decimal(s)
        else:
            amount = int(s) * 60 * 60
        return amount + convert_to_seconds(string[index + 1:])


def create_row_total(dataframe, employee):
    result = [f'{employee} total', 0]
    col_names = dataframe.columns.tolist()

    for c in col_names[2:]:
        if 'Mean' in c:
            result_column = dataframe[c].mean().round(2)
        else:
            result_column = dataframe[c].sum()
        result.append(result_column)

    return dict(zip(col_names, result))


class EmployeeData:

    def __init__(self, path_name, month, output_path_name):
        self.col_names = ['Agent', 'NaN', 'Queue', 'NaN', 'TotalLoggedInTime', 'CallsAnswered', '%CallsServiced',
                          'CallsAnsweredPerHour', 'RingTimeTotal', 'RingTimeTotalMean', 'TalkTimeTotal', 'TalkTimeMean']
        self.col_convert = ['TotalLoggedInTime', 'RingTimeTotal', 'RingTimeTotalMean', 'TalkTimeTotal', 'TalkTimeMean']
        self.month = month
        self.writer = pd.ExcelWriter(output_path_name, engine='openpyxl', mode='a')
        self.df = pd.read_excel(path_name, dtype=str)
        self.clean_up_df()
        self.convert_in_df()
        self.make_columns_numeric()
        self.employee_df = self.df.copy()
        self.labels_df = self.df.copy()

    def clean_up_df(self):
        self.df = self.df.iloc[5:len(self.df) - 2]
        self.df.columns = self.col_names
        del(self.df['NaN'])

    def make_columns_numeric(self):
        numeric_list = ['CallsAnswered', '%CallsServiced', 'CallsAnsweredPerHour']
        for i in numeric_list:
            self.df[i] = pd.to_numeric(self.df[i])

    def get_employees_list(self):
        return list(dict.fromkeys(self.employee_df['Agent'].tolist()))

    def get_labels_list(self):
        return list(dict.fromkeys(self.labels_df['Queue'].tolist()))

    def convert_in_df(self):
        for i in self.col_convert:
            self.df[i] = self.df[i].apply(lambda x: convert_to_seconds(x))

    def create_employee_dataframes(self):
        df_list = []
        employees = self.get_employees_list()

        for e in employees:
            df_to_append = self.employee_df[self.employee_df['Agent'].str.contains(e)]
            result = create_row_total(df_to_append, e)
            df_to_append = df_to_append.append(result, ignore_index=True)
            df_list.append(df_to_append)

        return df_list

    def select_employee_df(self, column_name, filter_condition):
        df_list = self.create_employee_dataframes()
        main_df = df_list[0]
        main_df = main_df.append([d for d in df_list[1:]])
        cleaned_df = main_df[main_df[column_name].str.contains(filter_condition)]
        cleaned_df = cleaned_df[cleaned_df['CallsAnswered'] > 0]

        return cleaned_df

    def create_employee_total_df(self):
        df = self.select_employee_df('Agent', 'total')
        df['Agent'] = df['Agent'].apply(lambda x: strip_name(x))
        title = f'CallsAnswered per person in {self.month}'

        ax = df.plot(x='Agent', y='CallsAnswered', kind='bar', title=title, rot=20, figsize=(17, 7), color='DarkOrange')
        for p in ax.patches:
            ax.annotate(str(p.get_height()), (p.get_x() * 1.007, p.get_height() * 1.007))

        image_name = BytesIO()

        plt.savefig(image_name,  bbox_inches='tight', pad_inches=0.1, format='png')
        # plt.clf()
        # testing if this will help the permission denied error in the main.exe
        plt.close()

        image_name.seek(0)

        sheet_name = f'CallsAnswered {self.month}'

        df.to_excel(self.writer, sheet_name=sheet_name)

        self.write_plot_to_excel(sheet_name, image_name, 'O2')

        self.writer.save()

    def write_plot_to_excel(self, sheet_name, image_name, cell_placement):
        worksheet = self.writer.sheets[sheet_name]
        img = openpyxl.drawing.image.Image(image_name)
        img.anchor = cell_placement
        worksheet.add_image(img)

    def get_filtered_labels(self, label_to_remove):
        labels = self.get_labels_list()
        for i in labels:
            if label_to_remove in i:
                labels.remove(i)

        return labels

    def remove_zero_calls_labels(self):
        label_to_remove = 'MatchQ'
        labels = self.get_filtered_labels(label_to_remove)
        for l in labels:
            if self.labels_df[self.labels_df['Queue'] == l]['CallsAnswered'].sum() < 1:
                labels.remove(l)

        return labels

    def create_data_labels(self):
        data_dict = {}
        employees = self.get_employees_list()
        labels = self.remove_zero_calls_labels()
        for emp in employees:
            emp_data = []
            for l in labels:
                try:
                    result = self.labels_df[(self.labels_df['Agent'] == emp) &
                                            (self.labels_df['Queue'] == l)]['CallsAnswered'].values[0]
                except IndexError:
                    result = 0
                emp_data.append(result)
            if sum(emp_data) > 0:
                name = strip_label_employee(emp) + ' - ' + str(sum(emp_data))
                data_dict[name] = emp_data

        return data_dict

    def create_labels_plot(self):
        color_list = [
            '#e6194B', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#42d4f4', '#f032e6', '#bfef45',
            '#fabed4', '#469990', '#dcbeff', '#9A6324', '#800000', '#aaffc3', '#808000', '#ffd8b1',
            '#000075', '#a9a9a9', '#ffffff', '#000000'
        ]
        data_dict = self.create_data_labels()
        labels = self.remove_zero_calls_labels()
        labels = [strip_label(l) for l in labels]
        plotdata = pd.DataFrame(data_dict, index=labels)
        plotdata = plotdata.sort_index(axis=0)
        title = f'CallsAnswered per label in {self.month}'

        plotdata.plot(kind="bar", stacked=True, figsize=(20, 8), rot=10, title=title, color=color_list)

        # the legend is now outside the plot -bbox_to_anchor=(1, 1)- but that makes the plot more wide.
        # currently removing the -bbox_to_anchor=(1, 1)-
        plt.legend(title='name person - #calls answered', loc='upper right', bbox_to_anchor=(1, 1))

        image_name = BytesIO()

        plt.savefig(image_name, bbox_inches='tight', pad_inches=0.1, format='png')
        # plt.clf()
        # testing if this will help the permission denied error in the main.exe
        plt.close()

        image_name.seek(0)

        sheet_name = f'CallsAnswered per label {self.month}'

        plotdata.to_excel(self.writer, sheet_name=sheet_name)

        self.write_plot_to_excel(sheet_name, image_name, 'T2')
        self.writer.save()

    def create_both_plots(self):
        self.create_employee_total_df()
        self.create_labels_plot()

