import os
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt


class ConvertData:

    def __init__(self, path_name, month, output_path_name,  has_file_columns=False):
        self.col_names = ["Queue", "NaN", "CallsTotal", "CallsAbandoned", "CallsAnswered", "CallsServicedPercentage",
                          "RingTimeTotal", "RingTimeAverage", "TalkTimeTotal", "TalkTimeAverage", "SuccessfulCallbacks"]
        self.month = month
        self.filter_list = ['Yellow', 'Green', 'Blue', 'All Teams', 'FS', 'HF', 'UBPlus', 'MatchQ']
        self.team_list = ['Total Yellow', 'Total Green', 'Total Blue', 'Total All Teams']
        if not os.path.exists(output_path_name):
            openpyxl.Workbook().save(output_path_name)
        self.writer = pd.ExcelWriter(output_path_name, engine='openpyxl', mode='a')

        if has_file_columns:
            self.df = pd.read_excel(path_name)
        else:
            self.df = pd.read_excel(path_name, names=self.col_names)

        self.clean_up_dataframe()
        self.calculate_sum_or_mean()

    def clean_up_dataframe(self):
        self.df = self.df.iloc[5:21]
        del(self.df["NaN"])
        self.df["Queue"] = self.df.Queue.apply(lambda x: x[6:])

    def convert_time_to_seconds(self, string):
        if len(string) == 3:
            s = int(string[:2])
            return s
        elif len(string) == 7:
            minutes = int(string[:2]) * 60
            return minutes + self.convert_time_to_seconds(string[4:])
        elif len(string) == 11:
            hours = int(string[:2]) * 60 * 60
            return hours + self.convert_time_to_seconds(string[4:])
        elif len(string) == 15:
            days = int(string[:2]) * 24 * 60 * 60
            return days + self.convert_time_to_seconds(string[4:])

    def convert_to_decimal(self, string):
        index = string.index(' %')
        decimal = float(string[:index]) / 100
        return decimal

    def calculate_sum_or_mean(self):
        self.df["RingTimeTotal"] = self.df.RingTimeTotal.apply(lambda x: self.convert_time_to_seconds(x))
        self.df["RingTimeAverage"] = self.df.RingTimeAverage.apply(lambda x: self.convert_time_to_seconds(x))
        self.df["TalkTimeTotal"] = self.df.TalkTimeTotal.apply(lambda x: self.convert_time_to_seconds(x))
        self.df["TalkTimeAverage"] = self.df.TalkTimeAverage.apply(lambda x: self.convert_time_to_seconds(x))
        self.df["CallsServicedPercentage"] = self.df.CallsServicedPercentage.apply(lambda x: self.convert_to_decimal(x))

    def get_results_in_dict(self, dataframe, team):
        specified_columns = ["Queue", "CallsTotal", "CallsAbandoned", "CallsAnswered", "RingTimeTotal",
                             "RingTimeAverage", "TalkTimeTotal", "TalkTimeAverage"]
        result = [f"Total {team}"]

        for s in specified_columns[1:]:
            if "Mean" in s:
                result_column = dataframe[s].mean()
            else:
                result_column = dataframe[s].sum()
            result.append(result_column)

        return dict(zip(specified_columns, result))

    def append_dataframes_to_list(self):
        df_list = []

        for t in self.filter_list:
            df_to_append = self.df[self.df["Queue"].str.contains(t)]
            result_dict = self.get_results_in_dict(df_to_append, t)
            df_to_append = df_to_append.append(result_dict, ignore_index=True)
            df_list.append(df_to_append)

        return df_list

    def select_data_from_all_dataframes(self, column_name, filter_list):
        df_list = self.append_dataframes_to_list()
        main_df = df_list[0]
        main_df = main_df.append([d for d in df_list[1:]])
        cleaned = main_df[main_df[column_name].isin(filter_list)]

        return cleaned

    def write_plot_to_excel(self, sheet_name, image_name, cell_placement):
        worksheet = self.writer.sheets[sheet_name]
        img = openpyxl.drawing.image.Image(image_name)
        img.anchor = cell_placement
        worksheet.add_image(img)

    def create_call_plot(self, column_name, title, column_list, sheet_name, sub_plot=False):
        df = self.select_data_from_all_dataframes(column_name, self.team_list)
        data_df = df[column_list]

        data_df = data_df.set_index('Queue')
        if sub_plot:
            data_df.loc['Total Overall', :] = round(data_df.sum(axis=0) / 4)
        else:
            data_df.loc['Total Overall', :] = data_df.sum(axis=0)

        if sub_plot:
            data_df = data_df.reset_index()

            fig, axes = plt.subplots(nrows=1, ncols=2, figsize=(12, 4))

            ax1, ax2 = axes

            data_df.plot(x="Queue", y=column_list[1], kind="bar", ax=ax1, rot=20)
            data_df.plot(x="Queue", y=column_list[2], kind="bar", ax=ax2, rot=20, color='red')

            # Increasing the y-axis on the second subplot, in order to place the legend higher
            ax2.set_ylim(0, max(data_df[column_list[2]]) + (max(data_df[column_list[2]]) / 5))

            ax1.title.set_text(f'{column_list[1]} in {self.month}')
            ax2.title.set_text(f'{column_list[2]} in {self.month}')

            ax1.set_ylabel('Time in seconds')
            ax2.set_ylabel('Time in seconds')

            for x in axes:
                for p in x.patches:
                    x.annotate(str(int(p.get_height())), (p.get_x() * 1.007, p.get_height() * 1.007))
        else:
            ax = data_df.plot(kind='bar', rot=20, title=title, figsize=(14, 4))
            for p in ax.patches:
                ax.annotate(str(int(p.get_height())), (p.get_x() * 1.007, p.get_height() * 1.007))

        image_name = f'{title}.png'

        plt.savefig(image_name,  bbox_inches='tight', pad_inches=0.1)
        plt.clf()

        data_df.to_excel(self.writer, sheet_name=sheet_name)

        self.write_plot_to_excel(sheet_name, image_name, 'G2')

    def create_missed_calls_plot(self, path_name):
        df = pd.read_excel(path_name)
        column_to_find = 'Software applicaties.Naam applicatie'
        columns = df.columns.tolist()
        index = columns.index(column_to_find)
        df[columns[index]] = df[columns[index]].fillna("Onbekend")
        data_df = df[columns[index]]

        counts = df[columns[index]].value_counts().to_dict()

        def absolute_value(val):
            a = round(val / 100. * len(df[columns[index]]), 0)
            return int(a)

        plt.pie([float(v) for v in counts.values()], labels=[k for k in counts.keys()], autopct=absolute_value)
        plt.title(f"Antwoordservice {self.month}")
        plt.suptitle(f'Total missed calls = {len(data_df)}', y=0.1)

        image_name = f'Piechart {self.month}.png'
        plt.savefig(image_name,  bbox_inches='tight', pad_inches=0.1)

        plt.clf()

        sheet_name = f'Antwoordservice {self.month}'

        data_df.to_excel(self.writer, sheet_name=sheet_name)

        self.write_plot_to_excel(sheet_name, image_name, 'G2')

    def write_dataframes_to_excel(self, create_plots=True, create_missed_calls=False, path_name_missed_calls=None):
        df_list = self.append_dataframes_to_list()
        main_df = df_list[0]

        for i in df_list[1:]:
            main_df = main_df.append(i)

        if create_plots:
            # Creating the calls plot
            self.create_call_plot(column_name='Queue', title=f'CallsAbandoned vs CallsAnswered in {self.month}',
                                  column_list=['Queue', 'CallsAbandoned', 'CallsAnswered'],
                                  sheet_name=f'Chart Calls {self.month}')
            # Creating the ring time plot
            self.create_call_plot(column_name='Queue', title=f'Ring Time in {self.month}',
                                  column_list=['Queue', 'RingTimeTotal', 'RingTimeAverage'],
                                  sheet_name=f'Chart RingTime {self.month}', sub_plot=True)
            # Creating the talk time plot
            self.create_call_plot(column_name='Queue', title=f'Talk Time in {self.month}',
                                  column_list=['Queue', 'TalkTimeTotal', 'TalkTimeAverage'],
                                  sheet_name=f'Chart TalkTime {self.month}', sub_plot=True)
        if create_missed_calls:
            self.create_missed_calls_plot(path_name_missed_calls)

        self.writer.save()
