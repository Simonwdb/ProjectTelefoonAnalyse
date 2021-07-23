import EmployeeData
import FileManagement
import ExtractData
import FrequentlyCallersData


def start_program(file_list):

    cs_overall_file2 = file_list[0]
    cs_employee_file2 = file_list[1]
    answer_services_file2 = file_list[2]
    output_file2 = file_list[3]
    month2 = file_list[4]

    ExtractData.ConvertData(cs_overall_file2, month2, output_file2, has_file_columns=False)\
        .write_dataframes_to_excel(create_plots=True, create_missed_calls=True,
                                   path_name_missed_calls=answer_services_file2)

    EmployeeData.EmployeeData(cs_employee_file2, month2, output_file2).create_both_plots()

    FrequentlyCallersData.FrequentlyCallersData(answer_services_file2, month2, output_file2).\
        write_data_to_excel()

    FileManagement.remove_images()

