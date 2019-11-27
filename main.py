from openpyxl import load_workbook


def get_temp_file():
    temp_file = load_workbook('/Users/u6000791/Box/Conservation/Rare Plants/Research Projects/Penstemon 2018-2019/'
                              'PENGRA_2019/DATA_PENGRA_2019/Measured variables_PENGRA.xlsx')
    return temp_file


def get_output_file():
    output_file = load_workbook('/Users/u6000791/Desktop/PENGRA_SEM.xlsx')
    return output_file
