import matplotlib.pyplot as plt
import numpy as np
import xlrd

__author__ = "Anders Rubach Ese"
__version__ = "1.1.0"
__email__ = "aes014@uit.no"
excel_files_path = 'assets/'
output = 'Output/'


def exoprtWeeklyWorkloadAllBooks():
    """
    Export weekly workload for all workbooks
    """

    file_name_dict = dict()
    file_name_dict['Adrian'] = 'temp_Timeliste UkeX Na1 Adrian Chirita.xlsm'
    file_name_dict['Eivind'] = 'temp_Timeliste UkeX Na1 Eriksen Eivind Blom.xlsm'
    file_name_dict['Anders'] = 'temp_Timeliste UkeX Na1 Ese Anders Rubach.xlsm'
    file_name_dict['Alexander'] = 'temp_Timeliste UkeX Na1 Hansen Alexander.xlsm'
    file_name_dict['Joachim'] = 'temp_Timeliste UkeX Na1 Kristensen Joachim.xlsm'
    file_name_dict['Martin'] = 'temp_Timeliste UkeX Na1 Mikalsen Martin.xlsm'
    file_name_dict['Vegard'] = 'temp_Timeliste UkeX Na1 Simonsen Vegard.xlsm'

    minY = 2
    maxY = 5
    startWeek = 3
    endWeek = 20
    sheetPrefix = 'Uke'
    cacheAssocDict = []

    assocDict = dict()
    for owner in file_name_dict:
        workbook = xlrd.open_workbook(excel_files_path + file_name_dict[owner])
        for i in range(6 if owner == 'Joachim' else startWeek, 20 if owner == 'Vegard' else endWeek):
            sheetName = f'{sheetPrefix}{i}'
            worksheet = workbook.sheet_by_name(sheetName)
            tempList = []
            for slideY in range(3 if owner == 'Eivind' else minY, maxY + 2 if owner == 'Eivind' else maxY + 1):
                py_date = xlrd.xldate.xldate_as_datetime(worksheet.cell(slideY, 4 if owner == 'Eivind' else 5).value,
                                                         workbook.datemode)
                hr_decimal = py_date.hour + py_date.minute / 60
                tempList.append(hr_decimal)
                assocDict[sheetName] = tempList
        cacheAssocDict.append(assocDict.copy())

    result = cacheAssocDict[0]
    for i in range(1, len(cacheAssocDict)):
        for key in cacheAssocDict[i].keys():
            for k in range(0, 3):
                result[key][k] += cacheAssocDict[i][key][k]

    for key in result:
        for i in range(0, len(result[key])):
            result[key][i] = result[key][i] / (len(cacheAssocDict))

        print(key)

    x_values = np.array([i for i in range(startWeek, endWeek)])
    y_values = list(result.values())
    plt.ylabel('Timer brukt')
    plt.xticks(x_values, list(result.keys()), rotation=90)
    plt.title(f'Gjennomsnittlig arbeidstid per type arbeid (uke{startWeek}-uke{endWeek - 1})')
    workload = ['Teori', 'Utvikling', 'Administrativt', 'Logg/Oppgaver']
    for i in range(len(y_values[0])):
        plt.plot(x_values, [pt[i] for pt in y_values], label=workload[i], linewidth=2)

    _total_sum_each_week = []
    for key in result.keys():
        cache = 0
        for i in range(0, len(result[key])):
            cache += result[key][i]
        _total_sum_each_week.append(cache)
    print(_total_sum_each_week)
    plt.plot(x_values, _total_sum_each_week, "-", label='alle typer', linewidth=0.2)
    plt.fill_between(x_values, _total_sum_each_week, color="#7492E3", alpha=0.16)
    plt.grid(linestyle=':', linewidth=1, alpha=0.6)

    plt.legend()

    sum = 0
    for item in cacheAssocDict:
        for key in item.keys():
            for i in range(0, len(item[key])):
                sum += item[key][i]
    print(sum)
    plt.text(15.8, 15.8, f'Totalt: {round(sum)}t', fontsize=12)

    plt.savefig(output + 'medianWorkloadPerWeek.png', dpi=2000)
    plt.show()
    print('Done')
    print(len(cacheAssocDict))


def exoprtWeeklyWorkloadOneBooks():
    x = 5
    minY = 2
    maxY = 5
    startWeek = 3
    endWeek = 21

    sheetPrefix = 'Uke'

    assocDict = dict()

    workbook = xlrd.open_workbook(excel_files_path + 'temp_Timeliste UkeX Na1 Ese Anders Rubach.xlsm')
    for i in range(startWeek, endWeek + 1):
        sheetName = f'{sheetPrefix}{i}'
        worksheet = workbook.sheet_by_name(sheetName)
        tempList = []
        for slideY in range(minY, maxY + 1):
            py_date = xlrd.xldate.xldate_as_datetime(worksheet.cell(slideY, x).value, workbook.datemode)
            hr_decimal = py_date.hour + py_date.minute / 60
            tempList.append(hr_decimal)
            assocDict[sheetName] = tempList

    print(assocDict)

    x_values = np.array([i for i in range(startWeek, endWeek + 1)])
    y_values = list(assocDict.values())
    for i in range(len(y_values[0])):
        plt.plot(x_values, [pt[i] for pt in y_values],
                 label=['Teori', 'Utvikling', 'Administrativt', 'Logg/Oppgaver'][i])
    plt.ylabel('Timer brukt')
    plt.xticks(x_values, list(assocDict.keys()), rotation=90)
    plt.title('Timer brukt per uke (Anders)')
    plt.legend()
    plt.savefig(output + 'Anders.png', dpi=500)
    plt.show()


# exoprtWeeklyWorkloadOneBooks()
exoprtWeeklyWorkloadAllBooks()
