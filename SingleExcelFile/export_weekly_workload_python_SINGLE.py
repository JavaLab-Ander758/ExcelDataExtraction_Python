import xlrd
import matplotlib.pyplot as plt
import numpy as np

__author__ = "Anders Rubach Ese"
__version__ = "1.1.0"
__email__ = "aes014@uit.no"


def export_weekly_workload(owner, file_path_excel):
    """
    Export weekly workload for workbook
    """

    file_name_dict = dict()
    file_name_dict[owner] = file_path_excel

    start_week = 3
    end_week = 22
    sheet_prefix = 'Uke'
    cache_assoc_dict = []

    assoc_dict = dict()
    workbook = xlrd.open_workbook(file_path_excel)
    for i in range(3, 22):
        sheet_name = f'{sheet_prefix}{i}'
        worksheet = workbook.sheet_by_name(sheet_name)
        temp_list = []
        for slideY in range(2, 6):
            temp_list.append(worksheet.cell_value(slideY, 5) * 24)
            assoc_dict[sheet_name] = temp_list
    cache_assoc_dict.append(assoc_dict.copy())

    result = cache_assoc_dict[0]
    for i in range(1, len(cache_assoc_dict)):
        for key in cache_assoc_dict[i].keys():
            for k in range(0, 3):
                result[key][k] += cache_assoc_dict[i][key][k]

    for key in result:
        for i in range(0, len(result[key])):
            result[key][i] = result[key][i] / (len(cache_assoc_dict))

    x_values = np.array([i for i in range(start_week, end_week)])
    y_values = list(result.values())
    plt.ylabel('Timer brukt')
    plt.xticks(x_values, list(result.keys()), rotation=90)
    plt.title(f'Ukentlig arbeidstid for \'{owner}\' per kategori (uke{start_week} - {end_week - 1})')
    workload = ['Teori', 'Utvikling', 'Administrativt', 'Logg/Oppgaver']
    for i in range(len(y_values[0])):
        plt.plot(x_values, [pt[i] for pt in y_values], label=workload[i], linewidth=2)

    _total_sum_each_week = []
    for key in result.keys():
        cache = 0
        for i in range(0, len(result[key])):
            cache += result[key][i]
        _total_sum_each_week.append(cache)

    plt.plot(x_values, _total_sum_each_week, "r--", label='alle typer', color="#000000", linewidth=0.7)
    plt.fill_between(x_values, _total_sum_each_week, color="#7492E3", alpha=0.2)
    plt.grid(linestyle=':', linewidth=1)

    plt.legend()

    tot_sum = 0
    for item in cache_assoc_dict:
        for key in item.keys():
            for i in range(0, len(item[key])):
                tot_sum += item[key][i]
    plt.text(17.8, 37.8, f'Totalt: {round(tot_sum)}t', fontsize=12)
    plt.savefig(f'Output/Ukentlig arbeidstid for {owner} per kategori (uke {start_week}-{end_week-1}) SystemUtvikling.png', dpi=1000)
    plt.show()
    print('Done')


export_weekly_workload('Anders', 'assets/temp_Timeliste UkeX Na1 Ese Anders Rubach.xlsm')