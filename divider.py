from openpyxl import Workbook, load_workbook

from colors import bcolors

def divider(column_specific, wb, daughter=True):
    s = wb['Sheet']
    assignation_code = {}
    i = 2
    while True:
        ass = s.cell(row=i, column=column_specific).value
        checker = s.cell(row=i, column=13).value == None and s.cell(row=i, column=19).value == None
        if checker:
            break
        if ass is None:
            i += 1
            continue
        if ass not in assignation_code:
            assignation_code[ass] = {}
            assignation_code[ass]['wb'] = Workbook()
            assignation_code[ass]['s'] = assignation_code[ass]['wb'].active
            assignation_code[ass]['row'] = 1
        for col_ in range(1, 30):
            assignation_code[ass]['s'].cell(row=assignation_code[ass]['row'], column=col_).value = s.cell(row=i,
                                                                                                          column=col_).value
        assignation_code[ass]['row'] += 1
        i += 1
        if not daughter:
            continue
        while s.cell(row=i, column=14).value is None:
            for col_ in range(19, 50):
                assignation_code[ass]['s'].cell(row=assignation_code[ass]['row'], column=col_).value = s.cell(row=i,
                                                                                                              column=col_).value
            assignation_code[ass]['row'] += 1
            i += 1
            if s.cell(row=i, column=19).value == None:
                break
    return assignation_code

class Divider():
    def __init__(self, wb=None, path='Выгрузка/2.xlsx'):
        if wb is not None:
            self.wb = wb
        else:
            self.wb = load_workbook(path)

    def divide_by_assignation_code(self):
        print('\rДелим на файлы по коду помещения', end='')
        assignation_code = divider(14, self.wb)
        for ass, item in assignation_code.items():
            item['wb'].save('Выгрузка/' + ass + '.xlsx')
        print(bcolors.OKGREEN + '\rФайл был разделен по коду жилого помещения' + bcolors.ENDC)

    def divide_by_type_uk(self):
        print('\rДелим на файлы по типу ук', end='')

        assignation_code = divider(1, self.wb)
        for ass, item in assignation_code.items():
            item['wb'].save('Выгрузка/' + ass + ' с дочерними.xlsx')

        assignation_code = divider(1, self.wb, daughter=False)
        for ass, item in assignation_code.items():
            item['wb'].save('Выгрузка/' + ass + ' без дочерних.xlsx')

        print(bcolors.OKGREEN + '\rФайл был разделен по типу ук' + bcolors.ENDC)