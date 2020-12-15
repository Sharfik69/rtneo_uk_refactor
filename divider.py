from openpyxl import Workbook, load_workbook

from colors import bcolors


class Divider():
    def __init__(self, wb = None, path = 'Выгрузка/2.xlsx'):
        if wb is not None:
            self.wb = wb
        else:
            self.wb = load_workbook(path)

    def divide_by_assignation_code(self):
        print('\rДелим на файлы по коду помещения', end='')
        s = self.wb['Sheet']
        assignation_code = {}
        i = 2
        while True:
            ass = s.cell(row=i, column=14).value
            checker = s.cell(row=i, column=19).value
            if ass == None and checker == None:
                break
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
            while s.cell(row=i, column=14).value is None:
                for col_ in range(19, 50):
                    assignation_code[ass]['s'].cell(row=assignation_code[ass]['row'], column=col_).value = s.cell(row=i,
                                                                                                                  column=col_).value
                assignation_code[ass]['row'] += 1
                i += 1
                if s.cell(row=i, column=19).value == None:
                    break
        print('\rСохраняем', end='')
        for ass, item in assignation_code.items():
            item['wb'].save('Выгрузка/' + ass + '.xlsx')
        print(bcolors.OKGREEN + '\rФайл был разделен по коду жилого помещения' + bcolors.ENDC)
