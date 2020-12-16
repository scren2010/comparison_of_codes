import xlsxwriter
import xlrd


def clear_exel(exle, row):
    return str(exle.row(row)[0]).replace('text:', '').replace("'", "")


def main():
    en_excel_data_file = xlrd.open_workbook('./EN.xlsx')
    ru_excel_data_file = xlrd.open_workbook('./RU.xlsx')
    en_sheet = en_excel_data_file.sheet_by_index(3)
    ru_sheet = ru_excel_data_file.sheet_by_index(3)

    en_row_number = en_sheet.nrows
    ru_row_number = ru_sheet.nrows

    eng = {}
    rus = {}
    if en_row_number > 0:
        for row in range(2, en_row_number):
            number = clear_exel(en_sheet, row)
            text = str(en_sheet.row(row)[1]).replace('text:', '').replace("'", "")
            eng.update({number: f'{text}'})

    if ru_row_number > 0:
        for row in range(1, ru_row_number):
            number = clear_exel(ru_sheet, row)
            text = str(ru_sheet.row(row)[1]).replace('text:', '').replace("'", "")
            rus.update({number: f'{text}'})

    first_num = str(en_sheet.row(1)[0]).replace("number:", '').replace('.0', '')
    first_text = str(en_sheet.row(1)[1]).replace('text:', '').replace("'", "")
    eng.update({first_num: f'{first_text}'})
    print(len(eng))
    print(len(rus))

    workbook = xlsxwriter.Workbook('allCodeTNved.xlsx')
    bold = workbook.add_format({'bold': True})
    center = workbook.add_format()
    center.set_align('center')
    worksheet = workbook.add_worksheet("code RU EN")
    worksheet.write('A1', 'Код', bold)
    worksheet.write('A1', 'Код', center)
    worksheet.write('B1', 'Название RU', bold)
    worksheet.write('B1', 'Название RU', center)
    worksheet.write('C1', 'Название EN', bold)
    worksheet.write('C1', 'Название EN', center)

    nums = 2

    for k1, v1 in rus.items():
        if k1 in eng:
            worksheet.write(f'A{nums}', k1)
            worksheet.write(f'B{nums}', v1)
            worksheet.write(f'C{nums}', eng[f"{k1}"])
            nums += 1
        else:
            worksheet.write(f'A{nums}', k1)
            worksheet.write(f'B{nums}', v1)
            worksheet.write(f'C{nums}', f'*****', bold)
            worksheet.write(f'C{nums}', None, center)
            nums += 1
            print(f'============================={k1, v1} нет в EN')

    print('!----------------------------------------------------------------------------------------------!')
    for k1, v1 in eng.items():
        if k1 in rus:
            pass
        else:
            worksheet.write(f'A{nums}', k1)
            worksheet.write(f'B{nums}', f'*****', bold)
            worksheet.write(f'B{nums}', None, center)
            worksheet.write(f'C{nums}', v1)
            nums += 1
            print(f'{k1, v1} нет в RU')

    workbook.close()


if __name__ == "__main__":
    main()
