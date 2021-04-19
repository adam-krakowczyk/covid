from os.path import isfile
import openpyxl
from pathlib import Path


SOURCE_EXCEL_1 = 'test_kwarantanna1.xlsx'
SOURCE_EXCEL_2 = 'test_wynikidodatnie1.xlsx'
DESTINATION_EXCEL ='test_wynik_pasuje.xlsx'

print ('start processing')

def xls_compare(file1,file2):
        print ('[checking files exist]')
        if not isfile(file1):
            print('[-] File doesnt exist: ', file1)
            return
        if not isfile(file2):
            print('[-] File doesnt exist: ', file2)
            return
        print('[+] Files exist')

        xlsx_file = Path('', file1)
        wb_obj = openpyxl.load_workbook(xlsx_file)

        # Read the active sheet:
        sheet = wb_obj.active

        adres_1=[]

        for i, row in enumerate(sheet.iter_rows(values_only=True)):
            if i == 0:      # nagłówek
                pass
            else:
                adres_1.append(str(row[4]).upper().replace('UL. ','').replace('AL. ','').replace('PL. ','').replace(', KATOWICE','').replace('ALEJA ','').strip())
        print(adres_1)
        print('-'*20)
        print('załadowano z pliku : ',file1, ' adresów: ', len(adres_1))
        print('-'*20)

        xlsx_file2 = Path('', file2)
        wb_obj2 = openpyxl.load_workbook(xlsx_file2)

        # Read the active sheet:
        sheet2 = wb_obj2.active
        adres_2=[]

        for i, row in enumerate(sheet2.iter_rows(values_only=True)):
            if i == 0:
                pass
            else:
                if len(row[6])>0 and str(row[6]).isalnum():
                        adres_2.append(str(row[4]).upper().replace('UL. ','').replace('AL. ','').replace('\\','').replace('_','').lstrip()
                               + ' '
                               + row[5]
                               + '/'
                               + row[6])
                elif len(row[6])==0:
                        adres_2.append(str(row[4]).upper().replace('UL. ','').replace('AL. ','').replace('\\','').replace('_','').lstrip()
                               + ' '
                               + row[5])

        print(adres_2)
        print('-'*20)
        print('załadowano z pliku : ',file2, ' adresów: ', len(adres_2))
        print('-'*20)
        adres_to_call =[]
        for adres in adres_1:
            if adres in adres_2:
                adres_to_call.append(adres)
        print('Spójne adresy: ', len(adres_to_call))
        print(adres_to_call)


xls_compare(SOURCE_EXCEL_1,SOURCE_EXCEL_2)
