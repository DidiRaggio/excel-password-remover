import os, sys
import glob
import win32com.client

def unprotect_xlsx(filename, password):
    xcl = win32com.client.Dispatch('Excel.Application')
    xcl.DisplayAlerts = False
    try:
        wb = xcl.workbooks.open(filename, 0, False, None,password, password)
        wb.Unprotect(password)
        wb.UnprotectSharing(password)
        wb.Save()
        xcl.Quit()

    except Exception as error:
        print("THERE WAS AN ERROR {}".format(error))


def get_xlsx_files(path=os.getcwd(), extension='xls*'):
    os.chdir(path+'\\excel_files\\')
    result = [path+'\\excel_files\\'+i for i in glob.glob('*.{}'.format(extension))]
    return result

def main(password="password"):
    xlsx_files = get_xlsx_files()
    for file in xlsx_files:
        unprotect_xlsx(file, password)

if __name__ == '__main__':
    print(len(sys.argv))
    if len(sys.argv) > 1:
        main(sys.argv[1])
    main()
