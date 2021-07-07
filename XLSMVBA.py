from __future__ import print_function
from pandas import ExcelWriter
from os.path import expanduser
from win32com.client import DispatchEx

def XLSM_Create(parent_Dir, folder_name_2, file_n):
    
    filename = (parent_Dir 
                + '\\' 
                + folder_name_2 
                + '\\'
                + file_n
                + '.xlsx')
    
    writer = ExcelWriter(filename, engine='xlsxwriter')
    
    
    #ATTACH MACRO AND CREATE .XLSM FILE# 
    filename_macro = (parent_Dir 
                      + '\\' 
                      + folder_name_2 
                      + '\\'
                      + file_n
                      + '.xlsm')
    
    workbook = writer.book
    workbook.filename = filename_macro
    
    workbook.add_vba_project(parent_Dir 
                             + '\\' 
                             + folder_name_2 
                             + '\\'
                             + 'vbaProject.bin')

    writer.save()

def XLSM_Run(parent_Dir, folder_name_2, file_n):
    
    VBA_fail = False
    #import unittest
    #class ExcelMacro(unittest.TestCase):
    #    def test_excel_macro(self):
    try:
        xlApp = DispatchEx('Excel.Application')
        
        xlsPath = expanduser(parent_Dir 
                                     + '\\' 
                                     + folder_name_2 
                                     + '\\'
                                     + file_n
                                     + '.xlsm')
        #xlsPath = os.path.expanduser("C:\\Users\\User\\Desktop\\py work\\Data verifier\\VBA_Test\\TEST1.xlsm")
        
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        xlApp.Run('GetSheets')
        wb.Save()
        xlApp.Quit()
        print("Macro ran successfully!")
    except:
        VBA_fail = True
        print("Error found while running the excel macro!")
        xlApp.Quit()
    #if __name__ == "__main__":
    #    unittest.main()
    
    return VBA_fail
       
    
    
    
#def Create2():   
#    import win32com.client as win32
#    import win32con
#    import win32gui
#    import time
#    
#    
#    import comtypes, comtypes.client
#    
#    
#    excel = win32.gencache.EnsureDispatch("Excel.Application")
#    #excel.Visible = True
#    workbook = excel.Workbooks.Add()
#    #sh = workbook.ActiveSheet
#    
#    xlmodule = workbook.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
#    
#    sCode = '''
#    Sub GetSheets()
#    'Updated by Extendoffice 2019/2/20
#    Path = "C:\\Users\\User\\Desktop\\py work\\Data verifier\\New_Excel_Files\\"
#    Filename = Dir(Path & "*.xls")
#      Do While Filename <> ""
#      Workbooks.Open Filename:=Path & Filename, ReadOnly:=True
#         For Each Sheet In ActiveWorkbook.Sheets
#         Sheet.Copy After:=ThisWorkbook.Sheets(1)
#      Next Sheet
#         Workbooks(Filename).Close
#         Filename = Dir()
#      Loop
#    End Sub
#    
#    
#    '''
#    
#    xlmodule.CodeModule.AddFromString(sCode)
