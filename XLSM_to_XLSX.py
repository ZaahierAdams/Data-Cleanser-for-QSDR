from win32com.client.gencache import EnsureDispatch

def XLSM_to_XLSX(parent_Dir, folder_name_2, folder_name_4, folder_name_5, file_n):
    
    excel = EnsureDispatch('Excel.Application')
    
    # Load the .XLSM file into Excel
    wb = excel.Workbooks.Open(parent_Dir
                              +'\\'
                              +folder_name_2
                              +'\\'
                              +folder_name_4
                              +'\\'
                              +file_n
                              +'.xlsm')
             
    #wb = excel.Workbooks.Open(r'C:\Users\User\Desktop\py work\Data verifier\VBA_Test\TEST1.xlsm')
    
    # Save it in .XLSX format to a different filename
    excel.DisplayAlerts = False
    wb.DoNotPromptForConvert = True
    wb.CheckCompatibility = False
    
    wb.SaveAs(parent_Dir
              +'\\'
              +folder_name_2
              +'\\'
              +folder_name_5
              +'\\'
              +file_n
              +'.xlsx'
              ,FileFormat=51
              ,ConflictResolution=2)
    
    #wb.SaveAs(r'C:\Users\User\Desktop\py work\Data verifier\VBA_Test\TEST1.xlsx', FileFormat=51, ConflictResolution=2)
    
    excel.Application.Quit()
