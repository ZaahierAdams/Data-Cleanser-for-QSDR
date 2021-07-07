from xlsxwriter import Workbook
from datetime import date, datetime

def Dirt_Report(parent_Dir, 
                folder_name_8,
                all_duplicate_rows,
                all_no_err_ID,
                all_unverify_ID,
                all_mis_age,
                all_mis_gender,
                all_mis_citizen,
                all_mis_youth,
                all_mis_name,
                all_mis_sname,
                all_mis_persal,
                all_mis_SLvL,
                all_mis_race,
                all_mis_disab,
                counter_all_sheets,
                counter_book,
                string_files,
                name,
                f_type,
                khangela,
                all_mis_occlvl,
                all_mis_jobtit,
                all_mis_res,
                all_unveri_TI,
                all_intervention,
                all_type,
                all_NQF,
                all_duration,
                all_provider,
                all_res_issues
                ):
                ## name = "All", "file name", "sheet name"
                ## f_type:





    if f_type == 3:
        wrkbk = Workbook(parent_Dir
                                    + '\\''' 
                                    + folder_name_8
                                    + '\\''Report - '
                                    + string_files 
                                    + ' ('
                                    + name
                                    + ')'
                                    + '.xlsx')
    elif f_type == 2:
        wrkbk = Workbook(parent_Dir
                                    + '\\''' 
                                    + folder_name_8
                                    + '\\''Report - '
                                    + name 
                                    + '.xlsx')
    
    else: 
        # if ftype == 1
        wrkbk = Workbook(parent_Dir
                                    + '\\''' 
                                    + folder_name_8
                                    + '\\''Report - '
                                    + name 
                                    + ' '
                                    + datetime.now().strftime('(%d-%m-%Y at %H.%M)')
                                    + '.xlsx')
    
    wrksht = wrkbk.add_worksheet()
    
    cell_format_1 = wrkbk.add_format({'bold': True, 'font_color': 'black'})
    wrksht.set_column(0, 0, 30)
    wrksht.set_column(1, 1, 10, cell_format_1)
    
    cell_format_2 = wrkbk.add_format()
    cell_format_2.set_font_color('#FFFFFF')
    cell_format_2.set_font_size(12)
    cell_format_2.set_bold()
    cell_format_2.set_pattern(1)  
    cell_format_2.set_bg_color('#02319A')
                               
    cell_format_3 = wrkbk.add_format()
    cell_format_3.set_pattern(1)  
    cell_format_3.set_bg_color('#AEE8FC')
    
                               
                               
                               
    wrksht.write('A1', 'ERRORS FIXED',                      cell_format_2) 
    wrksht.write('A2', '',                                  cell_format_2) 
    wrksht.write('A3', 'Source of Error',                   cell_format_2) 
    wrksht.write('A4', 'Duplicate rows',                    cell_format_3)
    wrksht.write('A5', 'ID Number',                         cell_format_3)
    wrksht.write('A6', 'Age blunders',                      cell_format_3)
    wrksht.write('A7', 'Gender blunders',                   cell_format_3)
    wrksht.write('A8', 'Citizen blunders',                  cell_format_3)
    wrksht.write('A9', 'Youth blunders',                    cell_format_3)
    wrksht.write('A10', 'First Name',                       cell_format_3)
    wrksht.write('A11', 'Surname' ,                         cell_format_3)
    wrksht.write('A12', 'Persal Number',                    cell_format_3)
    wrksht.write('A13', 'Salary Level' ,                    cell_format_3)
    wrksht.write('A14', 'Race',                             cell_format_3)
    wrksht.write('A15', 'Disability',                       cell_format_3)
    
    wrksht.write('A16', 'Occupation level',                 cell_format_3)
    wrksht.write('A17', 'Job Title',                        cell_format_3)
    wrksht.write('A18', 'Residence',                        cell_format_3)
    
    wrksht.write('A19', 'Unverified Training Intervention', cell_format_3)
    wrksht.write('A20', 'Intervention Name',                cell_format_3)
    wrksht.write('A21', 'Intervention Type',                cell_format_3)
    wrksht.write('A22', 'Intervention NQF',                 cell_format_3)
    wrksht.write('A23', 'Intervention Duration',            cell_format_3)
    wrksht.write('A24', 'Intervention Provider',            cell_format_3)
    
    wrksht.write('A25', 'Residual Issues',                  cell_format_3)
    
    wrksht.write('A26', 'All Sheets',                       cell_format_2)
    wrksht.write('A27', 'All Books',                        cell_format_2)
    
    
    wrksht.write('B1', '',                                  cell_format_2) 
    wrksht.write('B2', '',                                  cell_format_2) 
    wrksht.write('B3', 'Frequency',                         cell_format_2) 
    wrksht.write('B4',  all_duplicate_rows,                 cell_format_3)
    wrksht.write('B5',  all_no_err_ID,                      cell_format_3)
    wrksht.write('B6',  all_mis_age,                        cell_format_3)
    wrksht.write('B7',  all_mis_gender,                     cell_format_3)
    wrksht.write('B8',  all_mis_citizen,                    cell_format_3)
    wrksht.write('B9',  all_mis_youth,                      cell_format_3)
    wrksht.write('B10', all_mis_name,                       cell_format_3)
    wrksht.write('B11', all_mis_sname,                      cell_format_3)
    wrksht.write('B12', all_mis_persal,                     cell_format_3)
    wrksht.write('B13', all_mis_SLvL,                       cell_format_3)
    wrksht.write('B14', all_mis_race,                       cell_format_3)
    wrksht.write('B15', all_mis_disab,                      cell_format_3)
    
    wrksht.write('B16', all_mis_occlvl,                     cell_format_3)
    wrksht.write('B17', all_mis_jobtit,                     cell_format_3)
    wrksht.write('B18', all_mis_res,                        cell_format_3)
    
    wrksht.write('B19', all_unveri_TI,                      cell_format_3)
    wrksht.write('B20', all_intervention,                   cell_format_3)
    wrksht.write('B21', all_type,                           cell_format_3)
    wrksht.write('B22', all_NQF,                            cell_format_3)
    wrksht.write('B23', all_duration,                       cell_format_3)
    wrksht.write('B24', all_provider,                       cell_format_3)

    wrksht.write('B25', all_res_issues,                     cell_format_3)

    wrksht.write('B26', counter_all_sheets,                 cell_format_2)
    wrksht.write('B27', counter_book,                       cell_format_2)
    
    
    wrksht.write('A29', 'Workbooks:', cell_format_1)
    wrksht.write('A30', string_files)
    
    khangela = list(set(khangela))
    string_khangela = ''
    for jonga in khangela:
        string_khangela = string_khangela + jonga + ', '
    
    
    wrksht.write('A32', 'Unverified ID Numbers', cell_format_2)
    wrksht.write('B32', all_unverify_ID, cell_format_3)
    wrksht.write('A33', 'Contains Unverified ID Numbers:', cell_format_1)
    wrksht.write('A34', string_khangela)

    
    
    wrkbk.close() 