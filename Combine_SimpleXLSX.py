from os import chdir
from glob import glob
from pandas import read_excel, ExcelWriter

def Combine_XLSX(file_type,
                 DB_index_col,
                 file_n,
                 parent_Dir,
                 folder_name_2,
                 folder_name_3,
                 folder_name_7):
    
    
    ## Dir. of new file 
    comb_file_dir = (parent_Dir
                     + '\\''' 
                     + folder_name_2
                     + '\\''' 
                     + folder_name_7
                     + '\\''' 
                     + file_n
                     + file_type.replace('*','')
                     )
    print(comb_file_dir)
    
    ## List containing all iterated dataframes in directory ('Sheets')
    df_list=[]
    
    chdir(parent_Dir
             + '\\''' 
             + folder_name_2
             + '\\''' 
             + folder_name_3
             )
    print(parent_Dir
             + '\\''' 
             + folder_name_2
             + '\\''' 
             + folder_name_3
             )
    
    for file in glob(file_type):
        
        ## try/ except
        ##  in case there is no ID col. 
        try:
            sheety = read_excel(file, index_col = DB_index_col)
        except:
            sheety = read_excel(file)
        
        ## nested list containing:
        ##  (1) Sheet name
        ##  (2) DF 
        inside_list = [file, sheety]
        
        ## append nested list to DF list
        df_list.append(inside_list)
        
    
    ## Iterates through df list
    ## writes each to a unique sheet     
    with ExcelWriter(comb_file_dir) as writer:
        for file, sheety in df_list:
            sheety.to_excel(writer, 
                            sheet_name = file.replace(file_type.replace('*',''),''))
            