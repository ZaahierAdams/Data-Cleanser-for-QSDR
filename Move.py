from shutil import move, copy

def Move_file(parent_Dir, folder_name_2, folder_name_4, file_n):
    
    dir_1 = (parent_Dir
             + '\\'
             + folder_name_2
             + '\\'
             + file_n
             + '.xlsm'
             )
    
    dir_2 = (parent_Dir
             + '\\'
             + folder_name_2
             + '\\'
             + folder_name_4
             + '\\'
             + file_n
             + '.xlsm'
             )

    move(dir_1, dir_2)


def Copy_file(source_dir, target_dir):
    copy(source_dir, target_dir)
    
def Copy_file_2(parent_Dir, folder_name_2, folder_name_4, file_n):
    dir_1 = (parent_Dir
             + '\\'
             + folder_name_2
             + '\\'
             + file_n
             + '.xlsm'
             )
    
    dir_2 = (parent_Dir
             + '\\'
             + folder_name_2
             + '\\'
             + folder_name_4
             + '\\'
             + file_n
             + '.xlsm'
             )

    copy(dir_1, dir_2)