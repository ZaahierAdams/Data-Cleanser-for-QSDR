from os import listdir, path, unlink, remove
from shutil import rmtree

def Clear_Directory(parent_Dir, folder_name_2, folder_name_3):
    '''
    Scorch all files in specified dir.
    '''
    
    folder = (parent_Dir
              +'\\'
              +folder_name_2
              +'\\'
              +folder_name_3)
    
    for filename in listdir(folder):
        file_path = path.join(folder, filename)
        try:
            if path.isfile(file_path) or path.islink(file_path):
                unlink(file_path)
            elif path.isdir(file_path):
                rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
            

def Clear_file_type(parent_Dir, folder_name, file_extension):
    '''
    delete files in 1 level sub dir by extension 
    '''
    
    folder = (parent_Dir + '\\' + folder_name)
    

    for filename in listdir(folder):
        if filename.endswith(file_extension):
            remove(path.join(folder, filename))
        else:
            pass
            