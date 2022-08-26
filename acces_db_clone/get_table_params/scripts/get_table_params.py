import pandas as pd
import os
import pyodbc

paths = os.path.abspath(os.getcwd())

def leer_paths(key, paths=paths):
    '''
    Read path.xlsx file wich contains key input respective absolute folder path

    input:
        key: keyname of the path to search for (str)
        paths*: basolute path where the path.xlsx file is located default is script location folder (str)

    output:
        file_path: Location of the corresponding key input (str)
    '''
    for file in os.listdir(paths):
            file_path = os.path.join(paths,file)

            if 'path.xlsx' in file_path:
                df = pd.read_excel(file_path)
    
    return df[key][0]


def crear_paths():
    '''
    Create xlsx file with parent folder's folder paths dataframe at actual script execution folder
    '''
    local_path = os.path.abspath(os.getcwd())
    os.chdir('../')
    parent_path = os.path.abspath(os.getcwd())
    os.chdir(local_path)
    paths = []
    folders = []
    for file in os.listdir(parent_path):
                file_path = os.path.join(parent_path,file)
                paths.append(file_path)
                folders.append(file)

    ruta = dict(zip(folders, paths))

    ruta_df = pd.DataFrame([ruta])

    ruta_df.to_excel('path.xlsx', index=False)
    print('path.xlsx creado en ' + local_path)

def get_acces_tablenames(filepath):
    '''
    Read .mdb or .accdb file in filepath provided and returns all user created table names in a list\
    
    input: filepath (str)
    output: table_names (list[(str)])
    '''
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + filepath + ';')
    cursor = conn.cursor()
    table_names = []
    for row in cursor.tables():
        tname = row.table_name
        if ('MSys' not in tname) & ('qry' not in tname):
            table_names.append(tname)

    conn.close()
    return table_names

def get_acc_table_descr(acc_path,table):

    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + acc_path + ';')
    cursor = conn.cursor()

    qry = 'SELECT * FROM ' + table

    cursor.execute(qry)

    cols = [column[0] for column in cursor.description]
    a = [column for column in cursor.description]
    df = pd.DataFrame(a,columns=['col_name', 'type_code', 'display_size', 'internal_size', 'precision', 'scale', 'null_ok'],)

    conn.close()

    df = df[df.keys()].astype(str)


    return df

crear_paths()
acc_file = leer_paths('acc_file')
tables = leer_paths('tables')
scripts = leer_paths('scripts')

for file in os.listdir(acc_file):
                file_path = os.path.join(acc_file,file)
                filename = file.split('.')[0]

                tb = get_acces_tablenames(file_path)

                print('---- Creando '+ filename +' ----')

                for table in tb:

                    print_path = os.path.join(tables,filename + '.xlsx')
                    


                    with pd.ExcelWriter(print_path) as writer:
                        df = get_acc_table_descr(file_path,table)
                        df.to_excel(writer,sheet_name=table,index=False)

print('--- Proceso Completado ---')
