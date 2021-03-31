from win32com.client import Dispatch
import os
import pyodbc
import build
# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

dbname = r'C:/Users/geoff.ritchey/Desktop/NewDB.mdb'

connTime = pyodbc.connect(
    '''
    DRIVER={{SQL Server}};
    SERVER={2};
    DATABASE={0};
    UID=Avatar;
    PWD={1};
    '''.format(build.time_database, build.avatar_password, build.time_server)
)


class Fk:
    def __init__(self, pk, fk):
        self.pk = ''
        self.fk = ''

    def add(self, pk, fk):
        self.pk = self.pk + pk
        self.fk = self.fk + fk


def create_access():
    try:
        accApp = Dispatch("Access.Application")
        dbEngine = accApp.DBEngine
        workspace = dbEngine.Workspaces(0)

        dbLangGeneral = ';LANGID=0x0409;CP=1252;COUNTRY=0'
        newdb = workspace.CreateDatabase(dbname, dbLangGeneral, 64)

        curs = connTime.cursor()

        primary_keys = {}
        foreign_keys = {}
        for row in curs.tables(tableType='TABLE', schema='dbo').fetchall():
            query = "create table " + row.table_name + "("
            for column_data in curs.columns(table=row.table_name, schema='dbo').fetchall():
                if column_data.type_name in ('nvarchar', 'varchar'):
                    if column_data.column_size > 255:
                        query = query + f'[{column_data.column_name}] memo,'
                    else:
                        query = query + f'[{column_data.column_name}] varchar({column_data.column_size}),'
                elif column_data.type_name in ('numeric'):
                    query = query + f'[{column_data.column_name}] float,'
                elif column_data.type_name in ('tinyint'):
                    query = query + f'[{column_data.column_name}] smallint,'
                elif column_data.type_name in ('bit', 'datetime', 'int'):
                    query = query + f'[{column_data.column_name}] {column_data.type_name},'
                elif column_data.type_name in ('numeric() identity', 'int identity'):
                    query = query + f'[{column_data.column_name}] int,'
                else:
                    print(f"name = {column_data.column_name}: type={column_data.type_name}: size={column_data.column_size}")
            query = query[:-1] + ");"
            print(query)
            newdb.Execute(query)
            for primary_key in curs.primaryKeys(table=row.table_name):
                primary_keys.update({primary_key.table_name: str(primary_keys.get('primary_key.table_name') or '') + f", {primary_key.column_name}"})
            for foreign_key in curs.foreignKeys(table=row.table_name):
                foreign_keys.update({foreign_key.fk_name: str(foreign_keys.get('foreign_key.fk_name') or Fk()).add(f", {primary_key.column_name}"})
                print(foreign_key)

        for item, value in primary_keys.items():
            query = f"alter table {item} add primary key ([{value[2:]}]);"
            print(query)
            newdb.Execute(query)

    except Exception as e:
        print(e)

    finally:
        accApp.DoCmd.CloseDatabase
        accApp.Quit
        newdb = None
        workspace = None
        dbEngine = None
        accApp = None


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    os.remove(dbname)
    create_access()
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
