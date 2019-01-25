from shutil import copyfile
from win32com import client
from sqlite3 import connect
from collections import namedtuple
from os.path import basename

def copy_mdd_ddf_data(input_path, output_path, only_mdd=False):
    
    copyfile(f'{input_path}.mdd', f'{output_path}.mdd')
    if not only_mdd:
        copyfile(f'{input_path}.ddf', f'{output_path}.ddf')
        mdd = client.Dispatch('MDM.Document')
        mdd.Open(f'{output_path}.mdd')
        mdd.DataSources.Default.DBLocation =  basename(f'{output_path}.ddf')
        mdd.Save()
        mdd.Close()

################################
#
# TRANSFERING DATA IN ONE BLOCK
#
################################

class BlockTransferer:

    def __init__(self, mdd_path, ddf_path, block_name):
        self.mdd_path = mdd_path
        self.ddf_path = ddf_path
        self.block_name = block_name

    def update_mdd(self):
        
        mdd = client.Dispatch('MDM.Document')
        mdd.Open(self.mdd_path)

        # adds block_name as prefix to all types
        for t in mdd.Types:
            t.Name = f'{self.block_name}_{t.Name}'

        # saves script for all fields except system variables
        fields_script = '\r\n'.join(f.Script for f in mdd.Fields if not f.IsSystem)

        # removes all fields except system variables
        for f in mdd.Fields:
            if not f.IsSystem:
                mdd.Fields.Remove(f.Name)

        # adds new block
        new_block = mdd.CreateClass(self.block_name)
        mdd.Fields.Add(new_block)

        # adds fields script to the new block
        new_block.Fields.AddScript(fields_script)

        mdd.Save()
        mdd.Close()

    def update_ddf(self):

        # uses mdd to read the list of system variables
        mdd = client.Dispatch('MDM.Document')
        mdd.Open(self.mdd_path)
        system_variables = [v.FullName for v in mdd.Variables if v.IsSystemVariable]
        mdd.Close()
        
        # sets up sqlite database
        sqlite_conn = connect(self.ddf_path)
        sqlite_cursor = sqlite_conn.cursor()

        # update Levels table
        sqlite_cursor.execute(f"""
            UPDATE Levels
            SET DSCTableName = '{self.block_name}.' || DSCTableName
            WHERE ParentName = 'L1'""")

        # read columns from L1 table
        SQLiteVariable = namedtuple('SQLiteVariable', 'cid name type notnull dflt_value pk')
        current_variables = [SQLiteVariable(*row) for row in sqlite_cursor.execute('pragma table_info(L1)')]

        # builds list of renamed variables
        new_variables = [
            SQLiteVariable(
                v.cid,
                v.name if (v.name.split(':')[0] in system_variables
                    or v.name[0] == ':') else f'{self.block_name}.{v.name}',
                v.type,
                v.notnull,
                v.dflt_value,
                v.pk)
            for v in current_variables
        ]

        #renames old L1 table
        sqlite_cursor.execute('ALTER TABLE L1 RENAME TO temp_L1')   

        # creates new L1 table
        create_table_statement = '''CREATE TABLE L1 (
            ''' + ', '.join(
                    [f'[{v.name}] {v.type} {"not" if v.notnull else ""} null {"unique" if v.pk else ""}'
                    for v in new_variables]
                ) + ')'
        sqlite_cursor.execute(create_table_statement)

        # transfers data
        sqlite_cursor.execute('''
            INSERT INTO L1 
            SELECT *
            FROM temp_L1
        ''')

        # deletes old table
        sqlite_cursor.execute('DROP TABLE temp_L1')

        # commits transaction
        sqlite_conn.commit()
        sqlite_conn.close()


def remove_helper_fields(mdd_path):

    mdd = client.Dispatch('MDM.Document')
    mdd.Open(mdd_path)

    parent_fields = reversed([f.FullName for f in mdd.Fields.Expanded if len(f.HelperFields) > 0])

    for f in parent_fields:
        helper_fields = [hf.Name for hf in mdd.Fields[f].HelperFields if not hf.IsSystem]
        for hf in helper_fields:
            mdd.Fields[f].HelperFields.Remove(hf)

    mdd.Save()
    mdd.Close()