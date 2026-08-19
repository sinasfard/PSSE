import utils_standAlone as utils
from utils_standAlone import PSSE_PATH_OS, PSSE_PATH_SYS
import os,sys
import zipfile
import pyodbc
import psspy
from datetime import datetime
from log import MyLogger
import pandas as pd
import subprocess
from psspy import _i, _f, _s
from case_standAlone import ReadCase, get_case_df, reloadCase, saveCase, runIdevFromPSSE

class SQLDatabase:
    def __init__(self)->None:
        self.server = '''PRODSQL37\RBASQL'''
        self.database = '''NS_ShiftFactor'''
        self.connection()
        self.createTable()
        self.batchSize = 5000

    def createTable(self):
        # Check if the table exists
        table_exists_query = """
        SELECT * 
        FROM INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_NAME = 'ShiftFactor'
        """        
        self.cursor.execute(table_exists_query)
        table_exists = self.cursor.fetchone()
        if not table_exists:
            create_table_query = f"""
            CREATE TABLE ShiftFactor (
                ShiftFactorID INT IDENTITY(1,1) PRIMARY KEY,
                injectionBus INT,
                fromBus INT,
                toBus INT,
                genEffectivness real,
                branchID varchar(4),
                transmissionElement varchar(20),
                opposingBus INT,
                caseID INT
            )
            """    
            self.cursor.execute(create_table_query)
            self.conn.commit()   


    def writeTable(self, df):
        for i in range(0, len(df), self.batchSize):
            batch = df[i:i + self.batchSize]
            batch = batch.fillna("")

            # Insert data from the batch into the SQL table
            insert_query = f"INSERT INTO ShiftFactor VALUES (?, ?, ?, ?, ?, ?, ?, ?)"  
            params = [ tuple(row) for index, row in batch.iterrows()]
            self.cursor.executemany(insert_query, params)
            self.conn.commit()

    def connection(self):
        self.conn = pyodbc.connect(f'DRIVER={{SQL Server}};'
                                   f'Server={self.server};'
                                   f'Database=master;'
                                   f'Trusted_Connection=True'
                                    ,autocommit=True)
        self.cursor = self.conn.cursor()
        check_database_query = f"SELECT db_id('{self.database}')"
        self.cursor.execute(check_database_query)
        result = self.cursor.fetchone()
        if result[0] is None:
            # Create the target database if it doesn't exist
            create_database_query = f"CREATE DATABASE {self.database}"
            self.cursor.execute(create_database_query)
            self.conn.commit()
        self.conn.close()
        self.conn = pyodbc.connect(f'DRIVER={{SQL Server}};'
                                f'Server={self.server};'
                                f'Database={self.database};'
                                f'Trusted_Connection=True;'
                                    f'fast_executemany=True')
        self.cursor = self.conn.cursor()            
class ShiftFactor:
    def __init__(self) -> None:
        columns = {
                    'injectionBus': int,
                    'fromBus': int,
                    'toBus': int,
                    'genEffectivness': float,
                    'branchId': str,
                    'transmissionElement': str,
                    'opposingBus': int
                    }
        self.shiftFactor_df = pd.DataFrame(columns=columns.keys())
        self.directory = os.getcwd()
    def createSubFile(self, area, opposingBus):
        template = f"""/PSS(R)E 34
COM
COM SUBSYSTEM description file entry created by PSS(R)E Config File Builder
COM
SUBSYSTEM 'SOURCE'"""
        # Add Area entries
        for index, area_num in enumerate(area):
            template += f"\n AREA {area_num}"
        template += """
END
SUBSYSTEM 'SINK'"""
        # Add Area entries
        for index, area_num in enumerate(opposingBus):
            template += f"\n AREA {area_num}"
        template += """
END
COM
COM SUBSYSTEM description file entry created by PSS(R)E Config File Builder
COM
SUBSYSTEM 'MONITOR'"""    
        template += """\nKVRANGE 69.000 500.000
END
END"""

        filename = "temp.sub"
        filename = os.path.join(self.directory, filename)
        with open(filename, "w") as file:
            file.write(template)

    def creatConFile(self):
        content = """/PSS(R)E 35
COM
COM CONTINGENCY description file entry created by PSS(R)E Config File Builder
COM
SINGLE BRANCH IN SUBSYSTEM 'SOURCE'
END"""
        filename = "temp.con"
        filename = os.path.join(self.directory, filename)
        with open(filename, "w") as file:
            file.write(content)


    def creatMonFile(self):
        content = """/PSS(R)E 35
COM
COM MONITORED element file entry created by PSS(R)E Config File Builder
COM
MONITOR BRANCHES IN SUBSYSTEM 'MONITOR'
END"""
        filename = "temp.mon"
        filename = os.path.join(self.directory, filename)
        with open(filename, "w") as file:
            file.write(content)


    def createShiftFactorTable(self, filtered_df = None, opposingBus = None):
        rows = []
        filename = os.path.join(self.directory, 'results.sf')
        with open(filename, 'r') as file:
            readFromBus = False
            readToBus = False
            readBuses = False
            fromBus = -1
            toBus = -1
            branch_id = -1
            
            for line in file:
                line_data = line.split()
                
                if "SENSITIVITY FACTORS OF BRANCH FLOW (MW) ON " in line:
                    readBuses = False
                    next_index = 0
                    for index, str1 in enumerate(line_data):
                        if next_index < len(line_data) - 1:
                            next_index += 1
                            if str1 == 'ON':
                                readFromBus = True
                                continue
                            if readFromBus:
                                fromBus = int(str1)
                                readFromBus = False
                                continue
                            if str1 == "TO":
                                readToBus = True
                                continue
                            if readToBus:
                                try:
                                    toBus = int(str1)
                                except:
                                    toBus = -1
                                readToBus = False
                                continue
                            if line_data[next_index] == "MORE":
                                branch_id = str1
                    
                if "<----" in line:
                    readBuses = True
                elif readBuses and len(line_data) > 2:
                    effectiveness = 0
                    if fromBus<toBus: 
                        effectiveness = round(float(line_data[-1]), 3)
                    else: 
                        effectiveness = -1*round(float(line_data[-1]), 3)
                    rows.append({
                        'injectionBus': int(line_data[0]),
                        'fromBus': min(fromBus,toBus),
                        'toBus': max(fromBus,toBus),
                        'genEffectivness': effectiveness,
                        'branchId': branch_id
                    })
                else:
                    readBuses =  False
        temp_df = pd.DataFrame(rows)
        
        temp_df = pd.merge(temp_df, filtered_df, how = 'inner', left_on=['fromBus', 'toBus', 'branchId'], right_on=['Bus', 'Bus2', 'ID'])
        temp_df = temp_df[['injectionBus','fromBus','toBus','genEffectivness', 'branchId']]
        temp_df['opposingBus'] = opposingBus
        self.shiftFactor_df = pd.concat([self.shiftFactor_df, temp_df], ignore_index=True)

    def removeFiles(self):
        filename = os.path.join(self.directory, 'results.sf')
        if os.path.exists(filename):
            os.remove(filename)
        filename = os.path.join(self.directory, 'temp.mon')
        if os.path.exists(filename):
            os.remove(filename)
        filename = os.path.join(self.directory, 'temp.con')
        if os.path.exists(filename):
            os.remove(filename)
        filename = os.path.join(self.directory, 'temp.sub')
        if os.path.exists(filename):
            os.remove(filename)   
        filename = os.path.join(self.directory, 'temp.dfx')
        if os.path.exists(filename):
            os.remove(filename)  
	
    def Solve(self,i=0):
        islands=psspy.tree(1,-1)[1]
        while islands > 0:
            islands=psspy.tree(2,1)[1]
        psspy.fnsl([0,0,0,0,2,0,0,0])
        psspy.fnsl([0,0,0,0,2,0,0,0])
        if psspy.solved() == 0:
            return True
        elif psspy.solved()==1 and i<10:
            i += 1
            self.Solve(i)
            if psspy.solved()==0:
                return True
        else:
            return False    
                 
    def process(self, caseID, workingCase_sav, req_issue_dt):

        #sql = SQLDatabase()
        psspy.psseinit(16000) 
        _i=psspy.getdefaultint()
        _f=psspy.getdefaultreal()
        _s=psspy.getdefaultchar()
        #workingCase_sav = r"T:\Rshared\Delivery Shared\Python Files\Sina\GitRepo\cas-automation-tools\CPAT GUI\Functions\Code_test\powerflow\S5_NW_2029_SP_HG_Post.sav"
        psspy.case(workingCase_sav) 
        self.Solve()


	



        OutputDirectory = os.path.dirname(workingCase_sav)
        self.directory  = OutputDirectory
        working_case = ReadCase(branches=True)
        working_case_df = get_case_df(working_case)           
        areas = {4:[17,18,25],52:[17,18,25],48:[17,18,25],47:[17,18,25],43:[17,18,25],45:[17,18,25],49:[17,18,25],54:[17,18,25],55:[17,18,25],53:[17,18,25],46:[17,18,25],44:[17,18,25], 57:[17,18,25], 6:[17,18,25],40:[4,55], 60:[4,55], 31:[4,55],28:[22,20,17], 56:[22,20,17], 13:[22,20,17], 32:[22,20,17], 37:[22,20,17], 36:[22,20,17], 42:[22,20,17], 35:[22,20,17], 39:[22,20,17], 29:[13,37], 30:[48,4], 38:[48,4], 34:[48,4], 25:[4], 27:[55,53], 33:[55,53], 17:[55,53], 18:[55,53], 19:[55,53], 20:[55,53], 21:[55,53], 22:[55,53], 23:[55,53], 24:[55,53], 26:[55,53] }
        # areeas ={55:1036}
        # Create a new dictionary with values as keys and keys as values
        reversed_areas = {}
        for key, value in areas.items():
            if str(value) not in reversed_areas:
                reversed_areas[str(value)] = [key]
            else:
                reversed_areas[str(value)].append(key)        
        areasWithSameOpposingBus = [keys for keys in reversed_areas.values() if len(keys) > 1]
        print(areasWithSameOpposingBus)
        d=0
        for areaSet in areasWithSameOpposingBus:
            opposingBus = areas[areaSet[0]]
            self.createSubFile(areaSet, opposingBus)
            self.creatConFile()
            self.creatMonFile()            
            print(f'opposingBus is {opposingBus}')
            ierr= psspy.report_output(2,os.path.join(self.directory, "results.sf"),[0,0])
            ierr= psspy.dfax_2([1,0,0],os.path.join(self.directory, "temp.sub"),os.path.join(self.directory, "temp.mon"),os.path.join(self.directory, "temp.con"),os.path.join(self.directory, "temp.dfx"))

            ierr= psspy.sensitivity_flows([0,0],[0,1,0,0,0,0,0,0,0],[ 0.5, 0.03],[r"""SOURCE""","",r"""MONITOR"""],os.path.join(self.directory, "temp.dfx"))
    
            ierr = psspy.close_report()
            working_case_df['Area']=working_case_df['Area'].astype('int32')
            filtered_df =working_case_df.loc[~working_case_df['Area'].isin(opposingBus)]
            self.createShiftFactorTable(filtered_df) 
            self.removeFiles()
        brn_real = psspy.abrnchar(-1, 1, 3, 3, 1, ["BRANCHNAME","ID"])[1]
        brn_int = psspy.abrnint(-1, 1, 3, 3, 1, ["FROMNUMBER", "TONUMBER"])[1]
        line_df = pd.DataFrame(zip(brn_real[0], brn_real[1],brn_int[0], brn_int[1] ), columns=['BranchName', 'branchId', 'fromBus', 'toBus'])
        self.shiftFactor_df = self.shiftFactor_df.merge(line_df, how='left', on=[ 'branchId', 'fromBus', 'toBus'])
        # mach_real = psspy.amachreal(-1, 1, ["PGEN"])[1]
        # mach_int = psspy.amachint(-1, 1, ["NUMBER"])[1]
        # machine_df = pd.DataFrame(zip(mach_real[0], mach_int[0]), columns=['DispatchedGen','injectionBus'])       
        # self.shiftFactor_df = self.shiftFactor_df.merge(machine_df, how='left', on=[ 'injectionBus'])
        ierr = psspy.close_powerflow()
        self.shiftFactor_df['caseID'] = caseID
        outputPath = os.path.join(OutputDirectory, f"{req_issue_dt}.csv")
        self.shiftFactor_df.to_csv(outputPath)
        # sql.writeTable(self.shiftFactor_df)
if __name__ == "__main__":
        # dirname = sys.argv[1]
    local_run = 0
    try:
        if local_run == 1:
            pass
        else:
            session_id = sys.argv[1]
            workingCase_sav = sys.argv[2]
            req_issue_dt = sys.argv[3]

        SF = ShiftFactor()
        SF.process(session_id, workingCase_sav, req_issue_dt)
    except:
        pass
