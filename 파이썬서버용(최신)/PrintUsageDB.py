
import pymssql
import pyodbc 
conn = pymssql.connect(server="10.10.4.37:14233",user='genuine',password='ASD!!axzC$',database='EMSF', as_dict=True)
#conn = pyodbc.connect(driver='{SQL Server}',server='tcp:10.10.4.37,14233',database='EMSF', uid='genuine', pwd='ASD!!axzC$')
global cursor
cursor = conn.cursor()

def FindQuery(Part,Arguments):
    Query = """select B.DeptName, A.EmpName from _TDAEmp AS A             #not use
                  LEFT OUTER join _tcaUser AS C WITH(NOLOCK)  ON A.EmpSeq = C.EmpSeq
                  LEFT OUTER join _tdadept AS B WITH(NOLOCK)  ON C.DeptSeq = B.DeptSeq
                  where A.empid = '{}'""".format(Arguments)
                  
    Query2 = """SELECT	B.DeptName, A.EmpName FROM	_TDAEmp AS A  
		        JOIN _tcaUser AS C WITH(NOLOCK) ON A.EmpSeq = C.EmpSeq
		        JOIN _tdadept AS B WITH(NOLOCK) ON C.DeptSeq = B.DeptSeq
                WHERE	A.empid = '{}'
                UNION ALL
                SELECT	 T1.ValueText	AS DeptName
		                ,T0.Remark		AS EmpName
                FROM		_TDAUMinor	AS	T0	WITH (NOLOCK)
		        LEFT OUTER JOIN	_TDAUMinorValue	AS	T1	WITH (NOLOCK)	ON	T0.MajorSeq = T1.MajorSeq	AND	T0.MinorSeq	=	T1.MinorSeq		AND	T1.Serl	=	1000001
                WHERE	T0.CompanySeq = '1'		
                AND		T0.MajorSeq	= 1026002
                AND		T0.IsUse = '1'
                AND		T0.MinorName = '{}'""".format(Arguments,Arguments) #not use
    if Part == '에몬스앳홈':
        Query3 = """SELECT   T0.MinorName AS UserID, T0.Remark AS EmpName, T1.ValueText AS DeptName
                FROM		_TDAUMinor	AS	T0	WITH (NOLOCK)
		        LEFT OUTER JOIN	_TDAUMinorValue	AS	T1	WITH (NOLOCK)	ON	T0.MajorSeq = T1.MajorSeq	AND	T0.MinorSeq	=	T1.MinorSeq		AND	T1.Serl	=	1000001
                WHERE	T0.CompanySeq = '1'		
                AND		T0.MajorSeq	= 1026002
                AND		T0.IsUse = '1'
                AND 	T0.MinorName = '{}'""".format(Arguments)
    elif Part == '개발실':
        Query3 = """SELECT   T0.MinorName AS UserID, T0.Remark AS EmpName, T1.ValueText AS DeptName
                FROM		_TDAUMinor	AS	T0	WITH (NOLOCK)
		        LEFT OUTER JOIN	_TDAUMinorValue	AS	T1	WITH (NOLOCK)	ON	T0.MajorSeq = T1.MajorSeq	AND	T0.MinorSeq	=	T1.MinorSeq		AND	T1.Serl	=	1000001
                WHERE	T0.CompanySeq = '1'		
                AND		T0.MajorSeq	= 1026079
                AND		T0.IsUse = '1'
                AND 	T0.MinorName = '{}'""".format(Arguments)
                
    return Query3



def dbConnect():
    #conn = pyodbc.connect(driver='{SQL Server}',server='tcp:10.10.4.37,14233',database='EMSF', uid='genuine', pwd='ASD!!axzC$')
    conn = pymssql.connect(server="10.10.4.37:14233",user='genuine',password='ASD!!axzC$',database='EMSF', as_dict=True)
    cursor = conn.cursor()

def selectIndividual(Part):
    if Part == '에몬스앳홈':
        global query
        query = '''SELECT	B.DeptName, A.EmpName, A.Empid as UserID FROM	_TDAEmp AS A 
		        JOIN _tcaUser AS C WITH(NOLOCK) ON A.EmpSeq = C.EmpSeq
		        JOIN _tdadept AS B WITH(NOLOCK) ON C.DeptSeq = B.DeptSeq
                WHERE	A.empid = '{}'
                UNION ALL
                SELECT	 T1.ValueText	AS DeptName
		                ,T0.Remark		AS EmpName
		                ,T0.MinorName   AS UserID
								                
                FROM		_TDAUMinor	AS	T0	WITH (NOLOCK)
		        LEFT OUTER JOIN	_TDAUMinorValue	AS	T1	WITH (NOLOCK)	ON	T0.MajorSeq = T1.MajorSeq	AND	T0.MinorSeq	=	T1.MinorSeq		AND	T1.Serl	=	1000001
                WHERE	T0.CompanySeq = '1'		
                AND		T0.MajorSeq	=	1026002
                AND		T0.IsUse = '1'
                --AND		T0.MinorName = '{}' 
            '''
       
            
    elif Part == '개발실':
        query = '''SELECT	B.DeptName, A.EmpName, A.Empid as UserID FROM	_TDAEmp AS A 
		        JOIN _tcaUser AS C WITH(NOLOCK) ON A.EmpSeq = C.EmpSeq
		        JOIN _tdadept AS B WITH(NOLOCK) ON C.DeptSeq = B.DeptSeq
                WHERE	A.empid = '{}'
                UNION ALL
                SELECT	 T1.ValueText	AS DeptName
		                ,T0.Remark		AS EmpName
		                ,T0.MinorName   AS UserID
								                
                FROM		_TDAUMinor	AS	T0	WITH (NOLOCK)
		        LEFT OUTER JOIN	_TDAUMinorValue	AS	T1	WITH (NOLOCK)	ON	T0.MajorSeq = T1.MajorSeq	AND	T0.MinorSeq	=	T1.MinorSeq		AND	T1.Serl	=	1000001
                WHERE	T0.CompanySeq = '1'		
                AND		T0.MajorSeq	= 1026079
                AND		T0.IsUse = '1'
                --AND		T0.MinorName = '{}' 
            '''
    cursor.execute(query)
    result = cursor.fetchall()
    return result        
    
    

def dbInsertIndividual(NDate,UserID,DeptName,UserName,MonoPrint,ColorPrint,MonoCopy,ColorCopy,DeviceID,InsertTime):
    #query = '''INSERT INTO emsf_PrintUsageIndividual 
    #        (CompanySeq,NDate,UserID,DeptName,UserName,MonoPrint,ColorPrint,MonoCopy,ColorCopy,DeviceID,InsertTime) VALUES (?,?,?,?,?,?,?,?,?,?,?)'''
    query = '''INSERT INTO emsf_PrintUsageIndividual 
            (CompanySeq,NDate,UserID,DeptName,UserName,MonoPrint,ColorPrint,MonoCopy,ColorCopy,DeviceID,InsertTime) VALUES (%d,%d,%d,%s,%s,%d,%d,%d,%d,%s,%s)'''
    cursor.execute(query,(1,NDate,UserID,DeptName,UserName,MonoPrint,ColorPrint,MonoCopy,ColorCopy,DeviceID,InsertTime))
    conn.commit()



def dbInsertTotal(NDate,MonoTotal,ColorTotal,DeviceID,InsertTime):
    query = "INSERT INTO emsf_PrintUsageTotal (CompanySeq,NDate,MonoTotal,ColorTotal,DeviceID,InsertTime) VALUES (%d,%s,%d,%d,%s,%s)" #PC용
    #query = "INSERT INTO emsf_PrintUsageTotal (CompanySeq,NDate,MonoTotal,ColorTotal,DeviceID,InsertTime) VALUES (?,?,?,?,?,?)" #서버용
    cursor.execute(query,(1,NDate,MonoTotal,ColorTotal,DeviceID,InsertTime))
    conn.commit()

def dbExcute(excuteName):
    result = cursor.execute(excuteName)
    return result 
    
    
def dbFetchOne():
    result = cursor.fetchone()
    return result
    
def dbFetchAll():
    result = cursor.fetchall()
    return result
    
def dbClose():
    conn.close()