


import pyodbc
connection = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost\SQLEXPRESS;Database=master;Trusted_Connection=yes;')
cursor = connection.cursor()
def index(tagname):
    tagname = (f'%{tagname}%')
    # print(tagname)
    
    '''uses SQL code to execute a LIKE command, feed data containing keyword name to 
    return a SQL index position. '''
    select_query = '''SELECT TagIndex FROM TestDB.dbo.TagTable WHERE
     TagName LIKE (?)  '''
    cursor.execute(select_query, tagname)
    tagindex = cursor.fetchone()[0]
    connection.commit()
    return tagindex












