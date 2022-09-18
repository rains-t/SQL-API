import pyodbc
from indexing import index
from collector import data_collector

connection = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost\SQLEXPRESS;Database=master;Trusted_Connection=yes;')
cursor = connection.cursor()
pages = []
    
data = data_collector('04_2020.xlsx', (f'{i}' for i in range(1, 31)))
#(f'{i}' for i in range(1, 31) make iteration number 1 more than page needed
internal_timer = 0
item_counter = 0
for group in data:
    item_counter += 1
    
    
    for i in range(0, len(group)):
            
        if i == 0:
            date = group[i]
        if i == 1:
            time = group[i]
        if i == 2:
            Tagindex = index(group[i])
           
        if i == 3:
            val = group[i]
            date_time = f'{date} {time}'
            print(f'{date_time}, {Tagindex}, {val} wrote')

            

            cursor.execute('''INSERT INTO TestDB.dbo.FloatTable (DateAndTime, TagIndex, val) VALUES (?, ?, ?)''', (date_time, Tagindex, val))
                
            internal_timer += 1    
                


            if internal_timer == 25:
                connection.commit()
                internal_timer = 0  

connection.commit()  
            

            
print(item_counter,'items wrote')
