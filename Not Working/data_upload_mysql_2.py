# Python code to illustrate and create a 
# table in database 
import mysql.connector as mysql 

host = 'nadunjayathunga.mysql.pythonanywhere-services.com'
user = 'nadunjayathunga'
passowrd = '#a+n=*2023Happy*'
database = 'nadunjayathunga$elite_security'

# Open database connection 
db = mysql.connect(host=host,user=user,password=passowrd,database=database) 

cursor = db.cursor() 

# Drop table if it already exist using execute() 
cursor.execute("DROP TABLE IF EXISTS EMPLOYEE") 

# Create table as per requirement 
sql = "CREATE TABLE EMPLOYEE ( FNAME CHAR(20) NOT NULL, LNAME CHAR(20), AGE INT )"

cursor.execute(sql) #table created 

# disconnect from server 
db.close() 
