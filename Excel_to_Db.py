import xlrd
import MySQLdb

# Open the workbook and define the worksheet
book = xlrd.open_workbook("beginner_assignment.xls")
# To access the sheet we can either use its path or index Uncomment to use index
#sheet = book.sheet_by_index()
sheet1 = book.sheet_by_name("product_listing")
sheet2 = book.sheet_by_name("group_listing")

# Establish a MySQL connection, Enter your full hostname, username&password
database = MySQLdb.connect (host="localhost", user = "root", passwd = "", db = "mysqlPython")

# Get the cursor, which is used to traverse the database, line by line so you can perform a query
cursor = database.cursor()

# Create the INSERT INTO sql query
query1 = """INSERT INTO Product_listing (Product_Name, Model_Name, Product_Serial_No, Group_Name, Product_MRP) VALUES (%s, %s, %s, %s, %s)"""
query2 ="""INSERT INTO Group_listing (Group_name, Group_description,isActive) VALUES(%s, %s, %s)"""
# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
for r in range(1, sheet1.nrows):
		Product_Name	  = sheet1.cell(r,0).value
		Model_Name	      = sheet1.cell(r,1).value
		Product_Serial_No = sheet1.cell(r,2).value
		Group_Name        = sheet1.cell(r,3).value
		Product_MRP		  = sheet1.cell(r,4).value

		# Assign values from each row
		values1 = (Product_Name, Model_Name, Product_Serial_No, Group_name, Product_MRP)

		# Execute sql Query
		cursor.execute(query1, values1)
#Do the same for sheet2
for r in range(1, sheet2.nrows):
	    Group_name         = sheet2.cell(r,0).value
		Group_description   = sheet2.cell(r,1).value
		isActive           = sheet2.cell(r,2).value

		values2 = (Group_name, Group_description, isActive)

		cursor.execute(query2, values2)
# Close the cursor
cursor.close()

# Commit the transaction
database.commit()

# Close the database connection
database.close()
