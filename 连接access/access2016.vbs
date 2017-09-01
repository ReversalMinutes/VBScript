Set conn=CreateObject("adodb.connection")
dim connStr
connStr = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=D:\´úÂë\VBS\Á¬½Óaccess\Database1.mdb;UID=admin"
conn.Open connStr
conn.Close
