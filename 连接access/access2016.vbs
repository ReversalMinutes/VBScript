Set conn=CreateObject("adodb.connection")
dim connStr
connStr = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=D:\����\VBS\����access\Database1.mdb;UID=admin"
conn.Open connStr
conn.Close
