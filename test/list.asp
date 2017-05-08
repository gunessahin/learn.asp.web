<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Asp Test Page</title>
</head>
<body>
    <%=Response.Write("Connection Test: <b> List </b>") %>
    <hr />
    <%
        ' PROVIDERS

        ' Microsoft ACE OLEDB 12.0
        ' ==================================================

        ' Standart
        ' Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\myFolder\myAccessFile.accdb;Persist Security Info=False;

        ' With database password
        ' Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\myFolder\myAccessFile.accdb;Jet OLEDB:Database Password=MyDbPassword;

        ' --------------------------------------------------

        ' Microsoft ACE OLEDB 4.0
        ' ==================================================
        
        ' Standart Access 2003
        ' Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;User Id=admin;Password=;
        
        ' With database password Access 2003
        ' Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;Jet OLEDB:Database Password=MyDbPassword;

        ' --------------------------------------------------

        ' Microsoft Access accdb ODBC Driver connection strings
        ' ==================================================

        ' Standart Access 97 / Access 2000 / Access 2002 / Access 2002 / Access 2007
        ' Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\mydatabase.mdb;       
  

        ' Define
        db = "db97.mdb"
        dbPath = "/test/"
        dbq = Server.MapPath(dbPath & db)

        ' Test 1
        connectionProvider = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};"        
        connectionString = connectionProvider & "DBQ=" & dbq

        ' Test 2
        connectionProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"          
        connectionString = connectionProvider & "Data Source=" & dbq

        ' Connection
        set conn=Server.CreateObject("ADODB.Connection")        
        Response.Write("<b>Connection</b> :" & connectionString & "</br>")
    %>

    <%
    ' Proccess
    conn.Open connectionString
    Response.Write("Connection Open </br>")
    
    ' RecordSet
    set rs = Server.CreateObject("ADODB.recordset")
    query = "SELECT * FROM Kisi"
    rs.Open query, conn
    Response.Write("Record Open </br>")

    Response.Write("==================================================" & "</br>")
    
    ' Get
    do until rs.EOF                     ' Check
    
        Response.Write("Get Data </br>")

        ' List
        for each x in rs.Fields
            Response.Write(x.name)
            Response.Write(" - ")
            Response.Write(x.value & "<br>")
        next
    
        Response.Write("<br>")
    
        rs.MoveNext
    
    loop

    Response.Write("==================================================" & "</br>")
    
    Response.Write("List Complated </br>")
    
    ' Close
    rs.close
    conn.close

    Response.Write("Close Recordset and Connection </br>")
    %>
</body>
</html>
