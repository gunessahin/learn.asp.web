<%@LANGUAGE="VBSCRIPT" CODEPAGE="1254"%>
<%
id=Request.Form("id")
Veri_yolu=Server.MapPath("db/veri_tabani.mdb")
Bcumle="DRIVER={Microsoft Access Driver(*.mdb)};DBQ=" & Veri_yolu
Set bag=Server.CreateObject("ADODB.Connection")
bag.Open(Bcumle)
Set kset=Server.CreateObject("ADODB.Recordset")
sql="DELETE * FROM ögrenci WHERE id=" & id
Set kset=bag.execute(sql
set kset=Nothing
bag.close
set bag=nothing
%>