<%@LANGUAGE="VBSCRIPT" CODEPAGE="1254"%>
<%
id=Request.Form("id")
ad=Request.Form("ad")
soyad=Request.Form("soyad")

if ad="" or soyad=""
Response.Write "<script language='JavaScript'>alert('Bilgileri eksiksiz doldurunuz...'):history.back(-1);</script>"
Response.End
end if

Veri_yolu=Server.MapPath("db/veri_tabani.mdb")
Bcumle="DRIVER={Microsoft Access Driver(*.mdb)};DBQ=" & Veri_yolu
Set bag=Server.CreateObject("ADODB.Connection")
bag.Open(Bcumle)
Set kset=Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM ögrenci WHERE id=" & id
kset.open sql, bag, 1, 3 


kset("ad")=ad
kset("soyad")=soyad
kset.Update

kset.close
set kset=nothing
bag.close
bag=nothing

Response.Redirect("default.asp")
%>