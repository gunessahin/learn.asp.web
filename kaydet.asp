<%@LANGUAGE="VBSCRIPT" CODEPAGE="1254"%>
<%
Veri_yolu=Server.MapPath("Veritabani.mdb")
Bcumle="DRIVER={Microsoft Access Driver(*.mdb)};DBQ=" & Veri_yolu
Set bag=Server.CreateObject("ADODB.Connection")
bag.Open(Bcumle)
Set kset=Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM ögrenci"
kset.open sql, bag, 1, 3 

ad=Request.Form("ad")
soyad=Request.Form("soyad")

if ad="" or soyad=""
Response.Write "<script language='JavaScript'>alert('Bilgileri eksiksiz doldurunuz...'):history.back(-1);</script>"
Response.End
end if

kset.AddNew
kset("ad")=ad
kset("soyad")=soyad

kset.Update
kset.close
set kset=nothing
bag.close
bag=nothing

Response.Redirect("default.asp")
%>
