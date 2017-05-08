<%id=Request.QueryString("id")

Veri_yolu=Server.MapPath("Veritabani.mdb")
Bcumle="DRIVER={Microsoft Access Driver(*.mdb)};DBQ=" & Veri_yolu
Set bag=Server.CreateObject("ADODB.Connection")
bag.Open(Bcumle)
Set kset=bag.Execute("SELECT * FROM ogrenci WHERE id=" & id)
if kset.Eof then
response.write = "Olmayan kayıt istendi" 
%>
<a href="default.asp">ÖĞRENCİ LİSTESİ</a>

<form action="kayit_guncelle.asp?id=<%=kset("id")%>" method="post">
<table width="750px">
<tr>
<td width="250px">AD:</td>
<td width="500px"><input type="text" name="ad" size="50" style="border:1px" value="<%=kset("ad")%>"></td>
</tr>
<tr>
<td>SOYAD:</td>
<td><input type="text" name="soyad" size="50" style="border:1px" value="<%=kset("soyad")%>"></td>
</tr>
<tr>
<td>
  <input type="submit" name="button" id="button" value="DEĞİŞTİR"></td>
<td>&nbsp;</td>
</tr>
</table>
</form>
<%
kset.Close
Set kset=Nothing
bag.Close
Set bag=Nothing

%>
