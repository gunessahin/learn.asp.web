<%@LANGUAGE="VBSCRIPT" CODEPAGE="1254"%>

<%Veri_yolu=Server.MapPath("Veritabani.mdb")
Bcumle="DRIVER={Microsoft Access Driver(*.mdb)};DBQ=" & Veri_yolu
Set bag=Server.CreateObject("ADODB.Connection")
bag.Open(Bcumle)
Set kset=bag.Execute("SELECT * FROM ogrenci") %>

<% i=1%>
<p><a href="kayit_yeni.asp">YENİ KAYIT</a></p>
<table border=1>
<tr>
<th>#</th>
<th>AD</th>
<th>SOYAD</th>
<th>İŞLEMLER</th>
</tr>
<% Do While Not kset.EOF%>
<tr>
<td><%=i%></td>
<td><%=kset("ad")%></td>
<td><%=kset("soyad")%></td>
<td><a href="kayit_duzenle.asp?id=<%=kset("id")%>">DÜZENLE</a> | <a href="kayit_sil.asp?id=<%=kset("id")%>">SİL</a></td>
</tr>
<%kset.Movenext%>
<%i=i + 1%>
<%Loop%>
</table>


<%kset.Close
Set kset=Nothing
bag.Close
Set bag=Nothing%>
