<%
    ' Kayıt Bilgileri Alınıyor
    id = Request.QueryString("id")
%>

<%
    databaseAdi = "veri.mdb"
    databaseYolu = "/db/"
    databaseTamYol = Server.MapPath(databaseYolu & databaseAdi )
    
    'Sağlayıcı
    connectionProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"          
    connectionString = connectionProvider & "Data Source=" & databaseTamYol
    
    'Bağlan Database
    Set conn=Server.CreateObject("ADODB.Connection")
%>

<%
    ' İşlemlere Başla

    ' İlk Defa Bağlanıyr isem bağlantımı açmalıyım
    conn.Open connectionString
    
    ' Bağlantı Kayıt Nesnesi
    set rs =  Server.CreateObject("ADODB.recordset")    
            
    ' Sorgu
    sorgu = "DELETE * FROM Kisi Where id=" & id

    ' Çalıştır
    rs.Open sorgu, conn 'Bağlantı Açıldı        
    
    ' Kayıt Sil
    
    ' Tüm işlemler tamamlandı    

    ' Bağlantımı Kapat
    conn.Close

    ' Ana Sayfaya Geri Dönüş
    Response.Redirect("/")
%>