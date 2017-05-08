<%
    ' Kayıt Bilgileri Alınıyor
    id = Request.Form("id")
    ad = Request.Form("isim")
    soyad = Request.Form("soyad")    
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
    sorgu = "SELECT * FROM Kisi Where id=" & id

    ' Çalıştır
    rs.Open sorgu, conn ,1  , 3'Bağlantı Açıldı        
    
    ' Kayıt Güncelle    

    ' Veri Alanlarını Tanımla
    rs("Ad") = ad
    rs("Soyad") = soyad

    ' Veriyi Tabloya aktar
    rs.Update
    
    ' Tüm işlemler tamamlandı
    
    ' Kayıt Setimi Kapat
    rs.Close

    ' Bağlantımı Kapat
    conn.Close

    ' Ana Sayfaya Geri Dönüş
    Response.Redirect("/")
%>