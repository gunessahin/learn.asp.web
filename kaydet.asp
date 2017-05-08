<%
    ' Kayıt Bilgileri Alınıyor
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
            
    ' Kayıt
    sorgu = "SELECT * FROM Kisi"

    ' Çalıştır
    rs.Open sorgu, conn ,1  , 3'Bağlantı Açıldı        
    
    ' Yeni Kayıt
    rs.AddNew

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