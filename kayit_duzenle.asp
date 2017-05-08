<%
    ' Kayıt Bilgileri Alınıyor
    id = Request.QueryString("id")
    ad = ""
    soyad = ""
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Veritabanı İşlemleri</title>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" />
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" />
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
</head>

<body>
    <nav class="navbar">
        <div class="container">
            <div class="navbar-header">
                <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar">
                    <span class="sr-only">Toggle navigation</span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
            </div>
            <div id="navbar" class="collapse navbar-collapse">
                <ul class="nav navbar-nav">
                    <li class="active"><a href="/">Listeleme</a></li>
                    <li><a href="kayit_yeni.asp">Yeni</a></li>
                </ul>
            </div>
            <!--/.nav-collapse -->
        </div>
    </nav>
    <div class="container">
        <h3>Yeni Duzenle</h3>
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
    sorgu = "SELECT * FROM Kisi Where id=" & id

    ' Çalıştır
    rs.Open sorgu, conn ,1  , 3'Bağlantı Açıldı        
    
    ' Kayıt Bilgileri Al    

    ' Veri Alanlarını Tanımla
    ad = rs("Ad")
    soyad = rs("Soyad") 

    
    ' Tüm işlemler tamamlandı
    
    ' Kayıt Setimi Kapat
    rs.Close

    ' Bağlantımı Kapat
    conn.Close

        %>

        <form class="form" action="kayit_guncelle.asp" method="post">
            <div class="form-group">
                <label>id</label>
                <div class="form-control"><%=id %></div>
                <input type="hidden" class="form-control" id="id" name="id" placeholder="<%=id %>" value="<%=id %>">
            </div>
            <div class="form-group">
                <label>İsim</label>
                <input type="text" class="form-control" id="isim" name="isim" placeholder="isim" value="<%=ad %>" maxlength="50">
            </div>
            <div class="form-group">
                <label>Soyisim</label>
                <input type="text" class="form-control" id="soyad" name="soyad" placeholder="soyad" value="<%=soyad %>" maxlength="50">
            </div>
            <button type="submit" class="btn btn-default">Düzelt</button>
        </form>
    </div>
</body>
</html>
