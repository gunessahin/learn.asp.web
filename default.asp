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
        
    ' Listeleme 
    sorgu = "SELECT * FROM Kisi"

    ' Çalıştır
    rs.Open sorgu, conn 'Bağlantı Açıldı
    %>
    <div class="container">
        <h2>Kişiler</h2>
        <table class="table">

            <thead>
                <tr>
                    <th>#</th>
                    <th>Ad</th>
                    <th>Soyad</th>
                    <th>İşlemler</th>
                </tr>
            </thead>

            <tbody>

                <%
    ' Kontrol Tablo Kayıtları 
    i = 0
    do until rs.EOF
                %>

                <%    
    ' Bilgileri Tanımla    
    i = i + 1
    id = rs("id")
    ad = rs("Ad")
    soyad = rs("Soyad")

    ' Bilgileri Yazdır
                %>
                <tr>
                    <td><%=i %></td>
                    <td><%=ad%></td>
                    <td><%=soyad%></td>
                    <td>
                        <a href="kayit_sil.asp?id=<%=id%>">Sil</a> | <a href="kayit_duzenle.asp?id=<%=id%>">Düzelt</a>
                    </td>
                </tr>
                <%
    ' Sonraki Satıra Geçtiğimiz Alan
    rs.MoveNext
                %>

                <%
    loop
    ' Listeleme Tamalandı
                %>
            </tbody>
        </table>
    </div>

    <%
    ' Tüm işlemler tamamlandı

    ' Kayıt Setimi Kapat
    rs.Close

    ' Bağlantımı Kapat
    conn.Close
    %></body>
</html>
