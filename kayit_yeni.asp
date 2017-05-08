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
        <h3>Yeni Kayıt</h3>
        <form class="form" action="kaydet.asp" method="post">
            <input type="hidden" class="form-control" id="id" name="id">
            <div class="form-group">
                <label>İsim</label>
                <input type="text" class="form-control" id="isim" name="isim" placeholder="isim" maxlength="50" required>
            </div>
            <div class="form-group">
                <label>Soyisim</label>
                <input type="text" class="form-control" id="soyad" name="soyad" placeholder="soyad" maxlength="50" required>
            </div>
            <button type="submit" class="btn btn-default">Kaydet</button>
        </form>
    </div>
</body>
</html>
