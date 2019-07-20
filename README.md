# ClassicASP-Helper
------------------------
First Classic ASP Coding Helper Utility

## Usage

İlk olarak dosyayı include ediyoruz.
<!--#include file="casphelper.asp"-->

Eğer kendiniz kütüphaneyi başlatmak isterseniz bu kodu ekledin
	Set Query = New QueryManager

Artık tüm işlemleriniz için Query değişkenini kullanmanız yeterlidir.

	Dim Query
	Set Query = New QueryManager
	Query.Debug          = False
	Query.Host           = "localhost"
	Query.Database       = "my_db_name"
	Query.User           = "my_db_username"
	Query.Password       = "MyS3c3tP4ssw0d"
	Query.Connect()


## SQL Insert/Update İşlemi

Bir SQL sorgusunu INSERT yapmak istersek, form input name değerlerimizi, ilgili tablonun sütun isimleriyle aynı tutmamız gerekiyor. Kütüphane burada bir kaç işlem yapar.
* Gelen FORM/POST name parametreleri ve hedef tablo sütun isimleri eşleşiyor mu? Eşleşmiyor ise, alan dışında kalan veriler alınmaz.
* Gelen veriler, hedef sütun veri türü ile uyuşuyor mu? (INT, VARCHAR/LONGTEXT, DATE/DATETIME)
* Gelen veriler boşmu?

	Query.Debug = False
	Query.CollectForm("INSERT")
	Query.AppendRows    = "EKLENME_TARIHI"
	Query.AppendValues  = "'"& Now() &"'"
	Query.Run("INSERT INTO tbl_tableName("& Query.Rows &") VALUES("& Query.Values &")")
	Query.Go("?Cmd=OtelOdalari&ID="& Query.MaxID("tbl_tableName") &"&Msg=Success")


## SQL - MaxID İşlemi

	Query.MaxID("tbl_tableName")
	Query.MaxID("tbl_tableName WHERE ID = 1")

## SQL Parametrik Güncellemeler

Eğer bir form yada querystring verisi almak istiyorsanız Query.Data("anahtar") yada inline olarak {anahtar} şeklinde alabilirsiniz. Yazılımınız 404 url yapısında bile olsa tüm parametreleri yakalayacaktır.

	Query.Run("SELECT ID FROM tbl_tableName WHERE ID = {ID} ")
	Query.Run("SELECT ID FROM tbl_tableName WHERE ID = "& .Data("ID") &" ")
	Query.Run("SELECT ID FROM tbl_tableName WHERE ID = 1 ")

## Response.Redirect

İşleminizi tamamladıktan sonra kullandığınız Response.Redirect "url.asp?some=string" yerine kullanabileceğiniz bir komuttur. Amaç, Set edilen objelerin komple kapatılabilmesi içindir. Açtığınız DB bağlantıları, kapatılmamış Set tanımları gibi tüm açık objeleri kapatarak redirect işlemi yapabilirsiniz. Ayrıca güncel verilerinizi içine işleyebilirsiniz. Form yada Querystring parametresi çekmek için Parametrik Güncellemeler kullanılabilir.

	Query.Go("url.asp?some=string")
	Query.Go("url.asp?some={ID}")



