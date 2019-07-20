# ClassicASP-Helper
(TR) Bilinen ilk kompakt, Klasik ASP yardımcı kütüphanesidir (araştırmalarıma göre). Sıklıkla yaptığınız işlemleri kısaltan, özellikle veritabanı çalışmalarınızda ve yazılım geliştirme aşamalarında pratiklik ile hız kazanmanızı, geliştirmelerinizi daha kolay yapmanızı sağlayacak yapıdadır. Mevcut kütüphanelerinize entegre edebilir, geliştirebilir ve dağıtabilirsiniz. Lütfen Star vermeyi, Watch listenize eklemeyi unutmayın.

(EN) First Classic ASP Coding Helper Utility

# Usage / Kullanım
(TR) İlk olarak dosyayı fiziksel yolundan proje dosyanıza include edin.
(EN) ...

<!--#include file="/{path}/casphelper.asp"-->

(TR) Eğer kendiniz kütüphaneyi başlatmak isterseniz aşağıda ki kodu ilk sırada çalışacak şekilde projenize ekleyin
(EN) ...

	Set Query = New QueryManager

(TR) Artık tüm işlemleriniz için *Query* değişkenini kullanmanız yeterlidir.
(EN) ...

	Dim Query
	Set Query = New QueryManager
	  Query.Debug          = False
	  Query.Host           = "localhost"
	  Query.Database       = "my_db_name"
	  Query.User           = "my_db_username"
	  Query.Password       = "MyS3c3tP4ssw0d"
	  Query.Connect()

## SQL Insert/Update İşlemi
### fn: RunExtend

Bir SQL sorgusunu INSERT veya UPDATE yapmak istersek, form input name değerlerimizi, ilgili tablonun sütun isimleriyle aynı tutmamız gerekiyor. Kütüphane burada bir kaç işlem yapar.
* Gelen FORM/POST name parametreleri ve hedef tablo sütun isimleri eşleşiyor mu? Eşleşmiyor ise, alan dışında kalan veriler alınmaz.
* Gelen veriler, hedef sütun veri türü ile uyuşuyor mu? (INT, VARCHAR/LONGTEXT, DATE/DATETIME)
* Gelen veriler boşmu?

Sonuç olarak kütüphaneden 2 türde yanıt döner. 
* INSERT işlemi başarılı ise, ID parametresi ile son eklenen kayıt numarası (INT) döner
* UPDATE işlemi başarılı ise true, başarısız ise false değeri (BOOLEAN) döner.

- INSERT İşlemi
	Query.RunExtend("INSERT", "table_name", "")

-UPDATE İşlemi
	Query.RunExtend("UPDATE", "table_name", "ID={ID}")

> tbl_users
	ID(INT) Primary Key
	NAME(VARCHAR)
	SURNAME(VARCHAR)
	BIRTHDAY(DATE)

> POST.ASP
	<form action="/?Cmd=InsertSample" method="post">
		<input name="NAME" value="" />
		<input name="SURNAME" value="" />
		<input name="BIRTHDAY" value="" />
		<button type="submit">Submit</button>
	</form>

> CATCH.ASP
<%
If Query.Data("Cmd") = "InsertSample" Then 
	If Query.RunExtend("INSERT", "tbl_users", "") = True Then
		Response.Write "Başarılı / Success"
	Else 
		Response.Write "Başarısız / Failed"
	End If
End If
%>


### CollectForm & Run
Kütüphanenin ilk versiyonunda bulunan Collector ve Run komutlarının birleşimi aşağıda ki gibidir. CollectForm fonksiyonu, FORM Post methodu ile gelen Request.Form parametrelerini toplar ve INSERT yada UPDATE için birleştirir.
	Query.Debug = False
	Query.CollectForm("INSERT")
	Query.AppendRows    = "EKSTRA1, EKSTRA2"
	Query.AppendValues  = "'Manuel Eklenecek Veri 1', 'Manuel Eklenecek Veri 2'"
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



