# ClassicASP-Helper
----------------

(TR) Bilinen ilk kompakt, Klasik ASP yardımcı kütüphanesidir (araştırmalarıma göre). Sıklıkla yaptığınız işlemleri kısaltan, özellikle veritabanı çalışmalarınızda ve yazılım geliştirme aşamalarında pratiklik ile hız kazanmanızı, geliştirmelerinizi daha kolay yapmanızı sağlayacak yapıdadır. Mevcut kütüphanelerinize entegre edebilir, geliştirebilir ve dağıtabilirsiniz. Lütfen Star vermeyi, Watch listenize eklemeyi unutmayın.

(EN) First Classic ASP Coding Helper Utility

# Usage / Kullanım
----------------

> (TR) İlk olarak dosyayı fiziksel yolundan proje dosyanıza include edin.
> (EN) ...

```vb
<!--#include file="/{path}/casphelper.asp"-->
```

> (TR) Eğer kendiniz kütüphaneyi başlatmak isterseniz aşağıda ki kodu ilk sırada çalışacak şekilde projenize ekleyin
> (EN) ...

```vb
<%
Set Query = New QueryManager
%>
```

> (TR) Artık tüm işlemleriniz için *Query* değişkenini kullanmanız yeterlidir.
> (EN) ...

```vb
<%
Dim Query
Set Query = New QueryManager
  Query.Debug          = False
  Query.Host           = "localhost"
  Query.Database       = "my_db_name"
  Query.User           = "my_db_username"
  Query.Password       = "MyS3c3tP4ssw0d"
  Query.Connect()
%>
```

# SQL
## Insert/Update
> Bir SQL sorgusunu INSERT veya UPDATE yapmak istersek, form input name değerlerimizi, ilgili tablonun sütun isimleriyle aynı tutmamız gerekiyor. Kütüphane burada bir kaç işlem yapar.
* Gelen FORM/POST name parametreleri ve hedef tablo sütun isimleri eşleşiyor mu? Eşleşmiyor ise, alan dışında kalan veriler alınmaz.
* Gelen veriler, hedef sütun veri türü ile uyuşuyor mu? (INT, VARCHAR/LONGTEXT, DATE/DATETIME)
* Gelen veriler boşmu?

> Sonuç olarak kütüphaneden 2 türde yanıt döner. 
* INSERT işlemi başarılı ise, ID parametresi ile son eklenen kayıt numarası (INT) döner
* UPDATE işlemi başarılı ise true, başarısız ise false değeri (BOOLEAN) döner.

### - RunExtend(INSERT)
> (TR) RunExtend fonksiyonu basit bir return fonksiyonudur ve INSERT parametresi işlem sonucunda eklenen satırın primaryKey (ID) değerini (INT) döndürür. Bu sonuç, INSERT işleminin başarılı olup olmadığı bilgisini verir.
> (EN) ...

```vb
Query.RunExtend("INSERT", "table_name", Null)
```

> (TR) Örnek kullanım için aşağıdaki yapı kullanılabilir.
> (EN) ...

```vb 
<%
If Query.Data("Cmd") = "InsertSample" Then 
    Dim QueryResult
    QueryResult = Query.RunExtend("INSERT", "tbl_users", Null)
    If IsNumeric( QueryResult ) Then
        Response.Write "Başarılı / Success"
        Response.Write "ID: " & QueryResult
    Else 
        Response.Write "Başarısız / Failed"
    End If
End If
%>
```

> (TR) Form yapısı şu şekilde olmalıdır.
> (EN) ...

```html
<form action="/?Cmd=InsertSample" method="post">
    <input name="NAME" value="Anthony Burak" />
    <input name="SURNAME" value="Dursun" />
    <input name="BIRTHDAY" value="24.07.1986" />
    <button type="submit">Insert</button>
</form>
```

> (TR) Veritabanı yapısı ise aşağıdaki gibidir
> (EN) ...

| FIELD NAME | TYPE |
| ------ | ------ |
| ID | (INT) Primary Key
| NAME | (VARCHAR)
| SURNAME | (VARCHAR)
| BIRTHDAY | (DATE)

### - RunExtend(UPDATE)
> (TR) RunExtend fonksiyonu basit bir return fonksiyonudur ve UPDATE parametresi işlem sonucunda *true* veya *false* (boolean) dönüş yapar. Bu sonuç, UPDATE işleminin başarılı olup olmadığı bilgisini verir
> (EN) ...

```vb
Query.RunExtend("UPDATE", "table_name", "ID={ID}")
```

> (TR) Örnek kullanım için aşağıdaki yapı kullanılabilir.
> (EN) ...

```vb
<%
If Query.Data("Cmd") = "UpdateSample" Then 
    Dim QueryResult
    QueryResult = Query.RunExtend("UPDATE", "tbl_uyeler", "ID={ID}")
    If QueryResult = True Then 
        Response.Write "Başarılı / Success"
    Else 
        Response.Write "Başarısız / Failed"
    End If
End If
%>
```

> (TR) Form yapısı şu şekilde olmalıdır.
> (EN) ...

```html
<form action="/?Cmd=UpdateSample&ID=123" method="post">
    <input name="NAME" value="Anthony Burak" />
    <input name="SURNAME" value="Dursun" />
    <input name="BIRTHDAY" value="24.07.1986" />
    <button type="submit">Update</button>
</form>
```

> (TR) Veritabanı yapısı ise aşağıdaki gibidir
> (EN) ...

| FIELD NAME | TYPE |
| ------ | ------ |
| ID | (INT) Primary Key
| NAME | (VARCHAR)
| SURNAME | (VARCHAR)
| BIRTHDAY | (DATE)

### -CollectForm & Run (deprecate)
> (TR) Kütüphanenin ilk versiyonunda bulunan Collector ve Run komutlarının birleşimi aşağıda ki gibidir. CollectForm fonksiyonu, FORM Post methodu ile gelen Request.Form parametrelerini toplar ve INSERT yada UPDATE için birleştirir. Herhangi bir kontrol mekanizması yoktur. Parametre hatası Error Raise döner.
> (EN) ...

```vb
<%
If Query.Data("Cmd") = "UpdateSample" Then 
	Query.CollectForm("INSERT")
	Query.AppendRows    = "EKSTRA1, EKSTRA2"
	Query.AppendValues  = "'Manuel Eklenecek Veri 1', 'Manuel Eklenecek Veri 2'"
	Query.Run("INSERT INTO tbl_tableName("& Query.Rows &") VALUES("& Query.Values &")")
	Query.Go("?Msg=Success")
End If
%>
```

## RecordExist(sql)
> (TR) Bir SQL sorgusunun sonucunu *true* veya *false* olarak döner. Geleneksel yöntemlerde EOF muadili olarak kullanılır.
> (EN) ...

```vb
<%
Dim QueryResult
QueryResult = Query.RecordExist("SELECT ID FROM tbl_users WHERE ID = 1")
If QueryResult = True Then
    Response.Write "Record Exist"
Else
    Response.Write "Record Not Exist"
End If
%>
```

## MaxID
> (TR) Herhangi bir tabloda ve koşulda maksimum ID (PrimaryKey) değerinin döndürülmesini sağlar. Hata kontrolü yoktur.
> (EN) ...

```vb
Query.MaxID("tbl_tableName")
```

> (TR) Koşullu durumlar için
> (EN) ...

```vb
Query.MaxID("tbl_tableName WHERE EMAIL = 'badursun@gmail.com'")
```

## Execute SQL [ conn.Execute(sqlQuery) ]
> (TR) Bu fonksiyon için bulunan tek özelleştirme Request.Querystring ile alınacak verinin *Replace* edilebilir olmaısıdır. URL yapısı /?Cmd=Update&ID=123 olarak geliyorsa, sorgu içinde *{ID}* parametresi *123* olarak güncellenir. . Standart obj.Execute(sql) parametresini yerine getirir. 
> (EN) ...

```vb
<%
Query.Run("SELECT ID FROM tbl_tableName WHERE ID = {ID} ")
Query.Run("SELECT ID FROM tbl_tableName WHERE ID = "& Query.Data("ID") &" ")
Query.Run("SELECT ID FROM tbl_tableName WHERE ID = 1 ")
%>
```

-----------------

# Request ve Response
## - Request.Form/QueryString
> (TR) Eğer bir form yada querystring verisi almak istiyorsanız *Query.Data("anahtar")* yada inline olarak *{anahtar}* şeklinde alabilirsiniz. Yazılımınız 404 url yapısında bile olsa tüm parametreleri yakalayacaktır. Requet.Form(anahtar) veya Request.QueryString(anahtar) yerine kullanılabilir.
> (EN) ...

```vb
<%
Dim SampleValue
SampleValue = Query.Data("ID")
%>
```
> (TR) Verinin varlığı bulunmadıysa (Null, Empty) sonuç her zaman Empty ile karşılanabilir.
> (EN) ...

```vb
/script.asp?Cmd=Test&Data1=value&Data2=&Data3=value3
/404url/params/?Cmd=Test&Data1=value&Data2=&Data3=value3

<%
Response.Write Query.Data("Cmd")    ' return Test (String)
Response.Write Query.Data("Data1")  ' return value (String)
Response.Write Query.Data("Data2")  ' return 
Response.Write Query.Data("Data3")  ' return value3 (String)
%>
```
## - Response.Redirect
> (TR) İşleminizi tamamladıktan sonra kullandığınız Response.Redirect "url.asp?some=string" yerine kullanabileceğiniz bir komuttur.Güncel Request verilerini işleyebilirsiniz. Form yada Querystring parametresi çekmek için Parametrik Güncellemeler kullanılabilir.
> (EN) ...

```vb
<%
Query.Go("url.asp?some=string")
Query.Go("url.asp?some={ID}")
%>
```
## - Response.Write
> (TR) Standart *Response.Write("test")* kullanımı yerine *Query.Echo("test")* kullanılabilir.
> (EN) ...
```vb
<%
Query.Echo("test")
%>
```

## - Response.End
> (TR) Standart *Response.End()* kullanımı yerine *Query.Kill()* kullanılabilir.
> (EN) ...
```vb
<%
Query.Kill()
%>
```
---------------
# Helpers / Yardımcılar
## - Exist()
> (TR) Herhangi bir değişken için varlık kontorlü yapabilir. IsNull, IsEmpty, Len()>0 kontrolleri gerçekleştirir ve *true* yada *false* (boolean) sonuç döndürür
> (EN) ...
```vb
<%
str_value1 = ""
str_value2 = 2
If Query.Exist(str_value1) = True Then
    ' return true
End If

If Query.Exist(str_value2) = False Then
    ' return false
End If
%>
```

## - FindInArray(String, Array)
> (TR) Herhangi bir string veriyi, bir array öbeği içerisinde arar. Tam eşleşme kontrolü yapar, otomatik Trim() uygular. Sonuç bulunursa, index numarası döner. Sonuç bulunamazsa *Null* sonuç döner.
> (EN) ...

```vb
<%
Dim str_array
str_array = Array("test", "apple", "fruit", "banana", "mercedes")

Dim QueryResult
QueryResult = Query.FindInArray("apple", str_array)
If IsNull( QueryResult ) Then
    Query.Echo "Not Found"
Else
    Query.Echo "apple found in array index: " & QueryResult
End If
%>
```

## - AllowedMethod(MethodName)
> (TR) Bazı durumlarda ilgili işlem alanına sadece belirli method'lar ile erişilmesini sınırlayabilirsiniz. Örnek olarak bir form için Method="POST" kullanırsanız, karşılamada method'un gerçekten "POST" olduğunu teyit edebilirsiniz.
> (EN) ...

```vb
<%
If Query.AllowedMethod("POST") = False Then
    Query.Echo "Only POST Method Allowed"
    Query.Kill
End If
%>
```




