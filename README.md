# ClassicASP-Helper
------------------------
First Classic ASP Coding Helper Utility

## Usage

İlk olarak dosyayı include ediyoruz.
<!--#include file="casphelper.asp"-->

Eğer kendiniz kütüphaneyi başlatmak isterseniz bu kodu ekledin
	Set Query = New QueryManager

Artık tüm işlemleriniz için Query değişkenini kullanmanız yeterlidir.

## SQL Insert İşlemi

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


