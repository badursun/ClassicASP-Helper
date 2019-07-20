<%
Set Query = New QueryManager

'####################################
' Custom SQL Query Helper
' Classic ASP & MySQL Query Builder
' 2019  (c) Anthony Burak DURSUN
'####################################

' HELPER USAGE 												| TYPE 						| RETURN 					| DETAILS
'---------------------------------------------------------------------------------------------------------
' SIMPLE FUNCTIONS
'---------------------------------------------------------------------------------------------------------
' .BadWord( vString ) 								| Function 				| string					| BadWord Filter 		
' .BadWordList = "aa|bb|xx"						| Let 						| null 						| List of custom bad words seperate 
' .UTF8Converter(string) 							| Function 				| string 					| Convert htmlcode char turkish char
' .SQLInjectionBlocker(String) 				| Function 				| string 					| clean before insert sql injection code (Special chars...)
' .Exist(string) 											| Function 				| boolean 				| check value null,empty or len greater then zero
' .DateDiffDay(vStartDate, vEndDate)  | function 				| string 					| return two date difference with day, hour, second
' .SQLDateTime(varDate) 							| function 				| datetime 				| return sql formatted datetime (yyyy-mm-dd hh:mm:ss)
' .TimesAgo(varDate) 									| function 				| datetime 				| return times ago (string)
' .TrimWords(vLongWords, WordCount) 	| function 				| string 					| trim count words...
' .MoneyFormat(vPrice) 								| function 				| double 					| Money Format 1.000,00 => 1000.00

'---------------------------------------------------------------------------------------------------------
' MYSQL CONNECTION HEPLERS
'---------------------------------------------------------------------------------------------------------
' .Debug 															| (Let|Get) 			| null 						| write to screen sql query with pre tag. Repeat beafore all helper call
' .Host                      					| Let 						| null 						| MySQL Connection: Set Host
' .Database                  					| Let 						| null 						| MySQL Connection: Set Database Name
' .User                      					| Let 						| null 						| MySQL Connection: Set Database User
' .Password                  					| Let 						| null 						| MySQL Connection: Set Database Password 
' .Driver(DriverNameString) 					| Let 						| null 						| "MySQL ODBC 3.51 Driver|MySQL ODBC 5.2 ANSI Driver" 
' .Connect()  												| Get 						| object ( db ) 	| Mysql Connection Open
' .DisConnect() 											| Let 						| null 						| Mysql Connection Close
' .TableExist(tableName) 							| 
' .MaxID(tableName) 									| 
' .CountRow(tbl_name, col, where) 		| Let 						| numeric 				| CountRow Value
' .RunCount() 												| let 						| numeric 				| return countrow query run with SQL_CALC_FOUND_ROWS command...
' .InField(String, rsObj, dimension) 	| funtion 				| numeric|null 		| if exist, return index number, not exist, return null
' .ListField(rsObj) 									| Get 						| array 					| get rs elements col name to array 0,1,2

'---------------------------------------------------------------------------------------------------------
' MYSQL INSERT UPDATE DELETE CMD ETC...
'---------------------------------------------------------------------------------------------------------
' .CollectForm("INSERT|UPDATE") 			| let 						| 
' .Rows() 														| get  						| 
' .Values() 													| get  						| 
' .AppendValues=string 								| let 						|  
' .AppendRows=string 									| let 						| 
' .Run(sql) 													| let 						| 
' 					.Run("SELECT ID FROM tbl_tableName WHERE ID = {ID} ")
' 					.Run("SELECT ID FROM tbl_tableName WHERE ID = "& .Data("ID") &" ")
' 					.Run("SELECT ID FROM tbl_tableName WHERE ID = 1 ")
' .RunExtend("INSERT|UPDATE", vTable, "ID={ID}")  | If Method INSERT then, return record id else return true boolen

'---------------------------------------------------------------------------------------------------------
' TIMING FUNCTION
'---------------------------------------------------------------------------------------------------------
' .StartTime() 												| Let 						| big number 			| script execute start time (always return 0.0000000000 sec)
' .NowTime() 													| Let 						| big number 			| script execute time from start (all call return now time from start execute )
' .StopTime() 												| Let 						| big number 			| class terminate time

'---------------------------------------------------------------------------------------------------------
' USER DATAS
'---------------------------------------------------------------------------------------------------------
' .IPAddress() 												| Let 						| string 					| return user IP (if on cloudflare dns else...)

'---------------------------------------------------------------------------------------------------------
' JSON HEPLER
'---------------------------------------------------------------------------------------------------------
' .jsonResponse(200, "Ok") 						| Function 				| json 						| Basic Json Helper : return 2 params (string and int)

'---------------------------------------------------------------------------------------------------------
' SPECIAL FORCES
'---------------------------------------------------------------------------------------------------------
' .CallSub(subbName) 									| sub 						| evulate 				| call evulate a code block  CallSub("IPAddress()")
' .PageContentType("json|html|txt") 	| let 						| null 						| change page content mime type
' .echo(string) 											| sub 						| response 				| short for response.write command
' .Go(url) 														| sub 						| response 				| short for response.redirect command
' .Abandon()													| sub 						| response 				| short for Session.Abandon
' .Kill 															| sub 						| response 				| short for response.end

'---------------------------------------------------------------------------------------------------------
' INTERACTIONS
'--------------------------------------------------------------------------------------------------------- 
' .PushMesaj(Title, Mesaj) 						| function 				| json 						| PUSHBULLET INTEGRATION 
' 

'--------------------------------------------------------------------------------------------------------- 
' DATA VALUES
'--------------------------------------------------------------------------------------------------------- 
' .Data(valueKey) 										| function 				| string 					| capture key value from POST,GET,404_GET
' .URLFrom404(valueKey) 							| function 				| string 					| capture key value from 404_GET
' .DebugForm() 												| sub 						| response 				| capture all POST form, print to values and keys
' .FindInArray(String, Array) 				| function 				| null|index 			| search string in array list...
' .AllowedMethod("GET|PUT|POST|DELETE") 

'--------------------------------------------------------------------------------------------------------- 
' CLASS HELPERS & DEBUGS
'--------------------------------------------------------------------------------------------------------- 
' .CollectedError() 									| sub 						| response 				| Get collected error in array (0 dimension)
' .collectError(String) 							| sub  						| evulate 				| Collect execute handler errors
' .collectInfo() 											| sub 						| response 				| Get collected info in array (0 dimension)
' .CollectedInfo() 										| sub 						| response 				| Get collected info in array (0 dimension)




' Olası Hatalı Toplama
Dim QCollectedErrors()
CollectedErrorsSize=0
Dim CollectedInfos()
CollectedInfosSize=0
Dim db

Class QueryManager 
	'----------------------------------------
	' Public and Private Variable
	'----------------------------------------
	Private db
	Public ExecuteTime
	Public Sorgu, sql_add, sql1, sql2, TotalRows
	Public DebugMode
	Public bad_DBName, bad_DBUser, bad_DBPass, bad_DBServer, bad_DriverType
	Public BadWords

	'----------------------------------------
	' Class Initialize
	'----------------------------------------
	Private Sub Class_Initialize()
		Set db = Conn
		ExecuteTime 			= Timer

		bad_DBName 				= ""
		bad_DBUser 				= ""
		bad_DBPass 				= ""
		bad_DBServer 			= ""
		bad_DriverType 		= "MySQL ODBC 3.51 Driver" 'MySQL ODBC 5.2 ANSI Driver


		DebugMode 				= False
		sql_add 					= ""
		sql1 							= ""
		sql2 							= ""
		TotalRows 				= 0

		BadWords 					= "bok|yarak|am|amcık|amını|amini|yarrak|orospu|göt|salak|aptal|sik|sikik|sikko|manyak|gerizekalı|geri zekalı|sikis|sikti|sikme|sokus|sokma|domal|amcik|xxx|yarak|yarrak|domal|amcik|anal|azdirici|azgın|ateşli|sex|seks|bakire|sevişmek|fantazi|fantezi|erotik|erotic|porn|porno|ensest"
	End Sub

	'----------------------------------------
	' Class Terminate
	'----------------------------------------
	Private Sub Class_Terminate()
		'Set Run 					= Nothing
		Set Sorgu 				= Nothing
		Set db 						= Nothing

		'If Err <> 0 Then echo("<code>Err Found !</code>")
	End Sub


























































	'----------------------------------------
	' .ContentType Page Content Type
	'----------------------------------------
	Public Property Let PageContentType(vType)
		vType = LCase(vType)

		Select Case vType
			Case "json"
				response.ContentType = "application/json"
			Case "txt"
				response.ContentType = "text/plain"
			Case Else
				response.ContentType = "text/HTML"
		End Select
	End Property


'###################################################################################################
'################################ TIMER FUNCTION
'###################################################################################################
	'----------------------------------------
	' Script Execute Time (Start)
	'----------------------------------------
	Public Property Get StartTime()
		StartTime = ExecuteTime
	End Property
	'----------------------------------------
	' Script Execute Time (Now sec from start of execute)
	'----------------------------------------
	Public Property Get NowTime()
		NowTime = Timer-ExecuteTime
	End Property
	'----------------------------------------
	' Class Terminate Time
	'----------------------------------------
	Public Property Get StopTime()
		StopTime = Timer - ExecuteTime
	End Property


'###################################################################################################
'################################ SIMPLE FUNCTION
'###################################################################################################
	'----------------------------------------
	' UTF8 Turkih Char Converter
	'----------------------------------------
	Public Function UTF8Converter(gelenveri)
		gelenveri = Trim(gelenveri)
		gelenveri = Replace(gelenveri ,"%C4%B1","ı",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C4%B0","İ",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C5%9E","Ş",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C5%9F","ş",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C4%9E","Ğ",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C4%9F","ğ",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C3%87","Ç",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C3%A7","ç",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C3%96","Ö",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C3%B6","ö",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C3%9C","Ü",1,-1,1)
		gelenveri = Replace(gelenveri ,"%C3%BC","ü",1,-1,1)
		'gelenveri = Replace(gelenveri ,"aaaa","Ç",1,-1,1)
		gelenveri = Replace(gelenveri ,"+"," ",1,-1,1)
		gelenveri = Replace(gelenveri ,"%20"," ",1,-1,1)
		UTF8Converter = gelenveri
	End Function

	'----------------------------------------
	' .IPAddress() Return Host IP
	'----------------------------------------
	Public Function IPAddress()
    t_IPAdresi = Request.ServerVariables("HTTP_CF-Connecting-IP") & ""
    If Len(t_IPAdresi) < 2 Then t_IPAdresi = Request.ServerVariables("remote_addr")
    IPAddress = t_IPAdresi
	End Function

	'----------------------------------------
	' Badword Extractor
	'----------------------------------------
	Public Function BadWord(gelenveri)
		Dim RegExp
		Set RegExp = New RegExp

		RegExp.IgnoreCase = true
		RegExp.Global 		= True
		RegExp.Pattern 		= "\b(" & BadWords & ")\b"

		BadWord = RegExp.Replace(gelenveri,"*****")

		Set RegExp = Nothing
	End Function
	'----------------------------------------
	' Badword Extractor Set List
	'----------------------------------------
	Public Property Let BadWordList(vNewWords)
		BadWords = vNewWords
	End Property

'###################################################################################################
'################################ JSON
'###################################################################################################
	'----------------------------------------
	' .jsonResponse(200, "Ok")
	'----------------------------------------
	Public Function jsonResponse(vStatus, vMsg)
		Response.Write "{""status"": "&vStatus&", ""messages"": """& vMsg &"""}"
	End Function


'###################################################################################################
'################################ MYSQL 
'###################################################################################################
	'----------------------------------------
	' SQL Connection
	'----------------------------------------
	Public Property Get Connect()
		If Exist(bad_DBName) = False Then 			echo("<code>Please LET Database</code>") 	: Exit Property
		If Exist(bad_DBServer) = False Then 		echo("<code>Please LET Host</code>") 			: Exit Property
		If Exist(bad_DBUser) = False Then 			echo("<code>Please LET User</code>") 			: Exit Property
		If Exist(bad_DBPass) = False Then 			echo("<code>Please LET Password</code>") 	: Exit Property
		If Exist(bad_DriverType) = False Then 	echo("<code>Please LET Driver</code>") 		: Exit Property

		' If TypeName(db) = "Empty" Then 
		' 	Response.Write "<code>Please Set QueryManager db Connection</code>"
		' Else
		' 	Set ConnectionObj = db
		' End If

		'On Error Resume Next
		Set DBConnection = Server.CreateObject("ADODB.Connection")
				DBConnection.Open = "DRIVER={"&bad_DriverType&"};database="&bad_DBName&";server="&bad_DBServer&";uid="&bad_DBUser&";password="&bad_DBPass&";"
				DBConnection.Execute "SET NAMES 'utf8'"
				DBConnection.Execute "SET CHARACTER SET utf8"
				DBConnection.Execute "SET COLLATION_CONNECTION = 'utf8_general_ci'"
		If Err<> 0 Then 
			If DebugMode = True Then echo("<code data-toggle=""tooltip"" title=""Error Number: "& Err.Number &" .Connect()"">* Connection Error</code>")
		Else
			Set db = DBConnection
		End If

		DebugMode = False
	End Property

	'----------------------------------------
	' SQL Connection
	'----------------------------------------
	Public Property Get DisConnect()
		db.Close : Set db = Nothing
	End Property

	'----------------------------------------
	' SQL Connection
	'----------------------------------------
	Public Property Let Database(vData)
		bad_DBName = vData
	End Property

	'----------------------------------------
	' SQL Connection
	'----------------------------------------
	Public Property Let Host(vData)
		bad_DBServer = vData
	End Property

	'----------------------------------------
	' SQL Connection
	'----------------------------------------
	Public Property Let User(vData)
		bad_DBUser = vData
	End Property

	'----------------------------------------
	' SQL Connection
	'----------------------------------------
	Public Property Let Password(vData)
		bad_DBPass = vData
	End Property

	'----------------------------------------
	' SQL Connection
	'----------------------------------------
	Public Property Let Driver(vData)
		bad_DriverType = vData
	End Property

	'----------------------------------------
	' .MaxID("tbl_name") return numeric
	'----------------------------------------
	Public Property Get MaxID(tableName)
		Set LatestID = db.Execute("SELECT MAX(ID) FROM "& tableName &"")
			MaxID = LatestID(0)
		LatestID.Close : Set LatestID = Nothing
	End Property

	'----------------------------------------
	' .TableExist("tbl_name") return boolean
	'----------------------------------------
	Public Property Get TableExist(tableName)
		Set CheckTableExist = db.Execute("SHOW TABLES LIKE '"&tableName&"'")
		If CheckTableExist.Eof Then 
			TableExist = False
		Else
			TableExist = True
		End If
		CheckTableExist.Close : Set CheckTableExist = Nothing
	End Property

	'----------------------------------------
	' Form Collertor
	'----------------------------------------
	Public Sub CollectForm(CollectType)
		If CollectType = "INSERT" Then
		  For Each Item In Request.Form
		      fieldName = Item
		      fieldValue = Request.Form(Item)

		      If fieldName = "file" OR fieldName = "FAVICON" Then 

		      ElseIf Left(fieldName, 7) = "ATTACH_" Then

		      ElseIf fieldName = "PASSWORD" OR fieldName = "SIFRE" OR fieldName = "PAROLA" Then
		          SQLCumle1 = SQLCumle1 & fieldName &", "
		          SQLCumle2 = SQLCumle2 & "'"& MD5( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "KIMLIK_NO" Then
		          If IsNumeric(SQLInjectionBlocker(Request.Form(fieldName))) Then SQLCumle1 = SQLCumle1 & fieldName &", "
		          If IsNumeric(SQLInjectionBlocker(Request.Form(fieldName))) Then SQLCumle2 = SQLCumle2 & "'"& Trim(SQLInjectionBlocker(Request.Form(fieldName))) &"',"
		      ElseIf fieldName = "KURS_SURESI" Then
		          SQLCumle1 = SQLCumle1 & fieldName &", "
		          SQLCumle2 = SQLCumle2 & "'"& ParaTemizle( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
		       ElseIf fieldName = "UCRET" OR fieldName = "KURS_UCRETI" OR fieldName = "FIYAT" OR fieldName = "DERS_UCRETI" Then
		          If Len(SQLInjectionBlocker(Request.Form(fieldName))) > 4 Then SQLCumle1 = SQLCumle1 & fieldName &", "
		          If Len(SQLInjectionBlocker(Request.Form(fieldName))) > 4 Then SQLCumle2 = SQLCumle2 & "'"& ParaTemizle( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "KURS_YILI" OR fieldName = "YIL"  Then
		          SQLCumle1 = SQLCumle1 & fieldName &", "
		          SQLCumle2 = SQLCumle2 & "'"& SQLDate( SQLInjectionBlocker(Request.Form(fieldName)) & "-01-01" ) &"',"
		      ElseIf fieldName = "DOGUM_TARIHI" Then
		          SQLCumle1 = SQLCumle1 & fieldName &", "
		          SQLCumle2 = SQLCumle2 & "'"& SQLDate( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "BITIS_TARIHI" OR fieldName = "BASLANGIC_TARIHI" OR fieldName = "GUNCELLENME_TARIHI" OR fieldName = "EKLENME_TARIHI" OR fieldName = "OGRENCI_KAYIT_BASLANGIC" OR fieldName = "OGRENCI_KAYIT_BITIS" Then
		          If Len(SQLInjectionBlocker(Request.Form(fieldName))) > 4 Then SQLCumle1 = SQLCumle1 & fieldName &", "
		          If Len(SQLInjectionBlocker(Request.Form(fieldName))) > 4 Then SQLCumle2 = SQLCumle2 & "'"& BlogDateFunc( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
		      Else
		          SQLCumle1 = SQLCumle1 & ""& fieldName &","
		          SQLCumle2 = SQLCumle2 & "'"& SQLInjectionBlocker(Request.Form(fieldName)) &"',"
		      End If
		  Next
		End If

		If CollectType = "UPDATE" Then 
	    For Each Item In Request.Form
	        fieldName = Item
	        fieldValue = Request.Form(Item)

	        If fieldName = "file" OR fieldName = "FAVICON" Then 

	        ElseIf fieldName = "PASSWORD" Then
	            tmp_pass = Trim(SQLInjectionBlocker(Request.Form(fieldName)))
	            If Len(tmp_pass) > 4 Then  SQLCumle2 = SQLCumle2 & "PASSWORD = '"& MD5( tmp_pass ) &"',"
	        ElseIf fieldName = "KIMLIK_NO" Then
	            If IsNumeric(SQLInjectionBlocker(Request.Form(fieldName))) Then SQLCumle2 = SQLCumle2 & "KIMLIK_NO='"& SQLInjectionBlocker(Request.Form(fieldName)) &"',"
		      ElseIf fieldName = "KURS_SURESI" Then
		          SQLCumle2 = SQLCumle2 & "'"& ParaTemizle( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "UCRET" OR fieldName = "KURS_UCRETI" OR fieldName = "DERS_UCRETI" Then
		          If Len(SQLInjectionBlocker(Request.Form(fieldName))) > 4 Then SQLCumle2 = SQLCumle2 & "'"& ParaTemizle( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "KURS_YILI" OR fieldName = "YIL"  Then
		          SQLCumle2 = SQLCumle2 & "'"& SQLDate( SQLInjectionBlocker(Request.Form(fieldName)) & "-01-01" ) &"',"
	        ElseIf fieldName = "DOGUM_TARIHI" Then
	            SQLCumle2 = SQLCumle2 & ""&fieldName&"='"& SQLDate( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "BITIS_TARIHI" OR fieldName = "BASLANGIC_TARIHI" OR fieldName = "GUNCELLENME_TARIHI" OR fieldName = "EKLENME_TARIHI" OR fieldName = "OGRENCI_KAYIT_BASLANGIC" OR fieldName = "OGRENCI_KAYIT_BITIS" Then
	            SQLCumle2 = SQLCumle2 & ""&fieldName&"='"& BlogDateFunc( SQLInjectionBlocker(Request.Form(fieldName)) ) &"',"
	        Else
	            SQLCumle2 = SQLCumle2 & ""& fieldName &"='"& SQLInjectionBlocker(Request.Form(fieldName)) &"',"
	        End If
	    Next
		End If

		sql1 = SQLCumle1
		sql2 = SQLCumle2

		If DebugMode = True Then 
	  	With Response
	  		.Write "<pre>"
		  		.Write "<code>"
			  		.Write "Return Rows: "&sql1&"" & vbcrlf & vbcrlf
			  		.Write "Return Values: "&sql2&""& vbcrlf & vbcrlf
		  		.Write "</code>"
		  		.Write "-------- DEBUG MODE ON SYSTEM FORCE ENDED -------"
	  		.Write "</pre>"
	  	End With
		End If
	End Sub

	'----------------------------------------
	' Returned Row Count
	'----------------------------------------
	Public Property Get RunCount()
		RunCount = TotalRows
	End Property

	'----------------------------------------
	' .CountRow("tbl_name", "id", "WHERE X='y'") return numeric
	'----------------------------------------
	Public Property Get CountRow(table, CountRowName, WhereCondition)
		if Len(WhereCondition) > 0 Then NewWhereCondition = ReplaceChar(WhereCondition) : sql_add = " WHERE "& NewWhereCondition &""
		On Error Resume Next
		If TableExist(table) = False Then 
			CountRow = "<code data-toggle=""tooltip"" title=""Table Not Exist"">* Query Error</code>"
			Exit Property
		End If

		CountRow = db.Execute("SELECT COUNT("&CountRowName&") FROM "& table & sql_add &"")(0)
		if Err <> 0 then
			CountRow = "<code data-toggle=""tooltip"" title=""Error: "& Err.Description &"<br>Error Number: "& Err.Number &" .CountRow(a,b,c)"">* Query Error</code>"
		Else

		End If
	End Property

	'----------------------------------------
	' .DebugForm  return form collection debug
	'----------------------------------------
	Public Sub DebugForm()
    With Response
        .Write "<pre>"
        For each x in Request.Form
            .Write x & ": "
            .Write SQLInjectionBlocker(Request.Form(""& x &"")) & vbcrlf
        Next
        .Write "</pre><hr />"
    End With
	End Sub

	'----------------------------------------
	' Query.Run("SELECT ID FROM tbl_tableName WHERE ID = {ID} ")
	' Query.Run("SELECT ID FROM tbl_tableName WHERE ID = "& .Data("ID") &" ")
	' Query.Run("SELECT ID FROM tbl_tableName WHERE ID = 1 ")
	'----------------------------------------
	Public Property Get Run(sql)
		sql = ReplaceChar(sql)
		TotalRows=0
		If DebugMode = True Then
	  	With Response
	  		.Write "<pre>"
		  		.Write "<code>"
			  		.Write "Return SQL Structure: "&sql&"" & vbcrlf & vbcrlf
		  		.Write "</code>"
		  		.Write "-------- DEBUG MODE ON SYSTEM FORCE ENDED -------"
	  		.Write "</pre>"
	  	End With
	  	Response.End
		Else
			' Implement Query
			On Error Resume Next
			Set Sorgu = db.Execute(sql)
			' Errorleri Yakala
			If Err <> 0 Then 
			  str_frm = ""
			  For each x in Request.Form
			      str_frm = str_frm & x & ": "
			      str_frm = str_frm & LoginKontrol(Request.Form(""& x &""))
			  Next
			  If Len(str_frm) < 1 Then str_frm = "(Forms Empty)"
				Mesaj = "WEB: "& Request.ServerVariables("SERVER_NAME") &"\n\n URL: "& Request.ServerVariables("QUERY_STRING") &"\n\n FORM: "&str_frm&"\n\nError Code: "& Err.Description &" "
				a= PushMesaj("Title", Mesaj)
			
				Response.Write Err.Line 
				Response.Write "<br>"
				Response.Write Err.Description 
				Response.Write "<br>"
				Exit Property
			End If
			'On Error GoTo 0

			' Get All Row Count
			If Instr(1, sql, "SQL_CALC_FOUND_ROWS") <> 0 Then
				Set SorguCount = db.Execute("SELECT FOUND_ROWS()")
				TotalRows = SorguCount(0)
				If IsNumeric(TotalRows) Then TotalRows = Cint(TotalRows)
				SorguCount.Close : Set SorguCount = Nothing
			Else
				TotalRows = 0
			End If

			Set Run = Sorgu	
		End If

		' Debug Tekrar Kapat
		DebugMode = False
	End Property

	'----------------------------------------
	'
	'----------------------------------------
	Public Property Get Rows()
		Rows = sql1
	End Property
	Public Property Let AppendRows(data)
		sql1 = sql1 & data
	End Property

	'----------------------------------------
	'
	'----------------------------------------
	Public Property Get Values()
		Values = sql2
	End Property
	Public Property Let AppendValues(data)
		sql2 = sql2 & data
	End Property

	'----------------------------------------
	' .SQL Query ParamConverter  {Param}
	'----------------------------------------
	Public Function ReplaceChar(strString)
		Set KeyArr = RegExResults(strString, "{(.*?)\}")
		For each result in KeyArr
		    DegisimAnahtari = result.Submatches(0)
		    strString = Replace(strString , "{"& DegisimAnahtari &"}" , Query.Data(DegisimAnahtari) )
		Next
		Set KeyArr = Nothing

		ReplaceChar = strString
	End Function

	'----------------------------------------
	' .SQL Query ParamConverter  {Param}
	'----------------------------------------
	Public Function RegExResults(strTarget, strPattern)
	    Set regEx = New RegExp
	      regEx.Pattern = strPattern
	      regEx.Global = true
	    Set RegExResults = regEx.Execute(strTarget)
	    Set regEx = Nothing
	End Function

	'----------------------------------------
	' SQLInjection Blocker
	'----------------------------------------
	Public Function SQLInjectionBlocker(vData)
		vData = Replace(vData, "script", "&#115;cript", 1, -1, 0)
		vData = Replace(vData, "SCRIPT", "&#083;CRIPT", 1, -1, 0)
		vData = Replace(vData, "Script", "&#083;cript", 1, -1, 0)
		vData = Replace(vData, "script", "&#083;cript", 1, -1, 1)
		vData = Replace(vData, "object", "&#111;bject", 1, -1, 0)
		vData = Replace(vData, "OBJECT", "&#079;BJECT", 1, -1, 0)
		vData = Replace(vData, "Object", "&#079;bject", 1, -1, 0)
		vData = Replace(vData, "object", "&#079;bject", 1, -1, 1)
		vData = Replace(vData, "document", "&#100;ocument", 1, -1, 0)
		vData = Replace(vData, "DOCUMENT", "&#068;OCUMENT", 1, -1, 0)
		vData = Replace(vData, "Document", "&#068;ocument", 1, -1, 0)
		vData = Replace(vData, "document", "&#068;ocument", 1, -1, 1)
		vData = Replace(vData, "cookie", "&#099;ookie", 1, -1, 0)
		vData = Replace(vData, "COOKIE", "&#067;OOKIE", 1, -1, 0)
		vData = Replace(vData, "Cookie", "&#067;ookie", 1, -1, 0)
		vData = Replace(vData, "cookie", "&#067;ookie", 1, -1, 1)
		vData = Replace(vData, "document.cookie", "&#068;ocument.cookie", 1, -1, 1)
		vData = Replace(vData, "javascript:", "javascript ", 1, -1, 1)
		vData = Replace(vData, "applet", "&#097;pplet", 1, -1, 0)
		vData = Replace(vData, "APPLET", "&#065;PPLET", 1, -1, 0)
		vData = Replace(vData, "Applet", "&#065;pplet", 1, -1, 0)
		vData = Replace(vData, "applet", "&#065;pplet", 1, -1, 1)
		vData = Replace(vData, "UNION", "", 1, -1, 0)
		vData = Replace(vData, "union", "", 1, -1, 0)
		vData = Replace(vData, "Union", "", 1, -1, 0)
		vData = Replace(vData, "vbscript:", "vbscript ", 1, -1, 1)
		vData = Replace(vData, "'", "&apos;")
		vData = Replace(vData, chr(39), "&apos;")
		vData = Replace(vData, "%20", " ")
		SQLInjectionBlocker = vData
	End Function

	'----------------------------------------
	' .SQLDateTime(data)
	'----------------------------------------
	Public Function SQLDateTime(varDate)
	  If day(varDate) < 10 then
	    dd = "0" & day(varDate)
	  Else
	    dd = day(varDate)
	  End If
	  If month(varDate) < 10 then
	    mm = "0" & month(varDate)
	  Else
	    mm = month(varDate)
	  End If
	  If hour(varDate) < 10 then
	    hh = "0" & hour(varDate)
	  Else
	    hh = hour(varDate)
	  End If
	  If minute(varDate) < 10 then
	    mi = "0" & minute(varDate)
	  Else
	    mi = minute(varDate)
	  End If
	  If second(varDate) < 10 then
	    se = "0" & second(varDate)
	  Else
	    se = second(varDate)
	  End If
	  SQLDateTime = ""& year(varDate) &"-"& mm &"-"& dd &" "& hh &":"& mi &":"& se &""
	End Function

	'----------------------------------------
	' .DateDiffDay(data)
	'----------------------------------------
	Public Function DateDiffDay(vStartDate, vEndDate)
		If IsNull(vEndDate) Or IsEmpty(vEndDate) Or vEndDate = "" Then vEndDate = Now()
	
		If IsNull(vStartDate) OR IsEmpty(vStartDate) Then 
			Exit Function
		End If

		CountdownDate = vStartDate
		theDate     = Now()
		DaysLeft    = DateDiff("d",theDate,CountdownDate) '- 1
		
		theDate     = DateAdd("d",DaysLeft,theDate)
		HoursLeft   = DateDiff("h",theDate,CountdownDate) '- 1
		
		theDate     = DateAdd("h",HoursLeft,theDate)
		MinutesLeft = DateDiff("n",theDate,CountdownDate) '- 1
		
		theDate     = DateAdd("n",MinutesLeft,theDate)
		SecondsLeft = DateDiff("s",theDate,CountdownDate) '- 1

		DateDiffDay = ""& DaysLeft &" Gün "& HoursLeft &" Saat "& MinutesLeft &" Dakika"
	End Function

	'----------------------------------------
	' .DateDiffDay(data)
	'----------------------------------------
	Public Function TrimWords(Kelime,Karakter)
	    Kelime = Kelime & ""
			If IsNull(Karakter) OR IsEmpty(Karakter) Or Not IsNumeric(Karakter) Then Karakter = 100
	    If len(Kelime)>40 Then
	        Degiskensayma = mid(Kelime,Karakter,len(Kelime))
	        noktaninyeri = inStr(Degiskensayma," ")-1
	        TrimWords = Left(Kelime,Karakter+noktaninyeri)
	    Else
	        TrimWords = Kelime
	    End If
	End Function

	'----------------------------------------
	' .DateDiffDay(data)
	'----------------------------------------
	Public Function TimesAgo(dt)
	  If IsNull(dt) OR IsEmpty(dt) Then Exit Function
	  Dim t_SECOND : t_SECOND = 1
	  Dim t_MINUTE : t_MINUTE = 60 * t_SECOND
	  Dim t_HOUR : t_HOUR = 60 * t_MINUTE
	  Dim t_DAY : t_DAY = 24 * t_HOUR
	  Dim t_MONTH : t_MONTH = 30 * t_DAY
	  Dim delta : delta = DateDiff("s", dt, Now)
	  Dim strTime : strTime = ""

	  If (delta < 1 * t_MINUTE) Then
	    If delta = 0 Then
	      strTime = GetLang("simdi")
	    ElseIf delta = 1 Then
	      strTime = GetLang("bir_saniye_once")
	    Else
	      strTime = delta & GetLang("x_saniye_once")
	    End If
	  ElseIf (delta < 2 * t_MINUTE) Then
	    strTime = GetLang("bir_dakika_once")
	  ElseIf (delta < 50 * t_MINUTE) Then
	    'strTime = max(Round(delta / t_MINUTE), 2) & GetLang("x_dakika_once")
	    strTime = Round(delta / t_MINUTE) & GetLang("x_dakika_once")
	  ElseIf (delta < 90 * t_MINUTE) Then
	    strTime = GetLang("bir_saat_once")
	  ElseIf (delta < 24 * t_HOUR) Then
	    strTime = Round(delta / t_HOUR) & GetLang("x_saat_once")
	  ElseIf (delta < 48 * t_HOUR) Then
	    strTime = GetLang("dun")
	  ElseIf (delta < 30 * t_DAY) Then
	   strTime = Round(delta / t_DAY) & GetLang("x_gun_once")
	  ElseIf (delta < 12 * t_MONTH) Then
	    Dim months
	    months = Round(delta / t_MONTH)
	    If months <= 1 Then
	      strTime = GetLang("bir_ay_once")
	    Else
	      strTime = months & GetLang("x_ay_once")
	    End If
	  Else
	    Dim years : years = Round((delta / t_DAY) / 365)
	    If years <= 1 Then
	        strTime = GetLang("bir_yil_once")
	    Else
	      strTime = years & GetLang("x_yil_once")
	    End If
	  End If
	  TimesAgo = strTime
	End Function

'###################################################################################################
'################################ SPECIAL FORCES 
'###################################################################################################
	'----------------------------------------
	' Evulate Call Sub
	'----------------------------------------
	Public Sub CallSub(vData)
		If Not IsNull(vData) Or Not IsEmpty(vData) Or Not vData = "" Then Eval("Call "& vData &"")
	End Sub

	'----------------------------------------
	' Short to Response.Write 
	'----------------------------------------
	Public Sub echo(vData)
		Response.Write vData
	End Sub

	'----------------------------------------
	' .Go(ToURL)
	'----------------------------------------
	Public Sub Go(url)
		Response.Redirect ReplaceChar(url)
		'Response.End()
	End Sub

	'----------------------------------------
	' .Kill Terminate All Code
	'----------------------------------------
	Public Sub Kill()
		Response.End
	End Sub

	'----------------------------------------
	' .Kill Terminate All Code
	'----------------------------------------
	Public Sub Abandon()
		Session.Abandon()
	End Sub

'###################################################################################################
'################################ DATA
'###################################################################################################
	'----------------------------------------
	' Method Restrict
	'----------------------------------------
	Public Property Get AllowedMethod(vMethod)
		AllowedMethod = True
		MethodNot 		= Request.ServerVariables("REQUEST_METHOD")
		Select Case vMethod
			Case "GET"
				If Not MethodNot = vMethod Then  AllowedMethod = False
			Case "POST"
				If Not MethodNot = vMethod Then  AllowedMethod = False
			Case "PUT"
				If Not MethodNot = vMethod Then  AllowedMethod = False
			Case "DELETE"
				If Not MethodNot = vMethod Then  AllowedMethod = False
			Case Else 
				
		End Select

	End Property

	'----------------------------------------
	' Get String form request or 404 url params
	'----------------------------------------
	Public Property Get Data(vData)
		If IsNull(vData) Or IsEmpty(vData) Or Len(vData) < 0 Then 
			Data = ""
		Else 
			Data = Trim(SQLInjectionBlocker(Request(""&vData&"")))
			If IsNull(Data) Or IsEmpty(Data) OR Len(Data) < 1 Then Data = URLFrom404( vData )
			If IsNumeric(Data) Then Data = Trim(Data)
		End If
	End Property

	'----------------------------------------
	' URL Parse
	'----------------------------------------
	Public Function URLFrom404(Hangisi)
		DataVal 		= Request.ServerVariables("QUERY_STRING") & "&s=x"
	  Hangisi = Hangisi & "="

    dim sResult 
    dim lStart
    dim lEnd
    
    lStart = instr( 1, DataVal, Hangisi, 1 )
    if lStart > 0 then 
        lStart = lStart + len(Hangisi)
        lEnd = instr( lStart, DataVal, "&" )
        if lEnd = 0 then lEnd = len( DataVal )
        
        sResult = mid( DataVal, lStart, lEnd - lStart )
    end if
    
    URLFrom404 = SQLInjectionBlocker(sResult)
	End function

	'----------------------------------------
	' Value Exist ?
	'----------------------------------------
	Public Function Exist(vData)
		If IsNull(vData) OR IsEmpty(vData) Or Trim(vData) = "" OR Len(vData) = 0 Then
			Exist = False
		Else
			Exist = True
		End If
	End Function

	'----------------------------------------
	' String Search In Array
	'----------------------------------------
	Public Function FindInArray(vString, ArrayName)
	  Dim i
	  'FindInArray = False
	  FindInArray = Null
	  For i=0 To Ubound(ArrayName)
	    If Trim(ArrayName(i)) = Trim(vString) Then
	      'FindInArray = True
	      FindInArray = i
	      Exit Function      
	    End If
	  Next
	End Function

	Public Function FindInMDArray(vString, ArrayName, ArrayTotalBound, ArraySearchIndex)
	  Dim i
	  FindInMDArray = Null
	  For i=0 To ArrayTotalBound
	    If Trim(ArrayName(i, ArraySearchIndex)) = Trim(vString) AND Not Trim(ArrayName(i, ArraySearchIndex)) = "" Then
		  	'Response.Write ""& Trim(vString) &" Found At "& i &" ArrIndex: "& ArraySearchIndex &"<br>"
	      FindInMDArray = i
	      Exit Function
	    End If
	  Next
	End Function

	'----------------------------------------
	' Array Dimensions
	'----------------------------------------
	Public Function NumDimensions(arr)
		Dim dimensions : dimensions = 0
		On Error Resume Next
		Do While Err.number = 0
		    dimensions = dimensions + 1
		    UBound arr, dimensions
		Loop
		On Error Goto 0
		NumDimensions = dimensions - 1
	End Function

	'----------------------------------------
	' Recordset Field Name Exist return index number
	'----------------------------------------
	Public Function InField(fieldName, rsObj, Dimension)
		if IsNull(Dimension) OR IsEmpty(Dimension) OR Not IsNumeric(Dimension) OR Len(Dimension) < 0 OR Dimension = "" Then Dimension = 0

		If IsEmpty(fieldName) Or IsNull(fieldName) Then 
			InField = Null 
			Exit Function
		End If
		P=0 : Dim fieldArr()
		ReDim PRESERVE fieldArr(rsObj.Fields.count, 2)  
		For each item in rsObj.Fields
			fieldArr(P,0) = item.Name
			fieldArr(P,1) = item.Type
			fieldArr(P,2) = FieldTypeName(item.Type)
			P=P+1
		Next

		'InField = FindInMDArray(fieldName, fieldArr, rsObj.Fields.count, 0)
		InField = FindInMDArray(fieldName, fieldArr, rsObj.Fields.count, Dimension)
	End Function

	'----------------------------------------
	' Recordset Field Name to Array
	'----------------------------------------
	Public Property Get ListField(rsObj)
		P=0 : Dim fieldArr()
		ReDim PRESERVE fieldArr(rsObj.Fields.count, 2)  
		For each item in rsObj.Fields
			fieldArr(P,0) = item.Name
			fieldArr(P,1) = item.Type
			fieldArr(P,2) = FieldTypeName(item.Type)
			P=P+1
			'fieldArr(P) = item.Name & "-"& item.Type &"-"& FieldTypeName(item.Type) &" " : P=P+1
		Next
		ListField = fieldArr
	End Property

	Public Function FieldTypeName(vCode)
		Select Case vCode 
			Case 0 	: FieldTypeName = "NO_VALUE"
			Case 2 	: FieldTypeName = "INT"
			Case 3 	: FieldTypeName = "INT"
			Case 4 	: FieldTypeName = "INT"
			Case 5 	: FieldTypeName = "DOUBLE"
			Case 7 	: FieldTypeName = "DATE"
			Case 10 : FieldTypeName = "ERROR"
			Case 11 : FieldTypeName = "BOOLEAN"
			Case 12 : FieldTypeName = "VARIANT"
			Case 135 : FieldTypeName = "DATETIME"
			Case 200 : FieldTypeName = "VARCHAR"
			Case 201 : FieldTypeName = "VARCHAR"
			Case 202 : FieldTypeName = "VARCHAR"
			Case 203 : FieldTypeName = "VARCHAR"
			Case 204 : FieldTypeName = "VARCHAR"
			Case 205 : FieldTypeName = "VARCHAR"
			Case Else 
				FieldTypeName = "NO_INFO"
		End Select
	End Function

	Public Property Get GetFieldTypeName(rsObj, IndexName)
		GetFieldTypeName = FieldTypeName( rsObj.Fields(IndexName).Type )
	End Property

	'----------------------------------------
	' SQL Command Catcher
	'----------------------------------------
	Public Property Get RunExtend(vMethod, vTable, vUpdateKey)
		If Query.AllowedMethod("POST") = False Then
			collectError("RunExtend("&vMethod&", "&vTable&")->AllowedMethod->Method Not Allowed. Only POST Method ("& NowTime() &")")
			RunExtend = False
			Exit Property
		End If

		t="" : G=0 : Dim ExtArray() : Dim ExtArray2()

		' Gelen Formları Al
		FormData = FormsNameToArray()

		' Tabloyu Kontrol Et 
		If TableExist(vTable) = False Then
			collectError("RunExtend("&vMethod&", "&vTable&")->Case->TableExist("&vTable&")->Table Not Found. ("& NowTime() &")")
			RunExtend = False
			Exit Property
		End If

		' Tablo Fieldları Al
		Set tmp_rs = db.Execute("SELECT * FROM "& vTable &" LIMIT 1")
		'DataField= ListField(tmp_rs)

Select Case vMethod
			'--------------------------------------------------------------------------------------
Case "INSERT"
			'--------------------------------------------------------------------------------------
				ii=0
				For i=0 To UBound(FormData)
					str_field_name = FormData(i,0)
					str_field_val  = FormData(i,1)

					control_exist = TypeName( InField(str_field_name, tmp_rs, 0) )
					If control_exist = "Null" Then
						collectError("RunExtend("&vMethod&", "&vTable&")->Case->InField("&str_field_name&", obj, 0)->Field Not Found In Table. ("& NowTime() &")")
					Else
						ReDim PRESERVE ExtArray(ii)
						ReDim PRESERVE ExtArray2(ii)
							control_type = GetFieldTypeName(tmp_rs, str_field_name)
							Select Case control_type
								Case "DOUBLE"
									'ExtArray(ii) 		= ""& str_field_name &" = '"& str_field_val &"'["& GetFieldTypeName(tmp_rs, str_field_name) &"]"
									ExtArray(ii) 		= Trim(str_field_name)
									ExtArray2(ii) 	= ""& MoneyFormatter(str_field_val) &""
								Case "INT"
									'ExtArray(ii) 		= ""& str_field_name &" = '"& str_field_val &"'["& GetFieldTypeName(tmp_rs, str_field_name) &"]"
									ExtArray(ii) 		= Trim(str_field_name)
									ExtArray2(ii) 	= ""& Trim(str_field_val) &""
								Case "DATETIME"
									If IsDate(str_field_val) = True Then
										ExtArray(ii) 		= Trim(str_field_name)
										ExtArray2(ii) 	= "'"& SQLDateTime(Trim(str_field_val)) &"'"
									Else 
										ReDim PRESERVE ExtArray(ii-1) : ii=ii-1
									End If
								Case Else 
									'ExtArray(ii) 		= ""& str_field_name &" = '"& str_field_val &"'["& GetFieldTypeName(tmp_rs, str_field_name) &"]"
									ExtArray(ii) 		= Trim(str_field_name)
									ExtArray2(ii) 	= "'"& Trim(str_field_val) &"'"
							End Select

							collectInfo("RunExtend("&vMethod&", "&vTable&")->Case->InField("&str_field_name&", obj, 0)->Type Is: "& control_type &" ("& NowTime() &")")
							ii=ii+1
						't=t& "#"& InField(str_field_name, tmp_rs) &"#"
						't=t& str_field_name &"="& str_field_val & "<br><br>"
					End If
				Next

				If Not ii=0 Then ReDim PRESERVE ExtArray(ii-1)
				If Not ii=0 Then ReDim PRESERVE ExtArray2(ii-1)

				t_fields = Join( ExtArray, ", ")
				t_Values = Join( ExtArray2, ", ")
				t="INSERT INTO "& vTable &"("& t_fields &") VALUES("& t_Values &") "
				
				Run(t)
				
				ReturnValue = MaxID(vTable)
			'--------------------------------------------------------------------------------------
Case "UPDATE"
			'--------------------------------------------------------------------------------------
				If vUpdateKey = Null Or IsEmpty(vUpdateKey) Or Isnull(vUpdateKey) Then 
					collectError("RunExtend("&vMethod&", "&vTable&", "& vUpdateKey &")->Case->Update Key Error. ("& NowTime() &")")
					RunExtend = False
					Exit Property
				End If

				ii=0
				For i=0 To UBound(FormData)
					str_field_name = FormData(i,0)
					str_field_val  = FormData(i,1)

					control_exist = TypeName( InField(str_field_name, tmp_rs, 0) )
					If control_exist = "Null" Then
						collectError("RunExtend("&vMethod&", "&vTable&")->Case->InField("&str_field_name&", obj, 0)->Field Not Found In Table. ("& NowTime() &")")
					Else
						ReDim PRESERVE ExtArray(ii)
							control_type = GetFieldTypeName(tmp_rs, str_field_name)
							Select Case control_type
								Case "DOUBLE"
									'ExtArray(ii) 		= ""& str_field_name &" = '"& str_field_val &"'["& GetFieldTypeName(tmp_rs, str_field_name) &"]"
									ExtArray(ii) = ""& Trim(str_field_name) &"="& MoneyFormatter(str_field_val) &""
								Case "INT"
									'ExtArray(ii) = ""& str_field_name &" = '"& str_field_val &"'["& GetFieldTypeName(tmp_rs, str_field_name) &"]"
									ExtArray(ii) = ""& Trim(str_field_name) &"="& Trim(str_field_val) &""
								Case "DATETIME"
									If IsDate(str_field_val) = True Then
										ExtArray(ii) = ""& Trim(str_field_name) &"='"& SQLDateTime(Trim(str_field_val)) &"'"
									Else 
										ReDim PRESERVE ExtArray(ii-1) : ii=ii-1
									End If
								Case Else 
									'ExtArray(ii) = ""& str_field_name &" = '"& str_field_val &"'["& GetFieldTypeName(tmp_rs, str_field_name) &"]"
									ExtArray(ii) = ""& Trim(str_field_name) &"='"& Trim(str_field_val) &"'"
							End Select

							collectInfo("RunExtend("&vMethod&", "&vTable&")->Case->InField("&str_field_name&", obj, 0)->Type Is: "& control_type &" ("& NowTime() &")")
							ii=ii+1
						't=t& "#"& InField(str_field_name, tmp_rs) &"#"
						't=t& str_field_name &"="& str_field_val & "<br><br>"
					End If
				Next

				If Not ii=0 Then ReDim PRESERVE ExtArray(ii-1)

				t_KeysAndValues= Join( ExtArray, ", ")
				t="UPDATE "& vTable &" SET "& t_KeysAndValues &" WHERE "& ReplaceChar(vUpdateKey) &""

				Run(t)

				ReturnValue = True
				If Err <> 0 Then ReturnValue = False
			'--------------------------------------------------------------------------------------
Case Else
			'--------------------------------------------------------------------------------------
			ReturnValue = False
			'--------------------------------------------------------------------------------------
End Select

		tmp_rs.Close : Set tmp_rs = Nothing

		RunExtend = ReturnValue
		collectInfo("RunExtend()->["& ReturnValue &"] ("& NowTime() &")")
	End Property

	'----------------------------------------
	' .DebugForm  return form collection debug
	'----------------------------------------
	Public Property Get FormsNameToArray()
		P=0 : i=0 : Dim fieldArr()
		ReDim PRESERVE fieldArr(Request.Form.count, 1)
    For each x in Request.Form
    	fieldArr(i, 0) = SQLInjectionBlocker( x )
    	fieldArr(i, 1) = SQLInjectionBlocker( Request.Form(""& x &"") )
    	i=i+1
    Next

    FormsNameToArray = fieldArr
	End Property

'###################################################################################################
'################################ SETTINGS
'###################################################################################################
	'----------------------------------------
	' .Debug = True / False
	'----------------------------------------
	Public Property Get Debug()
		Debug = DebugMode
	End Property
	Public Property Let Debug(vDebug)
		DebugMode = vDebug
	End Property

	'----------------------------------------
	' .Collect Error
	'----------------------------------------
	Public Sub collectError(HataAciklamasi)
		ReDim PRESERVE QCollectedErrors(CollectedErrorsSize)
		CollectedErrorsSize=CollectedErrorsSize+1

		ReDim PRESERVE QCollectedErrors( CollectedErrorsSize )

		QCollectedErrors(CollectedErrorsSize-1) = HataAciklamasi
	End Sub

	'----------------------------------------
	' .Collected Error
	'----------------------------------------
	Public Sub CollectedError()
		t=""
		If CollectedErrorsSize > 0 Then 
			ErrorRaw=""
			t=t&"<pre>"
			t=t&"**** TOTAL ERROR "& ( CollectedErrorsSize ) &" ****" &vbcrlf
			For Erz=0 To UBound( QCollectedErrors )-1
				t=t& "ERROR("&Erz&"): "& QCollectedErrors(Erz) &"" &vbcrlf
			Next
			t=t&"</pre>"
		Else
			t=t& "<center>Nothing Found =)</center>"
		End If

		Response.Write t
	End Sub

	'----------------------------------------
	' .Collect Info
	'----------------------------------------
	Public Sub collectInfo(HataAciklamasi)
		ReDim PRESERVE CollectedInfos(CollectedInfosSize)
		CollectedInfosSize=CollectedInfosSize+1

		ReDim PRESERVE CollectedInfos( CollectedInfosSize )

		CollectedInfos(CollectedInfosSize-1) = HataAciklamasi
	End Sub

	'----------------------------------------
	' .Collected Error
	'----------------------------------------
	Public Sub CollectedInfo()
		t=""
		If CollectedInfosSize > 0 Then 
			ErrorRaw=""
			t=t&"<pre>"
			t=t&"**** TOTAL INFO "& ( CollectedInfosSize ) &" ****" &vbcrlf
			For Erz=0 To UBound( CollectedInfos )-1
				t=t& "ERROR("&Erz&"): "& CollectedInfos(Erz) &"" &vbcrlf
			Next
			t=t&"</pre>"
		Else
			t=t& "<center>Nothing Found =)</center>"
		End If

		Response.Write t
	End Sub

	'----------------------------------------
	' .Collected Error
	'----------------------------------------
	Public Sub CleanCollect()
		Erase CollectedInfos 		: CollectedInfosSize=0
		Erase QCollectedErrors 	: CollectedErrorsSize=0

		collectInfo("CleanCollect()->All Collect Data Earesed: ("& NowTime() &")")
	End Sub


	'----------------------------------------
	' .MoneyFormat(sVal)
	'----------------------------------------
	Public Function MoneyFormatter(sVal)
		tmp_LCID = Session.LCID
		Session.LCID = 1033
		sVal = Trim(sVal) & ""
		sVal = Replace(sVal ,".","",1,-1,1)
		sVal = Replace(sVal ,",",".",1,-1,1)
		sVal = Replace(sVal ," ","",1,-1,1)
		sVal = Replace(sVal ,"TL","",1,-1,1)
		sVal = Replace(sVal ,"TRY","",1,-1,1)
		MoneyFormatter = FormatNumber(sVal,2,-1,0,0)
		Session.LCID = tmp_LCID
	End Function

End Class
%>
