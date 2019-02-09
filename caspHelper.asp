<%
Set Query = New QueryManager

'####################################
' Custom SQL Query Helper
' Classic ASP & MySQL Query Builder
' 2019  (c) Anthony Burak DURSUN
'####################################
Class QueryManager 
	Private ConnectionObj
	Public Sorgu, sql_add, sql1, sql2, TotalRows
	Public DebugMode

	'----------------------------------------
	'
	'----------------------------------------
	Private Sub Class_Initialize()
		' Set Connection Obj
		Set ConnectionObj = Conn
		DebugMode 				= False
		sql_add 					= ""
		sql1 							= ""
		sql2 							= ""
		TotalRows 				= 0
	End Sub

	'----------------------------------------
	'
	'----------------------------------------
	Private Sub Class_Terminate()
		' Destroy Connection Obj
		Set ConnectionObj = Nothing
	End Sub


	'----------------------------------------
	' .SQL Query ParamConverter
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
			Set Sorgu = ConnectionObj.Execute(sql)
			' Get All Row Count
			If Instr(1, sql, "SQL_CALC_FOUND_ROWS") <> 0 Then
				Set SorguCount = ConnectionObj.Execute("SELECT FOUND_ROWS()")
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
	' Query Returned Total Row
	'----------------------------------------
	Public Property Get RunCount()
		RunCount = TotalRows
	End Property

  '----------------------------------------
	' Form Collector
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
		          SQLCumle2 = SQLCumle2 & "'"& MD5( LoginKontrol(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "KIMLIK_NO" Then
		          If IsNumeric(LoginKontrol(Request.Form(fieldName))) Then SQLCumle1 = SQLCumle1 & fieldName &", "
		          If IsNumeric(LoginKontrol(Request.Form(fieldName))) Then SQLCumle2 = SQLCumle2 & "'"& Trim(LoginKontrol(Request.Form(fieldName))) &"',"
		      ElseIf fieldName = "KURS_SURESI" Then
		          SQLCumle1 = SQLCumle1 & fieldName &", "
		          SQLCumle2 = SQLCumle2 & "'"& ParaTemizle( LoginKontrol(Request.Form(fieldName)) ) &"',"
		       ElseIf fieldName = "UCRET" OR fieldName = "KURS_UCRETI" OR fieldName = "DERS_UCRETI" Then
		          If Len(LoginKontrol(Request.Form(fieldName))) > 4 Then SQLCumle1 = SQLCumle1 & fieldName &", "
		          If Len(LoginKontrol(Request.Form(fieldName))) > 4 Then SQLCumle2 = SQLCumle2 & "'"& ParaTemizle( LoginKontrol(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "KURS_YILI" OR fieldName = "YIL"  Then
		          SQLCumle1 = SQLCumle1 & fieldName &", "
		          SQLCumle2 = SQLCumle2 & "'"& SQLDate( LoginKontrol(Request.Form(fieldName)) & "-01-01" ) &"',"
		      ElseIf fieldName = "DOGUM_TARIHI" Then
		          SQLCumle1 = SQLCumle1 & fieldName &", "
		          SQLCumle2 = SQLCumle2 & "'"& SQLDate( LoginKontrol(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "BITIS_TARIHI" OR fieldName = "BASLANGIC_TARIHI" OR fieldName = "GUNCELLENME_TARIHI" OR fieldName = "EKLENME_TARIHI" OR fieldName = "OGRENCI_KAYIT_BASLANGIC" OR fieldName = "OGRENCI_KAYIT_BITIS" Then
		          If Len(LoginKontrol(Request.Form(fieldName))) > 4 Then SQLCumle1 = SQLCumle1 & fieldName &", "
		          If Len(LoginKontrol(Request.Form(fieldName))) > 4 Then SQLCumle2 = SQLCumle2 & "'"& BlogDateFunc( LoginKontrol(Request.Form(fieldName)) ) &"',"
		      Else
		          SQLCumle1 = SQLCumle1 & ""& fieldName &","
		          SQLCumle2 = SQLCumle2 & "'"& LoginKontrol(Request.Form(fieldName)) &"',"
		      End If
		  Next
		End If

		If CollectType = "UPDATE" Then 
	    For Each Item In Request.Form
	        fieldName = Item
	        fieldValue = Request.Form(Item)

	        If fieldName = "file" OR fieldName = "FAVICON" Then 

	        ElseIf fieldName = "PASSWORD" Then
	            tmp_pass = Trim(LoginKontrol(Request.Form(fieldName)))
	            If Len(tmp_pass) > 4 Then  SQLCumle2 = SQLCumle2 & "PASSWORD = '"& MD5( tmp_pass ) &"',"
	        ElseIf fieldName = "KIMLIK_NO" Then
	            If IsNumeric(LoginKontrol(Request.Form(fieldName))) Then SQLCumle2 = SQLCumle2 & "KIMLIK_NO='"& LoginKontrol(Request.Form(fieldName)) &"',"
		      ElseIf fieldName = "KURS_SURESI" Then
		          SQLCumle2 = SQLCumle2 & "'"& ParaTemizle( LoginKontrol(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "UCRET" OR fieldName = "KURS_UCRETI" OR fieldName = "DERS_UCRETI" Then
		          If Len(LoginKontrol(Request.Form(fieldName))) > 4 Then SQLCumle2 = SQLCumle2 & "'"& ParaTemizle( LoginKontrol(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "KURS_YILI" OR fieldName = "YIL"  Then
		          SQLCumle2 = SQLCumle2 & "'"& SQLDate( LoginKontrol(Request.Form(fieldName)) & "-01-01" ) &"',"
	        ElseIf fieldName = "DOGUM_TARIHI" Then
	            SQLCumle2 = SQLCumle2 & ""&fieldName&"='"& SQLDate( LoginKontrol(Request.Form(fieldName)) ) &"',"
		      ElseIf fieldName = "BITIS_TARIHI" OR fieldName = "BASLANGIC_TARIHI" OR fieldName = "GUNCELLENME_TARIHI" OR fieldName = "EKLENME_TARIHI" OR fieldName = "OGRENCI_KAYIT_BASLANGIC" OR fieldName = "OGRENCI_KAYIT_BITIS" Then
	            SQLCumle2 = SQLCumle2 & ""&fieldName&"='"& BlogDateFunc( LoginKontrol(Request.Form(fieldName)) ) &"',"
	        Else
	            SQLCumle2 = SQLCumle2 & ""& fieldName &"='"& LoginKontrol(Request.Form(fieldName)) &"',"
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
	'
	'----------------------------------------
	Public Property Get Data(vData)
		If IsNull(vData) Or IsEmpty(vData) Or Len(vData) < 0 Then 
			Data = ""
		Else 
			Data = Trim( LoginKontrol( Request(""&vData&"") ) )
			If IsNumeric(Data) Then Data = Cint(Data)
		End If
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
	' .CountRow("tbl_name", "id", "WHERE X='y'") return numeric
	'----------------------------------------
	Public Property Get CountRow(table, CountRowName, WhereCondition)
		if Len(WhereCondition) > 0 Then sql_add = " WHERE "& WhereCondition &""
		CountRow = ConnectionObj.Execute("SELECT COUNT("&CountRowName&") FROM "& table & sql_add &"")(0)
	End Property

	'----------------------------------------
	' .MaxID("tbl_name") return numeric
	'----------------------------------------
	Public Property Get MaxID(tableName)
		Set LatestID = ConnectionObj.Execute("SELECT MAX(ID) FROM "& tableName &"")
			MaxID = LatestID(0)
		LatestID.Close : Set LatestID = Nothing
	End Property

	'----------------------------------------
	' .TableExist("tbl_name") return boolean
	'----------------------------------------
	Public Property Get TableExist(tableName)
		Set CheckTableExist = ConnectionObj.Execute("SHOW TABLES LIKE '"&tableName&"'")
		If CheckTableExist.Eof Then 
			TableExist = False
		Else
			TableExist = True
		End If
		CheckTableExist.Close : Set CheckTableExist = Nothing
	End Property

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
	' .Debug = True / False
	'----------------------------------------
	Public Sub Go(url)
		Response.Redirect url
		Response.End()
	End Sub

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










End Class
%>