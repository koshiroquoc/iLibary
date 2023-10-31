<!-- #INCLUDE FILE="../include/inc_parameter.asp" -->
<%
'Ham kiem tra
Function AddSlash(s)
	Dim StrKq
	Dim i,j
	Dim Lens, ch	
	Lens = Len(s)
	StrKq = ""
	For i = 1 to Lens
		ch = Mid(s,i,1)
		If (ch = "'") Then
			StrKq = StrKq + "'"
		End If
		If (ch = Chrw(34)) Then
			StrKq = StrKq + Chrw(34)
		End If
		StrKq = StrKq + ch
	Next
	AddSlash = StrKq
End Function

Function NgayVN_Text(datTheDay)
	intThisYear = Year(datTheDay)
	intThisMonth = Month(datTheDay)
	intThisDay = Day (datTheDay)	
	if intThisMonth < 10 then
		intThisMonth = "0" & intThisMonth
	end if
	if intThisDay < 10 then
		intThisDay = "0" & intThisDay
	end if	
	NgayVN_Text=intThisDay &" tháng " & intThisMonth&" n&#259;m "&intThisYear
End Function

Function NgayVN(datTheDay)
	intThisYear = Year(datTheDay)
	intThisMonth = Month(datTheDay)
	intThisDay = Day (datTheDay)	
	if intThisMonth < 10 then
		intThisMonth = "0" & intThisMonth
	end if
	if intThisDay < 10 then
		intThisDay = "0" & intThisDay
	end if	
	NgayVN=intThisDay&"/"&intThisMonth&"/"&intThisYear
End Function

Function GetID(strSQL,Conn)
	Set rsGet = Server.CreateObject("ADODB.Recordset")
	rsGet.CursorType = 2
	rsGet.LockType = 3
	rsGet.Open strSQL, Conn
	If rsGet.EOF Then
		ID = 1
	Else
		ID = rsGet.Fields("ID") + 1
	End If	
	rsGet.Close
	Set rsGet = Nothing
	GetID = ID
	Exit Function
End Function

Function CheckString(s,endChar)    
	pos=InStr(s,"'")
	While pos>0
		s=Mid(s,1,pos)&"'"&Mid(s,pos+1)
		pos=InStr(pos+2,s,"'")
	Wend
	CheckString="'"&s&"'"&endChar
End Function
Function AddCom(s)
	Dim StrKq
	Dim i,j
	Dim Lens, ch
	
	s = Replace(s,">","&gt;")
	s = Replace(s,"<","&lt;")
	AddCom = s
End Function

Sub ListCombo(inputQuery,index)
	SET rstemp = Conn.Execute(inputQuery)
	Do While Not rstemp.eof
		If index = rstemp(1) then
		%>
			<option selected value="<%=rstemp(1)%>"><%=rstemp(0)%></option>
		<% else %>
			<option value="<%=rstemp(1)%>"><%=rstemp(0)%></option>
		<%
		End if
	rstemp.MoveNext
	Loop 
	rstemp.Close
	Set rstemp = Nothing
End Sub


Sub ListComboCARD(inputQuery,index)
	SET rstemp = Conn.Execute(inputQuery)
	Do While Not rstemp.eof
		If index = rstemp(1) then
		%>
			<option selected value="<%=rstemp(0)%>"><%=rstemp(1)%></option>
		<% else %>
			<option value="<%=rstemp(0)%>"><%=rstemp(1)%></option>
		<%
		End if
	rstemp.MoveNext
	Loop 
	rstemp.Close
	Set rstemp = Nothing
End Sub


Sub ListComboCARD1(inputQuery,index)
	SET rstemp = Conn.Execute(inputQuery)
	Do While Not rstemp.eof
		If index = rstemp(1) then
		%>
			<option selected value="<%=rstemp(0)%>"><%=rstemp(0)%></option>
		<% else %>
			<option value="<%=rstemp(0)%>"><%=rstemp(0)%></option>
		<%
		End if
	rstemp.MoveNext
	Loop 
	rstemp.Close
	Set rstemp = Nothing
End Sub


Sub ListNumber(id,ic,index)
	For i = id to ic
		If i = index then
		%>
			<option value="<%=i%>" selected><%=i%></option>
		<% else %>
			<option value="<%=i%>"><%=i%></option>
		<%
		End if
	Next 
End Sub

Sub UpdateField(strSQL)
	Set rsSet = Server.CreateObject("ADODB.Recordset")
	rsSet.CursorType = 2
	rsSet.LockType = 3
	rsSet.Open strSQL, Conn
	Do While Not rsSet.EOF
		If rsSet("STATUS") = 1 Then
			rsSet("STATUS") = 0
			rsSet.Update
		End If	
	rsSet.MoveNext
	Loop
	rsSet.Close
	Set rsSet = Nothing
End Sub

Function ZenCardID(strCATACARDID)
	strClassID = strCATACARDID
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select CARD_ID From CARD Where Left(CARD_ID,2)='" & strClassID & "'"
	strSQL = strSQL & "Order By cInt(Right(CARD_ID,5)) Desc"
	Set rstemp = Conn.Execute(strSQL)
	If rstemp.Eof Then
		ZenCardID = strClassID & "00001"
	Else
		Index = Cint(Right(rstemp("CARD_ID"),5))
		Index = Index + 1
		If Index < 10 Then
			ZenCardID = strClassID & "0000" & Index
		ElseIf Index < 100 Then
			ZenCardID = strClassID & "000" & Index
		ElseIf Index < 1000 Then
			ZenCardID = strClassID & "00" & Index
		ElseIf Index < 10000 Then
			ZenCardID = strClassID & "0" & Index
		ElseIf Index < 100000 Then
			ZenCardID = strClassID & Index			
		End If	
	End If	 		
End Function

Function ImportExcel(strFileName,Conn)
	strConnection = "DBQ=" & strFileName & "; DRIVER={Microsoft Excel Driver (*.xls)};"
	Set cnExcel = Server.CreateObject("ADODB.Connection")
	Set rsExcel = Server.CreateObject("ADODB.Recordset")
	cnExcel.Open strConnection
	
	strSQLExcel="SELECT * FROM DANHSACH;"

	rsExcel.Open strSQLExcel, cnExcel, 3, 1
	
	Do While Not rsExcel.Eof
		strSQL = "INSERT INTO EXCEL_IMPORT(FIRSTNAME,LASTNAME,BIRTHDAY) VALUES("
		strSQL = strSQL & CheckString(rsExcel("FIRSTNAME"),",")
		strSQL = strSQL & CheckString(rsExcel("LASTNAME"),",")
		strSQL = strSQL & CheckString(rsExcel("BIRTHDAY"),")")
		Conn.Execute strSQL
	rsExcel.MoveNext
	Loop

	cnExcel.Close
	Set cnExcel = Nothing
End Function
	
Function CountBook()
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select BOOK_ID From BOOK"
	Set rstemp = Conn.Execute(strSQL)
	CountBook = rstemp.RecordCount
End Function

Function CountCard()
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select CARD_ID From CARD"
	Set rstemp = Conn.Execute(strSQL)
	CountCard = rstemp.RecordCount
End Function

Function CountClassCard(strClassID)
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select CLASS_ID From CARD WHERE CLASS_ID ='" & strClassID & "'"
	Set rstemp = Conn.Execute(strSQL)
	CountClassCard = rstemp.RecordCount
End Function

Function CountBreaching()
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select * From BORROW WHERE NOW()-DATE_INFORM>7"
	Set rstemp = Conn.Execute(strSQL)
	CountBreaching = rstemp.RecordCount
End Function

Function CountBreachingClass(strClassID)
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select * From BORROW WHERE NOW()-DATE_INFORM>7"
	strSQL = strSQL & " AND CLASS_ID ='" & strClassID & "'"
	Set rstemp = Conn.Execute(strSQL)
	CountBreachingClass = rstemp.RecordCount
End Function

Function CountCateBook(strCateBook)
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select BOOK_ID From BOOK Where Left(BOOK_ID,3)='" & strCateBook & "'"
	Set rstemp = Conn.Execute(strSQL)
	If rstemp.Eof Then
		CountCateBook = 0
	Else	
		CountCateBook = rstemp.RecordCount
	End If	
End Function

Function CountBorrow()
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select BOOK_ID From BORROW"
	Set rstemp = Conn.Execute(strSQL)
	CountBorrow = rstemp.RecordCount
End Function

Function CheckCountBorrow(strBookID)
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select BOOK_ID From BORROW WHERE BOOK_ID ='" & strBookID & "'"
	Set rstemp = Conn.Execute(strSQL)
	CheckCountBorrow = rstemp.RecordCount
End Function

Function CountCardID(strCardID)
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select ID From REGISTER WHERE CARD_ID ='" & strCardID & "'"
	Set rstemp = Conn.Execute(strSQL)
	CountCardID = rstemp.RecordCount
End Function

Function CheckCardID(strCardID)
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select CARD_ID From CARD WHERE CARD_ID ='" & strCardID & "'"
	Set rstemp = Conn.Execute(strSQL)
	CheckCardID = rstemp.RecordCount
End Function

Function CountRegister(strCardID,strBookID)
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select ID From REGISTER WHERE CARD_ID ='" & strCardID & "'"
	strSQL = strSQL & " AND BOOK_ID='" & strBookID & "'"
	Set rstemp = Conn.Execute(strSQL)
	CountRegister = rstemp.RecordCount
End Function

Function CountCateBookBorrow(strCateBook)
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select BOOK_ID From BORROW Where Left(BOOK_ID,3) ='" & strCateBook & "'"
	Set rstemp = Conn.Execute(strSQL)
	CountCateBookBorrow = rstemp.RecordCount
End Function

Function CountSumBorrow()
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select BOOK_ID From TEMP_BORROW"
	Set rstemp = Conn.Execute(strSQL)
	CountSumBorrow = rstemp.RecordCount
End Function

Function CountSumReturn()
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select BOOK_ID From TEMP_RETURN"
	Set rstemp = Conn.Execute(strSQL)
	CountSumReturn = rstemp.RecordCount
End Function

Function ZenBookID(strCategoryID)
	Dim Index
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select BOOK_ID From BOOK Where Left(BOOK_ID,3)='" & strCategoryID & "'"
	strSQL = strSQL & "Order By Right(BOOK_ID,3) Desc"
	Set rstemp = Conn.Execute(strSQL)
	If rstemp.Eof Then
		ZenBookID = strCategoryID & "001"
	Else	
		Index = Cint(Right(rstemp("BOOK_ID"),3))
		Index = Index + 1
		If Index < 10 Then
			ZenBookID = strCategoryID & "00" & Index
		ElseIf Index < 100 Then
			ZenBookID = strCategoryID & "0" & Index
		ElseIf Index < 1000 Then
			ZenBookID = strCategoryID & Index
		End If	 		
	End If	
	rstemp.Close
	Set rstemp = Nothing
End Function

Function ZenDocID(strCategoryID)
	Dim Index
	Set rstemp = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select DOCUMENT_ID From DOCUMENT Where Left(DOCUMENT_ID,3)='" & strCategoryID & "'"
	strSQL = strSQL & "Order By Right(DOCUMENT_ID,3) Desc"
	Set rstemp = Conn.Execute(strSQL)
	If rstemp.Eof Then
		ZenDocID = strCategoryID & "001"
	Else	
		Index = Cint(Right(rstemp("DOCUMENT_ID"),3))
		Index = Index + 1
		If Index < 10 Then
			ZenDocID = strCategoryID & "00" & Index
		ElseIf Index < 100 Then
			ZenDocID = strCategoryID & "0" & Index
		ElseIf Index < 1000 Then
			ZenDocID = strCategoryID & Index
		End If	 		
	End If	
	rstemp.Close
	Set rstemp = Nothing
End Function

Function SetPower(strUserName)
	Session("Username") = strUsername
	Set rsUser = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select * From USER Where USERNAME='" & strUserName & "'"
	Set rsUser = Conn.Execute(strSQL)
	If Not rsUser.Eof Then
		If rsUser("LEVEL")=1 Then
			Session("Admin") = True
			Session("Mod") = True
			Response.Cookies("Admin")("Login") = 1
			Session.Timeout = 60
			SetPower = "admin_default.asp"
		Else	
			Set rsManager = Server.CreateObject("ADODB.Recordset")
			Set rsUserFunction = Server.CreateObject("ADODB.Recordset")
			strSQL = "Select * From MODULE Where USERNAME ='" & rsUser("USERNAME") & "'"
			Set rsManager = Conn.Execute(strSQL)
			If Not rsManager.Eof Then
				Do While Not rsManager.Eof 
					strSQL = "Select * From FUNCTION Where ID =" & rsManager("FUNCTION_ID")
					Set rsUserFunction = Conn.Execute(strSQL)
					If Not rsUserFunction.Eof Then
						checkuser = 1
						Session(rsUserFunction("SHORT_NAME")) = True
					End If	
				rsManager.MoveNext
				Loop
				If checkuser = 1 Then
					Session.Timeout = 60
					Session("Mod") = True
					If Session("library") = True Then
						SetPower = "admin_library.asp"
					Else
						SetPower = "admin_default.asp"	
					End If	
				End If	
			Else
				Set rsGroup = Server.CreateObject("ADODB.Recordset")
				Set rsGroupFunction = Server.CreateObject("ADODB.Recordset")			
				strSQL = "Select * From MODULE Where GROUP_ID=" & rsUser("GROUP_ID")
				Set rsGroup = Conn.Execute(strSQL)
				If Not rsGroup.Eof Then
					Do While Not rsGroup.Eof
						strSQL = "Select * From FUNCTION WHERE ID=" & rsGroup("FUNCTION_ID")
						Set rsGroupFunction = Conn.Execute(strSQL)
						If Not rsGroupFunction.Eof Then
							checkgroup = 1
							Session(rsGroupFunction("SHORT_NAME")) = True
						End If	
					rsGroup.MoveNext
					Loop
					If checkgroup = 1 Then
						Session.Timeout = 60							
						Session("Mod") = True
						If Session("library") = True Then
							SetPower = "admin_library.asp"
						Else
							SetPower = "admin_default.asp"	
						End If	
					End If	
				End If
			End If
		End If	
	End If	
	Conn.Close
	Set Conn = Nothing
End Function

Function SearchNoneCategory (txtSearchKey,txtBookName,txtSummary,txtAuthorName)
	If txtBookName = False Then
		If txtSummary = False Then
			If txtAuthorName = False Then
				strSQL = "SELECT * FROM BOOK WHERE NAME LIKE '%" & txtSearchKey & "%'"
			Else
				strSQL = "SELECT * FROM BOOK WHERE AUTHOR LIKE '%" & txtSearchKey & "%'"
			End If	
		Else
			If txtAuthorName = False Then
				strSQL = "SELECT * FROM BOOK WHERE SUMMARY LIKE '%" & txtSearchKey & "%'"
			Else
				strSQL = "SELECT * FROM BOOK WHERE SUMMARY LIKE '%" & txtSearchKey & "%'"
				strSQL = strSQL & " OR AUTHOR LIKE '%" & txtSearchKey & "%'"
			End If				
		End If		
	Else
		If txtSummary = False Then
			If txtAuthorName = False Then
				strSQL = "SELECT * FROM BOOK WHERE NAME LIKE '%" & txtSearchKey & "%'"
			Else
				strSQL = "SELECT * FROM BOOK WHERE NAME LIKE '%" & txtSearchKey & "%'"
				strSQL = strSQL & " OR AUTHOR LIKE '%" & txtSearchKey & "%'"
			End If
		Else
			If txtAuthorName = False Then
				strSQL = "SELECT * FROM BOOK WHERE NAME LIKE '%" & txtSearchKey & "%'"
				strSQL = strSQL & " OR SUMMARY LIKE '%" & txtSearchKey & "%'"
			Else							
				strSQL = "SELECT * FROM BOOK WHERE NAME LIKE '%" & txtSearchKey & "%'"
				strSQL = strSQL & " OR SUMMARY LIKE '%" & txtSearchKey & "%'"				
				strSQL = strSQL & " OR AUTHOR LIKE '%" & txtSearchKey & "%'"
			End If	
		End If	
	End If	
	SearchNoneCategory = strSQL
End Function

Sub InsertSQL(strSQLInsert)
	Set rsCheck = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select * From TEMP"
	rsCheck.CursorType = 2
	rsCheck.LockType = 3
	rsCheck.Open strSQL, Conn
	If Not rsCheck.Eof Then	
		rsCheck.Fields("SQL") = strSQLInsert			
'		rsCheck.Update	
	Else
		strSQL = "SELECT * FROM TEMP Order by ID Desc"		
		txtID = GetID(strSQL,Conn)			
		strSQL = "INSERT INTO TEMP(ID,SQL)Values("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(strSQLInsert,")")
		Conn.Execute strSQL
	End If
End Sub

Function LoadSQL()	
	Set rsLoad = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select * From TEMP"
	Set rsLoad = Conn.Execute(strSQL)
	If Not rsLoad.Eof Then
		strSQL = rsLoad("SQL")
	End If
	LoadSQL = strSQL
End Function

Sub InsertAdminSQL(strSQLInsert)
	Set rsCheck = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select * From TEMP_LIST"
	rsCheck.CursorType = 2
	rsCheck.LockType = 3
	rsCheck.Open strSQL, Conn
	If Not rsCheck.Eof Then	
		rsCheck.Fields("SQL") = strSQLInsert			
		rsCheck.Update	
	Else
		strSQL = "SELECT * FROM TEMP_LIST Order by ID Desc"		
		txtID = GetID(strSQL,Conn)			
		strSQL = "INSERT INTO TEMP_LIST(ID,SQL)Values("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(strSQLInsert,")")
		Conn.Execute strSQL
	End If
End Sub

Function LoadAdminSQL()	
	Set rsLoad = Server.CreateObject("ADODB.Recordset")
	strSQL = "Select * From TEMP_LIST"
	Set rsLoad = Conn.Execute(strSQL)
	If Not rsLoad.Eof Then
		strSQL = rsLoad("SQL")
	End If
	LoadAdminSQL = strSQL
End Function

Function CountVote(rstemp)
	iSum = 0
	Do While Not rstemp.Eof
		iSum = iSum + rstemp("VALUE")
	rstemp.MoveNext
	Loop
	CountVote= iSum
End Function

%>
<Script language="JavaScript">
function openWindow(url) {
  popupWin = window.open(url,'new_page','width=230,height=140,left=200')
}
function openWindow2(url) {
  popupWin = window.open(url,'new_page','width=400,height=150,left=200')
}
function openWindowPrint(url) {
  popupWin = window.open(url,'new_page','width=600,height=540,left=150,scrollbars=yes,top=0')
}
function openWindow3(url) {
  popupWin = window.open(url,'new_page','width=445,height=235,left=150,top=20')
}
function InsertStr(strValue,anh)
{
	window.opener.document.form1[anh].value=strValue;
}
function doSubmit(url)
{
	document.frmList.action = url;
	document.frmList.submit();
}
function cboChange(strComboName){
	document.frmFilter.txtCategoryFilter.value = document.frmList[strComboName].value;
	document.frmFilter.submit();
}
function cboChangeClass(strComboName){
	document.frmFilter.txtCategoryFilter.value = document.frmList[strComboName].value;
	document.frmFilter.txtClassID.value = document.frmList[strComboName].value;
	document.frmFilter.submit();
}
</Script>