<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<%
	If Session("Username")= "" Then
		Response.Redirect("admin_login.asp")
	End If	
	If Session("book")= False Then
		If Session("Admin") = False Then
			Response.Redirect("admin_error.asp?type=5")
		End If	
	End If
%>

<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<%
		id	= Request.QueryString("id")	
		Set rsEdit = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM REGISTER WHERE ID="&id
		rsEdit.CursorType = 2
		rsEdit.LockType = 3
		rsEdit.Open strSQL, Conn
		
		Set rsCard = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM CARD WHERE CARD_ID='" & rsEdit("CARD_ID") & "'"
		rsCard.Open strSQL, Conn,3,1

		txtCardID = rsEdit("CARD_ID")
		txtBookID = rsEdit("BOOK_ID")
		txtClassID = rsCard("CLASS_ID")
		
		rsEdit.Close
		Set rsEdit = Nothing				
		rsCard.Close
		Set rsCard = Nothing						

		Set rsBreach = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT CARD_ID FROM BREACH WHERE CARD_ID='" & Trim(txtCardID) & "'"
		rsBreach.Open strSQL, Conn,3,1
		If Not rsBreach.Eof Then			
			rsBreach.Close
			Set rsBreach = Nothing
			Response.Redirect("admin_breach.asp")
		Else
			rsBreach.Close
			Set rsBreach = Nothing				
		End If
		
		Set rsDoneCheck = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT CARD_ID FROM BORROW WHERE CARD_ID='" & Trim(txtCardID) & "'"
		rsDoneCheck.Open strSQL, Conn,3,1
		If Not rsDoneCheck.Eof Then			
			rsDoneCheck.Close
			Set rsDoneCheck = Nothing
			Response.Redirect("admin_error.asp?type=11")
		Else
			rsDoneCheck.Close
			Set rsDoneCheck = Nothing			
		End If
		
		Set rsCheckBook = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT BOOK_ID, AMOUNT FROM BOOK WHERE BOOK_ID='" & Trim(txtBookID) & "'"
		rsCheckBook.Open strSQL, Conn,3,1
		If rsCheckBook("AMOUNT") = CheckCountBorrow(txtBookID) Then
			rsCheckBook.Close
			Set rsCheckBook = Nothing
			Response.Redirect("admin_error.asp?type=10")
		End If
		
		strSQL = "SELECT * FROM BORROW Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
		
		strSQL = "INSERT INTO BORROW(ID,CARD_ID,BOOK_ID,CLASS_ID,DATE_INFORM)Values("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCardID,",")
		strSQL = strSQL & CheckString(txtBookID,",") & CheckString(txtClassID,",")
		strSQL = strSQL & CheckString(Now(),")")
		Conn.Execute(strSQL)
		
		strSQL = "SELECT * FROM TEMP_BORROW Order by ID Desc"		
		txtID = GetID(strSQL,Conn)
		
		strSQL = "INSERT INTO TEMP_BORROW(ID,CARD_ID,BOOK_ID,CLASS_ID,DATE_INFORM)Values("
		strSQL = strSQL & CheckString(txtID,",") & CheckString(txtCardID,",")
		strSQL = strSQL & CheckString(txtBookID,",") & CheckString(txtClassID,",")
		strSQL = strSQL & CheckString(Now(),")")
		Conn.Execute(strSQL)
		
		Set rsRegis = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM REGISTER"
		rsRegis.CursorType = 2
		rsRegis.LockType = 3
		rsRegis.Open strSQL, Conn
		Do While Not rsRegis.Eof
			If rsRegis("CARD_ID") = txtCardID Then
				rsRegis.Delete
			End If	
		rsRegis.MoveNext
		Loop

		rsRegis.Close
		Set rsRegis = Nothing				
		Conn.Close
		Set Conn = Nothing
		
		Response.Redirect("admin_done.asp?page=admin_listborrow.asp")
%>					
