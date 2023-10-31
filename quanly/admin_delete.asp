<!--#INCLUDE FILE="../include/inc_connect.asp" -->
<%
	strCategory = Request.QueryString("category")
	id = Request.QueryString("id")
	
	If strCategory = "news" then	
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM NEWS WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL) 
		Next
		page="admin_listnews.asp"
	ElseIf strCategory = "book" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM BOOK WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL) 
		Next
		page="admin_listbook.asp"	
	ElseIf strCategory = "soft" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM SOFTWARE WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL) 
		Next
		page="admin_listsoft.asp"			
	ElseIf strCategory = "document" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM DOCUMENT WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL) 
		Next
		page="admin_listdoc.asp"									
	ElseIf strCategory = "relax" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM RELAX WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL) 
		Next
		page = "admin_listrelax.asp"
	ElseIf strCategory = "vote" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM QUESTION WHERE ID ="& Request.Form("Mid")(i)
			strSQL1 = "DELETE FROM VOTE WHERE CATEGORY_ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL)
			Conn.EXECUTE(strSQL1) 
		Next
		page = "admin_listvote.asp"		
	ElseIf strCategory = "notice" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM NOTICE WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL) 
		Next
		page = "admin_listnotice.asp"				
	ElseIf strCategory = "user" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM USER WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL) 
		Next
		page = "admin_listuser.asp"				
	ElseIf strCategory = "schedule" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM SCHEDULE WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL)
		Next	
		page = "admin_listschedule.asp"	
	ElseIf strCategory = "card" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM CARD WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL)
		Next	
		page = "admin_listcard.asp"
	ElseIf strCategory = "borrow" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM BORROW WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL)
		Next	
		page = "admin_listborrow.asp"						
	ElseIf strCategory = "breach" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM BREACH WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL)
		Next	
		page = "admin_listbreach.asp"
	ElseIf strCategory = "register" then
		For i = 1 To Request.Form("Mid").Count
			strSQL = "DELETE FROM REGISTER WHERE ID ="& Request.Form("Mid")(i)
			Conn.EXECUTE(strSQL)
		Next	
		page = "admin_listregister.asp"
	ElseIf strCategory = "contact" then
		If id = "" Then
			For i = 1 To Request.Form("Mid").Count
				strSQL = "DELETE * FROM CONTACT WHERE ID="& Request.Form("Mid")(i)
				Conn.EXECUTE(strSQL)
			Next
		Else
			strSQL = "DELETE * FROM CONTACT WHERE ID="& id
			Conn.EXECUTE(strSQL)		
		End If		
		page="admin_listcontact.asp"
	ElseIf strCategory = "newscategory" then
		strSQL = "DELETE * FROM CATEGORY_NEWS WHERE ID="&id
		Conn.EXECUTE(strSQL) 
		page="admin_newscategory.asp"
	ElseIf strCategory = "doccategory" then
		strSQL = "DELETE * FROM CATEGORY_DOCUMENT WHERE ID="&id						
		Conn.EXECUTE(strSQL) 		
		page="admin_doccategory.asp"			
	ElseIf strCategory = "bookcategory" then
		strSQL = "DELETE * FROM CATEGORY_BOOK WHERE ID="&id
		Conn.EXECUTE(strSQL) 		
		page="admin_bookcategory.asp"	
	ElseIf strCategory = "cardcategory" then
		strSQL = "DELETE * FROM CATEGORY_CARD WHERE ID="&id
		Conn.EXECUTE(strSQL) 		
		page="admin_cardcategory.asp"	
	ElseIf strCategory = "publisher" then
		strSQL = "DELETE * FROM PUBLISHER WHERE ID="&id
		Conn.EXECUTE(strSQL) 		
		page="admin_publisher.asp"			
	ElseIf strCategory = "relaxcategory" then
		strSQL = "DELETE * FROM CATEGORY_RELAX WHERE ID="&id
		Conn.EXECUTE(strSQL)
		page="admin_relaxcategory.asp"	
	ElseIf strCategory = "softcategory" then
		strSQL = "DELETE * FROM CATEGORY_SOFT WHERE ID="&id
		Conn.EXECUTE(strSQL) 		
		page="admin_softcategory.asp"
	ElseIf strCategory = "language" then
		strSQL = "DELETE * FROM LANGUAGE WHERE ID="&id
		Conn.EXECUTE(strSQL)
		page="admin_language.asp"			
	ElseIf strCategory = "usergroup" then
		strSQL = "DELETE * FROM USERGROUP WHERE ID="&id
		Conn.EXECUTE(strSQL)
		page="admin_usergroup.asp"		
	ElseIf strCategory = "ologycategory" then
		strSQL = "DELETE * FROM OLOGY WHERE ID="&id
		Conn.EXECUTE(strSQL)
		page="admin_ologycategory.asp"
	ElseIf strCategory = "classcategory" then
		strSQL = "DELETE * FROM CLASS WHERE ID="&id
		Conn.EXECUTE(strSQL)
		page="admin_classcategory.asp"			
	ElseIf strCategory = "genrecategory" then
		strSQL = "DELETE * FROM CATEGORY_GENRE WHERE ID="&id
		Conn.EXECUTE(strSQL)
		page="admin_genrecategory.asp"		
	Else
	End If
	Conn.Close
	Set Conn = Nothing
	Response.Redirect(page)
%>