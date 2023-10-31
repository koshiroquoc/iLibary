<%	@LANGUAGE="VBSCRIPT" %>
<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<!-- #INCLUDE FILE="../include/inc_hexpass.asp" -->
<%
	txtUsername = AddSlash(Request.Form("txtUsername"))
	If txtUsername ="" Then	
		Response.Redirect("admin_login.asp?error=1")
	End If	
	txtPassword = AddSlash(Request.Form("txtPassword"))
	txtPassword = HashEncode(txtPassword)

	strSQL = "SELECT * FROM USER WHERE USERNAME ='" & txtUsername & "'" 
	Set rsLogin = Server.CreateObject("ADODB.Recordset")
	rsLogin.Open strSQL,Conn,3,3


	If Not rsLogin.EOF Then
		If txtPassword = rsLogin("PASSWORD") Then
			page = SetPower(txtUsername)
			If page = "" Then
				Response.Redirect("admin_login.asp?error=2")
			Else
				Response.Redirect(page)
			End If	
		Else
			Response.Redirect("admin_login.asp?error=3")
		End If	
	Else
		Response.Redirect("admin_login.asp?error=4")
	End If	

%>