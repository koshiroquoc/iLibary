<!--#INCLUDE FILE="../include/inc_connect.asp" -->
<%
	Set rsCheckBreach = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM BREACH"
	rsCheckBreach.CursorType = 2
	rsCheckBreach.LockType = 3
	rsCheckBreach.Open strSQL, Conn
	Do While Not rsCheckBreach.Eof
		If Now() - rsCheckBreach("DATE_INFORM") >=7 Then
			rsCheckBreach.Delete
		End If
	rsCheckBreach.MoveNext
	Loop		
	rsCheckBreach.Close
	Set rsCheckBreach = Nothing
	
	Set rsCheckRegister = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM REGISTER"
	rsCheckRegister.CursorType = 2
	rsCheckRegister.LockType = 3
	rsCheckRegister.Open strSQL, Conn
	Do While Not rsCheckRegister.Eof
		If Now() - rsCheckRegister("DATE_INFORM") >=3 Then
			rsCheckRegister.Delete
		End If
	rsCheckRegister.MoveNext
	Loop		
	rsCheckRegister.Close
	Set rsCheckRegister = Nothing
	Conn.Close
	Set Conn = Nothing
%>