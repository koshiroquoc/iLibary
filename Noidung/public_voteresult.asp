<%Session.CodePage = 65001%>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Kết quả</title>
</head>
<%
	
	txtQuestionID = Request.Form("txtQuestionID")
	If txtQuestionID <> "" Then
		strSQL = "SELECT * FROM VOTE WHERE ID="& txtQuestionID
		Set rsVote = Server.CreateObject("ADODB.Recordset")
		rsVote.CursorType = 2
		rsVote.LockType = 3
		rsVote.Open strSQL,Conn
		value = rsVote.Fields("VALUE") + 1
		If Not rsVote.Eof Then
			rsVote.Fields("VALUE") = value
			rsVote.Update	
		End If
	End If	
		
	strSQL = "SELECT * FROM VOTE"
	Set rsCount = Server.CreateObject("ADODB.Recordset")
	rsCount.Open strSQL,Conn,3,1
	iSum = CountVote(rsCount)
%>
<link rel="stylesheet" type="text/css" href="../css/public.css">

<body leftmargin="4" topmargin="4">

</body>

</html>
<table border="1" width="230" id="table1" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC" height="104">
	<tr>
		<td>
<table border="0" width="230" id="table2" cellspacing="0" cellpadding="0" height="76">
	<tr>
		<td align="right" colspan="3" height="30">
		<p align="center" style="margin-top: 0; margin-bottom: 0"><b>
		<font color="#0000FF">KẾT QUẢ</font></b></td>
	</tr>
	<tr>
		<td width="100%" align="right" colspan="3">
		</td>
	</tr>
	<%
		strSQL = "SELECT * FROM VOTE WHERE CATEGORY_ID=" & txtQuestionID
		Set rsDisplay = Server.CreateObject("ADODB.Recordset")
		rsDisplay.Open strSQL,Conn,3,1
		Do While Not rsDisplay.Eof 
		iVote = Cint((rsDisplay("VALUE")/iSum)*100)
		If iVote = 0 Then 
			iVote = 1
		End If	
	%>	
	<tr>
		<td width="5" align="right">&nbsp;</td>
		<td width="<%=Len(rsDisplay("NAME"))*8%>"><%=rsDisplay("NAME")%></td>
		<td width="125">
		<p style="margin-top: 3px; margin-bottom: 3px">
		<img border="0" src="../images/vote.gif" width="<%=iVote%>" height="15"></td>
	</tr>
	<%
		rsDisplay.MoveNext
		Loop
		Conn.Close
		Set Conn = Nothing
	%>
	<tr>
		<td align="right" colspan="3" height="9"></td>
	</tr>
	<tr>
		<td align="right" colspan="3" height="15">
		<p align="center"><b><a href="JavaScript:closepopup();">Đóng</a></b></td>
	</tr>
	<tr>
		<td align="right" colspan="3" height="1"></td>
	</tr>
	</table>

		</td>
	</tr>
</table>
<SCRIPT LANGUAGE="JavaScript">
function closepopup()
{
closewindow = window.close(),"closewindow"
}

</SCRIPT>