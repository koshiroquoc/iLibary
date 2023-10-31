<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Trang chủ</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
</head>

<body>
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse"  width="155" id="table1" bgcolor="#FFFFFF">
<tr>
<td height="21" width="12" style="border-left: 1px solid #999999; border-top: 1px solid #999999" bgcolor="#800000">
<font color="#FFFFFF"><b>..::</b></font></td>
<td height="21" width="141" style="border-right: 1px solid #999999; border-top: 1px solid #999999" class="txt_titlemenu" bgcolor="#800000">
<font color="#FFFFFF">
<b>&nbsp;THĂM DÒ Ý KIẾN</b></font></td>
</tr>
<tr>
<td width="100%" colspan="2">
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse; border-left: 1px solid #999999; border-right: 1px solid #999999" bordercolor="#111111" width="100%" id="table2">
<tr>
<form method="POST" target="poll" name="frmVote" onSubmit="window.open('', 'poll', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=0,width=250,height=180')" action="public_voteresult.asp">
<td style="border:1px solid #999999; " valign="top" height="130">
<%
	strSQL = "SELECT * FROM QUESTION WHERE STATUS=1"
	Set rsQuestion = Server.CreateObject("ADODB.Recordset")
	rsQuestion.Open strSQL,Conn,3,1
	If Not rsQuestion.Eof Then
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="2" align="center"><b><%=rsQuestion("NAME")%></b></td>
	</tr>
	<tr>
		<td colspan="2">
		<img border="0" src="../images/spacer.gif" width="1" height="3"></td>
	</tr>
	<%
		strSQL = "SELECT * FROM VOTE WHERE CATEGORY_ID=" & rsQuestion("ID")
		Set rsVote = Server.CreateObject("ADODB.Recordset")
		rsVote.Open strSQL,Conn,3,1
		Do While Not rsVote.Eof
	%>	
	<tr>
		<td width="16%"><input type="radio" value="<%=rsVote("ID")%>" name="Choose"></td>
		<td width="84%"><%=rsVote("NAME")%></td>
	</tr>
	<%
		rsVote.MoveNext
		Loop
	%>
	<tr>
		<td colspan="2">
		<img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td colspan="2">
		<p align="center">		
		<input type="submit" value=" Chọn " name="B1" class="input_button"></td>
	</tr>
	<tr>
		<td colspan="2" height="7">
		</td>
	</tr>
</table>
<input type="hidden" name="txtQuestionID" value="<%=rsQuestion("ID")%>">
<%
	End If
	Conn.Close
	Set Conn = Nothing
%>
	</td>
</form>
</tr>
</table>
</td>
</tr>
</table>
</body>

</html>