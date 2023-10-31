<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
</head>

<body>
<%
	strSQL = "SELECT * FROM NEWS Order by DATE_INFORM Desc"
	Set rsNews = Server.CreateObject("ADODB.Recordset")
	rsNews.Open strSQL,Conn,3,1
%>
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<%
		If rsNews.Eof Then
	%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>	
	<tr>
		<td colspan="2">
		<p align="center">&nbsp; Xin lổi, tin tức chưa được cập nhật!</td>
	</tr>
	<%
		Else
		iCount = 1
		Do While Not rsNews.Eof And iCount <=4
	%>
	<tr>
		<td class="txt_title" background="../images/bg_bar.gif" height="25" colspan="2">
		<p style="margin-left: 4px">
		<a href="default.asp?name=newsdetail&id=<%=rsNews("ID")%>"><b><%=uCase(rsNews("TITLE"))%></b></a>	
		</td>
	</tr>
	<tr>
		<td colspan="2">
		<p align="justify">
		<img border="0" src="../images/spacer.gif" width="1" height="3"></td>
	</tr>	
	<%
		If rsNews("IMAGE") <> "" Then
	%>
	<tr>
		<td colspan="2"><p align="justify" style="margin-right: 3px">
		<img border="0" align="left" width="130" height="89" src="<%=rsNews("IMAGE")%>" style="border-style: outset; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" hspace="6">
		<%=rsNews("SUMMARY")%></td>
	</tr>
	<%
		Else
	%>		
	<tr>
		<td colspan="2"><p align="justify" style="margin-right: 3px">		
		<%=rsNews("SUMMARY")%></td>
	</tr>
	<%
		End If
	%>	
	<tr>
		<td colspan="2">
		<p align="justify">
		<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
	</tr>	
	<%
		iCount = iCount + 1
		rsNews.MoveNext
		Loop
		End If
	%>
	<tr>
		<td colspan="2">
		<p align="center" style="margin-top: 10px">
		<img border="0" src="../images/line.gif" width="405" height="5"></td>
	</tr>
	<%
		strSQL = "SELECT * FROM NEWS Order by DATE_INFORM Desc"
		Set rsLastNews = Server.CreateObject("ADODB.Recordset")
		rsLastNews.Open strSQL,Conn,3,1
		If rsLastNews.RecordCount > iCount Then
	%>
	<tr>
		<td colspan="2"><b>Các tin khác:</b></td>
	</tr>
	<%
		iLast = 1
		Do While Not rsLastNews.Eof And iLast <= 14
		If iLast >= iCount Then
	%>
	<tr>
		<td width="14">
		<p align="right">
		<img border="0" src="../images/blackArrow_right.gif" width="7" height="5"></td>
		<td width="422">
		<a href="default.asp?name=newsdetail&id=<%=rsLastNews("ID")%>"><%=rsLastNews("TITLE")%></a></td>
	</tr>
	<%
		End If
		iLast = iLast + 1
		rsLastNews.MoveNext
		Loop
		End If
	%>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
</table>

</div>

</body>

</html>