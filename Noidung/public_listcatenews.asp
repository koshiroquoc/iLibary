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
	id	= Request.QueryString("id")
	strSQL = "SELECT * FROM CATEGORY_NEWS WHERE ID="&id
	Set rsCategory = Server.CreateObject("ADODB.Recordset")
	rsCategory.Open strSQL,Conn,3,1
%>
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td>
		<table border="0" width="100%" id="table3" cellspacing="0" cellpadding="0">
			<tr>
				<td background="../images/bg_bar.gif" height="24">
				<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
					<tr>
						<td width="23">
						<p align="center">
						<img border="0" src="../images/Pic/bullet.gif" width="10" height="10" align="middle"></td>
						<td class="txt_title"><b><%=uCase(rsCategory("NAME"))%></b></td>
					</tr>
				</table>
				</td>
			</tr>
			<%
				strSQL = "SELECT * FROM NEWS WHERE CATEGORY_ID="&rsCategory("ID")
				strSQL = strSQL & " Order By DATE_INFORM Desc"
				Set rsNews = Server.CreateObject("ADODB.Recordset")
				rsNews.Open strSQL,Conn,3,1
				iCount = 1
				Do While Not rsNews.Eof and iCount <=5
			%>
			<tr>
				<td width="100%">
				<p style="margin-top: 4px; margin-bottom: 4px"><a href="default.asp?name=newsdetail&id=<%=rsNews("ID")%>"><b><%=rsNews("TITLE")%></a></b></td>
			</tr>
			<%
				If Not rsNews("IMAGE") = "" Then
			%>
			<tr>
				<td width="100%" valign="top">
				<img border="0" align="left" width="130" height="89" src="<%=rsNews("IMAGE")%>" style="border-style: outset; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" hspace="6">
				<p align="justify"><span class="Head"><%=rsNews("SUMMARY")%></span>
				</td>
			</tr>
			<%
				Else
			%>
			<tr>	
				<td width="91%" valign="top"><p align="justify"><span class="Head">
				<%=rsNews("SUMMARY")%>
				</span></td>
			</tr>
			<%
				End If
				If iCount <> rsNews.RecordCount	Then			
			%>
			<tr>	
				<td width="91%" align="center">
				<p style="margin-top: 4px; margin-bottom:4px">
				<img border="0" src="../images/line.gif" width="373" height="5"></td>
			</tr>			
			<%
				End If
				iCount = iCount + 1
				rsNews.MoveNext
				Loop
			%>	
			<tr>
				<td width="100%">
				<table border="0" width="100%" id="table5" cellspacing="0" cellpadding="0">
					<%
						strSQL = "SELECT * FROM NEWS WHERE CATEGORY_ID =" & rsCategory("ID")
						strSQL = strSQL & " Order By DATE_INFORM Desc"
						Set rsLastNews = Server.CreateObject("ADODB.Recordset")
						rsLastNews.Open strSQL,Conn,3,1		
						If rsLastNews.Eof Then
					%>					
					<tr>
						<td colspan="2">
						<p align="justify" style="margin-top: 4px; margin-bottom:4px">
						<font color="#666666"><b>Các tin đã đưa:</b></font></td>
					</tr>
						<%
							iCount = 1
							Do While Not rsLastNews.Eof And iCount <=15
							If iCount <=6 Then 
						%>
					<tr>
						<td width="4%">
						<p align="center" style="margin-top: 4px">
						<img border="0" src="../images/Pic/next_sm.gif" width="13" height="13" align="middle"></td>
						<td width="96%">
						<p style="margin-top: 2px; margin-bottom:2px">
						<a href="default.asp?name=newsdetail&id=<%=rsLastNews("ID")%>"><%=rsLastNews("TITLE")%></a></td>
					</tr>
						<%
							End If
							iCount = iCount + 1
							rsLastNews.MoveNext
							Loop
						End If	
						Conn.Close
						Set Conn = Nothing
						%>
				</table>
				</td>
			</tr>
			<tr>
				<td width="100%">
				<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
			</tr>
			<tr>
				<td width="100%">
				&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
</table>

</div>

</body>

</html>