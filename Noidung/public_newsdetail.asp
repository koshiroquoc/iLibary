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
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td>
	<%
		id	= Request.QueryString("id")
		strSQL = "SELECT * FROM NEWS WHERE ID =" & id
		Set rsNews = Server.CreateObject("ADODB.Recordset")
		rsNews.Open strSQL,Conn,3,1

		strSQL = "SELECT * FROM CATEGORY_NEWS WHERE ID =" & rsNews("CATEGORY_ID")
		Set rsCategory = Server.CreateObject("ADODB.Recordset")
		rsCategory.Open strSQL,Conn,3,1
		txtCategoryName	= rsCategory("NAME")
	%>
		<table border="0" width="100%" id="table3" cellspacing="0" cellpadding="0">
			<tr>
				<td background="../images/bg_bar.gif" height="24">
				<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
					<tr>
						<td width="23">
						<p align="center">
						<img border="0" src="../images/Pic/bullet.gif" width="10" height="10"></td>
						<td class="txt_title"><a href="default.asp?name=listcatenews&id=<%=rsCategory("ID")%>"><b><%=uCase(txtCategoryName)%></b></a></font></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="100%" class="news_title">
				<p style="margin-top: 4px; margin-bottom: 8px" class="title_news"><b><%=rsNews("TITLE")%></b></td>
			</tr>
			<tr>
				<td width="91%" valign="top">
				<img border="0" align="right" width="130" height="89" src="<%=rsNews("IMAGE")%>" style="border-style: outset; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" hspace="6">
				</p>
				<p align="justify" style="margin-right: 4px">
				<%=rsNews("CONTENT")%></td>
			</tr>
			<tr>
				<td align="right">
				<p style="margin-top: 4px"><b><%=rsNews("AUTHOR")%></b></td>
			</tr>			
			<tr>
				<td align="right">
				<p style="margin-top: 4px" align="center">
				<img border="0" src="../images/line.gif" width="373" height="5"></td>
			</tr>	
			<%
				strSQL = "SELECT * FROM NEWS WHERE CATEGORY_ID =" & rsNews("CATEGORY_ID")
				strSQL = strSQL  & " AND ID<>" & rsNews("ID") & " Order by DATE_INFORM Desc"
				Set rsLastNews = Server.CreateObject("ADODB.Recordset")
				rsLastNews.Open strSQL,Conn,3,1
				If Not rsLastNews.Eof Then
			%>					
			<tr>
				<td>
				<p align="justify" style="margin-top: 4px; margin-bottom:4px">
				<b><font color="#666666">Các tin đã đưa:</font></b></td>
			</tr>
			<tr>
				<td width="100%">
				<%
					iCount = 1			
					Do While Not rsLastNews.Eof And iCount<=10
				%>
				<table border="0" width="100%" id="table5" cellspacing="0" cellpadding="0">
					<tr>
						<td width="16">
						<p align="center" style="margin-top: 4px">
						<img border="0" src="../images/Pic/next_sm.gif" width="13" height="13"></td>
						<td><p style="margin-top: 4px"><a href="default.asp?name=newsdetail&id=<%=rsLastNews("ID")%>"><%=rsLastNews("TITLE")%></a></td>
					</tr>
				</table>
				<%
					iCount = iCount + 1
					rsLastNews.MoveNext
					Loop
				End If
				Conn.Close
				Set Conn = Nothing	
				%>
				</td>
			</tr>
			<tr>
				<td width="100%">&nbsp;</td>
			</tr>			
			<tr>
				<td align="right">
				<table border="1" width="100%" id="table6" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#EEEEEE" bordercolor="#EEEEEE">
					<tr>
						<td>
						<table border="0" width="100%" id="table7" cellspacing="0" cellpadding="0">
							<tr>
								<td width="23">
								<p align="center">
								<img border="0" src="../images/left.gif" width="22" height="20"></td>
								<td>
								<p align="left">
								<font style="font-weight: 700">
								<a href="JavaScript:history.back();">Quay lại</a></font></td>
								<td width="190">
								<p align="right"><span style="font-weight: 700">
								<a href="#">Đầu trang</a></span></td>
								<td width="20">
								<p align="center">
								<img border="0" src="../images/top.gif" width="20" height="22"></td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="100%">&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
</table>

</div>
</body>

</html>