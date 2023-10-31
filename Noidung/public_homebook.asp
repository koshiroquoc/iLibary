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
		<p align="center">
		<img border="0" src="../images/spacer.gif" width="1" height="10"></td>
	</tr>
		<%
			id	= Request.QueryString("id")
			strSQL = "SELECT * FROM BOOK WHERE ID="& id
				
			Set rsBookResult = Server.CreateObject("ADODB.Recordset")
			rsBookResult.Open strSQL,Conn,3,1

			strSQL = "SELECT * FROM CATEGORY_BOOK WHERE CATEGORY_ID='" & Left(rsBookResult("BOOK_ID"),3) & "'"
			Set rsCategoryBook = Server.CreateObject("ADODB.Recordset")
			rsCategoryBook.Open strSQL,Conn,3,1
			
			strSQL = "SELECT * FROM PUBLISHER WHERE ID=" & rsBookResult("PUBLISHER")
			Set rsPublisher = Server.CreateObject("ADODB.Recordset")
			rsPublisher.Open strSQL,Conn,3,1

			strSQL = "SELECT * FROM CATEGORY_GENRE WHERE ID=" & rsBookResult("GENRE")
			Set rsGenre = Server.CreateObject("ADODB.Recordset")
			rsGenre.Open strSQL,Conn,3,1
			
			strSQL = "SELECT * FROM LANGUAGE WHERE ID=" & rsBookResult("LANGUAGE")
			Set rsLanguage = Server.CreateObject("ADODB.Recordset")
			rsLanguage.Open strSQL,Conn,3,1
		%>
	<tr>
		<td>
		<table border="1" width="100%" id="table4"  cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" id="table5" cellspacing="0" cellpadding="0">
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Tên 
						sách</b></td>
						<td width="2%">
						<p style="margin-top: 2px"><b>:</b></td>
						<td width="56%"><font color="#0066FF"><b><%=rsBookResult("NAME")%></b></font></td>
						<td width="20%" rowspan="6">
						<p style="margin-top: 8px">
						<img border="0" width="77" height="91" align="right" src="<%=rsBookResult("IMAGE")%>"></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Mã sách</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsBookResult("BOOK_ID")%></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Thể 
						loại</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsGenre("NAME")%></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Lĩnh vực</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsCategoryBook("NAME")%></td>
						<td width="1%">&nbsp;</td>
					</tr>					
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<b>Tên tác giả</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsBookResult("AUTHOR")%></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Nhà xuất bản</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsPublisher("NAME")%></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 2px"><b>Ngôn ngữ</b></td>
						<td width="2%">
						<b>:</b></td>
						<td width="56%"><%=rsLanguage("NAME")%></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td width="1%">&nbsp;</td>
						<td width="20%">
						<p style="margin-top: 2px; margin-bottom: 8px"><b>Tóm tắt</b></td>
						<td width="2%">
						<p style="margin-bottom: 8px"><b>:</b></td>
						<td colspan="2">
						<p align="justify" style="margin-bottom: 8px"><%=rsBookResult("SUMMARY")%></td>
						<td width="1%">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="6">
						<p align="center" style="margin-top: 5px; margin-bottom: 5px">
						<img border="0" src="../images/line.gif" width="276" height="5"></td>
						</tr>
					<tr>
						<td colspan="6">
						<table border="0" width="100%" id="table7" cellspacing="0" cellpadding="0">
							<tr>
								<td width="147">&nbsp;</td>
								<td width="33">
								<p align="center">
								<img border="0" src="../images/abook_add.gif" width="17" height="19"></td>
								<td><b>
								<a href="#" onclick="JavaScript:openWindow3('public_register.asp?id=<%=rsBookResult("ID")%>')">
								Đăng ký mượn</a></b></td>
								<td width="136">&nbsp;</td>
							</tr>
						</table>
						</td>
						</tr>
					<tr>
						<td colspan="6" height="6">
						<p align="center"></td>
					</tr>
					</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="3"></td>
	</tr>
		<%
			Conn.Close
			Set Conn = Nothing
		%>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" id="table3" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC">
			<tr>
				<td bgcolor="#EEEEEE">
				<table border="0" width="100%" id="table6" cellspacing="0" cellpadding="0">
					<tr>
						<td width="18"><a href="JavaScript:history.back();">
						<img border="0" src="../images/left.gif" width="22" height="20" alt="Quay lại"></a></td>
						<td>&nbsp;</td>
						<td width="20"><a href="#">
						<img border="0" src="../images/top.gif" width="20" height="22" alt="Lên trên"></a></td>
						<td width="386">
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="2"></td>
	</tr>
</table>

</div>
</body>

</html>