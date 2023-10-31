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
		<td width="100%">
		<img border="0" src="../images/spacer.gif" width="1" height="5"></td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" id="table3" cellspacing="0" cellpadding="0">
			<tr>
				<td width="100%">
				<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
					<tr>
						<td>
						<table border="0" width="100%" cellspacing="0" cellpadding="0">
							<tr>
								<td width="10">
								<p style="margin-left: 3px; margin-top: 4px; margin-bottom: 4px">&nbsp;</td>
								<td width="304"><b>
								<a href="default.asp?name=notice">Quay lại</a></b></td>
								<td>
								<p align="right" style="margin-right: 3px">
						&nbsp;</td>
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
			<%
				id	= Request.QueryString("id")
				strSQL = "SELECT * FROM NOTICE WHERE ID =" & id
				Set rsNotice = Server.CreateObject("ADODB.Recordset")
				rsNotice.Open strSQL,Conn,3,1
			%>			
			<tr>
				<td width="100%">
				<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
					<tr>
						<td>
						<table border="0" width="100%" cellspacing="0" cellpadding="0">
							<tr>
								<td>
								<p align="center" style="margin-top: 0; margin-bottom: 10px"><b><%=UCase(rsNotice("TITLE"))%></b></td>
							</tr>
							<tr>
								<td>
								<p align="justify"><%=rsNotice("CONTENT")%></td>
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
								<a href="default.asp?name=notice">Quay lại</a></font></td>
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