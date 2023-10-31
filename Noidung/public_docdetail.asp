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
								<a href="default.asp?name=listdoc">
								<font color="#000080">Quay lại</font></a></b></td>
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
				strSQL = "SELECT * FROM DOCUMENT WHERE DOCUMENT_ID ='" & id & "'"
				Set rsDoc = Server.CreateObject("ADODB.Recordset")
				rsDoc.Open strSQL,Conn,3,1
		
				strSQL = "SELECT * FROM CATEGORY_DOCUMENT WHERE CATEGORY_ID ='" & LEFT(rsDoc("DOCUMENT_ID"),3) & "'"
				Set rsCategory = Server.CreateObject("ADODB.Recordset")
				rsCategory.Open strSQL,Conn,3,1
			%>			
			<tr>
				<td width="100%">
				<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
					<tr>
						<td>
						<table border="0" width="100%" cellspacing="0" cellpadding="0">
							<tr>
								<td>
								<p align="center" style="margin-top: 0; margin-bottom: 10px"><b><%=UCase(rsDoc("TITLE"))%></b></td>
							</tr>
							<tr>
								<td>
								<p align="justify"><%=rsDoc("CONTENT")%></td>
							</tr>
							<tr>
								<td>
								<p align="right" style="margin-top: 6px"><b><%=rsDoc("AUTHOR")%></b></td>
							</tr>
							<%
								If rsDoc("FILE_PATH") <>"" Then
							%>							
							<tr>
								<td>File đính kèm: 
								<b> 
								<a href="<%=rsDoc("FILE_PATH")%>">Download</a></b></td>
							</tr>
							<%
								End If
								Conn.Close
								Set Conn = Nothing
							%>
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