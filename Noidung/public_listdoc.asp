<%	Session.CodePage = 65001 %>
<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>New Page 1</title>
<link rel="stylesheet" type="text/css" href="../css/public.css">
</head>

<body>
<div align="center">
<table border="0" width="667" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="3">
		<img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>
		<table border="1" width="100%" id="table5" cellspacing="0" cellpadding="0" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" id="table6" cellspacing="0" cellpadding="0">
					<tr>
						<td width="5"><p align="center">&nbsp;</td>
						<td width="51"><p align="center">
						<b>
								<a href="default.asp?name=listdoc">Quay lại</a></b></td>
						<form method="POST" name="frmSearch" action="default.asp?name=resultdoc">
						<td width="229">
						<p align="right" style="margin-right: 4px">
						<font size="2">Tìm kiếm</font></td>
						<td width="102">
						<p align="right" style="margin-right: 1px; margin-top: 1px; margin-bottom: 1px">
						<font size="2">
						<input type="text" name="txtSearchKey" size="17" class="textbox"></font></td>
						<td>
						<button name="B1" class="input_button" type="submit">
						<font size="2">&nbsp;Tìm&nbsp;</font>
						</button></font></td>
						</form>
					</tr>
					</table>
				</td>
			</tr>
		</table>
		</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
	<%
		strSQL = "SELECT * FROM CATEGORY_DOCUMENT"
		Set rsCategory = Server.CreateObject("ADODB.Recordset")
		rsCategory.Open strSQL,Conn,3,1
		iIndex = 1
		Do While Not rsCategory.Eof and iIndex <=10
		strSQL = "SELECT * FROM DOCUMENT WHERE LEFT(DOCUMENT_ID,3)='" & rsCategory("CATEGORY_ID") & "'"
		Set rsListDoc = Server.CreateObject("ADODB.Recordset")
		rsListDoc.Open strSQL,Conn,3,1
		If Not rsListDoc.Eof Then
	%>
	<tr>
		<td width="2">&nbsp;</td>
		<td width="428">
		<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0">
			<tr>
				<td bgcolor="#F8F8F8">
				<table border="0" width="100%" id="table3" bordercolorlight="#CCCCCC" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
					<tr>
						<td class="doc_category" align="center">
						<a href="default.asp?name=listcatedoc&id=<%=rsCategory("CATEGORY_ID")%>">						
						<p style="margin-top: 2px; margin-bottom: 2px"><b><%=UCase(rsCategory("NAME"))%></b></a></td>
					</tr>
				</table>
				</td>
			</tr>
			<%
				iCount = 1
				Do While Not rsListDoc.Eof And iCount <=4
			%>			
			<tr>
				<td width="100%">
				<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
					<tr>
						<td>
						<p style="margin-top: 3px; margin-bottom: 2px">
						<a href="default.asp?name=docdetail&id=<%=rsListDoc("DOCUMENT_ID")%>"><b><%=rsListDoc("TITLE")%></b></a></td>
					</tr>
					<tr>
						<td class="doc_summary">
						<p align="justify"><%=rsListDoc("SUMMARY")%></td>
					</tr>
					<%
						If iCount <> rsListDoc.RecordCount Then
					%>					
					<tr>
						<td>
						<p align="center" style="margin-top: 4px; margin-bottom: 8px">
						<img border="0" src="../images/line.gif" width="373" height="5"></td>
					</tr>
					<%
						Else
					%>
					<tr>
					<td>
						<p align="center" style="margin-top: 4px; margin-bottom: 8px">
						<img border="0" src="../images/space.gif" width="1" height="1"></td>
					</tr>
					<%
						End If
					%>
				</table>
				</td>				
			</tr>
			<%				
				iCount = iCount + 1
				rsListDoc.MoveNext
				Loop
			%>	
			</tr>
			</table>
		</td>
		<td width="2">&nbsp;</td>
	</tr>
	<%
		iIndex = iIndex + 1
		End If	
		rsCategory.MoveNext
		Loop
	%>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
</table>

</div>
</body>

</html>