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
	<%
		id	= Request.QueryString("id")
		If id <> "" Then
			strSQL = "SELECT * FROM DOCUMENT WHERE LEFT(DOCUMENT_ID,3)='" & id & "'"
		Else
			strSQL = LoadSQL()		
		End If
		
		Set rsListDoc = Server.CreateObject("ADODB.Recordset")
		rsListDoc.Open strSQL,Conn,3,1
		InsertSQL(strSQL)
		If Request.QueryString("page") = "" Then
			intCurrentPage = 1
		Else
			intCurrentPage = CInt(Request.QueryString("page"))
		End If
		rsListDoc.PageSize = 10
		If rsListDoc.PageCount > 0 then
			rsListDoc.AbsolutePage = intCurrentPage
		Else
			intCurrentPage = 0
		End If

	%>
	<tr>
		<td width="2">&nbsp;</td>
		<td width="428">
		<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0">
			<%
				If rsListDoc.Eof Then
			%>
			<tr>
				<td>
				<p align="center" style="margin-top: 10px; margin-bottom:10px">
				Không tìm thấy tài liệu nào trong mục này<br>
				Bấm <a href="JavaScript:history.go(-1)">vào đây</a> để quay lại.</td>
			</tr>			
			<%
				Else
					strSQL = "SELECT * FROM CATEGORY_DOCUMENT WHERE CATEGORY_ID='"& LEFT(rsListDoc("DOCUMENT_ID"),3) & "'"
					Set rsCategory = Server.CreateObject("ADODB.Recordset")
					rsCategory.Open strSQL,Conn,3,1
			%>
			<tr>
				<td>
				<table border="1" width="100%" id="table9" cellspacing="0" cellpadding="0" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF">
					<tr>
						<td>
						<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table10">
							<tr>
								<td width="10">
								<p style="margin-left: 3px; margin-top: 4px; margin-bottom: 4px">&nbsp;</td>
								<td width="52"><b>
								<a href="default.asp?name=listdoc">Quay lại</a></b></td>
								<form method="POST" name="frmSearch" action="default.asp?name=resultdoc">
						<td width="229">
						<p align="right" style="margin-right: 4px">Tìm kiếm</td>
						<td width="102">
						<p align="right" style="margin-right: 1px; margin-top: 1px; margin-bottom: 1px">
						<input type="text" name="txtSearchKey" size="17" class="textbox"></td>
						<td>
						<button name="B1" class="input_button" type="submit">
						&nbsp;Tìm&nbsp;
						</button></td>
								</form>
								<td>
								&nbsp;</td>
							</tr>
						</table>
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td>
				<img border="0" src="../images/spacer.gif" width="1" height="4"></td>
			</tr>
			<tr>
				<td bgcolor="#F8F8F8">
				<table border="0" width="100%" id="table3" bordercolorlight="#CCCCCC" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
					<tr>
						<td class="doc_category" align="center">						
						<p style="margin-top: 2px; margin-bottom: 2px"><b><%=UCase(rsCategory("Name"))%></b></td>
					</tr>
				</table>
				</td>
			</tr>
			<%
				numRecord = rsListDoc.RecordCount
				numPage = rsListDoc.PageCount
				numResult = rsListDoc.RecordCount
				iCount = 1
				iIndex = 1
				Do While Not rsListDoc.Eof And iIndex<= rsListDoc.PageSize
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
						If iCount <> numRecord Then
					%>					
					<tr>
						<td>
						<p align="center" style="margin-top: 4px; margin-bottom: 8px">
						<img border="0" src="../images/line.gif" width="373" height="5"></td>
					</tr>
					<%
						Else
					%>
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
				iIndex = iIndex + 1
				rsListDoc.MoveNext
				Loop
			%>	
			</tr>
			<%
				End If
			%>						
			</table>
		</td>
		<td width="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="3">&nbsp;</td>
	</tr>
	<% If numPage >1 Then %>
	<tr>
		<td colspan="3">
		<table border="1" width="100%" id="table7" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC">
			<tr>
				<td bgcolor="#EEEEEE">
				<table border="0" width="100%" id="table8" cellspacing="0" cellpadding="0">
					<tr>
						<td width="18"><a href="JavaScript:history.back();">
						<img border="0" src="../images/left.gif" width="22" height="20" alt="Quay lại"></a></td>
						<td>&nbsp;</td>
						<td width="20"><a href="#">
						<img border="0" src="../images/top.gif" width="20" height="22" alt="Lên trên"></a></td>
						<td width="386">
							<p align="right">Trang: 
								<%
									for i=1 to numPage
								%>	
								<b><a href="default.asp?name=listcatedoc&page=<%=i%>"><%=i%></a> |</b>
								<%	
									Next
								%>						
						</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<%
		End If
	%>
</table>

</div>
</body>

</html>