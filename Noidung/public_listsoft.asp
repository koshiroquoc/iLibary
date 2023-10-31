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
		<td><img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<p align="center" style="margin-top: 3px; margin-bottom: 3px"><font color="#666666">Tại đây bạn có thể tải 
				các file, các phần mềm tiện ích <br>
				phục vụ cho nhu cầu giảng dạy và học tập của 
				bạn.</font></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="10">
						<p style="margin-left: 3px; margin-top: 4px; margin-bottom: 4px">&nbsp;</td>
						<td width="242"><b><a href="default.asp?name=listsoft">Quay lại</a></b></td>
						<td width="75">
						<p align="center">
						<span style="font-size: 8pt; font-weight: 700">Tìm nhanh</span></td>
						<form method="POST" name="frmList" action="default.asp?name=resultsoft">
						<td>
						<p align="center" style="margin-right: 3px">
						<input type="text" name="txtSoftName" size="11" class="textbox"></td>
						<td width="39">						
						<button name="B1" type="submit" class="input_button">&nbsp;Tìm
						</button></td>
						</form>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td><img border="0" src="../images/spacer.gif" width="1" height="6"></td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolorlight="#D1DCE9" bordercolordark="#FFFFFF">
			<tr>
				<td>
				<%
					strSQL = "SELECT * FROM CATEGORY_SOFT ORDER BY NAME"
					Set rsCategory = Server.CreateObject("ADODB.Recordset")
					rsCategory.Open strSQL,Conn,3,1
					strSQL = "SELECT * FROM SOFTWARE ORDER BY NAME"
					Set rsSoft = Server.CreateObject("ADODB.Recordset")
					rsSoft.Open strSQL,Conn,3,1
					numCategory = rsCategory.RecordCount
					numSoft = rsSoft.RecordCount
				%>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
					<%	
						iCount = 1
						Do While Not rsCategory.Eof 
						If iCount Mod 2 <> 0 Then
					%>						
						<td width="44">
						<p align="right" style="margin-right: 6px; margin-top:7px; margin-bottom:7px">
						<img border="0" src="../images/folder.gif" width="22" height="22"></td>
						<td width="206"><a href="default.asp?name=listcatesoft&id=<%=rsCategory("ID")%>"><%=rsCategory("NAME")%></a></td>
					<%
						Else
					%>	
						<td width="37">
						<p style="margin-right: 6px"><img border="0" src="../images/folder.gif" width="22" height="22" align="right"></td>
						<td><a href="default.asp?name=listcatesoft&id=<%=rsCategory("ID")%>"><%=rsCategory("NAME")%></a></td>
					</tr>
					<%
						End If
						iCount = iCount + 1
						rsCategory.MoveNext
						Loop
						Conn.Close
						Set Conn = Nothing						
					%>
					<tr>
						<td colspan="4">
						<p align="center" style="margin-top: 12px; margin-bottom: 12px">
						<img border="0" src="../images/line.gif" width="310" height="5"><br>
						<b>Có tất cả <%=numSoft%> file trong <%=numCategory%> mục.</b></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>

</div>
</body>

</html>