<!-- #INCLUDE FILE="../include/inc_connect.asp" -->
<!-- #INCLUDE FILE="../include/inc_function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>In danh sách</title>
<link rel="stylesheet" type="text/css" href="../css/admin.css">

<body>
<%
	Session.CodePage = 65001
	txtTypePrint = Request.QueryString("typeprint")
	txtClassID = Request.QueryString("class")
	If txtClassID <> "All" Then
		If txtTypePrint = "borrow" Then
			txtTitle = "DANH S&#193;CH S&#193;CH CHO M&#431;&#7906;N " & uCase(txtClassID)			
			strSQL = "SELECT * FROM BORROW WHERE CLASS_ID='" & txtClassID & "'" 
			strSQL = strSQL & " ORDER BY CARD_ID ASC"
		ElseIf txtTypePrint = "breaching" Then
			txtTitle = "DANH S&#193;CH M&#431;&#7906;N QU&#193; H&#7840;N " & uCase(txtClassID)			
			strSQL = "SELECT * FROM BORROW WHERE CLASS_ID='" & txtClassID & "'" 
			strSQL = strSQL & " AND NOW()-DATE_INFORM>7"			
			strSQL = strSQL & " ORDER BY CARD_ID ASC"			
		Else		
			txtTitle = "DANH S&#193;CH TH&#7866; VI PH&#7840;M " & uCase(txtClassID)			
			strSQL = "SELECT * FROM BREACH WHERE CLASS_ID='" & txtClassID & "'" 
			strSQL = strSQL & " ORDER BY CARD_ID ASC"			
		End If	
	Else
		If txtTypePrint = "borrow" Then
			txtTitle = "DANH S&#193;CH S&#193;CH CHO M&#431;&#7906;N"			
			strSQL = "SELECT * FROM BORROW" 
			strSQL = strSQL & " ORDER BY CARD_ID ASC"
		ElseIf txtTypePrint = "breaching" Then
			txtTitle = "DANH S&#193;CH M&#431;&#7906;N QU&#193; H&#7840;N"
			strSQL = "SELECT * FROM BORROW WHERE NOW()-DATE_INFORM>7"
			strSQL = strSQL & " ORDER BY CARD_ID ASC"			
		Else
			txtTitle = "DANH S&#193;CH TH&#7866; VI PH&#7840;M"		
			strSQL = "SELECT * FROM BREACH" 
			strSQL = strSQL & " ORDER BY CARD_ID ASC"			
		End If				
	End If
	
%>
<p align="center">
<b><font face="Times New Roman" size="3" color="#0000FF"><%=txtTitle%></font></b></p>
<%	
	Set rsSelect = Server.CreateObject("ADODB.Recordset")
	rsSelect.Open strSQL,Conn,3,1
	numRecord = rsSelect.RecordCount
%>		

<table border="1" width="100%" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC" id="table1">
	<tr>
		<td width="26" align="center" bgcolor="#EEEEEE" height="20"><b>STT</b></td>
		<td align="center" width="66" bgcolor="#EEEEEE" height="20">
		<b>Mã thẻ</b></td>
		<td align="center" width="180" bgcolor="#EEEEEE" height="20">
		<b>Họ và tên</b></td>
		<td align="center" width="74" bgcolor="#EEEEEE" height="20">
		<b>Lớp</b></td>
		<td align="center" width="325" bgcolor="#EEEEEE" height="20">
		<b>Tên sách</b></td>
		<td align="center" bgcolor="#EEEEEE" height="20" width="80">
		<b>Ngày mượn</b></td>
	</tr>
		<%
			Dim iCount
			iCount = 1
			Do While Not rsSelect.Eof and iCount <=rsSelect.PageSize
			strSQL = "SELECT BOOK_ID, NAME FROM BOOK WHERE BOOK_ID='" & rsSelect("BOOK_ID") & "'"
			Set rsCategory = Server.CreateObject("ADODB.Recordset")
			rsCategory.Open strSQL,Conn,3,1
			strSQL = "SELECT CARD_ID, FIRSTNAME,LASTNAME FROM CARD WHERE CARD_ID='" & rsSelect("CARD_ID") & "'"
			Set rsCard = Server.CreateObject("ADODB.Recordset")
			rsCard.Open strSQL,Conn,3,1
		%>
	<tr>
		<td width="26" align="center"><%=iCount%></td>
		<td width="66">
		<p align="center" style="margin:2px 4px; ">
		<%=rsSelect("CARD_ID")%></a></td>
		<td width="180">
		<p align="justify" style="margin:2px 4px; "><%=rsCard("FIRSTNAME") & " " & rsCard("LASTNAME")%></td>
		<td width="74">
		<p align="center" style="margin:2px 4px; ">
		<%=rsSelect("CLASS_ID")%></td>
		<td width="325">
		<p align="justify" style="margin:2px 4px; "><%=rsCategory("NAME")%></td>
		<td align ="center"><%=NgayVN(rsSelect("DATE_INFORM"))%></td>
	</tr>
	<%
		iCount = iCount + 1
		rsSelect.MoveNext
		Loop
		Conn.Close
		Set Conn = Nothing
	%>
	<tr>
		<td align="center" colspan="6">
		<p style="margin-top: 8px; margin-bottom: 8px"><i>
		<font style="font-size: 9pt; font-weight: 700">Có tất cả&nbsp;<%=numRecord%> 
		độc giả trong danh sách</font></i></td>
	</tr>
</table>
							
<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0">
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td>
		<table border="1" width="100%" bordercolorlight="#FFFFFF" cellspacing="0" cellpadding="0" bordercolordark="#CCCCCC" id="table3">
			<tr>
				<td>
				<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0">
					<tr>
						<td width="257" bgcolor="#F7F7F7">&nbsp;</td>
						<td width="16" bgcolor="#F7F7F7">
						<img border="0" src="../images/print.gif" width="16" height="16"></td>
						<td width="123" bgcolor="#F7F7F7">
						<p align="left" style="margin-top: 5px; margin-bottom: 5px; margin-left:8px">
						<b><a href="#" onclick="JavaScript:window.print();">In 
						danh sách</a></b></td>
						<td width="22" bgcolor="#F7F7F7">
						<p align="center" style="margin-top: 5px; margin-bottom: 5px">
						<img border="0" src="../images/close.gif" width="14" height="9"></td>
						<td width="62" bgcolor="#F7F7F7">
						<p align="center"><a href="#" onclick="JavaScript:window.close();">
						<b>Đóng lại</b></a></td>
						<td width="278" bgcolor="#F7F7F7">&nbsp;</td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
							
</body>