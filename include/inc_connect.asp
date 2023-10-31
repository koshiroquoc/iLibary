<%
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	strConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../dulieu/trungvuong.mdb")
	Conn.ConnectionString = strConn
	Conn.CursorLocation = 3
	Conn.Open
%>