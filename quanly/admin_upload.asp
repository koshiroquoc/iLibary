<%
Dim qStrDir
Dim qStrWin
Dim targetis
qStrDir		=	Request.QueryString("dir")
qStrWin		=	Request.QueryString("win")
targetis	=	Request.QueryString("targetis")
show		=	Request.QueryString("show")
%>
<html>
<head>
<title>Tải ảnh minh họa</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="JavaScript">
<!--
function insertStr(strValue,txtImage,txtDisplay){
	window.opener.document.frmAddNew[txtImage].value=strValue;
	window.opener.document.frmAddNew[txtDisplay].src=strValue;
}
function getFileExtension(filePath) { //v1.0
  fileName = ((filePath.indexOf('/') > -1) ? filePath.substring(filePath.lastIndexOf('/')+1,filePath.length) : filePath.substring(filePath.lastIndexOf('\\')+1,filePath.length));
  return fileName.substring(fileName.lastIndexOf('.')+1,fileName.length);
}

function checkFileUpload(form,extensions) { //v1.0
  document.MM_returnValue = true;
  if (extensions && extensions != '') {
    for (var i = 0; i<form.elements.length; i++) {
      field = form.elements[i];
      if (field.type.toUpperCase() != 'FILE') continue;
      if (field.value == '') {
        alert('File is required!');
        document.MM_returnValue = false;field.focus();break;
      }
      if (extensions.toUpperCase().indexOf(getFileExtension(field.value).toUpperCase()) == -1) {
        alert('This file is not allowed for uploading!');
        document.MM_returnValue = false;field.focus();break;
  } } }
}
//-->
</script>
<link rel="stylesheet" type="text/css" href="../css/admin.css">
</HEAD>
<Body>

<%
Sub BuildUploadRequest(RequestBin)

  PosBeg = 1
  PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
  if PosEnd = 0 then
    Response.Write "<b>Form was submitted with no ENCTYPE=""multipart/form-data""</b><br>"
    Response.Write "Please correct the form attributes and try again."
    Response.End
  end if
  boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
  boundaryPos = InstrB(1,RequestBin,boundary)

  Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))

    Dim UploadControl
    Set UploadControl = CreateObject("Scripting.Dictionary")

    Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
    Pos = InstrB(Pos,RequestBin,getByteString("name="))
    PosBeg = Pos+6
    PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
    Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
    PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
    PosBound = InstrB(PosEnd,RequestBin,boundary)

    If  PosFile<>0 AND (PosFile<PosBound) Then

      PosBeg = PosFile + 10
      PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
      FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      FileName = Mid(FileName,InStrRev(FileName,"\")+1)

      UploadControl.Add "FileName", FileName
      Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
      PosBeg = Pos+14
      PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))

      ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      UploadControl.Add "ContentType",ContentType

      PosBeg = PosEnd+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = FileName
      ValueBeg = PosBeg-1
      ValueLen = PosEnd-Posbeg
    Else

      Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
      PosBeg = Pos+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      ValueBeg = 0
      ValueEnd = 0
    End If

    UploadControl.Add "Value" , Value	
    UploadControl.Add "ValueBeg" , ValueBeg
    UploadControl.Add "ValueLen" , ValueLen	

    UploadRequest.Add name, UploadControl	

    BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
  Loop
End Sub


Function getByteString(StringStr)
  For i = 1 to Len(StringStr)
 	  char = Mid(StringStr,i,1)
	  getByteString = getByteString & chrB(AscB(char))
  Next
End Function

Function getString(StringBin)
  getString =""
  For intCount = 1 to LenB(StringBin)
	  getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
  Next
End Function

Function UploadFormRequest(name)
  on error resume next
  if UploadRequest.Item(name) then
    UploadFormRequest = UploadRequest.Item(name).Item("Value")
  end if  
End Function

UploadQueryString = Replace(Request.QueryString,"AF_upload=true","")
if mid(UploadQueryString,1,1) = "&" then
	UploadQueryString = Mid(UploadQueryString,2)
end if

AF_uploadAction = CStr(Request.ServerVariables("URL")) & "?AF_upload=true"
If (Request.QueryString <> "") Then  
  if UploadQueryString <> "" then
	  AF_uploadAction = AF_uploadAction & "&" & UploadQueryString
  end if 
End If

If (CStr(Request.QueryString("AF_upload")) <> "") Then
  AF_redirectPage = "success.asp"
  If (AF_redirectPage = "") Then
    AF_redirectPage = CStr(Request.ServerVariables("URL"))
  end if
    
  RequestBin = Request.BinaryRead(Request.TotalBytes)
  Dim UploadRequest
  Set UploadRequest = CreateObject("Scripting.Dictionary")  
  BuildUploadRequest RequestBin
  
  AF_keys = UploadRequest.Keys
  for AF_i = 0 to UploadRequest.Count - 1
    AF_curKey = AF_keys(AF_i)

    if UploadRequest.Item(AF_curKey).Item("FileName") <> "" then
      AF_value = UploadRequest.Item(AF_curKey).Item("Value")
      AF_valueBeg = UploadRequest.Item(AF_curKey).Item("ValueBeg")
      AF_valueLen = UploadRequest.Item(AF_curKey).Item("ValueLen")

      if AF_valueLen = 0 then
        Response.Write "<p><B>&#272;&#259; có m&#7897;t l&#7895;i x&#7843;y ra trong quá tr&#769;nh b&#7841;n upload file!</B><br><br>"
        Response.Write "Tên File: " & Trim(AF_curPath) & UploadRequest.Item(AF_curKey).Item("FileName") & "<br>"
        Response.Write "File không t&#7891;n t&#7841;i ho&#7863;c r&#7895;ng.<br>"
        Response.Write "B&#7841;n vui l&#803;ng ki&#7875;m tra và <A HREF=""javascript:history.back(1)"">th&#7917; l&#7841;i</a>"
	  	  response.End
	    end if
      
      Dim AF_strm1, AF_strm2
      Set AF_strm1 = Server.CreateObject("ADODB.Stream")
      Set AF_strm2 = Server.CreateObject("ADODB.Stream")
      
      AF_strm1.Open
      AF_strm1.Type = 1 'Binary
      AF_strm2.Open
      AF_strm2.Type = 1 'Binary
        
      AF_strm1.Write RequestBin
      AF_strm1.Position = AF_ValueBeg
      AF_strm1.CopyTo AF_strm2,AF_ValueLen
    
      AF_curPath = "../upload/"&qStrDir&"/"
      fname=UploadRequest.Item(AF_curKey).Item("FileName")
      on error resume next
      AF_strm2.SaveToFile Trim(Server.mappath(AF_curPath))& "\" & UploadRequest.Item(AF_curKey).Item("FileName"),2
      if err then
        Response.Write "<p><B>&#272;&#259; có m&#7897;t l&#7895;i x&#7843;y ra trong quá tr&#769;nh b&#7841;n upload file!</B><br><br>"
        Response.Write "Tên File: " & Trim(AF_curPath) & UploadRequest.Item(AF_curKey).Item("FileName") & "<br>"
        Response.Write "Th&#432; m&#7909;c &#273;&#7875; upload file vào không t&#7891;n t&#7841;i ho&#7863;c không &#273;&#432;&#7907;c phép upload file vào &#273;ó.<br>"
        Response.Write "B&#7841;n vui l&#803;ng ki&#7875;m tra và <A HREF=""javascript:history.back(1)"">th&#7917; l&#7841;i</a>"
  		  err.clear
	  	  response.End
	    end if
    end if
  next
  

  If (AF_redirectPage <> "") Then
 %>
<table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolorlight="#F4F2E8">
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="1">
			<tr>
				<td bgcolor="#EEEEEE"><b>
				&nbsp;Tải ảnh minh 
				họa</b></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" height="93">
			<tr>
				<td height="59">
				<p align="center" style="margin-top: 1px; margin-bottom: 1px">
				<b><font size="2">Tải ảnh thành công!</font></b></p>
				<p align="center" style="margin-top: 1px; margin-bottom: 1px">
				<font size="2">B&#7845;m vào &#273;ây 
				</font> 
				<a href="Javascript:insertStr('../upload/<%=qStrDir%>/<%=fname%>','<%=targetis%>','<%=show%>');">
				<font size="2">(<%=qStrDir%>/<%=fname%>)</font></a></p>
				<p align="center" style="margin-top: 1px; margin-bottom: 1px">
				<font size="2">&#273;&#7875; chèn &#273;&#432;&#7901;ng d&#7851;n 
				ảnh vào h&#7897;p văn bản ảnh minh họa.</font></td>
			</tr>
			<tr>
				<td>
				<p align="center">
							<input onclick="javascript:history.back();" type="button" value="Tiếp tục" name="Continue" class="input_button">
							<input onclick="javascript:window.close();" type="button" value=" Thoát " name="Close1" class="input_button"></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<%    
  end if  
  
Else
if UploadQueryString <> "" then
  UploadQueryString = UploadQueryString & "&AF_upload=true"
else  
  UploadQueryString = "AF_upload=true"
end if  
%>
<table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolorlight="#F4F2E8">
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="1">
			<tr>
				<td bgcolor="#EEEEEE"><b>
				&nbsp;Tải ảnh minh 
				họa</b></td>
			</tr>
			<tr>
			<form name="ASP" method="POST" enctype="multipart/form-data" action="admin_upload.asp?dir=<%=qStrDir%>&targetis=<%=targetis%>&show=<%=show%>&AF_upload=true">
				<td>
				<table border="0" width="100%">
					<tr>
						<td width="40%" height="10">						
						</td>
						<td width="60%" height="10">						
						</td>
					</tr>
					<tr>
						<td width="40%">						
						<p align="right"><b>Tên tập tin:</b></td>
						<td width="60%">						
							<input type="file" name="Files" size="25">
						</td>
					</tr>
					<tr>
						<td width="40%" height="5"></td>
						<td width="60%" height="5">
						</td>
					</tr>
					<tr>
						<td width="40%">&nbsp;</td>
						<td width="60%">
							<input type="submit" value="Tải ảnh" name="Submit1" class="input_button">
							<input onclick="javascript:window.close();" type="button" value=" Thoát " name="Close" class="input_button">
						</td>
					</tr>
					<tr>
						<td width="40%" height="5"></td>
						<td width="60%" height="5"></td>
					</tr>
				</table>
				</td>
			 </form>
			</tr>
		</table>
		</td>
	</tr>
</table>
<%End if%>
</Body>
</HTML>