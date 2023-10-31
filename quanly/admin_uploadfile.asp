<%
	qStrDir=Request.QueryString("dir")
	qStrWin=Request.QueryString("win")
	targetis=Request.QueryString("targetis")
	
	txtCheck = False
	
	Server.ScriptTimeout = 5000
	
	'Create upload form
	'Using Huge-ASP file upload
	'Dim Form: Set Form = Server.CreateObject("ScriptUtils.ASPForm")
	'Using Pure-ASP file upload
	Dim Form: Set Form = New ASPForm %><!--#INCLUDE FILE="../include/inc_upload.asp"--><% 

	Server.ScriptTimeout = 1000
	Form.SizeLimit = &HA00000'10MB
	
	Dim strFileName
	'was the Form successfully received?
	Const fsCompletted  = 0
	Dim DestinationFileName
	txtCurPath = "../upload/"&qStrDir&"/"
	txtFileName = Form("SourceFile").FileName
	
	If Form.State = fsCompletted Then 'Completted
		DestinationPath = Server.MapPath(txtCurPath)
  		DestinationFileName = DestinationPath & "\" & Form("SourceFile").FileName
		Form("SourceFile").SaveAs DestinationFileName
		txtCheck = True
	Else
		txtCheck = False	
	End If
%>
<html>
<head>
<title>Tải file đính kèm</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="JavaScript">
<!--
function insertStr(strValue,txtFilePath){
	window.opener.document.frmAddNew[txtFilePath].value=strValue;
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
	If txtCheck= True Then 'Completted
%>
<table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolorlight="#F4F2E8">
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="1">
			<tr>
				<td bgcolor="#EEEEEE"><b>&nbsp;&nbsp;Tải tập tin đính kèm</b></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" height="82">
			<tr>
				<td height="50">
				<p align="center"><b>Upload thành công!<br>
				</b>B&#7845;m vào &#273;ây 
				<a href="Javascript:insertStr('../upload/<%=qStrDir%>/<%=txtFileName%>','<%=targetis%>');">
				(<%=qStrDir%>/<%=txtFileName%>)</a><br>
				&#273;&#7875; &#273;&#432;a &#273;&#432;&#7901;ng d&#7851;n vào h&#7897;p nh&#7853;p.</p></td>
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
	Else
%>
<table border="1" width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse"  bordercolorlight="#F4F2E8">
	<tr>
		<td>
		<table border="0" width="100%" cellspacing="1">
			<tr>
				<td bgcolor="#EEEEEE"><b>&nbsp;&nbsp;Tải tập tin đính kèm</b></td>
			</tr>
			<tr>
			<form name="ASP" method="POST" enctype="multipart/form-data">
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
						<p align="right"><b>Tên tập 
						tin:
						</b>
						</td>
						<td width="60%">						
							<input type="file" name="SourceFile" size="25">
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
							<input type="submit" value="Tải file" name="Submit1" class="input_button">
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