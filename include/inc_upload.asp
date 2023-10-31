<%
'Pure-ASP upload v. 2.06 (with progress bar)
'This software is a FreeWare with limitted use.
'1. You can use this software to upload files with size up to 10MB for free
' if you want to upload bigger files, please register ScriptUtilities and Huge-ASP upload
' See license info at http://asp-upload.borec.net
'2. I'll be glad if you include <A Href="http://asp-upload.borec.net">Pure ASP file upload</A>
' or similar text link to http://www.motobit.com on the web site using Pure-ASP upload.


'Feel free to send comments/suggestions to help@pstruh.cz

Const adTypeBinary = 1
Const adTypeText = 2


Const xfsCompleted    = &H0 '0  Form was successfully processed. 
Const xfsNotPost      = &H1 '1  Request method is NOT post 
Const xfsZeroLength   = &H2 '2  Zero length request (there are no data in a source form) 
Const xfsInProgress   = &H3 '3  Form is in a middle of process. 
Const xfsNone         = &H5 '5  Initial form state 
Const xfsError        = &HA '10  
Const xfsNoBoundary   = &HB '11  Boundary of multipart/form-data is not specified. 
Const xfsUnknownType  = &HC '12  Unknown source form (Content-type must be multipart/form-data) 
Const xfsSizeLimit    = &HD '13  Form size exceeds allowed limit (ScriptUtils.ASPForm.SizeLimit) 
Const xfsTimeOut      = &HE '14  Upload time exceeds allowed limit (ScriptUtils.ASPForm.ReadTimeout) 
Const xfsNoConnected  = &HF '15  Client was disconnected before upload was completted.
Const xfsErrorBinaryRead = &H10 '16  Unexpected error from Request.BinaryRead method (ASP error).

'This class emulates ASPForm Class of Huge-ASP upload, http://upload.Motobit.cz
'See ScriptUtilities and Huge-ASP file upload help (ScptUtl.chm)

Class ASPForm
	Private m_ReadTime
	Public ChunkReadSize, BytesRead, TotalBytes, UploadID

	'non-used properties.
	Public TempPath, MaxMemoryStorage, CharSet, FormType, SourceData, ReadTimeout

	public Default Property Get Item(Key)
		Read
		Set Item = m_Items.Item(Key)
	End Property

	public Property Get Items
		Read
		Set Items = m_Items
	End Property

	public Property Get Files
		Read
		Set Files = m_Items.Files
	End Property

	public Property Get Texts
		Read
		Set Texts = m_Items.Texts
	End Property
	

	public Property Get NewUploadID
		Randomize
		NewUploadID = clng(rnd * &H7FFFFFFF)
	End Property

	Public Property Get ReadTime
		if isempty(m_ReadTime) then
			if not isempty(StartUploadTime) then ReadTime = Clng((Now() - StartUploadTime) * 86400 * 1000)
		else ' For progress window.
			ReadTime = m_ReadTime
		end if
	End Property

	Public Property Get State
		if m_State = xfsNone Then Read
		State = m_State
	End Property


	Private Function CheckRequestProperties
		'Wscript.Echo "**CheckRequestProperties"
	  If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then 'Request method must be "POST"
			m_State = xfsNotPost 
			Exit Function
		End If 'If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 
	
		Dim CT
		CT = Request.ServerVariables("HTTP_Content_Type") 'reads Content-Type header
		if len(CT) = 0 then CT = Request.ServerVariables("CONTENT_TYPE") 'reads Content-Type header UNIX/Linux 
	  If LCase(Left(CT, 19)) <> "multipart/form-data" Then 'Content-Type header must be "multipart/form-data"
			m_State = xfsUnknownType 
			Exit Function
		End If 'If LCase(Left(CT, 19)) <> "multipart/form-data" Then 

		Dim PosB 'help position variable
		'This is upload request.
		'Get the boundary and length from Content-Type header
		PosB = InStr(LCase(CT), "boundary=") 'Finds boundary
		If PosB = 0 Then
			m_State = xfsNoBoundary
			Exit Function
		End If 'If PosB = 0 Then
		If PosB > 0 Then Boundary = Mid(CT, PosB + 9) 'Separetes boundary

		'****** Error of IE5.01 - doubbles http header
		PosB = InStr(LCase(CT), "boundary=") 
		If PosB > 0 then 'Patch for the IE error
			PosB = InStr(Boundary, ",")
			If PosB > 0 Then Boundary = Left(Boundary, PosB - 1)
		
		end if
		
		'****** Error of IE5.01 - doubbles http header

		On Error Resume next
		TotalBytes = Request.TotalBytes
		If Err<>0 Then
			'For UNIX/Linux 
			
			TotalBytes = CLng(Request.ServerVariables("HTTP_Content_Length")) 'Get Content-Length header
			if len(TotalBytes)=0 then TotalBytes = CLng(Request.ServerVariables("CONTENT_LENGTH")) 'Get Content-Length header
		End If
		
		If TotalBytes = 0 then
			m_State = xfsZeroLength 
			Exit Function
		End If

		If IsInSizeLimit(TotalBytes) Then 'Form data are in allowed limit
			CheckRequestProperties = True
			m_State = xfsInProgress 
		Else   'Form data are in allowed limit
			'Form data exceeds the limit.
			m_State = xfsSizeLimit	
		End if 'Form data are in allowed limit

	End Function


	'reads source data using BinaryRead and store them in SourceData stream
	Public Sub Read()
		if m_State <> xfsNone Then Exit Sub
		'Wscript.Echo "**Read"
		If Not CheckRequestProperties Then 
			WriteProgressInfo
			Exit Sub
		End If

		'Initialize binary store stream
		if isempty(bSourceData) then Set bSourceData = createobject("ADODB.Stream")
		bSourceData.Open
		bSourceData.Type = 1 'Binary

		'Initialize Read variables.
		Dim DataPart, PartSize
		BytesRead = 0
		StartUploadTime = Now

		'read source data stream in chunks of ChunkReadSize
		Do While BytesRead < TotalBytes
			'Read chunk of data
			PartSize = ChunkReadSize
			if PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize
			'Wscript.Echo PartSize

			'Store the part size in our stream
			bSourceData.Write DataPart

			'Write progress info for secondary window.
			WriteProgressInfo

			'Check if the client is still connected
			If Not Response.IsClientConnected Then
				m_State = xfsNoConnected  
				Exit Sub
			End If
		Loop
		m_State = xfsCompleted

		'We have all source data in bSourceData stream
		ParseFormData
	End Sub

	Private Sub ParseFormData
		Dim Binary
		bSourceData.Position = 0
		Binary = bSourceData.Read
		'wscript.echo "Binary", LenB(Binary)
		m_Items.mpSeparateFields Binary, Boundary
	End Sub


	'This function reads progress info data from a temporary file.
	Public Function getForm(FormID)
		if isempty(ProgressFile.UploadID) Then 'Was UploadID of ProgressFile set?
			ProgressFile.UploadID = FormID
		End If

		'Get progress data
		Dim ProgressData
		
		ProgressData = ProgressFile
		
		if len(ProgressData) > 0 then 'There are some progress data
			if ProgressData = "DONE" Then 'Upload was done.
				ProgressFile.Done
				Err.Raise 1, "getForm", "Upload was done"
			Else ' if ProgressData = "DONE" Then 'Upload was done.
				'm_State & vbCrLf & TotalBytes & vbCrLf & BytesRead & vbCrLf & ReadTime
				ProgressData = Split (ProgressData, vbCrLf)
				if ubound(ProgressData) = 3 Then
					m_State = clng(ProgressData(0))
					TotalBytes = clng(ProgressData(1))
					BytesRead = clng(ProgressData(2))
					m_ReadTime = clng(ProgressData(3))
				End If'if ubound(ProgressData) = 3 Then
			End If'if ProgressData = "DONE" Then 'Upload was done.
		end if'if len(ProgressData) > 0 then 'There are some progress data
		Set getForm = Me
	End Function


	'This function writes progress info data to a temporary file.
	Private Sub WriteProgressInfo
		If UploadID > 0 Then ' Is the upload ID defined? (Upload is using progress)
			if isempty(ProgressFile.UploadID) Then 'Was UploadID of ProgressFile set?
				ProgressFile.UploadID = UploadID
			End If

			Dim ProgressData, FileName
			ProgressData = m_State & vbCrLf & TotalBytes & vbCrLf & BytesRead & vbCrLf & ReadTime
			ProgressFile.Contents = ProgressData
		End If
	End Sub

	'ASPForm Constructor 
	Private Sub Class_Initialize()
		ChunkReadSize = &H10000 '64 kB
		SizeLimit = &H100000 '1MB

		BytesRead = 0
		m_State = xfsNone
		
		TotalBytes = Request.TotalBytes

		Set ProgressFile = New cProgressFile
		Set m_Items = New cFormFields
	End Sub

	'ASPForm Destructor
	Private Sub Class_Terminate()
		If UploadID > 0 Then ' Is the upload ID defined? (Upload is using progress)
			'We have to close info about upload.
			ProgressFile.Contents = "DONE"
		End If
	End Sub

	Private Function IsInSizeLimit(TotalBytes)
		IsInSizeLimit = (m_SizeLimit = 0 or m_SizeLimit > TotalBytes) and (MaxLicensedLimit > TotalBytes)
	End Function

	Public Property Get SizeLimit
		SizeLimit = m_SizeLimit
	End Property 

	'Pure - ASP upload is a free script, but with 10MB upload limit
  ' if you want to upload bigger files, please register ScriptUtilities and Huge-ASP upload
	' at http://www.motobit.com/help/scptutl/lc2.htm
	Public Property Let SizeLimit(NewLimit)
	if NewLimit > MaxLicensedLimit Then
			Err.Raise 1, "ASPForm - SizeLimit", "This version of Pure-ASP upload is licensed with maximum limit of 10MB (" & MaxLicensedLimit & "B)"
			m_SizeLimit = MaxLicensedLimit
		Else
			m_SizeLimit = NewLimit
		end if
	End Property 

	Public Boundary
	Private m_Items 
	Private m_State
	Private m_SizeLimit 'Defined form size limit.
	Private bSourceData 'ADODB.Stream
	Private StartUploadTime , TempFiolder 
	Private ProgressFile 'File with info about current progress
End Class 'ASPForm
Const MaxLicensedLimit = &HA00000


'************************************************************************
'Emulates ScriptUtilities FormFields object
'We must have such class because of multiselect fields.
'See http://www.motobit.com
Class cFormFields
	Dim m_Keys()
	Dim m_Items()
	Dim m_Count
	

	Public Default Property Get Item(ByVal Key)
		If vartype(Key) = vbInteger or vartype(Key) = vbLong then
			'Numeric index
			if Key<1 or Key>m_Count Then Err.raise "Index out of bounds"
			Set Item = m_Items(Key-1)
			Exit Property
		end if

		'wscript.echo "**Item", Key
		Dim Count
		Count = ItemCount(Key)
		Key = LCase(Key)
		
		If Count > 0 then
			If Count>1 Then
				'More items
				'Get them All as an cFormFields
				Dim OutItem, ItemCounter
				Set OutItem = New cFormFields
				ItemCounter = 0
				
				For ItemCounter = 0 To Ubound(m_Keys)
					If LCase(m_Keys(ItemCounter)) = Key then OutItem.Add Key, m_Items(ItemCounter)
				Next
				Set Item = OutItem
				'wscript.echo "***Item-More", Key
			Else 
				For ItemCounter = 0 To Ubound(m_Keys)
					If LCase(m_Keys(ItemCounter)) = Key then exit for
				Next

				if isobject (m_Items(ItemCounter)) then
					Set Item = m_Items(ItemCounter)
				else
					Item = m_Items(ItemCounter)
				end if
				'wscript.echo "***Item-One", Key
			End If
		Else'No item 
			Set Item = New cFormField
		End if
	End Property

	Public Property Get MultiItem(ByVal Key)
		'returns an array of items with the same Key
		Dim Out: Set Out = New cFormFields
		Dim I, vItem 
		Dim Count
		Count = ItemCount(Key)
		
		if Count = 1 then
			'one key - get it from Item
			Out.Add Key, Item(Key)
		elseif Count > 1 then
			'more keys - enumerate them using Items
			For Each I In Item(Key).Items
				Out.Add Key, I
			Next
		End If

		Set MultiItem = Out
	End Property


	'For multiitem (I'm sorry, VBS does not support optional parameters for Item property)
	Public Property Get Value
		Dim I, V
		For Each I in m_Items
			V = V & ", " & I 
		Next
		V = Mid(V, 3)
		Value = V
	End Property


	Public Property Get xA_NewEnum
		Set xA_NewEnum = m_Items
	End Property

	Public Property Get Items()
		'Wscript.Echo "**cFormFields-Items"		
		Items = m_Items
	End Property

	Public Property Get Keys()
		Keys = m_Keys
	End Property

	public Property Get Files
		Dim cItem, OutItem, ItemCounter
		Set OutItem = New cFormFields 
		ItemCounter = 0
		if m_Count > 0 then ' Enumerate only non-empty form
			For ItemCounter = 0 To Ubound(m_Keys)
				Set cItem = m_Items(ItemCounter)
				if cItem.IsFile then
					OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter)
				end if
			Next
		End If
		Set Files = OutItem 
	End Property

	Public Property Get Texts
		Dim cItem, OutItem, ItemCounter
		Set OutItem = New cFormFields 
		ItemCounter = 0
		
		For ItemCounter = 0 To Ubound(m_Keys)
			Set cItem = m_Items(ItemCounter)
			if Not cItem.IsFile then
				OutItem.Add m_Keys(ItemCounter), m_Items(ItemCounter)
			end if
		Next
		Set Texts = OutItem
	End Property

	Public Sub Save(Path)
		Dim Item
		For Each Item In m_Items
			If Item.isFile Then
				Item.Save Path
			End If
		Next
	End Sub


	'Count of dictionary items within specified key
	Public Property Get ItemCount(ByVal Key)
		'wscript.echo "ItemCount"
		Dim cKey, Counter
		Counter = 0
		Key = LCase(Key)
		For Each cKey In m_Keys
			'wscript.echo "ItemCount", "cKey"
			If LCase(cKey) = Key then Counter = Counter + 1
		Next
		ItemCount = Counter
	End Property

	'Count of all dictionary items
	Public Property Get Count()
		Count = m_Count
	End Property

	Public Sub Add(byval Key, Item)
		Key = "" & Key
		ReDim Preserve m_Items(m_Count)
		ReDim Preserve m_Keys(m_Count)
		m_Keys(m_Count) = Key
		Set m_Items(m_Count) = Item
		m_Count = m_Count + 1
	End Sub

	Private Sub Class_Initialize()
		Dim vHelp()
		' I do not know why, but some of VBS verrsions declares m_Items and m_Keys as Empty,
		' not as Variant() - see class variables.
		' vHelp eliminates this problem. V. 2.03, 2.04
		On Error Resume Next
		m_Items = vHelp
		m_Keys = vHelp
		m_Count = 0
	End Sub


	'********************************** mpSeparateFields **********************************
	'This method retrieves the upload fields from binary data 
	'Binary is safearray ( VT_UI1 | VT_ARRAY ) of all multipart document raw binary data from input.
	Public Sub mpSeparateFields(Binary, ByVal Boundary)
		Dim PosOpenBoundary, PosCloseBoundary, PosEndOfHeader, isLastBoundary

		Boundary = "--" & Boundary			
		Boundary = StringToBinary(Boundary)

		PosOpenBoundary = InStrB(Binary, Boundary)
		PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary, 0)

		Do While (PosOpenBoundary > 0 And PosCloseBoundary > 0 And Not isLastBoundary)
			'Header and file/source field data
			Dim HeaderContent, bFieldContent
			'Header fields
			Dim Content_Disposition, FormFieldName, SourceFileName, Content_Type
			'Helping variables
			Dim TwoCharsAfterEndBoundary
			'Get end of header
			PosEndOfHeader = InStrB(PosOpenBoundary + Len(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))

			'Separates field header
			HeaderContent = MidB(Binary, PosOpenBoundary + LenB(Boundary) + 2, PosEndOfHeader - PosOpenBoundary - LenB(Boundary) - 2)
    
			'Separates field content
			bFieldContent = MidB(Binary, (PosEndOfHeader + 4), PosCloseBoundary - (PosEndOfHeader + 4) - 2)
			
			'Separates header fields from header
			GetHeadFields BinaryToString(HeaderContent), FormFieldName, SourceFileName, Content_Disposition, Content_Type

			'Create one field and assign parameters
			
			Dim Field        'All field values.
			Set Field = New cFormField

			Field.ByteArray = MultiByteToBinary(bFieldContent)

			Field.Name = FormFieldName
			Field.ContentDisposition = Content_Disposition
			if not isempty(SourceFileName) then
				Field.FilePath = SourceFileName
				Field.FileName = GetFileName(SourceFileName)
				Field.FileExt = GetFileExt(SourceFileName)
			else'if not isempty(SourceFileName) then
			End If'if not isempty(SourceFileName) then
			Field.ContentType = Content_Type
			
			Add FormFieldName, Field

			'Is this last boundary ?
			TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, PosCloseBoundary + LenB(Boundary), 2))
			isLastBoundary = TwoCharsAfterEndBoundary = "--"

			If Not isLastBoundary Then 'This is not last boundary - go to next form field.
				PosOpenBoundary = PosCloseBoundary
				PosCloseBoundary = InStrB(PosOpenBoundary + LenB(Boundary), Binary, Boundary)
			End If
		Loop
	End Sub
End Class 'cFormFields











'This class transfers data between primary (upload) and secondary (progress) window.
Class cProgressFile
	Private fs
	Public TempFolder
	Public m_UploadID
	Public TempFileName

	Public Default Property Get Contents()
		Contents = GetFile(TempFileName)
	End Property

	Public Property Let Contents(inContents)
		WriteFile TempFileName, inContents
	End Property

	Public Sub Done 'Delete temporary file when upload was done.
		FS.DeleteFile TempFileName
	End Sub

	Public Property Get UploadID()
		UploadID = m_UploadID
	End Property

	Public Property Let UploadID(inUploadID)
		if isempty(FS) then Set fs = CreateObject("Scripting.FileSystemObject")
		TempFolder = fs.GetSpecialFolder(2)

		m_UploadID = inUploadID
		TempFileName = TempFolder & "\pu" & m_UploadID & ".~tmp"
		
		Dim DateLastModified
		on error resume next
		DateLastModified = fs.GetFile(TempFileName).DateLastModified
		on error goto 0
		if isempty(DateLastModified) then 'OK
		elseif Now-DateLastModified>1 Then 'I think upload duration will be less than one day
			FS.DeleteFile TempFileName
		end if
	End Property

	Private Function GetFile(Byref FileName)
		
		Dim InStream
		On Error Resume Next
		Set InStream = fs.OpenTextFile(FileName, 1)
		GetFile = InStream.ReadAll
		On Error Goto 0
	End Function

	Private Function WriteFile(Byref FileName, Byref Contents)
		'wscript.echo "WriteFile", FileName, Contents
		Dim OutStream
		On Error Resume Next
		Set OutStream = fs.OpenTextFile(FileName, 2, True)
		OutStream.Write Contents
	End Function


	Private Sub Class_Initialize()
	End Sub
End Class 'cProgressFile



'******************************************************************************
'Emulates ScriptUtilities FormField object
'See http://www.motobit.com
Class cFormField
	'Used properties
	Public ContentDisposition, ContentType, FileName, FilePath, FileExt, Name
	Public ByteArray

	'non-used properties.
	Public CharSet, HexString, InProgress, SourceLength, RAWHeader, Index, ContentTransferEncoding
 
	Public Default Property Get String()
		'wscript.echo "**Field-String", Name, LenB(ByteArray)
		String = BinaryToString(ByteArray)
	End Property 

	Public Property Get IsFile()
		IsFile = not isempty(FileName)
	End Property

	Public Property Get Length()
		Length = LenB(ByteArray)
	End Property

	Public Property Get Value()
		Set Value = Me
	End Property

	Public Sub Save(Path)
	  '2.06 - and len(FileName)>0
		if IsFile and len(FileName)>0 Then
			Dim fullFileName
			fullFileName = Path & "\" & FileName
			SaveAs fullFileName
		Else
			'response.write "<br>" & typename(Name)
			'Err.Raise 1, "Text field " & Name & " does not have a file name"
		End If
	End Sub

	Public Sub SaveAs(newFileName)
		'2.06 - removed if len(ByteArray)>0 then 
		SaveBinaryData newFileName, ByteArray
	End Sub
	
End Class























Function StringToBinary(String)
  Dim I, B
  For I=1 to len(String)
    B = B & ChrB(Asc(Mid(String,I,1)))
  Next
  StringToBinary = B
End Function

Function BinaryToString(Binary)
  '2001 Antonin Foller, Motobit Software
  'Optimized version of PureASP conversion function
  'Selects the best algorithm to convert binary data to String data
  Dim TempString 

  On Error Resume Next
  'Recordset conversion has a best functionality
  TempString = RSBinaryToString(Binary)
  If Len(TempString) <> LenB(Binary) then'Conversion error
    'We have to use multibyte version of BinaryToString
    TempString = MBBinaryToString(Binary)
  end if
  BinaryToString = TempString
End Function


Function MBBinaryToString(Binary)
  '1999 Antonin Foller, Motobit Software
  'MultiByte version of BinaryToString function
	'Optimized version of simple BinaryToString algorithm.
  dim cl1, cl2, cl3, pl1, pl2, pl3
  Dim L', nullchar
  cl1 = 1
  cl2 = 1
  cl3 = 1
  L = LenB(Binary)
  
  Do While cl1<=L
    pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
    cl1 = cl1 + 1
    cl3 = cl3 + 1
    if cl3>300 then
      pl2 = pl2 & pl3
      pl3 = ""
      cl3 = 1
      cl2 = cl2 + 1
      if cl2>200 then
        pl1 = pl1 & pl2
        pl2 = ""
        cl2 = 1
      End If
    End If
  Loop
  MBBinaryToString = pl1 & pl2 & pl3
End Function


Function RSBinaryToString(xBinary)
  '1999 Antonin Foller, Motobit Software
  'This function converts binary data (VT_UI1 | VT_ARRAY or MultiByte string)
	'to string (BSTR) using ADO recordset
	'The fastest way - requires ADODB.Recordset
	'Use this function instead of MBBinaryToString if you have ADODB.Recordset installed
	'to eliminate problem with PureASP performance

	Dim Binary
	'MultiByte data must be converted to VT_UI1 | VT_ARRAY first.
	if vartype(xBinary) = 8 then Binary = MultiByteToBinary(xBinary) else Binary = xBinary
	
  Dim RS, LBinary
  Const adLongVarChar = 201
  Set RS = CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)
	
	if LBinary>0 then
		RS.Fields.Append "mBinary", adLongVarChar, LBinary
		RS.Open
		RS.AddNew
			RS("mBinary").AppendChunk Binary 
		RS.Update
		RSBinaryToString = RS("mBinary")
	Else
		RSBinaryToString = ""
	End If
End Function


Function MultiByteToBinary(MultiByte)
  ' This function converts multibyte string to real binary data (VT_UI1 | VT_ARRAY)
  ' Using recordset
  Dim RS, LMultiByte, Binary
  Const adLongVarBinary = 205
  Set RS = CreateObject("ADODB.Recordset")
  LMultiByte = LenB(MultiByte)
	if LMultiByte>0 then
		RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
		RS.Open
		RS.AddNew
			RS("mBinary").AppendChunk MultiByte & ChrB(0)
		RS.Update
		Binary = RS("mBinary").GetChunk(LMultiByte)
	End If
  MultiByteToBinary = Binary
End Function



'************** Upload Utilities 
'Separates header fields from upload header
Function GetHeadFields(ByVal Head, Name, FileName, Content_Disposition, Content_Type)
  'Get name of the field. Name is separated by name= and ;
  Name = (SeparateField(Head, "name=", ";")) 'ltrim
  'Remove quotes (if the field name is quoted)
  If Left(Name, 1) = """" Then Name = Mid(Name, 2, Len(Name) - 2)

  'Same for source filename
  FileName = (SeparateField(Head, "filename=", ";")) 'ltrim

  If Left(FileName, 1) = """" Then FileName = Mid(FileName, 2, Len(FileName) - 2)


  'Separate content-disposition and content-type header fields
  Content_Disposition = LTrim(SeparateField(Head, "content-disposition:", ";"))
  Content_Type = LTrim(SeparateField(Head, "content-type:", ";"))
End Function

'Separates one field between sStart and sEnd
Function SeparateField(From, ByVal sStart, ByVal sEnd)
  Dim PosB, PosE, sFrom
  sFrom = LCase(From)
  PosB = InStr(sFrom, sStart)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    PosE = InStr(PosB, sFrom, sEnd)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(From, PosB, PosE - PosB)
  Else
    SeparateField = Empty
  End If
End Function

Function SplitFileName(FullPath)
  Dim Pos, PosF
  PosF = 0
  For Pos = Len(FullPath) To 1 Step -1
    Select Case Mid(FullPath, Pos, 1)
      Case ":", "/", "\": PosF = Pos + 1: Pos = 0
    End Select
  Next
  If PosF = 0 Then PosF = 1
	SplitFileName = PosF
End Function

Function GetPath(FullPath)
  GetPath = left(FullPath, SplitFileName(FullPath)-1)
End Function

'Separetes file name from the full path of file
Function GetFileName(FullPath)
  GetFileName = Mid(FullPath, SplitFileName(FullPath))
End Function

'Separetes file name from the full path of file
Function GetFileExt(FullPath)
	Dim Pos: Pos = InStrRev(FullPath,".")
	if Pos>0 then GetFileExt = Mid(FullPath, Pos)
End Function


Function RecurseMKDir(ByVal Path)
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
	
  Path = Replace(Path, "/", "\")
  If Right(Path, 1) <> "\" Then Path = Path & "\"   '"
  Dim Pos, n
  Pos = 0: n = 0
  Pos = InStr(Pos + 1, Path, "\")   '"
  Do While Pos > 0
    On Error Resume Next
    FS.CreateFolder Left(Path, Pos - 1)
    If Err = 0 Then n = n + 1
    Pos = InStr(Pos + 1, Path, "\")   '"
  Loop
  RecurseMKDir = n
End Function

Function SaveBinaryData(FileName, ByteArray)
	SaveBinaryData = SaveBinaryDataStream(FileName, ByteArray)
End Function

Function SaveBinaryDataTextStream(FileName, ByteArray)
  Dim FS : Set FS = CreateObject("Scripting.FileSystemObject")
	On error Resume next
  Dim TextStream 
	Set TextStream = FS.CreateTextFile(FileName)
	if Err = &H4c then 'Path not found.
		On error Goto 0
		RecurseMKDir GetPath(FileName)
		On error Resume next
		Set TextStream = FS.CreateTextFile(FileName)
	end if
  TextStream.Write BinaryToString(ByteArray) 'BinaryToString is in upload.inc.
  TextStream.Close

	Dim ErrMessage, ErrNumber
	ErrMessage = Err.Description
	ErrNumber = Err

	On Error Goto 0
	if ErrNumber<>0 then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage 

End Function

Function SaveBinaryDataStream(FileName, ByteArray)
	Dim BinaryStream
	Set BinaryStream = createobject("ADODB.Stream")
	BinaryStream.Type = 1 'Binary
	BinaryStream.Open
	'2.06 - zero byte file is legal
	if lenb(ByteArray)>0 then BinaryStream.Write ByteArray
	On error Resume next
	
	BinaryStream.SaveToFile FileName, 2 'Overwrite

	if Err = &Hbbc then 'Path not found.
		On error Goto 0
		RecurseMKDir GetPath(FileName)
		On error Resume next
		BinaryStream.SaveToFile FileName, 2 'Overwrite
	end if
	Dim ErrMessage, ErrNumber
	
	ErrMessage = Err.Description
	ErrNumber = Err

	On Error Goto 0
	if ErrNumber<>0 then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage 
	
End Function
'************** Upload Utilities - end















'Emulates response object
Class cResponse
	Public Property Get IsClientConnected
		randomize
		IsClientConnected = cbool(clng(rnd * 4))
		IsClientConnected = True
	End Property 
End Class 


Class cRequest
	Private Readed

	Private BinaryStream
	public function ServerVariables(Name)	
		select case UCase(Name) 
			Case "CONTENT_TYPE": 
			Case "HTTP_CONTENT_TYPE": 
				ServerVariables = "multipart/form-data; boundary=---------------------------7d21960404e2"
			Case "CONTENT_LENGTH": 
			Case "HTTP_CONTENT_LENGTH": 
				ServerVariables = "" & TotalBytes
			Case "REQUEST_METHOD": 
				ServerVariables = "POST"
		End Select
	End Function

	public function BinaryRead(ByRef Bytes)
		If Bytes<=0 then Exit Function

		if Readed + Bytes > TotalBytes Then Bytes = TotalBytes - Readed
		BinaryRead = BinaryStream.Read(Bytes)
	End Function

	Public Property Get TotalBytes
		TotalBytes = BinaryStream.Size
	End Property

	Private Sub Class_Initialize()
		Set BinaryStream = createobject("ADODB.Stream")
		BinaryStream.Type = 1 'Binary
		BinaryStream.Open
		BinaryStream.LoadFromFile "F:\InetPub\Motobit\pureupload\2.txt"
		BinaryStream.Position = 0
		Readed = 0
	End Sub
end Class



%>