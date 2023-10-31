<%
Public UVowels

'-------------------------------------------------------------------------------------
Sub InitUnicode()  
   Dim TStr

   ' Initialise the list of Unicode Vowels, 67 lowerCase followed by 67 Uppercase
   ' Note that by using the Function chrW, the &HE1 Unicode character is stored internally
   ' as &HE100 for a String character

   TStr = TStr & ChrW(&HE1) & ChrW(&HE0) & ChrW(&H1EA3) & ChrW(&HE3) & ChrW(&H1EA1) & ChrW(&H103) & ChrW(&H1EAF) & ChrW(&H1EB1) & ChrW(&H1EB3) & ChrW(&H1EB5) & ChrW(&H1EB7) & ChrW(&HE2) & ChrW(&H1EA5) & ChrW(&H1EA7) & ChrW(&H1EA9) & ChrW(&H1EAB) & ChrW(&H1EAD) & ChrW(&HE9) & ChrW(&HE8) & ChrW(&H1EBB) 
   TStr = TStr & ChrW(&H1EBD) & ChrW(&H1EB9) & ChrW(&HEA) & ChrW(&H1EBF) & ChrW(&H1EC1) & ChrW(&H1EC3) & ChrW(&H1EC5) & ChrW(&H1EC7) & ChrW(&HED) & ChrW(&HEC) & ChrW(&H1EC9) & ChrW(&H129) & ChrW(&H1ECB) & ChrW(&HF3) & ChrW(&HF2) & ChrW(&H1ECF) & ChrW(&HF5) & ChrW(&H1ECD) & ChrW(&HF4) & ChrW(&H1ED1) 
   TStr = TStr & ChrW(&H1ED3) & ChrW(&H1ED5) & ChrW(&H1ED7) & ChrW(&H1ED9) & ChrW(&H1A1) & ChrW(&H1EDB) & ChrW(&H1EDD) & ChrW(&H1EDF) & ChrW(&H1EE1) & ChrW(&H1EE3) & ChrW(&HFA) & ChrW(&HF9) & ChrW(&H1EE7) & ChrW(&H169) & ChrW(&H1EE5) & ChrW(&H1B0) & ChrW(&H1EE9) & ChrW(&H1EEB) & ChrW(&H1EED) & ChrW(&H1EEF) 
   TStr = TStr & ChrW(&H1EF1) & ChrW(&HFD) & ChrW(&H1EF3) & ChrW(&H1EF7) & ChrW(&H1EF9) & ChrW(&H1EF5) & ChrW(&H111) & ChrW(&HC1) & ChrW(&HC0) & ChrW(&H1EA2) & ChrW(&HC3) & ChrW(&H1EA0) & ChrW(&H102) & ChrW(&H1EAE) & ChrW(&H1EB0) & ChrW(&H1EB2) & ChrW(&H1EB4) & ChrW(&H1EB6) & ChrW(&HC2) & ChrW(&H1EA4) 
   TStr = TStr & ChrW(&H1EA6) & ChrW(&H1EA8) & ChrW(&H1EAA) & ChrW(&H1EAC) & ChrW(&HC9) & ChrW(&HC8) & ChrW(&H1EBA) & ChrW(&H1EBC) & ChrW(&H1EB8) & ChrW(&HCA) & ChrW(&H1EBE) & ChrW(&H1EC0) & ChrW(&H1EC2) & ChrW(&H1EC4) & ChrW(&H1EC6) & ChrW(&HCD) & ChrW(&HCC) & ChrW(&H1EC8) & ChrW(&H128) & ChrW(&H1ECA) 
   TStr = TStr & ChrW(&HD3) & ChrW(&HD2) & ChrW(&H1ECE) & ChrW(&HD5) & ChrW(&H1ECC) & ChrW(&HD4) & ChrW(&H1ED0) & ChrW(&H1ED2) & ChrW(&H1ED4) & ChrW(&H1ED6) & ChrW(&H1ED8) & ChrW(&H1A0) & ChrW(&H1EDA) & ChrW(&H1EDC) & ChrW(&H1EDE) & ChrW(&H1EE0) & ChrW(&H1EE2) & ChrW(&HDA) & ChrW(&HD9) & ChrW(&H1EE6) 
   TStr = TStr & ChrW(&H168) & ChrW(&H1EE4) & ChrW(&H1AF) & ChrW(&H1EE8) & ChrW(&H1EEA) & ChrW(&H1EEC) & ChrW(&H1EEE) & ChrW(&H1EF0) & ChrW(&HDD) & ChrW(&H1EF2) & ChrW(&H1EF6) & ChrW(&H1EF8) & ChrW(&H1EF4) & ChrW(&H110) 

   UVowels = TStr  ' Assign to the Unicode Vowel list

End Sub 

'--------------------------------------------------------------------------------------

Function IsUnicode(ch)
	IsUnicode = (InStr(UVowels, Ch) > 0) 
End Function

'--------------------------------------------------------------------------------

Function WriteUnicodeOnWeb(s)
	Dim LenOfSt
	Dim sTemp
	Dim i
	Dim ch
	LenOfSt=len(s)
	sTemp =""
	If (isNull(s)) or (lenofst = 0) Then
		WriteUniCodeOnWeb = "&nbsp;"
		Exit Function
	Else
		For i=1 to  LenOfSt
			ch=Mid(s,i,1)
			If IsUnicode(ch) Then
				sTemp =sTemp+"&#"+CStr( ascw(ch))+";"
			Else
				sTemp = sTemp+ch
			End If
		Next
		WriteUnicodeOnWeb = sTemp
	End If
			
End Function
%>