<%
Function getSalt(intLen)
	Dim strSalt
	Dim intIndex, intRand

	If Not IsNumeric(intLen) Then
		getSalt = "00000000"
		exit function
	ElseIf CInt(intLen) <> CDbl(intLen) Or CInt(intLen) < 1 Then
		getSalt = "00000000"
		exit function
	End If

	Randomize

	For intIndex = 1 to CInt(intLen)
		intRand = CInt(Rnd * 1000) Mod 16
		strSalt = strSalt & getDecHex(intRand)
	Next
	
	getSalt = strSalt

End Function


Function HashEncode(strSecret)
    Dim strEncode, strH(4)
    Dim intPos
    
    
    If len(strSecret) = 0 or len(strSecret) >= 2^61 then
		HashEncode = "0000000000000000000000000000000000000000"
		exit function
    end if
    
    
    strH(0) = "FB0C14C2"
    strH(1) = "9F00AB2E"
    strH(2) = "991FFA67"
    strH(3) = "76FA2C3F"
    strH(4) = "ADE426FA"
    
    For intPos = 1 to len(strSecret) step 56
		
		strEncode = Mid(strSecret, intPos, 56) 'get 56 character chunks
		strEncode = WordToBinary(strEncode) 'convert to binary
		strEncode = PadBinary(strEncode) 'make it 512 bites
		strEncode = BlockToHex(strEncode) 'convert to hex value
		
		strEncode = DigestHex(strEncode, strH(0), strH(1), strH(2), strH(3), strH(4))

		strH(0) = HexAdd(left(strEncode, 8), strH(0))
		strH(1) = HexAdd(mid(strEncode, 9, 8), strH(1))
		strH(2) = HexAdd(mid(strEncode, 17, 8), strH(2))
		strH(3) = HexAdd(mid(strEncode, 25, 8), strH(3))
		strH(4) = HexAdd(right(strEncode, 8), strH(4))
		
    Next
    
    'This is the final Hex Digest
    HashEncode = strH(0) & strH(1) & strH(2) & strH(3) & strH(4)
    
End Function



Function HexToBinary(btHex)

    Select Case btHex
    Case "0"
        HexToBinary = "0000"
    Case "1"
        HexToBinary = "0001"
    Case "2"
        HexToBinary = "0010"
    Case "3"
        HexToBinary = "0011"
    Case "4"
        HexToBinary = "0100"
    Case "5"
        HexToBinary = "0101"
    Case "6"
        HexToBinary = "0110"
    Case "7"
        HexToBinary = "0111"
    Case "8"
        HexToBinary = "1000"
    Case "9"
        HexToBinary = "1001"
    Case "A"
        HexToBinary = "1010"
    Case "B"
        HexToBinary = "1011"
    Case "C"
        HexToBinary = "1100"
    Case "D"
        HexToBinary = "1101"
    Case "E"
        HexToBinary = "1110"
    Case "F"
        HexToBinary = "1111"
    Case Else
        HexToBinary = "2222"
    End Select
End Function

Function BinaryToHex(strBinary)

    Select Case strBinary
    Case "0000"
        BinaryToHex = "0"
    Case "0001"
        BinaryToHex = "1"
    Case "0010"
        BinaryToHex = "2"
    Case "0011"
        BinaryToHex = "3"
    Case "0100"
        BinaryToHex = "4"
    Case "0101"
        BinaryToHex = "5"
    Case "0110"
        BinaryToHex = "6"
    Case "0111"
        BinaryToHex = "7"
    Case "1000"
        BinaryToHex = "8"
    Case "1001"
        BinaryToHex = "9"
    Case "1010"
        BinaryToHex = "A"
    Case "1011"
        BinaryToHex = "B"
    Case "1100"
        BinaryToHex = "C"
    Case "1101"
        BinaryToHex = "D"
    Case "1110"
        BinaryToHex = "E"
    Case "1111"
        BinaryToHex = "F"
    Case Else
        BinaryToHex = "Z"
    End Select
End Function

Function WordToBinary(strWord)
	Dim strTemp, strBinary 
	Dim intPos

	For intPos = 1 To Len(strWord)
	    strTemp = Mid(strWord, cint(intPos), 1)
	    strBinary = strBinary & IntToBinary(Asc(strTemp))
	Next

	WordToBinary = strBinary

End Function

Function HexToInt(strHex)
	Dim intNew, intPos, intLen

	intNew = 0
	intLen = CDbl(len(strHex)) - 1
	
	For intPos = CDbl(intLen) to 0 step -1
	    Select Case Mid(strHex, CDbl(intPos) + 1, 1)       
	    Case "0"
			intNew = CDbl(intNew) + (0 * 16^CDbl(intLen - intPos))
	    Case "1"
	        intNew = CDbl(intNew) + (1 * 16^CDbl(intLen - intPos))
	    Case "2"
	        intNew = CDbl(intNew) + (2 * 16^CDbl(intLen - intPos))
	    Case "3"
	        intNew = CDbl(intNew) + (3 * 16^CDbl(intLen - intPos))
	    Case "4"
	        intNew = CDbl(intNew) + (4 * 16^CDbl(intLen - intPos))
	    Case "5"
	        intNew = CDbl(intNew) + (5 * 16^CDbl(intLen - intPos))
	    Case "6"
	        intNew = CDbl(intNew) + (6 * 16^CDbl(intLen - intPos))
	    Case "7"
	        intNew = CDbl(intNew) + (7 * 16^CDbl(intLen - intPos))
	    Case "8"
	        intNew = CDbl(intNew) + (8 * 16^CDbl(intLen - intPos))
	    Case "9"
	        intNew = CDbl(intNew) + (9 * 16^CDbl(intLen - intPos))
	    Case "A"
	        intNew = CDbl(intNew) + (10 * 16^CDbl(intLen - intPos))
	    Case "B"
	        intNew = CDbl(intNew) + (11 * 16^CDbl(intLen - intPos))
	    Case "C"
	        intNew = CDbl(intNew) + (12 * 16^CDbl(intLen - intPos))
	    Case "D"
	        intNew = CDbl(intNew) + (13 * 16^CDbl(intLen - intPos))
	    Case "E"
	        intNew = CDbl(intNew) + (14 * 16^CDbl(intLen - intPos))
	    Case "F"
	        intNew = CDbl(intNew) + (15 * 16^CDbl(intLen - intPos))
		End Select

	Next

	HexToInt = CDbl(intNew)
	
End Function

Function IntToBinary(intNum)

    Dim strBinary, strTemp
    Dim intNew, intTemp
    Dim dblNew
    
    intNew = intNum
    
    Do While intNew > 1
        dblNew = CDbl(intNew) / 2
        intNew = Round(CDbl(dblNew) - 0.1, 0)
        If CDbl(dblNew) = CDbl(intNew) Then
            strBinary = "0" & strBinary
        Else
            strBinary = "1" & strBinary
        End If

    Loop
    
    strBinary = intNew & strBinary
    
    intTemp = Len(strBinary) mod 8
    
    For intNew = intTemp To 7
        strBinary = "0" & strBinary
    Next
    
    IntToBinary = strBinary
    
End Function

Function PadBinary(strBinary)

	Dim intPos, intLen
	Dim strTemp
	    
	intLen = Len(strBinary)
	    
	strBinary = strBinary & "1"
	    
	For intPos = Len(strBinary) To 447
	    strBinary = strBinary & "0"
	Next
	    
	strTemp = IntToBinary(intLen)
	    
	For intPos = Len(strTemp) To 63
	    strTemp = "0" & strTemp
	Next
	    
	strBinary = strBinary & strTemp
	    
	PadBinary = strBinary
	    
End Function

Function BlockToHex(strBinary)

	Dim intPos
	Dim strHex

	For intPos = 1 To Len(strBinary) Step 4
	    strHex = strHex & BinaryToHex(Mid(strBinary, intPos, 4))
	Next

	BlockToHex = strHex

End Function

Function DigestHex(strHex, strH0, strH1, strH2, strH3, strH4)

	Dim strWords(79), adoConst(4), strTemp, strTemp1, strTemp2, strTemp3, strTemp4
	Dim intPos
	Dim strH(4), strA(4), strK(3)

	'Constant hex words are used for encryption, these can be any valid 8 digit hex value
    strK(0) = "5A827999"
    strK(1) = "6ED9EBA1"
    strK(2) = "8F1BBCDC"
    strK(3) = "CA62C1D6"
    
    'Hex words are used in the encryption process, these can be any valid 8 digit hex value
    strH(0) = strH0
    strH(1) = strH1
    strH(2) = strH2
    strH(3) = strH3
    strH(4) = strH4
    
	For intPos = 0 To (len(strHex) / 8) - 1
	    strWords(cint(intPos)) = Mid(strHex, (cint(intPos)*8) + 1, 8)
	Next


	For intPos = 16 To 79
	    strTemp = strWords(cint(intPos) - 3)
	    strTemp1 = HexBlockToBinary(strTemp)
	    strTemp = strWords(cint(intPos) - 8)
	    strTemp2 = HexBlockToBinary(strTemp)
	    strTemp = strWords(cint(intPos) - 14)
	    strTemp3 = HexBlockToBinary(strTemp)
	    strTemp = strWords(cint(intPos) - 16)
	    strTemp4 = HexBlockToBinary(strTemp)
	    strTemp = BinaryXOR(strTemp1, strTemp2)
	    strTemp = BinaryXOR(strTemp, strTemp3)
	    strTemp = BinaryXOR(strTemp, strTemp4)
	    strWords(cint(intPos)) = BlockToHex(BinaryShift(strTemp, 1))
	Next

	strA(0) = strH(0)
	strA(1) = strH(1)
	strA(2) = strH(2)
	strA(3) = strH(3)
	strA(4) = strH(4)

	'Main encryption loop on all 80 hex word positions
	For intPos = 0 To 79
	    strTemp = BinaryShift(HexBlockToBinary(strA(0)), 5)
	    strTemp1 = HexBlockToBinary(strA(3))
	    strTemp2 = HexBlockToBinary(strWords(cint(intPos)))
	    
	    Select Case intPos
	    
	    Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19
	        strTemp3 = HexBlockToBinary(strK(0))
	        strTemp4 = BinaryOR(BinaryAND(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(2))), BinaryAND(BinaryNOT(HexBlockToBinary(strA(1))), _
				HexBlockToBinary(strA(3))))
	    Case 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39
	        strTemp3 = HexBlockToBinary(strK(1))
	        strTemp4 = BinaryXOR(BinaryXOR(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(2))), HexBlockToBinary(strA(3)))
	    Case 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59
	        strTemp3 = HexBlockToBinary(strK(2))
	        strTemp4 = BinaryOR(BinaryOR(BinaryAND(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(2))), BinaryAND(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(3)))), BinaryAND(HexBlockToBinary(strA(2)), _
				HexBlockToBinary(strA(3))))
	    Case 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79
	        strTemp3 = HexBlockToBinary(strK(3))
	        strTemp4 = BinaryXOR(BinaryXOR(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(2))), HexBlockToBinary(strA(3)))
	    End Select
	    
	    strTemp = BlockToHex(strTemp)
	    strTemp1 = BlockToHex(strTemp1)
	    strTemp2 = BlockToHex(strTemp2)
	    strTemp3 = BlockToHex(strTemp3)
	    strTemp4 = BlockToHex(strTemp4)
	    
	    strTemp = HexAdd(strTemp, strTemp1)
	    strTemp = HexAdd(strTemp, strTemp2)
	    strTemp = HexAdd(strTemp, strTemp3)
	    strTemp = HexAdd(strTemp, strTemp4)
	    
	    strA(4) = strA(3)
	    strA(3) = strA(2)
	    strA(2) = BlockToHex(BinaryShift(HexBlockToBinary(strA(1)), 30))
	    strA(1) = strA(0)
	    strA(0) = strTemp
	    
	Next

	DigestHex = strA(0) & strA(1) & strA(2) & strA(3) & strA(4)

End Function

Function HexAdd(strHex1, strHex2)

    Dim intCalc
    Dim strNew
    
    intCalc = 0
    intCalc = CDbl(CDbl(HexToInt(strHex1)) + CDbl(HexToInt(strHex2)))
    Do While CDbl(intCalc) > 2^32
		intCalc = CDbl(intCalc) - 2^32
    Loop
       
    strNew = IntToBinary(CDbl(intCalc))
    Do While Len(strNew) < 32
        strNew = "0" & strNew
    Loop
    strNew = BlockToHex(strNew)
    
    if InStr(strNew, "00") = 1 and len(strNew) = 10 then
		strNew = right(strNew, 8)
    end if
    
    HexAdd = strNew

End Function

Function getHexDec(strHex)

    Select Case strHex
    Case "0"
        getHexDec = 0
    Case "1"
        getHexDec = 1
    Case "2"
        getHexDec = 2
    Case "3"
        getHexDec = 3
    Case "4"
        getHexDec = 4
    Case "5"
        getHexDec = 5
    Case "6"
        getHexDec = 6
    Case "7"
        getHexDec = 7
    Case "8"
        getHexDec = 8
    Case "9"
        getHexDec = 9
    Case "A"
        getHexDec = 10
    Case "B"
        getHexDec = 11
    Case "C"
        getHexDec = 12
    Case "D"
        getHexDec = 13
    Case "E"
        getHexDec = 14
    Case "F"
        getHexDec = 15
    Case Else
        getHexDec = -1
    End Select
End Function

Function getDecHex(strHex)

    Select Case CInt(strHex)
    Case 0
       getDecHex = "0"
    Case 1
       getDecHex = "1"
    Case 2
       getDecHex = "2"
    Case 3
       getDecHex = "3"
    Case 4
       getDecHex = "4"
    Case 5
       getDecHex = "5"
    Case 6
       getDecHex = "6"
    Case 7
       getDecHex = "7"
    Case 8
       getDecHex = "8"
    Case 9
       getDecHex = "9"
    Case 10
       getDecHex = "A"
    Case 11
       getDecHex = "B"
    Case 12
       getDecHex = "C"
    Case 13
       getDecHex = "D"
    Case 14
       getDecHex = "E"
    Case 15
       getDecHex = "F"
    Case Else
       getDecHex = "Z"
    End Select
End Function

Function BinaryShift(strBinary, intPos)

    BinaryShift = Right(strBinary, Len(strBinary) - cint(intPos)) & _
		Left(strBinary, cint(intPos))

End Function

Function BinaryXOR(strBin1, strBin2)
    Dim strBinaryFinal
    Dim intPos
    
    For intPos = 1 To Len(strBin1)
        Select Case Mid(strBin1, cint(intPos), 1)
        
        Case Mid(strBin2, cint(intPos), 1)
            strBinaryFinal = strBinaryFinal & "0"
        Case Else
            strBinaryFinal = strBinaryFinal & "1"
        End Select
    Next
    
    BinaryXOR = strBinaryFinal
    
End Function

Function BinaryOR(strBin1, strBin2)
    Dim strBinaryFinal
    Dim intPos
    
    For intPos = 1 To Len(strBin1)
        If Mid(strBin1, cint(intPos), 1) = "1" Or Mid(strBin2, cint(intPos), 1) = "1" Then
            strBinaryFinal = strBinaryFinal & "1"
        Else
            strBinaryFinal = strBinaryFinal & "0"
        End If
    Next
    
    BinaryOR = strBinaryFinal
End Function

Function BinaryAND(strBin1, strBin2)
    Dim strBinaryFinal
    Dim intPos
    
    For intPos = 1 To Len(strBin1)
        If Mid(strBin1, cint(intPos), 1) = "1" And Mid(strBin2, cint(intPos), 1) = "1" Then
            strBinaryFinal = strBinaryFinal & "1"
        Else
            strBinaryFinal = strBinaryFinal & "0"
        End If
    Next
    
    BinaryAND = strBinaryFinal
End Function

Function BinaryNOT(strBinary)
    Dim strBinaryFinal
    Dim intPos
    
    For intPos = 1 To Len(strBinary)
        If Mid(strBinary, cint(intPos), 1) = "1" Then
            strBinaryFinal = strBinaryFinal & "0"
        Else
            strBinaryFinal = strBinaryFinal & "1"
        End If
    Next
    
    BinaryNOT = strBinaryFinal
    
End Function

Function HexBlockToBinary(strHex)
    Dim intPos
    Dim strTemp
    
    For intPos = 1 To Len(strHex)
        strTemp = strTemp & HexToBinary(Mid(strHex, cint(intPos), 1))
    Next
    
    HexBlockToBinary = strTemp
    
End Function
%>