<%
Function MyIsNumeric(invalue)
	MyIsNumeric = True
	If IsNull(invalue) then
		MyIsNumeric = False
	ElseIf invalue = "" then
		MyIsNumeric = False
	Else
		Dim y
		For y = 1 to Len(invalue)
			If asc(Mid(invalue, y, 1)) >= 48 AND asc(Mid(invalue, y, 1)) <= 57 then
			'If Mid(invalue, y, 1) >= 0 AND Mid(invalue, y, 1) <= 9 then
				'Do nothing
			Else
				MyIsNumeric = False
				Exit Function
			End If
		Next
	End If
End Function

Function NumOnly(strInput)
	if len(strInput) = 0 then
		NumOnly = 0
	end if
	For i = 1 to Len(strInput)
		If isNumeric(Mid(strInput, i, 1)) then
			NumOnly = NumOnly & Mid(strInput, i, 1)
		End If
	Next
End Function

Function Encrypt(blnED, intCCNum)
	'Take Input CC and strip non-numeric chars from it.
	stripCC = NumOnly(intCCNum)
	'If blnED is true, then encrypt, else, decrypt
	Dim outNum
	If blnED then
		'encrypt
		'Response.Write "Encrypt<br>"
		For i = 1 to Len(stripCC)
			Dim tempNum
			tempNum = cint(Mid(stripCC, i, 1)) + 4
			'Response.Write "i = " & i & ", "
			'Response.Write "tempNum = " & tempNum & "<br>"
			If Len(tempNum) = 2 then
				outNum = cstr(outNum)& cstr(Mid(tempNum, 2, 1))
			ElseIf Len(tempNum) = 1 then
				outNum = cstr(outNum) & cStr(tempNum)
			Else
				'Response.Write outNum
				'Response.Write " Error"
				'Response.End
			End If
		Next
	Else
		'Decrypt
		For i = 1 to Len(stripCC)
			Dim tempDec
			tempDec = cint(Mid(stripCC, i, 1)) - 4
			Dim tempLT
			'Response.Write "i = " & i & ", "
			'Response.Write "tempDec = " & tempDec & "<br>"
			If tempDec < 0 then
				tempLT = "1" & Mid(stripCC, i, 1)
				tempLT = tempLT - 4
				outNum = cstr(outNum) & cstr(tempLT)
			Else
				outNum = cstr(outNum) & cstr(tempDec)
			End If
		Next
	End If
	Encrypt = outNum
End Function

Sub SendEmail(strTo, strCC, strBCC, strSubject, strBody)
	Dim objCDO
	Set objCDO = Server.CreateObject("CDONTS.NewMail")
	objCDO.From = "Cobra Errors <alex@epolk.net>"
	objCDO.To = strTo
	objCDO.CC = strCC
	objCDO.BCC = strBCC
	objCDO.Subject = strSubject
	objCDO.Body = strBody
	objCDO.Send
	Set objCDO = Nothing
End Sub

Function CheckTwoDigitDate(input)
	If len(input) = 1 then
		CheckTwoDigitDate = "0" & input
	ElseIf len(input) = 2 then
		CheckTwoDigitDate = input
	ElseIf len(input) > 2 then
		CheckTwoDigitDate = Right(input, 2)
	Else
		CheckTwoDigitDate = "00"
	End If
End Function

%>




