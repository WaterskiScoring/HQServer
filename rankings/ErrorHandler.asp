<%

' ----------------------------
' --- Initialize variables ---
' ----------------------------

Dim strErrorMessage
Dim bolErrors
strErrorMessage = ""	' --- The error messages for tech. support
bolErrors = False	'Have we found any errors yet?


' ----------------------------
SUB TrapError(strError)
' ----------------------------

  bolErrors = True	' --- Egad, we've found an error!
  strErrorMessage = strErrorMessage & strError & ", "

END SUB


' ----------------------------
 SUB ProcessErrors()
' ----------------------------
  IF bolErrors THEN
    ' --- Send the email
    Dim objCDO
    Set objCDO = Server.CreateObject("CDO.Message")

    objCDO.To = "Mark@productdesign-biz.com"
	objCDO.bcc = "eweiss@metisentry.com"
    objCDO.From = "RankingsErrors@usawaterski.org"
    objCDO.Subject = "ADO Error"
    ' objCDO.HTMLBody = eBody	
    objCDO.Body = "At " & Now & " the following errors occurred on " & _
		  "the page " & Request.ServerVariables("SCRIPT_NAME") & _
		  ": " & _
                  chr(10) & chr(10) & strErrorMessage


    objCDO.Send

    Set objCDO = Nothing

  END IF

END SUB  

%>