<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_registration2.asp"-->
<%




sTourID="13M967"
'DefineTourVariables_New

sFunctionName="ERROR-Test"
IF InStr(LCASE(sFunctionName),"error") > 0 or InStr(sFunctionName,"invalid") > 0 THEN
		response.write("<br>Found Error")
ELSE
		response.write("<br>Didn't Find Error")
END IF				


%>