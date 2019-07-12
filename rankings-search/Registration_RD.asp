<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<%


sTourID=LEFT(request("sTourID"),6)

DefineTourVariables_New

'response.write("<br>sPayPalOK="&sPayPalOK)
'response.write("<br>sPayPalAct="&sPayPalAct)
'response.write("<br>sUseOLReg="&sUseOLReg)
'response.write("<br>sOLR_PD="&sOLR_PD)

'response.write("<br>TESTING")


IF sPayPalOK=false OR sPayPalOK=0 OR TRIM(sPayPalAct)="" OR sUseOLReg=0 OR sOLR_PD=0 THEN 
'response.write("Opt 1")
'response.end

	response.redirect("/rankings/View-TournamentsHQ.asp?pvar=TourInfo&TourID="&sTourID&"&rg=1")
ELSE
'response.write("Opt 2")
'response.end

	response.redirect("/rankings/registration.asp?sTourID="&sTourID)
END IF 

%>

