<%

'Dim oSoapClient, mySoapClient

' --- Initialize the MSSOAP.SoapClient
'Set mySoapClient = Server.CreateObject("MSSOAP.SoapClient30")
'mySoapClient.ClientProperty("ServerHTTPRequest") = True

' --- Associate the WebService with the SoapClient
' --- The SoapClient object association needs the path of the WSDL file, and the Webservice's name
'mySoapClient.mssoapinit("http://www.usawaterski.org/sanctions/webservices/swiftservices.asmx?WSDL")


' ----------------------------------------------------------------------
' --- Populate sFunctionName with the results from the web service   ---
' ----------------------------------------------------------------------


' --- FROM Jim Meis - Dec 2009 ---
' --- The difference between the recordsets returned by the function names provided by GetOLRFunction and GetSearchFunction
' --- is in the where clause on the sanction side. 
' --- GetOLRFunction uses Where TournAppID = ?? and TSanType<> 2 and TStatus > 1
' --- 	It does not return advertisements and returns only tournaments which have achieved region approval.
' --- GetSearchFunction uses Where TournAppID = ?? and TYear > 2009



' --------------------------------------------------------------------------------------------------------------
' --- Define and call the function ---
'sFunctionName = "dbo."&mySoapClient.GetSearchFunction(sTournAppID)
'sFunctionName = "dbo."&mySoapClient.GetLJSearchFunction(sTournAppID)

' --- This format works but as of 1-15-2009 going with SEARCH function only as it returns ALL not just approved
' sFunctionName = "dbo."&mySoapClient.GetOLRFunction(sTournAppID)
' --------------------------------------------------------------------------------------------------------------

'Set mySoapClient = Nothing

%>

<html>
<head>
<title>Calling a webservice from classic ASP</title>
</head>
<body>
<%

Dim xmlhttp
Dim DataToSend
sTournAppID = "13S154"
DataToSend="TournAppID="&sTournAppID
Dim postUrl

'postUrl = "http://usawaterski.org/sanctions/webservices/GetLJSearchFunction"
postUrl = "http://www.usawaterski.org/sanctions/webservices/swiftservices.asmx/GetLJSearchFunction"

Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
xmlhttp.Open "POST",postUrl,false
xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
xmlhttp.send DataToSend

sXML = (xmlhttp.responseText)

response.write "Starting XML " & sXML & "</br>"

'This is what it looks like returned
'=<?xml version="1.0" encoding="utf-8"?> <string xmlns="http://usawaterski.org/sanctions/webservices">fn_LJsearch2010XTournAppID('13S154')</string>

'First look for the word error or invalid
if instr(sXML,"error") > 0 or instr(sXML,"invalid") > 0 then
else
	'First, find the string WebServices
	LocationofWordWebservies = instr(sXML,"webservices")
	if LocationofWordWebservies > 0 then
		sXML = mid(sXML, LocationofWordWebservies + 13,99)
		'now remove the last part
		LocationofwordString = instr(sXML,"</string>")
		if LocationofwordString > 0 then
			sXML = left(sXML, LocationofwordString)
		end if
	end if
end if
response.write "HHH" & sXML & "HHH </br>"
%>
