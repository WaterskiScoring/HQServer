<%@ Language=VBScript %>
<!--#include virtual="/rankings/settingsHQ.asp"-->

<%
'Relies on SOAP3 toolkit which is installed on the server.
'dim the input and output variables
dim sTournAppID, sFunctionName 
'supply a TournAppID
'sTournAppID = "09E076"
sTournAppID = "08S999"

'Declare a MSSOAP.SoapClient
Dim oSoapClient, mySoapClient

'Initialize the MSSOAP.SoapClient
Set mySoapClient = Server.CreateObject("MSSOAP.SoapClient30")
mySoapClient.ClientProperty("ServerHTTPRequest") = True


'Associate the WebService with the SoapClient
'The SoapClient object association needs the path of the WSDL file, and the Webservice's name
'It is supposed to be possible to do it with one line of code but I couldn't make that work so I used 2 lines

mySoapClient.mssoapinit("http://www.usawaterski.org/sanctions/webservices/swiftservices.asmx?WSDL")

'Populate sFunctionName with the results from the web service.
sFunctionName = mySoapClient.GetOLRFunction(sTournAppID)
'Clean up
Set mySoapClient = Nothing

response.write("sFunctionName = "&sFunctionName)



sSQL = "SELECT * FROM "&sFunctionName
response.write("<br>sSQL = "&sSQL)
'response.end

OpenConOLReg
set rsTSetUp=Server.CreateObject("ADODB.recordset")
rsTSetUp.open sSQL, sConnectionToOLRegFunction


response.write("<br>EOF = ")
response.write(rsTSetUp.eof)

%>




