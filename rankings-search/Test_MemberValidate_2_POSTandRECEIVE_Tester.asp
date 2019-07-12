<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%

Dim sMemberID


testMemberID = "200156835"
sMemberID = Request("sMemberID")
IF TRIM(sMemberID)="" THEN sMemberID = testMemberID

response.write("<br>TRUE = ")
response.write(TRIM(Request("sMemberID"))="")
response.write("<br>sMemberID = " & Request("sMemberID"))


action = Request("action")

SELECT CASE action
		CASE "post"
				Post_Test
		CASE ELSE
				' -- Do Nothing 
END SELECT 











ActionURL = "http://usawaterski.org/rankings/Test_MemberValidate_2_POSTandRECEIVE_Tester.asp?action=post"


%>
<br><br><br><br><br>
		    <table class="innertable" align="center" width=50%>
					<tr>
						<th align="center" height="30px">
							<font size="4" color="white"><b>Member Validation Testing</b></font>
						</th>	
					</tr>
					<tr>
						<td align="center" height="100px">
								<font size="3" color="black">This page tests the XML POST to Validate Member status</font>
						</td>
					</tr>
					<tr>
						<td align="center" height="100px">
								<font size="3" color="black">Test Status: Y</font>
						</td>
					</tr>
					<tr>
			      
			      <form action=<%=ActionURL%> method=post name="MVForm">
							<input type="text" name="sMemberID" value="<%=sMemberID%>">

	        	  <td colspan=3 align="center" valign="middle" height="50px">
	        	  	<input type="submit" name="Submit" value="Submit" style="width:9em; height:2.5em;" title="Submit this form">
			  			</td>
		      	</form>
		 			</tr>
		    </table>
<%
		    










' ------------------------
  SUB Post_Test
' ------------------------

  
' sMemberID = "000001151"
'sMemberID = Request("sMemberID")

url = "http://usawaterski.org/rankings/Test_MemberValidate.asp"
MemberRequest = "<USAWaterski><MemberStatus><ClubAccessCode>WBP001</ClubAccessCode><MemberID>"&sMemberID&"</MemberID></MemberStatus></USAWaterski>"

Set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
xmlhttp.Open "POST", url, false
xmlhttp.setRequestHeader "Content-Type", "text/xml" 
xmlhttp.send MemberRequest

'server.Createobject("XMLHttpRequest")
'Dim xmlhttp
' Set xmlhttp = server.CreateObject("MSXML2.DOMDocument.3.0")
'Set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
'Set xmlhttp = new XMLHttpRequest()
'xmlhttp = new XMLHttpRequest()
'xmlhttp.open "GET", url, false
'xmlhttp.setRequestHeader "Content-Type",  "text/xml"
'xmlhttp.send(null)
'xmlhttp.send MemberRequest
'alert(xmlhttp.responseXML)

Set xmlDoc=CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
xmlDoc.load(Request)



MemberID=""
MemberFound="no"
MemberExpireDate=""
MembershipType=""
CanSkiTour="no"


Set xmlList = xmlDoc.getElementsByTagName("MemberStatus")
NumNodes=xmlDoc.getElementsByTagName("MemberStatus").length

FOR Each xmlItem In xmlList
		' -- MemberFound = LCASE(xmlItem.childNodes(0).text)
		MemberID = xmlDoc.getElementsByTagName("MemberID")(0).text		
		MemberFound = xmlDoc.getElementsByTagName("MemberFound")(0).text
		MemberFirst = xmlDoc.getElementsByTagName("MemberFirst")(0).text
		MemberLast = xmlDoc.getElementsByTagName("MemberLast")(0).text
		MembershipType = xmlDoc.getElementsByTagName("MembershipType")(0).text
		MemberExpireDate = xmlDoc.getElementsByTagName("MemberExpireDate")(0).text
		CanSkiTour = xmlDoc.getElementsByTagName("CanSkiTour")(0).text
NEXT

response.write("TEST RESPONSE = " & MemberExpireDate)

END SUB






%>
