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





BuildCustomJavascript





' ActionURL = "http://usawaterski.org/rankings/Test_MemberValidate_POST_Tester.asp?action=post"
ActionURL = "http://usawaterski.org/rankings/Web_Services/MemberValidate.asp?action=post"

ClubAccessCode = "MAC001"


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


 			<div class="rankingsbody" style="width:97%; text-align:center; background-color:<%= scolor %>; border:0px solid white; padding:0px 2px 0px 3px; margin:0px 0px 0px 2px;">
					<a title="Test Member Validate" href="javascript:loadXMLDoc('<%=ClubAccessCode%>','<%=sMemberID%>')">
						<span class="span95" style="text-align:center; color:red; font-size:9pt; font-weight:normal;">** TEST VALIDATION **<br>Click here to Test the Member Validation</span>
					</a>
				</div>		    
<%
		    










' ------------------------
  SUB Post_Test
' ------------------------

  
' sMemberID = "000001151"
'sMemberID = Request("sMemberID")
ClubAccessCode = "MAC001"

url = "http://usawaterski.org/rankings/Test_MemberValidate.asp"
MemberRequest = "<USAWaterski><MemberStatus><ClubAccessCode>"&ClubAccessCode&"</ClubAccessCode><MemberID>"&sMemberID&"</MemberID></MemberStatus></USAWaterski>"

Set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
xmlhttp.Open "POST", url, false
xmlhttp.setRequestHeader "Content-Type", "text/xml" 
xmlhttp.send MemberRequest

END SUB





' ---------------------------
  SUB BuildCustomJavascript
' ---------------------------  

%>
<script>
	function loadXMLDoc(cac, mid) {
  	
  	// alert('Line 122');
  	var ClubAccessCode = cac
  	var sMemberID = mid
	
		// rankings/Web_Services/MemberValidate.asp
		var url = '/rankings/Web_Services/MemberValidate.asp?';
		var MemberRequest = '<USAWaterski><MemberStatus><ClubAccessCode>' + ClubAccessCode + '</ClubAccessCode><MemberID>' + sMemberID + '</MemberID></MemberStatus></USAWaterski>';
  	
  	// var PostURL = 'Personal_Best_Recording.asp?stid=' + stid + '&smid=' + smid + '&sevt=' + sevt;
  	var PostURL = url + MemberRequest;
  	alert('URL = ' + PostURL);
  	var xhttp = new XMLHttpRequest();
  	xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
	
				alert('this.responseText = ' + this.responseText);
				var responsetxxt =  this.responseText;
				// var xmlDoc = this.responseXML;
				// Set xmlList = xmlDoc.getElementsByTagName("result");

				parser = new DOMParser();
				xmlDoc = parser.parseFromString(responsetxxt,"text/xml");

				var sMemberID = xmlDoc.getElementsByTagName("memberid")[0].nodeValue;
				var sMemberFound = xmlDoc.getElementsByTagName("sMemberFound")[0].childNodes[0].nodeValue;
				// var sEventName = xmlDoc.getElementsByTagName("eventname")[0].childNodes[0].nodeValue;
				// var sScore = xmlDoc.getElementsByTagName("score")[0].childNodes[0].nodeValue;
				// var sUnits = xmlDoc.getElementsByTagName("units")[0].childNodes[0].nodeValue;
				// var scoreexists = xmlDoc.getElementsByTagName("scoreexists")[0].childNodes[0].nodeValue;
				// var emailexists = xmlDoc.getElementsByTagName("emailexists")[0].childNodes[0].nodeValue;

				


				messageBody = 'sMemberID returned:' + sMemberID + ' - sMemberFound:' + sMemberFound;
				
				// -- Displays notice to user --
				alert(messageBody);
				
				// alert('MemberID = ' + sMemberID);
				// alert('TourID = ' + sTourID);
				// alert('Event = ' + sEvent);
				// alert('Score = ' + sScore);
				// alert('scoreexists = ' + scoreexists);

    	}
  	};
  	xhttp.open("GET", "" + PostURL + "", true);
  	xhttp.send();
	}
</script>
<%
	

END SUB  



%>
