<%



' -- Read URL query string params and store in session variables --
' Session("m_param")=request("m")
Session("JobID")=request("j")
Session("ListID")=request("l")
Session("SubID")=request("sfmc_sub")
Session("LinkID")=request("u")
Session("BatchID")=request("jb")
Session("MID")=request("mid")

%>
<br><br><br>
<%
' response.write("<br>m_param = " &request("m")) 
response.write("<br>JobID = " &request("j")) 
response.write("<br>ListID = " &request("l")) 
response.write("<br>SubID = " &request("sfmc_sub")) 
response.write("<br>LinkID = " &request("u")) 
response.write("<br>BatchID = " &request("jb")) 
response.write("<br>MID = " &request("mid")) 
%>
<br><br>
<center><a href="Test_ConversionTracking_CONFIRMATION.asp"><h3>Click to Test Conversion Tracking</h3></a></center>

