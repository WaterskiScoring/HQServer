
<html>
<body>
<form action="http://www.google.com/drivingdirections"> 
<br><br>
This is a test of the Directions function
<br><br>

<%
sSiteAddress="1251 Holy Cow Road, Polk City, FL"
%>

<input type="hidden" name="ie" value="UTF8"> 
<input type="hidden" name="f" value="d"> Start Address
 <input type="text" style="width:20em" size="20" name="saddr" tabindex="1" maxlength="2048"/> 
<br>e.g. 1050 S Lake Sybelia Dr, Maitland, FL

 <!-- Code meeting location, date and time below --> 
 <input type="hidden" name="daddr" value="<%=sSiteAddress%>"

 <br>
<input type="submit" value="Get Directions" /> 
</form> 
</body> 
</html> 

