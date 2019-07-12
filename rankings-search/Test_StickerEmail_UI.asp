

<!DOCTYPE html>
<html>
<body>

<h2>Using the XMLHttpRequest Object</h2>

<%
	



%>
<div id="demo">
<button type="button" onclick="loadXMLDoc('18S110R','800074054','S')">Post Send Email</button>
<br>
<br>
<a href="javascript:loadXMLDoc('18S110R','800074054','S')">Post Send Email</a>

</div>



<script>
	function loadXMLDoc(stid, smid, sevt) {
  	
  	var stid = stid
  	var smid = smid
  	var sevt = sevt
  	
  	var PostURL = 'Personal_Best_Recording.asp?stid=' + stid + '&smid=' + smid + '&sevt=' + sevt;
  	// alert('URL = ' + PostURL);
  	var xhttp = new XMLHttpRequest();
  	xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
	
				// alert('this.responseText = ' + this.responseText);
				var responsetxxt =  this.responseText;
				// var xmlDoc = this.responseXML;
				// Set xmlList = xmlDoc.getElementsByTagName("result");

				parser = new DOMParser();
				xmlDoc = parser.parseFromString(responsetxxt,"text/xml");

				var sMemberID = xmlDoc.getElementsByTagName("memberid")[0].childNodes[0].nodeValue;
				var sTourID = xmlDoc.getElementsByTagName("tourid")[0].childNodes[0].nodeValue;
				var sEventName = xmlDoc.getElementsByTagName("eventname")[0].childNodes[0].nodeValue;
				var sScore = xmlDoc.getElementsByTagName("score")[0].childNodes[0].nodeValue;
				var sUnits = xmlDoc.getElementsByTagName("units")[0].childNodes[0].nodeValue;
				var scoreexists = xmlDoc.getElementsByTagName("scoreexists")[0].childNodes[0].nodeValue;

				if (scoreexists == 'N') {
						messageBody = 'Your request has been recorded for the score of ' + sScore + ' ' + sUnits +' in ' + sEventName + ' from sanction ' + sTourID + '. Please look for a confirmation email.'
				}
				else {
						messageBody = 'Thank you.  But, a previous request was received for the score of ' + sScore + ' ' + sUnits +' in ' + sEventName + ' from sanction ' + sTourID + '.'	
				}	
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

</body>
</html>
