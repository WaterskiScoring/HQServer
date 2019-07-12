

<!DOCTYPE html>
<html>
<body>

<h2>Using the XMLHttpRequest Object</h2>

<%
	
sTourID = "17S116C"
sMemberID = "900167964"
sEvent= "T"



%>

<br>
<br>
<br>


</div>

  			<div class="rankingsbody" style="width:97%; text-align:center; background-color:<%= scolor %>; border:0px solid white; padding:0px 2px 0px 3px; margin:0px 0px 0px 2px;">
					<a title="Request Personal Best Decal" href="javascript:loadXMLDoc('<%=sTourID%>','<%=sMemberID%>', '<%=sEvent%>')">
						<span class="span95" style="text-align:center; color:red; font-size:9pt; font-weight:normal;">** PERSONAL BEST **<br>Click here to request your <b>Personal Best</b> sticker</span>
					</a>
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
						messageBody = 'Thank you.  A previous request was received for the score of ' + sScore + ' ' + sUnits +' in ' + sEventName + ' from sanction ' + sTourID + '.'	
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
