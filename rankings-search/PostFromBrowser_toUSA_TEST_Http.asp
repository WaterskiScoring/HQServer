<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?php header('Access-Control-Allow-Origin: *'); ?>
<html>

<head>
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
	<script type="text/javascript">

		function loadDoc() {
  			
  			var xhttp = new XMLHttpRequest();
  			xhttp.onreadystatechange = function() { // this function handles the response
    		if (xhttp.readyState == 4 && xhttp.status == 200) {
    				// get as XML object
						xmlDoc = xhttp.responseXML;
						//direct output here
    				}
  				};


					// -- Sends the request
  				var ClubAccessCode = "MAC001"
					var sMemberID="000001151"
					var url = "http://usawaterski.org/rankings/MemberStatus_WebService_Dev.asmx"
					var url = "http://usawaterski.org/rankings/Test_MemberValidate.asp"
					var MemberStatusXML = "<USAWaterski><MemberStatus><ClubAccessCode>" + ClubAccessCode + "</ClubAccessCode><MemberID>" + sMemberID + "</MemberID></MemberStatus></USAWaterski>"
					
					alert('MemberStatus XML = ' + MemberStatusXML);
			    xhttp.open("POST", url, true);
			    xhttp.setRequestHeader("Content-Type", "text/xml"); 
					xhttp.send(MemberStatusXML);

				}

	</script>


</head>
<body>
	<br>
	<br>
	<br>
	<table width="300px" align="center">
		<tr>
		<td height="200px" align="center">Click button to test post</td> 
		</tr>	
		<tr>
			<td height="200px" align="center">
				<input type="button" name="testpost" value="Test Post" onclick="javascript:loadDoc();">
			</td> 
		</tr>	
	</table>
</body>	

</html>
