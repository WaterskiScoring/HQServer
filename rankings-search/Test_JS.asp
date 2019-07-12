<html>
	

<head>
<script type="text/javascript">


function calcMemberID() {
		// www.usawaterski.org/rankings/test_js.asp
		// www.usawaterski.org/rankings/test_js_orig.asp
		
		var PersonID = '72469';
		var PIDSum = 0;
		var PIDChar = PersonID;
		var PIDLen = PIDChar.length;
		var PIDPtr;
		
 		document.write('<br><br>NEW CALC');
 		document.write('<br><br>PersonID = ' + PersonID );
					
		for(PIDPtr=1; PIDPtr<=PIDLen; PIDPtr +=2 ) {
					
					document.write('<br><br><br>COUNTER PIDPtr = ' + PIDPtr);

					CurrSum = PIDSum;
					PIDSum = CurrSum + (3 * Number(PIDChar.substr(PIDPtr-1,1)));

					document.write('<br>PIDSum = ' + PIDSum);

				if(PIDPtr + 1 <= PIDLen) { 

						var NewSum = PIDSum;
						document.write('<br>INSIDE: Line 36 - PIDChar.substr(PIDStrPlus1,1) = ' + PIDChar.substr(PIDPtr,1));
						
						PIDSum = NewSum + Number(PIDChar.substr(PIDPtr,1));

						//document.write('<br>PIDChar.substr(PIDStrPlus1,1) = ' + PIDChar.substr(PIDStrPlus1,1));
						document.write('<br> AFTER Calc: PIDSum = ' + PIDSum);
					}	
				}

	
		var FirstPart = String(100 - PIDSum);
		document.write('<br>FirstPart = ' + FirstPart );
		
		// -- Javascript starts at zero --
		LenPID100MinusPIDSum = FirstPart.length;
		document.write('<br>PID100MinusPIDSum = ' + LenPID100MinusPIDSum );
		var SecondPart = FirstPart.substr(LenPID100MinusPIDSum-1,1);
		document.write('<br>SecondPart = ' + SecondPart );
		
		var ThirdPart = String(100000000 + Number(PersonID));
		document.write('<br>ThirdPart = ' + ThirdPart );
		
		var ThirdPartLen = ThirdPart.toString().length;
		var FourthPart = ThirdPart.substr(1,8);
		document.write('<br>FourthPart = ' + FourthPart );
		
		var MemberID = SecondPart + FourthPart;
		document.write('<br>MemberID = ' + MemberID );
		//PersonIDwritewChkDgt = right(100-PIDSum,1) & Right(100000000+PersonID,8);

	}



</script>
</head>


<body onload="javascript:calcMemberID();">

</body>
</html>



