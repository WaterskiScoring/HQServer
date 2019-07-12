<%@ Language=VBScript %>
<html>
<%	'This form sets up the where clause for the sql statement used in awsh_rlist.asp to display the tournament list.
'Make sure user is logged in
	If Session("LoggedIn")= false then
		'Response.Write("Not Logged In")
		'Response.End
		session.Abandon
		Response.Redirect ("../../default.html")
		Response.End
	end if
	'Make sure Region Level ID is correct for this page.
	if not Session("HQUser") = true then
		session.Abandon
		Response.Redirect ("../../default.html")
		Response.End
	end if
dim Conn2, rsWorking, sConn, SQL, sStartYr, sSptsGrpID, sLogo, sHeader	
Set Conn2 = Server.CreateObject("ADODB.Connection")
Set rsWorking = Server.CreateObject("ADODB.Recordset")
sConn = Application("PSAConnStr")
SQL = "SELECT DISTINCT Tschedul.TYear FROM Tschedul ORDER BY Tschedul.TYear"
sStartYr = cstr(year(date))
Conn2.Open    sConn    
Set rsWorking = Conn2.Execute (SQL)
	do until rsWorking.EOF
		if cint(rsWorking("TYear")) < cint(sStartYr) then
			sStartYr = cstr(rsWorking("TYear"))
		end if
		rsWorking.MoveNext
	loop
rsWorking.close
set rsWorking = nothing
Conn2.Close
set Conn2 = nothing
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<script language="javascript">

function YearPicked(Field) {
	// alert(Field.options [Field.selectedIndex].text);
	document.setwhere.Year.value = Field.options [Field.selectedIndex].text;
}
function FixRange(local) {
	FR_Msg = ''
	var chk1 = document.setwhere.ckWhere1;
	var chk2 = document.setwhere.ckWhere2;
	var field1 = document.setwhere.Where1;
	var field2 = document.setwhere.Where2;
	YearPicked(document.setwhere.YearPick)
	date1ok = false;
	date2ok = false;
	UseDate1 = false;
	UseDate2 = false;
	if (!chk1.checked) {
		field1.value = '';
	} else {
		date1ok = CheckDate(field1, 'true');
		if (date1ok == false) {
			FR_Msg += '"Scheduled After" Date is Invalid/n';
			chk1.checked = false;
		} else {
			UseDate1 = true;
			L_Value = field1.value ;
			L_Index = L_Value.indexOf('/');
			L_Index2 = L_Value.indexOf('/', L_Index + 1);
			L1_Year = L_Value.substring(L_Index2 + 1, L_Value.length);
			if (document.setwhere.Year.value != L1_Year) { FR_Msg += '"Scheduled After" date must be in Droplist Year.\n'};
		}
	}
	if (!chk2.checked) {
		field2.value = '';
	} else {
		date2ok	= CheckDate(field2, 'true');
		if (date2ok == false) {
			FR_Msg += '"Scheduled Before" Date is Invalid/n';
			chk2.checked = false;
		} else {
			UseDate2 = true;
			L_Value = field2.value ;
			L_Index = L_Value.indexOf('/');
			L_Index2 = L_Value.indexOf('/', L_Index + 1);
			L2_Year = L_Value.substring(L_Index2 + 1, L_Value.length);
			if (document.setwhere.Year.value != L2_Year) { FR_Msg += '"Scheduled Before" date must be in Droplist Year.\n'};
		}
	}
	
	if (UseDate1 == true
		&& UseDate2 == true) {
		if (field1.value > field2.value){ FR_Msg += '"Scheduled Before" date must be same as or later than "Scheduled After" date. \n'};
	} 
	if (local == true) {
		if (FR_Msg.length != 0) {
			alert(FR_Msg);
		}
	} else {
		if (FR_Msg.length == 0) {
			return true;
		} else {
			alert(FR_Msg);
			return false;
		}
	}
}


function FormValidation(theForm) {
	return (FixRange(false));
}

function CheckDate(Field, ShowAlert) {
	L_Msg = '';
	L_Value = Field.value ;
	daysInMonth = new Array(12);
	daysInMonth[1] = 31;
	daysInMonth[2] = 29;
	daysInMonth[3] = 31;
	daysInMonth[4] = 30;
	daysInMonth[5] = 31;
	daysInMonth[6] = 30;
	daysInMonth[7] = 31;
	daysInMonth[8] = 31;
	daysInMonth[9] = 30;
	daysInMonth[10] = 31;
	daysInMonth[11] = 30;
	daysInMonth[12] = 31;

	L_Index = L_Value.indexOf('/');
	L_Index2 = L_Value.indexOf('/', L_Index + 1);
	L_Month = L_Value.substring(0, L_Index);
	L_Day = L_Value.substring(L_Index + 1, L_Index2);
	L_Year = L_Value.substring(L_Index2 + 1, L_Value.length);

if (L_Month.length == 2 && L_Month.substring(0, 1) == 0) {L_Month =
L_Month.substring(1, 2);}
if (L_Day.length == 2 && L_Day.substring(0, 1) == 0) {L_Day =
L_Day.substring(1, 2);}

    if (isNaN(L_Year) || L_Year.length != 4 || parseInt(L_Year) < 1) { L_Msg += 'valid 4 digit year required\n'};
    if (isNaN(L_Month) || parseInt(L_Month) < 1 || parseInt(L_Month) > 12) { L_Msg += 'invalid month value\n'};
    if (isNaN(L_Day) || parseInt(L_Day) < 1 || parseInt(L_Day) > 31) { L_Msg += 'invalid day value\n'};	
	if (L_Msg.length == 0) {
		if (parseInt(L_Day) > daysInMonth[parseInt(L_Month)]) {
			L_Msg += 'days invalid for month\n';
		}
		if ( (parseInt(L_Month) == 2 && parseInt(L_Day) == 29)
			&& (parseInt(L_Year) % 4 > 0
				|| (parseInt(L_Year) % 100 == 0 && parseInt(L_Year) % 400 > 0))
			) {
			L_Msg += 'days invalid for month\n';
		}
	}

	if (ShowAlert) {
		if (L_Msg.length > 0) {
			alert(L_Msg + 'Dates should be mm/dd/yyyy format');
			Field.focus();
			return false;
		} else {
		    return true;
		}
	} else {
	    return L_Msg;
	}
}
</script>
<%
sSptsGrpID = Session("SptsGrpID")
Select case sSptsGrpID
	Case "AWS"
		sHeader = " AWSA 3 Event"
		sLogo = "../../images/" & Session("RegnLogo")
		sBGColor = "Maroon"
		sRegn = Session("Region")
	Case "ABC"
		sHeader = " ABC Barefoot"
		sLogo = "../../images/logo_abc_" & lcase(Session("RegnID")) & ".gif"
		sBGColor="Blue"
	Case "NCW"
		sHeader = " NCWSA Collegiate"
		sLogo = "../../images/logo_ncw_" & lcase(Session("RegnID")) & ".jpg"
		sBGColor="Red"	
end select

%>
<body>
<%'response.write(sLogo & sSptsGrpID)
'response.end%>
<table width="100%" bgcolor="#FFFFE0"><tr><td>
<h2><img src="../../images/usawski.gif" width="150" height="39" alt="USA Water Ski logo" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;USA Water Ski <br>&nbsp;&nbsp;&nbsp;&nbsp;Post Tournament Functions</h2>
<%Select Case UCase(sSptsGrpID)
	Case "AWS"
		sBGColor = "Maroon"%>
		<table><tr><td><img src="../../images/logo_awsa_sm.jpg" WIDTH="104" HEIGHT="66"></td><td><img src="<%=sLogo%>"></td></tr>
			</table>
<%	Case "ABC"
		sBGColor="Blue"%>
		<table><tr><td><img src="../../images/logo_abc_sm.jpg" WIDTH="141" HEIGHT="79"></td><td><img src="<%=sLogo%>"></td></tr>
		    </table>
<%	Case "NCW"
		sBGColor="Red"%>
		<table><tr><td><img src="../../images/logo_ncw_sm_dotcom.gif" WIDTH="260" HEIGHT="56"></td><td><img src="<%=sLogo%>"></td></tr>
		    </table>
<%	Case "USW"
		sBGColor="Black"%>
		<table><tr><td><img src="../../images/logo_usw_150.gif" WIDTH="75" HEIGHT="30" ></td><td></td></tr>
		    </table>   
<%	Case "AKA"
		sBGColor="Green"%>
		<table><tr><td><img src="../../images/logo_aka_150.jpg" WIDTH="75" HEIGHT="75"></td><td></td></tr>
		    </table>
<% End Select%>
<form name="setwhere" method="post" OnSubmit="return FormValidation(setwhere);" action="awsh_rlist.asp">
<table width="100%"><tr><td BGcolor="<%=sBGColor%>" width="50%"><font color="white" size="3"><b>1. Select the Calendar Year</b></font></td><td></td></tr>
	<tr><td colspan="2"><center><p>
			<select name="YearPick" ID="YearPick" size="1" onchange="YearPicked(this)">
				<option value>
					<%	do %>
						<option value="<%= sStartYr%>" <% if cint(sStartYr) = cint(year(date)) then Response.Write(" SELECTED")%>><%= sStartYr%>
						<% sStartYr = sStartYr + 1%>
					<% loop until cstr(year(date)+3) - sStartYr = 0%>
				</select>
			<input type="hidden" Name="Year" ID="Year"> 
			<p></center></td></tr>
	<tr><td BGcolor="<%=sBGColor%>"><font color="white" size="3"><b>2. Limit the tournaments to be displayed.</b></font></td><td></td></tr>
	<tr><td colspan="2" valign="top"><center><table border="1">	
			<tr><td colspan="2"><b>Select Tournament Status</b></td></tr>
			<tr><td colspan="2"></td></tr>
			<tr><td>Sanctioned Only</td><td><input type="radio" name="TStatus" value="2"></td></tr>
			<tr><td>Canceled Only</td><td><input type="radio" name="TStatus" value="3"></td></tr>
			<tr><td>Scored Only</td><td><input type="radio" name="TStatus" value="4"></td></tr>
			<tr><td>Archived Only</td><td><input type="radio" name="TStatus" value="5"></td></tr>
			<tr><td>All</td><td><input type="radio" name="TStatus" value="6"></td></tr>
			<tr><td>Sanctioned, Scored, &amp; Canceled Only</td><td><input type="radio" name="TStatus" value="7" CHECKED></td></tr>
		</table></center></td></tr>
		
		<tr><td colspan="2" valign="top"><center><table border="1">
			<tr><td colspan="2"><b>Optional Conditions</b><br>Year must match droplist. &nbsp;&nbsp;&nbsp; <font size="2">Date Format: MM/DD/YYYY</font>  </td><td>Use</td></tr>
			
			<tr><td>Scheduled After </td><td><input name="Where1" size="10" maxlength="10"></td><td><input type="checkbox" name="ckWhere1" onclick="FixRange(true)"></td></tr>
			
			<tr><td>Scheduled Before </td><td><input name="Where2" size="10" maxlength="10"></td><td><input type="checkbox" name="ckWhere2" onclick="FixRange(true)"></td></tr>

		</table></center></td></tr>

	<tr><td></td>
	<tr><td BGcolor="<%=sBGColor%>"><font color="white" size="3"><b>3. Click Go </b></font></td></tr>
	<tr><td align="right"><input type="submit" value="GO" id="submit1" name="submit1"><p><p></td>
	<tr><td><input type="reset" value="Reset" id="reset1" name="reset1"></td></tr>
</table>		
		
		</td></tr></table>
		
</form>
</body>
</html>

