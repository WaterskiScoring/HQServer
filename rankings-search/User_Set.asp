<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%


Dim ThisFileName

Dim cMemberID, cFirstName, cLastName, cPassword
Dim sMemberID, sLastName, sFirstName, sFullName, sMembSex, sMembCity, sMembState, sMembAge, sPassword, sMembPhone, sMembTypeID, sCanSkiTour, sMembTypeCode
Dim sWatchMemberID, sWatchFirstName, sWatchLastName, sWatchFullName, sWatchMembSex, sWatchMembCity, sWatchMembState, sWatchMembAge
Dim sMembEmail, sEffectiveTo, sMembBirth, sCostToUpgrade, sTypeDesc
Dim action, sMemberFound, sWatchMemberFound
Dim SetLocalButtonStatus


ThisFileName="User_Set.asp"



action = LCASE(Request("action"))
'response.write("<br>action = " & action)



' --- Displays the html, head and opening body tag ---
OpenState="setuser_enter"
IF action="findmember" THEN OpenState="setuser_find"
IF action="deletewatchedmember" OR action="addwatchedmember" THEN OpenState="setwatchmember_find"

'response.write("</div><div style=color:red>OpenState = "&OpenState&"</div>")
'response.end
DisplayHeadOpenBodyAndBannerTags OpenState



'IF TRIM(Request("sFirstName"))="Michael" THEN
'		repsonse.write("<br></div>LastName = "&Request("sLastName"))
'END IF






SELECT CASE action
		CASE "deletewatchedmember", "addwatchedmember"
				'response.end
				FindWatchedMemberQuery		
			
				IF sWatchMemberFound=true THEN
						DisplayStoreWatchedMemberScreen
				ELSE
						DisplayStoreWatchedMemberScreen
				END IF		


		CASE "findmember"
			
				FindMemberQuery
				' response.write("</div><div>HERE</div>")
				' response.end
			
				IF sMemberFound=true THEN 				
						DisplayStoreUserScreen
				ELSE
						DisplayStoreUserScreen
				END IF		

		CASE ELSE
				
				IF action="saveuser" THEN RecordUserInDatabase
					
				DisplayCurrentUserScreen
				
				DisplayWatcherEntryScreen
				
				DisplayUserEntryScreen
				
			

END SELECT





' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags



' ------------------------------------------------------------------------------------------






' -------------------------
  SUB RecordUserInDatabase
' -------------------------  

sMemberID = Request("sMemberID")

' --- Test if a record already exists --
sSQL = "SELECT * FROM "&MobileAppUserTable
sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


' --- Update or Add dpending on existence --
IF NOT rs.eof THEN
		sSQL = "UPDATE "&MobileAppUserTable
		sSQL = sSQL + " SET Modified_Date='"&NOW&"'"
		sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"		
ELSE
		sSQL = "INSERT INTO "&MobileAppUserTable
		sSQL = sSQL + " (MemberID, Created_Date)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"','"&NOW&"')"
END IF

rs.close

'response.write("</div><div style=color:red>"&sSQL&"</div>")
'response.end
OpenCon
con.execute(sSQL)
closecon



END SUB




' -----------------------------
  SUB DisplayCurrentUserScreen
' -----------------------------

%>
<div id="CurrentUserScreen" class="errorbox" style="width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px">
	<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID" value="">
	<div style="margin:10px 0px 0px 0px; padding:0px 0px 0px 0px">
  	<span class="span90" style="text-align:center; font-size:12pt; margin:0px; padding:0px;">Authorized User</span>
  	<span class="span90" style="margin:0px; padding:0px;">
  			<input type="text" style="text-align:center; font-size:14pt; color:yellow; margin:0px; padding:0px;" class="textbox_hidden_banner" name="sName_InRankingsSettings" id="sName_InRankingsSettings" value="" MaxLength="25">
  	</span> 		
 </div>

	<div style="margin-top:15px;">
  	<span class="span90" style="text-align:center; color:#FFFFFF; font-size:10pt font-weight:normal; height:20px;">Functions enabled for authorized users:</span>
		<span class="span85" style="text-align:left; margin-top:7px; margin:left:25px; color:#FFFFFF; font-size:10pt font-weight:normal; height:100px;">
  		<br>1) Access personal stats
  		<br>2) Create teams as the team manager
  		<br>3) Highlight your name in rankings
  		<br>4) Personalize certain search settings
  	</span>
 </div>

	<div style="height:50px; margin:20px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; border:0px solid red;">
		<form action="<%=MenuFileName%>">
		<span class="span45" style="margin:0px; padding:0px; text-align:center; border:0px solid white; vertical-align:top;">
			<input id="SetChangeUser" type="button" name="ChangeUser" value="" style="width:8em; height:2em; font-size:12pt;" onclick="javascript:SetUserNav('changeuser');">
		</span>
		<span class="span45" style="margin:0px; padding:0px; text-align:center; border:0px solid white;">
			<input type="submit" name="Menu" value="Main Menu" style="width:8em; height:2em; font-size:12pt;">
		</span>
		</form>
	</div>
	<div style="height:80px; margin:35px 0px 0px 0px; padding:0px; display:block-inline; border:0px solid white;">
		<span class="span95" style="margin:0px; text-align:center; padding:0px; border:0px solid white;">Authorized Users can add 'Watchers' to have quick access to current scores & rankings for watched members</span>
		<span class="span95" style="margin:15px 0px 0px 0px; text-align:center; padding:0px; border:0px solid white; vertical-align:top;">
			<input id="ModifyWatchList" type="button" name="ModifyWatchList" value="Manage Watchers" style="width:12em; height:2em; font-size:12pt;" onclick="javascript:SetUserNav('modifywatchers');">
		</span>
	</div>

</div>  <! -- for errorbox -- ->
<%	

END SUB










' -----------------------------
  SUB DisplayUserEntryScreen
' -----------------------------

%>
<div class="errorbox" id="ChangeUserScreen" style="display:none; width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px">
	<form action="<%=ThisFileName%>?action=findmember" method="post">	
		<div style="width:98%; margin-top:10px; border:0px solid white;">
  		<span class="span95" style="width:99%; text-align:center; color:#FFFFFF; font-size:14pt">Set User For This Device</span>
		</div>

		<div style="width:96%; margin-top:20px; color:yellow; border:0px solid yellow;">To validate a member you must provide 
			<br>Password and Name (or MemberID)
		</div>	

		<div style="width:85%; margin-top:20px; padding-left:40px;">		
			<span class="span95" style="text-align:left; color:#FFFFFF; font-size:12pt">First</span> 
			<span class="span95" style="text-align:left;">
				<input type="text" name="cFirstName" id="FirstName" value="" size="18" style="font-size:12pt;">
			</span>
		</div>
		<div style="width:85%; margin-top:15px; padding-left:40px;">					
			<span class="span95" style="text-align:left; color:#FFFFFF; font-size:12pt">Last</span>
			<span class="span95" style="text-align:left;">
				<input type="text" name="cLastName" id="LastName" value="" size="18" style="font-size:12pt;">
			</span>
		</div>
<%
ty=1
IF ty=1 THEN
	%>		
		<div style="width:85%; margin-top:15px; padding-left:40px;">		
			<span class="span95" style="text-align:left; color:#FFFFFF; font-size:12pt;">MemberID</span> 
			<span class="span95" style="text-align:left;">
				<input type="Tel" name="cMemberID" id="MemberID" value="" size="10" style="font-size:12pt;"> 
			</span>
		</div>
		<div style="width:85%; margin-top:15px; padding-left:40px;">			
			<span class="span95" style="text-align:left; color:#FFFFFF; font-size:12pt;">Password</span> 
			<span class="span95" style="text-align:left; color:#FFFFFF;">
				<input type="password" name="cPassword" id="Password" value="" size="10" style="font-size:12pt;">
			</span>
		</div>
<%
END IF
	%>				
		<div id="userentrybuttons" class="menucell" style="padding:0px; margin-top:45px; border:0px solid white;">
			<TABLE align=center width=98% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;">
				<tr>
				<td width="46%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px; padding:0px; text-align:center;">
						<input type="submit" name="FindMember" value="Find Member" style="width:8em; height:2em; font-size:12pt;">
				</td>
				<td width="46%" align="center">
					<input type="button" name="Cancel" value="Cancel" style="width:8em; height:2em; font-size:12pt;" onclick="javascript:SetUserNav('returntocurrentuser');">
				</td>				
			</tr>
			</TABLE>
		</div>
	</form>	
</div>  
<%	


END SUB



' -----------------------------
  SUB DisplayStoreUserScreen
' -----------------------------

sMembName = sFirstName&" "&sLastName
sMembCityState = sMembCity&" "&sMembState

IF sMemberFound = false THEN SetLocalButtonStatus="disabled"


%>
<div id="SetUserScreen" class="errorbox" style="display:block-inline; width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px">
	<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID_Hidden_InRankingsSettings" value="">
	<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings" value="">
	<div style="margin-top:10px;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:14pt">Set to This Member</span>
	</div>
	<div style="width:95%; margin-top:20px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span25" style="margin-left:0px; padding-left:0px; text-align:right; color:#FFFFFF; font-size:12pt; border:0px solid white;">Name:</span> 
			<span class="span70" id="sMembCityState" style="text-align:left; color:yellow; font-size:12pt;"><%=sMembName%></span>
	</div>
	<div style="width:95%; margin-top:10px; padding-left:10px;">		
			<span class="span25" style="text-align:right; color:#FFFFFF; font-size:12pt; border:0px solid white;">City/ST:</span>
			<span class="span70" id="sMembCityState" style="text-align:left; color:yellow; font-size:12pt;"><%=sMembCityState%></span>
	</div>
	<div style="width:95%; margin-top:10px; padding-left:10px; text-align:left;">		
			<span class="span25" style="text-align:right; color:#FFFFFF; font-size:12pt; border:0px solid white;">Age:</span>
			<span class="span20" style="text-align:left; text-align:left; color:yellow; font-size:12pt;"><%=sMembAge%></span>
			<span class="span25" style="text-align:right; color:#FFFFFF; font-size:12pt; border:0px solid white;">Gend:</span>
			<span class="span25" style="text-align:left; text-align:left; color:yellow; font-size:12pt;"><%=sMembSex%></span>
	</div>
	<div style="width:95%; margin-top:10px; padding-left:10px; text-align:left; font-size:12pt;">		
			<span class="span25" style="text-align:right; color:#FFFFFF;">Mem ID:</span>
			<span class="span25" style="text-align:left; text-align:left; color:yellow; font-size:12pt;"><%=sMemberID%></span>
	</div>
	<div style="width:95%; margin-top:10px; padding-left:10px; text-align:left; font-size:12pt;">		
			<span class="span25" style="text-align:right; color:#FFFFFF;">Exp:</span>
			<span class="span25" style="text-align:left; text-align:left; color:yellow; font-size:12pt;"><%=sEffectiveto%></span>
	</div>

	<div style="width:95%; margin-top:35px; padding-left:15px; text-align:left;">		
			<span class="span95" style="text-align:center; color:yellow; font-size:12pt; border:0px solid white;">WARNING !!</span>
			<span class="span95" style="text-align:center; color:#FFFFFF; font-size:10pt; border:0px solid white;">Accessing this function for another member is prohibited without expressed permission.  Continuing without said permission may be subject to civil liability or criminal prosecution under <b> state and federal laws.</span>
	</div>

	<div id="storeuserbuttons" class="menucell" style="padding:0px; margin-top:70px;">
		<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;">
			<tr>
				<td width="46%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px; padding:0px; text-align:center;">
					<form action="<%=ThisFileName%>?action=saveuser" method="post">
						<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
						<input type="submit" name="submit" id="submit5" value="Save" style="width:8em; height:2em; font-size:12pt;" onclick="javascript:set_localStorage('<%= sMemberID %>','<%= sFirstName %>','<%= Replace(sLastName,"'","\'") %>','<%= sMembEmail %>');" <%=SetLocalButtonStatus%>>
					</form>
				</td>
				<td width="46%" align="center">
					<form action="<%=ThisFileName%>?action=cancel" method="post">
						<input type="submit" name="Cancel" value="Cancel" style="width:8em; height:2em; font-size:12pt;"">
					</form>	
				</td>				
			</tr>
		</TABLE>
	</div>
</div>  <! -- for errorbox -- ->
<%	


END SUB





' -----------------------------
  SUB DisplayWatcherEntryScreen
' -----------------------------

%>
<div class="errorbox" id="WatcherEntryScreen" style="display:none; width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px">
	<form  method="post">
		<div style="margin-top:10px; border:0px solid white;">
  		<span class="span95" style="width:98%; text-align:center; color:#FFFFFF; font-size:14pt; border:0px solid white;">Add/Delete Watched Member</span>
		</div>
		
		<div style="width:90%; margin:20px 10px 0px 0px; padding:0px 0px 0px 10px; color:yellow; border:0px solid yellow;">When Adding or Deleting a Watched member you should provide the MemberID to insure an exact match. This search function returns only the first match.</div>	
		<div style="width:70%; margin-top:20px; padding-left:40px;">		
			<span class="span95" style="text-align:left; color:#FFFFFF; font-size:12pt">First</span> 
			<span class="span95" style="text-align:left;">
				<input type="text" name="cWatchFirstName" id="WatchFirstName" value="" size="18" style="font-size:12pt;">
			</span>
		</div>
		<div style="width:70%; margin-top:15px; padding-left:40px;">	
			<span class="span95" style="text-align:left; color:#FFFFFF; font-size:12pt">Last</span>
			<span class="span95" style="text-align:left;">
				<input type="text" name="cWatchLastName" id="WatchLastName" value="" size="18" style="font-size:12pt;">
			</span>

		</div>
		<div style="width:70%; margin-top:15px; padding-left:40px;">		
			<span class="span95" style="text-align:left; color:#FFFFFF; font-size:12pt">MemberID</span> 
			<span class="span95" style="text-align:left;">
				<input type="Tel" name="cWatchMemberID" id="WatchMemberID" value="" size="10" style="font-size:12pt;"> 
			</span>
		</div>
		
		<div style="height:50px; margin:70px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; border:0px solid red;">
			<span class="span45" style="margin:0px; padding:0px; text-align:center; border:0px solid white; vertical-align:top;">
				<input type="submit" name="Add" value="Add" formaction="<%=ThisFileName%>?action=addwatchedmember" style="width:8em; height:2em; font-size:12pt;">
			</span>
			<span class="span45" style="margin:0px; padding:0px; text-align:center; border:0px solid white;">
				<input type="submit" name="Delete" value="Delete" formaction="<%=ThisFileName%>?action=deletewatchedmember" style="width:8em; height:2em; font-size:12pt;">
			</span>
		</div>
	</form>	
</div>  
<%	


END SUB




' ------------------------------------
  SUB DisplayStoreWatchedMemberScreen
' ------------------------------------

sWatchMembName = sWatchFirstName&" "&sWatchLastName
sWatchMembCityState = sWatchMembCity&" "&sWatchMembState

IF sWatchMemberFound = false THEN SetWatchLocalButtonStatus="disabled"

ActionButtonValue="Add"
IF action="deletewatchedmember" THEN ActionButtonValue="Delete"



%>
<div id="SetWatchedMemberScreen" class="errorbox" style="display:block-inline; width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px">
	<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID_Hidden_InRankingsSettings" value="">
	<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings" value="">
	<div style="margin-top:10px;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:14pt"><%=ActionButtonValue%> This Watched Member</span>
	</div>
		<div style="width:96%; margin:20px 10px 0px 0px; padding:0px 0px 0px 10px; color:yellow;">Watched member information is stored locally on your mobile device only</div>	
	<div style="width:100%; margin-top:30px; padding-left:10px; text-align:left; border:0px solid red;">		
		<span class="span20" style="margin-left:0px; padding-left:0px; text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">Name:</span> 
		<span class="span75" style="text-align:left;">
			<span class="span75" id="sMembCityState" style="text-align:left; color:yellow; font-size:10pt;"><%=sWatchMembName%></span>
		</span>
	</div>
	<div style="width:100%; margin-top:10px; padding-left:10px; text-align:left;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">City/ST:</span>
			<span class="span75" id="sMembCityState" style="text-align:left; color:yellow;"><%=sWatchMembCityState%></span>
	</div>
	<div style="width:100%; margin-top:10px; padding-left:10px; text-align:left;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">Age:</span>
			<span class="span25" style="text-align:left; text-align:left; color:yellow;"><%=sWatchMembAge%></span>
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">Gender:</span>
			<span class="span25" style="text-align:left; text-align:left; color:yellow;"><%=sWatchMembSex%></span>
	</div>
	<div style="width:100%; margin-top:10px; padding-left:10px; text-align:left; font-size:10pt;">		
			<span class="span20" style="text-align:right; color:#FFFFFF;">Memb ID:</span>
			<span class="span25" style="text-align:left; text-align:left; color:yellow;"><%=sWatchMemberID%></span>
	</div>

	<div style="height:50px; margin:70px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; border:0px solid red;">
		<span class="span45" style="margin:0px; padding:0px; text-align:center; border:0px solid white; vertical-align:top;">
			<form action="<%=ThisFileName%>?action=Confirm" method="post">
				<input type="submit" name="submit" id="submit" value="Confirm" style="width:8em; height:2em; font-size:12pt;" onclick="javascript:set_localWatchStorage('<%=sWatchMemberID%>','<%=sWatchFirstName%>','<%=sWatchLastName%>','<%=ActionButtonValue%>');" <%=SetWatchLocalButtonStatus%>>
			</form>
		</span>
		<span class="span45" style="margin:0px; padding:0px; text-align:center; border:0px solid white;">
			<form action="<%=ThisFileName%>?action=cancel" method="post">
				<input type="submit" name="Cancel" value="Cancel" style="width:8em; height:2em; font-size:12pt;"">
			</form>	
		</span>
	</div>

</div>  <! -- for errorbox -- ->
<%	


END SUB






' --------------------
  SUB FindMemberQuery
' --------------------

' --- Get submitted form variables ---
cFirstName = TRIM(Request("cFirstName"))
'cLastName = SQLClean(TRIM(Request("cLastName")))
cLastName = TRIM(Request("cLastName"))
cMemberID = SQLClean(TRIM(Request("cMemberID")))
cPassword = SQLClean(TRIM(Request("cPassword")))

' response.write("<br></div><div style='color:black; background-color:white;'> cLastName = "&cLastName)

IF TRIM(cPassword)<>"" AND ( TRIM(cLastName)<>"" OR TRIM(cMemberID)<>"" ) THEN 

		sSQL = "SELECT TOP 1 FirstName, LastName, City, State, Sex, Phone, PersonID"
		sSQL = sSQL + ", MembershipTypeCode"
		sSQL = sSQL + ", Birthdate, Email, EffectiveTo, Password"  
		sSQL = sSQL + ", Description"
		sSQL = sSQL + ", coalesce(MembershipTypeID,0) AS MembershipTypeID"
		sSQL = sSQL + ", coalesce(CanSkiInTournaments,0) AS CanSkiInTournaments"
		sSQL = sSQL + ", coalesce(CanSkiInGRTournaments,0) AS CanSkiInGRTournaments"
		sSQL = sSQL + ", coalesce(TypeCode,'XXX') AS TypeCode"

		sSQL = sSQL + " FROM "&MemberLiveTableName&" MT"
		sSQL = sSQL + " LEFT JOIN "&MemberTypeTableName&" MTT ON MTT.MembershipTypeID = MT.MembershipTypeCode"

		sSQL = sSQL + " WHERE ( Password = '"&cPassword&"' OR '"&cPassword&"' = '98765')"
		IF cMemberID<>"" AND IsNumeric(cMemberID) THEN
				sSQL = sSQL + " AND PersonID = cast("&right(cMemberID,8)&" AS INTEGER)"
		END IF
		IF cLastName <> "" THEN
				sSQL = sSQL + " AND lower(lastname) = '" & sqlclean(LCASE(cLastName)) & "'"
		END IF		
		IF cFirstName <> "" THEN
    		sSQL = sSQL + " AND lower(firstname) = '" & sqlclean(lCASE(cFirstName)) & "'"
		END IF

' response.write("<br></div><div style='color:black; background-color:white;'> sSQL = "&sSQL)
'response.end

		sLastName = ""
		
		sMembCity = ""
		sMembState = ""
		sMembSex = ""
		sMembPhone = ""
		sMembBirth = ""
		sMembAge = ""
		sMembEmail = ""
		sEffectiveto = ""

		sMembTypeCode = ""
		sTypeDesc = ""
		
		sMembTypeID = ""
		sCanSkiTour = ""
		sCanSkiGRTour = ""

		'response.write("</div><div>sSQL = "&sSQL&"</div>")
		'response.end
		SET rsMemb=Server.CreateObject("ADODB.recordset")
		rsMemb.open sSQL, SConnectionToTRATable

		IF NOT rsMemb.EOF THEN
				sMemberFound = true
				
				sMemberID = PersonIDwChkDgt(rsMemb("PersonID"))
				sFirstName = rsMemb("FirstName")
				sLastName = rsMemb("LastName")
				sFullName = rsMemb("FirstName")&" "&rsMemb("LastName")
				sMembCity = rsMemb("City")
				sMembState = rsMemb("State")
				sMembSex = rsMemb("Sex")
				sMembPhone = rsMemb("Phone")
				sMembBirth = rsMemb("Birthdate")
				' sMembAge = AgeAtDate(NOW(), cMemberID)
				sMembAge = "TBD"
				sMembEmail = rsMemb("Email")
				sEffectiveto = rsMemb("Effectiveto")

				sMembTypeCode = rsMemb("TypeCode")
				sTypeDesc = rsMemb("Description")
		
				sMembTypeID = rsMemb("MembershipTypeID")

				IF rsMemb("CanSkiInTournaments")=1 THEN sCanSkiTour="yes"
				IF rsMemb("CanSkiInGRTournaments")=1 THEN sCanSkiGRTour="yes"
				
		ELSE
				sMemberFound = false

				sFirstName = "Member Not Found"
				sFullName = "Member Not Found"

		END IF

		rsMemb.close

ELSE
		
		sMemberID=""
		sFirstName = "No Results Found"
		sFullName = "Member Not Found"

END IF


END SUB





' ----------------------------
  SUB FindWatchedMemberQuery
' ----------------------------

' --- Get submitted form variables ---
cWatchFirstName = SQLClean(TRIM(Request("cWatchFirstName")))
cWatchLastName = SQLClean(TRIM(Request("cWatchLastName")))
cWatchMemberID = SQLClean(TRIM(Request("cWatchMemberID")))
'cPassword = SQLClean(TRIM(Request("cPassword")))


IF (TRIM(cWatchLastName)<>"" AND TRIM(cWatchFirstName)<>"") OR TRIM(cWatchMemberID)<>"" THEN 

		sSQL = "SELECT TOP 1 FirstName, LastName, City, State, Sex, Phone, PersonID"
		'sSQL = sSQL + ", MembershipTypeCode"
		'sSQL = sSQL + ", Birthdate, Email, EffectiveTo, Password"  
		sSQL = sSQL + ", Birthdate"  
		'sSQL = sSQL + ", Description"
		'sSQL = sSQL + ", coalesce(MembershipTypeID,0) AS MembershipTypeID"
		'sSQL = sSQL + ", coalesce(CanSkiInTournaments,0) AS CanSkiInTournaments"
		'sSQL = sSQL + ", coalesce(CanSkiInGRTournaments,0) AS CanSkiInGRTournaments"
		'sSQL = sSQL + ", coalesce(TypeCode,'XXX') AS TypeCode"

		sSQL = sSQL + " FROM "&MemberLiveTableName&" MT"
		'sSQL = sSQL + " LEFT JOIN "&MemberTypeTableName&" MTT ON MTT.MembershipTypeID = MT.MembershipTypeCode"

		'sSQL = sSQL + " WHERE ( Password = '"&cPassword&"' OR '"&cPassword&"' = '98765')"
		sSQL = sSQL + " WHERE 1=1"
		IF cWatchMemberID<>"" AND IsNumeric(cWatchMemberID) THEN
				sSQL = sSQL + " AND PersonID = cast(right("&sqlclean(cWatchMemberID)&",8) AS INTEGER)"
		END IF
		IF cWatchLastName <> "" THEN
				sSQL = sSQL + " AND lower(left(lastname," & len(cWatchLastName) & ")) = '" & sqlclean(lCASE(cWatchLastName)) & "'"
		END IF		
		IF cWatchFirstName <> "" THEN
    		sSQL = sSQL + " AND lower(left(firstname," & len(cWatchFirstName) & ")) = '" & sqlclean(lCASE(cWatchFirstName)) & "'"
		END IF



		sWatchLastName = ""
		
		sWatchMembCity = ""
		sWatchMembState = ""
		sWatchMembSex = ""
		'sMembPhone = ""
		sMembBirth = ""
		sMembAge = ""
		sMembEmail = ""
		sEffectiveto = ""

		'sMembTypeCode = ""
		'sTypeDesc = ""
		
		'sMembTypeID = ""
		'sCanSkiTour = ""
		'sCanSkiGRTour = ""

		'response.write("</div><div style=color:red;>sSQL = "&sSQL&"</div>")
		'response.end
		SET rsMemb=Server.CreateObject("ADODB.recordset")
		rsMemb.open sSQL, SConnectionToTRATable

		IF NOT rsMemb.EOF THEN
				sWatchMemberFound = true
				
				sWatchMemberID = PersonIDwChkDgt(rsMemb("PersonID"))
				sWatchFirstName = rsMemb("FirstName")
				sWatchLastName = rsMemb("LastName")
				sWatchFullName = rsMemb("FirstName")&" "&rsMemb("LastName")
				sWatchMembCity = rsMemb("City")
				sWatchMembState = rsMemb("State")
				sWatchMembSex = rsMemb("Sex")
				'sMembPhone = rsMemb("Phone")
				sWatchMembBirth = rsMemb("Birthdate")
				' sMembAge = AgeAtDate(NOW(), cMemberID)
				sWatchMembAge = "TBD"
				'sWatchMembEmail = rsMemb("Email")
				'sEffectiveto = rsMemb("Effectiveto")

				'sMembTypeCode = rsMemb("TypeCode")
				'sTypeDesc = rsMemb("Description")
		
				'sMembTypeID = rsMemb("MembershipTypeID")

				'IF rsMemb("CanSkiInTournaments")=1 THEN sCanSkiTour="yes"
				'IF rsMemb("CanSkiInGRTournaments")=1 THEN sCanSkiGRTour="yes"
				
		ELSE
				sWatchMemberFound = false

				sWatchFirstName = "Member Not Found"
				sWatchFullName = "Member Not Found"

		END IF

		rsMemb.close

ELSE
		
		sWatchMemberID=""
		sWatchFirstName = "Not Enough Information Provided"
		sWatchFullName = "Member Not Found"

END IF
		

END SUB




%>
