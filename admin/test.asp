
<% 
Function CalculateDivision(SkiAge, Gender)
Dim AgeDivision
if len(SkiAge) = 0 then
	AgeDivision = "-"
elseif SkiAge >= 0 AND SkiAge < 10 THEN '1' 
	AgeDivision = "1"
elseif  SkiAge >= 10 AND SkiAge < 14 THEN '2' 
	AgeDivision = "2"
elseif  SkiAge >= 14 AND SkiAge < 18 THEN '3' 
	AgeDivision = "3"
elseif  SkiAge >= 18 AND SkiAge < 25 THEN '1' 
	AgeDivision = "1"
elseif  SkiAge >= 25 AND SkiAge < 35 THEN '2' 
	AgeDivision = "2"
elseif  SkiAge >= 35 AND SkiAge < 45 THEN '3' 
	AgeDivision = "3"
elseif  SkiAge >= 45 AND SkiAge < 53 THEN '4' 
	AgeDivision = "4"
elseif  SkiAge >= 53 AND SkiAge < 60 THEN '5' 
	AgeDivision = "5"
elseif  SkiAge >= 60 AND SkiAge < 65 THEN '6' 
	AgeDivision = "6"
elseif  SkiAge >= 65 AND SkiAge < 70 THEN '7' 
	AgeDivision = "7"
elseif  SkiAge >= 70 AND SkiAge < 75 THEN '8' 
	AgeDivision = "8"
elseif  SkiAge >= 75 AND SkiAge < 80 THEN '9' 
	AgeDivision = "9"
elseif  SkiAge >= 80 AND SkiAge < 85 THEN 'A' 
	AgeDivision = "A"
elseif  SkiAge >= 85 THEN 'B' 
	AgeDivision = "B"
else
	AgeDivision = "-"
end if
					  
if Gender = "M" AND SkiAge < 18 THEN 'B' 
	SkiGender = "B"
elseif Gender = "M" AND SkiAge >= 18 THEN 'M' 
	SkiGender = "M"
elseif Gender = "F" AND SkiAge < 18 THEN 'G' 
	SkiGender = "G"
elseif Gender = "F" AND SkiAge >= 18 THEN 'W' 
	SkiGender = "W"
else 
	SkiGender = "-"
end if					  

CalculateDivision = SkiGender & AgeDivision
				  
End Function


'REMOVE me when yo umove over
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("WaterSkiConn")
'REMOVE me when yo umove over

'clear out the temp table of any entries for this session
objConn1.execute "Delete FROM [Temp Registration Template Export Table] where sessionid = " & (Session.SessionID)

Set MemberstoExport = Server.CreateObject("ADODB.RecordSet")
MemberstoExport.ActiveConnection = objConn1
MemberstoExport.Open "SELECT * FROM [Export Members to Excel] Where " & Session("StateSQL") & " ;" 

Dim TempTable
Set TempTable = Server.CreateObject("ADODB.RecordSet")
TempTable.ActiveConnection = objConn1
TempTable.LockType = 3	'adLockOptimistic
TempTable.Open "[Temp Registration Template Export Table]" 

Do until MemberstoExport.EOF

	SkiAge = Session("TournamentYear") - DATEPART("yyyy", MemberstoExport("BirthDate")) - 1

	TempTable.addnew
	TempTable("sessionid") = (Session.SessionID)
	TempTable("newmemid") = MemberstoExport("PersonIDwithCheckDigit")
	TempTable("lname") = MemberstoExport("lname")
	TempTable("fname") = MemberstoExport("fname")
	TempTable("Div") = CalculateDivision(SkiAge, MemberstoExport("Gender"))
	TempTable("SkiAge") = SkiAge
	TempTable("city") = MemberstoExport("city")
	TempTable("State") = MemberstoExport("State")
	TempTable("PrimaryRecord") = True
	
	if MemberstoExport("EffectiveTo") >= cdate(session("tournamentdate")) and MemberstoExport("CanSkiInTournaments") = True then
		TempTable("Active") = True		
	else
		TempTable("Active") = False
		if MemberstoExport("EffectiveTo") <= cdate(session("tournamentdate")) then
			TempTable("UpgradeDescription") = "Exp " & datepart("m",MemberstoExport("EffectiveTo")) & "/" & datepart("yyyy",MemberstoExport("EffectiveTo"))
		else
			TempTable("UpgradeDescription") = "Needs Upgrd" 
			TempTable("CosttoUpgrade") = MemberstoExport("CosttoUpgrade")
		end if
	end if
	TempTable.Update	
	
	MemberstoExport.MoveNext
Loop

'Now add to the temp table everyone that has scores under the extra divisions
'Extra Divisions with Scores to Add to Registration Template Export Grouped
MemberstoExport.Close 
MemberstoExport.Open "SELECT * FROM [Extra Divisions with Scores to Add to Registration Template Export Grouped] Where " & Session("StateSQL") & " ;" 

Do until MemberstoExport.EOF

	SkiAge = Session("TournamentYear") - DATEPART("yyyy", MemberstoExport("BirthDate")) - 1

	TempTable.addnew
	TempTable("sessionid") = (Session.SessionID)
	TempTable("newmemid") = MemberstoExport("PersonIDwithCheckDigit")
	TempTable("lname") = MemberstoExport("lname")
	TempTable("fname") = MemberstoExport("fname")
	TempTable("Div") = MemberstoExport("Div")
	TempTable("SkiAge") = SkiAge
	TempTable("city") = MemberstoExport("city")
	TempTable("State") = MemberstoExport("State")
	TempTable("PrimaryRecord") = False
	TempTable.Update	
	
	MemberstoExport.MoveNext
Loop

response.write "Done!!"

%>






