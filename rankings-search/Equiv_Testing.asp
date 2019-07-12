<!--#include virtual="/rankings/settingsHQ.asp"-->
<%


	sSQL = " SELECT row.MemberID, row.SkiYearID, Email, FirstName, LastName"
	sSQL = sSQL & " 	, S_First_Instance_E, T_First_Instance_E, J_First_Instance_E, O_First_Instance_E"
	sSQL = sSQL & " 	, S_First_Instance_S, T_First_Instance_S"
	sSQL = sSQL & " 	, row.skiyear-datepart(yyyy,mt.birthdate)-1 AS MemberAge"
	sSQL = sSQL & " FROM"

	sSQL = sSQL & " ( SELECT DISTINCT MemberID, l.SkiYearID, skiyear" 
	sSQL = sSQL & " 			FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 			JOIN " & SkiYearTableName & " sy ON sy.skiyearid = l.skiyearid"
	sSQL = sSQL & " 					WHERE sy.defaultyear = 1"
	sSQL = sSQL & " 						AND Event IN ('S','T','J')"
	sSQL = sSQL & " 						AND ( Sent_Notice IS NULL OR Sent_Notice='N')"
	sSQL = sSQL & "							AND First_Instance <  CAST('7/1/' + RIGHT(SkiYear,2) AS Date)"
	sSQL = sSQL & " ) row"
 	
 	sSQL = sSQL & " JOIN " & MemberShortTableName & " mt ON mt.PersonID = RIGHT(row.MemberID,8)"
 	
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS S_First_Instance_E" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='S' AND Div IN ('EM','EW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) se"
	sSQL = sSQL & " ON se.SkiYearID=row.SkiYearID AND se.MemberID=row.MemberID"	
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS S_First_Instance_S" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='S' AND Div IN ('SM','SW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) ss"
	sSQL = sSQL & " ON ss.SkiYearID=row.SkiYearID AND ss.MemberID=row.MemberID"	
					
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS T_First_Instance_E" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='T' AND Div IN ('EM','EW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) te"
	sSQL = sSQL & " ON te.SkiYearID=row.SkiYearID AND te.MemberID=row.MemberID"
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS T_First_Instance_S" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='T' AND Div IN ('SM','SW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) ts"
	sSQL = sSQL & " ON ts.SkiYearID=row.SkiYearID AND ts.MemberID=row.MemberID"

	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS J_First_Instance_E" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='J' AND Div IN ('EM','EW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) je"
	sSQL = sSQL & " ON je.SkiYearID=row.SkiYearID AND je.MemberID=row.MemberID"
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS O_First_Instance_E" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='O' AND Div IN ('EM','EW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) oe"
	sSQL = sSQL & " ON oe.SkiYearID=row.SkiYearID AND oe.MemberID=row.MemberID"

	sSQL = sSQL & " WHERE Email IS NOT NULL AND LEN(Email)>2"
	sSQL = sSQL & "              AND row.skiyear-datepart(yyyy,mt.birthdate)-1 >= 18" 
	

response.write(sSQL)

%>	