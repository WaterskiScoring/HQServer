<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%


' -- --------------------------------------------------------------------------------------------------------------
' -- PURPOSE:
' -- The view-mystats/asp has a feature that creates a highlighted row showing that a score is a personal best
' -- This module serves as an endpoint activtated by an XMLHttpRequest from the mystats page to 
' --   1) Record the member request for that Member, Tour Event PB.
' --   2) Return the contents for an acknowledgement to the my stats page, based on whether there is an email and 
' --        whether or not a previous request had been made for that score. 

' -- --------------------------------------------------------------------------------------------------------------




' -- Dimensioning
' -----------------------------------------


' -- sWatchMemberIDs

Dim sWatchMemberIDs, ebody
Dim eMailBody, eMailSubj, eMailTo, eMailCC, eMailBCC, eMailFrom, eMailReplyTo

Dim sEventName, ThisTourDate
' , sTourName
Dim sEvent, sScore, sUnits, sThisPBExists, sEmailExists
Dim sFirstName, sLastName, sAddress1, sCity, sState, sZip, sEmail	

sWatchMemberIDs = ""



' -- Parameters passed by post --
InputTest = "Post payload"
' InputTest = "Test Mode"

IF InputTest = "Post payload" THEN 
		sTourID = Request("stid") 
		sMemberID = Request("smid")
		sEvent = Request("sevt")

ELSEIF InputTest = "Test Mode" THEN 
		' sTourID = "18W062R"
		sTourID = "18W134L"
		sMemberID = "300150474"
		sEvent ="T"
END IF


' ----------------------------------
' -- Send the email and add to dB --
' ----------------------------------

SendRequestStickerEmail





' =======================================================================
' -- END OF PROGRAM --
' =======================================================================












' -----------------------------
  SUB SendRequestStickerEmail
' -----------------------------


	' -- Query to get Member, Tournament and Score data for parameters --
	sSQL = "SELECT FirstName, LastName, Address1, City, State, Zip, Email"
	sSQL = sSQL + ", TName, TCity, TState, TDateS, TDateE, s.Event, MAX(s.Div) AS Div, MAX(s.Score) AS Score"
	sSQL = sSQL + ", CASE WHEN MAX(pb.MemberID) IS NOT NULL THEN 'Y' ELSE 'N' END AS ThisPBExists"
	sSQL = sSQL + " FROM usawsrank.Scores s"
	sSQL = sSQL + " LEFT JOIN usawaterski.dbo.Members m ON m.PersonID=RIGHT(s.MemberID,8)"
	sSQL = sSQL + " LEFT JOIN sanctions.dbo.TSchedul st ON st.TournAppID=LEFT(s.TourID,6)"
	sSQL = sSQL + " LEFT JOIN usawsrank.Personal_Best_Stickers pb ON pb.MemberID=s.MemberID AND LEFT(pb.TourID,6)=LEFT(s.TourID,6) AND pb.Event=s.Event"
	
	sSQL = sSQL + " WHERE s.MemberID='"&sMemberID&"' AND s.TourID='"&sTourID&"' AND s.Event='"&sEvent&"'"
	sSQL = sSQL + " GROUP BY s.MemberID, s.TourID, FirstName, LastName, Address1, City, State, Zip, Email"
	sSQL = sSQL + ", TName, TCity, TState, TDateS, TDateE, s.Event"	

	'response.write(sSQL)
	'response.end

	SET rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, SConnectionToTRATable


	' -- Define variables --
	IF NOT rs.eof THEN
			sThisPBExists = rs("ThisPBExists")
			
			sFirstName = rs("FirstName") 
			sLastName = rs("LastName")		
			sAddress1 = rs("Address1") 
			sCity = rs("City")
			sState = rs("State")
			sZip = rs("Zip")
			sEmail = rs("Email")
			sEmailExists="N"
			IF TRIM(sEmail)<>"" AND INSTR(sEmail,"@")>0 AND INSTR(sEmail,".")>0 THEN sEmailExists="Y"

			sEvent = rs("Event")			
			tDiv = rs("Div")
			sScore = rs("Score")
			
			sTourName = rs("Tname")
			sTourCity = rs("TCity")
			sTourState = rs("TState")
			sTDateS = rs("TDateS")
			sTDateE = rs("TDateE")

			' -- Configure date for display --
			ThisTourDate = MONTH(sTDateS) &"/"&DAY(sTDateS)&" - "&MONTH(sTDateE) &"/"& DAY(sTDateE) &"/"& YEAR(sTDateE)
			IF MONTH(sTDateS) <> MONTH(sTDateE) THEN 
					ThisTourDate = MONTH(sTDateS)& "/" & DAY(sTDateS)& "/" &YEAR(sTDateS)& " - "&MONTH(sTDateE) &"/"& DAY(sTDateE) &"/"& YEAR(sTDateE)
			END IF
			
			
			SELECT CASE sEvent
					CASE "S"
						sEventName = "Slalom" 
						sScore = FormatNumber(rs("Score"),2)
						sUnits = "buoys"
	 				CASE "T"
						sEventName = "Tricks" 
						sScore = REPLACE(FormatNumber(rs("Score"),0),",","")
						sUnits = "points"
					CASE "J"
						sEventName = "Jumping" 
						sScore = FormatNumber(rs("Score"),0)
						sUnits = "feet"						
			END SELECT



			' -- Add a record to the table if it doesn't already exist for this Member + TourID + Event --
			IF sThisPBExists = "N" THEN 
					
					ThisDate = FormatDateTime(NOW,2)
					
					sSQL = "INSERT INTO usawsrank.Personal_Best_Stickers"
					sSQL = sSQL + " (MemberID, TourID, Event, Div, Score, Created_Date)"
					sSQL = sSQL + " VALUES ('" &sMemberID&"', '"&sTourID&"', '"&sEvent&"', '"&tDiv&"', '"&sScore&"', '" &ThisDate& "')"
			
					'response.write(sSQL)
					' response.end
					OpenCon
					con.execute(sSQL)
					CloseCon		


					' -- Deploy the email message --
					IF sEmailExists="Y" THEN
							CreateAndSendEmailMessage
					END IF
					
								
			END IF		


			' -- Write response to console for js Alert --
			WriteResponseToConsole 


	ELSE
	
		' -- Some response when NO records found in Query--


		
	END IF



END SUB






' ----------------------------
  SUB WriteResponseToConsole 
' ----------------------------  


' -- Build and write XML for response to XMLHttpResponse --
response.ContentType="text/xml"



RespStr = "<USAWaterski><result>"
RespStr = RespStr + "<memberid>" &sMemberID& " </memberid>"
RespStr = RespStr + "<tourid>" &sTourID& "</tourid>"
RespStr = RespStr + "<eventname>" &sEventName& "</eventname>"
RespStr = RespStr + "<score>" &sScore& "</score>"
RespStr = RespStr + "<units>" &sUnits& "</units>"
RespStr = RespStr + "<scoreexists>" &sThisPBExists& "</scoreexists>"
RespStr = RespStr + "<emailexists>" &sEmailExists& "</emailexists>"
RespStr = RespStr + "<status>200</status>"
RespStr = RespStr + "</result></USAWaterski>"

response.write(RespStr)


END SUB






' ---------------------------------
  SUB CreateAndSendEmailMessage 
' ---------------------------------

		' tcolor = "#0000b3"
		
 		'USAWS_Logo ="http://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg"
 		USAWS_Logo ="http://www.usawaterski.org/rankings/images/logos/USAWSWSlogo175Wpx.png"

    eMailBody = "<HTML><HEAD><TITLE>Message Preview</TITLE></HEAD>"
    eMailBody = eMailBody + "<BODY style='font-family: Arial, Helvetica, sans-serif; text-align:left; font-size:10pt;'>"
		eMailBody = eMailBody + "<div style='width:auto; height:35px; margin-top:40px; padding-top:9px; color:#FFFFFF; background-color:#0000b3; border:1px solid #0000b3; border-radius:20px 20px 0px 0px; text-align:center; font-size:16pt; font-weight:bold;'>PB Sticker Request Received</div>"

    eMailBody = eMailBody + "<div style='width:auto; padding:0px 10px 0px 10px; border:1px solid black; border-radius:0px 0px 20px 20px;'>"

		eMailBody = eMailBody + "	<div style='width:100%; padding-top:30px;'>Dear " &sFirstName& ",</div>" 

		eMailBody = eMailBody + "	<div style='width:100%; padding-top:10px;'>Congratulations on skiing a Personal Best in <b>" &sEventName& "</b> with a score of <b>"&sScore&"</b> "&sUnits&" at:</div>"  

		eMailBody = eMailBody + "	<div style='width:100%; padding-top:10px; padding-left:10px;'><b>" &LEFT(sTourName,35)& "</b></div>"		
		eMailBody = eMailBody + "	<div style='width:100%; padding-left:10px;'>" &sTourCity& ", " &sTourState& "</div>"
		eMailBody = eMailBody + "	<div style='width:100%; padding-left:10px;'>TourID: " &sTourID& "</div>"
  	eMailBody = eMailBody + "	<div style='width:100%; padding-left:10px; margin-top:0px;'>" &ThisTourDate& "</div>"
 				
		eMailBody = eMailBody + "	<div style='width:100%; padding-top:10px;'>We received your request for a new Personal Best sticker.  It will ship to you shortly at the following address. </div>"
	
		eMailBody = eMailBody + "	<div style='width:100%; padding-left:10px; padding-top:10px;'>" &sFirstName& " " &sLastName& "</div>"
		eMailBody = eMailBody + "	<div style='width:100%; padding-left:10px; text-decoration: none;'>" &sAddress1& "</div>"
		eMailBody = eMailBody + "	<div style='width:100%; padding-left:10px; text-decoration: none;'>" &sCity& ", " &sState& " " &sZip& "</div>"

		eMailBody = eMailBody + "	<div style='width:100%; padding-top:10px;'>Please allow 2 weeks for delivery. NOTE: Only one event score is used for each tournament.</div>"

	
		eMailBody = eMailBody + "	<div style='width:100%; margin-top:10px;'>Great skiing!!</div>"		
               
    eMailBody = eMailBody + "	<div style='width:100%; margin-top:20px; margin-top:15px;'>Sincerely,</div>"
    eMailBody = eMailBody + "	<div style='width:100%; margin-top:20px;'>Jeff Surdej</div>"
    eMailBody = eMailBody + "	<div style='width:100%; margin:0px 0px 30px 0px;'>AWSA President</div>" 
		eMailBody = eMailBody + "</div>"

    eMailBody = eMailBody + "<div style='width:100%; text-align:center; font-size:8pt; margin:15px 0px 0px 0px;'>A Service of</div>" 
    eMailBody = eMailBody + "<div style='width:100%; text-align:center; margin:10px 0px 0px 0px;'><img src='"&USAWS_Logo&"' style='width:100px;'></div>" 
    eMailBody = eMailBody + "<div style='width:100%; text-align:center; font-size:8pt; font-style:bold; margin:10px 0px 0px 0px;'>180 Holy Cow Rd<br>Polk City, FL 33883</div>" 

    eMailBody = eMailBody + "</BODY></HTML>"
		


		' -- Define mailing values
		eMailTo = sEmail
		' eMailTo = "cronemarka@gmail.com"
		eMailCC = ""
		eMailBCC = "cronemarka@gmail.com, j_surdej@yahoo.com"
		eMailFrom = "competition@usawaterski.org"
		eMailReplyTo = "competition@usawaterski.org"
		eMailSubj = "Personal Best Sticker - "&sFirstName&" "&sLastName


		

	
	
		' -- Send using Generic function in tools_registration16.asp --
		Dim eMailTo, eMailCC, eMailBCC, eMailFrom, eMailReplyTo, eMailSubj, eMailBody
		SendEmailFromGenericMethodAndReplyTo eMailTo, eMailCC, eMailBCC, eMailFrom, eMailReplyTo, eMailSubj, eMailBody



END SUB



%>

  				