<!--#include virtual="/rankings/settingsHQ.asp"-->
<link rel="stylesheet" href="css/stylesheet_mob_tours.css" media="screen">
<head>
	<style>
.widgettab {
		font-size:11px;
		margin:5px 0px 0px 0px;
    padding:0px 2px 0px 5px; 
    border-top-left-radius:10px;
    border-top-right-radius:10px;    
		border-top: 1px solid blue;
    border-left: 1px solid blue;
    border-right: 1px solid blue;
    text-align:center;
    width:96%;
    height:12px;
		}

.widgetbody {
		font-size:10pt;
		margin:0px 0px 0px 0px;
    padding:0px 2px 0px 5px; 
    border-left: 1px solid #000000;
    border-right: 1px solid #000000;
    text-align:center;
    width:96%;
    height:15px;
    height:auto;
    background-color:#FFFFFF;
		}

.widgetbottom {
		font-size:10pt;
		margin:0px 0px 0px 0px;
    padding:0px 2px 0px 5px; 
    border-bottom-left-radius:10px;
    border-bottom-right-radius:10px;    
    /*border-bottom: 2px solid; */
    border-left: 1px solid #000000;
    border-right: 1px solid #000000;
    border-bottom: 1px solid #000000;
    text-align:center;
    width:96%;
    height:15px;
    height:auto;
		}

	</style>
</head>
<%





Dim ThisFileName, action, sSuccess
Dim sTourSelected
Dim sLongitude, sLatitude, sSubmitMessage

ThisFileName = "Register_longlat_widget.asp"




WriteIndexPageHeader

ReadFormVariables


SELECT CASE action
		CASE "update"

				IF sSuccess="Y" THEN SubmitCoordinates
						
		CASE ELSE

END SELECT

		
DisplayForm


WriteIndexPageFooter










' -----------------
  SUB DisplayForm
' -----------------

' style="background-color:white; color:#000000;"

%>
		<form action="/rankings/<%=ThisFileName%>?action=update" method="post">
			<div class="widgettab" style="color:white; background-color:<%=HQSiteColor2%>; height:25px; font-size:12pt; padding-top:2px; text-align:center;">Tournament GPS Location Coordinates Widget</div>
				<div class="widgetbody"  style="height:30px; padding-top:10px;">Select a Tournament from the site where you want to update GPS coordinates. </div>
				<div class="widgetbody" style="height:40px;">
					<%
					TourDropBuild
					%>
				</div>
				<div class="widgetbody" style="height:40px; font-size:10pt;">	
					Latitude
					<input type="text" name="sLatitude" value="<%=sLatitude%>">
					Longitude
					<input type="text" name="sLongitude" value="<%=sLongitude%>">

					<br> (Example format: 37.781001 or -97.444876)
				</div>
				<div class="widgetbody" style="color:red;"><%=sSubmitMessage%></div>
				<div class="widgetbody" style="height:50px; padding-top:30px;">
					<input type="submit" value="Submit Coordinates" style="width:15em;">
				</div>
				<div class="widgetbody" style="color:blue;">The widget is for administrative purposes only</div>				
			<div class="widgetbottom">&nbsp;</div> 	
			
			
		</form>
<%

END SUB





' ---------------------------------
  SUB ReadFormVariables
' ---------------------------------


	action=request("action")
	sSubmitMessage=""
  
	sTourSelected = TRIM(Request("sTourSelected"))
	sLatitude = TRIM(Request("sLatitude"))
	sLongitude = TRIM(Request("sLongitude"))
	
	sSuccess = "N"
	IF action="update" THEN
			IF sTourSelected = "" THEN 
					sSubmitMessage="You must select a tournament"	
			ELSEIF LEN(sLatitude)<8 OR LEN(sLongitude)<8 OR LEN(sLatitude)>15 OR LEN(sLongitude)>15 THEN 
					sSubmitMessage="Longitude and Latitude must be between 10 and 15 digits"
			ELSE 
					sSubmitMessage="New coordinates have been submitted"	
					sSuccess = "Y"
			END IF		
	END IF




END SUB  



' ------------------------
  SUB SubmitCoordinates
' ------------------------  


'response.write("sTourSelected = "&sTourSelected)
'response.write("<br>sLatitude = "&sLatitude)
'response.write("<br>sLongitude = "&sLongitude)


		sSQL = "UPDATE s"
		sSQL = sSQL + " SET Latitude='"&sLatitude&"',Longitude='"&sLongitude&"'"
		sSQL = sSQL + " FROM sanctions.dbo.wssites s"
		sSQL = sSQL + " JOIN sanctions.dbo.TSchedul ts ON ts.TSiteID=s.TSiteID"
		sSQL = sSQL + " WHERE TournAppID='"&LEFT(sTourSelected,6)&"'"

		' response.write(sSQL)
		OpenCon
		con.execute(sSQL)
		CloseCon


END SUB



' -----------------------
   SUB TourDropBuild
' -----------------------

' ------------   Builds Tournament Drop Down list ----------------- 

' --- First find the two-degit year for SkiYearSelected ---
' sSQL="SELECT RIGHT(SkiYear,2) AS ThisYear FROM "&SkiYearTableName&" AS ST WHERE ST.SkiYearID = '"&sSYID&"'"

sSQL = "SELECT TournAppID, TName FROM "&SanctionTableName&" s"
sSQL = sSQL + " LEFT JOIN "&SkiYearTableName&" sy ON RIGHT(sy.SkiYear,2)=LEFT(s.TournAppID,2)"
sSQL = sSQL + " WHERE DefaultYear=1 AND SptsGrpID='AWS'"
sSQL = sSQL + "  ORDER BY TournAppID"


'response.write(sSQL)
'response.end


SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 1

%>
<SELECT name='sTourSelected' style="width:20em"><%

  response.write("<option value =''")
  IF sTourSelected = "select" THEN response.write(" Selected")
  response.write(">Select Tournament</option><br>")

  IF NOT rs.eof THEN
	rs.movefirst
	DO WHILE not rs.eof
	  response.write("<option value =""" & rs("TournAppID") & """")
	  response.write(" <a title="""&rs("TName")&"""")
	  IF trim(rs("TournAppID")) = sTourSelected THEN
	    	response.write(" selected")
	  END IF

	  response.write(">")
	  response.write(rs("TournAppID") & " - " &rs("TName"))
	  response.write("</a></option><br>")
	  rs.movenext
	LOOP
  END IF %>

</SELECT><%

END SUB





%>



