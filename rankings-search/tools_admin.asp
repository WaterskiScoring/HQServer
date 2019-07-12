<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%


Dim action, sconfirm, ThisFileName
Dim cancel_button
Dim div_count



ThisFileName = "tools_admin.asp"
action = Request("action")
sconfirm = LCASE(Request("sconfirm"))
cancel_button = LCASE(Request("cancel_button"))

IF (action="confirmednewdivrecs" AND sconfirm<>"yes") OR cancel_button="cancel" THEN
		action="canceldivadd"		
END IF


' response.write("<br>action = "&action)


DefineTRAStyles 


' -- Avoids writing page on cancel redirect --
IF action<>"canceldivadd" THEN 
		WriteIndexPageHeader
END IF



' -- Counts to make sure Division records don't already exist for MAX(SkiYear) --
CheckExistanceOfThisSkiYearInDivisionTable
				
IF div_count>0 THEN action="displaydivisionsexistnotice"
' response.write("<br>Line 40 div_count = "&div_count)									
	 


SELECT CASE action
		CASE "newskiyear"
				CreateNewSkiYear
					
		
		CASE "displaydivisionsexistnotice"
				DisplayExistsNotice
		
		
		CASE "confirmednewdivrecs"
								
				BuildNewDivisionControlRecords
				
				ConfirmNewDCTRecordsAdded

				
		CASE "canceldivadd"
				response.redirect("/rankings/defaulthq.asp")
		
		CASE ELSE
				ConfirmToBuildNewDivisionRecords 

	

END SELECT



WriteIndexPageFooter


' ==================================================================================================    	  
' --                 END OF MAIN PROGRAM 
' ==================================================================================================









' ----------------------
  SUB CreateNewSkiYear 
' ----------------------


' --  Code from DefaultHQ.asp rows 385-395 --



' ----------------------------------------------------------------------------------    	  
' --- Determine the values from the largest record based on the value of SkiYear ---
' ----------------------------------------------------------------------------------

sSQL = "SELECT TOP 1 SkiYearID, SkiYear FROM " & SkiYearTableName
sSQL = sSQL + " ORDER BY SkiYear DESC"  

SET rs = Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable
      
Dim NewPrevYearID, NewSkiYear
IF NOT rs.eof THEN
		NewPrevYearID=rs("SkiYearID")
		NewSkiYear=rs("SkiYear")+1
END IF
			   
      
sSQL = "INSERT into " & SkiYearTableName
sSQL = sSQL + " ([SkiYearName], [BeginDate], [EndDate], [DefaultYear], [LastRecalc], [RecalcUnderway], [SkiYear], [PrevYearID])"
sSQL = sSQL + " VALUES ("
sSQL = sSQL + "'Ski Year: " & year(SQLClean(request("EndDate"))) & "'"
sSQL = sSQL + ", '" & SQLClean(request("begindate")) & "'"
sSQL = sSQL + ", '" & SQLClean(request("enddate")) & "'"
sSQL = sSQL + ", '0','0','0'"
sSQL = sSQL + ", '"&NewSkiYear&"'"
sSQL = sSQL + ", '"&NewPrevYearID&"'"
sSQL = sSQL + ")"

'response.write(sSQL)
'response.end
      
Con.Execute(sSQL)

rs.close
CloseCon
set rs = nothing

' WriteIndexPageHeader
' NewsTitle="Add Ski Year"
' News="Enter the new Ski Year information. <br><br> No Tournaments in this ski year can begin before the beginning date. <br><br> No events in this 
' ski year can end  after the ending date."


END SUB





' -------------------------------------
  SUB ConfirmToBuildNewDivisionRecords 
' -------------------------------------  

%>
	<form action="/rankings/<%= ThisFileName %>" method="post">
		<input type="hidden" name="action" value="confirmednewdivrecs"> 

		<div style="width:100%; text-align:center; padding-left:50px; margin-top:50px;">		 
		<TABLE class="innertable" style="text-align:center; border-style:none; border:1px solid green; width:80%;" >
			<TR>
				<th colspan=8 align=center>
					<font face=<% =font1 %> size="3" Color="<%=TextColor5%>"><b>Create NEW SkiYear Division Control Entries</b></font>
					<br>
				</th>
			</TR>  
	  	<TR>
				<TD colspan=8 style="border-style:none; height:80px;">
					<br>
					<TABLE class="innertable" style="text-align:center; border-style:none;" width="90%" >
						<tr>
					    <td style="text-align:right; border-style:none;">
					    	<font style="color:#000000;" size=2 face=arial"><b>To create Division Control entries for the next ski year<br>Enter YES and press Confirm </b></FONT>
					    </td>
  	          <td style="text-align:left; border-style:none;">
  	          	<input type="text" id="sconfirm" name="sconfirm" maxlength=3 size=5>
  	          </td>
	  	      </tr>  
						<tr>
  	          <td colspan=2 style="text-align:left; border-style:none; padding-left:20px;">
  	          	<font style="color:#000000; font-size=10pt;">
  	          		CAUTION: 
  	          		  <br> 1) Do not use this function if you do not have Full Administrator Rights
  	          		  <br> 2) The Ski Year for which the DCT records belong must exist BEFORE running this function 
  	          	</font>
  	          </td>
	  	      </tr>  

					</TABLE>		
				</TD>
			</TR>	
	  	<TR>
				<TD colspan=4 style="border-style:none; width:45%; height:35px; text-align:center;">
					<input type="submit" name="confirm_button" style="width:9em;" value="Confirm">
				</TD>
				<TD colspan=4 style="border-style:none; width:45%; height:35px; text-align:center;">
					<input type="submit" name="cancel_button" style="width:9em;" value="Cancel">
				</TD>
			</TR>		
		</TABLE>
		</div>	
	</form>	
<%



END SUB




' ------------------------------
  SUB ConfirmNewDCTRecordsAdded
' ------------------------------  
  
%>
	<form action="/rankings/defaulthq.asp" method="post">
		<div style="width:100%; text-align:center; padding-left:50px; margin-top:50px;">		 
		<TABLE class="innertable" style="text-align:center; border-style:none; border:1px solid green; width:70%;" >
			<TR>
				<th colspan=8 align=center>
					<font face=<% =font1 %> size="3" Color="<%=TextColor5%>"><b>NEW SkiYear DCT Entries Added</b></font>
					<br>
				</th>
			</TR>
			<tr>
  	  	<td style="text-align:center; padding-top:15px border-style:none; width:90%">
  	       <font style="color:#000000; font-size=10pt;">Skier Qualification Committee edits may be made to the new entries</FONT>
  	    </td>
	  	</tr>  
	  	<TR>
				<TD colspan=8 style="border-style:none; width:45%; height:50px; text-align:center;">
					<input type="submit" name="mainmenu" style="width:11em;" value="Return to Main Menu">
				</TD>
			</TR>		
		</TABLE>
		</div>	
	</form>	
<%





END SUB





' ------------------------
  SUB DisplayExistsNotice
' ------------------------  

%>
	<form action="/rankings/defaulthq.asp" method="post">
		<div style="width:100%; text-align:center; padding-left:50px; margin-top:50px;">		 
		<TABLE class="innertable" style="text-align:center; border:1px solid blue; width:70%;" >
			<TR>
				<th colspan=8 align=center>
					<font face=<% =font1 %> size="3" Color="<%=TextColor5%>"><b>Division Records Already Exist</b></font>
					<br>
				</th>
			</TR>
			<tr>
  	  	<td style="text-align:center; padding-top:15px; width:90%; border-bottom:0px solid">
  	       <font style="color:#000000; font-size=10pt;">A new Ski Year must exist before new Divisions can be created.</FONT>
  	    </td>
	  	</tr>  
	  	<TR>
				<TD colspan=8 style="border-style:none; width:45%; height:50px; text-align:center;">
					<input type="submit" name="mainmenu" style="width:11em;" value="Return to Main Menu">
				</TD>
			</TR>		
		</TABLE>
		</div>	
	</form>	
<%






END SUB






' ------------------------------------------------
  SUB CheckExistanceOfThisSkiYearInDivisionTable
' ------------------------------------------------  



sSQL = " SELECT COUNT(*) AS Div_Count"
sSQL = sSQL + " FROM usawsrank.division"
sSQL = sSQL + " WHERE SkiYearID = (SELECT MAX(SkiYearID) AS SkiYearID FROM USAWSRank.SkiYear);"

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 1

div_count = rs("Div_Count")

' response.write("<br>"&sSQL)
' response.write("<br>Line 301 div_count = "&div_count)
rs.close


END SUB
  



      	

' -------------------------------------
  SUB BuildNewDivisionControlRecords 
' -------------------------------------  


' Both sections of code use a temporary table named usawsrank.divtemp as a work area, to copy the appropriate existing source rows 
' from the division control table into, then reset the skiyearid value in that temp table, then insert those rows back into the master table.  

' IS THIS NEEDED?
' In the case of the default ski year change, we also need a step to delete the existing rows for skiyearid = 1, before inserting the replacement rows.   
      	 

' -- Open connection --
OpenCon      	 

'---------------------------------
'-- Drop temp table (if exists) --
'---------------------------------

' sSQL = "DROP TABLE USAWSRank.DivTemp;"
sSQL = "DELETE FROM USAWSRank.DivTemp;"

Con.Execute(sSQL)




'------------------------------------------------ 
'-- Copy the 12 mo divisions to the Temp table --
'------------------------------------------------

sSQL = " INSERT INTO USAWSRank.DivTemp"
sSQL = sSQL + " SELECT * FROM USAWSRank.Division"
sSQL = sSQL + "   WHERE SkiYearID = 1;"

Con.Execute(sSQL)

 

'-----------------------------------------------------------
'-- Update the Temp table to be the most recent SkiYearID --
'-----------------------------------------------------------

sSQL = " UPDATE USAWSRank.DivTemp"
sSQL = sSQL + " SET SkiYearID = (SELECT MAX(SkiYearID) AS SkiYearID FROM USAWSRank.SkiYear);"

Con.Execute(sSQL)

 


'---------------------------------------------------------------------
'-- Insert all the rows from the Temp table into the division table --
'---------------------------------------------------------------------

sSQL = " INSERT INTO USAWSRank.Division"
sSQL = sSQL + " SELECT * FROM USAWSRank.DivTemp;"

Con.Execute(sSQL)



 
 

' -------------------------------------------------------------------------------------------
' -- SECTION #2:  Creates new 12 mo SkiYear Division records from the MOST RECENT SKI YEAR --
' -------------------------------------------------------------------------------------------



' ----------------------------------
' -- Drop the temp division table --
' ----------------------------------

' sSQL = "DROP TABLE USAWSRank.DivTemp;"
sSQL = "DELETE FROM USAWSRank.DivTemp;"

Con.Execute(sSQL)



' ------------------------------------------------------------------------------------------------
' -- Insert rows into the temp table from div table where previous active ski year (default=1)
' ------------------------------------------------------------------------------------------------

sSQL = " INSERT INTO USAWSRank.DivTemp"
sSQL = sSQL + " SELECT * FROM USAWSRank.Division"
sSQL = sSQL + " WHERE SkiYearID = (SELECT SkiYearID FROM USAWSRank.SkiYear WHERE defaultyear = 1);"

Con.Execute(sSQL)

 

' ------------------------------------------
' -- Change the temp table SkiYearID to 1 --
' ------------------------------------------

sSQL = " UPDATE USAWSRank.DivTemp"
sSQL = sSQL + " SET SkiYearID = 1;"

Con.Execute(sSQL)


 
' --------------------------------------------------
' -- Delete the 12 mo records from Division table --
' --------------------------------------------------

sSQL = " DELETE FROM USAWSRank.Division"
sSQL = sSQL + " WHERE SkiYearID = 1;"

Con.Execute(sSQL)

     

' --------------------------------------------------------
' -- Insert the rows from temp into LIVE Division table --
' --------------------------------------------------------

sSQL = " INSERT INTO USAWSRank.Division"
sSQL = sSQL + " SELECT * FROM USAWSRank.DivTemp;"

Con.Execute(sSQL)




' -- Close connection --
CloseCon
 


END SUB
 





' -------------------------------------
  SUB Update12Mo_DivisionControlRecords 
' -------------------------------------  

' -------------------------------------------------------------------------------------------
' -- SECTION #2 (dupe):  Replaces 12 mo SkiYear Division records from the MOST RECENT SKI YEAR --
' -------------------------------------------------------------------------------------------



' ----------------------------------
' -- Drop the temp division table --
' ----------------------------------

' sSQL = "DROP TABLE USAWSRank.DivTemp;"
sSQL = "DELETE FROM USAWSRank.DivTemp;"

Con.Execute(sSQL)



' ------------------------------------------------------------------------------------------------
' -- Insert rows into the temp table from div table where previous active ski year (default=1)
' ------------------------------------------------------------------------------------------------

sSQL = " INSERT INTO USAWSRank.DivTemp"
sSQL = sSQL + " SELECT * FROM USAWSRank.Division"
sSQL = sSQL + " WHERE SkiYearID = (SELECT SkiYearID FROM USAWSRank.SkiYear WHERE defaultyear = 1);"

Con.Execute(sSQL)

 

' ------------------------------------------
' -- Change the temp table SkiYearID to 1 --
' ------------------------------------------

sSQL = " UPDATE USAWSRank.DivTemp"
sSQL = sSQL + " SET SkiYearID = 1;"

Con.Execute(sSQL)


 
' --------------------------------------------------
' -- Delete the 12 mo records from Division table --
' --------------------------------------------------

sSQL = " DELETE FROM USAWSRank.Division"
sSQL = sSQL + " WHERE SkiYearID = 1;"

Con.Execute(sSQL)

     

' --------------------------------------------------------
' -- Insert the rows from temp into LIVE Division table --
' --------------------------------------------------------

sSQL = " INSERT INTO USAWSRank.Division"
sSQL = sSQL + " SELECT * FROM USAWSRank.DivTemp;"

Con.Execute(sSQL)




' -- Close connection --
CloseCon
 


END SUB



 
' ---------------------------------
  SUB RunSQCQueryForAuditingDCT
' ---------------------------------

' As part of the process carried out by the SQC committee each year, there are NOPS factors present in the Division Control table, which are 
' edited and maintained by the SQC chairman each year.  So at the time that SQC chairman has done that editing, they need to do an “Audit” of that data, 
' against the content of the NOPS calculator spreadsheet (which is what they work from as the source of the data).   Here below is a select query which 
' I’ve run for them which spits out the values for that new ski year (highest skiyearid value), in a form which they can then just copy and paste into 
' that spreadsheet, to compare with the master values therein.   I would suggest you add a function to “List NOPS factors for highest ski year” to the 
' TRA system tools sub-menu, to run this select query and display the answerset as a table.

 
 

' --------------------------------------------------------------------------------------
' -- Query to list NOPS factors for Max Ski Year, to compare to NOPS Calc Spreadsheet --
' --------------------------------------------------------------------------------------
 
 
sSQL = "SELECT sy.skiyearname, rx.Div, rx.Event, rx.Base, rx.Exp"
sSQL = sSQL + " FROM (SELECT skiyearid, div, '1-S' as Event, Over_S as Base, OverExp_S as Exp"
sSQL = sSQL + " 		FROM usawsrank.division"
sSQL = sSQL + " 				WHERE skiyearid = (SELECT MAX(skiyearid) FROM usawsrank.skiyear)"
sSQL = sSQL + "  						AND left(div,1) in ('B','G','M','W','O')"

sSQL = sSQL + " UNION ALL"

sSQL = sSQL + " SELECT skiyearid, div, '2-T' as Event, Over_T as Base, OverExp_T as Exp"
sSQL = sSQL + " 	FROM usawsrank.division"
sSQL = sSQL + "  		WHERE skiyearid = (SELECT MAX(skiyearid) FROM usawsrank.skiyear)"
sSQL = sSQL + "				AND left(div,1) in ('B','G','M','W','O')"

sSQL = sSQL + " UNION ALL"

sSQL = sSQL + " SELECT skiyearid, div, '3-J' as Event, Over_J as Base, OverExp_J as Exp"
sSQL = sSQL + "   FROM usawsrank.division"
sSQL = sSQL + "     WHERE skiyearid = (SELECT MAX(skiyearid) FROM usawsrank.skiyear)"
sSQL = sSQL + "		    AND left(div,1) in ('B','G','M','W','O')) RX"

sSQL = sSQL + " JOIN  usawsrank.skiyear sy ON sy.skiyearid = rx.skiyearid"

sSQL = sSQL + " ORDER BY CASE"
sSQL = sSQL + " WHEN LEFT(rx.div,1) = 'O' THEN '3' + rx.div"
sSQL = sSQL + " WHEN RIGHT(rx.div,1) in ('M','W') THEN '4' + rx.div"
sSQL = sSQL + " WHEN LEFT(rx.div,1) in ('B','M') THEN '1' + rx.div"
sSQL = sSQL + " WHEN LEFT(rx.div,1) in ('G','W') THEN '2' + rx.div"
sSQL = sSQL + "	ELSE '1' + rx.div END, rx.event;"


' The Order By clause above results in ordering the divisions in the same sequence as they appear in that NOPS calculator spreadsheet – ie first 
' comes Boys and Men, then Girls and Women, and then finally the Elite Divisions last, with Open then Masters.  

END SUB





%>