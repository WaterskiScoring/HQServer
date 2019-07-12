<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%


Dim action, sconfirm, ThisFileName
Dim cancel_button
Dim div_count



ThisFileName = "tools_admin_copy_default_to_12mo.asp"
action = Request("action")
sconfirm = LCASE(Request("sconfirm"))
cancel_button = LCASE(Request("cancel_button"))

IF (action="confirmedupdatedivrecs" AND sconfirm<>"yes") OR cancel_button="cancel" THEN
		action="canceldivupdate"		
END IF


' response.write("<br>action = "&action)


DefineTRAStyles 


' -- Avoids writing page on cancel redirect --
IF action<>"canceldivupdate" THEN 
		WriteIndexPageHeader
END IF


					
	 


SELECT CASE action

				
		CASE "confirmedupdatedivrecs"
								
				Update12Mo_DivisionControlRecords
				
				Confirm12MoDCTRecordsUpdated

				
		CASE "canceldivupdate"
				response.redirect("/rankings/defaulthq.asp")
		
		CASE ELSE
				ConfirmToUpdate12MoDivisionRecords 

	

END SELECT



WriteIndexPageFooter


' ==================================================================================================    	  
' --                 END OF MAIN PROGRAM 
' ==================================================================================================













' -----------------------------------------
  SUB ConfirmToUpdate12MoDivisionRecords 
' -----------------------------------------  

%>
	<form action="/rankings/<%= ThisFileName %>" method="post">
		<input type="hidden" name="action" value="confirmedupdatedivrecs"> 

		<div style="width:100%; text-align:center; padding-left:50px; margin-top:50px;">		 
		<TABLE class="innertable" style="text-align:center; border-style:none; border:1px solid green; width:80%;" >
			<TR>
				<th colspan=8 align=center>
					<font face=<% =font1 %> size="3" Color="<%=TextColor5%>"><b>Update 12 mo SkiYear Division Control Entries</b></font>
					<br>
				</th>
			</TR>  
	  	<TR>
				<TD colspan=8 style="border-style:none; height:80px;">
					<br>
					<TABLE class="innertable" style="text-align:center; border-style:none;" width="90%" >
						<tr>
					    <td style="text-align:right; border-style:none;">
					    	<font style="color:#000000;" size=2 face=arial"><b>To update 12mo Ski Year Div Control from Default Ski Year <br>Enter YES and press Confirm </b></FONT>
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
  	          		  <br> 2) Edit Default Ski Year records BEFORE running this function 
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
  SUB Confirm12MoDCTRecordsUpdated
' ------------------------------  
  
%>
	<form action="/rankings/defaulthq.asp" method="post">
		<div style="width:100%; text-align:center; padding-left:50px; margin-top:50px;">		 
		<TABLE class="innertable" style="text-align:center; border-style:none; border:1px solid green; width:70%;" >
			<TR>
				<th colspan=8 align=center>
					<font face=<% =font1 %> size="3" Color="<%=TextColor5%>"><b>12 Mo SkiYear DCT Updated</b></font>
					<br>
				</th>
			</TR>
			<tr>
  	  	<td style="text-align:center; padding-top:15px border-style:none; width:90%">
  	       <font style="color:#000000; font-size=10pt;">SkiYearID=1 has been updated with Default Ski Year data</FONT>
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






 




' -------------------------------------
  SUB Update12Mo_DivisionControlRecords 
' -------------------------------------  

' -------------------------------------------------------------------------------------------
' -- SECTION #2 (dupe):  Replaces 12 mo SkiYear Division records from the MOST RECENT SKI YEAR --
' -------------------------------------------------------------------------------------------

' response.write("<br>Line 194: In Update Section")
' response.end



' -- Open connection --
OpenCon      	 



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



 



%>