<!--#include file="secure-settings.asp"-->



<% 
If Session("userlevel") < 1 Then

  WriteIndexPageHeader
  %>
  <br><br><br>
  <font color=red><center> You do not have access to this function. </center></font>
  <font color=red><center> Your security level is too low.  Please contact an administrator. </center></font>
  <%
  WriteIndexPageFooter
  Response.End

End If

Dim sDSN, i
Dim ConCSV, rsCSV
Dim sSQL
       
dim filespec
dim CSVfile
dim objfso
dim objstream
dim linecount
dim LineOne
dim linetext
dim resulttext
dim AgeInYears
dim rs1, rs2, rs3


OpenCon
    
Set rs=Server.CreateObject("ADODB.recordset")
Set rs2=Server.CreateObject("ADODB.recordset")
set rsSelectFields = Server.CreateObject("ADODB.recordset")
Set ConCSV = Server.CreateObject("ADODB.Connection")
Set rsCSV=Server.CreateObject("ADODB.recordset")
    
sDSN = "FileDSN=" & PathToTRA & "WSPDelim.DSN;DefaultDir=" & PathToExceptions & "\;DBQ=" & PathToExceptions & "\;Extensions=csv,wsp;"

'markdebug(sDSN)

ConCSV.Open sDSN
    

Set objfso = CreateObject("Scripting.FileSystemObject")
' We will delete the CSV when we are done processing.
filespec = PathtoExceptions & "\" & request("file")
CSVfile=PathToExceptions & "\" & right(left(request("file"),18),7) & ".csv"
objfso.CopyFile filespec, CSVfile, 1


if (request("fix") <> "1" and request("fix") <> "2") and request("delete") <> "1" then %>
        
    <html><head><title>Exception Editor</title></head><body>
    
    <%
    WriteIndexPageHeader
    NewsPageNum = "4"
    %>
    
    <center>
    <strong>
        <%
        response.write (right(left(request("file"),18),7))
    
        set objstream=objFSO.opentextfile(filespec)
        
        do while not objstream.atendofStream
          objstream.skipline
        loop
        linecount = objstream.Line
        objstream.close
        %>
    </strong><br>
    Current record:
    <% ' WSP files have a header row which should not be counted.
    Response.Write Request("line")-1
    %>
     of 
    <% ' Linecount is always +1 too high ... and WSP files have a header row
    Response.Write linecount-2
    %>
    <br><br>
    
    <%
    set objstream=objFSO.opentextfile(filespec)
    LineOne = objstream.ReadLine
    
    do while (not objstream.atendofStream) and (objstream.line - request("line") <> 0)
       objstream.skipline
    loop
    
    if objstream.atendofstream then
       response.redirect("/rankings/defaultHQ.asp?process=endoffile&line=" & request("line"))
    else
       lineText=objstream.readline
    
    objstream.close
    %>
    
    <%If Request("line") > 2 Then%>
    <a href="/rankings/exceptionmgmt-wsp.asp?file=<%=Request("file")%>&line=<%=Request("line")-1%>">
    <img src="/rankings/images/buttons/left.gif" border=0 title="Display Record Number <%=Request("line")-2%>"></a>
    <%Else Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    end if%>
    &nbsp;&nbsp;&nbsp;
    <%If linecount - Request("line") > 1 Then%>
    <a href="/rankings/exceptionmgmt-wsp.asp?file=<%=Request("file")%>&line=<%=Request("line")+1%>">
    <img src="/rankings/images/buttons/right.gif" border=0 title="Display Record Number <%=Request("line")%>"></a>
    <%Else Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    End If%>
    
    <br><hr>
    <b><u>Reasons Why the Record Failed</u></b><br><br>
    <textarea rows=3 cols=50 readonly style="overflow:hidden"><%
    set objstream=objFSO.opentextfile(PathtoReasons & "/" & request("file"))
    
    do while (not objstream.atendofStream) and (objstream.line - request("line") <> 0)
       objstream.skipline
    loop
    
    if NOT objstream.atendofstream then
       resulttext = objstream.readline   
       response.write(resulttext)
    end if
    
    objstream.close
    
    ' This SQL opens the divisions table to get allowable division codes.
    ' We use this later to build the division drop down options.
    '

    ' sSQL = "Select distinct div from " & RankTableName & " order by div"
    ' rsSelectFields.open sSQL, SConnectionToTRATable

    sSQL = "Select distinct div from " & DivisionsTableName & " order by div"
    rsSelectFields.open sSQL, SConnectionToTRATable

    %>
    </textarea><br>
    
    <%   If left(resulttext,1) = "*" Then %>
         <br><center>
         <form method=post action="/rankings/exceptionmgmt-wsp.asp">
         <input type="hidden" name="line" value="<%=Request("line")%>">
         <input type="hidden" name="file" value="<%=Request("file")%>">
         <input type="hidden" name="fix" value="1">
         <input type="submit" value="Fix Member Information"></form>
         </center>
    <%End If
    
    ' Open the WSP file and get read to read the fields
    
    sSQL = "Select * from " & CSVfile
    rsCSV.open sSQL, sDSN
    
    ' The line number is always +1 in WSP files because the header row doesn't count
    ' in the ADO recordset.
    ' We have to do -2 here because we compensate for the header and we also 
    ' don't have to move anywhere for the first record -- which is where we start
    ' by default.  So -1 for header and -1 so we don't move on the first record = -2.
    For i = 1 to (request("line")-2)
      rsCSV.MoveNext
    Next
    
    %>
    
    
    <form method=post action="/rankings/updatefields.asp">
    <input type="hidden" name="linenum" value="<%=Request("line")%>">
    <input type="hidden" name="file" value="<%=Request("file")%>">
    
    <TABLE class="innertable" BORDER="1" CELLPADDING="10" CELLSPACING="0" BGCOLOR="#FFFFFF" width=65%>
    
    <TR>
    <td colspan="2"><b><center>Participant Data</center></b></td>
    </tr>
    
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Member Federation</FONT></TD>
    <%InputData = rsCSV.fields(0)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Member_Federation" rows=1 cols=25 style="overflow:hidden"><%Response.Write inputdata%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Member ID</FONT></TD>
    <%InputData = rsCSV.fields(1)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Member_ID" rows=1 cols=25 style="overflow:hidden" readonly><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Last Name</FONT></TD>
    <%InputData = rsCSV.fields(2)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Lastname" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">First Name</FONT></TD>
    <%InputData = rsCSV.fields(3)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Firstname" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </TR>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Gender</FONT></TD>
    <%InputData = rsCSV.fields(4)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Gender" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Birth Year</FONT></TD>
    <%InputData = rsCSV.fields(5)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Birthyear" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">State</FONT></TD>
    <%InputData = rsCSV.fields(6)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="State" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Region</FONT></TD>
    <%InputData = rsCSV.fields(7)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Region" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <TR>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Team</FONT></TD>
    <%InputData = rsCSV.fields(8)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Team" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </TR>
    
    <TR>
    <td colspan="2"><b><center>Tournament Data</center></b></td>
    </tr>
    
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Tour Federation</FONT></TD>
    <%InputData = rsCSV.fields(0).name%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="TourFederation" rows=1 cols=25 style="overflow:hidden" readonly><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Tour ID</FONT></TD>
    <%InputData = rsCSV.fields(1).name%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="TourID" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Homologation Class</FONT></TD>
    <%InputData = rsCSV.fields(4).name%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="Homologation" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Tour End Date</FONT></TD>
    <%InputData = right(left(rsCSV.fields(9).name,6),2)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2">
    <textarea name="TourMonth" rows=1 cols=4 style="overflow:hidden"><%Response.Write InputData%></textarea>
    /
    <%InputData = right(rsCSV.fields(9).name,2)%>
    <textarea name="TourDay" rows=1 cols=4 style="overflow:hidden"><%Response.Write InputData%></textarea>
    /
    <%InputData = left(rsCSV.fields(9).name,4)%>
    <textarea name="TourYear" rows=1 cols=8 style="overflow:hidden"><%Response.Write InputData%></textarea>
    </FONT></TD>
    </TR>
    
    <TR>
    <td colspan="2"><b><center>Performance Summary Data</center></b></td>
    </tr>
    
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Number of Rounds Reported</FONT></Center></TD>
    <%InputData = trim(rsCSV.fields(9))%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="NumberofRounds" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Event Placement</FONT></TD>
    <%InputData = trim(rsCSV.fields(10))%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="SlalomPlacement" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Placement Points</FONT></TD>
    <%InputData = rsCSV.fields(11)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="SlalomPlacementPoints" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Round # of Best Slalom Score</FONT></TD>
    <%InputData = rsCSV.fields(12)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="BestSlalomRound">
    <%
    if trim(InputData) = "" then
      %><option value="" selected> </option><%
    else
      %><option value=""> </option><%
    end if  
    For i = 1 to rsCSV.fields(9)
      if InputData = i then
        %><option value="<%=i%>" selected><%=i%></option><%
      else
        %><option value="<%=i%>"><%=i%></option><%
      end if
    Next
    %>
    </select></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Event Placement</FONT></TD>
    <%InputData = trim(rsCSV.fields(13))%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="TrickPlacement" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Placement Points</FONT></TD>
    <%InputData = rsCSV.fields(14)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="TrickPlacementPoints" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Round # of Best Trick Score</FONT></TD>
    <%InputData = rsCSV.fields(15)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="BestTrickRound">
    <%
    if trim(InputData) = "" then
      %><option value="" selected> </option><%
    else
      %><option value=""> </option><%
    end if  
    For i = 1 to rsCSV.fields(9)
      if InputData = i then
        %><option value="<%=i%>" selected><%=i%></option><%
      else
        %><option value="<%=i%>"><%=i%></option><%
      end if
    Next
    %>
    </select></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Event Placement</FONT></Center></TD>
    <%InputData = trim(rsCSV.fields(16))%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="JumpPlacement" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Placement Points</FONT></TD>
    <%InputData = rsCSV.fields(17)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="JumpPlacementPoints" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Round # of Best Jump Score</FONT></TD>
    <%InputData = rsCSV.fields(18)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="BestJumpRound">
    <%
    if trim(InputData) = "" then
      %><option value="" selected> </option><%
    else
      %><option value=""> </option><%
    end if  
    For i = 1 to rsCSV.fields(9)
      if InputData = i then
        %><option value="<%=i%>" selected><%=i%></option><%
      else
        %><option value="<%=i%>"><%=i%></option><%
      end if
    Next
    %>
    </select></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Over All Placement</FONT></Center></TD>
    <%InputData = trim(rsCSV.fields(19))%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="OverAllPlacement" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Over All Placement Points</FONT></TD>
    <%InputData = rsCSV.fields(20)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="OverAllPlacementPoints" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
    </tr>
    <tr>
    <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Round # of Best Over All Score</FONT></TD>
    <%InputData = rsCSV.fields(21)%>
    <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="BestOverAllRound">
    <%
    if trim(InputData) = "" then
      %><option value="" selected> </option><%
    else
      %><option value=""> </option><%
    end if  
    For i = 1 to rsCSV.fields(9)
      if InputData = i then
        %><option value="<%=i%>" selected><%=i%></option><%
      else
        %><option value="<%=i%>"><%=i%></option><%
      end if
    Next
    %>
    </select></FONT></TD>
    </tr>
    
    <% For i = 1 To rsCSV.fields(9) %>
      <TR>
      <td colspan="2"><b><center>Round <%=i%> Score Data</center></b></td>
      </tr>
    
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Sanction Class</FONT></TD>
      <%InputData = rsCSV.fields(i*22+1)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="SL_<%=i%>_Sanction">
      <%If trim(InputData) = "" Then
        %><option value="" selected> </option><%
      else
        %><option value=""> </option><%
      end if  
      if ucase(trim(InputData)) = "C" then
        %><option value="C" selected>C</option><%
      else
        %><option value="C">C</option><%
      end if
      if ucase(trim(InputData)) = "E" then
        %><option value="E" selected>E</option><%
      else
        %><option value="E">E</option><%
      end if
      if ucase(trim(InputData)) = "L" then
        %><option value="L" selected>L</option><%
      else
        %><option value="L">L</option><%
      end if
      if ucase(trim(InputData)) = "R" then
        %><option value="R" selected>R</option><%
      else
        %><option value="R">R</option><%
      end if
      if ucase(trim(InputData)) = "N" then
        %><option value="N" selected>N</option><%
      else
        %><option value="N">N</option><%
      end if
      %>
      </select></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Division Code</FONT></TD>
      <%InputData = rsCSV.fields(i*22+2)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="SL_<%=i%>_Division">
      <%rsSelectFields.MoveFirst
        if not rsSelectFields.eof then 
          if trim(InputData) = "" Then
            response.write("<option value="" "" selected> </option>")
          else
            response.write("<option value="" ""> </option>")
          end if
          do while not rsSelectFields.eof
            if ucase(trim(rsSelectFields.Fields(0).value)) = ucase(InputData) then
              response.write("<option value =""" & rsSelectFields.Fields(0).value &""" selected>" & rsSelectFields.Fields(0).value & "</option><br>")
            else
              response.write("<option value =""" & rsSelectFields.Fields(0).value &""">" & rsSelectFields.Fields(0).value & "</option><br>")
            end if
            rsSelectFields.movenext
          loop
        else
          response.write("<option value="" "" selected> </option>")
        end if
      %>
      </select></font></td>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Boat Model Code</FONT></TD>
      <%InputData = rsCSV.fields(i*22+3)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="SL_<%=i%>_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Score</FONT></TD>
      <%InputData = rsCSV.fields(i*22+4)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="SL_<%=i%>_EndPassScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Speed</FONT></TD>
      <%InputData = rsCSV.fields(i*22+5)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="SL_<%=i%>_EndPassSpeed" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Line</FONT></TD>
      <%InputData = rsCSV.fields(i*22+6)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="SL_<%=i%>_EndPassLine" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Total Score</FONT></TD>
      <%InputData = rsCSV.fields(i*22+7)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="SL_<%=i%>_TotalScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </TR>
      <TR>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Sanction Class</FONT></TD>
      <%InputData = rsCSV.fields(i*22+8)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="TR_<%=i%>_Sanction">
      <%If trim(InputData) = "" Then
        %><option value="" selected> </option><%
      else
        %><option value=""> </option><%
      end if  
      if ucase(trim(InputData)) = "C" then
        %><option value="C" selected>C</option><%
      else
        %><option value="C">C</option><%
      end if
      if ucase(trim(InputData)) = "E" then
        %><option value="E" selected>E</option><%
      else
        %><option value="E">E</option><%
      end if
      if ucase(trim(InputData)) = "L" then
        %><option value="L" selected>L</option><%
      else
        %><option value="L">L</option><%
      end if
      if ucase(trim(InputData)) = "R" then
        %><option value="R" selected>R</option><%
      else
        %><option value="R">R</option><%
      end if
      if ucase(trim(InputData)) = "N" then
        %><option value="N" selected>N</option><%
      else
        %><option value="N">N</option><%
      end if
      %>
      </select></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Division Code</FONT></TD>
      <%InputData = rsCSV.fields(i*22+9)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="TR_<%=i%>_Division">
      <%rsSelectFields.MoveFirst
        if not rsSelectFields.eof then 
          if trim(InputData) = "" Then
            response.write("<option value="" "" selected> </option>")
          else
            response.write("<option value="" ""> </option>")
          end if
          do while not rsSelectFields.eof
            if ucase(trim(rsSelectFields.Fields(0).value)) = ucase(InputData) then
              response.write("<option value =""" & rsSelectFields.Fields(0).value &""" selected>" & rsSelectFields.Fields(0).value & "</option><br>")
            else
              response.write("<option value =""" & rsSelectFields.Fields(0).value &""">" & rsSelectFields.Fields(0).value & "</option><br>")
            end if
            rsSelectFields.movenext
          loop
        else
          response.write("<option value="" "" selected> </option>")
        end if
      %>
      </select></font></td>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Boat Model Code</FONT></TD>
      <%InputData = rsCSV.fields(i*22+10)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="TR_<%=i%>_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </TR>
      <TR>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Total Score</FONT></TD>
      <%InputData = rsCSV.fields(i*22+11)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="TR_<%=i%>_TotalScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </TR>
      <TR>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Sanction Class</FONT></TD>
      <%InputData = rsCSV.fields(i*22+12)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="JM_<%=i%>_Sanction">
      <%If trim(InputData) = "" Then
        %><option value="" selected> </option><%
      else
        %><option value=""> </option><%
      end if  
      if ucase(trim(InputData)) = "C" then
        %><option value="C" selected>C</option><%
      else
        %><option value="C">C</option><%
      end if
      if ucase(trim(InputData)) = "E" then
        %><option value="E" selected>E</option><%
      else
        %><option value="E">E</option><%
      end if
      if ucase(trim(InputData)) = "L" then
        %><option value="L" selected>L</option><%
      else
        %><option value="L">L</option><%
      end if
      if ucase(trim(InputData)) = "R" then
        %><option value="R" selected>R</option><%
      else
        %><option value="R">R</option><%
      end if
      if ucase(trim(InputData)) = "N" then
        %><option value="N" selected>N</option><%
      else
        %><option value="N">N</option><%
      end if
      %>
      </select></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Division Code</FONT></TD>
      <%InputData = rsCSV.fields(i*22+13)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="JM_<%=i%>_Division">
      <%rsSelectFields.MoveFirst
        if not rsSelectFields.eof then 
          if trim(InputData) = "" Then
            response.write("<option value="" "" selected> </option>")
          else
            response.write("<option value="" ""> </option>")
          end if
          do while not rsSelectFields.eof
            if ucase(trim(rsSelectFields.Fields(0).value)) = ucase(InputData) then
              response.write("<option value =""" & rsSelectFields.Fields(0).value &""" selected>" & rsSelectFields.Fields(0).value & "</option><br>")
            else
              response.write("<option value =""" & rsSelectFields.Fields(0).value &""">" & rsSelectFields.Fields(0).value & "</option><br>")
            end if
            rsSelectFields.movenext
          loop
        else
          response.write("<option value="" "" selected> </option>")
        end if
      %>
      </select></font></td>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Boat Model Code</FONT></TD>
      <%InputData = rsCSV.fields(i*22+14)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="JM_<%=i%>_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Ramp Height Ratio</FONT></TD>
      <%InputData = rsCSV.fields(i*22+15)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="JM_<%=i%>_RampHeight" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Boat Speed</FONT></TD>
      <%InputData = rsCSV.fields(i*22+16)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="JM_<%=i%>_BoatSpeed" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Best Distance (Feet)</FONT></TD>
      <%InputData = rsCSV.fields(i*22+17)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="JM_<%=i%>_DistanceFeet" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Best Distance (Meters)</FONT></TD>
      <%InputData = rsCSV.fields(i*22+18)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="JM_<%=i%>_DistanceMeter" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </TR>
      <TR>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Over All Sanction Class</FONT></TD>
      <%InputData = rsCSV.fields(i*22+19)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="OverAll_<%=i%>_Sanction">
      <%If trim(InputData) = "" Then
        %><option value="" selected> </option><%
      else
        %><option value=""> </option><%
      end if  
      if ucase(trim(InputData)) = "C" then
        %><option value="C" selected>C</option><%
      else
        %><option value="C">C</option><%
      end if
      if ucase(trim(InputData)) = "E" then
        %><option value="E" selected>E</option><%
      else
        %><option value="E">E</option><%
      end if
      if ucase(trim(InputData)) = "L" then
        %><option value="L" selected>L</option><%
      else
        %><option value="L">L</option><%
      end if
      if ucase(trim(InputData)) = "R" then
        %><option value="R" selected>R</option><%
      else
        %><option value="R">R</option><%
      end if
      if ucase(trim(InputData)) = "N" then
        %><option value="N" selected>N</option><%
      else
        %><option value="N">N</option><%
      end if
      %>
      </select></FONT></TD>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Over All Division Code</FONT></TD>
      <%InputData = rsCSV.fields(i*22+20)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><select name="OverAll_<%=i%>_Division">
      <%rsSelectFields.MoveFirst
        if not rsSelectFields.eof then 
          if trim(InputData) = "" Then
            response.write("<option value="" "" selected> </option>")
          else
            response.write("<option value="" ""> </option>")
          end if
          do while not rsSelectFields.eof
            if ucase(trim(rsSelectFields.Fields(0).value)) = ucase(InputData) then
              response.write("<option value =""" & rsSelectFields.Fields(0).value &""" selected>" & rsSelectFields.Fields(0).value & "</option><br>")
            else
              response.write("<option value =""" & rsSelectFields.Fields(0).value &""">" & rsSelectFields.Fields(0).value & "</option><br>")
            end if
            rsSelectFields.movenext
          loop
        else
          response.write("<option value="" "" selected> </option>")
        end if
      %>
      </select></font></td>
      </tr>
      <tr>
      <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Over All Score</FONT></TD>
      <%InputData = rsCSV.fields(i*22+21)%>
      <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><textarea name="OverAll_<%=i%>_Score" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
      </tr>
           
    <% Next 
      
    For i = 1 to rsCSV.fields(9) %>
          <input type="hidden" name="Round<%=i%>" value="1">
    <% Next %>
    
    
    
    
    
    </table>
    <br><center><input type="submit" value="Save New Field Data"></center>
    </form>
    <br>
    <form method=post action="/rankings/exceptionmgmt-wsp.asp">
    <input type="hidden" name="line" value="<%=Request("line")%>">
    <input type="hidden" name="file" value="<%=Request("file")%>">
    <input type="hidden" name="delete" value="1">
    <center><input type="submit" value="Delete This Record"></center>
    </form>
    
    <%
    
    rsSelectFields.close
    
    'This code allowed the user to edit the raw data line rather then breaking
    'the data out into the various fields.  If you understand the format
    'of the data line, sometimes this is easier.  If you aren't careful, it can
    'cause a ton of problems though.
    '
    'Mark decided to take it out of the final version.
    '
    '
    '  <form method=post action="/rankings/updateraw.asp">
    '  <input type="hidden" name="linenum" value="
    '  =Request("line")
    '  "><input type="hidden" name="file" value="
    '  =Request("file")
    '  "><br><br><hr>
    '  <b><u>Raw Data</u></b><br><br>
    '  <textarea name=rawdata rows=4 cols=75 style="overflow:hidden">
    '  Response.Write linetext
    '  </textarea><br>
    '  <br><center><input type="submit" value="Save New Raw Data"></center>
    '  </form>
    '
    '
    '

    end if
    %>
    <center>
    <%If Request("line") > 2 Then%>
    <a href="/rankings/exceptionmgmt-wsp.asp?file=<%=Request("file")%>&line=<%=Request("line")-1%>">
    <img src="/rankings/images/buttons/left.gif" border=0 title="Display Record Number <%=Request("line")-2%>"></a>
    <%Else Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    end if%>
    &nbsp;&nbsp;&nbsp;
    <%If linecount - Request("line") > 1 Then%>
    <a href="/rankings/exceptionmgmt-wsp.asp?file=<%=Request("file")%>&line=<%=Request("line")+1%>">
    <img src="/rankings/images/buttons/right.gif" border=0 title="Display Record Number <%=Request("line")%>"></a>
    <%Else Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    end if%>
    <br>&nbsp;<br><a href="/rankings/defaultHQ.asp">Return to the Main Menu</a>
    <br>&nbsp;<br>
    </center>
    
    <%WriteIndexPageFooter%>
    
    
    <%

    rsCSV.close

End If

if request("fix") = "1" then 
    
    %>
    <html><head><title>Exception Editor</title></head><body>
    <%

    WriteIndexPageHeader
    NewsPageNum="5"

    %>
    <center><h4>Member Data Correction<br><br></h4></center>
    <%
    
    
    ' Open the WSP file and get read to read the fields
        
    sSQL = "Select * from " & CSVfile
    rsCSV.open sSQL, sDSN
        
    ' The line number is always +1 in WSP files because the header row doesn't count
    ' in the ADO recordset.
    ' We have to do -2 here because we compensate for the header and we also 
    ' don't have to move anywhere for the first record -- which is where we start
    ' by default.  So -1 for header and -1 so we don't move on the first record = -2.
    For i = 1 to (request("line")-2)
      rsCSV.MoveNext
    Next
    
    %>
    <table class="innertable" align="center" width=50%>
    <tr>
      <th align=center colspan=2>
        <font color="#FFFFFF" size="<%=fontsize4%>"><b>Original Record Data</b></font>
      </th>
    </tr>
    <tr>
    	<td width=40%><font size="<%=fontsize3%>">First Name:</font></td>
	<td width=60%><% If trim(rsCSV.fields(3)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(3) & "</b></font>") %></td>
      </tr>
      <tr>
	<td><font size="<%=fontsize3%>">Last Name: </font></td>
	<td><% If trim(rsCSV.fields(2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(2) & "</b></font>") %></td>
      </tr>
      <tr>
	<td><font size="<%=fontsize3%>">Member ID:</font></td>
	<td><% If trim(rsCSV.fields(1)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(1) & "</b></font>") %></td>
      </tr>
      <tr>
	<td><font size="<%=fontsize3%>">Tour ID:</font></td>
	<td><%
    
    sSQL = "Select top 1 TSanction,TName,TCity,TState,TDateE from "& SanctionTableName &" where lower(TournAppID) = '" & sqlclean(lcase(right(left(Request("file"),17),6))) & "'"
    rs.open sSQL, sConnectionToTRATable, 3, 1
    IF NOT rs.EOF THEN %>
	<a title="<% =rs("tname") %>&#13;<% =rs("tcity")%>, <% =rs("tstate")%>&#13;<% =rs("tdatee")%>"><% 
    END IF
    Response.Write ("<font size="&fontsize3&"><b>"&right(left(Request("file"),18),7)&"</b></font>")
    Response.Write ("</a>")
    rs.Close %>
    
       </td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">Federation:</font></td>
    	<td><% If trim(rsCSV.fields(0)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(0) & "</b></font>") %></td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">Gender:</font></td> 
    	<td><% If trim(rsCSV.fields(4)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(4) & "</b></font>") %></td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">Birth Year:</font></td> 
	<td><% If trim(rsCSV.fields(5)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(5) & "</b></font>") %></td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">State:</font></td> 
	<td><% If trim(rsCSV.fields(6)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(6) & "</b></font>") %></td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">Region: </font></td>
	<td><% If trim(rsCSV.fields(7)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(7) & "</b></font>") %></td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">Team: </font></td>
	<td><% If trim(rsCSV.fields(8)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(8) & "</b></font>") %></td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">Slalom Division:</font></td> 
	<td><% If trim(rsCSV.fields(24)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(24) & "</b></font>") %></td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">Trick Division:</font></td> 
	<td><% If trim(rsCSV.fields(31)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(31) & "</b></font>") %></td>
      </tr>
      <tr>
    	<td><font size="<%=fontsize3%>">Jump Division:</font></td> 
	<td><% If trim(rsCSV.fields(35)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<font size="&fontsize3&"><b>" & rsCSV.fields(35) & "</b></font>") %></td>
      </tr>
     </table><%
    
    If Request.Form("search") = "1" Then 
    
      sSQL = "Select top 10 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex from "&MemberTableName&" where "
      If Request.Form("Member_ID") <> "" Then
        sSQL = sSQL + "PersonIDWithCheckDigit LIKE '%" & SQLClean(Request.Form("Member_ID")) & "%'"
        If Request.Form("Last_Name") <> "" or Request.Form("First_Name") <> "" or request.form("state") <> "" or request.form("date_of_birth") <> "" or request.form("gender") <> "" Then
          sSQL = sSQL + " and "
        End If
      End If
      If Request.Form("Last_Name") <> "" Then
        sSQL = sSQL + "lower(lastname) LIKE '%" & SQLClean(lcase(Request.Form("Last_Name"))) & "%'"
        If Request.Form("First_Name") <> "" or request.form("state") <> "" or request.form("date_of_birth") <> "" or request.form("gender") <> "" Then
          sSQL = sSQL + " and "
        End If
      End If
      If Request.Form("First_Name") <> "" Then
        sSQL = sSQL + "lower(firstname) LIKE '%" & SQLClean(lcase(Request.Form("First_Name"))) & "%'"
        if request.form("state") <> "" or request.form("date_of_birth") <> "" or request.form("gender") <> "" then
          sSQL = sSQL + " and "
        end if
      End If
      If Request.Form("State") <> "" Then
        sSQL = sSQL + "lower(state) LIKE '%" & SQLClean(lcase(Request.Form("state"))) & "%'"
        if request.form("date_of_birth") <> "" or request.form("gender") <> "" then
          sSQL = sSQL + " and "
        end if
      End If
      If Request.Form("Date_Of_Birth") <> "" Then
        sSQL = sSQL + "lower(birthdate) LIKE '%" & SQLClean(lcase(Request.Form("date_of_birth"))) & "%'"
        if request.form("gender") <> "" then
          sSQL = sSQL + " and "
        end if
      End If
      If Request.Form("gender") <> "" Then
        sSQL = sSQL + "left(lower(sex),1) = '" & SQLClean(lcase(left(Request.Form("gender"),1))) & "'"
      End If
      sSQL = sSQL + " and membertypeid <> 2 order by PersonIDWithCheckDigit"
        
        ' If statement protects against someone searching for nothing.
        
      If len(sSQL) < 166 Then
        sSQL = "Select * from members where 1=0"
      end if
    
      rs.open sSQL, sConnectionToMemberTable, 3, 1
      %>
        <center>    
        <br><b>Search Results</b>
        <% If rs.EOF Then %>
            <br><font color="red"> No Records Found, Please Search Again </font><br><br>
        <% Else %>
            <TABLE class="innertable"align="center" width=65%>
            <TR>
            <TH ALIGN="Center" vAlign="top"><FONT SIZE="2" color="#FFFFFF">Member ID</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT SIZE="2" color="#FFFFFF">First Name</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT SIZE="2" color="#FFFFFF">Last Name</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT SIZE="2" color="#FFFFFF">State</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT SIZE="2" color="#FFFFFF">Age</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT SIZE="2" color="#FFFFFF">Gender</FONT></TH>
            </tr>
            <% Do While Not rs.EOF
                  If IsNull(rs("BirthDate")) Then
                    AgeInYears = "N/a"
                  Else
                    ' get absolute number of years 
                    AgeInYears = cint(datediff("YYYY", rs("BirthDate"), Date())) 
    
                    ' get date1's month and day in terms of date2's year 
                    If dateadd("yyyy", ageInYears, rs("BirthDate")) > Date() Then 
                      ' their birthday hasn't hit yet in date2's year 
                      AgeInYears = AgeInYears - 1 
                    End If 
                  End If %>
    
                <tr>
                <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><a href="/rankings/exceptionmgmt-wsp.asp?fix=2&line=<%=Request("line")%>&file=<%=Request("file")%>&MemberID=<%=rs("PersonIDwithCheckDigit")%>"><%=rs("PersonIDwithCheckDigit")%></a></FONT></TD>
                <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("FirstName")%></FONT></TD>
                <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("LastName")%></FONT></TD>
                <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("State")%></FONT></TD>
                <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=AgeInYears%></FONT></TD>
                <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("Sex")%></FONT></TD>
                </tr>
                <% rs.MoveNext %>
            <% Loop %>
            </table>
            <% If rs.recordcount > 9 Then %>
                <font color="red"><small><center>Only the top ten records were displayed.</center></small></font>
            <% End If %>
      <% End If
        rs.Close
    
    ELSE
      sSQL = "Select FROMsubQ.PersonIDwithCheckDigit, min(FROMsubQ.lastname) as LastName, min(FROMsubQ.firstname) as FirstName, min(FROMsubQ.state) as State, min(FROMsubQ.birthdate) as BirthDate, min(FROMsubQ.sex) as Sex, min(FROMsubQ.Rank) as NewRank from ("
      'Nasty SubQuery Alert!!
    
      ' First we select everyone with a matching Member ID
      sSQL = sSQL + "Select top 3 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex,'1' as Rank from "&MemberTableName&" where "
      sSQL = sSQL + "PersonIDwithCheckDigit = " & SQLClean(trim(rsCSV.fields(1))) & " and "
      sSQL = sSQL + "membertypeid <> 2"
      sSQL = sSQL + " UNION " 
    
      ' Next we select everyone who has 3 char matches for both first and last name
      sSQL = sSQL + "Select top 3 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex,'2' as Rank from "&MemberTableName&" where "
      sSQL = sSQL + "difference(lower(lastname),'" & SQLClean(lcase(trim(rsCSV.fields(2)))) & "') = 4 and "
      sSQL = sSQL + "difference(lower(firstname),'" & SQLClean(lcase(trim(rsCSV.fields(3)))) & "') = 4 and "
      If trim(rsCSV.fields(4)) <> "" then
        sSQL = sSQL + "left(lower(sex),1) = '" & SQLClean(lcase(trim(rsCSV.fields(4)))) & "' and "
      End If
      sSQL = sSQL + "membertypeid <> 2"
      sSQL = sSQL + " UNION " 
      
      ' Then we check for left 4 digits of the Last Name + left 1 digit of first name
      sSQL = sSQL + "Select top 3 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex,'3' as Rank from "&MemberTableName&" where "
      sSQL = sSQL + "lower(lastname) LIKE '%" & SQLClean(lcase(trim(rsCSV.fields(2)))) & "%' and "
      sSQL = sSQL + "left(lower(firstname),1) = '" & SQLClean(lcase(trim(left(rsCSV.fields(3),1)))) & "' and "
      If trim(rsCSV.fields(4)) <> "" then
        sSQL = sSQL + "left(lower(sex),1) = '" & SQLClean(lcase(trim(rsCSV.fields(4)))) & "' and "
      End If
      sSQL = sSQL + "membertypeid <> 2"
      sSQL = sSQL + " UNION " 
    
      ' Then we check for left 4 digits of the First Name + left 1 digit of last name
      sSQL = sSQL + "Select top 3 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex,'4' as Rank from "&MemberTableName&" where "
      sSQL = sSQL + "left(lower(lastname),1) = '" & SQLClean(lcase(trim(left(rsCSV.fields(2),1)))) & "' and "
      sSQL = sSQL + "lower(firstname) LIKE '%" & SQLClean(lcase(trim(rsCSV.fields(3)))) & "%' and "
      If trim(rsCSV.fields(4)) <> "" then
        sSQL = sSQL + "left(lower(sex),1) = '" & SQLClean(lcase(trim(rsCSV.fields(4)))) & "' and "
      End If
      sSQL = sSQL + "membertypeid <> 2"
      sSQL = sSQL + ") as FROMsubQ group by PersonIDwithCheckDigit order by NewRank"
      
       
      rs.open sSQL, sConnectionToTRATable, 3, 1
      If rs.EOF Then %>
          <br><font color=red>Unable to find potential match.</font>
    <% Else %>
        <br>
        <TABLE class="innertable"  align="center" BORDER="1" width=65%>
          <TR>
            <TH ALIGN="Center" colspan=6 vAlign="top"><FONT COlOR="#FFFFFF" SIZE="3">Possible Matching Members</FONT></TH>
	  </TR>
          <TR>
            <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">Member ID</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">First Name</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">Last Name</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">State</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">Age</FONT></TH>
            <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">Gender</FONT></TH>
          </TR>
       <% Do While Not rs.EOF
            If IsNull(rs("BirthDate")) Then
              AgeInYears = "n/a"
            Else
              ' get absolute number of years 
              AgeInYears = cint(datediff("YYYY", rs("BirthDate"), Date())) 
    
              ' get date1's month and day in terms of date2's year 
              If dateadd("yyyy", ageInYears, rs("BirthDate")) > Date() Then 
                ' their birthday hasn't hit yet in date2's year 
                AgeInYears = AgeInYears - 1 
              End If 
            End If %>
          <tr>
            <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><a href="/rankings/exceptionmgmt-wsp.asp?fix=2&line=<%=Request("line")%>&file=<%=Request("file")%>&MemberID=<%=rs("PersonIDwithCheckDigit")%>"><%=rs("PersonIDwithCheckDigit")%></a></FONT></TD>
            <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("FirstName")%></FONT></TD>
            <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("LastName")%></FONT></TD>
            <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("State")%></FONT></TD>
            <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=AgeInYears%></FONT></TD>
            <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("Sex")%></FONT></TD>
          </tr>
       <% rs.MoveNext %>
       <% Loop %>
       </table><% 
      END IF 
      rs.Close
    
    END IF
    
    %>
    </td>
    </tr>
    </table>
    <center>
    <form method=post action="/rankings/exceptionmgmt-wsp.asp">
    <input type="hidden" name="line" value="<%=Request("line")%>">
    <input type="hidden" name="file" value="<%=Request("file")%>">
    <input type="hidden" name="fix" value="0">
    <input type="submit" value="Return Without Making Changes">
    </form>
    </center>
    
    <b><center>Search Membership Database</center></b>
    <form method=post action="/rankings/exceptionmgmt-wsp.asp">
    <input type="hidden" name="line" value="<%=Request("line")%>">
    <input type="hidden" name="file" value="<%=Request("file")%>">
    <input type="hidden" name="fix" value="1">
    <input type="hidden" name="search" value="1">

    <TABLE align="center" class="innertable" BORDER="1" width=65%>
    <TR>
    <TH ALIGN="Left" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">Member ID</FONT></TH>
    <TH ALIGN="Left" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">First Name</FONT></TH>
    <TH ALIGN="Left" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">Last Name</FONT></TH>
    <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">State</FONT></TH>
    <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">Date of Birth</FONT></TH>
    <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="2">Gender</FONT></TH>
    </TR>
    
    <TR>
    <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" SIZE="2"><input type="text" name="Member_ID" size=15></input></FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" SIZE="2"><input type="text" name="First_Name" size=15></input></FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" SIZE="2"><input type="text" name="Last_Name" size=15></input></FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" SIZE="2"><input type="text" name="State" size=5></input></FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" SIZE="2"><input type="text" name="Date_of_Birth" size=10></input></FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" SIZE="2"><input type="text" name="Gender" size=5></input></FONT></Center></TD>
    </TR>
    </table>
    <br>
    <center> 	
    <input type="submit" value="Search Member Database"></form>
    </center>
    <%
    
    WriteIndexPageFooter

    rsCSV.close

End If 


if request("fix") = "2" then
    
    %>
    
    <html><head><title>Exception Editor</title></head><body>
    
    <%WriteIndexPageHeader%>
    
    <center>
    <br><h4>Member Data Correction</h4>
    
    <%
    
    ' Open the WSP file and get read to read the fields
        
    sSQL = "Select * from " & CSVfile
    rsCSV.open sSQL, sDSN
        
    ' The line number is always +1 in WSP files because the header row doesn't count
    ' in the ADO recordset.
    ' We have to do -2 here because we compensate for the header and we also 
    ' don't have to move anywhere for the first record -- which is where we start
    ' by default.  So -1 for header and -1 so we don't move on the first record = -2.
    For i = 1 to (request("line")-2)
      rsCSV.MoveNext
    Next
    
    
    %>
    
    
    <form method=post action="/rankings/updatefields.asp">
        <input type="hidden" name="linenum" value="<%=Request("line")%>">
        <input type="hidden" name="file" value="<%=Request("file")%>">
        <input type="hidden" name="Region" value="<%=rsCSV.fields(7)%>">
        <input type="hidden" name="Team" value="<%=rsCSV.fields(8)%>">
        <input type="hidden" name="TourFederation" value="<%=rsCSV.fields(0).name%>">
        <input type="hidden" name="TourID" value="<%=rsCSV.fields(1).name%>">
        <input type="hidden" name="Homologation" value="<%=rsCSV.fields(4).name%>">
        <input type="hidden" name="TourMonth" value="<%=right(left(rsCSV.fields(9).name,6),2)%>">
        <input type="hidden" name="TourDay" value="<%=right(rsCSV.fields(9).name,2)%>">
        <input type="hidden" name="TourYear" value="<%=left(rsCSV.fields(9).name,4)%>">
        <input type="hidden" name="NumberofRounds" value="<%=trim(rsCSV.fields(9))%>">
        <input type="hidden" name="SlalomPlacement" value="<%=trim(rsCSV.fields(10))%>">
        <input type="hidden" name="SlalomPlacementPoints" value="<%=trim(rsCSV.fields(11))%>">
        <input type="hidden" name="BestSlalomRound" value="<%=rsCSV.fields(12)%>">
        <input type="hidden" name="TrickPlacement" value="<%=trim(rsCSV.fields(13))%>">
        <input type="hidden" name="TrickPlacementPoints" value="<%=rsCSV.fields(14)%>">
        <input type="hidden" name="BestTrickRound" value="<%=rsCSV.fields(15)%>">
        <input type="hidden" name="JumpPlacement" value="<%=trim(rsCSV.fields(16))%>">
        <input type="hidden" name="JumpPlacementPoints" value="<%=rsCSV.fields(17)%>">
        <input type="hidden" name="BestJumpRound" value="<%=rsCSV.fields(18)%>">
        <input type="hidden" name="OverAllPlacement" value="<%=trim(rsCSV.fields(19))%>">
        <input type="hidden" name="OverAllPlacementPoints" value="<%=rsCSV.fields(20)%>">
        <input type="hidden" name="BestOverAllRound" value="<%=rsCSV.fields(21)%>">
    
    <% For i = 1 To rsCSV.fields(9) %>
        <input type="hidden" name="SL_<%=i%>_Sanction" value="<%=rsCSV.fields(i*22+1)%>">
        <input type="hidden" name="SL_<%=i%>_Division" value="<%=rsCSV.fields(i*22+2)%>">
        <input type="hidden" name="SL_<%=i%>_Boat" value="<%=rsCSV.fields(i*22+3)%>">
        <input type="hidden" name="SL_<%=i%>_EndPassScore" value="<%=rsCSV.fields(i*22+4)%>">
        <input type="hidden" name="SL_<%=i%>_EndPassSpeed" value="<%=rsCSV.fields(i*22+5)%>">
        <input type="hidden" name="SL_<%=i%>_EndPassLine" value="<%=rsCSV.fields(i*22+6)%>">
        <input type="hidden" name="SL_<%=i%>_TotalScore" value="<%=rsCSV.fields(i*22+7)%>">
        <input type="hidden" name="TR_<%=i%>_Sanction" value="<%=rsCSV.fields(i*22+8)%>">
        <input type="hidden" name="TR_<%=i%>_Division" value="<%=rsCSV.fields(i*22+9)%>">
        <input type="hidden" name="TR_<%=i%>_Boat" value="<%=rsCSV.fields(i*22+10)%>">
        <input type="hidden" name="TR_<%=i%>_TotalScore" value="<%=rsCSV.fields(i*22+11)%>">
        <input type="hidden" name="JM_<%=i%>_Sanction" value="<%=rsCSV.fields(i*22+12)%>">
        <input type="hidden" name="JM_<%=i%>_Division" value="<%=rsCSV.fields(i*22+13)%>">
        <input type="hidden" name="JM_<%=i%>_Boat" value="<%=rsCSV.fields(i*22+14)%>">
        <input type="hidden" name="JM_<%=i%>_RampHeight" value="<%=rsCSV.fields(i*22+15)%>">
        <input type="hidden" name="JM_<%=i%>_BoatSpeed" value="<%=rsCSV.fields(i*22+16)%>">
        <input type="hidden" name="JM_<%=i%>_DistanceFeet" value="<%=rsCSV.fields(i*22+17)%>">
        <input type="hidden" name="JM_<%=i%>_DistanceMeter" value="<%=rsCSV.fields(i*22+18)%>">
        <input type="hidden" name="OverAll_<%=i%>_Sanction" value="<%=rsCSV.fields(i*22+19)%>">
        <input type="hidden" name="OverAll_<%=i%>_Division" value="<%=rsCSV.fields(i*22+20)%>">
        <input type="hidden" name="OverAll_<%=i%>_Score" value="<%=rsCSV.fields(i*22+21)%>">
    <%Next
    
    For i = 1 To rsCSV.fields(9) %>
       <input type="hidden" name="Round<%=i%>" value="1">
    <%Next
    
    
        sSQL = "Select top 1 FederationCode,PersonIDwithCheckDigit,LastName,FirstName,BirthDate,Sex,"&MemberTableName&".State,"&RegionTableName&".Region from "&MemberTableName
        sSQL = sSQL + " left join " & RegionTableName & " on lower(" & MemberTableName & ".state) = lower(" & RegionTableName & ".state)"
        sSQL = sSQL + " where PersonIDwithCheckDigit = '" & SQLClean(Request("MemberID")) & "'"
    
        rs.open sSQL, sConnectionToTRATable, 3, 1
    
        If IsNull(rs("BirthDate")) Then
          AgeInYears = "n/a"
        Else
          ' get absolute number of years 
          AgeInYears = cint(datediff("YYYY", rs("BirthDate"), Date())) 
    
          ' get date1's month and day in terms of date2's year 
          If dateadd("yyyy", ageInYears, rs("BirthDate")) > Date() Then 
            ' their birthday hasn't hit yet in date2's year 
            AgeInYears = AgeInYears - 1 
          End If 
        End If %>
    
    
    
    <table class="outertable" align="center" width=65% border=1 >
    <tr>
      <td colspan=2 align="center">Are you sure you wish to save these changes?
	<br><br>
      </td>
    </tr>
    <tr>
      <td align=center>	
        <input type="hidden" name="Member_Federation" value="<%=rs("FederationCode")%>">
        <input type="hidden" name="Member_ID" value="<%=rs("PersonIDwithCheckDigit")%>">
        <input type="hidden" name="Lastname" value="<%=rs("LastName")%>">
        <input type="hidden" name="Firstname" value="<%=rs("FirstName")%>">
        <input type="hidden" name="Gender" value="<%=left(rs("Sex"),1)%>">
        <input type="hidden" name="Birthyear" value="<%=right(rs("BirthDate"),2)%>">
        <input type="hidden" name="State" value="<%=rs("State")%>">
        <input type="submit" value="Yes - Save Now">
	</td>
        </form>
        <form method=post action="/rankings/exceptionmgmt-wsp.asp">
	<td align=center>    
          <input type="hidden" name="line" value="<%=Request("line")%>">
          <input type="hidden" name="file" value="<%=Request("file")%>">
          <input type="hidden" name="fix" value="1">
          <input type="submit" value="NO - Do Not Save">
      </td>
        </form> 
      </tr>
    </table>
    <br>	
    <table class="innertable" align="center" border=1 width=65%>

      <tr><font size="3">
			<th>&nbsp;</th>	
      <th>
        <center><strong>Original Record Data</strong>
      </th>
      <th>
        <center><strong>New Record Data</strong></center>
      </th>
      </font></tr>

      <tr>
			<td><font size="2">First: </font></td> 
	<td><% If trim(rsCSV.fields(3)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(3) & "</font></strong>") %></td>
    	<td><font size="2"><b><%=rs("FirstName")%></b></font></td>
      </tr>

      <tr>
    	<td><font size="2">Last: </font></td>
	<td><% If trim(rsCSV.fields(2)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(2) & "</font></strong>") %></td>
    	<td><font size="2"><b><%=rs("LastName")%></b></font></td>
      </tr>

      <tr>
    	<td><font size="2">Mem ID: </font></td>
	<td><% If trim(rsCSV.fields(1)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(1) & "</font></strong>") %></td>
    	<td><font size="2"><b><%=rs("PersonIDwithCheckDigit")%></b></font></td>
      </tr>

      <tr>
    	<td><font size="2">Tour ID: </font></td>
	<td><font size="2"><strong><% =right(left(Request("file"),18),7) %></strong></font></td>
	<td>&nbsp;</td>
      </tr>

      <tr>
    	<td><font size="2">Federation: </font></td> 
	<td><% If trim(rsCSV.fields(0)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(0) & "</font></strong>") %></td>
    	<td><% If trim(rs("FederationCode")) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rs("FederationCode")&"</font></strong>") %></td>
      </tr>

      <tr>
    	<td><font size="2">Gender: </font></td> 
	<td><% If trim(rsCSV.fields(4)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(4) & "</font></strong>") %></td>
    	<td><% If trim(rs("Sex")) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & left(rs("Sex"),1) & "</font></strong>") %></td>
      </tr>

      <tr>
    	<td><font size="2">Age: </font></td> 
	<td><% If trim(rsCSV.fields(5)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & Year(Now) - 1900 - rsCSV.fields(5) & "</font></strong>") %></td>
    	<td><% If IsNull(rs("BirthDate")) Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & AgeInYears & "</font></strong>") %></td>
      </tr>

      <tr>
    	<td><font size="2">State: </font></td> 
	<td><% If trim(rsCSV.fields(6)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(6) & "</font></strong>") %></td>
    	<td><% If trim(rs("State")) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rs("State") & "</font></strong>") %></td>
      </tr>

      <tr>
    	<td><font size="2">Region: </font></td> 
	<td><% If trim(rsCSV.fields(7)) = "" Then Response.Write("<font  size=""2""color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(7) & "</font></strong>") %></td>
	<td><%
               If trim(rs("State")) = "" Then 
                 Response.Write("<font size=""2"" color=gray>n/a</font>") 
               Else 
                 Select case rs("Region")
                   Case 1
                     Response.Write("<strong><font size=""2"">C</font></strong>") 
                   Case 2
                     Response.Write("<strong><font size=""2"">M</font></strong>") 
                   Case 3
                     Response.Write("<strong><font size=""2"">W</font></strong>") 
                   Case 4
                     Response.Write("<strong><font size=""2"">S</font></strong>") 
                   Case 5
                     Response.Write("<strong><font size=""2"">E</font></strong>") 
                 End Select
               End If  %>
	</td>
      </tr>

      <tr>
    	<td><font size="2">Team: </font></td> 
	<td><% If trim(rsCSV.fields(8)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(8) & "</font></strong>") %></td>
	<td>&nbsp;</td>
      </tr>

      <tr>
    	<td><font size="2">Slalom Div: </font></td> 
	<td><% If trim(rsCSV.fields(24)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(24) & "</font></strong>") %></td>
	<td><%
                If ucase(right(left(Request("file"),18),1)) = "A" Then                
                  sDate = CDate("01/01/"&(1999+(right(left(Request("file"),13),2))))
                Else
                  sDate = CDate("01/01/"&(2000+(right(left(Request("file"),13),2))))
                End If
                
                sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&sDate&"' BETWEEN BeginDate and EndDate "
                rs2.open sSQL, SConnectionToTRATable, 3, 1
                
                ' We have to redo this because division is not based on REAL age, it's based on age relative to ski year.
                If NOT IsNull(rs("BirthDate")) Then
                  ' get absolute number of years 
                  AgeInYears = cint(datediff("YYYY", rs("BirthDate"), sDate)) - 1
                  If AgeInYears > 16 Then
                    sSQL = "Select top 1 div from " & DivisionsTableName & " where left(Div,1) ='" & ucase(left(rs("Sex"),1)) &"' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& rs2("SkiYearID") &" order by Div"
                  Else
                    If ucase(left(rs("Sex"),1)) = "M" Then
                      sSQL = "Select top 1 div from " & DivisionsTableName & " where left(Div,1) = 'B' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& rs2("SkiYearID") &" order by Div"
                    Else
                      sSQL = "Select top 1 div from " & DivisionsTableName & " where left(Div,1) = 'G' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& rs2("SkiYearID") &" order by Div"
                    End If
                  End If
                Else
                  sSQL = "Select top 1 div from " & DivisionsTableName & " where 0=1"
                End If
                rs2.close ' Close the Ski Year Table
                rs2.open sSQL, SConnectionToTRATable, 3, 1 ' Open the Division table
                  If rs2.EOF Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rs2("Div") & "</font></strong>")
                rs2.close  %>
	</td>
      </tr>

      <tr>
    	<td><font size="2">Trick Div: </font></td> 
	<td><% If trim(rsCSV.fields(31)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(31) & "</font></strong>") %></td>
	<td>&nbsp;</td>	
      </tr>
      <tr>
    	<td><font size="2">Jump Div: </font></td> 
	<td><% If trim(rsCSV.fields(35)) = "" Then Response.Write("<font size=""2"" color=gray>n/a</font>") Else Response.Write("<strong><font size=""2"">" & rsCSV.fields(35) & "</font></strong>") %></td>
	<td>&nbsp;</td>	
      </tr>

    </table>
    
		<br>&nbsp;<br>    
    
    </center>
    
    <%WriteIndexPageFooter%>
    
    <%  

    rs.Close
    rsCSV.close


End If  

If request("delete") = "1" then

      If lcase(Request("confirm")) <> "yes" and lcase(Request("confirm")) <> "" then
        response.redirect ("/rankings/exceptionmgmt-wsp.asp?file=" & Request("file") & "&line=" & Request("line"))
      End If
      
      If lcase(Request("confirm")) = "yes" then
    
          ' Remove the old bad record from the exception file.
              
          dim fileoutbad
          dim fileoutexplainations
          
          linecount = 0
          fileoutbad=PathtoExceptions & "\" & request("file")
          fileoutexplainations=PathtoReasons & "\" & request("file")
    
          set objstream=objFSO.opentextfile(fileoutbad)
    
          do while not objstream.atendofStream
            objstream.skipline
          loop
          linecount = objstream.Line
          objstream.close
    
          set objstream=objFSO.opentextfile(fileoutbad)
          textFile = "" ' this will hold the contents of the text file
    
          Do While not objStream.AtEndOfStream
            strFileLine = objStream.Readline
            if objstream.line - request("line") <> 1 then
              textFile = textFile & strFileLine & vbCrLf
            end if
          Loop
    
          objstream.close
          set objstream=objfso.opentextfile(fileoutbad,2,true)
          objstream.write(textfile)
          objstream.close
    
    ' Remove the old bad reason from the reasons file.
    
          set objstream=objFSO.opentextfile(fileoutexplainations)
          textFile = "" ' this will hold the contents of the text file
    
          Do While not objStream.AtEndOfStream
            strFileLine = objStream.Readline
            if objstream.line - request("line") <> 1 then
              textFile = textFile & strFileLine & vbCrLf
            end if
          Loop
    
          objstream.close
          set objstream=objfso.opentextfile(fileoutexplainations,2,true)
          objstream.write(textfile)
          objstream.close
    
    ' Display a success message and return to exception management.
    
          if linecount-1 > 2 then
            if request("line") > 2 then
              response.redirect "/rankings/exceptionmgmt-wsp.asp?file=" & request("file") & "&line=" & request("line")-1
            else
              response.redirect "/rankings/exceptionmgmt-wsp.asp?file=" & request("file") & "&line=2"
            end if 
          else
            WriteLog(date() &"  "& time() &"  "& fileoutbad & " is now corrected and has been automatically deleted.")
            objfso.DeleteFile(fileoutbad)
            objfso.DeleteFile(fileoutexplainations)
        
            WriteIndexPageHeader
            %>
            <h2> The last exception in the file <br>
            <%=Request("file")%> 
            <br>
            has been removed.  You will now return to the main menu.</h2>
            <p>
            <%
    
            Response.Write "<a href=""/rankings/defaultHQ.asp"">Return to Main Menu.</a>"
            WriteIndexPageFooter
          end if 

      End If

      If lcase(Request("confirm")) = "" then

        WriteIndexPageHeader
        %>  
            <br><br><h3>
            Type the word "YES" if you are sure you wish to delete the record.
            </h3>
            To cancel and return without deleting, use your BACK button or type the word "NO".
            <br><br>
            <form action="/rankings/exceptionmgmt-wsp.asp" method="post"> 
            <input type="hidden" name="line" value="<%=Request("line")%>">
            <input type="hidden" name="file" value="<%=Request("file")%>">
            <input type="hidden" name="delete" value="1">
            <input type="text" name="confirm" size="5">
            <input type="submit" value="Confirm Deletion?">
            </form>
        <%
        WriteIndexPageFooter

      End If

End If

'Close our recordset and connection
conCSV.close
set rs = Nothing
set rs2 = Nothing
set rsSelectFields = Nothing
Set rsCSV = Nothing
set conCSV = nothing

objfso.DeleteFile CSVfile

CloseCon

%>




