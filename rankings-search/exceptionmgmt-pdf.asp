<!--#include file="secure-settings.asp"-->

<% 

    dim objfso
    dim objstream
    dim linecount
    dim linetext
    dim resulttext
    dim AgeInYears
    dim rs1, rs2, rs3

if (request("fix") <> "1" and request("fix") <> "2") and request("delete") <> 1 then %>
    
<html><head><title>Exception Editor</title></head><body>

<%
WriteIndexPageHeader
NewsPageNum = "4"
%>

<center>
<strong>
    <%
    response.write (right(left(request("file"),18),7))

    filespec = PathtoExceptions & "\" & request("file")

markdebug(filespec)
    
    Set objfso = CreateObject("Scripting.FileSystemObject")
    set objstream=objFSO.opentextfile(filespec)
'c=0    
    do while not objstream.atendofStream
'c=c+1
'if c=1 THEN
 'markdebug("filespec - inside ="&filespec)
'END IF
      objstream.skipline
    loop
    linecount = objstream.Line
    objstream.close
    %>
</strong><br>
Current record:
<%Response.Write Request("line")%>
 of 
<%Response.Write linecount-1%>
<br><br>

<%
set objstream=objFSO.opentextfile(filespec)

do while (not objstream.atendofStream) and (objstream.line - request("line") <> 0)
   objstream.skipline
loop

if objstream.atendofstream then
   response.redirect("/rankings/defaultHQ.asp?process=endoffile&line=" & request("line"))
else
   lineText=objstream.readline

objstream.close
%>

<%If Request("line") > 1 Then%>
<a href="/rankings/exceptionmgmt-pdf.asp?file=<%=Request("file")%>&line=<%=Request("line")-1%>">
<img src="/rankings/images/buttons/left.gif" border=0 title="Display Record Number <%=Request("line")-1%>"></a>
<%Else Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
end if%>
&nbsp;&nbsp;&nbsp;
<%If linecount - Request("line") > 1 Then%>
<a href="/rankings/exceptionmgmt-pdf.asp?file=<%=Request("file")%>&line=<%=Request("line")+1%>">
<img src="/rankings/images/buttons/right.gif" border=0 title="Display Record Number <%=Request("line")+1%>"></a>
<%Else Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
End If%>

<br><hr>
<b><u>Reasons Why the Record Failed</u></b><br><br>
<textarea rows=3 cols=50 readonly style="overflow:hidden">
<%
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
OpenCon
set rsSelectFields = Server.CreateObject("ADODB.recordset")
sSQL = "Select distinct div from " & RankTableName & " order by div"
rsSelectFields.open sSQL, SConnectionToTRATable

%>
</textarea><br>

<%   If left(resulttext,1) = "*" Then %>
     <br><center>
     <form method=post action="/rankings/exceptionmgmt-pdf.asp">
     <input type="hidden" name="line" value="<%=Request("line")%>">
     <input type="hidden" name="file" value="<%=Request("file")%>">
     <input type="hidden" name="fix" value="1">
     <input type="submit" value="Fix Member Information"></form>
     </center>
<%End If%>


<form method=post action="/rankings/updatefields.asp">
<input type="hidden" name="linenum" value="<%=Request("line")%>">
<input type="hidden" name="file" value="<%=Request("file")%>">

<TABLE class="innertable" BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=75%>

<TR>
<td colspan="2"><b><center>Participant Data</center></b></td>
</tr>

<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Member Federation</font></TD>
<%InputData = trim(left(linetext,3))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Member_Federation" rows=1 cols=25 style="overflow:hidden"><%Response.Write inputdata%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Member ID</FONT></TD>
<%InputData = trim(right(left(linetext,12),9))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Member_ID" rows=1 cols=25 style="overflow:hidden" readonly><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Last Name</FONT></TD>
<%InputData = trim(right(left(linetext,31),17))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Lastname" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">First Name</FONT></TD>
<%InputData = trim(right(left(linetext,44),13))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Firstname" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</TR>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Gender</FONT></TD>
<%InputData = trim(right(left(linetext,45),1))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Gender" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Birth Year</FONT></TD>
<%InputData = trim(right(left(linetext,47),2))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Birthyear" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">State</FONT></TD>
<%InputData = trim(right(left(linetext,49),2))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="State" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Region</FONT></TD>
<%InputData = trim(right(left(linetext,50),1))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Region" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<TR>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Team</FONT></TD>
<%InputData = trim(right(left(linetext,54),4))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Team" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</TR>

<TR>
<td colspan="2"><b><center>Tournament Data</center></b></td>
</tr>

<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Tour Federation</FONT></TD>
<%InputData = trim(right(left(linetext,63),3))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TourFederation" rows=1 cols=25 style="overflow:hidden" readonly><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Tour ID</FONT></TD>
<%InputData = trim(right(left(linetext,71),8))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TourID" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Homologation Class</FONT></TD>
<%InputData = trim(right(left(linetext,72),1))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="Homologation" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Tour End Date</FONT></TD>
<%InputData = trim(right(left(linetext,78),2))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1">
<textarea name="TourMonth" rows=1 cols=4 style="overflow:hidden"><%Response.Write InputData%></textarea>
/
<%InputData = trim(right(left(linetext,80),2))%>
<textarea name="TourDay" rows=1 cols=4 style="overflow:hidden"><%Response.Write InputData%></textarea>
/
<%InputData = trim(right(left(linetext,76),4))%>
<textarea name="TourYear" rows=1 cols=8 style="overflow:hidden"><%Response.Write InputData%></textarea>
</FONT></TD>
</TR>

<TR>
<td colspan="2"><b><center>Performance Summary Data</center></b></td>
</tr>

<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Event Placement</FONT></TD>
<%InputData = trim(right(left(linetext,93),3))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SlalomPlacement" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Round # of Best Slalom Score</FONT></TD>
<%InputData = trim(right(left(linetext,94),1))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="BestSlalomRound">
<%
if trim(InputData) = "" then
  %><option value="" selected> </option><%
else
  %><option value=""> </option><%
end if  
if trim(InputData) = "1" then
  %><option value="1" selected>1</option><%
else
  %><option value="1">1</option><%
end if
if trim(InputData) = "2" then
  %><option value="2" selected>2</option><%
else
  %><option value="2">2</option><%
end if
if trim(InputData) = "3" then
  %><option value="3" selected>3</option><%
else
  %><option value="3">3</option><%
end if
%>
</select></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Event Placement</FONT></TD>
<%InputData = trim(right(left(linetext,97),3))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TrickPlacement" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Round # of Best Trick Score</FONT></TD>
<%InputData = trim(right(left(linetext,98),1))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="BestTrickRound">
<%
if trim(InputData) = "" then
  %><option value="" selected> </option><%
else
  %><option value=""> </option><%
end if  
if trim(InputData) = "1" then
  %><option value="1" selected>1</option><%
else
  %><option value="1">1</option><%
end if
if trim(InputData) = "2" then
  %><option value="2" selected>2</option><%
else
  %><option value="2">2</option><%
end if
if trim(InputData) = "3" then
  %><option value="3" selected>3</option><%
else
  %><option value="3">3</option><%
end if
%>
</select></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Event Placement</FONT></Center></TD>
<%InputData = trim(right(left(linetext,101),3))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JumpPlacement" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Round # of Best Jump Score</FONT></TD>
<%InputData = trim(right(left(linetext,102),1))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="BestJumpRound">
<%
if trim(InputData) = "" then
  %><option value="" selected> </option><%
else
  %><option value=""> </option><%
end if  
if trim(InputData) = "1" then
  %><option value="1" selected>1</option><%
else
  %><option value="1">1</option><%
end if
if trim(InputData) = "2" then
  %><option value="2" selected>2</option><%
else
  %><option value="2">2</option><%
end if
if trim(InputData) = "3" then
  %><option value="3" selected>3</option><%
else
  %><option value="3">3</option><%
end if
%>
</select></FONT></TD>
</tr>
<tr>
<TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Number of Rounds Reported</FONT></Center></TD>
<%InputData = trim(right(left(linetext,103),1))%>
<TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="NumberofRounds" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
</tr>

<%If len(linetext) > 164 Then%>
  <TR>
  <td colspan="2"><b><center>Round 1 Score Data</center></b></td>
  </tr>

  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,112),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="SL_1_Sanction">
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
  <%InputData = trim(right(left(linetext,114),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="SL_1_Division">
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
  <%InputData = trim(right(left(linetext,116),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_1_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Score</FONT></TD>
  <%InputData = trim(right(left(linetext,120),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_1_EndPassScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Speed</FONT></TD>
  <%InputData = trim(right(left(linetext,122),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_1_EndPassSpeed" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Line</FONT></TD>
  <%InputData = trim(right(left(linetext,126),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_1_EndPassLine" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Total Score</FONT></TD>
  <%InputData = trim(right(left(linetext,131),5))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_1_TotalScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,136),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="TR_1_Sanction">
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
  <%InputData = trim(right(left(linetext,138),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="TR_1_Division">
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
  <%InputData = trim(right(left(linetext,140),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TR_1_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Total Score</FONT></TD>
  <%InputData = trim(right(left(linetext,145),5))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TR_1_TotalScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,149),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="JM_1_Sanction">
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
  <%InputData = trim(right(left(linetext,151),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="JM_1_Division">
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
  <%InputData = trim(right(left(linetext,153),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_1_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Ramp Height Ratio</FONT></TD>
  <%InputData = trim(right(left(linetext,157),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_1_RampHeight" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Boat Speed</FONT></TD>
  <%InputData = trim(right(left(linetext,159),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_1_BoatSpeed" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Best Distance (Feet)</FONT></TD>
  <%InputData = trim(right(left(linetext,162),3))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_1_DistanceFeet" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Best Distance (Meters)</FONT></TD>
  <%InputData = trim(right(left(linetext,166),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_1_DistanceMeter" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
<%End If%>

<%If trim(len(linetext)) > 224 Then%>
  <TR>
  <td colspan="2"><b><center>Round 2 Score Data</center></b></td>
  </tr>

  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,172),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="SL_2_Sanction">
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
  <%InputData = trim(right(left(linetext,174),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="SL_2_Division">
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
  <%InputData = trim(right(left(linetext,176),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_2_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Score</FONT></TD>
  <%InputData = trim(right(left(linetext,180),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_2_EndPassScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Speed</FONT></TD>
  <%InputData = trim(right(left(linetext,182),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_2_EndPassSpeed" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Line</FONT></TD>
  <%InputData = trim(right(left(linetext,186),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_2_EndPassLine" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Total Score</FONT></TD>
  <%InputData = trim(right(left(linetext,191),5))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_2_TotalScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,196),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="TR_2_Sanction">
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
  <%InputData = trim(right(left(linetext,198),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="TR_2_Division">
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
  <%InputData = trim(right(left(linetext,200),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TR_2_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Total Score</FONT></TD>
  <%InputData = trim(right(left(linetext,205),5))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TR_2_TotalScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,209),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="JM_2_Sanction">
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
  <%InputData = trim(right(left(linetext,211),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="JM_2_Division">
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
  <%InputData = trim(right(left(linetext,213),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_2_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Ramp Height Ratio</FONT></TD>
  <%InputData = trim(right(left(linetext,217),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_2_RampHeight" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Boat Speed</FONT></TD>
  <%InputData = trim(right(left(linetext,219),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_2_BoatSpeed" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Best Distance (Feet)</FONT></TD>
  <%InputData = trim(right(left(linetext,222),3))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_2_DistanceFeet" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Best Distance (Meters)</FONT></TD>
  <%InputData = trim(right(left(linetext,226),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_2_DistanceMeter" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
<%End If%>
  
<%If trim(len(linetext)) > 284 Then%>
  <TR>
  <td colspan="2"><b><center>Round 3 Score Data</center></b></td>
  </tr>

  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,232),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="SL_3_Sanction">
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
  <%InputData = trim(right(left(linetext,234),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="SL_3_Division">
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
  <%InputData = trim(right(left(linetext,236),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_3_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Score</FONT></TD>
  <%InputData = trim(right(left(linetext,240),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_3_EndPassScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Speed</FONT></TD>
  <%InputData = trim(right(left(linetext,242),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_3_EndPassSpeed" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom End Pass Line</FONT></TD>
  <%InputData = trim(right(left(linetext,246),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_3_EndPassLine" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Slalom Total Score</FONT></TD>
  <%InputData = trim(right(left(linetext,251),5))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="SL_3_TotalScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,256),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="TR_3_Sanction">
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
  <%InputData = trim(right(left(linetext,258),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="TR_3_Division">
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
  <%InputData = trim(right(left(linetext,260),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TR_3_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Trick Total Score</FONT></TD>
  <%InputData = trim(right(left(linetext,265),5))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="TR_3_TotalScore" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
  <TR>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Sanction Class</FONT></TD>
  <%InputData = trim(right(left(linetext,269),1))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="JM_3_Sanction">
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
  <%InputData = trim(right(left(linetext,271),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><select name="JM_3_Division">
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
  <%InputData = trim(right(left(linetext,273),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_3_Boat" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Ramp Height Ratio</FONT></TD>
  <%InputData = trim(right(left(linetext,277),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_3_RampHeight" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Jump Boat Speed</FONT></TD>
  <%InputData = trim(right(left(linetext,279),2))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_3_BoatSpeed" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Best Distance (Feet)</FONT></TD>
  <%InputData = trim(right(left(linetext,282),3))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_3_DistanceFeet" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </tr>
  <tr>
  <TD ALIGN="Left" vAlign="center"><FONT COlOR="#000000" SIZE="2">Best Distance (Meters)</FONT></TD>
  <%InputData = trim(right(left(linetext,286),4))%>
  <TD ALIGN="Left" vAlign="center" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><textarea name="JM_3_DistanceMeter" rows=1 cols=25 style="overflow:hidden"><%Response.Write InputData%></textarea></FONT></TD>
  </TR>
<%End If
  
if instr(right(left(linetext,113),5),"1") then %>
  <input type="hidden" name="Round1" value="1">
<%End If
if instr(right(left(linetext,173),5),"2") then %>
  <input type="hidden" name="Round2" value="1">
<%End If
if instr(right(left(linetext,233),5),"3") then %>
  <input type="hidden" name="Round3" value="1">
<%End If%>






</table>
<br><center><input type="submit" value="Save New Field Data"></center>
</form>
<br>
<form method=post action="/rankings/exceptionmgmt-pdf.asp">
<input type="hidden" name="line" value="<%=Request("line")%>">
<input type="hidden" name="file" value="<%=Request("file")%>">
<input type="hidden" name="delete" value="1">
<center><input type="submit" value="Delete This Record"></center>
</form>

<%
CloseCon
rsSelectFields.close

'This code allowed the user to edit the raw data line rather then breaking
'the data out into the various fields.  If you understand the format
'of the data line, sometimes this is easier.  If you aren't careful, it can
'cause a ton of problems though.
'
'Mark decided to take it out of the final version.
'
'
'  <form method=post action="updateraw.asp">
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
%>


<br><br><%
end if
%>
<center>
<%If Request("line") > 1 Then%>
<a href="/rankings/exceptionmgmt-pdf.asp?file=<%=Request("file")%>&line=<%=Request("line")-1%>">
<img src="/rankings/images/buttons/left.gif" border=0 title="Display Record Number <%=Request("line")-1%>"></a>
<%Else Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
end if%>
&nbsp;&nbsp;&nbsp;
<%If linecount - Request("line") > 1 Then%>
<a href="/rankings/exceptionmgmt-pdf.asp?file=<%=Request("file")%>&line=<%=Request("line")+1%>">
<img src="/rankings/images/buttons/right.gif" border=0 title="Display Record Number <%=Request("line")+1%>"></a>
<%Else Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
end if%>
<br><br><a href="/rankings/defaultHQ.asp">Return to the Main Menu</a>
</center>

<%WriteIndexPageFooter%>


<%

End If

if request("fix") = "1" then 

%>

<html><head><title>Exception Editor</title></head><body>

<%WriteIndexPageHeader%>

<%
NewsPageNum="5"
%>


<center>
<br><h3>
Member Data Correction<br><br>    
</h3>
<hr><br>


<%
filespec = PathtoExceptions & "\" & request.form("file")

Set objfso = CreateObject("Scripting.FileSystemObject")
set objstream=objFSO.opentextfile(filespec)

do while not objstream.atendofStream
  objstream.skipline
loop
linecount = objstream.Line
objstream.close

set objstream=objFSO.opentextfile(filespec)

do while (not objstream.atendofStream) and (objstream.line - request.form("line") <> 0)
   objstream.skipline
loop

if objstream.atendofstream then
   Response.Redirect("/rankings/defaultHQ.asp?process=endoffile&line=" & Request.Form("line"))
else
   lineText=objstream.readline 
   objstream.close
end if
OpenCon
Set rs=Server.CreateObject("ADODB.recordset")
%>
    
<table border=1 CELLPADDING="3" CELLSPACING="0">
<tr>
  <td>
    <center><strong>Original Record Data</strong>
  </td>
  <td>
    <center><strong>Possible Matches</strong></center>
  </td>
</tr>
<tr><td>
First Name: <% If trim(right(left(linetext,44),13)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,44),13)) & "</strong>" %><br>
Last Name: <% If trim(right(left(linetext,31),17)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,31),17)) & "</strong>" %><br>
Member ID: <% If trim(right(left(linetext,12),9)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,12),9)) & "</strong>" %><br>
Tour ID: <strong><%

sSQL = "Select top 1 TSanction,TName,TCity,TState,TDateE from "& SanctionTableName &" where lower(TournAppID) = '" & sqlclean(lcase(right(left(Request("file"),17),6))) & "'"
rs.open sSQL, sConnectionToTRATable, 3, 1
If NOT rs.EOF Then
  %>
  <a title="<% =rs("tname") %>&#13;<% =rs("tcity")%>, <% =rs("tstate")%>&#13;<% =rs("tdatee")%>"><u>
  <% 
End If
Response.Write (right(left(Request("file"),18),7))
Response.Write ("</u></a></strong><br>")
rs.Close %>


Federation: <% If trim(left(linetext,3)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & left(linetext,3)) & "</strong>" %><br>
Gender: <% If trim(right(left(linetext,45),1)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,45),1)) & "</strong>" %><br>
Birth Year: <% If trim(right(left(linetext,47),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,47),2)) & "</strong>" %><br>
State: <% If trim(right(left(linetext,49),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,49),2)) & "</strong>" %><br>
Region: <% If trim(right(left(linetext,50),1)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,50),1)) & "</strong>" %><br>
Team: <% If trim(right(left(linetext,54),4)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,54),4)) & "</strong>" %><br>
Slalom Division: <% If trim(right(left(linetext,114),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,114),2)) & "</strong>" %><br>
Trick Division: <% If trim(right(left(linetext,138),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,138),2)) & "</strong>" %><br>
Jump Division: <% If trim(right(left(linetext,151),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,151),2)) & "</strong>" %><br>
</td>
<td>
<%

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
    
  if len(sSQL) < 166 then
    sSQL = "Select * from members where 1=0"
  end if

  rs.open sSQL, sConnectionToMemberTable, 3, 1
  %>
    <center>    
    <br><br><b>Search Results</b>
    <% If rs.EOF Then %>
        <br><font color="red"> No Records Found, Please Search Again </font><br><br>
    <% Else %>
        <TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=50%>
        <TR>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">Member ID</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">First Name</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">Last Name</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">State</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">Age</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">Gender</FONT></TD>
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
            <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><a href="/rankings/exceptionmgmt-pdf.asp?fix=2&line=<%=Request("line")%>&file=<%=Request("file")%>&MemberID=<%=rs("PersonIDwithCheckDigit")%>"><%=rs("PersonIDwithCheckDigit")%></a></FONT></TD>
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
            <font color="red"><small><center>Only the top ten records were displayed.</center></small></font><br>
        <% End If %>
        </center>
  <% End If
    rs.Close

ELSE
  sSQL = "Select FROMsubQ.PersonIDwithCheckDigit, min(FROMsubQ.lastname) as LastName, min(FROMsubQ.firstname) as FirstName, min(FROMsubQ.state) as State, min(FROMsubQ.birthdate) as BirthDate, min(FROMsubQ.sex) as Sex, min(FROMsubQ.Rank) as NewRank from ("
  'Nasty SubQuery Alert!!

  ' First we select everyone with a matching Member ID
  sSQL = sSQL + "Select top 3 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex,'1' as Rank from "&MemberTableName&" where "
  sSQL = sSQL + "PersonIDwithCheckDigit = " & SQLClean(trim(right(left(linetext,12),9))) & " and "
  sSQL = sSQL + "membertypeid <> 2"
  sSQL = sSQL + " UNION " 

  ' Next we select everyone who has 3 char matches for both first and last name
  sSQL = sSQL + "Select top 3 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex,'2' as Rank from "&MemberTableName&" where "
  sSQL = sSQL + "difference(lower(lastname),'" & SQLClean(lcase(trim(right(left(linetext,31),17)))) & "') = 4 and "
  sSQL = sSQL + "difference(lower(firstname),'" & SQLClean(lcase(trim(right(left(linetext,44),13)))) & "') = 4 and "
  If trim(right(left(linetext,45),1)) <> "" then
    sSQL = sSQL + "left(lower(sex),1) = '" & SQLClean(lcase(trim(right(left(linetext,45),1)))) & "' and "
  End If
  sSQL = sSQL + "membertypeid <> 2"
  sSQL = sSQL + " UNION " 
  
  ' Then we check for left 4 digits of the Last Name + left 1 digit of first name
  sSQL = sSQL + "Select top 3 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex,'3' as Rank from "&MemberTableName&" where "
  sSQL = sSQL + "lower(lastname) LIKE '%" & SQLClean(lcase(trim(right(left(linetext,18),4)))) & "%' and "
  sSQL = sSQL + "left(lower(firstname),1) = '" & SQLClean(lcase(trim(right(left(linetext,32),1)))) & "' and "
  If trim(right(left(linetext,45),1)) <> "" then
    sSQL = sSQL + "left(lower(sex),1) = '" & SQLClean(lcase(trim(right(left(linetext,45),1)))) & "' and "
  End If
  sSQL = sSQL + "membertypeid <> 2"
  sSQL = sSQL + " UNION " 

  ' Then we check for left 4 digits of the First Name + left 1 digit of last name
  sSQL = sSQL + "Select top 3 PersonIDwithCheckDigit,LastName,FirstName,State,BirthDate,Sex,'4' as Rank from "&MemberTableName&" where "
  sSQL = sSQL + "left(lower(lastname),1) = '" & SQLClean(lcase(trim(right(left(linetext,15),1)))) & "' and "
  sSQL = sSQL + "lower(firstname) LIKE '%" & SQLClean(lcase(trim(right(left(linetext,35),4)))) & "%' and "
  If trim(right(left(linetext,45),1)) <> "" then
    sSQL = sSQL + "left(lower(sex),1) = '" & SQLClean(lcase(trim(right(left(linetext,45),1)))) & "' and "
  End If
  sSQL = sSQL + "membertypeid <> 2"
  sSQL = sSQL + ") as FROMsubQ group by PersonIDwithCheckDigit order by NewRank"
  
   
  rs.open sSQL, sConnectionToTRATable, 3, 1
  If rs.EOF Then %>
    <br><br><br><font color=red>Unable to find potential match.</font>
<% Else %>
    <br><br><br>
    <TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=50%>
      <TR>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">Member ID</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">First Name</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">Last Name</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">State</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">Age</FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="2">Gender</FONT></TD>
      </tr>
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
        <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><a href="/rankings/exceptionmgmt-pdf.asp?fix=2&line=<%=Request("line")%>&file=<%=Request("file")%>&MemberID=<%=rs("PersonIDwithCheckDigit")%>"><%=rs("PersonIDwithCheckDigit")%></a></FONT></TD>
        <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("FirstName")%></FONT></TD>
        <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("LastName")%></FONT></TD>
        <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("State")%></FONT></TD>
        <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=AgeInYears%></FONT></TD>
        <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" SIZE="2"><%=rs("Sex")%></FONT></TD>
      </tr>
   <% rs.MoveNext %>
   <% Loop %>
   </table><br><br>
<% End If 
   rs.Close

END IF

Set rs = Nothing
CloseCon
%>
</td>
</tr>
</table>
<br>

<form method=post action="/rankings/exceptionmgmt-pdf.asp">
<input type="hidden" name="line" value="<%=Request("line")%>">
<input type="hidden" name="file" value="<%=Request("file")%>">
<input type="hidden" name="fix" value="0">
<input type="submit" value="Return Without Making Changes">
</form>

<b><center>Search Membership Database</center></b>
<form method=post action="/rankings/exceptionmgmt-pdf.asp">
<input type="hidden" name="line" value="<%=Request("line")%>">
<input type="hidden" name="file" value="<%=Request("file")%>">
<input type="hidden" name="fix" value="1">
<input type="hidden" name="search" value="1">
<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=60%>
<TR>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" SIZE="2">Member ID</FONT></Center></TD>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" SIZE="2">First Name</FONT></Center></TD>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" SIZE="2">Last Name</FONT></Center></TD>
<TD ALIGN="Center" vAlign="top"><Center><FONT COlOR="#000000" SIZE="2">State</FONT></TD>
<TD ALIGN="Center" vAlign="top"><Center><FONT COlOR="#000000" SIZE="2">Date of Birth</FONT></TD>
<TD ALIGN="Center" vAlign="top"><Center><FONT COlOR="#000000" SIZE="2">Gender</FONT></TD>
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
<input type="submit" value="Search Member Database"></form>

<%

WriteIndexPageFooter

End If 


if request("fix") = "2" then

%>

<html><head><title>Exception Editor</title></head><body>

<%WriteIndexPageHeader%>

<center>
<br><h3>
Member Data Correction    
</h3>

<br>

<%

filespec = PathtoExceptions & "\" & request("file")

Set objfso = CreateObject("Scripting.FileSystemObject")
set objstream=objFSO.opentextfile(filespec)

do while not objstream.atendofStream
  objstream.skipline
loop
linecount = objstream.Line
objstream.close


set objstream=objFSO.opentextfile(filespec)

do while (not objstream.atendofStream) and (objstream.line - request("line") <> 0)
   objstream.skipline
loop

if objstream.atendofstream then
   Response.Redirect("/rankings/defaultHQ.asp?process=endoffile&line=" & Request("line"))
else
   lineText=objstream.readline 
   objstream.close
end if
%>


<form method=post action="/rankings/updatefields.asp">
    <input type="hidden" name="linenum" value="<%=Request("line")%>">
    <input type="hidden" name="file" value="<%=Request("file")%>">
    <input type="hidden" name="TourFederation" value="<%=right(left(linetext,63),3)%>">
    <input type="hidden" name="TourID" value="<%=right(left(linetext,71),8)%>">
    <input type="hidden" name="Homologation" value="<%=right(left(linetext,72),1)%>">
    <input type="hidden" name="TourMonth" value="<%=right(left(linetext,78),2)%>">
    <input type="hidden" name="TourDay" value="<%=right(left(linetext,80),2)%>">
    <input type="hidden" name="TourYear" value="<%=right(left(linetext,76),4)%>">
    <input type="hidden" name="SlalomPlacement" value="<%=right(left(linetext,93),3)%>">
    <input type="hidden" name="BestSlalomRound" value="<%=right(left(linetext,94),1)%>">
    <input type="hidden" name="TrickPlacement" value="<%=right(left(linetext,97),3)%>">
    <input type="hidden" name="BestTrickRound" value="<%=right(left(linetext,98),1)%>">
    <input type="hidden" name="JumpPlacement" value="<%=right(left(linetext,101),3)%>">
    <input type="hidden" name="BestJumpRound" value="<%=right(left(linetext,102),1)%>">
    <input type="hidden" name="NumberofRounds" value="<%=right(left(linetext,103),1)%>">
<%If len(linetext) > 164 Then%>
    <input type="hidden" name="SL_1_Sanction" value="<%=right(left(linetext,112),1)%>">
    <input type="hidden" name="SL_1_Division" value="<%=right(left(linetext,114),2)%>">
    <input type="hidden" name="SL_1_Boat" value="<%=right(left(linetext,116),2)%>">
    <input type="hidden" name="SL_1_EndPassScore" value="<%=right(left(linetext,120),4)%>">
    <input type="hidden" name="SL_1_EndPassSpeed" value="<%=right(left(linetext,122),2)%>">
    <input type="hidden" name="SL_1_EndPassLine" value="<%=right(left(linetext,126),4)%>">
    <input type="hidden" name="SL_1_TotalScore" value="<%=right(left(linetext,131),5)%>">
    <input type="hidden" name="TR_1_Sanction" value="<%=right(left(linetext,136),1)%>">
    <input type="hidden" name="TR_1_Division" value="<%=right(left(linetext,138),2)%>">
    <input type="hidden" name="TR_1_Boat" value="<%=right(left(linetext,140),2)%>">
    <input type="hidden" name="TR_1_TotalScore" value="<%=right(left(linetext,145),5)%>">
    <input type="hidden" name="JM_1_Sanction" value="<%=right(left(linetext,149),1)%>">
    <input type="hidden" name="JM_1_Division" value="<%=right(left(linetext,151),2)%>">
    <input type="hidden" name="JM_1_Boat" value="<%=right(left(linetext,153),2)%>">
    <input type="hidden" name="JM_1_RampHeight" value="<%=right(left(linetext,157),4)%>">
    <input type="hidden" name="JM_1_BoatSpeed" value="<%=right(left(linetext,159),2)%>">
    <input type="hidden" name="JM_1_DistanceFeet" value="<%=right(left(linetext,162),3)%>">
    <input type="hidden" name="JM_1_DistanceMeter" value="<%=right(left(linetext,166),4)%>">
<%End If%>
<%If len(linetext) > 224 Then%>
    <input type="hidden" name="SL_2_Sanction" value="<%=right(left(linetext,172),1)%>">
    <input type="hidden" name="SL_2_Division" value="<%=right(left(linetext,174),2)%>">
    <input type="hidden" name="SL_2_Boat" value="<%=right(left(linetext,176),2)%>">
    <input type="hidden" name="SL_2_EndPassScore" value="<%=right(left(linetext,180),4)%>">
    <input type="hidden" name="SL_2_EndPassSpeed" value="<%=right(left(linetext,182),2)%>">
    <input type="hidden" name="SL_2_EndPassLine" value="<%=right(left(linetext,186),4)%>">
    <input type="hidden" name="SL_2_TotalScore" value="<%=right(left(linetext,191),5)%>">
    <input type="hidden" name="TR_2_Sanction" value="<%=right(left(linetext,196),1)%>">
    <input type="hidden" name="TR_2_Division" value="<%=right(left(linetext,198),2)%>">
    <input type="hidden" name="TR_2_Boat" value="<%=right(left(linetext,200),2)%>">
    <input type="hidden" name="TR_2_TotalScore" value="<%=right(left(linetext,205),5)%>">
    <input type="hidden" name="JM_2_Sanction" value="<%=right(left(linetext,209),1)%>">
    <input type="hidden" name="JM_2_Division" value="<%=right(left(linetext,211),2)%>">
    <input type="hidden" name="JM_2_Boat" value="<%=right(left(linetext,213),2)%>">
    <input type="hidden" name="JM_2_RampHeight" value="<%=right(left(linetext,217),4)%>">
    <input type="hidden" name="JM_2_BoatSpeed" value="<%=right(left(linetext,219),2)%>">
    <input type="hidden" name="JM_2_DistanceFeet" value="<%=right(left(linetext,222),3)%>">
    <input type="hidden" name="JM_2_DistanceMeter" value="<%=right(left(linetext,226),4)%>">
<%End If%>
<%If len(linetext) > 284 Then%>
    <input type="hidden" name="SL_3_Sanction" value="<%=right(left(linetext,232),1)%>">
    <input type="hidden" name="SL_3_Division" value="<%=right(left(linetext,234),2)%>">
    <input type="hidden" name="SL_3_Boat" value="<%=right(left(linetext,236),2)%>">
    <input type="hidden" name="SL_3_EndPassScore" value="<%=right(left(linetext,240),4)%>">
    <input type="hidden" name="SL_3_EndPassSpeed" value="<%=right(left(linetext,242),2)%>">
    <input type="hidden" name="SL_3_EndPassLine" value="<%=right(left(linetext,246),4)%>">
    <input type="hidden" name="SL_3_TotalScore" value="<%=right(left(linetext,251),5)%>">
    <input type="hidden" name="TR_3_Sanction" value="<%=right(left(linetext,256),1)%>">
    <input type="hidden" name="TR_3_Division" value="<%=right(left(linetext,258),2)%>">
    <input type="hidden" name="TR_3_Boat" value="<%=right(left(linetext,260),2)%>">
    <input type="hidden" name="TR_3_TotalScore" value="<%=right(left(linetext,265),5)%>">
    <input type="hidden" name="JM_3_Sanction" value="<%=right(left(linetext,269),1)%>">
    <input type="hidden" name="JM_3_Division" value="<%=right(left(linetext,271),2)%>">
    <input type="hidden" name="JM_3_Boat" value="<%=right(left(linetext,273),2)%>">
    <input type="hidden" name="JM_3_RampHeight" value="<%=right(left(linetext,277),4)%>">
    <input type="hidden" name="JM_3_BoatSpeed" value="<%=right(left(linetext,279),2)%>">
    <input type="hidden" name="JM_3_DistanceFeet" value="<%=right(left(linetext,282),3)%>">
    <input type="hidden" name="JM_3_DistanceMeter" value="<%=right(left(linetext,286),4)%>">
    <input type="hidden" name="Region" value="<%=right(left(linetext,50),1)%>">
    <input type="hidden" name="Team" value="<%=right(left(linetext,54),4)%>">

<%End If
if instr(right(left(linetext,113),5),"1") then %>
  <input type="hidden" name="Round1" value="1">
<%End If
if instr(right(left(linetext,173),5),"2") then %>
  <input type="hidden" name="Round2" value="1">
<%End If
if instr(right(left(linetext,233),5),"3") then %>
  <input type="hidden" name="Round3" value="1">
<%End If 


    OpenCon
    Set rs=Server.CreateObject("ADODB.recordset")
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



<table border=1 CELLPADDING="8" CELLSPACING="3">
<tr>
  <td>
    <center><strong>Original Record Data</strong>
  </td>
  <td>
    <center><strong>New Record Data</strong></center>
  </td>
  <td rowspan=2>
    <br><br><br>
    Are you sure you wish<br>
    to save these changes? <br><br>
    <center>
    <input type="hidden" name="Member_Federation" value="<%=rs("FederationCode")%>">
    <input type="hidden" name="Member_ID" value="<%=rs("PersonIDwithCheckDigit")%>">
    <input type="hidden" name="Lastname" value="<%=rs("LastName")%>">
    <input type="hidden" name="Firstname" value="<%=rs("FirstName")%>">
    <input type="hidden" name="Gender" value="<%=left(rs("Sex"),1)%>">
    <input type="hidden" name="Birthyear" value="<%=right(rs("BirthDate"),2)%>">
    <input type="hidden" name="State" value="<%=rs("State")%>">
    <input type="submit" value="Yes - Save Now">
    </form>

    <form method=post action="/rankings/exceptionmgmt-pdf.asp">
      <input type="hidden" name="line" value="<%=Request("line")%>">
      <input type="hidden" name="file" value="<%=Request("file")%>">
      <input type="hidden" name="fix" value="1">
      <input type="submit" value="NO - Do Not Save">
    </form> 
    </center>
  </td>
</tr>
<tr><td>
F.Name: <% If trim(right(left(linetext,44),13)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,44),13)) & "</strong>" %><br>
L.Name: <% If trim(right(left(linetext,31),17)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,31),17)) & "</strong>" %><br>
Mem ID: <% If trim(right(left(linetext,12),9)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,12),9)) & "</strong>" %><br>
Tour ID: <strong><% =right(left(Request("file"),18),7) %></strong><br>
Fed: <% If trim(left(linetext,3)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & left(linetext,3)) & "</strong>" %><br>
Gender: <% If trim(right(left(linetext,45),1)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,45),1)) & "</strong>" %><br>
Age: <% If trim(right(left(linetext,47),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & Year(Now) - 1900 - (right(left(linetext,47),2))) & "</strong>" %><br>
State: <% If trim(right(left(linetext,49),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,49),2)) & "</strong>" %><br>
Region: <% If trim(right(left(linetext,50),1)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,50),1)) & "</strong>" %><br>
Team: <% If trim(right(left(linetext,54),4)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,54),4)) & "</strong>" %><br>
Slalom Div: <% If trim(right(left(linetext,114),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,114),2)) & "</strong>" %><br>
Trick Div: <% If trim(right(left(linetext,138),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,138),2)) & "</strong>" %><br>
Jump Div: <% If trim(right(left(linetext,151),2)) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & right(left(linetext,151),2)) & "</strong>" %><br>
</td>

<td>
F.Name: <% Response.Write("<strong>" & rs("FirstName") & "</strong>") %><br>
L.Name: <% Response.Write("<strong>" & rs("LastName") & "</strong>") %><br>
Mem ID: <% Response.Write("<strong>" & rs("PersonIDwithCheckDigit") & "</strong>") %><br>
<br>
Fed: <% If trim(rs("FederationCode")) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & rs("FederationCode") & "</strong>") %><br>
Gender: <% If trim(rs("Sex")) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & left(rs("Sex"),1) & "</strong>") %><br>
Age: <% If IsNull(rs("BirthDate")) Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & AgeInYears & "</strong>") %><br>
State: <% If trim(rs("State")) = "" Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & rs("State") & "</strong>") %><br>
Region: <% 
           If trim(rs("State")) = "" Then 
             Response.Write("<font color=gray>n/a</font>") 
           Else 
             Select case rs("Region")
               Case 1
                 Response.Write("<strong>C</strong>") 
               Case 2
                 Response.Write("<strong>M</strong>") 
               Case 3
                 Response.Write("<strong>W</strong>") 
               Case 4
                 Response.Write("<strong>S</strong>") 
               Case 5
                 Response.Write("<strong>E</strong>") 
             End Select
           End If
         %><br>
<br>
Div: <%
            sDate = CDate("01/01/"&(2000+(right(left(Request("file"),13),2))))
            
            set rs2 = Server.CreateObject("ADODB.recordset")
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
              If rs2.EOF Then Response.Write("<font color=gray>n/a</font>") Else Response.Write("<strong>" & rs2("Div") & "</strong>")
            rs2.close
            set rs2 = Nothing
          %> <br>
<br><br>
</td>
</tr>
</table>



</center>

<%WriteIndexPageFooter%>

<%  

    rs.Close
    Set rs = Nothing
    CloseCon


End If  

if request("delete") = 1 then

  If lcase(Request("confirm")) <> "yes" and lcase(Request("confirm")) <> "" then
    Response.Redirect ("/rankings/exceptionmgmt-pdf.asp?file=" & Request("file") & "&line=" & Request("line"))
  End If
  
  If lcase(Request("confirm")) = "yes" then

' Remove the old bad record from the exception file.
      
      dim fileoutbad
      dim fileoutexplainations
      
      linecount = 0
      fileoutbad=PathtoExceptions & "\" & request("file")
      fileoutexplainations=PathtoReasons & "\" & request("file")

      set objFSO=server.createobject("scripting.filesystemObject")
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

      if linecount-1 > 1 then
        if request("line") > 1 then
          response.redirect "/rankings/exceptionmgmt-pdf.asp?file=" & request("file") & "&line=" & request("line")-1
        else
          Response.Redirect "/rankings/exceptionmgmt-pdf.asp?file=" & Request("file") & "&line=1"
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
    <form action="/rankings/exceptionmgmt-pdf.asp" method="post"> 
    <input type="hidden" name="line" value="<%=Request("line")%>">
    <input type="hidden" name="file" value="<%=Request("file")%>">
    <input type="hidden" name="delete" value="1">
    <input type="text" name="confirm" size="5">
    <input type="submit" value="Confirm Deletion?">
    </form>
<%
WriteIndexPageFooter

End If

end if

%>




