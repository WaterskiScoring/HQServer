<!--#include virtual="/rankings/secure-settings.asp"-->
<%
dim objfso
dim updatedlinePDF
dim updatedlineWSP
dim inputdata
dim objstream
dim linecount
dim filespec
dim i


filespec = PathtoExceptions & "\" & Request.Form("file")

Set objfso = Server.CreateObject("Scripting.FileSystemObject")

'markdebug("UpdateFields Line 17")

set objstream=objFSO.opentextfile(filespec)

textFile = "" ' this will hold the contents of the text file

updatedlineWSP = """" & request.form("Member_Federation") & """," & _
  """" & request.form("Member_ID") & """," & _
  """" & request.form("lastname") & """," & _
  """" & request.form("Firstname") & """," & _
  """" & request.form("Gender") & """," & _
  request.form("Birthyear") & "," & _
  """" & request.form("State") & """," & _
  """" & request.form("Region") & """," & _
  """" & request.form("Team") & """," & _
  request.form("NumberofRounds") & "," & _
  """" & request.form("SlalomPlacement") & """," & _
  request.form("SlalomPlacementPoints") & "," & _
  request.form("BestSlalomRound") & "," & _
  """" & request.form("TrickPlacement") & """," & _
  request.form("TrickPlacementPoints") & "," & _
  request.form("BestTrickRound") & "," & _
  """" & request.form("JumpPlacement") & """," & _
  request.form("JumpPlacementPoints") & "," & _
  request.form("BestJumpRound") & "," & _
  """" & request.form("OverAllPlacement") & """," & _
  request.form("OverAllPlacementPoints") & "," & _
  request.form("BestOverAllRound")
For i = 1 to 9
  If request.form("round" & i) = 1 Then
    updatedlineWSP = updatedlineWSP & "," & i & "," & _
    """" & request.form("SL_" & i & "_Sanction") & """," & _
    """" & request.form("SL_" & i & "_Division") & """," & _
    """" & request.form("SL_" & i & "_Boat") & """," & _
    request.form("SL_" & i & "_EndPassScore") & "," & _
    request.form("SL_" & i & "_EndPassSpeed") & "," & _
    request.form("SL_" & i & "_EndPassLine") & "," & _
    request.form("SL_" & i & "_TotalScore") & "," & _
    """" & request.form("TR_" & i & "_Sanction") & """," & _
    """" & request.form("TR_" & i & "_Division") & """," & _
    """" & request.form("TR_" & i & "_Boat") & """," & _
    request.form("TR_" & i & "_TotalScore") & "," & _
    """" & request.form("JM_" & i & "_Sanction") & """," & _
    """" & request.form("JM_" & i & "_Division") & """," & _
    """" & request.form("JM_" & i & "_Boat") & """," & _
    request.form("JM_" & i & "_RampHeight") & "," & _
    request.form("JM_" & i & "_BoatSpeed") & "," & _
    request.form("JM_" & i & "_DistanceFeet") & "," & _
    request.form("JM_" & i & "_DistanceMeter") & "," & _
    """" & request.form("OverAll_" & i & "_Sanction") & """," & _
    """" & request.form("OverAll_" & i & "_Division") & """," & _
    request.form("OverAll_" & i & "_Score")
  End If
Next



updatedlinePDF = left(request.form("Member_Federation") & "   ",3) & _
  left(request.form("Member_ID") & "         ",9) &  _
  "  " & _
  left(request.form("lastname") & "                 ",17) & _
  left(request.form("Firstname") & "             ",13) & _
  left(request.form("Gender") & " ",1) & _
  left(request.form("Birthyear") & "  ",2) & _
  left(request.form("State") & "  ",2) & _
  left(request.form("Region") & " ",1) & _
  left(request.form("Team") & "    ",4) & _
  "      " & _
  left(request.form("TourFederation") & "   ",3) & _
  left(request.form("TourID") & "        ",8) & _
  left(request.form("Homologation") & " ",1) & _
  left(request.form("TourYear") & "    ",4) & _
  left(request.form("TourMonth") & "  ",2) & _
  left(request.form("TourDay") & "  ",2) & _
  "          " & _
  left(request.form("SlalomPlacement") & "   ",3) & _
  left(request.form("BestSlalomRound") & " ",1) & _
  left(request.form("TrickPlacement") & "   ",3) & _
  left(request.form("BestTrickRound") & " ",1) & _
  left(request.form("JumpPlacement") & "   ",3) & _
  left(request.form("BestJumpRound") & " ",1) & _
  left(request.form("NumberofRounds") & " ",1)
  if request.form("round1") = 1 then
    updatedlinePDF = updatedlinePDF & "       1" & _
    left(request.form("SL_1_Sanction") & " ",1) & _
    left(request.form("SL_1_Division") & "  ",2) & _
    left(request.form("SL_1_Boat") & "  ",2) & _
    left(request.form("SL_1_EndPassScore") & "    ",4) & _
    left(request.form("SL_1_EndPassSpeed") & "  ",2) & _
    left(request.form("SL_1_EndPassLine") & "    ",4) & _
    left(request.form("SL_1_TotalScore") & "     ",5) & _
    "    " & _
    left(request.form("TR_1_Sanction") & " ",1) & _
    left(request.form("TR_1_Division") & "  ",2) & _
    left(request.form("TR_1_Boat") & "  ",2) & _
    right("     " & request.form("TR_1_TotalScore"),5) & _
    "   " & _
    left(request.form("JM_1_Sanction") & " ",1) & _
    left(request.form("JM_1_Division") & "  ",2) & _
    left(request.form("JM_1_Boat") & "  ",2) & _
    left(request.form("JM_1_RampHeight") & "    ",4) & _
    left(request.form("JM_1_BoatSpeed") & "  ",2) & _
    left(request.form("JM_1_DistanceFeet") & "   ",3) & _
    left(request.form("JM_1_DistanceMeter") & "    ",4)
  end if
  if request.form("round2") = 1 then
    updatedlinePDF = updatedlinePDF & "    2" & _
    left(request.form("SL_2_Sanction") & " ",1) & _
    left(request.form("SL_2_Division") & "  ",2) & _
    left(request.form("SL_2_Boat") & "  ",2) & _
    left(request.form("SL_2_EndPassScore") & "    ",4) & _
    left(request.form("SL_2_EndPassSpeed") & "  ",2) & _
    left(request.form("SL_2_EndPassLine") & "    ",4) & _
    left(request.form("SL_2_TotalScore") & "     ",5) & _
    "    " & _
    left(request.form("TR_2_Sanction") & " ",1) & _
    left(request.form("TR_2_Division") & "  ",2) & _
    left(request.form("TR_2_Boat") & "  ",2) & _
    right("     " & request.form("TR_2_TotalScore"),5) & _
    "   " & _
    left(request.form("JM_2_Sanction") & " ",1) & _
    left(request.form("JM_2_Division") & "  ",2) & _
    left(request.form("JM_2_Boat") & "  ",2) & _
    left(request.form("JM_2_RampHeight") & "    ",4) & _
    left(request.form("JM_2_BoatSpeed") & "  ",2) & _
    left(request.form("JM_2_DistanceFeet") & "   ",3) & _
    left(request.form("JM_2_DistanceMeter") & "    ",4)
  end if
  if request.form("round3") = 1 then
    updatedlinePDF = updatedlinePDF & "    3" & _
    left(request.form("SL_3_Sanction") & " ",1) & _
    left(request.form("SL_3_Division") & "  ",2) & _
    left(request.form("SL_3_Boat") & "  ",2) & _
    left(request.form("SL_3_EndPassScore") & "    ",4) & _
    left(request.form("SL_3_EndPassSpeed") & "  ",2) & _
    left(request.form("SL_3_EndPassLine") & "    ",4) & _
    left(request.form("SL_3_TotalScore") & "     ",5) & _
    "    " & _
    left(request.form("TR_3_Sanction") & " ",1) & _
    left(request.form("TR_3_Division") & "  ",2) & _
    left(request.form("TR_3_Boat") & "  ",2) & _
    right("     " & request.form("TR_3_TotalScore"),5) & _
    "   " & _
    left(request.form("JM_3_Sanction") & " ",1) & _
    left(request.form("JM_3_Division") & "  ",2) & _
    left(request.form("JM_3_Boat") & "  ",2) & _
    left(request.form("JM_3_RampHeight") & "    ",4) & _
    left(request.form("JM_3_BoatSpeed") & "  ",2) & _
    left(request.form("JM_3_DistanceFeet") & "   ",3) & _
    left(request.form("JM_3_DistanceMeter") & "    ",4)
  end if

Do While not objStream.AtEndOfStream
  strFileLine = objStream.Readline
  '
  ' This process reads through the entire exceptions file
  ' until it finds the right line (linenum - 1)
  ' and then it replaces that line with the new data
  ' that we built above.
  ' 
  if objstream.line - request.form("linenum") - 1 = 0 then
     If ucase(Right(Request.Form("file"),3)) = "WSP" Then textfile = textfile & updatedlineWSP & vbCrLf
     If ucase(Right(Request.Form("file"),3)) = "PDF" Then textfile = textfile & updatedlinePDF & vbCrLf
  else
     textFile = textFile & strFileLine & vbCrLf
  end if
Loop

objstream.close

set objstream=objfso.opentextfile(filespec,2,true)

objstream.write(textfile)

objstream.close

'markdebug("UpdateFields Line 193")
'markdebug(""&PathToTRA&"verify_record.asp?file=" & Request.Form("file") & "&line=" & Request.Form("linenum"))

'Response.Redirect ""&RankPath&"\verify_record.asp?file=" & Request.Form("file") & "&line=" & Request.Form("linenum")
'Response.Redirect ""&PathToTRA&"verify_record.asp?file=" & Request.Form("file") & "&line=" & Request.Form("linenum")
Response.Redirect "/rankings/verify_record.asp?file=" & Request.Form("file") & "&line=" & Request.Form("linenum")

'markdebug("UpdateFields Line 199")

%>




