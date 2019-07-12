<!--#include virtual="/rankings/secure-settings.asp"-->
<%

Dim sAction
Dim sTable
Dim sRecordSet
Dim UsingSQL
Dim ThisPage
Dim sID

if request("search") = "1" then
  session("SearchFilter") = "1"
  session("TourIDFilter") = request("Tour_ID")
  session("MemberIDFilter") = request("Member_ID")
  session("LastNameFilter") = request("Last_Name")
  session("FirstNameFilter") = request("First_Name")
  session("DivisionFilter") = request("Division")
  session("EventFilter") = request("Event")
end if
if request("search") = "clear" then
  session.contents.remove("SearchFilter")
  session.contents.remove("TourIDFilter")
  session.contents.remove("MemberIDFilter")
  session.contents.remove("LastNameFilter")
  session.contents.remove("FirstNameFilter")
  session.contents.remove("DivisionFilter")
  session.contents.remove("EventFilter")
end if

sAction = trim(Request("action"))
sTable = trim(Request("table"))
If sTable = "" Then sTable = RawScoresTableName

Dim currentPage, rowCount, i
currentPage = TRIM(Request("currentPage"))
if currentPage = "" then currentPage = 1

sID = trim(Request("id"))
If sID = "" then sID = 0


ThisPage = Request.ServerVariables("SCRIPT_NAME")


'---- DataTypeEnum Values ----'
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205


'-Main Page Code-----------------------------------'
if sAction = "" then sAction = "listrec"
WriteHeader
Select Case lcase(sAction)
case "listrec"
    WriteHeaders "List Records: " & sTable
    ListRecords
case "editrec"
    WriteHeaders "Edit Record: " & sID & " in " & sTable
    ShowEditor
case "saverec"
    WriteHeaders "Record Saved to: " & trim(Request("Table"))
    SaveRec
case "delrec"
    WriteHeaders "Delete Record: " & sID & " in " & sTable
    DeleteRec
end select
WriteFooter










Sub WriteHeaders(sTitle)
' Write Headers for DB Page

%>


<TABLE BORDER="0" CELLPADDING="6" CELLSPACING="0" WIDTH="100%" BGCOLOR="#C0C0C0" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0" >
<TR>
<TD ALIGN="Left"><Font Face="courier" COLOR="#000000" SIZE="4"><B><% Response.Write(sTitle) %></B></FONT></TD>
</TR>
</TABLE>
<BR>

<%
End Sub



Sub ListRecords
'  List DB Records

Response.Write("<BR>")

If sTable = RawScoresTableName Then
%> 
<center>
<h4>Search For Specific Records.</h4>

<form action="/rankings/viewdb.asp" method="post">
<input type="hidden" name="search" value="1">

<br>

<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=60%>
<TR>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Tour ID</FONT></Center></TD>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Member ID</FONT></Center></TD>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Last Name</FONT></Center></TD>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">First Name</FONT></Center></TD>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Division</FONT></Center></TD>
<TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Event</FONT></Center></TD>
</TR>

<TR>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1"><input type="text" name="Tour_ID" value="<%=Session("TourIDFilter")%>" size=9></input></FONT></Center></TD>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1"><input type="text" name="Member_ID" value="<%=Session("MemberIDFilter")%>" size=9></input></FONT></Center></TD>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1"><input type="text" name="Last_Name" value="<%=Session("LastNameFilter")%>" size=15></input></FONT></Center></TD>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1"><input type="text" name="First_Name" value="<%=Session("FirstNameFilter")%>" size=15></input></FONT></Center></TD>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1"><input type="text" name="Division" value="<%=Session("DivisionFilter")%>" size=5></input></FONT></Center></TD>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1"><input type="text" name="Event" value="<%=Session("EventFilter")%>" size=5></input></FONT></Center></TD>
</TR>
</table>
<br>
<input type="submit" value="Search"></form>
<form action="/rankings/viewdb.asp" method="post">
<input type="hidden" name="search" value="clear">
<input type="submit" value="Reset Search Filters"></form>
</center>
<%
end if
OpenCon
if session("SearchFilter") = "1" then
  sSQL = "Select * from " & RawScoresTableName & " where "
  if session("TourIDFilter") <> "" then
    sSQL = sSQL + "tourid LIKE '%" & SQLClean(Session("TourIDFilter")) & "%'"
    if session("MemberIDFilter") <> "" or session("LastNameFilter") <> "" or session("FirstNameFilter") <> "" or session("DivisionFilter") <> "" or session("EventFilter") <> "" then
      sSQL = sSQL + " and "
    end if
  end if
  if session("MemberIDFilter") <> "" then
    sSQL = sSQL + "memberid LIKE '%" & SQLClean(Session("MemberIDFilter")) & "%'"
    if session("LastNameFilter") <> "" or session("FirstNameFilter") <> "" or session("DivisionFilter") <> "" or session("EventFilter") <> "" then
      sSQL = sSQL + " and "
    end if
  end if
  if session("LastNameFilter") <> "" then
    sSQL = sSQL + "lower(lname) LIKE '%" & SQLClean(lcase(Session("LastNameFilter"))) & "%'"
    if session("FirstNameFilter") <> "" or session("DivisionFilter") <> "" or session("EventFilter") <> "" then
      sSQL = sSQL + " and "
    end if
  end if
  if session("FirstNameFilter") <> "" then
    sSQL = sSQL + "lower(fname) LIKE '%" & SQLClean(lcase(Session("FirstNameFilter"))) & "%'"
    if session("DivisionFilter") <> "" or session("EventFilter") <> "" then
      sSQL = sSQL + " and "
    end if
  end if
  if session("DivisionFilter") <> "" then
    sSQL = sSQL + "lower(div) LIKE '%" & SQLClean(lcase(Session("DivisionFilter"))) & "%'"    
    if session("EventFilter") <> "" then
      sSQL = sSQL + " and "
    end if
  end if
  if session("EventFilter") <> "" then
    sSQL = sSQL + "lower(event) LIKE '%" & SQLClean(lcase(Session("EventFilter"))) & "%'"
  end if
  ChoosePagesSQL sSQL,currentPage, 40
else
  ChoosePagesSQL "Select * from " & sTable,currentPage, 40
end if
rowCount = 0
%>
<BR>
<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" WIDTH="100%" BGCOLOR="#FFFFFF" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0">
<TR>
<TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1">Rec #</FONT></TD>
<TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1">Delete</FONT></TD>
<TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1">Edit</FONT></TD>
<%
For i = 1 to rs.fields.count - 1
%>
<TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><% Response.Write(Rs.Fields(i).name) %></FONT></TD>
<%
next
%>
</TR>
<%
'add table data here with paging'

do while not rs.eof
if rowCount = rs.PageSize then exit DO
%>
<TR>
<TD ALIGN="Left" vAlign="top"><FONT COlOR="#000000" FACE="courier" SIZE="1">&nbsp;<% =rs.fields(0).Value %></FONT></TD>
<TD ALIGN="Left" vAlign="top"><FONT COlOR="#000000" FACE="courier" SIZE="1">&nbsp;<% WriteLink "?action=delrec&id=" & rs.fields(0).Value & "&table=" & stable,"Delete","" %></FONT></TD>
<TD ALIGN="Left" vAlign="top"><FONT COlOR="#000000" FACE="courier" SIZE="1">&nbsp;<% WriteLink "?action=editrec&id=" & rs.fields(0).Value & "&table=" & stable,"Edit","" %></FONT></TD>
<%
For i = 1 to rs.fields.count - 1
%>
<TD ALIGN="Left" vAlign="top"><FONT COlOR="#000000" FACE="courier" SIZE="1">&nbsp;<%
Response.Write(Rs.Fields(i).value) 
%></FONT></TD>
<%
next
%>
</TR>

<% 
rowCount = rowCount + 1
rs.movenext
loop

%>
</TABLE>
<br><br>
<%
DoCount currentPage
CloseCon
End Sub


Sub ShowEditor
WriteLink "?action=listrec&table=" & stable,"Return To Database","<BR>"
OpenCon
ChooseSQL "Select * from " & sTable & " where(id=" & sID & ")"
%>
<FORM METHOD="POST" ACTION="<% Response.Write(ThisPage) %>?action=saverec&table=<% Response.Write(sTable) %>">
<TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0">

<%

For i = 0 to rs.fields.count - 1
%>
<TR>
<TD ALIGN="Left" vAlign="Top"><Font Face="courier" COLOR="#000000" SIZE="2"><B><% Response.Write(Rs.Fields(i).name) %></B></FONT></TD>
<TD ALIGN="Left" vAlign="Top" bgcolor="#C0C0C0"><Font COLOR="#000000" SIZE="2"><% WriteType i %></FONT></TD>
</TR>
<%
next
%>
</TABLE><BR>
<TABLE BORDER="0" CELLPADDING="3" CELLSPACING="0">
<TR>
<TD ALIGN="Left"><input type="submit" value="Save"></TD>
<TD ALIGN="Left"><input type="reset" value="Reset"></TD>
</TR>
</TABLE>

</FORM>
<%
CloseCon
End Sub


Sub SaveRec
'Save the record to the table'
OpenCon
ChooseSQL "Select * from " & sTable & " where(ID=" & sID & ")"
if sID = 0 then
    rs.addnew
else
    rs.movefirst
end if

For i = 1 to rs.fields.count - 1
    'set the field value'
      select case rs.fields(i).type
        case adBigInt
          rs.fields(i).value = csng(Request.Form(rs.fields(i).name))
        case adBoolean 
          if trim(Request.Form(rs.fields(i).name) = "") then
            rs.fields(i).value = False
          else
            rs.fields(i).value = True
          end if
        case adCurrency
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = ccur(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adDate,adDBDate,adDBTime,adDBTimeStamp
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = cdate(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adDecimal
          if Request.Form(rs.fields(i).name) <> "" then
             rs.fields(i).value = cdec(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adDouble
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = cdbl(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adInteger
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = cint(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adSingle
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = csng(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case else
          rs.fields(i).value = Request.Form(rs.fields(i).name)
     end select
next
rs.UpdateBatch 3
CloseCon
%>
Your record has been saved.<BR>
<%
WriteLink "?action=listrec&table=" & sTable,"Click here to continue.","<BR>"
End Sub

Sub DeleteRec
WriteLink "?action=listrec&table=" & sTable,"Return To Database","<BR><BR>"
OpenCon
ChooseSQL "Select * from " & sTable & " where(ID=" & sID & ")"
if lcase(Request("confirm")) = "yes" then
    'delete the record'
    WriteLog(date() &"  "& time() &"  "& rs.fields(0).value & " " & rs.fields(1).value & " " & rs.fields(2).value & " has been deleted.")
    if isrecordsetempty = false then
        rs.movefirst
        rs.delete
        rs.UpdateBatch 3
    end if
    CloseCon
    %>
    The record was deleted.<BR>
    <%
end if
if lcase(Request("confirm")) = "" then
%>  <br><br>
    Type the word "YES" if you are sure you wish to delete this record.
    <br><br>
    <form action="/rankings/viewdb.asp" method="post"> 
    <input type="hidden" name="action" value="delrec">
    <input type="hidden" name="id" value="<%=sid%>">
    <input type="hidden" name="table" value="<%=stable%>">
    <input type="text" name="confirm" size="5">
    <input type="submit" value="Confirm Deletion?">
    </form>
<%
    WriteLink "?action=listrec&table=" & stable,"No - do not delete the record","<BR><BR>"
end if
if lcase(Request("confirm")) <> "yes" and lcase(Request("confirm")) <> "" then
     %>  <br><br>
         The record was NOT deleted.
         <br><br>
     <%
end if

End Sub


Sub GetCheckValue(i)
    if lcase(sAction) = "editrec" then
        if rs.fields(i).value = 0 then
            Response.Write("")
        else
            Response.Write("checked")
        end if
    else
        Response.Write("")
    end if
End Sub



Sub GetFieldValue(i)
    if lcase(sAction) = "editrec" then
        Response.Write(rs.fields(i).value)
    else
        Response.Write("")
    end if
End Sub

Sub WriteType(I)
Select Case Rs.Fields(i).type
case 3 'primary key / auto number ?'
    if lcase(Rs.Fields(i).name) = "id" then
       if sid = 0 then
          %>
          <input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number (new)
          <%
       else
          %>
          <input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number (<% Response.Write(sID) %>)
          <%
       end if
    else
        %>
        <input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>">
        <%
    end if
case 20 'primary key / auto number ?'
    if lcase(Rs.Fields(i).name) = "id" then
       if sid = 0 then
          %>
          <input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number (new)
          <%
       else
          %>
          <input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number (<% Response.Write(sID) %>)
          <%
       end if
    else
        %>
        <input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>">
        <%
    end if
case 11 'boolean'
    %>
    <INPUT TYPE="checkbox" NAME="<% Response.Write(Rs.Fields(i).name) %>" VALUE="1" <% GetCheckValue i %>>
    <%
case 203 'memo'
    %>
    <TEXTAREA NAME="<% Response.Write(Rs.Fields(i).name) %>" ROWS="20" COLS="56"><% GetFieldValue i %></TEXTAREA>
    <%
case else 'not handled by this function'
    %>
    <input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>">
    <%
End Select

End Sub

Sub ChooseSQL(sSQL)
set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 3
End Sub


Sub WriteHeader
%>
<HTML>
<HEAD><TITLE>TRA Database Tool</TITLE>
</HEAD>

<BODY BGCOLOR="#FFFFFE" Text="#0A0D0A" LINK="#375AE2" VLINK="#36566D" ALINK="#3E85BB">
<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:#375AE2}  A:visited {text-decoration: none; color:#375AE2}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:#3E85BB; }-->
</style>
<%
End Sub

Sub WriteFooter
%>
<hr><h4><a href="/">Return to Main Menu</a></h4><hr>
</BODY>
</HTML>
<%
End Sub


Function IsRecordSetEmpty
if rs.bof = true and rs.eof = true then
    IsRecordSetEmpty = true
else
    IsRecordSetEmpty = false
end if
end Function

Sub ChoosePagesSQL(sSQL,sStart, sSize)
  set rs=Server.CreateObject("ADODB.recordset")
  sqlstmt = sSQL
  rs.CursorType = 3
  rs.PageSize = cint(sSize)
  rs.open sqlstmt, SConnectionToTRATable
  if isrecordsetempty = false then
    rs.AbsolutePage = cINT(sStart)
  end if
End Sub

Sub WriteLink(sParms,sDisplay,sBreak)
%><A HREF="<% Response.Write(ThisPage & sParms) %>"><% Response.Write(sDisplay) %></A><%
End Sub


Sub DoCount(currentPage) 
h = rs.PageCount

if h > 21 then
  if currentpage - 10 > 1 then
    Response.Write("... ")
  end if
  for i = 10 to 1 step -1
    if currentpage - i > 0 then
      Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage - i  & "&action=" & sAction & "&table=" & sTable & chr(34) & ">" & currentpage - i & "</a>")
    end if
  next
  Response.Write(" " & currentpage & " ")
  for i = 1 to 10
   if currentpage + i <= h then
      Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage + i  & "&action=" & sAction & "&table=" & sTable & chr(34) & ">" & currentpage + i & "</a>")
    end if
  next
  if currentpage + 10 <= h then
    Response.Write(" ...")
  end if
else
  for i = 1 to h
    Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  i  & "&action=" & sAction & "&table=" & sTable & chr(34) & ">" & i & "</a>")
  next
end if

if h = 0 then h = 1
Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")
end sub

Function GetFieldTypeName(I)
select case i
case 0
GetFieldTypeName = "Empty"
case 16
GetFieldTypeName = "TinyInt"
case 2
GetFieldTypeName = "SmallInt"
case 3
GetFieldTypeName = "Integer"
case 20
GetFieldTypeName = "BigInt"
case 17
GetFieldTypeName = "UnsignedTinyInt"
case 18
GetFieldTypeName = "UnsignedSmallInt"
case 19
GetFieldTypeName = "UnsignedInt"
case 21
GetFieldTypeName = "UnsignedBigInt"
case 4
GetFieldTypeName = "Single"
case 5
GetFieldTypeName = "Double"
case 6
GetFieldTypeName = "Currency"
case 14
GetFieldTypeName = "Decimal"
case 131
GetFieldTypeName = "Numeric"
case 11
GetFieldTypeName = "Boolean"
case 10
GetFieldTypeName = "Error"
case 132
GetFieldTypeName = "UserDefined"
case 12
GetFieldTypeName = "Variant"
case 9
GetFieldTypeName = "IDispatch"
case 13
GetFieldTypeName = "IUnknown"
case 72
GetFieldTypeName = "GUID"
case 7
GetFieldTypeName = "Date"
case 133
GetFieldTypeName = "DBDate"
case 134
GetFieldTypeName = "DBTime"
case 135
GetFieldTypeName = "DBTimeStamp"
case 8
GetFieldTypeName = "BSTR"
case 129
GetFieldTypeName = "Char"
case 200
GetFieldTypeName = "VarChar"
case 201
GetFieldTypeName = "LongVarChar"
case 130
GetFieldTypeName = "WChar"
case 202
GetFieldTypeName = "VarWChar"
case 203
GetFieldTypeName = "LongVarWChar"
case 128
GetFieldTypeName = "Binary"
case 204
GetFieldTypeName = "VarBinary"
case 205
GetFieldTypeName = "LongVarBinary"
End Select
End Function

Function GetFieldTypeCode(sTXT,sLen)
'I am not overly familar with this stuff'
'you may have to edit these values'
select case sTXT
case "Empty"
GetFieldTypeCode = "Empty"
case "TinyInt"
GetFieldTypeCode = "TinyInt"
case "SmallInt"
GetFieldTypeCode = "SmallInt"
case "Integer"
GetFieldTypeCode = "Integer"
case "BigInt"
GetFieldTypeCode = "BigInt"
case "UnsignedTinyInt"
GetFieldTypeCode = "UnsignedTinyInt"
case "UnsignedSmallInt"
GetFieldTypeCode = "UnsignedSmallInt"
case "UnsignedInt"
GetFieldTypeCode = "UnsignedInt"
case "UnsignedBigInt"
GetFieldTypeCode = "UnsignedBigInt"
case "Single"
GetFieldTypeCode = "Single"
case "Double"
GetFieldTypeCode = "Double"
case "Currency"
GetFieldTypeCode = "Currency"
case "Decimal"
GetFieldTypeCode = "Decimal"
case "Numeric"
GetFieldTypeCode = "Numeric"
case "Boolean"
GetFieldTypeCode = "Boolean"
case "Error"
GetFieldTypeCode = "Error"
case "UserDefined"
GetFieldTypeCode = "UserDefined"
case "Variant"
GetFieldTypeCode = "Variant"
case "IDispatch"
GetFieldTypeCode = "IDispatch"
case "IUnknown"
GetFieldTypeCode = "IUnknown"
case "GUID"
GetFieldTypeCode = "GUID"
case "Date"
GetFieldTypeCode = "Date"
case "DBDate"
GetFieldTypeCode = "DBDate"
case "DBTime"
GetFieldTypeCode = "DBTime"
case "DBTimeStamp"
GetFieldTypeCode = "DBTimeStamp"
case "BSTR"
GetFieldTypeCode = "BSTR(" & sLen & ")"
case "Char"
GetFieldTypeCode = "Char(" & sLen & ")"
case "VarChar"
GetFieldTypeCode = "VarChar(" & sLen & ")"
case "LongVarChar"
GetFieldTypeCode = "LongVarChar(" & sLen & ")"
case "WChar"
GetFieldTypeCode = "WChar(" & sLen & ")"
case "VarWChar"
GetFieldTypeCode = "VarWChar(" & sLen & ")"
case "LongVarWChar"
GetFieldTypeCode = "LongVarWChar(" & sLen & ")"
case "Binary"
GetFieldTypeCode = "Binary(" & sLen & ")"
case "VarBinary"
GetFieldTypeCode = "VarBinary(" & sLen & ")"
case "LongVarBinary"
GetFieldTypeCode = "LongVarBinary"
case else
GetFieldTypeCode = "IUnknown"
End Select
End Function

%>



