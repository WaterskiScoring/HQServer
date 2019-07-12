<%









Dim rs
Dim SConnectionToTRATable

Dim sMemberID, sFirst, sLast, sAddress1, sAddress2, sCity, sState, sZip
Dim sWorkPhone, sCellPhone, sEmail 
Dim StdTableWidth

Dim sColor1, sColor2, sColor3
Dim TextColor1, TextColor2, TextColor3, TextColor4, TextColor5
Dim TableColor1, BackgroundColor1

Dim fontsize1, fontsize2, fontsize3, fontsize4
Dim USStatesList

Dim MembersTableName
Dim HeaderImage1

StdTableWidth=900


' --- Table definitions ---
MembersTableName =  "usawsrank.Eago_Members"

HeaderImage1="images\eago\Eago_Header.jpg"


USStatesList = ",AL,AK,AR,AZ,CA,CO,CT,DE,FL,GA,HI,ID,IA,IL,IN,KS,KY,LA,MA,MD,ME,MI,MN,MO,MS,MT,NC,ND,NE,NH,NJ,NM,NV,NY,OH,OK,OR,PA,RI,SC,SD,TN,TX,UT,VT,VA,WA,WI,WV,WY"

font1="Verdana, Arial, Helvetica, sans-serif"
font2="arial"
fontsize1="0"
fontsize2="1"
fontsize3="2"
fontsize4="3"

' --- Text formatting
TextColor1="#000000"
TextColor2="#0000CD"
TextColor3="Red"
TextColor4="Green"
TextColor5="FFFFFF"

EagoColor1="#203f5e"
EagoColor2="#0f77da"
EagoColor3="#2F4F4F"
EagoColor2="#EEE8AA"				' Pale Goldenrod
EagoColor2="#D2B48C"				' Tan

' --- Used for formatting tables & headers ---
HeaderTextColor1="#8B4513"

TableColor1="#FFF8DC"
TableColor1="#F5FFFA"				' Mint Cream
TableColor1="#FFFAF0"				' Floral White

BackgroundColor1="#F5DEB3"	' Wheat
BackgroundColor1="#EEE8AA"	' Pale Goldenrod
BackgroundColor1="#FFFFE0"	' Light Yellow
BackgroundColor1="#FAF0E6"	' Linen

' --- Used for Formatting columns ---
scolor01 = "#FFFFFF"  ' Orignal  White (default)		
scolor02 = "#FFEBCD"  ' Bright Green  
scolor03 = "#FFFACD"  ' Lemon Chiffon 
scolor04 = "#F5DEB3"  ' Wheat
scolor05 = "#CC99FF"  ' Plum
scolor06 = "#CCCCFF"  ' Lt Steel Blue 
scolor07 = "#FFFF66"  ' Yellow 
scolor08 = "#CCFFCC"  ' Lt Green 
scolor09 = "#FFCCCC"  '   
scolor10 = "#DDA0DD"  ' Plum







' -------------------------------
' --- For SQL Server Connections 
' -------------------------------

HQUser = "trastand22"
HQPass = "ski33ret"
TRADBName = "00025"
MemberDBName = "USAWaterski"
SanctionDBName = "Sanctions"

sConnectionToTRATable = "Provider=SQLOLEDB;Data Source=jaguar.epolk.net;User ID=" & HQuser & ";Password=" & HQpass & ";Initial Catalog=cobra00025"






' *********************
    SUB OpenCon
' *********************
  Set Con = Server.CreateObject("ADODB.Connection")
  Con.ConnectionTimeout = 2000
  Con.CommandTimeout = 2000
  Con.open(SConnectionToTRATable)

END SUB






%>