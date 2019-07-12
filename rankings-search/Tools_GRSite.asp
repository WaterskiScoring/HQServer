<%

'--- Subroutines for GR Site


Dim ThisFileName
Dim GRTableColor1, GRTableColor2, GRTableColor3

ThisFileName="Tools_GRSite.asp"
GRTableColor1="#000000"
GRTableColor2="#303030"
GRTableColor3="#505050"
GRTableColor4="#DCDCDC"



' ------------------
  SUB DefineGRCSS
' ------------------

%>
<style type="text/css">
/*
/* this style applies to the GRtable1 table */
table.GRtable1 {padding:0px; border:1px solid <%=GRTableColor3%>; border-collapse:collapse;}
table.GRtable1 th {padding:1px; border:1px solid <%=GRTableColor1%>; border-style:solid; background-color:<%=GRTableColor3%>; vertical-align:bottom;} 
table.GRtable1 td {padding:3px; border:1px solid <%=GRTableColor2%>; border-style:solid; background-color:<%=GRTableColor4%>; vertical-align:middle; white-space:nowrap;} 

/* this style applies to the GRtable2 table */
table.GRtable2 {padding:0px; border:3px solid <%=GRTableColor3%>; border-collapse:collapse;}
table.GRTable2 th {padding:1px; border:0px solid <%=GRTableColor1%>; border-style:solid; background-color:<%=GRTableColor2%>; vertical-align:bottom;} 
table.GRTable2 td {padding-left:6px; padding-right:6px; padding-top:0px; padding-bottom:0px; border:0px solid <%=GRTableColor2%>; border-style:solid; vertical-align:middle; white-space:nowrap;} 




p
{
font-family: Verdana, Arial, Helvetica, sans-serif;
font-size:7pt;
color:<%=textcolor1%>;
} 

h1
{
font-family: Verdana, Arial, Helvetica, sans-serif;
font-size:24pt;
font-weight:bold;
color:white;
} 

h2
{
font-family: Verdana, Arial, Helvetica, sans-serif;
font-size:20pt;
font-weight:bold;
color:white;
} 

h3
{
font-family: Verdana, Arial, Helvetica, sans-serif;
font-size:16pt;
font-style: italic;
font-weight:bold;
color:red;
} 

h4
{
font-family: Verdana, Arial, Helvetica, sans-serif;
font-size:12pt;
font-style: italic;
font-weight:bold;
color:<%=textcolor1%>;
} 

h5
{
font-family: Verdana, Arial, Helvetica, sans-serif;
font-size:8pt;
color:<%=textcolor2%>;
} 

</style><%

END SUB



' ------------------
  SUB WriteGRHeader
' ------------------

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>GrassRoots Home</title>
<meta name="generator" content="WYSIWYG Web Builder - http://www.wysiwygwebbuilder.com">
<link rel="shortcut icon" href="usa.png">
<style type="text/css">
div#container
{
   width: 1024px;
   position: relative;
   margin-top: 0px;
   margin-left: auto;
   margin-right: auto;
   text-align: left;
}
body
{
   text-align: center;
   margin: 0;
}
</style>
<script type="text/javascript">
<!--
function PreloadImages()
{
   var imageObj = new Image();
   var images = new Array();
   images[0]="/GrassRootsSeries/images/img0001.gif";
   images[1]="/GrassRootsSeries/images/img0001_over.gif";
   images[2]="/GrassRootsSeries/images/img0002.gif";
   images[3]="/GrassRootsSeries/images/img0002_over.gif";
   images[4]="/GrassRootsSeries/images/img0003.gif";
   images[5]="/GrassRootsSeries/images/img0003_over.gif";
   images[6]="/GrassRootsSeries/images/img0004.gif";
   images[7]="/GrassRootsSeries/images/img0004_over.gif";
   images[8]="/GrassRootsSeries/images/img0005.gif";
   images[9]="/GrassRootsSeries/images/img0005_over.gif";
   images[10]="/GrassRootsSeries/images/a1.jpg";
   images[11]="/GrassRootsSeries/images/a2.jpg";
   images[12]="/GrassRootsSeries/images/a3.jpg";
   images[13]="/GrassRootsSeries/images/a4.jpg";
   images[14]="/GrassRootsSeries/images/a5.jpg";
   images[15]="/GrassRootsSeries/images/a6.jpg";
   images[16]="/GrassRootsSeries/images/a7.jpg";
   images[17]="/GrassRootsSeries/images/a8.jpg";
   images[18]="/GrassRootsSeries/images/a9.jpg";
   for (var i=0; i<=18; i++)
   {
      imageObj.src = images[i];
   }
}
// -->
</script>
<script language="JavaScript" type="text/javascript">
<!--
function FadeImage(id, duration)
{
   var millisec = Math.round(duration / 100);
   var timer = 0;
   for(i = 0; i <= 100; i++)
   {
      setTimeout("SetOpacity('" + id + "'," + i + ")",(timer * millisec));
      timer++;
   }
}
function SetOpacity(id, opacity)
{
   var element = document.getElementById(id).style;
   element.opacity = (opacity / 100);
   element.MozOpacity = (opacity / 100);
   element.KhtmlOpacity = (opacity / 100);
   element.filter = "alpha(opacity=" + opacity + ")";
}
//-->
</script>
<style type="text/css">
a
{
   color: #FFFFFF;
}
a:visited
{
   color: #7F7F7F;
}
a:active
{
   color: #99FF33;
}
a:hover
{
   color: #99FF33;
}
a.Menu_Links:link
{
   color: #FFFFFF;
   font-weight: bold;
   text-decoration: none;
}
a.Menu_Links:visited
{
   color: #FFFFFF;
   font-weight: bold;
   text-decoration: none;
}
a.Menu_Links:active
{
   color: #99FF33;
   font-weight: bold;
   text-decoration: none;
}
a.Menu_Links:hover
{
   color: #99FF33;
   font-weight: bold;
   text-decoration: none;
}
</style>
</head>

<% MainPageWidth=1050  
%>


<body bgcolor="#000000" text="#FFFFFF" >


<TABLE border=0 height=510px width=<%=MainPageWidth%>px align=center background="/GrassRootsSeries/images/BK.jpg" >
  <TR>
    <TD colspan=3 align=right valign=bottom>	
	  <div id="wb_TextMenu1" style="width:695px;height:31px;z-index:1;" align="center">
		<font style="font-size:24px;" color="#000000" face="Arial">
		[<a href="/GrassRootsSeries/index.html" class="Menu_Links">Home</a>]&nbsp;[<a href="/GrassrootsSeries/about.html" class="Menu_Links">About Us</a>]&nbsp;[<a href="http://www.usawaterski.org/rankings/View-TournamentsHQ.asp" class="Menu_Links" target="_blank">Events</a>]&nbsp;[<a href="http://www.usawaterski.org/rankings/view-grscores.asp" class="Menu_Links" target="_blank">Scores</a>]&nbsp;[<a href="http://www.usawaterski.org/rankings/view-grranking.asp" class="Menu_Links" target="_blank">Rankings</a>]&nbsp;[<a href="./contact.html" class="Menu_Links">Contact Us</a>]
		</font>
	  </div>

     </TD>
   </TR>	
</TABLE>


<TABLE border=0 bgcolor="black" align=center width=<%=MainPageWidth%>px height=400px>
   <TR>
     <TD width=15% align=center valign=top>

	<div id="wb_SlideShow1" style="width:235px;height:111px;z-index:8;overflow:hidden" align="left">


	<% ' --- This is the changing image of sponsors %>	
	<script language="JavaScript" type="text/javascript">
	<!--
	   var SlideShow1_Index = -1;
	   var SlideShow1_Images = new Array();
	   SlideShow1_Images[0] = ["/GrassRootsSeries/images/a1.jpg","http://www.skidim.com/","_blank"];
	   SlideShow1_Images[1] = ["/GrassRootsSeries/images/a2.jpg","http://www.obrien.com/","_blank"];
	   SlideShow1_Images[2] = ["/GrassRootsSeries/images/a3.jpg","http://www.mastercraft.com/","_blank"];
	   SlideShow1_Images[3] = ["/GrassRootsSeries/images/a4.jpg","http://www.nautiques.com/","_blank"];
	   SlideShow1_Images[4] = ["/GrassRootsSeries/images/a5.jpg","http://www.malibuboats.com/","_blank"];
	   SlideShow1_Images[5] = ["/GrassRootsSeries/images/a6.jpg","http://www.ojprops.com/","_blank"];
	   SlideShow1_Images[6] = ["/GrassRootsSeries/images/a7.jpg","http://www.hosports.com/","_blank"];
	   SlideShow1_Images[7] = ["/GrassRootsSeries/images/a8.jpg","https://www.quotemyboat.com/quote/","_blank"];
	   SlideShow1_Images[8] = ["/GrassRootsSeries/images/a9.jpg","http://www.indmar.com//","_blank"];

	   function SlideShow1ShowNext()
	   {
	      SlideShow1_Index = SlideShow1_Index + 1;
	      if (SlideShow1_Index > 8)
	         SlideShow1_Index = 0;
	      document.getElementById('SlideShow1_Fade').src = document.getElementById('SlideShow1').src;
	      SetOpacity('SlideShow1', 0);
	      eval("document.SlideShow1.src = SlideShow1_Images[" + SlideShow1_Index + "][0]");
	      setTimeout("SlideShow1ShowNext();", 3000);
	      FadeImage('SlideShow1', 1500);
	   }

	   function onSlideShow1Click()
	   {
	      if (SlideShow1_Images[SlideShow1_Index][2] == "")
	      {
	         targetwin = "_self";
	      }
	      else
	      {
	         targetwin = SlideShow1_Images[SlideShow1_Index][2];
	      }
	      eval("window.open(url = SlideShow1_Images[" + SlideShow1_Index + "][1],'" + targetwin +"');");
	   }
	// -->
	</script>

	<a href="#" onClick="onSlideShow1Click();return false;">
		<img src="/GrassRootsSeries/images/a1.jpg" id="SlideShow1_Fade" border="0" align="top" alt="" width="235" height="111" name="SlideShow1_Fade">
		<img src="/GrassRootsSeries/images/a1.jpg" id="SlideShow1" border="0" align="top" alt="" width="235" height="111" name="SlideShow1">
	</a>


	<script language="JavaScript" type="text/javascript">
	<!--
	  SlideShow1ShowNext();
	// -->
	</script>


	</div>

	<!-- Search -->
	<div id="Html1" style="width:235px;height:48px;z-index:12">
	
	  <div id="searchxmedia">
	    <script type="text/javascript">

		// Google Internal Site Search script- By JavaScriptKit.com (http://www.javascriptkit.com)
		// For this and over 400+ free scripts, visit JavaScript Kit- http://www.javascriptkit.com/
		// This notice must stay intact for use

		//Enter domain of site to search.
		var domainroot="http://www.usawaterski.org"

		function Gsitesearch(curobj){
		curobj.q.value="site:"+domainroot+" "+curobj.qfront.value
		}

	     </script>

	    <form action="http://www.google.com/search" method="get" onSubmit="Gsitesearch(this)">
		<input name="q" type="hidden" />
		<input name="qfront" type="text" style="width: 165px" /> <input type="submit" value="Search" />
	    </form>

	  </div>
	</div>


     <br>	
	<div id="wb_Image5" style="width:260px;height:174px;z-index:10;" align="left">
		<a href="./gallery.html"><img src="/GrassRootsSeries/images/gallery.png" id="Image5" alt="" align="top" border="0" style="width:260px;height:174px;"></a>
	</div>
     <br>	
	<div id="wb_Image4" style="width:259px;height:175px;z-index:9;" align="left">
		<a href="./competitors.html"><img src="/GrassRootsSeries/images/competitors.png" id="Image4" alt="" align="top" border="0" style="width:259px;height:175px;"></a>
	</div>
     <br>	
	<div id="wb_Image6" style="width:260px;height:176px;z-index:11;" align="left">
		<a href="./events.html"><img src="/GrassRootsSeries/images/events.png" id="Image6" alt="" align="top" border="0" style="width:260px;height:176px;"></a>
	</div>
     <br>	
	<div id="wb_Image1" style="width:201px;height:35px;z-index:5;" align="left">
		<a href="http://www.myspace.com/462502997" target="_blank"><img src="/GrassRootsSeries/images/myspace.png" id="Image1" alt="" align="top" border="0" style="width:201px;height:35px;"></a>
	</div>
     <br>	
	<div id="wb_Image2" style="width:109px;height:42px;z-index:6;" align="left">
		<a href="http://www.facebook.com/pages/USA-Water-Ski-GrassRoots-Series/76320551428?ref=s" target="_blank"><img src="/GrassRootsSeries/images/facebook.png" id="Image2" alt="" align="top" border="0" style="width:109px;height:42px;"></a>
	</div>
     <br>	
	<div id="wb_Image3" style="width:105px;height:52px;z-index:7;" align="left">
		<a href="http://www.usawaterski.org/" target="_blank"><img src="/GrassRootsSeries/images/usa.png" id="Image3" alt="" align="top" border="0" style="width:105px;height:52px;"></a>
	</div>



    </TD>


    <TD colspan=2 align=center>	<% ' --- Cell from overall page design to contain table containing area for content


	' --- Content table --- %>
	<TABLE Align=center style="background-color:black;" width=100% height=100%>
	  <TR>
	    <TD align=center valign=top>
		<%  ' --- Ends with the opening tag of the table area for content ---



END SUB





' --------------------
   SUB WriteGRFooter
' --------------------

	' --- Starts with the CLOSING tags of the table holding content --- %>

	    </TD>
	  </TR>
	</TABLE>

	<% ' --- Closing tags of cell from main page table --- %>	
    </TD>
  </TR>	
</TABLE>



</body>
</html>
<!-- www.000webhost.com Analytics Code -->
<script type="text/javascript" src="http://analytics.hosting24.com/count.php"></script>
<noscript><a href="http://www.hosting24.com/"><img src="http://analytics.hosting24.com/count.php" alt="web hosting" /></a></noscript>
<!-- End Of Analytics Code -->
<%

END SUB




' -----------------
  SUB OldContent
' -----------------

%>

	<div id="wb_Text1" style="position:absolute;left:723px;top:572px;width:176px;height:38px;z-index:2;" align="center">
		<font style="font-size:32px" color="#99FF33" face="Century Gothic"><b>WELCOME</b></font>
	</div>

	<div id="wb_Text3" style="position:absolute;left:273px;top:631px;width:117px;height:25px;z-index:3;" align="right">
		<font style="font-size:21px" color="#99FF33" face="Tahoma">Latest News</font>
	</div>

	<div id="wb_Text2" style="width:538px;height:160px;z-index:4;" align="left">
	   <font style="font-size:13px" color="#C0C0C0" face="Tahoma">
		Welcome to the new USA Water Ski GrassRoots Series Web site. This site will serve as a one-stop shop for everything related to the GrassRoots Series, including the latest news, events, photos, and information on how to host a GrassRoots event and register to participate in an event.<br>
		<br>
		USA Water Ski developed the GrassRoots Series with the ultimate goal of growing organized towed water sports in the United States. It is a great way to introduce new people to water sports and to let them see how much fun they can have on the water. Our challenge to all of the tournament organizers as well as the existing individual USA Water Ski members is to get friends and family involved in towed water sports by hosting a GrassRoots Series event.</font>
	</div>
<%

END SUB




' -----------------------------------------------
   SUB LoadGRTournamentList (GROnly)
' -----------------------------------------------


' ----  Define LEAGUE drop down from LeagueTableName ----
IF GROnly="yes" THEN
	set rsTour=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT TourID, ' *' AS Grass TName"
	sSQL = sSQL + " FROM "&RawGRScoresTableName
	sSQL = sSQL + " LEFT JOIN"
	sSQL = sSQL + " 	(SELECT TournAppID, TName FROM "&SanctionTableName&") AS ST"
	sSQL = sSQL + " ON ST.TournAppID=LEFT(TourID,6)"
	sSQL = sSQL + " GROUP BY TourID, TName"
	sSQL = sSQL + " ORDER BY TName"
ELSE
	' ----  Define LEAGUE drop down from LeagueTableName ----
	set rsTour=Server.CreateObject("ADODB.recordset")

	sSQL = "SELECT TourID, TName, Grass FROM"

	sSQL = sSQL + " (SELECT TourID, Class, ' **' AS Grass FROM "&RawGRScoresTableName&" WHERE Class IN ('F','N')"
	sSQL = sSQL + " UNION "
	sSQL = sSQL + " SELECT TourID, Class, ' ' AS Grass FROM "&RawScoresTableName&" WHERE Class IN ('F','N')) AS RS"

	sSQL = sSQL + " LEFT JOIN"
	sSQL = sSQL + " 	(SELECT TournAppID, TName, TDateE FROM "&SanctionTableName&") AS ST"
	sSQL = sSQL + " ON ST.TournAppID=LEFT(TourID,6)"

	sSQL = sSQL + "	LEFT JOIN "
	sSQL = sSQL + "		(SELECT SkiYearID, EndDate, BeginDate FROM "&SkiYearTableName&") AS SY "
	sSQL = sSQL + "	ON SY.SkiYearID='"&SkiYearSelected&"'"

	sSQL = sSQL + "	WHERE SY.EndDate>=ST.TDateE AND SY.BeginDate<=ST.TDateE "	

	sSQL = sSQL + " GROUP BY TourID, TName, Class, Grass"
	sSQL = sSQL + " ORDER BY TName"
END IF

'response.write(sSQL)
'response.end

rsTour.open sSQL, SConnectionToTRATable


%>
<select name="TourSelected" <%=LeagueStatus%> onchange=submit()><%

IF TRIM(TourSelected) = "" THEN
	response.write("<option value =""None"" selected>Select Tournament</option><br>")
ELSE
	response.write("<option value =""None"">Select Tournament</option><br>")
END IF

IF NOT rsTour.eof THEN 
  	rsTour.movefirst
  	DO WHILE NOT rsTour.eof
		IF TRIM(rsTour("TourID")) = TRIM(TourSelected) THEN
			response.write("<option value =""" & rsTour("TourID") &""" selected>"&rsTour("TName")&rsTour("TourID")&rsTour("Grass")&"</option><br>")
    		ELSE
			response.write("<option value =""" & rsTour("TourID") &""">"&rsTour("TName")&rsTour("TourID")&rsTour("Grass")&"</option><br>")
		END IF	
		rsTour.moveNEXT
	LOOP
ELSE
	response.write("<option value =""NA"" selected>None Available</option>")
END IF  %>

</select><%

rsTour.close

END SUB



%>


