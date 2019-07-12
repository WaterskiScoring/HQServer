<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



OpenState="rssfeed"
DisplayHeadOpenBodyAndBannerTags OpenState



	WriteRSSFeed



' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags




  SUB TY
  
  %>
<%

END SUB
  



' -------------------
  SUB WriteRSSFeed
' -------------------

Set xmlDOM = Server.CreateObject("MSXML2.DOMDocument")
xmlDOM.async = False
xmlDOM.setProperty "ServerHTTPRequest", True
xmlDOM.Load("http://www.usawaterski.org/rss.asp?Category=5")
 
'Set itemList = XMLDom.SelectNodes("news/article")
Set itemList = XMLDom.SelectNodes("rss/channel/item")


%>
		
<div class="scroll" style="background-color:white;">
<style type="text/css">
	img { width:auto; height:auto; max-width:50%; max-height:120px;" }
</style>		
	
	<%

 
For Each itemAttrib In itemList
   itemtitle =itemAttrib.SelectSingleNode("title").text
   itemdescription =itemAttrib.SelectSingleNode("description").text
   itemlink =itemAttrib.SelectSingleNode("link").text
   itempubDate =itemAttrib.SelectSingleNode("pubDate").text

		%>
		<div style="margin:20px 10px 0px 10px;">
      <div style="color:red; font-size:14pt;"><%=itemtitle%></div>
      <div style="color:black; font-size:10pt; margin:5px 0px 0px 0px;"><%=itemdescription%></div>
      <%
      tyt=1
      IF tyt=2 THEN 
      	%>
      	<div style="color:black; font-size:10pt;"><%=itemlink%></div>
      	<%
      END IF
      %>	
      <div style="color:black; font-size:8pt; margin:5px 0px 0px 0px;"><%=itempubDate%></div>
		</div>
		<%
Next
 
Set xmlDOM = Nothing
Set itemList = Nothing

%></div><%

END SUB


%>


