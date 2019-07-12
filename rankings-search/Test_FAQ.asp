<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



OpenState="rssfeed"
DisplayHeadOpenBodyAndBannerTags OpenState



	WriteFAQItem



' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags




  SUB TY
  
  %>
<iframe src="news/faq_rankings_m.htm" name="targetframe" allowTransparency="true" frameborder="0" ></iframe>
<%

END SUB
  






' -------------------
  SUB WriteFAQItem
' -------------------



%>
		
<div class="scroll" style="background-color:white; height:400px; font-family: Arial, Helvetica, sans-serif; font-size:12pt; width:100%;">
   	<!--#include virtual="/rankings/news/faq_rankings_m.htm"-->
   	
</div><%

END SUB


%>


