<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Welcome to USA Water Ski</title>
<link rel="stylesheet" type="text/css" href="/css/styles.css" />
<script language="javascript" type="text/JavaScript" src="/jscripts/scripts.js"></script>
<script language="javascript" type="text/javascript" src="/jscripts/swfobject.js"></script>
</head>

<body onload="MM_preloadImages('/images/interior/img_06_f2.jpg','/images/interior/img_08_f2.jpg','/images/interior/img_10_f2.jpg','/images/interior/img_12_f2.jpg','/images/interior/img_14_f2.jpg','/images/interior/img_16_f2.jpg','/images/interior/img_18_f2.jpg','/images/interior/img_20_f2.jpg','/images/interior/img_22_f2.jpg')">
<table cellspacing="0" class="layout">
  <tr>
    <td><img src="/sc_infodir/members/images/img_01.jpg" alt="" width="1014" height="35" /></td>
  </tr>
  <tr>
    <td><table width="1014" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td class="logo" width="241"><img src="/sc_infodir/members/images/img_02.jpg" alt="USA Water Ski" width="241" height="108" /></td>
          <td class="top"><table width="773" height="108" cellpadding="0" cellspacing="0">
              <tr>
                <td width="62"></td>
                <td align="center" width="468" class="top_ad_space"><div id="flashcontent" style="width:100%;text-align:center;"></div>
                    <!--#include virtual="/inc/banners.asp" -->
						<layer id="placeholderlayer"></layer><div id="placeholderdiv"></div></td>
                <td class="top_search"><!--#include virtual="/inc/search_form.asp" --></td>
              </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
  <!--#include virtual="/inc/icon_nav.asp" -->
  <tr>
    <td><img src="/sc_infodir/members/images/interior/img_24.jpg" alt="Having FUN Today...Building CHAMPIONS For Tomorrow" width="1014" height="18" /></td>
  </tr>
  <tr>
    <td><img src="/sc_infodir/members/images/img_26.jpg" alt="" width="1014" height="24" /></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="244" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
            <td class="sidebar"><!--#include virtual="/inc/nav.asp" --></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><img src="/sc_infodir/members/images/img_31.jpg" alt="" width="234" height="71" /></td>
          </tr>
        </table></td>
        <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="contentcontainer">
          <tr>
            <td class="content_4">
<%
If Len(Request.QueryString("insertPage")) > 0 Then
	framePage = Request.QueryString("insertPage")
Else
	framePage = "http://rankings.usawaterski.org/sample_insert.asp"
End If
%>
							<iframe src="<%= framePage %>" frameborder="0" scrolling="auto" height="600px" width="100%"></iframe>
            </td>
							</tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include virtual="/inc/footer.asp" --></td>
  </tr>
</table>
</body>
</html>




