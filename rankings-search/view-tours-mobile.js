/*
 * Javacript file for USA Waterski Tournament Listing Mobile
 * @author Mark Crone
 * @version 1.0.0
 * @date May 23, 2015
 * @copyright (c) USA Waterski
 */


		(function(document,navigator,standalone) {
			// prevents links from apps from oppening in mobile safari
			// this javascript must be the first script in your <head>

			//alert('document = ' + document);
			if ((standalone in navigator) && navigator[standalone]) {
				var curnode, location=document.location, stop=/^(a|html)$/i;
				document.addEventListener('click', function(e) {
					curnode=e.target;
					while (!(stop).test(curnode.nodeName)) {
						curnode=curnode.parentNode;
					}
					// Conditions to do this only on links to your own app
					// if you want all links, use if('href' in curnode) instead.
					if(
					
						'href' in curnode && // is a link
						(chref=curnode.href).replace(location.href,'').indexOf('#') && // is not an anchor
						(	!(/^[a-z\+\.\-]+:/i).test(chref) ||                       // either does not have a proper scheme (relative links)
							chref.indexOf(location.protocol+'//'+location.host)===0 ) // or is in the same protocol and domain
					) {
						e.preventDefault();
						location.href = curnode.href;
					}
				},false);
			}
			//alert('curnode = ' + curnode);
			//alert('href = ' + href);

		})(document,window.navigator,'standalone');
	

function hideURLbar() {
	if (window.location.hash.indexOf('#') == -1) {
		window.scrollTo(0, 1);
	}
}

if (navigator.userAgent.indexOf('iPhone') != -1 || navigator.userAgent.indexOf('Android') != -1) {
    addEventListener("load", function() {
            setTimeout(hideURLbar, 0);
    }, false);
}



function set_localStorage(sMemberID, sFirstName, sLastName) {
			
			// --- Sets local storage value for sMemberID 
			// localStorage.removeItem("sMemberID");
			//alert('SET LOCAL');
			// alert('SET LOCAL - sMemberID: ' +  sMemberID + ' <br> sFirstName: ' + sFirstName);
			
			localStorage.setItem("sMemberID",sMemberID);
			localStorage.setItem("sFirstName",sFirstName);
			localStorage.setItem("sLastName",sLastName);
		}

function set_localWatchStorage(sWatchMemberID, sWatchFirstName, sWatchLastName, ActionButtonValue) {
			
			// --- Sets local storage value for sMemberID 
			if (ActionButtonValue = 'Add') {
					var CurrIDs = localStorage.getItem("sWatchMemberIDs"));
					var CurrPos = CurrIDs.indexOf(sWatchMemberID);
					if (CurrPos != -1) {
							localStorage.setItem("sWatchMemberIDs",CurrIDs + ',' + sWatchMemberID);
						}

					//localStorage.setItem("sFirstName",sFirstName);
					//localStorage.setItem("sLastName",sLastName);
				}		
			if (ActionButtonValue = 'Delete') {
					//localStorage.setItem("sMemberID",sMemberID);
					//localStorage.setItem("sFirstName",sFirstName);
					//localStorage.setItem("sLastName",sLastName);
				}		
		}
		
		
		


function clear_localStorage() {
			localStorage.removeItem("sMemberID");
			alert('IN CLEAR');
			alert('IN CLEAR localStorage.getItem(sMemberID) = ' + localStorage.getItem("sMemberID"));
		}



function OnOpen(whichdisplay) {
		if (whichdisplay == 'norankingsfound') {		
			document.getElementById('norankingsfound').style.display = 'inline'; 		
		} 		
		else if (whichdisplay == 'nodivisionselected') {		
			document.getElementById('nodivisionselected').style.display = 'none'; 		
			document.getElementById('rankingssearchsettings').style.display = 'inline'; 
		} 		
		else if (whichdisplay == 'notournamentselected') {		
			document.getElementById('tournamentlisting').style.display = 'none'; 		
			document.getElementById('toursearchsettings').style.display = 'inline';
		} 		
		else if (whichdisplay == 'TourInfo') {		
			document.getElementById('tournamentlisting').style.display = 'none'; 		
			document.getElementById('toursearchsettings').style.display = 'inline';
		} 		
	}


function SetUserNav(whichdisplay) {
		if (whichdisplay == 'changeuser') {		
			document.getElementById('ChangeUserScreen').style.display = 'inline-block'; 		
			document.getElementById('CurrentUserScreen').style.display = 'none'; 		
		} 
		if (whichdisplay == 'returntocurrentuser') {		
			document.getElementById('ChangeUserScreen').style.display = 'none'; 		
			document.getElementById('CurrentUserScreen').style.display = 'inline-block'; 
		} 
		if (whichdisplay == 'load') {		
				if (localStorage.getItem("sMemberID") !== null) {
					document.getElementById('SetChangeUser').value = 'Change User'
					}
				else {
					document.getElementById('SetChangeUser').value = 'Set User'
					}
			} 
		if (whichdisplay == 'modifywatchers') {
			document.getElementById('CurrentUserScreen').style.display = 'none'; 		
			document.getElementById('WatcherEntryScreen').style.display = 'inline-block'; 
			}
	}			



function UpdateMemberID_IntoHidden_FromStart() {
		//alert('Line 32 START -  Local: ' + localStorage.getItem("sMemberID"));
		//alert('Line 32 START -  Local: ' + localStorage.getItem("sFirstName") + " " + localStorage.getItem("sLastName")); 
			if ( localStorage.getItem("sMemberID") != null && localStorage.getItem("sMemberID") != '999' ) {
					document.getElementById('sMemberID_Hidden_InRankingsSettings').value = localStorage.getItem("sMemberID");
					document.getElementById('sName_InRankingsSettings').value = localStorage.getItem("sFirstName") + " " + localStorage.getItem("sLastName");

				}
			else {
				//document.getElementById('sMemberID_InRankingsSettings').value = '999'
				document.getElementById('sName_InRankingsSettings').value = 'No Member Specified'
			}	
		//alert('Line 33 BEFORE');
		
	}

function DisplaySearchFilters(whichbutton) {
			if (whichbutton == 'searchfilters') {
				document.getElementById('rankingssearchsettings').style.display = 'inline'; 		
				document.getElementById('rankingslisting').style.display = 'none'; 		
				}
			else if (whichbutton == 'displayrankings') {
				document.getElementById('rankingssearchsettings').style.display = 'none'; 		
				// document.getElementById('rankingslisting').style.display = 'inline'; 		
				}
			else if (whichbutton == 'tournamentsearchfilters') {
				document.getElementById('tournamentlisting').style.display = 'none'; 
				document.getElementById('toursearchsettings').style.display = 'inline';				
			}				
			else if (whichbutton == 'norecordserror') {
				document.getElementById('notournamentselectederror').style.display = 'none'; 
				document.getElementById('toursearchsettings').style.display = 'inline';				
			}				
	} 


function MainMenuOptions(whichoption) {
			if (whichoption == 'twitter') {
				document.getElementById('twitterfeedscreen').style.display = 'inline'; 		
				document.getElementById('mainmenuscreen').style.display = 'none'; 		
				document.getElementById('AddIconNowInstruction').style.display = 'none';
				}
			else if (whichoption == 'faq') {
				document.getElementById('FAQMenuScreen').style.display = 'inline'; 		
				document.getElementById('mainmenuscreen').style.display = 'none'; 		
				document.getElementById('AddIconNowInstruction').style.display = 'none';
				}
			if (whichoption == 'iphone') {
				document.getElementById('LaunchfromiPhone').style.display = 'inline'; 		
				document.getElementById('SavingSearchSettings').style.display = 'none'; 		
				document.getElementById('AddIconNowInstruction').style.display = 'none';
				}
			if (whichoption == 'rulebook') {
				document.getElementById('DisplayRulebook').style.display = 'inline'; 		
				document.getElementById('mainmenuscreen').style.display = 'none'; 		
				document.getElementById('AddIconNowInstruction').style.display = 'none';
				}
			if (whichoption == 'HomeFromiPhone') {
				document.getElementById('iPhoneAddIcon').style.display = 'none'; 		
				document.getElementById('mainmenuscreen').style.display = 'inline'; 		
				document.getElementById('AddIconNowInstruction').style.display = 'inline';				
				if( ("standalone" in navigator) && window.navigator.standalone) {
				// alert('Condition TRUE - standalone');
					document.getElementById('AddIconNowInstruction').value = 'Home Screen icon is already installed on this phone';	
					}
				else {
					//	alert('Condition FALSE');
					document.getElementById('AddIconNowInstruction').value = 'iPhone users press the Bookmark icon in the center of the tray below to begin Add to Home Screen process';	
					}
				}
			if (whichoption == 'ResourcesFromHome') {
				document.getElementById('mainmenuscreen').style.display = 'none';
				document.getElementById('ResoucesMenuScreen').style.display = 'inline'; 		
 				}
		}

function LaunchPageOptions(whichoption) {
			if (whichoption == 'iphone') {
				document.getElementById('mainlaunch').style.display = 'none'; 		
				document.getElementById('iPhoneAddIcon').style.display = 'inline'; 		
				}
			if (whichoption == 'mainlaunch') {
				document.getElementById('iPhoneAddIcon').style.display = 'none'; 		
				document.getElementById('mainlaunch').style.display = 'inline'; 	
				document.getElementById('AddIconNowInstruction').value = 'iPhone users press the Bookmark icon in the center of the tray below to begin Add to Home Screen process';	
				}
		}

function FAQOptions(whichoption) {
			if (whichoption == 'iphone') {
				document.getElementById('iPhoneAddIcon').style.display = 'inline'; 		
				document.getElementById('FAQMenuScreen').style.display = 'none'; 		
				}
			if (whichoption == 'faqfromiPhone') {
				document.getElementById('FAQMenuScreen').style.display = 'inline'; 		
				document.getElementById('iPhoneAddIcon').style.display = 'none'; 		
				}
			if (whichoption == 'savesearch') {
				document.getElementById('SavingSearchSettings').style.display = 'inline'; 		
				document.getElementById('FAQMenuScreen').style.display = 'none'; 		
				}
			if (whichoption == 'faqfromSaveSearch') {
				document.getElementById('FAQMenuScreen').style.display = 'inline'; 		
				document.getElementById('SavingSearchSettings').style.display = 'none'; 		
				}
		}




function TeamCreateNav(whichoption) {
			if (whichoption == 'toteamnameentryfromtypeselect') {
				document.getElementById('teamnameentryscreen').style.display = 'inline-block'; 		
				document.getElementById('teamtypeselectionscreen').style.display = 'none'; 		
				}
			if (whichoption == 'totypeselectfromteamnameentry') {
				document.getElementById('teamtypeselectionscreen').style.display = 'inline-block'; 		
				document.getElementById('teamnameentryscreen').style.display = 'none'; 		
				}
		}


function TeamConfirmInviteNav(whichoption,sInviteMemberID,sInviteName,sInviteCityState,sInviteEmail,sTeam_ID) {
			if (whichoption == 'ToConfirmSendInviteFromMyTeamList') {
				//alert('IN ToConfirmSendInviteFromMyTeamList - sInviteEmail=' + sInviteEmail); 
				document.getElementById('ConfirmInviteNewMemberScreen').style.display = 'inline-block'; 		
				document.getElementById('TeamManageOptionsScreen').style.display = 'none'; 
				document.getElementById('sInviteMemberID').value = sInviteMemberID; 		
				document.getElementById('sInviteName').value = sInviteName;
				document.getElementById('sInviteCityState').value = sInviteCityState;
				document.getElementById('sInviteEmail').value = sInviteEmail;
				document.getElementById('sTeam_ID').value = sTeam_ID;
				
				//alert('IN ToConfirmSendInviteFromMyTeamList - END'); 
				}
			if (whichoption == 'ToMyTeamListFromConfirmSendInvite') {
				//alert('IN ToMyTeamListFromConfirmSendInvite'); 
				document.getElementById('ConfirmInviteNewMemberScreen').style.display = 'none'; 		
				document.getElementById('TeamManageOptionsScreen').style.display = 'inline-block'; 
				}

		}



function SelectFromTournamentMenu(whichbutton)
	{
		if (whichbutton == '1') {
			document.getElementById('tourmenubuttons').style.display = "none"; 		
			document.getElementById('tourlisting').style.display = "none"; 		
			document.getElementById('toursearchsettings').style.display = "inline"; 		
			document.getElementById('function_not_available').style.display = "none"; 		
			}
		else if (whichbutton == '2') {
			document.getElementById('tourmenubuttons').style.display = "none"; 		
			document.getElementById('tourlisting').style.display = "none"; 		
			document.getElementById('function_not_available').style.display = "inline"; 		
			}					
		else if (whichbutton == '3') {
			document.getElementById('tourmenubuttons').style.display = "none"; 		
			document.getElementById('rankingslisting').style.display = "none"; 		
			document.getElementById('rankingssearchsettings').style.display = "inline"; 		
			}					
		else if (whichbutton == '4') {
			document.getElementById('tourmenubuttons').style.display = "none"; 		
			document.getElementById('tourlisting').style.display = "none"; 		
			document.getElementById('function_not_available').style.display = "inline"; 		
			}								
	}




function ReturnToListing(whichdisplay) {
		if (whichdisplay == 'tournamentlisting') {		
			document.getElementById('tournamentlisting').style.display = 'inline'; 		
			document.getElementById('toursearchsettings').style.display = 'none';
		} 		
		else if (whichdisplay == 'tournamentlistingfromtourdetails') {		
			document.getElementById('tournamentlisting').style.display = 'inline'; 		
			document.getElementById('tourdetails').style.display = 'none';
		} 		
		else if (whichdisplay == 'searchfiltersfromrecalcerror') {		
			document.getElementById('rankingssearchsettings').style.display = 'inline'; 		
			document.getElementById('displayrankingunderwaymessage').style.display = 'none';
		} 		
		else if (whichdisplay == 'rankingslisting') {		
			document.getElementById('rankingssearchsettings').style.display = 'none';
			//document.getElementById('tourmenubuttons').style.display = 'inline'; 		
			document.getElementById('rankingslisting').style.display = 'inline'; 		
		} 		
		//else {
		//	document.getElementById('function_not_available').style.display = "none"; 			
		//}	
	}

//ReturnToListing('norankingsfound')

function StoreRankingsSettingsToLocalVar() {
		var RankingListType = document.getElementById('RankingListType').value;
		
		if (RankingListType == 'National' || RankingListType == 'Junior' ) {		
			localStorage.setItem("RankingsSkiYearID_National",document.getElementById('RankingsSkiYearIDSelected').value);
			localStorage.setItem("RankingsStateRegionSelected_National",document.getElementById('RankingsStateRegionSelected').value);
			localStorage.setItem("RankingsDivSelected_National",document.getElementById('RankingsDivSelected').value);
			localStorage.setItem("RankingsEventSelected_National",document.getElementById('RankingsEventSelected').value);
			localStorage.setItem("RankingsInclude_International_National",document.getElementById('Include_International').value);
			alert('Your National Rankings Settings Have Been Saved'); 		
		}
		else if (RankingListType == 'NCWSA') {		
			alert('Your NCWSA Rankings Settings Have Been Saved'); 		
			localStorage.setItem("RankingsSkiYearID_NCWSA",document.getElementById('RankingsSkiYearIDSelected').value);
			localStorage.setItem("RankingsStateRegionSelected_NCWSA",document.getElementById('RankingsStateRegionSelected').value);
			localStorage.setItem("RankingsDivSelected_NCWSA",document.getElementById('RankingsDivSelected').value);
			localStorage.setItem("RankingsEventSelected_NCWSA",document.getElementById('RankingsEventSelected').value);
			localStorage.setItem("RankingsInclude_International_NCWSA",document.getElementById('Include_International').value);		
		}
	}

function UpdateRankingsSettingsFromLocal() {
			var RankingListType = document.getElementById('RankingListType').value;
			//alert("RankingListType =" & RankingListType); 		
			if ( (RankingListType == 'National' || RankingListType == 'Junior' ) && localStorage.getItem("RankingsSkiYearID_National") != null && localStorage.getItem("RankingsSkiYearID_National") != '999' ) {
			//alert('National HERE'); 		
					document.getElementById('RankingsSkiYearIDSelected').value = localStorage.getItem("RankingsSkiYearID_National");
					document.getElementById('RankingsStateRegionSelected').value = localStorage.getItem("RankingsStateRegionSelected_National");
					document.getElementById('RankingsDivSelected').value = localStorage.getItem("RankingsDivSelected_National");
					document.getElementById('RankingsEventSelected').value = localStorage.getItem("RankingsEventSelected_National");
					document.getElementById('Include_International').value = localStorage.getItem("RankingsInclude_International_National");															
				}
			else if ( RankingListType == 'NCWSA' && localStorage.getItem("RankingsSkiYearID_NCWSA") != null && localStorage.getItem("RankingsSkiYearID_NCWSA") != '999' ) {
					document.getElementById('RankingsSkiYearIDSelected').value = localStorage.getItem("RankingsSkiYearID_NCWSA");
					document.getElementById('RankingsStateRegionSelected').value = localStorage.getItem("RankingsStateRegionSelected_NCWSA");
					document.getElementById('RankingsDivSelected').value = localStorage.getItem("RankingsDivSelected_NCWSA");
					document.getElementById('RankingsEventSelected').value = localStorage.getItem("RankingsEventSelected_NCWSA");
					document.getElementById('Include_International').value = localStorage.getItem("RankingsInclude_International_NCWSA");															
				}
		}


function StoreTournamentSettingsToLocalVar() {
			localStorage.setItem("Tournament_SportsGroup",document.getElementById('sSportsGroup').value);
			localStorage.setItem("Tournament_Region",document.getElementById('Region').value);
			localStorage.setItem("Tournament_TourLevel",document.getElementById('sTourLevel').value);
			localStorage.setItem("Tournament_State",document.getElementById('State').value);
			localStorage.setItem("Tournament_Class",document.getElementById('sClass').value);
			//localStorage.setItem("Tournament_StartMonth",document.getElementById('StartMonth').value);
			//localStorage.setItem("Tournament_EndMonth",document.getElementById('EndMonth').value);
			alert('Your Tournament Search Settings Have Been Saved'); 		
	}

function UpdateTournamentSettingsFromLocal() {
			document.getElementById('sSportsGroup').value = localStorage.getItem("Tournament_SportsGroup");
			document.getElementById('Region').value = localStorage.getItem("Tournament_Region");
			document.getElementById('sTourLevel').value = localStorage.getItem("Tournament_TourLevel");
			document.getElementById('State').value = localStorage.getItem("Tournament_State");
			document.getElementById('sClass').value = localStorage.getItem("Tournament_Class");

			//document.getElementById('StartMonth').value = localStorage.getItem("Tournament_StartMonth");
			//document.getElementById('EndMonth').value = localStorage.getItem("RankingsInclude_EndMonth");															
		}

function UpdateNOPSField() {
		var Div = document.getElementById('DivSelected').value;
		//alert('DivDropdown = ' + Div);
		
		var DivArray = document.getElementById('DivArray').value;
		var SlalomNOPSArray = document.getElementById('SlalomNOPSArray').value;
		var TrickNOPSArray = document.getElementById('TrickNOPSArray').value;
		var JumpNOPSArray = document.getElementById('JumpNOPSArray').value;
		var SlalomExpArray = document.getElementById('SlalomExpArray').value;
		var TrickExpArray = document.getElementById('TrickExpArray').value;
		var JumpExpArray = document.getElementById('JumpExpArray').value;		
		var OverPtsBySArray = document.getElementById('OverPtsBySArray').value;		

		var ThisDivArray = DivArray.split(',');
		var ThisSlalomNOPSArray = SlalomNOPSArray.split(',');
		var ThisTrickNOPSArray = TrickNOPSArray.split(',');
		var ThisJumpNOPSArray = JumpNOPSArray.split(',');
		var ThisSlalomExpArray = SlalomExpArray.split(',');
		var ThisTrickExpArray = TrickExpArray.split(',');
		var ThisJumpExpArray = JumpExpArray.split(',');		
		var ThisOverPtsBySArray = OverPtsBySArray.split(',');		
		
		var ThisIndexNo = ThisDivArray.indexOf(Div);
		var ThisDiv = ThisDivArray[ThisIndexNo];

		//alert('Index No = ' + ThisIndexNo + ' -- ThisDiv = ' +ThisDiv);

		var SlalomRecord = ThisSlalomNOPSArray[ThisIndexNo];
		var TrickRecord = ThisTrickNOPSArray[ThisIndexNo];
		var JumpRecord = ThisJumpNOPSArray[ThisIndexNo];
		var SlalomExp = ThisSlalomExpArray[ThisIndexNo];
		var TrickExp = ThisTrickExpArray[ThisIndexNo];
		var JumpExp = ThisJumpExpArray[ThisIndexNo];		
		var SlalomPts = ThisOverPtsBySArray[ThisIndexNo];

		//alert('SlalomRecord = ' + SlalomRecord + ' -- TrickRecord = ' + TrickRecord + ' -- JumpRecord = ' + JumpRecord)
		
		//alert('SlalomExp = ' + SlalomExp + ' --  TrickExp = ' + TrickExp + ' -- JumpExp = ' + JumpExp + ' --  SlalomPts = ' + SlalomPts)

		var RawScore_S = document.getElementById('RawScore_S').value;
		var RawScore_T = document.getElementById('RawScore_T').value;		
		var RawScore_J = document.getElementById('RawScore_J').value;

		//alert('RawScore_S = ' + RawScore_S + ' --  RawScore_T = ' + RawScore_T + ' -- RawScore_J = ' + RawScore_J)		
				
		var NOPS_S = 0;
		var NOPS_T = 0;
		var NOPS_J = 0;
		
		// --- NOPS calc SLALOM ---
		//NOPS_S = (0.5+10 * IF(RawScore_S<6,(RawScore_S*SlalomPts),(6*SlalomPts)+((1500-(6*SlalomPts))*((RawScore_S-6)/(SlalomRecord-6))^SlalomExp)))/10
		//INT(0.5+10*IF(Slalom_Score<6,(Slalom_Score*SLM_Pts),(6*SLM_Pts)+((1500-(6*SLM_Pts))*((Slalom_Score-6)/(Slm_Recd-6))^SLM_Exp)))/10
		//INT(0.5+10*((6*SLM_Pts)+((1500-(6*SLM_Pts))*((Slalom_Score-6)/(Slm_Recd-6))^SLM_Exp)))/10
		if ( RawScore_S < 6 )	{
			NOPS_S = parseInt(0.5 + 10 * (RawScore_S * SlalomPts))/10;
			}
		else {
			//NOPS_S = parseInt(0.5 + 10 * ((6 * SlalomPts) + ((1500 - (6 * SlalomPts)) *     ((RawScore_S - 6)/(SlalomRecord - 6))^SlalomExp   )))/10
			NOPS_S = parseInt(0.5 + 10 * ((6 * SlalomPts) + ((1500 - (6 * SlalomPts)) *    Math.pow( ((RawScore_S - 6)/(SlalomRecord - 6)),SlalomExp)   )))/10;
			}
		// --- NOPS calc TRICKS ---
		NOPS_T = parseInt(0.5 + 15000 * Math.pow(RawScore_T/TrickRecord,TrickExp))/10;

		// --- NOPS calc JUMP ---
		// NOPS_J = (0.5+10*IF(RawScore_J<(0.15*JumpRecord),0,1500*(((RawScore_J-(0.15*JumpRecord))/(JumpRecord-(0.15*JumpRecord)))^JumpExp)))/10
		if (RawScore_J < (0.15 * JumpRecord)) {
			NOPS_J = parseInt(0.5 + 10 * 0) / 10;

			}
		else	{
			//NOPS_J = parseInt(0.5 + 10 * 1500 * (  ((RawScore_J-(0.15*JumpRecord))/(JumpRecord-(0.15*JumpRecord))) ^JumpExp     )) / 10
			NOPS_J = parseInt(0.5 + 10 * 1500 * ( Math.pow(((RawScore_J-(0.15*JumpRecord))/(JumpRecord-(0.15*JumpRecord))),JumpExp) )) / 10;
			}

		var NOPS_O = 0;
		NOPS_O = NOPS_S + NOPS_T + NOPS_J; 
		
		//alert('NOPS_S = ' + NOPS_S + ' -- NOPS_T = ' + NOPS_T + ' -- NOPS_J = ' + NOPS_J);
			
		document.getElementById('Record_S').value = SlalomRecord;
		document.getElementById('Record_T').value = TrickRecord;
		document.getElementById('Record_J').value = JumpRecord;
		
		document.getElementById('NOPS_S').value = NOPS_S;
		document.getElementById('NOPS_T').value = NOPS_T;
		document.getElementById('NOPS_J').value = NOPS_J;		
		document.getElementById('NOPS_O').value = NOPS_O;
		
	}


	