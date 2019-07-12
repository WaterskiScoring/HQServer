<%	





' ----------------------------  END OF MAIN PROGRAM -----------------------

' ----------------------------------
   SUB DefineTourSessionVar (sTourID)
' ----------------------------------

	' --- dim these variables at page level scope --- 
	' include variables calculated by GetHCodes and all fields produced by fn_TschedulRegFieldsXTournAppID()
'	dim bMixedClass, bMixedClassS, bMixedClassT, bMixedClassJ, HClass, HClassLow, HClassStd, EventCount, GrtEventCount 
'	dim sTournAppID, sSptsGrpID, sTRegion, sTDateS, sTDateE, sTLateDate, sTLateFee, sTLFPerDay
'	dim	sTName, sTSite, sTCity, sTState, sTSponsor, sTSponsorID, sTYear, sTSkiYr, sTSTATUS, sCanceled0
'	dim TEventSlalom, sTEventJump, sTEventTrick, sTEventFun, sTEventNSL, sTEventNBL, sTEventNWL, sTEventCustom
'	dim sTDescription, sFDescription, sCDescription, sWDescription, sKDescription, sT5Star				
'	dim	sTOpenClosed, sTPandC, sTPandCPulls, sMaxPulls, sTDvOffered, sRestrictDv
'	dim sGrassRoots, sTEventFHF, sTEventFKB, sTEventFDA, sTEventF3ev, sTEventFB, sTEventFW				
'	dim sTHSClassF, sTHSClassN, sTHSClassI, sTHSClassC, sTHSClassE, sTHSClassL, sTHSClassR, sTHSClassCASH, sTHSClassX
'	dim sTHJClassF, sTHJClassN, sTHJClassI, sTHJClassC, sTHJClassE, sTHJClassL, sTHJClassR, sTHJClassCASH, sTHJClassX
'	dim sTHTClassF, sTHTClassN, sTHTClassI, sTHTClassC, sTHTClassE, sTHTClassL, sTHTClassR, sTHTClassCASH, sTHTClassX
'	dim sTRoundsS, sTRoundsT, sTRoundsJ, sTRoundsF
'	dim sTDirName, sTDirAddress, sTDirCity, sTDirState, sTDirZip, sTDirEmail, sTDirFAX, sTDirPhoneAm, sTDirPhonePm, sTsEmail
'	dim sTRegistrarName, sTRegistrarPhone, sTRegistrarEmail, sTRegistrarAddr, sTRegistrarCity, sTRegistrarState, sTRegistrarZip, sTRegistrarFax
'	dim sTEventClin, sClinNumParticipants, sClinLevel, sJDClin, sADClin
'	dim sTEventWake, sTRoundsWakeBd, sBoatWake, sCableWake, sTEventWSkate, sTRoundsWSkate, sBoatSkate, sCableSkate
'	dim sTEventWSurf, sTRoundsWSurf
'	dim sKEventFlip, sKRoundsFlip, sKFlipClassT, sKFlipClassQ, sKEventFree, sKRoundsFree, sKFreeClassT, sKFreeClassQ
'	dim sKSlalomClassT, sKSlalomClassQ, sKTrickClassT, sKTrickClassQ



	'Get the recordset and populate the variables - then close the recordset
	dim Conn, objCnn, objRS, SQL
	
	' --- Define from calling program ---
	sTAID = left(sTourID,6)

			'Calculate the SptsGrpID
			Select case lcase(mid(sTournAppID, 3, 1))
				case c, e, s, m, w
					sTSptsGrpID = "AWS"
				case x
					sTSptsGrpID = "USW"
				case b
					sTSptsGrpID = "ABC"
				case u
					sTSptsGrpID = "NCW"  
				case k
					sTSptsGrpID = "AKA"
			end select
			if len(sTAID) <> 6 then
				Response.Write("TAID= " & sTAID & "<br>Invalid TAID - use back button and try again.")
				Response.End
			end if


			SQL = "SELECT * FROM fn_TschedulRegFieldsXTournAppID('" & sTAID & "')"
			'Use a connection string that works from your location.
		'	sConn = Application("WSAConnStr") ' "Provider=SQLOLEDB;SERVER=jaguar.epolk.net;Database=Sanctions;uid=Sanctions_P;pwd=43qe9ho6"
			sConn = "Provider=SQLOLEDB;SERVER=jaguar.epolk.net;Database=Sanctions;uid=Sanctions_Admin;pwd=qej8h7w34w"
			Set Conn = Server.CreateObject("ADODB.Connection")
			Set objRS = server.CreateObject("ADODB.Recordset")
			Conn.Open sConn

			Set objRS = Conn.Execute (SQL)
			If objRS.EOF And objRS.BOF = True Then  'No match found 
				objRS.close
				Conn.close
				set objRS = nothing
				set Conn = nothing
				Response.Write("Did not record.")
				Response.end
			else 'found match - load event and class field values into page scope variables
'markdebug("Found Tour")
	'You should validate data coming in for each variable - at least specify type
	Session("sTSptsGrpID") = objRS("SptsGrpID")


	Session("sTRegion") = objRS("TRegion")
	Session("sTDateS") = objRS("TDateS")
	Session("sTDateE") = objRS("TDateE")
	Session("sTLateDate") = objRS("TLateDate")
	Session("sTLateFee") = objRS("TLateFee")
	Session("sTLFPerDay") = objRS("TLFPerDay")
	Session("sTName") = objRS("TName")
	Session("sTSite") = objRS("TSite")
	Session("sTCity") = objRS("TCity")
	Session("sTState") = objRS("TState")
	Session("sTSponsor") = objRS("TSponsor")
	Session("sTSponsorID") = objRS("TSponsor")
	Session("sTYear") = objRS("TYear")
	Session("sTSkiYr") = objRS("TSkiYr")
	Session("sTSTATUS") = objRS("TSTATUS")
	Session("sCanceled0") = objRS("Canceled0")
	Session("sTEventSlalom") = objRS("TEventSlalom")
	Session("sTEventJump") = objRS("TEventJump")
	Session("sTEventTrick") = objRS("TEventTrick")
	Session("sTEventFun") = objRS("TEventFun")
	Session("sTEventNSL") = objRS("TEventNSL")
	Session("sTEventNBL") = objRS("TEventNBL")
	Session("sTEventNWL") = objRS("TEventNWL")
	Session("sTEventCustom") = objRS("TEventCustom")
	Session("sTEventClin") = objRS("TEventClin")
	Session("sTDescription") = objRS("TDescription")
	Session("sFDescription") = objRS("FDescription")
	Session("sCDescription") = objRS("CDescription")
	Session("sWDescription") = objRS("WDescription")
	Session("sKDescription") = objRS("KDescription")
	Session("sT5Star") = objRS("T5Star")
	Session("sTOpenClosed") = objRS("TOpenClosed")
	Session("sTPandC") = objRS("TPandC")
	Session("sTPandCPulls") = objRS("TPandCPulls")
	Session("sMaxPulls") = objRS("MaxPulls")
	Session("sTDvOffered") = objRS("TDvOffered")
	Session("sRestrictDv") = objRS("RestrictDv")
	Session("sGrassRoots") = objRS("GrassRoots")
	Session("sTEventFHF") = objRS("TEventFHF")
	Session("sTEventFKB") = objRS("TEventFKB")
	Session("sTEventFDA") = objRS("TEventFDA")
	Session("sTEventF3ev") = objRS("TEventF3ev")
	Session("sTEventFB") = objRS("TEventFB")
	Session("sTEventFW") = objRS("TEventFW")
	Session("sTHSClassF") = objRS("THSClassF")
	Session("sTHSClassN") = objRS("THSClassN")
	Session("sTHSClassI") = objRS("THSClassI")
	Session("sTHSClassC") = objRS("THSClassC")
	Session("sTHSClassE") = objRS("THSClassE")
	Session("sTHSClassL") = objRS("THSClassL")
	Session("sTHSClassR") = objRS("THSClassR")
	Session("sTHSClassCASH") = objRS("THSClassCASH")
	Session("sTHSClassX") = objRS("THSClassX")
	Session("sTHJClassF") = objRS("THJClassF")
	Session("sTHJClassN") = objRS("THJClassN")
	Session("sTHJClassI") = objRS("THJClassI")
	Session("sTHJClassC") = objRS("THJClassC")
	Session("sTHJClassE") = objRS("THJClassE")
	Session("sTHJClassL") = objRS("THJClassL")
	Session("sTHJClassR") = objRS("THJClassR")
	Session("sTHJClassCASH") = objRS("THJClassCASH")
	Session("sTHJClassX") = objRS("THJClassX")
	Session("sTHTClassF") = objRS("THTClassF")
	Session("sTHTClassN") = objRS("THTClassN")
	Session("sTHTClassI") = objRS("THTClassI")
	Session("sTHTClassC") = objRS("THTClassC")
	Session("sTHTClassE") = objRS("THTClassE")
	Session("sTHTClassL") = objRS("THTClassL")
	Session("sTHTClassR") = objRS("THTClassR")
	Session("sTHTClassCASH") = objRS("THTClassCASH")
	Session("sTHTClassX") = objRS("THTClassX")
	Session("sTRoundsS") = objRS("TRoundsS")
	Session("sTRoundsT") = objRS("TRoundsT")
	Session("sTRoundsJ") = objRS("TRoundsJ")
	Session("sTRoundsF") = objRS("TRoundsF")
	Session("sTDirName") = objRS("TDirName")
	Session("sTDirAddress") = objRS("TDirAddress")
	Session("sTDirCity") = objRS("TDirCity")
	Session("sTDirState") = objRS("TDirState")
	Session("sTDirZip") = objRS("TDirZip")
	Session("sTDirEmail") = objRS("TDirEmail")
	Session("sTDirFAX") = objRS("TDirFAX")
	Session("sTDirPhoneAm") = objRS("TDirPhoneAm")
	Session("sTDirPhonePm") = objRS("TDirPhonePm")
	Session("sTsEmail") = objRS("TsEmail")
	Session("sTRegistrarName") = objRS("TRegistrarName")
	Session("sTRegistrarPhone") = objRS("TRegistrarPhone")
	Session("sTRegistrarEmail") = objRS("TRegistrarEmail")
	Session("sTRegistrarAddr") = objRS("TRegistrarAddr")
	Session("sTRegistrarCity") = objRS("TRegistrarCity")
	Session("sTRegistrarState") = objRS("TRegistrarState")
	Session("sTRegistrarZip") = objRS("TRegistrarZip")
	Session("sTRegistrarFax") = objRS("TRegistrarFax")
	Session("sClinNumParticipants") = objRS("ClinNumParticipants")
	Session("sClinLevel") = objRS("ClinLevel")
	Session("sJDClin") = objRS("JDClin")
	Session("sADClin") = objRS("ADClin")
	Session("sTEventWake") = objRS("TEventWake")
	Session("sTRoundsWakeBd") = objRS("TRoundsWakeBd")
	Session("sBoatWake") = objRS("BoatWake")
	Session("sCableWake") = objRS("CableWake")
	Session("sTEventWSkate") = objRS("TEventWSkate")
	Session("sTRoundsWSkate") = objRS("TEventWSkate")
	Session("sBoatSkate") = objRS("BoatSkate")
	Session("sCableSkate") = objRS("CableSkate")
	Session("sTEventWSurf") = objRS("TEventWSurf")
	Session("sTRoundsWSurf") = objRS("TRoundsWSurf")
	Session("sKEventFlip") = objRS("KEventFlip")
	Session("sKRoundsFlip") = objRS("KRoundsFlip")
	Session("sKFlipClassT") = objRS("KFlipClassT")
	Session("sKFlipClassQ") = objRS("KFlipClassQ")
	Session("sKEventFree") = objRS("KEventFree")
	Session("sKRoundsFree") = objRS("KRoundsFree")
	Session("sKFreeClassT") = objRS("KFreeClassT")
	Session("sKFreeClassQ") = objRS("KFreeClassQ")
	Session("sKSlalomClassT") = objRS("KSlalomClassT")
	Session("sKSlalomClassQ") = objRS("KSlalomClassQ")
	Session("sKTrickClassT") = objRS("KTrickClassT")
	Session("sKTrickClassQ") = objRS("KTrickClassQ")
	'Call the function that calculates MixedClass etc.
	' GetHCodes(objRS)  
	'destroy the recordset and close the connection		


	objRS.close
	Conn.close
	set objRS = nothing
	set Conn = nothing		
				
		end if

END SUB

%>




