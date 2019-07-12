
/*										*/
/*	Create Full View: USAWaterski.dbo.MembersLive.		*/
/*	Fully equivalent to USAWaterski.dbo.Members table,	*/
/*	except does not include PersonIDWithCheckDigit.		*/
/*	Hence must join by Person ID, using rightmost 8		*/
/*	positions of MemberID field from Rankings tables.	*/
/*										*/

USE USAWATERSKI
;

ALTER VIEW MembersLive as

SELECT PT.[Person ID] as PersonID, PT.[Name Prefix] as NamePrefix,
	 PT.[First Name] as FirstName, PT.[Middle] as MiddleName,
	 PT.[Last Name] as LastName, PT.[Name Suffix] as NameSuffix,
	 PT.SSN, PT.[Company Name] as CompanyName,
	 Substring(PT.Website,1,100) as Website, PT.Email, PT.MailPref,
	 PT.[Birth Date] as BirthDate, PT.Sex, D1.[Division Code] as DivisionCode1,
	 D2.[Division Code] as DivisionCode2, PT.[Federation Code] as FederationCode,
	 MT.MemberTypeID, PA.Phone, left(PA.Extension,4) as Extension,
	 PA.Fax, PA.[Business Phone] as BusinessPhone,
	 left(PA.[Business Extension],4) as BusinessExtension,
	 PA.[Mobile Phone] as MobilePhone, PA.Address1, PA.Address2,
	 PA.City, PA.State, PA.Zip, PA.[Country ID] as CountryID,
	 left(MH.[Membership Type Code],10) as MembershipTypeCode,
	 MH.EffectiveFrom, MH.EffectiveTo,
	 Case when PT.DoNotEMail=1 then '1' else '0' end as DoNotEMail,
	 Coalesce(TS.[Region Code], '6') as Region,
	 PT.[Member Since] as MemberSince, PT.[Date Updated] as DateUpdated,
	 Case when PT.DoNotCall=1 then '1' else '0' end as DoNotCall,
	 Left(MT.[Membership Type Description],10) as MembershipType,
	 Case when PT.Deceased=1 then '1' else '0' end as Deceased,
	 CWS.WaiverStatusID

FROM	Waterski.dbo.tblPeople PT, Waterski.dbo.[Membership History] MH,

	(Select MH.[Person ID] as PersonID, cast(substring(Max(convert(char(8),MH.EffectiveTo,112)
			+right(convert(char(8),10000000+MH.MembershipHistoryID),7)),9,7) as integer) as MemHistID
	 From Waterski.dbo.[Membership History] MH 
	 Join Waterski.dbo.[Current Waiver Status by Membership History ID] CWS
	   on CWS.MembershipHistoryID = MH.MembershipHistoryID
	 group by [Person ID]) MHL,

	 Waterski.dbo.[Current Waiver Status by Membership History ID] CWS,

	 Waterski.dbo.tblMembershipTypeCodes MT, Waterski.dbo.tblDivisionCodes D1,
	 Waterski.dbo.tblDivisionCodes D2, Waterski.dbo.tblPeopleAddresses PA
	 LEFT JOIN Waterski.dbo.tblStates TS ON PA.State = TS.[State Code]

WHERE PA.[Person ID] = PT.[Person ID] AND PA.[Primary] = 1
  AND PT.[Person ID] = MH.[Person ID]
  AND MH.[Person ID] = MHL.PersonID
  AND MH.MembershipHistoryID = MHL.MemHistID
  AND CWS.MembershipHistoryID = MHL.MemHistID 
  AND MH.[Membership Type Code] = MT.[Membership Type Code]
  AND MH.PrimaryDivisionCodeID = D1.DivisionCodeID
  AND MH.SecondaryDivisionCodeID = D2.DivisionCodeID
;


/*										*/
/*	Create short form view USAWaterski.dbo.MemberShort.	*/
/*	Similar to the above full view, except this one only 	*/
/*	accesses the tblPeople and tblPeopleAddresses tables, */
/*	and leaves out the complex joins involved in pulling	*/ 
/*	the latest mbr hist row and related status values.	*/
/*	So this should be speedier where all you need is the	*/
/*	Name and address information for the member.		*/
/*										*/


USE USAWATERSKI
;

ALTER VIEW MemberShort as

SELECT PT.[Person ID] as PersonID, PT.[Name Prefix] as NamePrefix,
	 PT.[First Name] as FirstName, PT.[Middle] as MiddleName,
	 PT.[Last Name] as LastName, PT.[Name Suffix] as NameSuffix,
	 PT.SSN, PT.[Company Name] as CompanyName,
	 Substring(PT.Website,1,100) as Website, PT.Email, PT.MailPref,
	 PT.[Birth Date] as BirthDate, PT.Sex,
	 PT.[Federation Code] as FederationCode,
	 PA.Phone, left(PA.Extension,4) as Extension,
	 PA.Fax, PA.[Business Phone] as BusinessPhone,
	 left(PA.[Business Extension],4) as BusinessExtension,
	 PA.[Mobile Phone] as MobilePhone, PA.Address1, PA.Address2,
	 PA.City, PA.State, PA.Zip, PA.[Country ID] as CountryID,
	 Case when PT.DoNotEMail=1 then '1' else '0' end as DoNotEMail,
	 Coalesce(TS.[Region Code], '6') as Region,
	 PT.[Member Since] as MemberSince, PT.[Date Updated] as DateUpdated,
	 Case when PT.DoNotCall=1 then '1' else '0' end as DoNotCall,
	 Case when PT.Deceased=1 then '1' else '0' end as Deceased

FROM	Waterski.dbo.tblPeople PT
JOIN	Waterski.dbo.tblPeopleAddresses PA
  ON	PA.[Person ID] = PT.[Person ID] AND PA.[Primary] = 1
LEFT JOIN Waterski.dbo.tblStates TS ON PA.State = TS.[State Code]
;