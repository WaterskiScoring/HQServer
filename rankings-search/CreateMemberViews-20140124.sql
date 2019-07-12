
/*																											*/
/*	Create Full View: USAWaterski.dbo.MembersLive.			*/
/*	Fully equivalent to USAWaterski.dbo.Members table,	*/
/*	except does not include PersonIDWithCheckDigit.			*/
/*	Hence must join by Person ID, using rightmost 8			*/
/*	positions of MemberID field from Rankings tables.		*/
/*																											*/

USE USAWATERSKI
;

ALTER VIEW MembersLive as

SELECT PT.[Person ID] as PersonID, PT.[Name Prefix] as NamePrefix,
	 PT.[First Name] as FirstName, PT.[Middle] as MiddleName,
	 PT.[Last Name] as LastName, PT.[Name Suffix] as NameSuffix,
	 PT.SSN, PT.Password, PT.[Company Name] as CompanyName,
	 Substring(PT.Website,1,100) as Website, PT.Email, PT.MailPref,
	 PT.[Birth Date] as BirthDate, PT.Sex, MS.DivisionCode1,
	 MS.DivisionCode2, PT.[Federation Code] as FederationCode,
	 MS.MemberTypeID, PA.Phone, left(PA.Extension,4) as Extension,
	 PA.Fax, PA.[Business Phone] as BusinessPhone,
	 left(PA.[Business Extension],4) as BusinessExtension,
	 PA.[Mobile Phone] as MobilePhone, PA.Address1, PA.Address2,
	 PA.City, PA.State, PA.Zip, PA.[Country ID] as CountryID,
	 MS.MembershipTypeCode, MS.EffectiveFrom, MS.EffectiveTo,
	 Case when PT.DoNotEMail=1 then '1' else '0' end as DoNotEMail,
	 Coalesce(TS.[Region Code], '6') as Region,
	 PT.[Member Since] as MemberSince, PT.[Date Updated] as DateUpdated,
	 Case when PT.DoNotCall=1 then '1' else '0' end as DoNotCall,
	 MS.MembershipType,
	 Case when PT.Deceased=1 then '1' else '0' end as Deceased,
	 MS.WaiverStatusID, MS.WaiverGoodTo, PT.ForeignFederationID

FROM	Waterski.dbo.tblPeople           PT

JOIN	USAWaterski.dbo.MemberStatus     MS
  ON	MS.PersonID = PT.[Person ID]

JOIN	Waterski.dbo.tblPeopleAddresses  PA
  ON	PA.[Person ID] = PT.[Person ID] AND PA.[Primary] = 1

LEFT JOIN Waterski.dbo.tblStates       TS 
  ON PA.State = TS.[State Code]
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
	 PT.SSN, PT.Password, PT.[Company Name] as CompanyName,
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
	 Case when PT.Deceased=1 then '1' else '0' end as Deceased, 
	 PT.ForeignFederationID

FROM	Waterski.dbo.tblPeople PT
JOIN	Waterski.dbo.tblPeopleAddresses PA
  ON	PA.[Person ID] = PT.[Person ID] AND PA.[Primary] = 1
LEFT JOIN Waterski.dbo.tblStates TS ON PA.State = TS.[State Code]
;



/*																										*/
/*	Create view that pulls data from the latest	 			*/
/*	Membership History row, keyed by PersonID.				*/
/*	Also pulls in the latest Annual Waiver Status.	 	*/
/*																										*/

USE USAWATERSKI
;

ALTER VIEW MemberStatus as

Select MH.[Person ID] as PersonID,
	 MHL.MemHistID as MemHistID,
	 D1.[Division Code] as DivisionCode1,
	 D2.[Division Code] as DivisionCode2,
	 left(MH.[Membership Type Code],10) as MembershipTypeCode,
	 MT.MemberTypeID,
	 Left(MT.[Membership Type Description],10) as MembershipType,
	 MH.EffectiveFrom, MH.EffectiveTo, 
	 Coalesce(MW.WaiverGoodTo,cast('2000-01-01' as date)) as WaiverGoodTo,
	 case when Coalesce(MW.WaiverGoodTo,cast('2000-01-01' as date)) >= GetDate()
	 	then 1 else 0 end as WaiverStatusID

FROM	(Select [Person ID] as PersonID, 
		cast(substring(Max(convert(char(8),EffectiveTo,112) 
			+ right(convert(char(8),10000000+MembershipHistoryID),7)),9,7) 
				as integer) as MemHistID
	 From Waterski.dbo.[Membership History]
	 group by [Person ID])                  MHL
	 
JOIN	Waterski.dbo.[Membership History]       MH
  ON	MH.[Person ID] = MHL.PersonID
  AND	MH.MembershipHistoryID = MHL.MemHistID

JOIN	Waterski.dbo.tblMembershipTypeCodes     MT
  ON	MT.[Membership Type Code] = MH.[Membership Type Code]

JOIN	Waterski.dbo.tblDivisionCodes           D1
  ON	D1.DivisionCodeID = MH.PrimaryDivisionCodeID

JOIN	Waterski.dbo.tblDivisionCodes           D2
  ON	D2.DivisionCodeID = MH.SecondaryDivisionCodeID

LEFT JOIN USAWaterski.dbo.MemberWaiver        MW	
  ON	MW.PersonID = MHL.PersonID
;



/*																													*/
/*	Create view that pulls the latest Annual Waiver Status.	*/
/*																													*/

USE USAWATERSKI
;

ALTER VIEW MemberWaiver as

Select MH.[Person ID] as PersonID,
	max(case when DateAdd(dd,364,HW.WaiverStartDate)
		> MH.EffectiveTo then MH.EffectiveTo else
		DateAdd(dd,364,HW.WaiverStartDate)		
			end) as WaiverGoodTo
From Waterski.dbo.[Membership History Waivers]	HW
Join Waterski.dbo.[Membership History]		MH
  on MH.MembershipHistoryID = HW.MembershipHistoryID
  and HW.WaiverStatusID > 0
Group by MH.[Person ID]
;

