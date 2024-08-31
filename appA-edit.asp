<%option explicit

if session("lng") = "fr" then
	Session.LCID = 1036
    else 
    Session.LCID =3081
    end if

'<!-------------------------------------------NOTES ----------
'MODS --
'7 DEC 03 LJL revised main query to use collated joins
'14 FEB 04 LJL added 2nd and 3rd nats for ILO
'23 MAR 04 LJL reduced start date years, lengthened end date years
'28 APR 04 LJL added webfamiliar per PIGUETA taken from additional information pages.
'29 MAY 04 LJL added maxrows to query to avoid double listings when refdb staff data duplicated
'11 JUN 04 LJL changed contract len query to only get those in this org
'7 AUG 04 LJL changed nat specs for internal candidates and tested
'1 NOV 04 LJL changed web familiars to only org specific
'26 MAY 05 LJL updated country list to only include those with names, and per org list
'15 JUN 05 LJL made the honorific only come up once in FR and ES as it was coming up twice - check if null
'17 JUN 05 LJL sex is changed to not modifiable field
'20 JUN 05 LJL adjusted nationalities per WTO - If intern, consultant, or regular member country, see notes within
'21 JUN 05 LJL added check to see if person claims staff number, but names do not match
'11 AUG 05 LJL revised gender selection back to selection
'THE VERSION OF THIS IS NOT THE LATEST.  VERIFY AGAINST CFM
'06 SEP 05 RR reviewed--headers/includes/footers
' 22 OCT 05 LJL added Nationality list instead of country name list for IFRC 7000
' 23 OCT 05 LJL modified so that WTO 3000 users who registered as INTERNS will have a different nat list to choose from.  Interns change over when they apply to a regular post.

'30 NOV 05 LJL lots of changes on this page for internship programme WTO 3000
'25 JAN 06 LJL added validation for valid birth date, as some are entering illegal birth dates.
'13 FEB 06 LJL changed thisorg_staffno to get right value from JAPINFO3 query, for 7000 IFRC
'19 Feb 06 RR added apostrophe code for passport place of issue.
'01 MAR 06 LJL adapted 1st nationality for UNESCO 2500 which have decided not to allow it to be changed so that they can have more control over it. - from experience, they have had problems with applicants over this.
'10 MAR 06 LJL nat assigning for previous nat was wrongly inserting in to the log file, with wrong values
'03 MAY 06 LJL added check for age limit, which originates in Data Elements/Adm Items
'06 MAY 06 LJL added focus back to field in javascript error messages
'16 MAY 06 LJL added set of the intern/regular candidate nationality include text to be used in the javascript validation and in the code below for nat
'18 MAY 06 LJL added set to upper case javascript for LAST NAME
'18 MAY 06 LJL js changed to not allow only spaces for required fields
'04 JUN 06 RR fixed intern year_option code
'20 JUN 06 LJl many small changes for WHO implementation - added ' to honor and sex update in UPDATE query
' NEEDNEED - WHO internals - not showing contract dates and info  - check on this and fix
'08 JUL 06 LJL moved notes to top, added update text as include, which used to be in the main page top include
'16 SEP 06 LJL added UNAIDS and adjusted internal staff for UNAIDS, ILO
'17 OCT 06 LJL contract internal for UNAIDS not right
'23 OCT 06 LJL 2nd and 3rd nats for UNAIDS
'31 OCT 06 LJL changed int employment link to AppF instead of AppF2
'28 Nov 06 ac added trim() to first, last, maiden name
'08 FEB 07 LJL country names len()
'02 APR 07 LJL auto set 1 for thisorg_short when thisorg_staffno is len()
'17 APR 07 LJL added to check on WHO staff and then set var to then change email to the staff email account, and move the primary email to secondary (other) email
'25 APR 07 LJL modified internal staff indicator for WHO
'01 MAY 07 INT Changed queries to parameterized
'14 JUn 07 LJL resolved problem with clearing of contract dates for WHO
'09 JUL 07 LJL ' ADDED INTERN CHECK FOR WTO	for contract dates
'17 JUL 07 LJl corrected problem with varchar in update statement with session("RSYSUSER") as it is varchar
'07 AUG 07 LJL change vals for new queries
'08 aUG 07 LJL internal staff for other orgs than WHO was blocking
'17 SEP 07 LJL interns for WTO were being showed "are you staff mamber" section
' 11 NOV 05 LJL CHANGED to the actual country member number for applicant member state
'03 FEB 08 LJL ALREADYOK was not allowing application to vacancy for WHO.  Changed public/hrd-cl-auth.asp to change the checknatok session var properly
'12 FEB 08 LJL revised the internal staff part for all
'26 JUN 08 LJL revised for GSM - had to add rtrim to all nat fields!!!
'29 LJL 08 LJL revised for GSM - new staff number checking and last name checking
'07 JUL 08 INT added new validation(DataValidation) script to check the start date could not be greater than end date.
'11 NOV 08 LJL revised the GSM DOB to be day month year in the inputs to update
'14 NOV 08 LJL modified email sent when GSM names do not match, WHO and UNAIDS
'14 APR 09 DD  added condition to ensure that internal staff query gets executed in case of session("CLI_INTERNAL_STAFF")
'26 APR 09 LJL added log entries into current year log table for orgs
'6 MAY 09 LJL modified showing of birth year in interns for WTO and others - end and start years were set with no step -1 years getting higher into the drop down, which was opposite of what should be.
'23 MAY 09 DD added check to update staff no. for WHO with upper case.
'01 JUN 09 LJL adjusted CLI_INTERNAL session set for 1500 UNAIDS
'21 JUL 09 DD  Validation added on onblur of last name field for lastname to have minimum 2 apha characters.
'05 Aug 09 DD declaration added for one command object, revised checkbox updatation on submit.
'06 AUG 09 LJL added coding to log when user ACCESSES, or UPDATES profile into td_rsys_candupdates_...
'06 AUG 09 LJL added additional logging for changes to STAFF INT number or status
'17 SEP 09 DD  Size of asterix are increased to bigger size
'24 SEP 09 LJL added 1200 to the OR statement
'26 APR 10 LJL NEW SITE - ADD NEW ORGS TO THIS LIST IF NEEDED CONTRACT DETAILS (most will need them if wanting internal staff info)
'26 APR 10 LJL NEW SITE - must decide if need internal staff number to be varchar or just integer.  Set it in the top of this page.
'26 APR 10 LJL removed the ' for month values for contract start and end months
'01 MAY 10 LJL moved header set and check login to before loading content -
	'order is Edit Section, check_complete, then page headers, then ejobs-updates include
'10 JUL 10 LJL removed <head around js
'13 AUG 10 LJL added old and new bday monitoring
'15 AUG 10 LJL fixed problem with staff number check doing a response.end when it should just add a "warning" text and push that along.
'17 AUG 10 lJL changed to not say 's' for the beginning of the code, so that non-WHO staff don't know.
'23 AUG 10 LJL added check to see if bday is same or not
'10 SEPT 10 DD Removed A/D/C for ITU only
'04 OCT 10 LJL added public PHF viewing indicator
'05 OCT 10 LJL only member states for nat list for ITU
'27 OCT 10 DD SQL query executed by ADODB Command object and NOT by Connection object.
'03 NOV 10 LJL add session set to internal so they can see and apply to internal posts
'03 NOV 10 LJL check if number entered for WIPO
'19 NOV 10 DD SQL query executed by new ADODB Command object always
'10 DEC 10 LJL moved pv_staff_numberer section from just above this section - for non WHO/UNAIDS orgs.  WHO will be set later below
'11 DEC 10 LJL massive re-do of WHO setting of internal or external with thisorg_short.  pv_isSTAFF is used to set this early on and then follow through the rest of the page
'16 DEC 10 LJL heavy revision of the internal/external and staff number system for WHO/UNAIDS
'11 JAN 11 LJL revised to check for and set the contract type for WHO and UNAIDS
'11 JAN 11 LJL changed the birthdate change to be disallowed for GSM entered dates
'06 FEB 11 LJL changed UNESCO nat output view
'07 FEB 11 LJL had problem with WHO - Admin user changed Fixed-Term Appointment to Fixed Term Appointment without hyphen and the system could then not assign it correctly for users.
'27 APR 11 LJL revised marital to be org reflective
'28 APR 11 LJL revised to used marital_id_ session code for marital
'09 MAY 11 DD Hided staff member contract information depending upon either applicant is staff member ONLY for WHO,ILO and UNAIDS orgs
'12 MAY 11 DD Hided staff member contract information depending upon either applicant is staff member for WIPO (new org is added for this functionality)
'22 JUL 11 LJL not show SQL statement for internals
'11 MAR 12 LJL changed ITU int/non-staff to just show and not be editable by applicant
'12 APR 12 LJL added help page for ITU internals
'16 APR 12 LJL added new mail configuration to avoid having to get schemas from Microsoft
'02 MAY 12 LJL parameterized the nats
'11 SEP 12 LJL removed contract internal check for 5500 UPU
'26 SEP 12 LJL modified the to from secant to talenti
'03 NOV 12 LJL order of webfamiliar
'08 JAN 13 LJL fixed month and day to be below 10 to add 0 in front
'17 Jan 13 Pramod Fixed the issue of the manual date change by user, added jquery datepicker. readonly
'27 MAR 13 LJL added Birth_date_new to INTERNAL WHO And UNAIDS staff birth date section
'09 OCT 13 LJL modify date process to deal with odd dates being entered in the system
'10 OCT 13 LJL revised birth date to be yyyy-mm-dd for the new adDBdate parameter setting for insertion
'12 OCT 13 LJL moved date values around in first birthday section
'02 APR 14 LJL changed from <> "" to len() on prev year
'20 DEC 14 LJL removed config for mail for localhost SMTP
'20 DEC 14 LJL added  AND countryquota_rank_i > 0 to the quota listing query
'21 DEC 14 LJL getnat movefirst removed
'25 DEC 14 LJL removed talenti email from notification
'13 JAN 15 LJL changed nat2 for WIPO and all ELSE to not have quota indicator
'31 MAR 15 LJL no staff number for WMO
'31 MAR 15 LJL no contract dates for WMO
'12 MAY 15 LJL ordered webfamiliar by alphabetical
'14 MAY 15 LJL no other names section for WMO	
'05 AUG 15 LJL revised marital to default to nothing initially
'05 AUG 15 LJL revised contract types for WMO to include empty option but check on it when submitted
'14 april line number 2598 manoj
'08 JUN 16 GG Added Check lenght befor left function
'24 mar 16 manoj added check for single and double quotes in names
'7june 18 manoj added validation in firstname,lastname and maiden name to avoid special character
'8 nov 18 manoj add check in second name to allow it to  leave it blank
'01 OCT 19 GG Added Server.HTMLEncode 
'17 AUG 20 LJL changed var MyReg = /^[\a-zA-ZàèìòùÀÈÌÒÙáéíóúýÁÉÍÓÚÝâêîôûÂÊÎÔÛãñõÃÑÕäëïöüÿÄËÏÖÜŸçÇßØøÅåÆæœ ]+$/i; to var MyReg = /^[A-Za-zÀ-ȕ ]+$/i;
'07 DEC 20 LJL moved page headers to above edit messages so that all messages are correctly accented in FR ES...
'04 MAY 22 LJL fixed the send commands for mail
'07 AUG 24 LJL update date
'20 AUG 24 LJL changed last name to UCASE as it wasn't keeping upper


'----------------------------------------------
' NOTES
'NOTE: INTERN SECTION to show about age restrictions if any (for WTO)

' pv_isSTAFF is used to check if the staff number is truly validated in GSM, and then used through the rest of the page to insert GSM data, or no

pv_col2right = "1"
pv_heavybottomvery = "1"


' CHECK LOGGED IN AND DIM DB's
' ***********************************************************%>
<!--#include file = "../includes/include_check_login.asp"-->
<!--#include file = "../../includes/rsys_db.asp"-->
<!--#include file = "../../includes/rsys_db_select.asp"-->
<!--#include file = "../../includes/rsys_int_select.asp"-->
<!--#include file = "../../includes/rsys_logs.asp"-->
<% '<<--Modified by Interface on 05/01/2007 %>
<!--#include file="../../sysdev/adovbs.inc" -->
<%
' ***********************************************************
' END CHECK LOGGED IN AND DIM DB's

dim ctrylang, cand_birth,  year_option, start_year,  end_year, yearget, cstart, year_option2, start_year2,end_year2, yearget2, cend_date,  Acount, cstart_date, birth_date, cend
dim GETNATsql, GETNAT, GETNAT2sql, GETNAT2, getMARITALsql, getMARITAL, getHONORsql, getHONOR, getMONTHSsql, getMONTHS, getFamiliarssql, getFamiliars
dim pv_nat1_quota,pv_nat1_country, pv_nat2_quota,pv_nat2_country, pv_nat3_quota,pv_nat3_country, checkdate, checkdate2, gAGEsql, gAGE
dim JAPINFO3sql, JAPINFO3, JAPINFO4sql, JAPINFO4, pv_birth_date_check, JAPINFO2sql, JAPINFO2, pv_intdetails, pv_nattext
dim GETCONTRACTLENsql, GETCONTRACTLEN, chkAsql, chkA, updAsql, updA, updArankssql, updAranks, goeditsql, goedit, logeditsql, logedit, gITEXTA, faqid, gAGE2, gAGE2sql
dim chkSTAFFsql, chkSTAFF, pv_isSTAFF, GETJAPNATsql, GETJAPNAT, pv_newval, pv_prevval, UPD7000sql, UPD7000
'TEST
dim obj_logs_CmdIX,logedit1,logeditsql1, chkSTAFFNUMBERsql, chkSTAFFNUMBER, pv_isSTAFFNUMBER, pv_staffhistoryold, pv_staffhistorynew, pv_warning, pv_bday_prev, pv_bday_changed, pv_bday_prevform

'<<--Added by Interface on 05/01/2007
dim obj_logs_CmdI, obj_logs_CmdII,obj_logs_CmdIII,obj_logs_CmdIV, obj_logs_CmdV, obj_logs_CmdVI, obj_logs_CmdVII, obj_db_CmdI, obj_db_CmdII, obj_db_select_CmdI, obj_db_select_CmdII
dim obj_db_select_Cmd3, obj_db_select_Cmd4, obj_db_select_Cmd2, obj_db_select_CmdN, obj_db_select_CmdSN
dim obj_db_select_CmdIII, obj_db_select_CmdIV, obj_db_select_CmdV, obj_db_select_CmdVI, obj_db_select_CmdVII, obj_db_select_CmdVIII, obj_db_select_CmdIX, obj_db_select_CmdX, obj_db_select_Cmd5, obj_db_select_CmdRank, obj_db_select_CmdEdit
'<<--Added on 11/24/2008 DD
dim currentYear, obj_logs_CmdBD, goBDsql, goBD, pv_birth_date, pv_bday_prev_month, CHECKWHOSTAFFsql , CHECKWHOSTAFF, pv_DEBUGMAIL

' IF SET TO 1, then mail is sent when testing, 0 = testing
' SET to 0 for testing
dim pv_sendemail
pv_sendemail = 1

pv_DEBUGMAIL = 1
pv_warning = ""
pv_bday_prev = ""
pv_bday_changed = ""
pv_bday_prevform = ""

'response.write "<Br>TESTER INT1: " & session("CLI_INTERNAL_STAFF")

'-->>
' BEGIN INCLUDE TEXT DB
dim gITEXTAsql
'<<--Modified by Interface on 05/01/2007
gITEXTAsql = " SELECT i_1, i_74, i_52, i_51, i_23, i_53, i_58, i_2, i_8, i_3, i_5, i_30, i_38, i_39, i_40, i_41, i_48, i_49, i_56, i_55, i_4, i_9, i_29, i_10, i_6, i_11, i_text2, i_64, i_67, i_68, i_69, i_7,  i_12, i_14, i_15, i_19, i_26, i_27, i_28, i_17, i_18, i_20, i_66, i_67, i_68, i_69, i_21, i_59, i_60, i_87, i_88, i_62, i_63, i_64, i_22, i_71, i_72, i_24, i_81, i_89, i_90 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'A' "
obj_int_select_Cmd.CommandText = gITEXTAsql
Set gITEXTA = obj_int_select_Cmd.Execute(,Array(session("template_org_code"),session("lng")))
'-->>
' END INCLUDE TEXT DB


pv_page_title = gITEXTA("i_1")

titletype = 6
'faqid=10
faqid=11
currentYear =  year(now())  'Added on 11/24/2008 DD, to get the value of current year.

pv_isSTAFF = 0
pv_isSTAFFNUMBER = 0


' ***********************************************************
' BEGIN INCLUDES
' ***********************************************************%>
<!--#include file = "../includes/include_check_complete.asp"-->
<!--#include file="../includes/include_publicedit_frame_top.asp"-->
<%
' ***********************************************************
' END INCLUDES
' ***********************************************************

'******************************
' BEGIN A EDIT
'******************************

If Request.form("DOeditA") = "99" then

dim pv_thisorg_staffno
if len(request.form("thisorg_staffno")) then
	pv_thisorg_staffno = Trim(Replace(request.form("thisorg_staffno"), "'", "''"))
else
	pv_thisorg_staffno = ""
end if

'06 AUG 09 LJL Added if they UPDATED THEIR PROFILE PORTION ----------------------------------------------------
dim obj_logs_CmdPU, PAEDITsql, PAEDIT
	set obj_logs_CmdPU = server.CreateObject("adodb.command")
	obj_logs_CmdPU.ActiveConnection = rsys_logs
	If session("RSYSUSER") <> "" then
		PAEDITsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  " (candupd_ip_c, cand_id_c, candupd_page_t,user_id_t,candupd_site_c) " &_
		"VALUES (?, ?, 'A-Edited', ?,?)"
		obj_logs_CmdPU.CommandText = PAEDITsql
		Set PAEDIT = obj_logs_CmdPU.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),session("RSYSUSER"),left(session("template_org_name"),20)))
	Else
		PAEDITsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  " (candupd_ip_c, cand_id_c, candupd_page_t,candupd_site_c) " &_
		"VALUES (?, ?, 'A-Edited',?)"
		obj_logs_CmdPU.CommandText = PAEDITsql
		Set PAEDIT = obj_logs_CmdPU.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),left(session("template_org_name"),20)))
	End if
' -------------------------------------------------------------------------------------------------------------------------------------------

if cend_date = "" then
	cend_date = ""
end if

if not cstart_date = "" then
	cstart_date = ""
end if


' NEEDNEEDNEED LISTGETAT REPLACEMENT MADE HERE
'02 MAY 12 LJL parameterized the nats
	dim nat11, nat12, nat21, nat22, nat31, nat32, pv_nat1, pv_nat2, pv_nat3
			' response.write "T START=="
			if len(trim(request.form("cand_nat_c"))) then
				pv_nat1 = trim(request.form("cand_nat_c"))
				nat11 = instrRev(pv_nat1,"|")+1
				nat12 = instr(pv_nat1,"|")-1
				' response.write "T nat " & request.form("cand_nat_c") & " T nat11" & nat11 & " T nat12" & nat12
				pv_nat1_quota = mid(pv_nat1,nat11, nat11)
				pv_nat1_country = left(pv_nat1,nat12)
			else
				pv_nat1_quota = ""
				pv_nat1_country = ""
			end if
			'29 JUN 08 LJL revised for GSM
	'if session("template_org_code") <> 1000 then
			if len(trim(request.form("cand_nat2_c"))) then
				pv_nat2 = trim(request.form("cand_nat2_c"))
				nat21 = instrRev(pv_nat2,"|")+1
				nat22 = instr(pv_nat2,"|")-1
				pv_nat2_quota = mid(pv_nat2,nat21, nat21)
				
				'08 JUN 16 GG Added Check lenght befor left function
				if nat22 > 0 Then
					pv_nat2_country = left(pv_nat2,nat22)
				Else
					pv_nat2_country = pv_nat2
				End if
			else
				pv_nat2_quota = ""
				pv_nat2_country = ""
			end if
			if len(trim(request.form("cand_nat3_c"))) then
				pv_nat3 = trim(request.form("cand_nat3_c"))
				nat31 = instrRev(pv_nat3,"|")+1
				nat32 = instr(pv_nat3,"|")-1
				pv_nat3_quota = mid(pv_nat3,nat31, nat31)
				pv_nat3_country = left(pv_nat3,nat32)
			else
				pv_nat3_quota = ""
				pv_nat3_country = ""
			end if
	'end if


'18 JUN 08 LJL added new check and language for WHO-----------------------------------------------
if session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500 then
	'10 DEC 10 lJL added check if set to yes for WHO/UNAIDS
	if request.form("thisorg_short") = "1" then
		If len(pv_thisorg_staffno) then

		else
		'17 AUG 10 lJL changed to not say 's' for the beginning of the code, so that non-WHO staff don't know.
		pv_warning = pv_warning & "Warning:<br>You must enter a valid Staff Number (must commence with appropriate code) or clear the field for the " &  session("template_org_name") & " staff Number.<br><br>Please hit BACK on your browser and either enter your " & session("template_org_name") & " staff number or leave it completely blank.<br><br>"
			'response.end
		End If
	end if

'else
'	If pv_thisorg_staffno = "" then

'	elseif len(pv_thisorg_staffno) AND (NOT isnumeric(pv_thisorg_staffno) OR NOT pv_thisorg_staffno > "0") then
'	pv_warning = pv_warning & "<br>You must enter only a number or leave clear the field for the  " &  session("template_org_name") & " staff Number.<br>Please hit BACK on your browser and either enter your  " &  session("template_org_name") & " staff number or leave it completely blank.<br><br>"
			'response.end
'	End If

'end if

'03 NOV 10 LJL check if number entered for WIPO ----------------------------------------------
elseif session("template_org_code") = 2800 then
	If request.form("thisorg_short") = "1" AND pv_thisorg_staffno = "" then
	'17 AUG 10 lJL changed to not say 's' for the beginning of the code, so that non-WHO staff don't know.
	pv_warning = pv_warning & "<br>Important:<br>If you are currently employed by " &  session("template_org_name") & ", you must enter a valid staff number (matricule).<br><br>To do so, click BACK on your browser window and enter your " & session("template_org_name") & " staff number (matricule).<br>If you are not a WIPO employee, then leave the Yes/No as 'No'.<br><br>"
			'response.end
	elseif len(pv_thisorg_staffno) AND (NOT isnumeric(pv_thisorg_staffno) OR NOT pv_thisorg_staffno > "0") then
	pv_warning = pv_warning & "<br>Important:<br>You must enter only a number or leave field clear for the  " &  session("template_org_name") & " staff Number.<br><Br>Please hit BACK on your browser and either enter your  " &  session("template_org_name") & " staff number or leave it completely blank.<br><br>"
			'response.end
	End If


'03 NOV 10 LJL check if number entered for WMO ----------------------------------------------
elseif session("template_org_code") = 2900 then
	If request.form("thisorg_short") = "1" AND request.form("thisorg_type") = "0" then
	'17 AUG 10 lJL changed to not say 's' for the beginning of the code, so that non-WHO staff don't know.
	pv_warning = pv_warning & gITEXTA("i_text2")
	'pv_warning = pv_warning & "<br><font color=maroon><strong>Important:</strong></font><br><br>If you are currently employed by " &  session("template_org_name") & ", you must enter a valid contract type.<br><Br>To do so, click BACK on your browser and enter the contract type.<br><br><br>"
			'response.end
	End If



else
	If request.form("thisorg_short") = "1" AND pv_thisorg_staffno = "" then

	elseif len(pv_thisorg_staffno) AND (NOT isnumeric(pv_thisorg_staffno) OR NOT pv_thisorg_staffno > "0") then
	pv_warning = pv_warning & "Warning:<br>You must enter only a number or leave clear the field for the  " &  session("template_org_name") & " staff Number.<br><br>Please hit BACK on your browser and either enter your  " &  session("template_org_name") & " staff number or leave it completely blank.<br><br>"
			'response.end
	End If

end if


'05 AUG 15 LJL check for marital.civil status
	If request.form("form_marital_id") = "0" then
		
  		'response.write("<font face=arial><br><br>Please go back and select your civil status<br><br><br>")
  		response.write "<font face=arial><br><br>" & gITEXTA("i_41") & "<br><br><br></font>"
  		response.end
	End If


     
		dim year_d, month_d, day_d, datepicker1nullcheck	
		
		'09 OCT 13 LJL modify date process to deal with odd dates being entered in the system
        datepicker1nullcheck=request.form("birth_date_new")
        if Trim(datepicker1nullcheck)="" then
		 response.write "Date of Birth is Required Field"
		 response.end
		 'alert('Date of Birth is Required Field');
         'datepicker1nullcheck="01-01-1900"
        end if



		if Trim(datepicker1nullcheck)<>"" then
		
        '***********************M.J VALIDATING DATE AND DATE FORMAT 11-22-2014**************
         dim validate '***Mj 26 nov change ***

          
          If isdate(datepicker1nullcheck) then
	      validate =datepicker1nullcheck
  	      Else
		  response.write "Date of Birth is Required Field"
		  response.end
		  'alert('Date of Birth is Required Field');
	      'validate = "01-01-1900"
	      End If
          
        ' validate=FormatDateTime(Request.form("birth_date_new"))		 
         birth_date =  year(validate) & "-" & month(validate) & "-" & day(validate)	 
	 	 pv_birth_date = birth_date 
	 	 pv_birth_date_check = day(validate) & "-" & month(validate) & "-" & year(validate)
	 	 '***********************M.J VALIDATING DATE AND DATE FORMAT 11-22-2014**************

		'09 OCT 13 LJL modify date process to deal with odd dates being entered in the system
		else
				response.write "<br>The date was not correctly entered.<br>It has been modified to a generic date.<br>If you continue to have issues, you may <a href='mailto:erecruit@erecruithelp.org?subject=EREC: Incorrect AppA date-" & request.form("birth_date_new") & "' target='_blank'>click here to contact us.</a><br><Br>"
			response.end
        '10 OCT 13 LJL revised birth date to be yyyy-mm-dd for the new adDBdate parameter setting for insertion
			birth_date = "1990-01-01" 
	 	    pv_birth_date = birth_date
	 	    pv_birth_date_check = "01-01-1990"
	 	
		end if 
        '13 AUG 10 LJL added old and new bday monitoring
		pv_bday_changed = pv_birth_date
        
        dim previous_bdy,previous_bdy_check '***26 nov mj added  
        

      previous_bdy_check=  Request.form("pv_bday_prev")
      If isdate(previous_bdy_check) then
	  previous_bdy =previous_bdy_check
  	  Else
	  previous_bdy = "01-01-1900"
	  End If
            

      '  previous_bdy=FormatDateTime(Request.form("pv_bday_prev"))      
      
      '26 nov 2014 end   
		pv_bday_prevform = year(previous_bdy) & "-" & month(previous_bdy) & "-" & day(previous_bdy)

        '23 AUG 10 LJL added check to see if bday is same or not
        if pv_bday_changed <> pv_bday_prevform then

                      set obj_logs_CmdBD = server.CreateObject("adodb.command")
                      obj_logs_CmdBD.ActiveConnection = rsys_logs
                      goBDsql = "INSERT INTO tx_rsys_log_candhistory_" & session("template_org_code") & "_" & currentYear &  " (cand_id_c, candhistory_type_c, candhistory_value_prev_t, candhistory_value_new_t, candhistory_ip_c, candhistory_type_t) " &_
                      "VALUES (?,'BD', ?,?,?, 'DOB' )"
                      obj_logs_CmdBD.CommandText = goBDsql
                      Set goBD = obj_logs_CmdBD.Execute(,Array(session("RSYS_EVAL"),pv_bday_prevform,pv_bday_changed,left(request.servervariables("REMOTE_ADDR"), 25)))

         end if

'response.write "<br>TEST DATE: " & birth_date
'response.end


		' CHECK BIRTH DATE TO VERIFY IF LEGAL (is a valid date at all)

	if isdate(pv_birth_date_check) = false then
			pv_warning = pv_warning & "<br>Please enter a valid date for your date of birth.<br><br>The combination entered does not seem to be a valid date.<br><br>" & pv_birth_date_check
			'response.end
	'else
			' response.write "OKAY BUD"
	end if


'---------------------------------------------------------------------------------------------------------------------------------------------
' ADDED INTERN CHECK FOR WTO
' CHECK contract start dates for any orgs that have dates
'21 JUL 10 LJL not for WIPO so far
'11 SEP 12 LJL removed contract internal check for 5500 UPU
'12 FEB 15 LJL added UNWOMEN and WMO
'31 MAR 15 LJL WMO doesn't need more info that the contract type for internals
'28 OCT 15 LJL added WTO to internal check

if session("template_org_code") <> 2400 AND session("template_org_code") <> 2500 AND session("template_org_code") <> 2800 AND session("template_org_code") <> 2900 AND session("template_org_code") <> 3000 AND session("template_org_code") <> 4000 AND session("template_org_code") <> 5500 then
if session("rsys_intern") = "0" AND request.form("thisorg_short") <> "" then

	If isnumeric(Request.form("cstart_month")) AND isnumeric(Request.form("cstart_day")) AND isnumeric(Request.form("cstart_year")) then
		checkdate = request.form("cstart_month") & "/" & request.form("cstart_day") & "/" & request.form("cstart_year")
  		If isdate(checkdate) then

  		Else
			pv_warning = pv_warning & "<br><br><font  color='Maroon'><strong>You have entered a contract START date which is not valid.<br>Please check the month, day and year for a correct date.<br><br>Example November does not have 31 days, but only 30.<br><br>You may return to the profile by clicking <a href='appA-edit.asp'>here</a>. (err: SD01)</font>"
		  End If
	Else
		If session("CLI_INTERNAL_STAFF") = "1" then
		  checkdate = ""
		Else
	  		If request.form("thisorg_short") = "1" then
  				pv_warning = pv_warning & "<font  color='Maroon'><strong>You must enter a contract START date if you have indicated that you are " & session("template_org_name") & " staff.<br><br>You may return to the profile by clicking <a href='appA-edit.asp'>here</a>. (err: SD02)</font>"
  				'response.end
  			End If
  		End If
	end if

	If isnumeric(Request.form("cend_month")) AND isnumeric(Request.form("cend_day")) AND isnumeric(Request.form("cend_year")) then
		checkdate2 = request.form("cend_month") & "/" & request.form("cend_day") & "/" & request.form("cend_year")
  		If isdate(checkdate2) then

  		Else
				pv_warning = pv_warning & "<br><br><font  color='Maroon'><strong>You have entered a contract END date which is not valid.<br>Please check the month, day and year for a correct date.<br><br>Example November does not have 31 days, but only 30.<br><br>You may return to the profile by clicking <a href='appA-edit.asp'>here</a>. (err: ED01)</font>"
  			End If
  Else
	If session("CLI_INTERNAL_STAFF") = "1" then
		checkdate2 = ""
	Else
		If request.form("thisorg_short") = "1" then
		pv_warning = pv_warning & "<font  color='Maroon'><strong>You must enter a contract END date if you have indicated that you are " & session("template_org_name") & " staff.<br><br>You may return to the profile by clicking <a href='appA-edit.asp'>here</a>. (err: ED02)</font>"
  		'response.end
		End If
	End If
  End If
end if
end if
'---------------------------------------------------------------------------------------------------------------------------------------------

' 17 APR 07 LJL added to check on WHO staff and then set var to then change email to the staff email account, and move the primary email to secondary (other) email
'25 APR 07 LJL set this as default
'pv_isSTAFF = 0
  set obj_db_select_Cmd = server.CreateObject("adodb.command")
  obj_db_select_Cmd.ActiveConnection = rsys_db

''''  WHO STAFF CHECK WAS HERE


'06 AUG 09 LJL -  FOR OTHER ORGS TO GET STAFF DETAILS TO SAVE IN HISTORY----------------------------
	dim  obj_db_select_CmdIS, chkINTSTAFFsql, chkINTSTAFF, pv_staff_temp
  set obj_db_select_CmdIS = server.CreateObject("adodb.command")
  obj_db_select_CmdIS.ActiveConnection = rsys_db_select
  chkINTSTAFFsql = " SELECT cand_thisorg_short_i_" & session("template_org_code") & " AS thisorg_isstaff, staff_nbr_" & session("template_org_code") & " AS thisorg_staffnumber FROM td_rsys_cand WHERE cand_id_c = ?"
  obj_db_select_CmdIS.CommandText = chkINTSTAFFsql
  Set chkINTSTAFF = obj_db_select_CmdIS.Execute(,Array(Trim(session("RSYS_EVAL"))))

'06 AUG 09 LJL -  FOR OTHER ORGS TO GET STAFF DETAILS TO SAVE IN HISTORY
  if chkINTSTAFF.eof = false then
  	if  chkINTSTAFF("thisorg_isstaff") = 1 then
  		pv_staff_temp = "Is Internal"
  	else
  		pv_staff_temp = "Not Internal"
  	end if
  	pv_staffhistoryold = pv_staff_temp & " | SN=" & chkINTSTAFF("thisorg_staffnumber")
  else
  	pv_staffhistoryold = "Not Internal | SN=[blank]"
  end if
  if len(request.form("thisorg_short")) OR len(pv_thisorg_staffno) then
  	if  request.form("thisorg_short") = 1 then
  		pv_staff_temp = "Is Internal"
  	else
  		pv_staff_temp = "Not Internal"
  	end if
  	pv_staffhistorynew = pv_staff_temp & " | SN=" & pv_thisorg_staffno
  else
  	pv_staffhistorynew = "Not Internal | SN=[blank]"
  end if
'---------------------------------------------------------------------------------------------------------------------------------------------
'06 AUG 09 LJL   <!------------------- INSERT HISTORY OF CHANGES ON STAFF NUMBER AND SETTING ------------------------------>
' If len(pv_thisorg_staffno) then
 '<<--Modified by Interface on 05/01/2007
 if pv_staffhistorynew <> pv_staffhistoryold then
  set obj_logs_CmdIII = server.CreateObject("adodb.command")
  obj_logs_CmdIII.ActiveConnection = rsys_logs
  goeditsql = "INSERT INTO tx_rsys_log_candhistory_" & session("template_org_code") & "_" & currentYear &  " (cand_id_c, candhistory_type_c, candhistory_value_prev_t, candhistory_value_new_t, candhistory_ip_c, candhistory_type_t) " &_
  "VALUES (?,'ST', ?,?,?, 'STAFF NUMBER' )"
  obj_logs_CmdIII.CommandText = goeditsql
  Set goedit = obj_logs_CmdIII.Execute(,Array(session("RSYS_EVAL"),pv_staffhistoryold,pv_staffhistorynew,left(request.servervariables("REMOTE_ADDR"), 25)))
  '-->>
'End If
end if
'-------------------------------------------------------------------------------------------------------

'<<--Modified by Interface on 05/01/2007
  set obj_db_select_CmdI = server.CreateObject("adodb.command")
  obj_db_select_CmdI.ActiveConnection = rsys_db
  chkAsql = " SELECT cand_io_i, cand_email_t, cand_wml_t FROM td_rsys_cand WHERE cand_id_c = ?"
  obj_db_select_CmdI.CommandText = chkAsql
  Set chkA = obj_db_select_CmdI.Execute(,Array(session("RSYS_EVAL")))
'-->>
 ' <!------------------------ moved to here from appJ per PIGUETA 28 APR 04 ---------------->
'<<--Modified by Interface on 05/01/2007
set obj_db_CmdII = server.CreateObject("adodb.command")
obj_db_CmdII.ActiveConnection = rsys_db

updAsql = " UPDATE td_rsys_cand SET webfamiliar_id_c = ?, webfamiliar_refer_t = ?, "
updAsql = updAsql & " cand_bthp_t = ?, "
updAsql = updAsql & " cand_bth_d = ?, "
updAsql = updAsql & " cand_expl_t = ?, "
updAsql = updAsql & " cand_gnd_i = ?, "
updAsql = updAsql & " cand_fnam_t = ?, "
updAsql = updAsql & " cand_lnam_t = ?, "
updAsql = updAsql & " cand_maiden_t = ?, "

'27 APR 11 LJL revised to used marital_id_ session code for marital
updAsql = updAsql & " marital_id_" & session("template_org_code") & " = ?, "

'' REMOVE THIS WHEN DONE
updAsql = updAsql & " cand_mar_st_c = ?, "

updAsql = updAsql & " cand_mnam_t = ?, "


if len(request.form("webfamiliar_id_c")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@webfamiliar_id_c" ,adInteger ,adParamInput,4, request.form("webfamiliar_id_c"))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@webfamiliar_id_c" ,adInteger ,adParamInput,4, null)
end if

if len(request.form("webfamiliar_refer_t")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@webfamiliar_refer_t" ,adVarChar ,adParamInput,510, Trim(request.form("webfamiliar_refer_t")))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@webfamiliar_refer_t" ,adVarChar ,adParamInput,510, null)
end if

if len(request.form("cand_bthp_t")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_bthp_t" ,adVarChar ,adParamInput,100, request.form("cand_bthp_t"))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_bthp_t" ,adVarChar ,adParamInput,100, request.form("cand_bthp_t"))
end if

obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_bth_d" ,adDBDate ,adParamInput,8, birth_date)

if len(request.form("cand_expl_t")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_expl_t" ,adVarChar ,adParamInput,510, Trim(request.form("cand_expl_t")))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_expl_t" ,adVarChar ,adParamInput,510, null)
end if

obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_gnd_i" ,adTinyInt,adParamInput,1, request.form("cand_gnd_i"))

obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_fnam_t" ,adVarChar ,adParamInput,510, Trim(request.form("cand_fnam_t")))

'20 AUG 24 LJL changed last name to UCASE as it wasn't keeping upper
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_lnam_t" ,adVarChar ,adParamInput,510, UCASE(Trim(request.form("cand_lnam_t"))))

if len(request.form("cand_maiden_t")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_maiden_t" ,adVarChar ,adParamInput,100, request.form("cand_maiden_t"))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_maiden_t" ,adVarChar ,adParamInput,100, null)
end if

'28 APR 11 LJL revised for marital issues
if len(request.form("form_marital_id")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@form_marital_id" ,adInteger ,adParamInput,4, request.form("form_marital_id"))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@form_marital_id" ,adInteger ,adParamInput,4, null)
end if


'''' REMOVED THIS WHEN DONE WITH MARITAL
if len(request.form("form_marital_id")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@form_marital_id" ,adInteger ,adParamInput,4, request.form("form_marital_id"))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@form_marital_id" ,adInteger ,adParamInput,4, null)
end if


if len(request.form("cand_mnam_t")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_mnam_t" ,adVarChar ,adParamInput,100, Trim(request.form("cand_mnam_t")))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_mnam_t" ,adVarChar ,adParamInput,100, null)
end if

'17 OCT 06 LJL added check for staff number
'if (session("template_org_code") = 1000 OR session("template_org_code") = 1500) then


updAsql = updAsql & " cand_nat_c = ?, "
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_nat_c" ,adVarChar ,adParamInput,6,pv_nat1_country)
'IF session("template_org_code") <> 1000 then
	updAsql = updAsql & " cand_nat2_c = ?, "
	updAsql = updAsql & " cand_nat3_c = ?, "
	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_nat2_c" ,adVarChar ,adParamInput,6,pv_nat2_country)
	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_nat3_c" ,adVarChar ,adParamInput,6,pv_nat3_country)
'end if

updAsql = updAsql & " cand_newnat_c = ?, "
updAsql = updAsql & " cand_newnat_d = ?, "
updAsql = updAsql & " cand_onam_t = ?, "
updAsql = updAsql & " cand_pnat_i = ?, "

if len(request.form("cand_newnat_c")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_newnat_c" ,adVarChar ,adParamInput,100,request.form("cand_newnat_c"))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_newnat_c" ,adVarChar ,adParamInput,100,null)
end if

if len(request.form("cand_newnat_d")) then     
    obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_newnat_d" ,adVarChar ,adParamInput,100,request.form("cand_newnat_d"))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_newnat_d" ,adVarChar ,adParamInput,100,null)
end if

if len(request.form("cand_onam_t")) then
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_onam_t" ,adVarChar ,adParamInput,100,Trim(request.form("cand_onam_t")))
else
obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_onam_t" ,adVarChar ,adParamInput,100,null)
end if

obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_pnat_i" ,adTinyInt ,adParamInput,1,request.form("cand_pnat_i"))

if request.form("thisorg_type") <> "" then

    dim paramcontracttype 

    if request.form("thisorg_short") ="0" then
    paramcontracttype ="0"
    else
    paramcontracttype =request.form("thisorg_type")
    end if 
	updAsql = updAsql & " cand_thisorg_jobtype_i_" & session("template_org_code") & "  = ?, "
	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_thisorg_jobtype_i_" ,adSmallInt ,adParamInput,2,paramcontracttype)
end if

if request.form("thisorg_prev") <> "" then
	updAsql = updAsql & " cand_thisorg_prev_i_" & session("template_org_code") & "  = ?,"
	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_thisorg_prev_i_" ,adTinyInt ,adParamInput,1,request.form("thisorg_prev"))
end if

' 17 APR 07 LJL added to check on WHO staff and then set var to then change email to the staff email account, and move the primary email to secondary (other) email
' MAKES SURE NOT to change it if they have already indicated their WHO email as primary
if session("template_org_code") = "1000" AND pv_isSTAFF = 1 then
	if chkSTAFF("email") <> chkA("cand_email_t") then
		updAsql = updAsql & " cand_email_t = ?, "
		updAsql = updAsql & " cand_wml_t = ?, "
		obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_email_t" ,adVarChar,adParamInput,510,chkSTAFF("email"))
		obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_wml_t" ,adVarChar ,adParamInput,510,chkA("cand_email_t"))
		'<<--Modified by Interface on 05/01/2007
  		set obj_logs_CmdI = server.CreateObject("adodb.command")
		obj_logs_CmdI.ActiveConnection = rsys_logs
  		goeditsql = " INSERT INTO tx_rsys_log_candhistory_" & session("template_org_code") & "_" & currentYear & " (cand_id_c, candhistory_type_c, candhistory_value_prev_t, candhistory_value_new_t, candhistory_ip_c, candhistory_type_t) VALUES "&_
  		" ( ?, 'EM', ?, ?, ?, 'Email-Secondary-WHO' ) "
  		obj_logs_CmdI.CommandText = goeditsql

		if len(chkA("cand_wml_t")) then
			pv_prevval = chkA("cand_wml_t")
		else
			pv_prevval = "NULL"
		end if
		if len(chkA("cand_email_t")) then
			pv_newval = chkA("cand_email_t")
		else
			pv_newval = "NULL"
		end if

		Set goedit = obj_logs_CmdI.Execute(,Array(session("RSYS_EVAL"),pv_prevval,pv_newval,left(request.servervariables("REMOTE_ADDR"), 25)))

  		set obj_logs_CmdII = server.CreateObject("adodb.command")
		obj_logs_CmdII.ActiveConnection = rsys_logs
  		goeditsql = "  INSERT INTO tx_rsys_log_candhistory_" & session("template_org_code") & "_" & currentYear &  " (cand_id_c, candhistory_type_c, candhistory_value_prev_t, candhistory_value_new_t, candhistory_ip_c, candhistory_type_t) VALUES "&_
  		" (?, 'EM', ?,?, ?, 'Email-Primary-WHO' )"
  		obj_logs_CmdII.CommandText = goeditsql
		Set goedit = obj_logs_CmdII.Execute(,Array(session("RSYS_EVAL"),chkA("cand_email_t"),chkSTAFF("email"),left(request.servervariables("REMOTE_ADDR"), 25)))
  		'-->>
	end if
end if


'10 DEC 10 LJL moved thisorg_short update to outside WHO/UNAIDS restriction below
if len(request.form("thisorg_short")) then

'10 DEC 10 LJL possibly set this to not WHO/UNAIDS so that it is set later with CLI_INTERNAL
	if NOT session("template_org_code") = 1000 AND NOT session("template_org_code") = 1500 then
		if request.form("thisorg_short") = 0  then
			pv_staffer_number = 0
			'03 NOV 10 LJL add session set to internal so they can see and apply to internal posts
			session("CLI_INTERNAL_STAFF") = 0
		else
			pv_staffer_number = 1
			'03 NOV 10 LJL add session set to internal so they can see and apply to internal posts
			session("CLI_INTERNAL_STAFF") = 1
		end if
	else
		pv_staffer_number = 1
	end if

		updAsql = updAsql & " cand_thisorg_short_i_" & session("template_org_code") & "  = ?, "
		obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_thisorg_short_i_" ,adTinyInt ,adParamInput,1,request.form("thisorg_short"))

end if

'05/23/2009 DD, Check is added as staffno of who is asked to be updated with upper case.
'24 SEP 09 LJL added 1200 and 1500 for internal update
if session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500 then

	if request.form("thisorg_short") = 1 then

		if len(pv_thisorg_staffno) AND pv_staffer_number = 1 then
			updAsql = updAsql & " staff_nbr_" & session("template_org_code") & " = ?, "
			obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@staff_nbr_" ,adVarChar ,adParamInput,24,trim(UCASE(pv_thisorg_staffno)))
		else
			updAsql = updAsql & " staff_nbr_" & session("template_org_code") & " = NULL, "
		end if
	' UPDATE to NULL if set to No
	else
			updAsql = updAsql & " staff_nbr_" & session("template_org_code") & " = NULL, "
	end if

else
		if len(pv_thisorg_staffno) AND pv_staffer_number = 1 then
			updAsql = updAsql & " staff_nbr_" & session("template_org_code") & " = ?, "
			obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@staff_nbr_" ,adVarChar ,adParamInput,24,trim(UCASE(pv_thisorg_staffno)))
		else
			updAsql = updAsql & " staff_nbr_" & session("template_org_code") & " = '', "
		end if

		'10 DEC 10 LJL moved from just above this section - for non WHO/UNAIDS orgs.  WHO will be set later below
		'06 AUG 09 LJL this section changes the short (staff) to yes if there is still a staff number.  Better to erase the staff number if set to NO??
		dim pv_staffer_number

end if

'if len(pv_thisorg_staffno) then
'	updAsql = updAsql & " cand_thisorg_short_i_" & session("template_org_code") & "  = 1, "
'elseif request.form("thisorg_short") <> "" then
'	updAsql = updAsql & " cand_thisorg_short_i_" & session("template_org_code") & "  = ?, "
'	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_thisorg_short_i_" ,adTinyInt ,adParamInput,1,request.form("thisorg_short"))
'else
'	updAsql = updAsql & " cand_thisorg_short_i_" & session("template_org_code") & "  = 0, "
'end if

'02 APR 14 LJL changed from <> "" to len() on prev year
if len(request.form("cand_prev_year_c")) then
	updAsql = updAsql & " cand_prev_year_c = ?,"
	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_prev_year_c" ,adVarChar,adParamInput,8,request.form("cand_prev_year_c"))
end if

	IF request.form("thisorg_short") = "1" AND chkA("cand_io_i") = 0 then
		updAsql = updAsql & " cand_io_i = 1,"
	end if

	updAsql = updAsql & " honor_id_c = ?, user_id_t = ?, upd_d = getdate() WHERE cand_id_c = ?"

if len(request.form("honor_id_c")) then
	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@honor_id_c" ,adSmallInt ,adParamInput,2,request.form("honor_id_c"))
else
	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@honor_id_c" ,adSmallInt ,adParamInput,2,null)
end if

	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@user_id_t" ,adVarChar,adParamInput,100,left(session("rsysuser"),50))
	obj_db_CmdII.Parameters.Append obj_db_CmdII.CreateParameter("@cand_id_c" ,adInteger ,adParamInput,4,session("RSYS_EVAL"))

obj_db_CmdII.CommandText = updAsql
'response.write "<br><Br>T SQL: " & updAsql
'response.write "<br><br>BTH: " & birth_date
'response.End
set updA = obj_db_CmdII.Execute()

'response.write "TESTERSQL: " & updAsql

'-->>

 ' UPDATE PASSPORT DETAILS IF NEEDED FOR IFRC 7000
if session("template_org_code") = 7000 then
'<<--Modified by Interface on 05/01/2007
	set obj_db_CmdI = server.CreateObject("adodb.command")
	obj_db_CmdI.ActiveConnection = rsys_db
	UPD7000sql = ""
	UPD7000sql = 	UPD7000sql & "UPDATE tx_rsys_candmisc SET candmisc_passport_c = ?"
	if len(request.form("candmisc_passport_c")) then
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@candmisc_passport_c" ,adVarChar ,adParamInput,100,request.form("candmisc_passport_c"))
	else
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@candmisc_passport_c" ,adVarChar ,adParamInput,100,null)
	end if

	UPD7000sql = 	UPD7000sql & ", candmisc_passport_place_c = ?"
	if len(request.form("candmisc_passport_place_c")) then
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@candmisc_passport_place_c" ,adVarChar ,adParamInput,100,request.form("candmisc_passport_place_c"))
	else
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@candmisc_passport_place_c" ,adVarChar ,adParamInput,100,null)
	end if

	UPD7000sql = 	UPD7000sql & ",candmisc_passport_issue_d = ?"
	if len(request.form("candmisc_passport_issue_d")) then
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@candmisc_passport_issue_d" ,adVarChar ,adParamInput,60,request.form("candmisc_passport_issue_d"))
	else
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@candmisc_passport_issue_d" ,adVarChar ,adParamInput,60,null)
	end if

	UPD7000sql = 	UPD7000sql & ", candmisc_passport_valid_d = ?"
	if len(request.form("candmisc_passport_valid_d")) then
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@candmisc_passport_valid_d" ,adVarChar ,adParamInput,60,request.form("candmisc_passport_valid_d"))
	else
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@candmisc_passport_valid_d" ,adVarChar ,adParamInput,60,null)
	end if

	UPD7000sql = 	UPD7000sql & " WHERE cand_id_c = ? "
	obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_id_c" ,adInteger ,adParamInput,4,session("RSYS_EVAL"))
	obj_db_CmdI.CommandText = UPD7000sql
	Set UPD7000 = obj_db_CmdI.Execute()
'-->>
end if


set obj_db_select_CmdRank = server.CreateObject("adodb.command")
obj_db_select_CmdRank.ActiveConnection = rsys_db

  updArankssql = "UPDATE tx_rsys_candrank SET candrank_c1_" & session("template_org_code") &  "= "
  	If len(pv_nat1_quota) then
		updArankssql = updArankssql & "'" & pv_nat1_quota & "' "
	Else
		updArankssql = updArankssql & "0 "
	End If
	'<!---------------- aDD ORG INFO 1000 2000 3000 4000 5000 19 JUN 05 LJL ------------->
	If session("template_org_code") <> 1000 then
	 	updArankssql = updArankssql & ", candrank_c2_" & session("template_org_code") &  "= "
	 	If len(pv_nat2_quota) then
	 		updArankssql = updArankssql & " '" & pv_nat2_quota & "' "
	 	Else
	 		updArankssql = updArankssql & "0 "
	 	End If
		updArankssql = updArankssql & ", candrank_c3_" & session("template_org_code") &  "= "
		 If len(pv_nat3_quota) then
	 		updArankssql = updArankssql & " '" & pv_nat3_quota & "' "
	 	Else
	 		updArankssql = updArankssql & "0 "
		end if
	end if
	updArankssql = updArankssql & ", upd_d = getdate() WHERE cand_id_c = " & session("RSYS_EVAL")
	 ' response.write "T <BR>" & updArankssql

  obj_db_select_CmdRank.CommandText = updArankssql
  'set updAranks =rsys_db.execute(updArankssql)
  set updAranks =obj_db_select_CmdRank.execute()


'  <!------------------- INSERT HISTORY OF CHANGES ON NATIONALITY 1 , internal cands not applicable ------------------------------>
' NEEDNEEDNEED verify that this is the right setting for previous values in database when changed
 If len(Request.form("prevnat1")) then
 '<<--Modified by Interface on 05/01/2007
	'06 AUG 09 LJL added check to make sure not the same as previous
	 if request.form("prevnat1") <> pv_nat1_country then
  		set obj_logs_CmdIII = server.CreateObject("adodb.command")
  		obj_logs_CmdIII.ActiveConnection = rsys_logs
  goeditsql = "INSERT INTO tx_rsys_log_candhistory_" & session("template_org_code") & "_" & currentYear &  " (cand_id_c, candhistory_type_c, candhistory_value_prev_t, candhistory_value_new_t, candhistory_ip_c, candhistory_type_t) " &_
  "VALUES (?,'NA',?,?,?, 'NAT 1' )"
  obj_logs_CmdIII.CommandText = goeditsql
  		Set goedit = obj_logs_CmdIII.Execute(,Array(session("RSYS_EVAL"),request.form("prevnat1"),pv_nat1_country,left(request.servervariables("REMOTE_ADDR"), 25)))
  		'-->>
	End if

End If




'  <!--------------------- ADD ORG INFO ------------------------------->
If session("template_org_code") <> 1000 then
'  <!------------------- INSERT HISTORY OF CHANGES ON NATIONALITY 2 ------------------------------>
If len(Request.form("prevnat2")) then
	'06 AUG 09 LJL added check to make sure not the same as previous
	 if request.form("prevnat2") <> pv_nat2_country then
'<<--Modified by Interface on 05/01/2007
  set obj_logs_CmdIV = server.CreateObject("adodb.command")
  obj_logs_CmdIV.ActiveConnection = rsys_logs
  goeditsql = "INSERT INTO tx_rsys_log_candhistory_" & session("template_org_code") & "_" & currentYear &  "(cand_id_c, candhistory_type_c, candhistory_value_prev_t, candhistory_value_new_t, candhistory_ip_c, candhistory_type_t) " &_
  "VALUES ( ?, 'NA', ?,?, ?, 'NAT 2' )"
  obj_logs_CmdIV.CommandText = goeditsql
    if pv_nat2_country = "" then
	pv_nat2_country = "-"
  else
	pv_nat2_country = pv_nat2_country
  end if
  Set goedit = obj_logs_CmdIV.Execute(,Array(session("RSYS_EVAL"),request.form("prevnat2"),pv_nat2_country,left(request.servervariables("REMOTE_ADDR"), 25)))
  '-->>
	End if
End If


'  <!------------------- INSERT HISTORY OF CHANGES ON NATIONALITY 3 ------------------------------>
If len(Request.form("prevnat3")) then
	'06 AUG 09 LJL added check to make sure not the same as previous
	 if request.form("prevnat3") <> pv_nat3_country then
'<<--Modified by Interface on 05/01/2007
  set obj_logs_CmdV = server.CreateObject("adodb.command")
  obj_logs_CmdV.ActiveConnection = rsys_logs
  goeditsql = "INSERT INTO tx_rsys_log_candhistory_" & session("template_org_code")  & "_" & currentYear &  "(cand_id_c, candhistory_type_c, candhistory_value_prev_t, candhistory_value_new_t, candhistory_ip_c, candhistory_type_t) " &_
  "VALUES ( ?, 'NA', ?, ?, ?, 'NAT 3' )"
  obj_logs_CmdV.CommandText = goeditsql
    if pv_nat3_country = "" then
	pv_nat3_country = "-"
  else
	pv_nat3_country = pv_nat3_country
  end if
  Set goedit = obj_logs_CmdV.Execute(,Array(session("RSYS_EVAL"),request.form("prevnat3"),pv_nat3_country,left(request.servervariables("REMOTE_ADDR"), 25)))
 '-->>
	End If
	End if

End If


  '<!------------------- INSERT HISTORY OF CHANGES ON NATIONALITY IN PROCESS OF CHANGE (4) ------------------------------>
If len(Request.form("prevnat4")) then
'<<--Modified by Interface on 05/01/2007
  set obj_logs_CmdVI = server.CreateObject("adodb.command")
  obj_logs_CmdVI.ActiveConnection = rsys_logs
  goeditsql = "INSERT INTO tx_rsys_log_candhistory_" & session("template_org_code")  & "_" & currentYear &  "(cand_id_c, candhistory_type_c, candhistory_value_prev_t, candhistory_value_new_t, candhistory_ip_c, candhistory_type_t) VALUES ( ?, 'NA', 'PREVNAT4', ?, ?,'NAT IN PROCESS' )"
  obj_logs_CmdVI.CommandText = goeditsql
  Set goedit = obj_logs_CmdVI.Execute(,Array(session("RSYS_EVAL"),request.form("prevnat4"),left(request.servervariables("REMOTE_ADDR"), 25)))
'-->>
End If


set obj_db_select_CmdEdit = server.CreateObject("adodb.command")
obj_db_select_CmdEdit.ActiveConnection = rsys_db

 goeditsql = "UPDATE tx_rsys_candedit SET upd_d = getdate(), editA = " & request.form("editA") & ", "
 if isdate(checkdate) then
	goeditsql = goeditsql & " cand_thisorg_start_d_" & session("template_org_code") & " = '" & checkdate & "', "
else
	goeditsql = goeditsql & " cand_thisorg_start_d_" & session("template_org_code") & " = NULL, "
end if
 if isdate(checkdate2) then
	goeditsql = goeditsql & " cand_thisorg_end_d_" & session("template_org_code") & " = '" & checkdate2 & "', "
else
	goeditsql = goeditsql & " cand_thisorg_end_d_" & session("template_org_code") & " = NULL, "
end if
	goeditsql = goeditsql & " editA_d = getdate() WHERE cand_id_c = " & session("RSYS_EVAL")
	' response.write goeditsql

obj_db_select_CmdEdit.CommandText = goeditsql
'set goedit =rsys_db.execute(goeditsql)
set goedit =obj_db_select_CmdEdit.execute()


' WHO specific check if staff, then must indicate in international employment
if session("template_org_code") = 1000 OR session("template_org_code") = 1500 OR session("template_org_code") = 1400 then
	If request.form("thisorg_short") = "1" AND chkA("cand_io_i") = 0 then
		session("OKF") = 0
  		pv_warning = pv_warning & "<font  color='Maroon'><strong>You have indicated that you are currently employed at " & session("template_org_name") & ", but have not indicated it in International Employment section of your profile.  Please do so.<br><br>You may return to the profile by clicking <a href='AppF-edit.asp'>here</a>."
  		'response.end
  End If


  end if



  if (session("template_org_code") = 1000 OR session("template_org_code") = 1500) AND len(pv_thisorg_staffno) then
'<<--Modified by Interface on 05/01/2007
  '08/05/2009 DD , declaration added.

'06 AUG 09 LJL revised for int change logging
    chkSTAFFsql = " SELECT s.email, c.cand_email_t, s.SID, c.honor_id_c, c.cand_lnam_t, c.cand_fnam_t, s.LastName, s.FirstName, c.staff_nbr_"& session("template_org_code") &" AS thisorg_staffno FROM dbo.td_rsys_cand c LEFT OUTER JOIN dbo.v_staff_list_"& session("template_org_code") &" s ON c.staff_nbr_"& session("template_org_code") &" = s.SID COLLATE SQL_Latin1_General_CP850_CI_AI WHERE s.SID = ? AND c.cand_id_c = ?"

    'SELECT SID, email, lastname, firstname, staff_nbr_" & session("template_org_code") & " AS thisorg_staffnumber, SID AS thisorg_isstaff FROM v_staff_list_" & session("template_org_code") & " WHERE NOT SID IS NULL AND SID = ?"

'  chkSTAFFsql = " SELECT SID, email, lastname, firstname FROM v_staff_list_" & session("template_org_code") WHERE NOT SID IS NULL AND SID = ?"
  obj_db_select_Cmd.CommandText = chkSTAFFsql
  Set chkSTAFF = obj_db_select_Cmd.Execute(,Array(pv_thisorg_staffno,Trim(session("RSYS_EVAL"))))
 '-->>
	if chkSTAFF.eof = false then

		if len(chkSTAFF("SID")) then
	  		if UCASE(chkSTAFF("cand_lnam_t")) = UCASE(chkSTAFF("lastname")) then
   				 ' response.write "T CCC22<br>"
  				session("CLI_INTERNAL_STAFF") = 1
				pv_isSTAFF = 1
				pv_isSTAFFNUMBER = 1
  			else
   		 		' response.write "T CCC33<br>"
  				session("CLI_INTERNAL_STAFF") = 2
				pv_isSTAFF = 2
				pv_isSTAFFNUMBER = 0
  			end if
		else
			pv_isSTAFF = 2
			pv_isSTAFFNUMBER = 0
		end if

	else
		session("CLI_INTERNAL_STAFF") = 0
		pv_isSTAFF = 0
		pv_isSTAFFNUMBER = 0
	end if
	
	'response.write "<br><br>T PVIS: " & pv_isSTAFF
	'response.write "<br><Br>T PVSTNUM: " & pv_isSTAFFNUMBER

		'if pv_isSTAFF = 1 then
		'	dim  obj_db_select_CmdCW, chkWHOSTAFFsql, chkWHOSTAFF
  		'	set obj_db_select_CmdCW = server.CreateObject("adodb.command")
  		'	obj_db_select_CmdCW.ActiveConnection = rsys_db_select
  		'	chkWHOSTAFFsql = " SELECT cand_lnam_t, staff_nbr_" & session("template_org_code") & " AS thisorg_staffnumber FROM td_rsys_cand WHERE cand_id_c = ?"
  		'	obj_db_select_CmdCW.CommandText = chkWHOSTAFFsql
  		'	Set chkWHOSTAFF = obj_db_select_CmdCW.Execute(,Array(Trim(session("RSYS_EVAL"))))

	  	'	if UCASE(chkWHOSTAFF("cand_lnam_t")) = UCASE(chkSTAFF("lastname")) then
   		'		 ' response.write "T CCC22<br>"
  		'		session("CLI_INTERNAL_STAFF") = 1
  		'	else
   		 '		' response.write "T CCC33<br>"
  		'		session("CLI_INTERNAL_STAFF") = 2
		'		pv_isSTAFF = 2
		'		pv_isSTAFFNUMBER = 0
  		'	end if
  		'	set obj_db_select_CmdCW = nothing

  		'end if
end if





else

'06 AUG 09 LJL Added if they ACCESSED THEIR PROFILE PORTION -------------------------------------------------
dim obj_logs_CmdPA, PACCsql, PACC
	set obj_logs_CmdPA = server.CreateObject("adodb.command")
	obj_logs_CmdPA.ActiveConnection = rsys_logs
	If session("RSYSUSER") <> "" then
		PACCsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  " (candupd_ip_c, cand_id_c, candupd_page_t,user_id_t,candupd_site_c) " &_
		"VALUES (?, ?, 'A-Viewed', ?,?)"
		obj_logs_CmdPA.CommandText = PACCsql
		Set PACC = obj_logs_CmdPA.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),session("RSYSUSER"),left(session("template_org_name"),20)))
	Else
		PACCsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  " (candupd_ip_c, cand_id_c, candupd_page_t,candupd_site_c) " &_
		"VALUES (?, ?, 'A-Viewed',?)"
		obj_logs_CmdPA.CommandText = PACCsql
		Set PACC = obj_logs_CmdPA.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),left(session("template_org_name"),20)))
	End if
' -------------------------------------------------------------------------------------------------------------------------------------------

'IF NOT EDITING, CHECK IF PERSON IS STAFF

'CHECK IF STAFF IF WHO OR UNAIDS
if (session("template_org_code") = 1000 OR session("template_org_code") = 1500) then
'<<--Modified by Interface on 05/01/2007
  '08/05/2009 DD , declaration added.

'06 AUG 09 LJL revised for int change logging
    chkSTAFFsql = " SELECT s.email, c.cand_email_t, s.SID, c.honor_id_c, c.cand_lnam_t, c.cand_fnam_t, s.LastName, s.FirstName, c.staff_nbr_"& session("template_org_code") &" AS thisorg_staffno FROM dbo.td_rsys_cand c LEFT OUTER JOIN dbo.v_staff_list_"& session("template_org_code") &" s ON c.staff_nbr_"& session("template_org_code") &" = s.SID COLLATE SQL_Latin1_General_CP850_CI_AI WHERE c.cand_id_c = ?"
    'SELECT SID, email, lastname, firstname, staff_nbr_" & session("template_org_code") & " AS thisorg_staffnumber, SID AS thisorg_isstaff FROM v_staff_list_" & session("template_org_code") & " WHERE NOT SID IS NULL AND SID = ?"

'  chkSTAFFsql = " SELECT SID, email, lastname, firstname FROM v_staff_list_" & session("template_org_code") WHERE NOT SID IS NULL AND SID = ?"
  obj_db_select_Cmd.CommandText = chkSTAFFsql
  Set chkSTAFF = obj_db_select_Cmd.Execute(,Array(Trim(session("RSYS_EVAL"))))
 '-->>
	if chkSTAFF.eof = false then

	if len(chkSTAFF("thisorg_staffno")) then

		if len(chkSTAFF("SID")) then
	  		if UCASE(chkSTAFF("cand_lnam_t")) = UCASE(chkSTAFF("lastname")) then
   				 ' response.write "T CCC22<br>"
  				session("CLI_INTERNAL_STAFF") = 1
				pv_isSTAFF = 1
				pv_isSTAFFNUMBER = 1
  			else
   		 		' response.write "T CCC33<br>"
  				session("CLI_INTERNAL_STAFF") = 2
				pv_isSTAFF = 2
				pv_isSTAFFNUMBER = 0
  			end if
		else
			pv_isSTAFF = 2
			pv_isSTAFFNUMBER = 0
		end if

	else
		session("CLI_INTERNAL_STAFF") = 0
		pv_isSTAFF = 0
		pv_isSTAFFNUMBER = 0
	end if

	else
		session("CLI_INTERNAL_STAFF") = 0
		pv_isSTAFF = 0
		pv_isSTAFFNUMBER = 0

	end if

end if





End If

'******************************
' END A EDIT
'******************************
'08 JUL 06 LJL added UPDATE PHRASE HERE%>
<!-- #include file="../edit/ejobs-updates.asp"-->
<%


 If session("lng") = "fr" then
	ctrylang = "country_french_name"
elseIf session("lng") = "es" then
	ctrylang = "country_spanish_name"
Else
	ctrylang = "country_english_name"
End If



set obj_db_select_CmdII = server.CreateObject("adodb.command")
obj_db_select_CmdII.ActiveConnection = rsys_db_select


gAGEsql = "			SELECT 			rtrim(admitem_default_" & session("template_org_code") & ") AS admitem 			FROM td_rsys_admitem 			WHERE admitem_ident_c = 'MAXIMUM-AGE' 		"
obj_db_select_CmdII.CommandText = gAGEsql
'set gAGE = rsys_db_select.execute(gAGEsql)
set gAGE = obj_db_select_CmdII.execute()

if gAGE.eof = true then
	response.write	("Error.  Alert Admin that minimum values not set - New Registration<br><br>")
    response.end
end if


set obj_db_select_CmdIII = server.CreateObject("adodb.command")
obj_db_select_CmdIII.ActiveConnection = rsys_db_select

 gAGE2sql = " SELECT rtrim(admitem_default_" & session("template_org_code") & ") AS admitem 	FROM td_rsys_admitem WHERE admitem_ident_c = 'MAXIMUM-AGE-INTERN' 		"
obj_db_select_CmdIII.CommandText = gAGE2sql
'set gAGE2 = rsys_db_select.execute(gAGE2sql)
set gAGE2 = obj_db_select_CmdIII.execute()

'<!----------------- moved here from additional info (j) 28 APR 04 by pigueta --------------------->

set obj_db_select_CmdIV = server.CreateObject("adodb.command")
obj_db_select_CmdIV.ActiveConnection = rsys_db_select

'03 NOV 12 LJL order of webfamiliar
getFamiliarssql = " SELECT webfamiliar_dsc_"& session("lng") &"_t as familiardsc, webfamiliar_id_c FROM core_webfamiliar 		WHERE webfamiliar_inactind_i <> 1 AND webfamiliar_thisorg_" & session("template_org_code") &  " = 1 ORDER BY 1 "
'webfamiliar_rank_i "

'response.write "<br>FAMIL:<BR>" & getFamiliarssql

obj_db_select_CmdIV.CommandText = getFamiliarssql
'set getFamiliars =rsys_db_select.execute(getFamiliarssql)
set getFamiliars =obj_db_select_CmdIV.execute()

set obj_db_select_CmdV = server.CreateObject("adodb.command")
obj_db_select_CmdV.ActiveConnection = rsys_db_select

'27 APR 11 LJL revised marital to be org reflective
GETMARITALsql = "SELECT marital_id_c, marital_"& session("lng") &"_dsc_" & session("template_org_code") & " AS maritaldsc FROM core_maritalt WHERE (NOT marital_" & session("lng") & "_dsc_" & session("template_org_code") & " IS NULL) AND marital_thisorg_" & session("template_org_code") & "= 1 ORDER BY 2 "
obj_db_select_CmdV.CommandText = GETMARITALsql
'response.write "<br><br>T MARITAL " & getMARITALsql
'set GETMARITAL =rsys_db_select.execute(GETMARITALsql)
set GETMARITAL =obj_db_select_CmdV.execute()

set obj_db_select_CmdVI = server.CreateObject("adodb.command")
obj_db_select_CmdVI.ActiveConnection = rsys_db_select

GETHONORsql = "SELECT honor_"& session("lng") &"_dsc_t AS honordsc, honor_id_c FROM core_honorific WHERE honor_thisorg_" & session("template_org_code") & "= 1 AND honor_public_i = 1 AND NOT honor_"& session("lng") & "_dsc_t IS NULL ORDER BY 1"
obj_db_select_CmdVI.CommandText = GETHONORsql
'set GETHONOR =rsys_db_select.execute(GETHONORsql)
set GETHONOR =obj_db_select_CmdVI.execute()

set obj_db_select_CmdVII = server.CreateObject("adodb.command")
obj_db_select_CmdVII.ActiveConnection = rsys_db_select

getmonthssql = "SELECT month_id_c, month_int_c, month_dsc_"& session("lng") &"_t AS monthname 	FROM core_montht 	WHERE month_id_c < 13 AND month_id_c > 0 	ORDER BY 1"
obj_db_select_CmdVII.CommandText = getmonthssql
'set getmonths =rsys_db_select.execute(getmonthssql)
set getmonths =obj_db_select_CmdVII.execute()

'04 OCT 10 LJL added public PHF viewing indicator
set obj_db_select_CmdVIII = server.CreateObject("adodb.command")
obj_db_select_CmdVIII.ActiveConnection = rsys_db_select

GETCONTRACTLENsql = "SELECT contractlen_id_c, contractlen_dsc_en_t, contractlen_dsc_"& session("lng") &"_t AS contractlendsc FROM tr_rsys_contractlen WHERE contractlen_thisorg_" & session("template_org_code") & " = 1 AND contractlen_pubphf_i = 1 ORDER BY 2"
obj_db_select_CmdVIII.CommandText = GETCONTRACTLENsql
'set GETCONTRACTLEN =rsys_db_select.execute(GETCONTRACTLENsql)
set GETCONTRACTLEN =obj_db_select_CmdVIII.execute()


' GET PASSPORT DETAILS IF NEEDED FOR IFRC 7000
if session("template_org_code") = 7000 then
'<<--Modified by Interface on 05/01/2007
	set obj_db_select_Cmd4 = server.CreateObject("adodb.command")
	obj_db_select_Cmd4.ActiveConnection = rsys_db

	JAPINFO4sql = "	SELECT candmisc_passport_c, candmisc_passport_place_c, candmisc_passport_issue_d, candmisc_passport_valid_d FROM tx_rsys_candmisc WHERE cand_id_c = ? "
	obj_db_select_Cmd4.CommandText = JAPINFO4sql
	Set JAPINFO4 = obj_db_select_Cmd4.Execute(,Array(session("RSYS_EVAL")))
'-->>
' GET APPTYPE DETAILS IF NEEDED FOR WTO 3000 (check if intern or not)
elseif session("template_org_code") = 3000 then
	set obj_db_select_Cmd4 = server.CreateObject("adodb.command")
	obj_db_select_Cmd4.ActiveConnection = rsys_db

'<<--Modified by Interface on 05/01/2007
	JAPINFO4sql = "SELECT candmisc_apptype_" & session("template_org_code") & " AS apptype FROM tx_rsys_candmisc WHERE cand_id_c = ? "
	obj_db_select_Cmd4.CommandText = JAPINFO4sql
	Set JAPINFO4 = obj_db_select_Cmd4.Execute(,Array(session("RSYS_EVAL")))
'-->>
end if

	set obj_db_select_Cmd3 = server.CreateObject("adodb.command")
	obj_db_select_Cmd3.ActiveConnection = rsys_db

'response.write "CLI_STAFF: " & session("CLI_INTERNAL_STAFF")

'		        <!-------- internal change here ------------------------> WHO ONLY INTERNAL
dim isIntSqlExecuted
isIntSqlExecuted = false

'response.write "TEST INTSTAFF : " & session("CLI_INTERNAL_STAFF")
'response.end


' ************************************************************************************************************************************
' WHO / UNAIDS - set the type of internal staff that they are
' ************************************************************************************************************************************


'response.write "<br>INTTY: " & session("CLI_INTERNAL_STAFF")
'response.write "<br>pv_isSTAFF: " & pv_isSTAFF

'pv_isSTAFFNUMBER = 0
  set obj_db_select_CmdSN = server.CreateObject("adodb.command")
  obj_db_select_CmdSN.ActiveConnection = rsys_db_select

If (session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500) AND pv_isSTAFF = "1" then

	'response.write "111<br>"
	'<<--Modified by Interface on 05/01/2007
	'28 APR 11 LJL revised marital desc
	JAPINFO3sql = "	SELECT c.cand_thisorg_jobtype_i_"& session("template_org_code") &" AS thisorg_type, c.honor_id_c, c.cand_lnam_t, c.cand_fnam_t, c.cand_mnam_t, c.cand_onam_t, c.cand_gnd_i, c.cand_maiden_t, c.cand_bthp_t, s.DOB AS cand_bth_d, c.cand_mar_st_c, rtrim(c.cand_nat_c) AS cand_nat_c, RTRIM(c.cand_nat2_c) AS cand_nat2_c, rtrim(c.cand_nat3_c) AS cand_nat3_c, c.cand_pnat_i, c.cand_expl_t, c.cand_newnat_c, c.cand_newnat_d, c.cand_thisorg_short_i_"& session("template_org_code") &" AS thisorg_short, c.cand_thisorg_prev_i_"& session("template_org_code") &" AS thisorg_prev, c.cand_thisorg_refn_c_"& session("template_org_code") &" AS thisorg_refn, 	c.cand_fil_t, cq.countryquota_rank_i AS cq, c.webfamiliar_id_c, c.webfamiliar_refer_t, s.LastName, s.FirstName, s.RoomNo, s.TelephoneNo, s.Phone, s.Office, s.UnitAcro, rtrim(s.Natlty) AS Natlty, s.DeptAcro, s.DutyStation, s.sex_code, s.unit_id, s.staff_type, s.contract_type, s.contract_start_date, s.staff_category, s.staff_grade, s.classified_grade_of_post, s.Email, s.StaffNo, s.Salutation, s.OtherName, s.SID, s.marital_" & session("lng") & "_dsc_t AS maritaldsc, s.marital_id_c as marital_id, cty.country_english_name AS internal_nat, 	c.staff_nbr_"& session("template_org_code") &" AS thisorg_staffno, s.contract_end_date FROM dbo.v_country_quota_check_"& session("template_org_code") &" cq RIGHT OUTER JOIN dbo.v_country_list_"& session("template_org_code") &" cty ON cq.who_country_code COLLATE SQL_Latin1_General_CP850_CI_AI = cty.who_country_code RIGHT OUTER JOIN dbo.td_rsys_cand c INNER JOIN dbo.v_staff_list_"& session("template_org_code") &" s ON c.staff_nbr_"& session("template_org_code") &" = s.SID COLLATE SQL_Latin1_General_CP850_CI_AI ON cty.who_country_code COLLATE SQL_Latin1_General_CP1_CI_AS = s.Natlty WHERE c.cand_id_c = ? "
	obj_db_select_Cmd3.CommandText = JAPINFO3sql

	'response.write JAPINFO3sql
		Set JAPINFO3 = obj_db_select_Cmd3.Execute(,Array(session("RSYS_EVAL")))
	isIntSqlExecuted = true

	'-->>
Elseif session("template_org_code") = 2500 then
	'response.write "222<br>"
	'<<--Modified by Interface on 05/01/2007
	JAPINFO3sql = "SELECT c.cand_prev_year_c, c.upd_d, c.cand_thisorg_jobtype_i_" & session("template_org_code") &  " AS thisorg_type, c.staff_nbr_" & session("template_org_code") &  " AS thisorg_staffno, c.honor_id_c, c.cand_lnam_t, c.cand_fnam_t, c.cand_mnam_t, c.cand_onam_t, c.cand_gnd_i, c.cand_maiden_t, c.cand_bthp_t, c.cand_bth_d, c.marital_id_" & session("template_org_code") & " AS marital_id, c.cand_mar_st_c, rtrim(c.cand_nat_c) AS cand_nat_c, RTRIM(c.cand_nat2_c) AS cand_nat2_c, rtrim(c.cand_nat3_c) AS cand_nat3_c, c.cand_pnat_i, c.cand_expl_t, c.webfamiliar_id_c, c.webfamiliar_refer_t, c.cand_newnat_c, c.cand_newnat_d, c.cand_thisorg_short_i_"& session("template_org_code") &" AS thisorg_short, c.cand_thisorg_prev_i_"& session("template_org_code") &" AS thisorg_prev, c.cand_thisorg_refn_c_"& session("template_org_code") &" AS thisorg_refn, c.cand_fil_t, cty.c_" & session("lng") & "_name AS ctyname FROM dbo.td_rsys_cand c INNER JOIN dbo.v_country_list_"& session("template_org_code") &" cty ON c.cand_nat_c = cty.who_country_code WHERE (c.cand_id_c = ?)"
	obj_db_select_Cmd3.CommandText = JAPINFO3sql
	Set JAPINFO3 = obj_db_select_Cmd3.Execute(,Array(session("RSYS_EVAL")))
	'-->>
	if JAPINFO3.eof = false then

	else
		response.write "The staff number is not correct or there is a problem with your data in the system.<br><br>Please contact Tech Support (msg:A 971)"
		'response.end
	end if
Else


	'response.write "333<br>"
	'<<--Modified by Interface on 05/01/2007
	JAPINFO3sql = "SELECT cand_thisorg_short_i_"& session("template_org_code") &" AS thisorg_short, cand_thisorg_prev_i_"& session("template_org_code") &" AS thisorg_prev, cand_prev_year_c, upd_d, cand_thisorg_jobtype_i_" & session("template_org_code") &  " AS thisorg_type, 		staff_nbr_"& session("template_org_code") &" AS thisorg_staffno, honor_id_c, cand_lnam_t, cand_fnam_t, cand_mnam_t, cand_onam_t, cand_gnd_i, cand_maiden_t, cand_bthp_t, cand_bth_d, marital_id_" & session("template_org_code") & " AS marital_id, cand_mar_st_c, rtrim(cand_nat_c) AS cand_nat_c, rtrim(cand_nat2_c) AS cand_nat2_c, rtrim(cand_nat3_c) AS cand_nat3_c, cand_pnat_i, cand_expl_t, webfamiliar_id_c, webfamiliar_refer_t, cand_newnat_c, cand_newnat_d, 	cand_thisorg_short_i_"& session("template_org_code") &" AS thisorg_short, cand_thisorg_prev_i_"& session("template_org_code") &" AS thisorg_prev, cand_thisorg_refn_c_"& session("template_org_code") &" AS thisorg_refn, cand_fil_t FROM td_rsys_cand WHERE cand_id_c = ? "
	obj_db_select_Cmd3.CommandText = JAPINFO3sql
	'response.write "<BR><BR>JAP3 " & JAPINFO3sql
	Set JAPINFO3 = obj_db_select_Cmd3.Execute(,Array(session("RSYS_EVAL")))
	'-->>
	if JAPINFO3.eof = false then

	else
		response.write "The staff number is not correct or there is a problem with your data in the system.<br><br>Please contact Tech Support (msg:A 971)"
		'response.end
	end if


End If

	'response.write "TT JAP3:" & JAPINFO3sql & "| ID:" & session("RSYS_EVAL")
	'response.end


	set obj_db_select_Cmd2 = server.CreateObject("adodb.command")
	obj_db_select_Cmd2.ActiveConnection = rsys_db_select

'<<--Modified by Interface on 05/01/2007
JAPINFO2sql = "	SELECT cand_thisorg_start_d_" & session("template_org_code") &  " AS thisorg_start, cand_thisorg_end_d_" & session("template_org_code") &  " AS thisorg_end, editA FROM tx_rsys_candedit WHERE cand_id_c = ?"
obj_db_select_Cmd2.CommandText = JAPINFO2sql
Set JAPINFO2 = obj_db_select_Cmd2.Execute(,Array(session("RSYS_EVAL")))
'-->>

set obj_db_select_CmdIX = server.CreateObject("adodb.command")
obj_db_select_CmdIX.ActiveConnection = rsys_db_select

set obj_db_select_CmdX = server.CreateObject("adodb.command")
obj_db_select_CmdX.ActiveConnection = rsys_db_select


If session("template_org_code") = "3000" then
	'<!---------------------------------- ADD ORG INFO 3000 may need to adjust for interns and consultants per WTO not sure yet 12 APR 05 -------------------->
	 If JAPINFO4("apptype") = "15" then
		'<!---------------------- INTERNS ONLY MEMBER AND ACCEDING COUNTRIES ------------------------------>
		GETNATsql = "SELECT TOP 100 PERCENT rtrim(cq.who_country_code) AS code, cq.c_"& session("lng") &"_name AS ctyname, cq.countryquota_rank_i AS cq 	FROM dbo.v_country_quota_check_"& session("template_org_code") &" cq INNER JOIN dbo.tr_rsys_country cty ON cq.who_country_code = cty.cty_id_c COLLATE SQL_Latin1_General_CP1_CI_AS 	WHERE (cty.country_member_"& session("template_org_code") &" = 1 OR cty.country_member_"& session("template_org_code") &" = 2) AND (len(cq.c_"& session("lng") &"_name) > 0) 	AND len(c_"& session("lng") & "_name) > 0 ORDER BY 2 	"
		obj_db_select_CmdIX.CommandText = GETNATsql
		'set GETNAT =rsys_db_select.execute(GETNATsql)
		set GETNAT =obj_db_select_CmdIX.execute()
	elseif JAPINFO4("apptype") = "20" then
		'<!---------------------- CONSULTANTS ONLY MEMBER AND ACCEDING COUNTRIES ------------------------------>
		GETNATsql = "SELECT TOP 100 PERCENT rtrim(cq.who_country_code) AS code, cq.c_"& session("lng") &"_name AS ctyname, cq.countryquota_rank_i AS cq 	FROM dbo.v_country_quota_check_"& session("template_org_code") &" cq INNER JOIN dbo.tr_rsys_country cty ON cq.who_country_code = cty.cty_id_c COLLATE SQL_Latin1_General_CP1_CI_AS 	WHERE (cty.country_member_"& session("template_org_code") &" = 1 OR cty.country_member_ "& session("template_org_code") &" = 2) AND (len(cq.c_"& session("lng") &"_name) > 0)	AND len(c_"& session("lng") & "_name) > 0 ORDER BY 2 	"
		obj_db_select_CmdIX.CommandText = GETNATsql
		'set GETNAT =rsys_db_select.execute(GETNATsql)
		set GETNAT =obj_db_select_CmdIX.execute()
	Else
		'<!---------------------- APPLICANTS TO POSTS  ONLY MEMBER COUNTRIES ------------------------------>
		GETNATsql = "SELECT TOP 100 PERCENT rtrim(cq.who_country_code) AS code, cq.c_"& session("lng") &"_name AS ctyname, cq.countryquota_rank_i AS cq 	FROM dbo.v_country_quota_check_"& session("template_org_code") &" cq INNER JOIN dbo.tr_rsys_country cty ON cq.who_country_code = cty.cty_id_c COLLATE SQL_Latin1_General_CP1_CI_AS 	WHERE (cty.country_member_"& session("template_org_code") &" = 1) 	AND (len(cq.c_"& session("lng") &"_name) > 0)	AND len(c_"& session("lng") & "_name) > 0 ORDER BY 2 	"
		obj_db_select_CmdIX.CommandText = GETNATsql
		'set GETNAT =rsys_db_select.execute(GETNATsql)
		set GETNAT =obj_db_select_CmdIX.execute()
	End If
		'<!---------------------- APPLICANTS TO POSTS  NAT2 and NAT3 - All countries allowed ------------------------------>
		GETNAT2sql = "SELECT TOP 100 PERCENT rtrim(cq.who_country_code) AS code, cq.c_"& session("lng") &"_name AS ctyname, cq.countryquota_rank_i AS cq 	FROM dbo.v_country_quota_check_"& session("template_org_code") &" cq INNER JOIN dbo.tr_rsys_country cty ON cq.who_country_code = cty.cty_id_c COLLATE SQL_Latin1_General_CP1_CI_AS 	WHERE (len(cq.c_"& session("lng") &"_name) > 0) 	AND len(c_"& session("lng") & "_name) > 0  ORDER BY 2 	"
		'response.write GETNAT2sql
		obj_db_select_CmdIX.CommandText = GETNAT2sql
		'set GETNAT2 =rsys_db_select.execute(GETNAT2sql)
		set GETNAT2 =obj_db_select_CmdIX.execute()

elseIf session("template_org_code") = "7000" then
'20 DEC 14 LJL added  AND countryquota_rank_i > 0 to the quota listing query
' 7000 IFRC USES NATIONALITY AND NOT COUNTRY NAME IN THE NAT DROPDOWN
	GETNATsql = " SELECT who_country_code AS code, c_"& session("lng") &"_nat AS ctyname, countryquota_rank_i AS cq FROM v_country_quota_check_" & session("template_org_code") &  " WHERE len(c_"& session("lng") & "_nat) > 0 AND len(c_"& session("lng") & "_name) > 0  AND countryquota_rank_i > 0 ORDER BY 2"
	obj_db_select_CmdIX.CommandText = GETNATsql
	'set GETNAT =rsys_db_select.execute(GETNATsql)
	set GETNAT =obj_db_select_CmdIX.execute()

	'response.write "TEST NAT:" & GETNATsql

	'<!---------------------- APPLICANTS TO POSTS  NAT2 and NAT3 - All countries allowed ------------------------------>
'20 DEC 14 LJL added  AND countryquota_rank_i > 0 to the quota listing query
	GETNAT2sql = "		SELECT who_country_code AS code, c_"& session("lng") &"_nat AS ctyname, countryquota_rank_i AS cq FROM v_country_quota_check_" & session("template_org_code") &  " WHERE len(c_"& session("lng") & "_nat) > 0  AND len(c_"& session("lng") & "_name) > 0 AND countryquota_rank_i > 0 ORDER BY 2"
	obj_db_select_CmdX.CommandText = GETNAT2sql
	'set GETNAT2 =rsys_db_select.execute(GETNAT2sql)
	set GETNAT2 =obj_db_select_CmdX.execute()

' UNESCO 2500
'05 JUN 10 LJL added check on language name of country to make the list tighter
Elseif session("template_org_code") = 2500 AND (instr(session("rsysuser"), "ADM") = false) then
		set obj_db_select_CmdN = server.CreateObject("adodb.command")
		obj_db_select_CmdN.ActiveConnection = rsys_db

	'<!---------------------- APPLICANTS TO POSTS  NAT 1 ------------------------------>
	'<<--Modified by Interface on 05/01/2007
'20 DEC 14 LJL added  AND countryquota_rank_i > 0 to the quota listing query
	GETNATsql = "SELECT rtrim(who_country_code) AS code, c_"& session("lng") &"_name AS ctyname, countryquota_rank_i AS cq FROM v_country_quota_check_" & session("template_org_code") &  " WHERE len(c_"& session("lng") & "_name) > 0 AND len(c_"& session("lng") & "_nat) > 0  AND len(c_"& session("lng") & "_name) > 0 AND countryquota_rank_i > 0 AND who_country_code = ? "
	obj_db_select_CmdN.CommandText = GETNATsql
	Set GETNAT = obj_db_select_CmdN.Execute(,Array(JAPINFO3("cand_nat_c")))
	'-->>
	'<!---------------------- APPLICANTS TO POSTS  NAT2 and NAT3 - All countries allowed ------------------------------>
'20 DEC 14 LJL added  AND countryquota_rank_i > 0 to the quota listing query
	GETNAT2sql = "		SELECT rtrim(who_country_code) AS code, c_"& session("lng") &"_name AS ctyname, countryquota_rank_i AS cq FROM v_country_quota_check_" & session("template_org_code") &  " WHERE len(c_"& session("lng") & "_name) > 0 AND len(c_"& session("lng") & "_nat) > 0  AND len(c_"& session("lng") & "_name) > 0  AND countryquota_rank_i > 0 ORDER BY 2"
	obj_db_select_CmdIX.CommandText = GETNAT2sql
	'set GETNAT2 =rsys_db_select.execute(GETNAT2sql)
	set GETNAT2 =obj_db_select_CmdIX.execute()

'05 OCT 10 LJL only member states for nat list for ITU
Elseif session("template_org_code") = 2400 then
	'<!---------------------- APPLICANTS TO POSTS  NAT 1 ------------------------------>
'20 DEC 14 LJL added  AND countryquota_rank_i > 0 to the quota listing query
	GETNATsql = " SELECT rtrim(who_country_code) AS code, c_"& session("lng") &"_name AS ctyname, countryquota_rank_i AS cq FROM v_country_quota_check_" & session("template_org_code") &  " WHERE country_member = 1  AND countryquota_rank_i > 0 ORDER BY 2"
	obj_db_select_CmdIX.CommandText = GETNATsql
	'set GETNAT =rsys_db_select.execute(GETNATsql)
	set GETNAT =obj_db_select_CmdIX.execute()
	'<!---------------------- APPLICANTS TO POSTS  NAT2 and NAT3 - All countries allowed ------------------------------>

'20 DEC 14 LJL added  AND countryquota_rank_i > 0 to the quota listing query
	GETNAT2sql = " SELECT rtrim(who_country_code) AS code, c_"& session("lng") &"_name AS ctyname, countryquota_rank_i AS cq FROM v_country_quota_check_" & session("template_org_code") &  " WHERE country_member = 1  AND countryquota_rank_i > 0 ORDER BY 2"
		'response.write GETNAT2sql
	obj_db_select_CmdX.CommandText = GETNAT2sql
	'set GETNAT2 =rsys_db_select.execute(GETNAT2sql)
	set GETNAT2 =obj_db_select_CmdX.execute()


' ALL OTHER ORGS
Else
	'<!---------------------- APPLICANTS TO POSTS  NAT 1 ------------------------------>
	GETNATsql = " SELECT rtrim(who_country_code) AS code, c_"& session("lng") &"_name AS ctyname, countryquota_rank_i AS cq FROM v_country_quota_check_" & session("template_org_code") &  " WHERE len(c_"& session("lng") & "_nat) > 0 AND len(c_"& session("lng") & "_name) > 0 ORDER BY 2"
	
	'response.write "<BR>NAT: "& GETNATsql
	obj_db_select_CmdIX.CommandText = GETNATsql
	'set GETNAT =rsys_db_select.execute(GETNATsql)
	set GETNAT =obj_db_select_CmdIX.execute()
	'<!---------------------- APPLICANTS TO POSTS  NAT2 and NAT3 - All countries allowed ------------------------------>
	GETNAT2sql = " SELECT rtrim(who_country_code) AS code, c_"& session("lng") &"_name AS ctyname, countryquota_rank_i AS cq FROM v_country_quota_check_" & session("template_org_code") &  " WHERE len(c_"& session("lng") & "_nat) > 0 AND len(c_"& session("lng") & "_name) > 0  ORDER BY 2"
	'AND countryquota_rank_i > 0 
		'response.write GETNAT2sql
	obj_db_select_CmdX.CommandText = GETNAT2sql
	'set GETNAT2 =rsys_db_select.execute(GETNAT2sql)
	set GETNAT2 =obj_db_select_CmdX.execute()
End If

'response.write getnat2sql


'03 FEB 08 LJL ALREADYOK was not allowing application to vacancy for WHO.  Changed public/hrd-cl-auth.asp to change the checknatok session var properly
session("checknatok") = 0
'01 AUG 07 LJL added NAT CHECK AND SESSION SET FOR WTO
	if session("template_org_code") = "3000" then
		'<!------------------------------ ADD ORG INFO 3000 31 MAY 05 LJL only for WTO, only nationals apply to posts ----------------------------->
		'<!---------------------- GET CAND NATIONALITY TO SEE IF MEMBER STATE -------------------------------------------------->
		'<<--Modified by Interface on 05/01/2007
		set obj_db_select_CmdN = server.CreateObject("adodb.command")
		obj_db_select_CmdN.ActiveConnection = rsys_db

		GETJAPNATsql = "SELECT ct.country_member_" & session("template_org_code") & " AS cty_member, c.cand_nat_c FROM td_rsys_cand c, tr_rsys_country ct WHERE c.cand_id_c = ? AND c.cand_nat_c = ct.cty_id_c "
		obj_db_select_CmdN.CommandText = GETJAPNATsql
		Set GETJAPNAT = obj_db_select_CmdN.Execute(,Array(session("RSYS_EVAL")))
		'-->>
		session("checknatok") = 1
	' 11 NOV 05 LJL CHANGED to the actual country member number
		if GETJAPNAT.eof= false then
			if GETJAPNAT("cty_member") > 0 then
				session("checknatok") = GETJAPNAT("cty_member")
			else
				session("checknatok") = GETJAPNAT("cty_member")
			end if
		else
			session("checknatok") = GETJAPNAT("cty_member")
		end if
	else
	' 11 NOV 05 LJL CHANGED to 0 for all others
		session("checknatok") = 1
	end if

'response.write "<br>CHECKNAT:" & session("CHECKNATOK")

' SET THE PROPER INCLUDE TEXT FOR EITHER INTERN OR REGULAR APPLICANT - MAINLY WTO ISSUE
	if session("rsys_intern") = "1" then
		pv_nattext = gITEXTA("i_67")
		' NORMAL TEXT
		'pv_nattext = gITEXTA("i_12")
	else
		pv_nattext = gITEXTA("i_49")
	end if

'response.write "NATTY:" & session("CHECKNATOK")
	%>
	
	


<SCRIPT language="JavaScript" type="text/javascript" src="../../js/Validation_Script.js"></script>
   
   <!--DATE PICKER CHANGES 11-22-2014-->
    <link rel="stylesheet" href="js/datepickercss.css" />
    <script src="js/jquery-1.8.2.js"></script>
    <script src="js/jquery-ui-date.js"></script> 



    <% if session("lng")="fr" then%>
<script src="js/datepicker-fr.js"></script> 
        <script>
            $(function () {
                $("#datepicker,#datepicker1").datepicker({
                    changeMonth: true,
                    changeYear: true,
                    dateFormat: 'd-M-yy',
                    yearRange: '-125:+0',
                    onChangeMonthYear:function(y, m, i){                                
                    var d = i.selectedDay;
                    $(this).datepicker('setDate', new Date(y, m-1, d));
                }

                }).datepicker($.datepicker.regional["fr"]);
            });
        </script>
<%end if%>
<% if session("lng")="en" then%> 
        <script>
            $(function () {
                $("#datepicker,#datepicker1").datepicker({
                    changeMonth: true,
                    changeYear: true,
                    dateFormat: 'd-M-yy',
                    yearRange: '-125:+0',
                    onChangeMonthYear: function (y, m, i) {
                        var d = i.selectedDay;
                        $(this).datepicker('setDate', new Date(y, m - 1, d));
                    }

                });
            });
        </script>
<%end if%>
 <!--DATE PICKER CHANGES 11-22-2014-->
<script>  
var act_form_object = "";
//Modified on 07/21/2009 DD,To check the valid last name and setting of upper case.
function setUpperCase(lname) {
	return false;
    var retval;
	var str_lname = "";
		str_lname = document.data.cand_lnam_t.value;
		str_lname = str_lname.replace(/^\s+/g, '').replace(/\s+$/g, '');
        if (str_lname == "") {
            alert("<% response.write gITEXTA("i_51")%>");
			setfocus(lname);
            return false;
        }
    retval = ValidateLastName(document.data.cand_lnam_t.value);
    if (retval == false)
    {
        //alert("Family/Last name should contain minimum two alphabatic characters.");
		alert("Family name should contain minimum two alphabatic characters.");
        setfocus(lname);
        return false;
    }
    document.data.cand_lnam_t.value = document.data.cand_lnam_t.value.toUpperCase();
    //return true;
}

$(document).ready(function(){
     if (document.data.thisorg_short.value =="0")
     { document.data.thisorg_type.style.display ='none'; document.getElementById("spancontract").style.display ='none'}
          
           if (document.data.thisorg_short.value =="1")
                 { document.data.thisorg_type.style.display ='block'; document.getElementById("spancontract").style.display ='block' }


                   $("#contract_short").change(function () {

            if (document.data.thisorg_short.value == "0")
            { document.data.thisorg_type.style.display ='none'; document.getElementById("spancontract").style.display ='none' }

            if (document.data.thisorg_short.value == "1")
            { document.data.thisorg_type.style.display ='block'; document.getElementById("spancontract").style.display ='block' }



        });
   
   
});

function DataValidation()
    {
        //var MyReg = /^[\a-zA-ZàèìòùÀÈÌÒÙáéíóúýÁÉÍÓÚÝâêîôûÂÊÎÔÛãñõÃÑÕäëïöüÿÄËÏÖÜŸçÇßØøÅåÆæœ ]+$/i;
        
        var MyReg = /^[A-Za-zÀ-ȕ ];
       
        if (document.data.honor_id_c.value == "") {
            alert("<% response.write gITEXTA("i_74")%>");
			document.data.honor_id_c.focus();
            return false;
        }
		var str_fname = "";
		str_fname = document.data.cand_fnam_t.value;
		
		str_fname = str_fname.replace(/^\s+/g, '').replace(/\s+$/g, '');

           
         if (str_fname.indexOf("''") != -1 || str_fname.indexOf('"') >= 0)
         {
            alert("double quotes are not allowed");
			document.data.cand_fnam_t.focus();
            return false;
         }






        if (str_fname == "") {
            alert("<% response.write gITEXTA("i_52")%>");
			document.data.cand_fnam_t.focus();
            return false;
        }

       if (!MyReg.test(str_fname))
           {
                alert("Only Alphabet allowed! in first name");
                document.data.cand_fnam_t.focus();
                 return false;
            } 



		var str_lname = "";
		str_lname = document.data.cand_lnam_t.value;
		str_lname = str_lname.replace(/^\s+/g, '').replace(/\s+$/g, '');

        if (str_lname.indexOf("''") != -1 || str_lname.indexOf('"') >= 0)
         {
            alert("double quotes are not allowed");
			document.data.cand_lnam_t.focus();
          
            return false;
         }


         //newvalidation not allow special chracter
           if (str_lname != "")
           {
            if (!MyReg.test(str_lname)) 
            {
                alert("only alphabet allowed! in last name");
                document.data.cand_lnam_t.focus();
                return false;
            } 
           }

         //new validation end


         //new validation for maiden name

         var cand_mnames_t = "";
		cand_mnames_t = document.data.cand_mnam_t.value;
		cand_mnames_t = cand_mnames_t.replace(/^\s+/g, '').replace(/\s+$/g, '');

        if (cand_mnames_t.indexOf("''") != -1 || cand_mnames_t.indexOf('"') >= 0)
         {
            alert("double quotes are not allowed");
			document.data.cand_mnam_t.focus();
            
            return false;
         }

          if (cand_mnames_t != "")
          {
            if (!MyReg.test(cand_mnames_t)) 
            {
                alert("Only Alphabet allowed! in Maiden or Second name");
                document.data.cand_mnam_t.focus();
                 return false;
            } 
          }
         //new validation end for maiden name

        if (str_lname == "") {
            alert("<% response.write gITEXTA("i_51")%>");
			document.data.cand_lnam_t.focus();
            return false;
        }
		var retval;
    	retval = ValidateLastName(document.data.cand_lnam_t.value);
	    if (retval == false)
    	{
            alert("Family name should contain minimum two alphabatic characters.");
			document.data.cand_lnam_t.focus();
            return false;
        }
        if (document.data.cand_gnd_i.value == "") {
            alert("<% response.write gITEXTA("i_53")%>");
			document.data.cand_gnd_i.focus();
            return false;
        }
//        if (document.data.birth_day.value == "") {
//            alert("<% response.write gITEXTA("i_58")%>");
//			document.data.birth_day.focus();
//            return false;
//        }
//        if (document.data.birth_month.value == "") {
//            alert("<% response.write gITEXTA("i_58")%>");
//			document.data.birth_month.focus();
//            return false;
//        }
//        if (document.data.birth_year.value == "") {
//            alert("<% response.write gITEXTA("i_58")%>");
//			document.data.birth_year.focus();
//            return false;
//        }
        if (document.data.cand_nat_c.value == "") {
            alert("<% response.write pv_nattext%>");
			document.data.cand_nat_c.focus();
            return false;
        }
        // To prevent 'script' to be included in text.
         if(!ValidateForm(document.forms[0])){
			return false;
		}

		<%if session("template_org_code") <> 2400 then %>
			var start_date, end_date;
			start_date = document.data.cstart_year.value + document.data.cstart_month.value;
			end_date = document.data.cend_year.value + document.data.cend_month.value;

			if (start_date > end_date) {
			    alert("The Start date must be less than the End date. Please revise your input");
			    return false;
				document.data.cstart_year.focus();
			}
		<%End If%>

}

function StaffInfo()
{
	//alert("testing...." + document.data.thisorg_short.value);
	if (document.data.thisorg_short.value == 1)
	{
		//alert("Yes selected")
		document.getElementById("staffno").style.display="";
		document.getElementById("contract_type").style.display="";

		<%if session("template_org_code") <> 2800 then%>
			document.getElementById("contract_start_date").style.display="";
			document.getElementById("contract_end_date").style.display="";
		<%End if%>

	}
	else
	{
		//Hide and clear fields here
		//alert("No selected")
		document.data.thisorg_staffno.value = "";
		document.getElementById("staffno").style.display="none";
		document.getElementById("contract_type").style.display="none";
		//document.data.cstart_day[0].selected = true

		<%if session("template_org_code") <> 2800 then%>
			document.getElementById("contract_start_date").style.display="none";
			document.getElementById("contract_end_date").style.display="none";
		<%End if%>
	}
}

//Added on 07/20/2009,DD, To validate the lastname to have minimum 2 apha characters.
function ValidateLastName(lname)
{
    var nm= lname;
    if (lname.length < 2)
    {
        return false;
    }
    else if(lname.length == 2)
    {
        //var nm= lname;
        for(var i=0;i<nm.length;i++)
        {
            if (!( (nm.charAt(i)>="a" && nm.charAt(i)<="z") || (nm.charAt(i)>="A" && nm.charAt(i)<="Z")))
            {
                return false;
            }
        }
    }
    else
    {
        //var nm= lname;
        var alphaCount = 0;
        for(var i=0;i<nm.length;i++)
        {
            if (( nm.charAt(i) >= "a" && nm.charAt(i)<="z") || ( nm.charAt(i)>="A" && nm.charAt(i)<="Z"))
            {
                alphaCount = alphaCount + 1;
                if  (alphaCount==2)
                {
                    return true;
                }
            }
        }
        if (i == nm.length)
        {
            return false;
        }
    }


}

// --------------------------------------------
//                  setfocus
// Delayed focus setting to get around IE bug
// --------------------------------------------

function setFocusDelayed()
{
  global_valfield.focus();
}

function setfocus(valfield)
{
  // save valfield in global variable so value retained when routine exits
  global_valfield = valfield;
  setTimeout( 'setFocusDelayed()', 100 );
}

</SCRIPT>


<%'response.write "STAGGER: " & pv_thisorg_staffno



if len(pv_warning) then
	response.write pv_warning
	response.end
end if
%>

<noscript>
<meta http-equiv="refresh" content="0; url=https://erecruit.who.int/demo/public/edit/CheckJavascript.asp" />
</noscript>

<form action="appA-edit.asp" method="POST" name="data" ONSUBMIT="return DataValidation();" class="appForms">
<TABLE cellpadding="0" BORDER="0" width="100%">
<% '<!-------- IF USER CLAIMS STAFF NUMBER WHICH VERIFIES AGAINST STAFF DIR, BUT IS NOT THE RIGHT LAST NAME, NOTIFY  21 JUN 05 LJL
' NEEDNEEDNEED TEST
If (session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500) AND session("CLI_INTERNAL_STAFF") = "2" then

	dim  obj_db_select_CmdVC, VERIFCANDsql, VERIFCAND
  set obj_db_select_CmdVC = server.CreateObject("adodb.command")
  obj_db_select_CmdVC.ActiveConnection = rsys_db_select


	VERIFCANDsql = "	SELECT c.honor_id_c, c.cand_lnam_t, c.cand_fnam_t, s.LastName, s.FirstName, s.RoomNo, s.TelephoneNo, s.Phone, s.Office, s.UnitAcro, rtrim(s.Natlty) AS Natlty, c.staff_nbr_"& session("template_org_code") &" AS thisorg_staffno FROM dbo.td_rsys_cand c LEFT OUTER JOIN dbo.v_staff_list_"& session("template_org_code") &" s ON c.staff_nbr_"& session("template_org_code") &" = s.SID COLLATE SQL_Latin1_General_CP850_CI_AI WHERE c.cand_id_c = ? "
	obj_db_select_CmdVC.CommandText = VERIFCANDsql

	'response.write JAPINFO3sql
		Set VERIFCAND = obj_db_select_CmdVC.Execute(,Array(session("RSYS_EVAL")))


	if pv_DEBUGMAIL = 1 then%>
<TR>
	<td class='alert' colspan='2' bgcolor="yellow"><strong>Your name does not match your name as staff, please notify the system administrator at <a class="topiclink" href='mailto:<%=session("hradminemail")%>'>e-Recruit Tech Support</a><br><br>&nbsp;</strong></td>
</tr>
<TR>
	<td class='alert' colspan='2'>&nbsp;</td>
</tr>

<%
'response.end

if pv_sendemail = 1 then

	if len(VERIFCAND("thisorg_staffno")) then
' ******************************
' MAIL CODE  BEGIN
' ******************************

'16 APR 12 LJL added new mail configuration to avoid having to get schemas from Microsoft

dim sch, cdoConfig, cdoMessage, aFrom
'Sub SendMailCDOCacheConf(aTo, Subject, TextBody, aFrom)
  'cached configuration  
  'Static Conf ' As New CDO.Configuration
'20 DEC 14 LJL removed config for mail for localhost SMTP
'  If IsEmpty(cdoConfig) Then
    'Const cdoOutlookExvbsss = 2
    'Const cdoIIS = 1
'    Set cdoConfig = CreateObject("CDO.Configuration")
'    cdoConfig.Load cdoIIS
'  End If
    
  'Create CDO message object
  '26 SEP 12 LJL modified the to from secant to talenti
  Set cdoMessage = CreateObject("CDO.Message")
    With cdoMessage
       '04 DEC 20 LJL added UTF-8 to send out emails with formatting and ACCENTED CHARS, had to adjust for an unknown reason 
        cdoMessage.BodyPart.Charset = "utf-8" 
'20 DEC 14 LJL removed config for mail for localhost SMTP
'        Set .Configuration = cdoConfig
        .From = "erecruit.alerts@who.int"
        .To = session("hradminemail") & ""
        '25 DEC 14 LJL removed talenti email from notification
        ',erecruit@talentisoft.com
        .Subject = Session("template_org_name") & " - " & VERIFCAND("lastname") & " - Names do not match"
        .HTMLBody = "Applicant ID  " & session("RSYS_EVAL") & " claims to be " & VERIFCAND("thisorg_staffno") & ". <br>NAMES DO NOT MATCH<br><br>GSM:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & VERIFCAND("lastname") & ", " & VERIFCAND("firstname") & "<br><br>e-Recruit:&nbsp;&nbsp;&nbsp;" & VERIFCAND("cand_lnam_t") & ", " & VERIFCAND("cand_fnam_t") & "<br><br>If there is a major difference, contact this applicant to resolve the issue.<br><br>If the applicant seems to be the same, please modify through the GSM name changer in Admin/Data Elements<br><br><br>It can be caused by:<br><br>1. Their not being the person they claim to be (trying to enter a Staff ID to get internal info)<br>2. A slight misspelling of the names<br>3. A change in marital status that has not been recorded in GSM or e-Recruitment<br>4. The legacy data in e-Recruitment may not be correct.	<br><br>-appA-edit-"
 
 
'15 JAN 21 LJL had to add config statements for sending mail - must change the SMTP server 
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
'Name or IP of remote SMTP server
	.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="80.80.227.83"
'Server port
	.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") =25 
	.Configuration.Fields.Update

 
    
    'Set sender address If specified.
    If Len(aFrom) > 0 Then .From = aFrom
    
    'Send the message
    .Send
  End With
    Set cdoMessage = Nothing
'20 DEC 14 LJL removed config for mail for localhost SMTP
'    Set cdoConfig = Nothing
'End Sub



'dim sch, cdoConfig, cdoMessage
'      sch = "http://schemas.microsoft.com/cdo/configuration/"

'    Set cdoConfig = CreateObject("CDO.Configuration")

'    With cdoConfig.Fields
'        .Item(sch & "sendusing") = 2 ' cdoSendUsingPort
'        .Item(sch & "smtpserver") = application("mail_server")
'        .update
'    End With

'    Set cdoMessage = CreateObject("CDO.Message")

'    With cdoMessage
'        Set .Configuration = cdoConfig
'        .From = "erecruit.alerts@who.int"
'        .To = session("hradminemail") & ",erecruit@secantsystems.com"
'        .Subject = Session("template_org_name") & " - " & VERIFCAND("lastname") & " - Names do not match"
'        .HTMLBody = "Applicant ID  " & session("RSYS_EVAL") & " claims to be " & VERIFCAND("thisorg_staffno") & ". <br>NAMES DO NOT MATCH<br><br>GSM:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & VERIFCAND("lastname") & ", " & VERIFCAND("firstname") & "<br><br>e-Recruit:&nbsp;&nbsp;&nbsp;" & VERIFCAND("cand_lnam_t") & ", " & VERIFCAND("cand_fnam_t") & "<br><br>If there is a major difference, contact this applicant to resolve the issue.<br><br>If the applicant seems to be the same, please modify through the GSM name changer in Admin/Data Elements<br><br><br>It can be caused by:<br><br>1. Their not being the person they claim to be (trying to enter a Staff ID to get internal info)<br>2. A slight misspelling of the names<br>3. A change in marital status that has not been recorded in GSM or e-Recruitment<br>4. The legacy data in e-Recruitment may not be correct.	<br><br>-appA-edit-"
'        .Send
'    End With

'    Set cdoMessage = Nothing
'    Set cdoConfig = Nothing
' ******************************
' MAIL CODE  END
' ******************************
		end if
end if

	set obj_db_select_CmdVC = nothing
	end if

End If

'		        <!-------- internal change here ------------------------> WHO ONLY INTERNAL

If (session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500) AND pv_isSTAFF = "1" then
    ' ' '     ' ' '     Do while JAPINFO.eof = false%>
<TR>
	<td><% response.write gITEXTA("i_2")%></td>
    <td><strong><%= JAPINFO3("Salutation")%></strong></td>
</tr>
<TR>
    <td ><% response.write gITEXTA("i_8")%><FONT SIZE='4' COLOR='Red'>*</FONT></td>
    <td><b><% response.write JAPINFO3("FirstName")
	'<!------ first and last name are required fields in table and used later in admin search --------->%></b>
    <INPUT TYPE="hidden" NAME="cand_lnam_t" VALUE="<%= UCASE(JAPINFO3("cand_lnam_t"))%>">
    <INPUT TYPE="hidden" NAME="cand_fnam_t" VALUE="<%= JAPINFO3("cand_fnam_t")%>">
    </td>
</TR>
<TR>
	<td><% response.write gITEXTA("i_3")%></td>
	<td><b><%= JAPINFO3("LastName")%></b></TD>
</TR>
<TR>
	<td><% response.write gITEXTA("i_5")%> </td>
	<td><b><% '20 JUN 08 LJL revised for GSM
	If UCASE(JAPINFO3("sex_code")) = "F" then
	response.write gITEXTA("i_30")%>
	<input type=hidden name=cand_gnd_i value=0>
    <%elseif UCASE(JAPINFO3("sex_code")) = "M" then
	response.write gITEXTA("i_29")%>
	<input type=hidden name=cand_gnd_i value=1>
	<%Else
	response.write("UNKN")
	End If%></b></td>
</TR>
<%
Else%>
<TR>
	<td width="20%"><% response.write gITEXTA("i_2")%><FONT SIZE='4' COLOR='Red'>*</FONT></td>
	<td width="80%"><select  name="honor_id_c">
    <%Do while GETHONOR.eof = false%>
    <OPTION value="<%=GETHONOR("honor_id_c")%>"
	<% If JAPINFO3("honor_id_c") = GETHONOR("honor_id_c") then%>SELECTED<% End If%>>
	<%=GETHONOR("honordsc")%>
    <%GETHONOR.movenext
    loop%>
    </select></td>
</tr>
<TR>
	<td><% response.write gITEXTA("i_8")%><FONT SIZE='4' COLOR='Red'>*</FONT></td>
    
    <td><INPUT size="15" TYPE="text" maxlength="255" NAME="cand_fnam_t" VALUE="<% response.write Server.HTMLEncode(JAPINFO3("cand_fnam_t") & "")%>"></td>
</TR>
<TR>
    <td><% response.write gITEXTA("i_3")%><FONT SIZE='4' COLOR='Red'>*</FONT></td>
    <td><INPUT size="15" TYPE="text"  maxlength="255"  NAME="cand_lnam_t" VALUE="<% response.write Server.HTMLEncode(JAPINFO3("cand_lnam_t") & "")%>" onBlur="javascript:setUpperCase(this);"></TD>
</TR>
<%' REMOVED FOR IFRC and WTO
if session("template_org_code") = 7000 OR session("template_org_code") = 3000 then%>
	<INPUT size="15" TYPE="hidden"  NAME="cand_mnam_t" VALUE="<%= JAPINFO3("cand_mnam_t")%>">
<%else%>
<TR>
    <td><% response.write gITEXTA("i_4")%></td>
    <td><INPUT size="15" TYPE="text"  maxlength="50" NAME="cand_mnam_t" VALUE="<%= Server.HTMLEncode(JAPINFO3("cand_mnam_t") & "")%>"></td>
</TR>
<%end if
'14 MAY 15 LJL no other names section for WMO	
if session("template_org_code") = 2900 then%>
	<INPUT TYPE="hidden"  NAME="cand_onam_t" VALUE="<%= JAPINFO3("cand_onam_t")%>">
<%else%>
<TR>
    <td><% response.write gITEXTA("i_9")%></td>
    <td><INPUT size="15" TYPE="text"  maxlength="50" NAME="cand_onam_t"  VALUE="<%= Server.HTMLEncode(JAPINFO3("cand_onam_t") & "")%>"></td>
</TR>
<%'14 MAY 15 LJL no other names section for WMO	
end if%>

<TR>
    <td><% response.write gITEXTA("i_5")%><FONT SIZE='4' COLOR='Red'>*</FONT> </td>
    <td><select  name="cand_gnd_i">
    <OPTION value="1" <% If JAPINFO3("cand_gnd_i") = "1" then%> SELECTED<%End If%>><% response.write gITEXTA("i_29")%>
    <OPTION value="0" <% If JAPINFO3("cand_gnd_i") = "0" then%> SELECTED<% End If%>><% response.write gITEXTA("i_30")%>
    </select></td>
</TR>
<%' REMOVED FOR IFRC
	'20 MAY 15 LJL not maiden separate for WMO
if session("template_org_code") = 7000 OR session("template_org_code") = 2900 then%>
	<INPUT size="15" TYPE="hidden"  NAME="cand_maiden_t" VALUE="<%= Server.HTMLEncode(JAPINFO3("cand_maiden_t") & "")%>">
<%else%>
<TR>
	<td><% response.write gITEXTA("i_10")%></td>
    <td><INPUT size="15" TYPE="text"  maxlength="50" NAME="cand_maiden_t" VALUE="<%= Server.HTMLEncode(JAPINFO3("cand_maiden_t") & "")%>"></td>
</TR>
<%end if

End If

'11 NOV 08 LJL revised to be day - month - year for American date format dd/mm/yyyy

dim pv_newmonth, pv_newday, pv_newmonth2, pv_newyear

'13 AUG 10 LJL added old and new bday monitoring
pv_bday_prev = JAPINFO3("cand_bth_d")

If isdate(JAPINFO3("cand_bth_d")) = true then
'11 NOV 08 LJL revised to have 0 in front of non two digit numbers
	if month(JAPINFO3("cand_bth_d")) < 10 then
		pv_newmonth = "0" & month(JAPINFO3("cand_bth_d"))
	else
		pv_newmonth = month(JAPINFO3("cand_bth_d"))
	end if
	if day(JAPINFO3("cand_bth_d")) < 10 then
		pv_newday = "0" & day(JAPINFO3("cand_bth_d"))
	else
		pv_newday = day(JAPINFO3("cand_bth_d"))
	end if
	pv_newyear = year(JAPINFO3("cand_bth_d"))
	cand_birth = year(JAPINFO3("cand_bth_d")) & pv_newmonth & pv_newday
	'cand_birth = month(JAPINFO3("cand_bth_d")) &"/" & day(JAPINFO3("cand_bth_d")) & "/"& year(JAPINFO3("cand_bth_d"))
Else
	pv_newmonth = "01"
	pv_newday = "01"
	pv_newyear = "1900"
	cand_birth = "19000101"
End If
	' ' ' response.write "TEST BD " & cand_birth%>
<TR>
	<td><% response.write gITEXTA("i_6")%> </td>
    <td><INPUT size="25" TYPE="text" NAME="cand_bthp_t" maxlength="50" VALUE="<%= Server.HTMLEncode(JAPINFO3("cand_bthp_t") & "")%>">
    <INPUT size="25" TYPE="hidden" NAME="pv_bday_prev" VALUE="<%=JAPINFO3("cand_bth_d")%>"></TD>
</TR>
<%'NOTE: INTERN SECTION to show about age restrictions if any (for WTO)
if session("rsys_intern") = "1" then%>
<TR>
	<td colspan="2" class="alert"><% response.write gITEXTA("i_89")%> </td>
</TR>
<%end if%>
<TR>
	<td><%response.write gITEXTA("i_11")%>
	<%'<!--------- REMOVED 30 SEP 05 <br>response.write gITEXTA("i_49") ------->
	%>
	<FONT SIZE='4' COLOR='Red'>*</FONT></td>
	<td nowrap>
	<%'11 JAN 11 LJL changed the birthdate change to be disallowed for GSM entered dates
	If (session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500) AND pv_isSTAFF = "1" then
	'if session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500 then
	'27 MAR 13 LJL added Birth_date_new to INTERNAL WHO And UNAIDS staff birth date section
			response.write "<strong>" & pv_newday & " " & monthname(pv_newmonth) & " " & pv_newyear & "</strong>"
			%>
            
			<input type="hidden" name=birth_day value="<%=pv_newday%>">
			<input type="hidden" name=birth_month value="<%=pv_newmonth%>">
			<input type="hidden" name=birth_year value="<%=pv_newyear%>">
			
			<input type="hidden" name="birth_date_new" size="12" maxlength="15" value="<%=pv_newday&"-"&monthname(pv_newmonth,true)&"-"&pv_newyear%>">
			 
	<%else%>
    <%
	dim date_d,date_m,date_y,dep_full_1
	dep_full_1 = now()
	date_d=day(dep_full_1)
	date_m=monthname(month(dep_full_1),true)
	date_y=year(dep_full_1)
	
	%>
    <input type="text" name="birth_date_new" size="12" maxlength="15" value="<%=pv_newday&"-"&monthname(pv_newmonth,true)&"-"&pv_newyear%>" id="datepicker1">
            	<!-- <select  name="birth_day" class="halfWidth">
            		<option value="01"<% If pv_newday  = "01" then%> SELECTED<% End If%>><%=("01")%>
            		<option value="02"<% If pv_newday  = "02" then%> SELECTED<% End If%>><%=("02")%>
            		<option value="03"<% If pv_newday  = "03" then%> SELECTED<% End If%>><%=("03")%>
            		<option value="04"<% If pv_newday  = "04" then%> SELECTED<% End If%>><%=("04")%>
            		<option value="05"<% If pv_newday  = "05" then%> SELECTED<% End If%>><%=("05")%>
            		<option value="06"<% If pv_newday  = "06" then%> SELECTED<% End If%>><%=("06")%>
            		<option value="07"<% If pv_newday  = "07" then%> SELECTED<% End If%>><%=("07")%>
            		<option value="08"<% If pv_newday  = "08" then%> SELECTED<% End If%>><%=("08")%>
            		<option value="09"<% If pv_newday  = "09" then%> SELECTED<% End If%>><%=("09")%>
            		<option value="10"<% If pv_newday  = "10" then%> SELECTED<% End If%>><%=("10")%>
            		<option value="11"<% If pv_newday  = "11" then%> SELECTED<% End If%>><%=("11")%>
            		<option value="12"<% If pv_newday  = "12" then%> SELECTED<% End If%>><%=("12")%>
            		<option value="13"<% If pv_newday  = "13" then%> SELECTED<% End If%>><%=("13")%>
            		<option value="14"<% If pv_newday  = "14" then%> SELECTED<% End If%>><%=("14")%>
            		<option value="15"<% If pv_newday  = "15" then%> SELECTED<% End If%>><%=("15")%>
            		<option value="16"<% If pv_newday  =  "16" then%> SELECTED<% End If%>><%=("16")%>
            		<option value="17"<% If pv_newday  =  "17" then%> SELECTED<% End If%>><%=("17")%>
            		<option value="18"<% If pv_newday  =  "18" then%> SELECTED<% End If%>><%=("18")%>
            		<option value="19"<% If pv_newday  =  "19" then%> SELECTED<% End If%>><%=("19")%>
            		<option value="20"<% If pv_newday  =  "20" then%> SELECTED<% End If%>><%=("20")%>
            		<option value="21"<% If pv_newday  =  "21" then%> SELECTED<% End If%>><%=("21")%>
            		<option value="22"<% If pv_newday  =  "22" then%> SELECTED<% End If%>><%=("22")%>
            		<option value="23"<% If pv_newday  =  "23" then%> SELECTED<% End If%>><%=("23")%>
            		<option value="24"<% If pv_newday  =  "24" then%> SELECTED<% End If%>><%=("24")%>
            		<option value="25"<% If pv_newday  =  "25" then%> SELECTED<% End If%>><%=("25")%>
            		<option value="26"<% If pv_newday  =  "26" then%> SELECTED<% End If%>><%=("26")%>
            		<option value="27"<% If pv_newday  =  "27" then%> SELECTED<% End If%>><%=("27")%>
            		<option value="28"<% If pv_newday  =  "28" then%> SELECTED<% End If%>><%=("28")%>
            		<option value="29"<% If pv_newday  =  "29" then%> SELECTED<% End If%>><%=("29")%>
            		<option value="30"<% If pv_newday  =  "30" then%> SELECTED<% End If%>><%=("30")%>
            		<option value="31"<% If pv_newday  =  "31" then%> SELECTED<% End If%>><%=("31")%>
            	</select>
      	<select  name="birth_month" class="halfWidth">
<%
	GETMONTHS.movefirst
      Do while getmonths.eof = false
       if getmonths("month_int_c") < 10 then
                          		pv_newmonth2 = "0" & getmonths("month_int_c")
                          	else
                          		pv_newmonth2 = getmonths("month_int_c")
                          	end if%>
      		<option value="<%=pv_newmonth2%>"<% if pv_newmonth2 = pv_newmonth then%> SELECTED<%end if%>>
			<%=getmonths("monthname")%>
      <%
      getmonths.movenext
      loop%>
      	</select>
<%year_option = ""
if session("rsys_intern") = "1" then
'30 APR 12 LJL removed additional start_year setting
      '   start_year = year(now())-30
'            	<!--------- SET TO 65 years per AP, 24 APR 03 ---------------->
           year_option = ""
	end_year = year(now())-21
    start_year = year(now())-gAGE2("admitem")
'    For yearget=  end_year   to   start_year
'    year_option = year_option & "<option value=" & yearget & ">" & yearget & "^^^</option>"
'    next
     If len(pv_newyear) = "" then
            For yearget=  end_year   to   start_year   step -1
'            		<!-------- check if the year is the same as the indicated birth year ----------->
            year_option = year_option & "<option value=" & yearget & ">" & yearget & "</option>"
            next
     Else
            For yearget=  end_year   to   start_year step -1
'            		<!-------- check if the year is the same as the indicated birth year ----------->
		If yearget = pv_newyear then
            year_option = year_option & "<option value=" & yearget & " SELECTED>" & yearget & "</option>"
        Else
            year_option = year_option & "<option value=" & yearget & ">" & yearget & "</option>"
        End If
            next
    End If
else
			start_year = year(now())-gAGE("admitem")
			' start_year = year(now())-65
'            	<!--------- SET TO 65 years per AP, 24 APR 03 ---------------->
            end_year = year(now())-16
     If len(pv_newyear) = "" then
            For yearget=  end_year   to   start_year   step -1
'            		<!-------- check if the year is the same as the indicated birth year ----------->
            year_option = year_option & "<option value=" & yearget & ">" & yearget & "-</option>"
            next
     Else
            For yearget=  end_year   to   start_year step -1
'            		<!-------- check if the year is the same as the indicated birth year ----------->
		If yearget = pv_newyear then
            year_option = year_option & "<option value=" & yearget & " SELECTED>" & yearget & "</option>"
        Else
            year_option = year_option & "<option value=" & yearget & ">" & yearget & "</option>"
        End If
            next
    End If
end if

            'response.write "YEARER: " & pv_newyear
            'response.write session("rsys_intern")%>
            	<select  name="birth_year" class="halfWidth">
            <% If len(cand_birth) then
            response.write year_option
            Else
			%>
            	<option value="" SELECTED><% response.write gITEXTA("i_64")
            response.write year_option
            End If%>
            	</select>
-->
    <input size="15" type="hidden" name="birth_day_required" maxlength="10" value="<% response.write gITEXTA("i_53")%>">
    <input size="15" type="hidden" name="birth_month_required" maxlength="10" value="<% response.write gITEXTA("i_53")%>">
    <input size="15" type="hidden" name="birth_year_required" maxlength="10" value="<% response.write gITEXTA("i_53")%>">
   <%end if%>
    </TD>
</TR>

<%'19 JUN 08 LJL revised for GSM
If (session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500) AND pv_isSTAFF = "1" then
'28 APR 11 LJL replace form field cand_mar_st_c with form_marital_id%>
<tr>
	<td><% response.write gITEXTA("i_7")%><FONT SIZE='4' COLOR='Red'> *</FONT></td>
    <td><input type="hidden" name="form_marital_id" value="<%=JAPINFO3("marital_id")%>"><strong><%=JAPINFO3("maritaldsc")%></strong></td>
</TR>
<%else%>
<tr>
	<td><% response.write gITEXTA("i_7")%><FONT SIZE='4' COLOR='Red'> *</FONT></td>
    <td><select  name="form_marital_id">
        <%
        Do while GETMARITAL.eof = false%>
       <OPTION value="<%=GETMARITAL("marital_id_c")%>"<%if JAPINFO3("marital_id") = GETMARITAL("marital_id_c") then%>
	   	SELECTED
	   <% End If%>>
	   <%response.write GETMARITAL("maritaldsc")
		GETMARITAL.movenext
        loop%>
        </select></td>
</TR>
<%end if%>

<tr>
	<td colspan='4'><HR size='1' width="100%"></td>
</tr>
    <% '		        <!-------- internal change here ------------------------> WHO ONLY INTERNAL
If (session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500) AND pv_isSTAFF = "1" then
%>
<TR>
    <td><% response.write gITEXTA("i_12")%><FONT SIZE='4' COLOR='Red'>*</FONT> </td>
    <td><b><%=JAPINFO3("internal_nat")%></b>	<input name="cand_nat_c" type="hidden" value="<%=JAPINFO3("Natlty") & "|" & JAPINFO3("cq")%>">
    <input name="prevnat1" type="hidden" value="Natlty">
    </TD>
</TR>

<%Else '	      <!------------------ for making history of changes ------------------------------>
	' UNESCO 2500, ONLY ALLOW 1st NAT change if Admin logged in
	if session("template_org_code") = 2500 AND (instr(session("rsysuser"), "ADM") = false) then	%>
<tr>
	<td colspan='4' class='alert'><em><%=gITEXTA("i_68")%></em></td>
</tr>
<TR>
    <td><% response.write gITEXTA("i_69")%></td>
	<td><strong><% response.write JAPINFO3("ctyname")%></strong>
	<input type="hidden" name="prevnat1" value="">
    <input type="hidden" name="cand_nat_c" value="<%=getnat("code") & "|" & getnat("cq")%>"></td>
</tr>

	<%else%>
	<TR>
    	<td><%' WTO interns use different message
		' REMOVED AND CORRECTED USING NEW NAT INCLUDE TEXT SET
	response.write pv_nattext
	'if session("rsys_intern") = "1" then
	'	response.write gITEXTA("i_67")
	'else
	'	response.write gITEXTA("i_12")
	'end if
	%><FONT SIZE='4' COLOR='Red'>*</FONT> </td>
	<td>
    <input type="hidden" name="prevnat1" value="<% response.write JAPINFO3("cand_nat_c")%>">
     <select  name="cand_nat_c">
    <option value="">-
      <%GETNAT.movefirst
      Do while getnat.eof = false
      If len(getnat("cq")) then%>
     <option value="<%=getnat("code") & "|" & getnat("cq")%>"
	 <% if getnat("code") = JAPINFO3("cand_nat_c") then%> SELECTED<%end if%>><% response.write getnat("ctyname")
		else%>
			<option value="<%=getnat("code") & "|0"%>"<%IF getnat("code") = JAPINFO3("cand_nat_c") then%>SELECTED<%end if%>><%response.write getnat("ctyname")
		end if
      getnat.movenext
      loop%>
      </SELECT>
      </td>
</tr>
	<%end if
 End If
'<!--------------------------- ADD ORG INFO 2000 3000 4000 4 MAR 05 3000 req 2------------------------------------->
'    <!------------------ for making history of changes ------------------------------>
 If session("template_org_code") <> 1000 OR session("template_org_code") <> 1500 then%>
<tr>
    	<td valign="top"><% response.write gITEXTA("i_14")%></TD>
  <td>
    <input type="hidden" name="prevnat2" value="<% response.write JAPINFO3("cand_nat2_c")%>">
    <select  name="cand_nat2_c">
    	<option value="">-
      <%GETNAT2.movefirst
      Do while getnat2.eof = false
      If len(getnat2("cq")) then%>
     	<option value="<%=getnat2("code") & "|" & getnat2("cq")%>"<% if getnat2("code") = JAPINFO3("cand_nat2_c") then%> SELECTED<%end if%>><% response.write getnat2("ctyname")
	 	else%>
		<option value="<%=getnat2("code") & "|0"%>"<%IF getnat2("code") = JAPINFO3("cand_nat2_c") then%>SELECTED<%end if%>><%response.write getnat2("ctyname")
						end if
    getnat2.movenext
    loop%>
    </SELECT>
    </td>
</tr>
<tr>
  	<td><% response.write gITEXTA("i_15")%></td>
  <td>
	<%'    <!------------------ for making history of changes ------------------------------>%>
    <input type="hidden" name="prevnat3" value="<% response.write JAPINFO3("cand_nat3_c")%>">
    <select  name="cand_nat3_c">
    	<option value="">-
      <%GETNAT2.movefirst
      Do while getnat2.eof = false
      If len(getnat2("cq")) then%>
     <option value="<%=getnat2("code") & "|" & getnat2("cq")%>"
	 <% if getnat2("code") = JAPINFO3("cand_nat3_c") then%> SELECTED<%end if%>><% response.write getnat2("ctyname")
		else%>
			<option value="<%=getnat2("code") & "|0"%>"<%IF getnat2("code") = JAPINFO3("cand_nat3_c") then%>SELECTED<%end if%>><%response.write getnat2("ctyname")
		end if
      getnat2.movenext
      loop%>
    </SELECT>
    </td>
</tr>
<%End If%>
<%' REMOVED FOR IFRC
if session("template_org_code") = 7000 then%>
<input TYPE="hidden" name="cand_pnat_i" value="<%=JAPINFO3("cand_pnat_i")%>">
<input TYPE="hidden" name="cand_newnat_c" value="<%=JAPINFO3("cand_newnat_c")%>">
<input TYPE="hidden" NAME="cand_newnat_d" size="12" maxlength="15" value="<%=JAPINFO3("cand_newnat_d")%>" >
<input TYPE="hidden" NAME="cand_expl_t" size="40" maxlength="255" value="<%=JAPINFO3("cand_expl_t")%>">

<%else%>

<TR>
  	<td><% response.write gITEXTA("i_19")%></td>
	<td><select  name="cand_pnat_i">
    <OPTION value="1" <% If JAPINFO3("cand_pnat_i") = "1" then%> SELECTED<% End If%>><% response.write gITEXTA("i_27")%>
    <OPTION value="0" <% If JAPINFO3("cand_pnat_i") = "0" then%> SELECTED<% End If%>><% response.write gITEXTA("i_28")%>
    </select></td>
</tr>
<tr>
  	<td><% response.write gITEXTA("i_39")%></td>
  	<td><input TYPE="text"  NAME="cand_expl_t" size="40" maxlength="255" value="<%= Server.HTMLEncode(JAPINFO3("cand_expl_t") & "")%>"></TD>
</TR>
<TR>
  	<td colspan='2'><% response.write gITEXTA("i_40")%></td>
</TR>
<TR>
  	<td><% response.write gITEXTA("i_12")%></TD>
  	<td>
    <input type="hidden" name="prevnat4" value="<% response.write JAPINFO3("cand_newnat_c")%>">
    	<select  name="cand_newnat_c">
    	<OPTION value="">-
    <%if session("template_org_code") <> 2500 then
    GETNAT.movefirst
    end if
    Do while GETNAT.eof = false
      If len(getnat("code")) then
        if (getnat("ctyname")<> " ") then			%>

    <OPTION value="<%=GETNAT("code")%>"
	<% IF GETNAT("code") = JAPINFO3("cand_newnat_c") then%> SELECTED <%end if%>><%=GETNAT("ctyname")%>
    <%	end if
       end if
    GETNAT.movenext
    loop%>
    	</select></td>
</TR>
<TR>
  	<td>
    <% response.write gITEXTA("i_26")%> </td>
  	<td><input TYPE="text"  NAME="cand_newnat_d" size="12" maxlength="15" value="<%=Server.HTMLEncode(JAPINFO3("cand_newnat_d") & "")%>" id="datepicker" /></TD>
</TR>

<% ' END NO IFRC PREV NAT OR EXPLANATION
end if%>
<TR>
  	<td colspan='2'><hr size="1" width="100%"></TD>
</TR>



<%
' ************************************************************************************************************************************
' BEGIN INTERNAL STAFF SECTION
' ************************************************************************************************************************************

' REMOVED FOR IFRC - NO STAFF OR INTERNAL DETAILS HERE
if session("template_org_code") = 7000 then
	if session("CLI_INTERNAL_STAFF") = "1" then%>
	<INPUT TYPE="hidden" name="thisorg_type" VALUE="1">
    <INPUT TYPE="hidden" NAME="thisorg_short" VALUE="1">
	<INPUT TYPE="hidden" NAME="thisorg_staffno" VALUE="<%=JAPINFO3("thisorg_staffno")%>">
	<%else%>
	<INPUT TYPE="hidden" name="thisorg_type" VALUE="0">
    <INPUT TYPE="hidden" NAME="thisorg_short" VALUE="0">
	<INPUT TYPE="hidden" NAME="thisorg_staffno" VALUE="<%=JAPINFO3("thisorg_staffno")%>">
	<%end if
' PASSPORT DETAILS FOR IFRC%>
<TR>
  	<td colspan='2'><% response.write gITEXTA("i_66")%></td>
</TR>
<TR>
  	<td><% response.write gITEXTA("i_17")%></td>
  	<td>
    	<INPUT size="15" TYPE="text"  maxlength="50" NAME="candmisc_passport_c" VALUE="<%=JAPINFO4("candmisc_passport_c")%>"></td>
</TR>
<TR>
  	<td><% response.write gITEXTA("i_18")%></td>
  	<td>
    	<INPUT size="15" TYPE="text"  maxlength="50" NAME="candmisc_passport_place_c" VALUE="<%=JAPINFO4("candmisc_passport_place_c")%>"></td>
</TR>
<TR>
  	<td><% response.write gITEXTA("i_67")%></td>
  	<td>
    	<INPUT size="15" TYPE="text"  maxlength="30" NAME="candmisc_passport_issue_d" VALUE="<%=JAPINFO4("candmisc_passport_issue_d")%>"></td>
</TR>
<TR>
  	<td><% response.write gITEXTA("i_20")%></td>
  	<td>
    	<INPUT size="15" TYPE="text"  maxlength="30" NAME="candmisc_passport_valid_d" VALUE="<%=JAPINFO4("candmisc_passport_valid_d")%>"></td>
</TR>
	<INPUT TYPE="hidden" name="cstart_day" VALUE="">
	<INPUT TYPE="hidden" name="cstart_month" VALUE="">
	<INPUT TYPE="hidden" name="cstart_year" VALUE="">

	<INPUT TYPE="hidden" name="cend_day" VALUE="">
	<INPUT TYPE="hidden" name="cend_month" VALUE="">
	<INPUT TYPE="hidden" name="cend_year" VALUE="">
	<INPUT TYPE="hidden" name="thisorg_prev" VALUE="0">
<%' REMOVED FOR IFRC - NO STAFF OR INTERNAL DETAILS HERE
' OTHER ORGS HERE

'28 OCT 15 LJL change WTO to use internal staff now 
'elseif session("template_org_code") = 3000 then 	'<!------------------------------------- ADD ORG INFO 3000 req 1 ------------------------------------>//>
'	<INPUT TYPE="hidden" NAME="thisorg_short" VALUE="<%=JAPINFO3("thisorg_short")///>">
'	<INPUT TYPE="hidden" NAME="thisorg_staffno" VALUE="<%=JAPINFO3("thisorg_staffno")///>">
'<INPUT TYPE="hidden" NAME="cstart_day" VALUE="<%=day(cstart)///>">
'<INPUT TYPE="hidden" NAME="cstart_month" VALUE="<%=month(cstart)///>">
'<INPUT TYPE="hidden" NAME="cstart_year" VALUE="<%=year(cstart)///>">
'<INPUT TYPE="hidden" NAME="cend_day" VALUE="<%=day(cend)///>">
'<INPUT TYPE="hidden" NAME="cend_month" VALUE="<%=month(cend)///>">
'<INPUT TYPE="hidden" NAME="cend_year" VALUE="<%=year(cend)///>">
'<%	' END FIRST WTO pv_intdetails
else
' ************************************************************************************************************************************
' BEGIN INTERNAL STAFF SECTION - WHO / UNAIDS
' ************************************************************************************************************************************

	'		if session("rsys_intern") = "1" then
	 '<!------------------------------------------- SHOW REFDB start and end dates for applicant if they are WHO staff ----------------------------->
If (session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500) AND pv_isSTAFF = "1" then
		'17 OCT 06 LJL added for WHO UNAIDS
		%>
<INPUT TYPE="hidden" NAME="thisorg_staffno" VALUE="<%=JAPINFO3("thisorg_staffno")%>">
<%'<INPUT TYPE="hidden" NAME="thisorg_type" VALUE="<%=JAPINFO3("thisorg_type")>">
%>
<INPUT TYPE="hidden" NAME="thisorg_short" VALUE="<%=JAPINFO3("thisorg_short")%>">
<%if isdate(JAPINFO3("contract_start_date")) then%>
<INPUT TYPE="hidden" NAME="cstart_day" VALUE="<%=day(JAPINFO3("contract_start_date"))%>">
<INPUT TYPE="hidden" NAME="cstart_month" VALUE="<%=month(JAPINFO3("contract_start_date"))%>">
<INPUT TYPE="hidden" NAME="cstart_year" VALUE="<%=year(JAPINFO3("contract_start_date"))%>">
<%else%>
<INPUT TYPE="hidden" NAME="cstart_day" VALUE="<%=day(cstart)%>">
<INPUT TYPE="hidden" NAME="cstart_month" VALUE="<%=month(cstart)%>">
<INPUT TYPE="hidden" NAME="cstart_year" VALUE="<%=year(cstart)%>">
<%end if%>

<%if isdate(JAPINFO3("contract_end_date")) then%>
<INPUT TYPE="hidden" NAME="cend_day" VALUE="<%=day(JAPINFO3("contract_end_date"))%>">
<INPUT TYPE="hidden" NAME="cend_month" VALUE="<%=month(JAPINFO3("contract_end_date"))%>">
<INPUT TYPE="hidden" NAME="cend_year" VALUE="<%=year(JAPINFO3("contract_end_date"))%>">
<%else%>

<INPUT TYPE="hidden" NAME="cend_day" VALUE="<%=day(cend)%>">
<INPUT TYPE="hidden" NAME="cend_month" VALUE="<%=month(cend)%>">
<INPUT TYPE="hidden" NAME="cend_year" VALUE="<%=year(cend)%>">
<%end if%>
<tr>
  	<td><% response.write gITEXTA("i_22")%></td>
	<td class="textbold"><%If JAPINFO3("thisorg_short") = "1" then
		response.write gITEXTA("i_27")
    elseIf JAPINFO3("thisorg_short") = "0" then
		response.write gITEXTA("i_28")
		end if%></td>
</TR>

<%if session("template_org_code") <> 2400 then%>
<TR>
  	<td><% response.write gITEXTA("i_48")%></td>
  	<td class="textbold"><%=JAPINFO3("thisorg_staffno") & " "%>
  	  	<%'29 JUN 08 LJL revised for GSM
  	  	'response.write "STA:" & pv_isstaff & "| STB: " & pv_isSTAFFNUMBER & " | "
  	if (session("template_org_code") = 1000 OR session("template_org_code") = 1500) then
  			'if pv_isSTAFF = 0 then
  			if session("CLI_INTERNAL_STAFF") = "1" then
  				response.write "<font color=green><i>Valid staff number</i></font>"
  				If session("CLI_INTERNAL_STAFF") = "2" then
  					response.write " <font color=maroon>Data does not match. Reported to erecruit.</font>"
  				end if
  			else
  				response.write "<font color=maroon><i>Invalid staff number</i></font>"
			end if
	end if%>

  	</td>
</TR>
<%End if%>

<%' 19 JUN 08 LJL revised for GSM
' ************************************************************************************************************************************
' BEGIN INTERNAL STAFF SECTION - WHO / UNAIDS GSM TRUE INTERNAL
' ************************************************************************************************************************************

'11 JAN 11 LJL revised to check for and set the contract type for WHO and UNAIDS

If (session("template_org_code") = 1000 OR session("template_org_code") = 1500) AND pv_isSTAFF = "1" then

Do while GETCONTRACTLEN.eof = false
IF UCASE(JAPINFO3("contract_type")) = UCASE(GETCONTRACTLEN("contractlen_dsc_en_t")) then%>
	<input type="hidden" name="thisorg_type" value="<%=GETCONTRACTLEN("contractlen_id_c")%>">
<%end if
GETCONTRACTLEN.movenext
loop
%>

<%'<tr>
  	'<td colspan="2"><i>Contract Type is presented for information only.  It is set by GSM for <%=session("template_org_name")//> staff.</i></td>
'</tr>
%>
<tr>
  	<td><% response.write gITEXTA("i_23")%></td>
  	<td><strong><%response.write JAPINFO3("contract_type")%></strong>
</td>
 <%'
	'<td><select  name="thisorg_type">
    '<%Do while GETCONTRACTLEN.eof = false//>
    '<OPTION value="<%=GETCONTRACTLEN("contractlen_id_c")//>"<% IF UCASE(JAPINFO3("contract_type")) = UCASE(GETCONTRACTLEN("contractlen_dsc_en_t")) then//> SELECTED<% end if//>>
	 '<%=GETCONTRACTLEN("contractlendsc")//>
    '<%GETCONTRACTLEN.movenext
    'loop//></select></td>
    %>
</TR>

<%else
' ************************************************************************************************************************************
' BEGIN INTERNAL STAFF SECTION - OTHER ORGS
' ************************************************************************************************************************************
' CONTRACT LENGTH FOR OTHER ORGS   %>

<tr>
  	<td><% response.write gITEXTA("i_23")%> </td>
	<td><span id="sporgtype1"> <select name="thisorg_type">
    <%Do while GETCONTRACTLEN.eof = false%>
    <OPTION value="<%=GETCONTRACTLEN("contractlen_id_c")%>"<% IF JAPINFO3("thisorg_type") = GETCONTRACTLEN("contractlen_id_c") then%> SELECTED<% end if%>>
	 <%=GETCONTRACTLEN("contractlendsc")%>
    <%GETCONTRACTLEN.movenext
    loop%></select></span></td>
</TR>
<%end if

' ************************************************************************************************************************************
' IF STAFF - CONTRACT DATES - NOT ITU
'31 MAR 15 LJL also no contract dates for WMO
'28 OCT 15 LJL no contract dates for WTO
' ************************************************************************************************************************************
if session("template_org_code") <> 2400 AND session("template_org_code") <> 2900 AND session("template_org_code") <> 3000 then%>
<TR>
  	<td><% response.write gITEXTA("i_68")%></td>
  	<td class="textbold"><%if len(JAPINFO3("contract_start_date")) then
  	response.write day(JAPINFO3("contract_start_date")) & "-" & monthname(month(JAPINFO3("contract_start_date")),1) & "-" & right(year(JAPINFO3("contract_start_date")), 2)
  	end if%></td>
</tr>
<TR>
  	<td><% response.write gITEXTA("i_69")%></td>
  	<td class="textbold"><%if len(JAPINFO3("contract_end_date")) then
 	response.write day(JAPINFO3("contract_end_date")) & "-" & monthname(month(JAPINFO3("contract_end_date")),1) & "-" & right(year(JAPINFO3("contract_end_date")), 2)
 	end if%></td>
</tr>
<%End If

' ************************************************************************************************************************************
' IF STAFF - INTERNAL INDICATOR - NOT ITU - IS WHO / UNAIDS - ILO - UNESCO
' ************************************************************************************************************************************

Else
		' WHOI/UNAIDS/ILO internal or staff information
		'2500 UNESCO among others
		' OTHER ORGS STAFF INFO
		
		'11 MAR 12 LJL changed ITU int/non-staff to just show and not be editable by applicant
		if session("template_org_code") = 2400 then
		'12 APR 12 LJL added help page for ITU internals%>
<tr>
	<td><% response.write gITEXTA("i_22")%> <a href="../pub-2400-internal-help.asp" target="popupWindow"		onclick="window.open('','popupWindow','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=620,height=auto,left=200,top=100');">[FAQ]</a></td>

  	<td>
		<% If JAPINFO3("thisorg_short") = "1" then
		'11 MAR 12 LJL changed ITU int/non-staff to just show and not be editable by applicant
			 	response.write "<font color=green>" & gITEXTA("i_27") & "</font>"
			else
				response.write gITEXTA("i_28")
			end if%>
	</td>
	</tr>		
		
		<%'11 MAR 12 LJL changed ITU int/non-staff to just show and not be editable by applicant ' 14 april manoj added session 3000 in or
else%>
<tr>
  	<td><% response.write gITEXTA("i_22")%></td>
	<td><select  id="contract_short" name="thisorg_short" <%If (session("template_org_code") = 1000 OR session("template_org_code") = 1500 OR session("template_org_code") = 2000 OR session("template_org_code") = 3000 OR session("template_org_code") = 2800) then%>onchange="javascript:StaffInfo();" <%end if%>>
    <OPTION value="1" <% If JAPINFO3("thisorg_short") = "1" then%> SELECTED<% End If%>><% response.write gITEXTA("i_27")%>
    <OPTION value="0" <% If JAPINFO3("thisorg_short") = "0" then%> SELECTED<% End If%>><% response.write gITEXTA("i_28")%>
    </select></td>

</TR>

<%'11 MAR 12 LJL changed ITU int/non-staff to just show and not be editable by applicant
end if


'31 MAR 15 LJL no staff number for WMO
if session("template_org_code") <> 2400 AND session("template_org_code") <> 2900 then
' ************************************************************************************************************************************
' IF STAFF - GSM VALIDATION - WHO / UNAIDS
' ************************************************************************************************************************************
%>
<TR id="staffno" <%if JAPINFO3("thisorg_short") <> "1" then%>style="display:none"<%end if%>>
  	<td><% response.write gITEXTA("i_48")%></td>
  	<td><INPUT size="15" TYPE="text"  maxlength="10" NAME="thisorg_staffno" VALUE="<%=JAPINFO3("thisorg_staffno")%>" validate="integer" message="You must enter a valid Staff ID number only in the Staff Number field"> &nbsp;
  	  	<%'29 JUN 08 LJL revised for GSM
  	if (session("template_org_code") = 1000 OR session("template_org_code") = 1500) then
  		if pv_isstaff = 1 then
  			if pv_isSTAFFNUMBER = 0 then
  				response.write "<font color=maroon><i>Invalid staff number (ST6)</i></font>"
  			else
  				response.write "<font color=green><i>Valid staff number</i></font>"
			end if
		elseif len(JAPINFO3("thisorg_staffno")) then
			response.write "<font color=maroon><i>Staff number not correct or not found (ST5)</i></font>"
		else
		end if

  		If session("CLI_INTERNAL_STAFF") = "2" then
  				response.write "<br><font color=maroon>Data does not match. Reported to Help Desk.</font>"
  		end if
	end if%>


  	</td>
</TR>
<%End If

	' BEGIN CHECK INTERN 2
	'if session("rsys_intern") = "1" then
	
		'11 MAR 12 LJL changed ITU int/non-staff to just show and not be editable by applicant
	if session("template_org_code") = 2400 then
	
	else %>
<tr id="contract_type">
  	<td><span id="spancontract"><% response.write gITEXTA("i_23")%> </span></td>
	<td><span id="sporgtype"> <select  name="thisorg_type">
    <%Do while GETCONTRACTLEN.eof = false%>
    <OPTION value="<%=GETCONTRACTLEN("contractlen_id_c")%>"<% IF JAPINFO3("thisorg_type") = GETCONTRACTLEN("contractlen_id_c") then%> SELECTED<% end if%>>
	 <%=GETCONTRACTLEN("contractlendsc")%>
    <%GETCONTRACTLEN.movenext
    loop
    GETCONTRACTLEN.movefirst%></select></span></td>
</TR>
<%'11 MAR 12 LJL changed ITU int/non-staff to just show and not be editable by applicant
end if

' ADD NEW ORGS TO THIS LIST IF NEEDED CONTRACT DETAILS (most will need them if wanting internal staff info)

' ************************************************************************************************************************************
' IF STAFF - CONTRACT DATES
' ************************************************************************************************************************************
'12 FEB 15 LJL added UNWOMEN, WMO
'31 MAR 15 LJL no contract dates for WMO


if session("template_org_code") = 1000 OR session("template_org_code") = 1500 Or session("template_org_code") = 2000 Or session("template_org_code") = 2400 Or session("template_org_code") = 2600 then
     if session("template_org_code") <> 2400 then
%>
<TR id="contract_start_date" <%if JAPINFO3("thisorg_short") <> "1" then%>style="display:none"<%end if%>>
  	<td><% response.write gITEXTA("i_68")%></td>
  	<td nowrap>
    <% If isdate(JAPINFO2("thisorg_start")) then
    cstart = JAPINFO2("thisorg_start")
    Else
    cstart = ""
    End If%>
    	<select  name="cstart_day">
    	<% If len(cstart) = 0 then%>
    		<option value=""<% If cstart = "" then%> SELECTED<% End If%>><% response.write gITEXTA("i_62")%>
    		<option value="1">01
    		<option value="2">02
    		<option value="3">03
    		<option value="4">04
    		<option value="5">05
    		<option value="6">06
    		<option value="7">07
    		<option value="8">08
    		<option value="9">09
    		<option value="10">10
    		<option value="11">11
    		<option value="12">12
    		<option value="13">13
    		<option value="14">14
    		<option value="15">15
    		<option value="16">16
    		<option value="17">17
    		<option value="18">18
    		<option value="19">19
    		<option value="20">20
    		<option value="21">21
    		<option value="22">22
    		<option value="23">23
    		<option value="24">24
    		<option value="25">25
    		<option value="26">26
    		<option value="27">27
    		<option value="28">28
    		<option value="29">29
    		<option value="30">30
    		<option value="31">31
    	<% Else%>
    		<option value="1" <% If day(cstart)  =  "1" then%> SELECTED<% End If%>><%="01"%>
    		<option value="2"<% If day(cstart)  = "2" then%> SELECTED<% End If%>><%="02"%>
    		<option value="3"<% If day(cstart)  =  "3" then%> SELECTED<% End If%>><%="03"%>
    		<option value="4"<% If day(cstart)  =  "4" then%> SELECTED<% End If%>><%="04"%>
    		<option value="5"<% If day(cstart)  =  "5" then%> SELECTED<% End If%>><%="05"%>
    		<option value="6"<% If day(cstart)  =  "6" then%> SELECTED<% End If%>><%="06"%>
    		<option value="7"<% If day(cstart)  =  "7" then%> SELECTED<% End If%>><%="07"%>
    		<option value="8"<% If day(cstart)  =  "8" then%> SELECTED<% End If%>><%="08"%>
    		<option value="9"<% If day(cstart)  =  "9" then%> SELECTED<% End If%>><%="09"%>
    		<option value="10"<% If day(cstart)  = "10" then%> SELECTED<% End If%>><%="10"%>
    		<option value="11"<% IF day(cstart)  = "11" then%> SELECTED<% End If%>><%="11"%>
    		<option value="12"<% IF day(cstart)  = "12" then%> SELECTED<% End If%>><%="12"%>
    		<option value="13"<% If day(cstart)  = "13" then%> SELECTED<% End If%>><%="13"%>
    		<option value="14"<% If day(cstart)  = "14" then%> SELECTED<% End If%>><%="14"%>
    		<option value="15"<% If day(cstart)  = "15" then%> SELECTED<% End If%>><%="15"%>
    		<option value="16"<% If day(cstart)  = "16" then%> SELECTED<% End If%>><%="16"%>
    		<option value="17"<% If day(cstart)  = "17" then%> SELECTED<% End If%>><%="17"%>
    		<option value="18"<% If day(cstart)  = "18" then%> SELECTED<% End If%>><%="18"%>
    		<option value="19"<% If day(cstart)  = "19" then%> SELECTED<% End If%>><%="19"%>
    		<option value="20"<% If day(cstart)  = "20" then%> SELECTED<% End If%>><%="20"%>
    		<option value="21"<% If day(cstart)  = "21" then%> SELECTED<% End If%>><%="21"%>
    		<option value="22"<% If day(cstart)  = "22" then%> SELECTED<% End If%>><%="22"%>
    		<option value="23"<% If day(cstart)  = "23" then%> SELECTED<% End If%>><%="23"%>
    		<option value="24"<% If day(cstart)  = "24" then%> SELECTED<% End If%>><%="24"%>
    		<option value="25"<% If day(cstart)  = "25" then%> SELECTED<% End If%>><%="25"%>
    		<option value="26"<% If day(cstart)  = "26" then%> SELECTED<% End If%>><%="26"%>
    		<option value="27"<% If day(cstart)  = "27" then%> SELECTED<% End If%>><%="27"%>
    		<option value="28"<% If day(cstart)  = "28" then%> SELECTED<% End If%>><%="28"%>
    		<option value="29"<% If day(cstart)  = "29" then%> SELECTED<% End If%>><%="29"%>
    		<option value="30"<% If day(cstart)  =  "30" then%> SELECTED<% End If%>><%="30"%>
    		<option value="31"<% If day(cstart)  =  "31" then%> SELECTED<% End If%>><%="31"%>
    		<option value="">-
    	<% End If%>
    	</select>
    	<select  name="cstart_month">
    	<% If len(cstart) = "0" OR cstart = "" then%>
    		<option value="" SELECTED><% response.write gITEXTA("i_63")
	GETMONTHS.movefirst
    Do while getmonths.eof = false%>
    		<option value="<%=getmonths("month_id_c")%>"><%=getmonths("monthname")%>
    <%getmonths.movenext
    loop
    Else
	GETMONTHS.movefirst
    Do while getmonths.eof = false%>
    		<option value="<%=getmonths("month_id_c")%>"<% If cdbl(month(cstart)) = cdbl(getmonths("month_id_c")) then%> SELECTED<% End If%>><%=getmonths("monthname")%>
    <%getmonths.movenext
    loop
    End If
    '26 APR 10 LJL removed
'    		<option value="">-
%>
    	</select>
       	<select  name="cstart_year">
            <%year_option = ""
            start_year = 1975
            end_year = year(now())+65
            If isdate(cstart) then
            	For yearget=  end_year   to  start_year   step -1
					If cdbl(yearget) = cdbl(year(cstart)) then
            			response.write "<option value=""" & yearget & """ SELECTED>" & yearget & "</option>"
         			Else
            			response.write "<option value=""" & yearget & """>" & yearget & "</option>"
        			End If
            	next
            Else%>
            		<option value="" SELECTED><%=gITEXTA("i_64")%></option>
			<%For yearget=  end_year   to   start_year   step -1
            		response.write "<option value=""" & yearget & """>" & yearget & "</option>"
            	next
            End If%>
		</select>
    </td>
</TR>
<TR id="contract_end_date" <%if JAPINFO3("thisorg_short") <> "1" then%>style="display:none"<%end if%>>
    	<td><% response.write gITEXTA("i_69")%></td>
    	<td nowrap>
      <% If isdate(JAPINFO2("thisorg_end")) then
      cend = JAPINFO2("thisorg_end")
      Else
      cend = ""
      End If%>
      	<select  name="cend_day">
      	<% If len(cend) = 0 then%>
      		<option value=""<% If cend = "" then%> SELECTED<% End If%>><% response.write gITEXTA("i_62")%>
      		<option value="1">01
      		<option value="2">02
      		<option value="3">03
      		<option value="4">04
      		<option value="5">05
      		<option value="6">06
      		<option value="7">07
      		<option value="8">08
      		<option value="9">09
      		<option value="10">10
      		<option value="11">11
      		<option value="12">12
      		<option value="13">13
      		<option value="14">14
      		<option value="15">15
      		<option value="16">16
      		<option value="17">17
      		<option value="18">18
      		<option value="19">19
      		<option value="20">20
      		<option value="21">21
      		<option value="22">22
      		<option value="23">23
      		<option value="24">24
      		<option value="25">25
      		<option value="26">26
      		<option value="27">27
      		<option value="28">28
      		<option value="29">29
      		<option value="30">30
      		<option value="31">31
      	<% Else%>
      		<option value="1"<% If day(cend) = "1" then%> SELECTED<% End If%>><%="01"%>
      		<option value="2"<%  If day(cend) = "2" then%> SELECTED<% End If%>><%="02"%>
      		<option value="3"<%  If day(cend) = "3" then%> SELECTED<% End If%>><%="03"%>
      		<option value="4"<%  If day(cend) = "4" then%> SELECTED<% End If%>><%="04"%>
      		<option value="5"<%  If day(cend) = "5" then%> SELECTED<% End If%>><%="05"%>
      		<option value="6"<%  If day(cend) = "6" then%> SELECTED<% End If%>><%="06"%>
      		<option value="7"<%  If day(cend) = "7" then%> SELECTED<% End If%>><%="07"%>
      		<option value="8"<%  If day(cend) = "8" then%> SELECTED<% End If%>><%="08"%>
      		<option value="9"<%  If day(cend) = "9" then%> SELECTED<% End If%>><%="09"%>
      		<option value="10"<%  If day(cend) = "10" then%> SELECTED<% End If%>><%="10"%>
      		<option value="11"<%  If day(cend) = "11" then%> SELECTED<% End If%>><%="11"%>
      		<option value="12"<%  If day(cend) = "12" then%> SELECTED<% End If%>><%="12"%>
      		<option value="13"<%  If day(cend) = "13" then%> SELECTED<% End If%>><%="13"%>
      		<option value="14"<%  If day(cend) = "14" then%> SELECTED<% End If%>><%="14"%>
      		<option value="15"<%  If day(cend) = "15" then%> SELECTED<% End If%>><%="15"%>
      		<option value="16"<%  If day(cend) = "16" then%> SELECTED<% End If%>><%="16"%>
      		<option value="17"<%  If day(cend) = "17" then%> SELECTED<% End If%>><%="17"%>
      		<option value="18"<%  If day(cend) = "18" then%> SELECTED<% End If%>><%="18"%>
      		<option value="19"<%  If day(cend) = "19" then%> SELECTED<% End If%>><%="19"%>
      		<option value="20"<%  If day(cend) = "20" then%> SELECTED<% End If%>><%="20"%>
      		<option value="21"<%  If day(cend) = "21" then%> SELECTED<% End If%>><%="21"%>
      		<option value="22"<%  If day(cend) = "22" then%> SELECTED<% End If%>><%="22"%>
      		<option value="23"<%  If day(cend) = "23" then%> SELECTED<% End If%>><%="23"%>
      		<option value="24"<%  If day(cend) = "24" then%> SELECTED<% End If%>><%="24"%>
      		<option value="25"<%  If day(cend) = "25" then%> SELECTED<% End If%>><%="25"%>
      		<option value="26"<%  If day(cend) = "26" then%> SELECTED<% End If%>><%="26"%>
      		<option value="27"<%  If day(cend) = "27" then%> SELECTED<% End If%>><%="27"%>
      		<option value="28"<%  If day(cend) = "28" then%> SELECTED<% End If%>><%="28"%>
      		<option value="29"<%  If day(cend) = "29" then%> SELECTED<% End If%>><%="29"%>
      		<option value="30"<%  If day(cend) = "30" then%> SELECTED<% End If%>><%="30"%>
      		<option value="31"<%  If day(cend) = "31" then%> SELECTED<% End If%>><%="31"%>
      		<option value="">-
      	<% End If%>
      	</select>
      	<select  name="cend_month">
    	<% If len(cend) = "0" OR cend = "" then%>
      		<option value="" SELECTED><% response.write gITEXTA("i_63")
	GETMONTHS.movefirst
      Do while getmonths.eof = false%>
      		<option value="<%=getmonths("month_id_c")%>"><%=getmonths("monthname")%>
      <%GETMONTHS.movenext
      loop
      Else
	GETMONTHS.movefirst
      Do while getmonths.eof = false%>
      		<option value="<%=getmonths("month_id_c")%>" <% if cdbl(month(cend)) = cdbl(getmonths("month_id_c")) then%> SELECTED<%end if%>><%=getmonths("monthname")%>
      <%getmonths.movenext
      loop
      End If
    '26 APR 10 LJL removed
'    		<option value="">-
%>
      	</select>
       	<select  name="cend_year">
            <%year_option = ""
            start_year = 1975
            end_year = year(now())+65
            If isdate(cend) then
            	For yearget=  end_year   to  start_year   step -1
					If cdbl(yearget) = cdbl(year(cend)) then
            			response.write "<option value=""" & yearget & """ SELECTED>" & yearget & "</option>"
         			Else
            			response.write "<option value=""" & yearget & """>" & yearget & "</option>"
        			End If
            	next
            Else%>
            		<option value="" SELECTED><%=gITEXTA("i_64")%></option>
			<%For yearget=  end_year   to   start_year   step -1
            		response.write "<option value=""" & yearget & """>" & yearget & "</option>"
            	next
            End If%>
		</select>
	</td>
</TR>
    <% End If
    %>
<tr>
	<td valign='top'>&nbsp;</td>
</TR>
<tr>
	<td><% response.write gITEXTA("i_81")%></td>
    <td><select  name="thisorg_prev">
    <OPTION value="0"<% If JAPINFO3("thisorg_prev") = "0" then%> SELECTED<% End If%>><% response.write gITEXTA("i_28")%>
    <OPTION value="1"<% If JAPINFO3("thisorg_prev") = "1" then%> SELECTED<% End If%>><% response.write gITEXTA("i_27")%>
    </select>
	<%' PREV APPLIED YEAR, ONLY WTO
	if session("template_org_code") = 3000 then
		 response.write gITEXTA("i_90")%>
       	<select  name="cand_prev_year_c">
            <%year_option = ""
            start_year = 1975
            end_year = year(now())+65
            If JAPINFO3("cand_prev_year_c") <> "" then
            	For yearget=  end_year   to  start_year   step -1
					If cdbl(yearget) = cdbl(JAPINFO3("cand_prev_year_c")) then
            			response.write "<option value=""" & yearget & """ SELECTED>" & yearget & "</option>"
         			Else
            			response.write "<option value=""" & yearget & """>" & yearget & "</option>"
        			End If
            	next
            Else
			%>
            		<option value="" SELECTED><%=gITEXTA("i_64")%></option>
			<%
            	For yearget=  end_year   to   start_year   step -1
            		response.write "<option value=""" & yearget & """>" & yearget & "</option>"
            	next
            End If%>
		</select>
		<%end if%>
	</td>
</TR>
<tr>
	<td>&nbsp;</td>
</tr>
<%' END IFRC NO STAFF OR INTERNAL DETAILS
end if
end if
	' END INTERNAL STAFF SECTION WHO/UNAIDS/ILO
end if


' ************************************************************************************************************************************
' END IS STAFF SECTION
' ************************************************************************************************************************************
%>




<TR>
	<td colspan='2'>&nbsp;</TD>
</TR>
<TR>
	<td><% response.write gITEXTA("i_71")%> <FONT SIZE='4' COLOR='Red'>*</FONT></TD>
	<td>
       <select  name="webfamiliar_id_c">
        <%Do while getfamiliars.eof = false%>
      <option value="<%=getfamiliars("webfamiliar_id_c")%>"<%if JAPINFO3("webfamiliar_id_c") = getfamiliars("webfamiliar_id_c") then%> SELECTED <%end if%>><%=getfamiliars("familiardsc")%>
        <%getfamiliars.movenext
        loop%>
        </select>
		<input type="hidden" name="webfamiliar_id_c_required" value="Kindly indicate how you became familiar with our web site">
		</td>
</TR>
<TR>
	<td><% response.write gITEXTA("i_72")%></TD>
	<td><INPUT size="45" maxlength="250" TYPE="text" NAME="webfamiliar_refer_t" VALUE="<%= Server.HTMLEncode(JAPINFO3("webfamiliar_refer_t") & "")%>" SIZE="8" MAXLENGTH="8"></TD>
</TR>
<TR>
	<td colspan='2'>&nbsp;</TD>
</TR>
<tr>
    <td colspan='2' align='center'>
    <INPUT TYPE="hidden" NAME="cand_id_c" VALUE="<%=session("RSYS_EVAL")%>">
	<INPUT TYPE="hidden" NAME="cand_ipa_c" VALUE="<% If Request.servervariables("REMOTE_ADDR") <> "" then
		response.write left(request.servervariables("REMOTE_ADDR"),15)
	Else%>webuser<% End If%>">
    <%Acount = int(japinfo2("EditA")+1)%>
    <INPUT TYPE="hidden" NAME="upd_d" value="<%=now()%>">
    <INPUT TYPE="hidden" NAME="editA" VALUE="<%=Acount%>">
    <INPUT TYPE="hidden" NAME="DOeditA" VALUE="99">
	<INPUT  TYPE="submit" class="login_submit" VALUE="<% response.write gITEXTA("i_24")%>   ">
	</td>
</tr>
</TABLE>
</form>
<%'response.write "<Br>TESTER INT1: " & session("CLI_INTERNAL_STAFF")
pv_last_update="20 Aug 24"
'<<--Added by Interface on 05/02/2007
set obj_int_select_Cmd = nothing
set obj_logs_CmdI = nothing
set obj_logs_CmdII = nothing
set obj_logs_CmdIII = nothing
set obj_logs_CmdIV = nothing
set obj_logs_CmdV = nothing
set obj_logs_CmdVI = nothing
set obj_logs_CmdVII = nothing
set obj_db_CmdI = nothing
set obj_db_CmdII = nothing
set obj_db_select_Cmd = nothing
set obj_db_select_CmdI = nothing
set obj_db_select_CmdII = nothing

set obj_db_select_CmdIII = nothing
set obj_db_select_CmdIV = nothing
set obj_db_select_CmdV = nothing
set obj_db_select_CmdVI = nothing
set obj_db_select_CmdVII = nothing
set obj_db_select_CmdVIII = nothing
set obj_db_select_CmdIX = nothing
set obj_db_select_CmdX = nothing
set obj_db_select_Cmd5 = nothing

set obj_db_select_CmdRank = nothing
set obj_db_select_CmdEdit = nothing

'-->>
rsys_db_select.close
Set rsys_db_select = nothing
rsys_int_select.close
Set rsys_int_select = nothing

	rsys_db.close
	Set rsys_db = nothing
	rsys_logs.close
	Set rsys_logs = nothing

' *****************************************************
' BEGIN BOTTOM INCLUDES
' *****************************************************%>
<!--#include file="../includes/include_pubedit_frame_bottom.asp"-->
<%
' *****************************************************
' END BOTTOM INCLUDES
' *****************************************************%>

