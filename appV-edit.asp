<%option explicit

if session("lng") = "fr" then
	Session.LCID = 1036
    else 
    Session.LCID =3081
    end if
'<!-------------------------------------------NOTES ----------
'MODS --
'21 NOV 03 LJL added log to update when a person completes an app, per TA
'16 AUG 04 LJL fixed verification page for those who have not completed their app yet.
'31 MAY 05 LJL added WTO nationality filter to stop non-nationals from succeeding in applying to a post
'16 JUN 05 LJL changed education scheme to be lines and not fixed
'15 OCT 05 ADDED ADDITIONAL DROP DOWN FOR LAW/VERIFICATION PER IFRC 7000
'15 OCT 05 FOUND that javascript validation code has to be BELOW the top include files or it changes the fonts and layout a bit.
'17 JAN 06 LJL size limit on entry for verif_name_t
'18 JAN 06 LJL updated query for WTO, no law indication for them on this page.
'19 JAN 06 LJL added age check for WTO, 62 years
'16 MAR 06 LJL added IFRC RC experience check for people applying to those jobs.
'28 MAR 06 LJL worked on RC experience check again, for release
'08 JUL 06 LJL moved notes to top, added update text as include, which used to be in the main page top include
'17 AUG 06 LJL added maxlength for the verification fields.  Not sure why they were not there before.
'17 SEP 06 LJL added Motivation (Covering) Letter for UNAIDS 1500 - covering letter set as type 13 for cl
'11 Nov 06 ac - this was not on the server so I put the one I have up. PLEASE overwrite this one with yours, you may have been working on this one?
'20 NOV 06 LJl had to add extra check for when a person is applying to a post.  Previously, if they messed up verif page for applying, it would add the doc anyway... - new textadd var value is 88 for this page, 99 for doc-edit page
'03 FEB 07 LJL not show the verif page if the person has just submitted it.
'01 MAY 07 INT Changed queries to parameterized
'08 MAY 07 LJl removed a line of code that halted progress on page, getjobname.movefirst
'29 MAY 07 INT Change Variable name obj_db_select_Cmd to obj_db_select_CmdI
'11 JUN 07 INT modify the query
'09 JUL 07 LJL modified how the page checks if the profile is complete, from teh include page session var appallok now
'23 JUL 07 LJL added js notification for not answering the YES at the top of the page, finally.
'27 JUL 07 LJL revised IFRC version of this page, and corrected verification validation error which was going in circles
'29 JUL 07 LJL had to change the obj name for the answer list query as it was outputting incorrect answers
'29 JUL 07 :LJL corrected the replace() for the SQL injection in query
'31 JUL 07 LJL changed postadding script to be outside the other scripts because not always on the page, only when applying to a VN.
 '01 AUG 07 LJL removed out setting of NATOK for WTO, and set it on the auth pages and appA, related to not allowing certain users to apply if not member state applicant.
'05 AUG 07 LJL corrected issue with applying and the notification of YES to confirmation
'01 SEP 07 LJL maxlength on verif location
'16 SEP 07 LJL added age checking for WTO
'30 OCT 07 LJL set true age output as one year  lower than the admin age for internships for WTO
'01 FEB 08 LJL removed AoE checking, per ILO
'11 FEB 08 LJL text is being too long for the question text to be inserted.  Maybe because of apostrophes?  Shorten the text length on the subsequent page...
'15 FEB 08 LJL changed the violations question to have to be answered each time, per UNAIDS.
'13 AUG 08 DD js validation added for UNAIDS to make explanation mandatory for violations question if answered yes. Along with changes to display previously selected answer.
'14 NOV 08 LJL added logging of person trying to apply to vac
'25 NOV 08 DD  modified rsyslog table names to conatin current year's tag.
'02 JUN 09 DD  Changes for checkbox updatation on profile page.
'17 SEP 09 DD  Size of asterisks are increased to bigger size
'14 OCT 09 LJL ADD VN FAMILIAR FEATURE - currently only for ILO and IFRC ---------------------------------------->
'02 DEC 09 LJL added check to make sure person indicates where they heard of the VN - ILO, IFRC
'25 JAn 10 LJL changed webfamiliar to be location = 2 or 9 for VN sourcing or all
'12 MAR 10 LJL reduced all text input fields by 2 chars to stop truncation problem
'23 APR 10 LJL revised the footers and format css for larger size across
'01 MAY 10 LJL lots of formatting and added new system to tell applicants which sections are not complete
'01 MAY 10 LJL moved header set and check login to before loading content -
	'order is Edit Section, check_complete, then page headers, then ejobs-updates include
'20 AUG 10 LJL default langcheck to 1 and then have it set back to 0 if any are not set right.
'20 AUG 10 defined the IFRC check if complete section.  Didn't have it set before.
'20 AUG 10 LJL added If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0" for Edu section
'05 OCT 10 LJL set cand_law_i as no default in the db table as it needs to be set by user
'03 NOV 10 DD SQL query executed by ADODB Command object and NOT by Connection object.
'16 DEC 10 DD SQL query executed by new ADODB Command object always
'01 MAR 11 LJL reduced candjob_webfamiliar_ field to 98 from 100 to see if that helps in truncation errors.
'06 MAR 11 LJL added check on internal staff for comptencies check on 1000 WHO complete
'07 MAR 11 LJL set to be the max age from admitem. It was set to 62 for some reason.
'06 JUN 12 LJL no Other Info section required for IFRC interns
'06 JUN 12 LJL added Areas of Expertise check for IFRC 7000
'11 SEP 12 LJL changed Qcount to allow for null value
'31 OCT 12 LJL order the webfamiliar items by rank_i
'03 NOV 12 LJL appV include text setting
'03 FEB 13 LJL IFRC Also requires covering letter
'20 MAY 13 LJL added different text for INTERNS for WTO when AppV is completed and not applying directly for a post. - added to ejobs-updates.asp as the include file - this is where "This is NOT an application to a vacancy. It is only the completion of your Personal History Form. You may now 'Apply to Vacancies' by clicking on the above link." is found.
'19 MAR 15 LJL modified order of web familiar to alpha, per ILO
'29 APR 15 LJL add VN sourcing for WMO
'05 MAY 15 LJL added WMO section of verification items
'12 MAY 15 LJL webfamiliar check for WMO
'06 AUG 15 LJL added WMO translation for required text box.
'01 SEP 15 LJL moved WMO law section to Other info section
'03 SEP 15 LJL added OI check for WMO
'22 Feb 16 LJL/MC Add emergency resource avail dates for WHO specific VNs
'26 March 2023 Manoj added Dismiss ,resigned and name include in UN fields to this form and validation

' SET TO HAVE LEFT MENU
pv_col2right = "1"
pv_heavybottomvery = "1"

' CHECK LOGGED IN AND DIM DB's
' ***********************************************************
%>
<!--#include file = "../includes/include_check_login.asp"-->
<!--#include file = "../../includes/rsys_db.asp"-->
<!--#include file = "../../includes/rsys_db_select.asp"-->
<!--#include file = "../../includes/rsys_int_select.asp"-->
<!--#include file = "../../includes/rsys_logs.asp"-->
<% '<<--Modified by Interface on 06/11/2007 %>
<!--#include file="../../sysdev/adovbs.inc" -->
<%
' ***********************************************************
' END CHECK LOGGED IN AND DIM DB's
dim JAPYsql, JAPY, JAPINFO2sql, JAPINFO2, GetAnssql, GetAns, updFootersql, updFooter, goeditsql, goedit, logeditsql, logedit, pv_newage
dim getjobnamesql, getjobname, GETJAPNATsql, GETJAPNAT, GetAnswerssql, GetAnswers, GetJAPAnswerssql, GetJAPAnswers
dim pv_qanswerid, faqid, question_counter,  jobtitle,  tablesize, goform, questionnumber, qcounter, qstid,  qrowcount,  Qcount, Signcount, dater
dim GETAGEsql, GETAGE, ageok, GETRCEXPsql, GETRCEXP, obj_db_select_aJobEditCmdV, gMINAGEsq, gMAXAGEsql, gAGEsql, gAGE, gMINAGEsql, gMINAGE, pv_viewvn, getfamiliars, getfamiliarssql

'<<--Modified by Interface on 05/29/2007
dim obj_logs_CmdI, obj_db_CmdI, obj_db_CmdII, obj_db_select_CmdI, obj_db_select_CmdII, obj_db_select_CmdIII, obj_db_select_CmdIV, obj_db_select_CmdV
'Added on 11/25/2008 DD
dim currentYear
'-->>
' BEGIN INCLUDE TEXT DB
dim gVTextsql, gVText
dim obj_int_select_Cmd10
 set obj_int_select_Cmd10 = server.CreateObject("adodb.command")
 obj_int_select_Cmd10.ActiveConnection = rsys_int_select
'<<--Modified by Interface on 05/01/2007
'22 Feb 16 LJL/MC Add emergency resource avail dates for WHO specific VNs
gVTextsql = "SELECT i_1, i_3, i_4, i_5, i_7, i_8, i_9, i_11, i_12, i_13, i_14, i_15, i_18, i_19, i_20, I_32, i_33, i_34, i_40, i_44, i_45, I_46, i_47, i_48, i_50, i_55, i_61, i_62, i_64, i_70, i_71, i_72, i_80, i_81, i_82, i_83, i_84, i_text1, i_text2, i_text3,i_94,i_95,i_96  FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'V' "
obj_int_select_Cmd10.CommandText = gVTextsql
Set gVText = obj_int_select_Cmd10.Execute(,Array(session("template_org_code"),session("lng")))
'-->>
' END INCLUDE TEXT DB
' BEGIN INCLUDE TEXT DB
dim gQTextsql, gQText
dim obj_int_select_Cmd11
 set obj_int_select_Cmd11 = server.CreateObject("adodb.command")
 obj_int_select_Cmd11.ActiveConnection = rsys_int_select
'<<--Modified by Interface on 05/01/2007
gQTextsql = "SELECT i_1, i_3, i_4, i_5, i_7, i_8, i_9, i_20, i_11, i_12, i_13, i_14, i_15, i_18, i_19, i_text1, i_text2, i_40, i_55, i_60, i_62, i_64 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'QST' "
obj_int_select_Cmd11.CommandText = gQTextsql
Set gQText = obj_int_select_Cmd11.Execute(,Array(session("template_org_code"),session("lng")))

dim gPHTextsql, gPHText
dim obj_int_select_Cmd12
 set obj_int_select_Cmd12 = server.CreateObject("adodb.command")
 obj_int_select_Cmd12.ActiveConnection = rsys_int_select
'<<--Modified by Interface on 05/01/2007
gPHTextsql = "SELECT i_65, i_66, i_67 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'PH' "
obj_int_select_Cmd12.CommandText = gPHTextsql
Set gPHText = obj_int_select_Cmd12.Execute(,Array(session("template_org_code"),session("lng")))


set obj_db_select_CmdII = server.CreateObject("adodb.command")
obj_db_select_CmdII.ActiveConnection = rsys_db_select

'-->>
' END INCLUDE TEXT DB

if Request.form("postadding") = 1 then
	pv_page_title= gVText("i_1")
else
	pv_page_title= gVText("i_50")
end if
titletype = 8
'faqid="11"
faqid=46



currentYear =  year(now())  'Added on 11/25/2008 DD, to get the value of current year.

'******************************
' BEGIN V EDIT
'******************************

If Request.form("GOEDITV") = "99" then

'<<--Modified by Interface on 06/11/2007
 set obj_db_CmdI = server.CreateObject("adodb.command")
 obj_db_CmdI.ActiveConnection = rsys_db
 updFootersql = "UPDATE td_rsys_cand SET cand_fil_d = getdate(), cand_fil_t = ?, cand_ipa_c = ?, cand_law_i = ?, cand_law_m = ?, cand_verif_ip_c = ?, cand_verif_name_t = ?, user_id_t = ?,cand_dismissed_i =? ,cand_dismissed_m =?, cand_resigned_i =?, cand_resigned_m =?, cand_nameinclude_i =?,cand_nameinclude_UN =? , upd_d = getdate() WHERE cand_id_c = ?"
 obj_db_CmdI.CommandText = updFootersql
if len(request.form("cand_fil_t")) then
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_fil_t" ,adVarChar,adParamInput,200,request.form("cand_fil_t"))
else
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_fil_t" ,adVarChar,adParamInput,200,NULL)
end if

 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_ipa_c" ,adVarChar,adParamInput,30,left(request.servervariables("REMOTE_ADDR"),20))
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_law_i" ,adTinyInt,adParamInput,1,request.form("cand_law_i"))

 if len(request.form("cand_law_m")) then
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_law_m" ,adLongVarChar,adParamInput,10000,request.form("cand_law_m"))
else
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_law_m" ,adLongVarChar,adParamInput,10000,NULL)
end if

 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_verif_ip_c" ,adVarChar ,adParamInput,100,left(request.servervariables("REMOTE_ADDR"),20))
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_verif_name_t" ,adVarChar,adParamInput,150,left(request.form("cand_verif_name_t"),75))
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@user_id_t" ,adVarChar,adParamInput,100,left(session("RSYSUSER"),20))
       obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_dismissed_i" ,adTinyInt,adParamInput,1,request.form("cand_dismissed_i"))
 if len(request.form("cand_dismissed_m")) then
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_dismissed_m" ,adLongVarChar,adParamInput,10000,request.form("cand_dismissed_m"))
else
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_dismissed_m" ,adLongVarChar,adParamInput,10000,NULL)
end if
    
     obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_resigned_i" ,adTinyInt,adParamInput,1,request.form("cand_resigned_i"))
     if len(request.form("cand_resigned_m")) then
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_resigned_m" ,adLongVarChar,adParamInput,10000,request.form("cand_resigned_m"))
else
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_resigned_m" ,adLongVarChar,adParamInput,10000,NULL)
end if
         obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_nameinclude_i" ,adTinyInt,adParamInput,1,request.form("cand_nameinclude_i"))
     if len(request.form("cand_nameinclude_UN")) then
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_nameinclude_UN" ,adLongVarChar,adParamInput,10000,request.form("cand_nameinclude_UN"))
else
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_nameinclude_UN" ,adLongVarChar,adParamInput,10000,NULL)
end if
 obj_db_CmdI.Parameters.Append obj_db_CmdI.CreateParameter("@cand_id_c" ,adInteger ,adParamInput,4,session("RSYS_EVAL"))

  

 Set updFooter = obj_db_CmdI.Execute()

 set obj_db_CmdII = server.CreateObject("adodb.command")
 obj_db_CmdII.ActiveConnection = rsys_db
 goeditsql = "UPDATE tx_rsys_candedit SET editFooter = ?, editFooter_d = getdate() WHERE cand_id_c = ?"
 obj_db_CmdII.CommandText = goeditsql
 Set goedit = obj_db_CmdII.Execute(,Array(request.form("editFooter"),session("RSYS_EVAL")))

'06 AUG 09 LJL Added if they UPDATED THEIR PROFILE PORTION ----------------------------------------------------
dim obj_logs_CmdPU, PAEDITsql, PAEDIT
	set obj_logs_CmdPU = server.CreateObject("adodb.command")
	obj_logs_CmdPU.ActiveConnection = rsys_logs
	If session("RSYSUSER") <> "" then
		PAEDITsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  " (candupd_ip_c, cand_id_c, candupd_page_t,user_id_t,candupd_site_c) " &_
		"VALUES (?, ?, 'V-Edited', ?,?)"
		obj_logs_CmdPU.CommandText = PAEDITsql
		Set PAEDIT = obj_logs_CmdPU.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),session("RSYSUSER"),left(session("template_org_name"),20)))
	Else
		PAEDITsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  " (candupd_ip_c, cand_id_c, candupd_page_t,candupd_site_c) " &_
		"VALUES (?, ?, 'V-Edited',?)"
		obj_logs_CmdPU.CommandText = PAEDITsql
		Set PAEDIT = obj_logs_CmdPU.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),left(session("template_org_name"),20)))
	End if
' -------------------------------------------------------------------------------------------------------------------------------------------

' -->>
else
'06 AUG 09 LJL Added if they ACCESSED THEIR PROFILE PORTION -------------------------------------------------
dim obj_logs_CmdPA, PACCsql, PACC
	set obj_logs_CmdPA = server.CreateObject("adodb.command")
	obj_logs_CmdPA.ActiveConnection = rsys_logs
	If session("RSYSUSER") <> "" then
		PACCsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  " (candupd_ip_c, cand_id_c, candupd_page_t,user_id_t,candupd_site_c) " &_
		"VALUES (?, ?, 'V-Viewed', ?,?)"
		obj_logs_CmdPA.CommandText = PACCsql
		Set PACC = obj_logs_CmdPA.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),session("RSYSUSER"),left(session("template_org_name"),20)))
	Else
		PACCsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  " (candupd_ip_c, cand_id_c, candupd_page_t,candupd_site_c) " &_
		"VALUES (?, ?, 'V-Viewed',?)"
		obj_logs_CmdPA.CommandText = PACCsql
		Set PACC = obj_logs_CmdPA.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),left(session("template_org_name"),20)))
	End if
' -------------------------------------------------------------------------------------------------------------------------------------------


End If


' **************************************************************************************************
'14 OCT 09 LJL ADD VN FAMILIAR FEATURE - currently only for ILO and IFRC ---------------------------------------->
' **************************************************************************************************
'31 OCT 12 LJL order the webfamiliar items by rank_i
'29 APR 15 LJL add VN sourcing for WMO
if session("template_org_code") = 2000 OR session("template_org_code") = 7000 OR session("template_org_code") = 2900 then
	getFamiliarssql = " SELECT webfamiliar_dsc_"& session("lng") &"_t as familiardsc, webfamiliar_id_c FROM core_webfamiliar WHERE webfamiliar_inactind_i <> 1 AND webfamiliar_location_i_" & session("template_org_code") & " IN (2,9) AND webfamiliar_thisorg_" & session("template_org_code") &  "= 1 ORDER BY 1 ASC"
	
	' webfamiliar_rank_i "
	'response.write getfamiliarsql
	'response.end

	obj_db_select_CmdII.CommandText = getFamiliarssql
	'set getFamiliars =rsys_db_select.execute(getFamiliarssql)
	set getFamiliars =obj_db_select_CmdII.execute()
end if
' **************************************************************************************************
'14 OCT 09 LJL END ADD VN FAMILIAR FEATURE - currently only for ILO and IFRC ---------------------------------------->
' **************************************************************************************************

'******************************
' END V EDIT
'******************************
'
' ***********************************************************
' BEGIN INCLUDES
' ***********************************************************
%>
<!--#include file = "../includes/include_check_complete.asp"-->
<!--#include file="../includes/include_pubedit_frame_top.asp"-->
<%
' ***********************************************************
' END INCLUDES
' ***********************************************************
'08 JUL 06 LJL added UPDATE PHRASE HERE
%>
<!-- #include file="../edit/ejobs-updates.asp"-->
<%

appcheck = 1
'  CHECKAPPOKsql = "{call erstp_cand_verify_param(" & session("RSYS_EVAL") & ")}"
'  set CHECKAPPOK =rsys_db_select.execute(CHECKAPPOKsql)

'  CHECKEMPOKsql = "{call erstp_cand_verify_employ_param(" & session("RSYS_EVAL") & ")} "
'  set CHECKEMPOK =rsys_db_select.execute(CHECKEMPOKsql)

	' EDUCATION CHECK
'If CHECKAPPOK("cand_edu_degree_n") = "0" OR CHECKAPPOK("cand_edu_degree_n") = "" OR CHECKAPPOK("edulines") < 0 then
'  appcheck = 0
'  session("educheck") = 0
'Else
'  session("educheck") = 1
'End If
	' response.write "T 002" & appcheck & "<br>"

  ' <!----------------------- verify that english language is done by the user ----------------->
'If CHECKAPPOK("cand_en_spk_n") = "" then
'  appcheck = 0
'  session("langcheck") = 0
'Else
'  session("langcheck") = 1
'End If
	' response.write "T 003" & appcheck & "<br>"

  ' <!----------------------- verify that french language is done by the user ----------------->
'If CHECKAPPOK("cand_fr_spk_n") = "" then
'  appcheck = 0
'  session("langcheck") = 0
'Else
'  session("langcheck") = session("langcheck")
'End If
	' response.write "T 004" & appcheck & "<br>"

  ' <!----------------------- verify that spanish language is done by the user ----------------->
'If CHECKAPPOK("cand_es_spk_n") = "" then
'  appcheck = 0
'  session("langcheck") = 0
'Else
'  session("langcheck") = session("langcheck")
'End If
	' response.write "T 005" & appcheck & "<br>"

'If CHECKEMPOK("jlecount") <= 0 then
'  	appcheck = 0
'  	session("employcheck") = 0
'Else
'  	session("employcheck") = 1
'End If

'<--Modified by DD on 06/04/2009
set obj_db_select_Cmd = server.CreateObject("adodb.command")
obj_db_select_Cmd.ActiveConnection = rsys_db_select

if Request.form("postadding") = 1 then

if session("appallok") = 0 then
	response.write gVText("i_text3") & "<br><br>" & gVText("i_61")
	response.end
end if

  	'<-- response.write "T JOB " & Request.form("jobinfo_uid_c")

	if Request.form("jobinfo_uid_c") <> "" then
		'<<--Modified by Interface on 05/01/2007
  		getjobnamesql = " SELECT jobinfo_thisorg_1000 ,isnull(jobinfo_indic_avail_i,0) as jobinfo_indic_avail_i, jobinfo_contracttype_id_c, status_id_c, jobinfo_job_en_t, jobinfo_job_fr_t, jobinfo_job_es_t, jobinfo_acl_d, jobinfo_en_ok_i, jobinfo_fr_ok_i, jobinfo_es_ok_i, jobinfo_vac2_c FROM td_rsys_jobinfo WHERE jobinfo_uid_c = ? "
		obj_db_select_Cmd.CommandText = getjobnamesql
		Set getjobname = obj_db_select_Cmd.Execute(,Array(Request.form("jobinfo_uid_c")))

' 14 NOV 08 LJL added logging of View of VN
pv_viewvn = "VN-AttemptApply " & getjobname("jobinfo_vac2_c") & " (" & Request.form("jobinfo_uid_c") & ")"

 set obj_logs_CmdI = server.CreateObject("adodb.command")
 obj_logs_CmdI.ActiveConnection = rsys_logs
 'Response.Write ("session(RSYSUSER)=" & session("RSYSUSER"))
 if session("RSYSUSER") = "" then
    //Added on 11/25/2008 DD, for entry to rsys logs for current year's table
	//logeditsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "(candupd_ip_c, cand_id_c, candupd_page_t, candupd_site_c) VALUES" &_
	logeditsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  "(candupd_ip_c, cand_id_c, candupd_page_t, candupd_site_c) VALUES" &_
	" (?, ?, '" & pv_viewvn & "', ?)"
	obj_logs_CmdI.CommandText = logeditsql
	Set logedit = obj_logs_CmdI.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),session("template_org_name")))
 else
	//logeditsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "(candupd_ip_c, cand_id_c, candupd_page_t, user_id_t, candupd_site_c) VALUES" &_
	logeditsql = "INSERT INTO td_rsys_log_candupdates_" & session("template_org_code") & "_" & currentYear &  "(candupd_ip_c, cand_id_c, candupd_page_t, user_id_t, candupd_site_c) VALUES" &_
	" (?, ?, '" & pv_viewvn & "', ?, ?)"
	obj_logs_CmdI.CommandText = logeditsql
	Set logedit = obj_logs_CmdI.Execute(,Array(left(request.servervariables("REMOTE_ADDR"), 25),session("RSYS_EVAL"),left(session("RSYSUSER"),20),session("template_org_name")))
	end if

  		'-->>
		' response.write "T SQL REC " & GETJOBNAME.eof
			'  	<!-------------------------- LANGUAGE ADJUSTMENT ---------------------------------->
			'28 MAR 06 LJL loop removed, should only be one post in each case.
		'Do while getjobname.eof = false
  			if session("lng") = "en" then
  				if getjobname("jobinfo_en_ok_i") then
  					jobtitle = UCASE(getjobname("jobinfo_job_en_t"))
  				elseif getjobname("jobinfo_en_ok_i") = 0 AND getjobname("jobinfo_fr_ok_i") = 1 then
  					jobtitle = UCASE(getjobname("jobinfo_job_fr_t"))
  				elseif getjobname("jobinfo_en_ok_i") = 0 AND getjobname("jobinfo_fr_ok_i") = 0 AND getjobname("jobinfo_es_ok_i") = 1 then
  					jobtitle = UCASE(getjobname("jobinfo_job_es_t"))
  				else
  					jobtitle = UCASE(getjobname("jobinfo_job_en_t"))
  				end if
  			elseif session("lng") = "fr" then
  				if getjobname("jobinfo_fr_ok_i") then
  					jobtitle = UCASE(getjobname("jobinfo_job_fr_t"))
  				elseif getjobname("jobinfo_fr_ok_i") = 0 AND getjobname("jobinfo_en_ok_i")= 1 then
  					jobtitle = UCASE(getjobname("jobinfo_job_en_t"))
  				elseif getjobname("jobinfo_fr_ok_i") = 0 AND getjobname("jobinfo_en_ok_i") = 0 AND getjobname("jobinfo_es_ok_i") = 1 then
  					jobtitle = UCASE(getjobname("jobinfo_job_es_t"))
  				else
  					jobtitle = UCASE(getjobname("jobinfo_job_en_t"))
  				end if
  			elseif session("lng") = "es" then
  				if getjobname("jobinfo_es_ok_i") then
  					jobtitle = UCASE(getjobname("jobinfo_job_es_t"))
  				elseif getjobname("jobinfo_es_ok_i") = 0 AND getjobname("jobinfo_en_ok_i") = 1 then
  					jobtitle = UCASE(getjobname("jobinfo_job_en_t"))
  				elseif getjobname("jobinfo_es_ok_i") = 0 AND getjobname("jobinfo_en_ok_i") = 0 AND getjobname("jobinfo_fr_ok_i") = 1 then
  					jobtitle = UCASE(getjobname("jobinfo_job_fr_t"))
  				else
  					jobtitle = UCASE(getjobname("jobinfo_job_en_t"))
  				end if
  			end if
		'getjobname.movenext
  		'loop
	else%>
  		You have reached this page in error.  Please contact <a href="mailto:erecruit@secantsystems.com?subject=error on appv-edit">with details on what you were attempting to do</a>
		<%response.end
	end if


	'<!------------------------------ ADD ORG INFO 3000 31 MAY 05 LJL only for WTO, only nationals apply to posts ----------------------------->
	if session("template_org_code") = "3000" then
		'<!------------------------------ ADD ORG INFO 3000 31 MAY 05 LJL only for WTO, only nationals apply to posts ----------------------------->
		'<!---------------------- GET CAND NATIONALITY TO SEE IF MEMBER STATE -------------------------------------------------->
		'<<--Modified by Interface on 05/01/2007

'01 AUG 07 LJL revised the WTO nat check
''		GETJAPNATsql = "SELECT c.cand_nat_c FROM td_rsys_cand c, tr_rsys_country ct WHERE c.cand_id_c = ? AND ct.country_member_" & session("template_org_code") & " = 1 AND c.cand_nat_c = ct.cty_id_c 	"
''		obj_db_select_Cmd.CommandText = GETJAPNATsql
''		Set GETJAPNAT = obj_db_select_Cmd.Execute(,Array(session("RSYS_EVAL")))
		'-->>
''		session("checknatok") = 1
''		if GETJAPNAT.eof= false then
''			session("checknatok") = 1
''		else
''			session("checknatok") = 0
''		end if

		ageok = 0

		'<<--Modified by Interface on 05/01/2007
dim obj_db_select_CmdA, obj_db_select_CmdAI, obj_db_select_CmdAII
 set obj_db_select_CmdA = server.CreateObject("adodb.command")
 obj_db_select_CmdA.ActiveConnection = rsys_db_select
		GETAGEsql = " SELECT c.cand_bth_d FROM td_rsys_cand c WHERE c.cand_id_c = ?"
		obj_db_select_CmdA.CommandText = GETAGEsql
		Set GETAGE = obj_db_select_CmdA.Execute(,Array(session("RSYS_EVAL")))

		if session("template_org_code") = 3000 then
			' GET MAX AGE FROM ADMIN ITEMS (DATA ELEMENTS)
			if GETJOBNAME("status_id_c") = "15" then

			 set obj_db_select_CmdAI = server.CreateObject("adodb.command")
			 obj_db_select_CmdAI.ActiveConnection = rsys_db_select

              gMINAGEsql = "	SELECT rtrim(admitem_default_" & session("template_org_code") & ") AS admitem 			FROM td_rsys_admitem 			WHERE admitem_ident_c = 'MINIMUM-AGE-INTERN' "
              obj_db_select_CmdAI.CommandText = gMINAGEsql
              'set gMINAGE = rsys_db_select.execute(gMINAGEsql)
              set gMINAGE = obj_db_select_CmdAI.execute()

              if gMINAGE.eof = true then
              		response.write	("Error.  Alert Admin that minimum values not set - New Registration2<br><br>")
              response.end
              end if
			  dim pv_age_forward, pv_age_days_older, pv_birth_daysince_older
			  ' GET THAT YEAR AND DATE WHEN IT WOULD BE FOR ORG MAX AGE
			 pv_age_forward = dateadd("yyyy", -1 * gMINAGE("admitem"), now())
			 pv_age_days_older = datediff("y", pv_age_forward, now())
			 ' GET DAYS SINCE THEIR BIRTH DATE
			'response.write "AGER:" & getAGE("cand_bth_d")
			'response.end
			pv_birth_daysince_older = datediff("y", getAGE("cand_bth_d"), now())
			' COMPARE TWO
			'response.write "<BR>DATE3:" & pv_age_days_older & ":|:" & pv_birth_daysince_older & "<br>"
			'if pv_birth_daysince >  pv_age_days then
			'	response.write "TOO OLD"
			'else
			'	response.write "JUST RIGHT"
			'end if
			'response.end
			if pv_birth_daysince_older < pv_age_days_older then
				response.write "<br><br>" & gPHText("i_65") & gMINAGE("admitem") & " " & gPHText("i_66") & "<br><br>" '& pv_birth_date_check
				response.end
			end if
              gAGEsql = "SELECT rtrim(admitem_default_" & session("template_org_code") & ") AS admitem 			FROM td_rsys_admitem 			WHERE admitem_ident_c = 'MAXIMUM-AGE-INTERN' "
			else
              gAGEsql = "SELECT rtrim(admitem_default_" & session("template_org_code") & ") AS admitem 			FROM td_rsys_admitem 			WHERE admitem_ident_c = 'MAXIMUM-AGE' "
			end if

			 set obj_db_select_CmdAII = server.CreateObject("adodb.command")
			 obj_db_select_CmdAII.ActiveConnection = rsys_db_select

            obj_db_select_CmdAII.CommandText = gAGEsql
            'set gAGE = rsys_db_select.execute(gAGEsql)
            set gAGE = obj_db_select_CmdAII.execute()

              if gAGE.eof = true then
              		response.write	"Error.  Alert Admin that minimum values not set - New Registration2<br><br>"
              response.end
              end if
			  dim pv_age_back, pv_age_days, pv_birth_daysince
			  ' GET THAT YEAR AND DATE WHEN IT WOULD BE FOR ORG MAX AGE
			 pv_age_back = dateadd("yyyy", -1 * gAGE("admitem"), now())
			 pv_age_days = datediff("y", pv_age_back, now())
			 ' GET DAYS SINCE THEIR BIRTH DATE
			pv_birth_daysince = datediff("y", getAGE("cand_bth_d"), now())
			' COMPARE TWO
			'response.write "<Br>DATE333:" & pv_age_days & ":|:" & pv_birth_daysince &"<BR>"
			'response.write "<BR>AGE:" & gAGE("admitem")
			'if pv_birth_daysince >  pv_age_days then
			'	response.write "TOO OLD"
			'else
			'	response.write "JUST RIGHT"
			'end if
			'response.end
			if pv_birth_daysince > pv_age_days then
			pv_newage = gAGE("admitem") - 1
			' HAD TO SET THE THRESHOLD AS LOWER THAN THE ADMIN AGE
				response.write "<br><br>" & gPHText("i_65") & pv_newage & " " & gPHText("i_67") & "<br><br>" '& pv_birth_date_check
				response.end
			end if
		end if
		'-->>
	' response.write "AGER: " & DATEDIFF("yyyy", GETAGE("cand_bth_d"),now()) & "<br>"
'07 MAR 11 LJL set to be the max age from admitem. It was set to 62 for some reason.
		if DATEDIFF("yyyy", GETAGE("cand_bth_d"),now()) > gAGE("admitem") then
			ageok = 1
		else
			ageok = 0
		end if

		if session("checknatok") = 0 then%>
<table>
  <TR>
    	<td valign="top" align="center" class="alert" colspan="6">
		<%=gVText("i_62") & "<br><br>" & gVText("i_61")%></td>
    </tr>
    <TR>
      	<td valign="top" align="center" colspan="6">&nbsp;</td>
    </tr>
  </table>
		<%
		response.end
		end if

		if ageok = 1 then%>
<table>
  <TR>
    	<td valign="top" align="center" class="alert" colspan="6">
		<%=gVText("i_64") & "<br><br>" & gVText("i_61")%></td>
    </tr>
    <TR>
      	<td valign="top" align="center" colspan="6">&nbsp;</td>
    </tr>
  </table>
		<%
		response.end
		end if


	end if



end if

	set obj_db_select_CmdIII = server.CreateObject("adodb.command")
	obj_db_select_CmdIII.ActiveConnection = rsys_db_select


  '<<--Modified by Interface on 05/01/2007
  JAPYsql = "SELECT upd_d, cand_fnam_t, cand_lnam_t, cand_fil_t, cand_fil_d, cand_verif_name_t, cand_law_i, cand_law_m,cand_dismissed_i, cand_dismissed_m,cand_resigned_i,cand_resigned_m,cand_nameinclude_i,cand_nameinclude_UN  	FROM td_rsys_cand WHERE cand_id_c = ?"
  obj_db_select_CmdIII.CommandText = JAPYsql
  Set JAPY = obj_db_select_CmdIII.Execute(,Array(session("RSYS_EVAL")))
  '-->>
  '<<--Modified by Interface on 05/29/2007
  set obj_db_select_CmdI = server.CreateObject("adodb.command")
  obj_db_select_CmdI.ActiveConnection = rsys_db_select
  JAPINFO2sql = "SELECT editFooter, EditQ FROM tx_rsys_candedit WHERE cand_id_c = ?	"
  obj_db_select_CmdI.CommandText = JAPINFO2sql
  Set JAPINFO2 = obj_db_select_CmdI.Execute(,Array(session("RSYS_EVAL")))
 '-->>
if Request.querystring("cand_fil_d") then
  	dater = Request.Querystring("cand_fil_d")
else
	dater = Now()
end if

  if Request.form("postadding") = 1 then
  ' was 581
  tablesize = "500"
  else
  ' was 431
  tablesize = "400"
  end if

%>
<head>
    <script src="js/jquery-1.8.2.js" type="text/javascript"></script>
<SCRIPT language="JavaScript" type="text/javascript" src="../../js/Validation_Script.js"></script>



<script>


$(document).ready(function(){
     if (document.data.cand_law_i.value =="0")
     {   
            document.data.cand_law_m.readOnly = true; document.data.cand_law_m.disabled = true; 
     }
     if (document.data.cand_law_i.value =="1")
     { 
        document.data.cand_law_m.readOnly = false; document.data.cand_law_m.disabled = false;
     }
   
    $("#cand_law_i").change(function(){
       if (document.data.cand_law_i.value =="0" || document.data.cand_law_i.value =="")
                 { document.data.cand_law_m.readOnly = true; document.data.cand_law_m.disabled = true;}
           if (document.data.cand_law_i.value =="1")
                  { document.data.cand_law_m.readOnly = false;  document.data.cand_law_m.disabled = false;}
                  

       
       
    });
    
       if (document.data.cand_dismissed_i.value =="0")
     {   
            document.data.cand_dismissed_m.readOnly = true; document.data.cand_dismissed_m.disabled = true; 
     }
     if (document.data.cand_dismissed_i.value =="1")
     { 
        document.data.cand_dismissed_m.readOnly = false; document.data.cand_dismissed_m.disabled = false;
     }
   
    $("#cand_dismissed_i").change(function(){
       if (document.data.cand_dismissed_i.value =="0" || document.data.cand_dismissed_i.value =="")
                 { document.data.cand_dismissed_m.readOnly = true; document.data.cand_dismissed_m.disabled = true;}
           if (document.data.cand_dismissed_i.value =="1")
                  { document.data.cand_dismissed_m.readOnly = false;  document.data.cand_dismissed_m.disabled = false;}
                  

       
       
    });

       if (document.data.cand_resigned_i.value =="0")
     {   
            document.data.cand_resigned_m.readOnly = true; document.data.cand_resigned_m.disabled = true; 
     }
     if (document.data.cand_resigned_i.value =="1")
     { 
        document.data.cand_resigned_m.readOnly = false; document.data.cand_resigned_m.disabled = false;
     }
   
    $("#cand_resigned_i").change(function(){
       if (document.data.cand_resigned_i.value =="0" || document.data.cand_resigned_i.value =="")
                 { document.data.cand_resigned_m.readOnly = true; document.data.cand_resigned_m.disabled = true;}
           if (document.data.cand_resigned_i.value =="1")
                  { document.data.cand_resigned_m.readOnly = false;  document.data.cand_resigned_m.disabled = false;}
                  

       
       
    });

      if (document.data.cand_nameinclude_i.value =="0")
     {   
            document.data.cand_nameinclude_UN.readOnly = true; document.data.cand_nameinclude_UN.disabled = true; 
     }
     if (document.data.cand_nameinclude_i.value =="1")
     { 
        document.data.cand_nameinclude_UN.readOnly = false; document.data.cand_nameinclude_UN.disabled = false;
     }
   
    $("#cand_nameinclude_i").change(function(){
       if (document.data.cand_nameinclude_i.value =="0" || document.data.cand_nameinclude_i.value =="")
                 { document.data.cand_nameinclude_UN.readOnly = true; document.data.cand_nameinclude_UN.disabled = true;}
           if (document.data.cand_nameinclude_i.value =="1")
                  { document.data.cand_nameinclude_UN.readOnly = false;  document.data.cand_nameinclude_UN.disabled = false;}
                  

       
       
    });



});
function DataValidation()


    {



    //       manoj changes
            $(".radio-field").each(function() {
  
                var id = $(this).attr("id"); //radiobutton id
      
                varidval =$(this).val();  //textbox id
   
                if($(this).is(":checked")) 
                { 
    
                if ($("#txt"+varidval).length>0)
                    {
                       var textBox = $("#txt"+varidval).val();
                      if(textBox =="")
                        {
                          alert(" Please explain: Enter details to the answer ");
                           event.preventDefault();
                        }
                    }
    
                }


            });


    //       manoj changes 

    
    	// Modified by Interface on 07/23/2007
		// To prevent 'script' to be included in text.
	    if(!ValidateForm(document.forms[0]))
		{
			return false;
		}


         if (document.data.cand_fil_t.value == "") {
            alert("<%=gVText("i_9")%>");
            return false;
        }

        var str_var1 = "";
		str_var1 = document.data.cand_verif_name_t.value;
		str_var1 = str_var1.replace(/^\s+/g, '').replace(/\s+$/g, '');
        if (str_var1 == "") {
            alert("<%=gVText("i_8")%>");
            return false;
        }


        if (document.data.cand_verif_name_t.value == "") {
            alert("<%=gVText("i_8")%>");
            return false;
        }


       

<%if Request.form("postadding") = 1 then%>
        if (document.data.postaddselect.value == "0") {
            alert("<%=gVText("i_55")%>");
            return false;
        }
	<%'29 APR 15 LJL add VN sourcing for WMO
	if session("template_org_code") = 2000 OR session("template_org_code") = 7000 OR session("template_org_code") = 2900 then%>
         if (document.data.candjob_webfamiliar_id_c.value == "") {
            alert("<% response.write gVText("i_80")%>");
            return false;
        }
	<%end if

else
'01 SEP 15 LJL moved WMO law section to Other info section
	 if (session("template_org_code") <> 3000 AND session("template_org_code") <> 2900) AND NOT Request.form("GOEDITV") = "99" then%>
         if (document.data.cand_law_i.value == "") {
            alert("<% response.write gVText("i_70")%>");
			document.data.cand_law_i.focus();
            return false;
        }

	<%end if

     if session("template_org_code") = 1500 then %>
       if (document.data.cand_law_i.value == 1)
        {
           if (document.data.cand_law_m.value == "")
           {
            alert("<% response.write gVText("i_72")%>");
			document.data.cand_law_m.focus();
            return false;
           }
        }
      if (document.data.cand_dismissed_i.value == 1)
        {
           if (document.data.cand_dismissed_m.value == "")
           {
            alert("<% response.write gVText("i_72")%>");
			document.data.cand_dismissed_m.focus();
            return false;
           }
        }
      if (document.data.cand_resigned_i.value == 1)
        {
           if (document.data.cand_resigned_m.value == "")
           {
            alert("<% response.write gVText("i_72")%>");
			document.data.cand_resigned_m.focus();
            return false;
           }
        }
      if (document.data.cand_nameinclude_i.value == 1)
        {
           if (document.data.cand_nameinclude_UN.value == "")
           {
            alert("<% response.write gVText("i_72")%>");
			document.data.cand_nameinclude_UN.focus();
            return false;
           }
        }

     <%end if

end if%>


       if (document.data.cand_law_i.value =="1")
       {
           if(document.data.cand_law_m.value =="")
           {
	       //please give complete details in the space below
           alert("<%=gVText("i_72")%>");
           return false;
           }
       }

       if (document.data.cand_dismissed_i.value == 1)
        {
           if (document.data.cand_dismissed_m.value == "")
           {
            alert("<% response.write gVText("i_72")%>");
			document.data.cand_dismissed_m.focus();
            return false;
           }
        }
     if (document.data.cand_resigned_i.value == 1)
        {
           if (document.data.cand_resigned_m.value == "")
           {
            alert("<% response.write gVText("i_72")%>");
			document.data.cand_resigned_m.focus();
            return false;
           }
        }
      
		
		if (document.data.cand_nameinclude_i.value == 1)
        {
           if (document.data.cand_nameinclude_UN.value == "")
           {
            alert("<% response.write gVText("i_72")%>");
			document.data.cand_nameinclude_UN.focus();
            return false;
           }
        }
		
        



<%if session("template_org_code") = 1500 then%>
        if (document.data.candtext_title_t.value == "") {
        alert("<% response.write gVText("i_48")%>");
            return false;
        }

<%end if%>

}
</SCRIPT>

</head>

<%if (Request.form("postadding") = 1 AND session("appallok") = 0) then
	if session("template_org_code") = 7000 then
		' IFRC ONLY - IF PERSON APPLYING TO RC POST, then check if they have RC experience as current

			set obj_db_select_CmdIV = server.CreateObject("adodb.command")
			obj_db_select_CmdIV.ActiveConnection = rsys_db_select

		  '<<--Modified by Interface on 05/01/2007
			GETRCEXPsql = " SELECT c.cand_employ1_pres_i FROM td_rsys_cand c WHERE c.cand_id_c = ?"
			obj_db_select_CmdIV.CommandText = GETRCEXPsql
			Set GETRCEXP = obj_db_select_CmdIV.Execute(,Array(session("RSYS_EVAL")))
			'-->>
			' CONTRACT TYPE MUST BE 11 for the IFRC RC experience check
			if GETJOBNAME("jobinfo_contracttype_id_c") = 11 then
				if GETRCEXP("cand_employ1_pres_i") = "0" then%>
	<table>
<TR>
    	<td valign="top" align="center" class="alert" colspan="6"><%=gVText("i_40")%></td>
</tr>
<TR>
      	<td valign="top" align="center" colspan="6">&nbsp;</td>
</tr>
	</table>
			<%response.end
			end if
		end if
	end if
end if

' ---------- AREA WHERE APPLICANT IS TOLD THAT THEIR SECTIONS ARE NOT COMPLETED

if (Request.form("postadding") = 1 AND session("appallok") = 0) then%>
<TABLE cellpadding="0" BORDER="0" bordercolor="yellow" width="<% response.write tablesize%>">
<%else%>
	<%if session("appallok") = 0 then%>
<TABLE cellpadding="0" BORDER="0" bordercolor="green" width="100%" class="tabelPadd">
<TR>
	<td width="100%" valign="top">

<TABLE cellpadding="0" BORDER="0" width="100%">
<TR>
	<td valign="top">&nbsp;</TD>
</tr>
<TR>
	<td valign="top" class="alert"><strong><%=gVText("i_20")%></strong>.</TD>
</TR>
<TR>
	<td valign="top">&nbsp;</TD>
</tr>
<%'<TR>
'	<td valign="top"><strong><i><%=gVText("i_32")></i></strong></TD>
'</TR>
'<TR>
'	<td valign="top">&nbsp;</TD>
'</tr>
%>
<TR>
	<td valign="top"><%=gVText("i_34")%></TD>
</TR>
<TR>
	<td valign="top">&nbsp;</TD>
</tr>
<tr>
	<td>
		<table border="0">
			<tr>
				<td><strong><font color="blue">
<% ' SECTION TO TELL APPLICANT WHICH SECTIONS ARE NOT COMPLETE
' WTO
if session("rsys_intern") = "1" then
	''response.write "<br>T 008 INTERN: " & session("rsys_intern")
	if session("OKA") = 0 then
		response.write gSIDETEXT("i_1") & "<br>"
	end if
	if session("OKAdd") = 0 then
		response.write gSIDETEXT("i_2") & "<br>"
	end if
	if session("OKEmailAdd") = 0 then
		response.write gSIDETEXT("i_51") & "<br>"
	end if
	If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0"then
		response.write gSIDETEXT("i_6") & "<br>"
	end if
	if session("OKC") = 0 then
		response.write gSIDETEXT("i_4") & "<br>"
	end if
	if session("OKD") = 0 then
		response.write gSIDETEXT("i_5") & "<br>"
	end if
	'06 JUN 12 LJL no Other Info section required for IFRC interns
	if session("template_org_code") = 3000 then
		if session("OKOI") = 0 then
			response.write gSIDETEXT("i_73") & "<br>"
		end if
	end if
else
	if session("template_org_code") = 1000 then

		if session("OKA") = 0 then
			response.write gSIDETEXT("i_1") & "<br>"
		end if
		if session("OKAdd") = 0 then
			response.write gSIDETEXT("i_2") & "<br>"
		end if
		if session("OKEmailAdd") = 0 then
			response.write gSIDETEXT("i_51") & "<br>"
		end if
		' AREAS OF EXPERTISE
		if session("OKS") = 0 then
			response.write gSIDETEXT("i_52") & "<br>"
		end if

	'06 MAR 11 LJL added check on internal staff for comptencies check on 1000 WHO complete
	if session("CLI_INTERNAL_STAFF") > 0 then
		' COMPETENCIES
		if session("OKRMCmp") = 0 then
			response.write gSIDETEXT("i_58") & "<br>"
		end if
		if session("OKRMPref") = 0 then
			response.write gSIDETEXT("i_59") & "<br>"
		end if
	end if

		if session("OKB") = 0 then
			response.write gSIDETEXT("i_3") & "<br>"
		end if
		if session("employcheck") = 0 then
			response.write gSIDETEXT("i_8") & "<br>"
		end if

		if session("OKC") = 0 then
			response.write gSIDETEXT("i_4") & "<br>"
		end if
		if session("OKIE") = 0  then
			response.write gSIDETEXT("i_7") & "<br>"
		end if

		If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0"then
			response.write gSIDETEXT("i_6") & "<br>"
		end if
		
		if session("OKD") = 0 then
			response.write gSIDETEXT("i_5") & "<br>"
		end if

		if session("OKG") = 0 then
			response.write gSIDETEXT("i_9") & "<br>"
		end if

		if session("refcheck") = 0 then
			response.write gSIDETEXT("i_63") & "<br>"
		end if


'ILO SETTINGS
	elseif session("template_org_code") = 2000 then

		if session("OKA") = 0 then
			response.write gSIDETEXT("i_1") & "<br>"
		end if
		if session("OKAdd") = 0 then
			response.write gSIDETEXT("i_2") & "<br>"
		end if
		if session("OKEmailAdd") = 0 then
			response.write gSIDETEXT("i_51") & "<br>"
		end if
		if session("OKB") = 0 then
			response.write gSIDETEXT("i_3") & "<br>"
		end if

		if session("employcheck") = 0 then
			response.write gSIDETEXT("i_8") & "<br>"
		end if

If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0"then
			response.write gSIDETEXT("i_6") & "<br>"
		end if
		if session("OKIE") = 0  then
			response.write gSIDETEXT("i_7") & "<br>"
		end if

		if session("OKC") = 0 then
			response.write gSIDETEXT("i_4") & "<br>"
		end if
		if session("OKD") = 0 then
			response.write gSIDETEXT("i_5") & "<br>"
		end if

		if session("OKG") = 0 then
			response.write gSIDETEXT("i_9") & "<br>"
		end if
		if session("refcheck") = 0 then
			response.write gSIDETEXT("i_63") & "<br>"
		end if

'WTO SETTINGS
	elseif session("template_org_code") = 3000 then

		if session("OKA") = 0 then
			response.write gSIDETEXT("i_1") & "<br>"
		end if
		if session("OKAdd") = 0 then
			response.write gSIDETEXT("i_2") & "<br>"
		end if
		if session("OKEmailAdd") = 0 then
			response.write gSIDETEXT("i_51") & "<br>"
		end if
If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0"then
			response.write gSIDETEXT("i_6") & "<br>"
		end if
		if session("OKC") = 0 then
			response.write gSIDETEXT("i_4") & "<br>"
		end if
		if session("OKD") = 0 then
			response.write gSIDETEXT("i_5") & "<br>"
		end if
		if session("OKOI") = 0 then
			response.write gSIDETEXT("i_73") & "<br>"
		end if

		if session("employcheck") = 0 then
			response.write gSIDETEXT("i_8") & "<br>"
		end if
		if session("OKG") = 0 then
			response.write gSIDETEXT("i_9") & "<br>"
		end if
		if session("refcheck") = 0 then
			response.write gSIDETEXT("i_63") & "<br>"
		end if
		if session("OKIE") = 0  then
			response.write gSIDETEXT("i_7") & "<br>"
		end if

'20 AUG 10 LJL IFRC section
	elseif session("template_org_code") = 7000 then

		if session("OKA") = 0 then
			response.write gSIDETEXT("i_1") & "<br>"
		end if
		if session("OKAdd") = 0 then
			response.write gSIDETEXT("i_2") & "<br>"
		end if

		if session("OKEmailAdd") = 0 then
			response.write gSIDETEXT("i_51") & "<br>"
		end if

		' AREAS OF EXPERTISE
		if session("OKS") = 0 then
			response.write gSIDETEXT("i_52") & "<br>"
		end if

' Other skills
		if session("OKOS") = 0 then
			response.write gSIDETEXT("i_66") & "<br>"
		end if

' Edu
	If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0"then
			response.write gSIDETEXT("i_6") & "<br>"
		end if

' contract
		if session("OKB") = 0 then
			response.write gSIDETEXT("i_3") & "<br>"
		end if

' employ
		if session("employcheck") = 0 then
			response.write gSIDETEXT("i_8") & "<br>"
		end if

' language
		if session("OKC") = 0 then
			response.write gSIDETEXT("i_4") & "<br>"
		end if

' int experience
		if session("OKIE") = 0  then
			response.write gSIDETEXT("i_7") & "<br>"
		end if

' RC/RC
		if session("OKRC") = 0 then
			response.write gSIDETEXT("i_64") & "<br>"
		end if

' RCOther
		if session("OKRCO") = 0 then
			response.write gSIDETEXT("i_68") & "<br>"
		end if

' Computer
		if session("OKD") = 0 then
			response.write gSIDETEXT("i_5") & "<br>"
		end if

' Family
		if session("OKG") = 0 then
			response.write gSIDETEXT("i_9") & "<br>"
		end if

' References
		if session("refcheck") = 0 then
			response.write gSIDETEXT("i_63") & "<br>"
		end if


' ITU UPDATES - includes competenecies
elseif session("template_org_code") = 2400 then

		if session("OKA") = 0 then
			response.write gSIDETEXT("i_1") & "<br>"
		end if
		if session("OKAdd") = 0 then
			response.write gSIDETEXT("i_2") & "<br>"
		end if
		if session("OKEmailAdd") = 0 then
			response.write gSIDETEXT("i_51") & "<br>"
		end if
		' AREAS OF EXPERTISE
		if session("OKS") = 0 then
			response.write gSIDETEXT("i_52") & "<br>"
		end if

		' COMPETENCIES
		if   session("OKRMCmp") = 0 then
			response.write gSIDETEXT("i_58") & "<br>"
		end if

		if session("OKB") = 0 then
			response.write gSIDETEXT("i_3") & "<br>"
		end if
		if session("employcheck") = 0 then
			response.write gSIDETEXT("i_8") & "<br>"
		end if

		if session("OKC") = 0 then
			response.write gSIDETEXT("i_4") & "<br>"
		end if
		if session("OKIE") = 0  then
			response.write gSIDETEXT("i_7") & "<br>"
		end if

If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0"then
			response.write gSIDETEXT("i_6") & "<br>"
		end if
		if session("OKD") = 0 then
			response.write gSIDETEXT("i_5") & "<br>"
		end if

		if session("OKG") = 0 then
			response.write gSIDETEXT("i_9") & "<br>"
		end if

		if session("refcheck") = 0 then
			response.write gSIDETEXT("i_63") & "<br>"
		end if

' WMO UPDATES - excludes competenecies
elseif session("template_org_code") = 2900 then

		if session("OKA") = 0 then
			response.write gSIDETEXT("i_1") & "<br>"
		end if
		if session("OKAdd") = 0 then
			response.write gSIDETEXT("i_2") & "<br>"
		end if
		if session("OKEmailAdd") = 0 then
			response.write gSIDETEXT("i_51") & "<br>"
		end if
		' AREAS OF EXPERTISE
		if session("OKS") = 0 then
			response.write gSIDETEXT("i_52") & "<br>"
		end if

		if session("OKB") = 0 then
			response.write gSIDETEXT("i_3") & "<br>"
		end if
		if session("employcheck") = 0 then
			response.write gSIDETEXT("i_8") & "<br>"
		end if

		if session("OKC") = 0 then
			response.write gSIDETEXT("i_4") & "<br>"
		end if
		if session("OKIE") = 0  then
			response.write gSIDETEXT("i_7") & "<br>"
		end if

If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0"then
			response.write gSIDETEXT("i_6") & "<br>"
		end if
		if session("OKD") = 0 then
			response.write gSIDETEXT("i_5") & "<br>"
		end if

		if session("OKG") = 0 then
			response.write gSIDETEXT("i_9") & "<br>"
		end if

		if session("refcheck") = 0 then
			response.write gSIDETEXT("i_63") & "<br>"
		end if
		'03 SEP 15 LJL added OI check for WMO
		if session("OKOI") = 0 then
			response.write gSIDETEXT("i_73") & "<br>"
		end if

' ALL OTHER ORG UPDATES
	else

		if session("OKA") = 0 then
			response.write gSIDETEXT("i_1") & "<br>"
		end if
		if session("OKAdd") = 0 then
			response.write gSIDETEXT("i_2") & "<br>"
		end if
		if session("OKEmailAdd") = 0 then
			response.write gSIDETEXT("i_51") & "<br>"
		end if
		' AREAS OF EXPERTISE
		if session("OKS") = 0 then
			response.write gSIDETEXT("i_52") & "<br>"
		end if

		' COMPETENCIES
		if   session("OKRMCmp") = 0 then
			response.write gSIDETEXT("i_58") & "<br>"
		end if

		if session("OKB") = 0 then
			response.write gSIDETEXT("i_3") & "<br>"
		end if
		if session("employcheck") = 0 then
			response.write gSIDETEXT("i_8") & "<br>"
		end if

		if session("OKC") = 0 then
			response.write gSIDETEXT("i_4") & "<br>"
		end if
		if session("OKIE") = 0  then
			response.write gSIDETEXT("i_7") & "<br>"
		end if

If session("OKE") = "0" OR session("educheck") = "0" OR session("educheck2") = "0"then
			response.write gSIDETEXT("i_6") & "<br>"
		end if
		if session("OKD") = 0 then
			response.write gSIDETEXT("i_5") & "<br>"
		end if

		if session("OKG") = 0 then
			response.write gSIDETEXT("i_9") & "<br>"
		end if

		if session("refcheck") = 0 then
			response.write gSIDETEXT("i_63") & "<br>"
		end if

	end if%>

</strong></font>
	</td>
	</tr>
	</table>

</td>
</tr>

<TR>
	<td valign="top">&nbsp;</TD>
</tr>
<TR>
	<td valign="top">&nbsp;</TD>
</tr>
<TR>
	<td valign="top"><%=gVText("i_33")%></TD>
</TR>

<TR>
	<td valign="top">&nbsp;</TD>
</tr>
</table>
<%end if%>
<%'22 FEB 16 LJL moved this end to stop big line breaks on top of applying page
%>
</td>
      </tr>
    </table>


<%end if%>


<%if session("appallok") = 0 then
		response.end
    end if
    if Request.form("postadding") = 1 then
    goform = "hrd-cl-vac-app.asp"
    else
    goform = "AppV-edit.asp"
    end if
    %>
<form action="<%=goform%>" method="POST" name="data" ONSUBMIT="return DataValidation();" class="language">
<%'response.write "PA:" & request.form("postadding")%>
<TABLE width="100%" border="0" bordercolor="blue" cellpadding="1" cellspacing="0" bgcolor="#ffffcc">
<%'      <!------------ ILO ADD 21 OCT 02 --------------->
if Request.form("postadding") = "1" then
	' NEEDNEEDNEED add this section
	'Logsssql = "  "& getjobname("jobinfo_a") & " 		"& ParseDateTime(Now()) &", 		" & session("template_org_name") & "-APPLY | JOBID " & Request.form("jobinfo_uid_c") & " CLOSE "& date(getjobname("jobinfo_acl_d"), "d-mmm-yy") &"| CANDID " & session("RSYS_EVAL") & "', 		'" & session("template_org_name") & "-APPLY', 		" & Request.form("jobinfo_uid_c") & ", 		" & session("RSYS_EVAL") & ", 		'"& left("REMOTE_ADDR","15")&"' 	) 	"") & "
	'set Logss = rsys_logs.execute(Logsssql)
else
	
	'01 SEP 15 LJL moved WMO law section to Other info section
if  (session("template_org_code") <> 3000 AND session("template_org_code") <> 2900) AND NOT Request.form("GOEDITV") = "99" then
%>

<tr>
	<td valign="top" class="text" colspan="2"><%=gVText("i_11")%>
	<br><br>
    <select  id="cand_law_i" name="cand_law_i">
    <%
    //Modified on 08/13/2008 , to display previously selected answer.
    'if len( trim(JAPY("cand_law_i")) ) then ///>
    if trim(JAPY("cand_law_i")) >= 0 then %>
    <option value=""><%=gVText("i_71")%>
    <option value="1" <% If trim(JAPY("cand_law_i")) = "1" then%> SELECTED<% End If%> ><%=gVText("i_18")%>
    <option value="0" <% If trim(JAPY("cand_law_i")) = "0" then%> SELECTED<% End If%> ><%=gVText("i_19")%>
    <%else%>
    <option value="" SELECTED><%=gVText("i_71")%>
    <option value="1"><%=gVText("i_18")%>
    <option value="0"><%=gVText("i_19")%>
    <%end if%>
    </select></TD>
</TR>
<tr>
    <td valign="top" colspan="2"><br><span id="span_law_anw"><%=gVText("i_12")%></span><Br>
    <textarea  wrap="soft" ROWS="3" NAME="cand_law_m" COLS="50"><% response.write JAPY("cand_law_m")%></TEXTAREA><br><br></td>
</tr>
<%end if


'<!---  manoj added info for Dismissed  -->
    if  (session("template_org_code") <> 3000 AND session("template_org_code") <> 2900) AND NOT Request.form("GOEDITV") = "99" then
%>

<tr>
	<td valign="top" class="text" colspan="2"><%=gVText("i_94")%>
	<br><br>
    <select  id="cand_dismissed_i" name="cand_dismissed_i">
    <%
    
    'if len( trim(JAPY("cand_law_i")) ) then ///>
    if trim(JAPY("cand_dismissed_i")) >= 0 then %>
    <option value=""><%=gVText("i_71")%>
    <option value="1" <% If trim(JAPY("cand_dismissed_i")) = "1" then%> SELECTED<% End If%> ><%=gVText("i_18")%>
    <option value="0" <% If trim(JAPY("cand_dismissed_i")) = "0" then%> SELECTED<% End If%> ><%=gVText("i_19")%>
    <%else%>
    <option value="" SELECTED><%=gVText("i_71")%>
    <option value="1"><%=gVText("i_18")%>
    <option value="0"><%=gVText("i_19")%>
    <%end if%>
    </select></TD>
</TR>
<tr>
    <td valign="top" colspan="2"><br><span id="span1"><%=gVText("i_12")%></span><Br>
    <textarea  wrap="soft" ROWS="3" NAME="cand_dismissed_m" COLS="50"><% response.write JAPY("cand_dismissed_m")%></TEXTAREA><br><br></td>
</tr>
<%end if


'dismissed reason end here

    '<!---  manoj added info for resigned  -->
    if  (session("template_org_code") <> 3000 AND session("template_org_code") <> 2900) AND NOT Request.form("GOEDITV") = "99" then
%>

<tr>
	<td valign="top" class="text" colspan="2"><%=gVText("i_95")%>
	<br><br>
    <select  id="cand_resigned_i" name="cand_resigned_i">
    <%
    
    'if len( trim(JAPY("cand_law_i")) ) then ///>
    if trim(JAPY("cand_resigned_i")) >= 0 then %>
    <option value=""><%=gVText("i_71")%>
    <option value="1" <% If trim(JAPY("cand_resigned_i")) = "1" then%> SELECTED<% End If%> ><%=gVText("i_18")%>
    <option value="0" <% If trim(JAPY("cand_resigned_i")) = "0" then%> SELECTED<% End If%> ><%=gVText("i_19")%>
    <%else%>
    <option value="" SELECTED><%=gVText("i_71")%>
    <option value="1"><%=gVText("i_18")%>
    <option value="0"><%=gVText("i_19")%>
    <%end if%>
    </select></TD>
</TR>
<tr>
    <td valign="top" colspan="2"><br><span id="span2"><%=gVText("i_12")%></span><Br>
    <textarea  wrap="soft" ROWS="3" NAME="cand_resigned_m" COLS="50"><% response.write JAPY("cand_resigned_m")%></TEXTAREA><br><br></td>
</tr>
<%end if


'resigned reason end here


     '<!---  manoj added info for name include in un  -->
    if  (session("template_org_code") <> 3000 AND session("template_org_code") <> 2900) AND NOT Request.form("GOEDITV") = "99" then
%>

<tr>
	<td valign="top" class="text" colspan="2"><%=gVText("i_96")%>
	<br><br>
    <select  id="cand_nameinclude_i" name="cand_nameinclude_i">
    <%
    
    'if len( trim(JAPY("cand_law_i")) ) then ///>
    if trim(JAPY("cand_nameinclude_i")) >= 0 then %>
    <option value=""><%=gVText("i_71")%>
    <option value="1" <% If trim(JAPY("cand_nameinclude_i")) = "1" then%> SELECTED<% End If%> ><%=gVText("i_18")%>
    <option value="0" <% If trim(JAPY("cand_nameinclude_i")) = "0" then%> SELECTED<% End If%> ><%=gVText("i_19")%>
    <%else%>
    <option value="" SELECTED><%=gVText("i_71")%>
    <option value="1"><%=gVText("i_18")%>
    <option value="0"><%=gVText("i_19")%>
    <%end if%>
    </select></TD>
</TR>
<tr>
    <td valign="top" colspan="2"><br><span id="span3"><%=gVText("i_12")%></span><Br>
    <textarea  wrap="soft" ROWS="3" NAME="cand_nameinclude_UN" COLS="50"><% response.write JAPY("cand_nameinclude_UN")%></TEXTAREA><br><br></td>
</tr>
<%end if


'resigned reason end here




end if
'      <!------------ ILO ADD 21 OCT 02 --------------->
if Request.form("postadding") = 1 then
	  	' response.write "T POST NAME aND REC" & GETJOBNAME.eof
	  	' REMOVED 08 MAY 07
		'getjobname.movefirst
Do while getjobname.eof = false
' INFO - SHOW THE POST NAME AND NUMBER%>
<TR>
	<td valign="top" colspan="2" align="center">
	<%=gVText("i_13")%>
    <Br>
    <span class="textbold"><% response.write jobtitle%></span>
<%if len(GETJOBNAME("jobinfo_vac2_c")) then%>- <% response.write GETJOBNAME("jobinfo_vac2_c")
end if%>
<br>
<br></td>
</tr>
<%
'**************************************************************************************************
' UNAIDS MOTIVATION LETTER REQUIRED
'03 FEB 13 LJL IFRC Also requires covering letter

'**************************************************************************************************
if session("template_org_code") = 1500 OR session("template_org_code") = 7000 then%>
<tr>
	<td colspan="2">
<TABLE cellpadding="0" BORDER="0" width="100%" bgcolor="white">
    <tr>
      	<td valign="bottom"  align='center'><font color="blue"><% response.write gVText("i_44")%></font></td>
    </tr>
    <TR>
      	<td valign='top'>&nbsp;</TD>
    </TR>
    <tr>
      	<td valign="bottom"  align='center'><% response.write gVText("i_45")%></td>
    </tr>
    <TR>
      	<td valign='top'>&nbsp;</TD>
    </TR>
    <tr>
      	<td valign="bottom"  align='center'><strong><% response.write gVText("i_46")%></strong></td>
    </tr>
    <TR>
      	<td valign='top'  align='center'><INPUT size="50" TYPE="text" maxlength="98" NAME="candtext_title_t" value=""></td>
    </TR>
    <TR>
      	<td valign='top'  align='center'>&nbsp;</TD>
    </TR>
    <tr>
      	<td valign='top'  align='center'><strong><% response.write gVText("i_47")%></strong></td>
      </tr>
      <TR>
        	<td valign='top'  align='center'>
          	<textarea  wrap="soft" NAME="candtext_text_en_m" cols="80" rows="10"></textarea>
			<br>&nbsp;<br>
          <INPUT TYPE="hidden" VALUE="13" name="candtext_type_c">
          <INPUT TYPE="hidden" VALUE="88" name="TextAdd">
          </td>
        </tr>
      </TABLE>
	  </td>
</tr>
<%end if
' **************************************************************************************************
' END UNAIDS MOTIVATION LETTER REQUIRED
'03 FEB 13 LJL IFRC Also requires covering letter

'******************************************************************************************************
%>
<tr bgcolor="#ffccff">
	<td valign="middle" align="left" bgcolor="#ffe4e1"><%=gVText("i_14")%></td>
    <td valign="middle" bgcolor="#ffe4e1">&nbsp; <select class="yesno" name="postaddselect">
    <option value="1"><%=gVText("i_18")%>
    <option value="0" SELECTED><%=gVText("i_19")%>
    </select> &nbsp; </TD>
</TR>



 <!-- manoj td_rsys_jobinfo.jobinfo_indic_avail_i = 1  and   jobinfo_thisorg_1000 = 1 -->

 
 <% if GETJOBNAME("jobinfo_thisorg_1000") = "1" and GETJOBNAME("jobinfo_indic_avail_i")="1"  then %>


 <tr bgcolor ="#FFFFFF">
 <td valign="top">
 </td>
 <td></td>
 
 </tr>
 <%'22 Feb 16 LJL/MC Add emergency resource avail dates for WHO specific VNs%>
 <tr  bgcolor ="#FFFFFF">
  <td valign="top"><strong><%'Kindly indicate available start and end dates for vacancies%><%=gVText("i_82")%></strong> </td>
  <td></td>
 </tr>
 <tr>
 <td>
 <table>
 <tr>
 <td>
 <strong><%'Start Date for vacancies   : %><%=gVText("i_83")%></strong>  &nbsp;
 </td>
 <td><input type="text" name="startdate" size="12"  maxlength="15" value="" id="datepicker" readonly></td>
 </tr>
 </table>
 
 </td>
 
 </tr>
 <tr>
 <td>
 <table>
 <tr>
 <td>
 <strong><%'End Date for vacancies   : %><%=gVText("i_84")%></strong>  &nbsp;
 </td>
 <td><input type="text" name="Enddate" size="12" maxlength="15" value="" id="datepicker1" readonly ></td>
 </tr>
 </table>
 
 </td>
 
 
 </tr>

  <%end if %>

  
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
                    yearRange: '-0:+20',
                    onChangeMonthYear: function (y, m, i) {
                        var d = i.selectedDay;
                        $(this).datepicker('setDate', new Date(y, m - 1, d));
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
                    yearRange: '-0:+20',
                    onChangeMonthYear: function (y, m, i) {
                        var d = i.selectedDay;
                        $(this).datepicker('setDate', new Date(y, m - 1, d));
                    }

                });
            });
        </script>
<%end if%>

  <!-- td_rsys_jobinfo.jobinfo_indic_avail_i = 1  and   jobinfo_thisorg_1000 = 1 -->



<%getjobname.movenext
loop

' **************************************************************************************************
'14 OCT 09 LJL ADD VN FAMILIAR FEATURE - currently only for ILO and IFRC ---------------------------------------->
' **************************************************************************************************
'29 APR 15 LJL add VN sourcing for WMO
if session("template_org_code") = 2000 OR session("template_org_code") = 7000 OR session("template_org_code") = 2900 then%>
<TR>
	<td colspan='2'>&nbsp;</TD>
</TR>
<TR>
	<td colspan='2'><% response.write gVText("i_80")%> <FONT SIZE='4' COLOR='Red'>*</FONT></TD>
</TR>
<TR>
	<td colspan='2'>
       <select  name="candjob_webfamiliar_id_c">
      <option value=''><% response.write gVText("i_72")%></option>
      <%Do while getfamiliars.eof = false%>
      <option value='<%=getfamiliars("webfamiliar_id_c")%>'><%=getfamiliars("familiardsc")%></option>
        <%getfamiliars.movenext
        loop%>
        </select></td>
</TR>
<TR>
	<td colspan='2'><% response.write gVText("i_81")%></TD>
</TR>
<TR>
	<td colspan='2'><INPUT size="45" maxlength="248" TYPE="text" NAME="candjob_webfamiliar_refer_t" VALUE="" SIZE="48" MAXLENGTH="98"></TD>
</TR>
<TR>
	<td colspan='2'>&nbsp;</TD>
</TR>
<%end if
' **************************************************************************************************
'14 OCT 09 LJL END ADD VN FAMILIAR FEATURE - currently only for ILO and IFRC ---------------------------------------->
' **************************************************************************************************
%>
</TABLE>
<%questionnumber = 1
	'      <!------------ KEEP THIS HERE AS IS DEPENDENT ON JOBINFO_UID_C inclusion -------------------->
  '      end if
  '<<--Modified by Interface on 05/01/2007

    set obj_db_select_CmdV = server.CreateObject("adodb.command")
    obj_db_select_CmdV.ActiveConnection = rsys_db_select

	GetAnssql = " SELECT j.jobinfo_uid_c, q.questions_dsc_" & session("lng") & "_t AS qstdsc, q.questions_id_c FROM tx_rsys_qj j INNER JOIN td_rsys_questions q ON j.questions_id_c = q.questions_id_c WHERE (q.questions_inactind_i = 0) AND j.jobinfo_uid_c = ? ORDER BY q.questions_order_c"
	obj_db_select_CmdV.CommandText = GetAnssql
	Set GetAns = obj_db_select_CmdV.Execute(,Array(Request.form("jobinfo_uid_c")))
	'-->>%>
 <table>


 <TR>
    <td valign="top">&nbsp;</td>
</tr>
<%if GETANS.eof = false then%>
      <input type="hidden" name="AnswerQuestions" value="1">
<%qcounter = 1%>
<TR>
	<td valign="top"><%=gQText("i_5")%></td>
</tr>
</table>
<TABLE>
<%'<TR>
'	<td valign="top" bgcolor="yellow"><font color=RED><strong>THIS PAGE IS BEING REVISED, it will be operational shortly.</font></strong></td>
'</tr>%>
<%question_counter = 0
      Do while NOT GETANS.eof
question_counter = question_counter + 1
      qstid = GetAns("questions_id_c")%>
<input type="hidden" name="ID" value="<% response.write qcounter%>">
<input type="hidden" name="qid<% response.write qcounter%>" value="<% response.write qstid%>">
<TR>
	<td valign="top"  colspan="10"><% response.write questionnumber%>. <b><% response.write GetAns("qstdsc")%></b></td>
</tr>
<%'<<--Modified by Interface on 05/01/2007

set obj_db_select_aJobEditCmdV = server.CreateObject("adodb.command")
obj_db_select_aJobEditCmdV.ActiveConnection = rsys_db_select
GETANSWERSsql = " SELECT qa.qanswer_id_c, qa.qanswer_dsc_" & session("lng") & "_t AS qadsc, qa.qanswer_right_i, qa.qanswer_text_i  FROM dbo.tr_rsys_qanswer qa  INNER JOIN dbo.td_rsys_questions q ON qa.questions_id_c = q.questions_id_c WHERE q.questions_id_c = ? ORDER BY qa.qanswer_rank_c, qa.qanswer_id_c "
obj_db_select_aJobEditCmdV.CommandText = GETANSWERSsql
Set GetANSWERS = obj_db_select_aJobEditCmdV.Execute(,Array(qstid))

'GETAnswerssql = " SELECT qa.qanswer_id_c, qa.qanswer_dsc_" & session("lng") & "_t AS qadsc, qa.qanswer_right_i, qa.qanswer_text_i  FROM dbo.tr_rsys_qanswer qa  INNER JOIN dbo.td_rsys_questions q ON qa.questions_id_c = q.questions_id_c WHERE q.questions_id_c = ? ORDER BY qa.qanswer_rank_c, qa.qanswer_id_c "
'response.write GETANSWERSSQL
'response.end
'obj_db_select_Cmd.CommandText = GETAnswerssql
'Set GetAnswers = obj_db_select_Cmd.Execute(,Array(qstid))

'-->>
qrowcount = 0

Do while GETANSWERS.eof = false
pv_qanswerid = GETANSWERS("qanswer_id_c")%>
<%if qrowcount = 0 then%>
<tr>
<%end if%>
<td valign="top"  colspan="1">
	<%'="Q:" & qstid & "QA:" & pv_qanswerid%>
		<input type="radio" id="qst<% response.write qcounter%>" class="radio-field" name="qst<% response.write qcounter%>" value="<% response.write pv_qanswerid%>"><% response.write GetAnswers("qadsc")
	if GetAnswers("qanswer_text_i") = "1" then%>
    <br>
    <%=gQText("i_20")%>
    <br>
    <input TYPE="text" id="txt<% response.write pv_qanswerid%>" class ="text-field"  name="txt<% response.write pv_qanswerid%>" value="" maxlength="253">
    <%'if GetAnswers("qanswer_id_c") = getjapanswers("qanswer_id_c") then   	'response.write GETJAPanswers("candquestion_other_t")  'end if%>
<%else%>
		<input type="hidden" id="hdn<% response.write pv_qanswerid%>" name="txt<% response.write pv_qanswerid%>" value="" maxlength="253">
<%end if%>
	</td>
<%qrowcount = qrowcount + 1
if qrowcount >= 5 then%>
</tr>
      <%if qrowcount > 4 then
      qrowcount = 0
      end if
      end if
      GETANSWERS.movenext
      loop
      if qrowcount <= 4 then%>
    </tr>
 <%end if%>
<TR>
	<td valign="top" colspan="10"><hr size="1"></td>
</tr>
<%qcounter = qcounter + 1
    questionnumber = questionnumber + 1
	qstid = ""
GETANS.movenext
loop
qcounter = qcounter - 1%>
<INPUT TYPE="hidden" NAME="question_counter" VALUE="<% response.write question_counter%>">
<INPUT TYPE="hidden" NAME="qcounter" VALUE="<% response.write qcounter%>">
<%else%>
<TR>
      	<td valign="top" class="textitalic" colspan="10"><%=gQText("i_14")%></td>
</tr><%end if
if japinfo2("EditQ") >=0 then
	Qcount = int(japinfo2("EditQ")+1)
else 
	Qcount = 0
end if%>
<INPUT TYPE="hidden" NAME="editQ" VALUE="<%=Qcount%>">
</table>
<%end if%>
<table>
<TR>
    <td valign="top" colspan="2">&nbsp;</TD>
</TR>
<%if len(request.form("GOEDITV")) then
' BEGIN not showing the verif page if the person has just submitted it.
else%>
<TR>
<%' CERTIFICATION STATEMENT %>
    <td valign="top" class="text" colspan="2" bgcolor="#ffe4e1"><%=gVText("i_text1")%> </TD>
</TR>
<TR>
    <td valign="top" colspan="2">&nbsp;</TD>
</TR>
<TR>
    <td valign="top"  colspan="2"><%=gVText("i_text2")%></TD>
</TR>
<TR>
    <td valign="top" colspan="2">&nbsp;</TD>
</TR>
<TR>
    <td valign="top" colspan="2">&nbsp;</TD>
</TR>
<TR>
	<td valign="top" width="30%"><%=gVText("i_3")%> </td>
    <td valign="top" width="70%"><INPUT size="15" TYPE="hidden" NAME="cand_fil_d"><b><%response.write day(now())  & "-" & monthname(month(now())) & "-" & year(now())%></b></TD>
</TR>
<TR>
<%if session("template_org_code") <> 3000 then
if Request.querystring("editFooter") <> "" then%>
	<td valign="top" align="left" ><%=gVText("i_4")%> <font size="4" color="red">*</font></td>
	<td valign="top" ><% response.write JAPY("cand_fil_t")%><INPUT size="15" TYPE="hidden" NAME="cand_fil_t" maxlength="98" VALUE="<% response.write JAPY("cand_fil_t")%>"></TD>
<%else%>
	<td valign="top" align="left" ><%=gVText("i_4")%> <font size="4" color="red">*</font></td>
	<td valign="top" ><input TYPE="text"  NAME="cand_fil_t" VALUE="" size="30" maxlength="98"><input type="hidden" name="cand_fil_t_required" value="<%=gVText("i_9")%>"></TD>
<%end if
	
	
else%>
	<td><INPUT size="15" TYPE="hidden" NAME="cand_fil_t" VALUE="Not indicated-WTO"></td>
<%end if%>
</TR>
<TR>
	<td valign="top" ><%=gVText("i_7")%> <font size="4" color="red">*</font></td>
<%if Request.querystring("editFooter") <> "" then%>
	<td valign="top" ><% response.write JAPY("cand_verif_name_t")%></TD>
    <INPUT size="15" TYPE="hidden" NAME="cand_verif_name_t" VALUE="<% response.write JAPY("cand_verif_name_t")%>" maxlength="73">


<%else%>
	<td valign="top"><input size="30" maxlength="73" TYPE="text" NAME="cand_verif_name_t" VALUE=""></TD>
	<input type="hidden" name="cand_verif_name_t_required" value="<%=gVText("i_8")%>">
<%end if%>
</tr>
<%'    <!------------ ILO ADD 21 OCT 02 --------------->
if Request.form("postadding") = 1 then%>
    	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<% response.write Request.form("jobinfo_uid_c")%>">
<%end if
Signcount = int(japinfo2("EditFooter") + 1)%>
	<INPUT TYPE="hidden" NAME="editFooter" VALUE="<% response.write Signcount%>">
	<INPUT TYPE="hidden" NAME="GOEDITV" VALUE="99">
<TR>
	<td valign="top">&nbsp;</TD>
</tr>
<tr valign="bottom">
	<td valign="bottom" onclick="" align="center" colspan="2">
    
<INPUT  TYPE="submit" VALUE="<%if Request.form("postadding") = 1 then
	response.write gVText("i_15")
	else
	response.write gVText("i_5")
	end if%>"></form></td>
</tr>
<%end if ' END not showing the verif page if the person has just submitted it.
end if%>
</TABLE>
<%pv_last_update="26 Mar 2023"
'<<--Modified by Interface on 05/29/2007
set obj_int_select_Cmd = nothing
set obj_db_select_Cmd = nothing

set obj_db_select_CmdA = nothing
set obj_db_select_CmdAI = nothing
set obj_db_select_CmdAII = nothing

set obj_db_select_CmdI = nothing
set obj_db_select_CmdII = nothing
set obj_db_select_CmdIII = nothing
set obj_db_select_CmdIV = nothing
set obj_db_select_CmdV = nothing

set obj_db_Cmd = nothing
set obj_logs_CmdI = nothing
set obj_db_CmdI = nothing
set obj_db_CmdII = nothing
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
' *****************************************************
%>
<%//Added on 06/02/2009,DD, following file included for checkbox updatation changes.%>
<!--#include file="../includes/include_pubedit_frame_bottom.asp"-->
<%
' *****************************************************
' END BOTTOM INCLUDES
' *****************************************************
%>

