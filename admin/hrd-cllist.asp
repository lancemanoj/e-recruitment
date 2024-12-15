<% option explicit
'<!-------------------------------------------NOTES ----------
'BE CAREFUL WITH COMPARISONS, INT() TO INT() as sometimes values are set as char by the system automatically

'MODS --
'6 FEB 04 LJL added UN (intl org) listing also
'6 FEB 04 LJL added ability to pass params in links for screening shortlisting etc in order to keep list as it was when they did the action. - keep lists short
'				Significantly changed the queries by compacting into one.
'18 FEB 04 LJL added ability to select line numbers to select for long lists and PDF's
'27 JUL 04 LJL changed selected, screened, shlisted to make it more optimized.  Was causing time outs constantly.
'8-9 AUG 04 LJL updating of the ranking so that it does not timeout as easily.
'24 AUG 04 LJL removed a bunch of things from the queries
'14 SEPT 04 LJL made queries from datasource which is select only
'6 NOV 04 LJL added eligible applicants link for ILO
'18 NOV 04 LJL revise country quota output for ILO
'1 FEB 05 LJL added non-eligible for list - ILO
'17 FEB 05 LJL added recommended section per ANDERSENT
'23 FEB 05 LJL org 1000 2000 3000 4000
'08 MAR 05 LJL revised test column to allow for separate email to testing candidates.
'09 MAR 05 LJL added questions and answers column per WTO req 37
'11 MAR 05 LJL added compare on dates for last updates and then shortlist, screened, application to post, per WTO req 31b
'12 MAR 05 LJL removed test columns per WTO req 35
'12 MAR 05 LJL comms only in English per WTO req 38
'10 MAY 05 LJL reduced out the inactive corres from the list of items to send or add to the applicants for communications
'07 JUN 05 LJL added ability to add posts to a list of applicants via MoveToPosts
'08 JUN 05 LJL added ability to evaluate applicants all at once.
'20 JUN 05 LJL added [Upd!] to each column that indicates when an applicant has updated their profile after the application to a post, screened, shlisted - may change also to a different color.
'20 JUL 05 LJL changed 	"AND ((NOT staff_nbr_response.write session("template_org_code") IS NULL) OR (cand_thisorg_short_i_ response.write session("template_org_code") = 1) OR (is_thisorg_staff > 1))" from staff_nbr_ response.write session("template_org_code") > 1
'22 JUL 05 LJL modified term [Tested] to [To be Tested] in links for listings of applicants
'19 JAN 06 LJL took out nat quota indicator for IFRC 7000
'17 FEB 06 LJL some small formatting modifs
'09 APR 06 LJL changed "select these" function for selecting only certain applicants
'07 Jul 06 LJL wording on PDF select button just above cand list numbering
'09 JUL 06 LJL 	' HAD TO REVISE THE TERM ELIGIBLE BECAUSE IT IS CONTAINED IN NONELIGIBLE
'12 SEP 06 LJL confusion on crd_d = date that admin sets for adding the VN to applicant, upd_d = last date that anything happens to candjob line (anything, selection, etc), candjob_truedate_d = date that the applicant is truly added to the VN
' (above done on cand-list5.asp)
'12 OCT 06 LJL fixed average age problem
'20 OCT 06 LJL defaulted Email? for online testing to checked as it is easier that way
'02 NOV 06 LJL checked what vars are passed to PDF maker
'02 NOV 06 LJL pass jobinfo id to admin_CVPDF_docsettings.asp so that it goes to reduce covering letters for the post, for PDF output, etc
' 03 NOV 06 LJl added check to see if candlist1 and 2 are numeric. If not, don't allow. (to check mark applicants in a list)
'17 Nov 06 ac added flush buffer to stop the script crashing (also in stage and unaids prod version)
'19 Nov 06 ac - fixed column break <font color='maroon">[arch]</font>, the color had a single quote casuing the closing </td> to fail.
'20 NOV 06 LJl revised the CHECKED section as the CHECK indication would be global from here.  Added below, just above the individual line
'27 nov 06 ac - added debug var
'27 fixed blue icon issue adding Q value to action url
'11 DEC 06 LJL added 9 as option for tested numbers for special circumstances - to be set manually by tech support into db table tx_rsys_candjob
'21 DEC 06 LJL adjusted the output of nationalities as it was incorrectly nestled within a 2000 ILO section.
'03 mar 07 ac increase db timeout
'11 mar 07 ac - only reverse applicant checkboxes
'21 mar 07 ac - added range check js fucntion removed related ASP code as it worked only for ranges from 1 to 9
'14 June 07 ac - chnaged "for (var i=0; i<elements.length; i++)" to	"for (i=0; i<elements.length; i++)" declared var i counter higer
'                in function. stopped object HTMLInputElement error
'29 JUL 07 LJL resolved js error in WTO IE
'24 AUG 07 LJL color of 1. Set option
'04 FEB 08 INT added drop down,checkboxe  and queries for getting type of candidate (Eraps)
'14 FEB 08 INT Icluded eraps-login-check.asp to check the login and access of eraps authorized user.
'16 FEB 08 INT applied session check for org(ILO-2000) for candidate types.
'03 mar 08 ac - still not saving; same issue as friday.
 '03 mar 08 ac - added the check against a post number as the candid_c is not unique to retrive cand. type shorlisted color
 '05 MAR 08 LJL added reduced rights for evals per ILO rights changes
 '06 mar 08 ac - add check for ILO and erpas for "updateonly" form or erro appears on other unshare orgs pages
 '06 MAR 08 LJL
  ' 15 mar 08 - ac - grab eRAPS VN flag here
 '17 MAR 08 INT access block for SUC
'29 APR 08 LJL changed UN to IntOrg
'12 jul 08 ac - don't need this check as it is done in eraps-login-check.asp
'30 SEP 08 LJL added REP for ILO criteria
 '30 SEP 08 LJL added new listings for ILO represented, under, etc
'23 APR 09 LJL changed out " and ' combination on [upd] screened applicants which was deleting them from the list per WHO and ILO
'21 MAY 09 LJL changed intra log insert to be year based table
'10 SEP 09 LJL added right for WHO to restrict online testing access
'20 OCT 09 LJL added WHO/UNAIDS/ICC check to distinguish GSM internals
'24 DEC 09 LJL slightly revised to move GSM int indicator from cand-list7.asp
'12 APR 10 DD Arranged sort by cand_lnam_t and then cand_fnam_t
'05 MAY 10 DD Arranged candidates belonging to same nationality by cand_lnam_t and then cand_fnam_t when sort by nationality
'26 MAy 10 DD Disallowed setting for Online Testing of EN Masse Update link if there is no data set for hours, days and  contact email of Online Testing of the VN
'08 JUN 10 LJL revised formatting for ILO RAPS columns
'21 JUN 10 LJL removed extraneous <tr and <td
'21 JUN 10 LJL added excel output
'21 JUN 10 LJL - modifed cllist to use the criteria that are set for internal, nats, etc for PDF and Excel outputs
'22 JUN 10 LJL added overallrank totater
'12 JUL 10 LJL revised sorting and added sorting to the Auto Eval system
'25 SEP 10 LJL added three new columns for ITU and others - pre screening items - ITU has prescreen, Edu, Lang levels, others can use as needed.
'	REMOVED GRID percentages, revised column sizings, revised some colors
'26 SEP 10 LJL not show prescreen items if not set for it
'28 SEPT 10 DD Applied DataValidation for EN mass update
'03 NOV 10 LJL added HTML output
'25 JAN 11 LJL added for ITU - language RU, AR, CN
'03 FEB 11 LJL removed <head> from around the script section
'09 FEB 11 LJL set PScr to PSel for ITU
'09 SEP 11 LJL added APlusREPctyvals to dims
'24 OCT 11 LJL changed GSM Int to just GSM for WHO/UNAIDS - was taking up another line if not.
'03 NOV 11 LJL removed template viewable, mainly for ILO, but probably should not be there for anyone other than rights
'06 DEC 11 LJL remove online testing column for IFRC, per VANELDEREN
'21 FEB 12 LJL not show Ex for ITU for internals/externals
'21 FEB 12 LJL change header to be for staff or not
'11 JUN 12 LJL made org name 6 chars instead of 3, as WIPO looked wrong in this list
'28 JUN 12 LJL revised ITU not consider list to be for cands with Contact HR ranks for overall, not per VN
'29 SEP 12 LJL added all nat quota indications for UPU 5500		
'14 FEB 13 LJL added sort on covering letter, requested by ILO
'20 MAR 13 LJL checked birthdate setting - some UNAIDS people are having very high ages.  Need to set birthdate in cand table from HQDATAHUB table better.
'01 NOV 14 LJL movetoposts section is candlist5.asp
'27 JUN 15 GG added candsCount to the rsys_matrix.asp form
'08 AUG 15 LJL REMOVED EXTRA NONSENSE - cand-list8, and then odd reference to query on that page
'08 AUG 15 LJL revise the entire select letter process - put into this page
'26 AUG 15 LJL add back in online testing for WTO
'04 SEP 15 LJL get if there are any screening questions for this list of applicants, used to be on cand8.asp include
'28 SEP 15 LJL reinstated cand-list8.asp to provide the basic VN info at the page header
'29 SEP 15 LJL added candrank_rem_m_" & session("template_org_code") & " AS thisorg_rankrem, candjob_rank_rem_m to getfields in cand-list7.asp
'05 FEB 16 GG added fixed Mixed Content Blocker problem 
'27 FEB 16 GG added 
'09 Mar 16 GG Fixed page split bug for matrix 
'23 MAR 16 LJL changed tested to capture all tested levels (in candlist7.asp)
'18 JUN 18 GG Fixed sort url 
'18 JUN 18 GG style selected page
'08 feb manoj uploaded latest changes of header
'28 APril 2018 manoj fix the space issue
'6june2018 manoj fix print privew issue
'07 OCT 24 LJL changed docsettings to admin_CVPDF_docsettings.asp for PDF revisions 
'12 OCT 24 LJL revised matrix section and worked on making it optimal
'14 OCT 24 LJL changed colors for background td for generation of PDF, Matrix, comms
' 03 NOV 24 LJL reiterated the stop of no selection letter being present




%>

 <link href="styles/bootstrap.min.css" rel="stylesheet" />
    <link href="styles/custom.min.css" rel="stylesheet" />


<!--#include file = "../includes/rsys_db_dim.asp"-->
<!--#include file = "../includes/rsys_db_select_dim.asp"-->
<!--#include file = "../includes/rsys_logs_dim.asp"-->
<!--#include file = "../includes/rsys_int_dim.asp"-->
<!--#include file = "../includes/rsys_int_select_dim.asp"-->
<!--#include file = "includes/include_check_login.asp"-->
<%
dim getstatsql, getstat, GetJobssql, GetJobs, GETJOBCTYsql, GETJOBCTY, GETJOBCTYNONsql, GETJOBCTYNON, GETJOBCTYAsql, GETJOBCTYA, GETJOBCTYBsql, GETJOBCTYB, GETJOBCTYB1sql, GETJOBCTYB1, GETJOBCTYB2sql, GETJOBCTYB2, GETJOBCTYCsql, GETJOBCTYC, ctyvals, NONREPctyvals, AREPctyvals, BREPctyvals, B1REPctyvals, B2REPctyvals, CREPctyvals, APlusREPctyvals

dim v_year, yearshort, urlsort, county, yo_tot, yo_count, japids
dim delsql, delit, footerama, cty, gcranksql, gcrank, defaultSort, sort, butty, namer, emailcheck, emailcheck2, japbirth, candcounter, gCandListsql, gCandList
dim Getstatusessql, Getstatuses, GetJQsql, GetJQ, getcorrid2sql, getcorrid2, getcorrid3sql, getcorrid3
dim gl_corres, NAMER2, linenum, yo, yo_total, jobsize, jobcolor, yo_result, GETJOBTESTEDsql, GETJOBTESTED, GETJOBSCREENEDsql, GETJOBSCREENED
dim GETJOBSHORTLISTsql, GETJOBSHORTLIST, GETJOBSELECTEDsql, GETJOBSELECTED
dim GETJOBRECOMMsql, GETJOBRECOMM, ByJ, getcorridsql, getcorrid, GETRANK1sql, GETRANK1, GETRANK2sql, GETRANK2
dim defIds1, defnames1, i1, i2, GETJAFsql, GETJAF, CHECKCANDY3sql, CHECKCANDY3, CHECKSHORT3sql, CHECKSHORT3, CHKtestsql, CHKtest
dim CHECKTEST1sql, CHECKTEST1, UPDCANDYsql, UPDCANDY, INSERTCORRsql, INSERTCORR
dim pv_emailerid, pv_testid, vacincl, CHECKTEST2sql, CHECKTEST2, CHECKINTERVIEW3sql, CHECKINTERVIEW3, CHECKSELECT3sql, CHECKSELECT3
dim GetCORRSsql, GetCORRS, CHECKTEST3sql, CHECKTEST3, CHECKCANDY1sql, CHECKCANDY1
dim ASParr, candyfind1, candyfind2, candylist1, candylist2, cand_rank1, cand_rank2
dim inscandjobsql, inscandjob, UPDcandjobsql, updcandjob, singleonly, GETYEARSsql, GETYEARS, short_year, GetVACSsql, GetVACS, yearlist
dim nocandcounter, CHECKVACsql, CHECKVAC, pv_corrid, status_job_id, CHECKCANDsql, CHECKCAND, clstat, systemmsg, GETCORRINFOsql, GETCORRINFO
dim GetVACsql, GetVAC, pv_CANDYid, mimeit, CHECKCANDY2sql, CHECKCANDY2, pv_UNCANDYid, pv_noCANDY
dim CHECKSHORT1sql, CHECKSHORT1, pv_unshortlistid, pv_shortlistid, pv_noshortlist, Logsssql, Logss, corr_id, pv_newq, pv_checkedit
dim pv_DEBUG, item,GetERapsInfoSql,GetERapsInfo,GetErapShtSql,GetErapSht,InstERapsInfoSql,InstERapsInfo,GetErapsCandInfoSql,GetErapsCandInfo,GetColorSql,GetColor
dim obj_int_select_CmdI,checkILOrolesql,checkILOrole,isHM,isRU,GetVacancysql,GetVacancy,currentStatus,cand_typ_val
dim lcand_type, ERapsVN,contact_prn_id,GETSUCsql,GETSUC,prn_id_suc,position_suc,width,checkErapsSql,checkEraps
dim pv_screener,  	pv_updonlyindic, pv_fullinternal
dim cranker, cranker1, cranker2, cranker3, cranktot, ranktot, ccog1, ccog2, ccog3, count, rank_official_i, ranktotal, job_count, qyear_option, qstart_year, qend_year
dim GETJAFINFO1sql, GETJAFINFO1, pv_overalltotaler, pv_ps1, pv_ps2, pv_ps3, dir_current, pv_showprescreen
dim candsCount
dim GETheadersql, GETheader, pv_totalapplicant, pv_totalinterview, pv_male, pv_female, pv_screen,pv_shortlisted
dim gSLsql, gSL, pv_selectletter

'27 FEB 16 GG added 
dim rowsCountConst, candsTo
rowsCountConst =1000

'18 JUN 18 GG style selected page
dim currentPage
currentPage = 0

candsTo = 0
If request.QueryString("candsto") <> "" then
	candsTo = int(request.QueryString("candsto"))
	currentPage = candsTo/rowsCountConst
End if

'09 Mar 16 GG Fixed page split bug for matrix 
Dim scriptUrl,candstoSubstr,strContent
Dim iStart
scriptUrl = Request.ServerVariables("QUERY_STRING")
iStart = InStr(1, scriptUrl, "candsto=", vbTextCompare)
if iStart > 0 Then
	candstoSubstr = Mid(scriptUrl, iStart, CInt(Len(scriptUrl) - iStart) + 1) 
End If
scriptUrl = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")

'Response.Write"<br>AAA " & candstoSubstr

'26 SEP 10 LJL not show prescreen items if not set for it
'07 APR 15 LJL removed extra columns for all orgs besides ITU, WMO, DEMO
dir_current = request.servervariables("PATH_TRANSLATED")
if session("template_org_code") <> 1000 AND session("template_org_code") <> 1500 AND session("template_org_code") <> 2000 AND session("template_org_code") <> 2800 AND session("template_org_code") <> 3000 AND session("template_org_code") <> 5500 then 
	'OR (UCASE(Instr(1, dir_current, "STAGE", 1)) > 0 OR UCASE(Instr(1, dir_current, "DEMO", 1)) > 0) then
	pv_showprescreen = 1
else
	pv_showprescreen = 0
end if



' SET THE TITLES for the Prescreening items - ITU wants Edu and Language Levels as them
if session("template_org_code") = 2400 then
	pv_ps1 = "PreSel"
	pv_ps2 = "Edu"
	pv_ps3 = "Lang"
else
	pv_ps1 = "PS1"
	pv_ps2 = "PS2"
	pv_ps3 = "PS3"
end if

ERapsVN = ""
japids = ""


' 27 nov 06 ac initialised,now need to call  - response.Flush
Response.Buffer=true
Server.ScriptTimeOut = 320

'27 nov 06 ac - debug
if instr(session("CLI_RSYS_ADMIN_USER"),"ILO-UNIT 1R") or instr(session("CLI_RSYS_ADMIN_USER"),"ILO-CURLEYA") or instr(session("CLI_RSYS_ADMIN_USER"),"ILO-MANAGER 1H") then
	pv_DEBUG = 0
else
	pv_DEBUG = 0
end if

If instr(Request.querystring("Q"),"Matrixonly")  then
    pv_page_title = "VACANCY APPLICANT MATRIX LISTING"
else
	pv_page_title = "APPLICANTS TO VACANCY"
end if

 	pv_updonlyindic = 0

 If instr(Request.querystring("Q"),"UpdOnly") then
 	pv_updonlyindic = 1
 else
 	pv_updonlyindic = 0
  end if

'27 Nov 06 ac -  - debug flag - could be moved to page header for all files?
if pv_DEBUG then
	response.write "<br>[hrd-cllist] -  pv_newq: " & pv_newq
	For Each Item In Request.Form
		response.write "<br>rf. " & Item &  " = " &  Request.Form(Item)
	Next
	Response.Write"<br>u. " & Request.QueryString
	response.write "<br>stop 1"
	'response.End()
end if

''<!-------------------- for the nationals or eligible types --------------->

if not cty then
	cty = "0"
end if

If request.form("jobinfo_uid_c") <> "" then
	newjobid = request.form("jobinfo_uid_c")
elseIf request.querystring("jobinfo_uid_c") <>"" then
	newjobid = request.querystring("jobinfo_uid_c")
else%>
	<br><br>
  	<font >You have reached this page in error.  Please <a href="index.asp" target="_parent">start here</a>.</font>
	<br><br>
  <%'<!-- Unsuported Tag
end if

' manoj dim GETheadersql, GETheader, pv_totalapplicant, pv_totalinterview, pv_male, pv_female, pv_screen,pv_shortlisted

	GETheadersql = "select count(candjob_id_c) as totalapplicant,count(shortlist) as shortlist,count(interv) as interv,count(screened) as screened,count(selected) as selected,COUNT(CASE WHEN sexc = 'M' then 1 ELSE NULL END) as Male,COUNT(CASE WHEN sexc = 'F' then 1 ELSE NULL END) as female,COUNT(CASE WHEN sexc = 'O' then 1 ELSE NULL END) as other" 
    GETheadersql= GETheadersql &" from(SELECT DISTINCT candjob_id_c, case candjob_shortlist_i when 1 then count(candjob_shortlist_i) end as shortlist, case candjob_interv_i when 1 then count(candjob_interv_i)   end as interv, case candjob_screened_i when 1 then count(candjob_screened_i) end as screened , case candjob_selected_i when 1 then count(candjob_selected_i) end as selected ,sexc   FROM v_rsys_jobs_sub_view_" & session("template_org_code") & "  WHERE jobinfo_uid_c = " & newjobid & " group by candjob_shortlist_i,candjob_id_c,candjob_interv_i,candjob_screened_i,candjob_selected_i,sexc   )     a"
	set GETheader = rsys_db_select.execute(GETheadersql)
  
  '  Response.Write "<br>T MAIN QUERY : " & GETheadersql  
if newjobid <> "" then
	GETJAFINFO1sql = "SELECT gl.gradetype_typ_t, j.jobinfo_job_en_t, j.jobinfo_uid_c, j.jobinfo_vac2_c, j.jobinfo_test_start_d, j.jobinfo_test_end_d, j.jobinfo_test_hour_c, j.jobinfo_test_contact_c, j.jobinfo_test_fax_c, j.jobinfo_test_score_c, j.jobinfo_test_lastday_c, j.jobinfo_test_minute_c, j.jobinfo_test_delaytime_c FROM td_rsys_jobinfo j LEFT OUTER JOIN tr_rsys_gradelevel gl ON j.gradelevel_id_c = gl.gradelevel_id_c WHERE j.jobinfo_uid_c = " & newjobid & ""
	set GETJAFINFO1 = rsys_db_select.execute(GETJAFINFO1sql)
end if



'  <!-------------QUOTA LIST
'  - this list is pulled from the database and then the c rank is found in the number list and then that index placement in that index is pulled from teh Alpha ranking list
'  ----------->

  gcranksql = "	SELECT Rtrim(countryquota_sht_t) AS crankeralpha, countryquota_rank_i AS cranker 	FROM tr_rsys_countryquota 	ORDER BY 1"

	rsys_db_select.CommandTimeout = 320
  set gcrank =rsys_db_select.execute(gcranksql)

  ' NEEDNEEDNEED cranklist_alpha = valuelist(gcrank.crankeralpha)
  ' NEEDNEEDNEED cranklist = valuelist(gcrank.cranker)

'  <!--- Added by LJL on July 2001 --->

 '12 APR 10 DD Arranged sort by cand_lnam_t and then cand_fnam_t
 'defaultSort = "cand_lnam_t,isNull(cand_gnd_i,sex_code),candjob_id_c,cand_fnam_t"
  defaultSort = "cand_lnam_t,cand_fnam_t,isNull(cand_gnd_i,sex_code),candjob_id_c"
  if not yo_tot then
  yo_tot = "0"
  end if
  if not yo_count then
  yo_count = "0"
  end if
'  if not response.write then
'  response.write  = "0"
'  end if

'21 JUN 10 LJL added excel
if request.querystring("goExcel") = 99 then

else%>
<!--#include file = "includes/include_admin_frame_top.asp"-->
<%end if



if session("template_org_code")= 2000 then %>
<!--#include file = "eraps-login-check.asp"-->
<% end if %>

<%
'21 JUN 10 LJL added excel
'03 FEB 11 LJL revised excel name


if request.querystring("goExcel") = 99 then
	dim pv_filenamer, pv_dayer
	if day(now()) < 10 then
		pv_dayer = 0 & day(now())
	else
		pv_dayer = day(now())
	end if
	pv_filenamer = GETJAFINFO1("jobinfo_vac2_c") & "_" & pv_dayer & monthname(month(now()),1) & right(year(now()),2)
	
	'Response.ContentType = "text/csv"
	'Response.AddHeader "Content-Disposition", "attachment; filename=" & pv_filenamer & ".csv"

	Response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "content-disposition", "inline; filename=" & pv_filenamer & ".xls"
end if%>



<!--manoj added div for number header -->
<div class="row tile_count">
            <div class="col-md-2 col-sm-1 col-xs-6 tile_stats_count">
              <span class="count_top"><i class="fa fa-user"></i> Applicants</span>
              <div class="count"><%= GETheader("totalapplicant") %></div>
              <span class="count_bottom"><i class="green"> </i> </span>
            </div>
         
            <div class="col-md-2 col-sm-1 col-xs-6 tile_stats_count">
              <span class="count_top"><i class="fa fa-user"></i> Females</span>
              <div class="count red"><%= GETheader("female") %></div>
              <span class="count_bottom"><i class="red"><i class="fa fa-sort-desc"></i> </i> </span>
            </div>

            <div class="col-md-2 col-sm-1 col-xs-6 tile_stats_count">
                <span class="count_top"><i class="fa fa-user"></i> Males</span>
                <div class="count blue"><%= GETheader("Male") %></div>
                <span class="count_bottom"><i class="green"><i class="fa fa-sort-asc"></i></i> </span>
            </div>

			<div class="col-md-2 col-sm-1 col-xs-6 tile_stats_count">
                <span class="count_top"><i class="fa fa-user"></i> Other</span>
                <div class="count"><%= GETheader("other") %></div>
                <span class="count_bottom"><i class="green"><i class="fa fa-sort-asc"></i></i> </span>
            </div>

            <div class="col-md-2 col-sm-1 col-xs-6 tile_stats_count">
              <span class="count_top"><i class="fa fa-user"></i> Screened</span>
              <div class="count orange"><%= GETheader("screened") %></div>
              <span class="count_bottom"><i class="green"><i class="fa fa-sort-asc"></i></i></span>
            </div>

            <div class="col-md-2 col-sm-1 col-xs-6 tile_stats_count">
              <span class="count_top"><i class="fa fa-clock-o"></i> Interviewed </span>
              <div class="count brown"><%= GETheader("interv") %></div>
              <span class="count_bottom"><i class="green"><i class="fa fa-sort-asc"></i> </i> </span>
            </div>
           
            <div class="col-md-2 col-sm-1 col-xs-6 tile_stats_count">
              <span class="count_top"><i class="fa fa-user"></i> Shortlisted</span>
              <div class="count purple"><%= GETheader("shortlist") %></div>
              <span class="count_bottom"><i class="green"><i class="fa fa-sort-asc"></i> </i> </span>
            </div>

            <div class="col-md-2 col-sm-1 col-xs-6 tile_stats_count">
              <span class="count_top"><i class="fa fa-clock-o"></i> Selected </span>
              <div class="count green"><%= GETheader("Selected") %></div>
              <span class="count_bottom"><i class="green"><i class="fa fa-sort-asc"></i> </i> </span>
            </div>

          </div>    

<!--#include file = "rsys_vac_admin_menu.asp"-->

<%
' added by INT on 03/17/08 for SUC candidate access block to this page
if session("template_org_code") = 2000 then
	'12 jul 08 ac - don't need this check as it is done in eraps-login-check.asp
	'checkErapsSql = "select jobinfo_eraps from td_rsys_jobinfo where jobinfo_uid_c = " & newjobid
	'set	checkEraps =rsys_db_select.execute(checkErapsSql)

if len(Request.QueryString ("jobinfo_uid_c")) then

 if ERapsVN = "2"  then

	getVACsql = "SELECT  jobcreate_auth1_c, jobcreate_auth2_c,jobcreate_auth3_c,jobcreate_auth4_c,jobcreate_auth5_c,jobcreate_hr_auth1_i , jobcreate_hr_auth2_i , s1.prn_lnam_t + ', ' + s1.prn_fnam_t AS s1name, s2.prn_lnam_t + ', ' + s2.prn_fnam_t AS s2name, s3.prn_lnam_t + ', ' + s3.prn_fnam_t AS s3name, s4.prn_lnam_t + ', ' + s4.prn_fnam_t AS s4name, s5.prn_lnam_t + ', ' + s5.prn_fnam_t AS s5name  FROM erec_test.dbo.td_rsys_jobcreate jc LEFT OUTER JOIN rsys_int.dbo.staff_" & session("template_org_code") & " s1 ON jc.jobcreate_auth1_c = s1.prn_id LEFT OUTER JOIN rsys_int.dbo.staff_" & session("template_org_code") & " s2 ON jc.jobcreate_auth2_c = s2.prn_id LEFT OUTER JOIN rsys_int.dbo.staff_" & session("template_org_code") & " s3 ON jc.jobcreate_auth3_c = s3.prn_id LEFT OUTER JOIN rsys_int.dbo.staff_" & session("template_org_code") & " s4 ON jc.jobcreate_auth4_c = s4.prn_id LEFT OUTER JOIN rsys_int.dbo.staff_" & session("template_org_code") & " s5 ON jc.jobcreate_auth5_c = s5.prn_id  WHERE jobinfo_uid_c = " & newjobid
	set getVAC = rsys_db_select.execute(getVACsql)
	contact_prn_id = trim(session("contactpersid"))

	if (getVAC.eof = false)  then
	  if (getVAC("jobcreate_auth4_c") <> " ") then

		GETSUCsql = "SELECT DISTINCT s.prn_fnam_t, s.prn_lnam_t, s.eml_add_t,rtrim(s.prn_id) as prn_id, s.prn_uid_c, s.prn_id_c,ISNULL(p.position_id_i, 0 ) as position_id_i FROM rsys_int.dbo.staff_2000  s INNER JOIN   rsys_int.dbo.position_2000 p ON s.position_id_i = p.position_id_i WHERE (p.position_id_i = 5) and prn_id =" & trim(getVAC("jobcreate_auth4_c"))
		set GETSUC = rsys_int.execute(GETSUCsql)

'	  end if
'	end if

	if (GETSUC.eof = false ) then
		prn_id_suc = trim(GETSUC("prn_id_c"))
		position_suc = trim(GETSUC("position_id_i"))

		if  position_suc <> 0 then
			if contact_prn_id = prn_id_suc  then
				Response.Write ("<div align='center'><Br><br><font class='alertbold'>You are not authorized to edit this vacancy (msg: cllist002)</font><br></div>")
				Response.End
			end if
		end if

	end if

	  end if
	end if


  end if 'checkEraps
end if
end if
%>



<!--#include file = "cand-list1.asp"-->
<!--#include file = "cand-list2.asp"-->

<%If Request.querystring("MailForm") <> "" then
		GetclInfosql = "SELECT td_rsys_cand.cand_email_t, td_rsys_cand.cand_lnam_t, isNull(cand_gnd_i,sex_code), candjob_id_c, td_rsys_cand.cand_fnam_t, firstname, lastname, td_rsys_cand.staff_nbr_" & session("template_org_code") & " AS thisorg_staffno, s.Email FROM td_rsys_cand LEFT OUTER JOIN v_staff_list_" & session("template_org_code") & " s ON td_rsys_cand.staff_nbr_" & session("template_org_code") & " = s.SID "
		IF session("temploate_org_code") = 1000 OR session("template_org_code") = 1500 then
			'<!---------------------------- ADD ORG INFO 1000 only --------------------------->
			GetclInfosql = GetclInfosql & "COLLATE SQL_Latin1_General_CP850_CI_AI "
		end if
		GetclInfosql = GetclInfosql & " WHERE td_rsys_cand.cand_id_c = cand_id_c "
		rsys_db_select.CommandTimeout = 320
		set GetclInfo =rsys_db_select.execute(GetclInfosql)
 		'          Logsssql = "		INSERT INTO tl_intra_log_" & session("template_org_code") & "_" & year(now()) & " 			(user_id_t, log_upd_d, log_beg_val, log_upd_val) 			VALUES 			('" &_
		'		  left(session("template_org_name") & "-SELECT+P11-" & request.querystring("cand_id_c") & ") '" & Now() & "', '" & session("template_org_name") & "-SELECT+P11-" & GetclInfo("cand_fnam_t") &   GetclInfo("cand_lnam_t") & "-" & newjobid & ", 'SENT P11' ) "
		'          set Logss = rsys_logs.execute(Logsssql)
end if





'03 FEB 11 LJL revised excel name
if request.querystring("goExcel") = 99 then

else %>
<script language="JavaScript">
  	<!--
function isInteger (s)
   {
      var i;

      if (isEmpty(s))
      if (isInteger.arguments.length == 1) return 0;
      else return (isInteger.arguments[1] == true);

      for (i = 0; i < s.length; i++)
      {
         var c = s.charAt(i);

         if (!isDigit(c)) return false;
      }

      return true;
   }

   function isEmpty(s)
   {
      return ((s == null) || (s.length == 0))
   }

   function isDigit (c)
   {
      return ((c >= "0") && (c <= "9"))
   }

function rangeCheck()
{
	var counter = 0
	var pagecount
	var start
	var st, ed, fail, i
	fail = 1
	st  = document.sortForm4c.candlist1.value
	endCounter  = document.sortForm4c.candlist2.value
	
	pagecount = parseInt(document.getElementsByName('candstoHiddenvalue')[0].value) / 1000
	if(pagecount>1)
	{
		counter = parseInt(document.getElementsByName('candstoHiddenvalue')[0].value) - 1000	
	}	
	
	if (!isInteger(st))
	{
		alert(st + " is not an integer")
		fail = 0
		return false
	}

	if (!isInteger(endCounter))
	{
		alert(endCounter + " is not an integer")
		fail = 0
		return false
	}


	with (document.pdfForm)
	{
		for (i=0; i<elements.length; i++)
		{
			if (elements[i].type=='checkbox')
			{
				/*only reverse applicants checkboxes*/
				elm = elements[i];
				if (elm.name ==  'japids')
				{
					counter = counter +1
					if (counter >= st && counter <= endCounter)
						elements[i].checked = true
					else
						elements[i].checked = false
				}
			}
		}
	}
}
function unCheck()
{
	with (document.pdfForm)
	{
		for (var i=0; i<elements.length; i++)
		{
			if (elements[i].type=='checkbox')
			{
				/*only reverse applicants checkboxes*/
				elm = elements[i];
				if (elm.name ==  'japids')
					elements[i].checked = !elements[i].checked;
			}
		}
	}
}
  	//-->

  	function DataValidation()
  	{

  		var found_checked_cand;
  		found_checked_cand = 0
  		//alert("testing");

  		for(var i=0; i < document.getElementsByName('japids').length; i++)
  		{
  			if (document.getElementsByName('japids')[i].checked == true)
  			{
  				found_checked_cand = 1
  				break;
  			}
  		}

  		if (found_checked_cand == 0)
  		{
  			alert("Please select at least one applicant.");
  			return false;
  		}

  		return true;
  	}

	</script>

<%end if
'<!--- isDefiend("MailForm") --->%>

<!--#include file = "cand-list3.asp"-->
<!--#include file = "cand-list4.asp"-->
<!--#include file = "cand-list5.asp"-->
<!--#include file = "cand-list6.asp"-->
<!--#include file = "cand-list7.asp"-->
<!--#include file = "cand-list8.asp"-->


<style type="text/css">


a { 
	text-decoration:none; 
	color:#00c6ff;
}

h1 {
	font: 4em normal Arial, Helvetica, sans-serif;
	padding: 20px;	margin: 0;
	text-align:center;
}

h1 small{
	font: 0.2em normal  Arial, Helvetica, sans-serif;
	text-transform:uppercase; letter-spacing: 0.2em; line-height: 5em;
	display: block;
}

h2 {
    font-weight:700;
    color:#bbb;
    font-size:20px;
}

h2, p {
	margin-bottom:10px;
}

.container {width: 960px; margin: 0 auto; overflow: hidden; height:910px;}

.tooltip {
	display:none;
	position:absolute;
	border:1px solid #333;
	background-color:#161616;
	border-radius:5px;
	padding:10px;
	color:#fff;
	font-size:12px Arial;
}



body {
  margin: 0 auto;
  padding: 0 20px;
  font-family: Arial, Helvetica, sans-serif;
  font-size: 11px;
  color: #555;
}

.table-wrapper table {
  border: 0;
  padding: 0;
  margin: 0 0 20px 0;
  border-collapse: collapse;
}

.table-wrapper th {
  padding: 5px;
  /* NOTE: th padding must be set explicitly in order to support IE */
  text-align: right;
  font-weight: bold;
  line-height: -1em;
  color: #FFF;
  background-color: #787878;
}

.table-wrapper tbody td {
  padding: 10px;
  line-height: 18px;
  border-top: 1px solid #E0E0E0;
}

.table-wrapper tbody tr:nth-child(2n) {
  background-color: #F7F7F7;
}

.table-wrapper tbody tr:hover {
  background-color: #EEEEEE;
}

.table-wrapper td {
  text-align: left;
}

.table-wrapper td:first-child,
th:first-child {
  text-align: left;
}


</style>

<style type="text/css" media="print">
    .NonPrintable
    {
      display: none;
    }
    
   .newanchor[href]:after {
    content: none;
   }
    
  
  
    
  </style>

<script type="text/javascript" src="https://code.jquery.com/jquery-2.2.0.min.js"></script>

<script type="text/javascript">

    $(document).ready(function () {


        $(".table-wrapper").stickyTableHeaders();


    });
</script>

 <script type="text/javascript">

     $(document).ready(function () {
         // Tooltip only Text
         $('.masterTooltip').hover(function () {


             // Hover over code
             var title = $(this).attr('title');
           
             $(this).data('tipText', title).removeAttr('title');
             $('<p class="tooltip"></p>')
                .text(title)
                .appendTo('body')
                .fadeIn('slow');
         }, function () {
             // Hover out code
             $(this).attr('title', $(this).data('tipText'));
             $('.tooltip').remove();
         }).mousemove(function (e) {
             var mousex = e.pageX + 20; //Get X coordinates
             var mousey = e.pageY + 10; //Get Y coordinates
             $('.tooltip')
                .css({ top: mousey, left: mousex })
         });
     });

</script>

<%
if pv_DEBUG then
	response.write "<br>C 7 - done"
end if

'21 JUN 10 LJL added excel
if request.querystring("goExcel") = 99 then%>

<Table>
<tr>
    <td align='left' width="50%" height="26" bgcolor="#ECF4FF" align='left' colspan=5><b><% response.write VacInfo("jobinfo_job_en_t")
    if   len(VacInfo("jobinfo_vac2_c")) then
    	response.write " " & VacInfo("jobinfo_vac2_c")
    end if%></b>
    </td>
	<td align='left' width="50%" height="26" bgcolor="<% response.write VacInfo("status_bg_color_c")%>" align='left' colspan=5><font color="<% response.write VacInfo("status_font_color_c")%>"><% response.write VacInfo("status_dsc_t")%></font></td>
</tr>
</TABLE>

<%end if
	
'08 AUG 15 LJL REMOVED EXTRA NONSENSE - cand-list8, and then odd reference to query on that page
' added missing end if from cand8
	end if
%>

<%'08 AUG 15 LJL REMOVED EXTRA NONSENSE - cand-list8, and then odd reference to query on that page
'<!--#include file = "cand-list8.asp"-->
%>


<%response.flush
if GCandList.eof = true and  session("template_org_code")<> 2000  then
candsCount = 0
%>
<br>
<hr size=1>
<br>
  <center><i><span class="alert" >No applicants meet the criteria you specified for this post or none have applied as of yet</span></i></center>
<Br><br>
<%
else
candsCount = GCandList.recordcount
%>
<table width="100%" border="0" bordercolor="lime" align="center">
<TR>
	<td valign="top" class="littletitle" align="center">
    <FONT color="navy">
    <%if session("template_org_code")<> 2000 then%>
	<%= candsCount %> total

	<%end if
	 if session("template_org_code") = 2000 and GCandList.eof = true then %>

	  <%=0%>  total
	<%elseif session("template_org_code") = 2000 and GCandList.eof = false then%>
		<%=candsCount%> total
	<%end if
	If instr(Request.querystring("Q"),"nats")  then%>
	<font color="navy">(qualified external nationals) </font>
	<%end if
	' HAD TO REVISE THE TERM ELIGIBLE BECAUSE IT IS CONTAINED IN NONELIGIBLE
      If instr(Request.querystring("Q"),"eligible")  then%>
	<font color="navy">(Qualified external nationals and Internals - All Eligible) </font>
	<%end if

     If  instr(Request.querystring("Q"),"REPNON") then%>
	<font color="navy">(Represented - Non) </font>
    <%end if
     If  instr(Request.querystring("Q"),"REPAplus") then%>
	<font color="navy">(Represented - A*) </font>
     <%elseIf  instr(Request.querystring("Q"),"REPA") then%>
	<font color="navy">(Represented - A) </font>
    <%end if
     If  instr(Request.querystring("Q"),"REPBB") then%>
	<font color="navy">(Represented - B-All) </font>
    <%end if
     If  instr(Request.querystring("Q"),"REPB1") then%>
	<font color="navy">(Represented - B1) </font>
    <%end if
     If  instr(Request.querystring("Q"),"REPB2") then%>
	<font color="navy">(Represented - B2) </font>
    <%end if
     If  instr(Request.querystring("Q"),"REPC") then%>
	<font color="navy">(Represented - C) </font>
    <%end if

     If  instr(Request.querystring("Q"),"nonELG") then%>
	<font color="navy">(NON-Eligible) </font>
    <%end if
	If  instr(Request.querystring("Q"),"ints") then%>
	<font color="navy">(Internal applicants)</font>
	<%end if
	If  instr(Request.querystring("Q"),"uns") then%>
	<font color="navy">(UN/Int'l Org applicants)</font>
	<%end if
          	If  instr(Request.querystring("Q"),"short") then%>
          		<i><font color="navy">(Short-term)</font>
			<%
          	end if
          	If  instr(Request.querystring("Q"),"others") then%>
          		<font color="navy">
			<%
          	end if
          	If  instr(Request.querystring("Q"),"shlisted") then%>
          		<font color="Blue">(Short-listed)</font>
			<%
          	end if
          	If  instr(Request.querystring("Q"),"selected") then%>
          		<font color="Blue">(Selected)</font>
			<%
          	end if
          	If  instr(Request.querystring("Q"),"screened") then%>
          		<font color="Blue">(Screened)</font>
			<%
          	end if
          	If  instr(Request.querystring("Q"),"UpdOnly") then%>
          		<font color="Blue">(Updated)</font>
			<%
          	end if
          	If  instr(Request.querystring("Q"),"app_test") then%>
          		<font color="Blue">(Tested)</font>
			<%
          	end if
          	If  instr(Request.querystring("Q"),"app_interv") then%>
          		<font color="Blue">(Interviewed)</font>
			<%
          	end if
			%>
          		listed
          	</TD>
      </tr>
<%
	
'04 SEP 15 LJL get if there are any screening questions for this list of applicants, used to be on cand8.asp include
GetJQsql = "SELECT qj_id_c FROM tx_rsys_qj WHERE jobinfo_uid_c = " & newjobid & " 	"
set GetJQ =rsys_db_select.execute(GetJQsql)
	
	
	'21 JUN 10 LJL added excel
if request.querystring("goExcel") = 99 then

else%>
<TR>
	<td class="textsmall" valign="top"  align="center">Submitted dates with Upd! indicate applicant has updated their profile since they applied to this post.
	<br>Screened and Shortlisted columns with Upd! indicate applicant has updated their profile since being screened or Shortlisted.</td>
</tr>
<%end if%>

</TABLE>

<%
'27 nov 06 ac - output to browser
response.flush
dim qstringer
dim qlist
qlist = ""
' GET Q values to pass to next page
For Each qstringer In Request.QueryString("Q")
   qlist = qlist & "&Q=" & qstringer
Next
' response.write "T QLISTER " & qlist

'07 APR 15 LJL set widths for columns in main table different for orgs with many columns vs orgs with fewer, ig WTO
dim pv_widthSEL, pv_width1, pv_width2, pv_width3, pv_width4, pv_width5, pv_width6, pv_width7, pv_width8, pv_width9, pv_width10, pv_width11, pv_width12, pv_width13, pv_width14, pv_width15, pv_width16, pv_width17, pv_width18, pv_width_table_cols, pv_width_table

if session("template_org_code") = 3000 then
	pv_widthSEL = "100%"
	pv_width1 = "5%"
	pv_width2 = "15%"
	pv_width3 = "5%"
	pv_width4 = "5%"
	pv_width5 = "5%"
	pv_width6 = "5%"
	pv_width7 = "5%"
	pv_width8 = "5%"
	pv_width9 = "5%"
	
	pv_width_table = "35%"
	pv_width_table_cols = "5"
	
	pv_width10 = "20%"
' PRE SCREEN COLUMNS
'	pv_width11 = 40
'	pv_width12 = 40
'	pv_width13 = 40
	pv_width14 = "20%"
' ONLINE TESTING
'	pv_width15 = "8%"
	pv_width16 = "20%"
	pv_width17 = "20%"
	pv_width18 = "20%"

elseif session("template_org_code") = 2400 OR session("template_org_code") = 2900 then
	pv_widthSEL = "100%"
	pv_width1 = "5%"
	pv_width2 = "15%"
	pv_width3 = "5%"
	pv_width4 = "5%"
	pv_width5 = "5%"
	pv_width6 = "5%"
	pv_width7 = "5%"
	pv_width8 = "5%"
	pv_width9 = "5%"
	
	pv_width_table = "40%"
	pv_width_table_cols = "9"
	
	pv_width10 = "11%"
' PRE SCREEN COLUMNS
	pv_width11 = "11%"
	pv_width12 = "11%"
	pv_width13 = "11%"
	pv_width14 = "11%"
' ONLINE TESTING
	pv_width15 = "11%"
	pv_width16 = "11%"
	pv_width17 = "11%"
	pv_width18 = "11%"
	
else
	pv_widthSEL = "100%"
	pv_width1 = "5%"
	pv_width2 = "15%"
	pv_width3 = "5%"
	pv_width4 = "6%"
	pv_width5 = "7%"
	pv_width6 = "7%"
	pv_width7 = "5%"
	pv_width8 = "5%"
	pv_width9 = "5%"

	pv_width_table = "40%"
	pv_width_table_cols = "6"
	
	pv_width10 = "16%"
' PRE SCREEN COLUMNS
'	pv_width11 = 40
'	pv_width12 = 40
'	pv_width13 = 40
	pv_width14 = "16%"
' ONLINE TESTING
	pv_width15 = "16%"
	pv_width16 = "16%"
	pv_width17 = "16%"
	pv_width18 = "16%"
	
end if%>


<% followurl = Request.ServerVariables("QUERY_STRING")
'18 JUN 18 GG Fixed sort url 
 %>

<table width="100%" border="0" align="center" bordercolor="red">

<%'21 JUN 10 LJL added excel
if request.querystring("goExcel") = 99 then%>

<tr>
	<td></td>
	<td>Name</td>
	<td>Gender/Int</td>
	<td>DOB</td>
	<td>Nat</td>
	<td>Applied</td>
	<td>CL/upd</td>
	<td>Eval</td>
	<td>Screen</td>
	<td>Shortlist</td>
	<td>Test</td>
	<td>Interview</td>
	<td>Rec'm</td>
	<td>Selected</td>
	<td></td>
</tr>


<%else%>

<TR>
	<form name="sortForm4c" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<td class="textsmall" valign="top"  align=center nowrap width="<%=pv_widthSEL%>">
	<hr>
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<INPUT TYPE="hidden" NAME="candstoHiddenvalue" VALUE="<%= candsto%>">
	To select large numbers of candidates, enter the line numbers from
	<input class="textsmall" type="text" name="candlist1" value="0" size="3">
	to <input class="textsmall" type="text" name="candlist2" value="0" size="3">
	<input class="textsmall" type="button" value=" Select these " onClick="rangeCheck();">
<%'20 NOV 06 LJl revised this section as the CHECK indication would be global from here.  Added below, just above the individual line%>
	<br>&nbsp;
	</td>
	</form>
</TR>
</table>



<div>Pages: &nbsp;&nbsp;
<% '27 Feb 16 GG added Pagination
If candsto>rowsCountConst then 
	If Len(candstoSubstr) Then
		strContent = Replace(scriptUrl,candstoSubstr, "candsto=" & (candsto-rowsCountConst)) 
	Else
		strContent = scriptUrl & "&candsto=" & (candsto-rowsCountConst)
	End If
	If request.form("sort") <> "" then
		strContent = strContent & "&sort=" & Replace(request.form("sort")," ", "+")
	elseif request.querystring("sort") <> "" then
		strContent = strContent & "&sort=" & Replace(request.querystring("sort")," ", "+")
	end if
%>
		<a href= '<% response.write strContent%>' ><b>Previous </b></a>
<% 
End IF 

if candsCount > rowsCountConst Then
	Dim pagesCount,pageIndex
	pagesCount = candsCount/rowsCountConst
	pageIndex = 0	
	
	Do while pagesCount > pageIndex
		pageIndex = pageIndex + 1
		
		If Len(candstoSubstr) Then
			strContent = Replace(scriptUrl,candstoSubstr, "candsto=" & (rowsCountConst*pageIndex)) 
		Else
			strContent = scriptUrl & "&candsto=" & (rowsCountConst*pageIndex)
		End If
		If request.form("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.form("sort")," ", "+")
		elseif request.querystring("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.querystring("sort")," ", "+")
		end if
%>
		<a <% 	If currentPage = pageIndex Then
					response.write "style='color:#990066;font-weight: bolder;'"
				End IF
		%> href='<% response.write strContent%>' ><b><% response.write pageIndex %></b></a>
<% 
	Loop
End If 

		If candsCount/rowsCountConst > 1 and candsCount > candsto then 
		If Len(candstoSubstr) Then
			strContent = Replace(scriptUrl,candstoSubstr, "candsto=" & (candsto+rowsCountConst)) 
		Else
			strContent = scriptUrl & "&candsto=" & (candsto+rowsCountConst+rowsCountConst) 
		End If
		If request.form("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.form("sort")," ", "+")
		elseif request.querystring("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.querystring("sort")," ", "+")
		end if
%>
	<a href='<%  response.write strContent %>' ><b>Next</b></a>
<% End If %>
</div>


<%' TABLE FOR ALL THE COLUMNS in the main page%>


<TABLE class="table-wrapper" border="0" bordercolor="orange" cellpadding="0" cellspacing="2" width="100%" align="center" >
    <thead>
<TR class=trheader9>
	<th class="textsmall" valign="top"  align=center nowrap width="<%=pv_width1%>">&nbsp;</th>

    <th class="textsmall" valign="top"  align=left nowrap width="<%=pv_width2%>">Name<Br><br>
    <form name="sortForm21" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<!--<input type="hidden" name="sort" value="cand_lnam_t,isNull(cand_gnd_i,sex_code),candjob_id_c,cand_fnam_t"> -->
	<input type="hidden" name="sort" value="cand_lnam_t ASC,cand_fnam_t ASC,isNull(cand_gnd_i,sex_code),candjob_id_c" ID="Hidden4">
	<%'<input src="../images/<%=session("template_org_code")>_admin/name.gif" type="image" onClick="document.sortForm2.submit();" value="Name">>
'	<input type="button" onClick="document.sortForm21.submit();" value=" + ">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm2.submit();">
	</form>

	<br>

    <form name="sortForm22" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<!--<input type="hidden" name="sort" value="cand_lnam_t,isNull(cand_gnd_i,sex_code),candjob_id_c,cand_fnam_t"> -->
	<input type="hidden" name="sort" value="cand_lnam_t DESC,cand_fnam_t DESC,isNull(cand_gnd_i,sex_code),candjob_id_c" ID="Hidden4">
	<%'<input src="../images/<%=session("template_org_code")>_admin/name.gif" type="image" onClick="document.sortForm2.submit();" value="Name">>
'	<input type="button" onClick="document.sortForm22.submit();" value=" - ">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm2.submit();">
	</form>
	</th>

	<th class="textsmall" valign="top" align=left nowrap width="<%=pv_width3%>">M/F<form name="sortForm4a1" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="isNull(cand_gnd_i,sex_code) ASC">
<%'	<input src="../images/<%=session("template_org_code")>_admin/gender.gif" type="image" onClick="document.sortForm4a.submit();" value="M/F">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm4a1.submit();">
<%'	<input type="button" onClick="document.sortForm4a1.submit();" value="-">%>
	</form>
	<br>
	<form name="sortForm4a2" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="isNull(cand_gnd_i,sex_code) DESC">
<%'	<input src="../images/<%=session("template_org_code")>_admin/gender.gif" type="image" onClick="document.sortForm4a.submit();" value="M/F">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm4a2.submit();">

<%'	<input type="button" onClick="document.sortForm4a2.submit();" value="-">%>
	</form>
<%'21 FEB 12 LJL not show Ex for ITU for internals/externals
'21 FEB 12 LJL change header to be for staff or not
if session("template_org_code") <> 2400 then%>
	Int/Ext
<%else%>
	Staff/Non
<%end if%>
	<form name="sortForm311" action="hrd-cllist.asp?<%=followurl%>" method="post"><INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")%>">
	<input type="hidden" name="sort" value=" thisorg_short DESC, cand_io_i DESC, cand_lnam_t, cand_fnam_t">
	<%'<input src="../images/<%=session("template_org_code")>_admin/int_ext.gif" type="image" onClick="document.sortForm3.submit();" value="Int/Ext"></th>>%>
	<%'<input type="button" onClick="document.sortForm31.submit();" value="+">%>
	<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm311.submit();">
	</form>

	<br>

	<form name="sortForm312" action="hrd-cllist.asp?<%=followurl%>" method="post"><INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")%>">
	<input type="hidden" name="sort" value=" thisorg_short ASC, cand_io_i ASC, cand_lnam_t, cand_fnam_t">
	<%'<input src="../images/<%=session("template_org_code")>_admin/int_ext.gif" type="image" onClick="document.sortForm3.submit();" value="Int/Ext"></th>>%>
	<%'<input type="button" onClick="document.sortForm31.submit();" value="+">%>
	<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm312.submit();">
	</form>

	</th>

	<th class="textsmall" valign="top"  align=left nowrap width="<%=pv_width4%>">DOB<br><Br>
	<form name="sortForm41" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="cand_bth_d ASC">
<%'	<input src="../images/<%=session("template_org_code")>_admin/age.gif" type="image" onClick="document.sortForm4.submit();" value="DOB">>
'	<input type="button" onClick="document.sortForm41.submit();" value=" + ">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm41.submit();">
	</form>
	<br>
	<form name="sortForm42" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="cand_bth_d DESC">
<%'	<input src="../images/<%=session("template_org_code")>_admin/age.gif" type="image" onClick="document.sortForm4.submit();" value="DOB">>
'	<input type="button" onClick="document.sortForm42.submit();" value=" - ">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm42.submit();">
	</form>

	</th>
	<th class="textsmall" valign="top"  align=left nowrap width="<%=pv_width5%>">Nats<br><Br>

	<form name="sortForm51" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="isNull(scty1,pcty1) DESC,cand_lnam_t,cand_fnam_t">
	<%'<input src="../images/<%=session("template_org_code")>_admin/nationality.gif" type="image" onClick="document.sortForm5.submit();" >value="Nationality">>
'	<input type="button" onClick="document.sortForm51.submit();" value=" + ">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm51.submit();">
	</form>
	<br>
	<form name="sortForm52" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="isNull(scty1,pcty1) ASC,cand_lnam_t,cand_fnam_t">
	<%'<input src="../images/<%=session("template_org_code")>_admin/nationality.gif" type="image" onClick="document.sortForm5.submit();" >value="Nationality">>
'	<input type="button" onClick="document.sortForm52.submit();" value=" - ">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm52.submit();">
	</form>
	</th>
<%'<!----------------------- CONSERVE SPACE IF NOT GENERAL VIEW ----------------->
'If  instr(Request.querystring("Q"),"EvalApp") then


'else%>
	<th class="textsmall" valign="top"  align=left width="<%=pv_width6%>">Applied<br><br>
	<form name="sortForm61" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value=" candjob_crd DESC">
	<%'<input src="../images/<%=session("template_org_code")>_admin/submitted.gif" type="image" onClick="document.sortForm6.submit();" value="Submitted">>
'	<input type="button" onClick="document.sortForm61.submit();" value=" + ">%>
	<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm61.submit();">

	</form>
	<br>
	<form name="sortForm62" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value=" candjob_crd ASC">
	<%'<input src="../images/<%=session("template_org_code")>_admin/submitted.gif" type="image" onClick="document.sortForm6.submit();" value="Submitted">>
'	<input type="button" onClick="document.sortForm62.submit();" value=" - ">%>
	<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm62.submit();">
	</form>
	</th>
<%'end if%>

	<th class="textsmall" valign="top"  align=center width="<%=pv_width7%>">Cov<br>Let	
	<form name="sortForm515" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="candtext_id_c DESC, cand_lnam_t, cand_fnam_t">
	<%'<input src="../images/<%=session("template_org_code")>_admin/nationality.gif" type="image" onClick="document.sortForm5.submit();" >value="Nationality">>
'	<input type="button" onClick="document.sortForm511.submit();" value=" + ">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm515.submit();">
	</form>
	<br>
	<form name="sortForm516" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="candtext_id_c ASC, cand_lnam_t, cand_fnam_t">
	<%'<input src="../images/<%=session("template_org_code")>_admin/nationality.gif" type="image" onClick="document.sortForm5.submit();" >value="Nationality">>
'	<input type="button" onClick="document.sortForm512.submit();" value=" - ">%>
	<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm516.submit();">
	</form>
QA</th>

<%'22 JUN 10 LJL added overallrank eval total%>
	<th class="textsmall" valign="top"  align=left width="<%=pv_width8%>">
	Auto<br>Eval<br>
	<form name="sortForm511" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="candjob_overallrank_c DESC,cand_lnam_t,cand_fnam_t">
	<%'<input src="../images/<%=session("template_org_code")>_admin/nationality.gif" type="image" onClick="document.sortForm5.submit();" >value="Nationality">>
'	<input type="button" onClick="document.sortForm511.submit();" value=" + ">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm511.submit();">
	</form>
	<br>
	<form name="sortForm512" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="candjob_overallrank_c ASC,cand_lnam_t,cand_fnam_t">
	<%'<input src="../images/<%=session("template_org_code")>_admin/nationality.gif" type="image" onClick="document.sortForm5.submit();" >value="Nationality">>
'	<input type="button" onClick="document.sortForm512.submit();" value=" - ">%>
	<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm512.submit();">
	</form>


	</th>

<%'<!------------------ ADD ORG INFO 2000 check for 3000 --------------------->
' 05 MAR 08 LJL added rights check per ILO
if   instr(Session("rightsgroup"),",12,") then
	If  instr(Request.querystring("Q"),"EvalApp") then%>
	<th class="textsmall" valign="top"  align=center colspan="2">Rank per this post
		
		<%'29 APR 15 LJL no overall ranking if don't have the right 133
if   instr(Session("rightsgroup"),",133,") then%>

		<br>Overall Rank of Applicant</th>
<%'29 APR 15 LJL no overall ranking if don't have the right 133
end if	
	else
	%>
	<th class="textsmall" valign="top"  align=center width="<%=pv_width9%>">Vac Rank<br>
		
<%'29 APR 15 LJL no overall ranking if don't have the right 133
if   instr(Session("rightsgroup"),",133,") then%>

	Overall Rank
<%
'29 APR 15 LJL no overall ranking if don't have the right 133
end if
end if%>

	<% ' Added by INT on 7-Feb-08 for drop down for selcting type of candidates (eraps)
	'08 JUN 10 LJL moved to here to stop craziness for ILO page - was messing up th's above
	'08 JUN 10 LJL moved to this position to conserve space on ILO site
if session("template_org_code") = 2000  and instr(Request.querystring("Q"),"UpdOnly") then%>
	<br>
<%  ' 15 mar 08 - ac - grab eRAPS VN flag here
	if ERapsVN = "2" then%>
'	<form name="candtypeForm" action="hrd-cllist.asp?<%=followurl &"X=1"%>"  method="post">
		<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
		<input type="hidden" name="sort" value="candjob_gridtotal_c DESC, cand_lnam_t ASC, cand_fnam_t ASC">
		<input type="hidden" name="test1" value="">
		<!---<input class="textsmall" type="submit" onclick="document.sortForm8.submit();" value="[-]" id=submit1 name=submit1>--->
		<%
			lcand_type = 0
			if len(request.QueryString ("cand_type")) then
				lcand_type = request.QueryString("cand_type")
			end if

		 dim selvalue
		 selvalue = 0
	'	 response.write "<br>cur status is" & VacInfo("status_id_c")
		 if trim(Request.QueryString("cand_type")) <> ""  then
		 	 selvalue = trim(Request.QueryString("cand_type"))
		 elseif trim(Request.Form("candydateType"))  <> "" then
			selvalue =trim(Request.Form("candydateType"))
		 elseif trim(Request.Form("candydateType")) = "" and trim(Request.QueryString("cand_type")) = "" then
		 	if VacInfo("status_id_c") = "63" or  VacInfo("status_id_c") = "64" then
		 		selvalue = "1"  ' Under Rep.
			elseif VacInfo("status_id_c") = "65" or  VacInfo("status_id_c") = "66" then
			 	selvalue = "2"  ' adj Rep.
			elseif VacInfo("status_id_c") = "67" or  VacInfo("status_id_c") = "68" then
			 	selvalue = "3"  ' Over Rep.
			else
				selvalue = "4" ' all
			end if
		end if

		' Response.Write "<br>selvalue= " & selvalue
			 %>
		<Select width="20" STYLE="width: 70px" id="submit1" NAME="candydateType" <%if isHM =1 then %> onChange="this.form.submit();" <%end if%> >
				<%Do while GetErapSht.eof = false %>
				<%if ((isHM =1) and GetErapSht("sht_id") <> "4") or (isRU = 1) then %>
				 <OPTION maxlength="18"  VALUE="<%=GetErapSht("sht_id")%>"<%if selvalue = trim(GetErapSht("sht_id")) then %>SELECTED<%elseif (lcand_type <> 0) then  if (lcand_type = trim(GetErapSht("sht_id"))) then%>selected  <%end if  end if%>><%=GetErapSht("sht_dsc_en")%>
				 <%end if%>
				 <%GetErapSht.movenext
	               loop%>
		</SELECT>
		<br>
		</form>
	 <% end if
	 end if%>
</th>


<%end if%>



<th colspan="<%=pv_width_table_cols%>" align="left" width="<%=pv_width_table%>" valign="top">



	
<table  border="0" bordercolor="navy" width="100%" cellpadding="0" cellspacing="0">
	<tr class ="NonPrintable">
    
	<th class="textsmall NonPrintable" valign="top" width="<%=pv_width10%>" align=center nowrap>
	Screen 
	<Br>
        <!-- 31-08-16 to show when en mass update link click --> <br> 
    <%if  pv_updonlyindic =1 then  %>
        
                <input type="checkbox" class="Screen_All"  />
        <%end if %>
         <br />

	<form name="sortForm711" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_screened_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm711.submit();">
	</form>
	<br>
	<form name="sortForm712" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_screened_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm712.submit();">
	</form>

</th>

<%'26 SEP 10 LJL not show prescreen items if not set for it
if pv_showprescreen = 1 then%>
<%'25 SEP 10 LJL  ADDED 3 preselection items per ITU and others%>

<th class="textsmall NonPrintable" valign="top" width="<%=pv_width11%>" align=center nowrap>
	<%if session("template_org_code") = 2400 then%>
	Pre<Br>Select
	<%else%>
	Pre<br>screen1 <br />  <!-- 31-08-16 to show when en mass update link click -->
     <%if  pv_updonlyindic =1 then  %><input type="checkbox" class="Screen_All1"  /> <%end if%> <br />
	<%end if%>

	<form name="sortForm811" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_prescreen1_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm811.submit();">
	</form>
	<br>
	<form name="sortForm812" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_prescreen1_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm812.submit();">
	</form>

</th>

<th class="textsmall NonPrintable" valign="top" width="<%=pv_width12%>" align=center nowrap>
	<%if session("template_org_code") = 2400 then%>
	Edu<br>Req
	<%else%>
	Pre<Br>screen2 <br />   <!-- 31-08-16 to show when en mass update link click -->
    <%if  pv_updonlyindic =1 then  %><input type="checkbox" class="Screen_All2"  /> <%else %> <br /> <%end if%>
	<%end if%>
	<form name="sortForm911" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_prescreen2_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm911.submit();">
	</form>
	<br>
	<form name="sortForm912" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_prescreen2_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm912.submit();">
	</form>
</th>

<th class="textsmall NonPrintable" valign="top" width="<%=pv_width13%>" align=center nowrap>
	<%if session("template_org_code") = 2400 then%>
	Lang<br>Req
	<%else%>
	Pre<br>screen3 <br /> <!-- 31-08-16 to show when en mass update link click -->
     <%if  pv_updonlyindic =1 then  %><input type="checkbox" class="Screen_All3" />  <%end if%>   <br />
	<%end if%>
	<form name="sortForm101" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_prescreen3_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm101.submit();">
	</form>
	<Br>
	<form name="sortForm102" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_prescreen3_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm102.submit();">
	</form>
</th>

<%'26 SEP 10 LJL not show prescreen items if not set for it
end if%>

	<th class="textsmall NonPrintable" valign="top" width="<%=pv_width14%>" align=center nowrap>
	Short<br>list <br />  <!-- 31-08-16 to show when en mass update link click -->
    <%if  pv_updonlyindic =1 then  %><input type="checkbox" class="Screen_All4" />  <% end if%> <br />

	<form name="sortForm111" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_shortlist_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm111.submit();">
	</form>
	<br>
	<form name="sortForm112" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_shortlist_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm112.submit();">
	</form>
</th>


<%'<!---------------------------- ADD ORG INFO 1000 2000 4000 - 3000 removed per WTO req 35 ---------------------------------------------->

'06 DEC 11 LJL remove online testing column for IFRC, per VANELDEREN
'26 AUG 15 LJL add back in online testing for WTO
'session("template_org_code") <> 3000 AND 
if session("template_org_code") <> 7000 then
'<!---------------------------- TESTED APPLICANTS --------------------->%>
	<th class="textsmall NonPrintable" valign="top" width="<%=pv_width15%>" align=center nowrap>
	Test
    <br><br> <!-- 31-08-16 to show when en mass update link click -->
    <%if  pv_updonlyindic =1 then  %><input type="checkbox" class="Screen_All5"  /> <% end if %> <br />
	<form name="sortForm121" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>"><%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="candjob_tested_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm121.submit();">
	</form>
	<br>
	<form name="sortForm122" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>"><%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<input type="hidden" name="sort" value="candjob_tested_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm122.submit();">
	</form>
	</th>
<%end if
'<!---------------------------- INTERVIEWED APPLICANTS --------------------->%>
	<th class="textsmall NonPrintable" valign="top" width="<%=pv_width16%>" align=center nowrap>
	Inter<br>viewed <br /><!-- 31-08-16 to show when en mass update link click -->
  <%if  pv_updonlyindic =1 then  %><input type="checkbox" class="Screen_All6"  /> <%end if %> <br />

	<form name="sortForm131" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_interv_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm131.submit();">
	</form>
	<br>
	<form name="sortForm132" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_interv_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm132.submit();">
	</form>

</th>
<th class="textsmall NonPrintable" valign="top" width="<%=pv_width17%>" align=center nowrap>
	Recm'd<br><br><!-- 31-08-16 to show when en mass update link click -->
  <%if  pv_updonlyindic =1 then  %><input type="checkbox" class="Screen_All7"  /> <% end if %> <br />
	<form name="sortForm141" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_recomm_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm141.submit();">
	</form>
	<br>
	<form name="sortForm142" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_recomm_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm142.submit();">
	</form>

</th>
<th class="textsmall NonPrintable" valign="top" width="<%=pv_width18%>" align=center nowrap>
	Select<br><Br><!-- 31-08-16 to show when en mass update link click -->
   <%if  pv_updonlyindic =1 then  %><input type="checkbox" class="Screen_All8"  /> <% end if %> <br />
	<form name="sortForm151" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_selected_i DESC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-up.gif" onClick="document.sortForm151.submit();">
	</form>
	<br>
	<form name="sortForm152" action="hrd-cllist.asp?<%=followurl%>" method="post">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<input type="hidden" name="sort" value="candjob_selected_i ASC, cand_lnam_t ASC, cand_fnam_t ASC">
<%'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]">%>
<input type="image"  img src="../css/<%=session("template_org_code")%>-css/arrow-down.gif" onClick="document.sortForm152.submit();">
	</form>
</th>
<script src="https://code.jquery.com/jquery-3.1.0.min.js"></script>
 <script src="jquery.stickytableheaders.min.js" type="text/javascript"></script>
<script type="text/javascript">
    $(document).ready(function () {
        $(".Screen_All").click(function () {
                    $('.Spec:enabled').prop('checked', this.checked);
                    
        });

        $(".Screen_All1").click(function () {
                    $('.Spec1:enabled').prop('checked', this.checked);

                });


                $(".Screen_All2").click(function () {
                    $('.Spec2:enabled').prop('checked', this.checked);

                });
                $(".Screen_All3").click(function () {
                    $('.Spec3:enabled').prop('checked', this.checked);

                });
                $(".Screen_All4").click(function () {
                    $('.Spec4:enabled').prop('checked', this.checked);

                });

                $(".Screen_All5").click(function () {
                    $('.Spec5:enabled').prop('checked', this.checked);

                });

                $(".Screen_All6").click(function () {
                    $('.Spec6:enabled').prop('checked', this.checked);

                });

                $(".Screen_All7").click(function () {
                    $('.Spec7:enabled').prop('checked', this.checked);

                });

                $(".Screen_All8").click(function () {
                    $('.Spec8:enabled').prop('checked', this.checked);

                });

                

        
    });
</script>

<%'	<th class="textsmall" valign="top" width="8%" align=center nowrap>
'	<form name="sortForm7" action="hrd-cllist.asp?<%=followurl>" method="post">
'	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid>">
'	<input type="hidden" name="sort" value="candjob_gridtotal_c DESC, cand_lnam_t ASC, cand_fnam_t ASC">
'	<input class="textsmall" type="submit" onClick="document.sortForm8.submit();" value="[-]"><br>Grid<br>score
'	</form></th>
%>


	</tr>
	</table>
	</th>
</tr>
</thead>

    
<tbody>
       
<TR>
<%
' **********************************************
' BEGIN FORM MAKER 1
' **********************************************
If instr(Request.querystring("Q"),"Matrixonly") then
'<!-------------- removed <cfif gCandList.recordcount GT 250> target="_blank"end if 13 OCT 03 ---------------->%>
<form name="pdfForm" id="pdfForm" action="rsys_matrix.asp?requesttimeout=20000" method="post">
				<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
				
				<INPUT TYPE="hidden" NAME="ByJ" VALUE="1">
				<INPUT TYPE="hidden" NAME="cand_type" VALUE=0>
				<INPUT TYPE="hidden" NAME="sort" VALUE="upper(cand_lnam_t),upper(cand_fnam_t)">
						<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
<%elseif instr(Request.querystring("Q"),"MoveToPosts")  then
'<!--- Added by LJL on July 2001 to generate PDF files --->%>
<form name="pdfForm" id="pdfForm" action="hrd-cllist.asp?Q=ByJ&Q=MoveToPosts&requesttimeout=5000" method="post">
		<INPUT TYPE="hidden" NAME="vacchoice" VALUE="<%= newjobid%>">
		<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
		<INPUT TYPE="hidden" NAME="cand_type" VALUE=0>
		<input type="hidden" name="count" value="<%= gCandList.RecordCount%>">
<%ELSEIF  instr(Request.querystring("Q"),"EvalApp") then
'<!--- Added by LJL on July 2001 to generate PDF files --->%>
<form name="pdfForm" id="pdfForm" action="hrd-cllist.asp?Q=ByJ&Q=EvalApp&requesttimeout=5000" method="post">
		<INPUT TYPE="hidden" NAME="vacchoice" VALUE="<%= newjobid%>">
		<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
		<INPUT TYPE="hidden" NAME="cand_type" VALUE=0>
		<input type="hidden" name="count" value="<%= gCandList.RecordCount%>">
<%ELSEIF  instr(Request.querystring("Q"),"wordonly") then%>
<form name="pdfForm"  id="pdfForm" action="admin_CVPDF_docsettings.asp?goWord=99&requesttimeout=6000" method="post" target="_blank">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<INPUT TYPE="hidden" NAME="ByJ" VALUE="1">
	<INPUT TYPE="hidden" NAME="cand_type" VALUE=0>
	<INPUT TYPE="hidden" NAME="sort" VALUE="upper(cand_lnam_t),upper(cand_fnam_t)">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<INPUT TYPE="hidden" NAME="vacchoice" VALUE="<%= newjobid%>">
	<input type="hidden" name="count" value="<%= gCandList.RecordCount%>">

<%'03 NOV 10 LJL added HTML output
ELSEIF  instr(Request.querystring("Q"),"HTMLonly") then%>
<form name="pdfForm"  id="pdfForm" action="admin_CVPDF_docsettings.asp?goHTML=99&requesttimeout=6000" method="post" target="_blank">
	<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<%= newjobid%>">
	<INPUT TYPE="hidden" NAME="ByJ" VALUE="1">
	<INPUT TYPE="hidden" NAME="cand_type" VALUE=0>
	<INPUT TYPE="hidden" NAME="sort" VALUE="upper(cand_lnam_t),upper(cand_fnam_t)">
	<%'<INPUT TYPE="hidden" NAME="Q" VALUE="<%=request.querystring("Q")">%>
	<INPUT TYPE="hidden" NAME="vacchoice" VALUE="<%= newjobid%>">
	<input type="hidden" name="count" value="<%= gCandList.RecordCount%>">

<%elseif  instr(Request.querystring("Q"),"UpdOnly") then%>
<form name="pdfForm" id="pdfForm"  action="hrd-cllist-updates.asp?Q=ByJ&Q=UpdOnly&requesttimeout=5000&interviewer=1" method="post" onsubmit="return DataValidation();">
<%'<form name="pdfForm" action="hrd-cllist.asp?Q=ByJ&Q=UpdOnly&requesttimeout=5000&interviewer=1" method="post">%>
		<INPUT TYPE="hidden" NAME="vacchoice" VALUE="<% response.write newjobid%>">
		<INPUT TYPE="hidden" NAME="jobinfo_uid_c" VALUE="<% response.write newjobid%>">
		<INPUT TYPE="hidden" NAME="cand_type" VALUE=0>
		<input type="hidden" name="count" value="<%= gCandList.RecordCount%>">
<%elseif  instr(Request.querystring("Q"),"Commonly") then
'<!--- Added by LJL on July 2001 to generate PDF files --->%>
<form name="pdfForm"  id="pdfForm" action="rsys-email-prep.asp?EMAWHOPT=1" method="post" target="_blank">
		<INPUT TYPE="hidden" NAME="vacchoice" VALUE="<% response.write newjobid%>">
		<input type="hidden" name="count" value="<%= gCandList.RecordCount%>">
		<INPUT TYPE="hidden" NAME="cand_type" VALUE=0>

<%else
'<!--- Added by LJL on July 2001 to generate PDF files --->%>
	<form name="pdfForm" id="pdfForm" action="admin_CVPDF_docsettings.asp" method="post" target="_blank">
	<INPUT TYPE="hidden" NAME="goPDF" VALUE="99">
	<INPUT TYPE="hidden" NAME="vacchoice" VALUE="<% response.write newjobid%>">
	<input type="hidden" name="count" value="<%=gCandList.RecordCount%>">
	<INPUT TYPE="hidden" NAME="cand_type" VALUE=0>
<%
' NEEDNEEDNEED recordcount above changed

end if


'21 JUN 10 LJL end goExcel portion
end if

'OUTOUTOUT

If instr(Request.querystring("Q"),"Matrixonly")  then
		butty = "Matrix"
elseif instr(Request.querystring("Q"),"MoveToPosts")  then
		butty = "Add"
elseif  instr(Request.querystring("Q"),"Commonly") then
		butty = "Select"
elseif  instr(Request.querystring("Q"),"wordonly") then
		butty = "Select"
elseif  instr(Request.querystring("Q"),"HTMLonly") then
		butty = "Select"
else
	butty = "Select"
end if

	 if GCandList.eof = true and  session("template_org_code")= 2000  then%>

	<table width="100%">
	<tr width="100%">

		<span class="alert" >No applicants meet the criteria you specified for this post or none have applied as of yet</span>

	 </tr>
	 </table>
<% Response.end
   end if

' **********************************************
' END FORM MAKER 1
' **********************************************
%>
	<td colspan="0" style="float:left"><input class='textsmall' type="Button" value="<% response.write butty%>" onClick="unCheck();"></td>
	
</tr>


<%NAMER = ""
	emailcheck = ""
	japbirth = ""
	candcounter = 0
	pv_screener = 0
'ac nov max at 200
Do while gCandList.eof = false 

if (candsto <> 0 and candsto-rowsCountConst > 0 and candsto-rowsCountConst > candcounter) then
	Do while candsto-rowsCountConst > candcounter
		candcounter = candcounter + 1
		gCandList.movenext
	Loop
End If

pv_screener = gCandList("candjob_screened_i")

'if (candsto <> 0 and candsto-rowsCountConst < candcounter) or (candsto = 0) then

if session("template_org_code") = 2000 then

 '03 mar 08 ac -  added the check against a post number as the candid_c is not unique
   GetColorSql  = " select e.sht_id,e.sht_color,t.candjob_hm_review,candjob_ru_review from  eraps_sht e  Inner join  tx_rsys_candjob  t on  (e.sht_id = t.candjob_hm_review or e.sht_id = t.candjob_ru_review )  and  t.cand_id_c =" & gcandlist("cand_id_c") & " and jobinfo_uid_c = " & newjobid

  ' response.write "<br>GetColorSql = " & GetColorSql
   Set GetColor =   rsys_db_select.execute(GetColorSql)
end if

'27 Nov 06 ac added flush buffer to stop the script crashing -  reduce to every 50 records
' releases output as is, so every 100 records it pushes to browser.
	if (candcounter mod 50) = 0 then
		response.Flush
	end if
	candcounter = candcounter + 1
NAMER2 = gCandList("cand_lnam_t")&gCandList("cand_fnam_t")
	emailcheck2 = gCandList("email")
	If Request.querystring("linenum") AND candcounter = linenum then%>
	<TR bgcolor="#800080">
  <%elseif NAMER = NAMER2 AND japbirth = gCandList("cand_bth_d") then
  if emailcheck <> emailcheck2  then%>
  			<TR bgcolor="yellow">
	<%else%>
    			<TR bgcolor="#ffccff">
	<%end if
	else%>
<tr <%
Dim x, bgcolor
 if x = 1 then
    Response.write "bgcolor='#EEEEEE'"
     x=2
Else
     response.write "bgcolor='#ffffff'"
    x=1
End if%>>
<%end if

	'20 OCT 09 LJL added to set var for hrd-cllist instead of check each time.
		if gcandlist("is_thisorg_staff") <> "" then
			pv_fullinternal = 1
		else
			pv_fullinternal = 0
		end if


'        	<!--------------------- ADD ORG INFO 3000 10 MAR 05 LJL req 31a
'
'
'        if upd > candjob_crd then
'
'        		<td valign="top"  nowrap bgcolor="00ff00">
'          	else
'          		<td valign="top"  nowrap>
'            	end if
'            	-------------------->
'20 NOV 06 ADDED THIS FOR THE CHECKING PER LINE
'response.write "TRUE99: " & candcounter & "::" & request.form("candlist1") & ":::" & request.form("candlist2")  &"<Br>"
' 03 NOV 06 LJl added check to see if candlist1 and 2 are numeric. If not, don't allow. (to check mark applicants in a list)
if (len(request.form("candlist1")) AND isnumeric(request.form("candlist1"))) AND (len(request.form("candlist2")) AND isnumeric(request.form("candlist2"))) then
	if request.form("candlist1") <= "0" AND request.form("candlist2") <= "0" then
		'response.write "TRUE4<Br>"
		pv_checkedit = " CHECKED"
	elseif (int(candcounter) >= int(request.form("candlist1"))) AND (int(candcounter) <= int(request.form("candlist2"))) then
		'response.write "TRUE5<Br>"
		pv_checkedit = " CHECKED"
	else
		'response.write "TRUE6<Br>"
		pv_checkedit = ""
	end if
else
	pv_checkedit = " CHECKED"
	'response.write "TRUE7<Br>"
	'response.write "<font class='alert'>Please enter numeric values for the applicant number selection fields.</font><br>"
end if
' HAD TO SET TO TRIM EACH, or doesn't work%>
	<td valign="top"  nowrap align="center" width="<%=pv_width1%>">
	<input type="Checkbox" name="japids" value="<%= gcandlist("cand_id_c")%>"<%=pv_checkedit%>>
	<% response.write candcounter%>
	<br>
	<%	'              		<!------------------ ADD ORG INFO 2000 check for 3000 --------------------->
	'              <!------------------------- BEGIN RANK OUTPUT ---------------------------------->
' 05 MAR 08 LJL added rights check per ILO
if   instr(Session("rightsgroup"),",12,") then
    if session("template_org_code")= 2000 then%>
		<br>
		<%if NOT pv_fullinternal = 1 then
              cranktot = 0
              cranker2 = 0
              cranker3 = 0
              if gcandlist("crank1") > 0 then
              cranker1 = gcandlist("crank1")
              cranktot = cranktot + 1
              else
              cranker1 = 0
              end if
              if gcandlist("crank2") > 0 then
              cranker2 = gcandlist("crank2")
              cranktot = cranktot + 1
              else
              cranker2 = 0
              end if
              if gcandlist("crank3") > 0 then
              cranker3 = gcandlist("crank3")
              cranktot = cranktot + 1
              else
              cranker3 = 0
              end if
              if cranktot = 0 then
              cranktot = 1
              end if
              cranker = ((cranker1 + cranker2 +cranker3)/cranktot)
              if NOT isnumeric(gcandlist("rank_official_i")) then
              rank_official_i = 0
			  else
			  rank_official_i = gcandlist("rank_official_i")
              end if
              if NOT isnumeric(cranker) then
              cranker = 0
              end if
			ranktotal = cranker + rank_official_i
			if ranktotal > 0 then
				ranktotal = Round(ranktotal,2)
			end if
			response.write "<b>" & ranktotal  &"</b>"
		end if
	end if
end if
'<!------------------------- END RANK OUTPUT ---------------------------------->
%>
	</td>

	<td valign="top" align="left" width="<%=pv_width2%>">
  		<a class ="newanchor" href='hrd-cand-info.asp?cand_id_c=<% response.write gcandlist("cand_id_c")%>' target="_blank"><strong><% if gcandlist("is_thisorg_staff") > "0" then
			if len(gcandlist("lastname")) > 15 then
				response.write trim(left(UCASE(gcandlist("lastname")),15)) & ".."
			else
				response.write trim(UCASE(gcandlist("lastname")))
			end if
			else
				if len(gcandlist("cand_lnam_t")) > 15 then
					response.write trim(left(UCASE(gcandlist("cand_lnam_t")),15)) & ".."
				else
					response.write trim(UCASE(gcandlist("cand_lnam_t")))
				end if
			end if	%></strong></a><br><%
			If gcandlist("is_thisorg_staff") > "0" then
			  if len(gcandlist("firstname")) > 15 then
		  		response.write trim(left(UCASE(gcandlist("firstname")),15)) & ".."
			  else
  				response.write trim(UCASE(gcandlist("firstname")))
			  end if
		  	else
				if len(gcandlist("cand_fnam_t")) > 15 then
			  		response.write trim(left(UCASE(gcandlist("cand_fnam_t")),15)) & ".."
				else
		  			response.write trim(UCASE(gcandlist("cand_fnam_t")))
  				end if
	 		end if
 if gcandlist("cand_archived_i") = "1" then
  '19 nov 06 ac changed single quote to double - was causing next closing cell tag to fail to be read correclty by browser %>
  <br> <font color="maroon">[arch]</font>
  <%
  end if
  %></td>
<!--'28 april 2018 manoj fix the space issue  -->
<td valign="top"  align=left class="textsmall" width="<%=pv_width3%>">

  <%
  if pv_fullinternal = 1 then
  response.write gcandlist("sexs")
	else
  response.write gcandlist("sexc")
  end if
  %>
  <br>
  <% ' 		<!-------- REVISED from SHORT TO WHO STAFF not in HQ ------->
	if gcandlist("thisorg_short") = "1" OR gcandlist("thisorg_staffno") > "0" then
	'11 JUN 12 LJL made org name 6 chars instead of 3, as WIPO looked wrong in this list%>
  		<font color=green><strong><% response.write left(session("template_org_name"),6)%></strong></font>
  		<br>
	<%elseif gcandlist("cand_io_i") = 1 then%>
			<font color=#0000cd><strong>IntOrg</strong></font>
			<br>
	<%elseif gcandlist("cand_io_i") = 0 then
	end if

'21 FEB 12 LJL not show Ex for ITU for internals/externals
if session("template_org_code") <> 2400 then

if pv_fullinternal = 1 then
'20 OCT 09 LJL added WHO/UNAIDS/ICC check to distinguish GSM internals
	if session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500 then
		response.write "<font color=red>GSM</font>"
	else
		response.write "Int"
	end if%>
			<b><font color="green">
	<%if len(gcandlist("staff_type")) then%>
		<strong><%UCASE(trim(gcandlist("staff_type")))%></strong>
	<%end if%>
	</font></b>
<%else%>
	<font color=navy>Ex</font>
<%end if

'21 FEB 12 LJL not show Ex for ITU for internals/externals
end if%>

</td>
<td valign="top"  align=left nowrap class='textsmall' width="<%=pv_width4%>">

  <%if isdate(gcandlist("cand_bth_d")) then
  			response.write day(gcandlist("cand_bth_d")) & "-" & monthname(month(gcandlist("cand_bth_d")),1) & "-" & right(year(gcandlist("cand_bth_d")),2)
  			yo_count = yo_count + 1
	else%>
  	N/A
<%yo_count = yo_count
	end if%>
  	<br>
  	Age
	<%
	if isdate(gcandlist("cand_bth_d")) then
	yo = datediff("yyyy",gcandlist("cand_bth_d"),now())
	'response.write "BDAY: " & gcandlist("cand_bth_d") & "|" & now()
  	'<!------------ SET THE AVERAGE AGE ONLY IF NOT INTERNAL STAFF FOR NOW ----------------->
			yo_total = yo_total + yo
			response.write yo
  		else
  			response.write "N/A"
  		end if	%>
  </td>
		<td valign="top"  align="left" nowrap class='textsmall' width="<%=pv_width5%>">
<%'<!------------------ ADD ORG INFO 2000 check for 3000 --------------------->
if session("template_org_code")= 2000 then
	'<!-------------QUOTA LIST'- this list is pulled from the database and then the c rank is found in the number list and then that index placement in that index is pulled from teh Alpha ranking list
	if pv_fullinternal = 1 then%>
		<font color="silver"><i>Int</i></font>
		<%IF gCandList("crank1") <= 0 OR gCandList("crank1") > 5 then
		else
			'NEEDNEEDNEED revise this area
			cra1 = listfind(variables.cranklist,crank1)	& "/"
			'#listgetat(variables.cranklist_alpha,variables.cra1)#
		end if
		IF gCandList("crank2") <= 0 OR gCandList("crank2") > 5 then
		ELSE
			cra2 = listfind(variables.cranklist,crank2)	& "/"
			 '#listgetat(variables.cranklist_alpha,variables.cra2)#
		end if
		IF gCandList("crank3") <= 0 OR gCandList("crank3") > 5 then
		ELSE
			cra3 = listfind(variables.cranklist,crank3)	& "/"
			 '#listgetat(variables.cranklist_alpha,variables.cra3)#
		end if
	end if
	response.write "<br>"
end if

	'<!----------------------------------------- ADD ORG INFO 1000 and others check for 3000 23 FEB 05 LJL ---------------->
	'29 SEP 12 LJL added nat quota output for UPU 5500
	if session("template_org_code")= 1000 OR session("template_org_code")= 1500 OR session("template_org_code")= 2000 then
		if pv_fullinternal = 1 then
			if Len(gCandList("scty1")) then
				response.write LEFT(gCandList("scty1"), 9)
			else%>
				<font color="silver">--</font>
			<%end if
		else
			if Len(gCandList("pcty1")) then
				response.write LEFT(gCandList("pcty1"), 9)
			else%>
				<font color="silver">--</font>
			<%end if
			if Len(gCandList("pcty2")) then%>
				<br>
			<%end if
			response.write LEFT(gCandList("pcty2"), 9)
			if Len(gCandList("pcty3")) then%>
			<br>
			<%end if
			response.write LEFT(gCandList("pcty3"), 9)
		end if
'29 SEP 12 LJL added all nat quota indications for UPU 5500		
	elseif session("template_org_code")= 5500 then
		if pv_fullinternal = 1 then
			if Len(gCandList("scty1")) then
				response.write LEFT(gCandList("scty1"), 9)
			else%>
			<font color="silver">--</font>
				<%end if
		else
			if Len(gCandList("pcty1")) then
				response.write LEFT(gCandList("pcty1"), 9) & " - " & gCandList("pcq1")
			else%>
				<font color="silver">--</font>
			<%end if
			if Len(gCandList("pcty2")) then
				response.write "<br>" & LEFT(gCandList("pcty2"), 9) & " - " & gCandList("pcq2")
			end if
			if Len(gCandList("pcty3")) then
				response.write "<br>" & LEFT(gCandList("pcty3"), 9) & " - " & gCandList("pcq3")
			end if
		end if
	else
		if pv_fullinternal = 1 then
			if Len(gCandList("scty1")) then
				response.write LEFT(gCandList("scty1"), 9)
			else%>
			<font color="silver">--</font>
				<%end if
		else
			if Len(gCandList("pcty1")) then
				response.write LEFT(gCandList("pcty1"), 9)
			else%>
				<font color="silver">--</font>
			<%end if
			if Len(gCandList("pcty2")) then%><br>
			<%end if
			response.write LEFT(gCandList("pcty2"), 9)
			if Len(gCandList("pcty3")) then%><br>
			<%end if
				response.write LEFT(gCandList("pcty3"), 9)
		end if
	end if
	
	'else
'	if gCandList("is_thisorg_staff") >= 1 then
'		if len(gCandList("nat")) >= 9 then
'			response.write left(gCandList("nat"),9) & "..."
'		else
'			response.write gCandList("nat")
'		end if
'	else
'		if session("template_org_code") <> 7000 then%
'			<strong><% response.write gCandList("pcq1")></strong>
'			<br>
'		<%end if
'		if len(gCandList("nat")) >= 9 then
'			response.write left(gCandList("nat"),9) & "..."
'		else
'			response.write gCandList("nat")
'		end if
'	end if
'end if%>
	</td>
	<td valign="top" align="left"  nowrap class='textsmall' width="<%=pv_width6%>">
<%'<!----------------------- CONSERVE SPACE IF NOT GENERAL VIEW ----------------->
'If  instr(Request.querystring("Q"),"EvalApp") then
'	response.write "///"
'else

	if Len(gCandList("candjob_crd")) then%>
		<%	'<!--------------------- ADD ORG INFO 3000 10 MAR 05 LJL req 31a 	-------------------->
		if gCandList("upd") > gCandList("candjob_crd") then
			response.write day(gcandlist("candjob_crd")) & "-" & monthname(month(gcandlist("candjob_crd")),1) & "-" & right(year(gcandlist("candjob_crd")),2)
		%><Br>
			<a  class ="NonPrintable" href='admin-cand-view-updates.asp?adminjap=<% response.write gCandList("cand_id_c")%>&upd_d=<% response.write gCandList("upd")%>&updtype=post' target="_blank"><font color='orange'>Upd!</font></a>
		<%else
			response.write day(gcandlist("candjob_crd")) & "-" & monthname(month(gcandlist("candjob_crd")),1) & "-" & right(year(gcandlist("candjob_crd")),2)
		end if
	else
    	if len(gCandList("upd")) then
  			response.write day(gcandlist("upd")) & monthname(month(gcandlist("upd")),1) & right(year(gcandlist("upd")),2)
		end if
	end if
'end if%>
	</td>


<%'<!------------------- COVERING LETTER REVISED 14 JUL 04 LJL ------------------------------>%>

    
<td valign="top"  align=center class='textsmall' width="<%=pv_width7%>">
  <% if len(gCandList("candtext_id_c")) then%>
	<a class ="NonPrintable" href='appDoc-adminviewclose.asp?candtext_id_c=<% response.write gCandList("candtext_id_c")%>&adminjap=<% response.write gCandList("cand_id_c")%>' target="_blank"><strong><font color="red">CL</font></strong></a><br>
  <%end if%>
  <%'08 AUG 15 LJL REMOVED EXTRA NONSENSE - cand-list8, and then odd reference to query on that page
	   if GETJQ.eof = false then%>
  	<a class ="NonPrintable" href='rsys-question-response.asp?adminjap=<% response.write gCandList("cand_id_c")%>&job_id=<% response.write newjobid%>' target="popupWindow" onClick="window.open('','popupWindow','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=700,height=400,left=200,top=100');">QA</a>
  <%'08 AUG 15 LJL REMOVED EXTRA NONSENSE - cand-list8, and then odd reference to query on that page
	  end if%>
  </td>

<%'22 JUN 10 LJL added overallrank totaler
pv_overalltotaler = gCandlist("candjob_overallrank_c")

				if pv_overalltotaler >= 85 AND pv_overalltotaler <= 100 then
					response.write "<td align=left bgcolor=#BAFEA3 width=" & pv_width8 & ">&nbsp;" & pv_overalltotaler & "</td>"
				elseif pv_overalltotaler >= 70 AND pv_overalltotaler < 85 then
					response.write "<td align=left bgcolor=#FFFFB5  width=" & pv_width8 & ">&nbsp;" & pv_overalltotaler & "</td>"
				elseif pv_overalltotaler >= 50 AND pv_overalltotaler < 70 then
					response.write "<td align=left bgcolor=#FFD7B3 width=" & pv_width8 & ">&nbsp;" &pv_overalltotaler & "</td>"
				elseif pv_overalltotaler < 50 AND pv_overalltotaler > 1 then
					response.write "<td align=left bgcolor=#EDA9BC width=" & pv_width8 & ">&nbsp;" & pv_overalltotaler & "</td>"
				else
					response.write "<td align=left bgcolor=silver width=" & pv_width8 & ">&nbsp;" & pv_overalltotaler & "</td>"
				end if%>

  <%  '<!--------------------------------------- EN MASSE EVALUATIONS ------------------------------------------------->
' 05 MAR 08 LJL added rights check per ILO
if   instr(Session("rightsgroup"),",12,") then
  If  instr(Request.querystring("Q"),"EvalApp") then
  i1 = ""
  i2 = ""%><td>
<table width="<%=pv_width9%>" border="0" bordercolor="green">
<tr>
	<%'  <!----------------- GOT RANK LISTINGS HIGHER UP IN PAGE ----------------------->
    '		<!---------------- THIS POST RANK ----------------------------->
    'Response.write "candrank2: " & gCANDLIST("candrank2") & " <br>" %>
  	<td valign="top"  nowrap class='textsmall' width="<%=pv_width9%>">
	  	
<%'29 APR 15 LJL no overall ranking if don't have the right 133
if   instr(Session("rightsgroup"),",133,") then%>
	  	
	  	<Select  NAME="C2_<% response.write gCandList("cand_id_c")%>">
    <%GETRANK1.movefirst
    Do while GETRANK1.eof = false%>
    <option value="<%=GETRANK1("rankid")%>"<% if gCANDLIST("candrank2") = GETRANK1("rankid") then%> SELECTED<%end if%>><% response.write GETRANK1("rankdsc")
	GETRANK1.movenext
    loop%></select>
    
    <%'29 APR 15 LJL no overall ranking if don't have the right 133
	    end if%>
    </td>
</tr>
<tr>
<%'<!---------------- OVERALL RANK ----------------------------->
'Response.write "candrank1: " & gCANDLIST("candrank1") & " <br>" %>
	<td valign="top"  nowrap class='textsmall' width="<%=pv_width9%>"><Select  NAME="C1_<% response.write gCandList("cand_id_c")%>">
	<%GETRANK1.movefirst
	Do while GETRANK1.eof = false%>
	<option value="<%=GETRANK1("rankid")%>"<% if gCANDLIST("candrank1") = GETRANK1("rankid") then%> SELECTED<%end if%>><% response.write GETRANK1("rankdsc")
	GETRANK1.movenext
	loop%></select></td>
    <input type="hidden" name="candeval_add" value="1">
</tr>
</table>
<%else%>
    		<td valign="top" align="center" nowrap class="textsmall" width="<%=pv_width9%>"><a  class="masterTooltip newanchor" title="<%=gCandList("candjob_rank_rem_m") %>"   href='hrd-cand-rank.asp?adminjap=<% response.write gCandList("cand_id_c")%>&cand_id_c=<% response.write gCandList("cand_id_c")%>&jobinfo_uid_c=<% response.write newjobid%>&candjob_id_c=<% response.write gCandList("candjob_id_c")%>' target="popupWindow"		onclick="window.open('','popupWindow','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=670,height=500,left=200,top=100');">
			<font color="<% response.write gCandList("rank_color_c")%>">
			<% if len(gCandList("rank_dsc")) > 7 then
				response.write left(gCandList("rank_dsc"),7) & ".."
			else
				 response.write gCandList("rank_dsc")
			end if%>
			<% if gCandList("rank_dsc") = "" then%>
			<br>To rank
			<%end if%></font></a>
			<br><a   class="masterTooltip newanchor" title="<%=gCandList("thisorg_rankrem") %>" href='hrd-cand-rank-over.asp?hdroff=1&adminjap=<% response.write gCandList("cand_id_c")%>&cand_id_c=<% response.write gCandList("cand_id_c")%>&jobinfo_uid_c=<% response.write newjobid%>&candjob_id_c=<% response.write gCandList("candjob_id_c")%>' target="popupWindow"		onclick="window.open('','popupWindow','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=670,height=500,left=200,top=100');">
			<font color="<% response.write gCandList("orank_color")%>">
			<% if len(gCandList("orank")) > 7 then
				response.write left(gCandList("orank"),7)
			else
				response.write gCandList("orank")
			end if%>
			<% if gCandList("orank") = "" then%>
			<br>To rank
			<%end if%></font></a>
	 <%end if
	' NEEDNEEDNEED this section is a problem, may be with the if statement or format of the line - may have an extra if/then somewhere which is upsetting this.
      'if gCandList("candedit_available_c") = "1" then%
	  '<br><a href='hrd-assess-cand.asp?getass=1&adminjap=% response.write gCandList("cand_id_c")%' target="_blank"><font color="green'>[Avail]</font></a>
	  '<%
	  'else
	  '	response.write "++"
	  'end if%>

	 	<% 'Added By INT on 08 FEB- displays the checkbox for selection of type of candidates.
	'08 JUN 10 LJL revised for ILO placement - MoVED TO SEPARATE COLUMN
	if session("template_org_code") = 2000 and  instr(Request.querystring("Q"),"UpdOnly") then
	  ' 15 mar 08 - ac - grab eRAPS VN flag here
	   if ERapsVN = "2" then%>
	  <Br>
	  <table>
	  <tr>
	 <td valign="top"  align=center <% if GetColor.eof = false then %>bgcolor=<%=GetColor("sht_color")%><%end if %> width="<%=pv_width9%>">
     <input type="hidden" name="cand_id" value="<%=gcandlist("cand_id_c")%>">
    <span title="Click to set as RAPS shortlisted"><input type="Checkbox" id = "eraps_sht_id" name="eraps_sht_id" value="<%=gcandlist("cand_id_c")%>" <%if gcandlist("candjob_ru_review")<>"Null" then%>CHECKED <%end if%>>
    </td>
    </tr>
    </table>
    <%  end if
	  end if
' END GRID SECTION     %>



</td>



<%end if%>


<td colspan="<%=pv_width_table_cols%>" align="left" width="<%=pv_width_table%>" valign="top">
	
<table border="0" bordercolor="navy" width="100%" cellpadding="0" cellspacing="0">
<tr class ="NonPrintable">


<%'SCREENED COLUMN
'if gCandList("candjob_screened_i") = 1 then
if pv_screener = 1 then
	if gCandList("upd") > gCandList("candjob_screened_d") then%>
	<td class ="NonPrintable" width="<%=pv_width10%>" valign="top" align=center bgcolor="#FFF1C6">Scr
	<br>
	<a  class ="NonPrintable" href='admin-cand-view-updates.asp?adminjap=<% response.write gCandList("cand_id_c")%>&upd_d=<% response.write gCandList("upd")%>&updtype=screen' target='_blank'><font color='red'>Upd!</font></a>
	<%else%>
  	<td class ="NonPrintable" width="<%=pv_width10%>" valign="top" align=center bgcolor="#FFF1C6">Scr
	<%end if
else%>
    <td class ="NonPrintable" width="<%=pv_width10%>" valign="top" align=center>
<%end if
  If pv_updonlyindic = 1 then%>
<Br><span title="Click to set as SCREENED"><input type="Checkbox" class="Spec NonPrintable" onmouseover='' name="screened_japid" value="<% response.write gCandList("cand_id_c")%>"<%if pv_screener = 1 then%> CHECKED<%end if%>></span>
<%else%>
<br>&nbsp;
<%end if%>
</td>

<%'26 SEP 10 LJL not show prescreen items if not set for it
if pv_showprescreen = 1 then%>

    <%' PRESCREEN1 COLUMN
    if gCandList("candjob_prescreen1_i") = 1 then
   	 if gCandList("upd") > gCandList("candjob_prescreen1_d") then%>
	<td width="<%=pv_width11%>" class ="NonPrintable" valign="top" align=center bgcolor="#95FF4F"><%=pv_ps1%>
    <br>
    <a class ="NonPrintable" href='admin-cand-view-updates.asp?adminjap=<% response.write gCandList("cand_id_c")%>&upd_d=<% response.write gCandList("upd")%>&updtype=prescreen1' target="_blank">
	<font color='orange'>Upd!</font></a>
    <%else%>
    <td class ="NonPrintable" width="<%=pv_width11%>" valign="top"  align=center bgcolor="#95FF4F"><%=pv_ps1%>
    <%end if
    else%>
    <td class ="NonPrintable" width="<%=pv_width11%>" valign="top"  align=center>
    <%end if
	' UNSURE
    If instr(Request.querystring("Q"),"UpdOnly") then%>
 		<br><span title="Click to set as <%=pv_ps1%>"><input type="Checkbox" class ="Spec1 NonPrintable" name="prescreen1_japid" value="<% response.write gCandList("cand_id_c")%>"<%if gCandList("candjob_prescreen1_i") = 1 then%> CHECKED<%end if%>></span>
 <%else%>
<br>&nbsp;
		<%end if%></td>


   <%' PRESCREEN2 COLUMN
    if gCandList("candjob_prescreen2_i") = 1 then
   	 if gCandList("upd") > gCandList("candjob_prescreen2_d") then%>
	<td class ="NonPrintable" width="<%=pv_width12%>" valign="top" align=center bgcolor="#D1FFB3"><%=pv_ps2%>
    <br>
    <a class ="NonPrintable" href='admin-cand-view-updates.asp?adminjap=<% response.write gCandList("cand_id_c")%>&upd_d=<% response.write gCandList("upd")%>&updtype=prescreen2' target="_blank">
	<font color='orange'>Upd!</font></a>
    <%else%>
    <td class ="NonPrintable" width="<%=pv_width12%>" valign="top"  align=center bgcolor="#D1FFB3"><%=pv_ps2%>
    <%end if
    else%>
    <td class ="NonPrintable" width="<%=pv_width12%>" valign="top"  align=center>
    <%end if
	' UNSURE
    If instr(Request.querystring("Q"),"UpdOnly") then%>
 		<br><span title="Click to set as <%=pv_ps2%>"><input type="Checkbox" class="Spec2 NonPrintable" name="prescreen2_japid" value="<% response.write gCandList("cand_id_c")%>"<%if gCandList("candjob_prescreen2_i") = 1 then%> CHECKED<%end if%>></span>
 <%else%>
<br>&nbsp;
		<%end if%></td>


   <%' PRESCREEN3 COLUMN
    if gCandList("candjob_prescreen3_i") = 1 then
   	 if gCandList("upd") > gCandList("candjob_prescreen3_d") then%>
	<td class ="NonPrintable" width="<%=pv_width13%>" valign="top" align=center bgcolor="#D2FFC4"><%=pv_ps3%>
    <br>
    <a class ="NonPrintable" href='admin-cand-view-updates.asp?adminjap=<% response.write gCandList("cand_id_c")%>&upd_d=<% response.write gCandList("upd")%>&updtype=prescreen3' target="_blank">
	<font color='orange'>Upd!</font></a>
    <%else%>
    <td class ="NonPrintable" width="<%=pv_width13%>" valign="top"  align=center bgcolor="#D2FFC4"><%=pv_ps3%>
    <%end if
    else%>
    <td class ="NonPrintable" width="<%=pv_width13%>" valign="top"  align=center>
    <%end if
	' UNSURE
    If instr(Request.querystring("Q"),"UpdOnly") then%>
 		<br><span title="Click to set as <%=pv_ps3%>"><input type="Checkbox" class ="Spec3 NonPrintable" name="prescreen3_japid" value="<% response.write gCandList("cand_id_c")%>"<%if gCandList("candjob_prescreen3_i") = 1 then%> CHECKED<%end if%>></span>
 <%else%>
<br>&nbsp;
		<%end if%></td>

<%'26 SEP 10 LJL not show prescreen items if not set for it
end if%>


    <%' SHORTLIST COLUMN
    if gCandList("candjob_shortlist_i") = 1 then
    if gCandList("upd") > gCandList("candjob_shortlist_d") then%>
	<td class ="NonPrintable" width="<%=pv_width14%>" valign="top" align=center bgcolor="#6495ed">SL
    <br>
    <a class ="NonPrintable" href='admin-cand-view-updates.asp?adminjap=<% response.write gCandList("cand_id_c")%>&upd_d=<% response.write gCandList("upd")%>&updtype=shortlist' target="_blank">
	<font color='orange'>Upd!</font></a>
    <%else%>
    <td class ="NonPrintable" width="<%=pv_width14%>" valign="top"  align=center bgcolor="#6495ed">SL
    <%end if
    else%>
    <td class ="NonPrintable" width="<%=pv_width14%>" valign="top"  align=center>
    <%end if

    If instr(Request.querystring("Q"),"UpdOnly") then%>
 		<br><span title="Click to set as SHORT LISTED"><input type="Checkbox" name="shlisted_japid" class="Spec4 NonPrintable" value="<% response.write gCandList("cand_id_c")%>"<%if gCandList("candjob_shortlist_i") = 1 then%> CHECKED<%end if%>></span>
 <%else%>
<br>&nbsp;
		<%end if%></td>


	<%'<!---------------------------- ADD ORG INFO 1000 2000 4000 - 3000 removed per WTO req 35 ---------------------------------------------->
	'06 DEC 11 LJL remove online testing column for IFRC, per VANELDEREN
'26 AUG 15 LJL add back in online testing for WTO
'session("template_org_code") <> 3000 AND 
if session("template_org_code") <> 7000 then

'10 SEP 09 LJL added right for WHO to restrict online testing access
pv_onlinetestright = 0
if session("template_org_code") = 1000 then
	if instr(Session("rightsgroup"),",130,") then
		pv_onlinetestright = 1
	else
		pv_onlinetestright = 0
	end if
else
	pv_onlinetestright = 1
end if

if pv_onlinetestright = 1 then


	if gCandList("candjob_tested_i") = 1 then%>
    <td  class ="NonPrintable" width="<%=pv_width15%>" valign="top"  align=center bgcolor="#ccffcc">Test<br>1. Set
 	<%If instr(Request.querystring("Q"),"UpdOnly") then%>
    	<span title="Click to set for ONLINE TESTING"><input type="Checkbox" name="tested_japid" class="Spec5 NonPrintable" value="<% response.write gCandList("cand_id_c")%>"<%if gCandList("candjob_tested_i") = 1 then%> CHECKED<%end if%>></span>
	<%end if
	 If instr(Request.querystring("Q"),"UpdOnly") then%>Email?
    <input type="Checkbox" class ="NonPrintable" name="emailed_japid" value="<% response.write gCandList("cand_id_c")%>" checked>
    <%end if%>
    </td>
        <%elseif gCandList("candjob_tested_i") = 2 then%>
	<td class ="NonPrintable" width="<%=pv_width15%>" valign="top"  align=center bgcolor="#ffcc99">Test<br>2. Emailed</td>
        <%elseif gCandList("candjob_tested_i") = 3 then%>
	<td class ="NonPrintable" width="<%=pv_width15%>" valign="top"  align=center bgcolor="#ccff66">Test<br>3. Conf'd</td>
        <%elseif gCandList("candjob_tested_i") = 4 then%>
    <td class ="NonPrintable" width="<%=pv_width15%>" valign="top"  align=center bgcolor="#99ff99">Test<br>4. Taken</td>
        <%elseif gCandList("candjob_tested_i") = 9 then%>
    <td class ="NonPrintable" width="<%=pv_width15%>" valign="top"  align=center bgcolor="#9999cc">Test<br>9. Hold-taken/confirmed
        <%elseif gCandList("candjob_tested_i") = 0 AND gCandList("candjob_shortlist_i") = 1 then
        '06 MAR 08 LJL removed 'no' for not yet tested or not to be tested shortlisted
        %>
    <td class ="NonPrintable" width="<%=pv_width15%>" valign="top"  align=center>
	<%If instr(Request.querystring("Q"),"UpdOnly") then%>
		<br>
		<span title="Click to set for ONLINE TESTING">
			<input type="Checkbox" class ="NonPrintable" name="tested_japid" value="<% response.write gCandList("cand_id_c")%>"
				<%if newjobid <> "" then%>
					<%if ((GETJAFINFO1("jobinfo_test_hour_c") <> "") and (GETJAFINFO1("jobinfo_test_minute_c") <> "") and (GETJAFINFO1("jobinfo_test_delaytime_c") <> "") and (GETJAFINFO1("jobinfo_test_lastday_c") <> "") and (GETJAFINFO1("jobinfo_test_contact_c") <> "")) then%>
					<%else%>
						disabled="disabled"
					<% end if %>
				<% end if %>
				<% if gCandList("candjob_tested_i") = 1 then%>
					CHECKED
				<%end if%>>
		</span>
	<%end if%>
	</td>
	<%else%>
	<td class ="NonPrintable" width="<%=pv_width15%>" valign="top"  align=center>
		<%If instr(Request.querystring("Q"),"UpdOnly") then%>
			<br>
			<span title="Click to set for ONLINE TESTING">
			<input class ="NonPrintable" type="Checkbox" name="tested_japid" value="<% response.write gCandList("cand_id_c")%>"
				<%if newjobid <> "" then%>
					<%if ((GETJAFINFO1("jobinfo_test_hour_c") <> "") and (GETJAFINFO1("jobinfo_test_minute_c") <> "") and (GETJAFINFO1("jobinfo_test_delaytime_c") <> "") and (GETJAFINFO1("jobinfo_test_lastday_c") <> "") and (GETJAFINFO1("jobinfo_test_contact_c") <> "")) then%>
					<%else%>
						disabled="disabled"
					<%end if%>
				<%end if%>
				<%
				if gCandList("candjob_tested_i") = 1 then%> CHECKED<%end if%>>
			</span>
		<%end if%>
	</td>
	<%end if

	else
	'10 SEP 09 LJL added right for WHO to restrict online testing access%>
	<td width="<%=pv_width15%>" valign="top"  align=center>No auth</td>

    <%end if
    end if


	if gCandList("candjob_interv_i") = 1 then%>
	<td class ="NonPrintable" width="<%=pv_width16%>" valign="top"  align=center bgcolor="#f0e68c">Intv
	<%else%>
    <td class ="NonPrintable" width="<%=pv_width16%>" valign="top"  align=center>
 	<%end if
    If instr(Request.querystring("Q"),"UpdOnly") then%>
		<br><span title="Click to set as INTERVIEWED"><input type="Checkbox" name="interviewed_japid" class="Spec6 NonPrintable" value="<% response.write gCandList("cand_id_c")%>"
    <%if gCandList("candjob_interv_i") = 1 then%> CHECKED<%end if%>></span><%end if%></td>
	<%if gCandList("candjob_recomm_i") = 1 then%>
	<td class ="NonPrintable" width="<%=pv_width16%>" valign="top"  align=center bgcolor="#33ff66">Recm
<%else%>
    <td class ="NonPrintable" width="<%=pv_width17%>" valign="top"  align=center>
 <%end if
    If instr(Request.querystring("Q"),"UpdOnly") then%>
    	<br><span title="Click to set as RECOMMENDED"><input type="Checkbox" name="recomm_japid" class ="Spec7 NonPrintable" value="<% response.write gCandList("cand_id_c")%>"<%if gCandList("candjob_recomm_i") = 1 then%> CHECKED<%end if%>></span><%end if%></td>
    	
	<%if gCandList("candjob_selected_i") = 1 then
	'08 AUG 15 LJL selected letter comes from cand-list8 file include
	' get the select letter appropriate to this gradetype/level
	'08 AUG 15 LJL revise the entire select letter process - put into this page
	
		if GETJAFINFO1("gradetype_typ_t") = "1" then
			pv_selectletter = "'CAND-SELECTLETTER-G'"
		else
			pv_selectletter = "'CAND-SELECTLETTER-P'"
		end if
			
		'  <!-------------------SEND CONFIRMATION EMAIL TO USER's EMAIL ADDRESS----------------------------------->
  		gSLsql = " SELECT corrshare_thisorg_" & session("template_org_code") & "_corrid as CORRID FROM tx_rsys_corrshare 	WHERE rtrim(corrshare_dsc_t) = " & pv_selectletter & " AND corrshare_thisorg_" & session("template_org_code") & " = 1 "
  		set gSL =rsys_db_select.execute(gSLsql)
  		
  		'response.write "<br>TEST GETSLsql:<Br>" & gSLsql
  		
	  	if gSL.eof = false then
			  gl_corres = gSL("corrid")
  	  	else
  			' 03 NOV 24 LJL reiterated the stop of no selection letter being present%>
  			<font color=red size=3><strong>Sel letter not set.  System stop.</strong></font>
 			<%	
  	 	response.end
  	 	end if
	
	%>
    <td class ="NonPrintable" width="<%=pv_width17%>" valign="top"  align=center bgcolor="#db7093">Select<br>Letter<Br><a class ="NonPrintable" href='rsys-email-prep.asp?cand_id_c=<% response.write gCandList("cand_id_c")%>&corr_id_c=<% response.write gl_corres%>&jobinfo_uid_c=<% response.write newjobid%>&lng=EN' target="_blank">EN</a>
	<a class ="NonPrintable" href='rsys-email-prep.asp?cand_id_c=<% response.write gCandList("cand_id_c")%>&corr_id_c=<% response.write gl_corres%>&jobinfo_uid_c=<% response.write newjobid%>&lng=FR' target="_blank">FR</a>
<%else%>
    <td class ="NonPrintable" width="<%=pv_width18%>" valign="top"  align=center>
 <%end if
If instr(Request.querystring("Q"),"UpdOnly") then%>
	<br><span title="Click to set as SELECTED"><input type="Checkbox" name="selected_japid" class ="Spec8 NonPrintable" value="<% response.write gCandList("cand_id_c")%>"<%if gCandList("candjob_selected_i") = 1 then%> CHECKED<%end if%>></span><%end if%></td>

<%' 10 MAY 10 LJL grid section
'jobsize = 0
'if gCandList("candjob_gridtotal_c") >= 100 then
'    jobsize = 100
'    jobcolor = "green"
'elseif gCandList("candjob_gridtotal_c") <= 0 OR ISNULL(gCandList("candjob_gridtotal_c")) then
'    jobsize = 0
'    jobcolor = "silver"
'else
'    jobsize = 0
'	jobsize = gCandList("candjob_gridtotal_c")
'    jobcolor = "#9999cc"
'end if
'>
'	<td nowrap colspan="1"><font color="<% response.write jobcolor>"> <% response.write jobsize> %</td>
%>
</tr>
</table>
</td>
</tr>

<%'21 JUN 10 LJL removed </tr>%>

<%NAMER = gCandList("cand_lnam_t")&gCandList("cand_fnam_t")
	emailcheck = gCandList("email")
    japbirth = gCandList("cand_bth_d")

'27 Feb 16 GG added
if candsTo = 0 and candsCount > rowsCountConst and candcounter >= rowsCountConst Then
	gCandList.MoveLast
	gCandList.movenext
ElseIf candsTo > 0 and candsCount > rowsCountConst and candcounter >= candsTo Then
	gCandList.MoveLast
	gCandList.movenext
	'response.write candcounter& "->"& 	candsTo
Else
	gCandList.movenext
End If
	
Loop%>

<!------- END LOOP OF CAND OUTPUT FOR VN -------------------->

<!--- remove extra line break section
<tr>
	<td colspan="16"><hr size="1"></td>
</tr>
---->

<%if session("template_org_code") = 2000 and GCandList.eof = false then
    if yo_total = "0" AND yo_count = "0" then
    	yo_result = 0
    else
    	yo_result = int(yo_total / yo_count)
    end if
 else
	    if yo_total = "0" AND yo_count = "0" then
    		yo_result = 0
		else
    		yo_result = int(yo_total / yo_count)
		end if
 end if
    %>
   <!----  REMOVE EXTRA SECTION BREAK
  <TR>
    	<td valign="top"  align="left">&nbsp;</TD>
  </tr> 
  ---> 
  <TR>
    	<td colspan="16" bgcolor="#FFCCFF" align="center">

<!-------------------- BEGIN PAGES INDICATOR SECTION ----->
<!----- PAGE OUTPUT if many more than page limit which is ....  ------>		
<div>

<%'13 OCT 24 LJL try to hide pages if no pages indicated
' if candsTo > 100 then %>
Prev and Next Pages (if list overly long): &nbsp;&nbsp;
<%'end if%>


<% '27 Feb 16 GG added Pagination
If candsto>rowsCountConst then 
	If Len(candstoSubstr) Then
		strContent = Replace(scriptUrl,candstoSubstr, "candsto=" & (candsto-rowsCountConst)) 
	Else
		strContent = scriptUrl & "&candsto=" & (candsto-rowsCountConst)
	End If
	If request.form("sort") <> "" then
		strContent = strContent & "&sort=" & Replace(request.form("sort")," ", "+")
	elseif request.querystring("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.querystring("sort")," ", "+")
	end if
%>
		<a class ="NonPrintable" href= '<% response.write strContent%>' ><b>Previous </b></a>
<% 
End IF 

if candsCount > rowsCountConst Then
	pagesCount = candsCount/rowsCountConst
	pageIndex = 0	
	
	Do while pagesCount > pageIndex
		pageIndex = pageIndex + 1
		
		If Len(candstoSubstr) Then
			strContent = Replace(scriptUrl,candstoSubstr, "candsto=" & (rowsCountConst*pageIndex)) 
		Else
			strContent = scriptUrl & "&candsto=" & (rowsCountConst*pageIndex)
		End If
		If request.form("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.form("sort")," ", "+")
		elseif request.querystring("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.querystring("sort")," ", "+")
		end if
%>
		<a class ="NonPrintable" <% 	If currentPage = pageIndex Then
					response.write "style='color:#990066;font-weight: bolder;'"
				End IF
		%> href='<% response.write strContent%>' ><b><% response.write pageIndex %></b></a>
<% 
	Loop
End If 

If candsCount/rowsCountConst > 1 and candsCount > candsto then 
If Len(candstoSubstr) Then
	strContent = Replace(scriptUrl,candstoSubstr, "candsto=" & (candsto+rowsCountConst)) 
Else
	strContent = scriptUrl & "&candsto=" & (candsto+rowsCountConst+rowsCountConst) 
End If
		If request.form("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.form("sort")," ", "+")
		elseif request.querystring("sort") <> "" then
			strContent = strContent & "&sort=" & Replace(request.querystring("sort")," ", "+")
		end if
%>
	<a class ="NonPrintable" href='<%  response.write strContent %>' ><b>Next</b></a>
<% End If %>
</div>
<!-------------------- END PAGES INDICATOR SECTION ----->


		</TD>
  </tr>
'  <!---  REMOVE EXTRA BREAK SECTION--->
<tr>
	<td colspan="16" bgcolor="#FFFFCC" align="center" ><strong><em>Average age of applicants is <% response.write yo_result%> years</em></tr></strong></td>
</tr>

<%'28 JUN 12 LJL revised ITU not consider list to be for cands with Contact HR ranks for overall, not per VN
'25 JUN 12 LJL add section for ITU to identify those applicants to not consider or select, per SUEDIS
		  if session("template_org_code") = 2400 then
	dim gNOCONSIDERsql, gNOCONSIDER
	         gNOCONSIDERsql = " {call erstp_cand_notconsider_2400 (" & newjobid & ")} "
			  ' , rank_dsc_en_t, rank_official_i, rank_color_c, rank_off_dsc_t 		
    	     set gNOCONSIDER =rsys_db_select.execute(gNOCONSIDERsql)
if gNOCONSIDER.eof = false then%>

<tr>
	<td colspan="16"><font color='maroon'><strong>List of applicants to not consider (overall rating):</strong></font></td>
</tr>

<%do while gNOCONSIDER.eof = false%>
<tr>
	<td colspan="4"><%=gNOCONSIDER("cand_lnam_t") & ", " & gNOCONSIDER("cand_fnam_t")%></td>
	<td colspan="4" ><font color='maroon'><%=gNOCONSIDER("RANKDSC")%></font></td>
</tr>


<%gNOCONSIDER.movenext
loop
end if

end if

If instr(Request.querystring("Q"),"Matrixonly") then%>

<tr>
<td colspan="16">
<!--#include file = "rsys_matrix_menu.asp""--->
</td>
</tr>

<tr>
	<td colspan="16" bgcolor="#9999CC" align="center"><input type="submit" value="Generate Matrix for the applicants above"></td>
</tr>
     </tr> 
    </tbody>
</TABLE> 
       

<input type="hidden" name="fm_jobid" value="<%=newjobid%>">
<input type="hidden" name="fm_japspost" value="99">
<!--27 JUN 15 GG added candsCount to the rsys_matrix.asp form-->
<INPUT TYPE="hidden" NAME="candsCount" VALUE="<%= candcounter %>">
</form>
<%
  elseif instr(Request.querystring("Q"),"MoveToPosts")  then
  if instr(session("rightsgroup"), ",5,") then
  	singleonly = 1
  else
  	singleonly = 0
  end if
  if len(request.querystring("v_year")) then
  	v_year = trim(request.querystring("v_year"))
  else
  	v_year = year(now())
  end if
  GETYEARSsql = "	 {call erstp_rsys_list_years_" & session("template_org_code") & "}"
	rsys_db_select.CommandTimeout = 320
  set GETYEARS =rsys_db_select.execute(GETYEARSsql)

    short_year = right(year(now()),2)
'    <!--------------------- GETVACS module shared on hrd-cand-info.asp, hrd-assess-cand.asp, hrd-japsearch2.asp, hrd-cllist.asp -------------------------->
'					  	<!---- UNITADMIN ADD  REV 9 FEB 05 LJL ---------->
    GetVACSsql = " "
	'IF session("CLI_ADMIN_UNITS") = "0" then
'	<!----------- SETS PAGE TO VIEW ONLY APPLICANT AND ONE SPECIFIED POST ---------------------->
'if ((request.querystring("singlejob") = 1) OR (singleonly = 1)) AND len(request.form("jobinfo_uid_c")) then
'    GetVACSsql = GetVACSsql & " SELECT   dbo.td_rsys_jobinfo.status_id_c, dbo.tr_rsys_status.status_dsc_t, dbo.td_rsys_jobinfo.jobinfo_uid_c, dbo.td_rsys_jobinfo.jobinfo_vac2_c, dbo.tr_rsys_status.status_sht_t, dbo.td_rsys_jobinfo.jobinfo_job_en_t FROM dbo.tx_rsys_candjob INNER JOIN dbo.td_rsys_jobinfo ON dbo.tx_rsys_candjob.jobinfo_uid_c = dbo.td_rsys_jobinfo.jobinfo_uid_c LEFT OUTER JOIN dbo.tr_rsys_status ON dbo.td_rsys_jobinfo.status_id_c = dbo.tr_rsys_status.status_id_c WHERE ((datepart(yyyy, td_rsys_jobinfo.jobinfo_acl_d) = " & v_year & ") OR td_rsys_jobinfo.jobinfo_vac2_c LIKE '%/" & short_year & "/%' OR tr_rsys_status.status_across_postdates_i = 1)) AND (org_wk_c IN (" & session("unitsgroup") & ") OR td_rsys_jobinfo.jobinfo_option2_i = 1 ) AND tx_rsys_candjob.jobinfo_uid_c = " & request.form("jobinfo_uid_c")
'03 NOV 11 LJL removed template viewable, mainly for ILO, but probably should not be there for anyone other than rights
   'OR td_rsys_jobinfo.jobinfo_template_i = 1 
   
'<!------------------------ ORG IDENT --------------------->
'	GetVACSsql = GetVACSsql & " AND jobinfo_thisorg_" & session("template_org_code") & " = 1 ORDER BY tr_rsys_status.status_order_c, td_rsys_jobinfo.jobinfo_job_en_t, td_rsys_jobinfo.jobinfo_vac2_c "
'	ELSE
'GetVACSsql = GetVACSsql & " SELECT tr_rsys_status.status_dsc_t, td_rsys_jobinfo.jobinfo_uid_c, td_rsys_jobinfo.jobinfo_vac2_c, tr_rsys_status.status_sht_t,  td_rsys_jobinfo.jobinfo_job_en_t FROM tr_rsys_status RIGHT OUTER JOIN td_rsys_jobinfo ON tr_rsys_status.status_id_c = td_rsys_jobinfo.status_id_c WHERE (((datepart(yyyy, td_rsys_jobinfo.jobinfo_acl_d) = " & v_year & ") AND (org_wk_c IN (" & session("unitsgroup") & ")) OR tr_rsys_status.status_across_postdates_i = 1)) "
'03 NOV 11 LJL removed template viewable, mainly for ILO, but probably should not be there for anyone other than rights
'OR td_rsys_jobinfo.jobinfo_template_i = 1 

'<!------------------------ ORG IDENT --------------------->
'	GetVACSsql = GetVACSsql & " AND jobinfo_thisorg_" & session("template_org_code") & " = 1 ORDER BY tr_rsys_status.status_order_c, td_rsys_jobinfo.jobinfo_job_en_t, td_rsys_jobinfo.jobinfo_vac2_c "
'	end if
'ELSE

if ((request.querystring("singlejob") = 1) OR (singleonly = 1)) AND len(request.form("jobinfo_uid_c")) then
GetVACSsql = GetVACSsql & " SELECT     dbo.tr_rsys_status.status_dsc_t, dbo.td_rsys_jobinfo.jobinfo_uid_c, dbo.td_rsys_jobinfo.jobinfo_vac2_c, dbo.tr_rsys_status.status_sht_t,  dbo.td_rsys_jobinfo.jobinfo_job_en_t FROM dbo.tx_rsys_candjob INNER JOIN dbo.td_rsys_jobinfo ON dbo.tx_rsys_candjob.jobinfo_uid_c = dbo.td_rsys_jobinfo.jobinfo_uid_c LEFT OUTER JOIN dbo.tr_rsys_status ON dbo.td_rsys_jobinfo.status_id_c = dbo.tr_rsys_status.status_id_c WHERE (j.jobinfo_thisorg_" & session("template_org_code") & " = 1) AND (((datepart(yyyy, td_rsys_jobinfo.jobinfo_acl_d) = " & v_year & " OR td_rsys_jobinfo.jobinfo_vac2_c LIKE '%/" & short_year & "/%' ) OR (tr_rsys_status.status_across_postdates_i = 1)) "

'03 NOV 11 LJL removed template viewable, mainly for ILO, but probably should not be there for anyone other than rights
'td_rsys_jobinfo.jobinfo_template_i = 1) OR 


'	<!----------- SETS PAGE TO VIEW ONLY APPLICANT AND ONE SPECIFIED POST ---------------------->
	  	GetVACSsql = GetVACSsql & " AND tx_rsys_candjob.jobinfo_uid_c = " & request.form("jobinfo_uid_c")
'<!------------------------ ORG IDENT --------------------->
	GetVACSsql = GetVACSsql & " AND jobinfo_thisorg_" & session("template_org_code") & " = 1 ORDER BY tr_rsys_status.status_order_c, td_rsys_jobinfo.jobinfo_job_en_t, td_rsys_jobinfo.jobinfo_vac2_c "
else
	'GetVACSsql = GetVACSsql & " {call erstp_rsys_jobs_by_year_" & session("template_org_code") & "(" & v_year & "," & short_year & ")} "
	
	GetVACSsql = GetVACSsql & " SELECT s.status_dsc_t, j.jobinfo_uid_c, j.jobinfo_vac2_c, s.status_sht_t,  j.jobinfo_job_en_t FROM tr_rsys_status s RIGHT OUTER JOIN td_rsys_jobinfo j ON s.status_id_c = j.status_id_c "
	if   session("CLI_ADMIN_CCOG") = "0"  then
		GetVACSsql = GetVACSsql & " INNER JOIN tx_rsys_jobccog c ON j.jobinfo_uid_c = c.jobinfo_uid_c "
	end if
	
	 GetVACSsql = GetVACSsql & "  WHERE (j.jobinfo_thisorg_" & session("template_org_code") & " = 1) AND ((datepart(yyyy, j.jobinfo_acl_d) = " & v_year & ") OR s.status_across_postdates_i = 1) "
'03 NOV 11 LJL removed template viewable, mainly for ILO, but probably should not be there for anyone other than rights
'OR td_rsys_jobinfo.jobinfo_template_i = 1 	
	
if   session("CLI_ADMIN_UNITS") = "0" OR session("CLI_ADMIN_POSTS") = "0" OR instr(Session("rightsgroup"),",106,") OR session("CLI_ADMIN_CCOG") = "0"  then
				GetVACSsql = GetVACSsql & "  AND (0=1 "
end if
          if   session("CLI_ADMIN_UNITS") = "0"  then
				GetVACSsql = GetVACSsql & "  OR (j.org_wk_c IN (" & session("unitsgroup") & "))"
          end if
          if   session("CLI_ADMIN_POSTS") = "0"  then
				GetVACSsql = GetVACSsql & "  OR  (j.jobinfo_uid_c IN (" & session("postsgroup") & "))"
          end if
          if   session("CLI_ADMIN_CCOG") = "0"  then
				GetVACSsql = GetVACSsql & "  OR (c.ccog_id_c IN (" & session("ccoggroup") & "))"
          end if
if   session("CLI_ADMIN_UNITS") = "0" OR session("CLI_ADMIN_POSTS") = "0" OR instr(Session("rightsgroup"),",106,") OR session("CLI_ADMIN_CCOG") = "0"  then
				GetVACSsql = GetVACSsql & " OR (j.jobinfo_option2_i = 1) )"
				'03 NOV 11 LJL removed template viewable, mainly for ILO, but probably should not be there for anyone other than rights
'a.jobinfo_template_i = 1 OR 
end if
	
	
	
	end if
'end if
	rsys_db_select.CommandTimeout = 320
    set GetVACS =rsys_db_select.execute(GetVACSsql)
    'response.write "<tr><td>GetVAC: " & Getvacssql & "</td></tr>"

'<!----------- SETS PAGE TO VIEW ONLY APPLICANT AND ONE SPECIFIED POST ---------------------->
if ((request.querystring("singlejob") = 1) OR (singleonly = 1)) AND len(request.form("jobinfo_uid_c")) then

else
	if len(session("orgid")) then
		else%>
  <tr>
    <td colspan="16">
      <TABLE border=0 frame=hsides width="100%">
        <TR>
          	<td class="contentbold" colspan="4" valign="top">Add Applicant to a Vacancy Notice
            <%
'            	<!--------- UNITADMIN ADD --------->
            if session("CLI_ADMIN_UNITS") = "0" then%>
            	 (Limited to your unit list)
 <%end if

'            			<!--------- POSTADMIN ADD --------->
            if session("CLI_ADMIN_POSTS") = "0" then%>
            	(Limited to your posts list)
 <%end if

'            		<!--------- CCOGADMIN ADD --------->
            if session("CLI_ADMIN_CCOG") = "0" then%>
            	(Limited to your posts list)
 <%end if%>
            </b></td>
        </tr>
        <TR>
          	<td class="contentbold" valign="top" colspan="4">Vacancies, listed by status, closing in year
            <%if v_year = 1925 then%>TDB
            <%elseif v_year = 2999 then%>Until Cand Identified
			<%else
			response. write v_year
			 end if%>
<Br>
            <%
            Do while GETYEARS.eof = false%><a href='hrd-cllist.asp?Q=MoveToPosts&Q=ByJ&jobinfo_uid_c=<%=newjobid%>&v_year=<% response.write GETYEARS("yearlist")%>' target="_parent">
			<% if yearlist = 1925 then%>TBD
			<% elseif yearlist = 2999 then%>Until Cand Identified
			<%else
				response.write GETYEARS("yearlist")
         	end if%>
</a>
            <%
            GETYEARS.movenext
            loop%></td>
        </tr>
        <TR class="bglight">
          	<td valign="top">Vacancy</td>
          	<td valign="top">Date applied</td>
          	<td valign="top">&nbsp;</td>
</tr>
<tr>
	<td valign="top">
	<Select  NAME="addpost_id">
    <%Do while GETVACS.eof = false%>
	<OPTION VALUE="<% response.write getvacs("jobinfo_uid_c")%>"><% response.write GetVACS("status_sht_t")%> -
	<%IF LEN(getvacs("jobinfo_job_en_t")) > 28 then
		response.write LEFT(getvacs("jobinfo_job_en_t"), 28) & "..."
	ELSE
		response.write getvacs("jobinfo_job_en_t")
	end if
	IF len(getvacs("jobinfo_vac2_c")) then
		response.write "(" & getvacs("jobinfo_vac2_c") & ")"
	end if
    GETVACS.movenext
    loop%></SELECT>
	<input type="hidden" name="addvac" value="1"></td>
    <td valign="top">
    <input type="text" name="upd_d" value="<%=day(now()) & "-" & monthname(month(now()),1) & "-" & year(now())%>" size="14" maxlength="16">
    <br>
    Comments (required)
    <br>
    <input type="text" name="candjob_add_comments_t" value="Added en masse" size="25" maxlength="75">
    <input type="hidden" name="candjob_add_comments_t_required" value="Please enter comments on why this post was added to this applicant">
	</td>
</tr>
<TR>
	<td class="contentlg" colspan="4" valign="top">&nbsp;</td>
</tr>
<TR>
          	<td class="contentlg" colspan="4" valign="top" width="100%"><hr></td>
        </tr>
      </table>
	  <%
      end if
      end if
	%>
      </td>
  </tr>

  	<tr>
    		<td colspan="16" bgcolor="white" align="center"><input type="submit" value="Add to Posts for all checked above"></td>
  	</tr>
	</table>
</form>
<%elseif  instr(Request.querystring("Q"),"EvalApp") then%>
	<tr>
  		<td colspan="16" bgcolor="white" align="center"><input type="hidden" name="UPDEVAL" value="1">
    		<input type="submit" value="Update applicants evaluations for all checked">
    		</td>
	</tr>
	</table>
<input type="hidden" name="fm_jobid" value="<%=newjobid%>">
<input type="hidden" name="fm_japspost" value="99">
</form>
<%elseif  instr(Request.querystring("Q"),"UpdOnly") and  session("template_org_code") = 2000 then
if GetERapsInfo.eof = false then %>
	<tr>
  		<td colspan="16" bgcolor="white" align="center">
		<input type="hidden" name="UPDATEPAGE" value="1">
    	<input type="button" value="Update applicants for all checked above" onClick="newSubmit() ">
    		</td>
	</tr>
	</table>
<input type="hidden" name="fm_jobid" value="<%=newjobid%>">
<input type="hidden" name="fm_japspost" value="99">
</form>
<%else%>
	<tr>
  		<td colspan="16" bgcolor="white" align="center">
		<input type="hidden" name="UPDATEPAGE" value="1" ID="Hidden1">
    	<input type="submit" value="Update applicants for all checked above" ID="Submit2" NAME="Submit2">
    		</td>
	</tr>
	</table>
<input type="hidden" name="fm_jobid" value="<%=newjobid%>" ID="Hidden2">
<input type="hidden" name="fm_japspost" value="99" ID="Hidden3">
</form>
<% end if
elseif  instr(Request.querystring("Q"),"UpdOnly")   then%>
	<tr>
  		<td colspan="16" bgcolor="white" align="center">
		<input type="hidden" name="UPDATEPAGE" value="1" ID="Hidden1">
    	<input type="submit" value="Update applicants for all checked above" ID="Submit2" NAME="Submit2">
    		</td>
	</tr>
	</table>
<input type="hidden" name="fm_jobid" value="<%=newjobid%>" ID="Hidden2">
<input type="hidden" name="fm_japspost" value="99" ID="Hidden3">
</form>

<%elseif  instr(Request.querystring("Q"),"wordonly") then%>
	<tr>
  		<td colspan="16" bgcolor="white" align="center"><input type="submit" value="Generate Word output for all checked applicants above"></td>
	</tr>
	</table>
<input type="hidden" name="fm_jobid" value="<%=newjobid%>">
<input type="hidden" name="fm_japspost" value="99">
</form>

<%elseif  instr(Request.querystring("Q"),"HTMLonly") then%>
	<tr>
  		<td colspan="16" bgcolor="white" align="center"><input type="submit" value="Generate HTML output for all checked applicants above"></td>
	</tr>
	</table>
<input type="hidden" name="fm_jobid" value="<%=newjobid%>">
<input type="hidden" name="fm_japspost" value="99">
</form>

<%elseif  instr(Request.querystring("Q"),"Commonly") then
GetCORRSsql = "	SELECT corr_id_c, corr_dsc_en_t, corr_name_t, corr_auto_archive_i, corr_is_email_i 	FROM core_corrtblf 	WHERE corr_inactind_i <> 1 	AND corr_mass_email_i = 1	AND corr_thisorg_" & session("template_org_code") & " = 1 	ORDER BY 3"
rsys_db_select.CommandTimeout = 320
set GetCORRS =rsys_db_select.execute(GetCORRSsql)%>
<tr>
  <td colspan="16" align="center">
    <br>
    <TABLE WIDTH="100%" border="0" frame=hsides bordercolor="navy" align="center">
      <tr>
        	<td class="black9" valign="top" colspan="2"><b>Send Communication or enter status notes</td>
      </tr>
      <tr>
        	<td class="black9" valign="top">Archive?</td>
        	<td class="black9" valign="top"><b><Select  NAME="archive">
          		<OPTION value="1">ARCHIVE THE FOLLOWING
          		<OPTION value="0" SELECTED>REMAIN ACTIVE
          	</select></b></td>
      </tr>
      <tr>
        	<td><input type="hidden" name="archive" value="1"></td>
      </tr>
      <TR>
        	<td class="black9" valign="top">Communication type</td>
        	<td class="black9" valign="top"><Select  NAME="corr_id_c">
          	    <OPTION VALUE="">- Choose one
          <%Do while GETCORRS.eof = false%>
          <OPTION VALUE="<% response.write GetCORRS("corr_id_c")%>">
          <%if GetCORRS("corr_is_email_i") = 0 then%>Status<%else%>E-mail<%end if%>
		  -
		  <% response.write GetCORRS("corr_name_t")
          GETCORRS.movenext
          loop%></SELECT>
        </tr>
<%' <!-------------------- SEND OVER THE JOB ID------------------------------------>%>
<INPUT TYPE="HIDDEN" NAME="jobinfo_uid_c" VALUE="<% response.write newjobid%>">
<%'  <!------------------------------- ADD ORG INFO 1000 2000 4000 - 3000 only in EN 12 MAR 05 LJL per WTO req 38 --------------------------------->
if session("template_org_code") <> 3000 then%>
        <tr>
          	<td class="black9" valign="top">Language</td>
          	<td class="black9" valign="top"><Select  NAME="lng">
            <OPTION VALUE='en' selected>English
            <OPTION VALUE='fr'>French
            <OPTION VALUE='es'>Spanish
            	</SELECT>
            <input type="hidden" name="goforit" value="1">
            </td>
        </tr>
<%else%>
        <INPUT TYPE="HIDDEN" NAME="lng" VALUE="en">
<%end if%>
      </table>
      </td>
  </tr>
  	<tr>
    		<td colspan="16" bgcolor="99CC66" align="center"><input type="submit" value="Generate Communication for all checked above"></td>
  	</tr>
	</table>
<input type="hidden" name="fm_jobid" value="<%=newjobid%>">
<input type="hidden" name="fm_japspost" value="99">
</form>
<%else
'21 JUN 10 LJL added excel
if request.querystring("goExcel") = 99 then

else%>
  <TR>
    	<td valign="top"  align="left">&nbsp;</TD>
  </tr>
  <!-- 14 OCT 24 LJL changed bg color of PDF generation block -->
<tr>
	<td colspan="16" bgcolor="#FF9999" align="center"><input type="submit" value="Generate PDF file for all checked applicants"></td>
</tr>
<%end if%>
</table>
<input type="hidden" name="fm_jobid" value="<%=newjobid%>">
<input type="hidden" name="fm_japspost" value="99">
</form>
<%end if
'21 JUN 10 LJL added excel
if request.querystring("goExcel") = 99 then

else%>
<TABLE WIDTH="100%" align="center" border="0" bordercolor="green">
   <TR>
    	<td valign="top"  align="left"><Br>NB: Applicant names highlighted in PINK indicate a possible duplicate application (If so, applicant may need to be deleted)<br>
      	Applicant names highlighted in YELLOW indicate a staff list duplicate where staff has email in data sources.
      	<br>
      	</TD>
  </tr>
</table>

<%'21 JUN 10 LJL for goExcel portion
end if
' IF CANDLIST.recordcount end if
end if
'ADDED
pv_last_update = "27 Oct 24"
rsys_logs.close
Set rsys_logs = nothing

rsys_db_select.close
Set rsys_db_select = nothing

rsys_db.close
Set rsys_db = nothing

rsys_int.close
Set rsys_int = nothing

rsys_int_select.close
Set rsys_int_select = nothing


'21 JUN 10 LJL added excel
if request.querystring("goExcel") = 99 then

else%>
<!--#include file="includes/include_admin_frame_bottom.asp"-->
<%end if%>