<% option explicit%>
<!--#include file = "includes/include_check_login.asp"-->
<!--#include file = "../includes/rsys_db_select_dim.asp"-->
<%

dim GETMALE, GETFEMALE, GETWHOMALE, GETWHOFEMALE, GetJAFsql, GetJAF, GetGENDERsql, GetGENDER
dim getmalesORGsql, getmalesORG, getfmalesORGsql, getfmalesORG, getmalessql, getmales, getfmalessql, getfmales
dim maler, fmaler, counter, total_maler, total_fmaler, pv_totalgender
'dim femalecnt, malecnt
dim getVacancySql, getVacancy, pv_jobid, pv_regionsok


'femalecnt = 0
'malecnt = 0

pv_page_title = "REGIONAL / GENDER DISTRIBUTION OF APPLICATIONS"

'<!-------------------------------------------NOTES ----------
'MODS --
'12 NOV 03 LJL ORG enabled - may need revision, totals not correct
'23 AUG 05 RRr reviewed asp
'21 OCT 08 DD  revised query for male count
'04 APR 10 DD Modified to get correct count of male and female
'28 APR 10 DD Added graph to this statistics
'03 MAY 10 DD Assigned value to statistical variables (to maler and fmaler)
'27 JAN 11 DD Y-axis name is changed
'17 FEB 11 DD Statistics of Gender only for applicants applied to post, it is for WTO only

' 30 AUG 05 LJL ready for testing
'15 JUN 06 LJL lock out non Full Admin users (meaning those who are NATSOC will not be able to get to certain parts or whole page
'26 JUN 08 LJL revised for GSM
'25 MAY 15 LJL no excel graph output
'03 JUN 15 LJL changed terminology for WMO applications
'03 JUN 15 LJL worked on not outputting graph stuff into excel
'04 AUG 15 LJL updated regions for all - Manoj - check ILO output which is not showing gender
'18 FEB 16 LJL added Vn number and title - output of corrected regions


' NEEDNEEDNEED - VERIFY TOTALS

if  Request.querystring("jobinfo_uid_c") <> "" then
	pv_jobid = Request.querystring("jobinfo_uid_c")
elseif Request.form("jobinfo_uid_c") <> "" then
	pv_jobid = Request.form("jobinfo_uid_c")
else
	pv_jobid = ""
end if

IF request.form("goexcel") <> "" then

else%>
<!--#include file = "includes/include_admin_frame_top.asp"-->
<%end if%>

<%'03 JUN 06 LJL added stop for not able to view posts right
if   instr(Session("rightsgroup"),",106,") then
' 03 JUN 06 LJL lock out those without VN view right 106
	response.write "Rights not provided for further Applicant access<br><br><br>"
else

'response.write "<br>JOBID: " & pv_jobid


' IF JOB ID INDICATED, show the vac header
if pv_jobid <> "" then

IF request.form("goexcel") <> "" then

else%>
<!--#include file = "rsys_vac_admin_menu.asp"-->
<br>
<%end if%>
<h3>Regional / Gender Applications</h3>
<br>
<%'GetJAFsql = "SELECT jobinfo_job_en_t, jobinfo_vac2_c FROM td_rsys_jobinfo WHERE jobinfo_uid_c = " & pv_jobid & " AND jobinfo_thisorg_" & Session("template_org_code")& " = 1 	"
    'set GetJAF =rsys_db_select.execute(GetJAFsql)
'if pv_jobid <> "" then//>
	'For vacancy: <% response.write "<strong>" & GetJAF("jobinfo_job_en_t") & "</strong> No. <strong>" & GetJAF("jobinfo_vac2_c") & "</strong>"
'end if
end if%>

<HEAD>
	
<%IF request.form("goexcel") <> "" then

else%>
	
<SCRIPT LANGUAGE="Javascript" SRC="exportchart/FusionCharts.js"></SCRIPT>
<script language="JavaScript" src="exportchart/FusionChartsExportComponent.js"></script>
  <script type="text/javascript">
      //Define a function, which will be invoked when user clicks the batch-export-initiate button
      function initiateExport() {
          myExportComponent.BeginExport();

      }
  </script>
	<style type="text/css">
	<!--
	body {
		font-family: Arial, Helvetica, sans-serif;
		font-size: 12px;
	}
	-->
	</style>
	
<%end if%>
	
</HEAD>


<%IF request.form("goexcel") <> "" then
 
Response.ContentType = "application/vnd.ms-excel" 
' Adds a header to give the document a name
Response.AddHeader "content-disposition", "inline; filename=ApplicationsGenderList-" & now() & ".xls"

else%>

<br>
<center>
<form name="excelForm" action="hrd-clgenderfull.asp" method="post">
<%	'<!--- Hidden fields to transmit info for refining query --->%>
	<!--#include file = "rsys_follow_vars.asp"-->
<INPUT type="hidden" name="goexcel" value="1">
<INPUT type="hidden" name="jobinfo_uid_c" value="<%=pv_jobid%>">
<input type="submit" value=" Export to Excel ">
</form>
</center>
<br>

<%end if%>

<%GetJAFsql = "SELECT jobinfo_job_en_t, jobinfo_vac2_c FROM td_rsys_jobinfo WHERE jobinfo_uid_c = " & pv_jobid & " AND jobinfo_thisorg_" & Session("template_org_code")& " = 1 	"
    set GetJAF =rsys_db_select.execute(GetJAFsql)
if   pv_jobid <> "" then%>
<Br>
<center>
	Vacancy Notice Number: <% response.write "<strong>" & GetJAF("jobinfo_vac2_c") & "</strong><br>"%>
	<% response.write "<b>" & GetJAF("jobinfo_job_en_t") & "</b><br><br></center>"
end if
%>


<div align="center">
<TABLE WIDTH="60%">

<%if  pv_jobid <> "" then

    if session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500 then
	    
        getmalessql = "SELECT     COUNT(c.cand_gnd_i) AS malecount FROM         dbo.td_rsys_cand c INNER JOIN                      dbo.tx_rsys_candjob cj ON c.cand_id_c = cj.cand_id_c LEFT OUTER JOIN                      dbo.v_staff_list_" & session("template_org_code") & " s ON c.staff_nbr_" & session("template_org_code") & " = s.StaffNo COLLATE SQL_Latin1_General_CP850_CI_AI WHERE     (cj.jobinfo_uid_c = " & pv_jobid & ") AND (c.cand_gnd_i = 1) AND (s.StaffNo IS NULL)"
        set getmales =rsys_db_select.execute(getmalessql)
        getmalesORGsql = "SELECT     COUNT(s.sex_code) AS malecount FROM         dbo.td_rsys_cand c INNER JOIN                      dbo.tx_rsys_candjob cj ON c.cand_id_c = cj.cand_id_c LEFT OUTER JOIN                      dbo.v_staff_list_" & session("template_org_code") & " s ON c.staff_nbr_" & session("template_org_code") & " = s.StaffNo COLLATE SQL_Latin1_General_CP850_CI_AI WHERE     (cj.jobinfo_uid_c = " & pv_jobid & ") AND (rtrim(UPPER(s.sex_code)) = 'M') AND NOT (s.StaffNo IS NULL)"
        set getmalesORG =rsys_db_select.execute(getmalesORGsql)


        getfmalessql = "SELECT     COUNT(c.cand_gnd_i) AS fmalecount FROM         dbo.td_rsys_cand c INNER JOIN                      dbo.tx_rsys_candjob cj ON c.cand_id_c = cj.cand_id_c LEFT OUTER JOIN                      dbo.v_staff_list_" & session("template_org_code") & " s ON c.staff_nbr_" & session("template_org_code") & " = s.StaffNo COLLATE SQL_Latin1_General_CP850_CI_AI WHERE     (cj.jobinfo_uid_c = " & pv_jobid & ") AND (c.cand_gnd_i = 0) AND (s.StaffNo IS NULL)"
        set getfmales =rsys_db_select.execute(getfmalessql)
        getfmalesORGsql = "SELECT     COUNT(s.sex_code) AS fmalecount FROM         dbo.td_rsys_cand c INNER JOIN                      dbo.tx_rsys_candjob cj ON c.cand_id_c = cj.cand_id_c LEFT OUTER JOIN                      dbo.v_staff_list_" & session("template_org_code") & " s ON c.staff_nbr_" & session("template_org_code") & " = s.StaffNo COLLATE SQL_Latin1_General_CP850_CI_AI WHERE     (cj.jobinfo_uid_c = " & pv_jobid & ") AND (rtrim(UPPER(s.sex_code)) = 'F') AND NOT (s.StaffNo IS NULL)"
        set getfmalesORG =rsys_db_select.execute(getfmalesORGsql)


        maler = 0
        If getmales.eof = false then
	        maler = maler + getmales("malecount")
        end if
        If getmalesORG.eof = false then
	        maler = maler + getmalesORG("malecount")
        end if

        fmaler = 0
        If getfmales.eof = false then
	        fmaler = fmaler + getfmales("fmalecount")
        end if
        If getfmalesORG.eof = false then
	        fmaler = fmaler + getfmalesORG("fmalecount")
        end if

        'response.write getmales("malecount") & "/" & getmalesORG("malecount") & "?" & getfmales("fmalecount") & ")" & getfmalesORG("fmalecount")
        %>
                        <tr>
                      	    <TD valign='top'  width="25%"><u><strong>Gender</strong></u></td>
                      	    <TD valign='top'><u><strong>Number of Applications</strong></u></td>
                    	</tr>
                   	<tr>
                    	<td>&nbsp;</td>
                    	</tr>
                        <tr>
                      	    <TD valign='top'  width="25%">Males</td>
                      	    <TD valign='top'><%=maler%></td>
                    	</tr>
                        <tr>
                      	    <TD valign='top'  width="25%">Females</td>
                      	    <TD valign='top'><%=fmaler%></td>
                    	</tr>
                    	<tr>
                    	<td>&nbsp;</td>
                    	</tr>
                        <tr>
                      	    <TD valign='top'  width="25%"><strong>Total: </strong></td>
                      	    <TD valign='top'><%=(maler+fmaler)%></td>
                    	</tr>
                        <%


    else



dim gREGIONsql, gREGION, pv_regtotal

gREGIONsql = "SELECT  r.geoloc_areaID_" & session("template_org_code") & " as AreaID, r.geoloc_en_" & session("template_org_code") & " AS regname, count(cj.cand_id_c) AS regcount FROM tx_rsys_candjob cj, td_rsys_jobinfo j, td_Rsys_cand c, tr_rsys_country cty, core_geolocf r WHERE (cj.cand_id_c = c.cand_id_c  AND cj.jobinfo_uid_c = j.jobinfo_uid_c AND cty.cty_id_c = c.cand_nat_c  AND cty.geoloc_id_c_" & session("template_org_code") & " = r.geoloc_id_c) AND j.jobinfo_thisorg_" & session("template_org_code") & " = 1 and c.cand_thisorg_" & session("template_org_code") & " = 1 AND j.jobinfo_uid_c = " & pv_jobid & " GROUP BY r.geoloc_areaID_" & session("template_org_code") & ", r.geoloc_en_" & session("template_org_code") & " "
	set gREGION =rsys_db_select.execute(gREGIONsql)
	
	pv_regtotal = 0

'**********************************************
' IF NO REGION VALUES, MOVE ON
'**********************************************
	
if gREGION.eof = true then
		'response.write "<br><Br>There are no values currently for this vacancy.<br><Br>"
		pv_regionsok = 0 
		
		
else
	pv_regionsok = 1

if pv_regionsok = 1 then
		
	Do while gREGION.eof = false
	
	pv_regtotal = pv_regtotal + gREGION("regcount")
	
    gREGION.movenext
    loop
	%>
	
	                   	<tr>
                    	<td>&nbsp;</td>
                    	</tr>
                   	<tr>
                    	<td><strong>Applications by Region</strong></td>
                    	</tr>
                  	<tr>
                    	<td>&nbsp;</td>
                    	</tr>
                    <tr>
                       	    <TD valign='top' width="50%"><strong><i>Regions</i></strong></td>
                       	    <TD valign='top' align="right" width="25%"><strong><i>Number of applications</i></strong></td>
                      	    <TD valign='top' align="right" width="25%"><strong><i>%</i></strong></td>
                    </tr>

	
	<%	' 05 MAR 15 REGIONAL OUTPUT FOR WMO

                  '30 may 2015 adding region graph paramter start
                        dim regionarr()
                        dim regionnumber()
                        dim m
                       
                        m=1

						gREGION.movefirst
                        Do while gREGION.eof = false
                        
						ReDim PRESERVE regionarr(m)
                        ReDim PRESERVE regionnumber(m)
                        %>
                    <tr>
                       	    <TD valign='top'><%=gREGION("regname")%></td>
                       	    <TD valign='top' align="right"><%=gREGION("regcount")%></td>
                      	    <TD valign='top' align="right"><%=FormatNumber((gREGION("regcount")/pv_regtotal)*100, 2)%> %</td>
                    </tr>
                        <%

                            regionarr(m)=gREGION("AreaID")
                            regionnumber(m)=gREGION("regcount")
                        m=m+1
                        gREGION.movenext
                        loop%>

                        
                        <% dim color(6)
                         color(1) ="AFD8F8"
                         color(2)="FF8000"
                         color(3)="Acb99"
                         color(4)="ba0000"
                         color(5)="bad769"
                         color(6)="8a8a8a"
                          %>

                     <tr>
                       	    <TD valign='top'>&nbsp;</td>
					  </tr>
                     <tr>
                       	    <TD valign='top'><strong>Total: </strong></td>
                       	    <TD valign='top' align="right"><strong><%=pv_regtotal%><strong></td>
                      	    <TD valign='top' align="right"><strong>100.00 %<strong></td>
                    </tr>
                       
                       <% 
'**********************************************
' IF NO REGION VALUES, MOVE ON
'**********************************************
end if   
end if                     


        getmalessql = "SELECT count(cand_gnd_i) AS malecount 	FROM td_rsys_cand c, tx_rsys_candjob cj 	WHERE  (cj.jobinfo_uid_c = " & pv_jobid & ") AND c.cand_id_c = cj.cand_id_c 	AND c.cand_gnd_i = 1"
        set getmales =rsys_db_select.execute(getmalessql)

        getfmalessql = "SELECT count(cand_gnd_i) AS fmalecount 	FROM td_rsys_cand c, tx_rsys_candjob cj 	WHERE (cj.jobinfo_uid_c = " & pv_jobid & ") AND c.cand_id_c = cj.cand_id_c 	AND c.cand_gnd_i = 0"
        set getfmales =rsys_db_select.execute(getfmalessql)
        
 							maler = GetMales("malecount")
       						fmaler = GetFmales("fmalecount")
       						pv_totalgender = maler+fmaler

		'03 MAY 10 DD Assigned value to statistical variables (to maler and fmaler)
                        'Do while GetMales.eof = false
							'maler = GetMales("malecount")
                        %>

                        </TABLE>

<% ' 30 may change for graph 
	IF request.form("goexcel") <> "" then

else

if pv_regionsok = 1 then	
'**********************************************
' IF NO REGION VALUES, MOVE ON
'**********************************************
%>

                        <CENTER>
<strong><CENTER>REGIONAL DISTRIBUTION OF APPLICATIONS</CENTER></strong>
<%
	'In this example, we plot a Combination chart from data contained
	'in an array. The array will have three columns - first one for Quarter Name
	'second one for sales figure and third one for quantity.

	
	'Now, we need to convert this data into combination XML.
	'We convert using string concatenation.
	'strXML - Stores the entire XML
	'strCategories - Stores XML for the <categories> and child <category> elements
	'strDataRev - Stores XML for current year's sales
	'strDataQty - Stores XML for previous year's sales
	Dim regionXML, regionCategories, regionDataRev, regionDataQty, rcount

	'Initialize <graph> element


	regionXML = "<graph caption='Region Statistics' PYAxisName='Count of applications' SYAxisName='' numberPrefix='' formatNumberScale='0' showValues='0' decimalPrecision='0' anchorSides='10' anchorRadius='3' anchorBorderColor='FF8000' exportEnabled='1' exportAtClient='1' exportHandler='fcBatchExporter'>"
     'regionXML=  "<graph caption='Region Statistics' xAxisName='' yAxisName='Count of applications' showNames='1' decimalPrecision='0' formatNumberScale='0'>"
	'Initialize <categories> element - necessary to generate a multi-series chart
	regionCategories = "<categories>"

	'Initiate <dataset> elements
	regionDataRev = "<dataset showValues='1' seriesName='Region' color='008000' >"
	regionDataQty = "<dataset   seriesName='Count of applications' parentYAxis='S' color='FF8000' >"


	'Iterate through the data
	For rcount=1 to UBound(regionarr)
		'Append <category name='...' /> to regionCategories
		regionCategories = regionCategories & "<category name='" & regionarr(rcount) & "' />"
		'Add <set value='...' color='...'/> to both the datasets

		regionDataRev = regionDataRev & "<set value='" & regionnumber(rcount) & "' />"
		regionDataQty = regionDataQty & "<set value='" & regionnumber(rcount) & "' />"
	Next

	'Close <categories> element
	regionCategories = regionCategories & "</categories>"

	'Close <dataset> elements
	regionDataRev = regionDataRev & "</dataset>"
	regionDataQty = regionDataQty & "</dataset>"

	'Assemble the entire XML now
	regionXML = regionXML & regionCategories & regionDataRev & regionDataQty & "</graph>"

	'Create the chart - MS Column 3D Line Combination Chart with data contained in regionXML
	Call renderChart("exportchart/FusionCharts/MSColumn3DLineDY.swf", "", regionXML, "Region_Statistics", 600, 300)
end if

'**********************************************
' IF NO REGION VALUES, MOVE ON
'**********************************************
end if%>

        
<BR><BR>



<TABLE WIDTH="60%">
                        <%' 30 may change end here for graph %>

                   	<tr>
                    	<td colspan="3" width="100%"><hr></td>
                    	</tr>
                   	<tr>
                    	<td>&nbsp;</td>
                    	</tr>
                   	<tr>
                    	<td><strong>Applications by Gender</strong></td>
                    	</tr>
                   	<tr>
                    	<td>&nbsp;</td>
                    	</tr>
                   <tr>
                       	    <TD valign='top'><strong><i>Gender</i></strong></td>
                       	    <TD valign='top' align=right><strong><i>Number of applications</i></strong></td>
                      	    <TD valign='top' align=right><strong><i>%</i></strong></td>
                    </tr>
                        <tr>
                      	    <TD valign='top'  width="50%">Males</td>
                      	    <!-- <TD valign='top'><%=GetMales("malecount")%></td> -->
                      	    <TD valign='top'  width="25%" align=right><%=maler%></td>
                      	    <TD valign='top' align=right width="26%"><%=FormatNumber(((maler/pv_totalgender)*100),2)%> %</td>
                    	    </tr>
                        <%
                        'GetMales.movenext
                        'loop
                        'Do while GetFMales.eof = false
							'fmaler = GetFmales("fmalecount")
                        %>
                        <tr>
                      	    <TD valign='top'  width="50%">Females</td>
                      	    <!-- <TD valign='top'><%=GetFmales("fmalecount")%></td> -->
                      	    <TD valign='top' align=right><%=fmaler%></td>
                      	    <TD valign='top' align=right><%=FormatNumber(((fmaler/pv_totalgender)*100), 2)%> %</td>
                        </tr>
                        <%
                        'GetFMales.movenext
                        'loop%>
                        <tr>
                      	    <TD valign='top' colspan="3" width="100%">&nbsp;</td>
                      	  </tr>
                        <tr>
                      	    <TD valign='top'  width="50%"><strong>Total: </strong></td>
                      	    <TD valign='top' align=right><strong><%=(pv_totalgender)%></strong></td>
                      	    <TD valign='top' align=right><strong>100.00 %</strong></td>
                    	</tr>

    <%end if

else

    if session("template_org_code") = 1000 OR session("template_org_code") = 1200 OR session("template_org_code") = 1500 then
        getmalessql = "SELECT     COUNT(c.cand_gnd_i) AS malecount FROM         dbo.td_rsys_cand c LEFT OUTER JOIN                      dbo.v_staff_list_" & session("template_org_code") & " s ON c.staff_nbr_" & session("template_org_code") & " = s.StaffNo COLLATE SQL_Latin1_General_CP850_CI_AI WHERE cand_thisorg_" & session("template_org_code") & " = 1 AND (c.cand_gnd_i = 1) AND (s.StaffNo IS NULL)"
        'response.write GETMALESSQL
        set getmales =rsys_db_select.execute(getmalessql)
        getmalesORGsql = "SELECT     COUNT(s.sex_code) AS malecount FROM         dbo.td_rsys_cand c LEFT OUTER JOIN                      dbo.v_staff_list_" & session("template_org_code") & " s ON c.staff_nbr_" & session("template_org_code") & " = s.StaffNo COLLATE SQL_Latin1_General_CP850_CI_AI WHERE  cand_thisorg_" & session("template_org_code") & " = 1 AND (rtrim(UPPER(s.sex_code)) = 'M') AND NOT (s.StaffNo IS NULL)"
        set getmalesORG =rsys_db_select.execute(getmalesORGsql)


        getfmalessql = "SELECT     COUNT(c.cand_gnd_i) AS fmalecount FROM         dbo.td_rsys_cand c LEFT OUTER JOIN                      dbo.v_staff_list_" & session("template_org_code") & " s ON c.staff_nbr_" & session("template_org_code") & " = s.StaffNo COLLATE SQL_Latin1_General_CP850_CI_AI WHERE cand_thisorg_" & session("template_org_code") & " = 1 AND (c.cand_gnd_i = 0) AND (s.StaffNo IS NULL)"
        set getfmales =rsys_db_select.execute(getfmalessql)
        getfmalesORGsql = "SELECT     COUNT(s.sex_code) AS fmalecount FROM         dbo.td_rsys_cand c LEFT OUTER JOIN                      dbo.v_staff_list_" & session("template_org_code") & " s ON c.staff_nbr_" & session("template_org_code") & " = s.StaffNo COLLATE SQL_Latin1_General_CP850_CI_AI WHERE  cand_thisorg_" & session("template_org_code") & " = 1 AND (rtrim(UPPER(s.sex_code)) = 'F') AND NOT (s.StaffNo IS NULL)"
        set getfmalesORG =rsys_db_select.execute(getfmalesORGsql)


        maler = 0
        If getmales.eof = false then
	        maler = maler + getmales("malecount")
        end if
        If getmalesORG.eof = false then
	        maler = maler + getmalesORG("malecount")
        end if

        fmaler = 0
        If getfmales.eof = false then
	        fmaler = fmaler + getfmales("fmalecount")
        end if
        If getfmalesORG.eof = false then
	        fmaler = fmaler + getfmalesORG("fmalecount")
        end if

        'response.write getmales("malecount") & "/" & getmalesORG("malecount") & "?" & getfmales("fmalecount") & ")" & getfmalesORG("fmalecount")
        %>
                    <tr>
                      	<TD valign='top'  width="50%"><u><strong>Gender</strong></u></td>
                      	<TD valign='top'><u><strong>Number of Applications</strong></u></td>
                    </tr>
                   	<tr>
                    	<td>&nbsp;</td>
                    	</tr>

                    <tr>
                      	<TD valign='top' width="50%">Males</td>
                      	<TD valign='top'><%=maler%></td>
                    </tr>
                    <tr>
                      	<TD valign='top' width="50%">Females</td>
                      	<TD valign='top'><%=fmaler%></td>
                    </tr>
                   	<tr>
                    	<td>&nbsp;</td>
                    	</tr>
                    <tr>
                      	<TD valign='top'  width="50%"><strong>Total: </strong></td>
                      	<TD valign='top'><%=(maler+fmaler)%> %</td>
                    </tr>
                    <%


    else

   if session("template_org_code") <> 3000 then

			//Modified on 10/21/2008, changes done to following query..
			//getmalessql = "SELECT count(cand_gnd_i) AS malecount 	FROM td_rsys_cand c WHERE c.cand_id_c = cj.cand_id_c AND c.cand_gnd_i = 1"

			'04 APR 10 DD Modified to get correct count of male and female
			'getmalessql = "SELECT count(cand_gnd_i) AS malecount 	FROM td_rsys_cand c WHERE c.cand_gnd_i = 1"
			getmalessql = "SELECT count(cand_gnd_i) AS malecount 	FROM td_rsys_cand c WHERE  c.cand_thisorg_" & session("template_org_code") & "=1 and c.cand_gnd_i = 1"
			set getmales =rsys_db_select.execute(getmalessql)

			'getfmalessql = "SELECT count(cand_gnd_i) AS fmalecount 	FROM td_rsys_cand c WHERE c.cand_gnd_i = 0"
			getfmalessql = "SELECT count(cand_gnd_i) AS fmalecount 	FROM td_rsys_cand c WHERE c.cand_thisorg_" & session("template_org_code") & "=1 and  c.cand_gnd_i = 0"
			set getfmales =rsys_db_select.execute(getfmalessql)

					Do while GetMales.eof = false
						maler = GetMales("malecount")
%>
						<tr>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<TD valign='top'  width="25%">Males</td>
							<!-- <TD valign='top'><%=GetMales("malecount")%></td> -->
							<TD valign='top'><%=maler%></td>
						</tr>

						<%
						GetMales.movenext
					loop

					Do while GetFMales.eof = false
						fmaler = GetFmales("fmalecount")
					%>
						<tr>
							<TD valign='top'  width="25%">Females</td>
							<!-- <TD valign='top'><%=GetFmales("fmalecount")%></td> -->
							<TD valign='top'><%=fmaler%></td>
						</tr>
					<%
						GetFMales.movenext
					loop
					%>

                   	<tr>
                    	<td>&nbsp;</td>
                    </tr>
					<tr>
						<TD valign='top'  width="25%"><strong>Total</strong></td>
						<TD valign='top'><strong><%=(maler + fmaler)%></strong></td>
					</tr>

<%	else
			getVacancySql = "SELECT DISTINCT a.jobinfo_uid_c as vac_id, a.jobinfo_vac2_c as vac_number, a.jobinfo_job_en_t as vac_title,a.jobinfo_acl_d FROM v_rsys_jobinfo_search_" & session("template_org_code") & " a  WHERE a.jobinfo_thisorg_" & session("template_org_code") & " = 1 ORDER BY jobinfo_job_en_t ASC, jobinfo_vac2_c ASC "
			set getVacancy =rsys_db_select.execute(getVacancySql)
%>
			<tr bgcolor="#8EB4E6"><th align="left">Sr. No.</th><th align="left">Vacancy Title</th><th align="left">Vacancy Number</th><th align="left">Count of Male Applications</th><th align="left">Count of Female Applicant</th></tr>

			<%
			counter = 0
			maler = 0
			fmaler = 0
			total_maler = 0
			total_fmaler = 0

			do while getVacancy.eof = false
				counter = counter + 1

				getmalessql = " SELECT count(cj.cand_id_c) AS malecount FROM tx_rsys_candjob cj INNER JOIN td_rsys_cand c ON cj.cand_id_c = c.cand_id_c WHERE c.cand_thisorg_" & session("template_org_code") & " = 1 AND c.cand_gnd_i = 1 and cj.jobinfo_uid_c = " & getVacancy("vac_id") & " "
				set getmales =rsys_db_select.execute(getmalessql)

				getfmalessql = " SELECT count(cj.cand_id_c) AS fmalecount FROM tx_rsys_candjob cj INNER JOIN td_rsys_cand c ON cj.cand_id_c = c.cand_id_c WHERE c.cand_thisorg_" & session("template_org_code") & " = 1 AND c.cand_gnd_i = 0 and cj.jobinfo_uid_c = " & getVacancy("vac_id") & " "
				
response.write "<br>GETFMALES1:<BR>" & getfmalessql

				
				set getfmales =rsys_db_select.execute(getfmalessql)

				Do while getmales.eof = false
						maler = getmales("malecount")
					getmales.movenext
				loop

				Do while getfmales.eof = false
						fmaler = getfmales("fmalecount")
					getfmales.movenext
				loop
				%>
				<tr>
					<td><%=counter%></td>
					<td><%=getVacancy("vac_title")%></td>
					<td><%=getVacancy("vac_number")%></td>
					<td><%=maler%></td>
					<td><%=fmaler%></td>
				</tr>
			<%
				total_maler = total_maler + maler
				total_fmaler = total_fmaler + fmaler
				getVacancy.movenext
			loop
			maler = total_maler
			fmaler = total_fmaler
			%>
			<tr bgcolor="#fafad2">
				<td colspan="3" align="left">Total</td>
				<td><%=maler%></td>
				<td><%=fmaler%></td>
			</tr>

<%
	end if

        end if

end if

%>
</table>
</div>
<br><br>


<%
IF request.form("goexcel") <> "" then

else


'We've included includes/FusionCharts.asp, which contains functions
'to help us easily embed the charts.
%>
<!-- #INCLUDE FILE="Includes/FusionCharts.asp" -->

<CENTER>
<strong><CENTER>GENDER DISTRIBUTION OF APPLICATIONS</CENTER></strong>
<%
	'In this example, we plot a Combination chart from data contained
	'in an array. The array will have three columns - first one for Quarter Name
	'second one for sales figure and third one for quantity.

	Dim arrData(2,3)
	''Store Quarter Name
	arrData(0,1) = "MALE"
	arrData(1,1) = "FEMALE"

	''Store revenue data
	arrData(0,2) = maler
	arrData(1,2) = fmaler

	''Store Quantity
	arrData(0,3) = maler
	arrData(1,3) = fmaler

	'Now, we need to convert this data into combination XML.
	'We convert using string concatenation.
	'strXML - Stores the entire XML
	'strCategories - Stores XML for the <categories> and child <category> elements
	'strDataRev - Stores XML for current year's sales
	'strDataQty - Stores XML for previous year's sales
	Dim strXML, strCategories, strDataRev, strDataQty, i

	'Initialize <graph> element
	strXML = "<graph caption='Male-Female Statistics' PYAxisName='Count of applications' SYAxisName='' numberPrefix='' formatNumberScale='0' showValues='0' decimalPrecision='0' anchorSides='10' anchorRadius='3' anchorBorderColor='FF8000'  exportEnabled='1' exportAtClient='1' exportHandler='fcBatchExporter'>"

	'Initialize <categories> element - necessary to generate a multi-series chart
	strCategories = "<categories>"

	'Initiate <dataset> elements
	strDataRev = "<dataset seriesName='MALE/FEMALE' color='AFD8F8' >"
	strDataQty = "<dataset seriesName='Count of applications' parentYAxis='S' color='FF8000' >"


	'Iterate through the data
	For i=0 to UBound(arrData)-1
		'Append <category name='...' /> to strCategories
		strCategories = strCategories & "<category name='" & arrData(i,1) & "' />"
		'Add <set value='...' color='...'/> to both the datasets

		strDataRev = strDataRev & "<set value='" & arrData(i,2) & "' />"
		strDataQty = strDataQty & "<set value='" & arrData(i,3) & "' />"
	Next

	'Close <categories> element
	strCategories = strCategories & "</categories>"

	'Close <dataset> elements
	strDataRev = strDataRev & "</dataset>"
	strDataQty = strDataQty & "</dataset>"

	'Assemble the entire XML now
	strXML = strXML & strCategories & strDataRev & strDataQty & "</graph>"

	'Create the chart - MS Column 3D Line Combination Chart with data contained in strXML
	Call renderChart("exportchart/FusionCharts/MSColumn3DLineDY.swf", "", strXML, "Male_Female_Statistics", 600, 300)

end if

IF request.form("goexcel") <> "" then

else
%>
<input type='button' onClick="javascript:initiateExport();" value="Export-(PNG/JPG/PDF)" />
        <div id="fcexpDiv" align="center">FusionCharts Export Handler Component</div>
<BR><BR>
</CENTER>
<script type="text/javascript">
    //Initialize Batch Exporter with DOM Id as fcBatchExporter
    var myExportComponent = new FusionChartsExportObject("fcBatchExporter", "exportchart/FusionCharts/FCExporter.swf");
    myExportComponent.debugMode = true;
    //Add the charts to queue. The charts are referred to by their DOM Id.
    myExportComponent.sourceCharts = ['Region_Statistics', 'Male_Female_Statistics'];
    myExportComponent.componentAttributes.defaultExportFileName = 'clgenderfull';
    //------ Export Component Attributes ------//
    //Set the mode as full mode
    myExportComponent.componentAttributes.fullMode = '1';
    //Set saving mode as both. This allows users to download individual charts/ as well as download all charts as a single file.
    myExportComponent.componentAttributes.saveMode = 'both';
    //Show allowed export format drop-down
    myExportComponent.componentAttributes.showAllowedTypes = '1';
    //Cosmetics 
    //Width and height
    myExportComponent.componentAttributes.width = '350';
    myExportComponent.componentAttributes.height = '140';
    //Message - caption of export component
    myExportComponent.componentAttributes.showMessage = '1';
    myExportComponent.componentAttributes.message = 'Click on button above to begin export of charts. Then save from here.';
    //Render the exporter SWF in our DIV fcexpDiv
    myExportComponent.Render("fcexpDiv");           
    </script>


<%' 03 JUN 06 LJL lock out those without VN view right 106
end if
end if

IF request.form("goexcel") <> "" then

else
pv_last_update = "18 Feb 16"
rsys_db_select.close
Set rsys_db_select = nothing%>
<!--#include file = "includes/include_admin_frame_bottom.asp"-->
<%end if%>
