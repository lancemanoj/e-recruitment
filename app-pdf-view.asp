<%option explicit
pv_col2left = "1"
' HEAVY BOTTOM
pv_heavybottomvery = "1"

' CHECK LOGGED IN AND DIM DB's
' ***********************************************************%>
<!--#include file = "../includes/include_check_login.asp"-->
<!--#include file = "../../includes/rsys_db_select.asp"-->
<!--#include file = "../../includes/rsys_int_select.asp"-->

<script src="js/jquery-1.8.2.js" type="text/javascript"></script>

        <script type ="text/javascript">
            $(document).ready(function () {
                $("#print").addClass("active");
            });
</script>
<%

' ***********************************************************
' END CHECK LOGGED IN AND DIM DB's
'<!-------------------------------------------NOTES ----------
'MODS --
'23 APR 07 INT Changed queries to parameterized
'20 APR 09 LJL added space at top of main page content
'29 March 17 manoj show higlighted menu bar
'01 SEP 24 LJL adjust fonts and styles

'--------------------------------------------------------------->
dim JAPINFO2sql, JAPINFO2, faqid

' BEGIN INCLUDE TEXT
dim gITEXTsql, gITEXT
'<<--Modified by Interface on 04/23/2007
gITEXTsql = "SELECT i_12, i_3, i_4, i_5, i_6, i_7, i_11 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'pdf' "
obj_int_select_Cmd.CommandText = gITEXTsql
Set gITEXT = obj_int_select_Cmd.Execute(,Array(session("template_org_code"),session("lng")))
'-->>
' END INCLUDE TEXT

pv_page_title = gITEXT("i_12")
titletype = 23
'faqid="25"
faqid=38


'<<--Modified by Interface on 04/23/2007
JAPINFOsql = "SELECT upd_d, cand_lnam_t, cand_fnam_t FROM td_rsys_cand WHERE cand_id_c = ? "
obj_db_select_Cmd.CommandText = JAPINFOsql
Set JAPINFO = obj_db_select_Cmd.Execute(,Array(session("RSYS_EVAL")))
'-->>


'if not Request.form("cand_email_t") then
'cand_email_t = ""
'end if
'if not Request.form("cand_pwd_c") then
'cand_pwd_c = ""
'end if


'Logsql = "INSERT INTO pub_log (upd_d,log_ip,log_host,pg_id_c, log_user, log_browser_type_c, log_referer_c, log_server_i) VALUES ('"& Date(Now() &"','"& request.servervariables("REMOTE_ADDR") &"','"& left(request.servervariables("REMOTE_HOST"), 50) &"','vac-pdf-"& session("lng") &"', "'"
'IF (session("rsysuser") and LEFT("rsysuser","20") then else PUBLIC-NLI</cfif>', 
'IF isdefined("HTTP_USER_AGENT")>#LEFT(TRIM(HTTP_USER_AGENT),50)#<CFELSE>UKN</CFIF>', 
			'<CFIF isdefined("HTTP_REFERER")>#RIGHT(TRIM(HTTP_REFERER),50)#<CFELSE>-</cfif>'server_int_ext 			) 	"
'set Log = rsys_logs.execute(Logsql)

' ***********************************************************
' BEGIN INCLUDES
' ***********************************************************%>
<!--#include file="../includes/include_pubedit_frame_top.asp"-->
<%
' ***********************************************************
' END INCLUDES
' ***********************************************************%> 
<TABLE cellpadding="0" BORDER="0" width="100%">
<TR>
	<td valign='top'>&nbsp;</td>
</tr>
<TR>
    	<td valign='top' class='blacktext' colspan='3'><% response.write gITEXT("i_3")%> :</td>
 </tr>
<TR>
	<td valign='top'>&nbsp;</td>
</tr>
<%
'<TR>
'	<td valign='top'>The creation of Personal History output is currently off line.  It should resume soon.</td>
'</tr>

'<TR>
 '   	<td valign='top'>1. &nbsp;</td>
  '  	<td valign='top'><a href='../pdf/docsettings.asp?goHTML=1' target='_blank'>[ <% response.write gITEXT("i_4")///> ]</a></td>
'    	<td valign='top'><% response.write gITEXT("i_5")///>.</td>
'</tr>
'<TR>
'	<td valign='top' colspan='3'>&nbsp;</td>
'</tr>
%>
<TR>
    	<td valign='top'>1. &nbsp;</td>
    	<td valign='top'nowrap><a href='../pdf/docsettings.asp?goPDF=1' target='_blank'>[ <% response.write gITEXT("i_6")%> ]</a></td>
    	<td valign='top'><% response.write gITEXT("i_7")%>.</td>
</tr>
<TR>
	<td valign='top' colspan='3'>&nbsp;</td>
</tr>
<TR>
    <td valign='top'>2. &nbsp;</td>
    <td valign='top' nowrap><a href='../pdf/docsettings.asp?goWord=1' target='_blank'>[ Word ]</a></td>
    <td valign='top'><% response.write gITEXT("i_11")%></td>
</tr>
<TR>
	<td valign='top' colspan='3'>&nbsp;</td>
</tr>
</table>
<%pv_last_update = "01 Sep 24"
' *****************************************************
' BEGIN BOTTOM INCLUDES
' *****************************************************%>
<!--#include file="../includes/include_pubedit_frame_bottom.asp"-->
<% 
' *****************************************************
' END BOTTOM INCLUDES
' *****************************************************%>

<%'<<--Added By Interface on 04/23/2007
Set obj_int_select_Cmd = nothing
Set obj_db_select_Cmd = nothing
'-->> 
rsys_db_select.close
Set rsys_db_select = nothing
rsys_int_select.close
Set rsys_int_select = nothing%>