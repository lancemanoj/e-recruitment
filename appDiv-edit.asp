<%option explicit

'<!-------------------------------------------NOTES ----------
'MAY HAVE ISSUES WITH SINGLE QUOTES IN DATA UPDATES.
'did set the query to be CFUPDATE, but that won't work with ASP and also,
'it was causing memory problems when updating and memory high on server. 
'MODS --
'14 JUN 05 LJL adjusted update query on app-edit-start.asp to be not CFUPDATE
'06 SEP 05 RR reviewed
' 25 SEPT LJL ready except for update text
'10 MAR 06 LJL updated logging table for candupdates
'08 JUL 06 LJL moved notes to top, added update text as include, which used to be in the main page top include
'01 NOV 06 LJL added space after button
'02 MAY 07 INT Changed queries to parameterized
'28 JUL 07 LJL modified adVarChar to adLongVarChar for text/memo fields - added replace check for NULL entries when it is to go to REPLACE()
'03 AUG 07 INT - Removed replace function from query parameter.
'17 AUG 07 LJL page was using wrong query
'24 NOV 08 DD  modified rsyslog table names to conatin current year's tag.
'06 AUG 09 LJL added coding to log when user ACCESSES, or UPDATES profile into td_rsys_candupdates_...
'01 MAY 10 LJL formatting
'01 MAY 10 LJL moved header set and check login to before loading content - 
	'order is Edit Section, check_complete, then page headers, then ejobs-updates include
'04 APR 14 LJL removed validation check script - wasn't allowing html scripting
'14 OCT 15 GG added additional checking for size of field

'11 OCT 19 GG Added Server.HTMLEncode 


'--------------------------------------------------------------->

' TEMPORARY TAKEOUT VARS

' SET TO HAVE LEFT MENU
pv_col3 = "1"
pv_heavybottom = "1"

' CHECK LOGGED IN AND DIM DB's
' ***********************************************************%>
<!--#include file = "../includes/include_check_login.asp"-->
<!--#include file = "../../includes/rsys_db.asp"-->
<!--#include file = "../../includes/rsys_db_select.asp"-->
<!--#include file = "../../includes/rsys_int_select.asp"-->
<!--#include file = "../../includes/rsys_logs.asp"-->
<% '<<--Modified by Interface on 05/02/2007 %>
<!--#include file="../../sysdev/adovbs.inc" -->
<%'-->>%>
<%
' ***********************************************************
' END CHECK LOGGED IN AND DIM DB's
dim JAPYsql, JAPY
dim JAPINFO2sql, JAPINFO2
DIM UPDDsql, UPDD, goeditsql, goedit, logeditsql, logedit
Dim Dcount, faqid
'<<--Added by Interface on 05/02/2007
dim obj_db_CmdI, obj_db_CmdII, obj_logs_CmdI,obj_db_select_CmdI
'-->>
'Added on 11/24/2008 DD
dim currentYear, pv_candidD

' BEGIN INCLUDE TEXT
dim gITEXTsql, gITEXT
'<<--Modified by Interface on 05/02/2007
gITEXTsql = "SELECT i_1, i_4, i_23, i_text1, i_27,i_5, i_16, i_17, i_18,i_32,i_15, i_9,i_11, i_10,i_19, i_20, i_29, i_22, i_21,i_13 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'Div' "
obj_int_select_Cmd.CommandText = gITEXTsql
Set gITEXT = obj_int_select_Cmd.Execute(,Array(session("template_org_code"),session("lng")))
'-->>
' END INCLUDE TEXT

pv_page_title = gITEXT("i_1")
titletype = 11
'faqid="14"
faqid=14
currentYear =  year(now())  'Added on 11/24/2008 DD, to get the value of current year.

pv_candidD = session("RSYS_EVAL")


'******************************
' BEGIN Div EDIT
'******************************
If Request.form("GOeditDiv") = "99" then


else

	
End If

'******************************
' END D EDIT
'******************************
' ***********************************************************
' BEGIN INCLUDES
' ***********************************************************%>
<!--#include file = "../includes/include_check_complete.asp"-->
<!--#include file="../includes/include_pubedit_frame_top.asp"-->
<%
' ***********************************************************
' END INCLUDES
' ***********************************************************
'08 JUL 06 LJL added UPDATE PHRASE HERE%>
<!-- #include file="../edit/ejobs-updates.asp"-->
<%  

'<<--Modified by Interface on 05/02/2007
//Modified on 06/01/2009,DD
set obj_db_select_Cmd = server.CreateObject("adodb.command")
obj_db_select_Cmd.ActiveConnection = rsys_db_select

JAPYsql = "SELECT upd_d, cand_lnam_t, cand_fnam_t, cand_comp_os_i, cand_comp_os_other_t, cand_wp_i, cand_wp_other_t, cand_sps_i, cand_sps_other_t, cand_db_i, cand_db_other_t, cand_pres_i, cand_pres_other_t, cand_web_i, cand_web_other_t, cand_prgming_i, cand_prgming_other_t, cand_pc_other_t, cand_pc_skills_t FROM td_rsys_cand 	WHERE cand_id_c = ?"
obj_db_select_Cmd.CommandText = JAPYsql
Set JAPY = obj_db_select_Cmd.Execute(,Array(pv_candidD)) 

'Modified by Interface , 30/07/2007
set obj_db_select_CmdI = server.CreateObject("adodb.command")
obj_db_select_CmdI.ActiveConnection = rsys_db_select
'----
JAPINFO2sql = "SELECT editD FROM tx_rsys_candedit WHERE cand_id_c = ? "
obj_db_select_CmdI.CommandText = JAPINFO2sql
Set JAPINFO2 = obj_db_select_CmdI.Execute(,Array(pv_candidD)) 
'-->>
%>

<%
'<head>
'<SCRIPT language="JavaScript" type="text/javascript" src="../../js/Validation_Script.js"></script>
'<SCRIPT language="JavaScript" type="text/javascript">
'function DataValidation()
 '   {
'		// To prevent 'script' to be included in text. 
'		// Modified by Interface on 07/23/2007 
'		return ValidateForm(document.forms[0]);
'	}
'</SCRIPT>

'</head>
%>

   
<form class="appForms" action="appD-edit.asp" method="POST" ONSUBMIT="return DataValidation();">
  <TABLE cellpadding="0" border="0" width="100%">
        <%' REMOVED Do while JAPINFO.eof = false%>
          		<tr>
            			<td valign='top' colspan="2"><%=gITEXT("i_4")%></td>
          		</TR>
          		<tr>
            			<td valign='top' colspan="2">&nbsp;</td>
          		</TR>
          		<tr>
            			<td valign='top' colspan="2"><%=gITEXT("i_23")%></td>
          		</TR>
          		<tr>
            			<td valign='top' colspan="2">&nbsp;</td>
          		</TR>
          		<tr>
            			<td valign='top'><%=gITEXT("i_text1")%></td>
          		</TR>
          		<tr>
            			<td valign='top' colspan="2">&nbsp;</td>
          		</TR>
          		<tr>
            			<td valign='top'><%=gITEXT("i_27")%></td>
          		</TR>
          		<tr>
            			<td valign='top'>&nbsp;</td>
          		</TR>
</table>

          		
<table cellpadding="0" border="0" width="100%">
          		<TR>
            			<td valign='top' valign='top' colspan='3' class="textbold"><%=gITEXT("i_22")%></td>
          		</tr>
          		<TR>
            			<td valign='top'><textarea  wrap="soft" ROWS="6" NAME="cand_pc_skills_t" COLS="50"><%=Server.HTMLEncode(JAPY("cand_pc_skills_t") & "")%></TEXTAREA></td>
          		</TR>
          		<TR>
            			<td valign='top' valign='top' colspan='3' class="textbold"><%=gITEXT("i_21")%></td>
          		</tr>
          		<TR>
            			<td valign='top'><textarea  wrap="soft" ROWS="6" NAME="cand_pc_other_t" COLS="50"><%=Server.HTMLEncode(JAPY("cand_pc_other_t") & "")%></TEXTAREA></td>
          		</TR>
          		<TR>
            			<td valign='top' valign='top' colspan='3' class="textbold">&nbsp;</td>
          		</tr>
          <%
          Dcount = int(japinfo2("EditD") + 1)
          'JAPINFO.movenext
          'loop%>
          <tr valign="bottom">
            	<td valign="bottom" align='center'>
              			<INPUT TYPE="hidden" NAME="cand_id_c" VALUE="<%=pv_candidD%>">
				<INPUT TYPE="hidden" NAME="editD" VALUE="<% response.write Dcount%>">
				<INPUT TYPE="hidden" NAME="GOeditD" VALUE="99">
                <INPUT  TYPE="submit" class="submit" VALUE="<% response.write gITEXT("i_13")%> ">
				
                	</td>
              </tr>
<tr valign="bottom">
	<td valign="bottom">&nbsp;</td>
</tr>
</TABLE>
</form>
<%pv_last_update="22 Aug 24"
'<<--Added by Interface on 05/02/2007
'set obj_int_select_Cmd = nothing
'set obj_db_CmdII = nothing
'set obj_logs_CmdI = nothing
'set obj_db_select_Cmd = nothing
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
//Added on 06/01/2009,DD, following include of checkbox updatation changes.%>
<!--#include file="../includes/include_pubedit_frame_bottom.asp"-->
<% 
' *****************************************************
' END BOTTOM INCLUDES
' *****************************************************%>
