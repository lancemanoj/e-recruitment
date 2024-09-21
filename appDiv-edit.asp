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
<!--#include file="functions.asp"-->
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
dim JAPINFO3sql,JAPINFO3
DIM UPDDsql, UPDD, goeditsql, goedit, logeditsql, logedit
Dim Dcount, faqid
'<<--Added by Interface on 05/02/2007
dim obj_db_CmdI, obj_db_CmdII, obj_logs_CmdI,obj_db_select_CmdI,obj_db_select_Cmd3
'-->>
'Added on 11/24/2008 DD
dim currentYear, pv_candidD
    Function CleanInput(input)
    CleanInput = Trim(input)
    If Right(CleanInput, 1) = "," Then
        CleanInput = Left(CleanInput, Len(CleanInput) - 1)
    End If
End Function
' BEGIN INCLUDE TEXT
dim gITEXTsql, gITEXT
'<<--Modified by Interface on 05/02/2007
gITEXTsql = "SELECT i_text3,i_text4,i_text5, i_1, i_4, i_23, i_text1, i_27,i_5, i_16, i_17, i_18,i_32,i_15, i_9,i_11, i_10,i_19, i_20, i_29,i_30, i_22, i_21,i_13,i_76,i_83,i_84, i_87,i_88,i_89,i_91,i_92,i_93,i_94,i_95,i_96,i_97,i_98,i_99,i_text7 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'Div' "
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


' Form Submission Logic
Dim action
action = Request.Form("action")


If action = "Save" Then
    ' Ensure race_ethnicity is treated as an array
    Dim race_ethnicity_value
    If IsArray(Request.Form("cand_div_race_ethnicity")) Then
        race_ethnicity_value = Join(Request.Form("cand_div_race_ethnicity"), ",")
    Else
        race_ethnicity_value = Request.Form("cand_div_race_ethnicity")
    End If
    ' Clean up Request.Form values

    ' Construct the SQL query string with parameters
     Call InsertOrUpdateCandidate(rsys_db, pv_candidD, race_ethnicity_value)

    ' Redirect back after save
    Response.Redirect "appDiv-edit.asp"

ElseIf action = gITEXT("i_13") Then
    ' Handle the comment action and redirect
    Response.Redirect "appD-edit.asp"
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

JAPYsql = "SELECT cand_gnd_i, upd_d, cand_lnam_t, cand_fnam_t, cand_comp_os_i, cand_comp_os_other_t, cand_wp_i, cand_wp_other_t, cand_sps_i, cand_sps_other_t, cand_db_i, cand_db_other_t, cand_pres_i, cand_pres_other_t, cand_web_i, cand_web_other_t, cand_prgming_i, cand_prgming_other_t, cand_pc_other_t, cand_pc_skills_t,cand_div_pronouns, cand_div_gender_doc, cand_div_gender_identity, cand_div_race_ethnicity, cand_div_disability, cand_div_accommodation, cand_div_key_population FROM td_rsys_cand 	WHERE cand_id_c = ?"
obj_db_select_Cmd.CommandText = JAPYsql
Set JAPY = obj_db_select_Cmd.Execute(,Array(pv_candidD)) 

'Modified by Interface , 30/07/2007
set obj_db_select_CmdI = server.CreateObject("adodb.command")
obj_db_select_CmdI.ActiveConnection = rsys_db_select
'----
JAPINFO2sql = "SELECT editD FROM tx_rsys_candedit WHERE cand_id_c = ? "
obj_db_select_CmdI.CommandText = JAPINFO2sql
Set JAPINFO2 = obj_db_select_CmdI.Execute(,Array(pv_candidD)) 


set obj_db_select_Cmd3 = server.CreateObject("adodb.command")
  obj_db_select_Cmd3.ActiveConnection = rsys_db_select
JAPINFO3sql = "SELECT  cand_div_keyPopulationsExplanation,cand_div_disability_accom, cand_div_other_raceethnicity, cand_div_other_genderidentity,cand_div_genderidentity, cand_other_gender,cand_gnd_i,cand_other_pronoun,canddiv_pronoun_id,cand_div_race_ethnicity,cand_div_disability,cand_div_disability_accommodation, cand_div_key_population FROM tx_rsys_candmisc WHERE cand_id_c = ? "
obj_db_select_Cmd3.CommandText = JAPINFO3sql

'response.write JAPINFO3sql
	Set JAPINFO3 = obj_db_select_Cmd3.Execute(,Array(pv_candidD))



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

<style>
    .wide {
        width: 100%;
    }
</style>
<script type="text/javascript">
      function toggleOther(textboxId,field) {
          var genderDropdown = document.getElementsByName(field)[0];
          var selectedValue = genderDropdown.value;
          var otherTextbox = document.getElementById(textboxId);
          
          if (selectedValue === "0" || selectedValue === "-1") { // Assuming "2" corresponds to "Other"
              otherTextbox.style.display = "block";
          } else {
              otherTextbox.style.display = "none";
          }
      }

      function toggleOtherTextbox() {
          var checkboxes = document.getElementsByName('cand_div_race_ethnicity');
          var otherTextbox = document.getElementById('OtherraceethnicityTextbox');
          var showOther = false;
          for (var i = 0; i < checkboxes.length; i++) {
              if (checkboxes[i].value == '-1' && checkboxes[i].checked) {
                  showOther = true;
                  break;
              }
          }
          otherTextbox.style.display = showOther ? 'block' : 'none';
      }

        function toggleDetails() {
        var detailsRow = document.getElementById('detailsRow');
        var button = event.target;

        if (detailsRow.style.display === 'none') {
            detailsRow.style.display = 'table-row';
        button.textContent = 'Hide Details';
        } else {
            detailsRow.style.display = 'none';
        button.textContent = 'Show Details';
        }
        }


        window.onload = function() {

            toggleOtherTextbox();
            

            var disabilitySelect = document.getElementsByName('cand_div_disability')[0];
            if (disabilitySelect.value === "0") {
                document.getElementById('otherdisabilityTextbox').style.display = 'table-row';
            }

            // Check for cand_div_disability_accommodation on page load
            var disabilityAccomSelect = document.getElementsByName('cand_div_disability_accommodation')[0];
            if (disabilityAccomSelect.value === "0") {
                document.getElementById('otherdisabilityaccomTextbox').style.display = 'table-row';
            } else {
                document.getElementById('otherdisabilityaccomTextbox').style.display = 'none';
            }

         //    Define an array of dropdown IDs and their corresponding textbox IDs
            var fields = [
                { dropdownId: 'canddiv_pronoun_id', textboxId: 'otherpronounTextbox' },
                 { dropdownId: 'cand_gnd_i', textboxId: 'otherGenderTextbox' },
                  { dropdownId: 'cand_div_genderidentity', textboxId: 'OtherIdentityTextbox' }
        
                // Add more fields here following the same pattern
            ];

            // Loop through each field and call toggleOther based on the current selected value
            fields.forEach(function(field) {
                toggleOther(field.textboxId, field.dropdownId);
            });
        };
     

</script>

   
<form class="appForms" action="appDiv-edit.asp" method="POST" >



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
    	<!-- Personal Pronouns -->
          		<tr>
           	<td><% response.write gITEXT("i_91")%> </td>

           	<td valign='top'>
           <select name="canddiv_pronoun_id" onchange="toggleOther('otherpronounTextbox', 'canddiv_pronoun_id');">
    <option value="select">select</option>
    <option value="<%= gITEXT("i_92") %>" <% If JAPINFO3("canddiv_pronoun_id") = gITEXT("i_92") Then Response.Write("selected") %>><%= gITEXT("i_92") %></option>
    <option value="<%= gITEXT("i_93") %>" <% If JAPINFO3("canddiv_pronoun_id") = gITEXT("i_93") Then Response.Write("selected") %>><%= gITEXT("i_93") %></option>
    <option value="<%= gITEXT("i_94") %>" <% If JAPINFO3("canddiv_pronoun_id") = gITEXT("i_94") Then Response.Write("selected") %>><%= gITEXT("i_94") %></option>
   <option value="0" <% If JAPINFO3("canddiv_pronoun_id") = "0" Then Response.Write("selected") %>><% response.write gITEXT("i_76")%></OPTION>
</select>

                         <span id="otherpronounTextbox" style="display:none;">
   
    
        <input type="text" placeholder="specify other" name="cand_other_pronoun"  value="<% =JAPINFO3("cand_other_pronoun") %>" />
        </span>
   

           	</td>
        </tr>
      

          		
          <TR>
    <td><% response.write gITEXT("i_87")%><FONT SIZE='4' COLOR='Red'>*</FONT> </td>
    <td>
     
         
                    <select name="cand_gnd_i" onchange="toggleOther('otherGenderTextbox','cand_gnd_i');">
                        <OPTION value="<%= gITEXT("i_88") %>" <% If JAPINFO3("cand_gnd_i") =gITEXT("i_88") then%> SELECTED<% End If%>><% response.write gITEXT("i_88")%></OPTION>
                        <OPTION value="<%= gITEXT("i_89") %>" <% If JAPINFO3("cand_gnd_i") = gITEXT("i_89") then%> SELECTED<% End If%>><% response.write gITEXT("i_89")%></OPTION>
                        <OPTION value="0" <% If JAPINFO3("cand_gnd_i") = "0" then%> SELECTED<% End If%>><% response.write gITEXT("i_76")%></OPTION>
                    </select>
               
                     <span id="otherGenderTextbox" style="display:none;">
                    <input type="text"  placeholder="specify other" name="cand_other_gender" value="<% =JAPINFO3("cand_other_gender") %>">
                </span>
        
    </td>
               <!-- The "Other" textbox, initially hidden, placed in the same row -->

<%
    Dim jsonText
    jsonText = gITEXT("i_text7") ' Fetch the JSON string from your function
    ' Call the function to generate the dropdown
    GenerateDropdown jsonText, "cand_div_genderidentity", JAPINFO3("cand_div_genderidentity"),"OtherIdentityTextbox",JAPINFO3("cand_div_other_genderidentity")
    %>

          		
<!-- Race/Ethnicity -->
<tr>
    

   
      <%
                ' Fetch the JSON data for the race/ethnicity checkboxes
                Dim jsonTextRaceEthnicity
                jsonTextRaceEthnicity = gITEXT("i_96") ' Use the correct JSON string function
                
                ' Call the GenerateCheckboxGroup function
                GenerateCheckboxGroup jsonTextRaceEthnicity, "cand_div_race_ethnicity", JAPINFO3("cand_div_race_ethnicity"), "OtherraceethnicityTextbox", JAPINFO3("cand_div_other_raceethnicity")
                %>
    </td>
</tr>
      
          		<!-- Disability Inclusion -->
          	    <tr>
       	<td><% response.write gITEXT("i_text3")%> </td>

       	<td valign='top'>
     <select name="cand_div_disability" class="wide" onchange="toggleOther('otherdisabilityTextbox','cand_div_disability')" >
      
         <option value ="select"> select</option>
    
            <option value="0" <% If JAPINFO3("cand_div_disability") = "0" Then Response.Write("selected") %>><% response.write gITEXT("i_text4")%></OPTION>
            <option value="1" <% If JAPINFO3("cand_div_disability") = "1" Then Response.Write("selected") %>><% response.write gITEXT("i_text5")%></OPTION>
           
     </select>

           

       	</td>
 </tr>


             		<!-- Disability acoomodation -->
         	    <tr id="otherdisabilityTextbox" style="display:none" >
      	<td><% response.write gITEXT("i_83")%> </td>

      	<td valign='top'>
    <select class="wide" name="cand_div_disability_accommodation"  onchange="toggleOther('otherdisabilityaccomTextbox','cand_div_disability_accommodation')"; >
     
        <option value ="select"> select</option>
   
           <option value="1" <% If JAPINFO3("cand_div_disability_accommodation") = "1" Then Response.Write("selected") %>><% response.write gITEXT("i_text5")%></OPTION>
           <option value="0" <% If JAPINFO3("cand_div_disability_accommodation") = "0" Then Response.Write("selected") %>><% response.write gITEXT("i_text4")%></OPTION>
          
    </select>

          

      	</td>
</tr>
    <tr id="otherdisabilityaccomTextbox" style="display:none">
    <td><% response.write gITEXT("i_84")%></td>
    <td valign='top'>
        <textarea class="wide" placeholder="Specify reasonable accommodation" name="cand_div_disability_accom" rows="4" cols="50">
            <% =JAPINFO3("cand_div_disability_accom") %>
        </textarea>
    </td>
</tr>






    
    <!-- Text Area Row -->
    <tr>
        <td>
            <label for="keyPopulationsExplanation">You may want to explain whether and how you identify as part of key populations. 
</label>
            <textarea id="cand_div_keyPopulationsExplanation" name="cand_div_keyPopulationsExplanation" rows="4" cols="50">
                <%=Server.HTMLEncode(JAPINFO3("cand_div_keyPopulationsExplanation") & "")%>
 
</textarea>
        </td>
    </tr>
    
          		<!-- Key Populations -->
          	<tr>
        <td>
            <button type="button" onclick="toggleDetails()">Hide Details</button>
        </td>
    </tr>
    <!-- Paragraph Row (Hidden by Default) -->
<tr id="detailsRow" style="display:block;">
    <td>
        <%=response.write(gITEXT("i_99") & "")%>
    </td>
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

             

            	<td  colspan="2" valign="bottom" align='center'>
              			<INPUT TYPE="hidden" NAME="cand_id_c" VALUE="<%=pv_candidD%>">
				<INPUT TYPE="hidden" NAME="editD" VALUE="<% response.write Dcount%>">
				<INPUT TYPE="hidden" NAME="GOeditD" VALUE="99">
                               <input type="submit" name="action" value="Save">
                <input type="submit" name="action" value="<%= gITEXT("i_13") %>">
				
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
