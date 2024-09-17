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
dim JAPINFO3sql,JAPINFO3
DIM UPDDsql, UPDD, goeditsql, goedit, logeditsql, logedit
Dim Dcount, faqid
'<<--Added by Interface on 05/02/2007
dim obj_db_CmdI, obj_db_CmdII, obj_logs_CmdI,obj_db_select_CmdI,obj_db_select_Cmd3
'-->>
'Added on 11/24/2008 DD
dim currentYear, pv_candidD

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

    ' Construct the SQL query string with parameters
    Dim updateSQL, queryString
    updateSQL = "UPDATE dbo.tx_rsys_candmisc SET canddiv_pronoun_id = ?, cand_div_race_ethnicity = ?, cand_div_disability = ?, cand_div_disability_accommodation = ?, cand_div_key_population = ? WHERE cand_id_c = ?"
    
    ' Build the query string with the parameters replaced for printing
    queryString = "UPDATE tx_rsys_candmisc SET " & _
        "cand_div_pronoun_id = '" & Request.Form("canddiv_pronoun_id") & "', " & _
        "cand_div_race_ethnicity = '" & race_ethnicity_value & "', " & _
        "cand_div_disability = '" & Request.Form("cand_div_disability") & "', " & _
        "cand_div_disability_accommodation = '" & Request.Form("cand_div_disability_accommodation") & "', " & _
        "cand_div_key_population = '" & Request.Form("cand_div_key_population") & "' " & _
        "WHERE cand_id_c = " & pv_candidD

    ' Print the query string for debugging
    Response.Write("<p><strong>SQL Query:</strong></p>")
    Response.Write("<pre>" & Server.HTMLEncode(queryString) & "</pre>")
    
    ' TO CHECK SQL STATEMENT
    'response.end

    ' Execute the SQL query with parameters
    Dim dbCmd
    Set dbCmd = Server.CreateObject("ADODB.Command")
    dbCmd.ActiveConnection = rsys_db
    dbCmd.CommandText = updateSQL
    dbCmd.Parameters.Append dbCmd.CreateParameter("@pronoun", adVarChar, adParamInput, 50, Request.Form("canddiv_pronoun_id"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@race_ethnicity", adVarChar, adParamInput, 255, race_ethnicity_value)
    dbCmd.Parameters.Append dbCmd.CreateParameter("@disability", adVarChar, adParamInput, 50, Request.Form("cand_div_disability"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@accommodation", adVarChar, adParamInput, 50, Request.Form("cand_div_disability_accommodation"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@key_population", adVarChar, adParamInput, 255, Request.Form("cand_div_key_population"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_id", adInteger, adParamInput, , pv_candidD)
    dbCmd.Execute()

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
JAPINFO3sql = "SELECT canddiv_pronoun_id,cand_div_race_ethnicity,cand_div_disability,cand_div_disability_accommodation, cand_div_key_population FROM tx_rsys_candmisc WHERE cand_id_c = ? "
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

    function checkOtherSelection() {
        var dropdown = document.getElementById('race_ethnicityDropdown');
        var otherTextbox = document.getElementById('otherrace_ethnicityTextbox');

        // Check if "Other" (value = 0) is selected
        var selectedOptions = Array.from(dropdown.selectedOptions);
        var isOtherSelected = selectedOptions.some(option => option.value === '0');

        if (isOtherSelected) {
            otherTextbox.style.display = 'inline'; // Show the textbox when "Other" is selected
        } else {
            otherTextbox.style.display = 'none';   // Hide the textbox when "Other" is not selected
        }
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
</script>

   
<form class="appForms" action="appDiv-edit.asp" method="POST" >
<%
Function GenerateDropdown(jsonText, dropdownName, selectedValue)
    Dim jsonKey, jsonOptions
    Dim keyStart, keyEnd, optionsStart, optionsEnd, optionsText
    Dim i, optionParts
    Dim optionValue, optionText
    
    ' Extract the key for the label
    keyStart = InStr(jsonText, "{") + 1
    keyEnd = InStr(keyStart, jsonText, ":{") - 1
    jsonKey = Mid(jsonText, keyStart, keyEnd - keyStart + 1)
    
    ' Remove any extra quotation marks
    jsonKey = Replace(jsonKey, """", "")

    ' Extract the options for the dropdown
    optionsStart = InStr(jsonText, ":{") + 2
    optionsEnd = InStrRev(jsonText, "}") - 1
    optionsText = Mid(jsonText, optionsStart, optionsEnd - optionsStart + 1)

    ' Remove extra characters like curly braces and quotes
    optionsText = Replace(optionsText, """", "")
    optionsText = Replace(optionsText, "{", "")
    optionsText = Replace(optionsText, "}", "")
    
    ' Convert optionsText to an array of key-value pairs
    Dim optionsArray
    optionsArray = Split(optionsText, ",")

    ' Form generation
%>  
    <tr>
        <td valign='top'>
            <label for="<%= dropdownName %>"><%= jsonKey %></label>
        </td>
        <td valign='top'>
            <select name="<%= dropdownName %>" id="<%= dropdownName %>" onchange="toggleOther('otherGenderRow', '<%= dropdownName %>')">
                <% 
                ' Populate dropdown options
                For i = 0 To UBound(optionsArray)
                    ' Extract key and value
                    optionParts = Split(optionsArray(i), ":")
                    If UBound(optionParts) = 1 Then
                        optionValue = Trim(optionParts(0))
                        optionText = Trim(optionParts(1))
                        
                        ' Write option to the dropdown
                        response.write "<option value=""" & optionValue & """"
                        If selectedValue = optionValue Then
                            response.write " selected"
                        End If
                        response.write ">" & optionText & "</option>"
                    End If
                Next
                %>
            </select>
            <span id="otherGenderRow" style="display:none;">
    <input type="text"  placeholder="specify other" name='<%= dropdownName %>' value="">
</span>

        </td>
    </tr>
    <% 
End Function
%>


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
            <select name="canddiv_pronoun_id"  onchange="toggleOther('otherpronounTextbox','canddiv_pronoun_id');" >
             
                <option value ="select"> select</option>
           
                   <option value="1" <% If JAPINFO3("canddiv_pronoun_id") = "1" Then Response.Write("selected") %>><% response.write gITEXT("i_92")%></OPTION>
                   <option value="2" <% If JAPINFO3("canddiv_pronoun_id") = "2" Then Response.Write("selected") %>><% response.write gITEXT("i_93")%></OPTION>
                   <option value="3" <% If JAPINFO3("canddiv_pronoun_id") = "3" Then Response.Write("selected") %>><% response.write gITEXT("i_94")%></OPTION>
                     <option value="0" <% If JAPINFO3("canddiv_pronoun_id") = "0" Then Response.Write("selected") %>><% response.write gITEXT("i_76")%></OPTION>
            </select>

                         <span id="otherpronounTextbox" style="display:none;">
   
    
        <input type="text" placeholder="specify other" name="cand_other_pronoun"  value="">
    
</span>

           	</td>
        </tr>
      

          		
          <TR>
    <td><% response.write gITEXT("i_87")%><FONT SIZE='4' COLOR='Red'>*</FONT> </td>
    <td>
     
         
                    <select name="cand_gnd_i" onchange="toggleOther('otherGenderTextbox','cand_gnd_i');">
                        <OPTION value="1" <% If JAPY("cand_gnd_i") = "1" then%> SELECTED<% End If%>><% response.write gITEXT("i_88")%></OPTION>
                        <OPTION value="2" <% If JAPY("cand_gnd_i") = "2" then%> SELECTED<% End If%>><% response.write gITEXT("i_89")%></OPTION>
                        <OPTION value="0" <% If JAPY("cand_gnd_i") = "0" then%> SELECTED<% End If%>><% response.write gITEXT("i_76")%></OPTION>
                    </select>
               
                     <span id="otherGenderTextbox" style="display:none;">
                    <input type="text"  placeholder="specify other" name="cand_other_gender" value="">
                </span>
         
    </td>
               <!-- The "Other" textbox, initially hidden, placed in the same row -->
<%
Dim jsonText
jsonText = gITEXT("i_text7") ' Fetch the JSON string from your function
' Call the function to generate the dropdown
GenerateDropdown jsonText, "genderDropdown", Request.Form("genderDropdown")
%>


          		
<!-- Race/Ethnicity -->
<tr>
    <td><% response.write gITEXT("i_95")%></td>

    <td valign='top'>
        <select name="cand_div_race_ethnicity[]" id="race_ethnicityDropdown" multiple onchange="checkOtherSelection();" size="5">
            <option value="1" <% If InStr(JAPINFO3("cand_div_race_ethnicity"), "1") > 0 Then Response.Write("selected") %>><% response.write gITEXT("i_96")%></option>
            <option value="2" <% If InStr(JAPINFO3("cand_div_race_ethnicity"), "2") > 0 Then Response.Write("selected") %>><% response.write gITEXT("i_97")%></option>
            <option value="3" <% If InStr(JAPINFO3("cand_div_race_ethnicity"), "3") > 0 Then Response.Write("selected") %>><% response.write gITEXT("i_98")%></option>
            <option value="4" <% If InStr(JAPINFO3("cand_div_race_ethnicity"), "4") > 0 Then Response.Write("selected") %>><% response.write gITEXT("i_99")%></option>
            <option value="0" <% If InStr(JAPINFO3("cand_div_race_ethnicity"), "0") > 0 Then Response.Write("selected") %>><% response.write gITEXT("i_76")%> <!-- "Other" --></option>
        </select>

        <!-- "Specify Other" textbox, initially hidden -->
        <span id="otherrace_ethnicityTextbox" style="display:none;">
            <input type="text" placeholder="Specify other race/ethnicity" name="cand_div_race_ethnicity_other" value="">
        </span>
    </td>
</tr>
      
          		<!-- Disability Inclusion -->
          	    <tr>
       	<td><% response.write gITEXT("i_text3")%> </td>

       	<td valign='top'>
     <select name="cand_div_disability" onchange="toggleOther('otherdisabilityTextbox','cand_div_disability')" >
      
         <option value ="select"> select</option>
    
            <option value="0" <% If JAPINFO3("cand_div_disability") = "0" Then Response.Write("selected") %>><% response.write gITEXT("i_text4")%></OPTION>
            <option value="2" <% If JAPINFO3("cand_div_disability") = "2" Then Response.Write("selected") %>><% response.write gITEXT("i_text5")%></OPTION>
           
     </select>

           

       	</td>
 </tr>


             		<!-- Disability acoomodation -->
         	    <tr id="otherdisabilityTextbox" style="display:none" >
      	<td><% response.write gITEXT("i_83")%> </td>

      	<td valign='top'>
    <select name="cand_div_disability_accommodation"  onchange="toggleOther('otherdisabilityaccomTextbox','cand_div_disability_accommodation')"; >
     
        <option value ="select"> select</option>
   
           <option value="2" <% If JAPINFO3("cand_div_disability_accommodation") = "2" Then Response.Write("selected") %>><% response.write gITEXT("i_text5")%></OPTION>
           <option value="0" <% If JAPINFO3("cand_div_disability_accommodation") = "0" Then Response.Write("selected") %>><% response.write gITEXT("i_text4")%></OPTION>
          
    </select>

          

      	</td>
</tr>

             	    <tr id="otherdisabilityaccomTextbox" style="display:none">
      	<td><% response.write gITEXT("i_84")%> </td>

      	<td valign='top'>
             

             <input type="text" placeholder="specify reasonable accomation" name="cand_div_disability_accom"  value="">

      	</td>
</tr>





    
    <!-- Text Area Row -->
    <tr>
        <td>
            <label for="keyPopulationsExplanation">You may want to explain whether and how you identify as part of key populations. 
</label>
            <textarea id="keyPopulationsExplanation" name="keyPopulationsExplanation" rows="4" cols="50"></textarea>
        </td>
    </tr>
    
          		<!-- Key Populations -->
          	<tr>
        <td>
            <button type="button" onclick="toggleDetails()">Show Details</button>
        </td>
    </tr>
    <!-- Paragraph Row (Hidden by Default) -->
<tr id="detailsRow" style="display:none;">
    <td>
        <p>
            The engagement of key populations is critical to a successful HIV response. Leadership by and greater involvement of communities living with and affected by HIV is essential to ending AIDS, therefore, applications from candidates belonging to these communities are especially welcome.
            <br><br>
            In light of these goals, we ask if you identify as part of key population(s).
            Key populations, or key populations at higher risk, are groups of people who are more likely to be exposed to HIV or to transmit it and whose engagement is critical to a successful HIV response. In all countries, key populations include people living with HIV. In most settings, men who have sex with men, trans(gender) people, people who inject drugs, and sex workers and their clients are at higher risk of exposure to HIV than other groups. These populations often suffer from punitive laws or stigmatizing policies, and they are among the most likely to be exposed to HIV.
        </p>
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
