<%option explicit
' 14 DEC 05 LJL made so that the sections that are complete in the profile are automatically checked in this page
' 14 DEC 05 LJL no contract details for WTO
'05 JUL 06 REMOVED GENERAL COVERING LETTER 
'31 JUL 06 LJL adjusted covering letters, post specific to show now.
'15 AUG 06 LJL added TOC (table of contents) setting
'16 dec 06 ac - add pdf to show pic if candidate wannts
'19 feb 07 ac - added new temp /public_hold when fixed changed back /public
'21 feb 07 ac - change pub.. back to public
'23 APR 07 INT Changed queries to parameterized
'01 FEB 08 LJL removed AoE per ILO
'16 MAR 08 LJL no photo for HTML versions
'03 APR 09 DD  added OTHER INFORMATION AND REFERENCE sections for WTO, and blocked Clearical Skills for this org.
'06 APR 09 DD  added Secretarial Skills for WTO org.
'11 AUG 10 DD Replaced text 'Include Table of Contents?' by 'RC/RC Experience' for selection checkbox
'20 AUG 10 DD Added checkbox to include country list of RC/RC Experience section
'04 FEB 12 LJL changed Covering letter text to be for all CL listed, not vacancy specific, as that does not apply for applicant outputs of CV
'03 MAR 15 LJL revised directories for WTO, IFRC for HTML file creation
'28 APR 15 LJL no clerical for WMO
'16 DEC 20 LJL checked and validated for foreign chars and code pages
'16 DEC 20 LJL tried a number of things to get correct accented chars outputted.  Ended up having to just add <head><meta http-equiv='Content-Type' content='text/html; charset=ISO-8859-1'> to admin/ACdoc-maker-new.asp for all outputs via the file stream write


'--------------------------------------------------------------->

' SET TO HAVE LEFT MENU
pv_col2left = "1"

' CHECK LOGGED IN AND DIM DB's
' ***********************************************************%>
<!--#include file = "../includes/include_check_login.asp"-->
<!--#include file = "../../includes/rsys_db_select.asp"-->
<!--#include file = "../../includes/rsys_int_select.asp"-->
<%
' ***********************************************************
dim dir_path
' END CHECK LOGGED IN AND DIM DB's
' BEGIN INCLUDE TEXT
dim faqid

dim gITEXTsql, gITEXT, gITEXTTsql, gITEXTT
'<<--Modified by Interface on 04/23/2007
gITEXTsql = "SELECT * FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'PDF' "
obj_int_select_Cmd.CommandText = gITEXTsql
Set gITEXT = obj_int_select_Cmd.Execute(,Array(session("template_org_code"),session("lng")))

//added by DD,04/06/2009, for 'secreterial skills' title for wto
gITEXTTsql=" SELECT i_1 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'T' "
obj_int_select_Cmd.CommandText = gITEXTTsql
set gITEXTT = obj_int_select_Cmd.Execute(,Array(session("template_org_code"),session("lng")))

'-->>
pv_page_title = gITEXT("i_12")
titletype = 23
faqid="25"
  

' ***********************************************************
' BEGIN INCLUDES
' ***********************************************************%>
<!--#include file="../includes/include_pubedit_frame_top.asp"-->
<%
' ***********************************************************
' END INCLUDES
' ***********************************************************

' response.write "<br>DIR: " & request.servervariables("PATH_TRANSLATED") & "<br>"
dir_path = Replace(request.servervariables("PATH_TRANSLATED"), "\public\pdf\docsettings.asp", "\public\pdf")
'dir_path = Replace(request.servervariables("PATH_TRANSLATED"), "\public\pdf\docsettings.asp", "\public\pdf")
' response.write "DIR2: " & dir_path & "<br>"
' response.end

'******************************************
'DELETE OLD PDF - more than an hour old
'******************************************
Dim filecount, fso, file, f :rem Set up a variable for counting the number of files
Set fso = CreateObject("Scripting.FileSystemObject")
' Call the file system object to manipulate files
Set f = fso.GetFolder(dir_path) :rem use any folder that you want here
' Assign "f" the folder J-ComEDI so that the files within this directory can be manipulated
filecount = 0 :rem Clear the variable before use in our filecount loop

' If there are files in the J-CommEDI folder with today's date on
' them then print the name of the file to the screen and increment
' the counter.
' RESPONSE.write "DIFF: " & dateadd("h", -1, now()) & "TIMENOW: " & now()

  On Error Resume Next

For Each file in f.Files
IF file.DateCreated < dateadd("h", -1, now()) and fso.GetExtensionName(file) = "pdf" then
	' response.write "The file " & file.Name & " was created on " & FormatDateTime(file.DateCreated,vbShortDate)
	' response.write "The file " & file.Name & " was created on " & file.DateCreated & "EXT: " & fso.GetExtensionName(file) & "<br>"
	filecount = filecount + 1
	' response.write "<br>PATH: " & file & "<br>"
     fso.DeleteFile(file)
' else
	' response.write "NADA"
End if 
next
f.Close ' Make sure you close it or it won't write it!!
Set f = Nothing
Set fso = Nothing 
'******************************************
'DELETE OLD PDF - more than an hour old
'******************************************
' INFO -- DO NOT COPY THIS FROM THE ADMIN VERSION OF DOCSETTINGS.asp

' END INCLUDE TEXT
%> 


<script language="Javascript" type="text/javascript">
function Controldisplay(chk)
{	
	if(chk.checked == 1){
		//alert("Checked");
		document.getElementById("countrylist").style.display="";
	}else{
		//alert("Unchecked");
		document.getElementById("countrylistcheckbox").checked = 0		
		document.getElementById("countrylist").style.display="none";		
	}
}
</script>


<table width="100%">
<tr>
	<td>&nbsp;</td>
</tr>
<%
if request.querystring("goPDF") = "1" then
'  <!-- <form action=../../doccreate/make-doc-prep.asp" method="post"> -->
%>
<form action="make-doc-prepNEW.asp" method="post">
<%
elseif request.querystring("goWord") = "1" then%>
<form action="make-doc-prepNEW.asp" method="post">
<input type="hidden" name="goWord" value="99">
<%
elseif request.querystring("goHTML") = "1" then%>
<form action="make-doc-prepNEW.asp" method="post">
<input type="hidden" name="goHTML" value="99">
<%end if%>
<tr>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="A"<%If session("OKA") <> "0" then%> checked<%end if%>> <%=gITEXT("i_19")%></td>
<%' NO CLERICAL FOR IFRC 7000
//Added by DD, 04/02/2009 ,  NO CLERICAL FOR WTO 3000
//if session("template_org_code") <> 7000 then
'28 APR 15 LJL no clerical for WMO

if session("template_org_code") <> 7000 AND session("template_org_code") <> 2900 then
    
    if session("template_org_code") <> 3000 then %>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="T"<%If session("OKT") <> "0" then%> checked<%end if%>> <%=gITEXT("i_22")%></td>
	
	<%else%>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="T"<%If session("OKT") <> "0" then%> checked<%end if%>> <%=gITEXTT("i_1")%></td>
	
	<%end if%>

<%elseif session("template_org_code") = 7000 then%>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="RC"<%If session("OKRC") <> "0" then%> checked<%end if%> onClick="Controldisplay(this)"> <%'=gITEXT("i_60")%> <%=gITEXT("i_42")%>  </td>
	<td id="countrylist" <%If session("OKRC") =  0 then%>style="display:none"<%End if%>>&nbsp;&nbsp;&nbsp;<input type='checkbox' id="countrylistcheckbox" name="app" value="RCCountryList"<%If session("OKRC") <> "0" then%> checked <%end if%>><%=gITEXT("i_65")%> (<%=gITEXT("i_42")%>)</td>	

<%elseif session("template_org_code") = 2900 then%>
	<td>&nbsp;</td>	


<%end if%>
</tr>
<tr>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="W"<% If session("OKAdd") <> "0" then%> checked<%end if%>> <%=gITEXT("i_48")%></td>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="D"<%If session("OKD") <> "0" then%> checked<%end if%>> <%=gITEXT("i_23")%>            		</td>
</tr>
<tr>
<% if session("template_org_code") <> 2000 then%>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="S"<%If session("OKS") <> "0" then%> checked<%end if%>> <%=gITEXT("i_25")%></td>
<% end if %>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="G"<%If session("OKG") <> "0" AND session("OKGref") <> "0" then%> checked<%end if%>> <%=gITEXT("i_27")%>            		</td>
</tr>
<tr>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="E"<%If session("OKE") <> "0" AND session("educheck") <> "0" AND session("educheck2") <> "0"then%> checked<%end if%>> <%=gITEXT("i_24")%></td>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="J"<%If session("OKJ") <> "0" then%> checked<%end if%>> <%=gITEXT("i_47")%></td>
</tr>
<tr>
<%if session("template_org_code") <> 3000 then%>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="B"<%If session("OKB") <> "0" then%> checked<%end if%>> <%=gITEXT("i_20")%></td>
<%end if%>
</tr>
<tr>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="F"<%If session("employcheck") <> "0" then%> checked<%end if%>> <%=gITEXT("i_26")%>            		</td>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="Y"> <%=gITEXT("i_28")%></td>
</tr>
<tr>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="C"<% If session("langcheck") <> "0" AND session("OKC") <> "0" then%> checked<%end if%>> <%=gITEXT("i_21")%></td>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="V"> <%=gITEXT("i_49")%></td>
</tr>
<tr>
	<%if session("template_org_code") = 3000 then%>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="OI"<%If session("OKG") <> "0" AND session("OKGref") <> "0" then%> checked<%end if%>> <%=gITEXT("i_50")%></td>		
	<%end if%>
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="X"> <%=gITEXT("i_70")%></td>
</tr>
<% 
//Added by DD, 04/03/2009, recommended for wto only.
if session("template_org_code") = 3000 then%>
<tr>	
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="GR"<%If session("OKG") <> "0" AND session("OKGref") <> "0" then%> checked<%end if%>> <%=gITEXT("i_43")%></td>	
	<td></td> 
</tr>
<%end if%>

    <% 

if session("template_org_code") = 1500 then%>
<tr>	
	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="Div"<%If session("OKG") <> "0" AND session("OKGref") <> "0" then%> checked<%end if%>> <%=gITEXT("i_51")%></td>	
	<td></td> 
</tr>
<%end if%>

<tr>
	<td></td>
</tr>

<%'05 JUL 06 REMOVED GENERAL COVERING LETTER 
'<tr>
'	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="H"> <%=gITEXT("i_30")></td>

'	<td>&nbsp;&nbsp;&nbsp;<input type='checkbox' name="app" value="X"> <%=gITEXT("i_29")></td>
'</tr>%>
<tr>
	<td>&nbsp;</td>
</tr>

<tr>
	<td colspan="2"><b><%=GITEXT("i_60")%></b> <select  name="usetoc">
	<option value="1" selected><%=GITEXT("i_39")%></option>
    <option value="0"><%=GITEXT("i_40")%></option>
    </select>
	</td>
</tr>
<%
'16 dec 06 ac - add pdf to show pic if candidate wannts
if request.querystring("goPDF") <> "" then%>
<tr>
	<td colspan="2"><b><%=GITEXT("i_10")%> ?</b> <select  name="usepix">
	<option value="0" selected><%=GITEXT("i_40")%></option>
    <option value="1"><%=GITEXT("i_39")%></option>
    </select>
	</td>
</tr>
<%end if%>
    <input type="hidden" name="useheader" value="0">
<tr>
	<td colspan="2"><b><%=gITEXT("i_61")%> </b> <select  name="format">
    <option value="A4" selected><%=gITEXT("i_62")%></option>
    <option value="Letter"><%=gITEXT("i_63")%></option>
    <option value="Universal"><%=gITEXT("i_64")%></option>
    </select></td>
</tr>
<tr>
    <td colspan="2">&nbsp;</td>
</tr>
<tr>
    <td align='center' colspan="2">
    <INPUT  TYPE="submit" value="    <%= GITEXT("i_46")%>   ">
    <input type="Reset" value="   <%= GITEXT("i_36")%>   ">
	</td>
</tr>
    </form>
<tr>
	<td align='center' colspan="2">&nbsp;</td>
</tr>
<tr>
	<td align='center' colspan="2"><font  color='Maroon'><% response.write GITEXT("i_37")%></font><a  href='http://www.adobe.com/products/acrobat/readstep.html'><img src="../../images/getacro.gif" border="0"></a></td>
</tr>
</table>
<%pv_last_update = "16 Dec 20"
'<<--Added By Interface on 04/23/2007
Set obj_int_select_Cmd = nothing
'-->> 
rsys_db_select.close
Set rsys_db_select = nothing 
rsys_int_select.close
Set rsys_int_select = nothing 
' *****************************************************
' BEGIN BOTTOM INCLUDES
' *****************************************************%>
<!--#include file="../includes/include_pubedit_frame_bottom.asp"-->
<% 
' *****************************************************
' END BOTTOM INCLUDES
' *****************************************************%>
