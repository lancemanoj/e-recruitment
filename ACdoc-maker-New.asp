<% Server.ScriptTimeout = 880%>
<% 
	Session.LCID = 9 'Language English 
	
	
'response.write "<br>THIS IS A TEST///"gITEXTV
	
'Response.ContentType = "text/html"
'Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
'Response.CodePage = 65001
'Response.CharSet = "UTF-8"
	
    dim JAPYsql, JAPY, JAPINFO2sql, JAPINFO2
	Dim pv_multi, pv_multicode, pv_pixok, pv_dater
    Dim pv_pic_path, pv_pic_temp_path, pv_pic_http_path, strError, src_pix_loc, des_pix_loc, pf
    Dim VN, VCDate, VacInfosql, VacInfo, objASPError, counter, pv_edulinecount, pv_new_sessioncode

    intHighNumber = 10000000
    intLowNumber = 1

sub StripSpecialChar(inFileName)
    Dim sOut, strorigFileName, arrSpecialChar, intCounter
    arrSpecialChar = Array("%20", "%", " ", "#", "+", "(", ")", "&", "$", "@", "!", "*", "<", ">", "?", "/", "|", "\", ",", "-")
    strorigFileName = inFileName
    intCounter = 0

    Do Until intCounter = 20
        sOut = replace(strorigFileName, arrSpecialChar(intCounter), "")
        intCounter = intCounter + 1
        strorigFileName = sOut
        'response.write strorigFileName
    Loop
    f.WriteLine(strorigFileName)
End Sub

'Response.ContentType = "text/html"
'Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
'Response.CodePage = 65001
'Response.CharSet = "UTF-8"

'response.write "<br>TEST FOREIGN CHARS"

' NOTES
' MAKE SURE THAT THE NAME IS in <h1></h1> brackets so that it will show as the chapter heading and header middle section of the PDF.
' <th colspan=""1""><font size=" & pv_THfontsize & "></th> and lower are to be used for the parts of the PHF for each applicant.


' 10 NOV 05 LJL revised to work with multiple candidates
'03 JAN 06 LJL worked on the output of additional docs, which was not right.
'07 LJL 06 LJL modified some bolding per IFRC, in RC section
'19 JAN 06 LJL many changes made for WTO format and also to clean up a bit
'30 Jan 06 RR Removed bold on perm. address and added driving license
'13 FEB 06 LJL/RR added SELECT statements for itext items, and session stop
'28 FEB 06 LJL added more checks for session vars and newlng vars   just before queries
'04 MAR 06 LJL string conversion wrong for if  JAPINFO1("cand_io_year_n") = "0" OR JAPINFO1("cand_io_year_n") = ""  then, was = 0
'05 MAR 06 LJL added Chinese and Russian for UNESCO 2500
'09 MAR 06 RR added Understanding for UNESCO 2500
'11 MAR 06 LJL added place and country to the education history lines per UNESCO
'12 MAR 06 LJL added fax to referal lines, per UNESCO
'25 MAR 06 LJL adjusted language level lines for all
'26 MAR 06 LJL revised formatting a bit to make tables fit better
'06 APR 06 LJL response buffer dying when trying to do 200 applicants, checking it out
'10 APR 06 LJL email address was not right
'10 APR 06 LJL corrected education section - if was looping for the highest edu degree and then the info below the edu lines, which it doesn't need to do, as those are single outputs, not loopable
'06 MAY 06 LJL changed template_org_code to pv_new_sessioncode
'11 MAY 06 LJL rsys_int to rsys_int_select
'17 JUN 06 LJL added if/then for crap MSIE browsers which cannot save a document that has its file name created on the fly. FFox can do this.
'09 JUL 06 LJL changed CV to Personal History as CV is not the term used by these organizations (WTO)
'10 AUG 06 RR added pv_new_sessioncode = 1500 for UNAIDS...
'28 AUG 06 LJL changed language level codes to +1
'17 SEP 06 LJL added rotmob to 1000 and 1500 WHO/UNAIDS
'19 OCT 06 LJL change : for UNAIDS and others
'23 OCT 06 LJL updated rotmob and punctuation, etc for UNAIDS
'31 Oct 06 ac added new pic location puller
'02 NOV 06 LJL adjusted internal section, additional information, tested internal for UNAIDS against my profile
'02 NOV 06 LJL adjusted pv_new_sessioncode to be pv_new_sessioncode in the picture finding code
'02 NOV 06 LJL changed the include text queries to only run if that section is selected - to optimize the page further
' ADDED ALSO the colors for the backgrounds of sections, per org, set in PH include text file
' 60 = section header color, 61 = section header font color, 62 = last update section color, 63 = last update font color
'02 NOV 06 LJL changed cand_law_m if empty to nothing written as output
'02 NOV 06 LJL adjusted to only get this org listing, for the post specific covering letter, set index in candtext table for this
'02 NOV 06 LJL added this to weed out per job covering letters if sent from admin/hrd-cllist to admin/docsettings.asp to /doccreate/make-doc-prep.asp
'05 Nov 06 ac to speed up page -  remove breaks bewtween html and asp for all "do while" and "if else" groups where possible (changes thru' out page
'07 Nov 06 ac- added new admin view thumb file for pic. pv_pic_path
'09 Nov 06 ac - open file, start writing
'21 Nov 06 ac - cleanup wrtielines.
'22 NOv 06 ac - open code up to others in stage.
'19 Dec 06 ac - added pv_new_sessioncode back as page is now integrated.
'20 DEC 06 LJL added header details and modified Section A, personal details, there are some extra </table and <tr codes that need to be taken out, by doing section by section
'20 DEC 06 LJL added this check on if it is truly a staff member for WHO and UNAIDS - stored proc dependent
'21 DEC 06 lJL removed line just under Areas of Expertise, per WTO request
'17 JAN 07 LJL int'l employ date wasn't set to right text var
'19 feb 07 ac - added new temp /public_hold when fixed changed back /public
'21 feb 07 ac - change back to public
'04 mar 07 - ac - get job title and VN number
'21 mar 07 - added line break cand_add_info; no change in gender required as both file were same
'29 MAR 07 LJl fixed intl employ first section of orgs
' 17 APR 07 LJL modif to properly assess if truly internal, mainly WHO  - thisorg_stafftrue is the staff number from WHO. If somethng there, then they are in WHO internal dir
'20 apr 07 ac - change CV file locations to match UNSHARE db (WHO, ILO, UNAIDS) and keep IFRC and WTO separate
'26 apr 07 ac - block new code revert to old file locs until duplicates removed
'29 APR 07 LJL corrected text for date of birth
'04 MAY 07 LJL moved language section to further up, per WHO
'09 MAY 07 LJL added the correct chapter headings by using <h1 for the name on the page, and then <TH for the parts of the PHF

'30 Nov 06 what type of file format does the user want?
'25 july 07 interface changed the  sql query into parameterized.
'18 AUG 07 LJL revised ONCE MORE the ridiculous WHO internal contract info and dates
'16 SEP 07 LJL updated lng, aoe, edu sections for langss
'27 SEp 07 LJL revised intl experience geo_exp_i which was missing i_7, had two i_6

'20 Nov 07 ac - add employment country location
'23 Nov 07 ac - only display country for IFRC
'25 Nov 07 ac - fix label for location of country as it uses rsys_int db and not IFRC prod db to i_88
'01 FEB 08 AC/DD photo view added.
'21 FEB 08 LJL internal cands for UNAIDS were showing wrong include text i_16 for nationality (staff nat)
'22 FEB 08 LJL centered the photo
'16 MAR 08 LJL format first page of multi PDF - show VN name and number
'27 MAR 09 DD  changes in additional information section for 'Other Information' related contents from WTO .
'03 APR 09 DD  added OTHER INFORMATION AND REFERENCE sections for WTO for this org.
'06 APR 09 DD  added Secretarial Skills for WTO org.
'17 JUN 09 LJL added to diminish VN specific covering letters output if admin is VN only
'20 JUL 09 LJL added language level names for each org
'17 SEP 09 LJL revised to not show this for any orgs that are not connected via true staff list (added 9999 as placeholder, stopper) code moves to ELSE
'07 OCT 09 DD Modified for Other Information section to contain the questions and answers even in the case of "No".
'14 OCT 09 DD Added the "if yes... " text part into the CV output
'28 OCT 09 DD Swapped Additional information to appear Relevant information first and includes more space to either side of "-" in Other documents section
'02 NOV 09 DD includes more space to either side of "-" in Other documents section
'08 NOV 09 LJL added Posts not to be enabled
'23 NOV 09 LJL corrected wrong label for ILO int't experience saying Country instead of From for header
'03 DEC 09 LJL adjusted employment section (Title) so that widths are more correctly spaced
'08 FEB 10 DD Blocked "In current country of residence" part for WTO by ANDing with 0
'20 AUG 10 DD Modified code to include country list of RC/RC experience if selected
'24 AUG 10 DD Modified code to put text 'Country list set to hidden' when include country list of RC/RC experience is not selected
'01 Sept 10 DD Added condition to avoid header (Additional Information) in geneated PDF if no data is available.
'01 SEP 10 LJL remove first horizontal line if first record
'08 SEP 10 LJL revised order of languages for ITU
'14 SEP 10 LJL removed contract details other than contract type, for ITU
'15 SEPT 10 DD Added new terms like pension of retiree and certification level of international experience and Clerical skills in PDf
'21 SEP 10 LJL changed to show the number for available levels in B
'05 OCT 10 LJL added for ITU, WIPO
'15 NOV 10 DD Now PDF contains Contract type only if applicant is staff member.
'10 DEC 10 LJL added new computer skills section, currently for 2400 and 2800 - ITU and WIPO
'11 DEC 10 LJL added ICDL for ITU
'24 JAN 11 LJL added new edu level names per org - set so that all are on the same level
'07 FEB 11 LJL added a space for before contractlendsc for WHO 1000 
'01 MAR 11  LJL added if/then to take the WHO GSM contract name and type over the user's entered contract.
'08 MAR 11 LJL added UN Typing test section for UN agencies
'27 MAR 11 LJL changed include text for A - previously applied, from being New Nationality include text
'11 MAY 11 LJL removed previously applied for WIPO
'11 MAY 11 LJL changed contract type output for internal staff to exclude for other orgs if from STAFF table join
'30 MAY 11 LJL changed candlng_un_typ_i to a string instead of int
'16 SEP 12 LJL PDF system reworked to ABCpdf
'19 SEP 12 LJL reworked for new PDF generation system.  Better TH, TD on sections.  
'04 JUL 13 added VnCloseDate and VNCDate
'08 NOV 14 LJL changed view to just country list for RotMob, added spaces before country answers in line
'11 DEC 14 GG cand image to center
'15 DEC 14 LJL added \docs\ to stage dirs
'30 JAN 15 LJL added UPU to image list
'04 FEB 15 LJL added coverter to make firstname of applicant without accents, per EBOLA and others
'10 FEB 15 LJL moved other text documents - some of the lines moved into the loop so that each text file has a separate header break and is in its own table element
'03 MAR 15 LJL added WTO, WMO, UNWomen to computer skills output
'20 MAY 15 LJL no contract start or end for WMO but show contract type
'22 MAY 15 GG set CAN hidden
'01 JUN 15 LJL added references as separate section header bar, WMO request				
'01 JUN 15 LJL made table 100% to fill page across for In city since, in country since
'02 JUN 15 GG move BODY  { background-color: #fff; } style only for html
'02 JUN 15 GG changed References section look like other sections
'02 JUN 15 GG added References section in GeneratePDF.aspx.cs
'03 JUN 15 GG changed headers color for Html version
'03 JUN 15 LJL no spanish as working lang for WMO
'06 JUN 15 GG Added Areas of expertise for WMO
'10 JUN 15 GG Changed i_39 to i_14 and moved here	
'11 JUN 15 GG Added countries for Geographical Experience
'11 JUN 15 GG Added EMPLOYMENT HISTORY Areas of Expertise
'09 JUL 15 GG Added French for Dates
'09 JUL 15 GG Fixed First 2 lines in international employment
'07 SEP 15 LJL added UPU to the new comp skills output
'16 FEB 16 LJL not sure why ILO 2000 was not to show dependants
'13 APR 16 GG added Linkedin Url
'25 APR 16 GG Added additional condition for Areas of Expertise
'30 MAY 16 GG Deleted additional condition for Areas of Expertise a.candccog_thisorg_2400 = 1 in many  rows this fiels is 0
'30 MAY 16 GG Deleted additional condition for Areas of Expertise a.candccog_thisorg_2400 = 1 
'17 JUN 16	GG Added colored line 
'10 07 17 GG fixed fr issue
'18 NOV 20 LJL removed linked in from PDF for UPU per them	
'16 DEC 20 LJL had to change encoding to ISO-8859-1 instead of UTF-8 for output to Word, PDF, HTML, not sure why, but if it works...
'16 DEC 20 LJL added goword check to stop the CAN (PDF chapter marker) from showing in word output 




if session("template_org_code") = "" then
	response.write "No session code was indicated or you have reached this page in error. <br>Please place your request once more or contact tech support if this error continues."
end if
if session("template_org_code") <> "" AND session("template_org_code") > 10 then
	pv_new_sessioncode = session("template_org_code")
else
	response.write "NO ORG CODE, ENDING"
	response.end
end if




PV_DEBUG = 0

Dim gITEXTPHsql, gITEXTPH, tempval1, obj_int_select_CmdII
Set obj_int_select_CmdII = Server.CreateObject("ADODB.Command")
obj_int_select_CmdII.ActiveConnection  = rsys_int_select


'<<--Modified by Interface on 07/25/2007
gITEXTPHsql = "	SELECT i_60, i_61, i_62, i_63 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'PH' "
obj_int_select_CmdII.CommandText = gITEXTPHsql
Set gITEXTPH = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
orgColor=gITEXTPH("i_60")
 
'-->>
' BEGIN INCLUDE TEXT DB
Dim gITEXTAsql, gITEXTA
'<<--Modified by Interface on 07/25/2007
gITEXTAsql = "SELECT i_1, i_2, i_3, i_8, i_4, i_9, i_10, i_5, i_22, i_23, i_27, i_28, i_29, i_89, i_30, i_38, i_48, i_7, i_6, i_74, i_19, i_72, i_11, i_16, i_14, i_15, i_27, i_26, i_40, i_81, i_28, i_21, i_67, i_68, i_69, i_70, i_73, i_77 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'A' "
obj_int_select_CmdII.CommandText = gITEXTAsql
Set gITEXTA = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
' END INCLUDE TEXT DB





' BEGIN CHECK IF PERSON SELECTED SECTION CONTACT
if instr(pv_parts,"W,") then
Dim gITEXTIsql, gITEXTI
'<<--Modified by Interface on 07/25/2007
	gITEXTIsql = "SELECT i_3, i_46, i_16, i_17, i_18, i_15, i_8, i_42, i_10, i_13, i_19 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'CONTACT' "
	obj_int_select_CmdII.CommandText = gITEXTIsql
	Set gITEXTI = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if

' BEGIN CHECK IF PERSON SELECTED SECTION B
if instr(pv_parts,"B,") then
Dim gITEXTBsql, gITEXTB
'<<--Modified by Interface on 07/25/2007
	gITEXTBsql = "SELECT i_1, i_4, i_8, i_10, i_17, i_21 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'B' "
	obj_int_select_CmdII.CommandText = gITEXTBsql
	Set gITEXTB = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if

' BEGIN CHECK IF PERSON SELECTED SECTION C
if instr(pv_parts,"C,") then
Dim gITEXTCsql, gITEXTC
'<<--Modified by Interface on 07/25/2007
	gITEXTCsql = "SELECT  i_1, i_3, i_21, i_5, i_6, i_7, i_10, i_11, i_12, i_13, i_23, i_14, i_15, i_16, i_17, i_35, i_36, i_37 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'C' "
	obj_int_select_CmdII.CommandText = gITEXTCsql
	Set gITEXTC = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if

' BEGIN CHECK IF PERSON SELECTED SECTION D
if instr(pv_parts,"D,") then
Dim gITEXTDsql, gITEXTD
'<<--Modified by Interface on 07/25/2007
	gITEXTDsql = "SELECT i_1, i_5, i_6, i_7, i_16, i_17, i_18, i_32, i_10, i_9, i_29, i_19, i_11, i_20, i_22, i_21, i_40, i_41, i_42, i_43, i_44 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'D' "
	obj_int_select_CmdII.CommandText = gITEXTDsql
	Set gITEXTD = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if

' BEGIN CHECK IF PERSON SELECTED SECTION E
if instr(pv_parts,"E,") then
Dim gITEXTEsql, gITEXTE
'<<--Modified by Interface on 07/25/2007
	gITEXTEsql = "SELECT i_1, i_37, i_38, i_23, i_8, i_9, i_10, i_39, i_26,i_82,i_22,i_84,i_14 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'E' "
	obj_int_select_CmdII.CommandText = gITEXTEsql
	Set gITEXTE = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if

' BEGIN CHECK IF PERSON SELECTED SECTION F
if instr(pv_parts,"F,") then
Dim gITEXTFsql, gITEXTF
'<<--Modified by Interface on 07/25/2007
	gITEXTFsql = "SELECT i_3, i_5, i_8, i_9, i_55, i_10, i_24, i_25, i_26, i_27, i_19, i_20, i_53, i_60, i_63, i_90, i_56, i_62, i_28, i_29, i_89, i_30, i_68, i_73, i_82, i_74, i_3, i_62, i_4, i_7, i_6, i_75, i_11, i_21, i_59, i_12, i_13, i_88, i_17, i_18, i_87, i_69, i_45, i_93, i_text4 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'F' "
	obj_int_select_CmdII.CommandText = gITEXTFsql
	Set gITEXTF = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if

' BEGIN CHECK IF PERSON SELECTED SECTION G
if instr(pv_parts,"G,") then
Dim gITEXTGsql, gITEXTG
'<<--Modified by Interface on 07/25/2007
	gITEXTGsql = "	SELECT i_21, i_22, i_23, i_1, i_3, i_28, i_5, i_54, i_7, i_55, i_29, i_12, i_15, i_30, i_17, i_18, i_19, i_75, i_68, i_24, i_27, i_80 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'G' "
	obj_int_select_CmdII.CommandText = gITEXTGsql
	Set gITEXTG = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
end if

' BEGIN INCLUDE TEXT DB
Dim gITEXTHsql, gITEXTH
'<<--Modified by Interface on 07/25/2007
gITEXTHsql = " SELECT i_4, i_11, i_5 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'H' "
obj_int_select_CmdII.CommandText = gITEXTHsql
Set gITEXTH = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
' END INCLUDE TEXT DB

' BEGIN CHECK IF PERSON SELECTED SECTION S
if instr(pv_parts,"S,") then
Dim gITEXTSsql, gITEXTS
'<<--Modified by Interface on 07/25/2007
	gITEXTSsql = " SELECT i_1, i_42, i_43, i_44, i_45, i_46, i_22, i_23, i_24, i_25, i_26, i_27, i_28, i_29, i_30 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'S' "
	obj_int_select_CmdII.CommandText = gITEXTSsql
	Set gITEXTS = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if



' BEGIN CHECK IF PERSON SELECTED SECTION J
if instr(pv_parts,"J,") then
Dim gITEXTJsql, gITEXTJ
'<<--Modified by Interface on 07/25/2007
'13 APR 16 GG added Linkedin Url Text
	gITEXTJsql = " SELECT i_1, i_3, i_12, i_13, i_71 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'J' "
	obj_int_select_CmdII.CommandText = gITEXTJsql
	Set gITEXTJ = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
'<<--Modified by DD on 03/27/2009
Dim gITEXTJFsql, gITEXTJF
Dim gITEXTJVsql, gITEXTJV
    if   pv_new_sessioncode = 3000 then
	    gITEXTJFsql = "SELECT i_text2  FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'F' "
	    obj_int_select_CmdII.CommandText = gITEXTJFsql
	    Set gITEXTJF = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))

	    gITEXTJVsql = "SELECT i_11 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'V'"
	    obj_int_select_CmdII.CommandText = gITEXTJVsql
	    Set gITEXTJV = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
	end if
'-->>
end if

' BEGIN CHECK IF PERSON SELECTED SECTION Other information
'14 OCT 09 DD Added the "if yes... " text part into the CV output
if instr(pv_parts,"OI,") then
Dim gITEXTJFsql2, gITEXTJF2
Dim gITEXTJVsql2, gITEXTJV2
    gITEXTJFsql2 = "SELECT i_text2, i_87  FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'F' "
    obj_int_select_CmdII.CommandText = gITEXTJFsql2
    Set gITEXTJF2 = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))

    gITEXTJVsql2 = "SELECT i_11, i_12 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'V'"
    obj_int_select_CmdII.CommandText = gITEXTJVsql2
    Set gITEXTJV2 = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
end if

' BEGIN CHECK IF PERSON SELECTED SECTION References
if instr(pv_parts,"GR,") then
Dim gITEXTGRefsql, gITEXTGRef
    gITEXTGRefsql = "SELECT i_15,i_30,i_5,  i_17, i_18, i_19, i_57, i_27,i_80,i_75,i_19  FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'G' "
    obj_int_select_CmdII.CommandText = gITEXTGRefsql
    set gITEXTGRef = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
end if

' BEGIN CHECK IF PERSON SELECTED SECTION V
if instr(pv_parts,"V,") then
Dim gITEXTVsql, gITEXTV
'<<--Modified by Interface on 07/25/2007
	gITEXTVsql = "SELECT i_1, i_text1, i_3, i_4, i_10,i_11,i_18,i_19,i_94,i_95,i_96,i_97,i_98 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'V' "
	obj_int_select_CmdII.CommandText = gITEXTVsql
	Set gITEXTV = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if

if pv_new_sessioncode = 7000 then
' BEGIN CHECK IF PERSON SELECTED SECTION RC
	if instr(pv_parts,"RC,") then
Dim gITEXTRCsql, gITEXTRC
		gITEXTRCsql = "		SELECT i_1, i_2, i_7, i_6, i_3, i_4, i_8, i_21, i_9, i_10, i_23, i_12, i_13, i_14, i_11, i_15, i_16, i_17, i_19, i_20, i_25 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'RC' "
' response.write gITEXTsql
		obj_int_select_CmdII.CommandText =gITEXTRCsql
		set gITEXTRC = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
' END INCLUDE TEXT DB
	end if
end if

if instr(pv_parts,"RM,") then
	if pv_new_sessioncode = 1000 OR pv_new_sessioncode = 1200  then ' OR pv_new_sessioncode = 1500 by Gevorg
' BEGIN INCLUDE TEXT DB
Dim gITEXTRMsql, gITEXTRM
		gITEXTRMsql = "	 SELECT i_1, i_3, i_4, i_5, i_6, i_7, i_8, i_9, i_11, i_14, i_15, i_16, i_17, i_20, i_21 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'Rmob' "
		obj_int_select_CmdII.CommandText = gITEXTRMsql
		set gITEXTRM = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
' END INCLUDE TEXT DB
	end if
end if


' BEGIN CHECK IF PERSON SELECTED SECTION V
if instr(pv_parts,"T,") then
Dim gITEXTTsql, gITEXTT
'<<--Modified by Interface on 07/25/2007
	gITEXTTsql = "SELECT i_1, i_4, i_5, i_6, i_7, i_8, i_9, i_3, i_10, i_11, i_12, i_15, i_13,i_14,i_16,i_17, i_18, i_40, i_41, i_42, i_21 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'T' "
	obj_int_select_CmdII.CommandText = gITEXTTsql
	Set gITEXTT = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))
'-->>
end if


if pv_doc_type = "pdf" then
	widther = "100%"
elseif pv_doc_type = "doc" then
	widther = "100%"
elseif pv_doc_type = "html" then
	widther = "100%"
else
	widther = "100%"
end if

Dim pv_THFontsize
pv_THfontsize = "10pt"

f.WriteLine ("<html>" )

'16 DEC 20 LJL had to change encoding to ISO-8859-1 instead of UTF-8 for output to Word, PDF, HTML, not sure why, but if it works...
'f.WriteLine "<head>"
f.WriteLine "<head><meta http-equiv='Content-Type' content='text/html; charset=ISO-8859-1'>"

'03 JUN 15 GG changed headers color for Html version
Dim g_headerColor
g_headerColor = gITEXTPH("i_61")

if pv_doc_type = "html" then
	f.WriteLine "<link href=""https://erecruit.itu.int/css/2900-css/admin-2900.css"" type=""text/css"" rel=""STYLESHEET"">"
	f.WriteLine "<style>" '02 JUN 15 GG move BODY  { background-color: #fff; } style only for html
	f.WriteLine "BODY  { background-color: #fff; }"
	f.WriteLine "</style>"
	g_headerColor = orgColor
end if

f.WriteLine "<style>"
'f.WriteLine "BODY  { FONT-FAMILY: Arial; FONT-SIZE: 9pt; }"
'f.WriteLine  "H1   { FONT-SIZE: 12pt; }"

'f.WriteLine  "TH   { text-align: left; background-color:" & gITEXTPH("i_60") & "; font-size:10px; color:" & g_headerColor & "; height:15px;} "

'f.WriteLine "TABLE { PADDING: 2px; }"
'f.WriteLine "TD    { FONT-SIZE: 9pt;  }"
f.WriteLine "</style>"

'04 mar 07 - ac - get job title and VN number
VN = " "
VCDate = " "
if len(pv_jobid) then
	VacInfosql = "SELECT jobinfo_job_en_t, jobinfo_vac2_c, jobinfo_acl_d FROM v_rsys_job_stat_simple WHERE jobinfo_uid_c = " & pv_jobid & " AND jobinfo_thisorg_" & pv_new_sessioncode & " = 1"
	Set VacInfo = Server.CreateObject("ADODB.RecordSet")
	VacInfo.Open VacInfosql, rsys_db_select, 1, 1
	if  VacInfo.eof = false then
		VN = VacInfo("jobinfo_job_en_t") & " - " &  VacInfo("jobinfo_vac2_c")
		VCDate = VacInfo("jobinfo_acl_d")
	else
		VN = " "
		VCDate = " "
	end if
end if
if pv_DEBUG then
	response.write "<br>IN DOCMAKER pv_jobid:" & pv_jobid
	response.write "<br>IN DOCMAKER VN:" & VacInfo("jobinfo_job_en_t")
	response.write "<br>IN DOCMAKER VCDate:" & VacInfo("jobinfo_acl_d")
'response.End()
end if

if multi <> "" then

'f.WriteLine "<title>" & VN & " - " & day(now()) & " " & monthname(month(now())) & " " & year(now()) & "</title>"
Vnname=VN
'04 JUL 13 added VnCloseDate and VNCDate
VnClosing=VCDate
'19 DEC 06 LJL moved to within the query and above section A - to put <h1> as the applicant name for headers in the PDF
'	f.WriteLine "<h1>Multiple applicant document</h1>"
	pv_multi = "1"
	if multicode <> "" then
		pv_multicode = multicode
	elseif multi <> "" then
		pv_multicode = multi
	end if
elseif(pv_doc_type <> "pdf") Then
	f.WriteLine "<title>Personal History - " & day(now()) & " " & monthname(month(now())) & " " & year(now()) & "</title>"
'19 DEC 06 LJL moved to within the query and above section A - to put <h1> as the applicant name for headers in the PDF
'	f.WriteLine "<h1>Personal History</h1>"
	pv_multi = "0"
end if

'19 DEC 06 LJL moved to within the query and above section A - to put <h1> as the applicant name for headers in the PDF
f.WriteLine "</head>"
f.WriteLine "<body>"


' TO CHECK THE VARS SENT TO THIS PAGE FOR SECTIONS TO SHOW
'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
if pv_DEBUG then
	response.write "<br>TEST-APP:" & request.form("app")
'response.End()
end if

' PAGE IS SAME IN HTDOCS/DOCMAKER, PUBLIC/PDF/, E_ASP/DOCMAKER
if pv_DEBUG then
	response.write "<br>T START: " & request.form("app")
	response.write "<br>stop 3.0"
'response.End()
end if


Dim new_lng_code, widther
Dim pv_newlevel, pv_loc
Dim JAPEMPsql, JAPEMP, JAPEMP0sql, JAPEMP0, pv_pagecount, JAPINFOIEsql, JAPINFOIE, GETMISCsql, GETMISC
Dim JAPINFORCsql, JAPINFORC, JAPOFFICE1sql, JAPOFFICE1, JAPINSTLIST11sql, JAPINSTLIST11, JAPINSTLIST12sql, JAPINSTLIST12, JAPGEOLIST21sql, JAPGEOLIST21
Dim JAPGEOLISTsql, JAPGEOLIST, JAPGEOLIST31sql, JAPGEOLIST31, JAPGEOLIST41sql, JAPGEOLIST41, JAPGEOLIST42sql, JAPGEOLIST42
Dim JAPGEOLIST90sql, JAPGEOLIST90, JAPGEOLIST61sql, JAPGEOLIST61, JAPGEOLIST62sql, JAPGEOLIST62, GETDROP2sql, GETDROP2, JAPMOBsql, JAPMOB
Dim GetEmploymentCountrySql, GetEmploymentCountry, JAPCOMPSsql, JAPCOMPS


' REMOVE AFTER TESTING
if request.querystring("lng") <>"" then
	new_lng_code = request.querystring("lng")
elseif request.form("lng") <>"" then
	new_lng_code = request.form("lng")
elseif session("lng") <> "" then
	new_lng_code = session("lng")	
else
	new_lng_code = "en"
end if

'09 JUL 15 GG Added French for Dates
if new_lng_code = "fr" then Session.LCID = 4108 'French - Switzerland

'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
if pv_DEBUG then
	response.write "<br>acdoc-maker-NEW.asp"
	response.write "<br> HTTP_REFERER = " & request.servervariables("HTTP_REFERER")
	response.write "<br> REMOTE_ADDR= " & request.servervariables("REMOTE_ADDR")
	For Each Item In Request.Form
		response.write "<br>rf. " & Item &  " = " &  Request.Form(Item)
	Next
	Response.Write"<br>u. " & Request.QueryString
	response.write "<br>stop 3.2"
'response.End()
end if

if instr(request.servervariables("HTTP_REFERER"),"/admin/docsettings.asp") then
	if (session("template_org_code") = 2600 OR session("template_org_code") = 2900) and session("adl_ok") = "1" then		
		'response.write "<br><Br>2900 ENTRY"
		
					applicant_id = ""
			if ids <> "" then
'widther = "650"
				applicant_id = ids
			elseif multi <> "" AND multicode <> "" then
'widther = "650"
				applicant_id = ""
			else
				response.write "ERR 909.2"
				response.end
			end if


	else
		If session("adl_ok") = "1" AND request.Cookies("aslogged") = "698" Then
'widther = "650"
			applicant_id = ""
			if ids <> "" then
'widther = "650"
				applicant_id = ids
			elseif multi <> "" AND multicode <> "" then
'widther = "650"
				applicant_id = ""
			else
				response.write "ERR 909.2"
				response.end
			end if

		else
			response.redirect "../login/index.asp"
		end if
	end if
'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
		if pv_DEBUG then
			response.write "<br>FROM: ADM docsettings PDF"
			response.write "<br>widther = " & widther
'response.End()
		end if


'19 feb 07 ac - added new temp /public_hold when fixed changed back /public
elseif instr(request.servervariables("HTTP_REFERER"),"/public/pdf/docsettings.asp") then
' CHECK IF LOGGED IN PUBLIC (copied from /public/include/include_check_login.asp)
		if pv_new_sessioncode = "" or session("RSYS_EVAL") = "" OR request.Cookies("curruser") = "" then
			response.redirect "../public/ejobs-login.asp"
		else
			widther = "100%"
'applicant_id = ids done in make_doc_prepNew
'cstr(session("RSYS_EVAL"))
			if pv_DEBUG then
				response.write "<br>applicant_id:=["
				 response.write  applicant_id
				 For each Item in Session.Contents
					Response.Write("<Br>Item=" & Item)
					Response.Write( " : Session value=" & Session.Contents(item) )
				next
				response.write "<br>stop 3.top"
'response.End()
			end if
'ac blocked already done pv_parts = replace(request.form("app")," ","")
			pv_loc = "../../"
		end if
'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
		if pv_DEBUG then
			response.write "<br> PUB do word or HTML"
		end if
else
response.redirect "../public/ejobs-login.asp"
end if

' INFO -- if no applicant code(s), then don't allow further
if pv_multicode <> "" OR applicant_id <> "" then

else
	response.write "<Br><font color='maroon'>IMPROPER ENTRY - NO CAND ID, please report how you entered this page to tech support <a href='mailto:contact@erecruithelp.org'>Tech Support</a></font>"
	response.end
end if

' SET ALL VARS AT TOP OF PAGE IF CAN
Dim japinfosql1
Dim JAPINFO1
Dim gAOEsql
Dim gAOE
Dim JAPEDUsql, JAPEDU, JAPREFsql, JAPREF, JAPDEPsql, JAPDEP, JAPEDULINESsql, JAPEDULINES, INTCURRENTsql, INTCURRENT, JAPRELsql, JAPREL, gCandsql, gCand
Dim GETFACTORSsql, GETFACTORS, GetDocsql, GetDoc, GetDoc2sql, GetDoc2, GetDoc3sql, GetDoc3, JAPPREFLIST, JAPPREFLISTsql
Dim pdf_lng, covlettype, covlettype2

' SET JOB ID IF SENT TO THIS PAGE
' ac blocked 24th nov - moved to ACmake-doc-prep.asp
if len(request.querystring("jobinfo_uid_c")) then
	vacchoice = request.querystring("jobinfo_uid_c")
else
	vacchoice = "0"
end if
' response.write "BEGIN 2"

' TEMPORARY TAKE OUT VARIABLES
'new_lng_code = "en"
pdf_lng = new_lng_code

' BEGIN CHECK IF PERSON SELECTED SECTION C
if instr(pv_parts,"C,") then

Dim GETLEVEL1sql, GETLEVEL2sql, GETLEVEL3sql, GETLEVEL4sql, GETLEVEL5sql, GETLEVEL6sql
Dim GETLEVEL1, GETLEVEL2, GETLEVEL3, GETLEVEL4, GETLEVEL5, GETLEVEL6
Dim pv_level1, pv_level2, pv_level3, pv_level4, pv_level5, pv_level6

GETLEVEL1sql = " SELECT lngl_dsc_"& new_lng_code & "_t_" & pv_new_sessioncode & " as LEVELDSC FROM core_lngltblf WHERE lngl_id_c = 1 AND lngl_thisorg_" & pv_new_sessioncode & " = 1 "
''d set GETLEVEL1 =rsys_db_select.execute(GETLEVEL1sql)
Set GETLEVEL1 = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
rsys_db_select.CommandTimeout = 320
GETLEVEL1.Open GETLEVEL1sql, rsys_db_select, 1, 1
if GETLEVEL1.eof = false then
pv_level1 = GETLEVEL1("leveldsc")
else
pv_level1 = ""
end if

GETLEVEL2sql = " SELECT lngl_dsc_"& new_lng_code & "_t_" & pv_new_sessioncode & " as LEVELDSC FROM core_lngltblf WHERE lngl_id_c = 2 AND lngl_thisorg_" & pv_new_sessioncode & " = 1 "
''d set GETLEVEL2 =rsys_db_select.execute(GETLEVEL2sql)
Set GETLEVEL2 = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
rsys_db_select.CommandTimeout = 320
GETLEVEL2.Open GETLEVEL2sql, rsys_db_select, 1, 1
if GETLEVEL2.eof = false then
pv_level2 = GETLEVEL2("leveldsc")
else
pv_level2 = ""
end if

GETLEVEL3sql = " SELECT lngl_dsc_"& new_lng_code & "_t_" & pv_new_sessioncode & " as LEVELDSC FROM core_lngltblf WHERE lngl_id_c = 3 AND lngl_thisorg_" & pv_new_sessioncode & " = 1 "
''d set GETLEVEL3 =rsys_db_select.execute(GETLEVEL3sql)
Set GETLEVEL3 = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
rsys_db_select.CommandTimeout = 320
GETLEVEL3.Open GETLEVEL3sql, rsys_db_select, 1, 1
if GETLEVEL3.eof = false then
pv_level3 = GETLEVEL3("leveldsc")
else
pv_level3 = ""
end if

GETLEVEL4sql = " SELECT lngl_dsc_"& new_lng_code & "_t_" & pv_new_sessioncode & " as LEVELDSC FROM core_lngltblf WHERE lngl_id_c = 4 AND lngl_thisorg_" & pv_new_sessioncode & " = 1 "
''d set GETLEVEL4 =rsys_db_select.execute(GETLEVEL4sql)
Set GETLEVEL4 = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
rsys_db_select.CommandTimeout = 320
GETLEVEL4.Open GETLEVEL4sql, rsys_db_select, 1, 1
if GETLEVEL4.eof = false then
pv_level4 = GETLEVEL4("leveldsc")
else
pv_level4 = ""
end if

GETLEVEL5sql = " SELECT lngl_dsc_"& new_lng_code & "_t_" & pv_new_sessioncode & " as LEVELDSC FROM core_lngltblf WHERE lngl_id_c = 5 AND lngl_thisorg_" & pv_new_sessioncode & " = 1 "
''d set GETLEVEL5 =rsys_db_select.execute(GETLEVEL5sql)
Set GETLEVEL5 = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
rsys_db_select.CommandTimeout = 320
GETLEVEL5.Open GETLEVEL5sql, rsys_db_select, 1, 1
if GETLEVEL5.eof = false then
pv_level5 = GETLEVEL5("leveldsc")
else
pv_level5 = ""
end if

GETLEVEL6sql = " SELECT lngl_dsc_"& new_lng_code & "_t_" & pv_new_sessioncode & " as LEVELDSC FROM core_lngltblf WHERE lngl_id_c = 6 AND lngl_thisorg_" & pv_new_sessioncode & " = 1 "
''d set GETLEVEL6 =rsys_db_select.execute(GETLEVEL6sql)
Set GETLEVEL6 = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
rsys_db_select.CommandTimeout = 320
GETLEVEL6.Open GETLEVEL6sql, rsys_db_select, 1, 1
if GETLEVEL6.eof = false then
pv_level6 = GETLEVEL6("leveldsc")
else
pv_level6 = ""
end if

' END CHECK IF PERSON SELECTED SECTION C
end if

' response.write ""
if pv_DEBUG then
	response.write "<br>T START2: "
	response.write "<br>stop 3.34"
'response.End()
end if
if new_lng_code = "" OR pv_new_sessioncode = "" then
	response.write "Cannot proceed.  Missing instruction (error: 111)" & "new_lng_code=" & new_lng_code & "Org. Code=" & pv_new_sessioncode
	response.end
end if

' -------------------------------------------------QUERY APPLICANT LIST -------------------------------------------------------------------------------

'japinfosql1 = "	 	{call erstp_rsys_cand_app_full_view_en_" & pv_new_sessioncode & "(" & applicant_id & ")} "
if pv_multi = "1" then
' INFO -- multicode comes across as VARCHAR from the querystring, in order to get the results of the storage table of cand_id's
	japinfosql1 = " {call erstp_rsys_cand_app_full_view_multi_" & new_lng_code & "_" & pv_new_sessioncode & "('" & pv_multicode & "')} "
'set JAPINFO1 =rsys_db_select.execute(japinfosql1)
'set JAPINFO =rsys_db_select.execute(JAPINFOsql)
'AC- changed how the connection was open
	Set JAPINFO1 = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
	rsys_db_select.CommandTimeout = 320
	JAPINFO1.Open japinfosql1, rsys_db_select, 0, 1
'adOpenForwardOnly,adLockReadOnly,adCmdText
'response.write "TESTOUT: MULTI"
else
	japinfosql1 = " {call erstp_rsys_cand_app_full_view_" & new_lng_code & "_" & pv_new_sessioncode & "(" & applicant_id & ")} "

'set JAPINFO1 =rsys_db_select.execute(japinfosql1)
'set JAPINFO =rsys_db_select.execute(japinfosql1)
'AC- changed how the connection was open
	Set JAPINFO1 = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
	rsys_db_select.CommandTimeout = 320
	JAPINFO1.Open japinfosql1, rsys_db_select, 0, 1
'response.write "TESTOUT: CCC"
end if

' -------------------------------------------------QUERY APPLICANT LIST -------------------------------------------------------------------------------

if pv_DEBUG then
	for each x in JAPINFO1.fields
	   response.write(x.name)
	   response.write(" = ")
	   response.write(x.value)
	next
	response.write "<br>T japinfosql1: " & japinfosql1
	response.write "<Br>JAPINFO1.RecordCount: DDD"
	response.write JAPINFO1.RecordCount
	response.write "<br>JAPINFO1->cand_id_c:="
			response.write JAPINFO1("cand_id_c")
	response.write "<br>stop 3.36"
	'response.End()
end if

'***************************jap detal for all org by manoj
       '<<--Modified by Interface on 05/01/2007
JAPYsql = "SELECT upd_d, cand_fnam_t, cand_lnam_t, cand_fil_t, cand_fil_d, cand_verif_name_t, cand_law_i, cand_law_m," & _
          " cand_dismissed_i, cand_dismissed_m, cand_resigned_i, cand_resigned_m, cand_nameinclude_i, cand_nameinclude_UN," & _
          " cand_sexual_i, cand_sexual_m, cand_teriminated_i, cand_terminated_m" & _
          " FROM td_rsys_cand WHERE cand_id_c = " & applicant_id

Set JAPY = Server.CreateObject("ADODB.RecordSet")
rsys_db_select.CommandTimeout = 320
JAPY.Open JAPYsql, rsys_db_select, 0, 1
  '-->>


'********************
' ******************************************************************************
' CONTENT TYPE TO WORD IF DESIRED
'if request.form("goWord") = "99" then
'	widther = "100%"
'	Response.Buffer = True
'	Response.ContentType = "application/vnd.ms-word"
'	' Adds a header to give the document a name
'	pv_dater = day(now()) & "_" & monthname(month(now()),2) & year(now()) & formatdatetime(now(), 4)
''	'17 JUN 06 LJL had to adjust for crap IE browser not able to save the file if title generated on the fly
'if instr(request.servervariables("HTTP_USER_AGENT"),"MSIE") then
'	response.AddHeader "content-disposition", "inline; filename=MyPH.doc"
'	else
'		response.AddHeader "content-disposition", "inline; filename=PH_" & pv_dater & ".doc"
'	end if
'end if


''-----------------------------------------NOTES ----------
'MODS --
'09 MAR 05 LJL removed UN test for WTO
'RR - LISTGETAT's
'--------------------------------------------------------------->
''-----------------------------------------NOTES ----------
'Computer skills PDF preparation page
'MODS --
'--------------------------------------------------------------->

' ******************************************************************************
' BEGIN OUTPUT FROM JAPINFO
' ******************************************************************************
Dim liner, GETPICsql, GETPIC
liner = 1
Do Until JAPINFO1.eof
' releases output as is, so every 50 records it pushes to browser.
		if (liner mod 2) = 0 then
			response.write "<progress value='"& liner & "' max=""" & candsCount & """></progress>" 'records processed..."
' increase script timeout
			response.flush
			Server.ScriptTimeout = 880
		end if

		if pv_DEBUG then
			response.write "<br>JAPINFO1->cand_id_c:="
			response.write JAPINFO1("cand_id_c")
			response.write "<br>stop 3.37"
'response.End()
		end if

		  applicant_id = JAPINFO1("cand_id_c")
		  liner = liner + 1


' GET PICTURE IF REQUIRED
		if pv_usepix = "1" then
			GETPICsql = " SELECT TOP 1 upl_snam_t	from td_rsys_upl where uplcategory_id_c = 4 AND cand_id_c = " & applicant_id
''d set GETPIC =rsys_db_select.execute(GETPICsql)
			Set GETPIC = Server.CreateObject("ADODB.RecordSet")
			GETPIC.Open GETPICsql, rsys_db_select, 0, 1
			if GETPIC.eof or GETPIC.bof then
				pv_pixok = "0"
			else
				pv_pixok = "1"
			end if
		else
			pv_pixok = "9"
		end if

	Randomize
    intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)

	'f.WriteLine "<STYLE TYPE='text/css'>' BODY   {font-family:sans-serif; font-size:10px;} --> </STYLE>"
	'f.WriteLine "body {font-family:arial; font-size:10px};"
	f.WriteLine "<!-- FOOTER CENTER """ & VN & """ -->"
	f.WriteLine "<br>"

	f.WriteLine "<h3 style=""text-align: center;"">" 
	Randomize
    intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)

'16 DEC 20 LJL added goword check to stop the CAN (PDF chapter marker) from showing in word output 
if request.form("goWord") = "99" then
	f.WriteLine "<a style=""abcpdf-tag-visible: true;display:none;"" id='" & trim(Ucase(JAPINFO1("cand_lnam_t"))) & ", " & JAPINFO1("cand_fnam_t")  & "_|"  & intNumber  & "_h1'></a>" 
else
f.WriteLine "<a style=""abcpdf-tag-visible: true;display:none;"" id='" & trim(Ucase(JAPINFO1("cand_lnam_t"))) & ", " & JAPINFO1("cand_fnam_t")  & "_|"  & intNumber  & "_h1'>CAN</a>" 
end if

	f.WriteLine trim(Ucase(JAPINFO1("cand_lnam_t"))) & ", " & JAPINFO1("cand_fnam_t")
	if multi = "" then
		'04 FEB 15 LJL added coverter to make firstname of applicant without accents, per EBOLA and others
		dim pv_newFname, pv_newLname, obj1, obj2

		Set obj1 = Server.CreateObject("ADODB.Stream")
		obj1.Charset = "ascii"
		obj1.Open
		obj1.WriteText JAPINFO1("cand_fnam_t")
		obj1.Position = 0
		pv_newFname = obj1.ReadText

		obj1.Close
		Set obj1 = Nothing

		Set obj2 = Server.CreateObject("ADODB.Stream")
		obj2.Charset = "ascii"
		obj2.Open
		obj2.WriteText JAPINFO1("cand_lnam_t")
		obj2.Position = 0
		pv_newLname = obj2.ReadText

		obj2.Close
		Set obj2 = Nothing

		ApplicantName = trim(Ucase(pv_newLname)) & "_" & pv_newFname
	end if
	f.WriteLine "</h3>"


' *******************************************************************************************************************
' PART A PERSONAL DETAILS
' *******************************************************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION A
if instr(pv_parts,"A,") then
	f.WriteLine "<br/>"
	f.WriteLine "<p><br></p>"
	f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTA("i_1") & "</font></h2>"	
	f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width='" &  widther & "' align=""center"">"
	if pv_DEBUG then
		response.write "<br>PV_PIXOK:= " & pv_pixok
	end if

	if pv_pixok = "1" then
	dir_current = UCASE(request.servervariables("PATH_TRANSLATED"))

	' if "stage" found in path then file is in developemnt stage
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		src_pix_loc = "E:\docs\stage\ALLORG\vac-cv\"
				
				
' production site storage file locations
'05 OCT 10 LJL added for ITU, WIPO
'19 DEC 14 LJL added WMO
'30 JAN 15 LJL added UPU to image list
				elseif pv_new_sessioncode = 1000 or pv_new_sessioncode = 1500 or pv_new_sessioncode = 2000 or pv_new_sessioncode = 2400 or pv_new_sessioncode = 2800 or pv_new_sessioncode = 2900 or pv_new_sessioncode = 5500 then
' ' new folder location WHO, UNAIDS & ILO (Production sites)
				src_pix_loc = "E:\docs\UNSHARE\vac-cv\"
			elseif Instr(1, dir_current, "WTO", 1) > 0 then
' ' new folder location WTO (Production site)
				src_pix_loc = "E:\docs\WTO\vac-cv\"
			elseif Instr(1, dir_current, "IFRC", 1) > 0 then
' ' new folder location IFRC (Production site)
				src_pix_loc = "E:\docs\IFRC\vac-cv\"
			end if

' set complate path and file for source of pix
			src_pix_loc = src_pix_loc & GETPIC("upl_snam_t")

'******************************************
'COPY image file to pdf/html/word location
'******************************************
			if pv_DEBUG then
				response.write "<br>S:= " & src_pix_loc
				response.write "<br>D:= " & dest_dir
			end if
			Set pf=Server.CreateObject("Scripting.FileSystemObject")
			if pf.FileExists(src_pix_loc) then
				pf.CopyFile src_pix_loc, dest_dir
				f.WriteLine "<tr>"
				f.WriteLine	"<td colspan=""2"" valign=""top"" align=""center"">"
				f.WriteLine "<table><tr><td></td><td>"
				f.WriteLine "<img src=""" & GETPIC("upl_snam_t") & """ border=0 alt="""" width=""150"">"
				f.WriteLine "</td><td></td></tr></table>"
				f.WriteLine "</td>"
				f.WriteLine"</tr>"
			else
				f.WriteLine "<tr>"
				f.WriteLine	"<td colspan=""2"" valign=""top"" align=""right"">"
				f.WriteLine "<br>2: File does not exist"
				f.WriteLine "</td>"
				f.WriteLine"</tr>"
			end if
			set pf = nothing
	end if
'f.WriteLine "<img src=""../css/2900-css/vn_logo.jpg"" width=""15"" height=""15"" border=""0"" >"
' INTERNAL AND WHO 1000
'20 DEC 06 LJL changed this to be veritable WHO or UNAIDS
if false then '(pv_new_sessioncode = 10000) OR pv_new_sessioncode = 1500 by Gevorg
' 17 APR 07 LJL modif to properly assess if truly internal, mainly WHO  - thisorg_stafftrue is the staff number from WHO. If somethng there, then they are in WHO internal dir
	if len(JAPINFO1("thisorg_stafftrue")) then
		f.WriteLine "<tr><td valign=""top"" colspan=""2"" align=""center""><strong>" & JAPINFO1("salutation") & " " & JAPINFO1("firstname") & " " & Ucase(JAPINFO1("lastname")) & "</strong></td></tr>"
		if len(JAPINFO1("cand_mnam_t")) OR len(JAPINFO1("cand_onam_t")) then
			f.WriteLine "<tr>"
			if len(JAPINFO1("cand_mnam_t")) then
				f.WriteLine "<td valign=""top"">" & gITEXTA("i_4") & " <b>" & JAPINFO1("cand_mnam_t")& "</b></td>"
			else
				f.WriteLine "<td>&nbsp;</td>"
			end if
			if pv_new_sessioncode <> 7000 and len(JAPINFO1("cand_onam_t")) then
				f.WriteLine "<td valign=""top"">" & gITEXTA("i_9") & " <b>" & JAPINFO1("cand_onam_t")& "</b></td>"
			else
				f.WriteLine "<td>&nbsp;</td>"
			end if
			f.WriteLine "</tr>"
			if len(JAPINFO1("cand_maiden_t")) then
'chnage by Atul at-- make colspan 2
				f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTA("i_10b") & " <b>" & JAPINFO1("cand_maiden_t") & "</b></td></tr>"
			end if
		end if

		f.WriteLine "<tr>"
		if   trim(JAPINFO1("sex_code")) = "M"  then
			f.WriteLine "<td valign=""top"">" & gITEXTA("i_5")& " <b>" & gITEXTA("i_29") & "</b></td>"
		else
			f.WriteLine "<td valign=""top"">" & gITEXTA("i_5")& " <b>" & gITEXTA("i_30") & "</b></td>"
		end if
		f.WriteLine "<td valign=""top"">" & gITEXTA("i_7") & " <b>" & JAPINFO1("maritaldsc") & "</b></td>"
		f.WriteLine "</tr>"
		
		f.WriteLine "<tr><td valign=""top"">" & gITEXTA("i_6") & " <b>" & JAPINFO1("cand_bthp_t") & "</b></TD>"
		if   len(JAPINFO1("cand_bth_d")) then
			f.WriteLine "<td valign=""top"">" & gITEXTA("i_11") & " <b>" & day(JAPINFO1("cand_bth_d")) & "-" & monthname(month(JAPINFO1("cand_bth_d")),1) & "-" & year(JAPINFO1("cand_bth_d")) & "</b></td>"
'Start chnage by Atul make colspan=2
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2
		end if
		f.WriteLine "</tr>"
'Start chnage by Atul make colspan=2
		f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTA("i_16") & " <b>" & JAPINFO1("s_nat") & "</b></TD></tr>"
' IS OR IS NOT INTERNAL AND NOT WHO 1000



' NOTES ------------------ FOR NON-INTERNAL APPLICANTS
	else
' THIS IS COPY OF BELOW FOR NON-WHO AND NON-UNAIDS
		f.WriteLine "<tr><td valign=""top"" align=""center"" colspan=""2""><strong>" & JAPINFO1("honordsc") & " " & Ucase(JAPINFO1("cand_lnam_t")) & " " & JAPINFO1("cand_fnam_t") & "</strong></td></tr>"
		if  len(JAPINFO1("cand_mnam_t")) OR len(JAPINFO1("cand_onam_t")) then
				f.WriteLine "<tr>"
			if  len(JAPINFO1("cand_mnam_t")) then
					f.WriteLine "<td valign=""top"">" & gITEXTA("i_4") & " <b>" & JAPINFO1("cand_mnam_t") & "</b></td>"
			else
					f.WriteLine "<td>&nbsp;</td>"
			end if
			if   len(JAPINFO1("cand_onam_t")) then
				f.WriteLine "<td valign=""top"">" & gITEXTA("i_9") & " <b>" & JAPINFO1("cand_onam_t") & "</b></td>"
			else
				f.WriteLine "<td>&nbsp;</td>"
			end if
			f.WriteLine "</tr>"
		end if

		if pv_new_sessioncode <> 7000 and len(JAPINFO1("cand_maiden_t")) then
'Start chnage by Atul make colspan=2
			f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTA("i_73") & " <b>" & JAPINFO1("cand_maiden_t") & "</b></td></tr>"
		end if

		f.WriteLine "<tr>"
		If JAPINFO1("cand_gnd_i") = 1  then
			f.WriteLine "<td valign=""top"">" & gITEXTA("i_5") & " <b>" & gITEXTA("i_29") & "</b></td>"
		elseif JAPINFO1("cand_gnd_i") = 0 then
				f.WriteLine  "<td valign=""top"">" & gITEXTA("i_5") & " <b>" & gITEXTA("i_30") & "</b></td>"
		end if
		f.WriteLine "<td valign=""top"">" & gITEXTA("i_7") & " <b>" & JAPINFO1("maritaldsc") & "</b></td>"
		f.WriteLine "</tr>"

		f.WriteLine "<tr>"
		f.WriteLine "<td valign=""top"">" & gITEXTA("i_6")& " <b>" &  JAPINFO1("cand_bthp_t") & "</b></TD>"
		if   len(JAPINFO1("cand_bth_d")) then
			f.WriteLine "<td valign=""top"">" & gITEXTA("i_11") & " <b>" & day(JAPINFO1("cand_bth_d")) & "-" & monthname(month(JAPINFO1("cand_bth_d")),1) & "-" & year(JAPINFO1("cand_bth_d")) & "</b></td>"
'Start chnage by Atul make colspan=2
		else
			f.WriteLine "<td>&nbsp;</td>"
'End chnage by Atul make colspan=2
		end if
		f.WriteLine "</tr>"
'Start chnage by Atul make colspan=2
		f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTA("i_16")& " <b>" &  JAPINFO1("natcty") & "</b></TD></tr>"
	end if
	
	
	
else
'Commented by Gevorg
'f.WriteLine "<tr><td valign=""top"" align=""center"" colspan=""2""><strong>" & JAPINFO1("honordsc") & " " & Ucase(JAPINFO1("cand_lnam_t")) & " " & JAPINFO1("cand_fnam_t") & "</strong></td></tr>"
		if  len(JAPINFO1("cand_mnam_t")) OR len(JAPINFO1("cand_onam_t")) then
				f.WriteLine "<tr>"
			if  len(JAPINFO1("cand_mnam_t")) then
					f.WriteLine "<td valign=""top"">" & gITEXTA("i_4") & " <b>" & JAPINFO1("cand_mnam_t") & "</b></td>"
			else
					f.WriteLine "<td>&nbsp;</td>"
			end if
			if   len(JAPINFO1("cand_onam_t")) then
				f.WriteLine "<td valign=""top"">" & gITEXTA("i_9") & " <b>" & JAPINFO1("cand_onam_t") & "</b></td>"
			else
				f.WriteLine "<td>&nbsp;</td>"
			end if
			f.WriteLine "</tr>"
		end if

		if pv_new_sessioncode <> 7000 and len(JAPINFO1("cand_maiden_t")) then
'Start chnage by Atul make colspan=2
			f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTA("i_73") & " <b>" & JAPINFO1("cand_maiden_t") & "</b></td></tr>"
		end if

		f.WriteLine "<tr>"
		If JAPINFO1("cand_gnd_i") = 1  then
			f.WriteLine "<td valign=""top"">" & gITEXTA("i_5") & " <b>" & gITEXTA("i_29") & "</b></td>"
		else
				f.WriteLine  "<td valign=""top"">" & gITEXTA("i_5") & " <b>" & gITEXTA("i_30") & "</b></td>"
		end if
		f.WriteLine "<td valign=""top"">" & gITEXTA("i_7") & " <b>" & JAPINFO1("maritaldsc") & "</b></td>"
		f.WriteLine "</tr>"

		f.WriteLine "<tr>"
		f.WriteLine "<td valign=""top"">" & gITEXTA("i_6")& " <b>" &  JAPINFO1("cand_bthp_t") & "</b></TD>"
		if   len(JAPINFO1("cand_bth_d")) then
			f.WriteLine "<td valign=""top"">" & gITEXTA("i_11") & " <b>" & day(JAPINFO1("cand_bth_d")) & "-" & monthname(month(JAPINFO1("cand_bth_d")),1) & "-" & year(JAPINFO1("cand_bth_d")) & "</b></td>"
'Start chnage by Atul make colspan=2
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->	
		end if
		f.WriteLine "</tr>"
'Start chnage by Atul make colspan=2-->
		f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTA("i_16")& " <b>" &  JAPINFO1("natcty") & "</b></TD></tr>"
' END   IS OR IS NOT INTERNAL AND NOT WHO 1000
end if

'  '--------------------- ADD ORG INFO 1000 2000 3000 4000 23 FEB 05 LJL ------------------------->
	if  pv_new_sessioncode <> 1000 then
			f.WriteLine "<TR><td valign=""top"">" & gITEXTA("i_14")& " <b>"
			if  len(JAPINFO1("natcty2")) then
				  f.WriteLine JAPINFO1("natcty2")
			else
				 f.WriteLine "-"
			end if
			f.WriteLine "</b></TD>"
			f.WriteLine "<td valign=""top"">" & gITEXTA("i_15") & " <b>"
			if   len(JAPINFO1("natcty3"))  then
				f.WriteLine JAPINFO1("natcty3")
			else
				f.WriteLine "-"
			end if
			f.WriteLine "</b></TD></tr>"
	end if

'  		'- Has your nationality ever been changed?  --->
	if pv_new_sessioncode <> 7000 AND JAPINFO1("cand_pnat_i") = "1"  then
		f.WriteLine "<tr>"
				f.WriteLine "<td valign=""top"">" & gITEXTA("i_19")
				if JAPINFO1("cand_pnat_i") = 1 then
					f.WriteLine " <b>" & gITEXTA("i_27") & "</b></td>"
				else
					f.WriteLine " <b>" & gITEXTA("i_28") & "</b></td>"
				end if

'27 MAR 11 LJL revised to include explanation of nat change
'if len(JAPINFO1("cand_expl_t")) then
		f.WriteLine "<td valign=""top"">" & gITEXTA("i_77") & " <b>" & JAPINFO1("cand_expl_t") & "</b></TD>"
'end if

		f.WriteLine "</tr>"
		
		if len(JAPINFO1("cand_newnat_c")) or len(JAPINFO1("cand_newnat_d")) then
				f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTA("i_40") & " <b><br>" & JAPINFO1("newnat") & "</b> " & gITEXTA("i_26") & "<b> " &  JAPINFO1("cand_newnat_d") & "</b></TD>"
' chnage by atul add this tr in if condition
				f.WriteLine "</tr>"			
		end if
'f.WriteLine "</tr>"
	end if

		if pv_DEBUG then
			response.write "<br>JAPINFO1->cand_id_c:="
			response.write JAPINFO1("cand_id_c")
			response.write "<br>stop 3.388"
'response.End()
		end if
'if pv_new_sessioncode <> 7000 AND JAPINFO1("cand_pnat_i") = 1  then
'<tr>
'	<td valign=""top"">" & gITEXTA("i_19") & " <b>" & gITEXTA("i_27") & "</b></TD>"
'	<td valign=""top"">" & gITEXTA("i_77")> & " <b>" & JAPINFO1("cand_expl_t") "</b></TD>"
'</tr>
'<else
'
'<tr>
'f.WriteLine "<td valign=""top""><b>" & gITEXTA("i_76")> & "</b></TD>"
'</tr>
'end if
'if pv_new_sessioncode <> 7000  AND (len(JAPINFO1("cand_newnat_c")) or len(JAPINFO1("cand_newnat_d")) ) then>
'<TR>
'f.WriteLine "<td valign=""top"">" &  gITEXTA("i_79") > - <b>" & JAPINFO1("cand_newnat_c") & "</b></TD>"
'f.WriteLine "<td valign=""top"">" & gITEXTA("i_20")> & " <b>" & JAPINFO1("cand_newnat_d") & "</b></TD>"
'</tr>
'end if


'  '---------------------------- IF NOT INTERNAL, SHOW ARE  YOU CURRENTLY ------------------>
' NOT FOR IFRC OR WTO
		if pv_new_sessioncode <> 7000 AND pv_new_sessioncode <> 3000 then
			f.WriteLine "<tr>"
'if JAPINFO1("thisorg_short") = 0  then
				f.WriteLine "<td valign=""top"">" & gITEXTA("i_22") & "<strong> "
				if JAPINFO1("thisorg_short") = 1 then
				'10 07 17 GG fixed fr issue
					if gITEXTA("i_27") = "" then
						if session("lng") = "en" then
							f.WriteLine "Yes"
						elseif session("lng") = "fr" then
							f.WriteLine "Oui"
						elseif session("lng") = "es" then
							f.WriteLine "sí"
						end if						
					else
						f.WriteLine gITEXTA("i_27")
					end if
				else
					if gITEXTA("i_28") = "" then 
						if session("lng") = "en" then
							f.WriteLine "No"
						elseif session("lng") = "fr" then
							f.WriteLine "Non"
						elseif session("lng") = "es" then
							f.WriteLine "No"
						end if
						'f.WriteLine "No"
					else
						f.WriteLine gITEXTA("i_28")
					end if					
				end if
				f.WriteLine "</strong></td>"
'27 MAR 11 LJL changed include text for A - previously applied, from being New Nationality include text
'11 MAY 11 LJL removed previously applied for WIPO
'20 MAY 15 LJL removed previously applied for WMO
			if pv_new_sessioncode <> 2800 AND pv_new_sessioncode <> 2900 then 
				f.WriteLine "<td valign=""top"">" & gITEXTA("i_81") & "<b> "
				if   JAPINFO1("thisorg_prev") = 1  then 
					if session("lng") = "en" then
						f.WriteLine "Yes"
					elseif session("lng") = "fr" then
						f.WriteLine "Oui"
					elseif session("lng") = "es" then
							f.WriteLine "sí"
					end if
						'f.WriteLine "Yes"  'f.WriteLine gITEXTA("i_27")
				else
						'f.WriteLine "No"   'f.WriteLine gITEXTA("i_28")
					if session("lng") = "en" then
						f.WriteLine "No"
					elseif session("lng") = "fr" then
						f.WriteLine "Non"
					elseif session("lng") = "es" then
							f.WriteLine "No"
					end if
				end if
				f.WriteLine "</b></td>"
'end if
		  f.WriteLine "</TR>"
		  end if
		end if

		
' NOT FOR 7000 IFRC OR WTO
		if pv_new_sessioncode <> 7000 AND pv_new_sessioncode <> 3000 then
		f.WriteLine "<tr>"
'f.WriteLine "<td valign=""top"">SHORT:" & JAPINFO1("thisorg_short") & "|CTYPE" & JAPINFO1("contractlendsc")& "</strong></TD>"
			if len(JAPINFO1("thisorg_staffno")) then
				f.WriteLine "<td valign=""top"">" & gITEXTA("i_48") & "<strong> " & JAPINFO1("thisorg_staffno") & " "
'20 DEC 06 LJL added this check on if it is truly a staff member for WHO and UNAIDS
				if pv_new_sessioncode = 1000 OR pv_new_sessioncode = 1200 then ' OR pv_new_sessioncode = 1500 by Gevorg
					If JAPINFO1("thisorg_stafftrue") > "0" then
						f.WriteLine "(Verified)"
					else
						f.WriteLine "(Applicant indicated only)"
					end if
				else
					f.WriteLine "(Applicant indicated)"
				end if
				f.WriteLine "</strong></TD>"
			end if

			if JAPINFO1("thisorg_short") = "1" then
'f.WriteLine "<td>CLDESC: : " & JAPINFO1("contractlendsc") & "|CTYPE: " & JAPINFO1("contract_type") & "</td>"
'01 MAR 11  LJL added if/then to take the WHO GSM contract name and type over the user's entered contract.
'11 MAY 11 LJL changed contract type output for internal staff to exclude for other orgs if from STAFF table join
				if len(JAPINFO1("contract_type")) AND (pv_new_sessioncode = 1000) then ' OR pv_new_sessioncode = 1500 by Gevorg
					f.WriteLine "<td valign=""top"">" & gITEXTA("i_23") & " <strong>" & JAPINFO1("contract_type")& "</strong></TD>"
				elseif  len(JAPINFO1("contractlendsc")) and NOT JAPINFO1("contractlendsc") = "-"  then
					f.WriteLine "<td valign=""top"">" & gITEXTA("i_23") & " <strong>" & JAPINFO1("contractlendsc")& "</strong></TD>"
				end if
			end if
			f.WriteLine "</TR>"
		end if

		
		
		if pv_DEBUG then
			response.write "<br>JAPINFO1->cand_id_c:="
			response.write JAPINFO1("cand_id_c")
			response.write "<br>stop 3.390"
'response.End()
		end if

' INTERNAL AND NOT IFRC 7000 or WTO 3000
'14 SEP 10 LJL removed contract details other than contract type, for ITU
			if pv_new_sessioncode <> 7000 AND pv_new_sessioncode <> 3000 AND pv_new_sessioncode <> 2400 AND pv_new_sessioncode <> 2800 then
				if pv_new_sessioncode = 1000 or pv_new_sessioncode = 1500 then
					if JAPINFO1("thisorg_stafftrue") > 0 then
						f.WriteLine "<TR>"
						f.WriteLine "<td valign=""top"">" & gITEXTA("i_68") & " <br><strong>"
						if isdate(JAPINFO1("contract_start_date")) then
							f.WriteLine day(JAPINFO1("contract_start_date")) & "-" & monthname(month(JAPINFO1("contract_start_date")),1) & "-" & year(JAPINFO1("contract_start_date"))
						end if
						f.WriteLine "</strong></td>"
						f.WriteLine "<td valign=""top"">" & gITEXTA("i_69") & " <br><strong>"
						if isdate(JAPINFO1("contract_end_date")) then
							f.WriteLine day(JAPINFO1("contract_end_date")) & "-" & monthname(month(JAPINFO1("contract_end_date")),1) & "-" & year(JAPINFO1("contract_end_date"))
						end if
						f.WriteLine "</strong></td>"
						f.WriteLine "</tr>"
					else
						if JAPINFO1("thisorg_short") = 1 then
							f.WriteLine "<TR>"
							f.WriteLine "<td valign=""top"">" & gITEXTA("i_68")
							if isdate(JAPINFO1("thisorg_start")) then
								f.WriteLine "<br><strong>" & day(JAPINFO1("thisorg_start")) & "-" & monthname(month(JAPINFO1("thisorg_start")),1) & "-" & year(JAPINFO1("thisorg_start"))
							end if
							f.WriteLine "</strong></td>"
							f.WriteLine "<td valign=""top"">" & gITEXTA("i_69")
							if isdate(JAPINFO1("thisorg_end")) then
								f.WriteLine "<br><strong>" & day(JAPINFO1("thisorg_end")) & "-" & monthname(month(JAPINFO1("thisorg_end")),1) & "-" & year(JAPINFO1("thisorg_end"))
							end if
							f.WriteLine "</strong></td>  "
							f.WriteLine "</tr>"
						end if
					end if
				else
				' OTHER ORGS STAFF CONTRACT DATES  - WMO etc
					if JAPINFO1("thisorg_short") = 1 then
						f.WriteLine "<TR>"
						if pv_new_sessioncode = 2900 then
							'CONTRACT TYPE HERE FOR WMO
							'20 MAY 15 LJL no contract start or end for WMO but show contract type
						f.WriteLine "<td valign=""top""><strong>"
'if  len(JAPINFO1("contractlendsc")) and NOT JAPINFO1("contractlendsc") = "-"  then
	f.WriteLine JAPINFO1("contractlendsc")
'end if
						'20 MAY 15 LJL no contract start or end for WMO but show contract type
						else	

						f.WriteLine "<td valign=""top"">" & gITEXTA("i_68")
						'20 MAY 15 LJL no contract start or end for WMO
						if isdate(JAPINFO1("thisorg_start")) then
							f.WriteLine "<br><strong>" & day(JAPINFO1("thisorg_start")) & "-" & monthname(month(JAPINFO1("thisorg_start")),1) & "-" & year(JAPINFO1("thisorg_start"))
						end if
						f.WriteLine "</strong></td>"
						f.WriteLine "<td valign=""top"">" & gITEXTA("i_69")
						if isdate(JAPINFO1("thisorg_end")) then
							f.WriteLine "<br><strong>" & day(JAPINFO1("thisorg_end")) & "-" & monthname(month(JAPINFO1("thisorg_end")),1) & "-" & year(JAPINFO1("thisorg_end"))
						end if
						
						'20 MAY 15 LJL no contract start or end for WMO but show contract type
						end if
						f.WriteLine "</strong></td>  "
						f.WriteLine "</tr>"
					end if
				end if
			end if
				f.WriteLine "</TABLE>"
			end if
'*************************************************************************************************-->

' ******************************************************************************
' SECTION CONTACT DETAILS Contract Information
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION CONTACT
if instr(pv_parts,"W,") then
f.WriteLine "<p><br></p>"
f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width='" &  widther & "' align=""center"">"
'start chnage by atul remove the hr  make line with tr and table-->
'f.WriteLine "<tr><td colspan=""2""><hr/></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"

' End chnage by atul remove the hr  make line with tr and table -->
	
			f.WriteLine "<tr><td colspan=""2"" width=""100%"">"
			f.WriteLine "<TABLE border=""0"" bordercolor=""black"" width=""100%"">"
			f.WriteLine "<TR><td valign=""top"">"
' until fax number
			f.WriteLine "<table width=""100%"">"
			f.WriteLine "<tr><td colspan=""3"" valign=""top"">" & gITEXTI("i_3") & "</TD></tr>"
			f.WriteLine "<tr><td colspan=""3"" valign=""top""><b>" & japinfo1("cand_padd_m") & "<br>" & JAPINFO1("cand_pcity_t") & "<br>" & JAPINFO1("cand_pzip_t") & "<br>" & JAPINFO1("permcty") & "</b></td></tr>"
			if len(JAPINFO1("cand_pphn_t")) or len(JAPINFO1("cand_wphn_t")) or len(JAPINFO1("cand_mphn_t")) then
				if len(JAPINFO1("cand_pphn_t")) then
					f.WriteLine  "<TR><td valign=""top""><I>" & gITEXTI("i_16") & "</i> <b>" & JAPINFO1("cand_pphn_t") & "</b></td>"
				else
					f.WriteLine  "<TR><td valign=""top"">" & gITEXTI("i_46") & "</td>"
				end if
				if   len(JAPINFO1("cand_wphn_t")) then
					f.WriteLine "<td valign=""top""><I>" & gITEXTI("i_17") & "</I> <b>" & JAPINFO1("cand_wphn_t") & "</b></td>"
				else
					f.WriteLine "<td valign=""top""><I>" & gITEXTI("i_17") & "</I></td>"
				end if
				if len(JAPINFO1("cand_mphn_t")) then
					f.WriteLine "<td valign=""top""><I>" & gITEXTI("i_18") & "</I> <b>" & JAPINFO1("cand_mphn_t")& "</b></td>"
				else
					f.WriteLine "<td valign=""top""></td>"
				end if
			end if
			if len(JAPINFO1("cand_pfax_t")) then
				f.WriteLine "<td valign=""top""><I>" & gITEXTI("i_15") & "</I> <b>" & JAPINFO1("cand_pfax_t") & "</b></td></tr>"
			else
				f.WriteLine "<td valign=""top""></td></tr>"
			end if
			f.WriteLine "</table></td>"			
			if   len(JAPINFO1("cand_radd_m")) then
				f.WriteLine "<td width=""50%"" valign=""top"">"
				f.WriteLine "<table>"
				f.WriteLine "<tr><td valign=""top"">" & gITEXTI("i_8") & "</TD></tr>"
				f.WriteLine "<tr><td valign=""top""><b>" & JAPINFO1("cand_radd_m") & "<br>" & JAPINFO1("cand_rcity_t") & "<br>" & JAPINFO1("cand_rzip_t") & "<br>" & JAPINFO1("rescty") & "</b></td></tr>"
				if len(JAPINFO1("cand_pphn_2_t")) or len(JAPINFO1("cand_wphn_2_t")) or len(JAPINFO1("cand_mphn_2_t")) then
					f.WriteLine "<TR>"
					if len(JAPINFO1("cand_pphn_2_t")) then
						f.WriteLine "<td valign=""top"">" &  gITEXTI("i_46") & "<I>" & gITEXTI("i_16") & "</i> <b>" & JAPINFO1("cand_pphn_2_t") & "</b></td>"
					else
						f.WriteLine "<td valign=""top"">" &  gITEXTI("i_46") & "</td>"
					end if
					f.WriteLine "</tr>"
					f.WriteLine "<tr>"
					if len(JAPINFO1("cand_wphn_2_t")) then
						f.WriteLine "<td valign=""top""><I>" & gITEXTI("i_17") & "</I> <b>" & JAPINFO1("cand_wphn_2_t")& "</b></td>"
					else
						f.WriteLine"<td valign=""top""></td>"
					end if
					f.WriteLine "</tr>"
					f.WriteLine "<tr>"
					if len(JAPINFO1("cand_mphn_2_t")) then
						f.WriteLine "<td valign=""top""><I>" & gITEXTI("i_18") & "</I> <b>" & JAPINFO1("cand_mphn_2_t") & "</b></td>"
					else
						f.WriteLine "<td valign=""top""></td>"
					end if
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					if len(JAPINFO1("cand_wfax_t")) then
						f.WriteLine "<td valign=""top""><I>" & gITEXTI("i_15") & "</I> <b>" & JAPINFO1("cand_wfax_t") & "</b></td>"
					else
						f.WriteLine "<td valign=""top""></td>"
					end if
					f.WriteLine "</tr>"
				end if
				f.WriteLine "</table></td>"
			end if

			f.WriteLine "</tr>"
			f.WriteLine "</table>"
			f.WriteLine "</TD></tr>"
			
			f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTI("i_42") & "</td></tr>"
			f.WriteLine "<tr><td valign=""top"">&nbsp;&nbsp;<I>" & gITEXTI("i_10") & "</i> <b>" & JAPINFO1("cand_email_t") & "</b></td>"
			if len(JAPINFO1("cand_wml_t")) then
				f.WriteLine "<td valign=""top"">&nbsp;&nbsp;<I>" & gITEXTI("i_13") & "</I> <b>" & JAPINFO1("cand_wml_t") & "</b></td>"
			else
				f.WriteLine "<td valign=""top"">&nbsp;</td>"
			end if
			f.WriteLine "</tr>"

			if len(JAPINFO1("cand_aci_t")) then
'01 SEP 10 LJL remove first horizontal line if first record
' start chnage by atul remove the hr  make line with tr and table-->
'f.WriteLine "<tr><td colspan=""2""><hr noshade size=""1""></td></tr>"
'f.WriteLine "<tr><td colspan=""2"">"
''f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
				f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
				f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
				f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
' End chnage by atul remove the hr  make line with tr and table -->
			
'f.WriteLine "<tr><th colspan=""2"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTI("i_19") & "_|" & intNumber  &  "'>" & gITEXTI("i_19") & "</a></th></tr>"
				f.WriteLine "<tr><td valign=""top"" colspan=""2""><p>" & gITEXTI("i_19") & "</p></td></tr>"
				
				f.WriteLine "<tr><td valign=""top"" colspan=""2""><b>" & JAPINFO1("cand_aci_t") & "</b></td></tr>"
			end if
		if pv_DEBUG then
			response.write "<br>JAPINFO1->cand_id_c:="
			response.write JAPINFO1("cand_id_c")
			response.write "<br>stop 3.389"
'response.End()
		end if

' SHOW THIS PART IF WTO 3000
'08 FEB 10 DD Blocked "In current country of residence" part for WTO by ANDing with 0
		if pv_new_sessioncode = 3000 and 0 then
			f.WriteLine "<tr><td valign=""top"" colspan=""2"">" & gITEXTG("i_21") & "</TD></tr>"
			f.WriteLine "<tr><td valign=""top"" VALIGN=""middle"" colspan=""4""><I>" & gITEXTG("i_22") & " " & gITEXTG("i_23") & "</i> <strong>"
			if   isdate(JAPINFO1("cand_cty_arr_d")) then
				f.WriteLine monthname(month(JAPINFO1("cand_cty_arr_d")),1) & "-" & year(JAPINFO1("cand_cty_arr_d"))
			end if
			f.WriteLine "</td></tr>"
		end if

' '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
		 if pv_DEBUG then
			response.write "<br>JAPINFO1->cand_id_c:="
			response.write JAPINFO1("cand_id_c")
			response.write "<br>stop 3.391"
'response.End()
		end if
			if Request.querystring("viewupd") = True and viewupd = "YES"  then
				if GETUPDS("edita_d") > GETAPPDATE("candjob_d")  then
					f.WriteLine "<TR><td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") & ">Last update date <strong>" & formatdatetime(GETUPDS("EditA_d"),1) & "AFTER APPLYING</strong></td></tr>"
				end if
			end if
' END CONTACT DETAILS
	f.WriteLine "</TABLE>"
end if
		if pv_DEBUG then
			response.write "<br>JAPINFO1->cand_id_c:="
			response.write JAPINFO1("cand_id_c")
			response.write "<br>stop 3.392"
'response.End()
		end if
' ******************************************************************************
' SECTION B
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION B
		f.WriteLine "<table><tr><td>&nbsp;</td></tr></table>"
		if instr(pv_parts,"B,") then

''-----------------------------------------NOTES ----------
'ColdFusion template for generating dynamically an html page for pdf
'conversion with the personal details.
'Must be included in the jap-pdf-single-admin.asp template to be functional.
'PDF modifications are essentialy removing the styles (css,...), removing
'the nested tables, closing all tags, etc...
'Modified from appB-view.asp
'Modified by LJL (SECANTSYS) lance.lundberg@dock.net
'MODS --
'23 FEB 05 LJL org check
'--------------------------------------------------------------->
		JAPPREFLISTsql = "	 		SELECT tx_rsys_candworkpref.candworkpref_id_c, tx_rsys_candworkpref.cand_id_c, 	core_emppref.pref_dsc_" & new_lng_code & "_t as pref_dsc_t 	FROM tx_rsys_candworkpref INNER JOIN 	core_emppref ON 	tx_rsys_candworkpref.candworkpref_id_c = core_emppref.pref_id_n 		WHERE cand_id_c = " & applicant_id
''d set JAPPREFLIST =rsys_db_select.execute(JAPPREFLISTsql)
		Set JAPPREFLIST = Server.CreateObject("ADODB.RecordSet")
		JAPPREFLIST.Open JAPPREFLISTsql, rsys_db_select, 1, 1
			
'Add by Atul
f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTB("i_1") & "</font></h2>"		
		
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""1"" width=""" & widther & """ align=""center"">"
'f.WriteLine "<tr>"
'f.WriteLine "<th colspan=""4""> <a style=""abcpdf-tag-visible: true;"" id='"   & gITEXTB("i_1") & "_|" & intNumber   & "'>"  & gITEXTB("i_1") & "</a> <a name='test' ></a> </th>"
'f.WriteLine "</tr>"
'Commented by atul
'f.WriteLine "<tr>"
'f.WriteLine "<td valign=""bottom"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTB("i_1") & "</font></h2></td>"
'f.WriteLine "</tr>"
		
		f.WriteLine "<tr>"
		f.WriteLine "<td width=""15%"" valign=""top"">" & gITEXTB("i_4") & "</TD>"
		f.WriteLine "<td width=""30%"" valign=""top"">"
			Do while not JAPPREFLIST.eof
				f.WriteLine "<b>" & JAPPREFLIST("pref_dsc_t")& "</b><br>"
				JAPPREFLIST.movenext
			loop
			 f.WriteLine "</TD>"


'    		' NEEDed Delta1 -->
'21 SEP 10 LJL changed to show the number for available levels in B
			f.WriteLine "<td width=""25%"" valign=""top"">" & gITEXTB("i_8") & "</TD>"
			f.WriteLine "<td width=""30%"" valign=""top""><b>"
'			  if JAPINFO1("cand_avl_time_t") = "100%"  then
'				f.WriteLine "100%"
'			  end if
'			  if   JAPINFO1("cand_avl_time_t") = "85%"  then
'				f.WriteLine "85%"
'			  end if
'			  if   JAPINFO1("cand_avl_time_t") = "80%"  then
'				f.WriteLine "80%"
'			  end if
'			  if   JAPINFO1("cand_avl_time_t") = "80%"  then
'				f.WriteLine "80%"
'			  end if
'			  if  JAPINFO1("cand_avl_time_t") = "50%"  then
'			  f.WriteLine "50%"
'			  end if
			  if   JAPINFO1("cand_avl_time_t") = "Any"  then
				f.WriteLine gITEXTB("i_21")
			else
				f.WriteLine JAPINFO1("cand_avl_time_t")
			  end if
			  f.WriteLine "</b></td>"
		f.WriteLine "</tr>"

'  		' NEEDed Delta1 -->
		if len(JAPINFO1("avl_dsc")) then
			f.WriteLine "<TR><td width=""15%"" valign=""top"" >" &  gITEXTB("i_17") & "</TD><td width=""30%"" valign=""top"" ><b>" & JAPINFO1("avl_dsc") & "</b></TD>"
		end if
		
		if len(JAPINFO1("cand_avl_other_t")) then
			f.WriteLine "'<td width=""25%"" valign=""top"" >" & gITEXTB("i_10") & "</TD><td width=""30%"" valign=""top"" ><b>" & JAPINFO1("cand_avl_other_t") & "</b></TD>"
			f.WriteLine "'</tr>"
		else
			f.WriteLine "<td width=""25%"" valign=""top"" ></TD><td width=""30%"" valign=""top"" ></TD></tr>"
		end if
' '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
		  if   Request.querystring("viewupd") = "YES"  then
		  if   GETUPDS("editB_d") > GETAPPDATE("candjob_d")  then
			f.WriteLine "<TR><td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") & """>Last update date <strong>" & date(GETUPDS("EditB_d"),1) & " AFTER APPLYING</strong></td>"
			f.WriteLine "</tr>						"
		  end if
		  end if
		f.WriteLine "</table>"

' END CHECK IF PERSON SELECTED SECTION B
		end if
			if pv_DEBUG then
				response.write "<br>JAPINFO1->cand_id_c:="
				response.write JAPINFO1("cand_id_c")
				response.write "<br>stop 3.393"
'	response.End()
			end if
' ******************************************************************************
' SECTION S - AREAS OF EXPERTISE
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION S
		if instr(pv_parts,"S,") then
Dim pv_rower_aoe
'25 APR 16 GG Added additional condition for Areas of Expertise
'30 MAY 16 GG Deleted additional condition for Areas of Expertise a.candccog_thisorg_2400 = 1 
gAOEsql = "SELECT c.ccog_dsc_" & new_lng_code & "_t AS ccogdsc, c.ccog_code_c, a.cand_id_c, a.ccog_id_c, a.crd_d, a.candccog_dsc_t, a.candccog_rank_i, cr.ccogrank_dsc_" & new_lng_code & "_t AS ccogrank, cr.ccogrank_font_color_c, cy.ccogyears_dsc_" & new_lng_code & "_t AS ccyears FROM dbo.core_ccogtblt c INNER JOIN dbo.tx_rsys_cand_ccog a ON c.ccog_id_c = a.ccog_id_c LEFT OUTER JOIN dbo.tr_rsys_ccogyears cy ON a.candccog_years_c = cy.ccogyears_id_c LEFT OUTER JOIN dbo.tr_rsys_ccogrank cr ON a.candccog_rank_i = cr.ccogrank_id_c WHERE (a.cand_id_c = " & applicant_id & ") AND a.candccog_type_i = 11 AND c.ccog_thisorg_" & pv_new_sessioncode & " = 1 ORDER BY c.ccog_dsc_" & new_lng_code & "_t"
'response.write "QOE:" & gAOEsql
'response.end
''d set gAOE = rsys_db_select.execute(gAOEsql)
			Set gAOE = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			gAOE.Open gAOEsql, rsys_db_select, 1, 1
''--------------------- SKILLS  EDITING-------------------------------->
		if not gAOE.eof then
			if pv_DEBUG then
				response.write "<br>not gAOE.eof"
				response.write "<br>stop 3.394"
'response.End()
			end if
'Add by Atul
f.WriteLine "<p><br></p>"
			f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTS("i_1") & "</font></h2>"
			f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=100% align=""center"">"
'21 DEC 06 lJL removed per WTO request
'f.WriteLine "<tr><td valign=""top"" colspan=""3"">&nbsp;</TD></tr>"
'f.WriteLine "<TR><th colspan=""3""> &lt;a style=""abcpdf-tag-visible: true;"" id='"  & gITEXTS("i_1") & "_|" & intNumber  &  "'&gt;"   & gITEXTS("i_1") & "&lt;/a&gt;</th></tr>"
'Commented by atul
'f.WriteLine "<TR><td valign=""top"" colspan=""3"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTS("i_1") & "</font></h2></TD></tr>"
			
			if not gAOE.eof then
				pv_rower_aoe = 0
				Do while not gAOE.eof
					pv_rower_aoe = pv_rower_aoe + 1
					if pv_rower_aoe = 1  then
						f.WriteLine "<tr>"
					end if
					f.WriteLine "<td valign=""top"" align=""left""> <strong> <font size=""1"">" & gAOE("ccogdsc") & "</font> </strong>"
					if len(gAOE("ccogrank")) then
						f.WriteLine "<br> &nbsp; &nbsp;<font color=""" & gAOE("ccogrank_font_color_c")& """>" & gAOE("ccogrank") & "</font>"
					end if
					if len(gAOE("ccyears")) then
						f.WriteLine "<br> &nbsp; &nbsp;<font color=""" & gAOE("ccogrank_font_color_c") & """>" & gAOE("ccyears") & "</font>"
					end if
					f.WriteLine "</td>"
					if   pv_rower_aoe = 5  then
						f.WriteLine "</tr>"
						pv_rower_aoe = 0
					end if
				gAOE.movenext
				loop
				if   pv_rower_aoe < 5  then
					f.WriteLine "</tr>"
				end if
			else
				f.WriteLine "<tr>"
				f.WriteLine "<td valign=""top"" colspan=""5"" class=""littletitle"" align=""left"">&nbsp;</td>"
				f.WriteLine "</tr>"
			end if
			if pv_DEBUG then
				response.write "<br>JAPINFO1->cand_id_c:="
				response.write JAPINFO1("cand_id_c")
				response.write "<br>stop 3.395"
'response.End()
			end if
'  '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
				if   Request.querystring("viewupd") = True and viewupd = "YES"  then
					if   GETUPDS("editS_d") > GETAPPDATE("candjob_d")  then
						f.WriteLine "<TR>"
						f.WriteLine "<td valign=""top"" colspan=""3"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") & """>Last update date: <strong>" & GETUPDS("EditS_d") & " AFTER APPLYING</strong></td>"
						f.WriteLine "</tr>"
					end if
				end if
				f.WriteLine "</TABLE>"
			end if
		end if
		
		
		if pv_DEBUG then
			response.write "<br>SECTION OTHER SKILLS - start"
			response.write JAPINFO1("cand_id_c")
			response.write "<br>stop 3.396"
'response.End()
		end if
		
' ******************************************************************************
' SECTION OTHER SKILLS - IFRC ONLY
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION RC
		if pv_new_sessioncode = 7000 then
			if instr(pv_parts,"RC,") then
				GETMISCsql = " SELECT candmisc_driving_i, candmisc_maxstaff_c, 	candmisc_maxteam_c, candmisc_emops_i, candmisc_emops_where_c, candmisc_emops_type_c, candmisc_emops_capacity_c 	FROM tx_rsys_candmisc WHERE cand_id_c = " & applicant_id & " 	"
''d set GETMISC =rsys_db_select.execute(GETMISCsql)
				Set GETMISC = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
				rsys_db_select.CommandTimeout = 320
				GETMISC.Open GETMISCsql, rsys_db_select, 1, 1
				if pv_DEBUG then
					response.write "<br>SSECTION E2 - EDUCATION DETAILS - start"
					response.write "<br>GETMISCsql=  " & GETMISCsql
					response.write "<br>stop 3.397"
'response.End()
				end if
				if not GETMISC.eof then
				f.WriteLine "<p><br></p>"
					f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTS("i_42") & "</font></h2>"
					f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" & widther& """ align=""center"">"
					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top"" colspan=""3"">&nbsp;</td>"
					f.WriteLine "</tr>"
					
'f.WriteLine "<tr>"
'f.WriteLine "<th colspan=""3"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTS("i_42") & "_|" & intNumber   & "'>" & gITEXTS("i_42") &  "</a></th>"
'f.WriteLine "</tr>"
					
'commented by atul
					
'f.WriteLine "<tr>"
'f.WriteLine "<td valign=""top"" colspan=""3"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTS("i_42") & "</font></h2></td>"
'f.WriteLine "</tr>"
					
					
					f.WriteLine "<tr>"
					f.WriteLine "<td><strong>" & gITEXTS("i_43") & "</strong></td>"
					f.WriteLine "</tr>"
					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""3"">" & gITEXTS("i_44") & "<strong>"
					if GETMISC("candmisc_driving_i") = 1 then
						f.WriteLine giTEXTS("i_45")
					else
						f.WriteLine gITEXTS("i_46")
					end if
					f.WriteLine "</strong></td>"
					f.WriteLine "</TR>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""3"">" & gITEXTS("i_22") & "</TD>"
					f.WriteLine "</tr>"
					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""3"">" & gITEXTS("i_23") & "<strong>"
					if GETMISC("candmisc_maxstaff_c") = 0 then f.WriteLine "0" end if
					if GETMISC("candmisc_maxstaff_c") = 10 then f.WriteLine "1 - 5" end if
					if GETMISC("candmisc_maxstaff_c") = 20 then f.WriteLine "6 - 10" end if
					if GETMISC("candmisc_maxstaff_c") = 30 then f.WriteLine "10+" end if
					f.WriteLine "</strong></td>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""3"">" & gITEXTS("i_24") & "<strong>"
					if GETMISC("candmisc_maxteam_c") = 0 then f.WriteLine "0" end if
					if GETMISC("candmisc_maxteam_c") = 40 then f.WriteLine "1 - 10" end if
					if GETMISC("candmisc_maxteam_c") = 50 then f.WriteLine "11 - 50" end if
					if GETMISC("candmisc_maxteam_c") = 60 then f.WriteLine "51 - 250" end if
					if GETMISC("candmisc_maxteam_c") = 70 then f.WriteLine "251+" end if
					f.WriteLine "</strong></td>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""3"">" & gITEXTS("i_25") & "</TD>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""3"">" & gITEXTS("i_26") & "<strong>"
					if GETMISC("candmisc_emops_i") = 1 then f.WriteLine "Yes" end if
					if GETMISC("candmisc_emops_i") = 0 then f.WriteLine "No" end if
					f.WriteLine "</strong></td>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
' chnage by atul make colspan=3 from 1"-->
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""3"">" & gITEXTS("i_27") & "</TD>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""1"">" & gITEXTS("i_28") & "</td>"
					f.WriteLine "<td valign=""top""  colspan=""2""><strong>" & gETMISC("candmisc_emops_type_c") & "</strong></TD>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""1"">" & gITEXTS("i_29") & "</TD>"
					f.WriteLine "<td valign=""top""  colspan=""2""><strong>" & gETMISC("candmisc_emops_where_c") & "</strong></TD>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""  align=""left"" colspan=""1"">" & gITEXTS("i_30") & "</TD>"
					f.WriteLine "<td valign=""top""  colspan=""2""><strong>" & gETMISC("candmisc_emops_capacity_c") & "</strong></TD>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
' chnage by atul make colspan=3 from 1"-->
					f.WriteLine "<td valign=""top"" class=""textbold"" align=""left"" colspan=""3"">&nbsp;</td>"
					f.WriteLine "</tr>"
					f.WriteLine "</table>"
				end if ' if rec ord exists
			end if
		end if

		if pv_DEBUG then
			response.write "<br>SSECTION E2 - EDUCATION DETAILS - start"
			response.write JAPINFO1("cand_id_c")
			response.write "<br>stop 3.398"
'response.End()
		end if

' ******************************************************************************
' SECTION C - LANGUAGES
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION C
		if instr(pv_parts,"C,") then

Dim GETCsql, GETC
		GETCsql = "{call erstp_rsys_cand_lng_full_" & new_lng_code & "(" & applicant_id & ")}	"
'response.write GETCsql
'response.end
''d set GETC =rsys_db_select.execute(GETCsql)
		Set GETC = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		GETC.Open GETCsql, rsys_db_select, 1, 1

''-----------------------------------------NOTES ----------
'MODS --
'09 MAR 05 LJL removed UN test for WTO
'RR - LISTGETAT's
'--------------------------------------------------------------->
'Added by atul
f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTC("i_1") & "</font></h2>"
		
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'    		'------------------------ had included width='100%' but that crashed the pdf creation ---------------->
'f.WriteLine "<TR>"
' atul make space in  th and  a tag-->
'f.WriteLine "<th colspan=""4"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTC("i_1") & "_|" & intNumber  &  "'>" & gITEXTC("i_1") & "</a></th>"
'f.WriteLine "</TR>"
'Commented by atul
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTC("i_1") & "</font></h2></td>"
'f.WriteLine "</TR>"
		

' IFRC 7000 NO
		if pv_new_sessioncode <> 7000 then
' atul make colspan=4-->
			f.WriteLine "<TR><td valign=""top"" colspan=""2"">" & gITEXTC("i_3") & " 1 : <b> " & gETC("mt1") & "</b></td>"
			if   GETC("mt2") <> "" then
				f.WriteLine "<td valign=""top"" colspan=""2"">" & gITEXTC("i_3") & " 2 : <b> " & gETC("mt2") & "</b></td>"
'Start chnage by Atul make colspan=2-->
			else
			f.WriteLine "<td colspan=""2"">&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
			end if
		f.WriteLine "</TR>"
		end if

' '------------------ ADD ORG INFO 1000 2000 3000 4000 09 MAR 05 LJL --------------------->
		if pv_new_sessioncode <> 3000  AND pv_new_sessioncode <> 7000 then
'atul make colspan =4 from 1-->
			f.WriteLine "<TR><td valign=""top"" colspan=""4""><I>" & gITEXTC("i_21") & "</I><b> "
			if GETC("cand_lpr_i") = 1 then
			f.WriteLine gITEXTC("i_5")
			else
			f.WriteLine gITEXTC("i_6")
			end if
			f.WriteLine "</b></TD></tr>"
			if   GETC("cand_lpr_i") = 1  then
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"" colspan=""4"">" & gITEXTC("i_7")& "<b> " &  GETC("cand_unp_t") & "</b></td>"
				f.WriteLine "</TR>"
			end if
		end if

		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""4"">"
		f.WriteLine "<TABLE border=""0"" bordercolor=""blue"" width=""100%"">"
		f.WriteLine "<TR><td valign=""top""  width=""20%""><I>" & gITEXTC("i_10") & "</I></td>"
		f.WriteLine "<td valign=""top""  width=""20%""><I>" & gITEXTC("i_11") & "</I></td>"
		f.WriteLine "<td valign=""top""  width=""20%""><I>" & gITEXTC("i_12") & "</I></td>"
		f.WriteLine "<td valign=""top""  width=""20%""><I>" & gITEXTC("i_13") & "</I></td>"
		if pv_new_sessioncode = "2500" then
		f.WriteLine "<td valign=""top""  width=""20%""><I>" & gITEXTC("i_37") & "</I></td>"
		else
'change by Atul add one column becuase colspan was not correct of above table
		f.WriteLine "<td valign=""top""  width=""20%"">&nbsp;</td>"
		end if
		f.WriteLine "</TR>"


'08 SEP 10 LJL revised to be in right order for languages - per ITU
if session("lng") = "en" then


' IFRC 7000 NO
		if pv_new_sessioncode = 7000 OR pv_new_sessioncode = 2500 OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_23") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b>"
' UNESCO 2500 ADD CHINESE
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_ar_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			
'Start chnage by Atul make colspan=2-->
			else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
			
			end if
			f.WriteLine "</TR>"
' END ARABIC FOR IFRC 7000
		end if

' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_35") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_cn_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
			else
				f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
			end if
			f.WriteLine "</TR>"
' END CHINESE FOR UNESCO 2500
		end if
		f.WriteLine "<TR>"

		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_14") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"

' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_en_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"

		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_15") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"

' UNESCO 2500 ADD
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_fr_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"

' UNESCO 2500 ADD RUSSIAN
		if pv_new_sessioncode = 2500 OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_36") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
		if pv_new_sessioncode = 2500 then
			pv_newlevel = "pv_level" & GETC("cand_ru_und_n")
' Atul . make out side  to close tag of tr and td -->	
			f.WriteLine eval(pv_newlevel) & "&nbsp;"
' END RUSSIAN FOR UNESCO 2500
		end if
		f.WriteLine "</b></TD></tr>"
		end if

'03 JUN 15 LJL no spanish as working lang for WMO
if pv_new_sessioncode = 2900 then

else

		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_16") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_es_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"

end if


elseif session("lng") = "fr" then

'Atul  Start TR tag-->
		f.WriteLine "<tr>"
		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_14") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"

' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_en_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"

' IFRC 7000 NO
		if pv_new_sessioncode = 7000 OR pv_new_sessioncode = 2500 OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_23") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b>"
' UNESCO 2500 ADD CHINESE
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_ar_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
			else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
			end if
			f.WriteLine "</TR>"
' END ARABIC FOR IFRC 7000
		end if

' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 	OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_35") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_cn_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
			else
				f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
			end if
			f.WriteLine "</TR>"
' END CHINESE FOR UNESCO 2500
		end if
		f.WriteLine "<TR>"

'03 JUN 15 LJL no spanish as working lang for WMO
if pv_new_sessioncode = 2900 then
	
else

		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_16") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_es_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"
end if

		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_15") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"

' UNESCO 2500 ADD
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_fr_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"

' UNESCO 2500 ADD RUSSIAN
		if pv_new_sessioncode = 2500 OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_36") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
		if pv_new_sessioncode = 2500 then
			pv_newlevel = "pv_level" & GETC("cand_ru_und_n")
' make tr and td close tag out side if condition-->
			f.WriteLine eval(pv_newlevel) & "&nbsp;" '</TD></tr>
' END RUSSIAN FOR UNESCO 2500
		end if
		f.WriteLine "</b></TD></tr>"
		end if



elseif session("lng") = "es" then



' IFRC 7000 NO
		if pv_new_sessioncode = 7000 OR pv_new_sessioncode = 2500 OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_23") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ar_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b>"
' UNESCO 2500 ADD CHINESE
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_ar_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
			else
				f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->	
			end if
			f.WriteLine "</TR>"
' END ARABIC FOR IFRC 7000
		end if

' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_35") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_cn_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_cn_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
			else
				f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->	
			end if
			f.WriteLine "</TR>"
' END CHINESE FOR UNESCO 2500
		end if
'Atul Comment tr tag-->
'f.WriteLine "<TR>"

'03 JUN 15 LJL no spanish as working lang for WMO
if pv_new_sessioncode = 2900 then
	
else
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_16") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_es_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_es_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"
end if

		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_15") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_fr_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"

' UNESCO 2500 ADD
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_fr_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"

' Atul Add TR Tag -->
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_14") & "</b></td>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_spk_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_rd_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
		f.WriteLine "<td valign=""top""><b>"
		pv_newlevel = "pv_level" & GETC("cand_en_wr_n")
		f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"

' UNESCO 2500 ADD CHINESE
		if pv_new_sessioncode = 2500 then
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_en_und_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
		else
			f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->		
		end if
		f.WriteLine "</TR>"


' UNESCO 2500 ADD RUSSIAN
		if pv_new_sessioncode = 2500 OR pv_new_sessioncode = 2400 then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gITEXTC("i_36") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ru_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			if pv_new_sessioncode = 2500 then
			pv_newlevel = "pv_level" & GETC("cand_ru_und_n")
' Atul make tr and td outside if -->  
			f.WriteLine eval(pv_newlevel) & "&nbsp;" '</TD></tr>
' END RUSSIAN FOR UNESCO 2500
			end if
		f.WriteLine "</b></TD></tr>"
		end if

'08 SEP 10 LJL end languages in session language order
end if


		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5""><I>" & gITEXTC("i_17") & "</I></td>"
		f.WriteLine "</TR>"

		if  GETC("ol1") <> "" and GETC("ol1") <> "-"  then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gETC("ol1") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol1_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol1_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol1_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
' UNESCO 2500 ADD CHINESE
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol1_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
			else
				f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->	
			end if
			f.WriteLine "</TR>"
		end if

		 if  GETC("ol2") <> "" and GETC("ol2") <> "-"  then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gETC("ol2") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol2_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol2_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol2_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
' UNESCO 2500 ADD CHINESE
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol2_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
			else
				f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->	
			end if
			f.WriteLine "</TR>"
		end if

		if  GETC("ol3") <> "" and GETC("ol3") <> "-"  then
			f.WriteLine "<TR><td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gETC("ol3") & "</b></td>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol3_spk_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol3_rd_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
			f.WriteLine "<td valign=""top""><b>"
			pv_newlevel = "pv_level" & GETC("cand_ol3_wr_n")
			f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
' UNESCO 2500 ADD CHINESE
			if pv_new_sessioncode = 2500 then
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol3_und_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
'Start chnage by Atul make colspan=2-->
			else
				f.WriteLine "<td>&nbsp;</td>"
' End chnage by Atul make colspan=2-->	
			end if
			f.WriteLine "</TR>"
		end if

' IFRC 7000 NO
		if pv_new_sessioncode = 7000 then
			if  GETC("ol4") <> "" and GETC("ol4") <> "-"  then
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gETC("ol4") & "</b></td>"
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol4_spk_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>" & "<td valign=""top"" & "'><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol4_rd_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD><td valign=""top"" & "'><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol4_wr_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
' UNESCO 2500 ADD CHINESE
' Atul make   td out isde if-->
				f.WriteLine "<td valign=""top""><b>"
				if pv_new_sessioncode = 2500 then
'f.WriteLine "<td valign=""top""><b>"
					pv_newlevel = "pv_level" & GETC("cand_ol4_und_n")
					f.WriteLine eval(pv_newlevel) & "&nbsp;"
				end if
				f.WriteLine "</b></TD>"
				f.WriteLine "</TR>"
			end if
			if  GETC("ol5") <> "" and GETC("ol5") <> "-"  then
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"">&nbsp; &nbsp; &nbsp; &nbsp; <b>" & gETC("ol5") & "</b></td>"
				f.WriteLine "<td valign=""top""><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol5_spk_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>" & "<td valign=""top"" & "'><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol5_rd_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>" & "<td valign=""top"" & "'><b>"
				pv_newlevel = "pv_level" & GETC("cand_ol5_wr_n")
				f.WriteLine eval(pv_newlevel) & "&nbsp;</b></TD>"
' UNESCO 2500 ADD CHINESE
' make TD out side if-->
				f.WriteLine "<td valign=""top""><b>"
				if pv_new_sessioncode = 2500 then
'f.WriteLine "<td valign=""top""><b>"
					 pv_newlevel = "pv_level" & GETC("cand_ol5_und_n")
					f.WriteLine eval(pv_newlevel) & "&nbsp;"
				end if
				f.WriteLine "</b></TD>"
				f.WriteLine "</TR>"
			end if
		end if
		f.WriteLine "</table>"
		f.WriteLine "</td>"
		f.WriteLine "</TR>"
' '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
		if  Request.querystring("viewupd") = "YES"  then
			if   GETUPDS("editC_d") > GETAPPDATE("candjob_d")  then
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63")  & """>Last update date <strong>" & formatdatetime(GETUPDS("EditC_d"),1) & " AFTER APPLYING</strong></td>"
				f.WriteLine "</TR>"
			end if
		 end if
		f.WriteLine "</TABLE>"
' END CHECK IF PERSON SELECTED SECTION C
		end if
		
		
' ******************************************************************************
' SECTION E2 - EDUCATION DETAILS
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION E
		if instr(pv_parts,"E,") then

'RR #paragraphformat line 80-95
'24 JAN 11 LJL added new edu level names per org - set so that all are on the same level
		JAPEDUsql = "{call erstp_rsys_candedu_main_" & new_lng_code & "_" & pv_new_sessioncode & "(" & applicant_id & ")}"
'response.write "JAPEDU:" & JAPEDUsql
'response.end
''d set JAPEDU =rsys_db_select.execute(JAPEDUsql)
		Set JAPEDU = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPEDU.Open JAPEDUsql, rsys_db_select, 0, 1

		JAPEDULINESsql = "{call erstp_rsys_candedu_list_" & new_lng_code & "_" & pv_new_sessioncode & "(" & applicant_id & ")}"
		On Error Resume Next
'set JAPEDULINES =rsys_db_select1.execute(JAPEDULINESsql)
		Set JAPEDULINES = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPEDULINES.Open JAPEDULINESsql, rsys_db_select1, 0, 1
		If Err.Number <> 0 Then
		   Response.Write "An error has occurred!<br>"
		   Response.Write "Error number:      " & Err.number & "<br>"
		   Response.Write "Error description: " & Err.description & "<br>"
   		end if

		if pv_DEBUG then

			response.write "<br>SECTION E2 - EDUCATION DETAILS - start"
			response.write "<br>JAPEDULINES.recordcount=  "
			response.write JAPEDULINES.recordcount
			response.write(JAPEDULINES.Fields(0).Status)
			response.write "<br>stop 3.399"
'response.End()
		end if
'Add by Atul
f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTE("i_1") & "</font></h2>"
				
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
		if JAPEDU.eof = false then
'f.WriteLine "<TR>"
' make a space i th and  a  tag-->
'f.WriteLine "<th colspan=""4""> <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTE("i_1") & "_|" & intNumber  &  "'>"  & gITEXTE("i_1") &  "</a></th></tr>"
		
'Comented by Atul
'f.WriteLine "<TR><td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """>"
'f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTE("i_1") & "</font></h2></TD></tr>"


		f.WriteLine "<TR><td valign=""top"" class=""teal10"" colspan=""2""><font color=""Teal""><b><i>" & gITEXTE("i_37")& "</i></b></font></TD>"
		f.WriteLine "<td valign=""top"" colspan=""2""><b>" & JAPEDU("edddsc") & "</b></TD></tr>"


		
' start chnage by atul remove the hr  make line with tr and table-->
'f.WriteLine "<TR><td valign=""top"" colspan=""4"" ><hr noshade size=""1""></TD></tr>"
		
'f.WriteLine "<tr><td valign=""top"" colspan=""4"" >"
'f.WriteLine "<table width=""100%"" style=""border: 2px solid black""><tr>"
'f.WriteLine "<td style=""background-color:black; width:2.5%;"" >&nbsp;</td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""4""><br></td><tr>"
		f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
		f.WriteLine "<td valign=""top"" colspan=""4"">&nbsp;</td>"
		f.WriteLine "</tr>"	
	
'10 JUN 15 GG Changed i_39 to i_14 and moved here	
		f.WriteLine "<tr><td valign=""top"" colspan=""4""><TABLE border=""0"">"
		if not JAPEDU.eof  then
			f.WriteLine "<TR><td valign=""top"" colspan=""4""><I>" & gITEXTE("i_14") & "</I></TD></tr>"
			f.WriteLine "<TR><td valign=""top"" colspan=""4""><b>" & JAPEDU("cand_edu_type_other_m") & "</b></TD></tr>"
		end if

''-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->

		if   Request.form("viewupd") = "YES"  then
			if   GETUPDS("editE_d") > GETAPPDATE("candjob_d")  then
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_62")& """><font color=""" & gITEXTPH("i_63") & """>Last update date <strong>" & formatdatetime(GETUPDS("EditE_d"),1) & " AFTER APPLYING</strong></td>"
				f.WriteLine "</tr>"
			end if
		end if

		f.WriteLine "</td></tr></TABLE>"

		
'06 JUN 15 GG Added Areas of expertise for WMO 
		IF pv_new_sessioncode = 2900 Then
			f.WriteLine "<tr>"
			f.WriteLine "<td valign=""top"" colspan=""4"">"			
			f.WriteLine "<table border=""0"" width=""100%"">"
			
 			dim areaOfExpSql, areaOfExp,countAOE,obj_select_db_CmdII,pv_areaOfExp		
			set obj_select_db_CmdII = server.CreateObject("adodb.command")
			obj_select_db_CmdII.ActiveConnection = rsys_db_select

			areaOfExpSql = "SELECT a.candccog_id_c, a.ccog_id_c, a.candccog_percentage_i ,a.candccog_type_i,b.ccog_dsc_" & new_lng_code & "_t as ccogdsc FROM tx_rsys_cand_ccog a LEFT OUTER JOIN core_ccogtblt b ON  a.ccog_id_c = b.ccog_id_c  WHERE cand_id_c = " & applicant_id & " AND candccog_type_i in (21,22,23) AND candccog_thisorg_"& pv_new_sessioncode &" = 1 ORDER BY candccog_type_i DESC"
			obj_select_db_CmdII.CommandText = areaOfExpSql
			set areaOfExp=obj_select_db_CmdII.execute()

			'if areaOfExp.eof = false then
			'    dim collItemsAOE
			'	collItemsAOE = areaOfExp.GetRows()
			'	countAOE= UBound(collItemsAOE, 2) + 1
			'	areaOfExp.movefirst
			'	pv_areaOfExp = areaOfExp("candccog_type_i")
				
			'else
			'	pv_areaOfExp = "21"
			'	countAOE = 0
			'end if
			
			'15 FEB 10 DD Provide validation for Area of Study for Education -- To collect Addition of percentages of selected AOE
		   ' if countAOE > 0 then
		   '     areaOfExp.movefirst
		   ' end if
			dim percentCounter
			percentCounter = 0
			'Do while areaOfExp.eof=false
				'percentCounter = percentCounter + areaOfExp("candccog_percentage_i")
				'areaOfExp.movenext
		   ' loop
		   
			'need to change Areas of expertise
			'f.WriteLine "<TR><td valign=""top"" colspan=""3"" ><strong>" & gITEXTE("i_38") & "</strong></TD></tr>"

			f.WriteLine "<tr>"
			f.WriteLine "<td width=""5%"">" & gITEXTE("i_82") & "</td>"
			f.WriteLine "<td width=""65%"">" & gITEXTE("i_22") & "</td>"
			f.WriteLine "<td width=""30%"">" & gITEXTE("i_84") & "</td>"
			'f.WriteLine "<td bgcolor=""pink""> </td>"
			f.WriteLine "</tr>"
			
			Dim count
			count = 1
		    Do while not areaOfExp.eof
				f.WriteLine "<tr>"
				f.WriteLine "<td width=""5%""><b>" & count & ".</b></td>"
				f.WriteLine "<td width=""65%""><b>" & areaOfExp("ccogdsc") & "</b></td>"
				f.WriteLine "<td width=""30%"" align=""center"" ><b>" & areaOfExp("candccog_percentage_i") & "%</b></td></tr> "
						
				count = count + 1
				areaOfExp.movenext
			loop
				
			f.WriteLine "</table>"
			f.WriteLine "</td></tr>"
			Set obj_select_db_CmdII = nothing
		End If
		
		
		
		
'f.WriteLine "<tr><td colspan=""4""><br></td><tr>"
'f.WriteLine "</td></tr>"
' End chnage by atul remove the hr  make line with tr and table -->

''------ insertion ------->
		f.WriteLine "<TR><td valign=""top"" colspan=""4"" ><strong>" & gITEXTE("i_38") & "</strong></TD></tr>"

		f.WriteLine "<TR><td valign=""top"" width=""5%""><em>" & gITEXTE("i_23") & "</em></td>"
		f.WriteLine "<td valign=""top"" width=""45%""><em>" & gITEXTE("i_8") & "</em></td>"
		f.WriteLine "<td valign=""top"" width=""10%""><em>" & gITEXTE("i_9") & "</em></td>"
		f.WriteLine "<td valign=""top"" width=""40%""><em>" & gITEXTE("i_10") & "</em></TD></tr>"
		end if

		if pv_DEBUG then
			response.write "<br>SSECTION E2 - EDUCATION DETAILS - start"
			response.write "<br>pv_newdb = "  & pv_newdb
			response.write "<br>JAPEDULINESsql=  " & JAPEDULINESsql

			response.write "<br> JAPEDULINES.edtdsc="
			response.write JAPEDULINES("edtdsc")
			response.write "<br> JAPEDULINES.ctyname"
			response.write JAPEDULINES("ctyname")
			 Set objASPError = Server.GetLastError
			    Response.Write("Description = " & objASPError.Description & "<br/>")
				Response.Write("ASPCode = " & objASPError.ASPCode & "<br/>")
			response.write "<br>stop 3.400"
'response.End()
		end if
		Do while not JAPEDULINES.eof
			f.WriteLine "<TR><td valign=""top"" width=""10%"">"
			if   JAPEDULINES("candedu_start_d") > 0 OR JAPEDULINES("candedu_start_d") <> ""  then
				f.WriteLine "<b>"  & JAPEDULINES("candedu_start_m") & "</b> "
				f.WriteLine "<b>"  & JAPEDULINES("candedu_start_d") & "</b>"
			end if
			f.WriteLine "<BR>"
			if   JAPEDULINES("candedu_end_d") > 0 OR JAPEDULINES("candedu_end_d") <> ""  then
				f.WriteLine "<b>" & JAPEDULINES("candedu_end_m") & "</b> "
				f.WriteLine "<b>" & JAPEDULINES("candedu_end_d") & "</b>"
			end if
			f.WriteLine "</TD>"
'commented by at due to html data issue - f.WriteLine "<td valign=""top""><b>" & JAPEDULINES("candedu_loc_m") & "</b><Br><b>" & JAPEDULINES("candedu_location_c") & "</b><br><b>" & JAPEDULINES("ctyname") & "</b></TD>"
''f.WriteLine "<td valign=""top"" colspan=""2""><b>" & JAPEDULINES("edtdsc") & "</b></TD>"
			
			f.WriteLine "<td valign=""top"" width=""40%""><b>" & JAPEDULINES("candedu_loc_m") & "</b><Br><b>" & JAPEDULINES("candedu_location_c") & "</b><br><b>" & JAPEDULINES("ctyname") & "</b></TD>"
			f.WriteLine "<td valign=""top"" width=""10%""><b>" & JAPEDULINES("edtdsc") & "</b></TD>"
			f.WriteLine "<td valign=""top"" width=""40%""><b>" & JAPEDULINES("candedu_rem_m") & "</b></TD></tr>"

			
'commented by at due to html data issue -f.WriteLine "<td valign=""top""><b>" & JAPEDULINES("candedu_rem_m") & "</b></TD>
'f.WriteLine "</tr>"
' commented by at due to html error -f.WriteLine "<tr>"
			
			if pv_DEBUG then
				response.write "<br>SSECTION E2 - EDUCATION DETAILS - cc"
				response.write "<br>GETMISCsql=  " & GETMISCsql
				response.write "<br>stop 3.401"
				response.flush()
'response.End()
			end if
			JAPEDULINES.movenext
		loop
		f.WriteLine "</TABLE>"
		
' END CHECK IF PERSON SELECTED SECTION E
	end if

' ******************************************************************************
' SECTION F2 - INTERNATIONAL EMPLOYMENT
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION F2
' DO NOT SHOW FOR IFRC AND WTO
	if instr(pv_parts,"F,") then
		if pv_new_sessioncode <> 3000 AND pv_new_sessioncode <> 7000 then
			JAPINFOIEsql = " SELECT     c.upd_d, c.cand_lnam_t, c.cand_fnam_t, c.cand_thisorg_short_i_"& pv_new_sessioncode &" AS thisorg_short, c.cand_thisorg_jobtype_i_"& pv_new_sessioncode &" AS thisorg_type, c.cand_geo_ctywork_m, c.cand_io_i, c.cand_io_t, c.cand_io_ds_t, c.cand_io_year_n, c.cand_io_grade_t, c.cand_io_m, c.cand_io_f_i, c.cand_io_f1_t, c.cand_io_f1_ds_t, c.cand_io_f1_year_n, c.cand_io_f1_year_end_n, c.cand_io_f1_grade_t, c.cand_io_f2_t, c.cand_io_f2_ds_t, c.cand_io_f2_year_n, c.cand_io_f2_year_end_n, c.cand_io_f2_grade_t, c.cand_io_f3_t, c.cand_io_f3_ds_t, c.cand_io_f3_year_n, c.cand_io_f3_year_end_n, c.cand_io_f3_grade_t, c.cand_geo_exp_i, c.cand_geo_res_i, c.cand_geo_res_m, d.drop_dsc_"& new_lng_code &"_t AS georesdsc, c.cand_retireepension_i, c.cand_security_certification FROM         dbo.td_rsys_cand c LEFT OUTER JOIN dbo.tr_rsys_drop d ON c.cand_geo_res_i = d.drop_id_c WHERE     (c.cand_id_c = "& applicant_id &")"
''d set JAPINFOIE =rsys_db_select.execute(JAPINFOIEsql)
			Set JAPINFOIE = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPINFOIE.Open JAPINFOIEsql, rsys_db_select, 1, 1

'10 OCT 10 LJL changed geo list to be org specific names
			JAPGEOLISTsql = "SELECT tx_rsys_candgeo.candgeo_id_c, tx_rsys_candgeo.candgeo_cand_c, 	g.geoloc_"& new_lng_code &"_" & pv_new_sessioncode & " as geolocdsc 	FROM tx_rsys_candgeo INNER JOIN 	core_geolocf g ON 	tx_rsys_candgeo.candgeo_id_c = g.geoloc_id_c 		WHERE tx_rsys_candgeo.candgeo_cand_c = " & applicant_id & " AND tx_rsys_candgeo.candgeo_type_i  = 0 "
''d set JAPGEOLIST =rsys_db_select.execute(JAPGEOLISTsql)
			Set JAPGEOLIST = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPGEOLIST.Open JAPGEOLISTsql, rsys_db_select, 1, 1
			
'Added by Atul
f.WriteLine "<p><br></p>"
			f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTF("i_3") & "</font></h2>"
			
			f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
' Atul make  space in between th tag abd  c tag-->
'f.WriteLine "<tr><th colspan=""4"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTF("i_3") & "_|" & intNumber  &  "'>" & gITEXTF("i_3") &  "</th></tr>"
'Comented by atul
'f.WriteLine "<tr><td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTF("i_3") & "</font></h2></td></tr>"
			
' INTERNAL AND WHO 1000
			if   JAPINFO1("thisorg_short") > "0"  AND (pv_new_sessioncode = 1000) then ' OR pv_new_sessioncode = 1500 by Gevorg
		  		intcurrentsql = "{call erstp_rsys_cand_internal_current_employ_" & pv_new_sessioncode & "(" & applicant_id & ")}"
'25 feb 07 ac increase db timeout
				rsys_db_select.CommandTimeout = 320
	  			set intcurrent =rsys_db_select.execute(intcurrentsql)
				if intcurrent.eof = false then
'  '----------- get WHO employ if staff member --------------->
'Do while intcurrent.eof = false
					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top"">" & gITEXTF("i_5") & "</TD>"
					f.WriteLine "<td valign=""top"">" & gITEXTF("i_8") & "</TD>"
'15 JAN 07 LJL wasn't set to right text var
					f.WriteLine "<td valign=""top"">" & gITEXTF("i_19") & "</TD>"
					f.WriteLine "<td valign=""top"">" & gITEXTF("i_10") & "</TD>"
					f.WriteLine "</TR>"

					f.WriteLine "<TR>"
					f.WriteLine "<td valign=""top""><strong>" & intcurrent("unitacro") & "</strong></TD>"
					f.WriteLine "<td valign=""top""><strong>" & intcurrent("DutyStation") & "</strong></TD>"
					f.WriteLine "<td valign=""top"">"
					if   len(intcurrent("contract_start_date"))then
						f.WriteLine "<strong>" & day(intcurrent("contract_start_date")) & "-" & monthname(month(intcurrent("contract_start_date")),1) & "-" & year(intcurrent("contract_start_date")) & "</strong>"
					end if
					f.WriteLine "</TD>"
					f.WriteLine "<td valign=""top""><strong>" & rtrim(intcurrent("staff_category")) & "-" & rtrim(intcurrent("staff_grade")) & " " & intcurrent("contract_type") & "</strong></TD>"
					f.WriteLine "</tr>"
				end if
'intcurrent.movenext
'loop

'17 SEP 09 LJL revised to not show this for any orgs that are not connected via true staff list (added 9999 as placeholder, stopper) code moves to ELSE
			elseif JAPINFO1("thisorg_short") = "1" AND pv_new_sessioncode = 9999 then
				intcurrentsql = "{call erstp_rsys_cand_internal_current_employ_"& pv_new_sessioncode &"(" & applicant_id & ")}"
''d set intcurrent =rsys_db_select.execute(intcurrentsql)
				Set intcurrent = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
				rsys_db_select.CommandTimeout = 320
				intcurrent.Open intcurrentsql, rsys_db_select, 1, 1
				if intcurrent.eof = false then
' Do while intcurrent.eof = false
					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top"" colspan=""4"">" & gITEXTF("i_4") & "<strong>"
					If JAPINFOIE("cand_io_i") = 0 then
						f.WriteLine gITEXTF("i_7")
					Else
						f.WriteLine gITEXTF("i_6")
					End If
					f.WriteLine "</strong></td>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top"">" & gITEXTF("i_5") & "</TD>"
					f.WriteLine "<td valign=""top"">" & gITEXTF("i_8") & "</TD>"
					f.WriteLine "<td valign=""top"">" & gITEXTF("i_19") & "</TD>"
					f.WriteLine "<td valign=""top"">" & gITEXTF("i_10") & "</TD>"
					f.WriteLine "</TR>"

					f.WriteLine "<TR>"
					f.WriteLine "<td valign=""top""><strong>" & intcurrent("unitacro") & " " & pv_new_sessioncode & "</strong></TD>"
					f.WriteLine "<td valign=""top""><strong>" & intcurrent("DutyStation") & "</strong></TD>"
					f.WriteLine "<td valign=""top"">"
					If len(intcurrent("contract_start_date")) then
						f.WriteLine "<strong>" & day(intcurrent("contract_start_date")) & "-" & monthname(month(intcurrent("contract_start_date")), 1) & "-" & right(year(intcurrent("contract_start_date")), 2) & "</strong>"
					End If
					f.WriteLine "</TD>"
					f.WriteLine "<td valign=""top""><strong>" & trim(intcurrent("staff_category")) & "-" & trim(intcurrent("staff_grade")) & " " & trim(intcurrent("contract_type")) & "</strong></TD>"
					f.WriteLine "</tr>"

'intcurrent.movenext
'loop
				end if
			else
				f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTF("i_4") & "<strong> "
				If JAPINFOIE("cand_io_i") = 0 then
					f.WriteLine gITEXTF("i_7")
				Else
					f.WriteLine gITEXTF("i_6")
				End If
				f.WriteLine "</strong></TD></tr>"
				if JAPINFO1("cand_io_i") = 1  then
					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""><I>" & gITEXTF("i_5") & "</I></td>"
					f.WriteLine "<td valign=""top""><I>" & gITEXTF("i_8") & "</I></td>"
					f.WriteLine "<td valign=""top""><I>" & gITEXTF("i_19") & "</I></td>"
					f.WriteLine "<td valign=""top""><I>" & gITEXTF("i_10") & "</I></td>"
					f.WriteLine "</tr>"

					f.WriteLine "<tr>"
					f.WriteLine "<td valign=""top""><b>" & JAPINFO1("cand_io_t") & "</b></TD>"
					f.WriteLine "<td valign=""top""><b>" & JAPINFO1("cand_io_ds_t") & "</b></TD>"
					f.WriteLine "<td valign=""top"">"
					if  JAPINFO1("cand_io_year_n") = "0" OR JAPINFO1("cand_io_year_n") = ""  then
					else
						f.WriteLine "<b>" & JAPINFO1("cand_io_year_n") & "</b>"
					end if
					f.WriteLine "</TD>"
					f.WriteLine "<td valign=""top""><b>" & JAPINFO1("cand_io_grade_t") & "</b></TD></tr>"
				end if
			end if


			if pv_new_sessioncode = 2400 then
				f.WriteLine "<tr>"
				f.WriteLine "<td valign=""top"" colspan=""4"">"  & gITEXTF("i_93") & "<strong> "
				If JAPINFOIE("cand_retireepension_i") = 0 then
					f.WriteLine gITEXTF("i_7")
				Else
					f.WriteLine gITEXTF("i_6")
				End If
				f.WriteLine "</strong></TD>"
				f.WriteLine "</tr>"
			end if


			f.WriteLine "<tr><td valign=""top"" colspan=""4"">" & gITEXTF("i_11") & " <strong> "
			If JAPINFOIE("cand_io_f_i") = 0 then
				f.WriteLine gITEXTF("i_7")
			else
				f.WriteLine gITEXTF("i_6")
			end if
			f.WriteLine "</strong></td></tr>"
			
			If JAPINFOIE("cand_io_f_i") = 0 and pv_new_sessioncode = 2900 then
				
			else
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"">" & gITEXTF("i_5") & "</TD>"
				f.WriteLine "<td valign=""top"">" & gITEXTF("i_8") & "</TD>"
				f.WriteLine "<td valign=""top"">" & gITEXTF("i_19") & " - " & gITEXTF("i_20") & "</TD>"
				f.WriteLine "<td valign=""top"">" & gITEXTF("i_10") & "</TD>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f1_t") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f1_ds_t") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f1_year_n") & " - " & JAPINFOIE("cand_io_f1_year_end_n") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f1_grade_t") & "</strong></TD>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				
				f.WriteLine "<td valign=""top""><strong>"
				StripSpecialChar(JAPINFOIE("cand_io_f2_t")) 
				f.WriteLine"</strong></TD>"
				
				f.WriteLine "<td valign=""top""><strong>" 
				StripSpecialChar(JAPINFOIE("cand_io_f2_ds_t")) 
				f.WriteLine"</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f2_year_n") & " - " & JAPINFOIE("cand_io_f2_year_end_n") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f2_grade_t") & "</strong></TD>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f3_t") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f3_ds_t") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f3_year_n") & " - " & JAPINFOIE("cand_io_f3_year_end_n") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPINFOIE("cand_io_f3_grade_t") & "</strong></TD>"
				f.WriteLine "</TR>"
				
				'------------ Additional information on international employment -------------------------------------------------->
				if len(JAPINFOIE("cand_io_m")) then
					f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTF("i_59") & " <strong>" & JAPINFOIE("cand_io_m") & "</strong></TD></TR>"
				end if
			end If
'          
'<TR>
'	<td valign=""top"" colspan=""5"">" & gITEXTF("i_12")></td>
'</tr>
			f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTF("i_13") & " <strong> "
			If trim(JAPINFOIE("cand_geo_exp_i")) = "1" then
				f.WriteLine gITEXTF("i_6")
			else
				f.WriteLine gITEXTF("i_7")
			End If
			'f.WriteLine "  JAPINFOIE||2"  'commented by Gevorg
'commented by at due to html data issue  f.WriteLine JAPINFOFOIE("cand_geo_exp_i") 
			f.WriteLine "</strong></td></TR>"
			
			'11 JUN 15 GG Added countries for Geographical Experience
			If JAPGEOLIST.eof = true then
				'f.WriteLine "<TR><td valign=""top"" colspan=""4""> <strong>" & gITEXTF("i_21") & " </strong> </td></TR>"
			Else
				f.WriteLine "<TR><td valign=""top"" colspan=""4"">"
				Do while JAPGEOLIST.eof = false
					f.WriteLine "<strong>" & JAPGEOLIST("geolocdsc") & "</strong><br>"
					JAPGEOLIST.movenext
				loop
				
				f.WriteLine "</td></TR>"
			End If
			
'01 JUN 15 LJL if international experience is yes, then show that additional info text
If trim(JAPINFOIE("cand_geo_exp_i")) = "1" then			
				
			f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTF("i_69") & " <strong>" & JAPINFOIE("cand_geo_ctywork_m") & "</strong></TD></TR>"
'<tr>
'	<td valign=""top"" colspan=""5"">" & gITEXTF("i_88")></td>
'</tr>

end if

			f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTF("i_17") & " <strong>" & JAPINFOIE("georesdsc") & "</strong></TD></TR>"
			f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTF("i_18") & " <strong>" & JAPINFOIE("cand_geo_res_m") & "</strong></TD></TR>"

			if pv_new_sessioncode = 2400 then
				f.WriteLine "<tr>"
				f.WriteLine "<td valign=""top"" colspan=""4"">"  & gITEXTF("i_text4") & "<strong> "
				If JAPINFOIE("cand_security_certification") = 0 then
					f.WriteLine gITEXTF("i_7")
				Else
					f.WriteLine gITEXTF("i_6")
				End If
				f.WriteLine "</strong></TD>"
				f.WriteLine "</tr>"
			end if


			f.WriteLine "</table>"
' END NOT 3000 and not 7000
		end if
' END CHECK IF PERSON SELECTED SECTION F
	end if


' ******************************************************************************
' SECTION F - EMPLOYMENT HISTORY LINES
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION F

		if instr(pv_parts,"F,") then
''-----------------------------------------NOTES ----------
'MODS --
'11 MAY 05 LJL added may we contact your employer
'31 MAY 05 LJL changed may we contact employer to 1=yes, 0=not yet indicated, 2=no
'--------------------------------------------------------------->
		JAPEMP0sql = "{call erstp_cand_full_emplist_" & new_lng_code & "_param(" & applicant_id & ")}"
'25 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		set JAPEMP0 =rsys_db_select.execute(JAPEMP0sql)
		pv_pagecount = JAPEMP0(0)
		set JAPEMP = JAPEMP0.NextrecordSet()
'Add by Atul
f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTF("i_24") & "</font></h2>"
		
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'f.WriteLine "<TR>"
' make  space in between th and a-->	
'f.WriteLine "<th colspan=""4"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTF("i_24") & "_|" & intNumber  &  "'>" & gITEXTF("i_24") &  "</a></th>"
'f.WriteLine "</TR>"
		
'Commented by atul
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTF("i_24") & "</font></h2></td>"
'f.WriteLine "</TR>"

		if pv_new_sessioncode <> 3000 then
			f.WriteLine "<TR><td valign=""top"" colspan=""4"" class=""teal10""><i><b>" & pv_pagecount & " " & gITEXTF("i_25")
			f.WriteLine "</b></i></TD></tr>"
		end if
		if   JAPEMP.eof = true  then
			f.WriteLine "<TR><td valign=""top"" colspan=""4""><FONT color=""Maroon"">" & gITEXTF("i_26") & "</font></TD></tr>"
		end if
		
	pv_edulinecount = 0
		Do while JAPEMP.eof = false			

''------ insertion ------->		
'17 JUN 16	GG Added colored line 
				f.WriteLine "<TR><td valign=""top"" colspan=""4"">"
				if pv_edulinecount > 0 then	
					f.WriteLine "<span><table border=""1"" cellpadding=""0"" cellspacing=""0"" bordercolor=""red""><tr><td>&nbsp;</td></tr></table></span>"
				end if
                 
				f.WriteLine "<table border=""0"" cellpadding=""1""><tr>"
				f.WriteLine "<td width=""30%""  valign=""top"" ><FONT color=""Teal""><i>" & gITEXTF("i_27") & "</I></font></td>"
				f.WriteLine "<td  width=""68%"" valign=""top""><b>" & JAPEMP("candemploy_title_t") & "</b></td>"
				f.WriteLine "<td width=""1%"" valign=""top"" ></td>"
				'f.WriteLine "<td align=""left"" width=""1%"" valign=""top""></td>"
				f.WriteLine "</tr></table></td></TR>"
				
				
				
		'11 JUN 15 GG Added EMPLOYMENT HISTORY Areas of Expertise for WMO	
			if session("template_org_code") = 2900 then
				dim obj_db_select_CmdXIII, obj_db_select_CmdXIV, obj_db_select_CmdV, obj_db_select_CmdXV, obj_db_select_CmdVI
				Dim EMPCCOG1sql, EMPCCOG2sql, EMPCCOG3sql, GetCCOGsql, GetCCOGRanksql
				Dim GetCCOG, GetCCOGRank, EMPCCOG1, EMPCCOG2, EMPCCOG3, candemploy_uid_c
				
				candemploy_uid_c = JAPEMP("candemploy_uid_c")
				
				
				set obj_db_select_CmdXIII = server.CreateObject("adodb.command")
				obj_db_select_CmdXIII.ActiveConnection = rsys_db_select
				'<!------------------- POST-ASP MODIF 12 AUG 05 LJL --------------------------------------->
				'<--Modified by Interface on 04/26/2007
				EMPCCOG1sql = "SELECT candccog_id_c, ccog_id_c, candccog_percentage_i, candccog_rank_i, candccog_dsc_t FROM tx_rsys_cand_ccog WHERE cand_id_c = ? AND candccog_type_i = 31 AND candccog_thisorg_" & session("template_org_code") & " = 1 AND candccog_candemploy_uid_c = ?"
				obj_db_select_CmdXIII.CommandText = EMPCCOG1sql
				Set EMPCCOG1 = obj_db_select_CmdXIII.Execute(,Array(session("RSYS_EVAL"),candemploy_uid_c))
				'-->>
				
				set obj_db_select_CmdV = server.CreateObject("adodb.command")
				obj_db_select_CmdV.ActiveConnection = rsys_db_select
				GetCCOGsql = "	SELECT ccog_dsc_" & session("lng") & "_t as ccogdsc, ccog_id_c, ccog_typ_c 	FROM core_ccogtblt 	WHERE ccog_subtype_i IN (1,3) AND ccog_thisorg_" & session("template_org_code") & " = 1 	ORDER BY 1"
				obj_db_select_CmdV.CommandText = GetCCOGsql
				set GetCCOG =obj_db_select_CmdV.execute()
				
				
				set obj_db_select_CmdVI = server.CreateObject("adodb.command")
				obj_db_select_CmdVI.ActiveConnection = rsys_db_select
				GetCCOGRanksql = "	SELECT ccogrank_dsc_"& session("lng") &"_t as ccogrankdsc, ccogrank_id_c FROM tr_rsys_ccogrank WHERE ccogrank_thisorg_" & session("template_org_code") & " = 1 ORDER BY 2"
				obj_db_select_CmdVI.CommandText = GetCCOGRanksql
				set GetCCOGRank =obj_db_select_CmdVI.execute()
				
				
				dim empccog1id, empccog1perc, empccog1cand, empccog1rank, empccog1dsc
				if EMPCCOG1.eof = true then
					empccog1id = ""
					empccog1perc = "0"
					empccog1cand = ""
					empccog1rank = "0"
					empccog1dsc = ""
				 else
					empccog1id = EMPCCOG1("ccog_id_c")
					empccog1perc = EMPCCOG1("candccog_percentage_i")
					empccog1cand = EMPCCOG1("candccog_id_c")
					empccog1rank = EMPCCOG1("candccog_rank_i")
					empccog1dsc = EMPCCOG1("candccog_dsc_t")
				 end if
				 
				set obj_db_select_CmdXIV = server.CreateObject("adodb.command")
				obj_db_select_CmdXIV.ActiveConnection = rsys_db_select

				'<--Modified by Interface on 04/26/2007
				 EMPCCOG2sql = "SELECT candccog_id_c, ccog_id_c, candccog_percentage_i, candccog_rank_i, candccog_dsc_t FROM tx_rsys_cand_ccog 	WHERE cand_id_c = ? AND candccog_type_i = 32 AND candccog_thisorg_" & session("template_org_code") & " = 1 AND candccog_candemploy_uid_c = ? "
				 obj_db_select_CmdXIV.CommandText = EMPCCOG2sql
				 Set EMPCCOG2 = obj_db_select_CmdXIV.Execute(,Array(session("RSYS_EVAL"), candemploy_uid_c))
				'-->>
				dim empccog2id, empccog2perc, empccog2cand, empccog2rank, empccog2dsc
				if EMPCCOG2.eof = true then
					empccog2id = ""
					empccog2perc = "0"
					empccog2cand = ""
					empccog2rank = "0"
					empccog2dsc = ""
				 else
					empccog2id = EMPCCOG2("ccog_id_c")
					empccog2perc = EMPCCOG2("candccog_percentage_i")
					empccog2cand = EMPCCOG2("candccog_id_c")
					empccog2rank = EMPCCOG2("candccog_rank_i")
					empccog2dsc = EMPCCOG2("candccog_dsc_t")
				 end if

					set obj_db_select_CmdXV = server.CreateObject("adodb.command")
					obj_db_select_CmdXV.ActiveConnection = rsys_db_select

				'<--Modified by Interface on 04/26/2007
					EMPCCOG3sql = "SELECT candccog_id_c, ccog_id_c, candccog_percentage_i, candccog_rank_i, candccog_dsc_t FROM tx_rsys_cand_ccog WHERE cand_id_c = ? AND candccog_type_i = 33 AND candccog_thisorg_" & session("template_org_code") & " = 1 AND candccog_candemploy_uid_c = ? "
					obj_db_select_CmdXV.CommandText = EMPCCOG3sql
					Set EMPCCOG3 = obj_db_select_CmdXV.Execute(,Array(session("RSYS_EVAL"), candemploy_uid_c))
				'-->>
				dim empccog3id, empccog3perc, empccog3cand, empccog3rank, empccog3dsc
				 if EMPCCOG3.eof = true then
					empccog3id = ""
					empccog3perc = "0"
					empccog3cand = ""
					empccog3rank = "0"
					empccog3dsc = ""
				  else
					empccog3id = EMPCCOG3("ccog_id_c")
					empccog3perc = EMPCCOG3("candccog_percentage_i")
					empccog3cand = EMPCCOG3("candccog_id_c")
					empccog3rank = EMPCCOG3("candccog_rank_i")
					empccog3dsc = EMPCCOG3("candccog_dsc_t")
				  end if
				
				f.WriteLine "<TR>"
				If EMPCCOG1id <> ""  and EMPCCOG1id <> "0" Then	
					GETCCOG.movefirst
					Do while GETCCOG.eof = false
						if GETCCOG("ccog_id_c") = EMPCCOG1id then
							f.WriteLine "<td valign=""top"" ><b>" & GetCCOG("ccogdsc") & "</b></TD>" 
						end if	
						GETCCOG.movenext
					Loop
				Else 
					f.WriteLine "<td ></TD>" 
				End if
				If EMPCCOG1rank <> "0" Then	
					GetCCOGRank.movefirst
					Do while GetCCOGRank.eof = false
						if GetCCOGRank("ccogrank_id_c") = EMPCCOG1rank then
							f.WriteLine "<td valign=""top"" colspan=""1""><b>" & GetCCOGRank("ccogrankdsc") & "</b></TD>" 
						end if	
						GetCCOGRank.movenext
					Loop
				Else 
					f.WriteLine "<td ></TD>" 
				End if
				
				If EMPCCOG1dsc <> "" Then	
					f.WriteLine "<td valign=""top"" colspan=""1""><b>" & EMPCCOG1dsc & "</b></TD>"	
				Else 
					f.WriteLine "<td ></TD>"
				End if
				If EMPCCOG1perc <> "0" Then					
					f.WriteLine "<td valign=""top"" colspan=""1""><b>" & EMPCCOG1perc & "% </b></TD>" 
				Else 
					f.WriteLine "<td ></TD>"
				End if					
				f.WriteLine "</TR>"
				
				f.WriteLine "<TR>"
				If EMPCCOG2id <> ""  and EMPCCOG2id <> "0" Then	
					GETCCOG.movefirst				
					Do while GETCCOG.eof = false
						if GETCCOG("ccog_id_c") = EMPCCOG2id then
							f.WriteLine "<td valign=""top"" ><b>" & GetCCOG("ccogdsc") & "</b></TD>" 
						end if	
						GETCCOG.movenext
					Loop
				Else 
					f.WriteLine "<td></TD>" 
				End if
				If EMPCCOG2rank <> "0" Then
					GetCCOGRank.movefirst
					Do while GetCCOGRank.eof = false
						if GetCCOGRank("ccogrank_id_c") = EMPCCOG2rank then
							f.WriteLine "<td valign=""top"" colspan=""1""><b>" & GetCCOGRank("ccogrankdsc") & "</b></TD>" 
						end if	
						GetCCOGRank.movenext
					Loop
				Else 
					f.WriteLine "<td ></TD>" 
				End if
				
				If EMPCCOG2dsc <> "" Then	
					f.WriteLine "<td valign=""top"" colspan=""1""><b>" & EMPCCOG2dsc & "</b></TD>"	
				Else 
					f.WriteLine "<td ></TD>"
				End if
				If EMPCCOG2perc <> "0" Then					
					f.WriteLine "<td valign=""top"" colspan=""1""><b>" & EMPCCOG2perc & "% </b></TD>" 
				Else 
					f.WriteLine "<td ></TD>"
				End if					
				f.WriteLine "</TR>"
				
				If EMPCCOG3.eof <> true Then
					f.WriteLine "<TR>"
					If EMPCCOG3id <> "" and EMPCCOG3id <> "0" Then
						GETCCOG.movefirst
						Do while GETCCOG.eof = false
							if GETCCOG("ccog_id_c") = EMPCCOG3id then
								f.WriteLine "<td valign=""top"" ><b>" & GetCCOG("ccogdsc") & "</b></TD>" 
							end if	
							GETCCOG.movenext
						Loop
					Else 
						f.WriteLine "<td></TD>" 
					End if
					If EMPCCOG3rank <> "0" Then	
						GetCCOGRank.movefirst
						Do while GetCCOGRank.eof = false
							if GetCCOGRank("ccogrank_id_c") = EMPCCOG3rank then
								f.WriteLine "<td valign=""top"" colspan=""1""><b>" & GetCCOGRank("ccogrankdsc") & "</b></TD>" 
							end if	
							GetCCOGRank.movenext
						Loop
					Else 
						f.WriteLine "<td></TD>" 
					End if
					
					If EMPCCOG3dsc <> "" Then	
						f.WriteLine "<td valign=""top"" colspan=""1""><b>" & EMPCCOG3dsc & "</b></TD>"	
					Else 
						f.WriteLine "<td ></TD>"
					End if
					If EMPCCOG3perc <> "0" Then					
						f.WriteLine "<td valign=""top"" colspan=""1""><b>" & EMPCCOG3perc & "% </b></TD>" 
					Else 
						f.WriteLine "<td></TD>"
					End if					
					f.WriteLine "</TR>"
				End if
				
			 end if		
				
				
				f.WriteLine "<TR><td valign=""top"" colspan=""4""><table border=""0"" cellpadding=""1""><tr>"
				f.WriteLine "<td width=""15%"" valign=""top"" ><i>" & gITEXTF("i_29") & "</I></td>"
				f.WriteLine "<td align=""left"" width=""40%"" valign=""top"" ><b>" & replace(JAPEMP("candemploy_add_m"),VbCrLf,"<br>")  & "</b></td>"
				f.WriteLine "<td width=""15%"" valign=""top"" ><i>" & gITEXTF("i_28") & "</I></td>"
				f.WriteLine "<td align=""left"" width=""30%"" valign=""top"" ><b>" & JAPEMP("candemploy_sup_t") & "</b></td>"				
				f.WriteLine "</tr></table></td></TR>"
				
				f.WriteLine "<TR><td valign=""top"" colspan=""4""><table border=""0"" cellpadding=""1""><tr>"
				f.WriteLine "<td valign=""top""  width=""15%""><i>" & gITEXTF("i_19") & "</i>-<i>" & gITEXTF("i_20") & "</I></td>"
				f.WriteLine "<td valign=""top"" width=""83%""><b>"
				if   JAPEMP("candemploy_start_month_d") > 12 OR JAPEMP("candemploy_start_year_d") = 9999  then
					f.WriteLine gITEXTF("i_53")
				else
					f.WriteLine JAPEMP("candemploy_start_month_d") & "/" & JAPEMP("candemploy_start_year_d")
				end if
				f.WriteLine " - "

				if   JAPEMP("candemploy_end_month_d") > 12 OR JAPEMP("candemploy_end_year_d") = 9999  then
					f.WriteLine gITEXTF("i_53")
				else
					f.WriteLine JAPEMP("candemploy_end_month_d") & "/" & JAPEMP("candemploy_end_year_d")
				end if
				f.WriteLine "</b></td><td width=""1%"" valign=""top"" ></td><td width=""1%"" valign=""top""></td>"
				f.WriteLine "</tr></table></td></TR>"

			if pv_new_sessioncode <> 7000 then ' NOT FOR IFRC SALARY
				f.WriteLine "<TR><td valign=""top"" colspan=""4"">"& " " & gITEXTF("i_45") & " " & gITEXTF("i_56") & " " & gITEXTF("i_63") & "<table  border=""0"" cellpadding=""1""><tr>"
				f.WriteLine "<td width=""15%"" valign=""top"" ><i>" & gITEXTF("i_60")  & "</I></td>"
				f.WriteLine "<td align=""left"" width=""40%"" valign=""top"" ><b>" & JAPEMP("candemploy_sal1_c") & "</b></td>"
				f.WriteLine "<td width=""15%"" valign=""top"" ><i>" & gITEXTF("i_62") & "</I></td>"
				f.WriteLine "<td align=""left"" width=""30%"" valign=""top"" ><b>" & JAPEMP("candemploy_sal2_c") & "</b></td>"
				f.WriteLine "</tr></table></td></TR>"
			end if


'23 Nov 07 ac - only diplay country for IFRC and added i_89 Text "city / Country location
				if   pv_new_sessioncode = 7000 then
'20 Nov 07 ac - add employment country location
					if len(JAPEMP("candemploy_cty_c")) then
						GetEmploymentCountrySql = "SELECT c_"& session("lng")&"_name AS ctyname "
						GetEmploymentCountrySql = GetEmploymentCountrySql & " FROM dbo.v_country_list_" & pv_new_sessioncode
						GetEmploymentCountrySql = GetEmploymentCountrySql & " WHERE (NOT (c_"& session("lng") &"_name IS NULL))"
						GetEmploymentCountrySql = GetEmploymentCountrySql & " and who_country_code = '" & JAPEMP("candemploy_cty_c") & "'"
'rsys_db_select.CommandTimeout = 320
						set GetEmploymentCountry = rsys_db_select1.execute(GetEmploymentCountrySql)
''d set GetDoc2 =rsys_db_select.execute(GetDoc2sql)
'	Set GetEmploymentCountry = Server.CreateObject("ADODB.RecordSet")
'	GetEmploymentCountry.Open GetEmploymentCountrySql, rsys_db_select, 1, 1
						f.WriteLine "<TR>"
						f.WriteLine "<td valign=""top"" colspan=""1""><i>" & gITEXTF("i_88") & "</I></td>"
						f.WriteLine "<td valign=""top"" colspan=""3""><b>" & GetEmploymentCountry("ctyname")
						f.WriteLine "</b></td></TR>"
					end if
				end if
' Atul Commnet tr  becuase no need here  -->	
'f.WriteLine "<TR>"
'f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"" colspan=""4""><i>" & gITEXTF("i_30") & "</I></td>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"" colspan=""4""><strong>"
				if len(JAPEMP("candemploy_dsc_m")) then
					f.WriteLine replace(JAPEMP("candemploy_dsc_m"),VbCrLf,"<br>")
				end if
				f.WriteLine "</strong></td>"
				f.WriteLine "</tr>"
			if   len(JAPEMP("candemploy_keywork_m")) then
				f.WriteLine "<tr>"
				f.WriteLine "<td valign=""top"" colspan=""4""><i>" & gITEXTF("i_68") & "</I></td>"
				f.WriteLine "</tr>"
				f.WriteLine "<tr>"
				f.WriteLine "<td valign=""top"" colspan=""4""><strong>"
				if len(JAPEMP("candemploy_keywork_m")) then
					f.WriteLine replace(JAPEMP("candemploy_keywork_m"),VbCrLf,"<br>")
				end if
				f.WriteLine "</strong></td>"
				f.WriteLine "</tr>"
			end if

			f.WriteLine "<TR><td valign=""top"" colspan=""4""><table border=""0"" cellpadding=""0""><tr>"
			f.WriteLine "<td width=""40%"" valign=""top""><i>" & gITEXTF("i_73") & "</I></td>"
			f.WriteLine "<td align=""left"" width=""5%"" valign=""top""><b>" & JAPEMP("candemploy_supervise_c") & "</b></td>"
			f.WriteLine "<td width=""35%"" valign=""top"" class=""navy9""><i>" & gITEXTF("i_82") & "</I></td>"
			f.WriteLine "<td align=""left"" width=""20%"" valign=""top"" class=""navy9""><b>" & JAPEMP("dropdsc") & "</b></td>"
			f.WriteLine "</tr></table></td></TR>"			
			f.WriteLine "<TR>"
			f.WriteLine "<td valign=""top"" colspan=""1""><i>" & gITEXTF("i_74") & "</I></td>"
			f.WriteLine "<td valign=""top"" colspan=""3""><b>" & JAPEMP("candemploy_whyleave_t") & "</b></td>"
			f.WriteLine "</TR>"
'01 SEP 10 LJL remove first horizontal line if first record
			JAPEMP.movenext
			pv_edulinecount = pv_edulinecount + 1
			loop
'f.WriteLine "<TR><td valign=""top"" colspan=""4""><hr noshade size=""1""></TD></tr>"
' '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
		if   Request.querystring("viewupd") = True and viewupd = "YES"  then
			if   GETUPDS("editF_d") > GETAPPDATE("candjob_d")  then
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") & """>Last update date <strong>" & formatdatetime(GETUPDS("EditF_d"),1) & " AFTER APPLYING</strong></td>"
		f.WriteLine "</TR>"
		  end if
		end if
		f.WriteLine "</table>"
	
' END CHECK IF PERSON SELECTED SECTION F
		end if

		

' ******************************************************************************
' SECTION RC - RC/RC EXPERIENCE IFRC ONLY
' ******************************************************************************

		if pv_new_sessioncode = 7000 then
' BEGIN CHECK IF PERSON SELECTED SECTION RC
		if instr(pv_parts,"RC,") then

		JAPINFORCsql = " SELECT     c.cand_employ3_pres_i, c.cand_employ3_prev_i, c.cand_employ4_pres_i, c.cand_employ4_prev_i, c.cand_employ5_pres_i, c.cand_employ5_prev_i, c.cand_employ4_when_c, c.upd_d, c.cand_lnam_t, c.cand_fnam_t, c.cand_io_m, c.cand_employ_office_c, c.cand_thisorg_prev_i_" & pv_new_sessioncode & " AS thisorg_prev, c.cand_employ1_pres_i, c.cand_employ1_prev_i, c.cand_employ2_pres_i, c.cand_employ2_prev_i, c.cand_employ4_where_c, c.cand_thisorg_short_i_" & pv_new_sessioncode & " AS thisorg_short, e.extorg_dsc_" & new_lng_code & "_t AS office_name FROM         dbo.td_rsys_cand c LEFT OUTER JOIN rsys_int.dbo.extorg_" & pv_new_sessioncode & " e ON c.cand_employ_office_c = e.extorg_id_c WHERE     (c.cand_id_c = " & applicant_id & ")"
''d set JAPINFORC =rsys_db_select.execute(JAPINFORCsql)
		Set JAPINFORC = Server.CreateObject("ADODB.RecordSet")
'25 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPINFORC.Open JAPINFORCsql, rsys_db_select, 1, 1
		JAPOFFICE1sql = "SELECT gl.candgeo_office_c, gl.candgeo_cand_c, g."& new_lng_code &"_dsc1 as officename, g.u1 FROM tx_rsys_candgeo gl INNER JOIN v_orgext_list_" & pv_new_sessioncode & " g ON 	gl.candgeo_office_c = g.u1 WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 91 "
''d set JAPOFFICE1 =rsys_db_select.execute(JAPOFFICE1sql)
		Set JAPOFFICE1 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPOFFICE1.Open JAPOFFICE1sql, rsys_db_select, 1, 1

		JAPINSTLIST11sql = "SELECT gl.candgeo_ins_c, gl.candgeo_cand_c, 	g.ins_dsc_"& new_lng_code &"_t as insname FROM tx_rsys_candgeo gl INNER JOIN td_inst_institutions g ON 	gl.candgeo_ins_c = g.ins_id_c WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 11 "
''d set JAPINSTLIST11 =rsys_db_select.execute(JAPINSTLIST11sql)
		Set JAPINSTLIST11 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPINSTLIST11.Open JAPINSTLIST11sql, rsys_db_select, 1, 1

		JAPINSTLIST12sql = "SELECT gl.candgeo_ins_c, gl.candgeo_cand_c, 	g.ins_dsc_"& new_lng_code &"_t as insname FROM tx_rsys_candgeo gl INNER JOIN td_inst_institutions g ON 	gl.candgeo_ins_c = g.ins_id_c WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 12 "
''d set JAPINSTLIST12 =rsys_db_select.execute(JAPINSTLIST12sql)
		Set JAPINSTLIST12 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPINSTLIST12.Open JAPINSTLIST12sql, rsys_db_select, 1, 1

		JAPGEOLIST21sql = "SELECT gl.candgeo_cty_c, gl.candgeo_cand_c, 	g.c_"& new_lng_code &"_name as ctyname FROM tx_rsys_candgeo gl INNER JOIN v_country_list_" & pv_new_sessioncode & " g ON 	gl.candgeo_cty_c = g. cty_uid_c WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 21 "
''d set JAPGEOLIST21 =rsys_db_select.execute(JAPGEOLIST21sql)
		Set JAPGEOLIST21 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPGEOLIST21.Open JAPGEOLIST21sql, rsys_db_select, 1, 1

		JAPGEOLISTsql = "SELECT tx_rsys_candgeo.candgeo_id_c, tx_rsys_candgeo.candgeo_cand_c, 	g.geoloc_"& new_lng_code &"_" & pv_new_sessioncode & " as geolocdsc 	FROM tx_rsys_candgeo INNER JOIN 	core_geolocf g ON 	tx_rsys_candgeo.candgeo_id_c = g.geoloc_id_c 		WHERE tx_rsys_candgeo.candgeo_cand_c = " & applicant_id & " AND tx_rsys_candgeo.candgeo_type_i  = 0 "
''d set JAPGEOLIST =rsys_db_select.execute(JAPGEOLISTsql)
		Set JAPGEOLIST = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPGEOLIST.Open JAPGEOLISTsql, rsys_db_select, 1, 1

		JAPGEOLIST31sql = "SELECT gl.candgeo_cty_c, gl.candgeo_cand_c, 	g.c_"& new_lng_code &"_name as ctyname FROM tx_rsys_candgeo gl INNER JOIN v_country_list_" & pv_new_sessioncode & " g ON 	gl.candgeo_cty_c = g.cty_uid_c WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 31 "
''d set JAPGEOLIST31 =rsys_db_select.execute(JAPGEOLIST31sql)
		Set JAPGEOLIST31 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPGEOLIST31.Open JAPGEOLIST31sql, rsys_db_select, 1, 1

		JAPGEOLIST41sql = "SELECT gl.candgeo_cty_c, gl.candgeo_cand_c, 	g.c_"& new_lng_code &"_name as ctyname FROM tx_rsys_candgeo gl INNER JOIN v_country_list_" & pv_new_sessioncode & " g ON 	gl.candgeo_cty_c = g.cty_uid_c WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 41 "
''dset JAPGEOLIST41 =rsys_db_select.execute(JAPGEOLIST41sql)
		Set JAPGEOLIST41 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPGEOLIST41.Open JAPGEOLIST41sql, rsys_db_select, 1, 1

		JAPGEOLIST42sql = "SELECT gl.candgeo_cty_c, gl.candgeo_cand_c, 	g.c_"& new_lng_code &"_name as ctyname FROM tx_rsys_candgeo gl INNER JOIN v_country_list_" & pv_new_sessioncode & " g ON 	gl.candgeo_cty_c = g.cty_uid_c WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 42 "
''d set JAPGEOLIST42 =rsys_db_select.execute(JAPGEOLIST42sql)
		Set JAPGEOLIST42 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPGEOLIST42.Open JAPGEOLIST42sql, rsys_db_select, 1, 1
'Add y Atul
f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTRC("i_1") & "</font></h2>"
		
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'f.WriteLine "<TR>"
' Atul space in between th and a-->
'f.WriteLine "<th colspan=""4"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTRC("i_1") & "_|" & intNumber  &  "'>" & gITEXTRC("i_1") & "</a></th>"
'f.WriteLine "</TR>"
		
'Commented by atul
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") &"""><h2><font color=""" & g_headerColor & """>" & gITEXTRC("i_1") & "</font></h2></td>"
'f.WriteLine "</TR>"
		
		f.WriteLine "<tr>"
		f.WriteLine "<td valign=""top"">"
		f.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""0"" width=""" & widther & """>"
' CURRENTLY STAFF MEMBER?
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top""  colspan=""2"">" & gITEXTRC("i_2") & " <strong>"
		If JAPINFORC("thisorg_short") = "0" then
			f.WriteLine gITEXTRC("i_7")
		else
			f.WriteLine gITEXTRC("i_6")
		End If
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"

' CURRENTLY STAFF MEMBER - OFFICE ************
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top""  colspan=""2"">" & gITEXTRC("i_3") & " <strong>" & JAPINFORC("office_name") & "</strong></TD>"
		f.WriteLine "</TR>"
		
		
' start chnage by atul remove the hr  make line with tr and table-->
'f.WriteLine "<TR>"
'f.WriteLine "<td colspan=""2""><hr size=""1"" width=""100%""></td>"
'f.WriteLine "</TR>"
'f.WriteLine "<tr><td colspan=""2"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
		f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
		f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
		f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
' End chnage by atul remove the hr  make line with tr and table -->
		
' PREVIOUSLY WORKED FOR?
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top""  colspan=""2"">"
		f.WriteLine gITEXTRC("i_4") & " <strong>"
		If JAPINFORC("thisorg_prev") = "0" then
			f.WriteLine gITEXTRC("i_7")
		else
			f.WriteLine gITEXTRC("i_6")
		end if
		f.WriteLine "</strong>"
		f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"

' NAME OF OFFICES *******************
' Atul  make colspan=2 from 1-->
		f.WriteLine "<td align=""left"" valign=""middle"" colspan=""2"" >" & gITEXTRC("i_8") & "</TD>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" align=""left"" valign=""middle""  colspan=""1"">"
'f.WriteLine "<font color=""blue"">"
		If JAPOFFICE1.eof = true then
			f.WriteLine "<b>" & gITEXTRC("i_21") & "</b>"
		Else
			counter = 0
			if instr(pv_parts,"RCCountryList,") then
				Do while JAPOFFICE1.eof = false
				counter = counter + 1
				f.WriteLine "<td valign=""top"" align=""left"" valign=""middle""  colspan=""1"">"
				f.WriteLine "<font color=""blue"">"

					f.WriteLine "<img src""../../images/" & pv_new_sessioncode & "_admin/images/admin/arrow-blk.gif""" &_
			" width=""4"" height=""7"" border=""0"" alt="""" name=""in1"">" & UCASE(JAPOFFICE1("officename"))
				f.WriteLine "</font>"
				f.WriteLine "</td>"
				JAPOFFICE1.movenext

				if counter mod 2 = 0 then
						f.WriteLine "</tr>"
						f.WriteLine "<tr>"
				end if

				loop
			End If
		End If
		
' start chnage by atul remove the hr  make line with tr and table-->
'f.WriteLine "<tr><td colspan=""2""><hr noshade size=""1""></td></tr>"
'f.WriteLine "<tr><td colspan=""2"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
		f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
		f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
		f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
' End chnage by atul remove the hr  make line with tr and table -->
				
		
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" align=""left"" colspan=""2"" class='textbold'>" & gITEXTRC("i_9") & "</TD>"
		f.WriteLine "</TR>"
		
' start chnage by atul remove the hr  make line with tr and table-->
'f.WriteLine "<tr><td colspan=""2""><hr noshade size=""1""></td></tr>"
'f.WriteLine "<tr><td colspan=""2"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
		f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
		f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
		f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
' End chnage by atul remove the hr  make line with tr and table -->
		
' RC/RC NAT SOC?
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""middle""  colspan=""2"">" & gITEXTRC("i_10") & "</TD>"
		f.WriteLine "</TR>"
' Atul  make colspan=2 from 1-->
		f.WriteLine "<TR>"
		f.WriteLine "<td align=""left"" valign=""middle"" colspan=""2""><span class=""alert"">" & gITEXTRC("i_23") & "</span></td>"
		f.WriteLine "</TR>"
' Atul  make colspan=2 from 1-->
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""2"" >"
		f.WriteLine "<span class=""alert"">" & gITEXTRC("i_12") & "</span> <strong>"
		If JAPINFORC("cand_employ1_pres_i") = "1" then
			f.WriteLine gITEXTRC("i_6")
		else
			f.WriteLine gITEXTRC("i_7")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
' COUNTRIES PRES/PREV 11/12 *******************
		f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" align=""left"" valign=""top"" colspan=""1"">"
'f.WriteLine "<font color=""blue"">"
		if JAPINSTLIST11.eof = true then
			f.WriteLine "<b>" & gITEXTRC("i_21") & "</b>"
		Else
			counter = 0
			if instr(pv_parts,"RCCountryList,") then
				Do while JAPINSTLIST11.eof = false
					counter = counter + 1
					f.WriteLine "<td valign=""top"" align=""left"" valign=""top"" colspan=""1"">"
					f.WriteLine "<font color=""blue"">"
					f.WriteLine "<img src=""../../images/" & pv_new_sessioncode & "_admin/images/admin/arrow-blk.gif""" &_
							" width=""4"" height=""7""  border=""0"" alt="""" name=""in1"">" & UCASE(JAPINSTLIST11("insname"))
					f.WriteLine "</font><Br>"
					f.WriteLine "</td>"
					JAPINSTLIST11.movenext
					if counter mod 2 = 0 then
							f.WriteLine "</tr>"
							f.WriteLine "<tr>"
					end if

				loop
			End If
		End If
'f.WriteLine "</font><Br>"
'f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
' make colspan=2 from 1
		f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""2"" ><span class=""alert"">" & gITEXTRC("i_13") & "</span><strong>"
		if JAPINFORC("cand_employ1_prev_i") = "0" then
			f.WriteLine gITEXTRC("i_7")
		else
			f.WriteLine gITEXTRC("i_6")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" align=""left"" valign=""top"" colspan=""1"">"
'f.WriteLine "<font color=""blue"">"
		If JAPINSTLIST12.eof = true then
			f.WriteLine "<b>" & gITEXTRC("i_21") & "</b>"
		Else
			counter = 0
			if instr(pv_parts,"RCCountryList,") then
				Do while JAPINSTLIST12.eof = false
					counter = counter + 1
					f.WriteLine "<td valign=""top"" align=""left"" valign=""top"" colspan=""1"">"
					f.WriteLine "<font color=""blue"">"
					f.WriteLine "<img src=""../../images/"& pv_new_sessioncode & "_admin/images/admin/arrow-blk.gif""" &_
			" width=""4"" height=""7"" border=""0"" alt="""" name=""in1"" >" & UCASE(JAPINSTLIST12("insname"))
					f.WriteLine "</font>"
					f.WriteLine "</td>"
					JAPINSTLIST12.movenext

					if counter mod 2 = 0 then
							f.WriteLine "</tr>"
							f.WriteLine "<tr>"
					end if

				loop
			End If
		End If
'f.WriteLine "</font>"
'f.WriteLine "</td>"
		f.WriteLine "</TR>"
		
'start chnage by atul remove the hr  make line with tr and table
'f.WriteLine "<TR>"
'f.WriteLine "<td colspan=""2""><hr size=""1"" width=""100%""></td>"
'f.WriteLine "</TR>"
'f.WriteLine "<tr><td colspan=""2"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
		f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
		f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
		f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
'End chnage by atul remove the hr  make line with tr and table
		
' ICRC ?
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""2"">" & gITEXTRC("i_14") & "</TD>"
		f.WriteLine "</TR>"
' Atul  make colspan=2  form 1
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""2"" ><span class=""alert"">" & gITEXTRC("i_12") & "</span><strong>"
		If JAPINFORC("cand_employ2_pres_i") = "1" then
			f.WriteLine gITEXTRC("i_6")
		else
			f.WriteLine gITEXTRC("i_7")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
' Atul  make colspan=2  form 1
		f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""2"" ><span class=""alert"">" & gITEXTRC("i_13") & "</span><strong>"
		If JAPINFORC("cand_employ2_prev_i") = "0" then
			f.WriteLine gITEXTRC("i_7")
		else
			f.WriteLine gITEXTRC("i_6")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
' COUNTRIES PRES/PREV 21/22 *******************
		if instr(pv_parts,"RCCountryList,") then
			f.WriteLine "<TR>"
' Atul  make colspan=2  form 1
			f.WriteLine "<td align=""left"" valign=""middle"" colspan=""2""  ><span class=""alert"">" & gITEXTRC("i_11") & "</span></td>"
			f.WriteLine "</TR>"
		else
			f.WriteLine "<TR>"
' Atul  make colspan=2  form 1
			f.WriteLine "<td align=""left"" valign=""middle"" colspan=""2""  ><span class=""alert"">" & gITEXTRC("i_25") & "</span></td>"
			f.WriteLine "</TR>"
		end if
		f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" align=""left"" valign=""middle""  nowrap colspan=""1"">"
'f.WriteLine "<font color=""blue"">"
		If JAPGEOLIST21.eof = true then
			f.WriteLine "<b>" & gITEXTRC("i_21") & "</b>"
		Else
			counter = 0
			if instr(pv_parts,"RCCountryList,") then
				Do while JAPGEOLIST21.eof = false
					counter = counter + 1
					f.WriteLine "<td valign=""top"" align=""left"" valign=""middle""  nowrap colspan=""1"">"
					f.WriteLine "<font color=""blue"">"
					f.WriteLine "<img src=""../../images/" & pv_new_sessioncode & "_admin/images/admin/arrow-blk.gif""" &_
							" width=""4"" height=""7"" border=""0"" alt="""" name=""in1"">" & UCASE(JAPGEOLIST21("ctyname"))
					f.WriteLine "</font>"
					f.WriteLine "</td>"
					JAPGEOLIST21.movenext

					if counter mod 2 = 0 then
							f.WriteLine "</tr>"
							f.WriteLine "<tr>"
					end if

				loop
			End If
		End If
'f.WriteLine "</font>"
'f.WriteLine "<Br>"
		f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "</table>"
		f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "</table>"

		f.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"">"
		
		f.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""0"" border=""0"">"
'start chnage by atul remove the hr  make line with tr and table
'f.WriteLine "<TR>"
'f.WriteLine "<td colspan=""2""><hr size=""1"" width=""100%""></td>"
'f.WriteLine "</TR>"
'f.WriteLine "<tr><td colspan=""2"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
			f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
			f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
			f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
'End chnage by atul remove the hr  make line with tr and table
						
' RC VOLUNTEER ?
' Atul make colspan=""2"" from 1
			f.WriteLine "<TR><td colspan=""2"">" & gITEXTRC("i_15") & " <strong>"
			If JAPINFORC("cand_employ3_pres_i") = "1" then
				f.WriteLine  gITEXTRC("i_6")
			else
				f.WriteLine gITEXTRC("i_7")
			end if
			f.WriteLine "</strong>"
			f.WriteLine "</TD></tr>"

' COUNTRIES PRES/PREV 31/32 *******************
			if instr(pv_parts,"RCCountryList,") then
				f.WriteLine "<TR><td align=""left""  colspan=""2""><span class=""alert"">" & gITEXTRC("i_11") & "</span></TD></tr>"
			else
				f.WriteLine "<TR><td align=""left""  colspan=""2""><span class=""alert"">" & gITEXTRC("i_25") & "</span></TD></tr>"
			end if
'f.WriteLine "<TR><td align=""left"" colspan=""1"" width=""40%"">"
'f.WriteLine "<font color=""blue"">"
			If JAPGEOLIST31.eof = true then
				f.WriteLine "<b>" & gITEXTRC("i_21") & "</b>"
			Else
				counter = 0
				if instr(pv_parts,"RCCountryList,") then
					Do while JAPGEOLIST31.eof = false
						counter = counter + 1
						f.WriteLine "<td align=""left"" width=""40%"">"
						f.WriteLine "<font color=""blue"">"
						f.WriteLine "<img src=""../../images/" &  pv_new_sessioncode & "_admin/images/admin/arrow-blk.gif""" &_
				"width=""4"" height=""7"" border=""0""  alt="""" name=""in1"">" & UCASE(JAPGEOLIST31("ctyname"))
						f.WriteLine "</font>"
						f.WriteLine "</td>"
					JAPGEOLIST31.movenext
					if counter mod 2 = 0 then
							f.WriteLine "</tr>"
							f.WriteLine "<tr>"
					end if
					loop
				End If
			End If
'f.WriteLine "</font>"
			f.WriteLine "</TR>"
			
			
'start chnage by atul remove the hr  make line with tr and table
'f.WriteLine "<TR><td colspan=""2""><hr size=""1"" width=""100%""></TD></tr>"
'f.WriteLine "<tr><td colspan=""2"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
			f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
			f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
			f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
'End chnage by atul remove the hr  make line with tr and table

' BASIC TRAINING ?
		f.WriteLine "<tr>"
		f.WriteLine "<td>" & gITEXTRC("i_16") & " <strong>"
		If JAPINFORC("cand_employ4_pres_i") = "1" then
			f.WriteLine  gITEXTRC("i_6")
		else
			f.WriteLine gITEXTRC("i_7")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"

' COUNTRIES PRES 41 *******************
		if instr(pv_parts,"RCCountryList,") then
			f.WriteLine "<tr>"
			f.WriteLine "<td align=""left""  colspan=""2""><span class=""alert"">" & gITEXTRC("i_11") & "</span></td>"
			f.WriteLine "</TR>"
		else
			f.WriteLine "<tr>"
			f.WriteLine "<td align=""left""  colspan=""2""><span class=""alert"">" & gITEXTRC("i_25") & "</span></td>"
			f.WriteLine "</TR>"
		end if
		f.WriteLine "<tr>"
'f.WriteLine "<td align=""left"" colspan=""1"">"
'f.WriteLine "<font color=""blue"">"
		If JAPGEOLIST41.eof = true then
			f.WriteLine "<b>" & gITEXTRC("i_21") & "</b>"
		Else
			counter = 0
			if instr(pv_parts,"RCCountryList,") then
				Do while JAPGEOLIST41.eof = false
					counter = counter + 1
					f.WriteLine "<td align=""left"">"
					f.WriteLine "<font color=""blue"">"
					f.WriteLine "<img src=""../../images/" & pv_new_sessioncode & "_admin/images/admin/arrow-blk.gif""" &_
							" width=""4"" height=""7"" border=""0"" alt="""" name=""in1"">" & UCASE(JAPGEOLIST41("ctyname"))
					f.WriteLine "</font>"
					f.WriteLine "</td>"
					JAPGEOLIST41.movenext
					if counter mod 2 = 0 then
							f.WriteLine "</tr>"
							f.WriteLine "<tr>"
					end if
				loop
			End If
		End If
'f.WriteLine "</font>"
'f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<tr>"
		f.WriteLine "<td colspan=""2"">&nbsp;</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<tr>"
'Atul make  colspan=2 from 1
		f.WriteLine "<td colspan=""2"">" & gITEXTRC("i_17") & " <strong>" & JAPINFORC("cand_employ4_when_c") & "</strong></TD>"
		f.WriteLine "</TR>"
		
'start chnage by atul remove the hr  make line with tr and table
'f.WriteLine "<TR>"
'f.WriteLine "<td colspan=""2""><hr size=""1"" width=""100%""></td>"
'f.WriteLine "</TR>"
'f.WriteLine "<tr><td colspan=""2"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
		f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
		f.WriteLine "<td valign=""top"" colspan=""2"">&nbsp;</td>"
		f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""2""><br></td><tr>"
'End chnage by atul remove the hr  make line with tr and table
		
' HUMANITARIAN ORG ?
		f.WriteLine "<tr>"
' Atul make colspan=2 from 1
		f.WriteLine "<td colspan=""2"">" & gITEXTRC("i_19") & "</TD>"
		f.WriteLine "</TR>"
		f.WriteLine "<tr>"
' Atul make colspan=2 from 1
		f.WriteLine "<td align=""left"" colspan=""2""><span class=""alert"">" & gITEXTRC("i_12") & " </span><strong>"
		If JAPINFORC("cand_employ5_pres_i") = "1" then
			f.WriteLine gITEXTRC("i_6")
		else
			f.WriteLine gITEXTRC("i_7")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<tr>"
' Atul make colspan=2 from 1
		f.WriteLine "<td align=""left"" colspan=""2""><span class=""alert"">" & gITEXTRC("i_13") & "</span><strong>"
		If JAPINFORC("cand_employ5_prev_i") = "0" then
		f.WriteLine gITEXTRC("i_7")
		else
		f.WriteLine gITEXTRC("i_6")
		end if
		f.WriteLine "</strong></td></TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td colspan=""2"">" & gITEXTRC("i_20") & " <strong>" & JAPINFORC("cand_io_m") & "</strong>"
		f.WriteLine "</TD></tr>"
		f.WriteLine "</table>"
		f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "</table>"
	end if
end if
 
' ******************************************************************************
' SECTION INT RC ExPERIENCE - IFRC ONLY
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION RC
		if pv_new_sessioncode = 7000 then
' BEGIN CHECK IF PERSON SELECTED SECTION RC
		if instr(pv_parts,"RC,") then

			JAPGEOLIST90sql = "SELECT tx_rsys_candgeo.candgeo_id_c, tx_rsys_candgeo.candgeo_cand_c, 	g.geoloc_"& new_lng_code &"_" & pv_new_sessioncode & " as geolocdsc 	FROM tx_rsys_candgeo INNER JOIN 	core_geolocf g ON 	tx_rsys_candgeo.candgeo_id_c = g.geoloc_id_c 		WHERE tx_rsys_candgeo.candgeo_cand_c = " & applicant_id & " AND tx_rsys_candgeo.candgeo_type_i  = 90 "
''d set JAPGEOLIST90 =rsys_db_select.execute(JAPGEOLIST90sql)
			Set JAPGEOLIST90 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPGEOLIST90.Open JAPGEOLIST90sql, rsys_db_select, 1, 1

			JAPGEOLIST61sql = "SELECT gl.candgeo_cty_c, gl.candgeo_cand_c, 	g.c_"& new_lng_code &"_name as ctyname FROM tx_rsys_candgeo gl INNER JOIN v_country_list_" & pv_new_sessioncode & " g ON 	gl.candgeo_cty_c = g.cty_uid_c WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 61 "
''d set JAPGEOLIST61 =rsys_db_select.execute(JAPGEOLIST61sql)
			Set JAPGEOLIST61 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPGEOLIST61.Open JAPGEOLIST61sql, rsys_db_select, 1, 1

			JAPGEOLIST62sql = "SELECT gl.candgeo_cty_c, gl.candgeo_cand_c, 	g.c_"& new_lng_code &"_name as ctyname FROM tx_rsys_candgeo gl INNER JOIN v_country_list_" & pv_new_sessioncode & " g ON 	gl.candgeo_cty_c = g.cty_uid_c WHERE gl.candgeo_cand_c = " & applicant_id & " AND gl.candgeo_type_i  = 62 "
''d set JAPGEOLIST62 =rsys_db_select.execute(JAPGEOLIST62sql)
			Set JAPGEOLIST62 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPGEOLIST62.Open JAPGEOLIST62sql, rsys_db_select, 1, 1

			JAPMOBsql = "SELECT     cm.candmisc_mob_relocate_i, d.drop_dsc_" & new_lng_code & "_t AS reloc FROM  dbo.tx_rsys_candmisc cm INNER JOIN dbo.tr_rsys_drop d ON cm.candmisc_mob_relocate_i = d.drop_id_c WHERE     (cm.cand_id_c = " & applicant_id & ")"
''d set JAPMOB =rsys_db_select.execute(JAPMOBsql)
			Set JAPMOB = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPMOB.Open JAPMOBsql, rsys_db_select, 1, 1

			JAPINFOIEsql = " SELECT     c.upd_d, c.cand_lnam_t, c.cand_fnam_t, c.cand_thisorg_short_i_"& pv_new_sessioncode &" AS thisorg_short, c.cand_thisorg_jobtype_i_"& pv_new_sessioncode &" AS thisorg_type, c.cand_geo_ctywork_m, c.cand_io_i, c.cand_io_t, c.cand_io_ds_t, c.cand_io_year_n, c.cand_io_grade_t, c.cand_io_m, c.cand_io_f_i, c.cand_io_f1_t, c.cand_io_f1_ds_t, c.cand_io_f1_year_n, c.cand_io_f1_year_end_n, c.cand_io_f1_grade_t, c.cand_io_f2_t, c.cand_io_f2_ds_t, c.cand_io_f2_year_n, c.cand_io_f2_year_end_n, c.cand_io_f2_grade_t, c.cand_io_f3_t, c.cand_io_f3_ds_t, c.cand_io_f3_year_n, c.cand_io_f3_year_end_n, c.cand_io_f3_grade_t, c.cand_geo_exp_i, c.cand_geo_res_i, c.cand_geo_res_m, d.drop_dsc_"& new_lng_code &"_t AS georesdsc FROM         dbo.td_rsys_cand c LEFT OUTER JOIN dbo.tr_rsys_drop d ON c.cand_geo_res_i = d.drop_id_c WHERE     (c.cand_id_c = "& applicant_id &")"
''d set JAPINFOIE =rsys_db_select.execute(JAPINFOIEsql)
			Set JAPINFOIE = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPINFOIE.Open JAPINFOIEsql, rsys_db_select, 1, 1
f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor& """>" & gITEXTF("i_3") & "</font></h2>"
		
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'f.WriteLine "<TR>"
' Atul make space in between th and a
'f.WriteLine "<th colspan=""4"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTF("i_3") & "_|" & intNumber  &  "'>" & gITEXTF("i_3") & "</a></th>"
'f.WriteLine "</TR>"
		
'Commented by atul
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60")& """><h2><font color=""" & g_headerColor& """>" & gITEXTF("i_3") & "</font></h2></td>"
'f.WriteLine "</TR>"
		
		
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""4"" class=""blacktext"">&nbsp;</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""4"">"
		f.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""0"" border=""0"" width=""" & widther & """>"
' ANOTHER GEO LIST AND ANOTHER CTY LIST
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">" & gITEXTF("i_8") & "</TD>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""5"">"
'f.WriteLine "<font color=""blue""><b>"
			if JAPGEOLIST90.eof = true then
				f.WriteLine gITEXTF("i_21")
			Else
			    counter = 0
			    if instr(pv_parts,"RCCountryList,") then
					Do while JAPGEOLIST90.eof = false
						counter = counter + 1
						f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""5"">"
						f.WriteLine "<font color=""blue""><b>"
						f.WriteLine JAPGEOLIST90("geolocdsc")
						JAPGEOLIST90.movenext
						f.WriteLine "</b></font></td>"
						if counter mod 2 = 0 then
								f.WriteLine "</tr>"
								f.WriteLine "<tr>"
						end if
					loop
				End If
			End If
'f.WriteLine "</b></font></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">" & gITEXTF("i_9") & "</TD>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""2"">"
'f.WriteLine "<font color=""blue""><b>"
			If JAPGEOLIST61.eof = true then
				f.WriteLine gITEXTF("i_21")
			Else
			    counter = 0
			    if instr(pv_parts,"RCCountryList,") then
					Do while JAPGEOLIST61.eof = false
						counter = counter + 1
						f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""2"">"
						f.WriteLine "<font color=""blue""><b>"
						f.WriteLine JAPGEOLIST61("ctyname")
						JAPGEOLIST61.movenext
						f.WriteLine "</b></font></td>"
						if counter mod 2 = 0 then
								f.WriteLine "</tr>"
								f.WriteLine "<tr>"
						end if
					loop
				End If
			End If
'f.WriteLine "</b></font></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"" class=""blacktext"">&nbsp;</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">"
		f.WriteLine gITEXTF("i_13") & " <strong>"
			If trim(JAPINFOIE("cand_geo_exp_i")) = "1" then
				f.WriteLine gITEXTF("i_6")
			else
				f.WriteLine gITEXTF("i_7")
			end if
			f.WriteLine "JAPINFOIE||" & JAPINFOFOIE("cand_geo_exp_i") & "</strong>"
		f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">" & gITEXTF("i_8") & "</TD>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""5"">"
'f.WriteLine "<font color=""blue""><b>"
			If JAPGEOLIST.eof = true then
				f.WriteLine gITEXTF("i_21")
			Else
				counter = 0
				if instr(pv_parts,"RCCountryList,") then
					Do while JAPGEOLIST.eof = false
						counter = counter + 1
						f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""5"">"
						f.WriteLine "<font color=""blue""><b>"
						f.WriteLine JAPGEOLIST("geolocdsc")
						f.WriteLine "</b></font></td>"
						JAPGEOLIST.movenext

						if counter mod 2 = 0 then
								f.WriteLine "</tr>"
								f.WriteLine "<tr>"
						end if
					loop
				End If
			End If
'f.WriteLine "</b></font></td>"
		f.WriteLine "</TR>"

' ADD ANOTHER CTY LIST FOR IFRC 7000
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">" & gITEXTF("i_9") & "</TD>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""5"">"
'f.WriteLine "<font color=""blue""><b>"
			If JAPGEOLIST62.eof = true then
				f.WriteLine gITEXTF("i_21")
			Else
				counter = 0
				if instr(pv_parts,"RCCountryList,") then
					Do while JAPGEOLIST62.eof = false
						counter = counter + 1
						f.WriteLine "<td valign=""top"" align=""left"" valign=""middle"" colspan=""5"">"
						f.WriteLine "<font color=""blue""><b>"
						f.WriteLine JAPGEOLIST62("ctyname")
						f.WriteLine "</b> </font>"
						f.WriteLine "</td>"
						JAPGEOLIST62.movenext
						if counter mod 2 = 0 then
								f.WriteLine "</tr>"
								f.WriteLine "<tr>"
						end if
					loop
				End If
			End If
'f.WriteLine "</b> </font>"
'f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">&nbsp;</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">" & gITEXTF("i_88") & "</TD>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">" & gITEXTF("i_17") & " <strong>" & JAPINFOIE("georesdsc")& "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">&nbsp;</td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">" & gITEXTF("i_18") & " <strong>" & JAPINFOIE("cand_geo_res_m")& "</strong></TD>"
		f.WriteLine "</TR>"
' END ADDITIONAL COUNTRY LIST FOR IFRC 7000

'IFRC 7000 NEW EXEMPTION
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""5"">&nbsp;</td>"
		f.WriteLine "</TR>"
		if JAPMOB.eof = false then
			f.WriteLine "<TR><td valign=""top"" colspan=""5"">" & gITEXTF("i_87") & " <strong>" & JAPMOB("reloc")& "</strong></TD></tr>"
		end if
'    '------------ Additional information on geographical info -------------------------------------------------->
		f.WriteLine "</table>"
		f.WriteLine "</td>"
		f.WriteLine "</TR>"
		f.WriteLine "</table>"
		end if
		end if



' ******************************************************************************
' SECTION D - COMPUTER SKILLS
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION D
		if instr(pv_parts,"D,") then
Dim viewupd

'BEGIN CHECK ON WHETHER USING NEW COMP SKILLS OR NOT
'03 MAR 15 LJL added WTO, WMO, UNWomen to computer skills output
'07 SEP 15 LJL added UPU to the new comp skills output
		if pv_new_sessioncode = 2400 OR pv_new_sessioncode = 2600 OR pv_new_sessioncode = 2800 OR pv_new_sessioncode = 2900 OR pv_new_sessioncode = 3000 OR pv_new_sessioncode = 5500 then
		
		
		JAPCOMPSsql = " SELECT tx_rsys_candcompitem. cand_id_c, tx_rsys_candcompitem.candcompitem_id_c, tx_rsys_candcompitem.compitem_level_c, tr_rsys_compitem.compitem_id_c, tr_rsys_compitem.compitem_dsc_"& new_lng_code &"_t AS q_elem, tr_rsys_compitem.compitem_rem_"& new_lng_code &"_m AS q_elemrem FROM tr_rsys_compitem, tx_rsys_candcompitem WHERE tr_rsys_compitem.compitem_thisorg_" & pv_new_sessioncode & "  = 1 AND tx_rsys_candcompitem.cand_id_c = " & applicant_id & " AND (tx_rsys_candcompitem.compitem_id_c = tr_rsys_compitem.compitem_id_c) ORDER BY tr_rsys_compitem.compitem_order_c, tr_rsys_compitem.compitem_dsc_"& new_lng_code &"_t "
''d set JAPINSTLIST12 =rsys_db_select.execute(JAPINSTLIST12sql)
		Set JAPCOMPS = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPCOMPS.Open JAPCOMPSsql, rsys_db_select, 1, 1
'response.write JAPCOMPSsql
'response.end
'chnage  by atul make table tag out side if condition becuase  we  need

if JAPCOMPS.eof = false then
f.WriteLine "<p><br></p>"
f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTD("i_1") & "</font></h2>"
end if 
f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
if JAPCOMPS.eof = false then
'f.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'f.WriteLine "<TR>"
' Atul make space in between th and  a
'f.WriteLine "<th colspan=""4"">  &lt;a style=""abcpdf-tag-visible: true;"" id='" & gITEXTD("i_1") & "_|" & intNumber  &  "'&gt;"  & gITEXTD("i_1") &  "&lt;/a&gt;</th>"
'f.WriteLine "</TR>"
	
'Commented by atul fro table of content
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTD("i_1") & "</font></h2></td>"
'f.WriteLine "</TR>"
	
'11 DEC 10 LJL added ICDL for ITU
	if pv_new_sessioncode = 2400 then
' INTERNATIONAL COMPUTER DRIVERS LICENSE - ITU ONLY	if pv_new_sessioncode = 2400 then
'f.WriteLine "<tr><Td colspan='2'><strong>" & gITEXTD("i_40") & "</strong></Td></tr>"
' Atul make  colspan=4 from 1
		f.WriteLine "<tr><Td colspan='4'>"& gITEXTD("i_41") & "&nbsp; <strong>"   				
		If JAPINFO1("cand_comp_icdl_i") = "0" then
			f.WriteLine gITEXTD("i_7")
		elseIf JAPINFO1("cand_comp_icdl_i") = "1" then
			f.WriteLine gITEXTD("i_6")
		else
			f.WriteLine gITEXTD("i_44")
		end if
		f.WriteLine "</strong></td></tr><TR><td valign='top' colspan='4'>" & gITEXTD("i_42") & " &nbsp; &nbsp; <strong>" & JAPINFO1("cand_comp_icdl_dt") & "</strong></TD></TR>"
		
'start chnage by atul remove the hr  make line with tr and table
'f.WriteLine "<tr><td colspan='4'><hr size='1'></td></tr>"
'f.WriteLine "<tr><td colspan=""4"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""4""><br></td><tr>"
		f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
		f.WriteLine "<td valign=""top"" colspan=""4"">&nbsp;</td>"
		f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""4""><br></td><tr>"
'End chnage by atul remove the hr  make line with tr and table
		
	end if
		
Dim pv_rower_comp
	pv_rower_comp = 0
	Do while JAPCOMPS.eof = false
		pv_rower_comp = pv_rower_comp + 1
		if   pv_rower_comp = 1  then
			f.WriteLine "<tr>"
		end if
	
		f.WriteLine "<td><table border=""0"" ><tr><td>" & JAPCOMPS("q_elemrem") & "<br><font color='blue'>&nbsp; "
		if JAPCOMPS("compitem_level_c") = "0" then
			f.WriteLine "-"
		elseif JAPCOMPS("compitem_level_c") = "1" then
			f.WriteLine gITEXTD("i_16")
		elseif JAPCOMPS("compitem_level_c") = "2" then
			f.WriteLine gITEXTD("i_17")
		elseif JAPCOMPS("compitem_level_c") = "3" then
			f.WriteLine gITEXTD("i_18")
		elseif JAPCOMPS("compitem_level_c") = "4" then
			f.WriteLine gITEXTD("i_32")
		end if
'change by atul  font was not closed
'f.WriteLine "</td></tr></table></td>"
		f.WriteLine "</font></td></tr></table></td>"
		if   pv_rower_comp = 4  then
			f.WriteLine "</tr>"
			pv_rower_comp = 0
		end if
 	JAPCOMPS.movenext
 	loop
	if   pv_rower_comp < 4 then
' Start chnage  by Atul  becuase  ti colspan html was coming wrong if there is less then 4 records
		if(pv_rower_comp=3) then 
			f.WriteLine "<td>&nbsp;</td>"
		elseif(pv_rower_comp=2) then
			f.WriteLine "<td colspan=""2"">&nbsp;</td>"			
		elseif(pv_rower_comp=1) then 	
			f.WriteLine "<td colspan=""3"">&nbsp;</td>"
		end if  
' End 
		f.WriteLine "</tr>"
	end if
  
end if
	              				
	if   len(JAPINFO1("cand_pc_skills_t")) then
		f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTD("i_22") & "<br><strong>" & JAPINFO1("cand_pc_skills_t") & "</strong></TD></tr>"
	end if
	if   len(JAPINFO1("cand_pc_other_t")) then
		f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTD("i_21") & "<br><strong>" & JAPINFO1("cand_pc_other_t")& "</strong>	</TD></tr>"
	end if

'    '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
	if   viewupd = "1" and viewupd = "YES"  then
		if   GETUPDS("editD_d") > GETAPPDATE("candjob_d")  then
			f.WriteLine "<TR>"
			f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") &""">Last update date <strong>" & formatdatetime(GETUPDS("EditD_d"),1) & " AFTER APPLYING</strong></td>"
			f.WriteLine "</tr>"
		end if
	end if
f.WriteLine "</TABLE>"
'Atul make commnet to  belwo tags
'f.WriteLine "</td>"
'f.WriteLine "</tr>"
'f.WriteLine "</table>"
'f.WriteLine "</TABLE>"
	

'ELSE CHECK ON WHETHER USING NEW COMP SKILLS OR NOT
		else
		f.WriteLine "<p><br></p>"
'Add by atul fro table of Content
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTD("i_1") & "</font></h2>"
		
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'f.WriteLine "<TR>"
' Atul make space in between th and  a
'f.WriteLine "<th colspan=""4"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTD("i_1") & "_|" & intNumber  &  "'>" & gITEXTD("i_1") & "</a></th>"
'f.WriteLine "</TR>"
'commented by atul fro table of Content
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTD("i_1") & "</font></h2></td>"
'f.WriteLine "</TR>"
		
		f.WriteLine "<TR>"
' Atul  make colspan=4 from 1
		f.WriteLine "<td valign=""top"" width=""100%"" colspan=""4"">"
		f.WriteLine "<TABLE border=""0"" bordercolor=""black"" width=""100%"">"
'    '------------- PROBLEM IN THIS PART  - IF PERSON PUTS IN VERY LONG STRING ------------------>
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" width=""25%"">" & gITEXTD("i_5") & "</TD>"
		f.WriteLine "<td valign=""top"" width=""50%""><strong>" & JAPINFO1("cand_wp_other_t") & "&nbsp;</strong></td>"
		f.WriteLine "<td valign=""top"" width=""25%""><strong>"
		if   JAPINFO1("cand_wp_i") = 0  then
			f.WriteLine "----"
		elseif   JAPINFO1("cand_wp_i") = 1  then
			f.WriteLine gITEXTD("i_16")
		elseif   JAPINFO1("cand_wp_i") = 2  then
			f.WriteLine gITEXTD("i_17")
		elseif   JAPINFO1("cand_wp_i") = 3  then
			f.WriteLine gITEXTD("i_18")
		elseif   JAPINFO1("cand_wp_i") = 4  then
			f.WriteLine gITEXTD("i_32")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" width=""25%"">" & gITEXTD("i_10") & "</TD>"
		f.WriteLine "<td valign=""top"" width=""50%""><strong>" & JAPINFO1("cand_db_other_t") & "&nbsp;</strong></td>"
		f.WriteLine "<td valign=""top"" width=""25%""><strong>"
		if   JAPINFO1("cand_db_i") = 0  then
			f.WriteLine "----"
		elseif   JAPINFO1("cand_db_i") = 1  then
			f.WriteLine gITEXTD("i_16")
		elseif   JAPINFO1("cand_db_i") = 2  then
			f.WriteLine gITEXTD("i_17")
		elseif   JAPINFO1("cand_db_i") = 3  then
			f.WriteLine gITEXTD("i_18")
		elseif   JAPINFO1("cand_db_i") = 4  then
			f.WriteLine gITEXTD("i_32")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR><td valign=""top"" width=""25%"">" & gITEXTD("i_9") & "</TD>"
			f.WriteLine "<td valign=""top"" width=""50%""><strong>" & JAPINFO1("cand_sps_other_t") & "&nbsp;</strong></TD>"
		f.WriteLine "<td valign=""top"" width=""25%""><strong>"
		if   JAPINFO1("cand_sps_i") = 0  then
			f.WriteLine "----"
		elseif   JAPINFO1("cand_sps_i") = 1  then
			f.WriteLine gITEXTD("i_16")
		elseif   JAPINFO1("cand_sps_i") = 2  then
			f.WriteLine gITEXTD("i_17")
		elseif   JAPINFO1("cand_sps_i") = 3  then
			f.WriteLine gITEXTD("i_18")
		elseif   JAPINFO1("cand_sps_i") = 4  then
			f.WriteLine gITEXTD("i_32")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" width=""25%"">" & gITEXTD("i_29") & "</TD>"
		f.WriteLine "<td valign=""top"" width=""50%""><strong>" & JAPINFO1("cand_comp_os_other_t") & "&nbsp;</strong></td>"
		f.WriteLine "<td valign=""top"" width=""25%""><strong>"
		if   JAPINFO1("cand_comp_os_i") = 0  then
			f.WriteLine "----"
		elseif   JAPINFO1("cand_comp_os_i") = 1  then
			f.WriteLine gITEXTD("i_16")
		elseif   JAPINFO1("cand_comp_os_i") = 2  then
			f.WriteLine gITEXTD("i_17")
		elseif   JAPINFO1("cand_comp_os_i") = 3  then
			f.WriteLine gITEXTD("i_18")
		elseif   JAPINFO1("cand_comp_os_i") = 4  then
			f.WriteLine gITEXTD("i_32")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" width=""25%"">" & gITEXTD("i_19") & "</TD>"
		f.WriteLine "<td valign=""top"" width=""50%""><strong>" & JAPINFO1("cand_web_other_t") & "&nbsp;</strong></TD>"
		f.WriteLine "<td valign=""top"" width=""25%""><strong>"
		if   JAPINFO1("cand_web_i") = 0  then
			f.WriteLine "----"
		elseif   JAPINFO1("cand_web_i") = 1  then
			f.WriteLine gITEXTD("i_16")
		elseif   JAPINFO1("cand_web_i") = 2  then
			f.WriteLine gITEXTD("i_17")
		elseif   JAPINFO1("cand_web_i") = 3  then
			f.WriteLine gITEXTD("i_18")
		elseif   JAPINFO1("cand_web_i") = 4  then
			f.WriteLine gITEXTD("i_32")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" width=""25%"">" & gITEXTD("i_11") & "</TD>"
		f.WriteLine "<td valign=""top"" width=""50%""><strong>" & JAPINFO1("cand_pres_other_t") & "&nbsp;</strong></TD>"
		f.WriteLine "<td valign=""top"" width=""25%""><strong>"
		if   JAPINFO1("cand_pres_i") = 0  then
			f.WriteLine "----"
		elseif   JAPINFO1("cand_pres_i") = 1  then
			f.WriteLine gITEXTD("i_16")
		elseif   JAPINFO1("cand_pres_i") = 2  then
			f.WriteLine gITEXTD("i_17")
		elseif   JAPINFO1("cand_pres_i") = 3  then
			f.WriteLine gITEXTD("i_18")
		elseif   JAPINFO1("cand_pres_i") = 4  then
			f.WriteLine gITEXTD("i_32")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" width=""25%"">" & gITEXTD("i_20") & "</TD>"
		f.WriteLine "<td valign=""top"" width=""50%""><strong>" & JAPINFO1("cand_prgming_other_t") & "&nbsp;</strong></TD>"
		f.WriteLine "<td valign=""top"" width=""25%""><strong>"
		if   JAPINFO1("cand_prgming_i") = 0  then
			f.WriteLine "----"
		elseif   JAPINFO1("cand_prgming_i") = 1  then
			f.WriteLine gITEXTD("i_16")
		elseif   JAPINFO1("cand_prgming_i") = 2  then
			f.WriteLine gITEXTD("i_17")
		elseif   JAPINFO1("cand_prgming_i") = 3  then
			f.WriteLine gITEXTD("i_18")
		elseif   JAPINFO1("cand_prgming_i") = 4  then
			f.WriteLine gITEXTD("i_32")
		end if
		f.WriteLine "</strong></td>"
		f.WriteLine "</TR>"
		f.WriteLine "</table>"
		f.WriteLine "</td>"
		f.WriteLine "</TR>"
		if   len(JAPINFO1("cand_pc_skills_t")) then
			f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTD("i_22") & "<br><strong> " & JAPINFO1("cand_pc_skills_t") & "</strong></TD></tr>"
		end if
		if   len(JAPINFO1("cand_pc_other_t")) then
			f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTD("i_21") & "<br><strong>" & JAPINFO1("cand_pc_other_t")& "</strong>	</TD></tr>"
		end if

'    '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
			if   viewupd = "1" and viewupd = "YES"  then
				if   GETUPDS("editD_d") > GETAPPDATE("candjob_d")  then
					f.WriteLine "<TR>"
					f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") &""">Last update date <strong>" & formatdatetime(GETUPDS("EditD_d"),1) & " AFTER APPLYING</strong></td>"
					f.WriteLine "</tr>"
				end if
			end if
		f.WriteLine "</TABLE>"
'f.WriteLine "</td>"
'f.WriteLine "</tr>"
'f.WriteLine "</table>"
		
' END CHECK ON WHETHER USING NEW COMP SKILLS OR NOT
		end if
		
' END CHECK IF PERSON SELECTED SECTION D
		end if




' ******************************************************************************
' SECTION G - DEPENDANTS / RELATIVES / REFERENCES
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION G
		if instr(pv_parts,"G,") then

'MODS --
'23 FEB 05 LJL org check
'09 MAR 05 LJL removed town arrival date for WTO
'--------------------------------------------------------------->
''---------------------- ADD ORG INFO 1000 check for 3000 23FEB 05 LJL --------------------------------------->

' NEEDNEEDNEED MOVE TO SP ASAP
'  '---------------------- ADD ORG INFO 1000/ 7000 check for 3000 23FEB 05 LJL --------------------------------------->
'16 FEB 16 LJL not sure why ILO 2000 was not to show dependants
		'if   pv_new_sessioncode <> 2000 then
			JAPDEPsql = "SELECT canddepend_nam_t, canddepend_bth_d, canddepend_rel_t, canddepend_id_c, canddepend_id_c 	from tx_rsys_canddepend WHERE canddepend_cand_c = " & applicant_id & " 	"
''d set JAPDEP =rsys_db_select.execute(JAPDEPsql)
			Set JAPDEP = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPDEP.Open JAPDEPsql, rsys_db_select, 1, 1
'16 FEB 16 LJL not sure why ILO 2000 was not to show dependants
		'end if

		if   pv_new_sessioncode <> 7000  then
			JAPRELsql = "SELECT candrelative_nam_t, candrelative_rel_t, candrelative_org_t, candrelative_cand_c, candrelative_id_c from tx_rsys_candrelative 		WHERE candrelative_cand_c = " & applicant_id & " 	"
''d set JAPREL =rsys_db_select.execute(JAPRELsql)
			Set JAPREL = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
			rsys_db_select.CommandTimeout = 320
			JAPREL.Open JAPRELsql, rsys_db_select, 1, 1
		end if

		JAPREFsql = "SELECT candrefer_fax_c, candrefer_known_t, candrefer_email_t, candrefer_nam_t, candrefer_adr_t, candrefer_phn_t, candrefer_occ_t, candrefer_cand_c, candrefer_id_c from tx_rsys_candrefer 		WHERE candrefer_cand_c = " & applicant_id & " 	"
''d set JAPREF =rsys_db_select.execute(JAPREFsql)
		Set JAPREF = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPREF.Open JAPREFsql, rsys_db_select, 1, 1

'Added by Atul fro table  of Content
		'Changed by Gevorg
		f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTG("i_1") & "</font></h2>"
		'f.WriteLine "<h2><font color=""Teal"">" & gITEXTG("i_1") & "</font></h2>"
		
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'f.WriteLine "<TR>"
' Atul make space i between th and  a tag
'f.WriteLine "<th colspan=""4"">  &lt;a style=""abcpdf-tag-visible: true;"" id='" & gITEXTG("i_1") & "_|" & intNumber  &  "'&gt;"  & gITEXTG("i_1") &  "&lt;/a&gt;</th>"
'f.WriteLine "</TR>"
'Commented by atul fro table of content
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTG("i_1") & "</font></h2></td>"
'f.WriteLine "</TR>"
		
'  '---------------------- ADD ORG INFO 1000/ 7000 check for 3000 23FEB 05 LJL --------------------------------------->
' ******************************************************************************
' SECTION G - DEPENDANTS
' ******************************************************************************
'16 FEB 16 LJL not sure why ILO 2000 was not to show dependants
		'if pv_new_sessioncode <> 2000 then
		f.WriteLine "<TR>"
		f.WriteLine "<td valign=""top"" colspan=""4""><font color=""Teal""><i><strong>" & gITEXTG("i_3") & "</strong></i></font></td>"
		f.WriteLine "</TR>"
		if  JAPDEP.eof = true then
			f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTG("i_28") & ".</TD></tr>"
		else
			f.WriteLine "<TR><td align=""left"" colspan=""2""><I>" & gITEXTG("i_5") & "</I></td>"
			f.WriteLine "<td align=""left""><I>" & gITEXTG("i_54") & "</I></td>"
			f.WriteLine "<td align=""left""><I>" & gITEXTG("i_7") & "</I></td>"
'--------- ADDED TO TEST PDF PROB ------------->
			f.WriteLine "</TR>"
		end if
			Do while JAPDEP.eof = false
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"" colspan=""2""><strong>" & JAPDEP("canddepend_nam_t") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>"
				if   isdate(JAPDEP("canddepend_bth_d")) then
					f.WriteLine day(JAPDEP("canddepend_bth_d")) & "-" & monthname(month(JAPDEP("canddepend_bth_d")),1) & "-" & year(JAPDEP("canddepend_bth_d"))
				end if
				f.WriteLine "</strong></td>"
				f.WriteLine "<td valign=""top""><strong>" & JAPDEP("canddepend_rel_t") & "</strong></TD>"
				f.WriteLine "</TR>"
				
'start chnage by atul remove the hr  make line with tr and table
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4""><hr noshade size=""1""></td>"
'f.WriteLine "</TR>"
'f.WriteLine "<tr><td valign=""top"" colspan=""4"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""4""><br></td><tr>"
				f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
				f.WriteLine "<td valign=""top"" colspan=""4"">&nbsp;</td>"
				f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""4""><br></td><tr>"
'End chnage by atul remove the hr  make line with tr and table
				
				JAPDEP.movenext
			loop
'16 FEB 16 LJL not sure why ILO 2000 was not to show dependants
		'end if

' ******************************************************************************
' SECTION G - RELATIVES
' ******************************************************************************
' NO RELATIVES FOR IFRC 7000
		if   pv_new_sessioncode <> 7000  then
			f.WriteLine "<TR><td valign=""top"" colspan=""4""><font color=""Teal""><i><strong>" & gITEXTG("i_55") & "</strong></i></font></TD></tr>"
			if   JAPREL.eof = true  then
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top"" colspan=""4"">" & gITEXTG("i_29") & ".</td>"
				f.WriteLine "</TR>"
			else
				f.WriteLine "<TR>"
				f.WriteLine "<td align=""left""><I>" & gITEXTG("i_5") & "</I></td>"
				f.WriteLine "<td align=""left"" colspan=""2""><I>" & gITEXTG("i_7") & "</I></td>"
				f.WriteLine "<td align=""left""><I>" & gITEXTG("i_12") & "</I></td>"
				f.WriteLine "</TR>"
			end if
			Do while JAPREL.eof = false
				f.WriteLine "<TR>"
				f.WriteLine "<td valign=""top""><strong>" & JAPREL("candrelative_nam_t") & "</strong></TD>"
				f.WriteLine "<td valign=""top"" colspan=""2""><strong>" & JAPREL("candrelative_rel_t") & "</strong></TD>"
				f.WriteLine "<td valign=""top""><strong>" & JAPREL("candrelative_org_t") & "</strong></TD>"
'    			'--------- ADDED TO TEST PDF PROB ------------->
				f.WriteLine "</TR>"
				
			
				
'start chnage by atul remove the hr  make line with tr and table
'f.WriteLine "<TR>"
'f.WriteLine "<td valign=""top"" colspan=""4""><hr noshade size=""1""></td>"
'f.WriteLine "</TR>"
'f.WriteLine "<tr><td valign=""top"" colspan=""4"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""4""><br></td><tr>"
				f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
				f.WriteLine "<td valign=""top"" colspan=""4"">&nbsp;</td>"
				f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""4""><br></td><tr>"
'End chnage by atul remove the hr  make line with tr and table
				
				JAPREL.movenext
			loop
' END IFRC 7000 NO RELATIVES
		end if

' ******************************************************************************
' SECTION G - REFERENCES
' ******************************************************************************
' WANTS AS A MAIN HEADER 7000 IFRC
		'04/06/2009 DD-There is separete for WTO references
		if pv_new_sessioncode <> 3000 then
			
			'f.WriteLine "TR><td valign=""top"" colspan=""4"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTG("i_15") & "///</font></h2></TD></tr>"

			
		    'if pv_new_sessioncode = 7000 then
				' Atul make space in between th and a tag 
				'f.WriteLine "<tr><th colspan=""4"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTG("i_15") & "_|" & intNumber  &  "'>" & gITEXTG("i_15") & "</a></th></tr>"
'01 JUN 15 LJL added references as separate section header bar, WMO request		
				f.WriteLine "</Table>"	
				f.WriteLine "<p><br></p>"
'02 JUN 15 GG changed References section look like other sections				
				f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTG("i_15") & "</font></h2>"
				f.WriteLine "<Table border=""1"" >"
		   ' else
				' Atul  kake colspan=4 from 1
			 '   f.WriteLine "<TR><td valign=""top"" colspan=""4""><font color=""Teal""><i><strong>" & gITEXTG("i_15") & "</strong></I></font></TD></tr>"
		    'end if

			' WANTS AS A MAIN HEADER 7000 IFRC
		    if   JAPREF.eof = true  then
			    f.WriteLine "<TR><td valign=""top"" colspan=""4"">" & gITEXTG("i_30") & ".</TD></tr>"
		    else
			    f.WriteLine "<TR><td align=""left"" valign=""top"" width=""15%""><em>" & gITEXTG("i_5") & "</em></td>"
			    f.WriteLine "<td align=""left"" valign=""top"" width=""35%""><em>" & gITEXTG("i_17") & "<br>" & gITEXTG("i_18")
			    f.WriteLine "<br>" & gITEXTG("i_80") & "</em></td>"
			    f.WriteLine "<td align=""left"" valign=""top"" width=""20%""><em>" & gITEXTG("i_75") & "<br>" & gITEXTG("i_19") & "</em></td>"
			    if pv_new_sessioncode <> 3000 then
				    f.WriteLine "<td align=""left"" valign=""top"" width=""30%""><em>" & gITEXTG("i_68") & "</em></td>"
			    end if
			    f.WriteLine "</TR>"
		    end if
		    Do while JAPREF.eof = false
			    f.WriteLine "<TR><td valign=""top"" width=""15%""><b>" & JAPREF("candrefer_nam_t") & "</b></TD>"
			    f.WriteLine "<td valign=""top"" width=""35%""><b>" & JAPREF("candrefer_adr_t")& "<Br>" & JAPREF("candrefer_phn_t")
			    if len(JAPREF("candrefer_fax_c")) then
				    f.WriteLine "<br>" & JAPREF("candrefer_fax_c")
			    end if
			    f.WriteLine "</b></td>"
			    f.WriteLine "<td valign=""top"" width=""20%""><b>" & JAPREF("candrefer_email_t") & "<Br>" & JAPREF("candrefer_occ_t") & "</b></TD>"
			    if pv_new_sessioncode <> 3000 then
				    f.WriteLine "<td valign=""top"" width=""30%""><b>" & JAPREF("candrefer_known_t") & "</b></TD>"
			    end if
				''--------- PROBLEM HERE 1-50 on FT145 for certain apps ---------------->
			    f.WriteLine "</TR>"
				JAPREF.movenext
		    loop
'01 JUN 15 LJL made table 100% to fill page across for In city since, in country since
			f.WriteLine "</table><table width=100% border=1 bordercolor=green>"
			'  '------------ TOOK OUT PDF TESTER 6	PROBLEM FT145 12,16,21,24 ERROR HERE  --------------------->
			'  '--------------- TOOK OUT PDF TESTER 6a	PROBLEM FT145 12,16,21,24 THIS PART OKAY ------------->
			'Do while JAPINFO.eof = false
			' NO IFRC 7000
		    if pv_new_sessioncode <> 7000 then
				' NOT SHOW THIS PART IF WTO 3000
				if pv_new_sessioncode <> 3000 then
					'Session.LCID = 127
					f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' ><td valign=""top"" colspan=""4"">&nbsp;</td></tr>"
					f.WriteLine "<TR><td valign=""top"" width=""25%"" colspan=""1"">" & gITEXTG("i_21") & "</TD>"
					f.WriteLine "<td width=""75%"" valign=""top"" valign=""middle"" colspan=""3"">"
					f.WriteLine "<I>" & gITEXTG("i_22") & " " & gITEXTG("i_23") & "</i>: <strong>"
					if isdate(JAPINFO1("cand_cty_arr_d")) then
						f.WriteLine day(JAPINFO1("cand_cty_arr_d")) & "-" & monthname(month(JAPINFO1("cand_cty_arr_d")),1) & "-" & year(JAPINFO1("cand_cty_arr_d"))
					end if
					f.WriteLine "</strong> <i>" & gITEXTG("i_24") & " " & gITEXTG("i_23") & "</i>: <strong>"
					if   isdate(JAPINFO1("cand_twn_arr_d")) then
						f.WriteLine day(JAPINFO1("cand_twn_arr_d")) & "-" & monthname(month(JAPINFO1("cand_twn_arr_d")),1) & "-" & year(JAPINFO1("cand_twn_arr_d"))
					end if
					f.WriteLine "</strong></TD></tr>"
				end if 
		    end if' NO IFRC 7000 
		 end if	' END CHECK IF org <> 3000
		 f.WriteLine "</table>"
end if' END CHECK IF PERSON SELECTED SECTION G

' ******************************************************************************
' SECTION J - ADDITIONAL INFORMATION
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION J
		f.WriteLine "<!--Test33-->"
		if instr(pv_parts,"J,") then
'01 Sept 10 DD Added condition to avoid header (Additional Information) in geneated PDF if no data is available.
			<!--Atul- place the table tag outside of the if condition-->
			if len(JAPINFO1("cand_add_info_m")) or len(JAPINFO1("cand_info_m")) then
			f.WriteLine "<p><br></p>"
				f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTJ("i_1") & "</font></h2>"
			end if
			f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
		    if len(JAPINFO1("cand_add_info_m")) or len(JAPINFO1("cand_info_m")) or len(JAPINFO1("cand_linkedin_url")) then
				f.WriteLine "<!--test18-->"
'f.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
' Atul- Make space in between th and a tag 
				
'f.WriteLine "<TR><th colspan=""1"">  <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTJ("i_1") & "'>"  & gITEXTJ("i_1") & "</a></th></tr>"
'Commented by Atul
'f.WriteLine "<TR><td valign=""top"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTJ("i_1") & "</font></h2></TD></tr>"
				
				f.WriteLine "<!--test19-->"
' '------------ TOOK OUT PDF TESTER 3 ----->
				if   len(JAPINFO1("cand_add_info_m")) then
'mar 21 07 - ac grabbed from one on UNSHARE
					f.WriteLine "<TR><td><i><font color=""Teal""><strong><i>" & gITEXTJ("i_12") & "</i></strong></font></I><strong><br>" & replace(JAPINFO1("cand_add_info_m"),VbCrLf,"<br>") & "</strong></TD></tr>"
					
'start chnage by atul remove the hr  make line with tr and table
'f.WriteLine "<TR><td valign=""top""><hr noshade size=""1""></TD></tr>"
'f.WriteLine "<tr><td valign=""top"">"
'f.WriteLine "<table height=""1px""  width=""100%""><tr style=""height:1px !important;"">"
'f.WriteLine "<td style=""background-color:black; width:2.5%;""></td>"
'f.WriteLine "</tr></table></td></tr>"
'f.WriteLine "<tr><td colspan=""1""><br></td><tr>"
					f.WriteLine "<tr border=""1"" bgcolor='" & orgColor & "' >"
					f.WriteLine "<td valign=""top"" colspan=""1"">&nbsp;</td>"
					f.WriteLine "</tr>"
'f.WriteLine "<tr><td colspan=""1""><br></td><tr>"
'End chnage by atul remove the hr  make line with tr and table
					
				end if
			f.WriteLine "<!--test20-->"
'  '------------ TOOK OUT PDF TESTER 4 ---->
				if   len(JAPINFO1("cand_info_m")) then
'  	' OLD NEEDer#Delta1# -->

					f.WriteLine "<TR valign=""top""><td valign=""top""><font color=""Teal""><i><strong>" & gITEXTJ("i_3")  & "</strong></I></FONT><br><strong>" & replace(JAPINFO1("cand_info_m"),VbCrLf,"<br>") & "</strong></TD></tr>"
'01 SEP 10 LJL remove first horizontal line if first record
'f.WriteLine "<TR><td valign=""top""><hr noshade size=""1""></TD></tr>"
			    end if
				
'18 NOV 20 LJL removed linked in from PDF for UPU per them				
	if pv_new_sessioncode = 5500 then
	else
	
				
'13 APR 16 GG added Linkedin Url	
			if len(JAPINFO1("cand_linkedin_url")) then
				f.WriteLine "<TR valign=""top""><td valign=""top""><font color=""Teal""><i><strong>" & gITEXTJ("i_71")  & "</strong></I></FONT><br><strong>" & JAPINFO1("cand_linkedin_url") & "</strong></TD></tr>"
			end if
		end if
'18 NOV 20 LJL removed linked in from PDF for UPU per them				
	end if	

		' Added by DD on 03/27/2009, change for wto only.
		f.WriteLine "<!--test22-->"
		if pv_new_sessioncode = 3000 then
		else
			f.WriteLine "<TR>"
		    f.WriteLine "<td valign=""top""> "
		    if   len(JAPINFO1("cand_law_m")) then
		        f.WriteLine "<i>" & gITEXTJ("i_13")  & "</i> <strong>"
		        f.WriteLine JAPINFO1("cand_law_m") & "</strong>"
		    else
		        f.WriteLine "&nbsp;"
		    end if
		    f.WriteLine "</td>"
		    f.WriteLine "</TR>"
		end if
		f.WriteLine "<!--test23-->"
' '--------- ADDED TO TEST PDF PROB ------------->
' '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
		if   Request.querystring("viewupd") = True and viewupd = "YES"  then
			if   GETUPDS("editG_d") > GETAPPDATE("candjob_d")  then
			f.WriteLine "<TR><td valign=""top"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") & """>Last update date <strong>" & formatdatetime(GETUPDS("EditG_d"),1) & " AFTER APPLYING</strong></TD></tr>"
			end if
		end if
		f.WriteLine "<!--test24-->"
		f.WriteLine "</table>"
' END CHECK IF PERSON SELECTED SECTION J
		end if
		f.WriteLine "<!--test25-->"
		
' ******************************************************************************
' SECTION ROT MOB RMOB
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION RM
		if instr(pv_parts,"RM,") then
' WHO ONLY 1000
		if pv_new_sessioncode = 1000 OR pv_new_sessioncode = 1200 then ' OR pv_new_sessioncode = 1500 by Gevorg
''-----------------------------------------NOTES ----------
'MODS --
'3 NOV 04 LJL moved rotmob to different table tx_rsys_candmob
'--------------------------------------------------------------->
		gCandsql = "SELECT c.cand_mob_interest_i, c.cand_mob_oloc_i, c.cand_mob_oloc1_c, c.cand_mob_oloc2_c, c.cand_mob_oloc3_c, c.cand_mob_oloc4_c, c.cand_mob_oloc5_c, c.cand_mob_special_i, c.cand_mob_special_t, aow1.aow_dsc_" & pdf_lng & "_t AS aow1, aow2.aow_dsc_" & pdf_lng & "_t AS aow2, aow3.aow_dsc_" & pdf_lng & "_t AS aow3, d5.c_" & pdf_lng & "_name AS d5, d4.c_" & pdf_lng & "_name AS d4, d3.c_" & pdf_lng & "_name AS d3, d2.c_" & pdf_lng & "_name AS d2, d1.c_" & pdf_lng & "_name AS d1 " &_
		" FROM dbo.tx_rsys_candmob c LEFT OUTER JOIN dbo.v_country_list_" & pv_new_sessioncode & " d5 ON c.cand_mob_oloc5_c = d5.who_country_code COLLATE SQL_Latin1_General_CP850_CI_AI LEFT OUTER JOIN dbo.v_country_list_" & pv_new_sessioncode & " d4 ON c.cand_mob_oloc4_c = d4.who_country_code COLLATE SQL_Latin1_General_CP850_CI_AI LEFT OUTER JOIN dbo.v_country_list_" & pv_new_sessioncode & " d3 ON c.cand_mob_oloc3_c = d3.who_country_code COLLATE SQL_Latin1_General_CP850_CI_AI LEFT OUTER JOIN dbo.v_country_list_" & pv_new_sessioncode & " d2 ON c.cand_mob_oloc2_c = d2.who_country_code COLLATE SQL_Latin1_General_CP850_CI_AI LEFT OUTER JOIN dbo.v_country_list_" & pv_new_sessioncode & " d1 ON c.cand_mob_oloc1_c = d1.who_country_code COLLATE SQL_Latin1_General_CP850_CI_AI LEFT OUTER JOIN dbo.tr_rsys_areasofwork aow3 ON c.cand_mob_aow3_c = aow3.aow_id_c LEFT OUTER JOIN dbo.tr_rsys_areasofwork aow2 ON c.cand_mob_aow2_c = aow2.aow_id_c LEFT OUTER JOIN dbo.tr_rsys_areasofwork aow1 ON c.cand_mob_aow1_c = aow1.aow_id_c 	WHERE c.cand_id_c = " & applicant_id & ""
''d set gCand =rsys_db_select.execute(gCandsql)
		Set gCand = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		gCand.Open gCandsql, rsys_db_select, 1, 1

''------------ LANG SPECIFIC ------------------------------------------------>
		f.WriteLine "<!--Test1 -->"
Dim aowname, ctyname
' MOVE VARS UP TO TOP
		f.WriteLine "<!--Test34-->"
		aowname = "q_aow_" & pdf_lng
		ctyname = "cty_name_" & pdf_lng
'Add by atul for table of content
f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTRM("i_1") & "</font></h2>"
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"		
		
		if   gCand.eof = false then
			Do while gCand.eof = false
				f.WriteLine "<TR>"
				f.WriteLine "<td height=""1""colspan=""2"">" & gITEXTRM("i_3") & "</TD>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td height=""1""colspan=""2"" ><b>" & gITEXTRM("i_4") & " " & gITEXTRM("i_5") & "</b></TD>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td height=""1""colspan=""2"">" & gITEXTRM("i_6") & " <strong>"
				if   gCand("cand_mob_interest_i") = 1  then
					f.WriteLine  gITEXTRM("i_7")
				elseif   gCand("cand_mob_interest_i") = 0  then
					f.WriteLine  gITEXTRM("i_8")
				end if
				f.WriteLine "</strong></td>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td height=""1""colspan=""2"">&nbsp;" & gITEXTRM("i_9") & "<br>"
				if   len(gCand("aow1")) then
					f.WriteLine  gITEXTRM("i_11") & " 1 <strong>" & gCand("aow1") & "</strong>"
				end if
				f.WriteLine "<br>"
				if   len(gCand("aow2")) then
					f.WriteLine  gITEXTRM("i_11") & " 2 <strong>" & gCand("aow2") & "</strong>"
				end if
				f.WriteLine "<br>"
				if   len(gCand("aow3")) then
					f.WriteLine  gITEXTRM("i_11") & " 3 <strong>" & gCand("aow3") & "</strong>"
				end if
				f.WriteLine "<br>"
				f.WriteLine "</td>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td height=""11"" colspan=""2"" class=""littletitle""><b>" & gITEXTRM("i_14") &  " " & gITEXTRM("i_15") & "</b></TD>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td height=""1""colspan=""2"">" & gITEXTRM("i_16") & " <strong>"
				if   gCand("cand_mob_oloc_i") = 1  then
					f.WriteLine  gITEXTRM("i_7")
				elseif gCand("cand_mob_oloc_i") = 0  then
					f.WriteLine  gITEXTRM("i_8")
				end if
				f.WriteLine "</strong></td>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td colspan=""2"">&nbsp; " & gITEXTRM("i_17") & "<br>"
				f.WriteLine  gITEXTRM("i_11") & " 1 "
				if   len(gCand("d1")) then
					f.WriteLine  "<strong>" & gCand("d1") & "</strong>"
				end if
				f.WriteLine "<br>" & gITEXTRM("i_11") & " 2 "
				if   len(gCand("d2")) then
					f.WriteLine  "<strong>" & gCand("d2") & "</strong>"
				end if
				f.WriteLine "<br>" & gITEXTRM("i_11") & " 3 "
				if   len(gCand("d3")) then
					f.WriteLine  "<strong>" & gCand("d3") & "</strong>"
				end if
				f.WriteLine "<br>" & gITEXTRM("i_11") & " 4 "
				if   len(gCand("d4")) then
					f.WriteLine  "<strong>" & gCand("d4") & "</strong>"
				end if
				f.WriteLine "<br>" & gITEXTRM("i_11") & " 5 "
				if   len(gCand("d5")) then
					f.WriteLine  "<strong>" & gCand("d5") & "</strong>"
				end if
				f.WriteLine "</td>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td height=""1""colspan=""2"">" & gITEXTRM("i_20") & " <strong>"
				if   gCand("cand_mob_special_i") = 1  then
					f.WriteLine  gITEXTRM("i_7")
				elseif   gCand("cand_mob_special_i") = 0  then
					f.WriteLine  gITEXTRM("i_8")
				end if
				f.WriteLine  "</strong></td>"
				f.WriteLine "</TR>"
				f.WriteLine "<TR>"
				f.WriteLine "<td height=""5"" colspan=""2"">&nbsp; " & gITEXTRM("i_21") & ": <b>" &  gCand("cand_mob_special_t") & "</b></td>"
				f.WriteLine "</TR>"
				gCand.movenext
			loop
''-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
		if  request.querystring("viewupd") = "YES"  then
			if   GETUPDS("editRMPref_d") > GETAPPDATE("candjob_d")  then
'Atul make colspan= 2 from 5
				f.WriteLine "<TR><td valign=""top"" colspan=""5"" bgcolor=""" & gITEXTPH("i_62") &"""><font color=""" & gITEXTPH("i_63") & """>Last update date <strong>" & formatdatetime(GETUPDS("editRMPref_d"),1) & " AFTER APPLYING</strong></td>"
				f.WriteLine "</TR>"
			end if
		end if
		else
			f.WriteLine "<TR><td>No rotation and mobility indicated.</TD></tr>"
		end if
			f.WriteLine "</table>"
			
			f.WriteLine "<!--Test2 -->"
' ******************************************************************************
' SECTION ROT MOB RCMP
' ******************************************************************************
' MOVE TO TOP?
		Dim pv_rower_rotmob
		f.WriteLine "<!--Test35-->"
		f.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"

		GETFACTORSsql = "SELECT TOP 100 PERCENT dbo.core_ccogfact.fact_dsc_" & new_lng_code & "_t as q_ldsc, dbo.tx_rsys_candcmp.cand_id_c, dbo.core_ccogelemt.elem_id_c, dbo.core_ccogelemt.elem_dsc_en_t AS q_elem, dbo.core_ccogelemt.elem_rem_en_m AS q_elemrem FROM dbo.core_ccogelemt INNER JOIN dbo.core_ccogfact ON dbo.core_ccogelemt.fac_id_c = dbo.core_ccogfact.fac_id_c LEFT OUTER JOIN dbo.tx_rsys_candcmp ON dbo.core_ccogelemt.elem_id_c = dbo.tx_rsys_candcmp.elem_id_c WHERE (dbo.core_ccogelemt.elem_thisorg_" & pv_new_sessioncode & " = 1) AND (dbo.tx_rsys_candcmp.cand_id_c = " & applicant_id & ") ORDER BY dbo.core_ccogfact.fact_dsc_en_t, dbo.core_ccogelemt.elem_value_t"
''d set GETFACTORS =rsys_db_select.execute(GETFACTORSsql)
		Set GETFACTORS = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		GETFACTORS.Open GETFACTORSsql, rsys_db_select, 1, 1

		if  getfactors.eof = false then
			f.WriteLine "<br>"
			pv_rower_rotmob = 0
			' Grouped Output is not currently supported by CFM2ASP.
			Dim pv_factor
			pv_factor = ""
			Do while GETFACTORS.eof = false
				if pv_factor <> GETFACTORS("q_LDSC") then
					pv_factor = GETFACTORS("q_LDSC")
					f.WriteLine "<TR>"
					f.WriteLine "<td><b>" & pv_factor & "</b></TD>"
					f.WriteLine "</TR>"
				end if
				f.WriteLine "<TR>"
				f.WriteLine "<td>&nbsp; &nbsp; &nbsp; <strong>" & gETFACTORS("q_elemrem") & "</strong></TD>"
				pv_factor = GETFACTORS("q_LDSC")
				GETFACTORS.movenext
			loop
			f.WriteLine "</TR>"
'        '-------------------------------------- TO SHOW UPDATES SINCE APPLICATION TO POST 15 MAR 05 LJL per WTO reqs ---------------------------------------------->
			if Request.querystring("viewupd") = "YES"  then
				if GETUPDS("editRMCmp_d") > GETAPPDATE("candjob_d")  then
					f.WriteLine "<TR>"
					f.WriteLine "<td valign=""top"" colspan=""1"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") & """>Last update date <strong>" & formatdatetime(GETUPDS("EditRMCmp_d"),1) & " AFTER APPLYING</strong></td>"
					f.WriteLine "</tr>"
				end if
			end if
		else
			f.WriteLine "<TR>"
			f.WriteLine "<td>No core competencies indicated.</td>"
			f.WriteLine "</TR>"
		end if
		f.WriteLine "</table>"
' WHO ONLY 1000
	end if
' END CHECK IF PERSON SELECTED SECTION J
end if
		f.WriteLine "<!--Test36-->"
' ******************************************************************************
' SECTION COVERLETTER type 8
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION Z
		f.WriteLine "<!--Test37-->"
		if instr(pv_parts,"H") then
		covlettype = "8"
		GetDocsql = "SELECT candtext_title_t, candtext_id_c, candtext_text_en_m, candtext_type_c from tx_rsys_candtext 		WHERE cand_id_c = " & applicant_id & " AND candtext_type_c = "& covlettype &" "
		set GetDoc =rsys_db_select.execute(GetDocsql)
		Set JAPGEOLIST42 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		JAPGEOLIST42.Open JAPGEOLIST42sql, rsys_db_select, 1, 1

		if GETDOC.eof = false then
		f.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
		if GETDOC.eof = false then
		Do while GetDoc.eof = false
			
'f.WriteLine "<TR>"
'f.WriteLine "<th colspan=""1"">  "
'if   covlettype = 8  then
'f.WriteLine " <a style=""abcpdf-tag-visible: true;"" id='" &  gITEXTH("i_4") & "'></a>" 
			
'f.WriteLine gITEXTH("i_4")
'else
'f.WriteLine " <a style=""abcpdf-tag-visible: true;"" id='" &  gITEXTH("i_11") & "'></a>" 
'f.WriteLine gITEXTH("i_11")
'end if
'f.WriteLine  " - " & GetDoc("candtext_title_t")
'f.WriteLine "</th>"
'f.WriteLine "</TR>"
			f.WriteLine "<TR>"
			f.WriteLine "<td valign=""top"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>"
			if   covlettype = 8  then
				f.WriteLine gITEXTH("i_4")
			else
				f.WriteLine gITEXTH("i_11")
			end if
			f.WriteLine  " - " & GetDoc("candtext_title_t")
			f.WriteLine "</font></h2></td>"
			f.WriteLine "</TR>"
			
			f.WriteLine "<TR>"
			f.WriteLine "<td>" & replace(GETDOC("candtext_text_en_m"),VbCrLf,"<br>") & "</TD>"
			f.WriteLine "</TR>"
			GetDoc.movenext
		loop
		end if
		f.WriteLine "</table>"
		end if
' END CHECK IF PERSON SELECTED SECTION H
		end if
		f.WriteLine "<!--Test38-->"
' ******************************************************************************
' SECTION COVER LETTER type 13 POST SPECIFIC
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION X
if instr(pv_parts,"X,") then

' MOVE TO TOP?
		covlettype2 = "13"
'02 NOV 06 LJL adjusted to only get this org listing
		GetDoc2sql = "SELECT j.jobinfo_vac2_c, t.candtext_title_t, t.candtext_id_c, t.candtext_text_en_m, t.candtext_type_c from tx_rsys_candtext t INNER JOIN td_rsys_jobinfo j ON t.jobinfo_uid_c = j.jobinfo_uid_c WHERE t.cand_id_c = " & applicant_id & " AND t.candtext_type_c = "& covlettype2 &" AND j.jobinfo_thisorg_" & pv_new_sessioncode & " = 1 "
'02 NOV 06 LJL added this to weed out per job covering letters if sent from admin/hrd-cllist to admin/docsettings.asp to /doccreate/make-doc-prep.asp
		if len(pv_jobid) then
			GetDoc2sql = GetDoc2sql & " AND (t.jobinfo_uid_c = " & pv_jobid & ") "
		end if
		if vacchoice > 0 then
			GetDoc2sql = GetDoc2sql & " AND t.jobinfo_uid_c = " & vacchoice
		end if
'17 JUN 09 LJL added to diminish VN specific covering letters output if admin is VN only
	IF session("CLI_ADMIN_POSTS") = "0" then
		GetDoc2sql = GetDoc2sql & " AND  (j.jobinfo_uid_c IN (" & session("postsgroup") & "))"
	end if
	IF session("CLI_ADMIN_UNITS") = "0" then
		GETJAPJAFSsql = GETJAPJAFSsql & " AND (j.org_wk_c IN (" & session("unitsgroup") & "))"
		pv_narrowed = "1"
	end if
'08 NOV 09 LJL added Posts not to be enabled
	IF session("CLI_ADMIN_POSTSNOT") = "0" then
		GetDoc2sql = GetDoc2sql & " AND NOT (j.jobinfo_uid_c IN (" & session("postsnotgroup") & "))"
	end if
''d set GetDoc2 =rsys_db_select.execute(GetDoc2sql)
		Set GetDoc2 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
		rsys_db_select.CommandTimeout = 320
		GetDoc2.Open GetDoc2sql, rsys_db_select, 1, 1

		if GETDOC2.eof = false then
			f.WriteLine "<p><br></p>"
			f.WriteLine "<h2><font color=""" & g_headerColor & """>"
				if covlettype2 = 8  then
					f.WriteLine gITEXTH("i_4")
				else
					f.WriteLine gITEXTH("i_11")
				end if
			f.WriteLine " - " & GetDoc2("candtext_title_t")
			if len(GETDOC2("jobinfo_vac2_c")) then
				f.WriteLine " - " & GETDOC2("jobinfo_vac2_c")
			end if
			f.WriteLine "</font></h2>"
			
			f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
			Do while GetDoc2.eof = false
'f.WriteLine "<tr><th colspan=""1""> "
'if   covlettype2 = 8  then
' 		f.WriteLine " <a style=""abcpdf-tag-visible: true;"" id='" &  gITEXTH("i_4") & "_|" & intNumber  &  "'></a>" 
'	        f.WriteLine gITEXTH("i_4")
'	else
'   	f.WriteLine " <a style=""abcpdf-tag-visible: true;"" id='" &  gITEXTH("i_11") & "_|" & intNumber  &  "'></a>" 
'	f.WriteLine gITEXTH("i_11")
'end if
'f.WriteLine " - " & GetDoc2("candtext_title_t")
'if len(GETDOC2("jobinfo_vac2_c")) then
'	f.WriteLine " - " & GETDOC2("jobinfo_vac2_c")
'end if
'f.WriteLine "</th></tr><tr><td>"
			
			f.WriteLine "<tr><td>"
			
			f.WriteLine  replace(GETDOC2("candtext_text_en_m"),VbCrLf,"<br>") & "</td></tr>"
			GetDoc2.movenext
		loop
		f.WriteLine "</table>"
		end if
' END CHECK IF PERSON SELECTED SECTION X
		end if
		f.WriteLine "<!--Test40-->"
' ******************************************************************************
' SECTION DOCS other
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION Y
		f.WriteLine "<!--Test41-->"
if instr(pv_parts,"Y,") then

	GetDoc3sql = "SELECT candtext_title_t, candtext_id_c, candtext_text_en_m, candtext_type_c 		from tx_rsys_candtext 		WHERE cand_id_c = " & applicant_id & " 		AND candtext_type_c NOT IN (1,4,8,13) 	"
''d set GetDoc3 =rsys_db_select.execute(GetDoc3sql)
	Set GetDoc3 = Server.CreateObject("ADODB.RecordSet")
'26 feb 07 ac increase db timeout
	rsys_db_select.CommandTimeout = 320
	GetDoc3.Open GetDoc3sql, rsys_db_select, 1, 1

	Do while GetDoc3.eof = false
'ac - removedf.WriteLine "<TABLE border=""0"" bordercolor=""black"" cellpadding=""2"" width=""100%"" align=""center"" border=""0"">"
		f.WriteLine "<p><br></p>"
		f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTH("i_5") & " -  " &  GETDOC3("candtext_title_t") & "</font></h2>"
		f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" & widther & """ align=""center"">"
			
	'f.WriteLine "<tr><th colspan=""1""> <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTH("i_5") & "_|" & intNumber  &  "'></a>" & gITEXTH("i_5") & " -  " &  GETDOC3("candtext_title_t") & "</th></tr>"
		'f.WriteLine "<tr><td valign=""top"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTH("i_5") & " -  " &  GETDOC3("candtext_title_t") & "</font></h2></td></tr>"
		f.WriteLine "<tr><td>"& replace(GETDOC3("candtext_text_en_m"),VbCrLf,"<br>") & "</TD></tr>"
		GetDoc3.movenext
		f.WriteLine "</table>"
		f.WriteLine "<!--Test42-->"
	loop
' END CHECK IF PERSON SELECTED SECTION Y
end if

' ******************************************************************************
' SECTION OTHER INFORMATION
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION OI
       '//if instr(pv_parts,"OI,") then
if instr(pv_parts,"OI,") AND pv_new_sessioncode = 3000  then
'commented by atul for table of content
f.WriteLine "<p><br></p>"
			f.WriteLine "<h2><font color=""" & g_headerColor & """>" & "Other Information" & "</font></h2>"
			
   			f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'f.WriteLine "<TR><th colspan=""1""> " & "Other Information" & "</th></tr>"
'Commented by atul for table of content
'f.WriteLine "<TR><td valign=""top"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & "Other Information" & "</font></h2></TD></tr>"

             //Que1
            //07 OCT 09 DD Modified for Other Information section to contain the questions and answers even in the case of "No".
		    //if JAPINFO1("cand_geo_res_i") <> 2 and JAPINFO1("cand_geo_res_i") > 0 then
		    if JAPINFO1("cand_geo_res_i") > 0 then
Dim GETDROP3sql, GETDROP3
		        GETDROP3sql = " SELECT drop_id_c AS code, drop_dsc_"& session("lng") &"_t AS dropname FROM tr_rsys_drop WHERE (NOT drop_dsc_"& session("lng") & "_t IS NULL) AND drop_thisorg_" & pv_new_sessioncode & " = 1 AND drop_type_i = 1 And drop_id_c =" & JAPINFO1("cand_geo_res_i")
                set GETDROP3 =rsys_db_select.execute(GETDROP3sql)
		        f.WriteLine "<TR>"
		        f.WriteLine "<td valign=""top"">"
		        f.WriteLine "<i>" & gITEXTJF2("i_text2") & "<br></i><strong>"
                f.WriteLine GETDROP3("dropname") & "</strong><br>"
                if len(trim(JAPINFO1("cand_geo_res_m"))) AND  JAPINFO1("cand_geo_res_i")= 1 then
                 //'14 OCT 09 DD Added the "if yes... " text part into the CV output
                 f.WriteLine "<br>" & gITEXTJF2("i_87") & "<br>"
                 f.WriteLine "<strong>" & JAPINFO1("cand_geo_res_m") & "</strong>"
                end if
                f.WriteLine "</td>"
		        f.WriteLine "</TR>"
		    end if
		    //Que2
		    f.WriteLine "<TR>"
		    f.WriteLine "<td valign=""top"">"

		    f.WriteLine "<i>" & gITEXTJV2("i_11")  & "<br></i>"

			//07 OCT 09 DD Modified for Other Information section to contain the questions and answers even in the case of "No".
Dim gITEXTYESNOsql, gITEXTYESNO
			gITEXTYESNOsql = "SELECT i_18, i_19 FROM tr_rsys_itext WHERE itext_thisorg_c = ? AND itext_lng_c = ? AND itext_page_c = 'V' "
			obj_int_select_CmdII.CommandText = gITEXTYESNOsql
			Set gITEXTYESNO = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))

			if JAPINFO1("cand_law_i")= 1 then
				f.WriteLine "<strong>" & gITEXTYESNO("i_18") & "</strong><br>"
			else
			 	f.WriteLine "<strong>" & gITEXTYESNO("i_19") & "</strong><br>"
			end if

		    if len(trim(JAPINFO1("cand_law_m"))) AND  JAPINFO1("cand_law_i")= 1 then
		       // f.WriteLine "<i>" & gITEXTJV2("i_11") & "<br></i><strong>"
			   // f.WriteLine JAPINFO1("cand_law_m") & "</strong>"
			   //'14 OCT 09 DD Added the "if yes... " text part into the CV output
			   f.WriteLine "<br>" & gITEXTJV2("i_12") & "<br>"
			   f.WriteLine "<strong>" & JAPINFO1("cand_law_m") & "</strong>"
		    else
			    f.WriteLine " "
		    end if
		    f.WriteLine "</td>"
		    f.WriteLine "</TR>"
		    f.WriteLine "</table>"
        end if
		f.WriteLine "<!--Test44-->"
' END CHECK IF PERSON SELECTED SECTION OI

' ******************************************************************************
' SECTION T Secretarial Skills   Clerical Skills
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION T	
if instr(pv_parts,"T,") then
	Dim gITEXTT2sql, gITEXTT2
	gITEXTT2sql=" SELECT i_1 FROM tr_rsys_itext WHERE itext_thisorg_c = ?  AND itext_lng_c = ? AND itext_page_c = 'T' "
	obj_int_select_CmdII.CommandText = gITEXTT2sql
	set gITEXTT2 = obj_int_select_CmdII.Execute(,Array(pv_new_sessioncode,session("lng")))

	f.WriteLine "<!--Test45-->"
'Add by atul for table od content
f.WriteLine "<p><br></p>"
			f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTT2("i_1") & "</font></h2>"
	        f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'Atul make  space and colspan=7 from 6
	        
'f.WriteLine "<TR><th colspan=""7""> <a style=""abcpdf-tag-visible: true;"" id='" & gITEXTT2("i_1") & "_|" & intNumber  &  "'>" & gITEXTT2("i_1") & "</a></th></tr>"
'Commented by atul for table of content
'f.WriteLine "<TR><td valign=""top"" colspan=""7"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & gITEXTT2("i_1") & "</font></h2></TD></tr>"
	        
            f.WriteLine "<tr>"
'Atul make  colspan=7 from 6
            f.WriteLine "<td valign=""bottom"" class=""textitalic"" colspan=""7"">" & gITEXTT("i_4")& "</td>"
            f.WriteLine "</tr>"

	        
'08 MAR 11 LJL added UN Typing test section for UN agencies
	if pv_new_sessioncode <> 3000 AND pv_new_sessioncode <> 7000 then
	        f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" colspan=""3"">" & gITEXTT("i_5")  & " <strong>" 
'f.WriteLine "|" & JAPINFO1("candlng_un_typ_i") & "|"
'30 MAY 11 LJL changed candlng_un_typ_i to a string instead of int
            if JAPINFO1("candlng_typ_un_i") = "1" then
           		 f.WriteLine gITEXTT("i_6") 
           	else
           		 f.WriteLine gITEXTT("i_7")            	
            end if
   		    f.WriteLine "</strong></td>"
            f.WriteLine "<td valign=""top"" colspan=""2"">" & gITEXTT("i_8") & " <strong>" & JAPINFO1("candlng_typ_un_y") & "</strong></td>"
            f.WriteLine "<td valign=""top"" colspan=""2"">" & gITEXTT("i_9") & " <strong>" & JAPINFO1("candlng_typ_un_t") & "</strong></td>"

            f.WriteLine "</tr>"

	end if
	        


	f.WriteLine "<tr>"
	f.WriteLine "<td valign=""top""  colspan=""1""></td>"
	f.WriteLine "<td valign=""top"" colspan=""2"">" & gITEXTT("i_10") & "</td>"
	'f.WriteLine "<td valign=""top""  colspan=""1""></td>"
	f.WriteLine "<td valign=""top"" colspan=""2"">" & gITEXTT("i_11") & "</td>"
	'f.WriteLine "<td valign=""top""  colspan=""1""></td>"
	if pv_new_sessioncode = 2400 then
		f.WriteLine "<td valign=""top"">" & gITEXTT("i_21") & "</td>"
		f.WriteLine "<td valign=""top""  colspan=""1""></td>"
	else
		 f.WriteLine "<td valign=""top""  colspan=""1""></td>"
		 f.WriteLine "<td valign=""top""  colspan=""1""></td>"
	end if
	f.WriteLine "</tr>"

'08 SEP 10 LJL revised order of languages for ITU
if session("lng") = "en" then

	if pv_new_sessioncode = 2400 then
            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_40") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "</tr>"

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_41") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "</tr>"
	end if

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_12") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
				f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_au_i") & "</strong></td>"
				f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			else
				f.WriteLine "<td valign=""top""  colspan=""1""></td>"
				f.WriteLine "<td valign=""top""  colspan=""1""></td>"
			end if
            f.WriteLine "</tr>"

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_13") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
			f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			else
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
'Atul-Add one column
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
			end if
            f.WriteLine "</tr>"

			if pv_new_sessioncode = 2400 then
			        f.WriteLine "<tr>"
			        f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_42") & "</td>"
			        f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_typ_i") & "</strong></td>"
					f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			        f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_sh_i") & "</strong></td>"
					f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			        if pv_new_sessioncode = 2400 then
					f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_au_i") & "</strong></td>"
					f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
					else
					f.WriteLine "<td valign=""top""  colspan=""1""></td>"
'Atul-  Add one column
					f.WriteLine "<td valign=""top""  colspan=""1""></td>"
					end if
			        f.WriteLine "</tr>"
			end if


            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_14") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_sh_i")  & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
			f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			else
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
'Atul- Add one column
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
			end if
            f.WriteLine "</tr>"



elseif session("lng") = "fr" then

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_12") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
			f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			else
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
' Atul Add one column
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
			end if
            f.WriteLine "</tr>"

	if pv_new_sessioncode = 2400 then
            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_40") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "</tr>"

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_41") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "</tr>"
	end if

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_14") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_sh_i")  & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
			f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_au_i")  & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			Else
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
' Atul- Add one new column
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
            End if
            f.WriteLine "</tr>"

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_13") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
			f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			Else
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
'Atul add one new  column
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
            End if
            f.WriteLine "</tr>"

	if pv_new_sessioncode = 2400 then
            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_42") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "</tr>"
	end if

elseif session("lng") = "es" then

	if pv_new_sessioncode = 2400 then
            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_40") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ar_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "</tr>"

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_41") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_cn_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "</tr>"
	end if

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_12") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
			f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_en_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			Else
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
' Atul- Add one new column
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
            End if
            f.WriteLine "</tr>"

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_13") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
			f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_fr_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			Else
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
'Add one new column
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
            End if
            f.WriteLine "</tr>"

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_14") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_sh_i")  & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            if pv_new_sessioncode = 2400 then
			f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_es_au_i")  & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
			Else
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
' Atul Add one new column
			f.WriteLine "<td valign=""top""  colspan=""1""></td>"
            End if
            f.WriteLine "</tr>"

	if pv_new_sessioncode = 2400 then
            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" class=""textbold"">" & gITEXTT("i_42") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_typ_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_sh_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "<td valign=""top""><strong>" & JAPINFO1("candlng_ru_au_i") & "</strong></td>"
            f.WriteLine "<td valign=""top"">" & gITEXTT("i_15") & "</td>"
            f.WriteLine "</tr>"
	end if


end if
            f.WriteLine "<tr>"
'Atul make colspan=7 from from 4
            f.WriteLine "<td valign=""top"" class=""textbold"" colspan=""7"">" & gITEXTT("i_16") & "</td>"
            f.WriteLine "</tr>"

            f.WriteLine "<tr>"
            f.WriteLine "<td valign=""top"" colspan=""1"">&nbsp;</td>"
            f.WriteLine "<td valign=""top"" colspan=""2""><strong>" & JAPINFO1("candlng_ol_typ_t") & "</strong></td>"
            f.WriteLine "<td valign=""top"" colspan=""2""><strong>" & JAPINFO1("candlng_ol_sh_t") & "</strong></td>"
            if pv_new_sessioncode = 2400 then
				f.WriteLine "<td valign=""top"" colspan=""2""><strong>" & JAPINFO1("candlng_ol_au_i") & "</strong></td>"
			Else
'Atul- make colspan=2 from 1
				f.WriteLine "<td valign=""top""  colspan=""2""></td>"
            End if

            f.WriteLine "</tr>"
            f.WriteLine "</table>"

		end if
' END CHECK IF PERSON SELECTED SECTION T

'f.WriteLine "<h2>gevorg</h2>"
' ******************************************************************************
' SECTION REFERENCES - WTO ONLY
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION GR
		if pv_new_sessioncode = 3000 then
			if instr(pv_parts,"GR,") then
'			response.Write "pv_parts=" & pv_parts & "<br>"
'			response.End
Dim getRefListSql, getRefList
            getRefListSql="SELECT candrefer_fax_c, candrefer_email_t, candrefer_nam_t, candrefer_adr_t, candrefer_phn_t, candrefer_occ_t, candrefer_cand_c, candrefer_known_t, candrefer_contact_i, candrefer_id_c from tx_rsys_candrefer WHERE candrefer_cand_c = " & applicant_id
            set getRefList = Server.CreateObject("ADODB.RecordSet")
'increase db timeout
		    rsys_db_select.CommandTimeout = 320
		    getRefList.Open getRefListSql, rsys_db_select, 1, 1
'set getRefList = rsys_db_select.execute(getRefListSql)
'Added by atul for table of content
f.WriteLine "<p><br></p>"
			f.WriteLine "<h2><font color=""" & g_headerColor & """>" & "References" & "</font></h2>"
   			f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"
'f.WriteLine "<TR><th colspan=""3""> " & "References" & "</th><</tr>"
'Commented by atul for table od content
'f.WriteLine "<TR><td colspan=""3"" valign=""top"" bgcolor=""" & gITEXTPH("i_60") & """><h2><font color=""" & g_headerColor & """>" & "References" & "</font></h2></TD></tr>"

			f.WriteLine "<TR class=""trheader10"">"
            f.WriteLine "<Td align=""left"" valign=""top"">" & gITEXTGRef("i_5")  & "</td>"
            f.WriteLine "<Td align=""left"" valign=""top"">" & gITEXTGRef("i_17") & "<br>" & gITEXTGRef("i_18") & "<Br>" & gITEXTGRef("i_80") & "</td>"
            f.WriteLine "<Td align=""left"" valign=""top"">" & gITEXTGRef("i_75") & "<br>" & gITEXTGRef("i_19") & "</td>"

            Do while getRefList.eof = false
                f.WriteLine "<TR>"

		        f.WriteLine "<td valign=""top"">"
		        if len(getRefList("candrefer_nam_t")) then
		            f.WriteLine getRefList("candrefer_nam_t")
		        end if
		        f.WriteLine "</td>"

		        f.WriteLine "<td valign=""top"">"
		        if len(getRefList("candrefer_adr_t")) then
		            f.WriteLine getRefList("candrefer_adr_t")  & "<br>"
		        end if
		        if len(getRefList("candrefer_phn_t")) then
		            f.WriteLine getRefList("candrefer_phn_t") & "<br>"
		        end if
		        if len(getRefList("candrefer_fax_c")) then
		            f.WriteLine getRefList("candrefer_fax_c")
		        end if
		        f.WriteLine "</td>"

		        f.WriteLine "<td valign=""top"">"
		        if len(getRefList("candrefer_email_t")) then
		            f.WriteLine getRefList("candrefer_email_t") & "<Br>"
		        end if
		        if len(getRefList("candrefer_occ_t")) then
		         f.WriteLine getRefList("candrefer_occ_t")
		        end if
		        f.WriteLine "</td>"
		        f.WriteLine "</TR>"
            getRefList.MoveNext
            loop

		    f.WriteLine "</table>"
            end if
        end if

' END CHECK IF PERSON SELECTED SECTION GR

' ******************************************************************************
' SECTION VERIFICATION
' ******************************************************************************
' BEGIN CHECK IF PERSON SELECTED SECTION V
	if instr(pv_parts,"V,") then
    f.WriteLine "<p><br></p>"
    f.WriteLine "<h2><font color=""" & g_headerColor & """>" & gITEXTV("i_1") & "</font></h2>"
    f.WriteLine "<TABLE border=""1"" bordercolor=""black"" cellpadding=""2"" width=""" &  widther & """ align=""center"">"

    ' Sexual misconduct
    f.WriteLine "<TR><td valign=""top"">"
    f.WriteLine gITEXTV("i_97")
    f.WriteLine "<p><br></p>"
    if trim(JAPY("cand_sexual_i")) = 1 then  
        f.WriteLine "<b>" & gITEXTV("i_18")  & ", " & JAPY("cand_sexual_m") & "</b>"
    else
        f.WriteLine "<b>" & gITEXTV("i_19") & "</b>"
    end if
    f.WriteLine "<p><br></p></td></TR>"

    ' Law issue
    f.WriteLine "<TR><td valign=""top"">"
    f.WriteLine gITEXTV("i_11")
    f.WriteLine "<p><br></p>"
    if trim(JAPY("cand_law_i")) = "1" then  
        f.WriteLine "<b>" & gITEXTV("i_18")  & ", " & JAPY("cand_law_m") & "</b>"
    else
        f.WriteLine "<b>" & gITEXTV("i_19") & "</b>"
    end if
    f.WriteLine "<p><br></p></td></TR>"

    ' Terminated
    f.WriteLine "<TR><td valign=""top"">"
    f.WriteLine gITEXTV("i_98")
    f.WriteLine "<p><br></p>"
    if trim(JAPY("cand_teriminated_i")) = "1" then  
        f.WriteLine "<b>" & gITEXTV("i_18")  & ", " & JAPY("cand_terminated_m") & "</b>"
    else
        f.WriteLine "<b>" & gITEXTV("i_19") & "</b>"
    end if
    f.WriteLine "<p><br></p></td></TR>"

    ' Dismissed
    f.WriteLine "<TR><td valign=""top"">"
    f.WriteLine gITEXTV("i_94")
    f.WriteLine "<p><br></p>"
    if trim(JAPY("cand_dismissed_i")) = "1" then  
        f.WriteLine "<b>" & gITEXTV("i_18")  & ", " & JAPY("cand_dismissed_m") & "</b>"
    else
        f.WriteLine "<b>" & gITEXTV("i_19") & "</b>"
    end if
    f.WriteLine "<p><br></p></td></TR>"

    ' Resigned
    f.WriteLine "<TR><td valign=""top"">"
    f.WriteLine gITEXTV("i_95")
    f.WriteLine "<p><br></p>"
    if trim(JAPY("cand_resigned_i")) = "1" then  
        f.WriteLine "<b>" & gITEXTV("i_18")  & ", " & JAPY("cand_resigned_m") & "</b>"
    else
        f.WriteLine "<b>" & gITEXTV("i_19") & "</b>"
    end if
    f.WriteLine "<p><br></p></td></TR>"

    ' Name included
    f.WriteLine "<TR><td valign=""top"">"
    f.WriteLine gITEXTV("i_96")
    f.WriteLine "<p><br></p>"
    if trim(JAPY("cand_nameinclude_i")) = "1" then  
        f.WriteLine "<b>" & gITEXTV("i_18")  & " , " & JAPY("cand_nameinclude_UN") & "</b>"
    else
        f.WriteLine "<b>" & gITEXTV("i_19") & "</b>"
    end if
    f.WriteLine "</td></TR>"

    ' Additional information
    f.WriteLine "<TR><td valign=""top"">" & gITEXTV("i_text1") & "</td></TR>"
    f.WriteLine "<TR><td valign=""top"" align=""left"">" & gITEXTV("i_3") & ": <strong>"
    if len(JAPINFO1("cand_fil_d")) then
        f.WriteLine day(JAPINFO1("cand_fil_d")) & " " & monthname(month(JAPINFO1("cand_fil_d")),2) & " " & year(JAPINFO1("cand_fil_d"))
    end if
    f.WriteLine "</strong></td></TR>"

    if pv_new_sessioncode <> 3000 then
        f.WriteLine "<TR><td valign=""top"" align=""left"">" & gITEXTV("i_4") & ": <strong>" & JAPINFO1("cand_fil_t") & "</strong></td></TR>"
    end if

    f.WriteLine "<TR><td valign=""top"" align=""left"">" & gITEXTV("i_10") & ": <strong>" & JAPINFO1("cand_verif_name_t") & "</strong></td></TR>"
    f.WriteLine "<TR><td valign=""top"">&nbsp;</td></TR>"

    if Request.querystring("viewupd") = "YES" then
        if GETUPDS("editFooter_d") > GETAPPDATE("candjob_d") then
            f.WriteLine "<TR><td valign=""top"" bgcolor=""" & gITEXTPH("i_62") & """><font color=""" & gITEXTPH("i_63") & """>Last update date <strong>" & formatdatetime(GETUPDS("EditFooter_d"),1) & " AFTER APPLYING</strong></td></TR>"
        end if
    end if

    f.WriteLine "</TABLE>"
end if

		if pv_multi = "1" then
'02 NOV 06 LJL add new page break here as in CFM version and per htmldoc instruction manual PDF www.easysw.com
		f.WriteLine "<!-- FOOTER CENTER """ & VN & """ -->"
		f.WriteLine "<!-- NEW PAGE -->"
		end if

		if pv_DEBUG then
			response.write "<br>T liner: " & liner
			response.write "<br>stop 3.99-" & liner
'response.End()
		end if

	JAPINFO1.movenext
	loop
	response.write "<br>" & liner-1 & " records processed..."
rsys_int_select.close
Set rsys_int_select = nothing
rsys_db_select.close
Set rsys_db_select = nothing
rsys_db_select1.close
Set rsys_db_select1 = nothing
%>
</body> </html>
<%
if pv_DEBUG then
	response.write "<br>T Finish: " & request.form("app")
	response.write "<br>stop 3.4"
    'response.End()
end if
%>
