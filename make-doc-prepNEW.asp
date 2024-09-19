<%@ Language = VBScript%>
<% Option Explicit
'15 AUG 06 LJL added TOC (table of contents) setting
'01 NOV 06 LJL REM'd out some of the testing information so not to confuse admins - tested for new internal data in UNAIDS
'02 NOV 06 LJL added fm_jobid TO SEND ACROSS THE JOB ID TO REDUCE COVERING LETTERS IN APPLICANT CV OUTPUT
'21 nov 06 AC change file to point to ACdoc-maker.asp
'30 Nov 06 what type of file format does the user want?
'16 dec 06 ac to handle document format page title request by public/pdf

'20 DEC 06 LJL added if/then to be able to do from admin site
'19 feb 07 ac - added new temp /public_hold when fixed changed back /public
'22 feb 07 ac - addedILO db
'23 feb 07 - ac added new var to reduce width of file delete code line
' 20 apr 07 ac - file locations changed 
'26 apr 07 ac - block new code revert to old file locs until duplicates removed
'16 may 07 ac - file locations changed - duplicates docs removed from server and db
'22 FEB 08 LJL ONLY for public
'08 SEP 09 LJL revised the directories for orgs
'16 SEP 09 DD increase the level of compression of generated pdf if photo includes in cv
'09 MAY 10 LJL changed pv_DEBUG settings
'07 AUG 10 LJL change ITU and WIPO to prod db and added demo spec for dirs
'04 OCT 10 LJL revised to include demo in directory string - for demo site
'12 NOV 14 LJL added further text about where PDF will go for user
'05 DEC 14 LJL added UPU for db
'05 DEC 14 LJL added WMO for db
'12 DEC 14 GG added Usepix parameter to the GeneratePDF.aspx Url
'17 DEC 14 GG added viewformat parameter to the GeneratePDF.aspx Url
'23 DEC 14 GG added pageinfo ApplicantName parametera to the GeneratePDF.aspx Url
'23 DEC 14 GG implemented first info page 
'23 DEC 14 LJL added UNSHARE to stage dir
'03 MAR 15 LJL revised directories for WTO, IFRC for HTML file creation
'31 DEC 15 GG Fixed file path
'16 DEC 20 LJL uses /admin/ACdoc-maker-new.asp and not public/pdf/acdoc-maker.asp
'16 DEC 20 LJL this is where word also is generated.  The warnings are in English and don't show for WORD as it is immediately processed in
'public-view-f-test.asp
'16 DEC 20 LJL tried a number of things to get correct accented chars outputted.  Ended up having to just add <head><meta http-equiv='Content-Type' content='text/html; charset=ISO-8859-1'> to admin/ACdoc-maker-new.asp for all outputs via the file stream write



' ***********************************************************
' BEGIN INCLUDES
' ***********************************************************
' public/pdf folder
'20 DEC 06 lJL added check for login%>
<!--#include file="../includes/include_pubedit_frame_top.asp"-->
<!--#include file="../includes/include_check_login_noform.asp"-->

<%
' Admin/pdf folder only----include file=&quot;../includes/include_ext_frame_top.asp %>
<%
' ***********************************************************
' END INCLUDES
' ***********************************************************



dim send_ids, doc_page, result_page, page_dest, dest_page, pv_parts, pv_fullcode, pv_htmldoc, pv_type, page_start, dest_dir, dir_path, pv_tocset,pv_usepix, pv_pagsize, pv_doc_type, pv_newdb

dim pv_jobid, vacchoice, multicode, multi, org_code, ids, pv_DEBUG, dest_finalfile, dir_current, dest_file, LA, applicant_id

dim obj_int_select_Cmd, obj_db_select_Cmd, faqid

Dim orgColor,Vnname,VnClosing,ApplicantName,PdfFileName
ApplicantName = ""
PdfFileName = ""

'dim pv_page_title

'pv_DEBUG = 1



pv_page_title = "CV CREATION"
'21 nov 06 ac changed to set test db on stage only
'04 OCT 10 LJL revised to include demo in directory string - for demo site
dir_current = UCASE(request.servervariables("PATH_TRANSLATED"))
'19 feb 07 - ac add db change for WHO

'04 OCT 10 LJL revised to include demo in directory string - for demo site
if session("template_org_code") = 1000 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_unshare"
	end if
elseif session("template_org_code") = 1200 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_test"
	end if
elseif session("template_org_code") = 1500 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_unshare"
	end if
elseif session("template_org_code") = 2000 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		pv_newdb = "erec_unshare"
	end if
elseif session("template_org_code") = 2400 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_unshare"
	end if
elseif session("template_org_code") = 2500 then
	pv_newdb = "erec_test"
	
elseif session("template_org_code") = 2600 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_test"
	end if

elseif session("template_org_code") = 2700 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_test"
	end if
elseif session("template_org_code") = 2800 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_unshare"
	end if

'05 DEC 14 LJL added WMO for db
elseif session("template_org_code") = 2900 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_unshare"
	end if



elseif session("template_org_code") = 3000 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		'test db
		pv_newdb = "erec_test"
	else
		pv_newdb = "erec_wto"
	end if
elseif session("template_org_code") = 3200 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_test"
	end if
elseif session("template_org_code") = 4000 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		'test db
		pv_newdb = "erec_test"
	else
		pv_newdb = "erec_test"
	end if
elseif session("template_org_code") = 5000 then
	pv_newdb = "erec_test"

'05 DEC 14 LJL added UPU for db
elseif session("template_org_code") = 5500 then
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_unshare"
	end if

elseif session("template_org_code") = 6000 then
	pv_newdb = "erec_test"
elseif session("template_org_code") = 7000 then
	'06 nov 06 ac changed to set test db on stage only
	' if "stage" found in path then file is in developemnt stage
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		' test db
		pv_newdb = "erec_test"
	else
		' live db
		pv_newdb = "erec_ifrc"
	end if
elseif session("template_org_code") = 8000 then
	pv_newdb = "erec_test"
elseif session("template_org_code") = 9000 then
	pv_newdb = "erec_test"
end if

if pv_DEBUG = 1 then
	response.write "<br>admin.docmaker pv_newdb=" & pv_newdb 
	response.write "<br>stop 3.33"
	'response.End()
end if
	
' DB IS DIMMED IN PAGE OR IN OTHER PAGE INCLUDE
'20 DEC 06 LJL in the login check page
' dim rsys_db_select
dim rsys_db_select1
Set rsys_db_select = Server.CreateObject("ADODB.Connection")
Set rsys_db_select1 = Server.CreateObject("ADODB.Connection")

rsys_db_select.Open="Provider=sqloledb.1;Data Source=192.168.106.4;Initial Catalog=" & pv_newdb & ";User Id=read_human;Password=pass1ve"
rsys_db_select1.Open="Provider=sqloledb.1;Data Source=192.168.106.4;Initial Catalog=" & pv_newdb & ";User Id=read_human;Password=pass1ve"

' DB IS DIMMED IN PAGE OR IN OTHER PAGE INCLUDE
'20 DEC 06 LJL in the login check page
' dim rsys_int_select
Set rsys_int_select = Server.CreateObject("ADODB.Connection")
rsys_int_select.Open="Provider=sqloledb;Data Source=192.168.106.4;Initial Catalog=rsys_int;User Id=read_human;Password=pass1ve"


' output document type
pv_doc_type = ""
'30 Nov 06 what type of file format does the user want?
if request.form("goPDF" ) <> "" then  
		pv_page_title = "PDF CREATION"
		pv_doc_type = "pdf" 
		widther = "650"
elseif request.form("goWord" ) <> "" then  
		pv_page_title = "WORD DOCUMENT CREATION"
		pv_doc_type = "doc" 
		widther = "100%"
elseif request.form("goHTML" ) <> "" then  
		pv_page_title = "HTML CREATION"
		pv_doc_type = "html" 
		widther = "100%"
else 
	pv_doc_type = "pdf" 
	widther = "650"
end if

' default do not print pic
pv_usepix = "0"
' asp page to create file
doc_page = "ACdoc-maker.asp"
' holds post number used to reduce the number of cover letters pulled
pv_jobid = ""
' set to pv_newcode
multicode = ""
' set to > "" if cvs to produce number > 1 
multi = ""
' set page size A4 etc
pv_pagsize = "" 

' sections of CV to output to file
pv_parts = ""
' default = 1 provide link to output file
' 0 = open with binary read when done 
pv_type = "0"
' asp file that creates html, word or pdf file
doc_page = ""
' destination path for pdf generator to dump file
'dest_page = "admin\pdf\"
' name of output file
result_page = ""
' pdf genrator code - incldue or exclude TOC
pv_tocset = ""

ids = "" ' using to pass publ cand id
'for all files
if request.form("format") <> "" then
	pv_pagsize = request.form("format")
end if 

	
	
' ac added
if session("template_org_code") <> "" then
org_code = session("template_org_code")
else 
	response.write "NO ORG CODE, ENDING"
	response.end
end if


' if "stage" found in path then file is in developemnt stage
' 20 apr 07 ac - file locations changed
' new off-site download doc sections for each organization. 
dir_current = UCASE(request.servervariables("PATH_TRANSLATED"))

'if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
'	' new folder location repalces docstore/vac-cv while in developement
'	dest_dir = "E:\stage\" & session("template_org_code") & "_docs\short\"
'else
	' new folder location repalces docstore/vac-cv while in developement
'	dest_dir = "E:\docs\" & session("template_org_code") & "_docs\short\"
'end if

' if "stage" found in path then file is in developemnt stage
'03 MAR 15 LJL revised directories for WTO, IFRC for HTML file creation
if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
	'new folder location for IFRC
	'if session("template_org_code") = 7000 then
		'10 DEC 14 LJL changed
		'dest_dir = "E:\docs\stage\IFRC\short\"
		dest_dir = "E:\docs\stage\IFRC\PDFmake\PDF-CV\"
	'new folder location for WTO
	'elseif session("template_org_code") = 3000 then
	'	'10 DEC 14 LJL changed
'		dest_dir = "E:\docs\stage\WTO\short\"
	'	dest_dir = "E:\docs\stage\WTO\PDFmake\PDF-CV\"
	' new folder location WHO, UNAIDS & ILO
	'else
		'10 DEC 14 LJL changed
		'23 DEC 14 LJL added UNSHARE to stage dir
		dest_dir = "E:\docs\stage\ALLORG\PDFmake\PDF-CV\"
		'dest_dir = "E:\docs\stage\UNSHARE\PDFmake\PDF-CV\"
	'end if
' production site storage file locations 
elseif session("template_org_code") = 3000 then
	'  new folder location WTO (Production site)
'	dest_dir = "E:\docs\WTO\short\"
	dest_dir = "E:\docs\WTO\PDFmake\PDF-CV\"
elseif session("template_org_code") = 7000 then
	' new folder location IFRC (Production site)
	'dest_dir = "E:\docs\IFRC\short\"				
	dest_dir = "E:\docs\IFRC\PDFmake\PDF-CV\"
else
	' new folder location WHO, UNAIDS & ILO (Production sites)
	dest_dir = "E:\docs\UNSHARE\PDFmake\PDF-CV\"
end if



'22 ac Nov 06
'if instr(session("CLI_RSYS_ADMIN_USER"),"UNAIDS-LUNDBERGL") or  instr(session("CLI_RSYS_ADMIN_USER"),"Fed-CURLEYA") then
'	pv_DEBUG = 0
'else 
'	pv_DEBUG = 0
'end if


'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
	if pv_DEBUG = 1 then
		response.write "<br>REF:= " & request.servervariables("HTTP_REFERER")
		response.write "<br>make doc pre"
		Dim Item,i
		For Each Item In Request.Form
			response.write "<br>rf. " & Item &  " = " &  Request.Form(Item)
		Next
		  Response.Write"<br>u. " & Request.QueryString
		response.write "<br>pv_newdb = " & pv_newdb & " for org = " & session("template_org_code")  
		response.write "<br>pv_htmldoc= " & pv_htmldoc 
		response.write "<br>multicode = " & multicode
		'response.End()
	end if
	
'*****************************************************************************************************
' TO MAKE SINGLE PDF OF APPLICANT PUBLIC SITE - PDF OR WORD
'16 DEC 20 LJL this is where word also is generated.  The warnings are in English and don't show for WORD as it is immediately processed in
'public-view-f-test.asp
'*****************************************************************************************************
'19 feb 07 ac - added new temp /public_hold when fixed changed back /public 
if instr(request.servervariables("HTTP_REFERER"), "/public/pdf/docsettings.asp") then
	if isnumeric(session("RSYS_EVAL")) then
		pv_type = "0"
		'response.write "Your CV in " & Ucase(pv_doc_type) & " is being created.<br><br>Please wait.<br><Br>Your PDF will be sent to your local download folder or area on your computer."
		response.write "Your " & Ucase(pv_doc_type) & " CV is being created.<br><br>Please wait.<br><Br>The resulting document will be sent to your local download folder/area on your computer.<br><br>"
		'ac blocked  send_ids = "&ids=" & session("RSYS_EVAL")
		'ids = session("RSYS_EVAL")
'20 DEC 06 LJL added to be able to do from admin site
		if session("RSYS_EVAL") <> "" then
			applicant_id = session("RSYS_EVAL")
		else
			applicant_id = form("adminjap")
		end if
		' TO SEND ACROSS THE JOB ID TO REDUCE COVERING LETTERS IN APPLICANT CV OUTPUT
		if request.form("fm_jobid") <> "" then
			'ac blocked  send_ids = send_ids & "&fm_jobid=" & request.form("fm_jobid")
			fm_jobid = request.form("fm_jobid")
		end if
		'07 Nov 06 ac- added the ability to preapre for pic in pdf 
		if request.querystring("usepix") = "1" OR request.form("usepix") = "1" then
			pv_usepix = "1"
		end if
		doc_page = "ACdoc-maker.asp"
		LA = 1
		'dest_page = "public\pdf\"
		result_page = "CV_" & year(now()) & month(now()) & day(now()) & minute(now()) & second(now()) & int(session("RSYS_EVAL")) * second(now()) & "_" & session("lng")
		'Add by gevorg 
		PdfFileName = "{0}"
		if request.form("usetoc") = "1" then
			pv_tocset = " --toctitle CV --toclevels 3 --tocheader DDD "
		else
			pv_tocset = " --no-toc "
		end if
		pv_htmldoc = "--webpage --size " & request.form("format") & pv_tocset & " --book --format pdf --portrait --pagelayout twoleft --bodyfont Helvetica --fontsize 10 --headingfont Arial --header .c. --footer d./ --headfootfont Helvetica --headfootsize 9 --no-title --pagemode document --top 0 --bottom 0 --right 1cm --left 1cm"
		'ac blcoked - pv_parts = "&parts=" & replace(request.form("app")," ","")
		'19 DEC 06 LJL had to add ,XXX to make the section come up if there is just one, such as just A section
		pv_parts = replace(request.form("app")," ","") & ",XXXX"
	else
		response.write "You must enter this page correctly."
		response.end
	end if
	
'*****************************************************************************************************
' ac yes, admin - multiple appl pdfs
'*****************************************************************************************************
elseif instr(request.servervariables("HTTP_REFERER"), "/admin/docsettings.asp") AND request.form("pv_newcode") <> "" then
	response.write "The ADMIN PDF file is being created.<br><br>Please wait until the link appears below (msg:2)"
	pv_multi = "1" ' to produce page breaks
	multi = "1" ' set to produce multiple cvs
	pv_type = "1" ' set to provide link 0 opens file when done
	doc_page = "ACdoc-maker.asp"
	'dest_page = "admin\pdf\"
	result_page = "MULTI_" & request.form("pv_newcode")& "_" & session("lng")
	'07 Nov 06 ac- added the ability to preapre for pic in pdf 
	if request.querystring("usepix") = "1" OR request.form("usepix") = "1" then
		pv_usepix = "1"
	end if
	' TO SEND ACROSS THE JOB ID TO REDUCE COVERING LETTERS IN APPLICANT CV OUTPUT
	' this is only set in - multiple appliancts from search by post
	if request.form("fm_jobid") <> "" then
		pv_jobid = request.form("fm_jobid")
		' TO SEND ACROSS THE JOB ID TO REDUCE COVERING LETTERS IN APPLICANT CV OUTPUT
		ids = request.form("fm_jobid")
	end if
	if request.form("vacchoice") then 
 		vacchoice = request.form("vacchoice")
	end if 
	if request.form("pv_newcode") <> "" then 
		multicode = request.form("pv_newcode")  	
		multi = request.form("pv_newcode")  
	end if 
	if request.form("app") <> "" then
		'19 DEC 06 LJL had to add ,XXX to make the section come up if there is just one, such as just A section
		pv_parts = replace(request.form("app")," ","") & ",XXXX"
		'pv_parts = replace(request.form("app")," ","")
	end if
	if request.form("usetoc") = "1" then
		pv_tocset = " --toctitle CV --toclevels 3 --tocheader DDD "
	else
		pv_tocset = " --no-toc "
	end if
	pv_htmldoc = "--book --size " & request.form("format") & pv_tocset & " --format pdf --portrait --pagelayout twoleft --bodyfont Helvetica --fontsize 9 --headingfont Arial --header .c. --footer d./ --headfootfont Helvetica --headfootsize 8 --pagemode outline --top 0 --bottom 0 --right 1cm --left 1cm"
	

'****************************************************************************************************
' TO MAKE SINGLE PDF OF APPLICANT FROM ADMIN
'*****************************************************************************************************
elseif instr(request.servervariables("HTTP_REFERER"), "/admin/docsettings.asp") AND (request.form("adminjap") <> "")  then
	response.write "The ADMIN PDF file is being created.<br><br>Please wait (3)"
    if request.form("adminjap") <> "" then 
		ids = request.form("adminjap") 
	end if
	' TO SEND ACROSS THE JOB ID TO REDUCE COVERING LETTERS IN APPLICANT CV OUTPUT
	if request.form("fm_jobid") <> "" then
		ids = request.form("fm_jobid")
	end if
	doc_page = "ACdoc-maker.asp"
	'dest_page = "admin\pdf\"
	' GET NAME OF FILE BY MULTIPLYING ADMNJAP BY THE SECONDS
	result_page = "ADMIN_CV_" & year(now()) & month(now()) & day(now()) & minute(now()) & second(now()) & request.querystring("adminjap")*second(now())& "_" & session("lng")
	if request.form("usetoc") = "1" then
		pv_tocset = " --toctitle CV --toclevels 3 --tocheader DDD "
	else
		pv_tocset = " --no-toc "
	end if	
	if request.form("app") <> "" then
		pv_parts = replace(request.form("app")," ","")
	end if
	pv_htmldoc = "--webpage --size " & request.form("format") & pv_tocset & " --book --format pdf --portrait --pagelayout twoleft --bodyfont Helvetica --fontsize 10 --headingfont Arial --header .c. --footer d./ --headfootfont Helvetica --headfootsize 9 --no-title --pagemode document --top 0 --bottom 0 --right 1cm --left 1cm"

	pv_type = "0"

'*****************************************************************************************************
' INFO -- TO MAKE THE PDF OF THE POST, not a CV
'*****************************************************************************************************
'19 feb 07 ac - added new temp /public_hold when fixed changed back /public 
elseif instr(request.servervariables("HTTP_REFERER"), "/public/hrd-cl-vac-view.asp") then
	response.write "The VACANCY PDF is being created.<br><br>Please wait (4)"
	send_ids = "&ids=" & request.querystring("jobinfo_uid_c")
	doc_page = "make-post-pdf.asp?jobinfo_uid_c=" & request.querystring("jobinfo_uid_c") & "&gopdf=1"
	'''''' FIX dest_page = "public\"
	' GET NAME OF FILE BY MULTIPLYING THE POST UNIQUE NUMBER BY THE SECONDS
	result_page = "VN_" & year(now()) & month(now()) & day(now()) & minute(now()) & second(now()) & request.querystring("jobinfo_uid_c")*second(now())& "_" & session("lng")
	if request.form("usetoc") = "1" then
		pv_tocset = " --toctitle CV --toclevels 3 --tocheader DDD "
	else
		pv_tocset = " --no-toc "
	end if
	pv_htmldoc = pv_tocset & "--webpage --format pdf --portrait --pagelayout twoleft --bodyfont Helvetica --toclevels 3 --fontsize 10 --headingfont Arial --header .c. --footer d./ --headfootfont Helvetica --headfootsize 9 --no-title --pagemode document --top 0 --bottom 0 --right 1cm --left 1cm"
	pv_parts = ""
	pv_type = "0"
'16 dec 06 ac - use same mechanism to create VN Notice for candidate
elseif instr(request.servervariables("HTTP_REFERER"), "ejobs-jobview.asp") then
	response.write "The vacancy PDF is being created.<br><br>Please wait"
	send_ids = request.querystring("jobinfo_uid_c")
	doc_page = "make-post-pdf.asp?org_code=" & session("template_org_code") & "&jobinfo_uid_c=" & send_ids & "&gopdf=1"
	dest_page = "public\"
	result_page = year(now()) & month(now()) & day(now()) & minute(now()) & second(now()) & send_ids*second(now())& "_" & session("lng")
else
	response.write "This page is not allowed."
	response.end
end if

 'response.write "<br>DIR: " & request.servervariables("PATH_TRANSLATED") & "<br>"
'21 Nov 06 ac CHANGE when making LIVE 
'if session("CLI_RSYS_ADMIN_USER") = "Fed-CURLEYA"  or session("CLI_RSYS_ADMIN_USER") = "UNAIDS-LUNDBERGL" or  session("CLI_RSYS_ADMIN_USER") = "Fed-LUNDBERGL" then


'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
if pv_DEBUG = 1 then
	response.write "<br>doc_page = " & doc_page
	''''''response.write "<br>dest_page = " & dest_page
	response.write "<br>pv_doc_type =" & pv_doc_type
	''''response.write "<br>dest_page= " & dest_page
	response.write "<br>dest_dir= " & dest_dir 
	response.write "<br>stop 1"
	'response.End()
end if


'******************************************
'DELETE OLD PDF - more than an hour old
'******************************************
Dim filecount, fso, file, f :rem Set up a variable for counting the number of files
Dim Ext
Set fso = CreateObject("Scripting.FileSystemObject")
' Call the file system object to manipulate files
Set f = fso.GetFolder(dest_dir) :rem use any folder that you want here
' Assign "f" the folder J-ComEDI so that the files within this directory can be manipulated
filecount = 0 :rem Clear the variable before use in our filecount loop

' If there are files in the J-CommEDI folder with today's date on
' them then print the name of the file to the screen and increment
' the counter.
' RESPONSE.write "DIFF: " & dateadd("h", -1, now()) & "TIMENOW: " & now()

On Error Resume Next

	For Each file in f.Files
		Ext =  fso.GetExtensionName(file) 
		if  (file.DateCreated < dateadd("n", -30, now()) and (Ext = "pdf" or Ext = "html" or Ext = "jpeg" or Ext = "jpg"  or Ext = "png" or Ext = "doc" ) ) then
	filecount = filecount + 1
			'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
			if pv_DEBUG = 1 then
				 response.write "<br>The file " & file.Name & "  created on " & FormatDateTime(file.DateCreated,vbShortDate)
				 response.write "<br>The file " & file.Name & "  created on " & file.DateCreated & "EXT: " & fso.GetExtensionName(file) 
			end if
			fso.DeleteFile(file)
			'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
			if pv_DEBUG = 1 then
				 response.write "<br>File deleted= " & file
			end if
		End if 
	next
f.Close ' Make sure you close it or it won't write it!!
Set f = Nothing
Set fso = Nothing 
'******************************************
'DELETE OLD PDF - more than an hour old
'******************************************
' nov 05 06 ac set for test, org on
'pager = " https://erecruit.safehost.net/stage/e-asp/admin/" & doc_page & pv_parts & send_ids & "&org_code=" & session("template_org_code") 

 

'21 nov o6 ac open firl to wrtie html output to
Dim fs, dest_htmlfile 
'30 nov 06 ac check if it is PDF or WORD or HTML that is required:pv_doc_type?
dest_htmlfile = dest_dir & result_page & "." & "html"

dest_file = result_page & "." & pv_doc_type
'30 Nov 06 what type of file format does the user want?
if request.form("goPDF" ) <> "" then  
		pv_doc_type = "pdf" 
		dest_finalfile = dest_dir & dest_file
elseif request.form("goWord" ) <> "" then  
		pv_doc_type = "doc" 
		dest_finalfile = dest_dir & dest_file
elseif request.form("goHTML" ) <> "" then  
		pv_doc_type = "html" 
		dest_finalfile = dest_dir & dest_file
else 
		pv_doc_type = "pdf" 
		dest_finalfile = dest_dir & dest_file
end if


'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
if pv_DEBUG = 1 then
	'response.write "<br>pager= " & pager
	response.write "<br>dest_htmlfile= " & dest_htmlfile
	response.write "<br>stop 2"
	'response.End()
end if

set fs=Server.CreateObject("Scripting.FileSystemObject")
' 8= append, 2= write,1= reading
set f=fs.OpenTextFile(dest_htmlfile,2,true)
'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
if pv_DEBUG = 1 then
	response.write "<br> -doc start-"
		response.write "<br>stop 2.1"
	'response.End()
end if 	
' ***********************************************************
' add write to file 
' ***********************************************************

'Created By Gevorg for test UPU
if false then
	Dim JAPINFO100
	Set JAPINFO100 = Server.CreateObject("ADODB.RecordSet")
	rsys_db_select.CommandTimeout = 320
	JAPINFO100.Open "{call erstp_rsys_cand_app_full_view_en_2400(22984)} ", rsys_db_select, 0, 1
	response.write "<Br>JAPINFO100.RecordCount = "
	response.write JAPINFO100.RecordCount
	
response.end
End if
%>

<!--#include file="../../admin/ACdoc-maker-New.asp"-->
<%

'31 DEC 2015 GG added this To show only specific text 
response.write "<table width=""100%""><tr><td align=""center"" colspan = '2'><br><p style=""font-size:14px;font-family:verdana;color:green;font-weight: bold;"">Finalizing document for output...</p><br><br></td></tr><tr><td align=""center"" colspan = '2'><span style=""display:none;"">"&session("template_org_code")&"</span></td></tr></table>"

' ***********************************************************
' end write to file 
' ***********************************************************
if pv_DEBUG = 1 then
	response.write "<br>stop 3"
	'response.End()
end if 
' file 
set f=Nothing
set fs=Nothing

'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
if pv_DEBUG = 1 then
	response.write "<br>dest_htmlfile= " & dest_htmlfile & " Created"
	response.write "<br>stop 4"
	'response.End()
end if 

'30 Nov 06 ac if pv_doc_type = PDF then send thru' generator below
if pv_doc_type = "pdf" then 
	Dim objExecutor 
	Dim sResult 
	'dim pager
	'Create the server-side object 
	''Set objExecutor = Server.CreateObject("ASPExec.Execute") 
	'Set the application name 
	'For NT systems, it's "cmd.exe" 
	'If you're running something else then I think you know 'what it is ) 
	''''' WORKS objExecutor.Application = "c:/htmldoc/ghtmldoc.exe -f E:\who-shared-hosting\Sites\erecruit\htdocs\e-asp\admin\pdfbook\readit.pdf C:\htmldoC:\readme.txt " 
	''page_dest = dest_dir &  result_page & ".pdf " & dest_htmlfile
	'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
	if pv_DEBUG = 1 then
		response.write "<br><br>TEST DEST page_dest= " & page_dest & "<b>"
	end if 
	'ac blocked  

	''pv_fullcode = "c:/htmldoc/ghtmldoc.exe --bodyfont Helvetica --footer / --webpage -f " & page_dest 
	' INFO -- orginal HTMLDOC instructions --verbose --compression=2 --bodyfont Helvetica --fontsize 9 --footer / --webpage
	'22 Nov 06 ac -  - debug flag - could be moved to page header for all files?
	if pv_DEBUG = 1 then
		response.write "<br>stop 5"
		'response.End()
	end if
	
	'16 SEP 09 DD increase the level of compression of generated pdf if photo includes in cv
	''if request.form("usepix") = "1" then
	''	objExecutor.Application = "c:/htmldoc/ghtmldoc.exe " & pv_htmldoc & " --compression=9 -f " & page_dest 
	''else	
	''	objExecutor.Application = "c:/htmldoc/ghtmldoc.exe " & pv_htmldoc & " --compression=3 -f " & page_dest 
	''End if	
	' response.write "<BR>c:/htmldoc/ghtmldoc.exe " & pv_htmldoc & " --compression=3 -f " & page_dest  & "<BR>"
	'Now set the parameters, very important! 
	'You want your DOS prompt to call your batch file 
	'For NT - it's "/c filename.exe" 
	''objExecutor.Parameters = ""
	''objExecutor.ShowWindow = False 
	
	' ADD TIMEOUT
	''objExecutor.TimeOut = "36000"
	
	'Here we execute the app and get the output to this string 
	''sResult = objExecutor.ExecuteWinApp
	'Response.Write "<br>Result " & sResult & "<p>"
	
	'  --no-toc --format pdf --portrait --pagelayout one --bodyfont Helvetica --fontsize 9 --headingfont Arial --header .c. --footer d./ --headfootfont Helvetica --headfootsize 9 --no-title --toclevels 1 --tocheader ...  --pagemode document --top 0 --bottom 0 --right 1cm --left 1cm

	
	
	
	
	'Dim theDoc , theID
	'int pageCount=1 
	'int i=1




	'Set theDoc = Server.CreateObject("ABCpdf8.Doc")
	'theDoc.Rect.Inset 20, 20

	'Response.Write "Checking page :" &  dest_htmlfile
	'theDoc.Page = theDoc.AddPage()
	'theID = theDoc.AddImageUrl("file://"  & dest_htmlfile)

	'Do
	'  theDoc.FrameRect ' add a black border
	'  If Not theDoc.Chainable(theID) Then Exit Do
	'  theDoc.Page = theDoc.AddPage()
	'  theID = theDoc.AddImageToChain(theID)
	'Loop


	'  For i = 1 To theDoc.PageCount
	'  theDoc.PageNumber = i
	'  theDoc.Flatten
	'Next
		 
	'theDoc.Save   dest_dir &  result_page & ".pdf "
	Dim myArray
	Dim htmlFileName  
	myArray = Split(dest_htmlfile,"\")
'31 DEC 15 GG Fixed file path
	if Instr(1, dir_current, "STAGE", 1) > 0 OR Instr(1, dir_current, "DEMO", 1) > 0 then
		htmlFileName = myArray(6)
	else
		htmlFileName = myArray(5)
	end if
	
	'09 NOV 2016 GG Commented for sequrity reason
	htmlFileName = replace(htmlFileName, ".html", "")
	
	ApplicantName = replace(ApplicantName,"\","_")
	ApplicantName = replace(ApplicantName,"/","_")
	
	'PdfFileName = replace(PdfFileName,"{0}","CV_" & ApplicantName) & Mid(htmlFileName, 3, 18) & ".pdf"
	
	orgColor=Right(orgColor, Len(orgColor) - 1)
	'response.write"TestTable.aspx?htmlfile="&myArray(4)&"clr="&orgColor&"Vnname="&Vnname&"UseToc="&request.form("usetoc")&"ASize="&request.form("format")
	 
	%>
<script>
  window.location="../../GeneratePDF.aspx?OutputType=PublicPdf&VnClosing=&htmlfile=<%=htmlFileName%>&clr=<%=orgColor%>&Org=<%=session("template_org_name")%>&Vnname=<%=Vnname%>&UseToc=<%=request.form("usetoc")%>&ASize=<%=request.form("format")%>&Lng=<%=session("lng")%>&ViewFormat=<%=request.form("viewformat")%>&Usepix=<%=request.form("usepix")%>&TextSize=<%=request.form("textsize")%>&InfoPage=<%=request.form("infopage")%>&ApplicantName=<%=ApplicantName%>";
</script>
<%
Response.end  


end if '30 nov 06 ac - end of create PDF from html generated file

 if pv_DEBUG = 1 then
		response.write "<br>dest_htmlfile = " & dest_htmlfile
		response.write "<br>dest_finalfile = " & dest_finalfile
		'response.End()
end if
	
if pv_doc_type = "doc" then 'rename file ext of html created file to doc
	'******************************************
	'rename file from ?.html to ?.doc
	'******************************************
	
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fso.MoveFile dest_htmlfile, dest_finalfile 
	Set fso = Nothing
end if

if pv_type = "0" then
' SET THIS PART TO BE REM'd if you want to test output, then put the URL that gets outputted into the browser on the erecruit SERVER.
response.redirect ("public-view-f.asp?name=" &dest_file)

'response.write "<br><br><font color=maroon><Strong>The PDF maker is being revised.  Please try later.</font></strong><br><br>"
'response.redirect ("public-view-f-test.asp?name=" &dest_file)
else%>

<br><br>
<div align="center"><a href='public-view-f.asp?name=<%=dest_file%>'>Please click here to retrieve your file</a> <%'=pv_htmldoc%></div>

<%end if

pv_last_update = "16 Dec 20"

' ***********************************************************
' BEGIN INCLUDES
' ***********************************************************%>
<!--#include file="../../includes/include_ext_frame_bottom.asp"-->
<% ' Admin/pdf folder only ----"../includes/include_ext_frame_bottom.asp"--- 

' ***********************************************************
' END INCLUDES
' ***********************************************************%>



