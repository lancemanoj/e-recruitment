<%
' Set response headers for CSV download
Response.ContentType = "application/octet-stream"
Response.AddHeader "Content-Disposition", "attachment; filename=recruit_report.csv"

' Get year range from form input (default to current year)
Dim fromYear, toYear
fromYear = Request.Form("fromYear")
toYear = Request.Form("toYear")
If fromYear = "" Then fromYear = Year(Date) ' Default to current year
If toYear = "" Then toYear = Year(Date)

' Validate input
fromYear = CInt(fromYear)
toYear = CInt(toYear)

' Build SQL query with year range filter
Dim sql
sql = "SELECT  j.jobinfo_vac2_c AS VN_Number, j.jobinfo_job_en_t AS VN_Title, " & _
      "CONVERT(VARCHAR(11), j.jobinfo_acl_d, 106) as VN_closing_date, " & _
      "c.cand_lnam_t as Cand_last_name, ISNULL(UPPER(c.cand_onam_t), '') as Cand_last_pref_name, " & _
      "c.cand_fnam_t as Cand_first_name, ISNULL(UPPER(c.cand_ofnam_t), '') as Cand_first_pref_name, " & _
      "c.cand_email_t AS Cand_email, c.cand_nat_c AS Cand_nat_1st, " & _
      "CASE WHEN cj.candjob_screened_i = '0' THEN '' WHEN cj.candjob_screened_i = '1' THEN 'Screened' END AS Screened, " & _
      "CASE WHEN cj.candjob_shortlist_i = '0' THEN '' WHEN cj.candjob_shortlist_i = '1' THEN 'Shortlisted' END AS Shortlisted, " & _
      "CASE WHEN cj.candjob_tested_i = '0' THEN '' WHEN cj.candjob_tested_i < 4 THEN 'Test given-incomplete' WHEN cj.candjob_tested_i = 4 THEN 'Test taken' END AS Tested, " & _
      "CASE WHEN cj.candjob_interv_i = '0' THEN '' WHEN cj.candjob_interv_i = '1' THEN 'Interviewed' END AS Interviewed, " & _
      "CASE WHEN cj.candjob_recomm_i = '0' THEN '' WHEN cj.candjob_recomm_i = '1' THEN 'Recommended' END AS Recommended, " & _
      "CASE WHEN cj.candjob_selected_i = '0' THEN '' WHEN cj.candjob_selected_i = '1' THEN 'Selected' END AS Selected, " & _
      "CONVERT(VARCHAR(11), e.editDiv_d, 106) as Div_added_date, e.editDiv AS Div_edited_howoften, " & _
      "CONVERT(VARCHAR(11), c.upd_d, 106) AS Profile_last_upd, CONVERT(VARCHAR(11), c.crd_d, 106) AS Profile_created_date, " & _
      "CASE WHEN c.cand_gnd_i = '0' THEN 'F' WHEN c.cand_gnd_i = '1' THEN 'M' ELSE 'Other' END AS Cand_gender, " & _
      "m.canddiv_pronoun_id AS Div_pronouns, m.cand_other_pronoun AS Div_pronouns_other, " & _
      "CASE WHEN m.cand_div_genderidentity = '-1' THEN 'Other' WHEN m.cand_div_genderidentity = '0' THEN '' ELSE m.cand_div_genderidentity END AS Div_gender_identity, " & _
      "m.cand_div_other_genderidentity AS Div_gender_identity_other, REPLACE(m.cand_div_race_ethnicity, '-1', 'Other') AS Div_race_ethnicity, " & _
      "m.cand_div_other_raceethnicity AS Div_race_ethnicity_other, ISNULL(m.cand_div_keyPopulationsExplanation, '') AS Div_key_populations_info, " & _
      "CASE WHEN m.cand_div_disability = '0' THEN '' WHEN m.cand_div_disability = '1' THEN 'Yes' ELSE m.cand_div_disability END AS Div_disability_YN, " & _
      "CASE WHEN m.cand_div_disability_accommodation = '1' THEN 'Yes' ELSE '' END AS Div_disability_accom_YN, " & _
      "ISNULL(m.cand_div_disability_accom, '') AS Div_disabilty_accom_comments, c.cand_id_c as Candidate_id " & _
      "FROM tx_rsys_candmisc m " & _
      "INNER JOIN td_rsys_cand c ON m.cand_id_c = c.cand_id_c " & _
      "INNER JOIN tx_rsys_candedit e ON e.cand_id_c = c.cand_id_c " & _
      "INNER JOIN tx_rsys_candjob cj ON cj.cand_id_c = c.cand_id_c " & _
      "INNER JOIN td_rsys_jobinfo j ON j.jobinfo_uid_c = cj.jobinfo_uid_c " & _
      "WHERE  YEAR(j.jobinfo_acl_d) BETWEEN " & fromYear & " AND " & toYear & " " & _
      "ORDER BY 1 DESC "

' Establish SQL Server connection
Set conn = Server.CreateObject("ADODB.Connection")
' Add your connection string here: conn.Open "your_connection_string"
conn.Open "Provider=MSOLEDBSQL;Data Source=192.168.106.4;Initial Catalog=erec_test;User Id=write_human;Password=act1ve44"
' Execute query
Set rs = conn.Execute(sql)

' Write complete CSV header matching all selected columns
Response.Write "VN_Number,VN_Title,VN_closing_date,Cand_last_name,Cand_last_pref_name,Cand_first_name,Cand_first_pref_name," & _
               "Cand_email,Cand_nat_1st,Screened,Shortlisted,Tested,Interviewed,Recommended,Selected,Div_added_date," & _
               "Div_edited_howoften,Profile_last_upd,Profile_created_date,Cand_gender,Div_pronouns,Div_pronouns_other," & _
               "Div_gender_identity,Div_gender_identity_other,Div_race_ethnicity,Div_race_ethnicity_other," & _
               "Div_key_populations_info,Div_disability_YN,Div_disability_accom_YN,Div_disabilty_accom_comments,Candidate_id" & vbCrLf

' Write all data rows matching headers 
Do While Not rs.EOF
    Response.Write """" & Replace(rs("VN_Number") & "", """", """""") & """," & _
                   """" & Replace(rs("VN_Title") & "", """", """""") & """," & _
                   """" & Replace(rs("VN_closing_date") & "", """", """""") & """," & _
                   """" & Replace(rs("Cand_last_name") & "", """", """""") & """," & _
                   """" & Replace(rs("Cand_last_pref_name") & "", """", """""") & """," & _
                   """" & Replace(rs("Cand_first_name") & "", """", """""") & """," & _
                   """" & Replace(rs("Cand_first_pref_name") & "", """", """""") & """," & _
                   """" & Replace(rs("Cand_email") & "", """", """""") & """," & _
                   """" & Replace(rs("Cand_nat_1st") & "", """", """""") & """," & _
                   """" & Replace(rs("Screened") & "", """", """""") & """," & _
                   """" & Replace(rs("Shortlisted") & "", """", """""") & """," & _
                   """" & Replace(rs("Tested") & "", """", """""") & """," & _
                   """" & Replace(rs("Interviewed") & "", """", """""") & """," & _
                   """" & Replace(rs("Recommended") & "", """", """""") & """," & _
                   """" & Replace(rs("Selected") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_added_date") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_edited_howoften") & "", """", """""") & """," & _
                   """" & Replace(rs("Profile_last_upd") & "", """", """""") & """," & _
                   """" & Replace(rs("Profile_created_date") & "", """", """""") & """," & _
                   """" & Replace(rs("Cand_gender") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_pronouns") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_pronouns_other") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_gender_identity") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_gender_identity_other") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_race_ethnicity") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_race_ethnicity_other") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_key_populations_info") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_disability_YN") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_disability_accom_YN") & "", """", """""") & """," & _
                   """" & Replace(rs("Div_disabilty_accom_comments") & "", """", """""") & """," & _
                   """" & Replace(rs("Candidate_id") & "", """", """""") & """" & vbCrLf
    rs.MoveNext
Loop

' Cleanup
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
Response.End
%>