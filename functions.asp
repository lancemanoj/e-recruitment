<%
Function GenerateDropdown(jsonText, dropdownName, selectedValue,OtherTextbox,OtherTextboxValue)
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
    
    ' Trim any trailing comma or spaces
    optionsText = Trim(optionsText)
    If Right(optionsText, 1) = "," Then
        optionsText = Left(optionsText, Len(optionsText) - 1)
    End If
    
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
            <select name="<%= dropdownName %>" id="<%= dropdownName %>" onchange="toggleOther('<%= OtherTextbox %>', '<%= dropdownName %>')">
                <% 
                ' Populate dropdown options
                For i = 0 To UBound(optionsArray)
                    ' Extract key and value
                    optionParts = Split(optionsArray(i), ":")
                    If UBound(optionParts) = 1 Then
                        optionValue = Trim(optionParts(0))
                        optionText = Trim(optionParts(1))
                        
                        ' Clean optionValue to avoid extra spaces or commas
                        optionValue = Replace(optionValue, " ", "")
                        optionValue = Replace(optionValue, ",", "")
                        
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
            <span id='<%= OtherTextbox %>' style="display:none;">
                <input type="text" placeholder="specify other" name='<%= OtherTextbox %>' value='<%= OtherTextboxValue %>'>
            </span>
        </td>
    </tr>

    <% 
End Function

Function GenerateCheckboxGroup(jsonText, checkboxName, selectedValues, OtherTextbox, OtherTextboxValue)
    Dim jsonKey, optionsStart, optionsEnd, optionsText
    Dim i, optionParts
    Dim optionValue, optionText
    Dim keyStart, keyEnd ' Ensure these variables are declared

    ' Extract the key for the label
    keyStart = InStr(jsonText, "{") + 1
    keyEnd = InStr(keyStart, jsonText, ":{") - 1
    If keyEnd > keyStart Then
        jsonKey = Mid(jsonText, keyStart, keyEnd - keyStart + 1)
    Else
        jsonKey = ""
    End If
    
    ' Remove any extra quotation marks
    jsonKey = Replace(jsonKey, """", "")
    Response.Write "<td valign='top'><label>" & jsonKey & "</label></td>"
        Response.Write "<td valign='top'>"
    ' Extract the options for the checkboxes
    optionsStart = InStr(jsonText, ":{") + 2
    optionsEnd = InStrRev(jsonText, "}") - 1
    If optionsEnd > optionsStart Then
        optionsText = Mid(jsonText, optionsStart, optionsEnd - optionsStart + 1)
    Else
        optionsText = ""
    End If

    ' Remove extra characters like curly braces and quotes
    optionsText = Replace(optionsText, """", "")
    optionsText = Replace(optionsText, "{", "")
    optionsText = Replace(optionsText, "}", "")

    ' Trim any trailing comma or spaces
    optionsText = Trim(optionsText)
    If Right(optionsText, 1) = "," Then
        optionsText = Left(optionsText, Len(optionsText) - 1)
    End If

    ' Convert optionsText to an array of key-value pairs
    Dim optionsArray
    optionsArray = Split(optionsText, ",")

    Response.Write "<fieldset>"
    
    For i = 0 To UBound(optionsArray)
        optionParts = Split(optionsArray(i), ":")
        If UBound(optionParts) = 1 Then
            optionValue = Trim(optionParts(0))
            optionText = Trim(optionParts(1))
            
            ' Check if this option is selected
            Dim checked
            checked = ""
            If InStr(selectedValues, optionValue) > 0 Then
                checked = "checked"
            End If
            
            ' Write checkbox with onchange event
            Response.Write "<label>"
            Response.Write "<input type='checkbox' name='" & checkboxName & "' value='" & optionValue & "' " & checked & " onchange='toggleOtherTextbox()'> " & optionText
            Response.Write "</label><br>"
        End If
    Next
    Response.Write "</fieldset>"

    ' Other textbox (initially hidden)
    Response.Write "<span id='" & OtherTextbox & "' style='display:none;'>"
    Response.Write "<input type='text' placeholder='Specify other " & jsonKey & "' name='" & OtherTextbox & "' value='" & OtherTextboxValue & "'>"
    Response.Write "</span>"

    ' JavaScript function to toggle the visibility of the "Other" textbox
    Response.Write "<script>"
    Response.Write "function toggleOtherTextbox() {"
    Response.Write "  var checkboxes = document.getElementsByName('" & checkboxName & "');"
    Response.Write "  var otherTextbox = document.getElementById('" & OtherTextbox & "');"
    Response.Write "  var showOther = false;"
    Response.Write "  for (var i = 0; i < checkboxes.length; i++) {"
    Response.Write "    if (checkboxes[i].value == '-1' && checkboxes[i].checked) {"
    Response.Write "      showOther = true;"
    Response.Write "      break;"
    Response.Write "    }"
    Response.Write "  }"
    Response.Write "  otherTextbox.style.display = showOther ? 'block' : 'none';"
    Response.Write "}"
    Response.Write "</script>"
End Function

        Function UpdateCandidate(rsys_db, pv_candidD, race_ethnicity_value)
    Dim dbCmd, updateSQL, fullQuery

    ' Update SQL query
    updateSQL = "UPDATE dbo.tx_rsys_candmisc SET canddiv_pronoun_id = ?, cand_div_race_ethnicity = ?, cand_div_disability = ?, cand_div_disability_accommodation = ?, cand_div_key_population = ?, cand_other_pronoun = ?, cand_gnd_i = ?, cand_other_gender = ?, cand_div_genderidentity = ?, cand_div_other_genderidentity = ?, cand_div_other_raceethnicity = ?, cand_div_disability_accom = ?, cand_div_keyPopulationsExplanation = ? WHERE cand_id_c = ?"
    
    ' Build query string for debugging
    fullQuery = "UPDATE dbo.tx_rsys_candmisc SET " & _
                "canddiv_pronoun_id = '" & Request.Form("canddiv_pronoun_id") & "', " & _
                "cand_div_race_ethnicity = '" & race_ethnicity_value & "', " & _
                "cand_div_disability = '" & Request.Form("cand_div_disability") & "', " & _
                "cand_div_disability_accommodation = '" & Request.Form("cand_div_disability_accommodation") & "', " & _
                "cand_div_key_population = '" & Request.Form("cand_div_key_population") & "', " & _
                "cand_other_pronoun = '" & Request.Form("cand_other_pronoun") & "', " & _
                "cand_gnd_i = '" & Request.Form("cand_gnd_i") & "', " & _
                "cand_other_gender = '" & Request.Form("cand_other_gender") & "', " & _
                "cand_div_genderidentity = '" & CleanInput(Request.Form("cand_div_genderidentity")) & "', " & _
                "cand_div_other_genderidentity = '" & Request.Form("OtherIdentityTextbox") & "', " & _
                "cand_div_other_raceethnicity = '" & Request.Form("OtherraceethnicityTextbox") & "', " & _
                "cand_div_disability_accom = '" & Request.Form("cand_div_disability_accom") & "', " & _
                "cand_div_keyPopulationsExplanation = '" & Request.Form("cand_div_keyPopulationsExplanation") & "' " & _
                "WHERE cand_id_c = '" & pv_candidD & "'"

    ' Print the query for debugging
    Response.Write("SQL Query: " & fullQuery)
   'Response.End

    ' Execute the update query
    Set dbCmd = Server.CreateObject("ADODB.Command")
    dbCmd.ActiveConnection = rsys_db
    dbCmd.CommandText = updateSQL

    ' Append parameters for update query
    dbCmd.Parameters.Append dbCmd.CreateParameter("@pronoun", adVarChar, adParamInput, 50, Request.Form("canddiv_pronoun_id"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@race_ethnicity", adVarChar, adParamInput, 255, race_ethnicity_value)
    dbCmd.Parameters.Append dbCmd.CreateParameter("@disability", adVarChar, adParamInput, 50, Request.Form("cand_div_disability"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@accommodation", adVarChar, adParamInput, 50, Request.Form("cand_div_disability_accommodation"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@key_population", adVarChar, adParamInput, 255, Request.Form("cand_div_key_population"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_other_pronoun", adVarChar, adParamInput, 255, Request.Form("cand_other_pronoun"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_gnd_i", adVarChar, adParamInput, 255, Request.Form("cand_gnd_i"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_other_gender", adVarChar, adParamInput, 255, Request.Form("cand_other_gender"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_div_genderidentity", adVarChar, adParamInput, 255, CleanInput(Request.Form("cand_div_genderidentity")))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@OtherIdentityTextbox", adVarChar, adParamInput, 255, Request.Form("OtherIdentityTextbox"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@OtherraceethnicityTextbox", adVarChar, adParamInput, 255, Request.Form("OtherraceethnicityTextbox"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_div_disability_accom", adVarChar, adParamInput, 255, Request.Form("cand_div_disability_accom"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_div_keyPopulationsExplanation", adVarChar, adParamInput, 255, Request.Form("cand_div_keyPopulationsExplanation"))
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_id", adInteger, adParamInput, , pv_candidD)

    dbCmd.Execute()
End Function

  Function InsertOrUpdateCandidate(rsys_db, pv_candidD, race_ethnicity_value)
    Dim dbCmd, checkSQL, insertSQL, updateSQL, rs, fullQuery

    ' Check if the record exists
    checkSQL = "SELECT COUNT(*) FROM dbo.tx_rsys_candmisc WHERE cand_id_c = ?"
    Set dbCmd = Server.CreateObject("ADODB.Command")
    dbCmd.ActiveConnection = rsys_db
    dbCmd.CommandText = checkSQL
    dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_id", adInteger, adParamInput, , pv_candidD)

    Set rs = dbCmd.Execute()

    If Not rs.EOF Then
        If rs(0) = 0 Then
            ' Record does not exist, so insert
            insertSQL = "INSERT INTO dbo.tx_rsys_candmisc (canddiv_pronoun_id, cand_div_race_ethnicity, cand_div_disability, cand_div_disability_accommodation, cand_div_key_population, cand_other_pronoun, cand_gnd_i, cand_other_gender, cand_div_genderidentity, cand_div_other_genderidentity, cand_div_other_raceethnicity, cand_div_disability_accom, cand_div_keyPopulationsExplanation, cand_id_c) " & _
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            fullQuery = "INSERT INTO dbo.tx_rsys_candmisc (canddiv_pronoun_id, cand_div_race_ethnicity, cand_div_disability, cand_div_disability_accommodation, cand_div_key_population, cand_other_pronoun, cand_gnd_i, cand_other_gender, cand_div_genderidentity, cand_div_other_genderidentity, cand_div_other_raceethnicity, cand_div_disability_accom, cand_div_keyPopulationsExplanation, cand_id_c) " & _
                        "VALUES ('" & Request.Form("canddiv_pronoun_id") & "', '" & race_ethnicity_value & "', '" & Request.Form("cand_div_disability") & "', '" & Request.Form("cand_div_disability_accommodation") & "', '" & Request.Form("cand_div_key_population") & "', '" & Request.Form("cand_other_pronoun") & "', '" & Request.Form("cand_gnd_i") & "', '" & Request.Form("cand_other_gender") & "', '" & CleanInput(Request.Form("cand_div_genderidentity")) & "', '" & Request.Form("OtherIdentityTextbox") & "', '" & Request.Form("OtherraceethnicityTextbox") & "', '" & Request.Form("cand_div_disability_accom") & "', '" & Request.Form("cand_div_keyPopulationsExplanation") & "', '" & pv_candidD & "')"
            
            ' Print the query for debugging
            Response.Write("SQL Query: " & fullQuery)
            Response.End

            Set dbCmd = Server.CreateObject("ADODB.Command")
            dbCmd.ActiveConnection = rsys_db
            dbCmd.CommandText = insertSQL

            ' Append parameters for insert query
            dbCmd.Parameters.Append dbCmd.CreateParameter("@pronoun", adVarChar, adParamInput, 50, Request.Form("canddiv_pronoun_id"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@race_ethnicity", adVarChar, adParamInput, 255, race_ethnicity_value)
            dbCmd.Parameters.Append dbCmd.CreateParameter("@disability", adVarChar, adParamInput, 50, Request.Form("cand_div_disability"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@accommodation", adVarChar, adParamInput, 50, Request.Form("cand_div_disability_accommodation"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@key_population", adVarChar, adParamInput, 255, Request.Form("cand_div_key_population"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_other_pronoun", adVarChar, adParamInput, 255, Request.Form("cand_other_pronoun"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_gnd_i", adVarChar, adParamInput, 255, Request.Form("cand_gnd_i"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_other_gender", adVarChar, adParamInput, 255, Request.Form("cand_other_gender"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_div_genderidentity", adVarChar, adParamInput, 255, CleanInput(Request.Form("cand_div_genderidentity")))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@OtherIdentityTextbox", adVarChar, adParamInput, 255, Request.Form("OtherIdentityTextbox"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@OtherraceethnicityTextbox", adVarChar, adParamInput, 255, Request.Form("OtherraceethnicityTextbox"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_div_disability_accom", adVarChar, adParamInput, 255, Request.Form("cand_div_disability_accom"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_div_keyPopulationsExplanation", adVarChar, adParamInput, 255, Request.Form("cand_div_keyPopulationsExplanation"))
            dbCmd.Parameters.Append dbCmd.CreateParameter("@cand_id", adInteger, adParamInput, , pv_candidD)

            dbCmd.Execute()

        Else
            ' Existing record, proceed with update
            Call UpdateCandidate(rsys_db, pv_candidD, race_ethnicity_value)
        End If
    End If
End Function

%>