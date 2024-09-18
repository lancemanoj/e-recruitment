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
%>