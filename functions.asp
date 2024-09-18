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
%>