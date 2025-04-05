<html>
<head>
    <title>Recruitment Report Export Form</title>

        <%
' Define page title as in rsys-check-visitor.asp

pv_page_title = "SEARCH FOR APPLICANTS IN DATABASE"
%>
    <!--#include file = "includes/include_admin_frame_top.asp"-->

 

    <style type="text/css">
        /* Inferred CSS styles from rsys-check-visitor.asp */
        .checkVistor {
            margin: 20px;
            font-family: Arial, Helvetica, sans-serif;
        }
        h2 {
            color: navy;
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 10px;
        }
        table {
            border-collapse: collapse;
            width: 90%;
            margin: 0 auto;
        }
        td {
            padding: 8px;
            vertical-align: middle;
        }
        label {
            font-size: 12px;
            margin-right: 5px;
        }
        .textsmall {
            font-size: 12px;
            padding: 4px;
            border: 1px solid #ccc;
            width: 80px; /* Matches typical size from rsys-check-visitor.asp */
        }
        .textsmall[type="submit"] {
            background-color: #f0f0f0;
            border: 1px solid #999;
            padding: 6px 12px;
            cursor: pointer;
        }
        .textsmall[type="submit"]:hover {
            background-color: #e0e0e0;
        }
        .alert {
            color: red;
            font-weight: bold;
        }
    </style>
</head>
<body>
   <div class="checkVistor">
  
        <form method="post" action="export.asp">
            <table width="90%" align="center" border="0" bordercolor="orange">
                <tr>
                    <td valign="middle">
                        <label>Closing Year From:</label>
                        <input type="number" name="fromYear" value="<%=Year(Date)-1%>" class="textsmall">
                    </td>
                    <td valign="middle">
                        <label>To:</label>
                        <input type="number" name="toYear" value="<%=Year(Date)%>" class="textsmall">
                    </td>
                    <td valign="middle" align="center">
                        <button type="submit" class="textsmall">Export</button>
                    </td>
                </tr>
            </table>
        </form>
    </div>
    <!--#include file = "includes/include_admin_frame_bottom.asp"-->
</body>
</html>