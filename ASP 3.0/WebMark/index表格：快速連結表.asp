    <table cellpadding="1" cellspacing="5" class="styleTopLink_1">
        <tr><td colspan="2" class="styleHierarchy1"><a name='QuickLink'>快速連結</a></td></tr>
        <tr><td class="styleHierarchy2">快速連結<br>(待分類)</td>
            <td>
                <table border="2" cellpadding="3" cellspacing="5" class="styleTopLink_2">
                    <tr>
<%
    '表格：快速連結表
    RS.MoveFirst
    HierarchyCount = 0
    Do While RS.EOF = False
        RS.Find = "[Top]=1"
        If NOT RS.EOF then
            response.write "<td class='styleHierarchy3'><a href='" & RS("Link") & "'>" & RS("Name") & "</a></td>"

            HierarchyCount = HierarchyCount +1
            IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "</tr>" & vbNewLine & "<tr>"           '當欄位數量到達設定時，轉到下一列
            RS.MoveNext
        End if
    Loop
    
    HierarchyCount = (HierarchyCount) MOD ColumnNumber_Total



    tmpStr = ""
    DO UNTIL HierarchyCount MOD ColumnNumber_Total = 0          '當表格欄位不足時，算出需要補上幾格空欄位。
        tmpStr = tmpStr & "<td>　</td>"
        HierarchyCount = (HierarchyCount +1) MOD ColumnNumber_Total         '統整目前欄位數
    LOOP
    IF NOT tmpStr = nul then response.write tmpStr & "</tr>" & vbNewLine
    response.write "</table></td></tr>" & vbNewLine
    response.write "</table>" & vbNewLine & "</td></tr></table>" & vbNewLine
%>
    <hr>