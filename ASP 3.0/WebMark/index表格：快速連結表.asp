    <table cellpadding="1" cellspacing="5" class="styleTopLink_1">
        <tr><td colspan="2" class="styleHierarchy1"><a name='QuickLink'>�ֳt�s��</a></td></tr>
        <tr><td class="styleHierarchy2">�ֳt�s��<br>(�ݤ���)</td>
            <td>
                <table border="2" cellpadding="3" cellspacing="5" class="styleTopLink_2">
                    <tr>
<%
    '���G�ֳt�s����
    RS.MoveFirst
    HierarchyCount = 0
    Do While RS.EOF = False
        RS.Find = "[Top]=1"
        If NOT RS.EOF then
            response.write "<td class='styleHierarchy3'><a href='" & RS("Link") & "'>" & RS("Name") & "</a></td>"

            HierarchyCount = HierarchyCount +1
            IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "</tr>" & vbNewLine & "<tr>"           '�����ƶq��F�]�w�ɡA���U�@�C
            RS.MoveNext
        End if
    Loop
    
    HierarchyCount = (HierarchyCount) MOD ColumnNumber_Total



    tmpStr = ""
    DO UNTIL HierarchyCount MOD ColumnNumber_Total = 0          '������줣���ɡA��X�ݭn�ɤW�X������C
        tmpStr = tmpStr & "<td>�@</td>"
        HierarchyCount = (HierarchyCount +1) MOD ColumnNumber_Total         '�ξ�ثe����
    LOOP
    IF NOT tmpStr = nul then response.write tmpStr & "</tr>" & vbNewLine
    response.write "</table></td></tr>" & vbNewLine
    response.write "</table>" & vbNewLine & "</td></tr></table>" & vbNewLine
%>
    <hr>