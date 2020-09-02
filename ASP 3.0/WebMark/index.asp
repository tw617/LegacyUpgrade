<%
    '�� ============================= �]�w�ŧi ============================= ��
    Dim WebMarkVer
        WebMarkVer = "1.0a"         '�M�μҲճ]�w�� �Ҳժ���
    Dim WebMarkVerDate
        WebMarkVerDate = "2013/10/25"
 
    Dim xlsDataBasePathFile         '�]�w xls �ɮצW��(�i�]�t�۹���|)
        xlsDataBasePathFile = "WebMark.xls"
    Dim hTitle         'Title �W��
        hTitle = Request("htmlTitle")
        If hTitle = Empty Then hTitle = "���ҤU����"
    Dim webLoadingFrequency            '���o�������J������
        webLoadingFrequency = Request("webLoadingFrequency")
        If webLoadingFrequency = null Then
            webLoadingFrequency = 1
        Else
            webLoadingFrequency = webLoadingFrequency + 1
        End If
    
    Dim ColumnNumber_Total          '�]�w��ƪ�����`��
        ColumnNumber_Total = CInt(Request("Textbox_ColumnNumber_Total"))
    Dim IndexTable_Name         '���ު����D�W��
        IndexTable_Name = Request("IndexTable_Name")
    Dim styleHierarchy1           '�������������e��
        styleHierarchy1 = Request("styleHierarchy1")
    Dim tableColumn2Width         '������檺���e��
        If webLoadingFrequency > 1 Then tableColumn2Width = 100 / ColumnNumber_Total
 
    Dim DBConnection
    Dim RS
 
    Dim HierarchyPrevious        '�e�@�Ӷ��ت����h(Hierarchy)
    Dim HierarchyCurrent        '��e���ت����h(Hierarchy)
    Dim HierarchyCount        '�p���e���جO�s��ġH�ӦP���h���A�Ψӭp��������

%>

<html>
<head runat="server">
<title><%=hTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-tw">
    <style type="text/css">
         .styleWebMarkVer
        {
            font-size: small;
            color: #C0C0C0;
        }
        .styleMainTable
        {
            width: 75%;
            border-style: solid;
            border-width: 2px;
        }
        .styleHalfWidth
        {
            width: 50%;
        }
        </style>
</head>

<body>
<p style="text-align: right" class="styleWebMarkVer">�Ҳժ����G<%=WebMarkVer%></p>
<form method="POST" action="index.asp">
    <table align="center" border="3" cellpadding="3" cellspacing="5" class="styleMainTable" style="border: thick double #800000;">
        <caption style="font-family: �رd������W9; font-size: xx-large; color: #FF0000;"><%=hTitle%></caption>
            <tr><td dir="rtl" class="styleHalfWidth">�W��</td>
                <td dir="ltr" class="styleHalfWidth">�]�w��</td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">���� Title �W��</td>
                <td dir="ltr" class="styleHalfWidth"><input type="text" name="htmlTitle" size="20" value="�x�x������" tabindex="1"></td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">���ƶq</td>
                <td dir="ltr" class="styleHalfWidth"><input type="text" name="Textbox_ColumnNumber_Total" size="4" value="4" tabindex="2"></td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">���ު��W��</td>
                <td dir="ltr" class="styleHalfWidth"><input type="text" name="IndexTable_Name" size="20" value="����"></td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">�������������e��</td>
                <td dir="ltr" class="styleHalfWidth"><input type="text" name="styleHierarchy1" size="4" value="10%"></td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">&nbsp;</td>
                <td dir="ltr" class="styleHalfWidth">&nbsp;</td></tr>
        </table>
    <p style="text-align: center"><input type="submit" value="�U��" name="B1"><input type="reset" value="���s�]�w" name="B2"></p>
    <input type="hidden" name="webLoadingFrequency" value="<%=webLoadingFrequency%>">
</form>


<p>&nbsp;</p>

</body>
</html>


<%
IF webLoadingFrequency < 2 then Response.End	'�������J�p�ơA��2������~���ͤU�����e�C


'�� ============================= ���ͤU�� ============================= ��
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=WebMark.htm"
Response.ContentType = "text"
'Response.End


'�� ============================= ��Ƴs�u ============================= ��
ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(xlsDataBasePathFile) & ";Extended Properties=Excel 8.0"
Set DBConnection = Server.CreateObject("ADODB.Connection")
DBConnection.Open ConnectionString
Set RS = Server.CreateObject("ADODB.Recordset")
rem RS.Open "[Content$]",DBConnection,?,?
rem RS.Close
RS.Open "[Content$]",DBConnection,3,1

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title><%=hTitle%></title>
    <base target="_blank">
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <meta http-equiv="Content-Language" content="zh-tw">
    <style type="text/css">
        .styleIndexTable
        /*���ު�*/
        {
            width: 100%;
            border-style: solid;
            border-width: 3px;
        }
        .styleIndexTable_2
        /*���ޤ���*/
        {
            width: 100%;
            border: 2px solid #800000;
        }
        .styleIndexTableTitle
        /*���ު����D�r��*/
        {
            font-family: �رd�֤k��rW7;
            font-size: xx-large; color: #800000;
        }
        .styleIndexTableHierarchy0
        /*���ު����h0�r��*/
        {
            font-family: �رd������W12;
            font-size: x-large; color: #009900;
            width: 100px;
        }
        .styleIndexTableHierarchy1
        /*���ު����h1�r��*/
        {
            font-size: larger;
            width: <%=tableColumn2Width%>%;
        }
        .styleIndexTableHierarchy2
        /*���ޤ������h2�r��*/
        {
            font-size: small;
        }
        .styleTopLink_1
        /*�ֳt�s����*/
        {
            width: 100%;
            border-style: solid;
            border-width: 3px;
        }
        .styleTopLink_2
        /*�ֳt�s������*/
        {
            width: 100%;
            border: 2px solid #800000;
        }
        .styleTableCategory
        /*���O�~��*/
        {
            width: 100%;
            border-style: solid;
            border-width: 3px;
        }
        .styleTableContent
        /*���O����*/
        {
            width: 100%;
            border: 2px solid #800000;
        }
        .styleHierarchy1
        /*���h1 ���D�r��*/
        {
            font-family: �رd�֤k��rW7;
            font-size: xx-large; color: #CC3399;
        }
        .styleHierarchy2
        /*���h2 ���D�r��*/
        /*���Ĥ@��e��*/
        {
            font-family: �رd������W12;
            font-size: x-large; color: #FF6600;
            width: 135px;
        }
        .styleHierarchy3
        /*���h3 �W�٦r��*/
        /*�����������e��*/
        {
            font-size: larger; color: #800000;
            width: <%=tableColumn2Width%>%;
        }
        .styleHierarchy4
        /*���h4 ��r�r��*/
        {
            font-size: small; color: #000000;
        }
    </style>
</head>
<body>
<div align="right"><font size="2" color="#C0C0C0">�����G<%=now%></font></div>

<%
    '���G���ު�
    RS.MoveFirst
    
    response.write "<table border='1' cellspacing='5' class='styleIndexTable'>" & vbNewLine
    response.write "    <tr><td colspan='2' class='styleIndexTableTitle'>" & IndexTable_Name & "</td></tr>" & vbNewLine

    Do While NOT RS.EOF
        If vartype(HierarchyCurrent) = 0 Then HierarchyCurrent = RS("Hierarchy")
        
        '��ƪ��e�q HTML �N�X
        Select Case HierarchyCurrent
            Case 0
                '�y�u���h0�v�A�����j���A�ȯ��ު�ϥΡC�z
                '�Ģ����h�A�ȯ��ު�ϥΡC
                response.write "    <tr>" & vbNewLine
                response.write "        <td class='styleIndexTableHierarchy0'>" & RS("Name") & "</td>" & vbNewLine
                response.write "        <td>" & vbNewLine
                response.write "            <table border='2' cellpadding='3' cellspacing='5' class='styleIndexTable_2'>" & vbNewLine
                HierarchyCount = 0
                
            Case 1
                '�y�u���h(���)�}�l�v��u���h1�v�����C�z
                IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "                <tr>" & vbNewLine           '�����ƶq��F�]�w�ɡA�_�l�@�C�C
                HierarchyCount = HierarchyCount +1
                response.write "                    <td class='styleIndexTableHierarchy1'><a target='_self' href='#" & RS("Link") & "'>" & RS("Name") & "</a>"
                
                
            Case 2
                '�y�u���h�}�l�v��u���h2�v�����C�z
                response.write RS("Name") & ", "
                
                
            Case 3
                '�y�u���h�}�l�v��u���h3�v�����C�z
                
                
            Case 4
                '�y�u���h�}�l�v��u���h4�v�����C�z
                
                
            Case ELSE
                response.write "<p>�y���~�GHierarchyPrevious < HierarchyCurrent�AIn���N�~�C�z<br>HierarchyPrevious =�y" & HierarchyPrevious & "�z�F<br>HierarchyCurrent =�y" & HierarchyCurrent & "�z�C</p>"
                response.END
            END Select
        
        
        
        '���Щ��U���ʪ�����
        HierarchyPrevious = HierarchyCurrent
        
        Do
            RS.MoveNext()
            IF NOT RS.EOF Then
                HierarchyCurrent = RS("Hierarchy")
            ELSE
                HierarchyCurrent = 0
                Exit Do
            END IF
            IF HierarchyCurrent <= 2 Then Exit Do        '���ު�檺���X�Ȭ�2�F���e��檺�Ȭ��̤j��4�C
        Loop
        
        
        
        '��ƪ���q HTML �N�X
        Select Case True
            Case HierarchyPrevious < HierarchyCurrent
                '�y�u�|�v�B�W�[�@�Ӷ��h�z
                Select Case HierarchyPrevious
                    Case 0
                        '�y�u���h0�v�A�����j���A�ȯ��ު�ϥΡC�z
                        
                    Case 1
                        '�y�u���h1�v����A���u���h2�v���u���h�}�l�v�C�z
                        response.write "<br>" & vbNewLine
                        response.write "                        <span class='styleIndexTableHierarchy2'>"
                        
                    Case 2
                        '�y�u���h2�v����A���u���h3�v���u���h�}�l�v�C�z
                        
                        
                    Case 3
                        '�y�u���h3�v����A���u���h4�v���u���h�}�l�v�C�z   

                        
                    Case 4
                        '�y�u���h4�v����A���u���h5�v���u���h�}�l�v�C�z
                        
                        
                    Case ELSE
                        response.write "<p>�y���~�G��q<br>HierarchyPrevious < HierarchyCurrent ���N�~�C�z<br>HierarchyPrevious =�y" & HierarchyPrevious & "�z�F<br>HierarchyCurrent =�y" & HierarchyCurrent & "�z�C</p>"
                        response.END
                        
                    END Select
                
                
            Case HierarchyPrevious >= HierarchyCurrent
                '�i��|�@�����n�X�Ӷ��h
                tmpHierarchy = HierarchyPrevious
                Do
                    Select Case tmpHierarchy
                        Case 0
                            '�y�u���h0�v�A�����j���A�ȯ��ު�ϥΡC�z
                            IF NOT HierarchyCount MOD ColumnNumber_Total = 0 then
                                tmpStr = "                    "
                                DO UNTIL HierarchyCount MOD ColumnNumber_Total = 0          '������줣���ɡA��X�ݭn�ɤW�X������C
                                    tmpStr = tmpStr & "<td>�@</td>"
                                    HierarchyCount = (HierarchyCount +1) MOD ColumnNumber_Total         '�ξ�ثe����
                                LOOP
                                response.write tmpStr & vbNewLine
                                response.write "                </tr>" & vbNewLine
                            END IF
                            
                            response.write "            </table>" & vbNewLine
                            response.write "        </td>" & vbNewLine
                            response.write "    </tr>" & vbNewLine
                            
                        Case 1
                            '�y�u���h1�v��q�������A�ᱵ�u���h0�v�A���q���Ψӷ��u��浲���v�C�z
                            response.write "</td>" & vbNewLine
                            IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "                </tr>" & vbNewLine           '�����ƶq��F�]�w�ɡA�����@�C�C
                            
                        Case 2
                            '�y�u���h2�v��q�������A�ᱵ�u���h1�v���u���h�}�l�v�C�z
                            
                        Case 3
                            '�y�u���h3�v��q�������A�ᱵ�u���h2�v���u���h�}�l�v�C�z
                            
                        Case 4
                            IF HierarchyPrevious = HierarchyCurrent then
                                '�y�u���h4�v��u���h4�v�C�z
                                'HierarchyPrevious = HierarchyCurrent
                            ELSE
                                '�y�u���h4�v��q�������A�ᱵ�u���h3�v���u���h�}�l�v�C�z
                                'HierarchyPrevious > HierarchyCurrent
                            END IF
                            
                            
                        Case ELSE
                            response.write "<p>�y���~�G��q<br>HierarchyPrevious > HierarchyCurrent ���N�~�C�z<br>HierarchyPrevious =�y" & HierarchyPrevious & "�z�F<br>HierarchyCurrent =�y" & HierarchyCurrent & "�z�C</p>"
                            response.END
                            
                        END Select
                        
                    tmpHierarchy = tmpHierarchy -1
                    Loop Until tmpHierarchy < HierarchyCurrent
                
            Case Else
                response.write "<p>�y���~�G��q Select<br>HierarchyPrevious�G�u" & CStr(HierarchyPrevious) & "�v<br>" & "HierarchyCurrent�G�u" & CStr(HierarchyCurrent) & "�v�z</p>" & vbNewLine
                response.END
                
            End Select
    Loop
    

    response.write "</table>" & vbNewLine
    HierarchyCurrent = Nul
    HierarchyPrevious = Nul
%>
<hr>
<%
    '���G���e���
    RS.MoveFirst
    Do While NOT RS.EOF
        If vartype(HierarchyCurrent) = 0 Then HierarchyCurrent = RS("Hierarchy")
        
        '��ƪ��e�q HTML �N�X
        Select Case HierarchyCurrent
            Case 0
                '�y�u���h0�v�A�����j���A�ȯ��ު�ϥΡC�z
                '�Ģ����h�A�ȯ��ު�ϥΡC
                
            Case 1
                '�y�u���h(���)�}�l�v��u���h1�v�����C�z
                response.write "<table border='1' cellspacing='5' class='styleTableCategory'>" & vbNewLine
                response.write "    <tr><td colspan='2' class='styleHierarchy1'><a name='" & RS("Link") & "'>" & RS("Name") & "</a></td></tr>" & vbNewLine
                
            Case 2
                '�y�u���h�}�l�v��u���h2�v�����C�z
                response.write "    <tr>" & vbNewLine
                response.write "        <td class='styleHierarchy2'>" & RS("Name") & "</td>" & vbNewLine
                response.write "        <td>" & vbNewLine
                response.write "            <table border='2' cellpadding='3' cellspacing='5' class='styleTableContent'>" & vbNewLine
                HierarchyCount = 0

                
            Case 3
                '�y�u���h�}�l�v��u���h3�v�����C�z
                IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "                <tr>" & vbNewLine           '�����ƶq��F�]�w�ɡA�_�l�@�C�C
                HierarchyCount = HierarchyCount +1
                response.write "                    <td class='styleHierarchy3'><a href='" & RS("Link") & "'>" & RS("Name") & "</a>"
                
                
            Case 4
                '�y�u���h�}�l�v��u���h4�v�����C�z
                response.write "<a href='" & RS("Link") & "'>" & RS("Name") & "</a>, "
                
                
            Case ELSE
                response.write "<p>�y���~�GHierarchyPrevious < HierarchyCurrent�AIn���N�~�C�z<br>HierarchyPrevious =�y" & HierarchyPrevious & "�z�F<br>HierarchyCurrent =�y" & HierarchyCurrent & "�z�C</p>"
                response.END
            END Select
        
        
        
        '���Щ��U���ʪ�����
        HierarchyPrevious = HierarchyCurrent
        
        Do
            RS.MoveNext()
            IF NOT RS.EOF Then
                HierarchyCurrent = RS("Hierarchy")
            ELSE
                HierarchyCurrent = 0
                Exit Do
            END IF
            IF HierarchyCurrent <= 4 Then Exit Do        '���ު��a���X�Ȭ�2
        Loop
        
        
        
        '��ƪ���q HTML �N�X
        Select Case True
            Case HierarchyPrevious < HierarchyCurrent
                '�y�u�|�v�B�W�[�@�Ӷ��h�z
                Select Case HierarchyPrevious
                    Case 0
                        '�y�u���h0�v�A�����j���A�ȯ��ު�ϥΡC�z
                        
                    Case 1
                        '�y�u���h1�v����A���u���h2�v���u���h�}�l�v�C�z
                        
                        
                    Case 2
                        '�y�u���h2�v����A���u���h3�v���u���h�}�l�v�C�z
                        
                        
                    Case 3
                        '�y�u���h3�v����A���u���h4�v���u���h�}�l�v�C�z
                        response.write "<br><span class='styleHierarchy4'>"
                        
                    Case 4
                        '�y�u���h4�v����A���u���h5�v���u���h�}�l�v�C�z
                        
                        
                    Case ELSE
                        response.write "<p>�y���~�G��q<br>HierarchyPrevious < HierarchyCurrent ���N�~�C�z<br>HierarchyPrevious =�y" & HierarchyPrevious & "�z�F<br>HierarchyCurrent =�y" & HierarchyCurrent & "�z�C</p>"
                        response.END
                        
                    END Select
                
                
            Case HierarchyPrevious >= HierarchyCurrent
                '�i��|�@�����n�X�Ӷ��h
                tmpHierarchy = HierarchyPrevious
                Do
                    Select Case tmpHierarchy
                        Case 0
                            '�y�u���h0�v�A�����j���A�ȯ��ު�ϥΡC�z
                            
                        Case 1
                            '�y�u���h1�v��q�������A�ᱵ�u���h0�v�A���q���Ψӷ��u��浲���v�C�z
                            response.write "</table>" & vbNewLine
                            response.write "<hr>" & vbNewLine
                            
                            
                        Case 2
                            '�y�u���h2�v��q�������A�ᱵ�u���h1�v���u���h�}�l�v�C�z
                            IF NOT HierarchyCount MOD ColumnNumber_Total = 0 then
                                tmpStr = "                    "
                                DO UNTIL HierarchyCount MOD ColumnNumber_Total = 0          '������줣���ɡA��X�ݭn�ɤW�X������C
                                    tmpStr = tmpStr & "<td>�@</td>"
                                    HierarchyCount = (HierarchyCount +1) MOD ColumnNumber_Total         '�ξ�ثe����
                                LOOP
                                response.write tmpStr & vbNewLine
                                response.write "                </tr>" & vbNewLine
                            END IF
                            
                            response.write "            </table>" & vbNewLine
                            response.write "        </td>" & vbNewLine
                            response.write "    </tr>" & vbNewLine
                            
                        Case 3
                            '�y�u���h3�v��q�������A�ᱵ�u���h2�v���u���h�}�l�v�C�z
                            response.write "</td>" & vbNewLine
                            IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "                </tr>" & vbNewLine           '�����ƶq��F�]�w�ɡA�����@�C�C
                            
                        Case 4
                            IF HierarchyPrevious = HierarchyCurrent then
                                '�y�u���h4�v��u���h4�v�C�z
                                'HierarchyPrevious = HierarchyCurrent
                            ELSE
                                '�y�u���h4�v��q�������A�ᱵ�u���h3�v���u���h�}�l�v�C�z
                                'HierarchyPrevious > HierarchyCurrent
                                response.write "</span>"
                            END IF
                            
                            
                        Case ELSE
                            response.write "<p>�y���~�G��q<br>HierarchyPrevious > HierarchyCurrent ���N�~�C�z<br>HierarchyPrevious =�y" & HierarchyPrevious & "�z�F<br>HierarchyCurrent =�y" & HierarchyCurrent & "�z�C</p>"
                            response.END
                            
                        END Select
                        
                    tmpHierarchy = tmpHierarchy -1
                    Loop Until tmpHierarchy < HierarchyCurrent
                
            Case Else
                response.write "<p>�y���~�G��q Select<br>HierarchyPrevious�G�u" & CStr(HierarchyPrevious) & "�v<br>" & "HierarchyCurrent�G�u" & CStr(HierarchyCurrent) & "�v�z</p>" & vbNewLine
                response.END
                
            End Select
    Loop
    
    RS.Close
    DBConnection.Close
%>
</body>
</html>