<%
    '● ============================= 設定宣告 ============================= ●
    Dim WebMarkVer
        WebMarkVer = "1.0a"         '套用模組設定的 模組版本
    Dim WebMarkVerDate
        WebMarkVerDate = "2013/10/25"
 
    Dim xlsDataBasePathFile         '設定 xls 檔案名稱(可包含相對路徑)
        xlsDataBasePathFile = "WebMark.xls"
    Dim hTitle         'Title 名稱
        hTitle = Request("htmlTitle")
        If hTitle = Empty Then hTitle = "書籤下載器"
    Dim webLoadingFrequency            '取得網頁載入的次數
        webLoadingFrequency = Request("webLoadingFrequency")
        If webLoadingFrequency = null Then
            webLoadingFrequency = 1
        Else
            webLoadingFrequency = webLoadingFrequency + 1
        End If
    
    Dim ColumnNumber_Total          '設定資料的欄位總數
        ColumnNumber_Total = CInt(Request("Textbox_ColumnNumber_Total"))
    Dim IndexTable_Name         '索引表的標題名稱
        IndexTable_Name = Request("IndexTable_Name")
    Dim styleHierarchy1           '網站類型的欄位寬度
        styleHierarchy1 = Request("styleHierarchy1")
    Dim tableColumn2Width         '網站表格的欄位寬度
        If webLoadingFrequency > 1 Then tableColumn2Width = 100 / ColumnNumber_Total
 
    Dim DBConnection
    Dim RS
 
    Dim HierarchyPrevious        '前一個項目的階層(Hierarchy)
    Dim HierarchyCurrent        '當前項目的階層(Hierarchy)
    Dim HierarchyCount        '計算當前項目是連續第？個同階層的，用來計算表格欄位數

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
<p style="text-align: right" class="styleWebMarkVer">模組版本：<%=WebMarkVer%></p>
<form method="POST" action="index.asp">
    <table align="center" border="3" cellpadding="3" cellspacing="5" class="styleMainTable" style="border: thick double #800000;">
        <caption style="font-family: 華康海報體W9; font-size: xx-large; color: #FF0000;"><%=hTitle%></caption>
            <tr><td dir="rtl" class="styleHalfWidth">名稱</td>
                <td dir="ltr" class="styleHalfWidth">設定值</td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">網頁 Title 名稱</td>
                <td dir="ltr" class="styleHalfWidth"><input type="text" name="htmlTitle" size="20" value="庭庭的書籤" tabindex="1"></td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">欄位數量</td>
                <td dir="ltr" class="styleHalfWidth"><input type="text" name="Textbox_ColumnNumber_Total" size="4" value="4" tabindex="2"></td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">索引表格名稱</td>
                <td dir="ltr" class="styleHalfWidth"><input type="text" name="IndexTable_Name" size="20" value="索引"></td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">網站類型的欄位寬度</td>
                <td dir="ltr" class="styleHalfWidth"><input type="text" name="styleHierarchy1" size="4" value="10%"></td></tr>
            <tr><td dir="rtl" class="styleHalfWidth">&nbsp;</td>
                <td dir="ltr" class="styleHalfWidth">&nbsp;</td></tr>
        </table>
    <p style="text-align: center"><input type="submit" value="下載" name="B1"><input type="reset" value="重新設定" name="B2"></p>
    <input type="hidden" name="webLoadingFrequency" value="<%=webLoadingFrequency%>">
</form>


<p>&nbsp;</p>

</body>
</html>


<%
IF webLoadingFrequency < 2 then Response.End	'網頁載入計數，第2次之後才產生下載內容。


'● ============================= 產生下載 ============================= ●
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=WebMark.htm"
Response.ContentType = "text"
'Response.End


'● ============================= 資料連線 ============================= ●
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
        /*索引表*/
        {
            width: 100%;
            border-style: solid;
            border-width: 3px;
        }
        .styleIndexTable_2
        /*索引內表*/
        {
            width: 100%;
            border: 2px solid #800000;
        }
        .styleIndexTableTitle
        /*索引表的標題字樣*/
        {
            font-family: 華康少女文字W7;
            font-size: xx-large; color: #800000;
        }
        .styleIndexTableHierarchy0
        /*索引表的階層0字樣*/
        {
            font-family: 華康海報體W12;
            font-size: x-large; color: #009900;
            width: 100px;
        }
        .styleIndexTableHierarchy1
        /*索引表的階層1字樣*/
        {
            font-size: larger;
            width: <%=tableColumn2Width%>%;
        }
        .styleIndexTableHierarchy2
        /*索引內表的階層2字樣*/
        {
            font-size: small;
        }
        .styleTopLink_1
        /*快速連結表*/
        {
            width: 100%;
            border-style: solid;
            border-width: 3px;
        }
        .styleTopLink_2
        /*快速連結內表*/
        {
            width: 100%;
            border: 2px solid #800000;
        }
        .styleTableCategory
        /*類別外表*/
        {
            width: 100%;
            border-style: solid;
            border-width: 3px;
        }
        .styleTableContent
        /*類別內表*/
        {
            width: 100%;
            border: 2px solid #800000;
        }
        .styleHierarchy1
        /*階層1 標題字樣*/
        {
            font-family: 華康少女文字W7;
            font-size: xx-large; color: #CC3399;
        }
        .styleHierarchy2
        /*階層2 標題字樣*/
        /*表格第一欄寬度*/
        {
            font-family: 華康海報體W12;
            font-size: x-large; color: #FF6600;
            width: 135px;
        }
        .styleHierarchy3
        /*階層3 名稱字樣*/
        /*表中表的平均欄位寬度*/
        {
            font-size: larger; color: #800000;
            width: <%=tableColumn2Width%>%;
        }
        .styleHierarchy4
        /*階層4 文字字樣*/
        {
            font-size: small; color: #000000;
        }
    </style>
</head>
<body>
<div align="right"><font size="2" color="#C0C0C0">版本：<%=now%></font></div>

<%
    '表格：索引表
    RS.MoveFirst
    
    response.write "<table border='1' cellspacing='5' class='styleIndexTable'>" & vbNewLine
    response.write "    <tr><td colspan='2' class='styleIndexTableTitle'>" & IndexTable_Name & "</td></tr>" & vbNewLine

    Do While NOT RS.EOF
        If vartype(HierarchyCurrent) = 0 Then HierarchyCurrent = RS("Hierarchy")
        
        '資料的前段 HTML 代碼
        Select Case HierarchyCurrent
            Case 0
                '『「階層0」，分類大項，僅索引表使用。』
                '第０階層，僅索引表使用。
                response.write "    <tr>" & vbNewLine
                response.write "        <td class='styleIndexTableHierarchy0'>" & RS("Name") & "</td>" & vbNewLine
                response.write "        <td>" & vbNewLine
                response.write "            <table border='2' cellpadding='3' cellspacing='5' class='styleIndexTable_2'>" & vbNewLine
                HierarchyCount = 0
                
            Case 1
                '『「階層(表格)開始」到「階層1」之間。』
                IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "                <tr>" & vbNewLine           '當欄位數量到達設定時，起始一列。
                HierarchyCount = HierarchyCount +1
                response.write "                    <td class='styleIndexTableHierarchy1'><a target='_self' href='#" & RS("Link") & "'>" & RS("Name") & "</a>"
                
                
            Case 2
                '『「階層開始」到「階層2」之間。』
                response.write RS("Name") & ", "
                
                
            Case 3
                '『「階層開始」到「階層3」之間。』
                
                
            Case 4
                '『「階層開始」到「階層4」之間。』
                
                
            Case ELSE
                response.write "<p>『錯誤：HierarchyPrevious < HierarchyCurrent，In的意外。』<br>HierarchyPrevious =『" & HierarchyPrevious & "』；<br>HierarchyCurrent =『" & HierarchyCurrent & "』。</p>"
                response.END
            END Select
        
        
        
        '指標往下移動的部分
        HierarchyPrevious = HierarchyCurrent
        
        Do
            RS.MoveNext()
            IF NOT RS.EOF Then
                HierarchyCurrent = RS("Hierarchy")
            ELSE
                HierarchyCurrent = 0
                Exit Do
            END IF
            IF HierarchyCurrent <= 2 Then Exit Do        '索引表格的跳出值為2；內容表格的值為最大值4。
        Loop
        
        
        
        '資料的後段 HTML 代碼
        Select Case True
            Case HierarchyPrevious < HierarchyCurrent
                '『只會逐步增加一個階層』
                Select Case HierarchyPrevious
                    Case 0
                        '『「階層0」，分類大項，僅索引表使用。』
                        
                    Case 1
                        '『「階層1」之後，接「階層2」的「階層開始」。』
                        response.write "<br>" & vbNewLine
                        response.write "                        <span class='styleIndexTableHierarchy2'>"
                        
                    Case 2
                        '『「階層2」之後，接「階層3」的「階層開始」。』
                        
                        
                    Case 3
                        '『「階層3」之後，接「階層4」的「階層開始」。』   

                        
                    Case 4
                        '『「階層4」之後，接「階層5」的「階層開始」。』
                        
                        
                    Case ELSE
                        response.write "<p>『錯誤：後段<br>HierarchyPrevious < HierarchyCurrent 的意外。』<br>HierarchyPrevious =『" & HierarchyPrevious & "』；<br>HierarchyCurrent =『" & HierarchyCurrent & "』。</p>"
                        response.END
                        
                    END Select
                
                
            Case HierarchyPrevious >= HierarchyCurrent
                '可能會一次降好幾個階層
                tmpHierarchy = HierarchyPrevious
                Do
                    Select Case tmpHierarchy
                        Case 0
                            '『「階層0」，分類大項，僅索引表使用。』
                            IF NOT HierarchyCount MOD ColumnNumber_Total = 0 then
                                tmpStr = "                    "
                                DO UNTIL HierarchyCount MOD ColumnNumber_Total = 0          '當表格欄位不足時，算出需要補上幾格空欄位。
                                    tmpStr = tmpStr & "<td>　</td>"
                                    HierarchyCount = (HierarchyCount +1) MOD ColumnNumber_Total         '統整目前欄位數
                                LOOP
                                response.write tmpStr & vbNewLine
                                response.write "                </tr>" & vbNewLine
                            END IF
                            
                            response.write "            </table>" & vbNewLine
                            response.write "        </td>" & vbNewLine
                            response.write "    </tr>" & vbNewLine
                            
                        Case 1
                            '『「階層1」到段落結尾，後接「階層0」，此段落用來當做「表格結束」。』
                            response.write "</td>" & vbNewLine
                            IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "                </tr>" & vbNewLine           '當欄位數量到達設定時，結束一列。
                            
                        Case 2
                            '『「階層2」到段落結尾，後接「階層1」的「階層開始」。』
                            
                        Case 3
                            '『「階層3」到段落結尾，後接「階層2」的「階層開始」。』
                            
                        Case 4
                            IF HierarchyPrevious = HierarchyCurrent then
                                '『「階層4」到「階層4」。』
                                'HierarchyPrevious = HierarchyCurrent
                            ELSE
                                '『「階層4」到段落結尾，後接「階層3」的「階層開始」。』
                                'HierarchyPrevious > HierarchyCurrent
                            END IF
                            
                            
                        Case ELSE
                            response.write "<p>『錯誤：後段<br>HierarchyPrevious > HierarchyCurrent 的意外。』<br>HierarchyPrevious =『" & HierarchyPrevious & "』；<br>HierarchyCurrent =『" & HierarchyCurrent & "』。</p>"
                            response.END
                            
                        END Select
                        
                    tmpHierarchy = tmpHierarchy -1
                    Loop Until tmpHierarchy < HierarchyCurrent
                
            Case Else
                response.write "<p>『錯誤：後段 Select<br>HierarchyPrevious：「" & CStr(HierarchyPrevious) & "」<br>" & "HierarchyCurrent：「" & CStr(HierarchyCurrent) & "」』</p>" & vbNewLine
                response.END
                
            End Select
    Loop
    

    response.write "</table>" & vbNewLine
    HierarchyCurrent = Nul
    HierarchyPrevious = Nul
%>
<hr>
<%
    '表格：內容表格
    RS.MoveFirst
    Do While NOT RS.EOF
        If vartype(HierarchyCurrent) = 0 Then HierarchyCurrent = RS("Hierarchy")
        
        '資料的前段 HTML 代碼
        Select Case HierarchyCurrent
            Case 0
                '『「階層0」，分類大項，僅索引表使用。』
                '第０階層，僅索引表使用。
                
            Case 1
                '『「階層(表格)開始」到「階層1」之間。』
                response.write "<table border='1' cellspacing='5' class='styleTableCategory'>" & vbNewLine
                response.write "    <tr><td colspan='2' class='styleHierarchy1'><a name='" & RS("Link") & "'>" & RS("Name") & "</a></td></tr>" & vbNewLine
                
            Case 2
                '『「階層開始」到「階層2」之間。』
                response.write "    <tr>" & vbNewLine
                response.write "        <td class='styleHierarchy2'>" & RS("Name") & "</td>" & vbNewLine
                response.write "        <td>" & vbNewLine
                response.write "            <table border='2' cellpadding='3' cellspacing='5' class='styleTableContent'>" & vbNewLine
                HierarchyCount = 0

                
            Case 3
                '『「階層開始」到「階層3」之間。』
                IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "                <tr>" & vbNewLine           '當欄位數量到達設定時，起始一列。
                HierarchyCount = HierarchyCount +1
                response.write "                    <td class='styleHierarchy3'><a href='" & RS("Link") & "'>" & RS("Name") & "</a>"
                
                
            Case 4
                '『「階層開始」到「階層4」之間。』
                response.write "<a href='" & RS("Link") & "'>" & RS("Name") & "</a>, "
                
                
            Case ELSE
                response.write "<p>『錯誤：HierarchyPrevious < HierarchyCurrent，In的意外。』<br>HierarchyPrevious =『" & HierarchyPrevious & "』；<br>HierarchyCurrent =『" & HierarchyCurrent & "』。</p>"
                response.END
            END Select
        
        
        
        '指標往下移動的部分
        HierarchyPrevious = HierarchyCurrent
        
        Do
            RS.MoveNext()
            IF NOT RS.EOF Then
                HierarchyCurrent = RS("Hierarchy")
            ELSE
                HierarchyCurrent = 0
                Exit Do
            END IF
            IF HierarchyCurrent <= 4 Then Exit Do        '索引表格地跳出值為2
        Loop
        
        
        
        '資料的後段 HTML 代碼
        Select Case True
            Case HierarchyPrevious < HierarchyCurrent
                '『只會逐步增加一個階層』
                Select Case HierarchyPrevious
                    Case 0
                        '『「階層0」，分類大項，僅索引表使用。』
                        
                    Case 1
                        '『「階層1」之後，接「階層2」的「階層開始」。』
                        
                        
                    Case 2
                        '『「階層2」之後，接「階層3」的「階層開始」。』
                        
                        
                    Case 3
                        '『「階層3」之後，接「階層4」的「階層開始」。』
                        response.write "<br><span class='styleHierarchy4'>"
                        
                    Case 4
                        '『「階層4」之後，接「階層5」的「階層開始」。』
                        
                        
                    Case ELSE
                        response.write "<p>『錯誤：後段<br>HierarchyPrevious < HierarchyCurrent 的意外。』<br>HierarchyPrevious =『" & HierarchyPrevious & "』；<br>HierarchyCurrent =『" & HierarchyCurrent & "』。</p>"
                        response.END
                        
                    END Select
                
                
            Case HierarchyPrevious >= HierarchyCurrent
                '可能會一次降好幾個階層
                tmpHierarchy = HierarchyPrevious
                Do
                    Select Case tmpHierarchy
                        Case 0
                            '『「階層0」，分類大項，僅索引表使用。』
                            
                        Case 1
                            '『「階層1」到段落結尾，後接「階層0」，此段落用來當做「表格結束」。』
                            response.write "</table>" & vbNewLine
                            response.write "<hr>" & vbNewLine
                            
                            
                        Case 2
                            '『「階層2」到段落結尾，後接「階層1」的「階層開始」。』
                            IF NOT HierarchyCount MOD ColumnNumber_Total = 0 then
                                tmpStr = "                    "
                                DO UNTIL HierarchyCount MOD ColumnNumber_Total = 0          '當表格欄位不足時，算出需要補上幾格空欄位。
                                    tmpStr = tmpStr & "<td>　</td>"
                                    HierarchyCount = (HierarchyCount +1) MOD ColumnNumber_Total         '統整目前欄位數
                                LOOP
                                response.write tmpStr & vbNewLine
                                response.write "                </tr>" & vbNewLine
                            END IF
                            
                            response.write "            </table>" & vbNewLine
                            response.write "        </td>" & vbNewLine
                            response.write "    </tr>" & vbNewLine
                            
                        Case 3
                            '『「階層3」到段落結尾，後接「階層2」的「階層開始」。』
                            response.write "</td>" & vbNewLine
                            IF HierarchyCount MOD ColumnNumber_Total = 0 then response.write "                </tr>" & vbNewLine           '當欄位數量到達設定時，結束一列。
                            
                        Case 4
                            IF HierarchyPrevious = HierarchyCurrent then
                                '『「階層4」到「階層4」。』
                                'HierarchyPrevious = HierarchyCurrent
                            ELSE
                                '『「階層4」到段落結尾，後接「階層3」的「階層開始」。』
                                'HierarchyPrevious > HierarchyCurrent
                                response.write "</span>"
                            END IF
                            
                            
                        Case ELSE
                            response.write "<p>『錯誤：後段<br>HierarchyPrevious > HierarchyCurrent 的意外。』<br>HierarchyPrevious =『" & HierarchyPrevious & "』；<br>HierarchyCurrent =『" & HierarchyCurrent & "』。</p>"
                            response.END
                            
                        END Select
                        
                    tmpHierarchy = tmpHierarchy -1
                    Loop Until tmpHierarchy < HierarchyCurrent
                
            Case Else
                response.write "<p>『錯誤：後段 Select<br>HierarchyPrevious：「" & CStr(HierarchyPrevious) & "」<br>" & "HierarchyCurrent：「" & CStr(HierarchyCurrent) & "」』</p>" & vbNewLine
                response.END
                
            End Select
    Loop
    
    RS.Close
    DBConnection.Close
%>
</body>
</html>