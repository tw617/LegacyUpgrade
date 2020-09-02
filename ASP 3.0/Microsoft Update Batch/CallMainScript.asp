<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<!-- #include file="CallMain_ConnectionDataTable.asp" -->
<%
'在網頁第一行加入下面這行，注意路徑要改，不含最前面的單引號'
'<!-- #include file="CallMainScript.asp" -->

'宣告通用變數
Dim RowData
%>

<%'● = = = = = = = = = = = = = = = Sub = = = = = = = = = = = = = = = ●%>

<%
'＝＝＝＝＝ 顯示連結: 清除 ＝＝＝＝＝
Sub SubTitleGetBatch()
    if NOT LastXlsFile = Empty then
        response.write "<td align=""center""><b>批次檔下載</b>　<a href=""index.asp""><font size=""2"">清除</font></a></td>" & vbNewLine
    End If
End Sub
%>

<%
'＝＝＝＝＝ 顯示下載連結 ＝＝＝＝＝
Sub SubGetBatch()
    if NOT LastXlsFile = Empty then
        RowData = Empty
        RowData = RowData & "<td rowspan=""8"">" & vbNewLine
        call FunctionGetBatch()
        RowData = RowData & "</td>" & vbNewLine
        response.write RowData
    End If
End Sub
%>

<%
'＝＝＝＝＝ 取得資料列數 ＝＝＝＝＝
Sub SubGetRowCount()
    RowCount = 0
    RS.MoveFirst
    Do While RS.EOF = False
        RowCount = RowCount +1
        RS.MoveNext
    Loop
    response.write Right("000" & RowCount,3)
End Sub
%>

<%
'＝＝＝＝＝ 取得現在時間 ＝＝＝＝＝
Sub SubGetNow()
    ResultTime1 = Year(now) & "/" & Right("0" & Month(now), 2) & "/" & Right("0" & Day(now), 2)
    ResultTime2 = Right("0" & Hour(now), 2) & ":" & Right("0" & Minute(now), 2) & ":" & Right("0" & Second(now), 2)
    response.write ResultTime1 & " " & ResultTime2
End Sub
%>

<%
'＝＝＝＝＝ 取得更新檔資料 ＝＝＝＝＝
Sub SubGetBatchData(Sheet)
    Reboot_MePriority=""
    RowNum=0
    RS.MoveFirst
    Do While RS.EOF = False
        RowNum = RowNum +1
        response.write "echo (" & right("000" & RowNum,3) & "/%" & Replace(Sheet," ","_") & "Num%). " & Trim(RS("Title")) & vbNewLine
        
        response.write "IF NOT EXIST " & Chr(34) & "%SYSTEMDRIVE%\_UpdateProgress\" & Sheet & Chr(95) & Trim(RS("FileName")) & ".UdB" & Chr(34) & " (" & vbNewLine
        response.write "    echo      現在時間: %time%" & vbNewLine
        response.write "    echo      " & Trim( Replace(Replace(RS("Content"),")", "^)"),"(", "^(")) & vbNewLine
        response.write "    " & Trim(RS("Command")) & " " & Trim(RS("FileName")) & " " & Trim(RS("Argument")) & vbNewLine
        response.write "    echo %time% > " & Chr(34) & "%SYSTEMDRIVE%\_UpdateProgress\" & Sheet & Chr(95) & Trim(RS("FileName")) & ".UdB" & Chr(34) & vbNewLine
        response.write "    start /wait " & Chr(34) & "GrandTotalReset" & Chr(34) & Chr(32) & Chr(34) & "..\GrandTotalReset.bat" & Chr(34)
        
        '強制重啟判斷: Priority_Reboot 每個都要重開機; Priority 不同組要重開機
        Select Case Sheet
            Case "Intervene"
                
            Case "Priority_Reboot"
                response.write " 1"
            Case Else
                Reboot_MePriority = Trim(RS("Priority"))
                RS.MoveNext
                IF RS.EOF = False Then
                    '還有資料
                    IF Reboot_MePriority <> Trim(RS("Priority")) Then response.write " 1"
                    RS.moveprevious
                ELSE
                    '資料結束
                    response.write " 1"
                End IF
        End Select
        response.write vbNewLine
        response.write "    )" & vbNewLine
        IF RS.EOF = False Then RS.MoveNext
    Loop
    response.write "echo " & Chr("7") & vbNewLine
    response.write "exit" & vbNewLine

End Sub
%>

<%'● = = = = = = = = = = = = = = = Function = = = = = = = = = = = = = = = ●%>

<%
'＝＝＝＝＝ index.asp 超連結 ＝＝＝＝＝
Function FunctionHyperlinks(ArrayXlsFile)
xlsFile=Split(ArrayXlsFile,";")(0)
Title=Split(ArrayXlsFile,";")(1)

    if LastXlsFile = Empty then
        response.write "<TD><a href=""index.asp?xlsFile=" & xlsFile & """>" & Title & "</a></TD>" & vbNewLine
    ELSE
        if xlsFile = LastXlsFile then
            response.write "<TD><a href=""index.asp?xlsFile=" & xlsFile & """><b>" & Title & "</b></a></TD>" & vbNewLine
        ELSE
            response.write "<TD bgcolor=""#C0C0C0""><a href=""index.asp?xlsFile=" & xlsFile & """><font color=""#FFFFFF"">" & Title & "</font></a></TD>" & vbNewLine
        End If
    End If
End Function
%>

<%
'＝＝＝＝＝ 取得下載連結 ＝＝＝＝＝
Function FunctionGetBatch()
    RowData = RowData & "<table border=""1"" width=""100%"" bordercolor=""#800080"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine
    call FunctionDataSource(xlsFilePath)
    RS.Open "[" & dboIndex & "]",DBConnection,3,1
        RS.MoveFirst
        Do While RS.EOF = False
        
            RowData = RowData & "<TR><TD>"
            RowData = RowData & "<a href=""BatchMaker.asp?xlsFilePath=" & xlsFilePath & "&Sheet=" & Trim(RS("Sheet")) & """>" & Trim(RS("Sheet")) & "</a>"
            RowData = RowData & "</TD></TR>" & vbNewLine

            RS.MoveNext
        Loop
    RS.Close
    DBConnection.close
    RowData = RowData & "</table>"
End Function
%>

