<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<!-- #include file="CallMain_ConnectionDataTable.asp" -->
<%
'�b�����Ĥ@��[�J�U���o��A�`�N���|�n��A���t�̫e������޸�'
'<!-- #include file="CallMainScript.asp" -->

'�ŧi�q���ܼ�
Dim RowData
%>

<%'�� = = = = = = = = = = = = = = = Sub = = = = = = = = = = = = = = = ��%>

<%
'�סססס� ��ܳs��: �M�� �סססס�
Sub SubTitleGetBatch()
    if NOT LastXlsFile = Empty then
        response.write "<td align=""center""><b>�妸�ɤU��</b>�@<a href=""index.asp""><font size=""2"">�M��</font></a></td>" & vbNewLine
    End If
End Sub
%>

<%
'�סססס� ��ܤU���s�� �סססס�
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
'�סססס� ���o��ƦC�� �סססס�
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
'�סססס� ���o�{�b�ɶ� �סססס�
Sub SubGetNow()
    ResultTime1 = Year(now) & "/" & Right("0" & Month(now), 2) & "/" & Right("0" & Day(now), 2)
    ResultTime2 = Right("0" & Hour(now), 2) & ":" & Right("0" & Minute(now), 2) & ":" & Right("0" & Second(now), 2)
    response.write ResultTime1 & " " & ResultTime2
End Sub
%>

<%
'�סססס� ���o��s�ɸ�� �סססס�
Sub SubGetBatchData(Sheet)
    Reboot_MePriority=""
    RowNum=0
    RS.MoveFirst
    Do While RS.EOF = False
        RowNum = RowNum +1
        response.write "echo (" & right("000" & RowNum,3) & "/%" & Replace(Sheet," ","_") & "Num%). " & Trim(RS("Title")) & vbNewLine
        
        response.write "IF NOT EXIST " & Chr(34) & "%SYSTEMDRIVE%\_UpdateProgress\" & Sheet & Chr(95) & Trim(RS("FileName")) & ".UdB" & Chr(34) & " (" & vbNewLine
        response.write "    echo      �{�b�ɶ�: %time%" & vbNewLine
        response.write "    echo      " & Trim( Replace(Replace(RS("Content"),")", "^)"),"(", "^(")) & vbNewLine
        response.write "    " & Trim(RS("Command")) & " " & Trim(RS("FileName")) & " " & Trim(RS("Argument")) & vbNewLine
        response.write "    echo %time% > " & Chr(34) & "%SYSTEMDRIVE%\_UpdateProgress\" & Sheet & Chr(95) & Trim(RS("FileName")) & ".UdB" & Chr(34) & vbNewLine
        response.write "    start /wait " & Chr(34) & "GrandTotalReset" & Chr(34) & Chr(32) & Chr(34) & "..\GrandTotalReset.bat" & Chr(34)
        
        '�j��ҧP�_: Priority_Reboot �C�ӳ��n���}��; Priority ���P�խn���}��
        Select Case Sheet
            Case "Intervene"
                
            Case "Priority_Reboot"
                response.write " 1"
            Case Else
                Reboot_MePriority = Trim(RS("Priority"))
                RS.MoveNext
                IF RS.EOF = False Then
                    '�٦����
                    IF Reboot_MePriority <> Trim(RS("Priority")) Then response.write " 1"
                    RS.moveprevious
                ELSE
                    '��Ƶ���
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

<%'�� = = = = = = = = = = = = = = = Function = = = = = = = = = = = = = = = ��%>

<%
'�סססס� index.asp �W�s�� �סססס�
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
'�סססס� ���o�U���s�� �סססס�
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

