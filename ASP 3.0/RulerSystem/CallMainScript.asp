<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<%
'�b�����Ĥ@��[�J�U���o��A�`�N���|�n��A���t�̫e������޸�'
'<!-- #include file="CallMainScript.asp" -->

'�ŧi�q���ܼ�
Dim Transmit, FrontTip, CounterNum, CounterNumView
%>

<%'�� = = = = = = = = = = = = = = = ���e�{�� = = = = = = = = = = = = = = = ��%>

<%
'�סססס� ���o�ƶq �סססס�
Function FunctionGetRowNum(OpenDboName)
'Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
    RS.MoveFirst
    ArrayRowNum=0
    Do While RS.EOF = False
        RS.Find="[RulerOS]='" & CStr(RulerOS) & "'"
        if NOT RS.EOF then
            Dim rsSubTitle
            rsSubTitle = CStr("" & RS("SubTitle"))
            if ((rsSubTitle = "") OR (rsSubTitle = SubTitle)) then 
                ArrayRowNum = ArrayRowNum +1
            End if
            RS.MoveNext
        End if
    Loop
RS.Close
End Function
%>

<%'�� = = = = = = = = = = = = = = = Sub = = = = = = = = = = = = = = = ��%>

<%
'�סססס� ��ĳ���A �סססס�
Sub SubStateToSuggestType()
    call FunctionStateToSuggestType(dboSS)
    Response.Write Transmit
End Sub
%>

<%
'�סססס� �w�]���A �סססס�
Sub SubStateToDefaultType()
    call FunctionStateToDefaultType(dboSS)
    Response.Write Transmit
End Sub
%>

<%
'�סססס� �ӧO�A�ȳ]�w(���D) �סססס�
Sub SubServiceParticularMenu()
    call FunctionServiceParticularMenu(dboSS)
    Transmit = Transmit & FrontSpace & "------------------------------"
    Response.Write Transmit
End Sub
%>

<%
'�סססס� �ӧO�A�ȳ]�w(�ﶵ) �סססס�
Sub SubServiceParticularOption()
    call FunctionServiceParticularOption(dboSS)
    Response.Write Transmit
End Sub
%>

<%
'�סססס� �t�Υ\����D �סססס�
Sub SubSystemFunctionMenu()
    call FunctionSystemFunctionMenu(dboSF)
    Transmit = Transmit & FrontSpace & "------------------------------"
    Response.Write Transmit
End Sub
%>

<%
'�סססס� �t�Υ\��ﶵ �סססס�
Sub SubSystemFunctionOption()
    call FunctionSystemFunctionOption(dboSF)
    Response.Write Transmit
End Sub
%>

<%'�� = = = = = = = = = = = = = = = Function = = = = = = = = = = = = = = = ��%>

<%
'�סססס� ��ĳ���A �סססס�
Function FunctionStateToSuggestType(OpenDboName)
'Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
    RS.MoveFirst
    Transmit=""
    Do While RS.EOF = False
        RS.Find="[RulerOS]='" & RulerOS & "'"
        if NOT RS.EOF then
            Dim rsSubTitle
            rsSubTitle = CStr("" & RS("SubTitle"))
            if ((rsSubTitle = "") OR (rsSubTitle = SubTitle)) then 
                Transmit = Transmit & FrontSpace & "echo �ܧ�A�ȦW��: ��" & Trim(RS("ServiceView")) & "��" & vbNewLine
                Transmit = Transmit & FrontSpace & "echo     ��ĳ�]�w���A: " & Trim(RS("ServiceSuggestType")) & vbNewLine
                Transmit = Transmit & FrontSpace & "sc config """ & Trim(RS("ServiceName")) & """ start= " & Trim(RS("ServiceSuggestType")) & vbNewLine
                Transmit = Transmit & vbNewLine
            End if
            RS.MoveNext
        End if
    Loop
RS.Close
End Function
%>

<%
'�סססס� �w�]���A �סססס�
Function FunctionStateToDefaultType(OpenDboName)
'Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
    RS.MoveFirst
    Transmit=""
    Do While RS.EOF = False
        RS.Find="[RulerOS]='" & RulerOS & "'"
        if NOT RS.EOF then
            Dim rsSubTitle
            rsSubTitle = CStr("" & RS("SubTitle"))
            if ((rsSubTitle = "") OR (rsSubTitle = SubTitle)) then 
                Transmit = Transmit & FrontSpace & "echo �ܧ�A�ȦW��: ��" & Trim(RS("ServiceView")) & "��" & vbNewLine
                Transmit = Transmit & FrontSpace & "echo     �w�]�]�w���A: " & Trim(RS("ServiceDefaultType")) & vbNewLine
                Transmit = Transmit & FrontSpace & "sc config """ & Trim(RS("ServiceName")) & """ start= " & Trim(RS("ServiceDefaultType")) & vbNewLine
                Transmit = Transmit & vbNewLine
            End if
            RS.MoveNext
        End if
    Loop
RS.Close
End Function
%>

<%
'�סססס� �ӧO�A�ȳ]�w(���D) �סססס�
Function FunctionServiceParticularMenu(OpenDboName)
'Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
    RS.MoveFirst
    Transmit=""
    CounterNum=0
    CounterNumView=0
    Do While RS.EOF = False
        RS.Find="[RulerOS]='" & RulerOS & "'"
        if NOT RS.EOF then
            Dim rsSubTitle
            rsSubTitle = CStr("" & RS("SubTitle"))
            if ((rsSubTitle = "") OR (rsSubTitle = SubTitle)) then 
                CounterNum = CounterNum +1
                
                '�N > 10 ���ȧ令�^�����
                if CounterNum < 10 then
                    CounterNumView=Chr(CounterNum+48)
                    ELSE
                    CounterNumView=Chr(CounterNum+97-10)
                End if
                
                Transmit = Transmit & FrontSpace & " " & CounterNumView & ". " & Trim(RS("ServiceView")) & vbNewLine
            End if
            RS.MoveNext
        End if
    Loop
RS.Close
End Function
%>

<%
'�סססס� �ӧO�A�ȳ]�w(�ﶵ) �סססס�
Function FunctionServiceParticularOption(OpenDboName)
'Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
    RS.MoveFirst
    Transmit=""
    CounterNum=0
    Do While RS.EOF = False
        RS.Find="[RulerOS]='" & RulerOS & "'"
        if NOT RS.EOF then
            Dim rsSubTitle
            rsSubTitle = CStr("" & RS("SubTitle"))
            if ((rsSubTitle = "") OR (rsSubTitle = SubTitle)) then 
                CounterNum = CounterNum +1
                Transmit = Transmit & FrontSpace & ":Ch_3-" & CounterNum & vbNewLine
                Transmit = Transmit & FrontSpace & "    rem ---------- �ާ@�A�ȸ��: " & Trim(RS("ServiceName")) & " ----------" & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceView=" & Trim(RS("ServiceView")) & """ " & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceName=" & Trim(RS("ServiceName")) & """ " & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceDescribe=�нs�覹�ɬd�� " & vbNewLine
                Transmit = Transmit & FrontSpace & "    rem ""ServiceDescribe=" & Trim(RS("ServiceDescribe")) & """ " & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceDefaultType=" & Trim(RS("ServiceDefaultType")) & """ " & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceSuggestType=" & Trim(RS("ServiceSuggestType")) & """ " & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceMemo=" & Trim(RS("ServiceMemo")) & """ " & vbNewLine
                Transmit = Transmit & FrontSpace & "    goto Ch_3-Menu" & vbNewLine
                Transmit = Transmit & FrontSpace & "    goto end" & vbNewLine & vbNewLine
            End if
            RS.MoveNext
        End if
    Loop
RS.Close
End Function
%>

<%
'�סססס� �t�Υ\����D �סססס�
Function FunctionSystemFunctionMenu(OpenDboName)
'Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
    RS.MoveFirst
    Transmit=""
    CounterNum=0
    Do While RS.EOF = False
        RS.Find="[RulerOS]='" & RulerOS & "'"
        if NOT RS.EOF then
            Dim rsSubTitle
            rsSubTitle = CStr("" & RS("SubTitle"))
            if ((rsSubTitle = "") OR (rsSubTitle = SubTitle)) then 
                CounterNum = CounterNum +1
                Transmit = Transmit & FrontSpace & " " & CounterNum & ". " & Trim(RS("Title")) & vbNewLine
            End if
            RS.MoveNext
        End if
    Loop
RS.Close
End Function
%>

<%
'�סססס� �t�Υ\��ﶵ �סססס�
Function FunctionSystemFunctionOption(OpenDboName)
'Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
    RS.MoveFirst
    Transmit=""
    CounterNum=0
    Do While RS.EOF = False
        RS.Find="[RulerOS]='" & RulerOS & "'"
        if NOT RS.EOF then
            Dim rsSubTitle
            rsSubTitle = CStr("" & RS("SubTitle"))
            if ((rsSubTitle = "") OR (rsSubTitle = SubTitle)) then 
                CounterNum = CounterNum +1
                Transmit = Transmit & FrontSpace & ":Ch_4-" & CounterNum & vbNewLine
                Transmit = Transmit & FrontSpace & "    rem ---------- " & Trim(RS("Title")) & " ----------" & vbNewLine
                Transmit = Transmit & FrontSpace & "    echo " & Trim(RS("Title")) & vbNewLine
                Transmit = Transmit & FrontSpace & "    echo " & Trim(RS("Description")) & vbNewLine
                Transmit = Transmit & FrontSpace & "    " & Trim(RS("Command")) & " """ & Trim(RS("Target")) & """ " & Trim(RS("Argument")) & vbNewLine
                Transmit = Transmit & FrontSpace & "    goto end" & vbNewLine & vbNewLine
            End if
            RS.MoveNext
        End if
    Loop
RS.Close
End Function
%>
