<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<%
'在網頁第一行加入下面這行，注意路徑要改，不含最前面的單引號'
'<!-- #include file="CallMainScript.asp" -->

'宣告通用變數
Dim Transmit, FrontTip, CounterNum, CounterNumView
%>

<%'● = = = = = = = = = = = = = = = 選單前程式 = = = = = = = = = = = = = = = ●%>

<%
'＝＝＝＝＝ 取得數量 ＝＝＝＝＝
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

<%'● = = = = = = = = = = = = = = = Sub = = = = = = = = = = = = = = = ●%>

<%
'＝＝＝＝＝ 建議狀態 ＝＝＝＝＝
Sub SubStateToSuggestType()
    call FunctionStateToSuggestType(dboSS)
    Response.Write Transmit
End Sub
%>

<%
'＝＝＝＝＝ 預設狀態 ＝＝＝＝＝
Sub SubStateToDefaultType()
    call FunctionStateToDefaultType(dboSS)
    Response.Write Transmit
End Sub
%>

<%
'＝＝＝＝＝ 個別服務設定(標題) ＝＝＝＝＝
Sub SubServiceParticularMenu()
    call FunctionServiceParticularMenu(dboSS)
    Transmit = Transmit & FrontSpace & "------------------------------"
    Response.Write Transmit
End Sub
%>

<%
'＝＝＝＝＝ 個別服務設定(選項) ＝＝＝＝＝
Sub SubServiceParticularOption()
    call FunctionServiceParticularOption(dboSS)
    Response.Write Transmit
End Sub
%>

<%
'＝＝＝＝＝ 系統功能標題 ＝＝＝＝＝
Sub SubSystemFunctionMenu()
    call FunctionSystemFunctionMenu(dboSF)
    Transmit = Transmit & FrontSpace & "------------------------------"
    Response.Write Transmit
End Sub
%>

<%
'＝＝＝＝＝ 系統功能選項 ＝＝＝＝＝
Sub SubSystemFunctionOption()
    call FunctionSystemFunctionOption(dboSF)
    Response.Write Transmit
End Sub
%>

<%'● = = = = = = = = = = = = = = = Function = = = = = = = = = = = = = = = ●%>

<%
'＝＝＝＝＝ 建議狀態 ＝＝＝＝＝
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
                Transmit = Transmit & FrontSpace & "echo 變更服務名稱: “" & Trim(RS("ServiceView")) & "”" & vbNewLine
                Transmit = Transmit & FrontSpace & "echo     建議設定狀態: " & Trim(RS("ServiceSuggestType")) & vbNewLine
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
'＝＝＝＝＝ 預設狀態 ＝＝＝＝＝
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
                Transmit = Transmit & FrontSpace & "echo 變更服務名稱: “" & Trim(RS("ServiceView")) & "”" & vbNewLine
                Transmit = Transmit & FrontSpace & "echo     預設設定狀態: " & Trim(RS("ServiceDefaultType")) & vbNewLine
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
'＝＝＝＝＝ 個別服務設定(標題) ＝＝＝＝＝
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
                
                '將 > 10 的值改成英文顯示
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
'＝＝＝＝＝ 個別服務設定(選項) ＝＝＝＝＝
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
                Transmit = Transmit & FrontSpace & "    rem ---------- 操作服務資料: " & Trim(RS("ServiceName")) & " ----------" & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceView=" & Trim(RS("ServiceView")) & """ " & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceName=" & Trim(RS("ServiceName")) & """ " & vbNewLine
                Transmit = Transmit & FrontSpace & "    set ""ServiceDescribe=請編輯此檔查看 " & vbNewLine
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
'＝＝＝＝＝ 系統功能標題 ＝＝＝＝＝
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
'＝＝＝＝＝ 系統功能選項 ＝＝＝＝＝
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
