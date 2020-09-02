<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<%
'在網頁第一行加入下面這行，注意路徑要改，不含最前面的單引號'
'<!-- #include file="CallMainScript.asp" -->

'宣告通用變數
Dim Transmit, FrontTip, PgFront, SvFront, SvRear, CounterNum, CounterNumView
%>


<%
'＝＝＝＝＝ 取得數量 ＝＝＝＝＝
Function FunctionGetRowNum(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	ArrayRowNum=0

	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			ArrayRowNum = ArrayRowNum +1
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%'● = = = = = = = = = = = = = = = Sub = = = = = = = = = = = = = = = ●%>

<%
'---------- 啟動 基本程式 ----------
Sub SubStartBasicPg()
		FrontTip="echo 啟動基本程式:"
		PgFront="start"

	call FunctionStartProgram(dboBsPg)
	Response.Write Transmit
End Sub
%>
<%
'---------- 終止 基本程式 ----------
Sub SubKillBasicPg()
		FrontTip="echo 終止基本程式:"
		PgFront="taskkill /F /IM"

	call FunctionKillProgram(dboBsPg)
	Response.Write Transmit
End Sub
%>

<%
'---------- 啟動 擴充程式 ----------
Sub SubStartExpandPg()
		FrontTip="echo 啟動擴充程式:"
		PgFront="start"

	call FunctionStartProgram(dboEpPg)
	Response.Write Transmit
End Sub
%>
<%
'---------- 終止 擴充程式 ----------
Sub SubKillExpandPg()
		FrontTip="echo 終止擴充程式:"
		PgFront="taskkill /F /IM"

	call FunctionKillProgram(dboEpPg)
	Response.Write Transmit
End Sub
%>

<%
'---------- 啟動 基本服務 ----------
Sub SubStartBasicSv()
		FrontTip="echo 啟動基本服務:"
		SvFront="net start"
		SvRear=""

	call FunctionService(dboBsSv)
	Response.Write Transmit
End Sub
%>
<%
'---------- 停止 基本服務 ----------
Sub SubKillBasicSv()
		FrontTip="echo 停止基本服務:"
		SvFront="net stop"
		SvRear=""

	call FunctionService(dboBsSv)
	Response.Write Transmit
End Sub
%>

<%
'---------- 啟動 擴充服務 ----------
Sub SubStartExpandSv()
		FrontTip="echo 啟動擴充服務:"
		SvFront="net start"
		SvRear=""

	call FunctionService(dboEpSv)
	Response.Write Transmit
End Sub
%>
<%
'---------- 停止 擴充服務 ----------
Sub SubKillExpandSv()
		FrontTip="echo 停止擴充服務:"
		SvFront="net stop"
		SvRear=""

	call FunctionService(dboEpSv)
	Response.Write Transmit
End Sub
%>


<%
'---------- 停用 (基本/擴充)服務 ----------
Sub SubStopSv(OpenDboName)
		FrontTip="echo 停用服務:"
		SvFront="sc config"
		SvRear="start= disabled"

	call FunctionService(OpenDboName)
	Response.Write Transmit
End Sub
%>


<%
'---------- (基本/擴充)服務 改為建議狀態 ----------
Sub SubStateToSuggestType(OpenDboName)
		FrontTip="echo "
		SvFront="sc config"
		SvRear=""

	call FunctionSuggestType(OpenDboName)
	Response.Write Transmit
End Sub
%>

<%
'---------- (基本/擴充)服務 改為預設狀態 ----------
Sub SubStateToDefaultType(OpenDboName)
		FrontTip="echo "
		SvFront="sc config"
		SvRear=""

	call FunctionDefaultType(OpenDboName)
	Response.Write Transmit
End Sub
%>

<%
'＝＝＝＝＝ 個別服務設定(目錄) ＝＝＝＝＝
Sub SubServiceParticularMenu()
	Transmit=""
	CounterNum=0
	CounterNumView=0

	Transmit = Transmit & FrontSpace & "------- 基　本　服　務 -------" & vbNewLine
	call FunctionServiceParticularMenu(dboBsSv)
	Transmit = Transmit & FrontSpace & "------- 擴　充　服　務 -------" & vbNewLine
	call FunctionServiceParticularMenu(dboEpSv)
	Transmit = Transmit & FrontSpace & "------------------------------" & vbNewLine
	Response.Write Transmit
End Sub
%>

<%
'＝＝＝＝＝ 個別服務設定(選項) ＝＝＝＝＝
Sub SubServiceParticularOption()
	Transmit=""
	CounterNum=0

	call FunctionServiceParticularOption(dboBsSv)
	call FunctionServiceParticularOption(dboEpSv)
	Response.Write Transmit
End Sub
%>

<%'● = = = = = = = = = = = = = = = Function = = = = = = = = = = = = = = = ●%>

<%
'功能: 啟動/停止 服務
Function FunctionService(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Transmit=""
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			Transmit = Transmit & FrontSpace & FrontTip & " " & Trim(RS("ServiceView")) & vbNewLine
			Transmit = Transmit & FrontSpace & SvFront & " """ & Trim(RS("ServiceName")) & """ " & SvRear & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%
'功能: 執行程式
Function FunctionStartProgram(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Transmit=""
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			Transmit = Transmit & FrontSpace & FrontTip & " " & Trim(RS("SoftMainFile")) & vbNewLine
			Transmit = Transmit & FrontSpace & PgFront & " """" """ & Trim(RS("SoftPath")) & "\" & Trim(RS("SoftMainFile")) & """ " & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%
'功能: 終止程式
Function FunctionKillProgram(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Transmit=""
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			Transmit = Transmit & FrontSpace & FrontTip & " " & Trim(RS("SoftMainFile")) & vbNewLine
			Transmit = Transmit & FrontSpace & PgFront & " """ & Trim(RS("SoftMainFile")) & """ " & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>





<%
'功能: (基本/擴充)服務 改為建議狀態
Function FunctionSuggestType(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Transmit=""
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			Transmit = Transmit & FrontSpace & FrontTip & "更改服務: " & Trim(RS("ServiceView")) & vbNewLine
			Transmit = Transmit & FrontSpace & FrontTip & "    建議狀態: " & Trim(RS("ServiceSuggestType")) & vbNewLine
			Transmit = Transmit & FrontSpace & SvFront & " """ & Trim(RS("ServiceName")) & """ start= " & Trim(RS("ServiceSuggestType")) & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%
'功能: (基本/擴充)服務 改為預設狀態
Function FunctionDefaultType(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Transmit=""
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			Transmit = Transmit & FrontSpace & FrontTip & "更改服務: " & Trim(RS("ServiceView")) & vbNewLine
			Transmit = Transmit & FrontSpace & FrontTip & "    預設狀態: " & Trim(RS("ServiceDefaultType")) & vbNewLine
			Transmit = Transmit & FrontSpace & SvFront & " """ & Trim(RS("ServiceName")) & """ start= " & Trim(RS("ServiceDefaultType")) & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%
'＝＝＝＝＝ 個別服務設定(目錄) ＝＝＝＝＝
Function FunctionServiceParticularMenu(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			CounterNum = CounterNum +1

			'將 > 10 的值改成英文顯示
			if CounterNum < 10 then
				CounterNumView=Chr(CounterNum+48)
				ELSE
				CounterNumView=Chr(CounterNum+97-10)
				end if

			Transmit = Transmit & FrontSpace & " " & CounterNumView & ". " & Trim(RS("ServiceName")) & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%
'＝＝＝＝＝ 個別服務設定(選項) ＝＝＝＝＝
Function FunctionServiceParticularOption(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			CounterNum = CounterNum +1
			Transmit = Transmit & FrontSpace & ":Ch_3-" & CounterNum & vbNewLine
			Transmit = Transmit & FrontSpace & "	rem ---------- 操作服務資料: " & Trim(RS("ServiceName")) & " ----------" & vbNewLine
			Transmit = Transmit & FrontSpace & "	set ""ServiceView=" & Trim(RS("ServiceView")) & """ " & vbNewLine
			Transmit = Transmit & FrontSpace & "	set ""ServiceName=" & Trim(RS("ServiceName")) & """ " & vbNewLine
			Transmit = Transmit & FrontSpace & "	set ""ServiceDefaultType=" & Trim(RS("ServiceDefaultType")) & """ " & vbNewLine
			Transmit = Transmit & FrontSpace & "	set ""ServiceSuggestType=" & Trim(RS("ServiceSuggestType")) & """ " & vbNewLine
			Transmit = Transmit & FrontSpace & "	goto Ch_3-Menu" & vbNewLine
			Transmit = Transmit & FrontSpace & "	goto end" & vbNewLine & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>