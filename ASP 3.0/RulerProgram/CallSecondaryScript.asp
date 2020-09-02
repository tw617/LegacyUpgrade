<%
'在網頁第一行加入下面這行，注意路徑要改，不含最前面的單引號'
'<!-- #include file="CallSecondaryScript.asp" -->

'宣告通用變數
Dim FileSourceName, FileTargetName
%>

<%'● = = = = = = = = = = = = = = = Sub = = = = = = = = = = = = = = = ●%>

<%
'＝＝＝＝＝ 更名 累贅程式 ＝＝＝＝＝
Sub SubRename()
		FrontTip="echo 更名: "
		FileSourceName="SoftFileOriginalName"
		FileTargetName="SoftFileNewName"

	call FunctionRename(dboEtRn)
	Response.Write Transmit
End Sub
%>


<%
'＝＝＝＝＝ 復原更名 累贅程式 ＝＝＝＝＝
Sub SubRestoreRename()
		FrontTip="echo 復原更名: "
		FileSourceName="SoftFileNewName"
		FileTargetName="SoftFileOriginalName"

	call FunctionRename(dboEtRn)
	Response.Write Transmit
End Sub
%>


<%
'＝＝＝＝＝ 延伸控制(標題) ＝＝＝＝＝
Sub SubCommandRulerMenu()
	Transmit=""
	CounterNum=0
	CounterNumView=0

	call FunctionCommandRulerMenu(dboEtCR)
	Transmit = Transmit & FrontSpace & "------------------------------"
	Response.Write Transmit
End Sub
%>


<%
'＝＝＝＝＝ 延伸控制(選項) ＝＝＝＝＝
Sub SubCommandRulerOption()
	Transmit=""
	CounterNum=0

	call FunctionCommandRulerOption(dboEtCR)
	Response.Write Transmit
End Sub
%>

<%'● = = = = = = = = = = = = = = = Function = = = = = = = = = = = = = = = ●%>

<%
'功能: 更名 累贅程式
Function FunctionRename(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Transmit=""
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			Transmit = Transmit & FrontSpace & FrontTip & " " & Trim(RS(FileSourceName)) & vbNewLine
			Transmit = Transmit & FrontSpace & "ren """ & Trim(RS("ProgramPath")) & "\" & Trim(RS(FileSourceName)) & """ """ & Trim(RS(FileTargetName)) & """ " & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%
'功能: 延伸控制(目錄)
Function FunctionCommandRulerMenu(OpenDboName)
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

			Transmit = Transmit & FrontSpace & " " & CounterNumView & ". " & Trim(RS("CommandTitle")) & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>


<%
'功能: 延伸控制(選項)
Function FunctionCommandRulerOption(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			CounterNum = CounterNum +1
			CommandRulerData = ""

			Transmit = Transmit & FrontSpace & ":Ch_6-" & CounterNum & vbNewLine
			Transmit = Transmit & FrontSpace & "	rem 操作: " & Trim(RS("CommandTitle")) & vbNewLine
			CommandRulerData = Trim(RS("CommandRuler")) & " " & Trim(RS("CommandOperation")) & " """ & Trim(RS("CommandPath")) & """ " & Trim(RS("CommandArgument1")) & " """ & Trim(RS("CommandTarget")) & """ " & Trim(RS("CommandArgument2")) & " "
			Transmit = Transmit & FrontSpace & "	" & CommandRulerData & vbNewLine
			Transmit = Transmit & FrontSpace & "	goto Ch_6" & vbNewLine & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>