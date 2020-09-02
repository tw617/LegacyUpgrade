<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<!-- #include file="CallMain_ConnectionDataTable.asp" -->
<!-- #include file="CallMainScript.asp" -->
<%
'● = = = = = = = = = = = = = = = 設定宣告 = = = = = = = = = = = = = = = ●
RulerOS=Request.QueryString("RulerOS")	'取得欲搜尋的資料
SubTitle=CStr("" & Request.QueryString("SubTitle"))	'取得欲搜尋的資料
RulerRobotVer="0.3"		'套用模組設定的 模組版本

'● = = = = = = = = = = = = = = = 說　　明 = = = = = = = = = = = = = = = ●
Dim OSName

Select Case CStr((RulerOS & SubTitle))
	Case "Win10.0"
		OSName="Windows 10"
	Case "Win6.1ServerR2"
		OSName="Windows Server 2008 R2"
	Case "Win6.1VM"
		OSName="Windows 7 for Virtual Machine"
	Case "Win6.1"
		OSName="Windows 7"
	Case "Win6.0Server"
		OSName="Windows Server 2008"
	Case "Win6.0"
		OSName="Windows Vista"
	Case "Win5.2Server"
		OSName="Windows Server 2003"
	Case "Win5.1"
		OSName="Windows XP"
	Case "Win5.0"
		OSName="Windows 2000"
End Select

'-------------------- 素材: 說明宣告 --------------------
ResultTime1 = Year(now) & "/" & Right("0" & Month(now), 2) & "/" & Right("0" & Day(now), 2)
ResultTime2 = Right("0" & Hour(now), 2) & ":" & Right("0" & Minute(now), 2) & ":" & Right("0" & Second(now), 2)
%>

<%
'產生下載命令 (1/2)
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=Ruler-System_" & Replace(OSName,Chr("32"),"%20") & ".bat"
Response.ContentType = "text"
Call FunctionDataSource(DataBase)
%>
@echo off
rem ＝＝＝＝＝＝＝＝＝＝　說　　明　＝＝＝＝＝＝＝＝＝＝
echo 請使用系統管理員身分執行
echo 模組名稱: <%=RulerOS & SubTitle%>
echo 選單版本: <%=RulerRobotVer%>
echo 製作日期: <%=ResultTime1 & " " & ResultTime2%>

rem ＝＝＝＝＝＝＝＝＝＝　檢查系統　＝＝＝＝＝＝＝＝＝＝
set "HostOS="
ver|find /i " 10.0">nul && set "HostOS=Win10.0"
ver|find /i " 6.1">nul && set "HostOS=Win6.1"
ver|find /i " 6.0">nul && set "HostOS=Win6.0"
ver|find /i " 5.2">nul && set "HostOS=Win5.2"
ver|find /i " 5.1">nul && set "HostOS=Win5.1"
ver|find /i " 5.0">nul && set "HostOS=Win5.0"

IF "%HostOS%" EQU "" (
		set "Er_Msg_=抱歉，無法支配這個詭異的系統。"
		goto Er_Value
	) ELSE (
		IF "%HostOS%" NEQ "<%=RulerOS%>" (
			set "Er_Msg_=抱歉，此指令檔並不適用於此系統。"
			goto Er_Value
		)
	)

<%
'● = = = = = = = = = = = = = = = 取得數量 = = = = = = = = = = = = = = = ●
	Dim ArrayRowNum
'-------------------- 系統功能 --------------------
	call FunctionGetRowNum(dboSF)
	MaxSystemFunction_Num = ArrayRowNum

'-------------------- 系統服務 --------------------
	call FunctionGetRowNum(dboSS)
	MaxSystemService_Num = ArrayRowNum
%>
Title 系統：<%=OSName%>
:rootMenu
cls
echo 請選擇!!　『<%=OSName%>』
echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
echo  0. 關閉 (*預設)
echo ----------------------------------------
echo  1. 部分系統服務設為 建議狀態 (<%=MaxSystemService_Num%>) (*建議)
echo  2. 部分系統服務設為 預設狀態 (<%=MaxSystemService_Num%>)
echo ----------------------------------------
echo  3. 服務個別設定 (<%=MaxSystemService_Num%>)
echo ----------------------------------------
echo  4. 系統功能 (<%=MaxSystemFunction_Num%>)
echo ----------------------------------------
echo  9. 系統關機/重開機
echo ----------------------------------------

CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
cls
if %ERRORLEVEL% GEQ 36 goto end
if %ERRORLEVEL% GEQ 9 goto Ch_9
if %ERRORLEVEL% GTR 4 goto rootMenu
goto Ch_%ERRORLEVEL%

rem ● = = = = = = = = = = = = = = = 選項部分 = = = = = = = = = = = = = = = ●
:Ch_1
<%FrontSpace="	"%>
	rem ---------- 建議狀態 ----------
<%	call SubStateToSuggestType()%>
	pause
	goto end


:Ch_2
	rem ---------- 預設狀態 ----------
<%	call SubStateToDefaultType()%>
	pause
	goto end


:Ch_3
<%FrontSpace="	echo "%>
	rem ---------- 個別服務設定 ----------
	cls
	echo.
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
	echo ------------------------------
<%	call SubServiceParticularMenu()%>
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 30 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR <%=MaxSystemService_Num%> goto Ch_3
	goto Ch_3-%ERRORLEVEL%

<%FrontSpace="	"%>
<%	call SubServiceParticularOption()%>

	:Ch_3-Menu
		rem ---------- 服務資料(顯示用) ----------
		echo  ●操作服務：%ServiceView%
		rem 檢查目前 類型、狀態
		set "ServiceMomentState=未知"
		set "ServiceMomentType=未知"
		sc query "%ServiceName%"|find /i "RUNNING">nul && set "ServiceMomentState=已啟動"
		sc query "%ServiceName%"|find /i "STOPPED">nul && set "ServiceMomentState=已停止"
		sc qc "%ServiceName%"|find /i "AUTO_START">nul && set "ServiceMomentType=自動"
		sc qc "%ServiceName%"|find /i "DELAYED">nul && set "ServiceMomentType=延遲"
		sc qc "%ServiceName%"|find /i "DEMAND_START">nul && set "ServiceMomentType=手動"
		sc qc "%ServiceName%"|find /i "DISABLED">nul && set "ServiceMomentType=停用"
		echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
		echo  ●顯示名稱：“%ServiceView%”
		echo  ●服務名稱：“%ServiceName%”
		echo  ●服務描述：“%ServiceDescribe%”
		echo  ●預設狀態：“%ServiceDefaultType%”
		echo  ●建議狀態：“%ServiceSuggestType%”
		echo  ●    (自動: auto；延遲: delayed-auto；手動: demand；停用: disabled)
		echo  -----------------------------
		echo  ●目前狀態：“%ServiceMomentState%”
		echo  ●目前類型：“%ServiceMomentType%”
		echo  -----------------------------
		echo  ●備忘紀錄：“%ServiceMemo%”
		echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
		echo   0. 回上一頁 (*預設)
		echo ------------------------------
		echo   1. 常用設定: 設為建議狀態 (*建議)
		echo   2. 常用設定: 設為預設狀態
		echo ------------------------------
		echo   3. 立即執行: 立即啟動 [停用狀態者無法啟動]
		echo   4. 立即執行: 立即停用
		echo ------------------------------
		echo   5. 開機狀態: 自動啟動
		echo   6. 開機狀態: 延遲啟動
		echo   7. 開機狀態: 手動啟動
		echo   8. 開機狀態: 停用狀態
		CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 30 /D 0
		cls
		if %ERRORLEVEL% GTR 8 goto Ch_3
		goto Ch_3-Menu-%ERRORLEVEL%

		:Ch_3-Menu-1
			rem ---------- 設為建議狀態 ----------
			echo 設為建議狀態：%ServiceView%
			sc config "%ServiceName%" start= %ServiceSuggestType%
			goto Ch_3-Menu

		:Ch_3-Menu-2
			rem ---------- 設為預設狀態 ----------
			echo 設為預設狀態：%ServiceView%
			sc config "%ServiceName%" start= %ServiceDefaultType%
			goto Ch_3-Menu

		:Ch_3-Menu-3
			rem ---------- 立即啟動 ----------
			echo 立即啟動：%ServiceView%
			net start %ServiceName%
			goto Ch_3-Menu

		:Ch_3-Menu-4
			rem ---------- 立即停用 ----------
			echo 立即停用：%ServiceView%
			net stop %ServiceName%
			goto Ch_3-Menu

		:Ch_3-Menu-5
			rem ---------- 自動啟動 ----------
			echo 自動啟動：%ServiceView%
			sc config "%ServiceName%" start= auto
			goto Ch_3-Menu

		:Ch_3-Menu-6
			rem ---------- 延遲啟動 ----------
			echo 延遲啟動：%ServiceView%
			sc config "%ServiceName%" start= delayed-auto
			goto Ch_3-Menu

		:Ch_3-Menu-7
			rem ---------- 手動啟動 ----------
			echo 手動啟動：%ServiceView%
			sc config "%ServiceName%" start= demand
			goto Ch_3-Menu

		:Ch_3-Menu-8
			rem ---------- 停用狀態 ----------
			echo 停用狀態：%ServiceView%
			sc config "%ServiceName%" start= disabled
			goto Ch_3-Menu


:Ch_4
<%FrontSpace="	echo "%>
	rem ---------- 系統功能 ----------
	echo.
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
	echo ------------------------------
<%	call SubSystemFunctionMenu()%>
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR <%=MaxSystemFunction_Num%> goto Ch_4
	goto Ch_4-%ERRORLEVEL%

<%FrontSpace="		"%>
<%	call SubSystemFunctionOption()%>


:Ch_9

	rem ---------- 系統關機/重開機 ----------
	echo.
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
	echo ------------------------------
	echo  1. 重開機
	echo  2. 關機
	echo  3. 取消關機
	echo ------------------------------
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR 3 goto Ch_9
	goto Ch_9-%ERRORLEVEL%

		:Ch_9-1
			rem ---------- 重新開機 ----------
			echo ------------------------------
			echo  選擇了“10秒後重開機: %computername%”確定嗎？
			CHOICE /C YNC /n /M "是(Y)、否(N)、取消(C)。
			if %ERRORLEVEL% == 1 shutdown -r -m \\%computername% -t 10 -c "關機中，準備重開機。" -f
			if %ERRORLEVEL% == 2 goto Ch_9-1
			goto end

		:Ch_9-2
			rem ---------- 關機 ----------
			echo ------------------------------
			echo 選擇了“10秒後關機: %computername%”確定嗎？
			CHOICE /C YNC /n /M "是(Y)、否(N)、取消(C)。
			if %ERRORLEVEL% == 1 shutdown -s -m \\%computername% -t 10 -c "關機中" -f
			if %ERRORLEVEL% == 2 goto Ch_9-1
			goto end

		:Ch_9-3
			rem ---------- 取消關機 ----------
			shutdown -a
			goto end


:end
	exit


rem ● = = = = = = = = = = = = = = = 錯誤訊息 = = = = = = = = = = = = = = = ●
:Er_Value
	cls
	echo ┌──────┐
	echo │　錯誤！！　│
	echo └──────┘
	echo Error Message:
	echo %Er_Msg_%
	echo.
	pause
	goto end

<%
DBConnection.close

'產生下載命令 (2/2)
Response.End
%>