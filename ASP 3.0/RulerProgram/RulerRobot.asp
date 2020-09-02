<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<!-- #include file="CallMain_ConnectionDataTable.asp" -->
<!-- #include file="CallMainScript.asp" -->
<!-- #include file="CallSecondaryScript.asp" -->
<%
'● = = = = = = = = = = = = = = = 設定宣告 = = = = = = = = = = = = = = = ●
Brief=Request.QueryString("Brief")	'取得欲搜尋的資料
RulerRobotVer="0.2"			'套用模組設定的 模組版本
'● = = = = = = = = = = = = = = = 說　　明 = = = = = = = = = = = = = = = ●

'-------------------- 素材: 說明宣告 --------------------
Call FunctionDataSource("RulerProgram.xls")
RS.Open "[DB_Index$]",DBConnection,3,1
RS.MoveFirst
RS.Find="[SoftBriefName]='" & Brief & "'"			'尋找目標資料
	SoftBriefName = Trim(RS("SoftBriefName"))
	SoftFullName = Trim(RS("SoftFullName"))
	LastUpdate = Trim(RS("LastUpdate"))
	ResultTime1 = Year(now) & "/" & Right("0" & Month(now), 2) & "/" & Right("0" & Day(now), 2)
	ResultTime2 = Right("0" & Hour(now), 2) & ":" & Right("0" & Minute(now), 2) & ":" & Right("0" & Second(now), 2)
	MainFilePath = Trim(RS("MainSoftPath"))
RS.Close
DBConnection.Close
%>

<%
'產生下載命令 (1/2)
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=Ruler-Program_" & Replace(SoftFullName,Chr("32"),"%20") & ".bat"
Response.ContentType = "text"
%>
@echo off

rem ＝＝＝＝＝＝＝＝＝＝　說　　明　＝＝＝＝＝＝＝＝＝＝
echo 模組名稱: <%=SoftBriefName%>
echo 模組版本: <%=LastUpdate%>
echo 選單版本: <%=RulerRobotVer%>
echo 製作日期: <%=ResultTime1 & " " & ResultTime2%>

rem ＝＝＝＝＝＝＝＝＝＝　系統位元　＝＝＝＝＝＝＝＝＝＝
set pf=%ProgramFiles%
set cpf=%CommonProgramFiles%


if %PROCESSOR_ARCHITECTURE% == x86 (
	echo 這是 x86 系統
	set pf32=%ProgramFiles%
	set cpf32=%CommonProgramFiles%
	x86 set RegSoftPath32=Software
	)


if %PROCESSOR_ARCHITECTURE% == AMD64 (
	echo 這是 x64 系統
	set pf32=%ProgramFiles(x86)%
	set pf64=%ProgramFiles%
	set cpf32=%CommonProgramFiles(x86)%
	set cpf64=%CommonProgramFiles%
	set RegSoftPath32=Software\Wow6432Node
	set RegSoftPath64=Software
	)

rem ＝＝＝＝＝＝＝＝＝＝　檢查檔案　＝＝＝＝＝＝＝＝＝＝
IF NOT EXIST "<%=MainFilePath%>" (
	set "Er_Msg_=找不到檔案：<%=MainFilePath%>"
	goto Er_Value
	)
<%
'● = = = = = = = = = = = = = = = 取得數量 = = = = = = = = = = = = = = = ●
	Dim ArrayRowNum
'-------------------- Basic Program --------------------
	call FunctionGetRowNum(dboBsPg)
	MaxBsPg_Num = ArrayRowNum

'-------------------- Basic Service --------------------
	call FunctionGetRowNum(dboBsSv)
	MaxBsSv_Num = ArrayRowNum

'-------------------- Expand Program --------------------
	call FunctionGetRowNum(dboEpPg)
	MaxEpPg_Num = ArrayRowNum

'-------------------- Expand Service --------------------
	call FunctionGetRowNum(dboEpSv)
	MaxEpSv_Num = ArrayRowNum

'-------------------- Extend Rename --------------------
	call FunctionGetRowNum(dboEtRn)
	MaxEtMd_Num = ArrayRowNum

'-------------------- Extend CommandRuler --------------------
	call FunctionGetRowNum(dboEtCR)
	MaxEtCR_Num = ArrayRowNum

%>
Title 設定服務：<%=SoftBriefName%>
:rootMenu
cls
echo 請選擇!!　『<%=SoftFullName%>』
echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
echo  0. 關閉 (*預設)
echo ----------------------------------------
echo  1. 啟動『<%=SoftFullName%>』的 程式(<%=MaxBsPg_Num+MaxEpPg_Num%>) 服務(<%=MaxBsSv_Num+MaxEpSv_Num%>)
echo  2. 關閉『<%=SoftFullName%>』的 程式(<%=MaxEpPg_Num+MaxBsPg_Num%>) 服務(<%=MaxEpSv_Num+MaxBsSv_Num%>)
echo ----------------------------------------
echo  3. 管理『<%=SoftFullName%>』個別的 服務(<%=MaxBsSv_Num+MaxEpSv_Num%>)
echo  4. 管理『<%=SoftFullName%>』全部的 服務(<%=MaxBsSv_Num+MaxEpSv_Num%>)
echo ----------------------------------------
echo  5. 檔案更名管理(<%=MaxEtMd_Num%>)
echo  6. 延伸控制管理(<%=MaxEtCR_Num%>)

CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
cls
if %ERRORLEVEL% GEQ 36 goto end
if %ERRORLEVEL% GTR 6 goto rootMenu
goto Ch_%ERRORLEVEL%

rem ● = = = = = = = = = = = = = = = 選項部分 = = = = = = = = = = = = = = = ●
<%FrontSpace="		"%>
:Ch_1
	echo 請選擇欲“啟動的程式與服務”
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
	echo ----------------------------------------
	echo  1. 基本 程式(<%=MaxBsPg_Num%>) 與 服務(<%=MaxBsSv_Num%>)
	echo  2. 全部 程式(<%=MaxBsPg_Num+MaxEpPg_Num%>) 與 服務(<%=MaxBsSv_Num+MaxEpSv_Num%>)
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	set "Option=%ERRORLEVEL%"
	if %Option% GTR 2 goto rootMenu

	rem ---------- 啟動 基本服務 ----------
<%		call SubStartBasicSv()%>
	rem ---------- 啟動 基本程式 ----------
<%		call SubStartBasicPg()%>

	if "%Option%"=="1" (pause && goto end)
	rem ---------- 啟動 擴充服務 ----------
<%		call SubStartExpandSv()%>
	rem ---------- 啟動 擴充程式 ----------
<%		call SubStartExpandPg()%>
	pause
	goto end


:Ch_2
	echo 請選擇欲“關閉的程式與服務”
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
	echo ----------------------------------------
	echo  1. 全部 程式(<%=MaxBsPg_Num+MaxEpPg_Num%>) 與 服務(<%=MaxBsSv_Num+MaxEpSv_Num%>)
	echo  2. 擴充 程式(<%=MaxEpPg_Num%>) 與 服務(<%=MaxEpSv_Num%>)
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	set "Option=%ERRORLEVEL%"
	if %Option% GTR 2 goto rootMenu

	rem ---------- 終止 擴充程式 ----------
<%		call SubKillExpandPg()%>
	rem ---------- 停止 擴充服務 ----------
<%		call SubKillExpandSv()%>

	if "%Option%"=="2" (pause && goto end)
	rem ---------- 終止 基本程式 ----------
<%		call SubKillBasicPg()%>
	rem ---------- 停止 基本服務 ----------
<%		call SubKillBasicSv()%>
	pause
	goto end


:Ch_3
<%FrontSpace="	echo "%>
	echo.
	echo 管理個別服務: 『<%=SoftFullName%>』
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
<%	call SubServiceParticularMenu()%>
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 30 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR <%=MaxBsSv_Num+MaxEpSv_Num%> goto Ch_3
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
		echo  ●預設狀態：“%ServiceDefaultType%”
		echo  ●建議狀態：“%ServiceSuggestType%”
		echo  ●    (自動: auto；延遲: delayed-auto；手動: demand；停用: disabled)
		echo  -----------------------------
		echo  ●目前狀態：“%ServiceMomentState%”
		echo  ●目前類型：“%ServiceMomentType%”
		echo ●○●○●○●○●○●○●○●
		echo  0. 回上一頁 (*預設)
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
<%FrontSpace="			"%>
	cls
	echo.
	echo 管理全部服務:『<%=SoftFullName%>』
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
	echo ------------------------------
	echo 　　　　開機 狀態 設定　　　　
	echo   1. 全部設為建議狀態(<%=MaxBsSv_Num+MaxEpSv_Num%>) (*建議)
	echo   2. 全部設為預設狀態(<%=MaxBsSv_Num+MaxEpSv_Num%>)
	echo   3. (開機時)停用 全部服務(<%=MaxBsSv_Num+MaxEpSv_Num%>)
	echo ------------------------------
	echo 　　　　立　即　處　理　　　　
	echo   4. (立即)啟動 服務(<%=MaxBsSv_Num+MaxEpSv_Num%>) [停用狀態者無法啟動]
	echo   5. (立即)終止 服務(<%=MaxEpSv_Num+MaxBsSv_Num%>)

	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR 5 goto Ch_4
	goto Ch_4-%ERRORLEVEL%

	:Ch_4-1
		rem ========== 全部設為 建議狀態 ==========
		rem ---------- 基本服務 改為建議狀態 ----------
<%			call SubStateToSuggestType(dboBsSv)%>
		rem ---------- 擴充服務 改為建議狀態 ----------
<%			call SubStateToSuggestType(dboEpSv)%>
		pause
		goto Ch_4

	:Ch_4-2
		rem ========== 全部設為 預設狀態 ==========
		rem ---------- 基本服務 改為預設狀態 ----------
<%			call SubStateToDefaultType(dboBsSv)%>
		rem ---------- 擴充服務 改為預設狀態 ----------
<%			call SubStateToDefaultType(dboEpSv)%>
		pause
		goto Ch_4

	:Ch_4-3
		rem ========== (開機時)停用 全部服務 ==========
		rem ---------- 停用 基本服務 ----------
<%			call SubStopSv(dboBsSv)%>
		rem ---------- 停用 擴充服務 ----------
<%			call SubStopSv(dboEpSv)%>
		pause
		goto Ch_4


<%FrontSpace="		"%>
	:Ch_4-4
		echo 請選擇欲“啟動的服務”
		echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
		echo  0. 回上層目錄 (*預設)
		echo ----------------------------------------
		echo  1. 基本 服務(<%=MaxBsSv_Num%>)
		echo  2. 全部 服務(<%=MaxBsSv_Num+MaxEpSv_Num%>)
		echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
		CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
		set "Option=%ERRORLEVEL%"
		if %Option% GTR 2 goto rootMenu

		rem ---------- 啟動 基本服務 ----------
<%		call SubStartBasicSv()%>

		if "%Option%"=="1" (pause && goto end)
		rem ---------- 啟動 擴充服務 ----------
<%		call SubStartExpandSv()%>
		pause
		goto end

	:Ch_4-5
		echo 請選擇欲“關閉的服務”
		echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
		echo  0. 回上層目錄 (*預設)
		echo ----------------------------------------
		echo  1. 全部 服務(<%=MaxBsSv_Num+MaxEpSv_Num%>)
		echo  2. 擴充 服務(<%=MaxEpSv_Num%>)
		echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
		CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
		set "Option=%ERRORLEVEL%"
		if %Option% GTR 2 goto rootMenu

		rem ---------- 停止 擴充服務 ----------
<%		call SubKillExpandSv()%>

		if "%Option%"=="2" (pause && goto end)
		rem ---------- 停止 基本服務 ----------
<%		call SubKillBasicSv()%>
		pause
		goto end


:Ch_5
	echo.
	echo 檔案更名管理:『<%=SoftFullName%>』
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
	echo ------------------------------
	echo   1. 更名累贅的 程式(<%=MaxEtMd_Num%>) (請先關閉該 程式/服務) (*建議)
	echo   2. 復原已更名的 程式(<%=MaxEtMd_Num%>)
	echo ------------------------------

	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	cls
	if %ERRORLEVEL% GTR 2 goto rootMenu
	goto Ch_5-%ERRORLEVEL%

	:Ch_5-1
		rem ---------- 更名 累贅程式 ----------
<%			call SubRename()%>
		pause
		goto rootMenu


	:Ch_5-2
		rem ---------- 復原更名 累贅程式 ----------
<%			call SubRestoreRename()%>
		pause
		goto rootMenu


:Ch_6
<%FrontSpace="	echo "%>
	echo.
	echo 延伸控制管理:『<%=SoftFullName%>』
	echo ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
	echo  0. 回上層目錄 (*預設)
	echo ------------------------------
<%	call SubCommandRulerMenu()%>
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 30 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR <%=MaxEtCR_Num%> goto Ch_6
	goto Ch_6-%ERRORLEVEL%

<%FrontSpace="	"%>
<%	call SubCommandRulerOption()%>

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
'產生下載命令 (2/2)
Response.End
%>