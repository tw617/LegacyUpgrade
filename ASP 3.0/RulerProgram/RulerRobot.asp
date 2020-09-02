<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<!-- #include file="CallMain_ConnectionDataTable.asp" -->
<!-- #include file="CallMainScript.asp" -->
<!-- #include file="CallSecondaryScript.asp" -->
<%
'�� = = = = = = = = = = = = = = = �]�w�ŧi = = = = = = = = = = = = = = = ��
Brief=Request.QueryString("Brief")	'���o���j�M�����
RulerRobotVer="0.2"			'�M�μҲճ]�w�� �Ҳժ���
'�� = = = = = = = = = = = = = = = ���@�@�� = = = = = = = = = = = = = = = ��

'-------------------- ����: �����ŧi --------------------
Call FunctionDataSource("RulerProgram.xls")
RS.Open "[DB_Index$]",DBConnection,3,1
RS.MoveFirst
RS.Find="[SoftBriefName]='" & Brief & "'"			'�M��ؼи��
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
'���ͤU���R�O (1/2)
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=Ruler-Program_" & Replace(SoftFullName,Chr("32"),"%20") & ".bat"
Response.ContentType = "text"
%>
@echo off

rem �סססססססססס@���@�@���@�ססססססססס�
echo �ҲզW��: <%=SoftBriefName%>
echo �Ҳժ���: <%=LastUpdate%>
echo ��檩��: <%=RulerRobotVer%>
echo �s�@���: <%=ResultTime1 & " " & ResultTime2%>

rem �סססססססססס@�t�Φ줸�@�ססססססססס�
set pf=%ProgramFiles%
set cpf=%CommonProgramFiles%


if %PROCESSOR_ARCHITECTURE% == x86 (
	echo �o�O x86 �t��
	set pf32=%ProgramFiles%
	set cpf32=%CommonProgramFiles%
	x86 set RegSoftPath32=Software
	)


if %PROCESSOR_ARCHITECTURE% == AMD64 (
	echo �o�O x64 �t��
	set pf32=%ProgramFiles(x86)%
	set pf64=%ProgramFiles%
	set cpf32=%CommonProgramFiles(x86)%
	set cpf64=%CommonProgramFiles%
	set RegSoftPath32=Software\Wow6432Node
	set RegSoftPath64=Software
	)

rem �סססססססססס@�ˬd�ɮס@�ססססססססס�
IF NOT EXIST "<%=MainFilePath%>" (
	set "Er_Msg_=�䤣���ɮסG<%=MainFilePath%>"
	goto Er_Value
	)
<%
'�� = = = = = = = = = = = = = = = ���o�ƶq = = = = = = = = = = = = = = = ��
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
Title �]�w�A�ȡG<%=SoftBriefName%>
:rootMenu
cls
echo �п��!!�@�y<%=SoftFullName%>�z
echo �ססססססססססססססססססס�
echo  0. ���� (*�w�])
echo ----------------------------------------
echo  1. �Ұʡy<%=SoftFullName%>�z�� �{��(<%=MaxBsPg_Num+MaxEpPg_Num%>) �A��(<%=MaxBsSv_Num+MaxEpSv_Num%>)
echo  2. �����y<%=SoftFullName%>�z�� �{��(<%=MaxEpPg_Num+MaxBsPg_Num%>) �A��(<%=MaxEpSv_Num+MaxBsSv_Num%>)
echo ----------------------------------------
echo  3. �޲z�y<%=SoftFullName%>�z�ӧO�� �A��(<%=MaxBsSv_Num+MaxEpSv_Num%>)
echo  4. �޲z�y<%=SoftFullName%>�z������ �A��(<%=MaxBsSv_Num+MaxEpSv_Num%>)
echo ----------------------------------------
echo  5. �ɮק�W�޲z(<%=MaxEtMd_Num%>)
echo  6. ��������޲z(<%=MaxEtCR_Num%>)

CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
cls
if %ERRORLEVEL% GEQ 36 goto end
if %ERRORLEVEL% GTR 6 goto rootMenu
goto Ch_%ERRORLEVEL%

rem �� = = = = = = = = = = = = = = = �ﶵ���� = = = = = = = = = = = = = = = ��
<%FrontSpace="		"%>
:Ch_1
	echo �п�ܱ����Ұʪ��{���P�A�ȡ�
	echo �ססססססססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
	echo ----------------------------------------
	echo  1. �� �{��(<%=MaxBsPg_Num%>) �P �A��(<%=MaxBsSv_Num%>)
	echo  2. ���� �{��(<%=MaxBsPg_Num+MaxEpPg_Num%>) �P �A��(<%=MaxBsSv_Num+MaxEpSv_Num%>)
	echo �ססססססססססססססססססס�
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	set "Option=%ERRORLEVEL%"
	if %Option% GTR 2 goto rootMenu

	rem ---------- �Ұ� �򥻪A�� ----------
<%		call SubStartBasicSv()%>
	rem ---------- �Ұ� �򥻵{�� ----------
<%		call SubStartBasicPg()%>

	if "%Option%"=="1" (pause && goto end)
	rem ---------- �Ұ� �X�R�A�� ----------
<%		call SubStartExpandSv()%>
	rem ---------- �Ұ� �X�R�{�� ----------
<%		call SubStartExpandPg()%>
	pause
	goto end


:Ch_2
	echo �п�ܱ����������{���P�A�ȡ�
	echo �ססססססססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
	echo ----------------------------------------
	echo  1. ���� �{��(<%=MaxBsPg_Num+MaxEpPg_Num%>) �P �A��(<%=MaxBsSv_Num+MaxEpSv_Num%>)
	echo  2. �X�R �{��(<%=MaxEpPg_Num%>) �P �A��(<%=MaxEpSv_Num%>)
	echo �ססססססססססססססססססס�
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	set "Option=%ERRORLEVEL%"
	if %Option% GTR 2 goto rootMenu

	rem ---------- �פ� �X�R�{�� ----------
<%		call SubKillExpandPg()%>
	rem ---------- ���� �X�R�A�� ----------
<%		call SubKillExpandSv()%>

	if "%Option%"=="2" (pause && goto end)
	rem ---------- �פ� �򥻵{�� ----------
<%		call SubKillBasicPg()%>
	rem ---------- ���� �򥻪A�� ----------
<%		call SubKillBasicSv()%>
	pause
	goto end


:Ch_3
<%FrontSpace="	echo "%>
	echo.
	echo �޲z�ӧO�A��: �y<%=SoftFullName%>�z
	echo �סססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
<%	call SubServiceParticularMenu()%>
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 30 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR <%=MaxBsSv_Num+MaxEpSv_Num%> goto Ch_3
	goto Ch_3-%ERRORLEVEL%

<%FrontSpace="	"%>
<%	call SubServiceParticularOption()%>

	:Ch_3-Menu
		rem ---------- �A�ȸ��(��ܥ�) ----------
		echo  ���ާ@�A�ȡG%ServiceView%
		rem �ˬd�ثe �����B���A
		set "ServiceMomentState=����"
		set "ServiceMomentType=����"
		sc query "%ServiceName%"|find /i "RUNNING">nul && set "ServiceMomentState=�w�Ұ�"
		sc query "%ServiceName%"|find /i "STOPPED">nul && set "ServiceMomentState=�w����"
		sc qc "%ServiceName%"|find /i "AUTO_START">nul && set "ServiceMomentType=�۰�"
		sc qc "%ServiceName%"|find /i "DELAYED">nul && set "ServiceMomentType=����"
		sc qc "%ServiceName%"|find /i "DEMAND_START">nul && set "ServiceMomentType=���"
		sc qc "%ServiceName%"|find /i "DISABLED">nul && set "ServiceMomentType=����"
		echo �סססססססססססססס�
		echo  ����ܦW�١G��%ServiceView%��
		echo  ���A�ȦW�١G��%ServiceName%��
		echo  ���w�]���A�G��%ServiceDefaultType%��
		echo  ����ĳ���A�G��%ServiceSuggestType%��
		echo  ��    (�۰�: auto�F����: delayed-auto�F���: demand�F����: disabled)
		echo  -----------------------------
		echo  ���ثe���A�G��%ServiceMomentState%��
		echo  ���ثe�����G��%ServiceMomentType%��
		echo ������������������������������
		echo  0. �^�W�@�� (*�w�])
		echo ------------------------------
		echo   1. �`�γ]�w: �]����ĳ���A (*��ĳ)
		echo   2. �`�γ]�w: �]���w�]���A
		echo ------------------------------
		echo   3. �ߧY����: �ߧY�Ұ� [���Ϊ��A�̵L�k�Ұ�]
		echo   4. �ߧY����: �ߧY����
		echo ------------------------------
		echo   5. �}�����A: �۰ʱҰ�
		echo   6. �}�����A: ����Ұ�
		echo   7. �}�����A: ��ʱҰ�
		echo   8. �}�����A: ���Ϊ��A
		CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 30 /D 0
		cls
		if %ERRORLEVEL% GTR 8 goto Ch_3
		goto Ch_3-Menu-%ERRORLEVEL%

		:Ch_3-Menu-1
			rem ---------- �]����ĳ���A ----------
			echo �]����ĳ���A�G%ServiceView%
			sc config "%ServiceName%" start= %ServiceSuggestType%
			goto Ch_3-Menu

		:Ch_3-Menu-2
			rem ---------- �]���w�]���A ----------
			echo �]���w�]���A�G%ServiceView%
			sc config "%ServiceName%" start= %ServiceDefaultType%
			goto Ch_3-Menu

		:Ch_3-Menu-3
			rem ---------- �ߧY�Ұ� ----------
			echo �ߧY�ҰʡG%ServiceView%
			net start %ServiceName%
			goto Ch_3-Menu

		:Ch_3-Menu-4
			rem ---------- �ߧY���� ----------
			echo �ߧY���ΡG%ServiceView%
			net stop %ServiceName%
			goto Ch_3-Menu

		:Ch_3-Menu-5
			rem ---------- �۰ʱҰ� ----------
			echo �۰ʱҰʡG%ServiceView%
			sc config "%ServiceName%" start= auto
			goto Ch_3-Menu

		:Ch_3-Menu-6
			rem ---------- ����Ұ� ----------
			echo ����ҰʡG%ServiceView%
			sc config "%ServiceName%" start= delayed-auto
			goto Ch_3-Menu

		:Ch_3-Menu-7
			rem ---------- ��ʱҰ� ----------
			echo ��ʱҰʡG%ServiceView%
			sc config "%ServiceName%" start= demand
			goto Ch_3-Menu

		:Ch_3-Menu-8
			rem ---------- ���Ϊ��A ----------
			echo ���Ϊ��A�G%ServiceView%
			sc config "%ServiceName%" start= disabled
			goto Ch_3-Menu


:Ch_4
<%FrontSpace="			"%>
	cls
	echo.
	echo �޲z�����A��:�y<%=SoftFullName%>�z
	echo �סססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
	echo ------------------------------
	echo �@�@�@�@�}�� ���A �]�w�@�@�@�@
	echo   1. �����]����ĳ���A(<%=MaxBsSv_Num+MaxEpSv_Num%>) (*��ĳ)
	echo   2. �����]���w�]���A(<%=MaxBsSv_Num+MaxEpSv_Num%>)
	echo   3. (�}����)���� �����A��(<%=MaxBsSv_Num+MaxEpSv_Num%>)
	echo ------------------------------
	echo �@�@�@�@�ߡ@�Y�@�B�@�z�@�@�@�@
	echo   4. (�ߧY)�Ұ� �A��(<%=MaxBsSv_Num+MaxEpSv_Num%>) [���Ϊ��A�̵L�k�Ұ�]
	echo   5. (�ߧY)�פ� �A��(<%=MaxEpSv_Num+MaxBsSv_Num%>)

	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR 5 goto Ch_4
	goto Ch_4-%ERRORLEVEL%

	:Ch_4-1
		rem ========== �����]�� ��ĳ���A ==========
		rem ---------- �򥻪A�� �אּ��ĳ���A ----------
<%			call SubStateToSuggestType(dboBsSv)%>
		rem ---------- �X�R�A�� �אּ��ĳ���A ----------
<%			call SubStateToSuggestType(dboEpSv)%>
		pause
		goto Ch_4

	:Ch_4-2
		rem ========== �����]�� �w�]���A ==========
		rem ---------- �򥻪A�� �אּ�w�]���A ----------
<%			call SubStateToDefaultType(dboBsSv)%>
		rem ---------- �X�R�A�� �אּ�w�]���A ----------
<%			call SubStateToDefaultType(dboEpSv)%>
		pause
		goto Ch_4

	:Ch_4-3
		rem ========== (�}����)���� �����A�� ==========
		rem ---------- ���� �򥻪A�� ----------
<%			call SubStopSv(dboBsSv)%>
		rem ---------- ���� �X�R�A�� ----------
<%			call SubStopSv(dboEpSv)%>
		pause
		goto Ch_4


<%FrontSpace="		"%>
	:Ch_4-4
		echo �п�ܱ����Ұʪ��A�ȡ�
		echo �ססססססססססססססססססס�
		echo  0. �^�W�h�ؿ� (*�w�])
		echo ----------------------------------------
		echo  1. �� �A��(<%=MaxBsSv_Num%>)
		echo  2. ���� �A��(<%=MaxBsSv_Num+MaxEpSv_Num%>)
		echo �ססססססססססססססססססס�
		CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
		set "Option=%ERRORLEVEL%"
		if %Option% GTR 2 goto rootMenu

		rem ---------- �Ұ� �򥻪A�� ----------
<%		call SubStartBasicSv()%>

		if "%Option%"=="1" (pause && goto end)
		rem ---------- �Ұ� �X�R�A�� ----------
<%		call SubStartExpandSv()%>
		pause
		goto end

	:Ch_4-5
		echo �п�ܱ����������A�ȡ�
		echo �ססססססססססססססססססס�
		echo  0. �^�W�h�ؿ� (*�w�])
		echo ----------------------------------------
		echo  1. ���� �A��(<%=MaxBsSv_Num+MaxEpSv_Num%>)
		echo  2. �X�R �A��(<%=MaxEpSv_Num%>)
		echo �ססססססססססססססססססס�
		CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
		set "Option=%ERRORLEVEL%"
		if %Option% GTR 2 goto rootMenu

		rem ---------- ���� �X�R�A�� ----------
<%		call SubKillExpandSv()%>

		if "%Option%"=="2" (pause && goto end)
		rem ---------- ���� �򥻪A�� ----------
<%		call SubKillBasicSv()%>
		pause
		goto end


:Ch_5
	echo.
	echo �ɮק�W�޲z:�y<%=SoftFullName%>�z
	echo �סססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
	echo ------------------------------
	echo   1. ��W���ت� �{��(<%=MaxEtMd_Num%>) (�Х������� �{��/�A��) (*��ĳ)
	echo   2. �_��w��W�� �{��(<%=MaxEtMd_Num%>)
	echo ------------------------------

	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	cls
	if %ERRORLEVEL% GTR 2 goto rootMenu
	goto Ch_5-%ERRORLEVEL%

	:Ch_5-1
		rem ---------- ��W ���ص{�� ----------
<%			call SubRename()%>
		pause
		goto rootMenu


	:Ch_5-2
		rem ---------- �_���W ���ص{�� ----------
<%			call SubRestoreRename()%>
		pause
		goto rootMenu


:Ch_6
<%FrontSpace="	echo "%>
	echo.
	echo ��������޲z:�y<%=SoftFullName%>�z
	echo �סססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
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


rem �� = = = = = = = = = = = = = = = ���~�T�� = = = = = = = = = = = = = = = ��
:Er_Value
	cls
	echo �z�w�w�w�w�w�w�{
	echo �x�@���~�I�I�@�x
	echo �|�w�w�w�w�w�w�}
	echo Error Message:
	echo %Er_Msg_%
	echo.
	pause
	goto end

<%
'���ͤU���R�O (2/2)
Response.End
%>