<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<!-- #include file="CallMain_ConnectionDataTable.asp" -->
<!-- #include file="CallMainScript.asp" -->
<%
'�� = = = = = = = = = = = = = = = �]�w�ŧi = = = = = = = = = = = = = = = ��
RulerOS=Request.QueryString("RulerOS")	'���o���j�M�����
SubTitle=CStr("" & Request.QueryString("SubTitle"))	'���o���j�M�����
RulerRobotVer="0.3"		'�M�μҲճ]�w�� �Ҳժ���

'�� = = = = = = = = = = = = = = = ���@�@�� = = = = = = = = = = = = = = = ��
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

'-------------------- ����: �����ŧi --------------------
ResultTime1 = Year(now) & "/" & Right("0" & Month(now), 2) & "/" & Right("0" & Day(now), 2)
ResultTime2 = Right("0" & Hour(now), 2) & ":" & Right("0" & Minute(now), 2) & ":" & Right("0" & Second(now), 2)
%>

<%
'���ͤU���R�O (1/2)
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=Ruler-System_" & Replace(OSName,Chr("32"),"%20") & ".bat"
Response.ContentType = "text"
Call FunctionDataSource(DataBase)
%>
@echo off
rem �סססססססססס@���@�@���@�ססססססססס�
echo �ШϥΨt�κ޲z����������
echo �ҲզW��: <%=RulerOS & SubTitle%>
echo ��檩��: <%=RulerRobotVer%>
echo �s�@���: <%=ResultTime1 & " " & ResultTime2%>

rem �סססססססססס@�ˬd�t�Ρ@�ססססססססס�
set "HostOS="
ver|find /i " 10.0">nul && set "HostOS=Win10.0"
ver|find /i " 6.1">nul && set "HostOS=Win6.1"
ver|find /i " 6.0">nul && set "HostOS=Win6.0"
ver|find /i " 5.2">nul && set "HostOS=Win5.2"
ver|find /i " 5.1">nul && set "HostOS=Win5.1"
ver|find /i " 5.0">nul && set "HostOS=Win5.0"

IF "%HostOS%" EQU "" (
		set "Er_Msg_=��p�A�L�k��t�o�Ӹ޲����t�ΡC"
		goto Er_Value
	) ELSE (
		IF "%HostOS%" NEQ "<%=RulerOS%>" (
			set "Er_Msg_=��p�A�����O�ɨä��A�Ω󦹨t�ΡC"
			goto Er_Value
		)
	)

<%
'�� = = = = = = = = = = = = = = = ���o�ƶq = = = = = = = = = = = = = = = ��
	Dim ArrayRowNum
'-------------------- �t�Υ\�� --------------------
	call FunctionGetRowNum(dboSF)
	MaxSystemFunction_Num = ArrayRowNum

'-------------------- �t�ΪA�� --------------------
	call FunctionGetRowNum(dboSS)
	MaxSystemService_Num = ArrayRowNum
%>
Title �t�ΡG<%=OSName%>
:rootMenu
cls
echo �п��!!�@�y<%=OSName%>�z
echo �ססססססססססססססססססס�
echo  0. ���� (*�w�])
echo ----------------------------------------
echo  1. �����t�ΪA�ȳ]�� ��ĳ���A (<%=MaxSystemService_Num%>) (*��ĳ)
echo  2. �����t�ΪA�ȳ]�� �w�]���A (<%=MaxSystemService_Num%>)
echo ----------------------------------------
echo  3. �A�ȭӧO�]�w (<%=MaxSystemService_Num%>)
echo ----------------------------------------
echo  4. �t�Υ\�� (<%=MaxSystemFunction_Num%>)
echo ----------------------------------------
echo  9. �t������/���}��
echo ----------------------------------------

CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
cls
if %ERRORLEVEL% GEQ 36 goto end
if %ERRORLEVEL% GEQ 9 goto Ch_9
if %ERRORLEVEL% GTR 4 goto rootMenu
goto Ch_%ERRORLEVEL%

rem �� = = = = = = = = = = = = = = = �ﶵ���� = = = = = = = = = = = = = = = ��
:Ch_1
<%FrontSpace="	"%>
	rem ---------- ��ĳ���A ----------
<%	call SubStateToSuggestType()%>
	pause
	goto end


:Ch_2
	rem ---------- �w�]���A ----------
<%	call SubStateToDefaultType()%>
	pause
	goto end


:Ch_3
<%FrontSpace="	echo "%>
	rem ---------- �ӧO�A�ȳ]�w ----------
	cls
	echo.
	echo �סססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
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
		echo  ���A�ȴy�z�G��%ServiceDescribe%��
		echo  ���w�]���A�G��%ServiceDefaultType%��
		echo  ����ĳ���A�G��%ServiceSuggestType%��
		echo  ��    (�۰�: auto�F����: delayed-auto�F���: demand�F����: disabled)
		echo  -----------------------------
		echo  ���ثe���A�G��%ServiceMomentState%��
		echo  ���ثe�����G��%ServiceMomentType%��
		echo  -----------------------------
		echo  ���ƧѬ����G��%ServiceMemo%��
		echo �סססססססססססססס�
		echo   0. �^�W�@�� (*�w�])
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
<%FrontSpace="	echo "%>
	rem ---------- �t�Υ\�� ----------
	echo.
	echo �סססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
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

	rem ---------- �t������/���}�� ----------
	echo.
	echo �סססססססססססססס�
	echo  0. �^�W�h�ؿ� (*�w�])
	echo ------------------------------
	echo  1. ���}��
	echo  2. ����
	echo  3. ��������
	echo ------------------------------
	CHOICE /C 123456789abcdefghijklmnopqrstuvwxyz0 /n /T 10 /D 0
	cls
	if %ERRORLEVEL% GEQ 36 goto rootMenu
	if %ERRORLEVEL% GTR 3 goto Ch_9
	goto Ch_9-%ERRORLEVEL%

		:Ch_9-1
			rem ---------- ���s�}�� ----------
			echo ------------------------------
			echo  ��ܤF��10��᭫�}��: %computername%���T�w�ܡH
			CHOICE /C YNC /n /M "�O(Y)�B�_(N)�B����(C)�C
			if %ERRORLEVEL% == 1 shutdown -r -m \\%computername% -t 10 -c "�������A�ǳƭ��}���C" -f
			if %ERRORLEVEL% == 2 goto Ch_9-1
			goto end

		:Ch_9-2
			rem ---------- ���� ----------
			echo ------------------------------
			echo ��ܤF��10�������: %computername%���T�w�ܡH
			CHOICE /C YNC /n /M "�O(Y)�B�_(N)�B����(C)�C
			if %ERRORLEVEL% == 1 shutdown -s -m \\%computername% -t 10 -c "������" -f
			if %ERRORLEVEL% == 2 goto Ch_9-1
			goto end

		:Ch_9-3
			rem ---------- �������� ----------
			shutdown -a
			goto end


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
DBConnection.close

'���ͤU���R�O (2/2)
Response.End
%>