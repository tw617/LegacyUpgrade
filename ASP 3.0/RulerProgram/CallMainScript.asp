<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<%
'�b�����Ĥ@��[�J�U���o��A�`�N���|�n��A���t�̫e������޸�'
'<!-- #include file="CallMainScript.asp" -->

'�ŧi�q���ܼ�
Dim Transmit, FrontTip, PgFront, SvFront, SvRear, CounterNum, CounterNumView
%>


<%
'�סססס� ���o�ƶq �סססס�
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

<%'�� = = = = = = = = = = = = = = = Sub = = = = = = = = = = = = = = = ��%>

<%
'---------- �Ұ� �򥻵{�� ----------
Sub SubStartBasicPg()
		FrontTip="echo �Ұʰ򥻵{��:"
		PgFront="start"

	call FunctionStartProgram(dboBsPg)
	Response.Write Transmit
End Sub
%>
<%
'---------- �פ� �򥻵{�� ----------
Sub SubKillBasicPg()
		FrontTip="echo �פ�򥻵{��:"
		PgFront="taskkill /F /IM"

	call FunctionKillProgram(dboBsPg)
	Response.Write Transmit
End Sub
%>

<%
'---------- �Ұ� �X�R�{�� ----------
Sub SubStartExpandPg()
		FrontTip="echo �Ұ��X�R�{��:"
		PgFront="start"

	call FunctionStartProgram(dboEpPg)
	Response.Write Transmit
End Sub
%>
<%
'---------- �פ� �X�R�{�� ----------
Sub SubKillExpandPg()
		FrontTip="echo �פ��X�R�{��:"
		PgFront="taskkill /F /IM"

	call FunctionKillProgram(dboEpPg)
	Response.Write Transmit
End Sub
%>

<%
'---------- �Ұ� �򥻪A�� ----------
Sub SubStartBasicSv()
		FrontTip="echo �Ұʰ򥻪A��:"
		SvFront="net start"
		SvRear=""

	call FunctionService(dboBsSv)
	Response.Write Transmit
End Sub
%>
<%
'---------- ���� �򥻪A�� ----------
Sub SubKillBasicSv()
		FrontTip="echo ����򥻪A��:"
		SvFront="net stop"
		SvRear=""

	call FunctionService(dboBsSv)
	Response.Write Transmit
End Sub
%>

<%
'---------- �Ұ� �X�R�A�� ----------
Sub SubStartExpandSv()
		FrontTip="echo �Ұ��X�R�A��:"
		SvFront="net start"
		SvRear=""

	call FunctionService(dboEpSv)
	Response.Write Transmit
End Sub
%>
<%
'---------- ���� �X�R�A�� ----------
Sub SubKillExpandSv()
		FrontTip="echo �����X�R�A��:"
		SvFront="net stop"
		SvRear=""

	call FunctionService(dboEpSv)
	Response.Write Transmit
End Sub
%>


<%
'---------- ���� (��/�X�R)�A�� ----------
Sub SubStopSv(OpenDboName)
		FrontTip="echo ���ΪA��:"
		SvFront="sc config"
		SvRear="start= disabled"

	call FunctionService(OpenDboName)
	Response.Write Transmit
End Sub
%>


<%
'---------- (��/�X�R)�A�� �אּ��ĳ���A ----------
Sub SubStateToSuggestType(OpenDboName)
		FrontTip="echo "
		SvFront="sc config"
		SvRear=""

	call FunctionSuggestType(OpenDboName)
	Response.Write Transmit
End Sub
%>

<%
'---------- (��/�X�R)�A�� �אּ�w�]���A ----------
Sub SubStateToDefaultType(OpenDboName)
		FrontTip="echo "
		SvFront="sc config"
		SvRear=""

	call FunctionDefaultType(OpenDboName)
	Response.Write Transmit
End Sub
%>

<%
'�סססס� �ӧO�A�ȳ]�w(�ؿ�) �סססס�
Sub SubServiceParticularMenu()
	Transmit=""
	CounterNum=0
	CounterNumView=0

	Transmit = Transmit & FrontSpace & "------- ��@���@�A�@�� -------" & vbNewLine
	call FunctionServiceParticularMenu(dboBsSv)
	Transmit = Transmit & FrontSpace & "------- �X�@�R�@�A�@�� -------" & vbNewLine
	call FunctionServiceParticularMenu(dboEpSv)
	Transmit = Transmit & FrontSpace & "------------------------------" & vbNewLine
	Response.Write Transmit
End Sub
%>

<%
'�סססס� �ӧO�A�ȳ]�w(�ﶵ) �סססס�
Sub SubServiceParticularOption()
	Transmit=""
	CounterNum=0

	call FunctionServiceParticularOption(dboBsSv)
	call FunctionServiceParticularOption(dboEpSv)
	Response.Write Transmit
End Sub
%>

<%'�� = = = = = = = = = = = = = = = Function = = = = = = = = = = = = = = = ��%>

<%
'�\��: �Ұ�/���� �A��
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
'�\��: ����{��
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
'�\��: �פ�{��
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
'�\��: (��/�X�R)�A�� �אּ��ĳ���A
Function FunctionSuggestType(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Transmit=""
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			Transmit = Transmit & FrontSpace & FrontTip & "���A��: " & Trim(RS("ServiceView")) & vbNewLine
			Transmit = Transmit & FrontSpace & FrontTip & "    ��ĳ���A: " & Trim(RS("ServiceSuggestType")) & vbNewLine
			Transmit = Transmit & FrontSpace & SvFront & " """ & Trim(RS("ServiceName")) & """ start= " & Trim(RS("ServiceSuggestType")) & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%
'�\��: (��/�X�R)�A�� �אּ�w�]���A
Function FunctionDefaultType(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Transmit=""
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			Transmit = Transmit & FrontSpace & FrontTip & "���A��: " & Trim(RS("ServiceView")) & vbNewLine
			Transmit = Transmit & FrontSpace & FrontTip & "    �w�]���A: " & Trim(RS("ServiceDefaultType")) & vbNewLine
			Transmit = Transmit & FrontSpace & SvFront & " """ & Trim(RS("ServiceName")) & """ start= " & Trim(RS("ServiceDefaultType")) & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>

<%
'�סססס� �ӧO�A�ȳ]�w(�ؿ�) �סססס�
Function FunctionServiceParticularMenu(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			CounterNum = CounterNum +1

			'�N > 10 ���ȧ令�^�����
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
'�סססס� �ӧO�A�ȳ]�w(�ﶵ) �סססס�
Function FunctionServiceParticularOption(OpenDboName)
Call FunctionDataSource(DataBase)
RS.Open "[" & OpenDboName & "$]",DBConnection,3,1
	RS.MoveFirst
	Do While RS.EOF = False
		RS.Find="[SoftBriefName]='" & Brief & "'"
		if NOT RS.EOF then
			CounterNum = CounterNum +1
			Transmit = Transmit & FrontSpace & ":Ch_3-" & CounterNum & vbNewLine
			Transmit = Transmit & FrontSpace & "	rem ---------- �ާ@�A�ȸ��: " & Trim(RS("ServiceName")) & " ----------" & vbNewLine
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