<%
'�b�����Ĥ@��[�J�U���o��A�`�N���|�n��A���t�̫e������޸�'
'<!-- #include file="CallSecondaryScript.asp" -->

'�ŧi�q���ܼ�
Dim FileSourceName, FileTargetName
%>

<%'�� = = = = = = = = = = = = = = = Sub = = = = = = = = = = = = = = = ��%>

<%
'�סססס� ��W ���ص{�� �סססס�
Sub SubRename()
		FrontTip="echo ��W: "
		FileSourceName="SoftFileOriginalName"
		FileTargetName="SoftFileNewName"

	call FunctionRename(dboEtRn)
	Response.Write Transmit
End Sub
%>


<%
'�סססס� �_���W ���ص{�� �סססס�
Sub SubRestoreRename()
		FrontTip="echo �_���W: "
		FileSourceName="SoftFileNewName"
		FileTargetName="SoftFileOriginalName"

	call FunctionRename(dboEtRn)
	Response.Write Transmit
End Sub
%>


<%
'�סססס� ��������(���D) �סססס�
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
'�סססס� ��������(�ﶵ) �סססס�
Sub SubCommandRulerOption()
	Transmit=""
	CounterNum=0

	call FunctionCommandRulerOption(dboEtCR)
	Response.Write Transmit
End Sub
%>

<%'�� = = = = = = = = = = = = = = = Function = = = = = = = = = = = = = = = ��%>

<%
'�\��: ��W ���ص{��
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
'�\��: ��������(�ؿ�)
Function FunctionCommandRulerMenu(OpenDboName)
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

			Transmit = Transmit & FrontSpace & " " & CounterNumView & ". " & Trim(RS("CommandTitle")) & vbNewLine
			RS.MoveNext
			End if
		Loop
RS.Close
DBConnection.Close
End Function
%>


<%
'�\��: ��������(�ﶵ)
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
			Transmit = Transmit & FrontSpace & "	rem �ާ@: " & Trim(RS("CommandTitle")) & vbNewLine
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