<html>
<head>
<title>�}���� �]�w</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-tw">
</head>

<%
'�� =================================== �]�w�ŧi =================================== ��
CrackScriptVer = "2012.08.26"

Dim Textbox_SoftTitle			'�n��W��
Dim Textarea_ExecutionFileName	'�n��D�{���ɦW
Dim Textarea_Crack				'���|�P�ɦW
Dim Count_Num					'�ɮ׼ƶq


Textbox_SoftTitle = Request("Textbox_SoftTitle")
Textarea_ExecutionFileName = Request("Textarea_ExecutionFileName")
Textarea_Crack = Request("Textarea_Crack")


CrackSplitCL=Split(Textarea_Crack, vbCrLf)
%>

<body>
<form Name="Crack" method="POST" action="crack.asp">
	<p align="center"><b><font size="7" color="#FF0000">�}���ɳ]�w</font></b></p>
	<div align="center">
	<table border="3" cellspacing="5" cellpadding="2">
		<tr>
			<td width="125" align="right">�W��</td>
			<td align="left">�]�w��</td>
		</tr>
		<tr>
			<td align="right">�n��W��</td>
			<td align="left">
			<input type="text" name="Textbox_SoftTitle" size="40" value="<%=Textbox_SoftTitle%>" tabindex="1"></td>
		</tr>
		<tr>
			<td align="right">�{�������ɦW��<br>(�t���ɦW)</td>
			<td align="left">
			<textarea rows="2" name="Textarea_ExecutionFileName" cols="40" tabindex="2"><%=Textarea_ExecutionFileName%></textarea></tr>
		<tr>
			<td align="right">�۹���|���ɦW<br>(�n�m�����ɮסA�t���ɦW)</td>
			<td align="left">
			<textarea rows="5" name="Textarea_Crack" cols="40" tabindex="3"><%=Textarea_Crack%></textarea></td>
		</tr>
		<tr>
			<td align="right">�@</td>
			<td align="left">�@</td>
		</tr>
		</table>
	</div>
	<p align="center"><input type="submit" value="�e�X" name="B1" tabindex="4"><input type="reset" value="���s�]�w" name="B2"></p>
</form>


</body>

</html>


<%
'�L�n��W�١A�h�����X����
If Request("Textbox_SoftTitle")="" then response.End

'���ͤU���R�O�G�_�l���� (1/2)
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=crack.bat"
Response.ContentType = "text"
%>

@echo off
rem =================================== ��@�@�T ===================================
Title <%=Textbox_SoftTitle%> �}��

echo Script �����G<%=CrackScriptVer%>
echo �妸�ɻs�@����G<%=NOW%>

echo �n��W�١G<%=Textbox_SoftTitle%>

echo �����ɦW�١G<%=replace(Textarea_ExecutionFileName,vbCrLf,vbNewLine & "echo �@�@�@�@�@�@")%>

echo �m���ɮ׼ƶq�G<%=UBound(CrackSplitCL) +1%>
echo �ɮצC��G<%=replace(Textarea_Crack,vbCrLf,vbNewLine & "echo �@�@�@�@�@")%>

rem =================================== �t�Φ줸 ===================================
if %PROCESSOR_ARCHITECTURE% == x86 (
	echo �o�O x86 �t��
	set pf32=%ProgramFiles%
	set cpf32=%CommonProgramFiles%
	)

if %PROCESSOR_ARCHITECTURE% == AMD64 (
	echo �o�O x64 �t��
	set pf32=%ProgramFiles(x86)%
	set cpf32=%CommonProgramFiles(x86)%
	set pf64=%ProgramFiles%
	set cpf64=%CommonProgramFiles%
	)

rem =================================== �}���}�l ===================================
echo �B���ɮסG�}�l
echo ------------------------- �Ĥ@���� -------------------------
rem �פ�{��
taskkill /F /IM "<%=replace(Textarea_ExecutionFileName,vbCrLf,Chr(34) & vbNewLine & "taskkill /F /IM " & Chr(34))%>"

echo ------------------------- �ĤG���� -------------------------
rem �ɮק�W (+.bak)
<%
'Crack_i = ���X Textarea_Crack �Ĥ@�C��̫�@�C����C���
For Crack_i = 0 to UBound(CrackSplitCL)
	
	'���O
	CrackRen_1 = "ren "
	
	'�ӷ��ɮצW��
	CrackRen_2 = Chr(34) & CrackSplitCL(Crack_i) & Chr(34)
		
	'�ت��ɮצW��
	CrackSplit92=Split(CrackSplitCL(Crack_i), Chr(92))	'�ϥ� \ ���� �� Textarea_Crack ���X���C�r��
	CrackSplit92U=CrackSplit92(UBound(CrackSplit92))	'�q���Ϊ��}�C���A�ϥγ̤j�����ޭȨӨ��o�ɮצW��
	CrackRen_3 = CrackSplit92U & ".bak"
	
	response.write CrackRen_1 & CrackRen_2 & Chr(32) & CrackRen_3 & vbnewline
	
	Next
%>

echo ------------------------- �ĤT���� -------------------------
rem �R�����Ʀs�b�ɮ�
rem �ק�ӷ��G�Ĥ@����
del /f "<%=replace(Textarea_Crack,vbCrLf,Chr(34) & vbNewLine & "del /f " & Chr(34))%>"

echo ------------------------- �ĥ|���� -------------------------
rem �ɮק�W (-.crack)
rem �ק�ӷ��G�ĤG����
<%
'Crack_i = ���X Textarea_Crack �Ĥ@�C��̫�@�C����C���
For Crack_i = 0 to UBound(CrackSplitCL)
	
	'���O
	CrackRen_1 = "ren "
	
	'�ӷ��ɮצW��
	CrackRen_2 = Chr(34) & CrackSplitCL(Crack_i) & ".crack" & Chr(34)
		
	'�ت��ɮצW��
	CrackSplit92=Split(CrackSplitCL(Crack_i), Chr(92))	'�ϥ� \ ���� �� Textarea_Crack ���X���C�r��
	CrackSplit92U=CrackSplit92(UBound(CrackSplit92))	'�q���Ϊ��}�C���A�ϥγ̤j�����ޭȨӨ��o�ɮצW��
	CrackRen_3 = CrackSplit92U
	
	response.write CrackRen_1 & CrackRen_2 & Chr(32) & CrackRen_3 & vbnewline
	
	Next
%>

echo ------------------------- �Ĥ����� -------------------------
rem �R���L�k��W�� .crack
rem �ק�ӷ��G�Ĥ@����
del /f "<%=replace(Textarea_Crack,vbCrLf,".crack" & Chr(34) & vbNewLine & "del /f " & Chr(34))%>.crack"


echo �B���ɮסG����
rem =================================== �}������ ===================================
del /f crack.bat
exit

<%
'���ͤU���R�O�G�������� (2/2)
Response.End
%>