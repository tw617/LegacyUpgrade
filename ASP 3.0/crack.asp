<html>
<head>
<title>破解檔 設定</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-tw">
</head>

<%
'● =================================== 設定宣告 =================================== ●
CrackScriptVer = "2012.08.26"

Dim Textbox_SoftTitle			'軟體名稱
Dim Textarea_ExecutionFileName	'軟體主程式檔名
Dim Textarea_Crack				'路徑與檔名
Dim Count_Num					'檔案數量


Textbox_SoftTitle = Request("Textbox_SoftTitle")
Textarea_ExecutionFileName = Request("Textarea_ExecutionFileName")
Textarea_Crack = Request("Textarea_Crack")


CrackSplitCL=Split(Textarea_Crack, vbCrLf)
%>

<body>
<form Name="Crack" method="POST" action="crack.asp">
	<p align="center"><b><font size="7" color="#FF0000">破解檔設定</font></b></p>
	<div align="center">
	<table border="3" cellspacing="5" cellpadding="2">
		<tr>
			<td width="125" align="right">名稱</td>
			<td align="left">設定值</td>
		</tr>
		<tr>
			<td align="right">軟體名稱</td>
			<td align="left">
			<input type="text" name="Textbox_SoftTitle" size="40" value="<%=Textbox_SoftTitle%>" tabindex="1"></td>
		</tr>
		<tr>
			<td align="right">程式執行檔名稱<br>(含副檔名)</td>
			<td align="left">
			<textarea rows="2" name="Textarea_ExecutionFileName" cols="40" tabindex="2"><%=Textarea_ExecutionFileName%></textarea></tr>
		<tr>
			<td align="right">相對路徑及檔名<br>(要置換的檔案，含副檔名)</td>
			<td align="left">
			<textarea rows="5" name="Textarea_Crack" cols="40" tabindex="3"><%=Textarea_Crack%></textarea></td>
		</tr>
		<tr>
			<td align="right">　</td>
			<td align="left">　</td>
		</tr>
		</table>
	</div>
	<p align="center"><input type="submit" value="送出" name="B1" tabindex="4"><input type="reset" value="重新設定" name="B2"></p>
</form>


</body>

</html>


<%
'無軟體名稱，則停止輸出網頁
If Request("Textbox_SoftTitle")="" then response.End

'產生下載命令：起始部分 (1/2)
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=crack.bat"
Response.ContentType = "text"
%>

@echo off
rem =================================== 資　　訊 ===================================
Title <%=Textbox_SoftTitle%> 破解

echo Script 版本：<%=CrackScriptVer%>
echo 批次檔製作日期：<%=NOW%>

echo 軟體名稱：<%=Textbox_SoftTitle%>

echo 執行檔名稱：<%=replace(Textarea_ExecutionFileName,vbCrLf,vbNewLine & "echo 　　　　　　")%>

echo 置換檔案數量：<%=UBound(CrackSplitCL) +1%>
echo 檔案列表：<%=replace(Textarea_Crack,vbCrLf,vbNewLine & "echo 　　　　　")%>

rem =================================== 系統位元 ===================================
if %PROCESSOR_ARCHITECTURE% == x86 (
	echo 這是 x86 系統
	set pf32=%ProgramFiles%
	set cpf32=%CommonProgramFiles%
	)

if %PROCESSOR_ARCHITECTURE% == AMD64 (
	echo 這是 x64 系統
	set pf32=%ProgramFiles(x86)%
	set cpf32=%CommonProgramFiles(x86)%
	set pf64=%ProgramFiles%
	set cpf64=%CommonProgramFiles%
	)

rem =================================== 腳本開始 ===================================
echo 處裡檔案：開始
echo ------------------------- 第一部份 -------------------------
rem 終止程式
taskkill /F /IM "<%=replace(Textarea_ExecutionFileName,vbCrLf,Chr(34) & vbNewLine & "taskkill /F /IM " & Chr(34))%>"

echo ------------------------- 第二部份 -------------------------
rem 檔案更名 (+.bak)
<%
'Crack_i = 取出 Textarea_Crack 第一列到最後一列的整列資料
For Crack_i = 0 to UBound(CrackSplitCL)
	
	'指令
	CrackRen_1 = "ren "
	
	'來源檔案名稱
	CrackRen_2 = Chr(34) & CrackSplitCL(Crack_i) & Chr(34)
		
	'目的檔案名稱
	CrackSplit92=Split(CrackSplitCL(Crack_i), Chr(92))	'使用 \ 分割 → Textarea_Crack 取出的列字串
	CrackSplit92U=CrackSplit92(UBound(CrackSplit92))	'從分割的陣列中，使用最大的索引值來取得檔案名稱
	CrackRen_3 = CrackSplit92U & ".bak"
	
	response.write CrackRen_1 & CrackRen_2 & Chr(32) & CrackRen_3 & vbnewline
	
	Next
%>

echo ------------------------- 第三部份 -------------------------
rem 刪除重複存在檔案
rem 修改來源：第一部份
del /f "<%=replace(Textarea_Crack,vbCrLf,Chr(34) & vbNewLine & "del /f " & Chr(34))%>"

echo ------------------------- 第四部份 -------------------------
rem 檔案更名 (-.crack)
rem 修改來源：第二部份
<%
'Crack_i = 取出 Textarea_Crack 第一列到最後一列的整列資料
For Crack_i = 0 to UBound(CrackSplitCL)
	
	'指令
	CrackRen_1 = "ren "
	
	'來源檔案名稱
	CrackRen_2 = Chr(34) & CrackSplitCL(Crack_i) & ".crack" & Chr(34)
		
	'目的檔案名稱
	CrackSplit92=Split(CrackSplitCL(Crack_i), Chr(92))	'使用 \ 分割 → Textarea_Crack 取出的列字串
	CrackSplit92U=CrackSplit92(UBound(CrackSplit92))	'從分割的陣列中，使用最大的索引值來取得檔案名稱
	CrackRen_3 = CrackSplit92U
	
	response.write CrackRen_1 & CrackRen_2 & Chr(32) & CrackRen_3 & vbnewline
	
	Next
%>

echo ------------------------- 第五部份 -------------------------
rem 刪除無法更名的 .crack
rem 修改來源：第一部份
del /f "<%=replace(Textarea_Crack,vbCrLf,".crack" & Chr(34) & vbNewLine & "del /f " & Chr(34))%>.crack"


echo 處裡檔案：結束
rem =================================== 腳本結束 ===================================
del /f crack.bat
exit

<%
'產生下載命令：結束部分 (2/2)
Response.End
%>