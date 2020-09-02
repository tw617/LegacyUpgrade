<!-- #include file="CallMainScript.asp" -->
<%
Dim xlsFile, LastXlsFile, xlsFilePath
LastXlsFile=Request.QueryString("xlsFile")
xlsFilePath=Request.QueryString("xlsFile") & ".xls"
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-tw">
<title>Microsoft Update List</title>
</head>

<body link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<p>請選擇要下載更新資料的系統軟體：</p>
<table border="5" cellspacing="3" cellpadding="5" bordercolorlight="#808080" bordercolordark="#800000">

<tr><td align="center"><b>更新 系統、軟體 名稱</b></td><%call SubTitleGetBatch()%></tr>
<TR><%call FunctionHyperlinks("5.1-WinXP_x32_sp3;Microsoft Windows XP x86 SP3")%><%call SubGetBatch()%></TR>
<TR><%call FunctionHyperlinks("5.1-WinXP_x64_sp2;Microsoft Windows XP x64 SP2")%></TR>
<TR><%call FunctionHyperlinks("5.2-WS2k3_x32_sp2;Microsoft Windows Server 2003 x86 SP2")%></TR>
<TR><%call FunctionHyperlinks("5.2-WS2k3_x64_sp2;Microsoft Windows Server 2003 x64 SP2")%></TR>
<TR><%call FunctionHyperlinks("6.1-Win7_x32_sp1;Microsoft Windows 7 x86 SP1")%></TR>
<TR><%call FunctionHyperlinks("6.1-Win7_x64_sp1;Microsoft Windows 7 x64 SP1")%></TR>
<TR><%call FunctionHyperlinks("WS2k8_x86;Microsoft Windows Server 2008 x86")%></TR>
<TR><%call FunctionHyperlinks("WS2k8_x64;Microsoft Windows Server 2008 x64")%></TR>
<TR><%call FunctionHyperlinks("O2k3_sp3;Microsoft Office 2003 SP3")%></TR>
<TR><%call FunctionHyperlinks("O2k7_sp2;Microsoft Office 2007 SP2")%></TR>
</table>
</div>
</body>

</html>