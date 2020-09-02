<!-- #include file="CallMain_ConnectionDataSource.asp" -->
<!-- #include file="CallMain_ConnectionDataTable.asp" -->
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-tw">
<title>Service Model</title>
</head>

<body link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<p>請選擇要下載的文件檔案：</p>
<table border="5" cellspacing="3" cellpadding="5" bordercolorlight="#808080" bordercolordark="#800000">
<tr><td align="center"><b>軟體名稱</b></td><td align="center"><b>最後更新日期</b></td></tr>
<%
Call FunctionDataSource(DataBase)
RS.Open "[" & dboIndex & "$]",DBConnection,0,1

if rs.EOF then
	response.write "<p align='center'><b><font size='7' face='標楷體' color='#FF0000'>查無資料<br></font></b></p>"
	end if

RS.MoveFirst		' 將目前資料錄移到第一筆
While Not RS.EOF	' 判斷是否過了最後一筆

	'==================== 改良樣板 ====================
	Row = "<TR>"
	Row = Row & "<TD><a href='RulerRobot.asp?Brief=" & Trim(RS("SoftBriefName")) & "'>" & Trim(RS("SoftFullName")) & "</a></TD>" & vbNewLine
	Row = Row & "    <TD>" & Trim(RS("LastUpdate")) & "</TD>" & vbNewLine
	'==================================================

	Response.Write Row & "    </TR>" & vbNewLine
	rs.MoveNext	' 移到下一筆
	Wend

RS.Close
DBConnection.close
%>
</table>
</div>
</body>
</html>
