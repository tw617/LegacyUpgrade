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
<p>�п�ܭn�U��������ɮסG</p>
<table border="5" cellspacing="3" cellpadding="5" bordercolorlight="#808080" bordercolordark="#800000">
<tr><td align="center"><b>�n��W��</b></td><td align="center"><b>�̫��s���</b></td></tr>
<%
Call FunctionDataSource(DataBase)
RS.Open "[" & dboIndex & "$]",DBConnection,0,1

if rs.EOF then
	response.write "<p align='center'><b><font size='7' face='�з���' color='#FF0000'>�d�L���<br></font></b></p>"
	end if

RS.MoveFirst		' �N�ثe��ƿ�����Ĥ@��
While Not RS.EOF	' �P�_�O�_�L�F�̫�@��

	'==================== ��}�˪O ====================
	Row = "<TR>"
	Row = Row & "<TD><a href='RulerRobot.asp?Brief=" & Trim(RS("SoftBriefName")) & "'>" & Trim(RS("SoftFullName")) & "</a></TD>" & vbNewLine
	Row = Row & "    <TD>" & Trim(RS("LastUpdate")) & "</TD>" & vbNewLine
	'==================================================

	Response.Write Row & "    </TR>" & vbNewLine
	rs.MoveNext	' ����U�@��
	Wend

RS.Close
DBConnection.close
%>
</table>
</div>
</body>
</html>
