<%
M_URLScriptVer = "2013.09.10"
M_URL = Request("M_URL")	'�����ƻs Windows Update �W����s�M��
%>

<html>

<head>
<title>Open Multi URL</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-tw">
</head>

<body>
<p>�K�W���}</p>
<form method="POST" action="OpenMultiURL.asp">
	<p><textarea rows="10" name="M_URL" cols="81" tabindex="1"><%=M_URL%></textarea></p>
	<p><input type="submit" value="�e�X" name="Submit" tabindex="2"><input type="reset" value="���s�]�w" name="Restore"></p>
</form>
<hr>

<p>
<%
response.write "<SCRIPT LANGUAGE='javascript'>"
tmpM_URL = split(M_URL, VBCRLF)

For tmpLoop = 0 to UBound(tmpM_URL) Step 1
  openURL = Trim(tmpM_URL(tmpLoop))
  IF NOT openURL = Empty then response.write "window.open('" & openURL & "');"
Next
response.write "</SCRIPT>"
%>
</p>
</body>

</html>
