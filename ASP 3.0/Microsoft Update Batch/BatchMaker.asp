<!-- #include file="CallMainScript.asp" -->
<%
'�� = = = = = = = = = = = = = = = �]�w�ŧi = = = = = = = = = = = = = = = ��
Dim xlsFilePath, Sheet
xlsFilePath=Request.QueryString("xlsFilePath")
Sheet=Request.QueryString("Sheet")

call FunctionDataSource(xlsFilePath)
RS.CursorLocation = 3
RS.Open "[" & Sheet & DataTableAdjustString & "]",DBConnection,3,1
%>
<%
'���ͤU���R�O (1/2)
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=" & Replace(Sheet," ","_") & ".bat"
Response.ContentType = "text"
%>
@echo off
Title <%=Sheet%> [<%call SubGetRowCount()%>��] �]<%call SubGetNow()%>�^
echo �w�˶i�סG<%=Sheet%>
set "<%=replace(Sheet," ","_")%>Num=<%call SubGetRowCount()%>"
<%
call SubGetBatchData(Sheet)
RS.Close
DBConnection.close

'���ͤU���R�O (2/2)
Response.End
%>
