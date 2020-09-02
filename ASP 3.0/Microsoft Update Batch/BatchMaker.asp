<!-- #include file="CallMainScript.asp" -->
<%
'● = = = = = = = = = = = = = = = 設定宣告 = = = = = = = = = = = = = = = ●
Dim xlsFilePath, Sheet
xlsFilePath=Request.QueryString("xlsFilePath")
Sheet=Request.QueryString("Sheet")

call FunctionDataSource(xlsFilePath)
RS.CursorLocation = 3
RS.Open "[" & Sheet & DataTableAdjustString & "]",DBConnection,3,1
%>
<%
'產生下載命令 (1/2)
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=" & Replace(Sheet," ","_") & ".bat"
Response.ContentType = "text"
%>
@echo off
Title <%=Sheet%> [<%call SubGetRowCount()%>個] （<%call SubGetNow()%>）
echo 安裝進度：<%=Sheet%>
set "<%=replace(Sheet," ","_")%>Num=<%call SubGetRowCount()%>"
<%
call SubGetBatchData(Sheet)
RS.Close
DBConnection.close

'產生下載命令 (2/2)
Response.End
%>
