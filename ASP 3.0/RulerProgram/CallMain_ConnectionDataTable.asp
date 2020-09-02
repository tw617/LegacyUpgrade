<%
'在網頁第一行加入下面這行，注意路徑要改，不含最前面的單引號'
'<!-- #include file="CallMain_ConnectionDataSource.asp" -->
'<!-- #include file="CallMain_ConnectionDataTable.asp" -->

'＝＝＝＝＝ 資料表名稱 ＝＝＝＝＝
dboIndex = "DB_Index"		'主索引
dboBsPg = "DB_BaseProgram"	'基本程式
dboBsSv = "DB_BaseService"	'基本服務
dboEpPg = "DB_ExpandProgram"	'擴充程式
dboEpSv = "DB_ExpandService"	'擴充服務
dboEtRn = "DB_ExtendRename"	'延伸 - 修改檔案列表
dboEtCR = "DB_ExtendCommandRuler"	'延伸 - 延伸控制
%>
