<%
'在網頁第一行加入下面這行，注意路徑要改，不含最前面的單引號'
'<!-- #include file="CallMain_ConnectionDataSource.asp" -->
'Dim SourceMode, ServerName, LoginID, LoginPw, DataBase, DataTableAdjustString, DBConnection, RS

SourceMode = "Excel2000"
	rem Excel2000: 資料來源為 Microsoft Excel 2000.xls
	rem MSSQL2000: 資料來源為 Microsoft SQL Server 2000

Select Case SourceMode
	Case "Excel2000"
		'設定 xls 檔案名稱(可包含相對路徑)
		DataBase = "RulerProgram.xls"
		DataTableAdjustString = "$"

	Case "MSSQL2000"
		'伺服器電腦名稱
		ServerName = "vm-ws2k3e-x86-1"

		'設定SQL讀取及寫入的帳號
		LoginID = "sa"

		'設定SQL讀取及寫入的密碼
		LoginPw = "0000"

		'設定SQL資料庫名稱
		DataBase = ""
		DataTableAdjustString = ""

	Case Else
		'未提供參數
		response.write vbNewLine
		response.write "錯誤:" & vbNewLine
		response.write "CallMain_ConnectionString.asp" & vbNewLine
		response.write "未設定資料來源模式" & vbNewLine
		response.End
End Select

Set DBConnection = Server.CreateObject("ADODB.Connection")
Set RS = Server.CreateObject("ADODB.Recordset")


'＝＝＝＝＝ 選擇資料來源 ＝＝＝＝＝
Function FunctionDataSource(DataBaseName)
	If DatabaseName = "" then DatabaseName = DataBase

	Select Case SourceMode
		Case "Excel2000"
			'設定 Microsoft Excel 2000.xls 的資料來源

				'產生錯誤	ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DBQ=" & Server.MapPath(DatabaseName) & ";Extended Properties=Excel 8.0"
			ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DatabaseName) & ";Extended Properties=Excel 8.0"
			DBConnection.Open ConnectionString
			rem RS.Open "[" & SheetName｛欲開啟資料之名稱，請使用變數｝ & "$]",DBConnection,?,?


		Case "MSSQL2000"
			'設定 Microsoft SQL Server 2000 的資料來源

			ConnectionString ="Provider=SQLOLEDB.1;Server=" & ServerName & ";UID=" & LoginID & ";PWD=" & LoginPw & ";Database=" & DatabaseName
			DBConnection.Open ConnectionString
			rem RS.Open (欲開啟資料之名稱，請使用變數),DBConnection,?,?

		Case Else
			'未提供參數
			response.write vbNewLine
			response.write "錯誤:" & vbNewLine
			response.write "CallMain_ConnectionString.asp" & vbNewLine
			response.write "未設定資料來源模式" & vbNewLine
			response.End
			
	End Select

End Function
%>