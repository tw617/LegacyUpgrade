<%
'在網頁第一行加入下面這行，注意路徑要改，不含最前面的單引號'
'<!-- #include file="ServerData.asp" -->
'※※※※※※※※※※※※※※※
'伺服器電腦名稱
ServerName="vm-ws2k3e-x86-1"

'設定SQL讀取及寫入的帳號
LoginID="sa"

'設定SQL讀取及寫入的密碼
LoginPw="0000"

'設定SQL資料庫名稱
DatabaseName="RulerSystem"
'※※※※※※※※※※※※※※※
'資料表名稱...

dboSF = "DB_SystemFunction"		'系統功能
dboSS = "DB_SystemService"		'系統服務

'※※※※※※※※※※※※※※※
'聯結資料庫指令

Set DBConnection = Server.CreateObject("ADODB.Connection")
ConnectionString ="Provider=SQLOLEDB.1;Server=" & ServerName & ";UID=" & LoginID & ";PWD=" & LoginPw & ";Database=" & DatabaseName
DBConnection.Open ConnectionString
Set RS=Server.CreateObject("ADODB.Recordset")
'RS.Open (欲開啟資料之名稱，請使用變數),DBConnection,?,?
'	...
'	...
'RS.Close
'DBConnection.close


'Rs.Open 資料來源，資料連結，指標型態，鎖定方式
'資料來源：指定資料表名稱
'資料連結：指定一個已Connection的物件
'
'指標型態：
'０：只能向前移動的指標，此為預設值
'１：無法讀取其他使用者新增的資料，更新的資料會立即反應
'２：可以即時反應其他使用著操作資料庫之狀況
'３：無法即時反應其他使用著操作相同資料庫的狀況，用於搜尋或新增記錄時使用
'
'鎖定方式：
'１：將 Recordset 開啟為唯讀狀態，此為預設值
'２：當使用著對 Recordset 中的某筆資料作編輯時，才鎖定記錄
'３：當使用著呼叫 Update 方法對 Recordset 做更新時，才鎖定記錄
'４：使用者做批次更新時，才鎖定記錄
'
'
'Options 選擇性參數[非必要]：
'一個 Long 值，表示提供者在 Source 引數代表 Command 物件以外的東西時應如何評估它，否則 Recordset 應從前次儲存的檔案還原。
'它可以是下列其中一種常數。
'
'常數說明
'adCmdText：提供者會將 Source 評估為指令的文字定義。
'AdCmdTable：ADO 會產生一個 SQL 查詢，從 Source 中指定的資料表傳回所有資料列。
'AdCmdTableDirect：提供者會從 Source 中指定的資料表傳回所有資料列。
'AdCmdStoredProc：提供者會將 Source 評估為一個預存程序。
'AdCmdUnknown：Source 引數中未知的指令類型。
'AdCommandFile：保留的 (已儲存的) Recordset 會從 Source 中指定的檔案還原。
'AdExecuteAsync：Source 作非同步執行。
'AdFetchAsync：表示在 CacheSize 屬性中指定的初始數量被抓取後，剩下的資料列就會被非同步地抓取。
%>