<%
'�b�����Ĥ@��[�J�U���o��A�`�N���|�n��A���t�̫e������޸�'
'<!-- #include file="CallMain_ConnectionDataSource.asp" -->
'Dim SourceMode, ServerName, LoginID, LoginPw, DataBase, DataTableAdjustString, DBConnection, RS

SourceMode = "Excel2000"
	rem Excel2000: ��ƨӷ��� Microsoft Excel 2000.xls
	rem MSSQL2000: ��ƨӷ��� Microsoft SQL Server 2000

Select Case SourceMode
	Case "Excel2000"
		'�]�w xls �ɮצW��(�i�]�t�۹���|)
		DataBase = "RulerProgram.xls"
		DataTableAdjustString = "$"

	Case "MSSQL2000"
		'���A���q���W��
		ServerName = "vm-ws2k3e-x86-1"

		'�]�wSQLŪ���μg�J���b��
		LoginID = "sa"

		'�]�wSQLŪ���μg�J���K�X
		LoginPw = "0000"

		'�]�wSQL��Ʈw�W��
		DataBase = ""
		DataTableAdjustString = ""

	Case Else
		'�����ѰѼ�
		response.write vbNewLine
		response.write "���~:" & vbNewLine
		response.write "CallMain_ConnectionString.asp" & vbNewLine
		response.write "���]�w��ƨӷ��Ҧ�" & vbNewLine
		response.End
End Select

Set DBConnection = Server.CreateObject("ADODB.Connection")
Set RS = Server.CreateObject("ADODB.Recordset")


'�סססס� ��ܸ�ƨӷ� �סססס�
Function FunctionDataSource(DataBaseName)
	If DatabaseName = "" then DatabaseName = DataBase

	Select Case SourceMode
		Case "Excel2000"
			'�]�w Microsoft Excel 2000.xls ����ƨӷ�

				'���Ϳ��~	ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DBQ=" & Server.MapPath(DatabaseName) & ";Extended Properties=Excel 8.0"
			ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DatabaseName) & ";Extended Properties=Excel 8.0"
			DBConnection.Open ConnectionString
			rem RS.Open "[" & SheetName�a���}�Ҹ�Ƥ��W�١A�Шϥ��ܼơb & "$]",DBConnection,?,?


		Case "MSSQL2000"
			'�]�w Microsoft SQL Server 2000 ����ƨӷ�

			ConnectionString ="Provider=SQLOLEDB.1;Server=" & ServerName & ";UID=" & LoginID & ";PWD=" & LoginPw & ";Database=" & DatabaseName
			DBConnection.Open ConnectionString
			rem RS.Open (���}�Ҹ�Ƥ��W�١A�Шϥ��ܼ�),DBConnection,?,?

		Case Else
			'�����ѰѼ�
			response.write vbNewLine
			response.write "���~:" & vbNewLine
			response.write "CallMain_ConnectionString.asp" & vbNewLine
			response.write "���]�w��ƨӷ��Ҧ�" & vbNewLine
			response.End
			
	End Select

End Function
%>