<%
'�b�����Ĥ@��[�J�U���o��A�`�N���|�n��A���t�̫e������޸�'
'<!-- #include file="ServerData.asp" -->
'������������������������������
'���A���q���W��
ServerName="vm-ws2k3e-x86-1"

'�]�wSQLŪ���μg�J���b��
LoginID="sa"

'�]�wSQLŪ���μg�J���K�X
LoginPw="0000"

'�]�wSQL��Ʈw�W��
DatabaseName="RulerSystem"
'������������������������������
'��ƪ��W��...

dboSF = "DB_SystemFunction"		'�t�Υ\��
dboSS = "DB_SystemService"		'�t�ΪA��

'������������������������������
'�p����Ʈw���O

Set DBConnection = Server.CreateObject("ADODB.Connection")
ConnectionString ="Provider=SQLOLEDB.1;Server=" & ServerName & ";UID=" & LoginID & ";PWD=" & LoginPw & ";Database=" & DatabaseName
DBConnection.Open ConnectionString
Set RS=Server.CreateObject("ADODB.Recordset")
'RS.Open (���}�Ҹ�Ƥ��W�١A�Шϥ��ܼ�),DBConnection,?,?
'	...
'	...
'RS.Close
'DBConnection.close


'Rs.Open ��ƨӷ��A��Ƴs���A���Ы��A�A��w�覡
'��ƨӷ��G���w��ƪ��W��
'��Ƴs���G���w�@�ӤwConnection������
'
'���Ы��A�G
'���G�u��V�e���ʪ����СA�����w�]��
'���G�L�kŪ����L�ϥΪ̷s�W����ơA��s����Ʒ|�ߧY����
'���G�i�H�Y�ɤ�����L�ϥε۾ާ@��Ʈw�����p
'���G�L�k�Y�ɤ�����L�ϥε۾ާ@�ۦP��Ʈw�����p�A�Ω�j�M�ηs�W�O���ɨϥ�
'
'��w�覡�G
'���G�N Recordset �}�Ҭ���Ū���A�A�����w�]��
'���G���ϥε۹� Recordset �����Y����Ƨ@�s��ɡA�~��w�O��
'���G���ϥε۩I�s Update ��k�� Recordset ����s�ɡA�~��w�O��
'���G�ϥΪ̰��妸��s�ɡA�~��w�O��
'
'
'Options ��ܩʰѼ�[�D���n]�G
'�@�� Long �ȡA���ܴ��Ѫ̦b Source �޼ƥN�� Command ����H�~���F������p��������A�_�h Recordset ���q�e���x�s���ɮ��٭�C
'���i�H�O�U�C�䤤�@�ر`�ơC
'
'�`�ƻ���
'adCmdText�G���Ѫ̷|�N Source ���������O����r�w�q�C
'AdCmdTable�GADO �|���ͤ@�� SQL �d�ߡA�q Source �����w����ƪ��Ǧ^�Ҧ���ƦC�C
'AdCmdTableDirect�G���Ѫ̷|�q Source �����w����ƪ��Ǧ^�Ҧ���ƦC�C
'AdCmdStoredProc�G���Ѫ̷|�N Source �������@�ӹw�s�{�ǡC
'AdCmdUnknown�GSource �޼Ƥ����������O�����C
'AdCommandFile�G�O�d�� (�w�x�s��) Recordset �|�q Source �����w���ɮ��٭�C
'AdExecuteAsync�GSource �@�D�P�B����C
'AdFetchAsync�G���ܦb CacheSize �ݩʤ����w����l�ƶq�Q�����A�ѤU����ƦC�N�|�Q�D�P�B�a����C
%>