q<%
ListLinkScriptVer = "2013/09/17"
SearchURL = "http://catalog.update.microsoft.com/v7/site/Search.aspx?q="


Dim tmpUdL, tmpStr, UdL

'tmpUdL �]�p�������ܼƪ��@���}�C�C
'  �N UpdateList ���Τ��᪺�C�@����Ƽg�J�}�C tmpUdL(x) ���C

'tmpStr �]�p���ϰ��ܼƪ��r��C
'  �N tmpUdL(x) ���ΡA�u�� tmpUdL(x)(0) �s�J�ϰ��ܼ� tmpStr ���C
'  ex:
'    tmpStr = "KB2425227�Gx64 �t�Ϊ� Windows 7 �w���ʧ�s"

'UdL �]�p���G���}�C�C
'  �B�z tmpStr �� KB �и��� Title �W�٤���A�s�J UdL(x,y) ���C
'  ��ƽd�ҡG
'    UdL(?,0)=KB2425227
'    UdL(?,1)=x64 �t�Ϊ� Windows 7 �w���ʧ�s


UpdateList = Request("UpdateList")	'�����ƻs Windows Update �W����s�M��
radioOS = Request.Form("radioOS")
%>

<html>

<head>
<title>MsUpdate-1_ListLink</title>
</head>

<body>
<p> �yMicrosoft Update ����s�M��z�ഫ���y�U�����|�z�C<br>
�@�@�i�����K�W Windows Update ����s�C��j</p>
<form method="POST" action="MsUpdate-1_ListLink.asp">
	<p><textarea rows="10" name="UpdateList" cols="81" tabindex="1"><%=UpdateList%></textarea></p>
	<p>�z���k�G<br>
	<input type="radio" name="radioOS" value="ver5.1x32">Microsoft Windows XP x32<br>
	<input type="radio" name="radioOS" value="ver5.1x64">Microsoft Windows XP x64 Edition<br>
	<input type="radio" name="radioOS" value="ver5.2R2x32">Microsoft Windows Server 2003 x32<br>
	<input type="radio" name="radioOS" value="ver5.2R2x64">Microsoft Windows Server 2003 x64</p>
	<input type="radio" name="radioOS" value="ver6.1">Windows 7 x32, x64<br>
	<input type="radio" name="radioOS" value="ver6.2x64">Windows Server 2008 x64</p>
	<p><input type="submit" value="�e�X" name="Submit" tabindex="2"><input type="reset" value="���s�]�w" name="Restore"></p>
</form>
<hr>

<p><a href="http://catalog.update.microsoft.com/">Microsoft Update Catalog</a><br>
<a href="http://catalog.update.microsoft.com/v7/site/ViewBasket.aspx">�U���x</a></p>
<p>
<%

'�B�z���1: �N UpdateList ���Τ��᪺�C�@����Ƽg�J�}�C tmpUdL(x) ���C

Select Case radioOS
      
    Case "ver5.1x32", "ver5.1x64", "ver5.2R2x32", "ver5.2R2x64"
      
      Select Case True	//�h�����Y
          
          Case InStr(UpdateList , "���u�����Ǫ���s ") > 0
              UpdateList = split(UpdateList, "���u�����Ǫ���s " & vbnewline)(1)
              
          Case InStr(UpdateList , "��Ϊ��n���s ") > 0
              UpdateList = split(UpdateList, "��Ϊ��n���s " & vbnewline)(1)
              
          Case ELSE
              Response.write "<p>�y�䤣��������ޭȡA�פ��X�I�z<br>"
              Response.write "ex:<br>"
              Response.write "�y���u�����Ǫ���s �z<br>"
              Response.write "�y��Ϊ��n���s �z</p>"
              response.END
              
          End Select
          
      Select Case radioOS	//�h���r�ˡC�Ӧr�˷|�Q��X�ɡA�~�ݰ����I
          
          Case "ver5.1x32"
              UpdateList = Replace(UpdateList, "Microsoft Windows XP" & vbNewLine & vbNewLine, vbNewLine, 1, 2, 1)
              
          Case "ver5.1x64"
              UpdateList = Replace(UpdateList, "Microsoft Windows XP x64 Edition" & vbNewLine & vbNewLine, vbNewLine, 1, 2, 1)
              
          Case "ver5.2R2x32"
              UpdateList = Replace(UpdateList, "Microsoft Windows Server 2003" & vbNewLine & vbNewLine, vbNewLine, 1, 2, 1)
              
          Case "ver5.2R2x64"
              UpdateList = Replace(UpdateList, "Microsoft Windows Server 2003" & vbNewLine & vbNewLine, vbNewLine, 1, 2, 1)
              
          Case ELSE
              Response.write "�yradioOS ���`�A�פ��X�C�z"
              response.END
              
      End Select
      
      tmpUdL = split(UpdateList, vbnewline & Chr(32) & vbnewline)
      
      
    Case "ver6.1"
      tmpUdL = split(UpdateList, vbnewline & vbnewline & vbnewline & vbnewline)
      
      
      
    Case ELSE
      Response.write "�yradioOS ���`�A�פ��X�C�z"
      response.END
    
END Select



'�B�z���2:
'  �N tmpUdL(x) ���ΡA�u�� tmpUdL(x)(0) �s�J�ϰ��ܼ� tmpStr ���C

'  UdL �]�p���G���}�C�C
'    �B�z tmpStr �� KB �и��� Title �W�٤���A�s�J UdL(x,y) ���C
'    ��ƽd�ҡG
'      UdL(?,0)=KB2425227
'      UdL(?,1)=x64 �t�Ϊ� Windows 7 �w���ʧ�s

ReDim UdL(UBound(tmpUdL), 1)	'�ŧi�X�i UdL �}�C�C

For tmpLoop = 0 to UBound(tmpUdL) step 1
  tmpStr = split(tmpUdL(tmpLoop), vbnewline)(0)			'���o�}�C���A�U�������D��r�C
  
  '�B�z���2-1: �� KB �и��Φ�m�C
  tmpStr_KBStart = Instr(tmpStr, "KB")			'���o���D�� KB �и����_�l��m�C
  
  For tmpStr_KBEnd = 3 to Len(tmpStr) Step 1			'�ϥ� ASCii ���o���D�� KB �и����פ��m�C
    tmp = Asc(Mid(tmpStr, tmpStr_KBStart + tmpStr_KBEnd, 1))
    if tmp < 48 or tmp > 57 then Exit For
  Next
  
  
  IF tmpStr_KBStart > 0 then			'�z�� KB �и�
    UdL(tmpLoop, 0) = mid(tmpStr, tmpStr_KBStart, tmpStr_KBEnd)			'�N tmpStr �g�J UDL �}�C���C
  ELSE
    UdL(tmpLoop, 0) = tmpStr
  End IF
  
  
  
  '�B�z���2-2: �� Title �W�١C
  
  IF UdL(tmpLoop, 0) = LEFT(tmpStr, LEN(UdL(tmpLoop, 0))) then			'�P�_ KB �и��O�_�b�̫e��
    
    UdL(tmpLoop, 1) = MID(tmpStr, Len(UdL(tmpLoop, 0)) +1 )			'UdL(tmpLoop, 1) = KB �и����᪺�r�C
    
  ELSE
    
    tmpStr_KBStart = Instr(tmpStr, UdL(tmpLoop, 0))			'���o���D�� KB �и����_�l��m�C
    UdL(tmpLoop, 1) = Left(tmpStr, tmpStr_KBStart -1 )			'UdL(tmpLoop, 1) = KB �и����e���r�C
    
  END IF
  
  
  '�S�������
  IF Instr(UdL(tmpLoop, 1), "C++") > 0 then UdL(tmpLoop, 1) = LEFT(Chr(32) & UdL(tmpLoop, 1), +17 )			'Microsoft Visual ������
  IF Instr(UdL(tmpLoop, 1), ".NET") > 0 then			'Microsoft .NET Framework �����ءA�L�o x86 / x64 �r�ˡC
    IF Instr(UdL(tmpLoop, 1), "x64") > 0 then
      UdL(tmpLoop, 1) = " x64"
    ELSE
      UdL(tmpLoop, 1) = " x86"
    END IF
  END IF
  
  
  IF UdL(tmpLoop, 1) = Empty then UdL(tmpLoop, 1) = tmpStr
  
  
  
  
  
  '2016/07/16 �j��N�Ƽ��D��Ρux64 7�v�r��
  UdL(tmpLoop, 1) = " x64 7"
  
  
  
  
  UdL(tmpLoop, 1) = UdL(tmpLoop, 0) & UdL(tmpLoop, 1)
  
  
  
Next





'��XURL�G
Session.CodePage="65001"	'���w��X�� UTF-8
tmp = "http://catalog.update.microsoft.com/" & vbnewline & "http://catalog.update.microsoft.com/v7/site/ViewBasket.aspx" & vbnewline
For tmpLoop = 0 to UBound(UdL,1) step 1
  If NOT UdL(tmpLoop, 0) = nul then
    tmp = tmp & SearchURL & Server.URLEncode(Left(UdL(tmpLoop, 1), 65)) & vbnewline
  END If
Next
Session.CodePage="950"	'���w��X�� Big5
response.write "<p><textarea rows='2' name='S1' cols='20'>" & tmp & "</textarea></p>"




'��X�`���ơG
response.write "<p>�@�� " & UBound(UDL)+1 & " ����ơC</p>" & vbnewLine




'��X��ơG
Session.CodePage="65001"	'���w��X�� UTF-8

response.write "<p>"
For tmpLoop = 0 to UBound(UdL,1) step 1
  If NOT UdL(tmpLoop, 0) = nul then
    response.write "<a href='" & SearchURL & Server.URLEncode(Left(UdL(tmpLoop, 1), 80)) & "'>" & tmpLoop & ": " & UdL(tmpLoop, 0) & "</a><br>" & vbnewline
  END If
Next
response.write "</p>"

Session.CodePage="950"	'���w��X�� Big5

%>
</p>
</body>

</html>
