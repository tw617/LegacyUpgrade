<%
FileReNameScriptVer = "2013/09/18"

Dim FsRN, FRNpath, tmpFPath_0, tmpFPath_1, newFsList, newFLPath, newFLName, newCPName

delFName = "AMD64-; AMD64_; IA64-; IA64_; X86-; X86_; all-; all_; zh-tw-"			'�]�w�R�����r���A�ϥ�; �@���j�C
rpcFName = "kb:KB; ie:IE; windows:Windows; dotnetfx:DotNETFx; server:Server; msxml:MsXML; media:Media; xp:XP"			'�]�w���N���r���A�C�@���ϥ�; �@���j�F���N�e���N��ϥ� : �@���j�C
ClassPathName = "Internet Explorer:Pkg-IE; .NET:DotNETFx; C++:Pkg-Plus; Silverlight:Pkg-Plus; Windows:HotFix"			'�]�w�����̾ڪ���Ƨ��W�١A���u�����ǡA�C�@���ϥ�; �@���j�F���N�e���N��ϥ� : �@���j�C
arrangeFolderName = "_arrange"			'��m��z�᪺�ɮפ���Ƨ��W��

FileReName = Request("FileReName")
Radio = Request.Form("Radio")
%>

<html>

<head>
<title>Win7Update2-FileReName</title>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>

<body>
<p>Microsoft Update Catalog �U�����ɮפ��M��C</p>
<form method="POST" action="MsUpdate-2_FileReName.asp">
	<p><textarea rows="10" name="FileReName" tabindex="11" style="width:100%"><%=FileReName%></textarea><br>
	<input type="radio" name="Radio" value="classTable" checked>������<br>
	<input type="radio" name="Radio" value="downBat">�U���ɮ�</p>
	<p><input type="submit" value="�e�X" name="Submit" tabindex="12"><input type="reset" value="���s�]�w" name="Restore"></p>
</form>
<hr>
<%
'�B�z���1: �ˬd�@�P���|
IF FileReName = nul then response.END			'�L��ƴN�����X�C
FsRN = split(FileReName, vbNewLine)			'FsRN = �����ɮ׻P���|���}�C
loopFRNpath = 0			'�ثe�ˬd�`��
Do
  tmpFPath_0 = Split(FsRN(0), Chr(92))(loopFRNpath)			'���Ĥ@���O������ơA�P������ϥΡC
  For loopFRN = 0 to UBound(FsRN)-1 step 1
    tmpFPath_1 = Split(FsRN(loopFRN), Chr(92))(loopFRNpath)			'tmpFPath_1 = ��e��諸���
    
    
    IF NOT tmpFPath_0 = tmpFPath_1 then			'�ˬd��e��諸���ؤ��W�١A���P�h���X
      EXIT DO
    END IF
  Next
  
  
  FRNpath = FRNpath & tmpFPath_0 & Chr(92)			'�@�P���|
  loopFRNpath = loopFRNpath +1
Loop


'�B�z���2: �R���@�P���|�A�üg�^ FsRN() ���C
  For loopFRN = 0 to UBound(FsRN)-1 step 1
    FsRN(loopFRN) = Replace(FsRN(loopFRN), FRNpath, empty)
  Next



'�B�z���3: �L�o�X�i�Ϊ��ɮפ�(�M��B���|���ɦW)�A�t�s newFsList, newFLPath; newFLName ���C
  For loopFRN = 0 to UBound(FsRN)-1 step 1
    IF Instr(FsRN(loopFRN), Chr(92)) > 0 then
      IF vartype(EffectiveFRN) = 0 then
        EffectiveFRN = loopFRN			'���ĲM��_�l����
        ReDim newFsList(UBound(FsRN) - EffectiveFRN)			'���s�ŧi�}�C newFsList ���e�q
        ReDim newFLPath(UBound(FsRN) - EffectiveFRN)			'���s�ŧi�}�C newFLPath ���e�q
        Redim newFLName(UBound(FsRN) - EffectiveFRN)			'���s�ŧi�}�C newFLName ���e�q
        
      END IF
      
      
      newFsList(loopFRN - EffectiveFRN) = FsRN(loopFRN)
      newFLPath(loopFRN - EffectiveFRN) = Split(FsRN(loopFRN), Chr(92))(0)
      newFLName(loopFRN - EffectiveFRN) = Split(FsRN(loopFRN), Chr(92))(1)
      
    END IF
  Next



'�B�z���4: �ק� newFLName ���s�W�١C
  delFArrayName = Split(delFName, Chr(59) & Chr(32))			'�]�w�R�����r�����}�C�A�ϥ�; �@���j�C
  rpcFArrayName = Split(rpcFName, Chr(59) & Chr(32))			'�]�w���N���r�����}�C�A�C�@���ϥ�; �@���j�F���N�e���N��ϥ� : �@���j�C
    
  For loopFRN = 0 to UBound(newFLName)-1 step 1
    newEffectiveName1 = InstrRev(newFLName(loopFRN), Chr(95))			'���o�����ɦW��m
    newEffectiveName2 = InstrRev(newFLName(loopFRN), Chr(46))			'���o���İ��ɦW��m
    newFLName(loopFRN) = LEFT(newFLName(loopFRN), newEffectiveName1 -1) & MID(newFLName(loopFRN), newEffectiveName2)			'���o�ɦW & ���ɦW
    
    
    
    FOR loopDelFRN = 0 to UBound(delFArrayName) Step 1			'�ɦW�����r��
      newFLName(loopFRN) = Replace(newFLName(loopFRN), delFArrayName(loopDelFRN), empty)
    NEXT
    
    
    
    FOR loopRpcFRN = 0 to UBound(rpcFArrayName) Step 1			'�ɦW�����r��
      rpcFArrayNam1 = Split(rpcFArrayName(loopRpcFRN), Chr(58))(0)
      rpcFArrayNam2 = Split(rpcFArrayName(loopRpcFRN), Chr(58))(1)
      newFLName(loopFRN) = Replace(newFLName(loopFRN), rpcFArrayNam1, rpcFArrayNam2)
    NEXT
    
    
    
    IF NOT Instr(newFLName(loopFRN), "KB") > 0 then			'�ɮצW�٤��]�t KB �и��A�h�b�e��[�J KB �и��C
      IF Instr(newFLPath(loopFRN), "KB") > 0 then			'�ˬd���|�W�٦��L�]�t KB �и��A
        tmpStr_KBStart = Instr(newFLPath(loopFRN), "KB")			'�q�ɮ׸��|�Ө��o���D�� KB �и����_�l��m�C
        
        
        For tmpStr_KBEnd = 3 to Len(newFLPath(loopFRN)) Step 1			'�ϥ� ASCii ���o���D�� KB �и����פ��m�C
          tmp = Asc(Mid(newFLPath(loopFRN), tmpStr_KBStart + tmpStr_KBEnd, 1))
          if tmp < 48 or tmp > 57 then Exit For
        Next
        
        newFLName(loopFRN) = mid(newFLPath(loopFRN), tmpStr_KBStart, tmpStr_KBEnd) & Chr(45) & newFLName(loopFRN)			'�N KB �и��g�J newFLName(loopFRN) �}�C�e�C
        
      ELSE
        
        newFLName(loopFRN) = newFLName(loopFRN)			'�L KB �и��A�����a�J�ɦW�C
        
      END IF
    END IF
    
  NEXT



'�B�z���5: �U�ɮפ����C
'ClassPathName			'�]�w�����̾ڪ���Ƨ��W�١A���u�����ǡA�C�@���ϥ�; �@���j�F���N�e���N��ϥ� : �@���j�C

ReDim newCPName(UBound(newFLPath))			'���s�ŧi�}�C newCPName ���e�q
CPArrayName = Split(ClassPathName, Chr(59) & Chr(32))			'�]�w��������Ƨ��W�٤��}�C

  For loopFRN = 0 to UBound(newFLPath) -1 step 1
    
    FOR loopCPNameFRN = 0 to UBound(CPArrayName) Step 1
      
      CPArrayNam1 = Split(CPArrayName(loopCPNameFRN), Chr(58))(0)
      CPArrayNam2 = Split(CPArrayName(loopCPNameFRN), Chr(58))(1)
      
      IF Instr(newFLPath(loopFRN), CStr(CPArrayNam1)) > 0 then
        newCPName(loopFRN) = CPArrayNam2
        EXIT FOR
      END IF
      
    NEXT
    
  Next


Select Case Radio
  Case "classTable"
    '�B�z���6: �̤�����X�ɮײM��C
    
    response.write vbNewLine & "<table border='1' width='100%'><tr>" & vbNewLine
    FOR loopFRN = 0 to UBound(CPArrayName) Step 1
    
      ClassList0 = Split(CPArrayName(loopFRN), Chr(58))(1)
      
      If NOT ClassList0 = ClassList1 then
        response.write "<td align='center'>" & ClassList0 & "<br>"
        response.write "<textarea rows='10' name='ClassList' tabindex='" & loopFRN +1 & "' style='width:100%'>"
        FOR loopNewFLName = 0 to UBound(newFLName) -1 Step 1
          IF newCPName(loopNewFLName) = ClassList0 then response.write newFLName(loopNewFLName) & vbNewLine
        NEXT
        response.write "</textarea>�@" & vbNewLine
      END IF
      
      ClassList1 = ClassList0
      
    NEXT
    
    response.write "</tr></table>"
  
  
  
  Case "downBat"
  '�B�z���7: ��X�妸�ҡC
  'arrangeFolderName		'��m��z�᪺�ɮפ���Ƨ��W��
  'newFsList(loopFRN) = �i�βM��
  'newFLName(loopFRN) = �s�ɦW��
  'newFLPath(loopFRN) = ���|�W��
  'newCPName(loopFRN) = �����W��
  
  '���ͤU���R�O (1/2)
    Response.Clear
    Response.AddHeader "Content-Disposition", "attachment; filename=_MoveFileReName.bat"
    Response.ContentType = "text"
    
    
    response.write "@echo off" & vbNewLine
    
    FOR loopCPNameFRN = 0 to UBound(CPArrayName) Step 1
      response.write "md " & Chr(34) & arrangeFolderName & Chr(92) & Split(CPArrayName(loopCPNameFRN), Chr(58))(1) & Chr(34) & vbNewLine
    NEXT
    
    FOR loopFRN = 0 to UBound(newFsList) -1 step 1
      response.write "move /-Y " & Chr(34) & newFsList(loopFRN) & Chr(34) & Chr(32) & Chr(34) & arrangeFolderName & Chr(92) & newCPName(loopFRN) & Chr(92) & newFLName(loopFRN) & vbNewLine
    NEXT
    
    response.write "exit" & vbNewLine
    
  '���ͤU���R�O (2/2)
    Response.END
  
  End Select
%>
</body>

</html>