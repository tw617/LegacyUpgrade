<%
FileReNameScriptVer = "2013/09/18"

Dim FsRN, FRNpath, tmpFPath_0, tmpFPath_1, newFsList, newFLPath, newFLName, newCPName

delFName = "AMD64-; AMD64_; IA64-; IA64_; X86-; X86_; all-; all_; zh-tw-"			'設定刪除的字元，使用; 作分隔。
rpcFName = "kb:KB; ie:IE; windows:Windows; dotnetfx:DotNETFx; server:Server; msxml:MsXML; media:Media; xp:XP"			'設定取代的字元，每一筆使用; 作分隔；取代前取代後使用 : 作分隔。
ClassPathName = "Internet Explorer:Pkg-IE; .NET:DotNETFx; C++:Pkg-Plus; Silverlight:Pkg-Plus; Windows:HotFix"			'設定分類依據的資料夾名稱，有優先順序，每一筆使用; 作分隔；取代前取代後使用 : 作分隔。
arrangeFolderName = "_arrange"			'放置整理後的檔案之資料夾名稱

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
<p>Microsoft Update Catalog 下載的檔案之清單。</p>
<form method="POST" action="MsUpdate-2_FileReName.asp">
	<p><textarea rows="10" name="FileReName" tabindex="11" style="width:100%"><%=FileReName%></textarea><br>
	<input type="radio" name="Radio" value="classTable" checked>分類表<br>
	<input type="radio" name="Radio" value="downBat">下載檔案</p>
	<p><input type="submit" value="送出" name="Submit" tabindex="12"><input type="reset" value="重新設定" name="Restore"></p>
</form>
<hr>
<%
'處理資料1: 檢查共同路徑
IF FileReName = nul then response.END			'無資料就停止輸出。
FsRN = split(FileReName, vbNewLine)			'FsRN = 全部檔案與路徑之陣列
loopFRNpath = 0			'目前檢查深度
Do
  tmpFPath_0 = Split(FsRN(0), Chr(92))(loopFRNpath)			'取第一筆記錄的資料，與後方比較使用。
  For loopFRN = 0 to UBound(FsRN)-1 step 1
    tmpFPath_1 = Split(FsRN(loopFRN), Chr(92))(loopFRNpath)			'tmpFPath_1 = 當前比對的資料
    
    
    IF NOT tmpFPath_0 = tmpFPath_1 then			'檢查當前比對的項目之名稱，不同則跳出
      EXIT DO
    END IF
  Next
  
  
  FRNpath = FRNpath & tmpFPath_0 & Chr(92)			'共同路徑
  loopFRNpath = loopFRNpath +1
Loop


'處理資料2: 刪除共同路徑，並寫回 FsRN() 中。
  For loopFRN = 0 to UBound(FsRN)-1 step 1
    FsRN(loopFRN) = Replace(FsRN(loopFRN), FRNpath, empty)
  Next



'處理資料3: 過濾出可用的檔案之(清單、路徑及檔名)，另存 newFsList, newFLPath; newFLName 中。
  For loopFRN = 0 to UBound(FsRN)-1 step 1
    IF Instr(FsRN(loopFRN), Chr(92)) > 0 then
      IF vartype(EffectiveFRN) = 0 then
        EffectiveFRN = loopFRN			'有效清單起始筆數
        ReDim newFsList(UBound(FsRN) - EffectiveFRN)			'重新宣告陣列 newFsList 的容量
        ReDim newFLPath(UBound(FsRN) - EffectiveFRN)			'重新宣告陣列 newFLPath 的容量
        Redim newFLName(UBound(FsRN) - EffectiveFRN)			'重新宣告陣列 newFLName 的容量
        
      END IF
      
      
      newFsList(loopFRN - EffectiveFRN) = FsRN(loopFRN)
      newFLPath(loopFRN - EffectiveFRN) = Split(FsRN(loopFRN), Chr(92))(0)
      newFLName(loopFRN - EffectiveFRN) = Split(FsRN(loopFRN), Chr(92))(1)
      
    END IF
  Next



'處理資料4: 修改 newFLName 的新名稱。
  delFArrayName = Split(delFName, Chr(59) & Chr(32))			'設定刪除的字元之陣列，使用; 作分隔。
  rpcFArrayName = Split(rpcFName, Chr(59) & Chr(32))			'設定取代的字元之陣列，每一筆使用; 作分隔；取代前取代後使用 : 作分隔。
    
  For loopFRN = 0 to UBound(newFLName)-1 step 1
    newEffectiveName1 = InstrRev(newFLName(loopFRN), Chr(95))			'取得有效檔名位置
    newEffectiveName2 = InstrRev(newFLName(loopFRN), Chr(46))			'取得有效副檔名位置
    newFLName(loopFRN) = LEFT(newFLName(loopFRN), newEffectiveName1 -1) & MID(newFLName(loopFRN), newEffectiveName2)			'取得檔名 & 副檔名
    
    
    
    FOR loopDelFRN = 0 to UBound(delFArrayName) Step 1			'檔名替除字樣
      newFLName(loopFRN) = Replace(newFLName(loopFRN), delFArrayName(loopDelFRN), empty)
    NEXT
    
    
    
    FOR loopRpcFRN = 0 to UBound(rpcFArrayName) Step 1			'檔名替換字樣
      rpcFArrayNam1 = Split(rpcFArrayName(loopRpcFRN), Chr(58))(0)
      rpcFArrayNam2 = Split(rpcFArrayName(loopRpcFRN), Chr(58))(1)
      newFLName(loopFRN) = Replace(newFLName(loopFRN), rpcFArrayNam1, rpcFArrayNam2)
    NEXT
    
    
    
    IF NOT Instr(newFLName(loopFRN), "KB") > 0 then			'檔案名稱不包含 KB 標號，則在前方加入 KB 標號。
      IF Instr(newFLPath(loopFRN), "KB") > 0 then			'檢查路徑名稱有無包含 KB 標號，
        tmpStr_KBStart = Instr(newFLPath(loopFRN), "KB")			'從檔案路徑來取得標題的 KB 標號之起始位置。
        
        
        For tmpStr_KBEnd = 3 to Len(newFLPath(loopFRN)) Step 1			'使用 ASCii 取得標題的 KB 標號之終止位置。
          tmp = Asc(Mid(newFLPath(loopFRN), tmpStr_KBStart + tmpStr_KBEnd, 1))
          if tmp < 48 or tmp > 57 then Exit For
        Next
        
        newFLName(loopFRN) = mid(newFLPath(loopFRN), tmpStr_KBStart, tmpStr_KBEnd) & Chr(45) & newFLName(loopFRN)			'將 KB 標號寫入 newFLName(loopFRN) 陣列前。
        
      ELSE
        
        newFLName(loopFRN) = newFLName(loopFRN)			'無 KB 標號，直接帶入檔名。
        
      END IF
    END IF
    
  NEXT



'處理資料5: 各檔案分類。
'ClassPathName			'設定分類依據的資料夾名稱，有優先順序，每一筆使用; 作分隔；取代前取代後使用 : 作分隔。

ReDim newCPName(UBound(newFLPath))			'重新宣告陣列 newCPName 的容量
CPArrayName = Split(ClassPathName, Chr(59) & Chr(32))			'設定分類的資料夾名稱之陣列

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
    '處理資料6: 依分類輸出檔案清單。
    
    response.write vbNewLine & "<table border='1' width='100%'><tr>" & vbNewLine
    FOR loopFRN = 0 to UBound(CPArrayName) Step 1
    
      ClassList0 = Split(CPArrayName(loopFRN), Chr(58))(1)
      
      If NOT ClassList0 = ClassList1 then
        response.write "<td align='center'>" & ClassList0 & "<br>"
        response.write "<textarea rows='10' name='ClassList' tabindex='" & loopFRN +1 & "' style='width:100%'>"
        FOR loopNewFLName = 0 to UBound(newFLName) -1 Step 1
          IF newCPName(loopNewFLName) = ClassList0 then response.write newFLName(loopNewFLName) & vbNewLine
        NEXT
        response.write "</textarea>　" & vbNewLine
      END IF
      
      ClassList1 = ClassList0
      
    NEXT
    
    response.write "</tr></table>"
  
  
  
  Case "downBat"
  '處理資料7: 輸出批次黨。
  'arrangeFolderName		'放置整理後的檔案之資料夾名稱
  'newFsList(loopFRN) = 可用清單
  'newFLName(loopFRN) = 新檔名稱
  'newFLPath(loopFRN) = 路徑名稱
  'newCPName(loopFRN) = 分類名稱
  
  '產生下載命令 (1/2)
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
    
  '產生下載命令 (2/2)
    Response.END
  
  End Select
%>
</body>

</html>