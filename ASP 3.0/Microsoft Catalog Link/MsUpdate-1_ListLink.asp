q<%
ListLinkScriptVer = "2013/09/17"
SearchURL = "http://catalog.update.microsoft.com/v7/site/Search.aspx?q="


Dim tmpUdL, tmpStr, UdL

'tmpUdL 設計成全域變數的一維陣列。
'  將 UpdateList 切割之後的每一筆資料寫入陣列 tmpUdL(x) 中。

'tmpStr 設計成區域變數的字串。
'  將 tmpUdL(x) 切割，只取 tmpUdL(x)(0) 存入區域變數 tmpStr 中。
'  ex:
'    tmpStr = "KB2425227：x64 系統的 Windows 7 安全性更新"

'UdL 設計成二維陣列。
'  處理 tmpStr 的 KB 標號及 Title 名稱之後，存入 UdL(x,y) 中。
'  資料範例：
'    UdL(?,0)=KB2425227
'    UdL(?,1)=x64 系統的 Windows 7 安全性更新


UpdateList = Request("UpdateList")	'直接複製 Windows Update 上的更新清單
radioOS = Request.Form("radioOS")
%>

<html>

<head>
<title>MsUpdate-1_ListLink</title>
</head>

<body>
<p> 『Microsoft Update 的更新清單』轉換成『下載捷徑』。<br>
　　【直接貼上 Windows Update 的更新列表】</p>
<form method="POST" action="MsUpdate-1_ListLink.asp">
	<p><textarea rows="10" name="UpdateList" cols="81" tabindex="1"><%=UpdateList%></textarea></p>
	<p>篩選方法：<br>
	<input type="radio" name="radioOS" value="ver5.1x32">Microsoft Windows XP x32<br>
	<input type="radio" name="radioOS" value="ver5.1x64">Microsoft Windows XP x64 Edition<br>
	<input type="radio" name="radioOS" value="ver5.2R2x32">Microsoft Windows Server 2003 x32<br>
	<input type="radio" name="radioOS" value="ver5.2R2x64">Microsoft Windows Server 2003 x64</p>
	<input type="radio" name="radioOS" value="ver6.1">Windows 7 x32, x64<br>
	<input type="radio" name="radioOS" value="ver6.2x64">Windows Server 2008 x64</p>
	<p><input type="submit" value="送出" name="Submit" tabindex="2"><input type="reset" value="重新設定" name="Restore"></p>
</form>
<hr>

<p><a href="http://catalog.update.microsoft.com/">Microsoft Update Catalog</a><br>
<a href="http://catalog.update.microsoft.com/v7/site/ViewBasket.aspx">下載籃</a></p>
<p>
<%

'處理資料1: 將 UpdateList 切割之後的每一筆資料寫入陣列 tmpUdL(x) 中。

Select Case radioOS
      
    Case "ver5.1x32", "ver5.1x64", "ver5.2R2x32", "ver5.2R2x64"
      
      Select Case True	//去除表頭
          
          Case InStr(UpdateList , "高優先順序的更新 ") > 0
              UpdateList = split(UpdateList, "高優先順序的更新 " & vbnewline)(1)
              
          Case InStr(UpdateList , "選用的軟體更新 ") > 0
              UpdateList = split(UpdateList, "選用的軟體更新 " & vbnewline)(1)
              
          Case ELSE
              Response.write "<p>『找不到分類索引值，終止輸出！』<br>"
              Response.write "ex:<br>"
              Response.write "『高優先順序的更新 』<br>"
              Response.write "『選用的軟體更新 』</p>"
              response.END
              
          End Select
          
      Select Case radioOS	//去除字樣。該字樣會被輸出時，才需除掉！
          
          Case "ver5.1x32"
              UpdateList = Replace(UpdateList, "Microsoft Windows XP" & vbNewLine & vbNewLine, vbNewLine, 1, 2, 1)
              
          Case "ver5.1x64"
              UpdateList = Replace(UpdateList, "Microsoft Windows XP x64 Edition" & vbNewLine & vbNewLine, vbNewLine, 1, 2, 1)
              
          Case "ver5.2R2x32"
              UpdateList = Replace(UpdateList, "Microsoft Windows Server 2003" & vbNewLine & vbNewLine, vbNewLine, 1, 2, 1)
              
          Case "ver5.2R2x64"
              UpdateList = Replace(UpdateList, "Microsoft Windows Server 2003" & vbNewLine & vbNewLine, vbNewLine, 1, 2, 1)
              
          Case ELSE
              Response.write "『radioOS 異常，終止輸出。』"
              response.END
              
      End Select
      
      tmpUdL = split(UpdateList, vbnewline & Chr(32) & vbnewline)
      
      
    Case "ver6.1"
      tmpUdL = split(UpdateList, vbnewline & vbnewline & vbnewline & vbnewline)
      
      
      
    Case ELSE
      Response.write "『radioOS 異常，終止輸出。』"
      response.END
    
END Select



'處理資料2:
'  將 tmpUdL(x) 切割，只取 tmpUdL(x)(0) 存入區域變數 tmpStr 中。

'  UdL 設計成二維陣列。
'    處理 tmpStr 的 KB 標號及 Title 名稱之後，存入 UdL(x,y) 中。
'    資料範例：
'      UdL(?,0)=KB2425227
'      UdL(?,1)=x64 系統的 Windows 7 安全性更新

ReDim UdL(UBound(tmpUdL), 1)	'宣告擴展 UdL 陣列。

For tmpLoop = 0 to UBound(tmpUdL) step 1
  tmpStr = split(tmpUdL(tmpLoop), vbnewline)(0)			'取得陣列中，各筆的標題文字。
  
  '處理資料2-1: 取 KB 標號及位置。
  tmpStr_KBStart = Instr(tmpStr, "KB")			'取得標題的 KB 標號之起始位置。
  
  For tmpStr_KBEnd = 3 to Len(tmpStr) Step 1			'使用 ASCii 取得標題的 KB 標號之終止位置。
    tmp = Asc(Mid(tmpStr, tmpStr_KBStart + tmpStr_KBEnd, 1))
    if tmp < 48 or tmp > 57 then Exit For
  Next
  
  
  IF tmpStr_KBStart > 0 then			'篩選 KB 標號
    UdL(tmpLoop, 0) = mid(tmpStr, tmpStr_KBStart, tmpStr_KBEnd)			'將 tmpStr 寫入 UDL 陣列中。
  ELSE
    UdL(tmpLoop, 0) = tmpStr
  End IF
  
  
  
  '處理資料2-2: 取 Title 名稱。
  
  IF UdL(tmpLoop, 0) = LEFT(tmpStr, LEN(UdL(tmpLoop, 0))) then			'判斷 KB 標號是否在最前端
    
    UdL(tmpLoop, 1) = MID(tmpStr, Len(UdL(tmpLoop, 0)) +1 )			'UdL(tmpLoop, 1) = KB 標號之後的字。
    
  ELSE
    
    tmpStr_KBStart = Instr(tmpStr, UdL(tmpLoop, 0))			'取得標題的 KB 標號之起始位置。
    UdL(tmpLoop, 1) = Left(tmpStr, tmpStr_KBStart -1 )			'UdL(tmpLoop, 1) = KB 標號之前的字。
    
  END IF
  
  
  '特殊限制的項目
  IF Instr(UdL(tmpLoop, 1), "C++") > 0 then UdL(tmpLoop, 1) = LEFT(Chr(32) & UdL(tmpLoop, 1), +17 )			'Microsoft Visual 的項目
  IF Instr(UdL(tmpLoop, 1), ".NET") > 0 then			'Microsoft .NET Framework 的項目，過濾 x86 / x64 字樣。
    IF Instr(UdL(tmpLoop, 1), "x64") > 0 then
      UdL(tmpLoop, 1) = " x64"
    ELSE
      UdL(tmpLoop, 1) = " x86"
    END IF
  END IF
  
  
  IF UdL(tmpLoop, 1) = Empty then UdL(tmpLoop, 1) = tmpStr
  
  
  
  
  
  '2016/07/16 強制將副標題改用「x64 7」字樣
  UdL(tmpLoop, 1) = " x64 7"
  
  
  
  
  UdL(tmpLoop, 1) = UdL(tmpLoop, 0) & UdL(tmpLoop, 1)
  
  
  
Next





'輸出URL：
Session.CodePage="65001"	'指定輸出成 UTF-8
tmp = "http://catalog.update.microsoft.com/" & vbnewline & "http://catalog.update.microsoft.com/v7/site/ViewBasket.aspx" & vbnewline
For tmpLoop = 0 to UBound(UdL,1) step 1
  If NOT UdL(tmpLoop, 0) = nul then
    tmp = tmp & SearchURL & Server.URLEncode(Left(UdL(tmpLoop, 1), 65)) & vbnewline
  END If
Next
Session.CodePage="950"	'指定輸出成 Big5
response.write "<p><textarea rows='2' name='S1' cols='20'>" & tmp & "</textarea></p>"




'輸出總筆數：
response.write "<p>共有 " & UBound(UDL)+1 & " 筆資料。</p>" & vbnewLine




'輸出資料：
Session.CodePage="65001"	'指定輸出成 UTF-8

response.write "<p>"
For tmpLoop = 0 to UBound(UdL,1) step 1
  If NOT UdL(tmpLoop, 0) = nul then
    response.write "<a href='" & SearchURL & Server.URLEncode(Left(UdL(tmpLoop, 1), 80)) & "'>" & tmpLoop & ": " & UdL(tmpLoop, 0) & "</a><br>" & vbnewline
  END If
Next
response.write "</p>"

Session.CodePage="950"	'指定輸出成 Big5

%>
</p>
</body>

</html>
