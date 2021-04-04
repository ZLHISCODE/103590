VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LX5000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:尿沉渣LX500
'厂商:灵丘县医院
'医院：双流县中医院
'方式：SQLserver数据库提取方式
'提取文件：Examdir_Data.mdf数据库中的ITEMLIST表
'作者：陈德荣(曾超在LX7680基础上修改)

Private FileNames As String

Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
 dwLowDateTime As Long
 dwHighDateTime As Long
End Type
 
Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Dim fso As New FileSystemObject
Dim fld As Folder

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Function StripNulls(OriginalStr As String) As String
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, _
      InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function

Public Function FindFilesAPI(ByVal Path As String, ByVal SearchStr As String)
   Dim Filename As String   ' Walking filename variable...
   Dim DirName As String    ' 子文件名称
   Dim dirNames() As String ' Buffer for directory name entries
   Dim nDir As Integer   ' Number of directories in this path
   Dim i As Integer
   Dim hSearch As Long   ' 查找句柄
   Dim WFD As WIN32_FIND_DATA
   Dim Cont As Integer
   Dim FT As FILETIME
   Dim ST As SYSTEMTIME
   Dim DateCStr As String, DateMStr As String
      
   If Right(Path, 1) <> "\" Then Path = Path & "\"
   ' Search for subdirectories.
   nDir = 0
   ReDim dirNames(nDir)
   Cont = True
   hSearch = FindFirstFile(Path & "*", WFD)
   If hSearch <> INVALID_HANDLE_VALUE Then
      Do While Cont
         DirName = StripNulls(WFD.cFileName)
         If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
               dirNames(nDir) = DirName
               nDir = nDir + 1
               ReDim Preserve dirNames(nDir)
            End If
         End If
         Cont = FindNextFile(hSearch, WFD) ' Get next subdirectory.
      Loop
      Cont = FindClose(hSearch)
   End If
 
   ' Walk through this directory and sum file sizes.
   hSearch = FindFirstFile(Path & SearchStr, WFD)
   Cont = True
   If hSearch <> INVALID_HANDLE_VALUE Then
      While Cont
         Filename = StripNulls(WFD.cFileName)
            If (Filename <> ".") And (Filename <> "..") And _
              ((GetFileAttributes(Path & Filename) And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
            FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
             MAXDWORD) + WFD.nFileSizeLow
             
           '文件日期信息
           FileTimeToLocalFileTime WFD.ftCreationTime, FT
           FileTimeToSystemTime FT, ST
           DateCStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
              " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
           ' and Last Modified Date
           FileTimeToLocalFileTime WFD.ftLastWriteTime, FT
           FileTimeToSystemTime FT, ST
           DateMStr = ST.wMonth & "/" & ST.wDay & "/" & ST.wYear & _
              " " & ST.wHour & ":" & ST.wMinute & ":" & ST.wSecond
             
           If FileNames <> "" Then FileNames = FileNames + "|"
           FileNames = FileNames & Filename
           'List1.AddItem Path & FileName & vbTab & _
           '   Format(DateCStr, "mm/dd/yyyy hh:nn:ss") _
           '   & vbTab & Format(DateMStr, "mm/dd/yyyy hh:nn:ss")
          End If
         Cont = FindNextFile(hSearch, WFD) '查找下一个文件
      Wend
      Cont = FindClose(hSearch)
   End If
 
    '遍历子文件夹
    If nDir > 0 Then
      ' Recursively walk into them...
      For i = 0 To nDir - 1
        FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) _
         & "\", SearchStr)
      Next i
   End If
End Function

Private Function FindFile(ByVal sFol As String, sFile As String) As Currency
   Dim tFld As Folder, tFil As File, Filename As String
    
   On Error GoTo Catch
   Set fld = fso.GetFolder(sFol)
   Filename = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
                  vbHidden Or vbSystem Or vbReadOnly)
   While Len(Filename) <> 0
      FindFile = FindFile + FileLen(fso.BuildPath(fld.Path, _
      Filename))
      
      If FileNames <> "" Then FileNames = FileNames + "|"
      FileNames = FileNames & fso.BuildPath("", Filename)
      
      Filename = Dir() ' Get next file
      DoEvents
   Wend
   If fld.SubFolders.count > 0 Then
      For Each tFld In fld.SubFolders
         DoEvents
         FindFile = FindFile + FindFile(tFld.Path, sFile)
      Next
   End If
   Exit Function
Catch: Filename = ""
       Resume Next
End Function

'调试说明:
'1.通讯名:clsLISDev_File_LX5000
'2.由于需要一些环境变量所以增加了配置文件，第一次调用这个接口时会在当前目录下产生一个LX5000.ini的配置文件
'配置说明:
'  Test:0 0=正常模式 1=调试模式 调试模式会在c盘下增加一个LisLog_LX5000_时间.txt的日志文件
'DatabaseName:  ExamDir 数据库名
'  SERVERName:(local)    服务器名
'USERName:  sa 用户名
'PASSWORD:       密码
'PicPath:        图像路径
'PicFix:         图像类型 (不需要修改)
'3.由于项目名都为中文名需要转换为英文名
'  红细胞计数 rbc
'  白细胞总数 TWBC
'  上皮细胞计数 spxb
'  病理性管型 blxgx
'  非晶形尿酸结晶 nsjj
'  透明管型 tmgx
'  粘液丝 nys
'  滴虫 dc
'  真菌 zj
'  细菌 xj
'  结晶 jj
'  Unknown kong
'  草酸钙结晶 csgjj
'  精子 jz
'  颗粒管型 klgx
'4.产生配置文件后，修改Test = 1为调试模式。如有问题可以查看C盘下的日志进行判断



Private Function FindFile2(ByVal sPath As String, sTyp As String) As String
   Dim aTmp As String, bTmp As String
   aTmp = Dir(sPath & sTyp)   ' 找寻第一项。
   bTmp = ""
   Do While aTmp <> ""   '开始循环。
     ' 跳过当前的目录及上层目录。
     If aTmp <> "." And aTmp <> ".." Then
       If bTmp <> "" Then bTmp = bTmp & "|"
         bTmp = bTmp & aTmp   ' 如果它是一个目录，将其名称显示出来。
     End If
    aTmp = Dir   ' 查找下一个目录。
   Loop
   FindFile2 = bTmp
End Function

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'标准函数：从文件获取检验结果
'   strFile：包含路径的文件名
'   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
'   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
'   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
'   返回的检验结果数组(各解析程序必须按以下标准组织结果)
'   元素之间以|分隔
'   第0个元素：检验时间
'   第1个元素：样本序号
'   第2个元素：检验人
'   第3个元素：标本
'   第4个元素：是否质控品
'   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
'       如：第5i个元素为检验项目，第5i+1个元素为检验结果
' Dim DBConn As New ADODB.Connection
'打开SQL数据库
  ' DBConn.Open "Provider   =   sqloledb;Data   Source= '服务器名 ';
   'Initial   Catalog= '数据库名称 ';User   Id= '用户名 ';Password= '密码 ' "
   '如不想将连接字符串写在程序里的话，还可以通过读INI文件，例如配置内容如下:
    '检验技师工作站指定文件时只需指定这个ini文件即可
    '[LX2000]
    'Database Name = ExamDir
    'SERVER Name = (local)
    'USER Name = sa
    'PASSWORD=
    Dim aResults() As String
    Dim ConnSQL As New ADODB.Connection
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim lngResultCount As Long, strBh As String
    Dim str日期 As String, strImg As String, strTmp As String
    Dim strTDM As String, strPicPath As String, strPicFix As String, strTmp2 As String, strTmp3 As String
    Dim intloop As Integer
    Dim aFiles() As String, strLog As String, strImgBmp As String, YearMonth As String, YearMonthDay As String, YearMonthSample As String
    
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream
    
    Dim intDebug As Integer
    
    
    Dim DatabaseName As String
    Dim SERVERName As String
    Dim USERName As String
    Dim PASSWORD As String
    
    
'    strLog = "已经调用部件"
'    Call writelislog(strLog)

    On Error GoTo errH
    
    strIniFile = App.Path & "\LX5000.ini"
    
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Debug]"
       objStream.WriteLine "Test=0"     '1＝调试模式会产生日志 0=正常模式不会产生日志
       
       objStream.WriteLine "[Conn]"
       objStream.WriteLine "DatabaseName=ExamDir"
       objStream.WriteLine "SERVERName=(local)"
       objStream.WriteLine "USERName=sa"
       objStream.WriteLine "PASSWORD="
       
       objStream.WriteLine "[Options]"
       objStream.WriteLine "PicPath=C:\Program Files\LX5000\Images"
       objStream.WriteLine "PicFix=U100N"
       objStream.Close
       Set objStream = Nothing
       strPicPath = "C:\Program Files\LX5000\Images"
       intDebug = 0
       DatabaseName = "ExamDir"
       SERVERName = "(local)"
       USERName = "sa"
       PASSWORD = ""
       
    Else
       strPicPath = ReadIni("Options", "PicPath", strIniFile)
       strPicFix = ReadIni("Options", "PicFix", strIniFile)
       If (Len(strPicPath) > 0) Then
         If Mid(strPicPath, Len(strPicPath)) <> "\" Then strPicPath = strPicPath + "\"
       End If
       
       intDebug = ReadIni("Debug", "Test", strIniFile, 0)
       DatabaseName = ReadIni("Conn", "DatabaseName", strIniFile, "ExamDir")
       SERVERName = ReadIni("Conn", "SERVERName", strIniFile, "(local)")
       USERName = ReadIni("Conn", "USERName", strIniFile, "sa")
       PASSWORD = ReadIni("Conn", "PASSWORD", strIniFile, "")
    End If
    
    With ConnSQL
        .ConnectionString = "Provider=SQLOLEDB;Data Source=" & SERVERName & ";user id=" & USERName & ";password=" & PASSWORD & ";initial catalog=" & DatabaseName
        .Open
        If .State = adStateClosed Then
          If intDebug = 1 Then
            strLog = "连接SQL SERVER数据库失败。"
            Call writelislog(strLog)
          End If
          Exit Function
        End If
        
        If intDebug = 1 Then
            strLog = "连接SQL SERVER数据库成功。"
            Call writelislog(strLog)
        End If
        strSQL = "SELECT EXAMINEDATE, SAMPLETYPE,SAMPLEID,ITEMCODE, ITEMVALUE,ITEMCHAR,QUERYID " & vbNewLine & _
                "From ITEMLIST " & vbNewLine & _
                "Where EXAMINEDATE >= '" & Format(dtStart, "YYYY-MM-DD") & "'" & _
                " and EXAMINEDATE <= '" & Format(dtEnd, "YYYY-MM-DD") & "'" & _
                 IIf(lngSampleNO = -1, "", " And SAMPLEID=" & lngSampleNO) & _
                " order by SAMPLEID,EXAMINEDATE,ITEMCODE"
        Set rsSQL = ConnSQL.Execute(strSQL)
        
        If rsSQL.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If
        
        rsSQL.MoveFirst
        lngResultCount = -1
        strBh = "***"
        Do While Not rsSQL.EOF
            If strBh <> rsSQL("SAMPLEID") Then '新标本
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsSQL("EXAMINEDATE"), "yyyy-mm-dd hh:mm:ss")  '检验时间
                
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsSQL("SAMPLEID"))   '样本序号   Val("" & rsSQL("SAMPLEID"))
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      "尿液"  '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                   ' Val(rsSQL("质控")) '质控
                    
                strBh = rsSQL("SAMPLEID")
                   
            End If
              '将中文通道码替换成英文
                strTDM = rsSQL("ITEMCODE")
                strTDM = Replace(strTDM, "红细胞计数", "rbc")
                strTDM = Replace(strTDM, "白细胞总数", "TWBC")
                strTDM = Replace(strTDM, "上皮细胞计数", "spxb")
                strTDM = Replace(strTDM, "病理性管型", "blxgx")
                strTDM = Replace(strTDM, "非晶形尿酸结晶", "nsjj")
                strTDM = Replace(strTDM, "透明管型", "tmgx")
                strTDM = Replace(strTDM, "粘液丝", "nys")
                
                strTDM = Replace(strTDM, "滴虫", "dc")
                strTDM = Replace(strTDM, "真菌", "zj")
                strTDM = Replace(strTDM, "细菌", "xj")
                strTDM = Replace(strTDM, "结晶", "jj")
                strTDM = Replace(strTDM, "Unknown", "kong")
                strTDM = Replace(strTDM, "草酸钙结晶", "csgjj")
                strTDM = Replace(strTDM, "精子", "jz")
                strTDM = Replace(strTDM, "颗粒管型", "klgx")
            If strTDM <> "" Or IsNull(strTDM) = False Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Trim(strTDM) & "|" & Trim(Replace(Replace(Replace(IIf(rsSQL("ITEMCHAR") = "", rsSQL("ITEMVALUE"), rsSQL("ITEMCHAR")), "+-", "弱阳性(+-)"), "1+", "阳性(1+)"), "2+", "阳性(2+)"))
            End If
            
            If Trim(rsSQL("ITEMCODE")) = "真菌" Then
              '读取本地选定图片
              strImgBmp = ""
              strTmp = ""
              YearMonth = Format(rsSQL("EXAMINEDATE"), "yyyymm")
              YearMonthDay = Format(rsSQL("EXAMINEDATE"), "yyyymmdd")
              YearMonthSample = YearMonthDay & Trim(rsSQL("SAMPLEID"))
              strImgBmp = strPicPath & YearMonth & "\" & YearMonthDay & "\" & YearMonthSample & "\"
              
              If intDebug = 1 Then
                strLog = "查找图片路径: " & strImgBmp
                Call writelislog(strLog)
              End If
              
              strPicFix = Trim(rsSQL("QUERYID")) & "*p.jpg"
              
              If intDebug = 1 Then
                strLog = "查找图片名称: " & strPicFix
                Call writelislog(strLog)
              End If
              
              'Call FindFilesAPI(strPicPath, strPicFix)
              'Call FindFile(strPicPath, strPicFix)
              FileNames = FindFile2(strImgBmp, strPicFix)
              aFiles = Split(FileNames, "|")
              For intloop = LBound(aFiles) To UBound(aFiles)
                strImg = aFiles(intloop)
                If (Trim(strImg) <> "") And (Dir(strImgBmp + strImg) <> "") Then
                   strTmp = strTmp & "^" & CStr(intloop + 1) & ";113;0;" & strImgBmp & strImg
'                  strTmp = strTmp & "^尿沉渣曲线" & CStr(intloop + 1) & ";113;0;" & strImgBmp & strImg
                  If intDebug = 1 Then
                    strLog = "找到选定图片: " & strImgBmp & strImg
                    Call writelislog(strLog)
                  End If
                End If
              Next
              If strTmp <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & Mid(strTmp, 2)
              If intDebug = 1 Then
                strLog = "当前数据: " & aResults(lngResultCount)
                Call writelislog(strLog)
              End If
            End If
            rsSQL.MoveNext
        Loop
    End With
    
    If rsSQL.EOF Then
        ConnSQL.Close
    End If
    ResultFromFile = aResults
    
    Exit Function
errH:
    If intDebug = 1 Then
        Call writelislog("错误号(" & Err.Number & "):(" & Err.Description & ")")
    End If
End Function

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_LX5000_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function




