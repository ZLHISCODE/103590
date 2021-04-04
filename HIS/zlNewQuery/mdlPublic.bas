Attribute VB_Name = "mdlPublic"

Public Type CLIENTRECT
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'输入法控制API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const KLF_REORDER = &H8
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000

Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24
'------
Public Type FILETIME
    dwLowDate  As Long
    dwHighDate As Long
End Type
 
Public Type SYSTEMTIME
    wYear      As Integer
    wMonth     As Integer
    wDayOfWeek As Integer
    wDay       As Integer
    wHour      As Integer
    wMinute    As Integer
    wSecond    As Integer
    wMillisecs As Integer
End Type
 
Public Type POINTAPI
     X As Long
     Y As Long
End Type

Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
 
Public Const READ_WRITE = 2
  
Public Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, ByVal MullP As Long, ByVal NullP2 As Long, lpLastWriteTime As FILETIME) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function mciSendString Lib "Winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

'下列语句用于检测是否合法调用
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public gobjFSO As New Scripting.FileSystemObject    'FSO对象

Public Function MusicPlayStatus(Optional ByVal strAlia As String = "MediaMusic") As Boolean
    Dim MCIStatusLen As Integer
    Dim MCIStatus As String
    Dim ret As Integer
        
    On Error Resume Next
    MCIStatusLen = 15
    MCIStatus = String(MCIStatusLen + 1, " ")
    ret = mciSendString("STATUS " & strAlia & " MODE", MCIStatus, MCIStatusLen, 0)
    Select Case Trim(UCase(Left$(MCIStatus, 7)))
    Case "PLAYING"
        MusicPlayStatus = True
    Case "STOPPED"
        MusicPlayStatus = False
    Case Else
        MusicPlayStatus = False
    End Select
    On Error GoTo 0
    
End Function
Public Sub MusicPlay(ByVal strSong As String, Optional ByVal strAlia As String = "MediaMusic")
    Dim ret As Integer
    Dim mciReturnLength  As Integer
    
    '打开音乐Media文件
    On Error Resume Next
    ret = mciSendString("open " & strSong & " type sequencer alias " & strAlia, 0&, mciReturnLength, 0)
    '播放
    ret = mciSendString("play " & strAlia & " notify", 0&, 0, 0)
    On Error GoTo 0
End Sub

Public Sub MusicClose(Optional ByVal strAlia As String = "MediaMusic")
    Dim ret As Integer
   '关闭
   On Error Resume Next
   ret = mciSendString("close " & strAlia, 0&, 0, 0)
   On Error GoTo 0
End Sub
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strName As String = "", Optional ByVal Off As Single = 0)
'功能：将PictureBox模拟成3D平面按钮
'参数：intStyle:0=平面,-1=凹下,1=凸起
    
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = Off + (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Sub PrintText(picBox As PictureBox, ByVal strText As String, ByVal bytAligment As Byte, Optional ByVal Off As Single = 0)
    If strText <> "" Then
        Select Case bytAligment
        Case 0
            picBox.CurrentX = Off
        Case 1
            picBox.CurrentX = (picBox.ScaleWidth - picBox.TextWidth(strText)) / 2
        Case 2
            picBox.CurrentX = (picBox.ScaleWidth - picBox.TextWidth(strText))
        End Select
        picBox.CurrentY = (picBox.ScaleHeight - picBox.TextHeight(strText)) / 2
        
        picBox.Print strText
    End If
End Sub

Public Function SetFileDateTime(ByVal strFileName As String, ByVal TheDate As String) As Boolean
    Dim lRet As Long
    Dim lngFileHand As Long
    Dim typFileTime As FILETIME
    Dim typLocalTime As FILETIME
    Dim typSystemTime As SYSTEMTIME
    
    If Dir(strFileName) = "" Then Exit Function
    If Not IsDate(TheDate) Then Exit Function
    
    With typSystemTime
        .wYear = Year(TheDate)
        .wMonth = Month(TheDate)
        .wDay = Day(TheDate)
        .wDayOfWeek = Weekday(TheDate) - 1
        .wHour = Hour(TheDate)
        .wMinute = Minute(TheDate)
        .wSecond = Second(TheDate)
    End With

    lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
    lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)
    
    lngFileHand = lopen(strFileName, READ_WRITE)
    
    lRet = SetFileTime(lngFileHand, ByVal 0&, ByVal 0&, typFileTime)
    CloseHandle lngFileHand
    
    SetFileDateTime = lRet > 0
    
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function FindCboIndex(cbo As ComboBox, lngID As Long) As Long
'功能：由项目值查找ComboBox的项目索引
    Dim i As Integer
    If lngID < 0 Then FindCboIndex = -1: Exit Function
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = lngID Then
            FindCboIndex = i
            Exit Function
        End If
    Next
    FindCboIndex = -1
End Function

Public Sub ShowFlatFlash(Optional strNote As String, Optional frmParent As Object)
    '------------------------------------------------
    '功能： 显示等待的动态窗体
    '参数：
    '   strNote:提示信息
    '   frmParent：关于窗体的父窗体
    '返回：
    '------------------------------------------------
    With frmFlash
        If strNote <> "" Then .lbl提示.Caption = strNote
        Err = 0
        On Error Resume Next
        .avi.Open gstrAviPath & "\" & "Findfile.avi"
        If Err <> 0 Then
            .lblFile.Visible = True
        End If
        .Refresh
        If frmParent Is Nothing Then
            .Show
        Else
            .Show , frmParent
        End If
        If Not .lblFile.Visible Then .avi.Play
    End With
End Sub

Public Sub StopFlatFlash()
    '------------------------------------------------
    '功能： 停止并关闭等待的动态窗体
    '参数：
    '返回：
    '------------------------------------------------
    On Error Resume Next
    frmFlash.avi.Stop
    Unload frmFlash
End Sub

Public Sub DrawColorToColor(picDraw As Object, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal blnVertical As Boolean = True, Optional ByVal blnBorder As Boolean = False)
'画出从一种颜色到另一种颜色的渐变
    Dim VR, VG, VB As Single
    Dim R, G, b, R2, G2, B2 As Integer
    Dim temp As Long, Y As Long, X As Long
    Dim tmpMode As Long
    Dim blnAutoRedraw As Boolean
    
    '只有窗体和图片可以画
    If Not (TypeOf picDraw Is PictureBox Or TypeOf picDraw Is Form) Then Exit Sub
    
    
    tmpMode = picDraw.ScaleMode
    blnAutoRedraw = picDraw.AutoRedraw
    
    picDraw.ScaleMode = 3
    picDraw.AutoRedraw = True
    
    temp = (Color1 And 255)
    R = temp And 255
    temp = Int(Color1 / 256)
    G = temp And 255
    temp = Int(Color1 / 65536)
    b = temp And 255
    temp = (Color2 And 255)
    R2 = temp And 255
    temp = Int(Color2 / 256)
    G2 = temp And 255
    temp = Int(Color2 / 65536)
    B2 = temp And 255

    If blnVertical Then
        VR = Abs(R - R2) / picDraw.ScaleHeight
        VG = Abs(G - G2) / picDraw.ScaleHeight
        VB = Abs(b - B2) / picDraw.ScaleHeight
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB
        For Y = 0 To picDraw.ScaleHeight
            R2 = R + VR * Y
            G2 = G + VG * Y
            B2 = b + VB * Y
            picDraw.Line (0, Y)-(picDraw.ScaleWidth, Y), RGB(R2, G2, B2)
        Next Y
    Else
        VR = Abs(R - R2) / picDraw.ScaleWidth
        VG = Abs(G - G2) / picDraw.ScaleWidth
        VB = Abs(b - B2) / picDraw.ScaleWidth
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < b Then VB = -VB
        For X = 0 To picDraw.ScaleWidth
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = b + VB * X
            picDraw.Line (X, 0)-(X, picDraw.ScaleHeight), RGB(R2, G2, B2)
        Next X
    End If
    
    If blnBorder Then
        picDraw.DrawWidth = 2
        picDraw.Line (1, 1)-(picDraw.ScaleWidth - 1, picDraw.ScaleHeight - 1), &HC000&, B
        picDraw.DrawWidth = 1
    End If
    
    picDraw.Refresh
    picDraw.ScaleMode = tmpMode
    picDraw.AutoRedraw = blnAutoRedraw
End Sub

Public Function Custom_WndMessage(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hwnd, msg, wp, lp)
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function LoadImageData(ByVal strPath As String, ByVal lngID As Long) As Boolean
        '从数据库读取图形数据，绘制后保存到指定的路径下。
        '入参：
        '   strPath 路径
        '   lngID   检验图像结果的ID
        '--如果有的话, 删除原来的临时图形文件
        Static objImg As Object
        Dim rsTmp As New ADODB.Recordset, rsImage As New ADODB.Recordset
        Dim rsItem As New ADODB.Recordset
        Dim strImageType As String
        Dim strImageData As String
        Dim DrawIndex As Integer
        Dim intLoop As Integer
        Dim lngStart As Long
        Dim strTmp As String
        Dim strSQL  As String
    
        Dim blnPic As Boolean '是否图片格式
        Dim lngFileNum As Long, lngCount As Long, lngBound As Long
        Dim aryChunk() As Byte, strFile As String
        Dim intLayOut As Integer
        Dim killFile As String
    
        Dim blnFtp As Boolean       'FTP是否可用
        Static strFtpPara As String       '保存FTP参数
        Dim strFtpUser As String, strFtpPass As String, strFtpIP As String, strFtpDir As String
        Dim strDownOk As String, strFtpPath   As String, strLocalFile As String
        Dim objStream As TextStream
    
        On Error GoTo ErrHandle
    
100     If Dir(strPath & "\" & lngID & ".cht") <> "" Then
102         LoadImageData = True
            Exit Function
        End If
    
        'FTP连接检查，有效则可以按FTP方式取图片
104     blnFtp = False
106     If strFtpPara = "" Then
108         strFtpPara = zlDatabase.GetPara("FTP设置", 100, 1208, "")
        End If
110     If UBound(Split(strFtpPara, ";")) >= 3 Then
112        strFtpUser = Split(strFtpPara, ";")(0)
114        strFtpPass = Split(strFtpPara, ";")(1)
116        strFtpIP = Split(strFtpPara, ";")(2)
118        strFtpDir = Split(strFtpPara, ";")(3)
120        If TestFTP(strFtpUser, strFtpPass, strFtpIP, strFtpDir) = "" Then
122             blnFtp = True
           End If
        End If
    
'124     mlngImageID = lngID
    
126     lngCount = 0
128     strFile = ""
   
130     strSQL = "select 标本id,图像类型,图像位置 from 检验图像结果 where id = [1] "
132     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lngID)
    
134     If rsTmp.EOF = True Then
            Exit Function
        End If
    
136     If objImg Is Nothing Then Set objImg = CreateObject("zlLisDev.clsDrawGraph")
    
138     Do Until rsTmp.EOF
140         strImageType = Trim("" & rsTmp("图像类型"))
142         strFtpPath = Trim("" & rsTmp!图像位置)
144         If InStr(strFtpPath, ";") <= 0 Or Not blnFtp Then
                '- 图像存在数据库中，按原来的方式处理
146             gstrSQL = "select Zl_FUN_Get检验图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "',0) from dual "
148             Set rsImage = zlDatabase.OpenSQLRecord(gstrSQL, "LoadImgData")
150             strTmp = Nvl(rsImage(0))
152             strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
            
154             If strImageData <> "" Then
156                 intLoop = 0
                
158                 If Val(Mid(strImageData, 1, 3)) >= 100 And Val(Mid(strImageData, 1, 3)) <= 227 And Mid(strImageData, 4, 1) = ";" Then
                
160                     blnPic = True
162                     If Mid(strImageData, 1, 3) >= 100 And Mid(strImageData, 1, 3) <= 107 Then
164                         strFile = App.Path & "\zlLisPic" & lngID & ".bmp"
166                     ElseIf Mid(strImageData, 1, 3) >= 110 And Mid(strImageData, 1, 3) <= 117 Then
168                         strFile = App.Path & "\zlLisPic" & lngID & ".jpg"
170                     ElseIf Mid(strImageData, 1, 3) >= 120 And Mid(strImageData, 1, 3) <= 127 Then
172                         strFile = App.Path & "\zlLisPic" & lngID & ".gif"
174                     ElseIf Mid(strImageData, 1, 3) >= 200 And Mid(strImageData, 1, 3) <= 227 Then
176                         If gobjFSO.FolderExists(App.Path & "\ZLLIS_ZIP") = False Then
178                             gobjFSO.CreateFolder App.Path & "\ZLLIS_ZIP"
                            End If
180                         If gobjFSO.FolderExists(App.Path & "\ZLLIS_ZIP\" & lngID) = False Then
182                             gobjFSO.CreateFolder App.Path & "\ZLLIS_ZIP\" & lngID
                            End If
184                         strFile = App.Path & "\ZLLIS_ZIP\" & lngID & "\ZLISPIC.ZIP"
                        End If
                    
                    
186                     intLayOut = Val(Mid(strImageData, 1, 3))
188                     strImageData = Mid(strImageData, 5)
190                     lngFileNum = FreeFile
192                     lngCount = 0
    
194                     If Dir(strFile) <> "" Then Kill strFile
196                     Open strFile For Binary As lngFileNum
198                     ReDim aryChunk(Len(strImageData) / 2 - 1) As Byte
200                     For lngBound = LBound(aryChunk) To UBound(aryChunk)
202                         aryChunk(lngBound) = CByte("&H" & Mid(strImageData, lngBound * 2 + 1, 2))
                        Next
                    
204                     Put lngFileNum, , aryChunk()
                    
                    End If
                    '-------保存为图片文件
206                 Do While strTmp <> ""
208                     intLoop = intLoop + 1
210                     gstrSQL = "select Zl_FUN_Get检验图像(" & rsTmp("标本id") & ",'" & Nvl(rsTmp("图像类型")) & "'," & intLoop & ") from dual "
212                     Set rsImage = zlDatabase.OpenSQLRecord(gstrSQL, "LoadImgData")
                    
214                     strTmp = Nvl(rsImage(0))
    
216                     If blnPic Then
                            '
218                         If strTmp <> "" Then
220                             ReDim aryChunk(Len(strTmp) / 2 - 1) As Byte
222                             For lngBound = LBound(aryChunk) To UBound(aryChunk)
224                                 aryChunk(lngBound) = CByte("&H" & Mid(strTmp, lngBound * 2 + 1, 2))
                                Next
                            
226                             Put lngFileNum, , aryChunk()
                            End If
                        Else
                            '图形数据
228                         strImageData = strImageData & Replace(Replace(Trim(strTmp), vbCr, ""), vbLf, "")
                        End If
                    Loop
                
230                 If blnPic Then
232                     strImageData = intLayOut & ";" & strFile
234                     Close lngFileNum
                    End If
                End If
            Else
                '图像存在FTP中，从FTP中取数据
                '图像位置的数据格式为：图像格式;FTP文件路径
            
236             intLayOut = Val(Split(strFtpPath, ";")(0))
238             strFtpPath = Trim(Split(strFtpPath, ";")(1))
240             strImageData = ""
242             If intLayOut >= 100 And intLayOut <= 227 Then
                    ' 图片文件，直接下载到本地
244                 strLocalFile = strPath & "\" & Split(strFtpPath, "/")(UBound(Split(strFtpPath, "/")))
246                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
248                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
250                 If strDownOk = "" Then
252                     strImageData = intLayOut & ";" & strLocalFile
                    End If
                Else
                    ' 图形数据，需要从下载的文本文件中读取数据
254                 strLocalFile = strPath & "\" & lngID & "_" & strImageType & ".txt"
256                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
258                 strDownOk = DownFile(strFtpUser, strFtpPass, strFtpIP, strFtpPath, strLocalFile)
260                 If strDownOk = "" Then
262                     Set objStream = gobjFSO.OpenTextFile(strLocalFile, ForReading)
264                     Do Until objStream.AtEndOfLine
266                         strImageData = strImageData & objStream.ReadLine
                        Loop
268                     objStream.Close
270                     Set objStream = Nothing
272                     strImageData = Replace(Replace(Trim(strImageData), vbCr, ""), vbLf, "")
274                     strImageData = intLayOut & ";" & strImageData
                    End If
276                 If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
                End If
            End If
        
278         If Len(strImageData) <> 0 Then
280             If Not objImg Is Nothing Then
282                 LoadImageData = objImg.DrawImg(strImageType, strImageData, strPath & "\" & lngID & ".cht")
                End If
            End If
        
284         strTmp = "": strImageData = ""
286         rsTmp.MoveNext
        Loop
        Exit Function
ErrHandle:
'        WriteLog "LoadImagedata" & CStr(Erl()) & "行，" & Err.Description
288     If ErrCenter() = 1 Then
290         Resume
        End If
End Function


'-----以下是 FTP 相关函数
Private Function TestFTP(ByVal strUser As String, ByVal strPassWord As String, _
                            ByVal strDevAdress As String, ByVal strFtpPath As String) As String
                            
    Dim FtpNet As New clsFtp, strPath As String, strTmpPath As String           'FTP类
    Dim lngFileNo As Long
    strPath = Format(Now, "yyyymmddHHMMSS")
    strTmpPath = IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & "temp.txt"
    lngFileNo = FreeFile
    Open strTmpPath For Output As lngFileNo
    Close lngFileNo
    If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir(strFtpPath, "FTP测试" & strPath) > 0 Then
            TestFTP = "在FTP上不能创建目录！"
        Else
            If FtpNet.FuncUploadFile(strFtpPath, strTmpPath, "temp.txt") > 0 Then
                TestFTP = "上传文件失败"
            Else
                FtpNet.FuncFtpDisConnect '先断开，再删除，不然删不掉
                If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) <= 0 Then
                     TestFTP = "FTP不能连接！"
                ElseIf FtpNet.FuncFtpDelDir(strFtpPath, "FTP测试" & strPath) > 0 Then
                    TestFTP = "在FTP上不能删除目录"
                Else
                    TestFTP = ""
                End If
            End If
        End If
    Else
        TestFTP = "不能连接FTP！"
    End If
    FtpNet.FuncFtpDisConnect
    Set FtpNet = Nothing
    Kill strTmpPath
End Function

Private Function DownFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strFile As String) As String
        '从FTP服务器下载文件。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpFile :FTP上的文件。
        'strFile    :本地文件全路径。
        '返回：空串表示成功，否则为错误提示。
        Dim objFtp As New clsFtp, lngReturn As Long, strFtpFileName As String, strLocaFile As String
        Dim strFtpDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFile = "请指定要下载的文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFtpDir = Replace(strFtpFile, "/" & strFtpFileName, "")
108     strLocaFile = strFile
110     If strLocaFile = "" Then
112         DownFile = "请指定下载的文件保存到何处！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFile = "要下载的文件已存在！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFile = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFile = "不能连接服务器！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFtpDir)
130     If lngReturn <> 0 Then
132         DownFile = "不能进入指定的目录，可能是权限不足或服务器上无此目录！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownloadFile(strFtpDir, strLocaFile, strFtpFileName)
136     If lngReturn <> 0 Then
138         DownFile = "下载失败，可能是权限不足或服务器上无此文件！"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFile = CStr(Erl()) & "行，" & Err.Description
End Function

Private Function UploadFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                            ByVal strFtpPath As String, ByVal strFile As String, Optional strNewFileName As String) As String
        '按本地文件名上传文件到FTP服务器。
        'strUser    :用户名
        'strPass    :密码
        'strServer  :服务器
        'strFtpPath :FTP上的目录，无目录会自动创建。
        'strFile    :本地文件全路径。
        'strNewFileName: 传到FTP上后的文件名，为空则按本地文件名保存
        '返回：空串表示成功，否则为错误提示。
    
        Dim objFtp As New clsFtp, lngReturn As Long, strFileName As String, strLocaFile As String
        On Error GoTo errH
    
    
100     If Left(strFtpPath, 1) = "/" Then strFtpPath = Mid$(strFtpPath, 2)
    
102     If strServer = "" Then
104         UploadFile = "请指定FTP服务器"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
106     strLocaFile = strFile
108     If Dir(strLocaFile) = "" Then
110         UploadFile = "文件" & strLocaFile & "不存在!"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        If strNewFileName = "" Then
112         strFileName = Split(strLocaFile, "\")(UBound(Split(strLocaFile, "\")))
        Else
            strFileName = strNewFileName
        End If
114     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
116     If lngReturn <> 0 Then
            '检查目录是否存在
118         lngReturn = objFtp.FuncChangeDir(strFtpPath)
120         If lngReturn <> 0 Then
122             lngReturn = objFtp.FuncFtpMkDir("/", strFtpPath)
124             If lngReturn <> 0 Then
126                 UploadFile = "创建目录失败！可能是权限不足！"
                    objFtp.FuncFtpDisConnect
                    Set objFtp = Nothing
                    Exit Function
                End If
            End If
        
128         lngReturn = objFtp.FuncUploadFile("/" & strFtpPath, strLocaFile, strFileName)
130         If lngReturn <> 0 Then
132             UploadFile = "上传文件失败，可能是权限不足！"
                objFtp.FuncFtpDisConnect
                Set objFtp = Nothing
                Exit Function

            Else
134             UploadFile = ""
            End If
        Else
136         UploadFile = "不能连接服务器！"
        End If
        objFtp.FuncFtpDisConnect
        Set objFtp = Nothing
        Exit Function
errH:
138     UploadFile = CStr(Erl()) & "行，" & Err.Description
End Function



Public Function ValEx(ByVal varInput As Variant) As Variant
'功能：由于Val只能以数字开头识别，ValEx以第一个数字进行识别
    Dim arrTmp As Variant, lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
End Function

