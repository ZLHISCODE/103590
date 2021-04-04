Attribute VB_Name = "mdlMipClientShell"
Option Explicit

'######################################################################################################################


Public Enum mTextAlign
    taLeftAlign = 0
    taCenterAlign = 1
    taRightAlign = 2
End Enum



Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrDbUser As String                 '当前数据库用户
Public glngSys As Long                      '当前系统
Public gfrmMain As Object                   '导航台窗口，主要用于作消息编辑窗口的父窗口

Public glngParentForm As Long
Public gcnOracle As ADODB.Connection
Public gfrmMipResource As frmMipResource

'######################################################################################################################

Public Function CreateCondition() As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Fields.Append "条件名称", adVarChar, 30
        .Fields.Append "条件结果", adVarChar, 4000
        .Fields.Append "条件类型", adVarChar, 30
        .Open
    End With
    
    Set CreateCondition = rs
    
End Function

Public Function SetCondition(ByRef rs As ADODB.Recordset, ByVal strConditionName As String, ByVal strConditionValue As String, Optional ByVal strConditionType As String = "文本") As Boolean
    
    rs.Filter = ""
    rs.Filter = "条件名称='" & strConditionName & "'"
    If rs.RecordCount = 0 Then rs.AddNew
    rs("条件名称").Value = strConditionName
    rs("条件结果").Value = strConditionValue
    rs("条件类型").Value = strConditionType
    
    SetCondition = True
    
End Function

Public Function GetCondition(ByRef rs As ADODB.Recordset, ByVal strConditionName As String) As String
    
    rs.Filter = ""
    rs.Filter = "条件名称='" & strConditionName & "'"
    If rs.RecordCount > 0 Then
        GetCondition = CStr(rs("条件结果").Value)
    End If
    
End Function


Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    '******************************************************************************************************************
    '功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    '参数：
    '返回：
    '******************************************************************************************************************
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub DrawAngle(picDraw As PictureBox, ByVal fAngle As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim iSize       As Integer
    Dim iFillStyle  As Integer
    Dim lFillColor  As Long
    Dim lForeColor  As Long
    Dim lRet        As Long
    Dim uaPts(3)    As PointAPI

    'Size arrow to best fit picDraw at any angle
    iSize = IIf(picDraw.ScaleHeight < picDraw.ScaleWidth, Int(picDraw.ScaleHeight / PI), Int(picDraw.ScaleWidth / PI))
    
    'Setup the 4 points of the arrow using the first point
    'as the center and the other points offset from the center.
    uaPts(0).X = picDraw.ScaleWidth / 2
    uaPts(0).Y = picDraw.ScaleHeight / 2
    uaPts(1).X = uaPts(0).X - iSize
    uaPts(1).Y = uaPts(0).Y - iSize
    uaPts(2).X = uaPts(0).X + iSize
    uaPts(2).Y = uaPts(0).Y
    uaPts(3).X = uaPts(0).X - iSize
    uaPts(3).Y = uaPts(0).Y + iSize
    
    'Rotate the arrow to the correct angle
    Call RotatePoints(uaPts(0), uaPts, fAngle)
    
    'Save picDraw settings
    iFillStyle = picDraw.FillStyle
    lFillColor = picDraw.FillColor
    lForeColor = picDraw.ForeColor
    
    'Setup picDraw to fill the arrow
    picDraw.FillStyle = vbFSSolid   'Solid Fill
    picDraw.FillColor = &HFFFFFF    'Inside = White
    picDraw.ForeColor = &H0&        'Border = Black
    
    'Draw the filled arrow
    lRet = Polygon(picDraw.hDC, uaPts(0), 4)
    
    'Restore picDraw settings
    picDraw.FillStyle = iFillStyle
    picDraw.FillColor = lFillColor
    picDraw.ForeColor = lForeColor

    'Free the memory
    Erase uaPts
    
End Sub

Private Sub RotatePoints(uAxisPt As PointAPI, uRotatePts() As PointAPI, fDegrees As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    
    'Rotates an array of PointAPI points around a center point by fDegrees
    
    Dim lIdx        As Long
    Dim fDX         As Single
    Dim fDY         As Single
    Dim fRadians    As Single

    fRadians = fDegrees * RADS
    
    For lIdx = 0 To UBound(uRotatePts)
        fDX = uRotatePts(lIdx).X - uAxisPt.X
        fDY = uRotatePts(lIdx).Y - uAxisPt.Y
        uRotatePts(lIdx).X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
        uRotatePts(lIdx).Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    Next lIdx
    
End Sub

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, ByVal strTip As String)
    '******************************************************************************************************************
    '功能：在任务栏上增加一个图标
    '参数：
    '说明：
    '******************************************************************************************************************
        
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub ModifyIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, ByVal strTip As String)
    '******************************************************************************************************************
    '功能：在任务栏上增加一个图标
    '参数：
    '说明：
    '******************************************************************************************************************
        
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_MODIFY, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    '******************************************************************************************************************
    '功能：从任务栏上删除图标
    '参数：
    '说明：
    '******************************************************************************************************************
    
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '事件发生的载体
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
    
End Sub

Public Function AnalyseComputer() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim strComputer As String * 256
    
    On Error Resume Next
    
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
    
End Function

Public Sub DrawColorToColor(picDraw As Object, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal blnVertical As Boolean = True, Optional ByVal blnBorder As Boolean = False)
    '******************************************************************************************************************
    '功能：画出从一种颜色到另一种颜色的渐变
    '参数：
    '说明：
    '******************************************************************************************************************

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

Public Sub PicShowFlat(objPic As Object, Optional intStyle As Integer = -1, Optional strName As String = "", Optional intAlign As mTextAlign)
'功能：将PictureBox模拟成凹下或凸起的面板
'参数：intStyle:0=平面,-1=凹下,1=凸起
'      intAlign=如果要显示文本,则指定对齐方式
    
    Dim vRect As RECT, lngTmp As Long
    
    With objPic
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            vRect.Left = .ScaleLeft
            vRect.Top = .ScaleTop
            vRect.Right = .ScaleWidth
            vRect.Bottom = .ScaleHeight
            DrawEdge .hDC, vRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If intAlign = taCenterAlign Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2 '中间对齐
            ElseIf intAlign = taRightAlign Then
                .CurrentX = .ScaleWidth - .TextWidth(strName) - 2 '右边对齐
            Else
                .CurrentX = 2 '左边对齐
            End If
            objPic.Print strName
        End If
    End With
End Sub

Public Sub PlayWave(ByVal lngKey As Long)
    '******************************************************************************************************************
    '功能：播放资源文件中的指定资源声音文件(wave)
    '参数：ID=资源号
    '说明：
    '******************************************************************************************************************
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String
    Dim objFso As New FileSystemObject
    
    On Error GoTo errHand
    
    If lngKey = 0 Then Exit Sub
        
    If objFso.FolderExists(App.Path & "\Data") = False Then
        Call objFso.CreateFolder(App.Path & "\Data")
    End If
    
    strFile = App.Path & "\Data\zlMipClientShell_Wave_" & lngKey & ".wav"
    
    If objFso.FileExists(strFile) = False Then
        arrData = LoadResData(lngKey, "WAVE")
        intFile = FreeFile
        Open strFile For Binary As intFile
        Put intFile, , arrData()
        Close intFile
    End If
    
    Call sndPlaySound(strFile, SND_NODEFAULT Or SND_ASYNC)
    
    Exit Sub
    
errHand:
    
End Sub

Public Function LoadIcon(Path As String, cx As Long, cy As Long) As Long
    LoadIcon = LoadImage(App.hInstance, App.Path + "\" + Path, IMAGE_ICON, cx, cy, LR_LOADFROMFILE)
End Function

Public Function CboLocate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
    'blnItem:True-表示根据ItemData的值定位下拉框;False-表示根据文本的内容定位下拉框
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If blnItem Then
            If cboObj.ItemData(lngLocate) = Val(strValue) Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboObj.List(lngLocate), InStr(1, cboObj.List(lngLocate), "-") + 1) = strValue Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function


Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    
End Function


Public Function GetBasePeriod(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '功能:获取特殊时间
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim intDay As Integer
    Dim varValue As Variant
    
    If Left(strMode, 3) = "自定义" Then
        '自定义:3,4
        varValue = Split(Mid(strMode, 5), ",")
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", Val(varValue(0)), Now), "yyyy-MM-dd") & " 00:00:00"
        Else
            If UBound(varValue) < 1 Then
                GetBasePeriod = Format(Now, "yyyy-MM-dd") & " 23:59:59"
            Else
                GetBasePeriod = Format(DateAdd("d", Val(varValue(1)), Now), "yyyy-MM-dd") & " 23:59:59"
            End If
        End If
            
        Exit Function
    End If
    
    Select Case strMode
    Case "所  有"
        GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
    Case "当  时"      '当时
        GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetBasePeriod = Format(Now, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(Now, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 7 - intDay, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetBasePeriod = Format(Now, "YYYY-MM") & "-01 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(Now, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(Now, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-04-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-10-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(Now, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetBasePeriod = Format(Now, "YYYY") & "-01-01 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -3, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -7, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -15, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -30, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -60, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -90, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -180, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 2, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function

Public Function InitXMLDoc() As Object

    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo errHand
        
    varXMLVersion = Split("6.0,4.0", ",")
    
    On Error Resume Next
    For intLoop = 0 To UBound(varXMLVersion)
        Err = 0
        Set objXML = CreateObject("MSXML2.DOMDocument." & varXMLVersion(intLoop))
        If Err = 0 Then
            strXMLVer = varXMLVersion(intLoop)
            Exit For
        End If
    Next
    On Error GoTo errHand
    
    If strXMLVer = "" Then
        MsgBox "创建MSXML2.DOMDocument对象失败"
        Exit Function
    End If
    
    Set InitXMLDoc = objXML
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description
End Function

