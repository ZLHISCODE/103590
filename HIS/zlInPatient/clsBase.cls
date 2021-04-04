VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const GWL_WNDPROC = -4

Public Enum COLOR_NativeXpPlain
    BackgroundDark = 14054755
    BackgroundLight = 15180411
    HighlightBorderBottomRight = 8388608
    HighlightBorderTopLeft = 8388608
    HighlightHot = 12775167
    HighlightPressed = 4096254
    HighlightSelected = 7323903
    NormalGroupCaptionDark = 14215660
    NormalGroupCaptionLight = 14215660
    NormalGroupCaptionTextHot = 0
    NormalGroupCaptionTextNormal = 0
    NormalGroupClient = 16244694
    NormalGroupClientBorder = 16777215
    NormalGroupClientLink = 12999969
    NormalGroupClientLinkHot = 16748098
    NormalGroupClientText = 0
    SpecialGroupCaptionDark = 14215660
    SpecialGroupCaptionLight = 14215660
    SpecialGroupCaptionTextHot = 0
    SpecialGroupCaptionTextSpecial = 0
    SpecialGroupClient = 16244694
    SpecialGroupClientBorder = 16777215
    SpecialGroupClientLink = 12999969
    SpecialGroupClientLinkHot = 16748098
    SpecialGroupClientText = 0
End Enum

Public Enum Color
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    
    原始单据 = 0
    冲销记录 = &HFF
    停用项目 = &H8000000C
    启用项目 = 0
    
    公共模块色 = &HC00000
    
    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
    
End Enum

Public Enum REGISTER
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum

Public Type POINTAPI
     X As Long
     Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowRect& Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private mstrSQL As String
Private mrsTmp As ADODB.Recordset
Private mstrTitle As String
Private mlngTXTProc As Long

'######################################################################################################################

Public Function CommandBarInit(ByRef cbsMain As Object, Optional ByVal blnEnableCustomization As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
    cbsMain.VisualTheme = xtpThemeOffice2003
        
    With cbsMain.Options
        .ShowExpandButtonAlways = blnEnableCustomization
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization blnEnableCustomization

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    
End Function

Public Function DockPannelInit(ByRef dkpMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function TabControlInit(ByRef tbc As Object, Optional ByVal bytPosition As XTPTabPosition = xtpTabPositionTop) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    With tbc
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
            .Position = bytPosition
        End With
        
'        Set .Icons = frmPubResource.imgApp.Icons
        
    End With

    TabControlInit = True
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    Select Case bytMode
    Case 1
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 99
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function ZVal(ByVal varValue As Variant) As String
    '******************************************************************************************************************
    '功能：将0零转换为"NULL"串,在生成SQL语句时用
    '******************************************************************************************************************
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function ReadPicture(rsTable As ADODB.Recordset, strField As String, Optional strFile As String) As String
    '******************************************************************************************************************
    '功能：将指定的记录集图形字段复制为图形临时文件
    '参数：
    '       rsTable   图形存储记录集
    '       strField  图形字段
    '       strFile   用户定义的文件名（可选项）
    '返回：
    '******************************************************************************************************************
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, FileNum, j
    Dim aryChunk() As Byte
    Dim strTempFile As String
    
    On Error GoTo errH
    lngFileSize = rsTable.Fields(strField).ActualSize
    If lngFileSize = 0 Then
        '未读取有效数据
        Exit Function
    End If
    
    FileNum = FreeFile
    If strFile = "" Then
        '当用户并没定义文件名时
        
        strFile = CreateTmpFile

    End If
    Open strFile For Binary As FileNum
    
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    rsTable.Move 0
    For j = 0 To intBolcks
        If j = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim aryChunk(lngCurSize - 1) As Byte
        aryChunk() = rsTable.Fields(strField).GetChunk(lngCurSize)
        Put FileNum, , aryChunk()
    Next
    Close FileNum
    ReadPicture = strFile
    Exit Function

errH:
    Close FileNum
'    Kill strFile
    ReadPicture = ""

End Function

Public Function CreateTmpFile(Optional ByVal strFileType As String = "tmp") As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim strFileTemp As String
       
    
    strFileTemp = GetTmpPath
    
    strFileTemp = strFileTemp & "zlNewPic" & Format(Now, "yyyymmdd") & Format(Timer, "0") & "." & strFileType
    
    CreateTmpFile = strFileTemp
End Function

Public Function GetTmpPath() As String
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
End Function

Public Function ParamCreate(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "参数名", adVarChar, 50
        .Fields.Append "参数值", adVarChar, 50
        
        .Open
        .ActiveConnection = Nothing
    End With
    
    ParamCreate = True
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ParamAdd(ByRef rs As ADODB.Recordset, ByVal strParamName As String, Optional ByVal strParamValue As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    
    rs("参数名").Value = strParamName
    rs("参数值").Value = strParamValue
    
    ParamAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function ParamRead(ByRef rs As ADODB.Recordset, ByVal strParamName As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.Filter = ""
    rs.Filter = "参数名='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        ParamRead = rs("参数值").Value
    End If
    
    Exit Function
    
errHand:
End Function

Public Function ParamWrite(ByRef rs As ADODB.Recordset, ByVal strParamName As String, ByVal strParamValue As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.Filter = ""
    rs.Filter = "参数名='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        rs("参数值").Value = strParamValue
    End If
    
    Exit Function
    
errHand:
End Function

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    '******************************************************************************************************************
    '
    '
    '
    '******************************************************************************************************************
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function Between(X, a, b) As Boolean
    '******************************************************************************************************************
    '功能：判断x是否在a和b之间
    '******************************************************************************************************************
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function


Public Function LoadTree(ByRef objTvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objNode As Node
    Dim strTmp As String
    Dim strIcon As String
    Dim strIconSel As String
    Dim blnIcon As Boolean
    Dim blnIconSel As Boolean
    
    On Error GoTo errHand

    On Error Resume Next
    
    blnIcon = (rs("图标").Name = "图标")
    blnIconSel = (rs("选中图标").Name = "选中图标")
    
    On Error GoTo errHand
    
    LockWindowUpdate objTvw.hWnd

    Do While Not rs.EOF
        strTmp = AppendCode(zlCommFun.Nvl(rs("名称").Value), zlCommFun.Nvl(rs("编码").Value))
        
        If blnIcon Then strIcon = zlCommFun.Nvl(rs("图标").Value)
        If blnIconSel Then strIconSel = zlCommFun.Nvl(rs("选中图标").Value)
        If strIconSel = "" Then strIconSel = strIcon
        
        If IsNull(rs("上级id").Value) Then
            Set objNode = objTvw.Nodes.Add(, , "K" & zlCommFun.Nvl(rs("ID").Value, 0), strTmp, strIcon, strIconSel)
        Else
            Set objNode = objTvw.Nodes.Add("K" & rs("上级id").Value, tvwChild, "K" & zlCommFun.Nvl(rs("ID").Value, 0), strTmp, strIcon, strIconSel)
        End If

        rs.MoveNext
    Loop

    LockWindowUpdate 0

    LoadTree = True

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CopyRecordStruct(ByVal rsFrom As ADODB.Recordset, Optional ByVal blnRowID As Boolean = False, Optional ByVal blnNotOpen As Boolean = False) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim lngLoop As Long
    Dim rs As ADODB.Recordset
    
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic

    For lngLoop = 0 To rsFrom.Fields.Count - 1

        Select Case rsFrom.Fields(lngLoop).Type
        Case 135            'Oracle的Date型
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, adVarChar, 30, rsFrom.Fields(lngLoop).Attributes
        Case 5
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, rsFrom.Fields(lngLoop).Type, 30, rsFrom.Fields(lngLoop).Attributes
        Case Else
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, adVarChar, rsFrom.Fields(lngLoop).DefinedSize + 12
        End Select

    Next
    If blnRowID Then
        rs.Fields.Append "行号", adVarChar, 30
    End If
    
    If blnNotOpen = False Then rs.Open: rs.ActiveConnection = Nothing
    
    Set CopyRecordStruct = rs
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CopyRecordData(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset, Optional blnAll As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    If blnAll Then
        If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    End If
    
    Do While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.Count - 1
            
            On Error Resume Next
            strTmp = ""
            strTmp = rsTo.Fields(rsFrom.Fields(lngLoop).Name).Name
            On Error GoTo errHand
            
            If UCase(strTmp) = UCase(rsFrom.Fields(lngLoop).Name) Then
                rsTo.Fields(strTmp).Value = Trim(zlCommFun.Nvl(rsFrom.Fields(lngLoop).Value))
            End If

        Next
        If blnAll = False Then Exit Do
        rsFrom.MoveNext
        rsTo.Update
    Loop
    
    CopyRecordData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function DeleteRecordData(rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If Not (rs Is Nothing) Then
        If rs.RecordCount > 0 Then rs.MoveFirst
        While Not rs.EOF
            rs.Delete
            rs.MoveNext
        Wend
    End If
    
    DeleteRecordData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub CopyRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset)
    '******************************************************************************************************************
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    Set rsTo = New ADODB.Recordset
    For lngLoop = 0 To rsFrom.Fields.Count - 1
        rsTo.Fields.Append rsFrom.Fields(lngLoop).Name, rsFrom.Fields(lngLoop).Type, rsFrom.Fields(lngLoop).DefinedSize
    Next
    rsTo.Open
    rsTo.ActiveConnection = Nothing
    
    If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.Count - 1
            rsTo.Fields(lngLoop).Value = rsFrom.Fields(lngLoop).Value
        Next
        rsFrom.MoveNext
    Wend
    
errHand:
    
End Sub

Public Sub DeleteRecord(rs As ADODB.Recordset)
    '******************************************************************************************************************
    '功能:删除记录集
    '参数:rs        要删除的记录集
    '返回:无
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If rs.RecordCount > 0 Then rs.MoveFirst
    While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Wend
    
errHand:
End Sub

Public Function AppendRecord(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer
    
    Do While Not rsFrom.EOF
        rsTo.AddNew
        For intCol = 0 To rsFrom.Fields.Count - 1
            rsTo(rsFrom.Fields(intCol).Name).Value = zlCommFun.Nvl(rsFrom.Fields(intCol).Value)
        Next
        
        rsFrom.MoveNext
    Loop
    
    AppendRecord = True
    
End Function

Public Function AddPeriodToComboBox(ByRef cbo As Object, Optional ByVal blnParameter As Boolean)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    With cbo
        .Clear
        .AddItem "今  天"
        .AddItem "昨  天"
        .AddItem "本  周"
        .AddItem "本  月"
        .AddItem "本  季"
        .AddItem "本半年"
        .AddItem "本  年"
        .AddItem "前三天"
        .AddItem "前一周"
        .AddItem "前半月"
        .AddItem "前一月"
        .AddItem "前二月"
        .AddItem "前三月"
        .AddItem "前半年"
        .AddItem "前一年"
        .AddItem "前二年"
        If blnParameter = False Then .AddItem "自定义"
    End With
    
    AddPeriodToComboBox = True
    
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
            GetBasePeriod = Format(DateAdd("d", Val(varValue(0)), zlDatabase.Currentdate), "yyyy-MM-dd") & " 00:00:00"
        Else
            If UBound(varValue) < 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59"
            Else
                GetBasePeriod = Format(DateAdd("d", Val(varValue(1)), zlDatabase.Currentdate), "yyyy-MM-dd") & " 23:59:59"
            End If
        End If
            
        Exit Function
    End If
    
    Select Case strMode
    Case "当  时"      '当时
        GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function


Public Function IncStr(ByVal strVal As String) As String
    '******************************************************************************************************************
    '功能：对一个字符串自动加1。
    '说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    '******************************************************************************************************************
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

Public Function IntEx(vNumber As Variant) As Variant
    '******************************************************************************************************************
    '功能：取大于指定数值的最小整数
    '******************************************************************************************************************
    IntEx = -1 * Int(-1 * vNumber)
End Function

Public Function InDesign() As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function Lpad(ByVal strText As String, ByVal lngLen As Long, ByVal strReplace As String) As String
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngL As Long
    
    lngL = Len(strText)
    If lngL > lngLen Then
        Lpad = Left(strText, lngLen)
    ElseIf lngL < lngLen Then
        Lpad = String(lngLen - lngL, strReplace) & strText
    Else
        Lpad = strText
    End If
End Function

Public Sub DrawPicture(pic As Object, objPic As StdPicture, ByVal W As Long, ByVal H As Long)
    '******************************************************************************************************************
    '功能：在PictureBox中央按适当比例画一幅图
    '参数：W,H=要作图的尺寸
    '******************************************************************************************************************
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    If W <= pic.ScaleWidth And H <= pic.ScaleHeight Then
        lngW = W: lngH = H
    Else
        sngW = W / pic.ScaleWidth
        sngH = H / pic.ScaleHeight
        If sngW > sngH Then
            lngW = W / sngW: lngH = H / sngW
        Else
            lngW = W / sngH: lngH = H / sngH
        End If
    End If
    
    pic.Cls
    On Error Resume Next
    pic.PaintPicture objPic, (pic.ScaleWidth - lngW) / 2, (pic.ScaleHeight - lngH) / 2, lngW, lngH
    
End Sub

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1表示开始;2表示结束
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
        .ActiveConnection = Nothing
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
    '******************************************************************************************************************
    '功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
    '参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    '******************************************************************************************************************
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
        End If
    End If
    FormatEx = strNumber
End Function

Public Function AppendCode(ByVal strName As String, ByVal strCode As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If strName <> "" And strCode <> "" Then
        AppendCode = "【" & strCode & "】" & strName
    Else
        AppendCode = strName
    End If
End Function

Public Function CheckAllNumber(ByVal strKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long

    For lngLoop = 1 To Len(strKey)
        If Mid(strKey, lngLoop, 1) < "0" Or Mid(strKey, lngLoop, 1) > "9" Then
            Exit Function
        End If
    Next

    CheckAllNumber = True
End Function

Public Sub LocationObj(ByRef objTxt As Object, Optional ByVal blnDoevents As Boolean = False)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If blnDoevents Then DoEvents
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
    
errHand:
    
End Sub

Public Sub LocationGrid(ByRef vsf As Object, Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error Resume Next
    
    If lngRow <> -1 Then vsf.Row = lngRow
    If lngCol <> -1 Then vsf.Col = lngCol
    
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True, _
                        Optional ByVal strShowField As String, _
                        Optional ByVal strKeyField As String, _
                        Optional ByVal strDefaultField As String)
    '******************************************************************************************************************
    '功能:装载数据入指定的组合下拉框或网格中的下拉框中
    '参数:
    '返回:
    '******************************************************************************************************************
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        
        If strShowField = "" Then strShowField = rsTemp1.Fields(0).Name
        If strKeyField = "" Then strKeyField = rsTemp1.Fields(1).Name
        
        If rsTemp1.Fields.Count > 2 Then
            If strDefaultField = "" Then strDefaultField = rsTemp1.Fields(2).Name
        End If
        
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1(strShowField).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1(strKeyField).Value)
            
            If rsTemp1.Fields.Count > 2 Then
                If Val(rsTemp1(strDefaultField).Value) = 1 Then
                    objSource.ListIndex = objSource.NewIndex
                End If
            End If
            
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub


Public Sub DrawLine(pic As Object, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1)
    '在(X1,Y1),(X2,Y2)之间使用ForeColor色画一直线
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    
    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)
    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth
End Sub


Public Function FillLvw(ByRef objLvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objItem As ListItem
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    LockWindowUpdate objLvw.hWnd
    
    Do While Not rs.EOF
        
        Set objItem = objLvw.ListItems.Add(, "K" & rs("ID").Value, rs("名称").Value, rs("图标").Value, rs("图标").Value)
        For lngLoop = 2 To objLvw.ColumnHeaders.Count
            objItem.SubItems(lngLoop - 1) = zlCommFun.Nvl(rs(objLvw.ColumnHeaders(lngLoop).Text).Value)
        Next
                        
        rs.MoveNext
    Loop
    
    LockWindowUpdate 0
    
    FillLvw = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function HaveExcel() As Boolean
    '******************************************************************************************************************
    '功能：判断本机上装有EXCEL没有
    '参数：
    '返回：有则返回True
    '******************************************************************************************************************

    On Error GoTo errHandle
    
    Dim objTemp  As Object
    
    Set objTemp = CreateObject("Excel.Application") '打开一个EXCEL程序
    
    Set objTemp = Nothing
    
    HaveExcel = True
    
    Exit Function

errHandle:
    Set objTemp = Nothing
    HaveExcel = False
End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '功能： 将指定的信息保存在注册表中
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strKeyValue-键值
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand

    Select Case enmRegister
    Case 注册信息

        Call SaveSetting("ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue)

    Case 私有模块

        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)

    Case 私有全局

        Call SaveSetting("ZLSOFT", "私有全局\" & UserInfo.用户名 & "\" & strSection, strKey, strKeyValue)

    Case 公共模块

        Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strKeyValue)

    Case 公共全局

        Call SaveSetting("ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue)

    End Select

    SetRegister = True

errHand:

End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '功能： 将指定的注册信息读取出来
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strDefKeyValue-缺省键值
    '返回： strKeyValue-键值
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        strValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, strDefKeyValue)
        
    Case 私有模块

        strValue = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 私有全局

        strValue = GetSetting("ZLSOFT", "私有全局\" & UserInfo.用户名 & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共模块

        strValue = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共全局
        
        strValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        cbrPopupItem.Parameter = cbrControl.Parameter
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrPopupItem2 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
                cbrPopupItem2.Parameter = cbrControl2.Parameter
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ShowPubDate(ByVal obj As Object, ByRef strInitKey As String) As Boolean
    '******************************************************************************************************************
    '功能：打开树型+列表结构,应用于表格控件
    '参数：
    '返回：0:取消选择;1:选择;2:无数据返回
    '******************************************************************************************************************

    Dim lngX As Long
    Dim lngY As Long
    Dim lngObjHeight As Long
    Dim objPoint As POINTAPI
    Dim strDate As String

    On Error GoTo errHand

    Call ClientToScreen(obj.hWnd, objPoint)

    Select Case TypeName(obj)
    Case "TextBox", "CommandButton"

        lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
        lngY = obj.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        lngObjHeight = obj.Height

    Case Else
        lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
        lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
        lngObjHeight = obj.CellHeight
    End Select

    strDate = strInitKey
    If frmTemplateDate.ShowDialog(lngX, lngY, strDate, lngObjHeight) Then

        strInitKey = strDate
        ShowPubDate = True

    End If

    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'Public Function ShowPubSelect(ByVal frmParent As Object, _
'                                ByVal obj As Object, _
'                                ByVal bytStyle As Byte, _
'                                ByVal strLvw As String, _
'                                ByVal strSavePath As String, _
'                                ByVal strDescrible As String, _
'                                ByVal rsData As ADODB.Recordset, _
'                                ByRef rsResult As ADODB.Recordset, _
'                                Optional ByVal lngCX As Long = 9000, _
'                                Optional ByVal lngCY As Long = 4500, _
'                                Optional ByVal blnMuliSel As Boolean = False, _
'                                Optional ByVal strInitKey As String = "", _
'                                Optional ByVal strFilterControl As String = "", _
'                                Optional ByVal blnReturnOneData As Boolean, _
'                                Optional ByVal blnLeftSelect As Boolean = False, _
'                                Optional ByVal strKind As String, _
'                                Optional ByVal strDefaultKind As String, _
'                                Optional ByVal blnAll As Boolean = True) As Byte
'    '******************************************************************************************************************
'    '功能：打开树型+列表结构,应用于表格控件
'    '参数：
'    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
'    '返回：0:取消选择;1:选择;2:无数据返回
'    '******************************************************************************************************************
'
'    Dim lngX As Long
'    Dim lngY As Long
'    Dim lngObjHeight As Long
'    Dim rs As New ADODB.Recordset
'    Dim objPoint As POINTAPI
'
'    On Error GoTo errHand
'
'    strSavePath = "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & strSavePath
'
'    If rsData.BOF Then
'        ShowPubSelect = 2
'        Exit Function
'    End If
'
'    If blnReturnOneData Then
'        If rsData.RecordCount = 1 Then
'            Set rsResult = rsData
'            ShowPubSelect = 1
'            Exit Function
'        End If
'    End If
'
'    If obj Is Nothing Then
'        lngX = (Screen.Width - lngCX) / 2
'        lngY = (Screen.Height - lngCY) / 2
'        lngObjHeight = 0
'    Else
'        Call ClientToScreen(obj.hWnd, objPoint)
'
'        Select Case TypeName(obj)
'        Case "TextBox", "CommandButton"
'
'            lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
'            lngY = obj.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
'            lngObjHeight = obj.Height
'
'        Case Else
'            lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
'            lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
'            lngObjHeight = obj.CellHeight
'        End Select
'    End If
'
'    ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, blnLeftSelect, False, blnMuliSel, strFilterControl, strKind, strDefaultKind, blnAll)
'
'    If ShowPubSelect = 1 Then
'        Set rsResult = rsData
'
'        If rsResult.BOF And blnMuliSel = False Then
'            ShowPubSelect = 0
'        End If
'
'    End If
'
'    Exit Function
'
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '******************************************************************************************************************
    '功能:检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    '参数:
    '返回:
    '******************************************************************************************************************
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, ParamInfo.系统名称
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字符！", vbExclamation, ParamInfo.系统名称
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function

Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '
    '******************************************************************************************************************
    Dim intCol As Integer

    With msf

        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub

Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal X As Single, ByVal Y As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long
            
    lngX = X / 15
    lngY = Y / 15
        
    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next
    
    lngXY = lngXY * lngY + lngX
    
    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub

Public Function DockPannelCreate(ByRef dkpMain As Object, ByVal intIndex As Integer, _
                                    ByVal lngCX As Long, ByVal lngCY As Long, _
                                    ByVal bytDirection As DockingDirection, _
                                    Optional ByVal objNeighbour As Pane = Nothing, _
                                    Optional ByVal strTitle As String, _
                                    Optional ByVal bytOptions As PaneOptions) As Pane
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Set DockPannelCreate = dkpMain.CreatePane(intIndex, lngCX, lngCY, bytDirection, objNeighbour)
    DockPannelCreate.Title = strTitle
    DockPannelCreate.Options = PaneNoCaption
    
End Function

Public Function GetTrayHeight() As Long
    '******************************************************************************************************************
    '功能：获取任务栏的高度
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
    '******************************************************************************************************************
    '功能：得控件中指定坐标在屏幕中的位置(Twip)
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
    '******************************************************************************************************************
    '功能：将VB的系统颜色转换为RGB色
    '参数：
    '返回：
    '******************************************************************************************************************

    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Public Function SetNewWindowLong(ByVal lngHwnd As Long, ByVal dwNewLong As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mlngTXTProc = GetWindowLong(lngHwnd, GWL_WNDPROC)
    Call SetWindowLong(lngHwnd, GWL_WNDPROC, dwNewLong)
        
    SetNewWindowLong = True
    
End Function

Public Function RestoreWindowLong(ByVal lngHwnd As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call SetWindowLong(lngHwnd, GWL_WNDPROC, mlngTXTProc)
    
    RestoreWindowLong = True
End Function

Public Function CallNewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call CallWindowProc(mlngTXTProc, hWnd, msg, wParam, lParam)
    
    CallNewWindowProc = True
End Function
'
'Public Function GetDefaultCode(ByVal strTable As String, ByVal lngUpKey As Long, ByVal str上级编码 As String, intMaxLength As Integer, Optional ByVal byt编码方式 As Byte = 1, Optional ByVal str前缀 As String) As String
'    '******************************************************************************************************************
'    '功能：产生缺省编码
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim rs As New ADODB.Recordset
'    Dim str最大编码 As String
'    Dim str编码 As String
'
'    Select Case byt编码方式
'    Case 0
'
'        Set rs = gclsPackage.GetMaxCode(strTable, 0, 2)
'        If rs.BOF = False Then
'            str最大编码 = Trim(zlCommFun.Nvl(rs("最大编码").Value))
'        End If
'        If str最大编码 = "" Then
'            str编码 = Format(Val(str最大编码) + 1, String(4, "0"))
'        Else
'            str编码 = Format(Val(str最大编码) + 1, String(Len(str最大编码), "0"))
'        End If
'
'    Case 1
'        Set rs = gclsPackage.GetMaxCode(strTable, lngUpKey, 3)
'        If rs.BOF = False Then
'            str最大编码 = Trim(zlCommFun.Nvl(rs("最大编码").Value))
'        End If
'        If str最大编码 = "" Then
'            str编码 = str上级编码 & "01"
'        ElseIf Len(str最大编码) > Len(str前缀 & str上级编码) Then
'            str最大编码 = Mid(str最大编码, Len(str前缀 & str上级编码) + 1)
'            str编码 = str前缀 & str上级编码 & Format(Val(str最大编码) + 1, String(Len(str最大编码), "0"))
'        Else
'            str编码 = str前缀 & str上级编码 & "01"
'        End If
'    End Select
'
'    If Len(str编码) <= intMaxLength Then
'        GetDefaultCode = str编码
'    End If
'
'End Function

Public Function SetColorIcon(frmMain As Object, ByVal Key As String, ByVal Color As OLE_COLOR, ByRef ils As Object) As Boolean
    '******************************************************************************************************************
    '功能：产生颜色块，并添加到ImageList对象中
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim ctlPictureBox As VB.PictureBox
    
    On Error GoTo errHand
    
    Set ctlPictureBox = frmMain.Controls.Add("VB.PictureBox", "ctlPictureBox1")
    
    Dim ListImage As ListImage
    Set ListImage = ils.ListImages("User")
    
    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = ils.MaskColor
    
    ctlPictureBox.Picture = ListImage.ExtractIcon
    
    If Color = vbWhite Then Color = RGB(254, 254, 254)
    
    ctlPictureBox.Cls
    ctlPictureBox.Line (30, 30)-(ctlPictureBox.width - 105, ctlPictureBox.Height - 105), Color, BF
    ctlPictureBox.Refresh

    On Error Resume Next
    ils.ListImages.Remove ils.ListImages(Key).Index
    On Error GoTo errHand
    
    ils.ListImages.Add , Key, ctlPictureBox.Image
    ils.ListImages(Key).Tag = CStr(Color)

    frmMain.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
    
    SetColorIcon = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AppendColorImageList(frmMain As Object, ils As Object, ByVal lngColor As Long) As Boolean
    '******************************************************************************************************************
    '功能：产生颜色块，并添加到ImageList对象中
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objItem As Object
    
    On Error Resume Next
    Set objItem = Nothing
    Set objItem = ils.ListImages("K" & lngColor)
    On Error GoTo errHand
    If objItem Is Nothing Then
        Call SetColorIcon(frmMain, "K" & lngColor, lngColor, ils)
    End If
    
    Exit Function
    
errHand:
        
End Function

'Public Function OpenColorDialog(ByVal lngColor As Long) As Long
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    With frmPubResource.dlg
'
'        .Color = lngColor
'        .ShowColor
'
'        If .Color <> lngColor Then
'            lngColor = .Color
'        End If
'    End With
'
'    OpenColorDialog = lngColor
'
'End Function

Public Function GetVBKey() As Long
    
    Dim strTmp As String
    
    strTmp = Timer
    strTmp = Replace(strTmp, ".", "")
    
    GetVBKey = Format(Now, "1dd") & strTmp
    
    GetVBKey = 0 - GetVBKey
    
End Function

Public Function DrawPlatControl(ByVal objForm As Object, picDraw As Object) As Boolean
    '******************************************************************************************************************
    '功能：将picDraw内的控件变为平面控件
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strName As String
    Dim objTxt As Object
    Dim objPic As Object
    Dim objTmp As Object
    Dim lngIndex As Long
    
    Dim objDTPickerL() As Object
    Dim objDTPickerT() As Object
    Dim objDTPickerB() As Object
    Dim objComboL() As Object
    Dim objComboT() As Object
    Dim objComboB() As Object
    
    On Error GoTo errHand
    
    ReDim objDTPickerL(1)
    ReDim objDTPickerT(1)
    ReDim objDTPickerB(1)
    ReDim objComboL(1)
    ReDim objComboT(1)
    ReDim objComboB(1)
    
    picDraw.Cls
    
    For Each objTmp In objForm.Controls
        
        Select Case TypeName(objTmp)
        Case "TextBox", "ComboBox", "DTPicker"
            If TypeName(objTmp.Container) = TypeName(picDraw) Then
                
                If objTmp.Visible Then
                    Select Case TypeName(objTmp)
                    '--------------------------------------------------------------------------------------------------
                    Case "TextBox"
                        
                        If objTmp.BorderStyle = 0 Then
                            objTmp.BackColor = picDraw.BackColor
                            picDraw.Line (objTmp.Left, objTmp.Top + objTmp.Height)-(objTmp.Left + objTmp.width, objTmp.Top + objTmp.Height)
                        End If
                        
                    '--------------------------------------------------------------------------------------------------
                    Case "ComboBox"
                        
                        On Error Resume Next
                                                
                        lngIndex = objTmp.Index

                        If UBound(objComboL) < lngIndex Then
                            ReDim Preserve objComboL(lngIndex)
                            ReDim Preserve objComboT(lngIndex)
                            ReDim Preserve objComboB(lngIndex)
                        End If
                        
                        If objComboL(lngIndex) Is Nothing Then
                            
                            strName = "picCombo_" & objTmp.Index & "_L"
                            Set objComboL(lngIndex) = CreatePictureBox(objForm, picDraw, strName)
                            
                            strName = "picCombo_" & objTmp.Index & "_T"
                            Set objComboT(lngIndex) = CreatePictureBox(objForm, picDraw, strName)
                            
                            strName = "picCombo_" & objTmp.Index & "_B"
                            Set objComboB(lngIndex) = CreatePictureBox(objForm, picDraw, strName)

                            objComboB(lngIndex).AutoRedraw = True
                            objComboB(lngIndex).ForeColor = picDraw.ForeColor
                            objComboB(lngIndex).Enabled = False
                            
                        End If
                        
                        objComboL(lngIndex).Move objTmp.Left, objTmp.Top, 30, objTmp.Height
                        objComboT(lngIndex).Move objTmp.Left, objTmp.Top, objTmp.width, 30
                        objComboB(lngIndex).Move objTmp.Left, objTmp.Top + objTmp.Height - 30, objTmp.width, 30
                        
                        objComboB(lngIndex).Cls
                        objComboB(lngIndex).Line (0, 0)-(objComboB(lngIndex).width - 30, 0)
                        
                        objComboL(lngIndex).Visible = True
                        objComboT(lngIndex).Visible = True
                        objComboB(lngIndex).Visible = True
                    '--------------------------------------------------------------------------------------------------
                    Case "DTPicker"
                        
                        On Error Resume Next
                        
                        lngIndex = objTmp.Index

                        If UBound(objDTPickerL) < lngIndex Then
                            ReDim Preserve objDTPickerL(lngIndex)
                            ReDim Preserve objDTPickerT(lngIndex)
                            ReDim Preserve objDTPickerB(lngIndex)
                        End If
                                                
                        If objDTPickerL(lngIndex) Is Nothing Then
                        
                            strName = "picDTPicker_" & objTmp.Index & "_L"
                            Set objDTPickerL(lngIndex) = CreatePictureBox(objForm, picDraw, strName)
                            
                            strName = "picDTPicker_" & objTmp.Index & "_T"
                            Set objDTPickerT(lngIndex) = CreatePictureBox(objForm, picDraw, strName)
                            
                            strName = "picDTPicker_" & objTmp.Index & "_B"
                            Set objDTPickerB(lngIndex) = CreatePictureBox(objForm, picDraw, strName)

                            objDTPickerB(lngIndex).AutoRedraw = True
                            objDTPickerB(lngIndex).ForeColor = picDraw.ForeColor
                            objDTPickerB(lngIndex).Enabled = False
                        End If
                        
                        objDTPickerL(lngIndex).Move objTmp.Left, objTmp.Top, 30, objTmp.Height
                        objDTPickerT(lngIndex).Move objTmp.Left, objTmp.Top, objTmp.width, 30
                        objDTPickerB(lngIndex).Move objTmp.Left, objTmp.Top + objTmp.Height - 30, objTmp.width, 30
                        
                        objDTPickerB(lngIndex).Cls
                        objDTPickerB(lngIndex).Line (0, 0)-(objDTPickerB(lngIndex).width - 30, 0)
                        
                        objDTPickerL(lngIndex).Visible = True
                        objDTPickerT(lngIndex).Visible = True
                        objDTPickerB(lngIndex).Visible = True
                    End Select
                End If
                
            End If
        End Select
    Next
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function CreatePictureBox(ByVal objForm As Object, picDraw As Object, ByVal strName As String) As Object
    
    Dim objPic As Object
    
    Set objPic = objForm.Controls.Add("VB.PictureBox", strName)
    Set objPic.Container = picDraw
    objPic.BackColor = picDraw.BackColor
    objPic.ZOrder
    objPic.AutoRedraw = True
    objPic.ForeColor = picDraw.ForeColor
    objPic.Enabled = False
                                
    Set CreatePictureBox = objPic
End Function


Public Function zlBlobRead(ByVal KeyWord As String, Optional ByRef strFile As String) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select zl24_Lob_Read('" & KeyWord & "'," & "[1]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

