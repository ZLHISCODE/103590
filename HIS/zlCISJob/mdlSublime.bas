Attribute VB_Name = "mdlSublime"
Option Explicit

Public Const GW_OWNER = 4
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
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
Public Const WM_CLOSE = &H10
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
Public Const SRCCOPY = &HCC0020
Public Const TOGGLE_HIDEWINDOW = &H80
Public Const TOGGLE_UNHIDEWINDOW = &H40
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'绘制矩形的一条或者多条边
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWHEELSCROLLLINES = 104
Public WHEEL_SCROLL_LINES As Long
Public gobjScroll As Object             '保存 Hook 对象
Global glngPrevWndProc As Long

Public Const WM_MOUSEWHEEL = &H20A

Public Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Public Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Public Const madDbDateDefault As Integer = 20               '日期型字段缺省长度


Public Sub Hook(ByVal objParent As Object)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error Resume Next
    Debug.Print 1 / 0
    If err <> 0 Then Exit Sub
    
    Set gobjScroll = objParent
    glngPrevWndProc = SetWindowLong(objParent.hwnd, GWL_WNDPROC, AddressOf WindowProc)

    '获取"控制面板"中的滚动行数值
    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WHEEL_SCROLL_LINES, 0)

    If WHEEL_SCROLL_LINES > objParent.HScr.Max Then WHEEL_SCROLL_LINES = objParent.HScr.Max
End Sub

Public Sub UnHook(ByVal objParent As Object)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngReturnValue As Long
    If glngPrevWndProc = 0 Then Exit Sub
    lngReturnValue = SetWindowLong(objParent.hwnd, GWL_WNDPROC, glngPrevWndProc)
    Set gobjScroll = Nothing
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '******************************************************************************************************************
    '功能：捕获系统事件并进行处理
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim pt As POINTAPI
    Dim wzDelta
    Dim wKeys As Integer
    
    Select Case uMsg
    Case WM_MOUSEWHEEL                          '滚轮事件
        wzDelta = OS.HIWORD(wParam)
        wKeys = OS.LOWORD(wParam)
        pt.X = OS.LOWORD(lParam)
        pt.Y = OS.HIWORD(lParam)
    
        '将屏幕坐标转换为frmCaseTendBody窗口坐标
        ScreenToClient gobjScroll.hwnd, pt
        
        With gobjScroll
            If .mintREPORTSEL <> -1 Then
                WindowProc = CallWindowProc(glngPrevWndProc, hw, uMsg, wParam, lParam)
            Else
                If .HScr.Visible Then
                    '判断坐标是否在frmCaseTendBody.BodyEdit窗口内
                    If pt.X > .Left / Screen.TwipsPerPixelX And pt.X < (.Left + .Width) / Screen.TwipsPerPixelX And pt.Y > .Top / Screen.TwipsPerPixelY And pt.Y < (.Top + .Height) / Screen.TwipsPerPixelY Then
            
                        If wKeys = 16 Then
                            '水平滚动
                            
                        Else
                            '垂直滚动
                            If Sgn(wzDelta) = 1 Then
                                .HScr.Value = IIf(.HScr.Value - WHEEL_SCROLL_LINES < .HScr.Min, .HScr.Min, .HScr.Value = .HScr.Value - WHEEL_SCROLL_LINES)
                            Else
                                .HScr.Value = IIf(.HScr.Value + WHEEL_SCROLL_LINES > .HScr.Max, .HScr.Max, .HScr.Value + WHEEL_SCROLL_LINES)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Case Else                                   '其他事件仍由系统缺省处理
        WindowProc = CallWindowProc(glngPrevWndProc, hw, uMsg, wParam, lParam)
    End Select
End Function



Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, _
        Optional ByVal blnDelete As Boolean = False, Optional ByVal strValueSplit As String = "|")
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名,值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, strValueSplit)
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

