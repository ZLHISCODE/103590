VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVsf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'################################################################################################################
'## 模 块 名：  clsVsf.cls
'## 创 建 人：  陈福容
'## 日    期：  2006年12月14日
'## 修 改 人：
'## 日    期：
'## 描    述：  VSFlexGrid辅助类
'## 需    要：
'##
'## 使用方法：
'################################################################################################################

Private Const GWL_WNDPROC = -4
Private Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Public Enum VsfEditControl
    '1-文本编辑;2-选择/过滤编辑;3-下拉数据编辑;4-日期时间编辑;5-复选编辑
    vbVsfEditText = 1
    vbVsfEditCommand = 2
    vbVsfEditCombox = 3
    vbVsfEditCheck = 4
End Enum

'光标经过列条件
Public Enum ECursorJump
    m未指定 = -1   '添加列时(AppendColumn)缺省值为-1，将使用表格初始化(Initialize)中的光标类型值
    m强制经过 = 0
    m无值时经过 = 1
    m有值时经过 = 2
    m不经过 = 3
End Enum
Private mbytCurSorJumpType As ECursorJump '缺省类型(初始化调用)

Private Type ColPropertys
    编辑状态 As Boolean
    允许编辑 As Boolean         '
    编辑控件 As VsfEditControl
    最大长度 As Integer         '
    下拉数据 As String          '
    调整列宽 As Boolean
    允许排序 As Boolean
    允许拖动 As Boolean
    自动行高 As Boolean
    系统隐藏 As Boolean
    自动换行 As Boolean
    光标跳跃 As ECursorJump
End Type

Public Enum mTextType
    整数型 = 1
    小数型 = 2
End Enum

Private Type POINTAPI
     X As Long
     Y As Long
End Type

Private mobjVsf As Object
Private mobjControls As Object
Private milsIcon As Object
Private mblnAppendRows As Boolean
Private mblnAllowEdit As Boolean
Private mblnAllowNewRow As Boolean
Private mblnAllowDeleteRow As Boolean
Private mblnAllowColMove As Boolean
Private mblnEditing As Boolean
Private mblnSelectTrace As Boolean
Private mblnLoading As Boolean
Private mbytDragState As Byte
Private mstrSvrKey As String
Private mintKeyCol As Integer
Private mlngStartPos As Long
Private mintAutoResizeColumn As Integer
Private mintIndicatorColumn As Integer
Private mColProperty() As ColPropertys
Private mintIndicator As Integer
Private micoIndicator As StdPicture
Private mblnAutoRowHeight As Boolean
Private mblnMoveColumn As Boolean
Public Event BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
Public Event AfterNewRow(ByVal Row As Long, Col As Long)
Public Event RsetFocuse(ByVal strEvent As String) '事件处理完后，供主界面重新设置焦点
Public Event BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
Public Event BeforeDeleteCell(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
Public Event BeforeComboList(ByVal OldCol As Long, ByVal NewCol As Long, ComboList As String, Cancel As Boolean)

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'######################################################################################################################
Public Property Let MoveColumn(ByVal blnData As Boolean)
    mblnMoveColumn = blnData
End Property

Public Property Get MoveColumn() As Boolean
    MoveColumn = mblnMoveColumn
End Property

Public Property Get VsfBody() As Object
    Set VsfBody = mobjVsf
End Property

Public Property Let DropColData(ByVal intCol As Integer, ByVal strDropData As String)
    
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    mColProperty(intCol).下拉数据 = strDropData
End Property

Public Property Get DropColData(ByVal intCol As Integer) As String

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    DropColData = mColProperty(intCol).下拉数据
End Property

Public Property Let EditControl(ByVal intCol As Integer, ByVal bytData As VsfEditControl)
    
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    mColProperty(intCol).编辑控件 = bytData
End Property

Public Property Get EditControl(ByVal intCol As Integer) As VsfEditControl

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    EditControl = mColProperty(intCol).编辑控件
End Property

Public Property Let EditSelStart(ByVal lngData As Long)
    mlngStartPos = lngData
End Property

Public Property Get EditSelStart() As Long
    EditSelStart = mlngStartPos
End Property


Public Property Let EditSelText(ByVal strData As String)
    Dim strTmpConv As String
    Dim strSvrTmp As String
    Dim strTmp As String
    
    With mobjVsf
        strSvrTmp = .TextMatrix(.Row, .Col)
        
        strTmpConv = StrConv(MidB(StrConv(strSvrTmp, vbFromUnicode), 1, mlngStartPos), vbUnicode)
        strTmp = strTmpConv & strData
        strTmpConv = StrConv(MidB(StrConv(strSvrTmp, vbFromUnicode), mlngStartPos + 1), vbUnicode)
        
        mlngStartPos = mlngStartPos + LenB(strData)
        
        strTmp = strTmp & strTmpConv
        
        .TextMatrix(.Row, .Col) = strTmp
        
    End With
End Property


Public Property Let AllowColDrag(ByVal intCol As Integer, ByVal blnAllowDrag As Boolean)
    
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    mColProperty(intCol).允许拖动 = blnAllowDrag
End Property

Public Property Get AllowColDrag(ByVal intCol As Integer) As Boolean

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    AllowColDrag = mColProperty(intCol).允许拖动
End Property

Public Property Let AllowColResize(ByVal intCol As Integer, ByVal blnAllowColResize As Boolean)
    
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    mColProperty(intCol).调整列宽 = blnAllowColResize
End Property

Public Property Get AllowColResize(ByVal intCol As Integer) As Boolean

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    AllowColResize = mColProperty(intCol).调整列宽
End Property

Public Property Let AllowColSort(ByVal intCol As Integer, ByVal blnAllowSort As Boolean)
    
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    mColProperty(intCol).允许排序 = blnAllowSort
End Property

Public Property Get AllowColSort(ByVal intCol As Integer) As Boolean

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    AllowColSort = mColProperty(intCol).允许排序
End Property

Public Property Let AllowColEdit(ByVal intCol As Integer, ByVal blnAllowEdit As Boolean)

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property

    mColProperty(intCol).允许编辑 = blnAllowEdit
End Property

Public Property Get AllowColEdit(ByVal intCol As Integer) As Boolean

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    AllowColEdit = mColProperty(intCol).允许编辑
End Property

Public Property Let SysHidden(ByVal intCol As Integer, ByVal blnSysHidden As Boolean)

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property

    mColProperty(intCol).系统隐藏 = blnSysHidden
End Property

Public Property Get SysHidden(ByVal intCol As Integer) As Boolean

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    SysHidden = mColProperty(intCol).系统隐藏
End Property

Public Property Let AutoLinefeed(ByVal intCol As Integer, ByVal blnAutoLineFeed As Boolean)
    Dim intLoop As Integer
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    If AllowColEdit(intCol) = True And blnAutoLineFeed = True Then '允许编辑的列才能设置自动换行
        For intLoop = 0 To mobjVsf.Cols - 1
            mColProperty(intLoop).自动换行 = False
        Next intLoop
        mColProperty(intCol).自动换行 = blnAutoLineFeed
    End If
End Property

Public Property Get AutoLinefeed(ByVal intCol As Integer) As Boolean

    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    
    AutoLinefeed = mColProperty(intCol).自动换行
End Property

Public Property Let CursorJump(ByVal intCol As Integer, ByVal intCursorJump As ECursorJump)
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    mColProperty(intCol).光标跳跃 = intCursorJump
End Property

Public Property Get CursorJump(ByVal intCol As Integer) As ECursorJump
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Property
    CursorJump = mColProperty(intCol).光标跳跃
End Property

Public Property Get ColAutoLineFeed() As Boolean
    '是否设置了换行
    Dim intCol As Integer
    If mobjVsf Is Nothing Then Exit Property
    For intCol = 0 To mobjVsf.Cols - 1
        If AllowColEdit(intCol) = True And AutoLinefeed(intCol) = True Then
            ColAutoLineFeed = True
            Exit For
        End If
    Next intCol
End Property

Public Property Get GetAutoLineFeedCol() As Integer
    '获取设置了换行的列
    Dim intCol As Integer
    GetAutoLineFeedCol = -1
    If mobjVsf Is Nothing Then Exit Property
    For intCol = 0 To mobjVsf.Cols - 1
        If AllowColEdit(intCol) = True And AutoLinefeed(intCol) = True Then
            GetAutoLineFeedCol = intCol
            Exit For
        End If
    Next intCol
End Property

Public Property Get ColIndex(ByVal strColKey As String) As Long
    
    ColIndex = mobjVsf.ColIndex(strColKey)
    
End Property

Public Property Let AutoRowHeight(ByVal blnData As Boolean)
    mblnAutoRowHeight = blnData
    If blnData Then mobjVsf.AutoSizeMode = flexAutoSizeRowHeight
    mobjVsf.WordWrap = blnData
End Property

Public Property Get AutoRowHeight() As Boolean
    AutoRowHeight = mblnAutoRowHeight
End Property

Public Property Let IndicatorCol(ByVal intCol As Integer)
    mintIndicator = intCol
    Call ShowIndicatorIcon
End Property

Public Property Get IndicatorCol() As Integer
    IndicatorCol = mintIndicator
End Property

Public Property Set IndicatorIcon(ByVal vData As StdPicture)
    Set micoIndicator = vData
    Call ShowIndicatorIcon
End Property

Public Property Let AutoResizeColumn(ByVal intCol As Integer)
    mintAutoResizeColumn = intCol
End Property

Public Property Get AutoResizeColumn() As Integer
    AutoResizeColumn = mintAutoResizeColumn
End Property

Public Property Let AllowEdit(ByVal blnAllowEdit As Boolean)
    mblnAllowEdit = blnAllowEdit
End Property

Public Property Get AllowEdit() As Boolean
    AllowEdit = mblnAllowEdit
End Property

Public Property Let DragCol(ByVal lngData As Long)
'    mlngDragCol = lngData
End Property

Public Property Get DragCol() As Long
'    DragCol = mlngDragCol
End Property

Public Property Let KeyCol(ByVal intData As Integer)
    mintKeyCol = intData
End Property

Public Property Get KeyCol() As Integer
    KeyCol = mintKeyCol
End Property

Public Property Let SaveKey(ByVal strSvrKey As String)
    mstrSvrKey = strSvrKey
End Property

Public Property Get SaveKey() As String
    SaveKey = mstrSvrKey
End Property

Public Property Let DragState(ByVal bytData As Byte)
    mbytDragState = bytData
End Property

Public Property Get DragState() As Byte
    DragState = mbytDragState
End Property

Public Property Let ExtendLastCol(ByVal blnData As Boolean)
    mobjVsf.ExtendLastCol = blnData
End Property

Public Property Get ExtendLastCol() As Boolean
    ExtendLastCol = mobjVsf.ExtendLastCol
End Property

Public Property Let AppendRows(ByVal blnData As Boolean)
    mblnAppendRows = blnData
    If mblnAppendRows Then Call DrawSpaceCellLine
End Property

Public Property Get AppendRows() As Boolean
    AppendRows = mblnAppendRows
End Property

Public Property Let ExplorerBar(ByVal bytData As ExplorerBarSettings)
     mobjVsf.ExplorerBar = bytData
End Property

Public Property Get ExplorerBar() As ExplorerBarSettings
     ExplorerBar = mobjVsf.ExplorerBar
End Property

Public Property Let TextMatrix(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strData As String)
    mobjVsf.TextMatrix(lngRow, lngCol) = strData
End Property

Public Property Get TextMatrix(ByVal lngRow As Long, ByVal lngCol As Long) As String
    TextMatrix = mobjVsf.TextMatrix(lngRow, lngCol)
End Property

Public Function Initialize(objControls As Object, objvsf As Object, _
                            Optional ByVal blnExtendLastCol As Boolean = True, _
                            Optional ByVal blnSelectTrace As Boolean = True, _
                            Optional ByVal ilsIcon As Object, _
                            Optional ByVal bytCursorJumpType As ECursorJump = m强制经过) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objTmp As Object
    Dim strName As String
    
    Set mobjVsf = objvsf
    Set mobjControls = objControls
    
    On Error GoTo errHand
    
    mblnSelectTrace = blnSelectTrace
    mobjVsf.ExtendLastCol = blnExtendLastCol
    mobjVsf.ExplorerBar = flexExSortShowAndMove
    Set milsIcon = ilsIcon
    mbytCurSorJumpType = IIf(bytCursorJumpType = m未指定, m强制经过, bytCursorJumpType)
    
    If mblnSelectTrace And mobjVsf.Row >= 0 Then Call SelectRow(-1, mobjVsf.Row, 16772055)
    
    Initialize = True
    
    Exit Function
    
errHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InitializeEdit(Optional ByVal blnAllowEdit As Boolean, _
                                Optional ByVal blnAllowNewRow As Boolean, _
                                Optional ByVal blnAllowDeleteRow As Boolean, _
                                Optional ByVal intExplorerBar As Integer = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowEdit = blnAllowEdit
    
    If blnAllowEdit Then
        mobjVsf.Editable = flexEDKbdMouse
    Else
        mobjVsf.Editable = flexEDNone
    End If
    
    mobjVsf.ExplorerBar = intExplorerBar

    mblnAllowNewRow = blnAllowNewRow
    mblnAllowDeleteRow = blnAllowDeleteRow

    '允许编辑时，强行禁止VSF控件的高亮度属性，同时把编辑框边色设为蓝色
    mobjVsf.HighLight = flexHighlightNever
    mobjVsf.BackColorSel = &H8000000D

End Function

Public Function InitializeIndicator(ByVal vData As StdPicture, Optional ByVal intCol As Integer = 0, Optional ByVal strIcon As String = "指示器") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    mintIndicator = intCol
    Set micoIndicator = vData
    
    Call ModifyColumn(intCol, "", 255, flexAlignCenterCenter, flexDTString, "", "[" & strIcon & "]", False)
    
    Call ShowIndicatorIcon
    
End Function

Public Function InitializeEditColumn(ByVal intCol As Integer, blnAllowEdit As Boolean, Optional bytEditControl As VsfEditControl = vbVsfEditText, Optional ByVal strDropData As String, Optional ByVal intMaxLength As Integer, Optional ByVal blnShowEditState As Boolean, Optional ByVal strColEditFormat As String = "") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Function
    
    mColProperty(intCol).允许编辑 = blnAllowEdit
    mColProperty(intCol).下拉数据 = strDropData
    mColProperty(intCol).最大长度 = intMaxLength
    mColProperty(intCol).编辑控件 = bytEditControl
    mColProperty(intCol).编辑状态 = blnShowEditState
    mobjVsf.ColEditMask(intCol) = strColEditFormat
    
End Function

Public Function AppendColumn(ByVal strColTitle As String, _
                            ByVal lngColWidth As Long, _
                            Optional ByVal bytColAlignment As AlignmentSettings = flexAlignGeneral, _
                            Optional ByVal bytColDataType As DataTypeSettings = flexDTString, _
                            Optional ByVal strColFormat As String = "", _
                            Optional ByVal strDataTitle As String = "", _
                            Optional ByVal blnAllowColResize As Boolean = True, _
                            Optional ByVal blnAllowColSort As Boolean = True, _
                            Optional ByVal blnAllowColMove As Boolean = True, _
                            Optional ByVal blnColHidden As Boolean = False, _
                            Optional ByVal blnAutoRowHeight As Boolean = False, _
                            Optional ByVal bytColCursorJump As ECursorJump = m未指定) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer
    Dim intPosLeft As Integer
    Dim intPosRight As Integer
    Dim strIcon As String
    Dim objItem As Object
    
    On Error GoTo errHand
    
    mobjVsf.Cols = mobjVsf.Cols + 1
    intCol = mobjVsf.Cols - 1
    
    '重新定义列编辑属性对象
    ReDim Preserve mColProperty(0 To intCol)

    mobjVsf.TextMatrix(0, intCol) = strColTitle
    
    If strDataTitle = "" Then
        mobjVsf.Cell(flexcpData, 0, intCol) = strColTitle
        mColProperty(intCol).允许排序 = blnAllowColSort
        mobjVsf.ColKey(intCol) = IIf(strColTitle = "", strIcon, strColTitle)
    Else
        mobjVsf.Cell(flexcpData, 0, intCol) = strDataTitle
        mColProperty(intCol).允许排序 = blnAllowColSort
        mobjVsf.ColKey(intCol) = strDataTitle
        
        intPosLeft = InStr(strDataTitle, "[")
        intPosRight = InStr(strDataTitle, "]")
        
        If intPosLeft > 0 And intPosRight > 0 And intPosRight > intPosLeft Then
            strIcon = Mid(strDataTitle, intPosLeft + 1, intPosRight - intPosLeft - 1)
            If strIcon <> "" Then
                mobjVsf.ColKey(intCol) = strIcon
                If Not (milsIcon Is Nothing) Then
                    Set objItem = Nothing
                    
                    On Error Resume Next
                    Set objItem = milsIcon.ListImages(strIcon)
                    On Error GoTo errHand
                     
                    If Not (objItem Is Nothing) Then
                        Set mobjVsf.Cell(flexcpPicture, 0, intCol) = objItem.Picture
                    End If
                End If
                mColProperty(intCol).允许排序 = False
            End If
        End If
    End If

    mobjVsf.ColWidth(intCol) = lngColWidth
    mobjVsf.ColAlignment(intCol) = bytColAlignment
    mobjVsf.ColHidden(intCol) = blnColHidden
    mobjVsf.ColDataType(intCol) = bytColDataType
    mobjVsf.ColFormat(intCol) = strColFormat
        
    mColProperty(intCol).系统隐藏 = blnColHidden
    mColProperty(intCol).调整列宽 = blnAllowColResize
    mColProperty(intCol).允许拖动 = blnAllowColMove
    mColProperty(intCol).自动行高 = blnAutoRowHeight
    mColProperty(intCol).系统隐藏 = False
    mColProperty(intCol).自动换行 = False
    If bytColCursorJump = m未指定 Then
        mColProperty(intCol).光标跳跃 = mbytCurSorJumpType
    Else
        mColProperty(intCol).光标跳跃 = bytColCursorJump
    End If
    
    AppendColumn = True
    Exit Function
    
errHand:
    
End Function

Public Function ModifyColumn(ByVal intCol As Integer, _
                            ByVal strColTitle As String, _
                            ByVal lngColWidth As Long, _
                            Optional ByVal bytColAlignment As AlignmentSettings = flexAlignGeneral, _
                            Optional ByVal bytColDataType As DataTypeSettings = flexDTString, _
                            Optional ByVal strColFormat As String = "", _
                            Optional ByVal strDataTitle As String = "", _
                            Optional ByVal blnAllowColResize As Boolean = True, _
                            Optional ByVal blnAllowColSort As Boolean = True, _
                            Optional ByVal blnAllowColMove As Boolean = True, _
                            Optional ByVal blnColHidden As Boolean = False, _
                            Optional ByVal blnAutoRowHeight As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
'    Dim intCol As Integer
    Dim intPosLeft As Integer
    Dim intPosRight As Integer
    Dim strIcon As String
    Dim objItem As Object
    
    On Error GoTo errHand
    
    If intCol < 0 Or intCol > mobjVsf.Cols - 1 Then Exit Function
    
    mobjVsf.TextMatrix(0, intCol) = strColTitle

    If strDataTitle = "" Then
        mobjVsf.Cell(flexcpData, 0, intCol) = strColTitle
        mColProperty(intCol).允许排序 = blnAllowColSort
    Else
        
        mobjVsf.Cell(flexcpData, 0, intCol) = strDataTitle
        mColProperty(intCol).允许排序 = blnAllowColSort
        
        intPosLeft = InStr(strDataTitle, "[")
        intPosRight = InStr(strDataTitle, "]")
        
        If intPosLeft > 0 And intPosRight > 0 And intPosRight > intPosLeft Then
            strIcon = Mid(strDataTitle, intPosLeft + 1, intPosRight - intPosLeft - 1)
            If strIcon <> "" Then
                If Not (milsIcon Is Nothing) Then
                    Set objItem = Nothing
                    
                    On Error Resume Next
                    Set objItem = milsIcon.ListImages(strIcon)
                    On Error GoTo errHand
                     
                    If Not (objItem Is Nothing) Then
                        Set mobjVsf.Cell(flexcpPicture, 0, intCol) = objItem.Picture
                    End If
                End If
                mColProperty(intCol).允许排序 = False
            End If
        End If
    End If
    
    mobjVsf.ColWidth(intCol) = lngColWidth
    mobjVsf.ColAlignment(intCol) = bytColAlignment
    mobjVsf.ColHidden(intCol) = blnColHidden
    mobjVsf.ColDataType(intCol) = bytColDataType
    mobjVsf.ColFormat(intCol) = strColFormat
    
    mColProperty(intCol).调整列宽 = blnAllowColResize
    mColProperty(intCol).允许拖动 = blnAllowColMove
    mColProperty(intCol).自动行高 = blnAutoRowHeight
    
    ModifyColumn = True
    Exit Function
    
errHand:
    
End Function

Public Function SaveStateToString() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    Dim strReturn As String
    Dim intCol As Integer
    
    For intCol = 0 To mobjVsf.Cols - 1
'        strTmp = mobjVsf.TextMatrix(0, intCol)
        strTmp = mobjVsf.ColKey(intCol)
        
        strTmp = strTmp & "'" & mobjVsf.Cell(flexcpData, 0, intCol)
        strTmp = strTmp & "'" & mobjVsf.ColWidth(intCol)
        strTmp = strTmp & "'" & mobjVsf.ColAlignment(intCol)
        strTmp = strTmp & "'" & mobjVsf.ColHidden(intCol)
        strTmp = strTmp & "'" & mobjVsf.ColDataType(intCol)
        strTmp = strTmp & "'" & mobjVsf.ColFormat(intCol)
        strTmp = strTmp & "'" & mColProperty(intCol).允许编辑
        strTmp = strTmp & "'" & mColProperty(intCol).编辑控件
        strTmp = strTmp & "'" & mColProperty(intCol).最大长度
        strTmp = strTmp & "'"
        strTmp = strTmp & "'" & mColProperty(intCol).调整列宽
        strTmp = strTmp & "'" & mColProperty(intCol).允许排序
        strTmp = strTmp & "'" & mColProperty(intCol).允许拖动
        strTmp = strTmp & "'" & mColProperty(intCol).自动行高
        strTmp = strTmp & "'" & mColProperty(intCol).系统隐藏
        strTmp = strTmp & "'" & mColProperty(intCol).编辑状态
        strTmp = strTmp & "'" & mColProperty(intCol).自动换行
        
        strReturn = strReturn & ";" & strTmp
    Next
    
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    SaveStateToString = strReturn
End Function

Public Function LoadStateFromString(ByVal strState As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim varProperty As Variant
    Dim varCol As Variant
    Dim intCol As Integer
    Dim intIndex As Integer, blnAutoLineFeed As Boolean
    
    If strState <> "" Then
        varCol = Split(strState, ";")
        
        For intCol = 0 To UBound(varCol)
            
            varProperty = Split(varCol(intCol), "'")
            
            intIndex = ColIndex(varProperty(0))
            
            If intIndex >= 0 Then
                mobjVsf.ColPosition(intIndex) = intCol
                intIndex = intCol
                
                mobjVsf.ColWidth(intIndex) = Val(varProperty(2))
                mobjVsf.ColAlignment(intIndex) = Val(varProperty(3))
                mobjVsf.ColHidden(intIndex) = varProperty(4)
                mobjVsf.ColDataType(intIndex) = Val(varProperty(5))
                mobjVsf.ColFormat(intIndex) = Trim(varProperty(6))
                
                mColProperty(intIndex).允许编辑 = Trim(varProperty(7))
                mColProperty(intIndex).编辑控件 = Val(varProperty(8))
                mColProperty(intIndex).最大长度 = Val(varProperty(9))
                'mColProperty(intIndex).下拉数据 = Val(varProperty(10))
                mColProperty(intIndex).调整列宽 = Trim(varProperty(11))
                mColProperty(intIndex).允许排序 = Trim(varProperty(12))
                mColProperty(intIndex).允许拖动 = Trim(varProperty(13))
                mColProperty(intIndex).自动行高 = Trim(varProperty(14))
                mColProperty(intIndex).系统隐藏 = Trim(varProperty(15))
                If UBound(varProperty) >= 16 Then mColProperty(intIndex).编辑状态 = Trim(varProperty(16))
                If UBound(varProperty) >= 17 And blnAutoLineFeed = False Then
                    mColProperty(intIndex).自动换行 = Trim(varProperty(17))
                    blnAutoLineFeed = mColProperty(intIndex).自动换行
                End If
                
                If UBound(varProperty) >= 18 Then
                    If 0 <= Val(Trim(varProperty(18))) <= 3 Then
                        mColProperty(intIndex).光标跳跃 = Val(Trim(varProperty(18)))
                    End If
                End If
            End If
        Next
    End If
    
    LoadStateFromString = True
    
End Function

Public Function AfterMoveColumn(ByVal intCol As Integer, ByVal intPostion As Integer) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim objColTmp As ColPropertys
        
    objColTmp.允许编辑 = mColProperty(intCol).允许编辑
    objColTmp.编辑控件 = mColProperty(intCol).编辑控件
    objColTmp.最大长度 = mColProperty(intCol).最大长度
    objColTmp.下拉数据 = mColProperty(intCol).下拉数据
    objColTmp.调整列宽 = mColProperty(intCol).调整列宽
    objColTmp.允许排序 = mColProperty(intCol).允许排序
    objColTmp.允许拖动 = mColProperty(intCol).允许拖动
    objColTmp.自动行高 = mColProperty(intCol).自动行高
    objColTmp.系统隐藏 = mColProperty(intCol).系统隐藏
    objColTmp.编辑状态 = mColProperty(intCol).编辑状态
    objColTmp.自动换行 = mColProperty(intCol).自动换行
    objColTmp.光标跳跃 = mColProperty(intCol).光标跳跃
    
    If intCol > intPostion Then
        '向前移动列
        For intLoop = intCol To intPostion + 1 Step -1
            mColProperty(intLoop).允许编辑 = mColProperty(intLoop - 1).允许编辑
            mColProperty(intLoop).编辑控件 = mColProperty(intLoop - 1).编辑控件
            mColProperty(intLoop).最大长度 = mColProperty(intLoop - 1).最大长度
            mColProperty(intLoop).下拉数据 = mColProperty(intLoop - 1).下拉数据
            mColProperty(intLoop).调整列宽 = mColProperty(intLoop - 1).调整列宽
            mColProperty(intLoop).允许排序 = mColProperty(intLoop - 1).允许排序
            mColProperty(intLoop).允许拖动 = mColProperty(intLoop - 1).允许拖动
            mColProperty(intLoop).自动行高 = mColProperty(intLoop - 1).自动行高
            mColProperty(intLoop).系统隐藏 = mColProperty(intLoop - 1).系统隐藏
            mColProperty(intLoop).编辑状态 = mColProperty(intLoop - 1).编辑状态
            mColProperty(intLoop).自动换行 = mColProperty(intLoop - 1).自动换行
            mColProperty(intLoop).光标跳跃 = mColProperty(intLoop - 1).光标跳跃
        Next
    ElseIf intCol < intPostion Then
        '向后移动列
        For intLoop = intCol To intPostion - 1
            mColProperty(intLoop).允许编辑 = mColProperty(intLoop + 1).允许编辑
            mColProperty(intLoop).编辑控件 = mColProperty(intLoop + 1).编辑控件
            mColProperty(intLoop).最大长度 = mColProperty(intLoop + 1).最大长度
            mColProperty(intLoop).下拉数据 = mColProperty(intLoop + 1).下拉数据
            mColProperty(intLoop).调整列宽 = mColProperty(intLoop + 1).调整列宽
            mColProperty(intLoop).允许排序 = mColProperty(intLoop + 1).允许排序
            mColProperty(intLoop).允许拖动 = mColProperty(intLoop + 1).允许拖动
            mColProperty(intLoop).自动行高 = mColProperty(intLoop + 1).自动行高
            mColProperty(intLoop).系统隐藏 = mColProperty(intLoop + 1).系统隐藏
            mColProperty(intLoop).编辑状态 = mColProperty(intLoop + 1).编辑状态
            mColProperty(intLoop).自动换行 = mColProperty(intLoop + 1).自动换行
            mColProperty(intLoop).光标跳跃 = mColProperty(intLoop + 1).光标跳跃
        Next
    End If

    mColProperty(intPostion).允许编辑 = objColTmp.允许编辑
    mColProperty(intPostion).编辑控件 = objColTmp.编辑控件
    mColProperty(intPostion).最大长度 = objColTmp.最大长度
    mColProperty(intPostion).下拉数据 = objColTmp.下拉数据
    mColProperty(intPostion).调整列宽 = objColTmp.调整列宽
    mColProperty(intPostion).允许排序 = objColTmp.允许排序
    mColProperty(intPostion).允许拖动 = objColTmp.允许拖动
    mColProperty(intPostion).自动行高 = objColTmp.自动行高
    mColProperty(intPostion).系统隐藏 = objColTmp.系统隐藏
    mColProperty(intPostion).编辑状态 = objColTmp.编辑状态
    mColProperty(intPostion).自动换行 = objColTmp.自动换行
    mColProperty(intPostion).光标跳跃 = objColTmp.光标跳跃

    AfterMoveColumn = True
    
End Function

Public Function AfterRowColChange(ByVal intOldRow As Integer, ByVal intOldCol As Integer, ByVal intNewRow As Integer, ByVal intNewCol As Integer) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnCancel As Boolean
    Dim ComboList As String
    
    If intNewRow < 0 Or intNewCol < 0 Then Exit Function
    If mblnSelectTrace Then
        If mobjVsf.Cell(flexcpBackColor, intNewRow, intNewCol) <> 16772055 Then
            Call SelectRow(intOldRow, intNewRow, 16772055)
        End If
    End If

    If mblnLoading Then Exit Function

    '显示指示器
     Call ShowIndicatorIcon

    mobjVsf.FocusRect = IIf(mColProperty(intNewCol).允许编辑 And AllowEdit, flexFocusSolid, flexFocusHeavy)

    If mColProperty(intNewCol).允许编辑 = False Or AllowEdit = False Then
        mobjVsf.ColComboList(intNewCol) = ""
    Else
        Select Case mColProperty(intNewCol).编辑控件
        Case 1
            mobjVsf.ColComboList(intNewCol) = ""
        Case 2
            mobjVsf.ColComboList(intNewCol) = "..."
        Case 3
            ComboList = mColProperty(intNewCol).下拉数据
            '下拉,传入记录集
            blnCancel = False

            RaiseEvent BeforeComboList(intOldRow, intNewCol, ComboList, blnCancel)
            If blnCancel = False Then mobjVsf.ColComboList(intNewCol) = ComboList
            If ComboList = "" Then ComboList = " |"
            mobjVsf.ColComboList(intNewCol) = ComboList

        Case Else
            mobjVsf.ColComboList(intNewCol) = ""
        End Select

        Call ShowEditState

    End If
    
End Function

Public Function SumVsfCol(ByVal intCol As Integer) As Double
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    SumVsfCol = mobjVsf.Aggregate(flexSTSum, mobjVsf.FixedRows, intCol, mobjVsf.Rows - 1, intCol)

End Function

Public Sub SetFocus(Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1, Optional ByVal blnDoevents As Boolean = False)
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If lngRow <> -1 Then mobjVsf.Row = lngRow
    If lngCol <> -1 Then mobjVsf.Col = lngCol
    
    If blnDoevents Then DoEvents
    
    mobjVsf.SetFocus
    mobjVsf.ShowCell mobjVsf.Row, mobjVsf.Col
    
errHand:
        
End Sub

Public Function LoadCrossGrid(ByVal rsData As ADODB.Recordset, ByVal aryPortraitField As Variant, Optional ByVal strLandscapeField As String, Optional ByVal strValueField As String, Optional ByVal blnClear As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:填充数据到交叉网格
    '参数:
    '返回:
    '示例:aryPortraitField="病人id,姓名,性别,健康号";aryLandscapeField="体检项目";aryValueField="体检结果"
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngFixedCol() As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strField As String
    Dim blnField As Boolean
    Dim strIcon As String
    Dim objItem As Object
    
    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    mblnLoading = True
    ReDim lngFixedCol(0 To UBound(aryPortraitField))
    
    '清除显示数据
    '------------------------------------------------------------------------------------------------------------------
    If blnClear Then
        mobjVsf.Rows = 2
        mobjVsf.RowData(1) = 0
        mobjVsf.Cell(flexcpText, 1, 0, 1, mobjVsf.Cols - 1) = ""
    Else
        If Val(mobjVsf.RowData(mobjVsf.Rows - 1)) <= 0 Then
            lngRow = mobjVsf.Rows - 2
        Else
            lngRow = mobjVsf.Rows - 1
        End If
    End If

    For lngLoop = 0 To UBound(aryPortraitField)
        For lngCol = 0 To mobjVsf.Cols - 1
            If mobjVsf.Cell(flexcpData, 0, lngCol) = aryPortraitField(lngLoop) Then
                lngFixedCol(lngLoop) = lngCol
                Exit For
            End If
        Next
    Next
    
        
    '填写纵向数据
    '------------------------------------------------------------------------------------------------------------------
    Do While Not rsData.EOF
        
        lngRow = FindRow(Val(Nvl(rsData("ID").Value)))
        
        If lngRow = -1 Then
            If Val(mobjVsf.RowData(mobjVsf.Rows - 1)) > 0 Then mobjVsf.Rows = mobjVsf.Rows + 1
            lngRow = mobjVsf.Rows - 1
            
            '纵向数据
            mobjVsf.RowData(lngRow) = CStr(Nvl(rsData("ID").Value))
            
            For lngLoop = 0 To UBound(aryPortraitField)
                
                lngCol = lngFixedCol(lngLoop)
                strField = mobjVsf.Cell(flexcpData, 0, lngCol)
                blnField = False

                If Left(strField, 1) = "[" Then
                
                    strField = Mid(strField, 2, Len(strField) - 2)
                    strIcon = ""
                    
                    On Error Resume Next
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    On Error GoTo errHand
                    
                    If blnField Then
                        If Not (milsIcon Is Nothing) Then
                            strIcon = Nvl(rsData(strField))
                            If strIcon <> "" Then
                            
                                Set objItem = Nothing
                                
                                On Error Resume Next
                                Set objItem = milsIcon.ListImages(strIcon)
                                On Error GoTo errHand
                                
                                If Not (objItem Is Nothing) Then
                                    Set mobjVsf.Cell(flexcpPicture, lngRow, lngCol, lngRow, lngCol) = objItem.Picture
                                Else
                                    Set mobjVsf.Cell(flexcpPicture, lngRow, lngCol, lngRow, lngCol) = Nothing
                                End If
                            End If
                        End If
                        
                        mobjVsf.TextMatrix(lngRow, lngCol) = Nvl(rsData(strField).Value)
                        mobjVsf.Cell(flexcpData, lngRow, lngCol) = strIcon
                    End If
                Else
                    On Error Resume Next
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    On Error GoTo errHand
                    
                    If blnField Then
                        mobjVsf.TextMatrix(lngRow, lngCol) = Nvl(rsData(strField).Value)
                    End If
                End If

            Next
            
        End If
        
        '横向数据
        strField = Nvl(rsData(strLandscapeField).Value)
        lngCol = -1
        For lngLoop = 0 To mobjVsf.Cols - 1
            If strField = mobjVsf.Cell(flexcpData, 0, lngLoop) Then
                lngCol = lngLoop
                Exit For
            End If
        Next
        
        If lngCol > 0 And lngRow > 0 Then
            mobjVsf.TextMatrix(lngRow, lngCol) = Nvl(rsData(strValueField).Value)
            'mobjVsf.Cell(flexcpData, lngRow, lngCol) = 1
            mobjVsf.Cell(flexcpBackColor, lngRow, lngCol) = COLOR.浅黄色
        End If
        
        rsData.MoveNext
    Loop


    '恢复设置
    '------------------------------------------------------------------------------------------------------------------
    If mstrSvrKey <> "" Then Call RestoreRow(mstrSvrKey)
    If mblnAppendRows Then Call DrawSpaceCellLine
    If mblnAllowEdit And mblnSelectTrace Then Call SelectRow(-1, 1, 16772055)

    mblnLoading = False
    LoadCrossGrid = True
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mblnLoading = False
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function DrawIcon(ByVal strIcon As String, ByVal lngStartRow As Long, ByVal lngEndRow As Long, ByVal lngCol As Long) As Boolean
    '******************************************************************************************************************
    '功能:填充数据到网格
    '参数:
    '返回:
    '******************************************************************************************************************
    
    Dim objItem As Object
    
    If strIcon <> "" Then

        On Error Resume Next
        Set objItem = milsIcon.ListImages(strIcon)
        On Error GoTo endHand

        If Not (objItem Is Nothing) Then
            Set mobjVsf.Cell(flexcpPicture, lngStartRow, lngCol, lngEndRow, lngCol) = objItem.Picture
            mobjVsf.Cell(flexcpData, lngStartRow, lngCol, lngEndRow, lngCol) = strIcon
        End If
        
    End If
    
endHand:
    
End Function

Public Function LoadGrid(ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal strStrikeRows As String = "") As Boolean
    '******************************************************************************************************************
    '功能:填充数据到网格
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    
    mblnLoading = True
    If blnClear Then
        mobjVsf.Rows = 2
        mobjVsf.RowData(1) = ""
        For lngCol = 0 To mobjVsf.Cols - 1
            mobjVsf.TextMatrix(1, lngCol) = ""
        Next
        lngRow = 0
    Else
        If Val(mobjVsf.RowData(mobjVsf.Rows - 1)) <= 0 Then
            lngRow = mobjVsf.Rows - 2
        Else
            lngRow = mobjVsf.Rows - 1
        End If
    End If
    
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If mobjVsf.Rows < lngRow + 1 Then mobjVsf.Rows = lngRow + 1
        
        mobjVsf.Cell(flexcpForeColor, lngRow, 0, lngRow, mobjVsf.Cols - 1) = vbBlack
        If LoadGridRow(lngRow, rsData, MaskArray) = False Then GoTo errHand
        '冲销行显示为蓝色
        If strStrikeRows <> "" Then
            If InStr(1, "," & strStrikeRows & ",", "," & lngRow & ",") > 0 Then
                mobjVsf.Cell(flexcpForeColor, lngRow, 0, lngRow, mobjVsf.Cols - 1) = vbBlue
            End If
        End If
        
        rsData.MoveNext
    Loop
    
    If mstrSvrKey <> "" Then Call RestoreRow(mstrSvrKey)
        
    If mblnSelectTrace Then

        If mobjVsf.Cell(flexcpBackColor, 1, 1) <> 16772055 Then
            Call SelectRow(-1, 1, 16772055)
        End If

    End If
    
    If AutoRowHeight Then
        For lngCol = 0 To mobjVsf.Cols - 1
            If mColProperty(lngCol).自动行高 Then
                Call mobjVsf.AutoSize(lngCol, lngCol)
            End If
        Next
    End If
    
    If mblnAppendRows Then Call DrawSpaceCellLine
    
    mblnLoading = False
    LoadGrid = True
    Exit Function
    
errHand:
    mblnLoading = False
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function LoadGridRow(ByVal lngRow As Long, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    
    Dim lngCol As Long
    Dim strMask As String
    Dim strField As String
    Dim strIcon As String
    Dim blnField As Boolean
    Dim objItem As Object
    Dim blnForeColor As Boolean
    Dim blnBkColor As Boolean

    On Error Resume Next
    
    mobjVsf.RowData(lngRow) = CStr(Nvl(rsData("ID")))
    blnForeColor = (rsData("前景色").Name = "前景色")
    blnBkColor = (rsData("背景色").Name = "背景色")
    
    On Error GoTo errHand
 
    For lngCol = 0 To mobjVsf.Cols - 1
            
        strField = mobjVsf.Cell(flexcpData, 0, lngCol)
            
        If Trim(strField) <> "" Then
        
            On Error Resume Next
            strMask = ""
            strMask = MaskArray(lngCol)
            On Error GoTo errHand
                
            If Left(strField, 1) = "[" Then
            
                strField = Mid(strField, 2, Len(strField) - 2)
                strIcon = ""
                
                On Error Resume Next
                blnField = False
                blnField = (UCase(rsData(strField).Name) = UCase(strField))
                If blnField = False Then GoTo NextCol
                On Error GoTo errHand
                    
                If Not (milsIcon Is Nothing) Then
                    strIcon = Nvl(rsData(strField))
                    If strIcon <> "" Then
                    
                        Set objItem = Nothing
                        
                        On Error Resume Next
                        Set objItem = milsIcon.ListImages(strIcon)
                        On Error GoTo errHand
                        
                        If Not (objItem Is Nothing) Then
                            Set mobjVsf.Cell(flexcpPicture, lngRow, lngCol, lngRow, lngCol) = objItem.Picture
                        Else
                            Set mobjVsf.Cell(flexcpPicture, lngRow, lngCol, lngRow, lngCol) = Nothing
                        End If
                    End If
                End If
                
                mobjVsf.TextMatrix(lngRow, lngCol) = Nvl(rsData(strField).Value)
                mobjVsf.Cell(flexcpData, lngRow, lngCol) = strIcon
            Else
                
                On Error Resume Next
                blnField = False
                blnField = (UCase(rsData(strField).Name) = UCase(strField))
                If blnField = False Then GoTo NextCol
                On Error GoTo errHand
                
                If strMask <> "" Then
                    mobjVsf.TextMatrix(lngRow, lngCol) = Format(Nvl(rsData(strField)), strMask)
                Else
                    mobjVsf.TextMatrix(lngRow, lngCol) = Nvl(rsData(strField))
                End If
            
                mobjVsf.Cell(flexcpData, lngRow, lngCol, lngRow, lngCol) = mobjVsf.TextMatrix(lngRow, lngCol)
            End If
            
        End If
NextCol:
            '下一列
    Next
    
    If blnForeColor Then mobjVsf.Cell(flexcpForeColor, lngRow, 0, lngRow, mobjVsf.Cols - 1) = Val(rsData("前景色").Value)
    If blnBkColor Then mobjVsf.Cell(flexcpBackColor, lngRow, 0, lngRow, mobjVsf.Cols - 1) = Val(rsData("背景色").Value)
    
    LoadGridRow = True
    
    Exit Function
    
errHand:

End Function

Public Function RestoreRow(ByVal strKey As String, Optional ByVal intCol As Integer = -1) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    RestoreRow = True
    
    For lngLoop = 1 To mobjVsf.Rows - 1
        If intCol = -1 Then
            If mobjVsf.RowData(lngLoop) = strKey And mobjVsf.Row <> lngLoop Then
                mobjVsf.Row = lngLoop
                Exit Function
            End If
        Else
            If mobjVsf.TextMatrix(lngLoop, intCol) = strKey And mobjVsf.Row <> lngLoop Then
                mobjVsf.Row = lngLoop
                Exit Function
            End If
        End If
    Next
    
    RestoreRow = False
End Function

Public Function FindRow(ByVal strKey As String, _
                        Optional ByVal intCol As Integer = -1, _
                        Optional ByVal bytMatch As Byte = 0, _
                        Optional ByVal lngStartRow As Long = -1) As Long
    '******************************************************************************************************************
    '功能:
    '参数:bytMatch=0，完全匹配；=1，左向匹配；=2；包含匹配；=3，右向匹配
    '返回:
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngLoop As Long
    
    If lngStartRow = -1 Then lngStartRow = mobjVsf.FixedRows
    lngRow = -1
    
    If intCol = -1 Then
        lngRow = mobjVsf.FindRow(strKey)
    Else
        Select Case bytMatch
        Case 0
            lngRow = mobjVsf.FindRow(strKey, lngStartRow, intCol, False, True)
        Case 1
            lngRow = mobjVsf.FindRow(strKey, lngStartRow, intCol, False, False)
        Case 2
            For lngLoop = lngStartRow To mobjVsf.Rows - 1
                If InStr(UCase(mobjVsf.TextMatrix(lngLoop, intCol)), UCase(strKey)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        Case 3
            For lngLoop = lngStartRow To mobjVsf.Rows - 1
                If Right(UCase(mobjVsf.TextMatrix(lngLoop, intCol)), Len(strKey)) = UCase(strKey) Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End Select
    End If
        
    FindRow = lngRow
    
End Function

Public Function FindGridRow(ByVal strKey As String, Optional ByVal intCol As Integer = -1) As Long
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngLoop As Long
        
    For lngLoop = 1 To mobjVsf.Rows - 1
        If intCol = -1 Then
            If mobjVsf.RowData(lngLoop) = strKey Then
                FindGridRow = lngLoop
                Exit Function
            End If
        Else
            If mobjVsf.TextMatrix(lngLoop, intCol) = strKey Then
                FindGridRow = lngLoop
                Exit Function
            End If
        End If
    Next
    FindGridRow = -1
End Function

Public Function ClearGrid(Optional ByVal Row As Long = 1) As Boolean
    '******************************************************************************************************************
    '功能:清除表格数据
    '
    '
    '******************************************************************************************************************
    Err = 0
    On Error GoTo errHand
    
    '先保存当前行
    
    mstrSvrKey = mobjVsf.RowData(mobjVsf.Row)
    
    mobjVsf.Rows = Row + 1
    mobjVsf.RowData(Row) = 0
    mobjVsf.Cell(flexcpText, Row, 0, Row, mobjVsf.Cols - 1) = ""
    mobjVsf.Cell(flexcpBackColor, Row, 0, Row, mobjVsf.Cols - 1) = mobjVsf.BackColor
    
    If mblnAllowEdit = False Then Set mobjVsf.Cell(flexcpPicture, Row, 0, Row, mobjVsf.Cols - 1) = Nothing
    
    If mblnAppendRows Then Call DrawSpaceCellLine
    
    If Err = 0 Then ClearGrid = True
    
    Exit Function
    
errHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CountDataRows() As Long
    '******************************************************************************************************************
    '功能:清除表格数据
    '
    '
    '******************************************************************************************************************
    Dim lngCount As Long
    Dim lngLoop As Long
    
    For lngLoop = 1 To mobjVsf.Rows - 1
        If Val(mobjVsf.RowData(lngLoop)) > 0 Then
            lngCount = lngCount + 1
        End If
    Next
    
    CountDataRows = lngCount
End Function

Public Function ClearColumn() As Boolean
    mobjVsf.Cols = 0
    ClearColumn = True
End Function

Public Sub AppendRow()

    If mobjVsf.Row <> mobjVsf.Rows - 1 Then mobjVsf.Row = mobjVsf.Rows - 1
    If mobjVsf.Col <> mobjVsf.Cols - 1 Then mobjVsf.Col = mobjVsf.Cols - 1
    
    Call LocationNextCell

End Sub

Public Sub DeleteRow(ByVal intRow As Integer)
    If mobjVsf.Rows - 1 < intRow Then Exit Sub
    
    If mobjVsf.Row <> intRow Then mobjVsf.Row = intRow
    
    Call KeyDown(vbKeyDelete, 0)
End Sub

Public Sub AutoAddRow(ByVal intRow As Integer, ByVal intCol As Integer)
    Dim blnCancel As Boolean
    
    If mblnAllowNewRow = False Then Exit Sub
    
    If intRow <> -1 Then Exit Sub

    '最后一行是否已输入完整
    RaiseEvent BeforeNewRow(mobjVsf.Rows - 1, mobjVsf.Col, blnCancel)
    If blnCancel Then Exit Sub
    
    '是否是最后一行的下一行
    
    mobjVsf.Rows = mobjVsf.Rows + 1
    
    If mblnAppendRows Then Call DrawSpaceCellLine
    mobjVsf.Row = mobjVsf.Rows - 1
    If intCol > -1 Then mobjVsf.Col = intCol
    RaiseEvent AfterNewRow(mobjVsf.Row, mobjVsf.Col)
    mobjVsf.SetFocus
    RaiseEvent RsetFocuse("AfterNewRow")
End Sub

Public Function CheckHave(ByVal strKey As String, Optional ByVal blnNotCurrentRow As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:  检查是否有重复的项目
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    For lngLoop = 1 To mobjVsf.Rows - 1
        If Trim(mobjVsf.RowData(lngLoop)) = strKey And ((mobjVsf.Row <> lngLoop And blnNotCurrentRow) Or blnNotCurrentRow = False) Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Public Function SetRowForeColor(ByVal intRow As Integer, ByVal lngColor As Long, Optional ByVal aryRow As Variant) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************

    mobjVsf.Cell(flexcpForeColor, intRow, 0, intRow, mobjVsf.Cols - 1) = lngColor
    
End Function

Public Sub SelectRow(ByVal OldRow As Long, ByVal NewRow As Long, Optional ByVal lngBackColor As Long = -1)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngColor As Long
    
    On Error Resume Next
    
    If lngBackColor = -1 Then
        lngColor = mobjVsf.BackColorSel
    Else
        lngColor = lngBackColor
    End If
    
    If OldRow + 1 > mobjVsf.FixedRows Then
'        mobjVsf.Cell(flexcpBackColor, OldRow, mobjVsf.FixedCols, OldRow, mobjVsf.Cols - 1) = mobjVsf.BackColor
        mobjVsf.Cell(flexcpBackColor, mobjVsf.FixedRows, mobjVsf.FixedCols, mobjVsf.Rows - 1, mobjVsf.Cols - 1) = mobjVsf.BackColor
    End If
    
    If NewRow + 1 > mobjVsf.FixedRows Then
        mobjVsf.Cell(flexcpBackColor, NewRow, mobjVsf.FixedCols, NewRow, mobjVsf.Cols - 1) = lngColor
    End If
    
End Sub

Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim blnCancel As Boolean
    
    On Error GoTo errHand
    
    If mblnAllowEdit = False Then Exit Sub
    
    Select Case KeyCode
    Case vbKeyDelete
        If Shift = 0 And mblnAllowDeleteRow Then
            '删除整行及内容
            
            blnCancel = False
            
            RaiseEvent BeforeDeleteRow(mobjVsf.Row, mobjVsf.Col, blnCancel)
            
            If blnCancel Then Exit Sub
            
            If mobjVsf.Rows > 1 Then
                If mobjVsf.Rows = 2 And mobjVsf.Row = 1 Then
                    Call ClearGrid
                Else
                    mobjVsf.RemoveItem mobjVsf.Row
                    If mblnAppendRows Then Call DrawSpaceCellLine
                End If
                Call AdjustRowFlag(mobjVsf.Row)
                
                RaiseEvent AfterDeleteRow(mobjVsf.Row, mobjVsf.Col)
            End If
            
        ElseIf Shift = 2 And mColProperty(mobjVsf.Col).允许编辑 Then
            '删除当前单元格的内容
            
            blnCancel = False
            RaiseEvent BeforeDeleteCell(mobjVsf.Row, mobjVsf.Col, blnCancel)
            If blnCancel Then Exit Sub
            
            mobjVsf.TextMatrix(mobjVsf.Row, mobjVsf.Col) = ""
            
            RaiseEvent AfterDeleteCell(mobjVsf.Row, mobjVsf.Col)
            
        End If
        
    Case Else
        If mColProperty(mobjVsf.Col).编辑控件 = 2 And Shift = 0 Then
            mobjVsf.ColComboList(mobjVsf.Col) = ""
        End If
    End Select
    
    Exit Sub
    
errHand:
        
End Sub

Public Sub KeyPress(KeyAscii As Integer)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call LocationNextCell
    Else
        If mblnAllowEdit = False Then Exit Sub
        If mColProperty(mobjVsf.Col).编辑控件 = 2 Then mobjVsf.ColComboList(mobjVsf.Col) = ""
    End If
End Sub

Public Sub KeyPressEdit(KeyAscii As Integer, Optional ByVal blnCursorJump As Boolean = False)
    '******************************************************************************************************************
    '功能：
    '参数：blnCursorJump:是否进入光标跳转
    '返回：
    '******************************************************************************************************************
    '进行下一个焦点定位，若找不到，且允许新增行时，自动增加一行
    Dim intCol As Integer
    If mblnAllowEdit = False Then Exit Sub
    If mColProperty(mobjVsf.Col).允许编辑 = False Then Exit Sub
    If blnCursorJump = False Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        intCol = GetAutoLineFeedCol
        If intCol <> -1 Then
            If mColProperty(intCol).光标跳跃 <> m强制经过 And intCol = mobjVsf.Col Then mColProperty(intCol).光标跳跃 = m强制经过
        End If
        Call LocationNextCell(True)
    End If
End Sub
    
Public Sub DbClick()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If mblnAllowEdit = False Then Exit Sub
    
    Call KeyPress(32)

End Sub

Private Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal X As Single, ByVal Y As Single)
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

Public Sub SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim lp As POINTAPI
    Dim lp1 As POINTAPI

    With mobjVsf
        
        If mblnAllowEdit = False Then Exit Sub
        
        If mColProperty(Col).编辑状态 And mColProperty(Col).允许编辑 Then
            Call GetCursorPos(lp)
            ClientToScreen .hWnd, lp1
            
            Call SendLMouseButton(EditWindow, lp.X * 15 - .Cell(flexcpLeft, Row, Col, Row, Col) - lp1.X * 15, lp.Y * 15 - .Cell(flexcpTop, Row, Col, Row, Col) - lp1.Y * 15)
            Call SendLMouseButton(EditWindow, lp.X * 15 - .Cell(flexcpLeft, Row, Col, Row, Col) - lp1.X * 15, lp.Y * 15 - .Cell(flexcpTop, Row, Col, Row, Col) - lp1.Y * 15)
        Else
            .EditSelStart = 0
            .EditSelLength = gobjCommFun.ActualLen(.EditText)
        End If
    End With
End Sub

Public Sub EditSelAll()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If mblnAllowEdit = False Then Exit Sub
    
    mobjVsf.EditSelStart = 0
    mobjVsf.EditSelLength = gobjCommFun.ActualLen(mobjVsf.EditText)
End Sub

Private Sub ShowEditState()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    If mblnAllowEdit = False Then Exit Sub
    
    With mobjVsf
        If mColProperty(.Col).编辑状态 And mColProperty(.Col).允许编辑 Then
            .EditCell
        End If
    End With
    
End Sub

Public Sub GotFocus()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Call ShowEditState
    
End Sub

Public Sub ChangeEdit(ByVal Row As Long, ByVal Col As Long)
    
    If mColProperty(Col).自动行高 And AutoRowHeight Then
        mobjVsf.TextMatrix(Row, Col) = mobjVsf.EditText
        Call mobjVsf.AutoSize(Col, Col)
        If mblnAppendRows Then Call DrawSpaceCellLine
    End If
    
End Sub

Public Sub AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If mblnAllowEdit = False Then Exit Sub
    
    mlngStartPos = mobjVsf.EditSelStart
    
    If mobjVsf.ColComboList(Col) = "" And mColProperty(Col).允许编辑 Then
        Select Case mColProperty(Col).编辑控件
        Case 2
            mobjVsf.ColComboList(Col) = "..."
        Case 3
            mobjVsf.ColComboList(Col) = " |"
        End Select
    End If
    
    If mColProperty(Col).允许编辑 Then
        If mColProperty(Col).自动行高 And AutoRowHeight Then
            mobjVsf.AutoSizeMode = flexAutoSizeRowHeight
            Call mobjVsf.AutoSize(Col, Col)
            If mblnAppendRows Then Call DrawSpaceCellLine
        End If
    End If
    
    If mColProperty(Col).编辑控件 = 2 Then
        mobjVsf.TextMatrix(Row, Col) = mobjVsf.Cell(flexcpData, Row, Col)
    End If
    
    mblnEditing = False
End Sub

Public Sub AfterResizeColumn(ByVal Col As Long)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngCol As Long
    
    If AutoRowHeight Then
        For lngCol = 0 To mobjVsf.Cols - 1
            If mColProperty(lngCol).自动行高 Then
                Call mobjVsf.AutoSize(lngCol, lngCol)
            End If
        Next
        If mblnAppendRows Then Call DrawSpaceCellLine
    End If
    
End Sub

Public Sub BeforeResizeColumn(ByVal Col As Long, Cancel As Boolean)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Cancel = Not mColProperty(Col).调整列宽
End Sub

Public Sub BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Cancel = Not mblnAllowEdit
    If Cancel Then Exit Sub
    
    '检查是否允许编辑，并作相应的处理
    Cancel = Not mColProperty(Col).允许编辑
    If Cancel Then Exit Sub
    
    '先保存原来的内容，在弹出选择取消或没有数据时恢复
    mobjVsf.Cell(flexcpData, Row, Col) = mobjVsf.TextMatrix(Row, Col)
        
    '设置可编辑时的长度
    mobjVsf.EditMaxLength = mColProperty(Col).最大长度
    
    mblnEditing = True
End Sub

Public Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngCol As Long
    
    With mobjVsf
        lngCol = .MouseCol
        If lngCol >= 0 And lngCol <= .Cols - 1 Then
            
            If mColProperty(lngCol).允许编辑 And mblnAllowEdit Then
                .MousePointer = IIf(mColProperty(lngCol).编辑状态, flexIBeam, flexDefault)
            Else
                .MousePointer = flexDefault
            End If
            
        Else
            .MousePointer = flexDefault
        End If
    End With
    
End Sub

Public Sub ValidateEdit(ByVal Col As Long, Cancel As Boolean)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Cancel = Not gobjCommFun.StrIsValid(mobjVsf.EditText, mColProperty(Col).最大长度)
End Sub

Public Function MoveRow(ByVal intCurRow As Integer, Optional ByVal intMove As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim strTmp As String, intflexChecked As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    
    On Error GoTo errHand
    
    If intMove > 0 Then
        For intRow = intCurRow + intMove To mobjVsf.Rows - 1
            If mobjVsf.RowHidden(intRow) = False Then
                intMove = intRow - intCurRow
                Exit For
            End If
        Next
        If intRow = mobjVsf.Rows Then Exit Function
    Else
        For intRow = intCurRow + intMove To 1 Step -1
            If mobjVsf.RowHidden(intRow) = False Then
                intMove = intRow - intCurRow
                Exit For
            End If
        Next
        If intRow = 0 Then Exit Function
    End If

    strTmp = CStr(mobjVsf.RowData(intCurRow))
    mobjVsf.RowData(intCurRow) = mobjVsf.RowData(intCurRow + intMove)
    mobjVsf.RowData(intCurRow + intMove) = Val(strTmp)
    
    For intCol = 0 To mobjVsf.Cols - 1
        
        strTmp = mobjVsf.TextMatrix(intCurRow, intCol)
        intflexChecked = mobjVsf.Cell(flexcpChecked, intCurRow, intCol)
        mobjVsf.TextMatrix(mobjVsf.Row, intCol) = mobjVsf.TextMatrix(intCurRow + intMove, intCol)
        mobjVsf.Cell(flexcpChecked, mobjVsf.Row, intCol) = mobjVsf.Cell(flexcpChecked, intCurRow + intMove, intCol)
        mobjVsf.TextMatrix(intCurRow + intMove, intCol) = strTmp
        mobjVsf.Cell(flexcpChecked, intCurRow + intMove, intCol) = intflexChecked
    Next
    
    MoveRow = True
    
    Exit Function
    
errHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub ClearSortIcon()
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intCol As Integer
    
    For intCol = 0 To mobjVsf.Cols - 1
        If mColProperty(intCol).允许排序 Then
            Set mobjVsf.Cell(flexcpPicture, 0, intCol, 0, intCol) = Nothing
        End If
    Next
End Sub

Private Sub AdjustRowFlag(ByVal intRow As Integer)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    If mobjVsf.FixedCols = 0 Then Exit Sub
    
    If Not (mobjVsf.Cell(flexcpPicture, intRow, 0) Is Nothing) Then Exit Sub
    Set mobjVsf.Cell(flexcpPicture, 1, 0, mobjVsf.Rows - 1, 0) = Nothing
'    Set mobjVsf.Cell(flexcpPicture, intRow, 0) = ils16.ListImages(1).Picture
    
End Sub

Public Function LocationNextCell(Optional ByVal blnEdit As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：blnEdit:表格列是否处于编辑状态
    '返回：
    '******************************************************************************************************************
    Dim lngCol As Long, lngCurCol As Long, lngRow As Long
    Dim blnCancel As Boolean
    Dim blnAutoFeeLine As Boolean, blnAutoSetLine As Boolean  '自动换行
    
    '91648:如果设置了自动换行，则自动完成换行处理(自动换行必须是可编辑的列，且输入了内容的列)
    If mColProperty(mobjVsf.Col).自动换行 = True And mColProperty(mobjVsf.Col).允许编辑 And mobjVsf.ColHidden(mobjVsf.Col) = False And AllowEdit Then
        If blnEdit = False Then
            blnAutoSetLine = (Trim(mobjVsf.TextMatrix(mobjVsf.Row, mobjVsf.Col)) <> "")
        Else
            blnAutoSetLine = (Trim(mobjVsf.EditText) <> "")
        End If
        If blnAutoSetLine = False And mColProperty(mobjVsf.Col).编辑控件 = vbVsfEditCombox Then
            blnAutoSetLine = (Trim(mobjVsf.EditText) <> "")
        End If
        lngCurCol = mobjVsf.Col
    End If
    blnAutoFeeLine = blnAutoSetLine
    If blnAutoFeeLine = False Then
        blnAutoFeeLine = (GetAllowCol(mobjVsf.Row, mobjVsf.Col + 1) > mobjVsf.Cols - 1)
    End If
    
    '说明：当表格存在内容校验(ValidateEdit), 列和行同时发生改变时会触发校验,当通过.ROW,.COL方式切换列，且校验失败(先切换行，在切换列)，
    '如果切换列和当前列编辑类型不同，不知怎么回事当前列的编辑方式会变为目标列的，调整切换列方式为:mobjVsf.Select lngRow, lngCol即可正常
    If blnAutoFeeLine = True Then
        '换行之前，先检查是否允许换行，即是否有必输的项目没有输入
        If mobjVsf.Row = mobjVsf.Rows - 1 Then
            If mblnAllowNewRow Then
                blnCancel = False
                lngCol = 1
                
                If blnAutoSetLine = False Then
                    RaiseEvent BeforeNewRow(mobjVsf.Row, lngCol, blnCancel)
                End If
                
                If blnCancel Then
                    mobjVsf.Col = lngCol
                    mobjVsf.ShowCell mobjVsf.Row, mobjVsf.Col
                    Exit Function
                End If
                
                If InsertNewRow = False Then Exit Function
                RaiseEvent AfterNewRow(mobjVsf.Row, mobjVsf.Col)
            End If
            lngRow = mobjVsf.Row
        Else
            lngRow = mobjVsf.Row + 1
        End If
        
        '找第一个可以编辑的列
        lngCol = GetAllowCol(lngRow, 1)
        If blnAutoSetLine = True Then
            If lngCol > lngCurCol Then lngCol = lngCurCol
        End If
        If lngCol >= mobjVsf.Cols Then lngCol = 1
        
        'mobjVsf.Col = lngCol
        mobjVsf.Select lngRow, lngCol
    Else
        '找下一个可以编辑的列
        lngCol = GetAllowCol(mobjVsf.Row, mobjVsf.Col + 1)
        If lngCol >= mobjVsf.Cols Then lngCol = 1
        
        mobjVsf.Col = lngCol
    End If
    
    mobjVsf.ShowCell mobjVsf.Row, mobjVsf.Col
    LocationNextCell = True
End Function

Private Function GetAllowCol(ByVal lngRow As Long, ByVal lngFromCol As Long) As Long
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    
    For lngLoop = lngFromCol To mobjVsf.Cols - 1
        If mColProperty(lngLoop).允许编辑 And mobjVsf.ColHidden(lngLoop) = False And AllowEdit Then
            If mColProperty(lngLoop).光标跳跃 = m强制经过 Then
                Exit For
            ElseIf mColProperty(lngLoop).光标跳跃 = m无值时经过 And IIf(mColProperty(lngLoop).编辑控件 = vbVsfEditCheck, mobjVsf.TextMatrix(lngRow, lngLoop) = 0, mobjVsf.TextMatrix(lngRow, lngLoop) = "") Then
                Exit For
            ElseIf mColProperty(lngLoop).光标跳跃 = m有值时经过 And IIf(mColProperty(lngLoop).编辑控件 = vbVsfEditCheck, UCase(mobjVsf.TextMatrix(lngRow, lngLoop)) = 1, mobjVsf.TextMatrix(lngRow, lngLoop) <> "") Then
                Exit For
            End If
        End If
    Next

    GetAllowCol = lngLoop

End Function

Private Function InsertNewRow() As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '******************************************************************************************************************
    If mobjVsf.Editable <> flexEDNone Then
        mobjVsf.AddItem "", mobjVsf.Rows
        mobjVsf.Row = mobjVsf.Rows - 1
        '无法定位到新加行说明列内容有效检查不通过，如：vsf_ValidateEdit中参数Cancle=True
        If mobjVsf.Row <> mobjVsf.Rows - 1 Then
            mobjVsf.RemoveItem mobjVsf.Rows - 1
            Exit Function
        End If
        If mblnAppendRows Then Call DrawSpaceCellLine
    Else
        mobjVsf.Row = mobjVsf.Rows - 1
        If mobjVsf.Row <> mobjVsf.Rows - 1 Then Exit Function
    End If
    
    InsertNewRow = True
End Function

'######################################################################################################################

Public Function DrawSpaceCellLine() As Boolean
    '******************************************************************************************************************
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '******************************************************************************************************************
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    Dim strName As String
    Dim strName2 As String
    Dim objTmp As Object
    Dim objAryX() As Object
    Dim objAryY() As Object
    Dim lngSumColWidth As Long
    
    If TypeName(mobjVsf) <> "VSFlexGrid" Then Exit Function
    
    On Error GoTo errHand
    
    ReDim objAryX(0)
    ReDim objAryY(0)
    
    If mobjVsf.Rows = 0 Or mobjVsf.Cols = 0 Then Exit Function
    
    
    If mintAutoResizeColumn >= 0 Then
        '调整列宽
        lngSumColWidth = mobjVsf.Cell(flexcpLeft, 0, mobjVsf.Cols - 1) + mobjVsf.ColWidth(mobjVsf.Cols - 1)
        lngSumColWidth = lngSumColWidth - mobjVsf.ColWidth(mintAutoResizeColumn)
        If lngSumColWidth < mobjVsf.Width Then
            mobjVsf.ColWidth(mintAutoResizeColumn) = mobjVsf.Width - lngSumColWidth - 15 * (mobjVsf.Cols - mintAutoResizeColumn)
        End If
    End If

    For lngLoop = mobjVsf.Rows - 1 To 1 Step -1
        If mobjVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = mobjVsf.Cell(flexcpTop, lngLastRow, 0) + mobjVsf.RowHeight(lngLastRow)
    
    On Error Resume Next
    Err = 0
    strName = mobjVsf.Name & mobjVsf.Index
    If Err <> 0 Then
        strName = mobjVsf.Name
    End If
    Err = 0
    On Error GoTo errHand
    
    '1.隐藏所有的线
    '------------------------------------------------------------------------------------------------------------------
    For Each objTmp In mobjControls
        
        If TypeName(objTmp) = "Line" And Left(objTmp.Name, Len("ln" & strName & "_")) = "ln" & strName & "_" Then

            objTmp.Visible = False
            lngIndex = Val(Mid(objTmp.Name, Len("ln" & strName & "_") + 2))
            
            Select Case Mid(objTmp.Name, Len("ln" & strName & "_") + 1, 1)
            Case "X"
                
                If UBound(objAryX) < lngIndex Then ReDim Preserve objAryX(lngIndex)
                Set objAryX(lngIndex) = objTmp
                
            Case "Y"
                
                If UBound(objAryY) < lngIndex Then ReDim Preserve objAryY(lngIndex)
                Set objAryY(lngIndex) = objTmp
                
            End Select
        End If
    Next

    '2.重新计算需要的纵线
    '------------------------------------------------------------------------------------------------------------------
    For lngLoop = 1 To mobjVsf.Cols - 1
        If mobjVsf.ColHidden(lngLoop) = False Then
            If UBound(objAryY) < lngLoop Then ReDim Preserve objAryY(lngLoop)
            
            If objAryY(lngLoop) Is Nothing Then
                Set objTmp = mobjControls.Add("VB.Line", "ln" & strName & "_Y" & lngLoop)
                Set objTmp.Container = mobjVsf
                Set objAryY(lngLoop) = objTmp
            End If
    
            With objAryY(lngLoop)
                .ZOrder
                .X1 = mobjVsf.Cell(flexcpLeft, 0, lngLoop) - 15
                .X2 = .X1
                .Y1 = lngTop
                .Y2 = mobjVsf.Height
                .BorderColor = mobjVsf.GridColor
                .Visible = True
            End With
        End If
    Next
    
    '3.计算最后一行的右边纵线
    '------------------------------------------------------------------------------------------------------------------
    If mobjVsf.ColIsVisible(mobjVsf.Cols - 1) Then
        
        lngLoop = mobjVsf.Cols
        
        If UBound(objAryY) < lngLoop Then ReDim Preserve objAryY(lngLoop)
        
        If objAryY(lngLoop) Is Nothing Then
            Set objTmp = mobjControls.Add("VB.Line", "ln" & strName & "_Y" & lngLoop)
            Set objTmp.Container = mobjVsf
            Set objAryY(lngLoop) = objTmp
        End If

        With objAryY(lngLoop)
            .ZOrder
            Select Case mobjVsf.Appearance
            Case flex3D
                .X1 = mobjVsf.Width - 75
            Case flexXPThemes, flex3DLight
                .X1 = mobjVsf.Width - 45
            Case Else
                .X1 = mobjVsf.Width - 15
            End Select
            
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = mobjVsf.Height
            .BorderColor = mobjVsf.GridColor
            .Visible = True
        End With
        
    End If
    
    
   '4.重新计算需要的横线
   '------------------------------------------------------------------------------------------------------------------
    lngIndex = 0
    Do While (lngTop + mobjVsf.RowHeightMin) < mobjVsf.Height

        lngIndex = lngIndex + 1
        
        If UBound(objAryX) < lngIndex Then ReDim Preserve objAryX(lngIndex)
        
        If objAryX(lngIndex) Is Nothing Then
            Set objTmp = mobjControls.Add("VB.Line", "ln" & strName & "_X" & lngIndex)
            Set objTmp.Container = mobjVsf
            Set objAryX(lngIndex) = objTmp
        End If
        
        With objAryX(lngIndex)
            .ZOrder
            .X1 = 0
            .X2 = mobjVsf.Width
            .Y1 = lngTop + mobjVsf.RowHeightMin + IIf(lngIndex = 1, 30, 0)
            .Y2 = .Y1
            .BorderColor = mobjVsf.GridColor
            .Visible = True
            lngTop = .Y1
        End With
    Loop
    
    DrawSpaceCellLine = True
            
    Exit Function
    
errHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowIndicatorIcon() As Boolean
    
    If micoIndicator Is Nothing Then Exit Function
    
    If mintIndicator >= 0 And mintIndicator < mobjVsf.Rows Then
        Set mobjVsf.Cell(flexcpPicture, 1, mintIndicator, mobjVsf.Rows - 1, mintIndicator) = Nothing
        Set mobjVsf.Cell(flexcpPicture, mobjVsf.Row, mintIndicator, mobjVsf.Row, mintIndicator) = micoIndicator
    End If
    
End Function

Private Sub Class_initialize()
    mintKeyCol = -1
    mintAutoResizeColumn = -1
    mintIndicator = -1
End Sub

Private Sub Class_Terminate()
    Set mobjVsf = Nothing
End Sub
Public Function vsfCheckKeyPress(KeyAscii As Integer, TextType As mTextType, Optional strEditText As String = "")
     '功能：控制Vsf控件中的输入值
     '入参：
     '返回：
     '调用：Vsf控件的KeyPressEdit事件
     '编写：程池富
     '日期：2009-02-26
    Select Case TextType
        Case 整数型
            If KeyAscii <> vbKeyReturn Then
                If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0: Exit Function
            End If
        Case 小数型
            If KeyAscii <> vbKeyReturn Then
                If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0: Exit Function
                If InStr(strEditText, ".") > 0 And KeyAscii = 46 Then KeyAscii = 0: Exit Function '小数点只准输入一次
            End If
   End Select
End Function

