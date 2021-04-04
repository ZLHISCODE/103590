VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVSFlexGridEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------
'用途：标准VSFlexGrid的扩展封装
'说明：
'1.方法执行顺序：
'      Recordset->Bunding->Init->Head->ColsReadonly->Editable->Repaint
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'-----------------------------------------

Private Const GSTR_MSG As String = "clsVSFlexGridEx类"

'绑定的VSFlexGrid控件
Private WithEvents mvsfBunding As VSFlexGrid
Attribute mvsfBunding.VB_VarHelpID = -1

'填充数据事件
Public Event EventFillData(ByVal Row As Long, ByVal Col As Long)

Public Enum enuEditable
    EM_View = 0
    EM_Modify = 1
    EM_Verify = 2
    EM_Display = 3
End Enum

Public Enum enuRepaintType
    RT_Columns = 0
    RT_Rows = 1
    RT_ColsAndRows = 2
End Enum

'绑定的记录集（数据展现）
Private mrsRecordset As ADODB.Recordset

'列头
Private mstrHead As String

'只读列字符串；例：列1[|列2[|...]]
Private mstrColsReadonly As String

'控件模式
Private mbytEditable As enuEditable

'Append方法要复制上一行列的值
Private mstrAppendFields As String

'隐藏列字符串；例：列1[|列2[|...]]
Private mstrColsHide As String

Public Property Get ColsHide() As String
    ColsHide = mstrColsHide
End Property

Public Property Let ColsHide(ByVal strValue As String)
    mstrColsHide = strValue
End Property

Public Property Get AppendFields() As String
    AppendFields = mstrAppendFields
End Property

Public Property Let AppendFields(ByVal strValue As String)
    mstrAppendFields = strValue
End Property

Public Property Get Editable() As enuEditable
    Editable = mbytEditable
End Property

Public Property Let Editable(ByVal BytValue As enuEditable)
    mbytEditable = BytValue
    If BytValue = EM_Modify Or BytValue = EM_Verify Then
        mvsfBunding.Editable = flexEDKbdMouse
    Else
        mvsfBunding.Editable = flexEDNone
    End If
End Property

Public Property Get ColsReadonly() As String
    ColsReadonly = mstrColsReadonly
End Property

Public Property Let ColsReadonly(ByVal strValue As String)
    mstrColsReadonly = strValue
End Property

Public Property Get Head() As String
    Head = mstrHead
End Property

Public Property Let Head(ByVal strValue As String)
    mstrHead = strValue
    '立即设置列
    Call SetVSFHead(mvsfBunding, mstrHead)
End Property

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mrsRecordset
End Property

Public Property Let Recordset(ByVal RecValue As ADODB.Recordset)
    Set mrsRecordset = RecValue
End Property

Public Property Get Bunding() As Object
    Set Bunding = mvsfBunding
End Property

Public Property Let Bunding(ByVal VSFValue As Object)
    On Error GoTo hErr
    Set mvsfBunding = VSFValue
    Exit Property
    
hErr:
    MsgBox "绑定的控件必须是VSFlexGrid！", vbInformation, App.ProductName
End Property

Public Sub Repaint(Optional ByVal typVar As enuRepaintType)
'功能：重绘VSFlexGrid控件
'参数：
'  typVar：重绘Grid

    If typVar = RT_Columns Or typVar = RT_ColsAndRows Then
        Call SetVSFHead(mvsfBunding, mstrHead)
        Call SetColsHide
    End If
    If typVar = RT_Rows Or typVar = RT_ColsAndRows Then
        Call FillData
    End If
End Sub

Private Sub FillData()
    Dim intCol As Integer, intField As Integer
    Dim lngRow As Long, lngOldRow As Long
    
    With mvsfBunding
        lngOldRow = .Row
        .Redraw = flexRDNone
        .Clear 1
        .Rows = 1
                
        If mrsRecordset.State <> adStateOpen Then
            .Redraw = True
            'MsgBox "记录集未打开！", vbInformation, App.ProductName
            Exit Sub
        End If
        
        If mrsRecordset.RecordCount > 0 Then mrsRecordset.MoveFirst
        Do While mrsRecordset.EOF = False
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            For intField = 0 To mrsRecordset.Fields.Count - 1
                intCol = .ColIndex(mrsRecordset.Fields(intField).Name)
                If intCol >= 0 Then
                    .TextMatrix(lngRow, intCol) = IIf(IsNull(mrsRecordset.Fields(intField).Value), "", mrsRecordset.Fields(intField).Value)
                    RaiseEvent EventFillData(lngRow, intCol)
                End If
            Next
            mrsRecordset.MoveNext
        Loop
        
        .Redraw = flexRDDirect
        
        '恢复行位置
        If .Rows > 1 Then
            If .Rows - 1 >= lngOldRow And lngOldRow > 0 Then
                .Row = lngOldRow
            Else
                .Row = 1
            End If
        End If
        
        '只读列的背景色
        Call SetColsBackColor
        
    End With
End Sub

Private Sub Class_Terminate()
    Set mrsRecordset = Nothing
    Set mvsfBunding = Nothing
End Sub

Private Sub SetVSFHead(ByVal vsfObject As Object, ByVal strHead As String)
'--------------------------------
'功能：初始化VSFlexGrid控件表格头
'参数：
'  vsfObject：目标控件；
'  strHead：表格头的初始化字串
'
'格式： "剂型,,3,1000,s|..."
'   元素1：Key值；
'   元素2：Caption值（默认为Key值）；
'   元素3：列属性（0：内部显示，可移动；1：内部隐藏，不可移动，不可显示；2：用户隐藏；3：用户显示(默认值)）
'   元素4：列宽度（默认0）；
'   元素5：显示格式；s(默认)：字符串； n：数字； d：日期； t：时间； dt：日期时间
'--------------------------------
    Dim arrCols As Variant, arrRows As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    arrRows = Split(strHead, "|")
    With vsfObject
        If .Rows = 0 Then .Rows = 1
        .Cols = UBound(arrRows) + 1
        For i = LBound(arrRows) To UBound(arrRows)
            If arrRows(i) <> "" Then
                arrCols = Split(arrRows(i), ",")
                '第1元素：Key值
                .ColKey(i) = arrCols(0)
                
                '第2元素：Caption值
                If arrCols(1) = "" Then
                    .TextMatrix(0, i) = arrCols(0)
                Else
                    .TextMatrix(0, i) = arrCols(1)
                End If
                
                '第3元素：列属性
                If arrCols(2) = "" Then
                    .colData(i) = 3
                Else
                    .colData(i) = Val(arrCols(2))
                End If
                
                '第4元素：宽度
                .ColWidth(i) = Val(arrCols(3))
                
                '第5元素：显示格式
                If UBound(arrCols) > 3 Then
                    If UCase(arrCols(4)) = "D" Then
                        .ColFormat(i) = "yyyy-mm-dd"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "T" Then
                        .ColFormat(i) = "hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "DT" Then
                        .ColFormat(i) = "yyyy-mm-dd hh:mm:ss"
                        .ColAlignment(i) = flexAlignCenterCenter
                    ElseIf UCase(arrCols(4)) = "N" Then
                        .ColAlignment(i) = flexAlignRightCenter
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                        If UCase(arrCols(4)) = "B" Then
                            .ColDataType(i) = flexDTBoolean
                        End If
                    End If
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
                
                '隐藏列
                If Val(arrCols(2)) = 1 Or Val(arrCols(2)) = 2 Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                End If
                
            End If
        Next
        
        If .Cols > 0 Then .Cell(2, 0, 0, 0, .Cols - 1) = 4
    End With
    Exit Sub
    
errHandle:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub

Private Sub mvsfBunding_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '只读列
    Cancel = InStr("|" & mstrColsReadonly & "|", "|" & mvsfBunding.ColKey(Col) & "|") > 0
End Sub

Private Sub mvsfBunding_GotFocus()
    '得到焦点当前行的背景色
    mvsfBunding.BackColorSel = &H8000000D
End Sub

Private Sub mvsfBunding_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then
            gobjComLib.zlCommFun.PressKey vbKeyLeft
        ElseIf Shift = 0 Then
            gobjComLib.zlCommFun.PressKey vbKeyRight
            
            With mvsfBunding
                .Redraw = False
                If .Col >= .RightCol Then
                    '最后一列时，列焦点定位至下一行
                    If mbytEditable = EM_Modify And .Row >= .Rows - 1 Then
                        '根据“AppendFields”属性添加新行
                        Call Append(mstrAppendFields)
                        .Row = .Rows - 1
                    Else
                        .Row = .Row + 1
                    End If
                    .ColSel = .LeftCol
                    
                    '重绘列的背景色（可能有行增加）
                    Call SetColsBackColor
                End If
                .Redraw = True
            End With
        End If
    ElseIf KeyCode = vbKeyDelete And mbytEditable = EM_Modify Then
        If mvsfBunding.Row <= 0 Then Exit Sub
        
        If MsgBox("确认删除该行吗？", vbQuestion + vbYesNo + vbDefaultButton2, GSTR_MSG) = vbYes Then
            If mvsfBunding.Rows > 2 Then
                mvsfBunding.RemoveItem mvsfBunding.Row
            Else
                mvsfBunding.Clear 1
            End If
            mvsfBunding.SetFocus
        End If
    End If
End Sub

Private Sub mvsfBunding_LostFocus()
    '失去焦点当前行的背景色
    mvsfBunding.BackColorSel = &H8000000C
End Sub

Public Sub Init(Optional ByVal blnPermitSort As Boolean = False)
'功能：初始化控件的基本属性
'参数：
'  blnPermitSort：True允许点击列头排序；False不允许

    If mvsfBunding Is Nothing Then Exit Sub
    
    With mvsfBunding
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .ExplorerBar = IIf(blnPermitSort, flexExSortAndMove, flexExMove)
        .AutoResize = True
        .SheetBorder = &H8000000F
        .BackColorBkg = &H80000005
        .RowHeightMin = 270
        .FocusRect = flexFocusHeavy
        .Rows = 1
    End With
End Sub

Public Sub Append(Optional ByVal strCopyCols As String)
'功能：为VSFlexGrid添加一行

    Dim i As Integer, intCol As Integer
    Dim cllSource As Collection
    Dim arrCopyCols As Variant
    
    arrCopyCols = Split(arrCopyCols, "|")
    
    With mvsfBunding
        If .Rows > 1 Then
            If IsEmpty(arrCopyCols) Then
                .Rows = .Rows + 1
                Exit Sub
            End If
            If UBound(arrCopyCols) >= 0 Then
                '先将最后一行的值放到集合保存
                Set cllSource = New Collection
                For i = LBound(arrCopyCols) To UBound(arrCopyCols)
                    intCol = .ColIndex(arrCopyCols(i))
                    If intCol >= 0 Then
                        cllSource.Add .TextMatrix(.Rows - 1, intCol), arrCopyCols(i)  '"_" & intCol
                    End If
                Next
                '添加一行后再回写新行
                .Rows = .Rows + 1
                For i = 1 To cllSource.Count
                    .TextMatrix(.Rows - 1, .ColIndex(arrCopyCols(i - 1))) = cllSource(i)
                Next
            Else
                .Rows = .Rows + 1
            End If
        Else
            .Rows = .Rows + 1
        End If
    End With
    
End Sub

Private Sub SetColsBackColor()
'功能：设置不可编辑列的背景色
    
    Dim intCol As Integer

    With mvsfBunding
        If .Rows <= 1 Then Exit Sub
    
        If mbytEditable = EM_Display Then
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = .BackColor
        ElseIf mbytEditable = EM_View Then
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = &H8000000F
        Else
            For intCol = 0 To .Cols - 1
                If InStr("|" & mstrColsReadonly & "|", "|" & .ColKey(intCol) & "|") > 0 Then
                    .Cell(flexcpBackColor, 1, intCol, .Rows - 1, intCol) = &H8000000F
                End If
            Next
        End If
    End With
End Sub

Public Sub SetColsHide()
'功能：隐藏列
'参数：
'  strCols：隐藏列字符串；例：列1[|列2[|...]]

    Dim arrCols As Variant
    Dim strCol As String
    Dim i As Integer, j As Integer
    Dim blnFind As Boolean
    
    If mvsfBunding Is Nothing Then Exit Sub
    
    With mvsfBunding
        For i = 0 To .Cols - 1
            blnFind = False
            arrCols = Split(mstrColsHide & "|", "|")
            For j = LBound(arrCols) To UBound(arrCols)
                strCol = UCase(Trim(arrCols(j)))
                If strCol = UCase(.ColKey(i)) Then
                    blnFind = True
                    Exit For
                End If
            Next
            .ColHidden(i) = blnFind
        Next
        Erase arrCols
    End With
End Sub
