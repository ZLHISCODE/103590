VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl VsfGrid 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.ImageList ils16 
      Left            =   1590
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VsfGrid.ctx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2085
      Left            =   285
      TabIndex        =   0
      Top             =   405
      Width           =   2895
      _cx             =   5106
      _cy             =   3678
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483634
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
End
Attribute VB_Name = "VsfGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Column
    ComboList As String
    EditMode As Byte
    MaxLength As Integer
End Type

Private mColumn() As Column
Private mblnLoading As Boolean
Private mblnNoDouble As Boolean
Private mblnEditIng As Boolean
Private mblnAppendRows As Boolean
Private mblnSelRow As Boolean
Private mblnComboEdit As Boolean
Private mblnSelEdit As Boolean

Public Event ChangeEdit(ByVal Row As Long, ByVal Col As Long)
Public Event StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
Public Event ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
Public Event KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
Public Event KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
Public Event KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

Public Event BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
Public Event BeforeDeleteCell(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
Public Event CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Public Event BeforeComboList(ByVal OldCol As Long, ByVal NewCol As Long, ComboList As String, Cancel As Boolean)
Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long)
Public Event AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Public Event BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
Public Event AfterUserResize(ByVal Row As Long, ByVal Col As Long)
Public Event AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)

Public Property Let SelectMode(ByVal vData As Boolean)
    mblnSelRow = vData
End Property

Public Property Get SelectMode() As Boolean
    SelectMode = mblnSelRow
End Property

Public Property Let ComboEdit(ByVal vData As Boolean)
    mblnComboEdit = vData
End Property

Public Property Get ComboEdit() As Boolean
    ComboEdit = mblnComboEdit
End Property

Public Property Let SelEdit(ByVal vData As Boolean)
    mblnSelEdit = vData
End Property

Public Property Get SelEdit() As Boolean
    SelEdit = mblnSelEdit
End Property

Private Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeightMin) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeightMin + IIf(lngIndex = 1, 30, 0)
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    vsf.Cell(flexcpBackColor, 1, 0, vsf.Rows - 1, 0) = vsf.BackColor
    
    Exit Function
    
errHand:
    
End Function

Public Sub RemoveItem(ByVal Row As Long)
    vsf.RemoveItem Row
End Sub

Public Property Let ExtendLastCol(ByVal vData As Boolean)
    vsf.ExtendLastCol = vData
End Property

Public Property Let AppendRow(ByVal vData As Boolean)
    
    mblnAppendRows = vData
    
    If mblnAppendRows Then Call AppendRows(vsf, lnX, lnY)
        
End Property

Public Property Let ColHidden(ByVal Col As Long, ByVal vData As Boolean)
    vsf.ColHidden(Col) = vData
End Property

Public Property Get ColHidden(ByVal Col As Long) As Boolean
    ColHidden = vsf.ColHidden(Col)
End Property

Public Property Let EditMode(ByVal Col As Long, ByVal vData As Byte)
    mColumn(Col).EditMode = vData
End Property

Public Property Get EditMode(ByVal Col As Long) As Byte
    EditMode = mColumn(Col).EditMode
End Property

Public Property Let MaxLength(ByVal Col As Long, ByVal vData As Integer)
    mColumn(Col).MaxLength = vData
End Property

Public Property Get MaxLength(ByVal Col As Long) As Integer
    MaxLength = mColumn(Col).MaxLength
End Property

Public Property Let ComboList(ByVal Col As Long, vData As String)
    mColumn(Col).ComboList = vData
End Property

Public Property Let VsfComboList(vData As String)
    vsf.ComboList = vData
End Property

Public Property Get IsEditing() As Boolean
    IsEditing = mblnEditIng
End Property

Public Property Get ComboList(ByVal Col As Long) As String
    ComboList = mColumn(Col).ComboList
End Property

Public Property Let ColDataType(ByVal Col As Long, ByVal vData As DataTypeSettings)
    vsf.ColDataType(Col) = vData
End Property

Public Property Let Cols(vData As Long)
    vsf.Cols = vData
End Property

Public Property Get Cols() As Long
    Cols = vsf.Cols
End Property

Public Property Let NoDouble(vData As Boolean)
    mblnNoDouble = vData
End Property

Public Property Let FixedCols(vData As Long)
    vsf.FixedCols = vData
End Property

Public Property Get FixedCols() As Long
    FixedCols = vsf.FixedCols
End Property

Public Property Let FixedRows(vData As Long)
    vsf.FixedRows = vData
End Property

Public Property Get FixedRows() As Long
    FixedRows = vsf.FixedRows
End Property

Public Property Let Rows(vData As Long)
    vsf.Rows = vData
End Property

Public Property Get Rows() As Long
    Rows = vsf.Rows
End Property

Public Property Let Col(vData As Long)
    vsf.Col = vData
End Property

Public Property Get Col() As Long
    Col = vsf.Col
End Property

Public Property Let Row(vData As Long)
    vsf.Row = vData
End Property

Public Property Get Row() As Long
    Row = vsf.Row
End Property

Public Property Get hWnd() As Long
    hWnd = vsf.hWnd
End Property

Public Property Get CellLeft() As Long
    CellLeft = vsf.CellLeft
End Property

Public Property Get CellTop() As Long
    CellTop = vsf.CellTop
End Property

Public Property Get CellHeight() As Long
    CellHeight = vsf.CellHeight
End Property

Public Property Get CellWidth() As Long
    CellWidth = vsf.CellWidth
End Property

Public Property Let EditText(ByVal vData As String)
    vsf.EditText = vData
End Property

Public Property Get EditText() As String
    EditText = vsf.EditText
End Property

Public Property Let RowData(ByVal Row As Long, ByVal vData As Variant)
    vsf.RowData(Row) = vData
End Property

Public Property Get RowData(ByVal Row As Long) As Variant
    RowData = vsf.RowData(Row)
End Property

Public Property Let ColData(ByVal Col As Long, ByVal vData As Variant)
    vsf.ColData(Col) = vData
End Property

Public Property Get ColData(ByVal Col As Long) As Variant
    ColData = vsf.ColData(Col)
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal vData As String)
    vsf.TextMatrix(Row, Col) = vData
End Property

Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
    TextMatrix = vsf.TextMatrix(Row, Col)
End Property

Public Property Let Cell(ByVal Setting As CellPropertySettings, Optional ByVal Row1 As Long, Optional ByVal Col1 As Long, Optional ByVal Row2 As Long, Optional ByVal Col2 As Long, ByVal vData As Variant)
    vsf.Cell(Setting, Row1, Col1, Row2, Col2) = vData
End Property

Public Property Get Cell(ByVal Setting As CellPropertySettings, Optional ByVal Row1 As Long, Optional ByVal Col1 As Long, Optional ByVal Row2 As Long, Optional ByVal Col2 As Long) As Variant
    Cell = vsf.Cell(Setting, Row1, Col1, Row2, Col2)
End Property

Public Property Get Body() As VSFlexGrid
    Set Body = vsf
End Property

Public Sub ShowCell(ByVal Row As Long, ByVal Col As Long)
    vsf.ShowCell Row, Col
    
End Sub






Public Sub NewColumn(ByVal vText As String, _
                    Optional ByVal vWidth As Single = 1200, _
                    Optional ByVal vAlignment As Byte = 9, _
                    Optional ByVal ComboList As String = "", _
                    Optional ByVal EditMode As Byte = 0, _
                    Optional ByVal MaxLength As Integer = 0, _
                    Optional ByVal DataType As DataTypeSettings = flexDTString)
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim i As Long
    
    mblnLoading = True
    
    vsf.Cols = vsf.Cols + 1
    i = vsf.Cols - 1
    
    vsf.TextMatrix(0, i) = vText
    vsf.ColWidth(i) = vWidth
    vsf.ColAlignment(i) = vAlignment
        
    ReDim Preserve mColumn(0 To i)
    
    mColumn(i).ComboList = ComboList
    mColumn(i).EditMode = EditMode
    mColumn(i).MaxLength = MaxLength
    
    vsf.ColDataType(i) = DataType
    
    On Error Resume Next
    
    vsf.ColAlignmentFixed(i) = vAlignment
    
    mblnLoading = False
    
End Sub


Private Sub UserControl_Resize()
    On Error Resume Next
    
    vsf.Left = 0
    vsf.Top = 0
    vsf.Width = UserControl.Width
    vsf.Height = UserControl.Height
    
    If mblnAppendRows Then Call AppendRows(vsf, lnX, lnY)
     
End Sub

Private Sub UserControl_Show()
    If vsf.Rows > 1 And SelectMode Then
        Call SelectRow(vsf, -1, vsf.FixedRows, 16772055)
    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsf.ComboList = "" Then vsf.ComboList = mColumn(NewCol).ComboList
    
    If mColumn(Col).ComboList = "..." And mblnSelEdit = False Then
        vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
    End If
    
    RaiseEvent AfterEdit(Row, Col)
    
    mblnEditIng = False
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    Dim ComboList As String
    Dim blnCancel As Boolean
    
    If mblnSelRow Then
        If vsf.Cell(flexcpBackColor, NewRow, NewCol) <> 16772055 Then
            Call SelectRow(vsf, OldRow, NewRow, 16772055)
        End If
    End If
    
    If mblnLoading Then Exit Sub
    
    RaiseEvent AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    
    If mColumn(NewCol).EditMode = 1 Then
        vsf.FocusRect = flexFocusSolid
    Else
        vsf.FocusRect = flexFocusHeavy
    End If
    
    Call AdjustRowFlag(vsf, NewRow)
'    If OldCol <> NewCol Then

        vsf.ComboList = mColumn(NewCol).ComboList
        
        
        If mColumn(NewCol).ComboList = " " Then
            
            '下拉,传入记录集
            blnCancel = False
            RaiseEvent BeforeComboList(OldRow, NewCol, ComboList, blnCancel)
            If blnCancel = False Then vsf.ComboList = ComboList
                        
        End If
        
 '   End If
        
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If mblnAppendRows Then Call AppendRows(vsf, lnX, lnY)
    RaiseEvent AfterScroll(OldTopRow, OldLeftCol, NewTopRow, NewLeftCol)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If mblnAppendRows Then Call AppendRows(vsf, lnX, lnY)
    RaiseEvent AfterUserResize(Row, Col)
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    RaiseEvent BeforeRowColChange(OldRow, OldCol, NewRow, NewCol, Cancel)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    If mColumn(Col).ComboList = "..." Then
        
        RaiseEvent CellButtonClick(Row, Col)
        
    End If
    
End Sub

Private Sub vsf_ChangeEdit()
    RaiseEvent ChangeEdit(Row, Col)
End Sub

Private Sub vsf_DblClick()
    If mblnNoDouble = False Then Call vsf_KeyPress(32)
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim lngLoop As Long
    Dim blnCancel As Boolean
    
    On Error GoTo errHand
    
    Select Case KeyCode
    Case vbKeyDelete
        
        If Shift = 0 And vsf.Editable <> flexEDNone Then
            '删除整行及内容
            
            blnCancel = False
            
            RaiseEvent BeforeDeleteRow(vsf.Row, vsf.Col, blnCancel)
            
            If blnCancel Then Exit Sub
            
            If vsf.Rows > 1 Then
                If vsf.Rows = 2 And vsf.Row = 1 Then
                    For lngLoop = 0 To vsf.Cols - 1
                        vsf.TextMatrix(1, lngLoop) = ""
                    Next
                    vsf.RowData(1) = ""
                Else
                    vsf.RemoveItem vsf.Row
                    If mblnAppendRows Then Call AppendRows(vsf, lnX, lnY)
                End If
                Call AdjustRowFlag(vsf, vsf.Row)
                
                RaiseEvent AfterDeleteRow(vsf.Row, vsf.Col)
            End If
            
        End If
        
        If Shift = 2 And vsf.Editable <> flexEDNone And mColumn(vsf.Col).EditMode = 1 Then
            '删除当前单元格的内容
            
            blnCancel = False
            RaiseEvent BeforeDeleteCell(vsf.Row, vsf.Col, blnCancel)
            If blnCancel Then Exit Sub
            
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
            
            RaiseEvent AfterDeleteCell(vsf.Row, vsf.Col)
            
        End If
        
    Case Else
        If (mColumn(vsf.Col).ComboList = "..." And Shift = 0) Or mblnComboEdit Then
            vsf.ComboList = ""
        End If
    End Select
    
    Exit Sub
    
errHand:
        
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim blnCancel As Boolean
    
    blnCancel = False
    RaiseEvent KeyDownEdit(Row, Col, mColumn(Col).ComboList, KeyCode, Shift, blnCancel)
    
    If blnCancel Then
        KeyCode = 0
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then vsf.ComboList = mColumn(Col).ComboList
        
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    RaiseEvent KeyPress(vsf.Row, vsf.Col, KeyAscii, blnCancel)
    
    If blnCancel Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call GoNextCell
    Else
        If vsf.ComboList = "..." Or mblnComboEdit Then
            vsf.ComboList = ""
        End If
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call GoNextCell
    End If
        
    RaiseEvent KeyPressEdit(Row, Col, KeyAscii)
    
End Sub

Private Sub vsf_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    RaiseEvent KeyUpEdit(Row, Col, KeyCode, Shift)
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsf.EditSelStart = 0
    vsf.EditSelLength = zlCommFun.ActualLen(vsf.EditText)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
                        
    RaiseEvent StartEdit(Row, Col, Cancel)
    If Cancel = True Then Exit Sub
    
    'Cancel = mblnNoDouble
    'If Cancel = True Then Exit Sub
    
    '先保存原来的内容，在弹出选择取消或没有数据时恢复
    vsf.Cell(flexcpData, Row, Col) = vsf.TextMatrix(Row, Col)
    '检查是否允许编辑，并作相应的处理
    Cancel = (mColumn(Col).EditMode = 0)
    If Cancel Then Exit Sub
    '设置可编辑时的长度
    vsf.EditMaxLength = mColumn(Col).MaxLength
    
    mblnEditIng = True
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Not StrIsValid(vsf.EditText, mColumn(Col).MaxLength)
    
    If Cancel Then Exit Sub
    
    RaiseEvent ValidateEdit(Row, Col, Cancel)
    
End Sub

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    On Error Resume Next
    
    If Not (objVsf.Cell(flexcpPicture, intRow, 0) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, 0, objVsf.Rows - 1, 0) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, 0) = ils16.ListImages(1).Picture
    
End Sub

Private Sub GoNextCell()
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '-----------------------------------------------------------------------------------------
    Dim lngCol As Long
    Dim blnCancel As Boolean
    
    If GetAllowCol(vsf.Col + 1) > vsf.Cols - 1 Then
        '换行之前，先检查是否允许换行，即是否有必输的项目没有输入
                
        If vsf.Row = vsf.Rows - 1 Then
            blnCancel = False
            
            lngCol = 1
            
            RaiseEvent BeforeNewRow(vsf.Row, lngCol, blnCancel)
            
            If blnCancel Then
                vsf.Col = lngCol
                vsf.ShowCell vsf.Row, vsf.Col
                Exit Sub
            End If
            
            Call InsertNewRow
        Else
            vsf.Row = vsf.Row + 1
        End If
        
        '找第一个可以编辑的列
        lngCol = GetAllowCol(1)
        If lngCol >= vsf.Cols Then lngCol = 1
        
        vsf.Col = lngCol
        
    Else
        '找下一个可以编辑的列
        lngCol = GetAllowCol(vsf.Col + 1)
        If lngCol >= vsf.Cols Then lngCol = 1
        
        vsf.Col = lngCol
    End If
    
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Private Sub InsertNewRow()
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    If vsf.Editable <> flexEDNone Then
        vsf.AddItem "", vsf.Rows
        vsf.Row = vsf.Rows - 1
        
        If mblnAppendRows Then Call AppendRows(vsf, lnX, lnY)
    Else
        vsf.Row = vsf.Rows - 1
    End If
End Sub

Private Function GetAllowCol(ByVal lngFromCol As Long) As Long
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '-----------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lngLoop As Long
    
    lngRow = vsf.Row
    
    For lngLoop = lngFromCol To vsf.Cols - 1
        If mColumn(lngLoop).EditMode = 1 Then Exit For
    Next

    GetAllowCol = lngLoop

End Function



Public Sub SelectRow(objVsf As Object, ByVal OldRow As Long, ByVal NewRow As Long, Optional ByVal lngBackColor As Long = -1)
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    On Error Resume Next
    
    If lngBackColor = -1 Then
        lngColor = objVsf.BackColorSel
    Else
        lngColor = lngBackColor
    End If
    
    If OldRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, OldRow, objVsf.FixedCols, OldRow, objVsf.Cols - 1) = objVsf.BackColor
    End If
    
    If NewRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, NewRow, objVsf.FixedCols, NewRow, objVsf.Cols - 1) = lngColor
    End If
    
End Sub
