VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPatiCureCardInvoice 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
      Height          =   1845
      Left            =   1920
      TabIndex        =   0
      Top             =   1470
      Width           =   1800
      _cx             =   3175
      _cy             =   3254
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
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
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmPatiCureCardInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnNOMoved As Boolean

Public Sub zlReLoadData(ByVal strNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取票据使用明细
    '入参:strNO-病人信息集
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-04-01 18:55:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    On Error GoTo errH
    If strNo = "" Then
        vsInvoice.Redraw = flexRDNone
        vsInvoice.Rows = 2
        vsInvoice.Clear 1
        vsInvoice.Redraw = flexRDBuffered: Exit Sub
    End If
    vsInvoice.Rows = 2
    vsInvoice.Clear 1
    
    If Not zl_ExseSvr_GetUseBillInfo(strNo, rsInvoice, True) Then Exit Sub
    
    If rsInvoice.RecordCount <> 0 Then rsInvoice.MoveFirst
    With vsInvoice
        .Rows = IIf(rsInvoice.RecordCount = 0, 1, rsInvoice.RecordCount) + 1
        i = 1
        Do While Not rsInvoice.EOF
            .TextMatrix(i, .ColIndex("ID")) = Nvl(rsInvoice!id)
            .TextMatrix(i, .ColIndex("票据号")) = Nvl(rsInvoice!票据号)
            .TextMatrix(i, .ColIndex("使用原因")) = Nvl(rsInvoice!使用原因)
            .TextMatrix(i, .ColIndex("使用时间")) = Nvl(rsInvoice!使用时间)
            .TextMatrix(i, .ColIndex("使用人")) = Nvl(rsInvoice!使用人)
            i = i + 1
            rsInvoice.MoveNext
        Loop
    End With
    vsInvoice.Redraw = flexRDBuffered
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call InitInvoiceGrid
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsInvoice
        .Top = Me.ScaleTop
        .Left = Me.ScaleLeft
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zl_vsGrid_Para_Save 1107, vsInvoice, Me.Name, "发票信息列表", False
End Sub

Private Sub vsInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsInvoice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsInvoice_GotFocus()
    zl_VsGridGotFocus vsInvoice, &HFFC0C0
End Sub

Private Sub vsInvoice_LostFocus()
    zl_VsGridLostFocus vsInvoice, , vsInvoice.Cell(flexcpForeColor, vsInvoice.Row, vsInvoice.Col)
End Sub
Private Sub InitInvoiceGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化发票网格控件
    '编制:刘兴洪
    '日期:2020-03-25 17:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsInvoice
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        .Clear 1: .Rows = 2
        .Cols = 5
        .TextMatrix(0, i) = "ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "票据号": .ColWidth(i) = 1000: i = i + 1
        .TextMatrix(0, i) = "使用原因": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "使用时间": .ColWidth(i) = 1200: i = i + 1
        .TextMatrix(0, i) = "使用人": .ColWidth(i) = 1000: i = i + 1
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter: .ColAlignment(i) = flexAlignLeftCenter
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = 1000
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Select Case .ColKey(i)
            Case "ID"
                .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Case "票据号"
                .ColAlignment(i) = flexAlignCenterCenter
            End Select
        Next
        
         .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .RowHeightMin = 350
        zl_vsGrid_Para_Restore 1107, vsInvoice, Me.Name, "发票信息列表", False
        If .Rows < 2 Then .Rows = 2
        .Redraw = flexRDBuffered
    End With
End Sub

