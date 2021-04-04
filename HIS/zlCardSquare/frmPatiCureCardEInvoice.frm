VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPatiCureCardEInvoice 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsEInvoice 
      Height          =   1845
      Left            =   615
      TabIndex        =   0
      Top             =   495
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
Attribute VB_Name = "frmPatiCureCardEInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mobjEInvoice As clsEInvoiceObj  '电子票据部件
Private mlng原结帐ID As Long
Private mlng电子票据ID As Long, mbln是否换开 As Boolean
Private mbln是否电子票据 As Boolean

Public Sub zlInitVar(ByRef objEInvoice As clsEInvoiceObj, ByVal strPrivs As String, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:objEinvoice-电子发票部件
    '     strPrivs-当前权限串
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-25 16:59:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    Set mobjEInvoice = objEInvoice: mstrPrivs = strPrivs: mlngModule = lngModule
End Sub

 
Public Property Get 是否电子票据() As Boolean
    是否电子票据 = mbln是否电子票据
End Property
Public Property Get 原结帐ID() As Long
    原结帐ID = mlng原结帐ID
End Property
Public Property Get 电子票据ID() As Long
    电子票据ID = mlng电子票据ID
End Property
Public Property Get 是否换开() As Boolean
    是否换开 = mbln是否换开
End Property
Public Sub zlReLoadData(ByVal strNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载电子发票信息
    '编制:刘兴洪
    '日期:2020-03-25 17:13:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSQL As String
    Dim rsEInvoice As ADODB.Recordset
  
    On Error GoTo errHandle
     
    mlng电子票据ID = 0: mbln是否换开 = False
    vsEInvoice.Clear 1: vsEInvoice.Rows = 2
     
    If strNo = "" Or mobjEInvoice Is Nothing Then Exit Sub
    
    mbln是否电子票据 = mobjEInvoice.zlIsStartEinvoicFromNO(strNo, mlng原结帐ID)
    If mlng原结帐ID = 0 Then Exit Sub

    If Not mobjEInvoice.zlGetEInvoiceInforFromBalanceID(mlng原结帐ID, rsEInvoice, 5, 0) Then Exit Sub
    If rsEInvoice.EOF Then Exit Sub
    
    With vsEInvoice
        If rsEInvoice.RecordCount <> 0 Then rsEInvoice.MoveFirst
        i = 1
        Do While Not rsEInvoice.EOF
            .TextMatrix(i, .ColIndex("ID")) = Nvl(rsEInvoice!id)
            .TextMatrix(i, .ColIndex("记录状态")) = Nvl(rsEInvoice!记录状态)
            .TextMatrix(i, .ColIndex("结算ID")) = Nvl(rsEInvoice!结算ID)
            .TextMatrix(i, .ColIndex("发票代码")) = Nvl(rsEInvoice!代码)
            .TextMatrix(i, .ColIndex("发票号码")) = Nvl(rsEInvoice!号码)
            .TextMatrix(i, .ColIndex("票据金额")) = Format(Nvl(rsEInvoice!票据金额), "###0.00;-###0.00;;")
            .TextMatrix(i, .ColIndex("生成时间")) = Format(rsEInvoice!生成时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("换开纸质发票")) = IIf(Val(Nvl(rsEInvoice!是否换开)) = 1, "已换开", "未换开")
            .TextMatrix(i, .ColIndex("纸质发票号")) = Nvl(rsEInvoice!纸质发票号)
            .TextMatrix(i, .ColIndex("备注")) = Nvl(rsEInvoice!备注)
            .TextMatrix(i, .ColIndex("操作员姓名")) = Nvl(rsEInvoice!操作员姓名)
            If Val(Nvl(rsEInvoice!记录状态)) = 1 Then
                mlng电子票据ID = Nvl(rsEInvoice!id): mbln是否换开 = Val(Nvl(rsEInvoice!是否换开)) = 1
                 .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = Me.ForeColor
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = IIf(Val(Nvl(rsEInvoice!记录状态)) = 2, vbRed, vbBlue)
            End If
            i = i + 1: .Rows = .Rows + 1
            rsEInvoice.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitEinvoiceGrid
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsEInvoice
        .Top = Me.ScaleTop
        .Left = Me.ScaleLeft
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zl_vsGrid_Para_Save 1107, vsEInvoice, Me.Name, "电子票据信息列表", False
    Set mobjEInvoice = Nothing
End Sub

Private Sub vsEInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsEInvoice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsEInvoice_GotFocus()
    zl_VsGridGotFocus vsEInvoice, &HFFC0C0
End Sub

Private Sub vsEInvoice_LostFocus()
    zl_VsGridLostFocus vsEInvoice, , vsEInvoice.Cell(flexcpForeColor, vsEInvoice.Row, vsEInvoice.Col)
End Sub
Private Sub InitEinvoiceGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化电子发票网格控件
    '编制:刘兴洪
    '日期:2020-03-25 17:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsEInvoice
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        .Clear 1: .Rows = 2
        .Cols = 11
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "记录状态": i = i + 1
        .TextMatrix(0, i) = "结算ID": i = i + 1
        .TextMatrix(0, i) = "发票代码": i = i + 1
        .TextMatrix(0, i) = "发票号码": i = i + 1
        .TextMatrix(0, i) = "票据金额": i = i + 1
        .TextMatrix(0, i) = "生成时间": i = i + 1
        .TextMatrix(0, i) = "换开纸质发票": i = i + 1
        .TextMatrix(0, i) = "纸质发票号": i = i + 1
        .TextMatrix(0, i) = "备注": i = i + 1
        .TextMatrix(0, i) = "操作员姓名": i = i + 1
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter: .ColAlignment(i) = flexAlignLeftCenter
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = 1000
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Select Case .ColKey(i)
            Case "记录状态"
                .ColHidden(i) = True: .ColWidth(i) = 0: .ColData(i) = "-1|1"
            Case "备注"
                .ColWidth(i) = 2000
            Case "操作员姓名"
                 .ColWidth(i) = 1000
            Case "票据金额"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
         .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .RowHeightMin = 350
        zl_vsGrid_Para_Restore 1107, vsEInvoice, Me.Name, "电子票据信息列表", False
        If .Rows < 2 Then .Rows = 2
        .Redraw = flexRDBuffered
    End With
End Sub
