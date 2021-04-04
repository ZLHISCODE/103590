VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEInvoiceBalance 
   BorderStyle     =   0  'None
   Caption         =   "开票结算对照"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vs开票结算 
      Height          =   1080
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   4035
      _cx             =   7117
      _cy             =   1905
      Appearance      =   0
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoiceBalance.frx":0000
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   0
      Top             =   2160
      Width           =   525
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2325
      _Version        =   589884
      _ExtentX        =   4101
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "开票渠道对照"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmEInvoiceBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar控件
Private mstrDBUser As String
Private mlngSys As Long, mlngModule As Long

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF5
            Call RefreshData
    End Select
End Sub

Private Sub Form_Load()
    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    vs开票结算.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub vs开票结算_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vs开票结算
        If .TextMatrix(Row, .ColIndex("结算方式")) = "" Then Exit Sub
        If .TextMatrix(Row, .ColIndex("开票结算方式")) = "" Then Exit Sub
        If Save开票结算对照(.TextMatrix(Row, .ColIndex("结算方式")), .TextMatrix(Row, .ColIndex("开票结算方式"))) = False Then Exit Sub
        Call RefreshData
    End With
End Sub

Private Sub vs开票结算_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
     If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs开票结算, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs开票结算_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs开票结算
        If Col <> .ColIndex("开票结算方式") Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vs开票结算_GotFocus()
    If vs开票结算.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs开票结算, &HFFEBD7
End Sub

Private Sub vs开票结算_LostFocus()
    If vs开票结算.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs开票结算
    OS.OpenIme False
End Sub

Private Sub vs开票结算_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    Call ShowPopup
End Sub

Public Sub RefreshData()
    '功能：刷新数据
    Dim strSql As String, i As Integer
    Dim rs开票结算 As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSql = "Select b.名称 as 结算方式, a.开票结算方式 From 开票结算对照 A, 结算方式 B Where a.结算方式(+) = b.名称 And b.性质 In (3, 4)"
    Set rs开票结算 = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    With vs开票结算
        .Clear 1
        .Rows = 2
        If rs开票结算.EOF Then Exit Sub
        For i = 1 To rs开票结算.RecordCount
            .TextMatrix(i, .ColIndex("结算方式")) = rs开票结算!结算方式
            .TextMatrix(i, .ColIndex("开票结算方式")) = Nvl(rs开票结算!开票结算方式)
            rs开票结算.MoveNext
            If i < rs开票结算.RecordCount Then .Rows = .Rows + 1
        Next
        .ColComboList(.ColIndex("开票结算方式")) = "个人账户支付|医保统筹基金支付|其他医保支付"
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo ErrHandler
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
        With cbrMenuBar.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)", cbrControl.index + 1): cbrControl.BeginGroup = True
        End With
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&E)")
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&E)", cbrControl.index + 1): cbrControl.BeginGroup = True
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("E"), conMenu_Edit_Delete
    End With
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnDelete As Boolean
    On Error Resume Next
    If Not Me.Visible Then Exit Sub

    blnDelete = vs开票结算.TextMatrix(vs开票结算.Row, vs开票结算.ColIndex("开票结算方式")) <> ""

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel…
        Control.Enabled = False
    Case conMenu_Edit_Delete
        Control.Enabled = blnDelete
    Case Else
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_Delete '删除
        Call Delete开票结算对照
    Case conMenu_View_Refresh '刷新数据
        Call RefreshData
    Case Else
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowPopup()
    '显示弹出菜单
    Dim objPopup As CommandBarPopup
    Err = 0: On Error GoTo ErrHandler
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function Save开票结算对照(ByVal str结算方式 As String, ByVal str开票结算方式 As String) As Boolean
    Dim strSql As String
    
    If str结算方式 = "" Then Exit Function
    If str开票结算方式 = "" Then Exit Function
    
    On Error GoTo errHandle
    'Zl_开票结算对照_Update
    strSql = "Zl_开票结算对照_Update("
    '操作类型_In In Number,
    strSql = strSql & 0 & ","
    '结算方式_In In 收费渠道对照.结算方式%Type,
    strSql = strSql & "'" & str结算方式 & "',"
    '开票结算方式_In In 开票结算对照.开票结算方式%Type := Null
    strSql = strSql & "'" & str开票结算方式 & "')"
    
    Call zlDatabase.ExecuteProcedure(strSql, "开票结算对照")
    
    Save开票结算对照 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Delete开票结算对照()
    Dim strSql As String, str结算方式 As String
    
    With vs开票结算
        If .Row = 0 Then Exit Sub
        str结算方式 = .TextMatrix(.Row, .ColIndex("结算方式"))
    End With
    If str结算方式 = "" Then Exit Sub
    
    On Error GoTo errHandle
    'Zl_开票结算对照_Update
    strSql = "Zl_开票结算对照_Update("
    '操作类型_In In Number,
    strSql = strSql & 1 & ","
    '结算方式_In In 收费渠道对照.结算方式%Type,
    strSql = strSql & "'" & str结算方式 & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "开票结算对照")

    Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
