VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEInvoiceFees 
   BorderStyle     =   0  'None
   Caption         =   "收据费目对照"
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8835
   Icon            =   "frmEInvoiceFees.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vs收据费目 
      Height          =   1080
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2715
      _cx             =   4789
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoiceFees.frx":6852
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1035
      Left            =   360
      Top             =   2760
      Width           =   525
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2325
      _Version        =   589884
      _ExtentX        =   4101
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "收据费目对照"
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
Attribute VB_Name = "frmEInvoiceFees"
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
    vs收据费目.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
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
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&R)", cbrControl.Index + 1): cbrControl.BeginGroup = True
        End With
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
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
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&E)", cbrControl.Index + 1)
    End With
    
    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("E"), conMenu_Edit_Delete
    End With
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnEnable As Boolean
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    blnEnable = Val(vs收据费目.TextMatrix(vs收据费目.Row, vs收据费目.ColIndex("ID"))) <> 0

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel…
        Control.Enabled = False
    Case conMenu_Edit_Modify
        Control.Enabled = blnEnable
    Case conMenu_Edit_Delete
        Control.Enabled = blnEnable
    Case Else
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim objfrmEInvoiceParaSet As frmEInvoiceParaSet
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    Case conMenu_Edit_NewItem '新增
        Call AddNewEInvoiceFees
    Case conMenu_Edit_Modify  '修改
        Call ModifyEInvoiceFees
    Case conMenu_Edit_Delete '删除
        Call DeleteEInvoiceFees
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

Public Sub AddNewEInvoiceFees()
    '新增收据费目对照
    Dim frmEdit As New frmEInvoiceFeeseSet
    Dim blnRefresh As Boolean
    
    On Error GoTo errHandle
    Call frmEdit.ShowMe(Me, 0, blnRefresh)
    If blnRefresh Then Call RefreshData
  
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DeleteEInvoiceFees()
    '删除收据费目对照
    On Error GoTo errHandle
    Dim lngID As Long
    Dim str名称 As String, strSql As String
    
    With vs收据费目
        If .Row = 0 Then Exit Sub
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        str名称 = .TextMatrix(.Row, .ColIndex("名称"))
        If lngID = 0 Then Exit Sub
        If MsgBox("你确认要删除名称为“" & str名称 & "”的收据费目对照吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Me.MousePointer = 11
        strSql = "Zl_收据费目对照_Update(2," & lngID & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        Me.MousePointer = 0
    End With

    Call RefreshData
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.MousePointer = 0
End Sub

Public Sub ModifyEInvoiceFees()
    '修改收据费目对照
    Dim frmEdit As New frmEInvoiceFeeseSet
    Dim lngID As Long, blnRefresh As Boolean
    On Error Resume Next
    With vs收据费目
        If .Row = 0 Then Exit Sub
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
    If lngID = 0 Then Exit Sub
    Call frmEdit.ShowMe(Me, lngID, blnRefresh)
    If blnRefresh Then Call RefreshData
End Sub

Public Sub RefreshData()
    Dim strSql As String
    Dim rs收据费目 As ADODB.Recordset

    On Error GoTo errHandle
    strSql = " Select ID, 收据费目, 编码, 名称, Decode(费用场合, 0, '不区分', 1, '门诊', '住院') As 费用场合 From 收据费目对照 order by 收据费目 "
    Set rs收据费目 = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set vs收据费目.DataSource = rs收据费目
    Call SetHeader
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Integer
    
    strHead = "ID,1,0|收据费目,4,2100|编码,4,2100|名称,4,2100|费用场合,4,1800"
    With vs收据费目
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            .ColKey(i) = UCase(Trim(.TextMatrix(0, i)))
        Next
        If Not Visible Then Call RestoreFlexState(vs收据费目, App.ProductName & "\" & Me.Name)
        .ColHidden(.ColIndex("ID")) = True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
        If .Rows = 1 Then .Rows = 2
        If .Rows > 1 Then .Row = 1
        .Redraw = True
    End With

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

Private Sub vs收据费目_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Or OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vs收据费目, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vs收据费目_DblClick()
     Call ModifyEInvoiceFees
End Sub

Private Sub vs收据费目_GotFocus()
    If vs收据费目.Row <= 0 Then Exit Sub
    zl_VsGridGotFocus vs收据费目, &HFFEBD7
End Sub

Private Sub vs收据费目_LostFocus()
    If vs收据费目.Row <= 0 Then Exit Sub
    zl_VsGridLOSTFOCUS vs收据费目
    OS.OpenIme False
End Sub

Private Sub vs收据费目_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (Button = vbRightButton) Then Exit Sub
    Call ShowPopup
End Sub
