VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmStoreSpace 
   Caption         =   "库房货位管理"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11475
   Icon            =   "frmStoreSpace.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11475
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   575
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   13575
      TabIndex        =   2
      Top             =   720
      Width           =   13575
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   $"frmStoreSpace.frx":6852
         Height          =   360
         Left            =   600
         TabIndex        =   3
         Top             =   150
         Width           =   10170
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   0
         Picture         =   "frmStoreSpace.frx":68C9
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList imgP 
      Left            =   3840
      Top             =   240
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
            Picture         =   "frmStoreSpace.frx":7193
            Key             =   "pic"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDetails 
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   1680
      Width           =   9855
      Begin VB.ComboBox cboRoom 
         Height          =   300
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   2400
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetails 
         Height          =   1245
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   6615
         _cx             =   11668
         _cy             =   2196
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmStoreSpace.frx":D9F5
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfNoStock 
         Height          =   1245
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   6615
         _cx             =   11668
         _cy             =   2196
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmStoreSpace.frx":DACD
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "库房"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7125
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStoreSpace.frx":DBC1
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15161
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmStoreSpace.frx":E453
      Left            =   2400
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmStoreSpace.frx":E467
   End
End
Attribute VB_Name = "frmStoreSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MCONMENU_Edit_ADJUST = 100    '未分配库房的货位处理
Private Const MCONMENU_Edit_ADD = 101
Private Const MCONMENU_Edit_UPDATE = 102
Private Const MCONMENU_Edit_DELETE = 103
Private Const MCONMENU_Edit_HELP = 104
Private Const MCONMENU_Edit_EXIT = 105
Private Const MCONMENU_Edit_ADJUSTSAVE = 106    '未分配库房的货位处理保存
Private Const MCONMENU_Edit_ADJUSTEXIT = 107    '退出未分配库房的货位处理

Private mobjPopup As CommandBar
Private mobjControl As CommandBarControl

Private Const mlngBorderColor As Long = &H0&    '选中行边框颜色
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' 没选中行边框颜色
Private Const mconlngCanColColor As Long = &HE7CFBA    '能修改列颜色为淡蓝色
Private mlng库房ID As Long
Private mblnNoStock As Boolean  '是否存在没有分批库房的货位 true-存在;false-不存在
Private mint编辑模式 As Integer '0-正常增删改模式，1-处理未分配库房的货位模式


Private Sub GetNoStockDetail()
    '查询未分配库房的货位
    Dim rsTemp As ADODB.Recordset

    On Error GoTo err

    gstrSql = "Select id,编码,名称,简码,nvl(库房id,0) as 库房id,备注 From 药品库房货位 Where 库房id is null Order by 编码 "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetNoStockDetail")
    
    mblnNoStock = Not rsTemp.EOF
    
    vsfNoStock.Rows = 1
    vsfNoStock.RowHeight(0) = 400
    Do While Not rsTemp.EOF
        With vsfNoStock
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTemp!ID
            .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTemp!编码
            .TextMatrix(.Rows - 1, .ColIndex("货位")) = rsTemp!名称
            .TextMatrix(.Rows - 1, .ColIndex("简码")) = NVL(rsTemp!简码, "")
            .TextMatrix(.Rows - 1, .ColIndex("备注")) = NVL(rsTemp!备注, "")
            
            .RowHeight(.Rows - 1) = 300
        End With

        rsTemp.MoveNext
    Loop
    
'    vsfNoStock.Cell(flexcpBackColor, 1, vsfNoStock.ColIndex("选择"), vsfNoStock.Rows - 1) = mconlngCanColColor

    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitComandBars()
    '初始化菜单：加载全部菜单，工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPublic.Icons
    
    '右键菜单
'    Set mobjPopup = cbsMain.Add("Popup", xtpBarPopup)
'    With mobjPopup.Controls
'        Set mobjControl = .Add(xtpControlButton, MCONMENU_Edit_ADJUST, "未分配货位处理")
'    End With
  
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_ADJUST, "未分配货位处理")
        cbrControlMain.Visible = mblnNoStock
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_ADD, "新增")
        cbrControlMain.BeginGroup = mblnNoStock
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_UPDATE, "修改")
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_DELETE, "删除")
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_ADJUSTSAVE, "保存分配")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_ADJUSTEXIT, "退出分配")
        cbrControlMain.Visible = False
        
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_HELP, "帮助")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_EXIT, "退出")
    End With
    
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
    cbsMain.Item(1).Visible = False
End Sub

Private Sub TBFunc_Add()
    If frmStoreSpaceCard.ShowMe(1, Val(cboRoom.ItemData(cboRoom.ListIndex)), 0, Me) = True Then
        Call GetDetails(Val(cboRoom.ItemData(cboRoom.ListIndex)))
    End If
End Sub

Private Sub TBFunc_SetNoStock(ByVal blnBegin As Boolean)
    '未分配库房的货位处理
    'blnBegin：true-开始处理；false-结束处理
    Dim objPopup As CommandBarControl
    
    mint编辑模式 = IIf(blnBegin, 1, 0)

    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_ADD, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = Not blnBegin
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_UPDATE, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = Not blnBegin
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_DELETE, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = Not blnBegin
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_ADJUST, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = Not blnBegin And mblnNoStock
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_ADJUSTSAVE, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = blnBegin
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_ADJUSTEXIT, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = blnBegin

    vsfDetails.Visible = Not blnBegin
    vsfNoStock.Visible = blnBegin
    
    If blnBegin Then
        lblComment.Caption = "说明：选择列表中的货位分配到指定库房，双击“选择”列进行选择和取消选择。"
    Else
        If mblnNoStock Then
            lblComment.Caption = "说明：1.未分配到库房的货位请选择菜单“未分配货位处理” 2.双击已有的货位进行编辑 3.按DEL键删除当前选择的货位"
        Else
            lblComment.Caption = "说明：1.双击已有的货位进行编辑 2.按DEL键删除当前选择的货位"
        End If
    End If
End Sub

Private Sub TBFunc_Update()
    If vsfDetails.Rows = 1 Then Exit Sub
    If vsfDetails.Row < 1 Then Exit Sub
    If vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("编码")) = "" Then Exit Sub
    
    If frmStoreSpaceCard.ShowMe(2, Val(cboRoom.ItemData(cboRoom.ListIndex)), Val(vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("id"))), Me) Then
        Call GetDetails(Val(cboRoom.ItemData(cboRoom.ListIndex)))
    End If
End Sub

Private Sub cboRoom_Click()
    err = 0: On Error GoTo ErrHand
    
    If mlng库房ID = cboRoom.ItemData(cboRoom.ListIndex) Then Exit Sub
    mlng库房ID = cboRoom.ItemData(cboRoom.ListIndex)
    
    Call GetDetails(mlng库房ID)
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case MCONMENU_Edit_ADJUST  '未分配货位处理
            Call TBFunc_SetNoStock(True)
        Case MCONMENU_Edit_ADD '新增
            Call TBFunc_Add
        Case MCONMENU_Edit_UPDATE '修改
            Call TBFunc_Update
        Case MCONMENU_Edit_DELETE '删除
            Call TBFunc_SetDelete
        Case MCONMENU_Edit_ADJUSTSAVE '未分配货位处理保存
            Call SetNoStock
        Case MCONMENU_Edit_ADJUSTEXIT '退出未分配货位处理
            Call TBFunc_SetNoStock(False)
            
            If cboRoom.ListIndex <> -1 Then
                Call GetDetails(Val(cboRoom.ItemData(cboRoom.ListIndex)))
            End If
        Case MCONMENU_Edit_EXIT '退出
            Unload Me
        Case MCONMENU_Edit_HELP '帮助
            Call TBFunc_SetHelp
    End Select
End Sub

Private Sub TBFunc_SetHelp()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub TBFunc_SetDelete()
    '删除货位
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    
    If vsfDetails.Rows = 1 Then Exit Sub
    If vsfDetails.Row < 1 Then Exit Sub
    If vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("编码")) = "" Then Exit Sub
    
    On Error GoTo errH
    
    gstrSql = "Select 1 From 药品货位对照 Where 货位id = [1] And Rownum < 2"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "SetDelete", vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("id")))
    
    If Not rsData.EOF Then
        strMsg = "已选定的货位已设置了存储药品，是否删除？"
    Else
        strMsg = "是否删除选定的货位？"
    End If
    
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        gstrSql = "Zl_药品库房货位_Delete("
        'id_In In 药品库房货位.id%Type
        gstrSql = gstrSql & vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("id"))
        gstrSql = gstrSql & ")"
        Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        MsgBox "删除成功！", vbInformation, gstrSysName
        
        vsfDetails.RemoveItem vsfDetails.Row
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub





Private Sub SetNoStock()
    Dim intRow As Integer
    Dim lng库房ID As Long
    Dim objCol As New Collection
    
    On Error GoTo err

    With vsfNoStock
        If .Rows <= 1 Then Exit Sub
        If cboRoom.ListIndex = -1 Then Exit Sub
        
        lng库房ID = Val(cboRoom.ItemData(cboRoom.ListIndex))

        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("选择")) = "√" Then
                '修改
                gstrSql = "Zl_药品库房货位_Update("
                'ID
                gstrSql = gstrSql & Val(.TextMatrix(intRow, .ColIndex("id")))
                '编码
                gstrSql = gstrSql & ",'" & .TextMatrix(intRow, .ColIndex("编码")) & "'"
                '名称_In   In 药品库房货位.名称%Type,
                gstrSql = gstrSql & ",'" & .TextMatrix(intRow, .ColIndex("货位")) & "'"
                '简码_In   In 药品库房货位.简码%Type,
                gstrSql = gstrSql & ",'" & .TextMatrix(intRow, .ColIndex("简码")) & "'"
                '库房id_In In 药品库房货位.库房id%Type
                gstrSql = gstrSql & "," & lng库房ID
                '备注_In In 药品库房货位.备注%Type
                gstrSql = gstrSql & "," & IIf(.TextMatrix(intRow, .ColIndex("备注")) = "", "null", "'" & .TextMatrix(intRow, .ColIndex("备注")) & "'")
                gstrSql = gstrSql & ")"
                
                objCol.Add gstrSql, "_" & objCol.Count + 1
            End If
        Next
    End With

    If objCol.Count = 0 Then
        MsgBox "未选择货位，不能保存！", vbInformation, gstrSysName
        Exit Sub
    End If
        
    For intRow = 1 To objCol.Count
        Call zldatabase.ExecuteProcedure(objCol(intRow), "货位分配")
    Next

    MsgBox "货位分配成功！", vbInformation, gstrSysName
    
    Call GetNoStockDetail
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picCondition.Move lngLeft, lngTop, lngRight - lngLeft

    Me.picDetails.Move lngLeft, picCondition.Top + picCondition.Height + 50, lngRight - lngLeft, _
        Me.ScaleHeight - Me.picCondition.Top - Me.picCondition.Height - stbThis.Height - 50
End Sub


Private Sub Form_Load()
    Dim rsData As ADODB.Recordset
    
    mint编辑模式 = 0
    
    Call GetNoStockDetail
    Call InitComandBars
    Call GetStock
    
    Call RestoreWinState(Me, App.ProductName, Me.Caption)
    
    If mblnNoStock Then
        lblComment.Caption = "说明：1.未分配到库房的货位请选择菜单“未分配货位处理” 2.双击已有的货位进行编辑 3.按DEL键删除当前选择的货位"
    Else
        lblComment.Caption = "说明：1.双击已有的货位进行编辑 2.按DEL键删除当前选择的货位"
    End If
End Sub



Private Sub GetStock()
    '获取库房信息
    Dim rsTemp As ADODB.Recordset

    On Error GoTo err

    If InStr(1, gstrPrivs, "所有库房") > 0 Then
        gstrSql = "Select distinct b.Id, b.编码, b.名称,b.简码" & vbNewLine & _
                "From 部门性质说明 A, 部门表 B" & vbNewLine & _
                "Where a.部门id = b.Id And (a.工作性质 Like '%药库' Or a.工作性质 Like '%药房' Or a.工作性质 = '制剂室')  order by b.名称"
    Else
        gstrSql = "Select distinct b.Id, b.编码, b.名称, b.简码" & vbNewLine & _
                "From 部门人员 A, 部门表 B, 部门性质说明 C" & vbNewLine & _
                "Where c.部门id = b.Id And a.部门id = b.Id And (c.工作性质 Like '%药库' Or c.工作性质 Like '%药房' Or c.工作性质 = '制剂室') And a.人员id = [1] order by b.名称"
    End If
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "部门", UserInfo.ID)
    
    cboRoom.Clear
    With rsTemp
        Do While Not rsTemp.EOF
            cboRoom.AddItem !编码 & "-" & !名称
            cboRoom.ItemData(Me.cboRoom.NewIndex) = !ID
            .MoveNext
        Loop
    End With

    If cboRoom.ListCount <= 0 Then
        MsgBox "未设置库房或当前人员不属于库房，无法设置库房货位！", vbExclamation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    Me.cboRoom.ListIndex = 0

    Exit Sub
err:

If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetDetails(ByVal lngDeptID As Long)
    '获取具体库房对应货位
    Dim rsTemp As ADODB.Recordset

    On Error GoTo err
    
    vsfDetails.RowHeight(0) = 400
    
    If lngDeptID = 0 Then
        vsfDetails.Rows = 1
        vsfDetails.Rows = 2
        Exit Sub
    End If

    gstrSql = "Select id,编码,名称,简码,nvl(库房id,0) as 库房id,备注 From 药品库房货位 Where 库房id = [1] Order by 编码 "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "货位", lngDeptID)

    vsfDetails.Rows = 1
    Do While Not rsTemp.EOF
        With vsfDetails
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTemp!ID
            .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTemp!编码
            .TextMatrix(.Rows - 1, .ColIndex("库房id")) = rsTemp!库房id
            .TextMatrix(.Rows - 1, .ColIndex("货位")) = rsTemp!名称
            .TextMatrix(.Rows - 1, .ColIndex("简码")) = NVL(rsTemp!简码, "")
            .TextMatrix(.Rows - 1, .ColIndex("备注")) = NVL(rsTemp!备注, "")
            
            .RowHeight(.Rows - 1) = 300
        End With

        rsTemp.MoveNext
    Loop

    Exit Sub
err:

If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picDetails_Resize()
    vsfDetails.Move 0, cboRoom.Top + cboRoom.Height + 100, picDetails.Width - 20, picDetails.Height - cboRoom.Top - cboRoom.Height - 100
    vsfNoStock.Move vsfDetails.Left, vsfDetails.Top, vsfDetails.Width, vsfDetails.Height
End Sub





Private Sub vsfDetails_DblClick()
    If vsfDetails.Rows = 1 Then Exit Sub
    If vsfDetails.MouseRow < 1 Then Exit Sub
    If vsfDetails.TextMatrix(vsfDetails.MouseRow, vsfDetails.ColIndex("编码")) = "" Then Exit Sub
    
    frmStoreSpaceCard.ShowMe 2, Val(cboRoom.ItemData(cboRoom.ListIndex)), Val(vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("id"))), Me
    Call GetDetails(Val(cboRoom.ItemData(cboRoom.ListIndex)))
End Sub

Private Sub vsfDetails_EnterCell()
    '设置行选中边框
    Dim intRow As Integer
    
    With vsfDetails
        If .Rows <> 1 Then
            For intRow = 0 To .Rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, .ColIndex("编码"), .Row, .ColIndex("备注"), mlngBorderColor, 0, 2, 0, 2, 0, 2
            .CellBorderRange .Row, .ColIndex("编码"), .Row, .ColIndex("编码"), mlngBorderColor, 2, 2, 0, 2, 0, 0
            .CellBorderRange .Row, .ColIndex("备注"), .Row, .ColIndex("备注"), mlngBorderColor, 0, 2, 2, 2, 0, 0
        End If
    End With
End Sub

Private Sub vsfDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfDetails
        If KeyCode = vbKeyDelete Then
            If .Rows = 1 Then Exit Sub
            If .Row < 1 Then Exit Sub
            
            Call TBFunc_SetDelete
        End If
    End With
End Sub

Private Sub vsfDetails_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    With vsfDetails
'        If Val(.TextMatrix(.Row, .ColIndex("库房id"))) = 0 And .ColHidden(.ColIndex("分配")) = False Then
'            If .Col = .ColIndex("分配") Then
'                mobjPopup.ShowPopup
'            End If
'        End If
    End With
End Sub


Private Sub vsfNoStock_DblClick()
    With vsfNoStock
        If .Row < 1 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .Col = .ColIndex("选择") Then
            If .TextMatrix(.Row, .Col) = "√" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "√"
            End If
        End If
    End With
End Sub

Private Sub vsfNoStock_EnterCell()
    '设置行选中边框
    Dim intRow As Integer
    
    With vsfNoStock
        If .Rows <> 1 Then
            For intRow = 1 To .Rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, .ColIndex("编码"), .Row, .ColIndex("备注"), mlngBorderColor, 0, 2, 0, 2, 0, 2
            .CellBorderRange .Row, .ColIndex("编码"), .Row, .ColIndex("编码"), mlngBorderColor, 2, 2, 0, 2, 0, 0
            .CellBorderRange .Row, .ColIndex("备注"), .Row, .ColIndex("备注"), mlngBorderColor, 0, 2, 2, 2, 0, 0
        End If
    End With
End Sub


