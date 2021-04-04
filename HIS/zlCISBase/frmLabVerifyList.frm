VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmLabVerifyList 
   Caption         =   "规则设置"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   12600
   Icon            =   "frmLabVerifyList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   12600
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7410
      Left            =   4500
      ScaleHeight     =   7410
      ScaleWidth      =   7995
      TabIndex        =   26
      Top             =   105
      Width           =   8000
      Begin VB.Frame fraRule 
         Caption         =   "规则定义"
         Height          =   2300
         Left            =   105
         TabIndex        =   33
         Top             =   60
         Width           =   7800
         Begin VB.ComboBox cboValid 
            Height          =   300
            Left            =   4005
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   570
            Width           =   3645
         End
         Begin VB.TextBox txt备注 
            Height          =   900
            Left            =   4005
            MaxLength       =   200
            TabIndex        =   7
            Top             =   1260
            Width           =   3645
         End
         Begin VB.TextBox txt项目 
            Height          =   300
            Left            =   690
            TabIndex        =   4
            ToolTipText     =   "按DEL键清除项目"
            Top             =   915
            Width           =   2500
         End
         Begin VB.CommandButton cmd项目 
            Caption         =   "…"
            Height          =   300
            Left            =   3195
            TabIndex        =   42
            Top             =   915
            Width           =   300
         End
         Begin VB.TextBox txt编码 
            Height          =   285
            Left            =   4005
            MaxLength       =   3
            TabIndex        =   1
            Top             =   240
            Width           =   3660
         End
         Begin VB.ComboBox cbo仪器 
            Height          =   300
            Left            =   4005
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   915
            Width           =   3660
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   690
            MaxLength       =   30
            TabIndex        =   2
            Top             =   578
            Width           =   2800
         End
         Begin VB.TextBox txtInfo 
            Height          =   900
            Left            =   690
            MaxLength       =   200
            TabIndex        =   6
            Top             =   1275
            Width           =   2820
         End
         Begin VB.ComboBox cbo分类 
            Height          =   300
            ItemData        =   "frmLabVerifyList.frx":6852
            Left            =   690
            List            =   "frmLabVerifyList.frx":6854
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2820
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "有效"
            Height          =   180
            Left            =   3585
            TabIndex        =   45
            Top             =   630
            Width           =   360
         End
         Begin VB.Label Label14 
            Caption         =   "备注"
            Height          =   165
            Left            =   3600
            TabIndex        =   43
            Top             =   1275
            Width           =   420
         End
         Begin VB.Label Label10 
            Caption         =   "编码"
            Height          =   165
            Left            =   3585
            TabIndex        =   41
            Top             =   285
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "仪器"
            Height          =   165
            Left            =   3585
            TabIndex        =   38
            Top             =   975
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "名称"
            Height          =   165
            Left            =   270
            TabIndex        =   37
            Top             =   630
            Width           =   420
         End
         Begin VB.Label Label3 
            Caption         =   "提示"
            Height          =   165
            Left            =   285
            TabIndex        =   36
            Top             =   1275
            Width           =   420
         End
         Begin VB.Label Label11 
            Caption         =   "项目"
            Height          =   165
            Left            =   285
            TabIndex        =   35
            Top             =   975
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "分类"
            Height          =   165
            Left            =   270
            TabIndex        =   34
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame fraWhere 
         Caption         =   "适用条件"
         Height          =   1365
         Left            =   90
         TabIndex        =   27
         Top             =   2490
         Width           =   7800
         Begin VB.CheckBox chk急诊 
            Alignment       =   1  'Right Justify
            Caption         =   "急诊"
            Height          =   195
            Left            =   6790
            TabIndex        =   15
            Top             =   615
            Width           =   800
         End
         Begin VB.ComboBox cbo年龄单位 
            Height          =   300
            Left            =   4740
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   195
            Width           =   800
         End
         Begin VB.ComboBox cbo性别 
            Height          =   300
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox txt年龄下限 
            Height          =   285
            Left            =   2850
            MaxLength       =   9
            TabIndex        =   9
            Top             =   195
            Width           =   800
         End
         Begin VB.ComboBox cbo科室 
            Height          =   300
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   577
            Width           =   2600
         End
         Begin VB.ComboBox cbo病人类型 
            Height          =   300
            Left            =   4740
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   577
            Width           =   1800
         End
         Begin VB.TextBox txt诊断 
            Height          =   285
            Left            =   1065
            MaxLength       =   500
            TabIndex        =   16
            ToolTipText     =   "按DEL键清除诊断"
            Top             =   960
            Width           =   6555
         End
         Begin VB.TextBox txt年龄上限 
            Height          =   285
            Left            =   3855
            MaxLength       =   9
            TabIndex        =   10
            Top             =   195
            Width           =   800
         End
         Begin VB.CheckBox chk禁止 
            Alignment       =   1  'Right Justify
            Caption         =   "符合规则禁止审核"
            Height          =   195
            Left            =   5820
            TabIndex        =   12
            Top             =   225
            Width           =   1770
         End
         Begin VB.Label Label4 
            Caption         =   "性别"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   255
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "年龄          －"
            Height          =   165
            Left            =   2415
            TabIndex        =   31
            Top             =   255
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "送检科室"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   630
            Width           =   780
         End
         Begin VB.Label Label7 
            Caption         =   "病人类型"
            Height          =   225
            Left            =   3915
            TabIndex        =   29
            Top             =   645
            Width           =   810
         End
         Begin VB.Label Label13 
            Caption         =   "临床诊断"
            Height          =   165
            Left            =   240
            TabIndex        =   28
            Top             =   1005
            Width           =   780
         End
      End
      Begin VB.TextBox txtRule 
         Height          =   2625
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   4335
         Width           =   3880
      End
      Begin VB.CommandButton cmdSetEspecial 
         Caption         =   "设置(&S)"
         Height          =   350
         Left            =   6780
         TabIndex        =   20
         Top             =   7005
         Width           =   1100
      End
      Begin VB.TextBox txtEspecial 
         Height          =   2610
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   4335
         Width           =   3880
      End
      Begin VB.CommandButton cmdSetRule 
         Caption         =   "编辑(&E)"
         Height          =   350
         Left            =   105
         TabIndex        =   17
         Top             =   6990
         Width           =   1100
      End
      Begin VB.OptionButton optAnd 
         Caption         =   "AND"
         Height          =   315
         Left            =   3480
         TabIndex        =   18
         Top             =   3945
         Width           =   720
      End
      Begin VB.OptionButton optOr 
         Caption         =   "OR"
         Height          =   315
         Left            =   4170
         TabIndex        =   19
         Top             =   3945
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "普通规则 "
         Height          =   225
         Left            =   2490
         TabIndex        =   40
         Top             =   4005
         Width           =   810
      End
      Begin VB.Label Label9 
         Caption         =   "特殊规则"
         Height          =   225
         Left            =   4830
         TabIndex        =   39
         Top             =   4005
         Width           =   870
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   180
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5625
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   255
      ScaleHeight     =   5295
      ScaleWidth      =   3540
      TabIndex        =   24
      Top             =   735
      Width           =   3540
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4590
         Left            =   15
         TabIndex        =   0
         Top             =   60
         Width           =   3390
         _Version        =   589884
         _ExtentX        =   5980
         _ExtentY        =   8096
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   90
         Top             =   4800
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
               Picture         =   "frmLabVerifyList.frx":6856
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   7605
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabVerifyList.frx":2E630
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17145
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   465
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabVerifyList.frx":2EEC2
      Left            =   2100
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabVerifyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    图标 = 0: 分类: ID: 编码: 名称: 项目: 项目id: 仪器: 仪器Id: 科室ID: 病人类型: 性别: 年龄下限: 年龄上限: 年龄单位: 诊断: 规则: 特殊规则: 规则关系: 提示信息: 有效: 审核: 备注
End Enum
Const conPane_List = 201
Const conPane_Edit = 202

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串

Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim mLngEditWidth As Long

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim mlngItemID As Long
Dim mintEditState As Integer '当前编辑状态：0-非编辑状态,1-编辑状态
Dim mstrMatch As String
Private mstr项目 As String

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cmdSetEspecial_Click()
    Dim strRule As String
    Dim lng诊疗项目ID As Long, lng仪器ID As Long
    
    strRule = Trim(Me.txtEspecial)
    lng诊疗项目ID = Val(Me.txt项目.Tag)
    lng仪器ID = Val(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex))

    Me.txtEspecial = frmLabVerifyEspecial.DefFormula(lng诊疗项目ID, lng仪器ID, strRule, Me)

End Sub

Private Sub cmdSetRule_Click()
    Dim strRule As String
    Dim lng诊疗项目ID As Long, lng仪器ID As Long
    
    strRule = Trim(Me.txtRule)
    lng诊疗项目ID = Val(Me.txt项目.Tag)
    lng仪器ID = Val(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex))

    Me.txtRule = frmLabVerifySet.DefFormula(lng诊疗项目ID, lng仪器ID, strRule, Me)
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        Item.Handle = Me.picEdit.hWnd
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRetuId As Long

    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me

    Case conMenu_Edit_Save
        lngRetuId = EditSave()
        If lngRetuId <> 0 Then
            mlngItemID = lngRetuId: Call RefList(mlngItemID)
            mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
        End If
    Case conMenu_Edit_Untread
        Call EditCancel: Call RefList(mlngItemID)
        mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
    Case conMenu_Edit_NewItem
        If EditStart(True, mlngItemID) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select
    Case conMenu_Edit_Modify
        If mlngItemID = 0 Then Exit Sub
        If EditStart(False, mlngItemID) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select

    Case conMenu_Edit_Delete
        Dim strMsg As String
        With Me.rptList
            strMsg = "真的删除该规则吗？" & vbCrLf & "——" & .FocusedRow.Record(mCol.名称).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_检验审核规则_Edit(3," & mlngItemID & ")"

            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Err = 0: On Error GoTo 0
            mlngItemID = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                lngRetuId = lngRetuId + 1
            ElseIf lngRetuId > 0 Then
                lngRetuId = lngRetuId - 1
            End If
            If .Rows(lngRetuId).GroupRow = False Then mlngItemID = .Rows(lngRetuId - 1).Record(mCol.ID).Value
            Call RefList(mlngItemID)
        End With

    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call RefList(mlngItemID)

    Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If

    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (mintEditState <> 0)
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
        If Control.Enabled Then Control.Enabled = mlngItemID <> 0
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs

    mintEditState = 0
    mlngItemID = 0
    mstr项目 = ""

    mstrMatch = gstrMatch

    mLngEditWidth = picEdit.Width

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With

    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next


    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane

    Set panThis = dkpMan.CreatePane(conPane_List, 450, 580, DockLeftOf, Nothing)
    panThis.Title = "规则列表"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panThis = dkpMan.CreatePane(conPane_Edit, 550, 580, DockRightOf, Nothing)
    panThis.Title = "规则编辑"
    panThis.Options = PaneNoCaption

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True

    '-----------------------------------------------------
    With Me.rptList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '必须在列设置之前设置，才能生效
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.分类, "分类", 70, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编码, "编码", 60, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "名称", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.项目, "项目", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.项目id, "项目ID", 120, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.仪器, "仪器", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.仪器Id, "仪器ID", 70, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.科室ID, "科室ID", 70, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.病人类型, "病人类型", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.性别, "性别", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.年龄上限, "年龄上限", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.年龄下限, "年龄下限", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.年龄单位, "年龄单位", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        'Set rptCol = .Columns.Add(mCol.疾病id, "疾病id", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.诊断, "诊断", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.规则, "规则", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.特殊规则, "特殊规则", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.规则关系, "规则关系", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.提示信息, "提示信息", 60, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.有效, "有效", 60, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.审核, "审核", 60, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False

        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入

    Call RefSelect
    Call RefList(0)
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub

    Set panThis = Me.dkpMan.FindPane(conPane_Edit)

    panThis.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize 0, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height

End Sub

Private Sub picEdit_Resize()
    Me.cmdSetRule.Left = Me.txtRule.Left
    Me.cmdSetRule.Top = Me.picEdit.ScaleHeight - Me.cmdSetRule.Height - 45
    Me.cmdSetEspecial.Left = Me.txtEspecial.Left + Me.txtEspecial.Width - Me.cmdSetEspecial.Width
    Me.cmdSetEspecial.Top = Me.cmdSetRule.Top
    
    With Me.txtRule
        .Height = Me.picEdit.ScaleHeight - .Top - Me.cmdSetRule.Height - 45
        Me.txtEspecial.Height = .Height
    End With
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_SelectionChanged()

    With Me.rptList
        If .FocusedRow Is Nothing Then
            mlngItemID = 0
        ElseIf .FocusedRow.GroupRow = True Then
            mlngItemID = 0
        Else
            mlngItemID = Val(.FocusedRow.Record.Item(mCol.ID).Value)
        End If
        Call RefRule(mlngItemID)
    End With
End Sub

Private Sub txt项目_GotFocus()
    Call zlControl.TxtSelAll(txt项目)
End Sub

Private Sub txt项目_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mintEditState = 1 Then
        '清除原来的设置
        mstr项目 = ""
        Me.txt项目.Tag = ""
    End If
End Sub

Private Sub txt项目_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset

    If KeyAscii = vbKeyReturn Then
        If Me.txt项目 <> mstr项目 Then
            Set rsTmp = Select项目(Trim(Me.txt项目))
            If rsTmp Is Nothing Then
                Me.txt项目 = mstr项目
            Else
                Me.txt项目 = rsTmp("名称") & "(" & rsTmp("编码") & ")": Me.txt项目.Tag = rsTmp("ID"): mstr项目 = Me.txt项目
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        Exit Sub

    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt项目_LostFocus()
    If Me.txt项目 <> mstr项目 Then Me.txt项目 = mstr项目
End Sub

Private Sub cmd项目_Click()
    Dim rsTmp As ADODB.Recordset

    Set rsTmp = Select项目
    If Not rsTmp Is Nothing Then
        Me.txt项目 = rsTmp("名称") & "(" & rsTmp("编码") & ")": Me.txt项目.Tag = rsTmp("ID"): mstr项目 = Me.txt项目
    End If
End Sub


Private Sub txt诊断_GotFocus()
    Call zlControl.TxtSelAll(txt诊断)
End Sub

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Private Sub RefRule(ByVal lngItem As Long)
    '功能：刷新规则
    Dim intIndex As Integer
    If rptList.FocusedRow Is Nothing Then
        Exit Sub
    End If
    If Not rptList.FocusedRow.GroupRow Then
        With rptList.FocusedRow.Record
            '- 规则定义
            Me.txt编码 = .Item(mCol.编码).Value
            Me.txtName = .Item(mCol.名称).Value
            Me.txtInfo = .Item(mCol.提示信息).Value
            Me.txt备注 = .Item(mCol.备注).Value

            If .Item(mCol.分类).Value = "X-未分类" Then
                cbo分类.ListIndex = 0
            Else
                For intIndex = 0 To cbo分类.ListCount - 1
                    If cbo分类.List(intIndex) = .Item(mCol.分类).Value Then
                        cbo分类.ListIndex = intIndex: Exit For
                    End If
                Next
            End If

            If Val(.Item(mCol.仪器Id).Value) = 0 Then
                cbo仪器.ListIndex = 0
            Else
                For intIndex = 0 To cbo仪器.ListCount - 1
                    If cbo仪器.ItemData(intIndex) = Val(.Item(mCol.仪器Id).Value) Then
                        cbo仪器.ListIndex = intIndex: Exit For
                    End If
                Next
            End If

            txt项目 = .Item(mCol.项目).Value: txt项目.Tag = Val(.Item(mCol.项目id).Value): mstr项目 = txt项目
            cboValid.ListIndex = Val(.Item(mCol.有效).Value)

            '--适用条件
            cbo性别.ListIndex = 0
            For intIndex = 0 To cbo性别.ListCount - 1
                If cbo性别.List(intIndex) = .Item(mCol.性别).Value Then
                    cbo性别.ListIndex = intIndex: Exit For
                End If
            Next

            txt年龄下限 = .Item(mCol.年龄下限).Value
            txt年龄上限 = .Item(mCol.年龄上限).Value

            cbo年龄单位.ListIndex = 0
            For intIndex = 0 To cbo年龄单位.ListCount - 1
                If cbo年龄单位.List(intIndex) = .Item(mCol.年龄单位).Value Then
                    cbo年龄单位.ListIndex = intIndex: Exit For
                End If
            Next

            chk禁止 = Val(.Item(mCol.审核).Value)

            cbo科室.ListIndex = 0
            For intIndex = 0 To cbo科室.ListCount - 1
                If cbo科室.ItemData(intIndex) = .Item(mCol.科室ID).Value Then
                    cbo科室.ListIndex = intIndex: Exit For
                End If
            Next

            cbo病人类型.ListIndex = 0
            For intIndex = 0 To cbo病人类型.ListCount - 1
                If Val(cbo病人类型.List(intIndex)) = Val(.Item(mCol.病人类型).Value) Then
                    cbo病人类型.ListIndex = intIndex: Exit For
                End If
            Next

            txt诊断 = .Item(mCol.诊断).Value

            '--设置
            txtRule = .Item(mCol.规则).Value
            txtEspecial = .Item(mCol.特殊规则).Value
            If UCase(.Item(mCol.规则关系).Value) = "AND" Then
                Me.optAnd = True
                Me.optOr = False
            Else
                Me.optAnd = False
                Me.optOr = True
            End If
        End With
    End If
End Sub

Private Function RefList(ByVal lngItemID As Long) As Long
    '功能：刷新列表
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    strSql = "Select  Nvl(A.分类, 'X-未分类') As 分类,A.ID, A.编码, A.名称,C.名称||'('||C.编码 || ')' As 项目, A.项目id, B.名称||'('||B.编码|| ')' As 仪器," & vbNewLine & _
            "       A.仪器id, A.科室id, A.病人类型, A.性别, A.年龄下限, A.年龄上限, A.年龄单位, A.诊断, A.规则, A.特殊规则, A.规则关系, A.提示信息, A.有效, A.审核," & vbNewLine & _
            "       A.备注" & vbNewLine & _
            "From 诊疗项目目录 C, 检验仪器 B, 检验审核规则 A" & vbNewLine & _
            "Where A.仪器id = B.ID(+) And A.项目id = C.ID(+) " & vbNewLine & _
            "Order By A.分类, A.编码"
    Err = 0: On Error GoTo ErrHand
    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsRecord
        Do While Not .EOF

            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem("0"): rptItem.Icon = 0

            rptRcd.AddItem CStr("" & !分类)
            rptRcd.AddItem CStr("" & !ID)
            Set rptItem = rptRcd.AddItem(CStr("" & !编码)): rptItem.SortPriority = Val(("" & !编码))
            If Val("" & !ID) = 0 Then
                rptRcd.AddItem CStr("...该分类下没有规则...")
            Else
                rptRcd.AddItem CStr("" & !名称)
            End If
            rptRcd.AddItem IIf(CStr("" & !项目) = "()", "", CStr("" & !项目))
            rptRcd.AddItem CStr("" & !项目id)
            rptRcd.AddItem IIf(CStr("" & !仪器) = "()", "", CStr("" & !仪器))
            rptRcd.AddItem CStr("" & !仪器Id)
            rptRcd.AddItem CStr("" & !科室ID)
            rptRcd.AddItem CStr("" & !病人类型)
            rptRcd.AddItem CStr("" & !性别)
            rptRcd.AddItem CStr("" & !年龄下限)
            rptRcd.AddItem CStr("" & !年龄上限)
            rptRcd.AddItem CStr("" & !年龄单位)
            rptRcd.AddItem CStr("" & !诊断)
            rptRcd.AddItem Tran显示公式(CStr("" & !规则))
            rptRcd.AddItem Tran显示公式(CStr("" & !特殊规则))
            rptRcd.AddItem CStr("" & !规则关系)
            rptRcd.AddItem CStr("" & !提示信息)
            rptRcd.AddItem CStr("" & !有效)
            rptRcd.AddItem CStr("" & !审核)
            rptRcd.AddItem CStr("" & !备注)

            .MoveNext
        Loop
    End With

    With Me.rptList
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mCol.分类)
        .GroupsOrder(0).SortAscending = True
        .Populate
    End With

    If lngItemID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngItemID Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    Call rptList_SelectionChanged

    RefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "共有" & Me.rptList.Records.Count & "条项目"
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    RefList = Me.rptList.Records.Count
End Function

Private Sub RefSelect()
    '刷新选择项目
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    '--分类
    On Error GoTo ErrHand
    cbo分类.Clear
    cbo分类.AddItem ""
    strSql = "Select 编码,名称 From 检验审核类别 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Do Until rsTmp.EOF
        cbo分类.AddItem "" & rsTmp.Fields("编码") & "-" & rsTmp.Fields("名称")
        rsTmp.MoveNext
    Loop
    cbo分类.ListIndex = 0

    '仪器
    cbo仪器.Clear
    cbo仪器.AddItem ""
    strSql = "Select ID, 编码, 名称 From 检验仪器 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Do Until rsTmp.EOF
        cbo仪器.AddItem rsTmp.Fields("名称") & "(" & rsTmp.Fields("编码") & ")"
        cbo仪器.ItemData(cbo仪器.ListCount - 1) = rsTmp.Fields("ID")
        rsTmp.MoveNext
    Loop
    cbo仪器.ListIndex = 0

    '性别
    cbo性别.Clear
    cbo性别.AddItem ""
    cbo性别.AddItem "男"
    cbo性别.AddItem "女"
    cbo性别.ListIndex = 0

    '年龄单位
    cbo年龄单位.Clear
    cbo年龄单位.AddItem ""
    cbo年龄单位.AddItem "小时"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.ListIndex = 0

    '送检科室
    cbo科室.Clear
    cbo科室.AddItem ""
    strSql = "Select A.ID, A.编码, A.名称, B.服务对象 From 部门性质说明 B, 部门表 A Where A.ID = B.部门id And B.工作性质 = '临床'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Do Until rsTmp.EOF
        cbo科室.AddItem rsTmp.Fields("名称") & "(" & rsTmp.Fields("编码") & ")"
        cbo科室.ItemData(cbo科室.ListCount - 1) = rsTmp.Fields("ID")
        rsTmp.MoveNext
    Loop
    cbo科室.ListIndex = 0

    '病人类型
    cbo病人类型.Clear
    cbo病人类型.AddItem ""
    cbo病人类型.AddItem "1-门诊"
    cbo病人类型.AddItem "2-住院"
    
    '有效性
    cboValid.Clear
    cboValid.AddItem "0-禁止使用该规则"
    cboValid.AddItem "1-审核时使用该规则"
    cboValid.AddItem "2-仅批量审核时使用该规则"
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function EditSave() As Long
    '保存数据
    Dim strSql As String
    Dim lngNewId As Long, rsTmp As ADODB.Recordset
    '一般数据检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: EditSave = 0: Exit Function
    End If
    If Val(Me.txt编码.Text) > Val(String(Me.txt编码.MaxLength, "9")) Then
        MsgBox "编码太大！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: EditSave = 0: Exit Function
    End If
    
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "请输入名称！", vbInformation, gstrSysName
        Me.txtName.SetFocus: EditSave = 0: Exit Function
    End If
    
    Err = 0: On Error GoTo ErrHand
    If Trim(Me.cbo分类.List(cbo分类.ListIndex)) <> "" Then
        strSql = "select 编码,名称 From 检验审核类别 where 编码=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Mid(Trim(Me.cbo分类.List(cbo分类.ListIndex)), 1, InStr(Trim(Me.cbo分类.List(cbo分类.ListIndex)), "-") - 1))
        If rsTmp.RecordCount <= 0 Then
            MsgBox Trim(Me.cbo分类.List(cbo分类.ListIndex)) & "已被其他人删除，请重新选择！", vbInformation, gstrSysName
            Me.cbo分类.SetFocus: EditSave = 0: Exit Function
        End If
    End If

    '数据保存语句组织
    strSql = "'" & Replace(Trim(Me.txt编码.Text), "'", "") & "','" & Replace(Trim(Me.txtName.Text), "'", "") & "','" & Me.cbo分类.List(Me.cbo分类.ListIndex) & "'" & _
              "," & IIf(Val(txt项目.Tag) = 0, "Null", Val(txt项目.Tag)) & "," & _
              IIf(Val(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)) = 0, "Null", Val(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex))) & "," & _
              IIf(Val(Me.cbo科室.ItemData(Me.cbo科室.ListIndex)) = 0, "Null", Val(Me.cbo科室.ItemData(Me.cbo科室.ListIndex))) & "," & _
              IIf(Val(Me.cbo病人类型.List(Me.cbo病人类型.ListIndex)) = 0, "Null", "'" & Val(Me.cbo病人类型.List(Me.cbo病人类型.ListIndex)) & "'") & "," & _
              IIf(Trim(Me.cbo性别.List(Me.cbo性别.ListIndex)) = "", "Null", "'" & Me.cbo性别.List(Me.cbo性别.ListIndex) & "'") & "," & _
              IIf(Val(txt年龄下限) = 0, "Null", "'" & Val(txt年龄下限) & "'") & "," & _
              IIf(Val(txt年龄上限) = 0, "Null", "'" & Val(txt年龄上限) & "'") & "," & _
              IIf(Trim(Me.cbo年龄单位.List(Me.cbo年龄单位.ListIndex)) = "", "Null", "'" & Me.cbo年龄单位.List(Me.cbo年龄单位.ListIndex) & "'") & "," & _
              "'" & Replace(Trim(txt诊断), "'", "") & "'," & _
              IIf(Trim(txtRule) = "", "Null", "'" & Tran保存公式(Replace(Trim(txtRule), "'", "''") & "'")) & "," & _
              IIf(Trim(txtEspecial) = "", "Null", "'" & Tran保存公式(Replace(Trim(txtEspecial), "'", "''") & "'")) & "," & _
              IIf(optAnd, "'AND'", "'OR'") & "," & _
              IIf(Trim(txtInfo) = "", "Null", "'" & Replace(Trim(txtInfo), "'", "''") & "'") & "," & _
              IIf(chk急诊.Value = 1, "'1'", "'0'") & "," & _
              IIf(chk禁止.Value = 1, "'1'", "'0'") & ",'" & _
              Val(cboValid.List(cboValid.ListIndex)) & "'," & _
              IIf(Trim(txt备注) = "", "Null", "'" & Replace(Trim(txt备注), "'", "''") & "'")

    lngNewId = mlngItemID
    If Me.picEdit.Tag = "增加" Then
        lngNewId = zlDatabase.GetNextId("检验审核规则")
        strSql = "Zl_检验审核规则_Edit(1," & lngNewId & "," & strSql & ")"
    Else
        strSql = "Zl_检验审核规则_Edit(2," & lngNewId & "," & strSql & ")"
    End If

    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    'Call SQLTest(App.ProductName, Me.Caption, strSQL): gcnOracle.Execute strSQL, , adCmdStoredProc: Call SQLTest

    If Me.picEdit.Tag = "增加" Then mlngItemID = lngNewId

    Me.picEdit.Tag = "":    mintEditState = 0
    Me.picList.Enabled = True: Me.picEdit.Enabled = False: Me.picEdit.BackColor = &H8000000F

    EditSave = mlngItemID: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    EditSave = 0: Exit Function
End Function

Private Function EditStart(blnAdd As Boolean, lngItemID As Long) As Boolean
    '开始编辑
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    If blnAdd Then
        '定义
        strSql = "Select Max(编码) as 编码 From 检验审核规则"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Me.txt编码 = zlCommFun.IncStr(IIf(Trim("" & rsTmp.Fields("编码")) = "", "000", Trim("" & rsTmp.Fields("编码"))))
        Else
            Me.txt编码 = "001"
        End If
        Me.txtName = ""
        If Me.cbo分类.ListCount > 0 Then Me.cbo分类.ListIndex = 0
        Me.txt项目 = "": Me.txt项目.Tag = "": mstr项目 = ""
        If Me.cbo仪器.ListCount > 0 Then Me.cbo仪器.ListIndex = 0
        cboValid.ListIndex = 1
        Me.txtInfo = ""
        Me.txt备注 = ""

        '适用条件
        If Me.cbo性别.ListCount > 0 Then Me.cbo性别.ListIndex = 0
        Me.txt年龄下限 = "": Me.txt年龄上限 = ""
        If Me.cbo年龄单位.ListCount > 0 Then Me.cbo年龄单位.ListIndex = 0
        chk禁止.Value = 0
        If Me.cbo科室.ListCount > 0 Then Me.cbo科室.ListIndex = 0
        If Me.cbo病人类型.ListCount > 0 Then Me.cbo病人类型.ListIndex = 0
        chk急诊.Value = 0
        Me.txt诊断 = ""
        '设置
        Me.optOr.Value = True
        Me.txtRule = "": Me.txtEspecial = ""

    End If
    picEdit.Tag = IIf(blnAdd, "增加", "修改")
    picList.Enabled = False
    picEdit.Enabled = True: Me.picEdit.BackColor = RGB(250, 250, 250)
    mintEditState = 1
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub EditCancel()
    '取消
    picList.Enabled = True
    picEdit.Enabled = False: Me.picEdit.BackColor = &H8000000F
    mintEditState = 0

End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub

    '-------------------------------------------------
    '复制数据表格
    If zlControl.RPTCopyToVSF(Me.rptList, Me.vfgList) Is Nothing Then Exit Sub
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "检验审核规则"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Function Select项目(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer
    
    On Error GoTo ErrHand
    If Len(strName) = 0 Then
        '所有项目
        strSql = "Select 0 As 末级, ID, 上级id, 编码, 名称" & vbNewLine & _
                "From 诊疗分类目录 A" & vbNewLine & _
                "Where 类型 = 5" & vbNewLine & _
                "Start With 上级id is null" & vbNewLine & _
                "Connect By Prior A.id = A.上级ID" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Distinct 1 As 末级, ID, 分类id, 编码, 名称 From 诊疗项目目录 Where Nvl(单独应用, 0) = 1 And 类别 = 'C'"
        Set Select项目 = zlDatabase.ShowSelect(Me, strSql, 2, "检验项目", , , , , True)
    Else
        '指定项目
        strSQLItem = " From 诊疗项目别名 B,诊疗项目目录 A" & _
            " Where A.ID=B.诊疗项目ID And Nvl(A.单独应用, 0) = 1 And A.类别 = 'C'" & _
            " And (Upper(A.编码) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.简码) Like '" & mstrMatch & UCase(strName) & "%')"
'
'        strSQL = "Select distinct  0 As 末级, ID, 上级id, 编码, 名称" & vbNewLine & _
'                "From 诊疗分类目录 A" & vbNewLine & _
'                "Where 类型 = 5" & vbNewLine & _
'                "Start With ID In (Select A.分类id " & strSQLItem & ")" & vbNewLine & _
'                "Connect By Prior A.上级id = A.ID" & vbNewLine & _
'                "Union All" & vbNewLine & _
'                "Select Distinct 1 As 末级, A.ID, A.分类id, A.编码, A.名称 " & strSQLItem
                
        strSql = "Select Distinct  A.ID, A.分类id, A.编码, A.名称 " & strSQLItem
        Set Select项目 = zlDatabase.ShowSelect(Me, strSql, 0, "检验项目", , , , , True)
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Select诊断(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer

    If Len(strName) = 0 Then
        '所有项目
        strSql = "Select Distinct 0 As 末级, ID, 上级id, to_char(序号) as 编码, 名称" & vbNewLine & _
                "From 疾病编码分类 A" & vbNewLine & _
                "Where 类别='D'" & vbNewLine & _
                "Start With 上级ID IS NULL " & vbNewLine & _
                "Connect By Prior A.id = A.上级ID" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Distinct 1 As 末级, ID, 分类id, 编码, 名称 From 疾病编码目录 Where 类别='D'"
    Else
        '指定项目
        strSQLItem = " From 疾病编码目录 A" & _
            " Where A.类别 = 'D'" & _
            " And (Upper(A.编码) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(A.简码) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = "Select Distinct 0 As 末级, ID, 上级id,to_char(序号) as 编码, 名称" & vbNewLine & _
                "From 疾病编码分类 A" & vbNewLine & _
                "Where 类别='D'" & vbNewLine & _
                "Start With ID In (Select A.分类id " & strSQLItem & ")" & vbNewLine & _
                "Connect By Prior A.上级id = A.ID" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Distinct 1 As 末级, A.ID, A.分类id, A.编码, A.名称 " & strSQLItem
    End If
    Set Select诊断 = zlDatabase.ShowSelect(Me, strSql, 2, "疾病", , , , , True)

End Function

Private Function Tran保存公式(ByVal str显示公式 As String) As String
    '将显示公式转为保存公式
    Dim strItem As String, strTmp As String, strLast As String
    Dim rsGS As ADODB.Recordset, lngLength As Long
    strItem = "": strTmp = ""
    On Error GoTo ErrHand
    If str显示公式 <> "" Then
        Do While str显示公式 Like "*[[]*[]]*"
            strTmp = strTmp & Mid(str显示公式, 1, InStr(str显示公式, "[") - 1)
            lngLength = InStr(str显示公式, "]") - InStr(str显示公式, "[") - 1
            strItem = Mid(str显示公式, InStr(str显示公式, "[") + 1, lngLength)
            If InStr(strItem, "_") > 0 Then
                strItem = Mid(strItem, 1, InStr(strItem, "_") - 1)
            End If
            If InStr(strItem, "上次.") > 0 Then
                strLast = "上次."
                strItem = Replace(strItem, "上次.", "")
            ElseIf InStr(strItem, "标记.") > 0 Then
                strLast = "标记."
                strItem = Replace(strItem, "标记.", "")
            Else
                strLast = ""
            End If
            gstrSql = "Select ID,英文名 From 诊治所见项目  Where (id=[1] or 编码=[2]) "
            Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(strItem), strItem)

            Do Until rsGS.EOF
                strTmp = strTmp & "[" & strLast & Val("" & rsGS.Fields("ID")) & "]"
                rsGS.MoveNext
            Loop
            str显示公式 = Mid(str显示公式, InStr(str显示公式, "]") + 1)
        Loop
        strTmp = strTmp & Mid(str显示公式, InStr(str显示公式, "]") + 1)
        strTmp = Replace(strTmp, "{D:漏项检查}", "{D:1}")
        strTmp = Replace(strTmp, "{D:多项检查}", "{D:2}")
        strTmp = Replace(strTmp, "{D:漏项多项检查}", "{D:3}")
        Tran保存公式 = strTmp
        
        
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Tran显示公式(ByVal str保存公式 As String) As String
    '将保存公式转为显示公式
    Dim strTmp As String, strItem As String, strLast As String
    Dim rsGS As ADODB.Recordset, lngLength As Long
    On Error GoTo ErrHand
    Do While str保存公式 Like "*[[]*[]]*"
        strTmp = strTmp & Mid(str保存公式, 1, InStr(str保存公式, "[") - 1)
        lngLength = InStr(str保存公式, "]") - InStr(str保存公式, "[") - 1
        strItem = Mid(str保存公式, InStr(str保存公式, "[") + 1, lngLength)
        If InStr(strItem, "上次.") > 0 Then
            strLast = "上次."
            strItem = Replace(strItem, "上次.", "")
        ElseIf InStr(strItem, "标记.") > 0 Then
            strLast = "标记."
            strItem = Replace(strItem, "标记.", "")
        Else
            strLast = ""
        End If
        gstrSql = "Select ID,英文名,编码 From 诊治所见项目 Where id=[1] "
        Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(strItem))
        Do Until rsGS.EOF
            If Trim("" & rsGS.Fields("英文名")) <> "" Then
                strTmp = strTmp & "[" & strLast & rsGS.Fields("编码") & "_" & Trim("" & rsGS.Fields("英文名")) & "]"
            Else
                strTmp = strTmp & "[" & strLast & Val(strItem) & "]"
            End If
            rsGS.MoveNext
        Loop
        str保存公式 = Mid(str保存公式, InStr(str保存公式, "]") + 1)
    Loop
    strTmp = strTmp & Mid(str保存公式, InStr(str保存公式, "]") + 1)
    strTmp = Replace(strTmp, "{D:1}", "{D:漏项检查}")
    strTmp = Replace(strTmp, "{D:2}", "{D:多项检查}")
    strTmp = Replace(strTmp, "{D:3}", "{D:漏项多项检查}")
    Tran显示公式 = strTmp
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
