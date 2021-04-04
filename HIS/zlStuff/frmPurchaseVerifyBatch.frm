VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPurchaseVerifyBatch 
   BackColor       =   &H8000000A&
   Caption         =   "备货卫材批量审核"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmPurchaseVerifyBatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11760
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSelPatient 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4440
      ScaleHeight     =   2055
      ScaleWidth      =   4995
      TabIndex        =   14
      Top             =   960
      Width           =   4995
      Begin VB.PictureBox picTitlePatient 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   3015
         TabIndex        =   18
         Top             =   0
         Width           =   3015
         Begin VB.Label lblSelPatient 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "已选择的病人列表"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   100
            Width           =   1560
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfPatient 
         Height          =   945
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   4545
         _cx             =   8017
         _cy             =   1667
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   14
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483628
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
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
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseVerifyBatch.frx":014A
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
         ExplorerBar     =   7
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
      Begin VB.Label lblCostAmount 
         AutoSize        =   -1  'True
         Caption         =   "合计成本金额："
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1260
      End
      Begin VB.Label lblIVAmount 
         AutoSize        =   -1  'True
         Caption         =   "合计发票金额："
         Height          =   180
         Left            =   2760
         TabIndex        =   17
         Top             =   1680
         Width           =   1260
      End
   End
   Begin VB.PictureBox picSelMaterial 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4560
      ScaleHeight     =   2055
      ScaleWidth      =   4995
      TabIndex        =   10
      Top             =   3600
      Width           =   4995
      Begin VB.PictureBox picTitleMaterial 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   4815
         TabIndex        =   20
         Top             =   0
         Width           =   4815
         Begin VB.TextBox txtFindMaterial 
            Height          =   270
            Left            =   2880
            TabIndex        =   21
            Top             =   70
            Width           =   1815
         End
         Begin VB.Label lblFindMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "查找材料(&M)"
            Height          =   180
            Left            =   1800
            TabIndex        =   23
            Top             =   100
            Width           =   990
         End
         Begin VB.Label lblMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "病人使用材料明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   100
            Width           =   1560
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMaterial 
         Height          =   945
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4545
         _cx             =   8017
         _cy             =   1667
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483628
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
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
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchaseVerifyBatch.frx":021F
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
         ExplorerBar     =   7
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
      Begin VB.Label lblIV 
         AutoSize        =   -1  'True
         Caption         =   "小计发票金额："
         Height          =   180
         Left            =   2760
         TabIndex        =   13
         Top             =   1680
         Width           =   1260
      End
      Begin VB.Label lblCost 
         AutoSize        =   -1  'True
         Caption         =   "小计成本金额："
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1260
      End
   End
   Begin VB.PictureBox pic参数区 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   360
      ScaleHeight     =   4455
      ScaleWidth      =   3615
      TabIndex        =   9
      Top             =   600
      Width           =   3615
      Begin VB.CheckBox chk全选 
         Caption         =   "全选"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtPatientInfo 
         Height          =   270
         Left            =   1200
         TabIndex        =   8
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "选择填制日期"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   615
         Width           =   690
      End
      Begin MSComctlLib.ListView lvwPatient 
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   885
         TabIndex        =   1
         Top             =   240
         Width           =   2200
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "…"
         Height          =   300
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtpDateBegin 
         Height          =   315
         Left            =   885
         TabIndex        =   4
         Top             =   600
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   50987011
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtpDateEnd 
         Height          =   315
         Left            =   885
         TabIndex        =   5
         Top             =   960
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   50987011
         CurrentDate     =   36263
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "查找病人(&F)"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Width           =   990
      End
      Begin VB.Label lblProvider 
         AutoSize        =   -1  'True
         Caption         =   "供应商"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   540
      End
   End
   Begin XtremeCommandBars.CommandBars cmbMain 
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPurchaseVerifyBatch.frx":02F4
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPurchaseVerifyBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MCON_模块号 = 1712

Private Enum enm_CommandBarID
    GetData = 3052
    Verify = 8044
    Cancel = 2613
End Enum

Private mlngStockID As Long
Private mFMT As g_FmtString
Private mintUnit As Integer              '0：散装单位； 1：包装单位
Private mbln需要核查 As Boolean
Private mstrPrivs As String
Private mintMode As Integer              '0：无核查环节的审核； 1：有核查环节的核查； 2：有核查环节的审核

Private mdblIVAmountOld As Double
Private mdblIVAmountNew As Double

Private mdatBegin As Date
Private mdatEnd As Date
Private mblnUpdate As Boolean                   '是否按新的售价更新单据，主要是针对审核时定价格不致的情况
Public Sub ShowMe(ByVal frmMain As Form, ByVal strPrivs As String, ByVal intMode As Integer, ByVal lngStockID As Long)
    mlngStockID = lngStockID
    mstrPrivs = strPrivs
    mintMode = intMode
    If mintMode = 1 Then
        Caption = "备货卫材批量核查"
    Else
        Caption = "备货卫材批量审核"
    End If
    Show vbModal, frmMain
End Sub

Private Sub InitDKPMain()
'初始化dkpMain
    Dim pneParameter As Pane, pneInvoice As Pane, pneMaterial As Pane
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        
        Set pneParameter = .CreatePane(1, ScaleHeight, 250, DockLeftOf)
        pneParameter.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        pneParameter.Title = "参数设置"
        pneParameter.MinTrackSize.Width = 150
        pneParameter.MaxTrackSize.Width = 350
        
        Set pneInvoice = .CreatePane(2, 1000, 10, DockRightOf)
        pneInvoice.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        pneInvoice.Title = "选择区A"
        pneInvoice.MinTrackSize.Height = 120
        
        Set pneMaterial = .CreatePane(3, 100, 10, DockBottomOf, pneInvoice)
        pneMaterial.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        pneMaterial.Title = "选择区B"
        pneMaterial.MinTrackSize.Height = 120
        
        If Not cmbMain Is Nothing Then Call .SetCommandBars(cmbMain)
    End With
End Sub

Private Sub InitCommandBar()
    Dim cbcControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    cmbMain.VisualTheme = xtpThemeOffice2003
    With cmbMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True                 '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
    End With
    cmbMain.EnableCustomization False
'    cmbMain.Icons = frmPubIcons.imgPublic.Icons
    Set cmbMain.Icons = zlCommFun.GetPubIcons

    Set cbrToolBar = cmbMain.Add("工具栏", xtpBarTop)
    'cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagAlignTop
    With cbrToolBar.Controls
        Set cbcControl = .Add(xtpControlButton, enm_CommandBarID.GetData, "提取")
        If mintMode = 1 Then
            Set cbcControl = .Add(xtpControlButton, enm_CommandBarID.Verify, "核查")
        Else
            Set cbcControl = .Add(xtpControlButton, enm_CommandBarID.Verify, "审核")
        End If
        Set cbcControl = .Add(xtpControlButton, enm_CommandBarID.Cancel, "关闭")
        cbcControl.BeginGroup = True
    End With
    For Each cbcControl In cbrToolBar.Controls
        If cbcControl.Type = xtpControlButton Then
            cbcControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub chkDate_Click()
    dtpDateBegin.Enabled = chkDate.Value = 1
    dtpDateEnd.Enabled = chkDate.Value = 1
End Sub

Private Sub chkDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub chk全选_Click()
    Dim i As Integer
    
    vsfPatient.Rows = 1
    vsfMaterial.Rows = 1
    With lvwPatient
        For i = 1 To .ListItems.Count
            .ListItems.Item(i).Checked = chk全选.Value
            If chk全选.Value = 1 Then
                Call SelPatient(.ListItems(i))
            Else
                Call CalPatient(.ListItems(i))
            End If
        Next
    End With
End Sub

Private Sub cmbMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case enm_CommandBarID.GetData
            Call GetCheckData
        Case enm_CommandBarID.Verify
            Control.Enabled = False
            Call VerifyBatch
            Control.Enabled = True
        Case enm_CommandBarID.Cancel
            Unload Me
    End Select
End Sub

Private Sub GetCheckData()
    If chkDate.Value = 1 And DateDiff("M", dtpDateBegin.Value, dtpDateEnd.Value) > 3 Then
        MsgBox "选择填制日期范围不能超过三个月！", vbInformation, gstrSysName
        Exit Sub
    End If
    If vsfPatient.Rows > 1 Or vsfMaterial.Rows > 1 Then
        If MsgBox("已有选定的数据未处理，要覆盖它吗？", vbInformation + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "请录入“供应商”信息！", vbInformation, gstrSysName
        Exit Sub
    End If
    Dim cbcControl As CommandBarControl
    
    Set cbcControl = Me.cmbMain.FindControl(, enm_CommandBarID.GetData)
    If cbcControl Is Nothing Then Exit Sub
    
    cbcControl.Enabled = False
    MousePointer = vbHourglass
    Call FillLVWPatient
    lvwPatient.SetFocus
    MousePointer = vbDefault
    cbcControl.Enabled = True
End Sub

Private Sub cmbMain_Resize()
    On Error Resume Next
    With txtProvider
        .Width = pic参数区.Width - .Left - cmdProvider.Width - 100
    End With
    With cmdProvider
        .Left = pic参数区.Width - .Width - 100
    End With
    With dtpDateBegin
        .Width = txtProvider.Width
    End With
    With dtpDateEnd
        .Width = txtProvider.Width
    End With
    With chk全选
        .Left = chkDate.Left
        .Top = dtpDateEnd.Top + dtpDateEnd.Height + 50
    End With
    With lvwPatient
        .Top = chk全选.Top + chk全选.Height + 50
        .Left = 0
        .Width = pic参数区.Width
        .Height = pic参数区.Height - dtpDateEnd.Top - dtpDateEnd.Height - txtPatientInfo.Height - chk全选.Height - 200
    End With
    With lblFind
        .Top = pic参数区.Height - txtPatientInfo.Height - 70
        .Left = lblProvider.Left
    End With
    With txtPatientInfo
        .Top = pic参数区.Height - txtPatientInfo.Height - 100
        .Left = lblFind.Left + lblFind.Width + 50
        .Width = pic参数区.Width - lblFind.Left - lblFind.Width - 50
    End With

    With picTitlePatient
        .Top = 0
        .Left = 0
        .Width = picSelPatient.Width
        .Height = 400
    End With
    With vsfPatient
        .Top = picTitlePatient.Height
        .Left = 0
        .Width = picSelPatient.Width
        .Height = picSelPatient.Height - .Top - lblCostAmount.Height - 100 * 2
    End With
    With lblCostAmount
        .Top = picSelPatient.Height - .Height - 100
        .Left = lblSelPatient.Left
    End With
    With lblIVAmount
        .Top = lblCostAmount.Top
        .Left = picSelPatient.Width / 2
    End With

    With picTitleMaterial
        .Top = 0
        .Left = 0
        .Width = picSelPatient.Width
        .Height = 400
    End With
    With lblFindMaterial
        .Left = picSelMaterial.Width - txtFindMaterial.Width - .Width - 100
    End With
    With txtFindMaterial
        .Left = lblFindMaterial.Left + lblFindMaterial.Width + 50
    End With
    With vsfMaterial
        .Top = picTitleMaterial.Height
        .Left = 0
        .Width = picSelMaterial.Width
        .Height = picSelMaterial.Height - .Top - lblCost.Height - 200
    End With
    With lblCost
        .Top = picSelMaterial.Height - .Height - 100
        .Left = lblMaterial.Left
    End With
    With lblIV
        .Top = lblCost.Top
        .Left = picSelMaterial.Width / 2
    End With
    err.Clear: On Error GoTo 0
End Sub

Private Sub cmdProvider_Click()
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT

    vRect = zlControl.GetControlRect(txtProvider.hwnd)
    
    gstrSQL = "" & _
        "   Select id,上级ID,编码,简码,名称,末级 " & _
        "   From 供应商 " & _
        "   Where  (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
        "       And (substr(类型,5,1)=1 And (站点=[1] or 站点 is null)  Or Nvl(末级,0)=0) " & _
        "   Start with 上级ID is null connect by prior ID =上级ID " & _
        "   Order by level,ID "
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 2, "供应商", True, "", "" _
              , True, True, False, vRect.Left - 15, vRect.Top, txtProvider.Height, blnCancel, False, False, gstrNodeNo)
    If blnCancel = False Then
        If Not rsTmp Is Nothing Then
            txtProvider.Text = zlStr.Nvl(rsTmp!名称)
            txtProvider.Tag = zlStr.Nvl(rsTmp!Id)
        Else
            txtProvider.Text = ""
            txtProvider.Tag = "0"
        End If
    End If
    txtProvider.SetFocus
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1: Item.Handle = pic参数区.hwnd
        Case 2: Item.Handle = picSelPatient.hwnd
        Case 3: Item.Handle = picSelMaterial.hwnd
    End Select
End Sub

Private Sub dtpDateBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub dtpDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mblnUpdate = False
    
    dtpDateBegin.Value = DateAdd("M", -1, sys.Currentdate)
    dtpDateEnd.Value = sys.Currentdate
    Call chkDate_Click
    
    mbln需要核查 = Val(zlDatabase.GetPara("卫材外购需要核查", glngSys, "0")) = 1
    
    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, MCON_模块号, "0"))
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(0, g_售价)
    End With
    
    cmbMain.ActiveMenuBar.Visible = False
    Call InitDKPMain
    Call InitCommandBar
    Call InitLVWPatient
    Call InitVSFPatient
    Call InitVSFMaterial
    
    If Val(zlDatabase.GetPara("使用个性化风格", 0, 0)) = 1 Then
        RestoreWinState Me, App.ProductName, Me.Caption
    End If
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Width < 10000 Then Width = 10000
    If Height < 6000 Then Height = 6000
End Sub

Private Sub InitLVWPatient()
    With lvwPatient
        .Checkboxes = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .Sorted = True
        
        .ColumnHeaders.Add , "科室", "科室", 1500
        .ColumnHeaders.Add , "姓名", "姓名", 800
        .ColumnHeaders.Add , "性别", "性别", 600
        .ColumnHeaders.Add , "年龄", "年龄", 600
        .ColumnHeaders.Add , "住院号", "住院号", 800
    End With
End Sub

Private Sub InitVSFPatient()
Const conHead = "H_病人ID,,|姓名,,1000|H_科室ID,,|科室,,2000|发票号,,1000|发票代码,,1000|发票日期,,1000,d|发票金额,,1000,N|成本金额,,1000,n"
    
    With vsfPatient
        .Rows = 1
        SetVSFHead vsfPatient, conHead
    End With
End Sub

Private Sub InitVSFMaterial()
Const conHead = "H_病人ID,,|H_材料ID,,|材料名称,,2000|规格,,1000|发票金额,,1000,N|单位,,500|数量,,1000,N|成本价,,1000,N|成本金额,,1000,N" & _
                "|H_NO,,|H_序号,,|H_库房ID,,|H_供药单位ID,,|H_产地,,|H_批号,,|H_生产日期,,|H_效期,,|H_灭菌日期,,|H_灭菌效期,," & _
                "|H_扣率,,|H_零售价,,|H_零售金额,,|H_差价,,|H_差价金额,,|H_摘要,,|H_内部条码,,|H_换算系数,," & _
                "|H_注册证号,,|H_填制人,,|H_填制日期,,|H_发票号,,|H_发票日期,,|H_发票金额,,|H_核查人,,|H_核查日期,,|H_批次,," & _
                "|H_高值材料,,|H_商品条码,,|H_费用ID,,|H_Verify,,|H_科室ID,,"
                
    With vsfMaterial
        .Rows = 1
        .ExplorerBar = flexExMove
        SetVSFHead vsfMaterial, conHead
    End With
End Sub

Private Sub FillLVWPatient()
    Dim rsTmp As ADODB.Recordset
    Dim lsItem As ListItem
    Dim i As Long
    Dim lngColor As Long
    
    On Error GoTo ErrHandle
    MousePointer = vbHourglass
    '填充
    If mbln需要核查 Then
        If mintMode = 1 Then
            '核查
            gstrSQL = "Select Distinct a.病人ID, a.姓名, a.性别, b.年龄, a.住院号, b.病人科室ID, d.名称 科室 " & _
                      "From 病人信息 A, 住院费用记录 B, 药品收发记录 C, 部门表 D " & _
                      "Where a.病人id = b.病人id And b.Id = c.费用id And b.病人科室id = d.Id " & _
                      "    And c.费用ID > 0 And c.配药日期 is null And c.供药单位ID + 0 = [1] And c.库房id = [2] " & _
                      "    And c.单据 = 15 "
        Else
            '审核
            gstrSQL = "Select Distinct a.病人ID, a.姓名, a.性别, b.年龄, a.住院号, b.病人科室ID, d.名称 科室 " & _
                      "From 病人信息 A, 住院费用记录 B, 药品收发记录 C, 部门表 D " & _
                      "Where a.病人id = b.病人id And b.Id = c.费用id And b.病人科室id = d.Id " & _
                      "    And c.费用ID > 0 And c.配药日期 is not null And c.审核日期 is null " & _
                      "    And c.供药单位ID + 0 = [1] And c.库房id = [2] And c.单据 = 15 "
        End If
    Else
        '直接审核
        gstrSQL = "Select Distinct a.病人ID, a.姓名, a.性别, b.年龄, a.住院号, b.病人科室ID, d.名称 科室 " & _
                  "From 病人信息 A, 住院费用记录 B, 药品收发记录 C, 部门表 D " & _
                  "Where a.病人id = b.病人id And b.Id = c.费用id And b.病人科室id = d.Id " & _
                  "    And c.费用ID > 0 And c.审核日期 is null And c.供药单位ID + 0 = [1] And c.库房id = [2] " & _
                  "    And c.单据 = 15 "
    End If
    
    If chkDate.Value = 1 Then
        mdatBegin = dtpDateBegin.Value
        mdatEnd = dtpDateEnd.Value
    Else
        mdatBegin = DateAdd("M", -1, sys.Currentdate)
        mdatEnd = sys.Currentdate
    End If
    gstrSQL = gstrSQL & " And c.填制日期 between to_date('" & Format(mdatBegin, "yyyy-mm-dd 00:00:00") & "', 'yyyy-mm-dd hh24:mi:ss') " & _
              " And to_date('" & Format(mdatEnd, "yyyy-mm-dd 23:59:59") & "', 'yyyy-mm-dd hh24:mi:ss') " & vbNewLine
    '合并门诊数据
    gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "住院费用记录", "门诊费用记录") & " Order By 科室, 姓名 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-病人信息", Val(txtProvider.Tag), mlngStockID)
    
    With vsfPatient
        For i = .Rows - 1 To 1 Step -1
            .RemoveItem i
        Next
    End With
    With vsfMaterial
        For i = .Rows - 1 To 1 Step -1
            .RemoveItem i
        Next
    End With
    
    MousePointer = vbDefault
    
    With lvwPatient
        .ListItems.Clear
        .Tag = txtProvider.Tag
        Do While Not rsTmp.EOF
            Set lsItem = .ListItems.Add(, "_" & zlStr.Nvl(rsTmp!病人ID) & "_" & zlStr.Nvl(rsTmp!病人科室ID), zlStr.Nvl(rsTmp!科室))
            lsItem.SubItems(1) = zlStr.Nvl(rsTmp!姓名)
            lsItem.SubItems(2) = zlStr.Nvl(rsTmp!性别)
            lsItem.SubItems(3) = zlStr.Nvl(rsTmp!年龄)
            lsItem.SubItems(4) = zlStr.Nvl(rsTmp!住院号)
            rsTmp.MoveNext
        Loop
        If rsTmp.RecordCount <= 0 Then
            MsgBox "无可“" & IIf(mintMode = 1, "核查", "审核") & "”的数据！", vbInformation, gstrSysName
        End If
    End With
    
'    If mintMode <> 1 Then
        lngColor = vbBlue
'    Else
'        lngColor = vbBlack
'    End If
    With vsfPatient
        .Cell(flexcpForeColor, 0, .ColIndex("发票号"), .Rows - 1, .ColIndex("发票号")) = lngColor
        .Cell(flexcpForeColor, 0, .ColIndex("发票代码"), .Rows - 1, .ColIndex("发票代码")) = lngColor
        .Cell(flexcpForeColor, 0, .ColIndex("发票日期"), .Rows - 1, .ColIndex("发票日期")) = lngColor
        .Cell(flexcpForeColor, 0, .ColIndex("发票金额"), .Rows - 1, .ColIndex("发票金额")) = lngColor
    End With
    With vsfMaterial
        .Cell(flexcpForeColor, 0, .ColIndex("发票金额"), .Rows - 1, .ColIndex("发票金额")) = lngColor
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Val(zlDatabase.GetPara("使用个性化风格", 0, 0)) = 1 Then
        SaveWinState Me, App.ProductName, Me.Caption
    End If
End Sub

Private Sub lvwPatient_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index >= 1 And ColumnHeader.Index <= 2 Or ColumnHeader.Index = 5 Then
        lvwPatient.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub lvwPatient_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        '确定选择
        SelPatient Item
    Else
        '取消选择
        If CalPatient(Item) = False Then Item.Checked = True
    End If
    Call ShowAmount
End Sub

Private Sub txtFindMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
        If Trim(txtFindMaterial.Text) = "" Then Exit Sub
        Dim i As Long, lngStart As Long
        With vsfMaterial
            lngStart = IIf(KeyCode = vbKeyF3, IIf(.Rows - 1 > .Row, .Row + 1, 1), 1)
            For i = lngStart To .Rows - 1
                If InStr(.TextMatrix(i, .ColIndex("材料名称")), Trim(txtFindMaterial.Text)) > 0 Then
                    .Row = i
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Sub txtPatientInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
        If Trim(txtPatientInfo.Text) = "" Then Exit Sub
        Dim i As Long, lngStart As Long
        With lvwPatient
            lngStart = IIf(KeyCode = vbKeyF3, IIf(.ListItems.Count > .SelectedItem.Index, .SelectedItem.Index + 1, 1), 1)
            For i = lngStart To .ListItems.Count
                If InStr(.ListItems(i).SubItems(1), Trim(txtPatientInfo.Text)) > 0 Then
                    .ListItems.Item(i).Selected = True
                    Exit For
                End If
            Next
        End With
    End If
End Sub

Private Sub txtProvider_Change()
    With txtProvider
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim rsProvider As Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    vRect = zlControl.GetControlRect(txtProvider.hwnd)
    
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = GetMatchingSting(UCase(.Text))
        
        gstrSQL = "" & _
            "   Select id,编码,名称,简码 " & _
            "   From 供应商 " & _
            "   Where (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
            "       And (站点=[2] or 站点 is null) And 末级=1 And (substr(类型,5,1) = 1 ) " & _
            "       And (简码 like [1] Or 编码 like [1] or upper(名称) like [1]) "
        Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "供应商", False, "", "", False, False, True, _
                            vRect.Left, vRect.Top, txtProvider.Height, blnCancel, False, False, _
                            strProviderText, gstrNodeNo)
        If Not rsProvider Is Nothing Then
            txtProvider.Text = zlStr.Nvl(rsProvider!名称)
            txtProvider.Tag = zlStr.Nvl(rsProvider!Id)
            chkDate.SetFocus
        Else
            txtProvider.Text = ""
            txtProvider.Tag = "0"
        End If
    End With
End Sub

Private Sub txtProvider_LostFocus()
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub txtProvider_Validate(Cancel As Boolean)
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
    If CheckQualifications(MCON_模块号, 2, Val(txtProvider.Tag)) = False Then
        txtProvider.Text = ""
        txtProvider.Tag = "0"
        Exit Sub
    End If
    
End Sub

Private Sub SetVSFHead(ByVal vsfObject As VSFlexGrid, ByVal strHead As String)
    Dim arrCols As Variant, arrRows As Variant
    Dim i As Integer
    
    arrRows = Split(strHead, "|")
    With vsfObject
        If .Rows = 0 Then .Rows = 1
        .Cols = UBound(arrRows) + 1
        For i = LBound(arrRows) To UBound(arrRows)
            If arrRows(i) = "" Then
                .TextMatrix(0, i) = ""
            Else
                arrCols = Split(arrRows(i), ",")
                '第1元素：显示值
                .TextMatrix(0, i) = arrCols(0)
                '第2元素：Key值
                If arrCols(1) = "" Then
                    If Left(arrCols(0), 2) = "H_" Then
                        .ColKey(i) = Mid(arrCols(0), 3, Len(arrCols(0)))
                    Else
                        .ColKey(i) = arrCols(0)
                    End If
                Else
                    .ColKey(i) = arrCols(1)
                End If
                '第3元素：宽度
                .ColWidth(i) = Val(arrCols(2))
                'H_为隐藏列
                If Left(arrCols(0), 2) = "H_" Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                    '第4元素：显示格式
                    If UBound(arrCols) > 2 Then
                        If UCase(arrCols(3)) = "D" Then
                            .ColFormat(i) = "yyyy-mm-dd"
                            .ColAlignment(i) = flexAlignCenterCenter
                        ElseIf UCase(arrCols(3)) = "T" Then
                            .ColFormat(i) = "hh:mi:ss"
                            .ColAlignment(i) = flexAlignCenterCenter
                        ElseIf UCase(arrCols(3)) = "DT" Then
                            .ColFormat(i) = "yyyy-mm-dd hh:mi:ss"
                            .ColAlignment(i) = flexAlignCenterCenter
                        ElseIf UCase(arrCols(3)) = "N" Then
                            .ColAlignment(i) = flexAlignRightCenter
                        Else
                            .ColAlignment(i) = flexAlignLeftCenter
                        End If
                    Else
                        .ColAlignment(i) = flexAlignLeftCenter
                    End If
                End If
            End If
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
        Next
    End With
End Sub

Private Sub SelPatient(ByVal lsItem As ListItem)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lngCurRow As Long
    Dim dblCostAmount As Double
    Dim dbl成本金额 As Double
    Dim lng病人科室ID As Long
    Dim blnInfo As Boolean
    Dim strIVNO As String, strIVCode As String, strIVDate As String
        
    lng病人科室ID = Val(Mid(lsItem.Key, InStr(2, lsItem.Key, "_") + 1))
    On Error GoTo ErrHandle
    
    '填充 vsfMaterial
    gstrSQL = _
        "   SELECT distinct a.药品id 材料id, A.NO, A.序号, ('[' || D.编码 || ']' || D.名称) AS 卫材信息, D.规格, D.产地 as 原产地, A.产地," & _
        "          A.批号, Nvl(A.批次,0) 批次, to_char(A.生产日期,'yyyy-mm-dd') 生产日期, A.效期, A.灭菌日期, A.灭菌效期, " & _
        IIf(mintUnit = 1, "ltrim(rtrim(to_char(A.成本价 * c.换算系数, " & gOraFmt_Max.FM_成本价 & "))) as 成本价, ", "A.成本价, ") & _
        IIf(mintUnit = 1, "ltrim(rtrim(to_char(A.实际数量 / c.换算系数, " & gOraFmt_Max.FM_数量 & "))) as 实际数量, ", "A.实际数量, ") & _
        "         decode(nvl(A.发药方式,0),1,-1,1) * A.成本金额 AS 结算金额, Nvl(A.发药方式,0) 退货, " & _
        "         DECODE(A.扣率, NULL, 0, A.扣率) AS 扣率, " & _
        IIf(mintUnit = 1, "ltrim(rtrim(to_char(A.零售价 * c.换算系数, " & gOraFmt_Max.FM_零售价 & "))) as 零售价, ", "A.零售价, ") & _
        "         decode(nvl(A.发药方式,0),1,-1,1)*A.零售金额 as 零售金额, decode(nvl(A.发药方式,0),1,-1,1)* A.差价 差价, " & _
        "         decode(nvl(A.发药方式,0),1,-1,1)*to_number(A.用法," & gOraFmt_Max.FM_金额 & " )  as 零售差价, " & _
        "         a.供药单位id,a.注册证号,a.商品条码, a.摘要,A.填制人,A.填制日期,A.配药人 as 核查人,A.配药日期 as 核查日期," & _
        "         a.库房id, a.内部条码, a.费用id, b.病人科室ID, " & _
        IIf(mintMode = 1, "'' 发票号, '' 发票代码, '' 发票日期, '' 发票金额, ", "f.发票号,f.发票代码, f.发票日期, f.发票金额, ") & _
        IIf(mintUnit = 1, " c.换算系数 as 换算系数, ", "1 as 换算系数,") & _
        IIf(mintUnit = 1, " c.包装单位 as 单位, ", " d.计算单位 as 单位, ") & _
        "         decode(E.收发ID, null, '', E.科室 || ',' || nvl(E.病人姓名,'') || ',' || nvl(E.住院号,'') || ',' || nvl(E.床号,'') ) AS 高值材料 " & _
        "       FROM 药品收发记录 A, 住院费用记录 B, 材料特性 C, 收费项目目录 D, 收发记录补充信息 E, 应付记录 F " & _
        "       Where A.费用ID = B.ID And a.药品ID = c.材料ID And a.药品id = d.ID And A.ID = E.收发ID(+) And a.id=f.收发id(+) And A.费用ID > 0 And " & _
        "         a.供药单位id + 0 = [1] and a.库房id = [2] AND A.记录状态 = 1 And A.单据 = 15 AND B.病人ID + 0 = [3] And B.病人科室ID = [4] And " & _
        "         a.填制日期 between to_date('" & Format(mdatBegin, "yyyy-mm-dd 00:00:00") & "', 'yyyy-mm-dd hh24:mi:ss') And " & _
        "           to_date('" & Format(mdatEnd, "yyyy-mm-dd 23:59:59") & "', 'yyyy-mm-dd hh24:mi:ss') "
    If mbln需要核查 Then
        If mintMode = 1 Then
            gstrSQL = gstrSQL & " And A.审核日期 is null And A.配药日期 is null "
        Else
            gstrSQL = gstrSQL & " And A.审核日期 is null And A.配药日期 is not null "
        End If
    Else
        gstrSQL = gstrSQL & " And A.审核日期 is null "
    End If
    gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "住院费用记录", "门诊费用记录") & " ORDER BY NO, 序号 "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "病人材料明细", Val(lvwPatient.Tag), mlngStockID, Val(Mid(lsItem.Key, 2)), lng病人科室ID)
    With vsfMaterial
        blnInfo = rsTmp.RecordCount > 0
        dblCostAmount = 0
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            lngCurRow = .Rows - 1
            .TextMatrix(lngCurRow, .ColIndex("病人ID")) = Val(Mid(lsItem.Key, 2))
            .TextMatrix(lngCurRow, .ColIndex("科室ID")) = lng病人科室ID
            .TextMatrix(lngCurRow, .ColIndex("材料ID")) = zlStr.Nvl(rsTmp!材料ID)
            .TextMatrix(lngCurRow, .ColIndex("材料名称")) = zlStr.Nvl(rsTmp!卫材信息)
            .TextMatrix(lngCurRow, .ColIndex("规格")) = zlStr.Nvl(rsTmp!规格)
            .TextMatrix(lngCurRow, .ColIndex("数量")) = zlStr.Nvl(rsTmp!实际数量)
            .TextMatrix(lngCurRow, .ColIndex("成本价")) = zlStr.Nvl(rsTmp!成本价)
            .TextMatrix(lngCurRow, .ColIndex("成本金额")) = zlStr.Nvl(rsTmp!结算金额)
            .TextMatrix(lngCurRow, .ColIndex("发票金额")) = IIf(mintMode = 1, zlStr.Nvl(rsTmp!结算金额), zlStr.Nvl(rsTmp!发票金额))
            .TextMatrix(lngCurRow, .ColIndex("NO")) = zlStr.Nvl(rsTmp!NO)
            .TextMatrix(lngCurRow, .ColIndex("序号")) = zlStr.Nvl(rsTmp!序号)
            .TextMatrix(lngCurRow, .ColIndex("库房ID")) = zlStr.Nvl(rsTmp!库房ID)
            .TextMatrix(lngCurRow, .ColIndex("供药单位ID")) = zlStr.Nvl(rsTmp!供药单位ID)
            .TextMatrix(lngCurRow, .ColIndex("产地")) = zlStr.Nvl(rsTmp!产地)
            .TextMatrix(lngCurRow, .ColIndex("批号")) = zlStr.Nvl(rsTmp!批号)
            .TextMatrix(lngCurRow, .ColIndex("换算系数")) = zlStr.Nvl(rsTmp!换算系数)
            .TextMatrix(lngCurRow, .ColIndex("单位")) = zlStr.Nvl(rsTmp!单位)
            .TextMatrix(lngCurRow, .ColIndex("生产日期")) = zlStr.Nvl(rsTmp!生产日期)
            .TextMatrix(lngCurRow, .ColIndex("效期")) = zlStr.Nvl(rsTmp!效期)
            .TextMatrix(lngCurRow, .ColIndex("灭菌日期")) = zlStr.Nvl(rsTmp!灭菌日期)
            .TextMatrix(lngCurRow, .ColIndex("灭菌效期")) = zlStr.Nvl(rsTmp!灭菌效期)
            .TextMatrix(lngCurRow, .ColIndex("扣率")) = zlStr.Nvl(rsTmp!扣率)
            .TextMatrix(lngCurRow, .ColIndex("零售价")) = zlStr.Nvl(rsTmp!零售价)
            .TextMatrix(lngCurRow, .ColIndex("零售金额")) = zlStr.Nvl(rsTmp!零售金额)
            .TextMatrix(lngCurRow, .ColIndex("差价")) = zlStr.Nvl(rsTmp!差价)
            .TextMatrix(lngCurRow, .ColIndex("差价金额")) = zlStr.Nvl(rsTmp!零售差价)
            .TextMatrix(lngCurRow, .ColIndex("摘要")) = zlStr.Nvl(rsTmp!摘要)
            .TextMatrix(lngCurRow, .ColIndex("注册证号")) = zlStr.Nvl(rsTmp!注册证号)
            .TextMatrix(lngCurRow, .ColIndex("填制人")) = zlStr.Nvl(rsTmp!填制人)
            .TextMatrix(lngCurRow, .ColIndex("填制日期")) = zlStr.Nvl(rsTmp!填制日期)
            .TextMatrix(lngCurRow, .ColIndex("核查人")) = zlStr.Nvl(rsTmp!核查人)
            .TextMatrix(lngCurRow, .ColIndex("核查日期")) = zlStr.Nvl(rsTmp!核查日期)
            .TextMatrix(lngCurRow, .ColIndex("批次")) = zlStr.Nvl(rsTmp!批次)
            .TextMatrix(lngCurRow, .ColIndex("高值材料")) = zlStr.Nvl(rsTmp!高值材料)
            .TextMatrix(lngCurRow, .ColIndex("商品条码")) = zlStr.Nvl(rsTmp!商品条码)
            .TextMatrix(lngCurRow, .ColIndex("内部条码")) = zlStr.Nvl(rsTmp!内部条码)
            .TextMatrix(lngCurRow, .ColIndex("费用ID")) = zlStr.Nvl(rsTmp!费用ID)
            If mintMode = 1 Then
                dblCostAmount = dblCostAmount + zlStr.Nvl(rsTmp!结算金额, 0)
            Else
                dblCostAmount = dblCostAmount + zlStr.Nvl(rsTmp!发票金额, 0)
            End If
            dbl成本金额 = dbl成本金额 + Nvl(rsTmp!结算金额)
            If strIVNO = "" Then strIVNO = zlStr.Nvl(rsTmp!发票号)
            If strIVCode = "" Then strIVCode = zlStr.Nvl(rsTmp!发票代码)
            If strIVDate = "" Then strIVDate = IIf(IsNull(rsTmp!发票日期), "", Format(rsTmp!发票日期, "yyyy-mm-dd"))
            rsTmp.MoveNext
        Loop
        If .Rows > 1 Then
            .Row = 1
        End If
    End With
    rsTmp.Close

    If blnInfo Then
        '填充 vsfPatient
        With vsfPatient
            .Rows = .Rows + 1
            lngCurRow = .Rows - 1
            .TextMatrix(lngCurRow, .ColIndex("病人ID")) = Val(Mid(lsItem.Key, 2))
            .TextMatrix(lngCurRow, .ColIndex("科室ID")) = lng病人科室ID
            .TextMatrix(lngCurRow, .ColIndex("科室")) = lsItem.Text
            .TextMatrix(lngCurRow, .ColIndex("姓名")) = lsItem.SubItems(1)
            .TextMatrix(lngCurRow, .ColIndex("成本金额")) = dbl成本金额
            .TextMatrix(lngCurRow, .ColIndex("发票金额")) = dblCostAmount
            .TextMatrix(lngCurRow, .ColIndex("发票号")) = strIVNO
            .TextMatrix(lngCurRow, .ColIndex("发票代码")) = strIVCode
            .TextMatrix(lngCurRow, .ColIndex("发票日期")) = strIVDate
            .Row = lngCurRow
        End With
    End If
    
    zl_VsGridLOSTFOCUS vsfPatient
    zl_VsGridLOSTFOCUS vsfMaterial
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CalPatient(ByVal lsItem As ListItem) As Boolean
    Dim i As Long
    Dim lngPatientID As Long, lngPatientDrugID As Long
    
    '病人ID
    lngPatientID = Val(Mid(lsItem.Key, 2))
    '病人科室ID
    lngPatientDrugID = Val(Mid(lsItem.Key, InStr(2, lsItem.Key, "_") + 1))
    
    '检查是否录入数据，有就询问提示
    With vsfPatient
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("病人ID"))) = lngPatientID And Val(.TextMatrix(i, .ColIndex("科室ID"))) = lngPatientDrugID Then
                If Trim(.TextMatrix(i, .ColIndex("发票号"))) <> "" Or _
                   Trim(.TextMatrix(i, .ColIndex("发票代码"))) <> "" Or Trim(.TextMatrix(i, .ColIndex("发票日期"))) <> "" Then
                    If MsgBox("发现你已经有录入发票信息，程序将自动清除，要继续吗？", vbInformation + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
                        CalPatient = False
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    
    '清理选定的数据
    With vsfPatient
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, .ColIndex("病人ID"))) = lngPatientID And Val(.TextMatrix(i, .ColIndex("科室ID"))) = lngPatientDrugID Then
                .RemoveItem i
                Exit For
            End If
        Next
    End With
    With vsfMaterial
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, .ColIndex("病人ID"))) = lngPatientID And Val(.TextMatrix(i, .ColIndex("科室ID"))) = lngPatientDrugID Then
                .RemoveItem i
            End If
        Next
    End With
    
    CalPatient = True
    
End Function

Private Sub vsfMaterial_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim dblTmp As Double
    
    With vsfMaterial
        If .ColIndex("发票金额") = Col Then
            For i = 1 To .Rows - 1
                If .RowHidden(i) = False Then
                    dblTmp = dblTmp + Val(.TextMatrix(i, .ColIndex("发票金额")))
                End If
            Next
            vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("发票金额")) = Format(dblTmp, mFMT.FM_金额)
            Call ShowAmount
        End If
    End With
End Sub

Private Sub vsfMaterial_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsfMaterial.Rows > 1 Then Call zl_VsGridRowChange(vsfMaterial, OldRow, NewRow, OldCol, NewCol)
    vsfMaterial.Cell(flexcpBackColor, 0, 0, 0, vsfMaterial.Cols - 1) = &H8000000F
End Sub

Private Sub vsfMaterial_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfMaterial
        If .ColIndex("发票金额") = Col Then
'            If mintMode <> 1 Then
                Cancel = False
'            Else
'                Cancel = True
'            End If
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfMaterial_GotFocus()
    If vsfMaterial.Rows > 1 Then zl_VsGridGotFocus vsfMaterial
End Sub

Private Sub vsfMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlVsMoveGridCell vsfMaterial
    End If
End Sub

Private Sub vsfMaterial_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = vsfMaterial.ColIndex("发票金额") Then
        VsFlxGridCheckKeyPress vsfMaterial, Row, Col, KeyAscii, m金额式
    End If
End Sub

Private Sub vsfMaterial_LostFocus()
    zl_VsGridLOSTFOCUS vsfMaterial
End Sub

Private Sub vsfPatient_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim blnEdit As Boolean
    
    With vsfPatient
        If .ColIndex("发票号") = Col Then
            .TextMatrix(Row, Col) = UCase(.TextMatrix(Row, Col))
            blnEdit = (.TextMatrix(Row, Col) <> "")
        ElseIf .ColIndex("发票代码") = Col Then
            .TextMatrix(Row, Col) = UCase(.TextMatrix(Row, Col))
            blnEdit = (.TextMatrix(Row, Col) <> "")
        ElseIf .ColIndex("发票日期") = Col Then
            If Not IsDate(Trim(.Text)) And Trim(.Text) <> "" Then
                .Text = ""
                MsgBox "“生产日期”录入格式有误！", vbInformation, gstrSysName
                Exit Sub
            End If
            blnEdit = True
        ElseIf .ColIndex("发票金额") = Col Then
            mdblIVAmountNew = Val(.Text)
            ApportionInvoiceAmount mdblIVAmountOld, mdblIVAmountNew
            Call ShowAmount
            blnEdit = False
        End If
        
        If blnEdit Then
            For i = 1 To .Rows - 1
                If .RowHidden(i) = False And Trim(.TextMatrix(i, Col)) = "" And i <> Row Then
                    .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfPatient_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfPatient
        If .Rows > 1 Then
            Call zl_VsGridRowChange(vsfPatient, OldRow, NewRow, OldCol, NewCol)
            Call FillMaterial(Val(.TextMatrix(NewRow, .ColIndex("病人ID"))), Val(.TextMatrix(NewRow, .ColIndex("科室ID"))))
            Call ShowAmount
        End If
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = &H8000000F
    End With
    
End Sub

Private Sub vsfPatient_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPatient
        If .ColIndex("发票号") = Col Or .ColIndex("发票代码") = Col Or .ColIndex("发票日期") = Col Or .ColIndex("发票金额") = Col Then
'            If mintMode <> 1 Then
                Cancel = False
                mdblIVAmountOld = 0
                If .ColIndex("发票金额") = Col Then
                    mdblIVAmountOld = Val(.Text)
                End If
'            Else
'                Cancel = True
'            End If
        Else
            Cancel = True
        End If
        
        
    End With
End Sub

Private Sub vsfPatient_GotFocus()
    If vsfMaterial.Rows > 1 Then zl_VsGridGotFocus vsfPatient
End Sub

Private Sub vsfPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, lngPatient As Long, lngPatientDrug As Long
        
    If KeyCode = vbKeyReturn Then
        zlVsMoveGridCell vsfPatient
    ElseIf KeyCode = vbKeyDelete Then
        lngPatient = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("病人ID")))
        lngPatientDrug = Val(vsfPatient.TextMatrix(vsfPatient.Row, vsfPatient.ColIndex("科室ID")))
        
        With lvwPatient
            For i = 1 To .ListItems.Count
                If Val(Mid(.ListItems(i).Key, 2)) = lngPatient And Val(Mid(.ListItems(i).Key, InStr(2, .ListItems(i).Key, "_") + 1)) = lngPatientDrug Then
                    .ListItems(i).Checked = False
                    lvwPatient_ItemCheck .ListItems(i)
                End If
            Next
        End With
        With vsfPatient
            If .Rows > 1 Then
                Call FillMaterial(Val(.TextMatrix(.Row, .ColIndex("病人ID"))), Val(.TextMatrix(.Row, .ColIndex("科室ID"))))
                Call ShowAmount
            End If
        End With
    End If
End Sub

Private Sub vsfPatient_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = vsfPatient.ColIndex("发票金额") Then
        VsFlxGridCheckKeyPress vsfPatient, Row, Col, KeyAscii, m金额式
    End If
    
    
    If KeyAscii <> 13 Then
        If Col = vsfPatient.ColIndex("发票代码") Then
            If InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
                If Len(vsfPatient.EditText) <= 19 Or KeyAscii = 8 Then
                    KeyAscii = KeyAscii
                Else
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
        End If
    End If
    
End Sub

Private Sub vsfPatient_LostFocus()
    zl_VsGridLOSTFOCUS vsfPatient
End Sub

Private Sub ShowAmount()
    Dim dblCostAmount As Double, dblIVAmount As Double
    Dim dblCost As Double, dblIV As Double
    Dim i As Long, crFore1 As Long, crFore2 As Long
    
    With vsfPatient
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                dblCostAmount = dblCostAmount + Val(.TextMatrix(i, .ColIndex("成本金额")))
                dblIVAmount = dblIVAmount + Val(.TextMatrix(i, .ColIndex("发票金额")))
            End If
        Next
    End With
    With vsfMaterial
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                dblCost = dblCost + Val(.TextMatrix(i, .ColIndex("成本金额")))
                dblIV = dblIV + Val(.TextMatrix(i, .ColIndex("发票金额")))
            End If
        Next
    End With
    lblCostAmount.Caption = "合计成本金额：" & Format(dblCostAmount, "###,###,###,##0.000")
    lblIVAmount.Caption = "合计发票金额：" & Format(dblIVAmount, "###,###,###,##0.000")
    lblCost.Caption = "小计成本金额：" & Format(dblCost, "###,###,###,##0.000")
    lblIV.Caption = "小计发票金额：" & Format(dblIV, "###,###,###,##0.000")
    If dblCostAmount <> dblIVAmount Then
        crFore1 = vbRed
    Else
        crFore1 = vbBlack
    End If
    If dblCost <> dblIV Then
        crFore2 = vbRed
    Else
        crFore2 = vbBlack
    End If
    lblCostAmount.ForeColor = crFore1
    lblIVAmount.ForeColor = crFore1
    lblCost.ForeColor = crFore2
    lblIV.ForeColor = crFore2
End Sub

Private Sub ApportionInvoiceAmount(ByVal dblAmountOld As Double, ByVal dblAmountNew As Double)
'分摊发票金额到明细的发票金额
    Dim i As Long, LngLastRow As Long
    Dim dblTmp As Double
    
    With vsfMaterial
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                LngLastRow = i
                 .TextMatrix(i, .ColIndex("发票金额")) = Format(Val(.TextMatrix(i, .ColIndex("发票金额"))) / IIf(dblAmountOld = 0, 1, dblAmountOld) * dblAmountNew, mFMT.FM_金额)
            End If
        Next
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                If LngLastRow = i Then
                    dblTmp = dblAmountNew - dblTmp
                    .TextMatrix(i, .ColIndex("发票金额")) = Format(dblTmp, mFMT.FM_金额)
                    Exit For
                Else
                    dblTmp = dblTmp + Val(.TextMatrix(i, .ColIndex("发票金额")))
                End If
            End If
        Next
    End With
End Sub

Private Sub FillMaterial(ByVal lngPatientID As Long, ByVal lngPatientDrugID As Long)
'填充当前病人使用的材料信息
    Dim i As Long, lngTop As Long
    With vsfMaterial
        For i = 1 To .Rows - 1
            If lngPatientID = Val(.TextMatrix(i, .ColIndex("病人ID"))) And lngPatientDrugID = Val(.TextMatrix(i, .ColIndex("科室ID"))) Then
                .RowHidden(i) = False
                If lngTop = 0 Then lngTop = i
            Else
                .RowHidden(i) = True
            End If
        Next
        If lngTop > 0 Then .Row = lngTop
    End With
End Sub

Private Sub VerifyBatch()
'批量审核
    Dim i As Long, j As Long
    Dim strNo As String
    Dim strTime_Start As String, strTime_End As String
    Dim lngPatientID As Long, lngPatientDrugID As Long
    Dim strNewPirce As String
    Dim intCount As Integer
    
    If vsfPatient.Rows <= 1 Or vsfMaterial.Rows <= 1 Then Exit Sub
    
    If MsgBox("你确定批量“" & IIf(mintMode = 1, "核查", "审核") & "”吗？", vbInformation + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    'If mintMode <> 1 Then
        With vsfPatient
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("发票号"))) = "" Then
                    MsgBox "“发票号”未录入！", vbInformation, gstrSysName
                    .Col = .ColIndex("发票号")
                    .Row = i
                    .SetFocus
                    Exit Sub
                End If
                If Trim(.TextMatrix(i, .ColIndex("发票代码"))) = "" Then
                    MsgBox "“发票代码”未录入！", vbInformation, gstrSysName
                    .Col = .ColIndex("发票代码")
                    .Row = i
                    .SetFocus
                    Exit Sub
                End If
                If Trim(.TextMatrix(i, .ColIndex("发票日期"))) = "" Then
                    MsgBox "“发票日期”未录入！", vbInformation, gstrSysName
                    .Col = .ColIndex("发票日期")
                    .Row = i
                    .SetFocus
                    Exit Sub
                End If
                If Val(.TextMatrix(i, .ColIndex("发票金额"))) < 0 Then
                    MsgBox "“发票金额”未录入！", vbInformation, gstrSysName
                    .Col = .ColIndex("发票金额")
                    .Row = i
                    .SetFocus
                    Exit Sub
                End If
            Next
        End With
    'End If
    
    With vsfMaterial
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("发票金额"))) < 0 Then
                MsgBox "“发票金额”未录入！", vbInformation, gstrSysName
                .Col = .ColIndex("发票金额")
                .Row = i
                .SetFocus
                Exit Sub
            End If
        Next
    End With
    
    '检查价格变动，如果价格变动则更新界面数据
    If mblnUpdate = False Then
        strNo = ""
        For i = 1 To vsfMaterial.Rows - 1
            strNo = vsfMaterial.TextMatrix(i, vsfMaterial.ColIndex("NO"))
            If Not CheckValuePrice(15, strNo) = True Then
                intCount = intCount + 1
                If intCount <= 5 Then
                    strNewPirce = IIf(strNewPirce = "", "", strNewPirce & vbCrLf) & vsfMaterial.TextMatrix(i, vsfMaterial.ColIndex("材料名称"))
                End If
            End If
        Next
        
        If strNewPirce <> "" Then
            ShowMsgBox "高值卫材入库单中价格已调价，程序已自动完成更新（售价、售价金额、，成本价、成本金额、差价）,请检查！" & vbCrLf & strNewPirce
            mblnUpdate = True
            Call ShowAmount
            Exit Sub
        End If
    End If
    
    strNo = ""
    For i = 1 To vsfMaterial.Rows - 1
        If strNo = vsfMaterial.TextMatrix(i, vsfMaterial.ColIndex("NO")) Then GoTo Continue
        
        strNo = vsfMaterial.TextMatrix(i, vsfMaterial.ColIndex("NO"))
                
        gcnOracle.BeginTrans
        
        '保存单据
        If SaveCard(strNo) = False Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        
        If mbln需要核查 Then
            If mintMode = 1 Then
                strTime_Start = GetBillInfo(15, strNo, False, True)
                strTime_End = GetBillInfo(15, strNo, False, True)
                If strTime_Start = "" Then strTime_Start = GetBillInfo(15, strNo)
                If strTime_End = "" Then strTime_End = GetBillInfo(15, strNo)
            Else
                strTime_Start = GetBillInfo(15, strNo)
                strTime_End = GetBillInfo(15, strNo)
            End If
        Else
            strTime_Start = GetBillInfo(15, strNo)
            strTime_End = GetBillInfo(15, strNo)
        End If
        If strTime_End = "" Then
            gcnOracle.RollbackTrans
            MsgBox "注意:" & vbCrLf & "  该单据已经被其他操作员删除,不能继续！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mintMode <> 1 Then
            '审核单据
            If SaveCheck(strNo) = True Then
                gcnOracle.CommitTrans
                '标记审核完成的单据
                For j = 1 To vsfMaterial.Rows - 1
                    If vsfMaterial.TextMatrix(j, vsfMaterial.ColIndex("NO")) = strNo Then
                        vsfMaterial.TextMatrix(j, vsfMaterial.ColIndex("Verify")) = "1"
                    End If
                Next
            Else
                gcnOracle.RollbackTrans
            End If
        Else
            gcnOracle.CommitTrans
            '标记完成的单据
            For j = 1 To vsfMaterial.Rows - 1
                If vsfMaterial.TextMatrix(j, vsfMaterial.ColIndex("NO")) = strNo Then
                    vsfMaterial.TextMatrix(j, vsfMaterial.ColIndex("Verify")) = "1"
                End If
            Next
        End If
        
Continue:
    Next
    
    '清理界面数据
    With vsfMaterial
        lngPatientID = 0
        '清理已选定的病人材料
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, .ColIndex("Verify")) = "1" Then
                'If lngPatientID <> Val(.TextMatrix(i, .ColIndex("病人ID"))) Then
                    lngPatientID = Val(.TextMatrix(i, .ColIndex("病人ID")))
                    lngPatientDrugID = Val(.TextMatrix(i, .ColIndex("科室ID")))
                    '清理已选定的病人
                    For j = vsfPatient.Rows - 1 To 1 Step -1
                        If Val(vsfPatient.TextMatrix(j, vsfPatient.ColIndex("病人ID"))) = lngPatientID And Val(vsfPatient.TextMatrix(j, vsfPatient.ColIndex("科室ID"))) = lngPatientDrugID Then
                            vsfPatient.RemoveItem j
                            Exit For
                        End If
                    Next
                    '清理病人列表
                    For j = 1 To lvwPatient.ListItems.Count
                        If Val(Mid(lvwPatient.ListItems(j).Key, 2)) = lngPatientID And Val(Mid(lvwPatient.ListItems(j).Key, InStr(2, lvwPatient.ListItems(j).Key, "_") + 1)) = lngPatientDrugID Then
                            lvwPatient.ListItems.Remove j
                            Exit For
                        End If
                    Next
                'End If
                .RemoveItem i
            End If
        Next
    End With
End Sub

Private Function CheckValuePrice(ByVal int单据 As Integer, ByVal strNo As String) As Boolean
    '检查高值卫材虚拟入库产生的入库单的价格，有价格变动时更新界面价格，金额
    '查找在入库单填制日期后是否存在同批次的调价记录，如果有调价记录，找最近的调价记录和当前入库单的价格进行比较
    '只检查时价卫材的售价和成本价
    '返回：true-检查通过,false-有价格变动
    Dim rsData As ADODB.Recordset
    Dim rsprice As ADODB.Recordset
    Dim lng材料ID As Long
    Dim lng批次 As Long
    Dim str填制日期 As String
    Dim dbl原价 As Double
    Dim dbl现售价 As Double
    Dim dbl现成本价 As Double
    Dim strAdjustList As String '需要变动的清单：材料id,批次,现售价(为0表示价格无变化),现成本价(为0表示价格无变化)
    Dim lngRow As Long
    Dim lngRows As Long
    Dim dbl数量 As Double
    Dim dbl成本金额 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim lng科室id As Long
    Dim lng病人id As Long
    Dim bln成本价变动 As Boolean
    Dim dbl成本金额合计 As Double
    Dim blnUpdate As Boolean
    
    gstrSQL = " Select '定价售价' As 类型, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, a.零售价 As 原价, a.填制日期 " & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, 收费项目目录 C" & _
            " Where a.单据 = [1] And a.No = [2] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价,2) <> Round(b.现价, 2) And" & _
              "    NVL(c.是否变价, 0) = 0 " & _
        " Union All" & vbNewLine & _
        "Select '时价售价' As 类型, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, a.零售价 As 原价, a.填制日期 " & vbNewLine & _
        " From 药品收发记录 A, 收费项目目录 C" & vbNewLine & _
        " Where a.单据 = [1] And a.No = [2] And c.Id = a.药品id And Nvl(c.是否变价, 0) = 1 And a.费用id > 0 And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From 药品收发记录 B" & vbNewLine & _
        "       Where a.药品id = b.药品id And a.批次 = b.批次 And b.单据 = 13 And b.审核日期 > a.填制日期 And b.摘要 = '卫材调价')" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select '成本价' As 类型, a.药品id As 材料id, Nvl(a.批次, 0) As 批次, a.成本价 As 原价, a.填制日期 " & vbNewLine & _
        " From 药品收发记录 A" & vbNewLine & _
        " Where a.单据 = [1] And a.No = [2] And a.费用id > 0 And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From 药品收发记录 B" & vbNewLine & _
        "       Where a.药品id = b.药品id And a.批次 = b.批次 And b.单据 = 18 And b.审核日期 > a.填制日期 And b.摘要 = '卫生材料成本价调价') "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", int单据, strNo)
        
    If rsData.RecordCount = 0 Then
        CheckValuePrice = True
        Exit Function
    End If
    
    '检查到有调价记录则比较价格，由于调价记录可能有多条，取最近一条价格来比较
    Do While Not rsData.EOF
        lng材料ID = rsData!材料ID
        lng批次 = rsData!批次
        str填制日期 = Format(rsData!填制日期, "yyyy-mm-dd hh:mm:ss")
        dbl原价 = rsData!原价
        
        dbl现售价 = 0
        dbl现成本价 = 0
        bln成本价变动 = False
        
        If rsData!类型 = "定价售价" Then
            gstrSQL = " Select nvl(现价,0) 现价 From 收费价目 " & _
            " Where 收费细目id=[1] and (终止日期 Is NULL Or sysdate Between 执行日期 And nvl(终止日期,to_date('3000-01-01','yyyy-MM-dd')))" & _
            GetPriceClassString("")
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng材料ID)
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!现价, 2) <> Round(dbl原价, 2) Then
                    dbl现售价 = rsprice!现价
                    blnUpdate = True
                End If
            End If
        End If
        
        If rsData!类型 = "时价售价" Then
            gstrSQL = "Select 零售价 As 现价 " & _
                " From 药品收发记录 " & _
                " Where ID = (Select Max(ID) " & _
                " From 药品收发记录 B " & _
                " Where b.药品id = [1] And b.批次 = [2] And b.单据 = 13 And b.审核日期 > [3] And b.摘要 = '卫材调价') "
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng材料ID, lng批次, CDate(str填制日期))
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!现价, 2) <> Round(dbl原价, 2) Then
                    dbl现售价 = rsprice!现价
                    blnUpdate = True
                End If
            End If
        End If
        
        If rsData!类型 = "成本价" Then
            gstrSQL = "Select 单量 As 现价 " & _
                " From 药品收发记录 " & _
                " Where ID = (Select Max(ID) " & _
                " From 药品收发记录 B " & _
                " Where b.药品id = [1] And b.批次 = [2] And b.单据 = 18 And b.审核日期 > [3] And b.摘要 = '卫生材料成本价调价') "
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, "CheckValuePrice", lng材料ID, lng批次, CDate(str填制日期))
            
            If rsprice.RecordCount > 0 Then
                If Round(rsprice!现价, 2) <> Round(dbl原价, 2) Then
                    dbl现成本价 = rsprice!现价
                    bln成本价变动 = True
                    blnUpdate = True
                End If
            End If
        End If
        
        '以当前最新价格最新单据相关数据（单价、零售金额、差价）
        With vsfMaterial
            lngRows = vsfMaterial.Rows - 1
            For lngRow = 1 To lngRows
                If strNo = .TextMatrix(lngRow, .ColIndex("NO")) And lng材料ID = Val(.TextMatrix(lngRow, .ColIndex("材料ID"))) And (dbl现售价 <> 0 Or dbl现成本价 <> 0) Then
                    dbl数量 = Val(.TextMatrix(lngRow, .ColIndex("数量")))
                    lng科室id = Val(.TextMatrix(lngRow, .ColIndex("科室ID")))
                    lng病人id = Val(.TextMatrix(lngRow, .ColIndex("病人ID")))
                    If dbl现售价 <> 0 Then
                        dbl现售价 = Val(Format(dbl现售价 * Val(.TextMatrix(lngRow, .ColIndex("换算系数"))), mFMT.FM_零售价))
                        dbl零售金额 = dbl现售价 * dbl数量
                    Else
                        dbl现售价 = Val(.TextMatrix(lngRow, .ColIndex("零售价")))
                        dbl零售金额 = Val(.TextMatrix(lngRow, .ColIndex("零售金额")))
                    End If
                    
                    If dbl现成本价 <> 0 Then
                        dbl现成本价 = Val(Format(dbl现成本价 * Val(.TextMatrix(lngRow, .ColIndex("换算系数"))), mFMT.FM_成本价))
                        dbl成本金额 = dbl现成本价 * dbl数量
                    Else
                        dbl现成本价 = Val(.TextMatrix(lngRow, .ColIndex("成本价")))
                        dbl成本金额 = Val(.TextMatrix(lngRow, .ColIndex("成本金额")))
                    End If
                    
                    dbl差价 = dbl零售金额 - dbl成本金额
                    
                    .TextMatrix(lngRow, .ColIndex("成本价")) = Format(dbl现成本价, mFMT.FM_成本价)
                    .TextMatrix(lngRow, .ColIndex("成本金额")) = Format(dbl成本金额, mFMT.FM_金额)
                    .TextMatrix(lngRow, .ColIndex("零售价")) = Format(dbl现售价, mFMT.FM_零售价)
                    .TextMatrix(lngRow, .ColIndex("零售金额")) = Format(dbl零售金额, mFMT.FM_金额)
                    .TextMatrix(lngRow, .ColIndex("差价")) = Format(dbl差价, mFMT.FM_金额)

                End If
            Next
                    
            dbl成本金额合计 = 0
            If bln成本价变动 = True Then
                For lngRow = 1 To lngRows
                    If lng科室id = Val(.TextMatrix(lngRow, .ColIndex("科室ID"))) And lng病人id = Val(.TextMatrix(lngRow, .ColIndex("病人ID"))) Then
                        dbl成本金额合计 = dbl成本金额合计 + .TextMatrix(lngRow, .ColIndex("成本金额"))
                    End If
                Next
                
                '更新病人列表成本金额
                lngRows = vsfPatient.Rows - 1
                For lngRow = 1 To lngRows
                    If lng科室id = Val(vsfPatient.TextMatrix(lngRow, vsfPatient.ColIndex("科室ID"))) And lng病人id = Val(vsfPatient.TextMatrix(lngRow, vsfPatient.ColIndex("病人ID"))) Then
                        vsfPatient.TextMatrix(lngRow, vsfPatient.ColIndex("成本金额")) = dbl成本金额合计
                        vsfPatient.Row = lngRow
                        vsfPatient.TopRow = lngRow
                    End If
                Next
            End If
        End With
        
        rsData.MoveNext
    Loop
    
    CheckValuePrice = Not blnUpdate
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function SaveCheck(Optional ByVal strNo As String = "") As Boolean
   
    gstrSQL = "zl_材料外购_Verify('" & strNo & "','" & UserInfo.用户名 & "')"
    
    On Error GoTo ErrHandle
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-审核")
    
    SaveCheck = True
    Exit Function

ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function SaveCard(ByVal strNo As String) As Boolean
    Dim lng序号 As Long
    Dim lngStockID As Long
    Dim lng供货单位id As Long
    Dim lng材料ID As Long
    Dim str批号 As String
    Dim str产地 As String
    Dim str效期 As String
    Dim dbl实际数量 As Double
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl扣率 As Double
    Dim dbl零售价 As Double
    Dim dbl零售金额 As Double
    Dim dbl差价 As Double
    Dim dbl零售差价 As Double
    Dim str摘要 As String
    Dim str填制人 As String
    Dim str填制日期 As String
    Dim str审核人 As String
    Dim datAssessDate As String
    Dim str发票号 As String
    Dim str发票代码 As String
    Dim str发票日期 As String
    Dim str灭菌日期 As String
    Dim str灭菌失效期 As String
    Dim dbl发票金额 As Double
    Dim str生产日期  As String
    Dim str核查人 As String
    Dim str核查日期 As String
    Dim str注册证号 As String
    Dim intUnit As Integer
    Dim strUnit As String
    Dim str指导批发价 As String
    Dim str随货单号 As String
    Dim str商品条码 As String
    Dim str内部条码 As String
    Dim lng费用ID As Long
    Dim intRow As Integer
    Dim str高值材料 As String
    Dim str批次 As String
    Dim strTmp As String
    
    
    SaveCard = False
    
    With vsfMaterial
        
        lngStockID = mlngStockID
        lng供货单位id = lvwPatient.Tag
        
        On Error GoTo ErrHandle
        
        gstrSQL = "zl_材料外购_Delete('" & strNo & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("NO")) <> strNo Then GoTo Continue
            
            lng序号 = lng序号 + 1
            str摘要 = Trim(.TextMatrix(intRow, .ColIndex("摘要")))
            str填制人 = Trim(.TextMatrix(intRow, .ColIndex("填制人")))
            str审核人 = UserInfo.姓名
            str填制日期 = Trim(.TextMatrix(intRow, .ColIndex("填制日期")))
            
            If mbln需要核查 Then
                If mintMode = 1 Then
                    str核查人 = UserInfo.姓名
                    str核查日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                Else
                    str核查人 = Trim(.TextMatrix(intRow, .ColIndex("核查人")))
                    str核查日期 = Trim(.TextMatrix(intRow, .ColIndex("核查日期")))
                    If str核查人 = "" Then
                        str核查人 = UserInfo.姓名
                    End If
                    If str核查日期 = "" Then
                        str核查日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    End If
                End If
            Else
                str核查人 = Trim(.TextMatrix(intRow, .ColIndex("核查人")))
                str核查日期 = Trim(.TextMatrix(intRow, .ColIndex("核查日期")))
                If str核查人 = "" Then
                    str核查人 = UserInfo.姓名
                End If
                If str核查日期 = "" Then
                    str核查日期 = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
            End If
            
            lng材料ID = .TextMatrix(intRow, .ColIndex("材料ID"))
            str产地 = .TextMatrix(intRow, .ColIndex("产地"))
            str批号 = .TextMatrix(intRow, .ColIndex("批号"))
            str效期 = .TextMatrix(intRow, .ColIndex("效期"))
                
            strTmp = Val(.TextMatrix(intRow, .ColIndex("数量"))) * Val(.TextMatrix(intRow, .ColIndex("换算系数")))
            dbl实际数量 = Format(strTmp, mFMT.FM_数量)
            
            strTmp = Val(.TextMatrix(intRow, .ColIndex("成本价"))) / Val(.TextMatrix(intRow, .ColIndex("换算系数")))
            dbl成本价 = Round(Val(strTmp), g_小数位数.obj_散装小数.成本价小数)
            
            strTmp = Val(.TextMatrix(intRow, .ColIndex("零售价"))) / Val(.TextMatrix(intRow, .ColIndex("换算系数")))
            dbl零售价 = Round(Val(strTmp), g_小数位数.obj_散装小数.零售价小数)
            
            dbl扣率 = Val(.TextMatrix(intRow, .ColIndex("扣率")))
            dbl成本金额 = Val(.TextMatrix(intRow, .ColIndex("成本金额")))
            dbl零售金额 = Val(.TextMatrix(intRow, .ColIndex("零售金额")))
            dbl差价 = Val(.TextMatrix(intRow, .ColIndex("差价")))
            dbl零售差价 = Val(.TextMatrix(intRow, .ColIndex("差价金额")))
            str随货单号 = ""
                
            If GetInvoiceInfo(.TextMatrix(intRow, .ColIndex("病人ID")), str发票号, str发票代码, str发票日期) = False Then
                Exit Function
            End If
            
            dbl发票金额 = Val(.TextMatrix(intRow, .ColIndex("发票金额")))
            str灭菌日期 = Trim(IIf(.TextMatrix(intRow, .ColIndex("灭菌日期")) = "", "", .TextMatrix(intRow, .ColIndex("灭菌日期"))))
            str灭菌失效期 = Trim(IIf(.TextMatrix(intRow, .ColIndex("灭菌效期")) = "", "", .TextMatrix(intRow, .ColIndex("灭菌效期"))))
            str生产日期 = Trim(IIf(.TextMatrix(intRow, .ColIndex("生产日期")) = "", "", .TextMatrix(intRow, .ColIndex("生产日期"))))
            str注册证号 = Trim(.TextMatrix(intRow, .ColIndex("注册证号")))
            str内部条码 = Trim(.TextMatrix(intRow, .ColIndex("内部条码")))
            str商品条码 = Trim(.TextMatrix(intRow, .ColIndex("商品条码")))
            lng费用ID = Val(.TextMatrix(intRow, .ColIndex("费用ID")))
            str高值材料 = Trim(.TextMatrix(intRow, .ColIndex("高值材料")))
            str批次 = Trim(.TextMatrix(intRow, .ColIndex("批次")))
                
            ' Zl_材料外购_Insert
            gstrSQL = "zl_材料外购_INSERT("
            '  No_In         In 药品收发记录.NO%Type,
            gstrSQL = gstrSQL & "'" & strNo & "',"
            '  序号_In       In 药品收发记录.序号%Type,
            gstrSQL = gstrSQL & "" & lng序号 & ","
            '  库房id_In     In 药品收发记录.库房id%Type,
            gstrSQL = gstrSQL & "" & lngStockID & ","
            '  供药单位id_In In 药品收发记录.供药单位id%Type,
            gstrSQL = gstrSQL & "" & lng供货单位id & ","
            '  材料id_In     In 药品收发记录.药品id%Type,
            gstrSQL = gstrSQL & "" & lng材料ID & ","
            '  产地_In       In 药品收发记录.产地%Type := Null,
            gstrSQL = gstrSQL & "'" & str产地 & "',"
            '  批号_In       In 药品收发记录.批号%Type := Null,
            gstrSQL = gstrSQL & "'" & str批号 & "',"
            '  生产日期_In   In 药品收发记录.生产日期%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str生产日期 = "", "Null", "to_date('" & Format(str生产日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  效期_In       In 药品收发记录.效期%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str灭菌日期 = "", "Null", "to_date('" & Format(str灭菌日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str灭菌失效期 = "", "Null", "to_date('" & Format(str灭菌失效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  实际数量_In   In 药品收发记录.实际数量%Type := Null,
            gstrSQL = gstrSQL & "" & dbl实际数量 & ","
            '  成本价_In     In 药品收发记录.成本价%Type := Null,
            gstrSQL = gstrSQL & "" & dbl成本价 & ","
            '  成本金额_In   In 药品收发记录.成本金额%Type := Null,
            gstrSQL = gstrSQL & "" & dbl成本金额 & ","
            '  扣率_In       In 药品收发记录.扣率%Type := Null,
            gstrSQL = gstrSQL & "" & dbl扣率 & ","
            '  零售价_In     In 药品收发记录.零售价%Type := Null,
            gstrSQL = gstrSQL & "" & dbl零售价 & ","
            '  零售金额_In   In 药品收发记录.零售金额%Type := Null,
            gstrSQL = gstrSQL & "" & dbl零售金额 & ","
            '  差价_In       In 药品收发记录.差价%Type := Null,
            gstrSQL = gstrSQL & "" & dbl差价 & ","
            '  零售差价_In   In 药品收发记录.差价%Type := Null,目前存放在用法字段
            gstrSQL = gstrSQL & "" & dbl零售差价 & ","
            '  摘要_In       In 药品收发记录.摘要%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str摘要 = "", "NULL", "'" & str摘要 & "'") & ","
            '  注册证号_In   In 药品收发记录.注册证号%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str注册证号 = "", "NULL", "'" & str注册证号 & "'") & ","
            '  填制人_In     In 药品收发记录.填制人%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str填制人 = "", "NULL", "'" & str填制人 & "'") & ","
            '  随货单号_In   In 应付记录.随货单号%Type := Null
            gstrSQL = gstrSQL & "" & IIf(str随货单号 = "", "NULL", "'" & str随货单号 & "'") & ","
            '  发票号_In     In 应付记录.发票号%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str发票号 = "", "NULL", "'" & str发票号 & "'") & ","
            '  发票日期_In   In 应付记录.发票日期%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
            '  发票金额_In   In 应付记录.发票金额%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(dbl发票金额 = 0, "Null", dbl发票金额) & ","
            '  填制日期_In   In 药品收发记录.填制日期%Type := Null,
            gstrSQL = gstrSQL & "to_date('" & str填制日期 & "','yyyy-mm-dd HH24:MI:SS'),"
            '  核查人_In     In 药品收发记录.配药人%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str核查人 = "", "NULL", "'" & str核查人 & "'") & ","
            '  核查日期_In   In 药品收发记录.配药日期%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(str核查日期 = "", "Null", "to_date('" & str核查日期 & "','yyyy-mm-dd hh24:mi:ss')") & ","
            '  批次_In       In 药品收发记录.批次%Type := 0,
            gstrSQL = gstrSQL & "" & IIf(str批次 = "", "Null", "'" & str批次 & "'") & ","
            '  退货_In       In Number := 1
            gstrSQL = gstrSQL & "1,"
            '  高值材料_In   In varchar2(250)
            gstrSQL = gstrSQL & "" & IIf(str高值材料 = "", "Null", "'" & str高值材料 & "'") & ","
            '  商品条码_In   In 药品收发记录.商品条码%Type :=Null
            gstrSQL = gstrSQL & "" & IIf(str商品条码 = "", "NULL", "'" & str商品条码 & "'") & ","
            '  内部条码
            gstrSQL = gstrSQL & IIf(str内部条码 = "", "Null", "'" & str内部条码 & "'") & ","
            '  费用ID
            gstrSQL = gstrSQL & IIf(lng费用ID = 0, "Null", lng费用ID) & ","
            '  发票代码
            gstrSQL = gstrSQL & "" & IIf(str发票代码 = "", "NULL", "'" & str发票代码 & "'")
            gstrSQL = gstrSQL & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
Continue:
        Next
        
    End With
    SaveCard = True
    Exit Function
    
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function GetInvoiceInfo(ByVal lngPatientID, ByRef strIVNO As String, ByRef strIVCode As String, ByRef strIVDate As String) As Boolean
'获取发票号
    Dim i As Long
    
    With vsfPatient
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("病人ID")) = lngPatientID Then
                strIVNO = .TextMatrix(i, .ColIndex("发票号"))
                strIVCode = .TextMatrix(i, .ColIndex("发票代码"))
                strIVDate = .TextMatrix(i, .ColIndex("发票日期"))
                'dblIVAmount = .TextMatrix(i, .ColIndex("发票金额"))
                GetInvoiceInfo = True
                Exit Function
            End If
        Next
        strIVNO = ""
        strIVDate = ""
        'dblIVAmount = 0
    End With
End Function


