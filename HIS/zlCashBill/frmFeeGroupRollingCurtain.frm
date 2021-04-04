VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFeeGroupRollingCurtain 
   BorderStyle     =   0  'None
   Caption         =   "财务组收款管理"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picCurrentMoney 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      ScaleHeight     =   225
      ScaleWidth      =   2265
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
      Begin VB.Label lblCurrentMoney 
         Appearance      =   0  'Flat
         Caption         =   "当前暂存金: "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   8355
      End
   End
   Begin VB.PictureBox picSendFeeDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   2280
      ScaleHeight     =   2295
      ScaleWidth      =   6375
      TabIndex        =   9
      Top             =   5160
      Width           =   6375
      Begin VB.PictureBox picImgPlanSub 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   13
         Top             =   30
         Width           =   210
         Begin VB.Image imgColPlanSub 
            Height          =   195
            Left            =   0
            Picture         =   "frmFeeGroupRollingCurtain.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsSubCollectorInfo 
         Height          =   1695
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   5535
         _cx             =   9763
         _cy             =   2990
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeGroupRollingCurtain.frx":054E
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
         ExplorerBar     =   5
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
   Begin VB.PictureBox picTabSendFee 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4920
      ScaleHeight     =   1575
      ScaleWidth      =   5175
      TabIndex        =   7
      Top             =   480
      Width           =   5175
      Begin XtremeSuiteControls.TabControl tabSubSendFee 
         Height          =   1935
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2295
         _Version        =   589884
         _ExtentX        =   4048
         _ExtentY        =   3413
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picLastTime 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   1320
      ScaleHeight     =   2775
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   2640
      Width           =   6975
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   14
         Top             =   510
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmFeeGroupRollingCurtain.frx":076F
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.CommandButton cmdSendFees 
         Caption         =   "轧帐(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   4
         Top             =   75
         Width           =   1300
      End
      Begin VB.CommandButton cmdReloadData 
         Caption         =   "重新提取轧账数据(&G)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   75
         Width           =   2355
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCollectHistory 
         Height          =   2055
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   3255
         _cx             =   5741
         _cy             =   3625
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeGroupRollingCurtain.frx":0CBD
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
         ExplorerBar     =   5
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
      Begin MSComCtl2.DTPicker dtpLastTime 
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Top             =   75
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116719617
         CurrentDate     =   41521
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   300
         Left            =   5160
         TabIndex        =   2
         Top             =   75
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116719617
         CurrentDate     =   41521
      End
      Begin VB.Label lblLastTime 
         AutoSize        =   -1  'True
         Caption         =   "上次轧帐时间                           截止时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4935
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1320
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpSendFees 
      Bindings        =   "frmFeeGroupRollingCurtain.frx":0E4D
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFeeGroupRollingCurtain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjChargeBillRC As New clsChargeBill, mfrmChargeBillTotalRC As Form    '收款信息和票据对象
Private mlngModule As Long, mstrPrivs As String
Private mlngGroupID As Long '缴款组ID
Private mcbrPopupSub As CommandBar

Private Enum EM_Tab
    EM_Tab_收款 = 1
    EM_Tab_轧帐 = 2
    EM_Tab_历史轧账信息 = 3
    EM_Tab_收款及票据汇总 = 4
    EM_Tab_收费员轧帐明细 = 5
    EM_Tab_组收款信息 = 6
    EM_Tab_收费员轧帐信息 = 7
End Enum

Private Sub dkpSendFees_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionAttaching Then Cancel = True
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dtpEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtpLastTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘尔旋
    '日期:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
        
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    '初始化设置
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    
    Set mcbrPopupSub = cbsThis.Add("弹出菜单2", xtpBarPopup)
    With mcbrPopupSub.Controls
        .Add xtpControlButton, conMenu_View_Detail, "显示明细"
    End With
    
    cbsThis.ActiveMenuBar.Visible = False
    
    zlDefCommandBars = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub ViewDetail()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:查看明细按钮操作
    '编制:刘尔旋
    '日期:2013-09-22
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim i As Integer, strIDs As String
    If ActiveControl = vsSubCollectorInfo Then
        With vsSubCollectorInfo
            For i = .Row To .RowSel
                strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            Next i
            strIDs = Mid(strIDs, 2)
            Call mobjChargeBillRC.ChargeRollingListShow(Me, EM_收费员轧帐, strIDs, mlngModule, mstrPrivs)
        End With
    Else
        With vsCollectHistory
            For i = .Row To .RowSel
                strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            Next i
            strIDs = Mid(strIDs, 2)
            Call mobjChargeBillRC.ChargeRollingListShow(Me, EM_小组收款, strIDs, mlngModule, mstrPrivs)
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub InitMe(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngGroupID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:初始化轧帐界面
    '入参:lngModule-模块号
    '     strPrivs-权限串
    '编制:刘尔旋
    '日期:2013-10-10
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngGroupID = lngGroupID
End Sub

Private Sub SetDockingPanel()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建DOCKINGPANEL控件
    '编制:刘尔旋
    '日期:2013-09-04
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    
    With dkpSendFees
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(1, 1000, 1000, DockTopOf)
        objPanel.Handle = picLastTime.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 150
        Set objPanel = .CreatePane(2, 1000, 1000, DockBottomOf, objPanel)
        objPanel.Handle = picTabSendFee.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 250
        Set objPanel = .CreatePane(3, 1000, 300, DockBottomOf, objPanel)
        objPanel.Handle = picCurrentMoney.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 35
        objPanel.MaxTrackSize.Height = 35
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_View_Detail
            Call ViewDetail
    End Select
End Sub

Private Sub cmdReloadData_Click()
    Call SetDefaultRollingCurtain(True)
End Sub

Private Sub cmdSendFees_Click()
    Call RollingCurtain
End Sub

Private Sub SaveRollingCurtain()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:进行扎帐操作
    '编制:刘尔旋
    '日期:2013-09-10
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNO As String, strSQL As String, strIDs As String, i As Integer, lngID As Long
    Dim strTemp As String, blnBatch As Boolean, colSql As New Collection, strFixedSql As String
    blnBatch = False
    
    '获取单据号与ID
    strNO = zlDatabase.GetNextNo(139)
    lngID = zlDatabase.GetNextId("人员收缴记录")
    strFixedSql = "Zl_小组轧帐记录_Insert(" & lngID & ",'" & strNO & "'," & mlngGroupID & "," & _
                  "to_date('" & dtpLastTime.Value & "','yyyy-MM-dd HH24:mi:ss'),to_date('" & dtpEndTime.Value & "','yyyy-MM-dd HH24:mi:ss'),'" & _
                   UserInfo.姓名 & "'," & UserInfo.ID & ",to_date('" & zlDatabase.Currentdate & "','yyyy-MM-dd HH24:mi:ss'),'"
    
    With vsCollectHistory
        For i = 1 To .Rows - 1
            strTemp = strIDs
            strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            If Len(strIDs) >= 4000 Then
                strTemp = Mid(strTemp, 2)
                If blnBatch = False Then
                    strSQL = strFixedSql & strTemp & "',1)"
                Else
                    strSQL = strFixedSql & strTemp & "',2)"
                End If
                blnBatch = True
                colSql.Add strSQL
                strIDs = "," & Val(.TextMatrix(i, .ColIndex("ID")))
            End If
        Next i
    End With
    
    strIDs = Mid(strIDs, 2)
    If strIDs <> "" Then
        If blnBatch = False Then
            strSQL = strFixedSql & strIDs & "',0)"
        Else
            strSQL = strFixedSql & strIDs & "',3)"
        End If
        colSql.Add strSQL
    End If
    
    On Error GoTo errSql
    Call zlExecuteProcedureArrAy(colSql, Me.Caption)
    Call frmFeeGroupManage.AutoPrint(lngID, strNO, 2)
    Exit Sub
errSql:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub SetDefaultRollingCurtain(ByVal blnReload As Boolean, Optional ByVal blnUpdateEndTime As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:设置默认轧帐界面信息
    '编制:刘尔旋
    '日期:2013-09-10
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strSQL As String, rsTmp As New ADODB.Recordset, i As Integer, strDate As String
    Dim strEndDate As String
    If blnReload = True Then GoTo Reload
    strDate = ""
    
    If strDate = "" Then
        strSQL = "Select 上次轧帐时间 From 财务组组长构成 Where 组Id= [1] And 组长Id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, UserInfo.ID)
        If rsTmp.RecordCount <> 0 Then
            strDate = Nvl(rsTmp!上次轧帐时间)
        End If
    End If
    If strDate = "" Then
        strSQL = "Select 上次轧帐时间 From 财务缴款分组 Where Id= [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
        If rsTmp.RecordCount <> 0 Then
            strDate = Nvl(rsTmp!上次轧帐时间)
        End If
    End If
    If strDate = "" Then
        strSQL = "Select 终止时间 From 人员收缴记录 Where 记录性质=3 And 作废时间 Is Null And 缴款组ID= [1] Order By 终止时间 desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
        If rsTmp.RecordCount <> 0 Then
            strDate = Nvl(rsTmp!终止时间)
        End If
    End If
    If strDate = "" Then
        strSQL = "Select 登记时间 From 人员收缴记录 Where 记录性质=2 And 作废时间 Is Null And 缴款组ID= [1] Order By 登记时间 asc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
        If rsTmp.RecordCount <> 0 Then
            strDate = Nvl(rsTmp!登记时间)
        End If
    End If
    
    If strDate = "" Then
        dtpLastTime.Enabled = True
        strDate = Format(DateAdd("d", -7, zlDatabase.Currentdate), "yyyy-mm-dd HH:MM:SS")
    End If
    dtpLastTime.Value = strDate
Reload:
    With vsCollectHistory
        .Rows = 1
        strEndDate = zlDatabase.Currentdate
        dtpEndTime.MaxDate = strEndDate
        If blnUpdateEndTime Then dtpEndTime.Value = strEndDate
        If CStr(dtpLastTime.Value) <> "" Then
            strSQL = "" & _
            "Select NO, 收款员, 登记时间, Trim(to_char(冲预交款,'99999999990.00')) As 冲预交款, Trim(to_char(借入合计,'99999999990.00')) As 借入合计, " & _
            "Trim(to_char(借出合计,'99999999990.00')) As 借出合计, 小组收款人, 小组收款时间, 摘要, ID" & vbNewLine & _
            "From 人员收缴记录" & vbNewLine & _
            "Where 记录性质 = 2 And 小组收款人 = [1] And 作废时间 Is Null" & vbNewLine & _
            "      And 小组收款时间 Between [2] And [3] And 小组轧账ID Is Null And 缴款组ID = [4] " & vbNewLine & _
            "Order By 登记时间,NO Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, CDate(dtpLastTime.Value), CDate(dtpEndTime.Value), mlngGroupID)
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("收款单号")) = Nvl(rsTmp!NO)
                .TextMatrix(.Rows - 1, .ColIndex("收款时间")) = Nvl(rsTmp!登记时间)
                .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = Nvl(rsTmp!冲预交款)
                .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = Nvl(rsTmp!借入合计)
                .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = Nvl(rsTmp!借出合计)
                .TextMatrix(.Rows - 1, .ColIndex("缴款人")) = Nvl(rsTmp!收款员)
                .TextMatrix(.Rows - 1, .ColIndex("小组收款人")) = Nvl(rsTmp!小组收款人)
                .TextMatrix(.Rows - 1, .ColIndex("小组收款时间")) = Nvl(rsTmp!小组收款时间)
                .TextMatrix(.Rows - 1, .ColIndex("备注")) = Nvl(rsTmp!摘要)
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTmp!ID)
                rsTmp.MoveNext
            Loop
            'Set .DataSource = rsTmp
            .AutoSize 1, .Cols - 1
            zl_vsGrid_Para_Restore mlngModule, vsCollectHistory, Me.Caption, "小组收款信息", False
        End If
        If .Rows = 1 Then .Rows = 2
    End With
    With vsSubCollectorInfo
        .Clear 1
        .Rows = 2
    End With
    With vsCollectHistory
        mobjChargeBillRC.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_小组收款, 0
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CancelCollect()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:收款作废操作
    '编制:刘尔旋
    '日期:2013-09-10
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNO As String, strSQL As String
    With vsCollectHistory
        strSQL = "Zl_小组收款记录_Cancel(" & Val(.TextMatrix(.RowSel, .ColIndex("ID"))) & ",'" & UserInfo.姓名 & _
                 "',to_date('" & zlDatabase.Currentdate & "','yyyy-MM-dd HH24:mi:ss'))"
    End With
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshCurrentMoney(ByVal intPanel As Integer)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:刷新界面暂存金
    '入参:intPanel-TAB界面序号
    '编制:刘尔旋
    '日期:2013-09-18
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 结算方式,余额 From 人员缴款余额 Where 收款员=[1] And 性质=1"
    If intPanel = 1 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名)
    End If
    lblCurrentMoney(intPanel).Caption = " 当前暂存金:   "
    If rsTmp.RecordCount <> 0 Then
        Do While Not rsTmp.EOF
            If Val(Nvl(rsTmp!余额)) <> 0 Then
                lblCurrentMoney(intPanel).Caption = lblCurrentMoney(intPanel).Caption & rsTmp!结算方式 & ":" & rsTmp!余额 & "元   "
            End If
            rsTmp.MoveNext
        Loop
    End If
    If intPanel = 1 Then
        vsCollectHistory.Select 0, 0
        vsSubCollectorInfo.Rows = 2
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub ButtonCancelCollect()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:收费作废按钮操作
    '编制:刘尔旋
    '日期:2013-09-22
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsCollectHistory
        If MsgBox("将收款记录[" & .TextMatrix(.RowSel, .ColIndex("收款单号")) & "]作废，确定作废？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End With
    
    Call CancelCollect
    Call SetDefaultRollingCurtain(True)
    Call RefreshCurrentMoney(1)
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub RollingCurtain()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:轧帐按钮操作
    '编制:刘尔旋
    '日期:2013-09-22
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strNOs As String, i As Integer
    With vsCollectHistory
        For i = 1 To .Rows - 1
            strNOs = strNOs & "," & .TextMatrix(i, .ColIndex("收款单号"))
        Next i
        strNOs = Mid(strNOs, 2)
    End With
    If MsgBox("是否针对如下的收款单据进行轧帐？" & vbCrLf & strNOs, vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Call SaveRollingCurtain
    Call SetDefaultRollingCurtain(False, True)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call zl_vsGrid_Para_Save(mlngModule, vsSubCollectorInfo, Me.Caption, "收费员轧帐明细", False)
    If Not mfrmChargeBillTotalRC Is Nothing Then Unload mfrmChargeBillTotalRC
    Set mobjChargeBillRC = Nothing
End Sub

Private Sub picCurrentMoney_Resize()
    On Error Resume Next
    With lblCurrentMoney(1)
        .Top = 15
        .Width = picCurrentMoney.Width - 15
        .Height = picCurrentMoney.Height - 15
    End With
End Sub

Private Sub vsCollectHistory_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strSQL As String, rsTmp As New ADODB.Recordset, i As Integer
    If OldRow = NewRow Then Exit Sub
    With vsCollectHistory
        'If .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then .Select 0, 0
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
    End With
    
    With vsSubCollectorInfo
        .Rows = 1
        strSQL = "Select No As 轧帐单号, 登记时间 As 轧帐时间, 开始时间 As 开始时间, 终止时间 As 终止时间, " & _
                 "       Trim(to_char(冲预交款,'99999999990.00')) As 冲预交款, Trim(to_char(借入合计,'99999999990.00')) As 借入合计," & _
                 "       Trim(to_char(借出合计,'99999999990.00')) As 借出合计, 登记人, 登记时间, 小组收款人, 小组收款时间, 摘要 As 备注, ID, 收款员" & vbNewLine & _
                 "From 人员收缴记录 " & vbNewLine & _
                 "Where 记录性质 = 1 And 作废时间 Is Null And 小组收款ID= [1]" & vbNewLine & _
                 "Order By 轧帐时间 Desc"
        With vsCollectHistory
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
        End With
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("轧帐单号")) = Nvl(rsTmp!轧帐单号)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐时间")) = Nvl(rsTmp!轧帐时间)
            .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = Nvl(rsTmp!开始时间)
            .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = Nvl(rsTmp!终止时间)
            .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = Nvl(rsTmp!冲预交款)
            .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = Nvl(rsTmp!借入合计)
            .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = Nvl(rsTmp!借出合计)
            .TextMatrix(.Rows - 1, .ColIndex("登记人")) = Nvl(rsTmp!登记人)
            .TextMatrix(.Rows - 1, .ColIndex("登记时间")) = Nvl(rsTmp!登记时间)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款人")) = Nvl(rsTmp!小组收款人)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款时间")) = Nvl(rsTmp!小组收款时间)
            .TextMatrix(.Rows - 1, .ColIndex("备注")) = Nvl(rsTmp!备注)
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTmp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("收款员")) = Nvl(rsTmp!收款员)
            rsTmp.MoveNext
        Loop
        'Set .DataSource = rsTmp
        .AutoSize 1, .Cols - 1
        zl_vsGrid_Para_Restore mlngModule, vsSubCollectorInfo, Me.Caption, "收费员轧帐明细", False
        With vsCollectHistory
            mobjChargeBillRC.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_小组收款, Val(.TextMatrix(.RowSel, .ColIndex("ID")))
        End With
        If .Rows = 1 Then .Rows = 2
    End With
    Call zl_VsGridRowChange(vsCollectHistory, OldRow, NewRow, OldCol, NewCol)
    With vsCollectHistory
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Public Sub ClearChargeAndBillTotalForm()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:外部调用清除票据窗体内容
    '编制:刘尔旋
    '日期:2013-10-12
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    mobjChargeBillRC.ClearChargeAndBillTotalForm
End Sub

Private Sub SetTabControl()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建TAB控件
    '编制:刘尔旋
    '日期:2013-09-04
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With tabSubSendFee
        Set .PaintManager.Font = lblLastTime.Font
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        .InsertItem EM_Tab_收款及票据汇总, " 收款及票据汇总  ", mfrmChargeBillTotalRC.hWnd, 0
        .InsertItem EM_Tab_收费员轧帐明细, " 收费员轧帐明细  ", picSendFeeDetail.hWnd, 0
        .Item(0).Selected = True
        .PaintManager.BoldSelected = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetDateUnit()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:设置日期控件格式属性
    '编制:刘尔旋
    '日期:2013-09-09
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    dtpLastTime.Format = dtpCustom
    dtpLastTime.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpEndTime.Format = dtpCustom
    dtpEndTime.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpEndTime.Value = dtpEndTime.Value + 1
End Sub

Private Sub picTabSendFee_Resize()
    On Error Resume Next
    tabSubSendFee.Width = picTabSendFee.Width
    tabSubSendFee.Height = picTabSendFee.Height
End Sub

Private Sub picLastTime_Resize()
    On Error Resume Next
    cmdSendFees.Left = picLastTime.Width - cmdSendFees.Width - 300
    cmdReloadData.Left = cmdSendFees.Left - cmdReloadData.Width - 300
    If cmdReloadData.Left < dtpEndTime.Left + dtpEndTime.Width + 200 Then
        cmdReloadData.Left = dtpEndTime.Left + dtpEndTime.Width + 200
        cmdSendFees.Left = cmdReloadData.Left + cmdReloadData.Width + 300
    End If
    With vsCollectHistory
        .Width = picLastTime.Width - 15
        .Height = picLastTime.Height - 430
    End With
End Sub

Private Sub vsCollectHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsCollectHistory, Me.Caption, "小组收款信息", False)
End Sub

Private Sub vsCollectHistory_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsCollectHistory_DblClick()
    With vsCollectHistory
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call mobjChargeBillRC.ChargeRollingListShow(Me, EM_小组收款, Val(.TextMatrix(.RowSel, .ColIndex("ID"))), mlngModule, mstrPrivs)
    End With
End Sub

Private Sub vsCollectHistory_GotFocus()
    Call zl_VsGridGotFocus(vsCollectHistory)
    With vsCollectHistory
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsCollectHistory_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsCollectHistory)
End Sub

Private Sub vsSubCollectorInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call zl_VsGridRowChange(vsSubCollectorInfo, OldRow, NewRow, OldCol, NewCol)
    With vsSubCollectorInfo
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsSubCollectorInfo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsSubCollectorInfo, Me.Caption, "收费员轧帐明细", False)
End Sub

Private Sub vsSubCollectorInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsSubCollectorInfo_DblClick()
    With vsSubCollectorInfo
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call mobjChargeBillRC.ChargeRollingListShow(Me, EM_收费员轧帐, Val(.TextMatrix(.RowSel, .ColIndex("ID"))), mlngModule, mstrPrivs)
    End With
End Sub

Private Sub vsSubCollectorInfo_GotFocus()
    Call zl_VsGridGotFocus(vsSubCollectorInfo)
    With vsSubCollectorInfo
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsSubCollectorInfo_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsSubCollectorInfo)
End Sub

Private Sub vsSubCollectorInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intRow As Integer
    With vsSubCollectorInfo
        If .TextMatrix(1, .ColIndex("ID")) = "" Then Exit Sub
        If Button = 2 Then
            If y <= 255 Then
                Exit Sub
            End If
            intRow = y \ 255
            If intRow > .Rows - 1 Then Exit Sub
            If .Enabled And .Visible Then .SetFocus
            .Select intRow, 0
            mcbrPopupSub.ShowPopup
        End If
    End With
End Sub

Private Sub SetGrid()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSF控件
    '编制:刘尔旋
    '日期:2013-10-13
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    With vsSubCollectorInfo
        For i = 0 To .Cols - 1
            If .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "收款员" Or .ColKey(i) = "过滤" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "轧帐单号" Or .ColKey(i) = "开始时间" Or .ColKey(i) = "终止时间" Then .ColData(i) = "1|0"
        Next
    End With
    
    With vsCollectHistory
        For i = 0 To .Cols - 1
            If .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Or .ColKey(i) = "小组收款人" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "过滤" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "收款单号" Or .ColKey(i) = "收款时间" Then .ColData(i) = "1|0"
        Next
    End With
    zl_vsGrid_Para_Restore mlngModule, vsSubCollectorInfo, Me.Caption, "收费员轧帐明细", False
    zl_vsGrid_Para_Restore mlngModule, vsCollectHistory, Me.Caption, "小组收款信息", False
End Sub

Public Sub RefreshPage()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:刷新界面
    '编制:刘尔旋
    '日期:2013-10-13
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Call SetDefaultRollingCurtain(True, True)
    Call RefreshCurrentMoney(1)
    vsCollectHistory.Select 0, 0
End Sub

Private Sub picSendFeeDetail_Resize()
    On Error Resume Next
    With vsSubCollectorInfo
        .Width = picSendFeeDetail.Width
        .Height = picSendFeeDetail.Height
    End With
End Sub

Private Sub dtpEndTime_Change()
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    strSQL = "Select 上次轧帐时间 From 财务组组长构成 Where 组Id= [1] And 组长Id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, UserInfo.ID)
    If rsTmp.RecordCount = 0 Then
        If IsNull(rsTmp!上次轧帐时间) Then
            strSQL = "Select 上次轧帐时间 From 财务缴款分组 Where Id= [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
        End If
    End If
    '无记录或无上次轧帐时间则退出
    If rsTmp.RecordCount = 0 Then Exit Sub
    If IsNull(rsTmp!上次轧帐时间) Then Exit Sub
    If dtpEndTime.Value <= CDate(rsTmp!上次轧帐时间) Then
        dtpEndTime.Value = rsTmp!上次轧帐时间
    End If
End Sub

Private Sub Form_Load()
    mobjChargeBillRC.SetFontSize lblCurrentMoney(1).Font.Size
    Set mfrmChargeBillTotalRC = mobjChargeBillRC.GetChargeAndBillTotalForm
    cmdSendFees.Visible = zlStr.IsHavePrivs(mstrPrivs, "轧帐")
    Call zlDefCommandBars
    Call SetDockingPanel
    Call SetTabControl
    Call SetDateUnit
    Call SetGrid
    '轧帐界面默认信息
    Call SetDefaultRollingCurtain(False)
End Sub

Private Sub imgColPlanSub_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlanSub.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlanSub.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsSubCollectorInfo, lngLeft, lngTop, imgColPlanSub.Height)
    zl_vsGrid_Para_Save mlngModule, vsSubCollectorInfo, Me.Caption, "收费员轧帐明细", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlanSub_Click()
    Call imgColPlanSub_Click
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCollectHistory, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsCollectHistory, Me.Caption, "小组收款信息", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
