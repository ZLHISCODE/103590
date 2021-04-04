VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClientUpgradeConfigure 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   16020
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkAllBefUpgrade 
      Caption         =   "预升级"
      Height          =   280
      Left            =   870
      TabIndex        =   38
      Top             =   600
      Width           =   870
   End
   Begin VB.CheckBox chkAllUpgrade 
      Caption         =   "升级"
      Height          =   280
      Left            =   165
      TabIndex        =   37
      Top             =   600
      Width           =   660
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3630
      ScaleHeight     =   255
      ScaleWidth      =   2625
      TabIndex        =   32
      Top             =   195
      Width           =   2650
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   45
         TabIndex        =   33
         Text            =   "请输入客户端、IP、部门、用途"
         Top             =   30
         Width           =   2650
      End
   End
   Begin VB.PictureBox picMonthSet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   28
      Top             =   4785
      Visible         =   0   'False
      Width           =   255
      Begin VB.CommandButton cmdMonthSet 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -30
         Picture         =   "frmClientUpgradeConfigure.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   -30
         Width           =   285
      End
   End
   Begin VB.PictureBox Picpgb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   15
      ScaleHeight     =   735
      ScaleWidth      =   5475
      TabIndex        =   24
      Top             =   5325
      Visible         =   0   'False
      Width           =   5500
      Begin MSComctlLib.ProgressBar pgbThis 
         Height          =   390
         Left            =   60
         TabIndex        =   25
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   688
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblPer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "100 %"
         ForeColor       =   &H80000001&
         Height          =   180
         Left            =   4920
         TabIndex        =   27
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "INFO"
         Height          =   180
         Left            =   105
         TabIndex        =   26
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.CheckBox chkClientRepair 
      Caption         =   "禁用客户端修复"
      Height          =   270
      Left            =   12900
      TabIndex        =   3
      Top             =   195
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "刷新(&Q)"
      Height          =   300
      Left            =   6435
      TabIndex        =   1
      Top             =   180
      Width           =   1000
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   800
      Index           =   2
      Left            =   5370
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3975
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   800
      Index           =   5
      Left            =   9270
      ScrollBars      =   1  'Horizontal
      TabIndex        =   15
      Top             =   4035
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   800
      Index           =   1
      Left            =   5355
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2805
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   800
      Index           =   0
      Left            =   5325
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1650
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7560
      ScaleHeight     =   315
      ScaleWidth      =   3525
      TabIndex        =   18
      Top             =   165
      Width           =   3525
      Begin VB.OptionButton optStatus 
         Caption         =   "升级失败"
         Height          =   270
         Index           =   2
         Left            =   2535
         TabIndex        =   7
         Top             =   45
         Width           =   1065
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "未升级"
         Height          =   270
         Index           =   0
         Left            =   1605
         TabIndex        =   6
         Top             =   45
         Width           =   960
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "所有"
         Height          =   270
         Index           =   4
         Left            =   930
         TabIndex        =   5
         Top             =   45
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.Label lbloptStatus 
         AutoSize        =   -1  'True
         Caption         =   "升级状态"
         Height          =   180
         Left            =   75
         TabIndex        =   19
         Top             =   75
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdAllCollect 
      Caption         =   "全部收集(&R)"
      Height          =   300
      Left            =   12900
      TabIndex        =   4
      Top             =   555
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picBtn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   60
      ScaleHeight     =   345
      ScaleWidth      =   15735
      TabIndex        =   17
      Top             =   6210
      Width           =   15735
      Begin VB.CommandButton cmdkillProcess 
         Caption         =   "客户端进程管理(&P)"
         Height          =   300
         Left            =   7065
         TabIndex        =   35
         Top             =   0
         Width           =   1800
      End
      Begin VB.CommandButton cmdClientModify 
         Caption         =   "客户端控制修改(&M)"
         Height          =   300
         Left            =   4875
         TabIndex        =   34
         Top             =   0
         Width           =   1965
      End
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   11835
         ScaleHeight     =   240
         ScaleWidth      =   1920
         TabIndex        =   30
         Top             =   15
         Width           =   1945
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   315
            Left            =   -30
            TabIndex        =   31
            ToolTipText     =   "该时间点前进行预升级，该时间点以后进行正式升级"
            Top             =   -30
            Width           =   2000
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   106299395
            CurrentDate     =   42691
         End
      End
      Begin VB.CommandButton cmdTimeSet 
         Caption         =   "预升级时间设置(&T)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   13905
         TabIndex        =   12
         Top             =   0
         Width           =   1800
      End
      Begin VB.OptionButton optUpgradeTime 
         Caption         =   "定时升级"
         Height          =   210
         Index           =   1
         Left            =   10695
         TabIndex        =   11
         ToolTipText     =   "可以设置定时升级时间点，该时间点以前会进行预升级，该时间点以后会进行正式升级"
         Top             =   60
         Width           =   1170
      End
      Begin VB.OptionButton optUpgradeTime 
         Caption         =   "立即升级"
         Height          =   210
         Index           =   0
         Left            =   9540
         TabIndex        =   10
         ToolTipText     =   "对客户端勾选升级后，客户端登陆即会自动正式升级"
         Top             =   60
         Width           =   1050
      End
      Begin VB.CommandButton cmdClientAaminSet 
         Caption         =   "客户端通用管理员设置(&K)"
         Height          =   300
         Left            =   2475
         TabIndex        =   9
         Top             =   0
         Width           =   2400
      End
      Begin VB.CommandButton cmdFileSeverSet 
         Caption         =   "升级文件服务器设置(&S)"
         Height          =   300
         Left            =   60
         TabIndex        =   8
         Top             =   0
         Width           =   2200
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   4185
      Left            =   120
      TabIndex        =   2
      Top             =   585
      Width           =   4905
      _cx             =   8652
      _cy             =   7382
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClientUpgradeConfigure.frx":00F6
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
   Begin MSComctlLib.ImageList imgEdit 
      Left            =   10170
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":01CD
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":0767
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":0D01
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":1053
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":78B5
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":E117
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientUpgradeConfigure.frx":E5DF
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblClientsList 
      AutoSize        =   -1  'True
      Caption         =   "客户端升级清单"
      Height          =   180
      Left            =   135
      TabIndex        =   36
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "主动修复说明"
      Height          =   180
      Index           =   2
      Left            =   5370
      TabIndex        =   23
      Top             =   3645
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "收集说明"
      Height          =   180
      Index           =   5
      Left            =   9270
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "预升级说明"
      Height          =   180
      Index           =   1
      Left            =   5370
      TabIndex        =   21
      Top             =   2505
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      Caption         =   "升级说明"
      Height          =   180
      Index           =   0
      Left            =   5340
      TabIndex        =   20
      Top             =   1350
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查找"
      Height          =   180
      Left            =   3135
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmClientUpgradeConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrLocationClientsName As String '定位上次选中行客户端名称
Private mblnFilter As Boolean '记录当前是否过滤过表格 true - 已过滤 false -为过滤
Private mblnCancel As Boolean
Private mlngClinetNum As Long           '客户端总数
Private mlngUpFailClinetNum As Long '升级失败客户端数量
Private mlngNotUpClinetNum As Long '未升级客户端数量
Public blnRefreshData As Boolean '界面切换刷新判断标志
Private mblnAllowEdit As Boolean '标记当前界面是否允许编辑
Private mblnAllUpdateClick As Boolean '因为设置chkAllUpgrade.value值的时候会隐式调用chkAllUpgrade_Click，故需要本值来屏蔽隐式调用
Private mblnAllBefUpgrade As Boolean  '和mblnAllUpdateClick一样

Private Enum SeverData
    Col_升级 = 0
    Col_预升级 = 1
    Col_收集 = 2
    Col_客户端 = 3
    Col_IP = 4
    Col_部门 = 5
    Col_用途 = 6
    Col_升级服务器 = 7
    Col_预升级时点 = 8
    Col_更新检查 = 9
    Col_升级结果 = 10
    Col_预升级结果 = 11
'    Col_收集结果 = 12
    Col_主动修复结果 = 12
    Col_升级说明 = 13
    Col_预升级说明 = 14
'    Col_收集说明 = 15
    Col_主动修复说明 = 15
    Col_管理员 = 16
    Col_密码 = 17
    Col_列数 = 18
End Enum

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
End Sub

Private Sub chkAllBefUpgrade_Click()
    If mblnAllBefUpgrade Then Exit Sub
    If chkAllBefUpgrade.value = 1 Then
        If MsgBox("是否要设置全部客户端启用预升级？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
            mblnAllBefUpgrade = True
            chkAllBefUpgrade.value = 0
            mblnAllBefUpgrade = False
            Exit Sub
        End If
    Else
        If MsgBox("是否要取消全部客户端启用预升级？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
            mblnAllBefUpgrade = True
            chkAllBefUpgrade.value = 1
            mblnAllBefUpgrade = False
            Exit Sub
        End If
    End If
    mblnAllBefUpgrade = False
    On Error GoTo errH
    Call UpdateData(Col_预升级, chkAllBefUpgrade.value)
    RefreshData
    '插入重要操作日志
    Call SaveAuditLog(2, "全部预升级/取消全部预升级", "对所有客户端执行" & IIf(chkAllBefUpgrade.value = 1, "", "取消") & "预升级操作")
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub chkAllUpgrade_Click()
    If mblnAllUpdateClick Then Exit Sub
    If chkAllUpgrade.value = 1 Then
        If MsgBox("是否要设置全部客户端启用升级？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
            mblnAllUpdateClick = True
            chkAllUpgrade.value = 0
            mblnAllUpdateClick = False
            Exit Sub
        End If
    Else
        If MsgBox("是否要取消全部客户端启用升级？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
            mblnAllUpdateClick = True
            chkAllUpgrade.value = 1
            mblnAllUpdateClick = False
            Exit Sub
        End If
    End If
    mblnAllUpdateClick = False
    On Error GoTo errH
    Call UpdateData(Col_升级, chkAllUpgrade.value)
    RefreshData
    '插入重要操作日志
    Call SaveAuditLog(2, "全部升级/取消全部升级", "对所有客户端进行" & IIf(chkAllUpgrade.value = 1, "", "取消") & "升级操作")
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub chkClientRepair_Click()
    Dim strSQL As String
    
    On Error Resume Next
    If chkClientRepair.value = 0 Then
        strSQL = "update zltools.ZLReginfo set 内容 = '" & 0 & "'where 项目 = '禁止客户端修复'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "update zltools.ZLReginfo set 内容 = '" & 1 & "'where 项目 = '禁止客户端修复'"
        gcnOracle.Execute strSQL
    End If
    
End Sub

Private Sub cmdAllCollect_Click()
    Dim i As Long
    Dim strSQL As String
    Dim strUpdateVal As String
    Dim strTemp As String
    
    strTemp = cmdAllCollect.Caption
    strUpdateVal = IIf(strTemp = "全部收集(&R)", "1", "0")
    If MsgBox("是否要" & IIf(strUpdateVal = "1", "设置", "取消") & "全部客户端启用收集？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub

    On Error GoTo errH
    Call UpdateData(Col_收集, strUpdateVal)
    cmdAllCollect.Caption = IIf(strTemp = "全部收集(&R)", "全部不收集(&R)", "全部收集(&R)")
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub cmdClientAaminSet_Click()
    Load frmClientUpgradeAdmin
    frmClientUpgradeAdmin.Show 1, frmMDIMain
    If frmClientUpgradeAdmin.mblnOk Then
    End If
    Exit Sub
End Sub

Private Sub cmdClientModify_Click()
    Dim blnReturn   As Boolean
    Dim strIp       As String
    Dim strName     As String
    Dim lngRow      As Long
    
    With vsfMain
        If .Row >= .FixedRows Then
            lngRow = .Row
            strIp = .TextMatrix(lngRow, Col_IP)
            strName = .TextMatrix(lngRow, Col_客户端)
            frmClientsEdit.ShowEdit strIp, strName, 1, blnReturn
            If Not blnReturn Then Exit Sub
            Call LoadClientsData
            lngRow = .FindRow(strName, , Col_客户端)
            If lngRow >= .FixedRows Then
                .SetFocus
                .Row = lngRow
                .ShowCell lngRow, Col_客户端
            End If
        End If
    End With
End Sub

Private Sub cmdFileSeverSet_Click()
    Dim frmSeverSet As New frmClientUpgradeSever
    If frmSeverSet.ShowMe(frmMDIMain) = True Then
        cmdRef_Click
    End If
End Sub

'设置需要查杀的进程列表
Private Sub cmdkillProcess_Click()
    frmKillProcessManage.ShowMe ("0307")
End Sub

Private Sub cmdRef_Click()
    Call RefreshData
End Sub

Private Sub cmdTimeSet_Click()
    Load frmClientUpgradeTime
    frmClientUpgradeTime.Show 1, frmMDIMain
    If frmClientUpgradeTime.mblnOk Then
        LoadClientsData
        FilterData (mstrLocationClientsName)
        InitCombolist
    End If
    Exit Sub
ErrHandle:
    MsgBox "保存失败。" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Sub

Private Sub dtpTime_Change()
'    Dim strNow As String
'    strNow = Format(CurrentDate(), "yyyy-MM-dd") & " 23:00"
'    dtpTime.value = strNow
    Call SaveUpgradeDate
End Sub

Private Sub Form_Load()
'    Call LoadClientsData
'    Call FilterData
    mblnAllowEdit = True
    Call InitVsfMain
    Call LoadSetting
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngTxtHeight As Long
    vsfMain.Height = Me.ScaleHeight - vsfMain.Top - 600
    
    lngTxtHeight = (vsfMain.Height - 960) / 3
    
    If lngTxtHeight < 300 Or Me.ScaleWidth < 8000 Then
        lblExplain.Item(0).Visible = False
        lblExplain.Item(1).Visible = False
        lblExplain.Item(2).Visible = False
'        lblExplain.Item(3).Visible = False
        txtExplain.Item(0).Visible = False
        txtExplain.Item(1).Visible = False
        txtExplain.Item(2).Visible = False
'        txtExplain.Item(3).Visible = False
        vsfMain.Width = Me.ScaleWidth - 100
        picStatus.Visible = False
    Else
        lblExplain.Item(0).Visible = True
        lblExplain.Item(1).Visible = True
        lblExplain.Item(2).Visible = True
'        lblExplain.Item(3).Visible = True
        txtExplain.Item(0).Visible = True
        txtExplain.Item(1).Visible = True
        txtExplain.Item(2).Visible = True
'        txtExplain.Item(3).Visible = True
        vsfMain.Width = Me.ScaleWidth - 100 - 2600
        picStatus.Visible = True
        With lblExplain
            .Item(0).Move vsfMain.Left + vsfMain.Width + 90, vsfMain.Top
            .Item(1).Move .Item(0).Left, .Item(0).Top + lngTxtHeight + 330
            .Item(2).Move .Item(1).Left, .Item(1).Top + lngTxtHeight + 330
'            .Item(3).Move .Item(2).Left, .Item(2).Top + lngTxtHeight + 250
            txtExplain.Item(0).Move .Item(0).Left, .Item(0).Top + 290, 2500, lngTxtHeight
            txtExplain.Item(1).Move .Item(1).Left, .Item(1).Top + 290, 2500, lngTxtHeight
            txtExplain.Item(2).Move .Item(2).Left, .Item(2).Top + 290, 2500, lngTxtHeight
'            txtExplain.Item(3).Move .Item(3).Left, .Item(3).Top + 210, 2500, lngTxtHeight
    '        cmdRef.Move .Item(0).Left + 750, picBtn.Top
        End With
    End If
'    picBtn.Top = Me.ScaleHeight - picBtn.Top - 60
    Call Picpgb.Move((Me.Width - Picpgb.Width) / 2, (Me.Top - Picpgb.Height) / 2 + 2000)
    picBtn.Top = vsfMain.Top + vsfMain.Height + 150
    picStatus.Left = vsfMain.Left + vsfMain.Width - picStatus.Width
    cmdRef.Left = picStatus.Left - cmdRef.Width - 300
    PicFind.Left = cmdRef.Left - PicFind.Width - 200
    lblFind.Left = PicFind.Left - lblFind.Width - 100
End Sub

Private Sub LoadClientsData()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim strArr() As String
    Dim blnUpgrade As Boolean
    Dim blnBefUpgrade As Boolean
    Dim blnCollect As Boolean
    Dim intBatch As Integer
    Dim i As Long
    
    With vsfMain
        .Rows = .FixedRows
        strSQL = "Select Max(内容) As 最新批次 From zlRegInfo Where 项目 = '最新升级文件批次'"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        intBatch = Val(rsTemp!最新批次 & "")
        
'        strSQL = "select 工作站,IP,部门,升级服务器,升级标志,是否预升级,收集标志,预升时点,升级情况,预升完成,收集状态,修复状态,升级说明,预升级说明,收集说明,修复说明 from zlclients"
        strSQL = "select A.工作站,A.IP,A.部门,A.用途,A.升级文件服务器,B.位置, A.升级标志,A.是否预升级,A.收集标志,A.预升时点,A.升级情况,A.预升完成,A.收集状态,A.修复状态,A.升级说明,A.预升级说明,A.收集说明,A.修复说明,A.批次,A.管理员用户,A.管理员密码 from zlclients A,zlupgradeserver B where A.升级文件服务器 = B.编号(+)"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
        '数据填入
        .Rows = rsTemp.RecordCount + 1
        .Redraw = flexRDNone
        i = 1
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, Col_客户端) = Nvl(rsTemp.Fields("工作站"))
            .Cell(flexcpText, i, Col_IP) = Nvl(rsTemp.Fields("IP"))
            .Cell(flexcpText, i, Col_部门) = Nvl(rsTemp.Fields("部门"))
            .Cell(flexcpText, i, Col_用途) = Nvl(rsTemp.Fields("用途"))
            
            strTemp = Trim(Nvl(rsTemp.Fields("升级文件服务器"), ""))
            If Trim(rsTemp.Fields("位置")) & "" = "" And strTemp <> "" Then
'                strSQL = "update ZLClients set 升级文件服务器 = null where 升级文件服务器 = " & vsfMain.TextMatrix(vsfMain.Row, Col_编号)
'                gcnOracle.Execute strSQL
                .Cell(flexcpText, i, Col_升级服务器) = ""
            Else
                .Cell(flexcpText, i, Col_升级服务器) = IIf(strTemp <> "" And Trim(rsTemp.Fields("位置")) <> "", Nvl(rsTemp.Fields("升级文件服务器"), "") & ":" & Nvl(rsTemp.Fields("位置"), ""), "")
            End If
            
'            If Val(rsTemp!批次 & "") < intBatch Then
'                .Cell(flexcpText, i, Col_更新检查) = "需要更新"
'            Else
'                .Cell(flexcpText, i, Col_更新检查) = "无需更新"
'            End If
            
'            .Cell(flexcpText, i, Col_升级) = IIf(Nvl(rsTemp.Fields("升级标志"), "") = "1", "√", "")
'            If .TextMatrix(i, Col_升级) = "" Then blnUpgrade = True
'            .Cell(flexcpText, i, Col_预升级) = IIf(Nvl(rsTemp.Fields("是否预升级"), "") = "1", "√", "")
'            If .TextMatrix(i, Col_预升级) = "" Then blnBefUpgrade = True
'            .Cell(flexcpText, i, Col_收集) = IIf(Nvl(rsTemp.Fields("收集标志"), "") = "1", "√", "")
'            If .TextMatrix(i, Col_收集) = "" Then blnCollect = True

            .Cell(flexcpText, i, Col_升级) = IIf(Nvl(rsTemp.Fields("升级标志"), "") = "1", True, False)
            If .Cell(flexcpText, i, Col_升级) = False Then blnUpgrade = True
            .Cell(flexcpText, i, Col_预升级) = IIf(Nvl(rsTemp.Fields("是否预升级"), "") = "1", True, False)
            If .Cell(flexcpText, i, Col_预升级) = False Then blnBefUpgrade = True
            
            .Cell(flexcpText, i, Col_预升级时点) = Format(Nvl(rsTemp.Fields("预升时点")), "hh:mm")
            
            strTemp = Nvl(rsTemp.Fields("升级情况"), "0")
            .Cell(flexcpData, i, Col_升级结果) = strTemp
            .Cell(flexcpText, i, Col_升级结果) = Decode(strTemp, "0", "未升级", "1", "完成", "2", "失败", "3", "正在升级", "")

            strTemp = Nvl(rsTemp.Fields("预升完成"), "0")
            .Cell(flexcpData, i, Col_预升级结果) = strTemp
            .Cell(flexcpText, i, Col_预升级结果) = Decode(strTemp, "0", "未升级", "1", "完成", "2", "失败", "3", "正在升级", "")

'            strTemp = Nvl(rsTemp.Fields("收集状态"), "0")
'            .Cell(flexcpData, i, Col_收集结果) = strTemp
'            .Cell(flexcpText, i, Col_收集结果) = Decode(strTemp, "0", "未收集", "1", "完成", "2", "失败", "3", "正在收集", "")
            
            strTemp = Nvl(rsTemp.Fields("修复状态"), "0")
            .Cell(flexcpData, i, Col_主动修复结果) = strTemp
            .Cell(flexcpText, i, Col_主动修复结果) = Decode(strTemp, "0", "未修复", "1", "完成", "2", "失败", "3", "正在修复", "")
            .Cell(flexcpText, i, Col_升级说明) = Nvl(rsTemp.Fields("升级说明"))
            .Cell(flexcpText, i, Col_预升级说明) = Nvl(rsTemp.Fields("预升级说明"))
'            .Cell(flexcpText, i, Col_收集说明) = Nvl(rsTemp.Fields("收集说明"))
            .Cell(flexcpText, i, Col_主动修复说明) = Nvl(rsTemp.Fields("修复说明"))
            .Cell(flexcpText, i, Col_管理员) = Nvl(rsTemp.Fields("管理员用户"))
            .Cell(flexcpText, i, Col_密码) = Decipher(Nvl(rsTemp.Fields("管理员密码")))
            rsTemp.MoveNext
            i = i + 1
        Loop
        '文本风格设置
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, Col_客户端, .Rows - 1, Col_客户端) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_IP, .Rows - 1, Col_IP) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_部门, .Rows - 1, Col_部门) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_升级服务器, .Rows - 1, Col_升级服务器) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_更新检查, .Rows - 1, Col_更新检查) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, Col_升级, .Rows - 1, Col_升级) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, Col_预升级, .Rows - 1, Col_预升级) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_收集, .Rows - 1, Col_收集) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_预升级时点, .Rows - 1, Col_预升级时点) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_升级结果, .Rows - 1, Col_升级结果) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_预升级结果, .Rows - 1, Col_预升级结果) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, Col_收集结果, .Rows - 1, Col_收集结果) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_主动修复结果, .Rows - 1, Col_主动修复结果) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_升级说明, .Rows - 1, Col_升级说明) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_预升级说明, .Rows - 1, Col_预升级说明) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, Col_收集说明) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_主动修复说明) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, Col_管理员, .Rows - 1, Col_管理员) = flexAlignLeftCenter
            .Cell(flexcpAlignment, .FixedRows, Col_密码, .Rows - 1, Col_密码) = flexAlignLeftCenter
            .Cell(flexcpBackColor, .FixedRows, Col_用途, .Rows - 1, Col_用途) = RGB(210, 240, 255)  ' RGB(247, 247, 247)
            .Cell(flexcpBackColor, .FixedRows, Col_升级服务器, .Rows - 1, Col_升级服务器) = RGB(210, 240, 255)   'RGB(247, 247, 247)
            .Cell(flexcpBackColor, .FixedRows, Col_升级, .Rows - 1, Col_收集) = RGB(210, 240, 255)   'RGB(247, 247, 247)
            .Cell(flexcpBackColor, .FixedRows, Col_预升级时点, .Rows - 1, Col_预升级时点) = RGB(210, 240, 255)   'RGB(247, 247, 247)
        End If
        .Redraw = flexRDDirect
    End With

    
    '设置按键状态
    mblnAllUpdateClick = True
    If blnUpgrade Then
        chkAllUpgrade.value = 0
    Else
        chkAllUpgrade.value = 1
    End If
    mblnAllUpdateClick = False
    mblnAllBefUpgrade = True
    If blnBefUpgrade Then
        chkAllBefUpgrade.value = 0
    Else
        chkAllBefUpgrade.value = 1
    End If
    mblnAllBefUpgrade = False
    If blnCollect Then
        cmdAllCollect.Caption = "全部收集(&R)"
    Else
        cmdAllCollect.Caption = "全部不收集(&R)"
    End If
    '加载表格下拉列表内容
    InitCombolist
End Sub

Public Sub SetMenu(Optional lngrows As Long = -1)
    If lngrows = -1 Then lngrows = vsfMain.Rows - 1
'    frmMDIMain.stbThis.Panels(2).Text = "列表中共显示有" & lngrows & "行数据。"
    frmMDIMain.stbThis.Panels(2).Text = "列表中共显示有" & mlngClinetNum & "个客户端，未升级的客户端有" & mlngNotUpClinetNum & "个，升级失败的客户端有" & mlngUpFailClinetNum & "个。"
End Sub

Private Function CheckIP(strIp As String) As Boolean
'检查IP格式是否正确
    Dim sTmp() As String
    Dim i As Integer
    
    If strIp = "" Then CheckIP = False: Exit Function
    
    sTmp = Split(strIp, ".")
    If UBound(sTmp) <> 3 Then CheckIP = False: Exit Function
    
    For i = 0 To UBound(sTmp)
        If sTmp(i) = "" Then CheckIP = False: Exit Function
        
        If CLng(sTmp(i)) > 255 Or CLng(sTmp(i)) < 0 Or i > 3 Then CheckIP = False: Exit Function
    Next i
    
    CheckIP = True
End Function

Private Sub LoadSetting()
    '界面控件设置、状态读取加载
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errH
    
    lbloptStatus.Tag = "4" '过滤默认为显示全部
    
    txtFind.Tag = "请输入客户端、IP、部门、用途"
    txtFind.Text = txtFind.Tag
    txtFind.ForeColor = vbGrayText

'    vsfMain_RowColChange
    
    '禁止客户端修复 设置读取（删除）
'    strSQL = "select 内容 from ZLReginfo where 项目 = '禁止客户端修复'"
'    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
'
'    If rsTemp.EOF Then
'        strSQL = "insert into zltools.ZLReginfo(项目,内容) select '禁止客户端修复','0' from dual where not Exists (select 1 from zltools.ZLReginfo where 项目 ='禁止客户端修复')"
'        gcnOracle.Execute strSQL
'        chkClientRepair.value = 0
'    Else
'        chkClientRepair.value = CInt(Nvl(rsTemp.Fields("内容"), "0"))
'    End If
    
    '定时升级设置读取
    strSQL = "select 内容 from ZLReginfo where 项目 = '客户端升级日期'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    If rsTemp.EOF = False Then
        dtpTime.value = Format(Nvl(rsTemp.Fields("内容"), CurrentDate()), "yyyy-MM-dd hh:mm")
        optUpgradeTime.Item(1).value = True
    Else
        dtpTime.value = Format(CurrentDate(), "yyyy-MM-dd") & " 23:00"
        optUpgradeTime.Item(0).value = True
    End If

    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 1 = 0 Then
        Resume
    End If
End Sub

Private Sub ShowPercent(Optional blnVisible As Boolean, Optional strInfo As String, Optional sngPer As Single = -1, Optional blnPer As Boolean = False)
'显示处理进度
    If blnVisible = False Then Picpgb.Visible = False: Exit Sub
    
    If sngPer = -1 Then
        pgbThis.value = 0
    Else
        If sngPer >= 1 Then
            pgbThis.value = CInt(sngPer)
        Else
            pgbThis.value = CInt(sngPer * 100)
        End If
    End If
    
    pgbThis.Max = 100

    lblInfo.Caption = strInfo
    lblPer.Caption = CInt(pgbThis.value) & " %"
    
    If blnPer = True Then
        lblPer.Visible = True
    Else
        lblPer.Visible = False
    End If
    Picpgb.Visible = True
    
End Sub

Private Sub InitCombolist()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim blnTemp As Boolean
    
    On Error GoTo errH:
    With vsfMain
        .Editable = flexEDKbdMouse
        '升级服务器下拉列表服务器加载
        strSQL = "select 编号,类型,位置,是否升级,是否缺省,是否收集 from zltools.zlupgradeserver"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        .ColComboList(Col_升级服务器) = "#10*2;" & "" & vbTab & "" & vbTab & " "
        i = 2
        Do Until rsTemp.EOF
            If Nvl(rsTemp.Fields("是否升级"), "0") = "1" Or Nvl(rsTemp.Fields("是否缺省"), "0") = "1" Or Nvl(rsTemp.Fields("是否收集"), "0") = "1" Then
                blnTemp = True
            Else
                blnTemp = False
            End If
            If blnTemp Then
                .ColComboList(Col_升级服务器) = .ColComboList(Col_升级服务器) & _
                "|#" & i * 10 & ";" & rsTemp.Fields("编号") & "号" & vbTab & IIf(rsTemp.Fields("类型") = 0, "共享", "FTP") & vbTab & rsTemp.Fields("编号") & ":" & rsTemp.Fields("位置")
            End If
            rsTemp.MoveNext
            i = i + 1
        Loop
        '用途下拉列表服务器加载
        strSQL = "select distinct 用途 from zlclients order by 用途"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        Do Until rsTemp.EOF
            If Nvl(rsTemp!用途, "") <> "" Then
                .ColComboList(Col_用途) = .ColComboList(Col_用途) & "|" & rsTemp!用途
            End If
            rsTemp.MoveNext
        Loop
        .ColComboList(Col_用途) = " " & .ColComboList(Col_用途)
        
        '预升级时间点下拉列表
        strSQL = "select 项目,内容 from zltools.ZLReginfo where 项目 = '客户端预升级时间点'"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        If Not rsTemp.EOF Then
            .ColComboList(Col_预升级时点) = Replace(Nvl(rsTemp!内容), ",", "|")
            .ColComboList(Col_预升级时点) = " |" & .ColComboList(Col_预升级时点)
        End If
'        详细报告
        .ColComboList(Col_升级结果) = "..."
        .ColComboList(Col_预升级结果) = "..."
        .ColComboList(Col_主动修复结果) = "..."
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub optStatus_Click(Index As Integer)
    Dim lngRowCount As Long
    Dim i As Long
    If optStatus(Index).Visible = False Then Exit Sub
    lbloptStatus.Tag = Index
    Call FilterData(mstrLocationClientsName)
End Sub

Private Sub optUpgradeTime_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpTime.Enabled = False
            cmdTimeSet.Enabled = False
            chkAllBefUpgrade.Visible = False
            vsfMain.ColHidden(Col_预升级) = True
            vsfMain.ColHidden(Col_预升级时点) = True
            vsfMain.ColHidden(Col_预升级结果) = True
            txtExplain(1).Enabled = False
            lblExplain(1).Enabled = False
            If optUpgradeTime(Index).Visible Then
                If chkAllBefUpgrade.value = 0 Then
                    Call chkAllBefUpgrade_Click
                Else
                    chkAllBefUpgrade.value = 0
                End If
                SaveUpgradeDate True
            End If
        Case 1
            dtpTime.Enabled = True
            cmdTimeSet.Enabled = True
            chkAllBefUpgrade.Visible = True
            vsfMain.ColHidden(Col_预升级) = False
            vsfMain.ColHidden(Col_预升级时点) = False
            vsfMain.ColHidden(Col_预升级结果) = False
            txtExplain(1).Enabled = True
            lblExplain(1).Enabled = True
            If optUpgradeTime(Index).Visible Then
                mblnAllBefUpgrade = True
                chkAllBefUpgrade.value = 0
                mblnAllBefUpgrade = False
                SaveUpgradeDate
            End If
    End Select
End Sub

Private Sub txtExplain_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = txtFind.Tag Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    FilterData mstrLocationClientsName
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = txtFind.Tag
        txtFind.ForeColor = vbGrayText
    End If
End Sub

Private Sub vsfMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strTemp As String
    Dim strSave() As String
    Dim strSQL As String
    
    With vsfMain
        Select Case Col
        Case Col_用途
            strTemp = Trim(.Cell(flexcpTextDisplay, .Row, Col_用途))
            If strTemp <> "" Then
                strSQL = "update zltools.zlclients set 用途 = '" & strTemp & "' where 工作站 = '" & .TextMatrix(.Row, Col_客户端) & "'"
                gcnOracle.Execute strSQL
            Else
                strSQL = "update zltools.zlclients set 用途 = null where 工作站 = '" & .TextMatrix(.Row, Col_客户端) & "'"
                gcnOracle.Execute strSQL
            End If
        Case Col_升级服务器
            strTemp = Trim(.Cell(flexcpTextDisplay, .Row, Col_升级服务器))
            If strTemp <> "" Then
                strSave = Split(strTemp, ":")
                If IsNumeric(strSave(0)) = False Then Exit Sub
                strSQL = "update zltools.zlclients set 升级文件服务器 = " & Trim(strSave(0)) & " where 工作站 = '" & .TextMatrix(.Row, Col_客户端) & "'"
                gcnOracle.Execute strSQL
            Else
                strSQL = "update zltools.zlclients set 升级文件服务器 = null where 工作站 = '" & .TextMatrix(.Row, Col_客户端) & "'"
                gcnOracle.Execute strSQL
            End If
        Case Col_预升级时点
            strTemp = Trim(.Cell(flexcpTextDisplay, .Row, Col_预升级时点))
            If strTemp <> "" Then
                strTemp = Format(Now(), "yyyy/MM/dd") & " " & Format(strTemp, "hh:mm:00")
                strTemp = "to_date('" & strTemp & "','YYYY/MM/DD HH24:MI:SS')"
            Else
                strTemp = "NULL"
            End If
            strSQL = "update zltools.zlclients set 预升时点 = " & strTemp & " where 工作站 = '" & .TextMatrix(.Row, Col_客户端) & "'"
            gcnOracle.Execute strSQL
        End Select
    End With
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    txtExplain(0).Text = ""
    txtExplain(1).Text = ""
    txtExplain(2).Text = ""
'    txtExplain(3).Text = ""
    If NewRow < vsfMain.FixedRows Or NewCol < vsfMain.FixedCols Then Exit Sub
    
    With vsfMain
        mstrLocationClientsName = .TextMatrix(NewRow, Col_客户端)
        txtExplain(0).Text = .TextMatrix(NewRow, Col_升级说明)
        txtExplain(1).Text = .TextMatrix(NewRow, Col_预升级说明)
'        txtExplain(2).Text = .TextMatrix(NewRow, Col_收集说明)
        txtExplain(2).Text = .TextMatrix(NewRow, Col_主动修复说明)
    End With
End Sub

Private Sub vsfMain_AfterSort(ByVal Col As Long, Order As Integer)
    vsfMain.Row = vsfMain.FindRow(mstrLocationClientsName, , Col_客户端)
    vsfMain.ShowCell vsfMain.Row, 0
End Sub

Private Sub vsfMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Col_升级服务器 And Col <> Col_预升级时点 And Col <> Col_用途 And Col <> Col_升级结果 And Col <> Col_预升级结果 And Col <> Col_主动修复结果 Then Cancel = True
End Sub

Private Sub vsfMain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = Col_升级 Or Col = Col_预升级 Then
        Cancel = True
    End If
End Sub

Private Sub vsfMain_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Row < vsfMain.FixedRows Then Exit Sub
    frmClientUpgradeLogView.mstrName = vsfMain.TextMatrix(Row, Col_客户端)
    Load frmClientUpgradeLogView
    frmClientUpgradeLogView.Show 1, frmMDIMain
End Sub

Private Sub vsfMain_DblClick()
    Dim strSQL As String
    
    If mblnAllowEdit = False Then Exit Sub
    With vsfMain
        If .MouseRow <> .Row Then Exit Sub '非选中行双击无效，屏蔽固定行双击
        Select Case .ColSel
        Case Col_升级
            .TextMatrix(.RowSel, .ColSel) = IIf(.TextMatrix(.RowSel, .ColSel) = True, False, True)
            strSQL = "Zl_Zlclients_Update('" & .TextMatrix(.RowSel, Col_客户端) & "'," & 0 & "," & IIf(.TextMatrix(.RowSel, .ColSel) = True, "1", "0") & ")"
            Call ExecuteProcedure(strSQL, Me.Caption)
            If .TextMatrix(.RowSel, .ColSel) = True Then
                If .TextMatrix(.Row, Col_升级结果) <> "未升级" Then
                    mlngNotUpClinetNum = mlngNotUpClinetNum + 1
                    If .TextMatrix(.Row, Col_升级结果) = "失败" Then
                        mlngUpFailClinetNum = mlngUpFailClinetNum - 1
                    End If
                    SetMenu
                End If
                .TextMatrix(.Row, Col_升级结果) = "未升级"
                .TextMatrix(.Row, Col_升级说明) = ""
                .TextMatrix(.Row, Col_主动修复结果) = "未修复"
                .TextMatrix(.Row, Col_主动修复说明) = ""
                txtExplain(0).Text = ""
                txtExplain(2).Text = ""
            End If
            '插入重要操作日志
            Call SaveAuditLog(2, "升级/取消升级", "对客户端“" & .TextMatrix(.Row, Col_客户端) & "”进行" & IIf(.TextMatrix(.RowSel, .ColSel) = True, "", "取消") & "升级操作")
        Case Col_预升级
            .TextMatrix(.RowSel, .ColSel) = IIf(.TextMatrix(.RowSel, .ColSel) = True, False, True)
            strSQL = "Zl_Zlclients_Update('" & .TextMatrix(.RowSel, Col_客户端) & "'," & 1 & "," & IIf(.TextMatrix(.RowSel, .ColSel) = True, "1", "0") & ")"
            Call ExecuteProcedure(strSQL, Me.Caption)
            If .TextMatrix(.RowSel, .ColSel) = True Then
                .TextMatrix(.Row, Col_预升级结果) = "未升级"
                .TextMatrix(.Row, Col_预升级说明) = ""
                txtExplain(1).Text = ""
            End If
            '插入重要操作日志
            Call SaveAuditLog(2, "预升级/取消预升级", "对客户端“" & .TextMatrix(.Row, Col_客户端) & "”进行" & IIf(.TextMatrix(.RowSel, .ColSel) = True, "", "取消") & "预升级操作")
        Case Col_收集
            .TextMatrix(.RowSel, .ColSel) = IIf(.TextMatrix(.RowSel, .ColSel) = "√", "", "√")
            strSQL = "update zltools.ZlClients set 收集标志 = " & IIf(.TextMatrix(.RowSel, .ColSel) = "√", "1", "0") & " where 工作站 = '" & .TextMatrix(.RowSel, Col_客户端) & "'"
            gcnOracle.Execute strSQL
        End Select
    End With
End Sub

Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题
    '-------------------------------------------------------------
    Dim rsOut As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    Set rsOut = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Current_date")
    If rsOut.RecordCount > 0 Then
        CurrentDate = IIf(IsNull(rsOut.Fields(0)), 0, rsOut.Fields(0))
    Else
        CurrentDate = 0
    End If
    Exit Function
ErrHandle:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If
End Function

Private Function UpdateData(strUpdateCol As String, strUpdateVal As String, Optional blnRef As Boolean = False) As Boolean
    Dim arrSQL() As Variant
    Dim i As Long, strName As String, blnTrans As Boolean
    Dim strUpdateField As String
    Dim strTemp As String
    Dim strSQL  As String
    Select Case strUpdateCol
        Case Col_升级
            strUpdateField = 0
        Case Col_预升级
            strUpdateField = 1
        Case Col_收集
            strUpdateField = 2
        Case Else
            UpdateData = False: Exit Function
    End Select
    
    On Error GoTo errH:
    arrSQL() = Array()
    With vsfMain
        If .Rows < 1 Then Exit Function
        Me.Enabled = False
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            If .RowHidden(i) = False Then
                If ActualLen(strName & "," & .TextMatrix(i, Col_客户端)) > 3900 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_Zlclients_Update('" & strName & "'," & strUpdateField & "," & strUpdateVal & ")"
                    
                    strName = Trim(.TextMatrix(i, Col_客户端))
                Else
                    strName = IIf(strName = "", "", strName & ",") & (.TextMatrix(i, Col_客户端))
                End If
                .TextMatrix(i, strUpdateCol) = IIf(strUpdateVal = "1", True, False)
            End If
        Next
        .Redraw = flexRDBuffered
        
        If strName <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_Zlclients_Update('" & strName & "'," & strUpdateField & "," & strUpdateVal & ")"
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
        
            strSQL = arrSQL(i)
            Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        Next
        gcnOracle.CommitTrans: blnTrans = False
                
        Me.Enabled = True
    End With
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    MsgBox err.Description, vbExclamation, gstrSysName
    If 1 = 0 Then
        Resume
    End If
End Function

Private Function SaveUpgradeDate(Optional blnDelete As Boolean = False) As Boolean
'    存储定时升级时间内容
'    blnDelete 删除时间，true-删除，false-不删除
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim strTime As String
    Dim intTemp As Integer
    
    On Error GoTo errH
    
    strTime = Trim(dtpTime.value)
    
     '删除客户端升级日期项目
    If blnDelete = True Then
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select 项目,内容 From zlRegInfo where 项目 = '客户端升级日期'"
        Call OpenRecordset(rsTmp, strSQL, Me.Caption)
        
        If rsTmp.EOF = False Then
            strSQL = "Delete from zltools.zlRegInfo Where 项目='客户端升级日期'"
            gcnOracle.Execute strSQL
        End If
        SaveUpgradeDate = True
        optUpgradeTime.Item(0).SetFocus
        Exit Function
    End If
    
    '正常存储
    Set rsTmp = New ADODB.Recordset
    strSQL = "Select 项目,内容 From zlRegInfo where 项目 = '客户端升级日期'"
    Call OpenRecordset(rsTmp, strSQL, Me.Caption)
        
    If rsTmp.EOF = False Then
        strSQL = "Update zlRegInfo Set 内容='" & strTime & "' Where 项目='客户端升级日期'"
        gcnOracle.Execute strSQL
    Else
        strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('客户端升级日期',Null,'" & strTime & "')"
        gcnOracle.Execute strSQL
    End If
    
'    MsgBox "设置指定升级时间完成!", vbInformation, gstrSysName

    If optUpgradeTime.Item(1).Visible = True Then optUpgradeTime.Item(1).SetFocus
    
    SaveUpgradeDate = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub FilterData(Optional strLocationClientsName As String)
    Dim strFind As String
    Dim strCondition As String
    Dim lngSelectRow As Long
    Dim i As Long
    
    On Error GoTo errH:
    strFind = IIf(txtFind.Tag <> txtFind.Text, txtFind.Text, "")
    strCondition = Decode(lbloptStatus.Tag, "4", "", "0", "未升级", "2", "失败", "")
    
    mlngClinetNum = 0
    mlngNotUpClinetNum = 0
    mlngUpFailClinetNum = 0
    
    With vsfMain
        If .Rows < 1 Then Exit Sub
        lngSelectRow = .Row
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, Col_升级结果) = strCondition Or strCondition = "") And (InStr(Trim(.TextMatrix(i, Col_IP)), strFind) > 0 Or InStr(Trim(.TextMatrix(i, Col_客户端)), strFind) > 0 Or InStr(Trim(.TextMatrix(i, Col_客户端)), UCase(strFind)) > 0 Or InStr(Trim(.TextMatrix(i, Col_部门)), strFind) > 0 Or InStr(Trim(.TextMatrix(i, Col_用途)), strFind) > 0) Then
                .RowHidden(i) = False
                mlngClinetNum = mlngClinetNum + 1
                Select Case .TextMatrix(i, Col_升级结果)
                    Case "未升级"
                        mlngNotUpClinetNum = mlngNotUpClinetNum + 1
                    Case "失败"
                        mlngUpFailClinetNum = mlngUpFailClinetNum + 1
                End Select
            Else
                .RowHidden(i) = True
            End If
        Next
        
        lngSelectRow = .FindRow(strLocationClientsName, , Col_客户端)
        If lngSelectRow > 0 Then
            If .RowHidden(lngSelectRow) = False Then
                .Row = lngSelectRow
            Else
                For i = 1 To .Rows - 1
                    If .RowHidden(i) = False Then .Row = i: Exit For
                Next
                If i = .Rows Then .Row = 0
            End If
        Else
            For i = 1 To .Rows - 1
                If .RowHidden(i) = False Then .Row = i: Exit For
            Next
            If i = .Rows Then .Row = 0
        End If
        vsfMain_AfterRowColChange -1, -1, .Row, Col_客户端
        .ShowCell .Row, 0
        .Redraw = flexRDBuffered
    End With
    mblnFilter = True
    SetMenu
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub
Private Sub InitVsfMain()
    With vsfMain
        .Rows = .FixedRows
        .Cols = Col_列数

        .Cell(flexcpText, 0, Col_客户端) = "客户端"
        .Cell(flexcpAlignment, 0, Col_客户端) = flexAlignCenterCenter
        .ColWidth(Col_客户端) = 1800
        
        .Cell(flexcpText, 0, Col_IP) = "IP"
        .Cell(flexcpAlignment, 0, Col_IP) = flexAlignCenterCenter
        .ColWidth(Col_IP) = 1500
        
        .Cell(flexcpText, 0, Col_部门) = "部门"
        .Cell(flexcpAlignment, 0, Col_部门) = flexAlignCenterCenter
        .ColWidth(Col_部门) = 1000
        
        .Cell(flexcpText, 0, Col_用途) = "用途"
        .Cell(flexcpAlignment, 0, Col_用途) = flexAlignCenterCenter
        .ColWidth(Col_用途) = 1000
        
        .Cell(flexcpText, 0, Col_升级服务器) = "升级服务器"
        .Cell(flexcpAlignment, 0, Col_升级服务器) = flexAlignCenterCenter
        .ColWidth(Col_升级服务器) = 3200
        
        .Cell(flexcpText, 0, Col_更新检查) = "更新检查"
        .Cell(flexcpAlignment, 0, Col_更新检查) = flexAlignCenterCenter
        .ColWidth(Col_更新检查) = 900
        .ColHidden(Col_更新检查) = True

        .Cell(flexcpText, 0, Col_升级) = "升级"
        .Cell(flexcpAlignment, 0, Col_升级) = flexAlignCenterCenter
        .ColWidth(Col_升级) = 700
        
        .Cell(flexcpText, 0, Col_预升级) = "预升级"
        .Cell(flexcpAlignment, 0, Col_预升级) = flexAlignCenterCenter
        .ColWidth(Col_预升级) = 900
        
        .Cell(flexcpText, 0, Col_收集) = "收集"
        .Cell(flexcpAlignment, 0, Col_收集) = flexAlignCenterCenter
        .ColWidth(Col_收集) = 700
        .ColHidden(Col_收集) = True
        
        .Cell(flexcpText, 0, Col_预升级时点) = "预升级时间"
        .Cell(flexcpAlignment, 0, Col_预升级时点) = flexAlignCenterCenter
        .ColWidth(Col_预升级时点) = 1000

        .Cell(flexcpText, 0, Col_升级结果) = "升级结果"
        .Cell(flexcpAlignment, 0, Col_升级结果) = flexAlignCenterCenter
        .ColWidth(Col_升级结果) = 1800

        .Cell(flexcpText, 0, Col_预升级结果) = "预升级结果"
        .Cell(flexcpAlignment, 0, Col_预升级结果) = flexAlignCenterCenter
        .ColWidth(Col_预升级结果) = 1200

'        .Cell(flexcpText, 0, Col_收集结果) = "收集结果"
'        .Cell(flexcpAlignment, 0, Col_收集结果) = flexAlignCenterCenter
'        .ColWidth(Col_收集结果) = 1300
        
        .Cell(flexcpText, 0, Col_主动修复结果) = "主动修复结果"
        .Cell(flexcpAlignment, 0, Col_主动修复结果) = flexAlignCenterCenter
        .ColWidth(Col_主动修复结果) = 1200

        .Cell(flexcpText, 0, Col_升级说明) = "升级说明"
        .ColWidth(Col_升级说明) = 10
        .ColHidden(Col_升级说明) = True
        
        .Cell(flexcpText, 0, Col_预升级说明) = "预升级说明"
        .ColWidth(Col_预升级说明) = 10
        .ColHidden(Col_预升级说明) = True
        
'        .Cell(flexcpText, 0, Col_收集说明) = "收集说明"
'        .ColWidth(Col_收集说明) = 10
'        .ColHidden(Col_收集说明) = True
        
        .Cell(flexcpText, 0, Col_主动修复说明) = "主动修复说明"
        .ColWidth(Col_主动修复说明) = 10
        .ColHidden(Col_主动修复说明) = True
        
        .Cell(flexcpText, 0, Col_管理员) = "管理员"
        .ColWidth(Col_管理员) = 1025
        .Cell(flexcpAlignment, 0, Col_管理员) = flexAlignCenterCenter
        
        
        .Cell(flexcpText, 0, Col_密码) = "密码"
        .ColWidth(Col_密码) = 1025
        .Cell(flexcpAlignment, 0, Col_密码) = flexAlignCenterCenter
        '选中框风格
        .FocusRect = flexFocusSolid
        '最后一列自动列宽
'        .ExtendLastCol = True
        '滚动画面跟随
        .ScrollTrack = True
        '自动换行
        .WordWrap = True
        '行高设置
        .RowHeightMin = 300
        .RowHeightMax = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
'        Call SetMenu
    End With
End Sub

Public Sub RefreshData()
    Call LoadClientsData
    Call FilterData(mstrLocationClientsName)
End Sub

Public Sub SetControlEnable(ByVal strProgFunc As String)
'根据权限字符串设置控件状态
'strProgFunc:权限字符串
    Dim arrFunc() As String
    Dim i As Long
    
    mblnAllowEdit = False
    arrFunc = Split(strProgFunc, "|")
    For i = 0 To UBound(arrFunc)
        If arrFunc(i) = "客户端升级配置" Then
            mblnAllowEdit = True
        End If
    Next
    '若没有权限，则将一些控件设为不可用
    If mblnAllowEdit = False Then
        chkClientRepair.Enabled = False
        chkAllUpgrade.Enabled = False
        chkAllBefUpgrade.Enabled = False
        cmdAllCollect.Enabled = False
        txtExplain(0).Enabled = False
        txtExplain(1).Enabled = False
        txtExplain(2).Enabled = False
        txtExplain(5).Enabled = False
        cmdFileSeverSet.Enabled = False
        cmdClientAaminSet.Enabled = False
        optUpgradeTime(0).Enabled = False
        optUpgradeTime(1).Enabled = False
        cmdTimeSet.Enabled = False
        cmdClientModify.Enabled = False
        vsfMain.Editable = flexEDNone
        cmdkillProcess.Enabled = False
    End If
End Sub
