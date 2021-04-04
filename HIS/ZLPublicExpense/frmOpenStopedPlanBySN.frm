VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmOpenStopedPlanBySN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "开放停诊号源"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpenStopedPlanBySN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleMode       =   0  'User
   ScaleWidth      =   8891.551
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSignalSourceSelect 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   4935
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1170
      Width           =   4935
      Begin VB.ComboBox cboDept 
         Height          =   330
         Left            =   525
         TabIndex        =   6
         Top             =   50
         Width           =   1785
      End
      Begin VB.ComboBox cboDoctor 
         Height          =   330
         ItemData        =   "frmOpenStopedPlanBySN.frx":000C
         Left            =   2985
         List            =   "frmOpenStopedPlanBySN.frx":000E
         TabIndex        =   7
         Top             =   50
         Width           =   1785
      End
      Begin VB.Label lblDeptFilter 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   210
         Left            =   60
         TabIndex        =   30
         Top             =   110
         Width           =   420
      End
      Begin VB.Label lblDoctorFilter 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         Height          =   210
         Left            =   2520
         TabIndex        =   29
         Top             =   105
         Width           =   420
      End
   End
   Begin VB.Frame fraSplitY 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   4440
      TabIndex        =   38
      Top             =   6570
      Width           =   2385
   End
   Begin VB.Frame fraRecordInfo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3570
      TabIndex        =   34
      Top             =   2190
      Width           =   7005
      Begin VB.TextBox txt停诊时间 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   0
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   50
         Width           =   1815
      End
      Begin VB.TextBox txt限号数 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   50
         Width           =   1005
      End
      Begin VB.TextBox txt限约数 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   0
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   50
         Width           =   1005
      End
      Begin VB.Label lbl停诊时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "停诊时间"
         Height          =   210
         Index           =   0
         Left            =   4290
         TabIndex        =   37
         Top             =   110
         Width           =   840
      End
      Begin VB.Label lbl限约数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "限约数"
         Height          =   210
         Index           =   0
         Left            =   2280
         TabIndex        =   36
         Top             =   110
         Width           =   630
      End
      Begin VB.Label lbl限号数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "限号数"
         Height          =   210
         Index           =   0
         Left            =   300
         TabIndex        =   35
         Top             =   110
         Width           =   630
      End
   End
   Begin VB.Frame fraSplit 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3750
      MousePointer    =   9  'Size W E
      TabIndex        =   33
      Top             =   6420
      Width           =   67
   End
   Begin VB.PictureBox picVisitDate 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6210
      ScaleHeight     =   405
      ScaleWidth      =   4725
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1350
      Width           =   4725
      Begin VB.TextBox txtWorkTime 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   50
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Left            =   930
         TabIndex        =   9
         Top             =   50
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   160104451
         CurrentDate     =   42713
      End
      Begin VB.Label lblWorkTime 
         AutoSize        =   -1  'True
         Caption         =   "上班时段"
         Height          =   210
         Left            =   2610
         TabIndex        =   32
         Top             =   110
         Width           =   840
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "出诊日期"
         Height          =   210
         Left            =   30
         TabIndex        =   31
         Top             =   110
         Width           =   840
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSignalSource 
      Height          =   4845
      Left            =   0
      TabIndex        =   11
      Top             =   2070
      Width           =   3465
      _cx             =   6112
      _cy             =   8546
      Appearance      =   2
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOpenStopedPlanBySN.frx":0010
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
   Begin VB.Frame fraSignalSource 
      Caption         =   "号源信息"
      Height          =   1095
      Left            =   0
      TabIndex        =   22
      Top             =   30
      Width           =   11895
      Begin VB.TextBox txt号类 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   285
         Width           =   1275
      End
      Begin VB.TextBox txtSignalNO 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   285
         Width           =   1035
      End
      Begin VB.TextBox txtDoctor 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   8910
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   285
         Width           =   2175
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   285
         Width           =   2625
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   675
         Width           =   3045
      End
      Begin VB.Label lblSignalNO 
         AutoSize        =   -1  'True
         Caption         =   "号码"
         Height          =   210
         Left            =   300
         TabIndex        =   27
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lbl号类 
         AutoSize        =   -1  'True
         Caption         =   "号类"
         Height          =   210
         Left            =   2070
         TabIndex        =   26
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "项目"
         Height          =   210
         Left            =   300
         TabIndex        =   25
         Top             =   735
         Width           =   420
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   210
         Left            =   4710
         TabIndex        =   24
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         Height          =   210
         Left            =   8460
         TabIndex        =   23
         Top             =   345
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   380
      Left            =   9555
      TabIndex        =   21
      Top             =   6990
      Width           =   1250
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   405
      Left            =   3570
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1710
      Width           =   2385
      _Version        =   589884
      _ExtentX        =   4207
      _ExtentY        =   714
      _StockProps     =   64
   End
   Begin VB.PictureBox picTimeWork 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   0
      Left            =   3570
      ScaleHeight     =   4065
      ScaleWidth      =   8385
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8385
      Begin VB.TextBox txtOpen 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   930
         TabIndex        =   18
         Text            =   "0"
         Top             =   3690
         Width           =   1170
      End
      Begin MSComCtl2.UpDown updOpen 
         Height          =   330
         Index           =   0
         Left            =   2130
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3690
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         BuddyControl    =   "txtOpen(0)"
         BuddyDispid     =   196641
         BuddyIndex      =   0
         OrigLeft        =   2355
         OrigTop         =   3690
         OrigRight       =   2610
         OrigBottom      =   4020
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfTimeWork 
         Height          =   3210
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   420
         Width           =   8220
         _cx             =   14499
         _cy             =   5662
         Appearance      =   2
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
         GridColor       =   12632256
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOpenStopedPlanBySN.frx":00EC
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
         PicturesOver    =   -1  'True
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
      Begin VB.Label lblToolTip 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   0
         Left            =   2460
         TabIndex        =   39
         Top             =   3750
         Width           =   420
      End
      Begin VB.Label lblOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开放数量"
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   28
         Top             =   3750
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   380
      Left            =   8250
      TabIndex        =   20
      Top             =   6990
      Width           =   1250
   End
End
Attribute VB_Name = "frmOpenStopedPlanBySN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数
Private mlngModule As Long
Private mlngDeptID As Long, mlngDoctorID As Long
Private mlng记录ID As Long '号源ID

Private mblnOK As Boolean, mblnFirst As Boolean
Private msngStartX As Single    '移动前鼠标的位置

'挂号序号状态：0-待挂或待预约的号;1-已挂;2-已经预约;3-预留号;4-已经退号;5-已经锁号;6-已停诊
Private Enum SNState
    正常 = 0
    已挂 = 1
    已约 = 2
    预留 = 3
    退号 = 4
    锁号 = 5
    停诊 = 6
End Enum

Private mrsDept As ADODB.Recordset, mrsDoctor As ADODB.Recordset
Private mrsRecord As ADODB.Recordset, mrsRecordCount As ADODB.Recordset
Private mblnNotChange As Boolean, mblnChanged As Boolean
Private mblnCboClick As Boolean     '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
'                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
Private mlngPreRow As Long

Public Function ShowMe(ByVal frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal lng记录ID As Long, _
    Optional ByVal lngDeptID As Long, _
    Optional ByVal lngDoctorID As Long) As Boolean
    '程序入口，开放停诊序号，安排必须是启用了序号控制且分时段的
    '入参：
    '   frmMain 调用的主窗体
    '   lngModule 调用模块号
    '   lng记录ID 记录ID,1114模块调用时传入
    '   lngDeptID 科室ID
    '   lngDoctorID 医生ID
    '说明：
    '   lngModule不等于1114时，
    '   1.如果传入了医生ID,则科室只能选择操作员所属科室,医生不能编辑
    '   2.如果传入了科室ID,则科室不能编辑
    '   日期缺省为当前日期
    mlngModule = lngModule
    mlngDeptID = lngDeptID: mlngDoctorID = lngDoctorID
    mlng记录ID = lng记录ID

    On Error Resume Next
    mblnOK = False
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    ShowMe = mblnOK
End Function

Private Sub cboDept_Click()
    Dim lngDept As Long, lngDoctor As Long
    
    On Error GoTo ErrHandler
    mblnCboClick = True
    If cboDept.ListIndex <> -1 Then
        lngDept = cboDept.ItemData(cboDept.ListIndex)
    End If
    If mlngDoctorID = 0 Then Call FillDoctor(lngDept, mlngDoctorID)
    If cboDoctor.ListIndex <> -1 Then
        lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), lngDept, lngDoctor
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub FillDoctor(Optional lng科室id As Long, Optional ByVal lngDefault As Long)
'功能：根据指定的开单科室ID读取并填写医生列表,缺省医生
    Dim strOldID As String
    
    On Error GoTo ErrHandler
    cboDoctor.Clear
    If mrsDoctor Is Nothing Then Set mrsDoctor = GetAllDoctor
    mrsDoctor.Filter = ""
    If lng科室id <> 0 Then
        mrsDoctor.Filter = "部门ID=" & lng科室id
    End If
    
    Do While Not mrsDoctor.EOF
        If InStr("," & strOldID & ",", "," & Val(Nvl(mrsDoctor!ID)) & ",") = 0 Then
            cboDoctor.AddItem Nvl(mrsDoctor!简码) & "-" & Nvl(mrsDoctor!姓名)
            cboDoctor.ItemData(cboDoctor.NewIndex) = Val(Nvl(mrsDoctor!ID))
            If lngDefault = Val(Nvl(mrsDoctor!ID)) Then
                cboDoctor.ListIndex = cboDoctor.NewIndex
            End If
            strOldID = strOldID & "," & Val(Nvl(mrsDoctor!ID))
        End If
        mrsDoctor.MoveNext
    Loop
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDept_GotFocus()
    gobjControl.TxtSelAll cboDept
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim lngDoctor As Long
    
    On Error GoTo ErrHandler
    mblnCboClick = False
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mblnCboClick = True
    If Trim(cboDept.Text) = "" Then
        Call FillDoctor(, mlngDoctorID)
        If cboDoctor.ListIndex <> -1 Then
            lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
        End If
        LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), , lngDoctor
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cboDept.ListIndex < 0 Then
        If mrsDept Is Nothing Then Call FillDept(mlngDeptID)
        If zlSelectDept(Me, mlngModule, cboDept, mrsDept, cboDept.Text) = False Then
            KeyAscii = 0: mblnCboClick = False
            Exit Sub
        End If
    Else
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub FillDept(Optional ByVal lngDefault As Long)
    '功能：读取并加载科室列表,缺省科室
    '参数：lngDefault - 缺省科室ID
    Dim strSQL As String, strOldID As String
    
    On Error GoTo ErrHandler
    cboDept.Clear
    If mrsDept Is Nothing Then
        Set mrsDept = GetDepartments("临床", "1,3", mlngDoctorID)
    End If
    mrsDept.Filter = ""
    Do While Not mrsDept.EOF
        If InStr("," & strOldID & ",", "," & Val(Nvl(mrsDept!ID)) & ",") = 0 Then '一个部门可能同时属于产科和临床,不加载相同的
            cboDept.AddItem Nvl(mrsDept!名称)
            cboDept.ItemData(cboDept.NewIndex) = Val(Nvl(mrsDept!ID))
            If lngDefault = Val(Nvl(mrsDept!ID)) Then
                cboDept.ListIndex = cboDept.NewIndex
            End If
            strOldID = strOldID & "," & Val(Nvl(mrsDept!ID))
        End If
        mrsDept.MoveNext
    Loop
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    Dim Index As Integer
    
    On Error GoTo ErrHandler
    If cboDept.Text = "" Then
        cboDept.ListIndex = -1
    Else
        Index = SeekCboIndex(cboDept, NeedName(cboDept.Text))
        If Index = -1 Then
            cboDept.ListIndex = -1: cboDept.Text = ""
        ElseIf cboDept.ListIndex <> Index Then
            cboDept.ListIndex = Index
        End If
    End If
    If mblnCboClick = False Then cboDept_Click
    mblnCboClick = False
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDoctor_Click()
    Dim lngDept As Long, lngDoctor As Long
    
    On Error GoTo ErrHandler
    mblnCboClick = True
    If cboDept.ListIndex <> -1 Then
        lngDept = cboDept.ItemData(cboDept.ListIndex)
    End If
    If cboDoctor.ListIndex <> -1 Then
        lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), lngDept, lngDoctor
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDoctor_GotFocus()
    gobjControl.TxtSelAll cboDoctor
End Sub

Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngDept As Long
    
    On Error GoTo ErrHandler
    mblnCboClick = False
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mblnCboClick = True
    If Trim(cboDoctor.Text) = "" Then
        cboDoctor.ListIndex = -1
        If cboDept.ListIndex <> -1 Then
            lngDept = cboDept.ItemData(cboDept.ListIndex)
        End If
        LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), lngDept
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cboDoctor.ListIndex < 0 Then
        If mrsDoctor Is Nothing Then Call FillDoctor(, mlngDoctorID)
        If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, cboDoctor.Text) = False Then
            KeyAscii = 0: mblnCboClick = False
            Exit Sub
        End If
    Else
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
    Dim Index As Integer
    
    On Error GoTo ErrHandler
    If cboDoctor.Text = "" Then
        cboDoctor.ListIndex = -1
    Else
        Index = SeekCboIndex(cboDoctor, NeedName(cboDoctor.Text))
        If Index = -1 Then
            cboDoctor.ListIndex = -1: cboDoctor.Text = ""
        ElseIf cboDoctor.ListIndex <> Index Then
            cboDoctor.ListIndex = Index
        End If
    End If
    If mblnCboClick = False Then cboDept_Click
    mblnCboClick = False
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer, cllSql As Collection
    Dim blnTrans As Boolean
    Dim strIDs As String
    
    On Error GoTo ErrHandler
    If mblnChanged = False Then
        MsgBox "您还未调整任何安排的开放数量，不能保存！", vbInformation, gstrSysName
        Exit Sub
    End If
    If mlngModule = 1114 Then
        If txtOpen(0).Enabled Then
            strSQL = "Select Nvl(a.是否序号控制, 0) * Nvl(a.是否分时段, 0) As 启用序号时段," & vbNewLine & _
                    "        Decode(a.停诊开始时间, Null, 0, 1) As 已停诊," & vbNewLine & _
                    "        Decode(Sign(Nvl(a.停诊开始时间, Sysdate) - Sysdate), -1, 1, 0) As 超时" & vbNewLine & _
                    " From 临床出诊记录 A" & vbNewLine & _
                    " Where a.Id = [1]"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查安排", mlng记录ID)
            If rsTemp.EOF Then
                MsgBox "未找到安排数据，不能开放停诊安排！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!启用序号时段)) = 0 Then
                MsgBox "出诊安排不是启用了序号且启用时段的，不能开放停诊安排！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!已停诊)) = 0 Then
                MsgBox "当前上班时段无停诊安排，不能调整开放数量！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!超时)) = 1 Then
                MsgBox "当前时间已大于了停诊开始时间，不能再开放停诊安排！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '"       And a.开始时间 <> a.终止时间" '开始时间与终止时间相等的是加号的序号
            strSQL = "Select Sum(Decode(a.是否停诊, 1, 0, 1) * Decode(Nvl(a.挂号状态, 0), 0, 0, 1)) As 最小数量," & vbNewLine & _
                    "        Count(1) As 最大数量" & vbNewLine & _
                    " From 临床出诊序号控制 A, 临床出诊记录 B" & vbNewLine & _
                    " Where a.记录id = b.Id And b.Id = [1] And a.开始时间 Between b.停诊开始时间 And b.停诊终止时间" & vbNewLine & _
                    "       And a.开始时间 <> a.终止时间"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查安排", mlng记录ID)
            If rsTemp.EOF Then
                MsgBox "未找到安排的停诊时间范围内的序号时段数据，不能开放停诊安排！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!最小数量)) > Val(txtOpen(0).Text) Then
                MsgBox "开放数量不能小于最小开放数量" & Val(Nvl(rsTemp!最小数量)) & "！", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(Nvl(rsTemp!最大数量)) < Val(txtOpen(0).Text) Then
                MsgBox "开放数量不能小于最小开放数量" & Val(Nvl(rsTemp!最大数量)) & "", vbInformation, gstrSysName
                Exit Sub
            End If
        
            'Procedure Zl_临床出诊序号控制_开放挂号(
            strSQL = "Zl_临床出诊序号控制_开放挂号("
            '  记录id_In 临床出诊记录.Id%Type,
            strSQL = strSQL & mlng记录ID & ","
            '  数量_In Number
            strSQL = strSQL & Val(txtOpen(0).Text) & ")"
            
            gobjDatabase.ExecuteProcedure strSQL, Me.Caption
            mblnOK = True
            mblnChanged = False
            Unload Me
        End If
    Else
        For i = 0 To tbPage.ItemCount - 1
            If txtOpen(i).Enabled <> 0 Then
                strIDs = strIDs & "," & Val(tbPage(i).Tag)
            End If
        Next
        If strIDs = "" Then
            MsgBox "当前没有调整任何安排的开放数量，无需保存！", vbInformation, gstrSysName
            Exit Sub
        Else
            strIDs = Mid(strIDs, 2)
        End If
        
        strSQL = "Select a.ID As 记录ID, Nvl(a.是否序号控制, 0) * Nvl(a.是否分时段, 0) As 启用序号时段," & vbNewLine & _
                "        Decode(a.停诊开始时间, Null, 0, 1) As 已停诊," & vbNewLine & _
                "        Decode(Sign(Nvl(a.停诊开始时间, Sysdate) - Sysdate), -1, 1, 0) As 超时" & vbNewLine & _
                " From 临床出诊记录 A, Table(f_Num2list([1])) B" & vbNewLine & _
                " Where a.Id = b.Column_Value"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查安排", strIDs)
        For i = 0 To tbPage.ItemCount - 1
            If txtOpen(i).Enabled Then
                rsTemp.Filter = "记录ID=" & Val(tbPage(i).Tag)
                If rsTemp.EOF Then
                    MsgBox "未找到[" & tbPage(i).Caption & "]安排数据，不能开放停诊安排！", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!启用序号时段)) = 0 Then
                    MsgBox "[" & tbPage(i).Caption & "]出诊安排不是启用了序号且启用时段的，不能开放停诊安排！", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!已停诊)) = 0 Then
                    MsgBox "[" & tbPage(i).Caption & "]时段无停诊安排，不能调整开放数量！", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!超时)) = 1 Then
                    MsgBox "当前时间已大于了[" & tbPage(i).Caption & "]停诊开始时间，不能再开放停诊安排！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '"       And a.开始时间 <> a.终止时间" '开始时间与终止时间相等的是加号的序号
        strSQL = "Select a.记录id, Sum(Decode(a.是否停诊, 1, 0, 1) * Decode(Nvl(a.挂号状态, 0), 0, 0, 1)) As 最小数量," & vbNewLine & _
                "        Count(1) As 最大数量" & vbNewLine & _
                " From 临床出诊序号控制 A, 临床出诊记录 B, Table(f_Num2list([1])) C" & vbNewLine & _
                " Where a.记录id = b.Id And b.Id = c.Column_Value" & vbNewLine & _
                "       And a.开始时间 Between b.停诊开始时间 And b.停诊终止时间" & vbNewLine & _
                "       And a.开始时间 <> a.终止时间" & vbNewLine & _
                " Group By a.记录id"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "检查安排序号", strIDs)
        For i = 0 To tbPage.ItemCount - 1
            If txtOpen(i).Enabled Then
                rsTemp.Filter = "记录ID=" & Val(tbPage(i).Tag)
                If rsTemp.EOF Then
                    MsgBox "未找到[" & tbPage(i).Caption & "]安排的停诊时间范围内的序号时段数据，不能开放停诊安排！", vbInformation, gstrSysName
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!最小数量)) > Val(txtOpen(0).Text) Then
                    MsgBox "[" & tbPage(i).Caption & "]的开放数量不能小于最小开放数量" & Val(Nvl(rsTemp!最小数量)) & "！", vbInformation, gstrSysName
                    tbPage(i).Selected = True
                    If txtOpen(i).Visible And txtOpen(i).Enabled Then txtOpen(i).SetFocus
                    Exit Sub
                End If
                If Val(Nvl(rsTemp!最大数量)) < Val(txtOpen(0).Text) Then
                    MsgBox "[" & tbPage(i).Caption & "]的开放数量不能小于最小开放数量" & Val(Nvl(rsTemp!最大数量)) & "", vbInformation, gstrSysName
                    tbPage(i).Selected = True
                    If txtOpen(i).Visible And txtOpen(i).Enabled Then txtOpen(i).SetFocus
                    Exit Sub
                End If
            End If
        Next
        
        Set cllSql = New Collection
        For i = 0 To tbPage.ItemCount - 1
            If txtOpen(i).Enabled Then
                'Procedure Zl_临床出诊序号控制_开放挂号(
                strSQL = "Zl_临床出诊序号控制_开放挂号("
                '  记录id_In 临床出诊记录.Id%Type,
                strSQL = strSQL & Val(tbPage(i).Tag) & ","
                '  数量_In Number
                strSQL = strSQL & Val(txtOpen(i).Text) & ")"
            
                zlAddArray cllSql, strSQL
            End If
        Next
        If cllSql.Count > 0 Then
            blnTrans = True
            zlExecuteProcedureArrAy cllSql, Me.Caption
            blnTrans = False
            mblnOK = True
            mblnChanged = False
            Unload Me
        End If
    End If
    Exit Sub
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        'Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub dtpDate_Change()
    Dim lngDept As Long, lngDoctor As Long
    
    On Error GoTo ErrHandler
    If cboDept.ListIndex <> -1 Then
        lngDept = cboDept.ItemData(cboDept.ListIndex)
    End If
    If cboDoctor.ListIndex <> -1 Then
        lngDoctor = cboDoctor.ItemData(cboDoctor.ListIndex)
    End If
    LoadSignalSource Format(dtpDate.Value, "yyyy-MM-dd"), lngDept, lngDoctor
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    On Error GoTo ErrHandler
    If cboDept.Visible And cboDept.Enabled Then
        cboDept.SetFocus
    ElseIf cboDoctor.Visible And cboDoctor.Enabled Then
        cboDoctor.SetFocus
    ElseIf dtpDate.Visible And dtpDate.Enabled Then
        dtpDate.SetFocus
    ElseIf txtOpen(0).Visible And txtOpen(0).Enabled Then
        txtOpen(0).SetFocus
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Function InitFace(ByVal lngModule As Long) As Boolean
    '初始化界面
    Err = 0: On Error GoTo ErrHandler
    Select Case lngModule
    Case 1114 '临床出诊安排
        picSignalSourceSelect.Visible = False
        vsfSignalSource.Visible = False
        fraSplit.Visible = False
        tbPage.Visible = False
        
        dtpDate.Enabled = False
        LoadControl 0
    Case Else
        fraSignalSource.Visible = False
        fraSplitY.Visible = False
        lblWorkTime.Visible = False: txtWorkTime.Visible = False
        Set fraRecordInfo(0).Container = picTimeWork(0)
        
        With tbPage
            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
            .PaintManager.BoldSelected = True
            .PaintManager.Layout = xtpTabLayoutAutoSize
            .PaintManager.StaticFrame = True
            .PaintManager.ClientFrame = xtpTabFrameBorder
            
            .InsertItem 0, "无上班时段", picTimeWork(0).hWnd, 0
        End With
    End Select
    InitFace = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Form_Load()
    mblnFirst = True
    If InitFace(mlngModule) = False Then Unload Me: Exit Sub
    If mlngModule = 1114 Then
        If LoadData(mlng记录ID) = False Then Unload Me: Exit Sub
    Else
        cboDept.Enabled = mlngDeptID = 0
        cboDoctor.Enabled = mlngDoctorID = 0
        
        Call FillDept(mlngDeptID) '如果传入了医生则只能选择人员所属科室
        Call FillDoctor(, mlngDoctorID)
        dtpDate.Value = Format(gobjDatabase.CurrentDate(), "yyyy-MM-dd")
        dtpDate.minDate = dtpDate.Value
        Call dtpDate_Change
    End If
End Sub

Private Function LoadSignalSource(ByVal str出诊日期 As String, _
    Optional ByVal lngDeptID As Long, _
    Optional ByVal lngDoctorID As Long) As Boolean
    '加载号源数据
    '入参：
    '   str出诊日期 格式：yyyy-mm-dd
    '   lngDeptID 科室ID
    '   lngDoctorID 医生ID
    Dim strSQL As String, strWhere As String
    Dim i As Integer, strIDs As String
    Dim lngRow As Long

    Err = 0: On Error GoTo ErrHandler
    vsfSignalSource.Clear 1: vsfSignalSource.Rows = 1
    For i = picTimeWork.UBound To 1 Step -1
        tbPage.RemoveItem i
        UnLoadControl i
    Next
    tbPage(0).Caption = "无上班时段"
    UnLoadControl 0
    mlngPreRow = -1
    
    Set mrsRecord = Nothing: Set mrsRecordCount = Nothing
    '号源信息
    If lngDeptID <> 0 Then strWhere = " And a.科室ID=[1]"
    If lngDoctorID <> 0 Then strWhere = strWhere & " And a.医生ID=[2]"
    strSQL = "Select a.号源id, b.号码, b.号类, m.名称 As 科室, a.医生姓名 As 医生, n.名称 As 收费项目," & vbNewLine & _
            "        a.Id As 记录id, a.出诊日期, a.上班时段, a.限号数, a.限约数, a.预约控制," & vbNewLine & _
            "        a.停诊开始时间, a.停诊终止时间" & vbNewLine & _
            " From 临床出诊记录 A, 临床出诊号源 B, 部门表 M, 收费项目目录 N" & vbNewLine & _
            " Where a.号源id = b.Id And a.科室id = m.Id And a.项目id = n.Id And a.出诊日期 = [3] And" & vbNewLine & _
            "       Nvl(a.是否序号控制, 0) = 1 And Nvl(a.是否分时段, 0) = 1" & strWhere & vbNewLine & _
            "       And (b.撤档时间 Is Null Or b.撤档时间>=To_Date('3000-01-01','yyyy-mm-dd'))" & vbNewLine & _
            "       And (m.站点='" & gstrNodeNo & "' Or m.站点 is Null)"
    Set mrsRecord = gobjDatabase.OpenSQLRecord(strSQL, "获取号源信息", lngDeptID, lngDoctorID, CDate(str出诊日期))
    If mrsRecord.EOF Then Set mrsRecord = Nothing: Exit Function
    
    '"       And b.开始时间 <> b.终止时间" '开始时间与终止时间相等的是加号的序号
    strSQL = "Select b.记录ID, Nvl(Max(Decode(b.是否停诊, 1, 0, b.序号)) - Min(b.序号) + 1, 0) As 停诊范围," & vbNewLine & _
            "        Sum(Decode(b.是否停诊, 1, 0, 1) * Decode(Nvl(b.挂号状态, 0), 0, 0, 1)) As 最小数量," & vbNewLine & _
            "        Count(1) As 最大数量," & vbNewLine & _
            "        Sum(Decode(b.是否停诊, 1, 0, 1)) As 上次开放数量" & vbNewLine & _
            " From 临床出诊序号控制 B, 临床出诊记录 A, 临床出诊号源 M, 部门表 N" & vbNewLine & _
            " Where b.记录id = a.Id And a.号源id = m.Id And a.科室id = n.Id" & vbNewLine & _
            "       And b.开始时间 Between a.停诊开始时间 And a.停诊终止时间" & vbNewLine & _
            "       And b.开始时间 <> b.终止时间" & vbNewLine & _
            "       And a.出诊日期 = [3] And Nvl(a.是否序号控制, 0) = 1 And Nvl(a.是否分时段, 0) = 1" & strWhere & vbNewLine & _
            "       And (m.撤档时间 Is Null Or m.撤档时间>=To_Date('3000-01-01','yyyy-mm-dd'))" & vbNewLine & _
            "       And (n.站点='" & gstrNodeNo & "' Or n.站点 is Null)" & vbNewLine & _
            " Group By b.记录ID"
    Set mrsRecordCount = gobjDatabase.OpenSQLRecord(strSQL, "获取序号控制数量", lngDeptID, lngDoctorID, CDate(str出诊日期))
    
    With vsfSignalSource
        .Redraw = flexRDNone
        lngRow = 1
        Do While Not mrsRecord.EOF
            If InStr("," & strIDs & ",", "," & Val(Nvl(mrsRecord!号源ID)) & ",") = 0 Then
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("号源ID")) = Val(Nvl(mrsRecord!号源ID))
                .TextMatrix(lngRow, .ColIndex("号码")) = Nvl(mrsRecord!号码)
                .TextMatrix(lngRow, .ColIndex("号类")) = Nvl(mrsRecord!号类)
                .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(mrsRecord!科室)
                .TextMatrix(lngRow, .ColIndex("医生")) = Nvl(mrsRecord!医生)
                .TextMatrix(lngRow, .ColIndex("收费项目")) = Nvl(mrsRecord!收费项目)
                lngRow = lngRow + 1
            End If
            strIDs = strIDs & "," & Val(Nvl(mrsRecord!号源ID))
            mrsRecord.MoveNext
        Loop
        .Redraw = flexRDBuffered
        If .Rows > 1 Then
            mblnNotChange = True
            .Row = 1
            mblnNotChange = False
            vsfSignalSource_EnterCell
        End If
    End With
    LoadSignalSource = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function LoadData(ByVal lng记录ID As Long) As Boolean
    '加载数据
    Dim strSQL As String, rsRecord As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim lngCanOpenMax As Long, lngCanOpenMin As Long, lngPreOpen As Long

    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select b.号码, b.号类, c.名称 As 科室, a.医生姓名, d.名称 As 收费项目," & vbNewLine & _
            "        a.Id As 记录id, a.出诊日期, a.上班时段, a.限号数, a.限约数, a.预约控制," & vbNewLine & _
            "        a.停诊开始时间, a.停诊终止时间" & vbNewLine & _
            " From 临床出诊记录 A, 临床出诊号源 B, 部门表 C, 收费项目目录 D" & vbNewLine & _
            " Where a.号源ID = b.ID And a.科室id = c.Id And a.项目id = d.Id" & vbNewLine & _
            "       And a.id = [1] And Nvl(a.是否序号控制, 0) = 1 And Nvl(a.是否分时段, 0) = 1"
    Set rsRecord = gobjDatabase.OpenSQLRecord(strSQL, "获取安排", lng记录ID)
    If rsRecord.EOF Then
        MsgBox "当前安排未启用序号控制或未启用分时段，不能开放停诊安排！", vbInformation, gstrSysName
        Exit Function
    End If
    
    txtSignalNO.Text = Nvl(rsRecord!号码)
    txt号类.Text = Nvl(rsRecord!号类)
    txtDept.Text = Nvl(rsRecord!科室)
    txtDoctor.Text = Nvl(rsRecord!医生姓名)
    txtItem.Text = Nvl(rsRecord!收费项目)
    
    dtpDate.Value = Format(Nvl(rsRecord!出诊日期), "yyyy-MM-dd")
    txtWorkTime.Text = Nvl(rsRecord!上班时段)
    txt限号数(0).Text = IIf(Val(Nvl(rsRecord!限号数)) = 0, "", Nvl(rsRecord!限号数))
    txt限约数(0).Text = IIf(Val(Nvl(rsRecord!预约控制)) = 1, "禁止预约", IIf(Val(Nvl(rsRecord!限约数)) = 0, txt限号数(0).Text, Nvl(rsRecord!限约数)))
    txt停诊时间(0).Text = Format(Nvl(rsRecord!停诊开始时间), "hh:mm") & "～" & Format(Nvl(rsRecord!停诊终止时间), "hh:mm")
    txt停诊时间(0).Tag = Format(Nvl(rsRecord!停诊开始时间), "yyyy-mm-dd hh:mm") & "～" & Format(Nvl(rsRecord!停诊终止时间), "yyyy-mm-dd hh:mm")
    If txt停诊时间(0).Text = "～" Then txt停诊时间(0).Text = ""
    
    '"       And a.开始时间 <> a.终止时间" '开始时间与终止时间相等的是加号的序号
    strSQL = "Select Nvl(Max(Decode(a.是否停诊, 1, 0, a.序号)) - Min(a.序号) + 1, 0) As 停诊范围," & vbNewLine & _
            "        Sum(Decode(a.是否停诊, 1, 0, 1) * Decode(Nvl(a.挂号状态, 0), 0, 0, 1)) As 最小数量," & vbNewLine & _
            "        Count(1) As 最大数量," & vbNewLine & _
            "        Sum(Decode(a.是否停诊, 1, 0, 1)) As 上次开放数量" & vbNewLine & _
            " From 临床出诊序号控制 A, 临床出诊记录 B" & vbNewLine & _
            " Where a.记录id = b.Id And b.Id = [1]" & vbNewLine & _
            "       And a.开始时间 Between b.停诊开始时间 And b.停诊终止时间" & vbNewLine & _
            "       And a.开始时间 <> a.终止时间"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取停诊序号数量", lng记录ID)
    If Not rsTemp.EOF Then
        vsfTimeWork(0).Tag = Val(Nvl(rsTemp!停诊范围))
        lngCanOpenMax = Val(Nvl(rsTemp!最大数量))
        lngCanOpenMin = Val(Nvl(rsTemp!最小数量))
        lngPreOpen = Val(Nvl(rsTemp!上次开放数量))
    End If
    mblnNotChange = True
    updOpen(0).Max = lngCanOpenMax
    updOpen(0).Min = lngCanOpenMin
    txtOpen(0).Text = lngPreOpen
    mblnNotChange = False
    
    '序号时段
    Call LoadDataToGrid(0, Val(Nvl(rsRecord!记录ID)), lngCanOpenMax, lngCanOpenMin)
    mblnChanged = False
    LoadData = True
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function LoadDataToGrid(ByVal Index As Integer, ByVal lng记录ID As Long, _
    ByVal lngCanOpenMax As Long, ByVal lngCanOpenMin As Long) As Boolean
    '加载时段数据到网格控件
    '入参:
    '   lng记录ID - 记录ID
    '返回:加载成功，返回true,否则返回Flase
    Dim objColAll As New Collection 'Array(序号,开始时间,终止时间,挂号状态,是否停诊)
    Dim strSQL As String, rsRecord As ADODB.Recordset

    On Error GoTo ErrHander
    '必须根据时间先后进行排序，不然要乱
    '"       And a.开始时间 <> a.终止时间" '开始时间与终止时间相等的是加号的序号
    strSQL = "Select a.序号, a.开始时间, a. 终止时间, a.挂号状态, a.是否停诊" & vbNewLine & _
            " From 临床出诊序号控制 A" & vbNewLine & _
            " Where a.记录ID = [1] And a.开始时间 <> a.终止时间" & vbNewLine & _
            " Order By a.序号"
    Set rsRecord = gobjDatabase.OpenSQLRecord(strSQL, "获取号序信息", lng记录ID)
    Do While Not rsRecord.EOF
        objColAll.Add Array(Val(Nvl(rsRecord!序号)), Format(Nvl(rsRecord!开始时间), "yyyy-MM-dd hh:mm:ss"), _
            Format(Nvl(rsRecord!终止时间), "yyyy-MM-dd hh:mm:ss"), _
            Val(Nvl(rsRecord!挂号状态)), Val(Nvl(rsRecord!是否停诊)))
        rsRecord.MoveNext
    Loop
    LoadDataToGrid = ShowTimeIntervals(Index, objColAll, lngCanOpenMax, lngCanOpenMin)
    Exit Function
ErrHander:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function ShowTimeIntervals(ByVal Index As Integer, objCol As Collection, _
    ByVal lngCanOpenMax As Long, ByVal lngCanOpenMin As Long) As Boolean
    '显示时段数据
    '入参：
    '   objCol:Array(序号,开始时间,终止时间,挂号状态,是否停诊)
    Dim varItem As Variant, varTemp As Variant
    Dim i As Integer, j As Integer, blnFind As Boolean
    Dim lngRow As Long, lngCol As Long, strCurTime As String
    Dim dtSys As Date, strToolTip As String
    Dim strStopStart As String, strStopEnd As String

    Err = 0: On Error GoTo ErrHander:
    lblToolTip(Index).Caption = ""
    With vsfTimeWork(Index)
        .Clear
        .Rows = 0
        If objCol Is Nothing Then Exit Function
        If objCol.Count = 0 Then Exit Function
        
        .Redraw = flexRDNone
        dtSys = gobjDatabase.CurrentDate
        strStopStart = Split(txt停诊时间(Index).Tag & "～", "～")(0)
        strStopEnd = Split(txt停诊时间(Index).Tag & "～", "～")(1)
        If IsDate(strStopStart) Then
            If DateDiff("n", strStopStart, dtSys) > 0 Then
               '当前时间已进入停诊时间范围
               strToolTip = "当前时间已大于停诊开始时间，不能调整开放数量！"
            End If
        Else
            '当前上班时段无停诊安排
            strToolTip = "当前上班时段无停诊安排，不能调整开放数量！"
        End If
        
        .Rows = 1: .Cols = 1
        .FixedCols = 1

        lngRow = -1: lngCol = 1: strCurTime = ""
        .FontSize = 9
        For Each varItem In objCol
            If strCurTime <> Format(varItem(1), "hh:00") Then
                strCurTime = Format(varItem(1), "hh:00")
                lngRow = lngRow + 1: lngCol = 1
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, 0) = strCurTime
            End If
            If lngCol > .Cols - 1 Then .Cols = .Cols + 1
            .TextMatrix(lngRow, lngCol) = varItem(0) & vbCrLf & _
                Format(varItem(1), "hh:mm") & "-" & Format(varItem(2), "hh:mm")
            .Cell(flexcpData, lngRow, lngCol) = Format(varItem(1), "yyyy-MM-dd hh:mm:ss") & "～" & Format(varItem(2), "yyyy-MM-dd hh:mm:ss")
            
            If Format(dtSys, "yyyy-mm-dd hh:mm:ss") >= Format(varItem(1), "yyyy-mm-dd hh:mm:ss") Then
                 '已失效的用下划线和灰色字体显示
                 .Cell(flexcpFontUnderline, lngRow, lngCol) = True
                 .Cell(flexcpForeColor, lngRow, lngCol) = vbGrayText
            End If
            
            Select Case varItem(3)
            Case SNState.已挂
                .Cell(flexcpForeColor, lngRow, lngCol) = &HC0&
                .Cell(flexcpFontStrikethru, lngRow, lngCol) = True
            Case SNState.已约
                .Cell(flexcpForeColor, lngRow, lngCol) = vbGreen
            Case SNState.预留
                .Cell(flexcpForeColor, lngRow, lngCol) = vbBlue
            Case SNState.退号
                .Cell(flexcpForeColor, lngRow, lngCol) = vbGrayText
                .Cell(flexcpFontStrikethru, lngRow, lngCol) = True
            Case SNState.锁号
                .Cell(flexcpForeColor, lngRow, lngCol) = &HC0&
            End Select
            
            '判断是否是停诊范围内开放的
            If varItem(4) = 1 Then
                '已停诊用红色背景显示
                .Cell(flexcpBackColor, lngRow, lngCol) = vbRed
            End If
            
            lngCol = lngCol + 1
        Next
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = flexAlignCenterTop
        .Cell(flexcpAlignment, 0, 1, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpFontSize, 0, 0, .Rows - 1) = 12
        .Cell(flexcpFontBold, 0, 0, .Rows - 1) = True
        .ColWidth(-1) = 1100: .ColWidth(0) = 1000: .RowHeight(-1) = 600
        
        If strToolTip = "" Then
            If lngCanOpenMax = 0 Then
                strToolTip = "停诊时间范围内不存在“待挂或待预约的号”，不能调整开放数量！"
                txtOpen(Index).Enabled = False: updOpen(Index).Enabled = False
            ElseIf lngCanOpenMin > 0 Then
                strToolTip = "上次开放的数量中已有 " & lngCanOpenMin & " 个被使用，本次设置的开放数量必须大于 " & lngCanOpenMin & " ！"
            End If
        Else
            txtOpen(Index).Enabled = False: updOpen(Index).Enabled = False
        End If
        If Trim(strToolTip) = "" Then
            lblToolTip(Index).Caption = ""
        Else
            lblToolTip(Index).Caption = "提示：" & strToolTip
        End If
        .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftTop
        .Redraw = flexRDBuffered
    End With
    ShowTimeIntervals = True
    Exit Function
ErrHander:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Select Case mlngModule
    Case 1114 '临床出诊安排
        fraSignalSource.Move 0, 10, Me.ScaleWidth
        picVisitDate.Move 0, fraSignalSource.Top + fraSignalSource.Height
        fraRecordInfo(0).Move picVisitDate.Left + picVisitDate.Width, picVisitDate.Top
        With picTimeWork(0)
            .Left = 0
            .Top = picVisitDate.Top + picVisitDate.Height
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight - 800 - .Top
        End With
        fraSplitY.Move -10, picTimeWork(0).Top + picTimeWork(0).Height, Me.ScaleWidth + 20
    Case Else
        picSignalSourceSelect.Move 0, 50
        picVisitDate.Move picSignalSourceSelect.Left + picSignalSourceSelect.Width, picSignalSourceSelect.Top
        
        With vsfSignalSource
            .Left = 0
            .Top = picSignalSourceSelect.Top + picSignalSourceSelect.Height
            .Height = Me.ScaleHeight - 800 - .Top
        End With
        fraSplit.Move vsfSignalSource.Left + vsfSignalSource.Width, vsfSignalSource.Top, fraSplit.Width, Me.ScaleHeight - vsfSignalSource.Top - 800
        With tbPage
            .Left = fraSplit.Left + fraSplit.Width
            .Top = fraSplit.Top
            .Width = Me.ScaleWidth - .Left
            .Height = fraSplit.Height
        End With
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If mblnChanged Then
        If MsgBox("当前安排的开放数量已改变，但您还未保存，是否不保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    Set mrsDept = Nothing
    Set mrsDoctor = Nothing
    Set mrsRecord = Nothing: Set mrsRecordCount = Nothing
End Sub

Private Sub fraSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then msngStartX = X
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    
    On Error Resume Next
    If Button = vbLeftButton Then
        sngTemp = fraSplit.Left + X - msngStartX
        If sngTemp > 500 And Me.ScaleWidth - (sngTemp + fraSplit.Width) > 500 Then
            fraSplit.Left = sngTemp
            vsfSignalSource.Width = fraSplit.Left - vsfSignalSource.Left
            tbPage.Move fraSplit.Left + fraSplit.Width, tbPage.Top, Me.ScaleWidth - (fraSplit.Left + fraSplit.Width)
        End If
    End If
End Sub

Private Sub picTimeWork_Resize(Index As Integer)
    Dim blnRecordInfo As Boolean
    
    On Error Resume Next
    If Index = 0 And fraRecordInfo(0).Container <> picTimeWork(0) Then
        vsfTimeWork(Index).Move 0, 30, picTimeWork(Index).ScaleWidth, picTimeWork(Index).ScaleHeight - 450 - 60
    Else
        fraRecordInfo(Index).Move 0, 0
        vsfTimeWork(Index).Move 0, 420, picTimeWork(Index).ScaleWidth, picTimeWork(Index).ScaleHeight - 450 - 420
    End If
    txtOpen(Index).Top = picTimeWork(Index).ScaleHeight - txtOpen(Index).Height - 60
    updOpen(Index).Top = txtOpen(Index).Top
    lblOpen(Index).Top = txtOpen(Index).Top + (txtOpen(Index).Height - lblOpen(Index).Height) / 2
    lblToolTip(Index).Top = lblOpen(Index).Top
    lblToolTip(Index).Width = picTimeWork(Index).ScaleWidth - lblToolTip(Index).Left
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If txtOpen(Item.Index).Visible And txtOpen(Item.Index).Enabled Then txtOpen(Item.Index).SetFocus
End Sub

Private Sub txtOpen_Change(Index As Integer)
    If mblnNotChange Then Exit Sub
    
    On Error GoTo ErrHandler
    mblnChanged = True
    mblnNotChange = True
    If Trim(txtOpen(Index).Text) = "" Then txtOpen(Index).Text = "0"
    If updOpen(Index).Max < Val(txtOpen(Index).Text) Then
        MsgBox "开放数量不能大于最大可开放数量(" & updOpen(Index).Max & ")！", vbExclamation, gstrSysName
        txtOpen(Index).Text = updOpen(Index).Max
        If txtOpen(Index).Visible And txtOpen(Index).Enabled Then txtOpen(Index).SetFocus
    End If
    If updOpen(Index).Min > Val(txtOpen(Index).Text) Then
        MsgBox "开放数量不能小于最小可开放数量(" & updOpen(Index).Min & ")！", vbExclamation, gstrSysName
        txtOpen(Index).Text = updOpen(Index).Min
        If txtOpen(Index).Visible And txtOpen(Index).Enabled Then txtOpen(Index).SetFocus
    End If
    mblnNotChange = False
    Call OpenSN(Index, Val(txtOpen(Index).Text))
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub OpenSN(Index As Integer, ByVal lngCount As Long)
    '开放序号
    Dim lngRow As Long, lngCol As Long
    Dim strStopStart As String, strStopEnd As String
    Dim strStart As String, blnStart As Boolean
    
    On Error GoTo ErrHandler
    strStopStart = Split(txt停诊时间(Index).Tag & "～", "～")(0)
    strStopEnd = Split(txt停诊时间(Index).Tag & "～", "～")(1)
    If Not (IsDate(strStopStart) And IsDate(strStopEnd)) Then Exit Sub
    With vsfTimeWork(Index)
        .Redraw = flexRDNone
        
        If Val(vsfTimeWork(Index).Tag) > lngCount Then
            OpenSN Index, vsfTimeWork(Index).Tag '先按增加恢复初始状态，保证从停诊开始序号到开放的最大序号都是连续的
            lngCount = Val(vsfTimeWork(Index).Tag) - lngCount
            
            '减少开放数量
            For lngRow = .Rows - 1 To .FixedRows Step -1
                For lngCol = .Cols - 1 To .FixedCols Step -1
                    If .TextMatrix(lngRow, lngCol) <> "" Then
                        strStart = Split(.Cell(flexcpData, lngRow, lngCol), "～")(0)
                        If DateDiff("n", strStopStart, strStart) >= 0 And DateDiff("n", strStopEnd, strStart) <= 0 Then
                            If .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor And blnStart = False Then
                                blnStart = True '标记开始
                            End If
                            If blnStart And lngCount > 0 Then
                                If (.Cell(flexcpForeColor, lngRow, lngCol)) = vbBlack Then '黑色字体的挂号状态为"0-待挂或待预约的号"
                                    .Cell(flexcpBackColor, lngRow, lngCol) = vbRed
                                    lngCount = lngCount - 1
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        Else
            For lngRow = .FixedRows To .Rows - 1
                For lngCol = .FixedCols To .Cols - 1
                    If .TextMatrix(lngRow, lngCol) <> "" Then
                        strStart = Split(.Cell(flexcpData, lngRow, lngCol), "～")(0)
                        If DateDiff("n", strStopStart, strStart) >= 0 And DateDiff("n", strStopEnd, strStart) <= 0 Then
                            If lngCount > 0 Then
                                .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor
                                lngCount = lngCount - 1
                            Else
                                .Cell(flexcpBackColor, lngRow, lngCol) = vbRed
                            End If
                        End If
                    End If
                Next
            Next
        End If
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub txtOpen_GotFocus(Index As Integer)
    gobjControl.TxtSelAll txtOpen(Index)
End Sub

Private Sub txtOpen_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call gobjCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub LoadControl(ByVal Index As Long)
    '增加控件
    On Error GoTo ErrHandler
    If ExistsControl(picTimeWork(Index)) Then
        txt限号数(Index).Text = ""
        txt限约数(Index).Text = ""
        txt停诊时间(Index).Text = "": txt停诊时间(Index).Tag = ""
        vsfTimeWork(Index).Clear: vsfTimeWork(Index).Rows = 0
        mblnNotChange = True
        txtOpen(Index).Text = "0"
        mblnNotChange = False
    Else
        Load picTimeWork(Index): picTimeWork(Index).Visible = True
        
        Load lbl限号数(Index): lbl限号数(Index).Visible = True: Set lbl限号数(Index).Container = picTimeWork(Index)
        Load txt限号数(Index): txt限号数(Index).Visible = True: Set txt限号数(Index).Container = picTimeWork(Index)
        Load lbl限约数(Index): lbl限约数(Index).Visible = True: Set lbl限约数(Index).Container = picTimeWork(Index)
        Load txt限约数(Index): txt限约数(Index).Visible = True: Set txt限约数(Index).Container = picTimeWork(Index)
        Load lbl停诊时间(Index): lbl停诊时间(Index).Visible = True: Set lbl停诊时间(Index).Container = picTimeWork(Index)
        Load txt停诊时间(Index): txt停诊时间(Index).Visible = True: Set txt停诊时间(Index).Container = picTimeWork(Index)
        
        Load vsfTimeWork(Index): vsfTimeWork(Index).Visible = True: Set vsfTimeWork(Index).Container = picTimeWork(Index)
        
        Load lblOpen(Index): lblOpen(Index).Visible = True: Set lblOpen(Index).Container = picTimeWork(Index)
        Load txtOpen(Index): txtOpen(Index).Visible = True: Set txtOpen(Index).Container = picTimeWork(Index)
        Load updOpen(Index): updOpen(Index).Visible = True: Set updOpen(Index).Container = picTimeWork(Index)
        updOpen(Index).BuddyControl = txtOpen(Index): updOpen(Index).BuddyProperty = "Text"
        Load lblToolTip(Index): lblToolTip(Index).Visible = True: Set lblToolTip(Index).Container = picTimeWork(Index)
    End If
    txtOpen(Index).Enabled = True: updOpen(Index).Enabled = True
    
    '特殊处理一下，因为在动态加载updOpen时前一个txtOpen的宽度会变
    Dim i As Integer
    For i = txtOpen.LBound To txtOpen.UBound
        txtOpen(i).Width = 1100: updOpen(i).Left = txtOpen(i).Left + txtOpen(i).Width + 10
    Next
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub UnLoadControl(ByVal Index As Long)
    '卸载控件
    On Error GoTo ErrHandler
    If ExistsControl(picTimeWork(Index)) Then
        If Index = 0 Then
            txt限号数(Index).Text = ""
            txt限约数(Index).Text = ""
            txt停诊时间(Index).Text = "": txt停诊时间(Index).Tag = ""
            vsfTimeWork(Index).Clear: vsfTimeWork(Index).Rows = 0
            mblnNotChange = True
            txtOpen(Index).Text = "0"
            mblnNotChange = False
            txtOpen(Index).Enabled = False: updOpen(Index).Enabled = False
            lblToolTip(Index).Caption = ""
            tbPage(Index).Tag = ""
        Else
            '不能卸载，在ComboBox_Click中不能卸载控件，报错"不能在该上下文中卸载（错误 365）"
'            Unload lbl限号数(index): Unload txt限号数(index)
'            Unload lbl限约数(index): Unload txt限约数(index)
'            Unload lbl停诊时间(index): Unload txt停诊时间(index)
'
'            Unload vsfTimeWork(index)
'
'            Unload lblOpen(index): Unload txtOpen(index): Unload updOpen(index): Unload lblToolTip(index)
'
'            Unload picTimeWork(index)
        End If
    End If
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Function ExistsControl(ByRef ctlVal As Control) As Boolean
    '判断控件是否实例
    Dim strTmp As String

    On Error GoTo ErrHandler
    strTmp = ctlVal.Name
    ExistsControl = True
    Exit Function
ErrHandler:
    ExistsControl = False
End Function

Private Sub vsfSignalSource_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    mlngPreRow = NewRow
End Sub

Private Sub vsfSignalSource_EnterCell()
    Dim lng号源ID As Long
    Dim lngRow As Long, i As Integer
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngCanOpenMax As Long, lngCanOpenMin As Long, lngPreOpen As Long

    Err = 0: On Error GoTo ErrHandler
    If mblnNotChange Then Exit Sub
    If vsfSignalSource.Row < vsfSignalSource.FixedRows Then Exit Sub
    If mrsRecord Is Nothing Then Exit Sub
    
    If mblnChanged Then
        If MsgBox("当前安排的开放数量已改变，但您还未保存，是否不保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            mblnNotChange = True
            vsfSignalSource.Row = mlngPreRow
            mblnNotChange = False
            Exit Sub
        Else
            mblnChanged = False
        End If
    End If
    
    For i = picTimeWork.UBound To 1 Step -1
        tbPage.RemoveItem i
        UnLoadControl i
    Next
    tbPage(0).Caption = "无上班时段"
    UnLoadControl 0
    
    lng号源ID = Val(vsfSignalSource.TextMatrix(vsfSignalSource.Row, vsfSignalSource.ColIndex("号源ID")))
    mrsRecord.Filter = "号源ID=" & lng号源ID
    lngRow = 0
    Do While Not mrsRecord.EOF
        LoadControl lngRow
        
        txt限号数(lngRow).Text = IIf(Val(Nvl(mrsRecord!限号数)) = 0, "", Nvl(mrsRecord!限号数))
        txt限约数(lngRow).Text = IIf(Val(Nvl(mrsRecord!预约控制)) = 1, "禁止预约", IIf(Val(Nvl(mrsRecord!限约数)) = 0, txt限号数(lngRow).Text, Nvl(mrsRecord!限约数)))
        txt停诊时间(lngRow).Text = Format(Nvl(mrsRecord!停诊开始时间), "hh:mm") & "～" & Format(Nvl(mrsRecord!停诊终止时间), "hh:mm")
        txt停诊时间(lngRow).Tag = Format(Nvl(mrsRecord!停诊开始时间), "yyyy-mm-dd hh:mm") & "～" & Format(Nvl(mrsRecord!停诊终止时间), "yyyy-mm-dd hh:mm")
        If txt停诊时间(lngRow).Text = "～" Then txt停诊时间(0).Text = ""
        
        If Not mrsRecordCount Is Nothing Then
            mrsRecordCount.Filter = "记录ID=" & Val(Nvl(mrsRecord!记录ID))
            If Not mrsRecordCount.EOF Then
                vsfTimeWork(lngRow).Tag = Val(Nvl(mrsRecordCount!停诊范围))
                lngCanOpenMax = Val(Nvl(mrsRecordCount!最大数量))
                lngCanOpenMin = Val(Nvl(mrsRecordCount!最小数量))
                lngPreOpen = Val(Nvl(mrsRecordCount!上次开放数量))
            End If
        End If
        
        mblnNotChange = True
        updOpen(lngRow).Max = lngCanOpenMax
        updOpen(lngRow).Min = lngCanOpenMin
        txtOpen(lngRow).Text = lngPreOpen
        mblnNotChange = False
        
        LoadDataToGrid lngRow, Val(Nvl(mrsRecord!记录ID)), lngCanOpenMax, lngCanOpenMin
        
        If lngRow = 0 Then
            tbPage(0).Caption = Nvl(mrsRecord!上班时段)
        Else
            tbPage.InsertItem lngRow, Nvl(mrsRecord!上班时段), picTimeWork(lngRow).hWnd, 0
        End If
        tbPage(lngRow).Tag = Nvl(mrsRecord!记录ID)
        
        lngRow = lngRow + 1
        mrsRecord.MoveNext
    Loop
    If txtOpen(0).Visible And txtOpen(0).Enabled Then txtOpen(0).SetFocus
    mblnChanged = False
    Exit Sub
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub vsfTimeWork_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    If vsfTimeWork(Index).TextMatrix(NewRow, NewCol) = "" Then Cancel = True: Exit Sub
End Sub

Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = True, Optional str所有部门 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参:cboDept-指定的部门部件
    '     rsDept-指定的部门
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有部门-所有部门名称
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strIDs As String, str简码 As String
    
    On Error GoTo ErrHandler
    '先复制记录集
    Set rsTemp = gobjDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(gSysPara.strLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf gobjCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有部门 <> "" Then
        str简码 = gobjCommFun.SpellCode(str所有部门)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有部门) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编码) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编码) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!名称)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编码)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!名称))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    End Select
    
    '弹出选择器
    If gobjDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "" & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then gobjCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.TxtSelAll cboDept
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = True, Optional str所有 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:人员选择选择器
    '入参:cboSel-指定的部门选择部件
    '     rsPerson-指定的人员信息(ID,编号,姓名,简码)
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有-所有名称(所有人,所有操作员等)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strIDs As String, str简码 As String, strLike As String
    
    On Error GoTo ErrHandler
    '先复制记录集
    Set rsTemp = gobjDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gSysPara.strLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf gobjCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有 <> "" Then
        str简码 = gobjCommFun.SpellCode(str所有)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编号 = "-"
                rsTemp!姓名 = str所有
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编号 = "-"
                rsTemp!姓名 = str所有
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编号) = strSearch Then lngID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编号)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编号) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编号) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!姓名)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编号)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!姓名))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call gobjDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编号"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编号"
    End Select
    
    '弹出选择器
    If gobjDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "部门ID" & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!ID))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.CboLocate cboSel, lngID, True
    If blnSendKeys Then gobjCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    gobjControl.TxtSelAll cboSel
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAllDoctor() As ADODB.Recordset
    '获取医生列表
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = "Select c.id,c.编号,c.姓名,c.简码,b.部门id" & vbNewLine & _
            " From 人员性质说明 A, 部门人员 B, 人员表 C" & vbNewLine & _
            " Where b.人员id=c.id And b.人员id=a.人员id And a.人员性质=[1]" & vbNewLine & _
            "       And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
            "       And (c.站点=[2] Or c.站点 is Null)" & vbNewLine & _
            " Order by c.编号"
    Set GetAllDoctor = gobjDatabase.OpenSQLRecord(strSQL, "获取医生", "医生", gstrNodeNo)
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetDepartments(ByVal str性质 As String, _
    ByVal str服务对象 As String, _
    Optional ByVal lng人员id As Long = 0, _
    Optional ByVal blnCheck站点 As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定性质的部门列表
    '入参:str性质='临床','护理','中药房',...,允许为空
    '     str服务对象:以,分离:如1,3
    '     lng人员ID-不等于0，则人员的所属部门
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    str性质 = Replace(str性质, "'", "")
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.工作性质||',')>0"
        Else
            strSQL = " And B.工作性质 = [1]"
        End If
    End If
    If lng人员id <> 0 Then strSQL = strSQL & "  And A.id=C.部门ID and C.人员id =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码 " & _
        " From 部门表 A,部门性质说明 B " & IIf(lng人员id <> 0, ",部门人员 C", "") & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And Instr(',' || [2]|| ',',',' || B.服务对象 || ',')>0 " & strSQL & _
         IIf(blnCheck站点, " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
        " Order by A.编码"
    Set GetDepartments = gobjDatabase.OpenSQLRecord(strSQL, "获取科室", str性质, str服务对象, lng人员id)
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

