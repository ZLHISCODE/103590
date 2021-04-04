VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanNumberLimitModify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "加号"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10335
   Icon            =   "frmClinicPlanNumberLimitModify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   9030
      TabIndex        =   39
      Top             =   6810
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   9030
      TabIndex        =   37
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9030
      TabIndex        =   38
      Top             =   810
      Width           =   1100
   End
   Begin VB.Frame fra号源信息 
      Caption         =   "号源基本信息"
      Height          =   1035
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   8835
      Begin VB.TextBox txtItem 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   10
         Top             =   645
         Width           =   3105
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4950
         TabIndex        =   6
         Top             =   285
         Width           =   1605
      End
      Begin VB.TextBox txtDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7290
         TabIndex        =   8
         Top             =   285
         Width           =   1365
      End
      Begin VB.TextBox txt假日控制 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4950
         TabIndex        =   12
         Top             =   645
         Width           =   1605
      End
      Begin VB.CheckBox chk建档 
         Caption         =   "挂号时必须建档"
         Enabled         =   0   'False
         Height          =   180
         Left            =   6900
         TabIndex        =   13
         Top             =   705
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtSignalNO 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   2
         Top             =   285
         Width           =   1035
      End
      Begin VB.TextBox txt号类 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2670
         TabIndex        =   4
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label lbl假日控制 
         AutoSize        =   -1  'True
         Caption         =   "假日控制"
         Height          =   180
         Left            =   4200
         TabIndex        =   11
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         Height          =   180
         Left            =   6900
         TabIndex        =   7
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   180
         Left            =   4560
         TabIndex        =   5
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "项目"
         Height          =   180
         Left            =   450
         TabIndex        =   9
         Top             =   705
         Width           =   360
      End
      Begin VB.Label lbl号类 
         AutoSize        =   -1  'True
         Caption         =   "号类"
         Height          =   180
         Left            =   2280
         TabIndex        =   3
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblSignalNO 
         AutoSize        =   -1  'True
         Caption         =   "号码"
         Height          =   180
         Left            =   450
         TabIndex        =   1
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame fra出诊信息 
      Caption         =   "出诊信息"
      Height          =   1035
      Left            =   30
      TabIndex        =   14
      Top             =   1140
      Width           =   8835
      Begin VB.TextBox txt上班时段 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2880
         TabIndex        =   18
         Top             =   285
         Width           =   1065
      End
      Begin VB.TextBox txt预约控制 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4950
         TabIndex        =   20
         Top             =   285
         Width           =   1605
      End
      Begin VB.CheckBox chk时段 
         Caption         =   "启用时段"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4230
         TabIndex        =   24
         Top             =   698
         Width           =   1035
      End
      Begin VB.CheckBox chk序号控制 
         Caption         =   "启用序号控制"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2130
         TabIndex        =   23
         Top             =   698
         Width           =   1395
      End
      Begin VB.TextBox txt替诊医生 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   22
         Top             =   660
         Width           =   1065
      End
      Begin VB.TextBox txt出诊日期 
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   16
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label lbl预约控制 
         AutoSize        =   -1  'True
         Caption         =   "预约控制"
         Height          =   180
         Left            =   4200
         TabIndex        =   19
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lbl上班时段 
         AutoSize        =   -1  'True
         Caption         =   "上班时段"
         Height          =   180
         Left            =   2130
         TabIndex        =   17
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lbl替诊医生 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "替诊医生"
         Height          =   180
         Left            =   90
         TabIndex        =   21
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl出诊日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出诊日期"
         Height          =   180
         Left            =   90
         TabIndex        =   15
         Top             =   345
         Width           =   720
      End
   End
   Begin zl9RegEvent.ClinicPlanWorkTimeNum cpWorkTimeNum 
      Height          =   4485
      Left            =   30
      TabIndex        =   36
      Top             =   2940
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   7911
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IsDataChanged   =   -1  'True
   End
   Begin VB.Frame fraLimitInfo 
      Caption         =   "限号信息"
      Height          =   675
      Left            =   30
      TabIndex        =   25
      Top             =   2220
      Width           =   8835
      Begin VB.TextBox txtAdd限约数 
         Height          =   300
         Left            =   7170
         MaxLength       =   6
         TabIndex        =   34
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox txtAdd限号数 
         Height          =   300
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   29
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox txt限约数 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5160
         TabIndex        =   32
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox txt限号数 
         Enabled         =   0   'False
         Height          =   300
         Left            =   810
         TabIndex        =   27
         Top             =   285
         Width           =   1095
      End
      Begin MSComCtl2.UpDown upd限号 
         Height          =   285
         Left            =   3960
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   293
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtAdd限号数"
         BuddyDispid     =   196640
         OrigLeft        =   1200
         OrigTop         =   120
         OrigRight       =   1455
         OrigBottom      =   420
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown upd限约 
         Height          =   285
         Left            =   8250
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   293
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtAdd限约数"
         BuddyDispid     =   196639
         OrigLeft        =   1200
         OrigTop         =   120
         OrigRight       =   1455
         OrigBottom      =   420
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblAdd限约数 
         AutoSize        =   -1  'True
         Caption         =   "本次增加"
         Height          =   180
         Left            =   6420
         TabIndex        =   33
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lblAdd限号数 
         AutoSize        =   -1  'True
         Caption         =   "本次增加"
         Height          =   180
         Left            =   2130
         TabIndex        =   28
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lbl限约数 
         AutoSize        =   -1  'True
         Caption         =   "限约数"
         Height          =   180
         Left            =   4590
         TabIndex        =   31
         Top             =   345
         Width           =   540
      End
      Begin VB.Label lbl限号数 
         AutoSize        =   -1  'True
         Caption         =   "限号数"
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   345
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmClinicPlanNumberLimitModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As Byte '1-加号，2-减号
Private mobj号源 As 出诊号源, mobj出诊记录 As 出诊记录

Private mblnOk As Boolean
Private mlngMinSN As Long '可以删除的时段的最小序号（分时段，启用序号）
Private mblnNotChanged As Boolean

Public Function ShowMe(frmParent As Form, ByVal bytFun As Byte, _
    ByVal obj号源 As 出诊号源, ByVal obj出诊记录 As 出诊记录) As Boolean
    
    If obj号源 Is Nothing Then Exit Function
    If obj出诊记录 Is Nothing Then Exit Function
    
    mbytFun = bytFun
    Set mobj号源 = obj号源: Set mobj出诊记录 = obj出诊记录
    
    If CheckDepend = False Then Exit Function
    mblnOk = False
    On Error Resume Next
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Function CheckDepend() As Boolean
    '功能:检查数据
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    '不能对历史的安排进行操作
    If DateDiff("s", mobj出诊记录.终止时间, zlDatabase.Currentdate) >= 0 Then
        MsgBox "当前系统时间已大于了安排时段的终止时间，不能进行" & IIf(mbytFun = 1, "加号", "减号") & "操作！", vbInformation, gstrSysName
        Exit Function
    End If
    '无限号数的不能调整
    If mobj出诊记录.限号数 = 0 Then
        MsgBox "当前安排时段为不限号，不能进行" & IIf(mbytFun = 1, "加号", "减号") & "操作！", vbInformation, gstrSysName
        Exit Function
    End If
    '已经停诊或未出诊安排的，不允许加号/减号
    strSQL = "Select 1 from 临床出诊记录 Where ID=[1] and 上班时段=[2] And 停诊开始时间 Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查出诊记录", mobj出诊记录.记录ID, mobj出诊记录.时间段)
    If rsTemp.EOF Then
        MsgBox "当前安排时段不存在或已停诊，不能进行" & IIf(mbytFun = 1, "加号", "减号") & "操作！", vbInformation, gstrSysName
        Exit Function
    End If
    '限号数已经被全部使用的，则不允许减号
    If mobj出诊记录.已挂数 >= mobj出诊记录.限号数 And mbytFun = 2 Then
        MsgBox "当前安排时段已全部挂号，不能进行减号操作！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckDepend = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    If cpWorkTimeNum.限号数 < 1 Then
        MsgBox IIf(mbytFun = 1, "", "剩余") & "限号数(" & cpWorkTimeNum.限号数 & ")小于了1！", vbInformation, gstrSysName
        Exit Sub
    End If
    If cpWorkTimeNum.限号数 < mlngMinSN Then
        MsgBox IIf(mbytFun = 1, "", "剩余") & "限号数(" & cpWorkTimeNum.限号数 & ")小于了已使用时段的最大序号(" & mlngMinSN & ")！", vbInformation, gstrSysName
        Exit Sub
    End If
    If cpWorkTimeNum.限约数 > cpWorkTimeNum.限号数 Then
        MsgBox IIf(mbytFun = 1, "", "剩余") & "限约数(" & cpWorkTimeNum.限约数 & ")大于了" & IIf(mbytFun = 1, "", "剩余") & "限号数(" & cpWorkTimeNum.限号数 & ")！", vbInformation, gstrSysName
        Exit Sub
    End If
    If cpWorkTimeNum.限号数 < mobj出诊记录.已挂数 Then
        MsgBox IIf(mbytFun = 1, "", "剩余") & "限号数(" & cpWorkTimeNum.限号数 & ")小于了已挂数(" & mobj出诊记录.已挂数 & ")！", vbInformation, gstrSysName
        Exit Sub
    End If
    '预约控制:0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
    If mobj出诊记录.预约控制 <> 1 And cpWorkTimeNum.限约数 < mobj出诊记录.已约数 Then
        MsgBox IIf(mbytFun = 1, "", "剩余") & "限约数(" & cpWorkTimeNum.限约数 & ")小于了已约数(" & mobj出诊记录.已约数 & ")！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtAdd限号数.Text) = 0 And Val(txtAdd限约数.Text) = 0 And cpWorkTimeNum.IsDataChanged = False Then
        If MsgBox("本次未进行任何调整，不需要保存！要退出调整吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Unload Me
        End If
        Exit Sub
    End If
    If mbytFun = 2 And mobj出诊记录.限约数 > 0 And cpWorkTimeNum.限约数 = 0 Then
        If MsgBox("剩余限约数为0表示禁止预约，你确定要对该出诊安排进行禁止预约吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    If cpWorkTimeNum.IsValied() = False Then Exit Sub
    
    If SaveData() = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    Call InitData
    Call SetEnabledBackColor(Me.Controls)
    
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function InitData() As Boolean
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler
    Select Case mbytFun
    Case 1 '加号
        Me.Caption = "加号"
        lblAdd限约数.Caption = "本次增加"
        lblAdd限号数.Caption = "本次增加"
    Case 2 '减号
        Me.Caption = "减号"
        lblAdd限约数.Caption = "本次减少"
        lblAdd限号数.Caption = "本次减少"
    End Select
    
    '号源信息
    txtSignalNO.Text = mobj号源.号码
    txt号类.Text = mobj号源.号类
    txtDept.Text = mobj号源.科室名称
    txtItem.Text = mobj号源.项目名称
    txtDoctor.Text = mobj号源.医生姓名
    txt假日控制.Text = Decode(mobj号源.假日控制状态, 1, "开放预约", 2, "禁止预约", 3, "受节假日设置控制", "不上班")
    chk建档.Value = IIf(mobj号源.是否建病案, vbChecked, vbUnchecked)
    
    If IsDate(mobj出诊记录.出诊日期) Then
        txt出诊日期.Text = Format(mobj出诊记录.出诊日期, "yyyy-mm-dd")
    Else
        txt出诊日期.Text = mobj出诊记录.出诊日期
    End If
    txt上班时段.Text = mobj出诊记录.时间段
    txt替诊医生.Text = mobj出诊记录.替诊医生
    txt预约控制.Text = Choose(mobj出诊记录.预约控制 + 1, "允许预约", "禁止预约", "仅禁止三方机构预约")
    chk序号控制.Value = IIf(mobj出诊记录.是否序号控制, vbChecked, vbUnchecked)
    chk时段.Value = IIf(mobj出诊记录.是否分时段, vbChecked, vbUnchecked)
    txt限号数.Text = IIf(mobj出诊记录.限号数 <> 0, mobj出诊记录.限号数, "")
    txt限约数.Text = IIf(mobj出诊记录.限约数 <> 0, mobj出诊记录.限约数, "")
    
    '禁止预约和不限制预约的不允许修改预约数
    txtAdd限约数.Enabled = Not (mobj出诊记录.预约控制 = 1 Or mobj出诊记录.限约数 = 0)
    upd限约.Enabled = txtAdd限约数.Enabled
    
    cpWorkTimeNum.CanReCalic = mobj出诊记录.已约数 = 0 And mobj出诊记录.已挂数 = 0 And mobj出诊记录.是否分时段
    cpWorkTimeNum.EditMode = ED_RegistPlan_NumLimitModify

    If cpWorkTimeNum.CanReCalic = False Or mobj出诊记录.是否分时段 Then
        '标记哪些号序不能修改
        Dim strSQL As String, rsTemp As ADODB.Recordset
        Dim cllFixedSN As New Collection
        strSQL = "Select a.序号, Nvl(Sum(a.数量), 0) As 已约数" & vbNewLine & _
                " From 临床出诊序号控制 A" & vbNewLine & _
                " Where a.记录id = [1] And Nvl(a.挂号状态, 0) <> 0" & vbNewLine & _
                " Group By a.序号" & vbNewLine & _
                " Order By a.序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobj出诊记录.记录ID)
        Do While Not rsTemp.EOF
            cllFixedSN.Add Array(Val(Nvl(rsTemp!序号)), Val(Nvl(rsTemp!已约数)))
            If mobj出诊记录.是否序号控制 Then
                If Val(Nvl(rsTemp!序号)) > mlngMinSN Then mlngMinSN = Val(Nvl(rsTemp!序号))
            End If
            rsTemp.MoveNext
        Loop
    End If
    InitData = cpWorkTimeNum.LoadData(mobj出诊记录.号序信息集.Clone, mobj出诊记录.上班时段, cllFixedSN)
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtAdd限号数_Change()
    If mblnNotChanged = False Then
        If mobj出诊记录.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text) < 1 Then
            MsgBox IIf(mbytFun = 1, "", "剩余") & "限号数(" & mobj出诊记录.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text) & ")不能小于1！", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd限号数.Text = Val(txtAdd限号数.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
        If mobj出诊记录.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text) < mlngMinSN Then
            MsgBox IIf(mbytFun = 1, "", "剩余") & "限号数(" & mobj出诊记录.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text) & ")不能小于已使用时段的最大序号(" & mlngMinSN & ")！", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd限号数.Text = Val(txtAdd限号数.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
        If mobj出诊记录.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text) < mobj出诊记录.已挂数 Then
            MsgBox IIf(mbytFun = 1, "", "剩余") & "限号数(" & mobj出诊记录.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text) & ")不能小于已挂数(" & mobj出诊记录.已挂数 & ")！", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd限号数.Text = Val(txtAdd限号数.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
    End If
    mblnNotChanged = True
    txtAdd限号数.Text = IIf(Val(txtAdd限号数.Text) = 0, "", txtAdd限号数.Text)
    mblnNotChanged = False
    
    If Val(txtAdd限号数.Tag) <> Val(txtAdd限号数.Text) Then
        txtAdd限号数.Tag = Val(txtAdd限号数.Text)
        cpWorkTimeNum.SetNewSN mobj出诊记录.号序信息集.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text), _
            IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text), mbytFun = 1
    End If
    
    If mobj出诊记录.号序信息集.限约数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限约数.Text) > mobj出诊记录.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text) Then
        mblnNotChanged = True
        txtAdd限约数.Text = Abs(mobj出诊记录.限号数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限号数.Text) - mobj出诊记录.号序信息集.限约数)
        txtAdd限约数.Text = IIf(Val(txtAdd限约数.Text) = 0, "", txtAdd限约数.Text)
        mblnNotChanged = False
    End If
End Sub

Private Sub txtAdd限号数_GotFocus()
    zlControl.TxtSelAll txtAdd限号数
End Sub

Private Sub txtAdd限号数_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtAdd限约数_Change()
    If mblnNotChanged = False Then
        If mobj出诊记录.号序信息集.限约数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限约数.Text) > cpWorkTimeNum.限号数 Then
            MsgBox IIf(mbytFun = 1, "", "剩余") & "限约数(" & mobj出诊记录.限约数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限约数.Text) & ")不能大于限号数(" & cpWorkTimeNum.限号数 & ")！", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd限约数.Text = Val(txtAdd限约数.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
        If mobj出诊记录.限约数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限约数.Text) < mobj出诊记录.已约数 Then
            MsgBox IIf(mbytFun = 1, "", "剩余") & "限约数(" & mobj出诊记录.限约数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限约数.Text) & ")不能小于已约数(" & mobj出诊记录.已约数 & ")！", vbInformation, gstrSysName
            mblnNotChanged = True
            txtAdd限约数.Text = Val(txtAdd限约数.Tag)
            mblnNotChanged = False
            Exit Sub
        End If
    End If
    mblnNotChanged = True
    txtAdd限约数.Text = IIf(Val(txtAdd限约数.Text) = 0, "", txtAdd限约数.Text)
    mblnNotChanged = False
    
    If Val(txtAdd限约数.Tag) <> Val(txtAdd限约数.Text) Then
        txtAdd限约数.Tag = Val(txtAdd限约数.Text)
        cpWorkTimeNum.限约数 = mobj出诊记录.号序信息集.限约数 + IIf(mbytFun = 1, 1, -1) * Val(txtAdd限约数.Text)
    End If
End Sub

Private Sub txtAdd限约数_GotFocus()
    zlControl.TxtSelAll txtAdd限约数
End Sub

Private Sub txtAdd限约数_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, cllPro As New Collection, i As Integer
    Dim obj号序 As 号序信息
    Dim cll号序 As Collection, str号序 As String, strTemp As String
    Dim blnTrans As Boolean
    
    Err = 0: On Error GoTo errHandler
    Set mobj出诊记录.号序信息集 = cpWorkTimeNum.Get号序集
    '插入变动记录
    'Zl_临床出诊序号控制变动(
    strSQL = "Zl_临床出诊序号控制变动("
    '记录id_In     临床出诊变动记录.记录id%Type,
    strSQL = strSQL & "" & mobj出诊记录.记录ID & ","
    '限号数_In     临床出诊记录.限号数%Type,
    strSQL = strSQL & "" & mobj出诊记录.号序信息集.限号数 & ","
    '限约数_In     临床出诊记录.限约数%Type,
    strSQL = strSQL & "" & mobj出诊记录.号序信息集.限约数 & ","
    '原已挂数_In   临床出诊记录.已挂数%Type,
    strSQL = strSQL & "" & mobj出诊记录.已挂数 & ","
    '原已约数_In   临床出诊记录.已约数%Type,
    strSQL = strSQL & "" & mobj出诊记录.已约数 & ","
    '操作员姓名_In 临床出诊变动记录.操作员姓名%Type := Null,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '登记时间_In   临床出诊变动记录.登记时间%Type := Null
    strSQL = strSQL & "" & ZDate(zlDatabase.Currentdate) & ")"
    cllPro.Add strSQL
    
    Set cll号序 = New Collection
    For Each obj号序 In mobj出诊记录.号序信息集
        strTemp = obj号序.序号 & "," & _
            GetWorkTrueDate(mobj出诊记录.开始时间, ZDate(obj号序.开始时间, mobj出诊记录.开始时间, False), , False) & "," & _
            GetWorkTrueDate(mobj出诊记录.开始时间, ZDate(obj号序.终止时间, mobj出诊记录.终止时间, False)) & "," & _
            obj号序.数量 & "," & IIf(obj号序.是否预约, 1, 0)
        If zlCommFun.ActualLen(str号序 & "|" & strTemp) > 2000 Then
            '时段_In:序号,开始时间,终止时间,限制数量,预约标志|...
            str号序 = Mid(str号序, 2)
            cll号序.Add str号序
            str号序 = ""
        End If
        str号序 = str号序 & "|" & strTemp
    Next
    If str号序 <> "" Then
        str号序 = Mid(str号序, 2)
        cll号序.Add str号序
    End If
    For i = 1 To IIf(cll号序.Count = 0, 1, cll号序.Count)
        'Zl_临床出诊序号控制_Update(
        strSQL = "Zl_临床出诊序号控制_Update("
        '记录id_In   临床出诊记录.Id%Type,
        strSQL = strSQL & "" & mobj出诊记录.记录ID & ","
        '时段_In     Varchar2 := Null,--序号,开始时间,终止时间,限制数量,预约标志|...
        str号序 = ""
        If cll号序.Count > 0 Then str号序 = cll号序(i)
        strSQL = strSQL & "'" & str号序 & "',"
        '删除序号_In Number:=0 --是否删除现有序号时段
        strSQL = strSQL & "" & IIf(i = 1, 1, 0) & ")"
        cllPro.Add strSQL
    Next
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    blnTrans = False
    SaveData = True
    Exit Function
errHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

