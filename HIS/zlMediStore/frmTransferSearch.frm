VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTransferSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   4260
   ClientLeft      =   3156
   ClientTop       =   3168
   ClientWidth     =   7560
   Icon            =   "frmTransferSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   1605
      Left            =   1200
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7853
      _ExtentY        =   2836
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   3975
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   6015
      _ExtentX        =   10605
      _ExtentY        =   7006
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmTransferSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmTransferSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   2850
         Left            =   -74760
         TabIndex        =   25
         Top             =   600
         Width           =   5520
         Begin MSComctlLib.ListView lvw剂型 
            Height          =   1755
            Left            =   1560
            TabIndex        =   35
            Top             =   2640
            Visible         =   0   'False
            Width           =   3885
            _ExtentX        =   6858
            _ExtentY        =   3090
            View            =   1
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            Icons           =   "imgsDrug"
            SmallIcons      =   "imgsDrug"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "名称"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.TreeView tvw类别 
            Height          =   2205
            Left            =   0
            TabIndex        =   36
            Top             =   2640
            Visible         =   0   'False
            Width           =   3645
            _ExtentX        =   6435
            _ExtentY        =   3895
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imgsDrug"
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.CheckBox chkClass 
            Caption         =   "药品分类"
            Height          =   300
            Left            =   360
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkJiXin 
            Caption         =   "药品剂型"
            Height          =   300
            Left            =   360
            TabIndex        =   33
            Top             =   680
            Width           =   1095
         End
         Begin VB.CheckBox Chk药品 
            Caption         =   "药品"
            Height          =   300
            Left            =   360
            TabIndex        =   32
            Top             =   1120
            Width           =   990
         End
         Begin VB.TextBox Txt药品 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            ScrollBars      =   3  'Both
            TabIndex        =   31
            Top             =   1120
            Width           =   3255
         End
         Begin VB.CommandButton cmdClass 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4800
            TabIndex        =   30
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdJiXin 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4800
            TabIndex        =   29
            Top             =   680
            Width           =   255
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   28
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox txtJiXing 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   27
            Top             =   680
            Width           =   3255
         End
         Begin VB.CheckBox Chk移入库房 
            Caption         =   "移入库房"
            Height          =   420
            Left            =   360
            TabIndex        =   9
            Top             =   1500
            Width           =   1110
         End
         Begin VB.CommandButton Cmd药品 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4800
            TabIndex        =   8
            Top             =   1120
            Width           =   255
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   11
            Top             =   2100
            Width           =   1845
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   12
            Top             =   2460
            Width           =   1845
         End
         Begin VB.ComboBox Cbo库房 
            Enabled         =   0   'False
            Height          =   276
            Left            =   1530
            TabIndex        =   10
            Text            =   "Cbo库房"
            Top             =   1560
            Width           =   3550
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   255
            Left            =   570
            TabIndex        =   18
            Top             =   2123
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   570
            TabIndex        =   19
            Top             =   2520
            Width           =   540
         End
      End
      Begin VB.Frame fra范围 
         Height          =   2850
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkStrike 
            Caption         =   "包含冲销"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   39
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CheckBox chkYesStrike 
            Caption         =   "已审核冲销"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   38
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkNoStrike 
            Caption         =   "未审核冲销"
            Height          =   300
            Left            =   720
            TabIndex        =   37
            Top             =   1400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txt开始No 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   0
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt结束NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chk填制 
            Caption         =   "未审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   2
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "已审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   5
            Top             =   1680
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   4
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   312
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   1968
            Width           =   1608
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   312
            Index           =   1
            Left            =   3588
            TabIndex        =   7
            Top             =   1968
            Width           =   1608
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   15
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   2640
            TabIndex        =   24
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核日期"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   17
            Top             =   2028
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   3
            Left            =   3348
            TabIndex        =   23
            Top             =   2028
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制日期"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   16
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   22
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   14
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   13
      Top             =   435
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgsDrug 
      Left            =   0
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferSearch.frx":0044
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferSearch.frx":12C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferSearch.frx":1860
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmTransferSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '查找字符串
Private BlnAdvance As Boolean '是否展开
Private mlngMode As Long    '单据类型
Private mdatStart As Date   '开始时间
Private mdatEnd As Date     '结束时间
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mfrmMain As Form    '父窗体
Private mstrSelectTag As String     '当前选择的对象
Private mblnStock As Boolean            '当前操作员是否是药库人员，仅对领用单据有效
Private mint入出类型 As Integer
Private mlngStoreId As Long     '当前库房id
Private mstrMatch As String '匹配方式 0-双向匹配 1-从左向右单向匹配
Private mint冲销申请 As Integer '0-不需要申请;1-需要申请

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    lng库房 As Long
    str填制人 As String
    str审核人 As String
    int填制审核一并查询 As Integer
    lng药品分类 As Long
    str剂型 As String
End Type

Private SQLCondition As Type_SQLCondition

Public Property Get In_入出类型() As Integer
    In_入出类型 = mint入出类型
End Property

Public Property Let In_入出类型(ByVal vNewValue As Integer)
    mint入出类型 = vNewValue
End Property
Private Function Check是否是药库人员() As Boolean
    Dim rsDepend As ADODB.Recordset
    
    On Error GoTo errHandle
    '判断是不是药库人员使用本模块
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = [2] Or a.站点 is Null) And c.工作性质 = b.名称 " _
            & "  AND Instr('HIJKLMN', b.编码, 1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " _
            & "  And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1]) "
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.用户ID, gstrNodeNo)
                  
    Check是否是药库人员 = (rsDepend.RecordCount <> 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSearch(ByVal FrmMain As Form, ByVal lngMode As Long, ByVal lngStoreid As Long, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strNO开始 As String, _
        ByRef strNO结束 As String, _
        ByRef date填制时间开始 As Date, _
        ByRef date填制时间结束 As Date, _
        ByRef date审核时间开始 As Date, _
        ByRef date审核时间结束 As Date, _
        ByRef lng药品 As Long, _
        ByRef lng库房 As Long, _
        ByRef str填制人 As String, _
        ByRef str审核人 As String, _
        ByRef lng药品分类 As Long, _
        ByRef str剂型 As String, _
        Optional ByRef intTmp As Integer = 0) As String
    mstrFind = ""
    mlngMode = lngMode
    mlngStoreId = lngStoreid
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    
    strNO开始 = SQLCondition.strNO开始
    strNO结束 = SQLCondition.strNO结束
    date填制时间开始 = SQLCondition.date填制时间开始
    date填制时间结束 = SQLCondition.date填制时间结束
    date审核时间开始 = SQLCondition.date审核时间开始
    date审核时间结束 = SQLCondition.date审核时间结束
    lng药品 = SQLCondition.lng药品
    lng库房 = SQLCondition.lng库房
    str审核人 = SQLCondition.str审核人
    str填制人 = SQLCondition.str填制人
    lng药品分类 = SQLCondition.lng药品分类
    str剂型 = SQLCondition.str剂型
    intTmp = SQLCondition.int填制审核一并查询
    
End Function

Private Sub Cbo库房_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    '获取可操作的库房
    Select Case mlngMode
        Case 模块号.药品移库
            str工作性质 = "H,I,J,K,L,M,N"
        Case 模块号.药品领用
            str工作性质 = "O"
    End Select
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo库房.ListCount = 0 Then Exit Sub
    
    If Cbo库房.ListIndex >= 0 Then
        If Val(Cbo库房.Tag) = Cbo库房.ItemData(Cbo库房.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, Cbo库房, Trim(Cbo库房.Text), str工作性质) = False Then
        Exit Sub
    End If
    If Cbo库房.ListIndex >= 0 Then
        Cbo库房.Tag = Cbo库房.ItemData(Cbo库房.ListIndex)
    End If
End Sub

Private Sub Cbo库房_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0

    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Cbo库房_Validate(Cancel As Boolean)
    If Cbo库房.ListCount > 0 Then
        If Cbo库房.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub chkClass_Click()
    If chkClass.Value = 1 Then
        txtClass.Enabled = True
        cmdClass.Enabled = True
    Else
        txtClass.Enabled = False
        cmdClass.Enabled = False
    End If
End Sub

Private Sub chkJiXin_Click()
    If chkJiXin.Value = 1 Then
        txtJiXing.Enabled = True
        cmdJiXin.Enabled = True
    Else
        txtJiXing.Enabled = False
        cmdJiXin.Enabled = False
    End If
End Sub

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmd确定.SetFocus
    End If
End Sub

Private Sub chk审核_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chk审核.Value = 1 Then
            SendKeys vbTab
        Else
            cmd确定.SetFocus
        End If
    End If
    
End Sub

Private Sub chk填制_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
End Sub

Private Sub Chk药品_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        Chk药品.SetFocus
    End If
    
End Sub

Private Sub Chk药品_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Chk移入库房_click()
    Cbo库房.Enabled = IIf(Chk移入库房.Value = 1, True, False)
End Sub

Private Sub Chk移入库房_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Chk移入库房.Value = 1 Then
        Cbo库房.SetFocus
    Else
        Txt填制人.SetFocus
    End If
End Sub
Private Sub chk填制_Click()
    dtp开始时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    dtp结束时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    chkNoStrike.Enabled = IIf(chk填制.Value = 1, True, False)
End Sub

Private Sub chk审核_Click()
    dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
    chkYesStrike.Enabled = IIf(chk审核.Value = 1, True, False)
End Sub

Private Sub Chk药品_Click()
    Txt药品.Enabled = IIf(Chk药品.Value = 1, True, False)
    Cmd药品.Enabled = IIf(Chk药品.Value = 1, True, False)
End Sub



Private Sub cmdClass_Click()
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng库房ID As Long
    Dim Int末级 As Integer
    
    On Error GoTo errHandle
    tvw类别.Left = txtClass.Left
    tvw类别.Top = txtClass.Top + txtClass.Height
    tvw类别.Visible = True
    tvw类别.SetFocus
        
    gstrSQL = "Select 编码, 名称 From 诊疗项目类别 " & _
              "Where Instr([1], 编码, 1) > 0 " & _
              "Order by 编码 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
    With tvw类别
        .Nodes.Clear
'        Set nodTmp = .Nodes.Add(, , "Root", "所有", 2, 2)
        Do While Not rsTmp.EOF
            Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!名称, rsTmp!名称, 2, 2)
            nodTmp.Tag = "Root" & rsTmp!编码
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End With
    
    gstrSQL = "Select ID, 上级ID, 名称, 1 as 末级, decode(类型,1,'西成药',2,'中成药','中草药') as 材质, 类型 " & _
                  "From 诊疗分类目录 " & _
                  "Where 类型 in (1,2,3) " & _
                  "Start With 上级ID IS NULL Connect By Prior ID=上级ID Order by level,ID "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取药品用途分类")
    
    With rsTmp
        If .EOF Then
            Exit Sub
        End If
        
        '将药品用途分类数据装入
        Do While Not .EOF
            Int末级 = IIf(!末级 = 1, 3, 2)
            If IsNull(!上级ID) Then
                Set nodTmp = tvw类别.Nodes.Add("Root" & !材质, 4, "K_" & !id, !名称, Int末级, Int末级)
            Else
                Set nodTmp = tvw类别.Nodes.Add("K_" & !上级ID, 4, "K_" & !id, !名称, Int末级, Int末级)
            End If
            nodTmp.Tag = !类型   '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With

    With tvw类别
        .Nodes(1).Selected = True
        If .Nodes(1).Children <> 0 Then
            Int末级 = 1
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(2).Children <> 0 Then
            Int末级 = 2
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(3).Children <> 0 Then
            Int末级 = 3
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        Else
            Int末级 = 0
            .Nodes(1).Selected = True
            .SelectedItem.Selected = True
        End If
        If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdJiXin_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lng库房ID As Long
    
    lvw剂型.Left = txtJiXing.Left
    lvw剂型.Top = txtJiXing.Top + txtJiXing.Height
    lvw剂型.Visible = True
    lvw剂型.SetFocus
    
    On Error GoTo errHandle
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If lng库房ID <> 0 Then
        '提取该库房现有剂型，供用户选择
        gstrSQL = "Select Distinct J.编码,J.名称 " & _
                  "From 诊疗执行科室 A, 药品特性 B, 药品剂型 J " & _
                  "Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称 And A.执行科室ID=[1] " & _
                  "Order by J.名称 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房ID)
    Else
        gstrSQL = "Select 编码,名称 From 药品剂型 order by 名称 "
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "提取所有药品剂型")
    End If
    
    With rsTmp
        lvw剂型.ListItems.Clear
        Do While Not .EOF
            lvw剂型.ListItems.Add , "K" & !编码, !名称, 1, 1
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd取消_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    '检查数据
    If chkClass.Value = 1 Then
        If txtClass.Tag = 0 Then
            MsgBox "请选择要查询的分类信息！", vbInformation, gstrSysName
            Me.txtClass.SetFocus
            Exit Sub
        End If
    End If
    If chkJiXin.Value = 1 Then
        If txtJiXing.Tag = "" Then
            MsgBox "请选择要查询的剂型信息！", vbInformation, gstrSysName
            Me.txtJiXing.SetFocus
            Exit Sub
        End If
    End If
    If Chk药品.Value = 1 Then
        If Txt药品.Tag = 0 Then
            MsgBox "请选择需查询的药品信息！", vbInformation, gstrSysName
            Me.Txt药品.SetFocus
            Exit Sub
        End If
    End If
    
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        MsgBox "对不起，必须选择一个填制日期或者审核日期!", vbInformation, gstrSysName
        chk填制.SetFocus
        Exit Sub
    End If
    
    mstrFind = ""
    '基本查询条件
    Dim i As Integer
    
    SQLCondition.int填制审核一并查询 = 0
    
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        SQLCondition.int填制审核一并查询 = 1
        If mlngMode <> 1304 Then '药品移库
            If chkStrike.Value = 1 Then
                If mlngMode <> 1306 Then
                    mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                            & " or (A.审核日期 Between [5] And [6]))"
                Else
                    mstrFind = " And 1=1 "
                End If
            Else
                If mlngMode <> 1306 Then
                    mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                            & " or (A.审核日期 Between [5] And [6] and a.记录状态 =1))  "
                Else
                    mstrFind = " And a.记录状态 =1 "
                End If
            End If
        Else
            If chkStrike.Value = 1 Then
                mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                        & " or (A.审核日期 Between [5] And [6]))"
            Else
                If chkNoStrike.Value = 1 And chkYesStrike.Value = 1 Then
                    mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                                & " or (A.审核日期 Between [5] And [6]))"
                ElseIf chkNoStrike.Value = 1 And chkYesStrike.Value = 0 Then
                    mstrFind = "And ((((Mod(a.记录状态, 3) = 0 And 审核日期 Is not Null) or (Mod(a.记录状态, 3) = 2  And 审核日期 Is Null)) and 填制日期 Between [3] and [4] " _
                               & "And Exists (Select 1 From 药品收发记录 B Where a.单据 = b.单据 and a.库房id =b.库房id and a.No = b.No And Mod(b.记录状态, 3) = 2 And b.审核日期 Is Null)) " _
                               & " or (A.审核日期 Between [5] And [6] and (a.记录状态 =1 or mod(A.记录状态,3)=0)" _
                               & "And Not Exists (Select 1 From 药品收发记录 Y Where a.单据 = y.单据 and a.库房id =y.库房id and a.No = y.No And Mod(y.记录状态, 3) = 2)))"
                ElseIf chkNoStrike.Value = 0 And chkYesStrike.Value = 1 Then
                    mstrFind = " and ((A.记录状态=2 or mod(A.记录状态,3)=2 or mod(A.记录状态,3)=0) And A.审核日期 Between [5] And [6] " _
                               & "And Not Exists (Select 1 From 药品收发记录 B Where a.单据 = b.单据 and a.库房id =b.库房id and a.No = b.No And Mod(b.记录状态, 3) = 2 And b.审核日期 Is Null) " _
                               & " or ((a.记录状态 =1 or mod(A.记录状态,3)=0) and (A.填制日期 Between [3] And [4]) and a.审核日期 is null)) "
                Else
                    mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                                & " or (A.审核日期 Between [5] And [6])) and (a.记录状态 =1 or mod(A.记录状态,3)=0) " _
                                & "And Not Exists (Select 1 From 药品收发记录 B Where a.单据 = b.单据 and a.库房id =b.库房id and a.No = b.No And Mod(b.记录状态, 3) = 2)"
                End If
            End If
        End If
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        If mlngMode <> 1304 Then '药品移库
            If chkStrike.Value = 1 Then
                mstrFind = " And A.审核日期 Between [5] And [6] "
            Else
                mstrFind = " And A.审核日期 Between [5] And [6] and a.记录状态 =1 "
                
            End If
        Else
            If chkStrike.Value = 1 Then
                mstrFind = " And A.审核日期 Between [5] And [6] "
            Else
                If chkYesStrike.Value = 1 Then
                    mstrFind = " and (A.记录状态=2 or mod(A.记录状态,3)=2 or mod(A.记录状态,3)=0) And A.审核日期 Between [5] And [6] " _
                               & "And Not Exists (Select 1 From 药品收发记录 B Where a.单据 = b.单据 and a.库房id =b.库房id and a.No = b.No And Mod(b.记录状态, 3) = 2 And b.审核日期 Is Null)"
                Else
                    mstrFind = " And A.审核日期 Between [5] And [6] and (a.记录状态 =1 or mod(A.记录状态,3)=0) " _
                               & "And Not Exists (Select 1 From 药品收发记录 B Where a.单据 = b.单据 and a.库房id =b.库房id and a.No = b.No And Mod(b.记录状态, 3) = 2)"
                End If
            End If
        End If
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        If mlngMode <> 1304 Then '药品移库
            mstrFind = " And (A.填制日期 Between [3] And [4]) and 审核日期 is null "
        Else
            If chkNoStrike.Value = 1 Then
                mstrFind = "And ((Mod(a.记录状态, 3) = 0 And 审核日期 Is not Null) or (Mod(a.记录状态, 3) = 2  And 审核日期 Is Null)) and 填制日期 Between [3] and [4] " _
                           & "And Exists (Select 1 From 药品收发记录 B Where a.单据 = b.单据 and a.库房id =b.库房id and a.No = b.No And Mod(b.记录状态, 3) = 2 And b.审核日期 Is Null)"
            Else
                mstrFind = " And (a.记录状态 =1 or mod(A.记录状态,3)=0) and (A.填制日期 Between [3] And [4]) and 审核日期 is null "
            End If
        End If
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, intNO, lng库房ID)
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, intNO, lng库房ID)
    End If
    
    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= [1] And A.No <=[2] "
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= [1] "
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= [2] "
    
    SQLCondition.strNO开始 = Me.txt开始No
    SQLCondition.strNO结束 = Me.txt结束NO
    SQLCondition.date填制时间开始 = CDate(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date审核时间开始 = CDate(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date审核时间结束 = CDate(Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59")
    
    '扩展查询条件
    SQLCondition.lng药品分类 = 0
    SQLCondition.str剂型 = ""
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If Chk药品.Value = 1 Then
        mstrFind = mstrFind & " And A.药品ID + 0 =[7] "
    End If
    
    If mlngMode = 模块号.其他出库 Then
        If Chk移入库房.Value = 1 Then mstrFind = mstrFind & " And A.入出类别ID=[8]"
    ElseIf mlngMode = 模块号.药品移库 Then
        If mint入出类型 = -1 Then
            If Chk移入库房.Value = 1 Then mstrFind = mstrFind & " And A.对方部门ID + 0 =[8]"
        Else
            If Chk移入库房.Value = 1 Then mstrFind = mstrFind & " And A.库房ID+0=[8]"
        End If
    Else
        If Chk移入库房.Value = 1 Then mstrFind = mstrFind & " And A.对方部门ID + 0 =[8]"
    End If
    If Me.Txt审核人 <> "" Then mstrFind = mstrFind & " And A.审核人 like [10] "
    If Me.Txt填制人 <> "" Then mstrFind = mstrFind & " And A.填制人 like [9] "
    
    If chkClass.Value = 1 Then
        SQLCondition.lng药品分类 = Val(txtClass.Tag)
    End If
    If chkJiXin.Value = 1 Then
        SQLCondition.str剂型 = txtJiXing.Tag
    End If
    SQLCondition.lng药品 = Val(Txt药品.Tag)
    If Cbo库房.Visible Then
        SQLCondition.lng库房 = Cbo库房.ItemData(Cbo库房.ListIndex)
    End If
    SQLCondition.str审核人 = Me.Txt审核人 & "%"
    SQLCondition.str填制人 = Me.Txt填制人 & "%"
    
    Unload Me
End Sub

Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "药品移库管理", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , True)
    
'    Set RecReturn = Frm药品选择器.ShowME(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品id
    
    If Chk移入库房.Visible = True Then
        Chk移入库房.SetFocus
    Else
        Txt填制人.SetFocus
    End If
End Sub

Private Sub dtp结束时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp开始时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.dtp结束时间(Index).SetFocus
End Sub

Private Sub Form_Load()
    Me.dtp结束时间(0) = Sys.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    mblnStock = Check是否是药库人员
    mstrMatch = IIf(zlDatabase.GetPara("输入匹配", , , 0) = "0", "%", "")
    
    Me.Txt药品.Tag = 0
    sstFilter.Tab = 0
    Select Case mlngMode
        Case 模块号.药品移库
            mint冲销申请 = Val(zlDatabase.GetPara("冲销申请", glngSys, 模块号.药品移库))
            If mint入出类型 = -1 Then
                Chk移入库房.Caption = "移入库房"
            Else
                Chk移入库房.Caption = "移出库房"
            End If
            If mint冲销申请 = 0 Then    '不需要申请
                chkStrike.Visible = True
                chkNoStrike.Visible = False
                chkYesStrike.Visible = False
            Else
                chkStrike.Visible = False
                chkNoStrike.Visible = True
                chkYesStrike.Visible = True
            End If
        Case 模块号.药品领用
            Chk移入库房.Caption = "领用部门"
        Case 模块号.其他出库
            Chk移入库房.Caption = "入出类别"
    End Select
    BlnAdvance = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Booker"
                Txt填制人.SetFocus
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
            Case "Verify"
                Txt审核人.SetFocus
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
        End Select
        Cancel = True
    End If
    Call ReleaseSelectorRS
End Sub

Private Sub lvw剂型_DblClick()
    Dim i As Integer
    Dim strName As String
    
    With lvw剂型
        For i = 1 To .ListItems.count
            If .ListItems(i).Checked = True Then
                strName = strName & .ListItems(i).Text & ","
            End If
        Next
        lvw剂型.Visible = False
        txtJiXing.Tag = strName
        txtJiXing.Text = strName
    End With
End Sub

Private Sub lvw剂型_LostFocus()
    lvw剂型.Visible = False
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    Dim rsDepartment As New Recordset
    Dim strStock As String
    Dim str站点限制 As String
    
    On Error GoTo errHandle
    str站点限制 = GetDeptStationNode(mlngStoreId)
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
            If Cbo库房.ListCount < 1 Then
                Select Case mlngMode
                    Case 1304
                        strStock = "HIJKLMN"
                        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                                & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                                & "Where " & IIf(str站点限制 <> "", " (a.站点 = [3] or a.站点 is null) AND ", "") & " c.工作性质 = b.名称 " _
                                & "  AND Instr([1],b.编码,1) > 0 " _
                                & "  AND a.id = c.部门id " _
                                & "  AND a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd')"
                    Case 1305
                        strStock = "O"
                        gstrSQL = " Select C.ID " & _
                            " From 部门性质说明 A,部门性质分类 B,部门表 C " & _
                            " Where " & IIf(str站点限制 <> "", " (c.站点 = [3] or c.站点 is null) AND ", "") & " A.工作性质=B.名称 And A.部门ID=C.ID " & _
                            "   AND TO_CHAR(C.撤档时间, 'yyyy-MM-dd')='3000-01-01' And B.编码='O'" & _
                            "   And C.ID IN (Select 部门ID From 部门人员 Where 人员ID=[2])"
                        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                            & "Where " & IIf(str站点限制 <> "", " (a.站点 = [3] or a.站点 is null) AND ", "") & " c.工作性质 = b.名称 " _
                            & "  AND Instr([1],b.编码,1) > 0 " _
                            & "  AND a.id = c.部门id " _
                            & "  AND a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd')" _
                            & IIf(mblnStock, "", " And a.ID IN (Select Distinct 领用部门ID From 药品领用控制 Where 领用部门ID IN (" & gstrSQL & "))")
                    Case 1306
                       gstrSQL = "SELECT b.Id,b.名称 " _
                               & "FROM 药品单据性质 A, 药品入出类别 B " _
                               & "Where A.类别id = B.ID AND A.单据 = 11 "
                    Case 1303, 1307
                        If Chk移入库房.Visible = True Then
                            Chk移入库房.Visible = False
                            Cbo库房.Visible = False
                        End If
                        Exit Sub
                End Select
                Set rsDepartment = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock, UserInfo.用户ID, gstrNodeNo)
            
                With Cbo库房
                    Do While Not rsDepartment.EOF
                        .AddItem rsDepartment.Fields(1)
                        .ItemData(.NewIndex) = rsDepartment.Fields(0)
                        rsDepartment.MoveNext
                    Loop
                    If .ListCount > 0 Then .ListIndex = 0
                End With
                rsDepartment.Close
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvw类别_DblClick()
    With tvw类别
        If .SelectedItem.Text <> "" Then
            If .SelectedItem.Key Like "Root*" Then Exit Sub
            txtClass.Tag = Mid(.SelectedItem.Key, InStr(1, .SelectedItem.Key, "_") + 1)
            txtClass.Text = .SelectedItem.Text
            .Visible = False
        End If
    End With
End Sub

Private Sub tvw类别_LostFocus()
    tvw类别.Visible = False
End Sub

Private Sub txtClass_GotFocus()
    txtClass.SelStart = 0
    txtClass.SelLength = 100
End Sub

Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTemp As String
    Dim nodTmp As Node
    Dim rsTmp As ADODB.Recordset
    Dim lng库房ID As Long
    Dim Int末级 As Integer
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        strTemp = UCase(Trim(txtClass.Text))
        If strTemp <> "" Then
            tvw类别.Left = txtClass.Left
            tvw类别.Top = txtClass.Top + txtClass.Height
            tvw类别.Visible = True
            tvw类别.SetFocus
            
            gstrSQL = "Select 编码, 名称 From 诊疗项目类别 " & _
                      "Where Instr([1], 编码, 1) > 0 " & _
                      "Order by 编码 "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
            
            With tvw类别
                .Nodes.Clear
                Do While Not rsTmp.EOF
                    Set nodTmp = .Nodes.Add(, , "Root" & rsTmp!名称, rsTmp!名称, 2, 2)
                    nodTmp.Tag = "Root" & rsTmp!编码
                    rsTmp.MoveNext
                Loop
                rsTmp.Close
            End With
            
            gstrSQL = "Select ID, 上级id, 名称, 1 As 末级, 材质, 类型" & _
                        " From (Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 材质, 类型" & _
                               " From 诊疗分类目录" & _
                               " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
                                     " (编码 Like [1] Or 名称 Like [1] Or 简码 Like [1])" & _
                               " Start With 上级id Is Null" & _
                               " Connect By Prior ID = 上级id" & _
                               " Union " & _
                               " Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 材质, 类型" & _
                               " From 诊疗分类目录" & _
                               " Where ID In (Select 上级id" & _
                                            " From 诊疗分类目录" & _
                                            " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
                                                  " (编码 Like [1] Or 名称 Like [1] Or 简码 Like [1])))" & _
                        " Start With 上级id Is Null" & _
                        " Connect By Prior ID = 上级id" & _
                        " Order By Level, ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "查询品种", "%" & strTemp & mstrMatch)
            
            With rsTmp
                If .EOF Then
                    Exit Sub
                End If
                
                '将药品用途分类数据装入
                Do While Not .EOF
                    Int末级 = IIf(!末级 = 1, 3, 2)
                    If IsNull(!上级ID) Then
                        Set nodTmp = tvw类别.Nodes.Add("Root" & !材质, 4, "K_" & !id, !名称, Int末级, Int末级)
                    Else
                        Set nodTmp = tvw类别.Nodes.Add("K_" & !上级ID, 4, "K_" & !id, !名称, Int末级, Int末级)
                    End If
                    nodTmp.Tag = !类型   '存放分类类型:1-西成药,2-中成药,3-中草药
                    .MoveNext
                Loop
            End With
        
            With tvw类别
                .Nodes(1).Selected = True
                If .Nodes(1).Children <> 0 Then
                    Int末级 = 1
                    .Nodes(Int末级).Child.Selected = True
                    .SelectedItem.Selected = True
                ElseIf .Nodes(2).Children <> 0 Then
                    Int末级 = 2
                    .Nodes(Int末级).Child.Selected = True
                    .SelectedItem.Selected = True
                ElseIf .Nodes(3).Children <> 0 Then
                    Int末级 = 3
                    .Nodes(Int末级).Child.Selected = True
                    .SelectedItem.Selected = True
                Else
                    Int末级 = 0
                    .Nodes(1).Selected = True
                    .SelectedItem.Selected = True
                End If
                If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
            End With
        End If
    ElseIf KeyCode = vbKeyDelete Then
        txtClass.Tag = 0
    End If
    
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtJiXing_GotFocus()
    txtJiXing.SelStart = 0
    txtJiXing.SelLength = 100
End Sub

Private Sub txtJiXing_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim lng库房ID As Long
    Dim strFind As String
    
    If KeyCode = vbKeyReturn Then
        strFind = UCase(Trim(txtJiXing.Text))
        If strFind = "" Then Exit Sub
        
        lvw剂型.Left = txtJiXing.Left
        lvw剂型.Top = txtJiXing.Top + txtJiXing.Height
        lvw剂型.Visible = True
        lvw剂型.SetFocus
        
        On Error GoTo errHandle
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        If lng库房ID <> 0 Then
            '提取该库房现有剂型，供用户选择
            gstrSQL = "Select Distinct J.编码,J.名称 " & _
                      "From 诊疗执行科室 A, 药品特性 B, 药品剂型 J " & _
                      "Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称 And A.执行科室ID=[1] and (j.编码 like [2] or j.名称 like [2] or j.简码 like [2]) " & _
                      "Order by J.名称 "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房ID, "%" & strFind & mstrMatch)
        Else
            gstrSQL = "Select 编码,名称 From 药品剂型 where 编码 like [1] or 名称 like [1] or 简码 like [1] order by 名称 "
            Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, "提取所有药品剂型", "%" & strFind & mstrMatch)
        End If
        
        With rsTmp
            lvw剂型.ListItems.Clear
            Do While Not .EOF
                lvw剂型.ListItems.Add , "K" & !编码, !名称, 1, 1
                .MoveNext
            Loop
        End With
    ElseIf KeyCode = vbKeyDelete Then
        txtJiXing.Tag = 0
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    If mlngMode = 1307 Then
        If mfrmMain.TabShow.Tab = 1 Then
            '盘点表
            intNO = 29
        Else
            '盘点记录单
            intNO = 62
        End If
    End If
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, intNO, lng库房ID)
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    If mlngMode = 1307 Then
        If mfrmMain.TabShow.Tab = 1 Then
            '盘点表
            intNO = 29
        Else
            '盘点记录单
            intNO = 62
        End If
    End If
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, intNO, lng库房ID)
        End If
        Me.txt结束NO.SetFocus
    End If
End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then cmd确定.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            cmd确定.SetFocus
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)

        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取审核人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", _
                        Me.Txt审核人 & "%", gstrNodeNo)

        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt审核人.Top + Txt审核人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt审核人.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt审核人.Top - Txt审核人.Height - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt审核人 = IIf(IsNull(!姓名), "", !姓名)
                cmd确定.SetFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Me.Txt审核人.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt填制人.Text) = "" Then
            Txt审核人.SetFocus
            Exit Sub
        End If
        Txt填制人.Text = UCase(Txt填制人.Text)

        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取填制人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", _
                        Me.Txt填制人 & "%", gstrNodeNo)

        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt填制人.Top + Txt填制人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt填制人.Left
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt填制人.Top - Txt填制人.Height - 50
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt填制人 = IIf(IsNull(!姓名), "", !姓名)
                Me.Txt审核人.SetFocus
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt填制人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药品.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fra附加条件.Left + Txt药品.Left
    sngTop = Me.Top + fra附加条件.Top + Txt药品.Top + Txt药品.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt药品.Height - 3630
    End If
    
    strkey = Trim(Txt药品.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "药品移库管理", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , True)
    
'    Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品id
    
    If Chk移入库房.Visible = True Then
        Chk移入库房.SetFocus
    Else
        Txt填制人.SetFocus
    End If
    
End Sub

Private Sub Txt药品_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Booker"
                    Txt填制人 = .TextMatrix(.Row, 2)
                    Txt审核人.SetFocus
                Case "Verify"
                    Txt审核人 = .TextMatrix(.Row, 2)
                    cmd确定.SetFocus
                
            End Select
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

