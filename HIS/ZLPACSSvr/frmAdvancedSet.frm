VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdvancedSet 
   Caption         =   "高级设置"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   300
      Left            =   5520
      TabIndex        =   2
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   5880
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "存储高级设置"
      TabPicture(0)   =   "frmAdvancedSet.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmAutoRoutSet"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Worklist高级设置"
      TabPicture(1)   =   "frmAdvancedSet.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdResetWLResult"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkForceResult"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkModel"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtDayInterval"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtDayInterval 
         Height          =   300
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   43
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox chkModel 
         Caption         =   "按检查设备过滤"
         Height          =   225
         Left            =   120
         TabIndex        =   42
         Top             =   525
         Width           =   1755
      End
      Begin VB.CheckBox chkForceResult 
         Caption         =   "使用强制结果"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   900
         Width           =   1515
      End
      Begin VB.CommandButton cmdResetWLResult 
         Caption         =   "恢复默认结果"
         Height          =   350
         Left            =   2760
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Frame8 
         Caption         =   "结果集设置"
         Height          =   4215
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   8175
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   2055
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   3625
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
         Begin VB.Frame frmSetResult 
            Height          =   1575
            Left            =   120
            TabIndex        =   32
            Top             =   2520
            Width           =   7935
            Begin VB.CheckBox chkMWLItem 
               Caption         =   "选择使用该结果"
               Height          =   180
               Left            =   120
               TabIndex        =   45
               Top             =   0
               Width           =   1575
            End
            Begin VB.TextBox txtResult 
               Height          =   300
               Index           =   0
               Left            =   1200
               TabIndex        =   36
               Top             =   720
               Width           =   5775
            End
            Begin VB.CheckBox chkResult 
               Caption         =   "是否递增"
               Height          =   255
               Left            =   6360
               TabIndex        =   35
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtResult 
               Height          =   300
               Index           =   1
               Left            =   1200
               TabIndex        =   34
               Top             =   1080
               Width           =   6135
            End
            Begin VB.CommandButton cmdBuildResult 
               Appearance      =   0  'Flat
               Caption         =   "…"
               Height          =   235
               Index           =   0
               Left            =   6990
               MaskColor       =   &H80000000&
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   765
               Width           =   315
            End
            Begin VB.Label lblResult 
               Caption         =   "结果集："
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   360
               Width           =   7215
            End
            Begin VB.Label Label11 
               Caption         =   "返回值"
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   743
               Width           =   735
            End
            Begin VB.Label Label12 
               Caption         =   "强制结果值"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   1110
               Width           =   975
            End
         End
         Begin VB.CheckBox chkUseResult 
            Caption         =   "选择使用该结果："
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   3000
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "基本参数"
         Height          =   855
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   8175
         Begin VB.ComboBox cboStoreDevice 
            Height          =   300
            ItemData        =   "frmAdvancedSet.frx":0038
            Left            =   1275
            List            =   "frmAdvancedSet.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cboEncode 
            Height          =   300
            ItemData        =   "frmAdvancedSet.frx":0054
            Left            =   5040
            List            =   "frmAdvancedSet.frx":0061
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   360
            Width           =   2835
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "存储设备(&F)"
            Height          =   180
            Index           =   8
            Left            =   240
            TabIndex        =   29
            Top             =   405
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "压缩方式(&Y)"
            Height          =   180
            Index           =   0
            Left            =   3960
            TabIndex        =   28
            Top             =   405
            Width           =   990
         End
      End
      Begin VB.Frame frmAutoRoutSet 
         Caption         =   "自动路由设置"
         Height          =   2145
         Left            =   -74880
         TabIndex        =   14
         Top             =   3360
         Width           =   8175
         Begin VB.ComboBox cboDestination 
            Height          =   300
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1305
            Width           =   1605
         End
         Begin VB.ComboBox cboCondition 
            Height          =   300
            Index           =   1
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   1365
         End
         Begin VB.ComboBox cboCondition 
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   6450
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   825
            Width           =   1365
         End
         Begin VB.OptionButton optType 
            Caption         =   "检查设备(&R)"
            Height          =   255
            Index           =   2
            Left            =   5040
            TabIndex        =   19
            Top             =   855
            Width           =   1335
         End
         Begin VB.OptionButton optType 
            Caption         =   "影像类别(&S)"
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   18
            Top             =   375
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelAutoRouting 
            Caption         =   "删除(&D)"
            Height          =   350
            Left            =   6960
            TabIndex        =   17
            Top             =   1680
            Width           =   1100
         End
         Begin VB.CommandButton cmdModiAutoRouting 
            Caption         =   "修改(&M)"
            Height          =   350
            Left            =   5880
            TabIndex        =   16
            Top             =   1680
            Width           =   1100
         End
         Begin VB.CommandButton cmdAddAutoRouting 
            Caption         =   "添加(&A)"
            Height          =   350
            Left            =   4800
            TabIndex        =   15
            Top             =   1680
            Width           =   1100
         End
         Begin MSFlexGridLib.MSFlexGrid MSFAutoRout 
            Height          =   1845
            Left            =   150
            TabIndex        =   23
            Top             =   270
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   3254
            _Version        =   393216
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "目的设备(&B)"
            Height          =   180
            Left            =   5070
            TabIndex        =   24
            Top             =   1365
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "自动匹配设置"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   3
         Top             =   1440
         Width           =   8175
         Begin VB.Frame Frame4 
            Caption         =   "图像项目"
            Height          =   1455
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2685
            Begin VB.OptionButton optImgMatch 
               Caption         =   "Patient ID"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   13
               Top             =   360
               Width           =   1335
            End
            Begin VB.OptionButton optImgMatch 
               Caption         =   "Accession Number"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   12
               Top             =   720
               Width           =   1815
            End
            Begin VB.OptionButton optImgMatch 
               Caption         =   "Patient Name"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   11
               Top             =   1080
               Width           =   1455
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "数据库项目"
            Height          =   1455
            Left            =   2880
            TabIndex        =   6
            Top             =   240
            Width           =   2805
            Begin VB.OptionButton optMatch 
               Caption         =   "检查号"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   9
               ToolTipText     =   "按检查号将病人和接收的影像进行匹配"
               Top             =   360
               Width           =   1065
            End
            Begin VB.OptionButton optMatch 
               Caption         =   "病人标识号（门诊/住院号）"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   8
               ToolTipText     =   "按病人标识号将病人和接收的影像进行匹配"
               Top             =   720
               Width           =   2655
            End
            Begin VB.OptionButton optMatch 
               Caption         =   "检查标识号"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   7
               ToolTipText     =   "按检查标识号将病人和接收的影像进行匹配"
               Top             =   1080
               Width           =   1335
            End
         End
         Begin VB.CheckBox chkMatchStudyUID 
            Caption         =   "启用 ""检查UID"" 匹配"
            Height          =   350
            Left            =   5880
            TabIndex        =   5
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox chkImageType 
            Caption         =   "根据图像类型拆分序列"
            Height          =   350
            Left            =   5880
            TabIndex        =   4
            Top             =   1200
            Width           =   2175
         End
      End
      Begin VB.Label Label9 
         Caption         =   "检索最近        天的申请"
         Height          =   195
         Left            =   2730
         TabIndex        =   44
         Top             =   540
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmAdvancedSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlngServiceID As Long           '服务ID
Private mstrServiceType As String   '服务类型

Private aDevices() As Variant       '存储设备列表


Public Sub ShowMe(parent As Object, strServiceType As String, lngServiceID As Long)
    mlngServiceID = lngServiceID
    mstrServiceType = strServiceType
    If mstrServiceType = ZLPACS_存储服务 Then
        Me.SSTab1.TabVisible(0) = True
        Me.SSTab1.TabVisible(1) = False
    ElseIf UCase(mstrServiceType) = UCase(ZLPACS_Worklist服务) Then
        Me.SSTab1.TabVisible(0) = False
        Me.SSTab1.TabVisible(1) = True
    End If
    
    Me.Show vbModal, parent
End Sub

Private Sub cmdAddAutoRouting_Click()
    Dim iType As Integer
    
    '检查输入是否有效
    iType = IIf(optType(1).value = True, 1, 2)
    If cboDestination.Text = "" Then MsgBox "请输入自动路由的目的设备。": Exit Sub
    If cboCondition(iType).Text = "" Then MsgBox IIf(iType = 1, "请输入影像类别", "请输入检查设备"): Exit Sub
    
    On Error GoTo errHand
    '插入数据
    gstrSQL = "Zl_影像自动路由设置_INSERT(" & mlngServiceID & ",'" & iType & "','" & cboCondition(iType).Text & "','" & _
                    GetDeviceNameNum(aDevices, cboDestination.Text, 1) & "')"
                        
    ExecuteProcedure "增加自动路由设置"
    '刷新列表
    Call subFillMSFAutoRouting(Me.MSFAutoRout.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelAutoRouting_Click()
    
    On Error GoTo errHand
    '删除数据
    gstrSQL = "Zl_影像自动路由设置_DELETE(" & Me.MSFAutoRout.TextMatrix(Me.MSFAutoRout.RowSel, 3) & ")"
                        
    ExecuteProcedure "删除自动路由设置"
    '刷新列表
    Call subFillMSFAutoRouting
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdModiAutoRouting_Click()
    Dim iType As Integer
    
    '检查输入是否有效
    iType = IIf(optType(1).value = True, 1, 2)
    If cboDestination.Text = "" Then MsgBox "请输入自动路由的目的设备。": Exit Sub
    If cboCondition(iType).Text = "" Then MsgBox IIf(iType = 1, "请输入影像类别", "请输入检查设备"): Exit Sub
    
    On Error GoTo errHand
    '修改数据
    gstrSQL = "Zl_影像自动路由设置_UPDATE(" & Me.MSFAutoRout.TextMatrix(Me.MSFAutoRout.RowSel, 3) & "," & mlngServiceID & ",'" & iType & "','" & cboCondition(iType).Text & "','" & _
                    GetDeviceNameNum(aDevices, cboDestination.Text, 1) & "')"
                        
    ExecuteProcedure "修改自动路由设置"
    '刷新列表
    Call subFillMSFAutoRouting(Me.MSFAutoRout.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub CmdOK_Click()
    Call subSaveServiceParas
    If mstrServiceType = ZLPACS_存储服务 Then
        '保存自动匹配设置
        Call subSaveMatch
    End If
    Unload Me
End Sub

Private Sub subSaveServiceParas()
    '保存基本参数
    Dim strValue As String
    If mstrServiceType = ZLPACS_存储服务 Then
        '存储设备号
        strValue = aDevices(0, cboStoreDevice.ListIndex)
        subSaveServicePara ZLPACS_存储设备号, strValue
        '压缩方式
        subSaveServicePara ZLPACS_压缩方式, cboEncode.ListIndex
        '启用检查UID匹配
        subSaveServicePara ZLPACS_启用检查UID匹配, chkMatchStudyUID.value
        '按图像类型拆分序列
        subSaveServicePara ZLPACS_按图像类型拆分序列, chkImageType.value
    Else
        '按检查设备过滤
        subSaveServicePara ZLPACS_MWL按设备过滤, chkModel.value
        '检索天数
        subSaveServicePara ZLPACS_MWL检索天数, txtDayInterval.Text
        '使用强制结果
        subSaveServicePara ZLPACS_MWL用强制结果, chkForceResult.value
    End If
End Sub

Private Sub subSaveServicePara(strParaName As String, strParaValue As String)
    On Error GoTo errHand
    '插入数据
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngServiceID & ",'" & strParaName & "','" & strParaValue & "')"
                        
    ExecuteProcedure "保存DICOM服务参数"
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub subSaveMatch()
    Dim intDBItem As Integer
    Dim intImageItem As Integer
    
    For intDBItem = 0 To optMatch.count - 1
        If optMatch(intDBItem).value Then Exit For
    Next
    If intDBItem > optMatch.count - 1 Then intDBItem = 0
    
    For intImageItem = 0 To optImgMatch.count - 1
        If optImgMatch(intImageItem).value Then Exit For
    Next
    If intImageItem > optImgMatch.count - 1 Then intImageItem = 0
    
    On Error GoTo errHand
    '插入数据
    gstrSQL = "Zl_影像自动匹配设置_SAVE(" & mlngServiceID & ",'" & intImageItem & "','" & intDBItem & "')"
                        
    ExecuteProcedure "保存自动匹配设置"
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Load()
    If mstrServiceType = ZLPACS_存储服务 Then
        '加载 存储设置
        '加载存储设备
        Call subFillcboStoreDevice
        '加载基本参数
        Call subReadPara(1)
        '加载匹配方式
        Call subFillMatch
        '加载自动路由
        Call subFillAutoRoutDevice
        Call subFillMSFAutoRouting
    Else
        '加载WORKLIST设置
        '加载基本参数
        Call subReadPara(2)
        '加载结果集设置
        
    End If
End Sub

Private Sub subFillMSFAutoRouting(Optional iRow As Integer = 1)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRowPos As Long
    
    strSQL = "Select 自动路由ID,服务ID,条件类型,条件值, 目的设备号 From 影像自动路由设置 Where 服务ID =[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "读取自动路由设置", mlngServiceID)
    
    With MSFAutoRout
        .Clear
        .Rows = 1
        .Cols = 4
        .ColWidth(1) = 2500
        .ColWidth(3) = 0
        .TextMatrix(0, 0) = "条件类型"
        .TextMatrix(0, 1) = "条件内容"
        .TextMatrix(0, 2) = "目的设备"
        .TextMatrix(0, 3) = "ID"
        lngRowPos = 1
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(lngRowPos, 0) = IIf(rsTmp!条件类型 = 1, "影像类别", "检查设备")
            .TextMatrix(lngRowPos, 1) = rsTmp!条件值
            .TextMatrix(lngRowPos, 2) = GetDeviceNameNum(aDevices, rsTmp!目的设备号, 0)
            .TextMatrix(lngRowPos, 3) = rsTmp!自动路由ID
            lngRowPos = .Rows
            rsTmp.MoveNext
        Wend
    End With
    
    Call subClickMSFAutoRouting(iRow)
End Sub

Private Sub subClickMSFAutoRouting(Optional iRow As Integer = 1)

    If iRow > Me.MSFAutoRout.Rows Then iRow = 1

    If Me.MSFAutoRout.Rows > 1 Then
        Me.MSFAutoRout.Row = iRow - 1
        Me.MSFAutoRout.RowSel = iRow
        Call MSFAutoRout_Click
    End If
End Sub

Private Function GetDeviceNameNum(aSource() As Variant, ByVal SeekString As String, iType As Integer) As String
    '获取设备的名称或设备号
    'iType=0---输入SeekString为设备号，返回设备名。
    'iType=1---输入SeekString为设备名，返回设备号。
    Dim i As Long
    For i = 0 To UBound(aSource, 2)
        If aSource(iType, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then GetDeviceNameNum = "": Exit Function
    GetDeviceNameNum = IIf(iType = 1, aSource(0, i), aSource(1, i))
End Function

Private Sub subFillMatch()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select 服务ID,图像项,数据库项 From 影像自动匹配设置 Where 服务ID = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "读取自动匹配设置", mlngServiceID)
    
    If Not rsTmp.EOF Then
        optImgMatch(Val(rsTmp!图像项)).value = True
        optMatch(Val(rsTmp!数据库项)).value = True
    Else
        optImgMatch(0).value = True
        optMatch(0).value = True
    End If
End Sub

Private Sub subFillcboStoreDevice()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型= [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "读取存储设备", 1)
    If rsTmp.EOF Then
        MsgBox "未定义影像存储设备，请到影像设备目录中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    aDevices = rsTmp.GetRows
    rsTmp.MoveFirst
    
    Me.cboStoreDevice.Clear
    Do While Not rsTmp.EOF
        cboStoreDevice.AddItem Nvl(rsTmp(1))
        '填充自动路由设置中的目的设备下拉列表
        cboDestination.AddItem Nvl(rsTmp(1))
        rsTmp.MoveNext
    Loop
End Sub

Private Sub subReadPara(intType As Integer)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select  a.服务参数ID,a.服务ID,a.参数名称,a.参数值 From 影像dicom服务参数 a Where a.服务ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "读取服务参数", mlngServiceID)
    
    If intType = 1 Then     '存储服务
        While Not rsTmp.EOF
            Select Case rsTmp!参数名称
            Case ZLPACS_存储设备号
                cboStoreDevice.ListIndex = GetComboxIndex(aDevices, Nvl(rsTmp!参数值))
                cboStoreDevice.Tag = 1
            Case ZLPACS_压缩方式
                cboEncode.ListIndex = Nvl(rsTmp!参数值, 0)
                cboEncode.Tag = 1
            Case ZLPACS_启用检查UID匹配
                chkMatchStudyUID.value = Nvl(rsTmp!参数值, 0)
                chkMatchStudyUID.Tag = 1
            Case ZLPACS_按图像类型拆分序列
                chkImageType.value = Nvl(rsTmp!参数值, 0)
                chkImageType.Tag = 1
            End Select
            rsTmp.MoveNext
        Wend
        '处理没有参数设置的项目，设置成默认值
        If cboStoreDevice.Tag = "" Then cboStoreDevice.ListIndex = 0
        If cboEncode.Tag = "" Then cboEncode.ListIndex = 0
        If chkMatchStudyUID.Tag = "" Then chkMatchStudyUID.value = 0
        If chkImageType.Tag = "" Then chkImageType.value = 0
    ElseIf intType = 2 Then 'worklist服务
        While Not rsTmp.EOF
            Select Case rsTmp!参数名称
            Case ZLPACS_MWL检索天数
                txtDayInterval.Text = Nvl(rsTmp!参数值, 3)
                txtDayInterval.Tag = 1
            Case ZLPACS_MWL按设备过滤
                chkModel.value = Nvl(rsTmp!参数值, 0)
                chkModel.Tag = 1
            Case ZLPACS_MWL用强制结果
                chkForceResult.value = Nvl(rsTmp!参数值, 0)
                chkForceResult.Tag = 1
            End Select
            rsTmp.MoveNext
        Wend
        '处理没有参数设置的项目，设置成默认值
        If txtDayInterval.Tag = "" Then txtDayInterval.Text = "3"
        If chkModel.Tag = "" Then chkModel.value = 0
        If chkForceResult.Tag = "" Then chkForceResult.value = 0
    End If
End Sub

Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function

Private Sub MSFAutoRout_Click()
    Dim iSelected As Integer
    
    With MSFAutoRout
        iSelected = .RowSel
        '填写条件类型
        Me.optType(IIf(.TextMatrix(iSelected, 0) = "影像类别", 1, 2)).value = True
        '填写条件值
        Me.cboCondition(IIf(.TextMatrix(iSelected, 0) = "影像类别", 1, 2)).Text = .TextMatrix(iSelected, 1)
        '填写目的设备号
        Me.cboDestination = .TextMatrix(iSelected, 2)
    End With
End Sub

Private Sub subFillAutoRoutDevice()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    '填充自动路由设置中，影像类别，和检查设备列表
    strSQL = "Select 编码 From 影像检查类别"
    Set rsTmp = OpenSQLRecord(strSQL, "高级设置")
    Do While Not rsTmp.EOF
        cboCondition(1).AddItem rsTmp(0)
        rsTmp.MoveNext
    Loop
    
    strSQL = "Select distinct 检查设备 From 影像检查记录"
    Set rsTmp = OpenSQLRecord(strSQL, "高级设置")
    Do While Not rsTmp.EOF
        cboCondition(2).AddItem Nvl(rsTmp(0))
        rsTmp.MoveNext
    Loop
End Sub

Private Sub optType_Click(Index As Integer)
    Me.cboCondition(Index).Enabled = True
    Me.cboCondition(IIf(Index = 1, 2, 1)).Enabled = False
End Sub

Private Sub txtDayInterval_KeyPress(KeyAscii As Integer)
    If (Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
