VERSION 5.00
Begin VB.Form frmDeviceReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设备信息"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7575
   Icon            =   "frmDeviceReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7575
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraFunc 
      Caption         =   "功能与业务"
      Height          =   2295
      Left            =   3840
      TabIndex        =   19
      Top             =   960
      Width           =   3615
      Begin VB.ComboBox cboDispense 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkDispensing 
         Caption         =   "启用药品处方发药"
         Height          =   255
         Left            =   1200
         TabIndex        =   22
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "配药功能："
         Height          =   180
         Index           =   11
         Left            =   240
         TabIndex        =   20
         Top             =   390
         Width           =   900
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发送功能："
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   21
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblDevice 
         BackStyle       =   0  'Transparent
         Caption         =   "  指定配药功能在HIS哪个业务进行药品明细上传"
         Height          =   420
         Index           =   13
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label lblDevice 
         BackStyle       =   0  'Transparent
         Caption         =   "  指定发送功能是否在处方发药进行"
         Height          =   420
         Index           =   14
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   3000
      End
   End
   Begin VB.Frame fraService 
      Caption         =   "服务对象"
      Height          =   735
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton optObject 
         Caption         =   "住院"
         Height          =   180
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optObject 
         Caption         =   "门诊"
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   6240
      TabIndex        =   25
      Top             =   3360
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   360
      Left            =   5040
      TabIndex        =   24
      Top             =   3360
      Width           =   1110
   End
   Begin VB.Frame fraDevice 
      Caption         =   "基本信息"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox cboLink 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtManufacturer 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtModel 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optState 
         Caption         =   "禁用"
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Top             =   750
         Width           =   855
      End
      Begin VB.OptionButton optState 
         Caption         =   "启用"
         Height          =   180
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   750
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "使用药房："
         Height          =   180
         Index           =   7
         Left            =   360
         TabIndex        =   14
         Top             =   2550
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "制造商："
         Height          =   180
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   2190
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "型号："
         Height          =   180
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Top             =   1830
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称："
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编码："
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设备状态："
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   750
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "连接名："
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   390
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmDeviceReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long

Public Sub ShowMe(ByVal frmOwner As Form, ByVal bytState As Byte, ByVal lngID As Long)
'功能：窗体入口
'参数：
'  frmOwner：主窗体对象
'  bytState：窗体状态；0-新增；1-修改
'  lngID：窗体状态为0（新增）时，表示连接ID；窗体状态为1（修改）时，表示设备ID
    
    mlngID = lngID
    Me.Tag = bytState
    
    Call Init
    Call FullData(mlngID)
    Call cboLink_Click
    If Val(Me.Tag) = 0 Then Call cboDept_Click
    
    Me.Show vbModal, frmOwner
    
End Sub

Private Sub cboDept_Click()
    cmdSave.Enabled = cboDept.ListIndex >= 0 And cboLink.ListIndex >= 0
    If cboDept.ListIndex < 0 Then
        optObject(0).Value = False
        optObject(1).Value = False
        optObject(0).Enabled = False
        optObject(1).Enabled = False
    Else
        '药房服务对象
        Dim rsTmp As ADODB.Recordset
        
        On Error GoTo errHandle
        gstrSQL = "Select 服务对象 From 部门性质说明 " & _
                  "Where 部门id = [1] And 服务对象 in (1,2,3) " & _
                  "Order By 服务对象 "
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取部门服务对象", cboDept.ItemData(cboDept.ListIndex))
        Do While rsTmp.EOF = False
            Select Case gobjComLib.zlCommFun.Nvl(rsTmp!服务对象, 0)
                Case 1                  '门诊病人
                    optObject(0).Value = True
                    optObject(0).Enabled = True
                    optObject(1).Enabled = False
                Case 2                  '住院病人
                    optObject(1).Value = True
                    optObject(1).Enabled = True
                    optObject(0).Enabled = False
                Case 3                  '门诊病人与住院病人
                    optObject(0).Enabled = True
                    optObject(1).Enabled = True
                Case Else               '非病人
                    optObject(0).Value = False
                    optObject(1).Value = False
                    optObject(0).Enabled = False
                    optObject(1).Enabled = False
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        Set rsTmp = Nothing
        
    End If
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub cboDispense_Click()
    chkDispensing.Value = False
    chkDispensing.Enabled = cboDispense.ListIndex = 2
End Sub

Private Sub cboLink_Click()
    cmdSave.Enabled = cboDept.ListIndex >= 0 And cboLink.ListIndex >= 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer

    '检查
    If Trim(txtCode.Text) = "" Then
        MsgBox "未填写“编码”！", vbInformation, GSTR_INTERFACE_NAME
        txtCode.SetFocus
        Exit Sub
    End If
    If Trim(txtName.Text) = "" Then
        MsgBox "未填写“名称”！", vbInformation, GSTR_INTERFACE_NAME
        txtName.SetFocus
        Exit Sub
    End If
    If cboLink.ListIndex < 0 Then
        MsgBox "未选择“连接名”！", vbInformation, GSTR_INTERFACE_NAME
        cboLink.SetFocus
        Exit Sub
    End If
    If cboDept.ListIndex < 0 Then
        MsgBox "未选择“使用药房”！", vbInformation, GSTR_INTERFACE_NAME
        cboDept.SetFocus
        Exit Sub
    End If
    If optObject(0).Value = False And optObject(1) = False Then
        MsgBox "“服务对象”必须二选一！", vbInformation, GSTR_INTERFACE_NAME
        optObject(0).SetFocus
        Exit Sub
    End If
    If cboDispense.Enabled Then
        If cboDispense.ListIndex < 0 Then
            MsgBox "未选择“配药功能”对应的业务！", vbInformation, GSTR_INTERFACE_NAME
            cboDispense.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(Me.Tag) = 1 Then
        '修改
        gstrSQL = "Zl_药房注册设备_Update("
        gstrSQL = gstrSQL & mlngID & ","
        gstrSQL = gstrSQL & "'" & txtCode.Text & "',"
        gstrSQL = gstrSQL & "'" & txtName.Text & "',"
        gstrSQL = gstrSQL & IIf(Trim(txtModel.Text) = "", "null", "'" & txtModel.Text & "'") & ","
        gstrSQL = gstrSQL & IIf(Trim(txtManufacturer.Text) = "", "null", "'" & txtManufacturer.Text & "'") & ","
        gstrSQL = gstrSQL & cboLink.ItemData(cboLink.ListIndex) & ","
        gstrSQL = gstrSQL & cboDept.ItemData(cboDept.ListIndex) & ","
        gstrSQL = gstrSQL & IIf(optState(0).Value, "1", "null") & ","
        gstrSQL = gstrSQL & IIf(optObject(0).Value, "'1'", "'2'") & ","
        If cboDispense.Enabled Then
            gstrSQL = gstrSQL & "'" & cboDispense.ListIndex + 1 & "',"
        Else
            gstrSQL = gstrSQL & "null,"
        End If
        If chkDispensing.Enabled Then
            gstrSQL = gstrSQL & IIf(chkDispensing.Value, "'1'", "null")
        Else
            gstrSQL = gstrSQL & "null"
        End If
        gstrSQL = gstrSQL & ")"
        
        On Error GoTo errHandle
        Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "药房注册设备-修改")
        
    Else
        '新增
        gstrSQL = "Zl_药房注册设备_Insert("
        gstrSQL = gstrSQL & "'" & txtCode.Text & "',"
        gstrSQL = gstrSQL & "'" & txtName.Text & "',"
        gstrSQL = gstrSQL & IIf(Trim(txtModel.Text) = "", "null", "'" & txtModel.Text & "'") & ","
        gstrSQL = gstrSQL & IIf(Trim(txtManufacturer.Text) = "", "null", "'" & txtManufacturer.Text & "'") & ","
        gstrSQL = gstrSQL & cboLink.ItemData(cboLink.ListIndex) & ","
        gstrSQL = gstrSQL & cboDept.ItemData(cboDept.ListIndex) & ","
        gstrSQL = gstrSQL & IIf(optState(0).Value, "1", "null") & ","
        gstrSQL = gstrSQL & IIf(optObject(0).Value, "'1'", "'2'") & ","
        If cboDispense.Enabled Then
            gstrSQL = gstrSQL & "'" & cboDispense.ListIndex + 1 & "',"
        Else
            gstrSQL = gstrSQL & "null,"
        End If
        If chkDispensing.Enabled Then
            gstrSQL = gstrSQL & IIf(chkDispensing.Value, "'1'", "null")
        Else
            gstrSQL = gstrSQL & "null"
        End If
        gstrSQL = gstrSQL & ")"
        
        On Error GoTo errHandle
        Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "药房注册设备-新增")
        
    End If

    Unload Me
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub Form_Load()
    '
End Sub

Private Sub Init()
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select ID, 名称 From 药房设备连接 Order by 名称 "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房设备连接")
    Do While rsTmp.EOF = False
        cboLink.AddItem rsTmp!名称
        cboLink.ItemData(cboLink.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    gstrSQL = "Select Distinct a.Id, '【' || a.编码 || '】' || a.名称 名称 " & _
              "From 部门表 A, 部门性质说明 B " & _
              "Where a.Id = b.部门id And b.工作性质 In ('西药房', '成药房', '中药房') And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000/1/1', 'YYYY/MM/DD')) " & _
              "Order By '【' || a.编码 || '】' || a.名称 "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房部门信息")
    Do While rsTmp.EOF = False
        cboDept.AddItem rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub FullData(ByVal lngID As Long)
'功能：填充数据入控件
'参数：
'  lngID：设备ID

    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    With cboDispense
        .Clear
        .AddItem "门诊收费", 0
        .AddItem "处方发药-配药", 1
        .AddItem "处方发药-发药", 2
    End With
    
    '新增
    If Val(Me.Tag) = 0 Then
        gstrSQL = "Select 名称 From 药房设备连接 Where ID = [1] "
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房设备连接", lngID)
        If rsTmp.EOF = False Then
            cboLink.Text = rsTmp!名称
        End If
        optObject(0).Enabled = False
        optObject(1).Enabled = False
        cboDispense.Enabled = False
        chkDispensing.Enabled = False
        Exit Sub
    End If
    
    '药房设备信息
    gstrSQL = "Select a.*, b.名称 连接名, '【' || c.编码 || '】' || c.名称 药房 " & _
              "From 药房注册设备 A, 药房设备连接 B, 部门表 C " & _
              "Where a.连接id = b.Id and a.部门id = c.Id and a.ID = [1] "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房注册设备", lngID)
    If rsTmp.EOF = False Then
        cboLink.Text = rsTmp!连接名
        txtCode.Text = rsTmp!编码
        txtName.Text = rsTmp!名称
        txtModel.Text = gobjComLib.zlCommFun.Nvl(rsTmp!型号)
        txtManufacturer.Text = gobjComLib.zlCommFun.Nvl(rsTmp!制造商)
        If gobjComLib.zlCommFun.Nvl(rsTmp!启用, 0) = 1 Then
            optState(0).Value = True
        Else
            optState(1).Value = True
        End If
        cboLink.Text = rsTmp!连接名
        cboDept.Text = rsTmp!药房
    End If
    rsTmp.Close
    
    '药房设备参数
    gstrSQL = "Select b.Id, b.编码, b.名称, b.型号, b.启用, a.参数号, c.参数值 " & _
              "From Zlparameters A, 药房注册设备 B, 药房设备参数 C " & _
              "Where a.Id = c.参数id And b.Id = c.设备id And a.系统 = 100 And a.模块 = [1] And b.Id = [2] "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房设备参数", GINT_INTERFACE_MODULENO, lngID)
    
    '服务对象
    optObject(0).Enabled = True
    optObject(1).Enabled = True
    rsTmp.Filter = "参数号=1"
    If rsTmp.EOF = False Then
        Select Case Val(rsTmp!参数值)
            Case 1      '门诊
                optObject(0).Value = True
            Case 2      '住院
                optObject(1).Value = True
            Case Else   '异常
                optObject(0).Value = False
                optObject(1).Value = False
        End Select
    Else
        optObject(0).Value = False
        optObject(1).Value = False
    End If
    
    '配药对应业务
    cboDispense.ListIndex = -1
    cboDispense.Enabled = optObject(0).Value
    rsTmp.Filter = "参数号=2"
    If rsTmp.EOF = False Then
        If optObject(0).Value Then
            '门诊
            cboDispense.ListIndex = Val(rsTmp!参数值) - 1
        End If
    End If
    
    '发送对应业务
    chkDispensing.Value = False
    chkDispensing.Enabled = True
    rsTmp.Filter = "参数号=3"
    If rsTmp.EOF = False Then
        If cboDispense.ListIndex = 2 Then
            '处方发药-发药
            chkDispensing.Value = Val(rsTmp!参数值)
        Else
            chkDispensing.Enabled = False
        End If
    End If
    
    rsTmp.Close
    Set rsTmp = Nothing
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub optObject_Click(Index As Integer)
    cboDispense.Enabled = Index = 0
End Sub
