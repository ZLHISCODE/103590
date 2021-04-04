VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmServiceSet 
   Caption         =   "服务设置"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   300
      Left            =   5520
      TabIndex        =   34
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   300
      Left            =   2160
      TabIndex        =   33
      Top             =   6000
      Width           =   1100
   End
   Begin TabDlg.SSTab tabService 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "设备设置"
      TabPicture(0)   =   "frmServiceSet.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(3)=   "MSFDevice"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "txtDeviceName"
      Tab(0).Control(6)=   "cmdAddDevice"
      Tab(0).Control(7)=   "cmdModiDevice"
      Tab(0).Control(8)=   "cmdDelDevice"
      Tab(0).Control(9)=   "txtIPAddr"
      Tab(0).Control(10)=   "cboModality"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "服务设置"
      TabPicture(1)   =   "frmServiceSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAddService"
      Tab(1).Control(1)=   "MSFService"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "cmdModiService"
      Tab(1).Control(4)=   "cmdDelService"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "DICOM服务设置"
      TabPicture(2)   =   "frmServiceSet.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "MSFDicomDevice"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "MSFDicomService"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "DICOM服务参数"
         Height          =   2415
         Left            =   120
         TabIndex        =   56
         Top             =   3240
         Width           =   8775
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDicomService 
         Height          =   1335
         Left            =   120
         TabIndex        =   55
         Top             =   1800
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDicomDevice 
         Height          =   1215
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin VB.ComboBox cboModality 
         Height          =   300
         ItemData        =   "frmServiceSet.frx":0054
         Left            =   -71040
         List            =   "frmServiceSet.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton cmdDelService 
         Caption         =   "删除"
         Height          =   300
         Left            =   -68160
         TabIndex        =   51
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdModiService 
         Caption         =   "修改"
         Height          =   300
         Left            =   -71040
         TabIndex        =   50
         Top             =   5280
         Width           =   1100
      End
      Begin VB.Frame Frame2 
         Caption         =   "服务设置"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   38
         Top             =   3600
         Width           =   8775
         Begin VB.ComboBox cboServiceType 
            Height          =   300
            ItemData        =   "frmServiceSet.frx":006E
            Left            =   3720
            List            =   "frmServiceSet.frx":007B
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdServiceSetup 
            Caption         =   "高级设置"
            Enabled         =   0   'False
            Height          =   300
            Left            =   7440
            TabIndex        =   43
            Top             =   360
            Width           =   1100
         End
         Begin VB.TextBox txtServiceName 
            Height          =   300
            Left            =   720
            MaxLength       =   20
            TabIndex        =   42
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtServicePort 
            Height          =   300
            Left            =   6840
            MaxLength       =   4
            TabIndex        =   41
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtServiceIP 
            Height          =   300
            Left            =   720
            MaxLength       =   15
            TabIndex        =   40
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtServiceAE 
            Height          =   300
            Left            =   3720
            MaxLength       =   20
            TabIndex        =   39
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "端口"
            Height          =   165
            Left            =   6240
            TabIndex        =   48
            Top             =   908
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "服务名"
            Height          =   165
            Left            =   120
            TabIndex        =   47
            Top             =   435
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "服务类型"
            Height          =   165
            Left            =   3000
            TabIndex        =   46
            Top             =   435
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "AE名称"
            Height          =   165
            Left            =   3000
            TabIndex        =   45
            Top             =   908
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "IP地址"
            Height          =   165
            Left            =   120
            TabIndex        =   44
            Top             =   908
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFService 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5318
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.TextBox txtIPAddr 
         Height          =   300
         Left            =   -67920
         MaxLength       =   15
         TabIndex        =   31
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelDevice 
         Caption         =   "删除"
         Height          =   300
         Left            =   -68160
         TabIndex        =   23
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdModiDevice 
         Caption         =   "修改"
         Height          =   300
         Left            =   -71040
         TabIndex        =   22
         Top             =   5280
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddDevice 
         Caption         =   "增加"
         Height          =   300
         Left            =   -73920
         TabIndex        =   21
         Top             =   5280
         Width           =   1100
      End
      Begin VB.TextBox txtDeviceName 
         Height          =   300
         Left            =   -73920
         MaxLength       =   100
         TabIndex        =   20
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "设备服务设置"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   4
         Top             =   3240
         Width           =   8775
         Begin VB.CheckBox chkService 
            Caption         =   "Q/R检索服务"
            Height          =   255
            Index           =   2
            Left            =   6000
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkService 
            Caption         =   "Worklist服务"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   36
            Top             =   240
            Width           =   1455
         End
         Begin VB.Frame frmService 
            Enabled         =   0   'False
            Height          =   1455
            Index           =   1
            Left            =   3000
            TabIndex        =   24
            Top             =   240
            Width           =   2800
            Begin VB.TextBox txtDeviceAE 
               Height          =   300
               Index           =   1
               Left            =   840
               MaxLength       =   20
               TabIndex        =   27
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtDevicePort 
               Height          =   300
               Index           =   1
               Left            =   840
               MaxLength       =   4
               TabIndex        =   26
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox cboService 
               Height          =   300
               Index           =   1
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label Label6 
               Caption         =   "AE名称"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   390
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "端口"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   750
               Width           =   615
            End
            Begin VB.Label Label8 
               Caption         =   "服务"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   1110
               Width           =   615
            End
         End
         Begin VB.CheckBox chkService 
            Caption         =   "图像存储服务"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.Frame frmService 
            Enabled         =   0   'False
            Height          =   1455
            Index           =   2
            Left            =   5880
            TabIndex        =   6
            Top             =   240
            Width           =   2800
            Begin VB.ComboBox cboService 
               Height          =   300
               Index           =   2
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   1080
               Width           =   1695
            End
            Begin VB.TextBox txtDevicePort 
               Height          =   300
               Index           =   2
               Left            =   840
               MaxLength       =   4
               TabIndex        =   18
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtDeviceAE 
               Height          =   300
               Index           =   2
               Left            =   840
               MaxLength       =   20
               TabIndex        =   17
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label11 
               Caption         =   "服务"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   1110
               Width           =   615
            End
            Begin VB.Label Label10 
               Caption         =   "端口"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   750
               Width           =   615
            End
            Begin VB.Label Label9 
               Caption         =   "AE名称"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   390
               Width           =   615
            End
         End
         Begin VB.Frame frmService 
            Enabled         =   0   'False
            Height          =   1455
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2800
            Begin VB.ComboBox cboService 
               Height          =   300
               Index           =   0
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   1080
               Width           =   1695
            End
            Begin VB.TextBox txtDevicePort 
               Height          =   300
               Index           =   0
               Left            =   840
               MaxLength       =   4
               TabIndex        =   11
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtDeviceAE 
               Height          =   300
               Index           =   0
               Left            =   840
               MaxLength       =   20
               TabIndex        =   10
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label5 
               Caption         =   "服务"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   1110
               Width           =   615
            End
            Begin VB.Label Label4 
               Caption         =   "端口"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   750
               Width           =   615
            End
            Begin VB.Label Label3 
               Caption         =   "AE名称"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   390
               Width           =   615
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDevice 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3836
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.CommandButton cmdAddService 
         Caption         =   "增加"
         Height          =   300
         Left            =   -73920
         TabIndex        =   49
         Top             =   5280
         Width           =   1100
      End
      Begin VB.Label Label17 
         Caption         =   "影像类别"
         Height          =   255
         Left            =   -71880
         TabIndex        =   35
         Top             =   2903
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "设备IP地址"
         Height          =   255
         Left            =   -68880
         TabIndex        =   3
         Top             =   2903
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "设备主机名"
         Height          =   255
         Left            =   -74880
         TabIndex        =   2
         Top             =   2903
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmServiceSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngServiceID As Long    '服务ID

Private Sub chkService_Click(Index As Integer)
    If Me.chkService(Index).value = 1 Then
        Me.frmService(Index).Enabled = True
    Else
        Me.frmService(Index).Enabled = False
        Me.txtDeviceAE(Index).Text = ""
        Me.txtDevicePort(Index).Text = ""
        Me.cboService(Index).ListIndex = -1
    End If
End Sub

Private Sub cmdAddDevice_Click()
    Dim i As Integer
    Dim blnResult As Boolean    '有效性检查结果
    '检查输入是否有效
    blnResult = funValidateDevice
    If blnResult = False Then Exit Sub
    
    On Error GoTo errHand
    '插入数据
    gstrSQL = "Zl_影像接入设备_INSERT('" & Me.txtIPAddr.Text & "','" & Me.txtDeviceName.Text & "','" & _
                    Left(cboModality.Text, InStr(cboModality.Text, "-") - 1) & "'"
    For i = 0 To 2
        If chkService(i).value = 1 Then
            gstrSQL = gstrSQL & ",'" & Me.txtDeviceAE(i).Text & "','" & Me.txtDevicePort(i).Text & "'," & Me.cboService(i).ItemData(Me.cboService(i).ListIndex)
        Else
            gstrSQL = gstrSQL & ",null,null,null"
        End If
    Next i
    gstrSQL = gstrSQL & ")"
                        
    ExecuteProcedure "增加DICOM设备"
    '刷新列表
    Call subFillMSFDevice(Me.MSFDevice.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdAddService_Click()
    Dim i As Integer
    Dim blnResult As Boolean    '有效性检查结果
    '检查输入是否有效
    blnResult = funValidateService
    If blnResult = False Then Exit Sub
    On Error GoTo errHand
    '插入数据
    gstrSQL = "Zl_影像DICOM服务_INSERT('" & Me.txtServiceName.Text & "','" & Me.txtServiceIP.Text & "','" & _
                    Me.txtServiceAE.Text & "','" & Me.txtServicePort.Text & "','" & Me.cboServiceType.Text & "')"
                        
    ExecuteProcedure "增加DICOM服务"
    '刷新列表
    Call subFillMSFService(Me.MSFService.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function funValidateService()
    Dim arrIP() As String
    Dim i As Integer
    
    '判断输入的设备信息是否有效
    '服务名不为空
    If Me.txtServiceName = "" Then GoTo inValidate
    '服务类型不为空
    If Me.cboServiceType.ListIndex = -1 Then GoTo inValidate
    'IP地址有效
    If Me.txtServiceIP.Text = "" Then
        GoTo inValidate
    Else
        arrIP = Split(Me.txtServiceIP.Text, ".")
        If UBound(arrIP) = 3 Then
            If IsNumeric(arrIP(0)) And IsNumeric(arrIP(1)) And IsNumeric(arrIP(2)) And IsNumeric(arrIP(3)) Then
                '有效，不处理
            Else
                GoTo inValidate
            End If
        Else
            GoTo inValidate
        End If
    End If
    'AE名称不为空
    If Me.txtServiceAE.Text = "" Then GoTo inValidate
    '端口有效
    If Me.txtServicePort.Text = "" Then GoTo inValidate
    
    funValidateService = True
    Exit Function
inValidate:
    MsgBox "输入数据有误或者不完整，请检查后重新输入", vbOKOnly, "输入数据有误”"
    Exit Function
End Function

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelDevice_Click()
    If MSFDevice.Rows <= 1 Then Exit Sub
    On Error GoTo errHand
    '删除数据
    gstrSQL = "Zl_影像接入设备_DELETE(" & Me.MSFDevice.TextMatrix(Me.MSFDevice.RowSel, 12) & ")"
                        
    ExecuteProcedure "修改DICOM设备"
    '刷新列表
    Call subFillMSFDevice
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdDelService_Click()
    If MSFService.Rows <= 1 Then Exit Sub
    On Error GoTo errHand
    '删除数据
    gstrSQL = "Zl_影像DICOM服务_DELETE(" & Me.MSFService.TextMatrix(Me.MSFService.RowSel, 0) & ")"
                        
    ExecuteProcedure "删除加DICOM服务"
    '刷新列表
    Call subFillMSFService
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdModiDevice_Click()
    Dim i As Integer
    Dim blnResult As Boolean    '有效性检查结果
    '检查输入是否有效
    blnResult = funValidateDevice
    If blnResult = False Then Exit Sub
    
    On Error GoTo errHand
    '更新数据
    gstrSQL = "Zl_影像接入设备_UPDATE(" & Me.MSFDevice.TextMatrix(Me.MSFDevice.RowSel, 12) & ",'" & Me.txtIPAddr.Text & "','" & Me.txtDeviceName.Text & "','" & _
                    Left(cboModality.Text, InStr(cboModality.Text, "-") - 1) & "'"
    For i = 0 To 2
        If chkService(i).value = 1 Then
            gstrSQL = gstrSQL & ",'" & Me.txtDeviceAE(i).Text & "','" & Me.txtDevicePort(i).Text & "'," & Me.cboService(i).ItemData(Me.cboService(i).ListIndex)
        Else
            gstrSQL = gstrSQL & ",null,null,null"
        End If
    Next i
    gstrSQL = gstrSQL & ")"
                        
    ExecuteProcedure "修改DICOM设备"
    '刷新列表
    Call subFillMSFDevice(Me.MSFDevice.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdModiService_Click()
    Dim i As Integer
    Dim blnResult As Boolean    '有效性检查结果
    '检查输入是否有效
    blnResult = funValidateService
    If blnResult = False Then Exit Sub
    On Error GoTo errHand
    '修改数据
    gstrSQL = "Zl_影像DICOM服务_UPDATE(" & Me.MSFService.TextMatrix(Me.MSFService.RowSel, 0) & ",'" & Me.txtServiceName.Text & "','" & Me.txtServiceIP.Text & "','" & _
                    Me.txtServiceAE.Text & "','" & Me.txtServicePort.Text & "','" & Me.cboServiceType.Text & "')"
                        
    ExecuteProcedure "修改DICOM服务"
    '刷新列表
    Call subFillMSFService(Me.MSFService.RowSel)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CmdOK_Click()
    Unload Me
End Sub

Private Sub cmdServiceSetup_Click()
    If Me.MSFService.RowSel <> 0 Then
        frmAdvancedSet.ShowMe Me, Me.cboServiceType.Text, lngServiceID
    End If
End Sub

Private Sub Form_Load()
    '填充影像类别列表
    Call subFillcboModality
    '填充可选服务列表
    Call subFillcboService
    '填充设备列表
    Call subFillMSFDevice
    '填充服务列表
    Call subFillMSFService
End Sub

Private Sub subFillcboModality()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    '从数据库中读取设备
    strSQL = "Select 编码,名称 From 影像检查类别 "
    Set rsTmp = OpenSQLRecord(strSQL, "读取影像类别")
    
    cboModality.Clear
    While Not rsTmp.EOF
        cboModality.AddItem rsTmp!编码 & "-" & (rsTmp!名称)
        rsTmp.MoveNext
    Wend
End Sub

Private Sub subFillMSFService(Optional iRow As Integer = 1)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intRowPos As Integer
    
    '从数据库中读取服务
    strSQL = "Select a.服务ID,a.服务名,a.服务IP,a.服务AE,a.服务端口,a.服务功能  From 影像DICOM服务 a "
    Set rsTmp = OpenSQLRecord(strSQL, "读取DICOM服务")
    
    With MSFService
        .Clear
        .Rows = 1
        .Cols = 6
        .ColWidth(0) = 800
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColWidth(5) = 800
        
        .FixedCols = 0
        .TextMatrix(0, 0) = "服务ID"
        .TextMatrix(0, 1) = "服务名"
        .TextMatrix(0, 2) = "服务功能"
        .TextMatrix(0, 3) = "服务IP"
        .TextMatrix(0, 4) = "服务AE"
        .TextMatrix(0, 5) = "服务端口"
        
        intRowPos = 1
        
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(intRowPos, 0) = Nvl(rsTmp!服务ID)
            .TextMatrix(intRowPos, 1) = Nvl(rsTmp!服务名)
            .TextMatrix(intRowPos, 2) = Nvl(rsTmp!服务功能)
            .TextMatrix(intRowPos, 3) = Nvl(rsTmp!服务IP)
            .TextMatrix(intRowPos, 4) = Nvl(rsTmp!服务AE)
            .TextMatrix(intRowPos, 5) = Nvl(rsTmp!服务端口)
            
            rsTmp.MoveNext
            intRowPos = .Rows
        Wend
    End With
    subClickMSFService iRow
End Sub

Private Function funValidateDevice() As Boolean
    Dim arrIP() As String
    Dim i As Integer
    
    '判断输入的设备信息是否有效
    '设备名不为空
    If Me.txtDeviceName = "" Then
        GoTo inValidate
    End If
    '设备IP地址有效
    If Me.txtIPAddr = "" Then
        GoTo inValidate
    Else
        arrIP = Split(Me.txtIPAddr.Text, ".")
        If UBound(arrIP) = 3 Then
            If IsNumeric(arrIP(0)) And IsNumeric(arrIP(1)) And IsNumeric(arrIP(2)) And IsNumeric(arrIP(3)) Then
                '有效，不处理
            Else
                GoTo inValidate
            End If
        Else
            GoTo inValidate
        End If
    End If
    '影像类别不为空
    If Me.cboModality.Text = "" Then
        GoTo inValidate
    End If
    '服务启动后，AE，端口，PACS服务
    For i = 0 To 2
        If chkService(i).value = 1 Then
            '只检查有用的数据
            If Me.txtDeviceAE(i).Text = "" Then GoTo inValidate
            If Me.txtDevicePort(i).Text = "" Then GoTo inValidate
            If Me.cboService(i).ListIndex = -1 Then GoTo inValidate
        End If
    Next i
    
    funValidateDevice = True
    Exit Function
inValidate:
    MsgBox "输入数据有误或者不完整，请检查后重新输入", vbOKOnly, "输入数据有误”"
    Exit Function
End Function

Private Sub subFillcboService()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    '从数据库中读取设备
    strSQL = "Select 服务ID,服务名,服务功能 From 影像DICOM服务 "
    Set rsTmp = OpenSQLRecord(strSQL, "读取可选DICOM服务")
    
    For i = 0 To 2
        cboService(i).Clear
    Next i
    
    While Not rsTmp.EOF
        Select Case rsTmp!服务功能
        Case ZLPACS_存储服务
            i = 0
        Case ZLPACS_Worklist服务
            i = 1
        Case ZLPACS_检索服务
            i = 2
        End Select

        cboService(i).AddItem (rsTmp!服务名)
        cboService(i).ItemData(cboService(i).ListCount - 1) = rsTmp!服务ID
        rsTmp.MoveNext
    Wend
    
End Sub

Private Sub subFillMSFDevice(Optional iRow As Integer = 1)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intRowPos As Integer
    
    '从数据库中读取设备
    strSQL = "Select a.接入id, a.ip地址, a.设备名称,a.影像类别,a.存储AE,a.存储端口,a.存储服务ID, b1.服务名 As 存储服务名, " & _
             "a.WORKLISTAE,a.WORKLIST端口,a.WORKLIST服务ID,b2.服务名 As WORKLIST服务名,a.检索AE,a.检索端口,a.检索服务ID, " & _
             "b3.服务名 As 检索服务名 From 影像接入设备 a ,影像DICOM服务 b1,影像DICOM服务 b2,影像DICOM服务 b3 Where " & _
             "a.存储服务id=b1.服务id(+) And a.worklist服务id=b2.服务id(+) And a.检索服务id = b3.服务id(+) "
    Set rsTmp = OpenSQLRecord(strSQL, "读取DICOM设备")
    
    With MSFDevice
        .Clear
        .Rows = 1
        .Cols = 13
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 1400
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 800
        .ColWidth(9) = 800
        .ColWidth(10) = 800
        .ColWidth(11) = 800
        .ColWidth(12) = 800
        
        .FixedCols = 0
        '.FixedRows = 1
        .TextMatrix(0, 0) = "设备名称"
        .TextMatrix(0, 1) = "影像类别"
        .TextMatrix(0, 2) = "IP地址"
        .TextMatrix(0, 3) = "存储AE"
        .TextMatrix(0, 4) = "存储端口"
        .TextMatrix(0, 5) = "存储服务名"
        .TextMatrix(0, 6) = "WorklistAE"
        .TextMatrix(0, 7) = "Worklist端口"
        .TextMatrix(0, 8) = "Worklist服务名"
        .TextMatrix(0, 9) = "检索AE"
        .TextMatrix(0, 10) = "检索端口"
        .TextMatrix(0, 11) = "检索服务名"
        .TextMatrix(0, 12) = "接入ID"
        intRowPos = 1
        
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(intRowPos, 0) = Nvl(rsTmp!设备名称)
            .TextMatrix(intRowPos, 1) = Nvl(rsTmp!影像类别)
            .TextMatrix(intRowPos, 2) = Nvl(rsTmp!IP地址)
            .TextMatrix(intRowPos, 3) = Nvl(rsTmp!存储AE)
            .TextMatrix(intRowPos, 4) = Nvl(rsTmp!存储端口)
            .TextMatrix(intRowPos, 5) = Nvl(rsTmp!存储服务名)
            .TextMatrix(intRowPos, 6) = Nvl(rsTmp!WORKLISTAE)
            .TextMatrix(intRowPos, 7) = Nvl(rsTmp!WORKLIST端口)
            .TextMatrix(intRowPos, 8) = Nvl(rsTmp!WORKLIST服务名)
            .TextMatrix(intRowPos, 9) = Nvl(rsTmp!检索AE)
            .TextMatrix(intRowPos, 10) = Nvl(rsTmp!检索端口)
            .TextMatrix(intRowPos, 11) = Nvl(rsTmp!检索服务名)
            .TextMatrix(intRowPos, 12) = Nvl(rsTmp!接入id)
            rsTmp.MoveNext
            intRowPos = .Rows
        Wend
    End With
    
    Call subClickMSFDevice(iRow)
End Sub

Private Sub MSFDevice_Click()
    Dim iSelected As Integer
    If MSFDevice.Rows <= 1 Then Exit Sub
    
    With MSFDevice
        iSelected = .RowSel
        '填写基本信息
        Me.txtDeviceName.Text = .TextMatrix(iSelected, 0)
        Me.cboModality.Text = funcGetModalityText(.TextMatrix(iSelected, 1))
        Me.txtIPAddr.Text = .TextMatrix(iSelected, 2)
        '填写存储服务
        If .TextMatrix(iSelected, 5) = "" Then
            Me.chkService(0).value = 0
        Else
            Me.chkService(0).value = 1
            Me.txtDeviceAE(0).Text = .TextMatrix(iSelected, 3)
            Me.txtDevicePort(0).Text = .TextMatrix(iSelected, 4)
            Me.cboService(0).Text = .TextMatrix(iSelected, 5)
        End If
        '填写WORKLIST服务
        If .TextMatrix(iSelected, 8) = "" Then
            Me.chkService(1).value = 0
        Else
            Me.chkService(1).value = 1
            Me.txtDeviceAE(1).Text = .TextMatrix(iSelected, 6)
            Me.txtDevicePort(1).Text = .TextMatrix(iSelected, 7)
            Me.cboService(1).Text = .TextMatrix(iSelected, 8)
        End If
        '填写检索服务
        If .TextMatrix(iSelected, 11) = "" Then
            Me.chkService(2).value = 0
        Else
            Me.chkService(2).value = 1
            Me.txtDeviceAE(2).Text = .TextMatrix(iSelected, 9)
            Me.txtDevicePort(2).Text = .TextMatrix(iSelected, 10)
            Me.cboService(2).Text = .TextMatrix(iSelected, 11)
        End If
    End With
End Sub

Private Function funcGetModalityText(strModality As String) As String
    Dim i As Integer
    For i = 0 To cboModality.ListCount - 1
        If Left(cboModality.list(i), InStr(cboModality.list(i), "-") - 1) = strModality Then
            funcGetModalityText = cboModality.list(i)
            Exit Function
        End If
    Next i
End Function

Private Sub MSFService_Click()
    Dim iSelected As Integer
    If MSFService.Rows <= 1 Then Exit Sub
    With MSFService
        iSelected = .RowSel
        lngServiceID = .TextMatrix(iSelected, 0)
        '填写服务名
        Me.txtServiceName.Text = .TextMatrix(iSelected, 1)
        '填写服务类型
        Me.cboServiceType.Text = .TextMatrix(iSelected, 2)
        '服务IP地址
        Me.txtServiceIP.Text = .TextMatrix(iSelected, 3)
        '服务AE
        Me.txtServiceAE.Text = .TextMatrix(iSelected, 4)
        '服务端口
        Me.txtServicePort.Text = .TextMatrix(iSelected, 5)
        If Me.cboServiceType.Text = ZLPACS_检索服务 Then
            cmdServiceSetup.Enabled = False
        Else
            cmdServiceSetup.Enabled = True
        End If
    End With
End Sub

Private Sub subClickMSFService(Optional iRow As Integer = 1)

    If iRow > Me.MSFService.Rows Or iRow < 1 Then iRow = 1

    If Me.MSFService.Rows > 1 Then
        Me.MSFService.Row = iRow - 1
        Me.MSFService.RowSel = iRow
        Call MSFService_Click
    End If
End Sub

Private Sub subClickMSFDevice(Optional iRow As Integer = 1)

    If iRow > Me.MSFDevice.Rows Or iRow < 1 Then iRow = 1

    If Me.MSFDevice.Rows > 1 Then
        Me.MSFDevice.Row = iRow - 1
        Me.MSFDevice.RowSel = iRow
        Call MSFDevice_Click
    End If
End Sub

Private Sub tabService_Click(PreviousTab As Integer)
    '如果打开设备页，则刷新设备中的可选服务列表
    If tabService.Tab = 0 Then
        '填充可选服务列表
        Call subFillcboService
        '填充设备列表
        Call subFillMSFDevice
    ElseIf tabService.Tab = 2 Then
        Call subFillMSFDicomDevice
       ' Call subFillMSFDicomService
    
    End If
End Sub

Private Sub subFillMSFDicomDevice()
    '填充DICOM设备
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intRowPos As Integer
    
    '从数据库中读取设备
    strSQL = "Select  设备号,设备名,类型,IP地址 From 影像设备目录 Where 类型 =4 "
    Set rsTmp = OpenSQLRecord(strSQL, "读取DICOM设备")
    
    With MSFDicomDevice
        .Clear
        .Rows = 1
        .Cols = 4
        .ColWidth(0) = 1400
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        
        .FixedCols = 0
        .TextMatrix(0, 0) = "设备号"
        .TextMatrix(0, 1) = "设备名称"
        .TextMatrix(0, 2) = "类型"
        .TextMatrix(0, 3) = "IP地址"
        intRowPos = 1
        
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(intRowPos, 0) = Nvl(rsTmp!设备号)
            .TextMatrix(intRowPos, 1) = Nvl(rsTmp!设备名)
            .TextMatrix(intRowPos, 2) = "DICOM影像设备"
            .TextMatrix(intRowPos, 3) = Nvl(rsTmp!IP地址)
            rsTmp.MoveNext
            intRowPos = .Rows
        Wend
    End With
    
   ' Call subClickMSFDevice(iRow)
End Sub

Private Sub txtDevicePort_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtServicePort_KeyPress(KeyAscii As Integer)
    If (Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
