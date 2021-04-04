VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDeviceBase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设备基本信息"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   Icon            =   "frmDeviceBase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6645
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab sstDevice 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基础信息(&0)"
      TabPicture(0)   =   "frmDeviceBase.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDevice(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDevice(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDevice(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDevice(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDevice(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDevice(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCode"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboDept"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtManufacturer"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtModel"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "picObject"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "连接方式(&1)"
      TabPicture(1)   =   "frmDeviceBase.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtConnectStr"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdBuild"
      Tab(1).Control(2)=   "optLink(0)"
      Tab(1).Control(3)=   "optLink(1)"
      Tab(1).Control(4)=   "fraWS"
      Tab(1).Control(5)=   "optLink(2)"
      Tab(1).Control(6)=   "txtDirectory"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdBrowser"
      Tab(1).ControlCount=   8
      Begin VB.PictureBox picObject 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   2175
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2560
         Width           =   2175
         Begin VB.OptionButton optObject 
            Caption         =   "门诊"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optObject 
            Caption         =   "住院"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "浏览(&B)"
         Height          =   360
         Left            =   -70080
         TabIndex        =   31
         Top             =   3795
         Width           =   990
      End
      Begin VB.TextBox txtDirectory 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   -74640
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3840
         Width           =   4455
      End
      Begin VB.OptionButton optLink 
         Caption         =   "共享目录(&D)"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   29
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Frame fraWS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74640
         TabIndex        =   19
         Top             =   1560
         Width           =   5775
         Begin VB.TextBox txtConfirm 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   28
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   26
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtUser 
            Height          =   285
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   24
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton cmdWSTest 
            Caption         =   "测试(&T)"
            Height          =   360
            Left            =   4560
            TabIndex        =   22
            Top             =   202
            Width           =   990
         End
         Begin VB.TextBox txtURL 
            Height          =   285
            Left            =   1200
            TabIndex        =   21
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "确认密码："
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   27
            Top             =   1350
            Width           =   900
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "密    码："
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   990
            Width           =   900
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "用    户："
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   630
            Width           =   900
         End
         Begin VB.Label lblLink 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "服务地址："
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.OptionButton optLink 
         Caption         =   "Web Services(&W)"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   18
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optLink 
         Caption         =   "连接串(&L)"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   15
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuild 
         Caption         =   "创建(&U)"
         Height          =   360
         Left            =   -70080
         TabIndex        =   17
         Top             =   915
         Width           =   990
      End
      Begin VB.TextBox txtConnectStr 
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   -74640
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtModel 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtManufacturer 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   1800
         Width           =   3495
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "服务对象"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   2565
         Width           =   720
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名    称"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "型    号"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   1485
         Width           =   720
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "制 造 商"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   1845
         Width           =   720
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "使用药房"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   2205
         Width           =   720
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编    码"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   2
         Top             =   765
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5280
      TabIndex        =   33
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   32
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Frame fraLine1 
      Height          =   75
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   600
      Width           =   7335
   End
   Begin VB.Label lblComment 
      Caption         =   "设置发药设备的基本信息和数据交换连接方式！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDeviceBase.frx":0342
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmDeviceBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte    '0-新增,1-修改

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (LpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDlist Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Type BROWSEINFO
    hOwner As Long
    pidlroot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lparam As Long
    iImage As Long
End Type
Private mobjDataLink As MSDASC.DataLinks
Private mcnTmp As New ADODB.Connection

Public Sub ShowMe(ByVal frmOwner As Form, ByVal lng设备id As Long, ByVal bytType As Integer)
    mbytType = bytType
    
    '获取药房列表
    Call GetDrugStock
    
    '获取设备信息(修改状态时)
    Call GetTheDeviceInfo(lng设备id)
    
    Call cboDept_Click
    
    Me.Show vbModal, frmOwner
    
    Exit Sub
End Sub

Private Sub GetTheDeviceInfo(ByVal lng设备id As Long)
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    txtCode.Text = ""
    txtName.Text = ""
    txtModel.Text = ""
    txtManufacturer.Text = ""
    txtConnectStr.Text = ""
    txtURL.Text = ""
    txtUser.Text = ""
    txtPass.Text = ""
    txtConfirm.Text = ""
    txtDirectory.Text = ""
    cboDept.ListIndex = -1
    
    On Error GoTo errHandle
    
    gstrSQL = "Select a.Id, a.编码, a.名称, a.型号, a.制造商, a.使用部门id, '【' || b.编码 || '】' || b.名称 As 使用部门, " & _
        " Decode(a.连接类型, 1, '数据库', 2, 'WebService', 3, '共享目录', '未知') As 连接类型, a.连接内容, a.服务对象, a.是否启用 " & _
        " From 药房发药设备 A, 部门表 B " & _
        " Where a.使用部门id = b.Id and a.id=[1] "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetTheDeviceInfo", lng设备id)
    
    If rsData.RecordCount > 0 Then
        txtCode.Tag = rsData!ID
        txtCode.Text = rsData!编码
        txtName.Text = rsData!名称
        txtModel.Text = gobjComLib.zlcommfun.NVL(rsData!型号)
        txtManufacturer.Text = gobjComLib.zlcommfun.NVL(rsData!制造商)
        
        If gobjComLib.zlcommfun.NVL(rsData!服务对象, 0) = 1 Then
            optObject(0).Value = True
        Else
            optObject(1).Value = True
        End If
        
        For i = 0 To cboDept.ListCount - 1
            If cboDept.ItemData(i) = rsData!使用部门id Then
                cboDept.ListIndex = i
                Exit For
            End If
        Next
        
        If rsData!连接类型 = "数据库" Then
            optLink(0).Value = True
            txtConnectStr.Text = gobjComLib.zlcommfun.NVL(rsData!连接内容)
            Call optLink_Click(0)
        ElseIf rsData!连接类型 = "WebService" Then
            optLink(1).Value = True
            txtURL.Text = GetConnectStrEle(gobjComLib.zlcommfun.NVL(rsData!连接内容), enuLinkType.WEBServices, "URL")
            txtUser.Text = GetConnectStrEle(gobjComLib.zlcommfun.NVL(rsData!连接内容), enuLinkType.WEBServices, "USER")
            txtPass.Text = GetConnectStrEle(gobjComLib.zlcommfun.NVL(rsData!连接内容), enuLinkType.WEBServices, "PASS")
            txtConfirm.Text = txtPass.Text
            Call optLink_Click(1)
        ElseIf rsData!连接类型 = "共享目录" Then
            optLink(2).Value = True
            txtDirectory.Text = gobjComLib.zlcommfun.NVL(rsData!连接内容)
            Call optLink_Click(2)
        Else
            optLink(0).Value = True
            txtConnectStr.Text = gobjComLib.zlcommfun.NVL(rsData!连接内容)
            Call optLink_Click(0)
        End If
    
    Else
        optLink(0).Value = True
        Call optLink_Click(0)
    End If
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub cboDept_Click()
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
            Select Case gobjComLib.zlcommfun.NVL(rsTmp!服务对象, 0)
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

Private Sub cmdBrowser_Click()
    Dim strPath As String
    strPath = GetFolder(Me.hWnd, "浏览文件夹")
    If strPath <> "" Then
        txtDirectory.Text = strPath
    End If
End Sub

Private Function GetFolder(ByVal hWnd As Long, Optional Title As String) As String
    Dim typBI As BROWSEINFO
    Dim lngPID As Long
    Dim strFolder As String
    
    strFolder = Space(255)
    With typBI
       If IsNumeric(hWnd) Then .hOwner = hWnd
       .ulFlags = BIF_RETURNONLYFSDIRS
       .pidlroot = 0
       If Title <> "" Then
          .lpszTitle = Title & Chr$(0)
       Else
          .lpszTitle = "选择目录" & Chr$(0)
        End If
    End With

    lngPID = SHBrowseForFolder(typBI)
    
    If SHGetPathFromIDlist(ByVal lngPID, ByVal strFolder) Then
        GetFolder = Left(strFolder, InStr(strFolder, Chr$(0)) - 1)
    Else
        GetFolder = ""
    End If
End Function
Private Sub cmdBuild_Click()
    
    On Error GoTo errHandle
    If mobjDataLink Is Nothing Then
        Set mobjDataLink = New MSDASC.DataLinks
    End If
    If mcnTmp Is Nothing Then
        Set mcnTmp = mobjDataLink.PromptNew
    Else
        mcnTmp.ConnectionString = txtConnectStr.Text
        On Error Resume Next
        mobjDataLink.PromptEdit mcnTmp
        If Err <> 0 Then
            Err.Clear: On Error GoTo errHandle
            
            Set mobjDataLink = New MSDASC.DataLinks
            Set mcnTmp = mobjDataLink.PromptNew
        Else
            On Error GoTo errHandle
        End If
    End If
    
    If Not mcnTmp Is Nothing Then
        txtConnectStr.Text = mcnTmp.ConnectionString
    End If
    
    Exit Sub
    
errHandle:
    gstrMessage = Err.Description
    If gobjComLib.ErrCenter = 1 Then Resume
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
    If cboDept.ListIndex < 0 Then
        MsgBox "未选择“使用药房”！", vbInformation, GSTR_INTERFACE_NAME
        cboDept.SetFocus
        Exit Sub
    End If
    If optObject(0).Value = False And optObject(1).Value = False Then
        MsgBox "未选择“服务对象”！", vbInformation, GSTR_INTERFACE_NAME
        optObject(0).SetFocus
    End If
    
    If mbytType = 1 Then
        '修改
        gstrSQL = "Zl_药房发药设备_Update("
        gstrSQL = gstrSQL & Val(txtCode.Tag) & ","
        gstrSQL = gstrSQL & "'" & txtCode.Text & "',"
        gstrSQL = gstrSQL & "'" & txtName.Text & "',"
        gstrSQL = gstrSQL & IIf(Trim(txtModel.Text) = "", "null", "'" & txtModel.Text & "'") & ","
        gstrSQL = gstrSQL & IIf(Trim(txtManufacturer.Text) = "", "null", "'" & txtManufacturer.Text & "'") & ","
        gstrSQL = gstrSQL & cboDept.ItemData(cboDept.ListIndex) & ","
        gstrSQL = gstrSQL & IIf(optLink(0).Value, 1, IIf(optLink(1).Value, 2, 3)) & ","
        If optLink(1).Value Then
            gstrSQL = gstrSQL & "'" & GetURL() & "',"
        Else
            gstrSQL = gstrSQL & IIf(optLink(0).Value, "'" & txtConnectStr & "'", "'" & txtDirectory.Text & "'") & ","
        End If
        gstrSQL = gstrSQL & "1,"
        gstrSQL = gstrSQL & IIf(optObject(0).Value, "1", "2")
        gstrSQL = gstrSQL & ")"
        
        On Error GoTo errHandle
        Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "药房发药设备-修改")
        
    Else
        '新增
        gstrSQL = "Zl_药房发药设备_Insert("
        gstrSQL = gstrSQL & "'" & txtCode.Text & "',"
        gstrSQL = gstrSQL & "'" & txtName.Text & "',"
        gstrSQL = gstrSQL & IIf(Trim(txtModel.Text) = "", "null", "'" & txtModel.Text & "'") & ","
        gstrSQL = gstrSQL & IIf(Trim(txtManufacturer.Text) = "", "null", "'" & txtManufacturer.Text & "'") & ","
        gstrSQL = gstrSQL & cboDept.ItemData(cboDept.ListIndex) & ","
        gstrSQL = gstrSQL & IIf(optLink(0).Value, 1, IIf(optLink(1).Value, 2, 3)) & ","
        If optLink(1).Value Then
            gstrSQL = gstrSQL & "'" & GetURL() & "',"
        Else
            gstrSQL = gstrSQL & IIf(optLink(0).Value, "'" & txtConnectStr & "'", "'" & txtDirectory.Text & "'") & ","
        End If
        gstrSQL = gstrSQL & "1,"
        gstrSQL = gstrSQL & IIf(optObject(0).Value, "1", "2")
        gstrSQL = gstrSQL & ")"
        
        On Error GoTo errHandle
        Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "药房发药设备-新增")
        
    End If

    Unload Me
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub GetDrugStock()
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select Distinct a.Id, '【' || a.编码 || '】' || a.名称 名称 " & _
              "From 部门表 A, 部门性质说明 B " & _
              "Where a.Id = b.部门id And b.工作性质 In ('西药房', '成药房', '中药房') And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000/1/1', 'YYYY/MM/DD')) " & _
              "Order By '【' || a.编码 || '】' || a.名称 "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取药房部门信息")
    
    cboDept.Clear
    Do While rsData.EOF = False
        cboDept.AddItem rsData!名称
        cboDept.ItemData(cboDept.NewIndex) = rsData!ID
        rsData.MoveNext
    Loop
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub cmdWSTest_Click()
    If TestURL(txtURL.Text) = False Then
        MsgBox "服务地址连接测试失败！" & vbNewLine & gstrMessage, vbInformation, GSTR_INTERFACE_NAME
    Else
        gstrMessage = ""
        MsgBox "连接测试成功！", vbInformation, GSTR_INTERFACE_NAME
    End If
End Sub

Private Function GetURL() As String
    Dim strTmp As String
    
    If Trim(txtURL.Text) <> "" Then
        strTmp = "URL=" & Trim(txtURL.Text)
        If Trim(txtUser.Text) <> "" Then
            strTmp = strTmp & ";USER=" & Trim(txtUser.Text)
        End If
        If Trim(txtPass.Text) <> "" Then
            strTmp = strTmp & ";PASS=" & Trim(txtPass.Text)
        End If
    End If
    GetURL = strTmp
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mcnTmp = Nothing
    Set mobjDataLink = Nothing
End Sub

Private Sub optLink_Click(Index As Integer)
    cmdBuild.Enabled = Index = 0
    
    cmdWSTest.Enabled = Index = 1
    txtURL.Enabled = Index = 1
    txtUser.Enabled = Index = 1
    txtPass.Enabled = Index = 1
    txtConfirm.Enabled = Index = 1
    
    cmdBrowser.Enabled = Index = 2
    txtDirectory.Enabled = Index = 2
    
    Select Case Index
    Case 0
        If txtConnectStr.Text = "" Then
            cmdBuild.Caption = "创建(&U)"
        Else
            cmdBuild.Caption = "编辑(&U)"
        End If
    
        If cmdBuild.Visible Then cmdBuild.SetFocus
    Case 1
        If txtURL.Visible Then txtURL.SetFocus
    Case 2
        If txtDirectory.Visible Then txtDirectory.SetFocus
    End Select
    
End Sub

Private Sub txtConnectStr_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
