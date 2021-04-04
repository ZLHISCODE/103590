VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFilesSeverEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "服务器编辑"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   Icon            =   "frmFilesSeverEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgList 
      Left            =   45
      Top             =   3885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesSeverEdit.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesSeverEdit.frx":83A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   32
      Top             =   990
      Width           =   5835
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   5205
      TabIndex        =   28
      Top             =   0
      Width           =   5205
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "服务器状态：启用后才能上传并设置给客户端"
         Height          =   180
         Index           =   2
         Left            =   1365
         TabIndex        =   31
         Top             =   675
         Width           =   3600
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编号：唯一确定一个服务器的标识"
         Height          =   180
         Index           =   1
         Left            =   1365
         TabIndex        =   30
         Top             =   135
         Width           =   2700
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "默认服务器：只能有一个默认缺省服务器"
         Height          =   225
         Index           =   0
         Left            =   1365
         TabIndex        =   29
         Top             =   405
         Width           =   3780
      End
      Begin VB.Image imgCaption 
         Height          =   720
         Left            =   405
         Picture         =   "frmFilesSeverEdit.frx":9EF6
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.CheckBox chkDef 
      Caption         =   "默认服务器"
      Height          =   225
      Left            =   3615
      TabIndex        =   27
      Top             =   1275
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Height          =   415
      Left            =   1380
      TabIndex        =   26
      Top             =   2025
      Width           =   3400
      Begin VB.OptionButton optType 
         Caption         =   "共享"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   150
         Value           =   -1  'True
         Width           =   720
      End
      Begin VB.OptionButton optType 
         Caption         =   "FTP"
         Height          =   210
         Index           =   1
         Left            =   1380
         TabIndex        =   4
         Top             =   150
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   415
      Left            =   1380
      TabIndex        =   25
      Top             =   1575
      Width           =   3400
      Begin VB.OptionButton optSeverStatus 
         Caption         =   "启用"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   150
         Width           =   885
      End
      Begin VB.OptionButton optSeverStatus 
         Caption         =   "停用"
         Height          =   210
         Index           =   1
         Left            =   1380
         TabIndex        =   2
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   0
      Left            =   -345
      TabIndex        =   23
      Top             =   4545
      Width           =   5835
   End
   Begin VB.CommandButton cmdCel 
      Caption         =   "取消(&Q)"
      Height          =   350
      Left            =   3645
      TabIndex        =   11
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   2385
      TabIndex        =   10
      Top             =   4725
      Width           =   1100
   End
   Begin VB.PictureBox picPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1380
      ScaleHeight     =   270
      ScaleWidth      =   840
      TabIndex        =   21
      Top             =   3975
      Width           =   870
      Begin VB.TextBox txtPort 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   45
         TabIndex        =   9
         Text            =   "测试"
         Top             =   30
         Width           =   900
      End
   End
   Begin VB.PictureBox picPassWord 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1380
      ScaleHeight     =   270
      ScaleWidth      =   3345
      TabIndex        =   20
      Top             =   3495
      Width           =   3375
      Begin VB.TextBox txtPassWord 
         BorderStyle     =   0  'None
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   45
         TabIndex        =   8
         Text            =   "key"
         Top             =   15
         Width           =   3400
      End
   End
   Begin VB.PictureBox picUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1380
      ScaleHeight     =   270
      ScaleWidth      =   3345
      TabIndex        =   19
      Top             =   3030
      Width           =   3375
      Begin VB.TextBox txtUser 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   45
         TabIndex        =   7
         Text            =   "测试"
         Top             =   30
         Width           =   3400
      End
   End
   Begin VB.PictureBox picServerAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1380
      ScaleHeight     =   270
      ScaleWidth      =   3345
      TabIndex        =   18
      Top             =   2565
      Width           =   3375
      Begin VB.CommandButton cmdFileList 
         Caption         =   "···"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   2970
         TabIndex        =   6
         Top             =   -60
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.TextBox txtServerAddress 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   45
         TabIndex        =   5
         Text            =   "测试"
         Top             =   30
         Width           =   3030
      End
   End
   Begin VB.PictureBox picNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1380
      ScaleHeight     =   270
      ScaleWidth      =   645
      TabIndex        =   17
      Top             =   1245
      Width           =   670
      Begin VB.TextBox txtNumber 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   45
         TabIndex        =   0
         Text            =   "0"
         Top             =   30
         Width           =   700
      End
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器状态"
      Height          =   180
      Index           =   9
      Left            =   330
      TabIndex        =   24
      Top             =   1725
      Width           =   960
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "编号"
      Height          =   180
      Index           =   0
      Left            =   885
      TabIndex        =   22
      Top             =   1305
      Width           =   360
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "端口"
      Height          =   180
      Index           =   5
      Left            =   900
      TabIndex        =   16
      Top             =   4020
      Width           =   360
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户密码"
      Height          =   180
      Index           =   4
      Left            =   525
      TabIndex        =   15
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名称"
      Height          =   180
      Index           =   3
      Left            =   525
      TabIndex        =   14
      Top             =   3075
      Width           =   720
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "共享目录"
      Height          =   180
      Index           =   2
      Left            =   525
      TabIndex        =   13
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器类型"
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   12
      Top             =   2175
      Width           =   900
   End
End
Attribute VB_Name = "frmFilesSeverEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintReturn As Integer   '窗体返回值 0-不刷新数据 1-刷新数据
Private mintEditType As Integer '窗体编辑类型 0-新增 1-修改
Private mrsTemp As New ADODB.Recordset
Private mblnDefSever As Boolean '缺省服务器
Private mblnFirstAdd As Boolean '添加第一个服务器
Private mblnSatTag As Boolean '防止checkbox，opt控件重复触发
Private mstrServerAddress As String '备份服务器地址，用来判断重复添加相同服务器地址

Public Function ShowMe(intEditType As Integer, blnFirstAdd As Boolean, Optional intNumber As Integer)
'intEditType 0 - 新增 1 - 修改
Dim strSQL As String
On Error Resume Next
    
    Select Case intEditType
        Case 0 '新增初始化
            imgCaption.Picture = imgList.ListImages(1).Picture
            strSQL = "select max(编号) from ZLUpgradeServer"
            Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
            Me.Caption = "服务器-新增"
            OptType_Click 0
            txtServerAddress.Text = ""
            txtUser.Text = ""
            txtPassWord.Text = ""
            txtPort.Text = ""
            txtNumber.Text = Nvl(mrsTemp.Fields("MAX(编号)"), 0) + 1
            If blnFirstAdd = True Then chkDef.value = 1: mblnFirstAdd = blnFirstAdd
'            txtServerAddress.SetFocus
        Case 1 '修改初始化
            imgCaption.Picture = imgList.ListImages(2).Picture
            strSQL = "select * from ZLUpgradeServer where 编号 = " & intNumber
            Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
            Me.Caption = "服务器-修改"
            If Nvl(mrsTemp.Fields("是否升级"), "0") = "0" And Nvl(mrsTemp.Fields("是否缺省"), "0") = "0" And Nvl(mrsTemp.Fields("是否收集"), "0") = "0" Then
                optSeverStatus_Click 1
            Else
                optSeverStatus_Click 0
            End If
            If Nvl(mrsTemp.Fields("类型"), "") = "1" Then
                OptType_Click 1
            Else
                OptType_Click 0
            End If
            If Nvl(mrsTemp.Fields("是否升级"), "") = "1" Then
                optSeverStatus(0).value = True
            Else
                optSeverStatus(0).value = False
            End If
'            If Nvl(mrsTemp.Fields("是否升级"), "") = "1" Then
'                optSeverType_Click 0
'            ElseIf Nvl(mrsTemp.Fields("是否收集"), "") = "1" Then
'                optSeverType_Click 1
'            End If
'            Call IIf(Nvl(mrsTemp.Fields("类型"), "") = "1", optType_Click(1), optType_Click(0))
'            Call IIf(Nvl(mrsTemp.Fields("是否升级"), "") = "1", optSeverType_Click(0), optSeverType_Click(1))
            chkDef.value = IIf(Nvl(mrsTemp.Fields("是否缺省"), "") = "1", 1, 0)
            mblnDefSever = IIf(Nvl(mrsTemp.Fields("是否缺省"), "") = "1", True, False)
            txtNumber.Text = intNumber
            txtServerAddress.Text = Nvl(mrsTemp.Fields("位置"), "")
            mstrServerAddress = txtServerAddress.Text
            txtUser.Text = Nvl(mrsTemp.Fields("用户名"), "")
            txtPassWord.Text = Decipher(Nvl(mrsTemp.Fields("密码"), ""))
            txtPort.Text = Nvl(mrsTemp.Fields("端口"), "")
    End Select
    mintEditType = intEditType
    Me.Show 1, frmMDIMain
    
    ShowMe = mintReturn
    
End Function

Private Sub chkDef_Click()
    If mblnSatTag = True Then Exit Sub
    If chkDef.value = 1 Then optSeverStatus(0).value = True
    If mblnDefSever = True Then   '修改缺省服务器
        mblnSatTag = True
        chkDef.value = 1
        optSeverStatus(0).value = True
        Call MsgBox("该服务器为当前默认缺省服务器，不能取消缺省设置，不能停用该服务器，若有需要请切换其他服务器设置为缺省！", vbInformation, gstrSysName)
    End If
    If mblnFirstAdd = True And chkDef.value = 0 Then '首次添加服务器
        mblnSatTag = True
        chkDef.value = 1
        optSeverStatus(0).value = True
        Call MsgBox("首次添加升级服务器，需要启用服务器并设置为默认缺省服务器，不能取消！", vbInformation, gstrSysName)
    End If
    
    mblnSatTag = False
End Sub

Private Sub cmdCel_Click()
    Unload Me
    mintReturn = 0
End Sub

Private Sub cmdFileList_Click()
    Dim strFolderName As String
    On Error Resume Next
    
    strFolderName = OpenFolder(Me, "选择最新部件的所在目录")
    
    If Len(strFolderName) = 3 Then
        MsgBox "不能选择根目录(" & strFolderName & ")!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
'    txtServerAddress.Text = strFolderName
'    Me.txtAccessDir.Tag = Trim(strFolderName)

    If InStr(1, strFolderName, "\\") <> 0 Then
        txtServerAddress.Text = strFolderName
    Else
        txtServerAddress.Text = "\\" & GetMyCompterName & Mid(strFolderName, 3)
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim strType As String   '类型 0 - 共享 1 - FTP
    Dim strIsUpgrade As String '是否升级
    Dim strIsCheck As String    '是否缺省
    Dim strIsCollect As String  '是否收集
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '数据初始化
    strIsUpgrade = "0": strIsCollect = "0": strIsCheck = "0"
    strType = IIf(optType.Item(0).value = True, "0", "1")
    
    strIsUpgrade = IIf(optSeverStatus(0).value = True, "1", "0")
    strIsCheck = IIf(chkDef.value = 1, "1", "0")
    
    '输入检查
    If txtServerAddress = "" Then Call MsgBox("请输入" & IIf(strType = 0, "共享目录", "IP地址") & " !", vbInformation, gstrSysName): txtServerAddress.SetFocus: Exit Sub
    If Len(txtServerAddress) > 95 Then Call MsgBox("服务器路径过长，请修改！", vbInformation, gstrSysName): txtServerAddress.SetFocus: Exit Sub
    If txtUser.Text = "" Then Call MsgBox("请输入用户名!", vbInformation, gstrSysName): txtUser.SetFocus: Exit Sub
    If txtPassWord.Text = "" Then Call MsgBox("请输入密码!", vbInformation, gstrSysName): txtPassWord.SetFocus: Exit Sub
    If txtPort.Text = "" And strType = "1" Then Call MsgBox("请输入端口号!", vbInformation, gstrSysName): txtPort.SetFocus: Exit Sub
    If InStr(1, txtUser.Text, "'") <> 0 Then MsgBox "用户名中不能存在单引号!", vbInformation + vbDefaultButton1, gstrSysName: txtUser.SetFocus: Exit Sub
    If InStr(1, txtPassWord.Text, "'") <> 0 Then MsgBox "密码中不能存在单引号!", vbInformation + vbDefaultButton1, gstrSysName: txtPassWord.SetFocus: Exit Sub
    
    '重复服务器目录检查
    If txtServerAddress.Text <> mstrServerAddress Then
        strSQL = "select 1 from zlupgradeserver where 位置 = '" & txtServerAddress & "'"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        If rsTemp.EOF = False Then
            MsgBox IIf(strType = "1", "IP地址", "共享目录") & "：" & txtServerAddress & vbNewLine & "已存在，请不要重复添加！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    '连接校验 停用状态服务器不需要连接校验
    If optSeverStatus.Item(0).value = True Then
        If MsgBox("是否进行连接校验？", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            If strType = "1" Then
                If CheckFTPServer(txtServerAddress.Text, txtUser.Text, txtPassWord.Text, txtPort.Text) = False Then Exit Sub 'FTP服务器连接校验
            Else
                If CheckFileServer(txtServerAddress.Text, txtUser.Text, txtPassWord.Text) = False Then Exit Sub '共享服务器连接校验
            End If
        End If
    End If
    
    '数据库存入数据
    Select Case mintEditType
        Case 0 '新增
            strSQL = "Zl_Zlupgradeserver_Insert('" & txtNumber.Text & "','" & strType & "','" & txtServerAddress.Text & "','" & txtUser.Text & "','" & Cipher(txtPassWord.Text) & "','" & txtPort.Text & "','" & strIsUpgrade & "','" & strIsCheck & "','" & strIsCollect & "','" & "" & "')"
            Call ExecuteProcedure(strSQL, Me.Caption)
            If strIsCheck = "1" Then
                strSQL = "ZLReginfo_DefaultServer('" & strType & "','" & txtServerAddress.Text & "','" & txtUser.Text & "','" & txtPassWord.Text & "','" & txtPort.Text & "')"
                Call ExecuteProcedure(strSQL, Me.Caption)
            End If
            MsgBox "添加成功", vbInformation, gstrSysName
        Case 1 '修改
            strSQL = "Zl_Zlupgradeserver_Update('" & txtNumber.Text & "','" & strType & "','" & txtServerAddress.Text & "','" & txtUser.Text & "','" & Cipher(txtPassWord.Text) & "','" & txtPort.Text & "','" & strIsUpgrade & "','" & strIsCheck & "','" & strIsCollect & "','" & "" & "','" & 0 & "')"
            Call ExecuteProcedure(strSQL, Me.Caption)
            If strIsCheck = "1" Then
                strSQL = "ZLReginfo_DefaultServer('" & strType & "','" & txtServerAddress.Text & "','" & txtUser.Text & "','" & txtPassWord.Text & "','" & txtPort.Text & "')"
                Call ExecuteProcedure(strSQL, Me.Caption)
            End If
            MsgBox "修改成功", vbInformation, gstrSysName
    End Select
        
    mintReturn = 1
    Unload Me
    Exit Sub
errHand:
    MsgBox err.Description, vbInformation, gstrSysName
    If 1 = 0 Then
        Resume
    End If
End Sub

Private Sub optSeverStatus_Click(Index As Integer)
        Select Case Index
        Case 0

        Case 1
            If chkDef.value = 1 Then chkDef.value = 0
    End Select
End Sub

Private Sub OptType_Click(Index As Integer)
    Select Case Index
        Case 0
            optType.Item(0).value = True
            optType.Item(1).value = False
            lblItem(2).Caption = "共享目录"
            lblItem(2).Left = 570
            lblItem(5).Enabled = False
            txtPort.Text = ""
            txtPort.Enabled = False
            cmdFileList.Visible = True
            txtServerAddress.Width = 3030
        Case 1
            optType.Item(1).value = True
            optType.Item(0).value = False
            lblItem(2).Caption = "IP地址"
            lblItem(2).Left = 735
            lblItem(5).Enabled = True
            txtPort.Enabled = True
            cmdFileList.Visible = False
            txtServerAddress.Width = 3400
    End Select
End Sub

Private Function CheckFTPServer(ByVal strIp As String, ByVal strUser As String, ByVal strPass As String, ByVal strPort As String) As Boolean
    '-----------------------------------------------------------------------------
    '功能:检查当前的FTP服务器是否正确
    '返回:当前的文件服务器的各项正确,返回true,否则返回False
    '编制:陈振原
    '日期:2016/07/05
    'strIp - FTP地址
    'strUser - 用户名
    'strPass - 密码
    'strPort - 端口
    '-----------------------------------------------------------------------------
    On Error GoTo errHand:
    
    If strIp = "" Or strUser = "" Or strPass = "" Or strPort = "" Then
        CheckFTPServer = False
        Exit Function
    End If
    
    If IsFtpServer(Trim(strIp), Trim(strUser), Trim(strPass), Trim(strPort)) Then
        CheckFTPServer = True
    Else
        CheckFTPServer = False
        MsgBox "不能连接升级服务器，请检查FTP服务器配置!", vbInformation + vbDefaultButton1, gstrSysName
    End If
    Exit Function
    
errHand:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function


Private Function CheckFileServer(ByVal strAddress As String, ByVal strUser As String, ByVal strPass As String) As Boolean
    '-----------------------------------------------------------------------------
    '功能:检查当前的文件服务器是否正确
    '返回:当前的文件服务器的各项正确,返回true,否则返回False
    '编制:陈振原
    '日期:2016/07/05
    'strAddress - 地址
    'strUser - 用户
    'strPass - 密码
    '-----------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT

    On Error GoTo errHand:
    
    If strAddress = "" Or strUser = "" Or strPass = "" Then
        CheckFileServer = False
        Exit Function
    End If
    
    If IsNetServer(Trim(strAddress), Trim(strUser), Trim(strPass)) = False Then
        MsgBox "升级文件的指定目录不存在,请重新设置!", vbInformation + vbDefaultButton1, gstrSysName
        CheckFileServer = False
    Else
        CheckFileServer = True
    End If
    Call CancelNetServer(Trim(strAddress))
    
    Exit Function
errHand:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function FindFile(ByVal strFileName As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--功能:查找指定的文件或文夹是否存在
    '--返回: 如果存在此文件为True,否则为Flase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

Private Sub txtPassWord_GotFocus()
    txtPassWord.SelStart = 0
    txtPassWord.SelLength = Len(txtPassWord)
End Sub

Private Sub txtPort_GotFocus()
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort)
End Sub

Private Sub txtServerAddress_GotFocus()
    txtServerAddress.SelStart = 0
    txtServerAddress.SelLength = Len(txtServerAddress)
End Sub

Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser)
End Sub
