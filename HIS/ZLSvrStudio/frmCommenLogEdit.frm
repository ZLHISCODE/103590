VERSION 5.00
Begin VB.Form frmCommenLogServer 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "日志服务器编辑"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   Icon            =   "frmCommenLogServer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLogUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "写日志用户"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   20
      Top             =   4200
      Width           =   5175
      Begin VB.TextBox Text4 
         Height          =   350
         Left            =   960
         TabIndex        =   22
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   960
         TabIndex        =   21
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "密码"
         Height          =   180
         Left            =   480
         TabIndex        =   24
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "用户"
         Height          =   180
         Left            =   480
         TabIndex        =   23
         Top             =   330
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "日志服务器管理用户"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      TabIndex        =   15
      Top             =   2880
      Width           =   5175
      Begin VB.TextBox Text3 
         Height          =   350
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Left            =   960
         TabIndex        =   16
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "用户"
         Height          =   180
         Left            =   480
         TabIndex        =   19
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "密码"
         Height          =   180
         Left            =   480
         TabIndex        =   18
         Top             =   810
         Width           =   360
      End
   End
   Begin VB.Frame fraServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "日志服务器"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtSID 
         Height          =   350
         Left            =   960
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtPort 
         Height          =   350
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtIP 
         Height          =   350
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblSID 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "SID"
         Height          =   180
         Left            =   480
         TabIndex        =   13
         Top             =   1290
         Width           =   270
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Port"
         Height          =   180
         Left            =   480
         TabIndex        =   11
         Top             =   810
         Width           =   360
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "IP"
         Height          =   180
         Left            =   480
         TabIndex        =   9
         Top             =   330
         Width           =   180
      End
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   7
      Top             =   990
      Width           =   5835
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   5205
      TabIndex        =   3
      Top             =   0
      Width           =   5205
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "服务器状态：启用后才能上传并设置给客户端"
         Height          =   180
         Index           =   2
         Left            =   1365
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   135
         Width           =   2700
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "默认服务器：只能有一个默认缺省服务器"
         Height          =   225
         Index           =   0
         Left            =   1365
         TabIndex        =   4
         Top             =   405
         Width           =   3780
      End
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   0
      Left            =   -360
      TabIndex        =   2
      Top             =   5760
      Width           =   5835
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&Q)"
      Height          =   350
      Left            =   3525
      TabIndex        =   1
      Top             =   5925
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   2265
      TabIndex        =   0
      Top             =   5925
      Width           =   1100
   End
End
Attribute VB_Name = "frmCommenLogServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=================================================================
'模块变量
'=================================================================
Private mblnOk              As Boolean      '是否数据处理成功
Private mlngServerNo        As Long         '服务器编号
Private mblnHaveDefault     As Boolean      '是否存在默认服务器
Private mblnChange          As Boolean
Private mblnCollect         As Boolean      '是否收集服务器
Private mstrFileType        As String       '收集类型
Private mblnLoad            As Boolean      '是否数据加载中
Private Enum ServerState
    SS_停用 = 1
    SS_启用 = 0
End Enum

Private Enum ServerType
    ST_共享 = 0
    ST_FTP = 1
End Enum
'=================================================================
'公共接口
'=================================================================
Public Function ShowMe(ByVal lngServerNO As Long, ByVal blnHaveDefault As Boolean) As Boolean
'功能：进行数据的增加修改
'intServerNO=要编辑的服务器编号，=0表示新增数据
'blnHaveDefault=已经存在默认升级服务器
'返回：True-成功，false-失败
    mlngServerNo = lngServerNO
    mblnHaveDefault = blnHaveDefault
    mblnCollect = False
    mstrFileType = ""
    mblnOk = False
    mblnChange = False
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOk
End Function

'=================================================================
'私有方法
'=================================================================
Private Sub chkDefault_Click()

    If chkDefault.Tag <> "" Or mblnLoad Then Exit Sub
    chkDefault.Tag = "设置中"
    If Not mblnHaveDefault Then
        chkDefault.value = 1
        Call MsgBox("首次添加升级服务器，需要启用服务器并设置为默认缺省服务器，不能取消！", vbInformation, gstrSysName)
        chkDefault.Tag = ""
        Exit Sub
    End If
    optServerState(SS_停用).Enabled = chkDefault.value = 1
    optServerState(SS_启用).Enabled = chkDefault.value = 1
    If chkDefault.value = 1 Then
        optServerState(SS_启用).value = True
    End If
    mblnChange = True
    chkDefault.Tag = ""
End Sub

Private Sub cmdCancel_Click()
    If mblnChange Then
        If MsgBox("是否放弃当前编辑内容？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim objConn As clsConnect, strErr As String
    Dim strSQL  As String
    On Error GoTo ErrH
    '输入检查
    If txtServerPath.Text = "" Then
        MsgBox "请输入" & IIf(optServerType(ST_共享).value, "共享目录", "IP地址") & " !", vbInformation, gstrSysName
        txtServerPath.SetFocus
        Exit Sub
    End If
    If ActualLen(txtServerPath.Text) > txtServerPath.MaxLength Then
        MsgBox IIf(optServerType(ST_共享).value, "共享目录", "IP地址") & "超过" & txtServerPath.MaxLength & "位字符长度，请重新输入。", vbInformation, gstrSysName
        txtServerPath.SetFocus
        Exit Sub
    End If
    
    If txtUser.Text = "" Then
        MsgBox "请输入用户名 !", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    End If
    If ActualLen(txtUser.Text) > txtUser.MaxLength Then
        MsgBox "用户名超过" & txtUser.MaxLength & "位字符长度，请重新输入。", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    End If
    
    If txtPWD.Text = "" Then
        MsgBox "请输入密码 !", vbInformation, gstrSysName
        txtPWD.SetFocus
        Exit Sub
    End If
    If ActualLen(txtPWD.Text) > txtPWD.MaxLength Then
        MsgBox "密码超过" & txtPWD.MaxLength & "位字符长度，请重新输入。", vbInformation, gstrSysName
        txtPWD.SetFocus
        Exit Sub
    End If
    
    If txtPort.Text = "" And txtPort.Enabled Then
        MsgBox "请输入端口号 !", vbInformation, gstrSysName
        txtPort.SetFocus
        Exit Sub
    End If
    If MsgBox("是否进行连接校验？", vbYesNo + vbInformation + vbDefaultButton1, gstrSysName) = vbYes Then
        Set objConn = New clsConnect
        If objConn.ToConnect(IIf(optServerType(ST_共享).value, SCT_Share, SCT_FTP), txtServerPath.Text, txtUser.Text, txtPWD.Text, val(txtPort.Text), "", False, strErr) Then
            Call objConn.CloseConnect
        Else
            MsgBox "连接测试失败！信息：" & strErr, vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    strSQL = "Zl_Zlupgradeserver_Update(1," & mlngServerNo & "," & IIf(optServerType(ST_共享).value, 0, 1) & ",'" & Trim(txtServerPath.Text) & "','" & Trim(txtUser.Text) & "'," & SQLAdjust(Cipher(Trim(txtPWD.Text))) & "," & ZVal(txtPort.Text) & "," & IIf(optServerState(SS_启用).value, 1, 0) & "," & IIf(chkDefault.value, 1, 0) & "," & IIf(optServerState(SS_启用).value, 0, IIf(mblnCollect, 1, 0)) & "," & SQLAdjust(IIf(optServerState(SS_启用).value, "", mstrFileType)) & "," & IIf(gblnDelFileServer, "NULL", SQLAdjust(Trim(txtPWD.Text))) & ")"
    Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
    If mlngServerNo = 0 Then
        '插入重要操作日志
        Call SaveAuditLog(1, "文件服务器配置-新增", "新增编号为" & txtno.Text & "的文件服务器")
    Else
        '插入重要操作日志
        Call SaveAuditLog(2, "文件服务器配置-修改", "修改编号为" & mlngServerNo & "的文件服务器")
    End If
    mblnOk = True
    Unload Me
    Exit Sub
ErrH:
    If 1 = 0 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdServerPath_Click()
    Dim strFolderName As String
    On Error Resume Next

    strFolderName = OpenFolder(Me, "选择最新部件的所在目录")
    If Len(strFolderName) = 3 Then
        MsgBox "不能选择根目录(" & strFolderName & ")!", vbInformation, gstrSysName
        Exit Sub
    End If
    If InStr(1, strFolderName, "\\") <> 0 Then
        txtServerPath.Text = strFolderName
    Else
        txtServerPath.Text = "\\" & GetMyCompterName & Mid(strFolderName, 3)
    End If
End Sub

Private Sub Form_Activate()
    If mlngServerNo = 0 Then
        txtServerPath.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    mblnLoad = True
    On Error GoTo ErrH
    
    HookDefend txtPWD.hwnd
    
    If mlngServerNo <> 0 Then
        strSQL = "Select 编号, 类型, 位置, 用户名, 密码, 端口, 是否升级, 是否缺省,是否收集,收集类型 From ZLTOOLS.Zlupgradeserver Where 编号=[1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mlngServerNo)
        If rsTmp.EOF Then
            If MsgBox("当前数据已经被删除，是否新增数据？", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                mlngServerNo = 0
            Else
                On Error Resume Next
                Unload Me
                Exit Sub
            End If
        End If
    End If
    If mlngServerNo = 0 Then
        Me.Caption = "新增文件服务器"
        strSQL = "Select Nvl(Max(编号), 0) + 1 编号 From Zlupgradeserver"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mlngServerNo)
        txtno.Text = rsTmp!编号
        imgCaption.Picture = imgList.ListImages("NEW").Picture
    Else
        Me.Caption = "修改文件服务器"
        imgCaption.Picture = imgList.ListImages("EDIT").Picture
        optServerState(SS_停用).value = val(rsTmp!是否升级 & "") = 0 And val(rsTmp!是否缺省 & "") = 0 And val(rsTmp!是否收集 & "") = 0
        optServerType(val(rsTmp!类型 & "")) = True
        txtServerPath.Text = rsTmp!位置 & ""
        txtUser.Text = rsTmp!用户名 & ""
        txtPWD.Text = Decipher(rsTmp!密码 & "")
        txtPort.Text = rsTmp!端口 & ""
        mblnCollect = val(rsTmp!是否收集 & "") = 1
        mstrFileType = rsTmp!收集类型 & ""
        chkDefault.value = val(rsTmp!是否缺省 & "")
    End If
    If Not mblnHaveDefault Then
        chkDefault.value = 1
        chkDefault.Enabled = False
        optServerState(SS_启用).value = True
        optServerState(SS_启用).Enabled = False
        optServerState(SS_停用).Enabled = False
    End If
    mblnLoad = False
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub optServerState_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub optServerType_Click(Index As Integer)
    mblnChange = True
    lblServerPath.Caption = IIf(Index = ST_共享, "共享目录", "IP地址")
    lblServerPath.Left = lblServerType.Left + lblServerType.Width - lblServerPath.Width
    cmdServerPath.Visible = Index = ST_共享
    txtPort.Enabled = Index = ST_FTP
    If Not txtPort.Enabled Then
        txtPort.Text = ""
    Else
        txtPort.Text = "24"
    End If
End Sub

Private Sub txtPort_Change()
    mblnChange = True
End Sub

Private Sub txtPort_GotFocus()
    Call gclsBase.TxtSelAll(txtPort)
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPWD_Change()
    mblnChange = True
End Sub

Private Sub txtPWD_GotFocus()
    Call gclsBase.TxtSelAll(txtPWD)
End Sub

Private Sub txtServerPath_Change()
    mblnChange = True
End Sub

Private Sub txtUser_Change()
    mblnChange = True
End Sub

Private Sub txtUser_GotFocus()
    Call gclsBase.TxtSelAll(txtUser)
End Sub

