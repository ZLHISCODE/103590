VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClientsEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "客户端设置"
   ClientHeight    =   5310
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmClientsEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame frmMul 
      Caption         =   "快捷设置"
      Height          =   1935
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   6975
      Begin VB.TextBox txtbeforeIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3690
         MaxLength       =   3
         TabIndex        =   40
         Tag             =   "IP地址"
         Top             =   1470
         Width           =   315
      End
      Begin VB.TextBox txtbeforeIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   39
         Tag             =   "IP地址"
         Top             =   1470
         Width           =   315
      End
      Begin VB.TextBox txtbeforeIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2790
         MaxLength       =   3
         TabIndex        =   38
         Tag             =   "IP地址"
         Top             =   1470
         Width           =   315
      End
      Begin VB.TextBox txtbeforeIp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2355
         MaxLength       =   3
         TabIndex        =   37
         Tag             =   "IP地址"
         Top             =   1470
         Width           =   315
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   12
         Left            =   4485
         MaxLength       =   3
         TabIndex        =   46
         Tag             =   "IP"
         Top             =   1425
         Width           =   390
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   11
         Left            =   2325
         MaxLength       =   20
         TabIndex        =   45
         TabStop         =   0   'False
         Tag             =   "IP"
         Text            =   "   ．   ．   ．"
         Top             =   1440
         Width           =   1725
      End
      Begin VB.OptionButton optLink 
         Caption         =   "当前客户端"
         Height          =   180
         Index           =   0
         Left            =   855
         TabIndex        =   44
         Top             =   1080
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optLink 
         Caption         =   "相同部门客户端"
         Height          =   180
         Index           =   1
         Left            =   2295
         TabIndex        =   43
         Top             =   1080
         Width           =   1560
      End
      Begin VB.OptionButton optLink 
         Caption         =   "所有客户端"
         Height          =   180
         Index           =   2
         Left            =   4185
         TabIndex        =   42
         Top             =   1080
         Width           =   1200
      End
      Begin VB.OptionButton optLink 
         Caption         =   "IP范围"
         Height          =   180
         Index           =   3
         Left            =   855
         TabIndex        =   41
         Top             =   1500
         Width           =   885
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   34
         Text            =   "1"
         Top             =   240
         Width           =   435
      End
      Begin VB.CheckBox chkNoLimit 
         Caption         =   "无限制(&T)"
         Height          =   195
         Left            =   3120
         TabIndex        =   32
         Top             =   285
         Value           =   1  'Checked
         Width           =   1110
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   10
         Left            =   3300
         MaxLength       =   1
         TabIndex        =   31
         Tag             =   "说明"
         Top             =   600
         Width           =   435
      End
      Begin VB.CheckBox chkStartupVideo 
         Caption         =   "启用视频源"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   615
         Width           =   1215
      End
      Begin MSComCtl2.UpDown udLink 
         Height          =   300
         Left            =   2700
         TabIndex        =   33
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtEdit(9)"
         BuddyDispid     =   196611
         BuddyIndex      =   9
         OrigLeft        =   1485
         OrigTop         =   3375
         OrigRight       =   1725
         OrigBottom      =   3675
         Max             =   9
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应用于"
         Height          =   180
         Left            =   240
         TabIndex        =   48
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Index           =   11
         Left            =   4605
         TabIndex        =   47
         Top             =   1380
         Width           =   90
      End
      Begin VB.Label lblLinkNumber 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "导航台登录允许的连接数(&L)"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   285
         Width           =   2250
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "院区编号(&C)"
         Height          =   180
         Index           =   9
         Left            =   2295
         TabIndex        =   35
         Top             =   645
         Width           =   990
      End
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   4440
      MaxLength       =   50
      TabIndex        =   28
      Tag             =   "用途"
      Top             =   1500
      Width           =   2795
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   21
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   12
      Tag             =   "CPU"
      ToolTipText     =   "用于远程控制和客户端升级"
      Top             =   2280
      Width           =   1725
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   13
      Left            =   4425
      MaxLength       =   30
      TabIndex        =   13
      Tag             =   "内存"
      Top             =   2280
      Width           =   2795
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2445
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "IP地址"
      Top             =   270
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1995
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "IP地址"
      Top             =   270
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1545
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "IP地址"
      Top             =   270
      Width           =   315
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "IP地址"
      Top             =   270
      Width           =   315
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   8
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   11
      Tag             =   "说明"
      Top             =   1920
      Width           =   6140
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   10
      Tag             =   "院区"
      Top             =   1515
      Width           =   1725
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   4425
      MaxLength       =   50
      TabIndex        =   9
      Tag             =   "操作系统"
      Top             =   1110
      Width           =   2795
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "硬盘"
      Top             =   1110
      Width           =   1725
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   4425
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "内存"
      Top             =   690
      Width           =   2795
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "CPU"
      Top             =   690
      Width           =   1725
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   4425
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "工作站"
      Top             =   240
      Width           =   2795
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "IP"
      Text            =   "   ．   ．   ．"
      Top             =   240
      Width           =   1725
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5055
      TabIndex        =   23
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   24
      Top             =   4845
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6165
      TabIndex        =   22
      Top             =   4845
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "用途(Y)"
      Height          =   180
      Index           =   7
      Left            =   3705
      TabIndex        =   29
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "操作系统(&X)"
      Height          =   180
      Index           =   5
      Left            =   3345
      TabIndex        =   27
      Top             =   1170
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "管理员用户"
      Height          =   180
      Index           =   12
      Left            =   120
      TabIndex        =   26
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "管理员密码"
      Height          =   180
      Index           =   10
      Left            =   3360
      TabIndex        =   25
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "说明(&S)"
      Height          =   180
      Index           =   8
      Left            =   390
      TabIndex        =   21
      Top             =   1980
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "部门(&B)"
      Height          =   180
      Index           =   6
      Left            =   390
      TabIndex        =   20
      Tag             =   "院区"
      Top             =   1575
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "硬盘(&E)"
      Height          =   180
      Index           =   4
      Left            =   390
      TabIndex        =   19
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "内存(&M)"
      Height          =   180
      Index           =   3
      Left            =   3690
      TabIndex        =   18
      Top             =   750
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "CPU(&U)"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   17
      Top             =   750
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "客户端(&G)"
      Height          =   180
      Index           =   1
      Left            =   3510
      TabIndex        =   16
      Top             =   300
      Width           =   810
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "IP(&I)"
      Height          =   180
      Index           =   0
      Left            =   570
      TabIndex        =   0
      Top             =   300
      Width           =   450
   End
End
Attribute VB_Name = "frmClientsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EditType
    新增 = 0
    修改
End Enum

Private mstrStartupVideo As String
Private mstrUser As String '远程控制帐号
Private mstrPwd As String   '远程控制密码
Private mstrRemarks As String
Dim mEditType As EditType
Dim mStrIP As String
Dim mstrCompterName As String

Dim mblnReturn As Boolean
Dim mblnFirst As Boolean

Private Sub chkNoLimit_Click()
    txtEdit(9).Enabled = chkNoLimit.value = 0
    udLink.Enabled = chkNoLimit.value = 0
    
    If txtEdit(9).Enabled Then
        txtEdit(9).BackColor = txtEdit(1).BackColor
    Else
        txtEdit(9).BackColor = Me.BackColor
    End If
        
    If Visible Then
        If Trim(txtIp(0).Text) <> "" And Trim(txtIp(1).Text) <> "" And Trim(txtIp(2).Text) <> "" And Trim(txtIp(3).Text) <> "" And Trim(txtEdit(1).Text) <> "" Then
            cmdSave.Enabled = True
        End If
    End If
End Sub

Private Sub chkStartupVideo_Click()
    If Visible Then
        If Trim(txtIp(0).Text) <> "" And Trim(txtIp(1).Text) <> "" And Trim(txtIp(2).Text) <> "" And Trim(txtIp(3).Text) <> "" And Trim(txtEdit(1).Text) <> "" Then
            cmdSave.Enabled = True
        End If
    End If
    
    mstrStartupVideo = chkStartupVideo.value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If mEditType = 修改 And optLink(0).value <> True Then
        '验证身份并输入操作说明
        mstrRemarks = "修改客户端：" & mstrCompterName
        If Not CheckAuditStatus("0308", "修改", mstrRemarks) Then Exit Sub
    End If
    If Not IsValid() Then Exit Sub
    If Not Save() Then Exit Sub
    If mEditType = 新增 Then
        Call ClearInfor
    Else
        Unload Me
    End If
    mblnReturn = True
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    Call InitCard '初始卡片
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{tab}"
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrRemarks = ""
End Sub

Private Sub optLink_Click(Index As Integer)
    Dim i As Integer
    If optLink(3).value Then
        For i = 0 To 3
            txtbeforeIp(i).Enabled = True
            txtEdit(12).Enabled = True
        Next
    Else
        For i = 0 To 3
            txtbeforeIp(i).Enabled = False
            txtEdit(12).Enabled = False
        Next
    End If
    If Trim(txtIp(0).Text) <> "" And Trim(txtIp(1).Text) <> "" And Trim(txtIp(2).Text) <> "" And Trim(txtIp(3).Text) <> "" And Trim(txtEdit(1).Text) <> "" Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Trim(txtIp(0).Text) <> "" And Trim(txtIp(1).Text) <> "" And Trim(txtIp(2).Text) <> "" And Trim(txtIp(3).Text) <> "" And Trim(txtEdit(1).Text) <> "" Then
        cmdSave.Enabled = True
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    txtEdit(Index).SelStart = 0
    txtEdit(Index).SelLength = Len(txtEdit(Index).Text)
End Sub

'ip, 客户端, cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级标志, 禁止使用,连接数
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 9 Then
        If InStr("123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    
    If Index = 10 Or Index = 12 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Function IsValid() As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能:检证数据的合法性
    '返回:数据合法返回true,否则false
    '----------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strTmp As String
    Dim strBeforeIp As String, strAfterIp As String, strErr As String
    
    For intIndex = 0 To 8
        strTmp = Trim(txtEdit(intIndex).Text)
        If intIndex = 0 Or intIndex = 1 Then
            If strTmp = "" Then
                MsgBox txtEdit(intIndex).Tag & "必需输入!", vbInformation + vbOKOnly, gstrSysName
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
        If strTmp <> "" Then
            If LenB(StrConv(strTmp, vbFromUnicode)) > txtEdit(intIndex).MaxLength Then
                MsgBox txtEdit(intIndex).Tag & "超长,最多能输入" & txtEdit(intIndex).MaxLength / 2 & "个汉字或" & txtEdit(intIndex).MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
        If InStr(1, strTmp, "'") <> 0 Then
            MsgBox txtEdit(intIndex).Tag & "不能输入单引号!", vbInformation + vbOKOnly, gstrSysName
            If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
            Exit Function
        End If
    Next
    For intIndex = 0 To 3
        strTmp = Trim(txtIp(intIndex).Text)
        If strTmp = "" Then
            MsgBox txtIp(intIndex).Tag & "必需输入!", vbInformation + vbOKOnly, gstrSysName
            If txtIp(intIndex).Enabled Then txtIp(intIndex).SetFocus
            Exit Function
        End If
        If LenB(StrConv(strTmp, vbFromUnicode)) > txtIp(intIndex).MaxLength Then
            MsgBox txtIp(intIndex).Tag & "超长,最多能输入" & txtIp(intIndex).MaxLength / 2 & "个汉字或" & txtIp(intIndex).MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
            If txtIp(intIndex).Enabled Then txtIp(intIndex).SetFocus
            Exit Function
        End If
        If InStr(1, strTmp, "'") <> 0 Then
            MsgBox txtIp(intIndex).Tag & "不能输入单引号!", vbInformation + vbOKOnly, gstrSysName
            If txtIp(intIndex).Enabled Then txtIp(intIndex).SetFocus
            Exit Function
        End If
        If Not IsNumeric(strTmp) Then
            MsgBox txtIp(intIndex).Tag & "只能输入数字!", vbInformation + vbOKOnly, gstrSysName
            If txtIp(intIndex).Enabled Then txtIp(intIndex).SetFocus
            Exit Function
        End If
    Next
    
    '连接数检查
    If txtEdit(9).Enabled Then
        If Val(txtEdit(9).Text) = 0 Then
            MsgBox "请输入正确的允许连接数!", vbInformation, gstrSysName
            txtEdit(9).SetFocus: Exit Function
        End If
    End If
    
    '检查IP范围
    If optLink(3).value Then
        strBeforeIp = Trim(txtbeforeIp(0).Text) & "." & Trim(txtbeforeIp(1).Text) & "." & Trim(txtbeforeIp(2).Text) & "." & Trim(txtbeforeIp(3).Text)
        strAfterIp = Trim(txtbeforeIp(0).Text) & "." & Trim(txtbeforeIp(1).Text) & "." & Trim(txtbeforeIp(2).Text) & "." & Trim(txtEdit(12).Text)
        CheckIpValidate strBeforeIp, strAfterIp, strErr
        If strErr <> "" Then
            MsgBox strErr, vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Private Function Save() As Boolean
    '----------------------------------------------------------------------------------------------------
    '功能:保存数据
    '返回:
    '----------------------------------------------------------------------------------------------------
    Dim strIp As String
    Dim str工作站 As String
    Dim strCPU As String
    Dim str内存 As String
    Dim str硬盘 As String
    Dim str操作系统 As String
    Dim str部门 As String
    Dim str用途 As String
    Dim str说明 As String
    Dim int连接数 As String
    Dim str站点   As String
    Dim str启用视频源 As String
    Dim str管理员用户 As String
    Dim str管理员密码 As String
    Dim rsTmp As New ADODB.Recordset
    Dim strVideoNum As String
    Dim strNote As String, strMsg As String
    Dim strBatchEditItems As String
    
    Dim strSQL As String, strSQLDel As String
    Dim strCurIP As String, strLink As String
    Dim strIpBegin As String, strIpEnd As String
    Dim lngApply As Long
    
    strIp = Trim(txtIp(0).Text) & "." & Trim(txtIp(1).Text) & "." & Trim(txtIp(2).Text) & "." & Trim(txtIp(3).Text)
    str工作站 = Trim(txtEdit(1).Text)
    strCPU = Trim(txtEdit(2).Text)
    str内存 = Trim(txtEdit(3).Text)
    str硬盘 = Trim(txtEdit(4).Text)
    str操作系统 = Trim(txtEdit(5).Text)
    str部门 = Trim(txtEdit(6).Text)
    str用途 = Trim(txtEdit(7).Text)
    str说明 = Trim(txtEdit(8).Text)
    int连接数 = IIf(chkNoLimit.value = 1, 0, Val(txtEdit(9).Text)) '0-表示无限制
    str站点 = Trim(txtEdit(10).Text)
    str启用视频源 = IIf(chkStartupVideo.value = 1, 1, 0)     '1-表示启用视频源
    str管理员用户 = Trim(txtEdit(21).Text)
    str管理员密码 = Trim(txtEdit(13).Text)
    mstrUser = str管理员用户
    mstrPwd = str管理员密码
    
    '连接数应用范围
    If txtEdit(9).Enabled Or chkNoLimit.Enabled Then
        If optLink(1).value Then
            '应用于相同部门客户端
            lngApply = 1
            strMsg = "所属部门为“" & str部门 & "”的客户端上"
        ElseIf optLink(2).value Then
            '应用于所有客户端
            lngApply = 2
            strMsg = "所有客户端上"
        ElseIf optLink(3).value Then
            lngApply = 3
            Call GetIpRange(strIpBegin, strIpEnd)
            If Not (strIpBegin = strIpEnd And strIpBegin = strIp) Then
                If strIpBegin = strIpEnd Then
                    strMsg = "IP为" & strIpBegin & "的客户端上"
                Else
                    strMsg = "IP范围为" & strIpBegin & "至" & strIpEnd & "的客户端上"
                End If
            End If
        End If
    End If
    
    Save = False
    
    strVideoNum = gobjRegister.zlRegInfo("影像视频设备数量")

    If Val(strVideoNum) > 0 And str启用视频源 = 1 Then
        strSQL = "select count(启用视频源) as 启用数量 from zlClients where 启用视频源=1"
    
        Call OpenRecordset(rsTmp, strSQL, Me.Caption)
    
        If Val(Nvl(rsTmp!启用数量)) >= Val(strVideoNum) Then
            MsgBox "可启用的影像视频设备数量已达到最大值,不能再启用!" & vbNewLine & err.Description, vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If
    
    If mEditType = 新增 Or mStrIP <> strIp Or mstrCompterName <> str工作站 Then
        If strVideoNum = "" Then str启用视频源 = ""
        
        strSQL = "Zl_Zlclients_Set(0,Null,'" & str工作站 & "','" & strIp & "'," & _
            IIf(strCPU = "", "Null", "'" & strCPU & "'") & "," & _
            IIf(str内存 = "", "Null", "'" & str内存 & "'") & "," & _
            IIf(str硬盘 = "", "Null", "'" & str硬盘 & "'") & "," & _
            IIf(str操作系统 = "", "Null", "'" & str操作系统 & "'") & "," & _
            IIf(str部门 = "", "Null", "'" & str部门 & "'") & "," & _
            IIf(str用途 = "", "Null", "'" & str用途 & "'") & "," & _
            IIf(str说明 = "", "Null", "'" & str说明 & "'") & ",Null,Null," & int连接数 & "," & _
            IIf(str站点 = "", "Null", "'" & str站点 & "'") & "," & lngApply & "," & _
            IIf(strIpBegin = "", "Null", "'" & strIpBegin & "'") & "," & _
            IIf(strIpEnd = "", "Null", "'" & strIpEnd & "'") & ",'" & str启用视频源 & "'," & _
            IIf(str管理员用户 = "", "'空空'", "'" & str管理员用户 & "'") & "," & _
            IIf(str管理员密码 = "", "'空空'", "" & SQLAdjust(Cipher(str管理员密码)) & "") & ")"
    Else
        strSQL = "Zl_Zlclients_Set(1,Null,'" & str工作站 & "','" & strIp & "'," & _
            IIf(strCPU = "", "Null", "'" & strCPU & "'") & "," & _
            IIf(str内存 = "", "Null", "'" & str内存 & "'") & "," & _
            IIf(str硬盘 = "", "Null", "'" & str硬盘 & "'") & "," & _
            IIf(str操作系统 = "", "Null", "'" & str操作系统 & "'") & "," & _
            IIf(str部门 = "", "Null", "'" & str部门 & "'") & "," & _
            IIf(str用途 = "", "Null", "'" & str用途 & "'") & "," & _
            IIf(str说明 = "", "Null", "'" & str说明 & "'") & ",Null,Null," & int连接数 & "," & _
            IIf(str站点 = "", "Null", "'" & str站点 & "'") & "," & lngApply & "," & _
            IIf(strIpBegin = "", "Null", "'" & strIpBegin & "'") & "," & _
            IIf(strIpEnd = "", "Null", "'" & strIpEnd & "'") & ",'" & mstrStartupVideo & "'," & _
            IIf(str管理员用户 = "", "'空空'", "'" & str管理员用户 & "'") & "," & _
            IIf(str管理员密码 = "", "'空空'", "" & SQLAdjust(Cipher(str管理员密码)) & "") & ")"
        strCurIP = mStrIP
    End If

    err = 0
    On Error Resume Next
    If mEditType = 修改 And (mStrIP <> strIp Or mstrCompterName <> str工作站) Then
        strSQLDel = "Zl_Zlclients_Delete('" & mstrCompterName & "','" & mStrIP & "')"
        Call ExecuteProcedure(strSQLDel, Me.Caption)
    End If
    
    Call ExecuteProcedure(strSQL, Me.Caption)
    If err <> 0 Then
        MsgBox "已经设置了相同IP地址或客户端,不能再设!" & vbNewLine & err.Description, vbInformation + vbDefaultButton1, gstrSysName
        err.Clear
        Exit Function
    End If
    Save = True
    If mEditType = 修改 Then
        If mstrCompterName <> str工作站 Then strNote = ",客户端由“" & mstrCompterName & "”修改为了“" & str工作站 & "”"
        If mStrIP <> strIp Then strNote = strNote & ",IP地址由“" & mStrIP & "”修改为了“" & strIp & "”"
        If lblEdit(2).Tag <> strCPU Then strNote = strNote & ",CPU由“" & IIf(lblEdit(2).Tag = "", "空", lblEdit(2).Tag) & "”修改为了“" & IIf(strCPU = "", "空", strCPU) & "”"
        If lblEdit(3).Tag <> str内存 Then strNote = strNote & ",内存由“" & IIf(lblEdit(3).Tag = "", "空", lblEdit(3).Tag) & "”修改为了“" & IIf(str内存 = "", "空", str内存) & "”"
        If lblEdit(4).Tag <> str硬盘 Then strNote = strNote & ",硬盘由“" & IIf(lblEdit(4).Tag = "", "空", lblEdit(4).Tag) & "”修改为了“" & IIf(str硬盘 = "", "空", str硬盘) & "”"
        If lblEdit(5).Tag <> str操作系统 Then strNote = strNote & ",操作系统由“" & IIf(lblEdit(5).Tag = "", "空", lblEdit(5).Tag) & "”修改为了“" & IIf(str操作系统 = "", "空", str操作系统) & "”"
        If lblEdit(6).Tag <> str部门 Then strNote = strNote & ",部门由“" & IIf(lblEdit(6).Tag = "", "空", lblEdit(6).Tag) & "”修改为了“" & IIf(str部门 = "", "空", str部门) & "”"
        If lblEdit(7).Tag <> str用途 Then strNote = strNote & ",用途由“" & IIf(lblEdit(7).Tag = "", "空", lblEdit(7).Tag) & "”修改为了“" & IIf(str用途 = "", "空", str用途) & "”"
        If lblEdit(12).Tag <> str管理员用户 Then strNote = strNote & ",管理员用户由“" & IIf(lblEdit(12).Tag = "", "空", lblEdit(12).Tag) & "”修改为了“" & IIf(str管理员用户 = "", "空", str管理员用户) & "”"
        If lblEdit(9).Tag <> str站点 Then
            strNote = strNote & ",院区由“" & IIf(lblEdit(9).Tag = "", "空", lblEdit(9).Tag) & "”修改为了“" & IIf(str站点 = "", "空", str站点) & "”"
            strBatchEditItems = "“院区”"
        End If
        If chkNoLimit.value = 0 Then
            If lblLinkNumber.Tag <> int连接数 Then
                strNote = strNote & ",连接数由“" & lblLinkNumber.Tag & "”修改为了“" & int连接数 & "”"
                strBatchEditItems = strBatchEditItems & "“连接数”"
            End If
        Else
            If lblLinkNumber.Tag > 0 Then
                strNote = strNote & ",连接数由“" & lblLinkNumber.Tag & "”修改为了无限制"
                strBatchEditItems = strBatchEditItems & "“连接数”"
            End If
        End If
        If mstrStartupVideo <> str启用视频源 And chkStartupVideo.Enabled = True Then
            If mstrStartupVideo = "1" Then
                strNote = strNote & ",并将视频源设置为启用"
                strBatchEditItems = strBatchEditItems & "“视频源”"
            Else
                strNote = strNote & ",并将视频源设置为不启用"
                strBatchEditItems = strBatchEditItems & "“视频源”"
            End If
        End If
        If strMsg <> "" Then
            strNote = strNote & ",并将修改项" & strBatchEditItems & "的值应用到" & strMsg
        End If
        '插入重要操作日志
        Call SaveAuditLog(2, "修改", "修改客户端“" & mstrCompterName & "”" & "成功" & _
             IIf(strNote = "", "", "，其中" & Mid(strNote, 2)), mstrRemarks, "0308")
    End If
End Function

Public Sub ShowEdit(ByVal strIp As String, ByVal strCompterName As String, ByVal intEditType As EditType, ByRef blnRetun As Boolean, _
                                Optional ByRef strUser As String, Optional ByRef strPwd As String)
    '-------------------------------------------------------------------------------
    '--功能：显示和编辑客户端信息
    '--参数：intEditType-编辑类型
    '        StrIP:IP地址
    '--返回：blnRetun-编辑成功返回true,否则返回false
    'strUser,strPwd 返回的远程登录帐号密码
    '-------------------------------------------------------------------------------
    mEditType = intEditType
    mStrIP = strIp
    mstrCompterName = strCompterName
    
    Me.cmdSave.Enabled = False
    mblnReturn = False
    Me.Show 1, frmMDIMain
    
    strUser = mstrUser
    strPwd = mstrPwd
    blnRetun = mblnReturn
    
End Sub

Private Sub InitCard()
    '功能:初始化卡片
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strArr
    Dim i As Integer
    Dim strVideoNum As String
    
    txtEdit(0).Enabled = False
    txtEdit(11).Enabled = False
    txtEdit(12).Enabled = False
    
On Error GoTo ErrHandle
    strVideoNum = gobjRegister.zlRegInfo("影像视频设备数量")
    
    If strVideoNum = "" Then
        chkStartupVideo.value = 1
        chkStartupVideo.Enabled = False
    ElseIf CInt(strVideoNum) <= 0 Then
        strSQL = "update zlClients set 启用视频源=0"
        gcnOracle.Execute strSQL
        
        chkStartupVideo.Enabled = False
    End If

    If mEditType = 新增 Then
        Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client_maxip")
        With rsTmp
            Do While Not .EOF
                strArr = Split(IIf(IsNull(!IP), "", !IP), ".")
                If UBound(strArr) = 3 Then
                    txtIp(0).Text = strArr(0)
                    txtIp(1).Text = strArr(1)
                    txtIp(2).Text = strArr(2)
                    txtIp(3).Text = Val(strArr(3)) + 1
                Else
                    txtIp(0).Text = ""
                    txtIp(1).Text = ""
                    txtIp(2).Text = ""
                    txtIp(3).Text = ""
                End If
                .MoveNext
            Loop
            .Close
        End With
        If Trim(txtIp(0).Text) = "" Then
            txtIp(0).SelStart = 0
            txtIp(0).SelLength = 3
            If txtIp(0).Enabled Then
                txtIp(0).SetFocus
            End If
        Else
            txtIp(3).SelStart = 0
            txtIp(3).SelLength = 3
            If txtIp(3).Enabled Then
                txtIp(3).SetFocus
            End If
        End If
        
        For i = 0 To 3
            txtbeforeIp(i).Text = txtIp(i).Text
            txtbeforeIp(i).Enabled = False
            txtEdit(12).Enabled = False
            If i = 3 Then
                txtEdit(12).Text = txtIp(i).Text
            End If
        Next
        
        Exit Sub
    End If
    
    Call ClearInfor
    
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client", mstrCompterName)
    
    With rsTmp
        Do While Not .EOF
            strArr = Split(IIf(IsNull(!IP), "", !IP), ".")
            If UBound(strArr) = 3 Then
                txtIp(0).Text = strArr(0)
                txtIp(1).Text = strArr(1)
                txtIp(2).Text = strArr(2)
                txtIp(3).Text = strArr(3)
            Else
                txtIp(0).Text = ""
                txtIp(1).Text = ""
                txtIp(2).Text = ""
                txtIp(3).Text = ""
            End If
            txtEdit(1).Text = Nvl(!工作站)
            txtEdit(2).Text = Nvl(!cpu)
            lblEdit(2).Tag = txtEdit(2).Text
            txtEdit(3).Text = Nvl(!内存)
            lblEdit(3).Tag = txtEdit(3).Text
            txtEdit(4).Text = Nvl(!硬盘)
            lblEdit(4).Tag = txtEdit(4).Text
            txtEdit(5).Text = Nvl(!操作系统)
            lblEdit(5).Tag = txtEdit(5).Text
            txtEdit(6).Text = Nvl(!部门)
            lblEdit(6).Tag = txtEdit(6).Text
            txtEdit(7).Text = Nvl(!用途)
            lblEdit(7).Tag = txtEdit(7).Text
            txtEdit(8).Text = Nvl(!说明)
            txtEdit(10).Text = Nvl(!站点)
            lblEdit(9).Tag = txtEdit(10).Text
            mstrStartupVideo = Nvl(!启用视频源, "")
            chkStartupVideo.value = IIf(strVideoNum = "", 1, Nvl(!启用视频源, 0))
            txtEdit(21).Text = Nvl(!管理员用户)
            lblEdit(12).Tag = txtEdit(21).Text
            txtEdit(13).Text = Decipher(Nvl(!管理员密码))
            
            lblLinkNumber.Tag = 0
            If Nvl(!连接数, 0) > 0 Then
                chkNoLimit.value = 0
                txtEdit(9).Text = Nvl(!连接数, 0)
                lblLinkNumber.Tag = txtEdit(9).Text
            End If
            If Nvl(!禁止使用, 0) = 1 Then
                txtEdit(9).BackColor = Me.BackColor
                txtEdit(9).Enabled = False
                chkNoLimit.Enabled = False
                chkStartupVideo.Enabled = False
                optLink(0).Enabled = False
                optLink(1).Enabled = False
                optLink(2).Enabled = False
                optLink(3).Enabled = False
            End If
            
            .MoveNext
        Loop
        .Close
    End With
    txtIp(3).SelStart = 0
    txtIp(3).SelLength = 3
    If txtIp(3).Enabled Then
        txtIp(3).SetFocus
    End If
    
    For i = 0 To 3
        txtbeforeIp(i).Text = txtIp(i).Text
        txtbeforeIp(i).Enabled = False
        txtEdit(12).Enabled = False
        If i = 3 Then
            txtEdit(12).Text = txtIp(i).Text
        End If
    Next
    
    cmdSave.Enabled = False
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub ClearInfor()
    '功能:将所有信息清除
    Dim intIndex As Integer
    For intIndex = 0 To 8
        txtEdit(intIndex) = ""
    Next
    txtEdit(0).Text = "   ．   ．   ．"
    optLink(0).value = True
End Sub

Private Sub txtIp_Change(Index As Integer)
    Dim lngLineNo As Long '行号
    Dim lngColNo As Long  '列号
    err = 0
    On Error Resume Next
    If Trim(txtIp(0).Text) <> "" And Trim(txtIp(1).Text) <> "" And Trim(txtIp(2).Text) <> "" And Trim(txtIp(3).Text) <> "" And Trim(txtEdit(1).Text) <> "" Then
        cmdSave.Enabled = True
    End If
    Call GetCursorPos(Me.txtIp(Index).hwnd, lngLineNo, lngColNo)
    If lngColNo > 3 Then
        If Index < 3 Then
            If txtIp(Index + 1).Enabled Then txtIp(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtIp_GotFocus(Index As Integer)
    txtIp(Index).SelStart = 0
    txtIp(Index).SelLength = Len(txtIp(Index).Text)
End Sub

Private Sub txtIp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lngLineNo As Long '行号
    Dim lngColNo As Long  '列号
    err = 0
    On Error Resume Next
    
    Call GetCursorPos(Me.txtIp(Index).hwnd, lngLineNo, lngColNo)
    
    Select Case KeyCode
    Case 37     '<-
        
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtIp(Index - 1).Enabled Then
                txtIp(Index - 1).SelStart = Len(txtIp(Index - 1))
                txtIp(Index - 1).SetFocus
            End If
        End If
    Case 39     '->
        If Index < 3 Then
            If lngColNo <= Len(txtIp(Index)) Then Exit Sub
            If txtIp(Index + 1).Enabled Then
                txtIp(Index + 1).SelStart = 0
                txtIp(Index + 1).SetFocus
            End If
        End If
    Case 8     'BACKSPACE
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtIp(Index - 1).Enabled Then
                txtIp(Index - 1).SelStart = Len(txtIp(Index - 1))
                txtIp(Index - 1).SetFocus
            End If
        End If
    Case Else
    End Select

End Sub

Private Sub txtIp_KeyPress(Index As Integer, KeyAscii As Integer)
    err = 0
    On Error Resume Next
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then
                If KeyAscii = Asc(".") Then
                    If Index < 3 And Index >= 0 And Trim(txtIp(Index)) <> "" Then
                        If txtIp(Index + 1).Enabled Then txtIp(Index + 1).SetFocus
                    End If
                End If
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Public Sub GetCursorPos(ByVal hwnd5 As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long
    Dim lParam As Long, wParam As Long
    Dim K As Long
    err = 0
    On Error Resume Next
    i = SendMessage(hwnd5, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16 '取得目前光标所在位置前有多少个Byte
    LineNo = SendMessage(hwnd5, EM_LINEFROMCHAR, j, 0) '取得光标前面有多少行
    LineNo = LineNo + 1
    K = SendMessage(hwnd5, EM_LINEINDEX, -1, 0)
    '取得目前光标所在行前面有多少个Byte
    ColNo = j - K + 1
End Sub

Private Sub txtbeforeIp_Change(Index As Integer)
    Dim lngLineNo As Long '行号
    Dim lngColNo  As Long '列号
    err = 0
    On Error Resume Next
    If Trim(txtbeforeIp(0).Text) <> "" And Trim(txtbeforeIp(1).Text) <> "" And Trim(txtbeforeIp(2).Text) <> "" And Trim(txtbeforeIp(3).Text) <> "" And Trim(txtEdit(10).Text) <> "" Then
        cmdSave.Enabled = True
    End If
    Call GetCursorPos(Me.txtbeforeIp(Index).hwnd, lngLineNo, lngColNo)
    If lngColNo > 3 Then
        If Index < 3 Then
            If txtbeforeIp(Index + 1).Enabled Then txtbeforeIp(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtbeforeIp_GotFocus(Index As Integer)
    txtbeforeIp(Index).SelStart = 0
    txtbeforeIp(Index).SelLength = Len(txtbeforeIp(Index).Text)
End Sub

Private Sub txtbeforeIp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lngLineNo As Long '行号
    Dim lngColNo  As Long '列号
    err = 0
    Call GetCursorPos(Me.txtbeforeIp(Index).hwnd, lngLineNo, lngColNo)
    
    Select Case KeyCode
    Case 37    '<-
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtbeforeIp(Index - 1).Enabled Then
                txtbeforeIp(Index - 1).SelStart = Len(txtbeforeIp(Index - 1))
                txtbeforeIp(Index - 1).SetFocus
            End If
        End If
    Case 39    '->
        If Index < 3 Then
            If lngColNo <= Len(txtbeforeIp(Index)) Then Exit Sub
            If txtbeforeIp(Index + 1).Enabled Then
                txtbeforeIp(Index + 1).SelStart = 0
                txtbeforeIp(Index + 1).SetFocus
            End If
        End If
    Case 8     'BACKSPACE
        If Index > 0 Then
            If lngColNo > 1 Then Exit Sub
            If txtbeforeIp(Index - 1).Enabled Then
                txtbeforeIp(Index - 1).SelStart = Len(txtbeforeIp(Index - 1))
                txtbeforeIp(Index - 1).SetFocus
            End If
        End If
    End Select
    
End Sub

Private Sub txtbeforeIp_KeyPress(Index As Integer, KeyAscii As Integer)
    err = 0
    On Error Resume Next
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then
                If KeyAscii = Asc(".") Then
                    If Index < 3 And Index >= 0 And Trim(txtIp(Index)) <> "" Then
                        If txtbeforeIp(Index + 1).Enabled Then txtbeforeIp(Index + 1).SetFocus
                    End If
                End If
                KeyAscii = 0
            End If
        End If
    End If
End Sub


Private Sub GetIpRange(ByRef strIpBegin As String, ByRef strIpEnd As String)
    Dim i As Long
    Dim strNullIndex As String
    Dim strBeforeIp As String
    Dim strLaterIp As String
    
    If Trim(txtbeforeIp(0).Text) = "" Or Trim(txtbeforeIp(1).Text) = "" Or Trim(txtbeforeIp(2).Text) = "" Or Trim(txtbeforeIp(3).Text) = "" Then
        For i = 0 To 3
          If Trim(txtbeforeIp(i).Text) = "" Then
             strNullIndex = i
             Exit For
          End If
        Next
        
        Select Case strNullIndex
        Case 0
        Case 1
             strBeforeIp = Val(Trim(txtbeforeIp(0).Text)) & "." & Val(Trim(txtbeforeIp(1).Text)) & "." & Val(Trim(txtbeforeIp(2).Text)) & "." & Val(Trim(txtbeforeIp(3).Text))
             strLaterIp = Val(Trim(txtbeforeIp(0).Text)) & "." & "255" & "." & "255" & "." & "255"
        Case 2
             strBeforeIp = Val(Trim(txtbeforeIp(0).Text)) & "." & Val(Trim(txtbeforeIp(1).Text)) & "." & Val(Trim(txtbeforeIp(2).Text)) & "." & Val(Trim(txtbeforeIp(3).Text))
             strLaterIp = Val(Trim(txtbeforeIp(0).Text)) & "." & Val(Trim(txtbeforeIp(1).Text)) & "." & "255" & "." & "255"
        Case 3
             strBeforeIp = Val(Trim(txtbeforeIp(0).Text)) & "." & Val(Trim(txtbeforeIp(1).Text)) & "." & Val(Trim(txtbeforeIp(2).Text)) & "." & Val(Trim(txtbeforeIp(3).Text))
             strLaterIp = Val(Trim(txtbeforeIp(0).Text)) & "." & Val(Trim(txtbeforeIp(1).Text)) & "." & Val(Trim(txtbeforeIp(2).Text)) & "." & Val(Trim(txtEdit(12).Text))
        End Select
    Else
        strBeforeIp = Val(Trim(txtbeforeIp(0).Text)) & "." & Val(Trim(txtbeforeIp(1).Text)) & "." & Val(Trim(txtbeforeIp(2).Text)) & "." & Val(Trim(txtbeforeIp(3).Text))
        strLaterIp = Val(Trim(txtbeforeIp(0).Text)) & "." & Val(Trim(txtbeforeIp(1).Text)) & "." & Val(Trim(txtbeforeIp(2).Text)) & "." & Val(Trim(txtEdit(12).Text))
    End If
    strIpBegin = strBeforeIp: strIpEnd = strLaterIp
End Sub

