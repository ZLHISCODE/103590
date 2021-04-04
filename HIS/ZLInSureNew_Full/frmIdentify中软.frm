VERSION 5.00
Begin VB.Form frmIdentify中软 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify中软.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txt新密码 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1635
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox txt新密码 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1635
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1290
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.CheckBox chk离休 
      Caption         =   "离休病人(&L)"
      Height          =   315
      Left            =   4350
      TabIndex        =   5
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Timer timRead 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   30
      Top             =   1650
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1635
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   825
      Width           =   2355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   4530
      TabIndex        =   3
      Top             =   210
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   4530
      TabIndex        =   4
      Top             =   870
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   4170
      TabIndex        =   7
      Top             =   -270
      Width           =   30
   End
   Begin VB.Label lbl新密码 
      AutoSize        =   -1  'True
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   810
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lbl新密码 
      AutoSize        =   -1  'True
      Caption         =   "新密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   555
      TabIndex        =   12
      Top             =   1350
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lbl姓名 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   810
      TabIndex        =   10
      Top             =   2790
      Width           =   510
   End
   Begin VB.Label lbl卡号 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   810
      TabIndex        =   9
      Top             =   2430
      Width           =   510
   End
   Begin VB.Label lblNote 
      Caption         =   "请在读卡器的绿灯亮了之后，输入密码。"
      Height          =   540
      Left            =   840
      TabIndex        =   6
      Top             =   165
      Width           =   3180
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   810
      TabIndex        =   8
      Top             =   885
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmIdentify中软.frx":030A
      Top             =   405
      Width           =   480
   End
   Begin VB.Label lbl背景 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   1  'Fixed Single
      Height          =   885
      Left            =   645
      TabIndex        =   11
      Top             =   2250
      Width           =   3345
   End
End
Attribute VB_Name = "frmIdentify中软"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'      1）读IC卡病人信息
Private Declare Function ReadICCard Lib "ICREAD.DLL" (iIC中软 As TIC中软) As Long
'      2）写IC卡病人信息
Private Declare Function WriteICCard Lib "ICWRITE.DLL" (iIC中软 As TIC中软) As Long

Private mIC中软 As TIC中软   '临时保存卡信息
Private mintTimes As Integer
Private mblnOK As Boolean
Private mbln判断在院 As Boolean      '是否需要对该病人在院与否进行判断

Private Sub chk离休_Click()
    txtPwd.Text = ""
    lbl卡号.Caption = "卡号："
    lbl姓名.Caption = "姓名："
    If chk离休.Value = 0 Then
        '普通病人
        txtPwd.MaxLength = Len(mIC中软.Password)
        txtPwd.PasswordChar = "*"
        lblPwd.Caption = "密码"
        lblNote.Caption = "请在读卡器的绿灯亮了之后，输入密码。"
        timRead.Enabled = True
        cmdOK.Default = True
    Else
        '离休病人
        txtPwd.MaxLength = 18
        txtPwd.PasswordChar = ""
        lblPwd.Caption = "身份证"
        lblNote.Caption = "请输入离休病人的身份证。"
        timRead.Enabled = False
        cmdOK.Default = False
    End If
    txtPwd.SetFocus
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    mintTimes = mintTimes + 1
    '暂时停止自动读取，以免与正常操作冲突
    timRead.Enabled = False
    If IsValid = False Then
        '恢复
        timRead.Enabled = True
        If mintTimes > 3 Then
            '密码错误次数太多
            Unload Me
        End If
        Exit Sub
    End If
    
    If txt新密码(0).Visible = True Then
        If SavePass() = False Then
            Exit Sub
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Private Function SavePass() As Boolean
'功能：修改用户IC卡密码
    Dim ic As TIC中软
    Dim lngReturn  As Long
    
    On Error GoTo errHandle
    ic = mIC中软
    ic.Password = txt新密码(0).Text
    MousePointer = vbHourglass
    lngReturn = WriteICCard(ic)
    MousePointer = vbDefault
    If lngReturn <> 0 Then
        '读取失败
        MsgBox 错误信息_中软(lngReturn), vbInformation, gstrSysName
        Exit Function
    End If
    mIC中软 = ic
    MsgBox "新密码保存成功。", vbInformation, gstrSysName
    SavePass = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MousePointer = vbDefault
End Function

Private Function IsValid() As Boolean
'功能：判断IC卡是否合法
    Dim rsTemp As New ADODB.Recordset
    Dim str有效期 As String
    Dim bln定点医疗 As Boolean
    Dim str参数值 As String
    Dim lngIndex As Long, lngCount As Long
    
    If ReadIC中软(True) = False Then
        '读卡失败
        Exit Function
    End If
    'H） 密码校验是否正确：验证IC卡中Password。（如果Password为9000不对密码进行验证）。
    If TruncZero(mIC中软.Password) <> "9000" Then
        If TruncZero(mIC中软.Password) <> txtPwd.Text Then
            MsgBox "密码输入错误。", vbInformation, gstrSysName
            txtPwd.Text = ""
            txtPwd.SetFocus
            Exit Function
        End If
    End If
    
    '进行合法性验证
    If txt新密码(0).Visible = False Then
        str有效期 = Get保险参数_中软(mIC中软.CenterCode, "有效期", True)
        bln定点医疗 = (Get保险参数_中软(mIC中软.CenterCode, "定点医疗机构", False) = "1")
        
        'B） 是否过有效期判断：即判断Center表中CenterCode等于IC卡中的CenterCode的记录的UseExpired字段信息是否小于当前日期。
        If IsDate(str有效期) = False Then
            MsgBox "请先从医保中心下载数据后再使用本功能。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        If CDate(str有效期) < CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")) Then
            MsgBox "病人所属医保中心已经过了有效期。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        'C） 判断个人账户是否出现负数。即判断IC卡中InPerAcc-OutPerAcc是否为负数。
        If mIC中软.InPerAcc - mIC中软.OutPerAcc < 0 Then
            MsgBox "病人个人账户已经出现负数。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        'D） 判断是否定点医疗机构：即判断Center表中CenterCode等于IC卡中的CenterCode的记录的IsAppoint字段信息，
        '    如果IsAppoint=1则是，IsAppoint=0则否。
        If bln定点医疗 = False Then
            MsgBox "本院不属于该病人的定点医疗机构。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        'E） 判断是否属于长驻外地职工：判断IC卡中DomainCode是否等于1
        If mIC中软.DomainCode = 1 Then
            MsgBox "该病人属于长驻外地职工。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        'F） 判断是否属于异地安置职工：判断IC卡中DomainCode是否等于2。
        If mIC中软.DomainCode = 2 Then
            MsgBox "该病人属于异地安置职工。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        'G） 判断职工是否在住院：判断IC卡中InpatientFlag。（住院结算不进行此判断）
        If mbln判断在院 = True Then
            If mIC中软.InpatientFlag = "1" Then
                MsgBox "该病人仍然在院。", vbInformation, gstrSysName
                mintTimes = 10 '直接退出当前窗口
                Exit Function
            End If
        End If
    Else
        '密码修改
        If txt新密码(0).Text <> txt新密码(1).Text Then
            txt新密码(0).Text = ""
            txt新密码(1).Text = ""
            txt新密码(0).SetFocus
            MsgBox "新密码与确认密码不相同。", vbInformation, gstrSysName
            Exit Function
        End If
        
        For lngIndex = 0 To 1
            If Len(txt新密码(lngIndex).Text) <> txt新密码(lngIndex).MaxLength Then
                txt新密码(lngIndex).Text = ""
                txt新密码(lngIndex).SetFocus
                MsgBox "密码长度不够。", vbInformation, gstrSysName
                Exit Function
            End If
            
            For lngCount = 1 To Len(txt新密码(lngIndex).Text)
                If InStr("0123456789", Mid(txt新密码(lngIndex).Text, lngCount, 1)) = 0 Then
                    txt新密码(lngIndex).Text = ""
                    txt新密码(lngIndex).SetFocus
                    MsgBox "密码只能由数字组成。", vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        Next
    End If
    IsValid = True
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If txt新密码(0).Visible = True Then Exit Sub
    
    If InStr("0123456789X", Chr(KeyAscii)) > 0 Then
        If Not ActiveControl Is txtPwd Then
            '直接进入密码输入框
            txtPwd.SetFocus
            DoEvents
            txtPwd.Text = Chr(KeyAscii)
            txtPwd.SelStart = Len(txtPwd.Text)
            txtPwd.SelLength = 0
        End If
    End If
End Sub

Private Sub timRead_Timer()
    Call ReadIC中软
End Sub

Private Sub txtPwd_GotFocus()
    zlControl.TxtSelAll txtPwd
End Sub

Public Function GetPatient(ByVal bln判断在院 As Boolean, Optional ByVal bln修改密码 As Boolean = False) As Boolean
    Dim lngIndex As Long
    
    mintTimes = 0
    mblnOK = False
    timRead.Enabled = True
    txtPwd.MaxLength = Len(mIC中软.Password)
    txt新密码(0).MaxLength = txtPwd.MaxLength
    txt新密码(1).MaxLength = txtPwd.MaxLength
    
    mbln判断在院 = bln判断在院
    '先预读一次
    Call ReadIC中软
    
    '根据是否修改密码，改变显示状态
    If bln修改密码 = False Then
        lbl背景.Top = lbl背景.Top - 900
        lbl卡号.Top = lbl卡号.Top - 900
        lbl姓名.Top = lbl姓名.Top - 900
        
        Me.Height = Me.Height - 900
    Else
        Me.Caption = "IC卡密码修改"
        cmdOK.Default = False
        chk离休.Visible = False
        
        For lngIndex = 0 To 1
            lbl新密码(lngIndex).Visible = True
            txt新密码(lngIndex).Visible = True
        Next
    End If
    
    frmIdentify中软.Show vbModal
    DoEvents
    '返回值
    If mblnOK = True Then
        gIC中软 = mIC中软
    End If
    GetPatient = mblnOK
    
End Function

Private Function ReadIC中软(Optional ByVal blnMessage As Boolean = False) As Boolean
'功能：读IC卡上的信息
    Dim lngReturn As Long
    
    If chk离休.Value = 0 Then
        lngReturn = ReadICCard(mIC中软)
    Else
        '从离休清单中读取病人情况，填入IC卡结构中
        If Get离休病人_中软(Trim(txtPwd.Text), mIC中软, False) = False Then
            Exit Function
        End If
    End If
    If lngReturn = 0 Then
        '读取成功
        lbl卡号.Caption = "卡号：" & TruncZero(mIC中软.Cardno)
        lbl姓名.Caption = "姓名：" & TruncZero(mIC中软.Name)
        
        ReadIC中软 = True
    Else
        '读取失败
        If blnMessage = True Then
            MsgBox 错误信息_中软(lngReturn), vbInformation, gstrSysName
        End If
        lbl卡号.Caption = "卡号："
        lbl姓名.Caption = "姓名："
    End If
End Function

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If ReadIC中软 = True Then
            If txt新密码(0).Visible = False Then
                cmdOK.SetFocus
            Else
                txt新密码(0).SetFocus
            End If
        End If
    End If
End Sub

Private Sub txt新密码_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt新密码(Index)
End Sub

Private Sub txt新密码_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
