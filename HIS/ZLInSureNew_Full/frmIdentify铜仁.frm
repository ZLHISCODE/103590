VERSION 5.00
Begin VB.Form frmIdentify铜仁 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   3300
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
   Icon            =   "frmIdentify铜仁.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   300
      Left            =   3660
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2190
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt病种 
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   1635
      MaxLength       =   8
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox txt新密码 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1635
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   9
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
      TabIndex        =   4
      Top             =   1290
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.CheckBox chk离休 
      Caption         =   "离休病人(&L)"
      Height          =   315
      Left            =   4350
      TabIndex        =   12
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
      TabIndex        =   2
      Top             =   825
      Width           =   2355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   4530
      TabIndex        =   10
      Top             =   210
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   4530
      TabIndex        =   11
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
      TabIndex        =   13
      Top             =   -270
      Width           =   30
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmIdentify铜仁.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lbl病种 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病种"
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
      TabIndex        =   5
      Top             =   2220
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lbl新密码 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lbl新密码 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   3
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   2430
      Width           =   510
   End
   Begin VB.Label lblNote 
      Caption         =   "请在读卡器的绿灯亮了之后，输入密码。"
      Height          =   540
      Left            =   840
      TabIndex        =   0
      Top             =   165
      Width           =   3180
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   1
      Top             =   885
      Width           =   510
   End
   Begin VB.Label lbl背景 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   1  'Fixed Single
      Height          =   885
      Left            =   645
      TabIndex        =   16
      Top             =   2280
      Width           =   3345
   End
End
Attribute VB_Name = "frmIdentify铜仁"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'      1）读IC卡病人信息
Private Declare Function ReadICCard Lib "ICREAD.DLL" (iIC铜仁 As TIC铜仁) As Long
'      2）写IC卡病人信息
Private Declare Function WriteICCard Lib "ICWRITE.DLL" (iIC铜仁 As TIC铜仁) As Long

Private mIC铜仁 As TIC铜仁   '临时保存卡信息

Private mintTimes As Integer
Private mblnOK As Boolean
Private mint场合 As Integer
Private mlng病种ID As Long
Private mstr病种编码 As String
Private mbln远程验证 As Boolean, mstr远程地址 As String
Private blnUpload As Boolean

Private Sub chk离休_Click()
    txtPwd.Text = ""
    lbl卡号.Caption = "卡号："
    lbl姓名.Caption = "姓名："
    
    Call SetFace
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
    mlng病种ID = Val(txt病种.Tag)
    
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
    Dim ic As TIC铜仁
    Dim lngReturn  As Long
    
    On Error GoTo errHandle
    ic = mIC铜仁
    ic.Password = txt新密码(0).Text
    MousePointer = vbHourglass
    lngReturn = WriteICCard(ic)
    MousePointer = vbDefault
    If lngReturn <> 0 Then
        '读取失败
        MsgBox 错误信息_铜仁(lngReturn), vbInformation, gstrSysName
        Exit Function
    End If
    mIC铜仁 = ic
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
    Dim dat有效期 As String
    Dim bln定点医疗 As Boolean
    Dim str参数值 As String, str病种类别 As String
    Dim lngIndex As Long, lngCount As Long
    
    If ReadIC铜仁(True) = False Then
        '读卡失败
        Exit Function
    End If
    'H） 密码校验是否正确：验证IC卡中Password。（如果Password为9000不对密码进行验证）。
    If TruncZero(mIC铜仁.Password) <> "9000" Then
        If mbln远程验证 = False Then
            If TruncZero(mIC铜仁.Password) <> txtPwd.Text Then
                MsgBox "密码输入错误。", vbInformation, gstrSysName
                txtPwd.Text = ""
                txtPwd.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '进行合法性验证
    If txt新密码(0).Visible = False Then
        If mint场合 = 1 And txt病种.Tag = "" Then
            MsgBox "入院必须选择病种。", vbInformation, gstrSysName
            Exit Function
        End If
        If mint场合 = 0 And txt病种.Tag <> "" Then
            '检查是否支付该报销
            gstrSQL = "SELECT A.类别 FROM 保险病种 A " & _
                      "  WHERE A.险类=81 AND A.编码='" & mstr病种编码 & "' and A.类别>'0'"
            If rsTemp.State = adStateOpen Then rsTemp.Close
            rsTemp.Open gstrSQL, gcn铜仁, adOpenStatic, adLockReadOnly
            If rsTemp.EOF = True Then
                MsgBox "在医保中心不能找到该病种。", vbInformation, gstrSysName
                Exit Function
            End If
            str病种类别 = rsTemp("类别")
        End If
        
        gstrSQL = "SELECT B.有效期,B.是否可用,A.运行模式,A.开展慢病报销,A.开展大病报销 " & _
                   " FROM 保险中心目录 A,保险主机 B " & _
                   " WHERE A.险类=" & TYPE_铜仁 & " AND A.编码='" & mIC铜仁.CenterCode & "' AND A.主机编码=B.编码 AND A.险类=B.险类 "
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcn铜仁, adOpenStatic, adLockReadOnly
        If rsTemp.EOF = False Then
            dat有效期 = Nvl(rsTemp("有效期"), Date)
            bln定点医疗 = Nvl(rsTemp("是否可用"), 0) And Nvl(rsTemp("运行模式"), 0)
        Else
            MsgBox "请先从医保中心下载数据后再使用本功能。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        If str病种类别 >= "1" And str病种类别 <= "5" Then
            If Nvl(rsTemp("开展慢病报销"), 0) <> 1 Then
                MsgBox "本院在病人所属中心未开展慢病报销。", vbInformation, gstrSysName
                Exit Function
            End If
        ElseIf str病种类别 >= "6" And str病种类别 <= "9" Then
            If Nvl(rsTemp("开展大病报销"), 0) <> 1 Then
                MsgBox "本院在病人所属中心未开展大病报销。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        'B） 是否过有效期判断：即判断Center表中CenterCode等于IC卡中的CenterCode的记录的UseExpired字段信息是否小于当前日期。
        If dat有效期 < CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")) Then
            MsgBox "病人所属医保中心已经过了有效期。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        'C） 判断个人账户是否出现负数。即判断IC卡中InPerAcc-OutPerAcc是否为负数。
        If mIC铜仁.InPerAcc - mIC铜仁.OutPerAcc < 0 Then
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
        If mIC铜仁.DomainCode = 1 Then
            MsgBox "该病人属于长驻外地职工。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        'F） 判断是否属于异地安置职工：判断IC卡中DomainCode是否等于2。
        If mIC铜仁.DomainCode = 2 Then
            MsgBox "该病人属于异地安置职工。", vbInformation, gstrSysName
            mintTimes = 10 '直接退出当前窗口
            Exit Function
        End If
        
        'G） 判断职工是否在住院：判断IC卡中InpatientFlag。（住院结算不进行此判断）
'        If mbln判断在院 = True Then
            If mIC铜仁.InpatientFlag = "1" Then
                MsgBox "该病人仍然在院。", vbInformation, gstrSysName
                mintTimes = 10 '直接退出当前窗口
                Exit Function
            End If
            
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 卡号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否已经在院", TYPE_铜仁, CStr(TrimStr(mIC铜仁.Cardno)))
            If Not rsTemp.EOF Then
                If rsTemp!状态 = 1 Then
                    MsgBox "该病人仍然在院。", vbInformation, gstrSysName
                    mintTimes = 10 '直接退出当前窗口
                    Exit Function
                End If
            End If
'        End If
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

Private Sub cmd病种_Click()
    Dim rs病种 As ADODB.Recordset
    
    '住院要选择普通病种
    '门诊选择慢特病
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=[1] And A.类别 IN (" & IIf(mint场合 = 0, "1,2)", "0)")
    Set rs病种 = New ADODB.Recordset
    Set rs病种 = zlDatabase.OpenSQLRecord(gstrSQL, "身份验证", TYPE_铜仁)
    If rs病种.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_铜仁, rs病种, "ID", "医保病种选择", "请选择医保病种：") = True Then
            txt病种.Text = rs病种("名称")
            txt病种.Tag = rs病种("ID")
            mstr病种编码 = rs病种("编码")
            txt病种.ForeColor = txtPwd.ForeColor
        End If
    End If
End Sub

Private Sub Form_Load()
    Shell "cmd /c route delete 0.0.0.0", vbNormal
    Shell "cmd /c route add 0.0.0.0 mask 0.0.0.0 192.168.0.1", vbNormal
End Sub

Private Sub timRead_Timer()
    If mbln远程验证 = False Then
        Call ReadIC铜仁
    End If
End Sub

Private Sub txtPwd_GotFocus()
    zlControl.TxtSelAll txtPwd
End Sub

Public Function GetPatient(ByVal int场合 As Integer, ByVal bln修改密码 As Boolean, 病种ID As Long) As Boolean
    Dim lngIndex As Long
    Dim bln远程验证 As Boolean, str远程地址 As String
    
    If Get保险参数_铜仁(bln远程验证, str远程地址) = False Then
        Exit Function
    End If
    If bln修改密码 = True And bln远程验证 = True Then
        MsgBox "现在采用远程身份验证，不能进行密码修改。", vbInformation, gstrSysName
        Exit Function
    End If
    mbln远程验证 = bln远程验证
    mstr远程地址 = str远程地址
    
    mintTimes = 0
    mblnOK = False
    timRead.Enabled = True
    txtPwd.MaxLength = Len(mIC铜仁.Password)
    txt新密码(0).MaxLength = txtPwd.MaxLength
    txt新密码(1).MaxLength = txtPwd.MaxLength
    
    mint场合 = int场合
    If int场合 = 0 Then
        '首先检查是否可以使用慢病
        
    End If
    
    '先预读一次
    blnUpload = False
    Call ReadIC铜仁
    
    '根据是否修改密码，改变显示状态
    If bln修改密码 = False Then
'        If int场合 = 0 Or int场合 = 1 Then
            '门诊与入院都要求输入病种
            lbl病种.Top = lbl新密码(0).Top
            txt病种.Top = txt新密码(0).Top
            cmd病种.Top = txt病种.Top + 30
            lbl病种.Visible = True
            txt病种.Visible = True
            cmd病种.Visible = True
'        End If
        lbl新密码(1).Caption = "密码"
        lbl新密码(1).Left = lblPwd.Left
        Call SetFace
    Else
        Me.Caption = "IC卡密码修改"
        cmdOK.Default = False
        chk离休.Visible = False
        
        For lngIndex = 0 To 1
            lbl新密码(lngIndex).Visible = True
            txt新密码(lngIndex).Visible = True
        Next
    End If
    
    frmIdentify铜仁.Show vbModal
    DoEvents
    '返回值
    If mblnOK = True Then
        病种ID = mlng病种ID
        gIC铜仁 = mIC铜仁
    End If
    GetPatient = mblnOK
    
End Function

Private Function ReadIC铜仁(Optional ByVal blnMessage As Boolean = False) As Boolean
'功能：读IC卡上的信息
    Dim lngReturn As Long
    
    If chk离休.Value = 0 Then
        If mbln远程验证 = False Then
            lngReturn = ReadICCard(mIC铜仁)
        Else
            '远程连接
            If Trim(txtPwd.Text) = "" Then
                If blnMessage = True Then MsgBox "请输入身份证号码。", vbInformation, gstrSysName
                Exit Function
            End If
            If blnUpload = False Then
                If frmSock铜仁.CommIC(mstr远程地址, True, IIf(mint场合 = 1, 1, 0), txtPwd.Text & "|" & txt新密码(1).Text) = False Then
                    Exit Function
                End If
                blnUpload = True
                mIC铜仁 = gIC铜仁Temp
            End If
        End If
    Else
        '从离休清单中读取病人情况，填入IC卡结构中
        If Get离休病人_铜仁(Trim(txtPwd.Text), mIC铜仁, False) = False Then
            If blnMessage = True Then MsgBox "未找到身份证号为 " & txtPwd.Text & " 的离休病人。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If lngReturn = 0 Then
        '读取成功
        lbl卡号.Caption = "卡号：" & TruncZero(mIC铜仁.Cardno)
        lbl姓名.Caption = "姓名：" & TruncZero(mIC铜仁.Name)
        
        ReadIC铜仁 = True
    Else
        '读取失败
        If blnMessage = True Then
            MsgBox 错误信息_铜仁(lngReturn), vbInformation, gstrSysName
        End If
        lbl卡号.Caption = "卡号："
        lbl姓名.Caption = "姓名："
    End If
End Function

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        blnUpload = False
        '当在身份证处按回车时,强制上传数据
        If ReadIC铜仁 = True Then
            If txt新密码(0).Visible = False Then
                txt病种.SetFocus
            Else
                txt新密码(0).SetFocus
            End If
        Else
            zlControl.TxtSelAll txtPwd
        End If
    End If
End Sub

Private Sub txt病种_Change()
    txt病种.Tag = ""
    txt病种.ForeColor = &HC0&
End Sub

Private Sub txt病种_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt病种.Text = "" Or txt病种.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt病种.Text
    gstrSQL = "Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特殊病','普通病') 类别 " & _
             "   FROM 保险病种 A WHERE A.险类=[1] And A.类别 IN ([2]) And (" & _
             " A.编码 like [3] || '%' or A.名称  like [3] || '%' or  A.简码  like [3] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_铜仁, IIf(mint场合 = 0, "1,2", "0"), strText)
    
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_铜仁, rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Text = rsTemp("名称")
        txt病种.Tag = rsTemp("ID")
        mstr病种编码 = rsTemp("编码")
        txt病种.ForeColor = txtPwd.ForeColor
        SendKeys "{TAB}"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt新密码_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt新密码(Index)
End Sub

Private Sub SetFace()
'功能：根据状态设置界面样式
    If chk离休.Value = 1 Or mbln远程验证 = True Then
        '离休病人
        txtPwd.MaxLength = 18
        txtPwd.PasswordChar = ""
        lblPwd.Caption = "身份证"
        lblNote.Caption = "请输入病人的身份证。"
        timRead.Enabled = False
        If chk离休.Value = 1 Then
            '离休不需要密码
            lbl新密码(1).Visible = False
            txt新密码(1).Visible = False
        Else
            '远程验证
            lbl新密码(1).Visible = True
            txt新密码(1).Visible = True
        End If
    Else
        '直接读IC卡
        txtPwd.MaxLength = Len(mIC铜仁.Password)
        txtPwd.PasswordChar = "*"
        lblPwd.Caption = "密码"
        lblNote.Caption = "请在读卡器的绿灯亮了之后，输入密码。"
        timRead.Enabled = True
    End If
End Sub

Private Sub txt新密码_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub


