VERSION 5.00
Begin VB.Form frmIdentify毕节 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmIdentify毕节.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   270
      Left            =   3510
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3390
      Width           =   285
   End
   Begin VB.TextBox txt病种 
      Height          =   300
      Left            =   1530
      TabIndex        =   19
      Top             =   3390
      Width           =   2265
   End
   Begin VB.CommandButton cmdChangePass 
      Caption         =   "改密码(&G)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4320
      TabIndex        =   24
      Top             =   3270
      Width           =   1100
   End
   Begin VB.TextBox txt统筹报销累计 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3030
      Width           =   2265
   End
   Begin VB.TextBox txt住院次数 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2670
      Width           =   2265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4320
      TabIndex        =   22
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4320
      TabIndex        =   23
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   4050
      TabIndex        =   21
      Top             =   -120
      Width           =   45
   End
   Begin VB.TextBox txt帐户余额 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2310
      Width           =   2265
   End
   Begin VB.TextBox txt参加工作日期 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1950
      Width           =   2265
   End
   Begin VB.TextBox txt出生日期 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1590
      Width           =   2265
   End
   Begin VB.TextBox txt性别 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1230
      Width           =   1245
   End
   Begin VB.TextBox txt姓名 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   870
      Width           =   2265
   End
   Begin VB.TextBox txt密码 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1530
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   150
      Width           =   2265
   End
   Begin VB.TextBox txt社会保障号 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   510
      Width           =   2265
   End
   Begin VB.Label lbl病种 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "病种(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   810
      TabIndex        =   18
      Top             =   3450
      Width           =   630
   End
   Begin VB.Label lbl统筹报销累计 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "统筹报销累计(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   90
      TabIndex        =   16
      Top             =   3090
      Width           =   1350
   End
   Begin VB.Label lbl住院次数 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "住院次数(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   14
      Top             =   2730
      Width           =   990
   End
   Begin VB.Label lbl帐户余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "帐户余额(&Y)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   12
      Top             =   2370
      Width           =   990
   End
   Begin VB.Label lbl参加工作日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "工作日期(&J)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   10
      Top             =   2010
      Width           =   990
   End
   Begin VB.Label lbl出生日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   8
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label lbl性别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   810
      TabIndex        =   6
      Top             =   1290
      Width           =   630
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   810
      TabIndex        =   4
      Top             =   930
      Width           =   630
   End
   Begin VB.Label lbl密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IC卡密码(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   210
      Width           =   990
   End
   Begin VB.Label lbl社会保障号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "社会保障号(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   2
      Top             =   570
      Width           =   1170
   End
End
Attribute VB_Name = "frmIdentify毕节"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long
Private mbytType As Byte
Private mstrReturn As String
Private mstrICData As String
'只有使用了IC卡密码的才进行密码校验,并允许修改密码
'需要对冻结状态进行检查,如果冻结则不允许就诊登记
'如果有效卡号与中心一致,才允许使用,否则视为废卡

Public Function GetIdentify(ByVal bytType As Byte, Optional lng病人ID As Long) As String
    mlng病人ID = lng病人ID
    mbytType = bytType
    mstrReturn = ""
    mstrICData = ""
    Me.Show 1
    GetIdentify = mstrReturn
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangePass_Click()
    Dim strOldPass As String, strNewPass As String
    Dim strCardData As String
    '因程序上已控制，只有正确输入密码才能修改密码，因此，对输入的旧密码不作判断
    strNewPass = frm修改密码.ChangePassword(Me.txt密码.Text, strOldPass)
    If strOldPass = strNewPass Then Exit Sub
    
    If strOldPass <> IC_Data_毕节.个人IC卡密码 Then
        MsgBox "输入的旧密码与卡内密码不符！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    IC_Data_毕节.个人IC卡密码 = strNewPass
    Call 数据转换_毕节(strCardData, False)
    Call gobjCenter.IC_ChangePass(strCardData)
End Sub

Private Sub cmdOK_Click()
    Dim lng年龄 As Long
    Dim strIdentify As String, strAddition As String
    Dim rsTmp As New ADODB.Recordset
    '检查是否在院，如果在院，则禁止就诊
    gstrSQL = "Select Nvl(当前状态,0) 状态 From 保险帐户 Where 险类=[1] And 医保号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_毕节, CStr(txt社会保障号.Text))
    If rsTmp.RecordCount = 1 Then
        If rsTmp!状态 = 1 Then
            MsgBox "当前病人目前正在院治疗，不允许在院治疗期间再次就诊！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '如果是入院，则必须选择病种
    If mbytType = 1 Then
        If Val(txt病种.Tag) = 0 Then
            MsgBox "请选择入院病种！", vbInformation, gstrSysName
            txt病种.SetFocus
            Exit Sub
        End If
    End If
    
    lng年龄 = GetAge(Format(zlDatabase.Currentdate, "yyyy-MM-dd"), Me.txt出生日期.Text)
    '产生病人信息
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    strIdentify = IC_Data_毕节.有效卡号                         '0卡号
    strIdentify = strIdentify & ";" & txt社会保障号.Text        '1医保号
    strIdentify = strIdentify & ";"                             '2密码
    strIdentify = strIdentify & ";" & txt姓名.Text              '3姓名
    strIdentify = strIdentify & ";" & txt性别.Text              '4性别
    strIdentify = strIdentify & ";" & txt出生日期.Text          '5出生日期
    strIdentify = strIdentify & ";"                             '6身份证
    strIdentify = strIdentify & ";" & IC_Data_毕节.单位代码     '7.单位名称(编码)
    strAddition = ";0"                                          '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";"                             '10人员身份
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & Val(txt病种.Tag)         '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '18帐户增加累计
    strAddition = strAddition & ";0"                            '19帐户支出累计
    strAddition = strAddition & ";" & Val(txt统筹报销累计.Text) '20上年工资总额
    strAddition = strAddition & ";" & Val(txt住院次数.Text)     '21住院次数累计

    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_毕节)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    Else
        Exit Sub
    End If
    
    '弹卡（由于门诊结算时要使用卡，所以不弹卡，而住院则可以弹卡）
    If mbytType = 1 Then Call IC_End(True)
    
    Unload Me
End Sub

Private Sub cmd病种_Click()
    Dim blnReturn As Boolean
    Dim rsTmp As New ADODB.Recordset
        
    gstrSQL = " Select ID,病种代码 As 编码,病种名称,中医名称,病种类别,个人自付比例,个人起付金额 " & _
              " From 病种目录表"
    With rsTmp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
    End With
    
    blnReturn = frmListSel.ShowSelect(TYPE_毕节, rsTmp, "ID", "医保病种选择", "请选择医保病种：")
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt病种.Text = lbl病种.Tag
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Tag = rsTmp!ID
        txt病种.Text = "(" & rsTmp!编码 & ")" & rsTmp!病种名称
        lbl病种.Tag = txt病种.Text '用于恢复显示
    End If
End Sub

Private Sub Form_Load()
    txt密码.Locked = Not gCominfo_毕节.blnICPassVerify
    
    Me.txt病种.Enabled = (mbytType = 1)
    Me.cmd病种.Enabled = (mbytType = 1)
End Sub

Private Sub txt密码_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim dbl余额 As Double
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Me.cmdOK.Enabled = False
    cmdChangePass.Enabled = False
    
    If Not gobjCenter.IC_ReadCard(mstrICData) Then Exit Sub
    Call 数据转换_毕节(mstrICData, True)
    
    If gCominfo_毕节.blnICPassVerify Then
        If txt密码.Text <> IC_Data_毕节.个人IC卡密码 Then
            MsgBox "IC卡密码错误，请重新输入！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '以下数据需要从中间库中获取
    gstrSQL = "Select 住院次数,帐户冻结,冻结原因,有效卡号,当年住院费用,冻结时间,冻结说明,余额 " & _
        " From 个人帐户余额表 Where 社会保障号='" & IC_Data_毕节.社会保障号 & "'"
    If gCominfo_毕节.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTmp) Then
            Call IC_End(True)
            Call gobjCenter.CloseConnector
            Exit Sub
        End If
    Else
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.Open gstrSQL, gcnGYBJYB
    End If
    
    With rsTmp
        If .RecordCount = 0 Then
            Call IC_End(True)
            MsgBox "没发现该病人的有效记录，请与中心联系！", vbInformation, gstrSysName
            Exit Sub
        End If
        If Nvl(!帐户冻结, "否") = "是" Then
            Call IC_End(True)
            MsgBox "该病人的帐户已经被冻结，只能以普通病人的身份办理！" & vbCrLf & "冻结原因：" & Nvl(!冻结原因) & vbCrLf & "冻结说明：" & Nvl(!冻结说明) & vbCrLf & "冻结时间：" & Nvl(!冻结时间), vbInformation, gstrSysName
            Exit Sub
        End If
        If Nvl(IC_Data_毕节.有效卡号, 0) <> Nvl(!有效卡号, 0) Then
            Call IC_End(True)
            MsgBox "当前的IC卡片是一张无效的卡！", vbInformation, gstrSysName
            Exit Sub
        End If
        dbl余额 = Nvl(!余额, 0)
        Me.txt住院次数.Text = Format(Nvl(!住院次数, 0), "#####0;-#####0; ;")
        Me.txt统筹报销累计.Text = Format(Nvl(!当年住院费用, 0), "#####0.00;-#####0.00; ;")
    End With
    
    If gCominfo_毕节.blnOnLine Then gobjCenter.CloseConnector
    
    '将IC卡数据显示出来
    Me.txt社会保障号.Text = IC_Data_毕节.社会保障号
    Me.txt姓名.Text = IC_Data_毕节.姓名
    Me.txt性别.Text = IC_Data_毕节.性别
    Me.txt出生日期.Text = Replace(IC_Data_毕节.出生日期, ".", "-")
    Me.txt参加工作日期.Text = Replace(IC_Data_毕节.参加工作日期, ".", "-")
    Me.txt帐户余额.Text = Format(IC_Data_毕节.个人帐户余额, "#####0.00;-#####0.00; ;")
    
    '如果是脱机系统，且当天消费过，以卡内余额为准，否则以中心为准
    If gCominfo_毕节.blnOnLine Or (gCominfo_毕节.blnOnLine = False And Format(zlDatabase.Currentdate, "yyyy.MM.dd") <> IC_Data_毕节.最后就诊日期) Then
        Me.txt帐户余额.Text = Format(dbl余额, "#####0.00;-#####0.00; ;")
    End If
    IC_Data_毕节.个人帐户余额 = Val(Me.txt帐户余额.Text)
    
    cmdOK.Enabled = True
    cmdChangePass.Enabled = gCominfo_毕节.blnICPassVerify
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt病种_GotFocus()
    Call zlControl.TxtSelAll(txt病种)
End Sub

Private Sub txt病种_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt病种.Text = "" And txt病种.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = UCase(txt病种.Text)
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    gstrSQL = " Select ID,病种代码 As 编码,病种名称,中医名称,病种类别,个人自付比例,个人起付金额 " & _
              " From 病种目录表 A" & _
              " Where (" & zlCommFun.GetLike("A", "病种代码", strText) & " or " & zlCommFun.GetLike("A", "病种名称", strText) & " or zlspellcode(病种名称) Like '" & strText & "%')"
    With rsTmp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
    End With
    
    If rsTmp.RecordCount = 0 Then
        MsgBox "不存在该病种，请重新输入！", vbInformation, gstrSysName
        txt病种.Text = lbl病种.Tag
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '出现选择器
        If rsTmp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_毕节, rsTmp, "ID", "医保病种选择", "请选择医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt病种.Text = lbl病种.Tag
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Tag = rsTmp!ID
        txt病种.Text = "(" & rsTmp!编码 & ")" & rsTmp!病种名称
        lbl病种.Tag = txt病种.Text '用于恢复显示
    End If
    
    Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub IC_End(Optional ByVal blnPull As Boolean = False)
    '在打开IC设备后，如果出错，是否仅仅弹卡，否则就在弹卡后关闭端口
    Call gobjCenter.IC_PullCard
    If blnPull Then Exit Sub
    
    Call gobjCenter.IC_CloseDevice
End Sub
