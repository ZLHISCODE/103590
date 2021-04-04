VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentify眉山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人档案"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmIdentify眉山.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdChange 
      Caption         =   "修改密码(&M)"
      Height          =   405
      Left            =   330
      TabIndex        =   34
      Top             =   3390
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fra基本 
      Caption         =   "病人基本信息"
      Height          =   3135
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6795
      Begin VB.TextBox txt帐户余额 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2670
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   2490
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt卡号 
         Height          =   300
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox txt密码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt人员身份 
         Height          =   300
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   11
         Top             =   1905
         Width           =   1455
      End
      Begin VB.TextBox txt住院次数 
         Height          =   300
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Index           =   1
         Left            =   6240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2700
         Width           =   255
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1125
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Index           =   2
         Left            =   6240
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1935
         Width           =   255
      End
      Begin VB.TextBox txt退休证号 
         Height          =   300
         Left            =   4440
         MaxLength       =   26
         TabIndex        =   28
         Top             =   2280
         Width           =   2085
      End
      Begin VB.ComboBox cmb中心 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   330
         Width           =   2085
      End
      Begin VB.ComboBox cmb性别 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1125
         Width           =   2085
      End
      Begin VB.TextBox txt身份证号 
         Height          =   300
         Left            =   4440
         MaxLength       =   18
         TabIndex        =   23
         Top             =   1515
         Width           =   2085
      End
      Begin VB.ComboBox Cbo当前状态 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp生日 
         Height          =   300
         Left            =   1320
         TabIndex        =   9
         Top             =   1515
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   87031811
         CurrentDate     =   36526
      End
      Begin VB.TextBox txt病种 
         Height          =   300
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2670
         Width           =   2085
      End
      Begin VB.TextBox txt单位编码 
         Height          =   300
         Left            =   4440
         MaxLength       =   8
         TabIndex        =   25
         Top             =   1905
         Width           =   2085
      End
      Begin VB.Label lbl帐户余额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "帐户余额(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   2730
         Width           =   990
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "卡号(&D)"
         Height          =   180
         Left            =   600
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl密码 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Left            =   600
         TabIndex        =   4
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl退休证号 
         AutoSize        =   -1  'True
         Caption         =   "退休证号(&Z)"
         Height          =   180
         Left            =   3360
         TabIndex        =   27
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl人员身份 
         AutoSize        =   -1  'True
         Caption         =   "人员身份(&E)"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl单位编码 
         AutoSize        =   -1  'True
         Caption         =   "单位编码(&U)"
         Height          =   180
         Left            =   3360
         TabIndex        =   24
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl病种 
         AutoSize        =   -1  'True
         Caption         =   "病种(&F)"
         Height          =   180
         Left            =   3720
         TabIndex        =   29
         Top             =   2730
         Width           =   630
      End
      Begin VB.Label lbl住院次数 
         AutoSize        =   -1  'True
         Caption         =   "住院次数(&S)"
         Height          =   180
         Left            =   3360
         TabIndex        =   18
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "姓名(&N)"
         Height          =   180
         Left            =   600
         TabIndex        =   6
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         Caption         =   "身份证号(&I)"
         Height          =   180
         Left            =   3360
         TabIndex        =   22
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl医保中心 
         AutoSize        =   -1  'True
         Caption         =   "医保中心(&R)"
         Height          =   180
         Left            =   3360
         TabIndex        =   16
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         Caption         =   "性别(&X)"
         Height          =   180
         Left            =   3720
         TabIndex        =   20
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         Caption         =   "出生日期(&B)"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl当前状态 
         AutoSize        =   -1  'True
         Caption         =   "人员类别(&K)"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   2340
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5760
      TabIndex        =   33
      Top             =   3450
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4350
      TabIndex        =   32
      Top             =   3450
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify眉山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _

Private Enum 选择Enum
    Select卡号 = 0
    Select病种 = 1
    Select单位 = 2
End Enum

Dim mstrIdentify As String
Dim mbytType As Byte        '0-门诊;1-住院;2-不分门诊与住院;3-补卡
Dim strNewPass As String
Dim mlng病人ID As Long

Public Function ShowCard(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：返回医保病人的身份信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型
    Dim rsTemp As New ADODB.Recordset
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrIdentify = ""
    
    cmb性别.Clear
    gstrSQL = "select 编码,名称 from 性别 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cmb性别.AddItem rsTemp("编码") & "." & rsTemp("名称")
        rsTemp.MoveNext
    Loop
    
    cmb中心.Clear
    gstrSQL = "select A.具有中心,B.序号,B.编码,B.名称 from 保险类别 A,保险中心目录 B where A.序号=[1] and A.序号=b.险类"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_四川眉山)
    
    If rsTemp("具有中心") = 0 Then
        lbl医保中心.Visible = False
        cmb中心.Visible = False
        cmb中心.AddItem "1.中心" '单中心
    End If
    Do Until rsTemp.EOF
        cmb中心.AddItem rsTemp("编码") & "." & rsTemp("名称")
        cmb中心.ItemData(cmb中心.NewIndex) = rsTemp("序号")
        rsTemp.MoveNext
    Loop
    cmb中心.ListIndex = 0
    
    '1-在职;2-退休;3-离休
    gstrSQL = "Select 序号,名称 From 保险人群 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_四川眉山)
    Cbo当前状态.Clear
    Do While Not rsTemp.EOF
        Cbo当前状态.AddItem rsTemp!名称
        Cbo当前状态.ItemData(Cbo当前状态.NewIndex) = rsTemp!序号
        rsTemp.MoveNext
    Loop
    Cbo当前状态.ListIndex = 0
    cmb性别.ListIndex = 0
        
    dtp生日.MaxDate = zlDatabase.Currentdate
    Call SetEnable(mbytType <> 2)
    txt帐户余额.Enabled = (mlng病人ID = 0 And mbytType = 2)
    If mlng病人ID <> 0 Then
        txt卡号 = "A" & mlng病人ID
        Call txt卡号_KeyPress(vbKeyReturn)
    End If
    frmIdentify眉山.Show vbModal
    ShowCard = mstrIdentify
End Function

Private Sub SetEnable(ByVal blnEnable As Boolean)
    cmdChange.Visible = blnEnable
    txt姓名.Enabled = Not blnEnable
    dtp生日.Enabled = Not blnEnable
    txt人员身份.Enabled = Not blnEnable
    Cbo当前状态.Enabled = Not blnEnable
    cmb中心.Enabled = Not blnEnable
    txt住院次数.Enabled = Not blnEnable
    cmb性别.Enabled = Not blnEnable
    txt身份证号.Enabled = Not blnEnable
    txt单位编码.Enabled = Not blnEnable
    txt退休证号.Enabled = Not blnEnable
    Me.cmdSelect(Select单位).Enabled = Not blnEnable
    
    txt密码.PasswordChar = IIf(blnEnable = False, "", "*")
End Sub

Private Sub Cbo当前状态_Click()
    txt退休证号.Enabled = (Cbo当前状态.ListIndex = 1 Or Cbo当前状态.ListIndex = 2)
End Sub

Private Sub cmb中心_Click()
    Dim lng卡号长度 As Long, lng退休证长度 As Long
    Dim rsTemp As New ADODB.Recordset
    
    '缺省值
    lng卡号长度 = 20
    lng退休证长度 = 26
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1] and 中心=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_四川眉山, CInt(cmb中心.ItemData(cmb中心.ListIndex)))
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "卡号长度"
                If IsNull(rsTemp("参数值")) = False Then lng卡号长度 = Val(rsTemp("参数值"))
            Case "退休证长度"
                If IsNull(rsTemp("参数值")) = False Then lng退休证长度 = Val(rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    
    txt卡号.MaxLength = lng卡号长度
    txt退休证号.MaxLength = lng退休证长度
End Sub

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdChange_Click()
    If IsValid() = False Then Exit Sub
    With frm修改密码
        strNewPass = .ChangePassword(txt密码.Text)
    End With
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strIdentify As String, strAddition As String
    Dim lng变动ID As Long, lng病人ID As Long, lng中心 As Long
    Dim blnTrans As Boolean
    Dim cur统筹累计 As Currency
    Dim lng年度 As Long
    Dim str比较串 As String
    
    '首先验数据的正确性
    If IsValid() = False Then
        Exit Sub
    End If
    
    '得到中心序号
    If cmb中心.Visible = False Then
        lng中心 = 0
    Else
        If cmb中心.ListIndex < 0 Then
            MsgBox "请选择病人所属医保中心。", vbInformation, gstrSysName
            cmb中心.SetFocus
            Exit Sub
        End If
        lng中心 = cmb中心.ItemData(cmb中心.ListIndex)
    End If
    
    '检查病人状态
    gstrSQL = "select nvl(当前状态,0) as 状态,灰度级,备注 from 保险帐户 where 险类=[1] and 中心=[2] and 医保号=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_四川眉山, lng中心, CStr(Trim(txt卡号.Text)))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("状态") > 0 Then
            MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
            Exit Sub
        End If
        Select Case Nvl(rsTemp!灰度级, 0)
        Case 1
            MsgBox "该医保卡已经封锁，不能使用！" & IIf(Nvl(rsTemp!备注) <> "", "（" & rsTemp!备注 & "）", ""), vbInformation, gstrSysName
            Exit Sub
        Case 9
            MsgBox "该医保卡已经撤销，不能使用！", vbInformation, gstrSysName
            Exit Sub
        End Select
    End If
    
    '检查数据库中的数据是否正确
    If Not 检查帐户信息_米易(txt卡号.Text) Then Exit Sub
    
    If strNewPass = "" Then strNewPass = Trim(txt密码.Text)
    '读取本年度统筹累计
    lng年度 = Format(zlDatabase.Currentdate, "yyyy")
    gstrSQL = " Select Nvl(进入统筹累计,0) 统筹累计 From 帐户年度信息 " & _
              " Where 年度=" & lng年度 & " And 病人ID =" & _
              "     (Select 病人ID From 保险帐户 where 险类=[1] and 中心=[2] and 医保号=[3])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取年度统筹累计", TYPE_四川眉山, lng中心, CStr(Trim(txt卡号.Text)))
    If Not rsTemp.EOF Then
        cur统筹累计 = rsTemp!统筹累计
    Else
        cur统筹累计 = 0
    End If
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型
    strIdentify = Trim(txt卡号.Text)                       '0卡号
    strIdentify = strIdentify & ";" & Trim(txt卡号.Text)   '1医保号
    strIdentify = strIdentify & ";" & strNewPass   '2密码
    strIdentify = strIdentify & ";" & Trim(txt姓名.Text)   '3姓名
    strIdentify = strIdentify & ";" & Replace(GetTextFromCombo(cmb性别, True), "'", "") '4性别
    strIdentify = strIdentify & ";" & Format(dtp生日.Value, "yyyy-MM-dd") '5出生日期
    strIdentify = strIdentify & ";" & Trim(txt身份证号.Text)    '6身份证
    strIdentify = strIdentify & ";" & Trim(txt单位编码.Text) & "(" & Trim(txt单位编码.Text) & ")"  '7.单位名称(编码)
    strAddition = ";" & lng中心                                 '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";" & Trim(txt人员身份.Text)       '10人员身份
    strAddition = strAddition & ";" & IIf(txt帐户余额.Enabled, Val(txt帐户余额.Text), "NULL")  '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & txt病种.Tag     '13病种ID
    strAddition = strAddition & ";" & Cbo当前状态.ItemData(Cbo当前状态.ListIndex) '14在职(1,2,3)
    strAddition = strAddition & ";" & Trim(txt退休证号.Text) '15退休证号
    strAddition = strAddition & ";" & DateDiff("yyyy", dtp生日.Value, dtp生日.MaxDate) '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)       '18帐户增加累计
    strAddition = strAddition & ";0"      '19帐户支出累计
    strAddition = strAddition & ";" & cur统筹累计      '20进入统筹累计
    strAddition = strAddition & ";0"       '21统筹报销累计
    strAddition = strAddition & ";" & txt住院次数.Text      '22住院次数累计
    strAddition = strAddition & ";"                                             '23就诊类型 (1、急诊门诊)
    
    gcnOracle.BeginTrans
    blnTrans = True
    lng病人ID = BuildPatiInfo_眉山(mbytType, strIdentify & strAddition, mlng病人ID)
    '返回格式:中间插入病人ID
    If lng病人ID > 0 Then
        mstrIdentify = strIdentify & ";" & lng病人ID & strAddition
    End If
    
    '如果是新建档案，则要插入帐户变动记录
    If txt帐户余额.Enabled Then
        Call 检查帐户信息_米易(txt卡号.Text, True)
        lng变动ID = zlDatabase.GetNextID("帐户变动记录")
        gstrSQL = "ZL_帐户变动记录_INSERT(" & _
                 lng变动ID & "," & TYPE_四川眉山 & ",1," & lng病人ID & "," & _
                 Val(txt帐户余额.Text) & ",'" & gstrUserName & "','创建医保病人档案时录入的初始值',1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    '更新个人帐户中的信息
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_四川眉山 & ",'密码','''" & strNewPass & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新密码")
    If mbytType = 4 Then
        Call 检查帐户信息_米易(txt卡号.Text, True)
    End If
    
    gcnOracle.CommitTrans
    
    '打印卡片
    If txt帐户余额.Enabled Or mbytType = 4 Then
        Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1604", Me, "险类=" & 25, "医保号=" & txt卡号, 2)
    End If
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim rsTemp As ADODB.Recordset
    
    Select Case Index
        Case Select卡号
            gstrSQL = " Select A.病人ID as ID,A.卡号,A.医保号,'******' 密码,B.姓名,B.性别,B.出生日期,B.身份证号,C.序号 as 中心ID " & _
                    " ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,A.在职 as 在职ID,A.退休证号,A.帐户余额" & _
                    " From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D" & _
                    "  where A.病人ID=B.病人ID And Nvl(A.灰度级,0)<>9 And A.险类=" & TYPE_四川眉山 & _
                    "  and A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+)"
            
            Call Get帐户情况
            zlControl.TxtSelAll txt卡号
            If txt卡号.Enabled Then txt卡号.SetFocus
        Case Select单位
            Set rsTemp = frmPubSel.ShowSelect(Me, _
                    " Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                    " Start With 上级ID is NULL Connect by Prior ID=上级ID", _
                    2, "工作单位", , txt单位编码.Text)
            
            If Not rsTemp Is Nothing Then
                txt单位编码.Text = rsTemp("编码")
                zlControl.TxtSelAll txt单位编码
            End If
            txt单位编码.SetFocus
        Case Select病种
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                    " From 保险病种 A where A.险类=" & TYPE_四川眉山
            
            Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txt病种.Text)
            If Not rsTemp Is Nothing Then
                txt病种.Text = rsTemp("名称")
                txt病种.Tag = rsTemp("ID")
                zlControl.TxtSelAll txt病种
            End If
            txt病种.SetFocus
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt病种_GotFocus()
    zlControl.TxtSelAll txt病种
End Sub

Private Sub txt病种_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt病种.Text = ""
        txt病种.Tag = ""
    End If
End Sub

Private Sub txt单位编码_GotFocus()
    zlControl.TxtSelAll txt单位编码
End Sub

Private Sub txt卡号_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt卡号
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    Dim strCode As String
    Dim str条件 As String
    Dim rsTemp As New ADODB.Recordset
    
    If Len(txt卡号.Text) = txt卡号.MaxLength Or KeyAscii = vbKeyReturn Then
        strCode = UCase(Replace(Trim(txt卡号.Text), "'", ""))
        If strCode = "" Then Exit Sub
        
        If IsNumeric(Mid(strCode, 1, Len(strCode) - 1)) Then '刷卡
            str条件 = " and A.卡号='" & strCode & "'"
        ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
            str条件 = " and A.病人ID=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号(对住(过)院的病人)
            str条件 = " and B.住院号=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '门诊号(仅对门诊病人)
            str条件 = " and B.门诊号=" & Mid(strCode, 2)
        Else '当作姓名
            str条件 = " and A.卡号='" & strCode & "'"
        End If
    
        gstrSQL = " Select A.病人ID as ID,A.卡号,A.医保号,'******' 密码,B.姓名,B.性别,B.出生日期,B.身份证号,C.序号 as 中心ID " & _
                " ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,A.在职 as 在职ID,A.退休证号,A.帐户余额" & _
                " From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D" & _
                "  where A.病人ID=B.病人ID And Nvl(A.灰度级,0)<>9 And A.险类=" & TYPE_四川眉山 & _
                "  and A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+)" & str条件
        
        Call Get帐户情况
    End If
End Sub

Private Sub txt卡号_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt人员身份_GotFocus()
    zlControl.TxtSelAll txt人员身份
End Sub

Private Sub txt身份证号_GotFocus()
    zlControl.TxtSelAll txt身份证号
End Sub

Private Sub txt退休证号_GotFocus()
    zlControl.TxtSelAll txt卡号
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub txt密码_GotFocus()
    zlControl.TxtSelAll txt密码
End Sub

Private Sub txt帐户余额_LostFocus()
    txt帐户余额 = Format(txt帐户余额, "#####0.00;-#####0.00; ;")
End Sub

Private Sub txt住院次数_GotFocus()
    zlControl.TxtSelAll txt住院次数
End Sub

Private Sub Get帐户情况()
'从已经存在的记录中读出帐户信息
    Dim rs帐户 As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    Set rs帐户 = frmPubSel.ShowSelect(Me, gstrSQL, 0, "保险帐户", , txt卡号.Text, "", False, True)
    If Not rs帐户 Is Nothing Then
        If mbytType = 2 Or mbytType = 4 Then txt卡号.Enabled = False
        txt卡号.Text = rs帐户("卡号")
        '其它可用的数据
        If mbytType = 2 Or mbytType = 4 Then
            gstrSQL = "Select 密码 From 保险帐户 Where 险类=[1] And 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_四川眉山, CStr(txt卡号.Text))
            
            txt密码.Text = Nvl(rsTemp!密码, "")
            txt密码.PasswordChar = "*"
            If mbytType = 4 Then
                If IsNumeric(Right(txt卡号.Text, 1)) Then
                    txt卡号.Text = txt卡号.Text & "A"
                Else
                    txt卡号.Text = Mid(txt卡号.Text, 1, Len(txt卡号.Text) - 1) & Chr(asc(Right(txt卡号.Text, 1)) + 1)
                End If
            End If
        End If
        txt姓名.Text = IIf(IsNull(rs帐户("姓名")), "", rs帐户("姓名"))
        txt身份证号.Text = IIf(IsNull(rs帐户("身份证号")), "", rs帐户("身份证号"))
        txt人员身份.Text = IIf(IsNull(rs帐户("人员身份")), "", rs帐户("人员身份"))
        txt单位编码.Text = IIf(IsNull(rs帐户("单位编码")), "", rs帐户("单位编码"))
        txt病种.Text = IIf(IsNull(rs帐户("病种")), "", rs帐户("病种"))
        txt病种.Tag = IIf(IsNull(rs帐户("病种ID")), "", rs帐户("病种ID"))
        
        Call SetComboByText(cmb性别, IIf(IsNull(rs帐户("性别")), "", rs帐户("性别")), True)
        Cbo当前状态.ListIndex = rs帐户("在职ID") - 1
        txt退休证号.Text = ""
        txt退休证号.Text = IIf(IsNull(rs帐户("退休证号")), "", rs帐户("退休证号"))
        If IsNull(rs帐户("出生日期")) = False Then
            dtp生日.Value = rs帐户("出生日期")
        End If
        
        For lngIndex = 0 To cmb中心.ListCount - 1
            If cmb中心.ItemData(lngIndex) = rs帐户("中心ID") Then
                cmb中心.ListIndex = lngIndex
                Exit For
            End If
        Next
        txt帐户余额 = Format(rs帐户!帐户余额, "#####0.00;-#####0.00; ;")
        txt帐户余额.Enabled = False
        
        '再读出帐户年度信息
        gstrSQL = "select * from 帐户年度信息 where 险类=[1] and 病人ID=[2] and 年度=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_四川眉山, CLng(rs帐户("ID")), Year(dtp生日.MaxDate))
        
        If rsTemp.EOF = False Then
            '设置帐户情况
            txt住院次数.Text = Nvl(rsTemp("住院次数累计"), 0) & "/" & Nvl(rsTemp("外院住院次数"), 0)
        Else
            txt住院次数.Text = "0/0"
        End If
    End If
End Sub

Private Function IsValid() As Boolean
'功能：检查数据的正确性
    Dim lngIndex As Long
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txt卡号) = "" Then
        MsgBox "卡号不能为空！", vbInformation, gstrSysName
        If txt卡号.Enabled Then txt卡号.SetFocus
        Exit Function
    End If
    If Trim(txt姓名) = "" Then
        MsgBox "姓名不能为空！", vbInformation, gstrSysName
        If txt姓名.Enabled Then txt姓名.SetFocus
        Exit Function
    End If
    If Trim(txt帐户余额) <> "" Then
        If Not IsNumeric(txt帐户余额) Then
            MsgBox "帐户余额中含有非法字符！", vbInformation, gstrSysName
            If txt帐户余额.Enabled Then txt帐户余额.SetFocus
            Exit Function
        End If
        If Val(txt帐户余额.Text) < 0 Then
            MsgBox "帐户余额不能小于零！", vbInformation, gstrSysName
            If txt帐户余额.Enabled Then txt帐户余额.SetFocus
            Exit Function
        End If
    End If
    If cmb中心.ListIndex < 0 Then
        MsgBox "必须要选择一个医保中心！", vbInformation, gstrSysName
        If cmb中心.Enabled Then cmb中心.SetFocus
        Exit Function
    End If
    If UBound(Split(txt住院次数.Text, "/")) <> 1 Then
        MsgBox "请输入本院住院次数及外院住院次数！（格式：本院/外院。如：1/1）", vbInformation, gstrSysName
        txt住院次数.SetFocus
        Exit Function
    End If
    If Trim(txt身份证号.Text) <> "" Then
        If Not IsNumeric(txt身份证号) Then
            MsgBox "身份证号中含有非法字符！", vbInformation, gstrSysName
            If txt身份证号.Enabled Then txt身份证号.SetFocus
            Exit Function
        End If
    End If
    
    '校验密码正确性
    gstrSQL = "Select 密码 From 保险帐户 Where 险类=[1] And 医保号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_四川眉山, Trim(txt卡号.Text))
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!密码, "") <> txt密码.Text Then
            MsgBox "密码错误，请检查！", vbInformation, gstrSysName
            If txt密码.Enabled Then txt密码.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Public Function BuildPatiInfo_眉山(ByVal bytType As Byte, ByVal strInfo As String, ByVal lng病人ID As Long) As Long
'功能：建立病人帐户信息
'参数：bytType=0-门诊,1-住院
'      strInfo='0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
'      8中心;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(1,2,3);15退休证号;16年龄段;17灰度级
'      18帐户增加累计;19帐户支出累计;20进入统筹累计;21统筹报销累计;22住院次数累计;23就诊类别
'      24本次起付线;25起付线累计;26基本统筹限额
'返回：病人ID
    Const MAX_BOUND = 26 '要求传入的信息段数
    
    Dim rsPati As ADODB.Recordset, str单位编码 As String, lng年龄 As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim lng中心 As Long, array信息 As Variant
    
    On Error GoTo errHandle
    
    If Len(Trim(strInfo)) <> 0 Then
        curDate = zlDatabase.Currentdate
        
        '200308z012:保证传入的信息串够用
        If UBound(Split(strInfo, ";")) < MAX_BOUND Then
            strInfo = strInfo & String(MAX_BOUND - UBound(Split(strInfo, ";")), ";")
        End If
        array信息 = Split(strInfo, ";")
        
        '从第7项内容中取出单位编码
        If array信息(7) Like "*(*" Then
            str单位编码 = Split(array信息(7), "(")(UBound(Split(array信息(7), "(")))
            str单位编码 = Mid(str单位编码, 1, Len(str单位编码) - 1)
        End If
        '取年龄
        If IsDate(array信息(5)) Then
            lng年龄 = Int(curDate - CDate(array信息(5))) / 365
        End If
        
        lng中心 = Val(array信息(8))
        
        If lng病人ID > 0 Then
            '该病人已经存在
            gstrSQL = "Select nvl(病人ID,0) 病人ID from 保险帐户 where 医保号=[1] and 险类=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "建立帐户", CStr(array信息(1)), TYPE_四川眉山)
            If rsTemp.EOF = False Then
                If rsTemp("病人ID") <> lng病人ID Then
                    MsgBox "已经存在相同医保号的另外一位病人，请在病人管理程序中将两位合并。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        '帐户唯一：险类,中心,医保号
        #If gverControl < 6 Then
            strSQL = "Select A.*,B.医保号 From 病人信息 A," & _
                " (Select * From 保险帐户" & _
                " Where 险类=" & TYPE_四川眉山 & _
                " And 医保号='" & CStr(array信息(1)) & "') B" & _
                " Where " & IIf(lng病人ID = 0, "A.病人ID=B.病人ID", "A.病人ID=B.病人ID(+) and A.病人ID=" & lng病人ID) '可能病人ID已经确定
        #Else
            strSQL = "Select A.病人id, A.门诊号, A.住院号, A.就诊卡号, A.卡验证码, A.费别, A.医疗付款方式, A.姓名, A.性别, A.年龄, A.出生日期, A.出生地点, A.身份证号, A.其他证件," & vbNewLine & _
                "A.身份, A.职业, A.民族, A.国籍, A.区域, A.学历, A.婚姻状况, A.家庭地址,A.家庭电话, A.家庭地址邮编 As 户口邮编, A.监护人, A.联系人姓名, A.联系人关系, A.联系人地址, " & vbNewLine & _
                "A.联系人电话, A.合同单位id, A.工作单位, A.单位电话, A.单位邮编, A.单位开户行, A.单位帐号, A.担保人, A.担保额, A.担保性质, A.就诊时间, A.就诊状态,A.就诊诊室, A.住院次数," & vbNewLine & _
                "A.当前科室id, A.当前病区id, A.当前床号, A.入院时间, A.出院时间, A.在院, A.Ic卡号, A.健康号, A.医保号, A.险类, A.查询密码, A.登记时间, A.停用时间, A.锁定,B.医保号 From 病人信息 A," & _
                " (Select * From 保险帐户" & _
                " Where 险类=[1]" & _
                " And 医保号=[2]) B" & _
                " Where " & IIf(lng病人ID = 0, "A.病人ID=B.病人ID", "A.病人ID=B.病人ID(+) and A.病人ID=[3]")  '可能病人ID已经确定
        #End If
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "医保接口", TYPE_四川眉山, CStr(array信息(1)), lng病人ID)
        
        If rsPati.EOF Then
            '无保险帐户则认为没有病人信息
            If lng病人ID = 0 Then lng病人ID = GetNextNO(1)
            strSQL = "zl_病人信息_Insert(" & lng病人ID & ",NULL,NULL,'社会基本医疗保险'," & _
                "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "NULL,'" & array信息(6) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,'" & array信息(7) & "',NULL,NULL,NULL," & _
                "NULL,NULL,NULL," & TYPE_四川眉山 & "," & _
                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call SQLTest(App.ProductName, "医保接口", strSQL)
            gcnOracle.Execute strSQL, , adCmdStoredProc
            Call SQLTest
        Else
            '有病人信息和保险帐户信息
            If rsPati("姓名") <> array信息(3) Then
                If MsgBox("病人原有登记的姓名是 " & rsPati("姓名") & " ，与刷卡得到的姓名 " & array信息(3) & " 不符，" & vbCrLf & _
                          "继续会更新病人原有的登记信息，是否确定？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
            If lng病人ID = 0 Then lng病人ID = rsPati!病人ID
            
            strSQL = "zl_病人信息_Update(" & _
                lng病人ID & "," & IIf(IsNull(rsPati!门诊号), "NULL", rsPati!门诊号) & "," & _
                IIf(IsNull(rsPati!住院号), "NULL", rsPati!住院号) & ",'" & IIf(IsNull(rsPati!费别), "", rsPati!费别) & "'," & _
                "'" & IIf(IsNull(rsPati!医疗付款方式), "", rsPati!医疗付款方式) & "'," & _
                "'" & array信息(3) & "','" & array信息(4) & "'," & IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
                "To_Date('" & Format(array信息(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "'" & IIf(IsNull(rsPati!出生地点), "", rsPati!出生地点) & "','" & array信息(6) & "'," & _
                "'" & IIf(IsNull(rsPati!身份), "", rsPati!身份) & "','" & IIf(IsNull(rsPati!职业), "", rsPati!职业) & "'," & _
                "'" & IIf(IsNull(rsPati!民族), "", rsPati!民族) & "','" & IIf(IsNull(rsPati!国籍), "", rsPati!国籍) & "'," & _
                "'" & IIf(IsNull(rsPati!学历), "", rsPati!学历) & "','" & IIf(IsNull(rsPati!婚姻状况), "", rsPati!婚姻状况) & "'," & _
                "'" & IIf(IsNull(rsPati!家庭地址), "", rsPati!家庭地址) & "','" & IIf(IsNull(rsPati!家庭电话), "", rsPati!家庭电话) & "'," & _
                "'" & IIf(IsNull(rsPati!户口邮编), "", rsPati!户口邮编) & "','" & IIf(IsNull(rsPati!联系人姓名), "", rsPati!联系人姓名) & "'," & _
                "'" & IIf(IsNull(rsPati!联系人关系), "", rsPati!联系人关系) & "','" & IIf(IsNull(rsPati!联系人地址), "", rsPati!联系人地址) & "'," & _
                "'" & IIf(IsNull(rsPati!联系人电话), "", rsPati!联系人电话) & "'," & IIf(IsNull(rsPati!合同单位ID), "NULL", rsPati!合同单位ID) & "," & _
                "'" & array信息(7) & "','" & IIf(IsNull(rsPati!单位电话), "", rsPati!单位电话) & "'," & _
                "'" & IIf(IsNull(rsPati!单位邮编), "", rsPati!单位邮编) & "','" & IIf(IsNull(rsPati!单位开户行), "", rsPati!单位开户行) & "'," & _
                "'" & IIf(IsNull(rsPati!单位帐号), "", rsPati!单位帐号) & "','" & IIf(IsNull(rsPati!担保人), "", rsPati!担保人) & "'," & _
                "" & IIf(IsNull(rsPati!担保额), "NULL", rsPati!担保额) & "," & TYPE_四川眉山 & ")"
            Call SQLTest(App.ProductName, "医保接口", strSQL)
            gcnOracle.Execute strSQL, , adCmdStoredProc
            Call SQLTest
        End If
        
        '插入或更新保险帐户信息(自动)
        strSQL = "zl_保险帐户_insert(" & lng病人ID & "," & TYPE_四川眉山 & "," & _
            lng中心 & "," & _
            "'" & IIf(array信息(0) = "-1", array信息(1), array信息(0)) & "'," & _
            "'" & array信息(1) & "'," & _
            "'" & array信息(2) & "'," & _
            "'" & array信息(9) & "'," & _
            "'" & array信息(15) & "'," & _
            "'" & array信息(10) & "'," & _
            "'" & str单位编码 & "'," & _
            array信息(11) & "," & _
            Val(array信息(12)) & "," & _
            IIf(Val(array信息(13)) = 0, "NULL", Val(array信息(13))) & "," & _
            IIf(Val(array信息(14)) = 0, 1, Val(array信息(14))) & "," & _
            IIf(Val(array信息(16)) = 0, lng年龄, Val(array信息(16))) & "," & _
            "'" & array信息(17) & "'," & _
            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call SQLTest(App.ProductName, "医保接口", strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
        
        '插入或更新帐户年度信息(自动)
        '200308z012:成都:保存"24本次起付线=zyjs,25起付线累计=tcbxbl,26基本统筹限额=zyxe"
        strSQL = "zl_帐户年度信息_Insert(" & lng病人ID & "," & TYPE_四川眉山 & "," & Year(curDate) & "," & _
            Val(array信息(18)) & "," & Val(array信息(19)) & "," & _
            Val(array信息(20)) & "," & Val(array信息(21)) & "," & _
            Val(Split(array信息(22), "/")(0)) & "," & Val(Split(array信息(22), "/")(1)) & "," & Val(array信息(24)) & "," & Val(array信息(25)) & "," & Val(array信息(26)) & ")"
        Call SQLTest(App.ProductName, "医保接口", strSQL)
        gcnOracle.Execute strSQL, , adCmdStoredProc
        Call SQLTest
    End If
    BuildPatiInfo_眉山 = lng病人ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
