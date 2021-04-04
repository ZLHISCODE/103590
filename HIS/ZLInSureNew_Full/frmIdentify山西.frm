VERSION 5.00
Begin VB.Form frmIdentify山西 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7215
   Icon            =   "frmIdentify山西.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd改密码 
      Caption         =   "修改密码(&P)"
      Height          =   375
      Left            =   540
      TabIndex        =   56
      Top             =   5220
      Width           =   1215
   End
   Begin VB.CommandButton cmd疾病信息 
      Caption         =   "…"
      Height          =   300
      Left            =   6510
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   4785
      Width           =   285
   End
   Begin VB.TextBox txtDiseaseName 
      Height          =   300
      Left            =   1260
      TabIndex        =   3
      Top             =   4785
      Width           =   5220
   End
   Begin VB.ComboBox cmb医疗类别 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4365
      Width           =   1950
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   5220
      Width           =   1215
   End
   Begin VB.TextBox txtPin 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4095
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4335
      Width           =   1440
   End
   Begin VB.CommandButton cmdReadCar 
      Caption         =   "读卡(&R)"
      Height          =   345
      Left            =   5730
      TabIndex        =   2
      Top             =   4320
      Width           =   960
   End
   Begin VB.Frame Frame2 
      Caption         =   "帐户基本信息"
      Height          =   1905
      Left            =   270
      TabIndex        =   7
      Top             =   2340
      Width           =   6645
      Begin VB.TextBox txtAcc00 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1425
         TabIndex        =   55
         Top             =   255
         Width           =   1200
      End
      Begin VB.TextBox txtAcc03 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1920
         TabIndex        =   54
         Top             =   562
         Width           =   1200
      End
      Begin VB.TextBox txtAcc05 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1890
         TabIndex        =   53
         Top             =   869
         Width           =   1200
      End
      Begin VB.TextBox txtAcc07 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1890
         TabIndex        =   52
         Top             =   1176
         Width           =   1200
      End
      Begin VB.TextBox txtAcc09 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1890
         TabIndex        =   51
         Top             =   1485
         Width           =   1200
      End
      Begin VB.TextBox txtAcc10 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5235
         TabIndex        =   47
         Top             =   1485
         Width           =   1200
      End
      Begin VB.TextBox txtAcc08 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5235
         TabIndex        =   45
         Top             =   1170
         Width           =   1200
      End
      Begin VB.TextBox txtAcc06 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5235
         TabIndex        =   44
         Top             =   870
         Width           =   1200
      End
      Begin VB.TextBox txtAcc04 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5235
         TabIndex        =   43
         Top             =   555
         Width           =   1200
      End
      Begin VB.TextBox txtAcc02 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5805
         TabIndex        =   42
         Top             =   255
         Width           =   630
      End
      Begin VB.TextBox txtAcc01 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3660
         TabIndex        =   41
         Top             =   255
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "本年公务员补助支出累计"
         Height          =   240
         Left            =   3120
         TabIndex        =   40
         Top             =   1515
         Width           =   2070
      End
      Begin VB.Label Label21 
         Caption         =   "本年现金支出累计"
         Height          =   240
         Left            =   210
         TabIndex        =   39
         Top             =   1515
         Width           =   1710
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "本年统筹支付累计"
         Height          =   240
         Left            =   3480
         TabIndex        =   38
         Top             =   1206
         Width           =   1710
      End
      Begin VB.Label Label19 
         Caption         =   "本年帐户支出累计"
         Height          =   240
         Left            =   210
         TabIndex        =   37
         Top             =   1206
         Width           =   1710
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "本年进入统筹累计"
         Height          =   240
         Left            =   3480
         TabIndex        =   36
         Top             =   899
         Width           =   1710
      End
      Begin VB.Label Label17 
         Caption         =   "本年自理累计"
         Height          =   240
         Left            =   210
         TabIndex        =   35
         Top             =   899
         Width           =   1710
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "本年自费累计"
         Height          =   240
         Left            =   3480
         TabIndex        =   34
         Top             =   592
         Width           =   1710
      End
      Begin VB.Label Label15 
         Caption         =   "本年总费用支出累计"
         Height          =   240
         Left            =   210
         TabIndex        =   33
         Top             =   592
         Width           =   1710
      End
      Begin VB.Label Label14 
         Caption         =   "本年住院次数"
         Height          =   240
         Left            =   4530
         TabIndex        =   32
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "帐户年度"
         Height          =   240
         Left            =   2790
         TabIndex        =   31
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "帐户结余金额"
         Height          =   240
         Left            =   210
         TabIndex        =   30
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "个人基本信息"
      Height          =   1995
      Left            =   270
      TabIndex        =   6
      Top             =   255
      Width           =   6645
      Begin VB.TextBox txtEmp10 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4605
         TabIndex        =   18
         Top             =   1500
         Width           =   1860
      End
      Begin VB.TextBox txtEmp09 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1500
         TabIndex        =   17
         Top             =   1500
         Width           =   1560
      End
      Begin VB.TextBox txtEmp08 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4605
         TabIndex        =   16
         Top             =   1185
         Width           =   1875
      End
      Begin VB.TextBox txtEmp07 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1110
         TabIndex        =   15
         Top             =   1185
         Width           =   1935
      End
      Begin VB.TextBox txtEmp06 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4605
         TabIndex        =   14
         Top             =   870
         Width           =   1890
      End
      Begin VB.TextBox txtEmp05 
         Enabled         =   0   'False
         Height          =   270
         Left            =   780
         TabIndex        =   13
         Top             =   870
         Width           =   1830
      End
      Begin VB.TextBox txtEmp04 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5340
         TabIndex        =   12
         Top             =   555
         Width           =   1155
      End
      Begin VB.TextBox txtEmp03 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3465
         TabIndex        =   11
         Top             =   555
         Width           =   480
      End
      Begin VB.TextBox txtEmp02 
         Enabled         =   0   'False
         Height          =   270
         Left            =   795
         TabIndex        =   10
         Top             =   555
         Width           =   1140
      End
      Begin VB.TextBox txtEmp01 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4665
         TabIndex        =   9
         Top             =   240
         Width           =   1830
      End
      Begin VB.TextBox txtEmp00 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1125
         TabIndex        =   8
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label11 
         Caption         =   "在院状态"
         Height          =   240
         Left            =   3795
         TabIndex        =   29
         Top             =   1530
         Width           =   810
      End
      Begin VB.Label Label10 
         Caption         =   "照顾人员标志"
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "医疗人员类别"
         Height          =   240
         Left            =   3420
         TabIndex        =   27
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "单位编号"
         Height          =   240
         Left            =   240
         TabIndex        =   26
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "医疗证号"
         Height          =   240
         Left            =   3810
         TabIndex        =   25
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "卡号"
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "出生日期"
         Height          =   240
         Left            =   4500
         TabIndex        =   23
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "性别"
         Height          =   240
         Left            =   3030
         TabIndex        =   22
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "姓名"
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "身份证号"
         Height          =   240
         Left            =   3885
         TabIndex        =   20
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "个人编号"
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   5580
      TabIndex        =   5
      Top             =   5220
      Width           =   1215
   End
   Begin VB.Label Label25 
      Caption         =   "入院病种"
      Height          =   210
      Left            =   315
      TabIndex        =   49
      Top             =   4830
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "医疗类别"
      Height          =   225
      Left            =   330
      TabIndex        =   48
      Top             =   4410
      Width           =   810
   End
   Begin VB.Label Label23 
      Caption         =   "密码"
      Height          =   270
      Left            =   3525
      TabIndex        =   46
      Top             =   4380
      Width           =   480
   End
End
Attribute VB_Name = "frmIdentify山西"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号
Private mlng病人ID As Long
Private mstrReturn As String

Function 身份标识(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = "-1"
    Me.Show 1
    lng病人ID = mlng病人ID
    身份标识 = mstrReturn
End Function

Private Sub CancelButton_Click()
        mstrReturn = "-1"
    Unload Me
End Sub

Private Sub cmdReadCar_Click()
    Dim str医疗人员类别 As String
    Dim str公务员标志 As String
    Dim str照顾人员标志 As String
    
    If Len(Trim(txtPin.Text)) = 0 Then
        MsgBox "请输入密码！", vbInformation, gstrSysName
        Exit Sub
    Else
        If 读卡审核(Trim(txtPin.Text)) Then
            txtEmp00.Text = IIf(IsNull(g个人基本信息.个人编号00), "", g个人基本信息.个人编号00)
            txtEmp01.Text = IIf(IsNull(g个人基本信息.身份证号01), "", g个人基本信息.身份证号01)
            txtEmp02.Text = IIf(IsNull(g个人基本信息.姓名02), "", g个人基本信息.姓名02)
            txtEmp03.Text = IIf(IsNull(g个人基本信息.性别03), "", g个人基本信息.性别03)
            txtEmp04.Text = IIf(IsNull(g个人基本信息.出生日期04), "", g个人基本信息.出生日期04)
            txtEmp05.Text = IIf(IsNull(g个人基本信息.卡号05), "", g个人基本信息.卡号05)
            txtEmp06.Text = IIf(IsNull(g个人基本信息.医疗证号06), "", g个人基本信息.医疗证号06)
            txtEmp07.Text = IIf(IsNull(g个人基本信息.单位编号07), "", g个人基本信息.单位编号07)
            
            str医疗人员类别 = IIf(IsNull(g个人基本信息.医疗人员类别08), "", g个人基本信息.医疗人员类别08)
            Select Case str医疗人员类别
                Case 11
                    str医疗人员类别 = "在职"
                Case 21
                    str医疗人员类别 = "退休"
                Case 33
                    str医疗人员类别 = "二等乙级伤残军人"
                Case 91
                    str医疗人员类别 = "其他人员"
            End Select
            txtEmp08.Text = str医疗人员类别
            
            str公务员标志 = IIf(IsNull(g个人基本信息.公务员标志09), "", g个人基本信息.公务员标志09)
            txtEmp09.Text = IIf(str公务员标志 = 0, "否", "是")
            
            str照顾人员标志 = IIf(IsNull(g个人基本信息.照顾人员标志10), "", g个人基本信息.照顾人员标志10)
            txtEmp10.Text = IIf(str照顾人员标志 = 0, "否", "是")
            
            txtAcc00.Text = IIf(IsNull(g帐户基本信息.帐户结余金额00), "0.00", g帐户基本信息.帐户结余金额00)
            txtAcc01.Text = IIf(IsNull(g帐户基本信息.帐户年度01), "", g帐户基本信息.帐户年度01)
            txtAcc02.Text = IIf(IsNull(g帐户基本信息.本年住院次数02), "0", g帐户基本信息.本年住院次数02)
            txtAcc03.Text = IIf(IsNull(g帐户基本信息.本年总费用支出累计03), "0.00", g帐户基本信息.本年总费用支出累计03)
            txtAcc04.Text = IIf(IsNull(g帐户基本信息.本年自费累计05), "0.00", g帐户基本信息.本年自费累计05)
            txtAcc05.Text = IIf(IsNull(g帐户基本信息.本年自理累计06), "0.00", g帐户基本信息.本年自理累计06)
            txtAcc06.Text = IIf(IsNull(g帐户基本信息.本年进入统筹累计07), "0.00", g帐户基本信息.本年进入统筹累计07)
            txtAcc07.Text = IIf(IsNull(g帐户基本信息.本年帐户支出累计08), "0.00", g帐户基本信息.本年帐户支出累计08)
            txtAcc08.Text = IIf(IsNull(g帐户基本信息.本年统筹支付累计10), "0.00", g帐户基本信息.本年统筹支付累计10)
            txtAcc09.Text = IIf(IsNull(g帐户基本信息.本年现金支出累计11), "0.00", g帐户基本信息.本年现金支出累计11)
            txtAcc10.Text = IIf(IsNull(g帐户基本信息.本年公务员补助支出累计13), "0.00", g帐户基本信息.本年公务员补助支出累计13)
            Call 提交_山西
            '此时读取的信息供显示用
            '这时提交一次，避免用户在界面上长时间停留，造成锁定和确定时，医保中心已经改了状态，造成不一致
            '确定时，还要再调一次读卡审核
             SendKeys ("{Tab}")
        Else
            mstrReturn = "-1"
        End If
    End If
    
End Sub


Private Sub cmd改密码_Click()
    Dim strOldPass As String, strNewPass As String
    
    strOldPass = Trim(txtPin.Text)
    strNewPass = ""
    
    strNewPass = frm修改密码.ChangePassword("", strOldPass)
    
    If Nvl(strNewPass) = "" Then Exit Sub
    
    If 修改密码_山西(strOldPass, strNewPass) Then
        txtPin.Text = strNewPass
    End If
End Sub

Private Sub Form_Load()

    '初始化医疗类别
    If mbytType = 0 Then
        '11      普通门诊
        '12  大额疾病门诊
        '14  定点药店购药
        '17  门诊急诊
        
       cmb医疗类别.AddItem "普通门诊"
       cmb医疗类别.ItemData(cmb医疗类别.NewIndex) = 11
       
       cmb医疗类别.AddItem "大额疾病门诊"
       cmb医疗类别.ItemData(cmb医疗类别.NewIndex) = 12
       
       cmb医疗类别.AddItem "门诊急诊"
       cmb医疗类别.ItemData(cmb医疗类别.NewIndex) = 17
       
       '设默认值
       cmb医疗类别.Text = "普通门诊"
       cmb医疗类别.Tag = 11
    End If
    
    If mbytType = 1 Then
        '21  普通住院
        '23  转外住院
        '24  转院住院
        '26  门诊急诊转入住院
       cmb医疗类别.AddItem "普通住院"
       cmb医疗类别.ItemData(cmb医疗类别.NewIndex) = 21
       
       cmb医疗类别.AddItem "转院住院"
       cmb医疗类别.ItemData(cmb医疗类别.NewIndex) = 24
       
       cmb医疗类别.AddItem "门诊急诊转入住院"
       cmb医疗类别.ItemData(cmb医疗类别.NewIndex) = 26
       '设默认值
       cmb医疗类别.Text = "普通住院"
       cmb医疗类别.Tag = 21
    End If
    
End Sub

Private Sub cmb医疗类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub OKButton_Click()

Dim strEmpInfo As String, straccinfo As String  ''存放个人基本信息和帐户信息
Dim strTmpSQL As String '临时SQL语句
Dim rsTmp As New ADODB.Recordset  '临时记录集
Dim cur病种ID  As Currency  '用currency不容易出现界面未知错误.
Dim str病种简码 As String

'病种选择没有，要判断
If txtDiseaseName.Tag = "" Then
     MsgBox "请选择病种！", vbInformation, gstrSysName
     mstrReturn = "-1"
     txtDiseaseName.SetFocus
     Exit Sub
End If
  
'读卡判断,再次读卡，
If 读卡审核(Trim(txtPin.Text)) Then

    '保存病种信息到保险病种表中
      '判断库中有没有这个病种,如有，则直接取得病种ID
    strTmpSQL = "select * from 保险病种 where 险类=" & TYPE_山西 & _
                                         " and 编码='" & txtDiseaseName.Tag & "'"
    Call OpenRecordset(rsTmp, "查病种ID", strTmpSQL)
    If rsTmp.EOF Then
        strTmpSQL = "select 保险病种_ID.NextVal as ID from Dual "
        Call OpenRecordset(rsTmp, "取病种ID", strTmpSQL)
        cur病种ID = 1
        If Not rsTmp.EOF Then cur病种ID = rsTmp!ID
        
        strTmpSQL = "select zlspellcode('" & txtDiseaseName.Text & "') as 简码 from dual"
        Call OpenRecordset(rsTmp, "取病种简码", strTmpSQL)
        str病种简码 = rsTmp!简码
        
        strTmpSQL = "zl_保险病种_insert(" & cur病种ID & "," & TYPE_山西 & ",'" & _
                                         txtDiseaseName.Tag & "','" & _
                                         txtDiseaseName.Text & "','" & _
                                         str病种简码 & "',1,NULL,NULL)"
        gcnOracle.Execute strTmpSQL, , adCmdStoredProc
        
        
        rsTmp.Close
        Set rsTmp = Nothing
    Else
       cur病种ID = rsTmp!ID
    End If
   
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    
    strEmpInfo = g个人基本信息.卡号05                               '0卡号
    strEmpInfo = strEmpInfo & ";" & g个人基本信息.个人编号00             '1医保号
    strEmpInfo = strEmpInfo & ";" & txtPin.Text               '2密码
    strEmpInfo = strEmpInfo & ";" & g个人基本信息.姓名02               '3姓名
    strEmpInfo = strEmpInfo & ";" & g个人基本信息.性别03               '4性别
    strEmpInfo = strEmpInfo & ";" & Mid(g个人基本信息.出生日期04, 1, 4) & "-" & Mid(g个人基本信息.出生日期04, 5, 2) & "-" & Mid(g个人基本信息.出生日期04, 7, 2)        '5出生日期
    strEmpInfo = strEmpInfo & ";" & g个人基本信息.身份证号01           '6身份证
    strEmpInfo = strEmpInfo & ";" & g个人基本信息.单位编号07         '7.单位名称(编码)
    
    straccinfo = ";0"                                          '8.中心代码
    straccinfo = straccinfo & ";"                    '9.顺序号
    straccinfo = straccinfo & ";" & g个人基本信息.医疗人员类别08           '10人员身份
    straccinfo = straccinfo & ";" & g帐户基本信息.帐户结余金额00      '11帐户余额
    straccinfo = straccinfo & ";0" ' & g个人基本信息.在院状态16                             '12当前状态
    straccinfo = straccinfo & ";" & cur病种ID                  '13病种ID
    straccinfo = straccinfo & ";1"                            '14在职(1,2,3)
    straccinfo = straccinfo & ";"                             '15退休证号
    straccinfo = straccinfo & ";"                             '16年龄段
    straccinfo = straccinfo & ";1"                            '17灰度级
    straccinfo = straccinfo & ";" & g帐户基本信息.帐户结余金额00      '18帐户增加累计
    straccinfo = straccinfo & ";0"                              '19帐户支出累计
    straccinfo = straccinfo & ";0"                            '20上年工资总额
    straccinfo = straccinfo & ";"      '21
    straccinfo = straccinfo & ";" & g帐户基本信息.本年住院次数02      '22住院次数累计
    
    mlng病人ID = BuildPatiInfo(0, strEmpInfo & straccinfo, mlng病人ID, TYPE_山西)
    
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_山西 & ",'就诊类别','''" & cmb医疗类别.ItemData(cmb医疗类别.ListIndex) & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "应诊类别")
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strEmpInfo & ";" & mlng病人ID & straccinfo
    End If
    Unload Me
Else
    mstrReturn = "-1"
End If


End Sub

Private Sub txtDiseaseName_KeyPress(KeyAscii As Integer)
  ''调病种选择器
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtDiseaseName.Text) = "" Then Exit Sub
    Call 病种选择
End Sub

Private Sub 病种选择(Optional strLoad As String = 1)
    Dim rsTmp As ADODB.Recordset
    Dim strTmpSQL As String
    If strLoad = 1 Then
        strTmpSQL = "select rownum as ID,aka120  病种编码,aka121 病种名称,aka066 助记码,aae035 变更日期 from ka06" & _
                    " where aka120 like '%" & Trim(txtDiseaseName.Text) & "%' or aka121 like '%" & Trim(txtDiseaseName.Text) & "%' or Upper(aka066) like '%" & UCase(Trim(txtDiseaseName.Text)) & "%'"
    Else
        strTmpSQL = "select rownum as ID,aka120  病种编码,aka121 病种名称,aka066 助记码,aae035 变更日期 from ka06"
    End If
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "病种", True, , , , , gcnSxDr)
    If rsTmp Is Nothing Then Exit Sub
    txtDiseaseName.Text = rsTmp!病种名称
    txtDiseaseName.Tag = rsTmp!病种编码
    OKButton.SetFocus
End Sub

Private Sub cmd疾病信息_Click()
    Call 病种选择(0)
End Sub

Private Sub txtPin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub


