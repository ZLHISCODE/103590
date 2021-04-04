VERSION 5.00
Begin VB.Form frmIdentify铜山县 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   4050
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5370
   Icon            =   "frmIdentify铜山县.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk急诊 
      Caption         =   "急诊"
      Height          =   285
      Left            =   4395
      TabIndex        =   6
      Top             =   3060
      Width           =   660
   End
   Begin VB.TextBox txtDiseaseName 
      Height          =   300
      Left            =   1170
      TabIndex        =   2
      Top             =   2670
      Width           =   3585
   End
   Begin VB.CommandButton cmd疾病信息 
      Caption         =   "…"
      Height          =   300
      Left            =   4770
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2670
      Width           =   285
   End
   Begin VB.TextBox txt医保卡序号 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4050
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txt特诊 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3735
      TabIndex        =   28
      Top             =   2175
      Width           =   1335
   End
   Begin VB.TextBox txt余额 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   26
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txt人员身份 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3735
      TabIndex        =   24
      Top             =   1740
      Width           =   1335
   End
   Begin VB.TextBox txt在职 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   22
      Top             =   1740
      Width           =   1335
   End
   Begin VB.TextBox txt工作单位 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   20
      Top             =   1335
      Width           =   3900
   End
   Begin VB.TextBox txt年龄 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4605
      TabIndex        =   18
      Top             =   525
      Width           =   450
   End
   Begin VB.TextBox txt出生年月 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4065
      TabIndex        =   16
      Top             =   930
      Width           =   1005
   End
   Begin VB.TextBox txt性别 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3105
      TabIndex        =   14
      Top             =   525
      Width           =   465
   End
   Begin VB.TextBox txt身份证 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   12
      Top             =   930
      Width           =   1785
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   10
      Top             =   525
      Width           =   915
   End
   Begin VB.TextBox txt个人编号 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdReadCard 
      Caption         =   "读卡(&R)"
      Height          =   375
      Left            =   420
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtCom 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   1170
      TabIndex        =   7
      Text            =   "1"
      Top             =   3105
      Width           =   450
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3945
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   2175
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "病种选择："
      Height          =   210
      Left            =   270
      TabIndex        =   32
      Top             =   2745
      Width           =   900
   End
   Begin VB.Label Label12 
      Caption         =   "医保卡序号："
      Height          =   255
      Left            =   2970
      TabIndex        =   30
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label Label11 
      Caption         =   "特诊人员："
      Height          =   195
      Left            =   2790
      TabIndex        =   29
      Top             =   2190
      Width           =   915
   End
   Begin VB.Label Label10 
      Caption         =   "帐户余额："
      Height          =   195
      Left            =   255
      TabIndex        =   27
      Top             =   2190
      Width           =   915
   End
   Begin VB.Label Label9 
      Caption         =   "人员身份："
      Height          =   195
      Left            =   2790
      TabIndex        =   25
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label8 
      Caption         =   "在职状态："
      Height          =   195
      Left            =   255
      TabIndex        =   23
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "工作单位："
      Height          =   195
      Left            =   255
      TabIndex        =   21
      Top             =   1410
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "年龄："
      Height          =   195
      Left            =   4080
      TabIndex        =   19
      Top             =   585
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "出生年月："
      Height          =   195
      Left            =   3135
      TabIndex        =   17
      Top             =   1005
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "性别："
      Height          =   195
      Left            =   2565
      TabIndex        =   15
      Top             =   585
      Width           =   540
   End
   Begin VB.Label Label3 
      Caption         =   "身份证号："
      Height          =   195
      Left            =   255
      TabIndex        =   13
      Top             =   1005
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "姓名："
      Enabled         =   0   'False
      Height          =   195
      Left            =   615
      TabIndex        =   11
      Top             =   585
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "个人编号："
      Height          =   255
      Left            =   255
      TabIndex        =   9
      Top             =   150
      Width           =   915
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "IC卡串口："
      Height          =   180
      Index           =   3
      Left            =   285
      TabIndex        =   8
      Top             =   3135
      Width           =   900
   End
End
Attribute VB_Name = "frmIdentify铜山县"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const P_ERRORMSG = 167782161
Private Const P_FILENAME = 167782162
Private Const P_FILEBUF = 201336595
Private Const P_USERNAME = 167782164
Private Const P_FILETIME = 167782165
Private Const P_FLAG = 167782166
Private Const P_OMSG = 167782167
Private Const P_LEN = 33564440
Private Const P_RWID = 167782169
Private Const P_RWMCH = 167782170
Private Const P_RYH = 167782181
Private Const P_MM = 167782182
Private Const P_RES = 167782183
Private Const P_MLIST = 167782184
Private Const P_TLIST = 167782185
Private Const P_WLIST = 167782186
Private Const P_PLIST = 167782187
Private Const P_JGM = 167782188
Private Const P_JGMCH = 167782189
Private Const P_TBR = 167782210
Private Const P_XM = 167782211
Private Const P_DWM = 167782212
Private Const P_KXH = 167782213
Private Const P_SHBZH = 167782214
Private Const P_YYDJ = 167782215
Private Const P_DWMCH = 167782216
Private Const P_CZRYH = 167782217
Private Const P_DJH = 167782263
Private Const P_CFH = 167782265
Private Const P_YBKSM = 167782266
Private Const P_YYKSM = 167782267
Private Const P_YSRYH = 167782268
Private Const P_YSXM = 167782269
Private Const P_BZM = 167782270
Private Const P_JZYMD = 167782271
Private Const P_RYLB = 167782272
Private Const P_XB = 167782274
Private Const P_NL = 33564547
Private Const P_DWXZ = 167782277
Private Const P_JZLB = 167782278
Private Const P_TSMZ = 167782279
Private Const P_TCBZ = 167782280
Private Const P_GWFL = 167782281
Private Const P_LYLB = 167782282
Private Const P_ZFY = 134227851
Private Const P_YF = 134227852
Private Const P_ZLXMF = 134227853
Private Const P_QCGRZH = 134227854
Private Const P_JBTCZF = 134227855
Private Const P_TSTCZF = 134227856
Private Const P_GWTCZF = 134227857
Private Const P_DBTCZF = 134227858
Private Const P_GRZHZF = 134227859
Private Const P_GRZF = 134227860
Private Const P_GRZL = 134227861
Private Const P_QMGRZH = 134227862
Private Const P_GRFD = 134227863
Private Const P_TZH = 167782297
Private Const P_SFGWYMM = 167782298
Private Const P_CZYMD = 167782300
Private Const P_ZYXH = 167782301
Private Const P_ZYH = 167782302
Private Const P_RYBQ = 167782303
Private Const P_RYCWH = 167782304
Private Const P_RYYMD = 167782305
Private Const P_YJHJ = 134227874
Private Const P_CYYMD = 167782307
Private Const P_CYXZ = 167782308
Private Const P_ZYTZ = 167782309
Private Const P_ZWYYM = 167782310
Private Const P_GHF = 134227879
Private Const P_ZLF = 134227880
Private Const P_TSRY = 167782361
Private Const P_QKBZ = 167782362
Private Const P_MSG = 167782400
Private Const P_CXLB = 167782363
Private Const P_ZBM = 167782364
Private Const P_JG = 134227933
Private Const P_SL = 33564638
Private Const P_YBFYLJ = 134227935
Private Const P_ZXYY = 167782370
Private Const P_CSNY = 167782460
Private Const P_RYTZ = 167782461
Private Const P_GRSF = 167782462
Private Const P_GRQK = 134228032
Private Const P_HISID = 167782465
Private Const P_GLY = 167782466
Private Const P_LB = 167782467
Private Const P_GHXH = 167782468
Private Const P_QZ_QFFD = 134228037
Private Const P_QZ_JBFD = 134228038
Private Const P_QZ_DBFD = 134228039
Private Const P_QZ_CFD = 134228040
Private Const P_BCJBFWFY = 134228041
Private Const P_BCJBTSFY = 134228042
Private Const P_BCDBFWFY = 134228043
Private Const P_BCDBTSFY = 134228044
Private Const P_LJQFFY = 134228045
Private Const P_LJJBFWFY = 134228046
Private Const P_LJDBFWFY = 134228047
Private Const P_LJCFDFY = 134228048
Private Const P_GYYMD = 167782481
Private Const P_LXDH = 167782482
Private Const P_CQZY = 167782483
Private Const P_YYBQM = 167782484
Private Const P_YYKSMCH = 167782485
Private Const P_CWS = 33564758
Private Const P_GJYS = 33564759
Private Const P_ZJYS = 33564760
Private Const P_CJYS = 33564761
Private Const P_PYM = 167782490
Private Const P_WBM = 167782491
Private Const P_YBBM = 167782492
Private Const P_TYM = 167782493
Private Const P_YWM = 167782494
Private Const P_GG = 167782495
Private Const P_ZXGG = 167782496
Private Const P_JXM = 167782497
Private Const P_FLM = 167782498
Private Const P_JLDW = 167782499
Private Const P_ZXJLDW = 167782500
Private Const P_DJ1 = 134228069
Private Const P_DJ2 = 134228070
Private Const P_FDBL = 134228071
Private Const P_HSXS = 134228072
Private Const P_ZFLB = 167782505
Private Const P_ZFBL = 134228074
Private Const P_ZFLB1 = 167782507
Private Const P_ZFBL1 = 134228076
Private Const P_YBXL = 167782509
Private Const P_YLDL = 167782510
Private Const P_YLXL = 167782511
Private Const P_GMPRZ = 167782512
Private Const P_CFYBZ = 167782513
Private Const P_CD = 167782514
Private Const P_CDTZ = 167782515
Private Const P_BZ = 167782516
Private Const P_CFQX = 167782517
Private Const P_ZYYSZH = 167782518
Private Const P_TSJZ = 167782519
Private Const P_SPM = 167782520
Private Const P_DJ = 134228089
Private Const P_KCTZ = 167782522
Private Const P_SHTZ = 167782523
Private Const P_SHRYH = 167782524
Private Const P_SHYMD = 167782525
Private Const P_MCH = 167782526
Private Const P_YMD1 = 167782527
Private Const P_YMD2 = 167782528
Private Const P_ZXDJ = 167782529
Private Const P_LSTZ = 167782530
Private Const P_JSTZ = 167782531
Private Const P_TPTZ = 167782532
Private Const P_TSBZ = 167782533
Private Const P_YMD = 167782534
Private Const P_BNZYCS = 33564832
Private Const P_ZWYYMCH = 167782561
Private Const P_SZDJH = 167782562
Private Const P_NAME = 167782563
Private Const P_VAL = 167782564
Private Const P_JZ = 167782565
Private Const P_YYM = 167782566
Private Const P_YJSRYH = 167782567
Private Const P_QZ_XXFD = 134228136
Private Const P_KDYSRYH = 167782569
Private Const P_KDYSXM = 167782570
Private Const P_DZ = 167782571
Private Const P_DH = 167782572
Private Const P_CBNS = 33564845
Private Const P_GWQK = 134228142
Private Const P_SPDJH = 167782576
Private Const P_SPSL = 33564849
Private Const P_SPJE = 134228146
Private Const P_SPBZ = 167782579
Private Const P_TCJJZF = 134228229
Private Const P_BCYLZF = 134228230
Private Const P_GWYBZZF = 134228231
Private Const P_LJZF = 134228232
Private Const P_LJZL = 134228233
Private Const P_FYYMD = 167782666
Private Const P_BM = 167782667
Private Const P_DYLB = 167782668
Private Const P_TYMBM = 167782669
Private Const P_SPMBM = 167782670
Private Const P_XXM = 167782671
Private Const P_SFJZ = 167782672
Private Const P_DJ3 = 134228241
Private Const P_SBGS = 167782680
Private Const P_LJ_MZQF = 134228249
Private Const P_LJ_ZYQF = 134228250
Private Const P_LJ_MZFY = 134228251
Private Const P_LJ_TCFD = 134228252
Private Const P_LJ_MZGW = 134228253
Private Const P_LJ_TC = 134228254
Private Const P_LJ_DB = 134228255
Private Const P_DTGYCS = 33564960
Private Const P_DYGYCS = 33564961
Private Const P_L1 = 33565033
Private Const P_L2 = 33565034
Private Const P_L3 = 33565035
Private Const P_L4 = 33565036
Private Const P_L5 = 33565037
Private Const P_D1 = 134228335
Private Const P_D2 = 134228336
Private Const P_D3 = 134228337
Private Const P_D4 = 134228338
Private Const P_D5 = 134228339
Private Const P_S1 = 167782772
Private Const P_S2 = 167782773
Private Const P_S3 = 167782774
Private Const P_S4 = 167782775
Private Const P_S5 = 167782776
Private Const P_S6 = 167782777
Private Const P_S7 = 167782778
Private Const P_S8 = 167782779
Private Const P_S9 = 167782780
Private Const P_S10 = 167782791
Private Const P_S11 = 167782792
Private Const P_S12 = 167782793
Private Const P_S13 = 167782794
Private Const P_S14 = 167782795
Private Const P_S15 = 167782796
Private Const P_D6 = 134228435
Private Const P_D7 = 134228436
Private Const P_D8 = 134228437
Private Const P_D9 = 134228438
Private Const P_D10 = 134228439

Dim mlngReturn As Long
Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号
Private mlng病人ID As Long
Private mstrReturn As String

Public Function GetIdentify(ByVal bytType As Byte, Optional ByVal lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    
    Me.Show 1
    lng病人ID = mlng病人ID
    GetIdentify = mstrReturn
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub



Private Sub chk急诊_LostFocus()
    If chk急诊.Value = 2 Then chk急诊.Value = 1
End Sub

Private Sub cmdReadCard_Click()
    Dim str个人编号 As String, str医保卡序号 As String
    Dim str姓名 As String, str身份证 As String, str性别 As String, str出生年月 As String
    Dim lng年龄 As Long, str工作单位 As String, str人员特征 As String, str特殊人员 As String
    Dim dbl帐户余额 As Double, str特殊就诊人员 As String, str出生日期 As String
    Dim intCOM As Long, strRead As String
    
    
    str个人编号 = Space(9): str医保卡序号 = Space(3)
    mlngReturn = tsx_read_ic(str个人编号, str医保卡序号)
    Call WriteBusinessLOG("tsx_read_ic", str个人编号 & "," & str医保卡序号, mlngReturn)
    If mlngReturn <> -1 Then
        txt个人编号.Text = str个人编号
        txt医保卡序号.Text = str医保卡序号
        Call tsx_取基本信息
    Else
        MsgBox tsx_getlasterr(), vbInformation, gstrSysName
    End If
    

End Sub

Private Sub cmdReadCard_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub Form_Load()
    txtCom.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", 0) + 1
    OKButton.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", CStr(txtCom.Text - 1)
End Sub

Private Sub OKButton_Click()
    Dim strEmpInfo As String
    Dim strAccinfo As String

    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    
    '病种选择没有
    If txtDiseaseName.Tag = "" Then
         MsgBox "请选择病种！", vbInformation, gstrSysName
         mstrReturn = ""
         txtDiseaseName.SetFocus
         Exit Sub
    End If
    
    strEmpInfo = txt医保卡序号.Text                               '0卡号
    strEmpInfo = strEmpInfo & ";" & txt个人编号.Text              '1医保号
    strEmpInfo = strEmpInfo & ";"              '2密码
    strEmpInfo = strEmpInfo & ";" & txt姓名.Text                '3姓名
    strEmpInfo = strEmpInfo & ";" & txt性别.Text                '4性别
    strEmpInfo = strEmpInfo & ";" & txt出生年月.Text         '5出生日期
    strEmpInfo = strEmpInfo & ";" & txt身份证.Text           '6身份证
    strEmpInfo = strEmpInfo & ";" & txt工作单位.Text           '7.单位名称(编码)
    
    strAccinfo = ";0"                                          '8.中心代码
    strAccinfo = strAccinfo & ";"                    '9.顺序号
    strAccinfo = strAccinfo & ";"              '10人员身份
    strAccinfo = strAccinfo & ";" & Val(txt余额.Text)        '11帐户余额
    strAccinfo = strAccinfo & ";"     ' & g个人基本信息.在院状态16                             '12当前状态
    strAccinfo = strAccinfo & ";" & txtDiseaseName.Tag                   '13病种ID
    strAccinfo = strAccinfo & ";"                           '14在职(1,2,3)
    strAccinfo = strAccinfo & ";"                             '15退休证号
    strAccinfo = strAccinfo & ";"                             '16年龄段
    strAccinfo = strAccinfo & ";1"                            '17灰度级
    strAccinfo = strAccinfo & ";" & Val(txt余额.Text)      '18帐户增加累计
    strAccinfo = strAccinfo & ";0"                              '19帐户支出累计
    strAccinfo = strAccinfo & ";0"                            '20上年工资总额
    strAccinfo = strAccinfo & ";"      '21
    strAccinfo = strAccinfo & ";"       '22住院次数累计
    
    mlng病人ID = BuildPatiInfo(mbytType, strEmpInfo & strAccinfo, mlng病人ID, TYPE_铜山县)
    
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_铜山县 & ",'就诊类别','''" & txt特诊.Text & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "应诊类别")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_铜山县 & ",'人员身份','''" & txt人员身份.Text & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "帐户状态")
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strEmpInfo & ";" & mlng病人ID & strAccinfo
        g常用信息_铜山.病人ID = mlng病人ID
        g常用信息_铜山.个人编号 = txt个人编号.Text
        g常用信息_铜山.医保卡序号 = txt医保卡序号.Text
        g常用信息_铜山.病种编码 = Mid(txtDiseaseName.Text, 2, InStr(txtDiseaseName.Text, "）") - 2)
        g常用信息_铜山.急诊否 = chk急诊.Value
    End If
    Unload Me
End Sub

Private Sub txtCom_Change()
    If InStr("123456789", txtCom.Text) <= 0 Then
        MsgBox "请输入数字!", vbInformation, gstrSysName
        txtCom.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", 0) + 1
    End If
End Sub

Private Sub txtDiseaseName_KeyPress(KeyAscii As Integer)
  ''调病种选择器
    Call WriteBusinessLOG("KeyAscii", "", KeyAscii)

    If KeyAscii <> vbKeyReturn Then Exit Sub
    Call WriteBusinessLOG("KeyAscii", "", KeyAscii)
    If Trim(txtDiseaseName.Text) = "" Then Exit Sub
    Call WriteBusinessLOG("病种选择 开始", "", "")
    
    Call 病种选择
    Call WriteBusinessLOG("病种选择 结束", "", "")
    
End Sub

Private Sub 病种选择(Optional strLoad As String = 1)
    Dim rsTmp As ADODB.Recordset, strtab As String
    Dim strTmpSQL As String
    
    If mbytType = 0 Then
        strtab = "MZBZ"
    Else
        strtab = "ICD10"
    End If
    
    If strLoad = 1 Then
        strTmpSQL = "select ID,病种编码,病种名称,拼音码 from " & strtab & _
                    " where 病种名称 like '%" & Trim(txtDiseaseName.Text) & "%' or 病种编码 like '%" & _
                    Trim(txtDiseaseName.Text) & "%' or Upper(拼音码) like '%" & _
                    UCase(Trim(txtDiseaseName.Text)) & "%' "
    Else
        strTmpSQL = "select ID,病种编码,病种名称,拼音码  from " & strtab
    End If
    Call WriteBusinessLOG("strSQL", "", strTmpSQL)
    
    Call WriteBusinessLOG("ShowSelect 开始", "", "")
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "病种", True, , , , False, gcn铜山县)
    Call WriteBusinessLOG("ShowSelect 结束", "", "")
    
    If rsTmp Is Nothing Then Exit Sub
    txtDiseaseName.Text = "（" & Trim(rsTmp!病种编码) & "）" & Trim(rsTmp!病种名称)
    txtDiseaseName.Tag = rsTmp!ID
    
End Sub

Private Sub cmd疾病信息_Click()
    Call 病种选择(0)
End Sub

Private Sub txt个人编号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub



Private Sub tsx_取基本信息()
    Dim str个人编号 As String, str医保卡序号 As String
    Dim str姓名 As String, str身份证 As String, str性别 As String, str出生年月 As String
    Dim lng年龄 As Long, str工作单位 As String, str人员特征 As String, str特殊人员 As String
    Dim dbl帐户余额 As Double, str特殊就诊人员 As String, str出生日期 As String
    str姓名 = Space(12): str身份证 = Space(20): str性别 = Space(6)
    str出生年月 = Space(6): str工作单位 = Space(50): str人员特征 = Space(10)
    str特殊人员 = Space(30)
    '1
    If Trim(txt个人编号.Text) = "" Then
        MsgBox "请输入个人编号", vbInformation, gstrSysName
        txt个人编号.SetFocus
        Exit Sub
    End If
    If Trim(txt医保卡序号.Text) = "" Then
        MsgBox "请输入医保卡序号", vbInformation, gstrSysName
        txt医保卡序号.SetFocus
        Exit Sub
    End If
        
    If tsx_createparams(1024, 1024) = -1 Then
         MsgBox "分配内存空间失败!", vbInformation, gstrSysName
         Exit Sub
    End If
    Call WriteBusinessLOG("1 tsx_createparams", "1024,1204", mlngReturn)
    '2
    str个人编号 = txt个人编号.Text
    mlngReturn = tsx_setstringparam(P_TBR, 0, txt个人编号.Text) '个人编号
    Call WriteBusinessLOG("2 tsx_setstringparam", "P_TBR" & ", 0," & str个人编号, mlngReturn)
    mlngReturn = tsx_setstringparam(P_LB, 0, "0") '查询类别    业务类别暂为'0'
    Call WriteBusinessLOG("2 tsx_setstringparam", "P_LB" & ", 0,'0'", mlngReturn)
    '3
    If tsx_jkcall("GETCBRYXX_T") = -1 Then
        Call WriteBusinessLOG("3 tsx_jkcall", "GETCBRYXX_T", mlngReturn)
        MsgBox tsx_getlasterr(), vbInformation, gstrSysName
        mlngReturn = tsx_destroyparams
        Call WriteBusinessLOG("5 tsx_destroyparams", "", mlngReturn)
        Exit Sub
    Else
        Call WriteBusinessLOG("3 tsx_jkcall", "GETCBRYXX_T", mlngReturn)
    '4 取返回值
        mlngReturn = tsx_getstringparam(P_XM, 0, str姓名)
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_XM, 0, " & str姓名, mlngReturn)
        mlngReturn = tsx_getstringparam(P_SHBZH, 0, str身份证) ' C20 公民身份证号
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_SHBZH, 0, " & str身份证, mlngReturn)
        mlngReturn = tsx_getstringparam(P_XB, 0, str性别) '    C6  性别    女/男/未知
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_XB, 0, " & str性别, mlngReturn)
        mlngReturn = tsx_getstringparam(P_CSNY, 0, str出生年月) '  C6  出生年月    格式(YYYYMM)
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_CSNY, 0, " & str出生年月, mlngReturn)
        mlngReturn = tsx_getlongparam(P_NL, 0, lng年龄)  '    L   年龄
        Call WriteBusinessLOG("4 tsx_getlongparam", "P_NL, 0, " & lng年龄, mlngReturn)
        mlngReturn = tsx_getstringparam(P_DWMCH, 0, str工作单位) ' C50 工作单位
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_DWMCH, 0, " & str工作单位, mlngReturn)
        mlngReturn = tsx_getstringparam(P_RYTZ, 0, str人员特征) '  C10 人员特征    0-在职,1-退休
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_RYTZ, 0, " & str人员特征, mlngReturn)
        mlngReturn = tsx_getstringparam(P_TSRY, 0, str特殊人员) '  C30 特殊人员    0-普通,L=离休,E=二乙,G1=公务员,G2参照公务员
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_TSRY, 0, " & str特殊人员, mlngReturn)
        mlngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl帐户余额)  '   D   个人帐户余额    当前个人帐户余额
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_QCGRZH, 0, " & dbl帐户余额, mlngReturn)
        mlngReturn = tsx_getstringparam(P_TSJZ, 0, str特殊就诊人员) '  C2      特殊就诊人员(0-普通,1-门慢,2-门特
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_TSJZ, 0, " & str特殊就诊人员, mlngReturn)
        txt姓名.Text = Trim(str姓名)
        txt身份证.Text = Trim(str身份证)
        txt性别.Text = Trim(str性别)
        
        If 身份证号转出生日期(Trim(str身份证), str出生日期) Then
            txt出生年月.Text = Mid(Trim(str出生日期), 1, 4) & "-" & Mid(Trim(str出生日期), 5, 2) & "-" & Mid(Trim(str出生日期), 7, 2)
        Else
            txt出生年月.Text = Mid(Trim(str出生年月), 1, 4) & "-" & Mid(Trim(str出生年月), 5) & "-" & "01"
        End If
        
        txt年龄.Text = lng年龄
        txt工作单位.Text = Trim(str工作单位)
        
        'Select Case Trim(str人员特征)
        '    Case "0"
        '        txt在职.Text = "在职"
        '    Case "1"
        '        txt在职.Text = "退休"
        'End Select
        
        txt在职.Text = Trim(str人员特征)
        'txt人员身份.Tag = Trim(str特殊人员)
        txt人员身份.Text = Trim(str特殊人员)
'        Select Case Trim(str特殊人员)
'            Case "0"
'                txt人员身份.Text = "普通"
'
'            Case "L"
'                txt人员身份.Text = "离休"
'            Case "E"
'                txt人员身份.Text = "二乙"
'            Case "G1"
'                txt人员身份.Text = "公务员"
'            Case "G2"
'                txt人员身份.Text = "参照公务员"
'        End Select
        txt余额.Text = Format(dbl帐户余额, "0.00")
        txt特诊.Text = Val(str特殊就诊人员)
'        txt特诊.Tag = Trim(str特殊就诊人员)
'        Select Case Trim(str特殊就诊人员)
'            Case "0"
'                txt特诊.Text = "普通"
'
'            Case "1"
'                txt特诊.Text = "门慢"
'            Case "2"
'                txt特诊.Text = "门特"
'        End Select
        
    End If
    '5 销毁已分配空间
    mlngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("5 tsx_destroyparams", "", mlngReturn)
    OKButton.Enabled = True
End Sub

Private Sub txt医保卡序号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub txt医保卡序号_LostFocus()
    Call tsx_取基本信息
End Sub


Private Function Read(ByVal intCOM As Integer) As String
    Dim lngReturn As Integer, strReturn As String
    
    strReturn = "无信息"
    lngReturn = init_com(intCOM)
    Call WriteBusinessLOG("init_com", intCOM, lngReturn)
    If lngReturn <> 0 Then
        MsgBox "初始化端口错误", vbInformation, "读卡"
        Exit Function
    End If
    
    lngReturn = sele_card(43)
    Call WriteBusinessLOG("sele_card", 43, lngReturn)
    
    If lngReturn <> 0 Then
        MsgBox "定义卡类型错误", vbInformation, "读卡"
        GoTo powerOFF
    End If
    
    If power_on() <> 0 Then
        MsgBox "卡上电错误", vbInformation, "读卡"
        GoTo powerOFF
    End If
    
    strReturn = Space(129)
    lngReturn = rd_str(1, 0, 128, strReturn)
   
    If lngReturn <> 0 Then
        MsgBox "读取卡信息错误", vbInformation, "读卡"
        GoTo powerOFF
    End If

powerOFF:
    Call power_off
    Call close_com
    Read = Split(strReturn, "@")(0) & ";" & Split(strReturn, "@")(2)
End Function


