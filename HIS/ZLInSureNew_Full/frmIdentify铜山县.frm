VERSION 5.00
Begin VB.Form frmIdentifyͭɽ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   4050
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5370
   Icon            =   "frmIdentifyͭɽ��.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chk���� 
      Caption         =   "����"
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
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Height          =   300
      Left            =   4770
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2670
      Width           =   285
   End
   Begin VB.TextBox txtҽ������� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4050
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3735
      TabIndex        =   28
      Top             =   2175
      Width           =   1335
   End
   Begin VB.TextBox txt��� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   26
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txt��Ա��� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3735
      TabIndex        =   24
      Top             =   1740
      Width           =   1335
   End
   Begin VB.TextBox txt��ְ 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   22
      Top             =   1740
      Width           =   1335
   End
   Begin VB.TextBox txt������λ 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   20
      Top             =   1335
      Width           =   3900
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4605
      TabIndex        =   18
      Top             =   525
      Width           =   450
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4065
      TabIndex        =   16
      Top             =   930
      Width           =   1005
   End
   Begin VB.TextBox txt�Ա� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3105
      TabIndex        =   14
      Top             =   525
      Width           =   465
   End
   Begin VB.TextBox txt���֤ 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   12
      Top             =   930
      Width           =   1785
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   10
      Top             =   525
      Width           =   915
   End
   Begin VB.TextBox txt���˱�� 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdReadCard 
      Caption         =   "����(&R)"
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
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   3945
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   2175
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "����ѡ��"
      Height          =   210
      Left            =   270
      TabIndex        =   32
      Top             =   2745
      Width           =   900
   End
   Begin VB.Label Label12 
      Caption         =   "ҽ������ţ�"
      Height          =   255
      Left            =   2970
      TabIndex        =   30
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label Label11 
      Caption         =   "������Ա��"
      Height          =   195
      Left            =   2790
      TabIndex        =   29
      Top             =   2190
      Width           =   915
   End
   Begin VB.Label Label10 
      Caption         =   "�ʻ���"
      Height          =   195
      Left            =   255
      TabIndex        =   27
      Top             =   2190
      Width           =   915
   End
   Begin VB.Label Label9 
      Caption         =   "��Ա��ݣ�"
      Height          =   195
      Left            =   2790
      TabIndex        =   25
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label8 
      Caption         =   "��ְ״̬��"
      Height          =   195
      Left            =   255
      TabIndex        =   23
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "������λ��"
      Height          =   195
      Left            =   255
      TabIndex        =   21
      Top             =   1410
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "���䣺"
      Height          =   195
      Left            =   4080
      TabIndex        =   19
      Top             =   585
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "�������£�"
      Height          =   195
      Left            =   3135
      TabIndex        =   17
      Top             =   1005
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "�Ա�"
      Height          =   195
      Left            =   2565
      TabIndex        =   15
      Top             =   585
      Width           =   540
   End
   Begin VB.Label Label3 
      Caption         =   "���֤�ţ�"
      Height          =   195
      Left            =   255
      TabIndex        =   13
      Top             =   1005
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "������"
      Enabled         =   0   'False
      Height          =   195
      Left            =   615
      TabIndex        =   11
      Top             =   585
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "���˱�ţ�"
      Height          =   255
      Left            =   255
      TabIndex        =   9
      Top             =   150
      Width           =   915
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "IC�����ڣ�"
      Height          =   180
      Index           =   3
      Left            =   285
      TabIndex        =   8
      Top             =   3135
      Width           =   900
   End
End
Attribute VB_Name = "frmIdentifyͭɽ��"
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
Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�
Private mlng����ID As Long
Private mstrReturn As String

Public Function GetIdentify(ByVal bytType As Byte, Optional ByVal lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    
    Me.Show 1
    lng����ID = mlng����ID
    GetIdentify = mstrReturn
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub



Private Sub chk����_LostFocus()
    If chk����.Value = 2 Then chk����.Value = 1
End Sub

Private Sub cmdReadCard_Click()
    Dim str���˱�� As String, strҽ������� As String
    Dim str���� As String, str���֤ As String, str�Ա� As String, str�������� As String
    Dim lng���� As Long, str������λ As String, str��Ա���� As String, str������Ա As String
    Dim dbl�ʻ���� As Double, str���������Ա As String, str�������� As String
    Dim intCOM As Long, strRead As String
    
    
    str���˱�� = Space(9): strҽ������� = Space(3)
    mlngReturn = tsx_read_ic(str���˱��, strҽ�������)
    Call WriteBusinessLOG("tsx_read_ic", str���˱�� & "," & strҽ�������, mlngReturn)
    If mlngReturn <> -1 Then
        txt���˱��.Text = str���˱��
        txtҽ�������.Text = strҽ�������
        Call tsx_ȡ������Ϣ
    Else
        MsgBox tsx_getlasterr(), vbInformation, gstrSysName
    End If
    

End Sub

Private Sub cmdReadCard_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub Form_Load()
    txtCom.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", 0) + 1
    OKButton.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", CStr(txtCom.Text - 1)
End Sub

Private Sub OKButton_Click()
    Dim strEmpInfo As String
    Dim strAccinfo As String

    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    
    '����ѡ��û��
    If txtDiseaseName.Tag = "" Then
         MsgBox "��ѡ���֣�", vbInformation, gstrSysName
         mstrReturn = ""
         txtDiseaseName.SetFocus
         Exit Sub
    End If
    
    strEmpInfo = txtҽ�������.Text                               '0����
    strEmpInfo = strEmpInfo & ";" & txt���˱��.Text              '1ҽ����
    strEmpInfo = strEmpInfo & ";"              '2����
    strEmpInfo = strEmpInfo & ";" & txt����.Text                '3����
    strEmpInfo = strEmpInfo & ";" & txt�Ա�.Text                '4�Ա�
    strEmpInfo = strEmpInfo & ";" & txt��������.Text         '5��������
    strEmpInfo = strEmpInfo & ";" & txt���֤.Text           '6���֤
    strEmpInfo = strEmpInfo & ";" & txt������λ.Text           '7.��λ����(����)
    
    strAccinfo = ";0"                                          '8.���Ĵ���
    strAccinfo = strAccinfo & ";"                    '9.˳���
    strAccinfo = strAccinfo & ";"              '10��Ա���
    strAccinfo = strAccinfo & ";" & Val(txt���.Text)        '11�ʻ����
    strAccinfo = strAccinfo & ";"     ' & g���˻�����Ϣ.��Ժ״̬16                             '12��ǰ״̬
    strAccinfo = strAccinfo & ";" & txtDiseaseName.Tag                   '13����ID
    strAccinfo = strAccinfo & ";"                           '14��ְ(1,2,3)
    strAccinfo = strAccinfo & ";"                             '15����֤��
    strAccinfo = strAccinfo & ";"                             '16�����
    strAccinfo = strAccinfo & ";1"                            '17�Ҷȼ�
    strAccinfo = strAccinfo & ";" & Val(txt���.Text)      '18�ʻ������ۼ�
    strAccinfo = strAccinfo & ";0"                              '19�ʻ�֧���ۼ�
    strAccinfo = strAccinfo & ";0"                            '20���깤���ܶ�
    strAccinfo = strAccinfo & ";"      '21
    strAccinfo = strAccinfo & ";"       '22סԺ�����ۼ�
    
    mlng����ID = BuildPatiInfo(mbytType, strEmpInfo & strAccinfo, mlng����ID, TYPE_ͭɽ��)
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_ͭɽ�� & ",'�������','''" & txt����.Text & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "Ӧ�����")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_ͭɽ�� & ",'��Ա���','''" & txt��Ա���.Text & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "�ʻ�״̬")
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strEmpInfo & ";" & mlng����ID & strAccinfo
        g������Ϣ_ͭɽ.����ID = mlng����ID
        g������Ϣ_ͭɽ.���˱�� = txt���˱��.Text
        g������Ϣ_ͭɽ.ҽ������� = txtҽ�������.Text
        g������Ϣ_ͭɽ.���ֱ��� = Mid(txtDiseaseName.Text, 2, InStr(txtDiseaseName.Text, "��") - 2)
        g������Ϣ_ͭɽ.����� = chk����.Value
    End If
    Unload Me
End Sub

Private Sub txtCom_Change()
    If InStr("123456789", txtCom.Text) <= 0 Then
        MsgBox "����������!", vbInformation, gstrSysName
        txtCom.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", 0) + 1
    End If
End Sub

Private Sub txtDiseaseName_KeyPress(KeyAscii As Integer)
  ''������ѡ����
    Call WriteBusinessLOG("KeyAscii", "", KeyAscii)

    If KeyAscii <> vbKeyReturn Then Exit Sub
    Call WriteBusinessLOG("KeyAscii", "", KeyAscii)
    If Trim(txtDiseaseName.Text) = "" Then Exit Sub
    Call WriteBusinessLOG("����ѡ�� ��ʼ", "", "")
    
    Call ����ѡ��
    Call WriteBusinessLOG("����ѡ�� ����", "", "")
    
End Sub

Private Sub ����ѡ��(Optional strLoad As String = 1)
    Dim rsTmp As ADODB.Recordset, strtab As String
    Dim strTmpSQL As String
    
    If mbytType = 0 Then
        strtab = "MZBZ"
    Else
        strtab = "ICD10"
    End If
    
    If strLoad = 1 Then
        strTmpSQL = "select ID,���ֱ���,��������,ƴ���� from " & strtab & _
                    " where �������� like '%" & Trim(txtDiseaseName.Text) & "%' or ���ֱ��� like '%" & _
                    Trim(txtDiseaseName.Text) & "%' or Upper(ƴ����) like '%" & _
                    UCase(Trim(txtDiseaseName.Text)) & "%' "
    Else
        strTmpSQL = "select ID,���ֱ���,��������,ƴ����  from " & strtab
    End If
    Call WriteBusinessLOG("strSQL", "", strTmpSQL)
    
    Call WriteBusinessLOG("ShowSelect ��ʼ", "", "")
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "����", True, , , , False, gcnͭɽ��)
    Call WriteBusinessLOG("ShowSelect ����", "", "")
    
    If rsTmp Is Nothing Then Exit Sub
    txtDiseaseName.Text = "��" & Trim(rsTmp!���ֱ���) & "��" & Trim(rsTmp!��������)
    txtDiseaseName.Tag = rsTmp!ID
    
End Sub

Private Sub cmd������Ϣ_Click()
    Call ����ѡ��(0)
End Sub

Private Sub txt���˱��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub



Private Sub tsx_ȡ������Ϣ()
    Dim str���˱�� As String, strҽ������� As String
    Dim str���� As String, str���֤ As String, str�Ա� As String, str�������� As String
    Dim lng���� As Long, str������λ As String, str��Ա���� As String, str������Ա As String
    Dim dbl�ʻ���� As Double, str���������Ա As String, str�������� As String
    str���� = Space(12): str���֤ = Space(20): str�Ա� = Space(6)
    str�������� = Space(6): str������λ = Space(50): str��Ա���� = Space(10)
    str������Ա = Space(30)
    '1
    If Trim(txt���˱��.Text) = "" Then
        MsgBox "��������˱��", vbInformation, gstrSysName
        txt���˱��.SetFocus
        Exit Sub
    End If
    If Trim(txtҽ�������.Text) = "" Then
        MsgBox "������ҽ�������", vbInformation, gstrSysName
        txtҽ�������.SetFocus
        Exit Sub
    End If
        
    If tsx_createparams(1024, 1024) = -1 Then
         MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
         Exit Sub
    End If
    Call WriteBusinessLOG("1 tsx_createparams", "1024,1204", mlngReturn)
    '2
    str���˱�� = txt���˱��.Text
    mlngReturn = tsx_setstringparam(P_TBR, 0, txt���˱��.Text) '���˱��
    Call WriteBusinessLOG("2 tsx_setstringparam", "P_TBR" & ", 0," & str���˱��, mlngReturn)
    mlngReturn = tsx_setstringparam(P_LB, 0, "0") '��ѯ���    ҵ�������Ϊ'0'
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
    '4 ȡ����ֵ
        mlngReturn = tsx_getstringparam(P_XM, 0, str����)
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_XM, 0, " & str����, mlngReturn)
        mlngReturn = tsx_getstringparam(P_SHBZH, 0, str���֤) ' C20 �������֤��
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_SHBZH, 0, " & str���֤, mlngReturn)
        mlngReturn = tsx_getstringparam(P_XB, 0, str�Ա�) '    C6  �Ա�    Ů/��/δ֪
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_XB, 0, " & str�Ա�, mlngReturn)
        mlngReturn = tsx_getstringparam(P_CSNY, 0, str��������) '  C6  ��������    ��ʽ(YYYYMM)
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_CSNY, 0, " & str��������, mlngReturn)
        mlngReturn = tsx_getlongparam(P_NL, 0, lng����)  '    L   ����
        Call WriteBusinessLOG("4 tsx_getlongparam", "P_NL, 0, " & lng����, mlngReturn)
        mlngReturn = tsx_getstringparam(P_DWMCH, 0, str������λ) ' C50 ������λ
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_DWMCH, 0, " & str������λ, mlngReturn)
        mlngReturn = tsx_getstringparam(P_RYTZ, 0, str��Ա����) '  C10 ��Ա����    0-��ְ,1-����
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_RYTZ, 0, " & str��Ա����, mlngReturn)
        mlngReturn = tsx_getstringparam(P_TSRY, 0, str������Ա) '  C30 ������Ա    0-��ͨ,L=����,E=����,G1=����Ա,G2���չ���Ա
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_TSRY, 0, " & str������Ա, mlngReturn)
        mlngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl�ʻ����)  '   D   �����ʻ����    ��ǰ�����ʻ����
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_QCGRZH, 0, " & dbl�ʻ����, mlngReturn)
        mlngReturn = tsx_getstringparam(P_TSJZ, 0, str���������Ա) '  C2      ���������Ա(0-��ͨ,1-����,2-����
        Call WriteBusinessLOG("4 tsx_getstringparam", "P_TSJZ, 0, " & str���������Ա, mlngReturn)
        txt����.Text = Trim(str����)
        txt���֤.Text = Trim(str���֤)
        txt�Ա�.Text = Trim(str�Ա�)
        
        If ���֤��ת��������(Trim(str���֤), str��������) Then
            txt��������.Text = Mid(Trim(str��������), 1, 4) & "-" & Mid(Trim(str��������), 5, 2) & "-" & Mid(Trim(str��������), 7, 2)
        Else
            txt��������.Text = Mid(Trim(str��������), 1, 4) & "-" & Mid(Trim(str��������), 5) & "-" & "01"
        End If
        
        txt����.Text = lng����
        txt������λ.Text = Trim(str������λ)
        
        'Select Case Trim(str��Ա����)
        '    Case "0"
        '        txt��ְ.Text = "��ְ"
        '    Case "1"
        '        txt��ְ.Text = "����"
        'End Select
        
        txt��ְ.Text = Trim(str��Ա����)
        'txt��Ա���.Tag = Trim(str������Ա)
        txt��Ա���.Text = Trim(str������Ա)
'        Select Case Trim(str������Ա)
'            Case "0"
'                txt��Ա���.Text = "��ͨ"
'
'            Case "L"
'                txt��Ա���.Text = "����"
'            Case "E"
'                txt��Ա���.Text = "����"
'            Case "G1"
'                txt��Ա���.Text = "����Ա"
'            Case "G2"
'                txt��Ա���.Text = "���չ���Ա"
'        End Select
        txt���.Text = Format(dbl�ʻ����, "0.00")
        txt����.Text = Val(str���������Ա)
'        txt����.Tag = Trim(str���������Ա)
'        Select Case Trim(str���������Ա)
'            Case "0"
'                txt����.Text = "��ͨ"
'
'            Case "1"
'                txt����.Text = "����"
'            Case "2"
'                txt����.Text = "����"
'        End Select
        
    End If
    '5 �����ѷ���ռ�
    mlngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("5 tsx_destroyparams", "", mlngReturn)
    OKButton.Enabled = True
End Sub

Private Sub txtҽ�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub

Private Sub txtҽ�������_LostFocus()
    Call tsx_ȡ������Ϣ
End Sub


Private Function Read(ByVal intCOM As Integer) As String
    Dim lngReturn As Integer, strReturn As String
    
    strReturn = "����Ϣ"
    lngReturn = init_com(intCOM)
    Call WriteBusinessLOG("init_com", intCOM, lngReturn)
    If lngReturn <> 0 Then
        MsgBox "��ʼ���˿ڴ���", vbInformation, "����"
        Exit Function
    End If
    
    lngReturn = sele_card(43)
    Call WriteBusinessLOG("sele_card", 43, lngReturn)
    
    If lngReturn <> 0 Then
        MsgBox "���忨���ʹ���", vbInformation, "����"
        GoTo powerOFF
    End If
    
    If power_on() <> 0 Then
        MsgBox "���ϵ����", vbInformation, "����"
        GoTo powerOFF
    End If
    
    strReturn = Space(129)
    lngReturn = rd_str(1, 0, 128, strReturn)
   
    If lngReturn <> 0 Then
        MsgBox "��ȡ����Ϣ����", vbInformation, "����"
        GoTo powerOFF
    End If

powerOFF:
    Call power_off
    Call close_com
    Read = Split(strReturn, "@")(0) & ";" & Split(strReturn, "@")(2)
End Function


