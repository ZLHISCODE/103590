Attribute VB_Name = "Mdlͭɽ��"
Option Explicit
#Const gverControl = 99
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


'���������淶:ȫ�ֱ�����g��ͷ,ģ�鼶������m��ͷ
'����һ��ģ�鼶������������¼���ӿ��Ƿ���������ʼ�������ⱻ�ظ���γ�ʼ��
Private mblnInit As Boolean
Private mblnICinit As Boolean
'����һ��ȫ�����Ӷ���
Public gcnͭɽ�� As New ADODB.Connection
'API��������ʾ��
Private Const gstrSysName = "�������"
'=========================================================================

Public Declare Function tsx_init_ic Lib "LesTsybjk.dll" Alias "init_ic" _
 (ByVal com As Long) As Long
 
Public Declare Function tsx_read_ic Lib "LesTsybjk.dll" Alias "read_ic" _
 (ByVal tbr As String, ByVal kxh As String) As Long

Public Declare Function tsx_exit_ic Lib "LesTsybjk.dll" Alias "exit_ic" _
 () As Long

'=========================================================================
Public Declare Function tsx_conn_ybzx Lib "LesTsybjk.dll" Alias "conn_ybzx" _
 (ByVal jgm As String, ByVal ryh As String, ByVal mm As String, _
  ByVal res As String) As Long
  
Public Declare Function tsx_disconn_ybzx Lib "LesTsybjk.dll" Alias "disconn_ybzx" _
 () As Long
 
'1
Public Declare Function tsx_createparams Lib "LesTsybjk.dll" Alias "createparams" _
(ByVal sendlen As Long, ByVal recvlen As Long) As Long
'2
Public Declare Function tsx_destroyparams Lib "LesTsybjk.dll" Alias "destroyparams" _
 () As Long
'3
Public Declare Function tsx_setstringparam Lib "LesTsybjk.dll" Alias "setstringparam" _
 (ByVal paramid As Long, ByVal Row As Long, ByVal Value As String) As Long
'4
Public Declare Function tsx_setlongparam Lib "LesTsybjk.dll" Alias "setlongparam" _
 (ByVal paramid As Long, ByVal Row As Long, ByVal Value As Long) As Long
'5
Public Declare Function tsx_setdoubleparam Lib "LesTsybjk.dll" Alias "setdoubleparam" _
 (ByVal paramid As Long, ByVal Row As Long, ByVal Value As Double) As Long
'6
Public Declare Function tsx_jkcall Lib "LesTsybjk.dll" Alias "jkcall" _
 (ByVal svcname As String) As Long
'7
Public Declare Function tsx_getstringparam Lib "LesTsybjk.dll" Alias "getstringparam" _
 (ByVal paramid As Long, ByVal Row As Long, ByVal Value As String) As Long
'8
Public Declare Function tsx_getlongparam Lib "LesTsybjk.dll" Alias "getlongparam" _
 (ByVal paramid As Long, ByVal Row As Long, ByRef Value As Long) As Long
'9
Public Declare Function tsx_getdoubleparam Lib "LesTsybjk.dll" Alias "getdoubleparam" _
 (ByVal paramid As Long, ByVal Row As Long, ByRef Value As Double) As Long
'10
Public Declare Function tsx_getrowcount Lib "LesTsybjk.dll" Alias "getrowcount" _
 (ByVal paramid As Long) As Long
'11
Public Declare Function tsx_getlasterr2 Lib "LesTsybjk.dll" Alias "getlasterr2" _
 (ByVal str_err As String) As Long
 

'Public Declare Function BJ_Hosp_Divide3 Lib "FYFJ.dll" Alias "Hosp_Divide3" (ByVal strIn As String) As Long
'������"TODO:�������ѵ�ʵ�ִ���"���ҵ��������㣬��Щ����㶼�Ǳ�����д�����
'TODO:��������
'�����ڴ����Ӵ��룬ʵ��XX����
'-------------------------------------------------------------------------------
'��̲���˵��
'1��Ϊ���ӿڲ�������������zl9I_xxx���籱��ҽ������������Ϊ��zl9I_BJYB��ע�⣬��ģ����Ҫ����Ϊ��clsI_xxx
'2�������Ҫ��������ҽ����ص����ݣ����½�һ���û����������ǳ�֮Ϊ�м��
'3����ҽ����صĲ������ã����м����û��������������������������ӱ��ղ������ô��壬��������frmSetҽ�����ƣ��磺frmSet������
'4����������ṩ����ҽ����Ŀ�嵥������Ŀ¼�ȣ����ڱ�����Ŀѡ�����Ŀ���°�ť����д���룬��ɴ��ļ������Ľ�����·����ݸ��µ�HIS����
'5����д�������ҽ����Ŀ����Ĺ���
'6����д������������֤����
'7���������º�������̵�������룬���ҽ���ӿڵ����幦��
'8�����ݽӿ����ʣ��޸���ģ����GetCapability()��������ز�����μ�mdlInsure�е�ö�ٱ���"ҽԺҵ��"
'9��������Ҫ�޸���ģ�������������ĵ��ô���
'10��������Ҫ���ӻ��޸Ĺ��������ģ��
'-------------------------------------------------------------------------------

Private Type ������Ϣ_ͭɽ
    Com��           As Long
    ����ID          As Long
    ҽ�������      As String
    ���˱��        As String
    ҽԺ��          As String
    ����Ա��        As String
    ����Ա����      As String
    ���ֱ���        As String
    �����          As String * 1
End Type

Public g������Ϣ_ͭɽ As ������Ϣ_ͭɽ
Dim mstrʹ�ø����ʻ�֧�� As String

'>>������
'Public clsTst As New clsLesybjk
'>>

Public Function ҽ����ʼ��_ͭɽ��(Optional ByVal blnTest As Boolean = False) As Boolean
'*****************************************************************************
'�����ߡ���������������clsInsure ��  InitInsure  ���̵���(�÷����ɷ��ò��������Ժ�������ҽ���ӿ���صĲ�������)
'����˵�����������������ҽ���ӿڳ�ʼ����صĹ�����ȫ�ֱ����ĳ�ʼ�������������ĳ�ʼ���ȣ�
'���ù����嵥��˵����
'    ��tsx_conn_ybzx������ҽ������
'������tsx_createparams������ռ�
'������tsx_setstringparam������string�Ͳ���
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_getstringparam��ȡstring�ͽӿڷ���ֵ
'������tsx_destroyparams�������ѷ���ռ�
'
''*****************************************************************************
    'TODO:ҽ����ʼ
    '�����ǲο�����
    Dim strUser As String, strServer As String, strPass As String
    Dim rsTemp As New ADODB.Recordset, lngReturn As Long, rsCsh As New ADODB.Recordset
    Dim strReturn As String, intCOM As Long, STRERR As String
    On Error GoTo errHand

    If mblnInit = False Then
        '��������ҽ��������������
        strUser = "tsxyb"
        strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
        strPass = "tsxyb"
        g������Ϣ_ͭɽ.Com�� = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", 0) + 1
'
    intCOM = g������Ϣ_ͭɽ.Com�� - 1
    Select Case intCOM
        Case 0
             lngReturn = tsx_init_ic(0)
        Case 1
             lngReturn = tsx_init_ic(1)
        Case 2
             lngReturn = tsx_init_ic(2)
        Case Else
             lngReturn = tsx_init_ic(3)
    End Select
    
    Call WriteBusinessLOG("tsx_init_ic", g������Ϣ_ͭɽ.Com��, lngReturn)
    If lngReturn <> -1 Then
        mblnICinit = True
    End If
''
        'ȡҽԺ����
        gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_ͭɽ��)
        g������Ϣ_ͭɽ.ҽԺ�� = Nvl(rsTemp!ҽԺ����)
        
        If OraDataOpen(gcnͭɽ��, strServer, strUser, strPass, False) = False Then
            MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
            Exit Function
        Else
            gstrSQL = "Select * from czry Where ��ԱID=" & UserInfo.ID
            Call OpenRecordset_OtherBase(rsTemp, "ͭɽ��ҽ��", gstrSQL, gcnͭɽ��)
            
            If rsTemp.EOF Then
                gstrSQL = "Select * from czry Where P_GLY=1"
                Call OpenRecordset_OtherBase(rsCsh, "ͭɽ��ҽ��", gstrSQL, gcnͭɽ��)
                
                If rsCsh.EOF Then
                    MsgBox "��Ҫ����Ա�ʻ���", vbInformation, gstrSysName
                    Exit Function
                End If
                g������Ϣ_ͭɽ.����Ա�� = rsCsh!P_RYH
                g������Ϣ_ͭɽ.����Ա���� = rsCsh!P_MM
                
                lngReturn = tsx_conn_ybzx(g������Ϣ_ͭɽ.ҽԺ��, g������Ϣ_ͭɽ.����Ա��, g������Ϣ_ͭɽ.����Ա����, "")
                Call WriteBusinessLOG("tsx_conn_ybzx", g������Ϣ_ͭɽ.ҽԺ�� & "," & g������Ϣ_ͭɽ.����Ա�� & "," & g������Ϣ_ͭɽ.����Ա����, lngReturn)
                
                If lngReturn = -1 Then
                    MsgBox "ҽ����ʼʧ�ܣ�" & vbCrLf & tsx_getlasterr(), vbInformation, gstrSysName
                    Exit Function
                End If

                '> Beging �м�����޴˲���Ա��Ϣ,���ýӿ�������Ա
                '1�����ڴ�
               If tsx_createparams(1024, 1024) = -1 Then
                    MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
                    Exit Function
               End If
               Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
               '2 Ϊ������ֵ
               lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
               Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", 0," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
               lngReturn = tsx_setstringparam(P_RYH, 0, "")
               Call WriteBusinessLOG("tsx_setstringparam", "P_RYH" & ", 0,''", lngReturn)
               lngReturn = tsx_setstringparam(P_XM, 0, UserInfo.����) '    C12 ����    סԺ����ʱ���ɵ�Ψһ��
               Call WriteBusinessLOG("tsx_setstringparam", "P_XM" & ", 0," & UserInfo.����, lngReturn)
               lngReturn = tsx_setstringparam(P_MM, 0, UserInfo.���� & "001") '    C10 ����    ��Աע�ᵽҽ�����ĵĲ�������
               Call WriteBusinessLOG("tsx_setstringparam", "P_MM" & ", 0," & UserInfo.����, lngReturn)
               lngReturn = tsx_setstringparam(P_GLY, 0, "0") '   C1  ����Ա  0-����1-��
               Call WriteBusinessLOG("tsx_setstringparam", "P_GLY" & ", 0,'0'", lngReturn)
               lngReturn = tsx_setstringparam(P_LB, 0, 1) ' C1  �������    0-ע����Ա(���ָܻ�)1-������Ա 2-�޸����������
               Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", 0,'1'", lngReturn)
               
               '3 ���ýӿ�
               If tsx_jkcall("CZRYWH") = -1 Then
                    Call WriteBusinessLOG("jkcall", "CZRYWH", -1)
                    STRERR = tsx_getlasterr()
                    MsgBox "�����Աʧ�ܣ�" & vbCrLf & STRERR, vbInformation, gstrSysName
                    Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
                    lngReturn = tsx_destroyparams()
                    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
                    Exit Function
               Else
               Call WriteBusinessLOG("jkcall", "CZRYWH", lngReturn)
               '4 ȡ����ֵ
                    strPass = Space(10)
                    lngReturn = tsx_getstringparam(P_RYH, 0, strPass)
                    Call WriteBusinessLOG("tsx_getstringparam", "P_RYH, 0, " & g������Ϣ_ͭɽ.����Ա��, lngReturn)
                    
               End If
               '5 �����ѷ���ռ�
               lngReturn = tsx_destroyparams()
               Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
               
               gstrSQL = "insert into czry(��ԱID,P_JGM,P_RYH,P_XM,P_MM) values(" & _
                                    UserInfo.ID & ",'" & g������Ϣ_ͭɽ.ҽԺ�� & "','" & _
                                    strPass & "','" & UserInfo.���� & "','" & _
                                    UserInfo.���� & "001')"
               gcnͭɽ��.Execute gstrSQL
               
               g������Ϣ_ͭɽ.����Ա�� = strPass
               g������Ϣ_ͭɽ.����Ա���� = UserInfo.���� & "001"
               '> End �м�����޴˲���Ա��Ϣ,���ýӿ�������Ա
            Else
                g������Ϣ_ͭɽ.����Ա�� = rsTemp!P_RYH
                g������Ϣ_ͭɽ.����Ա���� = rsTemp!P_MM
                
                lngReturn = tsx_conn_ybzx(g������Ϣ_ͭɽ.ҽԺ��, g������Ϣ_ͭɽ.����Ա��, g������Ϣ_ͭɽ.����Ա����, "")
                Call WriteBusinessLOG("tsx_conn_ybzx", g������Ϣ_ͭɽ.ҽԺ�� & "," & g������Ϣ_ͭɽ.����Ա�� & "," & g������Ϣ_ͭɽ.����Ա����, lngReturn)
                
                If lngReturn = -1 Then
                    MsgBox "ҽ����ʼʧ�ܣ�" & vbCrLf & tsx_getlasterr(), vbInformation, gstrSysName
                    Exit Function
                End If
                
            End If
            

        End If


    End If

    Set rsTemp = Nothing
    Set rsCsh = Nothing
    mblnInit = True
    ҽ����ʼ��_ͭɽ�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ֹ_ͭɽ��() As Boolean

'*****************************************************************************
'�����ߡ���������������clsInsure ��  EndInsure  ���̵���(�÷����ɷ��ò��������Ժ�������ҽ���ӿ���صĲ�������)
'����˵��������������������ͷţ��Ͽ����ӵ�
'���ù����嵥��˵����
'������tsx_disconn_ybzx��ҽ�������ر�
''*****************************************************************************
    'TODO:ҽ����ֹ
    '�����ǲο�����
    'Dim strReturn As String
    'Call ���ýӿ�(ֹͣ���߻�����, strReturn)
    Dim lngReturn As Long
    On Error GoTo errHand
    
    If mblnInit = True Then
    
        If mblnICinit = True Then
            lngReturn = tsx_exit_ic()
            Call WriteBusinessLOG("tsx_exit_ic", "", lngReturn)
        End If
        
        Call tsx_disconn_ybzx
        Call WriteBusinessLOG("tsx_disconn_ybzx", "", "")
        
        mblnInit = False
    End If
    
    ҽ����ֹ_ͭɽ�� = True
    Exit Function
errHand:
    Call WriteBusinessLOG("ErrHand", Err.Number, Err.Description)
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function ��ݱ�ʶ_ͭɽ��(ByVal bytType As Byte, Optional lng����ID As Long = 0, Optional ByRef intinsure As Integer) As String
'*************************************************
'�����ߡ���������������clsInsure �� Identify  ���̵���(�÷�����������ò���������ҺŲ�������Ժ�Ǽǲ�������)
'����˵��������������ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ������֤�ɹ��󣬽�������Ϣ�����ظ���������
'��������������������bytType-ʶ�����ͣ�0-���1-סԺ
'���ء����������������ջ���Ϣ��
'ע�⡡��������������1)��Ҫ���ýӿڵ����ʶ���ף�
'��������������������2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'��������������������3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
'���ù����嵥��˵����
'�������ޡ�
'*************************************************
'TODO: �����֤
   ' Dim rsSfyz As New ADODB.Recordset
    Dim strReturn As String, lngReturn As Long, str�Һŵ��� As String
    
    strReturn = frmIdentifyͭɽ��.GetIdentify(bytType, lng����ID)
'        If bytType = 0 Then
'           '>Beging �������Ǽǹ���
'           '1
'            If tsx_createparams(1024, 1024) = -1 Then
'                MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
'                Exit Function
'            End If
'            Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
'            '2 Ϊ������ֵ
'            lngReturn= tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
'            Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", 0," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
'
'            lngReturn= tsx_setstringparam(P_KXH, 0, g������Ϣ_ͭɽ.ҽ�������) '   C3  ҽ�������
'            Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", 0," & g������Ϣ_ͭɽ.ҽ�������, lngReturn)
'            lngReturn= tsx_setstringparam(P_TBR, 0, g������Ϣ_ͭɽ.���˱��) '   C9  ���˱��
'            Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", 0," & g������Ϣ_ͭɽ.���˱��, lngReturn)
'            lngReturn= tsx_setstringparam(P_YYKSM, 0, "001") ' C20 �Һſ�����
'            Call WriteBusinessLOG("tsx_setstringparam", "P_YYKSM" & ", 0,'001'", lngReturn)
'            lngReturn= tsx_setdoubleparam(P_GHF, 0, 0)   '   D   �Һŷ�
'            Call WriteBusinessLOG("tsx_setdoubleparam", "P_GHF" & ", 0,0", lngReturn)
'
'            lngReturn= tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) '  C10 ������Ա���
'            Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH" & ", 0,'" & g������Ϣ_ͭɽ.����Ա�� & "'", lngReturn)
'            '3 ���ýӿ�
'            If jkcall("MZGH") = -1 Then
'                 Call WriteBusinessLOG("jkcall", "MZGH", lngReturn)
'                 MsgBox tsx_getlasterr(), vbInformation, gstrSysName

'                 lngReturn= tsx_destroyparams()
'                 Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
'                 Exit Function
'            Else
'            Call WriteBusinessLOG("jkcall", "MZGH", lngReturn)
'            '4 ȡ����ֵ
'                 lngReturn= tsx_getstringparam(P_DJH, 0, str�Һŵ���)
'                 Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str�Һŵ���, lngReturn)
'
'            End If
'            '5 �����ѷ���ռ�
'            lngReturn= tsx_destroyparams()
'            Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
'
'            '����str�Һŵ��ŵ������ʻ���
'
'
'           '>End �������Ǽǹ���
'        End If
    ��ݱ�ʶ_ͭɽ�� = strReturn
    
End Function

Public Function ҽ������_ͭɽ��(ByVal intinsure As Integer) As Boolean
'**************************************
'�����ߡ���������������clsInsure��CodeMan��1600���� ����
'����˵��������������ҽ����������
'���ù����嵥��˵����
'��������������
'**************************************

    'ҽ������_���� = frmSet����.��������()
End Function

Public Function ����Һ�_ͭɽ��(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'******************************************************************************
'�����ߡ��������������÷���������ҺŲ�������
'����˵��������������ͨ������ҽ���̵�����ҺŽӿڣ��ֽⱾ�η�����ϸ���õ��������������ʻ����١�ͳ�������ٵȣ�������
'ע�����������������Ҫ���ù���zl_���˽����¼_Update�Բ���Ԥ����¼������������
'���ù����嵥��˵����
'��������������
''*****************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str������ As String, dbl�Һŷ�  As Double, str������ˮ�� As String
    Dim lngReturn As Long, lngCounter As Long, dbl�����Է� As Double, STRERR As String, str��ˮ�� As String
    On Error GoTo ErrH
    strSQL = "Select b.���� as ������,Sum(A.ʵ�ս��) as ʵ�ս�� " & _
              " From ������ü�¼ A,���ű� B" & _
              " Where a.���˿���Id=b.id and A.����ID=[1] And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0 group by b.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ȡ�Һŷ�", lng����ID)
    str������ = "": dbl�Һŷ� = 0
    Do Until rsTmp.EOF
        str������ = Trim("" & rsTmp!������)
        dbl�Һŷ� = Val("" & rsTmp!ʵ�ս��)
        rsTmp.MoveNext
    Loop
    
    If tsx_createparams(102400, 102400) = -1 Then
        Err.Raise 9000, gstrSysName, "�����ڴ�ռ�ʧ��!"
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    
    'P_JGM   C5  ҽԺ��
    lngReturn = tsx_setstringparam(P_JGM, lngCounter, g������Ϣ_ͭɽ.ҽԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
    'P_KXH   C3  ҽ�������
    lngReturn = tsx_setstringparam(P_KXH, lngCounter, g������Ϣ_ͭɽ.ҽ�������)
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.ҽ�������, lngReturn)
    'P_TBR   C9  ���˱��
    lngReturn = tsx_setstringparam(P_TBR, lngCounter, g������Ϣ_ͭɽ.���˱��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.���˱��, lngReturn)
    'P_YYKSM C20 �Һſ�����
    lngReturn = tsx_setstringparam(P_YYKSM, 0, str������)
    Call WriteBusinessLOG("tsx_setstringparam", "P_YYKSM,0," & str������, lngReturn)
    'P_GHF   D   �Һŷ�
     lngReturn = tsx_getdoubleparam(P_GHF, 0, dbl�Һŷ�)
     Call WriteBusinessLOG("tsx_getdoubleparam", "P_GHF, 0, " & dbl�����Է�, lngReturn)
    'P_CZRYH C10 ������Ա���
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

    If tsx_jkcall("MZGH") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "MZGH", -1)
         STRERR = tsx_getlasterr()
         Err.Raise 9000, gstrSysName, STRERR
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams()
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
         Call WriteBusinessLOG("tsx_jkcall", "MZGH", lngReturn)
         '4 ȡ����ֵ
         lngReturn = tsx_getstringparam(P_DJH, 0, str������ˮ��)
         Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str��ˮ��, lngReturn)
    End If

    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("����tsx_destroyparams", "", lngReturn)
    ����Һ�_ͭɽ�� = True
    '��������¼
    strSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_ͭɽ�� & "," & g������Ϣ_ͭɽ.����ID & "," & _
        Year(zlDatabase.Currentdate) & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & 0 & ",0,0,0," & dbl�Һŷ� & ",0,0," & _
        dbl�Һŷ� & "," & 0 & ",0," & 0 & "," & 0 & ",'" & str������ˮ�� & "',NULL,NULL,'" & "" & "')"
    Call WriteBusinessLOG("����Һ�", "���汣�ս����¼", gstrSQL)
    Call zlDatabase.ExecuteProcedure(strSQL, "ͭɽ��ҽ��")
    
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����Һų���_ͭɽ��(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'******************************************************************************
'�����ߡ��������������÷���������ҺŲ�������
'����˵��������������ͨ������ҽ���̵�����Һų����ӿڣ��������ҺŽ��������
'���ù����嵥��˵����
'��������������
''*****************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng����ID As Long
    Dim str������ As String, dbl�Һŷ�  As Double, str������ˮ�� As String
    Dim str������ˮ�� As String, lngReturn As Long, lngCounter As Long, STRERR As String
    On Error GoTo ErrH
    strSQL = " select distinct A.����ID,A.NO,B.���˿���ID,to_char(B.����ʱ��,'YYYYMMDD') as �������� " & _
             " from ������ü�¼ A,������ü�¼ B " & _
             " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˽����¼", lng����ID)
    lng����ID = Val("" & rsTmp!����ID)
    
    strSQL = "Select B.��Ա���,A.����ID,B.����,B.�ʻ���־,A.֧��˳���,A.�������ý��,B.��ע " & _
              " from �����ʻ� B,���ս����¼ A " & _
              " Where A.����=[2] And B.���� = [2]" & _
              " And B.����ID = A.����ID And A.��¼ID=[1]"
    Call WriteBusinessLOG("����Һų���", "�ᱻ������¼", strSQL)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������¼", lng����ID, intinsure)
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "�޽����¼,���ܳ���!"
        Exit Function
    End If
    
    Do Until rsTmp.EOF
        str������ˮ�� = Trim("" & rsTmp!֧��˳���)
        rsTmp.MoveNext
    Loop
    
    If tsx_createparams(102400, 102400) = -1 Then
        Err.Raise 9000, gstrSysName, "�����ڴ�ռ�ʧ��!"
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    
    'P_JGM   C5  ҽԺ��
    lngReturn = tsx_setstringparam(P_JGM, lngCounter, g������Ϣ_ͭɽ.ҽԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
    'P_TBR   C9  ���˱��
    lngReturn = tsx_setstringparam(P_TBR, lngCounter, g������Ϣ_ͭɽ.���˱��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.���˱��, lngReturn)
    'P_KXH   C3  ҽ�������
    lngReturn = tsx_setstringparam(P_KXH, lngCounter, g������Ϣ_ͭɽ.ҽ�������)
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.ҽ�������, lngReturn)
    'P_DJH   C20 ԭ�Һŵ��ݺ�
    lngReturn = tsx_setstringparam(P_DJH, 0, str������ˮ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_DJH,0," & str������ˮ��, lngReturn)
    'P_CZRYH C10 ������Ա���
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

    If tsx_jkcall("MZTH") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "MZTH", -1)
         STRERR = tsx_getlasterr()
         Err.Raise 9000, gstrSysName, STRERR
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams()
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
         Call WriteBusinessLOG("tsx_jkcall", "MZTH", lngReturn)
         '4 ȡ����ֵ
         lngReturn = tsx_getstringparam(P_DJH, 0, str������ˮ��)
         Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str������ˮ��, lngReturn)
    End If

    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    
    strSQL = "Select * from ���ս����¼ Where ��¼ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ս���", lng����ID)
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_ͭɽ�� & "," & rsTmp!����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTmp!�������ý��, 0) & "," & -1 * Nvl(rsTmp!ȫ�Ը����, 0) & _
        "," & -1 * Nvl(rsTmp!�����Ը����, 0) & "," & -1 * Nvl(rsTmp!����ͳ����, 0) & _
        "," & -1 * Nvl(rsTmp!ͳ�ﱨ�����, 0) & _
        "," & -1 * Nvl(rsTmp!���Ը����, 0) & "," & 0 & "," & -1 * Nvl(rsTmp!�����ʻ�֧��) & _
        ",'" & str������ˮ�� & "',null,null,'" & Nvl(rsTmp!��ע) & "')"
    Call WriteBusinessLOG("����Һų���", "���ս����¼", gstrSQL)
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function �����������_ͭɽ��(rsHis As ADODB.Recordset, str���㷽ʽ As String, ByVal intinsure As Integer, Optional lng����ID As Long = 0) As Boolean
'*******************************************
'�����ߡ���������������clsInsure �� ClinicPreSwap ���̵���(�÷�����������ò�������)
'����˵��������������ͨ������ҽ���̵�Ԥ���㷽�����ֽⱾ�η�����ϸ���õ��������������ʻ�
'�����������������������١�ͳ�������ٵȣ�����������������ʽ�����ڲ�����str���㷽ʽ����
'����˵��������������
'                   1������ӿ���Ҫ������÷�����ϸ�ϴ��ӿڣ���������ϸ�ϴ�
'                   2����������Ԥ����ӿ�
'                   3�������������涨��ʽ����
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string�Ͳ���
'������tsx_setdoubleparam������double�Ͳ���
'������tsx_setlongparam������long�Ͳ���
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_getstringparam��ȡstring�ͽӿڷ���ֵ
'������tsx_getdoubleparam��ȡdouble�ͽӿڷ���ֵ
'������tsx_destroyparams�������ѷ���ռ�
'*******************************************
    'TODO:�����������
    'rs��ϸ��¼�����Ǳ���¼������ﴦ����ϸ
    'str���㷽ʽ�ĸ�ʽ˵����������ʽ;���;�Ƿ������޸�|....
    Dim dblͳ����� As Double, dbl��֧�� As Double
    Dim lngReturn As Long, lngCounter As Long
    Dim rsMzxnjs As New ADODB.Recordset, str��Ŀ��� As String
    Dim blnErr As Boolean '��¼�ϴ�ʱ�Ƿ��д���
    Dim str���ｻ����ˮ�� As String, dbl�ܷ��� As Double, dbl�����Է� As Double
    Dim dbl�����ʻ�֧�� As Double, dblͳ�����֧�� As Double, dbl��ͳ��֧�� As Double
    Dim dbl����Ա����֧�� As Double, dbl��ĩ�����ʻ� As Double, dbl�ڳ������ʻ� As Double
    Dim dbl�����Ը� As Double, dbl���� As Double, strҽ���� As String, str���ұ��� As String
    Dim STRERR As String
    Dim str���˱�� As String, strҽ������� As String, str�������� As String
    Dim rs��ϸ As New ADODB.Recordset
    On Error GoTo errHandle
    
    '>>Beging ����ǰҪ��ˢ��
'    If lng����ID > 0 Then
'        lngReturn= tsx_init_ic(g������Ϣ_ͭɽ.Com��)
'        Call WriteBusinessLOG("tsx_init_ic", g������Ϣ_ͭɽ.Com��, lngReturn)
'        lngReturn= tsx_read_ic(str���˱��, strҽ�������)
'        Call WriteBusinessLOG("tsx_read_ic", str���˱�� & "," & strҽ�������, lngReturn)
'
'        If str���˱�� <> g������Ϣ_ͭɽ.���˱�� Then
'            MsgBox "����ʱ��ҽ�����������֤ʱ��ҽ�����Ų��������ܽ��㣡", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    '>>End ����ǰҪ��ˢ��
    gstrSQL = "Select ������� From �����ʻ� Where ����=[1] ANd ҽ����=[2]"
    Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", intinsure, g������Ϣ_ͭɽ.���˱��)
    str�������� = rsMzxnjs!�������
    
    rsHis.Filter = "ʵ�ս�� <> 0"
    Set rs��ϸ = rsHis
    
    gstrSQL = "Select * from ���ű� where ID=[1]"
    Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "���ű�", CLng(rs��ϸ!��������ID))
    str���ұ��� = rsMzxnjs!����
    gstrSQL = "Select ��� from ��Ա�� where ����=[1]"
    Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "���ű�", CStr(rs��ϸ!������))
    strҽ���� = rsMzxnjs!���


    '>>Beging �ϴ���ϸ========================================================================================
    '1
    'If lng����ID = 0 Then
        If tsx_createparams(102400, 102400) = -1 Then
            MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
   
        lngCounter = 0
        
        Do Until rs��ϸ.EOF
        
            '2 Ϊ������ֵ
            lngReturn = tsx_setstringparam(P_JGM, lngCounter, g������Ϣ_ͭɽ.ҽԺ��)
            Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
            
            lngReturn = tsx_setstringparam(P_TBR, lngCounter, g������Ϣ_ͭɽ.���˱��) '    C9  ���˱��    �α���Ա���˱��
            Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.���˱��, lngReturn)
            
            lngReturn = tsx_setstringparam(P_KXH, lngCounter, g������Ϣ_ͭɽ.ҽ�������) '   C3  ҽ�������  �α���ԱIC�����
            Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & lngCounter & "," & g������Ϣ_ͭɽ.ҽ�������, lngReturn)
            
            gstrSQL = "Select * from �շ�ϸĿ where ID=[1]"
            Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�շ����", CLng(rs��ϸ!�շ�ϸĿID))
            
            str��Ŀ��� = 1
            If InStr("567", rsMzxnjs!���) > 0 Then
               str��Ŀ��� = 0
            End If
            
            lngReturn = tsx_setstringparam(P_LB, lngCounter, str��Ŀ���) '    C1  ���    0-ҩƷ,1-������Ŀ
            Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & lngCounter & "," & str��Ŀ���, lngReturn)
            
            lngReturn = tsx_setstringparam(P_JZLB, lngCounter, "0") ' C1  �������    �ݹ̶�Ϊ'0'
            Call WriteBusinessLOG("tsx_setstringparam", "P_JZLB" & ", " & lngCounter & ",'0'", lngReturn)
            
            lngReturn = tsx_setstringparam(P_YBBM, lngCounter, "") '  C20 ҽ������    ��Ϊ�մ�""
            Call WriteBusinessLOG("tsx_setstringparam", "P_YBBM" & ", " & lngCounter & ",''", lngReturn)
            
            gstrSQL = "Select * from ypzlk where �շ�ϸĿID=" & rs��ϸ!�շ�ϸĿID
            
            Call OpenRecordset_OtherBase(rsMzxnjs, "ypzlk", , gcnͭɽ��)
            If rsMzxnjs.EOF = False Then
               lngReturn = tsx_setstringparam(P_ZBM, lngCounter, rsMzxnjs!�Ա���) '   C20 �Ա���  ҩƷ(��������Ŀ)�Ա���
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCounter & "," & rsMzxnjs!�Ա���, lngReturn)
            Else
                lngReturn = tsx_setstringparam(P_ZBM, lngCounter, rs��ϸ!�շ�ϸĿID) '   C20 �Ա���  ҩƷ(��������Ŀ)�Ա���
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCounter & "," & rs��ϸ!�շ�ϸĿID, lngReturn)
            End If
            
            'Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCounter & "," & rs��ϸ!�շ�ϸĿID, lngReturn)
            dbl���� = Format(rs��ϸ!ʵ�ս�� / IIf(Format(rs��ϸ!����, "0") = 0, 1, Format(rs��ϸ!����, "0")), "0.0000")
            lngReturn = tsx_setdoubleparam(P_JG, lngCounter, dbl����)    '    D   ����
            Call WriteBusinessLOG("tsx_setdoubleparam", "P_JG" & ", " & lngCounter & "," & rs��ϸ!ʵ�ս�� / rs��ϸ!����, lngReturn)
            
            lngReturn = tsx_setlongparam(P_SL, lngCounter, IIf(Format(rs��ϸ!����, "0") = 0, 1, Format(rs��ϸ!����, "0")))  '    L   ����
            Call WriteBusinessLOG("tsx_setlongparam", "P_SL" & ", " & lngCounter & "," & rs��ϸ!����, lngReturn)
            
            
            lngCounter = lngCounter + 1
            rs��ϸ.MoveNext
        Loop
    'End If
            '3 ���ýӿ�
            If tsx_jkcall("MZSF_SC") = -1 Then
                 Call WriteBusinessLOG("tsx_jkcall", "MZSF_SC", -1)
                 STRERR = tsx_getlasterr()
                 MsgBox STRERR, vbInformation, gstrSysName
                 Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
                 lngReturn = tsx_destroyparams
                 Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
                 Exit Function
            Else
            Call WriteBusinessLOG("tsx_jkcall", "MZSF_SC", lngReturn)
            '4 ȡ����ֵ
                 '�޷���ֵ
            End If
            '5 �����ѷ���ռ�
            lngReturn = tsx_destroyparams()
            Call WriteBusinessLOG("��ϸ�ϴ�tsx_destroyparams", "", lngReturn)
    
    '>>End �ϴ���ϸ========================================================================================
'
    
    '>>Beging ���ü���========================================================================================
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    '2 Ϊ������ֵ
    lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, g������Ϣ_ͭɽ.���˱��) '    C9  ���˱��    �α���Ա���˱��
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & g������Ϣ_ͭɽ.���˱��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, g������Ϣ_ͭɽ.ҽ�������) '   C3  ҽ�������  �α���ԱIC�����
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽ�������, lngReturn)

    lngReturn = tsx_setstringparam(P_TSMZ, 0, Nvl(str��������, "0")) '  C1  ����������
    Call WriteBusinessLOG("tsx_setstringparam", "P_TSMZ,0,'0'", lngReturn)
    
    lngReturn = tsx_setstringparam(P_YYKSM, 0, str���ұ���) 'C20 ҽԺ���ұ���
    Call WriteBusinessLOG("tsx_setstringparam", "P_YYKSM,0," & str���ұ���, lngReturn)
    
    lngReturn = tsx_setstringparam(P_YSRYH, 0, strҽ����) 'C10 ҽ����Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_YSRYH,0," & strҽ����, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CFH, 0, "") ' C20 ������
    Call WriteBusinessLOG("tsx_setstringparam", "P_CFH,0,''", lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZM, 0, g������Ϣ_ͭɽ.���ֱ���) '    C10 ������
    Call WriteBusinessLOG("tsx_setstringparam", "P_BZM,0," & g������Ϣ_ͭɽ.���ֱ���, lngReturn)
    
    lngReturn = tsx_setstringparam(P_JZ, 0, g������Ϣ_ͭɽ.�����)  'C1  �Ƿ���
    Call WriteBusinessLOG("tsx_setstringparam", "P_JZ,0," & g������Ϣ_ͭɽ.�����, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)
               
    '3 ���ýӿ�
    If tsx_jkcall("MZSF_JS") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "MZSF_JS", -1)
         STRERR = tsx_getlasterr()
         MsgBox STRERR, vbInformation, gstrSysName
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams()
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
         Call WriteBusinessLOG("tsx_jkcall", "MZSF_JS", lngReturn)
         '4 ȡ����ֵ
         lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl�ܷ���) 'D �ܷ���
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl�ܷ���, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl�����Է�) '  D   �����Է�
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl�����Է�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl�����Ը�)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl�����Ը�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl�����ʻ�֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl�����ʻ�֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dblͳ�����֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dblͳ�����֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl��ͳ��֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl��ͳ��֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl����Ա����֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl����Ա����֧��, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl�ڳ������ʻ�)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl�ڳ������ʻ�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl��ĩ�����ʻ�)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl�ڳ������ʻ�, lngReturn)
         
    End If
    

    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("����tsx_destroyparams", "", lngReturn)

    '>>End ���ü���========================================================================================
    
    
    '>>Beging ����ȷ��~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If lng����ID > 0 Then ''����ID>0
        If tsx_createparams(1024, 1024) = -1 Then
            MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
        '2 Ϊ������ֵ
        lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
        Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
        
        lngReturn = tsx_setstringparam(P_TBR, 0, g������Ϣ_ͭɽ.���˱��) '    C9  ���˱��    �α���Ա���˱��
        Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & g������Ϣ_ͭɽ.���˱��, lngReturn)
        
        lngReturn = tsx_setstringparam(P_KXH, 0, g������Ϣ_ͭɽ.ҽ�������) '   C3  ҽ�������  �α���ԱIC�����
        Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽ�������, lngReturn)
    
        lngReturn = tsx_setdoubleparam(P_QCGRZH, 0, Val(Format(dbl�ڳ������ʻ�, "0.00")))
        Call WriteBusinessLOG("tsx_setdoubleparam", "P_QCGRZH, 0, " & dbl�ڳ������ʻ�, lngReturn)
    
        lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
        Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)
                   
        '3 ���ýӿ�
        If tsx_jkcall("MZSF_QR") = -1 Then
             Call WriteBusinessLOG("tsx_jkcall", "MZSF_QR", -1)
             STRERR = tsx_getlasterr()
             MsgBox STRERR, vbInformation, gstrSysName
             Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
             lngReturn = tsx_destroyparams
             Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
             Exit Function
        Else
             Call WriteBusinessLOG("tsx_jkcall", "MZSF_QR", lngReturn)
             '4 ȡ����ֵ
             str���ｻ����ˮ�� = Space(20)
             lngReturn = tsx_getstringparam(P_DJH, 0, str���ｻ����ˮ��)  '   C20 ���ݺ�
             Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str���ｻ����ˮ��, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl�ܷ���) 'D �ܷ���
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl�ܷ���, lngReturn)
            
             lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl�����Է�) '  D   �����Է�
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl�����Է�, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl�����Ը�)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl�����Ը�, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl�����ʻ�֧��)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl�����ʻ�֧��, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dblͳ�����֧��)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dblͳ�����֧��, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl��ͳ��֧��)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl��ͳ��֧��, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl����Ա����֧��)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl����Ա����֧��, lngReturn)
            
             lngReturn = tsx_getdoubleparam(P_QMGRZH, 0, dbl��ĩ�����ʻ�)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_QMGRZH, 0, " & dbl��ĩ�����ʻ�, lngReturn)
        End If
        
        '5 �����ѷ���ռ�
        lngReturn = tsx_destroyparams()
        Call WriteBusinessLOG("�������tsx_destroyparams", "", lngReturn)

        '**���汣�ս����¼**
        If InStr(str���ｻ����ˮ��, Chr(0)) > 0 Then
            str���ｻ����ˮ�� = Mid(str���ｻ����ˮ��, 1, InStr(str���ｻ����ˮ��, Chr(0)) - 1)
        End If
        gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & intinsure & "," & g������Ϣ_ͭɽ.����ID & "," & _
            Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
            0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
            dbl�ܷ��� & "," & dbl�����Է� & "," & dbl�����Ը� & "," & dbl�ܷ��� - dbl�����Է� - dbl�����Ը� & "," & _
            dblͳ�����֧�� + dbl����Ա����֧�� & "," & dbl��ͳ��֧�� & "," & _
            0 & "," & dbl�����ʻ�֧�� & ",'" & str���ｻ����ˮ�� & "',Null,Null,Null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
        
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & g������Ϣ_ͭɽ.����ID & "," & TYPE_ͭɽ�� & ",'�ʻ����','''" & dbl��ĩ�����ʻ� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ�")
    End If '����ID>0
    '>>End ����ȷ��~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '��������֯���㷽ʽ����ʾ������
    str���㷽ʽ = "ͳ�����;" & dblͳ�����֧�� & ";0"
    str���㷽ʽ = str���㷽ʽ & "|��֧��;" & dbl��ͳ��֧�� & ";0"
    str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & dbl����Ա����֧�� & ";0"
    str���㷽ʽ = str���㷽ʽ & "|�����ʻ�;" & dbl�����ʻ�֧�� & ";0"
    
    �����������_ͭɽ�� = True
    
    Set rsMzxnjs = Nothing
    Call WriteBusinessLOG("����ǰ̨", "", str���㷽ʽ)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_ͭɽ��(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, ByVal intinsure As Integer) As Boolean
'*****************************************************************
'�����ߡ���������������clsInsure �� ClinicSwap ���̵���(�÷�����������ò�������)
'����˵�������������������������ӿ�
'����˵��������������
'��������������������1�������Ҫ�ϴ���ϸ�����������ϸ�ϴ��ӿ�
'��������������������2�������������ӿ�
'��������������������3������ɹ����򱣴汣�ս����¼
'���ù����嵥��˵����
'�����������������_ͭɽ�ء� �����ϸ�ϴ�,������㣬���㹦��
'*****************************************************************��
'TODO:�������
    Dim str���㷽ʽ As String
    Dim rsMzjs As New ADODB.Recordset
On Error GoTo errHandle
    
    gstrSQL = "Select ID,NO,���,��¼����,�Ǽ�ʱ�� as ����ʱ��,����ID,�շ����,�վݷ�Ŀ,���㵥λ,������,��������ID, " & _
                     "�շ�ϸĿID,nvl(����,0)*nvl(����,0) as ����,��׼���� as ����, " & _
                     "ʵ�ս��,ͳ����,���մ���ID ����֧������ID, " & _
                     " ժҪ,�Ƿ��� " & _
            "from ������ü�¼ " & _
            "where ����ID=[1]"
    Set rsMzjs = zlDatabase.OpenSQLRecord(gstrSQL, "", lng����ID)
    
    '**���汣�ս����¼**
'    gstrSQL = "zl_���ս����¼_insert(" & IIf(blnסԺ, 2, 1) & "," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
'        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
'        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
'        dbl�����ܶ� & "," & dbl�ֽ� & ",0,0," & dblͳ����� & "," & dbl�󲡲��� & "," & _
'        0 & ",0,'" & gComInfo.������ˮ�� & "',null,null,'" & gComInfo.ҵ������ & "')"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
  
    �������_ͭɽ�� = �����������_ͭɽ��(rsMzjs, str���㷽ʽ, intinsure, lng����ID)
    Set rsMzjs = Nothing
    
    Call WriteBusinessLOG("�ӽ��㷵��ǰ̨", "", str���㷽ʽ)

    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_ͭɽ��(ByVal lng����ID As Long, ByVal cur�����ʻ� As Currency, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'*************************************************************************
'�����ߡ���������������clsInsure �� ClinicDelSwap ���̵���
'����˵����������������������������Ͻӿ�
'����˵��������������1�����ӿڹ����ж��Ƿ��������һ�ξ�������ﵥ�ݿ�ʼ�˷�
'��������������������2����������������Ͻӿ�
'��������������������3�����汣�ս����¼
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string�Ͳ���
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_getstringparam��ȡstring�ͽӿڷ���ֵ
'������tsx_destroyparams�������ѷ���ռ�
'*************************************************************************
'TODO:�������
    Dim rsMzJsCx As New ADODB.Recordset
    Dim lngReturn As Long, str���˱�� As String, strҽ������� As String
    Dim strԭ���ݺ� As String, lng����ID As Long, STRERR As String
On Error GoTo errHand
    
    gstrSQL = "Select * from �����ʻ� where ����ID=[1] And ����=[2]"
    Set rsMzJsCx = zlDatabase.OpenSQLRecord(gstrSQL, "", lng����ID, TYPE_ͭɽ��)
    If rsMzJsCx.EOF Then
        Err.Raise 9000, gstrSysName, "����ͭɽ��ҽ���α���Ա,���ܳ���!"
        Exit Function
    End If
    str���˱�� = rsMzJsCx!ҽ����
    strҽ������� = rsMzJsCx!����
    
    gstrSQL = "select distinct A.����ID,A.NO from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsMzJsCx = zlDatabase.OpenSQLRecord(gstrSQL, "", lng����ID)
    lng����ID = rsMzJsCx!����ID
    
    gstrSQL = "Select * From ���ս����¼ Where ����=1 And ��¼ID=[1] and ����=[2]"
    Set rsMzJsCx = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ˮ��", lng����ID, TYPE_ͭɽ��)
    If rsMzJsCx.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "û���ҵ�ԭʼ�����¼���޷�����������������"
        Exit Function
    End If
    strԭ���ݺ� = rsMzJsCx!֧��˳���
    
    
    '1 ����ռ�
    If tsx_createparams(1024, 1024) = -1 Then
        Err.Raise 9000, gstrSysName, "�����ڴ�ռ�ʧ��!"
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)

    '2 Ϊ������ֵ
    lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, str���˱��) '    C9  ���˱��    �α���Ա���˱��
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str���˱��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, strҽ�������) '   C3  ҽ�������  �α���ԱIC�����
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & strҽ�������, lngReturn)
    
    lngReturn = tsx_setstringparam(P_DJH, 0, strԭ���ݺ�) '
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & strԭ���ݺ�, lngReturn)
   
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

    '3 ���ýӿ�
    If tsx_jkcall("MZCZ") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "MZCZ", -1)
         STRERR = tsx_getlasterr()
         MsgBox STRERR
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
        Call WriteBusinessLOG("tsx_jkcall", "MZCZ", 0)
    '4 ȡ����ֵ
        lngReturn = tsx_getstringparam(P_DJH, 0, strԭ���ݺ�)
        Call WriteBusinessLOG("tsx_getstringparam", "P_RYH, 0, " & strԭ���ݺ�, lngReturn)
         
    End If
    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("����tsx_destroyparams", "", lngReturn)


    '���汣�ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_ͭɽ�� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsMzJsCx!�������ý��, 0) & "," & -1 * Nvl(rsMzJsCx!ȫ�Ը����, 0) & _
        "," & -1 * Nvl(rsMzJsCx!�����Ը����, 0) & "," & -1 * Nvl(rsMzJsCx!����ͳ����, 0) & _
        "," & -1 * Nvl(rsMzJsCx!ͳ�ﱨ�����, 0) & _
        "," & -1 * Nvl(rsMzJsCx!���Ը����, 0) & "," & 0 & "," & -1 * Nvl(rsMzJsCx!�����ʻ�֧��) & _
        ",'" & strԭ���ݺ� & "',null,null,'" & Nvl(rsMzJsCx!��ע) & "')"
        
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
    
    ����������_ͭɽ�� = True
    '
    Set rsMzJsCx = Nothing
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ��Ժ�Ǽ�_ͭɽ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
'**************************************************************************************************
'�����ߡ���������������clsInsure �� ComeInSwap ���̵���(�ɲ�����Ժ��������)
'����˵��������������������Ժ�Ǽǽӿ�
'����˵��������������1���Ӳ�����ҳ����ȡ��Ժ���ڣ�������Ժ�Ǽ�Ҳ�ǵ��øýӿڣ���˲���ȡ��ǰ������Ϊ��Ժ�����ϴ���
'��������������������2��������Ժ�Ǽǽӿ�
'��������������������3��ִ����Ժ�Ǽǹ���(zl_�����ʻ�_��Ժ)�����Ĳ��˵ĵ�ǰ״̬
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string����
'������tsx_setdoubleparam������double����
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_getstringparam��ȡstring�ͽӿڷ���ֵ
'������tsx_destroyparams�������ѷ���ռ�
'***************************************************************************************************
    'TODO:��Ժ�Ǽ�
    Dim lngReturn As Long, str��Ժ���� As String, str��Ժ���� As String
    Dim str��Ժ��λ�� As String, strסԺ�� As String, str��ϵ�绰 As String
    Dim dbl��ԺѺ�� As Double, strסԺҽ�� As String, str����ҽ�� As String
    Dim strסԺ��ˮ�� As String, str���ұ��� As String, STRERR As String
    Dim rsRydj As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "Select * from �����ʻ� where ����ID=[1] and ����=[2] and nvl(��ע,'0')<>'0'"
                  
    Set rsRydj = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", lng����ID, TYPE_ͭɽ��)
    If Not rsRydj.EOF Then
        MsgBox "�òα���Ա��Ƿ���ؽύ�׻�δ��ɣ����ܰ�����Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select A.*,B.���� as ���� from ������ҳ A ,���ű� B" & _
               " where A.����ID=[1] And A.��ҳID=[2] And A.����=[3] And A.��Ժ����ID=B.ID"
               
    Set rsRydj = zlDatabase.OpenSQLRecord(gstrSQL, "������ҳ", lng����ID, lng��ҳID, TYPE_ͭɽ��)
    If rsRydj.EOF Then
        MsgBox "����ҽ�����˲��ܰ���ҽ����Ժ!", vbInformation, gstrSysName
        Exit Function
    End If
    str��Ժ���� = Format(rsRydj!��Ժ����, "yyyyMMdd")
    
    str��Ժ���� = rsRydj!��Ժ����ID: str��Ժ��λ�� = Nvl(rsRydj!��Ժ����, "")
    strסԺ�� = lng����ID & "_" & lng��ҳID: str��ϵ�绰 = Nvl(rsRydj!��ϵ�˵绰, "")
    strסԺҽ�� = Nvl(rsRydj!סԺҽʦ, ""): str����ҽ�� = Nvl(rsRydj!����ҽʦ, "")
    dbl��ԺѺ�� = 0
    
    str���ұ��� = rsRydj!����
    
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    '2 Ϊ������ֵ
    lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, g������Ϣ_ͭɽ.���˱��) '    C9  ���˱��    �α���Ա���˱��
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & g������Ϣ_ͭɽ.���˱��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, g������Ϣ_ͭɽ.ҽ�������) '   C3  ҽ�������  �α���ԱIC�����
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽ�������, lngReturn)

    lngReturn = tsx_setstringparam(P_YYKSM, 0, str���ұ���) 'C20 ҽԺ���ұ���
    Call WriteBusinessLOG("tsx_setstringparam", "P_YYKSM,0," & str���ұ���, lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZM, 0, "") '    C10 ������ ��Ϊ�մ�( g������Ϣ_ͭɽ.���ֱ���)
    Call WriteBusinessLOG("tsx_setstringparam", "P_BZM,0," & "", lngReturn)
    
    lngReturn = tsx_setstringparam(P_RYYMD, 0, str��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYYMD,0," & str��Ժ����, lngReturn)

    lngReturn = tsx_setstringparam(P_RYBQ, 0, str��Ժ����)  '
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ,0," & str��Ժ����, lngReturn)

    lngReturn = tsx_setstringparam(P_RYCWH, 0, str��Ժ��λ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH,0," & str��Ժ��λ��, lngReturn)

    lngReturn = tsx_setstringparam(P_ZYH, 0, strסԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYH,0," & strסԺ��, lngReturn)

    lngReturn = tsx_setstringparam(P_LXDH, 0, str��ϵ�绰)
    Call WriteBusinessLOG("tsx_setstringparam", "P_LXDH,0," & str��ϵ�绰, lngReturn)

    lngReturn = tsx_setdoubleparam(P_YJHJ, 0, dbl��ԺѺ��)
    Call WriteBusinessLOG("tsx_setdoubleparam", "P_YJHJ,0," & dbl��ԺѺ��, lngReturn)

    lngReturn = tsx_setstringparam(P_YSRYH, 0, strסԺҽ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_YSRYH,0," & strסԺҽ��, lngReturn)

    lngReturn = tsx_setstringparam(P_KDYSRYH, 0, str����ҽ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_YSRYH,0," & str����ҽ��, lngReturn)

    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

    '3 ���ýӿ�
    If tsx_jkcall("ZYDJ") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "ZYDJ", -1)
         STRERR = tsx_getlasterr()
         MsgBox STRERR, vbInformation, gstrSysName
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
        Call WriteBusinessLOG("tsx_jkcall", "ZYDJ", lngReturn)
    '4 ȡ����ֵ
         strסԺ��ˮ�� = Space(20)
         lngReturn = tsx_getstringparam(P_ZYXH, 0, strסԺ��ˮ��)
         Call WriteBusinessLOG("tsx_getstringparam", "P_ZYXH, 0, " & strסԺ��ˮ��, lngReturn)
         
    End If
    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)


    '�ı䲡��״̬
    If lngReturn = 0 Then
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_ͭɽ�� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
        strסԺ��ˮ�� = Mid(strסԺ��ˮ��, 1, InStr(strסԺ��ˮ��, Chr(0)) - 1)
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_ͭɽ�� & ",'˳���','''" & strסԺ��ˮ�� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������ˮ��")
        
        ��Ժ�Ǽ�_ͭɽ�� = True
    End If
    Set rsRydj = Nothing
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_ͭɽ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
'****************************************
'�����ߡ���������������clsInsure �� ComeInDelSwap  ���̵���
'����˵�����������������ó�����Ժ�Ǽǻ��Ժ�Ǽǽӿ�
'���ù����嵥��˵����
'����˵��������������1�����ӿڹ�����м�飨һ�㷢�����û���н�����Ĳ��ˣ���������øýӿڣ�
'��������������������2�����ó�����Ժ�ӿ�
'��������������������3��ִ�г�Ժ�Ǽǹ���(zl_�����ʻ�_��Ժ)�����Ĳ��˵ĵ�ǰ״̬
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string����
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_destroyparams�������ѷ���ռ�
'****************************************
'TODO:��Ժ�Ǽǳ���
    On Error GoTo errHand
    Dim rsRydjCx As New ADODB.Recordset
    Dim str��ˮ�� As String, STRERR As String
    Dim lngReturn As Long
    Dim str���˱�� As String, strҽ������ As String, str��ע As String
    '����ʵĲ�������
    gstrSQL = "Select sum(nvl(���ʽ��,0)) as ���ʽ�� from סԺ���ü�¼ where ����ID=[1] And ��ҳID=[2]"
    Set rsRydjCx = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", lng����ID, lng��ҳID)
    
    If rsRydjCx!���ʽ�� <> 0 Then
        MsgBox "�ѽ���ʵ�ҽ����Ա���ܳ�����Ժ�Ǽ�!", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select * from �����ʻ� where ����=[1] And ����ID=[2]"
    Set rsRydjCx = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", intinsure, lng����ID)
    str��ˮ�� = rsRydjCx!˳���
    str���˱�� = rsRydjCx!ҽ����
    strҽ������ = rsRydjCx!����
    str��ע = InputBox("����д��ע", gstrSysName)
    lngReturn = -1
    '>>>beging ������Ժ
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    '2 Ϊ������ֵ
    lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, str���˱��) '    C9  ���˱��    �α���Ա���˱��
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str���˱��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, strҽ������) '   C3  ҽ�������  �α���ԱIC�����
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & strҽ������, lngReturn)

    lngReturn = tsx_setstringparam(P_ZYXH, 0, str��ˮ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str��ˮ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZ, 0, str��ע) ' C10 ������Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)
    '3 ���ýӿ�
    If tsx_jkcall("ZYDJZX") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "ZYDJZX", -1)
         STRERR = tsx_getlasterr()
         MsgBox STRERR, vbInformation, gstrSysName
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
    Call WriteBusinessLOG("tsx_jkcall", "ZYDJZX", lngReturn)
    '4 ȡ����ֵ
       '�޷���
    End If
    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

    
    '>>>End ������Ժ
    If lngReturn = 0 Then
        '�ı��������ϴ���¼Ϊδ�ϴ�
        gstrSQL = "Update סԺ���ü�¼ Set �Ƿ��ϴ�=0 where nvl(�Ƿ��ϴ�,0)=1 and ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
        gcnOracle.Execute gstrSQL
    '    '�ı䲡��״̬
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
        ��Ժ�Ǽǳ���_ͭɽ�� = True
    End If
    Set rsRydjCx = Nothing
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_ͭɽ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
'****************************************
'�����ߡ���������������clsInsure �� LeaveSwap ���̵���
'����˵�����������������ó�Ժ�Ǽǽӿ�
'����˵��������������1�����ӿڹ�����м��
'��������������������2�����ó�Ժ�ӿ�
'��������������������3��ִ�г�Ժ�Ǽǹ���(zl_�����ʻ�_��Ժ)�����Ĳ��˵ĵ�ǰ״̬
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string����
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_destroyparams�������ѷ���ռ�
'****************************************
    'TODO:��Ժ�Ǽ�(strת��סԺ��Ҫ���� ����ע��)
    Dim rsCydj As New ADODB.Recordset, str��Ժ���� As String, strת��ҽԺ As String
    Dim strTmp As String, strOut As String, lng��� As Long, strTmp1 As String
    Dim str��Ժ���� As String, str��Ժ��λ As String, strסԺҽ�� As String, str����ҽ�� As String
    Dim str˳��� As String, lngReturn As Long, lng����ID As Long, rsTemp As New ADODB.Recordset
    Dim STRERR As String
    On Error GoTo errHand

    '�ı䲡��״̬
    gstrSQL = "Select * from ������ҳ  where ����ID=[1] And ��ҳID=[2]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "������ҳ", lng����ID, lng��ҳID)
    If rsCydj.EOF Then
        MsgBox "δ�ҵ���Ժ��¼!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str��Ժ���� = rsCydj!��Ժ����ID
    str��Ժ��λ = Nvl(rsCydj!��Ժ����, 0)
    strסԺҽ�� = Nvl(rsCydj!סԺҽʦ, "")
    str����ҽ�� = Nvl(rsCydj!����ҽʦ, "")
    
   
    Select Case rsCydj!��Ժ��ʽ
        Case "��ת"
            str��Ժ���� = 1
        Case "��ת"
            str��Ժ���� = 2
        Case Else
            str��Ժ���� = 0
    End Select
    
    gstrSQL = "Select * from ���ű� where ID=[1]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ����", CLng(str��Ժ����))
    If rsCydj.EOF Then
        MsgBox "��Ժ���Ҳ���!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str��Ժ���� = rsCydj!����

    gstrSQL = "Select ��� from ��Ա�� A,��Ա����˵�� B where A.ID=B.��ԱID and B.��Ա����='ҽ��' and  A.����=[1]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��", strסԺҽ��)
    If rsCydj.EOF Then
        MsgBox "סԺҽ������!", vbInformation, gstrSysName
        Exit Function
    End If
    
    strסԺҽ�� = rsCydj!���
    
    gstrSQL = "Select ��� from ��Ա�� A,��Ա����˵�� B where A.ID=B.��ԱID and B.��Ա����='ҽ��' and  A.����=[1]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��", str����ҽ��)
    If rsCydj.EOF Then
        MsgBox "����ҽ������!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str����ҽ�� = rsCydj!���
    
    If str��Ժ���� <> 0 Then
    
        gstrSQL = "Select ҽԺ����,ҽԺ���� from YYDA"
        Call OpenRecordset_OtherBase(rsCydj, "ҽԺ�б�", , gcnͭɽ��)
        strTmp = ""
        strTmp1 = ""
        Do Until rsCydj.EOF
            strTmp = strTmp & rsCydj!ҽԺ���� & ";"
            strTmp1 = strTmp1 & rsCydj!ҽԺ���� & ";"
            rsCydj.MoveNext
        Loop
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        strTmp1 = Mid(strTmp1, 1, Len(strTmp1) - 1)
        strOut = frmShowList.ShowME(strTmp & "||" & strTmp1, "ת��ҽԺ����||ת��ҽԺ����")
        
        strת��ҽԺ = Split(strOut, ";")(0)
    End If
    gstrSQL = "select * from �����ʻ� where ����ID=[1]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", lng����ID)
    If rsCydj.EOF Then
        MsgBox "���Ǳ�ҽ������!", vbInformation, gstrSysName
        Exit Function
    End If
   
    lng����ID = rsCydj!����ID
    str˳��� = rsCydj!˳���
    
    'Beging ��鲡��
    gstrSQL = "Select * from icd10 where ID=" & lng����ID
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "����ID", lng����ID)
    If rsCydj.EOF Then
            'ǿ��ѡ����
        gstrSQL = " Select A.ID,A.���ֱ���,A.��������,A.ƴ����" & _
                " From Icd10 A "
        Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ȷ�Ｒ��")
        If rsTemp.State = 1 Then
            gstrSQL = " ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_ͭɽ�� & ",'����ID','''" & rsTemp!ID & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ID")
        Else
            ��Ժ�Ǽ�_ͭɽ�� = False
            Exit Function
        End If
    End If
    
    'End  ��鲡��


    '1 ����ռ�
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    '2 Ϊ������ֵ
    lngReturn = tsx_setstringparam(P_ZYXH, 0, str˳���)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str˳���, lngReturn)
    
    lngReturn = tsx_setstringparam(P_RYBQ, 0, str��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & 0 & "," & str��Ժ����, lngReturn)
    
    lngReturn = tsx_setstringparam(P_RYCWH, 0, str��Ժ��λ) '
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & 0 & "," & str��Ժ��λ, lngReturn)
    
    lngReturn = tsx_setstringparam(P_YSRYH, 0, strסԺҽ��) '
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & 0 & "," & strסԺҽ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KDYSRYH, 0, str����ҽ��) '
    Call WriteBusinessLOG("tsx_setstringparam", "P_KDYSRYH" & ", " & 0 & "," & str����ҽ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

    '3 ���ýӿ�
    If tsx_jkcall("ZYBQCW") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "ZYBQCW", -1)
         STRERR = tsx_getlasterr()
         MsgBox STRERR, vbInformation, gstrSysName
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
    Call WriteBusinessLOG("tsx_jkcall", "ZYBQCW", lngReturn)
    '4 ȡ����ֵ
         
    End If
    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_ͭɽ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_ͭɽ�� & ",'��ע','''" & str��Ժ���� & strת��ҽԺ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����תԺ��Ϣ")
    
    ��Ժ�Ǽ�_ͭɽ�� = True
    Set rsCydj = Nothing
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_ͭɽ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
'*******************************************
'�����ߡ���������������clsInsure �� LeaveDelSwap ���̵���
'����˵�����������������ó�����Ժ�Ǽǻ���Ժ�Ǽǽӿ�
'����˵��������������1�����ӿڹ�����м��
'��������������������2�����ó�����Ժ�Ǽǻ���Ժ�Ǽǽӿ�
'��������������������3��ִ����Ժ�Ǽǹ���(zl_�����ʻ�_��Ժ)�����Ĳ��˵ĵ�ǰ״̬
'���ù����嵥��˵����
'�������ޡ�
'*******************************************
    'TODO:��Ժ����
    On Error GoTo errHand

    '�ı䲡��״̬

    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_ͭɽ�� & ",'��ע','0')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����תԺ��Ϣ")

    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_ͭɽ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")

    ��Ժ�Ǽǳ���_ͭɽ�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_ͭɽ��(ByVal lng����ID As Long, ByVal intinsure As Integer) As Currency
'*******************************************
'�����ߡ���������������clsInsure �� SelfBalance ���̵���
'����˵�����������������ø����ʻ�����ѯ�ӿڻ�ֱ�Ӵӱ����ʻ�������ȡ�����ʻ����
'����˵��������������1�����ò�ѯ�ӿڻ�ȡ�����ʻ������±����ʻ���
'��������������������2������ֱ�Ӵӱ����ʻ�����ȡ�����ʻ����
'���ù����嵥��˵����
'�������ޡ�
'******************************************
    'TODO:�������
    '������� = 0
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = "Select Nvl(�ʻ����,0) AS �����ʻ� From �����ʻ� " & _
              " Where ����ID=[1] and ����=[2]"
              
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID, TYPE_ͭɽ��)
    �������_ͭɽ�� = rsTemp!�����ʻ�
    Set rsTemp = Nothing
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If

    
End Function

Public Function סԺ����_ͭɽ��(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal intinsure As Integer, Optional ByRef strAdvance As String = "") As Boolean
'********************************
'�����ߡ���������������clsInsure �� SettleSwap ���̵���
'����˵����������������ɱ���סԺ���õ�ҽ������
'����˵��������������1������סԺ����ӿ�
'��������������������2�����סԺ���㷵�صĽ�������סԺԤ���㷵�صĲ�һ�£���Ҫ����
'��������������������zl_���˽����¼_Update���̽�������
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string�Ͳ���
'������tsx_setdoubleparam������double�Ͳ���
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_getstringparam��ȡstring�ͽӿڷ���ֵ
'������tsx_getdoubleparam��ȡdouble�ͽӿڷ���ֵ
'������tsx_destroyparams�������ѷ���ռ�
'*********************************
    
    'TODO:סԺ����
    Dim rsZyjs As New ADODB.Recordset, lngReturn As Long, lngǷ���ؽ�ID As Long, str���㷽ʽ As String
    Dim str��ˮ�� As String, strҽ������ As String, str���˱��� As String
    Dim str�����־ As String, str��Ժ���� As String, lng��ҳID As Long, str���ֱ��� As String, strתԺ���� As String
    Dim strת��ҽԺ�� As String
    Dim str��Ժ���� As String, strסԺ������ˮ�� As String
    Dim dbl�ܷ��� As Double, dbl�����Է� As Double
    Dim dbl�����ʻ�֧�� As Double, dblͳ�����֧�� As Double, dbl��ͳ��֧�� As Double
    Dim dbl����Ա����֧�� As Double, dbl��ĩ�����ʻ� As Double, dbl�ڳ������ʻ� As Double
    Dim dbl�����Ը� As Double, STRERR As String
    On Error GoTo errHandle
    
    '>Beging ׼������
    str�����־ = 1

    gstrSQL = "Select * from �����ʻ� where ����=[1] And ����ID=[2]"
    Call WriteBusinessLOG("", gstrSQL, "�������ʻ�")
    
    Set rsZyjs = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", intinsure, lng����ID)
    lngǷ���ؽ�ID = Nvl(rsZyjs!Ƿ���ؽ�ID, 0)
    str��ˮ�� = rsZyjs!˳���
    strҽ������ = rsZyjs!����
    str���˱��� = rsZyjs!ҽ����
    str���ֱ��� = rsZyjs!����ID
    
    If Nvl(rsZyjs!��ע, "0") = 0 Then
        strתԺ���� = "0"
        strת��ҽԺ�� = ""
    Else
        strתԺ���� = Mid(rsZyjs!��ע, 1, 1)
        If strתԺ���� = 0 Then
            strת��ҽԺ�� = ""
        Else
            strת��ҽԺ�� = Mid(rsZyjs!��ע, 2, 1)
        End If
    End If
    
    gstrSQL = "Select * from icd10 where ID=" & str���ֱ���
    Call OpenRecordset_OtherBase(rsZyjs, "ICD10", , gcnͭɽ��)
    If rsZyjs.RecordCount > 0 Then
        str���ֱ��� = rsZyjs!���ֱ���
    End If
    
    gstrSQL = "Select * from ������Ϣ where ����ID=[1] And ����=[2]"
    Set rsZyjs = zlDatabase.OpenSQLRecord(gstrSQL, "������ҳ", lng����ID, intinsure)
    If Nvl(rsZyjs!��Ժʱ��, 0) = 0 Then
        Err.Raise 9000, gstrSysName, "��֧����;���㣬���Ȱ����Ժ���ٽ��н��㣡"
        Exit Function
    End If
    str��Ժ���� = Format(Nvl(rsZyjs!��Ժʱ��, Now()), "yyyyMMdd")
    lng��ҳID = Nvl(rsZyjs!סԺ����, 0)
    
    gstrSQL = "Select *  From ������ Where ����ID=[1]" & _
              " And ��ҳID=[2] And �������=3 And ��ϴ���=1"
              
    Set rsZyjs = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ���", lng����ID, lng��ҳID)
    If rsZyjs.EOF Then
        str��Ժ���� = 2
    Else
    
        Select Case Nvl(rsZyjs!��Ժ���, "��ת")
            Case "����"
                str��Ժ���� = 1
            Case "��ת"
                str��Ժ���� = 2
            Case "δ��"
                str��Ժ���� = 3
            Case "����"
                str��Ժ���� = 4
            Case Else
                str��Ժ���� = 5
        End Select
    End If
   
    '>End ׼������
    
    '>>Beging Ƿ���ؽ�
    If lngǷ���ؽ�ID > 0 Then
        gstrSQL = "Select * from ���ս����¼ where ��¼ID=[1]"
        Set rsZyjs = zlDatabase.OpenSQLRecord(gstrSQL, "���ս����¼", lngǷ���ؽ�ID)
        If rsZyjs.EOF Then
            MsgBox "δ�ҵ�ԭʼ���ʼ�¼������ִ�к���������"
            Exit Function
        End If
        str��ˮ�� = rsZyjs!֧��˳���
        '>>Beging ��QKCJS����
        '1 ����ռ�
        If tsx_createparams(1024, 1024) = -1 Then
            Err.Raise 9000, gstrSysName, "�����ڴ�ռ�ʧ��!"
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
        
        '2 Ϊ������ֵ
        lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
        Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
        
        lngReturn = tsx_setstringparam(P_TBR, 0, str���˱���) '    C9  ���˱��    �α���Ա���˱��
        Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str���˱���, lngReturn)
            '20051103 �ĵ���Ҫ��P_ZYXH ,ʵ��Ӧ��P_DJH
        lngReturn = tsx_setstringparam(P_DJH, 0, str��ˮ��)
        Call WriteBusinessLOG("tsx_setstringparam", "P_DJH" & ", " & 0 & "," & str��ˮ��, lngReturn)

        lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
        Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

        '3 ���ýӿ�
        If tsx_jkcall("QKCJS") = -1 Then
             Call WriteBusinessLOG("tsx_jkcall", "QKCJS", -1)
             STRERR = tsx_getlasterr()
             Err.Raise 9000, gstrSysName, STRERR
             Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
             lngReturn = tsx_destroyparams
             Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
             Exit Function
        Else
          Call WriteBusinessLOG("tsx_jkcall", "QKCJS", lngReturn)
        '4 ȡ����ֵ
          strסԺ������ˮ�� = Space(32)
          lngReturn = tsx_getstringparam(P_DJH, 0, strסԺ������ˮ��)   '   C20 ���ݺ�
          Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & strסԺ������ˮ��, lngReturn)
          strסԺ������ˮ�� = Mid(strסԺ������ˮ��, 1, InStr(strסԺ������ˮ��, Chr(0)) - 1)
           
          lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl�����Է�) '  D   �����Է�
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl�����Է�, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl�����Ը�)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl�����Ը�, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl�����ʻ�֧��)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl�����ʻ�֧��, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dblͳ�����֧��)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dblͳ�����֧��, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl��ͳ��֧��)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl��ͳ��֧��, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl����Ա����֧��)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl����Ա����֧��, lngReturn)
         
          lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl�ڳ������ʻ�)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl�ڳ������ʻ�, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_QMGRZH, 0, dbl��ĩ�����ʻ�)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_QMGRZH, 0, " & dbl��ĩ�����ʻ�, lngReturn)
             
        End If
        '5 �����ѷ���ռ�
        lngReturn = tsx_destroyparams()
        Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
        '>>end ��QKCJS����
        dbl�ܷ��� = Abs(dbl�����Է�) + Abs(dbl�����Ը�) + dbl�����ʻ�֧�� + dblͳ�����֧�� + dbl��ͳ��֧�� _
                   + dbl��ͳ��֧�� + dbl����Ա����֧��
                   
         If dbl�����ʻ�֧�� <> 0 Then
             str���㷽ʽ = "||�����ʻ�|" & dbl�����ʻ�֧��
         End If
         
         '2
         If dblͳ�����֧�� <> 0 Then
             str���㷽ʽ = str���㷽ʽ & "||ͳ�����|" & dblͳ�����֧��
         End If
        
         '3
         If dbl����Ա����֧�� <> 0 Then
             str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & dbl����Ա����֧��
         End If
         '4
         If dbl��ͳ��֧�� <> 0 Then
             str���㷽ʽ = str���㷽ʽ & "||��֧��|" & dbl��ͳ��֧��
         End If
         
         '�������
         If str���㷽ʽ <> "" Then
             str���㷽ʽ = Mid(str���㷽ʽ, 3)
             #If gverControl < 2 Then
                gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
             #Else
                strAdvance = str���㷽ʽ
                gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
             #End If
                
             Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
         End If
                   
        '**���汣�ս����¼**
        #If gverControl < 2 Then
            gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
                Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
                0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
                dbl�ܷ��� & "," & Abs(dbl�����Է�) & "," & Abs(dbl�����Ը�) & "," & dbl�ܷ��� - dbl�����Է� - dbl�����Ը� & "," & _
                dblͳ�����֧�� + dbl����Ա����֧�� & "," & dbl��ͳ��֧�� & "," & _
                0 & "," & dbl�����ʻ�֧�� & ",'" & strסԺ������ˮ�� & "',null,null,Null)"
        #Else
            gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
                Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
                0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
                dbl�ܷ��� & "," & Abs(dbl�����Է�) & "," & Abs(dbl�����Ը�) & "," & dbl�ܷ��� - dbl�����Է� - dbl�����Ը� & "," & _
                dblͳ�����֧�� + dbl����Ա����֧�� & "," & dbl��ͳ��֧�� & "," & _
                0 & "," & dbl�����ʻ�֧�� & ",'" & strסԺ������ˮ�� & "',null,null,Null,1)"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
        
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'Ƿ���ؽ�ID','" & 0 & " ')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ȡ��Ƿ���ؽ�")
        
        סԺ����_ͭɽ�� = True
        Exit Function
    End If
    '>>End Ƿ���ؽ�
    
    '>>Beging ��cyjs ����
    '1 ����ռ�
    If tsx_createparams(1024, 1024) = -1 Then
        Err.Raise 9000, gstrSysName, "�����ڴ�ռ�ʧ��!"
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    
    '2 Ϊ������ֵ
    lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, str���˱���) '    C9  ���˱��    �α���Ա���˱��
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str���˱���, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, strҽ������) '   C3  ҽ�������  �α���ԱIC�����
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & strҽ������, lngReturn)
    
    lngReturn = tsx_setstringparam(P_ZYXH, 0, str��ˮ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str��ˮ��, lngReturn)

    lngReturn = tsx_setstringparam(P_LB, 0, str�����־) 'C1  ��ע    Ԥ��������(0-Ԥ����,1-��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & 0 & "," & str�����־, lngReturn)

    lngReturn = tsx_setstringparam(P_CYYMD, 0, str��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYYMD" & ", " & 0 & "," & str��Ժ����, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CYXZ, 0, str��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYXZ" & ", " & 0 & "," & str��Ժ����, lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZM, 0, str���ֱ���)
    Call WriteBusinessLOG("tsx_setstringparam", "P_BZM" & ", " & 0 & "," & str���ֱ���, lngReturn)
    
    lngReturn = tsx_setstringparam(P_ZYTZ, 0, strתԺ����) 'C1 0 - ����תԺ, 1 - ��ת, 2 - ��ת
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYTZ" & ", " & 0 & "," & strתԺ����, lngReturn)

    lngReturn = tsx_setstringparam(P_ZWYYM, 0, strת��ҽԺ��) 'C5  ת��ҽԺ��  ת��ҽԺ��ҽԺ����
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZWYYM" & ", " & 0 & "," & strת��ҽԺ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

    '3 ���ýӿ�
    If tsx_jkcall("CYJS") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "CYJS", -1)
         STRERR = tsx_getlasterr()
         Err.Raise 9000, gstrSysName, STRERR
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
        Call WriteBusinessLOG("tsx_jkcall", "CYJS", lngReturn)
    '4 ȡ����ֵ
         lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl�ܷ���) 'D �ܷ���
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl�ܷ���, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl�����Է�) '  D   �����Է�
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl�����Է�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl�����Ը�)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl�����Ը�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl�����ʻ�֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl�����ʻ�֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dblͳ�����֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dblͳ�����֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl��ͳ��֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl��ͳ��֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl����Ա����֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl����Ա����֧��, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl�ڳ������ʻ�)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl�ڳ������ʻ�, lngReturn)
         
    End If
    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
    '>>End ��cyjs ����
    
    '>>> Beging ����CYQR����
    
        '1 ����ռ�
      If tsx_createparams(1024, 1024) = -1 Then
          Err.Raise 9000, gstrSysName, "�����ڴ�ռ�ʧ��!"
          Exit Function
      End If
      Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
      
      '2 Ϊ������ֵ
      lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
      Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
      
      lngReturn = tsx_setstringparam(P_TBR, 0, str���˱���) '    C9  ���˱��    �α���Ա���˱��
      Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str���˱���, lngReturn)
      
      lngReturn = tsx_setstringparam(P_KXH, 0, strҽ������) '   C3  ҽ�������  �α���ԱIC�����
      Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & strҽ������, lngReturn)
      
      lngReturn = tsx_setstringparam(P_ZYXH, 0, str��ˮ��)
      Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str��ˮ��, lngReturn)
  
      lngReturn = tsx_setstringparam(P_CYYMD, 0, str��Ժ����)
      Call WriteBusinessLOG("tsx_setstringparam", "P_CYYMD" & ", " & 0 & "," & str��Ժ����, lngReturn)
      
      lngReturn = tsx_setstringparam(P_CYXZ, 0, str��Ժ����)
      Call WriteBusinessLOG("tsx_setstringparam", "P_CYXZ" & ", " & 0 & "," & str��Ժ����, lngReturn)
      
      lngReturn = tsx_setstringparam(P_BZM, 0, str���ֱ���)
      Call WriteBusinessLOG("tsx_setstringparam", "P_BZM" & ", " & 0 & "," & str���ֱ���, lngReturn)
      
      lngReturn = tsx_setstringparam(P_ZYTZ, 0, strתԺ����) 'C1 0 - ����תԺ, 1 - ��ת, 2 - ��ת
      Call WriteBusinessLOG("tsx_setstringparam", "P_ZYTZ" & ", " & 0 & "," & strתԺ����, lngReturn)
  
      lngReturn = tsx_setstringparam(P_ZWYYM, 0, strת��ҽԺ��) 'C5  ת��ҽԺ��  ת��ҽԺ��ҽԺ����
      Call WriteBusinessLOG("tsx_setstringparam", "P_ZWYYM" & ", " & 0 & "," & strת��ҽԺ��, lngReturn)
      
      lngReturn = tsx_setdoubleparam(P_QCGRZH, 0, Val(Format(dbl�ڳ������ʻ�, "0.00")))
      Call WriteBusinessLOG("tsx_setdoubleparam", "P_QCGRZH, 0, " & dbl�ڳ������ʻ�, lngReturn)
      
      lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
      Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

      lngReturn = tsx_setstringparam(P_CXLB, 0, mstrʹ�ø����ʻ�֧��) 'C1  �����ʻ�֧��    �Ƿ�ʹ�ø����ʻ�֧��(0-��,1-��
      Call WriteBusinessLOG("tsx_setstringparam", "P_CXLB" & ", " & 0 & "," & mstrʹ�ø����ʻ�֧��, lngReturn)

      '3 ���ýӿ�
      If tsx_jkcall("CYQR") = -1 Then
           Call WriteBusinessLOG("tsx_jkcall", "CYQR", -1)
           STRERR = tsx_getlasterr()
           Err.Raise 9000, gstrSysName, STRERR
           Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
           lngReturn = tsx_destroyparams
           Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
           Exit Function
      Else
          Call WriteBusinessLOG("tsx_jkcall", "CYQR", lngReturn)
      '4 ȡ����ֵ
           strסԺ������ˮ�� = Space(32)
           lngReturn = tsx_getstringparam(P_DJH, 0, strסԺ������ˮ��)  '   C20 ���ݺ�
           Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & strסԺ������ˮ��, lngReturn)
           strסԺ������ˮ�� = Mid(strסԺ������ˮ��, 1, InStr(strסԺ������ˮ��, Chr(0)) - 1)
           
           lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl�ܷ���) 'D �ܷ���
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl�ܷ���, lngReturn)
          
           lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl�����Է�) '  D   �����Է�
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl�����Է�, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl�����Ը�)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl�����Ը�, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl�����ʻ�֧��)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl�����ʻ�֧��, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dblͳ�����֧��)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dblͳ�����֧��, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl��ͳ��֧��)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl��ͳ��֧��, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl����Ա����֧��)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl����Ա����֧��, lngReturn)
          
           lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl�ڳ������ʻ�)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl�ڳ������ʻ�, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_QMGRZH, 0, dbl��ĩ�����ʻ�)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_QMGRZH, 0, " & dbl��ĩ�����ʻ�, lngReturn)
         
      End If
      '5 �����ѷ���ռ�
      lngReturn = tsx_destroyparams()
      Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
      
      '**���汣�ս����¼**
      gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
          Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
          0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
          dbl�ܷ��� & "," & dbl�����Է� & "," & dbl�����Ը� & "," & dbl�ܷ��� - dbl�����Է� - dbl�����Ը� & "," & _
          dblͳ�����֧�� + dbl����Ա����֧�� & "," & dbl��ͳ��֧�� & "," & _
          0 & "," & dbl�����ʻ�֧�� & ",'" & strסԺ������ˮ�� & "',null,null,Null)"
      Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")

    '>>> End ����CYQR����
    Set rsZyjs = Nothing
    סԺ����_ͭɽ�� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_ͭɽ��(rsExse As Recordset, lng����ID As Long, strҽ���� As String, str���� As String, ByVal intinsure As Integer) As String
'*********************************
'�����ߡ���������������clsInsure �� WipeoffMoney ���̵���
'����˵����������������ɱ���סԺ���õ�ҽ��Ԥ����
'����˵��������������1����Ҫ��δ�ϴ��Ĵ�����ϸ�ϴ������ģ����ƽ������ʱ��ʵʱ�ϴ���
'���������������������򱾴�ʵ�ʽ��ϴ����Զ����������ϸ��
'��������������������2�����ݽӿ����ʣ�ÿ�������ϴ������ϴ��������ɹ��ϴ�����ϸ�����ϴ����
'��������������������3������סԺԤ����ӿ�
'��������������������4�����涨��ʽ���ؽ�����������μ�����Ԥ����
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string�Ͳ���
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_getdoubleparam��ȡdouble�ͽӿڷ���ֵ
'������tsx_destroyparams�������ѷ���ռ�
'*********************************
    
    'TODO:�������

    Dim rsZyxnjs As New ADODB.Recordset, lngReturn As Long, str��Ϣ As String, str���㷽ʽ As String
    Dim str��ˮ�� As String, strҽ������ As String, str���˱��� As String, str���ֱ��� As String
    Dim strתԺ���� As String, strת��ҽԺ��      As String
    Dim str�����־ As String, str��Ժ���� As String, lng��ҳID As Long
    Dim str��Ժ���� As String, strסԺ������ˮ�� As String
    Dim dbl�ܷ��� As Double, dbl�����Է� As Double, dbl�����Ը� As Double, dblͳ���Ը� As Double, dbl���Ը� As Double
    Dim dbl�����ʻ�֧�� As Double, dblͳ�����֧�� As Double, dbl��ͳ��֧�� As Double, dbl�ⶥ�Ը� As Double
    Dim dbl����Ա����֧�� As Double, dbl��ĩ�����ʻ� As Double, dbl�ڳ������ʻ� As Double
    Dim dbl�����Ը� As Double, rsCFMX As New ADODB.Recordset, STRERR As String
    On Error GoTo errHandle
    
    '>beging ����δ�ϴ���¼
       ' ���δ�ϴ���¼��,��¼���� , ��¼״̬, NO ���ü����ϴ�
    gstrSQL = "Select distinct ��¼����,��¼״̬,NO From סԺ���ü�¼ A,�����ʻ� B,������Ϣ C " & _
              " Where A.����ID=B.����ID And A.����ID=C.����ID And A.��ҳID=C.סԺ����" & _
              " And nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 and A.���ʷ���=1 And A.����Ա���� is not null " & _
              " AND A.ʵ�ս�� IS NOT NULL And B.����ID=[1] And B.����=[2]"
     Call WriteBusinessLOG("ȡ���˷��ü�¼", gstrSQL, "")
     
     Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˷��ü�¼", lng����ID, intinsure)

     Do Until rsZyxnjs.EOF
         Call WriteBusinessLOG("���ô����ϴ�_ͭɽ��", "2" & "," & rsZyxnjs!NO & "," & rsZyxnjs!��¼���� & "," & rsZyxnjs!��¼״̬ & "," & str��Ϣ & "," & lng����ID & "," & intinsure, lngReturn)
         Call �����ϴ�_ͭɽ��(2, rsZyxnjs!NO, rsZyxnjs!��¼����, rsZyxnjs!��¼״̬, str��Ϣ, lng����ID, intinsure)
         Call WriteBusinessLOG("��ɵ��ô����ϴ�_ͭɽ��", "2" & "," & rsZyxnjs!NO & "," & rsZyxnjs!��¼���� & "," & rsZyxnjs!��¼״̬ & "," & str��Ϣ & "," & lng����ID & "," & intinsure, lngReturn)

         rsZyxnjs.MoveNext
         
'         If rsZyxnjs!��¼״̬ > 1 Then
'         gstrSQL = "Select * from ���˷��ü�¼ where NO='" & rsZyxnjs!NO & " And ��¼����=" & rsZyxnjs!��¼���� & _
'                    " and  ��¼״̬=" & rsZyxnjs!��¼״̬ & " And ����ID=" & lng����ID
'
'         Call OpenRecordset(rsCFMX, "������ϸ")
'         Do Until rsCFMX.EOF
'            rsCFMX.MoveNext
'         Loop
'         End If
     Loop
    '>end ����δ�ϴ���¼
    
    '>>Beging Ƿ���ؽ�
    gstrSQL = "Select * from �����ʻ� where ����=[1] And ����ID=[2] And nvl(Ƿ���ؽ�ID,0)>0"
    Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", intinsure, lng����ID)
    If Not rsZyxnjs.EOF Then
        MsgBox "��ʾ��Ƿ���ؽ�ʱ��Ԥ����Ϊȫ�Էѣ�", vbInformation, gstrSysName
        str���㷽ʽ = "ͳ�����;" & 0 & ";0"
        str���㷽ʽ = str���㷽ʽ & "|��֧��;" & 0 & ";0"
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & 0 & ";0"
        str���㷽ʽ = str���㷽ʽ & "|�����ʻ�;" & 0 & ";0"
        
        סԺ�������_ͭɽ�� = str���㷽ʽ
        Call WriteBusinessLOG("�˳�Ƿ���ؽ�Ԥ����!", "", "")
        Exit Function
    End If
    '>>End Ƿ���ؽ�
    
    
    '>Beging ׼������
    str�����־ = 0
    mstrʹ�ø����ʻ�֧�� = 0
    '������òŵ���
    
    '2008-12-12 ����ǿҪ��ȥ��
'    If str���� = 1 Then
'        If MsgBox("�Ƿ�ʹ�ø����ʻ�֧����", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbYes Then
'            mstrʹ�ø����ʻ�֧�� = 1
'        End If
'    End If

    gstrSQL = "Select * from �����ʻ� where ����=[1] And ����ID=[2]"
    Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", intinsure, lng����ID)
    
    If rsZyxnjs.EOF Then
        MsgBox "���Ǳ�ҽ������!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str��ˮ�� = rsZyxnjs!˳���
    strҽ������ = rsZyxnjs!����
    str���˱��� = rsZyxnjs!ҽ����
    str���ֱ��� = rsZyxnjs!����ID
    
    If Nvl(rsZyxnjs!��ע, "0") = 0 Then
        strתԺ���� = "0"
        strת��ҽԺ�� = ""
    Else
        strתԺ���� = Mid(rsZyxnjs!��ע, 1, 1)
        If strתԺ���� = 0 Then
            strת��ҽԺ�� = ""
        Else
            strת��ҽԺ�� = Mid(rsZyxnjs!��ע, 2, 1)
        End If
    End If
    
    gstrSQL = "Select * from icd10 where ID=" & str���ֱ���
    Call OpenRecordset_OtherBase(rsZyxnjs, "ICD10", , gcnͭɽ��)
    
    If rsZyxnjs.EOF Then
        MsgBox "δָ�����ֻ�ָ���Ĳ���δ����ҽ���涨����,���ܽ���!", vbInformation, gstrSysName
        Exit Function
    End If

    str���ֱ��� = rsZyxnjs!���ֱ���
    
    gstrSQL = "Select * from ������Ϣ where ����ID=[1] And ����=[2]"
    Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "������ҳ", lng����ID, intinsure)
    
    str��Ժ���� = Format(Nvl(rsZyxnjs!��Ժʱ��, Now()), "yyyyMMdd")
    lng��ҳID = Nvl(rsZyxnjs!סԺ����, 0)
    
    gstrSQL = "Select *  From ������ Where ����ID=[1] And ��ҳID=[2] And �������=3 And ��ϴ���=1"
              
    Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ���", lng����ID, lng��ҳID)
    If rsZyxnjs.EOF Then
        str��Ժ���� = 2
    Else
    
        Select Case Nvl(rsZyxnjs!��Ժ���, "��ת")
            Case "����"
                str��Ժ���� = 1
            Case "��ת"
                str��Ժ���� = 2
            Case "δ��"
                str��Ժ���� = 3
            Case "����"
                str��Ժ���� = 4
            Case Else
                str��Ժ���� = 5
        End Select
    End If
    '>End ׼������
    
    '>>Beging ��cyjs ����
    '1 ����ռ�
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    
    '2 Ϊ������ֵ
    lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, str���˱���) '    C9  ���˱��    �α���Ա���˱��
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str���˱���, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, strҽ������) '   C3  ҽ�������  �α���ԱIC�����
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & strҽ������, lngReturn)
    
    lngReturn = tsx_setstringparam(P_ZYXH, 0, str��ˮ��)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str��ˮ��, lngReturn)

    lngReturn = tsx_setstringparam(P_LB, 0, str�����־) 'C1  ��ע    Ԥ��������(0-Ԥ����,1-��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & 0 & "," & str�����־, lngReturn)

    lngReturn = tsx_setstringparam(P_CYYMD, 0, str��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYYMD" & ", " & 0 & "," & str��Ժ����, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CYXZ, 0, str��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYXZ" & ", " & 0 & "," & str��Ժ����, lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZM, 0, str���ֱ���)
    Call WriteBusinessLOG("tsx_setstringparam", "P_BZM" & ", " & 0 & "," & str���ֱ���, lngReturn)
    
    lngReturn = tsx_setstringparam(P_ZYTZ, 0, strתԺ����) 'C1 0 - ����תԺ, 1 - ��ת, 2 - ��ת
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYTZ" & ", " & 0 & "," & strתԺ����, lngReturn)

    lngReturn = tsx_setstringparam(P_ZWYYM, 0, strת��ҽԺ��) 'C5  ת��ҽԺ��  ת��ҽԺ��ҽԺ����
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZWYYM" & ", " & 0 & "," & strת��ҽԺ��, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

    lngReturn = tsx_setstringparam(P_CZYMD, 0, str��Ժ����)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYYMD" & ", " & 0 & "," & str��Ժ����, lngReturn)

    '3 ���ýӿ�
    If tsx_jkcall("CYJS") = -1 Then
         Call WriteBusinessLOG("tsx_jkcall", "CYJS", -1)
         STRERR = tsx_getlasterr()
         MsgBox STRERR, vbInformation, gstrSysName
         Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
         lngReturn = tsx_destroyparams
         Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
         Exit Function
    Else
        Call WriteBusinessLOG("tsx_jkcall", "CYJS", lngReturn)
    '4 ȡ����ֵ
         lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl�ܷ���) 'D �ܷ���
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl�ܷ���, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_QZ_XXFD, 0, dbl�����Ը�) ' ����ҩƷ�����Ը�����
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_XXFD, 0, " & dbl�����Ը�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl�����Է�) '  D   �����Է�
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl�����Է�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_QZ_QFFD, 0, dbl�����Ը�) '�����Ը����� P_QZ_QFFDD
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_QFFD, 0, " & dbl�����Ը�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dblͳ�����֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dblͳ�����֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_QZ_JBFD, 0, dblͳ���Ը�) 'ͳ��ֶ��Ը����� P_QZ_JBFD
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_JBFD, 0, " & dblͳ���Ը�, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl��ͳ��֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl��ͳ��֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_QZ_DBFD, 0, dbl���Ը�) ' ���Ը����� P_QZ_DBFD
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_DBFD, 0, " & dbl���Ը�, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl����Ա����֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl����Ա����֧��, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl�����ʻ�֧��)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl�����ʻ�֧��, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl�ڳ������ʻ�)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl�ڳ������ʻ�, lngReturn)
    
         lngReturn = tsx_getdoubleparam(P_QZ_CFD, 0, dbl�ⶥ�Ը�) '��ⶥ���Ը����� P_QZ_CFD
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_CFD, 0, " & dbl�ⶥ�Ը�, lngReturn)
         
    End If
    '5 �����ѷ���ռ�
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
    '>>End ��cyjs ����
    
    If lngReturn <> -1 Then
        '>Beging ����ܽ���Ƿ����
        gstrSQL = "Select sum(nvl(ʵ�ս��,0))-sum(nvl(���ʽ��,0)) as δ����� From סԺ���ü�¼ Where nvl(��¼״̬,0)<>0 and ���ʷ���=1 And ����ID=[1]"
        Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "δ�����", lng����ID)
        
        If Format(Val(Nvl(rsZyxnjs.Fields!δ�����, 0)), "0.00") <> Format(dbl�ܷ���, "0.00") Then
            Dim intButton As Integer
            intButton = MsgBox("ҽԺ�ķ����ܽ��(" & Format(Nvl(rsZyxnjs.Fields!δ�����, 0), "0.00") & ")��ҽ�����ĵķ����ܶ�(" & Format(dbl�ܷ���, "0.00") & ")���ȣ��Ƿ������" & vbNewLine & _
                             "ѡ[��]�����Դ����⣬�������㡣" & vbNewLine & _
                             "ѡ[��]��ֹͣ����������ش�������ϸ���������Ժ����½��㡣" & vbNewLine & _
                             "ѡ[ȡ��]��ֹͣ����������������ֹ�ȷ�Ϸ��ú������½��㡣", vbQuestion Or vbYesNoCancel Or vbDefaultButton2, gstrSysName)
            If intButton = vbNo Then
                gstrSQL = "Update סԺ���ü�¼ Set �Ƿ��ϴ�=0 Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID & " And �Ƿ��ϴ�=1"
                gcnOracle.Execute gstrSQL
                Exit Function
            ElseIf intButton = vbCancel Then
                Exit Function
            End If
        End If
        ''>End ����ܽ���Ƿ����
        
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_ͭɽ�� & ",'�ʻ����','''" & Format(dbl�ڳ������ʻ�, "0.00") & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ����")
        
        If mstrʹ�ø����ʻ�֧�� = 1 Then
            If dbl�ڳ������ʻ� >= dbl�����Ը� Then
                dbl�����ʻ�֧�� = dbl�����Ը�
                dbl�����Ը� = 0
            Else
                dbl�����ʻ�֧�� = dbl�ڳ������ʻ�
                dbl�����Ը� = dbl�����Ը� - dbl�����ʻ�֧��
            End If
        End If
        
        str���㷽ʽ = "ͳ�����;" & dblͳ�����֧�� & ";0"
        str���㷽ʽ = str���㷽ʽ & "|��֧��;" & dbl��ͳ��֧�� & ";0"
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & dbl����Ա����֧�� & ";0"
        str���㷽ʽ = str���㷽ʽ & "|�����ʻ�;" & dbl�����ʻ�֧�� & ";0"
        
        סԺ�������_ͭɽ�� = str���㷽ʽ
        dbl��ĩ�����ʻ� = dbl�ڳ������ʻ� - dbl�����ʻ�֧��

        '>>Beging д��xybjz��,����һ�����嵥��
        gstrSQL = "Delete xybjz where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
        gcnͭɽ��.Execute gstrSQL
        
        gstrSQL = "Insert into Xybjz(����id,��ҳid,�ܶ�, ��������, ��ȫ����, ͳ�����֧��,�󲡻���֧��, ����Աͳ��֧��,����֧��,ͳ�����֧��,�󲡻���֧��,�ⶥ���Ը�) values(" & _
                 lng����ID & "," & lng��ҳID & "," & dbl�ܷ��� & "," & dbl�����Ը� & "," & dbl�����Է� & "," & dblͳ�����֧�� & "," & _
                 dbl��ͳ��֧�� & "," & dbl����Ա����֧�� & "," & dbl�����Ը� & "," & dblͳ���Ը� & "," & dbl���Ը� & "," & dbl�ⶥ�Ը� & ")"
        gcnͭɽ��.Execute gstrSQL
        
        
        '>>End д��xybjz��,����һ�����嵥��
        
    Else
        MsgBox "Ԥ����ʧ��!", vbInformation, gstrSysName
        סԺ�������_ͭɽ�� = ""
    End If
    'Set rsZyxnjs = Nothing
    
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_ͭɽ��(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'********************************

'�����ߡ���������������clsInsure �� SettleDelSwap ���̵���
'����˵����������������ɱ���סԺ���������
'����˵��������������1������סԺ�������Ͻӿ�
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string�Ͳ���
'������tsx_setdoubleparam������double�Ͳ���
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_getstringparam��ȡstring�ͽӿڷ���ֵ
'������tsx_destroyparams�������ѷ���ռ�
'*********************************

    'TODO:סԺ�������
    Dim rsZyjsCx As New ADODB.Recordset, bln������� As Boolean, lng����ID As Long, lng����ID As Long
    Dim strԭ��ˮ�� As String, str���˱�� As String, strҽ������ As String
    Dim lngReturn As Long, blnǷ���ؽ� As Boolean, STRERR As String
    
    On Error GoTo errHand
    
    gstrSQL = "Select * From ���㷽ʽ Where ���� In (" & _
                    "Select ���㷽ʽ From ����Ԥ����¼ " & _
                    "where ����ID=[1]) And ����>=3 And  ����<=4"
    Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, "���㷽ʽ", lng����ID)
    If Not rsZyjsCx.EOF Then
        'MsgBox "��ҽ����֧�ֳ�������!" & vbCrLf & "��Ҫ�����ѽ��㵥�ݣ��뵽ҽ�����İ���", vbInformation, gstrSysName
        bln������� = True
    End If
    
    gstrSQL = "Select * From ���ս����¼ Where ����=2 And ��¼ID=[1] and ����=[2]"
    Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������ˮ��", lng����ID, intinsure)
    If rsZyjsCx.EOF Then
        Err.Raise 9000, gstrSysName, "û���ҵ�ԭʼ�����¼���޷�����סԺ���������"
        Exit Function
    End If
    lng����ID = rsZyjsCx!����ID
    strԭ��ˮ�� = rsZyjsCx!֧��˳���
    
    If bln������� = True Then
        blnǷ���ؽ� = False
    Else
        'Ƿ���ؽ�
        If MsgBox("ȷ���ǰ���Ƿ���ؽύ����", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            blnǷ���ؽ� = True
        Else
            blnǷ���ؽ� = False
        End If
    End If
    
    If blnǷ���ؽ� = True Then
      
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_ͭɽ�� & ",'Ƿ���ؽ�ID','" & lng����ID & " ')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ƿ���ؽ�")
        סԺ�������_ͭɽ�� = True
    Else
        '�������
        gstrSQL = "Select Distinct A.ID From ���˽��ʼ�¼ A,���˽��ʼ�¼ B Where A.No=B.No And A.��¼״̬=2 And B.Id=[1]"
        Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
        lng����ID = rsZyjsCx("ID")
        
        gstrSQL = "Select * from �����ʻ� where ����ID=[1] And ����=[2]"
        Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", lng����ID, TYPE_ͭɽ��)
        str���˱�� = rsZyjsCx!ҽ����
        strҽ������ = rsZyjsCx!����
        '1 ����ռ�
        If tsx_createparams(1024, 1024) = -1 Then
            Err.Raise 9000, gstrSysName, "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
        '2 Ϊ������ֵ
        lngReturn = tsx_setstringparam(P_JGM, 0, g������Ϣ_ͭɽ.ҽԺ��)
        Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
        
        lngReturn = tsx_setstringparam(P_TBR, 0, str���˱��) '    C9  ���˱��    �α���Ա���˱��
        Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str���˱��, lngReturn)
        
        lngReturn = tsx_setstringparam(P_KXH, 0, strҽ������) '   C3  ҽ�������  �α���ԱIC�����
        Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & strҽ������, lngReturn)
        
        lngReturn = tsx_setstringparam(P_DJH, 0, strԭ��ˮ��) '   ��ˮ��
        Call WriteBusinessLOG("tsx_setstringparam", "P_DJH" & ", " & 0 & "," & strԭ��ˮ��, lngReturn)
        
        lngReturn = tsx_setstringparam(P_CZRYH, 0, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
        Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g������Ϣ_ͭɽ.����Ա��, lngReturn)

        '3 ���ýӿ�
        If tsx_jkcall("ZYCZ") = -1 Then
             Call WriteBusinessLOG("tsx_jkcall", "ZYCZ", -1)
             STRERR = tsx_getlasterr()
             Err.Raise 9000, gstrSysName, tsx_getlasterr(), vbInformation, gstrSysName
             Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
             lngReturn = tsx_destroyparams
             Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
             Exit Function
        Else
            Call WriteBusinessLOG("tsx_jkcall", "ZYCZ", lngReturn)
            '4 ȡ����ֵ
            lngReturn = tsx_getstringparam(P_DJH, 0, strԭ��ˮ��)
            Call WriteBusinessLOG("tsx_getstringparam", "P_RYH, 0, " & strԭ��ˮ��, lngReturn)
             
        End If
        
        If lngReturn <> -1 Then
            gstrSQL = "Select * From ���ս����¼ Where ����=2 And ��¼ID=[1] and ����=[2]"
            Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, "���ս����¼", lng����ID, TYPE_ͭɽ��)
     
            gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_ͭɽ�� & "," & lng����ID & "," & _
                Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
                0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
                -1 * Nvl(rsZyjsCx!�������ý��, 0) & "," & _
                -1 * Nvl(rsZyjsCx!ȫ�Ը����, 0) & "," & _
                -1 * Nvl(rsZyjsCx!�����Ը����, 0) & "," & _
                -1 * Nvl(rsZyjsCx!����ͳ����, 0) & "," & _
                -1 * Nvl(rsZyjsCx!ͳ�ﱨ�����, 0) & "," & _
                -1 * Nvl(rsZyjsCx!���Ը����, 0) & "," & 0 & "," & _
                -1 * Nvl(rsZyjsCx!�����ʻ�֧��, 0) & ",'" & strԭ��ˮ�� & "',null,null,Null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
           
        End If
        '5 �����ѷ���ռ�
        lngReturn = tsx_destroyparams()
        Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

        סԺ�������_ͭɽ�� = True
    End If
    Set rsZyjsCx = Nothing
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function �����ϴ�_ͭɽ��(ByVal int��� As Integer, ByVal str���ݺ� As String, ByVal int���� As Integer, ByVal int״̬ As Integer, _
        str��Ϣ As String, Optional ByVal lng����ID As Long = 0, Optional ByVal intinsure As Integer = 0) As Boolean

'*******************************************************
'�����ߡ���������������clsInsure �� TranChargeDetail ���̵���
'����˵��������������סԺ���ʱ���ʱ�򱣴�󣬸��ݲ���������support...���ɲμ�Getcapability��
'����˵��������������1����ȡ�����ݵĴ�����ϸ
'��������������������2�����ϴ���ҽ���Ĳ��˴���
'��������������������3�����ݽӿ����ʣ�ÿ�������ϴ������ϴ��������ɹ��ϴ�����ϸ�����ϴ����
'���ù����嵥��˵����
'������tsx_createparams������ռ�
'������tsx_setstringparam������string�Ͳ���
'������tsx_setdoubleparam������double�Ͳ���
'������tsx_setlongparam������long�Ͳ���
'������tsx_getlasterr�� ȡ���ϴδ�����Ϣ
'������tsx_jkcall�� ���ýӿ�
'������tsx_getstringparam��ȡstring�ͽӿڷ���ֵ
'������tsx_getdoubleparam��ȡdouble�ͽӿڷ���ֵ
'������tsx_destroyparams�������ѷ���ռ�
'*******************************************************
    
    'TODO:�����ϴ�
    '�����Ǵ��ϴ���ǵ�ʾ�����루�����ֲ�ͬ�ķ�ʽ���ɸ��������Ҫʹ�ã�
    'ע���������Ǳ����ͬʱ�ϴ�����Ϊ����û���ύ����ʹ��ȫ�����Ӷ���gcnOracle�����ϴ���־��
    '����Ǳ�����ϴ��������´�һ�����Ӷ���Ҳ����ʹ��ȫ�����Ӷ���gcnOracle�����ϴ���־
'    gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & !NO & "'," & !��� & "," & !��¼���� & "," & !��¼״̬ & ")"
'    cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
'    '����
    Dim rs������ϸ As New ADODB.Recordset, rsMzxnjs As New ADODB.Recordset, rs���ü�¼ As New ADODB.Recordset, rsXybmx As New ADODB.Recordset
    Dim rsCfsc As New ADODB.Recordset, lngPatiID As Long '����ID,Ϊ�˺Ͳ���������,����ĸ�����ϴ����ʱ�
    Dim str��ˮ�� As String, lngReturn As Long, strҩƷ As String, lngPageID As Long
    Dim str���˱�� As String, str�������� As String, blnȫ������ As Boolean, dblҽ��Ƿ�� As Double
    Dim lngCount As Long '��¼���
    Dim bln�Ƿ���ýӿ� As Boolean
    Dim STRERR As String, str���� As String
    
    �����ϴ�_ͭɽ�� = True
        '����NO��,��ȡ����ID
    
    
    gstrSQL = "Select distinct  decode(nvl(A.ҽ�����,-99),-99,to_char(A.����ʱ��,'YYYY-MM-DD'),to_char(A.�Ǽ�ʱ��,'YYYY-MM-DD')) as ��������,A.����ID,A.��ҳID from סԺ���ü�¼ A,�����ʻ� B " & _
            "where A.����ID=B.����ID And nvl(A.���ӱ�־,0)<>9 And Nvl(a.ʵ�ս��, 0)<>0" & _
            " And A.��¼����=[1] and  A.��¼״̬=[2] And A.NO=[3]" & _
            " And B.����=[4]"
    If lng����ID <> 0 Then
        gstrSQL = gstrSQL & " And A.����ID=[5]"
    End If
    Call WriteBusinessLOG("ȡҪ�ϴ�����(�����ϴ�)", gstrSQL, "")
    Set rsCfsc = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", int����, int״̬, str���ݺ�, intinsure, lng����ID)
    
    Do Until rsCfsc.EOF
        lngPatiID = rsCfsc!����ID
        lngPageID = rsCfsc!��ҳID
        
        str�������� = rsCfsc!��������
        
        gstrSQL = "Select * from �����ʻ� where ����=[1] And ����ID=[2]"
        Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", intinsure, lngPatiID)
        
        str��ˮ�� = rsMzxnjs!˳���
        
        gstrSQL = "Select A.�շ�ϸĿID,decode(nvl(A.ҽ�����,-99),-99,to_char(A.����ʱ��,'YYYY-MM-DD'),to_char(A.�Ǽ�ʱ��,'YYYY-MM-DD')) as ����ʱ��,max(nvl(����,0)) as ����,decode(round(sum(nvl(A.����,1)*nvl(A.����,1))),0,1,round(sum(nvl(A.����,1)*nvl(A.����,1)))) as ����,sum(nvl(A.ʵ�ս��,0)) as ���," & _
                                  "sum(nvl(A.ʵ�ս��,0))/decode(round(sum(nvl(A.����,1)*nvl(A.����,1))),0,1,round(sum(nvl(A.����,1)*nvl(A.����,1))))as �۸�,max(C.����) as �������� " & _
                          " from סԺ���ü�¼ A,�����ʻ� B,���ű� C " & _
                          " where decode(nvl(A.ҽ�����,-99),-99,to_char(A.����ʱ��,'YYYY-MM-DD'),to_char(A.�Ǽ�ʱ��,'YYYY-MM-DD'))='" & str�������� & _
                                "' And nvl(A.���ӱ�־,0)<>9 And A.���ʷ���=1" & _
                                " And A.��¼״̬=" & 1 & _
                                " And A.����ID=B.����ID " & _
                                " and B.����=" & intinsure & _
                                " and A.���˲���ID=C.ID " & _
                                " ANd A.����ID=[1]" & _
                                " And A.��ҳID=[2]" & _
                                " And nvl(A.ʵ�ս��,0)<>0 " & _
                                " Group by A.�շ�ϸĿID,decode(nvl(A.ҽ�����,-99),-99,to_char(A.����ʱ��,'YYYY-MM-DD'),to_char(A.�Ǽ�ʱ��,'YYYY-MM-DD'))"
        Call WriteBusinessLOG("����ϸ", gstrSQL, "")
        Set rs������ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ", lngPatiID, lngPageID)
        lngCount = 0
        '1 ����ռ�
        Call WriteBusinessLOG("׼������ռ�", "1024*30", "")
        If tsx_createparams(1024 * 30, 1024 * 30) = -1 Then
            MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
        
        If rs������ϸ.EOF Then
            '����ķ���ȫ����������,�ʹ�һ��0��ȥ
            blnȫ������ = True
        End If
        
        Do Until rs������ϸ.EOF
            '2 Beging Ϊ������ֵ
            lngReturn = tsx_setstringparam(P_JGM, lngCount, g������Ϣ_ͭɽ.ҽԺ��)
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCount & "," & g������Ϣ_ͭɽ.ҽԺ��, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCount & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_ZYXH, lngCount, str��ˮ��)
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & lngCount & "," & str��ˮ��, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & lngCount & "," & str��ˮ��, lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_FYYMD, lngCount, Format(rs������ϸ!����ʱ��, "yyyyMMdd"))
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_FYYMD" & ", " & lngCount & "," & Format(rs������ϸ!����ʱ��, "yyyyMMdd"), STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_FYYMD" & ", " & lngCount & "," & Format(rs������ϸ!����ʱ��, "yyyyMMdd"), lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_RYBQ, lngCount, rs������ϸ!��������)
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & lngCount & "," & rs������ϸ!��������, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & lngCount & "," & rs������ϸ!��������, lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_RYCWH, lngCount, Nvl(rs������ϸ!����, 0))
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & lngCount & "," & Nvl(rs������ϸ!����, 0), STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & lngCount & "," & Nvl(rs������ϸ!����, 0), lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_JZLB, lngCount, 0) '�������    ��Ϊ0
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_JZLB" & ", " & lngCount & "," & 0, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_JZLB" & ", " & lngCount & "," & 0, lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_CZRYH, lngCount, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH," & lngCount & "," & g������Ϣ_ͭɽ.����Ա��, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH," & lngCount & "," & g������Ϣ_ͭɽ.����Ա��, lngReturn)
            End If
            
            gstrSQL = "Select * from ypzlk where �շ�ϸĿID=" & rs������ϸ!�շ�ϸĿID
            Call OpenRecordset_OtherBase(rsMzxnjs, "ypzlk", gstrSQL, gcnͭɽ��)
            If rsMzxnjs.EOF = False Then
                lngReturn = tsx_setstringparam(P_ZBM, lngCount, rsMzxnjs!�Ա���) 'C20 �Ա���  ������ϸ�Ա���
                If lngReturn = -1 Then
                    STRERR = tsx_getlasterr()
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rsMzxnjs!�Ա���, STRERR)
                Else
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rsMzxnjs!�Ա���, lngReturn)
                End If
            
            Else
                lngReturn = tsx_setstringparam(P_ZBM, lngCount, rs������ϸ!�շ�ϸĿID) 'C20 �Ա���  ������ϸ�Ա���
                
                If lngReturn = -1 Then
                    STRERR = tsx_getlasterr()
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rs������ϸ!�շ�ϸĿID, STRERR)
                Else
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rs������ϸ!�շ�ϸĿID, lngReturn)
                End If
            End If
            gstrSQL = "select A.*,B.ҩƷ���� from �շ�ϸĿ A,ҩƷ��Ϣ B,ҩƷĿ¼ C " & _
                      " where A.id=C.ҩƷID(+) and C.ҩ��ID=B.ҩ��ID(+) And A.id=[1]"
            Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ŀ���", CLng(rs������ϸ!�շ�ϸĿID))
            Select Case rsMzxnjs!���s
                Case "5", "6", "7"
                    strҩƷ = "0"
                Case Else
                    strҩƷ = "1"
            End Select
            lngReturn = tsx_setstringparam(P_LB, lngCount, strҩƷ)
            Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & lngCount & "," & strҩƷ, lngReturn)
            If lngReturn = -1 Then
                Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
            End If
            
            lngReturn = tsx_setdoubleparam(P_JG, lngCount, rs������ϸ!�۸�)
            Call WriteBusinessLOG("tsx_setdoubleparam", "P_JG" & ", " & lngCount & "," & rs������ϸ!�۸�, lngReturn)
            If lngReturn = -1 Then
                Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
            End If
             
            lngReturn = tsx_setlongparam(P_SL, lngCount, rs������ϸ!����)
            Call WriteBusinessLOG("tsx_setlongparam", "P_SL" & ", " & lngCount & "," & rs������ϸ!����, lngReturn)
            If lngReturn = -1 Then
                Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
            End If
            lngReturn = tsx_setstringparam(P_CZYMD, lngCount, Format(rs������ϸ!����ʱ��, "yyyyMMdd") & Format(Now(), "HHmmss"))
            Call WriteBusinessLOG("tsx_setstringparam", "P_CZYMD" & ", " & lngCount & "," & Format(rs������ϸ!����ʱ��, "yyyyMMdd") & Format(Now(), "HHmmss"), lngReturn)
            If lngReturn = -1 Then
                Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
            End If
            
            lngCount = lngCount + 1
            rs������ϸ.MoveNext
        Loop
        '2 End Ϊ������ֵ
        
        If blnȫ������ = True Then
            gstrSQL = "Select A.�շ�ϸĿID,decode(nvl(A.ҽ�����,-99),-99,to_char(A.����ʱ��,'YYYY-MM-DD'),to_char(A.�Ǽ�ʱ��,'YYYY-MM-DD')) as ����ʱ��,max(nvl(����,0)) as ����,sum(nvl(A.����,1)*nvl(A.����,0)) as ����,sum(nvl(A.ʵ�ս��,0)) as ���," & _
                              "sum(nvl(A.ʵ�ս��,0))/round(sum((nvl(A.����,1)*nvl(A.����,0)))) as �۸�,max(C.����) as �������� " & _
                      " from סԺ���ü�¼ A,�����ʻ� B,���ű� C " & _
                      " where decode(nvl(A.ҽ�����,-99),-99,to_char(A.����ʱ��,'YYYY-MM-DD'),to_char(A.�Ǽ�ʱ��,'YYYY-MM-DD'))='" & str�������� & _
                            "' And nvl(A.���ӱ�־,0)<>9 And A.��¼����=" & int���� & _
                            " And A.��¼״̬=" & int״̬ & _
                            " And A.����ID=B.����ID " & _
                            " and B.����=" & intinsure & _
                            " and A.���˲���ID=C.ID " & _
                            " ANd A.����ID=" & lngPatiID & _
                            " And A.��ҳID=" & lngPageID & _
                            " And nvl(A.ʵ�ս��,0)<>0 " & _
                            " Group by A.�շ�ϸĿID,decode(nvl(A.ҽ�����,-99),-99,to_char(A.����ʱ��,'YYYY-MM-DD'),to_char(A.�Ǽ�ʱ��,'YYYY-MM-DD'))"
            Call WriteBusinessLOG("������ϸ", gstrSQL, "")
                        
'            Call OpenRecordset(rs������ϸ, "������ϸ", gstrSQL)
            
                lngReturn = tsx_setstringparam(P_JGM, lngCount, g������Ϣ_ͭɽ.ҽԺ��)
                Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCount & "," & g������Ϣ_ͭɽ.ҽԺ��, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                
                lngReturn = tsx_setstringparam(P_ZYXH, lngCount, str��ˮ��)
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & lngCount & "," & str��ˮ��, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                
                lngReturn = tsx_setstringparam(P_FYYMD, lngCount, Replace(str��������, "-", ""))
                Call WriteBusinessLOG("tsx_setstringparam", "P_FYYMD" & ", " & lngCount & "," & Replace(str��������, "-", ""), lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                
                gstrSQL = "Select B.���� from ������ҳ A,���ű� B where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
                Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "zlypk", lngPatiID, lngPageID)
                If rsMzxnjs.EOF = False Then
                    str���� = rsMzxnjs!����
                    lngReturn = tsx_setstringparam(P_RYBQ, lngCount, str����)
                    Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & lngCount & "," & str����, lngReturn)
                    If lngReturn = -1 Then
                        Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                    End If
                Else
                    Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & lngCount & "," & "δ�ҵ�����", lngReturn)
                    If lngReturn = -1 Then
                        Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                    End If
                End If
                 
                lngReturn = tsx_setstringparam(P_RYCWH, lngCount, 0)
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & lngCount & "," & 0, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                 
                lngReturn = tsx_setstringparam(P_JZLB, lngCount, 0) '�������    ��Ϊ0
                Call WriteBusinessLOG("tsx_setstringparam", "P_JZLB" & ", " & lngCount & "," & 0, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                
                lngReturn = tsx_setstringparam(P_CZRYH, lngCount, g������Ϣ_ͭɽ.����Ա��) ' C10 ������Ա��
                Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH," & lngCount & "," & g������Ϣ_ͭɽ.����Ա��, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                
                gstrSQL = "Select * from ypzlk where rownum=1"
                Call OpenRecordset_OtherBase(rsMzxnjs, "zlypk", gstrSQL, gcnͭɽ��)
                If rsMzxnjs.EOF = False Then
                    lngReturn = tsx_setstringparam(P_ZBM, lngCount, rsMzxnjs!�Ա���) 'C20 �Ա���  ������ϸ�Ա���
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rsMzxnjs!�Ա���, lngReturn)
                    If lngReturn = -1 Then
                        Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                    End If
                Else
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & "δ�ҵ��Ա���", lngReturn)
                    If lngReturn = -1 Then
                        Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                    End If
                End If
                
                strҩƷ = "0"
                
                lngReturn = tsx_setstringparam(P_LB, lngCount, strҩƷ)
                Call WriteBusinessLOG("tsx_setstringparam", "P_JG" & ", " & lngCount & "," & strҩƷ, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                
                lngReturn = tsx_setdoubleparam(P_JG, lngCount, 0)
                Call WriteBusinessLOG("tsx_setdoubleparam", "P_LB" & ", " & lngCount & "," & 0, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                 
                lngReturn = tsx_setlongparam(P_SL, lngCount, 0)
                Call WriteBusinessLOG("tsx_setlongparam", "P_SL" & ", " & lngCount & "," & 0, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
                lngReturn = tsx_setstringparam(P_CZYMD, lngCount, Replace(str��������, "-", "") & Format(Now(), "HHmmss"))
                Call WriteBusinessLOG("tsx_setstringparam", "P_CZYMD" & ", " & lngCount & "," & Replace(str��������, "-", "") & Format(Now(), "HHmmss"), lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("������Ϣ", "", tsx_getlasterr())
                End If
      End If

        '3 ���ýӿ�
            If tsx_jkcall("ZYMX_SC") = -1 Then
                 Call WriteBusinessLOG("tsx_jkcall", "ZYMX_SC", -1)
                 STRERR = tsx_getlasterr()
                 MsgBox STRERR, vbInformation, gstrSysName
                 Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
                 lngReturn = tsx_destroyparams
                 Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
                 Exit Function
            Else
            Call WriteBusinessLOG("tsx_jkcall", "ZYMX_SC", lngReturn)
            '4 ȡ����ֵ
            
            '�޷���ֵ
                ''>>begin ����HIS���ϴ���־
                
                If lngReturn = -1 Then
                    If blnȫ������ = False Then
                        MsgBox "�ϴ���ϸ(" & rsMzxnjs!��� & "_" & rsMzxnjs!���� & ")" & rsMzxnjs!���� & "ʧ�ܡ�" & vbCrLf & Trim(tsx_getlasterr), vbInformation, gstrSysName
                    Else
                        Call WriteBusinessLOG("ȫ������", "��¼�����", "-1")
                    End If
                Else
                    gstrSQL = "Select distinct A.ID " & _
                      " from סԺ���ü�¼ A,�����ʻ� B,���ű� C " & _
                      " where decode(nvl(A.ҽ�����,-99),-99,to_char(A.����ʱ��,'YYYY-MM-DD'),to_char(A.�Ǽ�ʱ��,'YYYY-MM-DD'))='" & str�������� & _
                            "'And nvl(A.���ӱ�־,0)<>9 And a.���ʷ���=1" & _
                            " And A.NO='" & str���ݺ� & "' " & _
                            " and A.��¼����=" & int���� & _
                            " And A.��¼״̬=" & int״̬ & _
                            " And A.����ID=B.����ID " & _
                            " and B.����=" & intinsure & _
                            " and A.���˲���ID=C.ID " & _
                            " ANd A.����ID=[1]" & _
                            " And A.��ҳID=[2]" & _
                            " And nvl(A.ʵ�ս��,0)<>0 "
                    Call WriteBusinessLOG("��д�ϴ���־�ļ�¼", gstrSQL, "")
                    Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "���ϴ���ϸ��ID", lngPatiID, lngPageID)
                    Do Until rsMzxnjs.EOF
                        gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rsMzxnjs!ID & "," & 0 & ",NULL)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ���ֶ�")
                        
                        '>>Beging д��XYBMX����,���ڴ�ӡһ���嵥
                        
                        '>>>Beging ȡxybmx���״̬
                        gstrSQL = "Select * from �����ʻ� where ����=[1] And ����ID=[2]"
                        Call WriteBusinessLOG("ȡxybmx���״̬", gstrSQL, "")
                        Set rsXybmx = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", TYPE_ͭɽ��, lngPatiID)
                        str���˱�� = rsXybmx!ҽ����
                        
                        '1 ����ռ�
                        If tsx_createparams(1024, 1024) = -1 Then
                            MsgBox "�����ڴ�ռ�ʧ��!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)

                         '2 Ϊ������ֵ
                        
                         lngReturn = tsx_setstringparam(P_TBR, 0, str���˱��) '    C9  ���˱��    �α���Ա���˱��
                         Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str���˱��, lngReturn)
                         
                         lngReturn = tsx_setstringparam(P_LB, 0, 0) '
                         Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & 0 & "," & 0, lngReturn)
                                  
                        '3 ���ýӿ�
                        If tsx_jkcall("GETCBRYXX_T") = -1 Then
                             Call WriteBusinessLOG("tsx_jkcall", "GETCBRYXX_T", -1)
                             STRERR = tsx_getlasterr()
                             MsgBox STRERR, vbInformation, gstrSysName
                             Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
                             lngReturn = tsx_destroyparams
                             Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
                             Exit Function
                        Else
                            Call WriteBusinessLOG("tsx_jkcall", "GETCBRYXX_T", 1)
                        '4 ȡ����ֵ
                             lngReturn = tsx_getdoubleparam(P_GRQK, 0, dblҽ��Ƿ��)
                             Call WriteBusinessLOG("tsx_getstringparam", "P_RYH, 0, " & dblҽ��Ƿ��, lngReturn)
                             
                        End If
                        '5 �����ѷ���ռ�
                        lngReturn = tsx_destroyparams()
                        Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

                         
                         '>>>End ȡxybmx���״̬
                        
                        gstrSQL = "Select b.����, b.���㵥λ As ��λ, Nvl(a.ʵ�ս��, 0) As ʵ�ս��," & _
                                  "Decode(Nvl(c.��ע, '�Է�'), '����', 0, '����', 0.2, 1) * Nvl(a.ʵ�ս��, 0) As �Էѽ�� ," & _
                                  "Decode(Round(Nvl(a.����, 1) * Nvl(a.����, 1)), 0, 1, Round(Nvl(a.����, 1) * Nvl(a.����, 1))) As ����," & _
                                  "Nvl(a.ʵ�ս��, 0) /Decode(Round(Nvl(a.����, 1) * Nvl(a.����, 1)), 0, 1, Round(Nvl(a.����, 1) * Nvl(a.����, 1))) As �۸�," & _
                                  "To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����" & _
                                  " From סԺ���ü�¼ A,�շ�ϸĿ B,(Select * From ����֧����Ŀ where ����=[1]) C " & _
                                  " Where  Nvl(a.ʵ�ս��, 0) <>0 and A.�շ�ϸĿID=B.ID And A.�շ�ϸĿID=C.�շ�ϸĿID(+) and A.ID=[2]"
                        Call WriteBusinessLOG("�������ϸ׼��дxybmx", gstrSQL, "")
                        Set rs���ü�¼ = zlDatabase.OpenSQLRecord(gstrSQL, "���ü�¼", TYPE_ͭɽ��, CLng(rsMzxnjs!ID))
                        
                        If rs���ü�¼.EOF = False Then
                            gstrSQL = "Select * from XYBMX Where ��¼ID=" & rsMzxnjs!ID
                            Call WriteBusinessLOG("��xybmx", gstrSQL, "")
                            Call OpenRecordset_OtherBase(rsXybmx, "���ü�¼", , gcnͭɽ��)
                            
                            If rsXybmx.EOF Then
                                gstrSQL = "Insert into XYBMX(��Ŀ����,��λ,����,����,�ϼƽ��,�Ը����,����id,��ҳid,����,��¼id,ҽ��״̬) values('" & _
                                        rs���ü�¼!���� & "','" & rs���ü�¼!��λ & "'," & rs���ü�¼!�۸� & "," & rs���ü�¼!���� & "," & rs���ü�¼!ʵ�ս�� & "," & rs���ü�¼!�Էѽ�� & "," & lngPatiID & "," & lngPageID & "," & _
                                        "to_date('" & Format(rs���ü�¼!����, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & rsMzxnjs!ID & "," & dblҽ��Ƿ�� & ")"
                                Call WriteBusinessLOG("дxybmx", gstrSQL, "")
                                gcnͭɽ��.Execute gstrSQL
                            End If
                        
                        End If
                        
                        '>>End д��XYBMX����,���ڴ�ӡһ���嵥
                        
                        rsMzxnjs.MoveNext
                    Loop
                End If 'lngReturn = -1
                '>>End ����HIS���ϴ���־
                
            End If 'tsx_jkcall("ZYMX_SC")
                      '5 �����ѷ���ռ�
        lngReturn = tsx_destroyparams()
        Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

        rsCfsc.MoveNext
    Loop
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ���˱䶯_ͭɽ��(lngPatiID As Long, lngPageID As Long, ByVal intinsure As Integer) As Boolean

'*****************************************************************************
'�����ߡ�������������סԺ��Ϣ��������
'����˵����������������Ժ���˵Ĵ�λ�䶯,ת��,ҽ���仯���������������Ϣ�仯ʱ���ô˽ӿ�
'����˵��������������1����ȡ���˱䶯��Ϣ
'��������������������2�����ϴ���ҽ���Ĳ��˱䶯���
'��������������������3�����ݽӿ�Ҫ��,������Ӧ�����ϴ��䶯���
'���ù����嵥��˵����
'�������ޡ�
''*****************************************************************************
    '//TODO:���˱䶯
    '�����Ǵ��ϴ���ǵ�ʾ�����루�����ֲ�ͬ�ķ�ʽ���ɸ��������Ҫʹ�ã�
    'ע������:(��)
    ���˱䶯_ͭɽ�� = True
End Function

Public Function �ʻ�תԤ��_ͭɽ��(lngԤ��ID As Long, curMoney As Currency, ByVal intinsure As Integer) As Boolean
'*****************************************************************************
'�����ߡ�������������
'����˵������������������Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'����˵��������������1����ȡ���˱䶯��Ϣ
'��������������������2�����ϴ���ҽ���Ĳ��˱䶯���
'��������������������3�����ݽӿ�Ҫ��,������Ӧ�����ϴ��䶯���
'���ù����嵥��˵����
'��������������ҽ�������ر�
''*****************************************************************************
    '//TODO:�ʻ�תԤ��
    '
    'ע������:(��)

End Function

Public Function �ʻ�תԤ������_ͭɽ��(lngԤ��ID As Long, ByVal intinsure As Integer) As Boolean
'*****************************************************************************
'�����ߡ�������������
'����˵������������������Ҫ��Ԥ����ת������ʻ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'����˵��������������1����ȡ���˱䶯��Ϣ
'��������������������2�����ϴ���ҽ���Ĳ��˱䶯���
'��������������������3�����ݽӿ�Ҫ��,������Ӧ�����ϴ��䶯���
'���ù����嵥��˵����
'��������������ҽ�������ر�
''*****************************************************************************
    '//TODO:�ʻ�תԤ������
    '
    'ע������:(��)
End Function

Public Function ����ѡ��_ͭɽ��(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal intinsure As Integer)
'*****************************************************************************
'�����ߡ���������������clsInsure ��  ChooseDisease  ���̵���
'����˵��������������ѡ���˵ĳ�Ժ����
'���ù����嵥��˵����
'��������������ҽ�������ر�
''*****************************************************************************
'//TODO:����ѡ����ҽ��ǰ̨���������д˹���
    
End Function

Public Function ҽ����������_ͭɽ��(ByVal capҵ�� As ҽԺҵ��) As Boolean
'*****************************************************************************
'�����ߡ���������������clsInsure ��  GetCapability  ���̵���
'����˵���������������жϺ��������һЩҵ���ڲ�ͬ��ҽ������Ƿ�õ�֧��
'���ù����嵥��˵����
'�������ޡ�
''*****************************************************************************
'TODO:��������
    Select Case capҵ��
    
        Case support�����˷�, _
             support�����ϴ�, _
             support���������ϴ�, _
             supportҽ���ϴ�, _
             support������ɺ��ϴ�, _
             support������봫����ϸ, _
             support��Ժ��������Ժ, _
             supportδ�����Ժ, _
             support����ʹ�ø����ʻ�, _
             support����¼��������, _
             support������Ժ, _
             support�����˸����ʻ�, _
             support��Ժ���˽�������, _
             support����Ԥ��, _
             support�����ݳ�������, _
             support��������
             ҽ����������_ͭɽ�� = True
             
       Case support�����������, supportסԺ��������
             ҽ����������_ͭɽ�� = True
'
    End Select
    
End Function

Public Function ȡ������_ͭɽ��(ByVal bytType As Byte, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
'*****************************************************************************
'�����ߡ���������������clsInsure ��  IdentifyCancel  ���̵���
'����˵����������������������ҽ�����������֤�ɹ���ȡ����������
'���ù����嵥��˵����
'�������ޡ�
''*****************************************************************************
'TODO:ȡ������
    ȡ������_ͭɽ�� = True
    
End Function


Public Function tsx_getlasterr() As String
    Dim lngRetu As Long, STRERR As String
    
    STRERR = Space(512)
    lngRetu = tsx_getlasterr2(STRERR)
    tsx_getlasterr = STRERR
    
End Function

Public Function ҽ����Ŀ_ͭɽ��(����ID As Long, �շ�ϸĿID As Long, ��� As Currency, _
    ByVal bln���� As Boolean, Optional ByVal intinsure As Integer) As String
'*****************************************************************************
'�����ߡ���������������clsInsure ��  GetItemInsure  ���̵���
'����˵����������������Ҫ����ǰ̨��ʾ��ҽ���ķ������ͼ�����
'���ù����嵥��˵����
'�������ޡ�
''*****************************************************************************
'TODO:ҽ����Ŀ��Ϣ
    Dim rsYbxm As New ADODB.Recordset
    gstrSQL = "Select * from ypzlk where �շ�ϸĿID=" & �շ�ϸĿID
    Call OpenRecordset_OtherBase(rsYbxm, "ypzlk", , gcnͭɽ��)
    If rsYbxm.EOF = False Then
        ҽ����Ŀ_ͭɽ�� = rsYbxm!֧�����
    Else
        ҽ����Ŀ_ͭɽ�� = "δ����"
    End If
    Set rsYbxm = Nothing
    
End Function

Public Function ҽ����Ϣ_ͭɽ��(ByVal lngItemID As Long, Optional intType As Integer = 0) As String
'*****************************************************************************
'�����ߡ���������������clsInsure ��  GetItemInfo  ���̵���
'����˵����������������Ҫ����ǰ̨��ʾ��ҽ���ķ������ͼ�����
'���ù����嵥��˵����
'�������ޡ�
''*****************************************************************************
    Dim rsYbxm As New ADODB.Recordset
    'WriteBusinessLOG "������ ҽ����Ϣ_ͭɽ��", lngItemID, intType
    If intType = 0 Then 'ҽ����������ʾ
        gstrSQL = "Select * from ypzlk where �շ�ϸĿID=" & lngItemID
        Call OpenRecordset_OtherBase(rsYbxm, "ypzlk", , gcnͭɽ��)
        If rsYbxm.EOF = False Then
            ҽ����Ϣ_ͭɽ�� = rsYbxm!֧�����
        Else
            ҽ����Ϣ_ͭɽ�� = "δ����"
        End If
        Set rsYbxm = Nothing
        If ҽ����Ϣ_ͭɽ�� <> "" Then MsgBox "����Ŀ��ҽ�����Ϊ��" & ҽ����Ϣ_ͭɽ�� & "��", vbInformation, gstrSysName
        ҽ����Ϣ_ͭɽ�� = ""
    End If
End Function



