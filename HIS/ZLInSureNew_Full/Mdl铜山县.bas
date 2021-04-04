Attribute VB_Name = "Mdl铜山县"
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


'变量命名规范:全局变量以g打头,模块级变量以m打头
'定义一个模块级变量，用来记录本接口是否已正常初始化，避免被重复多次初始化
Private mblnInit As Boolean
Private mblnICinit As Boolean
'申请一个全局连接对象
Public gcn铜山县 As New ADODB.Connection
'API函数定义示范
Private Const gstrSysName = "中联软件"
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
'可搜索"TODO:增加自已的实现代码"，找到代码插入点，这些插入点都是必须填写代码的
'TODO:声明部分
'可以在此增加代码，实现XX功能
'-------------------------------------------------------------------------------
'编程步骤说明
'1、为本接口部件命名，规则：zl9I_xxx，如北京医保部件，命名为：zl9I_BJYB，注意，类模块需要命名为：clsI_xxx
'2、如果需要单独保存医保相关的数据，请新建一个用户来处理，我们称之为中间库
'3、与医保相关的参数设置（含中间库的用户名、密码与主机串），请增加保险参数设置窗体，命名规则：frmSet医保名称，如：frmSet北京市
'4、如果中心提供的有医保项目清单、病种目录等，请在保险项目选择的项目更新按钮中填写代码，完成从文件或中心将相关下发数据更新到HIS库中
'5、编写代码完成医保项目对码的功能
'6、编写代码完成身份验证窗体
'7、填入以下函数或过程的主体代码，完成医保接口的主体功能
'8、根据接口性质，修改类模块中GetCapability()方法，相关参数请参见mdlInsure中的枚举变量"医院业务"
'9、根据需要修改类模块中其他方法的调用代码
'10、根据需要增加或修改公共窗体或模块
'-------------------------------------------------------------------------------

Private Type 常用信息_铜山
    Com口           As Long
    病人ID          As Long
    医保卡序号      As String
    个人编号        As String
    医院码          As String
    操作员号        As String
    操作员密码      As String
    病种编码        As String
    急诊否          As String * 1
End Type

Public g常用信息_铜山 As 常用信息_铜山
Dim mstr使用个人帐户支付 As String

'>>调试用
'Public clsTst As New clsLesybjk
'>>

Public Function 医保初始化_铜山县(Optional ByVal blnTest As Boolean = False) As Boolean
'*****************************************************************************
'调用者　　　　　　：被clsInsure 的  InitInsure  过程调用(该方法由费用部件、入出院管理等与医保接口相关的部件调用)
'功能说明　　　　　：完成医保接口初始化相关的工作（全局变量的初始化，环境变量的初始化等）
'调用过程清单及说明：
'    【tsx_conn_ybzx】连接医保中心
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string型参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_getstringparam】取string型接口返回值
'　　【tsx_destroyparams】销毁已分配空间
'
''*****************************************************************************
    'TODO:医保初始
    '以下是参考代码
    Dim strUser As String, strServer As String, strPass As String
    Dim rsTemp As New ADODB.Recordset, lngReturn As Long, rsCsh As New ADODB.Recordset
    Dim strReturn As String, intCOM As Long, STRERR As String
    On Error GoTo errHand

    If mblnInit = False Then
        '读出连接医保服务器的配置
        strUser = "tsxyb"
        strServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
        strPass = "tsxyb"
        g常用信息_铜山.Com口 = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", 0) + 1
'
    intCOM = g常用信息_铜山.Com口 - 1
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
    
    Call WriteBusinessLOG("tsx_init_ic", g常用信息_铜山.Com口, lngReturn)
    If lngReturn <> -1 Then
        mblnICinit = True
    End If
''
        '取医院编码
        gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_铜山县)
        g常用信息_铜山.医院码 = Nvl(rsTemp!医院编码)
        
        If OraDataOpen(gcn铜山县, strServer, strUser, strPass, False) = False Then
            MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
            Exit Function
        Else
            gstrSQL = "Select * from czry Where 人员ID=" & UserInfo.ID
            Call OpenRecordset_OtherBase(rsTemp, "铜山县医保", gstrSQL, gcn铜山县)
            
            If rsTemp.EOF Then
                gstrSQL = "Select * from czry Where P_GLY=1"
                Call OpenRecordset_OtherBase(rsCsh, "铜山县医保", gstrSQL, gcn铜山县)
                
                If rsCsh.EOF Then
                    MsgBox "需要管理员帐户！", vbInformation, gstrSysName
                    Exit Function
                End If
                g常用信息_铜山.操作员号 = rsCsh!P_RYH
                g常用信息_铜山.操作员密码 = rsCsh!P_MM
                
                lngReturn = tsx_conn_ybzx(g常用信息_铜山.医院码, g常用信息_铜山.操作员号, g常用信息_铜山.操作员密码, "")
                Call WriteBusinessLOG("tsx_conn_ybzx", g常用信息_铜山.医院码 & "," & g常用信息_铜山.操作员号 & "," & g常用信息_铜山.操作员密码, lngReturn)
                
                If lngReturn = -1 Then
                    MsgBox "医保初始失败！" & vbCrLf & tsx_getlasterr(), vbInformation, gstrSysName
                    Exit Function
                End If

                '> Beging 中间库中无此操作员信息,调用接口增加人员
                '1分配内存
               If tsx_createparams(1024, 1024) = -1 Then
                    MsgBox "分配内存空间失败!", vbInformation, gstrSysName
                    Exit Function
               End If
               Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
               '2 为参数赋值
               lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
               Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", 0," & g常用信息_铜山.医院码, lngReturn)
               lngReturn = tsx_setstringparam(P_RYH, 0, "")
               Call WriteBusinessLOG("tsx_setstringparam", "P_RYH" & ", 0,''", lngReturn)
               lngReturn = tsx_setstringparam(P_XM, 0, UserInfo.姓名) '    C12 姓名    住院结算时生成的唯一号
               Call WriteBusinessLOG("tsx_setstringparam", "P_XM" & ", 0," & UserInfo.姓名, lngReturn)
               lngReturn = tsx_setstringparam(P_MM, 0, UserInfo.简码 & "001") '    C10 密码    人员注册到医保中心的操作密码
               Call WriteBusinessLOG("tsx_setstringparam", "P_MM" & ", 0," & UserInfo.简码, lngReturn)
               lngReturn = tsx_setstringparam(P_GLY, 0, "0") '   C1  管理员  0-不是1-是
               Call WriteBusinessLOG("tsx_setstringparam", "P_GLY" & ", 0,'0'", lngReturn)
               lngReturn = tsx_setstringparam(P_LB, 0, 1) ' C1  操作类别    0-注销人员(不能恢复)1-增加人员 2-修改密码或姓名
               Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", 0,'1'", lngReturn)
               
               '3 调用接口
               If tsx_jkcall("CZRYWH") = -1 Then
                    Call WriteBusinessLOG("jkcall", "CZRYWH", -1)
                    STRERR = tsx_getlasterr()
                    MsgBox "添加人员失败！" & vbCrLf & STRERR, vbInformation, gstrSysName
                    Call WriteBusinessLOG("tsx_getlasterr", "", STRERR)
                    lngReturn = tsx_destroyparams()
                    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
                    Exit Function
               Else
               Call WriteBusinessLOG("jkcall", "CZRYWH", lngReturn)
               '4 取返回值
                    strPass = Space(10)
                    lngReturn = tsx_getstringparam(P_RYH, 0, strPass)
                    Call WriteBusinessLOG("tsx_getstringparam", "P_RYH, 0, " & g常用信息_铜山.操作员号, lngReturn)
                    
               End If
               '5 销毁已分配空间
               lngReturn = tsx_destroyparams()
               Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
               
               gstrSQL = "insert into czry(人员ID,P_JGM,P_RYH,P_XM,P_MM) values(" & _
                                    UserInfo.ID & ",'" & g常用信息_铜山.医院码 & "','" & _
                                    strPass & "','" & UserInfo.姓名 & "','" & _
                                    UserInfo.简码 & "001')"
               gcn铜山县.Execute gstrSQL
               
               g常用信息_铜山.操作员号 = strPass
               g常用信息_铜山.操作员密码 = UserInfo.简码 & "001"
               '> End 中间库中无此操作员信息,调用接口增加人员
            Else
                g常用信息_铜山.操作员号 = rsTemp!P_RYH
                g常用信息_铜山.操作员密码 = rsTemp!P_MM
                
                lngReturn = tsx_conn_ybzx(g常用信息_铜山.医院码, g常用信息_铜山.操作员号, g常用信息_铜山.操作员密码, "")
                Call WriteBusinessLOG("tsx_conn_ybzx", g常用信息_铜山.医院码 & "," & g常用信息_铜山.操作员号 & "," & g常用信息_铜山.操作员密码, lngReturn)
                
                If lngReturn = -1 Then
                    MsgBox "医保初始失败！" & vbCrLf & tsx_getlasterr(), vbInformation, gstrSysName
                    Exit Function
                End If
                
            End If
            

        End If


    End If

    Set rsTemp = Nothing
    Set rsCsh = Nothing
    mblnInit = True
    医保初始化_铜山县 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保终止_铜山县() As Boolean

'*****************************************************************************
'调用者　　　　　　：被clsInsure 的  EndInsure  过程调用(该方法由费用部件、入出院管理等与医保接口相关的部件调用)
'功能说明　　　　　：对象的释放，断开连接等
'调用过程清单及说明：
'　　【tsx_disconn_ybzx】医保部件关闭
''*****************************************************************************
    'TODO:医保终止
    '以下是参考代码
    'Dim strReturn As String
    'Call 调用接口(停止政策机服务, strReturn)
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
    
    医保终止_铜山县 = True
    Exit Function
errHand:
    Call WriteBusinessLOG("ErrHand", Err.Number, Err.Description)
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function 身份标识_铜山县(ByVal bytType As Byte, Optional lng病人ID As Long = 0, Optional ByRef intinsure As Integer) As String
'*************************************************
'调用者　　　　　　：被clsInsure 的 Identify  过程调用(该方法由门诊费用部件、门诊挂号部件或入院登记部件调用)
'功能说明　　　　　：识别指定人员是否为参保病人，身份验证成功后，将病人信息串返回给主调程序
'参数　　　　　　　：bytType-识别类型，0-门诊，1-住院
'返回　　　　　　　：空或信息串
'注意　　　　　　　：1)主要利用接口的身份识别交易；
'　　　　　　　　　　2)如果识别错误，在此函数内直接提示错误信息；
'　　　　　　　　　　3)识别正确，而个人信息缺少某项，必须以空格填充；
'调用过程清单及说明：
'　　【无】
'*************************************************
'TODO: 身份验证
   ' Dim rsSfyz As New ADODB.Recordset
    Dim strReturn As String, lngReturn As Long, str挂号单号 As String
    
    strReturn = frmIdentify铜山县.GetIdentify(bytType, lng病人ID)
'        If bytType = 0 Then
'           '>Beging 门诊就诊登记功能
'           '1
'            If tsx_createparams(1024, 1024) = -1 Then
'                MsgBox "分配内存空间失败!", vbInformation, gstrSysName
'                Exit Function
'            End If
'            Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
'            '2 为参数赋值
'            lngReturn= tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
'            Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", 0," & g常用信息_铜山.医院码, lngReturn)
'
'            lngReturn= tsx_setstringparam(P_KXH, 0, g常用信息_铜山.医保卡序号) '   C3  医保卡序号
'            Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", 0," & g常用信息_铜山.医保卡序号, lngReturn)
'            lngReturn= tsx_setstringparam(P_TBR, 0, g常用信息_铜山.个人编号) '   C9  个人编号
'            Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", 0," & g常用信息_铜山.个人编号, lngReturn)
'            lngReturn= tsx_setstringparam(P_YYKSM, 0, "001") ' C20 挂号科室码
'            Call WriteBusinessLOG("tsx_setstringparam", "P_YYKSM" & ", 0,'001'", lngReturn)
'            lngReturn= tsx_setdoubleparam(P_GHF, 0, 0)   '   D   挂号费
'            Call WriteBusinessLOG("tsx_setdoubleparam", "P_GHF" & ", 0,0", lngReturn)
'
'            lngReturn= tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) '  C10 操作人员编号
'            Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH" & ", 0,'" & g常用信息_铜山.操作员号 & "'", lngReturn)
'            '3 调用接口
'            If jkcall("MZGH") = -1 Then
'                 Call WriteBusinessLOG("jkcall", "MZGH", lngReturn)
'                 MsgBox tsx_getlasterr(), vbInformation, gstrSysName

'                 lngReturn= tsx_destroyparams()
'                 Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
'                 Exit Function
'            Else
'            Call WriteBusinessLOG("jkcall", "MZGH", lngReturn)
'            '4 取返回值
'                 lngReturn= tsx_getstringparam(P_DJH, 0, str挂号单号)
'                 Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str挂号单号, lngReturn)
'
'            End If
'            '5 销毁已分配空间
'            lngReturn= tsx_destroyparams()
'            Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
'
'            '保存str挂号单号到保险帐户中
'
'
'           '>End 门诊就诊登记功能
'        End If
    身份标识_铜山县 = strReturn
    
End Function

Public Function 医保设置_铜山县(ByVal intinsure As Integer) As Boolean
'**************************************
'调用者　　　　　　：被clsInsure的CodeMan的1600功能 调用
'功能说明　　　　　：医保参数设置
'调用过程清单及说明：
'　　【　　　】
'**************************************

    '医保设置_北京 = frmSet北京.参数设置()
End Function

Public Function 门诊挂号_铜山县(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
'******************************************************************************
'调用者　　　　　　：该方法由门诊挂号部件调用
'功能说明　　　　　：通过调用医保商的门诊挂号接口，分解本次费用明细，得到结算结果（个人帐户多少、统筹基金多少等）并保存
'注意事项　　　　　：需要调用过程zl_病人结算记录_Update对病人预交记录进行数据修正
'调用过程清单及说明：
'　　【　　　】
''*****************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim str科室码 As String, dbl挂号费  As Double, str交易流水号 As String
    Dim lngReturn As Long, lngCounter As Long, dbl个人自费 As Double, STRERR As String, str流水号 As String
    On Error GoTo ErrH
    strSQL = "Select b.编码 as 科室码,Sum(A.实收金额) as 实收金额 " & _
              " From 门诊费用记录 A,部门表 B" & _
              " Where a.病人科室Id=b.id and A.结帐ID=[1] And Nvl(A.附加标志,0)<>9 And Nvl(A.记录状态,0)<>0 group by b.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "取挂号费", lng结帐ID)
    str科室码 = "": dbl挂号费 = 0
    Do Until rsTmp.EOF
        str科室码 = Trim("" & rsTmp!科室码)
        dbl挂号费 = Val("" & rsTmp!实收金额)
        rsTmp.MoveNext
    Loop
    
    If tsx_createparams(102400, 102400) = -1 Then
        Err.Raise 9000, gstrSysName, "分配内存空间失败!"
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    
    'P_JGM   C5  医院码
    lngReturn = tsx_setstringparam(P_JGM, lngCounter, g常用信息_铜山.医院码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCounter & "," & g常用信息_铜山.医院码, lngReturn)
    'P_KXH   C3  医保卡序号
    lngReturn = tsx_setstringparam(P_KXH, lngCounter, g常用信息_铜山.医保卡序号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & lngCounter & "," & g常用信息_铜山.医保卡序号, lngReturn)
    'P_TBR   C9  个人编号
    lngReturn = tsx_setstringparam(P_TBR, lngCounter, g常用信息_铜山.个人编号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & lngCounter & "," & g常用信息_铜山.个人编号, lngReturn)
    'P_YYKSM C20 挂号科室码
    lngReturn = tsx_setstringparam(P_YYKSM, 0, str科室码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_YYKSM,0," & str科室码, lngReturn)
    'P_GHF   D   挂号费
     lngReturn = tsx_getdoubleparam(P_GHF, 0, dbl挂号费)
     Call WriteBusinessLOG("tsx_getdoubleparam", "P_GHF, 0, " & dbl个人自费, lngReturn)
    'P_CZRYH C10 操作人员编号
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

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
         '4 取返回值
         lngReturn = tsx_getstringparam(P_DJH, 0, str交易流水号)
         Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str流水号, lngReturn)
    End If

    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("计算tsx_destroyparams", "", lngReturn)
    门诊挂号_铜山县 = True
    '保存结算记录
    strSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_铜山县 & "," & g常用信息_铜山.病人ID & "," & _
        Year(zlDatabase.Currentdate) & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & 0 & ",0,0,0," & dbl挂号费 & ",0,0," & _
        dbl挂号费 & "," & 0 & ",0," & 0 & "," & 0 & ",'" & str交易流水号 & "',NULL,NULL,'" & "" & "')"
    Call WriteBusinessLOG("门诊挂号", "保存保险结算记录", gstrSQL)
    Call zlDatabase.ExecuteProcedure(strSQL, "铜山县医保")
    
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊挂号冲销_铜山县(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
'******************************************************************************
'调用者　　　　　　：该方法由门诊挂号部件调用
'功能说明　　　　　：通过调用医保商的门诊挂号冲销接口，完成门诊挂号结算的作废
'调用过程清单及说明：
'　　【　　　】
''*****************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng冲销ID As Long
    Dim str科室码 As String, dbl挂号费  As Double, str交易流水号 As String
    Dim str冲销流水号 As String, lngReturn As Long, lngCounter As Long, STRERR As String
    On Error GoTo ErrH
    strSQL = " select distinct A.结帐ID,A.NO,B.病人科室ID,to_char(B.发生时间,'YYYYMMDD') as 发生日期 " & _
             " from 门诊费用记录 A,门诊费用记录 B " & _
             " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人结算记录", lng结帐ID)
    lng冲销ID = Val("" & rsTmp!结帐ID)
    
    strSQL = "Select B.人员身份,A.病人ID,B.卡号,B.帐户标志,A.支付顺序号,A.发生费用金额,B.备注 " & _
              " from 保险帐户 B,保险结算记录 A " & _
              " Where A.险类=[2] And B.险类 = [2]" & _
              " And B.病人ID = A.病人ID And A.记录ID=[1]"
    Call WriteBusinessLOG("门诊挂号冲销", "提被冲销记录", strSQL)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取冲销记录", lng结帐ID, intinsure)
    If rsTmp.EOF Then
        Err.Raise 9000, gstrSysName, "无结算记录,不能冲销!"
        Exit Function
    End If
    
    Do Until rsTmp.EOF
        str交易流水号 = Trim("" & rsTmp!支付顺序号)
        rsTmp.MoveNext
    Loop
    
    If tsx_createparams(102400, 102400) = -1 Then
        Err.Raise 9000, gstrSysName, "分配内存空间失败!"
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    
    'P_JGM   C5  医院码
    lngReturn = tsx_setstringparam(P_JGM, lngCounter, g常用信息_铜山.医院码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCounter & "," & g常用信息_铜山.医院码, lngReturn)
    'P_TBR   C9  个人编号
    lngReturn = tsx_setstringparam(P_TBR, lngCounter, g常用信息_铜山.个人编号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & lngCounter & "," & g常用信息_铜山.个人编号, lngReturn)
    'P_KXH   C3  医保卡序号
    lngReturn = tsx_setstringparam(P_KXH, lngCounter, g常用信息_铜山.医保卡序号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & lngCounter & "," & g常用信息_铜山.医保卡序号, lngReturn)
    'P_DJH   C20 原挂号单据号
    lngReturn = tsx_setstringparam(P_DJH, 0, str交易流水号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_DJH,0," & str交易流水号, lngReturn)
    'P_CZRYH C10 操作人员编号
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

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
         '4 取返回值
         lngReturn = tsx_getstringparam(P_DJH, 0, str冲销流水号)
         Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str冲销流水号, lngReturn)
    End If

    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    
    strSQL = "Select * from 保险结算记录 Where 记录ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "保险结算", lng结帐ID)
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_铜山县 & "," & rsTmp!病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTmp!发生费用金额, 0) & "," & -1 * Nvl(rsTmp!全自付金额, 0) & _
        "," & -1 * Nvl(rsTmp!首先自付金额, 0) & "," & -1 * Nvl(rsTmp!进入统筹金额, 0) & _
        "," & -1 * Nvl(rsTmp!统筹报销金额, 0) & _
        "," & -1 * Nvl(rsTmp!大病自付金额, 0) & "," & 0 & "," & -1 * Nvl(rsTmp!个人帐户支付) & _
        ",'" & str冲销流水号 & "',null,null,'" & Nvl(rsTmp!备注) & "')"
    Call WriteBusinessLOG("门诊挂号冲销", "保险结算记录", gstrSQL)
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊虚拟结算_铜山县(rsHis As ADODB.Recordset, str结算方式 As String, ByVal intinsure As Integer, Optional lng结帐ID As Long = 0) As Boolean
'*******************************************
'调用者　　　　　　：被clsInsure 的 ClinicPreSwap 过程调用(该方法由门诊费用部件调用)
'功能说明　　　　　：通过调用医保商的预结算方法，分解本次费用明细，得到结算结果（个人帐户
'　　　　　　　　　　多少、统筹基金多少等），并将结算结果按格式保存在参数“str结算方式”中
'步骤说明　　　　　：
'                   1、如果接口需要，请调用费用明细上传接口，将本次明细上传
'                   2、调用门诊预结算接口
'                   3、将结算结果按规定格式返回
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string型参数
'　　【tsx_setdoubleparam】设置double型参数
'　　【tsx_setlongparam】设置long型参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_getstringparam】取string型接口返回值
'　　【tsx_getdoubleparam】取double型接口返回值
'　　【tsx_destroyparams】销毁已分配空间
'*******************************************
    'TODO:门诊虚拟结算
    'rs明细记录集中是本次录入的门诊处方明细
    'str结算方式的格式说明：报销方式;金额;是否允许修改|....
    Dim dbl统筹基金 As Double, dbl大病支付 As Double
    Dim lngReturn As Long, lngCounter As Long
    Dim rsMzxnjs As New ADODB.Recordset, str项目类别 As String
    Dim blnErr As Boolean '记录上传时是否有错误
    Dim str门诊交易流水号 As String, dbl总费用 As Double, dbl个人自费 As Double
    Dim dbl个人帐户支付 As Double, dbl统筹基金支付 As Double, dbl大病统筹支付 As Double
    Dim dbl公务员基金支付 As Double, dbl期末个人帐户 As Double, dbl期初个人帐户 As Double
    Dim dbl个人自付 As Double, dbl单价 As Double, str医生号 As String, str科室编码 As String
    Dim STRERR As String
    Dim str个人编号 As String, str医保卡序号 As String, str就诊类型 As String
    Dim rs明细 As New ADODB.Recordset
    On Error GoTo errHandle
    
    '>>Beging 结算前要求刷卡
'    If lng结帐ID > 0 Then
'        lngReturn= tsx_init_ic(g常用信息_铜山.Com口)
'        Call WriteBusinessLOG("tsx_init_ic", g常用信息_铜山.Com口, lngReturn)
'        lngReturn= tsx_read_ic(str个人编号, str医保卡序号)
'        Call WriteBusinessLOG("tsx_read_ic", str个人编号 & "," & str医保卡序号, lngReturn)
'
'        If str个人编号 <> g常用信息_铜山.个人编号 Then
'            MsgBox "结算时的医保卡和身份验证时的医保卡号不符，不能结算！", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    '>>End 结算前要求刷卡
    gstrSQL = "Select 就诊类别 From 保险帐户 Where 险类=[1] ANd 医保号=[2]"
    Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "取就诊类别", intinsure, g常用信息_铜山.个人编号)
    str就诊类型 = rsMzxnjs!就诊类别
    
    rsHis.Filter = "实收金额 <> 0"
    Set rs明细 = rsHis
    
    gstrSQL = "Select * from 部门表 where ID=[1]"
    Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "部门表", CLng(rs明细!开单部门ID))
    str科室编码 = rsMzxnjs!编码
    gstrSQL = "Select 编号 from 人员表 where 姓名=[1]"
    Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "部门表", CStr(rs明细!开单人))
    str医生号 = rsMzxnjs!编号


    '>>Beging 上传明细========================================================================================
    '1
    'If lng结帐ID = 0 Then
        If tsx_createparams(102400, 102400) = -1 Then
            MsgBox "分配内存空间失败!", vbInformation, gstrSysName
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
   
        lngCounter = 0
        
        Do Until rs明细.EOF
        
            '2 为参数赋值
            lngReturn = tsx_setstringparam(P_JGM, lngCounter, g常用信息_铜山.医院码)
            Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCounter & "," & g常用信息_铜山.医院码, lngReturn)
            
            lngReturn = tsx_setstringparam(P_TBR, lngCounter, g常用信息_铜山.个人编号) '    C9  个人编号    参保人员个人编号
            Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & lngCounter & "," & g常用信息_铜山.个人编号, lngReturn)
            
            lngReturn = tsx_setstringparam(P_KXH, lngCounter, g常用信息_铜山.医保卡序号) '   C3  医保卡序号  参保人员IC卡序号
            Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & lngCounter & "," & g常用信息_铜山.医保卡序号, lngReturn)
            
            gstrSQL = "Select * from 收费细目 where ID=[1]"
            Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "取收费类别", CLng(rs明细!收费细目ID))
            
            str项目类别 = 1
            If InStr("567", rsMzxnjs!类别) > 0 Then
               str项目类别 = 0
            End If
            
            lngReturn = tsx_setstringparam(P_LB, lngCounter, str项目类别) '    C1  类别    0-药品,1-诊疗项目
            Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & lngCounter & "," & str项目类别, lngReturn)
            
            lngReturn = tsx_setstringparam(P_JZLB, lngCounter, "0") ' C1  就诊类别    暂固定为'0'
            Call WriteBusinessLOG("tsx_setstringparam", "P_JZLB" & ", " & lngCounter & ",'0'", lngReturn)
            
            lngReturn = tsx_setstringparam(P_YBBM, lngCounter, "") '  C20 医保编码    暂为空串""
            Call WriteBusinessLOG("tsx_setstringparam", "P_YBBM" & ", " & lngCounter & ",''", lngReturn)
            
            gstrSQL = "Select * from ypzlk where 收费细目ID=" & rs明细!收费细目ID
            
            Call OpenRecordset_OtherBase(rsMzxnjs, "ypzlk", , gcn铜山县)
            If rsMzxnjs.EOF = False Then
               lngReturn = tsx_setstringparam(P_ZBM, lngCounter, rsMzxnjs!自编码) '   C20 自编码  药品(或诊疗项目)自编码
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCounter & "," & rsMzxnjs!自编码, lngReturn)
            Else
                lngReturn = tsx_setstringparam(P_ZBM, lngCounter, rs明细!收费细目ID) '   C20 自编码  药品(或诊疗项目)自编码
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCounter & "," & rs明细!收费细目ID, lngReturn)
            End If
            
            'Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCounter & "," & rs明细!收费细目ID, lngReturn)
            dbl单价 = Format(rs明细!实收金额 / IIf(Format(rs明细!数量, "0") = 0, 1, Format(rs明细!数量, "0")), "0.0000")
            lngReturn = tsx_setdoubleparam(P_JG, lngCounter, dbl单价)    '    D   单价
            Call WriteBusinessLOG("tsx_setdoubleparam", "P_JG" & ", " & lngCounter & "," & rs明细!实收金额 / rs明细!数量, lngReturn)
            
            lngReturn = tsx_setlongparam(P_SL, lngCounter, IIf(Format(rs明细!数量, "0") = 0, 1, Format(rs明细!数量, "0")))  '    L   数量
            Call WriteBusinessLOG("tsx_setlongparam", "P_SL" & ", " & lngCounter & "," & rs明细!数量, lngReturn)
            
            
            lngCounter = lngCounter + 1
            rs明细.MoveNext
        Loop
    'End If
            '3 调用接口
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
            '4 取返回值
                 '无返回值
            End If
            '5 销毁已分配空间
            lngReturn = tsx_destroyparams()
            Call WriteBusinessLOG("明细上传tsx_destroyparams", "", lngReturn)
    
    '>>End 上传明细========================================================================================
'
    
    '>>Beging 费用计算========================================================================================
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "分配内存空间失败!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    '2 为参数赋值
    lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, g常用信息_铜山.个人编号) '    C9  个人编号    参保人员个人编号
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & g常用信息_铜山.个人编号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, g常用信息_铜山.医保卡序号) '   C3  医保卡序号  参保人员IC卡序号
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & g常用信息_铜山.医保卡序号, lngReturn)

    lngReturn = tsx_setstringparam(P_TSMZ, 0, Nvl(str就诊类型, "0")) '  C1  特殊就诊类别
    Call WriteBusinessLOG("tsx_setstringparam", "P_TSMZ,0,'0'", lngReturn)
    
    lngReturn = tsx_setstringparam(P_YYKSM, 0, str科室编码) 'C20 医院科室编码
    Call WriteBusinessLOG("tsx_setstringparam", "P_YYKSM,0," & str科室编码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_YSRYH, 0, str医生号) 'C10 医生人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_YSRYH,0," & str医生号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CFH, 0, "") ' C20 处方号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CFH,0,''", lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZM, 0, g常用信息_铜山.病种编码) '    C10 病种码
    Call WriteBusinessLOG("tsx_setstringparam", "P_BZM,0," & g常用信息_铜山.病种编码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_JZ, 0, g常用信息_铜山.急诊否)  'C1  是否急诊
    Call WriteBusinessLOG("tsx_setstringparam", "P_JZ,0," & g常用信息_铜山.急诊否, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)
               
    '3 调用接口
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
         '4 取返回值
         lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl总费用) 'D 总费用
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl总费用, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl个人自费) '  D   个人自费
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl个人自费, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl个人自付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl个人自付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl个人帐户支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl个人帐户支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dbl统筹基金支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dbl统筹基金支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl大病统筹支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl大病统筹支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl公务员基金支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl公务员基金支付, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl期初个人帐户)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl期初个人帐户, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl期末个人帐户)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl期初个人帐户, lngReturn)
         
    End If
    

    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("计算tsx_destroyparams", "", lngReturn)

    '>>End 费用计算========================================================================================
    
    
    '>>Beging 结算确认~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If lng结帐ID > 0 Then ''结帐ID>0
        If tsx_createparams(1024, 1024) = -1 Then
            MsgBox "分配内存空间失败!", vbInformation, gstrSysName
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
        '2 为参数赋值
        lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
        Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
        
        lngReturn = tsx_setstringparam(P_TBR, 0, g常用信息_铜山.个人编号) '    C9  个人编号    参保人员个人编号
        Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & g常用信息_铜山.个人编号, lngReturn)
        
        lngReturn = tsx_setstringparam(P_KXH, 0, g常用信息_铜山.医保卡序号) '   C3  医保卡序号  参保人员IC卡序号
        Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & g常用信息_铜山.医保卡序号, lngReturn)
    
        lngReturn = tsx_setdoubleparam(P_QCGRZH, 0, Val(Format(dbl期初个人帐户, "0.00")))
        Call WriteBusinessLOG("tsx_setdoubleparam", "P_QCGRZH, 0, " & dbl期初个人帐户, lngReturn)
    
        lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
        Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)
                   
        '3 调用接口
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
             '4 取返回值
             str门诊交易流水号 = Space(20)
             lngReturn = tsx_getstringparam(P_DJH, 0, str门诊交易流水号)  '   C20 单据号
             Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str门诊交易流水号, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl总费用) 'D 总费用
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl总费用, lngReturn)
            
             lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl个人自费) '  D   个人自费
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl个人自费, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl个人自付)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl个人自付, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl个人帐户支付)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl个人帐户支付, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dbl统筹基金支付)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dbl统筹基金支付, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl大病统筹支付)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl大病统筹支付, lngReturn)
             
             lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl公务员基金支付)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl公务员基金支付, lngReturn)
            
             lngReturn = tsx_getdoubleparam(P_QMGRZH, 0, dbl期末个人帐户)
             Call WriteBusinessLOG("tsx_getdoubleparam", "P_QMGRZH, 0, " & dbl期末个人帐户, lngReturn)
        End If
        
        '5 销毁已分配空间
        lngReturn = tsx_destroyparams()
        Call WriteBusinessLOG("结算完成tsx_destroyparams", "", lngReturn)

        '**保存保险结算记录**
        If InStr(str门诊交易流水号, Chr(0)) > 0 Then
            str门诊交易流水号 = Mid(str门诊交易流水号, 1, InStr(str门诊交易流水号, Chr(0)) - 1)
        End If
        gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & intinsure & "," & g常用信息_铜山.病人ID & "," & _
            Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
            0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
            dbl总费用 & "," & dbl个人自费 & "," & dbl个人自付 & "," & dbl总费用 - dbl个人自费 - dbl个人自付 & "," & _
            dbl统筹基金支付 + dbl公务员基金支付 & "," & dbl大病统筹支付 & "," & _
            0 & "," & dbl个人帐户支付 & ",'" & str门诊交易流水号 & "',Null,Null,Null)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        
        gstrSQL = "zl_保险帐户_更新信息(" & g常用信息_铜山.病人ID & "," & TYPE_铜山县 & ",'帐户余额','''" & dbl期末个人帐户 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保险帐户")
    End If '结帐ID>0
    '>>End 结算确认~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '以下是组织结算方式串的示范代码
    str结算方式 = "统筹基金;" & dbl统筹基金支付 & ";0"
    str结算方式 = str结算方式 & "|大病支付;" & dbl大病统筹支付 & ";0"
    str结算方式 = str结算方式 & "|公务员基金;" & dbl公务员基金支付 & ";0"
    str结算方式 = str结算方式 & "|个人帐户;" & dbl个人帐户支付 & ";0"
    
    门诊虚拟结算_铜山县 = True
    
    Set rsMzxnjs = Nothing
    Call WriteBusinessLOG("返回前台", "", str结算方式)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_铜山县(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String, ByVal intinsure As Integer) As Boolean
'*****************************************************************
'调用者　　　　　　：被clsInsure 的 ClinicSwap 过程调用(该方法由门诊费用部件调用)
'功能说明　　　　　：调用门诊结算接口
'步骤说明　　　　　：
'　　　　　　　　　　1、如果需要上传明细，则调处方明细上传接口
'　　　　　　　　　　2、调用门诊结算接口
'　　　　　　　　　　3、如果成功，则保存保险结算记录
'调用过程清单及说明：
'　　【门诊虚拟结算_铜山县】 完成明细上传,虚拟结算，结算功能
'*****************************************************************、
'TODO:门诊结算
    Dim str结算方式 As String
    Dim rsMzjs As New ADODB.Recordset
On Error GoTo errHandle
    
    gstrSQL = "Select ID,NO,序号,记录性质,登记时间 as 结算时间,病人ID,收费类别,收据费目,计算单位,开单人,开单部门ID, " & _
                     "收费细目ID,nvl(数次,0)*nvl(付数,0) as 数量,标准单价 as 单价, " & _
                     "实收金额,统筹金额,保险大类ID 保险支付大类ID, " & _
                     " 摘要,是否急诊 " & _
            "from 门诊费用记录 " & _
            "where 结帐ID=[1]"
    Set rsMzjs = zlDatabase.OpenSQLRecord(gstrSQL, "", lng结帐ID)
    
    '**保存保险结算记录**
'    gstrSQL = "zl_保险结算记录_insert(" & IIf(bln住院, 2, 1) & "," & lng结帐ID & "," & TYPE_北京 & "," & lng病人ID & "," & _
'        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
'        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
'        dbl费用总额 & "," & dbl现金 & ",0,0," & dbl统筹基金 & "," & dbl大病补助 & "," & _
'        0 & ",0,'" & gComInfo.交易流水号 & "',null,null,'" & gComInfo.业务类型 & "')"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
  
    门诊结算_铜山县 = 门诊虚拟结算_铜山县(rsMzjs, str结算方式, intinsure, lng结帐ID)
    Set rsMzjs = Nothing
    
    Call WriteBusinessLOG("从结算返回前台", "", str结算方式)

    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_铜山县(ByVal lng结帐ID As Long, ByVal cur个人帐户 As Currency, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'*************************************************************************
'调用者　　　　　　：被clsInsure 的 ClinicDelSwap 过程调用
'功能说明　　　　　：调用门诊结算作废接口
'步骤说明　　　　　：1、按接口规则判断是否必须从最后一次就诊的门诊单据开始退废
'　　　　　　　　　　2、调用门诊结算作废接口
'　　　　　　　　　　3、保存保险结算记录
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string型参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_getstringparam】取string型接口返回值
'　　【tsx_destroyparams】销毁已分配空间
'*************************************************************************
'TODO:门诊冲销
    Dim rsMzJsCx As New ADODB.Recordset
    Dim lngReturn As Long, str个人编号 As String, str医保卡序号 As String
    Dim str原单据号 As String, lng冲销ID As Long, STRERR As String
On Error GoTo errHand
    
    gstrSQL = "Select * from 保险帐户 where 病人ID=[1] And 险类=[2]"
    Set rsMzJsCx = zlDatabase.OpenSQLRecord(gstrSQL, "", lng病人ID, TYPE_铜山县)
    If rsMzJsCx.EOF Then
        Err.Raise 9000, gstrSysName, "不是铜山县医保参保人员,不能冲销!"
        Exit Function
    End If
    str个人编号 = rsMzJsCx!医保号
    str医保卡序号 = rsMzJsCx!卡号
    
    gstrSQL = "select distinct A.结帐ID,A.NO from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsMzJsCx = zlDatabase.OpenSQLRecord(gstrSQL, "", lng结帐ID)
    lng冲销ID = rsMzJsCx!结帐ID
    
    gstrSQL = "Select * From 保险结算记录 Where 性质=1 And 记录ID=[1] and 险类=[2]"
    Set rsMzJsCx = zlDatabase.OpenSQLRecord(gstrSQL, "取结算流水号", lng结帐ID, TYPE_铜山县)
    If rsMzJsCx.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "没有找到原始结算记录，无法进行门诊结算冲销！"
        Exit Function
    End If
    str原单据号 = rsMzJsCx!支付顺序号
    
    
    '1 分配空间
    If tsx_createparams(1024, 1024) = -1 Then
        Err.Raise 9000, gstrSysName, "分配内存空间失败!"
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)

    '2 为参数赋值
    lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, str个人编号) '    C9  个人编号    参保人员个人编号
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str个人编号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, str医保卡序号) '   C3  医保卡序号  参保人员IC卡序号
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & str医保卡序号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_DJH, 0, str原单据号) '
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & str原单据号, lngReturn)
   
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

    '3 调用接口
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
    '4 取返回值
        lngReturn = tsx_getstringparam(P_DJH, 0, str原单据号)
        Call WriteBusinessLOG("tsx_getstringparam", "P_RYH, 0, " & str原单据号, lngReturn)
         
    End If
    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("冲销tsx_destroyparams", "", lngReturn)


    '保存保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_铜山县 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsMzJsCx!发生费用金额, 0) & "," & -1 * Nvl(rsMzJsCx!全自付金额, 0) & _
        "," & -1 * Nvl(rsMzJsCx!首先自付金额, 0) & "," & -1 * Nvl(rsMzJsCx!进入统筹金额, 0) & _
        "," & -1 * Nvl(rsMzJsCx!统筹报销金额, 0) & _
        "," & -1 * Nvl(rsMzJsCx!大病自付金额, 0) & "," & 0 & "," & -1 * Nvl(rsMzJsCx!个人帐户支付) & _
        ",'" & str原单据号 & "',null,null,'" & Nvl(rsMzJsCx!备注) & "')"
        
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    
    门诊结算冲销_铜山县 = True
    '
    Set rsMzJsCx = Nothing
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 入院登记_铜山县(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
'**************************************************************************************************
'调用者　　　　　　：被clsInsure 的 ComeInSwap 过程调用(由病人入院部件调用)
'功能说明　　　　　：调用入院登记接口
'步骤说明　　　　　：1、从病案主页中提取入院日期（补充入院登记也是调用该接口，因此不能取当前日期做为入院日期上传）
'　　　　　　　　　　2、调用入院登记接口
'　　　　　　　　　　3、执行入院登记过程(zl_保险帐户_入院)，更改病人的当前状态
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string参数
'　　【tsx_setdoubleparam】设置double参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_getstringparam】取string型接口返回值
'　　【tsx_destroyparams】销毁已分配空间
'***************************************************************************************************
    'TODO:入院登记
    Dim lngReturn As Long, str入院日期 As String, str入院病区 As String
    Dim str入院床位号 As String, str住院号 As String, str联系电话 As String
    Dim dbl入院押金 As Double, str住院医生 As String, str门诊医生 As String
    Dim str住院流水号 As String, str科室编码 As String, STRERR As String
    Dim rsRydj As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "Select * from 保险帐户 where 病人ID=[1] and 险类=[2] and nvl(备注,'0')<>'0'"
                  
    Set rsRydj = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", lng病人ID, TYPE_铜山县)
    If Not rsRydj.EOF Then
        MsgBox "该参保人员的欠款重结交易还未完成，不能办理入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select A.*,B.编码 as 病区 from 病案主页 A ,部门表 B" & _
               " where A.病人ID=[1] And A.主页ID=[2] And A.险类=[3] And A.入院科室ID=B.ID"
               
    Set rsRydj = zlDatabase.OpenSQLRecord(gstrSQL, "病案主页", lng病人ID, lng主页ID, TYPE_铜山县)
    If rsRydj.EOF Then
        MsgBox "不是医保病人不能办理医保入院!", vbInformation, gstrSysName
        Exit Function
    End If
    str入院日期 = Format(rsRydj!入院日期, "yyyyMMdd")
    
    str入院病区 = rsRydj!入院科室ID: str入院床位号 = Nvl(rsRydj!入院病床, "")
    str住院号 = lng病人ID & "_" & lng主页ID: str联系电话 = Nvl(rsRydj!联系人电话, "")
    str住院医生 = Nvl(rsRydj!住院医师, ""): str门诊医生 = Nvl(rsRydj!门诊医师, "")
    dbl入院押金 = 0
    
    str科室编码 = rsRydj!病区
    
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "分配内存空间失败!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    '2 为参数赋值
    lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, g常用信息_铜山.个人编号) '    C9  个人编号    参保人员个人编号
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & g常用信息_铜山.个人编号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, g常用信息_铜山.医保卡序号) '   C3  医保卡序号  参保人员IC卡序号
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & g常用信息_铜山.医保卡序号, lngReturn)

    lngReturn = tsx_setstringparam(P_YYKSM, 0, str科室编码) 'C20 医院科室编码
    Call WriteBusinessLOG("tsx_setstringparam", "P_YYKSM,0," & str科室编码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZM, 0, "") '    C10 病种码 暂为空串( g常用信息_铜山.病种编码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_BZM,0," & "", lngReturn)
    
    lngReturn = tsx_setstringparam(P_RYYMD, 0, str入院日期)
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYYMD,0," & str入院日期, lngReturn)

    lngReturn = tsx_setstringparam(P_RYBQ, 0, str入院病区)  '
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ,0," & str入院病区, lngReturn)

    lngReturn = tsx_setstringparam(P_RYCWH, 0, str入院床位号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH,0," & str入院床位号, lngReturn)

    lngReturn = tsx_setstringparam(P_ZYH, 0, str住院号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYH,0," & str住院号, lngReturn)

    lngReturn = tsx_setstringparam(P_LXDH, 0, str联系电话)
    Call WriteBusinessLOG("tsx_setstringparam", "P_LXDH,0," & str联系电话, lngReturn)

    lngReturn = tsx_setdoubleparam(P_YJHJ, 0, dbl入院押金)
    Call WriteBusinessLOG("tsx_setdoubleparam", "P_YJHJ,0," & dbl入院押金, lngReturn)

    lngReturn = tsx_setstringparam(P_YSRYH, 0, str住院医生)
    Call WriteBusinessLOG("tsx_setstringparam", "P_YSRYH,0," & str住院医生, lngReturn)

    lngReturn = tsx_setstringparam(P_KDYSRYH, 0, str门诊医生)
    Call WriteBusinessLOG("tsx_setstringparam", "P_YSRYH,0," & str门诊医生, lngReturn)

    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

    '3 调用接口
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
    '4 取返回值
         str住院流水号 = Space(20)
         lngReturn = tsx_getstringparam(P_ZYXH, 0, str住院流水号)
         Call WriteBusinessLOG("tsx_getstringparam", "P_ZYXH, 0, " & str住院流水号, lngReturn)
         
    End If
    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)


    '改变病人状态
    If lngReturn = 0 Then
        gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_铜山县 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
        str住院流水号 = Mid(str住院流水号, 1, InStr(str住院流水号, Chr(0)) - 1)
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_铜山县 & ",'顺序号','''" & str住院流水号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存流水号")
        
        入院登记_铜山县 = True
    End If
    Set rsRydj = Nothing
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_铜山县(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
'****************************************
'调用者　　　　　　：被clsInsure 的 ComeInDelSwap  过程调用
'功能说明　　　　　：调用撤销入院登记或出院登记接口
'调用过程清单及说明：
'步骤说明　　　　　：1、按接口规则进行检查（一般发生费用或进行结算过的病人，不允许调用该接口）
'　　　　　　　　　　2、调用撤销入院接口
'　　　　　　　　　　3、执行出院登记过程(zl_保险帐户_出院)，更改病人的当前状态
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_destroyparams】销毁已分配空间
'****************************************
'TODO:入院登记撤销
    On Error GoTo errHand
    Dim rsRydjCx As New ADODB.Recordset
    Dim str流水号 As String, STRERR As String
    Dim lngReturn As Long
    Dim str个人编号 As String, str医保卡号 As String, str备注 As String
    '结过帐的不允许撤销
    gstrSQL = "Select sum(nvl(结帐金额,0)) as 结帐金额 from 住院费用记录 where 病人ID=[1] And 主页ID=[2]"
    Set rsRydjCx = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", lng病人ID, lng主页ID)
    
    If rsRydjCx!结帐金额 <> 0 Then
        MsgBox "已结过帐的医保人员不能撤消入院登记!", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select * from 保险帐户 where 险类=[1] And 病人ID=[2]"
    Set rsRydjCx = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", intinsure, lng病人ID)
    str流水号 = rsRydjCx!顺序号
    str个人编号 = rsRydjCx!医保号
    str医保卡号 = rsRydjCx!卡号
    str备注 = InputBox("请填写备注", gstrSysName)
    lngReturn = -1
    '>>>beging 撤消入院
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "分配内存空间失败!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    '2 为参数赋值
    lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, str个人编号) '    C9  个人编号    参保人员个人编号
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str个人编号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, str医保卡号) '   C3  医保卡序号  参保人员IC卡序号
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & str医保卡号, lngReturn)

    lngReturn = tsx_setstringparam(P_ZYXH, 0, str流水号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str流水号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZ, 0, str备注) ' C10 操作人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)
    '3 调用接口
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
    '4 取返回值
       '无返回
    End If
    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

    
    '>>>End 撤消入院
    If lngReturn = 0 Then
        '改变所有已上传记录为未上传
        gstrSQL = "Update 住院费用记录 Set 是否上传=0 where nvl(是否上传,0)=1 and 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
        gcnOracle.Execute gstrSQL
    '    '改变病人状态
        gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & intinsure & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
        入院登记撤销_铜山县 = True
    End If
    Set rsRydjCx = Nothing
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_铜山县(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
'****************************************
'调用者　　　　　　：被clsInsure 的 LeaveSwap 过程调用
'功能说明　　　　　：调用出院登记接口
'步骤说明　　　　　：1、按接口规则进行检查
'　　　　　　　　　　2、调用出院接口
'　　　　　　　　　　3、执行出院登记过程(zl_保险帐户_出院)，更改病人的当前状态
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_destroyparams】销毁已分配空间
'****************************************
    'TODO:出院登记(str转往住院码要保存 到备注中)
    Dim rsCydj As New ADODB.Recordset, str出院特征 As String, str转往医院 As String
    Dim strTmp As String, strOut As String, lng序号 As Long, strTmp1 As String
    Dim str入院科室 As String, str入院床位 As String, str住院医生 As String, str门诊医生 As String
    Dim str顺序号 As String, lngReturn As Long, lng病种ID As Long, rsTemp As New ADODB.Recordset
    Dim STRERR As String
    On Error GoTo errHand

    '改变病人状态
    gstrSQL = "Select * from 病案主页  where 病人ID=[1] And 主页ID=[2]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "病案主页", lng病人ID, lng主页ID)
    If rsCydj.EOF Then
        MsgBox "未找到入院记录!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str入院科室 = rsCydj!入院科室ID
    str入院床位 = Nvl(rsCydj!入院病床, 0)
    str住院医生 = Nvl(rsCydj!住院医师, "")
    str门诊医生 = Nvl(rsCydj!门诊医师, "")
    
   
    Select Case rsCydj!出院方式
        Case "上转"
            str出院特征 = 1
        Case "下转"
            str出院特征 = 2
        Case Else
            str出院特征 = 0
    End Select
    
    gstrSQL = "Select * from 部门表 where ID=[1]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "入院科室", CLng(str入院科室))
    If rsCydj.EOF Then
        MsgBox "入院科室不对!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str入院科室 = rsCydj!编码

    gstrSQL = "Select 编号 from 人员表 A,人员性质说明 B where A.ID=B.人员ID and B.人员性质='医生' and  A.姓名=[1]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "医生", str住院医生)
    If rsCydj.EOF Then
        MsgBox "住院医生不对!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str住院医生 = rsCydj!编号
    
    gstrSQL = "Select 编号 from 人员表 A,人员性质说明 B where A.ID=B.人员ID and B.人员性质='医生' and  A.姓名=[1]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "医生", str门诊医生)
    If rsCydj.EOF Then
        MsgBox "门诊医生不对!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str门诊医生 = rsCydj!编号
    
    If str出院特征 <> 0 Then
    
        gstrSQL = "Select 医院编码,医院名称 from YYDA"
        Call OpenRecordset_OtherBase(rsCydj, "医院列表", , gcn铜山县)
        strTmp = ""
        strTmp1 = ""
        Do Until rsCydj.EOF
            strTmp = strTmp & rsCydj!医院编码 & ";"
            strTmp1 = strTmp1 & rsCydj!医院名称 & ";"
            rsCydj.MoveNext
        Loop
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
        strTmp1 = Mid(strTmp1, 1, Len(strTmp1) - 1)
        strOut = frmShowList.ShowME(strTmp & "||" & strTmp1, "转往医院编码||转往医院名称")
        
        str转往医院 = Split(strOut, ";")(0)
    End If
    gstrSQL = "select * from 保险帐户 where 病人ID=[1]"
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", lng病人ID)
    If rsCydj.EOF Then
        MsgBox "不是本医保病人!", vbInformation, gstrSysName
        Exit Function
    End If
   
    lng病种ID = rsCydj!病种ID
    str顺序号 = rsCydj!顺序号
    
    'Beging 检查病种
    gstrSQL = "Select * from icd10 where ID=" & lng病种ID
    Set rsCydj = zlDatabase.OpenSQLRecord(gstrSQL, "病种ID", lng病种ID)
    If rsCydj.EOF Then
            '强制选择病种
        gstrSQL = " Select A.ID,A.病种编码,A.病种名称,A.拼音码" & _
                " From Icd10 A "
        Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "确诊疾病")
        If rsTemp.State = 1 Then
            gstrSQL = " ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_铜山县 & ",'病种ID','''" & rsTemp!ID & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "病种ID")
        Else
            出院登记_铜山县 = False
            Exit Function
        End If
    End If
    
    'End  检查病种


    '1 分配空间
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "分配内存空间失败!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    '2 为参数赋值
    lngReturn = tsx_setstringparam(P_ZYXH, 0, str顺序号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str顺序号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_RYBQ, 0, str入院科室)
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & 0 & "," & str入院科室, lngReturn)
    
    lngReturn = tsx_setstringparam(P_RYCWH, 0, str入院床位) '
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & 0 & "," & str入院床位, lngReturn)
    
    lngReturn = tsx_setstringparam(P_YSRYH, 0, str住院医生) '
    Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & 0 & "," & str住院医生, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KDYSRYH, 0, str门诊医生) '
    Call WriteBusinessLOG("tsx_setstringparam", "P_KDYSRYH" & ", " & 0 & "," & str门诊医生, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

    '3 调用接口
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
    '4 取返回值
         
    End If
    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_铜山县 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_铜山县 & ",'备注','''" & str出院特征 & str转往医院 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存转院信息")
    
    出院登记_铜山县 = True
    Set rsCydj = Nothing
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_铜山县(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
'*******************************************
'调用者　　　　　　：被clsInsure 的 LeaveDelSwap 过程调用
'功能说明　　　　　：调用撤销出院登记或入院登记接口
'步骤说明　　　　　：1、按接口规则进行检查
'　　　　　　　　　　2、调用撤销出院登记或入院登记接口
'　　　　　　　　　　3、执行入院登记过程(zl_保险帐户_入院)，更改病人的当前状态
'调用过程清单及说明：
'　　【无】
'*******************************************
    'TODO:出院撤销
    On Error GoTo errHand

    '改变病人状态

    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_铜山县 & ",'备注','0')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存转院信息")

    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_铜山县 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")

    出院登记撤销_铜山县 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_铜山县(ByVal lng病人ID As Long, ByVal intinsure As Integer) As Currency
'*******************************************
'调用者　　　　　　：被clsInsure 的 SelfBalance 过程调用
'功能说明　　　　　：调用个人帐户余额查询接口或直接从保险帐户表中提取个人帐户余额
'步骤说明　　　　　：1、调用查询接口获取个人帐户余额并更新保险帐户表
'　　　　　　　　　　2、或者直接从保险帐户中提取个人帐户余额
'调用过程清单及说明：
'　　【无】
'******************************************
    'TODO:个人余额
    '个人余额 = 0
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = "Select Nvl(帐户余额,0) AS 个人帐户 From 保险帐户 " & _
              " Where 病人ID=[1] and 险类=[2]"
              
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", lng病人ID, TYPE_铜山县)
    个人余额_铜山县 = rsTemp!个人帐户
    Set rsTemp = Nothing
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If

    
End Function

Public Function 住院结算_铜山县(ByVal lng结帐ID As Long, ByVal lng病人ID As Long, ByVal intinsure As Integer, Optional ByRef strAdvance As String = "") As Boolean
'********************************
'调用者　　　　　　：被clsInsure 的 SettleSwap 过程调用
'功能说明　　　　　：完成本次住院费用的医保结算
'步骤说明　　　　　：1、调用住院结算接口
'　　　　　　　　　　2、如果住院结算返回的结算结果与住院预结算返回的不一致，需要调用
'　　　　　　　　　　zl_病人结算记录_Update过程进行修正
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string型参数
'　　【tsx_setdoubleparam】设置double型参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_getstringparam】取string型接口返回值
'　　【tsx_getdoubleparam】取double型接口返回值
'　　【tsx_destroyparams】销毁已分配空间
'*********************************
    
    'TODO:住院结算
    Dim rsZyjs As New ADODB.Recordset, lngReturn As Long, lng欠款重结ID As Long, str结算方式 As String
    Dim str流水号 As String, str医保卡号 As String, str个人编码 As String
    Dim str结算标志 As String, str出院日期 As String, lng主页ID As Long, str病种编码 As String, str转院特征 As String
    Dim str转往医院码 As String
    Dim str出院性质 As String, str住院交易流水号 As String
    Dim dbl总费用 As Double, dbl个人自费 As Double
    Dim dbl个人帐户支付 As Double, dbl统筹基金支付 As Double, dbl大病统筹支付 As Double
    Dim dbl公务员基金支付 As Double, dbl期末个人帐户 As Double, dbl期初个人帐户 As Double
    Dim dbl个人自付 As Double, STRERR As String
    On Error GoTo errHandle
    
    '>Beging 准备参数
    str结算标志 = 1

    gstrSQL = "Select * from 保险帐户 where 险类=[1] And 病人ID=[2]"
    Call WriteBusinessLOG("", gstrSQL, "读保险帐户")
    
    Set rsZyjs = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", intinsure, lng病人ID)
    lng欠款重结ID = Nvl(rsZyjs!欠款重结ID, 0)
    str流水号 = rsZyjs!顺序号
    str医保卡号 = rsZyjs!卡号
    str个人编码 = rsZyjs!医保号
    str病种编码 = rsZyjs!病种ID
    
    If Nvl(rsZyjs!备注, "0") = 0 Then
        str转院特征 = "0"
        str转往医院码 = ""
    Else
        str转院特征 = Mid(rsZyjs!备注, 1, 1)
        If str转院特征 = 0 Then
            str转往医院码 = ""
        Else
            str转往医院码 = Mid(rsZyjs!备注, 2, 1)
        End If
    End If
    
    gstrSQL = "Select * from icd10 where ID=" & str病种编码
    Call OpenRecordset_OtherBase(rsZyjs, "ICD10", , gcn铜山县)
    If rsZyjs.RecordCount > 0 Then
        str病种编码 = rsZyjs!病种编码
    End If
    
    gstrSQL = "Select * from 病人信息 where 病人ID=[1] And 险类=[2]"
    Set rsZyjs = zlDatabase.OpenSQLRecord(gstrSQL, "病案主页", lng病人ID, intinsure)
    If Nvl(rsZyjs!出院时间, 0) = 0 Then
        Err.Raise 9000, gstrSysName, "不支持中途结算，请先办理出院后再进行结算！"
        Exit Function
    End If
    str出院日期 = Format(Nvl(rsZyjs!出院时间, Now()), "yyyyMMdd")
    lng主页ID = Nvl(rsZyjs!住院次数, 0)
    
    gstrSQL = "Select *  From 诊断情况 Where 病人ID=[1]" & _
              " And 主页ID=[2] And 诊断类型=3 And 诊断次序=1"
              
    Set rsZyjs = zlDatabase.OpenSQLRecord(gstrSQL, "出院情况", lng病人ID, lng主页ID)
    If rsZyjs.EOF Then
        str出院性质 = 2
    Else
    
        Select Case Nvl(rsZyjs!出院情况, "好转")
            Case "治愈"
                str出院性质 = 1
            Case "好转"
                str出院性质 = 2
            Case "未愈"
                str出院性质 = 3
            Case "死亡"
                str出院性质 = 4
            Case Else
                str出院性质 = 5
        End Select
    End If
   
    '>End 准备参数
    
    '>>Beging 欠款重结
    If lng欠款重结ID > 0 Then
        gstrSQL = "Select * from 保险结算记录 where 记录ID=[1]"
        Set rsZyjs = zlDatabase.OpenSQLRecord(gstrSQL, "保险结算记录", lng欠款重结ID)
        If rsZyjs.EOF Then
            MsgBox "未找到原始结帐记录，不能执行后续操作！"
            Exit Function
        End If
        str流水号 = rsZyjs!支付顺序号
        '>>Beging 调QKCJS交易
        '1 分配空间
        If tsx_createparams(1024, 1024) = -1 Then
            Err.Raise 9000, gstrSysName, "分配内存空间失败!"
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
        
        '2 为参数赋值
        lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
        Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
        
        lngReturn = tsx_setstringparam(P_TBR, 0, str个人编码) '    C9  个人编号    参保人员个人编号
        Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str个人编码, lngReturn)
            '20051103 文档中要求传P_ZYXH ,实际应传P_DJH
        lngReturn = tsx_setstringparam(P_DJH, 0, str流水号)
        Call WriteBusinessLOG("tsx_setstringparam", "P_DJH" & ", " & 0 & "," & str流水号, lngReturn)

        lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
        Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

        '3 调用接口
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
        '4 取返回值
          str住院交易流水号 = Space(32)
          lngReturn = tsx_getstringparam(P_DJH, 0, str住院交易流水号)   '   C20 单据号
          Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str住院交易流水号, lngReturn)
          str住院交易流水号 = Mid(str住院交易流水号, 1, InStr(str住院交易流水号, Chr(0)) - 1)
           
          lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl个人自费) '  D   个人自费
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl个人自费, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl个人自付)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl个人自付, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl个人帐户支付)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl个人帐户支付, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dbl统筹基金支付)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dbl统筹基金支付, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl大病统筹支付)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl大病统筹支付, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl公务员基金支付)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl公务员基金支付, lngReturn)
         
          lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl期初个人帐户)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl期初个人帐户, lngReturn)
          
          lngReturn = tsx_getdoubleparam(P_QMGRZH, 0, dbl期末个人帐户)
          Call WriteBusinessLOG("tsx_getdoubleparam", "P_QMGRZH, 0, " & dbl期末个人帐户, lngReturn)
             
        End If
        '5 销毁已分配空间
        lngReturn = tsx_destroyparams()
        Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
        '>>end 调QKCJS交易
        dbl总费用 = Abs(dbl个人自费) + Abs(dbl个人自付) + dbl个人帐户支付 + dbl统筹基金支付 + dbl大病统筹支付 _
                   + dbl大病统筹支付 + dbl公务员基金支付
                   
         If dbl个人帐户支付 <> 0 Then
             str结算方式 = "||个人帐户|" & dbl个人帐户支付
         End If
         
         '2
         If dbl统筹基金支付 <> 0 Then
             str结算方式 = str结算方式 & "||统筹基金|" & dbl统筹基金支付
         End If
        
         '3
         If dbl公务员基金支付 <> 0 Then
             str结算方式 = str结算方式 & "||公务员基金|" & dbl公务员基金支付
         End If
         '4
         If dbl大病统筹支付 <> 0 Then
             str结算方式 = str结算方式 & "||大病支付|" & dbl大病统筹支付
         End If
         
         '如果存在
         If str结算方式 <> "" Then
             str结算方式 = Mid(str结算方式, 3)
             #If gverControl < 2 Then
                gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
             #Else
                strAdvance = str结算方式
                gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
             #End If
                
             Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
         End If
                   
        '**保存保险结算记录**
        #If gverControl < 2 Then
            gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
                Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
                0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
                dbl总费用 & "," & Abs(dbl个人自费) & "," & Abs(dbl个人自付) & "," & dbl总费用 - dbl个人自费 - dbl个人自付 & "," & _
                dbl统筹基金支付 + dbl公务员基金支付 & "," & dbl大病统筹支付 & "," & _
                0 & "," & dbl个人帐户支付 & ",'" & str住院交易流水号 & "',null,null,Null)"
        #Else
            gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
                Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
                0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
                dbl总费用 & "," & Abs(dbl个人自费) & "," & Abs(dbl个人自付) & "," & dbl总费用 - dbl个人自费 - dbl个人自付 & "," & _
                dbl统筹基金支付 + dbl公务员基金支付 & "," & dbl大病统筹支付 & "," & _
                0 & "," & dbl个人帐户支付 & ",'" & str住院交易流水号 & "',null,null,Null,1)"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
        
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'欠款重结ID','" & 0 & " ')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "取消欠款重结")
        
        住院结算_铜山县 = True
        Exit Function
    End If
    '>>End 欠款重结
    
    '>>Beging 调cyjs 交易
    '1 分配空间
    If tsx_createparams(1024, 1024) = -1 Then
        Err.Raise 9000, gstrSysName, "分配内存空间失败!"
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    
    '2 为参数赋值
    lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, str个人编码) '    C9  个人编号    参保人员个人编号
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str个人编码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, str医保卡号) '   C3  医保卡序号  参保人员IC卡序号
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & str医保卡号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_ZYXH, 0, str流水号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str流水号, lngReturn)

    lngReturn = tsx_setstringparam(P_LB, 0, str结算标志) 'C1  备注    预结算特征(0-预结算,1-出院结算)
    Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & 0 & "," & str结算标志, lngReturn)

    lngReturn = tsx_setstringparam(P_CYYMD, 0, str出院日期)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYYMD" & ", " & 0 & "," & str出院日期, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CYXZ, 0, str出院性质)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYXZ" & ", " & 0 & "," & str出院性质, lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZM, 0, str病种编码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_BZM" & ", " & 0 & "," & str病种编码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_ZYTZ, 0, str转院特征) 'C1 0 - 不是转院, 1 - 上转, 2 - 下转
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYTZ" & ", " & 0 & "," & str转院特征, lngReturn)

    lngReturn = tsx_setstringparam(P_ZWYYM, 0, str转往医院码) 'C5  转往医院码  转往医院的医院编码
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZWYYM" & ", " & 0 & "," & str转往医院码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

    '3 调用接口
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
    '4 取返回值
         lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl总费用) 'D 总费用
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl总费用, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl个人自费) '  D   个人自费
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl个人自费, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl个人自付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl个人自付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl个人帐户支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl个人帐户支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dbl统筹基金支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dbl统筹基金支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl大病统筹支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl大病统筹支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl公务员基金支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl公务员基金支付, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl期初个人帐户)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl期初个人帐户, lngReturn)
         
    End If
    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
    '>>End 调cyjs 交易
    
    '>>> Beging 调用CYQR交易
    
        '1 分配空间
      If tsx_createparams(1024, 1024) = -1 Then
          Err.Raise 9000, gstrSysName, "分配内存空间失败!"
          Exit Function
      End If
      Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
      
      '2 为参数赋值
      lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
      Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
      
      lngReturn = tsx_setstringparam(P_TBR, 0, str个人编码) '    C9  个人编号    参保人员个人编号
      Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str个人编码, lngReturn)
      
      lngReturn = tsx_setstringparam(P_KXH, 0, str医保卡号) '   C3  医保卡序号  参保人员IC卡序号
      Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & str医保卡号, lngReturn)
      
      lngReturn = tsx_setstringparam(P_ZYXH, 0, str流水号)
      Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str流水号, lngReturn)
  
      lngReturn = tsx_setstringparam(P_CYYMD, 0, str出院日期)
      Call WriteBusinessLOG("tsx_setstringparam", "P_CYYMD" & ", " & 0 & "," & str出院日期, lngReturn)
      
      lngReturn = tsx_setstringparam(P_CYXZ, 0, str出院性质)
      Call WriteBusinessLOG("tsx_setstringparam", "P_CYXZ" & ", " & 0 & "," & str出院性质, lngReturn)
      
      lngReturn = tsx_setstringparam(P_BZM, 0, str病种编码)
      Call WriteBusinessLOG("tsx_setstringparam", "P_BZM" & ", " & 0 & "," & str病种编码, lngReturn)
      
      lngReturn = tsx_setstringparam(P_ZYTZ, 0, str转院特征) 'C1 0 - 不是转院, 1 - 上转, 2 - 下转
      Call WriteBusinessLOG("tsx_setstringparam", "P_ZYTZ" & ", " & 0 & "," & str转院特征, lngReturn)
  
      lngReturn = tsx_setstringparam(P_ZWYYM, 0, str转往医院码) 'C5  转往医院码  转往医院的医院编码
      Call WriteBusinessLOG("tsx_setstringparam", "P_ZWYYM" & ", " & 0 & "," & str转往医院码, lngReturn)
      
      lngReturn = tsx_setdoubleparam(P_QCGRZH, 0, Val(Format(dbl期初个人帐户, "0.00")))
      Call WriteBusinessLOG("tsx_setdoubleparam", "P_QCGRZH, 0, " & dbl期初个人帐户, lngReturn)
      
      lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
      Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

      lngReturn = tsx_setstringparam(P_CXLB, 0, mstr使用个人帐户支付) 'C1  个人帐户支付    是否使用个人帐户支付(0-否,1-是
      Call WriteBusinessLOG("tsx_setstringparam", "P_CXLB" & ", " & 0 & "," & mstr使用个人帐户支付, lngReturn)

      '3 调用接口
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
      '4 取返回值
           str住院交易流水号 = Space(32)
           lngReturn = tsx_getstringparam(P_DJH, 0, str住院交易流水号)  '   C20 单据号
           Call WriteBusinessLOG("tsx_getstringparam", "P_DJH, 0, " & str住院交易流水号, lngReturn)
           str住院交易流水号 = Mid(str住院交易流水号, 1, InStr(str住院交易流水号, Chr(0)) - 1)
           
           lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl总费用) 'D 总费用
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl总费用, lngReturn)
          
           lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl个人自费) '  D   个人自费
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl个人自费, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_GRZF, 0, dbl个人自付)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZF, 0, " & dbl个人自付, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl个人帐户支付)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl个人帐户支付, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dbl统筹基金支付)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dbl统筹基金支付, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl大病统筹支付)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl大病统筹支付, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl公务员基金支付)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl公务员基金支付, lngReturn)
          
           lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl期初个人帐户)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl期初个人帐户, lngReturn)
           
           lngReturn = tsx_getdoubleparam(P_QMGRZH, 0, dbl期末个人帐户)
           Call WriteBusinessLOG("tsx_getdoubleparam", "P_QMGRZH, 0, " & dbl期末个人帐户, lngReturn)
         
      End If
      '5 销毁已分配空间
      lngReturn = tsx_destroyparams()
      Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
      
      '**保存保险结算记录**
      gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
          Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
          0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
          dbl总费用 & "," & dbl个人自费 & "," & dbl个人自付 & "," & dbl总费用 - dbl个人自费 - dbl个人自付 & "," & _
          dbl统筹基金支付 + dbl公务员基金支付 & "," & dbl大病统筹支付 & "," & _
          0 & "," & dbl个人帐户支付 & ",'" & str住院交易流水号 & "',null,null,Null)"
      Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")

    '>>> End 调用CYQR交易
    Set rsZyjs = Nothing
    住院结算_铜山县 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院虚拟结算_铜山县(rsExse As Recordset, lng病人ID As Long, str医保号 As String, str密码 As String, ByVal intinsure As Integer) As String
'*********************************
'调用者　　　　　　：被clsInsure 的 WipeoffMoney 过程调用
'功能说明　　　　　：完成本次住院费用的医保预结算
'步骤说明　　　　　：1、需要将未上传的处方明细上传到中心（如果平常记帐时会实时上传，
'　　　　　　　　　　则本次实际仅上传了自动计算费用明细）
'　　　　　　　　　　2、根据接口性质（每条单独上传或打包上传），将成功上传的明细打上上传标记
'　　　　　　　　　　3、调用住院预结算接口
'　　　　　　　　　　4、按规定格式返回结算结果串，请参见门诊预结算
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string型参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_getdoubleparam】取double型接口返回值
'　　【tsx_destroyparams】销毁已分配空间
'*********************************
    
    'TODO:虚拟结算

    Dim rsZyxnjs As New ADODB.Recordset, lngReturn As Long, str消息 As String, str结算方式 As String
    Dim str流水号 As String, str医保卡号 As String, str个人编码 As String, str病种编码 As String
    Dim str转院特征 As String, str转往医院码      As String
    Dim str结算标志 As String, str出院日期 As String, lng主页ID As Long
    Dim str出院性质 As String, str住院交易流水号 As String
    Dim dbl总费用 As Double, dbl个人自费 As Double, dbl起付线自付 As Double, dbl统筹自付 As Double, dbl大病自付 As Double
    Dim dbl个人帐户支付 As Double, dbl统筹基金支付 As Double, dbl大病统筹支付 As Double, dbl封顶自付 As Double
    Dim dbl公务员基金支付 As Double, dbl期末个人帐户 As Double, dbl期初个人帐户 As Double
    Dim dbl个人自付 As Double, rsCFMX As New ADODB.Recordset, STRERR As String
    On Error GoTo errHandle
    
    '>beging 补传未上传记录
       ' 查出未上传记录的,记录性质 , 记录状态, NO 调用记帐上传
    gstrSQL = "Select distinct 记录性质,记录状态,NO From 住院费用记录 A,保险帐户 B,病人信息 C " & _
              " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID And A.主页ID=C.住院次数" & _
              " And nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 and A.记帐费用=1 And A.操作员姓名 is not null " & _
              " AND A.实收金额 IS NOT NULL And B.病人ID=[1] And B.险类=[2]"
     Call WriteBusinessLOG("取病人费用记录", gstrSQL, "")
     
     Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "取病人费用记录", lng病人ID, intinsure)

     Do Until rsZyxnjs.EOF
         Call WriteBusinessLOG("调用处方上传_铜山县", "2" & "," & rsZyxnjs!NO & "," & rsZyxnjs!记录性质 & "," & rsZyxnjs!记录状态 & "," & str消息 & "," & lng病人ID & "," & intinsure, lngReturn)
         Call 处方上传_铜山县(2, rsZyxnjs!NO, rsZyxnjs!记录性质, rsZyxnjs!记录状态, str消息, lng病人ID, intinsure)
         Call WriteBusinessLOG("完成调用处方上传_铜山县", "2" & "," & rsZyxnjs!NO & "," & rsZyxnjs!记录性质 & "," & rsZyxnjs!记录状态 & "," & str消息 & "," & lng病人ID & "," & intinsure, lngReturn)

         rsZyxnjs.MoveNext
         
'         If rsZyxnjs!记录状态 > 1 Then
'         gstrSQL = "Select * from 病人费用记录 where NO='" & rsZyxnjs!NO & " And 记录性质=" & rsZyxnjs!记录性质 & _
'                    " and  记录状态=" & rsZyxnjs!记录状态 & " And 病人ID=" & lng病人ID
'
'         Call OpenRecordset(rsCFMX, "处方明细")
'         Do Until rsCFMX.EOF
'            rsCFMX.MoveNext
'         Loop
'         End If
     Loop
    '>end 补传未上传记录
    
    '>>Beging 欠款重结
    gstrSQL = "Select * from 保险帐户 where 险类=[1] And 病人ID=[2] And nvl(欠款重结ID,0)>0"
    Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", intinsure, lng病人ID)
    If Not rsZyxnjs.EOF Then
        MsgBox "提示：欠款重结时，预结算为全自费！", vbInformation, gstrSysName
        str结算方式 = "统筹基金;" & 0 & ";0"
        str结算方式 = str结算方式 & "|大病支付;" & 0 & ";0"
        str结算方式 = str结算方式 & "|公务员基金;" & 0 & ";0"
        str结算方式 = str结算方式 & "|个人帐户;" & 0 & ";0"
        
        住院虚拟结算_铜山县 = str结算方式
        Call WriteBusinessLOG("退出欠款重结预结算!", "", "")
        Exit Function
    End If
    '>>End 欠款重结
    
    
    '>Beging 准备参数
    str结算标志 = 0
    mstr使用个人帐户支付 = 0
    '结算调用才调试
    
    '2008-12-12 陈玉强要求去掉
'    If str密码 = 1 Then
'        If MsgBox("是否使用个人帐户支付？", vbQuestion Or vbYesNo Or vbDefaultButton1, gstrSysName) = vbYes Then
'            mstr使用个人帐户支付 = 1
'        End If
'    End If

    gstrSQL = "Select * from 保险帐户 where 险类=[1] And 病人ID=[2]"
    Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", intinsure, lng病人ID)
    
    If rsZyxnjs.EOF Then
        MsgBox "不是本医保病人!", vbInformation, gstrSysName
        Exit Function
    End If
    
    str流水号 = rsZyxnjs!顺序号
    str医保卡号 = rsZyxnjs!卡号
    str个人编码 = rsZyxnjs!医保号
    str病种编码 = rsZyxnjs!病种ID
    
    If Nvl(rsZyxnjs!备注, "0") = 0 Then
        str转院特征 = "0"
        str转往医院码 = ""
    Else
        str转院特征 = Mid(rsZyxnjs!备注, 1, 1)
        If str转院特征 = 0 Then
            str转往医院码 = ""
        Else
            str转往医院码 = Mid(rsZyxnjs!备注, 2, 1)
        End If
    End If
    
    gstrSQL = "Select * from icd10 where ID=" & str病种编码
    Call OpenRecordset_OtherBase(rsZyxnjs, "ICD10", , gcn铜山县)
    
    If rsZyxnjs.EOF Then
        MsgBox "未指定病种或指定的病种未采用医保规定编码,不能结算!", vbInformation, gstrSysName
        Exit Function
    End If

    str病种编码 = rsZyxnjs!病种编码
    
    gstrSQL = "Select * from 病人信息 where 病人ID=[1] And 险类=[2]"
    Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "病案主页", lng病人ID, intinsure)
    
    str出院日期 = Format(Nvl(rsZyxnjs!出院时间, Now()), "yyyyMMdd")
    lng主页ID = Nvl(rsZyxnjs!住院次数, 0)
    
    gstrSQL = "Select *  From 诊断情况 Where 病人ID=[1] And 主页ID=[2] And 诊断类型=3 And 诊断次序=1"
              
    Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "出院情况", lng病人ID, lng主页ID)
    If rsZyxnjs.EOF Then
        str出院性质 = 2
    Else
    
        Select Case Nvl(rsZyxnjs!出院情况, "好转")
            Case "治愈"
                str出院性质 = 1
            Case "好转"
                str出院性质 = 2
            Case "未愈"
                str出院性质 = 3
            Case "死亡"
                str出院性质 = 4
            Case Else
                str出院性质 = 5
        End Select
    End If
    '>End 准备参数
    
    '>>Beging 调cyjs 交易
    '1 分配空间
    If tsx_createparams(1024, 1024) = -1 Then
        MsgBox "分配内存空间失败!", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
    
    '2 为参数赋值
    lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_TBR, 0, str个人编码) '    C9  个人编号    参保人员个人编号
    Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str个人编码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_KXH, 0, str医保卡号) '   C3  医保卡序号  参保人员IC卡序号
    Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & str医保卡号, lngReturn)
    
    lngReturn = tsx_setstringparam(P_ZYXH, 0, str流水号)
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & 0 & "," & str流水号, lngReturn)

    lngReturn = tsx_setstringparam(P_LB, 0, str结算标志) 'C1  备注    预结算特征(0-预结算,1-出院结算)
    Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & 0 & "," & str结算标志, lngReturn)

    lngReturn = tsx_setstringparam(P_CYYMD, 0, str出院日期)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYYMD" & ", " & 0 & "," & str出院日期, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CYXZ, 0, str出院性质)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYXZ" & ", " & 0 & "," & str出院性质, lngReturn)
    
    lngReturn = tsx_setstringparam(P_BZM, 0, str病种编码)
    Call WriteBusinessLOG("tsx_setstringparam", "P_BZM" & ", " & 0 & "," & str病种编码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_ZYTZ, 0, str转院特征) 'C1 0 - 不是转院, 1 - 上转, 2 - 下转
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZYTZ" & ", " & 0 & "," & str转院特征, lngReturn)

    lngReturn = tsx_setstringparam(P_ZWYYM, 0, str转往医院码) 'C5  转往医院码  转往医院的医院编码
    Call WriteBusinessLOG("tsx_setstringparam", "P_ZWYYM" & ", " & 0 & "," & str转往医院码, lngReturn)
    
    lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
    Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

    lngReturn = tsx_setstringparam(P_CZYMD, 0, str出院日期)
    Call WriteBusinessLOG("tsx_setstringparam", "P_CYYMD" & ", " & 0 & "," & str出院日期, lngReturn)

    '3 调用接口
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
    '4 取返回值
         lngReturn = tsx_getdoubleparam(P_ZFY, 0, dbl总费用) 'D 总费用
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_ZFY, 0, " & dbl总费用, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_QZ_XXFD, 0, dbl个人自付) ' 乙类药品先行自付部分
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_XXFD, 0, " & dbl个人自付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZL, 0, dbl个人自费) '  D   个人自费
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZL, 0, " & dbl个人自费, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_QZ_QFFD, 0, dbl起付线自付) '起付线自付部分 P_QZ_QFFDD
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_QFFD, 0, " & dbl起付线自付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_TCJJZF, 0, dbl统筹基金支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_TCJJZF, 0, " & dbl统筹基金支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_QZ_JBFD, 0, dbl统筹自付) '统筹分段自付部分 P_QZ_JBFD
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_JBFD, 0, " & dbl统筹自付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_BCYLZF, 0, dbl大病统筹支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_BCYLZF, 0, " & dbl大病统筹支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_QZ_DBFD, 0, dbl大病自付) ' 大病自付部分 P_QZ_DBFD
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_DBFD, 0, " & dbl大病自付, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_GWYBZZF, 0, dbl公务员基金支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GWYBZZF, 0, " & dbl公务员基金支付, lngReturn)
         
         lngReturn = tsx_getdoubleparam(P_GRZHZF, 0, dbl个人帐户支付)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_GRZHZF, 0, " & dbl个人帐户支付, lngReturn)
        
         lngReturn = tsx_getdoubleparam(P_QCGRZH, 0, dbl期初个人帐户)
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QCGRZH, 0, " & dbl期初个人帐户, lngReturn)
    
         lngReturn = tsx_getdoubleparam(P_QZ_CFD, 0, dbl封顶自付) '起封顶线自付部分 P_QZ_CFD
         Call WriteBusinessLOG("tsx_getdoubleparam", "P_QZ_CFD, 0, " & dbl封顶自付, lngReturn)
         
    End If
    '5 销毁已分配空间
    lngReturn = tsx_destroyparams()
    Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)
    '>>End 调cyjs 交易
    
    If lngReturn <> -1 Then
        '>Beging 检查总金额是否相等
        gstrSQL = "Select sum(nvl(实收金额,0))-sum(nvl(结帐金额,0)) as 未结费用 From 住院费用记录 Where nvl(记录状态,0)<>0 and 记帐费用=1 And 病人ID=[1]"
        Set rsZyxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "未结费用", lng病人ID)
        
        If Format(Val(Nvl(rsZyxnjs.Fields!未结费用, 0)), "0.00") <> Format(dbl总费用, "0.00") Then
            Dim intButton As Integer
            intButton = MsgBox("医院的费用总金额(" & Format(Nvl(rsZyxnjs.Fields!未结费用, 0), "0.00") & ")与医保中心的费用总额(" & Format(dbl总费用, "0.00") & ")不等，是否继续？" & vbNewLine & _
                             "选[是]，忽略此问题，继续结算。" & vbNewLine & _
                             "选[否]，停止结算操作，重传费用明细，您可以稍后重新结算。" & vbNewLine & _
                             "选[取消]，停止结算操作，您可以手工确认费用后再重新结算。", vbQuestion Or vbYesNoCancel Or vbDefaultButton2, gstrSysName)
            If intButton = vbNo Then
                gstrSQL = "Update 住院费用记录 Set 是否上传=0 Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID & " And 是否上传=1"
                gcnOracle.Execute gstrSQL
                Exit Function
            ElseIf intButton = vbCancel Then
                Exit Function
            End If
        End If
        ''>End 检查总金额是否相等
        
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_铜山县 & ",'帐户余额','''" & Format(dbl期初个人帐户, "0.00") & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新帐户余额")
        
        If mstr使用个人帐户支付 = 1 Then
            If dbl期初个人帐户 >= dbl个人自付 Then
                dbl个人帐户支付 = dbl个人自付
                dbl个人自付 = 0
            Else
                dbl个人帐户支付 = dbl期初个人帐户
                dbl个人自付 = dbl个人自付 - dbl个人帐户支付
            End If
        End If
        
        str结算方式 = "统筹基金;" & dbl统筹基金支付 & ";0"
        str结算方式 = str结算方式 & "|大病支付;" & dbl大病统筹支付 & ";0"
        str结算方式 = str结算方式 & "|公务员基金;" & dbl公务员基金支付 & ";0"
        str结算方式 = str结算方式 & "|个人帐户;" & dbl个人帐户支付 & ";0"
        
        住院虚拟结算_铜山县 = str结算方式
        dbl期末个人帐户 = dbl期初个人帐户 - dbl个人帐户支付

        '>>Beging 写入xybjz表,供打一日期清单用
        gstrSQL = "Delete xybjz where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID
        gcn铜山县.Execute gstrSQL
        
        gstrSQL = "Insert into Xybjz(病人id,主页id,总额, 部分自理, 完全自理, 统筹基金支付,大病基金支付, 公务员统筹支付,起付线支付,统筹基金不支付,大病基金不支付,封顶线自付) values(" & _
                 lng病人ID & "," & lng主页ID & "," & dbl总费用 & "," & dbl个人自付 & "," & dbl个人自费 & "," & dbl统筹基金支付 & "," & _
                 dbl大病统筹支付 & "," & dbl公务员基金支付 & "," & dbl起付线自付 & "," & dbl统筹自付 & "," & dbl大病自付 & "," & dbl封顶自付 & ")"
        gcn铜山县.Execute gstrSQL
        
        
        '>>End 写入xybjz表,供打一日期清单用
        
    Else
        MsgBox "预结算失败!", vbInformation, gstrSysName
        住院虚拟结算_铜山县 = ""
    End If
    'Set rsZyxnjs = Nothing
    
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_铜山县(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
'********************************

'调用者　　　　　　：被clsInsure 的 SettleDelSwap 过程调用
'功能说明　　　　　：完成本次住院结算的作废
'步骤说明　　　　　：1、调用住院结算作废接口
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string型参数
'　　【tsx_setdoubleparam】设置double型参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_getstringparam】取string型接口返回值
'　　【tsx_destroyparams】销毁已分配空间
'*********************************

    'TODO:住院结算冲销
    Dim rsZyjsCx As New ADODB.Recordset, bln结算冲销 As Boolean, lng病人ID As Long, lng冲销ID As Long
    Dim str原流水号 As String, str个人编号 As String, str医保卡号 As String
    Dim lngReturn As Long, bln欠款重结 As Boolean, STRERR As String
    
    On Error GoTo errHand
    
    gstrSQL = "Select * From 结算方式 Where 名称 In (" & _
                    "Select 结算方式 From 病人预交记录 " & _
                    "where 结帐ID=[1]) And 性质>=3 And  性质<=4"
    Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, "结算方式", lng结帐ID)
    If Not rsZyjsCx.EOF Then
        'MsgBox "本医保不支持冲销操作!" & vbCrLf & "如要冲销已结算单据，请到医保中心办理！", vbInformation, gstrSysName
        bln结算冲销 = True
    End If
    
    gstrSQL = "Select * From 保险结算记录 Where 性质=2 And 记录ID=[1] and 险类=[2]"
    Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, "取结算流水号", lng结帐ID, intinsure)
    If rsZyjsCx.EOF Then
        Err.Raise 9000, gstrSysName, "没有找到原始结算记录，无法进行住院结算冲销！"
        Exit Function
    End If
    lng病人ID = rsZyjsCx!病人ID
    str原流水号 = rsZyjsCx!支付顺序号
    
    If bln结算冲销 = True Then
        bln欠款重结 = False
    Else
        '欠款重结
        If MsgBox("确认是办理欠款重结交易吗？", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            bln欠款重结 = True
        Else
            bln欠款重结 = False
        End If
    End If
    
    If bln欠款重结 = True Then
      
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_铜山县 & ",'欠款重结ID','" & lng结帐ID & " ')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "设欠款重结")
        住院结算冲销_铜山县 = True
    Else
        '结算冲销
        gstrSQL = "Select Distinct A.ID From 病人结帐记录 A,病人结帐记录 B Where A.No=B.No And A.记录状态=2 And B.Id=[1]"
        Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
        lng冲销ID = rsZyjsCx("ID")
        
        gstrSQL = "Select * from 保险帐户 where 病人ID=[1] And 险类=[2]"
        Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", lng病人ID, TYPE_铜山县)
        str个人编号 = rsZyjsCx!医保号
        str医保卡号 = rsZyjsCx!卡号
        '1 分配空间
        If tsx_createparams(1024, 1024) = -1 Then
            Err.Raise 9000, gstrSysName, "分配内存空间失败!", vbInformation, gstrSysName
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
        '2 为参数赋值
        lngReturn = tsx_setstringparam(P_JGM, 0, g常用信息_铜山.医院码)
        Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & 0 & "," & g常用信息_铜山.医院码, lngReturn)
        
        lngReturn = tsx_setstringparam(P_TBR, 0, str个人编号) '    C9  个人编号    参保人员个人编号
        Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str个人编号, lngReturn)
        
        lngReturn = tsx_setstringparam(P_KXH, 0, str医保卡号) '   C3  医保卡序号  参保人员IC卡序号
        Call WriteBusinessLOG("tsx_setstringparam", "P_KXH" & ", " & 0 & "," & str医保卡号, lngReturn)
        
        lngReturn = tsx_setstringparam(P_DJH, 0, str原流水号) '   流水号
        Call WriteBusinessLOG("tsx_setstringparam", "P_DJH" & ", " & 0 & "," & str原流水号, lngReturn)
        
        lngReturn = tsx_setstringparam(P_CZRYH, 0, g常用信息_铜山.操作员号) ' C10 操作人员号
        Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH,0," & g常用信息_铜山.操作员号, lngReturn)

        '3 调用接口
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
            '4 取返回值
            lngReturn = tsx_getstringparam(P_DJH, 0, str原流水号)
            Call WriteBusinessLOG("tsx_getstringparam", "P_RYH, 0, " & str原流水号, lngReturn)
             
        End If
        
        If lngReturn <> -1 Then
            gstrSQL = "Select * From 保险结算记录 Where 性质=2 And 记录ID=[1] and 险类=[2]"
            Set rsZyjsCx = zlDatabase.OpenSQLRecord(gstrSQL, "保险结算记录", lng结帐ID, TYPE_铜山县)
     
            gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_铜山县 & "," & lng病人ID & "," & _
                Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
                0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
                -1 * Nvl(rsZyjsCx!发生费用金额, 0) & "," & _
                -1 * Nvl(rsZyjsCx!全自付金额, 0) & "," & _
                -1 * Nvl(rsZyjsCx!首先自付金额, 0) & "," & _
                -1 * Nvl(rsZyjsCx!进入统筹金额, 0) & "," & _
                -1 * Nvl(rsZyjsCx!统筹报销金额, 0) & "," & _
                -1 * Nvl(rsZyjsCx!大病自付金额, 0) & "," & 0 & "," & _
                -1 * Nvl(rsZyjsCx!个人帐户支付, 0) & ",'" & str原流水号 & "',null,null,Null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
           
        End If
        '5 销毁已分配空间
        lngReturn = tsx_destroyparams()
        Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

        住院结算冲销_铜山县 = True
    End If
    Set rsZyjsCx = Nothing
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 处方上传_铜山县(ByVal int类别 As Integer, ByVal str单据号 As String, ByVal int性质 As Integer, ByVal int状态 As Integer, _
        str消息 As String, Optional ByVal lng病人ID As Long = 0, Optional ByVal intinsure As Integer = 0) As Boolean

'*******************************************************
'调用者　　　　　　：被clsInsure 的 TranChargeDetail 过程调用
'功能说明　　　　　：住院记帐保存时或保存后，根据参数决定（support...，可参见Getcapability）
'步骤说明　　　　　：1、提取本单据的处方明细
'　　　　　　　　　　2、仅上传本医保的病人处方
'　　　　　　　　　　3、根据接口性质（每条单独上传或打包上传），将成功上传的明细打上上传标记
'调用过程清单及说明：
'　　【tsx_createparams】分配空间
'　　【tsx_setstringparam】设置string型参数
'　　【tsx_setdoubleparam】设置double型参数
'　　【tsx_setlongparam】设置long型参数
'　　【tsx_getlasterr】 取得上次错误信息
'　　【tsx_jkcall】 调用接口
'　　【tsx_getstringparam】取string型接口返回值
'　　【tsx_getdoubleparam】取double型接口返回值
'　　【tsx_destroyparams】销毁已分配空间
'*******************************************************
    
    'TODO:处方上传
    '以下是打上传标记的示范代码（有两种不同的方式，可根据你的需要使用）
    '注意事项，如果是保存的同时上传，因为事务没有提交，请使用全局连接对象gcnOracle更新上传标志；
    '如果是保存后上传，可以新打开一个连接对象，也可以使用全局连接对象gcnOracle更新上传标志
'    gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
'    cn上传.Execute gstrSQL, , adCmdStoredProc
'    '或者
    Dim rs记帐明细 As New ADODB.Recordset, rsMzxnjs As New ADODB.Recordset, rs费用记录 As New ADODB.Recordset, rsXybmx As New ADODB.Recordset
    Dim rsCfsc As New ADODB.Recordset, lngPatiID As Long '病人ID,为了和参数的区别,用字母用于上传记帐表
    Dim str流水号 As String, lngReturn As Long, str药品 As String, lngPageID As Long
    Dim str个人编号 As String, str费用日期 As String, bln全部冲销 As Boolean, dbl医保欠款 As Double
    Dim lngCount As Long '记录序号
    Dim bln是否调用接口 As Boolean
    Dim STRERR As String, str部门 As String
    
    处方上传_铜山县 = True
        '根据NO号,提取病人ID
    
    
    gstrSQL = "Select distinct  decode(nvl(A.医嘱序号,-99),-99,to_char(A.发生时间,'YYYY-MM-DD'),to_char(A.登记时间,'YYYY-MM-DD')) as 费用日期,A.病人ID,A.主页ID from 住院费用记录 A,保险帐户 B " & _
            "where A.病人ID=B.病人ID And nvl(A.附加标志,0)<>9 And Nvl(a.实收金额, 0)<>0" & _
            " And A.记录性质=[1] and  A.记录状态=[2] And A.NO=[3]" & _
            " And B.险类=[4]"
    If lng病人ID <> 0 Then
        gstrSQL = gstrSQL & " And A.病人ID=[5]"
    End If
    Call WriteBusinessLOG("取要上传单据(处方上传)", gstrSQL, "")
    Set rsCfsc = zlDatabase.OpenSQLRecord(gstrSQL, "取病人ID", int性质, int状态, str单据号, intinsure, lng病人ID)
    
    Do Until rsCfsc.EOF
        lngPatiID = rsCfsc!病人ID
        lngPageID = rsCfsc!主页ID
        
        str费用日期 = rsCfsc!费用日期
        
        gstrSQL = "Select * from 保险帐户 where 险类=[1] And 病人ID=[2]"
        Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", intinsure, lngPatiID)
        
        str流水号 = rsMzxnjs!顺序号
        
        gstrSQL = "Select A.收费细目ID,decode(nvl(A.医嘱序号,-99),-99,to_char(A.发生时间,'YYYY-MM-DD'),to_char(A.登记时间,'YYYY-MM-DD')) as 发生时间,max(nvl(床号,0)) as 床号,decode(round(sum(nvl(A.付数,1)*nvl(A.数次,1))),0,1,round(sum(nvl(A.付数,1)*nvl(A.数次,1)))) as 数量,sum(nvl(A.实收金额,0)) as 金额," & _
                                  "sum(nvl(A.实收金额,0))/decode(round(sum(nvl(A.付数,1)*nvl(A.数次,1))),0,1,round(sum(nvl(A.付数,1)*nvl(A.数次,1))))as 价格,max(C.编码) as 开单部门 " & _
                          " from 住院费用记录 A,保险帐户 B,部门表 C " & _
                          " where decode(nvl(A.医嘱序号,-99),-99,to_char(A.发生时间,'YYYY-MM-DD'),to_char(A.登记时间,'YYYY-MM-DD'))='" & str费用日期 & _
                                "' And nvl(A.附加标志,0)<>9 And A.记帐费用=1" & _
                                " And A.记录状态=" & 1 & _
                                " And A.病人ID=B.病人ID " & _
                                " and B.险类=" & intinsure & _
                                " and A.病人病区ID=C.ID " & _
                                " ANd A.病人ID=[1]" & _
                                " And A.主页ID=[2]" & _
                                " And nvl(A.实收金额,0)<>0 " & _
                                " Group by A.收费细目ID,decode(nvl(A.医嘱序号,-99),-99,to_char(A.发生时间,'YYYY-MM-DD'),to_char(A.登记时间,'YYYY-MM-DD'))"
        Call WriteBusinessLOG("传明细", gstrSQL, "")
        Set rs记帐明细 = zlDatabase.OpenSQLRecord(gstrSQL, "记帐明细", lngPatiID, lngPageID)
        lngCount = 0
        '1 分配空间
        Call WriteBusinessLOG("准备分配空间", "1024*30", "")
        If tsx_createparams(1024 * 30, 1024 * 30) = -1 Then
            MsgBox "分配内存空间失败!", vbInformation, gstrSysName
            Exit Function
        End If
        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)
        
        If rs记帐明细.EOF Then
            '当天的费用全都冲销完了,就传一个0上去
            bln全部冲销 = True
        End If
        
        Do Until rs记帐明细.EOF
            '2 Beging 为参数赋值
            lngReturn = tsx_setstringparam(P_JGM, lngCount, g常用信息_铜山.医院码)
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCount & "," & g常用信息_铜山.医院码, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCount & "," & g常用信息_铜山.医院码, lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_ZYXH, lngCount, str流水号)
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & lngCount & "," & str流水号, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & lngCount & "," & str流水号, lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_FYYMD, lngCount, Format(rs记帐明细!发生时间, "yyyyMMdd"))
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_FYYMD" & ", " & lngCount & "," & Format(rs记帐明细!发生时间, "yyyyMMdd"), STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_FYYMD" & ", " & lngCount & "," & Format(rs记帐明细!发生时间, "yyyyMMdd"), lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_RYBQ, lngCount, rs记帐明细!开单部门)
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & lngCount & "," & rs记帐明细!开单部门, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & lngCount & "," & rs记帐明细!开单部门, lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_RYCWH, lngCount, Nvl(rs记帐明细!床号, 0))
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & lngCount & "," & Nvl(rs记帐明细!床号, 0), STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & lngCount & "," & Nvl(rs记帐明细!床号, 0), lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_JZLB, lngCount, 0) '就诊类别    暂为0
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_JZLB" & ", " & lngCount & "," & 0, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_JZLB" & ", " & lngCount & "," & 0, lngReturn)
            End If
            
            lngReturn = tsx_setstringparam(P_CZRYH, lngCount, g常用信息_铜山.操作员号) ' C10 操作人员号
            If lngReturn = -1 Then
                STRERR = tsx_getlasterr()
                Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH," & lngCount & "," & g常用信息_铜山.操作员号, STRERR)
            Else
                Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH," & lngCount & "," & g常用信息_铜山.操作员号, lngReturn)
            End If
            
            gstrSQL = "Select * from ypzlk where 收费细目ID=" & rs记帐明细!收费细目ID
            Call OpenRecordset_OtherBase(rsMzxnjs, "ypzlk", gstrSQL, gcn铜山县)
            If rsMzxnjs.EOF = False Then
                lngReturn = tsx_setstringparam(P_ZBM, lngCount, rsMzxnjs!自编码) 'C20 自编码  费用明细自编码
                If lngReturn = -1 Then
                    STRERR = tsx_getlasterr()
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rsMzxnjs!自编码, STRERR)
                Else
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rsMzxnjs!自编码, lngReturn)
                End If
            
            Else
                lngReturn = tsx_setstringparam(P_ZBM, lngCount, rs记帐明细!收费细目ID) 'C20 自编码  费用明细自编码
                
                If lngReturn = -1 Then
                    STRERR = tsx_getlasterr()
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rs记帐明细!收费细目ID, STRERR)
                Else
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rs记帐明细!收费细目ID, lngReturn)
                End If
            End If
            gstrSQL = "select A.*,B.药品类型 from 收费细目 A,药品信息 B,药品目录 C " & _
                      " where A.id=C.药品ID(+) and C.药名ID=B.药名ID(+) And A.id=[1]"
            Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "取项目类别", CLng(rs记帐明细!收费细目ID))
            Select Case rsMzxnjs!类别s
                Case "5", "6", "7"
                    str药品 = "0"
                Case Else
                    str药品 = "1"
            End Select
            lngReturn = tsx_setstringparam(P_LB, lngCount, str药品)
            Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & lngCount & "," & str药品, lngReturn)
            If lngReturn = -1 Then
                Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
            End If
            
            lngReturn = tsx_setdoubleparam(P_JG, lngCount, rs记帐明细!价格)
            Call WriteBusinessLOG("tsx_setdoubleparam", "P_JG" & ", " & lngCount & "," & rs记帐明细!价格, lngReturn)
            If lngReturn = -1 Then
                Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
            End If
             
            lngReturn = tsx_setlongparam(P_SL, lngCount, rs记帐明细!数量)
            Call WriteBusinessLOG("tsx_setlongparam", "P_SL" & ", " & lngCount & "," & rs记帐明细!数量, lngReturn)
            If lngReturn = -1 Then
                Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
            End If
            lngReturn = tsx_setstringparam(P_CZYMD, lngCount, Format(rs记帐明细!发生时间, "yyyyMMdd") & Format(Now(), "HHmmss"))
            Call WriteBusinessLOG("tsx_setstringparam", "P_CZYMD" & ", " & lngCount & "," & Format(rs记帐明细!发生时间, "yyyyMMdd") & Format(Now(), "HHmmss"), lngReturn)
            If lngReturn = -1 Then
                Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
            End If
            
            lngCount = lngCount + 1
            rs记帐明细.MoveNext
        Loop
        '2 End 为参数赋值
        
        If bln全部冲销 = True Then
            gstrSQL = "Select A.收费细目ID,decode(nvl(A.医嘱序号,-99),-99,to_char(A.发生时间,'YYYY-MM-DD'),to_char(A.登记时间,'YYYY-MM-DD')) as 发生时间,max(nvl(床号,0)) as 床号,sum(nvl(A.付数,1)*nvl(A.数次,0)) as 数量,sum(nvl(A.实收金额,0)) as 金额," & _
                              "sum(nvl(A.实收金额,0))/round(sum((nvl(A.付数,1)*nvl(A.数次,0)))) as 价格,max(C.编码) as 开单部门 " & _
                      " from 住院费用记录 A,保险帐户 B,部门表 C " & _
                      " where decode(nvl(A.医嘱序号,-99),-99,to_char(A.发生时间,'YYYY-MM-DD'),to_char(A.登记时间,'YYYY-MM-DD'))='" & str费用日期 & _
                            "' And nvl(A.附加标志,0)<>9 And A.记录性质=" & int性质 & _
                            " And A.记录状态=" & int状态 & _
                            " And A.病人ID=B.病人ID " & _
                            " and B.险类=" & intinsure & _
                            " and A.病人病区ID=C.ID " & _
                            " ANd A.病人ID=" & lngPatiID & _
                            " And A.主页ID=" & lngPageID & _
                            " And nvl(A.实收金额,0)<>0 " & _
                            " Group by A.收费细目ID,decode(nvl(A.医嘱序号,-99),-99,to_char(A.发生时间,'YYYY-MM-DD'),to_char(A.登记时间,'YYYY-MM-DD'))"
            Call WriteBusinessLOG("传零明细", gstrSQL, "")
                        
'            Call OpenRecordset(rs记帐明细, "记帐明细", gstrSQL)
            
                lngReturn = tsx_setstringparam(P_JGM, lngCount, g常用信息_铜山.医院码)
                Call WriteBusinessLOG("tsx_setstringparam", "P_JGM" & ", " & lngCount & "," & g常用信息_铜山.医院码, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                
                lngReturn = tsx_setstringparam(P_ZYXH, lngCount, str流水号)
                Call WriteBusinessLOG("tsx_setstringparam", "P_ZYXH" & ", " & lngCount & "," & str流水号, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                
                lngReturn = tsx_setstringparam(P_FYYMD, lngCount, Replace(str费用日期, "-", ""))
                Call WriteBusinessLOG("tsx_setstringparam", "P_FYYMD" & ", " & lngCount & "," & Replace(str费用日期, "-", ""), lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                
                gstrSQL = "Select B.编码 from 病案主页 A,部门表 B where A.病人ID=[1] And A.主页ID=[2] And A.入院病区ID=B.ID"
                Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "zlypk", lngPatiID, lngPageID)
                If rsMzxnjs.EOF = False Then
                    str部门 = rsMzxnjs!编码
                    lngReturn = tsx_setstringparam(P_RYBQ, lngCount, str部门)
                    Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & lngCount & "," & str部门, lngReturn)
                    If lngReturn = -1 Then
                        Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                    End If
                Else
                    Call WriteBusinessLOG("tsx_setstringparam", "P_RYBQ" & ", " & lngCount & "," & "未找到部门", lngReturn)
                    If lngReturn = -1 Then
                        Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                    End If
                End If
                 
                lngReturn = tsx_setstringparam(P_RYCWH, lngCount, 0)
                Call WriteBusinessLOG("tsx_setstringparam", "P_RYCWH" & ", " & lngCount & "," & 0, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                 
                lngReturn = tsx_setstringparam(P_JZLB, lngCount, 0) '就诊类别    暂为0
                Call WriteBusinessLOG("tsx_setstringparam", "P_JZLB" & ", " & lngCount & "," & 0, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                
                lngReturn = tsx_setstringparam(P_CZRYH, lngCount, g常用信息_铜山.操作员号) ' C10 操作人员号
                Call WriteBusinessLOG("tsx_setstringparam", "P_CZRYH," & lngCount & "," & g常用信息_铜山.操作员号, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                
                gstrSQL = "Select * from ypzlk where rownum=1"
                Call OpenRecordset_OtherBase(rsMzxnjs, "zlypk", gstrSQL, gcn铜山县)
                If rsMzxnjs.EOF = False Then
                    lngReturn = tsx_setstringparam(P_ZBM, lngCount, rsMzxnjs!自编码) 'C20 自编码  费用明细自编码
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & rsMzxnjs!自编码, lngReturn)
                    If lngReturn = -1 Then
                        Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                    End If
                Else
                    Call WriteBusinessLOG("tsx_setstringparam", "P_ZBM" & ", " & lngCount & "," & "未找到自编码", lngReturn)
                    If lngReturn = -1 Then
                        Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                    End If
                End If
                
                str药品 = "0"
                
                lngReturn = tsx_setstringparam(P_LB, lngCount, str药品)
                Call WriteBusinessLOG("tsx_setstringparam", "P_JG" & ", " & lngCount & "," & str药品, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                
                lngReturn = tsx_setdoubleparam(P_JG, lngCount, 0)
                Call WriteBusinessLOG("tsx_setdoubleparam", "P_LB" & ", " & lngCount & "," & 0, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                 
                lngReturn = tsx_setlongparam(P_SL, lngCount, 0)
                Call WriteBusinessLOG("tsx_setlongparam", "P_SL" & ", " & lngCount & "," & 0, lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
                lngReturn = tsx_setstringparam(P_CZYMD, lngCount, Replace(str费用日期, "-", "") & Format(Now(), "HHmmss"))
                Call WriteBusinessLOG("tsx_setstringparam", "P_CZYMD" & ", " & lngCount & "," & Replace(str费用日期, "-", "") & Format(Now(), "HHmmss"), lngReturn)
                If lngReturn = -1 Then
                    Call WriteBusinessLOG("错误信息", "", tsx_getlasterr())
                End If
      End If

        '3 调用接口
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
            '4 取返回值
            
            '无反回值
                ''>>begin 更新HIS的上传标志
                
                If lngReturn = -1 Then
                    If bln全部冲销 = False Then
                        MsgBox "上传明细(" & rsMzxnjs!类别 & "_" & rsMzxnjs!编码 & ")" & rsMzxnjs!名称 & "失败。" & vbCrLf & Trim(tsx_getlasterr), vbInformation, gstrSysName
                    Else
                        Call WriteBusinessLOG("全部冲销", "记录零费用", "-1")
                    End If
                Else
                    gstrSQL = "Select distinct A.ID " & _
                      " from 住院费用记录 A,保险帐户 B,部门表 C " & _
                      " where decode(nvl(A.医嘱序号,-99),-99,to_char(A.发生时间,'YYYY-MM-DD'),to_char(A.登记时间,'YYYY-MM-DD'))='" & str费用日期 & _
                            "'And nvl(A.附加标志,0)<>9 And a.记帐费用=1" & _
                            " And A.NO='" & str单据号 & "' " & _
                            " and A.记录性质=" & int性质 & _
                            " And A.记录状态=" & int状态 & _
                            " And A.病人ID=B.病人ID " & _
                            " and B.险类=" & intinsure & _
                            " and A.病人病区ID=C.ID " & _
                            " ANd A.病人ID=[1]" & _
                            " And A.主页ID=[2]" & _
                            " And nvl(A.实收金额,0)<>0 "
                    Call WriteBusinessLOG("填写上传标志的记录", gstrSQL, "")
                    Set rsMzxnjs = zlDatabase.OpenSQLRecord(gstrSQL, "已上传明细的ID", lngPatiID, lngPageID)
                    Do Until rsMzxnjs.EOF
                        gstrSQL = "ZL_病人记帐记录_上传(" & rsMzxnjs!ID & "," & 0 & ",NULL)"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新医保字段")
                        
                        '>>Beging 写入XYBMX表中,用于打印一日清单
                        
                        '>>>Beging 取xybmx里的状态
                        gstrSQL = "Select * from 保险帐户 where 险类=[1] And 病人ID=[2]"
                        Call WriteBusinessLOG("取xybmx里的状态", gstrSQL, "")
                        Set rsXybmx = zlDatabase.OpenSQLRecord(gstrSQL, "保险帐户", TYPE_铜山县, lngPatiID)
                        str个人编号 = rsXybmx!医保号
                        
                        '1 分配空间
                        If tsx_createparams(1024, 1024) = -1 Then
                            MsgBox "分配内存空间失败!", vbInformation, gstrSysName
                            Exit Function
                        End If
                        Call WriteBusinessLOG("tsx_createparams", "1024,1204", lngReturn)

                         '2 为参数赋值
                        
                         lngReturn = tsx_setstringparam(P_TBR, 0, str个人编号) '    C9  个人编号    参保人员个人编号
                         Call WriteBusinessLOG("tsx_setstringparam", "P_TBR" & ", " & 0 & "," & str个人编号, lngReturn)
                         
                         lngReturn = tsx_setstringparam(P_LB, 0, 0) '
                         Call WriteBusinessLOG("tsx_setstringparam", "P_LB" & ", " & 0 & "," & 0, lngReturn)
                                  
                        '3 调用接口
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
                        '4 取返回值
                             lngReturn = tsx_getdoubleparam(P_GRQK, 0, dbl医保欠款)
                             Call WriteBusinessLOG("tsx_getstringparam", "P_RYH, 0, " & dbl医保欠款, lngReturn)
                             
                        End If
                        '5 销毁已分配空间
                        lngReturn = tsx_destroyparams()
                        Call WriteBusinessLOG("tsx_destroyparams", "", lngReturn)

                         
                         '>>>End 取xybmx里的状态
                        
                        gstrSQL = "Select b.名称, b.计算单位 As 单位, Nvl(a.实收金额, 0) As 实收金额," & _
                                  "Decode(Nvl(c.附注, '自费'), '甲类', 0, '乙类', 0.2, 1) * Nvl(a.实收金额, 0) As 自费金额 ," & _
                                  "Decode(Round(Nvl(a.付数, 1) * Nvl(a.数次, 1)), 0, 1, Round(Nvl(a.付数, 1) * Nvl(a.数次, 1))) As 数量," & _
                                  "Nvl(a.实收金额, 0) /Decode(Round(Nvl(a.付数, 1) * Nvl(a.数次, 1)), 0, 1, Round(Nvl(a.付数, 1) * Nvl(a.数次, 1))) As 价格," & _
                                  "To_Char(发生时间, 'YYYY-MM-DD HH24:MI:SS') As 日期" & _
                                  " From 住院费用记录 A,收费细目 B,(Select * From 保险支付项目 where 险类=[1]) C " & _
                                  " Where  Nvl(a.实收金额, 0) <>0 and A.收费细目ID=B.ID And A.收费细目ID=C.收费细目ID(+) and A.ID=[2]"
                        Call WriteBusinessLOG("查费用明细准备写xybmx", gstrSQL, "")
                        Set rs费用记录 = zlDatabase.OpenSQLRecord(gstrSQL, "费用记录", TYPE_铜山县, CLng(rsMzxnjs!ID))
                        
                        If rs费用记录.EOF = False Then
                            gstrSQL = "Select * from XYBMX Where 记录ID=" & rsMzxnjs!ID
                            Call WriteBusinessLOG("查xybmx", gstrSQL, "")
                            Call OpenRecordset_OtherBase(rsXybmx, "费用记录", , gcn铜山县)
                            
                            If rsXybmx.EOF Then
                                gstrSQL = "Insert into XYBMX(项目名称,单位,单价,数量,合计金额,自付金额,病人id,主页id,日期,记录id,医保状态) values('" & _
                                        rs费用记录!名称 & "','" & rs费用记录!单位 & "'," & rs费用记录!价格 & "," & rs费用记录!数量 & "," & rs费用记录!实收金额 & "," & rs费用记录!自费金额 & "," & lngPatiID & "," & lngPageID & "," & _
                                        "to_date('" & Format(rs费用记录!日期, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & rsMzxnjs!ID & "," & dbl医保欠款 & ")"
                                Call WriteBusinessLOG("写xybmx", gstrSQL, "")
                                gcn铜山县.Execute gstrSQL
                            End If
                        
                        End If
                        
                        '>>End 写入XYBMX表中,用于打印一日清单
                        
                        rsMzxnjs.MoveNext
                    Loop
                End If 'lngReturn = -1
                '>>End 更新HIS的上传标志
                
            End If 'tsx_jkcall("ZYMX_SC")
                      '5 销毁已分配空间
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

Public Function 病人变动_铜山县(lngPatiID As Long, lngPageID As Long, ByVal intinsure As Integer) As Boolean

'*****************************************************************************
'调用者　　　　　　：住院信息调整调用
'功能说明　　　　　：在院病人的床位变动,转科,医生变化及诊断情况等相关信息变化时调用此接口
'步骤说明　　　　　：1、提取病人变动信息
'　　　　　　　　　　2、仅上传本医保的病人变动情况
'　　　　　　　　　　3、根据接口要求,调用相应函数上传变动情况
'调用过程清单及说明：
'　　【无】
''*****************************************************************************
    '//TODO:病人变动
    '以下是打上传标记的示范代码（有两种不同的方式，可根据你的需要使用）
    '注意事项:(无)
    病人变动_铜山县 = True
End Function

Public Function 帐户转预交_铜山县(lng预交ID As Long, curMoney As Currency, ByVal intinsure As Integer) As Boolean
'*****************************************************************************
'调用者　　　　　　：
'功能说明　　　　　：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'步骤说明　　　　　：1、提取病人变动信息
'　　　　　　　　　　2、仅上传本医保的病人变动情况
'　　　　　　　　　　3、根据接口要求,调用相应函数上传变动情况
'调用过程清单及说明：
'　　【　　　】医保部件关闭
''*****************************************************************************
    '//TODO:帐户转预交
    '
    '注意事项:(无)

End Function

Public Function 帐户转预交冲销_铜山县(lng预交ID As Long, ByVal intinsure As Integer) As Boolean
'*****************************************************************************
'调用者　　　　　　：
'功能说明　　　　　：将需要从预交款转入个人帐户余额的数据记录发送医保前置服务器确认；
'步骤说明　　　　　：1、提取病人变动信息
'　　　　　　　　　　2、仅上传本医保的病人变动情况
'　　　　　　　　　　3、根据接口要求,调用相应函数上传变动情况
'调用过程清单及说明：
'　　【　　　】医保部件关闭
''*****************************************************************************
    '//TODO:帐户转预交撤销
    '
    '注意事项:(无)
End Function

Public Function 病种选择_铜山县(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal intinsure As Integer)
'*****************************************************************************
'调用者　　　　　　：被clsInsure 的  ChooseDisease  过程调用
'功能说明　　　　　：选择病人的出院病种
'调用过程清单及说明：
'　　【　　　】医保部件关闭
''*****************************************************************************
'//TODO:病种选择，在医保前台程序中已有此功能
    
End Function

Public Function 医保参数设置_铜山县(ByVal cap业务 As 医院业务) As Boolean
'*****************************************************************************
'调用者　　　　　　：被clsInsure 的  GetCapability  过程调用
'功能说明　　　　　：判断后来补充的一些业务在不同的医保软件是否得到支持
'调用过程清单及说明：
'　　【无】
''*****************************************************************************
'TODO:参数设置
    Select Case cap业务
    
        Case support门诊退费, _
             support记帐上传, _
             support记帐作废上传, _
             support医嘱上传, _
             support记帐完成后上传, _
             support门诊必须传递明细, _
             support出院结算必须出院, _
             support未结清出院, _
             support结算使用个人帐户, _
             support必须录入入出诊断, _
             support撤销出院, _
             support结帐退个人帐户, _
             support出院病人结算作废, _
             support门诊预算, _
             support允许部份冲销单据, _
             support负数记帐
             医保参数设置_铜山县 = True
             
       Case support门诊结算作废, support住院结算作废
             医保参数设置_铜山县 = True
'
    End Select
    
End Function

Public Function 取消就诊_铜山县(ByVal bytType As Byte, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
'*****************************************************************************
'调用者　　　　　　：被clsInsure 的  IdentifyCancel  过程调用
'功能说明　　　　　：用于门诊医保病人身份验证成功后，取消就诊的情况
'调用过程清单及说明：
'　　【无】
''*****************************************************************************
'TODO:取消就诊
    取消就诊_铜山县 = True
    
End Function


Public Function tsx_getlasterr() As String
    Dim lngRetu As Long, STRERR As String
    
    STRERR = Space(512)
    lngRetu = tsx_getlasterr2(STRERR)
    tsx_getlasterr = STRERR
    
End Function

Public Function 医保项目_铜山县(病人ID As Long, 收费细目ID As Long, 金额 As Currency, _
    ByVal bln门诊 As Boolean, Optional ByVal intinsure As Integer) As String
'*****************************************************************************
'调用者　　　　　　：被clsInsure 的  GetItemInsure  过程调用
'功能说明　　　　　：主要用于前台显示本医保的费用类型甲乙类
'调用过程清单及说明：
'　　【无】
''*****************************************************************************
'TODO:医保项目信息
    Dim rsYbxm As New ADODB.Recordset
    gstrSQL = "Select * from ypzlk where 收费细目ID=" & 收费细目ID
    Call OpenRecordset_OtherBase(rsYbxm, "ypzlk", , gcn铜山县)
    If rsYbxm.EOF = False Then
        医保项目_铜山县 = rsYbxm!支付类别
    Else
        医保项目_铜山县 = "未对码"
    End If
    Set rsYbxm = Nothing
    
End Function

Public Function 医保信息_铜山县(ByVal lngItemID As Long, Optional intType As Integer = 0) As String
'*****************************************************************************
'调用者　　　　　　：被clsInsure 的  GetItemInfo  过程调用
'功能说明　　　　　：主要用于前台显示本医保的费用类型甲乙类
'调用过程清单及说明：
'　　【无】
''*****************************************************************************
    Dim rsYbxm As New ADODB.Recordset
    'WriteBusinessLOG "调进了 医保信息_铜山县", lngItemID, intType
    If intType = 0 Then '医嘱调用则提示
        gstrSQL = "Select * from ypzlk where 收费细目ID=" & lngItemID
        Call OpenRecordset_OtherBase(rsYbxm, "ypzlk", , gcn铜山县)
        If rsYbxm.EOF = False Then
            医保信息_铜山县 = rsYbxm!支付类别
        Else
            医保信息_铜山县 = "未对码"
        End If
        Set rsYbxm = Nothing
        If 医保信息_铜山县 <> "" Then MsgBox "该项目的医保类别为“" & 医保信息_铜山县 & "”", vbInformation, gstrSysName
        医保信息_铜山县 = ""
    End If
End Function



