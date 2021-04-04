Attribute VB_Name = "mdlDocument"
    '======================================================================================================================
    '一、IC卡结构说明
    '┌────┬──┬────┬────┬────────────────┐
    '│字段    │长度│地址长度│地址    │备注                            │
    '├────┼──┼────┼────┼────────────────┤
    '│员工编码│8   │4       │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│身份证号│18  │9       │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│姓名    │8   │8       │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│性别    │1   │1       │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│出生年月│8   │4       │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│工作年月│8   │4       │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│工作单位│3   │3       │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│人员性质│1   │0.5     │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│用工性质│1   │0.5     │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│住院次数│6   │        │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│年累计额│8   │        │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│        │    │        │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│        │    │        │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│        │    │        │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│        │    │        │        │                                │
    '├────┼──┼────┼────┼────────────────┤
    '│住院号  │12  │        │        │最近5次住院记录                 │
    '├────┼──┼────┼────┤                                │
    '│入院日期│8   │        │        │                                │
    '├────┼──┼────┼────┤                                │
    '│出院日期│8   │        │        │                                │
    '├────┼──┼────┼────┤                                │
    '│费用合计│8   │        │        │                                │
    '├────┼──┼────┼────┤                                │
    '│报销金额│8   │        │        │                                │
    '├────┼──┼────┼────┤                                │
    '│自付金额│8   │        │        │                                │
    '└────┴──┴────┴────┴────────────────┘
    
    '======================================================================================================================
    '二、VFP示例代码
    
    '*程 序 名： Func_ICC.prg
    '*编制日期： 2002年9月20日
    '*功    能：  定义IC卡操作的有关函数
    '
    '**(0) declare_IC()        IC卡函数说明
    '**(1) init_card()         初始化卡机，必须开电源
    '**(2) close_card()        关闭卡机
    '**(3) check_card()        初始化IC卡，必须有合法的卡
    '
    '**(4) rd_kxh()            读卡序号
    '**(5) wr_kxh(kxh,fkcs)    写卡序号
    '
    '**(6) rd_sfz()            读身份证
    '**(7) wr_sfz(rybm,sfz)    写身份证
    '
    '**(8) rd_geren()          读个人信息
    '**(9) wr_geren()          写个人信息
    '**(10) ts_geren()         更新提示个人信息
    '
    '**(11) rd_nlje()          读住院年累计额
    '**(12) wr_nlje()          写住院年累计额
    '
    '**(13) rd_zyk()           读住院记录
    '**(14) wr_zyk()           写住院记录
    '**(15) init_zypiont       初始化住院记录指针
    '
    '**(16) check_hmd()        检验黑名单
    '**(17) check_rybm()       检查IC卡与数据记录是否是同一人
    '
    '*  =num_str()             数量转换为8位字符
    '*  =str_num()             8位字符转换为数量
    '*  =str_add0()            数字字符串前面加零
    '*  =STR_INC0()            数字字符串前面去零
    '
    '
    '****(0) IC卡函数说明 *********************************************
    'FUNC declare_IC
    '
    '*m_dll=p_server+"iccdll\mwic_32.dll"
    '
    'declare integer auto_init in mwic_32.dll short port,integer baud
    'declare short ic_exit in mwic_32.dll  integer icdev
    'declare short get_status in mwic_32.dll integer icdev,integer @status
    '
    'declare short chk_4442  in mwic_32.dll integer icdev
    'declare short srd_4442  in mwic_32.dll integer icdev, short offset,short len,string @buffer
    
    'declare short swr_4442  in mwic_32.dll integer icdev, short offset,short len,string @buffer
    
    'declare short csc_4442   in mwic_32.dll integer icdev, short len,string @buffer
    'declare short wsc_4442   in mwic_32.dll integer icdev, short len,string @buffer
    'declare short rsc_4442   in mwic_32.dll integer icdev, short len,string @buffer
    'declare short rsct_4442   in mwic_32.dll integer icdev, short len,string @buffer
    '
    'declare short ic_encrypt in mwic_32.dll  string @buffer, string buffer1,short len,string @buffer2
    'declare short ic_decrypt in mwic_32.dll  string @buffer, string @buffer1,short len,string @buffer2
    '
    'declare short asc_hex in mwic_32.dll string @buffer, string @buffer1,integer  len
    'declare short hex_asc in mwic_32.dll string @buffer, string @buffer1,integer len
    '
    'endfunc
    '
    '****(1) 初始化卡*********************************************
    'FUNC init_card
    '
    'declare_IC()   &&说明IC卡库函数
    '
    'm_com=0        &&COM1
    'm_baud=9600    &&波特率
    '
    'p_icdev = auto_init(m_com, m_baud)
    '
    'if p_icdev<0
    '   =messagebox("初始化IC卡错误！"+chr(13)+chr(13)+"请检查是否打开读卡机 ！",16,"错误...")
    '   return .F.
    'End If
    '
    'return .T.
    'endfunc
    '
    '****(2) 关闭卡机*******************************************
    'Function close_card()
    '
    '    if p_icdev<=0
    '*      messagebox("IC卡卡机没准备好 或 没初始化 ,不能关闭IC卡机！",16,"错误")
    '       return .F.
    '    End If
    '
    '    ic_exit (p_icdev)
    '    p_icdve = 0
    '
    'Return
    '
    '
    '****(3) 校验卡*********************************************
    'Function check_card()
    'local m.st
    '
    'if p_icdev<=0
    '   messagebox("IC卡卡机没准备好 或 没初始化 ！",16,"错误")
    '   return .F.
    'End If
    '
    'm_status=0            &&测试卡座是否有卡
    'st=get_status(p_icdev,@m_status)
    'if st<>0
    '    messagebox("没有插入IC卡！",16,"错误...")
    '    return .f.
    'End If
    '
    'st=chk_4442(p_icdev)  &&检查是不是4442卡
    'if st<>0
    '    messagebox("卡类型不对！",16,"错误...")
    '     return .f.
    'End If
    '
    'm_passwd0=space(3)  &&初始密码
    'm_passwd1=space(3)  &&本系统密码
    'st=asc_hex("FFFFFF",@m_passwd0,3)
    'st=asc_hex("995188",@m_passwd1,3)
    '
    'st = csc_4442(p_icdev, 3, m_passwd1)
    'if st <> 0
    '    st=csc_4442(p_icdev,3,m_passwd0)  &&校对原始密码
    '    if st=0
    '        st=wsc_4442(p_icdev,3,m_passwd1)  &&写新密码
    '        if st<0
    '            =messagebox("写卡密码错！",16,"错误...")
    '            return .F.
    '        End If
    '    Else
    '        =messagebox("非本系统卡，请与系统提供商联系！",16,"错误...")
    '        return .F.
    '    End If
    'End If
    '
    'return .T.
    '
    'endfunc
    '
    '****(4) 读卡序号 *********************************************
    'Function rd_kxh()
    '
    'local st, m_str1,m_return,m_offset,m_len
    '
    'm_offset=0x1b          &&地址
    'm_len=3+1               &&长度
    'm_str1=space(m_len)     &&字符串变量
    'm_return=space(m_len*2) &&返回字符串
    '
    'st=srd_4442(p_icdev,m_offset,m_len,@m_str1)
    'if st<0
    '   =messagebox('读数据错(卡序号)！',16,'错误...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_str1,@m_return,m_len)
    '
    'return m_return
    '
    '****(5) 写卡序号 *********************************************
    'Function wr_kxh()
    'PARA m_kxh, m_fkcs
    '
    'local st, m_str1,m_offset,m_len
    '
    'm_offset=0x1b       &&地址
    'm_len=3+1           &&长度
    'm_str1=space(m_len) &&字符串变量
    '
    'm_kxh = STR_add0(m_kxh, 6)
    'm_fkcs = STR_add0(m_fkcs, 2)
    '
    'st=asc_hex(m_kxh + m_fkcs , @m_str1 , m_len)
    '
    'st=swr_4442(p_icdev , m_offset , m_len , @m_str1)
    '
    '
    'if st < 0
    '   =messagebox('写数据错(卡序号)！',16,'错误...')
    '   return .F.
    'End If
    '
    'return .T.
    '
    '****(6) 读人员编码 *********************************************
    'Function rd_sfz()
    '
    'local st, m_str1,m_return,m_offset,m_len
    '
    'm_offset=0x08           &&地址
    'm_len=4+9               &&长度
    'm_str1=space(m_len)     &&字符串变量
    'm_return=space(m_len*2) &&返回字符串
    '
    'st=srd_4442(p_icdev,m_offset,m_len,@m_str1)
    'if st<0
    '   =messagebox('读数据错(人员编码)！',16,'错误...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_str1,@m_return,m_len)
    '
    'return m_return
    '
    '****(7) 写人员编码 *********************************************
    'Function wr_sfz()
    'PARA m_rybm, m_sfz
    '
    'local st, m_str1,m_offset,m_len
    '
    'm_offset=0x08      &&地址
    'm_len=4+9          &&长度
    'm_str1=space(m_len) &&字符串变量
    '
    'm_rybm = STR_add0(m_rybm, 8)
    'm_sfz = STR_add0(m_sfz, 18)
    '
    'st=asc_hex(m_rybm + m_sfz , @m_str1 , m_len)
    '
    'st=swr_4442(p_icdev , m_offset , m_len , @m_str1)
    '
    '
    'if st < 0
    '   =messagebox('写数据错(人员编码)！',16,'错误...')
    '   return .F.
    'End If
    '
    'return .T.
    '
    '
    '**** （8）读个人信息*********************************************
    '**姓名 性别 血型 出生年月 工作年月 工作单位 工作岗位 人员性质 用工性质
    'FUNC rd_geren
    'local st, m_str1,m_str2,m_return,m_offset,m_len
    '
    'm_offset=0x20              &&地址
    'm_len=24                   &&长度  8+2+2+4+4+1.5+1.5+0.5+0.5
    'm_len2=12                  &&转换数据长度
    'm_str1=space(m_len)        &&字符串变量
    'm_str2 = Space(m_len2 * 2)
    'm_return=space(m_len*2-12) &&返回字符串  姓名 性别 血型 不能压缩
    '
    'st=srd_4442(p_icdev , m_offset , m_len , @m_str1)
    'if m.st<0
    '   messagebox("读数据错(个人信息)！",16,"错误...")
    '   Return ''
    'End If
    
    'm_return=substr(m_str1,1,12)     &&姓名 性别 血型
    '
    'm_str1=substr(m_str1,13)     &&其它
    '
    'st=hex_asc(@m_str1,@m_str2,m_len2)
    '
    'm_return = alltrim(m_return) + m_str2
    '
    'return m_return
    '
    '****(9) 写个人信息*********************************************
    '**姓名 性别 血型 出生年月 工作年月 工作单位 工作岗位 人员性质 用工性质
    'FUNC wr_geren
    'PARA m_xm, m_xb, m_blood, m_csny, m_gzny, m_gzdw, m_gzgw, m_ryxz, m_ygxz
    '
    'local st, m_str1,m_offset,m_len
    '
    'm_offset=0x20       &&地址
    'm_len=24            &&长度 8+2+2+4+4+1.5+1.5+0.5+0.5
    'm_str1=space(m_len) &&字符串变量
    '
    'm_xm=str_add0(m_xm,8)         &&姓名
    'm_blood=str_add0(m_blood,2)   &&血型
    'm_xb=str_add0(m_xb,2)         &&性别
    '
    'if empty(m_csny)          &&出生日期
    '    m_csny='19000101'
    'Else
    '    m_csny = dtos(m_csny)
    'End If
    '
    'if empty(m_gzny)          &&工作日期
    '    m_gzny='19000101'
    'Else
    '    m_gzny = dtos(m_gzny)
    'End If
    '
    'm_gzdw=str_add0(m_gzdw,3)     &&单位编码
    'm_gzgw=str_add0(m_gzgw,3)     &&岗位编码
    'm_ryxz=str_add0(m_ryxz,1)     &&人员性质
    'm_ygxz=str_add0(m_ygxz,1)     &&人员性质
    '
    'st=asc_hex(m_csny + m_gzny + m_gzdw + m_gzgw + m_ryxz + m_ygxz , @m_str1 , m_len)
    '
    'm_str1=m_xm + m_xb +  m_blood + m_str1  &&这三个字段不能压缩
    '
    'st=swr_4442(p_icdev , m_offset , m_len , @m_str1)
    '
    '
    'if st < 0
    '   =messagebox('写数据错(个人信息)！',16,'错误...')
    '   return .F.
    'End If
    '
    'return .T.
    '
    '****(10) 更新提示个人信息 *********************************************
    'Function ts_geren()
    '
    '    m_str=rd_sfz()  && 身份证
    '    if empty(m_str)
    '        Return
    '    End If
    '
    '    m_rybm = STR_INC0(substr(m_str, 1, 8))
    '    m_sfzh = STR_INC0(substr(m_str, 9, 18))
    '
    '    m_str=rd_geren()  && 个人信息
    '    if empty(m_str)
    '        Return
    '    End If
    '
    '    m_xm = STR_INC0(substr(m_str, 1, 8))
    '    m_xb = substr(m_str, 9, 2)
    '    m_blood = STR_INC0(substr(m_str, 11, 2))
    '    m_csny = stod(substr(m_str, 13, 8))
    '    m_gzny = stod(substr(m_str, 21, 8))
    '    m_gzdw = substr(m_str, 29, 3)
    '    m_gzgw = substr(m_str, 32, 3)
    '    m_ryxz = substr(m_str, 35, 1)
    '    m_ygxz = substr(m_str, 36, 1)
    '
    '    if !used("rybm")
    '        use rybm in 0 share
    '        openflag=.F.
    '    Else
    '        openflag=.T.
    '    End If
    '
    '    select rybm
    '    locate for rybm=m_rybm
    '    if found()
    '        if m_xm<>alltrim(rybm.xm) or m_csny<> rybm.csny or m_gzny<>rybm.gzny  or ;
    '            m_gzgw<>rybm.gzgw or m_ryxz<>alltrim(rybm.ryxz) or m_sfzh<>alltrim(rybm.sfzh)
    '
    '            m_mess = "此人的员工档案信息已经更改，请到发卡单位更新卡片信息！"
    '            messagebox(m_mess,64)
    '
    '        End If
    '    End If
    '
    '    if !openflag
    '        select rybm
    '        use
    '    End If
    '
    'endfunc
    '
    '****(11) 读年累计额 *********************************************
    '** 住院次数 年累计额
    'Function rd_nlje()
    '
    'local st, m_str1,m_return,m_offset,m_len
    '
    'm_offset=0x38           &&地址
    'm_len=7                 &&长度 3+4
    'm_str1=space(m_len)     &&字符串变量
    'm_return=space(m_len*2) &&返回字符串
    '
    'st=srd_4442(p_icdev,m_offset,m_len,@m_str1)
    'if st<0
    '   =messagebox('读数据错(住院次数)！',16,'错误...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_str1,@m_return,m_len)
    '
    'return m_return
    '
    '****(12) 写年累计额 *********************************************
    '** 住院次数 年累计额
    'Function wr_nlje()
    'PARA m_zycs, m_nlje
    '
    'local st, m_str1,m_offset,m_len
    '
    'm_offset=0x38       &&地址
    'm_len=7             &&长度 3+4
    'm_str1=space(m_len) &&字符串变量
    '
    'm_zycs = STR_add0(m_zycs, 6)
    'm_nlje = num_str(m_nlje, 8)
    '
    'st=asc_hex(m_zycs + m_nlje , @m_str1 , m_len)
    '
    'st=swr_4442(p_icdev , m_offset , m_len , @m_str1)
    '
    '
    'if st < 0
    '   =messagebox('写数据错(住院次数)！',16,'错误...')
    '   return .F.
    'End If
    '
    'return .T.
    '
    '
    '****(13) 读住院记录 *********************************************
    '** 凭证号 出院日期 费用合计 报销金额
    'Function rd_zyk()
    'PARA m_recno
    'local st, m_str1,m_return,m_offset,m_len,m_rec,m_rec1,m_recc
    '
    '
    'm_offset1=0x40           &&地址
    'm_offset2=0x60           &&地址
    'm_offset3=0x80           &&地址
    'm_offset4=0xa0           &&地址
    'm_offset5=0xc0           &&地址
    'm_len=16                &&长度 5+3+4+4
    'm_str1=space(m_len)     &&字符串变量
    'm_return=space(m_len*2+4) &&返回字符串
    'm_rec=space(1)            &&记录指针
    'm_rec1 = Space(2)
    '
    'm.st=srd_4442(p_icdev,0x3F,1,@m_rec)  &&读当前记录指针
    'if m.st<0
    '   =messagebox('读数据错(记录指针)！',16,'错误...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_rec,@m_rec1,1)
    'm_recc = Val(substr(m_rec1, 1, 1))
    'm_rec = Val(substr(m_rec1, 2, 1))
    '
    'if m_recno>m_recc  &&起范围
    '    return ""
    'End If
    '
    'if m_rec>5 or m_rec<1
    '   =messagebox('读数据错(记录指针超范围)！',16,'错误...')
    '   Return ''
    'End If
    '
    '
    'm_recno = m_rec - m_recno + 1
    'if m_recno<1
    '    m_recno = m_recno + 5
    'End If
    '
    '
    'do case
    '    Case m_recno = 1
    '        st=srd_4442(p_icdev,m_offset1,m_len,@m_str1)
    '    Case m_recno = 2
    '        st=srd_4442(p_icdev,m_offset2,m_len,@m_str1)
    '    Case m_recno = 3
    '        st=srd_4442(p_icdev,m_offset3,m_len,@m_str1)
    '    Case m_recno = 4
    '        st=srd_4442(p_icdev,m_offset4,m_len,@m_str1)
    '    Case m_recno = 5
    '        st=srd_4442(p_icdev,m_offset5,m_len,@m_str1)
    '    otherwise
    '        st = -1
    'endcase
    '
    'if st<0
    '   =messagebox('读数据错(住院记录)！',16,'错误...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_str1,@m_return,m_len)
    '
    'if substr(m_return,1,2)="FF"  &&还没有写记录
    '    return ""
    'End If
    '
    'm_return = "20" + substr(m_return, 1, 10) + "20" + alltrim(substr(m_return, 11))
    '
    'return m_return
    '
    '****(14) 写住院记录 *********************************************
    '** 凭证号 出院日期 费用合计 报销金额
    'Function wr_zyk()
    'PARA m_pzh, m_cyrq, m_hj, m_bxje
    '
    'local st, m_str1,m_offset,m_len , m_rec , m_rec1 , m_recc
    '
    'm_offset1=0x40           &&地址
    'm_offset2=0x60           &&地址
    'm_offset3=0x80           &&地址
    'm_offset4=0xa0           &&地址
    'm_offset5=0xc0           &&地址
    'm_len=16                &&长度 5+3+4+4
    'm_str1=space(m_len)     &&字符串变量
    'm_rec1=space(2)            &&记录指针
    'm_rec = Space(2)
    '
    'm.st=srd_4442(p_icdev,0x3F,1,@m_rec)  &&读当前记录指针
    'if m.st<0
    '   =messagebox('读数据错(记录指针)！',16,'错误...')
    '   Return ''
    'End If
    'st=hex_asc(@m_rec,@m_rec1,1)
    'm_recc=substr(m_rec1,1,1)  &&现有记录数
    'm_rec =substr(m_rec1,2,1)   &&当前记录数
    '
    'if m_rec>"5" and m_rec<"0"
    '   =messagebox('读数据错(记录指针超范围)！',16,'错误...')
    '   Return ''
    'End If
    '
    'm_rec = Val(m_rec) + 1
    'if m_rec>5
    '    m_rec = 1
    'End If
    '
    'm_recc = Val(m_recc) + 1
    'if m_recc>5
    '    m_recc = 5
    'End If
    '
    'm_pzh=str_add0(substr(m_pzh,3,10),10)  &&凭证号
    '
    'if empty(m_cyrq)
    '    m_cyrq = Date
    'End If
    'm_cyrq=substr(dtos(m_cyrq),3,6)           &&出院日期
    '
    'm_hj=num_str(m_hj,8)        &&合计金额
    'm_bxje=num_str(m_bxje,8)    &&报销金额
    '
    'st=asc_hex(m_pzh + m_cyrq + m_hj + m_bxje , @m_str1 , m_len)
    '
    'do case
    '    Case m_rec = 1
    '        st=swr_4442(p_icdev , m_offset1 , m_len , @m_str1)
    '
    '    Case m_rec = 2
    '        st=swr_4442(p_icdev , m_offset2 , m_len , @m_str1)
    '
    '    Case m_rec = 3
    '        st=swr_4442(p_icdev , m_offset3 , m_len , @m_str1)
    '
    '    Case m_rec = 4
    '        st=swr_4442(p_icdev , m_offset4 , m_len , @m_str1)
    '
    '    Case m_rec = 5
    '        st=swr_4442(p_icdev , m_offset5 , m_len , @m_str1)
    '
    '    otherwise
    '        st = -1
    'endcase
    '
    'if st < 0
    '   =messagebox('写数据错(住院记录)！',16,'错误...')
    '   return .F.
    'End If
    '
    'm_rec = alltrim(Str(m_rec))
    'm_rec1 = alltrim(Str(m_recc)) + m_rec
    'st=asc_hex(@m_rec1,@m_rec,1)
    'st=swr_4442(p_icdev , 0x3F , 1 , @m_rec)  &&写记录指针
    '
    'return .T.
    '
    '****(15) 初始化住院记录指针 *********************************************
    'Function init_zypiont()
    'local m_rec
    '
    '    m_rec = Chr(0)
    '    st=swr_4442(p_icdev , 0x3F , 1 , @m_rec)
    '    if st < 0
    '        =messagebox('写数据错(住院记录指针)！',16,'错误...')
    '        return .F.
    '    End If
    '
    'return .T.
    '
    '****(16) 检查黑名单 *********************************************
    'Function check_hmd()
    '
    'return .T.
    '
    '****(17) 检查IC卡与数据记录是否是同一人 *********************************************
    'Function check_rybm()
    '
    'return .T.
    '
    '
    '********************************************* 1位小数位数值转换为8位字符串 *********************************************
    'FUNC num_str
    'PARA nnum, nstrlen
    'local m_str
    '
    'if nnum<0
    '   messagebox("数量转换为字符不能为负数!",16,"错误")
    '   Return ''
    'End If
    '
    'm_str=str(nnum*10,nstrlen)             &&去掉1位小数点,  前面加空格
    'm_str=strtran(m_str,' ','0')  &&用'0'代替空格
    '
    'return m_str
    '
    '********************************************* 8 字符串 转换为 数值 *********************************************
    'FUNC str_num
    'PARA m_str
    'local m_nnum
    '
    'set decimals to 0
    'm_nnum=val(m_str)/10    &&1位小数点
    '
    'return m_nnum
    '
    '********************************************* 字符串前面加'0' *********************************************
    'FUNC STR_add0
    'PARA cstring, nstrlen
    'local str_return,cstr,nstrlen0,m_len
    '
    'cstring = alltrim(cstring)
    'nstrlen0 = Len(cstring)
    'do case                                   &&实际长度与需要长度比较
    '   case nstrlen0<nstrlen                  &&前面加'0'
    '        m_len = nstrlen - nstrlen0
    '        cstr=space(m_len)+cstring
    '        str_return=strtran(cstr,' ','0')
    '   case nstrlen0=nstrlen                  &&不变
    '        str_return = cstring
    '   otherwise                              &&实际长度过大
    '        str_return = substr(cstring, 1, nstrlen)
    'endcase
    '
    'return str_return
    '
    '********************************************* 去掉字符串前面'0' *********************************************
    'FUNC STR_INC0
    'PARA cstring
    'local str_return,cstr,nstrlen
    '
    'do while .t.
    '    nstrlen = Len(cstring)
    '    if nstrlen=0
    '        exit
    '    End If
    '    cstring=iif(subs(cstring,1,1)='0',subs(cstring,2,nstrlen-1),cstring)
    '    if len(cstring)=nstrlen
    '        exit
    '    End If
    'enddo
    'str_return = cstring
    '
    'return str_return
    '
    '
    '********************************************* END *********************************************
