Attribute VB_Name = "mdlDocument"
    '======================================================================================================================
    'һ��IC���ṹ˵��
    '�����������Щ����Щ��������Щ��������Щ���������������������������������
    '���ֶ�    �����ȩ���ַ���ȩ���ַ    ����ע                            ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '��Ա�����멦8   ��4       ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '�����֤�ũ�18  ��9       ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '������    ��8   ��8       ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '���Ա�    ��1   ��1       ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '���������©�8   ��4       ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '���������©�8   ��4       ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '��������λ��3   ��3       ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '����Ա���ʩ�1   ��0.5     ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '���ù����ʩ�1   ��0.5     ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '��סԺ������6   ��        ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '�����ۼƶ8   ��        ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '��        ��    ��        ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '��        ��    ��        ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '��        ��    ��        ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '��        ��    ��        ��        ��                                ��
    '�����������੤���੤�������੤�������੤��������������������������������
    '��סԺ��  ��12  ��        ��        �����5��סԺ��¼                 ��
    '�����������੤���੤�������੤��������                                ��
    '����Ժ���ک�8   ��        ��        ��                                ��
    '�����������੤���੤�������੤��������                                ��
    '����Ժ���ک�8   ��        ��        ��                                ��
    '�����������੤���੤�������੤��������                                ��
    '�����úϼƩ�8   ��        ��        ��                                ��
    '�����������੤���੤�������੤��������                                ��
    '��������8   ��        ��        ��                                ��
    '�����������੤���੤�������੤��������                                ��
    '���Ը���8   ��        ��        ��                                ��
    '�����������ة����ة��������ة��������ة���������������������������������
    
    '======================================================================================================================
    '����VFPʾ������
    
    '*�� �� ���� Func_ICC.prg
    '*�������ڣ� 2002��9��20��
    '*��    �ܣ�  ����IC���������йغ���
    '
    '**(0) declare_IC()        IC������˵��
    '**(1) init_card()         ��ʼ�����������뿪��Դ
    '**(2) close_card()        �رտ���
    '**(3) check_card()        ��ʼ��IC���������кϷ��Ŀ�
    '
    '**(4) rd_kxh()            �������
    '**(5) wr_kxh(kxh,fkcs)    д�����
    '
    '**(6) rd_sfz()            �����֤
    '**(7) wr_sfz(rybm,sfz)    д���֤
    '
    '**(8) rd_geren()          ��������Ϣ
    '**(9) wr_geren()          д������Ϣ
    '**(10) ts_geren()         ������ʾ������Ϣ
    '
    '**(11) rd_nlje()          ��סԺ���ۼƶ�
    '**(12) wr_nlje()          дסԺ���ۼƶ�
    '
    '**(13) rd_zyk()           ��סԺ��¼
    '**(14) wr_zyk()           дסԺ��¼
    '**(15) init_zypiont       ��ʼ��סԺ��¼ָ��
    '
    '**(16) check_hmd()        ���������
    '**(17) check_rybm()       ���IC�������ݼ�¼�Ƿ���ͬһ��
    '
    '*  =num_str()             ����ת��Ϊ8λ�ַ�
    '*  =str_num()             8λ�ַ�ת��Ϊ����
    '*  =str_add0()            �����ַ���ǰ�����
    '*  =STR_INC0()            �����ַ���ǰ��ȥ��
    '
    '
    '****(0) IC������˵�� *********************************************
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
    '****(1) ��ʼ����*********************************************
    'FUNC init_card
    '
    'declare_IC()   &&˵��IC���⺯��
    '
    'm_com=0        &&COM1
    'm_baud=9600    &&������
    '
    'p_icdev = auto_init(m_com, m_baud)
    '
    'if p_icdev<0
    '   =messagebox("��ʼ��IC������"+chr(13)+chr(13)+"�����Ƿ�򿪶����� ��",16,"����...")
    '   return .F.
    'End If
    '
    'return .T.
    'endfunc
    '
    '****(2) �رտ���*******************************************
    'Function close_card()
    '
    '    if p_icdev<=0
    '*      messagebox("IC������û׼���� �� û��ʼ�� ,���ܹر�IC������",16,"����")
    '       return .F.
    '    End If
    '
    '    ic_exit (p_icdev)
    '    p_icdve = 0
    '
    'Return
    '
    '
    '****(3) У�鿨*********************************************
    'Function check_card()
    'local m.st
    '
    'if p_icdev<=0
    '   messagebox("IC������û׼���� �� û��ʼ�� ��",16,"����")
    '   return .F.
    'End If
    '
    'm_status=0            &&���Կ����Ƿ��п�
    'st=get_status(p_icdev,@m_status)
    'if st<>0
    '    messagebox("û�в���IC����",16,"����...")
    '    return .f.
    'End If
    '
    'st=chk_4442(p_icdev)  &&����ǲ���4442��
    'if st<>0
    '    messagebox("�����Ͳ��ԣ�",16,"����...")
    '     return .f.
    'End If
    '
    'm_passwd0=space(3)  &&��ʼ����
    'm_passwd1=space(3)  &&��ϵͳ����
    'st=asc_hex("FFFFFF",@m_passwd0,3)
    'st=asc_hex("995188",@m_passwd1,3)
    '
    'st = csc_4442(p_icdev, 3, m_passwd1)
    'if st <> 0
    '    st=csc_4442(p_icdev,3,m_passwd0)  &&У��ԭʼ����
    '    if st=0
    '        st=wsc_4442(p_icdev,3,m_passwd1)  &&д������
    '        if st<0
    '            =messagebox("д�������",16,"����...")
    '            return .F.
    '        End If
    '    Else
    '        =messagebox("�Ǳ�ϵͳ��������ϵͳ�ṩ����ϵ��",16,"����...")
    '        return .F.
    '    End If
    'End If
    '
    'return .T.
    '
    'endfunc
    '
    '****(4) ������� *********************************************
    'Function rd_kxh()
    '
    'local st, m_str1,m_return,m_offset,m_len
    '
    'm_offset=0x1b          &&��ַ
    'm_len=3+1               &&����
    'm_str1=space(m_len)     &&�ַ�������
    'm_return=space(m_len*2) &&�����ַ���
    '
    'st=srd_4442(p_icdev,m_offset,m_len,@m_str1)
    'if st<0
    '   =messagebox('�����ݴ�(�����)��',16,'����...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_str1,@m_return,m_len)
    '
    'return m_return
    '
    '****(5) д����� *********************************************
    'Function wr_kxh()
    'PARA m_kxh, m_fkcs
    '
    'local st, m_str1,m_offset,m_len
    '
    'm_offset=0x1b       &&��ַ
    'm_len=3+1           &&����
    'm_str1=space(m_len) &&�ַ�������
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
    '   =messagebox('д���ݴ�(�����)��',16,'����...')
    '   return .F.
    'End If
    '
    'return .T.
    '
    '****(6) ����Ա���� *********************************************
    'Function rd_sfz()
    '
    'local st, m_str1,m_return,m_offset,m_len
    '
    'm_offset=0x08           &&��ַ
    'm_len=4+9               &&����
    'm_str1=space(m_len)     &&�ַ�������
    'm_return=space(m_len*2) &&�����ַ���
    '
    'st=srd_4442(p_icdev,m_offset,m_len,@m_str1)
    'if st<0
    '   =messagebox('�����ݴ�(��Ա����)��',16,'����...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_str1,@m_return,m_len)
    '
    'return m_return
    '
    '****(7) д��Ա���� *********************************************
    'Function wr_sfz()
    'PARA m_rybm, m_sfz
    '
    'local st, m_str1,m_offset,m_len
    '
    'm_offset=0x08      &&��ַ
    'm_len=4+9          &&����
    'm_str1=space(m_len) &&�ַ�������
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
    '   =messagebox('д���ݴ�(��Ա����)��',16,'����...')
    '   return .F.
    'End If
    '
    'return .T.
    '
    '
    '**** ��8����������Ϣ*********************************************
    '**���� �Ա� Ѫ�� �������� �������� ������λ ������λ ��Ա���� �ù�����
    'FUNC rd_geren
    'local st, m_str1,m_str2,m_return,m_offset,m_len
    '
    'm_offset=0x20              &&��ַ
    'm_len=24                   &&����  8+2+2+4+4+1.5+1.5+0.5+0.5
    'm_len2=12                  &&ת�����ݳ���
    'm_str1=space(m_len)        &&�ַ�������
    'm_str2 = Space(m_len2 * 2)
    'm_return=space(m_len*2-12) &&�����ַ���  ���� �Ա� Ѫ�� ����ѹ��
    '
    'st=srd_4442(p_icdev , m_offset , m_len , @m_str1)
    'if m.st<0
    '   messagebox("�����ݴ�(������Ϣ)��",16,"����...")
    '   Return ''
    'End If
    
    'm_return=substr(m_str1,1,12)     &&���� �Ա� Ѫ��
    '
    'm_str1=substr(m_str1,13)     &&����
    '
    'st=hex_asc(@m_str1,@m_str2,m_len2)
    '
    'm_return = alltrim(m_return) + m_str2
    '
    'return m_return
    '
    '****(9) д������Ϣ*********************************************
    '**���� �Ա� Ѫ�� �������� �������� ������λ ������λ ��Ա���� �ù�����
    'FUNC wr_geren
    'PARA m_xm, m_xb, m_blood, m_csny, m_gzny, m_gzdw, m_gzgw, m_ryxz, m_ygxz
    '
    'local st, m_str1,m_offset,m_len
    '
    'm_offset=0x20       &&��ַ
    'm_len=24            &&���� 8+2+2+4+4+1.5+1.5+0.5+0.5
    'm_str1=space(m_len) &&�ַ�������
    '
    'm_xm=str_add0(m_xm,8)         &&����
    'm_blood=str_add0(m_blood,2)   &&Ѫ��
    'm_xb=str_add0(m_xb,2)         &&�Ա�
    '
    'if empty(m_csny)          &&��������
    '    m_csny='19000101'
    'Else
    '    m_csny = dtos(m_csny)
    'End If
    '
    'if empty(m_gzny)          &&��������
    '    m_gzny='19000101'
    'Else
    '    m_gzny = dtos(m_gzny)
    'End If
    '
    'm_gzdw=str_add0(m_gzdw,3)     &&��λ����
    'm_gzgw=str_add0(m_gzgw,3)     &&��λ����
    'm_ryxz=str_add0(m_ryxz,1)     &&��Ա����
    'm_ygxz=str_add0(m_ygxz,1)     &&��Ա����
    '
    'st=asc_hex(m_csny + m_gzny + m_gzdw + m_gzgw + m_ryxz + m_ygxz , @m_str1 , m_len)
    '
    'm_str1=m_xm + m_xb +  m_blood + m_str1  &&�������ֶβ���ѹ��
    '
    'st=swr_4442(p_icdev , m_offset , m_len , @m_str1)
    '
    '
    'if st < 0
    '   =messagebox('д���ݴ�(������Ϣ)��',16,'����...')
    '   return .F.
    'End If
    '
    'return .T.
    '
    '****(10) ������ʾ������Ϣ *********************************************
    'Function ts_geren()
    '
    '    m_str=rd_sfz()  && ���֤
    '    if empty(m_str)
    '        Return
    '    End If
    '
    '    m_rybm = STR_INC0(substr(m_str, 1, 8))
    '    m_sfzh = STR_INC0(substr(m_str, 9, 18))
    '
    '    m_str=rd_geren()  && ������Ϣ
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
    '            m_mess = "���˵�Ա��������Ϣ�Ѿ����ģ��뵽������λ���¿�Ƭ��Ϣ��"
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
    '****(11) �����ۼƶ� *********************************************
    '** סԺ���� ���ۼƶ�
    'Function rd_nlje()
    '
    'local st, m_str1,m_return,m_offset,m_len
    '
    'm_offset=0x38           &&��ַ
    'm_len=7                 &&���� 3+4
    'm_str1=space(m_len)     &&�ַ�������
    'm_return=space(m_len*2) &&�����ַ���
    '
    'st=srd_4442(p_icdev,m_offset,m_len,@m_str1)
    'if st<0
    '   =messagebox('�����ݴ�(סԺ����)��',16,'����...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_str1,@m_return,m_len)
    '
    'return m_return
    '
    '****(12) д���ۼƶ� *********************************************
    '** סԺ���� ���ۼƶ�
    'Function wr_nlje()
    'PARA m_zycs, m_nlje
    '
    'local st, m_str1,m_offset,m_len
    '
    'm_offset=0x38       &&��ַ
    'm_len=7             &&���� 3+4
    'm_str1=space(m_len) &&�ַ�������
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
    '   =messagebox('д���ݴ�(סԺ����)��',16,'����...')
    '   return .F.
    'End If
    '
    'return .T.
    '
    '
    '****(13) ��סԺ��¼ *********************************************
    '** ƾ֤�� ��Ժ���� ���úϼ� �������
    'Function rd_zyk()
    'PARA m_recno
    'local st, m_str1,m_return,m_offset,m_len,m_rec,m_rec1,m_recc
    '
    '
    'm_offset1=0x40           &&��ַ
    'm_offset2=0x60           &&��ַ
    'm_offset3=0x80           &&��ַ
    'm_offset4=0xa0           &&��ַ
    'm_offset5=0xc0           &&��ַ
    'm_len=16                &&���� 5+3+4+4
    'm_str1=space(m_len)     &&�ַ�������
    'm_return=space(m_len*2+4) &&�����ַ���
    'm_rec=space(1)            &&��¼ָ��
    'm_rec1 = Space(2)
    '
    'm.st=srd_4442(p_icdev,0x3F,1,@m_rec)  &&����ǰ��¼ָ��
    'if m.st<0
    '   =messagebox('�����ݴ�(��¼ָ��)��',16,'����...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_rec,@m_rec1,1)
    'm_recc = Val(substr(m_rec1, 1, 1))
    'm_rec = Val(substr(m_rec1, 2, 1))
    '
    'if m_recno>m_recc  &&��Χ
    '    return ""
    'End If
    '
    'if m_rec>5 or m_rec<1
    '   =messagebox('�����ݴ�(��¼ָ�볬��Χ)��',16,'����...')
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
    '   =messagebox('�����ݴ�(סԺ��¼)��',16,'����...')
    '   Return ''
    'End If
    '
    'st=hex_asc(@m_str1,@m_return,m_len)
    '
    'if substr(m_return,1,2)="FF"  &&��û��д��¼
    '    return ""
    'End If
    '
    'm_return = "20" + substr(m_return, 1, 10) + "20" + alltrim(substr(m_return, 11))
    '
    'return m_return
    '
    '****(14) дסԺ��¼ *********************************************
    '** ƾ֤�� ��Ժ���� ���úϼ� �������
    'Function wr_zyk()
    'PARA m_pzh, m_cyrq, m_hj, m_bxje
    '
    'local st, m_str1,m_offset,m_len , m_rec , m_rec1 , m_recc
    '
    'm_offset1=0x40           &&��ַ
    'm_offset2=0x60           &&��ַ
    'm_offset3=0x80           &&��ַ
    'm_offset4=0xa0           &&��ַ
    'm_offset5=0xc0           &&��ַ
    'm_len=16                &&���� 5+3+4+4
    'm_str1=space(m_len)     &&�ַ�������
    'm_rec1=space(2)            &&��¼ָ��
    'm_rec = Space(2)
    '
    'm.st=srd_4442(p_icdev,0x3F,1,@m_rec)  &&����ǰ��¼ָ��
    'if m.st<0
    '   =messagebox('�����ݴ�(��¼ָ��)��',16,'����...')
    '   Return ''
    'End If
    'st=hex_asc(@m_rec,@m_rec1,1)
    'm_recc=substr(m_rec1,1,1)  &&���м�¼��
    'm_rec =substr(m_rec1,2,1)   &&��ǰ��¼��
    '
    'if m_rec>"5" and m_rec<"0"
    '   =messagebox('�����ݴ�(��¼ָ�볬��Χ)��',16,'����...')
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
    'm_pzh=str_add0(substr(m_pzh,3,10),10)  &&ƾ֤��
    '
    'if empty(m_cyrq)
    '    m_cyrq = Date
    'End If
    'm_cyrq=substr(dtos(m_cyrq),3,6)           &&��Ժ����
    '
    'm_hj=num_str(m_hj,8)        &&�ϼƽ��
    'm_bxje=num_str(m_bxje,8)    &&�������
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
    '   =messagebox('д���ݴ�(סԺ��¼)��',16,'����...')
    '   return .F.
    'End If
    '
    'm_rec = alltrim(Str(m_rec))
    'm_rec1 = alltrim(Str(m_recc)) + m_rec
    'st=asc_hex(@m_rec1,@m_rec,1)
    'st=swr_4442(p_icdev , 0x3F , 1 , @m_rec)  &&д��¼ָ��
    '
    'return .T.
    '
    '****(15) ��ʼ��סԺ��¼ָ�� *********************************************
    'Function init_zypiont()
    'local m_rec
    '
    '    m_rec = Chr(0)
    '    st=swr_4442(p_icdev , 0x3F , 1 , @m_rec)
    '    if st < 0
    '        =messagebox('д���ݴ�(סԺ��¼ָ��)��',16,'����...')
    '        return .F.
    '    End If
    '
    'return .T.
    '
    '****(16) �������� *********************************************
    'Function check_hmd()
    '
    'return .T.
    '
    '****(17) ���IC�������ݼ�¼�Ƿ���ͬһ�� *********************************************
    'Function check_rybm()
    '
    'return .T.
    '
    '
    '********************************************* 1λС��λ��ֵת��Ϊ8λ�ַ��� *********************************************
    'FUNC num_str
    'PARA nnum, nstrlen
    'local m_str
    '
    'if nnum<0
    '   messagebox("����ת��Ϊ�ַ�����Ϊ����!",16,"����")
    '   Return ''
    'End If
    '
    'm_str=str(nnum*10,nstrlen)             &&ȥ��1λС����,  ǰ��ӿո�
    'm_str=strtran(m_str,' ','0')  &&��'0'����ո�
    '
    'return m_str
    '
    '********************************************* 8 �ַ��� ת��Ϊ ��ֵ *********************************************
    'FUNC str_num
    'PARA m_str
    'local m_nnum
    '
    'set decimals to 0
    'm_nnum=val(m_str)/10    &&1λС����
    '
    'return m_nnum
    '
    '********************************************* �ַ���ǰ���'0' *********************************************
    'FUNC STR_add0
    'PARA cstring, nstrlen
    'local str_return,cstr,nstrlen0,m_len
    '
    'cstring = alltrim(cstring)
    'nstrlen0 = Len(cstring)
    'do case                                   &&ʵ�ʳ�������Ҫ���ȱȽ�
    '   case nstrlen0<nstrlen                  &&ǰ���'0'
    '        m_len = nstrlen - nstrlen0
    '        cstr=space(m_len)+cstring
    '        str_return=strtran(cstr,' ','0')
    '   case nstrlen0=nstrlen                  &&����
    '        str_return = cstring
    '   otherwise                              &&ʵ�ʳ��ȹ���
    '        str_return = substr(cstring, 1, nstrlen)
    'endcase
    '
    'return str_return
    '
    '********************************************* ȥ���ַ���ǰ��'0' *********************************************
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
