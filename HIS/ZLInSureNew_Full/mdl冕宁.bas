Attribute VB_Name = "mdl����"
Option Explicit
'//����ҽԺϵͳ�û���Ų�ѯҽ��ϵͳ�û����
Public Declare Function finduser Lib "YBXT.dll" (ByVal yyid As String) As Boolean
'//�����û���Ϣ
Public Declare Function inuser Lib "YBXT.dll" (ByVal yyid As String, ByVal user_name As String) As Double
'//�޸��û���Ϣ
Public Declare Function edituser Lib "YBXT.dll" (ByVal yyid As String, ByVal user_name As String) As Double
'//ɾ���û���Ϣ
Public Declare Function deluser Lib "YBXT.dll" (ByVal yyid As String) As Double
'//����ҽԺϵͳ�Ʊ��Ų�ѯҽ��ϵͳ�Ʊ���
Public Declare Function findkb Lib "YBXT.dll" (ByVal yykbbh As String) As Boolean
'//�����Ʊ���Ϣ
Public Declare Function inkb Lib "YBXT.dll" (ByVal yykbbh As String, ByVal mc As String) As Double
'//�޸ĿƱ���Ϣ
Public Declare Function editkb Lib "YBXT.dll" (ByVal yykbbhas As String, ByVal mc As String) As Double
'//ɾ���Ʊ���Ϣ
Public Declare Function delkb Lib "YBXT.dll" (ByVal yykbbh As String) As Double
'//����ҽԺϵͳҽ����Ų�ѯҽ��ϵͳҽ�����
Public Declare Function findys Lib "YBXT.dll" (ByVal yyysbh As String) As Boolean
'//����ҽ����Ϣ
Public Declare Function inys Lib "YBXT.dll" (ByVal yyysbh As String, ByVal yykbbh As String, ByVal xm As String) As Double
'//�޸�ҽ����Ϣ
Public Declare Function editys Lib "YBXT.dll" (ByVal yyysbh As String, ByVal yykbbh As String, ByVal xm As String) As Double
'//ɾ��ҽ����Ϣ
Public Declare Function delys Lib "YBXT.dll" (ByVal yyysbh As String) As Double
'//����ҽԺ��ĿDM��ѯҽ����ĿDM
Public Declare Function findxm Lib "YBXT.dll" (ByVal dm As String) As String
'//������Ŀ
Public Declare Function inxm Lib "YBXT.dll" (ByVal yyxmdm As String, ByVal yyfldm As String, ByVal xmzl As String, _
                                            ByVal zlxmmc As String, ByVal zlflmc As String, ByVal pybh As String, _
                                            ByVal dj As String, ByVal jldw As String) As Double
'//�޸���Ŀ
Public Declare Function editxm Lib "YBXT.dll" (ByVal yyxmdm As String, ByVal zlxmmc As String, ByVal pybh As String, _
                                              ByVal dj As String, ByVal jldw As String) As Double
'//ɾ����Ŀ
Public Declare Function delxm Lib "YBXT.dll" (ByVal yyxmdm As String) As Double
'//��ѯָ����Ŀ��Ӧ��ҽ����Ŀ��Ϣ��
Public Declare Function xmcx Lib "YBXT.dll" (ByVal yyxmdm As String) As Double

'//��Ժ��ʼ��
Public Declare Function rycsh Lib "YBXT.dll" (ByRef sbh As String, ByRef zhzt As String, ByRef zffs As String, _
                                             ByRef net As String, ByRef Zhye As Double, ByRef tcqfx As Double) As Double
'//֧������[��Ժ�����ɷ���]
Public Declare Function zffp Lib "YBXT.dll" (ByVal zhzt As String, ByVal yjje As Double, ByVal Zhye As Double) As Double
'//д��Ժ��
Public Declare Function ryappend Lib "YBXT.dll" (ByVal sbh As String, ByVal yyzyh As String, ByVal yykb As String, _
                                                ByVal yyzdys As String, ByVal yysfydm As String, ByVal ryzd As String, _
                                                ByVal lxdh As String, ByVal jtzz As String, ByVal Bz As String, _
                                                ByVal net As String, ByVal yjje As Double, ByVal zhyjje As Double) As Double
'��λ���
Public Declare Function cwappend Lib "YBXT.dll" (ByVal yyzyh As String, ByVal yyczydm As String, ByVal cwh As String, ByVal djrq As String) As Double
'��λɾ��
Public Declare Function cwdel Lib "YBXT.dll" (ByVal yyzyh As String, ByVal cwh As String, ByVal djrq As String) As Double

'//���ɷ��ó�ʼ��
Public Declare Function bjcsh Lib "YBXT.dll" (ByVal sbh As String, ByVal zhzt As String, ByVal zffs As String, _
                                            ByVal net As String, ByVal yyzyh As String, ByVal Zhye As Double) As Double
'//д���ɷ��ñ�
Public Declare Function bjappend Lib "YBXT.dll" (ByVal sbh As String, ByVal sjh As String, ByVal yyzyh As String, _
                                                ByVal yysfydm As String, ByVal Bz As String, ByVal net As String, _
                                                ByVal bjje As Double, ByVal zhbjje As Double) As Double

'//��Ӽ�����Ϣ
Public Declare Function jzappend Lib "YBXT.dll" (ByVal yyzyh As String, ByVal yyxmdm As String, ByVal yyczydm As String, _
                                                ByVal yycfys As String, ByVal djh As String, ByVal Bz As String, _
                                                ByVal sl As String, ByVal je As String, ByVal cfrq As String) As Double
'//ɾ��������Ϣ
Public Declare Function jzdel Lib "YBXT.dll" (ByVal yyzyh As String, ByVal Bz As String) As Double

'//���㲡�˵�ǰ����������Ϣ
Public Declare Function jsbxf Lib "YBXT.dll" (ByVal yyzyh As String, ByRef qfx As Double, ByRef ylf As Double, _
                                            ByRef Tcje As Double, ByRef bxje As Double, ByRef Bcbxje As Double, _
                                            ByRef Jbbcbxje As Double, ByRef Gwybcbxje As Double, ByRef Tsbxje As Double) As Double

'//��Ժ��ʼ��
Public Declare Function cycsh Lib "YBXT.dll" (ByRef sbh As String, ByRef zhzt As String, ByRef zffs As String, _
                                            ByRef net As String, ByRef yyzyh As String, ByRef Mxbbz As String, _
                                            ByRef rzhye As Double, ByRef rylf As Double, ByRef rqfx As Double, _
                                            ByRef rtcje As Double, ByRef ryjje As Double, ByRef rzhyjje As Double, _
                                            ByRef rbxje As Double, ByRef rbcbxje As Double, ByRef rpbxje As Double, _
                                            ByRef rnbxje As Double, ByRef rpbcbxje As Double, ByRef rnbcbxje As Double, _
                                            ByRef rbcjs1 As Double, ByRef rbcjs2 As Double, ByRef rzhzf As Double, _
                                            ByRef rxjzf As Double, ByRef rzhtk As Double, ByRef rxjtk As Double, _
                                            ByRef rjbbcbxje As Double, ByRef rgwybcbxje As Double, ByRef rtsbxje As Double, _
                                            ByRef rpjbbcbxje As Double, ByRef rnjbbcbxje As Double, ByRef rpgwybcbxje As Double, _
                                            ByRef rngwybcbxje As Double, ByRef rptsbxje As Double, ByRef rntsbxje As Double) As Double
'//д��Ժ��
Public Declare Function cyappend Lib "YBXT.dll" (ByVal yyzyh As String, ByVal bah As String, ByVal Cyzd As String, _
                                                ByVal djh As String, ByVal yysfy As String, ByVal Bz As String, _
                                                ByVal Bl As String, ByVal Mxbbz As String, ByVal qfx As Double, _
                                                ByVal Tcje As Double, ByVal Zhzf As Double, ByVal Grzf As Double, _
                                                ByVal Zhtk As Double, ByVal Xjtk As Double, ByVal bxje As Double, _
                                                ByVal Bcbxje As Double, ByVal Pbxje As Double, ByVal Nbxje As Double, _
                                                ByVal Pbcbxje As Double, ByVal Nbcbxje As Double, ByVal Bcjs1 As Double, _
                                                ByVal Bcjs2 As Double, ByVal Jbbcbxje As Double, ByVal Gwybcbxje As Double, _
                                                ByVal Tsbxje As Double, ByVal Pjbbcbxje As Double, ByVal Njbbcbxje As Double, _
                                                ByVal Pgwybcbxje As Double, ByVal Ngwybcbxje As Double, ByVal Ptsbxje As Double, _
                                                ByVal Ntsbxje As Double) As Double
        
'//�����ʼ��
Public Declare Function mzcsh Lib "YBXT.dll" (ByRef sbh As String, ByRef net As String, _
                                     ByRef rylx As String, ByRef zhzt As String, _
                                     ByRef Zhye As Double) As Double
'//�����������ﲡ�˱������ã�����ÿ����ϸ����ͳ�ﲿ�֣�
Public Declare Function mzfyjs Lib "YBXT.dll" Alias "mzjs" (ByVal sbh As String, ByVal yyxmdm As String, _
                                            ByVal sl As Long, ByVal je As Double, _
                                            ByRef bxbl As Double, ByRef bxje As Double, _
                                            ByRef bsbhddj As Double) As Double
'�������ⲡ�˱������ü������֧������(������ϸ�Ľ���ͳ��ϼƣ���������������֣�
Public Declare Function mztsjs Lib "YBXT.dll" (ByVal sbh As String, ByVal yydjh As String, ByVal bxje As Double) As Double

'//д����֧����ϸ
Public Declare Function mzzfmxappend Lib "YBXT.dll" (ByVal sbh As String, ByVal yyxmdm As String, _
                                                    ByVal yydjh As String, ByVal yysfydm As String, _
                                                    ByVal rylx As String, ByVal sl As Long, _
                                                    ByVal je As Double, ByVal dj As Double, _
                                                    ByVal bxbl As Double, ByVal bxje As Double) As Double
'//д����֧����
Public Declare Function mzappend Lib "YBXT.dll" (ByVal sbh As String, ByVal yydjh As String, _
                                                ByVal yysfydm As String, ByVal yyysdm As String, _
                                                ByVal rylx As String, ByVal net As String, _
                                                ByVal Zhzf As Double, ByVal xjzf As Double, _
                                                ByVal bxje As Double) As Double

'//סԺ���籣����
Public Declare Sub zyjs Lib "YBMOD.dll" (ByVal yyczy As String)
'//�������籣����
Public Declare Sub mzjs Lib "YBMOD.dll" (ByVal yyczy As String)
'//��λ�Ǽ�
Public Declare Sub cwdj Lib "YBMOD.DLL " (ByVal yyczy As String)
'//ҽ������סԺ��Ϣ��ѯ
Public Declare Sub ybcx Lib "YBMOD.dll" (ByVal yyczy As String)
'//��ѯ��Աҽ����Ϣ
Public Declare Function getyhxx_vb Lib "YBXT.dll" (ByVal sbh As String, ByRef fhz As String) As Double
'//��ѯ��ѯ����סԺ�Ƿ�ͨ������
Public Declare Function getspxx Lib "YBXT.dll" (ByVal yyzyh As String) As Double

Public Declare Function mzzffp Lib "YBXT.dll" (ByVal zhzt As String, ByVal yjje As Double, ByVal Zhye As Double, ByVal yydjh As String) As Double

'//ɾ������Ԥ���ϴ�����ϸ
Public Declare Function mzdeltmp Lib "YBXT.dll" (ByVal yydjh As String) As Boolean

Private mrsCdTmp As New ADODB.Recordset   '��ʱ��¼��
Private mstrCdSql As String    '��ʱ���SQL���Public Sub ���²���_����(lngPatiID, lngPageID)
Public gstrPuser_id As String '����Ա����

Private mblnInit As Boolean

'/////��Ժ�Ǽ�Ҫ���Ĳ���̫�࣬���浽���ݿ��в����㣬���Զ���Ϊ�������������
Private m_yyzyh As String * 18 'סԺ�ţ��ǲ���סԺ��Ψһ��ʶ�š��������ɳ�Ժ��ʼ����������ֵ�õ�������Ϊ�գ�
Private m_bah As String '������ (����Ϊ��)
Private m_cyzd As String ' ��Ժ��� (����Ϊ��)
Private m_djh As String ' �վݺ� (����Ϊ��)
Private m_yysfy As String ': �շ�Ա����? (����Ϊ��)

Private m_Bz As String ': ��ע? (����Ϊ��)
Private m_Bl As String ': ������? (����Ϊ��)
Private m_Mxbbz As String * 1 ': ���Բ���־? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Qfx As Double ': ͳ�������? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Tcje As Double '��ͳ������ϱ��������Ľ������ɳ�Ժ��ʼ����������ֵ�õ���

Private m_Zhzf As Double ': ��Ժ�ʻ�֧�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Grzf As Double '����Ժ�ֽ�֧����xjzf�������ɳ�Ժ��ʼ����������ֵ�õ���
Private m_Zhtk As Double ' ��Ժ�ʻ��˿���? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Xjtk As Double ': ��Ժ�ֽ��˿���? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Bxje As Double ': �����������? (�ɳ�Ժ��ʼ����������ֵ�õ�)

Private m_Bcbxje As Double ': �߶�䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Pbxje As Double ': ����Ȼ����������? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Nbxje As Double ': ����Ȼ����������? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Pbcbxje As Double ': ����ȸ߶�䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Nbcbxje As Double ': ����ȸ߶�䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)

Private m_Bcjs1 As Double ': �������1? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Bcjs2 As Double ': �������2? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Jbbcbxje As Double ': �������䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Gwybcbxje As Double ' : ����Ա���䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Tsbxje As Double ': ���ⱨ�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)

Private m_Pjbbcbxje As Double ': ����Ȼ������䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Njbbcbxje As Double ': ����Ȼ������䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Pgwybcbxje As Double ': ����ȹ���Ա���䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Ngwybcbxje As Double ': ����ȹ���Ա���䱨�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)
Private m_Ptsbxje As Double ': ��������ⱨ�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)

Private m_Ntsbxje As Double ': ��������ⱨ�����? (�ɳ�Ժ��ʼ����������ֵ�õ�)

Private m_rylf As Double   'ҽ�Ʒ��ܶ�,�����Ҫ����
Private m_rxjzf As Double  '�ֽ�֧�����,�����Ҫ����
'/////��Ժ�Ǽ�Ҫ���Ĳ���̫�࣬���浽���ݿ��в����㣬���Զ���Ϊ�������������

Public Function ����������_����(lngStlID, curMoney, lng����ID) As Boolean
'��clsInsure ��  ClinicDelSwap ���̵���
'����˵��:��ҽ�����������
On Error GoTo ErrH
    Err.Raise 9000, gstrSysName, "����ҽ���涨:����������ѽ��㵥��!", vbInformation, gstrSysName
    ����������_���� = False
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �����������_����(rs��ϸ��¼ As ADODB.Recordset, str���㷽ʽ As String, Optional str���� As String) As Boolean
'��clsInsure �� ClinicPreSwap ���̵���
'����˵��:�������Ԥ���㹦��
'���ú����嵥��˵��
'
'mzfyjs       ������ü���
'mztsjs       �����������
'mzzfmxappend ������ϸ����
'mzzffp       ����֧������
'mzdeltmp     ������ʱ��¼ɾ��
'
Dim sbh As String  '�籣�� (����)
Dim yyxmdm As String 'ҽԺ��Ŀ���� (����)
Dim sl  As Long      '����(����)
Dim dj As Double   '���� (����)
Dim je As Double    '��� ( ����)
Dim tmp As Double '���նԷ����صĽ���״̬
Dim bxbl As Double '�������� (����)
Dim bxje As Double  '������� (����)
Dim sbhddj As Double '�籨�˶��� (����)
Dim lng����ID As Long
Dim bxjeHj As Double  '�������ϼ�
Dim ssjehj As Double    'ʵ�ս��ϼ�
Dim rylx As String    '��������(����)
Dim zhzt As String '�ʻ�״̬(����)
Dim yjje As Double  '�ɷѽ��(����) =ʵ�ս��ϼ�-�������ϼ�
Dim Zhye As Double  '�ʻ����
Dim yydjh As String  'ҽԺ���ݺ� (����) ��ʽ:1 & NO
Dim yysfydm As String 'ҽԺ�շ�Ա����(����)
Dim blReturn As Boolean  ' ����ɾ����ϸ�ķ���ֵ
Dim dbl�����ʻ� As Double
Dim dblͳ����� As Double

On Error GoTo errHandle
    If rs��ϸ��¼.RecordCount = 0 Then
        MsgBox "û�в��˷�����ϸ�����ܽ���ҽ������", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Beging ȡҽ����
    lng����ID = rs��ϸ��¼!����ID
    mstrCdSql = "Select * from �����ʻ� where ����ID=[1] And ����=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "������Ϣ", lng����ID, TYPE_����)
    
    If mrsCdTmp.EOF Then
        MsgBox "�������������ˣ�������ִ�н��ס�", vbInformation, gstrSysName
        Exit Function
    End If
    sbh = mrsCdTmp.Fields!ҽ����
    rylx = mrsCdTmp.Fields!�������
    Zhye = mrsCdTmp.Fields!�ʻ����
    zhzt = mrsCdTmp.Fields!��Ա���
    
    'Ene ȡҽ����
        
    'Beging �����ϸ��¼�Ƿ����,�Ե�����ҽ�������Ƿ����
    Do Until rs��ϸ��¼.EOF
        mstrCdSql = "select * from ����֧����Ŀ where ����=[1] and �շ�ϸĿID=[2]"
        Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "ҽ����ϸ", TYPE_����, CLng(rs��ϸ��¼!�շ�ϸĿID))
        If mrsCdTmp.EOF Then
            mstrCdSql = "Select * from �շ���ĿĿ¼ where ID=[1]"
            Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "�շ���ĿĿ¼", CLng(rs��ϸ��¼!�շ�ϸĿID))
            MsgBox mrsCdTmp!���� & "(" & mrsCdTmp!���� & ")" & "δ���룡" & vbCrLf & "��������ʹ�ô˹��ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        yyxmdm = mrsCdTmp!��Ŀ����
        '�����ҽ���Ƿ����
        'findxm (yyxmdm)
        rs��ϸ��¼.MoveNext
    Loop
    'End �����ϸ��¼�Ƿ����,�Ե�����ҽ�������Ƿ����
    
    
    'Beging ����ϴ���ϸ,������ÿ����ϸ�ı�������,�������,
    rs��ϸ��¼.MoveFirst
    
    bxbl = 0
    bxje = 0
    sbhddj = 0
    bxjeHj = 0
    ssjehj = 0
    Do Until rs��ϸ��¼.EOF
        mstrCdSql = "select * from ����֧����Ŀ where ����=[1] and �շ�ϸĿID=[2]"
        Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "ҽ����ϸ", TYPE_����, CLng(rs��ϸ��¼!�շ�ϸĿID))
        yyxmdm = mrsCdTmp!��Ŀ����
        sl = rs��ϸ��¼!����
        dj = rs��ϸ��¼!����
        je = rs��ϸ��¼!ʵ�ս��
        If Nvl(str����, 0) = 9 Then
            yydjh = "1" & rs��ϸ��¼!NO
        Else
            yydjh = lng����ID & Format(rs��ϸ��¼!����ʱ��, "yyMMddHHmmdd")
        End If
        ssjehj = ssjehj + je
        yysfydm = gstrPuser_id
        '��Ա����Ϊ3 �����ⲡ�ˣ�,Ҫ���㱨�����
        If rylx = 3 Then
            tmp = mzfyjs(sbh, yyxmdm, sl, je, bxbl, bxje, sbhddj)
            Call WriteBusinessLOG("mzfyjs", "sbh, yyxmdm, sl, je, bxbl, bxje, sbhddj", tmp & "," & sbh & "," & yyxmdm & "," & sl & "," & je & "," & bxje & "," & sbhddj)
            bxjeHj = bxjeHj + bxje
            
        End If
        
        
        tmp = mzzfmxappend(sbh, yyxmdm, yydjh, yysfydm, rylx, sl, je, dj, bxbl, bxje)
        Call WriteBusinessLOG("mzzfmxappend", "sbh, yyxmdm, yydjh, yysfydm, rylx, sl, je, dj, bxbl, bxje", tmp & "," & sbh & "," & yyxmdm & "," & yydjh & "," & yysfydm & "," & rylx & "," & sl & "," & je & "," & dj & "," & bxbl & "," & bxje)
        
        Select Case tmp
            Case 0
                If Nvl(str����, 0) = 9 Then
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rs��ϸ��¼!ID & "," & _
                            bxje & _
                            ",NULL,1,NULL,1," & bxbl & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ���ֶ�")
                    
                End If
            Case 1
                MsgBox "���ݺţ������ţ��ظ�", vbInformation, gstrSysName
                Exit Function
            Case 2
                MsgBox "������Ϣ����û�иòα���Ա", vbInformation, gstrSysName
                Exit Function
            Case 3
                MsgBox "ҽ����Ŀ��û�и���Ŀ", vbInformation, gstrSysName
                Exit Function
            Case 99
                MsgBox "����", vbInformation, gstrSysName
                Exit Function
        End Select
    
    
      rs��ϸ��¼.MoveNext
    Loop
    'end ����ϴ���ϸ
    
    '
    
    yjje = ssjehj - bxjeHj 'ʵ�ս��-�������
    dbl�����ʻ� = mzzffp(zhzt, yjje, Zhye, yydjh)
    Call WriteBusinessLOG("Mzzffp", "zhzt, yjje, zhye, yydjh", dbl�����ʻ� & "," & zhzt & "," & yjje & "," & Zhye & "," & yydjh)
    
    dblͳ����� = bxjeHj
    '��Ա����=3,�������>0
    If rylx = 3 And bxjeHj > 0 Then
        dblͳ����� = mztsjs(sbh, yydjh, bxjeHj)
        Call WriteBusinessLOG("mztsjs", "sbh, yydjh, bxjehj", dblͳ����� & "," & sbh & "," & yydjh & "," & bxjeHj)
    End If
    
    'beging ��Ԥ����,Ҫ����ɾ����ϸ
    If Nvl(str����, 0) <> 9 Then
        blReturn = mzdeltmp(yydjh)
        Call WriteBusinessLOG("mzdeltmp", yydjh, IIf(blReturn, "True", "False"))
    End If
    'end ��Ԥ����
    
        '����ֵ��HIS
    str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";1|ͳ�����;" & dblͳ����� & ";0|����Ա����;" & 0 & ";0"
    �����������_���� = True
    
    Exit Function

errHandle:
    �����������_���� = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID, cur����֧��, strҽ����) As Boolean
    '��clsInsure �� ClinicSwap ���̵���
    '����˵��:����������
    '���ù����嵥������˵��
    '    �����������_����: �������Ԥ���㹦��
    '    ҽԺҽʦ����_����: ��ѯ�Ƿ���ָ��ҽʦ,���û�������,����ҽʦ���
    '    mzappend         : �������Ǽ�
        
    Dim sbh As String  '�籣�� (����)
    Dim yydjh As String  'ҽԺ���ݺ� (����) ��ʽ:1 & NO
    Dim yysfydm As String 'ҽԺ�շ�Ա����(����)
    Dim yyysdm As String  'ҽԺҽʦ����(����)
    Dim rylx As String    '��������(����)
    Dim net As String '����״̬(����)
    Dim Zhzf As Double '�ʻ�֧��
    Dim xjzf As Double  '�ֽ�֧��
    Dim bxje As Double  '������� (����)
    Dim strԤ����Ϣ As String '����Ԥ�㷵�ص���Ϣ
    Dim tmp As Double '���նԷ����صĽ���״̬
    Dim lng����ID As Long
    
    Dim rscd��ϸ��¼ As New ADODB.Recordset
On Error GoTo ErrH
    mstrCdSql = "Select ID,NO,���,��¼����,�Ǽ�ʱ�� as ����ʱ��,����ID,�շ����,�վݷ�Ŀ,���㵥λ,������, " & _
                "�շ�ϸĿID,nvl(����,0)*nvl(����,0) as ����,��׼���� as ����, " & _
                "ʵ�ս��,ͳ����,���մ���ID ����֧������ID, " & _
                "ժҪ,�Ƿ��� " & _
                "from ������ü�¼ " & _
                "where ����ID=[1]"
    Set rscd��ϸ��¼ = zlDatabase.OpenSQLRecord(mstrCdSql, "���ü�¼", lng����ID)
    
    yyysdm = ҽԺҽʦ����_����(rscd��ϸ��¼!������)
    lng����ID = rscd��ϸ��¼!����ID
    yydjh = 1 & rscd��ϸ��¼!NO
    yysfydm = gstrPuser_id
    
    '���ϴ���ϸ
    If �����������_����(rscd��ϸ��¼, strԤ����Ϣ, 9) = False Then
        Exit Function
    End If
    
    mstrCdSql = "Select * from �����ʻ� where ����=[1] And ����ID=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "�����ʻ�", TYPE_����, lng����ID)
    sbh = mrsCdTmp.Fields!ҽ����
    rylx = mrsCdTmp.Fields!�������
    net = mrsCdTmp.Fields!����
    Zhzf = cur����֧��
    
    mstrCdSql = "Select sum(nvl(��Ԥ��,0)) as ��� from ����Ԥ����¼ where ���㷽ʽ='�ֽ�' And ����ID=[1]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "�����ʻ�", lng����ID)
    xjzf = Nvl(mrsCdTmp!���, 0)
    
    mstrCdSql = "Select sum(nvl(��Ԥ��,0)) as ��� from ����Ԥ����¼ where ���㷽ʽ='ͳ�����' And ����ID=[1]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "�����ʻ�", lng����ID)
    bxje = Nvl(mrsCdTmp!���, 0)
    
    tmp = mzappend(sbh, yydjh, yysfydm, yyysdm, rylx, net, Zhzf, xjzf, bxje)
    Call WriteBusinessLOG("mzappend", "sbh,yydjh,yysfydm,yyysdm,rylx,net,zhzf,xjzf,bxje", tmp & "," & sbh & "," & yydjh & "," & yysfydm & "," & yyysdm & "," & rylx & "," & net & "," & Zhzf & "," & xjzf & "," & bxje)
      
    Select Case tmp
        Case 0
            �������_���� = True
        Case 1
            Err.Raise 9000, gstrSysName, "���ݺţ������ţ��ظ�"
            Exit Function
        Case 2
            Err.Raise 9000, gstrSysName, "������Ϣ����û�иòα���Ա"
            Exit Function
        Case 3
            Err.Raise 9000, gstrSysName, "�����ˢ������"
            Exit Function
        Case Else
            Err.Raise 9000, gstrSysName, "����"
            Exit Function
    End Select
      
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & _
            lng����ID & "," & Format(Now(), "YYYY") & ",0,0, " & _
            "" & _
            0 & ",NULL,NULL,NULL,NULL,0," & _
            Zhzf + xjzf + bxje & "," & xjzf & ",NULL,NULL," & bxje & ",NULL,NULL," & _
            cur����֧�� & ",NULL,NULL,NULL,'" & rylx & "')"
            '                                 ��������
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ������Ժ�Ǽ�_����(lngPatiID, lngPageID) As Boolean
'��clsInsure �� ComeInDelSwap  ���̵���
'����: ��֧�ִ˽���
    ������Ժ�Ǽ�_���� = False
    MsgBox "����ҽ����֧�ֳ�����Ժ", vbInformation, gstrSysName
End Function

Public Function ��Ժ�Ǽ�_����(lngPatiID, lngPageID, strҽ����) As Boolean
'��clsInsure �� ComeInSwap ���̵���
'����:���ҽ��������Ժ�Ǽ�
'���ù����嵥��˵��:
'    ҽԺ���Ҵ���_���� : ��ѯ�Ƿ���ָ������,���û�������,���ؿ��ұ��
'    ҽԺҽʦ����_���� : ��ѯ�Ƿ���ָ��ҽʦ,���û�������,����ҽʦ���
'    ryappend          : ҽ����Ժ�Ǽǽ���

    Dim zt  As Double '���շ���ֵ
    Dim strMsg As String '��ʾ��Ϣ
    Dim sbh As String    '�籣��
    Dim yyzyh As String  'ҽԺסԺ��
    Dim yykb As String   'ҽ�����ڿƱ�
    Dim yyzdys As String 'ҽԺ���ҽʦ
    Dim yysfydm As String 'ҽԺ�շ�Ա����
    Dim ryzd As String '��Ժ���
    Dim lxdh  As String ' ��ϵ�绰
    Dim jtzz As String '��ͥסַ
    Dim Bz As String '��ע
    Dim net As String '����״̬
    Dim iyjje As Double 'Ԥ�����
    Dim izhyj As Double  '�ʻ�Ԥ��
    
    mstrCdSql = "Select * from �����ʻ� where ����=[1] And ����ID=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "�����ʻ�", TYPE_����, lngPatiID)
    sbh = mrsCdTmp!ҽ����
    yyzyh = lngPatiID & "_" & lngPageID
    net = mrsCdTmp.Fields!����
    
    mstrCdSql = "Select * from ������ҳ where ����=[1] And ����ID=[2] And ��ҳID=[3]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "������ҳ", TYPE_����, lngPatiID, lngPageID)
    
    yykb = ҽԺ���Ҵ���_����(mrsCdTmp!��Ժ����ID)
    
    yyzdys = ҽԺҽʦ����_����(mrsCdTmp!����ҽʦ)
    yysfydm = gstrPuser_id
    
    lxdh = Nvl(mrsCdTmp!��ϵ�˵绰, "")
    jtzz = Nvl(mrsCdTmp!��ͥ��ַ, "")
    Bz = Nvl(mrsCdTmp!��ע, "")
    iyjje = 0
    izhyj = 0
    zt = 99
    mstrCdSql = "select �������,������Ϣ From ������ Where ����ID=[1] And ��ҳID=[2] And �������>=1 and �������<=2 order by ������� desc"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(mstrCdSql, "��Ժ���", lngPatiID, lngPageID)
    
    If mrsCdTmp.RecordCount > 0 Then
        ryzd = Trim(Nvl(mrsCdTmp!������Ϣ, "��"))
    Else
        ryzd = "��"
    End If
    
    zt = ryappend(sbh, yyzyh, yykb, yyzdys, yysfydm, ryzd, lxdh, jtzz, Bz, net, iyjje, izhyj)
    Call WriteBusinessLOG("ryappend(�Ǽ���Ժ��Ϣ)", "sbh:" & sbh & ",yyzyh:" & yyzyh & _
                                                    ",yykb:" & yykb & ",yyzdys:" & yyzdys & _
                                                    ",yysfydm:" & yysfydm & ",ryzd:" & ryzd & _
                                                    ",lxdh:" & lxdh & ",jtzz" & jtzz & _
                                                    ",bz:" & Bz & ",net:" & net & _
                                                    ",iyjje:" & iyjje & ",izhyj:" & izhyj, zt)
                                                    
    
    Select Case zt
        Case 0
            ��Ժ�Ǽ�_���� = True
            gstrSQL = "zl_�����ʻ�_��Ժ(" & lngPatiID & "," & TYPE_���� & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "ҽ����Ժ")
        Case 1
            strMsg = "����������������Ϣ!"
        Case 2
            strMsg = "�ò����Ѿ���Ժ!"
        Case 3
            strMsg = "ҽ��סԺ���ظ������ܵǼ���Ժ!"
        Case 4
            strMsg = "�޸��籣����Ա!"
        Case 5
            strMsg = "�����ϴ������IC����������!"
        Case 99
            strMsg = "����!"
    End Select
        
    If zt <> 0 Then
        MsgBox strMsg, vbInformation, gstrSysName
        ��Ժ�Ǽ�_���� = False
    End If
            
End Function

Public Function ��ݱ�ʶ_����(bytType As Byte, lng����ID As Long) As String
'��clsInsure �� Identify  ���̵���
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
'���ù����嵥��˵��:
'     frmIdentify���� : ��������֤,���ز�����Ϣ

    Dim str������� As String
    Dim str����� As String
    Dim str���� As String
    Dim strIdeReturn As String  '���շ�����Ϣ
    
    If Not (bytType = 1 Or bytType = 0) Then Exit Function  '�������շѣ���Ժ�Ǽǲŵ���
    
    strIdeReturn = frmIdentify����.��ݱ�ʶ(bytType, lng����ID) '''����ӿ�����,readcard
    If strIdeReturn = "99" Then
        ��ݱ�ʶ_���� = ""
    Else
        ��ݱ�ʶ_���� = strIdeReturn
    End If
    
End Function

Public Sub ȡ������_����()
'��clsInsure �� IdentifyCancel  ���̵���
'����:
'���ù����嵥��˵��:(��)
End Sub

Public Function ҽ����ʼ��_����() As Boolean
'��clsInsure ��  InitInsure  ���̵���
'����: ҽ����ʼ��
'���ù����嵥��˵��:
'    �շ�Ա����_����: ��ѯ��ǰ��Ա�Ƿ���ҽ������ע��,���û�������,������Ա���
    ҽ����ʼ��_���� = True
    mblnInit = True
    gstrPuser_id = �շ�Ա����_����()
End Function

Public Function ��Ժ�Ǽǳ���_����(lngPatiID, lngPageID) As Boolean
'��clsInsure �� LeaveDelSwap ���̵���
'����: �α����˳�����Ժ
'���ù����嵥��˵��:
' (��)

Dim rsCd As New ADODB.Recordset

mstrCdSql = "select A.��Ժ����,A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
          "C.����,D.���� As ���ұ���,C.������� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
          "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
          "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]" & _
          " and C.����=" & TYPE_����

Set rsCd = zlDatabase.OpenSQLRecord(mstrCdSql, "���ղ���", lngPatiID, lngPageID)
If rsCd.EOF Then
    ��Ժ�Ǽǳ���_���� = False
    MsgBox "�ò���δͨ�������֤�����ܰ�������Ժ��", vbInformation, gstrSysName
    Exit Function
End If

��Ժ�Ǽǳ���_���� = True
gstrSQL = "zl_�����ʻ�_��Ժ(" & lngPatiID & "," & TYPE_���� & ")"
Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")

End Function

Public Function ��Ժ�Ǽ�_����(lngPatiID, lngPageID) As Boolean
'��clsInsure �� LeaveSwap ���̵���
'����: ����ҽ�����˳�Ժ
'���ù����嵥��˵��:
' (��)

    ��Ժ�Ǽ�_���� = True
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lngPatiID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
End Function

Public Function ���˱䶯��¼�ϴ�_����(lngPatiID, lngPageID) As Boolean
'��clsInsure �� ModiPatiSwap ���̵���
'����: �ϴ���λ��Ϣ��ҽ������
'���ù����嵥��˵��
'    cwappend: ��λ�Ǽ�
    Dim tmp As Double
    Dim yyzyh As String
    Dim yyczydm As String
    Dim cwh As String
    Dim djrq As String
    
    gstrSQL = "Select B.����||' '||A.����||'��' as ��λ��,to_char(��ʼʱ��,'YYYY-MM-DD') as ��ʼ����" & _
             " from ���˱䶯��¼ A,���ű� B " & _
            "where A.����ID=B.ID And A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��ֹʱ�� is null And A.���� is not null"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�䶯��¼", lngPatiID, lngPageID)
    
    Do Until mrsCdTmp.EOF
        yyzyh = lngPatiID & "_" & lngPageID
        yyczydm = �շ�Ա����_����()
        cwh = mrsCdTmp!��λ��
        djrq = mrsCdTmp!��ʼ����
        
        tmp = cwappend(yyzyh, yyczydm, cwh, djrq)
        Call WriteBusinessLOG("cwappend(��λ�䶯)", "yyzyh:" & yyzyh & ",yyczydm:" & yyczydm & ",cwh:" & cwh & "djrq:" & djrq, tmp)
        mrsCdTmp.MoveNext
    Loop
    ���˱䶯��¼�ϴ�_���� = True
End Function

Public Function �������_����(strҽ����) As Currency
'��clsInsure �� SelfBalance ���̵���
'����: ��ȡ�α����˸����ʻ����
'���ù����嵥��˵��:
' (��)

    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = "Select Nvl(�ʻ����,0) AS �����ʻ� From �����ʻ� " & _
              " Where ҽ����=[1] and ����=[2]"
              
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", strҽ����, TYPE_����)
    �������_���� = rsTemp!�����ʻ�
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function סԺ�������_����(lng����ID) As Boolean
'��clsInsure �� SettleDelSwap  ���̵���
'����:����ҽ����֧�ֽ������
On Error GoTo ErrH
    Err.Raise 9000, gstrSysName, "����ҽ����֧�ֽ������!"
    סԺ�������_���� = False
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ����_����(lng����ID) As Boolean
'��clsInsure �� SettleSwap ���̵���
'���ܣ����סԺ�Ǽǽ���
'���ù����嵥��˵��
'    �շ�Ա����_����:��ѯ��ǰ��Ա�Ƿ���ҽ������ע��,���û�������,������Ա���
'    cyappend       :д��Ժ��

    Dim lng����ID As Long, lng��ҳID As Long
    Dim blnOut As Boolean  '�Ƿ���;����
    Dim strNO As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMsg As String
    Dim tmp As Double
    
On Error GoTo ErrH

    m_bah = "" '������
    m_djh = "" '���ݺ�
    m_yysfy = �շ�Ա����_����()
    m_Bz = ""
    m_Bl = ""
    
'����������סԺ��ʼ������
    '��ȡ���ʵ���
    gstrSQL = "Select NO,����ID From ���˽��ʼ�¼ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ʵ���", lng����ID)
    strNO = "2" & rsTemp!NO
    lng����ID = rsTemp!����ID
    

    '��ȡ������ҳID����Ժ����
    gstrSQL = " Select A.��ҳID,A.��Ժ���� From ������ҳ A,������Ϣ B" & _
            " Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ҳID����Ժ����", lng����ID)
    blnOut = Not (IsNull(rsTemp!��Ժ����))
    lng��ҳID = rsTemp!��ҳID
    
    '2005 0520 ��ӳ�Ժ���
    gstrSQL = "select * From ������ Where ����ID=[1] And ��ҳID=[2] And �������=3"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������", lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
        m_cyzd = Trim(Nvl(rsTemp!������Ϣ, "��"))
    Else
        m_cyzd = "��"
    End If
    
    If m_cyzd = "��" Then
        gstrSQL = "select * From ������ Where ����ID=[1] And ��ҳID=[2] And �������=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������", lng����ID, lng��ҳID)
        If Not rsTemp.EOF Then
            m_cyzd = Trim(Nvl(rsTemp!������Ϣ, "��"))
        Else
            m_cyzd = "��"
        End If
        m_cyzd = InputBox("�������Ժ���", "��Ժ���¼��", m_cyzd)
    End If
    
    If Trim(m_cyzd) = "" Then m_cyzd = "��"
    
    '2005 0520 ��ӳ�Ժ���
    
    tmp = cyappend(m_yyzyh, m_bah, m_cyzd, m_djh, m_yysfy, _
                  m_Bz, m_Bl, m_Mxbbz, m_Qfx, m_Tcje, _
                  m_Zhzf, m_Grzf, m_Zhtk, m_Xjtk, m_Bxje, _
                  m_Bcbxje, m_Pbxje, m_Nbxje, m_Pbcbxje, m_Nbcbxje, _
                  m_Bcjs1, m_Bcjs2, m_Jbbcbxje, m_Gwybcbxje, m_Tsbxje, _
                  m_Pjbbcbxje, m_Njbbcbxje, m_Pgwybcbxje, m_Ngwybcbxje, m_Ptsbxje, _
                  m_Ntsbxje)
       
    Call WriteBusinessLOG("cyappend(д��Ժ��)", "m_yyzyh:" & m_yyzyh & ",m_bah:" & m_bah & ",m_cyzd:" & m_cyzd & ",m_djh:" & m_djh & ",m_yysfy:" & m_yysfy & "," & _
                  "m_Bz:" & m_Bz & ",m_Bl:" & m_Bl & ",m_Mxbbz:" & m_Mxbbz & ",m_Qfx:" & m_Qfx & ",m_Tcje" & m_Tcje & "," & _
                  "m_Zhzf:" & m_Zhzf & ",m_Grzf:" & m_Grzf & ",m_Zhtk:" & m_Zhtk & ",m_Xjtk:" & m_Xjtk & ",m_Bxje:" & m_Bxje & "," & _
                  "m_Bcbxje" & m_Bcbxje & ",m_Pbxje:" & m_Pbxje & ",m_Nbxje:" & m_Nbxje & ",m_Pbcbxje:" & m_Pbcbxje & ",m_Nbcbxje:" & m_Nbcbxje & "," & _
                  "m_Bcjs1:" & m_Bcjs1 & ",m_Bcjs2:" & m_Bcjs2 & ",m_Jbbcbxje:" & m_Jbbcbxje & ",m_Gwybcbxje:" & m_Gwybcbxje & ",m_Tsbxje:" & m_Tsbxje & "," & _
                  "m_Pjbbcbxje:" & m_Pjbbcbxje & ",m_Njbbcbxje:" & m_Njbbcbxje & ",m_Pgwybcbxje:" & m_Pgwybcbxje & ",m_Ngwybcbxje:" & m_Ngwybcbxje & ",m_Ptsbxje:" & m_Ptsbxje & "," & _
                  "m_Ntsbxje:" & m_Ntsbxje, tmp)

    '���汣�ս����¼
    '���Ը�=�󲡲���;�����Ը�=����Ա����
    Select Case tmp
        Case 0
            gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
                Year(zlDatabase.Currentdate()) & ",0,0,0,0,0,0,0,0," & m_rylf & "," & m_rxjzf & ",0," & _
                m_Tcje & "," & m_Bxje + m_Bcbxje + m_Jbbcbxje + m_Tsbxje & ",0," & m_Gwybcbxje & "," & m_Zhzf & _
                ",'" & lng����ID & "_" & lng��ҳID & "'," & lng��ҳID & "," & IIf(blnOut, 1, 0) & _
                ",'" & m_Bxje & ";" & m_Bcbxje & ";" & m_Jbbcbxje & ";" & m_Tsbxje & "')"
                 '  ^�������       ^���䱨��        ^�������䱨��      ^ ���ⱨ�����
            'gcnOracle.Execute gstrSQL, , adCmdStoredProc
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���±��ս����¼")
            סԺ����_���� = True
        Case 1
            strMsg = "�Ѿ���Ժ��"
        Case 2
            strMsg = "���粻ͨ��д������"
        Case 3
            strMsg = "������Ϣ��ȫ��"
        Case Else
            strMsg = "��������"
    End Select
    
    If tmp <> 0 Then
        סԺ����_���� = False
        Err.Raise 9000, gstrSysName, strMsg, vbInformation, gstrSysName
    End If
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �����ϴ�_����(int����, int״̬, str���ݺ�) As Boolean
'��clsInsure �� TranChargeDetail ���̵���
'����:סԺ�����ϴ���ɾ��
'���ù����嵥��˵��
'    ҽԺҽʦ����_����: ����ҽ�������Ƿ���ָ�����ҽʦ
'    jzappend         : ���ʷ������
'    jzdel            : ���ʷ���ɾ��
    Dim rsCd As New ADODB.Recordset
    Dim rs������ϸ As New ADODB.Recordset
    Dim lng����ID As Long
    
    Dim yyzyh As String 'סԺ��
    Dim yyxmdm As String '��Ŀ����
    Dim yyczydm As String '����Ա����
    Dim yycfys As String  '����ҽʦ
    Dim djh As String  '���ݺ�
    Dim Bz As String '��ע    ��ˮ��,Ψһ��ʶ
    Dim sl As String '����
    Dim je As String  '���
    Dim cfrq As String '��������  ��ʽ :YYYY-MM-DD
    
    Dim tmp As Double '���շ���ֵ
    Dim strMsg As String '�����ʾ��Ϣ
    
    ' �����¼״̬Ϊ1�ĵ��ݣ��и�����¼���������浥��
    
    If int״̬ = 1 Then
    gstrSQL = "Select distinct  A.����ID from סԺ���ü�¼ A,�����ʻ� B " & _
            "where A.����ID=B.����ID And A.��¼����=[1]" & _
            " And A.��¼״̬=[2] And A.NO=[3] " & _
            " And B.����=[4] And A.ʵ�ս��<0"
        Set rsCd = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ��и�����¼", int����, int״̬, str���ݺ�, TYPE_����)
        If Not rsCd.EOF Then
            MsgBox "��ҽ����֧�ָ�����¼���������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����NO��,��ȡ����ID
    gstrSQL = "Select distinct  A.����ID from סԺ���ü�¼ A,�����ʻ� B " & _
            "where A.����ID=B.����ID And A.��¼����=[1]" & _
            " And A.��¼״̬=[2] And A.NO=[3] " & _
            " And B.����=[4]"
    Set rsCd = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", int����, int״̬, str���ݺ�, TYPE_����)
    
    
    �����ϴ�_���� = True
    
    '>beging ����Ǽ��ʱ�,��Ҫ�����ʱ��ϵ�ҽ����������ϴ�
    Do Until rsCd.EOF
        lng����ID = rsCd!����ID
        'If int״̬ = 1 Then
            gstrSQL = "Select A.*,D.��Ŀ����,nvl(A.����,1)*nvl(A.����,0) as ����,A.ʵ�ս�� as ���," & _
                                      "nvl(A.ʵ�ս��,0)/(nvl(A.����,1)*nvl(A.����,0)) as �۸�,A.������ as ҽ��,C.���� as ��������,B.* " & _
                              " from סԺ���ü�¼ A,�����ʻ� B,���ű� C,����֧����Ŀ D " & _
                              " where A.NO=[1]" & _
                                    " And A.��¼����=[2]" & _
                                    " And A.��¼״̬=[3]" & _
                                    " And nvl(A.�Ƿ��ϴ�,0)=0 " & _
                                    " And A.����ID=B.����ID " & _
                                    " and B.����=[4]" & _
                                    " and A.��������ID=C.ID " & _
                                    " ANd A.����ID=[5]" & _
                                    " And A.�շ�ϸĿID=D.�շ�ϸĿID " & _
                                    " And D.����=[4]"
                                    
                                    
            Set rs������ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ", str���ݺ�, int����, int״̬, TYPE_����, lng����ID)
            Do Until rs������ϸ.EOF
                yyzyh = rs������ϸ!����ID & "_" & rs������ϸ!��ҳID
                yyxmdm = rs������ϸ!��Ŀ����
                yyczydm = gstrPuser_id
                yycfys = Nvl(ҽԺҽʦ����_����(rs������ϸ!ҽ��), "")
                djh = rs������ϸ!��¼���� & rs������ϸ!NO
                Bz = rs������ϸ!��¼���� & rs������ϸ!NO & "_" & rs������ϸ!���
                sl = rs������ϸ!����
                je = Format(rs������ϸ!���, "0.00")
                cfrq = Format(rs������ϸ!����ʱ��, "yyyy-MM-dd")
                '>>beging ��������
                If int״̬ = 1 Then
                    tmp = jzappend(yyzyh, yyxmdm, yyczydm, yycfys, djh, Bz, sl, je, cfrq)
                    Call WriteBusinessLOG("jzappend(���ʵǼ�)", "yyzyh:" & yyzyh & ",yyxmdm:" & yyxmdm & _
                                                              ",yyczydm:" & yyczydm & ",yycfys:" & yycfys & _
                                                               ",djh:" & djh & ",bz:" & Bz & _
                                                               ",sl:" & sl & ",je:" & je & _
                                                               ",cfrq:" & cfrq, tmp)
                
                End If
                '>> end ��������

                '>>beging ��������
                If int״̬ = 2 Then
                    tmp = jzdel(yyzyh, Bz)
                    Call WriteBusinessLOG("jzdel(����ɾ��)", "yyzyh:" & yyzyh & ",bz:" & Bz, tmp)
                    
                End If
                '>>end ��������

                If tmp = 0 Then
                    '�ϴ��ɹ�,����ϴ���־
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rs������ϸ!ID & "," & _
                            "0" & _
                            ",NULL,1,NULL,1,'" & Bz & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ���ֶ�")
                Else
                    strMsg = "�ϴ�����[" & rs������ϸ!NO & "]��" & rs������ϸ!��� & "����ϸʱ����" & vbCrLf & "��ϸ��Ϣ��" & vbCrLf
                    Select Case tmp
                    Case 1
                        If int״̬ = 1 Then
                            strMsg = strMsg & "û��סԺ��Ϣ?"
                        Else
                            strMsg = strMsg & "�Ҳ�����Ӧ������Ϣ?"
                        End If
                    Case 2
                         strMsg = strMsg & "�Ҳ�����Ӧ��Ŀ?"
                    Case Else
                         strMsg = strMsg & "����"
                    End Select
                    MsgBox strMsg, vbInformation, gstrSysName
                    'Exit Function
                End If
                rs������ϸ.MoveNext
            Loop
        'End If
        
    rsCd.MoveNext
    Loop
    '>end ����Ǽ��ʱ�,��Ҫ�����ʱ��ϵ�ҽ����������ϴ�
                                           
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, str���� As String) As String
'��clsInsure �� WipeoffMoney ���̵���
'����˵��:���סԺԤ���㹦��
'���ù����嵥������˵��:
'    jsbxf          :���㲡�˵�ǰ����������Ϣ
'    �����ϴ�_����  :סԺ�����ϴ���ɾ��
'    cycsh          :��Ժ��ʼ��

    Dim yyzyh As String 'סԺ��
    Dim qfx As Double '����
    Dim ylf As Double 'ҽ�Ʒ�
    Dim Tcje As Double '���ϱ���������ͳ���
    Dim bxje As Double '�����������
    Dim Bcbxje As Double '�߶�䱨�����
    Dim Jbbcbxje As Double '�������䱨�����
    Dim Gwybcbxje As Double '����Ա���䱨�����
    Dim Tsbxje As Double    '���ⱨ�����
    
    Dim sbh As String * 20   '�籣��
    Dim zhzt As String * 3 '�ʻ�״̬
    Dim zffs As String * 1 '�ʻ�֧����ʽ
    Dim net As String * 1  '����״̬
    Dim rzhye As Double  '�ʻ����
    Dim ryjje As Double  'Ԥ�ɽ��
    Dim rzhyjje As Double '�ʻ�Ԥ�ɽ��
    
    Dim tmp As Double '����ֵ
    Dim strMsg  As String '������Ϣ
    
    gstrSQL = "select * from ������Ϣ where ����ID=[1] And ����=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ϣ", lng����ID, TYPE_����)
    If mrsCdTmp.EOF Then
        MsgBox "��������ҽ������,����ִ�д˲���!", vbInformation, gstrSysName
        Exit Function
    End If
    
    '>beging ����δ�ϴ���¼
       ' ���δ�ϴ���¼��,��¼���� , ��¼״̬, NO ���ü����ϴ�
       gstrSQL = "Select distinct ��¼����,��¼״̬,NO From סԺ���ü�¼ A,�����ʻ� B,������Ϣ C " & _
                 " Where A.����ID=B.����ID And A.����ID=C.����ID And A.��ҳID=C.סԺ����" & _
                 " And nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 and A.���ʷ���=1 And A.����Ա���� is not null " & _
                 " AND A.ʵ�ս�� IS NOT NULL And B.����ID=[1] And B.����=[2]"
        Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���˷��ü�¼", lng����ID, TYPE_����)
        Do Until mrsCdTmp.EOF
            Call �����ϴ�_����(mrsCdTmp!��¼����, mrsCdTmp!��¼״̬, mrsCdTmp!NO)
            mrsCdTmp.MoveNext
        Loop
    
    '>end ����δ�ϴ���¼
    gstrSQL = "Select * from ������Ϣ where ����ID=[1] And ����=[2]"
    Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���˷��ü�¼", lng����ID, TYPE_����)
    yyzyh = lng����ID & "_" & mrsCdTmp!סԺ����
    '�������
    tmp = jsbxf(yyzyh, qfx, ylf, Tcje, bxje, Bcbxje, Jbbcbxje, Gwybcbxje, Tsbxje)
    Call WriteBusinessLOG("jsbxf(סԺ�������)", "yyzyh:" & yyzyh & ",qfx:" & qfx & ",ylf:" & ylf & ",tcje:" & Tcje & ",bxje:" & bxje & ",bcbxje:" & Bcbxje & ",jbbcbxje:" & Jbbcbxje & ",gwybcbxje:" & Gwybcbxje & ",tsbxje:" & Tsbxje, tmp)
    If tmp = 0 Then
        '���ܷ����Ƿ���ҽ���������
        
        gstrSQL = "Select sum(nvl(ʵ�ս��,0))-sum(nvl(���ʽ��,0)) as δ����� From סԺ���ü�¼ Where nvl(��¼״̬,0)<>0 and ���ʷ���=1 And ����ID= [1]"
        Set mrsCdTmp = zlDatabase.OpenSQLRecord(gstrSQL, "δ�����", lng����ID)
        
        If Val(Nvl(mrsCdTmp.Fields!δ�����, 0)) <> Val(Format(ylf, "0.00")) Then
            If MsgBox("ҽԺ�ķ����ܽ��(" & Nvl(mrsCdTmp.Fields!δ�����, 0) & ")��ҽ�����ĵķ����ܶ�(" & Val(Format(ylf, "#####0.00")) & ")���ȣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        '����ҽ��Ԥ���㲻�ܼ���������ʻ����,�����Ժˢ�����ܵõ�.
        '���ʴ�����(str���� = 1)����Ҫˢ����
        If str���� = 1 Then
            '�����ʼ��
                sbh = Space(20): zhzt = Space(3): zffs = Space(1): net = Space(1)
                m_yyzyh = Space(18): m_Mxbbz = Space(1)
                tmp = cycsh(sbh, zhzt, zffs, net, m_yyzyh, _
                            m_Mxbbz, rzhye, m_rylf, m_Qfx, m_Tcje, _
                            ryjje, rzhyjje, m_Bxje, m_Bcbxje, m_Pbxje, _
                            m_Nbxje, m_Pbcbxje, m_Nbcbxje, m_Bcjs1, m_Bcjs2, _
                            m_Zhzf, m_rxjzf, m_Zhtk, m_Xjtk, m_Jbbcbxje, _
                            m_Gwybcbxje, m_Tsbxje, m_Pjbbcbxje, m_Njbbcbxje, m_Pgwybcbxje, _
                            m_Ngwybcbxje, m_Ptsbxje, m_Ntsbxje)
            
                Call WriteBusinessLOG("crcsh(��Ժ��ʼ��)", "sbh:" & sbh & ",zhzt:" & zhzt & ", zffs:" & zffs & " , net:" & net & ", yyzyh:" & m_yyzyh & " ," & _
                            "mxbbz:" & m_Mxbbz & ", rzhye:" & rzhye & ", rylf:" & m_rylf & ", rqfx:" & m_Qfx & ", rtcje:" & m_Tcje & "," & _
                            "ryjje:" & ryjje & ", rzhyjje:" & rzhyjje & ", rbxje:" & m_Bxje & ", rbcbxje:" & m_Bcbxje & " , rpbxje:" & m_Pbxje & "," & _
                            "rnbxje:" & m_Nbxje & ", rpbcbxje:" & m_Pbcbxje & ", rnbcbxje:" & m_Nbcbxje & ", rbcjs1:" & m_Bcjs1 & ", rbcjs2:" & m_Bcjs2 & "," & _
                            "rzhzf:" & m_Zhzf & ", rxjzf:" & m_rxjzf & ", rzhtk:" & m_Zhtk & ", rxjtk:" & m_Xjtk & ", rjbbcbxje:" & m_Jbbcbxje & "," & _
                            "rgwybcbxje:" & m_Gwybcbxje & ", rtsbxje:" & m_Tsbxje & ", rpjbbcbxje:" & m_Pjbbcbxje & ", rnjbbcbxje:" & m_Njbbcbxje & ", rpgwybcbxje:" & m_Pgwybcbxje & "," & _
                            "rngwybcbxje:" & m_Ngwybcbxje & ", rptsbxje:" & m_Ptsbxje & ", rntsbxje:" & m_Ntsbxje, tmp)
 
                Select Case tmp
                    Case 0
                        סԺ�������_���� = "�����ʻ�;" & m_Zhzf & ";1"
                        סԺ�������_���� = סԺ�������_���� & "|ͳ�����;" & m_Bxje + m_Bcbxje + m_Jbbcbxje + m_Tsbxje & ";1"
                        סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & m_Gwybcbxje & ";1"
                    Case 1
                        strMsg = "������Ϣ�޸��籣��?"
                    Case 2
                        strMsg = "û�в��˵�סԺ��Ϣ?"
                    Case 3
                        strMsg = "���粻ͨ?"
                    Case 4
                        strMsg = "���ü������?"
                    Case Else
                        strMsg = " ��������?"
                End Select
        Else
            סԺ�������_���� = "�����ʻ�;" & 0 & ";1"
            סԺ�������_���� = סԺ�������_���� & "|ͳ�����;" & bxje + Bcbxje + Jbbcbxje + Tsbxje & ";1"
            סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & Gwybcbxje & ";1"
            
        End If
    End If
    
End Function


Public Function �շ�Ա����_����() As String
'����˵��:��ѯ��ǰ��Ա�Ƿ���ҽ������ע��,���û�������,������Ա���
'���ù����嵥��˵��:
'    finduser  :����ҽ�������Ƿ���ָ�������Ա
'    inuser    :���ҽʦ

Dim tmp As Double
Dim rsRy As New ADODB.Recordset '��Ա��
Dim blnReturn As Boolean '����ֵ

Dim str_SQL As String
str_SQL = "Select ������ѵ from ��Ա�� where id=[1]"
Set rsRy = zlDatabase.OpenSQLRecord(str_SQL, "��Ա��", UserInfo.ID)

If Nvl(rsRy.Fields!������ѵ, "9") <> "1" Then
    blnReturn = finduser(UserInfo.���)
    Call WriteBusinessLOG("finduser(������Ա)", "yyid:" & UserInfo.���, IIf(blnReturn, "True", "False"))
    If blnReturn = False Then
        tmp = inuser(UserInfo.���, UserInfo.����)
        Call WriteBusinessLOG("inuser(�����Ա)", "yyid:" & UserInfo.��� & ",user_name:" & UserInfo.����, tmp)
        str_SQL = "Update ��Ա�� set ������ѵ='1' where id=" & UserInfo.ID
        gcnOracle.Execute str_SQL
        gcnOracle.Execute "Commit"
    End If
End If
�շ�Ա����_���� = UserInfo.���

End Function



Public Function ҽԺҽʦ����_����(STR���� As String) As String
'����˵��:��ѯ�Ƿ���ָ��ҽʦ,���û�������,����ҽʦ���
'���ù����嵥��˵��:
'    findys  :����ҽ�������Ƿ���ָ�����ҽʦ
'    inys    :���ҽʦ
'    findkb  :����ҽ�������Ƿ���ָ����ſƱ�
'    inkb    :��ӿƱ�

    Dim rsTemp As New ADODB.Recordset
    Dim tmp As Double
    Dim strҽʦ���� As String
    Dim strҽʦ���� As String
    Dim lng��ԱID As Long
    Dim str�Ʊ��� As String
    Dim blnReturn As Boolean '����ֵ
    
    mstrCdSql = "Select ID,���,����,���˼�� from ��Ա�� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrCdSql, "��Ա��", STR����)
    lng��ԱID = rsTemp!ID
    strҽʦ���� = rsTemp!���
    strҽʦ���� = rsTemp!����
    
    If Nvl(rsTemp!���˼��, "9") <> "1" Then
    
        mstrCdSql = "select * from ���ű� where ID in (select ����ID from ������Ա where ȱʡ=1 and ��ԱID=[1])"
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrCdSql, "���ű�", lng��ԱID)
        
        If Nvl(rsTemp!λ��, "9") <> "1" Then
            blnReturn = findkb(rsTemp!����)
            Call WriteBusinessLOG("findkb(���ҿƱ�", "yykbbh:" & rsTemp!����, IIf(blnReturn, "True", "False"))
            If blnReturn = False Then
                tmp = inkb(rsTemp!����, rsTemp!����)
                Call WriteBusinessLOG("inkb(��ӿ���)", "yykbbh:" & rsTemp!���� & ",mc:" & rsTemp!����, tmp)
                mstrCdSql = "Update ���ű� set λ��='1' where id= " & rsTemp!ID
                gcnOracle.Execute mstrCdSql
                gcnOracle.Execute "Commit"
            End If
        End If
    
        str�Ʊ��� = rsTemp!����
        blnReturn = findys(strҽʦ����)
        Call WriteBusinessLOG("findys(����ҽʦ)", "yyysbh:" & strҽʦ����, IIf(blnReturn, "True", "False"))
        If blnReturn = False Then
            tmp = inys(strҽʦ����, str�Ʊ���, strҽʦ����)
            Call WriteBusinessLOG("inys(���ҽʦ)", "yyysbh:" & strҽʦ���� & ",yykbbh:" & str�Ʊ��� & ",xm:" & strҽʦ����, tmp)
            mstrCdSql = "Update ��Ա�� set ���˼��=1 where id=" & lng��ԱID
            gcnOracle.Execute mstrCdSql
            gcnOracle.Execute "Commit"
        End If
        
    End If
    ҽԺҽʦ����_���� = strҽʦ����
    
End Function

Public Function ҽԺ���Ҵ���_����(ByVal lng����ID As Long) As String
'����˵��:��ѯ�Ƿ���ָ������,���û�������,���ؿ��ұ��
'���ù����嵥��˵��:
'    findkb  :����ҽ�������Ƿ���ָ����ſƱ�
'    inkb    :��ӿƱ�
    Dim str�Ʊ��� As String
    Dim blnReturn As Boolean '����ֵ
    Dim tmp  As Double  '����ֵ
    
    Dim rsTemp As New ADODB.Recordset
    
    mstrCdSql = "select * from ���ű� where ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrCdSql, "���ű�", lng����ID)
    
    If Nvl(rsTemp!λ��, "9") <> "1" Then
        blnReturn = findkb(rsTemp!����)
        Call WriteBusinessLOG("findkb(���ҿƱ�", "yykbbh:" & rsTemp!����, IIf(blnReturn, "True", "False"))
        If blnReturn = False Then
            tmp = inkb(rsTemp!����, rsTemp!����)
            Call WriteBusinessLOG("inkb(��ӿ���)", "yykbbh:" & rsTemp!���� & ",mc:" & rsTemp!����, tmp)
            mstrCdSql = "Update ���ű� set λ��='1' where id= " & lng����ID
            gcnOracle.Execute mstrCdSql
        End If
    End If
    
    ҽԺ���Ҵ���_���� = rsTemp!����

End Function












