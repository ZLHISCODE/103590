Attribute VB_Name = "mdlInPatient"
Option Explicit 'Ҫ���������

Public gobjPatient As Object '���˹�����
Public gclsInsure As New clsInsure
Public gobjPublicPatient As Object  '������Ϣ�������� zlPublicPatient.clsPublicPatient
Public gobjPlugIn As Object    '�������zlPlugIn.clsPlugIn
Public gobjXWHIS As Object     '�����ӿڲ���zl9XWInterface.clsHISInner

'ϵͳ����--------------------------------
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"
Public gbytDec As Byte '���ý���С����λ��

Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gblnMyStyle As Boolean 'ʹ�ø��Ի����
Public gstrIme As String '�Զ��Ŀ������뷨
Public gbytCode As Byte '�������ɷ�ʽ��0-ƴ��,1-���,2-����
Public gblnXW As Boolean      'ϵͳ������������ҽѧӰ����Ϣϵͳרҵ��ӿڡ�
Public gobjPublicExpenseBillOperation As Object  '���ù����������û�ת����ת����

Public gint����������� As Integer
Public gintסԺ������� As Integer
Public gbln���ȷ������ȼ� As Boolean
Public gblnҽ��������ܳ�Ժ As Boolean 'ҽ���´��Ժҽ���������˳�Ժ
Public gblnҽ��������ܳ���Ԥ��Ժ As Boolean '���˳�Ժҽ�����������Ԥ��Ժ
Public gblnÿ��סԺ��סԺ�� As Boolean 'ÿ��סԺʹ���µ�סԺ��
Public gbln��Ժ���˲�׼��Ժ���� As Boolean '����δ��˵�����������,δ���˲�������г�Ժ

Public gbln��ԺԤ�� As Boolean '��Ժʱ��Ԥ����
Public gbln��Ժ���� As Boolean '��Ժʱ����￨
Public gbln��Ժ��� As Boolean '��Ժͬʱ���
Public gbyt���ʱ�� As Byte '0-��Ժʱ��,1-���ʱϵͳʱ��
Public gbln���ƿմ� As Boolean

'Public gblnShowCard As Boolean '�Ƿ���ʾ���Ŀ���
Public gblnCheckPass As Boolean 'ˢ��ʱҪ����������
Public gblnMultiBalance As Boolean  '���ŵ���ʹ�ö��ֽ��㷽ʽģʽ
   
'Public gbytCardNOLen As Byte '���￨�ų���
Public gbytPrepayLen As Byte 'Ʊ�ݺ��볤��
Public gblnPrepayStrict As Boolean '�Ƿ��ϸ����Ʊ��
'Public gblnMagcardStrict As Boolean '�Ƿ��ϸ����Ʊ��
Public gbyt��Ժʱ���δִ�� As Byte        '��Ժ�ͽ��ʳ�Ժʱ����Ƿ���δִ����Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gbytת��ʱ���δִ�� As Byte        'ת��ʱ����Ƿ���δִ����Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
'����30208 by lesfeng 2010-08-02 ���ֲ���22��32 ����154��155
Public gbyt��Ժʱ���ҩƷδִ�� As Byte    '�ڳ�Ժ���ʼ�������������г�Ժʱ�Ƿ��鲡�˵�δ��ҩƷ��Ŀ,0-�����,1-��鲢��ʾ,2-��鲢��ֹ
Public gbytת��ʱ���ҩƷδִ�� As Byte    '�ڲ������������ת��ʱ�Ƿ��鲡�˵�δ��ҩƷ��Ŀ:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
'���˺� ����:????    ����:2010-12-07 09:36:02
Public gintFeePrecision As Integer    '����С������
Public gstrFeePrecisionFmt As String '����С����ʽ:0.00000
'61347:������,2013-11-09
Public gbyt������˷�ʽ As Byte             '������˷�ʽ:0-δ��˲�������ʣ�ȱʡΪ0;1-���ʱ����������ú�ҽ��������ҽ�������ͷ��õ�����
'61492:������,2013-11-11
Public gbytת��ʱδ������ʵ��ݼ�� As Byte  'ת��ʱ�Ƿ��鲡�˴���δ��˵����ʵ���:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
'68953:������,2014-08-12
Public gbyt��Ժʱ���ڻ������ݼ�� As Byte    '��Ժʱ�Ƿ��鲡�˳�Ժʱ��֮����ڻ�������:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
'ҽ�����
Public gblnҩ�ƻ��۵� As Boolean
Public gbln�������۵� As Boolean
Public gblnִ�к���� As Boolean
'�ṹ����ַ
Public gbln���ýṹ����ַ As Boolean
Public gbln��ʾ���� As Boolean
Public gblnPatiByID As Boolean   'ͬһ���ֻ֤�ܶ�Ӧһ����������

'���ز���
Public gbln���� As Boolean '�Ƿ��������뵣����Ϣ
Public gblnSeekName As Boolean '�Ƿ�ͨ����������ģ������
Public gintNameDays As Integer 'ͨ������ģ����������
Public gstrԤ��ID As String   'Ʊ������ID
Public gbln���� As Long '���￨�����Լ��˷�ʽ��ȡ
Public gbln��ѡ���� As Boolean '��Ժʱ��ѡ����
Public gbln���ü��� As Boolean '����һ�ε���Ŀ�Ƿ�����Ժ�Ǽ���(δ��ס�����)
Public gblnת����ת���� As Boolean

Public gblnҽ�ƻ�������������¼�� As Boolean

Public gbytPrepayPrint As Byte ''0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ
Public gbytFPagePrint As Byte ''0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ
Public gbytWristletPrint As Byte ''0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ(������Ժ����)
Public gbytBabyWristletPrint As Byte ''0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ
Public gbytCourseWristletPrint As Byte ''0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ(�����������)

'��ʼ�� clsBase
Public gclsBase As New clsBase
Public Const ETO_OPAQUE = 2
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'��Ϣ�ṹ����
Public Enum enumXmlType
    xsString = 1
    xsNumber = 2
    xsDate = 3
    xsTime = 4
    xsDateTime = 5
End Enum


Public Enum EFun
    E��� = 0
    Eת�� = 1
    E���� = 2
    E���� = 3
    E��Ժ = 4
    EתΪסԺ = 5
    E���Ĵ�λ�ȼ� = 6
    E����������Ϣ = 7
    E�������Ǽ� = 8
    E������� = 9
    Eҽ������ѡ�� = 10
    E���� = 11
    E�޸ĳ�Ժʱ�� = 12
    E��λ�Ի� = 13
    Eתҽ��С�� = 14
    Eת���� = 15
    E�벡�� = 16
    E���˱�ע�༭ = 17
End Enum

Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29 '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    
End Enum
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
    g��������ģ�� = 5
    g����˽��ģ�� = 6
End Enum
'ϵͳ������Ϣ
'----------------------------------------------------------------------------------------------------------------------
Public Type SYSPARAM_INFO
    ���ý��С��λ�� As String
    �շ�������Ŀƥ�� As String
    ����Ʊ�ݺų��� As Integer
    �շ�Ʊ�ݺų��� As Integer
    ���￨���볤�� As Integer
    ���￨��ĸǰ׺ As String
    ���￨������ʾ As Boolean
    ��Ŀ����ƥ�䷽ʽ As Integer '0-˫��;1-����
    ϵͳ�� As Long
    ����ϵͳ�� As Long
    ϵͳ���� As String
    ��Ʒ���� As String
    ģ��� As Long
    ������ As String
    �շ�Ʊ�� As Integer
    ����Ʊ�� As Integer
    ����Ʊ���ϸ���� As Boolean
    �շ�Ʊ���ϸ���� As Boolean
    ����HIS���� As Byte
End Type

Public Enum ����Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum

'�ڲ�Ӧ��ģ��Ŷ���
Public Enum ENUM_INSIDE_PROGRAM
    P������λ���� = 1130
    P������Ժ���� = 1131
    P����������� = 1132
End Enum

'�ṹ����ַ���� 1-�����أ�2-����,3-��סַ,4-���ڵ�ַ,5-��ϵ�˵�ַ��6-��λ��ַ
Public Enum Enum_IX_ADDRESS
    E_IX_�����ص� = 1
    E_IX_���� = 2
    E_IX_��סַ = 3
    E_IX_���ڵ�ַ = 4
    E_IX_��ϵ�˵�ַ = 5
End Enum

Public ParamInfo As SYSPARAM_INFO

Public Function ExecPatiChange(ByVal bytFun As Byte, ByRef frmParent As Form, ByRef strPrivs As String, ParamArray arrPar() As Variant) As Boolean
'����:ִ�в��˱䶯��ع���
'����:bytFun:0-���,1-ת��
'     arrPar:���ݲ�ͬ�Ĺ��ܵ��ã����벻ͬ�Ĳ���
'            ���:����,����(��ס��Ŀ�괲λ,����Ϊ��),mlng��λ����ID,��Ʒ�ʽ(0-��Ժ��ƣ�1-ת�����)
'            ת��:����ID,��ҳId
'            ����:����ID,��ҳId,mbytInFun,mstrĿ�괲��
'            ��Ժ:����ID,��ҳId
'            תΪסԺ:����ID,��ҳId,סԺ��,����
'            ������λ�ȼ�:����ID,��ҳId,mstr����(��ǰ�����ȼ��Ĵ���)
'            �������:lng����ID, lng��ҳId, str����
    Dim strSql As String
    Dim blnReturn As Boolean
    Dim strTmp As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error Resume Next    '����Ϊ�����Form_Load��ִ��Unload me�����
    Select Case bytFun
    Case EFun.E���
        strTmp = CStr(arrPar(3))
        blnReturn = frmIn.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strTmp, Val(arrPar(4)), CStr(arrPar(5)), strPrivs)
        arrPar(3) = strTmp
    Case EFun.E�벡��
        strTmp = CStr(arrPar(3))
        blnReturn = frmCheckIn.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strTmp, Val(arrPar(4)), strPrivs)
        arrPar(3) = strTmp
    Case EFun.Eת����
        blnReturn = frmChangeUnit.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strPrivs)
    Case EFun.Eת��
        blnReturn = frmChange.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strPrivs)
    Case EFun.Eתҽ��С��
        'EFun.Eתҽ��С��, Me, mstrPrivs, mlngUnit, mrsBeds!����ID, mrsBeds!��ҳID, 3
        blnReturn = frmChangeGroup.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strPrivs)
    Case EFun.E����
        strTmp = CStr(arrPar(4))
        blnReturn = frmMove.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), Val(arrPar(3)), strTmp, CStr(arrPar(5)), strPrivs)
        arrPar(4) = strTmp
    Case EFun.E��λ�Ի�
        strTmp = CStr(arrPar(4))
        blnReturn = frmBedSwap.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), CStr(arrPar(3)), strTmp, strPrivs)
        arrPar(4) = strTmp
    Case EFun.E��Ժ
        blnReturn = frmOut.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), strPrivs)
    '����27392 by lesfeng 2010-01-14
    Case EFun.E�޸ĳ�Ժʱ��
        ExecPatiChange = frmModifOut.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), strPrivs)
    Case EFun.EתΪסԺ
        Dim strNote  As String, strסԺ�� As String
        'û��סԺ�������һ��
        '���� 26939 by lesfeng 2010-1-4 ��ִ�� ZL_���˱䶯��¼_תסԺ ʱû�н�סԺ�Ŵ���
        '77193:������,�޸�CStr(arrPar(2))=""����Ϊval(arrPar(2))=0,����ģ����ܽ�����val���´����סԺ��Ϊ0
        If Val(arrPar(2)) = 0 Then
            If gblnÿ��סԺ��סԺ�� = False Then
                strSql = " SELECT Nvl(a.סԺ��," & vbNewLine & _
                    "            (SELECT סԺ��" & vbNewLine & _
                    "             FROM ������ҳ" & vbNewLine & _
                    "             WHERE ����id = a.����id AND" & vbNewLine & _
                    "                   ��ҳid = (SELECT MAX(��ҳid) FROM ������ҳ WHERE ����id = a.����id AND סԺ�� IS NOT NULL))) סԺ��" & vbNewLine & _
                    " FROM ������Ϣ a" & vbNewLine & _
                    " WHERE ����id = [1]"
            
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡסԺ��", arrPar(0))
                If Not rsTemp.EOF Then
                    strסԺ�� = Nvl(rsTemp!סԺ��)
                End If
                If strסԺ�� = "" Then
                    strסԺ�� = zlDatabase.GetNextNo(2)
                    strNote = "�����۲��� " & CStr(arrPar(3)) & " תΪסԺ����֮ǰ������Ϊ�ò���ȷ��һ��סԺ�š�"
                    If Not frmInput.InputVal(frmParent, "סԺ��", strNote, strסԺ��, 1, 10, False, InStr(strPrivs, ";�޸�סԺ��;") <> 0) Then Exit Function
                End If
            Else
                strסԺ�� = zlDatabase.GetNextNo(2)
                strNote = "�����۲��� " & CStr(arrPar(3)) & " תΪסԺ����֮ǰ������Ϊ�ò���ȷ��һ��סԺ�š�"
                If Not frmInput.InputVal(frmParent, "סԺ��", strNote, strסԺ��, 1, 10, False, InStr(strPrivs, ";�޸�סԺ��;") <> 0) Then Exit Function
            End If
        Else
            strסԺ�� = CStr(arrPar(2))
        End If
        
        On Error GoTo errH
        strSql = "ZL_���˱䶯��¼_תסԺ(" & Val(arrPar(0)) & "," & Val(arrPar(1)) & "," & strסԺ�� & ")"
        zlDatabase.ExecuteProcedure strSql, App.ProductName
        gblnOK = True
        blnReturn = gblnOK
    Case EFun.E���Ĵ�λ�ȼ�
        blnReturn = frmBedLevel.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), CStr(arrPar(2)))
    Case EFun.E����������Ϣ
        blnReturn = frmEditPati.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), strPrivs)
    Case EFun.E�������Ǽ�
        blnReturn = frmBabyReg.ShowMe(Val(arrPar(0)), Val(arrPar(1)), strPrivs, frmParent)
    Case EFun.E���˱�ע�༭
        blnReturn = frmMemo.ShowMe(frmParent, Val(arrPar(0)), Val(arrPar(1)), strPrivs)
    Case EFun.E�������
             
        If MsgBox("��ȷ��Ҫ��[" & CStr(arrPar(2)) & "]��δ����ð���ǰ�ѱ�������?" & vbCrLf & vbCrLf & _
            "�������������˵�ǰ�ѱ��Ӧ���Żݱ��ʶ�δ��������½��д��ۼ���!", vbInformation + vbYesNo + vbDefaultButton1, App.ProductName) = vbNo Then
            Exit Function
        End If
        
        On Error GoTo errH
        strSql = "Zl_����δ�����_Recalc(" & Val(arrPar(0)) & "," & Val(arrPar(1)) & ")"
        zlDatabase.ExecuteProcedure strSql, App.ProductName
        gblnOK = True
        blnReturn = gblnOK
    Case EFun.Eҽ������ѡ��
        Call gclsInsure.ChooseDisease(Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)))
        blnReturn = True
    Case EFun.E����
        blnReturn = ExecUndo(frmParent, strPrivs, Val(arrPar(0)), Val(arrPar(1)), Val(arrPar(2)), Val(arrPar(3)), CStr(arrPar(4)))
        
    End Select
    ExecPatiChange = blnReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecUndo(ByRef frmParent As Form, ByVal strPrivs As String, ByVal lngUnit As Long, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal int���� As Integer, ByVal strType As String) As Boolean
    Dim str���� As String, strBeds As String, str����λ As String, str�ȼ�IDs As String
    Dim str��λ�� As String
    Dim strPreBed As String, strBed As String, strInfo As String
    Dim lngBeSwap����ID As Long, lngBeSwap��ҳID As Long
    Dim blnDie As Boolean, blnUndoOut As Boolean, blnUndoPreOut As Boolean
    Dim bln���� As Boolean, strSql As String, strUndoBeds As String
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean, blnSwapBed As Boolean
    Dim rsBeSwap As ADODB.Recordset
    Dim arrSQL() As String, intLoop As Integer
    
    Dim clsMipModule As zl9ComLib.clsMipModule
    Dim clsXML As zl9ComLib.clsXML
    Dim rsUndoBegin As ADODB.Recordset, rsUndoEnd As ADODB.Recordset
    Dim rsDeptOper As New ADODB.Recordset '������Ա��Ϣ
    Dim rsBedChange As New ADODB.Recordset
    Dim strBeforBed As String, strAfterBed As String
    Dim lngNextPati As Long, lngNextPage As Long '��λ�Ի��ڶ�������
    Dim bln��λ�Ի� As Boolean
        Dim colSQL As New Collection, i As Long, strSQLTmp As String, rsPati As Recordset
    
    blnUndoOut = strType = "��Ժ"
    blnUndoPreOut = strType = "Ԥ��Ժ"
    
    On Error GoTo errH
    
    '�����������֮ǰ���δִ����Ŀ
    If InStr(strType, "����") > 0 Then
        If gbytת��ʱ���δִ�� <> 0 Then
    
            strInfo = ExistWaitExe(lng����ID, lng��ҳID)
            If strInfo <> "" Then
                If gbytת��ʱ���δִ�� = 1 Then
                    If MsgBox("�ò��˴�����δִ����ɵ����ݣ�" & _
                        vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ����" & strType & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "�ò��˴�����δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "��������" & strType & "��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    
        If gbytת��ʱ���ҩƷδִ�� <> 0 Then
            strInfo = ExistWaitDrug(lng����ID, lng��ҳID)
            If strInfo <> "" Then
                If gbytת��ʱ���ҩƷδִ�� = 1 Then
                    If MsgBox("�ò���" & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ����" & strType & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "�ò���" & strInfo & vbCrLf & vbCrLf & "��������" & strType & "��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        '61429:������,2013-11-11,ת��ʱ����δ��˵��ݼ��
        If gbytת��ʱδ������ʵ��ݼ�� <> 0 Then
            strInfo = ""
            strInfo = ExistWaitQuittance(lng����ID, lng��ҳID)
            If strInfo <> "" Then
                If gbytת��ʱδ������ʵ��ݼ�� = 1 Then
                    If MsgBox("�ò���" & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ����" & strType & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "�ò���" & strInfo & vbCrLf & vbCrLf & "��������" & strType & "��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        '74262:������,2014-06-23,�������:��ֹʱ�� Is Null
        gstrSQL = "Select 1" & vbNewLine & _
                " From ����ҽ����¼ a" & vbNewLine & _
                " Where ����id = [1] And ��ҳid = [2] And ҽ��״̬ Not In (4, 8, 9) And Exists" & vbNewLine & _
                " (Select 1" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id = a.����id And ��ҳid = a.��ҳid And ��ʼԭ�� = 15 And ��ʼʱ�� < a.����ʱ�� And ��ֹʱ�� Is Null) And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ת����ʱ�ж��Ƿ������Чҽ��", lng����ID, lng��ҳID)
        
        If Not rsTmp.EOF Then
            MsgBox "�ò����ڵ�ǰ����������Чҽ������������" & strType & "��"
            Exit Function
        End If
    End If
    
    If InStr(strType, "����ȼ��䶯") > 0 Then
        gstrSQL = "SELECT 1" & vbNewLine & _
                    "FROM ���˱䶯��¼ a, ����ҽ����¼ b, �����շѹ�ϵ c" & vbNewLine & _
                    "WHERE a.����id = b.����id AND a.��ҳid = b.��ҳid AND a.����ȼ�id = c.�շ���Ŀid AND b.������Ŀid = c.������Ŀid AND a.����id = [1] AND" & vbNewLine & _
                    "      a.��ҳid = [2] AND a.��ʼԭ�� = 6 AND a.��ֹԭ�� IS NULL AND a.��ֹʱ�� IS NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��������ȼ��ж��Ƿ���ڻ���ȼ�ҽ��", lng����ID, lng��ҳID)
        
        If Not rsTmp.EOF Then
            MsgBox "�ò��˴��ڻ���ȼ�ҽ������������" & strType & "��"
            Exit Function
        End If

    End If

    If InStr(strType, "��ס") > 0 And InStr(strPrivs, "д�����������") = 0 Then '�������/����ת�����,����Ƿ�����д����Ժ/ת��ʱ��Ҫ�Ĳ���
        gstrSQL = "Select Count(ID) ��¼ From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ֹʱ�� Is Null And ��ʼԭ�� = 3"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ת����Ƽ�¼", lng����ID, lng��ҳID)
        
        gstrSQL = "Select Count(B.�ļ�id) ����" & vbNewLine & _
                    "From ����ʱ��Ҫ�� A, ���Ӳ�����¼ B" & vbNewLine & _
                    "Where Instr([3],A.�¼�)>0 And A.��дʱ�� > 0 And A.�ļ�id = B.�ļ�id And B.����id = [1] And B.��ҳid = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���д������", lng����ID, lng��ҳID, IIf(rsTmp!��¼ = 0, "��Ժ,�״���Ժ,�ٴ���Ժ", "ת��"))
        If rsTmp!���� > 0 Then
            MsgBox "�ò�������д����Ժ/ת��ʱ��Ҫ��Ĳ�������ֹ������ƣ�", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    '����ǰ��λ�ȼ����
    If blnUndoOut Or InStr(strType, "����") > 0 Or InStr(strType, "��ס") > 0 Then '������Ժ������������ת�����,�������Ҳ�ᱻ�жϵ���ȡ��������
        gstrSQL = "Select Distinct a.����, a.�ȼ�id ԭס�ȼ�, b.�ȼ�id ���еȼ�" & vbNewLine & _
                "From (Select a.����, a.��λ�ȼ�id �ȼ�id, a.����id" & vbNewLine & _
                "       From ���˱䶯��¼ a, ���˱䶯��¼ b" & vbNewLine & _
                "       Where a.����id = [1] And a.��ҳid = [2] And a.����id = b.����id And a.��ҳid = b.��ҳid And" & vbNewLine & _
                "             (b.��ֹʱ�� Is Null And b.��ʼԭ�� In(3,4) And a.��ֹʱ�� = b.��ʼʱ�� Or" & vbNewLine & _
                "             b.��ֹԭ�� = 1 And a.��ֹԭ�� = b.��ֹԭ��)) a, ��λ״����¼ b" & vbNewLine & _
                "Where Nvl(a.����, 0) = b.���� And a.����id = b.����id"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ϴ���ס��λ�仯", lng����ID, lng��ҳID)
        Do Until rsTmp.EOF
            If Nvl(rsTmp!ԭס�ȼ�, 0) <> Nvl(rsTmp!���еȼ�, 0) Then
                strUndoBeds = strUndoBeds & Nvl(rsTmp!����, 0) & " "
            End If
            rsTmp.MoveNext
        Loop
        If Trim(strUndoBeds) <> "" Then
            If MsgBox("��λ " & strUndoBeds & "�ȼ��벡���ϴ���סʱ�ĵȼ���һ��" & vbCrLf & "�Ƿ����(�������Զ�����Ϊ�ϴ���סʱ�ĵȼ�)��", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        End If
    End If

    '������Ժ
    '���ܴ�������************
    If blnUndoOut Then
        If GetOutState(lng����ID, lng��ҳID) = "����" Then blnDie = True
        bln���� = True
        Set rsTmp = GetMoneyInfo(lng����ID, , , , , , , lng��ҳID)
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 0 Then bln���� = (rsTmp!������� = 0)
        End If
    
        'Ȩ���ж�
        If HavedInCost(lng����ID, lng��ҳID) Then '����ò���"���峷����Ժ"Ȩ�޿���
            If bln���� And InStr(strPrivs, "���峷����Ժ") = 0 Then
                MsgBox "�ó�Ժ���˵ķ����ѽ��壬��û��Ȩ�޽��ò��˳�����Ժ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '�ѱ�Ŀ�����˵Ĳ���������Ժ
        If HaveCatalogue(lng����ID, lng��ҳID) Then
            MsgBox "�ò��˱���סԺ�Ĳ����Ѿ���Ŀ������������Ժ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        'ҽ�������ж�
        If int���� <> 0 Then
            If Not gclsInsure.GetCapability(support������Ժ, lng����ID, int����) Then
                MsgBox "���ղ��˲��ܳ�����Ժ��", vbInformation, gstrSysName
                Exit Function
            End If
            'ȥ����ҽ������ƥ����
        End If
    ElseIf int���� <> 0 Then
        If Not gclsInsure.GetCapability(support��������ѽ��ʵļ��ʵ���, lng����ID, int����) Then
            str���� = "1"   '����Զ����ʵķ����Ƿ��ѱ�����
        End If
    End If
    
    '��������˲����Ƿ���������Ժ����Ԥ��Ժ
    If blnUndoPreOut Or blnUndoOut Then
        If InStr(strPrivs, "�����˳�����Ժ") = 0 Then
            If CheckAudited(lng����ID, lng��ҳID) Then
                MsgBox "�ò��˷�������ˣ���������" & IIf(blnUndoPreOut, "Ԥ", "") & "��Ժ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '61347:������,2013-11-09,������˵ķ����Ѿ�������,���ҷ�����˷�ʽΪ"���ʱ����������ú�ҽ��"������������Ժ
        If gbyt������˷�ʽ = 1 And blnUndoOut = True Then
            If CheckAudited(lng����ID, lng��ҳID, 2) Then
                MsgBox "�ò��˷�����������,���ҷ������ʱ����������ú�ҽ��������������Ժ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If (blnUndoPreOut) Then
            '43579:����Ԥ��Ժʱ��飬��ϵͳ��������Ϊ���³�Ժҽ����׼��Ժ��ʱ���������ڽ����ϳ���Ԥ��Ժ��ֻ��ͨ�����˳�Ժҽ���ķ��ͷ�ʽ������Ԥ��Ժ����
'            If (gblnҽ��������ܳ�Ժ) Then
'                MsgBox "ֻ��ͨ�����˳�Ժҽ�����͵ķ�ʽ������Ԥ��Ժ��", vbInformation, gstrSysName
'                Exit Function
'            End If
            '--55791:������,2012-11-13,���˳�Ժҽ�����ܳ�����Ժ
             '����Ԥ��Ժʱ�����������ɳ�Ժҽ�������ĳ�Ժ������Ҫ���ݲ���'���˳�Ժҽ�����ܳ�����Ժ'�����Ƿ����ֱ��ȡ��Ԥ��Ժ����������ɳ�Ժҽ�������ĳ�Ժ�������ֱ�ӳ�����
            If Checkҽ���´��Ժҽ��(lng����ID, lng��ҳID) And gblnҽ��������ܳ���Ԥ��Ժ = True Then
                MsgBox "ֻ��ͨ������ҽ���ķ�ʽ������Ԥ��Ժ��", vbInformation, gstrSysName
                Exit Function
            End If
            
        End If
    End If
        
    If blnUndoOut And blnDie Then
        If MsgBox("�ò��˳�Ժʱ�ѵǼ�Ϊ����,ȷʵҪ������Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If MsgBox("�ò�������������" & strType & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    
    If strType Like "*תΪסԺ����*" Then
        If lng��ҳID = 1 Then
            If MsgBox("Ҫͬʱ����ò��˵�סԺ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then str���� = "1"
        Else
            str���� = ""
        End If
    End If
    
    '��Ժǰ��Ժ�Ĵ�λ�Ƿ�ռ��,��ռ���򷵻ر�ռ�ô���ԭ��ס��
    If blnUndoOut Then
        strBeds = GetUsedBeds(lng����ID, lng��ҳID, str�ȼ�IDs)
        If strBeds <> "" Then
            If MsgBox("����ԭסԺ��λ " & strBeds & "�ǿմ����Ƿ�����������λ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        
            Call ExecPatiChange(EFun.E����, frmParent, strPrivs, lngUnit, lng����ID, lng��ҳID, 2, str�ȼ�IDs, strBeds)
            If Not gblnOK Then Exit Function
            'str�ȼ�IDs-��������ס�Ĵ���
            strBeds = str�ȼ�IDs
            str����λ = Split(str�ȼ�IDs, ",")(0)
        End If
    End If
    '����28386 by lesfeng 2010-03-06 ���� strType = ��ơ���Ժ��ơ�ת����ƣ�ת�ƣ�����������λ�ǼǱ䶯������ǼǱ䶯������ҽʦ�Ķ������λ�ʿ�ı䡢תΪסԺ���ˡ�Ԥ��Ժ������ҽʦ�䶯������ҽʦ�䶯�������䶯�Լ���Ժ
    ReDim Preserve arrSQL(0)
    arrSQL(UBound(arrSQL)) = "zl_���˱䶯��¼_Undo(" & lng����ID & "," & lng��ҳID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "','" & str���� & "','" & strBeds & "','" & str����λ & "','" & strType & "')"
    
    '��ȡ����ǰ���������Ϣ
    Select Case strType
    Case "��Ժ"
        strSql = " Select b.����, b.�Ա�, b.סԺ��, a.Id �䶯id,����Id ����ǰ����ID,����Id ����ǰ����Id,b.��Ժ����,A.��ʼʱ��" & _
            "   From ���˱䶯��¼ a, ������ҳ b" & _
            "   Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.��ֹԭ�� = 1 And Nvl(a.���Ӵ�λ, 0) = 0 And b.����id = [1] And b.��ҳid = [2]"
        Set rsUndoBegin = zlDatabase.OpenSQLRecord(strSql, "�����䶯ǰ", lng����ID, lng��ҳID)
    Case "ת��", "ת����"
        strSql = "Select b.����, b.�Ա�, b.סԺ��, a.Id �䶯id,����Id ����ǰ����ID,����Id ����ǰ����Id,b.��Ժ����,A.��ʼʱ��" & _
            "   From ���˱䶯��¼ a, ������ҳ b" & _
            "   Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.��ʼԭ�� = [3] And a.��ʼʱ�� Is Null And Nvl(a.���Ӵ�λ, 0) = 0 And ��ֹʱ�� Is Null And b.����id = [1] And b.��ҳid = [2]"
        Set rsUndoBegin = zlDatabase.OpenSQLRecord(strSql, "�����䶯ǰ", lng����ID, lng��ҳID, IIf(strType = "ת��", 3, 15))
    Case Else
        strSql = "Select b.����, b.�Ա�, b.סԺ��, a.Id �䶯id,����Id ����ǰ����ID,����Id ����ǰ����Id,b.��Ժ����,A.��ʼʱ��" & _
            "   From ���˱䶯��¼ a, ������ҳ b" & _
            "   Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.��ʼʱ�� IS NOT NULL And Nvl(a.���Ӵ�λ, 0) = 0 And ��ֹʱ�� Is Null And b.����id = [1] And b.��ҳid = [2]"
        Set rsUndoBegin = zlDatabase.OpenSQLRecord(strSql, "�����䶯ǰ", lng����ID, lng��ҳID)
    End Select
    '����Ƿ�ȡ����λ�Ի�(��Ϣ������Ҫ)
    bln��λ�Ի� = False
    If strType = "����" And rsUndoBegin.RecordCount > 0 Then
        strSql = "Select a.����,NVL(a.���Ӵ�λ,0) ���Ӵ�λ , c.��Ժ����, c.��Ժ����id, c.��ǰ����id,b.����id" & vbNewLine & _
            "    From ���˱䶯��¼ A, ��λ״����¼ B, ������ҳ C" & vbNewLine & _
            "    Where a.����id = [1] And a.��ҳid = [2] And ��ֹԭ�� = 4 And a.����id = b.����id And a.����id = c.����id And" & vbNewLine & _
            "          a.��ҳid = c.��ҳid And a.���� = b.����" & vbNewLine & _
            " Order By a.��ֹʱ�� Desc,NVL(A.���Ӵ�λ,0) DESC, a.��ʼʱ�� Desc"
         Set rsBedChange = zlDatabase.OpenSQLRecord(strSql, "", lng����ID, lng��ҳID)
         If rsBedChange.RecordCount > 0 Then
            strBeforBed = Nvl(rsBedChange!����)
            strAfterBed = Nvl(rsBedChange!��Ժ����)
            If (Not IsNull(rsBedChange!����ID)) And Val(Nvl(rsBedChange!����ID, 0)) <> lng����ID And Nvl(rsBedChange!���Ӵ�λ, 0) = 0 Then
                strSql = "Select a.����id, c.��ҳid, a.����,NVL(a.���Ӵ�λ,0) ���Ӵ�λ, c.��Ժ����" & vbNewLine & _
                    " From ���˱䶯��¼ a, ������ҳ c" & vbNewLine & _
                    " Where a.����id = [1] And a.����id = c.����id And a.��ҳid = c.��ҳid And c.��Ժ���� Is Null And a.��ֹԭ�� = 4" & vbNewLine & _
                    " Order By ��ֹʱ�� Desc,NVL(A.���Ӵ�λ,0) DESC, ��ʼʱ�� Desc"
                Set rsBedChange = zlDatabase.OpenSQLRecord(strSql, "", rsBedChange!����ID, Nvl(rsBedChange!����))
                If rsBedChange.RecordCount > 0 Then
                    If Nvl(rsBedChange!����ID, 0) <> 0 And Nvl(rsBedChange!��ҳID, 0) <> 0 And strBeforBed = Nvl(rsBedChange!��Ժ����, 0) And strAfterBed = Nvl(rsBedChange!����) And Nvl(rsBedChange!���Ӵ�λ, 0) = 0 Then
                        lngNextPati = Val(rsBedChange!����ID)
                        lngNextPage = Val(rsBedChange!��ҳID)
                        
                        strSql = "Select b.����, b.�Ա�, b.סԺ��, a.Id �䶯id,����Id ����ǰ����ID,����Id ����ǰ����Id,b.��Ժ����" & _
                        "   From ���˱䶯��¼ a, ������ҳ b" & _
                        "   Where a.����id = b.����id And a.��ҳid = b.��ҳid And A.��ʼԭ�� = [3] And a.��ʼʱ�� IS NOT NULL And Nvl(a.���Ӵ�λ, 0) = 0 And ��ֹʱ�� Is Null And b.����id = [1] And b.��ҳid = [2]"
                        Set rsBedChange = zlDatabase.OpenSQLRecord(strSql, "�����䶯ǰ", lngNextPati, lngNextPage, 4)
                        bln��λ�Ի� = rsBedChange.RecordCount > 0
                    End If
                End If
            End If
         End If
    End If
        
    '����ʱδִ�з�����Ҫת��ԭ����
    If strType = "ת������ס" Then
        If CreatePublicExpenseBillOperation() And gblnת����ת���� Then
            strSQLTmp = "Select ID, ����id" & vbNewLine & _
                    "From ���˱䶯��¼" & vbNewLine & _
                    "Where ����id = [1] And ��ҳid = [2] And Nvl(���Ӵ�λ, 0) = 0 And ��ֹԭ�� = 15 And" & vbNewLine & _
                    "      ��ֹʱ�� = [3]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQLTmp, App.ProductName, lng����ID, lng��ҳID, CDate(rsUndoBegin!��ʼʱ��))
            If rsPati.RecordCount > 0 Then
                If gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(frmParent, 1, lng����ID, lng��ҳID, Val(rsPati!ID & ""), Val(rsUndoBegin!����ǰ����ID & ""), Val(rsPati!����ID & ""), colSQL) = False Then Exit Function
            End If
        End If
    End If
        
    gcnOracle.BeginTrans: blnTrans = True
    'ת��������ʱ������ִ��
    For i = 1 To colSQL.Count
        zlDatabase.ExecuteProcedure colSQL(i), App.ProductName
    Next
    For intLoop = LBound(arrSQL) To UBound(arrSQL)
        zlDatabase.ExecuteProcedure arrSQL(intLoop), App.ProductName
    Next
    
    '������Ժʱҽ������
    If blnUndoOut And int���� <> 0 Then
        If Not gclsInsure.LeaveDelSwap(lng����ID, lng��ҳID, int����) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    ExecUndo = True

    '��ɲ��˱䶯����س�������������Ϣ
    On Error Resume Next
    '������Ϣ����
    Set clsMipModule = New zl9ComLib.clsMipModule
    Call clsMipModule.InitMessage(glngSys, P�����������, strPrivs, frmParent.hWnd)
    Call AddMipModule(clsMipModule)
    Set clsXML = New zl9ComLib.clsXML
    
    If clsMipModule.IsConnect = True Then
        
        strSql = "Select d.Id ����id, d.���� ��������, p.Id ��Աid, p.���� ��Ա����" & vbNewLine & _
            "    From ��Ա�� p, ������Ա r, ���ű� d" & vbNewLine & _
            "    Where p.Id = r.��Աid And r.����id = d.Id"
        Call zlDatabase.OpenRecordset(rsDeptOper, strSql, "��ȡ������Ա��Ϣ")
PreNext:
        If strType = "ת����ס" Or strType = "ת������ס" Then
            strSql = "Select  a.����id ��������id, a.����id ���������id, a.����ҽʦ, a.����ҽʦ, a.����ҽʦ, a.����, c.���� �������, a.���λ�ʿ, B.��Ժ���� ����" & vbNewLine & _
                "   From ���˱䶯��¼ a, ������ҳ b, ���� c" & vbNewLine & _
                "   Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.��ʼԭ�� = [3] And a.��ʼʱ�� Is Null And Nvl(a.���Ӵ�λ, 0) = 0 And ��ֹʱ�� Is Null And" & vbNewLine & _
                "      a.���� = c.����(+) And b.����id = [1] And b.��ҳid = [2]"
            Set rsUndoEnd = zlDatabase.OpenSQLRecord(strSql, "�����䶯��", lng����ID, lng��ҳID, IIf(strType = "ת����ס", 3, 15))
        Else
            strSql = "Select a.����Id ��������ID,a.����Id ���������Id,a.����ҽʦ,a.����ҽʦ,a.����ҽʦ,a.����,c.���� �������,a.���λ�ʿ,B.��Ժ���� ����" & _
                "   From ���˱䶯��¼ a, ������ҳ b,���� C" & _
                "   Where  a.����id = b.����id And a.��ҳid = b.��ҳid And a.��ʼʱ�� IS NOT NULL And Nvl(a.���Ӵ�λ, 0) = 0 And ��ֹʱ�� Is Null And a.���� = c.����(+) And b.����id = [1] And b.��ҳid = [2]"
            Set rsUndoEnd = zlDatabase.OpenSQLRecord(strSql, "�����䶯��", lng����ID, lng��ҳID)
        End If
        
        clsXML.ClearXmlText '��������е�XML
        '--������Ϣ��װ
        '������Ϣ
        clsXML.AppendNode "in_patient"
        'patient_id      ����id  1   N
        clsXML.appendData "patient_id", lng����ID, xsNumber  '����ID
        'page_id     ��ҳid  1   N
        clsXML.appendData "page_id", lng��ҳID, xsNumber '��ҳID
        'patient_name        ����    1   S
        clsXML.appendData "patient_name", Nvl(rsUndoBegin!����), xsString '����
        'patient_sex     �Ա�    0..1    S
        clsXML.appendData "patient_sex", Nvl(rsUndoBegin!�Ա�), xsString '�Ա�
        'in_number       סԺ��  1   S
        clsXML.appendData "in_number", Nvl(rsUndoBegin!סԺ��), xsString 'סԺ��
        clsXML.AppendNode "in_patient", True
        
        'change_cancel 1
        clsXML.AppendNode "change_cancel"
        'change_id       �䶯id  1   N
        clsXML.appendData "change_id", rsUndoBegin!�䶯ID, xsNumber
        'cancel_kind     ������ʽ    1   S       ��������λ�Ի���ת�ơ�ת��������Ժ����Ժ��ס����ס��ת����ס����λ�ȼ��䶯������ȼ��䶯������ҽʦ�ı䡢���λ�ʿ�ı䡢תΪסԺ���ˡ�Ԥ��Ժ������ҽʦ�䶯������ҽʦ�䶯�������䶯��תҽ��С�顢ת������ס
        clsXML.appendData "cancel_kind", strType, xsString
        'before_area_id      �����䶯ǰ����id    0..1    N
        If Nvl(rsUndoBegin!����ǰ����ID, 0) <> 0 Then
            clsXML.appendData "before_area_id", Nvl(rsUndoBegin!����ǰ����ID, 0), xsNumber
        End If
        'before_dept_id      �����䶯ǰ����Id    0..1    N
        clsXML.appendData "before_dept_id", Nvl(rsUndoBegin!����ǰ����Id, 0), xsNumber
        'after_area_id       �����䶯����id    0..1    N
        If Nvl(rsUndoEnd!��������ID, 0) <> 0 Then
            clsXML.appendData "after_area_id", Nvl(rsUndoEnd!��������ID, 0), xsNumber
        End If
        'after_area_title �����䶯��������
        rsDeptOper.Filter = "����ID=" & Val(Nvl(rsUndoEnd!��������ID, 0))
        If rsDeptOper.RecordCount > 0 Then
            clsXML.appendData "after_area_title", Nvl(rsDeptOper!��������), xsString
        End If
        'after_dept_id       �����䶯�����id    0..1    N
        clsXML.appendData "after_dept_id", Nvl(rsUndoEnd!���������Id, 0), xsNumber
        'after_dept_title �����䶯���������
        rsDeptOper.Filter = "����ID=" & Val(Nvl(rsUndoEnd!���������Id, 0))
        If rsDeptOper.RecordCount > 0 Then
            clsXML.appendData "after_dept_title", Nvl(rsDeptOper!��������), xsString
        Else
            clsXML.appendData "after_dept_title", "", xsString
        End If
        Select Case strType
        Case "����"
            'after_duty_nurse_id �����䶯��ʿID
            rsDeptOper.Filter = "��Ա����='" & Nvl(rsUndoEnd!���λ�ʿ) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_duty_nurse_id", Val(Nvl(rsDeptOper!��ԱID)), xsNumber
            End If
            'after_duty_nurse �����䶯��ʿ����
            clsXML.appendData "after_duty_nurse", Nvl(rsUndoEnd!���λ�ʿ), xsString
            'after_bed_no �����䶯��Ĵ���
            clsXML.appendData "after_bed_no", Nvl(rsUndoEnd!����), xsString
        Case "�����䶯"
            'after_situation �����䶯��Ĳ���
            clsXML.appendData "after_situation", Nvl(rsUndoEnd!����), xsString
            'after_situation_code �����䶯��Ĳ������
            clsXML.appendData "after_situation_code", Nvl(rsUndoEnd!�������), xsString
        Case "ת��", "����ҽʦ�ı�", "���λ�ʿ�ı�", "����ҽʦ�䶯", "����ҽʦ�䶯", "תҽ��С��"
            'after_in_doctor_id �����䶯����ҽ��ID
            rsDeptOper.Filter = "��Ա����='" & Nvl(rsUndoEnd!����ҽʦ) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_in_doctor_id", Val(Nvl(rsDeptOper!��ԱID)), xsNumber
            End If
            'after_in_doctor �����䶯����ҽ������
            clsXML.appendData "after_in_doctor", Nvl(rsUndoEnd!����ҽʦ), xsString
            'after_treat_doctor_id �����䶯������ҽ��ID
            rsDeptOper.Filter = "��Ա����='" & Nvl(rsUndoEnd!����ҽʦ) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_treat_doctor_id", Val(Nvl(rsDeptOper!��ԱID)), xsNumber
            End If
            'after_treat_doctor �����䶯������ҽ������
            clsXML.appendData "after_treat_doctor", Nvl(rsUndoEnd!����ҽʦ), xsString
            'after_director_doctor_id �����䶯������ҽ��ID
            rsDeptOper.Filter = "��Ա����='" & Nvl(rsUndoEnd!����ҽʦ) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_director_doctor_id", Val(Nvl(rsDeptOper!��ԱID)), xsNumber
            End If
            'after_director_doctor �����䶯������ҽ������
            clsXML.appendData "after_director_doctor", Nvl(rsUndoEnd!����ҽʦ), xsString
            'after_duty_nurse_id �����䶯��ʿID
            rsDeptOper.Filter = "��Ա����='" & Nvl(rsUndoEnd!���λ�ʿ) & "'"
            If rsDeptOper.RecordCount > 0 Then
                clsXML.appendData "after_duty_nurse_id", Val(Nvl(rsDeptOper!��ԱID)), xsNumber
            End If
            'after_duty_nurse �����䶯��ʿ����
            clsXML.appendData "after_duty_nurse", Nvl(rsUndoEnd!���λ�ʿ), xsString
            'after_bed_no �����䶯��Ĵ���
            clsXML.appendData "after_bed_no", Nvl(rsUndoEnd!����), xsString
        End Select
        clsXML.AppendNode "change_cancel", True
        clsMipModule.CommitMessage "ZLHIS_PATIENT_006", clsXML.XmlText
        If bln��λ�Ի� = True Then
            bln��λ�Ի� = False
            lng����ID = lngNextPati
            lng��ҳID = lngNextPage
            Set rsUndoBegin = rsBedChange
            GoTo PreNext
        End If
    End If
    'ж����Ϣ����
    If Not (clsMipModule Is Nothing) Then
        Call clsMipModule.CloseMessage
        Call DelMipModule(clsMipModule)
        Set clsMipModule = Nothing
    End If
    If Not (clsXML Is Nothing) Then
        Set clsXML = Nothing
    End If
    
    If Err <> 0 Then Err.Clear
    gblnOK = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOutState(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select ��Ժ��ʽ From ������ҳ Where ����ID = [1] And ��ҳID = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, lng��ҳID)
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!��Ժ��ʽ) Then
            GetOutState = rsTmp!��Ժ��ʽ
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetUsedBeds(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef str�ȼ�IDs) As String
'����:����Ժ������ǰ�Ĵ�λ�Ƿ����ǿմ�,ֻҪ����֮һ����,�򷵻ز��˵����д���
'������str�ȼ�ID����������λ��Ӧ�ĵȼ�ID
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strBeds As String
 
    strSql = "Select A.����,B.״̬,A.��λ�ȼ�ID From ���˱䶯��¼ A,��λ״����¼ B  Where A.����id=[1] And A.��ҳid=[2] And A.��ֹԭ�� = 1" & _
        " And A.����id = B.����id And A.���� = B.����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lng����ID, lng��ҳID)
    
    rsTmp.Filter = "״̬<>'�մ�'"
    If rsTmp.RecordCount = 0 Then Exit Function
    
    rsTmp.Filter = ""
    Do While Not rsTmp.EOF
        strBeds = strBeds & "," & rsTmp!����
        str�ȼ�IDs = str�ȼ�IDs & "," & rsTmp!��λ�ȼ�id
        rsTmp.MoveNext
    Loop
    If strBeds <> "" Then
        GetUsedBeds = Mid(strBeds, 2)
        str�ȼ�IDs = Mid(str�ȼ�IDs, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function InitSysPar() As Boolean
'���ܣ���ʼ��ϵͳ����
    Dim strValue As String
    On Error Resume Next
        
    '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    '������ʾ��ʽ
    'gblnShowCard = zldatabase.GetPara(12, glngSys) = "0"
    
    '���ŵ���ʹ�ö��ֽ��㷽ʽģʽ
    gblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
       
    'Ʊ�ݺ��볤�ȡ����￨�ų���
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    gbytPrepayLen = Val(Split(strValue, "|")(1))
    'gbytCardNOLen = Val(Split(strValue, "|")(4))
    If gbytPrepayLen = 0 Then gbytPrepayLen = 7
    'If gbytCardNOLen = 0 Then gbytCardNOLen = 7
    
    'Ʊ���ϸ����
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    gblnPrepayStrict = Mid(strValue, 2, 1) = "1"
    'gblnMagcardStrict = Mid(strValue, 5, 1) = "1"
    
    
    gbln��ԺԤ�� = zlDatabase.GetPara(10, glngSys) = "1"
    gbln��Ժ���� = zlDatabase.GetPara(11, glngSys) = "1"
    gbln��Ժ��� = zlDatabase.GetPara(13, glngSys) = "1"
    gbln��Ժ���˲�׼��Ժ���� = Val(zlDatabase.GetPara(31, glngSys))
    gbyt��Ժʱ���δִ�� = Val(zlDatabase.GetPara(22, glngSys))
    gbytת��ʱ���δִ�� = Val(zlDatabase.GetPara(32, glngSys))
    '����30208 by lesfeng 2010-08-02 ���ֲ���22��32 ����154��155
    gbyt��Ժʱ���ҩƷδִ�� = Val(zlDatabase.GetPara(154, glngSys))
    gbytת��ʱ���ҩƷδִ�� = Val(zlDatabase.GetPara(155, glngSys))
    
    '61429:������,2013-11-11
    gbytת��ʱδ������ʵ��ݼ�� = Val(zlDatabase.GetPara(227, glngSys))
    '68953
    gbyt��Ժʱ���ڻ������ݼ�� = Val(zlDatabase.GetPara(235, glngSys))
    
    gblnҽ��������ܳ�Ժ = zlDatabase.GetPara(43, glngSys) = "1"
    '--55791:������,2012-11-13,���˳�Ժҽ�����ܳ�����Ժ
    gblnҽ��������ܳ���Ԥ��Ժ = (Val(zlDatabase.GetPara(192, glngSys)) = 1)
    
    '��Ժ�Ǽ�ʱˢ����������
    gblnCheckPass = Mid(zlDatabase.GetPara(46, glngSys, , "0000000000"), 5, 1) = "1"
        
    '������뷽ʽ
    strValue = zlDatabase.GetPara(65, glngSys, , "11")
    If Len(strValue) = 1 Then strValue = strValue & strValue
    gint����������� = Mid(strValue, 1, 1)
    gintסԺ������� = Mid(strValue, 2)
        
    gblnҩ�ƻ��۵� = zlDatabase.GetPara(79, glngSys) = "1"
    gbln�������۵� = zlDatabase.GetPara(80, glngSys) = "1"
    gblnִ�к���� = zlDatabase.GetPara(81, glngSys) = "1"
    gbln���ȷ������ȼ� = zlDatabase.GetPara(99, glngSys) = "1"
    gblnÿ��סԺ��סԺ�� = zlDatabase.GetPara(145, glngSys) = "1"
    '���˺� ����:????    ����:2010-12-06 23:38:53
    '���õ��۱���λ��
    gintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
    '61347:������,2013-11-09
    gbyt������˷�ʽ = Val(zlDatabase.GetPara(185, glngSys, , "0"))
    
    gblnXW = Val(zlDatabase.GetPara(255, glngSys)) = 1
    'ͬһ���ֻ֤�ܶ�Ӧһ����������
    gblnPatiByID = Val(zlDatabase.GetPara(279, glngSys)) = 1
    gblnҽ�ƻ�������������¼�� = zlDatabase.GetPara(287, glngSys, , "0") = "1"
    
    InitSysPar = True
End Function

Public Sub InitLocPar(ByVal lngModul As Long)
'���ܣ���ʼ��ģ�����
    Dim strValue As String
    On Error Resume Next

    gstrLike = IIf(zlDatabase.GetPara("����ƥ��") = 0, "%", "")
    strValue = zlDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    gbytCode = Val(zlDatabase.GetPara("���뷽ʽ"))
    gblnMyStyle = zlDatabase.GetPara("ʹ�ø��Ի����") = "1"
    
    gbln���ýṹ����ַ = Val(zlDatabase.GetPara("���˵�ַ�ṹ��¼��", glngSys)) <> 0
    gbln��ʾ���� = Val(zlDatabase.GetPara("�����ַ�ṹ��¼��", glngSys)) <> 0
          
    If lngModul = P������Ժ���� Then
        gstrԤ��ID = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, lngModul, "")
        If gstrԤ��ID <> "" Then Call UpdateShareID(lngModul, gstrԤ��ID, 2)
        'LED��������
        gblnLED = Val(GetSetting("ZLSOFT", "����ȫ��", "ʹ��", 0)) <> 0
        gblnLedWelcome = Val(zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, lngModul, 1)) <> 0
        
        gbln���� = zlDatabase.GetPara("������Ϣ", glngSys, lngModul) = "1"
        gblnSeekName = zlDatabase.GetPara("����ģ������", glngSys, lngModul) = "1"
        gintNameDays = Val(zlDatabase.GetPara("������������", glngSys, lngModul))
        gbln���� = zlDatabase.GetPara("���Ѽ���", glngSys, lngModul) = "1"
        gbln��ѡ���� = zlDatabase.GetPara("��ѡ����", glngSys, lngModul) = "1"
        gbln���ƿմ� = zlDatabase.GetPara("�������޿մ����ܵǼ�", glngSys, lngModul) = "1"
        
        gbytPrepayPrint = Val(zlDatabase.GetPara("Ԥ����Ʊ�ݴ�ӡ", glngSys, lngModul))
        gbytFPagePrint = Val(zlDatabase.GetPara("������ҳ��ӡ", glngSys, lngModul))
        gbytWristletPrint = Val(zlDatabase.GetPara("���������ӡ", glngSys, lngModul))
        
        '36454,������,2012-09-06
        gbln���ü��� = Val(zlDatabase.GetPara("���ü���ʱ��", glngSys, lngModul)) = 1
        
    ElseIf lngModul = P����������� Then
        gbyt���ʱ�� = Val(zlDatabase.GetPara("ȱʡ���ʱ��", glngSys, lngModul, "0"))
        gbytBabyWristletPrint = Val(zlDatabase.GetPara("Ӥ�������ӡ", glngSys, lngModul))
        '49854:������,2013-10-31,���������ӡ
        gbytCourseWristletPrint = Val(zlDatabase.GetPara("���������ӡ", glngSys, lngModul))
        gblnת����ת���� = zlDatabase.GetPara("ת����ת����", glngSys, lngModul) = "1"
    End If
End Sub

Public Function SaveIDCard(bytStyle As Byte, strNo As String, lng����ID As Long, lng��ҳID As Long, _
        lng���˲���ID As Long, lng���˿���ID As Long, str��ʶ�� As String, str�ѱ� As String, _
        strԭ���� As String, str���� As String, str�Ա� As String, str���� As String, str���� As String, str���� As String, _
        curӦ�ս�� As Currency, curʵ�ս�� As Currency, str���㷽ʽ As String, Dat����ʱ�� As Date, lng����ID As Long, rsMoney As ADODB.Recordset, ByVal strICCard As String) As String
'���ܣ�����һ�����￨���ü�¼SQL���
'������bytStyle=0-����,1-����,2-����
'      cur���=���￨���
'      str���㷽ʽ=���Ϊ��,��ʾ����,�����ֽ�
'      rsMoney:�������￨�շ���Ϣ�ļ�¼��
'      strԭ����=������ʱ��
'      lng����ID=��ǰ���õľ��￨����ID
'      strICCard=IC����,ͨ����IC����ʽ����ʱ,ͬʱ��д������Ϣ��IC���ֶ�
    Dim lngUnitID As Long
    Dim strSql As String
    
'    Select Case rsMoney!���ұ�־
'        Case 1 'ָ������
'            lngUnitID = GetItemUnitID(rsMoney!���ұ�־, rsMoney!�շ�ϸĿID)
'        Case 2 '���˿���
'            If lng���˿���ID <> 0 Then
'                lngUnitID = lng���˿���ID
'            Else
'                lngUnitID = UserInfo.����ID
'            End If
'        Case 0, 3, 5, 6
'            lngUnitID = UserInfo.����ID
'        Case 4 'ָ������
'            lngUnitID = GetItemUnitID(rsMoney!���ұ�־, rsMoney!�շ�ϸĿID)
'    End Select
    
    '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
    Select Case rsMoney!���ұ�־
        Case 4 'ָ������
            lngUnitID = GetItemUnitID(rsMoney!���ұ�־, rsMoney!�շ�ϸĿID)
        Case 1, 2 '���˿���
            If lng���˿���ID <> 0 Then
                lngUnitID = lng���˿���ID
            Else
                lngUnitID = UserInfo.����ID
            End If
        Case 0, 3, 5, 6
            lngUnitID = UserInfo.����ID
    End Select
    
    '���ù���"zl_���￨��¼_Insert"
    strSql = "zl_���￨��¼_INSERT(" & bytStyle & ",'" & strNo & "'," & lng����ID & "," & lng��ҳID & "," & _
        IIf(str��ʶ�� = "0", "NULL", str��ʶ��) & ",'" & str�ѱ� & "','" & strԭ���� & "','" & str���� & "','" & str���� & "','" & str���� & _
        "','" & str�Ա� & "','" & str���� & "'," & lng���˲���ID & "," & lng���˿���ID & "," & rsMoney!�շ�ϸĿID & _
        ",'" & rsMoney!�շ���� & "','" & IIf(IsNull(rsMoney!���㵥λ), "", rsMoney!���㵥λ) & "'," & _
        rsMoney!������ĿID & ",'" & rsMoney!�վݷ�Ŀ & "'," & curӦ�ս�� & "," & lngUnitID & "," & UserInfo.����ID & _
        ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(OverTime(Dat����ʱ��), "1", "0") & _
        ",To_Date('" & Format(Dat����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & str���㷽ʽ & "'," & IIf(lng����ID = 0, "NULL", lng����ID) & ",'" & strICCard & "'," & curӦ�ս�� & "," & curʵ�ս�� & ")"
    
    SaveIDCard = strSql
End Function

Private Function GetItemUnitID(bytFlag As Byte, lngID As Long) As Long
'���ܣ������շ��ض���Ŀ��ִ�п���
'������bytFlag=ִ�п��ұ�־,lngID=�շ�ϸĿID
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '����ȷ����
            GetItemUnitID = UserInfo.����ID 'ȡ����Ա���ڿ���
        Case 4 'ָ������
            strSql = "Select B.ִ�п���ID From �շ���ĿĿ¼ A,�շ�ִ�п��� B Where B.�շ�ϸĿID=A.ID And A.ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngID)
            
            If Not rsTmp.EOF Then
                GetItemUnitID = rsTmp!ִ�п���ID 'Ĭ��ȡ��һ��(���ж��)
            Else
                GetItemUnitID = UserInfo.����ID '��û��ָ������ȡ����Ա���ڿ���
            End If
        Case 1, 2, 3 '���˿���,����Ա����
            GetItemUnitID = UserInfo.����ID '��ȡ����Ա����
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetSelectPersonal(ByVal strType As String, ByVal strPosts As String, ByRef frmParent As Form) As ADODB.Recordset
'����:strType=��Ա����
'     strPosts=רҵ����ְ�� �Զ��ŷָ�
    Dim strSql As String
    
    On Error GoTo errH
    strSql = _
            "Select ID,�ϼ�ID,0 as ĩ��,���� as ����,����,����" & _
            " From ���ű� Where ���� Not like '-%' Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
            " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
            " Union All" & _
            " Select Distinct A.ID,C.����ID as �ϼ�ID,1 as ĩ��,A.���,A.����,A.����" & _
            " From ��Ա�� A,��Ա����˵�� B,������Ա C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID And B.��Ա����='" & strType & _
            "' And C.ȱʡ=1 And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
            
    If strPosts <> "" Then
        strSql = strSql & " And Instr(',' || '" & strPosts & "' || ',', ',' || A.רҵ����ְ�� || ',') > 0"
    End If
    
    Set GetSelectPersonal = frmPubSel.ShowSelect(frmParent, strSql, 2, "��Աѡ��")
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDepCharacter(ByVal lngDepID As Long) As String
'���ܣ���ȡ���Ź�������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select �������� From ��������˵�� Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngDepID)
    Do While Not rsTmp.EOF
        If InStr(1, GetDepCharacter & ",", "," & rsTmp!�������� & ",") = 0 Then
            GetDepCharacter = GetDepCharacter & "," & rsTmp!��������
        End If
        rsTmp.MoveNext
    Loop
    
    If GetDepCharacter <> "" Then GetDepCharacter = Mid(GetDepCharacter, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'����27392 by lesfeng 2010-01-14
Public Function GetPatiInfoModiOut(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���������Ϣ
'������byt��Ʒ�ʽ:0-��Ժ��ƣ�1-ת�����
    Dim strSql As String
    On Error GoTo errH
    '����28982 by lesfeng 2010-06-09 ���ӡ�ȷ�����ڡ�
    strSql = "Select NVl(B.����,A.����) ����, Nvl(NVL(B.�Ա�,A.�Ա�),'δ֪') �Ա�, NVL(B.����,A.����) ����, B.����, B.��������, B.��ǰ����, B.����ȼ�id, B.סԺҽʦ, B.����ҽʦ, B.���λ�ʿ, B.��Ժ����id, B.��Ժ����id ��ס����id," & vbNewLine & _
        "       To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.��ǰ����id, B.סԺ��, D.���� as ��ǰ����, B.��Ժ���� as ��Ҫ����," & vbNewLine & _
        "     To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժ����,�Ƿ�ȷ��,ȷ������,��Ժ��ʽ,�����־,��������,ʬ���־" & vbNewLine & _
        "From ������Ϣ A, ������ҳ B, ���ű� D" & vbNewLine & _
        "Where A.����id = B.����id And B.����id = [1] And B.��ҳid = [2] And B.��Ժ����id = D.id and B.��Ժ���� is not null"
    Set GetPatiInfoModiOut = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional byt��ס��ʽ As Byte) As ADODB.Recordset
'���ܣ���ȡ���������Ϣ
'������byt��ס��ʽ:0-��Ժ��ƣ�1-ת����ƣ�2-�벡��
    Dim strSql As String
    '����31652 by lesfeng �Ӳ�����ҳֱ����ȡȷ�����ڣ������Ƿ�ȷ�Ｐȷ������
    '49163,������,2012-09-07,�Ӳ�����ҳֱ����ȡ�����־����������
    On Error GoTo errH
    If byt��ס��ʽ = 0 Then '
        strSql = "Select NVL(B.����,A.����) ����, Nvl(NVL(B.�Ա�,A.�Ա�),'δ֪') �Ա�, B.����,To_Char(A.��������,'YYYY-MM-DD HH24:MI:SS') As ��������, B.����, B.��������, B.��ǰ����, B.����ȼ�id, B.סԺҽʦ, B.����ҽʦ, B.���λ�ʿ, B.��Ժ����id, B.��Ժ����id ��ס����id, B.ҽ��С��id, " & vbNewLine & _
            "       To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.��ǰ����id, B.סԺ��, D.���� as ��ǰ����,E.���� as ��ǰ����, B.��Ժ���� as ��Ҫ����,B.�Ƿ�ȷ��,B.ȷ������, B.��Ժ����, B.��Ժ��ʽ, B.ʬ���־,B.�����־,B.��������,B.����Ժ,B.�Һ�ID " & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ���ű� D,���ű� E" & vbNewLine & _
            "Where A.����id = B.����id And B.����id = [1] And B.��ҳid = [2] And B.��Ժ����id = D.id And B.��ǰ����Id=E.id(+)"
    ElseIf byt��ס��ʽ = 1 Then
        strSql = "Select NVL(B.����,A.����) ����, Nvl(NVL(B.�Ա�,A.�Ա�),'δ֪') �Ա�, B.����,To_Char(A.��������,'YYYY-MM-DD HH24:MI:SS') As ��������, B.����, B.��������, B.��ǰ����, B.����ȼ�id, B.סԺҽʦ, B.����ҽʦ, B.���λ�ʿ, B.��Ժ����id, C.����id As ��ס����id, B.ҽ��С��id, " & vbNewLine & _
            "       To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.��ǰ����id, B.סԺ��, D.���� as ��ǰ����,E.���� as ��ǰ����, B.��Ժ���� as ��Ҫ����,B.�Ƿ�ȷ��,B.ȷ������, B.��Ժ����, B.��Ժ��ʽ, B.ʬ���־,B.�����־,B.��������,B.����Ժ,B.�Һ�ID" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ���ű� D,���ű� E" & vbNewLine & _
            "Where A.����id = B.����id And B.����id = [1] And B.��ҳid = [2] And C.����id = B.����id And C.��ҳid = B.��ҳid And" & vbNewLine & _
            "      C.��ʼԭ�� = 3 And C.��ʼʱ�� Is Null And C.��ֹʱ�� Is Null And B.��Ժ����id = D.id And B.��ǰ����Id=E.id(+)"
    ElseIf byt��ס��ʽ = 2 Then
        strSql = "Select NVL(B.����,A.����) ����, Nvl(NVL(B.�Ա�,A.�Ա�),'δ֪') �Ա�, B.����,To_Char(A.��������,'YYYY-MM-DD HH24:MI:SS') As ��������, B.����, B.��������, B.��ǰ����, B.����ȼ�id, B.סԺҽʦ, B.����ҽʦ, B.���λ�ʿ, B.��Ժ����id, C.����id As ��ס����id, B.ҽ��С��id, " & vbNewLine & _
            "       To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.��ǰ����id, B.סԺ��, D.���� as ��ǰ����,E.���� as ��ǰ����, B.��Ժ���� as ��Ҫ����,B.�Ƿ�ȷ��,B.ȷ������, B.��Ժ����, B.��Ժ��ʽ, B.ʬ���־,B.�����־,B.��������,B.����Ժ,B.�Һ�ID" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ���ű� D,���ű� E" & vbNewLine & _
            "Where A.����id = B.����id And B.����id = [1] And B.��ҳid = [2] And C.����id = B.����id And C.��ҳid = B.��ҳid And" & vbNewLine & _
            "      C.��ʼԭ�� = 15 And C.��ʼʱ�� Is Null And C.��ֹʱ�� Is Null And B.��Ժ����id = D.id And B.��ǰ����Id=E.id(+)"
    End If
    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInDept(blnByDept As Boolean, DateBegin As Date, DateEnd As Date, ByVal strNodeNo As String, _
    Optional ByVal strDeptIDs As String) As ADODB.Recordset
'���ܣ���ȡָ����Ժʱ�䷶Χ�ڵĿ��һ���
'������strDeptIDs-��ѡ���һ���ID
    Dim strSql As String

    On Error GoTo errH
    If strDeptIDs <> "" Then strSql = strSql & "  And Instr(','||[4]||',',','||B.ID||',')>0 "
    strSql = "Select A.ID, B.����, B.����" & vbNewLine & _
            "From (Select Distinct " & IIf(blnByDept, "��Ժ����id", "��Ժ����id") & " ID" & vbNewLine & _
            "       From ������ҳ" & vbNewLine & _
            "       Where ��Ժ���� Between [1] And [2] And " & IIf(blnByDept, "��Ժ����id", "��Ժ����id") & " Is Not Null) A, ���ű� B" & vbNewLine & _
            "Where A.ID = B.ID  And (B.վ��=[3] Or B.վ�� is Null) " & strSql & vbNewLine & _
            "Order By B.����"

    Set GetInDept = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", DateBegin, DateEnd, strNodeNo, strDeptIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function SimilarIDs(str���� As String, str���� As String, dat�������� As Date, str�Ա� As String, str���� As String, str���֤�� As String) As ADODB.Recordset
'���ܣ���鲡���Ƿ����������Ϣ
'���أ����Ƽ�¼�Ĳ���ID��,��"234,235,236"
    On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Integer
    
    strSql = _
        " Select Rownum+1 ID,����ID,Nvl(���֤��,'δ�Ǽ�') ���֤��,�����,סԺ��,Nvl(��ͥ��ַ,'δ�Ǽ�') ��ַ,To_Char(�Ǽ�ʱ��,'YYYY-MM-DD') �Ǽ�ʱ�� " & _
        " From ������Ϣ Where (����=[1] And ����=[2] And �Ա�=[3] And ����=[4] " & _
        " And ��������=TO_DATE([5],'YYYY-MM-DD')) Or ���֤��=[6] " & _
        " Order by ����ID Desc"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", str����, str����, str�Ա�, str����, Format(dat��������, "YYYY-MM-DD"), str���֤��)
    
    Set SimilarIDs = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMax��ҳID(lng����ID As Long) As Long
'���ܣ���ȡ���˵���󲡰���ҳID
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Nvl(Max(��ҳID),0)+1 as ��ҳID From ������ҳ Where Nvl(��ҳID,0)<>0 And ����ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID)
    If rsTmp.EOF Then
        GetMax��ҳID = 1
    Else
        GetMax��ҳID = rsTmp!��ҳID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function NextBedNo(lngUnitID As Long, str�������� As String, str���� As String) As String
'���ܣ���ȡָ����������һ��λ��
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errH
    If str���� <> "" Then
        gstrSQL = _
                "Select nvl(Max(to_number(Substr(����, nvl(Length(B.����),0) + 1))),0) ����,nvl(max(length(����)-length(����)),1) λ�� " & _
                "From ��λ״����¼ A, ��λ���Ʒ��� B " & _
                "Where A.��λ���� = B.���� And ��λ����=[1] and  A.���� like [2] and A.����id=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInPatient", str��������, str���� & "%", lngUnitID)
    Else
        gstrSQL = _
                "Select nvl(Max(to_number(Substr(����, nvl(Length(B.����),0) + 1))),0) ����,nvl(max(length(����)-length(����)),1) λ�� " & _
                "From ��λ״����¼ A, ��λ���Ʒ��� B " & _
                "Where A.��λ���� = B.���� And ��λ����=[1] and  zl_to_number(A.����)>0 and A.����id=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "mdlInPatient", str��������, lngUnitID)

    End If
    
    If Not rsTmp.EOF Then
        If IsNumeric(rsTmp!����) Then
            strTmp = rsTmp!���� + 1
        Else
            strTmp = Val(rsTmp!����) + 1
        End If
    Else
        NextBedNo = 1
    End If
    
    If Len(strTmp) > rsTmp!λ�� Then
        NextBedNo = strTmp
    Else
        NextBedNo = Right("00000" & strTmp, rsTmp!λ��)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function isRepeat(lngUnitID As Long, strBeds As String) As String
'���ܣ��ж���ָ�������ڵ�һϵ�д����Ƿ��Ѿ�����
'������lngUnitID=����ID,strBeds=�����ַ���,��"12,13,15..."
'���أ���=��������,����"12,13..."��Щ�����ظ�
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strSql As String
    
    On Error GoTo errH
    
    strSql = "Select ���� From ��λ״����¼ Where ����ID=[1] And instr([2],','||����||',')>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngUnitID, "," & strBeds & ",")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            isRepeat = isRepeat & rsTmp!���� & ","
            rsTmp.MoveNext
        Next
        isRepeat = Left(isRepeat, Len(isRepeat) - 1)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInPatiNO(strסԺ�� As String, Optional lng����ID As Long, Optional bln�������� As Boolean) As Boolean
'���ܣ��ж�ָ��סԺ���Ƿ��Ѿ����������ݿ���,ÿ��סԺ��סԺ��ʱ,ԤԼ���޸�����ԭ����סԺ��,�ſ������Լ�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    If gblnÿ��סԺ��סԺ�� Then
        strSql = "Select 1 From ������ҳ Where סԺ��=[1]" & IIf(bln��������, " And ����ID<>[2]", "")
    Else
        strSql = _
            " Select 1 From ������Ϣ Where סԺ��=[1] And ����ID<>[2]" & vbNewLine & _
            " UNION ALL" & vbNewLine & _
            " Select 1 From ������ҳ Where סԺ��=[1] And ����ID<>[2]"
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", strסԺ��, lng����ID)
    If rsTmp.RecordCount > 0 Then ExistInPatiNO = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistInPatiID(lngPatientID As Long) As Boolean
'���ܣ��ж�ָ������ID�Ƿ��Ѿ����������ݿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errH
    
    strSql = "Select 1 From ������Ϣ Where ����ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngPatientID)
    If rsTmp.RecordCount > 0 Then ExistInPatiID = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check�ѱ����ÿ���(ByVal str�ѱ� As String, ByVal lng����ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errH
    
    '���������п���,��ǰָ������
    strSql = "Select 1" & vbNewLine & _
            "From Dual" & vbNewLine & _
            "Where Not Exists (Select 1 From �ѱ����ÿ��� Where �ѱ� = [1]) Or Exists" & vbNewLine & _
            " (Select 1 From �ѱ����ÿ��� Where �ѱ� = [1] And ����id = [2])"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", str�ѱ�, lng����ID)
    If rsTmp.RecordCount > 0 Then Check�ѱ����ÿ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMaxDate(lng����ID As Long, lng��ҳID As Long, Optional intԭ�� As Integer) As Date
'���ܣ���ȡת�Ʋ��������ϴα䶯ʱ��
'������intԭ��=�����ϴα䶯��ԭ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    GetMaxDate = #1/1/1900#
    intԭ�� = 0
    
    strSql = "Select ��ʼʱ��,��ʼԭ�� From ���˱䶯��¼" & _
        " Where ��ʼʱ�� is Not NULL And ��ֹʱ�� is NULL" & _
        " And ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        GetMaxDate = IIf(IsNull(rsTmp!��ʼʱ��), GetMaxDate, rsTmp!��ʼʱ��)
        intԭ�� = Nvl(rsTmp!��ʼԭ��, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastInfo(lngPatientID As Long) As String
'���ܣ���ȡ�������һ��Ԥ���λ��Ϣ
'���أ�"�ɿλ|��λ������|��λ�ʺ�"
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    'by lesfeng 2010-01-11 �����Ż�
    '��ģ���¼����=1
    strSql = "Select �ɿλ,��λ������,��λ�ʺ� From ����Ԥ����¼ Where (�ɿλ is Not NULL Or ��λ������ is Not NUll Or ��λ�ʺ� is Not NULL) And ��¼����=1 And ����ID=[1] Order by �տ�ʱ�� Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lngPatientID)
    
    If Not rsTmp.EOF Then
        GetLastInfo = IIf(IsNull(rsTmp!�ɿλ), "", rsTmp!�ɿλ) & "|" & IIf(IsNull(rsTmp!��λ������), "", rsTmp!��λ������) & "|" & IIf(IsNull(rsTmp!��λ�ʺ�), "", rsTmp!��λ�ʺ�)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Checkҽ���´��Ժҽ��(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ��жϲ����Ƿ���Ԥ��Ժ״̬,�Ҵ�����Ч�ĳ�Ժ(תԺ������)ҽ���������Ժ(��Ч��ҽ����ָ��ʼִ��ʱ����Ԥ��Ժʱ����ͬ���Ҵ����ѷ���״̬[ҽ��״̬=8])��
'������
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    '--55791:������,2012-11-13,���ϳ�Ժҽ�����ܳ�����Ժ
     '�������"���ϳ�Ժҽ�����ܳ�����Ժ"Ϊ�٣�ҽ����¼�ͺͲ��˱䶯��¼�޷���Ӧ��ԭ��SQL(ע������),�Ͳ���������Ԥ��Ժ���Ժ
'    strSQL = "Select a.Id" & vbNewLine & _
'            " From ����ҽ����¼ a, ���˱䶯��¼ b, ������ҳ c, ������ĿĿ¼ d" & vbNewLine & _
'            " Where a.����id = [1] And a.��ҳid = [2] And a.ҽ��״̬ = 8 And a.����id = b.����id And a.��ҳid = b.��ҳid And" & vbNewLine & _
'            "           a.��ʼִ��ʱ�� = b.��ʼʱ��+0 And b.��ʼԭ�� = 10 And b.����id = c.����id And b.��ҳid = c.��ҳid And" & vbNewLine & _
'            "           c.״̬ = 3 And d.���='Z' And d.�������� In ('5', '6', '11') And a.������Ŀid = d.Id"
    '102342 ���ò���ʱ��"�´��Ժҽ�����ܳ�Ժ",��Ӥ���´��Ժҽ��,���˱���δ�´�ʱ,�������˳�Ժ��
    strSql = "Select a.Id" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
            "Where a.����id = [1] And a.��ҳid = [2] And a.ҽ��״̬ = 8 And Nvl(a.Ӥ��, 0) = 0 And b.��� = 'Z' And b.�������� In ('5', '6', '11') And" & vbNewLine & _
            "      a.������Ŀid = b.Id "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    
    Checkҽ���´��Ժҽ�� = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check��������(lng����ID As Long, lng��ҳID As Long) As String
'���ܣ��жϲ����Ǵ������������¼
'������
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strInfo As String
    '56323:������,2013-02-18,��ǿ����Ϊ��˵��ݵ���ʾ��Ϣ����
    'strSQL = "Select 1 From סԺ���ü�¼ A, ���˷������� B Where A.����ID=[1] And A.��ҳID=[2] And A.Id = B.����ID And b.״̬=0"
    strSql = "Select distinct A.NO,C.���� ��˿���,D.���� ��Ŀ���� From סԺ���ü�¼ A, ���˷������� B,���ű� C,�շ���ĿĿ¼ D" & vbNewLine & _
        "        Where A.����ID=[1] And A.��ҳID=[2] And A.Id = B.����ID And b.״̬=0 And B.��˲���ID=C.ID And B.�շ�ϸĿID=D.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    With rsTmp
        Do While Not .EOF
            If strInfo = "" Then
                strInfo = "����[" & Nvl(!NO) & "]�е�" & Nvl(!��Ŀ����) & "����" & Nvl(!��˿���, "[δ������]") & "δ���"
            Else
                strInfo = strInfo & vbCrLf & "����[" & Nvl(!NO) & "]�е�" & Nvl(!��Ŀ����) & "����" & Nvl(!��˿���, "[δ������]") & "δ���"
            End If
        rsTmp.MoveNext
        Loop
    End With
    Check�������� = strInfo
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckAudited(lng����ID As Long, lng��ҳID As Long, Optional bytAudited As Byte = 1) As Boolean
'���ܣ��жϲ����Ƿ������
'������
'      lng����ID:����ID
'      lng��ҳID:��ҳID
'      bytAudited:0-δ���;1-�����;2-������
'����:TRUE OR FALSE
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    If bytAudited = 0 Then
        'δ���
        strSql = "Select ����id From ������ҳ Where ����id=[1] And ��ҳid=[2] And Nvl(��˱�־,0)=0"
    ElseIf bytAudited = 1 Then
        '�Ѿ����
        strSql = "Select ����id From ������ҳ Where ����id=[1] And ��ҳid=[2] And Nvl(��˱�־,0)>=1"
    Else
        '������
        strSql = "Select ����id From ������ҳ Where ����id=[1] And ��ҳid=[2] And Nvl(��˱�־,0)=2"
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    CheckAudited = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMaxBedLen(Optional lng����ID As Long, Optional blnռ�� As Boolean) As Integer
'���ܣ���ȡָ�����ŵĴ�λ�ŵ���󳤶�
'������lng����ID=����ID�����ID,Ϊ0��ʾ���в��������
'      blnռ��=�Ƿ�ֻ�ܱ�ռ�õĴ�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    If Not blnռ�� Or lng����ID = 0 Then
        strSql = "Select Max(Lengthb(����)) as ���� From ��λ״����¼ Where  ����ID" & IIf(lng����ID = 0, " is Not NULL", "=[1]")
    Else
        strSql = "Select Max(Lengthb(����)) as ���� From ��λ״����¼ Where ״̬='ռ��' And ����ID" & IIf(lng����ID = 0, " is Not NULL", "=[1]")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInExse", lng����ID)
    If Not rsTmp.EOF Then GetMaxBedLen = IIf(IsNull(rsTmp!����), 0, rsTmp!����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPatiLog(lng����ID As Long, lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���˱䶯��¼
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    'by lesfeng 2010-01-11 �����Ż�
    strSql = "Select id,����ID ,��ҳID ,��ʼʱ�� ,��ʼԭ��,���Ӵ�λ,����id,����id,����ȼ�id,��λ�ȼ�id,����," & _
             "       ���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,��ֹ��Ա,��ֹʱ��,��ֹԭ��,����Ա���,����Ա����,�ϴμ���ʱ�� " & _
             "  From ���˱䶯��¼" & _
             " Where Nvl(���Ӵ�λ,0)=0 And ����ID=[1] And ��ҳID=[2] " & _
             " Order by ��ֹʱ�� Desc,��ʼʱ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then Set GetPatiLog = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetDiagnosticInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                                  ByVal str������� As String, ByVal str��¼��Դ As String, Optional ByVal lngDeptID As Long = 0) As ADODB.Recordset
'���ܣ���ȡָ�����˵���ϼ�¼'
'����:
'�������-1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
'        11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���
'��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����(����ҽ��վ,���ժҪ);
'lngDeptID:��Ժ����ID������ԤԼ���ˣ�
    On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intDiagDays As Integer
    
    If lng��ҳID = 0 And InStr(1, "," & str��¼��Դ & ",", ",3,") > 0 And (InStr(1, "," & str������� & ",", ",1,") > 0 Or InStr(1, "," & str������� & ",", ",11,") > 0) Then
        '107823����ԤԼ����ʱ��ȡ������Ϲ���:
        '1-����3���ڹҺŵ�,���Ȿ��û�йҺŵĲ��˰��ϴε���ϼ�¼����
        '2-����ȡ��Ч�����ڣ���Ժ����ID��Ӧ��Ժҽ������Ӧ�����
        '3-δȡ����Ժҽ����Ӧ�����,��ȡ��Ч�����ڵ����һ�ιҺż�¼��Ӧ�ĵ�һ���
        intDiagDays = Val(zlDatabase.GetPara("��ϲ�������", glngSys, glngModul, "3"))
        If lngDeptID = 0 Then
            strSql = "Select ��Ժ����id��from ������ҳ Where ����id = [1] And ��ҳid = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
            If Not rsTmp.EOF Then lngDeptID = Val(rsTmp!��Ժ����ID & "")
        End If
        strSql = "Select a.�������, a.��¼��Դ, a.�������, a.����id, a.���id, a.��Ժ���, a.��¼����, a.�Ƿ�����, a.��ҳid" & vbNewLine & _
                "From ������ϼ�¼ A, �������ҽ�� B, ����ҽ����¼ C, ������ĿĿ¼ D, ���˹Һż�¼ E" & vbNewLine & _
                "Where a.Id = b.���id And b.ҽ��id = c.Id And c.������Ŀid + 0 = d.Id And a.��ҳid = e.Id And a.����id = [1] And a.��¼��Դ = 3  And INSTR([2], ',' || A.������� || ',') > 0 And" & vbNewLine & _
                "      e.��¼���� = 1 And e.��¼״̬ = 1 And e.�Ǽ�ʱ�� + 0 > Trunc(Sysdate - [3]) And c.ҽ��״̬ In (3, 8) And d.��� = 'Z' And" & vbNewLine & _
                "      Instr(',1,2,', d.��������) > 0 And c.ִ�п���id = [4]" & vbNewLine & _
                "Order By a.��¼���� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, "," & str������� & ",", intDiagDays, lngDeptID)
        If rsTmp.EOF Then
            strSql = "Select a.�������, a.��¼��Դ, a.�������, a.����id, a.���id, a.��Ժ���, a.��¼����, a.�Ƿ�����, a.��ҳid" & vbNewLine & _
                    "From ������ϼ�¼ A, ���˹Һż�¼ B" & vbNewLine & _
                    "Where a.��ҳid = b.Id And a.����id = [1] And b.��¼���� = 1 And b.��¼״̬ = 1 And b.�Ǽ�ʱ�� + 0 > Trunc(Sysdate - [3]) And a.��ϴ��� = 1 And" & vbNewLine & _
                    "      Instr([2], ',' || a.������� || ',') > 0 And a.��¼��Դ = 3 " & vbNewLine & _
                    "Order By a.��¼���� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, "," & str������� & ",", intDiagDays)
        End If
    Else
        strSql = " Select �������,��¼��Դ,�������,����ID,���ID,��Ժ���,��¼����,�Ƿ����� From ������ϼ�¼ " & _
                 " Where ����ID=[1] And Nvl(��ҳID,0)=[2]" & _
                 " And ��ϴ���=1  And NVL(�������,1) = 1 And instr([3],','||�������||',')>0 And ��¼��Դ in (" & str��¼��Դ & ")" & _
                 " Order by ��¼���� Desc"
        '��ϴ���-��Ժʱ,������ҳ�����п�����д��Ҫ���,��Ҫ��ϵȶ�����¼
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID, "," & str������� & ",")
    End If
    
    If lng��ҳID = 0 Then
        If Not rsTmp.EOF Then
            '103952��ιҺż�¼,ȡ���һ�ιҺż�¼
            lng��ҳID = rsTmp!��ҳID
            rsTmp.Filter = "��ҳID =" & lng��ҳID
        End If
        If Not rsTmp.EOF Then
            Set GetDiagnosticInfo = zlDatabase.CopyNewRec(rsTmp)
        Else
            Set GetDiagnosticInfo = Nothing
        End If
    Else
        If Not rsTmp.EOF Then Set GetDiagnosticInfo = rsTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetInsureInfo(lng����ID As Long) As String
'���ܣ���ȡסԺ���˱����ʻ���Ϣ
'���أ�"������;ҽ����"
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '���Ӳ�����ҳ,ȷ������סԺ�Ǳ��ղ���,����һ����Ժ
    strSql = "Select A.����,B.ҽ����" & _
        " From ������� A,�����ʻ� B,������Ϣ C,������ҳ D" & _
        " Where A.���=B.���� And B.����ID=C.����ID" & _
        " And B.����=D.���� And C.����ID=D.����ID" & _
        " And D.��ҳID=C.��ҳID And C.����ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID)
    
    If Not rsTmp.EOF Then GetInsureInfo = rsTmp!���� & ";" & rsTmp!ҽ����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function



Public Function GetLastAdviceTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Date
'���ܣ���ȡָ���������һ����Ч��ҽ����ʱ��
'˵�������ڲ��˳�Ժʱ�жϳ�Ժʱ�������ڸ�ʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    GetLastAdviceTime = CDate("1900-01-01")
    
    On Error GoTo errH
    
    '�Գ������ִ��ʱ��Ϊ׼�ж�,��ʱ�ſ������Գ���
    '��������Ժ��ҩ�����,����"��Ժ"ҽ��Ϊ׼,��Ժʱ�䱾���ͱ�����ڸñ䶯ʱ�䡣
    strSql = "Select Max(Nvl(ִ����ֹʱ��,Nvl(�ϴ�ִ��ʱ��,��ʼִ��ʱ��))) as ʱ��" & _
        " From ����ҽ����¼" & _
        " Where Nvl(ҽ����Ч,0)=0 And ҽ��״̬ Not IN(1,2,4)" & _
        " And Not (ִ��ʱ�䷽�� is NULL And (Nvl(Ƶ�ʴ���, 0) = 0 Or Nvl(Ƶ�ʼ��, 0) = 0 Or �����λ is NULL))" & _
        " And ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!ʱ��) Then
            GetLastAdviceTime = rsTmp!ʱ��
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveCatalogue(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��жϱ���סԺ�����Ƿ��ѱ�Ŀ
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select ��Ŀ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        HaveCatalogue = Not IsNull(rsTmp!��Ŀ����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function inBlackList(ByVal lng����ID As Long) As String
'���ܣ��жϲ����Ƿ��ں�������,�����ؼ���ԭ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select ����ԭ�� From ���ⲡ�� Where ����ʱ�� is NULL And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPublic", lng����ID)
    
    If Not rsTmp.EOF Then inBlackList = rsTmp!����ԭ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitExe(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���鲡����ҽ�������Ƿ���δִ�����(δִ�л�����ִ��)����Ŀ
'���أ�ҽ����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select Zl_Pati_Check_Execute(2,[1],[2]) as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExistWaitExe", lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        ExistWaitExe = Nvl(rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitDrug(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���鲡����ҩ���Ƿ���δ��ҩ��ҩƷ������
'���أ�ҩ���ͷ��ϲ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select Zl_Pati_Check_Execute(1,[1],[2]) as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExistWaitDrug", lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        ExistWaitDrug = Nvl(rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExistWaitBool(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ���鲡����Ѫ���Ƿ���δ����Ѫ
'���أ�Ѫ�ⲿ������
'������:������
'����ţ�30339,2012-09-14
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select Zl_Pati_Check_Execute(3,[1],[2]) as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExistWaitBool", lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        ExistWaitBool = Nvl(rsTmp!����)
    End If
    
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Public Function ExistWaitTest(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
''���ܣ������Ѿ����ִ�еļ���ͼ�飬�жϼ��鼰��鱨���Ƿ���д
''���أ�����ͼ����Ŀ��ִ�в���
''������:������
''����ţ�51613,2012-17-10
''����ţ�69009,2013-01-03,ȡ���ü��(�Ѿ������˶�δִ����Ŀ�ļ��)
'    Dim strSQL As String
'    Dim rsTmp As New ADODB.Recordset
'    Dim strDrug As String, strStuff As String, strTest As String
'    On Error GoTo errH
'
'    '�жϼ��鼰��鱨���Ƿ���д
'    strSQL = _
'        " Select Distinct C.���, C.���� As ��Ŀ, d.���� As ����" & vbNewLine & _
'        " From ����ҽ����¼ A,����ҽ������ B,����ҽ������ E,������ĿĿ¼ C,���ű� D" & vbNewLine & _
'        " Where a.����id = [1] And Nvl(a.��ҳid, 0) = [2]" & vbNewLine & _
'        "   And a.Id = b.ҽ��id And A.ID = E.ҽ��ID(+) And B.ִ��״̬ = 1 And a.������Ŀid = c.Id" & vbNewLine & _
'        "   And b.ִ�в���id + 0 = d.Id(+) And E.ҽ��ID is null" & vbNewLine & _
'        "   And (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null)" & vbNewLine & _
'        "   And exists (select ID from ����ҽ����¼" & vbNewLine & _
'        "         where (������� = 'C' And ���id Is not Null And A.ID = ���ID)" & vbNewLine & _
'        "            OR (������� = 'D' And ���id Is Null And A.ID = ID))" & vbNewLine & _
'        " order by C.���, C.����"
'
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�жϼ��鼰��鱨���Ƿ���д", lng����ID, lng��ҳID)
'    strDrug = "": strStuff = ""
'    Do While Not rsTmp.EOF
'            If UCase(rsTmp!���) = "E" Then '������Ŀ�����������ΪE�ı걾ҽ��Ϊ����¼
'                If strDrug = "" Then
'                    strDrug = Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ������]") & "δ��д"
'                Else
'                    If InStr(1, vbCrLf & strDrug & vbCrLf, vbCrLf & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ������]") & "δ��д" & vbCrLf) = 0 Then
'                        If LenB(StrConv(strDrug & vbCrLf & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ������]") & "δ��д", vbFromUnicode)) <= 1000 Then
'                            strDrug = strDrug & vbCrLf & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ������]") & "δ��д"
'                        Else
'                            strDrug = strDrug & vbCrLf & "... ..."
'                        End If
'                    End If
'                End If
'            Else
'                If strStuff = "" Then
'                    strStuff = Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ������]") & "δ��д"
'                Else
'                    If InStr(1, vbCrLf & strStuff & vbCrLf, vbCrLf & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ������]") & "δ��д" & vbCrLf) = 0 Then
'                        If LenB(StrConv(strStuff & vbCrLf & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ������]") & "δ��д", vbFromUnicode)) <= 1000 Then
'                            strStuff = strStuff & vbCrLf & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ������]") & "δ��д"
'                        Else
'                            strStuff = strStuff & vbCrLf & "... ..."
'                        End If
'                    End If
'                End If
'            End If
'
'    rsTmp.MoveNext
'    Loop
'    If strDrug <> "" Then
'        strDrug = "����δ��д�ļ��鱨�棺" & vbCrLf & vbCrLf & strDrug
'    End If
'    If strStuff <> "" Then
'        strStuff = "����δ��д�ļ�鱨�棺" & vbCrLf & vbCrLf & strStuff
'    End If
'    strTest = ""
'    If strDrug <> "" And strStuff <> "" Then
'      strTest = strDrug & vbCrLf & vbCrLf & strStuff
'    ElseIf strDrug <> "" Then
'      strTest = strDrug
'    ElseIf strStuff <> "" Then
'      strTest = strStuff
'    End If
'
'    ExistWaitTest = strTest
'
'    Exit Function
'errH:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
'End Function

Public Function ExistNurseData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal dOutTime As Date) As String
'����:��鲡�˳�Ժʱ��֮���Ƿ���ڻ�������
    Dim strSql As String
    Dim strDrug As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    dOutTime = CDate(Format(dOutTime + 1 / 24 / 60, "YYYY-MM-DD HH:MM") & ":00")
    strSql = "Select ID From ���˻����ļ� Where ����ID=[1] and ��ҳID=[2] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ�����°滤���ļ�", lng����ID, lng��ҳID)
    If rsTemp.RecordCount > 0 Then
        '�°����:���˱�����ֱ�Ӽ���Ժ֮����ڻ�������,Ӥ��ֻ���ĸӤͬʱ��Ժ�����
        strSql = _
            " Select Distinct NVL(Ӥ��,0) ���, �ļ�����" & vbNewLine & _
            " From ���˻����ļ� a, ���˻������� b" & vbNewLine & _
            " Where a.Id = b.�ļ�id And b.����ʱ�� >= [3] And a.����id = [1] And a.��ҳid = [2] And" & vbNewLine & _
            "      (Nvl(a.Ӥ��, 0) = 0 Or" & vbNewLine & _
            "      (Nvl(a.Ӥ��, 0) <> 0 And Not Exists" & vbNewLine & _
            "       (Select e.Ӥ��" & vbNewLine & _
            "         From ����ҽ����¼ e, ������ĿĿ¼ f" & vbNewLine & _
            "         Where e.������Ŀid + 0 = f.Id And e.ҽ��״̬ = 8 And Nvl(e.Ӥ��, 0) <> 0 And f.��� = 'Z' And" & vbNewLine & _
            "               Instr([4], ',' || f.�������� || ',', 1) > 0 And e.����id = a.����id And e.��ҳid = a.��ҳid And e.Ӥ�� = a.Ӥ��)))"
            
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�������ݼ��", lng����ID, lng��ҳID, dOutTime, ",3,5,11,")
        If rsTemp.EOF Then ExistNurseData = "": Exit Function
        rsTemp.Sort = "���"
        Do While Not rsTemp.EOF
            If strDrug = "" Then
                strDrug = Nvl(rsTemp!�ļ�����) & IIf(Val(rsTemp!���) = 0, "", Space(6) & "Ӥ�����:" & Val(rsTemp!���))
            Else
                strDrug = strDrug & vbCrLf & Nvl(rsTemp!�ļ�����) & IIf(Val(rsTemp!���) = 0, "", Space(6) & "Ӥ�����:" & Val(rsTemp!���))
            End If
        rsTemp.MoveNext
        Loop
        
        If strDrug <> "" Then
            strDrug = "���ڻ������ݵ��ļ����ƣ�" & vbCrLf & vbCrLf & strDrug
        End If
    Else
        '�ϰ�:
        strSql = "Select 1 From ���˻����¼ Where ����id = [1] And ��ҳid = [2] And ������Դ = 2 And ����ʱ�� >= [3] And RowNum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�������ݼ��", lng����ID, lng��ҳID, dOutTime)
        If rsTemp.RecordCount > 0 Then
            strDrug = "OK"
        End If
    End If
    ExistNurseData = strDrug
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ExistWaitQuittance(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
'���ܣ����ڼ���Ƿ����δ������ʵĵ���
'���أ����Һ͵���
'������:������
'����ţ�61429,2013-11-11
    
    Dim strSql As String
    Dim strDrug As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    strSql = _
        " Select Distinct a.No, d.���� ��Ŀ, c.���� As ����" & vbNewLine & _
        " From סԺ���ü�¼ a, ���˷������� b, ���ű� c, �շ���ĿĿ¼ d" & vbNewLine & _
        " Where a.Id = b.����id And a.�շ�ϸĿid = d.Id And b.��˲���id = c.Id(+) And b.���ʱ�� Is Null And a.����id = [1] And" & vbNewLine & _
        "      Nvl(a.��ҳid, 0) = [2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����Ƿ����δ������ʵĵ���", lng����ID, lng��ҳID)
    strDrug = ""
    Do While Not rsTmp.EOF
        If strDrug = "" Then
            strDrug = "����[" & Nvl(rsTmp!NO) & "]�е�" & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ֪����]") & "δ���"
        Else
            If InStr(1, vbCrLf & strDrug & vbCrLf, vbCrLf & "����[" & Nvl(rsTmp!NO) & "]�е�" & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ֪����]") & "δ���" & vbCrLf) = 0 Then
                If LenB(StrConv(strDrug & vbCrLf & "����[" & Nvl(rsTmp!NO) & "]�е�" & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ֪����]") & "δ���", vbFromUnicode)) <= 1000 Then
                    strDrug = strDrug & vbCrLf & "����[" & Nvl(rsTmp!NO) & "]�е�" & Nvl(rsTmp!��Ŀ) & "����" & Nvl(rsTmp!����, "[δ֪����]") & "δ���"
                Else
                    strDrug = strDrug & vbCrLf & "... ..."
                End If
            End If
        End If
    rsTmp.MoveNext
    Loop
    
    If strDrug <> "" Then
        strDrug = "����δ������ʵĵ��ݣ�" & vbCrLf & vbCrLf & strDrug
    End If
    ExistWaitQuittance = strDrug
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExistFeeInsurePatient(lng����ID As Long) As Boolean
'���ܣ��ж�ҽ�������Ƿ����δ�����
'���أ�
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    strSql = "Select Nvl(sum(B.�������),0) ������� From ������Ϣ A,������� B Where A.����ID=B.����ID And Nvl(A.����,0)<>0 And A.����ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlPatient", lng����ID)
    
    If Not rsTmp.EOF Then ExistFeeInsurePatient = (rsTmp!������� <> 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNurseGrade() As ADODB.Recordset
'���ܣ���ȡ����ȼ�
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ID,����,���� From �շ���ĿĿ¼" & _
        " Where ���='H' And ��Ŀ����>=1 And (����ʱ�� is NULL or Trunc(����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by ����"
    Set GetNurseGrade = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'����27370 by lesfeng 2010-01-26
Public Function InputDept(ByRef frmParent As Object, ByVal fra��Ժ As Control, ByVal obj As Control, ByVal str���� As String, ByVal str������� As String, _
ByVal strInput As String, ByRef blnCancel As Boolean, Optional ByVal intFlag = -1, Optional ByVal lngDeptID = 0, Optional ByVal bln������Ա���� As Boolean = False) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ����ָ�����ʵĲ����б�
    '���:frmParent ���ڶ���
    '     fra��Ժ �ؼ� ��������� fra��Ժ ��Ҫ���㵯��ѡ������λ��
    '     obj �ؼ� ��������� cbo��Ժ���� ���� cbo��Ժ����
    '     str����='�ٴ�','����','��ҩ��',...,����Ϊ��
    '     str�������:��,����:��1,3
    '     strInput ֧��������롢���롢���ƽ���ƥ��
    '     blnCancel ���سɹ����
    '     intFlag �ж����뷽ʽ����ѡ���һ�����ѡ���� ��ʼ-1�������������Ҷ�Ӧ
    '     lngDeptId ��intFlag��Ϊ-1ʱ�������������Ҷ�Ӧ�Ŀ��һ��߲�����ID
    '     bln������Ա����-����Ա����������
    '����:
    '����:
    '����:lesfeng
    '����:2010-01-25 16:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim lngTxtHeight As Long, vPoint As POINTAPI
    Dim strNo As String, strInputN As String
    Dim strFrom As String, strWhere As String
    
    On Error GoTo errH
    
    vPoint = zlControl.GetCoordPos(fra��Ժ.hWnd, obj.Left, obj.Top)
    lngTxtHeight = obj.Height
    
    strInputN = gstrLike & strInput & "%"
    strNo = strInput & "%"
                    
    str���� = Replace(str����, "'", "")
    If str���� <> "" Then
        If InStr(1, str����, ",") > 0 Then
            strSql = " And Instr(','||[1]||',',','||B.��������||',')>0"
        Else
            strSql = " And B.�������� = [1]"
        End If
    End If
    
    If zlCommFun.IsCharChinese(strInput) Or InStr(1, strInput, "-", 0) <> 0 Then
        strSql = strSql & " And (A.���� Like [4] or A.����||'-'||A.���� Like [4])" '���뺺��ʱֻƥ������
    Else
        strSql = strSql & " And (A.���� Like [5] Or A.���� Like [4] Or A.���� Like [4])"
    End If
    
    If intFlag = -1 Then
        strFrom = ""
        strWhere = ""
    ElseIf intFlag = 1 Then '�������� gbln��ѡ����
        strFrom = ",�������Ҷ�Ӧ D"
        strWhere = " And A.ID = D.����ID And D.����ID = [6]"
    Else '�������� ��ѡ����
        strFrom = ",�������Ҷ�Ӧ D"
        strWhere = " And A.ID = D.����ID And D.����ID = [6]"
    End If

    If bln������Ա���� Then strSql = strSql & "  And A.id=C.����ID and C.��Աid =[3]"
    
    strSql = " Select 1 as ����ID, A.ID,A.����,A.����,A.����,B.��������,B.�������" & _
        " From ���ű� A,��������˵�� B " & IIf(bln������Ա����, ",������Ա C", "") & strFrom & _
        " Where B.����ID=A.ID And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And  Instr(','||[2]|| ',',','|| B.�������|| ',')>0 " & strSql & strWhere & _
        " Order by A.����  Desc"
        '" And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) Order by A.����  Desc"

    Set InputDept = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "����", 1, "", "��ѡ��", False, False, True, vPoint.X, vPoint.Y, lngTxtHeight, blnCancel, False, True, str����, str�������, UserInfo.ID, strInputN, strNo, lngDeptID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'����27370 by lesfeng 2010-02-03
Public Function InputDoctors(ByRef frmParent As Object, ByVal fra��Ժ As Control, ByVal obj As Control, ByVal bytType As Byte, ByVal str������� As String, _
ByVal strInput As String, ByRef blnCancel As Boolean, Optional ByVal strUnits As String = "") As ADODB.Recordset
'---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ����ʿ�б�.
    '���:frmParent ���ڶ���
    '     fra��Ժ �ؼ� ��������� fra��Ժ ��Ҫ���㵯��ѡ������λ��
    '     obj �ؼ� ��������� cbo����ҽʦ
    '     bytType=0-ҽ����1-��ʿ
    '     str�������:��,����:��1,2,3
    '     strInput ֧�������š����롢���ƽ���ƥ��
    '     blnCancel ���سɹ����
    '     strUnits=���һ���ID��,��:18,26,31
    '����:
    '����:
    '����:lesfeng
    '����:2010-01-25 16:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim lngTxtHeight As Long, vPoint As POINTAPI
    Dim strNo As String, strInputN As String
    Dim strFrom As String, strWhere As String
    
    On Error GoTo errH
    
    vPoint = zlControl.GetCoordPos(fra��Ժ.hWnd, obj.Left, obj.Top)
    lngTxtHeight = obj.Height
    
    strInputN = gstrLike & strInput & "%"
    strNo = strInput & "%"
    
    On Error GoTo errH
    If strUnits <> "" Then
        If InStr(1, strUnits, ",") > 0 Then
            strSql = " And Instr(','|| [3] || ',',',' || B.����ID || ',')>0"
        Else
            strSql = " And B.����ID=[3]"
        End If
    End If
    
    If zlCommFun.IsCharChinese(strInput) Or InStr(1, strInput, "-", 0) <> 0 Then
        strSql = strSql & " And (A.���� Like [4] or A.����||'-'||A.���� Like [4])" '���뺺��ʱֻƥ������
    Else
        strSql = strSql & " And (A.��� Like [5] Or A.���� Like [4] Or A.���� Like [4])"
    End If
    
    strSql = "Select Distinct A.ID,A.���,A.����,A.����,C.��Ա����" & _
             "  From ��Ա�� A,������Ա B,��Ա����˵�� C,��������˵�� D" & _
             " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա���� = [1] And B.����ID=D.����ID" & _
             "   And  Instr(','||[2]|| ',',','|| D.�������|| ',')>0 " & strSql & _
             "   And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
             "   And (A.վ��=[6] Or A.վ�� is Null)" & _
             " Order by ����"
    Set InputDoctors = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "ҽ��ѡ��", 1, "", "��ѡ��", False, False, True, vPoint.X, vPoint.Y, lngTxtHeight, blnCancel, False, True, IIf(bytType = 0, "ҽ��", "��ʿ"), str�������, strUnits, strInputN, strNo, gstrNodeNo)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDeptDoctors(ByVal lng����ID As Long) As String
'���ܣ���ȡָ������������������ҽ��/��ʿIDs
'���أ�ҽ��ID1,ҽ��ID2,...
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    
    On Error GoTo errH
    'by lesfeng 2010-01-11 �����Ż�
    strSql = "Select Distinct A.ID From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա���� IN('ҽ��','��ʿ') " & _
                " And B.����ID=[1] And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
                " And (A.վ��=[2] Or A.վ�� is Null)" & _
                " Order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, gstrNodeNo)
    
    strSql = ""
    For i = 1 To rsTmp.RecordCount
        strSql = strSql & "," & rsTmp!ID
        rsTmp.MoveNext
    Next
    GetDeptDoctors = Mid(strSql, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetArea(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'���ܣ���ȡ�����б��ѡ��ĵ���
'������
    Dim strSql As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtInput.hWnd)
    If Not blnShowAll Then
        strSql = " Select ���� as ID,����,����,���� From ����" & _
                 " Where (���� Like [1] Or upper(����) Like '" & gstrLike & "'||[1]||'%' Or ���� Like '" & gstrLike & "'||[1]||'%') And  NVL(����,0)<3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "����", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSql = "Select ���� as ID,����,����,���� From ���� Where NVL(����,0)<3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "����", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAddress(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'���ܣ���ȡ�����б��ѡ��ĵ���
'������
    Dim strSql As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    If Not blnShowAll Then
        strSql = " Select ���� as ID,����,����,���� From ����" & _
                 " Where ���� Like [1] Or ���� Like [1] Or ���� Like [1]"
        vRect = zlControl.GetControlRect(txtInput.hWnd)
        Set GetAddress = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "��ַ", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSql = " Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
                " Substr(����,1,2) as ���� From ����" & _
                " Union All" & _
                " Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
                " From ���� Order by ����"
        Set GetAddress = zlDatabase.ShowSQLSelect(frmParent, strSql, 2, "��ַ", True, txtInput.Text, "", True, True, False, 0, 0, 0, blnCancel, True, True)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Getҽ�ƻ���(ByVal objText As Object, ByVal objFrom As Form, ByVal bytStyle As Byte, ByVal strCaption As String, ByVal strMsg As String, ByRef vPoint As POINTAPI, ByVal blnCancel As Boolean)
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����,ĩ�� From ҽ�ƻ��� Order By ����"
    Set rsTmp = zlDatabase.ShowSelect(objFrom, strSql, bytStyle, strCaption, , , , , True, True, vPoint.X, vPoint.Y, objText.Height, blnCancel)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û������""" & strCaption & """���ݣ����ȵ�" & strMsg & "�����á�", vbInformation, gstrSysName
        End If
        objText.Tag = ""
        zlControl.ControlSetFocus objText
    Else
        objText.Text = rsTmp!����
        zlControl.ControlSetFocus objText
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetSpcҽ�ƻ���(ByVal objText As Object, ByVal objFrom As Form, ByVal strCaption As String, ByVal strSeek As String, ByVal strNote As String, ByVal blnCancel As Boolean, ByVal blnĩ�� As Boolean, ByRef vPoint As POINTAPI)
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    If zlCommFun.IsCharChinese(objText.Text) Then
        strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where ���� Like [1]"
    Else
        If gbytCode = 1 Then
            strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where zlWbCode(����) Like [1]"
        ElseIf gbytCode = 0 Then
            strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where ���� Like [1]"
        End If
    End If
    Set rsTmp = zlDatabase.ShowSQLSelect(objFrom, strSql, 0, strCaption, blnĩ��, strSeek, strNote, False, _
        False, True, vPoint.X, vPoint.Y, objFrom.Height, blnCancel, False, False, _
        gstrLike & UCase(objText.Text) & "%")
    If Not rsTmp Is Nothing Then
        objText.Text = rsTmp!����
    Else
        objText.Tag = ""
        If gblnҽ�ƻ�������������¼�� Then
            MsgBox "���ֵ����δ�ҵ�������,������¼�룡", vbInformation, gstrSysName
            objText.Text = ""
            objText.SetFocus
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetOrgAddress(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'���ܣ���ȡ��Լ��λ�б�
'������
    Dim strSql As String, blnCancel As Boolean
    Dim vRect As RECT
    '����27040 by lesfeng �Ժ�Լ��λ���ϳ���ʱ��Ĵ���
    On Error GoTo errH
    If Not blnShowAll Then
        strSql = " Select ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                    " Where ĩ��=1 And (���� Like [1] Or ���� Like [1] Or ���� Like [1]) " & _
                    " and (����ʱ�� IS NULL OR TO_CHAR(����ʱ��, 'yyyy-MM-dd') = '3000-01-01') "
        vRect = zlControl.GetControlRect(txtInput.hWnd)
        Set GetOrgAddress = zlDatabase.ShowSQLSelect(frmParent, strSql, 0, "��λ", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSql = " Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & _
                 "  Where (����ʱ�� IS NULL OR TO_CHAR(����ʱ��, 'yyyy-MM-dd') = '3000-01-01') " & _
                 " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
        Set GetOrgAddress = zlDatabase.ShowSQLSelect(frmParent, strSql, 2, "��λ", True, txtInput.Text, "", True, True, False, 0, 0, 0, blnCancel, True, True)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPatiBeds(ByVal lng����ID As Long, Optional ByVal strBed As String) As ADODB.Recordset
'���ܣ���ȡ������ռ�õĴ�λ����Ϣ
'������lng����ID=����ID
'���أ�����ռ�õĴ�λ��¼��
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select A.����, A.�����, A.�Ա����, A.�ȼ�id As ��λ�ȼ�id, B.���� As ��λ�ȼ�, A.����ID, A.����, A.״̬, C.�Ա�" & vbNewLine & _
        "       From ��λ״����¼ A, �շ���ĿĿ¼ B, ������Ϣ C" & vbNewLine & _
        "       Where A.�ȼ�id = B.ID And A.����ID = C.����ID(+) And A.����id = [1]" & IIf(strBed = "", "", " And ���� = [2]")
    'ע��:��ͥ��������û�д�λ
    Set GetPatiBeds = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, strBed)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFreeBeds(ByVal lng����ID As Long, ByVal lng����ID As Long, str�Ա� As String, Optional lng����ID As Long = 0) As ADODB.Recordset
'���ܣ���ȡָ�������Ϳ��ҵĿմ�
'������
    Dim strSql As String, strTmp As String
    
    If InStr(str�Ա�, "��") > 0 Then
        strTmp = "�д�,���޴�"
    ElseIf InStr(str�Ա�, "Ů") > 0 Then
        strTmp = "Ů��,���޴�"
    Else
        strTmp = "���޴�"
    End If
    
    On Error GoTo errH
    
'    Select ����, a.�����, �Ա����, ��λ����, �ȼ�id, ����, b.�Ա�
'From (Select ����, �����, �Ա����, ��λ����, �ȼ�id, ����
'       From (Select ����, �����, �Ա����, ��λ����, �ȼ�id, ����
'              From ��λ״����¼
'              Where ����id = 57 And ����id Is Null And ״̬ = '�մ�' And Instr('�д�,���޴�', �Ա����) > 0 And (����id = 57 Or ����id Is Null)
'              Union
'              Select ����, �����, �Ա����, ��λ����, �ȼ�id, ����
'              From ��λ״����¼
'              Where ����id = 701100 And ���� = 1 And ����id = 57)
'       Order By LPad(����, 10, ' ')) A,
'     (Select m.�����, Wmsys.Wm_Concat(Distinct n.�Ա�) As �Ա�
'       From ��λ״����¼ M, ������Ϣ N
'       Where m.����id = n.����id(+) And m.����id = 57 And ����� Is Not Null
'       Group By m.�����) B
'Where a.����� = b.�����(+)

    strSql = "Select ����, a.�����, �Ա����, ��λ����, �ȼ�id, ��λ�ȼ�, ����, b.�Ա� From( " & _
        "Select ����, �����, �Ա����, ��λ����, �ȼ�id, ��λ�ȼ�, ����" & vbNewLine & _
        "From (" & vbNewLine & _
                "Select ����, �����, �Ա����, ��λ����, �ȼ�id, J.���� AS ��λ�ȼ�, ����" & vbNewLine & _
                    "From ��λ״����¼ I, �շ���ĿĿ¼ J " & vbNewLine & _
                    "Where I.�ȼ�ID = J.ID And ����id = [1] And ����id Is Null And ״̬ = '�մ�' And Instr([3],�Ա����)>0" & _
                    IIf(lng����ID = 0, "", " And (����ID = [2] Or ����ID is Null)")
    If lng����ID <> 0 Then '-----------------------------------------------------------------���˿�����ס���ò������ò���ԭס��λ
        strSql = strSql & " Union " & _
                        "Select ����, �����, �Ա����, ��λ����, �ȼ�id, Q.���� AS ��λ�ȼ�, ����" & vbNewLine & _
                        "       From ��λ״����¼ P, �շ���ĿĿ¼ Q " & vbNewLine & _
                        "       Where P.�ȼ�ID = Q.ID And ����id = [4] And ���� = 1 And ����id = [1]"
    End If
    strSql = strSql & ") ORDER BY LPAD(����,10,' ')) A ," & _
            "(Select m.�����, f_List2str(Cast(COLLECT(Distinct n.�Ա�) as t_Strlist)) As �Ա�" & vbNewLine & _
            "       From ��λ״����¼ M, ������Ϣ N" & vbNewLine & _
            "       Where m.����id = n.����id(+) And m.����id = [1] And ����� Is Not Null" & vbNewLine & _
            "       Group By m.�����) B Where a.����� = b.�����(+)"

    Set GetFreeBeds = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng����ID, strTmp, lng����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDiagnosticOtherInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                                  ByVal str������� As String, ByVal str��¼��Դ As String) As ADODB.Recordset
'���ܣ���ȡָ�����˵�������ϼ�¼'
'����:
'�������-1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;
'        11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���
'��¼��Դ:1-������2-��Ժ�Ǽǣ�3-��ҳ����(����ҽ��վ,���ժҪ);

    On Local Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    If InStr(1, "," & str��¼��Դ & ",", ",3,") > 0 And _
        (InStr(1, "," & str������� & ",", ",1,") > 0 Or InStr(1, "," & str������� & ",", ",11,") > 0) Then
        strSql = " Select A.�������,A.��¼��Դ,A.�������,A.����ID,A.���ID,A.��Ժ���,A.��¼����,A.�Ƿ�����,A.��ϴ���,C.���� From ������ϼ�¼ A,��������Ŀ¼ C " & _
                 " Where A.����ID=[1] And Nvl(A.��ҳID,0)=[2] And A.����ID = C.ID(+)" & _
                 " And A.��ϴ���>1  And NVL(A.�������,1) = 1 And instr([3],','||A.�������||',')>0 And A.��¼��Դ in (" & str��¼��Դ & ")" & _
                 " Union " & _
                 " Select a.�������,a.��¼��Դ,a.�������,a.����ID,a.���ID,a.��Ժ���,a.��¼����,a.�Ƿ�����,a.��ϴ���,C.���� From ������ϼ�¼ a,���˹Һż�¼ b,��������Ŀ¼ C  " & _
                 " Where a.����ID=[1] And A.����ID = C.ID(+) " & _
                 " And a.����ID=b.����ID And b.�Ǽ�ʱ��>trunc(sysdate-3) And a.��ҳID=b.id" & _
                 " And a.��ϴ���>1 And NVL(A.�������,1) = 1 And instr([3],','||a.�������||',')>0 And a.��¼��Դ=3 And B.��¼����=1 and B.��¼״̬=1" & _
                 " Order by ��ϴ��� Asc"
    Else
        
        strSql = " Select A.�������,A.��¼��Դ,A.�������,A.����ID,A.���ID,A.��Ժ���,A.��¼����,A.�Ƿ�����,A.��ϴ���,C.���� From ������ϼ�¼ A,��������Ŀ¼ C " & _
                 " Where A.����ID=[1] And Nvl(A.��ҳID,0)=[2] And A.����ID = C.ID(+)" & _
                 " And A.��ϴ���>1 And NVL(A.�������,1) = 1 And instr([3],','||A.�������||',')>0 And A.��¼��Դ in (" & str��¼��Դ & ")" & _
                 " Order by ��ϴ��� Asc"
                 
        '��ϴ���-��Ժʱ,������ҳ�����п�����д��Ҫ���,��Ҫ��ϵȶ�����¼
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlInPatient", lng����ID, lng��ҳID, "," & str������� & ",")
    If Not rsTmp.EOF Then Set GetDiagnosticOtherInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SetVsFlexGridChangeHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid, lngNo As Long)
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1
        If lngNo = 0 Then
            .FixedCols = 0
            .Cols = .FixedCols + UBound(arrHead) + 1
            .Rows = .FixedRows + 1
        Else
            .FixedCols = 1
            .Cols = .FixedCols + UBound(arrHead)
            .Rows = .FixedRows + 1
        End If

        For i = 0 To UBound(arrHead)
            If .FixedCols > 0 Then
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            Else
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            End If
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               'Ϊ��֧��zl9PrintMode
                If .FixedCols > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .colAlignment(i) = Val(Split(arrHead(i), ",")(2))
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(.FixedCols + i) = False
                    .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                    .colAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
'                    .ColData
                    'Ϊ��֧��zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  'Ϊ��֧��zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
                End If
            End If
            .ColData(i) = Val(Split(arrHead(i), ",")(3)) '��������Ϊ����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        Next
        
        '�̶������־���
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
    End With
End Sub
 
Public Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '����:������
    '����:vsGrid-�����е�������
    '     lngRow-��ǰ��
    '     blnBefor-��lngrow֮���֮��.true:֮��,false-֮��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        If blnBefor Then
            .AddItem "", lngRow
        Else
            .AddItem "", lngRow + 1
        End If
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = lngRow
            Else
                .Row = lngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlPvVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng���� As Long = -1, Optional lngβ�� As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1, Optional strValue As String)
    ', Optional strHeadMove As String
    '-----------------------------------------------------------------------------------------------------------

    '����:�ƶ���Ԫ�����
    '���:blnEdit-��ǰ�����ڱ༭״̬,����������
    '     lng����-����,���<0,������Ϊ0��,����Ϊָ������
    '     lngβ��-β��,���<0,������Ϊ.cols-1,����Ϊָ������
    '����:lngRow-������ڲ�����,�򷵻ر�������к�,���򷵻�-1
    '����:
    '����:���˺�
    '����:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------

    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    Dim lngValue As Long
    Dim arrHead As Variant
    Dim j As Long
    Dim lngColValue As Long
    
    Err = 0: On Error GoTo ErrHand:
    'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)

    If lng���� <> -1 Then
        lngCol = lng����
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lngβ�� < 0, vsGrid.Cols - 1, lngβ��)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                If IsNull(strValue) Or strValue = "" Then
                    arrSplit = Split(.ColData(i) & "||", "||")
                    If IsNull(arrSplit(1)) Or Trim(arrSplit(1)) = "" Then
                        lngValue = 0
                    Else
                        lngValue = Val(arrSplit(1))
                    End If
                Else
                    arrHead = Split(strValue, ";")
                    For j = 0 To UBound(arrHead)
                        lngValue = 1
                        lngColValue = Val(Split(arrHead(j), "||")(0))
                        If i = lngColValue Then
                            lngValue = Val(Split(arrHead(j), "||")(1))
                            Exit For
                        End If
                    Next
                End If
                If .ColHidden(i) Or lngValue >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
ErrHand:
End Sub

Public Function zl_VsGrid_SaveToPara(ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, _
ByVal lngMoudel As Long, ByVal strParaName As String, Optional ByVal bln˽�� As Boolean = True, _
    Optional ByVal blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:����vsFlex�Ŀ�ȵ�������
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     lngMoudel-ģ���
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------

    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If blnǿ�ƻָ����� = False Then
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    End If

    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '�����ʽ:������,�п�,������|������,�п�,������|...
    zlDatabase.SetPara strParaName, strCol, glngSys, lngMoudel ', bln˽��
    zl_VsGrid_SaveToPara = True
End Function

Public Function zl_VsGrid_FromParaRestore(ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal lngMoudle As Long, _
    ByVal strParaName As String, Optional bln˽�� As Boolean = True, _
    Optional ByVal blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�Ӳ������лָ�����Ŀ��
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     lngMoudle-ģ���
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------

    Dim strParaValue As String, intCols As Integer, arrReg As Variant, arrtemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String

    If blnǿ�ƻָ����� = False Then
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    End If

    strParaValue = zlDatabase.GetPara(strParaName, glngSys, lngMoudle, "")
    If strParaValue = "" Then Exit Function
    
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...

    Err = 0: On Error GoTo ErrHand:

    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            arrtemp = Split(arrReg(intCol) & ",,", ",")
            strColName = arrtemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(arrtemp(1))
                If Val(arrtemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_VsGrid_FromParaRestore = True
    Exit Function
ErrHand:
End Function
Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub
Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo ErrHand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
ErrHand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("������������", , , True)) = 1
    Err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
ErrHand:
End Function
Public Function GetBalanceDeposit(ByVal lngBalanceID As Long, ByVal blnNOMoved As Boolean) As ADODB.Recordset
    '���ܣ���ȡһ�Ž��ʵ��ݵĳ�Ԥ����¼
    Dim strSql As String
    On Error GoTo errH
    strSql = _
        "Select A.ID,A.NO ���ݺ�,A.ʵ��Ʊ�� Ʊ�ݺ�,To_Char(A.�տ�ʱ��,'YYYY-MM-DD') as ����,A.���㷽ʽ," & _
        " Ltrim(To_Char(A.��Ԥ��,'9999999990.00')) as ���" & _
        " From " & IIf(blnNOMoved, "H", "") & "����Ԥ����¼ A " & _
        " Where mod(A.��¼����,10)=1 And A.����ID = [1]  " & _
        " Order by A.����,A.���㷽ʽ"
    Set GetBalanceDeposit = zlDatabase.OpenSQLRecord(strSql, App.ProductName, lngBalanceID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlIsExistsSquareCard(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ�Ϊ�����㵥��
    '���:strNos-���ݺ�(����Ϊ����,�ö��ŷ���)
    '����:
    '����:����,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, strNoIns As String
    strNoIns = Replace(strNos, "'", "")
    On Error GoTo errHandle
    strSql = "Select A.ID As ������id " & _
    "   From ���˿������¼ A, ����Ԥ����¼ B, " & _
    "        (Select Column_Value From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist))) J " & _
    "   Where A.����id = B.ID and ( B.��¼����=2 or B.��¼����=12) And B.NO = J.Column_Value And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�����ʵ��Ƿ����ˢ����¼", strNoIns)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSql As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSql, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSql As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSql = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSql, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub
Public Function StrToNum(ByVal strNumber As String) As Double
    '����:���ַ���ת��������
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function
Public Function zl_Getҽ�ƿ�����(lngTypeId As Long) As String()
    '-----------------------------------------------------------------------------------------------------------
    '����:����ҽ������ID��ȡҽ������
    '���:lngTypeID-ҽ�ƿ�����ID
    '����:���Ͷ���
    '����:����
    '����:2012-07-06
    '�����:51072
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    Dim arr(3) As String
    
    strSql = "" & _
    "       Select ���볤��,������������,�Ƿ�ȱʡ���� " & _
    "       From ҽ�ƿ���� " & _
    "       Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡҽ�ƿ����", lngTypeId)
    If rsTemp Is Nothing Then zl_Getҽ�ƿ����� = arr: Exit Function
    If rsTemp.RecordCount <= 0 Then zl_Getҽ�ƿ����� = arr: Exit Function
    rsTemp.MoveFirst
    arr(0) = Nvl(rsTemp!���볤��, "0")
    arr(1) = Nvl(rsTemp!������������, "0")
    arr(2) = Nvl(rsTemp!�Ƿ�ȱʡ����, "0")
    zl_Getҽ�ƿ����� = arr
End Function

Public Function �Ƿ��Ѿ�ǩԼ(strCardNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ҫ�󶨵Ŀ����Ƿ��Ѿ�ǩԼ
    '���:�󶨿���
    '����:����
    '����:2012-08-31 11:32:14
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand:
    strSql = "" & _
    "   Select Count(1) as �Ƿ�ǩԼ From ����ҽ�ƿ���Ϣ Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "ҽ�ƿ���", strCardNO)
    �Ƿ��Ѿ�ǩԼ = rsTemp!�Ƿ�ǩԼ > 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function


Public Sub AddSQL�󶨿�(ByVal lng����ID As Long, �����ID As Long, strCard As String, strPassWord As String, ByVal dtCurdate As Date, blnICCard As Boolean, ByRef cllPro As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�󶨿�����
    '���:lng����ID;strCard-�󶨿���;strPassWord-��������
    '����:lngCard����ID-���ѵĽ���ID
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim str�䶯ԭ�� As String
    Dim strICCard As String
    
    strICCard = IIf(blnICCard, strCard, "")
    str�䶯ԭ�� = "���˹Һŷ���"
          'Zl_ҽ�ƿ��䶯_Insert
          strSql = "Zl_ҽ�ƿ��䶯_Insert("
          '      �䶯����_In   Number,
          '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
          strSql = strSql & "" & 11 & ","
          '      ����id_In     סԺ���ü�¼.����id%Type,
          strSql = strSql & "" & lng����ID & ","
          '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
          strSql = strSql & "" & �����ID & ","
          '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
          strSql = strSql & "'',"
          '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
          strSql = strSql & "'" & strCard & "',"
          '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
          '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
          strSql = strSql & "'" & str�䶯ԭ�� & "',"
          '      ����_In       ������Ϣ.����֤��%Type,
          strSql = strSql & "'" & strPassWord & "',"
          '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
          strSql = strSql & "'" & UserInfo.���� & "',"
          '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
          strSql = strSql & "to_date('" & Format(dtCurdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
          '      Ic����_In     ������Ϣ.Ic����%Type := Null,
          strSql = strSql & "'" & strICCard & "',"
          '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
          strSql = strSql & "NULL)"
     zlAddArray cllPro, strSql
End Sub

Public Function Getҽ�ƿ����ID(strTypeName As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ�ƿ����ID
    '���:strTypeName ҽ�ƿ��������
    '����:ҽ�ƿ����ID
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand
    strSql = "" & _
    "   Select ID From ҽ�ƿ���� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "ҽ�ƿ����", strTypeName)
    If rsTemp Is Nothing Then Getҽ�ƿ����ID = 0: Exit Function
    If rsTemp.RecordCount <= 0 Then Getҽ�ƿ����ID = 0: Exit Function
    Getҽ�ƿ����ID = rsTemp!ID
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zl��ǰ�û����֤�Ƿ��(lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ�û����֤�Ƿ��ѱ���
    '���:lng����ID
    '����:True �Ѱ� false δ��
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As Recordset
    On Error GoTo ErrHand
    strSql = "" & _
    " Select count(1) as �Ƿ�� From ������Ϣ A,����ҽ�ƿ���Ϣ B Where A.���֤�� =B.���� And A.����ID=B.����ID And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "ҽ�ƿ���", lng����ID)
    zl��ǰ�û����֤�Ƿ�� = rsTemp!�Ƿ�� > 0
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Public Function SetNullValue(varObj As Variant, Optional strDefault As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿ�ֵĬ��ֵ
    '���:varObj�����ֶζ���,strDefault Ĭ��ֵ
    '����:���ú��ֵ
    '����:����
    '����:2012-08-31 04:36:33
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If IsNull(varObj) Then
        SetNullValue = strDefault
        Exit Function
    End If
    SetNullValue = CStr(varObj)
End Function


Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: �رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

Public Function CheckBillExistReplenishData(intType As Integer, _
    Optional lngBalance As Long, Optional strNos As String, _
    Optional ByRef strReplenishNo As String, Optional ByRef blnErrBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ���ڶ��ν���
    '���:intType:0-�շ����ݣ�ʹ��lngBalanceΪ�������
    '     intType:1-�շ����ݣ�ʹ��strNosΪ���ݺ�
    '���Σ�
    '   strReplenishNo ������㵥�ݺ�
    '   blnErrBill �Ƿ��쳣���㵥��
    '����:True-���ڶ��ν������� False-�����ڶ��ν�������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strReplenishNo = ""
    If intType = 0 Then
        strSql = _
            " Select Max(a.NO) As No,Max(a.����״̬) As ����״̬" & vbNewLine & _
            " From ���ò����¼ A, (Select Distinct ����id From ����Ԥ����¼ Where ������� = [1]) B" & vbNewLine & _
            " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2"
        strSql = strSql & _
            " Union All" & _
            " Select Max(a.NO) As No,Max(a.����״̬) As ����״̬ From ���ò����¼ A Where a.������� = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����ν���", lngBalance)
    Else
        strSql = _
            " Select Max(a.NO) As No,Max(a.����״̬) As ����״̬" & vbNewLine & _
            " From ���ò����¼ A," & vbNewLine & _
            "      (Select Distinct ����id" & vbNewLine & _
            "       From ������ü�¼" & vbNewLine & _
            "       Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(f_Str2list([1])))) B" & vbNewLine & _
            " Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 0 And Nvl(a.����״̬,0) <> 2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����ν���", strNos)
    End If
    
    strReplenishNo = Nvl(rsTmp!NO)
    blnErrBill = Val(Nvl(rsTmp!����״̬)) = 1
    CheckBillExistReplenishData = strReplenishNo <> ""
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CreateSquareCardObject(ByRef frmMain As Object, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '��������
    '���˺�:���ӽ��㿨�Ľ���:ִ�л��˷�ʱ
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
End Sub

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthDay As String = "", Optional ByVal datCalc As Date) As String
    '����:����Ϸ��Լ��
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    strBirthDay = Format(strBirthDay, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthDay) Then
        If datCalc = CDate(0) Then
            strSql = "select Zl_Age_Check([1],[2]) From dual"
        Else
            strSql = "select Zl_Age_Check([1],[2],[3]) From dual"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Zl_Age_Check", strAge, CDate(strBirthDay), datCalc)
    Else
        strSql = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Zl_Age_Check", strAge)
    End If
    CheckAge = Nvl(rsTemp.Fields(0).Value)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreatePublicPatient() As Boolean
'����:����������Ϣ������������
    If gobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set gobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        If gobjPublicPatient Is Nothing Then
            MsgBox "����������Ϣ��������(zlPublicPatient.clsPublicPatient)ʧ��!", vbInformation, gstrSysName
        Else
            Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser)
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjPublicPatient Is Nothing Then CreatePublicPatient = True
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject(, "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
    
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String, Optional ByRef strErr As String = "0")
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    Dim strMsg As String
    
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        strMsg = "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description
        If strErr = "0" Then
            MsgBox strMsg, vbInformation, gstrSysName
        Else
            strErr = strMsg
        End If
    End If
End Sub

Public Function Get������ҳ�ӱ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str��Ϣ�� As String) As ADODB.Recordset
'���ܣ�
'    ��ȡ������ҳ�ӱ���
'����:
    Dim strSql As String
    Dim intRet As Integer
    
    intRet = UBound(Split(str��Ϣ��, ","))
    If intRet = -1 Then '��ȡ�������дӱ���Ϣ
        strSql = "Select ��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where ����ID =[1] And ��ҳID =[2] And ��Ϣֵ is Not Null"
    ElseIf intRet = 0 Then '��ȡָ��ĳ���ӱ���Ϣ
        strSql = "Select ��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where ����ID =[1] And ��ҳID =[2] and ��Ϣ��='" & Split(str��Ϣ��, ",")(0) & "'" & " And ��Ϣֵ is Not Null "
    ElseIf intRet > 0 Then '��ȡָ���Ķ���ӱ���Ϣֵ
        strSql = "Select ��Ϣ��, ��Ϣֵ" & vbNewLine & _
            "From ������ҳ�ӱ�" & vbNewLine & _
            "Where ����id = [1] And ��ҳid = [2] And" & vbNewLine & _
            "      ��Ϣ�� In (Select * From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) And ��Ϣֵ is Not Null "
    End If
    
    On Error GoTo errH
    Set Get������ҳ�ӱ� = zlDatabase.OpenSQLRecord(strSql, "��ȡ������ҳ�ӱ�", lng����ID, lng��ҳID, str��Ϣ��)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub LoadStructAddressDef(ByRef strAddress() As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������е�ȱʡ��ַ
    '���:PatiAddress-�ṹ����ַ�ؼ�
    '����:
    '����:��ΰ��
    '����:2016/1/7
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    On Error GoTo errH
    strSql = "Select ����,����,level From ���� " & _
            " Start With ȱʡ��־=1 " & _
            " Connect by Prior �ϼ�����=���� " & _
            " Order by level Desc "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ȱʡ����")
    If rsTmp.RecordCount = 0 Then Exit Sub
    Do While Not rsTmp.EOF
        strAddress(Val(Nvl(rsTmp!����))) = Nvl(rsTmp!����)
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReadStructAddress(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef PatiAddress As Object)
'����:��ȡ�ṹ����ַ
    Dim i As Long
    Dim rsStruct As ADODB.Recordset
    Dim rsAddress As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select a.ʡ, a.��, a.��, a.����, a.����, a.��ַ��� From ���˵�ַ��Ϣ A Where a.����id = [1] And NVL(a.��ҳid,0) = [2]"
    Set rsStruct = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˽ṹ����ַ", lng����ID, lng��ҳID)
    
    For i = PatiAddress.LBound To PatiAddress.UBound
        rsStruct.Filter = "��ַ���=" & i
        If rsStruct.RecordCount > 0 Then
            Call PatiAddress(i).LoadStructAdress(rsStruct!ʡ & "", rsStruct!�� & "", rsStruct!�� & "", rsStruct!���� & "", rsStruct!���� & "")
        Else
            If rsAddress Is Nothing Then
                'ͬһ������ֻ��ȡһ��
                If lng��ҳID <> 0 Then
                    strSql = "Select c.�����ص�, c.����, Nvl(b.��ͥ��ַ, c.��ͥ��ַ) As ��סַ, Nvl(b.���ڵ�ַ, c.���ڵ�ַ) As ���ڵ�ַ, Nvl(b.��ϵ�˵�ַ, c.��ϵ�˵�ַ) As ��ϵ�˵�ַ" & vbNewLine & _
                        "From ������ҳ B, ������Ϣ C" & vbNewLine & _
                        "Where b.����id = c.����id And b.����id = [1] And b.��ҳid = [2] "
                Else
                    strSql = "Select c.�����ص�, c.����, c.��ͥ��ַ As ��סַ,  c.���ڵ�ַ As ���ڵ�ַ,c.��ϵ�˵�ַ As ��ϵ�˵�ַ " & vbNewLine & _
                        "From ������Ϣ C" & vbNewLine & _
                        "Where c.����id = [1] "
                End If
                Set rsAddress = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˽ṹ����ַ", lng����ID, lng��ҳID)
            End If
            If rsAddress.RecordCount > 0 Then
                If Nvl(rsAddress.Fields(PatiAddress(i).Tag).Value, "") <> "" Then
                    PatiAddress(i).Value = Nvl(rsAddress.Fields(PatiAddress(i).Tag).Value, "")    '�������ýṹ����ַ֮ǰ������
                End If
            End If
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub CreateStructAddressSQL(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef arrSQL As Variant, ByRef PatiAddress As Object, Optional ByVal bytFunc As Byte = 0)
'����:�����ṹ����ַSQL
'����:
'PatiAddress-�ṹ����ַ�ؼ�������
'arrSQL-���ص�SQL���鼯��
'bytFunc ��ѡ����:=1���ؼ�ֵΪ��ʱ,����ɾ��
    Dim i As Long
    
    For i = PatiAddress.LBound To PatiAddress.UBound
        If PatiAddress(i).Value <> "" Then
            '����\�޸�
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(1," & lng����ID & "," & lng��ҳID & "," & i & ",'" & PatiAddress(i).valueʡ & "','" & PatiAddress(i).value�� & "','" & PatiAddress(i).value���� & "','" & PatiAddress(i).value���� & "','" & PatiAddress(i).value��ϸ��ַ & "','" & PatiAddress(i).Code & "')"
        Else
            'ɾ��
            If bytFunc = 1 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(2," & lng����ID & "," & lng��ҳID & "," & i & ")"
            End If
        End If
    Next

End Sub

Public Function CreateXWHIS(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��ж� RIS�ӿڲ���(zl9XWInterface.clsHISInner) �Ƿ���ڣ�������
'������blnMsg������ʧ��ʱ�Ƿ���ʾ

    If Not gblnXW Then Exit Function
    If Not gobjXWHIS Is Nothing Then CreateXWHIS = True: Exit Function
    
    On Error Resume Next
    Set gobjXWHIS = GetObject(, "zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    On Error Resume Next
    If gobjXWHIS Is Nothing Then Set gobjXWHIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    If gobjXWHIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CreateXWHIS = True
End Function

Public Function CreatePublicExpenseBillOperation() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ò���
    '���:
    '����:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjPublicExpenseBillOperation Is Nothing Then
        Set gobjPublicExpenseBillOperation = CreateObject("zlPublicExpense.clsBillOperation")
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)����ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Function
        End If
    Else
        CreatePublicExpenseBillOperation = True
        Exit Function
    End If
    If gobjPublicExpenseBillOperation Is Nothing Then Exit Function
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ�ż��������
    '���:lngSys-ϵͳ��
    '     cnOracle-���ݿ����Ӷ���
    '     strDBUser-���ݿ�������
    '����:��ʼ���ɹ�,����true,���򷵻�False
    If gobjPublicExpenseBillOperation.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then
         MsgBox "ע��:" & vbCrLf & "   ���ù�������(zl9PublicExpense)��ʼ��ʧ�ܣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
         Exit Function
    End If
    CreatePublicExpenseBillOperation = True
End Function

Public Function IsAppointPati(ByVal lngRegID As Long, ByRef strBedNO As String) As Boolean
'����:��鲡���Ƿ���ԤԼ���Ĳ���
'����:lngRegID-�Һ�ID;
'     strPatiInfo-������Ϣ  ��ʽ  ��λ��,��Ժ����ID,��Ժ����ID
'���Ե�ַ��http://192.168.32.201:8889/bizdomain/9e404039-9c1a-48a7-b283-093992bffe4a
'���Է���ֵ:[{"RGST_ID":76237,"RGST_NO":"S0000021","PID":84414,"PAT_NAME":"�󷽲���","IBA_PAT_SEX":"Ů","PAT_AGE":"30��","INP_BED_NO":"12","DSTRBT_INP_DEPT_ID":122,"DSTRBT_INP_DEPT":"��Ѫ���ڿ�","IBA_WARD":"��Ѫ���ڿ�","DSTRBT_INP_WARD_ID":122,"ORDER_ID":1214753,"HOME_PHNO":"123","CONTACTS_PHNO":"234"}]

    Dim strRet As String
    Dim blnRet As Boolean
    
    blnRet = Sys.NewSystemSvr("ԤԼ����", "ԤԼ���Ų�ѯ", "{""rgst_id_in"":""" & lngRegID & """}", strRet)
    If blnRet And strRet <> "" Then
        strRet = Mid(strRet, 2, Len(strRet) - 2)
        If strRet = "" Then Exit Function   'δ�ҵ���λ
        strBedNO = zlStr.JSONParse("INP_BED_NO", strRet)   '��λ��
    End If
    IsAppointPati = blnRet
End Function

