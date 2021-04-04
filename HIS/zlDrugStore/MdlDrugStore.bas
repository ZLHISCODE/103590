Attribute VB_Name = "MdlDrugStore"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gstrSQL As String
Public gobjBrower As Object

Public glngModul As Long
Public glngSys As Long                      'ϵͳ��Ų���
Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrUnitName As String               '�û���λ����
Public gstrDbUser As String                 '��ǰ�û�����
Public gstrprivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrMatchMethod As String            'ƥ�䷽ʽ:0��ʾ˫��ƥ��

Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������
Public gstrTryUse As String                 '���÷�
Public gbytSimpleCodeTrans As Byte          '��Ƭ�����Ƿ���������л�����

Public gobjCharge As Object                 '���۲���
Public gobjStuff As Object                  '���Ĳ���

Public Const gint����ҩ�� As Integer = 2
Public Const gintסԺҩ�� As Integer = 3

Public gobjESign As Object '����ǩ���ӿ�
Public gblnESign������ҩ As Boolean         '������ҩ�����Ƿ�����
Public gblnESign���ŷ�ҩ As Boolean         '���ŷ�ҩ�����Ƿ�����
Public gblnESignUserStoped As Boolean       '�û�����ǩ��֤���Ƿ�ͣ��

Public grsMaster As New ADODB.Recordset        'ҩƷѡ������ҩƷ��񻺴����ݼ�
Public grsMasterInput As New ADODB.Recordset   'ҩƷѡ������ҩƷ���¼�����ʱ�Ļ������ݼ�
Public grsSlave As New ADODB.Recordset         'ҩƷѡ���������λ������ݼ�

Public gstrPriceClass As String         '�۸�ȼ�

Public Enum EsignTache
    Dosage = 1  '��ҩ
    send = 2    '��ҩ
    returnStep = 3 '��ҩ
End Enum

Public Const DblFrmHeight As Double = 3630

Public Const glngRowByFocus = &HFFE3C8
Public Const glngRowByNotFocus = &HF4F4EA
Public Const glngFixedForeColorByFocus = &HFF0000
Public Const glngFixedForeColorNotFocus = &H80000012

Public Const gstrUnit_DYEY = "����ҽ�ƴ�ѧ�����ڶ�ҽԺ"
Public Const gstrUnit_DLSY = "�����е�������ҽԺ"

'���������ŷ�ҩ������ɫ����
Public Const glng��ҩ As Long = &HC0&
Public Const glng��ҩ As Long = &HC00000
Public Const glng���� As Long = &H80000008
Public Const strAsc As String = "��"                   '����
Public Const strDesc As String = "��"                  '����

'LED��ʾ��ر���
Public glngLEDModal As Long                'LEDģ�����
Public grsLEDComponent As New ADODB.Recordset  'LED���������ݿ���Ϣ
Public gobjLEDShow As Object               'LED����

'ģ���
Public Enum ģ���
    �⹺��� = 1300
    ������� = 1301
    ������� = 1302
    ��۵��� = 1303
    ҩƷ�ƿ� = 1304
    ҩƷ���� = 1305
    �������� = 1306
    ҩƷ�̵� = 1307
    ҩƷ�ƻ� = 1330
    �������� = 1331
End Enum


'�û���Ϣ------------------------
Public Type TYPE_USER_INFO
    �û�ID As Long
    �û����� As String
    �û����� As String
    �û����� As String
    ����ID As Long
    ���ű��� As String
    �������� As String
    strMaterial As String
End Type
Public UserInfo As TYPE_USER_INFO

'���ŷ�ҩ�и�����ɫ����
Public Enum mListColor
    LowerLimit = &HC000C0                       '���ڿ�����ޣ���ɫ
    SumTotal = vbBlue                           'С�ơ��ϼƣ���ɫ
    State_Send = &HFFDDDD                       '��ҩ״̬��ǳ��ɫ
    State_UnProcess = &H80000005                '������״̬����ɫ
    State_Reject = &HDBDBDB                     '�ܷ�״̬��ǳ��ɫ
    State_Shortage = &HD7D7FF                   'ȱҩ״̬��ǳ��ɫ
    State_RejectRestore = &HD7D7FF              '�ܷ��ָ�״̬��ǳ��ɫ
    State_RejectUnProcess = &H80000005          '�ܷ�������״̬����ɫ
    Return_Original = &H80000008                '��ҩ״̬��ԭʼ���ݣ�����ɫ
    Return_Sended = &HC00000                    '��ҩ״̬���ѷ�ҩ���ݣ�����ɫ
    Return_Returned = &HC0&                     '��ҩ״̬����ҩ���ݣ�����ɫ
End Enum

'ҩ��ģ��Ҫʹ�õ���ϵͳ����
Public Type Type_SysParms
    P6_δ��˼��ʴ�����ҩ As Integer
    P9_���ý���λ�� As Integer
    P15_�����շ��뷢ҩ���� As Integer
    P16_סԺ�����뷢ҩ���� As Integer
    P23_�ѽ��ʵ��ݲ��� As Integer
    P25_ʹ�õ���ǩ�� As Integer
    P26_����ǩ������ As String
    P28_���ﲡ������ʱ��Ҫˢ����֤ As String
    P29_ָ�������۶��۵�λ As Integer
    P44_����ƥ�� As String
    P54_ʱ��ҩƷ�ԼӼ������ As Integer
    P64_������� As Integer
    P68_����ҩ�������Ϻ���ҩ As Integer
    P70_�����Ǽ���Ч���� As Integer
    P75_�⹺�����Ҫ�˲� As Integer
    P76_ʱ��ҩƷֱ��ȷ���ۼ� As Integer
    P85_ҩ���鿴���ݳɱ��� As Integer
    P96_ҩƷ��¿��ÿ�� As Integer
    P98_���ʱ����������۷��� As Integer
    P126_ʱ��ҩƷ�ۼۼӳɷ�ʽ As Integer
    P148_δ�շѴ�����ҩ As Integer
    P149_Ч����ʾ��ʽ As Integer
    P150_ҩƷ���������㷨 As Integer
    P153_�������� As Long
    P163_��Ŀִ��ǰ�������շѻ��ȼ������ As Integer
    Para_���뷽ʽ As String
    P214_�״�ҽ��ִ����Ҫ���  As Integer
    P221_ҩƷ���ʱ�� As Integer
    P174_ҩƷ�ƿ���ȷ���� As Integer
    P175_ҩƷ������ȷ���� As Integer
    P222_ҩ���Զ�����ҩ�ӿ� As Integer
    P240_ҩ��������� As Integer
    P241_�������ʱ�� As Integer
    P275_���۹���ģʽ As Integer
    P213_��ҩ�䷽ÿ����ҩζ�� As Integer
    
End Type
Public gtype_UserSysParms As Type_SysParms     'ϵͳ����

'����ģ�����
Public gstrLike As String                       '����ƥ��
Public gblnMyStyle As Boolean                   '���Ի����

Public gint���뷽ʽ As Integer              '0-ƴ����1-���
Public gintҩƷ������ʾ As Integer          '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
Public gint����ҩƷ��ʾ As Integer          '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��

'ҵ�񵥾ݺ�
Public Enum ���ݺ�
    �⹺��� = 1
    ������� = 2
    Эҩ��� = 3
    ������� = 4
    ��۵��� = 5
    ҩƷ�ƿ� = 6
    ҩƷ���� = 7
    �շѴ�����ҩ = 8
    ���ʵ�������ҩ = 9
    ���ʱ�����ҩ = 10
    �������� = 11
    �̵�� = 12
    ���۱䶯 = 13
    �̵㵥 = 14
    �����¼ = 27
End Enum

'˽�С�����ģ�����
Public Enum ����_Э�����_˽��
    P1_�Ƿ�ѡ��ⷿ = 1
    P2_���̴�ӡ = 2
    P3_��˴�ӡ = 3
End Enum

Public Enum ����_ҩƷ����_˽��
    P1_�Ƿ�ѡ��ⷿ = 1
    P2_ҩƷ��λ = 2
    P3_���� = 3
    P4_���̴�ӡ = 4
    P5_��˴�ӡ = 5
    P6_��ѯ���� = 6
End Enum

Public Enum ����_������ҩ_˽��
    P1_������ = 1
    P2_���� = 2
End Enum

Public Enum ����_������ҩ_����
    P1_�շѴ�����ʾ��ʽ = 1
    P2_���ʴ�����ʾ��ʽ = 2
    P3_��ѯ���� = 3
    P4_������ɫ = 4
    P5_��ӡ�������ʵ� = 5
    P6_��ӡ�˷ѵ��ݼ�� = 6
    P7_��ӡ�ӳ� = 7
    P8_��ʾ�������� = 8
    P9_ˢ�¼�� = 9
    P10_У�鷢ҩ�� = 10
    P11_У�鷽ʽ = 11
    P12_У����ҩ�� = 12
    P13_�Զ����� = 13
End Enum

Public Enum ����_���ŷ�ҩ_˽��
    P1_������ = 1
    P2_���� = 2
End Enum

Public Enum ����_���ŷ�ҩ_����
    P1_��ѯ���� = 1
    P2_��ҩ���� = 2
    P3_��Ҫ���� = 3
    P4_��ҩ��ǩ�� = 4
    P5_ȱҩ��� = 5
    P6_��ҩ��ǩ�� = 6
    P7_������ҩ��ʽ = 7
    P8_�Զ�ˢ��δ��ҩ�嵥 = 8
End Enum

Public Enum ����_�������_����
    P1_����׼ = 1
End Enum

'ҩƷ���۸�������󾫶�
Public Type Type_Digits
    Digit_��� As Integer
    Digit_�ɱ��� As Integer
    Digit_���ۼ� As Integer
    Digit_���� As Integer
End Type
Public gtype_UserDrugDigits As Type_Digits

Public Type Type_SaleDigits
    Digit_�ɱ��� As Integer
    Digit_���ۼ� As Integer
    Digit_���� As Integer
End Type
Public gtype_UserSaleDigits As Type_SaleDigits

'���ݲ�������
Private Type Type_BillControl
    bln�Ƿ���� As Boolean
    intʱ������ As Integer
    bln���˵��� As Boolean
    dbl������� As Double
End Type
Private gtype_myBillControl As Type_BillControl


'�����������ƣ���˳����;�ָ�
Public Const gconstrRecipeType = "��ͨ;����;����;����;��һ;����"

'Ĭ�ϴ�����ɫ����ͨ����ɫ���������ɫ�����ƣ�����ɫ��������һ������ɫ����������ɫ
Private Const gconlng��ͨ = &HFFFFFF
Private Const gconlng���� = &HC0FFC0
Private Const gconlng���� = &HC0FFFF
Private Const gconlng���� = &HFFFFFF
Private Const gconlng��һ = &HC0C0FF
Private Const gconlng���� = &HC0C0FF

Public Type InOutType
    bln�⹺��� As Boolean
    bln������� As Boolean
    blnЭҩ��� As Boolean
    bln������� As Boolean
    bln��۵��� As Boolean
    blnҩƷ�ƿ� As Boolean
    blnҩƷ���� As Boolean
    bln�շѴ�����ҩ As Boolean
    bln���ʵ�������ҩ As Boolean
    bln���ʱ�����ҩ As Boolean
    bln�������� As Boolean
    bln�̵�� As Boolean
    bln���۱䶯 As Boolean
    bln�̵㵥 As Boolean
    blnҩƷ���� As Boolean
End Type
Public gInOutType As InOutType

Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    
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
    support��Ժ��ʵ�ʽ��� = 29       '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    support�����ֳ�����ϸ = 32    '�������סԺ���ʴ�����ÿ����ϸ���в��ֳ���
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    supportסԺ�������� = 34        'HISʼ����ΪסԺ֧�ֽ������ϣ������֧����ҽ���ӿ��ڲ��������ؼټ��ɣ����Ӹò�����Ϊ�����GetCapability�����������ֽ��㷽ʽ�Ƿ�֧��ȫ��
    support�������� = 35            '�Ƿ����������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
    support����_ָ��סԺ���� = 36   '�Ƿ�֧��ָ��סԺ��������ҽ������
    support����_ָ�����ڷ�Χ = 37   '�Ƿ�֧��ָ���������ڷ�Χ����ҽ������
    support����_����Ӥ�������� = 38 '�Ƿ���������Ӥ��������
    
    support������� = 41            '�Ƿ�֧������ҽ�����˵ļ��ʷ���ʹ��������������
End Enum

Public Sub setNOtExcetePrice()
    '����ʱ�仹δִ�е���ҩƷִ�е���
    Dim rstemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct i.Id As ҩƷid " & _
               " From �շ���ĿĿ¼ I, �շѼ�Ŀ N, ҩƷ��� P" & _
               " Where i.Id = n.�շ�ϸĿid And i.Id = p.ҩƷid And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & _
                   " n.�䶯ԭ�� = 0 And Sysdate>n.ִ������" & GetPriceClassString("N") & _
               " Union " & _
               " Select Distinct a.ҩƷid From ҩƷ�۸��¼ A Where a.��¼״̬ = 0 And a.ִ������ <= Sysdate " & _
               " Order By ҩƷid "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "ִ�е���")
    
    If rstemp.RecordCount = 0 Then Exit Sub
    
    For i = 0 To rstemp.RecordCount - 1
        gstrSQL = "Zl_ҩƷ�շ���¼_Adjust(" & rstemp!ҩƷID & ")"
        zldatabase.ExecuteProcedure gstrSQL, "ִ�е���"
        rstemp.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function CheckPriceAdjustByNO(ByVal Int���� As Integer, ByVal lng�ⷿid As Long, ByVal strNo As String, Optional ByVal lng���ⷿid As Long = 0) As Boolean
    '�����ݺż������
    Dim rsData As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '���û����ȫ�ֵ����۹����򲻽��к�����飬����true
    If Val(zldatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjustByNO = True: Exit Function
    
    gstrSQL = "Select ҩƷid, Nvl(����, 0) As ���� From ҩƷ�շ���¼ " & _
        " Where ���� = [1] And �ⷿid = [2] And NO = [3] And ������� Is Null"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjustByNO", Int����, lng�ⷿid, strNo)
    
    If rsData.EOF Then CheckPriceAdjustByNO = True: Exit Function
    
    Do While Not rsData.EOF
        If IsPriceAdjustMod(Val(rsData!ҩƷID)) = True Then
            If CheckPriceAdjust(rsData!ҩƷID, IIf(lng���ⷿid = 0, lng�ⷿid, lng���ⷿid), rsData!����) = False Then
                gstrSQL = "Select '[' || ���� || ']' || ���� || '(' || ��� || ')' as ҩƷ From �շ���ĿĿ¼ Where ID = [1]"
                Set rsItem = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjustByNO", Val(rsData!ҩƷID))
                If Not rsItem.EOF Then
                    MsgBox "����[" & strNo & "]�е�ҩƷ " & rsItem!ҩƷ & "�����������۹�������ҩ�����ۼۺͳɱ��۲�һ�£����ܷ�ҩ ��" & _
                     vbCrLf & "���Ƚ��е����ٷ�ҩ��", vbInformation, gstrSysName
                Else
                    MsgBox "����[" & strNo & "]�е�ҩƷ�����������۹�������ҩ�����ۼۺͳɱ��۲�һ�£����ܷ�ҩ ��" & _
                     vbCrLf & "���Ƚ��е����ٷ�ҩ��", vbInformation, gstrSysName
                End If
                
                Exit Function
            End If
        End If
        
        rsData.MoveNext
    Loop
    
    CheckPriceAdjustByNO = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPriceAdjust(ByVal lngҩƷid As Long, ByVal lng�ⷿid As Long, ByVal lng���� As Long) As Boolean
    '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
    '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    '�޿��ʱ���ɱ���ȡҩƷ���ĳɱ���
    '������lngҩƷid-ҩƷ���ID��Ϊ0��������ҩƷ��lng�ⷿid-��Ӧ�ĿⷿID��Ϊ0�������пⷿ��lng����-��Ӧ�����Σ��������-1�򲻹�������
    '���أ�True-������false-�в��������۹���Ҫ���ҩƷ
    '
    Dim rsData As ADODB.Recordset
    Dim str���� As String
    
    On Error GoTo errHandle
    
    '���û����ȫ�ֵ����۹����򲻽��к�����飬����true
    If Val(zldatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjust = True: Exit Function
    
    '������޿��
    If lngҩƷid > 0 Then
        If lng�ⷿid > 0 Then
            gstrSQL = "Select 1 from ҩƷ��� Where ����=1 and ҩƷid=[1] and �ⷿid=[2] " & _
                " And Not (nvl(����,0) = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0)"
            
            If lng���� > 0 Then
                gstrSQL = gstrSQL & " and Nvl(����,0)=[3] "
            End If
        Else
            gstrSQL = "Select 1 from ҩƷ��� Where ����=1 and ҩƷid=[1] " & _
                " And Not (nvl(����,0) = 0 And �������� < 0 And ʵ������ = 0 And ʵ�ʽ�� = 0 And ʵ�ʲ�� = 0)"
        End If
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lngҩƷid, lng�ⷿid, lng����)
        
        If rsData.EOF Then
            '�޿��ʱ�����շѼ�Ŀȡ�ۼۣ���ҩƷ���ȡ�ɱ��ۣ����Ƚϼ۸�
            gstrSQL = "Select a.�ɱ���, b.�ּ� As �ۼ� " & _
                " From ҩƷ��� A, �շѼ�Ŀ B " & _
                " Where a.ҩƷid = b.�շ�ϸĿid And (Sysdate Between b.ִ������ And b.��ֹ����) And Nvl(a.�Ƿ����۹���, 0) = 1 " & _
                " And b.�ּ� <> a.�ɱ��� And a.ҩƷid = [1] " & GetPriceClassString("B")
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lngҩƷid)
            
            If rsData.EOF Then
                'û�ҵ���ʾ�۸�һ��
                CheckPriceAdjust = True
            Else
                '�ҵ���ʾ�۸�һ��
                CheckPriceAdjust = False
            End If
            
            Exit Function
        End If
    End If
    
    If lngҩƷid > 0 Then
        str���� = IIf(str���� = "", "", str����) & " and a.ҩƷid=[1] "
    End If
    
    If lng�ⷿid > 0 Then
        str���� = IIf(str���� = "", "", str����) & " and d.�ⷿid=[2] "
    End If
    
    If lng���� >= 0 Then
        str���� = IIf(str���� = "", "", str����) & " and nvl(d.����,0)=[3] "
    End If
    
    gstrSQL = "Select a.ҩƷid, '['|| c.���� || ']'|| c.����||decode(c.����,null,null,'('||c.����||')') ||c.��� As ͨ���� " & vbNewLine & _
        "       From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D" & vbNewLine & _
        "       Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And" & vbNewLine & _
        "             (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0 And Nvl(a.�Ƿ����۹���, 0) = 1 And" & vbNewLine & _
        "             b.�ּ� <> nvl(d.ƽ���ɱ���,a.�ɱ���) " & str���� & GetPriceClassString("B") & vbNewLine & _
        "  And Not (nvl(D.����,0) = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0) " & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select a.ҩƷid, '['|| c.���� || ']'|| c.����||decode(c.����,null,null,'('||c.����||')') ||c.��� As ͨ���� " & vbNewLine & _
        " From ҩƷ��� A, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E" & vbNewLine & _
        " Where a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And c.�Ƿ��� = 1 And" & vbNewLine & _
        "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.�Ƿ����۹���, 0) = 1 And nvl(d.���ۼ�,0) <> nvl(d.ƽ���ɱ���,a.�ɱ���) " & str���� & _
        "  And Not (nvl(D.����,0) = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjust", lngҩƷid, lng�ⷿid, lng����)
    
    'û�ҵ����������۹���Ҫ��ļ�¼������true
    If rsData.EOF Then CheckPriceAdjust = True: Exit Function
    
    CheckPriceAdjust = False
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPriceAdjustBatch(ByVal lng�ⷿid As Long, ByVal strҩƷ���� As String) As String
    '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£���������ѯģʽ
    '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
    '�޿��ʱ���ɱ���ȡҩƷ���ĳɱ���
    '������lng�ⷿid-��Ӧ�ĿⷿID��strҩƷ����-��Ӧ��ҩƷID�����Σ���ʽΪ��ҩƷID,����|ҩƷID,����...��
    '���أ��մ�-�������ǿմ�-�в��������۹���Ҫ���ҩƷ����
    '
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    Dim strInfo As String
    
    On Error GoTo errHandle
    
    '���û����ȫ�ֵ����۹����򲻽��к�����飬����true
    If Val(zldatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjustBatch = "": Exit Function

    gstrSQL = "Select Distinct a.ҩƷid, '[' || c.���� || ']' || c.���� || Decode(c.����, Null, Null, '(' || c.���� || ')') || c.��� As ͨ����" & vbNewLine & _
        "From ҩƷ��� A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ҩƷ��� D, Table(f_Num2list2([2], '|', ',')) T" & vbNewLine & _
        "Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.���� = 1 And (Sysdate Between b.ִ������ And b.��ֹ����) And" & vbNewLine & _
        "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.�Ƿ��� = 0 And Nvl(a.�Ƿ����۹���, 0) = 1 And" & vbNewLine & _
        "      b.�ּ� <> Nvl(d.ƽ���ɱ���, a.�ɱ���) And a.ҩƷid = t.C1 And d.�ⷿid = [1] And Nvl(d.����, 0) = t.C2 And b.�۸�ȼ� Is Null" & vbNewLine & _
        " And Not (nvl(D.����,0) = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0) " & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select Distinct a.ҩƷid, '[' || c.���� || ']' || c.���� || Decode(c.����, Null, Null, '(' || c.���� || ')') || c.��� As ͨ����" & vbNewLine & _
        "From ҩƷ��� A, �շ���ĿĿ¼ C, ҩƷ��� D, ���ű� E, Table(f_Num2list2([2], '|', ',')) T" & vbNewLine & _
        "Where a.ҩƷid = c.Id And a.ҩƷid = d.ҩƷid And d.�ⷿid = e.Id And d.���� = 1 And c.�Ƿ��� = 1 And" & vbNewLine & _
        "      (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(a.�Ƿ����۹���, 0) = 1 And" & vbNewLine & _
        "      Nvl(d.���ۼ�, 0) <> Nvl(d.ƽ���ɱ���, a.�ɱ���) And a.ҩƷid = t.C1 And d.�ⷿid = [1] And Nvl(d.����, 0) = t.C2 " & vbNewLine & _
        " And Not (nvl(D.����,0) = 0 And D.�������� < 0 And D.ʵ������ = 0 And D.ʵ�ʽ�� = 0 And D.ʵ�ʲ�� = 0) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjustBatch", lng�ⷿid, strҩƷ����)

    'û�ҵ����������۹���Ҫ��ļ�¼������true
    If rsData.EOF Then CheckPriceAdjustBatch = "": Exit Function
    
    For i = 1 To rsData.RecordCount
        If i > 3 Then
            Exit For
        End If
         
        strInfo = IIf(strInfo = "", "", strInfo & vbCrLf) & rsData!ͨ����
        
        rsData.MoveNext
    Next

    CheckPriceAdjustBatch = strInfo
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsPriceAdjustMod(ByVal lngҩƷid As Long) As Boolean
    '�ж�ҩƷ�Ƿ��������۹���
    Dim rsData As ADODB.Recordset
    
    gstrSQL = "Select Nvl(�Ƿ����۹���, 0) As �Ƿ����۹��� From ҩƷ��� Where ҩƷid = [1] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "IsPriceAdjustMod", lngҩƷid)
    
    If rsData.EOF Then IsPriceAdjustMod = False: Exit Function
    
    IsPriceAdjustMod = (rsData!�Ƿ����۹��� = 1)
End Function
Public Function GetMediPackerDetail(ByVal lng�շ�ID As Long, ByVal str���� As String, ByVal str���� As String) As String
    '����ҩƷ�ְ����ӿ�
    '�����շ�ID������Ҫ����ְ����ӿڵ���ϸ�ַ���
    '���ص��ַ�����һ��˳��͸�ʽ
    
    Dim rsData As ADODB.Recordset
    Dim rsGetNext As ADODB.Recordset
    Dim n As Integer
    Dim strReturn As String
    Dim strLastTime As String
    Dim intCount As Integer
    Dim blnErr As Boolean
    
     gstrSQL = "Select /*+ Rule */ A.�շ�id, A.סԺ��, A.����id, A.����, A.��������, A.��������, A.������, A.����, A.�÷�, A.ҩƷ����, A.ҩƷ����, A.���, A.����ϵ��, A.������λ, A.��������,A.�����," & _
        " A.�״�ʱ��, A.ĩ��ʱ��,A.��ʼִ��ʱ��, A.Ƶ�ʼ��, A.�����λ, A.ִ��ʱ�䷽��, Nvl(B.��������, 0) As ����, A.��ҩ����,����װ " & _
        " From (Select Distinct A.ID As �շ�id, B.��ʶ�� As סԺ��, B.����id, B.����, C.���� As ��������, C.���� As ��������, B.������, B.����, A.�÷�,A.�����," & _
        " D.���� As ҩƷ����, D.���� As ҩƷ����, D.���, E.����ϵ��, F.���㵥λ As ������λ, H.�������� / E.����ϵ�� As ��������, G.�״�ʱ��, G.ĩ��ʱ��," & _
        " H.��ʼִ��ʱ�� , H.Ƶ�ʼ��, H.�����λ, H.ִ��ʱ�䷽��, H.���id, G.���ͺ�, A.ʵ������ * Nvl(A.����, 1) / E.סԺ��װ As ��ҩ����,decode(mod(A.ʵ������ * Nvl(A.����, 1) , E.ҩ���װ),0,1,0) ����װ " & _
        " From ҩƷ�շ���¼ A, סԺ���ü�¼ B, ���ű� C, �շ���ĿĿ¼ D, ҩƷ��� E, ������ĿĿ¼ F, ����ҽ������ G, ����ҽ����¼ H "
    If str���� <> "����" Then
        gstrSQL = gstrSQL & " ,ҩƷ���� I, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) J "
    End If
    gstrSQL = gstrSQL & " Where A.����id = B.ID And B.���˲���id = C.ID And A.ҩƷid = D.ID And A.ҩƷid = E.ҩƷid And E.ҩ��id = F.ID And " & _
        " B.ҽ����� = G.ҽ��id And B.NO = G.NO And B.ҽ����� = H.ID And A.ID = [1] "
    If str���� <> "����" Then
        gstrSQL = gstrSQL & "And E.ҩ��id = I.ҩ��id And I.ҩƷ���� = J.Column_Value "
    End If
    gstrSQL = gstrSQL & ") A, ����ҽ������ B " & _
        " Where A.���id = B.ҽ��id(+) And A.���ͺ� = B.���ͺ�(+) "

    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ҽ����ϸ", lng�շ�ID, str����)
    
    If rsData.RecordCount = 0 Then Exit Function
    
    With rsData
        If Not .EOF Then
            '�������������������װ�������򲻷��͵���ҩ��
            If !����װ = 0 Or str���� = "����" Then
                If Val(Nvl(!Ƶ�ʼ��, 0)) = 0 Or Nvl(!�����λ, "") = "" Or Nvl(!ִ��ʱ�䷽��, "") = "" Or str���� = "����" Then
                    intCount = 1
                Else
                    intCount = Val(!����)
                    If intCount = 0 Then
                        gstrSQL = "Select Zl_Gettransexenumber([1],[2],[3],[4],[5],[6]) From Dual "
                        Set rsGetNext = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�´�ִ��ʱ��", CDate(!��ʼִ��ʱ��), CDate(!�״�ʱ��), CDate(!ĩ��ʱ��), Val(!Ƶ�ʼ��), !�����λ, !ִ��ʱ�䷽��)
                        If Not rsGetNext.EOF Then
                            intCount = Val(rsGetNext.Fields(0).Value)
                        End If
                    End If
                    If intCount = 0 Then
                        intCount = 1
                        blnErr = True
                    End If
                End If
                
                For n = 1 To intCount
                    strReturn = IIf(strReturn = "", "", strReturn & "|")
                    strReturn = strReturn & !�շ�ID
                    strReturn = strReturn & ";" & !סԺ��
                    strReturn = strReturn & ";" & !����ID
                    strReturn = strReturn & ";" & Replace(Replace(!����, ";", ""), "|", "")
                    strReturn = strReturn & ";" & !��������
                    strReturn = strReturn & ";" & Replace(Replace(!��������, ";", ""), "|", "")
                    strReturn = strReturn & ";" & Replace(Replace(!������, ";", ""), "|", "")
                    strReturn = strReturn & ";" & Replace(Replace(Nvl(!����, ""), ";", ""), "|", "")
                    strReturn = strReturn & ";" & Replace(Replace(Nvl(!�÷�, ""), ";", ""), "|", "")
                    strReturn = strReturn & ";" & ""    '����ʱ��˵��
                    strReturn = strReturn & ";" & !ҩƷ����
                    strReturn = strReturn & ";" & Replace(Replace(!ҩƷ����, ";", ""), "|", "")
                    strReturn = strReturn & ";" & Replace(Replace(!���, ";", ""), "|", "")
                    strReturn = strReturn & ";" & !����ϵ��
                    strReturn = strReturn & ";" & !������λ
                    
                    If str���� = "����" Then
                        strReturn = strReturn & ";" & !��ҩ����
                    Else
                        strReturn = strReturn & ";" & IIf(blnErr = False, !��������, !��ҩ����)
                    End If
                    
                    If n = 1 Then
                        strLastTime = Format(!�״�ʱ��, "YYYY-MM-DD HH:MM:SS")
                    Else
                        gstrSQL = "Select Zl_Gettransexetime([1],[2],[3],[4],[5]) From Dual "
                        Set rsGetNext = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�´�ִ��ʱ��", CDate(!��ʼִ��ʱ��), CDate(strLastTime), Val(!Ƶ�ʼ��), !�����λ, !ִ��ʱ�䷽��)
                        If Not rsGetNext.EOF Then
                            strLastTime = Format(rsGetNext.Fields(0).Value, "YYYY-MM-DD HH:MM:SS")
                        End If
                    End If
                    
                    strReturn = strReturn & ";" & strLastTime
                    strReturn = strReturn & ";" & "1"           '�ְ��豸���
                    strReturn = strReturn & ";" & "0"           '���ȱ��
                    
                    If str���� = "����" Then
                        strReturn = strReturn & ";" & "1"
                    Else
                        strReturn = strReturn & ";" & "0"
                    End If
                    
                    strReturn = strReturn & ";" & !�����
                Next
            End If
        End If
    End With
    
    GetMediPackerDetail = strReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub OutPutData(ByVal strOutput As String)
    '���ڱ������û��������ԣ����������Ի򲻷����������ʱʹ��
    '������ִ�еĹؼ����̣�����������ⲿ��־�ļ����Դ˷����������
    'ע�⣺����Ҫ����ʱ�ֹ�����ָ������־�ļ�������󻷾�ʱ�ŵ�����̨��������Ŀ¼��Դ���뻷��ʱ�ŵ������ļ�����Ŀ¼
    'ע�⣺�������Ҫ������Ҫ��ʱɾ����־�ļ���������־�ļ����ܻ��������ر����û������������������Ͽ�
    '��ϵͳ����ָ����ͬ����־�ļ���
    '��־�����Զ��壬�ο���ʽ��ʱ��+�����ڲ�����/����+ҵ������/����+�ؼ�����
    'Ĭ�ϵĴ�������ʱ�䣬�������Ҫ����ȥ��
    Dim objFile As New FileSystemObject
    Dim objTarget As TextStream
    Const STR_CONS_FILENAME As String = "zlDrugStore.log"
    
    err = 0
    
    On Error Resume Next
    
    '����ļ��Ƿ����
    Set objTarget = objFile.OpenTextFile(App.Path & "\" & STR_CONS_FILENAME)
    
    '������������������
    If objTarget Is Nothing Then Exit Sub
    
'    If err <> 0 Then
'        '����Ŀ���ļ�
'        Set objFile = CreateObject("Scripting.FileSystemObject")
'        Set objTarget = objFile.CreateTextFile(App.Path & "\" & STR_CONS_FILENAME, True)
'        objTarget.Close
'    End If
    
    err.Clear
    On Error GoTo ErrHand
    
    Open App.Path & "\" & STR_CONS_FILENAME For Append Shared As #1
    
    Print #1, Now & vbCrLf & strOutput
    Close #1
    
    Exit Sub
ErrHand:
    Close #1
'    MsgBox err.Description, vbExclamation + vbOKOnly
End Sub

Public Sub AutoAdjustPrice_ByID(ByVal lngDrugID As Long)
    '��������ѵ�ִ�����ڶ��۸�δִ�е�ҩƷ��ִ�е��۹���
    '��ָ��ҩƷID���
    '��ҩƷѡ�����е���
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    
    On Error GoTo errHandle

    gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & lngDrugID & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice_ByID")

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckNotVerifyClosingAccount() As ADODB.Recordset
    '��ѯ��ǰ����Ա�����Ĳ����Ƿ����δ��˵Ľ���¼
    Dim rsData As ADODB.Recordset
    Dim strDept As String
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.Id, b.����, 'δ������' As ����" & vbNewLine & _
            "From ������Ա A, ���ű� B, ��������˵�� C, ҩƷ����¼ D, ҩƷ������ E" & vbNewLine & _
            "Where a.����id = b.Id And b.Id = c.����id And b.Id = d.�ⷿid And d.Id = e.���id And a.��Աid = [1] And" & vbNewLine & _
            "      c.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���') And d.������� Is Null" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select Distinct b.Id, b.����, 'δ��˽��' As ����" & vbNewLine & _
            "From ������Ա A, ���ű� B, ��������˵�� C" & vbNewLine & _
            "Where a.����id = b.Id And b.Id = c.����id And a.��Աid = [1] And c.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���') And" & vbNewLine & _
            "      Exists (Select 1 From ҩƷ����¼ D Where b.Id = d.�ⷿid And d.������� Is Null)"

    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "����ѯ", UserInfo.�û�ID)
    
    Set CheckNotVerifyClosingAccount = rsData
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AutoAdjustPrice_ByNO(ByVal Int���� As Integer, ByVal strNo As String)
    '��������ѵ�ִ�����ڶ��۸�δִ�е�ҩƷ��ִ�е��۹���
    '��ָ������,NO�е�ҩƷ�Ž��м��
    '����ͨҵ��ģ������ʱ����
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    Dim blnMore As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.ҩƷid " & _
        " From �շѼ�Ŀ A, ҩƷ�շ���¼ B, �շ���ĿĿ¼ C " & _
        " Where a.�շ�ϸĿid = b.ҩƷid And a.�շ�ϸĿid = c.Id And Nvl(c.�Ƿ���, 0) = 0 And a.�䶯ԭ�� = 0 And a.ִ������ <= Sysdate And b.������� Is Null " & _
        " And b.���� = [1] And b.No = [2]" & GetPriceClassString("A") & _
        " Union " & _
        " Select Distinct a.ҩƷid " & _
        " From ҩƷ�۸��¼ A, ҩƷ�շ���¼ B " & _
        " Where a.ҩƷid = b.ҩƷid And a.��¼״̬ = 0 And a.ִ������ <= Sysdate And b.������� Is Null And " & _
        " b.���� = [1] And b.No = [2] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice", Int����, strNo)

    With rsData
        If .RecordCount > 20 Then
            blnMore = True
            Call FS.ShowFlash("��������ִ�е��ۣ����Ժ�......")
        End If
        
        Do While Not .EOF
            lngAdjustID = !ҩƷID
            gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & lngAdjustID & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice")
            
            .MoveNext
        Loop
        
        If blnMore = True Then
            Call FS.StopFlash
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub AutoAdjustPrice_Batch()
    '��������ѵ�ִ�����ڶ��۸�δִ�е�ҩƷ��ִ�е��۹���
    '�������ҩƷ
    '��ҩƷѡ�������ݼ���ʼ��ʱ����
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    Dim blnMore As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct a.�շ�ϸĿid As ҩƷid" & vbNewLine & _
        "From �շѼ�Ŀ A, �շ���ĿĿ¼ B" & vbNewLine & _
        "Where a.�շ�ϸĿid = b.Id And b.��� In ('5', '6', '7') And Nvl(b.�Ƿ���, 0) = 0 And a.�䶯ԭ�� = 0 " & _
        "And a.ִ������ <= Sysdate" & GetPriceClassString("A") & vbNewLine & _
        "Union" & vbNewLine & _
        "Select Distinct a.ҩƷid From ҩƷ�۸��¼ A Where a.��¼״̬ = 0 And a.ִ������ <= Sysdate"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice")

    With rsData
        If .RecordCount > 20 Then
            blnMore = True
            Call FS.ShowFlash("��������ִ�е��ۣ����Ժ�......")
        End If
        
        Do While Not .EOF
            lngAdjustID = !ҩƷID
            gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & lngAdjustID & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice")
            
            .MoveNext
        Loop
        
        If blnMore = True Then
            Call FS.StopFlash
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Get�ɱ���(ByVal lngҩƷid As Long, ByVal lng�ⷿid As Long, ByVal lng���� As Long) As Double
'���ܣ���ȡ��ǰҩƷ�ĳɱ��۸�
'������ҩƷid,�ⷿid,����
'����ֵ�� �ɱ��۸�
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo errHandle
    
    gstrSQL = "select ƽ���ɱ��� from ҩƷ��� where ����=1 and ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷid, lng�ⷿid, lng����)
    
    If rsData.EOF Then
        blnNullPrice = True
    ElseIf IsNull(rsData!ƽ���ɱ���) = True Then
        blnNullPrice = True
    ElseIf Val(rsData!ƽ���ɱ���) < 0 Then
        blnNullPrice = True
    End If
    
    If Not blnNullPrice Then
        Get�ɱ��� = rsData!ƽ���ɱ���
    Else
        '����޷��ӿ����ȡ�ɱ��ۣ����ҩƷ�����ȡ
        gstrSQL = "select �ɱ��� from ҩƷ��� where ҩƷid=[1]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷid)
        If Not rsData.EOF Then
            If Val(Nvl(rsData!�ɱ���, 0)) > 0 Then
                Get�ɱ��� = rsData!�ɱ���
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�ۼ�(ByVal bln�Ƿ�ʱ�� As Boolean, lngҩƷid As Long, ByVal lng�ⷿid As Long, ByVal lng���� As Long) As Double
    '���ܣ���ȡԭʼ���ۼ۵�λ�ۼۣ���Ҫ���ڳ���
    '����: bln�Ƿ�ʱ��:false-����,true-ʱ��
    '����ֵ����С��λ�ļ۸�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo errHandle

    'ȡ����ҩƷ�ۼ�
    If bln�Ƿ�ʱ�� = False Then
        gstrSQL = "Select �ּ� " & _
            " From �շѼ�Ŀ A, ҩƷ��� B " & _
            " Where A.�շ�ϸĿid = B.ҩƷid And A.�շ�ϸĿID=[1] And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) " & GetPriceClassString("A")
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "Get�ۼ�-ȡ����ҩƷ�ۼ�", lngҩƷid)
        
        If Not rsData.EOF Then
            Get�ۼ� = rsData!�ּ�
        End If
    Else
        'ȡʱ��ҩƷ�ۼ�
        gstrSQL = "select Decode(Nvl(���ۼ�, 0), 0, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� " & _
            " from ҩƷ��� where ����=1 and  ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-���ۼ�", lngҩƷid, lng�ⷿid, lng����)
        
        If rsData.EOF Then
            'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
            '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
            '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
            gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷid)
            dblָ�����ۼ� = rsData!ָ�����ۼ�
            dbl��������� = rsData!���������
            
            Get�ۼ� = 0
            dbl�ɱ��� = Get�ɱ���(lngҩƷid, lng�ⷿid, lng����)
            dbl�ӳ��� = rsData!�ӳ��� / 100
            dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
            dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
            Get�ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
        Else
            If rsData!���ۼ� = 0 Then
                gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
                Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷid)
                dblָ�����ۼ� = rsData!ָ�����ۼ�
                dbl��������� = rsData!���������
                
                Get�ۼ� = 0
                dbl�ɱ��� = Get�ɱ���(lngҩƷid, lng�ⷿid, lng����)
                dbl�ӳ��� = rsData!�ӳ��� / 100
                dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                Get�ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
            Else
                Get�ۼ� = rsData!���ۼ�
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Getʱ�����ۼ�(ByVal lngҩƷid As Long, ByVal lng�ⷿid As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double) As Double
    '���ܣ���ȡʱ��ҩƷ��ǰҩƷ�����ۼ�
    '����:ҩƷid,�ⷿid,����
    '����ֵ�����ۼ�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo errHandle
    gstrSQL = "select Decode(Nvl(���ۼ�, 0), 0, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� from ҩƷ��� where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷid, lng�ⷿid, lng����)
    
    If rsData.EOF Then
        'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
        '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
        '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
        gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷid)
        dblָ�����ۼ� = rsData!ָ�����ۼ�
        dbl��������� = rsData!���������
        
        Getʱ�����ۼ� = 0
        dbl�ɱ��� = Get�ɱ���(lngҩƷid, lng�ⷿid, lng����)
        dbl�ӳ��� = rsData!�ӳ��� / 100
        dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
        dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
        Getʱ�����ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
    Else
        If rsData!���ۼ� = 0 Then
            gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷid)
            dblָ�����ۼ� = rsData!ָ�����ۼ�
            dbl��������� = rsData!���������
            
            Getʱ�����ۼ� = 0
            dbl�ɱ��� = Get�ɱ���(lngҩƷid, lng�ⷿid, lng����)
            dbl�ӳ��� = rsData!�ӳ��� / 100
            dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
            dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
            Getʱ�����ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
        Else
            Getʱ�����ۼ� = rsData!���ۼ� * dbl����ϵ��
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get���ۼ�(ByVal bln�Ƿ�ʱ�� As Boolean, lngҩƷid As Long, ByVal lng�ⷿid As Long, ByVal lng���� As Long) As Double
    '���ܣ���ȡԭʼ���ۼ۵�λ�ۼۣ���Ҫ���ڳ���
    '����: bln�Ƿ�ʱ��:false-����,true-ʱ��
    '����ֵ����С��λ�ļ۸�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo errHandle

    'ȡ����ҩƷ�ۼ�
    If bln�Ƿ�ʱ�� = False Then
        gstrSQL = "Select �ּ� " & _
            " From �շѼ�Ŀ A, ҩƷ��� B " & _
            " Where A.�շ�ϸĿid = B.ҩƷid And A.�շ�ϸĿID=[1] And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) " & GetPriceClassString("A")
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "Get���ۼ�-ȡ����ҩƷ�ۼ�", lngҩƷid)
        
        If Not rsData.EOF Then
            Get���ۼ� = rsData!�ּ�
        End If
    Else
        'ȡʱ��ҩƷ�ۼ�
        gstrSQL = "select Decode(Nvl(���ۼ�, 0), 0, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� " & _
            " from ҩƷ��� where ����=1 and  ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-���ۼ�", lngҩƷid, lng�ⷿid, lng����)
        
        If rsData.EOF Then
            'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
            '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
            '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
            gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷid)
            dblָ�����ۼ� = rsData!ָ�����ۼ�
            dbl��������� = rsData!���������
            
            Get���ۼ� = 0
            dbl�ɱ��� = Get�ɱ���(lngҩƷid, lng�ⷿid, lng����)
            dbl�ӳ��� = rsData!�ӳ��� / 100
            dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
            dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
            Get���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
        Else
            If rsData!���ۼ� = 0 Then
                gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
                Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷid)
                dblָ�����ۼ� = rsData!ָ�����ۼ�
                dbl��������� = rsData!���������
                
                Get���ۼ� = 0
                dbl�ɱ��� = Get�ɱ���(lngҩƷid, lng�ⷿid, lng����)
                dbl�ӳ��� = rsData!�ӳ��� / 100
                dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                Get���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
            Else
                Get���ۼ� = rsData!���ۼ�
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function CreateObject_LED(lngLEDModal As Long) As Boolean
    '����LED��ʾ����
    
    Dim strsql As String
    Dim strObject As String

    On Error GoTo ErrHand
    
    '��ȡ��LED��ʾ�ӿڵ�ע����Ϣ
    If grsLEDComponent.State = 0 Then
        strsql = "Select ��������,������,Nvl(����,0) AS ���� From �Ŷ�LED��ʾ����  "
        Set grsLEDComponent = zldatabase.OpenSQLRecord(strsql, "��ȡ��LED��ʾ�ӿڵ�ע����Ϣ")
    End If
    grsLEDComponent.Filter = "��������=" & lngLEDModal
    If grsLEDComponent.RecordCount = 0 Then
        grsLEDComponent.Filter = 0
        MsgBox "��LED�ӿڻ�δע�ᣡ ���=" & lngLEDModal, vbInformation, "�Ŷӽк�ϵͳ"
        Exit Function
    End If
    strObject = UCase(grsLEDComponent!������)
    grsLEDComponent.Filter = 0
    
    '���ö����Ƿ����
    On Error Resume Next
    If Not gobjLEDShow Is Nothing Then
        CreateObject_LED = True
        Exit Function
    End If
    
    'ȥ���ļ�����׺
    strObject = Mid(strObject, 1, Len(strObject) - 4)
    '��������
    Set gobjLEDShow = CreateObject(strObject & ".Cls" & Mid(strObject, 4))
    
    
    '���ó�ʼ������
    CreateObject_LED = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SetGridFocus(ByVal objGrid As VSFlexGrid, ByVal blnGetFoucs As Boolean)
    With objGrid
        If blnGetFoucs Then
            .GridColorFixed = &H80000008
            .GridColor = &H80000008
            .ForeColorFixed = glngFixedForeColorByFocus
            .BackColorSel = glngRowByFocus
        Else
            .GridColorFixed = &H80000011
            .GridColor = &H80000011
            .ForeColorFixed = glngFixedForeColorNotFocus
            .BackColorSel = glngRowByNotFocus
        End If
    End With
End Sub
Public Function CheckIsCenter(ByVal lngStockid As Long) As Boolean
    '����ҩ���Ƿ���С��������ġ�����
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ��������˵�� Where �������� = '��������' And ����id = [1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ����������������", lngStockid)
    
    If Not rsTmp.EOF Then CheckIsCenter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Get�ּ�(ByVal lngҩƷid As Long) As Double
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select �ּ� " & _
            " From �շѼ�Ŀ A, ҩƷ��� B " & _
            " Where A.�շ�ϸĿid = B.ҩƷid And A.�շ�ϸĿID=[1] And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) " & GetPriceClassString("A")
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "[��ȡ��ҩƷ�����۵�λ�۸�]", lngҩƷid)
    
    If Not rstemp.EOF Then
        Get�ּ� = rstemp!�ּ�
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetDefaultRecipeColor() As String
    GetDefaultRecipeColor = CStr(gconlng��ͨ) & ";" & _
                    CStr(gconlng����) & ";" & _
                    CStr(gconlng����) & ";" & _
                    CStr(gconlng����) & ";" & _
                    CStr(gconlng��һ) & ";" & _
                    CStr(gconlng����)

End Function
Public Sub DeptSendWork_CheckDrugstore(ByVal strPrivs As String, ByVal lngUserID As Long, ByVal strStateNo As String)
    '���ҩ�����÷�(��ҩ������ҩ������ҩ��)������ģ���ж�Ӧ���ڴ�ǰ���
    'strPrivs��Ȩ�ޣ�
    'lngUserID����ǰ�û�ID��
    'strStateNo����ǰϵͳվ���ţ�
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo errHandle
    If IsInString(strPrivs, "����ҩ��", ";") Then
        gstrSQL = "(Select Distinct ����ID From ��������˵�� Where �������� Like '%ҩ��' And ������� IN (2,3))"
    Else
        gstrSQL = "(Select distinct A.����ID From ������Ա A,��������˵�� B " & _
                 " Where A.��ԱID=[1] And A.����ID=B.����ID And B.�������� Like '%ҩ��' And B.������� IN (2,3))"
    End If
    gstrSQL = " Select Distinct P.ID,P.���� From ���ű� P " & _
             " Where (P.վ�� = '" & strStateNo & "' Or P.վ�� is Null) And P.ID In " & gstrSQL & _
             " And (P.����ʱ�� Is Null Or P.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ҩ������", lngUserID)
    
    With rsData
        If .EOF Then
           If IsInString(strPrivs, "����ҩ��", ";") Then
               strMsg = "���ʼ��ҩ�������Ź���"
           Else
               strMsg = "�㲻��ҩ��������Ա�����ܲ�����ģ�飡"
           End If
           MsgBox strMsg, vbInformation, gstrSysName
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function DeptSendWork_GetDrugstore(ByVal strPrivs As String, ByVal lngUserID As Long, ByVal strStateNo As String) As ADODB.Recordset
    'ȡ��Ӧ����Ա���������ҩ��
    'strPrivs��Ȩ�ޣ�
    'lngUserID����ǰ�û�ID��
    'strStateNo����ǰϵͳվ���ţ�
    
    On Error GoTo errHandle
    If IsInString(strPrivs, "����ҩ��", ";") Then
        gstrSQL = "(Select Distinct ����ID From ��������˵�� Where �������� Like '%ҩ��' And ������� IN (2,3))"
    Else
        gstrSQL = "(Select distinct A.����ID From ������Ա A,��������˵�� B " & _
                 " Where A.��ԱID=[1] And A.����ID=B.����ID And B.�������� Like '%ҩ��' And B.������� IN (2,3))"
    End If
    gstrSQL = " Select Distinct P.ID,P.���� From ���ű� P " & _
             " Where (P.վ�� = '" & strStateNo & "' Or P.վ�� is Null) And P.ID In " & gstrSQL & _
             " And (P.����ʱ�� Is Null Or P.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set DeptSendWork_GetDrugstore = zldatabase.OpenSQLRecord(gstrSQL, "ȡ����ҩ��ҩ��", lngUserID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get��ҩ;��() As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select ���� as �÷� ,�걾��λ As ���� From ������ĿĿ¼ Where ���='E' And ��������='2'And (�������=2 Or �������=3) " & _
            " And (����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or ����ʱ�� Is Null) Order by ���� "
    Set DeptSendWork_Get��ҩ;�� = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ;��")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get�Զ��巢ҩ����() As ADODB.Recordset
    On Error GoTo errHandle
    '��ȡ��ҩ����
    gstrSQL = "Select ���� From ��ҩ���� Order By ����"
    Set DeptSendWork_Get�Զ��巢ҩ���� = zldatabase.OpenSQLRecord(gstrSQL, "[��ȡ��ҩ����]")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub MediWork_CheckInOutType()
    '���ҩƷ������
    Dim rsData As ADODB.Recordset
    Dim int��ϵ�� As Integer, int��ϵ�� As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select A.����, A.���id, B.ID, B.����, B.����, B.ϵ�� " & _
        " From ҩƷ�������� A, ҩƷ������ B " & _
        " Where A.���id = B.Id " & _
        " Order By ����"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���������")
    
    With rsData
        If .EOF Then Exit Sub
        
        rsData.Filter = "����=1"
        gInOutType.bln�⹺��� = Not .EOF
        
        rsData.Filter = "����=2"
        gInOutType.bln������� = Not .EOF
        
        rsData.Filter = "����=3"
        gInOutType.blnЭҩ��� = Not .EOF
        
        rsData.Filter = "����=4"
        gInOutType.bln������� = Not .EOF
        
        rsData.Filter = "����=5"
        gInOutType.bln��۵��� = Not .EOF
        
        rsData.Filter = "����=6"
        gInOutType.blnҩƷ�ƿ� = Not .EOF
        
        rsData.Filter = "����=7"
        gInOutType.blnҩƷ���� = Not .EOF
        
        rsData.Filter = "����=8"
        gInOutType.bln�շѴ�����ҩ = Not .EOF
        
        rsData.Filter = "����=9"
        gInOutType.bln���ʵ�������ҩ = Not .EOF
        
        rsData.Filter = "����=10"
        gInOutType.bln���ʱ�����ҩ = Not .EOF
        
        rsData.Filter = "����=11"
        gInOutType.bln�������� = Not .EOF
        
        rsData.Filter = "����=27"
        gInOutType.blnҩƷ���� = Not .EOF
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function DeptSendWork_CheckBill(ByVal IntOper As Integer, ByVal lng�շ�ID As Long, ByVal bln����δ��˴�����ҩ As Boolean) As Integer
    '--���ݽ�Ҫִ�еĲ������ж��Ƿ�����--
    '0-�ܷ�;1-��ҩ;2-��ҩ
    '����:
    '0-�������
    '1-�ѷ�ҩ
    '2-��ɾ��
    '3-δ��ҩ
    
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select A.NO,Nvl(B.��¼״̬,0) AS ��˱�־,A.�����,Decode(Nvl(A.ժҪ,'С��'),'�ܷ�',3,B.ִ��״̬) ִ��״̬,A.��ҩ��ʽ From ҩƷ�շ���¼ A,סԺ���ü�¼ B " & _
             " Where A.����ID=B.ID And A.ID=[1] "
    If IntOper = 2 Then
        gstrSQL = gstrSQL & " And A.����� IS Not Null"
    Else
        gstrSQL = gstrSQL & " And A.����� IS Null"
    End If
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "��鵥��״̬", lng�շ�ID)
    
    With rsData
        If .EOF Then
            DeptSendWork_CheckBill = 2
            MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Not IsNull(!�����) Then
            If IntOper <> 2 Then
                DeptSendWork_CheckBill = 1
                MsgBox "�ô���[" & !NO & "]�ѱ���������Ա��ҩ������������ֹ��", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If IntOper = 2 Then
                DeptSendWork_CheckBill = 3
                MsgBox "�ô���[" & !NO & "]��δ��ҩ������������ֹ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If IntOper = 1 Then
            If !ִ��״̬ = 3 Then
                DeptSendWork_CheckBill = 2
                MsgBox "�ô���[" & !NO & "]�Ѿܷ�������������ֹ��", vbInformation, gstrSysName
                Exit Function
            End If
            
            If !��˱�־ = 0 And bln����δ��˴�����ҩ = False Then
                DeptSendWork_CheckBill = 4
                MsgBox "�ô���[" & !NO & "]��δ��ˣ�����������ֹ��", vbInformation, gstrSysName
                Exit Function
            End If
            
            If Nvl(!��ҩ��ʽ, 0) = -1 Then
                DeptSendWork_CheckBill = 5
                MsgBox "�ô���[" & !NO & "]��ֹͣ��ҩ������������ֹ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    DeptSendWork_CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MediWork_CheckStorageStock(ByVal lngStockid As Long, ByVal lngMediID As Long) As Boolean
    '���ָ��ҩƷ�Ƿ����ô洢�ⷿ
    'lngStockID���ⷿID
    'lngMediID��ҩƷID
    
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1] And �շ�ϸĿid = [2]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "���ҩƷ�洢�ⷿ", lngStockid, lngMediID)
    
    MediWork_CheckStorageStock = Not rsData.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get��ҩ��(ByVal lngҩ��id As Long) As ADODB.Recordset
    '��ȡҩ����Ա
    On Error GoTo errHandle
    gstrSQL = "Select Distinct A.����||'-'||A.���� As ����,A.���� ����" & _
             " From ��Ա�� A,������Ա B,��������˵�� C,��Ա����˵�� D " & _
             " Where A.Id=B.��Աid And B.����id=C.����Id And D.��Աid=A.Id And D.��Ա���� = 'ҩ����ҩ��' " & _
             " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) AND B.����id=[1] " & _
             " ORDER BY ���� "

    Set DeptSendWork_Get��ҩ�� = zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ����Ա", lngҩ��id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DeptSendWork_Get�˲���(ByVal lngҩ��id As Long) As ADODB.Recordset
    On Error GoTo errHandle
    '��ȡҩ����Ա
    gstrSQL = "Select ����||'-'||���� As ����,���� As ���� From ��Ա�� Where Id In (Select ��Աid from ������Ա Where ����id=[1]) " & _
             " And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) " & _
             " ORDER BY ���� "

    Set DeptSendWork_Get�˲��� = zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ����Ա", lngҩ��id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get��ҩ����ʽ(ByVal strRPTCode As String) As ADODB.Recordset
    '��ȡ�����ʽ����
    '������strRPTCode-�������
    On Error GoTo errHandle
    gstrSQL = "Select ˵�� As ��ʽ From zltools.zlRPTFMTs Where ����id = (Select ID From zltools.zlReports Where ϵͳ = [1] And ��� = [2]) Order By ���"
    Set DeptSendWork_Get��ҩ����ʽ = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ҩ����ʽ", glngSys, strRPTCode)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get����(ByVal lng�ⷿid As Long) As ADODB.Recordset
    '��ȡ���м���
    On Error GoTo errHandle
    gstrSQL = "Select Distinct J.����||'-'||J.���� ����" & _
         " From ����ִ�п��� A,ҩƷ���� B,ҩƷ���� J " & _
         " Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.����" & _
         " And A.ִ�п���ID=[1]" & _
         " Order By j.���� || '-' || j.���� "
    Set DeptSendWork_Get���� = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�ⷿҩƷ����", lng�ⷿid)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DeptSendWork_Get��ҩ;������() As ADODB.Recordset
    '��ȡ��ҩ;������
    On Error GoTo errHandle
    gstrSQL = "Select Distinct �걾��λ As ���� From ������ĿĿ¼ Where ��� = 'E' And �������� = '2' And �걾��λ Is Not Null"
    Set DeptSendWork_Get��ҩ;������ = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ҩ;������")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function IsInString(ByVal strTarget As String, ByVal strOrigin As String, Optional strSplit As String = "") As Boolean
    'ĳ���ַ����Ƿ������һ���ַ���
    'strTarget��Ŀ���ַ���
    'strOrigin��ԭ�ַ���
    'strSplit���ָ�������Ϊ��ʱΪ��ȷƥ�䣩
    '��strTarget���Ƿ����strOrigin
    
    IsInString = InStrB(strSplit & strTarget & strSplit, strSplit & strOrigin & strSplit) > 0
End Function

Public Function MediWork_GetCheckStockRule(ByVal lng�ⷿid As Long) As Integer
    'ȡ���������
    Dim rsData As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡ���������", lng�ⷿid)

    If Not rsData.EOF Then
        MediWork_GetCheckStockRule = rsData!�����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function MediWork_GetMediRealAmount(ByVal lng�ⷿid As Long, ByVal lngҩƷid As Long, ByVal lng���� As Long) As Double
    'ȡҩƷʵ�ʿ��
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(ʵ������, 0) As ʵ������ " & _
            " From ҩƷ��� " & _
            " Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And Nvl(����, 0) = [3] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ���ʵ������", lng�ⷿid, lngҩƷid, lng����)

    If Not rsData.EOF Then
        MediWork_GetMediRealAmount = rsData!ʵ������
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function RecipeSendWork_GetDiagnosis(ByVal int���� As Integer, ByVal LngID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal int��ʾ������� As Integer) As String
    'ȡ���������Ϣ
    '1.���ﲡ�ˣ�����ҽ��ID��ȡ��ϼ�¼
    '2��3סԺ���ˣ����ݲ���ID����ҳID��ȡ��ϼ�¼
    'int��ʾ�������:1-ֻ��ѯ��ҩ���;2-ֻ��ѯ��ҩ���
    Dim rsData As ADODB.Recordset
    Dim strTmp, strFilter As String
    Dim strReturn As String
    Dim str��¼����, str������� As String
    Dim int�������, n As Integer
    Dim str�����Ϣ, str������� As String
    
    '1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���,8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
    
    If LngID = 0 Then Exit Function
    On Error GoTo errHandle
    If int���� = 1 Then
        gstrSQL = "Select A.�������,A.�Ƿ����� From ������ϼ�¼ A, �������ҽ�� B Where A.ID = B.���id And B.ҽ��id = [1] And ȡ��ʱ�� Is Null "

        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "RecipeSendWork_GetDiagnosis", LngID)
        
        With rsData
            Do While Not .EOF
                If Nvl(!�������, "") <> "" Then
                    strReturn = IIf(strReturn = "", "", strReturn & "|") & !������� & IIf(Nvl(rsData!�Ƿ�����, 0) = 1, "������", "")
                End If
                
                .MoveNext
            Loop
        End With
    ElseIf int���� = 2 Then
        '����Ժ����Ժ��������ѡ˳�򷵻����
        '����ֵ��ʽ���������,�������;�������|�������,�������;�������...
        gstrSQL = "Select ��¼��Դ,�������,��ϴ���,�������,�Ƿ�����,Mod(�������,10) as ���� From ������ϼ�¼" & _
            " Where ����ID=[1] And ��ҳID=[2] And ������� IN(" & IIf(int��ʾ������� = 1, "11,12,13", IIf(int��ʾ������� = 2, "1,2,3", "1,2,3,11,12,13")) & ")" & _
            " Order by ��¼��Դ,�������,��ϴ���"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "RecipeSendWork_GetDiagnosis", LngID, lng��ҳID)
        
        '�Ȱ���Դ����˳�����
        rsData.Filter = "��¼��Դ=3" '��ҳ����
        If rsData.EOF Then rsData.Filter = "��¼��Դ=2" '��Ժ�Ǽ�
        If rsData.EOF Then rsData.Filter = "��¼��Դ=1" '����
        If rsData.EOF Then rsData.Filter = "��¼��Դ=4" '������¼��
        
'        'סԺ�ٰ���������˳�����
'        If Not rsData.EOF And int���� = 2 Then
'            gstrSQL = rsData.Filter
'            rsData.Filter = gstrSQL & " And ����=3"
'            If rsData.EOF Then rsData.Filter = gstrSQL & " And ����=2"
'            If rsData.EOF Then rsData.Filter = gstrSQL & " And ����=1"
'        End If
        
        'סԺ�ٰ���������˳�����
        strFilter = rsData.Filter
        For n = 3 To 1 Step -1
            str������� = ""
            rsData.Filter = strFilter & " And ����=" & n
            Do While Not rsData.EOF
                Select Case rsData!�������
                    Case 1
                        str������� = "��ҽ�������"
                    Case 2
                        str������� = "��ҽ��Ժ���"
                    Case 3
                        str������� = "��ҽ��Ժ���"
                    Case 11
                        str������� = "��ҽ�������"
                    Case 12
                        str������� = "��ҽ��Ժ���"
                    Case 13
                        str������� = "��ҽ��Ժ���"
                End Select
                
                If Not IsNull(rsData!�������) Then
                    str������� = IIf(str������� = "", "", str������� & ";") & rsData!������� & IIf(Nvl(rsData!�Ƿ�����, 0) = 1, "������", "")
                End If
                
                rsData.MoveNext
                
                If rsData.EOF Then str�����Ϣ = IIf(str�����Ϣ = "", "", str�����Ϣ & "|") & str������� & "," & str�������
            Loop
        Next
        
        strReturn = str�����Ϣ
    Else
        '��ȡ���
        '����ֵ��ʽ���������,�������;�������|�������,�������;�������...
        '����¼���ڵ����������������ͣ���¼ʱ��һ����ϲ��������
        gstrSQL = "Select �������, �������, ��¼����,�Ƿ����� " & _
            " From ������ϼ�¼ " & _
            " Where ����id = [1] And ��ҳid = [2] " & _
            " Order By ��¼���� Desc,�������  Desc"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "RecipeSendWork_GetDiagnosis", LngID, lng��ҳID)
        
        Do While Not rsData.EOF
            If str��¼���� & "," & int������� <> Format(rsData!��¼����, "YYYY-MM-DD HH:MM:SS") & "," & rsData!������� Then
                '������ͣ���¼�����뵱ǰ��¼����ͬʱ
                If str��¼���� <> "" Then
                    str�����Ϣ = IIf(str�����Ϣ = "", "", str�����Ϣ & "|") & str������� & "," & str�������
                End If
                
                Select Case rsData!�������
                    Case 1
                        str������� = "��ҽ�������"
                    Case 2
                        str������� = "��ҽ��Ժ���"
                    Case 3
                        str������� = "��ҽ��Ժ���"
                    Case 5
                        str������� = "Ժ�ڸ�Ⱦ"
                    Case 6
                        str������� = "�������"
                    Case 7
                        str������� = "�����ж���"
                    Case 8
                        str������� = "��ǰ���"
                    Case 9
                        str������� = "�������"
                    Case 10
                        str������� = "����֢"
                    Case 11
                        str������� = "��ҽ�������"
                    Case 12
                        str������� = "��ҽ��Ժ���"
                    Case 13
                        str������� = "��ҽ��Ժ���"
                    Case 21
                        str������� = "��ԭѧ���"
                    Case 22
                        str������� = "Ӱ��ѧ���"
                End Select
                
                str��¼���� = Format(rsData!��¼����, "YYYY-MM-DD HH:MM:SS")
                int������� = rsData!�������
                str������� = rsData!������� & IIf(Nvl(rsData!�Ƿ�����, 0) = 1, "������", "")
            Else
                '������ͣ���¼�����뵱ǰ��¼��ͬʱ��ϲ��������
                str������� = IIf(str������� = "", "", str������� & ";") & rsData!������� & IIf(Nvl(rsData!�Ƿ�����, 0) = 1, "������", "")
            End If
            
            rsData.MoveNext
            
            If rsData.EOF Then str�����Ϣ = IIf(str�����Ϣ = "", "", str�����Ϣ & "|") & str������� & "," & str�������
        Loop
        
        strReturn = str�����Ϣ
    End If
    
    RecipeSendWork_GetDiagnosis = strReturn
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetƤ�Խ��(ByVal lng����ID As Long, ByVal lngҩ��id As Long, ByVal dateCurrent As Date, ByVal date����ʱ�� As Date, Optional lng��ҳID As Long) As String
    'ȡ����Ƥ�Խ����ǰ���Ǵ�������ҩƷ��������Ҫ��Ƥ�Ե�ҩƷ
    '1�������ǰʱ���ڣ�Ƥ�Խ����Ч�������ã���Ƥ�Խ������ʹ�����Ƥ�Խ����ҩ��ʹ��ģ�����ʾΪ"����","����"����"����"��
    '2�������ǰʱ���ڣ�Ƥ�Խ����Ч�������ã���û��Ƥ�Խ�����͸���ҽ���Ŀ�ʼִ��ʱ������һ��Ƥ�Խ���Ǽ�ʱ��Ƚϣ������Ƥ�Խ����Ч���������ڣ���ʹ�����Ƥ�Խ����ҩ��ʹ��ģ�����ʾΪ��������ҩ����
    '3�����1��2������������ʾ����Ƥ�Խ����
    Dim rsData As ADODB.Recordset
    
    If lng����ID = 0 Then Exit Function
    
    On Error GoTo errHandle
    
'    gstrSQL = "Select ���,��¼ʱ�� From ���˹�����¼ Where ����id=[1] And ҩ��ID=[2] Order By ��¼ʱ�� Desc "
    
    gstrSQL = "Select Decode(���, 1, '(+)', '(-)') As ���, ��¼ʱ�� As ʱ��" & vbNewLine & _
        "From ���˹�����¼" & vbNewLine & _
        "Where ����id = [1] And ҩ��id = [2]" & vbNewLine & _
        IIf(lng��ҳID = 0, "", " And ��ҳid=[3] ") & _
        "Union All" & vbNewLine & _
        "Select '(��)' As ���, a.����ʱ�� As ʱ��" & vbNewLine & _
        "From ����ҽ����¼ A" & vbNewLine & _
        "Where a.����id = [1] And a.������� = 'E' And Ƥ�Խ��='����' And Exists" & vbNewLine & _
        " (Select 1" & vbNewLine & _
        "       From ������ĿĿ¼ B, �����÷����� C" & vbNewLine & _
        "       Where b.Id = c.�÷�id And b.��� = 'E' And b.�������� = '1' And b.Id = a.������Ŀid And c.��Ŀid = [2])" & vbNewLine & _
        IIf(lng��ҳID = 0, "", " And A.��ҳid=[3] ") & _
        "Order By ʱ�� Desc"

    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡƤ�Խ��", lng����ID, lngҩ��id, lng��ҳID)
    
    If rsData.RecordCount = 0 Then
        GetƤ�Խ�� = "<��>"
        Exit Function
    ElseIf DateDiff("D", rsData!ʱ��, dateCurrent) > gtype_UserSysParms.P70_�����Ǽ���Ч���� Then
        'Ƥ��ʱ����뵱ǰʱ�䳬����������
        If DateDiff("D", rsData!ʱ��, date����ʱ��) > gtype_UserSysParms.P70_�����Ǽ���Ч���� Then
            '����ʱ����������ǰ���е�Ƥ�Խ����Ч
            GetƤ�Խ�� = "<��>"
            Exit Function
        Else
            '����ʱ�����������ڽ��е�Ƥ����Ч
            GetƤ�Խ�� = rsData!��� & "<��>"
            Exit Function
        End If
    Else
        'Ƥ��ʱ����뵱ǰʱ��������������
        GetƤ�Խ�� = rsData!���
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function RecipeSendWork_Getҽ��() As ADODB.Recordset
    'ȡҽ��
    On Error GoTo errHandle
    gstrSQL = " Select Distinct A.����||'-'||A.���� ҽ�� From ��Ա�� A,��Ա����˵�� B" & _
             " Where B.��Ա����='ҽ��' And A.ID=B.��ԱID" & _
             " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) " & _
             " Order by ҽ��"
    Set RecipeSendWork_Getҽ�� = zldatabase.OpenSQLRecord(gstrSQL, "ȡҽ��")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function RecipeSendWork_JudgeSign(ByVal Int���� As Integer, ByVal strNo As String, Optional int�ɲ��� As Integer, Optional ByVal lng�շ�ID As Long, Optional ByVal dateʱ�� As Date) As Boolean
    Dim rsTmp As ADODB.Recordset
    
    '�жϴ����Ƿ��ѽ����˵���ǩ�����������ʾ���е���ǩ��
    On Error GoTo errHandle
    If lng�շ�ID = 0 Then
        gstrSQL = "Select 1 From ҩƷǩ����ϸ " & _
            " Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] )  And Rownum = 1 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "�жϴ����Ƿ��ѽ����˵���ǩ��", Int����, strNo)
    Else
        gstrSQL = "Select 1 From ҩƷǩ����ϸ " & _
            " Where �շ�id in (Select ID From ҩƷ�շ���¼ Where Id=[3] And ���� = [1] And NO = [2]) And  Rownum = 1 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "�жϴ����Ƿ��ѽ����˵���ǩ��", Int����, strNo, lng�շ�ID)
    End If
    RecipeSendWork_JudgeSign = (rsTmp.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function RecipeSendWork_DispensingMedi(ByVal lngҩ��id As Long, bln�Ƿ���ҩȷ�� As Boolean) As Boolean
    'ҩ���Ƿ���Ҫ��ҩ
    Dim rsData As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(��ҩ,0) AS ��ҩ,nvl(��ҩȷ��,0) as ��ҩȷ��,���� From ҩ����ҩ���� Where ҩ��ID=[1] Order by ����"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩ����ҩ����", lngҩ��id)
    
    'ֻҪ��һ���ʾ��Ҫ������ҩ���̵ģ����Ϊ��Ҫ��ҩ
    Do While Not rsData.EOF
        If rsData!��ҩ = 1 Then
            RecipeSendWork_DispensingMedi = True
        End If
        If rsData!���� = 1 Then
            If rsData!��ҩȷ�� = 1 Then
                bln�Ƿ���ҩȷ�� = True
            End If
        End If
        rsData.MoveNext
    Loop
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж��룩
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColData(intCol) = lngColWidth
        
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Public Function TvwCheckNode(ByVal Node As Object, blnCheck As Boolean, Optional ByVal blnAutoExpand As Boolean = False)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        If blnAutoExpand = True Then Node.Expanded = blnCheck
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If blnAutoExpand = True Then Node.Expanded = blnCheck
            If Node.Children > 0 Then
                TvwCheckNode Node, blnCheck, blnAutoExpand
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function
Public Sub TvwSetParentNode(ByVal tvwObj As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.index
            Do While intIdx <> Node.LastSibling.index
                If tvwObj.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvwObj.Nodes(intIdx).Next.index
            Loop
            If intIdx = Node.LastSibling.index Then
                If tvwObj.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            TvwSetParentNode tvwObj, Node, blnCheck
        End If
    End If
End Sub

Public Function VerifySignatureRecored_bak(ByVal intTache As Integer, ByVal Int���� As Integer, ByVal strNo As String, _
    ByVal lngҩ��id As Long, Optional ByVal LngID As Long, Optional ByVal date���� As Date) As Boolean
    '����ǩ��
    'intTache:1-��ҩ;2-��ҩ
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    Dim lng֤��ID As Long
    Dim strTimeStamp As String
    Dim strSignDate As String
    Dim intRule As Integer
    Dim lngǩ��id As Long
    
    'Ŀǰʹ�ù���
    intRule = 2
    On Error GoTo errHandle
    '��ȡǩ��Դ��
    gstrSQL = "Select A.ID, A.����, A.NO, A.���, A.�ⷿid, A.������id, A.�Է�����id, A.���ϵ��, A.ҩƷid, nvl(A.����,0) ����, " & _
        " A.������, To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') As ��������, A.��ҩ��, To_Char(A.��ҩ����,'yyyy-MM-dd hh24:mi:ss') As ��ҩ����, A.�����, To_Char(A.�������,'yyyy-MM-dd hh24:mi:ss') As �������, " & _
        " A.����id, A.����, A.Ƶ��, A.�÷�, Nvl(B.ǩ��ID, 0) As ǩ��ID " & _
        " From ҩƷ�շ���¼ A, ҩƷǩ����ϸ B,ҩƷǩ����¼ C " & _
        " Where A.id=B.�շ�id and B.ǩ��id=C.id and  ���� = [1] And No = [2] And �ⷿid + 0 = [3] "
    If LngID <> 0 Then
        gstrSQL = gstrSQL & " And a.id=[4] "
    Else
        If intTache = EsignTache.Dosage Then
            gstrSQL = gstrSQL & " And ��ҩ����=[4] And C.����=1"
        Else
            gstrSQL = gstrSQL & " And �������=[4] And C.����<>1"
        End If
    End If
    
    gstrSQL = gstrSQL & " Order By ����, NO, ���,A.id "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", Int����, strNo, lngҩ��id, IIf(LngID = 0, date����, LngID))
    
    With rsTmp
        If Not .EOF Then
            strSignSource = !���� & "," & !NO & "," & !�ⷿid & "," & !������id & "," & !�Է�����id & "," & !���ϵ��
            
            If intTache = EsignTache.Dosage Then
                strSignSource = strSignSource & "," & !��ҩ�� & "," & !��ҩ����
            Else
                strSignSource = strSignSource & "," & !����� & "," & !�������
            End If
        Else
            Exit Function
        End If
        
        strSignSource = strSignSource & "|"
        
        Do While Not .EOF
            lngǩ��id = !ǩ��ID
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !Id & "," & !��� & "," & !ҩƷID & "," & Val(Nvl(!����)) & "," & !����ID & "," & !���� & "," & !Ƶ�� & "," & !�÷�
            .MoveNext
        Loop
        
        strSignSource = strSignSource & strDetail
    End With
    
    '��֤ǩ��
    Call gobjESign.VerifySignature(strSignSource, lngǩ��id, 3)
    
    VerifySignatureRecored_bak = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function VerifySignatureRecoredGather(ByVal intTache As Integer, ByVal LngID As Long) As Boolean
    '��֤����ǩ�������ڻ��ܷ�ҩǩ��ʱ��ǩ��һ�Σ���֤ʱֻ�ܻ������з�ҩ��¼����Ϣ����֤
    '��Ҫע�Ᵽ�ֺ�ǩ��ʱ����Ϣ��ɸ�ʽһ��
    'intTache:1-��ҩ;2-��ҩ
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    Dim lng֤��ID As Long
    Dim strTimeStamp As String
    Dim strSignDate As String
    Dim intRule As Integer
    Dim lngǩ��id As Long
    Dim Int���� As Integer
    Dim strNo As String
    
    'Ŀǰʹ�ù���
    intRule = 2
    
    On Error GoTo errHandle
    
    'ȡǩ��ID
    gstrSQL = "Select b.ǩ��id From ҩƷǩ����¼ A, ҩƷǩ����ϸ B Where b.�շ�id = [1] And a.Id = b.ǩ��id And a.���� = 2 "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡǩ��ID", LngID)
    If rsTmp.RecordCount > 0 Then
        lngǩ��id = rsTmp!ǩ��ID
    Else
        Exit Function
    End If
        
    '��ȡǩ��Դ�ģ����ݵ�ǰ��¼�ҵ����ܷ�ҩʱһ��ǩ�������м�¼
    gstrSQL = "Select A.ID, A.����, A.NO, A.���, A.�ⷿid, A.������id, A.�Է�����id, A.���ϵ��, A.ҩƷid, nvl(A.����,0) ����, " & _
        " A.������, To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') As ��������, A.��ҩ��, To_Char(A.��ҩ����,'yyyy-MM-dd hh24:mi:ss') As ��ҩ����, A.�����, To_Char(A.�������,'yyyy-MM-dd hh24:mi:ss') As �������, " & _
        " A.����id, A.����, A.Ƶ��, A.�÷�, i.���㵥λ " & _
        " From ҩƷ�շ���¼ A, ������ĿĿ¼ I, ҩƷ��� B " & _
        " Where a.ҩƷid = b.ҩƷid And i.Id = b.ҩ��id And a.Id In (Select �շ�id From ҩƷǩ����ϸ Where ǩ��id = [1]) " & _
        " Order By ����, NO, ��� "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", lngǩ��id)
    
    With rsTmp
        Do While Not .EOF
            If Int���� <> !���� Or strNo <> !NO Then
                '������Ϣ����ϸ��Ϣ֮����|�ָ�
                If strDetail <> "" Then strSignSource = strSignSource & "|" & strDetail
                
                '��ͬ����֮����#�ָ�
                strSignSource = IIf(strSignSource = "", "", strSignSource & "#") & !���� & "," & !NO & "," & !�ⷿid & "," & !������id & "," & !�Է�����id & "," & !���ϵ��
                If intTache = EsignTache.send Or intTache = EsignTache.returnStep Then
                    strSignSource = strSignSource & "," & IIf(IsNull(!�����), "", !�����) & "," & IIf(IsNull(!�������), "", Format(!�������, "yyyy-MM-dd HH:mm:ss"))
                End If
                
                Int���� = !����
                strNo = !NO
                strDetail = ""
            End If
            
            'ͬһ���ݲ�ͬ��ϸ֮����;�ָ�
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !Id & "," & !��� & "," & !ҩƷID & "," & Val(Nvl(!����)) & "," & !����ID & "," & IIf(IsNull(!����), "", FormatEx(!����, 5) & Nvl(!���㵥λ)) & "," & IIf(IsNull(!Ƶ��), "", !Ƶ��) & "," & IIf(IsNull(!�÷�), "", !�÷�)
            
            .MoveNext
        Loop
    End With
    
    If strDetail <> "" Then strSignSource = strSignSource & "|" & strDetail

    '��֤ǩ��
    Call gobjESign.VerifySignature(strSignSource, lngǩ��id, 3)
    
    VerifySignatureRecoredGather = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetCheck�ⷿ(ByVal lng�ⷿid As Long) As Integer
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1] "
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ƿ���������", lng�ⷿid)
    If Not rstemp.EOF Then GetCheck�ⷿ = Nvl(rstemp!�����, 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSignatureRecored(ByVal intTache As Integer, ByVal Int���� As Integer, ByVal strNo As String, _
        ByVal lngҩ��id As Long, ByRef strǩ����¼ As String, Optional ByVal LngID As Long, _
        Optional ByVal date���� As Date, Optional str������ As String, Optional lng��ҩ��id As Long = 0, Optional blnCheck As Boolean) As Boolean
    '����ǩ��
    'intTache:1-��ҩ;2-��ҩ
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    Dim lng֤��ID As Long
    Dim strTimeStamp As String
    Dim strTimeStampInfo As String
    Dim str�շ�ids As String
    Dim strSignDate As String
    Dim intRule As Integer
    
    'Ŀǰʹ�ù���
    intRule = 2
    
    gstrSQL = "Select ID, ����, NO, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, nvl(����,0) ����, " & _
        " ������, To_Char(��������,'yyyy-MM-dd hh24:mi:ss') As ��������, ��ҩ��, To_Char(��ҩ����,'yyyy-MM-dd hh24:mi:ss') As ��ҩ����, �����, To_Char(�������,'yyyy-MM-dd hh24:mi:ss') As �������, " & _
        " ����id, ����, Ƶ��, �÷� " & _
        " From ҩƷ�շ���¼ " & _
        " Where  ���� = [1] And No = [2] And �ⷿid + 0 = [3] "
    If LngID <> 0 Then
        gstrSQL = gstrSQL & " And id=[4] "
    Else
        If intTache = EsignTache.Dosage Then
            gstrSQL = gstrSQL & " And Mod(��¼״̬,3)=1  And ����� Is Null "
        ElseIf intTache = EsignTache.send Then
            gstrSQL = gstrSQL & " And ����� Is Null  "
        ElseIf intTache = EsignTache.returnStep Then
            gstrSQL = gstrSQL & " And �������=[4] "
        End If
            
    End If
    
    gstrSQL = gstrSQL & " Order By ����, NO, ��� "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", Int����, strNo, lngҩ��id, IIf(LngID = 0, date����, LngID))
    
    With rsTmp
        If Not .EOF Then
            strSignSource = !���� & "," & !NO & "," & IIf(lng��ҩ��id = 0, !�ⷿid, lng��ҩ��id) & "," & !������id & "," & !�Է�����id & "," & !���ϵ��
                
            If intTache = EsignTache.Dosage Then
                If str������ <> "" Then
                    strSignSource = strSignSource & "," & str������ & "," & Format(date����, "yyyy-mm-dd hh:mm:ss")
                Else
                    strSignSource = strSignSource & "," & !��ҩ�� & "," & !��ҩ����
                End If
            ElseIf intTache = EsignTache.send Then
                strSignSource = strSignSource & "," & str������ & "," & Format(date����, "yyyy-mm-dd hh:mm:ss")
            ElseIf intTache = EsignTache.returnStep Then
                strSignSource = strSignSource & "," & "," & !����� & "," & !�������
            End If
        Else
            Exit Function
        End If
        
        strSignSource = strSignSource & "|"
        
        Do While Not .EOF
            str�շ�ids = IIf(str�շ�ids = "", "", str�շ�ids & ",") & !Id
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !Id & "," & !��� & "," & !ҩƷID & "," & Val(Nvl(!����)) & "," & !����ID & "," & !���� & "," & !Ƶ�� & "," & !�÷�
            .MoveNext
        Loop
        
        strSignSource = strSignSource & strDetail
    End With
    
    strSign = gobjESign.Signature(strSignSource, gstrDbUser, lng֤��ID, strTimeStamp, , strTimeStampInfo, blnCheck)
    If strSign <> "" Then
        If strTimeStamp <> "" Then
            strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
        Else
            strTimeStamp = "NULL"
        End If
        
        If strTimeStampInfo = "" Then strTimeStampInfo = "NULL"
        
        strǩ����¼ = intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & "," & strTimeStamp & ",'" & strTimeStampInfo & "'," & intTache & ",'" & str�շ�ids & "'"
        GetSignatureRecored = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetSignatureRecoredGather(ByVal intTache As Integer, ByVal rsData As ADODB.Recordset, ByVal lng�ⷿid As Long, ByVal str��ҩ�� As String, ByVal str����� As String, ByVal str������� As String, ByRef strǩ����¼ As String, Optional blnCheck As Boolean) As Boolean
    '����ǩ�������ڻ��ܷ�ҩ��ÿ�η�ҩ������һ��ǩ��
    'ֱ�Ӵӷ�ҩ���ݼ���֯���ݣ����ٶ�ȡ���ݿ����
    'intTache:1-��ҩ;2-��ҩ
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    Dim lng֤��ID As Long
    Dim strTimeStamp As String
    Dim strTimeStampInfo As String
    Dim str�շ�ids As String
    Dim strSignDate As String
    Dim intRule As Integer
    Dim Int���� As Integer
    Dim strNo As String
    
    'Ŀǰʹ�ù���
    intRule = 2
    
    With rsData
'        .Filter = "ִ��״̬=1"
        
        '���򷽷����ܱ�
        .Sort = "����,NO,���"
    
        Do While Not .EOF
            If Int���� <> !���� Or strNo <> !NO Then
                '������Ϣ����ϸ��Ϣ֮����|�ָ�
                If strDetail <> "" Then strSignSource = strSignSource & "|" & strDetail
                
                '��ͬ����֮����#�ָ�
                strSignSource = IIf(strSignSource = "", "", strSignSource & "#") & !���� & "," & !NO & "," & lng�ⷿid & "," & !������id & "," & !��ҩ����ID & "," & !���ϵ��
                If intTache = EsignTache.send Or intTache = EsignTache.returnStep Then
                    strSignSource = strSignSource & "," & str����� & "," & str�������
                End If
                
                Int���� = !����
                strNo = !NO
                strDetail = ""
            End If
            
            str�շ�ids = IIf(str�շ�ids = "", "", str�շ�ids & ",") & !�շ�ID
            
            'ͬһ���ݲ�ͬ��ϸ֮����;�ָ�
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !�շ�ID & "," & !��� & "," & !ҩƷID & "," & Val(Nvl(!����)) & "," & !����ID & "," & !ԭʼ���� & "," & !Ƶ�� & "," & !�÷�
            
            .MoveNext
        Loop
    End With
    
    If strDetail <> "" Then strSignSource = strSignSource & "|" & strDetail
    
    '��ȡǩ����Ϣ
    strSign = gobjESign.Signature(strSignSource, gstrDbUser, lng֤��ID, strTimeStamp, , strTimeStampInfo, blnCheck)
    If strSign <> "" Then
        If strTimeStamp <> "" Then
            strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
        Else
            strTimeStamp = "NULL"
        End If
        
        If strTimeStampInfo = "" Then strTimeStampInfo = "NULL"
        
        strǩ����¼ = intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & "," & strTimeStamp & ",'" & strTimeStampInfo & "'," & intTache & ",'" & str�շ�ids & "'"
        GetSignatureRecoredGather = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function






Public Function DelSignatureRecored_Check(ByVal intTache As Integer, ByVal Int���� As Integer, ByVal strNo As String, ByVal lngҩ��id As Long, ByRef lngǩ��id As Long, Optional ByVal LngID As Long, Optional ByVal date���� As Date) As Boolean
    'intRule:1-��ҩ;2-��ҩ
    Dim rsTmp As ADODB.Recordset
    Dim strSignSource As String
    Dim strDetail As String
    Dim strSign As String
    
    On Error GoTo errHandle
    '��ȡǩ��Դ��
    gstrSQL = "Select A.ID, A.����, A.NO, A.���, A.�ⷿid, A.������id, A.�Է�����id, A.���ϵ��, A.ҩƷid, nvl(A.����,0) ����, " & _
        " A.������, To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') As ��������, A.��ҩ��, To_Char(A.��ҩ����,'yyyy-MM-dd hh24:mi:ss') As ��ҩ����, A.�����, To_Char(A.�������,'yyyy-MM-dd hh24:mi:ss') As �������, " & _
        " A.����id, A.����, A.Ƶ��, A.�÷�, Nvl(B.ǩ��ID, 0) As ǩ��ID " & _
        " From ҩƷ�շ���¼ A, ҩƷǩ����ϸ B" & _
        " Where A.id=B.�շ�id(+) and ���� = [1] And No = [2] And �ⷿid + 0 = [3] "
    If LngID <> 0 Then
        gstrSQL = gstrSQL & " And A.id=[4] "
    Else
        If intTache = EsignTache.Dosage Then
            gstrSQL = gstrSQL & " And Mod(��¼״̬,3)=1 "
        Else
            gstrSQL = gstrSQL & " And �������=[4] "
        End If
            
    End If
    
    gstrSQL = gstrSQL & " Order By ����, NO, ���,Id "
    
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", Int����, strNo, lngҩ��id, IIf(LngID = 0, date����, LngID))
    
    With rsTmp
        If Not .EOF Then
            If CLng(!ǩ��ID) = 0 Then
                '�������������ҵ����ʱû��ʹ�õ���ǩ����������ʹ���ˣ���������Ͳ��������ǩ��������������ҩ��������
                DelSignatureRecored_Check = True
                Exit Function
            End If
            
            '���USB-KEY
            If Not gobjESign.CheckCertificate(gstrDbUser) Then Exit Function
            
            lngǩ��id = CLng(!ǩ��ID)
            strSignSource = !���� & "," & !NO & "," & !�ⷿid & "," & !������id & "," & !�Է�����id & "," & !���ϵ�� & "," & _
                !������ & "," & !�������� & "," & !��ҩ�� & "," & !��ҩ����
            If intTache = EsignTache.send Then
                strSignSource = strSignSource & "," & !����� & "," & !�������
            End If
        Else
            Exit Function
        End If
        
        strSignSource = strSignSource & "|"
        
        Do While Not .EOF
            strDetail = IIf(strDetail = "", "", strDetail & ";") & !Id & "," & !��� & "," & !ҩƷID & "," & Val(Nvl(!����)) & "," & !����ID & "," & !���� & "," & !Ƶ�� & "," & !�÷�
            .MoveNext
        Loop
        
        strSignSource = strSignSource & strDetail
    End With
    
    '��֤ǩ��
    If Not gobjESign.VerifySignature(strSignSource, lngǩ��id, 3) Then Exit Function
    DelSignatureRecored_Check = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub CheckStopMedi(ByVal varInput As Variant, Optional ByRef Int��ҩ As Integer)
    '���ҩƷ�Ƿ�ͣ��
    'varInput���ָ�ʽ�����뵥����Ϣ������|No��;����ҩƷID������ʽ��ҩƷID1��ҩƷID2.....��
    'int��ҩ:0-������ҩ��1-��ҩ��2-��ҩ����ͣ��ҩƷ
    Dim rstemp As ADODB.Recordset
    Dim strMsg As String
    Dim Int���� As Integer
    Dim strNo As String
    Dim n As Integer
    
    On Error GoTo errHandle
    If InStr(varInput, "|") > 0 Then
        Int���� = Mid(varInput, 1, InStr(varInput, "|") - 1)
        strNo = Mid(varInput, InStr(varInput, "|") + 1)
        
        gstrSQL = "Select /*+rule*/ Distinct '(' || C.���� || ')' || Nvl(B.����, C.����) As ҩƷ��Ϣ " & _
                " From ҩƷ�շ���¼ A, �շ���Ŀ���� B, �շ���ĿĿ¼ C " & _
                " Where A.ҩƷid = C.ID And A.ҩƷid = B.�շ�ϸĿid(+) And B.����(+) = 3 " & _
                " And Nvl(C.����ʱ��, To_Date('3000-01-01', 'yyyy-MM-dd')) <> To_Date('3000-01-01', 'yyyy-MM-dd') " & _
                " And A.���� = [1] And A.NO = [2]"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "���ͣ��ҩƷ", Int����, strNo)
    Else
        gstrSQL = "Select /*+ Rule*/ Distinct '(' || C.���� || ')' || Nvl(B.����, C.����) As ҩƷ��Ϣ " & _
                " From Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) A, �շ���Ŀ���� B, �շ���ĿĿ¼ C " & _
                " Where A.Column_Value = C.ID  And A.Column_Value = B.�շ�ϸĿid(+) And B.����(+) = 3 " & _
                " And Nvl(C.����ʱ��, To_Date('3000-01-01', 'yyyy-MM-dd')) <> To_Date('3000-01-01', 'yyyy-MM-dd') "
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "���ͣ��ҩƷ", varInput)
    End If
    
    With rstemp
        If Not .EOF Then
            For n = 1 To .RecordCount
                If n > 5 Then
                    strMsg = strMsg & vbCrLf & "��������" & .RecordCount - 5 & "��ҩƷ......"
                    Exit For
                End If
                strMsg = IIf(strMsg = "", "", strMsg & "," & vbCrLf) & !ҩƷ��Ϣ
                .MoveNext
            Next
            
            strMsg = "ע�⣬����ҩƷ�ѱ�ͣ�ã�" & vbCrLf & strMsg
        End If
    End With
    
    If strMsg <> "" Then
        If Int��ҩ <> 0 Then
            MsgBox strMsg & vbCrLf & "ͣ�õ�ҩƷ��������ҩ���������ø�ҩƷ���ſ��Խ�����ҩ����", vbInformation, gstrSysName
            Int��ҩ = 2
        Else
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Check�ⷿ����(ByVal lng�ⷿid As Long, ByVal lngҩƷid As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim bln����Ƿ���� As Boolean, bln���� As Boolean, bln�ⷿ As Boolean
    '��������true������������false
    On Error GoTo errHandle
    Check�ⷿ���� = False
    
    '���ж��Ƿ��ǿⷿ
    gstrSQL = "select ����ID from ��������˵�� where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng�ⷿid)
    
    bln�ⷿ = (rsCheck.EOF)
        
    '�ж϶�Ӧ��ҩƷĿ¼�еķ�������
    gstrSQL = " Select Nvl(ҩ�����,0) ��������,nvl(ҩ������,0) ҩ���������� " & _
              " From ҩƷ��� Where ҩƷID=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩƷĿ¼�еķ�������", lngҩƷid)
              
    If bln�ⷿ Then
        Check�ⷿ���� = (rsCheck!�������� = 1)
    Else
        Check�ⷿ���� = (rsCheck!ҩ���������� = 1)
    End If

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckNumStock(ByVal objVSF As BillEdit, ByVal lng�ⷿid As Long, ByVal lntColҩƷid As Integer, _
    ByVal intCol���� As Integer, ByVal intCol���� As Integer, ByVal intCol����ϵ�� As Integer, _
    ByVal intMethod As Integer, Optional int���ҵ�� As Integer = 0, Optional int��ʵ���� As Integer = 0) As String
    '���ܣ���˳����൥��ʱ��������ʵ�������Ƿ��㹻
    '������objVSF-��Ҫ���ı��;lng�ⷿid��intcol����-���������У�intCol����-���������У�intCol����ϵ��-����ϵ��������
    '������int��ʵ����-��ʵ����������(������ˣ�����)��intMethod��1-������ˣ�2-������3-�˿����
    '������int���ҵ��0-��⣻1-����
    '����ֵ�����о����ҩƷ���ƣ�Ϊ��-���ͨ�����������㣻��Ϊ��-���δͨ��������������
    Dim objCol As Collection         '��ʹ�õ���������
    Dim dblNum As Double
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim strTemp As String
    Dim lngҩƷid As Long
    Dim lng���� As Long
    Dim rsData As ADODB.Recordset
    Dim strKey As String
    Dim vardrug As Variant
    Dim lngRow As Long
    Dim strArray As String
    
    On Error GoTo errHandle
    
    '����ϱ�������������������Ҫ�ǿ��ǲ����������
    Set objCol = New Collection
    With objVSF
        If .rows < 2 Then Exit Function
        For lngRow = 1 To .rows - 1
            dblNum = 0
            If .TextMatrix(lngRow, lntColҩƷid) <> "" Then
                For Each vardrug In objCol
                    If vardrug(0) = .TextMatrix(lngRow, lntColҩƷid) & "," & Val(.TextMatrix(lngRow, intCol����)) Then
                        dblNum = vardrug(1)
                        objCol.Remove vardrug(0)
                        Exit For
                    End If
                Next
                strKey = .TextMatrix(lngRow, lntColҩƷid) & "," & Val(.TextMatrix(lngRow, intCol����))
                
                '�������������С������ԭʼ���ݿ������������
                If Fix(Val(.TextMatrix(lngRow, intCol����))) <> Val(.TextMatrix(lngRow, intCol����)) And int��ʵ���� <> 0 Then
                    strArray = dblNum + Val(.TextMatrix(lngRow, int��ʵ����))
                Else
                    strArray = dblNum + (Val(.TextMatrix(lngRow, intCol����)) * Val(.TextMatrix(lngRow, intCol����ϵ��)))
                End If
                
                objCol.Add Array(strKey, strArray), strKey
            End If
        Next
    End With
    
    For Each varNum In objCol
        strTemp = varNum(0)  '��ʽ��ҩƷid,����
        dblNum = varNum(1)
        varTemp = Split(strTemp, ",")
        If int���ҵ�� = 0 Then '���
            If intMethod = 1 Then '�������
                If dblNum < 0 Then
                    '������⣬��Ҫ����棬������Ҫ�жϿ���Ƿ����
                    dblNum = Abs(dblNum)
                Else
                    '������⣬������棬���Բ����
                    dblNum = 0
                End If
            ElseIf intMethod = 2 Then
                '����
                If dblNum < 0 Then
                    dblNum = 0
                Else
                    dblNum = dblNum
                End If
            ElseIf intMethod = 3 Then
                '�˿���ˣ��˿����¼������
                dblNum = dblNum
            End If
        Else    '����
            If intMethod = 1 Then '�������
                If dblNum < 0 Then
                    '������⣬��Ҫ����棬������Ҫ�жϿ���Ƿ����
                    dblNum = 0
                Else
                    '������⣬������棬���Բ����
                    dblNum = dblNum
                End If
            ElseIf intMethod = 2 Then
                '����
                If dblNum < 0 Then
                    dblNum = Abs(dblNum)
                Else
                    dblNum = 0
                End If
            End If
        End If
        
        'ֻ�����������ж�
        If dblNum > 0 Then
            lngҩƷid = varTemp(0)
            lng���� = varTemp(1)
            If Check�ⷿ����(lng�ⷿid, lngҩƷid) = False Then
                lng���� = 0
            End If
            
            gstrSQL = "Select (a.ʵ������ - [1]) As ʣ������, b.����, b.����" & vbNewLine & _
                        "From ҩƷ��� A, �շ���ĿĿ¼ B" & vbNewLine & _
                        "Where a.ҩƷid = b.Id And a.ҩƷid = [2] And a.�ⷿid = [3] And Nvl(a.����, 0) = [4] And b.��� In ('5', '6', '7') And a.���� = 1"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "�����", dblNum, lngҩƷid, lng�ⷿid, lng����)
            
            If rsData.RecordCount = 0 Then
                gstrSQL = "select ����,���� from �շ���ĿĿ¼ where id=[1]"
                Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "�����", lngҩƷid)
                CheckNumStock = "[" & rsData!���� & "]" & rsData!����
                Exit Function
            Else
                If rsData!ʣ������ >= 0 Then
                    CheckNumStock = ""
                Else
                    CheckNumStock = "[" & rsData!���� & "]" & rsData!����
                    Exit Function
                End If
            End If
        End If
    Next
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckUsableNum( _
    ByVal lng�ⷿid As Long, _
    ByVal lngҩƷid As Long, _
    ByVal lng���� As Long, _
    ByVal dbl��д���� As Double, _
    ByVal dbl����ϵ�� As Double, _
    ByVal strNo As String, _
    ByVal Int���� As Integer, _
    ByVal int����� As Integer, _
    ByVal int�������� As Integer, _
    Optional int��� As Integer, _
    Optional dblSum As Double) As Boolean
    '������д����ʱ���������������Ƿ��㹻����������/�޸ģ����������
    '����ֵ true-ͨ����飬false-û��ͨ�����
    '��Σ�dbl��д�����ǽ��浥λ����
    '      strNo="", ��-� �ǿ�-�޸ģ��޸�ʱ��Ҫ�ų���ǰ��������
    '      dblSum �����ҩƷ����д�����������ڳ���/�������ʱ
    '1.���δ���0�ǰ����μ�飬����=0���Ǳ�ʾ�������飻�޸�״̬ʱҪ����ԭ����������������Ҫ���ǿ��ܱ�����δ�������ηֽ��ҵ��ռ�õ�����
    '2.�������Ҫ�����ľͲ��õ��ú�������������
    '3.����/�ƿⵥ�ݳ���ʱ���⴦��:
    '�������ȡԭ�������Σ�ע��Ҫ��ԭ������ⷿ(����ʱΪ���ⷿ)������ʱ��֧�ֶԳ���������޸ģ����Բ��������е��ݵ������Ҫ�ӽ��洫��������
    '4.���ѻ��ֹʱ���ݷ���������������������ͬ
    Dim dblNum As Double
    Dim rsData As ADODB.Recordset
    Dim dblCheck As Boolean
    Dim bln�������� As Boolean
    Dim bln�������� As Boolean
    Dim strSqlStock As String, strSqlStockBatch As String  '����������������ͷ�������
    Dim strSqlSum As String, strSqlSumBatch As String      '���ϲ�δ��˵��������������ͷ�������
    Dim lng�������� As Long
    Dim blnNewNo As Boolean '�Ƿ���������
    Dim dbl����д���� As Double
    
    On Error GoTo errHandle
    
    If int����� = 0 Then CheckUsableNum = True: Exit Function

    If Int���� = 6 And int��� > 0 Then
        blnNewNo = True
        
        'ȡԭ�����Ǳʵ�����
        gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ�շ���¼ Where  " & _
            " �ⷿid=[1] And ���� = [2] And NO = [3] And ��� = [4] And ҩƷid = [5] And ���ϵ�� = 1"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�������", lng�ⷿid, Int����, strNo, int��� + 1, lngҩƷid)
        
        If rsData.RecordCount = 0 Then Exit Function
        
        lng�������� = rsData!����
        
        If lng�������� = 0 Then
            '��������Ϊ��������������������
            dbl����д���� = dblSum
        Else
            '��������Ϊ����������������ε���д����
            dbl����д���� = dbl��д����
        End If
    Else
        blnNewNo = (strNo = "")
        lng�������� = lng����
        dbl����д���� = dbl��д����
    End If
        
    strSqlStock = "Select Sum(Nvl(��������, 0)) As �������� From ҩƷ��� Where ����=1 And �ⷿid = [1] And ҩƷid = [2]"
    strSqlStockBatch = "Select Sum(Nvl(��������, 0)) As �������� From ҩƷ��� Where ����=1 And �ⷿid = [1] And ҩƷid = [2] And nvl(����,0) = [3] "
    strSqlSum = "Select Sum(��������) As ��������" & vbNewLine & _
                " From (Select Nvl(��������, 0) As ��������" & vbNewLine & _
                "       From ҩƷ���" & vbNewLine & _
                "       Where ����=1 And �ⷿid = [1] And ҩƷid = [2] " & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Abs(a.ʵ������ * Nvl(a.����, 1)) As ��������" & vbNewLine & _
                "       From ҩƷ�շ���¼ A" & vbNewLine & _
                "       Where a.������� Is Null And a.�ⷿid = [1] And a.ҩƷid + 0 = [2]  And a.No = [4] And a.���� = [5])"
    strSqlSumBatch = "Select Sum(��������) As ��������" & vbNewLine & _
                    " From (Select Nvl(��������, 0) As ��������" & vbNewLine & _
                    "       From ҩƷ���" & vbNewLine & _
                    "       Where ����=1 And �ⷿid = [1] And ҩƷid = [2]  And nvl(����,0) = [3] " & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Abs(a.ʵ������ * Nvl(a.����, 1)) As ��������" & vbNewLine & _
                    "       From ҩƷ�շ���¼ A" & vbNewLine & _
                    "       Where a.������� Is Null And a.�ⷿid = [1] And a.ҩƷid + 0 = [2]  And a.No = [4] And a.���� = [5]  And nvl(����,0) = [3] )"
    
    If lng���� = 0 Then
        '1.�����������
        If blnNewNo = True Then
            '1.1����ǵ�������״̬��ֱ�ӿ�����ܿ��������Ƿ��㹻
            gstrSQL = strSqlStock
        Else
            '1.2����ǵ����޸�״̬��Ҫ�ϲ�ԭ��������
            gstrSQL = strSqlSum
        End If
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "��������", lng�ⷿid, lngҩƷid, lng��������, strNo, Int����)
        
        If Nvl(rsData.Fields(0), 0) > 0 Then
            dblNum = zlStr.FormatEx(rsData.Fields(0) / dbl����ϵ��, int��������, True, False)
        End If
        
        If dblNum < dbl����д���� Then
            dblCheck = True
            bln�������� = True
        End If
    Else
        '2.���������
        If blnNewNo = True Then
            '2.1����ǵ�������״̬��ֱ�ӿ�����ܿ��������Ƿ��㹻
            gstrSQL = strSqlStockBatch
        Else
            '2.2����ǵ����޸�״̬��Ҫ�ϲ�ԭ��������
            gstrSQL = strSqlSumBatch
        End If
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "��������", lng�ⷿid, lngҩƷid, lng��������, strNo, Int����)

        If Nvl(rsData.Fields(0), 0) > 0 Then
            dblNum = zlStr.FormatEx(rsData.Fields(0) / dbl����ϵ��, int��������, True, False)
        End If
        
        If dblNum < dbl����д���� Then
            '2.2.1��������
            dblCheck = True
            bln�������� = True
        End If
    End If
        
    '��治��ʱ���ѻ��ֹ
    If dblCheck = True Then
        gstrSQL = "select ����,���� from �շ���ĿĿ¼ where id=[1]"
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "�����", lngҩƷid)
                    
        Select Case int�����
        Case 1  '��ʾ
            If Int���� = 2 Then '�������
                If bln�������� = True Then
                    If MsgBox("���ҩƷ��[" & rsData!���� & "]" & rsData!���� & "���Ŀ��ÿ�治�㣬���ܱ�����δ��˵���ռ�ã��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf bln�������� = True Then
                    If MsgBox("���ҩƷ��[" & rsData!���� & "]" & rsData!���� & "���������������˿��ÿ��" & dblNum & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            Else
                If bln�������� = True Then
                    If MsgBox("��[" & rsData!���� & "]" & rsData!���� & "���Ŀ��ÿ�治�㣬���ܱ�����δ��˵���ռ�ã��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf bln�������� = True Then
                    If MsgBox("��[" & rsData!���� & "]" & rsData!���� & "���������������˿��ÿ��" & dblNum & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        Case 2  '��ֹ
            If Int���� = 2 Then '�������
                If bln�������� = True Then
                    MsgBox "���ҩƷ��[" & rsData!���� & "]" & rsData!���� & "���Ŀ��ÿ�治�㣬���ܱ�����δ��˵���ռ�ã����ܳ��⣡", vbInformation, gstrSysName
                ElseIf bln�������� = True Then
                    MsgBox "���ҩƷ��[" & rsData!���� & "]" & rsData!���� & "���������������˿��ÿ��" & dblNum & "�����ܳ��⣡", vbInformation, gstrSysName
                End If
            Else
                If bln�������� = True Then
                    MsgBox "��[" & rsData!���� & "]" & rsData!���� & "���Ŀ��ÿ�治�㣬���ܱ�����δ��˵���ռ�ã����ܳ��⣡", vbInformation, gstrSysName
                ElseIf bln�������� = True Then
                    MsgBox "��[" & rsData!���� & "]" & rsData!���� & "���������������˿��ÿ��" & dblNum & "�����ܳ��⣡", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End Select
    End If
    CheckUsableNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get��������(ByVal lng�ⷿid As Long, ByVal lngҩƷid As Long) As Integer
    '����ָ���ⷿ��ָ��ҩƷ�ķ�������
    '���أ�0-��������1-����
    Dim rsCheck As New ADODB.Recordset
    Dim int���� As Integer
    Dim blnҩ�� As Boolean
    Dim strsql As String
        
    On Error GoTo errHandle
    
    '�ж��Ƿ���ҩ�����Ƽ���
    strsql = "select ����ID from ��������˵�� where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(strsql, "Get��������", lng�ⷿid)

    blnҩ�� = (Not rsCheck.EOF)
        
    '�ж϶�Ӧ��ҩƷĿ¼�еķ�������
    strsql = " Select Nvl(ҩ�����,0) As ҩ�����,nvl(ҩ������,0) As ҩ������ " & _
              " From ҩƷ��� Where ҩƷID=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(strsql, "Get��������", lngҩƷid)
              
    If blnҩ�� Then
        int���� = rsCheck!ҩ������
    Else
        int���� = rsCheck!ҩ�����
    End If
    
    Get�������� = int����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




Public Function CheckStrickUsable(ByVal Int���� As Integer, ByVal lng�ⷿid As Long, _
        ByVal lngҩƷid As Long, ByVal strҩƷ���� As String, _
        ByVal lng���� As Long, ByVal dbl�������� As Double, ByVal int����� As Integer, _
        Optional ByVal strNo As String = "", Optional ByVal int��� As Integer = 0) As Boolean
    '��������ʱ��飺ԭ�������ⷿ�Ƿ���������㹻�������������ڻ�С��ʵ����������ʵ�ʳ����������ܴ��ڿ�������
    '�����ƿⵥ�ݡ�����ⵥ����Ҫȡԭ��������Ǳʵ����Σ��ٸ���������ȡ����������
    '����������⡢Э����ⵥ�ݣ�������ȫ�����������Ը��ݵ��ݺţ������ȡ���������������Ϳ����������Ƚ�
    '�������ݿ�ֱ�Ӹ�������ȡ����������
    'int����飺��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    'ֻ�г���ʱ�ǳ������ͣ�ԭ������������ͣ���Ҫ���˼�飺�����⡢������⣨ԭ��������Ǳʣ���Э����⣨ԭ��������Ǳʣ���������⡢�ƿ⣨ԭ��������Ǳʣ�
    
    Dim rstemp As ADODB.Recordset
    Dim lng������� As Long
    Dim dbl�������� As Double
    
    On Error GoTo errHandle
    
    If int����� = 0 Then CheckStrickUsable = True: Exit Function
    
    If Int���� = 2 Or Int���� = 3 Then  '������⡢Э����ⵥ��
        If strNo = "" Or int��� = 0 Then Exit Function
        gstrSQL = "Select 1 From ҩƷ�շ���¼ A, ҩƷ��� B " & _
            " Where A.���� = [1] And A.NO = [2] And A.��� = [3] And A.��¼״̬ = 1 And A.���ϵ�� = 1 And B.���� = 1 And A.�ⷿid = B.�ⷿid And A.ҩƷid = B.ҩƷid And " & _
            " Nvl(A.����, 0) = Nvl(B.����, 0) And A.ʵ������ > B.�������� And Rownum = 1"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "����������", Int����, strNo, int���)
        
        '���������̽�����ʾ���ֹ
        If rstemp.RecordCount > 0 Then
            Select Case int�����
            Case 1  '��ʾ
                If MsgBox(strҩƷ���� & "�Ŀ�治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Case 2  '��ֹ
                MsgBox strҩƷ���� & "�Ŀ�治�㣡", vbInformation, gstrSysName
                Exit Function
            End Select
        End If
    Else
        If Int���� = 6 Or Int���� = 4 Then   '�ƿⵥ��������ⵥ
            If strNo = "" Or int��� = 0 Then Exit Function
            
            gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And ��� = [3] And ҩƷid = [4] And ���ϵ�� = 1"
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ�������", Int����, strNo, int���, lngҩƷid)
            
            If rstemp.RecordCount = 0 Then Exit Function
            
            lng������� = rstemp!����
        Else
            '�������ݸ��ݴ����������ȡ����������
            lng������� = lng����
        End If
        
        gstrSQL = "Select Nvl(��������, 0) �������� From ҩƷ��� Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And Nvl(����, 0) = [3] "
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng�ⷿid, lngҩƷid, lng�������)
        
        If rstemp.RecordCount > 0 Then
            dbl�������� = rstemp!��������
        End If
        
        '���������̽�����ʾ���ֹ
        If dbl�������� < Abs(dbl��������) Then
            Select Case int�����
            Case 1  '��ʾ
                If MsgBox(strҩƷ���� & "�Ŀ�治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Case 2  '��ֹ
                MsgBox strҩƷ���� & "�Ŀ�治�㣡", vbInformation, gstrSysName
                Exit Function
            End Select
        End If
    End If
    
    CheckStrickUsable = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub LoadBillControl()
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(ʱ������, 0) ʱ������, Nvl(���˵���, 0) ���˵���, Nvl(�������, 0) ������� From ���ݲ������� Where ��Աid = [1] And ���� = 9"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "���ݲ�������", glngUserId)
    
    If Not rsTmp.EOF Then
        gtype_myBillControl.bln�Ƿ���� = True
        gtype_myBillControl.intʱ������ = rsTmp!ʱ������
        gtype_myBillControl.bln���˵��� = (rsTmp!���˵��� = 1)
        gtype_myBillControl.dbl������� = rsTmp!�������
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckBillControl(ByVal IntOper As Integer, ByVal IntBillStyle As Integer, ByVal strNo As String, ByVal dblMoney As Double) As Boolean
    '--���ݵ��ݲ������Ʊ���鵱ǰ����Ա�Ƿ������������
    'IntOper:1-��ҩ;2-ȡ����ҩ;3-��ҩ;4-��ҩ
    Dim rstemp As New ADODB.Recordset
    Dim bln�Ƿ���η�ҩ As Boolean
    
    On Error GoTo errHandle
    If gtype_myBillControl.bln�Ƿ���� = False Then
        CheckBillControl = True
        Exit Function
    End If
    
    
    '���ʱ������
    If gtype_myBillControl.intʱ������ > 0 Then
        If IntOper <> 4 Then
            gstrSQL = "Select Distinct �������� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And Mod(��¼״̬, 3) = 1 And ��¼״̬ <> 1 And ����� Is Null"
        Else
            gstrSQL = "Select Distinct �������� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And Mod(��¼״̬, 3) = 1 And ����� Is Not Null"
        End If
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��鵥�ݲ�������", IntBillStyle, strNo)
         
        If Not rstemp.EOF Then
            If DateDiff("d", Format(rstemp!��������, "yyyy-mm-dd hh:mm:ss"), Sys.Currentdate) > gtype_myBillControl.intʱ������ Then
                MsgBox "����[" & strNo & "]���������������ʱ�ޣ����ܽ��в�����"
                Exit Function
            End If
        Else
            bln�Ƿ���η�ҩ = True
        End If
    End If
    
    '����Ƿ�����������˵���
    If gtype_myBillControl.bln���˵��� Then
        If IntOper <> 4 Then
            gstrSQL = "Select ����� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And Mod(��¼״̬, 3) = 2 And ����� Is Not Null Order By ������� Desc"
        Else
            gstrSQL = "Select ����� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And Mod(��¼״̬, 3) = 1 And ����� Is Not Null Order By ������� Desc"
        End If
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��鵥�ݲ�������", IntBillStyle, strNo)
         
        If Not rstemp.EOF Then
            If rstemp!����� <> gstrUserName Then
                MsgBox "����[" & strNo & "]�ϴβ����˲��ǵ�ǰ����Ա�����ܽ��в�����"
                Exit Function
            End If
        End If
    End If
    
    '���������
    If gtype_myBillControl.dbl������� > 0 And bln�Ƿ���η�ҩ = False Then
        If gtype_myBillControl.dbl������� < dblMoney Then
            MsgBox "����[" & strNo & "]��������������������ܽ��в�����"
            Exit Function
        End If
    End If
    
    CheckBillControl = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPrice(ByVal lngBillId As Long, ByRef strMsg As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    '�ж��ۼ��Ƿ��ǵ�ǰ�����ۼ�
    
    CheckPrice = True       'ͳһ��Oracle�洢���̼��
    Exit Function
    
    'ȡԭʼ�۸���ּ�
    On Error GoTo errHandle
    gstrSQL = _
        "Select Round(a.ԭ��, b.����) ԭ��, Round(a.�ּ�, b.����) �ּ�, a.�Ƿ���" & vbNewLine & _
        "From (Select Nvl(a.���ۼ�, 0) ԭ��, b.�ּ�, Nvl(c.�Ƿ���, 0) �Ƿ���" & vbNewLine & _
        "       From ҩƷ�շ���¼ A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, ������ü�¼ D" & vbNewLine & _
        "       Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.����id = d.Id And" & vbNewLine & _
        "             (Sysdate Between b.ִ������ And b.��ֹ���� Or Sysdate >= b.ִ������ And b.��ֹ���� Is Null) And a.Id = [1]" & _
        GetPriceClassString("B") & ") A, ҩƷ���ľ��� B" & vbNewLine & _
        "Where b.���� = 0 And b.��� = 1 And b.���� = 2 And b.��λ = 1" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select Round(a.���ۼ�, Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))) ԭ��," & vbNewLine & _
        "       Round(b.�ּ�, Zl_To_Number(Nvl(zl_GetSysParameter(157), '5'))) �ּ�, Nvl(c.�Ƿ���, 0) �Ƿ���" & vbNewLine & _
        "From ҩƷ�շ���¼ A, �շѼ�Ŀ B, �շ���ĿĿ¼ C, סԺ���ü�¼ D" & vbNewLine & _
        "Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And a.����id = d.Id And" & vbNewLine & _
        "      (Sysdate Between b.ִ������ And b.��ֹ���� Or Sysdate >= b.ִ������ And b.��ֹ���� Is Null) And a.Id = [1] " & GetPriceClassString("B")
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "[ȡԭʼ�۸�����¼۸�]", lngBillId)
    
    If rstemp.RecordCount = 0 Then
        strMsg = "ҩƷ���ľ��ȱ������ݣ�"
        CheckPrice = True
        Exit Function
    End If
    
    'ʱ��ҩƷ������
    If rstemp!�Ƿ��� = 1 Then
        CheckPrice = True
        Exit Function
    End If
    
    '�Ƚϼ۸�
    If rstemp!ԭ�� <> rstemp!�ּ� Then
        strMsg = "ԭ��Ϊ" & rstemp!ԭ�� & ",�ּ�Ϊ" & rstemp!�ּ� & "��" & vbCrLf & Space(4) & "��ҩ������������ҩ��ϸ��¼���Ƿ������ҩ? "
        CheckPrice = False
        Exit Function
    End If
    
    CheckPrice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'ȡϵͳ����ֵ
Public Sub GetSysParms()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    
    gtype_UserSysParms.P6_δ��˼��ʴ�����ҩ = zldatabase.GetPara(6, glngSys, , 0)
    gtype_UserSysParms.P9_���ý���λ�� = zldatabase.GetPara(9, glngSys, , 0)
    gtype_UserSysParms.P15_�����շ��뷢ҩ���� = zldatabase.GetPara(15, glngSys, , 0)
    gtype_UserSysParms.P16_סԺ�����뷢ҩ���� = zldatabase.GetPara(16, glngSys, , 0)
    gtype_UserSysParms.P23_�ѽ��ʵ��ݲ��� = zldatabase.GetPara(23, glngSys, , 0)
    gtype_UserSysParms.P25_ʹ�õ���ǩ�� = zldatabase.GetPara(25, glngSys, , 0)
    gtype_UserSysParms.P26_����ǩ������ = zldatabase.GetPara(26, glngSys, , 0)
    gtype_UserSysParms.P28_���ﲡ������ʱ��Ҫˢ����֤ = zldatabase.GetPara(28, glngSys, , "1|0")
    gtype_UserSysParms.P29_ָ�������۶��۵�λ = zldatabase.GetPara(29, glngSys, , 0)
    gtype_UserSysParms.P44_����ƥ�� = zldatabase.GetPara(44, glngSys, , 0)
    gtype_UserSysParms.P54_ʱ��ҩƷ�ԼӼ������ = zldatabase.GetPara(54, glngSys, , 0)
    gtype_UserSysParms.P64_������� = zldatabase.GetPara(64, glngSys, , 0)
    gtype_UserSysParms.P68_����ҩ�������Ϻ���ҩ = zldatabase.GetPara(68, glngSys, , 0)
    gtype_UserSysParms.P70_�����Ǽ���Ч���� = zldatabase.GetPara(70, glngSys, , 0)
    gtype_UserSysParms.P75_�⹺�����Ҫ�˲� = zldatabase.GetPara(75, glngSys, , 0)
    gtype_UserSysParms.P76_ʱ��ҩƷֱ��ȷ���ۼ� = zldatabase.GetPara(76, glngSys, , 0)
    gtype_UserSysParms.P85_ҩ���鿴���ݳɱ��� = zldatabase.GetPara(85, glngSys, , 0)
    gtype_UserSysParms.P96_ҩƷ��¿��ÿ�� = zldatabase.GetPara(96, glngSys, , 0)
    gtype_UserSysParms.P98_���ʱ����������۷��� = zldatabase.GetPara(98, glngSys, , 0)
    gtype_UserSysParms.P126_ʱ��ҩƷ�ۼۼӳɷ�ʽ = zldatabase.GetPara(126, glngSys, , 0)
    gtype_UserSysParms.P148_δ�շѴ�����ҩ = zldatabase.GetPara(148, glngSys, , 0)
    gtype_UserSysParms.P149_Ч����ʾ��ʽ = zldatabase.GetPara(149, glngSys, , 0)
    gtype_UserSysParms.P150_ҩƷ���������㷨 = zldatabase.GetPara(150, glngSys, , 0)
    gtype_UserSysParms.P153_�������� = zldatabase.GetPara(153, glngSys, , 0)
    gtype_UserSysParms.P163_��Ŀִ��ǰ�������շѻ��ȼ������ = zldatabase.GetPara(163, glngSys, , 0)
    gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = zldatabase.GetPara(174, glngSys, , 0)
    gtype_UserSysParms.P175_ҩƷ������ȷ���� = zldatabase.GetPara(175, glngSys, , 0)
    gtype_UserSysParms.P214_�״�ҽ��ִ����Ҫ��� = zldatabase.GetPara(214, glngSys, , 0)
    gtype_UserSysParms.P221_ҩƷ���ʱ�� = zldatabase.GetPara(221, glngSys, , 0)
    gtype_UserSysParms.P222_ҩ���Զ�����ҩ�ӿ� = zldatabase.GetPara(222, glngSys, , 0)
    gtype_UserSysParms.P240_ҩ��������� = zldatabase.GetPara(241, glngSys, , 0)
    gtype_UserSysParms.P241_�������ʱ�� = zldatabase.GetPara(242, glngSys, , 0)
    gtype_UserSysParms.Para_���뷽ʽ = zldatabase.GetPara(44, glngSys, , 11)
    gtype_UserSysParms.P275_���۹���ģʽ = Val(zldatabase.GetPara(275, glngSys, , 0))
    gtype_UserSysParms.P213_��ҩ�䷽ÿ����ҩζ�� = Val(zldatabase.GetPara(213, glngSys, , 0))
    
    'ȡҩƷ���������
    gstrSQL = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum < 1"
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ����")
    gtype_UserDrugDigits.Digit_��� = rs.Fields(0).NumericScale
    gtype_UserDrugDigits.Digit_�ɱ��� = rs.Fields(1).NumericScale
    gtype_UserDrugDigits.Digit_���ۼ� = rs.Fields(2).NumericScale
    gtype_UserDrugDigits.Digit_���� = rs.Fields(3).NumericScale
    
    'ȡҩƷ�ۼ۵�λС��λ��
    gstrSQL = "Select ����, Nvl(����, 0) ���� From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ��λ = 1 "
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ�ۼ۵�λС��λ��")
    
    If rs.RecordCount > 0 Then
        rs.Filter = "����=1"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_�ɱ��� = rs!����
        
        rs.Filter = "����=2"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_���ۼ� = rs!����
        
        rs.Filter = "����=3"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_���� = rs!����
        
        If gtype_UserSaleDigits.Digit_�ɱ��� < 2 Or gtype_UserSaleDigits.Digit_�ɱ��� > gtype_UserDrugDigits.Digit_�ɱ��� Then
            gtype_UserSaleDigits.Digit_�ɱ��� = gtype_UserDrugDigits.Digit_�ɱ���
        End If
        
        If gtype_UserSaleDigits.Digit_���ۼ� < 2 Or gtype_UserSaleDigits.Digit_���ۼ� > gtype_UserDrugDigits.Digit_���ۼ� Then
            gtype_UserSaleDigits.Digit_���ۼ� = gtype_UserDrugDigits.Digit_���ۼ�
        End If
        
        If gtype_UserSaleDigits.Digit_���� < 2 Or gtype_UserSaleDigits.Digit_���� > gtype_UserDrugDigits.Digit_���� Then
            gtype_UserSaleDigits.Digit_���� = gtype_UserDrugDigits.Digit_����
        End If
    End If
    
    '����ȫ�ֲ���
    gstrLike = IIf(Val(zldatabase.GetPara("����ƥ��")) = 0, "%", "")
    gblnMyStyle = zldatabase.GetPara("ʹ�ø��Ի����") = "1"
        
    'ҩƷ������ʾ��ʽ
    gintҩƷ������ʾ = Val(zldatabase.GetPara("ҩƷ������ʾ", , , 2))
    gint����ҩƷ��ʾ = Val(zldatabase.GetPara("����ҩƷ��ʾ"))
    
    If gintҩƷ������ʾ < 0 Or gintҩƷ������ʾ > 2 Then gintҩƷ������ʾ = 2
    If gint����ҩƷ��ʾ < 0 Or gint����ҩƷ��ʾ > 1 Then gint����ҩƷ��ʾ = 0
    
    '���뷽ʽ
    gint���뷽ʽ = Val(zldatabase.GetPara("���뷽ʽ"))
    If gint���뷽ʽ < 0 Or gint���뷽ʽ > 1 Then gint���뷽ʽ = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function EsignIsOpen(ByVal lng����ID As Long) As Boolean
    Dim rstemp As Recordset
    
    On Error GoTo errH
    gstrSQL = "select Zl_Fun_Getsignpar(5,[1]) �Ƿ����� from dual"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "����ǩ��ʹ�ò���", lng����ID)

    If Not rstemp.EOF Then
        EsignIsOpen = (rstemp!�Ƿ����� = 1)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer, Optional ByVal lng����ID As Long) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String, intType As Integer
    Dim curDate As Date
    Dim intYear As Integer
    Dim PreFixNO As String  '���ǰ׺
    Dim strPre As String    '���������ǰ2λ
    Dim str��� As String
    Dim dateCurDate As Date
    Dim intMonth As Integer
    Dim strMonth As String
    
    On Error GoTo errH
    
    dateCurDate = Sys.Currentdate
    intYear = Format(dateCurDate, "YYYY") - 1990
    PreFixNO = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
    intMonth = Month(dateCurDate)
    strMonth = intMonth
    strMonth = String(2 - Len(strMonth), "0") & strMonth
    
    If Len(strNo) >= 8 Then
        GetFullNO = Right(strNo, 8)
        Exit Function
    ElseIf Len(strNo) = 7 Then
        GetFullNO = PreFixNO & strNo
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNo, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNo
    
    strsql = "Select ��Ź���,������,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, "GetFullNO", intNum)
        
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
        strPre = Left(Nvl(rsTmp!������, PreFixNO & "0"), 2)
    End If
    
    If intType = 0 Then
        '������
        GetFullNO = strPre & Format(Right(strNo, 6), "000000")
    ElseIf intType = 1 Then
        '���ձ��
        strsql = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strsql & Format(Right(strNo, 4), "0000")
    ElseIf intType = 2 Then
        '�����ҷ��±���
        gstrSQL = "Select ��� From ���Һ���� Where ��Ŀ���=[1] And Nvl(����ID,0)=[2]"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "GetFullNO", intNum, lng����ID)
        
        If rsTmp.RecordCount = 0 Then
            MsgBox "��δ���ÿ��ұ�ţ��޷��������룡", vbInformation, gstrSysName
            Exit Function
        End If
        If Nvl(rsTmp!���) = "" Then
            MsgBox "��δ���ÿ��ұ�ţ��޷��������룡", vbInformation, gstrSysName
            Exit Function
        End If
        str��� = Nvl(rsTmp!���)
        
        'С����λ�������²�������
        '��λ����λ������Ϊ��ָ���·ݵĺ���
        '��λ������Ϊ�ǲ�������ָ�����ҡ��·ݵĺ���
        '���ڵ��ڰ�λ��������
        If Len(strNo) <= 4 Then
            GetFullNO = PreFixNO & str��� & strMonth & String(4 - Len(strNo), "0") & strNo
        ElseIf Len(strNo) <= 6 Then
            GetFullNO = String(6 - Len(strNo), "0") & GetFullNO
            GetFullNO = PreFixNO & str��� & GetFullNO
        ElseIf Len(strNo) = 7 Then
            GetFullNO = PreFixNO & GetFullNO
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function


Public Function CopyNewRec(ByVal SourceRec As ADODB.Recordset) As ADODB.Recordset
        Dim RecTarget As New ADODB.Recordset
        Dim IntFields As Integer, LngLocate As Long
        '������:����
        '��������:2000-11-02
        '�ü�¼����ƾ֤�ؼ���Ӧ
        'Ҳʹ���ڱ���
        
        LngLocate = -1
        Set RecTarget = New ADODB.Recordset
        With RecTarget
                If .State = 1 Then .Close
                If SourceRec.RecordCount <> 0 Then
                        On Error Resume Next
                        err = 0
                        LngLocate = SourceRec.AbsolutePosition
                        If err <> 0 Then LngLocate = -1
                        SourceRec.MoveFirst
                End If
                For IntFields = 0 To SourceRec.Fields.count - 1
                        .Fields.Append SourceRec.Fields(IntFields).Name, SourceRec.Fields(IntFields).Type, SourceRec.Fields(IntFields).DefinedSize, adFldIsNullable     '0:��ʾ����
                Next
                
                .CursorLocation = adUseClient
                .CursorType = adOpenStatic
                .LockType = adLockOptimistic
                .Open
                
                If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
                Do While Not SourceRec.EOF
                        .AddNew
                        For IntFields = 0 To SourceRec.Fields.count - 1
                                .Fields(IntFields) = SourceRec.Fields(IntFields).Value
                        Next
                        .Update
                        SourceRec.MoveNext
                Loop
        End With
        
        If SourceRec.RecordCount <> 0 Then SourceRec.MoveFirst
        If LngLocate > 0 Then SourceRec.Move LngLocate - 1
        Set CopyNewRec = RecTarget
End Function




Public Function GetUserInfo() As Boolean
    Dim rsUser As ADODB.Recordset
    
    Set rsUser = Sys.GetUserInfo
    
    With rsUser
        If Not .EOF Then
            glngUserId = !Id '��ǰ�û�id
            UserInfo.�û�ID = !Id
            gstrUserCode = !��� '��ǰ�û�����
            UserInfo.�û����� = !���
            gstrUserName = IIf(IsNull(!����), "", !����) '��ǰ�û�����
            UserInfo.�û����� = IIf(IsNull(!����), "", !����)
            gstrUserAbbr = IIf(IsNull(!����), "", !����) '��ǰ�û�����
            UserInfo.�û����� = IIf(IsNull(!����), "", !����)
            glngDeptId = !����ID '��ǰ�û�����id
            UserInfo.����ID = !����ID
            gstrDeptCode = !������ '��ǰ�û�
            UserInfo.���ű��� = !������
            gstrDeptName = !������ '��ǰ�û�
            UserInfo.�������� = !������
            GetUserInfo = True
        Else
            glngUserId = 0 '��ǰ�û�id
            gstrUserCode = "" '��ǰ�û�����
            gstrUserName = "" '��ǰ�û�����
            gstrUserAbbr = "" '��ǰ�û�����
            glngDeptId = 0 '��ǰ�û�����id
            gstrDeptCode = "" '��ǰ�û�
            gstrDeptName = "" '��ǰ�û�
            
            
            UserInfo.�û�ID = 0
            UserInfo.�û����� = ""
            UserInfo.�û����� = ""
            UserInfo.�û����� = ""
            UserInfo.����ID = 0
            UserInfo.���ű��� = ""
            UserInfo.�������� = ""
        End If
    End With
End Function
Public Function GetUnit(ByVal lngҩ��id As Long, ByVal Int���� As Integer, ByVal strNo As String, ByVal int�����־ As Integer) As String
    '����ָ���ⷿ�����ݡ�NO���õ�ҩƷ��λ
    Dim intUnit As Integer
    Dim blnMoved As Boolean
    Dim rstemp As New ADODB.Recordset
    
    '����ϵͳ�����趨�ĵ�λ��ʾ����
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341, 0))
    If intUnit = 0 Then
        intUnit = int�����־
    End If
    If intUnit = 1 Or intUnit = 4 Then
        GetUnit = GetSpecUnit(lngҩ��id, gint����ҩ��)
    Else
        GetUnit = GetSpecUnit(lngҩ��id, gintסԺҩ��)
    End If
End Function
Public Function GetSpecUnit(ByVal lng�ⷿid As Long, ByVal int��Χ As Integer) As String
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    Dim strsql As String
    
    '����ָ���ָⷿ�����÷�Χ�ĵ�λ
    On Error GoTo ErrHand
    
    gstrSQL = "Select Nvl(����,1) AS ��λ From ҩƷ�ⷿ��λ Where �ⷿID=[1] And ���÷�Χ=[2] "
    Set rsProperty = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ", lng�ⷿid, int��Χ)
   
    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!��λ
    Else
        gstrSQL = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
        Set rsProperty = zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��λ", lng�ⷿid)
    
        'ȡ������󼰲�������
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            'סԺ��λ
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '���ﵥλ
            strUnit = 2
        ElseIf InStr(strWorkTemp, "ҩ��") <> 0 Then
            'ҩ�ⵥλ
            strUnit = 4
        Else
            '�ۼ۵�λ����Ҫ���Ƽ���
            strUnit = 1
        End If
    End If
    
    'ת��Ϊ��ʵ�ĵ�λ���ظ�������
    GetSpecUnit = Switch(strUnit = 1, "�ۼ۵�λ", strUnit = 2, "���ﵥλ", strUnit = 3, "סԺ��λ", strUnit = 4, "ҩ�ⵥλ")
    If glngSys / 100 = 8 Then
        'ҩ��ֻ���ۼ۵�λ��ҩ�ⵥλ
        GetSpecUnit = IIf(strUnit = 1, "�ۼ۵�λ", "ҩ�ⵥλ")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

'ȡҩƷ��λ����
Public Function GetDrugUnit(ByVal lng�ⷿid As Long, ByVal frmCaption As String, Optional ByVal bln���� As Boolean = True) As String
    Dim rsProperty As New Recordset
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim intUnit As Integer, strUnit As String
    Dim blnȱʡ As Boolean
    Dim lngModul As Long
    
    On Error GoTo ErrHand
    
    If frmCaption Like "ҩƷ�������*" Then
        lngModul = 1343
    ElseIf frmCaption Like "Э��ҩƷ���*" Then
        lngModul = 1344
    ElseIf frmCaption Like "ҩƷ�ƿ����*" Then
        lngModul = 1304
    End If
    
    intUnit = 0
    '��������쵥����ֱ�ӷ���ע����еĵ�λ
    If lngModul = 1343 Or lngModul = 1304 Or lngModul = 1344 Then
        intUnit = Val(zldatabase.GetPara("ҩƷ��λ", glngSys, lngModul))
        '���ز������õĵ�λ˳�����£�0-ȱʡ;1-ҩ��;2-����;3-סԺ;4-�ۼۣ���Ҫת��Ϊ��ϵͳ������һ��
        If intUnit = 1 Then
            intUnit = 4
        ElseIf intUnit = 4 Then
            intUnit = 1
        End If
        strUnit = intUnit
    End If
    
    If intUnit = 0 Then
        gstrSQL = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
        Set rsProperty = zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��λ", lng�ⷿid)
        
        'ȡ������󼰲�������
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        
        If InStr(strWorkTemp, "ҩ��") <> 0 Then
            'ҩ�ⵥλ
            intUnit = 1
            strUnit = 4
        ElseIf InStr(strobjTemp, "1") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '���ﵥλ
            intUnit = 2
            strUnit = 2
        ElseIf InStr(strobjTemp, "2") <> 0 Then
            'סԺ��λ
            intUnit = 3
            strUnit = 3
        Else
            '�ۼ۵�λ����Ҫ���Ƽ���
            intUnit = 4
            strUnit = 1
        End If
        
        'ȡ��ҩ��ȱʡ��ʹ�õĵ�λ
        GetDrugUnit = GetSpecUnit(lng�ⷿid, intUnit)
    Else
        GetDrugUnit = Switch(strUnit = 1, "�ۼ۵�λ", strUnit = 2, "���ﵥλ", strUnit = 3, "סԺ��λ", strUnit = 4, "ҩ�ⵥλ")
    End If
    
    'ת��Ϊ��ʵ�ĵ�λ���ظ�������
    
    If glngSys / 100 = 8 Then
        'ҩ��ֻ���ۼ۵�λ��ҩ�ⵥλ
        GetDrugUnit = IIf(strUnit = 1, "�ۼ۵�λ", "ҩ�ⵥλ")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDrugUnit = "�ۼ۵�λ"
End Function



'�����룬���ƣ���������ĳһ��
Public Function FindRow(ByVal mshBill As BillEdit, ByVal int�Ƚ��� As Integer, _
    ByVal str�Ƚ�ֵ As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim StrCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo errHandle
    FindRow = True
    With mshBill
        If .rows = 2 Then Exit Function
        If str�Ƚ�ֵ = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                StrCode = .TextMatrix(intRow, int�Ƚ���)
                If InStr(1, UCase(StrCode), UCase(str�Ƚ�ֵ)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int�Ƚ���
                    .SetRowColor CLng(intRow), &HFFCECE, True
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = " SELECT DISTINCT b.���� " & _
                  " FROM " & _
                  "    (SELECT DISTINCT A.�շ�ϸĿid " & _
                  "    FROM �շ���Ŀ���� A" & _
                  "    Where A.���� LIKE [1]) a," & _
                  " �շ���ĿĿ¼ B " & _
                  " Where a.�շ�ϸĿid = b.ID"
        Set rsCode = zldatabase.OpenSQLRecord(gstrSQL, "����ҩƷ", IIf(gstrMatchMethod = "0", "%", "") & str�Ƚ�ֵ & "%")
        
        If rsCode.EOF Then
            FindRow = False
            Exit Function
        End If
        
        For intRow = intStartRow To .rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                StrCode = .TextMatrix(intRow, int�Ƚ���)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(StrCode), UCase(rsCode!����)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int�Ƚ���
                        .SetRowColor CLng(intRow), &HFFCECE, True
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    FindRow = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetLength(ByVal strTable As String, ByVal strColumn As String) As Integer
    Dim rsPar As New ADODB.Recordset
    '��ȡָ�����ض��ֶεĳ���
    
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    GetLength = 0
    
    With rsPar
        gstrSQL = "Select " & strColumn & " From " & strTable & " Where Rownum<1"
        Call zldatabase.OpenRecordset(rsPar, gstrSQL, "��ȡ����")
        
        If err <> 0 Then
            MsgBox "���ݱ�[" & strTable & "]�����ڣ����뿪������ϵ��", vbInformation, gstrSysName
        End If
        GetLength = .Fields(0).DefinedSize
        .Close
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReturnSQL(ByVal lng�ⷿid As Long, ByVal strCaption As String, _
    Optional ByVal bln���� As Boolean = True, _
    Optional ByVal lngModuleNO As Long = 0) As ADODB.Recordset
    
    Dim str�ⷿ���� As String, strҩƷ���� As String, strվ������ As String, strsql As String
    '����ҩƷ������Ʊ�����ݣ���ȡ�Է��ⷿ
    'Writed by zyb
    '-----------------����-----------------
    '���ڿⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� In (1"������Է��ⷿ",3"��˫����ͨ")
    '�Է��ⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� IN (2"���������ڿⷿ",3"��˫����ͨ")
    '-----------------����-----------------
    '���ڿⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� In (2"���������ڿⷿ",3"��˫����ͨ")
    '�Է��ⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� IN (1"������Է��ⷿ",3"��˫����ͨ")
    
    On Error GoTo errHandle
    strվ������ = GetDeptStationNode(lng�ⷿid)
    str�ⷿ���� = "('H','I','J','K','L','M','N')"
    
    strҩƷ���� = ",(Select �Է��ⷿID ID From ҩƷ�������" & _
            " Where ���ڿⷿID=[1] And ���� In (" & IIf(bln����, 1, 2) & ",3)" & _
            " Union" & _
            " Select ���ڿⷿID ID From ҩƷ�������" & _
            " Where �Է��ⷿID=[1] And ���� In (" & IIf(bln����, 2, 1) & ",3)) D"
    Select Case lngModuleNO
        Case 1343   'ҩƷ�������
            strsql = " SELECT DISTINCT a.id,a.����,a.����, Decode(Instr(',H,I,J,', ',' || b.���� || ','), 0, 0, 1) As ҩ������ " & _
                    " FROM ��������˵�� c, �������ʷ��� b, ���ű� a" & strҩƷ���� & _
                    " Where c.�������� = b.����" & _
                    " AND b.����||'' in " & str�ⷿ���� & _
                    " AND a.id = c.����id And A.ID=D.ID " & _
                    " AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
                    " Order by a.����"
        Case Else
            strsql = " SELECT DISTINCT a.id,a.����,a.����, Decode(Instr(',H,I,J,', ',' || b.���� || ','), 0, 0, 1) As ҩ������ " & _
                    " FROM ��������˵�� c, �������ʷ��� b, ���ű� a" & strҩƷ���� & _
                    " Where c.�������� = b.����" & _
                    " AND b.����||'' in " & str�ⷿ���� & _
                    " AND a.id = c.����id And A.ID=D.ID" & IIf(strվ������ <> "", " AND a.վ��=[2] ", "") & _
                    " AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
                    " Order by a.����"
    End Select
    Set ReturnSQL = zldatabase.OpenSQLRecord(strsql, strCaption, lng�ⷿid, strվ������)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckRepeatMedicine(ByVal MyBill As Object, ByVal strDrugInfo As String, ByVal intExceptRow As Integer) As Boolean
    'ҩƷ��ͨ�༭������¼���ҩƷ�Ƿ��ظ�
    'MyBill�����ؼ���ҩƷ�б�
    'strDrugInfo��ҩƷID�����μ���Ӧ�кţ���ʽ��ҩƷID,ҩƷID�к�|����,�����кţ�
    'intExceptRow���ų�ָ�����У��������һ�У�
    Dim n As Integer
    Dim lngҩƷid As Long
    Dim intҩƷID�к� As Integer
    Dim lng���� As Long
    Dim int�����к� As Integer
    
    On Error GoTo errHandle
    lngҩƷid = Val(Split(Split(strDrugInfo, "|")(0), ",")(0))
    intҩƷID�к� = Val(Split(Split(strDrugInfo, "|")(0), ",")(1))
    lng���� = Val(Split(Split(strDrugInfo, "|")(1), ",")(0))
    int�����к� = Val(Split(Split(strDrugInfo, "|")(1), ",")(1))
    
    With MyBill
        For n = 1 To .rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If n <> intExceptRow And Val(.TextMatrix(n, intҩƷID�к�)) = lngҩƷid And Val(.TextMatrix(n, int�����к�)) = lng���� Then
                    MsgBox "�Բ������и�ҩƷ���ҩƷ����ͬ���Σ������ظ����룡", vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    CheckRepeatMedicine = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function












Public Sub CheckLapse(ByVal strЧ�� As String)
    'ʧЧҩƷ���
    If Not IsDate(strЧ��) Then Exit Sub
    If Format(strЧ��, "yyyy-MM-dd") < Format(Sys.Currentdate, "yyyy-MM-dd") Then
        MsgBox "��ҩƷ�Ѿ�ʧЧ�ˣ�", vbInformation, gstrSysName
    End If
End Sub

Public Sub zlPlugIn_Ini(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object)
    '�����չ�ӿڳ�ʼ��
    If objPlugIn Is Nothing Then
        On Error Resume Next
        Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not objPlugIn Is Nothing Then
            Call objPlugIn.Initialize(gcnOracle, lngSys, lngModul)
            If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub


Public Sub zlPlugIn_SetMenu(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object, _
    cbcMain As CommandBarControls, ByVal lngMenuPlugInMain As Long)
    '������չ���ܵĲ˵���Ŀ
    '������lngSys-ϵͳ��lngModul-ģ��ţ�objPlugIn-��չ��Ҷ���cbcMain-CommandBar���˵�����lngMenuPlugInMain-��Ҳ˵�
    Dim strFunc As String, strFuncName As String '��¼��չ����
    Dim lngFuncID As Long
    Dim blnGroup As Boolean
    Dim intToolbarCount As Integer
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim i As Integer
    
    If Not objPlugIn Is Nothing Then
        On Error Resume Next
        
        '��Ҳ�������չ����
        strFunc = objPlugIn.GetFuncNames(lngSys, lngModul)
        If InStr(",438,0,", "," & err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� GetFuncNames ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
        End If
        
        err.Clear: On Error GoTo 0
    End If
    
    If strFunc = "" Then Exit Sub
    
    With cbcMain
        Set objPopup = .Add(xtpControlButtonPopup, lngMenuPlugInMain, "��չ(&E)")
        objPopup.BeginGroup = True
        
        With objPopup.CommandBar.Controls
            For i = 0 To UBound(Split(strFunc, ","))
                lngFuncID = lngMenuPlugInMain + i + 1
                strFuncName = Split(strFunc, ",")(i)
                
                blnGroup = InStr(strFuncName, "|") > 0
                strFuncName = Replace(strFuncName, "InTool:", "")
                strFuncName = Replace(strFuncName, "|:", "")
                
                Set cbrControl = .Add(xtpControlButton, lngFuncID, strFuncName)
                If i <= 9 Then cbrControl.Caption = cbrControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                cbrControl.IconId = lngMenuPlugInMain
                cbrControl.Parameter = strFuncName
                cbrControl.BeginGroup = blnGroup
            Next
        End With
    End With
End Sub

Public Sub zlPlugIn_SetToolbar(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object, _
    cbrToolBar As CommandBarControls, ByVal lngMenuPlugInMain As Long)
    '������չ���ܵĹ�������Ŀ
    '������lngSys-ϵͳ��lngModul-ģ��ţ�objPlugIn-��չ��Ҷ���cbrToolBar-CommandBar����������lngMenuPlugInMain-��Ҳ˵�
    Dim strFunc As String, strFuncName As String '��¼��չ����
    Dim lngFuncID As Long
    Dim blnGroup As Boolean
    Dim intToolbarCount As Integer
    Dim cbrControl As CommandBarControl
    Dim i As Integer
    
    If Not objPlugIn Is Nothing Then
        On Error Resume Next
        
        '��Ҳ�������չ����
        strFunc = objPlugIn.GetFuncNames(lngSys, lngModul)
        If InStr(",438,0,", "," & err.Number & ",") = 0 Then
            MsgBox "zlPlugIn ��Ҳ���ִ�� GetFuncNames ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
        End If
        
        err.Clear: On Error GoTo 0
    End If
    
    If strFunc = "" Then Exit Sub
    
    With cbrToolBar
        For i = 0 To UBound(Split(strFunc, ","))
            strFuncName = Split(strFunc, ",")(i)
            lngFuncID = lngMenuPlugInMain + i + 1
            
            If InStr(strFuncName, "InTool:") > 0 Then
                intToolbarCount = intToolbarCount + 1
                blnGroup = (intToolbarCount = 1 Or InStr(strFuncName, "|") > 0)
                strFuncName = Replace(strFuncName, "InTool:", "")
                strFuncName = Replace(strFuncName, "|:", "")
    
                Set cbrControl = .Add(xtpControlButton, lngFuncID, strFuncName)
                cbrControl.IconId = lngMenuPlugInMain
                cbrControl.Parameter = strFuncName
                cbrControl.BeginGroup = blnGroup
            End If
        Next
    End With
End Sub


Public Sub zlPlugIn_Unload(objPlugIn As Object)
    'ж����ҽӿ�
    Set objPlugIn = Nothing
End Sub


Public Function ���������(ByVal lng�ⷿid As Long, ByVal lngҩƷid As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim bln����Ƿ���� As Boolean, bln���� As Boolean, bln�ⷿ As Boolean
    'ͨ��ҩƷѡ��������ҩƷʱ�����ҩƷ����е�������Ӳ������ʡ�ҩƷĿ¼�еķ��������жϳ��Ĳ�һ�£��򱨴�
    On Error GoTo errHandle
    ��������� = False
    
    '���û�п���¼����ֱ���˳�
    gstrSQL = " Select Count(*) ��¼�� From ҩƷ��� " & _
              " Where �ⷿID=[1] And ����=1 And ҩƷID=[2]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "�Ƿ���ڿ������", lng�ⷿid, lngҩƷid)
    
    If rsCheck!��¼�� = 0 Then
        ��������� = True
        Exit Function
    End If
    
    '���ڷ�����¼���������
    gstrSQL = " Select Count(*) ���� From ҩƷ��� " & _
              " Where �ⷿID=[1] And ����=1 And Nvl(����,0)<>0 And ҩƷID=[2]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "���������", lng�ⷿid, lngҩƷid)
    
    bln����Ƿ���� = (rsCheck!���� <> 0)
    
    '���ж��Ƿ��ǿⷿ
    gstrSQL = "select ����ID from ��������˵�� where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng�ⷿid)
    
    bln�ⷿ = (rsCheck.EOF)
        
    '�ж϶�Ӧ��ҩƷĿ¼�еķ�������
    gstrSQL = " Select Nvl(ҩ�����,0) ��������,nvl(ҩ������,0) ҩ���������� " & _
              " From ҩƷ��� Where ҩƷID=[1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩƷĿ¼�еķ�������", lngҩƷid)
              
    If bln�ⷿ Then
        bln���� = (rsCheck!�������� = 1)
    Else
        bln���� = (rsCheck!ҩ���������� = 1)
    End If
    
    ��������� = (bln����Ƿ���� = bln����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




'ȡҩƷ���۸��������С��λ��
Public Function GetDigit(ByVal int���� As Integer, ByVal int��� As Integer, ByVal int���� As Integer, Optional ByVal int��λ As Integer) As Integer
    'int���ʣ�0-���㾫��;1-��ʾ����
    'int���1-ҩƷ;2-����
    'int���ݣ�1-�ɱ���;2-���ۼ�;3-����;4-���
    'int��λ�������ȡ���λ�������Բ�����ò���
    '         ҩƷ��λ:1-�ۼ�;2-����;3-סԺ;4-ҩ��;
    '         ���ĵ�λ:1-ɢװ;2-��װ
    '���أ���С2�����Ϊ���ݿ����С��λ��
    
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If int���� = 4 Then
        int��λ = 5
    End If
    gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ���ľ��� Where ���� = [1] And ��� = [2] And ���� = [3] And ��λ = [4] "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ" & Choose(int����, "�ɱ���", "���ۼ�", "����") & "С��λ��", int����, int���, int����, int��λ)
    
    If rsTmp.RecordCount > 0 Then
        GetDigit = rsTmp!����
    End If
    
    If GetDigit = 0 Then
        '���û�����þ��ȣ���ȡ���ݿ���������λ��
        GetDigit = Choose(int����, gtype_UserDrugDigits.Digit_�ɱ���, gtype_UserDrugDigits.Digit_���ۼ�, gtype_UserDrugDigits.Digit_����, gtype_UserDrugDigits.Digit_���)
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigit = Choose(int����, gtype_UserDrugDigits.Digit_�ɱ���, gtype_UserDrugDigits.Digit_���ۼ�, gtype_UserDrugDigits.Digit_����, gtype_UserDrugDigits.Digit_���)
End Function

'���ݿⷿ�İ�װ��λ��ȡҩƷ�ļ۸����������С��λ��
Public Sub GetDrugDigit(ByRef lng�ⷿid As Long, ByVal frmCaption As String, ByRef intUnit As Integer, ByRef intCostDigit As Integer, ByRef intPricedigit As Integer, ByRef intNumberDigit As Integer, ByRef intMoneyDigit As Integer)
    Dim strUnit As String
    
    Const conInt���㾫�� As Integer = 0
    
    Const conIntҩƷ As Integer = 1
    
    Const conint�ۼ۵�λ As Integer = 1
    Const conint���ﵥλ As Integer = 2
    Const conintסԺ��λ As Integer = 3
    Const conintҩ�ⵥλ As Integer = 4
        
    Const conInt�ɱ��� As Integer = 1
    Const conInt�ۼ� As Integer = 2
    Const conInt���� As Integer = 3
    Const conInt��� As Integer = 4
    
    strUnit = GetDrugUnit(lng�ⷿid, frmCaption)
    
    Select Case strUnit
        Case "�ۼ۵�λ"             '�ۼ۵�λ����Ҫ���Ƽ���
            intUnit = conint�ۼ۵�λ
        Case "���ﵥλ"
            intUnit = conint���ﵥλ
        Case "סԺ��λ"
            intUnit = conintסԺ��λ
        Case "ҩ�ⵥλ"
            intUnit = conintҩ�ⵥλ
    End Select

    '�ֱ�ȡҩƷ�ɱ��ۡ��ۼۡ�����������С��λ��
    intCostDigit = GetDigit(conInt���㾫��, conIntҩƷ, conInt�ɱ���, intUnit)
    intPricedigit = GetDigit(conInt���㾫��, conIntҩƷ, conInt�ۼ�, intUnit)
    intNumberDigit = GetDigit(conInt���㾫��, conIntҩƷ, conInt����, intUnit)
    intMoneyDigit = GetDigit(conInt���㾫��, conIntҩƷ, conInt���)

End Sub



Public Function ҩƷ�������(ByVal str������ As String) As Boolean
    'ҩƷ�������ʱ���Ƿ��ж�������������ˣ��䷵����˽��
    Dim blnBillVerify As Boolean
    Dim rsSystemPara As New Recordset
    Dim intTemp As Integer
    
    On Error GoTo errHandle
    
    ҩƷ������� = True
    
    intTemp = Val(zldatabase.GetPara(64, glngSys, 0))
    If intTemp = 0 Then
        blnBillVerify = False
        Exit Function
    Else
        blnBillVerify = True
    End If
    
    If Not blnBillVerify Then Exit Function
    
    ҩƷ������� = (Trim(str������) <> Trim(gstrUserName))
    If Not ҩƷ������� Then MsgBox "������������˲�����ͬһ�ˣ����飡", vbInformation, gstrSysName
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBillInfo(ByVal lng���� As Long, ByVal strNo As String, Optional ByVal bln�������� As Boolean = True) As String
    Dim rsBillInfo As New ADODB.Recordset
    '��ȡ���ݵ�����޸�ʱ��
    
    On Error GoTo errHandle
    gstrSQL = " Select to_char(Max(" & IIf(bln��������, "��������", "�������") & "),'yyyyMMddhh24miss') ���� From ҩƷ�շ���¼ " & _
            " Where ����=[1] And NO=[2]"
    Set rsBillInfo = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���ݵ�����޸�ʱ��", lng����, strNo)
    
    With rsBillInfo
        '���ؿգ���ʾ�Ѿ�ɾ��
        If .EOF Then Exit Function
        If IsNull(!����) Then Exit Function
        GetBillInfo = !����
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ��鵥��(ByVal lng���� As Long, ByVal strNo As String, Optional ByVal blnmsg As Boolean = True, Optional ByVal bln�ƿⵥ As Boolean = False) As Boolean
    Dim rsPrice As New ADODB.Recordset
    Dim lngҩƷ_Last As Long, lngҩƷ_Cur As Long
    Dim intPricedigit As Integer
    Dim intCostDigit As Integer
    '���ҩƷ�ļ۸��Ƿ�Ϊ���µļ۸񣨰�ҩ�ⵥλ���бȽϣ��������������
    '�����ڱ���ǰ�жϺ��鷳���Ҹ��ֵ��ݵı���б�������ݲ�һ������ˣ����������֮�����ύǰ���ѱ�������ݽ��м��
    'ҩƷ��ͬ�ļ�¼�Թ�
    
    '�Զ�������鲢ִ�е���
    On Error GoTo errHandle
    
    Call AutoAdjustPrice_ByNO(lng����, strNo)
    intPricedigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
    
    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid , 0 ԭ��, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, �շ���ĿĿ¼ C" & _
            " Where a.���� = [1] And a.No = [2] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & intPricedigit & ") <> Round(b.�ּ�, " & intPricedigit & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0 " & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid , 0 ԭ��, decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C ," & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 1 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = [1] And a.No = [2] And c.Id = a.ҩƷid And Round(a.���ۼ�," & intPricedigit & ") <> Round(decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�), " & intPricedigit & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
                  " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid , 0 ԭ��, decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B ," & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 2 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = [1] And a.No = [2] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & intCostDigit & ")<>round(decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�)," & intCostDigit & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1  and b.����=1" & _
            " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Order By ����, ҩƷid, ���"
    Set rsPrice = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ�۸�", lng����, strNo)
            
    If rsPrice.EOF Then
        ��鵥�� = True
        Exit Function
    End If
    
    lngҩƷ_Last = 0
    With rsPrice
        Do While Not .EOF
            lngҩƷ_Cur = !ҩƷID
            If lngҩƷ_Cur <> lngҩƷ_Last Then
                If blnmsg Then
                    MsgBox "��" & IIf(bln�ƿⵥ, Round(!��� / 2 + 0.49), !���) & "��ҩƷ��" & !���� & "�������¼۸񣬽��������¼۸���½��棡", vbInformation, gstrSysName
                    Exit Function
                Else
                    Exit Function
                End If
            End If
            
            lngҩƷ_Last = lngҩƷ_Cur
            .MoveNext
        Loop
        ��鵥�� = True
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DepotProperty(ByVal lng��Աid As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '����ָ����Ա�Ƿ����ҩ������
    On Error GoTo errHandle
    gstrSQL = "Select Distinct �������� From ������Ա B,��������˵�� A " & _
             " Where A.�������� like '%ҩ��' And " & _
             " A.����id = B.����id And B.��Աid = [1]"
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng��Աid)
    If rsCheck.RecordCount <> 0 Then
        DepotProperty = True
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowCostPrice() As Boolean
    'ҩ����Ա���ܣ�ֻ��ҩ����Ա���Բ�������Ϊ׼
    If DepotProperty(glngUserId) Then
        ShowCostPrice = True
    Else
        ShowCostPrice = (gtype_UserSysParms.P85_ҩ���鿴���ݳɱ��� = 1)
    End If
End Function
Public Function IsOwner(ByVal strUser As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    '�жϴ�����û��ǲ��������߻�DBA�û�
    On Error GoTo errHandle
    gstrSQL = "SELECT 1 FROM DUAL " & _
            " WHERE EXISTS(SELECT 1 FROM ZLSYSTEMS WHERE ������=[1])"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "�жϸ��û��ǲ���������", UCase(strUser))
    IsOwner = (rstemp.RecordCount <> 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Is��ҩ�ⷿ(ByVal lngҩ��id As Long) As Boolean
    Dim rstemp As New ADODB.Recordset
    Dim str�ⷿ���� As String
    
    gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", lngҩ��id)
    Do While Not rstemp.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rstemp!��������
        rstemp.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Then
        Is��ҩ�ⷿ = True
    Else
        Is��ҩ�ⷿ = False
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsLowerLimit(ByVal lng�ⷿid As Long, ByVal lngҩƷid As Long) As Boolean
    '�жϸ�ҩƷ�ڵ�ǰ���Ŀ���Ƿ���ڿ�����ޣ����򷵻���
    Dim dbl������� As Double, dbl���� As Double
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '��ȡ�������
    gstrSQL = " Select Sum(Nvl(ʵ������,0)) AS ������� From ҩƷ���" & _
              " Where ����=1 And �ⷿID=[1] And ҩƷID=[2]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡָ���ⷿ��ʵ�ʿ��", lng�ⷿid, lngҩƷid)
              
    If rstemp.RecordCount = 1 Then dbl������� = Nvl(rstemp!�������, 0)
    
    '��ȡ�����޶��е�����
    gstrSQL = " Select Nvl(����,0) AS ���� From ҩƷ�����޶�" & _
              " Where �ⷿID=[1] And ҩƷID=[2]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�����޶��е�����", lng�ⷿid, lngҩƷid)
    
    If rstemp.RecordCount = 1 Then dbl���� = rstemp!����
    
    IsLowerLimit = (dbl������� < dbl����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsReceiptBalance_Charge(ByVal intType As Integer, ByVal strȨ�� As String, ByVal lng���� As Long, ByVal strNo As String, ByVal str��� As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer, Optional ByVal lngModle As Long) As Boolean
    'intType    ��0-��ҩ;1-��ҩ
    'strȨ��    ����ǰ����Աӵ�е�Ȩ��
    'lng����    ����ǰ��������
    'strNO      ����ǰ���ݺ�
    'str���    ���������
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If lng���� = 8 Then
        IsReceiptBalance_Charge = True
        Exit Function
    End If
    
    '��ҩ����ҩ״̬�ֱ����Ƿ���Ȩ�ޡ����ѽ��ʴ������͡����ѽ��ʴ����������ô����Ƿ��ѽ��ʣ��ѽ��ʴ�����������ҩ����
    If (intType = 0 And InStr(1, strȨ��, "���ѽ��ʴ���") = 0) Or (intType = 1 And InStr(1, strȨ��, "���ѽ��ʴ���") = 0) Then
        '�ϲ����סԺ���ü�¼�������˽�������
        gstrSQL = "Select Nvl(Sum(Nvl(���ʽ��,0)),0) AS ���ʽ��   " & _
                 "  From ������ü�¼   " & _
                 "  Where Instr([1], ',' || ��� || ',') > 0 " & _
                 "  And Mod(��¼����,10) = 2 and NO = [2] "
        If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        End If
        gstrSQL = gstrSQL & " Order By ���ʽ�� Desc"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��ѽ���", "," & str��� & ",", strNo)
        
        If Nvl(rstemp!���ʽ��, 0) <> 0 Then
            If lngModle = 1 Then
                MsgBox "�����ѽ��ʣ���û�ж��ѽ��ʲ��˵���Һ������������˵�Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            ElseIf lngModle = 2 Then
                MsgBox "�����ѽ��ʣ���û�ж��ѽ��ʲ��˵���Һ�����а�ҩ��Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            ElseIf lngModle = 3 Then
                MsgBox "�����ѽ��ʣ���û�ж��ѽ��ʲ��˵���Һ������ȡ����ҩ��Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            ElseIf lngModle = 4 Then
                MsgBox "�����ѽ��ʣ���û�ж��ѽ��ʲ��˵���Һ������ȡ����ҩ��Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            Else
                MsgBox "�ڴ���[" & strNo & "]�����ѽ��ˣ���û�ж��ѽ��˲��˵Ĵ������з�ҩ����ҩ��Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End If
    
    IsReceiptBalance_Charge = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsOutPatient(ByVal strȨ�� As String, ByVal lng���� As Long, ByVal strNo As String, _
    ByVal int��¼���� As Integer, ByVal int�����־ As Integer, Optional ByVal lng����ID As Long, _
    Optional ByVal lng��ҳID As Long, Optional ByVal lngModle As Long, Optional ByVal str���� As String) As Boolean
    '����˵���������ǰ������סԺ���ˣ����û��Ȩ�ޡ����˳�Ժ���˴���������������ҩ����
    Const str���˳�Ժ���˴��� As String = "���˳�Ժ���˴���"
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    '�����Ȩ�ޡ����˳�Ժ���˴���������������ҩ����
    If InStr(1, strȨ��, str���˳�Ժ���˴���) > 0 Then
        IsOutPatient = True
        Exit Function
    End If
    
    '�����ǰ������סԺ���ˣ����û��Ȩ�ޡ����˳�Ժ���˴���������������ҩ����
    '���δ���벡��ID�����Զ���ȡ
    If lng����ID = 0 Then
'        gstrSQL = "Select A.����ID,c.��ҳid From ������ü�¼ A, ҩƷ�շ���¼ B,����ҽ����¼ C Where A.ID = B.����ID  And A.ҽ�����=C.id And b.���� = [1] And b.No = [2] And Rownum = 1 "
'
'        If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
'        Else
'            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
'        End If
        
        gstrSQL = "Select a.����id, c.��ҳid" & vbNewLine & _
                "From ������ü�¼ A, ҩƷ�շ���¼ B, ����ҽ����¼ C" & vbNewLine & _
                "Where a.Id = b.����id And a.ҽ����� = c.Id And b.���� = [1] And b.No = [2] And Rownum = 1" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select d.����id, d.��ҳid" & vbNewLine & _
                "From סԺ���ü�¼ D, ҩƷ�շ���¼ B" & vbNewLine & _
                "Where d.Id = b.����id And b.���� = [1] And b.No = [2] And Rownum = 1"

        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����, strNo)
        
        '����������Ҳ�������ID�򲻽�����һ�����
        If rstemp.EOF Then
            IsOutPatient = True
            Exit Function
        End If
        
        lng����ID = rstemp!����ID
        lng��ҳID = Nvl(rstemp!��ҳid, 0)
    End If
    
    'ȡ��������
    If str���� = "" Then
        gstrSQL = "Select ���� From ������Ϣ Where ����ID=[1]"
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng����ID)

        str���� = rstemp!����
    End If
    
    '��鲡����Ԥ��Ժ���Ժ
    gstrSQL = " Select 1 From ������ҳ" & _
              " Where ����ID=[1] and ��ҳid=[2] " & _
              " And (��Ժ���� Is Not NULL)"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��ѳ�Ժ", lng����ID, lng��ҳID)
    
    If rstemp.RecordCount <> 0 Then
        If lngModle = 1 Then
            MsgBox "���ˡ�" & str���� & "���ѳ�Ժ����û�ж��ѳ�Ժ���˵���Һ������������˵�Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
        ElseIf lngModle = 2 Then
            MsgBox "���ˡ�" & str���� & "���ѳ�Ժ����û�ж��ѳ�Ժ���˵���Һ�����а�ҩ��Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
        ElseIf lngModle = 3 Then
            MsgBox "���ˡ�" & str���� & "���ѳ�Ժ����û�ж��ѳ�Ժ���˵���Һ������ȡ����ҩ��Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
        ElseIf lngModle = 4 Then
            MsgBox "���ˡ�" & str���� & "���ѳ�Ժ����û�ж��ѳ�Ժ���˵���Һ������ȡ����ҩ��Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
        Else
            MsgBox "�ڴ���[" & strNo & "]�У����ˡ�" & str���� & "���ѳ�Ժ����û�ж��ѳ�Ժ���˵Ĵ������з�ҩ����ҩ��Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    IsOutPatient = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Calc_Clique(ByVal lngҩƷid As Long, ByVal dbl�������� As Double, Optional ByVal blnУ�� As Boolean = False) As Double
    Dim dblʵ������ As Double
    Dim dbl�� As Double, dbl�� As Double, dbl��ֵ As Double
    Dim rstemp As New ADODB.Recordset
    '�������췧ֵ����ó�ʵ��������������ҩƷ����ʱ������������������ܾ��ǿ����������ʱУ�����Ϊ�棬������Ľ�����ܴ�������������Ҳ���ǿ������
    '����������ȷ�ģ���϶�����У����Ӧ�������죩
'    On Error Resume Next

'    err = 0
    On Error GoTo errHandle
    Calc_Clique = dbl��������
    
    '��ȡ��ҩƷ�����췧ֵ��Ϊ����ֱ���˳�
    gstrSQL = "Select Nvl(���췧ֵ,0) AS ��ֵ From ҩƷ��� Where ҩƷID=[1]"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩƷ�����췧ֵ", lngҩƷid)

    If err <> 0 Then Exit Function
    If rstemp!��ֵ = 0 Then Exit Function
    dbl��ֵ = rstemp!��ֵ
    
    '�㷨(�����뷧ֵ��һ����бȽϣ��������룩�����С�ڣ�������������λ
    dbl�� = Int(dbl�������� / dbl��ֵ)
    dbl�� = dbl�������� - (dbl��ֵ * dbl��)
    If dbl�� >= (dbl��ֵ / 2) And Not blnУ�� Then
        dbl�� = dbl�� + 1
    End If
    
    dblʵ������ = dbl�� * dbl��ֵ
    Calc_Clique = dblʵ������
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Logogram(ByVal staVal As StatusBar, ByVal bytType As Byte)
'���뷽ʽ
'staVal: StartusBar�ؼ�
'bytType: 0=ƴ��; 1=���;  ��ǰ����״̬
    Dim i As Integer
    For i = 1 To staVal.Panels.count
        If staVal.Panels(i).Key = "PY" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrInset
                zldatabase.SetPara "���뷽ʽ", 0
                gint���뷽ʽ = 0
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrRaised
            End If
        ElseIf staVal.Panels(i).Key = "WB" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrRaised
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrInset
                zldatabase.SetPara "���뷽ʽ", 1
                gint���뷽ʽ = 1
            End If
        End If
    Next
End Sub

Public Function GetDeptStationNode(ByVal lngDeptId As Long) As String
'��ȡ��������վ����Ϣ
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    strTmp = "select վ�� from ���ű� where id=[1]"
    Set rsSQL = zldatabase.OpenSQLRecord(strTmp, "��ȡ��������վ����Ϣ", lngDeptId)
    If Not rsSQL.EOF Then
        GetDeptStationNode = Nvl(rsSQL!վ��)
    End If
    rsSQL.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '����:�������ƶ��ؼ���:2008-07-08 16:48:35
    err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    objCtl.SetFocus
End Sub

Public Function Select����ѡ����(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str�������� As String = "", _
    Optional bln����Ա As Boolean = False, _
    Optional ByVal str������� As String, _
    Optional strsql As String = "") As Boolean
    '------------------------------------------------------------------------------
    '����:����ѡ����
    '����:objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '     str��������-��������:��"V,W,K"
    '     bln����Ա-�Ƿ�Ӳ���Ա����
    '     strSQL-ֱ�Ӹ���SQL��ȡ����(�����ű�ı���һ��Ҫ��A)
    '����:�ɹ�,����true,���򷵻�False
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rstemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    On Error GoTo errHandle
    strTittle = "����ѡ����"
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    If strsql <> "" Then
    
        gstrSQL = strsql
    Else
        gstrSQL = "" & _
        "   Select /*+ Rule*/ distinct a.Id,a.�ϼ�id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
        "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��"
    
        If str�������� = "" And bln����Ա = False Then
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a" & _
            "   Where 1=1"
        Else
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a, �������ʷ��� b,��������˵�� c," & _
            IIf(str�������� = "", "", "       (Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.�������� = b.����" & IIf(str�������� = "", "(+)", " and B.����=J.column_value ") & _
            "         AND a.id = c.����id and c.������� in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))" & _
            IIf(bln����Ա = False, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') or a.����ʱ�� is null ) And (a.վ��=[5] or a.վ�� is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.���� like upper([3]) or a.���� like upper([3]) or a.���� like [3] )"
        If IsNumeric(strSearch) Then                         '���������,��ֻȡ����
            If Mid(gtype_UserSysParms.Para_���뷽ʽ, 1, 1) = "1" Then strFind = " And (A.���� Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            '.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ" ))
            If Mid(gtype_UserSysParms.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And  (a.���� Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  'ȫ����
            strFind = " And a.���� Like [3] "
        End If
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strsql = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.�ϼ�id Is Null Connect By Prior A.ID = A.�ϼ�id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.����"
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strsql = "" Then
        '�����¼�
        Set rstemp = zldatabase.ShowSQLSelect(FrmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, str�������)
    Else
        Set rstemp = zldatabase.ShowSQLSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, glngUserId, str��������, strKey, str�������, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    If rstemp Is Nothing Then
        MsgBox "û�����������Ĳ���,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    Call zlCtlSetFocus(objCtl, True)
    If UCase(TypeName(objCtl)) = UCase("ComboBox") Then
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rstemp!Id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            MsgBox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = Nvl(rstemp!����) & "-" & Nvl(rstemp!����)
        objCtl.Tag = Val(rstemp!Id)
    End If
    OS.PressKey vbKeyTab
    Select����ѡ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Sub SetTip(ByVal objControl As Object, ByVal strTip As String)
    '����:������ʾָ���Ŀؼ�����ʾ�ı�
    'objControl:��Ҫ��ʾ�Ŀؼ�
    'strTip:����֯�õ���ʾ�ı�
    Call zlCommFun.ShowTipInfo(objControl.hWnd, strTip, True, True, 8800)
End Sub

Public Sub SetSelectorRS( _
    ByVal byt�༭ģʽ As Byte, _
    ByVal strModeName As String, _
    Optional ByVal lng��Դ�ⷿ As Long = 0, _
    Optional ByVal lngĿ��ⷿ As Long = 0, _
    Optional ByVal lngʹ�ò��� As Long = 0, _
    Optional ByVal lng��Ӧ�� As Long = 0, _
    Optional ByVal byt���÷�ʽ As Byte = 0, _
    Optional ByVal bln����ͣ��ҩƷ As Boolean = False, _
    Optional ByVal bln���޴洢�ⷿҩƷ As Boolean = False, _
    Optional ByVal byt�̵㵥�� As Byte = 0, _
    Optional ByVal bln����� As Boolean = True _
    )
'----------------------------------------------------------------------------------------
'���ܣ���ʼ��grsMaster��grsMasterInput��grsSlave����
'      Ϊ����ҩƷѡ����(frmSelector)������׼����
'������
'  byt�༭ģʽ�� 1����⣻ 2������
'  lng��Դ�ⷿ��
'----------------------------------------------------------------------------------------
    Const CON_FMT = "'999999999990.99999'"
    
    Dim strsql As String, strTmp As String
    Dim strUnit As String, strConversionUnit As String
    Dim rstemp As ADODB.Recordset
    Dim IntStockCheck As Integer
    Dim intUnit As Integer, intCostDigit As Integer, intPricedigit As Integer, intNumberDigit As Integer, intMoneyDigit As Integer
    
    On Error GoTo errHandle
    With grsMaster
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    With grsMasterInput
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    With grsSlave
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    
    '������λ
    If strModeName = "ҩƷ�������" Or strModeName = "ҩƷ�ƿ����" Then
        Call GetDrugDigit(lngʹ�ò���, strModeName, intUnit, intCostDigit, intPricedigit, intNumberDigit, intMoneyDigit)
    Else
        Call GetDrugDigit(IIf(lng��Դ�ⷿ = 0, lngĿ��ⷿ, lng��Դ�ⷿ), strModeName, intUnit, intCostDigit, intPricedigit, intNumberDigit, intMoneyDigit)
    End If
    Select Case intUnit
        Case 1: strConversionUnit = "1"
        Case 2: strConversionUnit = "d.�����װ"
        Case 3: strConversionUnit = "d.סԺ��װ"
        Case Else
            strConversionUnit = "d.ҩ���װ"
    End Select
    
    '�����
    If bln����� = True And strModeName = "ҩƷ�������" Then
        bln����� = (Val(zldatabase.GetPara("ҩƷ�����γ���", glngSys, 1343, 0)) = 1)
    End If
    
    '��鲢ִ�е���
    Call AutoAdjustPrice_Batch
    
    '��ȡ����������ȷ����治��Ĳ���ȡ����
    strsql = "Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1] "
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "��ȡ�Ƿ���������", lng��Դ�ⷿ)
    If Not rstemp.EOF Then IntStockCheck = Nvl(rstemp!�����, 0)
    rstemp.Close
    
    '*ѡ��ģʽ�����ݼ�*'
    strsql = _
        "Select " & _
        " d.����,d.��ҩ��̬, d.ҩ������, d.ͨ������, d.ҩƷ��Դ As ��Դ, d.����ҩ��, d.ҩ��id, d.��;����id, d.������λ, d.ҩƷ����, d.ҩƷ����, " & _
        " d.��Ʒ��, d.���, d.���� As ������, Decode(s.ԭ����, Null, d.ԭ����, s.ԭ����) as ԭ����, d.ҩ��id, d.ҩƷid," & _
        " trim(to_char(d.��ʼ�ɱ��� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) �ϴβɹ���, " & _
        " trim(to_char(Decode(d.ʱ��, '��', Decode(s.ƽ���ۼ�, Null, Nvl(d.�ϴ��ۼ�,p.�ۼ�), s.ƽ���ۼ�), p.�ۼ�) * " & strConversionUnit & ", '99999999999990." & String(intPricedigit, "0") & "')) �ۼ�, " & _
        " d.�ۼ۵�λ, d.����ϵ�� As �ۼ۰�װ," & _
        " d.���ﵥλ, d.�����װ, d.סԺ��λ, d.סԺ��װ, d.ҩ�ⵥλ, d.ҩ���װ, " & _
        " nvl(trim(to_char(s.�������� / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')),0) ��������, " & _
        " s.�������,s.�����,s.�����,d.���Ч�� ��Ч��, d.ҩ�����, d.ҩ������, d.ʱ��," & _
        " trim(to_char(d.ָ�������� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as ָ��������, " & _
        " d.�ӳ���, e.�ⷿ��λ, d.��׼�ĺ�, s.������� ʵ������, " & _
        " s.��������, d.��ͬ��λ, d.ҩ�ۼ���,e.���ñ�־,d.ͣ��,d.�ϴι�Ӧ�� " & vbNewLine & _
        "From " & vbNewLine & _
        "  (SELECT DISTINCT J.���� ����,Decode(c.���, '7', Decode(d.��ҩ��̬, 1, '��Ƭ', 2, '����', 'ɢװ'), '') As ��ҩ��̬,A.���� ��Ʒ��, C.���� ҩ������,C.���� ͨ������, 0 AS ҩ��ID,C.���� ҩƷ����,C.���� ҩƷ����," & vbNewLine & _
        "     C.���,C.����,d.ԭ����,C.���,C.���㵥λ AS �ۼ۵�λ,DECODE(C.�Ƿ���,1,'��','��') ʱ��,D.ҩƷ��Դ,D.����ҩ��,D.��׼�ĺ�, D.ҩ��ID," & vbNewLine & _
        "     D.ҩƷID, nvl(to_char(D.���Ч��,'9999990'),0) ���Ч��," & vbNewLine & _
        "     DECODE(D.ҩ�����,1,'��','��') ҩ�����,DECODE(D.ҩ������,1,'��','��') ҩ������," & vbNewLine & _
        "     to_char(D.����ϵ��, " & CON_FMT & ") ����ϵ��," & vbLf & _
        "     D.���ﵥλ, to_char(D.�����װ, " & CON_FMT & ") �����װ," & vbNewLine & _
        "     D.סԺ��λ, to_char(D.סԺ��װ, " & CON_FMT & ") סԺ��װ," & vbNewLine & _
        "     D.ҩ�ⵥλ, to_char(D.ҩ���װ, " & CON_FMT & ") ҩ���װ," & vbNewLine & _
        "     D.ָ��������, nvl(D.�ɱ���,0) ��ʼ�ɱ���,D.�ӳ���,D.ҩ�ۼ���," & vbNewLine & _
        "     M.����ID AS ��;����ID,M.���㵥λ AS ������λ,Q.���� As ��ͬ��λ,Decode(Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), '��','��') As ͣ��,d.�ϴ��ۼ�,f.���� �ϴι�Ӧ��  " & vbNewLine
    strsql = strsql & _
        "   FROM �շ���ĿĿ¼ C,ҩƷ��� D,�շ���Ŀ���� A,ҩƷ���� J,ҩƷ���� T,������ĿĿ¼ M,��Ӧ�� Q, ���Ʒ���Ŀ¼ E,��Ӧ�� F " & vbNewLine & _
        IIf(lng��Դ�ⷿ <> 0, "     ,(Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[2] Group By ִ�п���ID,�շ�ϸĿID) K", "") & vbNewLine & _
        IIf(lngĿ��ⷿ <> 0, "     ,(Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[3] Group By ִ�п���ID,�շ�ϸĿID) I ", "") & vbNewLine & _
        "   WHERE C.ID=D.ҩƷID AND D.ҩ��ID=T.ҩ��ID AND T.ҩ��ID=M.ID and m.����id=e.id AND M.��� IN ('5','6','7') " & _
        IIf(lng��Դ�ⷿ <> 0, "     And D.ҩƷID=K.�շ�ϸĿID" & IIf(bln���޴洢�ⷿҩƷ, "(+)", ""), "") & _
        IIf(lngĿ��ⷿ <> 0, "     And D.ҩƷID=I.�շ�ϸĿID" & IIf(bln���޴洢�ⷿҩƷ, "(+)", ""), "") & _
        "     AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 " & _
        "     And (C.վ�� = [1] or c.վ�� is null) AND T.ҩƷ����=J.����(+) And D.��ͬ��λID=Q.ID(+) And D.�ϴι�Ӧ��ID=f.ID(+) " & _
        IIf(bln����ͣ��ҩƷ = False, " And (C.����ʱ�� Is Null Or To_char(C.����ʱ��,'yyyy-MM-dd')='3000-01-01') ) D,", ") D,") & vbNewLine & _
        "(Select �շ�ϸĿid, �ּ� �ۼ� " & _
        " From �շѼ�Ŀ Where (Sysdate Between ִ������ And ��ֹ���� or Sysdate>=ִ������ And ��ֹ���� Is Null)" & _
        GetPriceClassString("") & ") P," & vbNewLine
    If byt���÷�ʽ = 1 Then
       '��������ҩ
       strsql = strsql & _
           "(Select a.ҩƷid,Max(�ϴβ���) AS ����,max(a.ԭ����) as ԭ����,Sum(a.��������) ��������," & _
           " To_Char(Sum(a.ʵ������), " & CON_FMT & ") �������," & _
           " To_Char(Sum(a.ʵ�ʽ��), " & CON_FMT & ") �����," & _
           " To_Char(Sum(a.ʵ�ʲ��), " & CON_FMT & ") �����," & _
           " Decode(Sum(nvl(ʵ������,0)), 0, null, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������)) As ƽ���ۼ�," & _
           " To_Char(Sum(b.ʵ������), '99999999999990.99') �������� " & vbNewLine & _
           "From ҩƷ��� A, ҩƷ���� B " & vbNewLine & _
           "Where a.����=1 and a.ҩƷid=b.ҩƷid And a.�ⷿid=b.�ⷿid and b.����id=[3] and b.�ڼ�=to_date(sysdate,'yyyy') "
    Else
       '��ҩ����ҩ
       strsql = strsql & _
           "(Select a.ҩƷid, Max(a.�ϴβ���) AS ����, max(a.ԭ����) as ԭ����,Sum(a.��������) ��������," & _
           " Sum(a.ʵ������) �������," & _
           " Sum(a.ʵ�ʽ��) �����," & _
           " Sum(a.ʵ�ʲ��) �����," & _
           " Decode(Sum(nvl(ʵ������,0)), 0, null, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������)) As ƽ���ۼ�," & _
           " '' �������� " & vbNewLine & _
           "From ҩƷ��� A " & vbNewLine & _
           "Where ����=1 "
    End If
    If lng��Դ�ⷿ <> 0 Or lngĿ��ⷿ <> 0 Then
       strsql = strsql & " And a.�ⷿID=" & IIf(lng��Դ�ⷿ = 0, "[3]", "[2]")
    End If
    strsql = strsql & vbNewLine & _
       "Group By a.ҩƷid) S," & vbNewLine & _
       "(Select ҩƷID,�ⷿID,�ⷿ��λ,���ñ�־ From ҩƷ�����޶� Where �ⷿID=[2]) E " & vbNewLine & _
       "Where D.ҩƷID=P.�շ�ϸĿID And D.ҩƷID=S.ҩƷID" & IIf(Not (IntStockCheck = 2 And byt�༭ģʽ = 2) Or byt�̵㵥�� = 1 Or Not bln�����, "(+)", "") & _
       "  And D.ҩƷID=E.ҩƷID(+) " & vbNewLine & _
       "Order By D.ҩ������,D.ҩƷ���� "
    Set grsMaster = zldatabase.OpenSQLRecord(strsql, "ҩƷ���", gstrNodeNo, lng��Դ�ⷿ, lngĿ��ⷿ)
    
    
    '*¼��ģʽ�����ݼ�*'
    strsql = _
        "Select " & _
        " d.����, d.ҩ������, d.ͨ������, d.ҩƷ��Դ ��Դ, d.����ҩ��, d.ҩ��id, d.��;����id, d.������λ, d.ҩƷ����, f.���� ҩƷ����, " & _
        " d.��Ʒ��, d.���, d.���� As ������, Decode(s.ԭ����, Null, d.ԭ����, s.ԭ����) as ԭ����, d.ҩ��id, d.ҩƷid," & _
        " trim(to_char(d.��ʼ�ɱ��� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) �ϴβɹ���, " & _
        " trim(to_char(Decode(d.ʱ��, '��', Decode(s.ƽ���ۼ�, Null, p.�ۼ�, s.ƽ���ۼ�), p.�ۼ�) * " & strConversionUnit & ", '99999999999990." & String(intPricedigit, "0") & "')) �ۼ�, " & _
        " d.�ۼ۵�λ, d.����ϵ�� �ۼ۰�װ,d.���ﵥλ, d.�����װ, d.סԺ��λ, d.סԺ��װ, d.ҩ�ⵥλ, d.ҩ���װ, " & _
        " nvl(trim(to_char(s.�������� / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')),0) ��������, " & _
        " s.�������,s.�����,s.�����, " & _
        " d.���Ч�� ��Ч��, d.ҩ�����, d.ҩ������, d.ʱ��," & _
        " trim(to_char(d.ָ�������� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as ָ��������, " & _
        " trim(to_char(d.ָ�����ۼ� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as ָ�����ۼ�, " & _
        "d.�ӳ���, e.�ⷿ��λ, d.��׼�ĺ�, s.������� ʵ������," & _
        " s.��������, d.��ͬ��λ, d.ҩ�ۼ���,e.���ñ�־, Max(Decode(����, '1', ����, Null)) ����, Max(Decode(����, '3', ����, Null)) ���ּ���,Max(Decode(����, '2', ����, Null)) �����,d.ͣ��,d.�ϴι�Ӧ�� " & vbNewLine & _
        "From " & vbNewLine & _
        "  (SELECT DISTINCT J.���� ����,C.���� ҩ������,C.���� AS ͨ������,0 AS ҩ��ID,M.����ID AS ��;����ID,M.���㵥λ AS ������λ, " & _
        "   C.���� AS ҩƷ����, a.���� As ��Ʒ��, c.���, c.����, d.ԭ����, d.ҩƷ��Դ, d.����ҩ��, d.��׼�ĺ�, d.ҩ��id, " & _
        "   d.ҩƷid, c.���㵥λ As �ۼ۵�λ, nvl(to_char(d.���Ч��, '9999990'),0) ���Ч��, " & _
        "   DECODE(D.ҩ�����,1,'��','��') ҩ�����, DECODE(D.ҩ������,1,'��','��') ҩ������, " & _
        "   to_char(D.����ϵ��, " & CON_FMT & ") ����ϵ��," & vbLf & _
        "   D.���ﵥλ, to_char(D.�����װ, " & CON_FMT & ") �����װ," & vbNewLine & _
        "   D.סԺ��λ, to_char(D.סԺ��װ, " & CON_FMT & ") סԺ��װ," & vbNewLine & _
        "   D.ҩ�ⵥλ, to_char(D.ҩ���װ, " & CON_FMT & ") ҩ���װ," & vbNewLine & _
        "   D.ָ��������,d.ָ�����ۼ�,nvl(D.�ɱ���,0) ��ʼ�ɱ���, D.�ӳ���, q.���� ��ͬ��λ, D.ҩ�ۼ���, " & _
        "   DECODE(C.�Ƿ���,1,'��','��') ʱ��,Decode(Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), '��','��') As ͣ��,f.���� �ϴι�Ӧ�� " & vbNewLine
    strsql = strsql & _
        "   From  �շ���ĿĿ¼ C,ҩƷ��� D,�շ���Ŀ���� A,ҩƷ���� J,ҩƷ���� T,������ĿĿ¼ M,��Ӧ�� Q,��Ӧ�� F" & vbNewLine & _
        IIf(lng��Դ�ⷿ <> 0, "     ,(Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[2] Group By ִ�п���ID,�շ�ϸĿID) K", "") & vbNewLine & _
        IIf(lngĿ��ⷿ <> 0, "     ,(Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID=[3] Group By ִ�п���ID,�շ�ϸĿID) I ", "") & vbNewLine & _
        "   Where c.Id = d.ҩƷid And d.ҩ��id = t.ҩ��id And t.ҩ��id = m.Id And m.��� In ('5', '6', '7') And d.ҩƷid = a.�շ�ϸĿid(+) " & _
        "     And a.����(+) = 3 And t.ҩƷ���� = j.����(+) And d.��ͬ��λid = q.Id(+) And D.�ϴι�Ӧ��ID=f.ID(+) " & _
        IIf(lng��Դ�ⷿ <> 0, "     And D.ҩƷID=K.�շ�ϸĿID" & IIf(bln���޴洢�ⷿҩƷ, "(+)", ""), "") & _
        IIf(lngĿ��ⷿ <> 0, "     And D.ҩƷID=I.�շ�ϸĿID" & IIf(bln���޴洢�ⷿҩƷ, "(+)", ""), "") & _
        IIf(bln����ͣ��ҩƷ = False, " And (C.����ʱ�� Is Null Or To_char(C.����ʱ��,'yyyy-MM-dd')='3000-01-01') ) D,", ") D,") & vbNewLine & _
        "  (Select �շ�ϸĿid, Trim(To_Char(�ּ�, '999999999990." & String(7, "0") & "')) �ۼ� " & _
        "   From �շѼ�Ŀ Where (Sysdate Between ִ������ And ��ֹ���� or Sysdate>=ִ������ And ��ֹ���� Is Null)" & _
        GetPriceClassString("") & ") P," & vbNewLine

    If byt���÷�ʽ = 1 Then
       '��������ҩ
       strsql = strsql & _
           "(Select a.ҩƷid,Max(�ϴβ���) AS ����, max(a.ԭ����) as ԭ����,Sum(a.��������) ��������," & _
           " To_Char(Sum(a.ʵ������), " & CON_FMT & ") �������," & _
           " To_Char(Sum(a.ʵ�ʽ��), " & CON_FMT & ") �����," & _
           " To_Char(Sum(a.ʵ�ʲ��), " & CON_FMT & ") �����," & _
           " Decode(Sum(Nvl(ʵ������, 0)), 0, Null, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������)) As ƽ���ۼ�, " & _
           " To_Char(Sum(b.ʵ������), '99999999999990.99') �������� " & vbNewLine & _
           "From ҩƷ��� A, ҩƷ���� B " & vbNewLine & _
           "Where a.����=1 and a.ҩƷid=b.ҩƷid And a.�ⷿid=b.�ⷿid and b.����id=[3] and b.�ڼ�=to_date(sysdate,'yyyy') "
    Else
       '��ҩ����ҩ
       strsql = strsql & _
           "(Select a.ҩƷid, Max(a.�ϴβ���) AS ����, max(a.ԭ����) as ԭ����,Sum(a.��������) ��������," & _
           " To_Char(Sum(a.ʵ������), " & CON_FMT & ") �������," & _
           " To_Char(Sum(a.ʵ�ʽ��), " & CON_FMT & ") �����," & _
           " To_Char(Sum(a.ʵ�ʲ��), " & CON_FMT & ") �����," & _
           " Decode(Sum(Nvl(ʵ������, 0)), 0, Null, Sum(a.ʵ�ʽ��) / Sum(a.ʵ������)) As ƽ���ۼ�, " & _
           " '' �������� " & vbNewLine & _
           "From ҩƷ��� A " & vbNewLine & _
           "Where ����=1 "
    End If
    If lng��Դ�ⷿ <> 0 Or lngĿ��ⷿ <> 0 Then
       strsql = strsql & " And a.�ⷿID=" & IIf(lng��Դ�ⷿ = 0, "[3]", "[2]")
    End If
    strsql = strsql & vbNewLine & _
       "Group By a.ҩƷid) S," & vbNewLine & _
       "(Select ҩƷID,�ⷿID,�ⷿ��λ,���ñ�־ From ҩƷ�����޶� Where �ⷿID=" & IIf(byt�༭ģʽ = 2, "[2]", "[3]") & ") E, �շ���Ŀ���� F " & vbNewLine & _
       "Where D.ҩƷID=P.�շ�ϸĿID And D.ҩƷID=S.ҩƷID" & IIf(Not (IntStockCheck = 2 And byt�༭ģʽ = 2) Or byt�̵㵥�� = 1 Or Not bln�����, "(+)", "") & _
       "  And D.ҩƷID=E.ҩƷID(+) And d.ҩƷid = f.�շ�ϸĿid(+) " & vbNewLine & _
       "Group By d.����, d.ҩ������, d.ͨ������, d.ҩƷ��Դ, d.����ҩ��, d.ҩ��id, d.��;����id, d.������λ, d.ҩƷ����, f.����, d.��Ʒ��, d.���, d.����," & vbNewLine & _
       "  Decode(s.ԭ����, Null, d.ԭ����, s.ԭ����), d.ҩ��id, d.ҩƷid, " & vbNewLine & _
       "  trim(to_char(d.��ʼ�ɱ��� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')), " & vbNewLine & _
       "  trim(to_char(Decode(d.ʱ��, '��', Decode(s.ƽ���ۼ�, Null, p.�ۼ�, s.ƽ���ۼ�), p.�ۼ�) * " & strConversionUnit & ", '99999999999990." & String(intPricedigit, "0") & "'))," & vbNewLine & _
       "   d.�ۼ۵�λ,d.����ϵ��, d.���ﵥλ, d.�����װ, d.סԺ��λ, d.סԺ��װ, d.ҩ�ⵥλ, d.ҩ���װ," & vbNewLine & _
       "  nvl(trim(to_char(s.�������� / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')),0), s.�������, s.�����, s.�����, " & vbNewLine & _
       "  d.���Ч��, d.ҩ�����, d.ҩ������, d.ʱ��,trim(to_char(d.ָ�������� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')), " & vbNewLine & _
       "  trim(to_char(d.ָ�����ۼ� * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')), d.�ӳ���,e.�ⷿ��λ, d.��׼�ĺ�, s.�������," & vbNewLine & _
       "  s.��������, d.��ͬ��λ, d.ҩ�ۼ���, e.���ñ�־, d.ͣ��, d.�ϴι�Ӧ�� " & vbNewLine & _
       "Order By D.ҩ������,D.ҩƷ���� "
    Set grsMasterInput = zldatabase.OpenSQLRecord(strsql, "ҩƷ���", gstrNodeNo, lng��Դ�ⷿ, lngĿ��ⷿ, IIf(gint���뷽ʽ = 0, 1, 2))
    
    '*ҩƷ����*'
    If byt�༭ģʽ = 2 Then
        strsql = _
            "Select Rid,�ⷿ,ҩƷID,����,�������,����,��������,���� as ������,ԭ����,�ɱ���,�ۼ�,ʱ��,���ﵥλ,�����װ,סԺ��λ,סԺ��װ,ҩ�ⵥλ,ҩ���װ," & _
            "  ��Ч��,ʵ������,��������,�������,�����,�����,�ϴι�Ӧ��ID,��׼�ĺ�,��Ӧ�� " & vbLf & _
            "From (Select Distinct 2 Rid, p.���� �ⷿ, k.ҩƷid, nvl(k.����,0) ����, To_Char(b.�������, 'YYYY-MM-DD') As �������, k.�ϴ����� ����," & _
            "  To_Char(k.�ϴ���������, 'YYYY-MM-DD') ��������, k.�ϴβ��� ����, Decode(k.ԭ����, Null, d.ԭ����, k.ԭ����) as ԭ����, k.ƽ���ɱ��� �ɱ���," & _
            "  Decode(Nvl(k.����, 0), 0, Decode(Sign(k.ʵ������), 1, k.ʵ�ʽ�� / decode(nvl(k.ʵ������,0), 0, 1, k.ʵ������), A.�ּ�) " & _
            "        ,Nvl(k.���ۼ�, k.ʵ�ʽ�� / decode(nvl(k.ʵ������,0), 0, 1, k.ʵ������) ) ) �ۼ�," & _
            "  Nvl(k.���ۼ�, k.ʵ�ʽ�� / decode(nvl(k.ʵ������,0), 0, 1, k.ʵ������) ) ʱ��," & _
            "  D.���ﵥλ, to_char(D.�����װ, " & CON_FMT & ") �����װ," & _
            "  D.סԺ��λ, to_char(D.סԺ��װ, " & CON_FMT & ") סԺ��װ," & _
            "  D.ҩ�ⵥλ, to_char(D.ҩ���װ, " & CON_FMT & ") ҩ���װ," & _
            "  k.Ч��" & IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "-1", "") & " ��Ч��," & _
            "  k.ʵ������, k.��������, k.ʵ������ �������, k.ʵ�ʽ�� �����, k.ʵ�ʲ�� �����, k.�ϴι�Ӧ��id, k.��׼�ĺ�,f.���� ��Ӧ�� " & vbNewLine & _
            "From ���ű� P, ҩƷ��� D, ҩƷ��� K, ҩƷ�����Ϣ B, �շѼ�Ŀ A,��Ӧ�� F " & vbNewLine & _
            "Where k.�ⷿid = p.Id And d.ҩƷid = k.ҩƷid And d.ҩƷid=a.�շ�ϸĿid " & GetPriceClassString("A") & _
            "  And k.���� = 1 And k.ҩƷid = b.ҩƷid(+) And k.�ⷿid = b.�ⷿid(+) And nvl(k.����,0)  = nvl(b.����(+),0) And k.�ⷿid = [1] And K.�ϴι�Ӧ��ID=f.ID(+) "
        If byt�̵㵥�� = 1 Then
            strsql = strsql & " And (K.ʵ������<>0 Or K.ʵ�ʽ��<>0 Or K.ʵ�ʲ��<>0) ) " & vbNewLine
'        ElseIf byt�̵㵥�� = 2 Then
'            '1303 ����ǿ���۵���ģ�飬��������˿������Ϊ0��ҩƷ��¼
'            gstrSQL = strSQL & " ) " & vbNewLine
        Else
            strsql = strsql & " And K.ʵ������<>0 ) " & vbNewLine
        End If
        If gtype_UserSysParms.P150_ҩƷ���������㷨 = 0 Then
            strsql = strsql & "Order By ҩƷid, ���� "
        Else
            strsql = strsql & "Order By ҩƷid, ��Ч��, ���� "
        End If

        Set grsSlave = zldatabase.OpenSQLRecord(strsql, "ҩƷ����", IIf(lng��Դ�ⷿ = 0, lngĿ��ⷿ, lng��Դ�ⷿ))
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReleaseSelectorRS()
    If Not grsMaster Is Nothing Then
        If grsMaster.State = adStateOpen Then grsMaster.Close
        Set grsMaster = Nothing
    End If
    
    If Not grsMasterInput Is Nothing Then
        If grsMasterInput.State = adStateOpen Then grsMasterInput.Close
        Set grsMasterInput = Nothing
    End If
    
    If Not grsSlave Is Nothing Then
        If grsSlave.State = adStateOpen Then grsSlave.Close
        Set grsSlave = Nothing
    End If
End Sub

Public Function GetVSFlexRows(ByVal vsfVal As VSFlexGrid, Optional ByVal blnHidden = False) As Long
'--------------------------------------------------------------
'���ܣ���VSFlexGrid��������������ͷ��
'������
'  blnHidden��True��������ص�������False�������ص�������
'���أ�������
'--------------------------------------------------------------
    Dim i As Long, lngRows As Long
    For i = 0 To vsfVal.rows - 1
        If blnHidden Then
            If vsfVal.RowHidden(i) Then lngRows = lngRows + 1
        Else
            If vsfVal.RowHidden(i) = False Then lngRows = lngRows + 1
        End If
    Next
    GetVSFlexRows = lngRows
End Function


Public Sub GetPriceClass()
    '���ݵ�¼վ���ȡҩƷ�ļ۸�ȼ�
    Dim rsData As ADODB.Recordset
    
    If gstrNodeNo <> "" And gstrNodeNo <> "-" Then
        gstrSQL = " Select a.�۸�ȼ� " & _
            " From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ� B " & _
            " Where a.�۸�ȼ� = b.���� And a.���� = 0 And b.�Ƿ�����ҩƷ = 1 And a.վ�� = [1] And Nvl(b.����ʱ��, Sysdate + 1) > Sysdate "
        Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "GetPriceClass", gstrNodeNo)
        
        If rsData.RecordCount > 0 Then gstrPriceClass = rsData!�۸�ȼ�
    End If
End Sub


Public Function GetPriceClassString(strTableName As String) As String
    '���ݴ����ı������ؼ۸�ȼ���������
    GetPriceClassString = " And " & IIf(strTableName = "", "�۸�ȼ� Is Null ", strTableName & ".�۸�ȼ� Is Null ")
    
End Function
