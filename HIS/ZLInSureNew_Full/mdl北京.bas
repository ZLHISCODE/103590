Attribute VB_Name = "mdl����"
Option Explicit
'----------------------------------------------------------------
'����Ŀ¼�У���W01��ͷ�ģ����Ƿ�����ʩ��Ŀ�����м���У����=2
'----------------------------------------------------------------
'������ˮ�ţ�ÿ�ν��㶼Ҫ�²���������ҽԺ���루8λ��+����ʱ��(YMDHms)+����ʱ��(YMDHms)
'��Ҫ�ṩС���ߣ�
'1�����ڽ���ʱ��ϲ�һ�������ˣ���˸ù����������ϴ�ǰ���޸ĳ�Ժ�����Ϣ�����Ժ�ѧ�򣺵����������̫һ������Ϊȷʵ�������ڶ�ʱ���ڳ���ϱȽ����ѣ�������˵�̫��ʱ�䣬ҽԺ�϶����⵽Ͷ�ߵġ�
'2���������ӡ�ɾ�����޸�ָ����ϵ
'3���޸ĳ���ʱע�⣬�󲿷�ģ���ڶ����г���

Private mblnInit As Boolean                 'ҽ����ʼ���ɹ���־
Public gcnBJYB As New ADODB.Connection

Private Type ComInfo_����
    ҽԺ����    As String
    ���Ŀ¼    As String
    ����Ŀ¼    As String
    ҵ������    As String
    ������ˮ��  As String
    ����        As String   '����/�ֲ��(�ֲ����S��β)
End Type
Public gComInfo_���� As ComInfo_����

Private Enum �ӿڹ���
    �������߻�����              'StartPolicy
    ֹͣ���߻�����              'StopPolicy
    ��ȡ�ֿ�������Ϣ            'GetPersonCommInfo
    ��ȡ�ֿ����˴�����Ϣ        'Get_SumInfo
    ��ȡ�ֲᲡ�˴�����Ϣ        'Get_SumInfo2
    ��ȡ���ⲡ��Ϣ              'Get_SpecInfo
    �ʻ�֧��                    'Reg_Account    ���ֿ����˴��ڿۼ��ʻ�
    ����ȷ��                    'Reg            ���ֿ�������Ҫִ�д˺���
    �ӿڰ汾��Ϣ                'Get_Ver        ��ȡ�ӿڵİ汾��Ϣ
    ��ȡ������Ϣ                'Get_ErrInfo    ��ȡ������Ϣ
    '���÷ֽ⺯��
    ���÷ֽ�_ͨ��1              'Divide
    ���÷ֽ�_ͨ��2              'Divide2
    ���÷ֽ�_��ͨ����           'Poli_Divide
    ���÷ֽ�_��������1          'Spec_Divide
    ���÷ֽ�_��������2          'Spec_Divide2
    ���÷ֽ�_��ͥ����1          'Home_Divide
    ���÷ֽ�_��ͥ����2          'Home_Divide2
    ���÷ֽ�_סԺ1              'Hosp_Divide
    ���÷ֽ�_סԺ2              'Hosp_Divide2
    ���÷ֽ�_סԺ3              'Hosp_Divide3
End Enum

Private Enum ���״���
    ���� = 0
    �������
    ��������ҽ����
    ����������֧��
    ����ͳ��֧��
    ������֧��
    ��������ۼ�
    ��������ʼ����
    �����ڽ��׺�
    ������ҽ����
    ������ͳ��֧��
    �����ڴ��֧��
    סԺ��ͥ�������
    ����סԺ��ͥ������ʶ
    ����סԺ��ͥ�������׺�
    ����סԺ��ͥ������ʼ����
    ����סԺ��ͥ����ҽ����
    ����סԺ��ͥ����ͳ��֧��
    ����סԺ���֧��
End Enum

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
Private Const mstrSplit As String = "|"

'�ӿ�����
'--��������˵��
Public Declare Function BJ_StartPolicy Lib "FYFJ.dll" Alias "StartPolicy" (ByVal ShowProgress As Long) As Long
Public Declare Function BJ_StopPolicy Lib "FYFJ.dll" Alias "StopPolicy" () As Long
Public Declare Function BJ_GetPersonCommInfo Lib "FYFJ.dll" Alias "GetPersonCommInfo" (ByVal strPersonalInfo As String) As Long
Public Declare Function BJ_Get_SumInfo Lib "FYFJ.dll" Alias "Get_SumInfo" (ByVal strSumInfo As String) As Long
Public Declare Function BJ_Get_SumInfo2 Lib "FYFJ.dll" Alias "Get_SumInfo2" (ByVal strPersonalInfo As String, ByVal strSumInfo As String) As Long
Public Declare Function BJ_Get_SpecInfo Lib "FYFJ.dll" Alias "Get_SpecInfo" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Divide Lib "FYFJ.dll" Alias "Divide" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Divide2 Lib "FYFJ.dll" Alias "Divide2" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Reg_Account Lib "FYFJ.dll" Alias "Reg_Account" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Get_Ver Lib "FYFJ.dll" Alias "Get_Ver" (ByVal strDllVer As String, ByVal strDateVer As String) As Long
Public Declare Function BJ_Get_ErrInfo Lib "FYFJ.dll" Alias "Get_ErrInfo" (ByVal strErrMsg As String) As Long
Public Declare Function BJ_Reg Lib "FYFJ.dll" Alias "Reg" (ByVal strIn As String, ByVal strOut As String) As Long
'--�����Ƿ��÷ֽ⺯��
Public Declare Function BJ_Poli_Divide Lib "FYFJ.dll" Alias "Poli_Divide" (ByVal strIn As String) As Long
Public Declare Function BJ_Spec_Divide Lib "FYFJ.dll" Alias "Spec_Divide" (ByVal strIn As String) As Long
Public Declare Function BJ_Spec_Divide2 Lib "FYFJ.dll" Alias "Spec_Divide2" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Home_Divide Lib "FYFJ.dll" Alias "Home_Divide" (ByVal strIn As String) As Long
Public Declare Function BJ_Home_Divide2 Lib "FYFJ.dll" Alias "Home_Divide2" (ByVal strIn As String, ByVal strOut As String) As Long
Public Declare Function BJ_Hosp_Divide Lib "FYFJ.dll" Alias "Hosp_Divide" (ByVal strIn As String) As Long
Public Declare Function BJ_Hosp_Divide2 Lib "FYFJ.dll" Alias "Hosp_Divide2" (ByVal strIn As String) As Long
Public Declare Function BJ_Hosp_Divide3 Lib "FYFJ.dll" Alias "Hosp_Divide3" (ByVal strIn As String) As Long


'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'�����ǹ��ܺ���
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
Private Function GetDeal(ByVal lng����ID As Long, ByVal str���� As String) As ADODB.Recordset
    '----------------------------------------------------------------
    '��������   �����ھ���Ǽǡ����㡢�������ʱ����ȡָ��������ǰ�Ĵ�����Ϣ
    '��д��     ������
    '��д����   ��2004-06-28
    '----------------------------------------------------------------
    On Error GoTo errHand
    Dim rsDeal As New ADODB.Recordset
    '��ȡָ��������ǰ�Ĵ�����Ϣ�����ֲ����Ѽ�¼����ȡ�������Ҫ����һ���ֲ����Ѽ�¼��
    Call DebugTool("��ȡ" & str���� & "��ǰ�Ĵ�����Ϣ")
    Call WriteBusinessLOG("��ȡ" & str���� & "��ǰ�Ĵ�����Ϣ", "", "")
    gstrSQL = " Select A.ҽ�����,to_Char(A.��Ժ����,'yyyyMMdd'),to_Char(A.��Ժ����,'yyyyMMdd'),A.��Ժ����,A.��Ժ����," & _
              "        A.�����ܶ�,A.ͳ��֧��,A.���֧��,A.�����Ը�,A.�����Է�,A.ͳ��ⶥ��ҽ����" & _
              " From �ֲ����Ѽ�¼ A,�����ʻ� B" & _
              " Where A.����=B.���� And B.����ID=" & lng����ID & " And A.��Ժ����<=TO_DATE('" & str���� & "','yyyy-MM-dd')"
    If rsDeal.State = 1 Then rsDeal.Close
    Call SQLTest(App.Title, "ZL9INSURE\GETDEAL", gstrSQL): rsDeal.Open gstrSQL, gcnBJYB: Call SQLTest
    Call DebugTool("������Ϣ��¼����Ϊ��" & rsDeal.RecordCount)
    Call WriteBusinessLOG("������Ϣ��¼����Ϊ��" & rsDeal.RecordCount, "", "")
    Set GetDeal = rsDeal
    Exit Function
errHand:
    Call DebugTool("��ȡ������Ϣʱ��������" & vbCrLf & "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("��ȡ������Ϣʱ��������" & vbCrLf & "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
End Function

Private Function MakeFile_Center(ByVal rsData As ADODB.Recordset, ByVal int���� As �ӿڹ���) As Boolean
    Dim strFile_In As String
    '----------------------------------------------------------------
    '��������   ��������ģ�飩���ݴ���ļ�¼�����ݣ�������Ӧ���ļ�
    '��д��     ������
    '��д����   ��2004-07-03
    '����˵��   ��֧�ֵĽӿڣ���ȡ�ֲᲡ�˴�����Ϣ�����÷ֽ�_��ͨ������÷ֽ�_��������1�����÷ֽ�_��ͥ����1�����÷ֽ�_סԺ1��������
    '  ����ļ� ��Ŀǰֻ��һ������ļ�������strFile_In
    '  �����ļ� �������ж������"||"�ָ�������strFile_Out
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    Select Case int����
    Case �ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ
        strFile_In = "med_note.in"
        '----�����ļ�˵��
'        ���    ������  ����    ��󳤶�    ˵��
'        1      ҽ�����    C   2   �μ���׼AKA130
'        2      ��Ժ����    D   8   סԺΪ�ֲ�����Ժ���ڣ��ֶ���ʼ���ڣ� ��ͨ�ż���������ⲡ����ͥ����Ϊ�ֲ��о�������
'        3      ��Ժ����    D   8   סԺΪ�ֲ��г�Ժ���ڣ��ֶν������ڣ� ��ͨ�ż���������ⲡ����ͥ����Ϊ�ֲ��о�������
'        4      ��Ժ����    C   2   סԺΪ��0-��ͨסԺ��1-���ⲡ����סԺ��2-������ֲסԺ��3-����סԺ��4-��ҽҽԺ��Ŀ�סԺ ��ͨ�ż���������ⲡ����ͥ����Ϊ��
'        5      ��Ժ����    C   1   0-������1-ת��Ժ����ͨ�ż���������ⲡ����ͥ����Ϊ�գ�
'        6      �����ܽ��  N   8,2 �ֲ��Ӧ��¼��
'        7      ͳ��֧�����    N   8,2 �ֲ��Ӧ��¼��
'        8      ������Ա������֧�����  N   8,2 �ǹ���ԱΪ�����ã�����ԱΪ����Ա�������
'        9      �����Ը����    N   8,2 �ֲ��Ӧ��¼��
'        10     �����Էѽ��    N   8,2 �ֲ��Ӧ��¼��
'        11     ͳ��ⶥ��ҽ���ڽ��    N   8,2 �ǹ���ԱΪ0������Ա���ֲ��Ӧ��¼��
    Case �ӿڹ���.���÷ֽ�_��ͨ����
        strFile_In = "Poli_Divide.in"
        '----�����ļ�˵��
'        ���    ������  ����    ��󳤶�    ˵��
'        1      �ļ���¼����    C   4   �ı��ļ������ļ�¼������
'        2      ������ˮ��  C   20  ҽ�ƻ������루8����룩+ҽԺ��ˮ�ţ�12�Ҷ��룩�м䲹�㡣��ҽԺ�˲���
'        3      ҽ�����    C   2   AKA130
'        4      �����ʶ    C   2   0-��ͨ��2-������ֲ
'        5      ���ν��׷����ܽ��  N   8,2
'        �ڶ�����Ϊ������ϸ��Ϣ�����������������б�
'        ���    ������  ����    ��󳤶�    ˵��
'        1      ��Ŀ���    C   9   ˳���
'        2      ������  C   20  �μ���׼AKC220���ɿ�
'        3      ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ��һ������
'        4      ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'        5      ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'        6      ����    N   10,4    AKC225
'        7      ����    N   8,2 AKC226
'        8      �����ܽ��  N   10,4    ʵ�ʽ�����
'        9      ���÷�������    D   8   ���÷�������
    Case �ӿڹ���.���÷ֽ�_��������1
        strFile_In = "Spec_Divide.in"
        '----�����ļ�˵��
'        ���    ������  ����    ��󳤶�    ˵��
'        1   ��Ŀ���    C   9   ˳���
'        2   ������  C   20  �μ���׼AKC220���ɿ�
'        3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'        4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'        5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'        6   ����    N   10,4    AKC225
'        7   ����    N   8,2 AKC226
'        8   �����ܽ��  N   10,4    ʵ�ʽ�����
'        9   ���÷�������    D   8   YYYYMMDD
    Case �ӿڹ���.���÷ֽ�_��ͥ����1
        strFile_In = "Home_Divide.in"
        '----�����ļ�˵��
'        ���    ������  ����    ��󳤶�    ˵��
'        1   ��Ŀ���    C   9   ˳���
'        2   ������  C   20  �μ���׼AKC220���ɿ�
'        3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'        4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'        5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'        6   ����    N   10,4    AKC225
'        7   ����    N   8,2 AKC226
'        8   �����ܽ��  N   10,4    ʵ�ʽ�����
'        9   ���÷�������    D   8   YYYYMMDD
    Case �ӿڹ���.���÷ֽ�_סԺ1
        strFile_In = "Hosp_Divide.in"
        '----�����ļ�˵��
'        ���    ������  ����    ��󳤶�    ˵��
'        1   ��Ŀ���    C   9   ˳���
'        2   ҽ����  C   20  ���ɿգ�
'        3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'        4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'        5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'        6   ����    N   10,4    AKC225
'        7   ����    N   8,2 AKC226
'        8   �����ܽ��  N   10,4    ʵ�ʽ�����
'        9   ���÷�������    D   8   YYYYMMDD
    
    Case Else
        Exit Function       '��֧�ָù��ܺ���
    End Select
    
    MakeFile_Center = MakeFile(rsData, strFile_In)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MakeFile(ByVal rsData As ADODB.Recordset, ByVal strFile As String) As Boolean
    Dim lng��¼�� As Long
    Dim dbl�����ܶ� As Double
    Dim strLine As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim lngCol As Long, lngCols As Long
    
    On Error GoTo errHand
    '----------------------------------------------------------------
    '��������   �����ݴ���ļ�¼�����ݣ�������Ӧ���ļ�
    '��д��     ������
    '��д����   ��2004-07-03
    '����˵��   ��֧�ֵĽӿڣ���ȡ�ֲᲡ�˴�����Ϣ�����÷ֽ�_��ͨ������÷ֽ�_��������1�����÷ֽ�_��ͥ����1�����÷ֽ�_סԺ1��������
    '  ����ļ� ��Ŀǰֻ��һ������ļ�������strFile_In
    '  �����ļ� �������ж������"||"�ָ�������strFile_Out
    '----------------------------------------------------------------
    
    '����ļ����ڣ�ɾ��
    
    Call DebugTool("����ļ����ھ�ɾ�������´���(zl9INSURE\MakeFile)" & vbCrLf & _
        "���strFile=" & strFile)
    Call WriteBusinessLOG("����ļ����ھ�ɾ�������´���(zl9INSURE\MakeFile)" & vbCrLf & _
        "���strFile=" & strFile, "", "")
    strFile = gComInfo_����.���Ŀ¼ & "\" & strFile
    If objFileSystem.FileExists(strFile) Then Call objFileSystem.DeleteFile(strFile, True)
    Set objStream = objFileSystem.CreateTextFile(strFile)
    
    '�˴�Ϊ��Ҫ��������Ĳ��֣�Ŀǰֻ����ͨ���
    If strFile = "Poli_Divide.in" Then
        '----�����ļ�˵��
'        ���    ������  ����    ��󳤶�    ˵��
'        1      �ļ���¼����    C   4   �ı��ļ������ļ�¼������
'        2      ������ˮ��  C   20  ҽ�ƻ������루8����룩+ҽԺ��ˮ�ţ�12�Ҷ��룩�м䲹�㡣��ҽԺ�˲���
'        3      ҽ�����    C   2   AKA130
'        4      �����ʶ    C   2   0-��ͨ��2-������ֲ
'        5      ���ν��׷����ܽ��  N   8,2
        With rsData
            lng��¼�� = .RecordCount
            Do While Not .EOF
                dbl�����ܶ� = dbl�����ܶ� + Nvl(!ʵ�ս��)
                .MoveNext
            Loop
        End With
        dbl�����ܶ� = Format(dbl�����ܶ�, "#####0.00;-#####0.00;0;")
        strLine = lng��¼�� & mstrSplit & gComInfo_����.������ˮ�� & mstrSplit & _
            gComInfo_����.ҵ������ & mstrSplit & "0" & mstrSplit & dbl�����ܶ�
        Call objStream.WriteLine(strLine)
    End If
    
    'ͳһ�����������¼�������ļ�
    With rsData
        lngCols = .Fields.Count
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strLine = ""
            For lngCol = 0 To lngCols - 1
                strLine = strLine & mstrSplit & Nvl(.Fields(lngCol).Value)
            Next
            strLine = Mid(strLine, Len(mstrSplit) + 1)
            objStream.WriteLine (strLine)
            .MoveNext
        Loop
    End With
    
    objStream.Close
    Set objStream = Nothing
    MakeFile = True
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\MakeFile)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\MakeFile)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
    If ErrCenter = 1 Then Resume
End Function

Private Function AnalyFile_Center(ByVal rsData As ADODB.Recordset, ByVal int���� As �ӿڹ���, _
    Optional ByVal blnԤ�� As Boolean = True) As Boolean
    Dim strReturn As String
    '----------------------------------------------------------------
    '��������   ��������ģ�飩�����ӿں���������ļ�������Ϊ�ڲ���¼�����������ݸ��µ����ݿ���
    '��д��     ������
    '��д����   ��2004-07-03
    '����˵��   ��֧�ֵĽӿڣ����÷ֽ�_��ͨ������÷ֽ�_��������1�����÷ֽ�_��ͥ����1�����÷ֽ�_סԺ1
    '  ����ļ� ��Ŀǰֻ��һ������ļ�������strFile_In
    '  �����ļ� �������ж������"||"�ָ�������strFile_Out
    '  ��blnԤ��Ϊ��ʱ������Ҫ����ֽ�����Ҳ����˵��������ϸ������һ�������Բ�����
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    Select Case int����
    Case �ӿڹ���.���÷ֽ�_��ͨ����
        strReturn = AnalyFile_��ͨ����(blnԤ��)
    Case �ӿڹ���.���÷ֽ�_��������1
        strReturn = AnalyFile_��������1(blnԤ��)
    Case �ӿڹ���.���÷ֽ�_��ͥ����1
        strReturn = AnalyFile_��ͥ����1(blnԤ��)
    Case �ӿڹ���.���÷ֽ�_סԺ1
        strReturn = AnalyFile_סԺ1(blnԤ��)
    Case Else
        Exit Function       '��֧�ָù��ܺ���
    End Select
    
    AnalyFile_Center = (strReturn <> "")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function AnalyFile_��ͨ����(Optional ByVal blnԤ�� As Boolean = True) As String
    '���ػ���������
    Dim strTotal As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Const strFile_Out As String = "Poli_Divide.out"
    On Error GoTo errHand
'    ----�����ļ�˵��
'    �ļ���һ��Ϊ�ܵķֽ��������������������б�
'    ���    ������  ����    ��󳤶�    ˵��
'    1      �ļ���¼����    C   4   �ı��ļ������ļ�¼������
'    2      ������ˮ��  C   20  ҽ�ƻ������루8����룩+ҽԺ��ˮ�ţ�12�Ҷ��룩�м䲹�㡣��ҽԺ�˲���
'    3      ҽ�����    C   2   �μ���׼AKA130
'    4      ���ν��׷����ܽ��  N   8,2
'    5      ҽ�����ܷ���    N   8,2
'    6      ҽ�����ܷ���    N   8,2
'    7      ���/����Ա֧����� N   8,2
'    8      ���/����Ա�Ը���� N   8,2
'    9      ����Ӧ���ܽ��  N   8,2
'    �ڶ�����Ϊ������ϸ��Ϣ�ķֽ�������ʽΪ��
'    ���    ������  ����    ��󳤶�    ˵��
'    1      ��Ŀ���    C   9   ˳���
'    2      ������  C   20  �μ���׼AKC220���ɿ�
'    3      ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'    4      ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'    5      ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'    6      ����    N   10,4
'    7      ����    N   8,2
'    8      �����ܽ��  N   10,4
'    9      ���÷�������    D   8   ���÷�������
'    10     ҽ���ڷ���  N   10,4
'    11     ҽ�������  N   10,4
'    12     �ֽ�״̬    C   1   0-������1-�����������ʶ��2-ҽ��Ŀ¼�ڲ����ڣ�3-���մ���
    
    objStream.Close
    Set objStream = Nothing
    AnalyFile_��ͨ���� = strTotal
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\MakeFile)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\MakeFile)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Private Function AnalyFile_��������1(Optional ByVal blnԤ�� As Boolean = True, _
    Optional ByVal lng����ID As Long = 0, Optional ByVal blnסԺ As Boolean = False) As String        '���ػ���������
    Dim lngRow As Long, lngRows As Long     '��ǰ�м�������
    Dim strTotal As String
    Dim arrRow
    Dim strNO As String                     '���ﵥ�ݺ�
    Dim str�����ڽ��׺� As String           '�����ڽ��׺ţ�������Ϣ�л�ȡ��
    Dim strҽ�� As String
    Dim str�������� As String               '���÷���ʱ��
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset     'ҽ����Ŀ��Ϣ
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Const strFile_Out As String = "Spec_Divide.out"
    Dim dbl�ֽ� As Double, dblҽ������ As Double, dbl���� As Double, dbl�����ܶ� As Double
    '�������ļ��ֶγ����������У�
    Const col��¼�� As Integer = 0
    Const col������ˮ�� As Integer = 1
    Const col�����ܶ� As Integer = 2
    Const col��ͨ����ҽ���� As Integer = 3
    Const col��ͨ����ҽ���� As Integer = 4
    Const colͳ��֧�� As Integer = 5
    Const colͳ���Ը� As Integer = 6
    Const col���֧�� As Integer = 7
    Const col����Ը� As Integer = 8
    Const col���ⲡ�Ը� As Integer = 9
    Const col���ⲡҽ���� As Integer = 10
    Const col�����Ը� As Integer = 11
    Const colͳ��ⶥ��ҽ���� As Integer = 12
    '�������ļ��ֶγ�����������ϸ��
    Const col��Ŀ��� As Integer = 0
    Const col������ As Integer = 1
    Const col��Ŀ���� As Integer = 2
    Const col��Ŀ���� As Integer = 3
    Const col��Ŀ��� As Integer = 4
    Const col���� As Integer = 5
    Const col���� As Integer = 6
    Const col��ϸ�ܶ� As Integer = 7
    Const col���÷������� As Integer = 8
    Const colҽ���� As Integer = 9
    Const colҽ���� As Integer = 10
    Const col��ϸ_�����Ը� As Integer = 11
    Const col�ֽ�״̬ As Integer = 12
    
    On Error GoTo errHand
'    ----�����ļ�˵��
'    �ú�����һ�������Ĵ����������ļ�Spec_Divide.out���ļ�������Ϊ�������ⲡ��������Ŀ�ķ��÷ֽ⣬���ļ�����HIS�ͻ��˺ͱ�DLL�Ľ���Ŀ¼SWAP_PATH���ڻ��������ж��壩���ļ������У���һ��Ϊ�ܵķֽ��������������������б�
'    ���    ������  ����    ��󳤶�    ˵��
'    1   �ļ���¼����    C   4   �ı��ļ������ļ�¼������
'    2   ������ˮ��  C   20  ҽ�ƻ������루8����룩+ҽԺ��ˮ�ţ�12�Ҷ��룩�м䲹�㡣��ҽԺ�˲���
'    3   ���ν��׷����ܽ��  N   8,2
'    4   ��ͨ����ҽ���ڷ���  N   8,2
'    5   ��ͨ����ҽ�������  N   8,2
'    6   ͳ��(���ⲡ)֧�����    N   8,2
'    7   ͳ����ⲡ���Ը����  N   8,2
'    8   ���/����Ա(���ⲡ)֧����� N   8,2
'    9   ���/����Ա(���ⲡ)�Ը���� N   8,2
'    10  ���ⲡ�����Ը����  N   8,2 ���ⲡ��ҩ������ҽ���ڸ����Ը����
'    11  ���ⲡҽ������    N   8,2 ���ⲡ��ҩ������ҽ������
'    12  �����Ը������  N   8,2 ������Ŀ���������˸�������
'    13  ���ν���ͳ��ⶥ��ҽ���ڽ��    N   8,2
'    �ڶ�����Ϊ������ϸ��Ϣ�ķֽ��������������������б�
'    ���    ������  ����    ��󳤶�    ˵��
'    1   ��Ŀ���    C   9   ˳���
'    2   ������  C   20  �μ���׼AKC220���ɿ�
'    3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'    4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'    5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'    6   ����    N   10,4
'    7   ����    N   8,2
'    8   �����ܽ��  N   10,4    ʵ�ʽ�����
'    9   ���÷�������    D   8   YYYYMMDD
'    10  ҽ���ڷ���  N   10,4
'    11  ҽ�������  N   10,4
'    12  �����Ը������  N   8,2 ������Ŀ���������˸�������
'    13  �ֽ�״̬    C   1   0-������1-�����������ʶ��2-ҽ��Ŀ¼�ڲ����ڣ�3-���մ���
    
    '�����Ԥ�㣬���������������ݣ���ֱ�ӷ���
    Call DebugTool("����(zl9Insure\AnalyFile_��������1)")
    Call WriteBusinessLOG("����(zl9Insure\AnalyFile_��������1)", "", "")
    If Not objFileSystem.FileExists(gComInfo_����.����Ŀ¼ & "\" & strFile_Out) Then Exit Function
    Set objStream = objFileSystem.OpenTextFile(gComInfo_����.����Ŀ¼ & "\" & strFile_Out)
    
    Call DebugTool("��ȡ����������(zl9Insure\AnalyFile_��������1)")
    Call WriteBusinessLOG("��ȡ����������(zl9Insure\AnalyFile_��������1)", "", "")
    '����ÿ���ı����Ի��з���������VB�ϵ��ǻس����У����������ݶ��������ˣ���ҪSPLIT
    strTotal = objStream.ReadLine
    arrRow = Split(strTotal, vbCr)
    objStream.Close
    Set objStream = Nothing
    
    lngRows = UBound(arrRow)
    'Ԥ�������Ҫ���������ݣ����ֽ��㷽ʽ֧���
    If blnԤ�� Then
        '�Է��÷ֽ⺯�����صķ�����ϸ���м�飬������ڲ������ķֽ�״̬����ʾ���˳�
        Call DebugTool("�Է��÷ֽ⺯�����صķ�����ϸ�ķֽ�״̬���м��(zl9Insure\AnalyFile_��������1)")
        Call WriteBusinessLOG("�Է��÷ֽ⺯�����صķ�����ϸ�ķֽ�״̬���м��(zl9Insure\AnalyFile_��������1)", "", "")
        For lngRow = 1 To lngRows
            If CheckDetail(Split(arrRow(lngRow), mstrSplit)(col��Ŀ����), Val(Split(arrRow(lngRow), mstrSplit)(col�ֽ�״̬))) Then Exit Function
        Next
        
        AnalyFile_��������1 = arrRow(0)
        Exit Function
    End If
    
    str�������� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    '��ȡ�շѵ��ݺ�
    Call DebugTool("��ȡ�շѵ��ݺ�(zl9Insure\AnalyFile_��������1")
    Call WriteBusinessLOG("��ȡ�շѵ��ݺ�(zl9Insure\AnalyFile_��������1", "", "")
    gstrSQL = "" & _
        " SELECT ʵ��Ʊ��,������" & _
        " From ������ü�¼" & _
        " WHERE ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շѵ��ݺ�", lng����ID)
    strNO = Nvl(rsTemp!ʵ��Ʊ��)
    strҽ�� = Nvl(rsTemp!������)
    
    If blnסԺ Then
        '��ȡ�շѵ��ݺ�
        Call DebugTool("�ӽ��ʵ���ȡ��Ʊ��(zl9Insure\AnalyFile_��������1")
        Call WriteBusinessLOG("�ӽ��ʵ���ȡ��Ʊ��(zl9Insure\AnalyFile_��������1", "", "")
        gstrSQL = "" & _
            " SELECT ʵ��Ʊ��" & _
            " From ���˽��ʼ�¼" & _
            " WHERE ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ӽ��ʵ���ȡ��Ʊ��", lng����ID)
        strNO = Nvl(rsTemp!ʵ��Ʊ��)
    End If
    
    '��ȡ�����ڽ��׺�
    Call DebugTool("�ӽ��״�����Ϣ����ȡ�����ڽ��׺�(zl9Insure\AnalyFile_��������1")
    Call WriteBusinessLOG("�ӽ��״�����Ϣ����ȡ�����ڽ��׺�(zl9Insure\AnalyFile_��������1", "", "")
    gstrSQL = "" & _
        " SELECT �����ڽ��׺�" & _
        " From ���״�����Ϣ" & _
        " WHERE ������ˮ��='" & gComInfo_����.������ˮ�� & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\AnalyFile_��������1", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str�����ڽ��׺� = Nvl(rsTemp!�����ڽ��׺�)
    
    '��ȡ�α��˻�����Ϣ��ֻ����ǰ��ʹ��rsTemp��¼������Ϊ����ֱ��ʹ���˸ü�¼�������ݣ�
    Call DebugTool("��ȡ�α��˻�����Ϣ(zl9Insure\AnalyFile_��������1")
    Call WriteBusinessLOG("��ȡ�α��˻�����Ϣ(zl9Insure\AnalyFile_��������1", "", "")
    gstrSQL = "" & _
        " SELECT ����,�籣֤��,�ɷѵ�������,ҵ������,�α����,����Ա,����Ա����,���ֱ�ʶ,���ⲡ��ֹ����" & _
        " From �����ʻ�" & _
        " WHERE ����='" & gComInfo_����.���� & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\AnalyFile_��������1", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    
    '��ʽ���㣬��Ҫ������������ݣ����ｻ����Ϣ����������ϸ���ݣ����������ϸ��
    '��������У���������
    '������ˮ��,�ɷѵ�������,�籣֤��,����,�ն˻����,�շѵ��ݺ�,ҵ������,�����ʶ,�����ڽ��׺�,����ʱ��
    '����Ա����(�ɿ�),��֤ͨ����ʽ(�ɿ�),�����ܶ�,��ͨ����ҽ����,��ͨ����ҽ����,ͳ��֧��,ͳ���Ը�,���֧��
    '����Ը�,���ⲡ�Ը�,���ⲡҽ����,�����Ը�,�����ʻ�֧��,�ֽ�֧��,�����ʻ����Ѻ����
    '���������ʶ(�ɿ�),MAC1(�ɿ�),�ϴ�
    dbl�����ܶ� = Val(Split(arrRow(0), mstrSplit)(col�����ܶ�))
    dblҽ������ = Val(Split(arrRow(0), mstrSplit)(colͳ��֧��))
    dbl���� = Val(Split(arrRow(0), mstrSplit)(col���֧��))
    dbl�ֽ� = dbl�����ܶ� - dbl���� - dblҽ������
    gstrSQL = "ZL_���ｻ����Ϣ_INSERT(" & _
        "'" & gComInfo_����.������ˮ�� & "','" & rsTemp!�ɷѵ������� & "','" & rsTemp!�籣֤�� & "'," & _
        "'" & rsTemp!���� & "',NULL,'" & strNO & "','" & rsTemp!ҵ������ & "','0','" & str�����ڽ��׺� & "'," & _
        "To_Date('" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & _
        "NULL,NULL," & Val(Split(arrRow(0), mstrSplit)(col�����ܶ�)) & "," & Val(Split(arrRow(0), mstrSplit)(col��ͨ����ҽ����)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col��ͨ����ҽ����)) & "," & Val(Split(arrRow(0), mstrSplit)(colͳ��֧��)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(colͳ���Ը�)) & "," & Val(Split(arrRow(0), mstrSplit)(col���֧��)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col����Ը�)) & "," & Val(Split(arrRow(0), mstrSplit)(col���ⲡ�Ը�)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col���ⲡҽ����)) & "," & Val(Split(arrRow(0), mstrSplit)(col�����Ը�)) & "," & _
        "0," & dbl�ֽ� & ",0,NULL,NULL,0)"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    '������ϸ��������ϸ������������
    '������ˮ��,��Ŀ���,����,������,ҽ������,��Ŀ����,��Ŀ���,����,���,����,����
    '��������,�����ܶ�,�շ����,ҽʦ����,ҽ���ڷ���,ҽ�������,�����Ը�,�ֽ�״̬,�ϴ�
    For lngRow = 1 To lngRows
        '��Ҫȡ����ҽ����Ŀ�ļ��͡�����շ���������ϴ���HIS�ж����Ŀ��Ӧһ��ҽ����Ŀʱ�����ܷ��ض�����¼��
        gstrSQL = "" & _
            " Select A.���,C.�շ����,C.����" & _
            " From �շ�ϸĿ A,����֧����Ŀ B," & GetUser & ".ҩƷĿ¼ C,������ü�¼ F" & _
            " Where A.ID=F.�շ�ϸĿID ANd A.ID=B.�շ�ϸĿID And B.��Ŀ����=C.���� " & _
            " And F.����ID=" & lng����ID & " AND F.���=" & Split(arrRow(lngRow), mstrSplit)(col��Ŀ���) & "" & _
            " Union" & _
            " Select A.���,C.�շ����,'' AS ����" & _
            " From �շ�ϸĿ A,����֧����Ŀ B," & GetUser & ".����Ŀ¼ C,������ü�¼ F" & _
            " Where A.ID=F.�շ�ϸĿID ANd A.ID=B.�շ�ϸĿID And B.��Ŀ����=C.���� " & _
            " And F.����ID=[1] AND F.���=[2]" & _
            " And A.��� Not In ('5','6','7')"
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ��Ϣ", lng����ID, CLng(Split(arrRow(lngRow), mstrSplit)(col��Ŀ���)))
        
        '���Ϊ�գ�˵���Ƿ�ҽ����Ŀ�����̶���ֵ
        If rsDetail.RecordCount = 0 Then
            gstrSQL = " Select A.���,A.��ʶ�� As �շ����,'' AS ���� " & _
                      " From ҩƷĿ¼ A,������ü�¼ F" & _
                      " Where A.ҩƷID=F.�շ�ϸĿID And F.����ID=" & lng����ID & " AND F.���=[2]"
            gstrSQL = gstrSQL & " UNION " & _
                      " Select A.���,A.��ʶ���� As �շ����,'' AS ���� " & _
                      " From �շ�ϸĿ A,������ü�¼ F" & _
                      " Where A.ID=F.�շ�ϸĿID And F.����ID=[1] AND F.���=[2]" & _
                      " And A.��� Not In ('5','6','7')"
            Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҽ����Ŀ����Ϣ", lng����ID, CLng(Split(arrRow(lngRow), mstrSplit)(col��Ŀ���)))
        End If
        
        gstrSQL = "ZL_���������ϸ_INSERT(" & _
            "'" & gComInfo_����.������ˮ�� & "','" & Val(Split(arrRow(lngRow), mstrSplit)(col��Ŀ���)) & "'," & _
            "'" & gComInfo_����.���� & "','" & strNO & "','" & Split(arrRow(lngRow), mstrSplit)(col��Ŀ����) & "'," & _
            "'" & Split(arrRow(lngRow), mstrSplit)(col��Ŀ����) & "','" & Split(arrRow(lngRow), mstrSplit)(col��Ŀ���) & "'," & _
            "'" & Nvl(rsDetail!����) & "','" & ToVarchar(Nvl(rsDetail!���), 40) & "'," & Val(Split(arrRow(lngRow), mstrSplit)(col����)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col����)) & ",TO_Date('" & str�������� & "','yyyy-MM-dd hh24:mi:ss')," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ�ܶ�)) & ",'" & Nvl(rsDetail!�շ����) & "','" & strҽ�� & "'," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(colҽ����)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(colҽ����)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_�����Ը�)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col�ֽ�״̬)) & ",0" & _
            ")"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    Next
    
    AnalyFile_��������1 = arrRow(0)
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\MakeFile)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\MakeFile)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Private Function AnalyFile_��ͥ����1(Optional ByVal blnԤ�� As Boolean = True) As String
    '���ػ���������
    Dim strTotal As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Const strFile_Out As String = "Home_Divide.out"
    On Error GoTo errHand
'    ----�����ļ�˵��
'    �ú�����һ�������Ĵ����������ļ�Home_Divide.out������������ϸ��Ϣ�����ļ�����HIS�ͻ��˺ͱ�DLL�Ľ���Ŀ¼SWAP_PATH���ڻ��������ж��壩���ļ������У���һ��Ϊ�ܵķֽ��������������������б�
'    ���    ������  ����    ��󳤶�    ˵��
'    1   �ļ���¼����    C   4
'    2   ������ˮ��  C   20  ҽ�ƻ������루8����룩+ҽԺ��ˮ�ţ�12�Ҷ��룩�м䲹�㡣��ҽԺ�˲���
'    3   ���ν��׷����ܽ��  N   8,2
'    4   ���ν���ҽ���ڷ����ܽ��    N   8,2
'    5   ͳ��֧�����    N   8,2
'    6   ͳ���Ը����    N   8,2
'    7   ���/����Ա֧����� N   8,2
'    8   ���/����Ա�Ը���� N   8,2
'    9   ����Ӧ���ܽ��  N   8,2
'    10  �����Ը������  N   8,2 ������Ŀ���������˸�������
'    11  ���ν���ͳ��ⶥ��ҽ���ڽ��    N   8,2
'    �ڶ�����Ϊ������ϸ��Ϣ�ֽ�������ʽΪ��
'    ���    ������  ����    ��󳤶�    ˵��
'    1   ��Ŀ���    C   9   ˳���
'    2   ������  C   20  �μ���׼AKC220���ɿ�
'    3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'    4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'    5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'    6   ����    N   10,4
'    7   ����    N   8,2
'    8   �����ܽ��  N   10,4    ʵ�ʽ�����
'    9   ���÷�������    D   8   YYYYMMDD
'    10  ҽ���ڷ���  N   10,4
'    11  ҽ�������  N   10,4
'    12  �����Ը������  N   8,2 ������Ŀ���������˸�������
'    13  �ֽ�״̬    C   1   0-������1-�����������ʶ��2-ҽ��Ŀ¼�ڲ����ڣ�3-���մ���
    
    objStream.Close
    Set objStream = Nothing
    AnalyFile_��ͥ����1 = strTotal
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\AnalyFile_��ͥ����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\AnalyFile_��ͥ����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Private Function AnalyFile_סԺ1(Optional ByVal blnԤ�� As Boolean = True, Optional ByVal lng����ID As Long = 0) As String
    '���ػ���������
    Dim strTotal As String
    Dim strUser As String
    Dim strNO As String
    Dim str��Ժ���� As String
    Dim str��ʼ���� As String, str�ֶο�ʼ���� As String, str�ֶν������� As String
    Dim arrRow
    Dim lngRow As Long, lngRows As Long
    Dim lng����ID As Long, lng��ҳID As Long
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim rs���÷������� As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsDetail As New ADODB.Recordset
    Dim dbl�����ܶ� As Double, dblҽ������ As Double, dbl���� As Double, dbl�ֽ� As Double
    Const strFile_Out As String = "Hosp_Divide.out"
    
    '�ֶγ���˵��
    Const col����_���÷ֶ��� As Integer = 0
    Const col����_������ϸ�� As Integer = 1
    Const col����_������ˮ�� As Integer = 2
    Const col����_�����ܶ� As Integer = 3
    Const col����_ҽ���� As Integer = 4
    Const col����_ͳ��֧�� As Integer = 5
    Const col����_ͳ���Ը� As Integer = 6
    Const col����_���֧�� As Integer = 7
    Const col����_����Ը� As Integer = 8
    Const col����_����Ӧ�� As Integer = 9
    Const col����_�����Ը� As Integer = 10
    Const col����_ͳ��ⶥ��ҽ���� As Integer = 11
    Const col�ֶ�_������� As Integer = 0
    Const col�ֶ�_������� As Integer = 1
    Const col�ֶ�_��ʼ���� As Integer = 2
    Const col�ֶ�_��ֹ���� As Integer = 3
    Const col�ֶ�_�����ܶ� As Integer = 4
    Const col�ֶ�_ҽ���� As Integer = 5
    Const col�ֶ�_ͳ��֧�� As Integer = 6
    Const col�ֶ�_ͳ���Ը� As Integer = 7
    Const col�ֶ�_���֧�� As Integer = 8
    Const col�ֶ�_����Ը� As Integer = 9
    Const col�ֶ�_�����Ը� As Integer = 10
    Const col�ֶ�_����Ӧ�� As Integer = 11
    Const col��ϸ_��Ŀ��� As Integer = 0
    Const col��ϸ_ҽ���� As Integer = 1
    Const col��ϸ_��Ŀ���� As Integer = 2
    Const col��ϸ_��Ŀ���� As Integer = 3
    Const col��ϸ_��Ŀ��� As Integer = 4
    Const col��ϸ_���� As Integer = 5
    Const col��ϸ_���� As Integer = 6
    Const col��ϸ_�����ܶ� As Integer = 7
    Const col��ϸ_���÷������� As Integer = 8
    Const col��ϸ_ҽ���� As Integer = 9
    Const col��ϸ_ҽ���� As Integer = 10
    Const col��ϸ_�����Ը� As Integer = 11
    Const col��ϸ_�ֽ�״̬ As Integer = 12
    On Error GoTo errHand
'    ----�����ļ�˵��
'    �ú�����һ�������Ĵ����������ļ�Hosp_Divide.out���������÷ֽ��������ļ�����HIS�ͻ��˺ͱ�DLL�Ľ���Ŀ¼SWAP_PATH���ڻ��������ж��壩���ļ���Ϊ�������֣���һ����Ϊ�ܵķֽ���������Ϊһ�С����������������б�
'    ���    ������  ����    ��󳤶�    ˵��
'    1   ���÷ֶηֽ�����    N   2
'    2   ������ϸ��¼����    N   4   �ı��ļ���������ϸ��¼������
'    3   ������ˮ��  C   20  ҽ�ƻ������루8����룩+ҽԺ��ˮ�ţ�12�Ҷ��룩�м䲹�㡣�޳�20λ����ҽԺ��Ԥ�������������ʧ��������
'    4   ���ν��׷����ܽ��  N   8,2
'    5   ���ν���ҽ�����ܽ��    N   8,2
'    6   ͳ��֧�����    N   8,2
'    7   ͳ���Ը����    N   8,2
'    8   ���/����Ա֧����� N   8,2
'    9   ���/����Ա�Ը���� N   8,2
'    10  ����Ӧ���ܽ��  N   8,2
'    11  �����Ը������  N   8,2 ������Ŀ���������˸�������
'    12  ���ν���ͳ��ⶥ��ҽ���ڽ��    N   8,2
'    �ڶ�����Ϊ���÷ֶηֽ�ķֽ���������Ϊ��һ���еķ��÷ֶηֽ����������������������б�
'    ���    ������  ����    ��󳤶�    ˵��
'    1   ���ý����������    N   2
'    2   �������    C   4
'    3   ���η�����ʼ����    D   8
'    4   ���η��ý�ֹ����    D   8
'    5   ���η����ܽ��  N   8,2
'    6   ���η���ҽ�����ܽ��    N   8,2
'    7   ���η���ͳ��֧�����    N   8,2
'    8   ���η���ͳ���Ը����    N   8,2
'    9   ���η��ô��/����Ա֧����� N   8,2
'    10  ���η��ô��/����Ա�Ը���� N   8,2
'    11  ���θ����Ը������  N   8,2 ������Ŀ���������˸�������
'    12  ���η��ø���Ӧ���ܽ��  N   8,2
'    ��������Ϊ������ϸ��Ϣ�ķֽ���������Ϊ��һ���еķ�����ϸ��¼���������������������б�
'    ���    ������  ����    ��󳤶�    ˵��
'    1   ��Ŀ���    C   9   ˳���
'    2   ҽ����  C   20  �ɿ�
'    3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'    4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'    5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'    6   ����    N   10,4
'    7   ����    N   8,2
'    8   �����ܽ��  N   10,4
'    9   ���÷�������    D   8   YYYYMMDD
'    10  ҽ���ڷ���  N   10,4
'    11  ҽ�������  N   10,4
'    12  �����Ը������  N   8,2 ������Ŀ���������˸�������
'    13  �ֽ�״̬    C   1   0-������1-�����������ʶ��2-ҽ��Ŀ¼�ڲ����ڣ�3-���մ���
    
    '�����Ԥ�㣬���������������ݣ���ֱ�ӷ���
    Call DebugTool("����(zl9Insure\AnalyFile_סԺ1)")
    Call WriteBusinessLOG("����(zl9Insure\AnalyFile_סԺ1)", "", "")
    If Not objFileSystem.FileExists(gComInfo_����.����Ŀ¼ & "\" & strFile_Out) Then Exit Function
    Set objStream = objFileSystem.OpenTextFile(gComInfo_����.����Ŀ¼ & "\" & strFile_Out)
    
    Call DebugTool("��ȡ����������(zl9Insure\AnalyFile_סԺ1)")
    Call WriteBusinessLOG("��ȡ����������(zl9Insure\AnalyFile_סԺ1)", "", "")
    '����ÿ���ı����Ի��з���������VB�ϵ��ǻس����У����������ݶ��������ˣ���ҪSPLIT
    strTotal = objStream.ReadLine
    arrRow = Split(strTotal, vbCr)
    objStream.Close
    Set objStream = Nothing
    
    'Ԥ�������Ҫ���������ݣ����ֽ��㷽ʽ֧���
    If blnԤ�� Then
        '�Է��÷ֽ⺯�����صķ�����ϸ���м�飬������ڲ������ķֽ�״̬����ʾ���˳�
        Call DebugTool("�Է��÷ֽ⺯�����صķ�����ϸ�ķֽ�״̬���м��(zl9Insure\AnalyFile_סԺ1)")
        Call WriteBusinessLOG("�Է��÷ֽ⺯�����صķ�����ϸ�ķֽ�״̬���м��(zl9Insure\AnalyFile_סԺ1)", "", "")
        For lngRow = (1 + Val(Split(arrRow(0), "|")(col����_���÷ֶ���))) To lngRows
            If CheckDetail(Split(arrRow(lngRow), mstrSplit)(col��ϸ_��Ŀ����), Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_�ֽ�״̬))) Then Exit Function
        Next
        
        AnalyFile_סԺ1 = arrRow(0)
        Exit Function
    End If
    
    strUser = GetUser
    '��ȡ�շѵ��ݺ�
    Call DebugTool("��ȡ�շѵ��ݺ�(zl9Insure\AnalyFile_��������1")
    Call WriteBusinessLOG("��ȡ�շѵ��ݺ�(zl9Insure\AnalyFile_��������1", "", "")
    gstrSQL = "" & _
        " SELECT ����ID,��ҳID" & _
        " From סԺ���ü�¼" & _
        " WHERE ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շѵ��ݺ�", lng����ID)
    lng����ID = rsTemp!����ID
    lng��ҳID = rsTemp!��ҳID
    gstrSQL = "" & _
        " SELECT ʵ��Ʊ��" & _
        " From ���˽��ʼ�¼" & _
        " WHERE ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շѵ��ݺ�", lng����ID)
    strNO = Nvl(rsTemp!ʵ��Ʊ��)
    
    '��ȡ���ÿ�ʼ����
    Call DebugTool("��ȡ���ÿ�ʼ����(zl9Insure\AnalyFile_סԺ1")
    Call WriteBusinessLOG("��ȡ���ÿ�ʼ����(zl9Insure\AnalyFile_סԺ1", "", "")
    gstrSQL = "Select min(�Ǽ�ʱ��) ��ʼ���� From סԺ���ü�¼ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ÿ�ʼ����", lng����ID)
    str��ʼ���� = Format(rsTemp!��ʼ����, "yyyy-MM-dd HH:mm:ss")
    
    '��ȡ��Ժ��ʽ
    Call DebugTool("��ȡ��Ժ���ڼ���Ժ��ʽ(zl9Insure\AnalyFile_סԺ1")
    Call WriteBusinessLOG("��ȡ��Ժ���ڼ���Ժ��ʽ(zl9Insure\AnalyFile_סԺ1", "", "")
    gstrSQL = "Select ��Ժ����,��Ժ��ʽ From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ���ڼ���Ժ��ʽ", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        MsgBox "δ�ҵ��ò��˵Ĳ�����Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��Ժ����:0-��Ժ,1-ת��Ժ��2-��;����
    If IsNull(rsTemp!��Ժ����) Then
        '�϶�����;����
        str��Ժ���� = 2
    Else
        '�жϳ�Ժ��ʽ
        str��Ժ���� = IIf(rsTemp!��Ժ��ʽ = "תԺ", 1, 0)
    End If
    
    '��ȡ�α��˻�����Ϣ��ֻ����ǰ��ʹ��rsTemp��¼������Ϊ����ֱ��ʹ���˸ü�¼�������ݣ�
    Call DebugTool("��ȡ�α��˻�����Ϣ(zl9Insure\AnalyFile_סԺ1")
    Call WriteBusinessLOG("��ȡ�α��˻�����Ϣ(zl9Insure\AnalyFile_סԺ1", "", "")
    gstrSQL = "" & _
        " SELECT A.����,A.�籣֤��,A.�ɷѵ�������,A.ҵ������,A.�α����,A.����Ա," & _
        "     A.����Ա����,A.���ֱ�ʶ,A.���ⲡ��ֹ����,B.��Ժ�ǼǺ�,B.��Ժ����,B.��Ժ��ʽ," & _
        "     to_Char(B.��Ժ����,'yyyy-MM-dd hh24:mi:ss') AS �����Ժ����,D.��Ժ����" & _
        " From " & strUser & ".�����ʻ� A," & strUser & ".��Ժ��Ϣ B,������Ϣ C," & strUser & ".��Ժ�����Ϣ D" & _
        " WHERE A.����ID=B.����ID And A.����ID=C.����ID And B.��ҳID=C.סԺ���� And B.����ID=D.����ID(+) And B.��ҳID=D.��ҳID(+) And A.����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�α��˻�����Ϣ", gComInfo_����.����)
    
    '��ʽ���㣬��Ҫ������������ݣ�סԺ������Ϣ����������ϸ���ݣ�סԺ������ϸ��
    '��������У���������
    '������ˮ��,�ɷѵ�������,�籣֤��,��Ժ�ǼǺ�,����,�ն˻����,��Ժ����,��Ժ��ʽ,��Ժ����,�����Ժ����,��Ժ����
    '��Ժ����,��������,�շѵ��ݺ�,����Ա����,��֤ͨ����ʽ,�����ܶ�,ҽ���ڽ��,ҽ������,ͳ��֧��,ͳ���Ը�,���֧��
    '����Ը�,�����Ը�,�ʻ�֧��,�ֽ�֧��,�����ʻ����Ѻ����,ͳ�ﶨ��,����,�����Ը�����Ա����,���˶����Ը�,
    '�������,���������ʶ,MAC1,ҽԺ�˽��㷽ʽ,�ϴ�
    dbl�����ܶ� = Val(Split(arrRow(0), mstrSplit)(col����_�����ܶ�))
    dblҽ������ = Val(Split(arrRow(0), mstrSplit)(col����_ͳ��֧��))
    dbl���� = Val(Split(arrRow(0), mstrSplit)(col����_���֧��))
    dbl�ֽ� = dbl�����ܶ� - dbl���� - dblҽ������
    
    'ҽԺ�˽��㷽ʽ:0����ͨ;1:������;2���ܶ�Ԥ��
    gstrSQL = "ZL_סԺ������Ϣ_INSERT(" & _
        "'" & gComInfo_����.������ˮ�� & "','" & rsTemp!�ɷѵ������� & "','" & rsTemp!�籣֤�� & "','" & rsTemp!��Ժ�ǼǺ� & "'," & _
        "'" & rsTemp!���� & "',NULL,'" & rsTemp!��Ժ���� & "','" & rsTemp!��Ժ��ʽ & "'," & _
        "to_Date('" & str��ʼ���� & "','yyyy-MM-dd hh24:mi:ss'),to_Date('" & rsTemp!�����Ժ���� & "','yyyy-MM-dd hh24:mi:ss')," & _
        "'" & str��Ժ���� & "',to_Date('" & IIf(IsNull(rsTemp!��Ժ����), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"), Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss")) & "','yyyy-MM-dd hh24:mi:ss')" & "," & _
        "to_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'),'" & strNO & "'," & _
        "NULL,NULL," & Val(Split(arrRow(0), mstrSplit)(col����_�����ܶ�)) & "," & Val(Split(arrRow(0), mstrSplit)(col����_ҽ����)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col����_�����ܶ�)) - Val(Split(arrRow(0), mstrSplit)(col����_ҽ����)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col����_ͳ��֧��)) & "," & Val(Split(arrRow(0), mstrSplit)(col����_ͳ���Ը�)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col����_���֧��)) & "," & Val(Split(arrRow(0), mstrSplit)(col����_����Ը�)) & "," & _
        "" & Val(Split(arrRow(0), mstrSplit)(col����_�����Ը�)) & ",0," & dbl�ֽ� & ",0,0,0,0,0,0,NULL,NULL,0,0)"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    '����ֶ���ϸ,��������
    lngRows = Val(Split(arrRow(0), mstrSplit)(col����_���÷ֶ���))
    '������ˮ��,���ý����������,�������,��Ժ�ǼǺ�,����,��ʼ����,��ֹ����,�����ܶ�
    'ҽ����,ͳ��֧��,ͳ���Ը�,���֧��,����Ը�,����Ӧ��,�����Ը�,ͳ�ﶨ��
    '����,�����Ը�����Ա����,���˶����Ը�,�������,�ϴ�
    For lngRow = 1 To lngRows
        str�ֶο�ʼ���� = Split(arrRow(lngRow), mstrSplit)(col�ֶ�_��ʼ����)
        str�ֶο�ʼ���� = Mid(str�ֶο�ʼ����, 1, 4) & "-" & Mid(str�ֶο�ʼ����, 5, 2) & "-" & Mid(str�ֶο�ʼ����, 7, 2)
        str�ֶν������� = Split(arrRow(lngRow), mstrSplit)(col�ֶ�_��ֹ����)
        str�ֶν������� = Mid(str�ֶν�������, 1, 4) & "-" & Mid(str�ֶν�������, 5, 2) & "-" & Mid(str�ֶν�������, 7, 2)
        gstrSQL = "ZL_סԺ���÷ֶ���ϸ_INSERT(" & _
                "'" & gComInfo_����.������ˮ�� & "','" & Split(arrRow(lngRow), mstrSplit)(col�ֶ�_�������) & "'," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_�������)) & ",'" & rsTemp!��Ժ�ǼǺ� & "','" & gComInfo_����.���� & "',to_Date('" & str�ֶο�ʼ���� & "','yyyy-MM-dd hh24:mi:ss')," & _
                "" & "to_Date('" & str�ֶν������� & "','yyyy-MM-dd hh24:mi:ss')," & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_�����ܶ�)) & "," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_ҽ����)) & "," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_ͳ��֧��)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_ͳ���Ը�)) & "," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_���֧��)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_����Ը�)) & "," & _
                "" & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_����Ӧ��)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col�ֶ�_�����Ը�)) & "," & _
                "" & "0,0,0,0,0,0)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    Next
    
    '������ϸ��������ϸ������������
    '������ˮ��,��Ŀ���,��Ժ�ǼǺ�,����,ҽ����,ҽ������,��Ŀ����,��Ŀ���,����,���,����,����
    '�����ܶ�,��������,�շ����,ҽ���ڷ���,ҽ�������,�����Ը�,�ֽ�״̬,�����ʶ,�ϴ�
    Dim str��ϸ_NO As String, str��ϸ_���� As String, str��ϸ_״̬ As String, str��ϸ_��� As String, str��ʶ As String
    lngRows = Val(Split(arrRow(0), mstrSplit)(col����_������ϸ��)) + (Val(Split(arrRow(0), "|")(col����_���÷ֶ���)))
    For lngRow = 1 + (Val(Split(arrRow(0), "|")(col����_���÷ֶ���))) To lngRows
        'col��ϸ_ҽ����=NO|��¼����|��¼״̬|���
        '�ֽ��NO����¼���ʡ���¼״̬�����
        str��ʶ = Split(arrRow(lngRow), mstrSplit)(col��ϸ_ҽ����)
        str��ϸ_NO = Split(str��ʶ, "*")(0)
        str��ϸ_���� = Split(str��ʶ, "*")(1)
        str��ϸ_״̬ = Split(str��ʶ, "*")(2)
        str��ϸ_��� = Split(str��ʶ, "*")(3)
        
        '��Ҫȡ����ҽ����Ŀ�ļ��͡�����շ���������ϴ���HIS�ж����Ŀ��Ӧһ��ҽ����Ŀʱ�����ܷ��ض�����¼��
        gstrSQL = "" & _
            " Select A.���,C.�շ����,C.����" & _
            " From �շ�ϸĿ A,����֧����Ŀ B," & strUser & ".ҩƷĿ¼ C,סԺ���ü�¼ F" & _
            " Where A.ID=F.�շ�ϸĿID ANd A.ID=B.�շ�ϸĿID And B.��Ŀ����=C.���� " & _
            " And F.NO=[1] And F.��¼����=[2] ANd F.��¼״̬=[3] ANd F.���=[4]" & _
            " Union" & _
            " Select A.���,C.�շ����,'' AS ����" & _
            " From �շ�ϸĿ A,����֧����Ŀ B," & strUser & ".����Ŀ¼ C,סԺ���ü�¼ F" & _
            " Where A.ID=F.�շ�ϸĿID ANd A.ID=B.�շ�ϸĿID And B.��Ŀ����=C.���� " & _
            " And F.NO=[1] And F.��¼����=[2] ANd F.��¼״̬=[3] ANd F.���=[4]" & _
            " And A.��� Not In ('5','6','7')"
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ��Ϣ", str��ϸ_NO, str��ϸ_����, str��ϸ_״̬, str��ϸ_���)
        
        '���Ϊ�գ�˵���Ƿ�ҽ����Ŀ�����̶���ֵ
        If rsDetail.RecordCount = 0 Then
            gstrSQL = " Select A.���,A.��ʶ�� As �շ����,'' AS ���� " & _
                      " From ҩƷĿ¼ A,סԺ���ü�¼ F" & _
                      " Where A.ҩƷID=F.�շ�ϸĿID And F.NO=[1] And F.��¼����=[2] ANd F.��¼״̬=[3] ANd F.���=[4]"
            gstrSQL = gstrSQL & " UNION " & _
                      " Select A.���,A.��ʶ���� As �շ����,'' AS ���� " & _
                      " From �շ�ϸĿ A,סԺ���ü�¼ F" & _
                      " Where A.ID=F.�շ�ϸĿID And F.NO=[1] And F.��¼����=[2] ANd F.��¼״̬=[3] ANd F.���=[4]" & _
                      " And A.��� Not In ('5','6','7')"
            Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҽ����Ŀ����Ϣ", str��ϸ_NO, str��ϸ_����, str��ϸ_״̬, str��ϸ_���)
        End If
        
        gstrSQL = "Select to_Char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') AS ����ʱ�� From סԺ���ü�¼ F" & _
            " WHERE F.NO=[1] And F.��¼����=[2] ANd F.��¼״̬=[3] ANd F.���=[4]"
        Set rs���÷������� = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ü�¼�ķ��÷���ʱ��", str��ϸ_NO, str��ϸ_����, str��ϸ_״̬, str��ϸ_���)
        str��ʼ���� = rs���÷�������!����ʱ��
        
        gstrSQL = "ZL_סԺ������ϸ_INSERT(" & _
            "'" & gComInfo_����.������ˮ�� & "','" & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_��Ŀ���)) & "'," & _
            "'" & rsTemp!��Ժ�ǼǺ� & "','" & gComInfo_����.���� & "',NULL,'" & Split(arrRow(lngRow), mstrSplit)(col��ϸ_��Ŀ����) & "'," & _
            "'" & Split(arrRow(lngRow), mstrSplit)(col��ϸ_��Ŀ����) & "','" & Split(arrRow(lngRow), mstrSplit)(col��ϸ_��Ŀ���) & "'," & _
            "'" & Nvl(rsDetail!����) & "','" & ToVarchar(Nvl(rsDetail!���), 40) & "'," & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_����)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_����)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_�����ܶ�)) & "," & _
            "to_Date('" & str��ʼ���� & "','yyyy-MM-dd hh24:mi:ss'),'" & Nvl(rsDetail!�շ����) & "'," & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_ҽ����)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_ҽ����)) & "," & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_�����Ը�)) & "," & _
            "" & Val(Split(arrRow(lngRow), mstrSplit)(col��ϸ_�ֽ�״̬)) & ",0,0)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    Next
    
    AnalyFile_סԺ1 = arrRow(0)
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\AnalyFile_סԺ1)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\AnalyFile_סԺ1)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
    If Not objStream Is Nothing Then
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Private Function CheckDetail(ByVal str��Ŀ���� As String, ByVal int�ֽ�״̬ As Integer) As Boolean
    Dim strMsg As String
    '----------------------------------------------------------------
    '��������   ����������ϸ�ֽ����������ڲ���������Ŀ��������ʾ
    '��д��     ������
    '��д����   ��2004-07-03
    '----------------------------------------------------------------
    Exit Function
    
    strMsg = "����Ŀ[" & str��Ŀ���� & "]�ķֽ�״̬��������������Ϣ���£�" & vbCrLf
    Select Case int�ֽ�״̬
    Case 0
        Exit Function
    Case 1
        strMsg = strMsg & "�����������ʶ"
    Case 2
        strMsg = strMsg & "ҽ��Ŀ¼�ڲ����ڸ���Ŀ"
    Case 3
        strMsg = strMsg & "ҽԺ��Ŀ��ҽ����Ŀ�������"
    Case Else
        strMsg = strMsg & "δ֪���󣬿�����ҽ���˽ӿ������б䶯����HIS�˵�ҽ���ӿ�δ����"
    End Select
    MsgBox strMsg, vbInformation, gstrSysName
    CheckDetail = True
End Function

Private Function CheckBlockage(ByVal StrInput As String, Optional ByVal bln���� As Boolean = True) As Boolean
    '----------------------------------------------------------------
    '��������   ��������С�ķ������ڣ�ʹ�ÿ��Ĳ��жϳ�ֵ��������
    '��д��     ������
    '��д����   ��2004-06-28
    '1.  ���˺��������粡�˿��ţ��ֲ�ţ��ڸ��˺������У� ����ʾ����Ա�Ҹÿ���ҽ���ֲᣩ����ʹ�ã���Ҫ��������Ϊ��ҽ�����ˣ�
    '2.  ��ֵ���������粡�˿����ڳ�ֵ�������У�����ʾ����Ӧȥ��ֵ�����ײ��ܽ��У�
    '3.  ��λ���������粡���籣�Ǽ�֤���ڵ�λ�������У�����ʾ���˲�����ҽ�����������з��þ�Ϊ�Էѣ������׿��ø����ʻ�����֧����
    '����˵��   ��bln����=False:strInput=����ID;����,strInput=����|�籣֤��
    '----------------------------------------------------------------
    Dim str���� As String, str�籣֤�� As String
    Dim bln�ÿ���־ As Boolean
    Dim rsPersonal As New ADODB.Recordset
    Dim rsBlockage As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡ���˵Ŀ��ż��籣֤��
    Call DebugTool("��ȡ���˵Ŀ��ż��籣֤��")
    Call WriteBusinessLOG("��ȡ���˵Ŀ��ż��籣֤��", "", "")
    If Not bln���� Then
        gstrSQL = "Select ����,�籣֤�� " & _
            " From �����ʻ� " & _
            " Where ����ID=" & Val(str����)
        If rsPersonal.State = 1 Then rsPersonal.Close
        Call SQLTest(App.Title, "ZL9INSURE\CheckBlockage", gstrSQL): rsPersonal.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsPersonal.EOF Then
            Call DebugTool("δ�ҵ��òα��˵��ʻ���Ϣ[zlBJ.�����ʻ�]")
            Call WriteBusinessLOG("δ�ҵ��òα��˵��ʻ���Ϣ[zlBJ.�����ʻ�]", "", "")
            Exit Function
        End If
        str���� = Nvl(rsPersonal!����)
        str�籣֤�� = Nvl(rsPersonal!�籣֤��)
    Else
        str���� = Split(StrInput, mstrSplit)(0)
        str�籣֤�� = Split(StrInput, mstrSplit)(1)
    End If
    bln�ÿ���־ = Not (UCase(Right(str����, 1)) = "S")
    
    '�����жϸ��˺���������ֵ����������λ������
    Call DebugTool("�жϸ��˺�����")
    Call WriteBusinessLOG("�жϸ��˺�����", "", "")
    gstrSQL = "SELECT ����ԭ�� FROM ���˺����� WHERE ����='" & str���� & "'"
    If rsBlockage.State = 1 Then rsBlockage.Close
    Call SQLTest(App.Title, "ZL9INSURE\CheckBlockage", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsBlockage.RecordCount <> 0 Then
        MsgBox "��ҽ���ֲ���ڸ��˺������У�������ʹ�ã��밴��ͨ���˰���" & vbCrLf & _
            "����ԭ��" & rsBlockage!����ԭ��, vbInformation, gstrSysName
        Exit Function
    End If
    
    If bln�ÿ���־ Then
        Call DebugTool("�жϳ�ֵ������")
        Call WriteBusinessLOG("�жϳ�ֵ������", "", "")
        gstrSQL = "SELECT ����ԭ�� FROM ��ֵ������ WHERE ����='" & str���� & "'"
        If rsBlockage.State = 1 Then rsBlockage.Close
        Call SQLTest(App.Title, "ZL9INSURE\CheckBlockage", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsBlockage.RecordCount <> 0 Then
            MsgBox "�ÿ������ֵ����ܼ������ף�" & vbCrLf & _
                "����ԭ��" & rsBlockage!����ԭ��, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Call DebugTool("�жϵ�λ������")
    Call WriteBusinessLOG("�жϵ�λ������", "", "")
    gstrSQL = "SELECT ����ԭ�� FROM ��λ������ WHERE �籣֤��='" & str�籣֤�� & "'"
    If rsBlockage.State = 1 Then rsBlockage.Close
    Call SQLTest(App.Title, "ZL9INSURE\CheckBlockage", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsBlockage.RecordCount <> 0 Then
        MsgBox "��ҽ���ֲ���ڵ�λ�������У���Ҫ���������з��þ�Ϊ�Էѣ�������ʱ�����ʻ�֧����" & vbCrLf & _
            "����ԭ��" & rsBlockage!����ԭ��, vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckBlockage = True
    Exit Function
errHand:
    Call DebugTool("������Ϣ���ʱ��������" & vbCrLf & "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("������Ϣ���ʱ��������" & vbCrLf & "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
End Function

Private Function GetBlockage(ByVal StrInput As String, Optional ByVal bln���� As Boolean = True) As String
    '----------------------------------------------------------------
    '��������   ��������С�ķ������ڣ�ʹ�ÿ��Ĳ��жϳ�ֵ��������
    '��д��     ������
    '��д����   ��2004-06-28
    '1.  ���˺��������粡�˿��ţ��ֲ�ţ��ڸ��˺������У� ����ʾ����Ա�Ҹÿ���ҽ���ֲᣩ����ʹ�ã���Ҫ��������Ϊ��ҽ�����ˣ�
    '2.  ��ֵ���������粡�˿����ڳ�ֵ�������У�����ʾ����Ӧȥ��ֵ�����ײ��ܽ��У�
    '3.  ��λ���������粡���籣�Ǽ�֤���ڵ�λ�������У�����ʾ���˲�����ҽ�����������з��þ�Ϊ�Էѣ������׿��ø����ʻ�����֧����
    '����˵��   ��bln����=False:strInput=����ID;����,strInput=����|�籣֤��
    '----------------------------------------------------------------
    Dim str���� As String, str�籣֤�� As String
    Dim bln�ÿ���־ As Boolean
    Dim rsPersonal As New ADODB.Recordset
    Dim rsBlockage As New ADODB.Recordset
    Dim str�������� As String
    Dim str���˺����� As String, str��ֵ������ As String, str��λ������ As String
    On Error GoTo errHand
    
    '��ȡ���˵Ŀ��ż��籣֤��
    Call DebugTool("��ȡ���˵Ŀ��ż��籣֤��")
    Call WriteBusinessLOG("��ȡ���˵Ŀ��ż��籣֤��", "", "")
    If Not bln���� Then
        gstrSQL = "Select ����,�籣֤�� " & _
            " From �����ʻ� " & _
            " Where ����ID=" & Val(str����)
        If rsPersonal.State = 1 Then rsPersonal.Close
        Call SQLTest(App.Title, "ZL9INSURE\GETBLOCKAGE", gstrSQL): rsPersonal.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsPersonal.EOF Then
            Call DebugTool("δ�ҵ��òα��˵��ʻ���Ϣ[zlBJ.�����ʻ�]")
            Call WriteBusinessLOG("δ�ҵ��òα��˵��ʻ���Ϣ[zlBJ.�����ʻ�]", "", "")
            Exit Function
        End If
        str���� = Nvl(rsPersonal!����)
        str�籣֤�� = Nvl(rsPersonal!�籣֤��)
    Else
        str���� = Split(StrInput, mstrSplit)(0)
        str�籣֤�� = Split(StrInput, mstrSplit)(1)
    End If
    bln�ÿ���־ = Not (UCase(Right(str����, 1)) = "S")
    
    '�����жϸ��˺���������ֵ����������λ������
    Call DebugTool("�жϸ��˺�����")
    Call WriteBusinessLOG("�жϸ��˺�����", "", "")
    gstrSQL = "SELECT to_char(��������,'yyyy-MM-dd') As �������� FROM ���˺����� WHERE ����='" & str���� & "'"
    If rsBlockage.State = 1 Then rsBlockage.Close
    Call SQLTest(App.Title, "ZL9INSURE\GETBLOCKAGE", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsBlockage.RecordCount <> 0 Then
        str���˺����� = Nvl(rsBlockage!��������)
    End If
    
    If bln�ÿ���־ Then
        Call DebugTool("�жϳ�ֵ������")
        Call WriteBusinessLOG("�жϳ�ֵ������", "", "")
        gstrSQL = "SELECT to_char(��������,'yyyy-MM-dd') As �������� FROM ��ֵ������ WHERE ����='" & str���� & "'"
        If rsBlockage.State = 1 Then rsBlockage.Close
        Call SQLTest(App.Title, "ZL9INSURE\GETBLOCKAGE", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsBlockage.RecordCount <> 0 Then
            str��ֵ������ = Nvl(rsBlockage!��������)
        End If
    End If
    
    Call DebugTool("�жϵ�λ������")
    Call WriteBusinessLOG("�жϵ�λ������", "", "")
    gstrSQL = "SELECT to_char(��������,'yyyy-MM-dd') As �������� FROM ��λ������ WHERE �籣֤��='" & str�籣֤�� & "'"
    If rsBlockage.State = 1 Then rsBlockage.Close
    Call SQLTest(App.Title, "ZL9INSURE\GETBLOCKAGE", gstrSQL): rsBlockage.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsBlockage.RecordCount <> 0 Then
        str��λ������ = Nvl(rsBlockage!��������)
    End If
    
    If str���˺����� <> "" Then str�������� = str���˺�����
    If str��ֵ������ <> "" Then
        If str�������� <> "" Then
            If str��ֵ������ < str�������� Then
                str�������� = str��ֵ������
            End If
        Else
            str�������� = str��ֵ������
        End If
    End If
    If str��λ������ <> "" Then
        If str�������� <> "" Then
            If str��λ������ < str�������� Then
                str�������� = str��λ������
            End If
        Else
            str�������� = str��λ������
        End If
    End If
    
    GetBlockage = str��������
    Exit Function
errHand:
    Call DebugTool("��ȡ��������������ʱ��������" & vbCrLf & "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("��ȡ��������������ʱ��������" & vbCrLf & "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
End Function

Private Function SaveBusinessDeal(ByVal strDeal As String) As Boolean
    '----------------------------------------------------------------
    '��������   �����潻�״�����Ϣ
    '��д��     ������
    '��д����   ��2004-07-06
    '----------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim str�������� As String           '��������
    Dim int������ As Integer            '�Ƿ��ڵ�λ��������(0-��;1-��)
    Dim arrDeal
    
    On Error GoTo errHand
    
    Call DebugTool("����(zl9Insure\SaveBusinessDeal")
    Call WriteBusinessLOG("����(zl9Insure\SaveBusinessDeal", "", "")
    arrDeal = Split(strDeal, mstrSplit)
    
    '�жϴ˿��Ƿ��ڵ�λ��������
    Call DebugTool("�жϴ˿��Ƿ��ڵ�λ��������(zl9Insure\SaveBusinessDeal")
    Call WriteBusinessLOG("�жϴ˿��Ƿ��ڵ�λ��������(zl9Insure\SaveBusinessDeal", "", "")
    gstrSQL = "Select �������� From ��λ������ Where �籣֤��=" & _
        " (Select �籣֤�� From �����ʻ� Where ����='" & gComInfo_����.���� & "')"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\SaveBusinessDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp.EOF Then
        int������ = 1
        str�������� = ""
    Else
        int������ = 0
        If Not IsNull(rsTemp!��������) Then
            str�������� = Format(rsTemp!��������, "yyyy-MM-dd") '������Ϊ��
        End If
    End If
    
    '��ȡ�α��˻�����Ϣ
    Call DebugTool("��ȡ�α��˻�����Ϣ(zl9Insure\SaveBusinessDeal")
    Call WriteBusinessLOG("��ȡ�α��˻�����Ϣ(zl9Insure\SaveBusinessDeal", "", "")
    gstrSQL = "" & _
        " SELECT ҵ������,�α����,����Ա,����Ա����,���ֱ�ʶ,���ⲡ��ֹ����" & _
        " From �����ʻ�" & _
        " WHERE ����='" & gComInfo_����.���� & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\SaveBusinessDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    
    '׼�����潻�״�����Ϣ
    Call DebugTool("���ս��״�����Ϣ(zl9Insure\SaveBusinessDeal")
    Call WriteBusinessLOG("���ս��״�����Ϣ(zl9Insure\SaveBusinessDeal", "", "")
'    ҵ������,����,������ˮ��,�α����,����Ա������ʶ,����Ա����,���ⲡ��ʶ,���ⲡ��ֹ����,�������,
'    ��������ҽ����,����������֧��,����ͳ��֧��,����ͳ����,��������ۼ�,��������ʼ����,�����ڽ��׺�,
'    ������ҽ����,������ͳ��֧��,�����ڴ��֧��,סԺ��ͥ�������,����סԺ��ͥ������ʶ,����סԺ��ͥ�������׺�,
'    ����סԺ��ͥ������ʼ����,����סԺ��ͥ����ҽ����,����סԺ��ͥ����ͳ��֧��,
'    ����סԺ���֧��,����ʱ�Ƿ��ڵ�λ��������,��λ��������������,�ϴ�
    gstrSQL = "ZL_���״�����Ϣ_INSERT(" & _
              "'" & rsTemp!ҵ������ & "','" & gComInfo_����.���� & "','" & gComInfo_����.������ˮ�� & "'," & _
              "'" & rsTemp!�α���� & "','" & rsTemp!����Ա & "','" & IIf(rsTemp!����Ա���� = -1, "", rsTemp!����Ա����) & "'," & _
              "" & rsTemp!���ֱ�ʶ & "," & IIf(IsNull(rsTemp!���ⲡ��ֹ����), "NULL", "to_Date('" & rsTemp!���ⲡ��ֹ���� & "','yyyy-MM-dd')") & "," & _
              "" & arrDeal(���״���.�������) & "," & arrDeal(���״���.��������ҽ����) & "," & arrDeal(���״���.����������֧��) & "," & _
              "" & arrDeal(���״���.����ͳ��֧��) & "," & arrDeal(���״���.������֧��) & "," & arrDeal(���״���.��������ۼ�) & "," & _
              "" & IIf(Trim(arrDeal(���״���.��������ʼ����)) = "", "NULL", "To_Date('" & arrDeal(���״���.��������ʼ����) & "','yyyy-MM-dd')") & "," & _
              "'" & arrDeal(���״���.�����ڽ��׺�) & "'," & arrDeal(���״���.������ҽ����) & "," & _
              "" & arrDeal(���״���.������ͳ��֧��) & "," & arrDeal(���״���.�����ڴ��֧��) & "," & arrDeal(���״���.סԺ��ͥ�������) & "," & _
              "'" & arrDeal(���״���.����סԺ��ͥ������ʶ) & "','" & arrDeal(���״���.����סԺ��ͥ�������׺�) & "'," & _
              "" & IIf(Trim(arrDeal(���״���.����סԺ��ͥ������ʼ����)) = "", "NULL", "To_Date('" & arrDeal(���״���.����סԺ��ͥ������ʼ����) & "','yyyy-MM-dd')") & "," & _
              "" & arrDeal(���״���.����סԺ��ͥ����ҽ����) & "," & arrDeal(���״���.����סԺ��ͥ����ͳ��֧��) & "," & arrDeal(���״���.����סԺ���֧��) & "," & _
              "'" & int������ & "'," & IIf(str�������� = "", "NULL", "to_Date('" & str�������� & "','yyyy-MM-dd')") & _
              ")"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    SaveBusinessDeal = True
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\SaveBusinessDeal)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\SaveBusinessDeal)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveBusinessVersion(Optional ByVal bln���� As Boolean = True) As Boolean
    Dim strReturn As String
    Dim strDllVersion As String, strDataVersion As String, strHISVersion As String
    Dim strCompanyVersion As String, strPersonalVersion As String, strCostVersion As String
    '----------------------------------------------------------------
    '��������   �����潻�װ汾��Ϣ
    '��д��     ������
    '��д����   ��2004-07-06
    '----------------------------------------------------------------
    On Error GoTo errHand
    Dim rsVersion As New ADODB.Recordset
    
    Call DebugTool("׼�����ýӿڻ�ȡDLL�汾��(zl9Insure\SaveBusinessVersion)")
    Call WriteBusinessLOG("׼�����ýӿڻ�ȡDLL�汾��(zl9Insure\SaveBusinessVersion)", "", "")
    '��ȡ�ӿڲ���������汾�š����ݰ��汾��
    If Not ���ýӿ�_����(�ӿڰ汾��Ϣ, strReturn) Then Exit Function
    strDllVersion = Split(strReturn, "||")(0)
    strDataVersion = Split(strReturn, "||")(1)
    
    Call DebugTool("׼����ȡ�������汾��(zl9Insure\SaveBusinessVersion)")
    Call WriteBusinessLOG("׼����ȡ�������汾��(zl9Insure\SaveBusinessVersion)", "", "")
    '��ȡ��λ�����������˺���������ֵ�������汾��(05�����˺�����;06����λ������;07����ֵ������)
    gstrSQL = "Select �ļ�����,�汾�� From �汾���� Where �ļ����� IN ('05','55','06','56','07','57')"
    If rsVersion.State = 1 Then rsVersion.Close
    rsVersion.Open gstrSQL, gcnBJYB
    '������û�м�¼�������������û���µĺ��������������ϴεĺ����������ݿ���ֻ�������µļ�¼
    '������ȫ���İ汾�ſ϶���һ�µ�
    Do While Not rsVersion.EOF
        Select Case rsVersion!�ļ�����
        Case "05", "55"
            If strPersonalVersion = "" Then strPersonalVersion = Nvl(rsVersion!�汾��)
        Case "06", "56"
            If strCompanyVersion = "" Then strCompanyVersion = Nvl(rsVersion!�汾��)
        Case "07", "57"
            If strCostVersion = "" Then strCostVersion = Nvl(rsVersion!�汾��)
        End Select
        rsVersion.MoveNext
    Loop
    
    '��ȡHIS�����汾��
    Call DebugTool("׼����ȡҽ�������汾��(zl9Insure\SaveBusinessVersion)")
    Call WriteBusinessLOG("׼����ȡҽ�������汾��(zl9Insure\SaveBusinessVersion)", "", "")
    strHISVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    '����һ�����ν��׵Ľ��װ汾��Ϣ
'    ����,������ˮ��,����,DLL���,DLL���ݰ�,��λ������,���˺�����,��ֵ������,HIS����,�ϴ�
    Call DebugTool("���뽻�װ汾��Ϣ(zl9Insure\SaveBusinessVersion)")
    Call WriteBusinessLOG("���뽻�װ汾��Ϣ(zl9Insure\SaveBusinessVersion)", "", "")
    gstrSQL = "ZL_���װ汾��Ϣ_INSERT(" & IIf(bln����, 1, 2) & ",'" & gComInfo_����.������ˮ�� & "'," & _
              "'" & gComInfo_����.���� & "','" & strDllVersion & "','" & strDataVersion & "'," & _
              "'" & strCompanyVersion & "','" & strPersonalVersion & "','" & strCostVersion & "'," & _
              "'" & strHISVersion & "')"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    SaveBusinessVersion = True
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\SaveBusinessVersion)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\SaveBusinessVersion)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then Resume
End Function

Private Function GetUser() As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=[1] And ������='ҽ���û���'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�м���û���", TYPE_����)
    GetUser = rsTemp!����ֵ
End Function

Private Function TransationSpec(ByVal str���� As String) As String
    Dim arr����
    Dim strReturn As String
    Dim str���÷������� As String
    Dim rsTemp As New ADODB.Recordset
    
    '----------------------------------------------------------------
    '��������   �����潻�װ汾��Ϣ
    '��д��     ������
    '��д����   ��2004-07-06
    '����˵��   ��bln����=TRUE�����,FALSE��סԺ��
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    arr���� = Split(str����, mstrSplit)
    str���÷������� = Format(zlDatabase.Currentdate(), "yyyyMMdd")
    '��ȡ�α��˻�����Ϣ
    Call DebugTool("��ȡ�α��˻�����Ϣ(zl9Insure\TransationSpec")
    Call WriteBusinessLOG("��ȡ�α��˻�����Ϣ(zl9Insure\TransationSpec", "", "")
    gstrSQL = "" & _
        " SELECT ҵ������,�α����,����Ա,����Ա����,���ֱ�ʶ,TO_CHAR(���ⲡ��ֹ����,'yyyyMMdd') AS ���ⲡ��ֹ���� " & _
        " From �����ʻ�" & _
        " WHERE ����='" & gComInfo_����.���� & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\TransationSpec", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    
    strReturn = gComInfo_����.������ˮ�� & mstrSplit & gComInfo_����.���� & mstrSplit & _
        rsTemp!�α���� & mstrSplit & rsTemp!����Ա & mstrSplit & IIf(rsTemp!����Ա���� = -1, "", rsTemp!����Ա����) & mstrSplit & _
        rsTemp!���ֱ�ʶ & mstrSplit & Nvl(rsTemp!���ⲡ��ֹ����) & mstrSplit & arr����(���״���.�������) & mstrSplit & _
        arr����(���״���.��������ҽ����) & mstrSplit & arr����(���״���.����������֧��) & mstrSplit & arr����(���״���.����ͳ��֧��) & mstrSplit & _
        arr����(���״���.������֧��) & mstrSplit & arr����(���״���.��������ۼ�) & mstrSplit & arr����(���״���.��������ʼ����) & mstrSplit & _
        arr����(���״���.�����ڽ��׺�) & mstrSplit & arr����(���״���.������ҽ����) & mstrSplit & arr����(���״���.������ͳ��֧��) & mstrSplit & _
        arr����(���״���.�����ڴ��֧��) & mstrSplit & str���÷�������
        
    TransationSpec = strReturn
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\TransationSpec)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\TransationSpec)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function TransationHosp(ByVal str���� As String) As String
    Dim arr����
    Dim strReturn As String
    Dim strUser As String
    Dim str��Ժ���� As String               '������Ժ����
    Dim str��Ժ���� As String               '��;���㽫��ǰ������Ϊ��Ժ����
    Dim str�������� As String               '��С�ķ�������
    Dim rsTemp As New ADODB.Recordset
    
    '----------------------------------------------------------------
    '��������   �����潻�װ汾��Ϣ
    '��д��     ������
    '��д����   ��2004-07-06
    '����˵��   ��bln����=TRUE�����,FALSE��סԺ��
    '----------------------------------------------------------------
    On Error GoTo errHand
    
    arr���� = Split(str����, mstrSplit)
    strUser = GetUser()
    
    '��ȡ�α��˱���סԺ������Ϣ
    Call DebugTool("��ȡ�α��˱���סԺ������Ϣ(zl9Insure\TransationHosp")
    Call WriteBusinessLOG("��ȡ�α��˱���סԺ������Ϣ(zl9Insure\TransationHosp", "", "")
    gstrSQL = "" & _
        " SELECT A.�籣֤��,A.ҵ������,A.�α����,A.����Ա,A.����Ա����,A.���ֱ�ʶ," & _
        "        TO_CHAR(A.���ⲡ��ֹ����,'yyyyMMdd') AS ���ⲡ��ֹ����," & _
        "        B.��Ժ����,B.��Ժ��ʽ,B.��Ժ����,C.��Ժ����" & _
        " From " & strUser & ".�����ʻ� A," & strUser & ".��Ժ��Ϣ B," & strUser & ".��Ժ�����Ϣ C,������Ϣ D" & _
        " WHERE A.����ID=B.����ID And B.����ID=C.����ID(+) and B.��ҳID=C.��ҳID(+)" & _
        " And A.����ID=D.����ID And B.��ҳID=D.סԺ���� And A.����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�α��˱���סԺ������Ϣ", gComInfo_����.����)
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyyMMdd")
    If Not IsNull(rsTemp!��Ժ����) Then
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyyMMdd")
    Else
        str��Ժ���� = Format(zlDatabase.Currentdate, "yyyyMMdd")
    End If
    
    '���������ڸ�ʽת����yyyyMMdd
    str�������� = GetBlockage(gComInfo_����.���� & mstrSplit & rsTemp!�籣֤��)
    If str�������� <> "" Then
        str�������� = Format(str��������, "yyyyMMdd")
    End If
    strReturn = gComInfo_����.������ˮ�� & mstrSplit & gComInfo_����.���� & mstrSplit & _
        rsTemp!�α���� & mstrSplit & rsTemp!����Ա & mstrSplit & IIf(rsTemp!����Ա���� = -1, "", rsTemp!����Ա����) & mstrSplit & _
        rsTemp!��Ժ���� & mstrSplit & rsTemp!���ֱ�ʶ & mstrSplit & Nvl(rsTemp!���ⲡ��ֹ����) & mstrSplit & _
        rsTemp!��Ժ��ʽ & mstrSplit & str��Ժ���� & mstrSplit & str��Ժ���� & mstrSplit & arr����(���״���.�������) & mstrSplit & _
        arr����(���״���.��������ҽ����) & mstrSplit & arr����(���״���.����������֧��) & mstrSplit & arr����(���״���.����ͳ��֧��) & mstrSplit & _
        arr����(���״���.������֧��) & mstrSplit & arr����(���״���.��������ۼ�) & mstrSplit & arr����(���״���.��������ʼ����) & mstrSplit & _
        arr����(���״���.�����ڽ��׺�) & mstrSplit & arr����(���״���.������ҽ����) & mstrSplit & arr����(���״���.������ͳ��֧��) & mstrSplit & _
        arr����(���״���.�����ڴ��֧��) & mstrSplit & arr����(���״���.סԺ��ͥ�������) & mstrSplit & arr����(���״���.����סԺ��ͥ������ʶ) & mstrSplit & _
        arr����(���״���.����סԺ��ͥ�������׺�) & mstrSplit & arr����(���״���.����סԺ��ͥ������ʼ����) & mstrSplit & arr����(���״���.����סԺ��ͥ����ҽ����) & mstrSplit & _
        arr����(���״���.����סԺ��ͥ����ͳ��֧��) & mstrSplit & arr����(���״���.����סԺ���֧��) & mstrSplit & _
        IIf(str�������� = "", 1, 0) & mstrSplit & IIf(str�������� = "", "", str��������)

    TransationHosp = strReturn
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\TransationHosp)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\TransationHosp)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveDeal(Optional ByVal bln���� As Boolean = True, Optional ByVal bln���� As Boolean = False) As Boolean
    '----------------------------------------------------------------
    '��������   �����潻�װ汾��Ϣ
    '��д��     ������
    '��д����   ��2004-07-06
    '����˵��   ��bln����=TRUE�����,FALSE��סԺ��
    '����:
    '    �ܽ�� = ���ⲡҽ���� + ���ⲡҽ����
    '    ͳ�����֧�� = ͳ��֧��
    '    ���ҽ��/����Ա�������� = ���֧��
    '    �����Ը����Ը�1 = ���ⲡҽ���� - ͳ��֧�� - ���֧�����Ը�2 = �Ը�2
    '    �����Է� = ���ⲡҽ����
    '    ͳ��ⶥ��ҽ���ڽ�� = ͳ��ⶥ��ҽ���ڽ��
    'סԺ:
    '    �ܽ�� = ���η����ܽ��
    '    ͳ�����֧�� = ����ͳ��֧��
    '    ���ҽ��/����Ա�������� = ���δ��֧��
    '    �����Ը����Ը�1 = ���η����ܽ�� - ����ͳ��֧�� - ���δ��֧�����Ը�2 = �����Ը�2
    '    �����Է� = ����ҽ����
    '    ͳ��ⶥ��ҽ���ڽ�� = ����ͳ��ⶥ��ҽ���ڽ��
    '----------------------------------------------------------------
    Dim rsHead As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim str��ʼ���� As String, str��ֹ���� As String, str�������� As String, str��Ժ���� As String
    Dim dbl�����ܶ� As Double, dblͳ��֧�� As Double, dbl���֧�� As Double
    Dim dbl�����Ը� As Double, dbl�����Է� As Double, dbl�����Ը� As Double, dblͳ��ⶥ��ҽ���� As Double
    Dim strҽԺ���� As String, strҽԺ�ȼ� As String
    Dim strFields As String, strValues As String
    Dim strTotal As String
    Dim strFile_Out As String
    
    Dim arrRow
    Dim lngRow As Long, lngRows As Long
    Dim objStream As TextStream
    Dim objFileSys As New FileSystemObject
    
    Const str���� As String = "Spec_Divide.out"
    Const strסԺ As String = "Hosp_Divide.out"
    Const col������_�ֶ����� As Integer = 0
    Const col�������ⲡ_��¼�� As Integer = 0
    Const col�������ⲡ_������ˮ�� As Integer = 1
    Const col�������ⲡ_�����ܶ� As Integer = 2
    Const col�������ⲡ_��ͨ����ҽ���� As Integer = 3
    Const col�������ⲡ_��ͨ����ҽ���� As Integer = 4
    Const col�������ⲡ_ͳ��֧�� As Integer = 5
    Const col�������ⲡ_ͳ���Ը� As Integer = 6
    Const col�������ⲡ_���֧�� As Integer = 7
    Const col�������ⲡ_����Ը� As Integer = 8
    Const col�������ⲡ_���ⲡ�Ը� As Integer = 9
    Const col�������ⲡ_���ⲡҽ���� As Integer = 10
    Const col�������ⲡ_�����Ը� As Integer = 11
    Const col�������ⲡ_ͳ��ⶥ��ҽ���� As Integer = 12
    Const colסԺ_������� As Integer = 0
    Const colסԺ_������� As Integer = 1
    Const colסԺ_��ʼ���� As Integer = 2
    Const colסԺ_��ֹ���� As Integer = 3
    Const colסԺ_�����ܶ� As Integer = 4
    Const colסԺ_ҽ���� As Integer = 5
    Const colסԺ_ͳ��֧�� As Integer = 6
    Const colסԺ_ͳ���Ը� As Integer = 7
    Const colסԺ_���֧�� As Integer = 8
    Const colסԺ_����Ը� As Integer = 9
    Const colסԺ_�����Ը� As Integer = 10
    Const colסԺ_����Ӧ�� As Integer = 11
    On Error GoTo errHand
    
    str�������� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    '��ȡҽԺ���ƣ�������д�ֲ����Ѽ�¼
    Call DebugTool("��ȡҽԺ���Ƽ�ҽԺ�ȼ�(zl9Insure\SaveDeal")
    gstrSQL = "Select A.ҽԺ����,A.ҽԺ����,B.���� AS ҽԺ�ȼ� " & _
            " From ҽԺ�ȼ� A,(Select B.����,B.���� From ָ������ A,ָ����ϵ���ձ� B Where A.���=B.��� And A.����='ҽԺ�ȼ�') B" & _
            " Where A.ҽԺ�ȼ�=B.���� And A.ҽԺ����='" & gComInfo_����.ҽԺ���� & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\SaveDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp.RecordCount = 0 Then
        MsgBox "ҽԺ��������ҽԺ�ȼ��嵥��û�ж�Ӧ��ҽԺ��Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    strҽԺ���� = rsTemp!ҽԺ���� & vbCrLf & rsTemp!ҽԺ����
    strҽԺ�ȼ� = rsTemp!ҽԺ�ȼ�
    
    If Not bln���� Then
        '��ȡ�α�����Ժ����
        Call DebugTool("��ȡ�α�����Ժ����(zl9Insure\SaveDeal")
        gstrSQL = "Select B.���� AS ��Ժ���" & _
                " From �����ʻ� A,(Select B.����,B.���� From ָ������ A,ָ����ϵ���ձ� B Where A.���=B.��� And A.����='��Ժ���') B" & _
                " Where A.��Ժ���=B.���� And A.����='" & gComInfo_����.���� & "'"
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "ZL9INSURE\SaveDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
        str��Ժ���� = rsTemp!��Ժ���
    End If
    
    '��ʼ����¼��
    If bln���� Then
        strFile_Out = str����
        strFields = "ҽԺ����," & adLongVarChar & ",100|" & _
                    "��������," & adLongVarChar & ",100|" & _
                    "ҽԺ����," & adLongVarChar & ",100|" & _
                    "�������," & adLongVarChar & ",100"
        Call Record_Init(rsHead, strFields)
        strFields = "�����ܶ�," & adLongVarChar & ",100|" & _
                    "ͳ��֧��," & adLongVarChar & ",100|" & _
                    "���֧��," & adLongVarChar & ",100|" & _
                    "�����Ը�," & adLongVarChar & ",100|" & _
                    "�����Ը�," & adLongVarChar & ",100|" & _
                    "�����Է�," & adLongVarChar & ",100|" & _
                    "ͳ��ⶥ��ҽ���ڽ��," & adLongVarChar & ",100|" & _
                    "��ʼ����," & adLongVarChar & ",100|" & _
                    "��ֹ����," & adLongVarChar & ",100|" & _
                    "��������," & adLongVarChar & ",100"
        Call Record_Init(rsDeal, strFields)
    Else
        strFile_Out = strסԺ
        strFields = "ҽԺ����," & adLongVarChar & ",100|" & _
                    "��Ժ����," & adLongVarChar & ",100|" & _
                    "ҽԺ����," & adLongVarChar & ",100|" & _
                    "�������," & adLongVarChar & ",100|" & _
                    "��Ժ����," & adLongVarChar & ",100|" & _
                    "ת������," & adLongVarChar & ",100"
        Call Record_Init(rsHead, strFields)
        strFields = "�����ܶ�," & adLongVarChar & ",100|" & _
                    "ͳ��֧��," & adLongVarChar & ",100|" & _
                    "���֧��," & adLongVarChar & ",100|" & _
                    "�����Ը�," & adLongVarChar & ",100|" & _
                    "�����Ը�," & adLongVarChar & ",100|" & _
                    "�����Է�," & adLongVarChar & ",100|" & _
                    "ͳ��ⶥ��ҽ���ڽ��," & adLongVarChar & ",100|" & _
                    "��ʼ����," & adLongVarChar & ",100|" & _
                    "��ֹ����," & adLongVarChar & ",100|" & _
                    "��������," & adLongVarChar & ",100"
        Call Record_Init(rsDeal, strFields)
    End If
    
    '׼�����������ļ������������¼�����Ա���ʾ������Ϣ
    Call DebugTool("׼�����������ļ������������¼�����Ա���ʾ������Ϣ(zl9Insure\SaveDeal)")
    Call WriteBusinessLOG("׼�����������ļ������������¼�����Ա���ʾ������Ϣ(zl9Insure\SaveDeal)", "", "")
    If Not objFileSys.FileExists(gComInfo_����.����Ŀ¼ & "\" & strFile_Out) Then Exit Function
    Set objStream = objFileSys.OpenTextFile(gComInfo_����.����Ŀ¼ & "\" & strFile_Out)
    
    Call DebugTool("��ȡ����������(zl9Insure\SaveDeal)")
    Call WriteBusinessLOG("��ȡ����������(zl9Insure\SaveDeal)", "", "")
    '����ÿ���ı����Ի��з���������VB�ϵ��ǻس����У����������ݶ��������ˣ���ҪSPLIT
    strTotal = objStream.ReadLine
    arrRow = Split(strTotal, vbCr)
    objStream.Close
    Set objStream = Nothing
    
    If bln���� Then
        
        strFields = "ҽԺ����|��������|ҽԺ����|�������"
        strValues = strҽԺ���� & "|" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd") & "|" & strҽԺ�ȼ� & "|"
        Call Record_Add(rsHead, strFields, strValues)
        
        '����������м���
        dbl�����ܶ� = Val(Split(arrRow(0), "|")(col�������ⲡ_�����ܶ�)) - Val(Split(arrRow(0), "|")(col�������ⲡ_��ͨ����ҽ����)) - Val(Split(arrRow(0), "|")(col�������ⲡ_��ͨ����ҽ����))
        dbl���֧�� = Val(Split(arrRow(0), "|")(col�������ⲡ_���֧��))
        dblͳ��֧�� = Val(Split(arrRow(0), "|")(col�������ⲡ_ͳ��֧��))
        dblͳ��ⶥ��ҽ���� = Val(Split(arrRow(0), "|")(col�������ⲡ_ͳ��ⶥ��ҽ����))
        dbl�����Է� = Val(Split(arrRow(0), "|")(col�������ⲡ_���ⲡҽ����))
        '�����Ը�=���ⲡҽ����-ͳ��֧��-���֧��
        dbl�����Ը� = (dbl�����ܶ� - dbl�����Է�) - dbl���֧�� - dblͳ��֧��
        dbl�����Ը� = Val(Split(arrRow(0), "|")(col�������ⲡ_�����Ը�))
        
        strFields = "�����ܶ�|ͳ��֧��|���֧��|�����Ը�|�����Ը�|�����Է�|ͳ��ⶥ��ҽ���ڽ��|��ʼ����|��ֹ����|��������"
        strValues = dbl�����ܶ� & "|" & dblͳ��֧�� & "|" & dbl���֧�� & "|" & dbl�����Ը� & "|" & dbl�����Ը� & _
            "|" & dbl�����Է� & "|" & dblͳ��ⶥ��ҽ���� & "|" & str�������� & "|" & str�������� & "|" & str��������
        Call Record_Add(rsDeal, strFields, strValues)
    Else
        lngRows = Val(Split(arrRow(0), "|")(col������_�ֶ�����))
        '��Ҫ����ֶηֽ���ϸ��
        For lngRow = 1 To lngRows
            strFields = "ҽԺ����|��Ժ����|ҽԺ����|�������|��Ժ����|ת������"
            str��ʼ���� = Split(arrRow(lngRow), "|")(colסԺ_��ʼ����)
            str��ֹ���� = Split(arrRow(lngRow), "|")(colסԺ_��ֹ����)
            strValues = strҽԺ���� & "|" & str��ʼ���� & "|" & strҽԺ�ȼ� & "||" & str��Ժ���� & "|" & str��ֹ����
            Call Record_Add(rsHead, strFields, strValues)
            
            dbl�����ܶ� = Val(Split(arrRow(lngRow), "|")(colסԺ_�����ܶ�))
            dbl���֧�� = Val(Split(arrRow(lngRow), "|")(colסԺ_���֧��))
            dblͳ��֧�� = Val(Split(arrRow(lngRow), "|")(colסԺ_ͳ��֧��))
            If lngRow = lngRows Then
                dblͳ��ⶥ��ҽ���� = Val(Split(arrRow(0), "|")(11))
            Else
                dblͳ��ⶥ��ҽ���� = 0
            End If
            dbl�����Է� = dbl�����ܶ� - Val(Split(arrRow(lngRow), "|")(colסԺ_ҽ����))
            dbl�����Ը� = Val(Split(arrRow(lngRow), "|")(colסԺ_ҽ����)) - dbl���֧�� - dblͳ��֧��
            dbl�����Ը� = Val(Split(arrRow(lngRow), "|")(colסԺ_�����Ը�))
            
            strFields = "�����ܶ�|ͳ��֧��|���֧��|�����Ը�|�����Ը�|�����Է�|ͳ��ⶥ��ҽ���ڽ��|��ʼ����|��ֹ����|��������"
            strValues = dbl�����ܶ� & "|" & dblͳ��֧�� & "|" & dbl���֧�� & "|" & dbl�����Ը� & "|" & dbl�����Ը� & _
                "|" & dbl�����Է� & "|" & dblͳ��ⶥ��ҽ���� & "|" & str��ʼ���� & "|" & str��ֹ���� & "|" & str��������
            Call Record_Add(rsDeal, strFields, strValues)
        Next
    End If
    
    If Not bln���� Then
        '��ʾ��ϸ���ֲ�������Ϣ
        Call frm�ֲ�����.ShowBalance(rsHead, rsDeal, bln����)
        SaveDeal = True
        Exit Function
    End If
    
    '��ȡ���ν��������Ϣ
    If bln���� Then
        Call DebugTool("��ȡ�������������Ϣ(zl9Insure\SaveDeal")
        Call WriteBusinessLOG("��ȡ�������������Ϣ(zl9Insure\SaveDeal", "", "")
        gstrSQL = "" & _
            " SELECT A.�籣֤��,A.ҵ������,A.�α����,A.����Ա,A.����Ա����,A.���ֱ�ʶ," & _
            "        TO_CHAR(A.���ⲡ��ֹ����,'yyyyMMdd') AS ���ⲡ��ֹ����," & _
            "        0 AS ��Ժ����,0 AS ��Ժ����," & _
            "        to_Char(sysdate,'yyyy-MM-dd hh24:mi:ss') As ��Ժ����," & _
            "        to_Char(sysdate,'yyyy-MM-dd hh24:mi:ss') AS ��Ժ����" & _
            " From �����ʻ� A,���ｻ����Ϣ C" & _
            " WHERE A.����=C.���� And C.������ˮ��='" & gComInfo_����.������ˮ�� & "'"
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "ZL9INSURE\SaveDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    Else
        Call DebugTool("��ȡ����סԺ������Ϣ(zl9Insure\SaveDeal")
        Call WriteBusinessLOG("��ȡ����סԺ������Ϣ(zl9Insure\SaveDeal", "", "")
        gstrSQL = "" & _
            " SELECT A.�籣֤��,A.ҵ������,A.�α����,A.����Ա,A.����Ա����,A.���ֱ�ʶ," & _
            "        TO_CHAR(A.���ⲡ��ֹ����,'yyyyMMdd') AS ���ⲡ��ֹ����," & _
            "        B.��Ժ����,B.��Ժ��ʽ,C.��Ժ����," & _
            "        to_Char(C.��Ժ����,'yyyy-MM-dd hh24:mi:ss') As ��Ժ����," & _
            "        to_Char(C.��Ժ����,'yyyy-MM-dd hh24:mi:ss') AS ��Ժ����" & _
            " From �����ʻ� A,��Ժ��Ϣ B,סԺ������Ϣ C" & _
            " WHERE A.����=B.���� And B.����=C.���� And C.������ˮ��='" & gComInfo_����.������ˮ�� & "'"
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "ZL9INSURE\SaveDeal", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    End If
    
    '��������
    '����,ҽ�ƻ���,ҽ�����,��Ժ����,��Ժ����,��Ժ����,��Ժ����,ͳ��֧��
    '���֧��,�����Ը�,�����Է�,ͳ��ⶥ��ҽ����,������ˮ��,�����ʷ��¼
    With rsDeal
        If .RecordCount <> 0 Then .MoveFirst
        
        Do While Not .EOF
            gstrSQL = "ZL_�ֲ����Ѽ�¼_INSERT(" & _
                      "'" & gComInfo_����.���� & "','" & strҽԺ���� & "','" & rsTemp!ҵ������ & "'," & _
                      "'" & rsTemp!��Ժ���� & "',TO_DATE('" & !��ʼ���� & "','yyyy-MM-dd hh24:mi:ss')," & _
                      "'" & IIf(.AbsolutePosition <> .RecordCount, "2", "0") & "',TO_DATE('" & !��ֹ���� & "','yyyy-MM-dd hh24:mi:ss')," & _
                      "" & Val(!ͳ��֧��) & "," & Val(!���֧��) & "," & Val(!�����Ը�) & "," & Val(!�����Է�) & "," & _
                      "" & Val(!ͳ��ⶥ��ҽ���ڽ��) & ",'" & gComInfo_����.������ˮ�� & "',0)"
            gcnBJYB.Execute gstrSQL, , adCmdStoredProc
            .MoveNext
        Loop
    End With
    
    SaveDeal = True
    Exit Function
errHand:
    Call DebugTool("(zl9INSURE\SaveDeal)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\SaveDeal)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, mstrSplit)
    arrValues = Split(strValues, mstrSplit)
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, mstrSplit)
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub










'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'�����ǻ����ӿ�
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
Public Function CheckTradeName(ByVal lng�շ�ϸĿID As Long, ByVal strҽ������ As String) As Boolean
    Dim strUser As String
    Dim rsTemp As New ADODB.Recordset
    '�����Ʒ��������Ƿ���ҽ���涨����Ʒ��������У�����ǲ��������ö���
    On Error GoTo errHand
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=[1] And ������='ҽ���û���'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�м����û���", TYPE_����)
    strUser = rsTemp!����ֵ
    
    gstrSQL = " Select 1 From " & strUser & ".ҩƷ���� " & _
              " Where ����='" & strҽ������ & "'" & _
              " And ���� IN ( " & _
              "     Select ���� From ҩƷ���� Where ҩ��ID= " & _
              "         (Select ҩ��ID From ҩƷĿ¼ Where ҩƷID=[1]))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", lng�շ�ϸĿID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "ҽ�������·���ҩƷ������û�д���Ŀ����Ʒ���ͱ�����������ѡ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckTradeName = True
errHand:
    Exit Function
End Function

Private Sub GetSequence_����(ByVal lng����ID As Long)
    Dim strText As String
    Dim str����ʱ�� As String
    Dim str��ǰʱ�� As String
    Dim strSequence As String
    Dim rsTemp As New ADODB.Recordset
    Dim intDO As Integer, intCOUNT As Integer, intPos As Integer
    '����ǰʱ�䡢����ʱ����д���ת��ΪΨһ����ˮ�ű�ʶ
    '���˼·�����ꡢ�¡��ա�ʱ���֡��붼ת��Ϊһ����ĸ����ʽ��ʾ����Ϊһ��ֻ��20λ����ǰ��8λҽԺ������ǰ���ַ�
    intCOUNT = 6
    intPos = 1
    str��ǰʱ�� = Format(zlDatabase.Currentdate, "yyMMddHHmmss")
    
    Call DebugTool("׼������˳���(zl9Insure\GetSequenct_����)")
    Call WriteBusinessLOG("׼������˳���(zl9Insure\GetSequenct_����)", "", "")
    '��ȡ�ò��˵ľ���ʱ��
    gstrSQL = "Select to_char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') As ����ʱ�� From �����ʻ�" & _
        " Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˵ľ���ʱ��", TYPE_����, lng����ID)
    '�϶�����Ϊ��
    str����ʱ�� = Format(rsTemp!����ʱ��, "yyMMddHHmmss")
    
    For intDO = 1 To intCOUNT
        strText = Mid(str����ʱ��, intPos, 2)
        intPos = intPos + 2
        strSequence = strSequence & Chr(asc("0") + Val(strText))
    Next
    intPos = 1
    For intDO = 1 To intCOUNT
        strText = Mid(str��ǰʱ��, intPos, 2)
        intPos = intPos + 2
        strSequence = strSequence & Chr(asc("0") + Val(strText))
    Next
    gComInfo_����.������ˮ�� = gComInfo_����.ҽԺ���� & strSequence
    Call DebugTool("˳���:" & gComInfo_����.������ˮ�� & "(zl9Insure\GetSequenct_����)")
    Call WriteBusinessLOG("˳���:" & gComInfo_����.������ˮ�� & "(zl9Insure\GetSequenct_����)", "", "")
End Sub

Public Function �������_����(ByVal StrInput As String, Optional ByVal bln���� As Boolean = True) As Boolean
    �������_���� = CheckBlockage(StrInput, bln����)
End Function

Public Function ��ݱ�ʶ_����(ByVal bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim strReturn As String
    strReturn = frmIdentify����.GetIdentify(bytType, lng����ID)
    ��ݱ�ʶ_���� = strReturn
End Function

Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.��������()
End Function

Public Function ҽ����ʼ��_����(Optional ByVal blnTest As Boolean = False) As Boolean
    Dim strUser As String, strServer As String, strPass As String
    Dim rsTemp As New ADODB.Recordset
    Dim strReturn As String
    On Error GoTo errHand
    
    If mblnInit = False Then
        '��������ҽ��������������
        gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_����)
        
        Do Until rsTemp.EOF
            Select Case rsTemp("������")
                Case "ҽ���û���"
                    strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                Case "ҽ��������"
                    strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                Case "ҽ���û�����"
                    strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
                Case "���Ŀ¼"
                    gComInfo_����.���Ŀ¼ = Nvl(rsTemp!����ֵ)
                Case "����Ŀ¼"
                    gComInfo_����.����Ŀ¼ = Nvl(rsTemp!����ֵ)
            End Select
            rsTemp.MoveNext
        Loop
        
        If OraDataOpen(gcnBJYB, strServer, strUser, strPass, False) = False Then
            MsgBox "�޷����ӵ��м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Not blnTest Then
            If Nvl(gComInfo_����.���Ŀ¼) = "" Then
                MsgBox "��Ϊ��ҽ����������ļ�Ŀ¼��", vbInformation, gstrSysName
                Exit Function
            End If
            If Nvl(gComInfo_����.����Ŀ¼) = "" Then
                MsgBox "��Ϊ��ҽ�����ó����ļ�Ŀ¼��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        'ȡҽԺ����
        gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_����)
        gComInfo_����.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
        
        '�������߻�����
        If Not ���ýӿ�_����(�������߻�����, strReturn) Then Exit Function
    End If
    
    mblnInit = True
    ҽ����ʼ��_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ֹ_����() As Boolean
    Dim strReturn As String
    Call ���ýӿ�_����(ֹͣ���߻�����, strReturn)
End Function

Private Function ���ýӿ�_����(ByVal int���� As �ӿڹ���, ByRef str���� As String) As Boolean
    '----------------------------------------------------------------
    '��������   �����ýӿں���
    '��д��     ������
    '��д����   ��2004-07-03
    '����˵��   ��str����=��Σ���ֵ����Ϊ�գ�����������ڶ����Σ�һ��ֻ��������Σ�����"||"�ָ���Ŀǰ�޶����εĺ���
    '                     ���Σ���ֵ����Ϊ�գ�����������ڳ��Σ��򷵻س��ε�ֵ����"||�ָ���Ŀǰ����ȡ�汾�ӿ�Ҫ���ض������
    '----------------------------------------------------------------
    Dim lngReturn As Long           '��������ֵ
    Dim strFunction As String       '��ǰִ�к�����
    Dim strInPara As String
    Dim strOutPara1 As String * 2000      '����
    Dim strOutPara2 As String * 2000      '����
    Dim strErrMsg As String * 255
    Dim arrPara                     '�������
    On Error GoTo errHand
    
    strInPara = str����
    Select Case int����
    Case �ӿڹ���.�������߻�����
        strFunction = "StartPolicy[INTERFACE]"
        '�������ʧ���ˣ���Ҫ����ֹͣ���߻�����
        lngReturn = BJ_StartPolicy(0)
        Call WriteBusinessLOG(strFunction, "", "")
        If lngReturn <> 0 Then
            Call BJ_StopPolicy
            Call WriteBusinessLOG("StopPolicy[INTERFACE]", "", "")
        End If
        lngReturn = 0
    Case �ӿڹ���.ֹͣ���߻�����
        strFunction = "StopPolicy[INTERFACE]"
        Call BJ_StopPolicy
        Call WriteBusinessLOG(strFunction, "", "")
    Case �ӿڹ���.��ȡ�ֿ�������Ϣ
        strFunction = "GetPersonCommInfo[INTERFACE]"
        lngReturn = BJ_GetPersonCommInfo(str����)
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.��ȡ�ֿ����˴�����Ϣ
        strFunction = "GetSumInfo[INTERFACE]"
        lngReturn = BJ_Get_SumInfo(str����)
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ
        strFunction = "GetSumInfo2[INTERFACE]"
        lngReturn = BJ_Get_SumInfo2(str����, strOutPara1)
        str���� = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.��ȡ���ⲡ��Ϣ
        strFunction = "GetSpecInfo[INTERFACE]"
        lngReturn = BJ_Get_SpecInfo(str����, strOutPara1)
        str���� = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.�ʻ�֧��
        strFunction = "RegAccount[INTERFACE]"
        lngReturn = BJ_Reg_Account(str����, strOutPara1)
        str���� = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.����ȷ��
        strFunction = "Reg[INTERFACE]"
        lngReturn = BJ_Reg(str����, strOutPara1)
        str���� = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.�ӿڰ汾��Ϣ
        strFunction = "GetVer[INTERFACE]"
        lngReturn = BJ_Get_Ver(strOutPara1, strOutPara2)
        str���� = TrimStr(strOutPara1) & "||" & TrimStr(strOutPara2)
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_ͨ��1
        strFunction = "Divide[INTERFACE]"
        lngReturn = BJ_Divide(str����, strOutPara1)
        str���� = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_ͨ��2
        strFunction = "Divide2[INTERFACE]"
        lngReturn = BJ_Divide2(str����, strOutPara1)
        str���� = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_��ͨ����
        strFunction = "Poli_Divide[INTERFACE]"
        lngReturn = BJ_Poli_Divide(str����)
        str���� = ""
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_��������1
        strFunction = "Spec_Divide[INTERFACE]"
        lngReturn = BJ_Spec_Divide(str����)
        str���� = ""
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_��������2
        strFunction = "Spec_Divide2[INTERFACE]"
        lngReturn = BJ_Spec_Divide2(str����, strOutPara1)
        str���� = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_��ͥ����1
        strFunction = "Home_Divide1[INTERFACE]"
        lngReturn = BJ_Home_Divide(str����)
        str���� = ""
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_��ͥ����2
        strFunction = "Home_Divide2[INTERFACE]"
        lngReturn = BJ_Home_Divide2(str����, strOutPara1)
        str���� = strOutPara1
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_סԺ1
        strFunction = "Hosp_Divide1[INTERFACE]"
        lngReturn = BJ_Hosp_Divide(str����)
        str���� = ""
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_סԺ2
        strFunction = "Hosp_Divide2[INTERFACE]"
        lngReturn = BJ_Hosp_Divide2(str����)
        str���� = ""
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    Case �ӿڹ���.���÷ֽ�_סԺ3
        strFunction = "Hosp_Divide3[INTERFACE]"
        lngReturn = BJ_Hosp_Divide3(str����)
        str���� = ""
        Call WriteBusinessLOG(strFunction, strInPara, str����)
    End Select
    
    '��ʾ������
    If lngReturn <> 0 Then
        lngReturn = BJ_Get_ErrInfo(strErrMsg)
        If lngReturn <> 0 Then
            strErrMsg = "ִ�к���[" & strFunction & "]ʱ����δ֪����"
        Else
            strErrMsg = "ִ�к���[" & strFunction & "]ʱ�������´���" & vbCrLf & strErrMsg
        End If
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    str���� = TrimStr(str����)
    ���ýӿ�_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub WriteBusinessLOG(ByVal strFunc As String, ByVal StrInput As String, ByVal strOutput As String)
    '���±������ڼ�¼���ýӿڵ����
    Const strFile As String = "C:\Business_"
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    If gintDebug = -1 Then gintDebug = Val(GetSetting("ZLSOFT", "ҽ��", "����", 0))
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    If gintDebug = 0 Then Exit Sub
    strFileName = strFile & Format(Date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (String(50, "-"))
    objStream.WriteLine ("ִ��ʱ��:" & strDate)
    objStream.WriteLine ("������:" & strFunc)
    objStream.WriteLine ("  ���:" & StrInput)
    objStream.WriteLine ("  ����:" & strOutput)
    objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
End Sub

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    Dim int��Ŀ��� As Integer  '0-ҩƷ 1-������Ŀ 2-������ʩ
    Dim lng����ID As Long
    Dim str�������� As String, str���÷������� As String
    Dim strҽ������ As String, strHIS��Ŀ���� As String, str����Ա��ʶ As String, strReturn As String
    Dim str���� As String, str���÷ֽ���� As String
    Dim strFields As String, strValues As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim arr���㷽ʽ
    Const intҽ������ As Integer = 5
    Const int�󲡲��� As Integer = 7
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    '����˵����Ŀǰ��֧�����ⲡ����
    On Error GoTo errHandle
    lng����ID = rs��ϸ!����ID
    str�������� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    str���÷������� = Format(zlDatabase.Currentdate(), "yyyyMMdd")
    
    '�Ȼ�ȡ���ν�����ˮ��
    Call DebugTool("������ˮ��(zl9Insure\�����������)")
    Call WriteBusinessLOG("������ˮ��(zl9Insure\�����������)", "", "")
    Call GetSequence_����(lng����ID)
    
    '��ȡ�α��˹���Ա��ʶ
    Call DebugTool("��ȡ�α��˹���Ա��ʶ(zl9Insure\�����������)")
    Call WriteBusinessLOG("��ȡ�α��˹���Ա��ʶ(zl9Insure\�����������)", "", "")
    gstrSQL = "Select ����Ա From �����ʻ� Where ����='" & gComInfo_����.���� & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\�����������_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str����Ա��ʶ = rsTemp!����Ա
    
    '��ȡ�˿̵Ĵ�����Ϣ�����÷ֽ⺯������Σ�
    Call DebugTool("��ȡ�˿̵���ʷ���Ѽ�¼�����÷ֽ⺯������Σ�(zl9Insure\�����������)")
    Call WriteBusinessLOG("��ȡ�˿̵���ʷ���Ѽ�¼�����÷ֽ⺯������Σ�(zl9Insure\�����������)", "", "")
    Set rsDeal = GetDeal(lng����ID, str��������)
    '����¼�����ʷ���Ѽ�¼����������ļ�
    Call DebugTool("����¼�����ʷ���Ѽ�¼����������ļ�(zl9Insure\�����������)")
    Call WriteBusinessLOG("����¼�����ʷ���Ѽ�¼����������ļ�(zl9Insure\�����������)", "", "")
    If Not MakeFile_Center(rsDeal, �ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ) Then Exit Function
    '�õ��ӿڷ��صĴ�����
    strReturn = gComInfo_����.���� & mstrSplit & str����Ա��ʶ
    Call DebugTool("���û�ȡ������Ϣ�ӿ�(zl9Insure\�����������)")
    Call WriteBusinessLOG("���û�ȡ������Ϣ�ӿ�(zl9Insure\�����������)", "", "")
    If Not ���ýӿ�_����(�ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ, strReturn) Then Exit Function   'strReturn�����ݽ����ڷ��÷ֽ⣬���²��ܽ��и�ֵ
    str���� = strReturn
    
    '��������ϸ������ϸ�ļ���Ҳ�Ƿ��÷ֽ⺯������Σ�
'    ----�����ļ�˵��
'    ���    ������  ����    ��󳤶�    ˵��
'    1   ��Ŀ���    C   9   ˳���
'    2   ������  C   20  �μ���׼AKC220���ɿ�
'    3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'    4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'    5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'    6   ����    N   10,4    AKC225
'    7   ����    N   8,2 AKC226
'    8   �����ܽ��  N   10,4    ʵ�ʽ�����
'    9   ���÷�������    D   8   YYYYMMDD
    strFields = "��Ŀ���," & adLongVarChar & ",9" & mstrSplit & _
                "������," & adLongVarChar & ",20" & mstrSplit & _
                "��Ŀ����," & adLongVarChar & ",20" & mstrSplit & _
                "��Ŀ����," & adLongVarChar & ",100" & mstrSplit & _
                "��Ŀ���," & adLongVarChar & ",3" & mstrSplit & _
                "����," & adDouble & ",18" & mstrSplit & _
                "����," & adDouble & ",18" & mstrSplit & _
                "�����ܽ��," & adDouble & ",18" & mstrSplit & _
                "���÷�������," & adLongVarChar & ",20"
    Call Record_Init(rsCharge, strFields)
    
    '�õ�����¼��ķ�����ϸ�����ҽ����Ϣ��������w01��ͷ����ʾ������ʩ����Ŀ��
    Call DebugTool("����������ϸ��¼��(zl9Insure\�����������)")
    Call WriteBusinessLOG("����������ϸ��¼��(zl9Insure\�����������)", "", "")
    strFields = "��Ŀ���|������|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|���÷�������"
    If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
    Do Until rs��ϸ.EOF
        '��ҩƷ��ͨ�����ƻ��ҩƷ��Ŀ������ȡ����
        gstrSQL = "Select A.��Ŀ���� As ҽ������,C.ͨ������ AS ��Ŀ���� " & _
            " From ����֧����Ŀ A,ҩƷĿ¼ B,ҩƷ��Ϣ C" & _
            " Where A.����(+)=[1] And A.�շ�ϸĿID(+)=B.ҩƷID And B.ҩ��ID=C.ҩ��ID " & _
            " AND B.ҩƷID=[2]" & _
            " UNION " & _
            " Select A.��Ŀ���� As ҽ������,B.���� AS ��Ŀ����" & _
            " From ����֧����Ŀ A,�շ�ϸĿ B" & _
            " Where A.����(+)=[1] AND B.ID=[2]" & _
            " And A.�շ�ϸĿID(+)=B.ID AND B.��� Not In ('5','6','7')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", TYPE_����, CLng(rs��ϸ!�շ�ϸĿID))
        If rsTemp.EOF Then
            MsgBox "������Ŀû�к�ҽ����Ŀ���ö��չ�ϵ��[������Ŀ]", vbInformation, gstrSysName
            Exit Function
        End If
        strҽ������ = Nvl(rsTemp!ҽ������, 0)
        strHIS��Ŀ���� = rsTemp!��Ŀ����
        
        If InStr(1, "5,6,7", rs��ϸ!�շ����) <> 0 Then
            int��Ŀ��� = 0 '0-ҩƷ
        Else
            int��Ŀ��� = (IIf(strҽ������ Like "w01*", 2, 1)) '1-������Ŀ 2-������ʩ
        End If
        
        '����������ϸ��¼��,�Թ���������ļ�
        strValues = rs��ϸ.AbsolutePosition & mstrSplit & "" & mstrSplit & _
            strҽ������ & mstrSplit & strHIS��Ŀ���� & mstrSplit & _
            int��Ŀ��� & mstrSplit & Format(rs��ϸ!����, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs��ϸ!����, "#####0.00;-#####0.00;0;") & mstrSplit & Format(rs��ϸ!ʵ�ս��, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            str���÷�������
        Call Record_Add(rsCharge, strFields, strValues)
        
        rs��ϸ.MoveNext
    Loop
    
    '����������ϸ�ļ�
    Call DebugTool("���ݷ�����ϸ����������ļ�(zl9Insure\�����������)")
    Call WriteBusinessLOG("���ݷ�����ϸ����������ļ�(zl9Insure\�����������)", "", "")
    If Not MakeFile_Center(rsCharge, ���÷ֽ�_��������1) Then Exit Function
    
    '�����صĴ�����Ϣת��Ϊ���÷ֽ�����ĸ�ʽ
    str���÷ֽ���� = TransationSpec(str����)
    
    '���÷��÷ֽ⺯����strReturn���Ǵ�����Ϣ��
    Call DebugTool("���÷�����ϸ�ֽ⺯��(zl9Insure\�����������)")
    Call WriteBusinessLOG("���÷�����ϸ�ֽ⺯��(zl9Insure\�����������)", "", "")
    strReturn = str���÷ֽ����
    If Not ���ýӿ�_����(���÷ֽ�_��������1, strReturn) Then Exit Function
    strReturn = AnalyFile_��������1(True)
    If strReturn = "" Then Exit Function
    
    If Not SaveDeal(True, False) Then Exit Function
    
    '��֯�ɽ��㷽ʽ��������
    Call DebugTool("�������÷ֽⷵ�صĻ������ݣ���������Ϣ���ظ���������(zl9Insure\�����������)")
    Call WriteBusinessLOG("�������÷ֽⷵ�صĻ������ݣ���������Ϣ���ظ���������(zl9Insure\�����������)", "", "")
    arr���㷽ʽ = Split(strReturn, mstrSplit)
    str���㷽ʽ = mstrSplit & "ͳ��֧��;" & arr���㷽ʽ(intҽ������) & ";0"
    str���㷽ʽ = str���㷽ʽ & mstrSplit & "���֧��;" & arr���㷽ʽ(int�󲡲���) & ";0"
    str���㷽ʽ = Mid(str���㷽ʽ, 2)
    
    �����������_���� = True
    Exit Function
errHandle:
    Call DebugTool("(zl9INSURE\�����������_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\�����������_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, _
    Optional ByVal blnסԺ As Boolean = False) As Boolean
    Dim lng����ID As Long
    Dim blnTrans As Boolean
    Dim str�������� As String, str���÷������� As String
    Dim str����Ա��ʶ As String, strReturn As String
    Dim str���� As String, str���÷ֽ���� As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim arr���㷽ʽ
    Dim dbl�����ܶ� As Double, dblҽ������ As Double, dbl�󲡲��� As Double, dbl�ֽ� As Double
    Dim dblҽ���� As Double, dbl���ⲡ�����Ը� As Double, dbl���ⲡҽ���� As Double, dbl�����Ը� As Double, dblͳ��ⶥ��ҽ���� As Double
    Const int�����ܶ� As Integer = 2
    Const intҽ���� As Integer = 3
    Const intҽ������ As Integer = 5
    Const int�󲡲��� As Integer = 7
    Const int���ⲡ�����Ը� As Integer = 9
    Const int���ⲡҽ���� As Integer = 10
    Const int�����Ը� As Integer = 11
    Const intͳ��ⶥ��ҽ���� As Integer = 12
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    '���ڽ���������������㣬������е�����ļ��Ѿ��������˹��̲�����Ҫ������ϸ��ֱ����֯һ�¼���
    On Error GoTo errHandle
    '��ȡ����ID
    gstrSQL = "Select ����ID From ������ü�¼ Where ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
    lng����ID = rsTemp!����ID
    str�������� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    str���÷������� = Format(zlDatabase.Currentdate(), "yyyyMMdd")
    
    '��ȡ�α��˹���Ա��ʶ
    Call DebugTool("��ȡ�α��˹���Ա��ʶ(zl9Insure\�������)")
    Call WriteBusinessLOG("��ȡ�α��˹���Ա��ʶ(zl9Insure\�������)", "", "")
    gstrSQL = "Select ����,����Ա From �����ʻ� Where ����ID=" & lng����ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\�������_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    gComInfo_����.���� = rsTemp!����
    str����Ա��ʶ = rsTemp!����Ա
    
    '��ȡ�˿̵Ĵ�����Ϣ�����÷ֽ⺯������Σ�
    Call DebugTool("��ȡ�˿̵���ʷ���Ѽ�¼�����÷ֽ⺯������Σ�(zl9Insure\�������)")
    Call WriteBusinessLOG("��ȡ�˿̵���ʷ���Ѽ�¼�����÷ֽ⺯������Σ�(zl9Insure\�������)", "", "")
    Set rsDeal = GetDeal(lng����ID, str��������)
    '�õ��ӿڷ��صĴ�����������Ԥ����ʱ���Ѳ�����ȡ����������ļ����˴����ٲ������ļ���ֱ�ӵ��ýӿڣ�
    strReturn = gComInfo_����.���� & mstrSplit & str����Ա��ʶ
    Call DebugTool("���û�ȡ������Ϣ�ӿ�(zl9Insure\�������)")
    Call WriteBusinessLOG("���û�ȡ������Ϣ�ӿ�(zl9Insure\�������)", "", "")
    If Not ���ýӿ�_����(�ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ, strReturn) Then  'strReturn�����ݽ����ڷ��÷ֽ⣬���²��ܽ��и�ֵ
        Exit Function
    End If
    str���� = strReturn
    '�����صĴ�����Ϣת��Ϊ���÷ֽ�����ĸ�ʽ
    str���÷ֽ���� = TransationSpec(str����)
    '���÷��÷ֽ⺯����strReturn���Ǵ�����Ϣ��
    Call DebugTool("���÷�����ϸ�ֽ⺯��(zl9Insure\�������)")
    Call WriteBusinessLOG("���÷�����ϸ�ֽ⺯��(zl9Insure\�������)", "", "")
    strReturn = str���÷ֽ����
    '���÷��÷ֽ⺯��������Ԥ����ʱ���Ѳ������÷ֽ⺯��������ļ����˴����ٲ������ļ���ֱ�ӵ��ýӿڣ�
    If Not ���ýӿ�_����(���÷ֽ�_��������1, strReturn) Then
        Exit Function
    End If
    
    '��֯�ɽ��㷽ʽ��������
    Call DebugTool("�������÷ֽⷵ�صĻ������ݣ��Ա㱣�浽���ս����¼��(zl9Insure\�������)")
    Call WriteBusinessLOG("�������÷ֽⷵ�صĻ������ݣ��Ա㱣�浽���ս����¼��(zl9Insure\�������)", "", "")
    '�Ȱ�Ԥ����ģʽ��ȡ�û���������
    strReturn = AnalyFile_��������1(True)
    If strReturn <> "" Then
        arr���㷽ʽ = Split(strReturn, mstrSplit)
        dbl�����ܶ� = Val(arr���㷽ʽ(int�����ܶ�))
        dblҽ������ = Val(arr���㷽ʽ(intҽ������))
        dbl�󲡲��� = Val(arr���㷽ʽ(int�󲡲���))
        
        dblҽ���� = Val(arr���㷽ʽ(intҽ����))
        dbl���ⲡ�����Ը� = Val(arr���㷽ʽ(int���ⲡ�����Ը�))
        dbl���ⲡҽ���� = Val(arr���㷽ʽ(int���ⲡҽ����))
        dbl�����Ը� = Val(arr���㷽ʽ(int�����Ը�))
        dblͳ��ⶥ��ҽ���� = Val(arr���㷽ʽ(intͳ��ⶥ��ҽ����))
    End If
    dbl�ֽ� = dbl�����ܶ� - dblҽ������ - dbl�󲡲���
    
    '*****************�����м�⣬�Ա��ϴ�*****************
    '**���汾�εĽ��״�����Ϣ**
    gcnBJYB.BeginTrans
    blnTrans = True
    Call DebugTool("���汾�εĽ��״�����Ϣ(zl9Insure\�������)")
    Call WriteBusinessLOG("���汾�εĽ��״�����Ϣ(zl9Insure\�������)", "", "")
    If Not SaveBusinessDeal(str����) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**���潻�װ汾��Ϣ**
    Call DebugTool("���汾�εĽ��װ汾��Ϣ(zl9Insure\�������)")
    Call WriteBusinessLOG("���汾�εĽ��װ汾��Ϣ(zl9Insure\�������)", "", "")
    If Not SaveBusinessVersion(True) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**�������ｻ����Ϣ�����������ϸ**
    Call DebugTool("�������ｻ����Ϣ�����������ϸ(zl9Insure\�������)")
    Call WriteBusinessLOG("�������ｻ����Ϣ�����������ϸ(zl9Insure\�������)", "", "")
    strReturn = AnalyFile_��������1(False, lng����ID, blnסԺ)
    If strReturn = "" Then
        '����ʧ��
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    If Not SaveDeal(True, True) Then Exit Function
    
    '**���汣�ս����¼**
    '����ͳ����=����Ա��������;ͳ�ﱨ�����=ͳ�����;���Ը����=�󲡻���
    '֧��˳���=������ˮ��,��ע=ҵ������
    Call DebugTool("���汣�ս����¼(zl9Insure\�������)")
    Call WriteBusinessLOG("���汣�ս����¼(zl9Insure\�������)", "", "")
    gstrSQL = "zl_���ս����¼_insert(" & IIf(blnסԺ, 2, 1) & "," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        dbl�����ܶ� & "," & dbl�ֽ� & ",0,0," & dblҽ������ & "," & dbl�󲡲��� & "," & _
        0 & ",0,'" & gComInfo_����.������ˮ�� & "',null,null,'" & gComInfo_����.ҵ������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
    
    gcnBJYB.CommitTrans
    blnTrans = False
    �������_���� = True
    Exit Function
errHandle:
    Call DebugTool("(zl9Insure\�������_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\�������_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Function

Public Function ����������_����(ByVal lng����ID As Long, ByVal cur�����ʻ� As Currency, ByVal lng����ID As Long) As Boolean
    On Error GoTo errHand
    Dim blnTrans As Boolean
    Dim lng����ID As Long
    Dim str���� As String, strԭ������ˮ�� As String
    Dim str����ʱ�� As String, str�˷�ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    
    Call GetSequence_����(lng����ID)
    
    'ȡ��ԭʼ�����¼�Ľ���ʱ��
    Call DebugTool("ȡ��ԭʼ�����¼�Ľ���ʱ��(zl9Insure\����������)")
    Call WriteBusinessLOG("ȡ��ԭʼ�����¼�Ľ���ʱ��(zl9Insure\����������)", "", "")
    gstrSQL = "Select ����,������ˮ��,����ʱ�� From " & GetUser & ".���״�����Ϣ " & _
            " Where ������ˮ��=(Select ֧��˳��� From ���ս����¼ Where ����=1 AND ��¼ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ԭʼ�����¼�Ľ���ʱ��", lng����ID)
    str���� = rsTemp!����
    strԭ������ˮ�� = rsTemp!������ˮ��
    str����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    str�˷�ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '��飬����ڸý��׺󣬻��������������ף���������г����������˷ѵļ�¼�ſ���
    Call DebugTool("����Ƿ�����һ�ʿ�ʼ�˷�(zl9Insure\����������)")
    Call WriteBusinessLOG("����Ƿ�����һ�ʿ�ʼ�˷�(zl9Insure\����������)", "", "")
    gstrSQL = "Select Count(*) AS Records From ���״�����Ϣ A,�����ʻ� B" & _
            " Where A.����=B.���� And B.����ID=" & lng����ID & _
            " And ����ʱ��>to_Date('" & str����ʱ�� & "','yyyy-MM-dd hh24:mi:ss')" & _
            " And ������ˮ�� Not In (Select ԭ������ˮ�� From �˷���Ϣ Where ����='" & str���� & "')"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\�������_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp!Records > 0 Then
        MsgBox "ҽ���ӿڲ�������м俪ʼ�˷ѣ���ֻ�ܴ����һ��ҵ��ʼ�˷ѣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ������¼�Ľ���ID�����ݺ�
    Call DebugTool("ȡ����ID(zl9Insure\����������)")
    Call WriteBusinessLOG("ȡ����ID(zl9Insure\����������)", "", "")
    gstrSQL = "select distinct A.����ID,A.NO from ������ü�¼ A,������ü�¼ B where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng����ID = rsTemp!����ID
    
    '��ȡԭʼ�ı��ս����¼
    Call DebugTool("���汣�ս����¼(zl9Insure\����������)")
    Call WriteBusinessLOG("���汣�ս����¼(zl9Insure\����������)", "", "")
    gstrSQL = "Select * From ���ս����¼ Where ����=1 AND ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭʼ�ı��ս����¼", lng����ID)
    
    '���汣�ս����¼
    Call DebugTool("���汣�ս����¼(zl9Insure\����������)")
    Call WriteBusinessLOG("���汣�ս����¼(zl9Insure\����������)", "", "")
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & ",0,0," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & _
        "," & -1 * Nvl(rsTemp!���Ը����, 0) & "," & 0 & ",0,'" & gComInfo_����.������ˮ�� & "',null,null,'" & Nvl(rsTemp!��ע) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
    
    blnTrans = True
    gcnBJYB.BeginTrans
    '����һ���˷Ѽ�¼���ɣ���������
    '��彻����ˮ��,����,ԭ������ˮ��,ԭ��������,�˷�����,����Ա����,�ϴ�
    Call DebugTool("��������˷Ѽ�¼(zl9Insure\����������)")
    Call WriteBusinessLOG("��������˷Ѽ�¼(zl9Insure\����������)", "", "")
    gstrSQL = "ZL_�˷���Ϣ_INSERT(" & _
            "'" & gComInfo_����.������ˮ�� & "','" & str���� & "','" & strԭ������ˮ�� & "'," & _
            "to_Date('" & str����ʱ�� & "','yyyy-MM-dd hh24:mi:ss')," & _
            "to_Date('" & str�˷�ʱ�� & "','yyyy-MM-dd hh24:mi:ss')," & _
            "'" & UserInfo.���� & "',0)"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
'    **סԺ�������ʱ��Ҫɾ���ϴε���ʷ�����¼**
    gstrSQL = "ZL_�ֲ����Ѽ�¼_DELETE('" & strԭ������ˮ�� & "')"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc

    gcnBJYB.CommitTrans
    blnTrans = False
    ����������_���� = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\����������)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\����������)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Function

Public Function ��Ժ�Ǽ�_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    On Error GoTo errHand
    Dim str��Ժ���� As String
    Dim rsTemp As New ADODB.Recordset
    
    Call DebugTool("��ȡ��Ժ�Ǽ���ˮ��(zl9Insure\��Ժ�Ǽ�)")
    Call WriteBusinessLOG("��ȡ��Ժ�Ǽ���ˮ��(zl9Insure\��Ժ�Ǽ�)", "", "")
    Call GetSequence_����(lng����ID)
    '��סԺ��ˮ��ֻ��18λ������ˮ���ǰ�20λ���������ȥ�����������ַ�
    gComInfo_����.������ˮ�� = Mid(gComInfo_����.������ˮ��, 1, 8) & Mid(gComInfo_����.������ˮ��, 11)
    
    gstrSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ����", lng����ID, lng��ҳID)
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss")
    
    '��ȡ�˴���Ժ�ķ�ʽ������
    Call DebugTool("��ȡ�ò��˵���Ժ���͡���Ժ��ʽ����Ժ����(zl9Insure\��Ժ�Ǽ�)")
    Call WriteBusinessLOG("��ȡ�ò��˵���Ժ���͡���Ժ��ʽ����Ժ����(zl9Insure\��Ժ�Ǽ�)", "", "")
    gstrSQL = "Select ҵ������,����,��Ժ���,��Ժ��ʽ,��Ժ���� From �����ʻ� Where ����='" & gComInfo_����.���� & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\�������_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    
    If Not �������ⲡ(lng����ID) Then
        '������Ժ��¼����������
        '����ID,��ҳID,����,��Ժ�ǼǺ�,��Ժ����,��Ժ��ʽ,��Ժ����,�ϴ�
        Call DebugTool("������Ժ��Ϣ(zl9Insure\��Ժ�Ǽ�)")
        Call WriteBusinessLOG("������Ժ��Ϣ(zl9Insure\��Ժ�Ǽ�)", "", "")
        gstrSQL = "ZL_��Ժ��Ϣ_INSERT(" & _
            "" & lng����ID & "," & lng��ҳID & ",'" & rsTemp!���� & "','" & gComInfo_����.������ˮ�� & "'," & _
            "" & rsTemp!��Ժ��� & "," & rsTemp!��Ժ��ʽ & ",to_Date('" & str��Ժ���� & "','yyyy-MM-dd hh24:mi:ss')" & ",0)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    '�ı䲡��״̬
    Call DebugTool("�޸Ĳ��˵ĵ�ǰ״̬(zl9Insure\��Ժ�Ǽ�)")
    Call WriteBusinessLOG("�޸Ĳ��˵ĵ�ǰ״̬(zl9Insure\��Ժ�Ǽ�)", "", "")
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\��Ժ�Ǽ�)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\��Ժ�Ǽ�)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    On Error GoTo errHand
    Dim str��Ժ�ǼǺ� As String
    Dim rsTemp As New ADODB.Recordset
    
    '����ü�¼δ�ϴ�,����������Ժ��סԺ����ֻ�г�Ժ�󣬲Ż��ϴ���
    Call DebugTool("��飺���ϴ�����������Ժ(zl9Insure\��Ժ�Ǽǳ���)")
    Call WriteBusinessLOG("��飺���ϴ�����������Ժ(zl9Insure\��Ժ�Ǽǳ���)", "", "")
    gstrSQL = "Select Nvl(�ϴ�,0) �ϴ�,��Ժ�ǼǺ� From ��Ժ��Ϣ Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\�������_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp!�ϴ� = 1 Then
        MsgBox "�ò��˵����м�¼�����ϴ���ҽ�����ģ���������г�����Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    str��Ժ�ǼǺ� = Nvl(rsTemp!��Ժ�ǼǺ�)
    
    '�ı䲡��״̬
    Call DebugTool("�޸�״̬Ϊ����Ժ(zl9Insure\��Ժ�Ǽǳ���)")
    Call WriteBusinessLOG("�޸�״̬Ϊ����Ժ(zl9Insure\��Ժ�Ǽǳ���)", "", "")
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    If Not �������ⲡ(lng����ID) Then
        '�����Ժ��¼��ͬʱ�����Ժ�����Ϣ��
        Call DebugTool("ɾ��������Ժ��Ϣ(zl9Insure\��Ժ�Ǽǳ���)")
        Call WriteBusinessLOG("ɾ��������Ժ��Ϣ(zl9Insure\��Ժ�Ǽǳ���)", "", "")
        gstrSQL = "ZL_��Ժ��Ϣ_DELETE('" & str��Ժ�ǼǺ� & "')"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\��Ժ�Ǽǳ���)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\��Ժ�Ǽǳ���)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    On Error GoTo errHand
    Dim str��Ժ��� As String, str��Ժ���� As String, str��Ժ���� As String, int��Ժ��� As Integer
    Dim str��Ժ�ǼǺ� As String, str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    '���û�з����κη��ã�����������Ժ
    Call DebugTool("δ�������ã�ֻ�ܳ�����Ժ(zl9Insure\��Ժ�Ǽ�)")
    Call WriteBusinessLOG("δ�������ã�ֻ�ܳ�����Ժ(zl9Insure\��Ժ�Ǽ�)", "", "")
    gstrSQL = "Select Count(*) AS Records From סԺ���ü�¼" & _
        " Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "'���û�з����κη��ã�����������Ժ", lng����ID, lng��ҳID)
    If rsTemp!Records = 0 Then
        MsgBox "�������޷��ó�Ժ���볷����Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not �������ⲡ(lng����ID) Then
        '��д��Ժ�����Ϣ��������Ϣ��Ҫ��С���߲��䣬�����ϴ���
        Call DebugTool("��ȡ��Ժ��ϡ���Ժ���ҡ���Ժ���(zl9Insure\��Ժ�Ǽ�)")
        Call WriteBusinessLOG("��ȡ��Ժ��ϡ���Ժ���ҡ���Ժ���(zl9Insure\��Ժ�Ǽ�)", "", "")
        str��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True, True)
        '��ȡ�ٴ��������ʼ���Ժ���
        gstrSQL = " Select B.��������,A.��Ժ��ʽ,A.��Ժ���� " & _
                  " From ������ҳ A,�ٴ����� B " & _
                  " Where A.��Ժ����ID=B.����ID(+) And A.����ID=[1] And A.��ҳID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ٴ��������ʼ���Ժ���", lng����ID, lng��ҳID)
        str��Ժ���� = Nvl(rsTemp!��������)
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss")
        '����-1,��ת-2,δ��-3,����-4,תԺ-5,ת��-6,����-9
        Select Case rsTemp!��Ժ��ʽ
        Case "����"
            int��Ժ��� = 1
        Case "��ת"
            int��Ժ��� = 2
        Case "δ��"
            int��Ժ��� = 3
        Case "����"
            int��Ժ��� = 4
        Case "תԺ"
            int��Ժ��� = 5
        Case "ת��"
            int��Ժ��� = 6
        Case Else
            int��Ժ��� = 9
        End Select
        
        'ȡ����Ժ�ǼǺż�����
        Call DebugTool("����Ժ��Ϣ��ȡ����Ժ�ǼǺż�����(zl9Insure\��Ժ�Ǽ�)")
        Call WriteBusinessLOG("����Ժ��Ϣ��ȡ����Ժ�ǼǺż�����(zl9Insure\��Ժ�Ǽ�)", "", "")
        gstrSQL = "Select ��Ժ�ǼǺ�,���� From ��Ժ��Ϣ Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "zl9Insure\��Ժ�Ǽ�", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
        str��Ժ�ǼǺ� = Nvl(rsTemp!��Ժ�ǼǺ�)
        str���� = Nvl(rsTemp!����)
        
        If str��Ժ���� = "" Then
            MsgBox "�������ó�Ժ������ҽ�����ҵĶ���[���Ź���]", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    If Not �������ⲡ(lng����ID) Then
        '�������£�
        '��Ժ�ǼǺ�,����,��Ժ�Ʊ����,��Ҫ���,��Ҫ��������,��������,�������Ʊ���,��Ժ�������,�ϴ�
        Call DebugTool("������Ժ�����Ϣ��¼(zl9Insure\��Ժ�Ǽ�)")
        Call WriteBusinessLOG("������Ժ�����Ϣ��¼(zl9Insure\��Ժ�Ǽ�)", "", "")
        gstrSQL = "ZL_��Ժ�����Ϣ_UPDATE('" & str��Ժ�ǼǺ� & "'," & _
                "NULL,'" & str���� & "'," & lng����ID & "," & lng��ҳID & ",'" & str��Ժ���� & "'," & _
                "'" & Replace(Split(str��Ժ���, mstrSplit)(0), "(" & Split(str��Ժ���, mstrSplit)(1) & ")", "") & "','" & Split(str��Ժ���, mstrSplit)(1) & "'," & _
                "NULL,NULL," & int��Ժ��� & ",to_Date('" & str��Ժ���� & "','yyyy-MM-dd hh24:mi:ss')" & ",0)"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\��Ժ�Ǽ�)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\��Ժ�Ǽ�)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    
    If Not �������ⲡ(lng����ID) Then
        '��Ժ�����Ϣ���ϴ��Ĳ�������
        gstrSQL = " Select Nvl(�ϴ�,0) �ϴ� From ��Ժ�����Ϣ " & _
                  " Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
        If rsTemp.State = 1 Then rsTemp.Close
        Call SQLTest(App.Title, "zl9Insure\��Ժ�Ǽǳ���", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
        If rsTemp!�ϴ� = 1 Then
            MsgBox "�ò��˱���סԺ���������ϴ������ģ�����������Ժ�Ǽǳ�����", vbInformation, gstrSysName
            Exit Function
        End If
        
        'ɾ����Ժ�����Ϣ
        gstrSQL = "zl_��Ժ�����Ϣ_DELETE(" & lng����ID & "," & lng��ҳID & ")"
        gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\��Ժ�Ǽǳ���)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\��Ժ�Ǽǳ���)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_����(ByVal lng����ID As Long) As Currency
    �������_���� = 0
End Function

Public Function סԺ����_����(ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
    '������סԺ���������з��ý��з��÷ֽ�
    Dim int��Ŀ��� As Integer  '0-ҩƷ 1-������Ŀ 2-������ʩ
    Dim str�������� As String
    Dim strҽ������ As String, strHIS��Ŀ���� As String, str����Ա��ʶ As String, strReturn As String
    Dim str���� As String, str���÷ֽ���� As String, str���㷽ʽ As String
    Dim dbl�����ܶ� As Double, dblҽ������ As Double, dbl�󲡲��� As Double, dbl�ֽ� As Double
    Dim dblҽ���� As Double, dbl����Ӧ�� As Double, dbl�����Ը� As Double, dblͳ��ⶥ��ҽ���� As Double
    Dim strFields As String, strValues As String
    Dim blnTrans As Boolean
    Dim lng��ҳID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim arr���㷽ʽ
    Const int�����ܶ� As Integer = 3
    Const intҽ���� As Integer = 4
    Const intҽ������ As Integer = 5
    Const int�󲡲��� As Integer = 7
    Const int����Ӧ�� As Integer = 9
    Const int�����Ը� As Integer = 10
    Const intͳ��ⶥ��ҽ���� As Integer = 11
    On Error GoTo errHandle
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "��֧����;���㣬��Ϊ�ò��˰����Ժ���ٽ��г�Ժ���㣡", vbInformation, gstrSysName
        Exit Function
    End If
    
    str�������� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    
    '��ȡ���η�����ϸ
    gstrSQL = " Select ��ҳID From סԺ���ü�¼ Where ����ID=[1] ANd Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���η�����ϸ", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    '��ȡ�α��˹���Ա��ʶ
    Call DebugTool("��ȡ�α��˹���Ա��ʶ(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��ȡ�α��˹���Ա��ʶ(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "Select ҵ������,����,����Ա From �����ʻ� Where ����ID=" & lng����ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\סԺ�������_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str����Ա��ʶ = rsTemp!����Ա
    gComInfo_����.���� = rsTemp!����
    gComInfo_����.ҵ������ = rsTemp!ҵ������
    
    '����������
    If gComInfo_����.ҵ������ = "12" Then
        סԺ����_���� = �������_����(lng����ID, 0, "", True)
        Exit Function
    End If
    
    '�õ��ӿڷ��صĴ�����
    strReturn = gComInfo_����.���� & mstrSplit & str����Ա��ʶ
    Call DebugTool("���û�ȡ������Ϣ�ӿ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���û�ȡ������Ϣ�ӿ�(zl9Insure\סԺ�������)", "", "")
    If Not ���ýӿ�_����(�ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ, strReturn) Then Exit Function   'strReturn�����ݽ����ڷ��÷ֽ⣬���²��ܽ��и�ֵ
    str���� = strReturn
    
    '�����صĴ�����Ϣת��Ϊ���÷ֽ�����ĸ�ʽ
    str���÷ֽ���� = TransationHosp(str����)
    
    '���÷��÷ֽ⺯����strReturn���Ǵ�����Ϣ��
    Call DebugTool("���÷�����ϸ�ֽ⺯��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���÷�����ϸ�ֽ⺯��(zl9Insure\סԺ�������)", "", "")
    strReturn = str���÷ֽ����
    If Not ���ýӿ�_����(���÷ֽ�_סԺ1, strReturn) Then Exit Function
    strReturn = AnalyFile_סԺ1(True, lng����ID)
    If strReturn <> "" Then
        arr���㷽ʽ = Split(strReturn, mstrSplit)
        dbl�����ܶ� = Val(arr���㷽ʽ(int�����ܶ�))
        dblҽ������ = Val(arr���㷽ʽ(intҽ������))
        dbl�󲡲��� = Val(arr���㷽ʽ(int�󲡲���))
        
        dblҽ���� = Val(arr���㷽ʽ(intҽ����))
        dbl����Ӧ�� = Val(arr���㷽ʽ(int����Ӧ��))
        dbl�����Ը� = Val(arr���㷽ʽ(int�����Ը�))
        dblͳ��ⶥ��ҽ���� = Val(arr���㷽ʽ(intͳ��ⶥ��ҽ����))
    End If
    dbl�ֽ� = dbl�����ܶ� - dblҽ������ - dbl�󲡲���
    
    
    '*****************�����м�⣬�Ա��ϴ�*****************
    blnTrans = True
    gcnBJYB.BeginTrans
    '**���汾�εĽ��״�����Ϣ**
    Call DebugTool("���汾�εĽ��״�����Ϣ(zl9Insure\סԺ����)")
    Call WriteBusinessLOG("���汾�εĽ��״�����Ϣ(zl9Insure\סԺ����)", "", "")
    If Not SaveBusinessDeal(str����) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**���潻�װ汾��Ϣ**
    Call DebugTool("���汾�εĽ��װ汾��Ϣ(zl9Insure\סԺ����)")
    Call WriteBusinessLOG("���汾�εĽ��װ汾��Ϣ(zl9Insure\סԺ����)", "", "")
    If Not SaveBusinessVersion(False) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**����סԺ������Ϣ��סԺ�ֶμ�סԺ������ϸ**
    Call DebugTool("����סԺ������Ϣ��סԺ�ֶμ�סԺ������ϸ(zl9Insure\סԺ����)")
    Call WriteBusinessLOG("����סԺ������Ϣ��סԺ�ֶμ�סԺ������ϸ(zl9Insure\סԺ����)", "", "")
    strReturn = AnalyFile_סԺ1(False, lng����ID)
    If strReturn = "" Then
        '����ʧ��
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '**���������ֲ����Ѽ�¼**
    Call DebugTool("���������ֲ����Ѽ�¼(zl9Insure\סԺ����)")
    Call WriteBusinessLOG("���������ֲ����Ѽ�¼(zl9Insure\סԺ����)", "", "")
    If Not SaveDeal(False, True) Then
        gcnBJYB.RollbackTrans
        Exit Function
    End If
    
    '���³�Ժ�����Ϣ�еĽ�����ˮ��
    gstrSQL = "Update ��Ժ�����Ϣ Set ������ˮ��='" & gComInfo_����.������ˮ�� & "' Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID
    gcnBJYB.Execute gstrSQL
    
    '**���汣�ս����¼**
    '����ͳ����=����Ա��������;ͳ�ﱨ�����=ͳ�����;���Ը����=�󲡻���
    '֧��˳���=������ˮ��,��ע=ҵ������
    Call DebugTool("���汣�ս����¼(zl9Insure\סԺ����)")
    Call WriteBusinessLOG("���汣�ս����¼(zl9Insure\סԺ����)", "", "")
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng��ҳID & "," & 0 & "," & 0 & "," & 0 & "," & _
        dbl�����ܶ� & "," & dbl�ֽ� & ",0,0," & dblҽ������ & "," & dbl�󲡲��� & "," & _
        0 & ",0,'" & gComInfo_����.������ˮ�� & "'," & lng��ҳID & ",null,'" & gComInfo_����.ҵ������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
    
    gcnBJYB.CommitTrans
    blnTrans = False
    סԺ����_���� = True
    Exit Function
errHandle:
    Call DebugTool("(zl9INSURE\סԺ�������_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\סԺ�������_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Function

Public Function סԺ�������_����(ByVal rs��ϸ As ADODB.Recordset, ByVal lng����ID As Long) As String
    '������סԺ���������з��ý��з��÷ֽ�
    Dim int��Ŀ��� As Integer  '0-ҩƷ 1-������Ŀ 2-������ʩ
    Dim str�������� As String
    Dim strҽ������ As String, strHIS��Ŀ���� As String, str����Ա��ʶ As String, strReturn As String, strUser As String
    Dim str���� As String, str���÷ֽ���� As String, str���㷽ʽ As String
    Dim strFilter As String
    Dim strFields As String, strValues As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rsLimit As New ADODB.Recordset
    Dim arr���㷽ʽ
    Const intҽ������ As Integer = 5
    Const int�󲡲��� As Integer = 7
    'rs��ϸ��¼���е��ֶ��嵥
    'ID,��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    
    On Error GoTo errHandle
    
    strUser = GetUser
    str�������� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    
    '�Ȼ�ȡ���ν�����ˮ��
    Call DebugTool("������ˮ��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("������ˮ��(zl9Insure\סԺ�������)", "", "")
    Call GetSequence_����(lng����ID)
    
    '��ȡ�α��˹���Ա��ʶ
    Call DebugTool("��ȡ�α��˹���Ա��ʶ(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��ȡ�α��˹���Ա��ʶ(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "Select ����,����Ա From �����ʻ� Where ����ID=" & lng����ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\סԺ�������_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str����Ա��ʶ = rsTemp!����Ա
    gComInfo_����.���� = rsTemp!����
    
    '��ȡ�α��˱��ξ����ҵ�����ͣ���������ⲡ������е�������
    If �������ⲡ(lng����ID) Then
        '���������ⲡ����
        If Not �������ⲡ_����(rs��ϸ, str���㷽ʽ) Then Exit Function
        סԺ�������_���� = str���㷽ʽ
        Exit Function
    End If
    
    '��ȡ�˿̵Ĵ�����Ϣ�����÷ֽ⺯������Σ�
    Call DebugTool("��ȡ�˿̵���ʷ���Ѽ�¼�����÷ֽ⺯������Σ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��ȡ�˿̵���ʷ���Ѽ�¼�����÷ֽ⺯������Σ�(zl9Insure\סԺ�������)", "", "")
    Set rsDeal = GetDeal(lng����ID, str��������)
    '����¼�����ʷ���Ѽ�¼����������ļ�
    Call DebugTool("����¼�����ʷ���Ѽ�¼����������ļ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("����¼�����ʷ���Ѽ�¼����������ļ�(zl9Insure\סԺ�������)", "", "")
    If Not MakeFile_Center(rsDeal, �ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ) Then Exit Function
    '�õ��ӿڷ��صĴ�����
    strReturn = gComInfo_����.���� & mstrSplit & str����Ա��ʶ
    Call DebugTool("���û�ȡ������Ϣ�ӿ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���û�ȡ������Ϣ�ӿ�(zl9Insure\סԺ�������)", "", "")
    If Not ���ýӿ�_����(�ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ, strReturn) Then Exit Function   'strReturn�����ݽ����ڷ��÷ֽ⣬���²��ܽ��и�ֵ
    str���� = strReturn
    
    '�ж��Ƿ��������ʹ�õ���Ŀ��������ڣ�����Ҫ����Աȷ���Ƿ�ʹ����ҽ����
    Call DebugTool("��������ʹ����Ŀ��¼����������Աѡ���Ƿ�����ҽ����(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��������ʹ����Ŀ��¼����������Աѡ���Ƿ�����ҽ����(zl9Insure\סԺ�������)", "", "")
    strFields = "ҽ����," & adLongVarChar & ",1" & mstrSplit & _
              "NO," & adLongVarChar & ",8" & mstrSplit & _
              "��¼����," & adLongVarChar & ",3" & mstrSplit & _
              "��¼״̬," & adLongVarChar & ",3" & mstrSplit & _
              "���," & adLongVarChar & ",5" & mstrSplit & _
              "HIS��Ŀ����," & adLongVarChar & ",100" & mstrSplit & _
              "ҽ������," & adLongVarChar & ",100" & mstrSplit & _
              "����," & adLongVarChar & ",20" & mstrSplit & _
              "����," & adLongVarChar & ",20" & mstrSplit & _
              "���," & adLongVarChar & ",20" & mstrSplit & _
              "����," & adLongVarChar & ",500" & mstrSplit & _
              "��ע," & adLongVarChar & ",500" & mstrSplit & _
              "����ҽ��," & adLongVarChar & ",20" & mstrSplit & _
              "�Ǽ�ʱ��," & adLongVarChar & ",20"
    Call Record_Init(rsLimit, strFields)
    strFields = "ҽ����|NO|��¼����|��¼״̬|���|HIS��Ŀ����|ҽ������|����|����|���|����|��ע|����ҽ��|�Ǽ�ʱ��"
    
    If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
    Do Until rs��ϸ.EOF
        'ֻ��ҩƷ��Ŀ�Ŵ������Ƶȼ�
        If rs��ϸ!��¼״̬ = 1 Or rs��ϸ!��¼״̬ = 3 Then
            gstrSQL = "Select A.��Ŀ���� As ҽ������,C.ͨ������ AS ��Ŀ����,D.���� AS ���Ƶȼ�,F.��ע " & _
                " From ����֧����Ŀ A,ҩƷĿ¼ B,ҩƷ��Ϣ C," & strUser & ".ҩƷĿ¼ F," & _
                " (Select B.����,B.���� From " & strUser & ".ָ������ A," & strUser & ".ָ����ϵ���ձ� B Where A.���=B.��� And A.����='ʹ�����Ƶȼ�') D" & _
                " Where A.����=[1] And A.�շ�ϸĿID=B.ҩƷID And B.ҩ��ID=C.ҩ��ID And F.����=A.��Ŀ���� " & _
                " AND B.ҩƷID=[2] ANd F.ʹ�����Ƶȼ�=D.����"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺԤ��", TYPE_����, CLng(rs��ϸ!�շ�ϸĿID))
            If Not rsTemp.EOF Then
                '������¼�����Ա����Ա���ҽ�������ȷ��
                strValues = "��" & mstrSplit & rs��ϸ!NO & mstrSplit & rs��ϸ!��¼���� & mstrSplit & rs��ϸ!��¼״̬ & mstrSplit & rs��ϸ!��� & mstrSplit & _
                    rsTemp!��Ŀ���� & mstrSplit & rsTemp!ҽ������ & mstrSplit & _
                    Format(rs��ϸ!����, "#####0.00;-#####0.00;0;") & mstrSplit & Format(rs��ϸ!�۸�, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
                    Format(rs��ϸ!���, "#####0.0000;-#####0.0000;0;") & mstrSplit & Nvl(rsTemp!���Ƶȼ�) & mstrSplit & Nvl(rsTemp!��ע) & mstrSplit & _
                    Nvl(rs��ϸ!ҽ��) & mstrSplit & Format(rs��ϸ!����ʱ��, "yyyyMMdd")
                Call Record_Add(rsLimit, strFields, strValues)
            End If
        End If
        
        rs��ϸ.MoveNext
    Loop
    '��Ҫ����Աȷ��������Ŀ��ҽ������
    If rsLimit.RecordCount <> 0 Then Call frm������ҩҽ�����⻮��.ShowEditor(rsLimit)
    
    '��������ϸ������ϸ�ļ���Ҳ�Ƿ��÷ֽ⺯������Σ�
'    ----�����ļ�˵��
'    ���    ������  ����    ��󳤶�    ˵��
'    1   ��Ŀ���    C   9   ˳���
'    2   ������  C   20  �μ���׼AKC220���ɿ�
'    3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ���루��Ӧ֢��ҩ��Ҫ����Աȷ����Щ��ҽ������ҩ����ҽ������ҩ���㣩
'    4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'    5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'    6   ����    N   10,4    AKC225
'    7   ����    N   8,2 AKC226
'    8   �����ܽ��  N   10,4    ʵ�ʽ�����
'    9   ���÷�������    D   8   YYYYMMDD
    strFields = "��Ŀ���," & adLongVarChar & ",9" & mstrSplit & _
                "ҽ����," & adLongVarChar & ",20" & mstrSplit & _
                "��Ŀ����," & adLongVarChar & ",20" & mstrSplit & _
                "��Ŀ����," & adLongVarChar & ",100" & mstrSplit & _
                "��Ŀ���," & adLongVarChar & ",3" & mstrSplit & _
                "����," & adDouble & ",18" & mstrSplit & _
                "����," & adDouble & ",18" & mstrSplit & _
                "�����ܽ��," & adDouble & ",18" & mstrSplit & _
                "���÷�������," & adLongVarChar & ",20"
    Call Record_Init(rsCharge, strFields)
    
    '�õ�����¼��ķ�����ϸ�����ҽ����Ϣ��������w01��ͷ����ʾ������ʩ����Ŀ��
    Call DebugTool("����������ϸ��¼��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("����������ϸ��¼��(zl9Insure\סԺ�������)", "", "")
    strFields = "��Ŀ���|ҽ����|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|���÷�������"
    If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
    Do Until rs��ϸ.EOF
        '��ҩƷ��ͨ�����ƻ��ҩƷ��Ŀ������ȡ����
        gstrSQL = "Select A.��Ŀ���� As ҽ������,C.ͨ������ AS ��Ŀ���� " & _
            " From ����֧����Ŀ A,ҩƷĿ¼ B,ҩƷ��Ϣ C" & _
            " Where A.����(+)=[1] And A.�շ�ϸĿID(+)=B.ҩƷID And B.ҩ��ID=C.ҩ��ID " & _
            " AND B.ҩƷID=[2]" & _
            " UNION " & _
            " Select A.��Ŀ���� As ҽ������,B.���� AS ��Ŀ����" & _
            " From ����֧����Ŀ A,�շ�ϸĿ B" & _
            " Where A.����(+)=[1] AND B.ID=[2]" & _
            " And A.�շ�ϸĿID(+)=B.ID AND B.��� Not In ('5','6','7')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺԤ��", TYPE_����, CLng(rs��ϸ!�շ�ϸĿID))
        If rsTemp.EOF Then
            MsgBox "������Ŀû�к�ҽ����Ŀ���ö��չ�ϵ��[������Ŀ]", vbInformation, gstrSysName
            Exit Function
        End If
        
        '���ж��Ƿ���������ҩ������ǣ����ݲ���Ա��ѡ��ȷ��
        strFilter = "NO='" & rs��ϸ!NO & "' And ��¼����=" & rs��ϸ!��¼���� & IIf(rs��ϸ!��¼״̬ = "2", " And ��¼״̬=3", " And ��¼״̬=1") & " And ���=" & rs��ϸ!���
        rsLimit.Filter = strFilter
        If rsLimit.RecordCount = 0 Then
            strҽ������ = Nvl(rsTemp!ҽ������, 0)
        Else
            If rsLimit!ҽ���� = 1 Or rsLimit!ҽ���� = "��" Then
                strҽ������ = Nvl(rsTemp!ҽ������, 0)
            Else
                strҽ������ = 0
            End If
        End If
        strHIS��Ŀ���� = rsTemp!��Ŀ����
        
        If InStr(1, "5,6,7", rs��ϸ!�շ����) <> 0 Then
            int��Ŀ��� = 0 '0-ҩƷ
        Else
            int��Ŀ��� = (IIf(strҽ������ Like "w01*", 2, 1)) '1-������Ŀ 2-������ʩ
        End If
        
        '����������ϸ��¼��,�Թ���������ļ�
        strValues = rs��ϸ.AbsolutePosition & mstrSplit & (rs��ϸ!NO & "*" & rs��ϸ!��¼���� & "*" & rs��ϸ!��¼״̬ & "*" & rs��ϸ!���) & mstrSplit & _
            strҽ������ & mstrSplit & strHIS��Ŀ���� & mstrSplit & _
            int��Ŀ��� & mstrSplit & Format(rs��ϸ!�۸�, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs��ϸ!����, "#####0.00;-#####0.00;0;") & mstrSplit & Format(rs��ϸ!���, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs��ϸ!����ʱ��, "yyyyMMdd")
        Call Record_Add(rsCharge, strFields, strValues)
        
        rs��ϸ.MoveNext
    Loop
    
    '����������ϸ�ļ�
    Call DebugTool("���ݷ�����ϸ����������ļ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���ݷ�����ϸ����������ļ�(zl9Insure\סԺ�������)", "", "")
    If Not MakeFile_Center(rsCharge, ���÷ֽ�_סԺ1) Then Exit Function
    
    '�����صĴ�����Ϣת��Ϊ���÷ֽ�����ĸ�ʽ
    str���÷ֽ���� = TransationHosp(str����)
    
    '���÷��÷ֽ⺯����strReturn���Ǵ�����Ϣ��
    Call DebugTool("���÷�����ϸ�ֽ⺯��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���÷�����ϸ�ֽ⺯��(zl9Insure\סԺ�������)", "", "")
    strReturn = str���÷ֽ����
    If Not ���ýӿ�_����(���÷ֽ�_סԺ1, strReturn) Then Exit Function
    strReturn = AnalyFile_סԺ1(True)
    If strReturn = "" Then Exit Function
    
    If Not SaveDeal(False, False) Then Exit Function
    
    '��֯�ɽ��㷽ʽ��������
    Call DebugTool("�������÷ֽⷵ�صĻ������ݣ���������Ϣ���ظ���������(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("�������÷ֽⷵ�صĻ������ݣ���������Ϣ���ظ���������(zl9Insure\סԺ�������)", "", "")
    arr���㷽ʽ = Split(strReturn, mstrSplit)
    str���㷽ʽ = mstrSplit & "ͳ��֧��;" & arr���㷽ʽ(intҽ������) & ";0"
    str���㷽ʽ = str���㷽ʽ & mstrSplit & "���֧��;" & arr���㷽ʽ(int�󲡲���) & ";0"
    str���㷽ʽ = Mid(str���㷽ʽ, 2)
    
    סԺ�������_���� = str���㷽ʽ
    Exit Function
errHandle:
    Call DebugTool("(zl9INSURE\סԺ�������_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\סԺ�������_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(ByVal lng����ID As Long) As Boolean
    '**סԺ�������ʱ��Ҫɾ���ϴε���ʷ�����¼**
    On Error GoTo errHand
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim blnTrans As Boolean
    Dim str���� As String, strԭ������ˮ�� As String
    Dim str����ʱ�� As String, str�˷�ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    
    '��ȡ����ID
    Call DebugTool("��ȡ����ID(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��ȡ����ID(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "Select ����ID From סԺ���ü�¼ Where ����ID=[1] ANd Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID)
    lng����ID = rsTemp!����ID
    
    Call DebugTool("����������ˮ��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("����������ˮ��(zl9Insure\סԺ�������)", "", "")
    Call GetSequence_����(lng����ID)
    
    'ȡ��ԭʼ�����¼�Ľ���ʱ��
    Call DebugTool("ȡ��ԭʼ�����¼�Ľ���ʱ��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("ȡ��ԭʼ�����¼�Ľ���ʱ��(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "Select ����,������ˮ��,����ʱ�� From " & GetUser & ".���״�����Ϣ " & _
            " Where ������ˮ��=(Select ֧��˳��� From ���ս����¼ Where ����=2 AND ��¼ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ԭʼ�����¼�Ľ���ʱ��", lng����ID)
    str���� = rsTemp!����
    strԭ������ˮ�� = rsTemp!������ˮ��
    str����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    str�˷�ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '��飬����ڸý��׺󣬻��������������ף���������г����������˷ѵļ�¼�ſ���
    Call DebugTool("����Ƿ�����һ�ʿ�ʼ�˷�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("����Ƿ�����һ�ʿ�ʼ�˷�(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "Select Count(*) AS Records From ���״�����Ϣ A,�����ʻ� B" & _
            " Where A.����=B.���� And B.����ID=" & lng����ID & _
            " And ����ʱ��>to_Date('" & str����ʱ�� & "','yyyy-MM-dd hh24:mi:ss')" & _
            " And ������ˮ�� Not In (Select ԭ������ˮ�� From �˷���Ϣ Where ����='" & str���� & "')"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\�������_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    If rsTemp!Records > 0 Then
        MsgBox "ҽ���ӿڲ�������м俪ʼ�˷ѣ���ֻ�ܴ����һ��ҵ��ʼ�˷ѣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ������¼�Ľ���ID�����ݺ�
    Call DebugTool("ȡ����ID(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("ȡ����ID(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���²����Ľ���ID", lng����ID)
    lng����ID = rsTemp!ID
    
    '��ȡԭʼ�ı��ս����¼
    Call DebugTool("��ȡԭʼ�ı��ս����¼(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��ȡԭʼ�ı��ս����¼(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "Select * From ���ս����¼ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡԭʼ�ı��ս����¼", lng����ID)
    
    '���汣�ս����¼
    Call DebugTool("���汣�ս����¼(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���汣�ս����¼(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "zl_���ս����¼_insert(" & rsTemp!���� & "," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & Nvl(rsTemp!��ҳID, 0) & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & ",0,0," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & _
        "," & -1 * Nvl(rsTemp!���Ը����, 0) & "," & 0 & ",0,'" & gComInfo_����.������ˮ�� & "'," & Nvl(rsTemp!��ҳID, 0) & ",null,'" & Nvl(rsTemp!��ע) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�ս����¼")
    
    blnTrans = True
    gcnBJYB.BeginTrans
    '����һ���˷Ѽ�¼���ɣ���������
    '��彻����ˮ��,����,ԭ������ˮ��,ԭ��������,�˷�����,����Ա����,�ϴ�
    Call DebugTool("��������˷Ѽ�¼(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��������˷Ѽ�¼(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "ZL_�˷���Ϣ_INSERT(" & _
            "'" & gComInfo_����.������ˮ�� & "','" & str���� & "','" & strԭ������ˮ�� & "'," & _
            "to_Date('" & str����ʱ�� & "','yyyy-MM-dd hh24:mi:ss')," & _
            "to_Date('" & str�˷�ʱ�� & "','yyyy-MM-dd hh24:mi:ss')," & _
            "'" & UserInfo.���� & "',0)"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    '**סԺ�������ʱ��Ҫɾ���ϴε���ʷ�����¼**
    gstrSQL = "ZL_�ֲ����Ѽ�¼_DELETE('" & strԭ������ˮ�� & "')"
    gcnBJYB.Execute gstrSQL, , adCmdStoredProc
    
    gcnBJYB.CommitTrans
    סԺ�������_���� = True
    Exit Function
errHand:
    Call DebugTool("(zl9Insure\סԺ�������)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9Insure\סԺ�������)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnBJYB.RollbackTrans
End Function

Private Function �������ⲡ(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ҵ������ From �����ʻ� Where ����ID=" & lng����ID
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "zl9Insure\�������ⲡ", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    �������ⲡ = (rsTemp!ҵ������ = "12")
End Function

Private Function �������ⲡ_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    Dim int��Ŀ��� As Integer  '0-ҩƷ 1-������Ŀ 2-������ʩ
    Dim lng����ID As Long
    Dim str�������� As String, str���÷������� As String
    Dim strҽ������ As String, strHIS��Ŀ���� As String, str����Ա��ʶ As String, strReturn As String
    Dim str���� As String, str���÷ֽ���� As String
    Dim strFields As String, strValues As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDeal As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim arr���㷽ʽ
    Const intҽ������ As Integer = 5
    Const int�󲡲��� As Integer = 7
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    'rs��ϸ��¼���е��ֶ��嵥
    'ID,��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    On Error GoTo errHandle
    lng����ID = rs��ϸ!����ID
'    gstrSQL = "Select ��Ժ���� From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� And B.����ID=" & lng����ID
'    Call OpenRecordset(rsTemp, "��ȡ������Ժ����")
'    str�������� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    str�������� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    
    '�Ȼ�ȡ���ν�����ˮ��
    Call DebugTool("������ˮ��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("������ˮ��(zl9Insure\סԺ�������)", "", "")
    Call GetSequence_����(lng����ID)
    
    '��ȡ�α��˹���Ա��ʶ
    Call DebugTool("��ȡ�α��˹���Ա��ʶ(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��ȡ�α��˹���Ա��ʶ(zl9Insure\סԺ�������)", "", "")
    gstrSQL = "Select ����Ա From �����ʻ� Where ����='" & gComInfo_����.���� & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.Title, "ZL9INSURE\�������ⲡ_����", gstrSQL): rsTemp.Open gstrSQL, gcnBJYB: Call SQLTest
    str����Ա��ʶ = rsTemp!����Ա
    
    '��ȡ�˿̵Ĵ�����Ϣ�����÷ֽ⺯������Σ�
    Call DebugTool("��ȡ�˿̵���ʷ���Ѽ�¼�����÷ֽ⺯������Σ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("��ȡ�˿̵���ʷ���Ѽ�¼�����÷ֽ⺯������Σ�(zl9Insure\סԺ�������)", "", "")
    Set rsDeal = GetDeal(lng����ID, str��������)
    '����¼�����ʷ���Ѽ�¼����������ļ�
    Call DebugTool("����¼�����ʷ���Ѽ�¼����������ļ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("����¼�����ʷ���Ѽ�¼����������ļ�(zl9Insure\סԺ�������)", "", "")
    If Not MakeFile_Center(rsDeal, �ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ) Then Exit Function
    '�õ��ӿڷ��صĴ�����
    strReturn = gComInfo_����.���� & mstrSplit & str����Ա��ʶ
    Call DebugTool("���û�ȡ������Ϣ�ӿ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���û�ȡ������Ϣ�ӿ�(zl9Insure\סԺ�������)", "", "")
    If Not ���ýӿ�_����(�ӿڹ���.��ȡ�ֲᲡ�˴�����Ϣ, strReturn) Then Exit Function   'strReturn�����ݽ����ڷ��÷ֽ⣬���²��ܽ��и�ֵ
    str���� = strReturn
    
    '��������ϸ������ϸ�ļ���Ҳ�Ƿ��÷ֽ⺯������Σ�
'    ----�����ļ�˵��
'    ���    ������  ����    ��󳤶�    ˵��
'    1   ��Ŀ���    C   9   ˳���
'    2   ������  C   20  �μ���׼AKC220���ɿ�
'    3   ��Ŀ����    C   20  ҩƷ��������Ŀ�������ʩ����
'    4   ��Ŀ����    C   100 ��ҽԺ��Ŀ����
'    5   ��Ŀ���    C   3   0-ҩƷ 1-������Ŀ 2-������ʩ
'    6   ����    N   10,4    AKC225
'    7   ����    N   8,2 AKC226
'    8   �����ܽ��  N   10,4    ʵ�ʽ�����
'    9   ���÷�������    D   8   YYYYMMDD
    strFields = "��Ŀ���," & adLongVarChar & ",9" & mstrSplit & _
                "������," & adLongVarChar & ",20" & mstrSplit & _
                "��Ŀ����," & adLongVarChar & ",20" & mstrSplit & _
                "��Ŀ����," & adLongVarChar & ",100" & mstrSplit & _
                "��Ŀ���," & adLongVarChar & ",3" & mstrSplit & _
                "����," & adDouble & ",18" & mstrSplit & _
                "����," & adDouble & ",18" & mstrSplit & _
                "�����ܽ��," & adDouble & ",18" & mstrSplit & _
                "���÷�������," & adLongVarChar & ",20"
    Call Record_Init(rsCharge, strFields)
    
    '�õ�����¼��ķ�����ϸ�����ҽ����Ϣ��������w01��ͷ����ʾ������ʩ����Ŀ��
    Call DebugTool("����������ϸ��¼��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("����������ϸ��¼��(zl9Insure\סԺ�������)", "", "")
    strFields = "��Ŀ���|������|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|���÷�������"
    If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
    Do Until rs��ϸ.EOF
        '��ҩƷ��ͨ�����ƻ��ҩƷ��Ŀ������ȡ����
        gstrSQL = "Select A.��Ŀ���� As ҽ������,C.ͨ������ AS ��Ŀ���� " & _
            " From ����֧����Ŀ A,ҩƷĿ¼ B,ҩƷ��Ϣ C" & _
            " Where A.����(+)=[1] And A.�շ�ϸĿID(+)=B.ҩƷID And B.ҩ��ID=C.ҩ��ID " & _
            " AND B.ҩƷID=[2]" & _
            " UNION " & _
            " Select A.��Ŀ���� As ҽ������,B.���� AS ��Ŀ����" & _
            " From ����֧����Ŀ A,�շ�ϸĿ B" & _
            " Where A.����(+)=[1] AND B.ID=[2]" & _
            " And A.�շ�ϸĿID(+)=B.ID AND B.��� Not In ('5','6','7')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", TYPE_����, CLng(rs��ϸ!�շ�ϸĿID))
        If rsTemp.EOF Then
            MsgBox "������Ŀû�к�ҽ����Ŀ���ö��չ�ϵ��[������Ŀ]", vbInformation, gstrSysName
            Exit Function
        End If
        strҽ������ = Nvl(rsTemp!ҽ������, 0)
        strHIS��Ŀ���� = rsTemp!��Ŀ����
        
        If InStr(1, "5,6,7", rs��ϸ!�շ����) <> 0 Then
            int��Ŀ��� = 0 '0-ҩƷ
        Else
            int��Ŀ��� = (IIf(strҽ������ Like "w01*", 2, 1)) '1-������Ŀ 2-������ʩ
        End If
        
        '����������ϸ��¼��,�Թ���������ļ�
        strValues = rs��ϸ.AbsolutePosition & mstrSplit & "" & mstrSplit & _
            strҽ������ & mstrSplit & strHIS��Ŀ���� & mstrSplit & _
            int��Ŀ��� & mstrSplit & Format(rs��ϸ!�۸�, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs��ϸ!����, "#####0.00;-#####0.00;0;") & mstrSplit & Format(rs��ϸ!���, "#####0.0000;-#####0.0000;0;") & mstrSplit & _
            Format(rs��ϸ!�Ǽ�ʱ��, "yyyyMMdd")
        Call Record_Add(rsCharge, strFields, strValues)
        
        rs��ϸ.MoveNext
    Loop
    
    '����������ϸ�ļ�
    Call DebugTool("���ݷ�����ϸ����������ļ�(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���ݷ�����ϸ����������ļ�(zl9Insure\סԺ�������)", "", "")
    If Not MakeFile_Center(rsCharge, ���÷ֽ�_��������1) Then Exit Function
    
    '�����صĴ�����Ϣת��Ϊ���÷ֽ�����ĸ�ʽ
    str���÷ֽ���� = TransationSpec(str����)
    
    '���÷��÷ֽ⺯����strReturn���Ǵ�����Ϣ��
    Call DebugTool("���÷�����ϸ�ֽ⺯��(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("���÷�����ϸ�ֽ⺯��(zl9Insure\סԺ�������)", "", "")
    strReturn = str���÷ֽ����
    If Not ���ýӿ�_����(���÷ֽ�_��������1, strReturn) Then Exit Function
    strReturn = AnalyFile_��������1(True)
    If strReturn = "" Then Exit Function
    
    If Not SaveDeal(True, False) Then Exit Function
    
    '��֯�ɽ��㷽ʽ��������
    Call DebugTool("�������÷ֽⷵ�صĻ������ݣ���������Ϣ���ظ���������(zl9Insure\סԺ�������)")
    Call WriteBusinessLOG("�������÷ֽⷵ�صĻ������ݣ���������Ϣ���ظ���������(zl9Insure\סԺ�������)", "", "")
    arr���㷽ʽ = Split(strReturn, mstrSplit)
    str���㷽ʽ = mstrSplit & "ͳ��֧��;" & arr���㷽ʽ(intҽ������) & ";0"
    str���㷽ʽ = str���㷽ʽ & mstrSplit & "���֧��;" & arr���㷽ʽ(int�󲡲���) & ";0"
    str���㷽ʽ = Mid(str���㷽ʽ, 2)
    
    �������ⲡ_���� = True
    Exit Function
errHandle:
    Call DebugTool("(zl9INSURE\�������ⲡ_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description)
    Call WriteBusinessLOG("(zl9INSURE\�������ⲡ_����)" & vbCrLf & _
        "�����:" & Err.Number & "|������Ϣ:" & Err.Description, "", "")
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
