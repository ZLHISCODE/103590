Attribute VB_Name = "mdl����"
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--�������ӿ�
    '����˵��:
    '   msgType-ҵ����������,�����µĲ�����
    '   packageType-���ݽ�����ʽ���ͣ�ϵͳ��������ʱʹ��,�����µĲ�����
    '   packageLength-���ݴ��ĳ���,�����µĲ�����
    '   str-���ݴ�,����ʱ��ͨ�����ݴ������������������ʱ�����ݴ��а������ص�����
    '   strCom:�������󴮿ڣ����ݶ��������λ�ã�����������ȡֵ��'com1','com2')
    '����:
    '   I.  ����������ֵ����0ʱ����ʾ�ɹ����ַ����а�����ҵ����󷵻ص�����
    '   II. ����������ֵ������0ʱ���μ��������һ����Ӧ����Ҫ�����������Ȼ������ʵ��Ĵ���

'��������������
Private Declare Function IC_Read_Base Lib "ICCNII32.DLL" (ByVal szData As String) As Long
Private Declare Function IC_Read_Plus Lib "ICCNII32.DLL" (nSequence As Long, ByVal szData As String) As Long

'Private Declare Function KfqTransData Lib "OltpTransKfq03.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
    ByVal str As String, ByVal strCom As String) As Long
    
'2005-08-02 �ܺ�ȫ
'������ҽ������
Private Declare Function KfqTransData Lib "OltpTransKfq05.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
    ByVal str As String, ByVal strCom As String) As Long
    
'--��ͨ�ӿ�
Private Declare Function OltpTransData Lib "OltpTransIc04.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
    ByVal str As String, ByVal strCom As String) As Long
    '����Ϊ�������Ĳ�����
''ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
''------------    ----------------    --------------         -----------------------------------------------
''1001            101                 95                     ʵʱ�鿨���������鿨��
''1002            12                  420                    ʵʱ����
''1003            7                   297                    ʵʱҽ����ϸ�����ύ
''1004            9                   136                    ʵʱסԺ�Ǽ������ύ
''1006            12                  420                    ʵʱ����Ԥ��
''1008            101                 95                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�

'2005-08-02 �ܺ�ȫ
'�����������޸�
'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
'------------    ----------------    --------------         -----------------------------------------------
'1001            101                 96                     ʵʱ�鿨���������鿨��
'1002            12                  509                    ʵʱ����
'1003            7                   240                    ʵʱҽ����ϸ�����ύ
'1004            9                   210                    ʵʱסԺ�Ǽ������ύ
'1006            12                  509                    ʵʱ����Ԥ��
'1008            101                 96                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�

'����Ϊ�����еĲ�����
'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
'------------    ----------------    --------------         -----------------------------------------------
'1001            101                 94                     ʵʱ�鿨���������鿨��
'1002            12                  424                    ʵʱ����
'1003            7                   230                    ʵʱҽ����ϸ�����ύ
'1004            9                   206                    ʵʱסԺ�Ǽ������ύ
'1006            12                  424                    ʵʱ����Ԥ��
'1008            101                 94                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�
'1005            8                   274                    ʵʱҽ������
'1007            2                   55                     �����ʻ���ѯ
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public gblnKFQCom_����  As Boolean   'true-�������ӿ�,False-��ͨ�ӿ�

Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

Public g�������_���� As �������
Private Type �������
    ���˱��            As String
    ����                As String
    �Ա�                As String
    ��������            As String
    ����                As Integer
    ���֤��            As String
    IC����              As Long
    �������            As Long
    ְ����ҽ���        As String
    ���������ʻ����    As Double
    ���������ʻ����    As Double
    ��ǰ״̬            As Double
    ͳ���ۼ�            As Double
    �½ɷѻ���          As Double
    �ʻ�״̬            As String
    �α����1           As String
    �α����2           As String
    �α����3           As String
    �α����4           As String
    �α����5           As String
    
    ת�ﵥ��            As String           '�����֤ʱ����
    ҽ������            As Long             '�����֤ʱѡ��,��������
    �������            As Long             '�����֤ʱѡ��,������ǽ��㷽ʽ����
    ֧�����            As Double           '
    ��ϱ���            As String           '��ϱ���ʱ����,������Ч
    �������            As String           '�������ʱ����,������Ч
    
    �����ʻ�ԭʼֵ      As Double          '������ѯ��ȡ
    �����ʻ���ǰֵ      As Double          '������ѯ��ȡ
    �����ʻ�״̬        As Double          '������ѯ��ȡ
    ����              As Double
    
    ��ǰ�����ʻ����    As Double
    ��ǰ�����ʻ����    As Double
    ��ǰͳ���ۼ�            As Double
    ���㿪ʼ            As Boolean             '��ҪӦ��������൥��,�账��ǰ�����
    ���ν���            As Boolean             '���ν���
    
End Type

Public gblnģ��ӿ�   As Boolean      'ģ��ӿ�����

Public gstrҽԺ����_���� As String        'ҽԺ����,ֻ��Ϊ4λ
Public gintComPort_���� As Integer
Public gbln������ϸʱʵ�ϴ� As Boolean
Public gblnסԺ��ϸʱʵ�ϴ� As Boolean
Public gbln������D As Boolean           '�����г�Ժ��������ʾ
Public gbln������K As Boolean           '��������Ժ��������ʾ
Public gblnDebug As Boolean             '������
Private mblnInit As Boolean     '�Ƿ񱻳�ʼ��

Private Function Readģ������(ByVal lng���Ĵ��� As Long, _
        msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
        str As String)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ���ù��ܶ�ȡģ������,��������
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strArr
    Dim strArr1
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    
    strFile = App.Path & "\����ҽ��ģ���ύ������" & lng���Ĵ��� & ".txt"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    
    objText.WriteLine msgType & Space(10) & packageType & Space(10) & packageLength & "|| " & str
    objText.Close
    
    If Dir(Left(App.Path, 18) & "\ҽ������\����ҽ��\����ҽ��ģ������" & lng���Ĵ��� & ".txt") <> "" Then
            Set objText = objFile.OpenTextFile(Left(App.Path, 18) & "\ҽ������\����ҽ��\����ҽ��ģ������" & lng���Ĵ��� & ".txt")
            Do While Not objText.AtEndOfStream
                strTemp = Trim(objText.ReadLine)
                strArr = Split(strTemp, "||")
                strArr1 = Split(strArr(0), "|")
                If Val(strArr1(0)) = msgType Then
                     str = strArr(1)
                     Exit Do
                End If
            Loop
            objText.Close
    End If
    
End Function
Public Function ��ȡ�������_����(ByVal lng���Ĵ��� As Long, ByVal intinsure As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵�������,������Ϣ����g�������_����
    '--�����:lng���Ĵ���(2��������)
    '--������:
    '--��  ��:��ȡ�ɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    
    Dim strInfor As String
    Dim rsTempset As New ADODB.Recordset
    Dim lngReturn As Long
    Dim int�Ա� As Integer
    
    ��ȡ�������_���� = False
    Err = 0
    On Error GoTo errHand:
    '�ܺ�ȫ���� 2003-12-17
    '����˴�������ո�ֵʱ�������������˴���ֱ���˳�
    strInfor = Space(100)
    If gblnģ��ӿ� Then
        Readģ������ lng���Ĵ���, 1001, 101, 94, strInfor
        If strInfor = "" Then Exit Function
    Else
        If lng���Ĵ��� = 2 Then
            '1001    101 95  ʵʱ�鿨���������鿨��
            lngReturn = KfqTransData(1001, 101, 95, strInfor, "com" & gintComPort_����)
        Else
            '1001    101 94  ʵʱ�鿨���������鿨��
            lngReturn = OltpTransData(1001, 101, 94, strInfor, "com" & gintComPort_����)
        End If
        If lngReturn <> 0 Or strInfor = "" Then
            ShowMsgbox GetErrInfo(CStr(lngReturn), intinsure)
            Exit Function
        End If
    End If
    'ȡ���ظ�
    strInfor = Mid(strInfor, 2)
    With g�������_����
        .ҽ������ = lng���Ĵ���
        If lng���Ĵ��� = 2 Then
            .���˱�� = Substr(strInfor, 1, 10)         '���˱���    1   10      ���ķ���
            .���� = Trim(Substr(strInfor, 11, 8))       '����    11  8       ���ķ���
            .���֤�� = Substr(strInfor, 19, 18)        '���֤��    19  18      ���ķ���
            .IC���� = Substr(strInfor, 37, 7)           'IC����  37  7       ���ķ���
            .������� = Val(Substr(strInfor, 44, 4))    '�������    44  4       ���ķ���
            .ְ����ҽ��� = Substr(strInfor, 48, 1)     'ְ����ҽ���    48  1   A��ְ��B����    ���ķ���
            .���������ʻ���� = Val(Substr(strInfor, 49, 10)) '���������ʻ����    49  10      ���ķ���
            .���������ʻ���� = Val(Substr(strInfor, 59, 10)) '���������ʻ����    59  10      ���ķ���
            .ͳ���ۼ� = Val(Substr(strInfor, 69, 10)) 'ͳ���ۼ�    69  10      ���ķ���
            .�½ɷѻ��� = Val(Substr(strInfor, 79, 10)) '�½ɷѻ���  79  10  �½ɷѹ���  ���ķ���
            .�ʻ�״̬ = Substr(strInfor, 89, 1) '�ʻ�״̬    89  1   A������B��ֹ����Cȫֹ����D����  ���ķ���
            .�α����1 = Substr(strInfor, 90, 1) '�α����1   90  1   �Ƿ����ܸ߶� 1 ���� 0 ������    ���ķ���
            .�α����2 = Substr(strInfor, 91, 1) '�α����2   91  1   �Ƿ����ܲ�������ҵ����������Ա������'0 ������ 1 ��ҵ 2 ����Ա    ���ķ���
            .�α����3 = Substr(strInfor, 92, 1) '�α����3   92  1   0 �󱣡�1 �±���2 ��������Ա(��������������)
            .�α����4 = Substr(strInfor, 93, 1) '�α����4   93  1   ����    ���ķ���
            .�α����5 = Substr(strInfor, 94, 1) '�α����5   94  1   ����    ���ķ���
        Else
            .���˱�� = Substr(strInfor, 1, 8)  '���˱��    CHAR    1   8   ҽ�����    ����
            
            ' 2004-09-23    �ܺ�ȫ
            '���ڲ��Կ�������Ϊȫ���֣���Ҫ����Ϊ����
            .���� = Trim(Substr(strInfor, 9, 8))                            '����    CHAR    9   8       ����
            .���� = IIf(IsNumeric(.����), Trim(.����) & "����", .����)      '����    CHAR    9   8       ����
            
            .���֤�� = Substr(strInfor, 17, 18)    '���֤��    CHAR    17  18  18λ��15λ  ����
            .IC���� = Substr(strInfor, 35, 7)       'IC����  NUM 35  7       ����
            .������� = Val(Substr(strInfor, 42, 4))    '�������    NUM 42  4       ����
            
            '�ܺ�ȫ���� 2003-12-17
            '���룺Q��ҵ����
            .ְ����ҽ��� = Substr(strInfor, 46, 1)     'ְ����ҽ���    CHAR    46  1   A��ְ��B���ݡ�L���ݡ�T���Q��ҵ����  ����
            .���������ʻ���� = Val(Substr(strInfor, 47, 10))   '���������ʻ����    NUM 47  10      ����
            .���������ʻ���� = Val(Substr(strInfor, 57, 10))   '���������ʻ����    NUM 57  10  �����ڹ���Ա��������    ����
            .ͳ���ۼ� = Val(Substr(strInfor, 67, 10))   'ͳ���ۼ�    NUM 67  10      ����
            .�½ɷѻ��� = Val(Substr(strInfor, 77, 10)) '�½ɷѻ���  NUM 77  10  �½ɷѹ���  ����
            .�ʻ�״̬ = Substr(strInfor, 87, 1)         '�ʻ�״̬    CHAR    87  1   A������B��ֹ����Cȫֹ����D����  ����
            .�α����1 = Substr(strInfor, 88, 1)        '�α����1   CHAR    88  1   �Ƿ����ܸ߶�: 0 �����ܸ߶1 ���ܸ߶2 ҽ�Ʊ��ղ�����    ����
            .�α����2 = Substr(strInfor, 89, 1)        '�α����2   CHAR    89  1   �Ƿ����ܲ�������ҵ����������Ա������0 ������ 1 ��ҵ 2 ����Ա    ����
            .�α����3 = Substr(strInfor, 90, 1)        '�α����3   CHAR    90  1   0 �󱣡�1 �±�  ����
            .�α����4 = Substr(strInfor, 91, 1)        '�α����4   CHAR    91  1   0���������á�1��������  ����
            .�α����5 = Substr(strInfor, 92, 1)        '�α����5   CHAR    92  1   0���˲����á�1���˿���  ����
        End If
        
        '��ȡ�����ʻ��ĵ�ǰ״̬
        gstrSQL = "select ��ǰ״̬ from �����ʻ� where ����=" & Decode(lng���Ĵ���, 2, 83, 1, 82) & " and ҽ����='" & g�������_����.���˱�� & "'"
        zlDatabase.OpenRecordset rsTempset, gstrSQL, "��ȡ�����ʻ���ǰ״̬"
        If Not rsTempset.EOF Then
            g�������_����.��ǰ״̬ = rsTempset!��ǰ״̬
        End If
        
        int�Ա� = Val(IIf(Len(.���֤��) = 18, Mid(.���֤��, 17, 1), Right(.���֤��, 1))) Mod 2
        '�������֤ȡ����Ӧ���Ա�
        .�Ա� = IIf(int�Ա� = 0, "Ů", "��")
        .�������� = zlCommFun.GetIDCardDate(Trim(.���֤��))
        '��������
        If IsDate(.��������) And .�������� <> "" Then
            '.���� = Abs(Int((zlDatabase.Currentdate - CDate(.��������)) / 366))
            gstrSQL = "Select Months_between(Trunc(SysDate),To_Date('" & .�������� & "','YYYY-MM-DD'))/12 As ���� From Dual"
            zlDatabase.OpenRecordset rsTempset, gstrSQL, "��ȡ����"
            If rsTempset.RecordCount <= 0 Then
                .���� = 0
            Else
                .���� = rsTempset!����
            End If
        Else
            .���� = 0
        End If
        
    End With
    ��ȡ�������_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    ��ȡ�������_���� = False
End Function

Public Function ҵ������_����( _
            ByVal lng���Ĵ��� As Long, _
            ByVal lngMsgType As Long, _
            strTans As String, _
            ByVal intinsure As Integer _
    ) As Boolean
    
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:����ص�ҵ������,��������Ӧ�Ľ��
    '--�����:lng���Ĵ���(2��������)
    '   lngMsgType-ҵ����������
    '   lngPackageType-���ݽ�����ʽ����
    '   lngPackageLength-���ݴ��ĳ���
    '   strTans-���ݴ�,����ʱ��ͨ�����ݴ������������������ʱ�����ݴ��а������ص�����
    '����:
    '   �ɹ�-true,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim lngPackageType As Long
    Dim lngPackageLength As Long
    Dim i As Long
    Dim strTmp As String
    Dim strReg As String
    
    i = lngMsgType
    
    '����Ϊ�������Ĳ�����
'    'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
'    '------------    ----------------    --------------         -----------------------------------------------
'    '1001            101                 95                     ʵʱ�鿨���������鿨��
'    '1002            12                  420                    ʵʱ����
'    '1003            7                   297                    ʵʱҽ����ϸ�����ύ
'    '1004            9                   136                    ʵʱסԺ�Ǽ������ύ
'    '1006            12                  420                    ʵʱ����Ԥ��
'    '1008            101                 95                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�

    '2005-08-02 �ܺ�ȫ
    '�����������޸�
    'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
    '------------    ----------------    --------------         -----------------------------------------------
    '1001            101                 96                     ʵʱ�鿨���������鿨��
    '1002            12                  509                    ʵʱ����
    '1003            7                   240                    ʵʱҽ����ϸ�����ύ
    '1004            9                   210                    ʵʱסԺ�Ǽ������ύ
    '1006            12                  509                    ʵʱ����Ԥ��
    '1008            101                 96                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�
    
    '����Ϊ�����еĲ�����
    'ҵ����������    ���ݽ�����ʽ����    ���ݴ���С����         ˵��
    '------------    ----------------    --------------         -----------------------------------------------
    '1001            101                 94                     ʵʱ�鿨���������鿨��
    '1002            12                  424                    ʵʱ����
    '1003            7                   230                    ʵʱҽ����ϸ�����ύ
    '1004            9                   206                    ʵʱסԺ�Ǽ������ύ
    '1006            12                  424                    ʵʱ����Ԥ��
    '1008            101                 94                     ʵʱ��ѯ��ֱ�Ӳ�ѯ�������ݣ�
    '1005            8                   274                    ʵʱҽ������
    '1007            2                   55                     �����ʻ���ѯ
    
    Dim strInfor As String
    Dim strSQL As String
    Dim lngReturn As Long
    ҵ������_���� = False
    Err = 0
    On Error Resume Next
    If lng���Ĵ��� = 2 Then
        strTmp = Switch(i = 1001, "101|96", i = 1002, "12|509", i = 1003, "7|240", i = 1004, "9|210", i = 1006, "12|509", _
            i = 1008, "101|96")
        If Err <> 0 Then
            strTmp = "|"
        End If
    Else
            strTmp = Switch(i = 1001, "101|94", i = 1002, "12|475", i = 1003, "7|230", i = 1004, "9|206", i = 1006, "12|475", _
                i = 1008, "101|94", i = 1005, "8|221", i = 1007, "2|55")
        If Err <> 0 Then
            strTmp = "|"
        End If
    End If
    lngPackageType = Val(Split(strTmp, "|")(0))
    lngPackageLength = Val(Split(strTmp, "|")(1))
    
    Err = 0
    On Error GoTo errHand:
    strInfor = strTans
    If gblnģ��ӿ� Then
        Readģ������ lng���Ĵ���, lngMsgType, lngPackageType, lngPackageLength, strInfor
        If strInfor = "" Then
            strTans = strInfor
            Exit Function
        End If
    Else
        '������˵,������ҵ�����͵������ж���ǰ�ӿո�.�����ؼ���:" " &
        strInfor = " " & strInfor
        
        '----��ʱΪ�л����ݽ��зָ����ݲ����ṩ����ϵͳ���м���
        If lngMsgType = 1006 Then
            If Val(GetSetting("ZLSOFT", "����ģ��\zl9insure\����", "�ָ����", 0)) = 1 Then
                strSQL = "insert into ��������ָ�(ҵ������,����,����ʱ��,��Ϣ��) values(" & lngMsgType & ",Substrb(substr('" & strInfor & "',2,1000), 14, 7),sysdate,'" & strInfor & "')"
                gcnOracle.Execute strSQL
                MsgBox "�ָ����ݲ����ɹ�,�˳�Ԥ����!"
                ҵ������_���� = False
                Exit Function
            End If
        End If
        If lng���Ĵ��� = 2 Then
            lngReturn = KfqTransData(lngMsgType, lngPackageType, lngPackageLength, strInfor, "com" & gintComPort_����)
        Else
            lngReturn = OltpTransData(lngMsgType, lngPackageType, lngPackageLength, strInfor, "com" & gintComPort_����)
        End If
        If lngReturn <> 0 Or strInfor = "" Then
            ShowMsgbox GetErrInfo(CStr(lngReturn), intinsure)
            strTans = ""
            Exit Function
        End If
    End If
    'ȡ���ظ�
    strInfor = Mid(strInfor, 2)
    
    strTans = strInfor
    ҵ������_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    strTans = ""
    ҵ������_���� = False
End Function


Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '--�����:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '--������:
    '--��  ��:�Ӵ�
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    Err = 0
    On Error GoTo errHand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
'    strTmp = Right(Substr, 1)
'    If zlCommFun.ActualLen(strTmp) = 1 Then
'        If asc(strTmp) < 32 Or asc(strTmp) > 126 Then
'            Substr = Left(Substr, Len(Substr) - 1)
'        End If
'    End If
    Exit Function
errHand:
    Substr = ""
End Function

Public Function ҽ����ʼ��_����(ByVal intinsure As Integer) As Boolean

    Dim rsTemp  As New ADODB.Recordset
    Dim strReg As String
    
    '���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
    '���أ���ʼ���ɹ�������true�����򣬷���false
    
    On Error Resume Next
    Err = 0
    On Error GoTo 0
    
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & intinsure
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡҽԺ����")
    gstrҽԺ����_���� = Nvl(rsTemp!ҽԺ����, "")
    
    '���ö˿ں�
    Call GetRegInFor(g����ģ��, "����", "�˿ں�", strReg)

    If Val(strReg) = 0 Then
        gintComPort_���� = 1
    Else
        gintComPort_���� = IIf(Val(strReg) > 99, 1, Val(strReg))
    End If
    
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        gblnģ��ӿ� = True
    Else
        gblnģ��ӿ� = False
    End If
    Call GetRegInFor(g����ģ��, "����", "������", strReg)
    
    If intinsure = TYPE_���������� Then
        gblnKFQCom_���� = True
    Else
        gblnKFQCom_���� = False
    End If
    Call GetRegInFor(g����ģ��, "����", "����", strReg)
    
    gblnDebug = strReg = "1"
    
    '�����ϴ���ϸ����
    gstrSQL = "Select * From ���ղ��� where ������ in ('������ϸʱʵ�ϴ�','סԺ��ϸʱʵ�ϴ�','�����ֳ�Ժ��ʾ') and ����=" & intinsure
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ղ���"
    gbln������ϸʱʵ�ϴ� = True
    gblnסԺ��ϸʱʵ�ϴ� = True
    
    '�����Ժ��ϵ�������ʾ
    Call GetRegInFor(g����ģ��, "����", "�����е�����", strReg)
    gbln������D = strReg = "1"
    Call GetRegInFor(g����ģ��, "����", "������������", strReg)
    gbln������K = strReg = "1"

    Do While Not rsTemp.EOF
        Select Case Nvl(rsTemp!������)
        Case "������ϸʱʵ�ϴ�"
            gbln������ϸʱʵ�ϴ� = IIf(Val(Nvl(rsTemp!����ֵ)) = 1, True, False)
        Case "סԺ��ϸʱʵ�ϴ�"
            gblnסԺ��ϸʱʵ�ϴ� = IIf(Val(Nvl(rsTemp!����ֵ)) = 1, True, False)
'        Case "ҽ����ϸʱʵ�ϴ�"
'            gblnҽ����ϸʱʵ�ϴ� = IIf(Val(Nvl(rsTemp!����ֵ)) = 1, True, False)
        Case "�����ֳ�Ժ��ʾ"
            If intinsure = 82 Then
                gbln������D = IIf(Val(Nvl(rsTemp!����ֵ)) = 1, True, False)
            Else
                gbln������K = IIf(Val(Nvl(rsTemp!����ֵ)) = 1, True, False)
            End If
        End Select
        rsTemp.MoveNext
    Loop
    
    mblnInit = True
    ҽ����ʼ��_���� = True
End Function

Public Function �������_����(ByVal lng����ID As Long, ByVal intinsure As Integer) As Currency
    '����: ���ݲ���idȡ�����
    '����: ����id
    '����: ���ظ����ʻ����
    Dim rsAcc As New ADODB.Recordset
    
    
    '����ʧ�����˳�
    gstrSQL = "Select Nvl(�ʻ����,0) �ʻ����,����֤�� From �����ʻ� Where ����=[1] And ����id=[2]"
    
    Set rsAcc = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ����", intinsure, lng����ID)
    
    With g�������_����
        .���������ʻ���� = Nvl(rsAcc!�ʻ����, 0)
        .���������ʻ���� = Val(Nvl(rsAcc!����֤��))
        �������_���� = .���������ʻ���� + .���������ʻ����
    End With
'    Call WriteDebugInfor_����("�������_���� ", lng����id)
End Function

Public Function ҽ������_����(ByVal lng���� As Long, ByVal lngҽ������ As Integer) As Boolean
    ҽ������_���� = frmSet����.ShowME(lng����, lngҽ������)
    'Call WriteDebugInfor_����("ҽ������_���� ", lng����id)
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long, Optional intinsure As Integer) As String
    Dim str��ע As String, RSPATIENT As New ADODB.Recordset
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    
    ��ݱ�ʶ_���� = frmIdentify����.GetPatient(intinsure, bytType, lng����ID)
'    Call WriteDebugInfor_����("��ݱ�ʶ_���� byttype:" & bytType, lng����id)
    
End Function
Public Function ��ݱ�ʶ_����2(ByVal strCard As String, ByVal strPass As String, Optional lng����ID As Long, Optional intinsure As Integer) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    '/**?
    ��ݱ�ʶ_����2 = frmIdentify����.GetPatient(intinsure, 3, lng����ID)
'    Call WriteDebugInfor_����("��ݱ�ʶ_����2", lng����id)
    
End Function

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '��Ҫ�пո������
        strTmp = Substr(strCode, 1, lngLen)
    End If
    'ȡ��������ַ�
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function
Private Function Get�������(ByVal bytҵ�� As Byte, ByVal int���� As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��������ʶ
    '--�����:bytҵ��-(0-����,1-����)
    '         int���� ����:(1-��ͨ����,2-��������,3-�����,4-������������)
    '                 סԺ:(5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ)
    '--������:
    '--��  ��:ҽ�����ĵķ����ʶ
    '-----------------------------------------------------------------------------------------------------------
    'ҽ�����ĵľ������Ķ�Ӧϵͳ
    
    '1 �������
    'A ����������
    '3 �������
    '7 ����������
    '5 ����󲡽���
    'B ����󲡳���
    'S ������������
    'T ������������
    
    '2 סԺ����
    'D סԺ�����
    '9 סԺ�岹��  �˹����ݲ���
    '4 ��ͥ��������
    'C ��ͥ���������
    '8 ��ͥ��������     '�˹����ݲ���
    'O ��������סԺ����
    'P ��������סԺ����
    'Q ���˱��ս���
    'R ���˱��ճ���


    Dim i As Integer
    Dim strTmp As String
    i = int����
    
    '���˺��ע:200404
    '     ����:1-1,2-3,3-5,4-"S"
    '     סԺ:5-2,6-4,7-"O",8-"Q"
            
            
    Select Case int����
        Case 1  '1-��ͨ����
            strTmp = Decode(bytҵ��, 0, "1", "A")
        Case 2  '2-��������
            strTmp = Decode(bytҵ��, 0, "3", "7")
        Case 3  '3-�����
            strTmp = Decode(bytҵ��, 0, "5", "B")
        Case 4  '4-������������
            strTmp = Decode(bytҵ��, 0, "S", "T")
        Case 5  '5-��ͨסԺ,
            strTmp = Decode(bytҵ��, 0, "2", "D")
        Case 6  '6-��ͥ����סԺ,
            strTmp = Decode(bytҵ��, 0, "4", "C")
        Case 7  '7-��������סԺ
            strTmp = Decode(bytҵ��, 0, "O", "P")
        Case 8  '8-���˱���סԺ
            strTmp = Decode(bytҵ��, 0, "Q", "R")
        Case Else
            strTmp = ""
    End Select
    Get������� = strTmp
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, ByVal intinsure As Integer) As Boolean
    Dim curTotal As Double, cur�����ʻ� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    
    Dim strInfor As String  '�������ķ��ش�
    Dim dbl���� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double
    
    '����������
    Dim dblҩ���Է� As Double
'    Dim dbl����ǰ����ͳ���ۼ� As Double
'    Dim dbl����󼲲�ͳ���ۼ� As Double
    
    Dim dbl��ҩ�� As Double, dbl���� As Double, dbl���Ʒ� As Double
    Dim dbl���� As Double, dbl����Է� As Double
    Dim dbl�������Ʒ� As Double, dbl���������Է� As Double
    Dim dbl�������Էѷ��� As Double, dbl�Ǳ��շ��� As Double
    Dim dbl������ As Double    '��Դ�����������
    Dim dblѪ�� As Double, dblѪ���Է� As Double
    Dim dblͳ����� As Double, dbl�𸶱�׼ As Double
    Dim lng����ID As Long
    
    Dim strҽʦ���� As String
    Dim str����Ա���� As String
    Dim str���������ʶ As String
    Dim strTmp As String
    Dim strҽ�� As String
    Dim str��ϸ As String       '��ϸ��
    Dim str���ұ��� As String
    Dim dbl���� As Double
    Dim str��Ŀͳ�Ʒ��� As String
    Dim str��Ŀ���� As String
    Dim dbl��Ŀ���� As Double
    Static str����ʱ�� As String
    Static lng����id1 As Long
On Error GoTo ErrH
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '��ϸ�ֶ�
    '   ����ID,�շ����,�վݷ�Ŀ,���㵥λ,������,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��,ժҪ,�Ƿ���
    
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ�
    
    '������֧��������ڱ���,�Ա���㱣�Ѽ��Է�
    gstrSQL = "Select * From ����֧������"
    zlDatabase.OpenRecordset rs����, gstrSQL, "����֧������"
    
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lng����ID As Long
    With rs��ϸ
    
        '��Ҫ������շѵ���
        If str����ʱ�� <> Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS") Or lng����id1 <> Nvl(!����ID, 0) Then
              str����ʱ�� = Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS")
              lng����id1 = Nvl(!����ID, 0)
              g�������_����.���㿪ʼ = True
        Else
              g�������_����.���㿪ʼ = False
        End If
    
        
        'ȷ������
        If Not .EOF Then
            lng����ID = Nvl(!����ID, 0)
            gstrSQL = "  select ����id from �����ʻ� where ����id=" & lng����ID & "  and ����=" & intinsure & "  and ҽ����='" & g�������_����.���˱�� & "'"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
            If Not rsTemp.EOF Then
                lng����ID = Nvl(rsTemp!����ID, 0)
            Else
                lng����ID = 0
            End If
          '����׼��Ŀ
            gstrSQL = "Select * from ������׼��Ŀ  where ����ID=  " & lng����ID
            zlDatabase.OpenRecordset rs��׼��Ŀ, gstrSQL, "��ȡ������Ŀ����"
            
        End If
        
        'ȡ�����η������õĽ��ϼ�
        Do While Not .EOF
            '---��˳��,�Խ���Ƿ�Ϊ���������ж�,���Ϊ������׼ִ��ҽ���շ�
            If !ʵ�ս�� < 0 Then
                ShowMsgbox "�õ����а����н��Ϊ��������Ŀ,����ִ��ҽ���շ�!����������շ�"
                �����������_���� = False
                Exit Function
            End If
            
            If lng����ID <> 0 Then
                    '��һ��,ȷ��������շ�ϸĿ
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=0 And ����=1 and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
                    If rs��׼��Ŀ.EOF Then
                        gstrSQL = "Select ����,���� from �շ�ϸĿ where id=" & Nvl(!�շ�ϸĿID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�շ�ϸĿ"
                        Err.Raise 9000, gstrSysName, "�շ�ϸĿΪ��" & Nvl(rsTemp!����) & "������Ŀ���ǲ��������趨����Ŀ."
                        Exit Function
                    End If
                    
                    '�ڶ���,ȷ������ı��մ���
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=1 And ����=1 and  �շ�ϸĿid=" & Nvl(!����֧������ID, 0)
                    If rs��׼��Ŀ.EOF Then
                        Err.Raise 9000, gstrSysName, "�ڽ����д����˽�������ı���֧������,���ܼ�����"
                        Exit Function
                    End If
                    '������,'ȷ����ֹ���շ�ϸĿ
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=0 And ����=2 and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
                    If Not rs��׼��Ŀ.EOF Then
                        gstrSQL = "Select ����,���� from �շ�ϸĿ where id=" & Nvl(!�շ�ϸĿID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�շ�ϸĿ"
                        Err.Raise 9000, gstrSysName, "�շ�ϸĿΪ��" & Nvl(rsTemp!����) & "������Ŀ�Ǳ���ֹʹ�õ���Ŀ." & vbCrLf & "���ܼ���!"
                        Exit Function
                    End If
                    '���Ĳ�,'ȷ����ֹ�Ĵ���
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=1 And ����=2 and �շ�ϸĿid=" & Nvl(!����֧������ID, 0)
                    If Not rs��׼��Ŀ.EOF Then
                        Err.Raise 9000, gstrSysName, "�ڽ����д����˽�ֹʹ�õı���֧������,���ܼ�����"
                    End If
            End If
        
            '���ж��Ƿ�������ҽ����Ӧ��Ŀ����
            gstrSQL = " Select ��Ŀ����,��Ŀ���� From ����֧����Ŀ" & _
                      " Where ����=[1] And �շ�ϸĿID=[2]"
                      
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ������˶�Ӧ��ҽ����Ŀ", intinsure, CLng(!�շ�ϸĿID))
            If rsTemp.EOF = True Then
                Err.Raise 9000, gstrSysName, "����Ŀδ����ҽ����Ŀ�����ܽ��㡣"
                Exit Function
            End If
            If strҽ�� = "" Then
                strҽ�� = Nvl(!������)
            End If
            
            str��Ŀ���� = Nvl(rsTemp!��Ŀ����)
            dbl��Ŀ���� = Val(Nvl(rsTemp!��Ŀ����)) / 100
            lng����ID = Nvl(!����ID, 0)
            
            gstrSQL = "" & _
                " Select b.������,b.����ֵ from �շ���� a,���ղ��� b " & _
                " Where a.���=b.������ and b.����=" & intinsure & _
                "        and a.����='" & Nvl(!�շ����) & "'"
            
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "���Ѽ���"
            
            If rsTemp.EOF Then
                strTmp = ""
            Else
                strTmp = Nvl(rsTemp!����ֵ)
            End If
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '���㱣��
                rs����.Find "id=" & Nvl(!����֧������ID, 0), , adSearchForward, 1
                If Not rs����.EOF Then
                    '2005-10-17 ZHQ
                    '�ж�����󲡻�������������ʹ��סԺ����
                    If (g�������_����.������� = 3 And IsParaBig(intinsure)) Or _
                        (IsParaQ(intinsure) And intinsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q") Then
                        dblͳ����� = Nvl(rs����!סԺ�ȶ�, 0) / 100
                    Else
                        dblͳ����� = Nvl(rs����!ͳ��ȶ�, 0) / 100
                    End If
                Else
                    dblͳ����� = 1
                End If
                
                '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,Q��ҵ����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
                If intinsure <> TYPE_���������� And g�������_����.ְ����ҽ��� = "L" _
                    And g�������_����.�α����3 = "0" And Nvl(!�Ƿ�ҽ��, 0) = 1 Then  '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '  ������  ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�����ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dblͳ����� = 1
                End If
                
                If intinsure = TYPE_������ And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dblͳ����� = dbl��Ŀ����
                End If
                
                If intinsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                    '�����Q��ҵ����,�������Ϊ100�Է�,�������Ǳ��շ�����
                    If dblͳ����� = 0 Then
                        '�Է�100
                        strTmp = ""
                    Else
                        '�ԷѲ��ַ��� �������Էѷ�����
                    End If
                End If
                                
                '�ܺ�ȫ���� 2003-12-17
                '����������Ŀ��ֻҪ�Ǳ�ʶΪ�����Ρ��ģ���Ӧ�����������
                'If NVL(!�շ����) = "����" And str��Ŀ���� = "����" Then
                If str��Ŀ���� = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If str��Ŀ���� = "���" Then
                    strTmp = "����"
                End If
                '����۳��ԷѲ��ֵķ���.��Ϊֻ�д�λ�����Ƕ����,���������ﲻ�ᷢ����λ����
                'һ�����﷢����λ����,������ͳ�����=0���м���
                If Not rsTemp.EOF Then
                    If dblͳ����� <> 0 Then
                        Select Case strTmp
                            Case "����"
                                
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս��, 0) * dblͳ�����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                               
                            Case "��ҩ��"
                                
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dblͳ�����, 5)
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                                Else    '����������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                                End If
                            Case "��ҩ��"
                                
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dblͳ�����, 5)
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                                Else    '����������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                                End If
                            Case "��ҩ��"
                                
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dblͳ�����, 5)
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                                Else    '����������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                                End If
                               
                            Case "����"
                                
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս��, 0) * dblͳ�����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                                
                            Case "���Ʒ�"
                                
                                dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!ʵ�ս��, 0) * dblͳ�����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)

                            Case "����"
                                '�����кͿ������Դ����ô���ͬ,
                                '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                          
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս�� * dblͳ�����, 0), 5)
                                
                                dbl����Է� = dbl����Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                                
                            Case "Ѫ��"
                                '�����кͿ������Դ����ô���ͬ,
                                '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                          
                                dblѪ�� = dblѪ�� + Round(Nvl(!ʵ�ս�� * dblͳ�����, 0), 5)
                                dblѪ���Է� = dblѪ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                            Case "�������Ʒ�"
                                '2004/9/11��ǰ:�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                                '2004/9/11�Ժ�:�������뿪�������㷽ʽһ�£�����ͳ�ﲿ�ֵĽ��
                                'If intinsure = TYPE_������ Then
                                '   dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!ʵ�ս��, 0), 5)
                                'Else
                                    dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!ʵ�ս��, 0) * dblͳ�����, 5)
                                'End If
                                
                                '�������Ʒ��Էѵļ��㷽ʽ��ͬ,
                                '�����нӿ��ڴ�����ܽ��ʱֻ���������Ʒѽ��л���,���������ԷѲ��ֲ��ټ���
                                dbl���������Է� = dbl���������Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                            
                        End Select
                    Else
                        'ȫ���Ǳ���Ϊ0����Ŀ(��������),�ֱ�Դ������ǿ����������жϷ��ڲ�ͬ���ֶ�
                        If intinsure = TYPE_������ Then
                            '�����з���dbl�Ǳ��շ���
                            dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(!ʵ�ս��, 5)
                        Else
                            '����������dbl������
                            dbl������ = dbl������ + Round(!ʵ�ս��, 5)
                        End If
                    
                    End If

                End If
            End If
            curTotal = curTotal + Round(Nvl(!ʵ�ս��, 0), 5)
            .MoveNext
        Loop
    End With
    
    '��������
    If strҽ�� <> "" Then
        gstrSQL = "Select ��� From ��Ա��  where ����='" & strҽ�� & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽ�����"
        If Not rsTemp.EOF Then
            strҽ�� = Nvl(rsTemp!���)
            If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                strҽ�� = Substr(strҽ��, 1, 6)
            End If
        Else
            strҽ�� = ""
        End If
    End If
    
    '���������Ϊ��
    If ����ҽ������(intinsure, 0, lng����ID, 0, 0, 0, False, True, 0, dbl����, dbl��ҩ��, dbl��ҩ��, dbl��ҩ��, dbl����, dbl���Ʒ�, dblѪ��, dblѪ���Է�, dbl����, dbl����Է�, dbl�������Ʒ�, dbl���������Է�, dbl�������Էѷ���, dbl�Ǳ��շ���, dbl������, dblҩ���Է�, curTotal, strҽ��, str���㷽ʽ) = False Then
        Exit Function
    End If
'    Call WriteDebugInfor_����("�����������_����", lng����id)
    �����������_���� = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function Get���㷽ʽ(ByVal strOutput As String, ByVal intinsure As Integer) As String
    '����:��ȡ���㷽ʽ
    '����:strOutPut-�����
    '����:���㷽ʽ
    '������:
    '    ���λ��������ʻ�֧��    NUM 225 10      ���ķ���
    '    ���β��������ʻ�֧��    NUM 235 10      ���ķ���
    '    ���λ���ͳ��֧��    NUM 245 10      ���ķ���
    '    ���λ���ͳ���Ը�    NUM 255 10      ���ķ���
    '    ���β���ͳ��֧��    NUM 265 10      ���ķ���
    '    ���β���ͳ���Ը�    NUM 275 10      ���ķ���
    '    ���λ�����������֧��    NUM 285 10  ����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
    '    ���ηǻ�����������֧��  NUM 295 10  ����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
    '    ���α��շ�Χ���Ը�  NUM 305 10  �޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
 
    Dim i As Long
    Dim str���㷽ʽ As String
    
    If intinsure = TYPE_���������� Then
    
        '���� 2005-08-16 ������ҽ������
        'i = 225 - 10
         i = 275 - 10
        'ȷ�����ν��㷽ʽ
        str���㷽ʽ = "�����ʻ�;" & Format(Val(Substr(strOutput, i + 10, 10)) + Val(Substr(strOutput, i + 20, 10)), "###0.00;-###0.00;0;0") & ";0" '���λ��������ʻ�֧��,�������޸�
        str���㷽ʽ = str���㷽ʽ & "|" & "����ͳ��;" & Format(Val(Substr(strOutput, i + 30, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
        str���㷽ʽ = str���㷽ʽ & "|" & "����ͳ��;" & Format(Val(Substr(strOutput, i + 50, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
        
        '���� 2005-08-16 ������ҽ������,ȡ���˲������պͷǲ�������,�����˹���Ա��������ҵ���ղ���
        '����Ա����:=���ι���Ա�𸶱�׼����֧��+���ι���Ա������������֧��+���ι���Ա�ǻ�����������֧��
        'Val (Substr(strOutPut, i + 70, 10)) + Val(Substr(strOutPut, i + 80, 10)) + Val(Substr(strOutPut, i + 90, 10))
        'str���㷽ʽ = str���㷽ʽ & "|" & "��������;" & Format(Val(Substr(strOutPut, i + 70, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
        'str���㷽ʽ = str���㷽ʽ & "|" & "�ǲ�������;" & Format(Val(Substr(strOutPut, i + 80, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
        str���㷽ʽ = str���㷽ʽ & "|" & "����Ա����;" & Format(Val(Substr(strOutput, i + 70, 10)) + Val(Substr(strOutput, i + 80, 10)) + Val(Substr(strOutput, i + 90, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
        str���㷽ʽ = str���㷽ʽ & "|" & "��ҵ���ղ���;" & Format(Val(Substr(strOutput, i + 100, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�

        Get���㷽ʽ = str���㷽ʽ
        Exit Function
    End If
    
   '������:
    '   3.0.2��
    '    ���λ��������ʻ�֧��    NUM 211 10  ������������㣬��ʾ�����ʻ�֧��    ����
    '    ���β��������ʻ�֧��    NUM 221 10  ������������㷵��0 ����
    '    ���λ���ͳ��֧��    NUM 231 10      ����
    '    ���λ���ͳ���Ը�    NUM 241 10      ����
    '    ���β���ͳ��֧��    NUM 251 10  ������������㣬���ֶ����ڴ����������֧��  ����
    '    ���β���ͳ���Ը�    NUM 261 10      ����
    '    ���λ�����������֧��    NUM 271 10  1�� �������ҵ���ո��ֶΰ�������ͳ���Ը����ֵ���ҵ����֧�� 2�� ����ǹ���Ա�������ֶΰ����ż��Ѳ������֡�����ͳ���Ը����ֵĹ���Ա����֧��������ͳ������޶��ڹ���Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����  ����
    '    ���ηǻ�����������֧��  NUM 281 10  1�� �������ҵ���ո��ֶ��ǲ���ͳ���Ը����ֵ���ҵ����֧��   2�� ����ǹ���Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧������������ͳ������޶��Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����    ����
    '    ���α��շ�Χ���Ը�  NUM 291 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣��������Էѷ��ã��Ǳ��շ���+����Է�   ����
    '    ����Ԥ��ʱ�������ʻ��Ͳ��������ʻ��ϼƺ���ΪԤ����ʱ�ʻ�֧�������ж�ֵ
    '   4.0.1��
    '   ���λ��������ʻ�֧��    NUM 241 10  ������������㣬��ʾ�����ʻ�֧��
    '   ���β��������ʻ�֧��    NUM 251 10  ������������㷵��0
    '   ���λ���ͳ��֧��    NUM 261 10
    '   ���λ���ͳ���Ը�    NUM 271 10
    '   ���β���ͳ��֧��    NUM 281 10  ������������㣬���ֶ����ڴ����������֧��
    '   ���β���ͳ���Ը�    NUM 291 10
    '   ���ι���Ա�𸶱�׼����֧��  NUM 301 10
    '   ���ι���Ա������������֧��  NUM 311 10
    '   ���ι���Ա�ǻ�����������֧��    NUM 321 10
    '   ������ҵ���ղ���֧��    NUM 331 10
    '   ���α������Ը�  NUM 341 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣����������Էѷ���+����Է�+Ѫ���Է�+�����Էѣ��ĸ����˻���ֺ�ķ���
    '   ���ηǱ����Ը�  NUM 351 10
    i = 241 - 10
    Dim dblMoney As Double
    With g�������_����
        .��ǰ�����ʻ���� = .��ǰ�����ʻ���� - Format(Val(Substr(strOutput, i + 10, 10)), "###0.00;-###0.00;0;0")
        .��ǰ�����ʻ���� = .��ǰ�����ʻ���� - Format(Val(Substr(strOutput, i + 20, 10)), "###0.00;-###0.00;0;0")
        .��ǰͳ���ۼ� = .��ǰͳ���ۼ� + Format(Val(Substr(strOutput, i + 30, 10)), "###0.00;-###0.00;0;0") + Format(Val(Substr(strOutput, i + 50, 10)), "###0.00;-###0.00;0;0")
    End With
    
    'ȷ�����ν��㷽ʽ
    str���㷽ʽ = "�����ʻ�;" & Format(Val(Substr(strOutput, i + 10, 10)) + Val(Substr(strOutput, i + 20, 10)), "###0.00;-###0.00;0;0") & ";0" '���λ��������ʻ�֧��,�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "����ͳ��;" & Format(Val(Substr(strOutput, i + 30, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "����ͳ��;" & Format(Val(Substr(strOutput, i + 50, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    '����Ա����:=���ι���Ա�𸶱�׼����֧��+���ι���Ա������������֧��+���ι���Ա�ǻ�����������֧��
    dblMoney = Val(Substr(strOutput, i + 70, 10)) + Val(Substr(strOutput, i + 80, 10)) + Val(Substr(strOutput, i + 90, 10))
    
    '2004/09/11���˺�,ȡ���˲������պͷǲ�������,�����˹���Ա��������ҵ���ղ���
    str���㷽ʽ = str���㷽ʽ & "|" & "����Ա����;" & Format(dblMoney, "###0.00;-###0.00;0;0") & ";0" '�������޸�
    str���㷽ʽ = str���㷽ʽ & "|" & "��ҵ���ղ���;" & Format(Val(Substr(strOutput, i + 100, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    
    'str���㷽ʽ = str���㷽ʽ & "|" & "��������;" & Format(Val(Substr(strOutPut, i + 70, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    'str���㷽ʽ = str���㷽ʽ & "|" & "�ǲ�������;" & Format(Val(Substr(strOutPut, i + 80, 10)), "###0.00;-###0.00;0;0") & ";0" '�������޸�
    '������ҵ���ݲ��˽��㷽ʽΪ'����ͳ��',���Ҷ����ݲ��˵�����ͳ��=�Ǳ��շ���
    If g�������_����.ְ����ҽ��� = "Q" Then
        'str���㷽ʽ = str���㷽ʽ & "|" & "���ݲ���;" & Format(Round(dbl���� + dbl��ҩ�� + dbl��ҩ�� + dbl��ҩ�� + dbl���� + dbl���Ʒ� + dbl���� + dbl�������Ʒ� + dbl����Է� + dbl�������Էѷ���, 2), "###0.00;-###0.00;0;0") & ";0" '�������޸�
        '����"���ηǱ����Ը�"�����������"�Ը�"Ϊ��β���ֶ�֮�ϼ�ΪҽԺӦ���滼�߸����ķ���
        '���ݲ���=���λ���ͳ���Ը�+���β���ͳ���Ը�+���α������Ը�
        dblMoney = Val(Substr(strOutput, i + 40, 10)) + Val(Substr(strOutput, i + 60, 10)) + Val(Substr(strOutput, i + 110, 10))
        str���㷽ʽ = str���㷽ʽ & "|" & "���ݲ���;" & Format(dblMoney, "###0.00;-###0.00;0;0") & ";0" '�������޸�
    End If
    Get���㷽ʽ = str���㷽ʽ
    '�ֽ�=���ηǱ����Ը�
End Function
Public Function ��������������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, ByVal intinsure As Integer) As Boolean
    '������rsExse     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '��ϸ�ֶ�
    '��¼����,��¼״̬,NO,���,�����־,����ID,��ҳID,Ӥ����,
    'ҽ����Ŀ����,���մ���ID,�շ����,�շ�ϸĿID,�շ�����,
    '���㵥λ,��������,���,����,����,Decode(A.����,0,0,Round(A.���/A.����,4)) as �۸�,
    'A.���,A.ҽ��,A.����ʱ�� as �Ǽ�ʱ��,�Ƿ��ϴ� ,�Ƿ���,������Ŀ��,ժҪ
    
    
    Dim curTotal As Double, cur�����ʻ� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    
    Dim strInfor As String  '�������ķ��ش�
    Dim dbl���� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double
    
    '2005-08-02����������
    Dim dblҩ���Է� As Double
    
    Dim dbl���� As Double, dbl���Ʒ� As Double, dbl���� As Double, dbl����Է� As Double
    Dim dbl�������Ʒ� As Double, dbl���������Է� As Double
    Dim dbl�������Էѷ��� As Double, dbl�Ǳ��շ��� As Double, dblͳ����� As Double
    Dim dblѪ�� As Double, dblѪ���Է� As Double
    Dim dbl������ As Double, dbl�𸶱�׼ As Double
    Dim lng����ID As Long
    
    Dim str��ϱ��� As String, strҽʦ���� As String, str����Ա���� As String
    Dim str������� As String, str���������ʶ As String
    Dim strTmp As String, strҽ�� As String, str��ϸ As String      '��ϸ��
    Dim str���ұ��� As String, str��Ŀͳ�Ʒ��� As String, str��Ŀ���� As String, dbl��Ŀ���� As Double
    '----------------��rsExse�в��ֲ���ȷ���ֶ����¸�ֵ
    Dim int����id As Integer
    Dim int������Ŀ�� As Integer
    '--------------------------------------------------
    
    Dim intMouse As Integer
      
    intMouse = Screen.MousePointer
    int����id = 0
    int������Ŀ�� = 0
    
    '���������ǰ����֤���
    Screen.MousePointer = 1
    If ��ݱ�ʶ_����(0, lng����ID, intinsure) = "" Then
        Screen.MousePointer = intMouse
        ��������������_���� = False
        MsgBox "������������֤ʧ��,���ܽ��н���"
        Exit Function
    End If
    Screen.MousePointer = intMouse

    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ�
    
    '������֧��������ڱ���,�Ա���㱣�Ѽ��Է�
    gstrSQL = "Select * From ����֧������ "
    zlDatabase.OpenRecordset rs����, gstrSQL, "����֧������"
    
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lng����ID As Long
    With rs��ϸ
        'ȷ������
        If Not .EOF Then
            lng����ID = Nvl(!����ID, 0)
            gstrSQL = "  select ����id from �����ʻ� where ����id=" & lng����ID & "  and ����=" & intinsure & "  and ҽ����='" & g�������_����.���˱�� & "'"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
            If Not rsTemp.EOF Then
                lng����ID = Nvl(rsTemp!����ID, 0)
            Else
                lng����ID = 0
            End If
          '����׼��Ŀ
            gstrSQL = "Select * from ������׼��Ŀ  where ����ID=  " & lng����ID
            zlDatabase.OpenRecordset rs��׼��Ŀ, gstrSQL, "��ȡ������Ŀ����"
            
        End If
        
        'ȡ�����η������õĽ��ϼ�
        Do While Not .EOF
            If lng����ID <> 0 Then
                    '��һ��,ȷ��������շ�ϸĿ
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=0 And ����=1 and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
                    If rs��׼��Ŀ.EOF Then
                        gstrSQL = "Select ����,���� from �շ�ϸĿ where id=" & Nvl(!�շ�ϸĿID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�շ�ϸĿ"
                        ShowMsgbox "�շ�ϸĿΪ��" & Nvl(rsTemp!����) & "������Ŀ���ǲ��������趨����Ŀ."
                        Exit Function
                    End If
                    
                    '�ڶ���,ȷ������ı��մ���
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=1 And ����=1 and  �շ�ϸĿid=" & Nvl(!����֧������ID, 0)
                    If rs��׼��Ŀ.EOF Then
                        ShowMsgbox "�ڽ����д����˽�������ı���֧������,���ܼ�����"
                        Exit Function
                    End If
                    '������,'ȷ����ֹ���շ�ϸĿ
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=0 And ����=2 and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
                    If Not rs��׼��Ŀ.EOF Then
                        gstrSQL = "Select ����,���� from �շ�ϸĿ where id=" & Nvl(!�շ�ϸĿID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�շ�ϸĿ"
                        ShowMsgbox "�շ�ϸĿΪ��" & Nvl(rsTemp!����) & "������Ŀ�Ǳ���ֹʹ�õ���Ŀ." & vbCrLf & "���ܼ���!"
                        Exit Function
                    End If
                    '���Ĳ�,'ȷ����ֹ�Ĵ���
                    rs��׼��Ŀ.Filter = 0
                    rs��׼��Ŀ.Filter = "����=1 And ����=2 and �շ�ϸĿid=" & Nvl(!����֧������ID, 0)
                    If Not rs��׼��Ŀ.EOF Then
                        ShowMsgbox "�ڽ����д����˽�ֹʹ�õı���֧������,���ܼ�����"
                    End If
            End If
        
            '���ж��Ƿ�������ҽ����Ӧ��Ŀ����
            gstrSQL = " Select ��Ŀ����,��Ŀ����,����id,�Ƿ�ҽ�� From ����֧����Ŀ" & _
                      " Where ����=[1] And �շ�ϸĿID=[2]"
                      
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ������˶�Ӧ��ҽ����Ŀ", intinsure, CLng(!�շ�ϸĿID))
            
            If rsTemp.EOF = True Then
                MsgBox "����Ŀδ����ҽ����Ŀ�����ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            Else
                '�������������Ŀ�϶Է��ü�¼û����д���մ���id�ͱ�����Ŀ����ֶ�,
                '�����ȡ�������������ֶβ�����ȷ��ӳ��ʵֵ,��Ҫ�ڱ��ε�ѡ����������¸�ֵ
                int����id = Nvl(rsTemp!����id, 0)
                int������Ŀ�� = Nvl(rsTemp!�Ƿ�ҽ��, 0)
            End If
            
            If strҽ�� = "" Then
                strҽ�� = Nvl(!ҽ��)
            End If
            
            str��Ŀ���� = Nvl(rsTemp!��Ŀ����)
            dbl��Ŀ���� = Val(Nvl(rsTemp!��Ŀ����)) / 100
            
            lng����ID = Nvl(!����ID, 0)
            gstrSQL = "" & _
                " Select b.������,b.����ֵ from �շ���� a,���ղ��� b " & _
                " Where a.���=b.������ and b.����=" & intinsure & _
                "        and a.����='" & Nvl(!�շ����) & "'"
            
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "���Ѽ���"
            
            If rsTemp.EOF Then
                strTmp = ""
            Else
                strTmp = Nvl(rsTemp!����ֵ)
            End If
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '���㱣��
                rs����.Find "id=" & int����id, , adSearchForward, 1
                If Not rs����.EOF Then
                    dblͳ����� = Nvl(rs����!ͳ��ȶ�, 0) / 100
                Else
                    dblͳ����� = 1
                End If
                '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,Q��ҵ����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
                If intinsure <> TYPE_���������� And g�������_����.ְ����ҽ��� = "L" _
                    And g�������_����.�α����3 = "0" And Nvl(int������Ŀ��, 0) = 1 Then  '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '  ������  ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�����ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dblͳ����� = 1
                End If
                
                If intinsure = TYPE_������ And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dblͳ����� = dbl��Ŀ����
                End If
                                                
                '�ܺ�ȫ���� 2003-12-17
                '����������Ŀ��ֻҪ�Ǳ�ʶΪ�����Ρ��ģ���Ӧ�����������
                'If NVL(!�շ����) = "����" And str��Ŀ���� = "����" Then
                If str��Ŀ���� = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If str��Ŀ���� = "���" Then
                    strTmp = "����"
                End If
                '����۳��ԷѲ��ֵķ���.��Ϊֻ�д�λ�����Ƕ����,���������ﲻ�ᷢ����λ����
                'һ�����﷢����λ����,������ͳ�����=0���м���
                If Not rsTemp.EOF Then
                    If dblͳ����� <> 0 Then
                        Select Case strTmp
                            Case "����"
                                
                                dbl���� = dbl���� + Round(Nvl(!���, 0) * dblͳ�����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                               
                            Case "��ҩ��"
                                
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���, 0) * dblͳ�����, 5)
                                
                                '2005-08-02����������
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                                End If
                                
                            Case "��ҩ��"
                                
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���, 0) * dblͳ�����, 5)
                                '2005-08-02����������
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                                End If
                                
                            Case "��ҩ��"
                                
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���, 0) * dblͳ�����, 5)
                                '2005-08-02����������
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                                End If
                                
                            Case "����"
                                
                                dbl���� = dbl���� + Round(Nvl(!���, 0) * dblͳ�����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                                
                            Case "���Ʒ�"
                                
                                dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!���, 0) * dblͳ�����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)

                                '�ܺ�ȫ���� 2003-12-17
                                '������ҽ�������������޷���Ӧ��Ŀ�����������ȡ�õģ�
                            Case "����"
                                '�����кͿ������Դ����ô���ͬ,
                                '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                          
                                dbl���� = dbl���� + Round(Nvl(!��� * dblͳ�����, 0), 5)
                                
                                dbl����Է� = dbl����Է� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                            Case "Ѫ��"
                                '�����кͿ������Դ����ô���ͬ,
                                '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                          
                                dblѪ�� = dblѪ�� + Round(Nvl(!ʵ�ս�� * dblͳ�����, 0), 5)
                                dblѪ���Է� = dblѪ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dblͳ�����), 5)
                            Case "�������Ʒ�"
                                '2004/9/11��ǰ:�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                                '2004/9/11�Ժ�:�������뿪�������㷽ʽһ�£�����ͳ�ﲿ�ֵĽ��
                                'If intinsure = TYPE_������ Then
                                '   dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!ʵ�ս��, 0), 5)
                                'Else
                                    dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!ʵ�ս��, 0) * dblͳ�����, 5)
                                'End If
                                
                                '�������Ʒ��Էѵļ��㷽ʽ��ͬ,
                                '�����нӿ��ڴ�����ܽ��ʱֻ���������Ʒѽ��л���,���������ԷѲ��ֲ��ټ���
                                dbl���������Է� = dbl���������Է� + Round(Nvl(!���, 0) * (1 - dblͳ�����), 5)
                                
                        End Select
                    Else
                
                        'ȫ���Ǳ���Ϊ0����Ŀ(��������),�ֱ�Դ������ǿ����������жϷ��ڲ�ͬ���ֶ�
                        If intinsure = TYPE_������ Then
                            '�����з���dbl�Ǳ��շ���
                            dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(!���, 5)
                        Else
                            '����������dbl������
                            dbl������ = dbl������ + Round(!���, 5)
                        End If
                    
                    End If

                End If
            End If
            curTotal = curTotal + Round(Nvl(!���, 0), 5)
            .MoveNext
        Loop
    End With
  
    
    If strҽ�� <> "" Then
        gstrSQL = "Select ��� From ��Ա��  where ����='" & strҽ�� & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽ�����"
        If Not rsTemp.EOF Then
            strҽ�� = Nvl(rsTemp!���)
            If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                strҽ�� = Substr(strҽ��, 1, 6)
            End If
        Else
            strҽ�� = ""
        End If
    End If
    
    '���������Ϊ��
    dbl�𸶱�׼ = g�������_����.����
    If ����ҽ������(intinsure, 0, lng����ID, 0, 0, 0, False, True, dbl�𸶱�׼, dbl����, dbl��ҩ��, dbl��ҩ��, dbl��ҩ��, _
        dbl����, dbl���Ʒ�, dblѪ��, dblѪ���Է�, dbl����, dbl����Է�, dbl�������Ʒ�, dbl���������Է�, dbl�������Էѷ���, _
        dbl�Ǳ��շ���, dbl������, dblҩ���Է�, curTotal, strҽ��, str���㷽ʽ) = False Then
        Exit Function
    End If
    
    ��������������_���� = True
End Function

Private Function ������ʽ��㼰����_����(ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal ԭ����id As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean

  '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID��
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Dim curTotal As Double
    Dim rsTemp As New ADODB.Recordset, rs��ϸ As New ADODB.Recordset
    Dim strInfor As String  '�������ķ��ش�
    Dim dbl���� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double
    
    '2005-08-02����������
    Dim dblҩ���Է� As Double
    
    Dim dbl���� As Double, dbl���Ʒ� As Double
    Dim dbl���� As Double, dbl����Է� As Double, dbl�������Ʒ� As Double, dbl���������Է� As Double, dbl�������Էѷ��� As Double
    Dim dbl�Ǳ��շ��� As Double, dbl������ As Double    '��Դ�����������
    Dim dblѪ�� As Double, dblѪ���Է� As Double
    Dim dbl�𸶱�׼ As Double, dbl���� As Double
    Dim strҽ�� As String, str��ϸ As String      '��ϸ��
    Dim str���ұ��� As String, str��Ŀ���� As String, str��Ŀͳ�Ʒ��� As String, strTmp As String
    Dim intҵ�� As Integer, lng����ID As Long
    Dim strNO As String, lng��¼���� As Long
    
    Dim lng������� As Long, dbl�����ʻ���� As Double
    Dim dblͳ��֧���ۼ� As Double, dbl�����ʻ�֧�� As Double
    Dim dbl�����ʻ�֧�� As Double
    Dim dbl����ͳ��֧�� As Double
    Dim dbl����ͳ���Ը� As Double
    Dim dbl����ͳ��֧�� As Double
    Dim dbl����ͳ���Ը� As Double
    Dim dbl��������֧�� As Double
    Dim dbl�ǲ�������֧�� As Double
    Dim dbl���շ�Χ���Ը� As Double
    
    Dim dbl����ǰ�����ʻ����  As Double
    Dim dbl����ǰ�����˻����  As Double
    Dim dbl����ǰͳ���ۼ�  As Double
    Dim lngTmp As Long
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim lng����ID As Long
    
    intҵ�� = IIf(bln����, 1, 0)
     ������ʽ��㼰����_���� = False
   
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ���������ڽ��㣬����ʹ��Ȧ��ӿڣ����������Ǯ���������ڣ������ӿ��ڽ��
    '���������Ҫͨ��������������ȡ����Ȧ�����ǽӿڷ��أ���Ҫ�޸�
    
    On Error GoTo errHand
    
    '���¶���
    If ��ȡ�������_����(IIf(intinsure = TYPE_����������, 2, 1), intinsure) = False Then
        Exit Function
    End If
    
    If bln���� Then
        lng����ID = ԭ����id
        '��֤�Ƿ�Ϊ�ò��˵�IC��
        gstrSQL = "Select * From  �����ʻ� where ����id=" & lng����ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵�ҽ����"
        If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "�ò����ڱ����ʻ����޼�¼!"
            Exit Function
        End If
        
        If g�������_����.IC���� <> Nvl(rsTemp!����) Then
            Err.Raise 9000, gstrSysName, "�ò��˵�IC���������,�����ǲ����������˵�IC��!"
            Exit Function
        End If
        'ȷ���������,ת�ﵥ��,��ϱ���,�������
        ' ֧��˳���_IN(�������;ת�ﵥ��;��ϱ���),��ע(�������_IN)
        gstrSQL = "Select ֧��˳���,��ע from ���ս����¼  where ��¼ID=" & lng����ID
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�������"
        If rsTemp.RecordCount = 0 Then
            Err.Raise 9000, gstrSysName, "�ڽ����¼���޽����¼!"
            Exit Function
        End If
        Dim strArr
        strArr = Split(Nvl(rsTemp!֧��˳���), ";")
        
        '�������;ת�ﵥ��;��ϱ���
        '1-��ͨ����("1", "A"),2-��������("3", "7")
        '3-�����("5", "B"),4-������������("S", "T")
        If UBound(strArr) >= 2 Then
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g�������_����.ת�ﵥ�� = strArr(1)
            g�������_����.��ϱ��� = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g�������_����.ת�ﵥ�� = strArr(1)
        Else
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
        End If
        g�������_����.������� = Nvl(rsTemp!��ע)
        
        
        'ȷ���˷Ѽ�¼
        '�˷�
        gstrSQL = "select id from ������ü�¼ where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����˷�", lng����ID)
        If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "�����ڲ��˷��ó�����¼!"
            Exit Function
        End If
         
    End If
    '�򿪱��ν�����ϸ��¼ '--���ұ���Ӧ���Ǳ�ʶ����+����
    gstrSQL = " " & _
        "  Select Rownum ��ʶ��,A.ID,A.����ID,A.�շ�ϸĿid,A.NO,A.���,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ��,A.������ as ҽ��,H.��� as ҽ�����, " & _
        "      A.����*A.���� as ����,A.���㵥λ,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ�� as ʵ�ս��,F.����ֵ,G.id as ����id,G.ͳ��ȶ�, " & _
        "      a.ҽ�����, A.�շ����,B.���� as ��Ŀ����,B.���� as ��Ŀ����,Nvl(J.��ʶ��,Nvl(B.��ʶ����||B.��ʶ����,B.����)) as ���ұ���, " & _
        "      D.��Ŀ���� ҽ������,D.��Ŀ���� as ҽ������,J.���� as ����,D.�Ƿ�ҽ��,C.���� ��������,E.���� �ܵ�����, " & _
        "      L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.��λ����,L.˳���,L.����֤��,L.�ʻ����,L.��ǰ״̬,L.����ID,L.��ְ,L.�����,L.�Ҷȼ�,L.����ʱ�� " & _
        "  From (Select * From ������ü�¼ Where ��¼״̬<>0 and ����ID=" & IIf(bln����, lng����ID, lng����ID) & " and  Nvl(���ӱ�־,0)<>9 ) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,���ű� E,  " & _
        "       (Select U.*,K.����ֵ From �շ���� U,���ղ��� K where U.���=K.������ and K.����=" & intinsure & "  ) F, " & _
        "       (Select distinct Q.ҩƷid,Q.��ʶ��,T.���� From ҩƷĿ¼ Q,ҩƷ��Ϣ R,ҩƷ���� T  Where  Q.ҩ��id=R.ҩ��id and R.����=T.���� ) J, " & _
        "       ����֧������ G,��Ա�� H,�����ʻ� L" & _
        "  Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID(+) and A.����id=L.����id and L.����=" & intinsure & " and A.�շ����=F.����(+)  and d.����id=G.id and a.�շ�ϸĿid=J.ҩƷid(+) " & _
        "        And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����= " & intinsure & " and a.������=H.����(+) " & _
        "  Order by A.ID"
        
    '�ϴ�������ϸ��¼
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡ���ν��ʷ�����ϸ"
    
    With rs��ϸ
        If Not .EOF Then
            lng����ID = Nvl(!����ID, 0)
            strҽ�� = Nvl(!ҽ�����)
            If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                strҽ�� = Substr(strҽ��, 1, 6)
            End If
            lng����ID = Nvl(!����ID, 0)
            '����׼��Ŀ
            gstrSQL = "Select * from ������׼��Ŀ  where ����ID=  " & lng����ID
            zlDatabase.OpenRecordset rs��׼��Ŀ, gstrSQL, "��ȡ������Ŀ����"
        End If
        
        Do While Not .EOF
            If lng����ID <> 0 And bln���� = False Then
                '��һ��,ȷ��������շ�ϸĿ
                rs��׼��Ŀ.Filter = 0
                rs��׼��Ŀ.Filter = "����=0 And ����=1 and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
                If rs��׼��Ŀ.EOF Then
                    Err.Raise 9000, gstrSysName, "�շ�ϸĿΪ��" & Nvl(!��Ŀ����) & "������Ŀ���ǲ��������趨����Ŀ."
                    Exit Function
                End If
                '�ڶ���,ȷ������ı��մ���
                rs��׼��Ŀ.Filter = 0
                rs��׼��Ŀ.Filter = "����=1 And ����=1 and  �շ�ϸĿid=" & Nvl(!����id, 0)
                If rs��׼��Ŀ.EOF Then
                    Err.Raise 9000, gstrSysName, "�ڽ����д����˽�������ı���֧������,���ܼ�����"
                    Exit Function
                End If
                '������,'ȷ����ֹ���շ�ϸĿ
                rs��׼��Ŀ.Filter = 0
                rs��׼��Ŀ.Filter = "����=0 And ����=2 and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
                If Not rs��׼��Ŀ.EOF Then
                    Err.Raise 9000, gstrSysName, "�շ�ϸĿΪ��" & Nvl(!��Ŀ����) & "������Ŀ�Ǳ���ֹʹ�õ���Ŀ." & vbCrLf & "���ܼ���!"
                    Exit Function
                End If
                '���Ĳ�,'ȷ����ֹ�Ĵ���
                rs��׼��Ŀ.Filter = 0
                rs��׼��Ŀ.Filter = "����=1 And ����=2 and �շ�ϸĿid=" & Nvl(!����id, 0)
                If Not rs��׼��Ŀ.EOF Then
                    Err.Raise 9000, gstrSysName, "�ڽ����д����˽�ֹʹ�õı���֧������,���ܼ�����"
                End If
            End If
            strTmp = Nvl(!����ֵ)
            lng����ID = Nvl(!����ID, 0)
            'ȷ���������
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str��Ŀͳ�Ʒ��� = ""
                Else
                    str��Ŀͳ�Ʒ��� = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                '����
                '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
                    
                If Nvl(!����, 0) <> TYPE_���������� And Val(Nvl(!��λ����, "99")) = 0 And Nvl(!��ְ, 0) = 3 And Nvl(!�Ƿ�ҽ��, 0) = 1 Then   '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '������    ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�������ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dbl���� = 1
                Else
                    dbl���� = Nvl(!ͳ��ȶ�, 0) / 100
                End If
                
                If Nvl(!����, 0) = TYPE_������ And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dbl���� = Val(Nvl(!ҽ������)) / 100
                End If
                
                If Nvl(!ҽ������) = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If Nvl(!ҽ������) = "���" Then
                    strTmp = "����"
                End If

                If dbl���� <> 0 Then
                    
                    Select Case strTmp
                            Case "����"
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                            
                            Case "��ҩ��"
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                
                                '2005-08-02����������
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                End If
                                
                            Case "��ҩ��"
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                
                                '2005-08-02����������
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                End If
                                
                            Case "��ҩ��"
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                
                                '2005-08-02����������
                                If intinsure = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                End If
                                
                            Case "����"
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                
                            Case "���Ʒ�"
                                dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                '�ܺ�ȫ���� 2003-12-17
                                '������ҽ�������������޷���Ӧ��Ŀ�����������ȡ�õģ�
                            Case "����"
                                '�����кͿ������Դ����ô���ͬ,
                                '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                          
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս�� * dbl����, 0), 5)
                                
                                dbl����Է� = dbl����Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                            Case "Ѫ��"
                                '�����кͿ������Դ����ô���ͬ,
                                '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                          
                                dblѪ�� = dblѪ�� + Round(Nvl(!ʵ�ս�� * dbl����, 0), 5)
                                dblѪ���Է� = dblѪ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                            Case "�������Ʒ�"
                                '2004/9/11��ǰ:�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                                '2004/9/11�Ժ�:�������뿪�������㷽ʽһ�£�����ͳ�ﲿ�ֵĽ��
                                'If intinsure = TYPE_������ Then
                                '   dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!ʵ�ս��, 0), 5)
                                'Else
                                    dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                'End If
                                
                                '�������Ʒ��Էѵļ��㷽ʽ��ͬ,
                                '�����нӿ��ڴ�����ܽ��ʱֻ���������Ʒѽ��л���,���������ԷѲ��ֲ��ټ���
                                dbl���������Է� = dbl���������Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                
                        End Select
                    Else
                
                        'ȫ���Ǳ���Ϊ0����Ŀ(��������),�ֱ�Դ������ǿ����������жϷ��ڲ�ͬ���ֶ�
                        If intinsure = TYPE_������ Then
                            '�����з���dbl�Ǳ��շ���
                            dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(!ʵ�ս��, 5)
                        Else
                            '����������dbl������
                            dbl������ = dbl������ + Round(!ʵ�ս��, 5)
                        End If
                    
                    End If
            Else
                dbl���� = 1
                str��Ŀͳ�Ʒ��� = ""
            End If

            '�ϴ���ϸ��¼,ʵʱҽ����ϸ����
            '����������ϸ�ϴ�
            If gbln������ϸʱʵ�ϴ� Then
                
                    If Nvl(!����, 0) = TYPE_���������� Then '������
                        str��ϸ = Lpad(gstrҽԺ����_����, 6)     'ҽԺ����    CHAR    1   6       Ժ����д
                        str��ϸ = str��ϸ & Lpad(Nvl(!ҽ����), 10)  '���ձ��    CHAR    7   10      Ժ����д
                    Else
                        str��ϸ = Lpad(gstrҽԺ����_����, 4)     'ҽԺ����    CHAR    1   4       Ժ��
                        str��ϸ = str��ϸ & Lpad(Nvl(!ҽ����), 8)   '���˱��    CHAR    5   8       Ժ��
                    End If
                
                    str��ϸ = str��ϸ & Space(10)   '��־��  CHAR    13  10  ������ϸ�Կո�λ,סԺ��סԺ��  Ժ��
                    str��ϸ = str��ϸ & Lpad(g�������_����.�������, 4)     '�������
                    
                    'Modified By ���� 2004-07-29 ԭ�򣺴���NO��
                    str��ϸ = str��ϸ & Lpad(Mid(Nvl(!NO, "00000000"), 2, 7), 10)      '������  NUM 27  10      Ժ��
                    str��ϸ = str��ϸ & Lpad(!���, 10)       '������Ŀ���    NUM 37  10  ��Ӧ�����ŵļǼ���Ŀ���    Ժ��
                    
                    '������Ϊ���ݺ�  CHAR    41  10  ҽ���ţ�    Ժ����д
                    str��ϸ = str��ϸ & Space(10)       'ҽ����  CHAR    47  10  ������Ӧҽ����ҽ����¼�ţ�������ϸ��û��ҽ����ҽԺ�Կո�λ    Ժ��
                    str��ϸ = str��ϸ & Get�������(intҵ��, Nvl(!�Ҷȼ�, 0))         '�������    CHAR    57  1   ȡֵ���"�������"˵��  Ժ��
                    str��ϸ = str��ϸ & Rpad(Format(!�Ǽ�ʱ��, "yyyymmddHHmmss"), 16)      '��������ʱ�䣨Ͷҩʱ�䣩    DATETIME    58  16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ    Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!���ұ���), 20)      '��Ŀ����    CHAR    74  20  �Ƽ���Ŀ����    Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!��Ŀ����), 20)      '��Ŀ����    CHAR    94  20      Ժ��
        
                    If !�Ƿ�ҽ�� = 1 Then
                        str��ϸ = str��ϸ & Lpad(1 - dbl����, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                    Else
                        str��ϸ = str��ϸ & Lpad(1, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(str��Ŀͳ�Ʒ���, 1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                    
                    '2005-08-02����������
                    If Nvl(!����, 0) = TYPE_���������� Then '������
                        str��ϸ = str��ϸ & Lpad(Nvl(!����), 10)  '����    NUM 121 10   �巽����Ϊ��ֵ  Ժ��
                        str��ϸ = str��ϸ & Lpad(Abs(Nvl(!ʵ�ʼ۸�)), 10) '����    NUM 127 10   ��������ָ�ֵ  Ժ��
                    Else
                        str��ϸ = str��ϸ & Lpad(Nvl(!����), 6)  '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                        str��ϸ = str��ϸ & Lpad(Abs(Nvl(!ʵ�ʼ۸�)), 8) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(Nvl(!���㵥λ), 4) '��λ    CHAR    135 4       Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!����), 20)      '����    CHAR    139 20  �����Ƭ����    Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!ҽ��), 8)      'ҽʦ����    CHAR    159 8       Ժ��
                    str��ϸ = str��ϸ & Lpad(g�������_����.��ϱ���, 16)      '��ϱ���    CHAR    167 16      Ժ��
                    str��ϸ = str��ϸ & Lpad(Substr(g�������_����.�������, 1, 28), 30)   '�������    CHAR    183 30      Ժ��
                    str��ϸ = str��ϸ & Space(16)     '����ʱ��    DATETIME    213 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ��Ժ�˿ո�λ  ����
                
                '�ϴ���ϸ
                '1003    7   230 ʵʱҽ����ϸ�����ύ
                ������ʽ��㼰����_���� = ҵ������_����(IIf(Nvl(!����, 0) = TYPE_����������, 2, 1), 1003, str��ϸ, intinsure)
                If ������ʽ��㼰����_���� = False Then
                    Err.Raise 9000, gstrSysName, "������ʽ���ʱҽ����ϸ�����ύʧ��,���ܼ���!"
                    Exit Function
                End If
                '�ϴ�ҽ����ϸ
                If Nvl(!ҽ�����, 0) <> 0 Then
                    If ҽ����ϸ�����ύ(!ҽ�����, "", str��Ŀͳ�Ʒ���, intinsure) = False Then
                        Err.Raise 9000, gstrSysName, "ҽ����ϸ�����ύʧ��,���ܼ���!"
                        Exit Function
                    End If
                End If
                'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            End If
            '�����ܶ�,����
            curTotal = curTotal + Round(Nvl(!ʵ�ս��, 0), 5)
            .MoveNext
        Loop
    End With
    ������ʽ��㼰����_���� = False
    
    '��������
    dbl�𸶱�׼ = g�������_����.����
    
    If ����ҽ������(intinsure, 2, lng����ID, 0, IIf(bln����, lng����ID, lng����ID), ԭ����id, bln����, False, 0, _
        dbl����, dbl��ҩ��, dbl��ҩ��, dbl��ҩ��, dbl����, dbl���Ʒ�, dblѪ��, dblѪ���Է�, dbl����, dbl����Է�, _
        dbl�������Ʒ�, dbl���������Է�, dbl�������Էѷ���, dbl�Ǳ��շ���, dbl������, dblҩ���Է�, curTotal, strҽ��, strInfor) = False Then
        Exit Function
    End If
    ������ʽ��㼰����_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function ����ҽ������(ByVal intinsure As Integer, ByVal bytType As Byte, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lngԭ����ID As Long, ByVal bln���� As Boolean, ByVal bln������� As Boolean, dbl�𸶱�׼ As Double, _
        dbl���� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double, _
        dbl���� As Double, dbl���Ʒ� As Double, dblѪ�� As Double, dblѪ���Է� As Double, dbl���� As Double, dbl����Է�, _
        dbl�������Ʒ� As Double, dbl���������Է� As Double, dbl�������Էѷ��� As Double, dbl�Ǳ��շ��� As Double, _
        dbl������ As Double, dblҩ���Է� As Double, curTotal As Double, strҽ�� As String, str���㷽ʽ As String, Optional strסԺ�� As String = "", Optional str��Ժ���� As String = "") As Boolean
        
        '����:���н���
        '����:bytType-0����,1סԺ,2�������
        '   lng����id-����idֵ
        '   bln�������=�Ƿ��������
        '   bln���� =����
        '   dbl��ͷ�Ĵ���ط���
        ' ����:
        '   str���㷽ʽ
        '����:�ɹ�,����true,���򷵻�False
        
        
        Dim rsTemp As New ADODB.Recordset
        Dim dbl�����ʻ���� As Double, dblͳ��֧���ۼ� As Double, dbl�����ʻ�֧�� As Double
        Dim dbl�����ʻ�֧�� As Double, dbl����ͳ��֧�� As Double, dbl����ͳ���Ը� As Double
        Dim dbl����ͳ��֧�� As Double, dbl����ͳ���Ը� As Double, dbl��������֧�� As Double
        Dim dbl�ǲ�������֧�� As Double, dbl���շ�Χ���Ը� As Double
        Dim dbl����Ա�𸶱�׼���� As Double, dbl����Ա�������� As Double, dbl����Ա�ǻ������� As Double
        Dim dbl��ҵ���ղ��� As Double, dbl�Ǳ����Ը� As Double
        Dim intҵ�� As Integer
        Dim dbl����ǰ�����ʻ����  As Double
        Dim dbl����ǰ�����˻����  As Double
        Dim dbl����ǰͳ���ۼ�  As Double
        Dim dbl���� As Double
        
        '2005-08-02 ����������
        Dim dbl����ǰ����ͳ���ۼ� As Double
        Dim dbl����󼲲�ͳ���ۼ� As Double
        
        Dim strInputString As String
        
        intҵ�� = IIf(bln����, 1, 0)
        
        '20040727:ͬ����������ɶԲ�����,�����Ȼ��ܺ�����������
        dbl���� = Round(dbl����, 2)
        dbl��ҩ�� = Round(dbl��ҩ��, 2)
        dbl��ҩ�� = Round(dbl��ҩ��, 2)
        dbl��ҩ�� = Round(dbl��ҩ��, 2)
        dbl���� = Round(dbl����, 2)
        dbl���Ʒ� = Round(dbl���Ʒ�, 2)
        dbl���� = Round(dbl����, 2)
        dbl����Է� = Round(dbl����Է�, 2)
    
        '2004/9/11 ���˺�����Ѫ��
        dblѪ�� = Round(dblѪ��, 2)
        dblѪ���Է� = Round(dblѪ���Է�, 2)
        
        dbl�������Ʒ� = Round(dbl�������Ʒ�, 2)
        dbl���������Է� = Round(dbl���������Է�, 2)
        dbl�������Էѷ��� = Round(dbl�������Էѷ���, 2)
        dbl�Ǳ��շ��� = Round(dbl�Ǳ��շ���, 2)
        dbl������ = Round(dbl������, 2)
        
        '����������
        dblҩ���Է� = Round(dblҩ���Է�, 2)
        curTotal = Round(curTotal, 2)
    
            
        Err = 0
        On Error GoTo errHand:
        
        ����ҽ������ = False
        If bln���� Then
            gstrSQL = "" & _
                "   Select *  " & _
                "   From ���ս����¼ " & _
                "   Where ��¼id=" & lngԭ����ID
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�����շ�ʱ���ص�����"
            If rsTemp.RecordCount = 0 Then
                ShowMsgbox "�������ϴ��շѵĽ����¼!"
                Exit Function
            End If
            dbl�����ʻ���� = Round(Nvl(rsTemp!�ʻ��ۼ�����, 0), 2)
            dblͳ��֧���ۼ� = Round(Nvl(rsTemp!�ʻ��ۼ�֧��, 0), 2)
            dbl��������֧�� = Round(Nvl(rsTemp!�ۼƽ���ͳ��, 0), 2)
            dbl�����ʻ�֧�� = Round(Nvl(rsTemp!�ۼ�ͳ�ﱨ��, 0), 2)
            dbl�𸶱�׼ = Round(Nvl(rsTemp!����, 0), 2)
            dbl���շ�Χ���Ը� = Round(Nvl(rsTemp!�ⶥ��, 0), 2)
            dbl����ͳ��֧�� = Round(Nvl(rsTemp!ȫ�Ը����, 0), 2)
            dbl����ͳ���Ը� = Round(Nvl(rsTemp!�����Ը����, 0), 2)
            dbl����ͳ��֧�� = Round(Nvl(rsTemp!����ͳ����, 0), 2)
            dbl����ͳ���Ը� = Round(Nvl(rsTemp!ͳ�ﱨ�����, 0), 2)
            dbl�ǲ�������֧�� = Round(Nvl(rsTemp!���Ը����, 0), 2)
            dbl�����ʻ�֧�� = Round(Nvl(rsTemp!�����ʻ�֧��, 0), 2)
            dbl����ǰ�����ʻ���� = Round(Nvl(rsTemp!����ǰ�����ʻ����, 0), 2)
            dbl����ǰ�����˻���� = Round(Nvl(rsTemp!����ǰ�����˻����, 0), 2)
            dbl����ǰͳ���ۼ� = Round(Nvl(rsTemp!����ǰͳ���ۼ�, 0), 2)
            
            dbl����Ա�𸶱�׼���� = Round(Nvl(rsTemp!����Ա�𸶱�׼����, 0), 2)
            dbl����Ա�������� = Round(Nvl(rsTemp!����Ա��������, 0), 2)
            dbl����Ա�ǻ������� = Round(Nvl(rsTemp!����Ա�ǻ�������, 0), 2)
            dbl��ҵ���ղ��� = Round(Nvl(rsTemp!��ҵ���ղ���, 0), 2)
            dbl�Ǳ����Ը� = Round(Nvl(rsTemp!�Ǳ����Ը�, 0), 2)
            '���жϳ����������,�������,�����ϴν���ķ��ö�
            
            dbl���� = Round(Nvl(rsTemp!����, 0), 2)
            dbl��ҩ�� = Round(Nvl(rsTemp!��ҩ��, 0), 2)
            dbl��ҩ�� = Round(Nvl(rsTemp!��ҩ��, 0), 2)
            dbl��ҩ�� = Round(Nvl(rsTemp!��ҩ��, 0), 2)
            dbl���� = Round(Nvl(rsTemp!����, 0), 2)
            dbl���Ʒ� = Round(Nvl(rsTemp!���Ʒ�, 0), 2)
            dbl���� = Round(Nvl(rsTemp!����, 0), 2)
            dbl����Է� = Round(Nvl(rsTemp!����Է�, 0), 2)
            
            '2004/9/11 ���˺�����Ѫ��
            dblѪ�� = Round(Nvl(rsTemp!Ѫ��, 0), 2)
            dblѪ���Է� = Round(Nvl(rsTemp!Ѫ���Է�, 0), 2)
            
            dbl�������Ʒ� = Round(Nvl(rsTemp!�������Ʒ�, 0), 2)
            dbl���������Է� = Round(Nvl(rsTemp!���������Է�, 0), 2)
            dbl�������Էѷ��� = Round(Nvl(rsTemp!�������Էѷ���, 0), 2)
            dbl�Ǳ��շ��� = Round(Nvl(rsTemp!�Ǳ��շ���, 0), 2)
            dbl������ = Round(Nvl(rsTemp!������, 0), 2)
            
            '2005-08-02 ����������
            dblҩ���Է� = Round(Nvl(rsTemp!ҩ���Է�, 0), 2)
            dbl����ǰ����ͳ���ۼ� = Round(Nvl(rsTemp!����ǰ����ͳ���ۼ�, 0), 2)
            dbl����󼲲�ͳ���ۼ� = Round(Nvl(rsTemp!����󼲲�ͳ���ۼ�, 0), 2)
            
            curTotal = Round(Nvl(rsTemp!�������ý��, 0), 2)
        End If
        
    With g�������_����
        If intinsure = TYPE_���������� Then    '������
            strInputString = Lpad(gstrҽԺ����_����, 6)       'ҽԺ����
        Else
            strInputString = Lpad(gstrҽԺ����_����, 4)       'ҽԺ����
        End If
        strInputString = strInputString & " "      '�������ʶ
        If intinsure = TYPE_���������� Then   '������
            strInputString = strInputString & Lpad(.���˱��, 10)         '���˱��
        Else
            strInputString = strInputString & Lpad(.���˱��, 8)      '���˱��
        End If
        strInputString = strInputString & Lpad(.IC����, 7)       'IC����
        If bln������� Then
            strInputString = strInputString & Lpad(.������� + 1, 4)     '�������
        Else
            .������� = .������� + 1
            strInputString = strInputString & Lpad(.�������, 4)       '�������
        End If
        
        strInputString = strInputString & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '����ʱ��
        If bytType = 1 Then
            'סԺ��סԺ��
            strInputString = strInputString & Lpad(strסԺ��, 10)  '��־��
                
        Else
            strInputString = strInputString & String(10, " ") '��־��
        End If
        
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl����, 2))), 10) '����
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10) '��ҩ��
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10) '��ҩ��
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl��ҩ��, 2))), 10)  '��ҩ��
        If intinsure = TYPE_������ Then
        Else
            '2005-08-02����������
            strInputString = strInputString & Lpad(Trim(CStr(Round(dblҩ���Է�, 2))), 10)  'ҩ���Է�
        End If
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl����, 2))), 10)  '����
        strInputString = strInputString & Lpad(Trim(CStr(Round(dbl���Ʒ�, 2))), 10)   '���Ʒ�
        
        If intinsure = TYPE_������ Then
            '2004/9/11 ���˺�����Ѫ��,ͬʱ�ı���˳��
            strInputString = strInputString & Lpad(Trim(CStr(Round(dblѪ��, 2))), 10)  'Ѫ��
            strInputString = strInputString & Lpad(Trim(CStr(Round(dblѪ���Է�, 2))), 10)   'Ѫ���Է�
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl����, 2))), 10)   '����
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl����Է�, 2))), 10)   '����Է�
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl�������Ʒ�, 2))), 10)   '�������Ʒ�
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl���������Է�, 2))), 10)    '�����Է�    NUM 145 10      Ժ����д
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl�������Էѷ���, 2))), 10)    '�������Էѷ���
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl�Ǳ��շ���, 2))), 10)    '�Ǳ��շ���
        Else
            '2005-08-02����������
            strInputString = strInputString & Lpad(Trim(CStr(Round(dblѪ��, 2))), 10)  'Ѫ��
            strInputString = strInputString & Lpad(Trim(CStr(Round(dblѪ���Է�, 2))), 10)   'Ѫ���Է�
            
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl����, 2))), 10)   '����
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl����Է�, 2))), 10)   '����Է�
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl�������Ʒ�, 2))), 10)   '�������Ʒ�
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl���������Է�, 2))), 10)    '�����Է�    NUM 145 10      Ժ����д
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl�������Էѷ���, 2))), 10)    '�������Էѷ���
            strInputString = strInputString & Lpad(Trim(CStr(Round(dbl������, 2))), 10)    '�������Է�  NUM 165 10  ��ҽ����ҩ�ԷѲ���  Ժ����д

        End If
        
        If bln���� Then
            strInputString = strInputString & Lpad(dbl�����ʻ����, 10)
            strInputString = strInputString & Lpad(dblͳ��֧���ۼ�, 10)
            If intinsure = TYPE_������ Then
            Else
                '2005-08-02����������
                strInputString = strInputString & Lpad(dbl����󼲲�ͳ���ۼ�, 10)
            End If
        Else
            strInputString = strInputString & String(10, " ") 'Lpad(dbl�����ʻ����, 10)
            strInputString = strInputString & String(10, " ") 'Lpad(dblͳ��֧���ۼ�, 10)
            If intinsure = TYPE_������ Then
            Else
                '2005-08-02����������
                strInputString = strInputString & String(10, " ") 'Lpad(dbl����󼲲�ͳ���ۼ�, 10)
            End If
        End If
        '����ǰ�����ʻ��������鿨���ؽ�����������������Ӧ����������ѯ�������ʻ������д��
        Dim dbl����ǰ���(1 To 3) As Double '1-����ǰ�����ʻ����,2-����ǰ�����˻����,3-����ǰͳ��֧���ۼ�
        
        If bytType = 0 And bln���� = False And .���㿪ʼ Then
            .��ǰ�����ʻ���� = Format(.���������ʻ����, "####0.00;-####0.00;0;0")
            .��ǰ�����ʻ���� = Format(.���������ʻ����, "####0.00;-####0.00;0;0")
            .��ǰͳ���ۼ� = Format(.ͳ���ۼ�, "####0.00;-####0.00;0;0")
        End If
        
        If bytType = 0 And bln���� = False Then
            dbl����ǰ���(1) = Format(.��ǰ�����ʻ����, "####0.00;-####0.00;0;0")
            dbl����ǰ���(2) = Format(.��ǰ�����ʻ����, "####0.00;-####0.00;0;0")
            dbl����ǰ���(3) = Format(.��ǰͳ���ۼ�, "####0.00;-####0.00;0;0")
        Else
            dbl����ǰ���(1) = Format(.���������ʻ����, "####0.00;-####0.00;0;0")
            dbl����ǰ���(2) = Format(.���������ʻ����, "####0.00;-####0.00;0;0")
            dbl����ǰ���(3) = Format(.ͳ���ۼ�, "####0.00;-####0.00;0;0")
        End If
        
        '����ǰ�����ʻ��������鿨���ؽ�����������������Ӧ����������ѯ�������ʻ������д��
        If bln���� Then
                strInputString = strInputString & Lpad(dbl����ǰ�����ʻ����, 10)   '����ǰ�����ʻ����
                strInputString = strInputString & Lpad(dbl����ǰ�����˻����, 10)    '����ǰ�����˻����(�����鿨���ؽ�������������������0)
                strInputString = strInputString & Lpad(dbl����ǰͳ���ۼ�, 10)     '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
        Else
            If intinsure <> TYPE_���������� And Get�������(0, .�������) = "S" Then
                '�� ����ǻ���ҽ�ƽ����ʾ: ���������ʻ���� ���������ʻ����
                '�� ��������������ʾ: �����ʻ����
                strInputString = strInputString & Lpad(.�����ʻ���ǰֵ, 10)   '����ǰ�����ʻ����
                strInputString = strInputString & Lpad("0", 10)   '����ǰ�����˻����(�����鿨���ؽ�������������������0)
                strInputString = strInputString & Lpad("0", 10)   '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
                dbl����ǰ���(1) = .�����ʻ���ǰֵ
                dbl����ǰ���(2) = 0
                dbl����ǰ���(3) = 0
            Else
                
                strInputString = strInputString & Lpad(dbl����ǰ���(1), 10)   '����ǰ�����ʻ����
                strInputString = strInputString & Lpad(Trim(CStr(dbl����ǰ���(2))), 10)    '����ǰ�����˻����(�����鿨���ؽ�������������������0)
                strInputString = strInputString & Lpad(Trim(CStr(dbl����ǰ���(3))), 10)    '����ǰͳ��֧���ۼ�:�����鿨���ؽ�������������������0
            End If
        End If
              
        If bln���� Then
            '���ϴ���Ӧ��ֵ
            If intinsure = TYPE_������ Then
            Else
                '2005-08-02����������
                strInputString = strInputString & Lpad(dbl����ǰ����ͳ���ۼ�, 10)
            End If
            strInputString = strInputString & Lpad(dbl�����ʻ�֧��, 10) '���ķ���:���λ��������ʻ�֧��(������������㣬��ʾ�����ʻ�֧��)
            strInputString = strInputString & Lpad(dbl�����ʻ�֧��, 10) '���ķ���:���β��������ʻ�֧��(������������㷵��0)
            strInputString = strInputString & Lpad(dbl����ͳ��֧��, 10) '���ķ���:���λ���ͳ��֧��
            strInputString = strInputString & Lpad(dbl����ͳ���Ը�, 10) '���ķ���:���λ���ͳ���Ը�
            strInputString = strInputString & Lpad(dbl����ͳ��֧��, 10)  '���ķ���:���β���ͳ��֧��
            strInputString = strInputString & Lpad(dbl����ͳ���Ը�, 10) '���ķ���:���β���ͳ���Ը�
            
            If intinsure = TYPE_������ Then
                '2004/9/11 ����ֵ�䶯��ԭ���Ĺ���Ա��������ҵ��������һ���ֶθ�Ϊ����Ա������Ŀ����ҵ����������¼���ֱ��ڵ�33��34��35��36���ֶ�������
                strInputString = strInputString & Lpad(dbl����Ա�𸶱�׼����, 10)   '33  ���ι���Ա�𸶱�׼����֧��  NUM 301 10      ����
                strInputString = strInputString & Lpad(dbl����Ա��������, 10)       '34  ���ι���Ա������������֧��  NUM 311 10      ����
                strInputString = strInputString & Lpad(dbl����Ա�ǻ�������, 10)     '35  ���ι���Ա�ǻ�����������֧��    NUM 321 10      ����
                strInputString = strInputString & Lpad(dbl��ҵ���ղ���, 10)         '36  ������ҵ���ղ���֧��    NUM 331 10      ����
                strInputString = strInputString & Lpad(dbl���շ�Χ���Ը�, 10)       '37  ���α������Ը�  NUM 341 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣����������Էѷ���+����Է�+Ѫ���Է�+�����Էѣ��ĸ����˻���ֺ�ķ���   ����
                strInputString = strInputString & Lpad(dbl�Ǳ����Ը�, 10)           '38  ���ηǱ����Ը�  NUM 351 10      ����
            Else
                '2005-08-02����������
                strInputString = strInputString & Lpad(dbl����Ա�𸶱�׼����, 10)   '36  ���ι���Ա�𸶱�׼����֧��  NUM 335 10      ����
                strInputString = strInputString & Lpad(dbl����Ա��������, 10)       '37  ���ι���Ա������������֧��  NUM 345 10      ����
                strInputString = strInputString & Lpad(dbl����Ա�ǻ�������, 10)     '38  ���ι���Ա�ǻ�����������֧��    NUM 355 10      ����
                strInputString = strInputString & Lpad(dbl��ҵ���ղ���, 10)         '39  ������ҵ���ղ���֧��    NUM 365 10      ����
                strInputString = strInputString & Lpad(dbl���շ�Χ���Ը�, 10)       '40  ���α������Ը�  NUM 375 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣����������Էѷ���+����Է�+Ѫ���Է�+�����Էѣ��ĸ����˻���ֺ�ķ���   ����
                strInputString = strInputString & Lpad(dbl�Ǳ����Ը�, 10)           '41  ���ηǱ����Ը�  NUM 385 10      ����
                
'                strInputString = strInputString & Lpad(dbl��������֧��, 10) '���ķ���:���λ�����������֧�� ��������:����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
'                strInputString = strInputString & Lpad(dbl�ǲ�������֧��, 10)   '���ķ���:���ηǻ�����������֧����������:����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
'                strInputString = strInputString & Lpad(dbl���շ�Χ���Ը�, 10)   '���ķ���:���α��շ�Χ���Ը���������:�޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
            End If
        Else
            If intinsure = TYPE_������ Then
            Else
                '2005-08-02����������
                strInputString = strInputString & String(10, " ")   'Lpad(dbl����ǰ����ͳ���ۼ�, 10)
            End If
            strInputString = strInputString & String(10, " ")    '���ķ���:���λ��������ʻ�֧��(������������㣬��ʾ�����ʻ�֧��)
            strInputString = strInputString & String(10, " ")    '���ķ���:���β��������ʻ�֧��(������������㷵��0)
            strInputString = strInputString & String(10, " ")    '���ķ���:���λ���ͳ��֧��
            strInputString = strInputString & String(10, " ")    '���ķ���:���λ���ͳ���Ը�
            strInputString = strInputString & String(10, " ")    '���ķ���:���β���ͳ��֧��
            strInputString = strInputString & String(10, " ")    '���ķ���:���β���ͳ���Ը�
            
            If intinsure = TYPE_������ Then
                '2004/9/11 ����ֵ�䶯��ԭ���Ĺ���Ա��������ҵ��������һ���ֶθ�Ϊ����Ա������Ŀ����ҵ����������¼���ֱ��ڵ�33��34��35��36���ֶ�������
                strInputString = strInputString & String(10, " ")   '33  ���ι���Ա�𸶱�׼����֧��  NUM 301 10      ����
                strInputString = strInputString & String(10, " ")   '34  ���ι���Ա������������֧��  NUM 311 10      ����
                strInputString = strInputString & String(10, " ")   '35  ���ι���Ա�ǻ�����������֧��    NUM 321 10      ����
                strInputString = strInputString & String(10, " ")   '36  ������ҵ���ղ���֧��    NUM 331 10      ����
                strInputString = strInputString & String(10, " ")   '37  ���α������Ը�  NUM 341 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣����������Էѷ���+����Է�+Ѫ���Է�+�����Էѣ��ĸ����˻���ֺ�ķ���   ����
                strInputString = strInputString & String(10, " ")   '38  ���ηǱ����Ը�  NUM 351 10      ����
            Else
                '2005-08-02����������
                strInputString = strInputString & String(10, " ")   '36  ���ι���Ա�𸶱�׼����֧��  NUM 335 10      ����
                strInputString = strInputString & String(10, " ")   '37  ���ι���Ա������������֧��  NUM 345 10      ����
                strInputString = strInputString & String(10, " ")   '38  ���ι���Ա�ǻ�����������֧��    NUM 355 10      ����
                strInputString = strInputString & String(10, " ")   '39  ������ҵ���ղ���֧��    NUM 365 10      ����
                strInputString = strInputString & String(10, " ")   '40  ���α������Ը�  NUM 375 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣����������Էѷ���+����Է�+Ѫ���Է�+�����Էѣ��ĸ����˻���ֺ�ķ���   ����
                strInputString = strInputString & String(10, " ")   '41  ���ηǱ����Ը�  NUM 385 10      ����
                
'                strInputString = strInputString & String(10, " ")    '������:����Ա�������ֶΰ����ż��Ѳ������ֺͻ���ͳ���Ը����ֵĹ���Ա����֧�� ���ķ���
'                strInputString = strInputString & String(10, " ")    '������:����Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧�����ò��֣���������ͳ������޶�֣���ȥ����Ա����֧����ȫ������"���α��շ�Χ���Ը�"����  ���ķ���
'                strInputString = strInputString & String(10, " ")    '������:�޶����⣫�ż����Ը����֣������ʻ���ֺ󣩣������Է�ȥ����������    ���ķ���
            End If
        End If
                
        '����ӦΪ��
        strInputString = strInputString & Lpad(Trim(CStr(dbl�𸶱�׼)), 10)    '�𸶱�׼��������:����סԺ�ż���  NUM 315 10      Ժ����д
        strInputString = strInputString & Lpad(.ת�ﵥ��, 6)     'ת�ﵥ��
        strInputString = strInputString & Lpad(Get�������(intҵ��, .�������), 1)      '�������
        
       '���� 2005-08-16 �˴������жϴ����кͿ�����
'        If intInsure <> TYPE_���������� Then
            '2004/9/11 ���Ӳα����2
            strInputString = strInputString & Lpad(.�α����2, 1)    ''42  �α����2   CHAR    378 1   0 ������ 1 ��ҵ 2 ����Ա�����鿨���    Ժ��
            strInputString = strInputString & Lpad(.�α����3, 1)    '�α����3:0 �󱣡�1 �±��������鿨���
'        End If
        strInputString = strInputString & Lpad(.ְ����ҽ���, 1)       'ְ����ҽ���
        
        strInputString = strInputString & Lpad(.��ϱ���, 16)    '��ϱ���
        
        strInputString = strInputString & Lpad(strҽ��, 6)    'ҽʦ����
        strInputString = strInputString & Lpad(UserInfo.���, 6)    '����Ա����
        strInputString = strInputString & Lpad(Substr(.�������, 1, 28), 30)  '�������
        
        '??49  ���������ʶ    CHAR    439 1   1������2��ת��3δ����4������5������סԺ���� Ժ��
        'A-������B-��ת��C-δ����D-������E-����
        If bytType = 1 Then
            strInputString = strInputString & Lpad(Get�������_����(lng����ID, lng��ҳID), 1)
            strInputString = strInputString & Lpad(str��Ժ����, 8)      '��Ժ����
        Else
            strInputString = strInputString & "1"    '���������ʶ
            strInputString = strInputString & String(8, " ")      '��Ժ����
        End If
        
        '���� 2005-08-16,�˴������ж�
'        If intInsure = TYPE_���������� Then       '������
'        Else
            strInputString = strInputString & String(16, " ")      '����ʱ��
'        End If
        strInputString = strInputString & String(10, " ")      '�������
    End With
    
    Dim blnReturn As Boolean
    
    'ҵ������
    If bln������� Then
        blnReturn = ҵ������_����(IIf(intinsure = TYPE_����������, 2, 1), 1006, strInputString, intinsure)
        
    Else
        blnReturn = ҵ������_����(IIf(intinsure = TYPE_����������, 2, 1), 1002, strInputString, intinsure)
    End If
    
    If blnReturn = False Then Exit Function
    
    If bln������� = True Then
        str���㷽ʽ = Get���㷽ʽ(strInputString, intinsure)
        ����ҽ������ = True
        Exit Function
    End If
    Dim i As Long
    If intinsure = TYPE_���������� Then
'        i = 225 - 10
        '2005-08-02����������
        i = 275 - 10
    Else
        i = 241 - 10
    End If
    
    If intinsure = TYPE_���������� Then
        '2005-08-02����������
        dbl�����ʻ���� = Val(Substr(strInputString, i - 60, 10))
        dblͳ��֧���ۼ� = Val(Substr(strInputString, i - 50, 10))  '�����ͳ��֧���ۼ�=����ͳ���ۼƣ�����ͳ���ۼ�
        dbl����ǰ����ͳ���ۼ� = Val(Substr(strInputString, i, 10))
        dbl����󼲲�ͳ���ۼ� = Val(Substr(strInputString, i - 40, 10))
    Else
        dbl�����ʻ���� = Val(Substr(strInputString, i - 40, 10))
        dblͳ��֧���ۼ� = Val(Substr(strInputString, i - 30, 10))  '�����ͳ��֧���ۼ�=����ͳ���ۼƣ�����ͳ���ۼ�
    End If
    
    dbl�����ʻ�֧�� = Val(Substr(strInputString, i + 10, 10))   '���λ��������ʻ�֧��=������������㣬��ʾ�����ʻ�֧��
    dbl�����ʻ�֧�� = Val(Substr(strInputString, i + 20, 10))   '���β��������ʻ�֧��    NUM 221 10  ������������㷵��0
    dbl����ͳ��֧�� = Val(Substr(strInputString, i + 30, 10))   '���λ���ͳ��֧��    NUM 231 10      ����
    dbl����ͳ���Ը� = Val(Substr(strInputString, i + 40, 10))   '���λ���ͳ���Ը�    NUM 241 10      ����
    dbl����ͳ��֧�� = Val(Substr(strInputString, i + 50, 10))   '���β���ͳ��֧��    NUM 251 10      ����
    dbl����ͳ���Ը� = Val(Substr(strInputString, i + 60, 10))   '���β���ͳ���Ը�    NUM 261 10      ����
    
    With g�������_����
        .��ǰ�����ʻ���� = .��ǰ�����ʻ���� - dbl�����ʻ�֧��
        .��ǰ�����ʻ���� = .��ǰ�����ʻ���� - dbl�����ʻ�֧��
        .��ǰͳ���ۼ� = .��ǰͳ���ۼ� + dbl����ͳ��֧�� + dbl����ͳ��֧��
    End With
    
    If intinsure = TYPE_������ Then
        '2004/9/11 ����ֵ�䶯��ԭ���Ĺ���Ա��������ҵ��������һ���ֶθ�Ϊ����Ա������Ŀ����ҵ����������¼���ֱ��ڵ�33��34��35��36���ֶ�������
        dbl����Ա�𸶱�׼���� = Val(Substr(strInputString, i + 70, 10)) '33  ���ι���Ա�𸶱�׼����֧��  NUM 301 10      ����
        dbl����Ա�������� = Val(Substr(strInputString, i + 80, 10))     '34  ���ι���Ա������������֧��  NUM 311 10      ����
        dbl����Ա�ǻ������� = Val(Substr(strInputString, i + 90, 10))   '35  ���ι���Ա�ǻ�����������֧��    NUM 321 10      ����
        dbl��ҵ���ղ��� = Val(Substr(strInputString, i + 100, 10))      '36  ������ҵ���ղ���֧��    NUM 331 10      ����
        dbl���շ�Χ���Ը� = Val(Substr(strInputString, i + 110, 10))    '37  ���α������Ը�  NUM 341 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣����������Էѷ���+����Է�+Ѫ���Է�+�����Էѣ��ĸ����˻���ֺ�ķ���   ����
        dbl�Ǳ����Ը� = Val(Substr(strInputString, i + 120, 10))        '38  ���ηǱ����Ը�  NUM 351 10      ����
    Else
        '2005-08-02����������
        dbl����Ա�𸶱�׼���� = Val(Substr(strInputString, i + 70, 10)) '36  ���ι���Ա�𸶱�׼����֧��  NUM 335 10      ����
        dbl����Ա�������� = Val(Substr(strInputString, i + 80, 10))     '37  ���ι���Ա������������֧��  NUM 345 10      ����
        dbl����Ա�ǻ������� = Val(Substr(strInputString, i + 90, 10))   '38  ���ι���Ա�ǻ�����������֧��    NUM 355 10      ����
        dbl��ҵ���ղ��� = Val(Substr(strInputString, i + 100, 10))      '39  ������ҵ���ղ���֧��    NUM 365 10      ����
        dbl���շ�Χ���Ը� = Val(Substr(strInputString, i + 110, 10))    '40  ���α������Ը�  NUM 375 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣����������Էѷ���+����Է�+Ѫ���Է�+�����Էѣ��ĸ����˻���ֺ�ķ���   ����
        dbl�Ǳ����Ը� = Val(Substr(strInputString, i + 120, 10))        '41  ���ηǱ����Ը�  NUM 385 10      ����
        
'        dbl��������֧�� = Val(Substr(strInputString, i + 70, 10))     '���λ�����������֧��    NUM 271 10  1�� �������ҵ���ո��ֶΰ�������ͳ���Ը����ֵ���ҵ����֧��2��   ����ǹ���Ա�������ֶΰ����ż��Ѳ������֡�����ͳ���Ը����ֵĹ���Ա����֧��������ͳ������޶��ڹ���Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����  ����
'        dbl�ǲ�������֧�� = Val(Substr(strInputString, i + 80, 10))     '���ηǻ�����������֧��  NUM 281 10  1�� �������ҵ���ո��ֶ��ǲ���ͳ���Ը����ֵ���ҵ����֧��2�� ����ǹ���Ա�������ֶ��ǳ�������ͳ������޶�ֵĹ���Ա����֧������������ͳ������޶��Ա����֧����ʣ��������"���α��շ�Χ���Ը�"����
'        dbl���շ�Χ���Ը� = Val(Substr(strInputString, i + 90, 10))     '���α��շ�Χ���Ը�  NUM 291 10  �޶����⣨ȥ�������󣩣��ż����Ը����֣������ʻ���ֺ󣩣��������Էѷ��ã��Ǳ��շ���+����Է�   ����
    End If
    
    '���̲���:
    '   ����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,
    '   �ʻ��ۼ�����_IN(�����ʻ����),�ʻ��ۼ�֧��_IN(ͳ��֧���ۼ�),�ۼƽ���ͳ��_IN(��������֧��),�ۼ�ͳ�ﱨ��_IN(�����ʻ�֧��),סԺ����_IN(�������),����_IN(�𸶱�׼),�ⶥ��_IN(���շ�Χ���Ը�),ʵ������_IN(�𸶱�׼),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(����ͳ��֧��),�����Ը����_IN(����ͳ���Ը�),����ͳ����_IN(����ͳ��֧��),ͳ�ﱨ�����_IN(����ͳ���Ը�),���Ը����_IN(�ǲ�������֧��),�����Ը����_IN(��),
    '   �����ʻ�֧��_IN(�����ʻ�֧��),֧��˳���_IN(�������;ת�ﵥ��;��ϱ���),��ҳID_IN(��ҳid),��;����_IN(Null),��ע_IN(�������),
    '   ����_IN,��ҩ��_IN,��ҩ��_IN,��ҩ��_IN,����_IN,���Ʒ�_IN,����_IN,����Է�_IN,�������Ʒ�_IN,���������Է�_IN,
    '   �������Էѷ���_IN,�Ǳ��շ���_IN,ͳ�����_IN,������_IN,Ѫ��_IN,Ѫ���Է�_IN,����ǰ�����ʻ����_IN,����ǰ�����˻����_IN,����ǰͳ���ۼ�_IN,
    '   ����Ա�𸶱�׼����_IN , ����Ա��������_IN, ����Ա�ǻ�������_IN, ��ҵ���ղ���_IN, �Ǳ����Ը�_IN
'2005-08-02����������
    '   ,ҩ���Է�_IN,����ǰ����ͳ���ۼ�_IN,����󼲲�ͳ���ۼ�_IN
    
    '�ܺ�ȫ���� 2004-10-20
    '����ʱ�����ֵӦ��Ϊ����:
    '   1)����/��ҩ��/��ҩ��/��ҩ��/����/���Ʒ�/Ѫ��/Ѫ���Է�/����/����Է�/�������Ʒ�/���������Է�/�������Էѷ���/��ҽ�Ʊ��շ���/ҩ���Է�
    '   2)���������ܶ�ҲӦ��ͬʱΪ��

    If bln���� Then
        gstrSQL = "zl_���ս����¼_insert(" & IIf(bytType = 0, 1, 2) & "," & lng����ID & "," & intinsure & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
            dbl�����ʻ���� & "," & dblͳ��֧���ۼ� & "," & dbl��������֧�� & "," & dbl�����ʻ�֧�� & "," & g�������_����.������� & "," & dbl�𸶱�׼ & "," & dbl���շ�Χ���Ը� & "," & dbl�𸶱�׼ & "," & _
            -curTotal & "," & dbl����ͳ��֧�� & "," & dbl����ͳ���Ը� & "," & dbl����ͳ��֧�� & "," & dbl����ͳ���Ը� & "," & dbl�ǲ�������֧�� & ",Null," & _
            dbl�����ʻ�֧�� & ",'" & Get�������(intҵ��, g�������_����.�������) & ";" & g�������_����.ת�ﵥ�� & ";" & g�������_����.��ϱ��� & "'," & lng��ҳID & ",null,'" & Lpad(Substr(g�������_����.�������, 1, 28), 30) & "'," & _
            -dbl���� & "," & -dbl��ҩ�� & "," & -dbl��ҩ�� & "," & -dbl��ҩ�� & "," & -dbl���� & "," & -dbl���Ʒ� & "," & -dbl���� & "," & -dbl����Է� & "," & -dbl�������Ʒ� & "," & -dbl���������Է� & "," & _
            -dbl�������Էѷ��� & "," & -Abs(dbl�Ǳ��շ���) & "," & dbl���� & "," & -dbl������ & "," & -dblѪ�� & "," & -dblѪ���Է� & "," & dbl����ǰ���(1) & "," & dbl����ǰ���(2) & "," & dbl����ǰ���(3) & "," & _
            dbl����Ա�𸶱�׼���� & "," & dbl����Ա�������� & "," & dbl����Ա�ǻ������� & "," & dbl��ҵ���ղ��� & "," & -Abs(dbl�Ǳ����Ը�) & "," & _
             -Abs(dblҩ���Է�) & "," & dbl����ǰ����ͳ���ۼ� & "," & dbl����󼲲�ͳ���ۼ� & " )"
    Else
        gstrSQL = "zl_���ս����¼_insert(" & IIf(bytType = 0, 1, 2) & "," & lng����ID & "," & intinsure & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
            dbl�����ʻ���� & "," & dblͳ��֧���ۼ� & "," & dbl��������֧�� & "," & dbl�����ʻ�֧�� & "," & g�������_����.������� & "," & dbl�𸶱�׼ & "," & dbl���շ�Χ���Ը� & "," & dbl�𸶱�׼ & "," & _
            curTotal & "," & dbl����ͳ��֧�� & "," & dbl����ͳ���Ը� & "," & dbl����ͳ��֧�� & "," & dbl����ͳ���Ը� & "," & dbl�ǲ�������֧�� & ",Null," & _
            dbl�����ʻ�֧�� & ",'" & Get�������(intҵ��, g�������_����.�������) & ";" & g�������_����.ת�ﵥ�� & ";" & g�������_����.��ϱ��� & "'," & lng��ҳID & ",null,'" & Lpad(Substr(g�������_����.�������, 1, 28), 30) & "'," & _
            dbl���� & "," & dbl��ҩ�� & "," & dbl��ҩ�� & "," & dbl��ҩ�� & "," & dbl���� & "," & dbl���Ʒ� & "," & dbl���� & "," & dbl����Է� & "," & dbl�������Ʒ� & "," & dbl���������Է� & "," & _
            dbl�������Էѷ��� & "," & dbl�Ǳ��շ��� & "," & dbl���� & "," & dbl������ & "," & dblѪ�� & "," & dblѪ���Է� & "," & dbl����ǰ���(1) & "," & dbl����ǰ���(2) & "," & dbl����ǰ���(3) & "," & _
            dbl����Ա�𸶱�׼���� & "," & dbl����Ա�������� & "," & dbl����Ա�ǻ������� & "," & dbl��ҵ���ղ��� & "," & dbl�Ǳ����Ը� & "," & _
            dblҩ���Է� & "," & dbl����ǰ����ͳ���ۼ� & "," & dbl����󼲲�ͳ���ۼ� & " )"
    End If
    zlDatabase.ExecuteProcedure gstrSQL, "����" & IIf(bytType = 1, "סԺ", "����") & "�շ�����"
    ����ҽ������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, ByVal intinsure As Integer) As Boolean
    Dim lng����ID As Long
    �������_���� = Set�����������(False, lng����ID, cur�����ʻ�, lng����ID, strSelfNo, intinsure)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    �������_���� = False
End Function
Private Function Set�����������(ByVal bln���� As Boolean, lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, strSelfNo As String, ByVal intinsure As Integer) As Boolean
  '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID��
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Dim curTotal As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim strInfor As String  '�������ķ��ش�
    Dim dbl���� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double
    
    '2005-08-02����������
    Dim dblҩ���Է� As Double, dbl����ǰ����ͳ���ۼ� As Double, dbl����󼲲�ͳ���ۼ� As Double
    
    Dim dbl���� As Double, dbl���Ʒ� As Double, dbl���� As Double, dbl����Է� As Double
    Dim dbl�������Ʒ� As Double, dbl���������Է� As Double, dbl�������Էѷ��� As Double
    Dim dbl�Ǳ��շ��� As Double, dblѪ�� As Double, dblѪ���Է� As Double
    Dim dbl������ As Double     '��Դ�����������
    Dim dbl���� As Double
    Dim strҽ�� As String, str��ϸ As String      '��ϸ��
    Dim str���ұ��� As String, str��Ŀ���� As String
    Dim str��Ŀͳ�Ʒ��� As String, strTmp As String
    Dim intҵ�� As Integer, lng����ID As Long
    Dim strNO As String, lng��¼���� As Long
    
    Dim lng������� As Long
    Dim dbl�����ʻ���� As Double, dblͳ��֧���ۼ� As Double, dbl�����ʻ�֧�� As Double
    Dim dbl�����ʻ�֧�� As Double, dbl����ͳ��֧�� As Double, dbl����ͳ���Ը� As Double
    Dim dbl����ͳ��֧�� As Double, dbl����ͳ���Ը� As Double, dbl��������֧�� As Double
    Dim dbl�ǲ�������֧�� As Double, dbl���շ�Χ���Ը� As Double
    
    Dim dbl����ǰ�����ʻ����  As Double, dbl����ǰ�����˻���� As Double, dbl����ǰͳ���ۼ� As Double
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim lngTmp As Long, lng����ID As Long
    Dim strInsertSQL As String
    Static str����ʱ�� As String
    Static lng����id1 As Long
    
    intҵ�� = IIf(bln����, 1, 0)
     Set����������� = False
   
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ���������ڽ��㣬����ʹ��Ȧ��ӿڣ����������Ǯ���������ڣ������ӿ��ڽ��
    '���������Ҫͨ��������������ȡ����Ȧ�����ǽӿڷ��أ���Ҫ�޸�
    
    On Error GoTo errHand
    
    '���¶���
    If ��ȡ�������_����(IIf(intinsure = TYPE_����������, 2, 1), intinsure) = False Then
        Exit Function
    End If
    
    If bln���� Then
        '��֤�Ƿ�Ϊ�ò��˵�IC��
        gstrSQL = "Select ����ID,���� From  �����ʻ� where ����id=" & lng����ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵�ҽ����"
        If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "�ò����ڱ����ʻ����޼�¼!"
            Exit Function
        End If
        
        If g�������_����.IC���� <> Nvl(rsTemp!����) Then
            ShowMsgbox "�ò��˵�IC���������,�����ǲ����������˵�IC��!"
            Exit Function
        End If
        'ȷ���������,ת�ﵥ��,��ϱ���,�������
        ' ֧��˳���_IN(�������;ת�ﵥ��;��ϱ���),��ע(�������_IN)
        gstrSQL = "Select ֧��˳���,��ע from ���ս����¼  where ��¼ID=" & lng����ID
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�������"
        If rsTemp.RecordCount = 0 Then
            Err.Raise 9000, gstrSysName, "�ڽ����¼���޽����¼!"
            Exit Function
        End If
        Dim strArr
        strArr = Split(Nvl(rsTemp!֧��˳���), ";")
        
        '�������;ת�ﵥ��;��ϱ���
        '1-��ͨ����("1", "A"),2-��������("3", "7")
        '3-�����("5", "B"),4-������������("S", "T")
        If UBound(strArr) >= 2 Then
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g�������_����.ת�ﵥ�� = strArr(1)
            g�������_����.��ϱ��� = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g�������_����.ת�ﵥ�� = strArr(1)
        Else
            g�������_����.������� = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
        End If
        g�������_����.������� = Nvl(rsTemp!��ע)
        
        
        'ȷ���˷Ѽ�¼
        '�˷�
          gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
                    " where A.NO=B.NO and A.��¼����=B.��¼����  and A.��¼״̬=2 and B.����ID=[1]"
          Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����˷�", lng����ID)
          If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "�����ڲ��˷��ó�����¼!"
            Exit Function
          Else
            lng����ID = rsTemp("����ID")
          End If
          
    End If
    '�򿪱��ν�����ϸ��¼ '--���ұ���Ӧ���Ǳ�ʶ����+����
    gstrSQL = " " & _
        "  Select Rownum ��ʶ��,A.ID,A.����ID,A.�շ�ϸĿid,A.NO,A.���,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ��,A.������ as ҽ��,H.��� as ҽ�����, " & _
        "      A.����*A.���� as ����,A.���㵥λ,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ�� as ʵ�ս��,F.����ֵ,G.id as ����id,G.ͳ��ȶ�,G.סԺ�ȶ�, " & _
        "      A.ҽ�����,A.�շ����,B.���� as ��Ŀ����,B.���� as ��Ŀ����,Nvl(J.��ʶ��,Nvl(B.��ʶ����||B.��ʶ����,B.����)) as ���ұ���, " & _
        "      D.��Ŀ���� ҽ������,D.��Ŀ���� as ҽ������,J.���� as ����,D.�Ƿ�ҽ��,C.���� ��������,E.���� �ܵ�����, " & _
        "      L.����,L.����,L.����,L.ҽ����,L.��Ա���,L.��λ����,L.˳���,L.����֤��,L.�ʻ����,L.��ǰ״̬,L.����ID,L.��ְ,L.�����,L.�Ҷȼ�,L.����ʱ�� " & _
        "  From (Select * From ������ü�¼ Where ��¼״̬<>0 and ����ID=" & IIf(bln����, lng����ID, lng����ID) & " and  Nvl(���ӱ�־,0)<>9 ) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,���ű� E,  " & _
        "       (Select U.*,K.����ֵ From �շ���� U,���ղ��� K where U.���=K.������ and K.����=" & intinsure & "  ) F, " & _
        "       (Select distinct Q.ҩƷid,Q.��ʶ��,T.���� From ҩƷĿ¼ Q,ҩƷ��Ϣ R,ҩƷ���� T  Where  Q.ҩ��id=R.ҩ��id and R.����=T.���� ) J, " & _
        "       ����֧������ G,��Ա�� H,�����ʻ� L" & _
        "  Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID(+) and A.����id=L.����id and L.����=" & intinsure & " and A.�շ����=F.����(+)  and d.����id=G.id and a.�շ�ϸĿid=J.ҩƷid(+) " & _
        "        And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����= " & intinsure & " and a.������=H.����(+) " & _
        "  Order by A.ID"
        
    '�ϴ�������ϸ��¼
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡ���ν��ʷ�����ϸ"
    
    With rs��ϸ
        '��Ҫ������շѵ���
        If str����ʱ�� <> Format(!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS") Or lng����id1 <> Nvl(!����ID, 0) Then
              str����ʱ�� = Format(!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
              lng����id1 = Nvl(!����ID, 0)
              g�������_����.���㿪ʼ = True
        Else
              g�������_����.���㿪ʼ = False
        End If
    
        If Not .EOF Then
            lng����ID = Nvl(!����ID, 0)
            strҽ�� = Nvl(!ҽ�����)
            If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                strҽ�� = Substr(strҽ��, 1, 6)
            End If
            lng����ID = Nvl(!����ID, 0)
            '����׼��Ŀ
            gstrSQL = "Select * from ������׼��Ŀ  where ����ID=  " & lng����ID
            zlDatabase.OpenRecordset rs��׼��Ŀ, gstrSQL, "��ȡ������Ŀ����"
        End If
        Do While Not .EOF
        
            If lng����ID <> 0 And bln���� = False Then
                '��һ��,ȷ��������շ�ϸĿ
                rs��׼��Ŀ.Filter = 0
                rs��׼��Ŀ.Filter = "����=0 And ����=1 and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
                If rs��׼��Ŀ.EOF Then
                    Err.Raise 9000, gstrSysName, "�շ�ϸĿΪ��" & Nvl(!��Ŀ����) & "������Ŀ���ǲ��������趨����Ŀ."
                    Exit Function
                End If
                '�ڶ���,ȷ������ı��մ���
                rs��׼��Ŀ.Filter = 0
                rs��׼��Ŀ.Filter = "����=1 And ����=1 and  �շ�ϸĿid=" & Nvl(!����id, 0)
                If rs��׼��Ŀ.EOF Then
                    Err.Raise 9000, gstrSysName, "�ڽ����д����˽�������ı���֧������,���ܼ�����"
                    Exit Function
                End If
                '������,'ȷ����ֹ���շ�ϸĿ
                rs��׼��Ŀ.Filter = 0
                rs��׼��Ŀ.Filter = "����=0 And ����=2 and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
                If Not rs��׼��Ŀ.EOF Then
                    Err.Raise 9000, gstrSysName, "�շ�ϸĿΪ��" & Nvl(!��Ŀ����) & "������Ŀ�Ǳ���ֹʹ�õ���Ŀ." & vbCrLf & "���ܼ���!"
                    Exit Function
                End If
                '���Ĳ�,'ȷ����ֹ�Ĵ���
                rs��׼��Ŀ.Filter = 0
                rs��׼��Ŀ.Filter = "����=1 And ����=2 and �շ�ϸĿid=" & Nvl(!����id, 0)
                If Not rs��׼��Ŀ.EOF Then
                    Err.Raise 9000, gstrSysName, "�ڽ����д����˽�ֹʹ�õı���֧������,���ܼ�����"
                End If
            End If
            strTmp = Nvl(!����ֵ)
            lng����ID = Nvl(!����ID, 0)
            'ȷ���������
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str��Ŀͳ�Ʒ��� = ""
                Else
                    str��Ŀͳ�Ʒ��� = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                '����
                '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
                    
                If Nvl(!����, 0) <> TYPE_���������� And Val(Nvl(!��λ����, "99")) = 0 And Nvl(!��ְ, 0) = 3 And Nvl(!�Ƿ�ҽ��, 0) = 1 Then   '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '������    ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�������ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dbl���� = 1
                Else
                    '2005-10-14 ZHQ
                    '����󲡻��������Ƿ�סԺ�ȶ����
                    If (g�������_����.������� = 3 And IsParaBig(intinsure)) Or _
                        (IsParaQ(intinsure) And intinsure = TYPE_������ And g�������_����.ְ����ҽ��� = "Q") Then
                        dbl���� = Nvl(!סԺ�ȶ�, 0) / 100
                    Else
                        dbl���� = Nvl(!ͳ��ȶ�, 0) / 100
                    End If
                End If
                
                If Nvl(!����, 0) = TYPE_������ And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dbl���� = Val(Nvl(!ҽ������)) / 100
                End If
                
                If Nvl(!����, 0) = TYPE_������ And g�������_����.ְ����ҽ��� = "Q" Then
                    '�����Q��ҵ����,�������Ϊ100�Է�,�������Ǳ��շ�����
                    If dbl���� = 0 Then
                        '�Է�100
                        strTmp = ""
                    Else
                        '�ԷѲ��ַ��� �������Էѷ�����
                    End If
                End If
                
                If Nvl(!ҽ������) = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If Nvl(!ҽ������) = "���" Then
                    strTmp = "����"
                End If

'-----------------------------------------------------������ʼ-------------------------------------------------------
                If dbl���� <> 0 Then
                    Select Case strTmp
                            Case "����"
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                            
                            Case "��ҩ��"
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                '2005-08-02����������
                                If Nvl(!����, 0) = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                End If
                                
                            Case "��ҩ��"
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                '2005-08-02����������
                                If Nvl(!����, 0) = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                End If
                                
                            Case "��ҩ��"
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                '2005-08-02����������
                                If Nvl(!����, 0) = TYPE_������ Then
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                Else
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                End If
                                
                            Case "����"
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                
                            Case "���Ʒ�"
                                dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                            Case "����"
                                '�����кͿ������Դ����ô���ͬ,
                                '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                          
                                dbl���� = dbl���� + Round(Nvl(!ʵ�ս�� * dbl����, 0), 5)
                                dbl����Է� = dbl����Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                            
                            Case "Ѫ��"
                                '2004/9/11 ���˺�����Ѫ��
                                dblѪ�� = dblѪ�� + Round(Nvl(!ʵ�ս�� * dbl����, 0), 5)
                                dblѪ���Է� = dblѪ���Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                          
                            Case "�������Ʒ�"
                                '2004/9/11��ǰ:�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                                '2004/9/11�Ժ�:�������뿪�������㷽ʽһ�£�����ͳ�ﲿ�ֵĽ��
                                'If intinsure = TYPE_������ Then
                                '    dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!ʵ�ս��, 0), 5)
                                'Else
                                    dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!ʵ�ս��, 0) * dbl����, 5)
                                'End If
                                
                                '�������Ʒ��Էѵļ��㷽ʽ��ͬ,
                                '�����нӿ��ڴ�����ܽ��ʱֻ���������Ʒѽ��л���,���������ԷѲ��ֲ��ټ���
                                dbl���������Է� = dbl���������Է� + Round(Nvl(!ʵ�ս��, 0) * (1 - dbl����), 5)
                                
                        End Select
                    Else
                
                        'ȫ���Ǳ���Ϊ0����Ŀ(��������),�ֱ�Դ������ǿ����������жϷ��ڲ�ͬ���ֶ�
                        If intinsure = TYPE_������ Then
                            '�����з���dbl�Ǳ��շ���
                            dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(!ʵ�ս��, 5)
                        Else
                            '����������dbl������
                            dbl������ = dbl������ + Round(!ʵ�ս��, 5)
                        End If
                    End If
            Else
                dbl���� = 1
                str��Ŀͳ�Ʒ��� = ""
            End If

            '�ϴ���ϸ��¼,ʵʱҽ����ϸ����
            '����������ϸ�ϴ�
            If gbln������ϸʱʵ�ϴ� Then
                
                    If Nvl(!����, 0) = TYPE_���������� Then '������
                        str��ϸ = Lpad(gstrҽԺ����_����, 6)     'ҽԺ����    CHAR    1   6       Ժ����д
                        str��ϸ = str��ϸ & Lpad(Nvl(!ҽ����), 10)  '���ձ��    CHAR    7   10      Ժ����д
                    Else
                        str��ϸ = Lpad(gstrҽԺ����_����, 4)     'ҽԺ����    CHAR    1   4       Ժ��
                        str��ϸ = str��ϸ & Lpad(Nvl(!ҽ����), 8)   '���˱��    CHAR    5   8       Ժ��
                    End If
                    
                    str��ϸ = str��ϸ & Space(10)   '��־��  CHAR    13  10  ������ϸ�Կո�λ,סԺ��סԺ��  Ժ��
                    str��ϸ = str��ϸ & Lpad(g�������_����.�������, 4)     '�������
                    
                    'Modified By ���� 2004-07-29 ԭ�򣺴���NO��
                    str��ϸ = str��ϸ & Lpad(Mid(Nvl(!NO, "00000000"), 2, 7), 10)   '������  NUM 27  10      Ժ��
                    str��ϸ = str��ϸ & Lpad(CStr(.AbsolutePosition), 10)           '������Ŀ���    NUM 37  10  ��Ӧ�����ŵļǼ���Ŀ���    Ժ��
                    
                    str��ϸ = str��ϸ & Space(10)           'ҽ����  CHAR    47  10  ������Ӧҽ����ҽ����¼�ţ�������ϸ��û��ҽ����ҽԺ�Կո�λ    Ժ��
                    str��ϸ = str��ϸ & Get�������(intҵ��, Nvl(!�Ҷȼ�, 0))         '�������    CHAR    57  1   ȡֵ���"�������"˵��  Ժ��
                    
                    str��ϸ = str��ϸ & Rpad(Format(!�Ǽ�ʱ��, "yyyymmddHHmmss"), 16)      '��������ʱ�䣨Ͷҩʱ�䣩    DATETIME    58  16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ    Ժ��
                    
                    str��ϸ = str��ϸ & Lpad(Nvl(!���ұ���), 20)      '��Ŀ����    CHAR    74  20  �Ƽ���Ŀ����    Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!��Ŀ����), 20)      '��Ŀ����    CHAR    94  20      Ժ��
        
                    If !�Ƿ�ҽ�� = 1 Then
                        str��ϸ = str��ϸ & Lpad(1 - dbl����, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                    Else
                        str��ϸ = str��ϸ & Lpad(1, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(str��Ŀͳ�Ʒ���, 1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                    
                    If Nvl(!����, 0) = TYPE_���������� Then
                        '2005-08-02����������
                        str��ϸ = str��ϸ & Lpad(Nvl(!����), 10)  '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                        str��ϸ = str��ϸ & Lpad(Nvl(!ʵ�ʼ۸�), 10) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
                    Else
                        str��ϸ = str��ϸ & Lpad(Nvl(!����), 6)  '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                        str��ϸ = str��ϸ & Lpad(Nvl(!ʵ�ʼ۸�), 8) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(Nvl(!���㵥λ), 4) '��λ    CHAR    135 4       Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!����), 20)      '����    CHAR    139 20  �����Ƭ����    Ժ��
                
                    str��ϸ = str��ϸ & Lpad(Nvl(!ҽ��), 8)      'ҽʦ����    CHAR    159 8       Ժ��
                    str��ϸ = str��ϸ & Lpad(g�������_����.��ϱ���, 16)      '��ϱ���    CHAR    167 16      Ժ��
                    str��ϸ = str��ϸ & Lpad(Substr(g�������_����.�������, 1, 28), 30)   '�������    CHAR    183 30      Ժ��
                    str��ϸ = str��ϸ & Space(16)     '����ʱ��    DATETIME    213 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ��Ժ�˿ո�λ  ����
                                        
                    '�ϴ���ϸ
                    '1003    7   230 ʵʱҽ����ϸ�����ύ
                    Set����������� = ҵ������_����(IIf(Nvl(!����, 0) = TYPE_����������, 2, 1), 1003, str��ϸ, intinsure)
                    If Set����������� = False Then
                        ShowMsgbox "�������ʱҽ����ϸ�����ύʧ��,���ܼ���!"
                        Exit Function
                    End If
                    
                    '�ϴ�ҽ����ϸ
                    If Nvl(!ҽ�����, 0) <> 0 Then
                        If ҽ����ϸ�����ύ(!ҽ�����, "", str��Ŀͳ�Ʒ���, intinsure) = False Then
                            ShowMsgbox "ҽ����ϸ�����ύʧ��,���ܼ���!"
                            Exit Function
                        End If
                    End If
                    
                    'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                    'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                    zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            End If
            '�����ܶ�,����
            curTotal = curTotal + Round(Nvl(!ʵ�ս��, 0), 5)
            .MoveNext
        Loop
    End With
    
    Set����������� = False
    
    '���������Ϊ��
    If ����ҽ������(intinsure, 0, lng����ID, 0, IIf(bln����, lng����ID, lng����ID), lng����ID, bln����, False, 0, _
        dbl����, dbl��ҩ��, dbl��ҩ��, dbl��ҩ��, dbl����, dbl���Ʒ�, dblѪ��, dblѪ���Է�, dbl����, dbl����Է�, _
        dbl�������Ʒ�, dbl���������Է�, dbl�������Էѷ���, dbl�Ǳ��շ���, dbl������, dblҩ���Է�, curTotal, strҽ��, strInfor) = False Then
        Exit Function
    End If
    
    Set����������� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, ByVal intinsure As Integer) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Err = 0
    On Error GoTo errHand:
    ����������_���� = Set�����������(True, lng����ID, cur�����ʻ�, lng����ID, "", intinsure)
    Exit Function
errHand:
    ����������_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String, ByVal intinsure As Integer) As Boolean
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��λ�� As String
    Dim strת�ﵥ�� As String
    Dim lng���� As Long
    Dim int���� As Long
    
    
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo errHand
    
    '��ȡ���˵���ر�����Ϣ

    gstrSQL = "select ����,����ID,��Ա���,ҽ����,˳���,�Ҷȼ� From �����ʻ� where  ����=" & intinsure & "  and ����id=" & lng����ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    int���� = Nvl(rsTemp!����)
    strת�ﵥ�� = Nvl(rsTemp!��Ա���)
    lng���� = IIf(intinsure = 83, 2, 1)
    If lng���� = 2 Then
        strInfor = Lpad(gstrҽԺ����_����, 6)                   'ҽԺ����    CHAR    1   6      Y   Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ����), 10)      '���ձ��    CHAR    7   10      Ժ����д
    Else
        strInfor = Lpad(gstrҽԺ����_����, 4)                   'ҽԺ����    CHAR    1   4       Y   Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ����), 8)       '���ձ��    CHAR    5   8       Y   Ժ��
    End If
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!˳���, 1), 4)        '�������    NUM 13  4   ���������Ժʱ�������  Y   Ժ��
    
    
    '�ڲ���ʶ:5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ
    'ҽ����ʶ:2-סԺ����,4-��ͥ��������,O-��������סԺ����,Q-���˱��ս���
    
    str������� = Decode(Val(Nvl(rsTemp!�Ҷȼ�, 0)), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '��ȡ������Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') ��Ժ����ʱ��," & _
            " to_char(A.��Ժ����,'yyyyMMdd') ��Ժ����" & _
            " From ������ҳ A,���ű� B,������Ϣ C" & _
            " Where A.����id=C.����id and C.����id=[1]" & _
            "       and A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ��Ϣ", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    str��Ժ���� = Nvl(rsTemp!��Ժ����)
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!סԺ��, 0), 10)       '��־��  CHAR    17  10      Y   Ժ�������ݶ�Ϊ�գ�סԺ��ΪסԺ��
    strInfor = strInfor & Lpad(Nvl(rsTemp!��Ժ����), 8)         '��Ժ���� Date 27  8   ����ʵ����Ժ���ڣ���ʽΪyyyymmdd    Y   Ժ��
    strInfor = strInfor & Rpad(Nvl(rsTemp!��Ժ����ʱ��), 16)    '�Ǽ�ʱ��    DATETIME    35  16  ��ȷ���룬���ݷ��غ��ʽΪyyyymmddhhmiss�����Կո�λ  Y   Ժ��
    strInfor = strInfor & Lpad(str�������, 1)                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��

    gstrSQL = "Select ����ID,����,����� From ��λ״����¼ D where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ��Ϣ", CLng(Nvl(rsTemp!��ǰ����ID, 0)), CLng(Nvl(rsTemp!��ǰ����, 0)))
    If rsTemp.EOF Then
        str��λ�� = Space(10)
    Else
        str��λ�� = Trim(Nvl(rsTemp!�����)) & "��" & Trim(Nvl(rsTemp!����)) & "��"
        str��λ�� = Lpad(str��λ��, 10)
        str��λ�� = Substr(str��λ��, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����||'~^||'||b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1 and a.����id=" & lng����ID & " and a.��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    Dim str��Ժ��ϱ��� As String
    Dim str��Ժ�������  As String
    Dim strȷ����ϱ��� As String
    Dim strȷ���������  As String
    
    If rsTemp.EOF Then
        str��Ժ��ϱ��� = ""
        str��Ժ������� = ""
        strȷ����ϱ��� = ""
        strȷ��������� = ""
    Else
        str��Ժ������� = Nvl(rsTemp!��Ժ���)
        strȷ��������� = Nvl(rsTemp!ȷ�����)
        If InStr(1, str��Ժ�������, "~^||") <> 0 Then
            str��Ժ��ϱ��� = Split(str��Ժ�������, "~^||")(0)
            str��Ժ������� = Split(str��Ժ�������, "~^||")(1)
        Else
            str��Ժ��ϱ��� = ""
            str��Ժ������� = ""
        End If
        If InStr(1, strȷ���������, "~^||") <> 0 Then
            strȷ����ϱ��� = Split(strȷ���������, "~^||")(0)
            strȷ��������� = Split(strȷ���������, "~^||")(1)
        Else
            strȷ����ϱ��� = ""
            strȷ��������� = ""
        End If
    End If
    '2006-02-20 ZHQ Modify
    '������Ժ���Ĭ��������ϣ���������ϲ�Ϊ��׼��ϣ�������������ϣ����Ǽ�����ʱ��������Ժ
    If Len(Trim(str��Ժ��ϱ���)) = 0 Then
        MsgBox "��Ժ��ϷǱ�׼ICD-10��ϣ�ҽ�������������Ժ��", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)     '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
    strInfor = strInfor & Lpad(Substr(str��Ժ�������, 1, 28), 30) '��Ժ�������    CHAR    68  30      y Ժ��
    strInfor = strInfor & Lpad(strȷ����ϱ���, 16)     'ȷ����ϱ���    CHAR    98  16      N   Ժ��
    strInfor = strInfor & Lpad(Substr(strȷ���������, 1, 28), 30) 'ȷ���������    CHAR    114 30      N   Ժ��
    strInfor = strInfor & Lpad(str��Ժ����, 20)         '�Ʊ�����    CHAR    144 20  �磺�ڿ�    Y   Ժ��
    strInfor = strInfor & Lpad(Substr(str��λ��, 1, 10), 10)             '��λ��  CHAR    164 10  �磺2003��12��  N   Ժ��
    strInfor = strInfor & Lpad(strת�ﵥ��, 6)          'ת�ﵥ��    CHAR    174 6       N   Ժ��
    strInfor = strInfor & Space(8)                      '��Ժʱ��    DATE    180 8   ϵͳ���û��߽������ݵĳ�Ժʱ���Զ����ɣ�ҽԺ���ÿո�λ���ɡ�  N   ��
    strInfor = strInfor & "A"                           '�����־    CHAR    188 1   A ��Ժ�Ǽǣ�M �޸���Ժ״̬��Cȡ����Ժ�Ǽ�   Y   Ժ��
    strInfor = strInfor & Space(16)                     '����ʱ��    DATATIME    189 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  N   ����
    
    '1004    9   206 ʵʱסԺ�Ǽ������ύ
    ��Ժ�Ǽ�_���� = ҵ������_����(lng����, 1004, strInfor, intinsure)
    If ��Ժ�Ǽ�_���� = False Then
        Err.Raise 9000, gstrSysName, "ʵʱסԺ�Ǽ������ύʧ��!"
        Exit Function
    End If
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & int���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
                
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��λ�� As String
    Dim strת�ﵥ�� As String
    Dim lng���� As Long
    Dim int���� As Integer
        
    On Error GoTo errHand
    
    '��ȡ���˵���ر�����Ϣ

    gstrSQL = "select ����,����ID,��Ա���,ҽ����,˳���,�Ҷȼ� From �����ʻ� where ����id=" & lng����ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "������Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    int���� = rsTemp!����
    strת�ﵥ�� = Nvl(rsTemp!��Ա���)
    lng���� = IIf(rsTemp!���� = 83, 2, 1)

    If lng���� = 2 Then
        strInfor = Lpad(gstrҽԺ����_����, 6) 'ҽԺ����    CHAR    1   6      Y   Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ����), 10)     '���ձ��    CHAR    7   10      Ժ����д
    Else
        strInfor = Lpad(gstrҽԺ����_����, 4) 'ҽԺ����    CHAR    1   4       Y   Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ����), 8)     '���ձ��    CHAR    5   8       Y   Ժ��
    End If
    strInfor = strInfor & Lpad(Nvl(rsTemp!˳���, 1), 4)      '�������    NUM 13  4   ���������Ժʱ�������  Y   Ժ��
    
    '�ڲ���ʶ:5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ
    'ҽ����ʶ:2-סԺ����,4-��ͥ��������,O-��������סԺ����,Q-���˱��ս���
    
    str������� = Decode(Val(Nvl(rsTemp!�Ҷȼ�, 0)), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '��ȡ������Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') ��Ժ����ʱ��," & _
            " to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����" & _
            " From ������ҳ A,���ű� B,������Ϣ C" & _
            " Where A.����id=C.����id and C.����id=" & lng����ID & _
            "       and A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ��Ϣ", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    str��Ժ���� = Nvl(rsTemp!��Ժ����)
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!סԺ��, 0), 10)      '��־��  CHAR    17  10      Y   Ժ�������ݶ�Ϊ�գ�סԺ��ΪסԺ��
    strInfor = strInfor & Lpad(Nvl(rsTemp!��Ժ����), 8)      '��Ժ���� Date 27  8   ����ʵ����Ժ���ڣ���ʽΪyyyymmdd    Y   Ժ��
    strInfor = strInfor & Rpad(Nvl(rsTemp!��Ժ����ʱ��), 16)      '�Ǽ�ʱ��    DATETIME    35  16  ��ȷ���룬���ݷ��غ��ʽΪyyyymmddhhmiss�����Կո�λ  Y   Ժ��
    
    strInfor = strInfor & Lpad(str�������, 1)                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��
    
    gstrSQL = "Select ����ID,����,����� From ��λ״����¼ D where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ��Ϣ", CLng(Nvl(rsTemp!��ǰ����ID, 0)), CLng(Nvl(rsTemp!��ǰ����, 0)))
    If rsTemp.EOF Then
        str��λ�� = Space(10)
    Else
        str��λ�� = Trim(Nvl(rsTemp!�����)) & "��" & Trim(Nvl(rsTemp!����)) & "��"
        str��λ�� = Lpad(str��λ��, 10)
        str��λ�� = Substr(str��λ��, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����||'~^||'||b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1 and a.����id=" & lng����ID & " and a.��ҳid=" & lng��ҳID
         
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    Dim str��Ժ��ϱ��� As String
    Dim str��Ժ�������  As String
    Dim strȷ����ϱ��� As String
    Dim strȷ���������  As String
    
    If rsTemp.EOF Then
        str��Ժ��ϱ��� = ""
        str��Ժ������� = ""
        strȷ����ϱ��� = ""
        strȷ��������� = ""
    Else
        str��Ժ������� = Nvl(rsTemp!��Ժ���)
        strȷ��������� = Nvl(rsTemp!ȷ�����)
        If InStr(1, str��Ժ�������, "~^||") <> 0 Then
            str��Ժ��ϱ��� = Split(str��Ժ�������, "~^||")(0)
            str��Ժ������� = Split(str��Ժ�������, "~^||")(1)
        Else
            str��Ժ��ϱ��� = ""
            str��Ժ������� = ""
        End If
        If InStr(1, strȷ���������, "~^||") <> 0 Then
            strȷ����ϱ��� = Split(strȷ���������, "~^||")(0)
            strȷ��������� = Split(strȷ���������, "~^||")(1)
        Else
            strȷ����ϱ��� = ""
            strȷ��������� = ""
        End If
    End If
    
    strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
    strInfor = strInfor & Lpad(Substr(str��Ժ�������, 1, 28), 30) '��Ժ�������    CHAR    68  30      y Ժ��
    strInfor = strInfor & Lpad(strȷ����ϱ���, 16)  'ȷ����ϱ���    CHAR    98  16      N   Ժ��
    strInfor = strInfor & Lpad(Substr(strȷ���������, 1, 28), 30) 'ȷ���������    CHAR    114 30      N   Ժ��
    
    strInfor = strInfor & Lpad(str��Ժ����, 20)  '�Ʊ�����    CHAR    144 20  �磺�ڿ�    Y   Ժ��
    strInfor = strInfor & str��λ��              '��λ��  CHAR    164 10  �磺2003��12��  N   Ժ��
    
    strInfor = strInfor & Lpad(strת�ﵥ��, 6)   'ת�ﵥ��    CHAR    174 6       N   Ժ��
    strInfor = strInfor & Space(8)   '��Ժʱ��    DATE    180 8   ϵͳ���û��߽������ݵĳ�Ժʱ���Զ����ɣ�ҽԺ���ÿո�λ���ɡ�  N   ��
    strInfor = strInfor & "C"   '�����־    CHAR    188 1   A ��Ժ�Ǽǣ�M �޸���Ժ״̬��Cȡ����Ժ�Ǽ�   Y   Ժ��
    strInfor = strInfor & Space(16)   '����ʱ��    DATATIME    189 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  N   ����
    
    '1004    9   206 ʵʱסԺ�Ǽ������ύ
    ��Ժ�Ǽǳ���_���� = ҵ������_����(lng����, 1004, strInfor, intinsure)
    If ��Ժ�Ǽǳ���_���� = False Then
        ShowMsgbox "ʵʱסԺ�Ǽǳ��������ύʧ��!"
        Exit Function
    End If
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & int���� & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���� = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '
    '����HIS��Ժ
    Dim rsTemp As New ADODB.Recordset
    Dim int����
    '---����HIS�����ҽ������Ƿֿ�ִ��,��˲����ڰ����Ժ��ʱ��ı�ҽ���ʻ���״̬,ֱ�ӷ���Ϊ��
    '--�ܺ�ȫ 2005-07-21 ���뵥���ֵ���ʾ
    gstrSQL = "select Nvl(a.����ֵ,'0') as ����ֵ From ���ղ��� a,������ҳ b " & _
            "  where a.����(+)=b.���� and a.������(+)='�����ֳ�Ժ��ʾ' and b.����id=" & lng����ID & " and b.��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�Ƿ񵥲�����ʾ"
    
    If rsTemp!����ֵ = "1" Then
        If MsgBox("ҽ���ر����ѣ�" & vbCr & vbCr & "    �����ּ�飬�Ƿ�����������", _
            vbYesNo + vbDefaultButton1 + vbInformation, gstrSysName) = vbYes Then
            ��Ժ�Ǽ�_���� = False
        Else
            ��Ժ�Ǽ�_���� = True
        End If
    Else
        ��Ժ�Ǽ�_���� = True
    End If
    
'    '----
'    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & rsTemp!���� & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    Dim int����
'    gstrSQL = "select ���� From ������ҳ where  ����id=" & lng����ID & " and ��ҳid=" & lng��ҳID
'
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵Ĳα�����"
'
'    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & rsTemp!���� & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    gstrSQL = "select ��ǰ״̬ from �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˱����ʻ�����Ժ״̬"
    
    If Not rsTemp.EOF Then
        If rsTemp!��ǰ״̬ = 1 Then
            ��Ժ�Ǽǳ���_���� = True
        Else
            MsgBox "�ò��˵ı����ʻ���ǰΪ��Ժ״̬,���ܽ��г���!����ȡ����Ժ��������ִ�б�����"
            ��Ժ�Ǽǳ���_���� = False
        End If
    Else
        ��Ժ�Ǽǳ���_���� = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ��ȡ������Ϣ_����(ByVal lng����ID As Long, ByVal intinsure As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵������Ϣ,����ֵ����G�������
    '--�����:lng����id
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    '��ȡҽ�����������Ϣ�������¹��ýṹ��
        
    gstrSQL = "" & _
        "   Select *" & _
        "   From �����ʻ�" & _
        "   Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˵������Ϣ", intinsure, lng����ID)
    
    If Not rsTemp.EOF Then
        With g�������_����
            .IC���� = Nvl(rsTemp!����, 0)
            .���˱�� = Nvl(rsTemp!ҽ����)
            .ҽ������ = IIf(intinsure = 83, 2, 1) ' NVL(rsTemp!����, 1)
            .������� = Nvl(rsTemp!˳���, 0)
            .ת�ﵥ�� = Nvl(rsTemp!��Ա���)
            .���������ʻ���� = Nvl(rsTemp!�ʻ����, 0)
            .���������ʻ���� = Val(Nvl(rsTemp!����֤��))
            
            .ְ����ҽ��� = Decode(Nvl(rsTemp!��ְ, 1), 1, "A", 2, "B", 3, "L", 4, "T", 5, "Q", "E", 6, "")
            .������� = Nvl(rsTemp!�Ҷȼ�, 0)
            .�α����3 = Nvl(rsTemp!��λ����, 0)
            '.���� = NVL(rsTemp!ͳ�ﱨ���ۼ�, 0)
        End With
    End If
End Sub
Private Function Get�������_����(lng����ID As Long, lng��ҳID As Long) As String
    '����:��ȡ���������ʶ
    '     A-������B-��ת��C-δ����D-������E-����
    '??49  ���������ʶ    CHAR    439 1   1������2��ת��3δ����4������5������סԺ���� Ժ��
    'A-������B-��ת��C-δ����D-������E-����
    
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.��Ժ���" & _
             " From ������ A,��������Ŀ¼ B " & _
             " Where A.����ID=[1] And A.����ID=B.ID(+) And A.��ҳID=[2]" & _
             "       And A.������� in (2,3)" & _
             " Order by A.������� Desc"
    Set rsInNote = zlDatabase.OpenSQLRecord(strTmp, "ҽ���ӿ�", lng����ID, lng��ҳID)
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = Nvl(rsInNote!��Ժ���)
    End If
    Get�������_���� = Decode(strTmp, "����", "1", "��ת", "2", "δ��", "3", "����", "4", "����", "5", "1")
End Function

Public Function IS��ǰסԺ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '�ж�סԺ�����Ƿ��ǵ�ǰ��סԺ����
    '2004/09/21
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Max(��ҳid) as ��ҳid From ������ҳ where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������ҳid"
    IS��ǰסԺ = False
    If rsTemp.EOF Then
        IS��ǰסԺ = True
        Exit Function
    End If
    IS��ǰסԺ = lng��ҳID >= Nvl(rsTemp!��ҳID)
End Function

Private Function ���������ö�(ByVal intinsure As Integer, ByVal lng����ID As Long, ByVal str���� As String, ByVal int�㷨 As Byte, ByVal str�շ���� As String, _
    ByVal int���� As Integer, ByVal strҽ����Ŀ���� As String, ByVal str����ʱ�� As String, ByVal int������Ŀ�� As Integer, _
    ByVal dbl���� As Double, ByVal dbl���� As Double, ByVal dbl��� As Double, ByVal dblӤ���� As Double, ByVal dblͳ����� As Double, ByVal dbl��ҵ���� As Double, _
    ByVal dbl��׼���� As Double, dblMoney As Variant) As Boolean
        
        '���������õĽ��
        '   2004/9/21
        Dim dbl���� As Double, strTmp As String
        Dim dblTemp(0 To 20) As Double
        Dim rsTemp As New ADODB.Recordset
        Dim i As Long
        Const strTemp = "����;��ҩ��;��ҩ��;��ҩ��;����;���Ʒ�;����;����Է�;Ѫ��;Ѫ���Է�;�������Ʒ�;���������Է�;�������Էѷ���;�Ǳ��շ���;������;ҩ���Է�"
        Dim strArr
        
        ���������ö� = False
    
        strArr = Split(strTemp, ";")
        If str���� = "" Or InStr(1, str����, ";") = 0 Then
            ShowMsgbox "δ�����շ����Ķ�Ӧ��ϵ,�뵽�������������!"
            Exit Function
        End If
        Err = 0
        On Error GoTo errHand:
        
        dbl���� = dblͳ����� / 100
        strTmp = Split(str����, ";")(0)
        
        '-----------------------------------------------------------------------
        '���㱣��
        '---�����㷨=2(�������)����Ŀ,���ڲ��ǰ��ձ������м���,�����ҪԤ������dbl����=1,
        '---ʹ����뱨���ָ���ж�,�������Ϊ��ʼֵ=0���±���Ϊȫ�Է���Ŀ�����ָ�
        
        '-------------------------------------------------------------------------
        '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
        If g�������_����.ҽ������ <> 2 And g�������_����.ְ����ҽ��� = "L" And g�������_����.�α����3 = "0" And int������Ŀ�� = 1 Then '���󱣺�������Ա����ҽ����Ŀ
            '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
            '  ������  ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�������ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
            dbl���� = 1
        End If
        '-------------------------------------------------------------------------
        If strҽ����Ŀ���� = "����" Then
            strTmp = "�������Ʒ�"
        End If
        If strҽ����Ŀ���� = "���" Then
            strTmp = "����"
        End If
                
        '---����һ�����������ƷѺʹ����ñ��������ļ��
        If (strTmp = "�������Ʒ�" Or strTmp = "����") And dbl���� = 0 Then
            MsgBox "�������Ʒ��û�����õı�������Ϊ0,���ȼ�������Ƿ���ȷ"
            Exit Function
        End If
                
        If g�������_����.ҽ������ <> 2 And (g�������_����.ְ����ҽ��� = "L" Or _
             g�������_����.ְ����ҽ��� = "T") Then
            '�����L���ݺ�T����ľͰ���ҵ��������
            dbl���� = dbl��ҵ���� / 100
        End If
                
        '����Ǵ�λ,���谴���·�ʽ����,��������ͳ���������,������Ϊ100���Է�,���ֿ������ʹ�����
        If str�շ���� = "J" Then
            
            gstrSQL = "" & _
                "   Select ���Ӵ�λ From ���˱䶯��¼ " & _
                "   Where ����=" & int���� & " and ����id=" & lng����ID & _
                "         And ( (to_date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600 between ��ʼʱ�� and ��ֹʱ��) or" & _
                "               ( ��ֹʱ�� is null  and ��ʼʱ��<=to_date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600)) " & _
                "         And ���� is not null"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ���Ƿ�Ϊ����!"
            
            If rsTemp.RecordCount >= 1 Then
            '--������Ų��������䶯��¼�еĴ�(�ֹ��Ƿ�),�򱾴�ѡ�����ΪNOTHING,��Ҫ��dbl����=0
            '--����м�¼,���ж��Ƿ�Ϊ���Ӵ�λ,�ǵĻ�dbl����=0
                If rsTemp!���Ӵ�λ = 1 Then
                    '��ʾ������λ,Ϊȫ�Է�
                    dbl���� = 0
                Else
                    If int�㷨 = 2 Then
                         dbl���� = 1
                    End If
                End If
            End If
        End If
                
        '--Ӥ����ֱ�ӹ���ȫ�ԷѲ���
        If dblӤ���� <> 0 Then
            dbl���� = 0
        End If
        '------------------------------------------------------�������-----------------------------------------------------
        '---��Ϊֻ�д�λ���ô����޶��,����ֻ�Դ�λ�����޶��жϲ�����
        '---�Ƿ�����ȷ�ĸ��Ӵ�λ,��Ҫ���ݱ䶯��¼�����ж� ����\ʱ��,����ȷ���Ǳ��˵��Զ����㴲λ��
        '---(���Ƕ��ֹ�¼��Ĵ�λ��,�ж��������б仯)
        '-------------------------------------------------------------------------------------------------------------------
        If dbl���� <> 0 Then
            For i = 0 To UBound(strArr)
                If strTmp = strArr(i) Then
                    '"����;��ҩ��;��ҩ��;��ҩ��;����;���Ʒ�;����;����Է�;Ѫ��;Ѫ���Է�;�������Ʒ�;���������Է�;�������Էѷ���;�Ǳ��շ���;������;ҩ���Է�"
                    Select Case strTmp
                        Case "����", "����"
                            dblTemp(i) = dblTemp(i) + Round(dbl��� * dbl����, 5)
                            '���㱣�����Է�
                            '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                            '---��������졢������������ԷѲ��ּ���dbl������
                            dblTemp(12) = dblTemp(12) + Round(dbl��� * (1 - dbl����), 5)
                        Case "��ҩ��", "��ҩ��", "��ҩ��"
                            dblTemp(i) = dblTemp(i) + Round(dbl��� * dbl����, 5)
                            '2005-08-02����������
                            'ҩ���Էѷ����ض��ֶ���
                            If intinsure = TYPE_������ Then
                                dblTemp(12) = dblTemp(12) + Round(dbl��� * (1 - dbl����), 5)
                                dblTemp(15) = 0
                            Else
                                dblTemp(15) = dblTemp(15) + Round(dbl��� * (1 - dbl����), 5)
                            End If
                        Case "���Ʒ�"
                            If int�㷨 = 2 Then
                                '���ڿ��ܴ����������ʺ��ٽ���,���Ƿ񳬹��޶���밴�վ���ֵ�����ж�
                                If dbl��׼���� <= Abs(dbl����) Then
                                    '����޶�<����,������������Ϊ�޶�*����
                                    dblTemp(i) = dblTemp(i) + dbl��׼���� * dbl���� * Sgn(dbl���)
                                    '���޲���=(����-�޶�)*����,  �����ɽ��ķ��ž���,�����м���dbl�������Էѷ���
                                    dblTemp(12) = dblTemp(12) + Round((Abs(dbl����) - dbl��׼����) * Abs(dbl����) * Sgn(dbl���), 5)
                                Else
                                    '����޶�>=����,��ȫ���������Ʒ���
                                    dblTemp(12) = dblTemp(12) + Round(dbl���, 5)
                                End If
                            Else
                                dblTemp(i) = dblTemp(i) + Round(dbl��� * dbl����, 5)
                                dblTemp(12) = dblTemp(12) + Round(dbl��� * (1 - dbl����), 5)
                            End If
                        Case "����", "Ѫ��", "�������Ʒ�"
                                dblTemp(i) = dblTemp(i) + Round(dbl��� * dbl����, 5)
                                dblTemp(i + 1) = dblTemp(i + 1) + Round(dbl��� * (1 - dbl����), 5)
                        End Select
                        Exit For
                End If
            Next
        Else
            'ȫ���Ǳ���Ϊ0����Ŀ(��������),�ֱ�Դ������ǿ����������жϷ��ڲ�ͬ���ֶ�
            If intinsure = TYPE_������ Then
                '�����з���dbl�Ǳ��շ���
                dblTemp(12) = dblTemp(12) + Round(dbl���, 5)
            Else
                '����������dbl������
                dblTemp(14) = dblTemp(14) + Round(dbl���, 5)
            End If
        End If
        dblMoney = dblTemp
        ���������ö� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    Exit Function
End Function

Private Function Get��ȡ���β��˽�����Ϣ(ByVal lng����ID As Long, lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
    '��ȡ���β��˽���ʱ����Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim strArr
    Get��ȡ���β��˽�����Ϣ = False
    
    '��ȡҽ�����������Ϣ�������¹��ýṹ��
    gstrSQL = "" & _
            "   Select A.����,A.ҽ����,A.��ְ,A.�α����1,A.�α����2,a.�α����3,A.�α����4,A.�α����5," & _
            "          b.����,B.�Ա�,B.��������,(sysdate-b.��������)/365 as ���� ,b.���֤��," & _
            "          c.סԺ����,c.ʵ������ �𸶱�׼,C.֧��˳���,c.�ʻ��ۼ�֧��,c.����ǰ�����ʻ����,c.����ǰ�����˻����,c.����ǰͳ���ۼ�" & _
            "   From �����ʻ� A,������Ϣ B,���ս����¼ C " & _
            "   Where A.����=[1] And A.����ID=[2]" & _
            "         and A.����id=B.����id and A.����id=C.����id and C.����=2 and C.��ҳID=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˵������Ϣ", intinsure, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        ShowMsgbox "û�����ҽ��������Ϣ��"
        Exit Function
    End If
    
    strArr = Split(Nvl(rsTemp!֧��˳���, ";;") & ";;", ";")
    
    With g�������_����
        .IC���� = Nvl(rsTemp!����, 0)
        .���˱�� = Nvl(rsTemp!ҽ����)
        .ҽ������ = IIf(intinsure = 83, 2, 1) ' NVL(rsTemp!����, 1)
        
        .���� = Nvl(rsTemp!����)
        .�Ա� = Nvl(rsTemp!�Ա�)
        .�������� = Format(rsTemp!��������, "yyyy-mm-dd")
        .���� = Nvl(rsTemp!����, 0)
        .���֤�� = Nvl(rsTemp!���֤��)
        .������� = Nvl(rsTemp!סԺ����, 1) - 1
        
        .ְ����ҽ��� = Decode(Nvl(rsTemp!��ְ, 1), 1, "A", 2, "B", 3, "L", 4, "T", 5, "Q", "E", 6, "")
        .���������ʻ���� = Nvl(rsTemp!����ǰ�����ʻ����, 0)
        .���������ʻ���� = Nvl(rsTemp!����ǰ�����˻����)
        .��ǰ״̬ = 1
        .ͳ���ۼ� = Nvl(rsTemp!����ǰͳ���ۼ�, 0)
        .�½ɷѻ��� = 0
        .�α����1 = Nvl(rsTemp!�α����1)
        .�α����2 = Nvl(rsTemp!�α����2)
        .�α����3 = Nvl(rsTemp!�α����3)
        .�α����4 = Nvl(rsTemp!�α����4)
        .�α����5 = Nvl(rsTemp!�α����5)
        .�ʻ�״̬ = 0
         'C.֧��˳��� �� �������;ת�ﵥ��;��ϱ��룩
        .ת�ﵥ�� = strArr(1)
        .������� = strArr(0)
        .���� = Nvl(rsTemp!�𸶱�׼, 0)
    End With
    Get��ȡ���β��˽�����Ϣ = True
    Exit Function
errHand:
        If ErrCenter = 1 Then Resume
End Function

Private Function ���η��ý���Ԥ����(rsExse As Recordset, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As String
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '      �ֶ�:��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID, _
    '           �շ����,�շ�ϸĿID,�շ�����,��������,���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��, _
    '           �Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    
    Dim rsTemp As New ADODB.Recordset, rs���� As New ADODB.Recordset, rs�ʻ�״̬ As New ADODB.Recordset
    '�����Ӧ��ϵͳ
    Dim dblMoney(0 To 20) As Double
    '����;��ҩ��;��ҩ��;��ҩ��;����;���Ʒ�;����;����Է�;Ѫ��;Ѫ���Է�;�������Ʒ�;���������Է�;�������Էѷ���;�Ǳ��շ���;������;ҩ���Է�
    Dim curTotal As Double
    Dim strInfor As String  '�������ķ��ش�
    Dim str���㷽ʽ  As String
    Dim strҽ�� As String
    Dim str��Ժ���� As String
    Dim strסԺ�� As String, strTmp As String
    Dim intMouse As Integer
    
    On Error GoTo errHand
    intMouse = Screen.MousePointer
    
    '���������ǰ����֤���
    Screen.MousePointer = 1
    
    
    '��ȡ���β�����Ϣ
    If Get��ȡ���β��˽�����Ϣ(lng����ID, lng��ҳID, intinsure) = False Then
        Screen.MousePointer = intMouse
        ���η��ý���Ԥ���� = ""
        Exit Function
    End If
    
    Screen.MousePointer = intMouse
    
    gstrSQL = " Select B.סԺ���� ��ҳID,to_char(A.��Ժ����,'yyyy') ��Ժ���,A.��Ժ����,B.סԺ��" & _
              " From ������ҳ A,������Ϣ B" & _
              " Where B.����ID=[1] And A.��ҳID=[2] And A.����ID=B.����ID"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ժʱ��", lng����ID, lng��ҳID)
    
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyymmdd")
    strסԺ�� = Nvl(rsTemp!סԺ��)
    
    
    '��ȷ�������� ���˺�2004/06/15,��Ϊ���ֽ���ʱ��������
    If rsTemp.EOF Then
        strTmp = Get��Ժ���(lng����ID, lng��ҳID, , True)
    Else
        If IsNull(rsTemp!��Ժ����) Then
            strTmp = Get��Ժ���(lng����ID, lng��ҳID, , True)
        Else
            strTmp = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, , True)
        End If
    End If
    If InStr(1, strTmp, "|") <> 0 Then
        g�������_����.��ϱ��� = Split(strTmp, "|")(1)
        g�������_����.������� = Split(strTmp, "|")(0)
    End If
    strTmp = ""

    '���»�ȡ��¼
    Set rs���� = GetסԺ�����¼(lng����ID, lng��ҳID, intinsure)
    If rs����.RecordCount <= 0 Then
        ShowMsgbox "����δ���õ�ҽ����Ŀ�����ܽ���!"
        Exit Function
    End If
    
    
    With rs����
        '�ϴ�������ϸ
        curTotal = 0
        strҽ�� = Nvl(!ҽ�����)
        If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
            strҽ�� = Substr(strҽ��, 1, 6)
        End If
        
        Do While Not .EOF
            lng����ID = Nvl(!����ID, 0)
            
            If ���������ö�(intinsure, lng����ID, Nvl(!����ֵ), Nvl(!�㷨, 0), Nvl(!�շ����), _
                Nvl(!����, 0), Nvl(!ҽ����Ŀ����), Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS"), _
                Nvl(!������Ŀ��, 0), Nvl(!����, 0), Nvl(!�۸�, 0), Nvl(!���, 0), Nvl(!Ӥ����, 0), _
                Nvl(!סԺ�ȶ�, 0), Val(Nvl(!ҽ����Ŀ����, 0)), Nvl(!��׼����, 0), dblMoney) = False Then
                Exit Function
            End If
            
            curTotal = curTotal + Nvl(!���, 0)
            .MoveNext
        Loop
        
        '����;��ҩ��;��ҩ��;��ҩ��;����;���Ʒ�;����;����Է�;Ѫ��;Ѫ���Է�;�������Ʒ�;���������Է�;�������Էѷ���;�Ǳ��շ���;������;ҩ���Է�
        If ����ҽ������(intinsure, 1, lng����ID, lng��ҳID, 0, 0, False, True, g�������_����.����, _
            dblMoney(0), dblMoney(1), dblMoney(2), dblMoney(3), dblMoney(4), dblMoney(5), dblMoney(8), _
            dblMoney(9), dblMoney(6), dblMoney(7), dblMoney(10), dblMoney(11), dblMoney(12), dblMoney(13), _
            dblMoney(14), dblMoney(15), curTotal, strҽ��, str���㷽ʽ, strסԺ��, str��Ժ����) = False Then
            Exit Function
        End If
        
        g�������_����.֧����� = curTotal
    End With
    
    ���η��ý���Ԥ���� = str���㷽ʽ
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, ByVal intinsure As Integer) As String
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '      �ֶ�:��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID, _
    '           �շ����,�շ�ϸĿID,�շ�����,��������,���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��, _
    '           �Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    Dim rsTemp As New ADODB.Recordset, rs���� As New ADODB.Recordset, rs�ʻ�״̬ As New ADODB.Recordset
    Dim curTotal As Double
    Dim lng��ҳID As Long
    Dim cur�����Ը� As Currency, cur�����ʻ� As Currency
    Dim str��Ժ��� As String, str������� As String, str����ʱ�� As String, str����ʱ�� As String
    Dim strInfor As String  '�������ķ��ش�
    Dim dbl���� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double
    
    '2005-08-02����������
    Dim dblҩ���Է� As Double
    
    Dim dbl���� As Double, dbl���Ʒ� As Double, dbl���� As Double, dbl����Է� As Double
    Dim dbl�������Ʒ� As Double, dbl���������Է� As Double, dbl�������Էѷ��� As Double
    Dim dbl�Ǳ��շ��� As Double, dbl������ As Double    '��Դ�����������
    Dim dblѪ�� As Double, dblѪ���Է� As Double
       
    Dim dbl���� As Double, dbl�𸶱�׼ As Double
    Dim str���㷽ʽ  As String
    
    Dim strҽʦ���� As String, str����Ա���� As String, str���������ʶ As String, strTmp As String
    Dim strҽ�� As String, str��ϸ As String, str���ұ��� As String, str��Ŀͳ�Ʒ��� As String
    Dim str��Ժ���� As String, dbl��Ŀ���� As Double
    Dim strסԺ�� As String, str������ʽ��㷽ʽ As String
    Dim intMouse As Integer
    
    On Error GoTo errHand
    intMouse = Screen.MousePointer
    
    
    If rsExse.EOF Then
        ShowMsgbox "��ǰû����ϸ��¼!"
        Exit Function
    End If
    
    lng��ҳID = Nvl(rsExse!��ҳID, 0)
    rsExse.MoveLast
    If Nvl(rsExse!��ҳID, 0) <> lng��ҳID Then
        ShowMsgbox "��׼�Զ��סԺ�Ĳ��˽���һ�ν���"
         Exit Function
    End If
    rsExse.MoveFirst
    g�������_����.���ν��� = IIf(IS��ǰסԺ(lng����ID, lng��ҳID) = False, True, False)
    
    '--�ж��Ƿ��ܹ���Ϊ������ʽ��н���
    gstrSQL = "select ��ǰ״̬ from �����ʻ� where ����id=[1]"
    Set rs�ʻ�״̬ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ�״̬", lng����ID)
    
    If Nvl(rs�ʻ�״̬!��ǰ״̬, 0) = 0 And g�������_����.���ν��� = False Then
        '���ν��㲻��������
        '����ʻ�״̬=0,���жϴ���������Ƿ����סԺ����,�������ܽ���
        Do While Not rsExse.EOF
            '������סԺ���þͲ�׼�����﷽ʽ����
            If rsExse!�����־ = 2 Then
                MsgBox "�ò��˱����ʻ�δ�Ǽ���Ժ,���Ǵ�������а���סԺ����,���鲡�˵��ʻ�״̬�����½���"
                סԺ�������_���� = ""
                Exit Function
            End If
            rsExse.MoveNext
        Loop
        
        If MsgBox("�ò��˴��������ֻ��������ʷ���,�Ƿ���������н���?", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then
            rsExse.MoveFirst
            'ѡ�����﷽ʽ�����,Ԥ��ɹ����ؽ���ֵ
            If ��������������_����(rsExse, str������ʽ��㷽ʽ, intinsure) Then
                סԺ�������_���� = str������ʽ��㷽ʽ
                Exit Function
            Else
            '����Ԥ��ʧ�ܷ��ؿմ�
                סԺ�������_���� = ""
                Exit Function
            End If
        Else
            '���������﷽ʽ�����ֱ�ӷ��ؿմ�
            סԺ�������_���� = ""
            Exit Function
        End If
    Else
        rsExse.Filter = 0
        rsExse.Filter = "�����־=1"
        '������˵��ʻ�״̬Ϊ��Ժ,����밴��סԺ��ʽ����;
        '������������ʷ���,��ʾ�Ƿ���סԺ��ʽ����,
        '�����ͬ�����,ֱ��ȡ���������ݺ������½���
        If Not rsExse.EOF Then
            If MsgBox("������ð���������ʷ���,�Ƿ���סԺ��ʽ����Щ���ý��н���?", vbQuestion + vbOKCancel + vbDefaultButton2) = vbCancel Then
                סԺ�������_���� = ""
                Exit Function
            End If
        End If
    End If
    '----------------------���������ʷ��ý��㷽ʽ�ж�------------
    
    
    '���������ǰ����֤���
    Screen.MousePointer = 1
    If ��ݱ�ʶ_����(4, lng����ID, intinsure) = "" Then
        Screen.MousePointer = intMouse
'        Call WriteDebugInfor_����("סԺ�������_����", lng����id)
        סԺ�������_���� = ""
        Exit Function
    End If
    
    Screen.MousePointer = intMouse
    
    cur�����ʻ� = g�������_����.���������ʻ����

    gstrSQL = " Select B.סԺ���� ��ҳID,to_char(A.��Ժ����,'yyyy') ��Ժ���,A.��Ժ����,B.סԺ��" & _
              " From ������ҳ A,������Ϣ B" & _
              " Where B.����ID=[1] And A.��ҳID=[2] And A.����ID=B.����ID"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ժʱ��", lng����ID, lng��ҳID)
    str��Ժ��� = rsTemp!��Ժ���
    'lng��ҳID = rsTemp!��ҳID
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyymmdd")
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str����ʱ�� = str����ʱ��
    str������� = Mid(str����ʱ��, 1, 4)
    strסԺ�� = Nvl(rsTemp!סԺ��)
    
    '��ȷ�������� ���˺�2004/06/15,��Ϊ���ֽ���ʱ��������
    If rsTemp.EOF Then
        strTmp = Get��Ժ���(lng����ID, lng��ҳID, , True)
    Else
        If IsNull(rsTemp!��Ժ����) Then
            strTmp = Get��Ժ���(lng����ID, lng��ҳID, , True)
        Else
            strTmp = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, , True)
        End If
    End If
    If InStr(1, strTmp, "|") <> 0 Then
        g�������_����.��ϱ��� = Split(strTmp, "|")(1)
        g�������_����.������� = Split(strTmp, "|")(0)
    End If
    
    
    strTmp = ""

    '���»�ȡ��¼
    Set rs���� = GetסԺ�����¼(lng����ID, lng��ҳID, intinsure)
    If rs����.RecordCount <= 0 Then
        ShowMsgbox "����Ŀδ����ҽ����Ŀ�����ܽ���!"
        Exit Function
    End If
    dbl�𸶱�׼ = g�������_����.����
    

    With rs����
        '�ϴ�������ϸ
        curTotal = 0
        Do While Not .EOF
        
            If strҽ�� = "" Then
                strҽ�� = Nvl(!ҽ�����)
                If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                    strҽ�� = Substr(strҽ��, 1, 6)
                End If
            End If
            curTotal = curTotal + Nvl(!���, 0)
            
            lng����ID = Nvl(!����ID, 0)
            strTmp = Nvl(!����ֵ)
            
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '-----------------------------------------------------------------------
                '���㱣��
                '---�����㷨=2(�������)����Ŀ,���ڲ��ǰ��ձ������м���,�����ҪԤ������dbl����=1,
                '---ʹ����뱨���ָ���ж�,�������Ϊ��ʼֵ=0���±���Ϊȫ�Է���Ŀ�����ָ�
                
                dbl���� = Nvl(!סԺ�ȶ�, 0) / 100

                
                '---��˳��,�˴�������������Ч����,��������
                '-------------------------------------------------------------------------
                '����Ϊ:A��ְ��B���ݡ�L���ݡ�T����,����Ĭ��Ϊ1��ְ��2���ݡ�3���ݡ�4����
                If g�������_����.ҽ������ <> 2 And g�������_����.ְ����ҽ��� = "L" And g�������_����.�α����3 = "0" And Nvl(!������Ŀ��, 0) = 1 Then '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '  ������  ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�������ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dbl���� = 1
                End If
                '-------------------------------------------------------------------------
                
                If Nvl(!ҽ����Ŀ����) = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If Nvl(!ҽ����Ŀ����) = "���" Then
                    strTmp = "����"
                End If
                
                '---����һ�����������ƷѺʹ����ñ��������ļ��
                If (strTmp = "�������Ʒ�" Or strTmp = "����") And dbl���� = 0 Then
                    MsgBox "�������Ʒ��û�����õı�������Ϊ0,���ȼ�������Ƿ���ȷ"
                    Exit Function
                End If
                
                
                If g�������_����.ҽ������ <> 2 And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dbl���� = Val(Nvl(!ҽ����Ŀ����)) / 100
                End If
                
                '����Ǵ�λ,���谴���·�ʽ����,��������ͳ���������,������Ϊ100���Է�,���ֿ������ʹ�����
                If Nvl(!�շ����) = "J" Then
                    
                    gstrSQL = "" & _
                        "   Select ���Ӵ�λ From ���˱䶯��¼ " & _
                        "   Where ����=" & Nvl(!����, 0) & " and ����id=" & lng����ID & _
                        "         And ( (to_date('" & Format(!����ʱ��, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600 between ��ʼʱ�� and ��ֹʱ��) or" & _
                        "               ( ��ֹʱ�� is null  and ��ʼʱ��<=to_date('" & Format(!����ʱ��, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600)) " & _
                        "         And ���� is not null"
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ���Ƿ�Ϊ����!"
                    If rsTemp.RecordCount >= 1 Then
                    '--������Ų��������䶯��¼�еĴ�(�ֹ��Ƿ�),�򱾴�ѡ�����ΪNOTHING,��Ҫ��dbl����=0
                    '--����м�¼,���ж��Ƿ�Ϊ���Ӵ�λ,�ǵĻ�dbl����=0
                        If rsTemp!���Ӵ�λ = 1 Then
                            '��ʾ������λ,Ϊȫ�Է�
                            dbl���� = 0
                        Else
                            If !�㷨 = 2 Then
                                 dbl���� = 1
                            End If
                        End If
                    End If
                End If
                
                '--Ӥ����ֱ�ӹ���ȫ�ԷѲ���
                If !Ӥ���� <> 0 Then
                    dbl���� = 0
                End If
                
                
            End If
'------------------------------------------------------�������-----------------------------------------------------
'---��Ϊֻ�д�λ���ô����޶��,����ֻ�Դ�λ�����޶��жϲ�����
'---�Ƿ�����ȷ�ĸ��Ӵ�λ,��Ҫ���ݱ䶯��¼�����ж� ����\ʱ��,����ȷ���Ǳ��˵��Զ����㴲λ��
'---(���Ƕ��ֹ�¼��Ĵ�λ��,�ж��������б仯)
'-------------------------------------------------------------------------------------------------------------------
                If dbl���� <> 0 Then
                    Select Case strTmp
                        Case "����"
                                
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl���� = dbl���� + Round(Nvl(!���, 0) * dbl����, 5)
                                
                                If intinsure = TYPE_������ Then
                                
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                End If

                        Case "��ҩ��"
                                    
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���, 0) * dbl����, 5)
                                
                                If intinsure = TYPE_������ Then
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                End If
                                
                        Case "��ҩ��"
                                
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���, 0) * dbl����, 5)

                                If intinsure = TYPE_������ Then
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                End If

                        Case "��ҩ��"
                                
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���, 0) * dbl����, 5)

                                If intinsure = TYPE_������ Then
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                End If

                        Case "����"
                                
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl���� = dbl���� + Round(Nvl(!���, 0) * dbl����, 5)

                                If intinsure = TYPE_������ Then
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                End If

                        Case "���Ʒ�"
                                '---���λ����������������Ϊ���Ʒ��ô���,������Ʒ�ר�����㷨�ж�
                                
                                If Nvl(!�㷨, 0) = 2 Then
                                    '���ڿ��ܴ����������ʺ��ٽ���,���Ƿ񳬹��޶���밴�վ���ֵ�����ж�
                                    If Nvl(!��׼����, 0) <= Nvl(Abs(!�۸�), 0) Then
                                    
                                        '����޶�<����,������������Ϊ�޶�*����
                                        dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!��׼����, 0), 5) * Nvl(Abs(!����), 0) * Sgn(!���)
                                        
                                        If intinsure = TYPE_������ Then
                                            '���޲���=(����-�޶�)*����,  �����ɽ��ķ��ž���,�����м���dbl�������Էѷ���
'                                            dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(Abs(!�۸�) - !��׼����, 0) * Abs(!����) * Sgn(!���), 5)
                                            If g�������_����.ְ����ҽ��� = "Q" Then
                                                '������ҵ������Գ��޲�����Ҫ�Էѣ������������
                                                '�ܺ�ȫ 2005-02-17
                                                dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(Nvl(Abs(!�۸�) - !��׼����, 0) * Abs(!����) * Sgn(!���), 5)
                                            Else
                                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(Abs(!�۸�) - !��׼����, 0) * Abs(!����) * Sgn(!���), 5)
                                            End If
                                        Else
                                            '���޲���=(����-�޶�)*����,  �����ɽ��ķ��ž���,����������dbl������
                                            'dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(Abs(!�۸�) - !��׼����, 0) * Abs(!����) * Sgn(!���), 5)
                                            
                                            '2005-08-02����������
                                            dbl������ = dbl������ + Round(Nvl(Abs(!�۸�) - !��׼����, 0) * Abs(!����) * Sgn(!���), 5)
                                        End If
                                        
                                    Else
                                        '����޶�>=����,��ȫ���������Ʒ���
                                        dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!���, 0), 5)
                                    End If
                                Else

                                    '--�۳��ԷѲ��ַ��뱾��
                                    dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!���, 0) * dbl����, 5)
    
                                    If intinsure = TYPE_������ Then
                                        '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                    Else
                                        '---��������졢������������ԷѲ��ּ���dbl������
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                                    End If
                                End If

                        Case "����"
                                
                                '---�����кͿ������ڴ������ϵĴ�����ȫһ��
                                '---�����Ŀ���ı������ּ������
                                dbl���� = dbl���� + Round(Nvl(!���, 0) * dbl����, 5)
                                '---�����Ŀ�����ԷѲ��ּ������Է�
                                dbl����Է� = dbl����Է� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                        Case "Ѫ��"
                            '�����кͿ������Դ����ô���ͬ,
                            '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                      
                            dblѪ�� = dblѪ�� + Round(Nvl(!��� * dbl����, 0), 5)
                            dblѪ���Է� = dblѪ���Է� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                        Case "�������Ʒ�"
                                '---�����кͿ��������������Ʒ��ϵĴ���������ͬ
                                '2004/9/11��ǰ:�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                                '2004/9/11�Ժ�:�������뿪�������㷽ʽһ�£�����ͳ�ﲿ�ֵĽ��
                                dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!���, 0) * dbl����, 5)
                                dbl���������Է� = dbl���������Է� + Round(Nvl(!���, 0) * (1 - dbl����), 5)
                    End Select
                Else
                
                    'ȫ���Ǳ���Ϊ0����Ŀ(��������),�ֱ�Դ������ǿ����������жϷ��ڲ�ͬ���ֶ�
                    If intinsure = TYPE_������ Then
                        '�����з���dbl�Ǳ��շ���
                        dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(!���, 5)
                    Else
                        '����������dbl������
                        dbl������ = dbl������ + Round(!���, 5)
                    End If
                    
                End If
            
            .MoveNext
        Loop

        If ����ҽ������(intinsure, 1, lng����ID, lng��ҳID, 0, 0, False, True, dbl�𸶱�׼, dbl����, dbl��ҩ��, dbl��ҩ��, dbl��ҩ��, _
            dbl����, dbl���Ʒ�, dblѪ��, dblѪ���Է�, dbl����, dbl����Է�, dbl�������Ʒ�, dbl���������Է�, dbl�������Էѷ���, _
            dbl�Ǳ��շ���, dbl������, dblҩ���Է�, curTotal, strҽ��, str���㷽ʽ, strסԺ��, str��Ժ����) = False Then
            Exit Function
        End If
        g�������_����.֧����� = curTotal
    End With
    
    סԺ�������_���� = str���㷽ʽ
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function Get�����ʻ����_����(ByVal intinsure As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�����ʻ����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    'ҽԺ����    CHAR    1   4       Ժ��
    '���˱��    CHAR    5   8       Ժ��
    '��������    CHAR    13  16  ĿǰΪ: WZMB    Ժ��
    '�������    NUM 29  4       ����
    '�����ʻ�ԭʼֵ  NUM 33  10  ÿ�β������ۼ�ֵ    ����
    '�����ʻ���ǰֵ  NUM 43  10      ����
    '�ʻ�״̬    CHAR    53  1   A������Cֹ��    ����

    
    Dim strTmp As String
    Err = 0
    On Error GoTo errHand:
    With g�������_����
        strTmp = Lpad(gstrҽԺ����_����, 4)      'ҽԺ����    CHAR    1   4       Ժ��
        strTmp = strTmp & Lpad(.���˱��, 8) '���˱��    CHAR    5   8       Ժ��
        strTmp = strTmp & Lpad("WZMB", 16)  '��������    CHAR    13  16  ĿǰΪ: WZMB    Ժ��
        strTmp = strTmp & Space(4)  '�������    NUM 29  4       ����
        strTmp = strTmp & Space(10)  '�����ʻ�ԭʼֵ  NUM 33  10  ÿ�β������ۼ�ֵ    ����
        strTmp = strTmp & Space(10)  '�����ʻ���ǰֵ  NUM 43  10      ����
        strTmp = strTmp & Space(1)   '�ʻ�״̬    CHAR    53  1   A������Cֹ��    ����
        '��ҽ�����Ĳ�ѯ����
        '   1007    2   55  �����ʻ���ѯ
        Get�����ʻ����_���� = ҵ������_����(.ҽ������, 1007, strTmp, intinsure)
        If Get�����ʻ����_���� = False Then
            .�����ʻ�ԭʼֵ = 0
            .�����ʻ���ǰֵ = 0
            Exit Function
        End If
        .�����ʻ�ԭʼֵ = Val(Substr(strTmp, 33, 10))
        .�����ʻ���ǰֵ = Val(Substr(strTmp, 43, 10))
    End With
    
    Exit Function
errHand:
End Function

Private Function סԺ���㼰����_����(ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal ԭ����id As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean

    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strTmp As String, strסԺ�� As String, str��Ŀͳ�Ʒ��� As String, strInsertSQL As String
    Dim strInfor As String  '�������ķ��ش�
    Dim curTotal As Double
    Dim dbl���� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double, dbl��ҩ�� As Double
    
    '2005-08-02����������
    Dim dblҩ���Է� As Double
    
    Dim dbl���� As Double, dbl���Ʒ� As Double, dbl���� As Double, dbl����Է� As Double
    Dim dbl�������Ʒ� As Double, dbl���������Է� As Double, dbl�������Էѷ��� As Double, dbl�Ǳ��շ��� As Double
    Dim dblѪ�� As Double, dblѪ���Է� As Double
    Dim dbl������ As Double     '��Դ�����������
    Dim dbl���� As Double, dbl�𸶱�׼ As Double
    
    Dim strҽ�� As String, str��ϸ As String, str���ұ��� As String, str��Ժ���� As String
    Dim intҵ�� As Integer
    
    intҵ�� = IIf(bln����, 1, 0)
    
    Err = 0
    On Error GoTo errHand:
    
    'סԺӦ�ñ���֧�������е�סԺ�ȶ�
    '4-26,��˳������Ӥ�����ж�,�ų�Ӥ���Ѳ������籣����  '--���ұ���=����+����
    gstrSQL = " " & _
        "        select a.ʵ�ս��,a.id,a.��¼����,a.��ҳid,a.��¼״̬,a.����ʱ��,a.�Ǽ�ʱ��,a.no,a.���˲���id,a.����,a.���,a.��ʶ�� as סԺ��,a.���˿���id,a.����id,a.�շ����,b.���,a.���㵥λ, " & _
        "               A.���㵥λ,A.����*Nvl(A.����,1) ����,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ�� ,a.������ as ҽ��,c.��� as ҽ�����, " & _
        "               a.ҽ�����,nvl(a.Ӥ����,0) as Ӥ����, A.ʵ�ս��,nvl(A.�Ƿ��ϴ�,0) as �Ƿ��ϴ�, " & _
        "               F.����ֵ,D.���� as ��Ŀ����,D.���� as ��Ŀ����,Nvl(J.��ʶ��,Nvl(D.��ʶ����||D.��ʶ����,D.����)) as ���ұ���, " & _
        "               E.��Ŀ���� as ҽ������,E.��Ŀ���� as ҽ������,e.�Ƿ�ҽ��,e.����id,G.סԺ�ȶ� as ͳ��ȶ�,G.��׼����,G.�㷨,H.���� as ��������,J.���� as ����, " & _
        "               L.����,l.���� , l.����, l.ҽ����, l.��Ա���, l.��λ����, l.˳���, l.����֤��, l.�ʻ����, l.��ǰ״̬, l.����ID, l.��ְ, l.�����, l.�Ҷȼ�, l.����ʱ�� " & _
        "        from סԺ���ü�¼ a,�շ���� b,��Ա�� c,�շ�ϸĿ D,����֧����Ŀ E,����֧������ G,�����ʻ� L,���ű� H, " & _
        "             (Select U.*,K.����ֵ From �շ���� U,���ղ��� K where U.���=K.������ and K.����=" & intinsure & "  ) F ," & _
        "             (Select distinct Q.ҩƷid,Q.��ʶ��,T.���� From ҩƷĿ¼ Q,ҩƷ��Ϣ R,ҩƷ���� T  Where  Q.ҩ��id=R.ҩ��id and R.����=T.���� ) J " & _
        "        where a.��¼״̬<>0 and  a.�շ����=b.���� and a.�շ�ϸĿid=J.ҩƷid(+)   and  Nvl(a.���ӱ�־,0)<>9 and a.�շ�ϸĿid=D.id and a.������=c.����(+) and a.�շ����=F.����(+) and " & _
        "              a.�շ�ϸĿid=E.�շ�ϸĿID and E.����id=G.id and a.����id=L.����ID and a.��������id=h.id  and " & _
        "              a.����ID = " & lng����ID & " And a.����ID = " & lng����ID & " And E.���� = " & intinsure
        
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡסԺ������ϸ"
    
    
    'ȷ���ò����Ƿ��Ѿ���Ժ
    gstrSQL = "Select ����ID,��ҳID,��Ժ���� From ������ҳ where ����id=" & lng����ID & " and ��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�����Ƿ��Ժ"
    str��Ժ���� = ""
    If rsTemp.EOF Then
        strTmp = Get��Ժ���(lng����ID, lng��ҳID, False, True)
    Else
        If IsNull(rsTemp!��Ժ����) Then
            strTmp = Get��Ժ���(lng����ID, lng��ҳID, False, True)
        Else
            strTmp = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, , True)
            str��Ժ���� = Format(rsTemp!��Ժ����, "yyyymmdd")
        End If
    End If
    
    If InStr(1, strTmp, "|") <> 0 Then
        g�������_����.��ϱ��� = Split(strTmp, "|")(1)
        g�������_����.������� = Split(strTmp, "|")(0)
    End If
    
    With rs��ϸ
        If Not .EOF Then
            strסԺ�� = Nvl(!סԺ��)
        End If
        Do While Not .EOF
            strTmp = Nvl(!����ֵ)
            lng����ID = Nvl(!����ID, 0)
            If strҽ�� = "" Then
                strҽ�� = Nvl(!ҽ�����)
                If LenB(StrConv(strҽ��, vbFromUnicode)) > 6 Then
                    strҽ�� = Substr(strҽ��, 1, 6)
                End If
            End If
            'ȷ���������
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str��Ŀͳ�Ʒ��� = ""
                Else
                    str��Ŀͳ�Ʒ��� = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                
                '����
                 dbl���� = Nvl(!ͳ��ȶ�, 0) / 100
                
                 '---��˳��,�˴�������������Ч����,��������
                '-------------------------------------------------------------------------
                If Nvl(!����, 0) <> TYPE_���������� And Val(Nvl(!��λ����, "99")) = 0 And Nvl(!��ְ, 0) = 3 And Nvl(!�Ƿ�ҽ��, 0) = 1 Then '���󱣺�������Ա����ҽ����Ŀ
                    '��λ����洢���ǲα����3   CHAR    90  1   0 �󱣡�1 �±�
                    '    ��ҵ��λ����ҽ��������ȫִ��ҽ�����ߣ�������ͨҽ��20%��10%�ԷѲ��ֲ�����ҽ�����ֽ�֧���������ಡ�������ԷѲ��ּ���ҽ������ӡҽ���վݣ�ֻ��100%�Է����Ը��ֽ𣬿��ֽ�Ʊ������дʵ�֣�ע��: ���ֲ������ڲ��ҽԺ��λ
                    dbl���� = 1
                End If
                '--------------------------------------------------------------------------
                
                If Nvl(!ҽ������) = "����" Then
                    strTmp = "�������Ʒ�"
                End If
                If Nvl(!ҽ������) = "���" Then
                    strTmp = "����"
                End If
                If Nvl(!����, 0) = TYPE_������ And (g�������_����.ְ����ҽ��� = "L" Or _
                     g�������_����.ְ����ҽ��� = "T") Then
                    '�����L���ݺ�T����ľͰ���ҵ��������
                    dbl���� = Val(Nvl(!ҽ������)) / 100
                End If
                
                 '------------------------------�˴����ε�---------------------------------------------
                '����Ǵ�λ,���谴���·�ʽ����,��������ͳ���������,������Ϊ100���Է�,���ֿ������ʹ�����
                If Nvl(!�շ����) = "J" Then
                    
                    gstrSQL = "" & _
                        "   Select ���Ӵ�λ From ���˱䶯��¼ " & _
                        "   Where ����=" & Nvl(!����, 0) & " and ����id=" & lng����ID & _
                        "         And ( (to_date('" & Format(!����ʱ��, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600 between ��ʼʱ�� and ��ֹʱ��) or" & _
                        "               ( ��ֹʱ�� is null  and ��ʼʱ��<=to_date('" & Format(!����ʱ��, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')+1-1/24/3600)) " & _
                        "         And ���� is not null"
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ���Ƿ�Ϊ����!"
                    If rsTemp.RecordCount >= 1 Then
                       If rsTemp!���Ӵ�λ = 1 Then
                            '��ʾ������λ,Ϊȫ�Է�
                            dbl���� = 0
                       Else
                            If !�㷨 = 2 Then
                            dbl���� = 1
                            End If
                       End If
                    End If
                End If
                
                '--���Ӷ�Ӥ���ѵ��ж�,Ӥ���������κ�����¶����豨��,ȫ�������ԷѲ���
                If !Ӥ���� <> 0 Then
                    dbl���� = 0
                End If
                
            '-----------------------------------------------�㷨����----------------------------------------------------
            '---���սӿ��ĵ����з��÷ָ�
            '-----------------------------------------------------------------------------------------------------------
                If dbl���� <> 0 Then
                    
                    Select Case strTmp
                        Case "����"
                                
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl���� = dbl���� + Round(Nvl(!���ʽ��, 0) * dbl����, 5)
                                
                                If intinsure = TYPE_������ Then
                                
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                End If

                        Case "��ҩ��"
                                    
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���ʽ��, 0) * dbl����, 5)
                                
                                If intinsure = TYPE_������ Then
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                End If
                                
                        Case "��ҩ��"
                                
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���ʽ��, 0) * dbl����, 5)

                                If intinsure = TYPE_������ Then
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                End If

                        Case "��ҩ��"
                                
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl��ҩ�� = dbl��ҩ�� + Round(Nvl(!���ʽ��, 0) * dbl����, 5)

                                If intinsure = TYPE_������ Then
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dblҩ���Է� = dblҩ���Է� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                End If

                        Case "����"
                                
                                '--�۳��ԷѲ��ַ��뱾��
                                dbl���� = dbl���� + Round(Nvl(!���ʽ��, 0) * dbl����, 5)

                                If intinsure = TYPE_������ Then
                                    '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                Else
                                    '---��������졢������������ԷѲ��ּ���dbl������
                                    dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                End If

                        Case "���Ʒ�"
                                '---���λ����������������Ϊ���Ʒ��ô���,������Ʒ�ר�����㷨�ж�
                                
                                If Nvl(!�㷨, 0) = 2 Then
                                    '���ڿ��ܴ����������ʺ��ٽ���,���Ƿ񳬹��޶���밴�վ���ֵ�����ж�
                                    If Nvl(!��׼����, 0) <= Nvl(Abs(!ʵ�ʼ۸�), 0) Then
                                    
                                        '����޶�<����,������������Ϊ�޶�*����
                                        dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!��׼����, 0), 5) * Nvl(Abs(!����), 0) * Sgn(!���ʽ��)
                                        
                                        If intinsure = TYPE_������ Then
                                            '���޲���=(����-�޶�)*����,  �����ɽ��ķ��ž���,�����м���dbl�������Էѷ���
                                            If g�������_����.ְ����ҽ��� = "Q" Then
                                                '������ҵ������Գ��޲�����Ҫ�Էѣ������������
                                                '�ܺ�ȫ 2005-02-17
                                                dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(Nvl(Abs(!ʵ�ʼ۸�) - !��׼����, 0) * Abs(!����) * Sgn(!���ʽ��), 5)
                                            Else
                                                dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(Abs(!ʵ�ʼ۸�) - !��׼����, 0) * Abs(!����) * Sgn(!���ʽ��), 5)
                                            End If
                                        Else
                                            '���޲���=(����-�޶�)*����,  �����ɽ��ķ��ž���,����������dbl������
                                            dbl������ = dbl������ + Round(Nvl(Abs(!ʵ�ʼ۸�) - !��׼����, 0) * Abs(!����) * Sgn(!���ʽ��), 5)
                                        End If
                                        
                                    Else
                                        '����޶�>=����,��ȫ���������Ʒ���
                                        dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!���ʽ��, 0), 5)
                                    End If
                                Else

                                    '--�۳��ԷѲ��ַ��뱾��
                                    dbl���Ʒ� = dbl���Ʒ� + Round(Nvl(!���ʽ��, 0) * dbl����, 5)
    
                                    If intinsure = TYPE_������ Then
                                        '---�����д�졢������������ԷѲ��ּ���dbl�������Էѷ���
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                    Else
                                        '---��������졢������������ԷѲ��ּ���dbl������
                                        dbl�������Էѷ��� = dbl�������Էѷ��� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                                    End If
                                End If

                        Case "����"
                                
                                '---�����кͿ������ڴ������ϵĴ�����ȫһ��
                                '---�����Ŀ���ı������ּ������
                                dbl���� = dbl���� + Round(Nvl(!���ʽ��, 0) * dbl����, 5)
                                '---�����Ŀ�����ԷѲ��ּ������Է�
                                dbl����Է� = dbl����Է� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                        Case "Ѫ��"
                            '�����кͿ������Դ����ô���ͬ,
                            '������Ϊ�۳������Ŀ���۳�����ԷѵĽ��,���е����ݲ��˵Ĵ���Է�ȫ�����������Է�
                                      
                            dblѪ�� = dblѪ�� + Round(Nvl(!���ʽ�� * dbl����, 0), 5)
                            dblѪ���Է� = dblѪ���Է� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                        Case "�������Ʒ�"
                                '---�����кͿ��������������Ʒ��ϵĴ���������ͬ
                                '2004/9/11��ǰ:�������뿪�������㷽ʽ��һ�£����������ܶ����������ͳ�ﲿ��
                                '2004/9/11�Ժ�:�������뿪�������㷽ʽһ�£�����ͳ�ﲿ�ֵĽ��
                                dbl�������Ʒ� = dbl�������Ʒ� + Round(Nvl(!���ʽ��, 0) * dbl����, 5)
                                dbl���������Է� = dbl���������Է� + Round(Nvl(!���ʽ��, 0) * (1 - dbl����), 5)
                    End Select
                Else
                
                    'ȫ���Ǳ���Ϊ0����Ŀ(��������),�ֱ�Դ������ǿ����������жϷ��ڲ�ͬ���ֶ�
                    If intinsure = TYPE_������ Then
                        '�����з���dbl�Ǳ��շ���
                        dbl�Ǳ��շ��� = dbl�Ǳ��շ��� + Round(!���ʽ��, 5)
                    Else
                        '����������dbl������
                        dbl������ = dbl������ + Round(!���ʽ��, 5)
                    End If
                    
                End If

            Else
                dbl���� = 1
                str��Ŀͳ�Ʒ��� = ""
            End If

 
            '�ϴ���ϸ��¼,ʵʱҽ����ϸ����
            If gblnסԺ��ϸʱʵ�ϴ� And bln���� = False And Nvl(!�Ƿ��ϴ�, 0) = 0 And Nvl(!���ʽ��, 0) <> 0 And Nvl(!ʵ�ս��, 0) <> 0 Then
                    If Nvl(!����, 0) = TYPE_���������� Then '������
                        str��ϸ = Lpad(gstrҽԺ����_����, 6)     'ҽԺ����    CHAR    1   6       Ժ����д
                        str��ϸ = str��ϸ & Lpad(Nvl(!ҽ����), 10)  '���ձ��    CHAR    7   10      Ժ����д
                    Else
                        str��ϸ = Lpad(gstrҽԺ����_����, 4)     'ҽԺ����    CHAR    1   4       Ժ��
                        str��ϸ = str��ϸ & Lpad(Nvl(!ҽ����), 8)   '���˱��    CHAR    5   8       Ժ��
                    End If
                    
                    str��ϸ = str��ϸ & Lpad(Nvl(!סԺ��, 0), 10) '��־��  CHAR    13  10  ������ϸ�Կո�λ,סԺ��סԺ��  Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!˳���, 0), 4)   '�������    NUM 23  4   סԺ��ϸ�����������Ժ�Ǽ�ʱ�������������ϸ:                         ������ڱ��ν���������� Ժ��
                    
                    'Modified By ���� 2004-07-29 ԭ�򣺴���NO��
                    str��ϸ = str��ϸ & Lpad(Mid(Nvl(!NO, "00000000"), 2, 7), 10)     '������  NUM 27  10      Ժ��

                    str��ϸ = str��ϸ & Lpad(Nvl(!ID, 0), 10)      '������Ŀ���    NUM 37  10  ��Ӧ�����ŵļǼ���Ŀ���    Ժ��
                    
                    '������Ϊ���ݺ�  CHAR    41  10  ҽ���ţ�    Ժ����д
                    str��ϸ = str��ϸ & Lpad(Nvl(!ҽ�����, " "), 10)     'ҽ����  CHAR    47  10  ������Ӧҽ����ҽ����¼�ţ�������ϸ��û��ҽ����ҽԺ�Կո�λ    Ժ��
                    g�������_����.������� = Nvl(!�Ҷȼ�, 0)
                    
                    str��ϸ = str��ϸ & Get�������(intҵ��, Nvl(!�Ҷȼ�, 0))         '�������    CHAR    57  1   ȡֵ���"�������"˵��  Ժ��
                    str��ϸ = str��ϸ & Rpad(Format(!�Ǽ�ʱ��, "yyyymmddHHmmss"), 16)      '��������ʱ�䣨Ͷҩʱ�䣩    DATETIME    58  16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ    Ժ��
                    
                    str��ϸ = str��ϸ & Lpad(Nvl(!���ұ���), 20)      '��Ŀ����    CHAR    74  20  �Ƽ���Ŀ����    Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!��Ŀ����), 20)      '��Ŀ����    CHAR    94  20      Ժ��
        
                    If !�Ƿ�ҽ�� = 1 Then
                        str��ϸ = str��ϸ & Lpad(1 - dbl����, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                    Else
                        str��ϸ = str��ϸ & Lpad(1, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(str��Ŀͳ�Ʒ���, 1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                    
                    '2005-08-02����������
                    If intinsure = TYPE_���������� Then
                        str��ϸ = str��ϸ & Lpad(Abs(Nvl(!����)) * Sgn(!���ʽ��), 10) '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                        str��ϸ = str��ϸ & Lpad(Abs(Nvl(!ʵ�ʼ۸�)), 10) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
                    Else
                        str��ϸ = str��ϸ & Lpad(Abs(Nvl(!����)) * Sgn(!���ʽ��), 6) '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                        str��ϸ = str��ϸ & Lpad(Abs(Nvl(!ʵ�ʼ۸�)), 8) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
                    End If
                    str��ϸ = str��ϸ & Lpad(Nvl(!���㵥λ), 4) '��λ    CHAR    135 4       Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!����), 20)      '����    CHAR    139 20  �����Ƭ����    Ժ��
                    str��ϸ = str��ϸ & Lpad(Nvl(!ҽ��), 8)      'ҽʦ����    CHAR    159 8       Ժ��
                    str��ϸ = str��ϸ & Lpad(g�������_����.��ϱ���, 16)      '��ϱ���    CHAR    167 16      Ժ��
                    str��ϸ = str��ϸ & Lpad(Substr(g�������_����.�������, 1, 28), 30)   '�������    CHAR    183 30      Ժ��
                    str��ϸ = str��ϸ & Space(16)     '����ʱ��    DATETIME    213 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ��Ժ�˿ո�λ  ����
                 
                    '�ϴ���ϸ
                    '1003    7   230 ʵʱҽ����ϸ�����ύ
                    סԺ���㼰����_���� = ҵ������_����(IIf(Nvl(!����, 0) = TYPE_����������, 2, 1), 1003, str��ϸ, intinsure)
                    If סԺ���㼰����_���� = False Then
                        ShowMsgbox "סԺ����������ϸ�����ύʧ��,���ܼ���!"
                        Exit Function
                    End If
                    
                    '�ϴ�ҽ����ϸ
                    If Nvl(!ҽ�����, 0) <> 0 Then
                        If ҽ����ϸ�����ύ(!ҽ�����, Nvl(!סԺ��), str��Ŀͳ�Ʒ���, intinsure) = False Then
                            ShowMsgbox "ҽ����ϸ�����ύʧ��,���ܼ���!"
                            Exit Function
                        End If
                    End If

                    'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                    'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                    zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            End If
            '�����ܶ�,����
            curTotal = curTotal + Round(Nvl(!���ʽ��, 0), 5)
            .MoveNext
        Loop
    End With
    
 
    '��д�����¼
    '��������
    dbl�𸶱�׼ = g�������_����.����
        
    If ����ҽ������(intinsure, 1, lng����ID, lng��ҳID, lng����ID, ԭ����id, bln����, False, dbl�𸶱�׼, dbl����, dbl��ҩ��, dbl��ҩ��, dbl��ҩ��, _
        dbl����, dbl���Ʒ�, dblѪ��, dblѪ���Է�, dbl����, dbl����Է�, dbl�������Ʒ�, dbl���������Է�, dbl�������Էѷ���, dbl�Ǳ��շ���, _
        dbl������, dblҩ���Է�, curTotal, strҽ��, strInfor, strסԺ��, str��Ժ����) = False Then
        Exit Function
    End If
    '��������ν���,�Ͳ��������
    If g�������_����.���ν��� Then
    Else
        If bln���� Then
            gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
            zlDatabase.ExecuteProcedure gstrSQL, "ҽ���ʻ���Ժ"
        Else
            gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
            zlDatabase.ExecuteProcedure gstrSQL, "ҽ���ʻ���Ժ"
        End If
    End If
    סԺ���㼰����_���� = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean

    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
    'Dim cur�����ʻ� As Currency
    Dim lng��ҳID As Long
    Dim int��ҳid As Long
    Dim blnError As Boolean
    'Dim str��Ժ��� As String, str������� As String
    'Dim str����ʱ�� As String, str����ʱ�� As String
    'Dim str������ As String
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    
    '=================================================================================
    '�������ҽ������
    '���ݴ���Ľ���id�жϱ��ν����Ƿ�������㣬��������ϱ�������ҳid,����Ϊ��סԺ����
    '��������������ӿڶԲ��˷��ü�¼����
'    gstrSQL = "select nvl(��ҳid,0) as ��ҳid from ����Ԥ����¼ " & _
'            "where mod(��¼����,10)=2 and ����id=" & lng����ID & " and ����id=" & lng����ID
    gstrSQL = "select Nvl(��ҳid,0) as ��ҳid from סԺ���ü�¼ " & _
            "where rownum=1 and ����id=[1] and ����id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鲡���Ƿ�������ʽ���", lng����ID, lng����ID)
    
    lng��ҳID = Nvl(rsTemp!��ҳID, 0)
    
    Do While Not rsTemp.EOF
        int��ҳid = int��ҳid + rsTemp!��ҳID
        rsTemp.MoveNext
    Loop
    
    '����Ƿ�����ҳid<>0�����,���Ƿ���סԺ����,
    '����������������ʽ���,ֱ�Ӱ���סԺ��ʽ���㡣����ִ���������
    If int��ҳid = 0 Then
        If ������ʽ��㼰����_����(False, lng����ID, lng����ID, lng����ID, 0, intinsure) = True Then
            '�������ɹ�����HIS���سɹ���־
            סԺ����_���� = True
            Exit Function
        Else
            סԺ����_���� = False
            Exit Function
        End If
    End If
    
    '=================================================================================
    '20041021��
    If g�������_����.���ν��� Then
    Else
        Call �������_����(lng����ID, intinsure)
    End If
    סԺ����_���� = סԺ���㼰����_����(False, lng����ID, lng����ID, lng����ID, lng��ҳID, intinsure)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function Get����(ByVal strְ����ҽ��� As String, ByVal lng���� As Long, ByVal intinsure As Integer) As Double
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ����
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------

       Dim strCaption As String
       Dim rsTmp As New ADODB.Recordset
       
       '20040911�������ϲ���
       strCaption = Decode(strְ����ҽ���, "A", "��ְ", "B", "����", "L", "����", "T", "����", "Q", "��ҵ����", "E", "����", "��ְ")
    
        gstrSQL = "" & _
            "   Select d.���*a.����/100 as ����" & _
            "   From ����֧������ a,������Ⱥ b, " & _
            "      (Select * From ���������  " & _
            "       where ((" & lng���� & ">=���� and " & lng���� & "<=����) or (" & lng���� & ">���� and ����=0) ) and ����=" & intinsure & _
            "       ) c,����֧���޶� d " & _
            " where a.����=" & intinsure & " and b.���� =a.���� and a.��ְ=b.��� and b.����='" & strCaption & "' and  " & _
            "       a.�����=c.����� and a.��ְ=c.��ְ and a.����=d.���� and d.���='" & Format(zlDatabase.Currentdate, "yyyy") & "' and d.����='1'"
    
       Err = 0
       On Error GoTo errHand:
       zlDatabase.OpenRecordset rsTmp, gstrSQL, "��������"
       If Not rsTmp.EOF Then
            Get���� = Nvl(rsTmp!����, 0)
       Else
            Get���� = 0
       End If
       Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
       Get���� = 0
   
End Function

Public Function סԺ�������_����(lng����ID As Long, ByVal intinsure As Integer) As Boolean
    Dim lng����ID As Long
    Dim str�˵���� As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    On Error GoTo errHand
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID

    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "Select ��¼ID,����ID,��ҳID,֧��˳���,��ע " & _
              "  From ���ս����¼ Where ����=2 and ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ��ս����¼���޸ý����¼!"
        Exit Function
    End If
         
    lng����ID = Nvl(rsTemp!����ID, 0)
    lng��ҳID = Nvl(rsTemp!��ҳID, 0)
        
    '��������¼����ҳid��дΪ��,���ʾ��������ʽ����¼,��Ҫ�������﷽ʽ���г���
    If lng��ҳID = 0 Then
        If ������ʽ��㼰����_����(True, lng����ID, lng����ID, lng����ID, 0, intinsure) Then
            סԺ�������_���� = True
            Exit Function
        Else
            סԺ�������_���� = False
            Exit Function
        End If
    End If
    '------------------------------������ʽ���������
        
    '���¶���
    If ��ȡ�������_����(IIf(intinsure = TYPE_����������, 2, 1), intinsure) = False Then
        Exit Function
    End If
    
    Dim strArr
    strArr = Split(Nvl(rsTemp!֧��˳���), ";")
    
    '�������;ת�ﵥ��;��ϱ���
    '5-��ͨסԺ("2", "D"),6-��ͥ����סԺ("4", "C")
    '7-��������סԺ("O", "P"),8-���˱���סԺ("Q", "R")
    With rsTemp
        If UBound(strArr) >= 2 Then
            g�������_����.������� = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
            g�������_����.ת�ﵥ�� = strArr(1)
            g�������_����.��ϱ��� = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g�������_����.������� = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
            g�������_����.ת�ﵥ�� = strArr(1)
        Else
            g�������_����.������� = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
        End If
        g�������_����.������� = Nvl(rsTemp!��ע)
    End With
    
    '��֤�Ƿ�Ϊ�ò��˵�IC��
    gstrSQL = "Select ����,����ID,���� From �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵�ҽ����"
    If rsTemp.EOF Then
        ShowMsgbox "�ò����ڱ����ʻ����޼�¼!"
        Exit Function
    End If
    
    If g�������_����.IC���� <> Nvl(rsTemp!����) Then
        ShowMsgbox "�ò��˵�IC���������,�����ǲ����������˵�IC��!"
        Exit Function
    End If
    
    '--------------------------------------------
    '���ó�������ӿ�
    סԺ�������_���� = סԺ���㼰����_����(True, lng����ID, lng����ID, lng����ID, lng��ҳID, intinsure)
    If סԺ�������_���� = False Then Exit Function
    
    סԺ�������_���� = False
    
    '----------------------------------------------
    '��ѯ����������Ժ�������Ϣ��
    '----------------------------------------------
    Dim str��Ժ����ʱ�� As String
    Dim strInfor As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��λ�� As String
    Dim strת�ﵥ�� As String
    Dim lng���� As Long
    
    '�������鿨
     lng���� = IIf(intinsure = 83, 2, 1)
    
     If ��ȡ�������_����(lng����, intinsure) = False Then Exit Function
    
    '����δ����õĲ��˲�������HIS��Ժ��������Ϊ�Ѱ���ҽ����Ժ���������ٰ���HIS��Ժ
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "ҽ���ѳ�Ժ�Ĳ��˲���������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
               
    On Error GoTo errHand
    
    '��ȡ���˵���ر�����Ϣ
    gstrSQL = "select ����,����ID,��Ա���,ҽ����,�Ҷȼ� From �����ʻ� where  ����=" & intinsure & "  and ����id=" & lng����ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "������Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    
    strת�ﵥ�� = Nvl(rsTemp!��Ա���)
    
    If lng���� = 2 Then
        strInfor = Lpad(gstrҽԺ����_����, 6) 'ҽԺ����    CHAR    1   6      Y   Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ����), 10)     '���ձ��    CHAR    7   10      Ժ����д
    Else
        strInfor = Lpad(gstrҽԺ����_����, 4) 'ҽԺ����    CHAR    1   4       Y   Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ����), 8)     '���ձ��    CHAR    5   8       Y   Ժ��
    End If
    
    strInfor = strInfor & Lpad(g�������_����.�������, 4)       '�������    NUM 13  4   ���������Ժʱ�������  Y   Ժ��
    
    
    '�ڲ���ʶ:5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ
    'ҽ����ʶ:2-סԺ����,4-��ͥ��������,O-��������סԺ����,Q-���˱��ս���
    If intinsure = TYPE_������ Then
        str������� = Decode(Nvl(rsTemp!�Ҷȼ�, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    Else
        'ҽ����ʶ:5-��ͨסԺ
        'ҽ����ʶ:2-סԺ����
        str������� = Decode(Nvl(rsTemp!�Ҷȼ�, 0), 5, "2", 6, "4", "2")
    End If
    '��ȡ������Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') ��Ժ����ʱ��," & _
            " to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����" & _
            " From ������ҳ A,���ű� B,������Ϣ C" & _
            " Where A.����id=C.����id and C.����id=" & lng����ID & _
            "       and A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ��Ϣ", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    str��Ժ���� = Nvl(rsTemp!��Ժ����)
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!סԺ��, 0), 10)       '��־��  CHAR    17  10      Y   Ժ�������ݶ�Ϊ�գ�סԺ��ΪסԺ��
    strInfor = strInfor & Lpad(Nvl(rsTemp!��Ժ����), 8)         '��Ժ���� Date 27  8   ����ʵ����Ժ���ڣ���ʽΪyyyymmdd    Y   Ժ��
    strInfor = strInfor & Rpad(Nvl(rsTemp!��Ժ����ʱ��), 16)    '�Ǽ�ʱ��    DATETIME    35  16  ��ȷ���룬���ݷ��غ��ʽΪyyyymmddhhmiss�����Կո�λ  Y   Ժ��
    
    strInfor = strInfor & Lpad(str�������, 1)                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��
    
    gstrSQL = "Select ����ID,����,����� From ��λ״����¼ D where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ��Ϣ", CLng(Nvl(rsTemp!��ǰ����ID, 0)), CLng(Nvl(rsTemp!��ǰ����, 0)))
    If rsTemp.EOF Then
        str��λ�� = Space(10)
    Else
        str��λ�� = Trim(Nvl(rsTemp!�����)) & "��" & Trim(Nvl(rsTemp!����)) & "��"
        str��λ�� = Lpad(str��λ��, 10)
        str��λ�� = Substr(str��λ��, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����||'~^||'||b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1 and a.����id=" & lng����ID & " and a.��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    
    Dim str��Ժ��ϱ��� As String
    Dim str��Ժ�������  As String
    Dim strȷ����ϱ��� As String
    Dim strȷ���������  As String
    
    If rsTemp.EOF Then
        str��Ժ��ϱ��� = ""
        str��Ժ������� = ""
        strȷ����ϱ��� = ""
        strȷ��������� = ""
    Else
        str��Ժ������� = Nvl(rsTemp!��Ժ���)
        strȷ��������� = Nvl(rsTemp!ȷ�����)
        If InStr(1, str��Ժ�������, "~^||") <> 0 Then
            str��Ժ��ϱ��� = Split(str��Ժ�������, "~^||")(0)
            str��Ժ������� = Split(str��Ժ�������, "~^||")(1)
        Else
            str��Ժ��ϱ��� = ""
            str��Ժ������� = ""
        End If
        If InStr(1, strȷ���������, "~^||") <> 0 Then
            strȷ����ϱ��� = Split(strȷ���������, "~^||")(0)
            strȷ��������� = Split(strȷ���������, "~^||")(1)
        Else
            strȷ����ϱ��� = ""
            strȷ��������� = ""
        End If
    End If
        
    strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
    strInfor = strInfor & Lpad(Substr(str��Ժ�������, 1, 28), 30) '��Ժ�������    CHAR    68  30      y Ժ��
    strInfor = strInfor & Lpad(strȷ����ϱ���, 16)  'ȷ����ϱ���    CHAR    98  16      N   Ժ��
    strInfor = strInfor & Lpad(Substr(strȷ���������, 1, 28), 30) 'ȷ���������    CHAR    114 30      N   Ժ��
    strInfor = strInfor & Lpad(str��Ժ����, 20)  '�Ʊ�����    CHAR    144 20  �磺�ڿ�    Y   Ժ��
    strInfor = strInfor & str��λ��              '��λ��  CHAR    164 10  �磺2003��12��  N   Ժ��
    strInfor = strInfor & Lpad(strת�ﵥ��, 6)   'ת�ﵥ��    CHAR    174 6       N   Ժ��
    strInfor = strInfor & Space(8)   '��Ժʱ��    DATE    180 8   ϵͳ���û��߽������ݵĳ�Ժʱ���Զ����ɣ�ҽԺ���ÿո�λ���ɡ�  N   ��
    
    strInfor = strInfor & "A"   '�����־    CHAR    188 1   A ��Ժ�Ǽǣ�M �޸���Ժ״̬��Cȡ����Ժ�Ǽ�   Y   Ժ��
    strInfor = strInfor & Space(16)   '����ʱ��    DATATIME    189 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  N   ����
    
    '--------------------------------------------
    '����ҽ���ӿڲ���
    '1004    9   206 ʵʱסԺ�Ǽ������ύ
    סԺ�������_���� = ҵ������_����(lng����, 1004, strInfor, intinsure)
    If סԺ�������_���� = False Then
        'Modify by ZHQ 2005-11-30
        '��������Ժ�Ǽ�ʧ�ܺ���Լ����������䲹��ǼǼ���
        ShowMsgbox "ʵʱסԺ�Ǽ�ʧ��,����ҽ���ʻ������н��в���Ǽ�!"
        סԺ�������_���� = True
        Exit Function
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ҽ����ֹ_����() As Boolean
    mblnInit = False
    ҽ����ֹ_���� = True
End Function

Public Function �����Ǽ�_����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String, ByVal intinsure As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp   As ADODB.Recordset
    '��д�뵥��ͷ����д�뵥����
    '��¼״̬��1-����;����Ϊɾ���������ô�����ֻ�����ŵ���ɾ�����ٲ����µ���
    On Error GoTo errHand
    �����Ǽ�_���� = False
    If gblnסԺ��ϸʱʵ�ϴ� = False Then
        �����Ǽ�_���� = True
        Exit Function
    End If
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then
        gstrSQL = " " & _
            " Select A.id,A.����ID,F.סԺ��,A.NO,A.���,A.ҽ�����,A.��¼����,A.��¼״̬,A.�շ����,D.���,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') �Ǽ�ʱ��, " & _
            "        A.������ ҽ��,V.��� AS ҽ�����,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,Round(A.ʵ�ս��/(A.����*A.����),2) as ʵ�ʼ۸�,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "        C.��Ŀ���� ҽ����Ŀ���� ,C.�Ƿ�ҽ��,decode(a.�����־,2,C.סԺ�ȶ�,C.ͳ��ȶ�) as ͳ��ȶ�,F.סԺ���� AS ��ҳid, " & _
            "        Nvl(K.��ʶ��,Nvl(G.��ʶ����||G.��ʶ����,G.����)) AS ���ұ���,G.���� AS ��Ŀ����,K.���� AS ����, " & _
            "        E.����,E.����,E.����,E.ҽ����,E.����,E.��Ա���,E.��λ����,E.˳���,E.����֤��,E.�ʻ����,E.��ǰ״̬, " & _
            "        E.����ID,E.��ְ,E.�����,E.�Ҷȼ�,to_char(E.����ʱ��,'yyyyMMddhh24miss') ����ʱ�� " & _
            " From סԺ���ü�¼ A,���ű� B,�շ���� D,�����ʻ� E,������Ϣ F,������ҳ F1,�շ�ϸĿ G,��Ա�� V," & _
            "       (Select J.����,O.ҩƷid,O.��ʶ�� From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K, " & _
            "       (Select M.��Ŀ����,M.��Ŀ����,M.�Ƿ�ҽ��,M.�շ�ϸĿid,Q.ͳ��ȶ�,Q.סԺ�ȶ�  From ����֧����Ŀ M,����֧������ Q Where M.����=" & TYPE_������ & " and M.����ID=Q.id) C " & _
            " Where   a.��¼״̬<>0 and   a.����id=E.����ID AND a.����id=F.����ID AND A.����id=F1.����id and F1.����=82 AND F.��ҳid= F1.��ҳid  AND  a.������=V.����(+) AND a.�շ�ϸĿid=k.ҩƷid(+) AND a.�շ�ϸĿid=G.id AND E.����=" & TYPE_������ & "   AND A.�շ����=D.���� AND  " & _
            "           A.��¼����=" & lng��¼���� & " and  A.��¼״̬=" & lng��¼״̬ & " And A.NO='" & str���ݺ� & "'" & _
            "           And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 "
        
        gstrSQL = gstrSQL & " Union all " & _
            " Select A.id,A.����ID,F.סԺ��,A.NO,A.���,A.ҽ�����,A.��¼����,A.��¼״̬,A.�շ����,D.���,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') �Ǽ�ʱ��, " & _
            "        A.������ ҽ��,V.��� AS ҽ�����,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,Round(A.ʵ�ս��/(A.����*A.����),2) as ʵ�ʼ۸�,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "        C.��Ŀ���� ҽ����Ŀ���� ,C.�Ƿ�ҽ��,decode(a.�����־,2,C.סԺ�ȶ�,C.ͳ��ȶ�) as ͳ��ȶ�,F.סԺ���� AS ��ҳid, " & _
            "        Nvl(K.��ʶ��,Nvl(G.��ʶ����||G.��ʶ����,G.����)) AS ���ұ���,G.���� AS ��Ŀ����,K.���� AS ����, " & _
            "        E.����,E.����,E.����,E.ҽ����,E.����,E.��Ա���,E.��λ����,E.˳���,E.����֤��,E.�ʻ����,E.��ǰ״̬, " & _
            "        E.����ID,E.��ְ,E.�����,E.�Ҷȼ�,to_char(E.����ʱ��,'yyyyMMddhh24miss') ����ʱ�� " & _
            " From סԺ���ü�¼ A,���ű� B,�շ���� D,�����ʻ� E,������Ϣ F,������ҳ F1,�շ�ϸĿ G,��Ա�� V," & _
            "       (Select J.����,O.ҩƷid,O.��ʶ�� From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K, " & _
            "       (Select M.��Ŀ����,M.��Ŀ����,M.�Ƿ�ҽ��,M.�շ�ϸĿid,Q.ͳ��ȶ�,Q.סԺ�ȶ�  From ����֧����Ŀ M,����֧������ Q Where M.����=" & TYPE_���������� & " and M.����ID=Q.id) C " & _
            " Where  a.��¼״̬<>0 and a.����id=E.����ID AND a.����id=F.����ID AND A.����id=F1.����id and F1.����=83 AND F.��ҳid= F1.��ҳid  AND  a.������=V.����(+) AND a.�շ�ϸĿid=k.ҩƷid(+) AND a.�շ�ϸĿid=G.id AND E.����=" & TYPE_���������� & "   AND A.�շ����=D.���� AND  " & _
            "           A.��¼����=[1] and  A.��¼״̬=[2] And A.NO=[3]" & _
            "           And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
            " Order by ����ID"
    Else
        gstrSQL = " " & _
            " Select A.id,A.����ID,F.סԺ��,A.NO,A.���,A.ҽ�����,A.��¼����,A.��¼״̬,A.�շ����,D.���,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') �Ǽ�ʱ��, " & _
            "        A.������ ҽ��,V.��� AS ҽ�����,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,Round(A.ʵ�ս��/(A.����*A.����),2) as ʵ�ʼ۸�,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "        C.��Ŀ���� ҽ����Ŀ���� ,C.�Ƿ�ҽ��,decode(a.�����־,2,C.סԺ�ȶ�,C.ͳ��ȶ�) as ͳ��ȶ�,F.סԺ���� AS ��ҳid, " & _
            "        Nvl(K.��ʶ��,Nvl(G.��ʶ����||G.��ʶ����,G.����)) AS ���ұ���,G.���� AS ��Ŀ����,K.���� AS ����, " & _
            "        E.����,E.����,E.����,E.ҽ����,E.����,E.��Ա���,E.��λ����,E.˳���,E.����֤��,E.�ʻ����,E.��ǰ״̬, " & _
            "        E.����ID,E.��ְ,E.�����,E.�Ҷȼ�,to_char(E.����ʱ��,'yyyyMMddhh24miss') ����ʱ�� " & _
            " From סԺ���ü�¼ A,���ű� B,�շ���� D,�����ʻ� E,������Ϣ F,������ҳ F1,�շ�ϸĿ G,��Ա�� V," & _
            "       (Select J.����,O.ҩƷid,O.��ʶ�� From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K, " & _
            "       (Select M.��Ŀ����,M.��Ŀ����,M.�Ƿ�ҽ��,M.�շ�ϸĿid,Q.ͳ��ȶ�,Q.סԺ�ȶ�  From ����֧����Ŀ M,����֧������ Q Where M.����=" & TYPE_������ & " and M.����ID=Q.id) C " & _
            " Where   a.��¼״̬<>0 and   a.����id=E.����ID AND a.����id=F.����ID AND A.����id=F1.����id and F1.����=82 AND F.סԺ����= F1.��ҳid  AND  a.������=V.����(+) AND a.�շ�ϸĿid=k.ҩƷid(+) AND a.�շ�ϸĿid=G.id AND E.����=" & TYPE_������ & "   AND A.�շ����=D.���� AND  " & _
            "           A.��¼����=" & lng��¼���� & " and  A.��¼״̬=" & lng��¼״̬ & " And A.NO='" & str���ݺ� & "'" & _
            "           And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 "
        
        gstrSQL = gstrSQL & " Union all " & _
            " Select A.id,A.����ID,F.סԺ��,A.NO,A.���,A.ҽ�����,A.��¼����,A.��¼״̬,A.�շ����,D.���,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') �Ǽ�ʱ��, " & _
            "        A.������ ҽ��,V.��� AS ҽ�����,B.���� ��������,A.�շ�ϸĿID,A.���㵥λ,Round(A.ʵ�ս��/(A.����*A.����),2) as ʵ�ʼ۸�,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�, " & _
            "        C.��Ŀ���� ҽ����Ŀ���� ,C.�Ƿ�ҽ��,decode(a.�����־,2,C.סԺ�ȶ�,C.ͳ��ȶ�) as ͳ��ȶ�,F.סԺ���� AS ��ҳid, " & _
            "        Nvl(K.��ʶ��,Nvl(G.��ʶ����||G.��ʶ����,G.����)) AS ���ұ���,G.���� AS ��Ŀ����,K.���� AS ����, " & _
            "        E.����,E.����,E.����,E.ҽ����,E.����,E.��Ա���,E.��λ����,E.˳���,E.����֤��,E.�ʻ����,E.��ǰ״̬, " & _
            "        E.����ID,E.��ְ,E.�����,E.�Ҷȼ�,to_char(E.����ʱ��,'yyyyMMddhh24miss') ����ʱ�� " & _
            " From סԺ���ü�¼ A,���ű� B,�շ���� D,�����ʻ� E,������Ϣ F,������ҳ F1,�շ�ϸĿ G,��Ա�� V," & _
            "       (Select J.����,O.ҩƷid,O.��ʶ�� From ҩƷĿ¼ O, ҩƷ��Ϣ H,ҩƷ���� J WHERE O.ҩ��id=H.ҩ��id and H.����=J.����) K, " & _
            "       (Select M.��Ŀ����,M.��Ŀ����,M.�Ƿ�ҽ��,M.�շ�ϸĿid,Q.ͳ��ȶ�,Q.סԺ�ȶ�  From ����֧����Ŀ M,����֧������ Q Where M.����=" & TYPE_���������� & " and M.����ID=Q.id) C " & _
            " Where  a.��¼״̬<>0 and a.����id=E.����ID AND a.����id=F.����ID AND A.����id=F1.����id and F1.����=83 AND F.סԺ����= F1.��ҳid  AND  a.������=V.����(+) AND a.�շ�ϸĿid=k.ҩƷid(+) AND a.�շ�ϸĿid=G.id AND E.����=" & TYPE_���������� & "   AND A.�շ����=D.���� AND  " & _
            "           A.��¼����=[1] and  A.��¼״̬=[2] And A.NO=[3]" & _
            "           And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
            " Order by ����ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ǽ�", lng��¼����, lng��¼״̬, str���ݺ�)
    If rsTemp.RecordCount = 0 Then
        MsgBox "δ�ҵ�������¼����ҽ����������������ʧ�ܣ�[�����Ǽ�]", vbInformation, gstrSysName
        Exit Function
    End If
    �����Ǽ�_���� = �ϴ�����_����(rsTemp, intinsure)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function �ϴ�����_����(ByVal rsExse As ADODB.Recordset, ByVal intinsure As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:rsExse-��ϸ����
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------


    Dim lng����ID As Long
    Dim curTotal As Currency
    Dim blnUpload As Boolean
    Dim rsPara As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim str��ϸ As String
    Dim str��Ŀͳ�Ʒ��� As String
    Dim str��ϱ��� As String
    Dim str������� As String
    Dim strInsertSQL As String
    
    
    Dim strTmp As String
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select * from ���ղ��� where ���� in (82,83)"
    zlDatabase.OpenRecordset rsPara, gstrSQL, "�ϴ�������ȡ����"
    With rsExse
        Do While Not .EOF
            lng����ID = Nvl(!����ID, 0)
            'ȷ���������
            '�ϴ���ϸ��¼,ʵʱҽ����ϸ����
                
            If Nvl(!����, 0) = TYPE_���������� Then '������
                str��ϸ = Lpad(gstrҽԺ����_����, 6)     'ҽԺ����    CHAR    1   6       Ժ����д
                str��ϸ = str��ϸ & Lpad(Nvl(!ҽ����), 10)  '���ձ��    CHAR    7   10      Ժ����д
            Else
                str��ϸ = Lpad(gstrҽԺ����_����, 4)     'ҽԺ����    CHAR    1   4       Ժ��
                str��ϸ = str��ϸ & Lpad(Nvl(!ҽ����), 8)   '���˱��    CHAR    5   8       Ժ��
            End If
            
            str��ϸ = str��ϸ & Lpad(Nvl(!סԺ��, 0), 10) '��־��  CHAR    13  10  ������ϸ�Կո�λ,סԺ��סԺ��  Ժ��
            str��ϸ = str��ϸ & Lpad(Nvl(!˳���, 0), 4)   '�������    NUM 23  4   סԺ��ϸ�����������Ժ�Ǽ�ʱ�������������ϸ:                         ������ڱ��ν���������� Ժ��
            
            'Modified By ���� 2004-07-29 ԭ�򣺴���NO��
            str��ϸ = str��ϸ & Lpad(Mid(Nvl(!NO, "00000000"), 2, 7), 10)     '������  NUM 27  10      Ժ��
            str��ϸ = str��ϸ & Lpad(CStr(Nvl(!���, 0)), 10)      '������Ŀ���    NUM 37  10  ��Ӧ�����ŵļǼ���Ŀ���    Ժ��
            
            '������Ϊ���ݺ�  CHAR    41  10  ҽ���ţ�    Ժ����д
            str��ϸ = str��ϸ & Lpad(Nvl(!ҽ�����, " "), 10)     'ҽ����  CHAR    47  10  ������Ӧҽ����ҽ����¼�ţ�������ϸ��û��ҽ����ҽԺ�Կո�λ    Ժ��
            str��ϸ = str��ϸ & Get�������(0, Nvl(!�Ҷȼ�))      '�������    CHAR    57  1   ȡֵ���"�������"˵��  Ժ��
            str��ϸ = str��ϸ & Rpad(Nvl(!�Ǽ�ʱ��), 16)      '��������ʱ�䣨Ͷҩʱ�䣩    DATETIME    58  16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ    Ժ��
            str��ϸ = str��ϸ & Lpad(Nvl(!���ұ���), 20)      '��Ŀ����    CHAR    74  20  �Ƽ���Ŀ����    Ժ��
            str��ϸ = str��ϸ & Lpad(Nvl(!��Ŀ����), 20)      '��Ŀ����    CHAR    94  20      Ժ��

            If !�Ƿ�ҽ�� = 1 Then
                str��ϸ = str��ϸ & Lpad(1 - Nvl(!ͳ��ȶ�, 0) / 100, 6) '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
            Else
                str��ϸ = str��ϸ & Lpad(1, 6)    '�Էѱ��� Char 114 6   ����Ǳ��շ�Χ�ڷ��ã��Էѱ�������Ϊ��0����0.1��0����10������ ����Ǳ��շ�Χ����ҩ�Էѱ���Ϊ��1��100����  Ժ��
            End If
            rsPara.Filter = 0
            rsPara.Filter = " ������='" & Nvl(!���) & "' and ����=" & Nvl(!����, 0)
            str��Ŀͳ�Ʒ��� = ""
            If Not rsPara.EOF Then
                strTmp = Nvl(rsPara!����ֵ)
                If InStr(1, strTmp, ";") <> 0 And strTmp <> ";" Then
                    strTmp = Split(strTmp, ";")(1)
                    If strTmp <> "" Then
                        str��Ŀͳ�Ʒ��� = Substr(strTmp, 1, 1)
                        str��ϸ = str��ϸ & Substr(strTmp, 1, 1)   '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                    Else
                        str��ϸ = str��ϸ & Space(1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                    End If
                Else
                    str��ϸ = str��ϸ & Space(1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
                End If
            Else
                    str��ϸ = str��ϸ & Space(1)    '��Ŀͳ�Ʒ���    CHAR    120 1   ���ע��,����ʵ�ַ�ʽ?  Ժ��
            End If
            
            '2005-08-02����������
            If Nvl(!����, 0) = TYPE_���������� Then
                str��ϸ = str��ϸ & Lpad(Nvl(!����), 10)  '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                str��ϸ = str��ϸ & Lpad(Nvl(!ʵ�ʼ۸�), 10) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
            Else
                str��ϸ = str��ϸ & Lpad(Nvl(!����), 6)  '����    NUM 121 6   �巽����Ϊ��ֵ  Ժ��
                str��ϸ = str��ϸ & Lpad(Nvl(!ʵ�ʼ۸�), 8) '����    NUM 127 8   ��������ָ�ֵ  Ժ��
            End If
            str��ϸ = str��ϸ & Lpad(Nvl(!���㵥λ), 4) '��λ    CHAR    135 4       Ժ��
            str��ϸ = str��ϸ & Lpad(Nvl(!����), 20)      '����    CHAR    139 20  �����Ƭ����    Ժ��
            str��ϸ = str��ϸ & Lpad(Nvl(!ҽ��), 8)      'ҽʦ����    CHAR    159 8       Ժ��
            'ȷ��������
            
            strTmp = Get��Ժ���(Nvl(!����ID), Nvl(!��ҳID, 0), False, True)
            If InStr(1, strTmp, "|") <> 0 Then
                
                str��ϸ = str��ϸ & Lpad(Split(strTmp, "|")(1), 16)     '��ϱ���    CHAR    167 16      Ժ��
                strTmp = Split(strTmp, "|")(0)
                strTmp = Lpad(strTmp, 30)
                str��ϱ��� = Split(strTmp, "|")(1)
                str������� = Split(strTmp, "|")(0)
                    
                str��ϸ = str��ϸ & strTmp     '�������    CHAR    183 30      Ժ��
            Else
                str��ϸ = str��ϸ & Space(16)      '��ϱ���    CHAR    167 16      Ժ��
                str��ϸ = str��ϸ & Space(30)     '�������    CHAR    183 30      Ժ��
                str��ϱ��� = ""
                str������� = ""
            End If
            
            str��ϸ = str��ϸ & Space(16)     '����ʱ��    DATETIME    213 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ��Ժ�˿ո�λ  ����
     
            '�ϴ���ϸ
            '1003    7   230 ʵʱҽ����ϸ�����ύ
            �ϴ�����_���� = ҵ������_����(IIf(Nvl(!����, 0) = TYPE_����������, 2, 1), 1003, str��ϸ, intinsure)
            If �ϴ�����_���� = False Then
                ShowMsgbox "�������ʱҽ����ϸ�����ύʧ��,���ܼ���!"
                Exit Function
            End If

            'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
            'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
            zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            .MoveNext
        Loop
    End With
    �ϴ�����_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
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
    On Error GoTo errHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
errHand:
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
    On Error GoTo errHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
errHand:
End Sub

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Private Function ҽ����ϸ�����ύ(ByVal lngҽ��ID As Long, ByVal strסԺ�� As String, ByVal str��Ŀͳ�Ʒ��� As String, ByVal intinsure As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡҽ����ϸ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    
    '5.  ʵʱҽ�������ύ�ӿ�
    
    '--����Ŀǰ������ҽ���ϴ��Ƿ�ǿ��Ҫ��δȷ��,ҽ����ϸ�ύ��ʱ����,ֱ�ӷ���Ϊ�ɹ�״̬��־
    ҽ����ϸ�����ύ = True
    Exit Function

    '��������ҽ���ӿ�
    If intinsure = TYPE_���������� Then
        ҽ����ϸ�����ύ = True
        Exit Function
    End If
    
    
    gstrSQL = " " & _
         " select A.ID,A.���id as �����,decode(A.ҽ����Ч,1,1,0) as  ҽ������,A.ҽ������," & _
         "          A.����ҽ�� as ��ҽ��ҽ��,to_char(A.��ʼִ��ʱ��,'yyyymmddhh24miss') as ��ʼִ��ʱ��,A.У�Ի�ʿ as ִ��ҽ����ʿ����," & _
         "          A.ͣ��ҽ�� as ͣҽ��ҽ��,to_char(A.ͣ��ʱ��,'yyyymmddhh24miss') as ͣҽ��ʱ��,A.ҽ������ as ����˵��, " & _
         "          Decode(B.���,'Z',decode(B.��������,'5','0000','6','0000',a.ҽ������),a.ҽ������) as ҽ������," & _
         "          A.�������� as ҩƷ����,B.���㵥λ as ������λ," & _
         "        A.ִ��Ƶ��,A.Ƶ�ʴ���, " & _
         "        A.����ʱ�� as ��ҽ��ʱ��" & _
         "        " & _
         " from ����ҽ����¼ A,������ĿĿ¼ B  " & _
         " Where A.������Ŀid=B.id and A.id=" & lngҽ��ID
         
    Err = 0
    On Error GoTo errHand:
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽ����ϸ��¼"
    
    If rsTemp.EOF Then
        ShowMsgbox "�޶�Ӧ��ҽ����¼!"
        Exit Function
    End If
    
    With g�������_����
        strInfor = Lpad(gstrҽԺ����_����, 4)   '1   ҽԺ����    CHAR    1   4       Ժ��
        strInfor = strInfor & Lpad(.���˱��, 8) '2   ���˱��    CHAR    5   8       Ժ��
        strInfor = strInfor & Lpad(.�������, 4)     '3   �������    NUM 13  4   ���������Ժʱ�������  Ժ��
        strInfor = strInfor & Lpad(strסԺ��, 10)    '4   ��־��  CHAR    17  10  Ҫ����ϸ���ݡ���־�š���Ӧ  Ժ��
        strInfor = strInfor & Lpad(lngҽ��ID, 10)    '5   ҽ����  CHAR    27  10  ��Ӧ��ϸ���ݵ�ҽ����    Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ������, 0), 1)   '6   ҽ������    CHAR    37  1   1 ������0 ��ʱҽ��
        strInfor = strInfor & Substr(Lpad(Nvl(rsTemp!ҽ������), 80), 1, 80) '7   ҽ����������??  CHAR    38  80  ���������Ժ��Ϣ���á�0000��    Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!��ҽ��ҽ��), 8)   '8   ��ҽ��ҽʦ����  CHAR    118 8       Ժ��
        strInfor = strInfor & Rpad(Nvl(rsTemp!��ʼִ��ʱ��), 16) '9   ��ʼִ��ʱ��    DATATIME    126 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ���������  Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ִ��ҽ����ʿ����), 8)  '10  ִ��ҽ����ʿ����    CHAR    142 8       Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ͣҽ��ҽ��), 8)  '11  ��ֹҽ��ҽʦ����    CHAR    150 8       Ժ��
        strInfor = strInfor & Rpad(Nvl(rsTemp!ͣҽ��ʱ��), 16) '12  ��ֹҽ��ʱ��    DATATIME    158 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڳ���ҽ��������������ʱҽ�����Կո�λ  Ժ��
        strInfor = strInfor & Substr(Lpad(Nvl(rsTemp!����˵��), 30), 1, 30) '13  ��ע    CHAR    174 30  ������ʱҽ��ִ�з���������������    Ժ��
        strInfor = strInfor & Space(16)                  '14  ����ʱ��    DATATIME    204 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  ����
    End With
    
    '1005    8   274 ʵʱҽ������
    ҽ����ϸ�����ύ = ҵ������_����(g�������_����.ҽ������, 1005, strInfor, intinsure)
    Exit Function
errHand:
    '���ûװҽ���Ͳ�ִ��
    ҽ����ϸ�����ύ = True
End Function

Private Function Get���˱䶯��¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵ı䶯���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
        "   Select  ����,���Ӵ�λ,��ʼʱ��,��ֹʱ��,��λ�ȼ�id " & _
        "   From ���˱䶯��¼  " & _
        "   Where  ����id=" & lng����ID & " and ��ҳid=" & lng��ҳID & " and ���� is not null"
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˱䶯���"
    Set Get���˱䶯��¼ = rsTemp
'    Call WriteDebugInfor_����("Get���˱䶯��¼", lng����id)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Set Get���˱䶯��¼ = Nothing
    Exit Function
End Function
Private Function GetסԺ�����¼(ByVal lng����ID As Long, Optional lng��ҳID As Long = 0, Optional ByVal intinsure As Integer) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ��������δ���¼
    '--�����:
    '--������:
    '--��  ��:δ�����
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset


    '--��Ԫ������,��Ҫ�ų����ڳ������ʵ��������ظ���ϸ��¼
    '--4-26,��˳�����Ӷ�Ӥ���ѵ��ж�,�ų�Ӥ���Ѳ��������
    If lng��ҳID <> 0 Then
        '��Ҫ�Ƕ����εĽ�������������
        strSQL = _
            "   Select  A.��¼����,A.NO,A.���,A.����," & _
            "           A.����ID,A.��ҳID,nvl(A.Ӥ����,0) as Ӥ����," & _
            "           A.���մ���ID,A.�շ����,A.�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������," & _
            "           Decode(Sign(Instr(B.���,'��')),0,B.���,Substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
            "           Decode(Sign(Instr(B.���,'��')),0,B.���,Substr(B.���,Instr(B.���,'��')+1)) as ����," & _
            "           A.����,Decode(A.����,0,0,Round(A.���/A.����,4)) as �۸�,A.���,A.ҽ��,w.��� as ҽ�����,A.����ʱ��,A.�Ǽ�ʱ��," & _
            "           A.�Ƿ���,A.������Ŀ��,A.ժҪ,C.��Ŀ���� as ҽ����Ŀ����," & _
            "           C.��Ŀ���� as ҽ����Ŀ����,Q.����ֵ,Q.������,J.ͳ��ȶ�,J.סԺ�ȶ�,J.��׼����,J.�㷨" & _
            "   From (" & _
            "           Select  C.����,Mod(A.��¼����,10) as ��¼����,A.����,A.NO,Nvl(A.�۸񸸺�,���) as ���,A.����ID,A.��ҳID,Nvl(A.Ӥ����,0) as Ӥ����," & _
            "                   A.������ as ҽ��,A.��������ID,A.�շ����,A.�շ�ϸĿID,Nvl(A.���մ���ID,0) as ���մ���ID,Avg(Nvl(A.����,1)*A.����) as ����," & _
            "                   Sum(A.��׼����) as ��׼����,Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as ���,A.����ʱ��,A.�Ǽ�ʱ��,Nvl(A.�Ƿ���,0) as �Ƿ���,Nvl(A.������Ŀ��,0) as ������Ŀ��,A.ժҪ" & _
            "           From סԺ���ü�¼ A,������Ŀ B,�����ʻ� C" & _
            "           Where a.��¼״̬<>0 and A.��ҳid=" & lng��ҳID & " and  A.���ʷ���=1 and A.����id=C.����id And A.������ĿID=B.ID And nvl(A.Ӥ����,0)=0 And A.����ID=" & lng����ID & " and C.����=" & intinsure & _
            "           Group by    C.����,Mod(A.��¼����,10),A.NO,Nvl(A.�۸񸸺�,���),A.����ID,A.��ҳID,A.����,Nvl(A.Ӥ����,0),A.������," & _
            "                       A.��������ID,A.�շ����,A.�շ�ϸĿID,Nvl(A.���մ���ID,0),A.����ʱ��,A.�Ǽ�ʱ��,Nvl(A.�Ƿ���,0),Nvl(A.������Ŀ��,0),A.ժҪ" & _
            "           Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0) A,�շ�ϸĿ B,���ű� X," & _
            "           (Select * From ����֧����Ŀ Where ����=" & intinsure & ") C," & _
            "           (Select M.����, L.������,L.����ֵ from �շ���� M,���ղ��� L  Where M.���=L.������ and L.����=" & intinsure & ")  Q," & _
            "           (Select * from ����֧������  Where ����=" & intinsure & ")  J,��Ա�� W" & _
            "   Where     A.�շ�ϸĿID=B.ID and a.ҽ��=w.����(+) and C.����id=J.ID and a.�շ����=Q.����(+) And A.�շ�ϸĿID=C.�շ�ϸĿID And A.��������ID=X.ID"
    Else
        strSQL = _
            "   Select  A.��¼����,A.NO,A.���,A.����," & _
            "           A.����ID,A.��ҳID,nvl(A.Ӥ����,0) as Ӥ����," & _
            "           A.���մ���ID,A.�շ����,A.�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������," & _
            "           Decode(Sign(Instr(B.���,'��')),0,B.���,Substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
            "           Decode(Sign(Instr(B.���,'��')),0,B.���,Substr(B.���,Instr(B.���,'��')+1)) as ����," & _
            "           A.����,Decode(A.����,0,0,Round(A.���/A.����,4)) as �۸�,A.���,A.ҽ��,w.��� as ҽ�����,A.����ʱ��,A.�Ǽ�ʱ��," & _
            "           A.�Ƿ���,A.������Ŀ��,A.ժҪ,C.��Ŀ���� as ҽ����Ŀ����," & _
            "           C.��Ŀ���� as ҽ����Ŀ����,Q.����ֵ,Q.������,J.ͳ��ȶ�,J.סԺ�ȶ�,J.��׼����,J.�㷨" & _
            "   From (" & _
            "           Select  C.����,Mod(A.��¼����,10) as ��¼����,A.����,A.NO,Nvl(A.�۸񸸺�,���) as ���,A.����ID,A.��ҳID,Nvl(A.Ӥ����,0) as Ӥ����," & _
            "                   A.������ as ҽ��,A.��������ID,A.�շ����,A.�շ�ϸĿID,Nvl(A.���մ���ID,0) as ���մ���ID,Avg(Nvl(A.����,1)*A.����) as ����," & _
            "                   Sum(A.��׼����) as ��׼����,Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as ���,A.����ʱ��,A.�Ǽ�ʱ��,Nvl(A.�Ƿ���,0) as �Ƿ���,Nvl(A.������Ŀ��,0) as ������Ŀ��,A.ժҪ" & _
            "           From סԺ���ü�¼ A,������Ŀ B,�����ʻ� C" & _
            "           Where a.��¼״̬<>0 and  A.���ʷ���=1 and A.����id=C.����id And A.������ĿID=B.ID And nvl(A.Ӥ����,0)=0 And A.����ID=" & lng����ID & " and C.����=" & intinsure & _
            "           Group by    C.����,Mod(A.��¼����,10),A.NO,Nvl(A.�۸񸸺�,���),A.����ID,A.��ҳID,A.����,Nvl(A.Ӥ����,0),A.������," & _
            "                       A.��������ID,A.�շ����,A.�շ�ϸĿID,Nvl(A.���մ���ID,0),A.����ʱ��,A.�Ǽ�ʱ��,Nvl(A.�Ƿ���,0),Nvl(A.������Ŀ��,0),A.ժҪ" & _
            "           Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0) A,�շ�ϸĿ B,���ű� X," & _
            "           (Select * From ����֧����Ŀ Where ����=" & intinsure & ") C," & _
            "           (Select M.����, L.������,L.����ֵ from �շ���� M,���ղ��� L  Where M.���=L.������ and L.����=" & intinsure & ")  Q," & _
            "           (Select * from ����֧������  Where ����=" & intinsure & ")  J,��Ա�� W" & _
            "   Where     A.�շ�ϸĿID=B.ID and a.ҽ��=w.����(+) and C.����id=J.ID and a.�շ����=Q.����(+) And A.�շ�ϸĿID=C.�շ�ϸĿID And A.��������ID=X.ID"
    End If
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTmp, strSQL, "��ȡ����ҽ��δ�����"
    Set GetסԺ�����¼ = rsTmp
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    Set GetסԺ�����¼ = Nothing
    Exit Function
End Function


Private Function Set����ҺŽ�������(ByVal bln���� As Boolean, lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, strSelfNo As String) As Boolean
  '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID��
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Set����ҺŽ������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �ҺŽ���_����(ByVal lng����ID As Long) As Boolean
     �ҺŽ���_���� = Set����ҺŽ�������(False, lng����ID, 0, 0, 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function �Һų���_����(ByVal lng����ID As Long) As Boolean
    �Һų���_���� = Set����ҺŽ�������(False, lng����ID, 0, 0, 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function ��Ժ������Ϣ_����(lng����ID As Long, lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str������� As String
    Dim str��Ժ���� As String
    Dim str��λ�� As String
    Dim strת�ﵥ�� As String
    Dim lng���� As Long
    
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo errHand
    
    '��ȡ���˵���ر�����Ϣ

    gstrSQL = "select ����,����ID,��Ա���,ҽ����,˳���,�Ҷȼ� From �����ʻ� where  ����=" & intinsure & "  and ����id=" & lng����ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    strת�ﵥ�� = Nvl(rsTemp!��Ա���)
    lng���� = IIf(intinsure = 83, 2, 1)
    If lng���� = 2 Then
        strInfor = Lpad(gstrҽԺ����_����, 6) 'ҽԺ����    CHAR    1   6      Y   Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ����), 10)     '���ձ��    CHAR    7   10      Ժ����д
    Else
        strInfor = Lpad(gstrҽԺ����_����, 4) 'ҽԺ����    CHAR    1   4       Y   Ժ��
        strInfor = strInfor & Lpad(Nvl(rsTemp!ҽ����), 8)     '���ձ��    CHAR    5   8       Y   Ժ��
    End If
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!˳���, 1), 4)      '�������    NUM 13  4   ���������Ժʱ�������  Y   Ժ��
    
    '�ڲ���ʶ:5-��ͨסԺ,6-��ͥ����סԺ,7-��������סԺ,8-���˱���סԺ
    'ҽ����ʶ:2-סԺ����,4-��ͥ��������,O-��������סԺ����,Q-���˱��ս���
    
    str������� = Decode(Nvl(rsTemp!�Ҷȼ�, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '��ȡ������Ϣ
    gstrSQL = "Select C.סԺ��,C.��ǰ����id,C.��ǰ����,A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyyMMddhh24miss') ��Ժ����ʱ��," & _
            " to_char(A.�Ǽ�ʱ��,'yyyyMMdd') ��Ժ����" & _
            " From ������ҳ A,���ű� B,������Ϣ C" & _
            " Where A.����id=C.����id and C.����id=[1]" & _
            "       and A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ��Ϣ", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        ShowMsgbox "�ڲ�����ҳ���޴˲���!"
        Exit Function
    End If
    
    str��Ժ���� = Nvl(rsTemp!��Ժ����)
    
    strInfor = strInfor & Lpad(Nvl(rsTemp!סԺ��, 0), 10)      '��־��  CHAR    17  10      Y   Ժ�������ݶ�Ϊ�գ�סԺ��ΪסԺ��
    strInfor = strInfor & Lpad(Nvl(rsTemp!��Ժ����), 8)      '��Ժ���� Date 27  8   ����ʵ����Ժ���ڣ���ʽΪyyyymmdd    Y   Ժ��
    strInfor = strInfor & Rpad(Nvl(rsTemp!��Ժ����ʱ��), 16)     '�Ǽ�ʱ��    DATETIME    35  16  ��ȷ���룬���ݷ��غ��ʽΪyyyymmddhhmiss�����Կո�λ  Y   Ժ��
    strInfor = strInfor & Lpad(str�������, 1)                  '�������    CHAR    51  1   2סԺ��4�Ҵ���O������   Y   Ժ��

    gstrSQL = "Select ����ID,����,����� From ��λ״����¼ D where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ��Ϣ", CLng(Nvl(rsTemp!��ǰ����ID, 0)), CLng(Nvl(rsTemp!��ǰ����, 0)))
    If rsTemp.EOF Then
        str��λ�� = Space(10)
    Else
        str��λ�� = Trim(Nvl(rsTemp!�����)) & "��" & Trim(Nvl(rsTemp!����)) & "��"
        str��λ�� = Lpad(str��λ��, 10)
        str��λ�� = Substr(str��λ��, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.�������,1,b.����||'~^||'||b.����,null)) as ��Ժ���,  " & _
         "        max(decode(A.�������,1,null,b.����||'~^||'||b.����)) as ȷ����� " & _
         " from ������ A,��������Ŀ¼ b " & _
         " where a.����id=b.id and  a.������� in(1,2) and a.��ϴ���=1 and a.����id=" & lng����ID & "and a.��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����ϱ��������"
    Dim str��Ժ��ϱ��� As String
    Dim str��Ժ�������  As String
    Dim strȷ����ϱ��� As String
    Dim strȷ���������  As String
    
    If rsTemp.EOF Then
        str��Ժ��ϱ��� = ""
        str��Ժ������� = ""
        strȷ����ϱ��� = ""
        strȷ��������� = ""
    Else
        str��Ժ������� = Nvl(rsTemp!��Ժ���)
        strȷ��������� = Nvl(rsTemp!ȷ�����)
        If InStr(1, str��Ժ�������, "~^||") <> 0 Then
            str��Ժ��ϱ��� = Split(str��Ժ�������, "~^||")(0)
            str��Ժ������� = Split(str��Ժ�������, "~^||")(1)
        Else
            str��Ժ��ϱ��� = ""
            str��Ժ������� = ""
        End If
        If InStr(1, strȷ���������, "~^||") <> 0 Then
            strȷ����ϱ��� = Split(strȷ���������, "~^||")(0)
            strȷ��������� = Split(strȷ���������, "~^||")(1)
        Else
            strȷ����ϱ��� = ""
            strȷ��������� = ""
        End If
    End If
    strInfor = strInfor & Lpad(str��Ժ��ϱ���, 16)  '��Ժ��ϱ���    CHAR    52  16      Y   Ժ��
    strInfor = strInfor & Lpad(Substr(str��Ժ�������, 1, 28), 30) '��Ժ�������    CHAR    68  30      y Ժ��
    strInfor = strInfor & Lpad(strȷ����ϱ���, 16)  'ȷ����ϱ���    CHAR    98  16      N   Ժ��
    strInfor = strInfor & Lpad(Substr(strȷ���������, 1, 28), 30) 'ȷ���������    CHAR    114 30      N   Ժ��
    strInfor = strInfor & Lpad(str��Ժ����, 20)  '�Ʊ�����    CHAR    144 20  �磺�ڿ�    Y   Ժ��
    strInfor = strInfor & Lpad(Substr(str��λ��, 1, 10), 10)         '��λ��  CHAR    164 10  �磺2003��12��  N   Ժ��
    strInfor = strInfor & Lpad(strת�ﵥ��, 6)   'ת�ﵥ��    CHAR    174 6       N   Ժ��
    strInfor = strInfor & Space(8)   '��Ժʱ��    DATE    180 8   ϵͳ���û��߽������ݵĳ�Ժʱ���Զ����ɣ�ҽԺ���ÿո�λ���ɡ�  N   ��
    strInfor = strInfor & "M"   '�����־    CHAR    188 1   A ��Ժ�Ǽǣ�M �޸���Ժ״̬��Cȡ����Ժ�Ǽ�   Y   Ժ��
    strInfor = strInfor & Space(16)   '����ʱ��    DATATIME    189 16  ��ȷ�����ʽΪ��yyyymmddhhmiss�����Կո�λ�����ڼ�¼���ݵ���ҽ�����ĵ�ʱ�䣬Ժ�˿ո�λ  N   ����
    
    '1004    9   206 ʵʱסԺ�Ǽ������ύ
    ��Ժ������Ϣ_���� = ҵ������_����(lng����, 1004, strInfor, intinsure)
    If ��Ժ������Ϣ_���� = False Then
        ShowMsgbox "ʵʱסԺ�Ǽ������ύʧ��!"
        Exit Function
    End If
    ��Ժ������Ϣ_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetItemInfo_����(ByVal lngPatiID As Long, ByVal lngItemID As Long, ByVal intinsure As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�������˵������ʾ��Ϣ
    '--�����:
    '--������:
    '--��  ��:��ʾ��
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strҽ�Ƹ��ʽ As String
    Dim int���� As Integer
    Dim bln��Ժ As Boolean
    Dim dblͳ����� As Double
    Dim strMsgInfor As String
    
'    '��һ��:ȷ���Ƿ�ҽ������
'    gstrSQL = "Select ����id,����,nvl(��ǰ״̬,0) as ״̬ from �����ʻ�  where ����id=" & lngPatiID & " and ����=" & intinsure
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�ж��Ƿ�Ϊҽ������!"
'    If rsTemp.EOF Then
'        rsTemp.Close
'        GetItemInfo_���� = ""
'        Exit Function
'    End If
'
'    int���� = NVL(rsTemp!����, 0)
'    bln��Ժ = NVL(rsTemp!״̬, 0) > 0
    
    '�ڶ���:ȷ��ҽ�Ƹ��ʽ
    gstrSQL = "Select ҽ�Ƹ��ʽ,decode(��ǰ����id,null,0,1) as ��Ժ״̬,nvl(����,0) as ���� from ������Ϣ where ����id=" & lngPatiID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽ�Ƹ���ʽ"
    
    strҽ�Ƹ��ʽ = Nvl(rsTemp!ҽ�Ƹ��ʽ)
    int���� = rsTemp!����
    If rsTemp!��Ժ״̬ = 1 Then
        bln��Ժ = True
    Else
        bln��Ժ = False
    End If
    
        
    '��������ȷ���շ�ϸĿ���������
    gstrSQL = "" & _
        "   Select a.����,b.����,b.����,b.����,b.�㷨,a.��Ŀ����,b.ͳ��ȶ�,b.��׼����,b.סԺ�ȶ�,a.�Ƿ�ҽ�� " & _
        "   From ����֧����Ŀ a,����֧������ b " & _
        "   where a.����id=b.id and a.����=b.���� and a.�շ�ϸĿid=" & lngItemID & _
        "     and a.����=decode('" & strҽ�Ƹ��ʽ & "'," & _
        "    '������ҽ�Ʊ���',decode(" & int���� & ",0,82," & int���� & "),82)"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����֧������"
    strMsgInfor = ""
    If rsTemp.RecordCount = 0 Then
        GetItemInfo_���� = ""
        ShowMsgbox "δ�ڱ���֧����Ŀ��������ر������Ӧ��ϵ,����!"
        Exit Function
    End If
    If InStr(1, "������ҽ�Ʊ���;��ҵ����;���˱���;��������;��ҵ����;", IIf(strҽ�Ƹ��ʽ = "", "D", strҽ�Ƹ��ʽ) & ";") <> 0 Then
        '   ҽ�Ƹ��ʽΪ������ҽ�������ҽ������ҵ���ݡ����˱��ա��������ա���ҵ���յģ������������մ�����ҽ���ӿڵ�ҽ����Ŀ�����е�ҽ�����ඨ���еı�������������ʾ
        '   ҽ�Ƹ��ʽΪ������ҽ���ģ������������տ�����ҽ���ӿڵ�ҽ�����ඨ���еı�������������ʾ
        If bln��Ժ Then
            If Nvl(rsTemp!�㷨, 0) = 2 Then
                 '���˺�:200404,���õ��㷨2(���ö����),������б�������
                 strMsgInfor = "����Ŀ�̶�����:ÿ��" & Format(Nvl(rsTemp!��׼����, 0), "#####0.00;-####0.00; ;") & "Ԫ"
            Else
                 If rsTemp!סԺ�ȶ� < 100 Then
                 strMsgInfor = "����ĿסԺ�Ը�����:" & Format(100 - Nvl(rsTemp!סԺ�ȶ�, 0), "#####0.00;-####0.00; ;") & "%"
                 End If
            End If
        Else
                 If rsTemp!ͳ��ȶ� < 100 Then
                 strMsgInfor = "����Ŀ�����Ը�����:" & Format(100 - Nvl(rsTemp!ͳ��ȶ�, 0), "#####0.00;-####0.00; ;") & "%"
                 End If
        End If
    ElseIf InStr(1, "����ҽ��;��ͬ��λ;", IIf(strҽ�Ƹ��ʽ = "", "D", strҽ�Ƹ��ʽ) & ";") <> 0 Then
        '   ҽ�Ƹ��ʽΪ����ҽ�ơ���ͬ��λ�ģ������������մ�����ҽ���ӿڵ�ҽ����Ŀ�����е���ҵ���ѱ������������ʾ��
        If Val(Nvl(rsTemp!��Ŀ����)) < 100 Then
        strMsgInfor = "����Ŀ�Ը�����:" & Format(100 - Val(Nvl(rsTemp!��Ŀ����)), "#####0.00;-####0.00; ;") & "%"
        End If
    End If
    If strMsgInfor <> "" Then
        ShowMsgbox strMsgInfor
    End If
    GetItemInfo_���� = ""
End Function

Public Sub WriteDebugInfor_����(ByVal strCallFunctionName As String, ByVal lng����ID As Long)
'��������Ϣд���ļ���
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        If gblnDebug = False Then Exit Sub
        
        Dim strFile As String
        Dim rsTemp As New ADODB.Recordset
        
        gstrSQL = "Select '����id:'||a.����id||'����'||a.����||'סԺ��:'||a.סԺ�� as ��Ϣ  From ������Ϣ a,�����ʻ� b Where a.��Ժʱ�� Is Null and a.��Ժʱ�� is not null and  b.��ǰ״̬=0 And a.����ID=" & lng����ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��ǰ״̬"
        If Not rsTemp.EOF Then
            '���ڱ仯,���¼����
            strFile = App.Path & "\ҽ�����˵�ǰ״̬�仯����.txt"
            
            If Not Dir(strFile) <> "" Then
                objFile.CreateTextFile strFile
            End If
            Set objText = objFile.OpenTextFile(strFile, ForAppending)
            objText.WriteLine strCallFunctionName & Space(10) & Nvl(rsTemp!��Ϣ)
            objText.WriteLine Format(Now, "yyyy-mm-dd")
            objText.Close
        End If
End Sub

Public Sub ���²���_����(ByVal lngPatiID As Long, ByVal lngPageID As Long, intinsure As Integer)
    '�����޸���һ�γ�Ժ�����
    Dim strIcdCode As String, strIcdName As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrName
    
    gstrSQL = "select ������Ϣ from ������ where �������=3 and ��ϴ���=1 and ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ϴγ�Ժ���", lngPatiID, lngPageID - 1)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    strIcdCode = Mid(rsTemp(0).Value, 2, InStr(1, rsTemp(0).Value, ")") - 2)
    strIcdName = Mid(rsTemp(0).Value, InStr(1, rsTemp(0).Value, ")") + 1)
    Call frm�����Ϣ.ShowME(lngPatiID, strIcdCode, strIcdName)
    
    If strIcdCode = "" Then Exit Sub
    
    'ZHQ 2006-03-15 modify
    '����HISͣ�á�������������Ϊ��ͼ��������Ҫ�޸Ĵ˱�UPDATE���
'    gstrSQL = "Update ������ Set ������Ϣ='(" & strIcdCode & ")" & strIcdName & "'" & _
'            "  Where �������=3 And ��ϴ���=1 And ����ID=" & lngPatiID & " And ��ҳID=" & lngPageID - 1
    gstrSQL = "Update ������ϼ�¼ Set �������='(" & strIcdCode & ")" & strIcdName & "'" & _
            "  Where ��¼��Դ=2 And �������=3 And ��ϴ���=1 And ����ID=" & lngPatiID & " And ��ҳID=" & lngPageID - 1
    
    Call SQLTest(App.ProductName, "���³�Ժ���", gstrSQL)
    gcnOracle.Execute gstrSQL
    Call SQLTest
    Exit Sub

ErrName:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function IsParaBig(intinsure) As Boolean
    '��鵱ǰҽ���Ƿ������������סԺ����
    Dim rsTemp As New ADODB.Recordset
    
    If intinsure <> 82 And intinsure <> 83 Then
        Exit Function
    End If
    On Error GoTo ErrName
    gstrSQL = "Select ����ֵ From ���ղ��� where ������='�����ʹ��סԺ����' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", intinsure)
    
    If rsTemp.RecordCount <= 0 Then
        MsgBox "ϵͳ������Ӳ�����������ϵͳ����Ա��ϵ��"
        Exit Function
    End If
    If rsTemp!����ֵ = 1 Then
        IsParaBig = True
    Else
        IsParaBig = False
    End If
    Exit Function

ErrName:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function IsParaQ(intinsure) As Boolean
    '��鵱ǰҽ���Ƿ���������������סԺ����
    Dim rsTemp As New ADODB.Recordset
    
    If intinsure <> 82 Then
        Exit Function
    End If
    On Error GoTo ErrName
    gstrSQL = "Select ����ֵ From ���ղ��� where ������='��������ʹ��סԺ����' and ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", intinsure)
    
    If rsTemp.RecordCount <= 0 Then
        MsgBox "ϵͳ������Ӳ�����������ϵͳ����Ա��ϵ��"
        Exit Function
    End If
    If rsTemp!����ֵ = 1 Then
        IsParaQ = True
    Else
        IsParaQ = False
    End If
    Exit Function

ErrName:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function






