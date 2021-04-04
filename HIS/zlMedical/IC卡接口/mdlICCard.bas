Attribute VB_Name = "mdlICCard"
Option Explicit

'######################################################################################################################
'������API����������

Public gstrSysName As String

'��������غ�����MW-ET-G��
Public Declare Function auto_init Lib "mwic_32.dll" (ByVal port As Integer, ByVal baud As Long) As Long
Public Declare Function ic_init Lib "mwic_32.dll" (ByVal port As Integer, ByVal baud As Long) As Long
Public Declare Function get_status Lib "mwic_32.dll" (ByVal icdev As Long, ByRef status As Integer) As Integer
Public Declare Function set_baud Lib "mwic_32.dll" (ByVal icdev As Long, ByVal baud As Long) As Integer
Public Declare Function cmp_dvsc Lib "mwic_32.dll" (ByVal icdev As Long, ByVal length As Integer, ByVal data_buffer As String) As Integer
Public Declare Function srd_dvsc Lib "mwic_32.dll" (ByVal icdev As Long, ByVal length As Long, ByVal data_buffer As String) As Integer
Public Declare Function swr_dvsc Lib "mwic_32.dll" (ByVal icdev As Long, ByVal length As Integer, ByVal data_buffer As String) As Integer
Public Declare Function setsc_md Lib "mwic_32.dll" (ByVal icdev As Long, ByVal mode As Integer) As Integer
Public Declare Function turn_on Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Public Declare Function turn_off Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Public Declare Function srd_ver Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByVal data_buffer As String) As Integer
Public Declare Function auto_pull Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Public Declare Function dv_beep Lib "mwic_32.dll" (ByVal icdev As Long, ByVal time As Integer) As Integer
Public Declare Function ic_exit Lib "mwic_32.dll" (ByVal icdev As Long) As Integer

'IC����غ�����SLE4442��
Public Declare Function srd_4442 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal offset As Integer, ByVal le As Integer, ByVal data_buffer As String) As Integer
Public Declare Function srd_4442_hex Lib "mwic_32.dll" Alias "srd_4442" (ByVal icdev As Long, ByVal offset As Integer, ByVal le As Integer, ByRef data_buff As Byte) As Integer
Public Declare Function swr_4442 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal offset As Integer, ByVal le As Integer, ByVal data_buffer As String) As Integer
Public Declare Function swr_4442_hex Lib "mwic_32.dll" Alias "swr_4442" (ByVal icdev As Long, ByVal offset As Integer, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Public Declare Function prd_4442 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByVal data_buffer As String) As Integer
Public Declare Function pwr_4442 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal offset As Integer, ByVal le As Integer, ByVal data_buffer As String) As Integer
Public Declare Function chk_4442 Lib "mwic_32.dll" (ByVal icdev As Long) As Integer
Public Declare Function csc_4442 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Public Declare Function wsc_4442 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Public Declare Function rsc_4442 Lib "mwic_32.dll" (ByVal icdev As Long, ByVal le As Integer, ByRef data_buffer As Byte) As Integer
Public Declare Function rsct_4442 Lib "mwic_32.dll" (ByVal icdev As Long, counter As Integer) As Integer
Public Declare Function asc_hex Lib "mwic_32.dll" (ByVal asc As String, ByRef hex As Byte, ByVal le As Integer) As Integer
Public Declare Function hex_asc Lib "mwic_32.dll" (ByRef hex As Byte, ByVal asc As String, ByVal le As Integer) As Integer
Public Declare Function ic_encrypt Lib "mwic_32.dll" (ByVal key As String, ByVal ptrsource As String, ByVal le As Integer, ByRef ptrdest As Byte) As Integer
Public Declare Function ic_decrypt Lib "mwic_32.dll" (ByVal key As String, ByRef ptrdest As Byte, ByVal le As Integer, ByVal ptrsource As String) As Integer

'######################################################################################################################
'�Զ�����̡�����

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��ʾ��ʾ����
    '������ strInfo     Ҫ��ʾ������
    '���أ� ��
    '------------------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Sub ErrorCenter(ByVal intErrorNumber As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��������IC��������
    '������ intErrorNumber      �����
    '���أ� ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strError As String
    
    Select Case intErrorNumber
    Case 100
        strError = "���ڳ�ʼ����������IC���豸�Ƿ������ȷ�Ĵ��ڻ��Դ�Ƿ��Ѿ��򿪣�"
    Case 101
        strError = "���棺����IC���ļ������Ѿ�Ϊ1�ˣ�"
    Case 102
        strError = "д�뱾ϵͳIC���������"
    Case 103
        strError = "У��IC��ԭʼ�������"
    Case 200
        strError = "IC������û׼���û���û�г�ʼ�� ��"
    Case 300
        strError = "��������û�в���IC����"
    Case 400
        strError = "�����������Ŀ����Ͳ��ԣ�"
    Case 500
        strError = "д�������"
    Case 600
        strError = "�Ǳ�ϵͳ��������ϵͳ�ṩ����ϵ��"
    Case 700
        strError = ""
    Case 800
        strError = "д���ݴ�(������Ϣ)��"
    Case Else
        strError = "δ֪IC����д����"
    End Select
    
    ShowSimpleMsg strError
    
End Sub

Public Function Lpad(ByVal strInput As String, ByVal intLen As Integer, Optional ByVal strPad As String = "0") As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ַ���ǰ�油��ָ���ַ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    
    Dim intL As Integer
        
    strInput = Trim(strInput)
    
    intL = LenB(StrConv(strInput, vbFromUnicode))
    If intLen <= intL Then
                
        Lpad = StrConv(MidB(StrConv(strInput, vbFromUnicode), 1, intLen), vbUnicode)
        
    Else
        Lpad = String(intLen - intL, strPad) & strInput
    End If
    
End Function

Public Function GetSubStr(ByVal strInput As String, ByVal intStart As Integer, ByVal intLen As Integer) As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ�Ӵ�����
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    strTmp = Trim(StrConv(MidB(StrConv(strInput, vbFromUnicode), intStart, intLen), vbUnicode))
    
    '�ٽ�ȡǰ��0
    GetSubStr = LTrim(strTmp)
    
End Function

Public Function LTrim(ByVal strInput As String, Optional ByVal strTrim As String = "0") As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ȥ���ַ���ǰ���ָ���ַ�
    '������ strInput          Ҫ������ַ�������
    '       strTrim           Ҫȥ��ǰ����ַ�������
    '���أ� ��ȥ��ָ���ַ����ַ���
    '------------------------------------------------------------------------------------------------------------------
    Dim intLen As Integer
    Dim blnDo As Boolean
    
    On Error GoTo errHand
    
    blnDo = True
    
    Do While blnDo = True
        
        intLen = Len(strInput)
        
        If intLen = 0 Then Exit Do
        
        strInput = IIf(Left(strInput, 1) = strTrim, Mid(strInput, 2), strInput)
        
        If Len(strInput) = intLen Then blnDo = False
        
    Loop
    
    LTrim = strInput
    
    Exit Function
    
errHand:

    MsgBox Err.Description
    
End Function

Public Function SDate(ByVal strTmp As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ����ͨ�ַ���ת��Ϊ�����ַ���
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    If strTmp = "19000101" Then
        SDate = ""
    ElseIf Len(strTmp) = 8 Then
        SDate = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2)
    End If
    
    If IsDate(SDate) = False Then
        SDate = ""
    End If
    
End Function

Public Function DString(ByVal strTmp As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� �������ַ���ת��Ϊ��ͨ�ַ���
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    If strTmp = "" Then
        DString = "19000101"
    ElseIf Len(strTmp) >= 10 Then
        
        If IsDate(strTmp) = False Then
            strTmp = ""
        Else
            DString = Mid(strTmp, 1, 4) & Mid(strTmp, 6, 2) & Mid(strTmp, 9, 2)
        End If
        
    End If
    
End Function

Public Function GetAryValue(strInfo() As String, ByVal strKey As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��ȡ���������е�ָ�����ֵ
    '������ strInfo()        ����
    '       strKey           ��Ŀ
    '���أ� strKey��Ӧ��ֵ
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strItem As String
    Dim strValue As String
    Dim lngPos As Long
    
    For lngLoop = LBound(strInfo) To UBound(strInfo)
        
        lngPos = InStr(strInfo(lngLoop), "=")
        If lngPos > 0 Then
            strItem = Trim(Mid(strInfo(lngLoop), 1, lngPos - 1))
            strValue = Trim(Mid(strInfo(lngLoop), lngPos + 1))
        End If
        
        If strItem = strKey Then
            GetAryValue = strValue
            Exit For
        End If
    Next
    
End Function
