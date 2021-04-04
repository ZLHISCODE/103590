VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_cd1800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Cell Dyn 1800
'����:��������
'�û�:�����ж�Ժ,��������ҽԺ
'�޸�:2007-03-28 �¶�
'+ ͼ�����ݽ���

Private mstrRBC As String, mstrWBC As String, mstrPLT As String '��ͼ�����ݡ�

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim i As Integer
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strValues As String
    Dim lng_S As Long, lng_E As Long, str_Record As String, Str_Img As String
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource Like "***" Then
        
        strReserved = strSource
        
        Do
            lng_S = InStr(strReserved, "")
            lng_E = InStr(strReserved, "")
            str_Record = Mid(strReserved, lng_S, lng_E - lng_S + 1)
            
            strReserved = Mid(strReserved, lng_E + 1)
            
            If Mid(str_Record, 3, 3) = Space(3) Then
                '---- ��������
                
                iStrPoint = 1
                Do While iStrPoint <= Len(str_Record)
                    strBuffer = ReadSubString(str_Record, iStrPoint, 1)
                    '�ж�һ�����Ƿ�ʼ
                    If strBuffer = "" Then
                        '��¼һ���������
                        iCurrGrpEntry = iStrPoint - 1 '63530918
                        
                        iStrPoint = iStrPoint + 8
                        strBuffer = ReadSubString(str_Record, iStrPoint, 6)
                        If strBuffer = "CD1800" Then
                            If Len(str_Record) - iStrPoint >= 280 Then
                                '����ʱ�� '060520061632
                                iStrPoint = iStrPoint + 32
                                strData = ReadSubString(str_Record, iStrPoint, 2)   '��
                                iStrPoint = iStrPoint + 1
                                strData = strData & "-" & ReadSubString(str_Record, iStrPoint, 2) '��
                                iStrPoint = iStrPoint + 1
                                strData = ReadSubString(str_Record, iStrPoint, 2) & "-" & strData '��
                                iStrPoint = iStrPoint + 3
                                strData = strData & " " & ReadSubString(str_Record, iStrPoint, 5) 'ʱ��
                                'ʱ���ʽ����Ϊ��yyyy-mm-dd hh:mm:ss
                                strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                                
                                '������ţ�����Ϊ���� 8λ
                                iStrPoint = iStrPoint + 3 '��λ���걾λ
                               ' strSampleID = Trim(ReadSubString(strSource, iStrPoint, 11))
                                strData = strData & "|" & Val(ReadSubString(str_Record, iStrPoint, 16))
                                '������
                                strData = strData & "| " '��������֧��
                                '�걾
                                strData = strData & "|ѪҺ" '��������֧��
                                '�ʿ�Ʒ
                                strData = strData & "|0" '��������֧�֣���Ϊ���ʿ�Ʒ
                                '������Ŀ
                                iStrPoint = iStrPoint + 95
                                For i = 1 To 21
                                    strData = strData & "|" & i
                                    iStrPoint = iStrPoint + 1
                                    Select Case i
                                        Case 1, 2, 3, 4, 5, 6, 9, 10, 11, 15, 17, 13
                                            strValues = ReadSubString(str_Record, iStrPoint, 5)
                                            If Val(strValues) <> 0 Then
                                                strData = strData & "|" & Val(strValues) / 10
                                            Else
                                                strData = strData & "|----"
                                            End If
                                        Case 8, 12, 14, 18
                                            strValues = ReadSubString(str_Record, iStrPoint, 5)
                                            If Val(strValues) <> 0 Then
                                                strData = strData & "|" & Val(strValues)
                                            Else
                                                strData = strData & "|----"
                                            End If
                                           
                                        Case 19, 21, 20
                                           
                                            strValues = ReadSubString(str_Record, iStrPoint, 5)
                                            If Val(strValues) <> 0 Then
                                                strData = strData & "|" & Val(strValues) / 1000
                                            Else
                                                strData = strData & "|----"
                                            End If
                                         Case 7, 16
                                            strValues = ReadSubString(str_Record, iStrPoint, 5)
                                            If Val(strValues) <> 0 Then
                                                strData = strData & "|" & Val(strValues) / 100
                                            Else
                                                strData = strData & "|----"
                                            End If
                                    End Select
                                    '�������ѽ���
                                    iCurrGrpEntry = 0
                                Next
                                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                            End If
                        Else
                            '��������Ч
                            If Len(strBuffer) = 2 Then iCurrGrpEntry = 0
                        End If
                        'ָ�붨λ����һ������֮��
                        MoveToNextLine str_Record, iStrPoint
                    End If
                Loop   '��������
                
                
                If strResult <> "" Then
                    ''�н������ͼ�����ݸ��ں��档
                    Str_Img = ""
                    
                    If mstrWBC <> "" Then
                        If Split(mstrWBC, "|")(0) = Mid(str_Record, 46, 38) Then
                            '���ڣ��걾����ͬ������ͼ������
                            Str_Img = Str_Img & "^" & Split(mstrWBC, "|")(1)
                            mstrWBC = ""
                        End If
                    End If
                    
                    If mstrRBC <> "" Then
                        If Split(mstrRBC, "|")(0) = Mid(str_Record, 46, 38) Then
                            '���ڣ��걾����ͬ������ͼ������
                            Str_Img = Str_Img & "^" & Split(mstrRBC, "|")(1)
                            mstrRBC = ""
                        End If
                    End If
                    
                    If mstrPLT <> "" Then
                        If Split(mstrPLT, "|")(0) = Mid(str_Record, 46, 38) Then
                            '���ڣ��걾����ͬ������ͼ������
                            Str_Img = Str_Img & "^" & Split(mstrPLT, "|")(1)
                            mstrPLT = ""
                        End If
                    End If
                    If Str_Img <> "" Then strResult = strResult & vbCrLf & Mid(Str_Img, 2)
                End If
            
            ElseIf Mid(str_Record, 3, 3) = "WBC" Then
                mstrWBC = Mid(str_Record, 46, 38) & "|WBC;0;120,262,43.33-,86.66-100,129.996-,173.326-200,216.656-,259.986-300;" & Replace(Mid(str_Record, 179, 1029), ",", ";")
            ElseIf Mid(str_Record, 3, 3) = "RBC" Then
                mstrRBC = Mid(str_Record, 46, 38) & "|RBC;0;120,262,52-,104-100,156-,208-200,260-;" & Replace(Mid(str_Record, 179, 1029), ",", ";")
            ElseIf Mid(str_Record, 3, 3) = "PLT" Then
                mstrPLT = Mid(str_Record, 46, 38) & "|PLT;0;120,262,65-10,130-20,195-30,260-;" & Replace(Mid(str_Record, 179, 1029), ",", ";")
            End If
            
            If InStr(strReserved, "") <= 0 Then Exit Do
        Loop '���봮��ȫ��������
        'strCmd = Chr(&H6)
    Else                    '���ݲ�ȫ���´ν���
        'strCmd = Chr(&H6)
        strReserved = strSource
        Exit Sub
    End If
    
    
''    'û���ҵ�����λʱ�˳�
''    If InStr(strSource, "") = 0 Then
''        strReserved = strSource
''        Exit Sub
''    End If
''
''    '����δ������ԭʼ��
''    If iCurrGrpEntry > 0 Then
''        strReserved = Mid(strSource, iCurrGrpEntry)
''    Else
''        '���豸���͡������յ�������
''         strCmd = Chr(&H6&)
''    End If
    
    
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    'ָ�붨λ����һ������֮��
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = "]" Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '��ȡ�Ӵ���ͬʱָ���ƶ�����ȡ���ַ���֮��
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function