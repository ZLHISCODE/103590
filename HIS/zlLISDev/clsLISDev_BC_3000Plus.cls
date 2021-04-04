VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BC_3000Plus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BC-3000plus
'����:��������
'����:�¶�
'�û�:�ɶ���ͯר��ҽԺ(���Ǹ�)

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    'GetStartCmd = Chr(6)
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    '   ͼ�����ݷ��ڼ��������ûس����з�(vbcrlf)�ָ�
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    '��ʽһ(�ɶ���ͯ)
    'A000000170080220070853006300210007003533211355547712802960907026816543201591031681630530000000000000
    '��ʽ�����Ϻ�������
    'AAAI10P19000000000137010302007145401570019001101271190738083691160324097003141613570189074157139054000000000000000000000

    Dim varLine As Variant, strItem As String, str���� As String, str�걾�� As String
    Dim iRow As Integer, str�����˱걾�ʿ� As String, strTmp As String, strLine As String, str��Ŀ As String
    Dim strValues As String, lngLength As Long, lngType As Long
    Dim blnEnd As Boolean '�Ƿ��ѽ�����
    Dim StrRbc As String, StrWbc As String, StrPLt As String
    
    '��ʼ������ֵ
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "**" & Chr(&H1A&) & "*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
            lngLength = InStr(strReserved, Chr(&H1A&)) - InStr(strReserved, "") + 1
            strLine = Mid(strReserved, InStr(strReserved, ""), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H1A&)) + 1)
            
            str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = "": strTmp = ""
            '����
            
            strItem = ""
           
            If strLine Like "A*" Then
                str���� = Mid(strLine, 16, 4) & "-" & Mid(strLine, 12, 2) & "-" & Mid(strLine, 14, 2)
                
                str�걾�� = Val(Mid(strLine, 3, 8))
                If str�걾�� <> 0 And IsDate(str����) Then
                    lngType = 1
                Else
                    '��һ�ָ�ʽ
                    str�걾�� = Val(Mid(strLine, 11, 12))
                    str���� = Mid(strLine, 28, 4) & "-" & Mid(strLine, 24, 2) & "-" & Mid(strLine, 26, 2)
                    If str�걾�� <> 0 And IsDate(str����) Then
                        lngType = 2
                    End If
                End If
                str�����˱걾�ʿ� = " |ѪҺ|0"
                Select Case lngType
                Case 1
                    strItem = strItem & "|WBC|" & Val(Mid(strLine, 24, 4)) / (10 ^ 1)    'WBC,24,4,3
                    strItem = strItem & "|LYMPHA|" & Val(Mid(strLine, 28, 4)) / (10 ^ 1) 'LYMPH#,28,4,3
                    strItem = strItem & "|MIDA|" & Val(Mid(strLine, 32, 4)) / (10 ^ 1)   'Mid#,32,4,3
                    strItem = strItem & "|GRANA|" & Val(Mid(strLine, 36, 4)) / (10 ^ 1)  'Gran#,36,4,3
                    strItem = strItem & "|LYMPHP|" & Val(Mid(strLine, 40, 3)) / (10 ^ 1) 'Lymph%#,40,3,1
                    strItem = strItem & "|MIDP|" & Val(Mid(strLine, 43, 3)) / (10 ^ 1)   'Mid%,43,3,1
                    strItem = strItem & "|GRANP|" & Val(Mid(strLine, 46, 3)) / (10 ^ 1)  'Gran%,46,3,1
                    strItem = strItem & "|RBC|" & Val(Mid(strLine, 49, 3)) / (10 ^ 2)    'RBC,49,3,2
                    strItem = strItem & "|HGB|" & Val(Mid(strLine, 52, 4)) / (10 ^ 1)    'HGB,52,4,1
                    strItem = strItem & "|MCHC|" & Val(Mid(strLine, 56, 3))              'MCHC,56,4,1
                    strItem = strItem & "|MCV|" & Val(Mid(strLine, 59, 4)) / (10 ^ 1)    'MCV,60,4,2
                    strItem = strItem & "|MCH|" & Val(Mid(strLine, 63, 4)) / (10 ^ 1)    'MCH,64,3,1
                    strItem = strItem & "|RDWCV|" & Val(Mid(strLine, 67, 3)) / (10 ^ 1)  'RDW-CV,67,3,1
                    strItem = strItem & "|HCT|" & Val(Mid(strLine, 70, 4)) / (10 ^ 4)    'HCT,70,4,2
                    strItem = strItem & "|PLT|" & Val(Mid(strLine, 74, 3))               'PLT,74,3,0
                    strItem = strItem & "|MPV|" & Val(Mid(strLine, 77, 3)) / (10 ^ 1)    'MPV,77,3,1
                    strItem = strItem & "|PDW|" & Val(Mid(strLine, 80, 3)) / (10 ^ 1)    'PDW,80,3,1
                    strItem = strItem & "|PCT|" & Val(Mid(strLine, 83, 4)) / (10 ^ 4)    'PCT,83,4,4
                    strItem = strItem & "|RDWSD|" & Val(Mid(strLine, 87, 3)) / (10 ^ 1)  'RDW-SD,87,3,1
                    StrWbc = Mid(strLine, 162, 795)
                    StrRbc = Mid(strLine, 957, 750)
                    StrPLt = Mid(strLine, 1707, 660)
                Case 2
                    strItem = strItem & "|WBC|" & Val(Mid(strLine, 36, 4)) / (10 ^ 1)    'WBC,36,4,1
                    strItem = strItem & "|LYMPHA|" & Val(Mid(strLine, 40, 4)) / (10 ^ 1) 'LYMPH#,40,4,1
                    strItem = strItem & "|MONA|" & Val(Mid(strLine, 44, 4)) / (10 ^ 1)   'Mon#,32,4,3
                    strItem = strItem & "|GRANA|" & Val(Mid(strLine, 48, 4)) / (10 ^ 1)  'Gran#,36,4,3
                    strItem = strItem & "|LYMPHP|" & Val(Mid(strLine, 52, 3)) / (10 ^ 1) 'Lymph%#,40,3,1
                    strItem = strItem & "|MONP|" & Val(Mid(strLine, 55, 3)) / (10 ^ 1)   'Mon%,43,3,1
                    strItem = strItem & "|GRANP|" & Val(Mid(strLine, 58, 3)) / (10 ^ 1)  'Gran%,46,3,1
                    strItem = strItem & "|RBC|" & Val(Mid(strLine, 61, 3)) / (10 ^ 2)    'RBC,49,3,2
                    strItem = strItem & "|HGB|" & Val(Mid(strLine, 64, 4)) / (10 ^ 1)    'HGB,52,4,1
                    strItem = strItem & "|MCHC|" & Val(Mid(strLine, 68, 3))              'MCHC,56,3,1
                    strItem = strItem & "|MCV|" & Val(Mid(strLine, 71, 4)) / (10 ^ 1)    'MCV,60,4,2
                    strItem = strItem & "|MCH|" & Val(Mid(strLine, 75, 4)) / (10 ^ 1)    'MCH,64,3,1
                    strItem = strItem & "|RDWCV|" & Val(Mid(strLine, 79, 3)) / (10 ^ 1)  'RDW-CV,67,3,1
                    strItem = strItem & "|HCT|" & Val(Mid(strLine, 82, 4)) / (10 ^ 4)    'HCT,70,4,2
                    strItem = strItem & "|PLT|" & Val(Mid(strLine, 86, 3))               'PLT,74,3,0
                    strItem = strItem & "|MPV|" & Val(Mid(strLine, 89, 3)) / (10 ^ 1)    'MPV,77,3,1
                    strItem = strItem & "|PDW|" & Val(Mid(strLine, 92, 3)) / (10 ^ 1)    'PDW,80,3,1
                    strItem = strItem & "|PCT|" & Val(Mid(strLine, 95, 4)) / (10 ^ 4)    'PCT,83,4,4
                    strItem = strItem & "|RDWSD|" & Val(Mid(strLine, 99, 3)) / (10 ^ 1)  'RDW-SD,87,3,1
                    
                    StrWbc = Mid(strLine, 162, 795)
                    StrRbc = Mid(strLine, 957, 750)
                    StrPLt = Mid(strLine, 1707, 660)
                End Select

            End If
            
            If strItem <> "" And Val(str�걾��) <> 0 Then
                strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                
                strResult = strResult & vbCrLf & "WBC;0;256,340,25-,60-100,95-,130-200,165-,200-300,235-,270-400" & BC3000(StrWbc)
                strResult = strResult & "^RBC;0;256,370,20-50,70-100,120-150,170-200,220-,270-300,320-" & BC3000(StrRbc)
                strResult = strResult & "^PLT;0;256,370,30-2,90-,140-20,190-,240-40,300-,360-60" & BC3000(StrPLt)
            End If
            
            If Not strReserved Like "**" & Chr(&H1A&) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        '�յ�����Ϣ��ȫ,����,���´��ٽ���
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    'strCmd = Chr(6)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:������Ҫ���������͵ı걾��Ϣ
    '           �걾֮����||�ָ�
    '           Ԫ��֮����|�ָ�
    '           ��0��Ԫ�أ�����ʱ��
    '           ��1��Ԫ�أ��������
    '           ��2��Ԫ�أ�������
    '           ��3��Ԫ�أ��걾
    '           ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '           ��5��9Ԫ�أ�ϵͳ����
    '           �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    '   iSendStep  : ����
    '   blnSuccess : �Ƿ���ʧ��
    '   strResponse: ������Ӧ����Ϣ
    '   blnUndo    : �Ƿ�ȡ��.
    '
    '   ����ֵ�������������ִ���ָ�
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function

Private Function BC3000(ByVal strIn As String) As String
    Dim strImg As String
    Dim lngI As Long
    
    For lngI = 0 To Len(strIn) / 3 - 1
        strImg = strImg & ";" & Mid(strIn, 3 * lngI + 1, 3)
    Next
    BC3000 = strImg
    
End Function

