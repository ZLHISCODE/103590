VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BC1800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:BC-1800
'����:��������
'����:ŷ��ʤ
'�û�:��ݸ�߈�ҽԺ

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
 
        On Error GoTo Analyse_Err

  
        Dim strItem As String, str���� As String, str�걾�� As String
        Dim str�����˱걾�ʿ� As String, strLine As String
        Dim lngLength As Long
        Dim blnEnd As Boolean '�Ƿ��ѽ�����
   
    
        '��ʼ������ֵ
100     strResult = "": strReserved = "": strCmd = ""
102     If strSource Like "**" & Chr(&H1A&) & "*" Then
104         blnEnd = False
106         strReserved = strSource
108         Do While Not blnEnd
                'ȡһ�������ļ�¼,ʣ��Ĵ浽strReserved��,�´ν���
110             lngLength = InStr(strReserved, Chr(&H1A&)) - InStr(strReserved, "") + 1
112             strLine = Mid(strReserved, InStr(strReserved, ""), lngLength)
114             strReserved = Mid(strReserved, InStr(strReserved, Chr(&H1A&)) + 1)
            
116             str���� = "": str�걾�� = "": str�����˱걾�ʿ� = "": strItem = ""
                '����
            
118             strItem = ""
120             If strLine Like "A*" Then
122                 str���� = Mid(strLine, 16, 4) & "-" & Mid(strLine, 12, 2) & "-" & Mid(strLine, 14, 2) & " " & Mid(strLine, 20, 2) & ":" & Mid(strLine, 22, 2)
124                 str�걾�� = Val(Mid(strLine, 3, 8))
126                 str�����˱걾�ʿ� = " |ѪҺ|0"
128                 strItem = strItem & "|WBC|" & Val(Mid(strLine, 24, 4)) / (10 ^ 1)    'WBC,24,4,3
130                 strItem = strItem & "|LYMPHA|" & Val(Mid(strLine, 28, 4)) / (10 ^ 1) 'LYMPH#,28,4,3
132                 strItem = strItem & "|MIDA|" & Val(Mid(strLine, 32, 4)) / (10 ^ 1)   'Mid#,32,4,3
134                 strItem = strItem & "|GRANA|" & Val(Mid(strLine, 36, 4)) / (10 ^ 1)  'Gran#,36,4,3
136                 strItem = strItem & "|LYMPHP|" & Val(Mid(strLine, 40, 3)) / (10 ^ 1) 'Lymph%#,40,3,1
138                 strItem = strItem & "|MIDP|" & Val(Mid(strLine, 43, 3)) / (10 ^ 1)   'Mid%,43,3,1
140                 strItem = strItem & "|GRANP|" & Val(Mid(strLine, 46, 3)) / (10 ^ 1)  'Gran%,46,3,1
142                 strItem = strItem & "|RBC|" & Val(Mid(strLine, 49, 3)) / (10 ^ 2)    'RBC,49,3,2
144                 strItem = strItem & "|HGB|" & Val(Mid(strLine, 52, 4)) / (10 ^ 1)    'HGB,52,4,1
146                 strItem = strItem & "|MCHC|" & Val(Mid(strLine, 56, 3))              'MCHC,56,4,1
148                 strItem = strItem & "|MCV|" & Val(Mid(strLine, 59, 4)) / (10 ^ 1)    'MCV,60,4,2
150                 strItem = strItem & "|MCH|" & Val(Mid(strLine, 63, 4)) / (10 ^ 1)    'MCH,64,3,1
152                 strItem = strItem & "|RDWCV|" & Val(Mid(strLine, 67, 3)) / (10 ^ 1)  'RDW-CV,67,3,1
154                 strItem = strItem & "|HCT|" & Val(Mid(strLine, 70, 4)) / (10 ^ 4)    'HCT,70,4,2
156                 strItem = strItem & "|PLT|" & Val(Mid(strLine, 74, 3))               'PLT,74,3,0
158                 strItem = strItem & "|MPV|" & Val(Mid(strLine, 77, 3)) / (10 ^ 1)    'MPV,77,3,1
160                 strItem = strItem & "|PDW|" & Val(Mid(strLine, 80, 3)) / (10 ^ 1)    'PDW,80,3,1
162                 strItem = strItem & "|PCT|" & Val(Mid(strLine, 83, 4)) / (10 ^ 4)    'PCT,83,4,4
164                 strItem = strItem & "|RDWSD|" & Val(Mid(strLine, 87, 3)) / (10 ^ 1)  'RDW-SD,87,3,1
                End If
            
166             If strItem <> "" And Val(str�걾��) <> 0 Then
168                 strResult = strResult & "||" & str���� & "|" & str�걾�� & "|" & str�����˱걾�ʿ� & strItem
                End If
            
170             If Not strReserved Like "**" & Chr(&H1A&) & "*" Then
172                 blnEnd = True
                End If
            Loop
        Else
            '�յ�����Ϣ��ȫ,����,���´��ٽ���
174         strReserved = strSource
        End If
176     If strResult <> "" Then
178         If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
        End If

        Exit Sub

Analyse_Err:
    If Error = 1 Then
        Resume
    End If
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



