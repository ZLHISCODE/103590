VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AA7000W"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:΢��Ԫ�ط����� AA7000W
'����:����
'����:����
'�û�:��ͩҽԺ
Private mblnStart As Boolean '��ʼ
Private mblnEnd As Boolean   '����
Private mstrData As String '����������

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
    
    Dim lng_Stx As String, lng_Etx As Long, str_Code As String
    Dim varIn As Variant '�����������
    Dim strLine As String
    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String, strֵ As String
    Dim strҪ��ȡָ�� As String, varҪ��ȡָ�� As Variant
    Dim iStrPoint As Integer
    Dim aData() As String, aItem() As String
    Dim intLoop  As Integer, intItem As Integer
    Dim strData As String
    
    
'    aData = Split(strSource, ",")
'    If UBound(aData) < 2 Then
'        strResult = strSource
'    End If
'
'    For intLoop = 0 To UBound(aData)
'        If intItem < 4 Then
'            strData = strData & "," & aData(intLoop)
'            intItem = intItem + 1
'        Else
'            strData = strData & ";" & aData(intLoop)
'            intItem = 1
'        End If
'    Next
    
    If InStr(strSource, ";") <= 0 Then
        strResult = strSource
    End If
    
    aData = Split(strSource, ";")
    For intLoop = 0 To UBound(aData)
        If aData(intLoop) <> "" Then
            aItem = Split(aData(intLoop), ",")
            If UBound(aItem) >= 3 And Len(Trim(aItem(3))) > 0 Then
                strResult = strResult & "||" & Now & "|" & aItem(0) & "| |ѪҺ|0|" & aItem(3) & "|" & aItem(4)
            Else
                strReserved = strReserved & "," & aData(intLoop)
            End If
        Else
            strReserved = strReserved & "," & aData(intLoop)
        End If
    Next
    strReserved = Mid(strReserved, 2)
    strResult = Mid(strResult, 3)
    
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



