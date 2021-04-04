VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_RxL_Max_HM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:����
'�ͺ�:RxL_Max_HM
'����:
'����:����

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
    
    Dim iStrPoint As Integer '�ַ�ָ��
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '��ǰ������������ַ�λ��
    
    Dim varIn As Variant '�����������
    Dim StrTime As String
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strLine As String
    Dim str�걾�� As String, str���� As String, str�����˱걾�ʿ� As String, strָ�� As String
    Dim aLine() As String
    Dim strDate As String
    Dim intLoop As Integer
    Dim aItem() As String
    Dim strNO As String
    Dim strNow As String
    
    
    On Error GoTo errHandle
    
    '��ʼ������ֵ
    strReserved = "": strResult = "": strCmd = ""
    
    'û�п�ʼ�ͽ����ַ�ʱ�˳�
    If InStr(strSource, Chr(&H2)) = 0 Or InStr(strSource, Chr(&H3)) = 0 Then
        strReserved = strSource
        strCmd = Chr(&H6)
        Exit Sub
    End If
    
    strNow = Mid(strSource, InStr(strSource, Chr(&H2)), InStr(strSource, Chr(&H3)))
    strReserved = Mid(strSource, InStr(strSource, Chr(&H3)) + 1)
    
    '��ȡ�걾�ź�������Ϣ
    aLine = Split(strNow, Chr(&H1C))
    
    strData = "20" & Mid(aLine(7), 11, 2) & "-" & Mid(aLine(7), 9, 2) & "-" & Mid(aLine(7), 7, 2) & _
              " " & Mid(aLine(7), 5, 2) & ":" & Mid(aLine(7), 3, 2) & ":" & Mid(aLine(7), 1, 2) & _
              "|" & aLine(3) & "| |Ѫ��|0"
              
    '��ʼ�������
'    aLine = Split(strNow, Chr(&H1C) & Chr(&H1C))
    
    For intLoop = 11 To UBound(aLine) Step 4
        If intLoop + 4 <= UBound(aLine) Then
            strData = strData & "|" & aLine(intLoop) & "|" & IIf(aLine(intLoop + 1) = "", aLine(intLoop + 3), aLine(intLoop + 1))
        End If
    Next
    strResult = strData
    strCmd = Chr(&H6)
    Exit Sub
errHandle:
    WriteLog "RxL_Max_HM", Err.Description, ""

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







