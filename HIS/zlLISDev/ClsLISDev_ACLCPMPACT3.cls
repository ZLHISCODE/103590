VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_ACLCPMPACT3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
        
    '�豸��׼�ӿ�-----U200
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������^�Ƿ���^�����  ����ֻ���������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim IntStar As Integer
    Dim IntGetStar As Integer
    Dim IntGetEnd As Integer
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    Dim Intlen As Long
    Dim StrStar As String
    Dim StrRes As String
    Dim StrItem As String
    Dim StrItemRes As String
    Dim StrTime As String
    Dim StrNum As String
    Dim StrInspe As String
    Dim StrResultOne As String 'һ�������Ž���ַ���
    Dim StrWbc As String
    Dim StrRbc As String
    Dim StrPLt As String
    Dim adata() As String
    Dim aitem() As String
    Dim lngloop As Long
    Dim intloop As Integer
    Dim strData As String
    Dim strTmp As String
    
    
    
    IntGetStar = 0
    IntGetEnd = 0
    pos2 = 1
    
   Intlen = Len(strSource)
  
    If InStr(strSource, "") > 0 Then
        
        adata = Split(strSource, "")
        
        For lngloop = 1 To UBound(adata)
            
            aitem = Split(adata(lngloop), vbCr) 'vbCr �����س�
            
            'û�д�������ʱ�˳�
            If UBound(aitem) < 45 And lngloop = UBound(adata) Then
                If strData <> "" Then
                    strResult = strData
                End If
                strReserved = adata(lngloop)
                Exit Sub
            End If
            
             StrRes = ""

            For intloop = 13 To 34
                StrItem = ""
                StrItemRes = ""
                If intloop = 13 Then StrTime = Trim(Mid(aitem(intloop), 5, 15))
                If intloop = 15 Then StrNum = Trim(Mid(aitem(intloop), 6, 11))
                If intloop = 19 Or intloop = 21 Or intloop = 23 Or intloop = 24 Then
                    StrItem = Trim(Mid(aitem(intloop), 1, 6))
                    StrItemRes = Trim(Mid(aitem(intloop), 7, 6))
                End If
                If intloop = 16 Or intloop = 18 Or intloop = 22 Or intloop = 25 Or _
                   intloop = 26 Or intloop = 27 Or intloop = 28 Or intloop = 29 Or _
                   intloop = 30 Or intloop = 31 Or intloop = 32 Or intloop = 34 Then
                   StrItem = Trim(Mid(aitem(intloop), 1, 9))
                   StrItemRes = Trim(Mid(aitem(intloop), 10, 10))
                End If
                
                If StrItem <> "" Then StrRes = StrRes + StrItem + "|" + StrItemRes + "|"
                
            Next
               strResult = strResult + CStr(Now) + "|" + StrNum + "|" + " | |0|" + StrRes + "|"
        Next
    Else
        strReserved = strSource
    End If
    strResult = Left(strResult, Len(strResult) - 2) 'ȥ����������ַ�
    strCmd = Chr(&H6&)
End Sub



Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
End Function







