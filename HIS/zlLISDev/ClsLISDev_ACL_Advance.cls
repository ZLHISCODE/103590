VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_ACL_Advance"
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
    Dim Intpos1 As Integer
    Dim Intpos2 As Integer
    Dim Intpos3 As Integer
    Dim lngloop As Long
    Dim Intlen As Integer
    
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
    Dim aData() As String
    Dim aItem() As String
    Dim aResult() As String
    Dim aNum() As String
    
   strReserved = ""
   strResult = ""
    
   Intlen = Len(strSource)
  
    If Intlen = 1 Then '���ݷ����ַ��ж��Ƿ�Ҫ����ѯ���ַ�
       If Asc(strSource) = &H3 Then strCmd = &H6
       Exit Sub
    End If
      
    If InStr(strSource, "O|") > 0 Then
        aData = Split(strSource, "O|")    '��������
        For i = 1 To UBound(aData)
            aItem = Split(aData(i), "R|") '��Ŀ����
            If UBound(aItem) < 4 Then
               strReserved = strSource
               Exit Sub
            End If
            aNum = Split(aItem(0), "|")
            StrNum = Trim(Mid(aNum(1), 2, Len(aNum(1)) - 1))
            StrTime = Trim(Mid(aItem(1), 1, 12))
            StrRes = ""
            For j = 2 To 4
                StrItem = Trim(Mid(aItem(j), InStr(aItem(j), "|^^^") + 4, 3) + "-" + Trim(Str(j))) 'һ����Ŀ�������
                aResult = Split(aItem(j), "|") '�������
                StrItemRes = Trim(aResult(2))
                StrRes = StrRes + StrItem + "|" + StrItemRes + "|"
            Next
            strResult = strResult + StrTime + "|" + StrNum + "|" + " | |0|" + StrRes + "|"
            
        Next
        
    Else
        strReserved = strSource
    End If
  
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







