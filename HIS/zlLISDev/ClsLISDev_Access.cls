VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_Access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstr_Tmp As String '��ʱӦ��

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
    Dim lngLoop As Long
    Dim Intlen As Integer
    
    Dim StrStar As String
    Dim StrRes As String
    Dim strItem As String
    Dim StrItemRes As String
    Dim StrTime As String
    Dim StrNum As String
    Dim StrInspe As String
    Dim StrResultOne As String 'һ�������Ž���ַ���
    Dim StrWbc As String
    Dim StrRbc As String
    Dim StrPLt As String
    Dim aData() As String
    Dim aitem() As String
    Dim aResult() As String
    
   strReserved = ""
   strResult = ""
    
   Intlen = Len(strSource)
  
'    If Intlen = 1 Then '���ݷ����ַ��ж��Ƿ�Ҫ����ѯ���ַ�
'       strCmd = Chr(&H6)
'       Exit Sub
'    End If
    If Replace(strSource, mstr_Tmp, "") Like "***" & Chr(&HD) & Chr(&HA) & "*" Then
        strCmd = Chr(&H6)
        mstr_Tmp = mstr_Tmp & strSource
    ElseIf Replace(strSource, mstr_Tmp, "") Like "*" & Chr(&H5) & "*" Then
        strCmd = Chr(&H6)
        mstr_Tmp = mstr_Tmp & strSource
    End If
    
    If InStr(strSource, "O|") > 0 And InStr(strSource, "L|1") Then
        aData = Split(strSource, "O|")    '��������
        For i = 1 To UBound(aData)
            aitem = Split(aData(i), "R|") '��Ŀ����
            If UBound(aitem) < 1 Then
                strReserved = strSource
                Exit Sub
            End If
            StrNum = Trim(Mid(aData(i), InStr(aData(i), "1|") + 2, InStr(aData(i), "|^") - InStr(aData(i), "1|") - 2))
            StrRes = ""
            For j = 1 To UBound(aitem)
                aResult = Split(aitem(j), "|") '�������
                If InStr(aResult(1), "^^^") + 3 > 0 And InStr(aResult(1), "^1") > 0 Then
                    strItem = Mid(aResult(1), InStr(aResult(1), "^^^") + 3, InStr(aResult(1), "^1") - InStr(aResult(1), "^^^") - 3) 'һ����Ŀ���
                    StrItemRes = aResult(2)
                    StrRes = StrRes + "|" & strItem + "|" + StrItemRes
                End If
            Next
            If StrRes <> "" Then
                strResult = strResult + "||" + Format(Now, "yyyy-MM-dd") + "|" + StrNum + "|" + " | |0|" + Mid(StrRes, 2)
            End If
            
        Next
        
    Else
        strReserved = strSource
    End If
    If strResult <> "" Then strResult = Mid(strResult, 3)

    
    
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






