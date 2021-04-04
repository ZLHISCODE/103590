VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AFT500AU_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:AFT-500AU ����ʷ�����
'����:÷�ݿ���
'����:���
'�û�:�ɶ��»�

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
   
    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim mstrCode As String
    Dim str_Code As String 'һ��������¼
    Dim str_Line As String 'һ������
    Dim str_�걾�� As String, str_���� As String, str_������Ŀ As String
    Dim varTmp() As String
    strResult = ""
    
    If Len(strSource) < 33 * 3 Then
        strReserved = strSource
        Exit Sub
    Else
        str_Code = Mid(strSource, 2)
    End If
    
    While str_Code <> ""
    
        varTmp = Split(str_Code, ",")
        str_���� = Format(varTmp(4) & "-" & varTmp(3) & "-" & varTmp(2) & " " & varTmp(5) & ":" & varTmp(6) & ":" & varTmp(1), "yyyy-MM-dd hh:mm:ss")
        str_�걾�� = CInt("&H" & varTmp(11))
        str_������Ŀ = "|K|" & CalcResult(varTmp(13), varTmp(14)) & "|Na|" & CalcResult(varTmp(16), varTmp(17)) & "|Cl|" & _
                        CalcResult(varTmp(19), varTmp(20)) & "|nCa|" & CalcResult(varTmp(22), varTmp(23)) & "|tCa|" & CalcResult(varTmp(28), varTmp(29))
        
        If Val(str_�걾��) <> 0 And IsDate(str_����) And str_������Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |ѪҺ|0" & str_������Ŀ
        End If
        
        If Len(mstrCode) < 33 * 3 Then
            strReserved = mstrCode
            str_Code = ""
        Else
            str_Code = Mid(mstrCode, 2)
            mstrCode = ""
        End If
    Wend
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Debug.Print strResult
    
    
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    SendSample = ""

End Function

Public Function CalcResult(ByVal strH As String, ByVal strL As String) As String
    CalcResult = Round((CDbl("&h" & strH) * 256 + CDbl("&H" & strL)) / 256, 2)
End Function



