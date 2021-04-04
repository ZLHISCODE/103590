VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_RT_2204C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Ѫ�������� RT-2204C
'����:Rayto(�׶�)
'����:�¶�
'�û�:�뽭����ҽԺ

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
    
    Dim strLine As String, varLine As Variant, var��Ŀ As Variant, strData As String
    Dim str_�걾�� As String, str_���� As String, str_��Ŀ As String
    Dim intRow As Integer, strTmp As String, strR As String
    
    strData = strSource
    strResult = ""
    
    If strData Like "*R,*" & Chr(&HA) & "*" Then
        
        varLine = Split(strData, Chr(&HA))
        str_�걾�� = "": str_���� = "": str_��Ŀ = ""
        
        For intRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(intRow)
            var��Ŀ = Split(strLine, ",")
            
            If UBound(var��Ŀ) >= 17 Then
                str_���� = var��Ŀ(8)
                str_�걾�� = var��Ŀ(1)
                str_��Ŀ = var��Ŀ(9)
                
                If str_��Ŀ = "PT" Then
                    str_��Ŀ = "|PT|" & Val(var��Ŀ(12))
                ElseIf str_��Ŀ = "DFIB" Then
                    str_��Ŀ = "|DFIB|" & Val(var��Ŀ(14))
                ElseIf str_��Ŀ = "APTT" Then
                    str_��Ŀ = "|APTT|" & Val(var��Ŀ(11))
                ElseIf str_��Ŀ = "TT" Then
                    str_��Ŀ = "|APTT|" & Val(var��Ŀ(11))
                End If
                
                If IsDate(str_����) And Val(str_�걾��) > 0 And str_��Ŀ <> "" Then
                    strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |ѪҺ|0" & str_��Ŀ
                End If
            Else
                If strLine Like "R,*" Or strLine = "R" Then
                    strReserved = strReserved & strLine
                End If
            End If
        Next
    End If
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   �걾֮����||�ָ�
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function