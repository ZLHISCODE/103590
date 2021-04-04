VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Dimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCode As String
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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant
    Dim str_�걾�� As String, str_���� As String, str_��Ŀ As String, str_��� As String, i As Integer
    
    mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    
    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '---------------------------------------------------
        var_Code = Split(str_Code, Chr(&H1C))
        str_�걾�� = Val(var_Code(3))
        str_���� = var_Code(7)
        If Len(str_����) >= 12 And IsNumeric(str_����) Then
            str_���� = "20" & Mid(str_����, 11, 2) & "-" & Mid(str_����, 9, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 5, 2) & ":" & Mid(str_����, 3, 2)
        Else
            str_���� = Format(Now, "yyyy-MM-dd HH:mm")
        End If
        str_��Ŀ = ""
        
        If UBound(var_Code) > 11 Then
            For i = 11 To UBound(var_Code)
                If InStr(",TBIL,TP,XADA,XCK,TGL,CHOL,ALP,AST,URCA,XALT,DBIL,GLU,ALB,BUN,LDH,GGT,CREA,", "," & Trim(var_Code(i)) & ",") > 0 And Trim(var_Code(i)) <> "" Then
                    str_��� = Trim(var_Code(i + 1))
                    If str_��� <> "" Then
                        str_��Ŀ = str_��Ŀ & "|" & Trim(var_Code(i)) & "|" & str_���
                    End If
                End If
            Next
        End If
        If Val(str_�걾��) > 0 And IsDate(str_����) And str_��Ŀ <> "" Then
            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |Ѫ��|0" & str_��Ŀ
        End If
        
        '---------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop
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
