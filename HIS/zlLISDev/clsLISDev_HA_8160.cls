VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HA_8160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ADAMS~(TM)A1 c HA-8160ȫ�Զ��ǻ�Ѫ�쵰�׷�����
'����:
'����:ŷ��ʤ
'�û�:�񺽹���ҽԺ
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
        Dim lngS As Long, lngE As Long, strLine As String
        Dim varTmp As Variant, intloop As Integer, intCount As Integer
        Dim str_��� As String, strSampleNo As String, str_���� As String, strItem As String
        Dim var_Code As Variant
        On Error GoTo hErr
    
        mstrCode = mstrCode & strSource
        lngE = InStr(mstrCode, "")
        strResult = ""
        strReserved = ""
        'strDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
     If lngE > 0 Then
        var_Code = Split(mstrCode, ",")
        str_���� = var_Code(4) & var_Code(5)
        str_���� = Mid(str_����, 2, 10) & " " & Mid(str_����, 13, 5)
        If var_Code(7) = "K0" Then
            strSampleNo = Val(Mid(var_Code(6), 2))
        ElseIf var_Code(7) = "K4" Then
'            If var_Code(9) Like "P9*" Then
'                strSampleNo = Val(Mid(var_Code(9), 2))
'            End If
            strSampleNo = Val(Mid(var_Code(9), 2))
        Else
            strSampleNo = Val(Mid(var_Code(6), 2))
        End If
        str_��� = Val(var_Code(11))
        strResult = str_���� & "|" & strSampleNo & "| |ѪҺ|0" & "|" & "HbA1c" & "|" & str_���
    End If
      If strResult <> "" Then
        mstrCode = ""
    End If
        Exit Sub
       
hErr:
        WriteLog "HA8180 Line:" & CStr(Erl()), mstrCode, strResult & vbNewLine & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String

    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function









