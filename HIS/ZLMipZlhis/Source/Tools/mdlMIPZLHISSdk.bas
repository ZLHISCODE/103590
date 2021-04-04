Attribute VB_Name = "mdlMIPZLHISSdk"
Option Explicit

Public gstrDataPath As String
Private gstrSysName As String
Public gcnOracle As New ADODB.Connection

Public Function OpenDialog(ByRef objDlg As Object, ByVal strDialogTitle As String, ByVal strFilter As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Dim strTmp As String
            
    With objDlg
        .DialogTitle = strDialogTitle
        .Filter = strFilter
    
        On Error Resume Next
    
        .Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowOpen
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            OpenDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    MsgBox "���ܴ��ļ�(" & strTmp & "),���ļ���������ʹ�û��ļ�������!"
End Function


Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
