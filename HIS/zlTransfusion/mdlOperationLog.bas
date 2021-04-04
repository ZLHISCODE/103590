Attribute VB_Name = "mdlOperationLog"
Option Explicit
'-- ������Һ������־ģ��
Public Enum OPERTYPE
    QUEUE = 1       '�ŶӲ�����־
    MEDICAL = 2     'ҽ��
    CALLS = 3       '����
    SEAT = 4        '��λ
End Enum

Public Sub SaveOperLog(ByVal lngDeptID As Long, ByVal varNO As Variant, ByVal Oper As OPERTYPE, ByVal strLogInfo As String)
'���ܣ�������־
'������
'  lngDeptID��ִ�п���ID
'  varNO��0_����ID_�Һŵ��ţ������1_����ID_��ҳID���������ۣ����� ���˶���(cPatient)
'  Oper����������
'  strLogInfo����־����
'���أ�

    Dim strSQL As String, strBillNO As String
    Dim lngID As Long, lngPatiID As Long, lngPageID As Long
    Dim strNO As String
    Dim objPati As cPatient
    
    If UCase(TypeName(varNO)) = "CPATIENT" Then
        Set objPati = varNO
        If Not objPati Is Nothing Then
            If objPati.������Դ = 1 Then
                strNO = "1_" & objPati.Key
            Else
                strNO = "0_" & objPati.����ID & "_" & objPati.�Һŵ�
            End If
        End If
        strNO = strNO & "__"
    Else
        strNO = varNO & "__"
    End If
    
    If Val(strNO) = 0 Then
        '����
        lngPatiID = Val(Split(strNO, "_")(1))
        strBillNO = Trim(Split(strNO, "_")(2))
    Else
        '��������
        lngPatiID = Val(Split(strNO, "_")(1))
        lngPageID = Val(Split(strNO, "_")(2))
    End If
    
    On Error GoTo hErr
    
    lngID = zldatabase.GetNextId("������Һ������־")
    '--1-�ŶӲ�����־ 2��ҽ��������־ 3-���в�����־ 4-��λ������־
    strLogInfo = DelInvalidChar(strLogInfo, "%|""?")
    
    'strSQL = "ZL_������Һ������־_Add(" & lngID & "," & lngDeptID & ",'" & strNO & "'," & Oper & ",'" & strLogInfo & "','" & UserInfo.�û��� & "')"
    strSQL = "ZL_������Һ������־_Add(" & lngID & "," & IIf(lngDeptID = 0, "null", lngDeptID) & "," & _
                IIf(lngPatiID <= 0, "Null", lngPatiID) & "," & _
                IIf(strBillNO = "", "Null", "'" & strBillNO & "'") & "," & _
                IIf(lngPageID <= 0, "Null", lngPageID) & "," & _
                Oper & ",'" & strLogInfo & "','" & UserInfo.�û��� & "')"
    Call zldatabase.ExecuteProcedure(strSQL, "���������־")
    Exit Sub
    
hErr:
    SaveErrLog
End Sub
