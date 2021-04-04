Attribute VB_Name = "mdlPubInterface"
Option Explicit

Public Function CheckGrantKey(ByVal cnOracle As ADODB.Connection, ByVal strKey As String, Optional ByRef strErrNote As String) As Boolean
'���ܣ��������Ӽ����Ȩ���Ƿ�Ϸ�
'������cnOracle-���Ӷ���
'      strKey-��Ȩ��
'���أ�True-�Ϸ���False-���Ϸ�
    Dim strSQL      As String
    Dim rstmp       As ADODB.Recordset
    On Error GoTo errh
    strSQL = "Select Key, To_Char(Starttime, 'YYYY-MM-DD hh24:mi:ss') Starttime, To_Char(Stoptime, 'YYYY-MM-DD hh24:mi:ss') Stoptime," & vbNewLine & _
            "       To_Char(Sysdate, 'YYYY-MM-DD hh24:mi:ss') Curtime, State" & vbNewLine & _
            "From Zlinterface" & vbNewLine & _
            "Where Key=[1]"
    Set rstmp = OpenSQLRecord(cnOracle, strSQL, "CheckGrantKey", Sm4EncryptEcb(strKey, GetGeneralAccountKey(G_APP_KEY)))
    If rstmp.EOF Then
        strErrNote = "�޸���Ȩ�롣"
    Else
        If Val(rstmp!State & "") = 1 Then
            strErrNote = "��Ȩ���Ѿ�ͣ�á�"
        ElseIf Not IsNull(rstmp!Stoptime) Then
            If rstmp!Curtime & "" < rstmp!Starttime & "" Then
                strErrNote = "��Ȩ����δ��Ч����Чʱ�䣺" & rstmp!Starttime & "����"
            ElseIf rstmp!Curtime & "" > rstmp!Stoptime & "" Then
                strErrNote = "��Ȩ���Ѿ����ڣ�����ʱ�䣺" & rstmp!Stoptime & "����"
            Else
                CheckGrantKey = True
            End If
        Else
            CheckGrantKey = True
        End If
    End If
    Exit Function
errh:
    strErrNote = "��Ȩ��У��ʧ�ܣ�(" & Err.Description & ")" & Err.Description
    Err.Clear
End Function

Public Function GetZLInterfacePWD(ByVal cnOracle As ADODB.Connection, Optional ByRef strErrNote As String) As String
    Dim strSQL  As String, strErr       As String
    Dim rstmp   As ADODB.Recordset

    On Error GoTo errh
    strSQL = "Select Max(����) ���� From zlRegInfo A Where a.��Ŀ = [1]"
    Set rstmp = OpenSQLRecord(cnOracle, strSQL, "GetZLInterfacePWD", "�����ӿ�����")
    If Trim(rstmp!���� & "") <> "" Then
        GetZLInterfacePWD = Sm4DecryptEcb(rstmp!���� & "", GetGeneralAccountKey(G_INTERFACE_KEY))
        If GetZLInterfacePWD = "" Then
            strErrNote = "�����ӿ������ȡʧ�ܣ���¼������������������Ȩ��������˻��޸�����"
        End If
    Else
        strErrNote = "�����ӿ������ȡʧ�ܣ���¼������������������Ȩ��������˻��޸�����"
    End If
    Exit Function
errh:
    strErrNote = "��ȡZLInterface�����ȡʧ��ʧ�ܣ���¼������������������Ȩ��������˻��޸�����(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function
'
