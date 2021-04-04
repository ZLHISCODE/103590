Attribute VB_Name = "mdlOLED"
Option Explicit

Public Function CheckOLEDBDriver(Optional ByRef strErr As String) As Boolean
'���ܣ����OLEDB�����Ƿ�ע�ᣬû��ע�����Զ�ע��
    Dim objFSO          As New FileSystemObject
    Dim strOLEDB        As String, i            As Integer
    Dim strOracleHome   As String, strCLSID     As String
    Dim blnOk           As Boolean

    On Error Resume Next
    strErr = ""
    If GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Classes\OraOLEDB.Oracle\CLSID", "", strCLSID) Then
        blnOk = strCLSID <> ""
    End If
    If Not blnOk Then
        Call gobjTrace.OpenTace("OO4O", gstrAPPPath)
        On Error GoTo ErrH
        'OracleHOme��ȡ
        strOracleHome = GetOracleHome()
        If strOracleHome = "" Then
            strErr = "δ�ҵ�32λORACLE�ͻ��˰�װ��Ϣ"
            Exit Function
        End If
        If objFSO.FileExists(strOracleHome & "\Bin\OraOLEDB.dll") Then
            strOLEDB = strOracleHome & "\Bin\OraOLEDB.dll"
        Else
            For i = 8 To 12
                If objFSO.FileExists(strOracleHome & "\Bin\OraOLEDB" & i & ".dll") Then
                    strOLEDB = strOracleHome & "\Bin\OraOLEDB" & i & ".dll"
                    Exit For
                End If
            Next
        End If
        If strOLEDB = "" Then
            strErr = "δ�ҵ�OLEDB�����ļ�(" & strOracleHome & "\Bin\OraOLEDB*.dll)����ǰ�ͻ��˿���δ��װ"
        Else
            If Not gclsRegCom.RegCom(strOLEDB, RFT_NormalReg, strErr) Then
                strErr = "OLEDB�����ļ�(" & strOLEDB & ")ע��ʧ��"
            Else
                strCLSID = ""
                If GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Classes\OraOLEDB.Oracle\CLSID", "", strCLSID) Then
                    CheckOLEDBDriver = True
                End If
            End If
        End If
    Else
        CheckOLEDBDriver = True
    End If
    Exit Function
ErrH:
    strErr = "OLEDB�������ʧ�ܣ�������Ϣ��" & Err.Description
    Err.Clear
End Function
