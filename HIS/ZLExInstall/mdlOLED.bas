Attribute VB_Name = "mdlOLED"
Option Explicit

Public Function CheckOLEDBDriver(Optional ByRef strErr As String) As Boolean
'功能：检查OLEDB驱动是否注册，没有注册则自动注册
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
        'OracleHOme获取
        strOracleHome = GetOracleHome()
        If strOracleHome = "" Then
            strErr = "未找到32位ORACLE客户端安装信息"
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
            strErr = "未找到OLEDB驱动文件(" & strOracleHome & "\Bin\OraOLEDB*.dll)，当前客户端可能未安装"
        Else
            If Not gclsRegCom.RegCom(strOLEDB, RFT_NormalReg, strErr) Then
                strErr = "OLEDB驱动文件(" & strOLEDB & ")注册失败"
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
    strErr = "OLEDB驱动检测失败，错误信息：" & Err.Description
    Err.Clear
End Function
