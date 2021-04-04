Attribute VB_Name = "mdlPubInterface"
Option Explicit

Public Function CheckGrantKey(ByVal cnOracle As ADODB.Connection, ByVal strKey As String, Optional ByRef strErrNote As String) As Boolean
'功能：根据连接检查授权码是否合法
'参数：cnOracle-连接对象
'      strKey-授权码
'返回：True-合法，False-不合法
    Dim strSQL      As String
    Dim rstmp       As ADODB.Recordset
    On Error GoTo errh
    strSQL = "Select Key, To_Char(Starttime, 'YYYY-MM-DD hh24:mi:ss') Starttime, To_Char(Stoptime, 'YYYY-MM-DD hh24:mi:ss') Stoptime," & vbNewLine & _
            "       To_Char(Sysdate, 'YYYY-MM-DD hh24:mi:ss') Curtime, State" & vbNewLine & _
            "From Zlinterface" & vbNewLine & _
            "Where Key=[1]"
    Set rstmp = OpenSQLRecord(cnOracle, strSQL, "CheckGrantKey", Sm4EncryptEcb(strKey, GetGeneralAccountKey(G_APP_KEY)))
    If rstmp.EOF Then
        strErrNote = "无该授权码。"
    Else
        If Val(rstmp!State & "") = 1 Then
            strErrNote = "授权码已经停用。"
        ElseIf Not IsNull(rstmp!Stoptime) Then
            If rstmp!Curtime & "" < rstmp!Starttime & "" Then
                strErrNote = "授权码尚未生效（生效时间：" & rstmp!Starttime & "）。"
            ElseIf rstmp!Curtime & "" > rstmp!Stoptime & "" Then
                strErrNote = "授权码已经过期（过期时间：" & rstmp!Stoptime & "）。"
            Else
                CheckGrantKey = True
            End If
        Else
            CheckGrantKey = True
        End If
    End If
    Exit Function
errh:
    strErrNote = "授权码校验失败，(" & Err.Description & ")" & Err.Description
    Err.Clear
End Function

Public Function GetZLInterfacePWD(ByVal cnOracle As ADODB.Connection, Optional ByRef strErrNote As String) As String
    Dim strSQL  As String, strErr       As String
    Dim rstmp   As ADODB.Recordset

    On Error GoTo errh
    strSQL = "Select Max(内容) 内容 From zlRegInfo A Where a.项目 = [1]"
    Set rstmp = OpenSQLRecord(cnOracle, strSQL, "GetZLInterfacePWD", "三方接口密码")
    If Trim(rstmp!内容 & "") <> "" Then
        GetZLInterfacePWD = Sm4DecryptEcb(rstmp!内容 & "", GetGeneralAccountKey(G_INTERFACE_KEY))
        If GetZLInterfacePWD = "" Then
            strErrNote = "三方接口密码获取失败（登录服务器管理工具三方授权管理进行账户修复）。"
        End If
    Else
        strErrNote = "三方接口密码获取失败（登录服务器管理工具三方授权管理进行账户修复）。"
    End If
    Exit Function
errh:
    strErrNote = "获取ZLInterface密码获取失败失败（登录服务器管理工具三方授权管理进行账户修复）。(" & Err.Number & ")" & Err.Description
    Err.Clear
End Function
'
