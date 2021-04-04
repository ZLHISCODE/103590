Attribute VB_Name = "mdlService"
Option Explicit

Public gobjService As Object

Public Function InitSvr() As Boolean
'功能：初始化服务接口部件
    If gobjService Is Nothing Then
        On Error Resume Next
        Set gobjService = CreateObject("zlServiceCall.clsServiceCall")
        If Not gobjService.InitService(gcnOracle, gstrDBUser, glngSys) Then
            Set gobjService = Nothing
        End If
        Err.Clear: On Error GoTo 0
    End If
    If gobjService Is Nothing Then
        MsgBox "zlServiceCall.clsServiceCall创建失败!", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not gobjService Is Nothing Then InitSvr = True
End Function

Public Function CallService(ByVal strServiceName As String, _
    ByVal strJson_In As String, Optional ByRef strJson_out As String, Optional ByVal strTittle As String, _
    Optional lngModule As Long, Optional blnShowErrMsg As Boolean = True, Optional ByVal strAskDate As String, _
    Optional varExpend As String, Optional lngSys As Long, Optional blnReadServiceErr As Boolean) As Boolean
'功能：调用服务
'相关说明见 zlServiceCall.clsServiceCall.CallService 接口
    If InitSvr() Then
        If Not gobjService.CallService(strServiceName, strJson_In, strJson_out, strTittle, lngModule, blnShowErrMsg, strAskDate, varExpend, lngSys, blnReadServiceErr) Then Exit Function
        If Not blnShowErrMsg Then
            If gobjService.GetJsonNodeValue("output.code") & "" = "0" Then
                varExpend = gobjService.GetJsonNodeValue("output.message")
                CallService = False: Exit Function
            End If
        End If
        CallService = True
    End If
End Function

