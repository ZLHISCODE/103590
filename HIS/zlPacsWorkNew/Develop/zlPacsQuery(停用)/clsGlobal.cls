VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjCall As Object

Public Function ExecuteSql(ByVal strSql As String, ByVal strTitle As String, _
    ParamArray arrInput() As Variant) As ADODB.Recordset
    
    Dim varPars() As Variant
    
    varPars = arrInput
    
    Set ExecuteSql = mdlAdo.ExecuteCore(strSql, strTitle, varPars, False, False)
End Function


Public Sub ExecuteCmd(ByVal strSql As String, ByVal strTitle As String)
'执行存储过程
    Dim varPars() As Variant
    
    Call mdlAdo.ExecuteCore(strSql, strTitle, varPars, True)
End Sub

Private Function GetMethodName(strScript As String) As String
'获取可运行的脚本语句
    Dim lngIndex As Long
        
    GetMethodName = ""
    
    lngIndex = InStr(UCase(strScript), "FUNCTION")

    If lngIndex <= 0 Then
        lngIndex = InStr(UCase(strScript), "SUB")
        '设置过程名的开始位置
        lngIndex = lngIndex + 4
    Else
        lngIndex = lngIndex + 9
    End If
    
    GetMethodName = Mid(strScript, lngIndex, InStr(strScript, "(") - lngIndex)
    
End Function

Public Function ExecuteScript(strScript As String, ParamArray arrInput() As Variant)
'调用vbs脚本实现功能
    Dim strMethodName As String
    
On Error GoTo ErrorHnad
    
    '创建脚本执行对象
    If mobjCall Is Nothing Then
        Set mobjCall = CreateObject("ScriptControl")
        
        mobjCall.Timeout = 60000
        mobjCall.Language = "vbscript"
    End If
    
    mobjCall.Reset

'    Call objCall.addobject("Me", Me, True)
'    Call objCall.addobject("zlDatabase", zlDatabase, True)
'    Call objCall.addobject("zlCommFun", zlCommFun, True)
'    Call objCall.addobject("zlControl", zlControl, True)
    
    strMethodName = GetMethodName(strScript)
    
    Call mobjCall.AddCode(strScript)
    
    Select Case UBound(arrInput)
        Case -1
            ExecuteScript = mobjCall.Run(Trim(strMethodName))
        Case 0
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0))
        Case 1
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1))
        Case 2
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1), _
                                                            arrInput(2))
        Case 3
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1), _
                                                            arrInput(2), _
                                                            arrInput(3))
        Case 4
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1), _
                                                            arrInput(2), _
                                                            arrInput(3), _
                                                            arrInput(4))
        Case 5
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1), _
                                                            arrInput(2), _
                                                            arrInput(3), _
                                                            arrInput(4), _
                                                            arrInput(5))
        Case 6
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1), _
                                                            arrInput(2), _
                                                            arrInput(3), _
                                                            arrInput(4), _
                                                            arrInput(5), _
                                                            arrInput(6))
        Case 7
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1), _
                                                            arrInput(2), _
                                                            arrInput(3), _
                                                            arrInput(4), _
                                                            arrInput(5), _
                                                            arrInput(6), _
                                                            arrInput(7))
        Case 8
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1), _
                                                            arrInput(2), _
                                                            arrInput(3), _
                                                            arrInput(4), _
                                                            arrInput(5), _
                                                            arrInput(6), _
                                                            arrInput(7), _
                                                            arrInput(8))
        Case 9
            ExecuteScript = mobjCall.Run(Trim(strMethodName), arrInput(0), _
                                                            arrInput(1), _
                                                            arrInput(2), _
                                                            arrInput(3), _
                                                            arrInput(4), _
                                                            arrInput(5), _
                                                            arrInput(6), _
                                                            arrInput(7), _
                                                            arrInput(8), _
                                                            arrInput(9))
    End Select
    
    Exit Function
ErrorHnad:
    Err.Raise -1, "clsGlobal.ExecuteScript", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Sub Class_Terminate()
    Set mobjCall = Nothing
End Sub
