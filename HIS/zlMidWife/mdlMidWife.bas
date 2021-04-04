Attribute VB_Name = "mdlMidWife"
Option Explicit
Public gcnOracle As ADODB.Connection

Public glngInstance As Long  '类的实例计数
Public gstrUser As String    '当前登录导航台的用户
Public gstrURL As String     '单病人页的URL
Public gstrURLLogin As String '登录助产信息主界面的初始URL
Public glngPatiID As Long, glngPageID As Long '记录上一次病人的病人id和主页id

Public Function GetEncrypt(ByVal strCode As String) As String
'功能：处理Url及加密方法：
    Dim tmp() As Byte, strResult As String
    Dim i As Integer
    
    tmp = StrConv(strCode, vbFromUnicode)
    strResult = URLEncode(tmp(0) + UBound(tmp) + 1)
    
    For i = 1 To UBound(tmp)
        strResult = strResult & URLEncode(tmp(i) + tmp(i - 1))
    Next
    GetEncrypt = strResult
End Function

Public Function URLEncode(ByVal intValue As Integer) As String
    Dim s As String
    
    If (intValue >= 48 And intValue <= 57) Or (intValue >= 65 And intValue <= 90) Or (intValue >= 97 And intValue <= 122) Then
      s = Chr(intValue)
    ElseIf intValue = 32 Then
      s = "+"
    Else
      s = "%" & Hex(intValue)
    End If

    URLEncode = s
End Function

Public Function GetSysPar(lngPar As Long, lngSys As Long) As String
    Dim strSql As String, rstmp As ADODB.Recordset
    
    On Error GoTo errHandle
    '执行：Set cmdData.ActiveConnection = gcnOracle
    '会报“参数类型不正确”，可能是ActiveExe方式对传入的连接对象有什么限制，所以不能使用公共部件函数OpenSqlRecord
    
    strSql = "Select Nvl(参数值,缺省值) as 参数值 From zlParameters Where 参数号= " & lngPar & " And 系统 = " & lngSys & " And Nvl(模块,0)=0"
    Set rstmp = gcnOracle.Execute(strSql)
    If rstmp.RecordCount > 0 Then GetSysPar = "" & rstmp!参数值

    Exit Function
errHandle:
    MsgBox Err.Description & strSql, vbExclamation, "读取参数"
    GetSysPar = ""
End Function




