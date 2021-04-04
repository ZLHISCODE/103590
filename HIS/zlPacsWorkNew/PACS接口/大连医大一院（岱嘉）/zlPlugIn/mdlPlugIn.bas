Attribute VB_Name = "mdlPlugIn"
Option Explicit

Public gstrSql      As String
Public gcnOracle    As ADODB.Connection
Public gComLib      As Object
Public gDatabase    As Object

Public Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal wFlags As Long) As Long

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

'========================================================================================
'=名稱:检测(ChkRsState)
'=入口参数:Rs               类型:ADODB.Recordset
'=出口参数:ChkRsState       类型:Boolean
'=功能:检测记录集的状态
'=日期:2004-07-08
'=編程:谢荣
'========================================================================================
Function ChkRsState(rs As ADODB.Recordset) As Boolean
On Error GoTo ErrH:
    With rs
        If rs Is Nothing Then
            ChkRsState = True
            Exit Function
        Else
            ChkRsState = False
        End If
        If rs.State = 0 Then
            ChkRsState = True
            Exit Function
        Else
            ChkRsState = False
        End If
        If .RecordCount < 1 Then
            ChkRsState = True
        Else
            ChkRsState = False
        End If
        If .EOF Or .BOF Then
            ChkRsState = True
        Else
            ChkRsState = False
        End If
    End With
    Exit Function
ErrH:
    err.Clear
End Function

'检测长度是否超过长度(字节数)
Function ChkStrUniCode(ByVal mStr As String, Optional ByVal mLen As Long = 0) As String
    Dim strL        As String
On Error GoTo ErrH
    mStr = ConvertString(mStr)
    If mLen <= 0 Then
        ChkStrUniCode = mStr
        Exit Function
    Else
        strL = StrConv(mStr, vbFromUnicode)
        strL = LeftB(strL, mLen)
        ChkStrUniCode = StrConv(strL, vbUnicode)
    End If
    Exit Function
ErrH:
    err.Clear
    ChkStrUniCode = ""
    Exit Function
End Function

'==================================================================================================
'=名称:去掉字符串中的单引号("'")(ConvertString)
'=入口参数:
'=1).sStr          类型:String
'=出口参数:空
'=功能:去掉字符串(sStr)中的单引号
'=日期:2004-12-11
'=编程:欧阳
'=说明:在SQL语句中不能带单引号
'======================================sssssss============================================================
Function ConvertString(ByVal sStr As String) As String
On Error GoTo ErrH
    ConvertString = Replace(sStr, "'", "")
    ConvertString = Replace(ConvertString, "—", "")
    ConvertString = Replace(ConvertString, "&", "")
    Exit Function
ErrH:
    err.Clear
    ConvertString = ""
End Function
