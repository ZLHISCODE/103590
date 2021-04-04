Attribute VB_Name = "mdlDefine"
Option Explicit

Public Type TYPE_PARAMS
    定时周期 As Integer
    有效天数 As Integer
    显示最大行数 As Integer
    输出日志 As Boolean
    详细日志 As Boolean
    保存日志天数 As Integer
    业务数据 As String
End Type

Public Const GSTR_MSG As String = "定时部件"

Public Function GetParameter(ByVal objXML As clsXML, ByVal strName As String, Optional ByVal strDefaultVal As String) As String
'功能：从zlDrugMachine.cfg文件中获取指定参数的值
'参数：
'  objXML：cfg文件的内容加载后的XML对象
'  strName：参数名称，即：XML结点名称
'返回：参数值

    Dim strValue As String

    If objXML Is Nothing Then
        GetParameter = strDefaultVal
        Exit Function
    End If
    
    strName = LCase(strName)
    
    If objXML.GetSingleNodeValue(strName, strValue) Then
        GetParameter = strValue
    Else
        GetParameter = strDefaultVal
    End If

End Function
