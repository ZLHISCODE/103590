Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Function HIS2Auto_DrugInfo(ByVal objConn As clsConnect, ByVal strContent As String) As String
'功能：上传药品基本信息
'参数：
'   objConnect：连接对象
'   strContent：已格式化后数据
'返回：消息内容
    
    
End Function

Public Function HIS2Auto_Dispense(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'功能：向自动化系统发送配药信息
'参数：
'   objDevice：设备对象
'   strContent：已格式化后数据
'返回：True成功；False失败

    
End Function

Public Function HIS2Auto_Dispensing(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'功能：向自动化系统发送发药信息
'参数：
'   objDevice：设备对象
'   strContent：已格式化后数据
'返回：True成功；False失败

    HIS2Auto_Dispensing = True

End Function
