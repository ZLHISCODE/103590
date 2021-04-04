Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Function DrugInfo(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'功能：上传药品基本信息
'参数：
'   objDevice：设备对象
'   strContent：已格式化后数据
'返回：True成功、False失败
    
    If objDevice Is Nothing Then Exit Function
    If objDevice.Status = False Then Exit Function

    On Error GoTo errHandle
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        'DB连接类型的数据上传。如调用存储过程、直接接入数据等
        DrugInfo = True
    
    Case enuLinkType.WEBServices
        'WebServices类型的数据上传，通过调用WebServices开放的接口函数上传数据。
        'objDevice.WSConnect.?????
        DrugInfo = True
        
    Case enuLinkType.Directory
        '文件类型的数据上传。
        DrugInfo = True
        
    End Select
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
    gobjComLib.ErrCenter
End Function

Public Function DrugStock(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'功能：上传药品库存信息
'参数：
'   objDevice：设备对象
'   strContent：已格式化后数据
'返回：True成功、False失败

    If objDevice Is Nothing Then Exit Function
    If objDevice.Status = False Then Exit Function
    
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        'DB连接类型的数据上传。如调用存储过程、直接接入数据等
        DrugStock = True
        
    Case enuLinkType.WEBServices
        'WebServices类型的数据上传，通过调用WebServices开放的接口函数上传数据。
        DrugStock = True
        
    Case enuLinkType.Directory
        '文件类型的数据上传。
        DrugStock = True
    
    End Select
    
errHandle:
    gstrMessage = Err.Description
    gobjComLib.ErrCenter
End Function

Public Function Dispense(ByVal objDevice As clsDevice, ByVal strNO As String, ByVal int单据 As Integer, ByVal strContent As String) As Boolean
'功能：向自动化系统发送配药信息
'参数：
'   objDevice：设备对象
'   strNO：单据号
'   strContent：已格式化后数据
'返回：True成功；False失败

    If objDevice Is Nothing Then Exit Function
    If objDevice.Status = False Then Exit Function

    Select Case objDevice.LinkType
    Case enuLinkType.DB
        'DB连接类型的数据上传。如调用存储过程、直接接入数据等
        Dispense = True
        
    Case enuLinkType.WEBServices
        'WebServices类型的数据上传，通过调用WebServices开放的接口函数上传数据。
        Dispense = True
        
    Case enuLinkType.Directory
        '文件类型的数据上传。
        Dispense = True
    
    End Select
    
    '调整发药窗口
    'If SetSendWin(药房ID, 单据号, 单据, 发药窗口) = False Then gstrMessage = "调整处方的发药窗口失败！"
    
errHandle:
    gstrMessage = Err.Description
End Function

Public Function Dispensing(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'功能：向自动化系统发送发药信息
'参数：
'   objDevice：设备对象
'   strContent：已格式化后数据
'返回：True成功；False失败

    If objDevice Is Nothing Then Exit Function
    If objDevice.Status = False Then Exit Function
    
    On Error GoTo errHandle
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        'DB连接类型的数据上传。如调用存储过程、直接接入数据等
        Dispensing = True
    
    Case enuLinkType.WEBServices
        'WebServices类型的数据上传，通过调用WebServices开放的接口函数上传数据。
        Dispensing = True
        
    Case enuLinkType.Directory
        '文件类型的数据上传。
        Dispensing = True
        
    End Select
    
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
End Function
