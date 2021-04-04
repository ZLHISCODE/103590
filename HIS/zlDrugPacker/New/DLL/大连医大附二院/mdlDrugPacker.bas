Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Function DrugInfo(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'功能：上传药品基本信息
'参数：
'   objDevice：设备对象
'   strContent：已格式化后数据
'返回：True成功、False失败
    
    Dim cmOutside As New ADODB.Command
    Dim strResume As String, strIP As String
    
    strIP = GetLocalIP
    
    On Error GoTo errHandle
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        On Error GoTo errDB
        objDevice.DBConnect.BeginTrans
        Set cmOutside.ActiveConnection = objDevice.DBConnect
        cmOutside.CommandText = strContent
        cmOutside.Execute
        objDevice.DBConnect.CommitTrans
        DrugInfo = True
    
    Case enuLinkType.WEBServices
        'TransConsiData
        '参数1：操作ID
        '参数2：业务类型
        '参数3：XML文本
        '参数4：本机IP
        '参数5：HIS用户编号
        '参数6：HIS用户姓名
        '参数7：返回信息
        If objDevice.WSConnect.TransConsisData(1, 101, strContent, strIP, gstrUserCode, gstrUserName, strResume) <> 1 Then
            gstrMessage = strResume
        Else
            DrugInfo = True
        End If
        
    Case enuLinkType.Directory
        
        
    End Select
    Exit Function
    
errDB:
    objDevice.DBConnect.RollbackTrans
    Exit Function
    
errHandle:
    gobjComLib.ErrCenter
End Function

Public Function DrugStock(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'功能：上传药品库存信息
'参数：
'   objDevice：设备对象
'   strContent：已格式化后数据
'返回：True成功、False失败
    
    Dim strIP As String, strResume As String
    Dim intRetval As Integer
    
    strIP = GetLocalIP
    
    Select Case objDevice.LinkType
    Case enuLinkType.DB
        
    Case enuLinkType.WEBServices
        If objDevice.WSConnect.TransConsisData(1, 102, strContent, strIP, gstrUserCode, gstrUserName, intRetval, strResume) <> 1 Then
            gstrMessage = strResume
        Else
            DrugStock = True
        End If
        
    Case enuLinkType.Directory
    
    End Select
    
End Function

Public Function Dispense(ByVal objDevice As clsDevice, ByVal strNO As String, ByVal int单据 As Integer, ByVal strContent As String) As Boolean
'功能：向自动化系统发送配药信息
'参数：
'   objDevice：设备对象
'   strNO：单据号
'   strContent：已格式化后数据
'返回：True成功；False失败

    Dim strIP As String, strResume As String
    Dim intRetval As Integer
    
    strIP = GetLocalIP
    
    If objDevice.Status Then
        
        On Error GoTo errHandle
    
        Select Case objDevice.LinkType
        Case enuLinkType.DB
            '
        Case enuLinkType.WEBServices
            If objDevice.WSConnect.TransConsisData(1, 201, strContent, strIP, gstrUserCode, gstrUserName, intRetval, strResume) <> 1 Then
                gstrMessage = strResume
            Else
                Dispense = True
            End If
            
            '发药窗口
            If SetSendWin(objDevice.DeptID, strNO, int单据, intRetval) = False Then gstrMessage = "调整处方的发药窗口失败！"
            
            
        Case enuLinkType.Directory
        End Select
        
    End If
    
    Exit Function

errHandle:
    gstrMessage = Err.Description
End Function

Public Function Dispensing(ByVal objDevice As clsDevice, ByVal strContent As String) As Boolean
'功能：向自动化系统发送发药信息
'参数：
'   objDevice：设备对象
'   strContent：已格式化后数据
'返回：True成功；False失败

    Dim strIP As String, strResume As String
    Dim cmOutside As New ADODB.Command
    Dim intRetval As Integer
    
    strIP = GetLocalIP
    
    On Error GoTo errHandle
    If objDevice.Status Then
        Select Case objDevice.LinkType
        Case enuLinkType.DB
            Set cmOutside.ActiveConnection = objDevice.DBConnect
            cmOutside.CommandText = strContent
            cmOutside.Execute
            Dispensing = True
        
        Case enuLinkType.WEBServices
            If objDevice.WSConnect.TransConsisData(1, 202, strContent, strIP, gstrUserCode, gstrUserName, intRetval, strResume) <> 1 Then
                gstrMessage = strResume
            Else
                Dispensing = True
            End If
            
        Case enuLinkType.Directory
        End Select
    End If
    Exit Function
    
errHandle:
    gstrMessage = Err.Description
End Function
