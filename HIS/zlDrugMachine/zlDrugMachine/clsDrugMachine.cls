VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'------------------------------------------------------------------------------
'说明：药品自动化设备接口类
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'注意：Init方法、Operation方法、Version属性，都是ZLHIS的调用接口，请勿随意变更
'------------------------------------------------------------------------------

Private mstrVersion As String
Private mblnInit As Boolean
Private mobjComLib As Object
Private mclsLog As New clsLog
Private mtypParams As TYPE_PARAMS
Private mintType As Integer
Private mstrUser As String
Private mblnDrugMachine As Boolean          '启用药品自动化
Private mblnIIP As Boolean                  '启用信息交互平台
Private mstrIIP_URL As String               '信息交互平台地址
Private mstrKey As String                   '信息交互平台密钥
Private mstrToken As String                 '信息交互平台令牌

Public Property Get Version() As String
Attribute Version.VB_UserMemId = 0
    mstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    Version = mstrVersion
End Property

Public Sub SaveLog(ByVal strContent As String)
'功能：将日志内容保存到zlDrugMachine.log文档中
'参数：
'  strContent：日志内容
    
    Dim strMessage As String

    If mblnInit = False Then
        strMessage = "未调用初始化方法或初始化未成功"
        mclsLog.Add strMessage
    End If
    
    mclsLog.Add strContent
    mclsLog.Save
End Sub

Public Function Operation(ByVal strUser As String, ByVal intType As Integer, ByVal strData As String, _
    ByRef strMessage As String) As Boolean
'功能：传送数据至接口
'参数：
'  strUser：登录ZLHIS的用户名
'  intType：业务类别
'           1-部门信息；
'           2-人员信息；
'           3-药品目录；
'           4-药品库存与库位；
'           5-发药窗口；
'           21-配药[门诊和住院处方明细上传]；
'           22-开始发药；
'           23-完成发药；
'           24-处方(完整/部分)退药；
'           25-处方完整退药；
'  strData：业务数据
'           业务类别=1：接口编号|部门性质1;部门性质2;…
'           业务类别=2：接口编号|人员性质1;人员性质2;…
'           业务类别=3：接口编号|剂型编码1;剂型编码2;…
'           业务类别=4：接口编号|库房id1;库房id2;…
'           业务类别=5：接口编号|库房id1;库房id2;…
'           业务类别=21 to 23：
'               门诊：1|单据1,处方号1;单据2,处方号2;…  说明：单据=8表收费单；单据=9表记帐单； 例：1|8,I0000001;8,I0000002
'               住院：2|收发ID1;收发ID2;…
'           业务类别=24：收费ID1,退药数量1;收费ID2,退药数量2;…   退药数量的单位为ZLHIS药品最小单位
'           业务类别=25：单据1,处方号1,库房id1;单据2,处方号2,库房id2;…
'  strMessage：返回消息
'
'返回：True成功；False失败

    Const INT_LEN As Integer = 5     'strData长度超过后，递归处理

    Dim objINF As clsINF
    Dim rsData As ADODB.Recordset
    Dim arrCode As Variant
    Dim strCode As String
    Dim l As Long
    Dim blnOK As Boolean
    Dim intIO As Integer    '1-门诊；2-住院；
    
'分解参数过大的问题
    Dim strIDs As String, strTmp As String
    Dim i As Long

'    If intType = 21 And Val(strData) = 2 Then
'        strIDs = Split(strData, "|")(1)     '收发id
'        If Len(strIDs) > INT_LEN Then
'            '分解收发id
'            For i = INT_LEN To 1 Step -1
'                If Mid(strIDs, i, 1) = ";" Then
'                    strTmp = "2|" & Mid(strIDs, i + 1)
'                    Exit For
'                End If
'            Next
'
'            '递归
'            Call Operation(strUser, intType, strTmp, strMessage)
'        Else
'            '递归最底层
'            With mdlDefine.gtypYUYAMA
'                .BottomLevel = True
'                .MacNO = 0
'                .SendIDs = ""
'            End With
'        End If
'
'        '排除已发送的ID
'        If mdlDefine.gtypYUYAMA.SendIDs <> "" And InStrRev(strData, mdlDefine.gtypYUYAMA.SendIDs) > 0 Then
'            strData = Left(strData, InStrRev(strData, mdlDefine.gtypYUYAMA.SendIDs) - 2)
'        End If
'        mdlDefine.gtypYUYAMA.SendIDs = Split(strData, "|")(1)
'        Debug.Print strData
'    End If
    mclsLog.AddTime
    mstrUser = strUser

    If mblnInit = False Then
        strMessage = "未调用初始化方法或初始化未成功"
        mclsLog.Add strMessage
        mclsLog.Save
        Exit Function
    End If
    
    If mblnDrugMachine = False Then
        mclsLog.Add "未启用药品自动化设备接口"
        mclsLog.Save
        Exit Function
    End If
    
    mclsLog.Add "开始传送数据至接口"
    mclsLog.Add "参数1:" & strUser, 1, 1
    mclsLog.Add "参数2:" & intType, 1, 1
    mclsLog.Add "参数3:" & strData, 1, 1
    
    '初始化
    Set objINF = New clsINF
    With objINF
        Set .mobjComLib = mobjComLib
        Set .mobjLog = mclsLog
        
'        .mobjLog.Path = mclsLog.Path
'        .mobjLog.FileName = mclsLog.FileName
'        .mobjLog.Action = mclsLog.Action
'        .mobjLog.SaveDays = mclsLog.SaveDays
'        .mobjLog.Type_ = mclsLog.Type_
    End With
    
    '获取ZLHIS的标准数据
    Set rsData = objINF.GetData(intType, strData, blnOK)      '记录集按接口排序
    If rsData Is Nothing Then
        If blnOK = False Then
            strMessage = "获取ZLHIS标准数据失败"
            Operation = False
        Else
            strMessage = ""
            Operation = True
        End If
        
        Exit Function
    Else
        If rsData.RecordCount <= 0 Then
            mclsLog.Add "获取ZLHIS标准数据暂时无"
            strMessage = ""
            Operation = True
            Exit Function
        End If
    End If
    
    If intType > 20 Then
        '业务数据
        
        If intType >= 21 And intType <= 23 Then     '21-配药；22-开始发药；23-完成发药
            '获取（1-门诊、2-住院）信息
            intIO = Val(strData)
        End If
    
        '获取数据的接口编号
        arrCode = Split(GetINFCodes(rsData), ";")
        
        '遍历接口
        For l = LBound(arrCode) To UBound(arrCode)
            If Trim(arrCode(l)) <> "" Then
                Call OperationEx(intType, intIO, strData, rsData, objINF, arrCode(l), strMessage)
            Else
                mclsLog.Add "业务数据的接口编号无"
            End If
        Next
        
        If UBound(arrCode) < 0 Then mclsLog.Add "业务数据的接口编号无"
        
    Else
        '基础数据
        
        '获取接口编号
        strCode = Split(strData, "|")(0)
        
        If Trim(strCode) <> "" Then
            If OperationEx(intType, intIO, "", rsData, objINF, strCode, strMessage) = False Then
                Operation = False
                If Not rsData Is Nothing Then
                    If rsData.State = adStateOpen Then rsData.Close
                    Set rsData = Nothing
                End If
                Exit Function
            End If
        Else
            mclsLog.Add "基础数据的接口编号无"
        End If
        
    End If
    
    '关闭记录集
    If Not rsData Is Nothing Then
        If rsData.State = adStateOpen Then rsData.Close
        Set rsData = Nothing
    End If
    
    '完成传送操作
    mdlDefine.gtypYUYAMA.BottomLevel = False
    mclsLog.Add "完成传送数据至接口"
    mclsLog.Save
    strMessage = ""
    Operation = True
    
    Exit Function
    
hErr:
    strMessage = Err.Number & "：" & Err.Description
    mclsLog.Add strMessage
    mclsLog.Add "传送数据至接口失败"
    mclsLog.Save
End Function

Private Function OperationEx(ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strData As String, ByVal rsData As ADODB.Recordset, _
    ByVal objBase As clsINF, ByVal strCode As String, ByRef strMessage As String) As Boolean
'参数：
'  intType：业务类型
'  intIO：1-门诊；2-住院
'  strData：业务数据
'  rsData：业务数据的记录集对象
'  objBase：接口基类
'  strCode：接口编号

    Dim objThird As clsINF
    Dim rsFinish As ADODB.Recordset
    Dim blnOK As Boolean

    '区分接口
    
    If mblnIIP Then
        '中联信息交互平台
        Set objThird = New clsINF_IIP
        With objThird
            Set .mobjComLib = objBase.mobjComLib
            Set .mobjLog = objBase.mobjLog
            .URL = mstrIIP_URL
            .KEY = mstrKey
            .Token = mstrToken
            .UserName = mstrUser
        End With
    Else
        '第三方接口
        Select Case GetINFType(strCode)
        Case 1          '韦乐海茨
            '请注意“韦乐海茨”目前有两个不同版本的类模块
            Set objThird = New clsINF_WillachHeise
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'WebService地址
                .UserName = mstrUser
                
                '如果要获取韦乐海茨的特殊数据
                'If rsData.State = adStateOpen Then rsData.Close
                'Set rsData = .GetData(intType, strData, blnOK)
            End With
        Case 2          'TOSHO
            Set objThird = New clsINF_TOSHO
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'OLEDB连接串
                .UserName = mstrUser
                
                '如果要获取TOSHO的特殊数据
                'If rsData.State = adStateOpen Then rsData.Close
                'Set rsData = .GetData(intType, strData, blnOK)
            End With
        Case 3          '蝶和
            Set objThird = New clsINF_DIH
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'WebService地址
                .UserName = mstrUser
            
                '重新获取蝶和的特殊数据
                If intType = Val("22-开始发药") Or intType = Val("23-完成发药") Then
                    If rsData.State = adStateOpen Then rsData.Close
                    Set rsData = .GetData(intType, strData, blnOK)
                    If rsData Is Nothing Then
                        If blnOK = False Then
                            strMessage = "获取ZLHIS标准数据（蝶和）失败"
                            OperationEx = False
                        Else
                            strMessage = ""
                            OperationEx = True
                        End If
                        Exit Function
                    Else
                        If rsData.RecordCount <= 0 Then
                            mclsLog.Add "获取ZLHIS标准数据（蝶和）暂时无"
                            strMessage = ""
                            OperationEx = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        Case 4      '关拉尼
            Set objThird = New clsINF_Grani
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'WebService地址
                .UserName = mstrUser
                
                '如果要获取关拉尼的特殊数据
                'If rsData.State = adStateOpen Then rsData.Close
                'Set rsData = .GetData(intType, strData, blnOK)
            End With
        Case 5      'YUYAMA
            Set objThird = New clsINF_YUYAMA
            With objThird
                Set .mobjComLib = objBase.mobjComLib
                Set .mobjLog = objBase.mobjLog
                .URL = mdlDrugMachine.GetInterfaceLink(objBase, strCode)    'OLEDB连接串
                .UserName = mstrUser
                
                '如果要获取YUYAMA的特殊数据
                'If rsData.State = adStateOpen Then rsData.Close
                'Set rsData = .GetData(intType, strData, blnOK)
            End With
        Case Else
            strMessage = "未知的接口类型"
            mclsLog.Add strMessage, 1
            Exit Function
        End Select
    End If

    '保存单据信息
    Select Case intType
    Case 21, 23
        objThird.mstrNOs = Split(strData, "|")(1)
    Case 24, 25
        objThird.mstrNOs = strData
    Case Else
        objThird.mstrNOs = ""
    End Select
    
    '接口开启
    objThird.Action = True
    If objThird.Action = False Then Exit Function
    
    '过滤对应接口编号的数据
    If objThird.FilterData(strCode, intType, rsData, rsFinish) = False Then Exit Function
    
    '传送数据   mintType是应用类别；intType是业务类型
    OperationEx = objThird.SendData(mintType, intType, intIO, strCode, rsFinish, strMessage)
    
    '结束
    objThird.Action = False
    
End Function

Public Function Init(ByVal intType As Byte, ByVal objControl As Object, ByRef strMessage As String) As Boolean
'功能：初始化
'参数：
'  intType：应用类别。1-ZLHIS；2-接口管理工具（定时向第三方接口传送数据）；3-移动（支付宝）；4-中联自助平台
'  objControl：clsComLib对象；如果应用类别为“3-支付宝”，请传入ADO的连接对象
'  strMessage：返回消息
'返回：True成功；False失败

    Dim strFile As String, strTmp As String

    mblnInit = False
    mintType = intType
    strMessage = ""
    
    '检查参数文件（调试专用）
    If LCase(App.Path) Like "*\apply" Then
        strFile = App.Path & "\" & GSTR_CONFIG_FILE
    ElseIf LCase(App.Path) Like "*\apply\*" Then
        strFile = Left(App.Path, InStr(LCase(App.Path), "\apply\") + Len("\apply\") - 1) & GSTR_CONFIG_FILE
    ElseIf LCase(App.Path) Like "*zldrugmachinemanage*" Or LCase(App.Path) Like "*zldrugmachine\*" Or LCase(App.Path) Like "*zldrugmachine" Then
        strFile = Replace(App.Path, "\" & App.EXEName, "") & "\" & App.EXEName & "\zlDrugMachineManage\" & GSTR_CONFIG_FILE
    Else
        strMessage = "“" & App.Path & "\" & GSTR_CONFIG_FILE & "”文件路径不存在"
        Exit Function
    End If
    
    If mdlDrugMachine.VerifyConfigFile(strFile) = False Then
        strMessage = "创建“" & GSTR_CONFIG_FILE & "”配置文件失败"
    End If
    If strMessage <> "" Then Exit Function
    
    '获取参数
    Call mdlDrugMachine.ReadParams(mtypParams)
    
    '设置日志对象的属性
    With mclsLog
        .Action = mtypParams.输出日志
        .Type_ = IIf(mtypParams.详细日志, 1, 0)
        .SaveDays = mtypParams.保存日志天数
    End With
    
    '日志
    mclsLog.AddTime
    If strMessage <> "" Then mclsLog.Add strMessage
    mclsLog.Add "更新参数完成"
    
    '检查入参
    If intType < 1 Or intType > 4 Then
        strMessage = "应用类别未传入"
        mclsLog.Add strMessage
        mclsLog.Save
        Exit Function
    End If
    If objControl Is Nothing Then
        strMessage = "对象未传入"
        mclsLog.Add strMessage
        mclsLog.Save
        Exit Function
    End If
    
    '初始化部件
    On Error GoTo hErr
    If intType = Val("3-支付宝") Then
        Set mobjComLib = New clsPublic
        If TypeName(objControl) <> "Connection" Then
            strMessage = "支付宝未传入数据连接对象"
            mclsLog.Add strMessage
            mclsLog.Save
            Exit Function
        End If
        strMessage = "调用“zl9CommLib.InitCommon”方法失败，请检查zl9ComLib部件版本！"
        mobjComLib.InitCommon objControl
    Else
        If TypeName(objControl) <> "clsComLib" Then
            strMessage = "未传入clsComLib对象"
            mclsLog.Add strMessage
            mclsLog.Save
            Exit Function
        End If
    
        Set mobjComLib = objControl
        strMessage = "调用“zl9CommLib.CurrentConn”方法失败，请检查zl9ComLib部件版本！"
        If mobjComLib.CurrentConn Is Nothing Then
            strMessage = "zlDrugMachine部件的数据连接创建"
            mclsLog.Add strMessage
            mclsLog.Save
            Exit Function
        End If
    End If
    strMessage = ""
    
    '获取系统的模块参数
    mclsLog.Add "获取模块参数"
    If intType = Val("3-支付宝") Then
        strTmp = mobjComLib.GetPara("启用信息交互平台", GLNG_SYS, GLNG_MODULE)
        mblnDrugMachine = Val(mobjComLib.GetPara("启用药品自动化设备接口", GLNG_SYS, GLNG_MODULE)) = 1
        mstrKey = mobjComLib.GetPara("信息交互平台密钥", GLNG_SYS, GLNG_MODULE)
        mstrToken = mobjComLib.GetPara("信息交互平台令牌", GLNG_SYS, GLNG_MODULE)
    Else
        strTmp = mobjComLib.zlDatabase.GetPara("启用信息交互平台", GLNG_SYS, GLNG_MODULE)
        mblnDrugMachine = Val(mobjComLib.zlDatabase.GetPara("启用药品自动化设备接口", GLNG_SYS, GLNG_MODULE)) = 1
        mstrKey = mobjComLib.zlDatabase.GetPara("信息交互平台密钥", GLNG_SYS, GLNG_MODULE)
        mstrToken = mobjComLib.zlDatabase.GetPara("信息交互平台令牌", GLNG_SYS, GLNG_MODULE)
    End If
    mblnIIP = Val(strTmp) = 1
    mstrIIP_URL = Mid(strTmp, InStr(strTmp, "|") + 1)
    
    If mblnDrugMachine = False Then
        strMessage = "未启用药品自动化设备接口"
        mclsLog.Add strMessage
        mclsLog.Save
        Exit Function
    End If
    
    '初始化完成
    mblnInit = True
    Init = True
    
    mclsLog.Add "初始化完成"
    mclsLog.Save
    
    Exit Function
    
hErr:
    mclsLog.Add IIf(strMessage = "", Err.Number & "：" & Err.Description, strMessage)
    mclsLog.Save
End Function

Private Sub Class_Initialize()
    With mclsLog
        .Path = App.Path
        .FileName = App.EXEName
    End With
End Sub

Private Sub Class_Terminate()
    Dim frmThis As Form
    
    mclsLog.Save
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next

    Set mclsLog = Nothing
    Set mobjComLib = Nothing
    Set gcnThird = Nothing
End Sub

Private Function GetINFCodes(ByVal rsData As ADODB.Recordset) As String
'功能：从标准记录集对象得到所有接口编号
'参数：
'  rsData：标准记录集对象
'返回：接口编号1[;接口编号2[...]]

    Dim l As Long
    Dim strTmp As String, strResult As String
    
    With rsData
        If .State <> adStateOpen Then Exit Function
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            If mdlDrugMachine.NVL(!接口编号) <> "" Then
                strTmp = Trim(!接口编号)
                If Not ";" & strResult & ";" Like "*;" & strTmp & ";*" Then
                    strResult = strResult & ";" & strTmp
                End If
            End If
            
            .MoveNext
        Loop
        If Left(strResult, 1) = ";" Then strResult = Mid(strResult, 2)
    End With
    
    GetINFCodes = strResult

End Function

Private Function GetINFType(ByVal strCode As String) As Integer
'功能：获取指定接口编号的接口类型
'参数：
'  strCode：接口编号
'返回：接口类型；1-韦乐海茨；2-TOSHO；3-蝶和；4-关拉尼；5-YUYAMA

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    
    strSQL = "Select 类型 From 药品设备接口 Where 编号 = [1] "
    
    If mintType = Val("3-支付宝") Then
        Set rsTmp = mobjComLib.OpenSQLRecord(strSQL, "获取接口的类型", strCode)
    Else
        Set rsTmp = mobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取接口的类型", strCode)
    End If
    If rsTmp.EOF = False Then
        GetINFType = IIf(IsNull(rsTmp!类型), 0, rsTmp!类型)
    End If
    rsTmp.Close
    
    Exit Function

hErr:
    mclsLog.Add Err.Number & "：" & Err.Description
    mclsLog.Save
End Function
