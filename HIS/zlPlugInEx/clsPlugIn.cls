VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------------
'功能：扩展接口插件模板
'如果要做限制使用 请实现 GetUserName 方法
'---------------------------------------------------------------------------
Public gblnMec As Boolean '是否开启扩展部件检查病案首页
Public gblnLoadMec As Boolean '是否开启扩展部件加载病案首页数据

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    Set gcnOracle = cnOracle
    
    gblnMec = True '开启自定义检查病案首页数据
    gblnLoadMec = True '开启自定义加载病案首页数据
End Sub

Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口终止
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
     
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1, Optional strReserve As String) As String
'功能：获取外挂功能列表；该接口一般在初始化之后调用一次，调用时机和频度可能因不同的主程序而不同
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号，病历调用时 lngModual=1070，无int场合参数
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=保留参数,用于扩展使用
'      医技站中的调用时,传入为当前执行的前提医嘱ID
'返回：外挂功能列表，格式为"功能名1,功能名2,功能名3,..."
'      如果某个功能需要在病人切换(或刷新)时自动调用，则在功能名前面加上"Auto:"前缀，这种功能只允许有一个。

'特殊说明：当插件在人员管理模块调用时只传入lngSys和lngModual，并且其中的lngModual＝0
'          自动调用功能在人员管理模块调用时不会自动调用
    Dim strFuncNames As String
    
    Dim str菜单栏 As String '菜单栏中放置的功能，菜单栏中应该包含所有功能名， str菜单栏 = str工具栏 ∪ str弹出栏
    Dim str工具栏 As String
    Dim str弹出栏 As String
    
    Select Case lngModual
    Case m医嘱附费模块
        '■医技工作站，附费页签
        strFuncNames = "功能Ex"
        
        str菜单栏 = "功能Ex"
        str工具栏 = "|:InTool:功能Ex"
        str弹出栏 = "InTool:功能Ex"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
    Case m门诊医嘱模块
        If int场合 = 0 Then
            '■门诊医生工作站，医嘱页签
            strFuncNames = "功能1,功能2,Auto:功能3"
            
            str菜单栏 = "功能Ex"
            str工具栏 = "|:InTool:功能Ex"
            str弹出栏 = "InTool:功能Ex"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        Else
            '■医技工作决，医嘱页签
            strFuncNames = "功能1,功能2,Auto:功能3"
            
            str菜单栏 = "功能Ex"
            str工具栏 = "|:InTool:功能Ex"
            str弹出栏 = "InTool:功能Ex"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        End If
    Case m住院医嘱模块
        If int场合 = 0 Then
            '■住院医生工作站，医嘱页签
            strFuncNames = "功能Ex"
            
            str菜单栏 = "功能Ex"
            str工具栏 = "|:InTool:功能Ex"
            str弹出栏 = "InTool:功能Ex"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        ElseIf int场合 = 1 Then
            '■老版护士工作站，医嘱页签/新版护士站病人事务窗口，医嘱页签
            strFuncNames = "功能Ex"
            
            str菜单栏 = "功能Ex"
            str工具栏 = "|:InTool:功能Ex"
            str弹出栏 = "InTool:功能Ex"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        Else
            '■医技工作站，医嘱页签
            strFuncNames = "功能Ex"
            
            str菜单栏 = "功能Ex"
            str工具栏 = "|:InTool:功能Ex"
            str弹出栏 = "InTool:功能Ex"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        End If
    Case m临床路径模块
        If int场合 = 0 Then
            '■住院医生工作站，临床路径页签
            strFuncNames = "功能Ex"
            
            str菜单栏 = "功能Ex"
            str工具栏 = "|:InTool:功能Ex"
            str弹出栏 = "InTool:功能Ex"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        Else
            '■护士工作站，临床路径页签/新版护士站病人事务窗口，临床路径页签
            strFuncNames = "功能Ex"
            str菜单栏 = "功能Ex"
            str工具栏 = "|:InTool:功能Ex"
            str弹出栏 = "InTool:功能Ex"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        End If
    Case m新版护士站
        '■新版护士工作站主界面
        strFuncNames = "功能Ex"
        
        str菜单栏 = "功能Ex"
        str工具栏 = "|:InTool:功能Ex"
        str弹出栏 = "InTool:功能Ex"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
    Case m病历模块
        '■病历书写页面
        strFuncNames = "功能Ex"
    Case m人员管理模块
        '■人员管理页面
        strFuncNames = "功能Ex"
    Case m体检中心管理
        '■体检中心管理
        strFuncNames = "功能Ex1,功能Ex2,InTool:功能Ex3"
    Case m体检总检登记
        '■体检总检登记
        strFuncNames = "功能Ex1,功能Ex2,InTool:功能Ex3"
    Case m体检分科执行
        '■体检分科执行
        strFuncNames = "功能Ex1,功能Ex2,InTool:功能Ex3"
    Case m体检结果登记
        '■体检结果登记
        strFuncNames = "功能Ex1,功能Ex2,InTool:功能Ex3"
    End Select
    
    GetFuncNames = strFuncNames
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int场合 As Integer = -1)
'功能：执行某个外挂功能
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号，病历调用时 lngModual＝1070，无int场合参数
'      strFuncName=要执行的功能名，该功能名之前由GetFuncNames函数返回
'      lngPatiID=当前病人ID,人员管理调用时，传入的人员ID。
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如医嘱ID
'           病历调用时 varKeyId=电子病历记录.id，在新增编辑未保存前其值=0,保存后或修改、审订时执行有值；
'           临床路径 varKeyId 病人路径id
'      strReserve=保留参数,用于扩展使用 当是病历调用时 strReserve＝ 病历文件列表.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

'特殊说明：当插件在人员管理模块调用时只传入lngSys、lngModual、strFuncName、lngPatiID、varRecId、varKeyId，
'          lngPatiID就是人员id，如果lngPatiID＝0表示界面上未选中人员；其中的varRecId和varKeyId等值都是传的0

    '注意：在采用分离插件后，功能的执行时，应该先对功能明进行判断。
    '先检查功能是否在本插件（zlPlugIn）中，然后再去执行扩展插件
    
    Select Case lngModual
        Case m门诊医嘱模块
            If int场合 = 0 Then
                MsgBox strFuncName & "执行成功！" '门诊医生站医嘱页面
            ElseIf int场合 = 2 Then
                MsgBox strFuncName & "执行成功！" '医技站门诊医嘱页面
            End If
        Case m住院医嘱模块
            If int场合 = 0 Then
                MsgBox strFuncName & "执行成功！" '住院医生站医嘱页面
            ElseIf int场合 = 2 Then
                MsgBox strFuncName & "执行成功！" '医技站住院医嘱页面
            End If
        Case m住院护士站模块
            MsgBox strFuncName & "执行成功！" '住院护士站医嘱页面
        Case m临床路径模块
            If int场合 = 0 Then
                MsgBox strFuncName & "执行成功！" '住院医生站临床路径页面
            ElseIf int场合 = 1 Then
                MsgBox strFuncName & "执行成功！" '住院护士站临床路径页面
            End If
        Case m病历模块
            MsgBox strFuncName & "执行成功！" '病历书写页面
        Case m人员管理模块
            MsgBox strFuncName & "执行成功！" '人员管理页面
        Case m医嘱附费模块
            MsgBox strFuncName & "执行成功！" '医技站医嘱附费页面
        Case m体检中心管理
            MsgBox strFuncName & "执行成功！" '体检中心管理
        Case m体检总检登记
            MsgBox strFuncName & "执行成功！" '体检总检登记
        Case m体检分科执行
            MsgBox strFuncName & "执行成功！" '体检分科执行
        Case m体检结果登记
            MsgBox strFuncName & "执行成功！" '体检结果登记
    End Select
End Sub
 
Public Function GetFormCaption(ByVal lngSys As Long, ByVal lngModual As Long) As String
'功能：在住院医生工作站中添加卡片的名称，逗号分割。
        Exit Function '如使用此功能请修改接口内部实现
    Select Case lngModual
    Case m住院医生工作站
        GetFormCaption = "Test3"
    Case m住院护士工作站
        GetFormCaption = "Test3"
    Case m门诊医生工作站
        GetFormCaption = "Test3"
    Case m医技工作站
        GetFormCaption = "Test3"
    End Select
End Function

Public Function GetForm(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'功能：根据索引值返回对象。工作站窗体加载时调用
        Exit Function '如使用此功能请修改接口内部实现
    Select Case lngModual
    Case m住院医生工作站
        Select Case strName
        Case "Test3"
            Set GetForm = New frmDockEx
        End Select
    Case m住院护士工作站
        Select Case strName
        Case "Test3"
            Set GetForm = New frmDockEx
        End Select
    Case m门诊医生工作站
        Select Case strName
        Case "Test3"
            Set GetForm = New frmDockEx
        End Select
    Case m医技工作站
        Select Case strName
        Case "Test3"
            Set GetForm = New frmDockEx
        End Select
    End Select
 
End Function

Public Sub RefreshForm(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String, ByVal lngPatiID As Long, ByVal str挂号单 As String, ByVal lng主页ID As Long, _
    Optional ByVal blnMoved As Boolean, Optional ByVal lng界面科室ID As Long, Optional ByVal lng医护科室ID As Long, Optional ByVal lng病区ID As Long, Optional ByVal lng科室id As Long, _
    Optional ByVal lng前提ID As Long, Optional ByVal int状态 As Integer, Optional ByVal int执行状态 As Integer, Optional ByVal lng路径状态 As Long = -1)
'功能：刷新内部界面
'参数：strName 当前窗体名称
'      lngPatiID 病人id,str挂号单 挂号单号门诊病人才有否则为空，lng主页ID 住院病人才有否则为0
'      blnMoved=该病人的数据是否已转出
'      lng界面科室ID=如果当前医生站是会诊病人，则为会诊科室ID；如果是医技站调用，则为医技科室ID
'      lng医护科室ID=医护站界面ID
'      lng病区ID，lng科室ID=当“5-最近转科病人”时为病人原病区或原科室
'      lng前提ID=当由医技站调用时传入
'      int状态=0-在院病人,1-预出院病人,2-出院病人,3-待会诊病人(在院),4-已会诊病人,5-最近转科病人(在院),6-转科或转病区待入住病人(不含入院待入住)
'      int执行状态=当由医技站调用时传入,医技执行项目的执行状态
'      lng路径状态=-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
    Exit Sub '如使用此功能请修改接口内部实现
    If objForm Is Nothing Then Exit Sub
    Call objForm.RefreshInSide
       
End Sub

Public Function GetButtomName(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String) As String
'功能：添加到菜单和工具栏上的按钮名称，菜单上级为 扩展，工具栏上直接加按钮
'参数：strName 当前窗体名称
        Exit Function '如使用此功能请修改接口内部实现
    If objForm Is Nothing Then Exit Function
    GetButtomName = objForm.GetInSideFunc
 
End Function

Public Function ExeButtomClick(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String, ByVal strFuncName As String, _
    ByVal lngPatiID As Long, ByVal lng主页ID As Long, ByVal str挂号单 As String, Optional ByVal lng医嘱ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：执行窗体内部的功能
'参数：strName 当前窗体名称，strFuncName 菜名称(窗体中功能名称)
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lng主页ID,str挂号单
'      lng医嘱ID=只有医技站传调用时会传入， 否则没传值
    Exit Function '如使用此功能请修改接口内部实现
    If objForm Is Nothing Then Exit Function
    ExeButtomClick = objForm.ExecuteFunc(strFuncName)
End Function

Public Function GetUserName() As String
'功能：获取单位名称
'说明：用户名空时表示不限制
'例  GetUserName = "中联技术中心医院信息系统"

    GetUserName = ""
End Function
 
 
'*****************************************************************
'                   特殊说明请注意
'*****************************************************************
'-------从此行开始以下的接口在使用扩展部件时有如下限制------------
'1.以下接口仅支持互斥扩展，即接口只能扩展一次；
'      例如 QueryPatiInfo ，扩展为一个独立部件zlPlugInQP，在zlPlugInQP中实现了QueryPatiInfo方法，此时就不能再增加一个扩展部件对其进行扩展了。
'      原因：a)接口本身会对主体程序带来影响，返回值，引用类型参数等；
'            b)如果接口多次扩展后，返回值的数据不能叠加；
'            c)主插件在调用扩展插件功能时，是遍历试调用，调用的先后顺序会影响最终结果；
'               如：接Fun1，被两个扩展部件(部件A和B)实现，如果A和B部件同时存在，则A.Fun1和B.Fun1都会被调用，后执行的接口的返回值会覆盖前一个，先A再B则A失效，反之B失效，
'                   主插件在调用扩展部件接口的顺序目前是不能预先判定的，所以就要求，只能被扩展一次。
'2.扩展部件实现接口数量的关系，用一个集合关系来说明：
'  ALL={接口1,接口2,....};所有接口
'  EX1={接口1,接口3};扩展部件EX1实现的接口;
'  EX2={接口5,接口6,接口8,接口9};扩展部件EX2实现的接口;
'  .
'  .
'  .
'  EXn={接口n,.....};扩展部件EX1实现的接口;
'  EXn≤ ALL ， EX1+EX2+...+EXn ≤ ALL ， EXn∩EXn-1 = 空，
'3.对于不实现的接口请删除其它接口声明；

'以上说明如看不懂请咨询重庆中联总部研发中心
 
Public Sub LogInAfter()
'功能：登陆航导台之后

End Sub

Public Sub LogOutBefore()
'功能：关闭航导台或者注销当前
    
End Sub

Public Function QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal str卡号 As String, ByRef rsInfo As ADODB.Recordset) As Boolean
'说明见主插件模版
    QueryPatiInfo = True
End Function

Public Function CommitPatiInfo(ByVal str卡号 As String, ByRef rsInfo As ADODB.Recordset) As Boolean
'说明见主插件模版
    CommitPatiInfo = True
End Function
 
Public Function ClinicReceive(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'说明见主插件模版
    ClinicReceive = True
End Function

Public Function ClinicFinish(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'说明见主插件模版
    ClinicFinish = True
End Function

Public Function DiagnosisEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1) As Boolean
'说明见主插件模版
    DiagnosisEnter = True
End Function

Public Function DiagnosisOtherEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1) As Boolean
'说明见主插件模版
    DiagnosisOtherEnter = True
End Function

Public Function DiagnosisDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional ByVal int场合 As Integer = -1) As Boolean
'说明见主插件模版
    DiagnosisDeleted = True
End Function

Public Function AdviceEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional ByVal int场合 As Integer = -1) As Boolean
'说明见主插件模版
    AdviceEnter = True
End Function

Public Function AdviceEditAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional ByVal int场合 As Integer = -1) As Boolean
'说明见主插件模版
    AdviceEditAfter = True
End Function

Public Function AdviceSave(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int场合 As Integer = -1) As Boolean
'说明见主插件模版
    AdviceSave = True
End Function

Public Function AdviceSend(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngSendID As Long) As Boolean
'说明见主插件模版
    AdviceSend = True
End Function

Public Function AdviceSendEnd(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strSendIDs As String) As Boolean
'说明见主插件模版
    AdviceSendEnd = True
End Function

Public Function AdviceRowChange(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int场合 As Integer = -1) As Boolean
'说明见主插件模版
    AdviceRowChange = True
End Function

Public Function AdviceDeletBefor(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'说明见主插件模版
    AdviceDeletBefor = True
End Function

Public Function AdviceDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'说明见主插件模版
    AdviceDeleted = True
End Function

Public Function AdviceRevoked(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'说明见主插件模版
    AdviceRevoked = True
End Function

Public Function AdviceRevokedBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int场合 As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    AdviceRevokedBefore = True
End Function
 
Public Function AdviceRollSend(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int场合 As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    AdviceRollSend = True
End Function
 
Public Function AdviceRollSendBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int场合 As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    AdviceRollSendBefore = True
End Function

Public Function AddEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'说明见主插件模版
    AddEMRBefore = True
End Function

Public Function BeforeEMRDocSave(ByVal lngSource As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFileId As Long, ByVal strTitle As String, ByVal strContent As String, ParamArray arrInput() As Variant) As Boolean
'功能：电子病历保存前调用
'参数：lngSource 1 门诊 2 住院
'      lngPatientID=病人ID,lngPageId=就诊ID/
'      lngFileId=病历文件列表.id
'      arrInput是数组：扩展参数用于扩展使用，通过下标访一直接问时可能会报两种错，9－下标越界/13－类型不匹配：均表示参数没有传入
'           所以访问参数前应先判断：If TypeName(arrInput(i)) <> "Error" Then 满足这个条件后才使用这个参数，否则就认为这个参数没有传入
'           arrInput 目前未传入任何参数
    BeforeEMRDocSave = True: Exit Function
End Function

Public Function SignEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'说明见主插件模版
    SignEMRBefore = True
End Function

Public Sub SignEMRAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant)
'说明见主插件模版
End Sub

Public Function PathImportBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径ID As Long, Optional ByVal lng版本号ID As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng疾病ID As Long, Optional ByVal lng诊断ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    PathImportBefore = True
End Function

Public Function PathImportAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径ID As Long, Optional ByVal lng版本号ID As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng疾病ID As Long, Optional ByVal lng诊断ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    PathImportAfter = True
End Function

Public Function PathEvaluateBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long, Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
   PathEvaluateBefore = True
End Function

Public Function PathEvaluateAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long, Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
   PathEvaluateAfter = True
End Function

Public Function DiagnosisEnterIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal int诊断类型 As Integer, Optional ByVal int诊断次序 As Integer, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    DiagnosisEnterIn = True
End Function

Public Function GetFormHwnd(Optional ByVal strReserve As String) As Long
'说明见主插件模版
    GetFormHwnd = 0
End Function

Public Function PatiInfoSaveBefore(ByVal lng病人ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    PatiInfoSaveBefore = True
End Function

Public Function PatiInfoSaveAfter(ByVal lng病人ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    PatiInfoSaveAfter = True
End Function

'--业务评价接口
Public Function PatiPrePayAfter(ByVal lng病人ID As Long, ByVal int预交类型 As Integer, ByVal lng预交ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    PatiPrePayAfter = True
End Function

Public Function OutPatiRegisterAfter(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    OutPatiRegisterAfter = True
End Function
 
Public Function OutPatiExseAfter(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    OutPatiExseAfter = True
End Function
  
Public Function OutPatiInvoicePrintAfter(ByVal lng病人ID As Long, ByVal str结帐IDs As String, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    OutPatiInvoicePrintAfter = True
End Function
 
Public Function OutPatiMedicineAfter(ByVal lng病人ID As Long, ByVal strNO As String, ByVal int单据 As Integer, ByVal lng药房ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    OutPatiMedicineAfter = True
End Function
 
Public Function DrugStuffWorkNoramal(ByVal lngModual As Long, ByVal strFuncName As String, ByVal lng库房ID As Long, ByVal strNO As String, ByVal int单据 As Integer, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    DrugStuffWorkNoramal = True
End Function

Public Function InPatiCheckInAfter(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    InPatiCheckInAfter = True
End Function
 
Public Function InPatiCheckInBranchAfter(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    InPatiCheckInBranchAfter = True
End Function
 
Public Function InPatiCheckOutAfter(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    InPatiCheckOutAfter = True
End Function

Public Function InPatiCashierAfter(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    InPatiCashierAfter = True
End Function

Public Sub AdviceBeforeSend(ByVal strEndDate As String, ByRef rsData As ADODB.Recordset, ByRef strAdviceIDs As String, ByRef strMsg As String)
'说明见主插件模版
End Sub

Public Function AdviceCheckSendFee(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal cur合计 As Currency, Optional ByVal int场合 As Integer, Optional strReserve As String) As Boolean
'说明见主插件模版
    AdviceCheckSendFee = True
End Function


Public Function PatiFeeAuditingAndCancelCheck(ByVal lngModual As Long, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal blnCancel As Boolean = False, _
    Optional strExpendXML As String) As Boolean
'说明见主插件模版
    PatiFeeAuditingAndCancelCheck = True
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'血库相关接口
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function BloodSendAfter(ByVal lngAdviceID As Long, ByVal strBloodIDs As String, Optional strReserve As String) As Boolean
'说明见主插件模版
    BloodSendAfter = True
End Function

Public Function BloodSendReverseAfter(ByVal lngAdviceID As Long, ByVal strBloodIDs As String, Optional strReserve As String) As Boolean
'说明见主插件模版
    BloodSendReverseAfter = True
End Function

Public Function BloodScrapAfter(ByVal strBloodIDs As String, Optional ByVal intDelete As Integer = 0) As Boolean
'说明见主插件模版
    BloodScrapAfter = True
End Function

Public Function BloodStoreImport(ByRef lng入库方式 As Long, ByRef rsData As ADODB.Recordset)
'说明见主插件模版
    BloodStoreImport = True
End Function

Public Function EMPI_QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal rsIn As ADODB.Recordset, ByRef rsOut As ADODB.Recordset, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    EMPI_QueryPatiInfo = False
End Function

Public Function EMPI_AddPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngClinicID As Long, ByRef strMsg As String, Optional ByRef strReserve As String) As Long
'说明见主插件模版
    EMPI_AddPatiInfo = 1
End Function

Public Function EMPI_ModifyPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngClinicID As Long, ByRef strMsg As String, Optional ByRef strReserve As String) As Long
'说明见主插件模版
    EMPI_ModifyPatiInfo = 1
End Function

Public Function PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    PatiValiedCheck = True
End Function

Public Function PatiRegTriageCheck(lngSys As Long, ByVal lngModule As Long, bytFun As Byte, ByVal lng挂号ID As Long, _
    Optional ByVal str医生 As String, Optional ByVal str诊室 As String, Optional str新号码 As String = "", Optional lng新出诊记录ID As Long = 0) As Boolean
'说明见主插件模版
    PatiRegTriageCheck = True
End Function

Public Function CheckChargeItem(lngSys As Long, ByVal lngModule As Long, ByVal intType As Integer, ByVal intMode As Integer, _
    ByRef rsDetail As ADODB.Recordset, Optional strExpend As String = "") As Boolean
'说明见主插件模版
    CheckChargeItem = True
End Function

Public Function CheckMecInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByRef colCtl As Collection, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    CheckMecInfo = True
End Function

Public Function LoadMecInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatiType As Long, ByRef colCtl As Collection) As Boolean
'说明见主插件模版
    LoadMecInfo = True
End Function

Public Function AdviceMakeFee(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strCurFee As String, ByVal rsNowFee As ADODB.Recordset, ByRef dblCount As Double, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
        dblCount = -1
        AdviceMakeFee = True
End Function

Public Function AdviceLoadApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng就诊ID As String, ByVal lng文件ID As Long, ByRef strXML As String, Optional ByVal lng医嘱ID As Long, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    AdviceLoadApplyCustom = True
End Function

Public Function AdviceSaveApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng就诊ID As String, ByVal lng文件ID As Long, ByRef strXML As String, ByRef objWeb As Object, Optional ByVal lng医嘱ID As Long, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    AdviceSaveApplyCustom = True

End Function

Public Function AdviceCheckApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng就诊ID As String, ByVal lng文件ID As Long, ByVal strXML As String, ByRef objWeb As Object, Optional ByVal lng医嘱ID As Long, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    AdviceCheckApplyCustom = True
End Function

Public Function CheckPathInItem(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, ByRef lng路径项目ID As Long, ByRef str分类 As String, Optional ByVal int场合 As Integer = -1, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    CheckPathInItem = True
End Function

Public Function InPatiCancelOutBefore(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    InPatiCancelOutBefore = True
End Function

Public Function InPatiCheckInCancel(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngType As Long, Optional ByVal strReserve As String) As Boolean
'说明见主插件模版
    InPatiCheckInCancel = True
End Function

Public Function TransfusionCall(ByVal lngSys As Long, ByVal lngModual As Long, ByVal str挂号单 As String, ByVal lng流水号 As Long, ByRef strContent As String, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    TransfusionCall = True
End Function

Public Function TransfusionShowPatiList(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lng科室id As Long, ByVal BeginTime As Date, ByVal EndTime As Date, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    TransfusionShowPatiList = True
End Function

Public Function ShowClinicHelp(lngOwnHwnd As Long, ByVal int场合 As Integer, ByVal int病人来源 As Integer, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str项目IDs As String, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    ShowClinicHelp = False
End Function

Public Function SetPassDiag(ByVal lngPatientID As Long, ByVal lng主页ID As Long, ByRef rsDiag As ADODB.Recordset) As Boolean
'说明见主插件模版
    SetPassDiag = True
End Function

Public Function SetPassExams(ByVal lngPatientID As Long, ByVal lng主页ID As Long, ByVal lng挂号ID As Long, ByRef strExamsInfo As String) As Boolean
'说明见主插件模版
    SetPassExams = True
End Function

Public Function BeforeVitalSignsSave(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng文件ID As Long, ByVal strDataXML As String) As Boolean
'说明见主插件模版
    BeforeVitalSignsSave = True
End Function

Public Sub AfterSelectorReady(ByVal intType As Integer, ByVal strTitle As String, ByVal objList As Object, _
    Optional ByVal frmMain As Object, Optional ByVal strExt As String)
'说明见主插件模版
End Sub


Public Function AdviceEditBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int场合 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str挂号单 As String, ByVal lng前提ID As Long, ByVal int婴儿 As Integer, ByVal lng界面科室ID As Long, ByVal lng危急值ID As Long, ByVal lng会诊医嘱ID As Long, Optional ByRef strReserve As String) As Boolean
'说明见主插件模版
    AdviceEditBefore = True
End Function

Public Function PeisDelItemCheck(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngTaskId As Long, _
    ByVal lngListId As Long, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'功能：体检项目页签，删除行时调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngTaskId=任务id
'      lngListID=清单id
'      strMsg   提示信息，接口内部组织返回给主程序进行提示，空串则不提示
'      strReserve   扩展参数
'返回：
    
    strMsg = "" '当需要控制不能删除行时，返回strMsg为提示的信息
    PeisDelItemCheck = True
End Function

Public Function PeisAbortSignInCheck(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngTaskId As Long, _
    ByVal lngListId As Long, ByVal strNO As String, ByVal lngRecType As Long) As Boolean
'功能：体检中心管理，取消报到时调用，用于检查卫材发料状态是从程序校验还是接口校验返回
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngTaskId=任务id
'      lngListID=清单id
'      strNo   费用单据号
'      lngRecType   费用记录性质
'返回：true-卫材发料检查通过 false-卫材发料检查不通过
 
    PeisAbortSignInCheck = True
End Function

