VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBloodPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'功能：扩展接口插件模板
'如果要做限制使用 请实现 GetUserName 方法
Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合  调用场合:-1
'      病历调用 lngModual=1070，无int场合参数
    Set gcnOracle = cnOracle
End Sub

Public Function GetFormCaption(ByVal lngSys As Long, ByVal lngModual As Long) As String
'功能：在科室配血管理中添加卡片的名称，逗号分割。
    Dim strTmp As String, strThis As String
    
    Exit Function '如使用此功能请修改接口内部实现
    
    'strTmp = GetFormCaptionEx(lngSys, lngModual)
    
    Select Case lngModual
    Case m科室配血管理
        strThis = "Test1,Test2"
    End Select
    GetFormCaption = strThis & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Function GetForm(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'功能：根据索引值返回对象。工作站窗体加载时调用
    
    Exit Function '如使用此功能请修改接口内部实现
    
    Select Case lngModual
    Case m科室配血管理
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            'Set GetForm = GetFormEx(lngSys, lngModual, strName)
        End Select
    End Select
End Function

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1, Optional strReserve As String) As String
'功能：获取外挂功能列表；该接口一般在初始化之后调用一次，调用时机和频度可能因不同的主程序而不同
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号，病历调用时 lngModual=1070，无int场合参数
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS),3-门诊挂号调用,4-门诊收费调用,5-门诊医嘱下达诊断扩展,6-首页诊断扩展
'      strReserve=保留参数,用于扩展使用

'返回：外挂功能列表，格式为"功能名1,功能名2,功能名3,..."
'       "Auto"   关键字支持新版护士站-病人事务窗口：病人医嘱(门诊和住院)Dock页签界面、临床路径Dock页签界面，医嘱附费Dock页签界面，当切换病人后会自动执行该功能，该功能只充许有一个
'       "InTool" 表示当前这个功能，是否是直接显示还是以下拉方式二级菜单显示，
'说明：当插件在人员管理模块调用时只传入lngSys和lngModual，并且其中的lngModual＝0

'strReserve 扩展后被用于控件功能在ZLHIS工作站中菜单的显示样式，如果不对此参数赋值，则仍沿用以前返回值方式
'应用场合:
'   1.科室配血管理
'当窗口切换当对应页签后生效:
'      a.常规复查Dock页签界面，[菜单栏－医嘱、工具按钮下;工具栏；右键弹出随 医嘱  菜单弹出]；
'      b.配血记录Dock页签界面，[菜单栏－路径、工具 按钮下；工具栏；右键弹出随 路径 菜单弹出]；
'格式为: xml结构 "<menubar>...</menubar><toolbar>...</toolbar><popbar>...</popbar><hotkey>...</hotkey>"
'menubar   窗体主界面菜单栏"工具"中展示样式，结点内容逗号分割，"功能1,功能2,功能3,功能4"，
'          每个功能前可以前缀 “|:Auto:InTool:功能1”
'          "|"竖线，表示当前位进行菜单的分割就是菜单之前的分割线；
'          "Auto"   关键字支持新版护士站-病人事务窗口：病人医嘱(门诊和住院)Dock页签界面、临床路径Dock页签界面，医嘱附费Dock页签界面，当切换病人后会自动执行该功能，该功能只充许有一个
'          "InTool" 表示当前这个功能，是否是直接显示还是以下拉方式二级菜单显示，
'toolbar  窗体主界面工具栏中展示样式，结点内容逗号分割，"功能1,功能2,功能3,功能4"，
'         每个功能前可以前缀 “|:InTool:功能1”
'         "|"竖线，表示当前位进行菜单的分割就是菜单之前的分割线；
'         "InTool" 表示当前这个功能，是否是直接显示还是以下拉方式二级菜单显示，
'popbar   右键弹出菜单展示样式，结点内容逗号分割，"功能1,功能2,功能3,功能4"，
'         每个功能前可以前缀 “|:InTool:功能1”
'         "|"竖线，表示当前位进行菜单的分割就是菜单之前的分割线；
'         "InTool" 表示当前这个功能，是否是直接显示还是以下拉方式二级菜单显示，
'hotkey   窗体中快捷建，功能名|KeyCode|Shift 组合成成，例："快捷键功能1|116|0,快捷键功能2|123|0,快捷键功能3|49|4
'         原理：功能由 父窗体的Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)事件进行触发
'         该功能仅门诊医生工作站医嘱下达界面-医嘱编辑 页签生效。
    Dim strFuncNames As String '所有功能名称串，逗号分割，可含一个 Auto关键字，接口中应该对此变量赋值，以兼容以前老版本产品
    
    Dim str菜单栏 As String '菜单栏中放置的功能，菜单栏中应该包含所有功能名， str菜单栏 = str工具栏 ∪ str弹出栏
    Dim str工具栏 As String
    Dim str弹出栏 As String
    Dim str快捷键 As String '固定格式  F5--快捷键功能1;F12--快捷键功能2,Alt+1 --快捷键功能3  对应的串 ："快捷键功能1|116|0,快捷键功能2|123|0,快捷键功能3|49|4"
    
    Select Case lngModual
    Case m科室配血管理_常规复查
        '■科室配血管理_常规复查界面
        strFuncNames = "功能1,功能2,功能3,功能4"
        
        str菜单栏 = "功能1,功能2,功能3,功能4"
        str工具栏 = "InTool:功能1,|:InTool:功能3"
        str弹出栏 = "InTool:功能1"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
    Case m科室配血管理_配血记录
        '■科室配血管理_配血记录界面
        strFuncNames = "功能1,功能2,功能3,功能4"
        
        str菜单栏 = "功能1,功能2,功能3,功能4"
        str工具栏 = "InTool:功能1,|:InTool:功能3"
        str弹出栏 = "InTool:功能1"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
    Case m科室发血管理
        '■科室发血管理界面
        strFuncNames = "功能1,功能2,功能3,功能4"
        
        str菜单栏 = "功能1,功能2,功能3,功能4"
        str工具栏 = "InTool:功能1,|:InTool:功能3"
        str弹出栏 = "InTool:功能1"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
    Case m血液输血反应
        '■输血反应主界面
        strFuncNames = "功能1,功能2,功能3,功能4"
        
        str菜单栏 = "功能1,功能2,功能3,功能4"
        str工具栏 = "InTool:功能1,|:InTool:功能3"
        str弹出栏 = "InTool:功能1"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
    End Select
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int场合 As Integer = -1, Optional ByVal strXML As String = "")
'功能：执行某个外挂功能
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号，病历调用时 lngModual＝1070，无int场合参数
'      strFuncName=要执行的功能名，该功能名之前由GetFuncNames函数返回
'      lngPatiID=当前病人ID
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如科室配血管理为，科室发血管理为，输血反应为
'      strReserve=保留参数,用于扩展使用
'      int场合  调用场合:0-,1-科室配血调用,2-科室发血调用,3-输血反应调用

'      strXml为XML格式的字符串,例：<ROOT><XXX>XXX</XXX></ROOT>
'       当门诊和住院医嘱界面调用时，strXml传入的节点：诊疗项目名称
'       场合为5和6时，strXml传入的节点：诊断ID、疾病ID

'特殊说明：当插件在人员管理模块调用时只传入lngSys、lngModual、strFuncName、lngPatiID、varRecId、varKeyId，
'          lngPatiID就是人员id，如果lngPatiID＝0表示界面上未选中人员；其中的varRecId和varKeyId等值都是传的0

    '注意：在采用分离插件后，功能的执行时，应该先对功能明进行判断。
    '先检查功能是否在本插件（zlPlugIn）中，然后再去执行扩展插件
    
    Select Case lngModual
        Case m科室配血管理_常规复查
            MsgBox strFuncName & "执行成功！" '科室配血管理页面
        Case m科室配血管理_配血记录
            MsgBox strFuncName & "执行成功！" '科室配血管理页面
        Case m科室发血管理
            MsgBox strFuncName & "执行成功！" '科室发血管理页面
        Case m血液输血反应
            'lngPatiID和varRecId为0；varKeyId:输血反应记录.收发ID
            MsgBox strFuncName & "执行成功！" '血液输血反应页面
    End Select
End Sub

Public Function GetUserName() As String
'功能：获取单位名称
'说明：用户名空时表示不限制
'例  GetUserName = "中联技术中心医院信息系统"

    GetUserName = ""
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

Public Function GetButtomName(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String) As String
'功能：添加到菜单和工具栏上的按钮名称，菜单上级为 扩展，工具栏上直接加按钮
'参数：strName 当前窗体名称
        'Exit Function '如使用此功能请修改接口内部实现
    If objForm Is Nothing Then Exit Function
    GetButtomName = objForm.GetInSideFunc
 
End Function

Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口终止
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合=-1

End Sub

Public Function AfterBloodPrepareRecept(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng申请ID As Long) As Boolean
'功能：接受配血后调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      intMode：1-接受申请，2-取消接受申请
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lng申请ID=医嘱id(病人医嘱状态.医嘱id)

    AfterBloodPrepareRecept = True
End Function

Public Function AfterBloodVerify(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lng配发ID As Long, _
    ByVal strIDs As String, Optional ByVal strTime As String = "") As Boolean
'说明见主插件模版
    AfterBloodVerify = True
End Function

Public Function AfterSendBlood(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lng配发ID As Long, _
        ByVal strIDs As String) As Boolean
'说明见主插件模版
    AfterSendBlood = True
End Function

Public Function AfterBloodPurchase(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal strNO As String) As Boolean
'说明见主插件模版
    AfterBloodPurchase = True
End Function

Public Function AfterBloodReaction(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lng收发id As String) As Boolean
'说明见主插件模版
    AfterBloodReaction = True
End Function

Public Function AfterBloodScrap(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal strNO As String) As Boolean
'说明见主插件模版
    AfterBloodScrap = True
End Function

Public Function BloodStoreImport(ByRef lng入库方式 As Long, ByRef rsData As ADODB.Recordset)
'说明见主插件模版
    BloodStoreImport = True
End Function
