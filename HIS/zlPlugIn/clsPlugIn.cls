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
'功能：外挂功能接口部件，用于一些工作站集成外挂调用功能，并传入相应的业务参数
'编制人：周韬
'编制时间：2010-07-22
'说明：
'  2010-07-23:集成到了门诊/住院医嘱虚拟模块中
'  2010-11-25:增加了对自动调用功能的支持,以及对医技站前提ID的传入
'---------------------------------------------------------------------------
Public gblnMec As Boolean '是否开启扩展部件检查病案首页
Public gblnLoadMec As Boolean '是否开启扩展部件加载病案首页数据

Public gblnfrmMec As Boolean '是否开启扩展部件加载病案首页附页
Public gcolMeRec As New Collection    '病案附页集合

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      病历调用 lngModual=1070，无int场合参数
    If Not gblnInited Then
        Set gcnOracle = cnOracle
        Call HandlePlugIn(1, lngSys, lngModual, cnOracle, int场合)
        gblnInited = True
    End If
    gblnMec = True '开启自定义检查病案首页数据
    gblnLoadMec = True '开启自定义加载病案首页数据
    
    gblnfrmMec = False '开启自定义显示病案首页附页
End Sub

Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口终止
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    Call HandlePlugIn(4, lngSys, lngModual, , int场合)
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1, Optional strReserve As String) As String
'功能：获取外挂功能列表；该接口一般在初始化之后调用一次，调用时机和频度可能因不同的主程序而不同
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号，病历调用时 lngModual=1070，无int场合参数
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS),3-门诊挂号调用,4-门诊收费调用,5-门诊医嘱下达诊断扩展,6-首页诊断扩展,7-门诊结帐;8-住院结帐
'      strReserve=保留参数,用于扩展使用

'返回：外挂功能列表，格式为"功能名1,功能名2,功能名3,..."
'       "Auto"   关键字支持新版护士站-病人事务窗口：病人医嘱(门诊和住院)Dock页签界面、临床路径Dock页签界面，医嘱附费Dock页签界面，当切换病人后会自动执行该功能，该功能只充许有一个
'       "InTool" 表示当前这个功能，是否是直接显示还是以下拉方式二级菜单显示，
'说明：当插件在人员管理模块调用时只传入lngSys和lngModual，并且其中的lngModual＝0

'strReserve 扩展后被用于控件功能在ZLHIS工作站中菜单的显示样式，如果不对此参数赋值，则仍沿用以前返回值方式
'应用场合:
'   1.新版护士站主界面（菜单栏-工具-按钮下；工具栏；床头卡右键菜单）
'   2.包含（门诊/住院医嘱清单、临床路径，医嘱附费）模块嵌入的Dock页签显示的工作站界面切换到医嘱信息页签后窗口功能显示
'     如：门诊医生站-医嘱信息，住院医生站-医嘱信息，老版护士站-医嘱信息，新版护士站病-病人事物窗中-医嘱信息等。
'当窗口切换当对应页签后生效:
'      a.病人医嘱(门诊和住院)Dock页签界面，[菜单栏－医嘱、工具按钮下;工具栏；右键弹出随 医嘱  菜单弹出]；
'      b.临床路径Dock页签界面，[菜单栏－路径、工具 按钮下；工具栏；右键弹出随 路径 菜单弹出]；
'      c.医嘱附费Dock页签界面，[菜单栏－费用、工具 按钮下；工具栏；右键弹出随 费用 菜单弹出]；
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


    Dim strTmp As String, strFunc As String
    Dim strReserveOther As String
    
    Dim strFuncNames As String '所有功能名称串，逗号分割，可含一个 Auto关键字，接口中应该对此变量赋值，以兼容以前老版本产品
    
    Dim str菜单栏 As String '菜单栏中放置的功能，菜单栏中应该包含所有功能名， str菜单栏 = str工具栏 ∪ str弹出栏
    Dim str工具栏 As String
    Dim str弹出栏 As String
    Dim str快捷键 As String '固定格式  F5--快捷键功能1;F12--快捷键功能2,Alt+1 --快捷键功能3  对应的串 ："快捷键功能1|116|0,快捷键功能2|123|0,快捷键功能3|49|4"
    
'==================================================================================================================================================
'***************************************************此区域以下为示例代码可进行调整*****************************************************************
    
    Select Case lngModual
    Case m医嘱附费模块
        '■医技工作站，附费页签
        strFuncNames = "功能1,功能2,功能3,功能4"
        
        str菜单栏 = "功能1,功能2,功能3,功能4"
        str工具栏 = "InTool:功能1,|:InTool:功能3"
        str弹出栏 = "InTool:功能1"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
    Case m门诊医嘱模块
        If int场合 = 0 Then
            '■门诊医生工作站，医嘱页签
            strFuncNames = "功能1,功能2,Auto:功能3,功能4"
        
            str菜单栏 = "功能1,功能2,Auto:功能3,功能4"
            str工具栏 = "InTool:功能1,|:InTool:功能3"
            str弹出栏 = "InTool:功能1"
            str快捷键 = "快捷键功能1|116|0,快捷键功能2|123|0,快捷键功能3|49|4"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar><hotkey>" & str快捷键 & "</hotkey>"
        ElseIf int场合 = 5 Then
            '■门诊医嘱编辑界面,加载诊断右键菜单
            strFuncNames = "功能1,功能2"
        Else
            '■医技工作决，医嘱页签
            strFuncNames = "功能1,功能2,Auto:功能3,功能4"
        
            str菜单栏 = "功能1,功能2,Auto:功能3,功能4"
            str工具栏 = "InTool:功能1,|:InTool:功能3"
            str弹出栏 = "InTool:功能1"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        End If
    Case m住院医嘱模块
        If int场合 = 0 Then
            '■住院医生工作站，医嘱页签
            strFuncNames = "功能1,功能2,Auto:功能3,功能4"
        
            str菜单栏 = "功能1,功能2,Auto:功能3,功能4"
            str工具栏 = "InTool:功能1,|:InTool:功能3"
            str弹出栏 = "InTool:功能1"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        Else
            '■医技工作站，医嘱页签
            strFuncNames = "功能1,功能2,Auto:功能3,功能4"
        
            str菜单栏 = "功能1,功能2,Auto:功能3,功能4"
            str工具栏 = "InTool:功能1,|:InTool:功能3"
            str弹出栏 = "InTool:功能1"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        End If
    Case m临床路径模块
        If int场合 = 0 Then
            '■住院医生工作站，临床路径页签
            strFuncNames = "功能1,功能2,Auto:功能3,功能4"
        
            str菜单栏 = "功能1,功能2,Auto:功能3,功能4"
            str工具栏 = "InTool:功能1,|:InTool:功能3"
            str弹出栏 = "InTool:功能1"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        Else
            '■护士工作站，临床路径页签/新版护士站病人事务窗口，临床路径页签
            strFuncNames = "病人一览,护士评估,护理评分,Auto:护理计划,健康宣教,其它文书,交班报告,配液情况,输液状态,输液巡视,护理巡视,你好,你来,走吧,去了"
            
            str菜单栏 = "病人一览,护士评估,护理评分,InTool:护理计划,健康宣教,其它文书,交班报告,配液情况,输液状态,输液巡视,护理巡视,你好,你来,走吧,去了"
            str工具栏 = "|:InTool:交班报告,InTool:配液情况,InTool:输液状态,InTool:输液巡视,InTool:护理巡视,你好,你来"
            str弹出栏 = "InTool:病人一览,InTool:护士评估,InTool:护理评分,InTool:健康宣教,走吧,去了"
            strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        End If
    
    Case m住院护士站模块
        '■老版护士工作站，医嘱页签/新版护士站病人事务窗口，医嘱页签
        strFuncNames = "病人一览,护士评估,护理评分,Auto:护理计划,健康宣教,其它文书,交班报告,配液情况,输液状态,输液巡视,护理巡视,功能一,功能二,功能三,功能四"
        
        str菜单栏 = "病人一览,护士评估,护理评分,Auto:InTool:护理计划,健康宣教,其它文书,交班报告,配液情况,输液状态,输液巡视,护理巡视,功能一,功能二,功能三,功能四"
        str工具栏 = "|:InTool:交班报告,InTool:配液情况,InTool:输液状态,InTool:输液巡视,InTool:护理巡视,功能二,功能三"
        str弹出栏 = "InTool:病人一览,InTool:护士评估,InTool:护理评分,InTool:健康宣教,功能一,功能四"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
    
    Case m新版护士站
        '■新版护士工作站主界面
        strFuncNames = "病人一览,护士评估,护理评分,Auto:护理计划,健康宣教,其它文书,交班报告,配液情况,输液状态,输液巡视,护理巡视,功能一,功能二,功能三,功能四"
        
        str菜单栏 = "病人一览,护士评估,护理评分,Auto:InTool:护理计划,健康宣教,其它文书,交班报告,配液情况,输液状态,输液巡视,护理巡视,功能一,功能二,功能三,功能四"
        str工具栏 = "|:InTool:交班报告,InTool:配液情况,InTool:输液状态,InTool:输液巡视,InTool:护理巡视,功能二,功能三"
        str弹出栏 = "InTool:病人一览,InTool:护士评估,InTool:护理评分,InTool:健康宣教,功能一,功能四"
        strReserve = "<menubar>" & str菜单栏 & "</menubar><toolbar>" & str工具栏 & "</toolbar><popbar>" & str弹出栏 & "</popbar>"
        
    Case m病历模块
        '■病历书写页面
        strFuncNames = "功能1,功能2"
    Case m人员管理模块
        '■人员管理页面
        strFuncNames = "功能1,功能2"
    Case m临床实验室管理
        '■临床实验室管理主界面
        strFuncNames = "功能1,功能2"
    Case m病人信息管理
        '■病人信息管理主界面
        strFuncNames = "功能1,功能2"
    Case m电子病案打印
        '■电子病案打印
        strFuncNames = "功能1,功能2"
    Case m卫材外购入库管理     '药品卫材模块
        '■卫材外购管理界面
        strFuncNames = "功能1,功能2"
    Case m药品外购入库管理
        '■药品外购管理界面
        strFuncNames = "功能1,功能2"
    Case m药品计划管理
        '■药品计划管理界面
        strFuncNames = "功能1,功能2"
    Case m药品处方发药
        '■药品处方发药界面
        '格式说明：InTool表示放到工具栏，|表示分组
        '例：功能2，功能3要放到工具栏，并且功能2和功能3分组显示
        strFuncNames = "处方发药功能1,InTool:处方发药功能2,|:InTool:处方发药功能3"
    Case m药品部门发药
        '■药品部门发药界面
        '格式说明：InTool表示放到工具栏，|表示分组
        '例：功能2，功能3要放到工具栏，并且功能2和功能3分组显示
        strFuncNames = "部门发药功能1,InTool:部门发药功能2,|:InTool:部门发药功能3"
    Case m输液配置中心
        '■输液配置中心界面
        '格式说明：InTool表示放到工具栏，|表示分组
        '例：功能2，功能3要放到工具栏，并且功能2和功能3分组显示
        strFuncNames = "输液配置中心1,InTool:输液配置中心2,|:InTool:输液配置中心3"
    Case m病人挂号模块
        '■病人挂号管理界面
        strFuncNames = "功能1,功能2,InTool:功能3"
    Case m病人收费模块
        '■病人收费管理界面
        strFuncNames = "功能1,功能2,InTool:功能3"
    Case m住院医生工作站
        '■住院首页右键诊断扩展
        If int场合 = 6 Then
            strFuncNames = "功能1,功能2,功能3"
        End If
    Case m病人结帐处理
        '■病人结帐界面
        strFuncNames = "功能1,功能2,InTool:功能3"
    Case m体检中心管理
        '■体检中心管理
        strFuncNames = "功能1,功能2,InTool:功能3"
    Case m体检总检登记
        '■体检总检登记
        strFuncNames = "功能1,功能2,InTool:功能3"
    Case m体检分科执行
        '■体检分科执行
        strFuncNames = "功能1,功能2,InTool:功能3"
    Case m体检结果登记
        '■体检结果登记
        strFuncNames = "功能1,功能2,InTool:功能3"
    End Select
    
'***************************************************此区域以上为示例代码可进行调整*****************************************************************
'==================================================================================================================================================

    strTmp = ""
    strReserveOther = ""
    Call HandlePlugIn(2, lngSys, lngModual, , int场合, strReserveOther, strTmp)
    If strReserveOther <> "" Then strReserve = strReserve & strReserveOther
    GetFuncNames = strFuncNames & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String, Optional ByVal int场合 As Integer = -1, Optional ByVal strXML As String = "")
'功能：执行某个外挂功能
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号，病历调用时 lngModual＝1070，无int场合参数
'      strFuncName=要执行的功能名，该功能名之前由GetFuncNames函数返回
'      lngPatiID=当前病人ID,人员管理调用时，传入的人员ID， 新版LIS默认传入0
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID，新版LIS默认传入0；挂号调用时，为挂号单号；收费调用时，为收费单号
'           老版体检系统：varRecId=任务id
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如医嘱ID
'           病历调用时 varKeyId=电子病历记录.id，在新增编辑未保存前其值=0,保存后或修改、审订时执行有值；
'           临床路径 varKeyId 病人路径id
'           新版LIS   varKeyId = 标本id ,当前有标本为标本id ,无标本则为0
'           老版体检系统   varKeyId = 清单id
'      strReserve=保留参数,用于扩展使用 当是病历调用时 strReserve＝ 病历文件列表.id 。当场合为5、6调用时,strReserve＝诊断内容
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS),3-挂号调用,4-收费调用,5-门诊医嘱下达诊断扩展,6-首页诊断扩展

'      strXml为XML格式的字符串,例：<ROOT><XXX>XXX</XXX></ROOT>
'       当门诊和住院医嘱界面调用时，strXml传入的节点：诊疗项目名称
'       当门诊医嘱界面调用时，strXml传入的节点：诊疗项目名称、界面科室ID、挂号科室ID
'       当住院医嘱界面调用时，strXml传入的节点：诊疗项目名称、病人科室ID、病人病区ID、界面科室ID、医护科室ID
'       场合为5和6时，strXml传入的节点：诊断ID、疾病ID

'特殊说明：当插件在人员管理模块调用时只传入lngSys、lngModual、strFuncName、lngPatiID、varRecId、varKeyId，
'          lngPatiID就是人员id，如果lngPatiID＝0表示界面上未选中人员；其中的varRecId和varKeyId等值都是传的0

    '注意：在采用分离插件后，功能的执行时，应该先对功能明进行判断。
    '先检查功能是否在本插件（zlPlugIn）中，然后再去执行扩展插件
    
    Select Case lngModual
        Case m门诊医嘱模块
            If int场合 = 0 Then
                MsgBox strFuncName & "执行成功！" '门诊医生站医嘱页面
            ElseIf int场合 = 5 Then
                MsgBox strFuncName & "执行成功！" '门诊医生站诊断扩展
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
        Case m临床实验室管理
            MsgBox strFuncName & "执行成功！"  '临床实验室管理
        Case m病人信息管理
            MsgBox strFuncName & "执行成功！"  '病人信息管理
        Case m电子病案打印
            MsgBox strFuncName & "执行成功！"  '电子病案打印
        Case m病人挂号模块
            MsgBox strFuncName & "执行成功！" '病人挂号管理页面
        Case m病人收费模块
            MsgBox strFuncName & "执行成功！" '病人收费管理页面
        Case m住院医生工作站
            '■住院首页右键诊断扩展执行
            If int场合 = 6 Then
                MsgBox strReserve
            End If
        Case m体检中心管理
            MsgBox strFuncName & "执行成功！" '体检中心管理
        Case m体检总检登记
            MsgBox strFuncName & "执行成功！" '体检总检登记
        Case m体检分科执行
            MsgBox strFuncName & "执行成功！" '体检分科执行
        Case m体检结果登记
            MsgBox strFuncName & "执行成功！" '体检结果登记
        End Select
    Call HandlePlugIn(3, lngSys, lngModual, , int场合, strReserve, strFuncName, lngPatiID, varRecId, varKeyId)
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

Public Function GetFormCaption(ByVal lngSys As Long, ByVal lngModual As Long) As String
'功能：在住院医生工作站中添加卡片的名称，逗号分割。
    Dim strTmp As String, strThis As String
    Exit Function '如使用此功能请修改接口内部实现
    strTmp = GetFormCaptionEx(lngSys, lngModual)
    
    Select Case lngModual
    Case m住院医生工作站
        strThis = "Test1,Test2"
    Case m住院护士工作站
        strThis = "Test1,Test2"
    Case m门诊医生工作站
        strThis = "Test1,Test2"
    Case m医技工作站
        strThis = "Test1,Test2"
    Case m住院医嘱下达
        strThis = "Test1,Test2"
    Case m门诊医嘱下达
        strThis = "Test1,Test2"
    End Select
    GetFormCaption = strThis & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Function GetForm(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'功能：根据索引值返回对象。工作站窗体加载时调用
        Exit Function '如使用此功能请修改接口内部实现
    Select Case lngModual
    Case m住院医生工作站
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            Set GetForm = GetFormEx(lngSys, lngModual, strName)
        End Select
    Case m住院护士工作站
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            Set GetForm = GetFormEx(lngSys, lngModual, strName)
        End Select
    Case m门诊医生工作站
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            Set GetForm = GetFormEx(lngSys, lngModual, strName)
        End Select
    Case m医技工作站
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            Set GetForm = GetFormEx(lngSys, lngModual, strName)
        End Select
    Case m门诊医嘱下达
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            Set GetForm = GetFormEx(lngSys, lngModual, strName)
        End Select
    Case m住院医嘱下达
        Select Case strName
        Case "Test1"
            Set GetForm = New frmDockEx
        Case "Test2"
            Set GetForm = New frmDockEx
        Case Else
            Set GetForm = GetFormEx(lngSys, lngModual, strName)
        End Select
    End Select
End Function

Public Sub RefreshForm(ByVal lngSys As Long, ByVal lngModual As Long, ByRef objForm As Object, ByVal strName As String, ByVal lngPatiID As Long, ByVal str挂号单 As String, ByVal lng主页ID As Long, _
    Optional ByVal blnMoved As Boolean, Optional ByVal lng界面科室ID As Long, Optional ByVal lng医护科室ID As Long, Optional ByVal lng病区ID As Long, Optional ByVal lng科室ID As Long, _
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

Public Sub LogInAfter()
'功能：登陆航导台之后

'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            Call gcolPlugIn(i).LogInAfter
            Err.Clear
        Next i
    End If
End Sub

Public Sub LogOutBefore()
'功能：关闭航导台或者注销当前
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            Call gcolPlugIn(i).LogOutBefore
            Err.Clear
        Next i
    End If
End Sub

Public Function zlSendInfor(ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPatiXML As String, ByVal int场合 As Integer, Optional ByRef strFeedBack As String, Optional ByRef strErrMsg As String, Optional ByRef strReserve As String) As Boolean
'功能：向短信平台发送消息
'返回：True是成功，False是失败
'入参：lngSys,lngModual：当前调用接口的主程序系统号及模块号
'      strPatiXML：包含病人信息的XML字符串
'      int场合：1-发卡时短信验证，2-账户余额变化提醒
'出参：strFeedBack：场合=1时返回验证码以及等待时间,格式:验证码|倒计时秒数
'      strErrMsg：错误信息
'      strReserve：扩展信息

    zlSendInfor = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            zlSendInfor = gcolPlugIn(i).zlSendInfor(lngSys, lngModule, strPatiXML, int场合, strFeedBack, strErrMsg, strReserve)
            Err.Clear
        Next i
    End If
    
'    If int场合 = 1 Then
'        strFeedBack = "111111|60"
'    End If
    zlSendInfor = True
End Function

Public Function DrugSendCustomCheckSet(ByRef strFunName As String) As Boolean
'功能：药品部门发药，自定义审核功能。返回true时部门发药主界面开启自定义审核功能按钮
'参数：strFunName：可以自定义功能名称，返回作为主界面工具栏按钮的名称
'返回：false-不启用，true-启用
    
    '以下用于测试
    MsgBox "部门发药调用PlugIn中DrugSendCustomCheckSet接口执行成功！"
    
    strFunName = "医嘱审核"
    
    DrugSendCustomCheckSet = False
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugSendCustomCheckSet = gcolPlugIn(i).DrugSendCustomCheckSet()
            Err.Clear
        Next i
    End If
End Function

Public Function DrugSendCustomCheck(ByVal str收发ids As String, ByRef str返回收发ids As String, Optional ByVal strReserve As String) As Boolean
'功能：药品部门发药主界面调用，按收发id传送，主要通过收发ID来获取其他信息
'参数：
'      str收发ids=收发id串，收发id1,收发id2...
'      str返回收发ids=返回主界面的收发ids，用来控制界面数据是否发药，默认返回的收发ids不发药
'      strReserve=扩展参数
'      返回：True-继续发药，False-不继续发药
    Dim str医嘱ids As String    '主医嘱id和对应的药嘱id：主医嘱id,药嘱id1,药嘱id2...|主医嘱id,药嘱id1,药嘱id2...
    Dim strSql As String
    Dim lngCount As Long
    
'    '需要用到医嘱id的参考SQL
'    strSql = "Select Distinct a.相关id, a.Id as 医嘱id " & _
'        " From 病人医嘱记录 A, 住院费用记录 B, 药品收发记录 D, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) C " & _
'        " Where a.Id = b.医嘱序号 And b.Id = d.费用id And d.Id = c.Column_Value " & _
'        " Order By a.相关id, a.Id "
'    Set rstemp = zlDatabase.OpenSQLRecord(strSql, "DrugBeforeSendBySumID", str收发ids)
    
    ''''以下用于测试
    MsgBox "部门发药调用PlugIn中DrugSendCustomCheck接口执行成功！"
        
    If InStr(str收发ids, ",") = 0 Then
        '只有一个收发ID就直接返回
        str返回收发ids = str收发ids
    Else
        '有多个收发id，返回一半
        For lngCount = 0 To UBound(Split(str收发ids, ","))
            If lngCount Mod 2 = 0 Then
                str返回收发ids = IIf(str返回收发ids = "", "", str返回收发ids & ",") & Split(str收发ids, ",")(lngCount)
            End If
        Next
    End If
    ''''以上用于测试
        
    DrugSendCustomCheck = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugSendCustomCheck = gcolPlugIn(i).DrugSendCustomCheck(str收发ids, str返回收发ids, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DrugBeforeSendByRecipe(ByVal lng库房ID As Long, ByVal strNos As String, Optional ByVal strReserve As String) As Boolean
'功能：药品处方发药前调用，按处方发药
'参数：lng库房ID  库房ID；
'      strNOs 一次发药可能多个处方，处方类型和处方号，格式；单据,NO|单据,NO...
'      strReserve=扩展参数
'      返回：True-继续发药，False-不继续发药
    Dim str医嘱ids As String    '主医嘱id和对应的药嘱id：主医嘱id,药嘱id1,药嘱id2...|主医嘱id,药嘱id1,药嘱id2...
    Dim strSql As String
    
'    '需要用到医嘱id的参考SQL
'    strSql = "Select Distinct a.相关id, a.Id " & _
'        " From 病人医嘱记录 A, 门诊费用记录 B, 药品收发记录 C, Table(Cast(f_Str2list2([2]) As Zltools.t_Strlist2)) D " & _
'        " Where a.Id = b.医嘱序号 And b.Id = c.费用id And c.单据 = d.C1 And c.No = d.C2 And c.库房id = [1] " & _
'        " Order By a.相关id, a.Id "
'    Set rstemp = zlDatabase.OpenSQLRecord(strSql, "DrugBeforeSendByRecipe", lng库房ID, strNOs)

    '以下用于测试
    If MsgBox("处方发药前调用PlugIn中DrugBeforeSendByRecipe接口执行成功！待发药信息如下，是否继续发药？" & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "发药处方=" & strNos & vbCrLf, _
        vbQuestion + vbYesNo + vbDefaultButton2, "zlPlugIn") = vbNo Then
        
        Exit Function
    End If

    DrugBeforeSendByRecipe = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugBeforeSendByRecipe = gcolPlugIn(i).DrugBeforeSendByRecipe(lng库房ID, strNos, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DrugListWork(ByVal intType As Integer, ByVal str药品IDs As String, Optional ByVal strReserve As String) As Boolean
'功能：药品目录通用业务接口
'      在完成药品目录中药品规格主要业务功能后调用该接口（增删改停用启用）
'参数：intType：业务类型，0-新增，1-修改，2-删除，3-停用，4-启用
'      str药品IDs ：药品id串，格式“药品id1,药品id2,...”
'      strReserve=扩展参数
'说明：业务类型=2修改时，正常情况是规格属性发生变化，也可能是因为该规格对应的品种属性修改，具体查询后比对
'      业务类型=3删除时，该接口是在ZLHIS功能执行后调用，在事务不同步情况下该药品已经在ZLHIS中删除了
    Dim strType As String
    
    Select Case intType
        Case 0
            strType = "新增"
        Case 1
            strType = "修改"
        Case 2
            strType = "删除"
        Case 3
            strType = "停用"
        Case 4
            strType = "启用"
    End Select
    
    MsgBox "调用接口DrugListWork成功！" & vbCrLf & _
        "intType=" & intType & "-" & strType & vbCrLf & _
        "str药品IDs=" & str药品IDs

   
    DrugListWork = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugListWork = gcolPlugIn(i).DrugListWork(intType, str药品IDs, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DrugStuffWork(ByVal intType As Integer, ByVal lng库房ID As Long, ByVal strNO As String, _
    ByVal int单据 As Integer, Optional ByVal strDate As String, _
    Optional ByVal strReserve As String) As Boolean
'功能：药品卫材流通通用业务接口
'      在完成流通主要业务功能后调用该接口（删除业务是在之前调用）
'参数：intType：业务类型，0-新增，1-修改，2-删除，3-审核，4-冲销
'      lng库房ID ：库房ID
'      strNO：单据号
'      int单据：单据类型
'      药品：0-计划；1-外购入库；2-自制入库；3-协药入库；4-其他入库；6-库房移出；7-部门领用；11-其他出库；12-盘点；
'      卫材：0-计划；15-材料外购入库,16-材料自制入库,17-材料其他入库,19-材料移库,20-部门材料领用,21-材料其他出库,22-材料盘点
'      strDate：单据日期，只有冲销时传入（因为允许部分冲销，用时间来区分本次冲销的是哪些数据，如果不传就是不允许部分冲销）
'      strReserve=扩展参数
'说明：1)根据传入的单据区分药品卫材和具体的业务类型
'      2)该接口可以通用于药品卫材所有的流通（包括计划，不包括发药/发料）的单据主要业务功能（增删改，审核，冲销），目前实现的功能如下
'      3)10.35.120 在药品申领、移库、外购、其他入库、其他出库、领用、盘点的审核功能中调用了接口
    
    '目前只处理审核业务
    If intType <> 3 Then Exit Function
    
    MsgBox "调用接口DrugStuffWork成功！" & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "NO=" & strNO & vbCrLf & _
        "单据=" & int单据

   
    DrugStuffWork = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugStuffWork = gcolPlugIn(i).DrugStuffWork(lng库房ID, strNO, int单据, strDate, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DrugBeforeSendBySumID(ByVal lng库房ID As Long, ByVal str收发ids As String, Optional ByVal strReserve As String) As Boolean
'功能：药品部门发药前调用，按收发id传送，主要通过收发ID来获取其他信息
'参数：lng库房ID  库房ID；
'      str收发ids 收发id串，收发id1,收发id2...
'      strReserve=扩展参数
'      返回：True-继续发药，False-不继续发药
    Dim str医嘱ids As String    '主医嘱id和对应的药嘱id：主医嘱id,药嘱id1,药嘱id2...|主医嘱id,药嘱id1,药嘱id2...
    Dim strSql As String
    
'    '需要用到医嘱id的参考SQL
'    strSql = "Select Distinct a.相关id, a.Id as 医嘱id " & _
'        " From 病人医嘱记录 A, 住院费用记录 B, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) C " & _
'        " Where a.Id = b.医嘱序号 And b.Id = c.Column_Value " & _
'        " Order By a.相关id, a.Id "
'    Set rstemp = zlDatabase.OpenSQLRecord(strSql, "DrugBeforeSendBySumID", str收发ids)
    
    '以下用于测试
    If MsgBox("部门发药前调用PlugIn中DrugBeforeSendBySumID接口执行成功！是否继续发药？", _
        vbQuestion + vbYesNo + vbDefaultButton2, "zlPlugIn") = vbNo Then
        Exit Function
    End If
        
    DrugBeforeSendBySumID = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugBeforeSendBySumID = gcolPlugIn(i).DrugBeforeSendBySumID(lng库房ID, str收发ids, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function StuffReturnByID(ByVal lng库房ID As Long, ByVal strIDs As String, ByVal dteSendTime As Date, Optional ByVal strReserve As String) As Boolean
'功能：卫材退料后调用
'参数：lng库房ID  库房ID；
'      strIDs 一次退料操作可能退多个卫材记录，格式；收发id,退料数量|收发id,退料数量...
'      dteSendTime 退料时间，一条收发id可能退多次，加上时间可以用于区别是哪次退料；
'      strReserve=扩展参数
       
    MsgBox "退料后调用PlugIn中StuffReturnByID接口执行成功！退料信息如下..." & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "退料记录=" & strIDs & vbCrLf & _
        "退料时间=" & Format(dteSendTime, "yyyy-mm-dd hh:mm:ss")
   
    StuffReturnByID = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            StuffReturnByID = gcolPlugIn(i).StuffReturnByID(lng库房ID, strIDs, dteSendTime, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function StuffSendByRecipe(ByVal lng库房ID As Long, ByVal strNos As String, ByVal dteSendTime As Date, Optional ByVal strReserve As String) As Boolean
'功能：卫材处方发料后调用，按处方发料
'参数：lng库房ID  库房ID；
'      strNOs 一次发料可能多个处方，处方类型和处方号，格式；单据,NO|单据,NO...
'      dteSendTime 发料时间，可能一个处方退料了之后再发，加上时间用于区别是哪次发料；
'      strReserve=扩展参数
       
    MsgBox "处方发料后调用PlugIn中StuffSendByRecipe接口执行成功！发料信息如下..." & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "发料处方=" & strNos & vbCrLf & _
        "发料时间=" & Format(dteSendTime, "yyyy-mm-dd hh:mm:ss")
   
    StuffSendByRecipe = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            StuffSendByRecipe = gcolPlugIn(i).StuffSendByRecipe(lng库房ID, strNos, dteSendTime, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function StuffSendBySumID(ByVal lng库房ID As Long, ByVal strSumNO As String, Optional ByVal strReserve As String) As Boolean
'功能：卫材发料后调用，按收发id发料
'参数：lng库房ID  库房ID；
'      strSumNO 汇总发料号
'      strReserve=扩展参数
       
    MsgBox "批量发料后调用PlugIn中StuffSendBySumID接口执行成功！发料信息如下..." & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "发料记录=" & strSumNO
        
    StuffSendBySumID = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            StuffSendBySumID = gcolPlugIn(i).StuffSendBySumID(lng库房ID, strSumNO, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal str卡号 As String, ByRef rsInfo As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回病人信息集。病人信息集为动态记录集，在接口中预先写好构建及返回字段模板；如果返回字段中"病人ID"<>0，则表示接口返回前已新建档案，挂号程序直接提取病人ID对应的档案信息，而不直接读取其他字段；
    '入参:lngModule-调用的模块号
    '出参:rsInfo-返回的病人信息集(直接更新该表,字段为:病人ID,姓名,性别,年龄,出生日期,出生地点,身份证号,其他证件,身份,职业,家庭地址,家庭电话,家庭邮编,工作单位,单位邮编,医保号,医疗付款方式,费别,国籍,民族,婚姻状况,区域)
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:33:19
    '说明:准格尔旗区域卫生系统接口(读取病人信息)
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '例:
    '    With rsInfo
    '        .AddNew
    '            !病人ID = Null
    '            !姓名 = "区域卫生" & Int(Rnd * 10)
    '            !性别 = "女"
    '            !年龄 = "44"
    '            !出生日期 = Now - 3000
    '            !出生地点 = "重庆市北碚区万寿桥村1组"
    '            !身份证号 = "80223197405197890"
    '            !其他证件 = Null
    '            !身份 = "工人"
    '            !职业 = "IT"
    '            !家庭地址 = "渝北区工业园区"
    '            !家庭电话 = "67136410"
    '            !家庭邮编 = "5123"
    '            !工作单位 = "中联信息产业"
    '            !单位邮编 = "423"
    '            !医保号 = "42143Yb"
    '            !医疗付款方式 = "生育保险"
    '            !费别 = "普通"
    '            !国籍 = "中国"
    '            !民族 = "汉族"
    '            !婚姻状况 = "已婚"
    '            !区域 = "渝北区"
    '        .Update
    '    End With
    QueryPatiInfo = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            QueryPatiInfo = gcolPlugIn(i).QueryPatiInfo(lngSys, lngModule, str卡号, rsInfo)
            Err.Clear
        Next i
    End If
End Function

Public Function CommitPatiInfo(ByVal str卡号 As String, ByRef rsInfo As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:提交发卡信息
    '入参:str卡号-卡号
    '        rsInfo-返回的病人信息集(直接更新该表,字段为:病人ID,姓名,性别,年龄,出生日期,出生地点,身份证号,其他证件,身份,职业,家庭地址,家庭电话,家庭邮编,工作单位,单位邮编,医保号,医疗付款方式,费别,国籍,民族,婚姻状况,区域)
    '返回:调用成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:33:19
    '说明:准格尔旗区域卫生系统接口(提交发卡信息)
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CommitPatiInfo = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            CommitPatiInfo = gcolPlugIn(i).CommitPatiInfo(str卡号, rsInfo)
            Err.Clear
        Next i
    End If
End Function

Public Function GetMeRecFormCol(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngPatiType As Long) As Collection
    '初始化外挂附页窗体对象的集合，注：zlhis34最多可以扩展3个页签窗体
    
    Dim frmTmp As Form
    Dim colTmp As New Collection
    
    '获取新的外挂附页窗体对象,用于zlhis35以上病案外挂附页支持鼠标滚轮
    Call CreateMecOK
    
    '加载外挂附页
    Set frmTmp = New frmMeRec
        frmTmp.glngSys = lngSys
        frmTmp.glngModule = lngModule
        frmTmp.glngPatiID = lngPatiID
        frmTmp.glngPageID = lngPageID
        frmTmp.glngPatiType = lngPatiType
    colTmp.Add frmTmp, "外挂附页1"
    
    Set frmTmp = New frmMeRec
        frmTmp.glngSys = lngSys
        frmTmp.glngModule = lngModule
        frmTmp.glngPatiID = lngPatiID
        frmTmp.glngPageID = lngPageID
        frmTmp.glngPatiType = lngPatiType
    colTmp.Add frmTmp, "外挂附页2"
    
    Set gcolMeRec = colTmp
    Set GetMeRecFormCol = gcolMeRec

End Function

Public Function CreateMecOK() As Boolean
'功能：病案部件创建与检查
    If gobjMec Is Nothing Then
        On Error Resume Next
        Set gobjMec = CreateObject("zlMedRecPage.clsInOutMedRec")
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjMec Is Nothing Then CreateMecOK = True
End Function

Public Function ClinicReceive(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'功能：病人接诊操作后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）

    ClinicReceive = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            ClinicReceive = gcolPlugIn(i).ClinicReceive(lngSys, lngModual, lngPatientID, lngClinicID)
            Err.Clear
        Next i
    End If
End Function

Public Function ClinicFinish(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Boolean
'功能：病人完成接诊后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）

    ClinicFinish = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            ClinicFinish = gcolPlugIn(i).ClinicFinish(lngSys, lngModual, lngPatientID, lngClinicID)
            Err.Clear
        Next i
    End If
End Function

Public Function DiagnosisEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：门诊病案中输入主要诊断后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）
'      lngDiagID=诊断ID或疾病ID(根据系统参数决定),strDiagDesc=诊断描述
'      lng原诊断id=0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id的值就是修改前的 诊断ID或疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
      
    DiagnosisEnter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DiagnosisEnter = gcolPlugIn(i).DiagnosisEnter(lngSys, lngModual, lngPatientID, lngClinicID, lngDiagID, strDiagDesc, lng原诊断id, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function DiagnosisOtherEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：门诊病案中输入次要诊断后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）
'      lngDiagID=诊断ID或疾病ID(根据系统参数决定),strDiagDesc=诊断描述
'      lng原诊断id=0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id的值就是修改前的 诊断ID或疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    
    DiagnosisOtherEnter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DiagnosisOtherEnter = gcolPlugIn(i).DiagnosisOtherEnter(lngSys, lngModual, lngPatientID, lngClinicID, lngDiagID, strDiagDesc, lng原诊断id, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function DiagnosisDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：门诊病案中删除诊断后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）
'      lngDiagID=诊断ID或疾病ID(根据系统参数决定),strDiagDesc=诊断描述
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    
    DiagnosisDeleted = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DiagnosisDeleted = gcolPlugIn(i).DiagnosisDeleted(lngSys, lngModual, lngPatientID, lngClinicID, lngDiagID, strDiagDesc, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceEnter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：每输入一条医嘱项目时调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      rsAdvice=当前录入的一条或一组医嘱记录（中药配方、成套方案、批量复制医嘱在输入时不调用）
'      字段为表“病人医嘱记录”的大部分字段，包括：ID, 相关id, 序号, 病人来源（1=门诊,2=住院), 病人id, 挂号单/主页ID,
'                姓名, 性别, 年龄, 婴儿, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id, 收费细目id, 天数, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 标本部位,
'                检查方法 , 执行标记, 执行频次, 频率次数, 频率间隔, 间隔单位, 执行时间方案, 计价特性, 执行科室id, 执行性质, 紧急标志, 可否分零,
'                开始执行时间,执行终止时间(住院调用才有，没填时为cdate(0)), 病人科室id, 开嘱科室id, 开嘱医生, 开嘱时间, 挂号单, 前提id, 摘要
'                ,除此之外，门诊调用时，还包含医嘱对应的诊断的疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceEnter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceEnter = gcolPlugIn(i).AdviceEnter(lngSys, lngModual, lngPatientID, lngClinicID, rsAdvice, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceEditAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional ByVal int场合 As Integer = -1) As Boolean
'功能：每输入完一条医嘱项目进行下一条医嘱输入时调用，或者在表格换行时调用(传入的医嘱是上一行的内容－－新开或修改过的医嘱)
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      rsAdvice=当前录入的一条或一组医嘱记录（中药配方、成套方案、批量复制医嘱在输入时不调用）
'      字段为表“病人医嘱记录”的大部分字段，包括：ID, 相关id, 序号, 病人来源（1=门诊,2=住院), 病人id, 挂号单/主页ID,
'                姓名, 性别, 年龄, 婴儿, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id, 收费细目id, 天数, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 标本部位,
'                检查方法 , 执行标记, 执行频次, 频率次数, 频率间隔, 间隔单位, 执行时间方案, 计价特性, 执行科室id, 执行性质, 紧急标志, 可否分零,
'                开始执行时间,执行终止时间(住院调用才有，没填时为cdate(0)), 病人科室id, 开嘱科室id, 开嘱医生, 开嘱时间, 挂号单, 前提id, 摘要
'                ,除此之外，门诊调用时，还包含医嘱对应的诊断的疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)


    AdviceEditAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceEditAfter = gcolPlugIn(i).AdviceEditAfter(lngSys, lngModual, lngPatientID, lngClinicID, rsAdvice, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceSave(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int场合 As Integer = -1) As Boolean
'功能：保存医嘱时(事务开始前)调用,返回False时不保存医嘱
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      rsAdvice=当前病人本次新开或修改的所有医嘱记录，字段同AdviceEnter接口的该参数,除此之外，门诊调用时，还包含医嘱对应的诊断的疾病ID
'               字段 EditState  '1-新增，2－修改
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)


'    记录集 rsAdvice 包含的字段如下：
'    strSQL = "select a.ID,a.相关ID,a.前提ID,a.病人来源,a.病人ID,a.主页ID,a.挂号单,a.婴儿,a.姓名,a.性别,a.年龄,a.病人科室ID,a.序号,a.医嘱状态,a.医嘱期效," & _
'        "a.诊疗类别,a.诊疗项目ID,a.标本部位,a.检查方法,a.收费细目ID,a.天数,a.单次用量,a.总给予量,a.医嘱内容,a.医生嘱托,a.执行科室ID,a.执行频次," & _
'        "a.频率次数,a.频率间隔,a.间隔单位,a.执行时间方案,a.计价特性,a.执行性质,a.执行标记,a.可否分零,a.紧急标志,a.开始执行时间," & _
'        "a.执行终止时间,a.开嘱科室ID,a.开嘱医生,a.开嘱时间,a.摘要,a.用药目的,a.用药理由,a.申请序号,a.手术情况,a.ID as 疾病ID,0 as EditState,a.医嘱内容 as 附项," & _
'        "a.医生嘱托 as Apply,null 诊断信息" & _
'        " from 病人医嘱记录 a where a.id=-1"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel")

'   诊断信息 -- 字段说明，仅诊断医嘱下达保存时rsAdvice记录的第一行有值，表示了门诊医嘱下达界面的诊断表格信息，字符串格式，形式为一个 二维数组，固定字分割符【<SPCOLS><SPROWS>】分割，
'              例：
'                 序号 , 诊断描述, 诊断编码, 疾病编码, 诊断ID, 疾病ID
'                 1<SPCOLS>额窦恶性肿瘤<SPCOLS><SPCOLS>C31.201<SPCOLS>0<SPCOLS>942<SPCOLS>0<SPROWS>2<SPCOLS>类结核型麻风<SPCOLS><SPCOLS>A30.101<SPCOLS>0<SPCOLS>283<SPCOLS>0......

    AdviceSave = True
    
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceSave = gcolPlugIn(i).AdviceSave(lngSys, lngModual, lngPatientID, lngClinicID, rsAdvice, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceSend(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngSendID As Long) As Boolean
'功能：发送医嘱完成后调用，当住院护士同时发送多个病人医嘱时一个病人会调用一次。
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达(门诊没有“发送”模块),1254=住院医嘱发送)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngSendID=发送号(病人医嘱发送.id)

    AdviceSend = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceSend = gcolPlugIn(i).AdviceSend(lngSys, lngModual, lngPatientID, lngClinicID, lngSendID)
            Err.Clear
        Next i
    End If
End Function




Public Function AdviceSendEnd(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strSendIDs As String) As Boolean
'功能：一次发送医嘱完成后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达(门诊没有“发送”模块),1254=住院医嘱发送)
'      strSendIDs 发送号可能有多个，用逗号分割

    AdviceSendEnd = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceSendEnd = gcolPlugIn(i).AdviceSendEnd(lngSys, lngModual, strSendIDs)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceRowChange(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, Optional int场合 As Integer = -1) As Boolean
'功能：医嘱编辑时切换医嘱行时(医嘱编辑界面医嘱表格行变换)调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      rsAdvice=当前病人本次新开的所有医嘱记录，字段同AdviceEnter接口的该参数,除此之外，门诊调用时，还包含医嘱对应的诊断的疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceRowChange = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceRowChange = gcolPlugIn(i).AdviceRowChange(lngSys, lngModual, lngPatientID, lngClinicID, rsAdvice, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceDeletBefor(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'功能：医嘱删除前调用（返回False则取消删除医嘱）
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceDeletBefor = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceDeletBefor = gcolPlugIn(i).AdviceDeletBefor(lngSys, lngModual, lngPatientID, lngClinicID, lngAdviceID, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceDeleted(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'功能：医嘱删除后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceDeleted = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceDeleted = gcolPlugIn(i).AdviceDeleted(lngSys, lngModual, lngPatientID, lngClinicID, lngAdviceID, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceRevoked(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, Optional int场合 As Integer = -1) As Boolean
'功能：医嘱作废后调用，一条医嘱调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

    AdviceRevoked = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceRevoked = gcolPlugIn(i).AdviceRevoked(lngSys, lngModual, lngPatientID, lngClinicID, lngAdviceID, int场合)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceRevokedBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int场合 As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'功能：医嘱作废前调用，一条医嘱调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strMsg   提示信息，接口内部组织返回给主程序进行提示，空串则不提示
'      strReserve   扩展参数
'返回：true允许作废，false不能作废
 
    AdviceRevokedBefore = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceRevokedBefore = gcolPlugIn(i).AdviceRevokedBefore(lngSys, lngModual, lngPatientID, lngClinicID, lngAdviceID, int场合, strMsg, strReserve)
            Err.Clear
        Next i
    End If
End Function



 
Public Function AdviceRollSend(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int场合 As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'功能：医嘱回退发送后调用，一条医嘱调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strMsg   提示信息，接口内部组织返回给主程序进行提示，空串则不提示
'      strReserve   扩展参数
 
    AdviceRollSend = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceRollSend = gcolPlugIn(i).AdviceRollSend(lngSys, lngModual, lngPatientID, lngClinicID, lngAdviceID, int场合, strMsg, strReserve)
            Err.Clear
        Next i
    End If
End Function
 
Public Function AdviceRollSendBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal int场合 As Integer, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
'功能：医嘱回退发送前调用，一条医嘱调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strMsg   提示信息，接口内部组织返回给主程序进行提示，空串则不提示
'      strReserve   扩展参数
'返回：true允许回退，false不能回退
 
    AdviceRollSendBefore = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceRollSendBefore = gcolPlugIn(i).AdviceRollSendBefore(lngSys, lngModual, lngPatientID, lngClinicID, lngAdviceID, int场合, strMsg, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function AddEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'功能：新增电子病历时候调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(门诊病历 1250,住院病历 1251,护理病历 1255)
'      lngPatientID=病人ID,lngClinicID=就诊ID/
'      EMRFileID=病历文件列表.id
'      arrInput是数组：扩展参数用于扩展使用，通过下标访一直接问时可能会报两种错，9－下标越界/13－类型不匹配：均表示参数没有传入
'           所以访问参数前应先判断：If TypeName(arrInput(i)) <> "Error" Then 满足这个条件后才使用这个参数，否则就认为这个参数没有传入
'           arrInput 目前未传入任何参数
    AddEMRBefore = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    Dim arrPars() As Variant
    arrPars = arrInput
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AddEMRBefore = gcolPlugIn(i).AddEMRBefore(lngSys, lngModual, lngPatientID, lngClinicID, EMRFileID, arrPars)
            Err.Clear
        Next i
    End If
End Function

Public Function BeforeEMRDocSave(ByVal lngSource As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal strTitle As String, ByVal strContent As String, ParamArray arrInput() As Variant) As Boolean
'功能：电子病历保存前调用
'参数：lngSource 1 门诊 2 住院
'      lngPatientID=病人ID,lngPageId=就诊ID/
'      lngFileId=病历文件列表.id
'      arrInput是数组：扩展参数用于扩展使用，通过下标访一直接问时可能会报两种错，9－下标越界/13－类型不匹配：均表示参数没有传入
'           所以访问参数前应先判断：If TypeName(arrInput(i)) <> "Error" Then 满足这个条件后才使用这个参数，否则就认为这个参数没有传入
'           arrInput 目前未传入任何参数
    BeforeEMRDocSave = True: Exit Function
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    Dim arrPars() As Variant
    arrPars = arrInput
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            BeforeEMRDocSave = gcolPlugIn(i).BeforeEMRDocSave(lngSource, lngPatiID, lngPageID, lngFileId, strTitle, strContent, arrPars)
            Err.Clear
        Next i
    End If
End Function

Public Function SignEMRBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant) As Boolean
'功能：电子病历签名前调用，事务未开启时
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号，lngModual固定传递1070
'      lngPatientID=病人ID,lngClinicID=就诊ID/主页ID
'      EMRFileID=电子病历记录.id
'      arrInput是数组：扩展参数用于扩展使用，通过下标访一直接问时可能会报两种错，9－下标越界/13－类型不匹配：均表示参数没有传入
'           所以访问参数前应先判断：If TypeName(arrInput(i)) <> "Error" Then 满足这个条件后才使用这个参数，否则就认为这个参数没有传入
'           arrInput 目前未传入任何参数
'返回：true时，允许继续签名，否则不弹出签名窗口，病历内部不做提示。
    SignEMRBefore = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    Dim arrPars() As Variant
    arrPars = arrInput
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            SignEMRBefore = gcolPlugIn(i).SignEMRBefore(lngSys, lngModual, lngPatientID, lngClinicID, EMRFileID, arrPars)
            Err.Clear
        Next i
    End If
End Function

Public Sub SignEMRAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal EMRFileID As Long, ParamArray arrInput() As Variant)
'功能：在签名窗口关闭且文档数据保存后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号，lngModual固定传递1070
'      lngPatientID=病人ID,lngClinicID=就诊ID/主页ID
'      EMRFileID=电子病历记录.id  在新增编辑未保存前其值=0，保存后签名、修改时签名、审签有值；
'      arrInput是数组：扩展参数用于扩展使用，通过下标访一直接问时可能会报两种错，9－下标越界/13－类型不匹配：均表示参数没有传入
'           所以访问参数前应先判断：If TypeName(arrInput(i)) <> "Error" Then 满足这个条件后才使用这个参数，否则就认为这个参数没有传入
'           arrInput(0)－String：签名人的姓名


'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    Dim arrPars() As Variant
    arrPars = arrInput
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            Call gcolPlugIn(i).SignEMRAfter(lngSys, lngModual, lngPatientID, lngClinicID, EMRFileID, arrPars)
            Err.Clear
        Next i
    End If
End Sub

Public Function PathImportBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径ID As Long, Optional ByVal lng版本号ID As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng疾病ID As Long, Optional ByVal lng诊断ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'功能：路径导入前
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lng路径ID=临床路径目录.id
'      lng版本号ID=导入的版本号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      lngDiagnosisType 导入诊断类型 '诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
'      lng疾病ID=导入疾病（疾病编码目录.ID）、lng诊断ID=导入诊断(疾病诊断目录.ID)
'      strReserve=扩展参数

    PathImportBefore = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PathImportBefore = gcolPlugIn(i).PathImportBefore(lngSys, lngModual, lngPatientID, lngClinicID, lng路径ID, lng版本号ID, int场合, lngDiagnosisType, _
                    lng疾病ID, lng诊断ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PathImportAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径ID As Long, Optional ByVal lng版本号ID As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal lngDiagnosisType As Long, Optional ByVal lng疾病ID As Long, Optional ByVal lng诊断ID As Long, _
    Optional ByVal strReserve As String) As Boolean
'功能：路径导入前
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lng路径ID=临床路径目录.id
'      lng版本号ID=导入的版本号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=扩展参数

    PathImportAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PathImportAfter = gcolPlugIn(i).PathImportAfter(lngSys, lngModual, lngPatientID, lngClinicID, lng路径ID, lng版本号ID, int场合, lngDiagnosisType, _
                    lng疾病ID, lng诊断ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PathEvaluateBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long, Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String) As Boolean
'功能：评估路径评估前
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号1256=临床路径应用
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lng路径记录ID=病人临床路径.id
'      lng阶段ID=当前评估的阶段ID
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=扩展参数

   PathEvaluateBefore = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PathEvaluateBefore = gcolPlugIn(i).PathEvaluateBefore(lngSys, lngModual, lngPatientID, lngClinicID, lng路径记录ID, lng阶段ID, int场合, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PathEvaluateAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long, Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String) As Boolean
'功能：评估路径评估后
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号1256=临床路径应用
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lng路径记录ID=病人临床路径.id
'      lng阶段ID=当前评估的阶段ID
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=扩展参数

   PathEvaluateAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PathEvaluateAfter = gcolPlugIn(i).PathEvaluateAfter(lngSys, lngModual, lngPatientID, lngClinicID, lng路径记录ID, lng阶段ID, int场合, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DiagnosisEnterIn(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngDiagID As Long, ByVal strDiagDesc As String, Optional lng原诊断id As Long, Optional ByVal int场合 As Integer = -1, _
    Optional ByVal int诊断类型 As Integer, Optional ByVal int诊断次序 As Integer, Optional ByVal strReserve As String) As Boolean
'功能：住院病案中输入诊断后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatientID=病人ID,lngClinicID=主页ID
'      lngDiagID=诊断ID或疾病ID(根据系统参数决定),strDiagDesc=诊断描述
'      lng原诊断id=0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id的值就是修改前的诊断ID或疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      int诊断类型=1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
'      int诊断次序=诊断类型下的顺序号
'      strReserve=扩展参数
      
    DiagnosisEnterIn = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DiagnosisEnterIn = gcolPlugIn(i).DiagnosisEnterIn(lngSys, lngModual, lngPatientID, lngClinicID, _
                lngDiagID, strDiagDesc, lng原诊断id, int场合, int诊断类型, int诊断次序, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function GetFormHwnd(Optional ByVal strReserve As String) As Long
'功能：病人信息管理 模块中调用 获取要嵌入的窗体句柄，
'参数：
'      strReserve=扩展参数
    GetFormHwnd = 0
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            GetFormHwnd = gcolPlugIn(i).GetFormHwnd(strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PatiInfoSaveBefore(ByVal lng病人ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：病人信息管理 模块中调用 病人信息管理中保存病人信息前，用于控制一些检查类的业务，返回False终止保存操作。
'参数：lng病人ID   病人ID
'      strReserve=扩展参数
    PatiInfoSaveBefore = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PatiInfoSaveBefore = gcolPlugIn(i).PatiInfoSaveBefore(lng病人ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PatiInfoSaveAfter(ByVal lng病人ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：病人信息管理 模块中调用 病人信息管理中保存病人信息后，事务提交后调用，插件中渠道自行处理数据保存
'参数：lng病人ID   病人ID
'      strReserve=扩展参数
    PatiInfoSaveAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PatiInfoSaveAfter = gcolPlugIn(i).PatiInfoSaveAfter(lng病人ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

'--业务评价接口
Public Function PatiPrePayAfter(ByVal lng病人ID As Long, ByVal int预交类型 As Integer, ByVal lng预交ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 缴纳预交款后调用（门诊/住院通用，由参数 int预交类型 来区分）
'参数：lng病人ID   病人ID
'      int预交类型：0-门诊，1-住院
'      lng预交ID  预交记录ID
'      strReserve=扩展参数
    PatiPrePayAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PatiPrePayAfter = gcolPlugIn(i).PatiPrePayAfter(lng病人ID, int预交类型, lng预交ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function OutPatiRegisterAfter(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 门诊挂号完成后调用
'参数：lng病人ID   病人ID；lng挂号ID 病人挂号ID
'      strReserve=扩展参数
    OutPatiRegisterAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            OutPatiRegisterAfter = gcolPlugIn(i).OutPatiRegisterAfter(lng病人ID, lng挂号ID, strReserve)
            Err.Clear
        Next i
    End If
End Function
 
Public Function OutPatiExseAfter(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 门诊收费完成后调用
'参数：lng病人ID   病人ID；lng结帐ID 本次收费后产生病人结帐记录ID
'      strReserve=扩展参数
    OutPatiExseAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            OutPatiExseAfter = gcolPlugIn(i).OutPatiExseAfter(lng病人ID, lng结帐ID, strReserve)
            Err.Clear
        Next i
    End If
End Function
  
Public Function OutPatiInvoicePrintAfter(ByVal lng病人ID As Long, ByVal str结帐IDs As String, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 门诊补打发票后调用
'参数：lng病人ID   病人ID；str结帐IDs 本次补打的病人结帐记录ID串，用逗号分割。
'      strReserve=扩展参数
    OutPatiInvoicePrintAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            OutPatiInvoicePrintAfter = gcolPlugIn(i).OutPatiInvoicePrintAfter(lng病人ID, str结帐IDs, strReserve)
            Err.Clear
        Next i
    End If
End Function
 
Public Function OutPatiMedicineAfter(ByVal lng病人ID As Long, ByVal strNO As String, ByVal int单据 As Integer, ByVal lng药房ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 门诊发药后调用
'参数：lng病人ID   病人ID；strNO 单据号；int单据 单据类型；lng药房ID 药房ID
'      strReserve=扩展参数
    OutPatiMedicineAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            OutPatiMedicineAfter = gcolPlugIn(i).OutPatiMedicineAfter(lng病人ID, strNO, int单据, lng药房ID, strReserve)
            Err.Clear
        Next i
    End If
End Function
 
Public Function DrugStuffWorkNoramal(ByVal lngModual As Long, ByVal strFuncName As String, ByVal lng库房ID As Long, ByVal strNO As String, ByVal int单据 As Integer, Optional ByVal strReserve As String) As Boolean
'功能：药品卫材流通管理业务调用
'参数：lngModual  模块号；
'      strFuncName  功能名称；
'      lng库房ID  库房ID；
'      strNO 单据号；
'      int单据 单据类型；
'      strReserve=扩展参数
'例：卫材外购入库调用接口，传入库房ID，NO号，单据类型
'    计划管理，单据类型=0
    Select Case lngModual
        Case m卫材外购入库管理
            '卫材外购入库
            MsgBox strFuncName & "卫材外购管理界面调用PlugIn执行成功！"
        Case m药品外购入库管理
            MsgBox strFuncName & "药品外购管理界面调用PlugIn执行成功！"
        Case m药品计划管理
            MsgBox strFuncName & "药品计划管理界面调用PlugIn执行成功！"
    End Select
   
    DrugStuffWorkNoramal = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugStuffWorkNoramal = gcolPlugIn(i).DrugStuffWorkNoramal(lngModual, strFuncName, lng库房ID, strNO, int单据, strReserve)
            Err.Clear
        Next i
    End If
End Function
 
Public Function DrugSendWorkNormal(ByVal lngModual As Long, ByVal strFuncName As String, ByVal lng库房ID As Long, ByVal strNO As String, ByVal int单据 As Integer, Optional ByVal strReserve As String) As Boolean
'功能：药品发药业务调用
'参数：lngModual  模块号；
'      strFuncName  功能名称；
'      lng库房ID  库房ID；
'      strNO 单据号；
'      int单据 单据类型；
'      strReserve=扩展参数
'例：药品业务调用接口，传入库房ID，NO号，单据类型
    Select Case lngModual
        Case m药品处方发药
            MsgBox strFuncName & "发药界面调用PlugIn执行成功！"
        Case m药品部门发药
            MsgBox strFuncName & "发药界面调用PlugIn执行成功！"
    End Select
   
    DrugSendWorkNormal = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugSendWorkNormal = gcolPlugIn(i).DrugSendWorkNormal(lngModual, strFuncName, lng库房ID, strNO, int单据, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DrugSendByRecipe(ByVal lng库房ID As Long, ByVal strNos As String, ByVal dteSendTime As Date, Optional ByVal strReserve As String) As Boolean
'功能：药品处方发药后调用，按处方发药
'参数：lng库房ID  库房ID；
'      strNOs 一次发药可能多个处方，处方类型和处方号，格式；单据,NO|单据,NO...
'      dteSendTime 发药时间，可能一个处方退药了之后再发，加上时间用于区别是哪次发药；
'      strReserve=扩展参数
       
    MsgBox "处方发药后调用PlugIn中DrugSendByRecipe接口执行成功！发药信息如下..." & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "发药处方=" & strNos & vbCrLf & _
        "发药时间=" & Format(dteSendTime, "yyyy-mm-dd hh:mm:ss")
   
    DrugSendByRecipe = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugSendByRecipe = gcolPlugIn(i).DrugSendByRecipe(lng库房ID, strNos, dteSendTime, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DrugDosageByRecipe(ByVal lng库房ID As Long, ByVal strNos As String, ByVal dteSendTime As Date, Optional ByVal strReserve As String) As Boolean
'功能：药品处方配药后调用，按处方配药
'参数：lng库房ID  库房ID；
'      strNOs 一次配药可能多个处方，处方类型和处方号，格式；单据,NO|单据,NO...
'      dteSendTime 配药时间，加上时间用于区别是哪次配药；
'      strReserve=扩展参数
       
    MsgBox "处方配药后调用PlugIn中DrugDosageByRecipe接口执行成功！配药信息如下..." & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "配药处方=" & strNos & vbCrLf & _
        "配药时间=" & Format(dteSendTime, "yyyy-mm-dd hh:mm:ss")
   
    DrugDosageByRecipe = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugDosageByRecipe = gcolPlugIn(i).DrugDosageByRecipe(lng库房ID, strNos, dteSendTime, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DrugSendFilterAfter(ByVal lng库房ID As Long, ByVal intOper As Integer, ByVal strNos As String, Optional ByVal strReserve As String) As Boolean
'功能：药品处方配药后调用，按处方配药
'参数：lng库房ID  库房ID；
'      intOper 操作状态  0-配药确认; 1-待配药; 3-待发药; 4-超时未发; 5-退药
'      strNOs 一次过滤可能多个处方，处方类型和处方号，格式；单据,NO|单据,NO...
'      strReserve=扩展参数
       
       Dim strOper As String
       
       Select Case intOper
            Case 0
                strOper = "配药确认"
            Case 1
                strOper = "待配药"
            Case 3
                strOper = "待发药"
            Case 4
                strOper = "超时未发"
            Case 5
                strOper = "退药"
        End Select
        
    MsgBox "处方过滤后调用PlugIn中DrugSendFilterAfter接口执行成功！处方信息如下..." & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "操作状态=" & strOper & vbCrLf & _
        "处方信息=" & strNos
   
    DrugSendFilterAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugSendFilterAfter = gcolPlugIn(i).DrugSendFilterAfter(lng库房ID, intOper, strNos, strReserve)
            Err.Clear
        Next i
    End If
End Function
Public Function DrugReturnByID(ByVal lng库房ID As Long, ByVal strIDs As String, ByVal dteSendTime As Date, Optional ByVal strReserve As String) As Boolean
'功能：药品退药后调用，处方发药/部门发药业务都适用
'参数：lng库房ID  库房ID；
'      strIDs 一次退药操作可能退多个药品记录，格式；收发id,退药数量|收发id,退药数量...
'      dteSendTime 退药时间，一条收发id可能退多次，加上时间可以用于区别是哪次退药；
'      strReserve=扩展参数
       
    MsgBox "退药后调用PlugIn中DrugReturnByID接口执行成功！退药信息如下..." & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "退药记录=" & strIDs & vbCrLf & _
        "退药时间=" & Format(dteSendTime, "yyyy-mm-dd hh:mm:ss")
   
    DrugReturnByID = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugReturnByID = gcolPlugIn(i).DrugReturnByID(lng库房ID, strIDs, dteSendTime, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function DrugSendBySumID(ByVal lng库房ID As Long, ByVal strSumNO As String, Optional ByVal strReserve As String) As Boolean
'功能：药品部门发药后调用，按收发id发药
'参数：lng库房ID  库房ID；
'      strSumNO 汇总发药号
'      strReserve=扩展参数
       
    MsgBox "部门发药后调用PlugIn中DrugSendBySumID接口执行成功！发药信息如下..." & vbCrLf & _
        "库房ID=" & lng库房ID & vbCrLf & _
        "发药记录=" & strSumNO
        
    DrugSendBySumID = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            DrugSendBySumID = gcolPlugIn(i).DrugSendBySumID(lng库房ID, strSumNO, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PivaWorkNormal(ByVal lngModual As Long, ByVal strFuncName As String, ByVal lng库房ID As Long, ByVal lng配药id As Long, Optional ByVal strReserve As String) As Boolean
'功能：配置中心业务调用
'参数：lngModual  模块号；
'      strFuncName  功能名称；
'      lng库房ID  库房ID；
'      lng配药id 输液单id；
'      strReserve=扩展参数
'例：配置中心业务调用接口，传入库房ID，配药id
    Select Case lngModual
        Case m输液配置中心
            MsgBox strFuncName & "配置中心界面调用PlugIn执行成功！"
    End Select
   
    PivaWorkNormal = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PivaWorkNormal = gcolPlugIn(i).PivaWorkNormal(lngModual, strFuncName, lng库房ID, lng配药id, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function InPatiCheckInAfter(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 入院处办理入院后调用
'参数：lng病人ID   lng主页ID
'      strReserve=扩展参数：入院登记调用时代表：0-正常登记,1-预约登记,2-接收预约
    InPatiCheckInAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            InPatiCheckInAfter = gcolPlugIn(i).InPatiCheckInAfter(lng病人ID, lng主页ID, strReserve)
            Err.Clear
        Next i
    End If
End Function
 
Public Function InPatiCheckInBranchAfter(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 护士站办理入科后调用
'参数：lng病人ID   lng主页ID
'      strReserve=扩展参数
    InPatiCheckInBranchAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            InPatiCheckInBranchAfter = gcolPlugIn(i).InPatiCheckInBranchAfter(lng病人ID, lng主页ID, strReserve)
            Err.Clear
        Next i
    End If
End Function
 
Public Function InPatiCheckOutAfter(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 护士站办理出院后调用(病人入出管理)
'参数：lng病人ID   病人ID
'      strReserve=扩展参数
    InPatiCheckOutAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            InPatiCheckOutAfter = gcolPlugIn(i).InPatiCheckOutAfter(lng病人ID, lng主页ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function InPatiCashierAfter(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional ByVal strReserve As String) As Boolean
'功能：业务评价 出院结账后调用
'参数：lng病人ID   病人ID；lng结帐ID 收费后产生病人结帐记录ID
'      strReserve=扩展参数
    InPatiCashierAfter = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            InPatiCashierAfter = gcolPlugIn(i).InPatiCashierAfter(lng病人ID, lng结帐ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Sub AdviceBeforeSend(ByVal strEndDate As String, ByRef rsData As ADODB.Recordset, ByRef strAdviceIDs As String, ByRef strMsg As String)
'功能：医嘱发送前调用。
'参数：strEndDate 医嘱发送的结束时间(如果医嘱为长嘱则需要此参数) 格式：2015-11-18 18:00:00
'      rsData 从ZLHIS程序中组织传入，本次发送的所有医嘱，含以下字段。
'             病人ID
'             就诊ID    主页ID/挂号ID
'             挂号单    如果是门诊病人 病人挂号记录.NO；住院病为空。
'             医嘱ID    医嘱ID   病人医嘱记录.ID
'             相关ID    如果是主医嘱 该字段为0
'             收费细目ID
'             分解时间  医嘱执行分解的执行时间点格式：2015-11-18 12:00:00,2015-11-18 18:00:00,...
'             次数      执行数
'             单量
'             单量单位
'             总量
'             总量单位
'             场合   0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strAdviceIDs 出参  本次不允许发送的医嘱ID串，用逗号分割，如：23423,3123,64563；
'      strMsg     出参  提示信息。strMsg="" － 不提示；strMsg<>"" － ZLHIS程以 msgbox 方式弹出提示。
'                       当提示信息过多时，msgbox方式提示信息显示不完成，可以将strMsg置空，在接口内部自定义提示信息的展示方式，例如：加一个模态窗体来显示提示信息。
 
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            Call gcolPlugIn(i).AdviceBeforeSend(strEndDate, rsData, strAdviceIDs, strMsg)
            Err.Clear
        Next i
    End If
End Sub

Public Function AdviceCheckSendFee(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal cur合计 As Currency, Optional ByVal int场合 As Integer, Optional strReserve As String) As Boolean
'功能：医嘱发送前费用检查：检查通过返回true,检查不通过返回false
'参数：lngSys,lngModual = 当前调用接口的主程序系统号及模块号
'      lngPatientID = 病人ID,
'      lngClinicID = 就诊ID（挂号记录ID）/主页ID(住院调用时)
'      cur合计 = 病人本次要发送医嘱的记帐金额合计,包含记帐划价单的金额
'      int场合 = 调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve = 保留参数,用于扩展使用


    AdviceCheckSendFee = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceCheckSendFee = gcolPlugIn(i).AdviceCheckSendFee(lngSys, lngModual, lngPatientID, lngClinicID, cur合计, int场合, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PatiFeeAuditingAndCancelCheck(ByVal lngModual As Long, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal blnCancel As Boolean = False, _
    Optional strExpendXML As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人费用审核和取消审核检查
    '入参:lngModual-调用的模块号
    '     lng病人ID-病人ID
    '     lng主页ID-主页ID
    '     blnCancel-取消执行时，传入True,执行时传入False
    '     strExpendXML-保留参数,暂无，待以后扩展
    '出参:strExpendXML-保留参数,暂无，待以后扩展
    '返回:检查合法，返回True,否则返回False
    '编制:刘兴洪
    '日期:2016-09-01 17:09:06
    '调用：
    '  病人费用查询--病人审核和取消病人审核时调用(数据保存前调用)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    PatiFeeAuditingAndCancelCheck = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PatiFeeAuditingAndCancelCheck = gcolPlugIn(i).PatiFeeAuditingAndCancelCheck(lngModual, lng病人ID, lng主页ID, blnCancel, strExpendXML)
            Err.Clear
        Next i
    End If
End Function

Public Function EMPI_QueryPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal rsIn As ADODB.Recordset, ByRef rsOut As ADODB.Recordset, Optional ByRef strReserve As String) As Boolean
'功能：从EMPI平台中读取病人信息
'返回：是否查找成功，true－找到病人，false-未找到任何信息
'参数：
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      strReserve=保留参数,用于扩展使用：入院登记调用时代表：0-正常登记,1-预约登记,2-接收预约
'      rsIn HIS传入的病人信息
'      rsOut 返回的病人信息结果此记录集信息由接口内部进行赋值
'      rsInt和rsOut记录集字段相同，字段均为字符串类型，包含这些字段："病人ID,主页ID,挂号ID,门诊号,住院号,医保号,身份证号,其他证件,姓名,性别,出生日期,出生地点,国籍,民族,学历,职业,工作单位,邮箱,婚姻状况,家庭电话,联系人电话,单位电话,家庭地址,家庭地址邮编,户口地址,户口地址邮编,单位邮编,联系人地址,联系人关系,联系人姓名"
'说明：如果His没有找到病人而接口中找了病人或者两者的病人ID不一致，接口决定返回病人ID告诉His以接口的病人ID为准，则需要准确核实这个病人ID是否属于当前病人。
'       因为基本信息可能因EMPI发生变化，His不能判断信息的正确性。
    Dim rsTmp As ADODB.Recordset
    
    If 2 = 1 Then
    '示例代码
    Set rsTmp = New ADODB.Recordset
    '总述，有些字段在接口内部需要转换，如一些基础字典表：性别，学历，国籍，区域，民族，籍贯，职业，婚姻状况 等
    '日期格式也需要转换
    With rsTmp.Fields
        'HIS关键，回传时可以不写入
        .Append "病人ID", adBigInt
        .Append "主页ID", adBigInt
        .Append "挂号ID", adBigInt
        '-------------------------------
        
        .Append "门诊号", adVarChar, 18
        .Append "住院号", adVarChar, 18
        .Append "医保号", adVarChar, 30
        
        'HIS-病人信息.身份证号；EMPI－身份证－IdCard
        .Append "身份证号", adVarChar, 18
        
        'HIS-病人信息.其他证件；EMPI－其它证件号码－OtherCardNo
        .Append "其他证件", adVarChar, 20
        
        .Append "姓名", adVarChar, 100
        
        'HIS-病人信息.性别；EMPI－性别名称
        .Append "性别", adVarChar, 4
    
        .Append "出生日期", adVarChar, 20 '日期格式：YYYY-MM-DD HH:MM:SS
        
        'HIS-病人信息.出生地点；EMPI－出生地－BirthPlace
        .Append "出生地点", adVarChar, 100
        
        'HIS-病人信息.国籍；EMPI－国家名称；
        .Append "国籍", adVarChar, 30
        
        'HIS-病人信息.民族；EMPI－民族名称；
        .Append "民族", adVarChar, 20
        
        'HIS-病人信息.学历；EMPI－学历名称；
        .Append "学历", adVarChar, 10
        
        'HIS-病人信息.职业；EMPI－职业特征名称
        .Append "职业", adVarChar, 80
        
        'HIS-病人信息.工作单位；EMPI－工作单位名称－CompanyName
        .Append "工作单位", adVarChar, 100
        
        'HIS-病人信息.Email；EMPI－电子邮箱－Email
        .Append "邮箱", adVarChar, 30
        
        'HIS-病人信息.婚姻状况；EMPI－婚姻状况名称
        .Append "婚姻状况", adVarChar, 4
        
        'HIS-病人信息.家庭电话；EMPI－家庭电话号码－HomePhone
        .Append "家庭电话", adVarChar, 20
        
        'HIS-病人信息.联系人电话；EMPI－联系人电话号码－RelationPhone
        .Append "联系人电话", adVarChar, 20
        
        'HIS－病人信息.单位电话；EMPI－工作单位电话号码－WorkPhone
        .Append "单位电话", adVarChar, 20
        
        'HIS－病人信息.家庭地址；EMPI－ 家庭地址-地址  － HomeAddress
        .Append "家庭地址", adVarChar, 100
        
        'HIS－病人信息.家庭地址邮编；EMPI－ 家庭地址-邮编  － HomePostcode
        .Append "家庭地址邮编", adVarChar, 6
        
        'HIS－病人信息.户口地址；EMPI－ 户籍地址-地址 - NativeAddress
        .Append "户口地址", adVarChar, 100
        
        'HIS－病人信息.户口地址邮编；EMPI－  户籍地址-邮编  -  NativePostcode
        .Append "户口地址邮编", adVarChar, 6
        
        'HIS－病人信息.单位邮编；EMPI－ 工作单位地址-邮编 - WorkPostcode
        .Append "单位邮编", adVarChar, 6
        
        'HIS－病人信息.联系人地址；EMPI－  联系人地址-地址  -  RelationAddress
        .Append "联系人地址", adVarChar, 100
        
        'HIS－病人信息.联系人关系；EMPI－ 联系人关系名称
        .Append "联系人关系", adVarChar, 30
 
        .Append "联系人姓名", adVarChar, 64
    End With
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    End If
    
    EMPI_QueryPatiInfo = False
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            EMPI_QueryPatiInfo = gcolPlugIn(i).EMPI_QueryPatiInfo(lngSys, lngModule, rsIn, rsOut, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function EMPI_AddPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngClinicID As Long, ByRef strMsg As String, Optional ByRef strReserve As String) As Long
'功能：向EMPI平台创建新病人信息
'返回：1是成功，0是失败
'参数：strMsg 出参接口内部赋值，成功、失败说明信息
'      strReserve=保留参数,用于扩展使用
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPageID－主页ID；lngClinicID－挂号ID；住院病人lngClinicID传0，门诊病人lngPageID传0

    EMPI_AddPatiInfo = 1
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            EMPI_AddPatiInfo = gcolPlugIn(i).EMPI_AddPatiInfo(lngSys, lngModule, lngPatiID, lngPageID, lngClinicID, strMsg, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function EMPI_ModifyPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngClinicID As Long, ByRef strMsg As String, Optional ByRef strReserve As String) As Long
'功能：更新EMPI平台中的病人信息
'返回：1是成功，0是失败
'参数：strMsg 出参接口内部赋值，成功、失败说明信息
'      strReserve=保留参数,用于扩展使用
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPageID－主页ID；lngClinicID－挂号ID；住院病人lngClinicID传0，门诊病人lngPageID传0

    EMPI_ModifyPatiInfo = 1
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            EMPI_ModifyPatiInfo = gcolPlugIn(i).EMPI_ModifyPatiInfo(lngSys, lngModule, lngPatiID, lngPageID, lngClinicID, strMsg, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
'功能：检查当前病人是否是指定的特殊病人
'返回：true时允许继续操作，False时不允许操作
'参数：
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngType 操作类型：1－门诊挂号，2－住院入院，3－门诊收费，4－住院结帐，5-门诊结帐。
'      lngPatiID-病人ID: 新建档的，为0,否则传入建档病人ID
'      lngPageID-主页ID: 新建档的，为0,否则传入建档主页ID(住院传入主页ID) 特殊说明：仅 lngType=4 时才传入 lngPageID，其它均传0
'      strPatiInforXML-病人信息:针对未建档病人传入，"姓名，性别，年龄，出生日期，医保号，身份证号"，出生日期 格式:2016-11-11 12:12:12,操作类型=1时，还会传入当前挂号的医生姓名
'                      固定格式：<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH><YSXM></YSXM>
'      strReserve=保留参数,用于扩展使用
    PatiValiedCheck = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PatiValiedCheck = gcolPlugIn(i).PatiValiedCheck(lngSys, lngModule, lngType, lngPatiID, lngPageID, strPatiInforXML, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function SaveCtlInfoCheck(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal rsCtlValueInfo As ADODB.Recordset, ByRef strError As String, Optional ByRef rsBackInfo As ADODB.Recordset, Optional ByRef strReserve As String) As Boolean
'功能：检查当前病人录入的信息是否准确
'返回：检查到不合格的信息后返回错误信息,格式(信息名|控件名|Index|错误信息)
'参数：
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号。
'      lngPatiID-病人ID: 新建档的，为0,否则传入建档病人ID
'      lngPageID-主页ID: 新建档的，为0,否则传入建档主页ID(住院传入主页ID) 特殊说明：仅 lngType=4 时才传入 lngPageID，其它均传0
'      rsCtlValueInfo:界面上录入的病人信息记录集:序号,信息名,控件名,Index(控件Index),ExpState(是否扩展),信息原值,信息现值,Edit(是否可编辑),是否改变
'      记录集数据有(控件名,信息名;): "txtPatient", "病人ID"; "txt住院号", "住院号"; "txt姓名", "姓名"; "txt医保号", "医保号"; "txt身份证号", "身份证号"; "cboIDNumber", "身份证未录原因";"txt支付密码", "支付密码"; "txt验证密码", "验证密码";
       '"txt出生日期", "出生日期"'"txt出生时间", "出生时间"' "txt年龄", "年龄"' "cbo年龄单位", "年龄单位"' "cbo性别", "性别"; "cbo费别", "费别"; "cbo医疗付款", "医疗付款方式"; "txt险类", "险类名";"cbo国籍","国籍"; "cbo民族", "民族";
       '"cbo身份", "身份"; "txt其他证件", "其他证件";"cbo学历", "学历"; "cbo婚姻状况", "婚姻状况"; "cbo职业", "职业"; "txt家庭地址", "现住址";"txt家庭电话", "家庭电话";"txt家庭地址邮编", "家庭地址邮编"; "txt户口地址", "户口地址";
       '"txt户口地址邮编", "户口地址邮编"; "txt籍贯", "籍贯"; "txt出生地点", "出生地点"; "txt区域", "区域"; "cbo病人类型", "病人类型"; "txt工作单位", "工作单位";"txt单位电话","单位电话"; "txt单位邮编", "单位邮编";"txt单位开户行", "单位开户行";
       '"txt单位帐号", "单位帐号";"txtMobile", "手机号"; "cbo联系人关系", "联系人关系"; "txt联系人姓名", "联系人姓名"; "txt联系人电话", "联系人电话"; "txt联系人身份证号", "联系人身份证号"; "txt联系人地址", "联系人地址"; "txt担保人", "担保人";
       '"chkUnlimit", "是否不限担保额"; "txt担保额", "担保额"; "dtp担保时限", "担保时限"; "chk临时担保", "是否临时担保";"txtReason", "担保原因"; "cbo入院科室","入院科室";"cbo入院病区", "入院病区"; "cbo床位","入院床位"; "txtTimes","第几次住院";
       '"cbo护理等级", "护理等级";"cbo门诊医师", "门诊医师"; "txt入院时间", "入院时间"; "chk再入院", "是否再入院";"chk陪伴", "是否陪伴"; "cbo住院目的", "住院目的"; "cbo入院病况", "入院病况"; "cbo入院属性", "入院属性"; "chk二级院转入", "是否二级院转入";
       '"txt门诊诊断", "门诊诊断";"cbo入院方式", "入院方式"; "txt转入", "转入医疗机构";"txt中医诊断", "中医诊断"; "txt备注", "备注";"PatiAddress(1)","出生地点";"PatiAddress(2)", "籍贯";"PatiAddress(3)", "现住址";"PatiAddress(4)", "户口地址";
       '"PatiAddress(5)", "联系人地址";"txtPages", "主页ID";"txtLinkManInfo","关系",
'      rsBackInfo:扩展参数,用于记录检查出的所有错误信息，返回到界面上显示出来
'      strReserve=保留参数,用于扩展使用
    Dim str信息值 As String
'例子:通过“控件名+信息名+Index” 的方式来过滤出信息值
'    If Not rsCtlValueInfo Is Nothing Then
'        rsCtlValueInfo.MoveFirst
'        If Not rsCtlValueInfo.EOF Then
'            rsCtlValueInfo.Filter = "信息名='手机号' And 控件名='txtMobile' And Index=-1"
'            If Not rsCtlValueInfo.EOF Then
'                str信息值 = "" & rsCtlValueInfo!信息现值
'                If Not IsNumeric(str信息值) Then
'                    strError = "手机号|txtMobile|0|手机号输入格式不正确，请检查！"
'                     MsgBox "手机号输入格式不正确，请检查！", vbInformation
'                    Exit Function
'                End If
'            End If
'            '数组控件
'            rsCtlValueInfo.Filter = "信息名='出生地点' And 控件名='PatiAddress' And Index=1"
'            If Not rsCtlValueInfo.EOF Then
'                str信息值 = "" & rsCtlValueInfo!信息现值
'                If str信息值 <> "重庆市丰都县" Then
'                    strError = "出生地点|PatiAddress|1|出生地点输入错误，请检查！"
'                     MsgBox "出生地点输入错误，请检查！", vbInformation
'                    Exit Function
'                End If
'            End If
'        End If
'    End If
    SaveCtlInfoCheck = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            SaveCtlInfoCheck = gcolPlugIn(i).SaveCtlInfoCheck(lngSys, lngModule, lngPatiID, lngPageID, rsCtlValueInfo, strError, rsBackInfo, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function PatiRegTriageCheck(lngSys As Long, ByVal lngModule As Long, bytFun As Byte, ByVal lng挂号ID As Long, _
    Optional ByVal str医生 As String, Optional ByVal str诊室 As String, Optional str新号码 As String = "", Optional lng新出诊记录ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:挂号分诊检测接口
    '入参:lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '     bytFun- 0-分诊;1-换号;2-完成就诊(13-恢复就诊);3-标记为不就诊;4-签道(14-取消签道);5-回诊(15-取消回诊);6-病人待诊
    '     lng挂号ID-挂号ID
    '     str医生-医生姓名(bytfun=0,1,2,4时传入)
    '     str诊室-诊室名称(bytfun=0,1,2,4时传入)
    '     str新号码-换号时的新号码(bytfun=1时传入）
    '     lng新出诊记录ID-换号时的新号码(bytfun=1时传入）
    '出参:
    '返回:分诊合法返回true,否则返回False
    '编制:刘尔旋
    '日期:2017-03-24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    PatiRegTriageCheck = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PatiRegTriageCheck = gcolPlugIn(i).PatiRegTriageCheck(lngSys, lngModule, bytFun, lng挂号ID, str医生, str诊室, str新号码, lng新出诊记录ID)
            Err.Clear
        Next i
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckChargeItem(lngSys As Long, ByVal lngModule As Long, ByVal intType As Integer, ByVal intMode As Integer, _
    ByRef rsDetail As ADODB.Recordset, Optional strExpend As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收费项目有效性检查
    '入参:lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '     intType:0-门诊;1-住院
    '     intMode:0-录入明细时的常规检查;1-保存单据前的汇总检查
    '     rsDetail-病人ID，主页ID，收费类别，收费细目ID，数量，单价，实收金额，开单人，开单科室,
    '                  执行科室ID, 单据性质（1-收费单,2-记帐单), 是否划价(1-划价;0-正常的收费及记帐单)

    '     strExpend-待以后扩展，暂无用
    '出参:strExpend-待以后扩展，暂无用
    '返回:数据合法返回true,否则返回False
    '编制:冉俊明
    '日期:2017-04-21 10:09:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    CheckChargeItem = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            CheckChargeItem = gcolPlugIn(i).CheckChargeItem(lngSys, lngModule, intType, intMode, rsDetail, strExpend)
            Err.Clear
        Next i
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckMecInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef colCtl As Collection, ByRef strMsg As String, Optional ByRef strReserve As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:首页输入有效性检查
    '返回：True是成功，False是失败
    '参数：strMsg 返回参数 格式 例: 提示消息-(提醒:1/禁止:0)|控件Key值-提示消息-(提醒:1/禁止:0)|表格控件Key值-提示消息-(提醒:1/禁止:0)-Row-Col
    '      strReserve=保留参数,用于扩展使用 固定格式，用<sTab>分割，
    '      lngType执行 & "<sTab>" &...&  "<sTab>" & ... &  "<sTab>"
    '      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '      colCtl 病案首页控件集合
    '      lngPageID－主页ID；
    '------------------------------------------------------------colCtl控件集合Key值--------------------------------------------------------------
    
    '结构化地址控件:出生地点(结构化), 籍贯(结构化), 现住址(结构化), 户口地址(结构化), 联系人地址(结构化),单位地址(结构化)
    '结构化地址控件[属性]:value省,value市,value区县,value乡镇,value详细地址
    
    '文本框:出生地点, 籍贯, 现住址, 户口地址, 联系人地址, 病人区域, 单位地址, 病案号, 档案号, X线号, 姓名, 其他证件, 联系人姓名, 入院科室, 入院病房,
    '       出院科室, 出院病房, 医保号, 摘要, 门诊号, 监护人, 发病地址, 医学警示, 其他医学警示, 病理号, 死亡原因, 病原学诊断, 抢救病因, 输其他,
    '       转入医疗机构, 31天内再住院, 转科1, 转科2, 转科3, 退出原因, 变异原因, 重症监护室名称, 肿瘤T, 肿瘤N, 肿瘤M, Email, 其他会诊, 引发药物,
    '       临床表现, 透析尿素氮值, 其他关系, 入院转入, 单位电话, 单位邮编, 家庭电话, 家庭邮编, 户口邮编, 身高, 身高单位, 体重, 体重单位, 体温,
    '       入院次数, 收缩压, 舒张压, 联系人电话, 年龄, 婴幼儿年龄, 新生儿出生体重, 新生儿入院体重, 住院天数, 住院号, 抢救次数, 成功次数, 特护,
    '       一级护理, 二级护理, 三级护理, ICU, CCU, 输红细胞, 输血小板, 输血浆, 输全血, 自体回收, 呼吸机使用, 昏迷时间入院前_天, 昏迷时间入院前_小时,
    '       昏迷时间入院前_分钟, 昏迷时间入院后_天, 昏迷时间入院后_小时, 昏迷时间入院后_分钟, 随诊期限, 费用和, 约束总时间, 重症监护天, 重症监护小时,
    '       Apgar , QQ, 输白蛋白, 院内会诊, 外院会诊, 距上次住院时间,婴幼儿年龄_DAY,GC_监护人身份证号
    
    '下拉框: 门诊医师 , 科主任, 主任或副主任, 进修医师, 主治医师, 住院医师, 研究生医师, 实习医师, 质控医师, 质控护士, 责任护士, 编目员, 主诊医师, 随诊期限(下拉)
    '        付款方式, 性别, 婚姻, 职业, 国籍, 民族, 关系, 入院途径, 文化程度, 去向, 感染与死亡关系, 入院情况, 分化程度, 最高诊断依据, 门诊与出院XY,入院与出院XY,
    '        门诊与入院, 术前与术后, 放射与病理, 临床与病理, 死亡期间, 临床与尸检, 门诊与出院ZY, 入院与出院ZY, 辩证, 治法, 方药, 治疗类别, 中医诊疗设备, 抢救方法,
    '        中医诊疗技术, 自制中药制剂, 辨证施护, 病案质量, 病例分型, HBsAg, 血型, HCVAb, RH, HIVAb, 输液反应, 输血反应, 输血前9项检查, 生育状况, 出院方式,
    '        再入院计划天数, 压疮发生期间, 压疮分期, 跌倒或坠床伤害, 跌倒或坠床原因, 距上次住院时间(下拉), 重返间隔时间, 约束方式, 约束工具, 约束原因,
    '        新生儿离院方式 , 肿瘤分期, 临床路径管理, 法定传染病, 实施DGRS管理, 死亡患者尸检, 身份证, 变异原因(下拉), 健康卡号, 年龄(下拉), 婴幼儿年龄(下拉),联系人身份证
    
    '时间控件: 出生日期 , 入院时间, 出院时间, 确诊日期, 死亡时间, 发病日期, 发病时间, 质控日期, 编目日期, 收回日期, 入科时间
    
    '复选框: 再入院 , 入院前外院治疗, 是否确诊, 病原学检查, 新发肿瘤, 危重, 急症, 疑难, 示教病案, 科研病案, 疑难病例, 随诊, CT, MRI, 彩色多普勒,
    '        传染病上传, 围术期死亡, 术后猝死, 进入路径, 完成路径, 变异, 住院出现危重, 是否同一疾病, 人工气道脱出, 重返重症医学科, 住院物理约束, 单病种管理,
    '        细菌标本送检 , 会诊情况, 无过敏记录, 未找到时自由录入,无手术记录,无重症监护记录,无转科记录
    
    '单选按钮: 再住院无 , 再住院有, 初诊, 复诊, ICU无, ICU有, XY按诊断输入, XY按疾病编码输入, ZY按诊断输入, ZY按疾病编码输入, 按诊断输入,
    '          按疾病编码输入, 按药品目录输入, 按过敏源输入
    
    'VS表格: 转科情况 (表格), 西医诊断(表格), 中医诊断(表格), 过敏信息(表格), 手术记录(表格), 住院费用(表格), 化疗记录信息(表格), 放疗记录信息(表格),
    '        抗菌药物使用情况(表格), 抗精神病治疗情况(表格), 重症监护情况(表格), 病案附加项目(表格), 特殊检查情况(表格),器械导管使用情况(表格),病人感染记录(表格),
    '        感染部位(表格),病人病原学检查(表格)
    
    'list列表控件: 不良事件(表格), 感染因素(表格)
    
    '树形表控件:  住院费用(树形表)
    '------------------------------------------------------------表格控件列枚举--------------------------------------------------------------------
    '过敏信息列枚举 : 过敏药物=0 ,过敏反应=1 ,过敏时间=2 ,过敏源编码=3 ,药物ID=4 ,过敏来源=5
    
    '(西医诊断/中医诊断)列枚举 : 诊断类型=0 ,关联=1 ,诊断编码=2 ,诊断描述=3 ,中医证候=4 ,发病时间=5 ,备注=6 ,入院病情=7 ,出院情况=8 ,ICD附码=9 ,是否未治=10
    '            是否疑诊=11 ,增加=12 ,Del=13 ,诊断ID=14 ,疾病ID=15 ,证候ID=16 ,医嘱IDs=17  ,诊断分类=18 ,固定附码=19 ,是否病人=20
    '            诊断来源=24 ,疾病编码=25 ,疾病类别=26 ,证候编码=27 ,记录日期=28 ,记录人员=29 ,疗效限制=21 ,分娩信息=22 ,附码ID=23
    
    '手术列枚举 : Copy=0 ,手术日期=1 ,结束日期=2 ,抗菌用药时间=3 ,手术情况=4 ,准备天数=5 ,手术编码=6 ,手术名称=7 ,再次手术=8 ,主刀医师=9 ,助产护士=10 ,助手1=11
    '    助手2=12 ,麻醉开始时间=13 ,麻醉类型 (界面名称麻醉方式)= 14  ,ASA分级=15 ,NNIS分级=16 ,手术级别=17 ,麻醉医师=18 ,切口愈合=19 ,切口部位=20 ,重返手术室计划=21
    '    重返手术室目的=22 ,切口感染=23 ,并发症=24 ,预防用抗菌药=25 ,抗菌药天数=26 ,非预期的二次手术=27 ,麻醉并发症=28 ,术中异物遗留=29 ,手术并发症=30 ,术后出血或血肿=31
    '    手术伤口裂开=32 ,术后深静脉血栓=33 ,术后生理代谢紊乱=34 ,术后呼吸衰竭=35 ,术后肺栓塞=36 ,术后败血症=37 ,术后髋关节骨折=38 ,手术操作ID=39 ,诊疗项目ID=40 ,麻醉ID=41
    '    麻醉方式(界面名称麻醉类型)=42  ,手麻来源=43
    
    '化疗列枚举 : 化学治疗编码=0 ,开始日期=1 ,结束日期=2 ,疗程数=3 ,化疗方案=4 ,总量=5 ,化疗效果=6 ,疾病ID=7
    
    '放疗列枚举 : 放射治疗编码=0 ,开始日期=1 ,结束日期=2 ,设野部位=3 ,放射剂量=4 ,累计量=5 ,放疗效果=6 ,疾病ID=7
    
    '精神药品列枚举 : 药物名称=0 ,疗程=1 ,最高日量=2 ,特殊反应=3 ,疗效=4 ,药品ID=5
    
    '重症监护枚举 : 序号=0 ,监护室名称=1 ,进入时间=2 ,退出时间=3 ,再入住计划=4 ,再入住原因=5
    
    '重症监护器械枚举 : ICU类型=0 ,器械及导管=1 ,开始时间=2 ,结束时间=3 ,感染累计小时=4
    
    '医院感染枚举 : 确诊日期=0 ,感染部位=1 ,医院感染名称=2 ,医院感染编码=3
    
    '标本来源枚举 : 标本=0 ,病原学代码及名称=1 ,送检日期=2
    
    '抗菌药枚举 : 序号=0 ,抗菌药物名=1 ,用药目的=2 ,使用阶段=3 ,使用天数=4 ,一类切口预防用=5 ,DDD数=6 ,联合用药=7
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '使用itemData的控件:
    '
    '涉及动态加载下拉框: 用于存储数据的id
    '
    '不良事件 (表格), 感染因素(表格): 用于存储数据的编码
    '
    '使用rowdata的控件:
    '
    '过敏信息 (表格), 西医诊断(表格), 中医诊断(表格), 手术记录(表格): 用于存储数据的id
    '
    '化疗记录信息 (表格), 重症监护情况(表格), 放疗记录信息(表格), 抗精神病治疗情况(表格): 用于存储数据的序号
    '
    '抗菌药物使用情况 (表格): 用于存储数据的药名id
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    CheckMecInfo = True
'    Dim i As Integer
'    Dim txt As TextBox
'    On Error GoTo errHandle
'    Dim varTmp As Variant
'    Dim lngType执行 As Long
'    CheckMecInfo = True
'    strMsg = ""
'    '文本框示例
'    Set txt = colCtl.Item("家庭邮编")
'    If txt.Text = "" Then
'        strMsg = strMsg & "家庭邮编-家庭邮编不能为空,请重新输入-1|"
'    End If
'    '下拉框示例
'    If colCtl.Item("住院医师").Text = "" Then
'        strMsg = strMsg & "住院医师-住院医生不能为空,请重新输入-1|"
'    End If
'    '时间控件示例
'    If colCtl.Item("发病日期").Text = "" Then
'        strMsg = strMsg & "发病日期-发病日期不能为空,请重新输入-0|"
'    End If
'    '复选框示例
'    If colCtl.Item("再入院").Value = 0 Then
'        strMsg = strMsg & "再入院-检测到尚未勾选再入院-1|"
'    End If
'    '单选框示例
'    If colCtl.Item("再住院无").Value = 0 Then
'        strMsg = strMsg & "再住院无-检测到尚未勾选再住院无-0|"
'    End If
'    '结构化地址控件示例
'    If colCtl.Item("现住址(结构化)").value市 = "" Then
'        strMsg = strMsg & "现住址(结构化)-现住址不能为空,请重新输入-1|"
'    End If
'    '表格控件示例
'    For i = 0 To colCtl.Item("西医诊断(表格)").Rows - 1
'        If colCtl.Item("西医诊断(表格)").TextMatrix(i, 3) = "" Then
'            strMsg = strMsg & "西医诊断(表格)-西医诊断的诊断描述不能为空,请重新输入-1-" & i & "-" & 3 & "|"
'        End If
'    Next
'    '返回参数格式示例
'    strMsg = strMsg & "不绑定控件的禁止信息-0|门诊医师-绑定控件的提示信息-1|西医诊断(表格)-绑定表格控件的提示信息-1-2-2"
'    If strMsg <> "" Then
'        CheckMecInfo = False
'    End If
'
'    If strReserve<>"" Then
'       varTmp = Split(strReserve, "<sTab>")
'        lngType执行= varTmp(0) 'lngType执行:设置 = 0,预览 = 1,打印 = 2,确定 = 3,确定并退出 = 4
'    End IF
'    If lngType执行=2 Then
'        If colCtl.Item("住院医师").Text = "" Then
'            strMsg = strMsg & "住院医师-住院医生不能为空,请重新输入-1|"
'        End If
'    End IF
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            CheckMecInfo = gcolPlugIn(i).CheckMecInfo(lngSys, lngModule, lngPatiID, lngPageID, colCtl, strMsg, strReserve)
            Err.Clear
        Next i
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function LoadMecInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngPatiType As Long, ByRef colCtl As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:首页加载自定义数据
    '返回：True是成功，False是失败
    '      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '      colCtl 病案首页控件集合
    '      lngPageID－主页ID；
    '      lngPatiType－病人类型:1-门诊,2-住院
    '      lngPatiID-病人id
    '------------------------------------------------------------colCtl控件集合Key值--------------------------------------------------------------
    
    '结构化地址控件:出生地点(结构化), 籍贯(结构化), 现住址(结构化), 户口地址(结构化), 联系人地址(结构化)
    '结构化地址控件[属性]:value省,value市,value区县,value乡镇,value详细地址
    
    '文本框:出生地点, 籍贯, 现住址, 户口地址, 联系人地址, 病人区域, 单位地址, 病案号, 档案号, X线号, 姓名, 其他证件, 联系人姓名, 入院科室, 入院病房,
    '       出院科室, 出院病房, 医保号, 摘要, 门诊号, 监护人, 发病地址, 医学警示, 其他医学警示, 病理号, 死亡原因, 病原学诊断, 抢救病因, 输其他,
    '       转入医疗机构, 31天内再住院, 转科1, 转科2, 转科3, 退出原因, 变异原因, 重症监护室名称, 肿瘤T, 肿瘤N, 肿瘤M, Email, 其他会诊, 引发药物,
    '       临床表现, 透析尿素氮值, 其他关系, 入院转入, 单位电话, 单位邮编, 家庭电话, 家庭邮编, 户口邮编, 身高, 身高单位, 体重, 体重单位, 体温,
    '       入院次数, 收缩压, 舒张压, 联系人电话, 年龄, 婴幼儿年龄, 新生儿出生体重, 新生儿入院体重, 住院天数, 住院号, 抢救次数, 成功次数, 特护,
    '       一级护理, 二级护理, 三级护理, ICU, CCU, 输红细胞, 输血小板, 输血浆, 输全血, 自体回收, 呼吸机使用, 昏迷时间入院前_天, 昏迷时间入院前_小时,
    '       昏迷时间入院前_分钟, 昏迷时间入院后_天, 昏迷时间入院后_小时, 昏迷时间入院后_分钟, 随诊期限, 费用和, 约束总时间, 重症监护天, 重症监护小时,
    '       Apgar , QQ, 输白蛋白, 院内会诊, 外院会诊, 距上次住院时间
    
    '下拉框: 门诊医师 , 科主任, 主任或副主任, 进修医师, 主治医师, 住院医师, 研究生医师, 实习医师, 质控医师, 质控护士, 责任护士, 编目员, 主诊医师, 随诊期限(下拉)
    '        付款方式, 性别, 婚姻, 职业, 国籍, 民族, 关系, 入院途径, 文化程度, 去向, 感染与死亡关系, 入院情况, 分化程度, 最高诊断依据, 门诊与出院XY,入院与出院XY,
    '        门诊与入院, 术前与术后, 放射与病理, 临床与病理, 死亡期间, 临床与尸检, 门诊与出院ZY, 入院与出院ZY, 辩证, 治法, 方药, 治疗类别, 中医诊疗设备, 抢救方法,
    '        中医诊疗技术, 自制中药制剂, 辨证施护, 病案质量, 病例分型, HBsAg, 血型, HCVAb, RH, HIVAb, 输液反应, 输血反应, 输血前9项检查, 生育状况, 出院方式,
    '        再入院计划天数, 压疮发生期间, 压疮分期, 跌倒或坠床伤害, 跌倒或坠床原因, 距上次住院时间(下拉), 重返间隔时间, 约束方式, 约束工具, 约束原因,
    '        新生儿离院方式 , 肿瘤分期, 临床路径管理, 法定传染病, 实施DGRS管理, 死亡患者尸检, 身份证, 变异原因(下拉), 健康卡号, 年龄(下拉), 婴幼儿年龄(下拉),
    
    '时间控件: 出生日期 , 入院时间, 出院时间, 确诊日期, 死亡时间, 发病日期, 发病时间, 质控日期, 编目日期, 收回日期, 入科时间
    
    '复选框: 再入院 , 入院前外院治疗, 是否确诊, 病原学检查, 新发肿瘤, 危重, 急症, 疑难, 示教病案, 科研病案, 疑难病例, 随诊, CT, MRI, 彩色多普勒,
    '        传染病上传, 围术期死亡, 术后猝死, 进入路径, 完成路径, 变异, 住院出现危重, 是否同一疾病, 人工气道脱出, 重返重症医学科, 住院物理约束, 单病种管理,
    '        细菌标本送检 , 会诊情况, 无过敏记录, 未找到时自由录入
    
    '单选按钮: 再住院无 , 再住院有, 初诊, 复诊, ICU无, ICU有, XY按诊断输入, XY按疾病编码输入, ZY按诊断输入, ZY按疾病编码输入, 按诊断输入,
    '          按疾病编码输入, 按药品目录输入, 按过敏源输入
    
    'VS表格: 转科情况 (表格), 西医诊断(表格), 中医诊断(表格), 过敏信息(表格), 手术记录(表格), 住院费用(表格), 化疗记录信息(表格), 放疗记录信息(表格),
    '        抗菌药物使用情况(表格), 抗精神病治疗情况(表格), 重症监护情况(表格), 病案附加项目(表格), 特殊检查情况(表格)
    
    'list列表控件: 不良事件(表格), 感染因素(表格)
    
    '树形表控件:  住院费用(树形表)
    '------------------------------------------------------------表格控件列枚举--------------------------------------------------------------------
    '过敏信息列枚举 : 过敏药物=0 ,过敏反应=1 ,过敏时间=2 ,过敏源编码=3 ,药物ID=4 ,过敏来源=5
    
    '(西医诊断/中医诊断)列枚举 : 诊断类型=0 ,关联=1 ,诊断编码=2 ,诊断描述=3 ,中医证候=4 ,发病时间=5 ,备注=6 ,入院病情=7 ,出院情况=8 ,ICD附码=9 ,是否未治=10
    '            是否疑诊=11 ,增加=12 ,Del=13 ,诊断ID=14 ,疾病ID=15 ,证候ID=16 ,医嘱IDs=17  ,诊断分类=18 ,固定附码=19 ,是否病人=20
    '            诊断来源=24 ,疾病编码=25 ,疾病类别=26 ,证候编码=27 ,记录日期=28 ,记录人员=29 ,疗效限制=21 ,分娩信息=22 ,附码ID=23
    
    '手术列枚举 : Copy=0 ,手术日期=1 ,结束日期=2 ,抗菌用药时间=3 ,手术情况=4 ,准备天数=5 ,手术编码=6 ,手术名称=7 ,再次手术=8 ,主刀医师=9 ,助产护士=10 ,助手1=11
    '    助手2=12 ,麻醉开始时间=13 ,麻醉类型 (界面名称麻醉方式)= 14  ,ASA分级=15 ,NNIS分级=16 ,手术级别=17 ,麻醉医师=18 ,切口愈合=19 ,切口部位=20 ,重返手术室计划=21
    '    重返手术室目的=22 ,切口感染=23 ,并发症=24 ,预防用抗菌药=25 ,抗菌药天数=26 ,非预期的二次手术=27 ,麻醉并发症=28 ,术中异物遗留=29 ,手术并发症=30 ,术后出血或血肿=31
    '    手术伤口裂开=32 ,术后深静脉血栓=33 ,术后生理代谢紊乱=34 ,术后呼吸衰竭=35 ,术后肺栓塞=36 ,术后败血症=37 ,术后髋关节骨折=38 ,手术操作ID=39 ,诊疗项目ID=40 ,麻醉ID=41
    '    麻醉方式(界面名称麻醉类型)=42  ,手麻来源=43
    
    '化疗列枚举 : 化学治疗编码=0 ,开始日期=1 ,结束日期=2 ,疗程数=3 ,化疗方案=4 ,总量=5 ,化疗效果=6 ,疾病ID=7
    
    '放疗列枚举 : 放射治疗编码=0 ,开始日期=1 ,结束日期=2 ,设野部位=3 ,放射剂量=4 ,累计量=5 ,放疗效果=6 ,疾病ID=7
    
    '精神药品列枚举 : 药物名称=0 ,疗程=1 ,最高日量=2 ,特殊反应=3 ,疗效=4 ,药品ID=5
    
    '重症监护枚举 : 序号=0 ,监护室名称=1 ,进入时间=2 ,退出时间=3 ,再入住计划=4 ,再入住原因=5
    
    '重症监护器械枚举 : ICU类型=0 ,器械及导管=1 ,开始时间=2 ,结束时间=3 ,感染累计小时=4
    
    '医院感染枚举 : 确诊日期=0 ,感染部位=1 ,医院感染名称=2 ,医院感染编码=3
    
    '标本来源枚举 : 标本=0 ,病原学代码及名称=1 ,送检日期=2
    
    '抗菌药枚举 : 序号=0 ,抗菌药物名=1 ,用药目的=2 ,使用阶段=3 ,使用天数=4 ,一类切口预防用=5 ,DDD数=6 ,联合用药=7

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '使用itemData的控件:
    '
    '涉及动态加载下拉框: 用于存储数据的id
    '
    '不良事件 (表格), 感染因素(表格): 用于存储数据的编码
    '
    '使用rowdata的控件:
    '
    '过敏信息 (表格), 西医诊断(表格), 中医诊断(表格), 手术记录(表格): 用于存储数据的id
    '
    '化疗记录信息 (表格), 重症监护情况(表格), 放疗记录信息(表格), 抗精神病治疗情况(表格): 用于存储数据的序号
    '
    '抗菌药物使用情况 (表格): 用于存储数据的药名id
    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim CBO As ComboBox
'   On Error GoTo errHandle

'    例:自定义家庭邮编
'    colCtl.Item("家庭邮编").Text = "123456"

'    '例:自定义下拉框内容
'    Set CBO = colCtl.Item("入院情况")
'    CBO.AddItem ""
'    CBO.ListIndex = CBO.ListCount - 1
   On Error GoTo errHandle
    LoadMecInfo = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            LoadMecInfo = gcolPlugIn(i).LoadMecInfo(lngSys, lngModule, lngPatiID, lngPageID, lngPatiType, colCtl)
            Err.Clear
        Next i
    End If
    Exit Function
errHandle:
    MsgBox Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function AdviceMakeFee(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strCurFee As String, ByVal rsNowFee As ADODB.Recordset, ByRef dblCount As Double, Optional ByRef strReserve As String) As Boolean
'功能：获取检查医嘱子医嘱行对应的费用收取次数，具体收几次费由接口内部处理，医嘱发送时调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      strCurFee当前医嘱记录行对的相关信息，固定格式，用<sTab>分割，
'              lng医嘱ID & "<sTab>" & lng当前医嘱ID & "<sTab>" & lng诊疗项目ID & "<sTab>" & lng收费项目ID & "<sTab>" & int收费方式 & "<sTab>" & str部位 & "<sTab>" & str方法 & "<sTab>" & (dbl总量 * dbl计价数量)
'      rsNowFee 本次发送中已经处理的费用项目 包含的字段
            '"医嘱ID", adBigInt '一组医嘱的ID
            '"诊疗项目ID", adBigInt
            '"收费项目ID", adBigInt
            '"收费方式", adInteger
            '"收费时间", adVarChar, 10
            '"执行部门ID", adBigInt
            '"子医嘱ID", adBigInt '相关ID不为空的医嘱行的医嘱ID
            '"检查部位", adVarChar, 30
            '"检查方法", adVarChar, 30
            '"数量", adDouble '收费数量
'      dblCount 出参数取值-1，0，>0；-1表示按主程序原有逻辑收取相当于是正常收费、为0表示本次不用收取、>0表示本次收费次数
'      strReserve 保留扩展参数
'特殊说明，rsNowFee 记录集只能读取不能修改删除
    Dim lng医嘱ID As Long '主医嘱ID
    Dim lng当前医嘱ID As Long '子医嘱ID
    Dim lng诊疗项目ID As Long
    Dim lng收费项目ID As Long
    Dim int收费方式 As Integer
    Dim str部位 As String
    Dim str方法 As String
    Dim dbl总量 As Double
    Dim i As Long
    Dim varTmp As Variant
    Dim dblTmp总量 As Double
    
    dblCount = -1
    
    If 0 = 1 Then
        '示例代码功能描述，一组检查医嘱中的子医嘱中相同收费项目中只收费费用次最多的一个费用对照
        varTmp = Split(strCurFee, "<sTab>")
        lng医嘱ID = varTmp(0)
        lng当前医嘱ID = varTmp(1)
        lng诊疗项目ID = varTmp(2)
        lng收费项目ID = varTmp(3)
        int收费方式 = varTmp(4)
        str部位 = varTmp(5)
        str方法 = varTmp(6)
        dbl总量 = varTmp(7)
       
        '当前这组检查医嘱中排除
        rsNowFee.Filter = "医嘱ID=" & lng医嘱ID & " and 收费项目ID=" & lng收费项目ID & " and 收费方式=" & int收费方式
        If Not rsNowFee.EOF Then
            For i = 1 To rsNowFee.RecordCount
                dblTmp总量 = dblTmp总量 + Val(rsNowFee!数量 & "")
                rsNowFee.MoveNext
            Next
        End If
        dbl总量 = dbl总量 - dblTmp总量
        If dbl总量 <= 0 Then
            dblCount = 0
        Else
            dblCount = dbl总量
        End If
    End If
    
    AdviceMakeFee = True
    
    
    '----------扩展部件接口遍历调用------------------------
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceMakeFee = gcolPlugIn(i).AdviceMakeFee(lngSys, lngModual, strCurFee, rsNowFee, dblCount)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceLoadApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng就诊ID As String, ByVal lng文件ID As Long, ByRef strXML As String, Optional ByVal lng医嘱ID As Long, Optional ByRef strReserve As String) As Boolean
'功能：门诊住院自定义申请单读取数据处理(XML处理)
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatiID=病人ID,lng就诊ID=门诊挂号单号,住院主页ID
'      lng文件ID=申请单对应的病历文件列表.ID；lng医嘱ID 修改申请单时传入
'      strXML=自定义申请单的XML数据(出参)
'      lng医嘱ID=修改和查看时，医嘱ID不为空，新开时为空
'      strReserve=扩展参数

'示例-----------------------------------------------------------------------------
'    Dim objXML As Object
'
'    Set objXML = CreateObject("MSXML2.DOMDocument")
'    If objXML Is Nothing Then
'        MsgBox "创建MSXML2.DOMDocument对象失败", vbExclamation, Me.Caption
'        Exit Function
'    End If
'    Call objXML.loadXML(strXML)
'    On Error Resume Next
'    objXML.selectSingleNode(".//" & "姓名").Text = mstr病人姓名
'---------------------------------------------------------------------------------
    
    AdviceLoadApplyCustom = True
    '----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceLoadApplyCustom = gcolPlugIn(i).AdviceLoadApplyCustom(lngSys, lngModual, lngPatiID, lng就诊ID, lng文件ID, strXML, lng医嘱ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceSaveApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng就诊ID As String, ByVal lng文件ID As Long, ByRef strXML As String, ByRef objWeb As Object, Optional ByVal lng医嘱ID As Long, Optional ByRef strReserve As String) As Boolean
'功能：门诊住院自定义申请单保存数据处理(XML处理)
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatiID=病人ID,lng就诊ID=门诊挂号单号,住院主页ID
'      lng文件ID=申请单对应的病历文件列表.ID；lng医嘱ID 修改申请单时传入
'      strXML=自定义申请单的XML数据(出参)
'      objWeb=WebBrowser浏览器控件，可通过objWeb.document 对象访问网页表单数据进行检查
'      strReserve=扩展参数

'示例-----------------------------------------------------------------------------
'    Dim objXML As Object
'
'    Set objXML = CreateObject("MSXML2.DOMDocument")
'    If objXML Is Nothing Then
'        MsgBox "创建MSXML2.DOMDocument对象失败", vbExclamation, Me.Caption
'        Exit Function
'    End If
'    Call objXML.loadXML(strXML)
'    On Error Resume Next
'    objXML.selectSingleNode(".//" & "姓名").Text = mstr病人姓名

'    Set vDoc = webSub.Document
'    For i = 0 To vDoc.All.Length - 1
'        If UCase(vDoc.All(i).tagName) = "INPUT" Or UCase(vDoc.All(i).tagName) = "SELECT" Then
'            Set vtag = vDoc.All(i)
'            If vtag.Type = "text" Or vtag.Type = "password" Or vtag.Type = "hidden" Or vtag.Type = "select-one" Then
'                If vtag.Value = XXXX Then XXXXX
'            End If
'        End If
'    Next
'---------------------------------------------------------------------------------
    
    AdviceSaveApplyCustom = True
    '----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceSaveApplyCustom = gcolPlugIn(i).AdviceSaveApplyCustom(lngSys, lngModual, lngPatiID, lng就诊ID, lng文件ID, strXML, objWeb, lng医嘱ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function AdviceCheckApplyCustom(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatiID As Long, ByVal lng就诊ID As String, ByVal lng文件ID As Long, ByVal strXML As String, ByRef objWeb As Object, Optional ByVal lng医嘱ID As Long, Optional ByRef strReserve As String) As Boolean
'功能：门诊住院自定义申请单检查表单数据
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lngPatiID=病人ID,lng就诊ID=门诊挂号单号,住院主页ID
'      lng文件ID=申请单对应的病历文件列表.ID；lng医嘱ID 修改申请单时传入
'      strXML=自定义申请单的XML数据
'      objWeb=WebBrowser浏览器控件，可通过objWeb.document 对象访问网页表单数据进行检查
'      strReserve=扩展参数

'示例-----------------------------------------------------------------------------
'    Dim objXML As Object
'
'    Set objXML = CreateObject("MSXML2.DOMDocument")
'    If objXML Is Nothing Then
'        MsgBox "创建MSXML2.DOMDocument对象失败", vbExclamation, Me.Caption
'        Exit Function
'    End If
'    Call objXML.loadXML(strXML)
'    On Error Resume Next
'    objXML.selectSingleNode(".//" & "姓名").Text = mstr病人姓名

'    Set vDoc = webSub.Document
'    For i = 0 To vDoc.All.Length - 1
'        If UCase(vDoc.All(i).tagName) = "INPUT" Or UCase(vDoc.All(i).tagName) = "SELECT" Then
'            Set vtag = vDoc.All(i)
'            If vtag.Type = "text" Or vtag.Type = "password" Or vtag.Type = "hidden" Or vtag.Type = "select-one" Then
'                If vtag.Value = XXXX Then XXXXX
'            End If
'        End If
'    Next
'---------------------------------------------------------------------------------
    
    AdviceCheckApplyCustom = True
    '----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceCheckApplyCustom = gcolPlugIn(i).AdviceCheckApplyCustom(lngSys, lngModual, lngPatiID, lng就诊ID, lng文件ID, strXML, objWeb, lng医嘱ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function LisPrintCodeBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal rsNumber As ADODB.Recordset) As Boolean
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'   rsNumber 本次打印的条码记录集中 包含的字段
'   rsNumber 中可能会存在多个条码
'    rsNumber.Fields.Append "类别", adVarChar, 20
'    rsNumber.Fields.Append "管码", adVarChar, 18
'    rsNumber.Fields.Append "相关ID", adBigInt
'    rsNumber.Fields.Append "样本条码", adVarChar, 18
'    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
'    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
'    rsNumber.Fields.Append "婴儿", adBigInt
'    rsNumber.Fields.Append "紧急标志", adBigInt
'    rsNumber.Fields.Append "标本", adVarChar, 30
'    rsNumber.Fields.Append "医嘱内容", adVarChar, 500
'    rsNumber.Fields.Append "采集方式", adVarChar, 100
'    rsNumber.Fields.Append "开嘱医生", adVarChar, 50
'    rsNumber.Fields.Append "开嘱时间", adDate
'    rsNumber.Fields.Append "采样人", adVarChar, 50
'    rsNumber.Fields.Append "采样时间", adDate
'    rsNumber.Fields.Append "采血量", adVarChar, 20
'    rsNumber.Fields.Append "试管名称", adVarChar, 50
'    rsNumber.Fields.Append "病人来源", adInteger
'    rsNumber.Fields.Append "医嘱ID串", adVarChar, 500
'    rsNumber.Fields.Append "执行科室", adVarChar, 50
'    rsNumber.Fields.Append "婴儿姓名", adVarChar, 50
'    rsNumber.Fields.Append "婴儿性别", adVarChar, 50
'    rsNumber.Fields.Append "申请科室", adVarChar, 50
'  返回值未false 时候，条码不进行打印。
    LisPrintCodeBefore = True
    '----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            LisPrintCodeBefore = gcolPlugIn(i).LisPrintCodeBefore(lngSys, lngModual, rsNumber)
            Err.Clear
        Next i
    End If

End Function

Public Function SetPassDiag(ByVal lngPatientID As Long, ByVal lng主页ID As Long, ByRef rsDiag As ADODB.Recordset) As Boolean
'功能：合理用药监测传入诊断拦截
'参数：
'      lngPatientID=病人ID,lng主页ID= 主页ID
'      rsDiag=传入Pass的诊断记录;此记录由合理用药监测传入,此记录集允许增删改操作

'出参:
'      rsDiag返回加工后的诊断记录

    SetPassDiag = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            SetPassDiag = gcolPlugIn(i).SetPassDiag(lngPatientID, lng主页ID, rsDiag)
            Err.Clear
        Next i
    End If
End Function

Public Function SetPassExams(ByVal lngPatientID As Long, ByVal lng主页ID As Long, ByVal lng挂号ID As Long, ByRef strExamsInfo As String) As Boolean
'功能：合理用药监测传入检验信息
'参数：
'      lngPatientID=病人ID,lng主页ID= 主页ID,lng挂号ID-挂号ID
'出参:
'      strExamsInfo返回检验信息MXL\JSON字符串

    SetPassExams = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            SetPassExams = gcolPlugIn(i).SetPassExams(lngPatientID, lng主页ID, lng挂号ID, strExamsInfo)
            Err.Clear
        Next i
    End If
End Function

Public Function AdvcieBeforToReview(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lngAdviceID As Long, ByVal lngSendID As Long, ByVal str核对人 As String, ByVal str核对时间 As String, ByVal str执行人 As String, ByVal strXML As String) As Boolean
'功能：核对前调用，批量核对时每条医嘱调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1261=住院医生站,1262=住院护士站)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lngAdviceID=病人医嘱记录.id
'      lngSendID=医嘱发送号
'      str核对人=核对人
'      str核对时间=核对时间
'      str执行人=执行人
'      strXML=XML扩展参数
'返回：true允许核对，false不能核对

    AdvcieBeforToReview = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdvcieBeforToReview = gcolPlugIn(i).AdvcieBeforToReview(lngSys, lngModual, lngPatientID, lngClinicID, lngAdviceID, lngSendID, str核对人, str核对时间, str执行人, strXML)
            Err.Clear
        Next i
    End If
End Function

Public Function ChargeSaveValied(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int记录性质 As Long, ByVal bln门诊 As Boolean, _
                                 ByVal bln划价单 As Boolean, ByVal strNos As String, ByVal rsSaveItems As ADODB.Recordset) As Boolean
'功能：收费合法性检查
'入参：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int记录性质
'      bln门诊
'             --是否门诊费用
'      bln划价单
'      strNOs-门诊收费时，传入的划价单号（对本次收费的划价单号)
'      rsSaveItems
'             --当前保存的项目集，字段:病人ID，主页ID,单据序号, 序号,价格父号,收费细目ID，收入项目id，付数 ，数次，标准单价，应收金额 ，
'                                      实收金额，发生时间，项目编码，项目名称，费用类别,开单部门ID,开单人,执行部门ID


    ChargeSaveValied = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            ChargeSaveValied = gcolPlugIn(i).ChargeSaveValied(lngSys, lngModual, int记录性质, bln门诊, bln划价单, strNos, rsSaveItems)
            Err.Clear
        Next i
    End If
End Function

Public Sub ChargeSaveAfter(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, _
                                ByVal lng主页ID As Long, ByVal bln门诊 As Boolean, ByVal int记录性质 As Integer, ByVal strNos As String)
'功能：在保存记帐、收费单后(事务结束后)调用
'入参：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      lng病人ID（记帐表时，传入0)
'      lng主页ID（记帐表时，传入0)
'      bln门诊 -是否门诊费用
'      int记录性质-1-收费;2-记帐
'      strNO-单据号

'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            Call gcolPlugIn(i).ChargeSaveAfter(lngSys, lngModual, lngPatientID, lng主页ID, bln门诊, int记录性质, strNos)
            Err.Clear
        Next i
    End If
End Sub

Public Function CheckPathInItem(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal rsAdvice As ADODB.Recordset, ByRef lng路径项目ID As Long, ByRef str分类 As String, Optional ByVal int场合 As Integer = -1, Optional ByRef strReserve As String) As Boolean
'功能：对临床路径病人下医嘱时，传入当前下达的医嘱信息，由接口内部判断是否为路径外项目
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      rsAdvice=当前录入的一条或一组医嘱记录（中药配方、成套方案、批量复制医嘱在输入时不调用）
'      字段为表“病人医嘱记录”的大部分字段，包括：ID, 相关id, 序号, 病人来源（1=门诊,2=住院), 病人id, 挂号单/主页ID,
'                姓名, 性别, 年龄, 婴儿, 医嘱状态, 医嘱期效, 诊疗类别, 诊疗项目id, 收费细目id, 天数, 单次用量, 总给予量, 医嘱内容, 医生嘱托, 标本部位,
'                检查方法 , 执行标记, 执行频次, 频率次数, 频率间隔, 间隔单位, 执行时间方案, 计价特性, 执行科室id, 执行性质, 紧急标志, 可否分零,
'                开始执行时间,执行终止时间(住院调用才有，没填时为cdate(0)), 病人科室id, 开嘱科室id, 开嘱医生, 开嘱时间, 挂号单, 前提id, 摘要
'                ,除此之外，门诊调用时，还包含医嘱对应的诊断的疾病ID
'      int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=保留参数,用于扩展使用
'      lng路径项目ID ， str分类：即做为入参数，也当做出参影响主程序判断。 当传入的传不为0和空值，表示主程序已经匹配成功能，接口内部仍可以对参数赋值设为不匹配。如果要保留ZLHIS主程序判断结果，则原值返回。
'            1.当前接判断为不是路径外项目时，则返回 路径项目ID和分类名称，lng路径项目ID - 病人路径执行.项目id，str分类 - 临床路径项目.分类
'            2.当前接判断为是路径外项目时，则返回0和空串，lng路径项目ID=0，str分类=""
 
    CheckPathInItem = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            CheckPathInItem = gcolPlugIn(i).CheckPathInItem(lngSys, lngModual, lngPatientID, lngClinicID, rsAdvice, lng路径项目ID, str分类, int场合, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function InPatiCancelOutBefore(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByRef strReserve As String) As Boolean
'功能：病人撤消出院之前调用(病人入出管理撤消出院操作时调用)
'参数：lng病人ID   病人ID
'      strReserve=扩展参数
'返回：true 允许撤消出院，false不能撤消出院

    InPatiCancelOutBefore = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            InPatiCancelOutBefore = gcolPlugIn(i).InPatiCancelOutBefore(lng病人ID, lng主页ID, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function InPatiCheckInCancel(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngType As Long, Optional ByVal strReserve As String) As Boolean
'功能：入院登记界面关闭或者点取消退出按钮时调用
'参数：lng病人ID   lng主页ID，这两个参数可能没有值，当登记界面被清空或者未录入时这两个值均为0
'      lngType 0-正常登记,1-预约登记,2-接收预约
'      strReserve=扩展参数
    
    InPatiCheckInCancel = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            InPatiCheckInCancel = gcolPlugIn(i).InPatiCheckInCancel(lng病人ID, lng主页ID, lngType, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function TransfusionCall(ByVal lngSys As Long, ByVal lngModual As Long, ByVal str挂号单 As String, ByVal lng流水号 As Long, ByRef strContent As String, Optional ByRef strReserve As String) As Boolean
'功能：门诊输液叫号
'参数:lngSys,lngModual=当前调用接口的主程序系统号及模块号 1264 门诊输液模块
'    str挂号单 挂号的单据号
'    lng流水号 记录哪几组医嘱一起执行的,病人医嘱执行.流水号,接单时的流水号
'    strContent 播报内容，传入的是his这边固定的播报内容，如果觉得不够全面，可以自己组织播报内容，是可以返回到his这边的
'    strReserve 扩展参数
'返回值:true-外挂接口播报,false-his这边播报
'列子:
'strContent = "请小康到第2号诊室输液"
    TransfusionCall = True
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            TransfusionCall = gcolPlugIn(i).TransfusionCall(lngSys, lngModual, str挂号单, lng流水号, strContent, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function TransfusionShowPatiList(ByVal lngSys As Long, ByVal lngModual As Long, ByVal lng科室ID As Long, ByVal BeginTime As Date, ByVal EndTime As Date, Optional ByRef strReserve As String) As Boolean
'功能：门诊输液系统病人列表刷新时调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号 1264 门诊输液模块
'      lng科室id 科室，
'      BeginTime , EndTime 界面过滤时间范围，开始时间，结束时间
'      strReserve=扩展参数

    TransfusionShowPatiList = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            TransfusionShowPatiList = gcolPlugIn(i).TransfusionShowPatiList(lngSys, lngModual, lng科室ID, BeginTime, EndTime, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function ChargeDeleteAndStrictValied(ByVal lngModule As Long, ByVal int记录性质 As Integer, ByVal bln门诊 As Boolean, _
    ByVal bln划价单 As Boolean, ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查删除划价单或对收费进行退费、和记帐，检查保存数据的合法性
    '入参:lngModule-模块号
    '     int记录性质-1-收费单;2-记帐单
    '     bln划价单-是否当前是保存的划价单
    '     strNOs-格式1:NO1:序号1,序号2,||NO2:序号1,序号2
    '            络式2:NO1||NO2:序号1,序号2||...
    '            络式3:NO1||NO2||...
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2017-12-13 17:55:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    ChargeDeleteAndStrictValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowClinicHelp(lngOwnHwnd As Long, ByVal int场合 As Integer, ByVal int病人来源 As Integer, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str项目IDs As String, Optional ByRef strReserve As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:外接第三方诊疗参考
'入参:lngOwnHwnd：父窗体句柄
'         int场合：0-申请，1-报告
'         int病人来源：1-门诊，2-住院
'         lng病人ID：病人ID
'         lng就诊ID：主页ID/挂号ID
'         str项目IDs：诊疗项目IDs（例：111,222,333）
'         strReserve 扩展参数，格式特殊符号分割,保留关键字 <ITEM> <SP1> <SP2>
'               信息1<ITEM>信息2<ITEM>信息3....
'               当传入的是检查类诊疗项目时 信息1 被使用通过特殊符号 <SP1>，<SP2>分割
'               [检查项目ID]<SP2>[部位1]<SP2>[部位2]<SP2>[部位2]<SP1>[检查项目ID]<SP2>[部位1]<SP2>[部位2]<SP2>[部位2]......
'               例:231<SP2>头部<SP2>胸部<SP2>腹部<SP1>232<SP2>心脑<SP2>锁骨<SP2>颈椎
'               也可以不包含部位 231<SP1>232<SP1>233<SP2>颈椎
'         注意：应工作站调用时可能未指定病人,则入参[int病人来源、lng病人ID、lng就诊ID、str项目IDs、strReserve]传入值为空

    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:蒋廷中
    '日期:2019-6-12 10:39:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str检查项目信息 As String
    Dim i As Long
    Dim j As Long
    Dim var项目 As Variant
    Dim var明细 As Variant
    
    '示例代码，获取检查项目部位信息，为了保持兼容性请通过Split函来获取检查部位信息
    If strReserve <> "" Then
        str检查项目信息 = Split(strReserve, "<ITEM>")(0)
        
        If str检查项目信息 <> "" Then
        
            var项目 = Split(str检查项目信息, "<SP1>")
            
            For i = 0 To UBound(var项目)
            
                var明细 = Split(var项目(i), "<SP2>")
                
                '项目ID var明细(0)
                '部位1  var明细(1)
                '部位2  var明细(2)
                '......
                
            Next
        End If
    End If
    
    ShowClinicHelp = False
    
'----------扩展部件接口遍历调用------------------------
 
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            ShowClinicHelp = gcolPlugIn(i).ShowClinicHelp(lngOwnHwnd, int场合, int病人来源, lng病人ID, lng就诊ID, str项目IDs, strReserve)
            Err.Clear
        Next i
    End If
End Function

Public Function BeforeVitalSignsSave(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng文件ID As Long, ByVal strDataXML As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:体温单记录保存时调用此接口，接口返回成功体温单/记录单数据才允许保存
    '入参:
    '         lng病人ID：病人ID
    '         lng主页ID：主页ID
    '         lng文件ID：护理文件ID
    '         strDataXML：数据XML串，格式如下：
                        '<DATALIST>
                        '    <ITEMLIST>
                        '        <FSSJ>2019-6-12 15:34</FSSJ>
                        '        <ITEM>
                        '            <XMXH>1</XMXH>
                        '            <XMMC>体温</XMMC>
                        '            <DATA>37.5</DATA>
                        '            <BW>腋温</BW>
                        '        </TEIM>
                        '        <ITEM>
                        '            <XMXH>2</XMXH>
                        '            <XMMC>脉搏</XMMC>
                        '            <DATA>110</DATA>
                        '            <BW></BW>
                        '        </TEIM>
                        '        .......
                        '    <ITEMLIST>
                        '    <ITEMLIST>
                        '        <FSSJ>2019-6-12 14:34</FSSJ>
                        '        <ITEM>
                        '            <XMXH>1</XMXH>
                        '            <XMMC>体温</XMMC>
                        '            <DATA>37</DATA>
                        '            <BW>腋温</BW>
                        '        </TEIM>
                        '        <ITEM>
                        '            <XMXH>2</XMXH>
                        '            <XMMC>脉搏</XMMC>
                        '            <DATA>120</DATA>
                        '            <BW></BW>
                        '        </TEIM>
                        '        .......
                        '    <ITEMLIST>
                        '    .........
                        '</DATALIST>
    '出参:
    '返回:成功返回true,否则返回Fale
    '编制:刘鹏飞
    '日期:2019-6-12 15:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    BeforeVitalSignsSave = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            BeforeVitalSignsSave = gcolPlugIn(i).BeforeVitalSignsSave(lng病人ID, lng主页ID, lng文件ID, strDataXML)
            Err.Clear
        Next i
    End If
End Function

Public Function zlGetSelfPaymentMode(ByVal lngModule As Long, ByVal lng病人ID As Long, ByVal str主页Ids As String, _
                                                             ByVal lng结帐ID As Long, ByVal rsFeeLists As ADODB.Recordset, _
                                                             ByRef strSelfPaymentMode_Out As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取自费方式
    '入参:lngModule-调用模块号
    '       lng病人ID-病人ID
    '       str主页Ids-当前主页IDs,多个用逗号分离
    '       lng结帐ID-如果是医保病人，在医保结算后，调用本接口，传入当前的结帐ID,
    '                       如果是普通病人，在读取费用明细后，调用本接口，结帐ID为0,但传入了rsFeeLists记录集
    '       rsFeeLists-传入的本次结帐的费用明细数据（住院,单据号 ,项目,费目,ID,序号,记录性质,记录状态,执行状态,主页ID,计算单位,
    '                       数量, 价格,应收金额,实收金额,未结金额,结帐金额, 统筹金额,类型, 收费类别,收费类别名,费别, 婴儿费, 执行部门id,
    '                       科室,开单部门ID,开单人, 保险大类id,收费细目ID,门诊标志,排序, 医嘱序号,时间,登记时间）(即传入现结帐程序中的费用明细集）
    '出参:strSelfPaymentMode_Out-返回结算方式,本接口返回true时有效,格式为:结算方式1,结算金额1,结算号码1||结算方式2,结算金额2,结算号码2||....
    '返回:获取返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-05-05 14:41:27
    '调用时机:
    '     1.普通病人:在读取病人费用明细后调用本接口
    '     2.医保病人：进行医保结算后，调用本接口
    '应用场景：慈善救助
    '    1、慈善救助得患者需病区得护士进行身份的核对，患者在结帐的时候不会主动说自己是慈善救助的对象，所以让结账员手工录入的话，容易漏掉这部分。
    '    2、慈善救助金额算法是：如果是医保病人，需要在医保结算完成之后才按照自费金额的xx%计算的，普通病人则直接是自费费用的xx%。
    '--------------------------------------------------------------------------------------------------------------------------------------

    zlGetSelfPaymentMode = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            zlGetSelfPaymentMode = gcolPlugIn(i).zlGetSelfPaymentMode(lngModule, lng病人ID, str主页Ids, lng结帐ID, rsFeeLists, strSelfPaymentMode_Out)
            Err.Clear
        Next i
    End If

End Function

Public Function SaveBalanceBeforCheck(ByVal lngModule As Long, ByVal blnMZBalace As Boolean, _
                                                             ByVal rsBalanceListData As ADODB.Recordset, _
                                                             ByVal strCurBalance As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------
    '功能:结帐合法性检查
    '入参:lngModule-调用模块号
    '       blnMZBalace-是否门诊结帐
    '       rsBalanceListData-结算列表数据：结算方式,金额,摘要,结算号码,卡类别ID,消费卡,卡号,交易流水号,交易说明,关联交易ID,结帐ID
    '       strCurBalance:结算方式||结算金额||结算号码||摘要
    '出参:
    '返回:检查成功返回true,否则返回False
    '调用时机:
    '     在结帐数据保存前调用
    '     医保接口结帐时不调有
    '使用程序：病人结帐处理
    '--------------------------------------------------------------------------------------------------------------------------------------

    SaveBalanceBeforCheck = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            SaveBalanceBeforCheck = gcolPlugIn(i).SaveBalanceBeforCheck(lngModule, blnMZBalace, rsBalanceListData, strCurBalance)
            Err.Clear
        Next i
    End If

End Function
 
Public Function GetfinancialTrancsBalanceInfor(ByVal str结算方式 As String, ByVal dblMoney As Double, _
    ByVal lngModule As Long, ByRef str结算信息_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取财务转帐的结算信息
    '入参:str结算方式-当前转账交易的结算方式
    '     dblMoney-转账金额
    '出参:str结算信息_Out-结算号码||结算摘要||卡号
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-09-02 15:32:47
    '调用时机:
    '   根据参数”结帐预交退款控制”来控制，如果为1时，存在预交款退款且存在缺省为退款方式时，调用本接口
    '---------------------------------------------------------------------------------------------------------------------------------------------
    str结算信息_Out = ""  ' "JS0021||测试||Card0001"
    GetfinancialTrancsBalanceInfor = True
    '----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
            On Error Resume Next
            For i = 1 To gcolPlugIn.Count
                GetfinancialTrancsBalanceInfor = gcolPlugIn(i).GetfinancialTrancsBalanceInfor(str结算方式, dblMoney, lngModule, str结算信息_Out)
                Err.Clear
            Next i
    End If
End Function

Public Function SaveRegisterAfter(ByVal lngSys As Long, ByVal lngModule As Long, ByVal strNO As String, ByVal bln预约 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 挂号完成后调用
    ' 入参 : lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '        strNo-挂号单号
    '        bln预约-预约挂号
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/22 10:03
    '---------------------------------------------------------------------------------------
    
    SaveRegisterAfter = True
    
    '----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            SaveRegisterAfter = gcolPlugIn(i).SaveRegisterAfter(lngSys, lngModule, strNO, bln预约)
            Err.Clear
        Next i
    End If
End Function

Public Sub AfterSelectorReady(ByVal intType As Integer, ByVal strTitle As String, ByVal objList As Object, _
    Optional ByVal frmMain As Object, Optional ByVal strExt As String)
'编制：周韬，2019-11-08
'功能：在业务功能的选择器数据准备好之后调用，可对选择器中的数据显示控件进行二次开发处理
'参数：
'    intType：1-诊疗项目选择器，2-收费项目选择器，99-通用选择器(内容依据SQL动态构建)
'    strTitle：选择器功能标题，不一定严格描述
'    objList：选择器中显示数据清单的控件，可能是不同类型，如：VSFlexGrid，ListView
'    frmMain：选择器所在的主窗体(某些情况下可能没传为Nothing
'    strExt：扩展参数
'       当intType=1时：strExt="调用场合;病人病区ID;病人科室ID;医技科室ID"。
'           其中调用场合=(-1)-成套方案编辑,0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'       当intType=2时：strExt="病人来源：1-门诊，2-住院"
    
    Dim lngIDCol As Long
    Dim i As Long
    
    Exit Sub
    
    If intType = 1 Then '开医嘱的诊疗项目选择器，VSFlexGrid
        '关键列：诊疗项目ID,收费项目ID(药品ID),类别ID(类别字母),类别(类别名称),编码,名称
        If TypeName(objList) <> "VSFlexGrid" Then Exit Sub
        
        lngIDCol = GetFlexCol(objList, "诊疗项目ID")
        If lngIDCol = -1 Then Exit Sub
        
        For i = objList.Rows - 1 To objList.FixedRows Step -1
            If Val(objList.TextMatrix(i, lngIDCol)) = -1234 Then
                objList.RemoveItem i
            End If
        Next
        
        If objList.Rows > objList.FixedRows Then
            objList.Row = objList.FixedRows
        End If
    ElseIf intType = 2 Then '收费、记帐、补费等用的收费项目选择器，VSFlexGrid
        '关键列：ID,类别ID(类别字母),类别(类别名称),编码,名称
        If TypeName(objList) <> "VSFlexGrid" Then Exit Sub
        
        lngIDCol = GetFlexCol(objList, "ID")
        If lngIDCol = -1 Then Exit Sub
        
        For i = objList.Rows - 1 To objList.FixedRows Step -1
            If Val(objList.TextMatrix(i, lngIDCol)) = -1234 Then
                objList.RemoveItem i
            End If
        Next
        
        If objList.Rows > objList.FixedRows Then
            objList.Row = objList.FixedRows
        End If
    ElseIf intType = 99 Then '其他公共选择器，ListView
        If TypeName(objList) <> "ListView" Then Exit Sub
        If objList.ListItems.Count = 0 Then Exit Sub
        
        If strTitle = "XX选择" Or frmMain.Caption = "XX功能" Then
            For i = objList.ListItems.Count To 1 Step -1
                'Key的形式为"_项目ID"，或"行号_项目ID"
                If Val(Split(objList.ListItems(i).Key, "_")(1)) = -1234 Then
                    objList.ListItems.Remove i
                End If
            Next
            
            If objList.ListItems.Count > 0 And objList.SelectedItem Is Nothing Then
                objList.ListItems(1).Selected = True
            End If
        End If
    End If
End Sub

Public Function InPatiLoadBedInfo(ByVal lng科室ID As Long, ByVal lng病区ID As Long, ByRef strBedInfo As String, ByRef blnNoEmptyBed As Boolean, Optional ByVal strReserve As String) As Boolean
'功能：入院登记界面选择病区或选择科室时候调用;科室，病区显示病人在院人数，床位数
'参数：lng科室ID
'      lng病区ID
'      strReserve=扩展参数
'出参:
'      strBedInfo- 返回在院人数及床位数信息。 格式示例:  "XX科室:在院人数 M，床位数 N"
'      blnNoEmptyBed-当参数:"科室下无空床不能登记"启用时,blnNoEmptyBed的值决定是否弹出提示,禁止入院登记。
'                    T-当前科室没有空床位;F-当前科室有空床位
'返回：true-外挂接口规则显示在院人数,床位数,false-HIS规则显示 在院人数,床位数 (默认false)
    InPatiLoadBedInfo = False
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            InPatiLoadBedInfo = gcolPlugIn(i).InPatiLoadBedInfo(lng科室ID, lng病区ID, strBedInfo, blnNoEmptyBed, strReserve)
            Err.Clear
        Next i
    End If
End Function


Public Function AdviceEditBefore(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int场合 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str挂号单 As String, ByVal lng前提ID As Long, ByVal int婴儿 As Integer, ByVal lng界面科室ID As Long, ByVal lng危急值ID As Long, ByVal lng会诊医嘱ID As Long, Optional ByRef strReserve As String) As Boolean
'功能：新增医嘱窗体打开前调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号(1252=门诊医嘱下达,1253=住院医嘱下达)
'      int场合       =调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      lng病人ID     =病人ID
'      lng主页ID     =主页ID(住院时调用)
'      str挂号单     =挂号单(门诊时调用)
'      lng前提ID     =医技工作站下医嘱时用
'      int婴儿       0=表示病人本人
'      lng界面科室ID 当前工作站界面科室ID
'      lng危急值ID
'      lng会诊医嘱ID
'      strReserve   扩展参数
'返回：true允许编辑，false不能编辑
 
    AdviceEditBefore = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AdviceEditBefore = gcolPlugIn(i).AdviceEditBefore(lngSys, lngModual, int场合, lng病人ID, lng主页ID, str挂号单, lng前提ID, int婴儿, lng界面科室ID, lng危急值ID, lng会诊医嘱ID, strReserve)
            Err.Clear
        Next i
    End If
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
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PeisDelItemCheck = gcolPlugIn(i).PeisDelItemCheck(lngSys, lngModual, lngPatientID, lngTaskId, lngListId, strMsg, strReserve)
            Err.Clear
        Next i
    End If
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
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            PeisAbortSignInCheck = gcolPlugIn(i).PeisAbortSignInCheck(lngSys, lngModual, lngPatientID, lngTaskId, lngListId, strNO, lngRecType)
            Err.Clear
        Next i
    End If
End Function

Public Sub zlChargeApplyAudit(ByVal frmMain As Object, ByVal lngModule As Long, ByVal str费用IDs As String)
    '-------------------------------------------------------------------------------------------------------------------------------------
    '功能:在销账审核完成后，获取到销账审核的数量
    '入参:lngModule-调用模块号
    '       frmMain-调用的主窗体
    '       str费用IDs-费用IDs：费用id1,费用id2,费用id3......
    '出参:
    '返回:调用成功返回true,否则返回False
    '调用时机:
    '     在销账审核完成后，调用该接口
    '使用程序：住院护士工作站
    '--------------------------------------------------------------------------------------------------------------------------------------
    
    '----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            Call gcolPlugIn(i).zlChargeApplyAudit(frmMain, lngModule, str费用IDs)
            Err.Clear
        Next i
    End If

End Sub

Public Function ChargePriceVerfiyValied(ByVal lngSys As Long, ByVal lngModule As Long, ByVal bln门诊 As Boolean, ByVal strNos As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐划价审核合法性检查接口
    '入参：lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '         bln门诊 -是否门诊费用
    '         strNos-划价单据号1:序号|划价单据号2:序号|...    （序号-每张单据审核的行号,多个用逗号分离,为空代表整张单据审核),各单据之间用"|"分隔
    '出参:
    '返回:调用成功返回true,否则返回False
    '调用时机:
    '     保存记帐审核 前调用
    '使用程序：费用系统1.门诊记帐划价审核;
    '2 住院记帐划价审核:记帐单划价审核，记账表划价审核、按病人审核，批量审核
    '--------------------------------------------------------------------------------------------------------------------------------------

    ChargePriceVerfiyValied = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            ChargePriceVerfiyValied = gcolPlugIn(i).ChargePriceVerfiyValied(lngSys, lngModule, bln门诊, strNos)
            Err.Clear
        Next i
    End If

End Function

Public Sub ChargePriceVerfiyAfter(ByVal lngSys As Long, ByVal lngModule As Long, ByVal bln门诊 As Boolean, ByVal strNos As String)
    '-------------------------------------------------------------------------------------------------------------------------------------
    '功能:在保存记帐划价审核数据后(事务结束后)调用
    '入参：lngSys,lngModual=当前调用接口的主程序系统号及模块号
    '         bln门诊 -是否门诊费用
    '         strNos-划价单据号1:序号|划价单据号2:序号|...    （序号-每张单据审核的行号,多个用逗号分离,为空代表整张单据审核),各单据之间用"|"分隔
    '出参:
    '返回:调用成功返回true,否则返回False
    '调用时机:
    '     保存记帐审核 前调用
    '使用程序：费用系统1.门诊记帐划价审核;
    '2 住院记帐划价审核:记帐单划价审核，记账表划价审核、按病人审核，批量审核
    '--------------------------------------------------------------------------------------------------------------------------------------
    
    '----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            Call gcolPlugIn(i).ChargePriceVerfiyAfter(lngSys, lngModule, bln门诊, strNos)
            Err.Clear
        Next i
    End If

End Sub

Public Sub QueueCallBefore(ByVal QueueID As Long, ByVal CallType As Integer, ByVal ServiceType As Integer, ByRef blnContinue As Boolean)
'QueueID: 排队叫号队列.ID:用于查询排队叫号数据的相关信息
'CallType:  呼叫操作类型 :0-顺呼，1-直呼，2-重呼，3-弃呼
'ServiceType: 业务类型:   0表示门诊排队，1表示pacs排队，2表示lis排队

'blnContinue 是否继续: 如果使用了自定义的叫号方式,这个参数应传递true  否则传递false
    MsgBox "收到：" & QueueID
    blnContinue = True
End Sub

