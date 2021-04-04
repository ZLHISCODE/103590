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

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal int场合 As Integer = -1)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合  调用场合:-1
'      病历调用 lngModual=1070，无int场合参数
    If Not gblnInited Then
        Set gcnOracle = cnOracle
        Call HandlePlugIn(1, lngSys, lngModual, cnOracle, int场合)
        gblnInited = True
    End If
End Sub

Public Function GetFormCaption(ByVal lngSys As Long, ByVal lngModual As Long) As String
'功能：在科室配血管理中添加卡片的名称，逗号分割。
    Dim strTmp As String, strThis As String
    'Exit Function '如使用此功能请修改接口内部实现
    strTmp = GetFormCaptionEx(lngSys, lngModual)
    
    Select Case lngModual
    Case m科室配血管理
        strThis = "Test1,Test2"
    End Select
    GetFormCaption = strThis & IIf(strTmp = "", "", "," & strTmp)
End Function

Public Function GetForm(ByVal lngSys As Long, ByVal lngModual As Long, ByVal strName As String) As Object
'功能：根据索引值返回对象。工作站窗体加载时调用
        'Exit Function '如使用此功能请修改接口内部实现
    Select Case lngModual
    Case m科室配血管理
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
    Call HandlePlugIn(3, lngSys, lngModual, , int场合, strReserve, strFuncName, lngPatiID, varRecId, varKeyId)
End Sub

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
    Call HandlePlugIn(4, lngSys, lngModual, , int场合)
End Sub

Public Function AfterBloodPrepareRecept(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lngClinicID As Long, _
    ByVal lng申请ID As Long) As Boolean
'功能：接受配血后调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      intMode：1-接受申请，2-取消接受申请
'      lngPatientID=病人ID,lngClinicID=就诊ID（挂号记录ID）/主页ID(住院调用时)
'      lng申请ID=医嘱id(病人医嘱状态.医嘱id)

    AfterBloodPrepareRecept = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodPrepareRecept = gcolPlugIn(i).AfterBloodPrepareRecept(lngSys, lngModual, intMode, lngPatientID, lngClinicID, lng申请ID)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterBloodVerify(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lng配发ID As Long, _
    ByVal strIDs As String, Optional ByVal strTime As String = "") As Boolean
'功能：备血完成后调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      intMode：1-配血审核，2-取消配血审核
'      lng配发ID = 配发ID
'      strIDs：收发ID串，以“，”分隔
'      strTime：审核时间，可能不传入（为空）

    AfterBloodVerify = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodVerify = gcolPlugIn(i).AfterBloodVerify(lngSys, lngModual, intMode, strIDs, strTime)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterSendBlood(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lngPatientID As Long, ByVal lng配发ID As Long, _
        ByVal strIDs As String) As Boolean
'功能：发血或退血完成后调用一次
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      intMode：1-发血，2-退血
'      lngPatientID = 病人ID
'      lng配发ID = 配发ID
'      strIDs = 收发ID串，以“，”分隔

    AfterSendBlood = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterSendBlood = gcolPlugIn(i).AfterSendBlood(lngSys, lngModual, intMode, lngPatientID, lng配发ID, strIDs)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterBloodPurchase(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal strNO As String) As Boolean
'功能：血液供应入库保存时调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      intMode：1-供应入库审核，2-供应入库冲销
'      strNO-单据号

    AfterBloodPurchase = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodPurchase = gcolPlugIn(i).AfterBloodPurchase(lngSys, lngModual, intMode, strNO)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterBloodReaction(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal lng收发id As String) As Boolean
'功能：输血反应记录保存后调用
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      intMode：1-医生提交或输血科保存时，2-医生回退或者输血科删除时
'      lng收发id-输血反应记录.收发ID

    AfterBloodReaction = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodReaction = gcolPlugIn(i).AfterBloodReaction(lngSys, lngModual, intMode, lng收发id)
            Err.Clear
        Next i
    End If
End Function

Public Function AfterBloodScrap(ByVal lngSys As Long, ByVal lngModual As Long, ByVal intMode As Integer, ByVal strNO As String) As Boolean
'功能:血液报废审核或冲销后调用(血液报废出库)
'参数：strBloodIDs=本次发送血液的血液收发ID串，多个收发ID之间用","分割，如：1001,1002
'          intMode=0 ：报废登记;=1 报废登记删除
'说明：可根据strBloodIDs从"血液收发ID中获取报废相关信息"
AfterBloodScrap = True
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            AfterBloodScrap = gcolPlugIn(i).AfterBloodScrap(lngSys, lngModual, intMode, strNO)
            Err.Clear
        Next i
    End If
End Function

Public Function BloodStoreImport(ByRef lng入库方式 As Long, ByRef rsData As ADODB.Recordset)
'功能：血液入库批量导入，由血库系统进行调用
'参数：rsData 数据集 数据在此接口进行添加或加载
'      lng入库方式 '1-外购入库;3 ：其他入库
'说明：参数用引用方式传入，在接口中进行赋值，返回后由调用方会直接使用这两个参数；记录集 rsData 请按以下字段组织记录集即可
'处理方式：首先需要将中心的规格和HIS中的血液规格进行对照工作(可在血液目录管理提供的扩展功能中完成)，导入时根据中心的规格编码获取到HIS中对照的规格ID。然后再根据规格ID查找记录集中相关字段信息
'字段：
'  ID           Number(18),--血液规格ID
' 行次         Number(5), --序号：从1开始，按照记录数排序
'  血液名称  Varchar2(200), --品种名称
'  规格         Varchar2(100), --收费项目目录中的规格
'  ABO         Varchar2(20),
'  RH(D)       Varchar2(1),
'  RH(C)       Varchar2(2),
'  RH (c)      Varchar2(2),
'  RH(E)       Varchar2(2),
'  RH (e)      Varchar2(2),
'  单位         Varchar2(20), --收费项目目录中的计算单位
'  生产时间     Date, --供应入库填写
'  效期         Date, --供应入库填写
'  颜色         Varchar2(20),
'  外观         Varchar2(20), --供应入库填写
'  渗漏         Number(1), --0 否 1是
'  数量         Number(16, 5), --袋数，一般都是1
'  售价         Number(16, 5), -- 收费价目中的现价
'  采购价      Number(16, 5), --血液规格中的成本价
'  差价金额     Number(16, 5),--售价-采购价
'  血袋编号     Varchar2(30), --血袋的唯一标识
'  复查ABO      Varchar2(20),
'  复查RH(D)    Varchar2(1),
'  复查RH(C)    Varchar2(2),
'  复查RH (c)   Varchar2(2),
'  复查RH(E)    Varchar2(2),
'  复查RH (e)   Varchar2(2),
'  采血人       Varchar2(20), --供应入库填写
'  献血者       Varchar2(20),
'  献血时间     Date,
'  送血者       Varchar2(20), --供应入库填写
'  送血时间     Date, --供应入库填写
'  送血箱内温度 Number(16, 5), --供应入库填写
'  接收者       Varchar2(20)  --供应入库填写
    Dim strFileds As String, strValues As String
    '记录集初始化
    Set rsData = New ADODB.Recordset
    strFileds = "ID," & adLongVarBinary & ",18|行次," & adLongVarBinary & ",5|序号," & adLongVarBinary & ",5|血液名称," & adLongVarChar & ",200|规格," & adLongVarChar & ",100|" & _
            "ABO," & adLongVarChar & ",20|RH(D)," & adVarChar & ",1|RH(C)," & adVarChar & ",2|RH (c)," & adVarChar & ",5|RH(E) ," & adVarChar & ",5|" & _
            "RH (e)," & adVarChar & ",5|单位," & adLongVarChar & ",20|生产时间," & adDate & ",20|效期," & adDate & ",20|颜色," & adLongVarChar & ",20|" & _
            "外观," & adLongVarChar & ",20|渗漏," & adNumeric & ",1|数量," & adDouble & ",16|售价," & adDouble & ",16|采购价," & adDouble & ",16|" & _
            "差价金额," & adDouble & ",16|血袋编号," & adLongVarChar & ",30|复查ABO," & adLongVarChar & ",20|复查RH(D)," & adVarChar & ",1|复查RH(C)," & adVarChar & ",2|" & _
            "复查RH (c," & adVarChar & ",2|复查RH(E)," & adVarChar & ",2|复查RH (e)," & adVarChar & ",2|采血人," & adLongVarChar & ",20|献血者," & adLongVarChar & ",20|" & _
            "献血时间," & adDate & ",20|送血者," & adLongVarChar & ",20|送血时间," & adDate & ",20|送血箱内温度," & adDouble & ",16|接收者," & adLongVarChar & ",20"
    Call Record_Init(rsData, strFileds)
    '调用Record_Update或Record_Add完成数据增加或更新操作
    BloodStoreImport = True
    
'----------扩展部件接口遍历调用------------------------
    Dim i As Long
    If Not gcolPlugIn Is Nothing Then
        On Error Resume Next
        For i = 1 To gcolPlugIn.Count
            BloodStoreImport = gcolPlugIn(i).BloodStoreImport(lng入库方式, rsData)
            Err.Clear
        Next i
    End If
End Function
