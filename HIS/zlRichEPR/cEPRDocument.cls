VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'################################################################################################################
'## 枚举常量
'################################################################################################################

'当前编辑模式
Public Enum EditModeEnum
    cprEM_新增 = 0              '新增
    cprEM_修改 = 1              '修改
End Enum

'当前编辑器的编辑类型
Public Enum EditTypeEnum
    cprET_病历文件定义 = 0      '病历文件定义
    cprET_全文示范编辑 = 1      '全文示范编辑
    cprET_单病历编辑 = 2        '单病历编辑
    cprET_单病历审核 = 3        '单病历审核
End Enum

'文件种类  1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-诊断文书;6-知情文件;7-诊疗报告;8-诊疗申请
Public Enum EPRDocTypeEnum
    cpr门诊病历 = 1
    cpr住院病历 = 2
    cpr护理记录 = 3
    cpr护理病历 = 4
    cpr诊断文书 = 5
    cpr知情文件 = 6
    cpr诊疗报告 = 7             '诊疗单据：报告
    cpr诊疗申请 = 8             '诊疗单据：申请
End Enum

'病人来源
Public Enum PatiFromEnum
    cprPF_门诊 = 1              '1-门诊；
    cprPF_住院 = 2              '2-住院；
    cprPF_外来 = 3              '3-外来；
    cprPF_体检 = 4              '4-体检
End Enum

'签名状态
Public Enum EPRSignLevelEnum
    cprSL_空白 = 0              '未签名
    cprSL_经治 = 1              '经治医师签名
    cprSL_主治 = 2              '主治医师签名
    cprSL_主任 = 3              '主任医师签名
    cprSL_正高 = 4              '正高：签名级别不包含，只表示人员居右正高职称，以便区别副主任医师
End Enum

'文本审核状态
Public Enum TextAuditStateEnum
    cprAS_NormalText = 0        '普通文本
    cprAS_Deletedtext = -1      '已删除文本
    cprAS_NewText = 1           '新增文本
    cprAS_MixedText = 2         '混合文本
    cprAS_ProtedText = 3        '手工设置的保护文本
End Enum

'################################################################################################################
'## 公共变量
'################################################################################################################

Public Tables As cEPRTables                     '表格集合
Public Pictures As cEPRPictures                 '图片集合
Public Compends As cEPRCompends                 '提纲集合
Public Elements As cEPRElements                 '诊治要素集合
Public Diagnosises As cEPRDiagnosises           '诊断集合
Public Signs As cEPRSigns                       '签名集合

Public EPRFileInfo As cEPRFileDefineInfo        '病历文件定义信息
Public EPRDemoInfo As cEPRDemoInfo              '病历示范信息
Public EPRPatiRecInfo As cEPRPatiRecordInfo     '病人电子病历信息

Public EditMode As EditModeEnum                     '编辑模式（新增、修改）
Public EditType As EditTypeEnum                     '编辑类型（病历文件定义、病历示范编辑、单病历编辑、单病历审核）

Public 用户签名级别 As EPRSignLevelEnum             '用户签名级别，判断用户是否有权限审签

Public IsVerifyMode As Boolean                      '是否是验证模式（不编辑）
Public KeepRTF As Boolean                           '是否保持RTF文本不被替换

Public mfrmParent As Object                         '父窗体
Public frmEditor As Object                          '主编辑窗体
Private mTextIDCol As Collection
Private mIntTextIDPos As Integer

Public Event BeforeSaved(lngRecordId As Long, Cancel As Boolean)    'EMR文档保存事件，返回记录集
Public Event AfterSaved(lngRecordId As Long)                        'EMR文档保存事件，返回记录集
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)       '点击诊断参考事件，以便调阅参考

Public Sub Event_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub
Public Sub AfterClosed(lngOrder As Long)
    If Me.EPRPatiRecInfo.病历种类 = cpr诊疗报告 And (Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核) Then
        If InStr(mfrmParent.Caption, "诊疗报告管理") > 0 Then
            Call mfrmParent.EditorClosed(lngOrder)
        End If
    End If
End Sub
Public Sub AfterPrinted(lngRecordId As Long)
    If mfrmParent Is Nothing Then Exit Sub
    If (Me.EPRPatiRecInfo.病历种类 = cpr诊疗报告 Or Me.EPRPatiRecInfo.病历种类 = cpr住院病历) _
        And (Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核) Then
        If InStr("住院病历记录,诊疗报告管理", mfrmParent.Caption) > 0 Then
            Call mfrmParent.Event_AfterPrinted(lngRecordId)
            Exit Sub
        End If
    End If
    On Error Resume Next
    mfrmParent.RefreshList: Err.Clear
End Sub
'################################################################################################################
'## 功能：  判断当前的病历文件是否采用“多文档编辑”模式打开
'##
'## 返回：  如果是，返回True；否则，返回False。
'##
'## 注释：  如果是多文档中的第一篇文档，则采用“单文档编辑”方式。默认只显示一个月内的文档。
'################################################################################################################
Public Function IsMultiEPRDoc() As Boolean
    On Error Resume Next
    Dim i As Long
    Dim rs As New ADODB.Recordset
    
    If Me.EditType = cprET_病历文件定义 Or Me.EditType = cprET_全文示范编辑 Then IsMultiEPRDoc = False: Exit Function
    gstrSQL = "Select count(l.Id) 数量 From 病历文件列表 l, 病历文件列表 m Where l.种类 = m.种类 And l.页面 = m.页面 And m.Id =[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", Me.EPRFileInfo.ID)
    '如果该页面的共用文件数目小于或等于1，那么不是多文件编辑
    If rs(0) <= 1 Then
        IsMultiEPRDoc = False
        Exit Function
    End If
    rs.Close
    
    Dim eEPRType As EPRDocTypeEnum
    If Me.EditMode = cprEM_新增 Then
        eEPRType = EPRFileInfo.种类
    Else
        eEPRType = EPRPatiRecInfo.病历种类
    End If
    
    If eEPRType = 0 Then
        IsMultiEPRDoc = False
        Exit Function
    End If
    
    '如果是共用文件，判断一月内是否书写过共用病历或转入后要求填写病历
    gstrSQL = "Select Count(r.Id) " & _
        " From 电子病历记录 r, " & _
        "      (Select l.Id From 病历文件列表 l, 病历文件列表 m Where l.种类 = m.种类 And l.页面 = m.页面 And m.Id =[1]) f" & _
        " Where r.文件id = f.Id And r.病人id = [2] And r.主页id = [3] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", EPRFileInfo.ID, EPRPatiRecInfo.病人ID, EPRPatiRecInfo.主页ID)
    
    If rs.EOF Then
        IsMultiEPRDoc = False
    ElseIf rs(0) > 1 Then
        IsMultiEPRDoc = True
    ElseIf rs(0) = 1 Then
        If Me.EditMode = cprEM_新增 Then
            '增加第二篇共用页面文档。
            IsMultiEPRDoc = True
        Else
            '第一篇共用文档的修改，此时也是“单文档编辑”模式。
            IsMultiEPRDoc = False
        End If
    Else
        '首次增加共用页面文档。
        IsMultiEPRDoc = False
    End If
    rs.Close
    Set rs = Nothing
End Function

'################################################################################################################
'## 功能：  初始化并打开指定文档（指定记录ID和版本号），用于签名验证
'##
'## 参数：  lngRecID        :电子病历记录ID
'##         lngVersion      :版本号
'##         blnIsVerifyMode :是否是签名验证模式
'##         blnKeepRTF      :是否保持RTF文本不变？用于导出到XML文件（不变）
'##
'## 说明：  该方法用于生成用于验证的文本字符串，打开指定ID和版本的文档。
'################################################################################################################
Public Sub InitAndOpenEPR(ByVal lngRecId As Long, _
    ByVal lngVersion As Long, _
    Optional ByVal blnIsVerifyMode As Boolean = True, _
    Optional ByVal blnKeepRTF As Boolean = False)
    
    Dim rs As New ADODB.Recordset
    '1、基础信息初始化
    Me.IsVerifyMode = blnIsVerifyMode   '是否属于验证模式
    Me.KeepRTF = blnKeepRTF
    Me.EditType = cprET_单病历编辑
    
    gstrSQL = "Select ID, 病人来源, 病人ID, 主页ID, 婴儿, 科室ID, 病历种类, 文件ID, " & _
        " 病历名称, 归档人, 归档日期, 保存人, 保存时间, 创建人, 创建时间, " & _
        " 最后版本 , 签名级别 , 完成时间 From 电子病历记录 Where ID=[1] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngRecId)  '此时的lngRecId表示病人电子病历的ID
    With EPRPatiRecInfo
        .ID = lngRecId
        If Not rs.EOF Then
            .病人来源 = NVL(rs("病人来源"), 1)
            .病人ID = NVL(rs("病人ID"), 0)
            .主页ID = NVL(rs("主页ID"), 0)
            .婴儿 = NVL(rs("婴儿"), 0)
            .科室ID = NVL(rs("科室ID"), 0)
            .病历种类 = NVL(rs("病历种类"), 0)
            .文件ID = NVL(rs("文件ID"), 0)
            .病历名称 = NVL(rs("病历名称"))
            .创建人 = NVL(rs("创建人"))
            .创建时间 = NVL(rs("创建时间"), 0)
            .保存人 = NVL(rs("保存人"))
            .保存时间 = NVL(rs("保存时间"), 0)
            .归档人 = NVL(rs("归档人"))
            .归档日期 = NVL(rs("归档日期"), 0)
            .完成时间 = NVL(rs("完成时间"), 0)
            .最后版本 = NVL(rs("最后版本"), 1)
            .签名级别 = NVL(rs("签名级别"), 0)     '记录本次编辑的初始版本状态
        End If
    End With
    rs.Close
    
    gstrSQL = "Select a.ID, a.种类, a.编号, a.名称, a.说明, a.页面, a.保留, a.通用, b.名称 As 页面名称, b.报表, b.格式, b.页眉, b.页脚 " & _
        " From 病历文件列表 a, 病历页面格式 b " & _
        " Where a.页面 = b.编号 And a.种类 = b.种类 And a.Id = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", EPRPatiRecInfo.文件ID)
    With EPRFileInfo
        If Not rs.EOF Then
            .ID = NVL(rs("ID"), 0)
            .保留 = NVL(rs("保留"), 0)
            .编号 = NVL(rs("编号"), 0)
            .名称 = NVL(rs("名称"))
            .说明 = NVL(rs("说明"))
            .通用 = NVL(rs("通用"), 0)
            .页面 = NVL(rs("页面"))
            .种类 = NVL(rs("种类"), 1)
            .报表 = NVL(rs("报表"), 0)
            .页面名称 = NVL(rs("页面名称"))
            .格式 = NVL(rs("格式"))
            .页眉 = NVL(rs("页眉"))
            .页脚 = NVL(rs("页脚"))
        End If
    End With
    
    '2、调入文件结构和内容
    Screen.MousePointer = vbHourglass
    frmEditor.Editor1.Tag = "禁止同步"
    frmEditor.Editor1.InProcessing = True
    Set Me.Compends = New cEPRCompends
    Set Me.Pictures = New cEPRPictures
    Set Me.Tables = New cEPRTables
    Set Me.Elements = New cEPRElements
    Set Me.Diagnosises = New cEPRDiagnosises
    Set Me.Signs = New cEPRSigns
    
    Call ReadFileStructure(frmEditor.Editor1)
    Screen.MousePointer = 0
    frmEditor.Editor1.InProcessing = False
    frmEditor.Editor1.Modified = False
    Set frmEditor.Document = Me
    frmEditor.Editor1.Tag = ""
End Sub

'################################################################################################################
'## 功能：  类初始化
'##
'## 参数：  eEdtMode    :当前编辑模式（新增、修改）
'##         eEdtType    :当前编辑方式（文件定义、示范编辑、单病历编辑、单病历审核）
'##         lngFileID   :文件ID（根据编辑方式的不同，可以表示文件定义ID、范文ID或者病人病历ID）
'##
'##         (以下为可选参数，用于病人电子病历编辑和审核):
'##
'##         PatiFrom    :病人来源
'##         lngPatiID   :病人ID
'##         intTimesNo  :住院次数，主页ID；门诊则为挂号ID
'##         intBabyOrder:婴儿数
'##         intCurDeptID:当前病人病历填写科室ID
'##         医嘱ID      :医嘱ID，用于特殊病历
'################################################################################################################
Public Sub InitEPRDoc(ByVal eEdtMode As EditModeEnum, _
    ByVal eEdtType As EditTypeEnum, _
    ByVal lngFileID As Long, _
    Optional ByVal PatiFrom As PatiFromEnum = cprPF_住院, _
    Optional ByVal lngPatiID As Long = 0, _
    Optional ByVal intTimesNo As Long = 0, _
    Optional ByVal intBabyOrder As Integer = 0, _
    Optional ByVal intCurDeptID As Long = 0, _
    Optional ByVal lng医嘱id As Long = 0, _
    Optional ByVal blnMoved As Boolean)
    
'    On Error GoTo LL
    Me.EPRPatiRecInfo.医嘱id = lng医嘱id
    
    Me.用户签名级别 = GetUserSignLevel(glngUserId, , lngPatiID, intTimesNo)   '用户的签名级别
    
    Me.EditMode = eEdtMode      '新增、修改；
    Me.EditType = eEdtType      '病历文件定义、全文示范编辑、单病历编辑、单病历审核
    
    frmEditor.Editor1.AuditMode = (Me.EditType = cprET_单病历审核)      '是否是审核模式
    
    '信息提取
    Dim rs As New ADODB.Recordset, lngTMP As Long
    rs.CursorLocation = adUseClient
    Select Case EditType
    Case cprET_病历文件定义
        If eEdtMode = cprEM_修改 Then
            '修改文件定义
            gstrSQL = "Select a.ID, a.种类, a.编号, a.名称, a.说明, a.页面, a.保留, a.通用, b.名称 As 页面名称, b.报表, b.格式, b.页眉, b.页脚 " & _
                " From 病历文件列表 a, 病历页面格式 b " & _
                " Where a.页面 = b.编号 And a.种类 = b.种类 And a.Id = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngFileID)
            With EPRFileInfo
                .ID = lngFileID
                If Not rs.EOF Then
                    .种类 = NVL(rs("种类"), 1)  '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明报告;6-知情文件;7-诊疗报告;8-诊疗申请
                    .编号 = NVL(rs("编号"), 0)
                    .名称 = NVL(rs("名称"))
                    .说明 = NVL(rs("说明"))
                    .页面 = NVL(rs("页面"))
                    .保留 = NVL(rs("保留"), 0)
                    .通用 = NVL(rs("通用"), 0)  '除诊疗单据外有效:0-未用;1-通用;2-专科
                    .报表 = NVL(rs("报表"), 0)
                    .页面名称 = NVL(rs("页面名称"))
                    .格式 = NVL(rs("格式"))
                    .页眉 = NVL(rs("页眉"))
                    .页脚 = NVL(rs("页脚"))
                End If
            End With
        End If
    Case cprET_全文示范编辑
        If eEdtMode = cprEM_修改 Then
            '修改范文
            gstrSQL = "Select ID, 文件ID, 编号, 名称, 性质, 说明, 通用级, 科室ID, 人员ID From 病历范文目录 where ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngFileID)  '此时的lngFileId表示文件示范的ID
            With EPRDemoInfo
                .ID = lngFileID
                If Not rs.EOF Then
                    .文件ID = NVL(rs("文件ID"), 0)
                    .编号 = NVL(rs("编号"), 0)
                    .名称 = NVL(rs("名称"))
                    .性质 = NVL(rs("性质"), 0)
                    .通用级 = NVL(rs("通用级"), 0)
                    .科室ID = NVL(rs("科室ID"), 0)
                    .人员ID = NVL(rs("人员ID"), 0)
                End If
            End With
            rs.Close
            gstrSQL = "Select a.ID, a.种类, a.编号, a.名称, a.说明, a.页面, a.保留, a.通用, b.名称 As 页面名称, b.报表, b.格式, b.页眉, b.页脚 " & _
                " From 病历文件列表 a, 病历页面格式 b " & _
                " Where a.页面 = b.编号 And a.种类 = b.种类 And a.Id = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", EPRDemoInfo.文件ID)
            With EPRFileInfo
                If Not rs.EOF Then
                    .ID = NVL(rs("ID"), 0)
                    .保留 = NVL(rs("保留"), 0)
                    .编号 = NVL(rs("编号"), 0)
                    .名称 = NVL(rs("名称"))
                    .说明 = NVL(rs("说明"))
                    .通用 = NVL(rs("通用"), 0)
                    .页面 = NVL(rs("页面"))
                    .种类 = NVL(rs("种类"), 1)
                    .报表 = NVL(rs("报表"), 0)
                    .页面名称 = NVL(rs("页面名称"))
                    .格式 = NVL(rs("格式"))
                    .页眉 = NVL(rs("页眉"))
                    .页脚 = NVL(rs("页脚"))
                End If
            End With
        Else
            '新增范文
            gstrSQL = "Select a.ID, a.种类, a.编号, a.名称, a.说明, a.页面, a.保留, a.通用, b.名称 As 页面名称, b.报表, b.格式, b.页眉, b.页脚 " & _
                " From 病历文件列表 a, 病历页面格式 b " & _
                " Where a.页面 = b.编号 And a.种类 = b.种类 And a.Id = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngFileID)  '此时的lngFileId表示文件定义的ID
            With EPRFileInfo
                .ID = lngFileID
                If Not rs.EOF Then
                    .保留 = NVL(rs("保留"), 0)
                    .编号 = NVL(rs("编号"), 0)
                    .名称 = NVL(rs("名称"))
                    .说明 = NVL(rs("说明"))
                    .通用 = NVL(rs("通用"), 0)
                    .页面 = NVL(rs("页面"))
                    .种类 = NVL(rs("种类"), 1)
                    .报表 = NVL(rs("报表"), 0)
                    .页面名称 = NVL(rs("页面名称"))
                    .格式 = NVL(rs("格式"))
                    .页眉 = NVL(rs("页眉"))
                    .页脚 = NVL(rs("页脚"))
                End If
           End With
        End If
    Case cprET_单病历编辑, cprET_单病历审核
        If eEdtMode = cprEM_修改 Or EditType = cprET_单病历审核 Then
            gstrSQL = "Select ID, 病人来源, 病人ID, 主页ID, 婴儿, 科室ID, 病历种类, 文件ID, " & _
                " 病历名称, 归档人, 归档日期, 保存人, 保存时间, 创建人, 创建时间, " & _
                " 最后版本 , 签名级别 , 完成时间 ,编辑方式 From 电子病历记录 Where ID=[1] "
            If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngFileID)  '此时的lngFileId表示病人电子病历的ID
            With EPRPatiRecInfo
                .ID = lngFileID
                If Not rs.EOF Then
                    .病人来源 = NVL(rs("病人来源"), 1)
                    .病人ID = NVL(rs("病人ID"), 0)
                    .主页ID = NVL(rs("主页ID"), 0)
                    .婴儿 = NVL(rs("婴儿"), 0)
                    .科室ID = NVL(rs("科室ID"), 0)
                    .病历种类 = NVL(rs("病历种类"), 0)
                    .文件ID = NVL(rs("文件ID"), 0)
                    .病历名称 = NVL(rs("病历名称"))
                    .创建人 = NVL(rs("创建人"))
                    .创建时间 = NVL(rs("创建时间"), 0)
                    .保存人 = NVL(rs("保存人"))
                    .保存时间 = NVL(rs("保存时间"), 0)
                    .归档人 = NVL(rs("归档人"))
                    .归档日期 = NVL(rs("归档日期"), 0)
                    .完成时间 = NVL(rs("完成时间"), 0)
                    .最后版本 = NVL(rs("最后版本"), 1)
                    .编辑方式 = NVL(rs("编辑方式"), 0)
                    lngTMP = NVL(rs("签名级别"), 0)     '记录本次编辑的初始版本状态
                    If lngTMP >= 4 Then
                        .签名级别 = cprSL_主任
                    ElseIf lngTMP >= 2 Then
                        .签名级别 = cprSL_主治
                    ElseIf lngTMP >= 1 Then
                        .签名级别 = cprSL_经治
                    Else
                        .签名级别 = cprSL_空白
                    End If
                End If
            End With
            rs.Close
            
            gstrSQL = "Select a.ID, a.种类, a.编号, a.名称, a.说明, a.页面, a.保留, a.通用, b.名称 As 页面名称, b.报表, b.格式, b.页眉, b.页脚 " & _
                " From 病历文件列表 a, 病历页面格式 b " & _
                " Where a.页面 = b.编号 And a.种类 = b.种类 And a.Id = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", EPRPatiRecInfo.文件ID)
            With EPRFileInfo
                If Not rs.EOF Then
                    .ID = NVL(rs("ID"), 0)
                    .保留 = NVL(rs("保留"), 0)
                    .编号 = NVL(rs("编号"), 0)
                    .名称 = NVL(rs("名称"))
                    .说明 = NVL(rs("说明"))
                    .通用 = NVL(rs("通用"), 0)
                    .页面 = NVL(rs("页面"))
                    .种类 = NVL(rs("种类"), 1)
                    .报表 = NVL(rs("报表"), 0)
                    .页面名称 = NVL(rs("页面名称"))
                    .格式 = NVL(rs("格式"))
                    .页眉 = NVL(rs("页眉"))
                    .页脚 = NVL(rs("页脚"))
                End If
            End With
        Else
            '新增电子病历（第1版本）
            gstrSQL = "Select a.ID, a.种类, a.编号, a.名称, a.说明, a.页面, a.保留, a.通用, b.名称 As 页面名称, b.报表, b.格式, b.页眉, b.页脚 " & _
                " From 病历文件列表 a, 病历页面格式 b " & _
                " Where a.页面 = b.编号 And a.种类 = b.种类 And a.Id = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngFileID)  '此时的lngFileId表示文件定义的ID
            With EPRFileInfo
                .ID = lngFileID
                If Not rs.EOF Then
                    .保留 = NVL(rs("保留"), 0)
                    .编号 = NVL(rs("编号"), 0)
                    .名称 = NVL(rs("名称"))
                    .说明 = NVL(rs("说明"))
                    .通用 = NVL(rs("通用"), 0)
                    .页面 = NVL(rs("页面"))
                    .种类 = NVL(rs("种类"), 1)
                    .报表 = NVL(rs("报表"), 0)
                    .页面名称 = NVL(rs("页面名称"))
                    .格式 = NVL(rs("格式"))
                    .页眉 = NVL(rs("页眉"))
                    .页脚 = NVL(rs("页脚"))
                End If
           End With
            EPRPatiRecInfo.科室ID = intCurDeptID
            With EPRPatiRecInfo
                .ID = 0                         '新增记录，ID为0
                .文件ID = EPRFileInfo.ID        '文件ID
                .最后版本 = 1                   '最后版本
                .签名级别 = cprSL_空白          '当前签名级别为空白(0)，表示未签名
                .保存人 = gstrUserName          '保存人
                .病历名称 = EPRFileInfo.名称    '病历名称
                .病历种类 = EPRFileInfo.种类    '病历种类
                .创建时间 = zlDatabase.Currentdate
                
                '直接赋值病人信息
                .病人来源 = PatiFrom
                .病人ID = lngPatiID
                .主页ID = intTimesNo
                .婴儿 = intBabyOrder
            End With
        End If
        
        '对于护理病历，只有2级签名（护士、护士长）
        If Me.EPRPatiRecInfo.病历种类 = cpr护理病历 And Me.用户签名级别 = cprSL_主治 Then Me.用户签名级别 = cprSL_主任
    End Select
    rs.Close
    Set rs = Nothing
    Exit Sub
LL:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'################################################################################################################
'## 功能：  显示主编辑器窗体
'##
'## 参数：  frmParent       :父窗体
'##         blnNoBorder     :是否没有边框
'##         blnCanPrint     :是否允许预览、打印
'################################################################################################################
Public Sub ShowEPREditor(frmParent As Object, Optional ByVal blnCanPrint As Boolean = True, Optional ByVal intStyle As Integer)
'    If Me.EditType = cprET_单病历审核 And Me.用户签名级别 < Me.EPRPatiRecInfo.签名级别 Then
'        '如果在审核模式下，当前用户的权限小于当前文档的签名级别，那么不允许编辑！
'        DoEvents
'        MsgBox "对不起，您没有此病历的审核权限！", vbOKOnly + vbInformation, gstrSysName
'        Exit Sub
'    End If
    OpenEPRDoc frmEditor.Editor1    '打开病历文件
    Set mfrmParent = frmParent
    
    Set frmEditor.Document = Me     '设置关联
    frmEditor.RefreshPatiInfo       '刷新病人信息栏
    frmEditor.Editor1.ReadOnly = (Me.EditType = cprET_单病历编辑 And Me.Signs.Count > 0)
    frmEditor.ShowMe frmParent, True, blnCanPrint, intStyle '显示编辑器
    If intStyle = -1 Then
        frmEditor.ChildMode = True
    End If
End Sub

'################################################################################################################
'## 功能：  导入指定ID的历史病历
'##
'## 参数：  edtThis     :RTB编辑器控件
'##         lngEPRID    :历史病历ID
'##
'## 返回：  打开成功，返回Ture；否则返回False。
'################################################################################################################
Public Function ImportOldEPRFile(ByRef edtThis As Object, ByVal lngEPRid As Long, Optional blnNotAsk As Boolean) As Boolean
    If blnNotAsk = False Then
        If MsgBox("注意：导入该历史文件将丢失已编辑的所有内容，是否确定导入？", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
            ImportOldEPRFile = False: Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo LL
    Dim eEdtMode As EditModeEnum, eEdtType As EditTypeEnum, lngFileID As Long, PatiFrom As PatiFromEnum
    Dim lngPatiID As Long, intTimesNo As Long, intBabyOrder As Integer, intCurDeptID As Long, lngOrderId As Long
    Dim rs As New ADODB.Recordset, lKey As Long, i As Long
    
    '保存历史属性
    eEdtMode = Me.EditMode
    eEdtType = Me.EditType
    If eEdtType = cprET_全文示范编辑 Then
        lngFileID = Me.EPRDemoInfo.ID
    Else
        lngFileID = IIf(Me.EditMode = cprEM_新增, EPRFileInfo.ID, Me.EPRPatiRecInfo.ID)
    End If
    PatiFrom = EPRPatiRecInfo.病人来源
    lngPatiID = EPRPatiRecInfo.病人ID
    intTimesNo = EPRPatiRecInfo.主页ID
    intBabyOrder = EPRPatiRecInfo.婴儿
    intCurDeptID = EPRPatiRecInfo.科室ID
    lngOrderId = EPRPatiRecInfo.医嘱id
    gstrSQL = "Select 病人来源, 病人ID, 主页ID, 婴儿, 科室ID From 电子病历记录 where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ImportOldEPRFile", lngEPRid)
    If rs.EOF Then
        '该条记录不存在
        ImportOldEPRFile = False
        Exit Function
    Else
        '重新初始化病历（注意最后要恢复状态）
        Call Me.InitEPRDoc(cprEM_修改, cprET_单病历审核, lngEPRid, NVL(rs("病人来源"), 1), NVL(rs("病人ID"), 0), _
            NVL(rs("主页ID"), 0), NVL(rs("婴儿"), 0), NVL(rs("科室ID"), 0), 1)
    End If
    rs.Close

    zlCommFun.ShowFlash "正在导入历史病历，请稍候..."
    Screen.MousePointer = vbHourglass
    edtThis.Tag = "禁止同步"
    edtThis.InProcessing = True
    edtThis.Freeze  '冻结
    edtThis.NewDoc  '清空
    
    '重新初始化各个对象集合
    Set Me.Compends = New cEPRCompends
    Set Me.Pictures = New cEPRPictures
    Set Me.Tables = New cEPRTables
    Set Me.Elements = New cEPRElements
    Set Me.Diagnosises = New cEPRDiagnosises
    Set Me.Signs = New cEPRSigns
    
    rs.CursorLocation = adUseClient
    gstrSQL = "select Count(ID) as 数目 from 电子病历内容 where 文件ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngEPRid)
    If rs("数目") = 0 Then
        '电子病历内容为空
        ImportOldEPRFile = False
        Exit Function
    Else
        '读取文件内容结构
        Call ReadFileStructure(edtThis)
        '必须清除所有原始ID，非常重要！！！！！
        Call ClearAllIDs
    End If
    
    edtThis.Title = Me.EPRFileInfo.名称
    '刷新页面格式
    Me.EPRFileInfo.SetFormat edtThis, Me.EPRFileInfo.格式
    Me.EPRFileInfo.ReadPageHead edtThis
    Me.EPRFileInfo.ReadPageFoot edtThis
    
    '清洁模式
    With edtThis
        .ForceEdit = True
        If Me.EditType = cprET_单病历审核 Then
            .SelectAll
            .AcceptAuditText    '清洁文本
        End If
        '清除所有的签名对象
        For i = Me.Signs.Count To 1 Step -1
            lKey = Me.Signs(i).Key
            Me.Signs(i).DeleteFromEditor edtThis
            Me.Signs.Remove "K" & lKey
        Next
        '修改所有对象的版本信息
        For i = Me.Elements.Count To 1 Step -1
            If Me.Elements(i).终止版 > 0 Then
                lKey = Me.Elements(i).Key
                Me.Elements(i).DeleteFromEditor edtThis
                Me.Elements.Remove "K" & lKey
            Else
                Me.Elements(i).开始版 = 1
            End If
        Next
        For i = Me.Diagnosises.Count To 1 Step -1
            If Me.Diagnosises(i).终止版 > 0 Then
                lKey = Me.Diagnosises(i).Key
                Me.Diagnosises(i).DeleteFromEditor edtThis
                Me.Diagnosises.Remove "K" & lKey
            Else
                Me.Diagnosises(i).开始版 = 1
            End If
        Next
        .ForceEdit = False
    End With
    
    '恢复属性
    If eEdtType = cprET_全文示范编辑 Then
        Call Me.InitEPRDoc(eEdtMode, eEdtType, Me.EPRDemoInfo.ID)
    Else
        Call Me.InitEPRDoc(eEdtMode, eEdtType, lngFileID, PatiFrom, lngPatiID, intTimesNo, intBabyOrder, intCurDeptID, lngOrderId)
    End If
    
    If Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核 Then '在病历编辑和审核状态，需要进行页眉页脚的替换处理
        Call Me.GetReplacedHeadFootString(edtThis)
    End If
    
    '重新执行替换
    Call AfterOpeningRTFText(edtThis, True)
    
    ImportOldEPRFile = True
    Screen.MousePointer = 0
    zlCommFun.StopFlash
    edtThis.UnFreeze
    edtThis.Filename = ""
    edtThis.Tag = ""
    edtThis.Modified = True
    edtThis.InProcessing = False
    
    Exit Function
LL:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ImportOldEPRFile = False
    Screen.MousePointer = 0
    zlCommFun.StopFlash
    edtThis.Tag = ""
    edtThis.InProcessing = False
End Function

'################################################################################################################
'## 功能：  导入指定ID的病历示范
'##
'## 参数：  edtThis     : RTB编辑器控件
'##         lngDemoID   :示范ID
'##
'## 返回：  打开成功，返回Ture；否则返回False。
'################################################################################################################
Public Function ImportEPRDemo(ByRef edtThis As Object, ByVal lngDemoId As Long, Optional blnNotAsk As Boolean) As Boolean
    If blnNotAsk = False Then
        If MsgBox("注意：导入全文示范将丢失已经编辑的所有内容，是否确定导入？", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
            ImportEPRDemo = False: Exit Function
        End If
    End If
        
    Err = 0: On Error GoTo LL
    zlCommFun.ShowFlash "正在导入全文示范，请稍候..."
    Screen.MousePointer = vbHourglass
    edtThis.Tag = "禁止同步"
    edtThis.InProcessing = True
    edtThis.Freeze  '冻结
    edtThis.NewDoc  '清空
    
    '重新初始化各个对象集合
    Set Me.Compends = New cEPRCompends
    Set Me.Pictures = New cEPRPictures
    Set Me.Tables = New cEPRTables
    Set Me.Elements = New cEPRElements
    Set Me.Diagnosises = New cEPRDiagnosises
    Set Me.Signs = New cEPRSigns
    
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    gstrSQL = "select Count(ID) as 数目 from 病历范文内容 where 文件ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngDemoId)
    If rs("数目") = 0 Then
        '如果范文内容为空，那么读取其文件定义的内容
        Call ReadInitFileStructure(edtThis)
        '必须清除所有原始ID，非常重要！！！！！
        Call ClearAllIDs
    Else
        '读取范文的内容结构
        Call ReadEPRDemoStructure(edtThis, lngDemoId)
        '必须清除所有原始ID，非常重要！！！！！
        Call ClearAllIDs
    End If
    
    edtThis.Title = Me.EPRFileInfo.名称
    '设置页面格式
    Me.EPRFileInfo.SetFormat edtThis, Me.EPRFileInfo.格式
    Me.EPRFileInfo.ReadPageHead edtThis
    Me.EPRFileInfo.ReadPageFoot edtThis
    If Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核 Then Call Me.GetReplacedHeadFootString(edtThis)
    
    ImportEPRDemo = True
    Screen.MousePointer = 0
    zlCommFun.StopFlash
    edtThis.UnFreeze
    edtThis.Filename = ""
    edtThis.Tag = ""
    edtThis.Modified = True
    edtThis.InProcessing = False
    Exit Function
LL:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ImportEPRDemo = False
    Screen.MousePointer = 0
    zlCommFun.StopFlash
    edtThis.Tag = ""
    edtThis.InProcessing = False
End Function

'################################################################################################################
'## 功能：  打开并显示当前的病历文件
'##
'## 参数：  edtThis: RTB编辑器控件
'##
'## 返回：  打开成功，返回Ture；否则返回False。特殊的，如果当前为修改模式，而数据库中内容为空，则返回False
'##
'## 说明：  如果RTFText内容存在，则直接显示；否则，需要通过数据库内容重构原始病历文件（丢失格式）。
'################################################################################################################
Public Function OpenEPRDoc(ByRef edtThis As Object, Optional ByVal blnMoved As Boolean) As Boolean
'blnMoved 病人数据已转储，需要查转储后备表,只在frmPrintPreview调用时可能会为真
    Dim strEleValue As String

    Err = 0: On Error GoTo LL
    
    edtThis.Tag = "禁止同步"
    edtThis.InProcessing = True
    
    '重新初始化各个对象集合
    Set Me.Compends = New cEPRCompends
    Set Me.Pictures = New cEPRPictures
    Set Me.Tables = New cEPRTables
    Set Me.Elements = New cEPRElements
    Set Me.Diagnosises = New cEPRDiagnosises
    Set Me.Signs = New cEPRSigns
    
    '设置页面格式
    Me.EPRFileInfo.SetFormat edtThis, Me.EPRFileInfo.格式
    Me.EPRFileInfo.ReadPageHead edtThis
    Me.EPRFileInfo.ReadPageFoot edtThis
    edtThis.ResetWYSIWYG    '刷新所见即所得（WYSIWYG）显示
    If EditMode <> cprEM_新增 Or EditType <> cprET_病历文件定义 Then
        Set edtThis.Picture = EPRFileInfo.ReadPagePicture
    End If
    
    Err = 0: On Error GoTo LL
    If EditMode = cprEM_新增 Then
        '新增
        Select Case EditType
        Case cprET_病历文件定义
            Call ReadInitFileStructure(edtThis)
        Case cprET_全文示范编辑
            Call ReadInitFileStructure(edtThis)
        Case cprET_单病历编辑
            Call ReadInitFileStructure(edtThis)
        Case cprET_单病历审核
            Call ReadFileStructure(edtThis)
        End Select
        '必须清除所有原始ID，非常重要！！！！！
        Call ClearAllIDs
        edtThis.Title = Me.EPRFileInfo.名称
        '设置页面格式
        If Me.EditType <> cprET_病历文件定义 Then Call Me.GetReplacedHeadFootString(edtThis)
    Else
        '修改
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        Select Case EditType
        Case cprET_病历文件定义
            gstrSQL = "select ID from 病历文件结构 where 文件ID=[1] and Rownum=1"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", Me.EPRFileInfo.ID)
            If rs.EOF Then
                '修改病历文件定义时，如果内容为空，不处理，供继续新增
            Else
                '否则读取文件内容结构
                Call ReadInitFileStructure(edtThis)
            End If
            edtThis.Title = Me.EPRFileInfo.名称
        Case cprET_全文示范编辑
            gstrSQL = "select ID from 病历范文内容 where 文件ID=[1] and Rownum=1"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", Me.EPRDemoInfo.ID)
            If rs.EOF Then
                '如果范文内容为空，那么读取其原始文件定义的内容
                Call ReadInitFileStructure(edtThis)
                '注意必须清除ID
                Call ClearAllIDs
            Else
                '否则读取范文内容结构
                Call ReadEPRDemoStructure(edtThis, EPRDemoInfo.ID)
            End If
            edtThis.Title = Me.EPRFileInfo.名称
        Case cprET_单病历编辑, cprET_单病历审核
            gstrSQL = "select ID from 电子病历内容 where 文件ID=[1] and Rownum=1"
            If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", Me.EPRPatiRecInfo.ID)
            If rs.EOF Then
                '如果电子病历内容为空，那么读取其原始文件定义的内容
                Call ReadInitFileStructure(edtThis)
                '注意清除ID
                Call ClearAllIDs
            Else
                '否则读取电子病历内容结构
                Call ReadFileStructure(edtThis, blnMoved)
            End If
            edtThis.Title = Me.EPRFileInfo.名称
        End Select
        If Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核 Then '在病历编辑和审核状态，需要进行页眉页脚的替换处理
            Me.GetReplacedHeadFootString edtThis
        Else                                '否则，只进行病历名称的替换
            strEleValue = Me.EPRFileInfo.名称
            Call edtThis.DocHeadReplaceKey("{病历名称}", strEleValue)
            Call edtThis.DocFootReplaceKey("{病历名称}", strEleValue)
        End If
        If Me.EditType = cprET_单病历编辑 And Me.Signs.Count > 0 Then edtThis.ReadOnly = True
    End If
    
    OpenEPRDoc = True
    Screen.MousePointer = 0
    zlCommFun.StopFlash
    edtThis.Filename = ""
    edtThis.Tag = ""
    edtThis.InProcessing = False
    edtThis.Modified = False
    Call frmEditor.ClearUndoList
    If Me.EPRPatiRecInfo.病历种类 = cpr诊疗报告 Then
        Call frmEditor.ClsPacsPic
    End If
    Exit Function
LL:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    OpenEPRDoc = False
    Screen.MousePointer = 0
    zlCommFun.StopFlash
    edtThis.Tag = ""
    edtThis.InProcessing = False
End Function

'################################################################################################################
'## 功能：  清除所有对象的ID号
'################################################################################################################
Public Sub ClearAllIDs()
    Dim i As Long, j As Long

    Set mTextIDCol = New Collection
	mIntTextIDPos = 0

    For i = 1 To Me.Compends.Count
        Me.Compends(i).ID = 0
    Next
    For i = 1 To Me.Pictures.Count
        Me.Pictures(i).ID = 0
    Next
    For i = 1 To Me.Elements.Count
        Me.Elements(i).ID = 0
    Next
    For i = 1 To Me.Signs.Count
        Me.Signs(i).ID = 0
    Next
    For i = 1 To Me.Diagnosises.Count
        Me.Diagnosises(i).ID = 0
    Next
    For i = 1 To Me.Tables.Count
        Me.Tables(i).ID = 0
        For j = 1 To Me.Tables(i).Cells.Count
            Me.Tables(i).Cells(j).ID = 0
        Next
        For j = 1 To Me.Tables(i).Elements.Count
            Me.Tables(i).Elements(j).ID = 0
        Next
        For j = 1 To Me.Tables(i).Pictures.Count
            Me.Tables(i).Pictures(j).ID = 0
        Next
    Next
End Sub

'################################################################################################################
'## 功能：  读取初始文件结构（从“病历文件结构”表中读取，用于新增病人病历时的内容模板初始化）
'##
'## 参数：  edtThis: RTB编辑器控件
'##
'## 返回：  打开成功，返回Ture；否则返回False。特殊的，如果当前为修改模式，而数据库中内容为空，则返回False
'################################################################################################################
Public Function ReadInitFileStructure(ByRef edtThis As Object) As Boolean
    Dim lngKey As Long, lngStart As Long, lngLen As Long, lngLen2 As Long, strTmp As String, strTmp2 As String
    Dim rs As New ADODB.Recordset
    Dim blnHaveRTF As Boolean       '是否存在RTF（文件格式）
    Dim bln保护文本 As Boolean
    
    rs.CursorLocation = adUseClient
    
    '直接恢复RTF文本
    Dim strTemp As String, strZipFile As String
    edtThis.Tag = "ReadInitFileStructure"
    edtThis.Freeze
    If Me.EditType = cprET_全文示范编辑 And Me.EPRDemoInfo.性质 = 1 Then
        blnHaveRTF = False
    Else
        strZipFile = zlBlobRead(1, EPRFileInfo.ID)
        If gobjFSO.FileExists(strZipFile) Then
            strTemp = zlFileUnzip(strZipFile)
            If gobjFSO.FileExists(strTemp) Then
                blnHaveRTF = True
                edtThis.OpenDoc strTemp
                gobjFSO.DeleteFile strTemp, True
            Else
                blnHaveRTF = False
            End If
            gobjFSO.DeleteFile strZipFile, True
        End If
    End If

    Err = 0: On Error GoTo errHand
    '读取元素集合
    gstrSQL = "Select Level, ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id,定义提纲ID, 复用提纲, 使用时机," & vbNewLine & _
                "       诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "From (Select ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行,预制提纲id,ID 定义提纲ID,复用提纲,使用时机," & vbNewLine & _
                "              诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "       From 病历文件结构" & vbNewLine & _
                "       Where 文件id = [1] And 对象序号 > 0)" & vbNewLine & _
                "Start With 父id Is Null" & vbNewLine & _
                "Connect By Prior ID = 父id" & vbNewLine & _
                "Order By 对象序号, 内容行次"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", EPRFileInfo.ID)

    If Not rs.EOF Then
        With edtThis
            .ForceEdit = True
            Do While Not rs.EOF
                Select Case NVL(rs("对象类型"), 2)
                    Case 1  '提纲
                        lngKey = Me.Compends.Add(NVL(rs("对象标记"), 0)) '恢复Key值！
                        Call Me.Compends("K" & lngKey).FillCompendMember(rs, "病历文件结构")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Compends("K" & lngKey).InsertIntoEditor edtThis, lngStart, False, Me
                        End If
                    Case 2  '文本
                        If blnHaveRTF = False Then
                            '恢复RTF内容
                            lngStart = Len(.Text)
                            strTmp = NVL(rs("内容文本")) + IIf(NVL(rs("是否换行"), 0) = 0, "", vbCrLf)
                            lngLen = Len(strTmp)
                            
                            .Range(lngStart, lngStart) = strTmp
                            SetCommonStyle edtThis, "正文", lngStart, lngStart + lngLen, True
                            
                            If Me.EditType = cprET_全文示范编辑 And Me.EPRDemoInfo.性质 = 1 Then
                                bln保护文本 = False
                            Else
                                bln保护文本 = IIf(NVL(rs("对象属性"), 0) = 0, False, True)
                            End If
                            If bln保护文本 Then
                                '自定义保护文本
                                .Range(lngStart, lngStart + lngLen).Font.Strikethrough = False
                                .Range(lngStart, lngStart + lngLen).Font.ForeColor = PROTECT_FORECOLOR
                                .Range(lngStart, lngStart + lngLen).Font.Protected = True
                            Else
                                '已删除文本
                                .Range(lngStart, lngStart + lngLen).Font.Strikethrough = False
                                .Range(lngStart, lngStart + lngLen).Font.ForeColor = tomAutoColor
                                .Range(lngStart, lngStart + lngLen).Font.Protected = False
                            End If
                            .Range(lngStart, lngStart + lngLen).Font.Hidden = False
                            .Range(lngStart + lngLen, lngStart + lngLen).Font.ForeColor = tomAutoColor
                            .Range(lngStart + lngLen, lngStart + lngLen).Font.Strikethrough = False
                        End If
                    Case 3  '表格
                        lngKey = Me.Tables.Add(NVL(rs("对象标记"), 0))                  '恢复Key值！
                        Call Me.Tables("K" & lngKey).FillTableMember(rs, "病历文件结构")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Tables("K" & lngKey).InsertIntoEditor edtThis, lngStart, , False
                        End If
                    Case 4  '要素
                        lngKey = Me.Elements.Add(NVL(rs("对象标记"), 0))
                        Call Me.Elements("K" & lngKey).FillElementMember(rs, "病历文件结构")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Elements("K" & lngKey).InsertIntoEditor edtThis, lngStart
                        End If
                    Case 5  '图片
                        lngKey = Me.Pictures.Add(NVL(rs("对象标记"), 0))
                        Call Me.Pictures("K" & lngKey).FillPictureMember(rs, "病历文件结构")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Pictures("K" & lngKey).InsertIntoEditor edtThis, lngStart, True
                        End If
                    Case 6  '标记
                    Case 7  '诊断
                        lngKey = Me.Diagnosises.Add(NVL(rs("对象标记"), 0))
                        Call Me.Diagnosises("K" & lngKey).FillDiagnosisMember(rs, "病历文件结构")

                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Diagnosises("K" & lngKey).InsertIntoEditor edtThis, lngStart
                        End If
                    Case 8  '签名
                        '病历文件初始调入时，不允许有签名
                End Select
                rs.MoveNext
            Loop
         
            '恢复RTF文本中的图片
            Call AfterOpeningRTFText(edtThis)
            
            .Range(0, 0).Selected
            .ForceEdit = False
            If Me.EditType <> cprET_病历文件定义 Then
                .Modified = True    '此时可以直接保存
            Else
                .Modified = False
            End If
        End With
        
        edtThis.UnFreeze
        Me.Compends.UpdateParentKeys            '更新父Key
        Me.Compends.FillTree frmEditor.mfrmCompends.Tree
        ReadInitFileStructure = True
    Else
        edtThis.UnFreeze
        ReadInitFileStructure = False
    End If
    edtThis.Tag = ""
    Exit Function
errHand:
    edtThis.UnFreeze
    edtThis.Tag = ""
    If ErrCenter() = 1 Then
        Resume
    End If
    ReadInitFileStructure = False
End Function

'################################################################################################################
'## 功能：  读取文件结构
'##
'## 参数：  edtThis: RTB编辑器控件
'##
'## 返回：  打开成功，返回Ture；否则返回False。特殊的，如果当前为修改模式，而数据库中内容为空，则返回False
'################################################################################################################
Public Function ReadFileStructure(ByRef edtThis As Object, Optional ByVal blnMoved As Boolean) As Boolean
'blnMoved 病人数据已转储，需要查转储后备表,只在frmPrintPreview调用时可能会为真
    On Error GoTo LL
    Dim lngKey As Long, lngStart As Long, lngLen As Long, strTmp As String
    Dim rs As New ADODB.Recordset, strTableName As String
    Dim blnHaveRTF As Boolean
    Dim bln保护文本 As Boolean
    Dim lng开始版 As Long, lng终止版 As Long
    
    rs.CursorLocation = adUseClient
    
    '直接恢复RTF文本
    Dim strTemp As String, strZipFile As String
    edtThis.Tag = "ReadFileStructure"

    strZipFile = zlBlobRead(5, EPRPatiRecInfo.ID, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strTemp) Then
            blnHaveRTF = True
            edtThis.OpenDoc strTemp
            gobjFSO.DeleteFile strTemp, True
        Else
            blnHaveRTF = False
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    
    strTableName = IIf(blnMoved, "H电子病历内容", "电子病历内容")
    gstrSQL = "Select Level, ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 定义提纲id, 复用提纲, 使用时机," & vbNewLine & _
                "       诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "From (Select ID, 文件id, 开始版, 终止版, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 定义提纲id, 复用提纲, 使用时机, 诊治要素id," & vbNewLine & _
                "              替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "       From 电子病历内容" & vbNewLine & _
                "       Where 文件id = [1] And 对象序号 <> ID)" & vbNewLine & _
                "Start With 父id Is Null" & vbNewLine & _
                "Connect By Prior ID = 父id" & vbNewLine & _
                "Order By 对象序号, 内容行次"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", EPRPatiRecInfo.ID)
    '对于表格内子对象，对象序号只有电子病历内容填写的单元格自身ID，病历文件结构和病历范文目录填NULL,只处只提取提纲下的子成员，不提取子成员的子成员
    
    If Not rs.EOF Then
        With edtThis
            .Freeze
            .ForceEdit = True
            Do While Not rs.EOF
                    Select Case NVL(rs("对象类型"), 2)
                    Case 1  '提纲
                        lngKey = Me.Compends.Add(NVL(rs("对象标记"), 0))                  '恢复Key值！
                        Call Me.Compends("K" & lngKey).FillCompendMember(rs, strTableName)
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Compends("K" & lngKey).InsertIntoEditor edtThis, lngStart, False, Me
                        End If
                    Case 2  '文本
                        mTextIDCol.Add rs("ID").Value, "K" & rs!ID
                        If blnHaveRTF = False Then
                            '恢复RTF内容
                            lngStart = Len(.Text)
                            strTmp = NVL(rs("内容文本")) + IIf(NVL(rs("是否换行"), 0) = 0, "", vbCrLf)
                            lngLen = Len(strTmp)
                            
                            .Range(lngStart, lngStart) = strTmp
                            SetCommonStyle edtThis, "正文", lngStart, lngStart + lngLen, True
                            bln保护文本 = IIf(NVL(rs("对象属性"), 0) = 0, False, True)
                            If bln保护文本 Then
                                '自定义保护文本
                                .Range(lngStart, lngStart + lngLen).Font.Strikethrough = False
                                .Range(lngStart, lngStart + lngLen).Font.ForeColor = PROTECT_FORECOLOR
                                .Range(lngStart, lngStart + lngLen).Font.Protected = True
                            Else
                                lng开始版 = NVL(rs("开始版"), 1)
                                lng终止版 = NVL(rs("终止版"), 0)
                                .Range(lngStart, lngStart + lngLen).Font.Strikethrough = (lng终止版 > 0 And lng终止版 <= Me.EPRPatiRecInfo.最后版本)
                                .Range(lngStart, lngStart + lngLen).Font.ForeColor = GetCharColor(lng开始版, lng终止版)
                                .Range(lngStart, lngStart + lngLen).Font.Protected = False
                            End If
                            
                            .Range(lngStart, lngStart + lngLen).Font.Hidden = False
                            .Range(lngStart + lngLen, lngStart + lngLen).Font.ForeColor = tomAutoColor
                            .Range(lngStart + lngLen, lngStart + lngLen).Font.Strikethrough = False
                        End If
                    Case 3  '表格
                        lngKey = Me.Tables.Add(NVL(rs("对象标记"), 0))                  '恢复Key值！
                        Call Me.Tables("K" & lngKey).FillTableMember(rs, strTableName)
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Tables("K" & lngKey).InsertIntoEditor edtThis, lngStart, , False
                        End If
                    Case 4  '要素
                        lngKey = Me.Elements.Add(NVL(rs("对象标记"), 0))
                        Call Me.Elements("K" & lngKey).FillElementMember(rs, strTableName)
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Elements("K" & lngKey).InsertIntoEditor edtThis, lngStart
                        End If
                    Case 5  '图片
                        lngKey = Me.Pictures.Add(NVL(rs("对象标记"), 0))
                        Call Me.Pictures("K" & lngKey).FillPictureMember(rs, strTableName)
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Pictures("K" & lngKey).InsertIntoEditor edtThis, lngStart, True
                        End If
                    Case 6  '标记
                    Case 7  '诊断
                        lngKey = Me.Diagnosises.Add(NVL(rs("对象标记"), 0))
                        Call Me.Diagnosises("K" & lngKey).FillDiagnosisMember(rs, strTableName)
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Diagnosises("K" & lngKey).InsertIntoEditor edtThis, lngStart
                        End If
                    Case 8  '签名
                        '临时写入签名信息
                        lngKey = Me.Signs.Add(NVL(rs("对象标记"), 1))  '签名级别
                        Call Me.Signs("K" & lngKey).FillElementMember(rs, strTableName)
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Signs("K" & lngKey).InsertIntoEditor edtThis, lngStart
                        End If
                    End Select
                rs.MoveNext
            Loop
         
            '恢复RTF文本中的图片（如果是验证签名模式，则不需要替换）
            If KeepRTF = False Then Call AfterOpeningRTFText(edtThis)
                        
            .Range(0, 0).Selected
            .ForceEdit = False
            .UnFreeze
            If Me.EditMode = cprEM_新增 And Me.EditType = cprET_单病历编辑 Then
                .Modified = True    '此时可以直接保存
            Else
                .Modified = False
            End If
        End With
        
        Me.Compends.UpdateParentKeys            '更新父Key
        If Me.IsVerifyMode = False Then
            Me.Compends.FillTree frmEditor.mfrmCompends.Tree
        End If
        ReadFileStructure = True
    Else
        edtThis.UnFreeze
        ReadFileStructure = False
    End If
    edtThis.Tag = ""
    Exit Function
LL:
    edtThis.UnFreeze
    edtThis.Tag = ""
    If ErrCenter() = 1 Then
        Resume
    End If
    ReadFileStructure = False
End Function

'################################################################################################################
'## 功能：  读取全文示范的文件结构
'##
'## 参数：  edtThis         : RTB编辑器控件
'##         lngEPRDemoID    : 全文示范ID
'##
'## 返回：  打开成功，返回Ture；否则返回False。特殊的，如果当前为修改模式，而数据库中内容为空，则返回False
'################################################################################################################
Public Function ReadEPRDemoStructure(ByRef edtThis As Object, ByVal lngEPRDemoID As Long) As Boolean
    Dim lngKey As Long, lngStart As Long, lngLen As Long, lngLen2 As Long, strTmp As String, strTmp2 As String
    Dim rs As New ADODB.Recordset
    Dim blnHaveRTF As Boolean, lngFileID As Long
    Dim eEditType As EditTypeEnum
    Dim bln保护文本 As Boolean
    
    eEditType = Me.EditType            '记录当前编辑模式
    Me.EditType = cprET_全文示范编辑
    
    rs.CursorLocation = adUseClient
    edtThis.Tag = "ReadEPRDemoStructure"
    
    '直接恢复RTF文本
    Dim strTemp As String, strZipFile As String
    
    strZipFile = zlBlobRead(3, lngEPRDemoID)
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strTemp) Then
            blnHaveRTF = True
            edtThis.OpenDoc strTemp
            gobjFSO.DeleteFile strTemp, True
        Else
            blnHaveRTF = False
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If

    lngFileID = lngEPRDemoID
    gstrSQL = "Select Level, ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id,定义提纲ID, 复用提纲, 使用时机," & vbNewLine & _
            "       诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
            "From (Select ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行,预制提纲id,定义提纲ID,复用提纲,使用时机," & vbNewLine & _
            "              诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
            "       From 病历范文内容" & vbNewLine & _
            "       Where 文件id = [1] And 对象序号 > 0)" & vbNewLine & _
            "Start With 父id Is Null" & vbNewLine & _
            "Connect By Prior ID = 父id" & vbNewLine & _
            "Order By 对象序号, 内容行次"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngFileID)

    If Not rs.EOF Then
        With edtThis
            .Freeze
            .ForceEdit = True
            Do While Not rs.EOF
                Select Case NVL(rs("对象类型"), 2)
                    Case 1  '提纲
                        lngKey = Me.Compends.Add(NVL(rs("对象标记"), 0))                  '恢复Key值！
                        Call Me.Compends("K" & lngKey).FillCompendMember(rs, "病历范文内容")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Compends("K" & lngKey).InsertIntoEditor edtThis, lngStart, False, Me
                        End If
                    Case 2  '文本
                        If blnHaveRTF = False Then
                            '恢复RTF内容
                            lngStart = Len(.Text)
                            strTmp = NVL(rs("内容文本")) + IIf(NVL(rs("是否换行"), 0) = 0, "", vbCrLf)
                            lngLen = Len(strTmp)
                            
                            .Range(lngStart, lngStart) = strTmp
                            SetCommonStyle edtThis, "正文", lngStart, lngStart + lngLen, True
                            
                            bln保护文本 = IIf(NVL(rs("对象属性"), 0) = 0, False, True)
                            If bln保护文本 Then
                                '自定义保护文本
                                .Range(lngStart, lngStart + lngLen).Font.Strikethrough = False
                                .Range(lngStart, lngStart + lngLen).Font.ForeColor = PROTECT_FORECOLOR
                                .Range(lngStart, lngStart + lngLen).Font.Protected = True
                            Else
                                '已删除文本
                                .Range(lngStart, lngStart + lngLen).Font.Strikethrough = False
                                .Range(lngStart, lngStart + lngLen).Font.ForeColor = tomAutoColor
                                .Range(lngStart, lngStart + lngLen).Font.Protected = False
                            End If
                            .Range(lngStart, lngStart + lngLen).Font.Hidden = False
                            .Range(lngStart + lngLen, lngStart + lngLen).Font.ForeColor = tomAutoColor
                            .Range(lngStart + lngLen, lngStart + lngLen).Font.Strikethrough = False
                        End If
                    Case 3  '表格
                        lngKey = Me.Tables.Add(NVL(rs("对象标记"), 0))                  '恢复Key值！
                        Call Me.Tables("K" & lngKey).FillTableMember(rs, "病历范文内容")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Tables("K" & lngKey).InsertIntoEditor edtThis, lngStart, , False
                        End If
                    Case 4  '要素
                        lngKey = Me.Elements.Add(NVL(rs("对象标记"), 0))
                        Call Me.Elements("K" & lngKey).FillElementMember(rs, "病历范文内容")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Elements("K" & lngKey).InsertIntoEditor edtThis, lngStart
                        End If
                    Case 5  '图片
                        lngKey = Me.Pictures.Add(NVL(rs("对象标记"), 0))
                        Call Me.Pictures("K" & lngKey).FillPictureMember(rs, "病历范文内容")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Pictures("K" & lngKey).InsertIntoEditor edtThis, lngStart, True
                        End If
                    Case 6  '标记
                    Case 7  '诊断
                        lngKey = Me.Diagnosises.Add(NVL(rs("对象标记"), 0))
                        Call Me.Diagnosises("K" & lngKey).FillDiagnosisMember(rs, "病历范文内容")
                        
                        If blnHaveRTF = False Then '用数据库数据刷新RTF控件内容
                            lngStart = Len(.Text)
                            Me.Diagnosises("K" & lngKey).InsertIntoEditor edtThis, lngStart
                        End If
                    Case 8  '签名
                        '范文中也不允许有签名
                End Select
                rs.MoveNext
            Loop
            
            '必须先恢复编辑模式
            Me.EditType = eEditType
            
            '然后恢复RTF文本中的图片，所有替换项目清空！！！
            Call AfterOpeningRTFText(edtThis, True)
                                              
            .Range(0, 0).Selected
            .ForceEdit = False
            .UnFreeze
            .Modified = False
        End With
        
        Me.Compends.UpdateParentKeys            '更新父Key
        Me.Compends.FillTree frmEditor.mfrmCompends.Tree
        ReadEPRDemoStructure = True
    Else
        edtThis.UnFreeze
        ReadEPRDemoStructure = False
    End If
    edtThis.Tag = ""
    Exit Function
LL:
    edtThis.UnFreeze
    edtThis.Tag = ""
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ReadEPRDemoStructure = False
End Function

'################################################################################################################
'## 功能：  更新电子病历记录的“保存时间”和“签名信息”、“最后版本”
'################################################################################################################
Public Sub UpdateDocInfo(ByRef edtThis As Object)
    Dim rs As New ADODB.Recordset, lngTMP As Long
    If Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核 Then
        InitEPRDoc cprEM_修改, Me.EditType, Me.EPRPatiRecInfo.ID, Me.EPRPatiRecInfo.病人来源, _
            Me.EPRPatiRecInfo.病人ID, Me.EPRPatiRecInfo.主页ID, Me.EPRPatiRecInfo.婴儿, _
            Me.EPRPatiRecInfo.科室ID, Me.EPRPatiRecInfo.医嘱id
    End If
End Sub

'################################################################################################################
'## 功能：  保存 RTFText 的预处理
'##
'## 参数：  edtThis     :编辑器控件
'##
'## 说明：  图片、表格内容的清空
'################################################################################################################
Public Function PreSavingRTFText(ByRef edtThis As Object) As Boolean
    On Error GoTo LL
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim i As Long, blnForceEdit As Boolean
    
    blnForceEdit = edtThis.ForceEdit
    With edtThis
        .Tag = "PreSavingRTFText"
'        .Freeze
        .ForceEdit = True
        For i = 1 To Me.Pictures.Count
            bFinded = FindKey(edtThis, "P", Me.Pictures(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then
                .Range(lKSE, lKES) = "□"
            End If
        Next
        For i = 1 To Me.Tables.Count
            bFinded = FindKey(edtThis, "T", Me.Tables(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then
                .Range(lKSE, lKES) = "□"
            End If
        Next
        .ForceEdit = blnForceEdit
        .Tag = ""
'        .UnFreeze
    End With
    PreSavingRTFText = True
    Exit Function
LL:
    PreSavingRTFText = False
End Function

'################################################################################################################
'## 功能：  将要素转化为纯文本
'##
'## 参数：  edtThis     :编辑器
'##         Ele         :诊治要素
'##         RemoveIt    :删除该要素，默认为True（表示要删除！）
'################################################################################################################
Public Sub EleToString(ByRef edtThis As Object, Ele As cEPRElement, Optional RemoveIt As Boolean = True)
    Dim sKeyType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bNeeded As Boolean, bBeteenKeys As Boolean
    Dim bForce As Boolean, strOldTag As String
    bBeteenKeys = FindKey(edtThis, "E", Ele.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bBeteenKeys Then
        Dim lngLen As Long, str内容 As String
        str内容 = Ele.内容文本
        lngLen = Len(str内容)
        With edtThis
            .Freeze
            strOldTag = .Tag
            .Tag = "EleToString"
            bForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = str内容
            .Range(lKSS, lKSS + lngLen).Font.Protected = False
            .Range(lKSS, lKSS + lngLen).Font.Hidden = False
            .Range(lKSS, lKSS + lngLen).Font.BackColor = tomAutoColor
            .Range(lKSS, lKSS + lngLen).Font.Underline = cprNone
            .ForceEdit = bForce
            .UnFreeze
            .Tag = strOldTag
        End With
        If RemoveIt Then Me.Elements.Remove "K" & Ele.Key
    End If
End Sub

'################################################################################################################
'## 功能：  打开 RTFText 的后期处理
'##
'## 参数：  edtThis             :编辑器控件
'##         blnImportEPRDemo    :是否是从全文示范中导入内容
'##
'## 说明：  图片、表格内容的还原
'################################################################################################################
Public Function AfterOpeningRTFText(ByRef edtThis As Object, _
    Optional ByVal blnImportEPRDemo As Boolean = False) As Boolean
    
'    On Error GoTo LL
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim i As Long, j As Long, blnForceEdit As Boolean, p As Long
    Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat
'    Dim frmTMP As New frmPreVerify
    
    blnForceEdit = edtThis.ForceEdit
    With edtThis
'        .Freeze
        .Tag = "AfterOpeningRTFText"
        .ForceEdit = True
        If (Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核 Or Me.EditType = cprET_全文示范编辑) And Me.IsVerifyMode = False Then
            '替换项目的处理
            For i = 1 To Me.Elements.Count
                If Me.Elements(i).替换域 = 1 And (Me.EditMode = cprEM_新增 Or blnImportEPRDemo Or Trim(Me.Elements(i).内容文本) = "") Then
                    Me.Elements(i).内容文本 = GetReplaceEleValue(Me.Elements(i).要素名称, _
                        Me.EPRPatiRecInfo.病人ID, _
                        Me.EPRPatiRecInfo.主页ID, _
                        Me.EPRPatiRecInfo.病人来源, _
                        Me.EPRPatiRecInfo.医嘱id, _
                        Me.EPRPatiRecInfo.婴儿)
                    bFinded = FindKey(edtThis, "E", Me.Elements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
                    Me.Elements(i).Refresh edtThis
                End If
                If Me.Elements(i).替换域 = 1 And Me.Elements(i).自动转文本 Then
                    EleToString edtThis, Me.Elements(i), False      '自动转化为纯文本（暂时不删除该要素）
                End If
            Next
            For i = Me.Elements.Count To 1 Step -1
                If Me.Elements(i).替换域 = 1 And Me.Elements(i).自动转文本 Then
                     Me.Elements.Remove "K" & Me.Elements(i).Key    '实际删除该要素节点
                End If
            Next
            For i = 1 To Me.Tables.Count
                For j = 1 To Me.Tables(i).Elements.Count
                    If Me.Tables(i).Elements(j).替换域 = 1 And (Me.EditMode = cprEM_新增 Or blnImportEPRDemo Or Left(Me.Tables(i).Elements(j).内容文本, 1) = "{") Then
                        Me.Tables(i).Elements(j).内容文本 = GetReplaceEleValue(Me.Tables(i).Elements(j).要素名称, _
                            Me.EPRPatiRecInfo.病人ID, _
                            Me.EPRPatiRecInfo.主页ID, _
                            Me.EPRPatiRecInfo.病人来源, _
                            Me.EPRPatiRecInfo.医嘱id, _
                            Me.EPRPatiRecInfo.婴儿)
                    End If
                Next
                For j = 1 To Me.Tables(i).Cells.Count
                    If Me.Tables(i).Cells(j).ElementKey > 0 Then
                        Me.Tables(i).Cells(j).内容文本 = Me.Tables(i).Elements("K" & Me.Tables(i).Cells(j).ElementKey).内容文本
'                            Me.Tables(i).Cells(j).CellFormat.FontColor = IIf(Me.Tables(i).Cells(j).内容文本 = "", TABLEELE_FORECOLOR, vbBlack)
'                        Me.Tables(i).Cells(j).CellFormat.FontUnderline = IIf(Me.Tables(i).Cells(j).内容文本 = "", True, False)
                    End If
                Next
            Next
        End If
        
        For i = 1 To Me.Pictures.Count
            bFinded = FindKey(edtThis, "P", Me.Pictures(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then
                '还原图片
                If IsVerifyMode = True Then
                    '验证模式，只还原文本
                    .Range(lKSE, lKES) = "?"
                Else
                    Me.Pictures(i).DeleteFromEditor edtThis
                    Me.Pictures(i).InsertIntoEditor edtThis, -1, True
                End If
            End If
        Next
        For i = 1 To Me.Tables.Count
            bFinded = FindKey(edtThis, "T", Me.Tables(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then
                '还原表格
                If IsVerifyMode = True Then
                    '验证模式，只还原文本
                    .Range(lKSE, lKES) = "?"
                Else
                    Set ParaFmt = .Range(lKSE, lKES).Para.GetParaFmt
                    Set FontFmt = .Range(lKSE, lKES).Font.GetFontFmt
                    
                    If Me.Tables(i).是否换行 Then
                        .Range(lKSS, lKEE + 2).Text = ""
                    Else
                        .Range(lKSS, lKEE).Text = ""
                    End If
                    '重新插入
                    Me.Tables(i).InsertIntoEditor edtThis, lKSS, , , True
                    
                    .Range(lKSE, lKES).Para.SetParaFmt ParaFmt
                    .Range(lKSE, lKES).Font.SetFontFmt FontFmt
                    .Range(lKSS, lKEE).Font.Protected = True
                End If
            End If
        Next
        
        If IsVerifyMode = False Then
            If blnImportEPRDemo Then
                '如果是从病历示范中导入，清空所有签名信息
                Set Me.Signs = New cEPRSigns
                For i = 1 To Me.Signs.Count
                    Me.Signs(i).DeleteFromEditor edtThis
                    Me.Signs.Remove "K" & Me.Signs(i).Key
                Next
            End If
        End If
        
        .ForceEdit = blnForceEdit
        .Tag = ""
'        .UnFreeze
    End With
'    Unload frmTMP
'    Set frmTMP = Nothing
    AfterOpeningRTFText = True
    Exit Function
LL:
    edtThis.UnFreeze
    edtThis.Tag = ""
    AfterOpeningRTFText = False
End Function

'################################################################################################################
'## 功能：  获取审核模式下分段文本到数据库的SQL语句
'##
'## 参数：  ArraySQL()      :IN/OUT，   SQL数组
'##         edtThis         :IN，       Editor控件
'##         lStart,lEnd     :IN，       需要保存的文本的起始、终止位置
'##         lng序号         :IN，       序号
'##         lng父ID         :IN，       父ID
'##         bln是否换行     :IN，       是否换行
'##
'## 说明：  审核模式下的保存需要对文本进行分段保存！！！    长度大于4000的字符串，分行存储，序号递增之！
'##         另外，因为审核模式下回车可能有不同状态，因此“是否换行”恒等于False，回车符保存在“内容文本”中
'################################################################################################################
Public Function GetPlainTextSaveSQL(ByRef ArraySQL() As String, _
    ByRef edtThis As Object, _
    ByVal lStart As Long, _
    ByVal lEnd As Long, _
    ByRef lng序号 As Long, _
    ByVal lng父ID As Double, _
    ByVal bln是否换行 As Boolean) As Boolean
    
    Dim i As Long, p As Long, q As Long
    Dim lngState1 As Integer                            '是否是保护文本 cprAS_ProtedText
    Dim lngState2 As Integer                            '是否是保护文本 cprAS_ProtedText
    Dim lng开始版1 As Long, lng终止版1 As Long          '版本
    Dim lng开始版2 As Long, lng终止版2 As Long          '版本
    
    If lStart = lEnd Then GoTo LL
    
    With edtThis
        If lEnd = lStart + 1 Then
            '只选中了一个字符
            lngState1 = GetTextState(edtThis, lStart, lEnd, lng开始版1, lng终止版1)
            GetTextSaveSQL Me.EditType, ArraySQL, .Range(lStart, lEnd), lng序号, lng父ID, bln是否换行, lng开始版1, lng终止版1, IIf(lngState1 = cprAS_ProtedText, 1, 0)
        Else
            p = lStart
            lngState1 = GetTextState(edtThis, p, p + 1, lng开始版1, lng终止版1)
            For i = lStart + 1 To lEnd - 1
                lngState2 = GetTextState(edtThis, i, i + 1, lng开始版2, lng终止版2)
                If lngState1 = lngState2 And lng开始版1 = lng开始版2 And lng终止版1 = lng终止版2 Then
                    If i = lEnd - 1 Then
                        '最末位置判断
                        GetTextSaveSQL Me.EditType, ArraySQL, .Range(p, lEnd), lng序号, lng父ID, bln是否换行, lng开始版1, lng终止版1, IIf(lngState1 = cprAS_ProtedText, 1, 0)
                    End If
                Else
                    '状态不同
                    q = i
                    GetTextSaveSQL Me.EditType, ArraySQL, .Range(p, q), lng序号, lng父ID, False, lng开始版1, lng终止版1, IIf(lngState1 = cprAS_ProtedText, 1, 0)
                    p = i
                    lngState1 = lngState2
                    lng开始版1 = lng开始版2
                    lng终止版1 = lng终止版2
                End If
            Next
        End If
    End With
    
    GetPlainTextSaveSQL = True
    Exit Function
LL:
    GetPlainTextSaveSQL = False
End Function

'################################################################################################################
'## 功能：  返回指定文本段的审阅状态
'##
'## 参数：  edtThis         :IN，       Editor控件
'##         lStart,lEnd     :IN，       需要保存的文本的起始、终止位置
'##
'## 返回：  -1－删除文本；0－普通文本；1－新增文本；2－混合文本；3－手工设置的保护文本
'################################################################################################################
Public Function GetTextState(ByRef edtThis As Object, _
    ByVal lStart As Long, ByVal lEnd As Long, ByRef lng开始版 As Long, lng终止版 As Long) As TextAuditStateEnum
    
    With edtThis
        If .Range(lStart, lEnd).Font.ForeColor = PROTECT_FORECOLOR And .Range(lStart, lEnd).Font.Protected Then
            '手工设置的保护文本
            GetTextState = cprAS_ProtedText
        Else
            lng开始版 = Get开始版(.Range(lStart, lEnd).Font.ForeColor)
            lng终止版 = Get终止版(.Range(lStart, lEnd).Font.ForeColor)
        End If
    End With
End Function

'################################################################################################################
'## 功能：  获取纯文本到数据库的SQL语句
'##
'## 参数：  eEdtType        :IN，       当前编辑模式
'##         ArraySQL()      :IN/OUT，   SQL数组
'##         strIn           :IN，       需要保存的字符串
'##         lng序号         :IN，       序号
'##         lng父ID         :IN，       父ID
'##         bln是否换行     :IN，       是否换行
'##         int文本状态     :IN，       文本的“对象属性”（-1－删除文本；0－普通文本；1－新增文本；2－保护文本）
'##         lng开始版       :IN，       开始版
'##         lng终止版       :IN，       终止版
'##
'## 说明：  长度大于4000的字符串，分行存储，序号递增之！
'################################################################################################################
Public Function GetTextSaveSQL(ByVal eEdtType As EditTypeEnum, _
    ByRef ArraySQL() As String, _
    ByVal strIn As String, _
    ByRef lng序号 As Long, _
    ByVal lng父ID As Double, _
    ByVal bln是否换行 As Boolean, _
    ByVal lng开始版 As Long, _
    ByVal lng终止版 As Long, _
    Optional ByVal int文本状态 As Integer = 0) As Boolean
    
    Dim lngLen As Long, strSub As String, i As Long, lngId As Double, blnCRLF As Boolean
    Dim lngCount As Long, lID As Long
'    strIn = Replace(strIn, "'", "' || chr(39) || '")
'    strIn = Replace(strIn, vbCrLf, "' || chr(13) || chr(10) || '")  '本来strIn是不允许有vbCrlf的。
    lngLen = Len(strIn)
    
    '按照4000为界分段存储。
    i = 0
    Do While (i * 2000 + 1 <= lngLen)
        lngCount = UBound(ArraySQL) + 1
        ReDim Preserve ArraySQL(1 To lngCount) As String

        strSub = Mid(strIn, i * 2000 + 1, 2000)
        
        strSub = Replace(strSub, "'", "' || chr(39) || '")
        strSub = Replace(strSub, vbCrLf, "' || chr(13) || chr(10) || '")  '本来strIn是不允许有vbCrlf的。
    
        If (i + 1) * 2000 > lngLen Then
            blnCRLF = bln是否换行
        Else
            blnCRLF = False
        End If
        
        Select Case eEdtType
        Case cprET_病历文件定义
            lngId = zlDatabase.GetNextId("病历文件结构")
            gstrSQL = "Zl_病历文件结构_Update(" & lngId & "," & Me.EPRFileInfo.ID & "," & IIf(lng父ID = 0, "NULL", lng父ID) & "," & lng序号 & "," & _
                " 2,NULL,NULL," & int文本状态 & ",NULL,'" & strSub & "'," & IIf(blnCRLF, 1, 0) & ")"
        Case cprET_全文示范编辑
            lngId = zlDatabase.GetNextId("病历范文内容")
            gstrSQL = "Zl_病历范文内容_Update(" & lngId & "," & Me.EPRDemoInfo.ID & "," & IIf(lng父ID = 0, "NULL", lng父ID) & "," & lng序号 & "," & _
                " 2,NULL,NULL," & int文本状态 & ",NULL,'" & strSub & "'," & IIf(blnCRLF, 1, 0) & ")"
        Case cprET_单病历编辑
            lngId = GetTextID
            gstrSQL = "ZL_电子病历内容_Update(" & lngId & "," & Me.EPRPatiRecInfo.ID & "," & lng开始版 & "," & lng终止版 & "," & IIf(lng父ID = 0, "NULL", lng父ID) & "," & lng序号 & "," & _
                " 2,NULL,NULL," & int文本状态 & ",NULL,'" & strSub & "'," & IIf(blnCRLF, 1, 0) & ")"
        Case cprET_单病历审核
            lngId = GetTextID
            gstrSQL = "ZL_电子病历内容_Update(" & lngId & "," & Me.EPRPatiRecInfo.ID & "," & lng开始版 & "," & lng终止版 & "," & IIf(lng父ID = 0, "NULL", lng父ID) & "," & lng序号 & "," & _
                " 2,NULL,NULL," & int文本状态 & ",NULL,'" & strSub & "'," & IIf(blnCRLF, 1, 0) & ")"
        End Select
        
        ArraySQL(lngCount) = gstrSQL
       
        lng序号 = lng序号 + 1
        i = i + 1
    Loop
    GetTextSaveSQL = True
End Function
Private Function GetTextID() As Double
Dim dblID As Double
'mIntTextIDPos为ID集合中可用ID位置，当无可用ID时取ORACLE序列，同时集合增加、位置增加
    
    If mIntTextIDPos = 0 Then mIntTextIDPos = 1 '集合从1开始  首次调用时赋值1,保存完毕后清0以备下次保存时可再用
    
    Err.Clear
    On Error Resume Next
    If mTextIDCol.Count > 0 And mIntTextIDPos <= mTextIDCol.Count Then
        dblID = mTextIDCol(mIntTextIDPos)
        mIntTextIDPos = mIntTextIDPos + 1
    Else
        dblID = GetEPRContentNextId
        mTextIDCol.Add dblID, "K" & dblID
        mIntTextIDPos = mIntTextIDPos + 1
    End If
    
    If Err.Number <> 0 Then Err.Clear
    
    If dblID = 0 Then
        dblID = GetEPRContentNextId
        mTextIDCol.Add dblID, "K" & dblID
        mIntTextIDPos = mIntTextIDPos + 1
    End If
    
    GetTextID = dblID
End Function
'################################################################################################################
'## 用途：  根据文章中的指定位置获取当前提纲Key值。
'##
'## 参数：  edtThis     :   RTB编辑器控件
'##         lngStartPos :   当前位置
'##
'## 返回：  返回找到的提纲Key值。
'################################################################################################################
Public Function GetCurCompendKey(ByRef edtThis As Object, Optional ByVal lngStartPos As Long = -1) As Long
    Dim i As Long, j As Long, lKey As Long, lLen As Long
    If lngStartPos = -1 Then lngStartPos = edtThis.SelStart
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = lngStartPos + 1
        lLen = Len(sText)
        If i > lLen And lLen > 0 Then i = lLen
LL1:
        i = InStrRev(sText, "OS(", i, vbTextCompare)    '反向搜索第一个 OS(...) 关键字
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            '查找结束关键字
            j = i + 16
LL2:
            j = InStr(j, sText, "OE(", vbTextCompare)   '正向搜索其匹配的结束关键字 OE(...)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                lKey = Val(.TOM.TextDocument.Range(i + 2, i + 10))
                GetCurCompendKey = lKey
            End If
        End If
    End With
End Function

'################################################################################################################
'## 功能： 根据文章中的指定位置获取相应父ID，并在提纲树中将对应提纲高亮显示。
'##
'## 参数：  edtThis     :编辑器控件
'##         Tree        :树型列表
'################################################################################################################
Public Sub HighlightCurCompend(ByRef edtThis As Object, ByRef Tree As Object)
    Dim i As Long, j As Long, lKey As Long, lLen As Long
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = .SelStart + 3
        lLen = Len(sText)
        If i > lLen And lLen > 0 Then i = lLen
LL1:
        i = InStrRev(sText, "OS(", i, vbTextCompare)    '反向搜索第一个 OS(...) 关键字
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            '查找结束关键字
            j = i + 16
LL2:
            j = InStr(j, sText, "OE(", vbTextCompare)   '正向搜索其匹配的结束关键字 OE(...)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                lKey = Val(.TOM.TextDocument.Range(i + 2, i + 10))
'                Debug.Print lKey
                
                '高亮显示该提纲
                For i = 1 To Tree.Nodes.Count
                    If Tree.Nodes(i).Tag = lKey Then
                        '找到该提纲
                        Tree.Nodes(i).Selected = True
                    End If
                Next
            End If
        End If
    End With
End Sub

Public Sub GetReplacedHeadFootString(ByRef edtThis As Object, Optional ByVal blnReplaceSign As Boolean)
    '******************************************************************************************************************
    '功能： 获取页眉/页脚中的替换要素内容
    '参数： edtThis 编辑控件
    '******************************************************************************************************************
    
    Dim strElements As String
    Dim lngLoop As Long
    Dim aryElements() As String
    Dim strEleValue As String
    Dim lngStartPos As Long
    Dim lngEndPos As Long
    Dim strHead As String, strFoot As String
    
    If Me.EditType <> cprET_单病历编辑 And Me.EditType <> cprET_单病历审核 Then Exit Sub
    '先替换掉与本病历相关的内容部份
    '------------------------------------------------------------------------------------------------------------------
    strEleValue = Me.EPRFileInfo.名称
    Call edtThis.DocHeadReplaceKey("{病历名称}", strEleValue)
    Call edtThis.DocFootReplaceKey("{病历名称}", strEleValue)
    
    strEleValue = Me.EPRPatiRecInfo.书写部门
    Call edtThis.DocHeadReplaceKey("{书写部门}", strEleValue)
    Call edtThis.DocFootReplaceKey("{书写部门}", strEleValue)
    
    If blnReplaceSign Then
        strEleValue = Me.EPRPatiRecInfo.书写签名
        Call edtThis.DocHeadReplaceKey("{书写签名}", strEleValue)
        Call edtThis.DocFootReplaceKey("{书写签名}", strEleValue)
        
        strEleValue = Me.EPRPatiRecInfo.医生签名
        Call edtThis.DocHeadReplaceKey("{医生签名}", strEleValue)
        Call edtThis.DocFootReplaceKey("{医生签名}", strEleValue)
        
        strEleValue = Me.EPRPatiRecInfo.主治签名
        Call edtThis.DocHeadReplaceKey("{主治签名}", strEleValue)
        Call edtThis.DocFootReplaceKey("{主治签名}", strEleValue)
        
        strEleValue = Me.EPRPatiRecInfo.主任签名
        Call edtThis.DocHeadReplaceKey("{主任签名}", strEleValue)
        Call edtThis.DocFootReplaceKey("{主任签名}", strEleValue)
    End If
    
    strEleValue = Format(IIf(Me.EPRPatiRecInfo.完成时间 = 0, "", Me.EPRPatiRecInfo.完成时间), "yyyy-MM-dd hh:mm")
    Call edtThis.DocHeadReplaceKey("{完成时间}", strEleValue)
    Call edtThis.DocFootReplaceKey("{完成时间}", strEleValue)
    
    '从strHead中分析出替换要素，并放到aryElements数组中
    '------------------------------------------------------------------------------------------------------------------
    strHead = edtThis.HeadFileText
    lngStartPos = 0
    lngEndPos = 0
    strElements = ""
    For lngLoop = 1 To Len(strHead)
        If Mid(strHead, lngLoop, 1) = "{" Then lngStartPos = lngLoop
        If Mid(strHead, lngLoop, 1) = "}" Then lngEndPos = lngLoop

        If lngStartPos > 0 And lngEndPos > 0 Then
            If lngEndPos > lngStartPos + 1 Then
                strElements = strElements & ";" & Mid(strHead, lngStartPos + 1, lngEndPos - lngStartPos - 1)
            End If

            lngStartPos = 0
            lngEndPos = 0
        End If
    Next

    '从strFoot中分析出替换要素，并放到aryElements数组中
    '------------------------------------------------------------------------------------------------------------------
    strFoot = edtThis.FootFileText
    lngStartPos = 0
    lngEndPos = 0
    For lngLoop = 1 To Len(strFoot)
        If Mid(strFoot, lngLoop, 1) = "{" Then lngStartPos = lngLoop
        If Mid(strFoot, lngLoop, 1) = "}" Then lngEndPos = lngLoop

        If lngStartPos > 0 And lngEndPos > 0 Then
            If lngEndPos > lngStartPos + 1 Then
                strElements = strElements & ";" & Mid(strFoot, lngStartPos + 1, lngEndPos - lngStartPos - 1)
            End If

            lngStartPos = 0
            lngEndPos = 0
        End If
    Next
    If strElements <> "" Then strElements = Mid(strElements, 2)
    
    
    aryElements = Split(strElements, ";")
    
    '开始替换要素成具体的内容
    '------------------------------------------------------------------------------------------------------------------
    With Me.EPRPatiRecInfo
        For lngLoop = 0 To UBound(aryElements)
            Select Case aryElements(lngLoop)
            Case "页码", "总页数", "标题", "文件名", "路径", "打印日期", "打印时间", "书写部门", "书写签名", "完成时间", "病历名称", "医生签名", "主治签名", "主任签名"
                
            Case Else
                If InStr(strHead, "{" & aryElements(lngLoop) & "}") > 0 Or InStr(strFoot, "{" & aryElements(lngLoop) & "}") > 0 Then
                    strEleValue = GetReplaceEleValue(aryElements(lngLoop), .病人ID, .主页ID, .病人来源, .医嘱id, .婴儿)
                    Call edtThis.DocHeadReplaceKey("{" & aryElements(lngLoop) & "}", strEleValue)
                    Call edtThis.DocFootReplaceKey("{" & aryElements(lngLoop) & "}", strEleValue)
                End If
            End Select
        Next
    End With
End Sub

Private Sub Class_initialize()
    Set Tables = New cEPRTables
    Set Pictures = New cEPRPictures
    Set Compends = New cEPRCompends
    Set Elements = New cEPRElements
    Set Diagnosises = New cEPRDiagnosises
    Set Signs = New cEPRSigns
    
    Set EPRFileInfo = New cEPRFileDefineInfo
    Set EPRDemoInfo = New cEPRDemoInfo
    Set EPRPatiRecInfo = New cEPRPatiRecordInfo
    Set mTextIDCol = New Collection
    
    Set mfrmParent = Nothing
    Set frmEditor = New frmMain
    
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set Tables = Nothing
    Set Pictures = Nothing
    Set Compends = Nothing
    Set Elements = Nothing
    Set Diagnosises = Nothing
    Set Signs = Nothing

    Set EPRFileInfo = Nothing
    Set EPRDemoInfo = Nothing
    Set EPRPatiRecInfo = Nothing
    Set mTextIDCol = Nothing
    
    Set mfrmParent = Nothing
    Unload frmEditor
    Set frmEditor = Nothing
End Sub

'################################################################################################################
'## 功能：  保存当前病历文件到数据库中
'##
'## 参数：  edtThis             : RTB编辑器控件
'##         blnSavePageSetup    : 是否保存非定义状态的页面设置
'##
'## 返回：  保存成功，返回Ture；否则返回False。
'##
'## 说明：  需要临时生成控件ctlEditor，保存edtThis的所有内容，保证替换图片时不影响到edtThis的内容。
'################################################################################################################
Public Function SaveEPRDoc(ByRef edtThis As Object, Optional blnSavePageSetup As Boolean = False) As Boolean
    '保存病历文件到数据库中
Dim rs As New ADODB.Recordset, ArraySQL() As String, blnForce As Boolean, blnReadOnly As Boolean
Dim i As Long, j As Long, blnHaveSign As Boolean, strHead As String, strFoot As String, strFileContent As String
Dim blnSave As Boolean, strZipFile As String, lS As Long, lE As Long, blnTran As Boolean

    Err = 0: On Error GoTo errHand

    If Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核 Then
        Dim strDTOld As String, strDTNew As String, lngFidOld As Long, lngFidNew As Long
        '判断在编辑期间该病历文件是否已经被修改过！如果是，则不允许保存！
        strDTOld = Format(Me.EPRPatiRecInfo.保存时间, "YYYYMMDDHHmmss")
        lngFidOld = Me.EPRPatiRecInfo.文件ID
        gstrSQL = "select 保存时间,文件ID from 电子病历记录 where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", Me.EPRPatiRecInfo.ID)
        If Not rs.EOF Then
            strDTNew = Format(NVL(rs("保存时间"), 0), "YYYYMMDDHHmmss")
            lngFidNew = rs!文件ID '文件ID的比对主要考虑是在病历排序界面中有可能不能文件ID对调功能
            If strDTNew <> strDTOld Or lngFidOld <> lngFidNew Then
                MsgBox "对不起，同一份病历文件在其他地方已经被修改过。" & vbCrLf & _
                    "无法保存该病历文件，需要重新进入！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        rs.Close
    End If
    
    Screen.MousePointer = vbHourglass

    '先更新所有对象相关信息（同步其“文件ID”值）
    Select Case EditType
    Case cprET_病历文件定义
        If Me.EditMode = cprEM_新增 Then Me.EPRPatiRecInfo.ID = zlDatabase.GetNextId("病历文件列表")
        For i = 1 To Me.Pictures.Count: Me.Pictures(i).文件ID = Me.EPRFileInfo.ID: Next
        For i = 1 To Me.Elements.Count: Me.Elements(i).文件ID = Me.EPRFileInfo.ID: Next
        For i = 1 To Me.Tables.Count: Me.Tables(i).文件ID = Me.EPRFileInfo.ID:  Next
        For i = 1 To Me.Compends.Count: Me.Compends(i).文件ID = Me.EPRFileInfo.ID: Next
        For i = 1 To Me.Diagnosises.Count: Me.Diagnosises(i).文件ID = Me.EPRFileInfo.ID: Next
        For i = 1 To Me.Signs.Count:  Me.Signs(i).文件ID = Me.EPRFileInfo.ID: Next
    Case cprET_全文示范编辑
        If Me.EditMode = cprEM_新增 Then Me.EPRPatiRecInfo.ID = zlDatabase.GetNextId("病历范文目录")
        For i = 1 To Me.Pictures.Count: Me.Pictures(i).文件ID = Me.EPRDemoInfo.ID: Next
        For i = 1 To Me.Elements.Count: Me.Elements(i).文件ID = Me.EPRDemoInfo.ID: Next
        For i = 1 To Me.Tables.Count: Me.Tables(i).文件ID = Me.EPRDemoInfo.ID:  Next
        For i = 1 To Me.Compends.Count: Me.Compends(i).文件ID = Me.EPRDemoInfo.ID: Next
        For i = 1 To Me.Diagnosises.Count: Me.Diagnosises(i).文件ID = Me.EPRDemoInfo.ID: Next
        For i = 1 To Me.Signs.Count:  Me.Signs(i).文件ID = Me.EPRDemoInfo.ID: Next
    
    Case cprET_单病历编辑, cprET_单病历审核
        If Me.EditType = cprET_单病历编辑 And Me.EPRPatiRecInfo.ID = 0 Then Me.EPRPatiRecInfo.ID = zlDatabase.GetNextId("电子病历记录")
        For i = 1 To Me.Pictures.Count: Me.Pictures(i).文件ID = Me.EPRPatiRecInfo.ID: Next
        For i = 1 To Me.Elements.Count: Me.Elements(i).文件ID = Me.EPRPatiRecInfo.ID: Next
        For i = 1 To Me.Tables.Count: Me.Tables(i).文件ID = Me.EPRPatiRecInfo.ID:  Next
        For i = 1 To Me.Compends.Count: Me.Compends(i).文件ID = Me.EPRPatiRecInfo.ID: Next
        For i = 1 To Me.Diagnosises.Count: Me.Diagnosises(i).文件ID = Me.EPRPatiRecInfo.ID: Next
        For i = 1 To Me.Signs.Count:  Me.Signs(i).文件ID = Me.EPRPatiRecInfo.ID: Next
    End Select
        
    'ID号有效性检测！
    Call Me.Elements.CheckValidIDs(EditType)
    Call Me.Tables.CheckValidIDs(EditType)
    Call Me.Pictures.CheckValidIDs(EditType)
    Call Me.Compends.CheckValidIDs(EditType)
    Call Me.Compends.UpdateParentIDs
    Call Me.Diagnosises.CheckValidIDs(EditType)
    Call Me.Signs.CheckValidIDs(EditType)
    
    '保存预处理
    ReDim ArraySQL(1 To 1) As String
    If Me.EditType = cprET_病历文件定义 Then
        If gobjFSO.FileExists(App.Path & "\Head.rtf") Then gobjFSO.DeleteFile App.Path & "\Head.rtf", True
        edtThis.SaveHead App.Path & "\Head.rtf"
        If gobjFSO.FileExists(App.Path & "\Head.rtf") Then
            Call zlBlobSql(12, EPRFileInfo.种类 & "-" & EPRFileInfo.页面, App.Path & "\Head.rtf", ArraySQL)
            gobjFSO.DeleteFile App.Path & "\Head.rtf", True
        End If
        strHead = zl9ComLib.zlStr.ToVarchar(edtThis.HeadFileText, 1000)
        
        If gobjFSO.FileExists(App.Path & "\Foot.rtf") Then gobjFSO.DeleteFile App.Path & "\Foot.rtf", True
        edtThis.SaveFoot App.Path & "\Foot.rtf"
        If gobjFSO.FileExists(App.Path & "\Foot.rtf") Then
            Call zlBlobSql(13, EPRFileInfo.种类 & "-" & EPRFileInfo.页面, App.Path & "\Foot.rtf", ArraySQL)
            gobjFSO.DeleteFile App.Path & "\Foot.rtf", True
        End If
        strFoot = zl9ComLib.zlStr.ToVarchar(edtThis.FootFileText, 1000)
        
        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
        ArraySQL(UBound(ArraySQL)) = "Zl_病历页面格式_Update(" & Me.EPRFileInfo.种类 & ",'" & Me.EPRFileInfo.页面 & "','" & Me.EPRFileInfo.页面名称 & "'," & _
            Me.EPRFileInfo.报表 & ",'" & Me.EPRFileInfo.GetFormat(edtThis) & "'," & _
            "'" & Replace(strHead, "'", "’") & "','" & Replace(strFoot, "'", "’") & "')"
        
        '页眉图片的保存SQL
        Dim strPicFile As String
        If Not (edtThis.Picture Is Nothing) Then
            If edtThis.Picture.Handle <> 0 Then
                strPicFile = App.Path & "\Head-" & Me.EPRFileInfo.种类 & Me.EPRFileInfo.页面 & ".BMP"
                If gobjFSO.FileExists(strPicFile) Then gobjFSO.DeleteFile strPicFile, True
                SavePicture edtThis.Picture, strPicFile
                If gobjFSO.FileExists(strPicFile) Then
                    Call zlBlobSql(7, Me.EPRFileInfo.种类 & "-" & Me.EPRFileInfo.页面, strPicFile, ArraySQL)
                    gobjFSO.DeleteFile strPicFile, True      '删除临时文件
                End If
            End If
        End If
    End If
    
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lngEnd As Long, M As Long, N As Long, p As Long
    Dim lng序号 As Long, StrText As String, lngCompendKey As Long, lngCompendID As Double
    Dim bln是否换行 As Boolean

    lng序号 = 1     '按照CRLF来分段
    StrText = edtThis.Text
    p = 0
    lngEnd = Len(edtThis.Text)
    Do While p < lngEnd
        '获取关键字位置 M
        bFinded = FindNextAnyKey(edtThis, p + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            M = lKSS
        Else
            M = lngEnd
        End If
        '获取vbCrlf位置 N
        N = InStr(p + 1, StrText, vbCrLf, vbTextCompare)
        If N > 0 Then
            N = N - 1
        Else
            N = lngEnd
        End If

        lngCompendKey = GetCurCompendKey(edtThis, p)
        If lngCompendKey = 0 Then
            lngCompendID = 0
        Else
            lngCompendID = Me.Compends("K" & lngCompendKey).ID
        End If
        If M < N Then
            '保存文本
            GetPlainTextSaveSQL ArraySQL, edtThis, p, M, lng序号, lngCompendID, False

            '保存对象
            If bFinded Then
                Select Case sKeyType
                Case "O"
                    Me.Compends("K" & lKey).对象序号 = lng序号
'                    Me.Compends("K" & lKey).父ID = lngCompendID
                    p = lKEE    '调整当前位置
                    Call Me.Compends("K" & lKey).GetSaveSQL(Me.EditType, ArraySQL)
                    lng序号 = lng序号 + 1
                Case "P"
                    Me.Pictures("K" & lKey).对象序号 = lng序号
                    Me.Pictures("K" & lKey).父ID = lngCompendID
                    p = lKEE    '调整当前位置
                    If edtThis.Range(lKEE, lKEE + 2) = vbCrLf Then
                        Me.Pictures("K" & lKey).是否换行 = True
                        p = p + 2
                    Else
                        Me.Pictures("K" & lKey).是否换行 = False
                    End If
                    Call Me.Pictures("K" & lKey).GetSaveSQL(Me.EditType, ArraySQL)
                    lng序号 = lng序号 + 1
                Case "T"
                    Me.Tables("K" & lKey).对象序号 = lng序号
                    Me.Tables("K" & lKey).父ID = lngCompendID
                    p = lKEE    '调整当前位置
                    If edtThis.Range(lKEE, lKEE + 2) = vbCrLf Then
                        Me.Tables("K" & lKey).是否换行 = True
                        p = p + 2
                    Else
                        Me.Tables("K" & lKey).是否换行 = False
                    End If
                    If EditType = cprET_单病历编辑 Or EditType = cprET_单病历审核 Then
                        Call Me.Tables("K" & lKey).GetSaveSQL(Me.EditType, ArraySQL, Me.EPRPatiRecInfo.ID)
                    Else
                        Call Me.Tables("K" & lKey).GetSaveSQL(Me.EditType, ArraySQL)
                    End If
                    lng序号 = lng序号 + 1
                Case "E"
                    Me.Elements("K" & lKey).对象序号 = lng序号
                    Me.Elements("K" & lKey).父ID = lngCompendID
                    p = lKEE    '调整当前位置
                    If edtThis.Range(lKEE, lKEE + 2) = vbCrLf Then
                        Me.Elements("K" & lKey).是否换行 = True
                        p = p + 2
                    Else
                        Me.Elements("K" & lKey).是否换行 = False
                    End If
                    Call Me.Elements("K" & lKey).GetSaveSQL(Me.EditType, ArraySQL)
                    lng序号 = lng序号 + 1
                Case "D"
                    Me.Diagnosises("K" & lKey).对象序号 = lng序号
                    Me.Diagnosises("K" & lKey).父ID = lngCompendID
                    p = lKEE    '调整当前位置
                    Call Me.Diagnosises("K" & lKey).GetSaveSQL(Me.EditType, ArraySQL)
                    lng序号 = lng序号 + 1
                Case "S"
                    If lKey > 0 And (Me.EditType = cprET_单病历编辑 Or Me.EditType = cprET_单病历审核) Then
                        '只有在病人病历编辑时才允许保存签名
                        Me.Signs("K" & lKey).父ID = lngCompendID
                        Me.Signs("K" & lKey).GetSaveSQL ArraySQL, lng序号
                        lng序号 = lng序号 + 1
                    End If
                    p = lKEE    '调整当前位置
                Case Else
                    p = M
                End Select
            Else
                p = M
            End If
        Else
            '以回车结尾的纯文本的保存
            If edtThis.Range(N, N + 2) = vbCrLf And edtThis.Range(N, N + 2).Font.Protected = True Then
                '是否换行＝False，该回车属于下一个对象（图片或者表格）
                If Me.EditType = cprET_单病历审核 Then
                    GetPlainTextSaveSQL ArraySQL, edtThis, p, N, lng序号, lngCompendID, False
                Else
                    If p < N Then GetPlainTextSaveSQL ArraySQL, edtThis, p, N, lng序号, lngCompendID, False
                End If
            Else
                '是否换行＝True
                If Me.EditType = cprET_单病历审核 Then
                    GetPlainTextSaveSQL ArraySQL, edtThis, p, N + 2, lng序号, lngCompendID, False
                Else
                    GetPlainTextSaveSQL ArraySQL, edtThis, p, N, lng序号, lngCompendID, True
                End If
            End If
            p = N + 2
        End If
    Loop
    
    '保存后期处理
    mIntTextIDPos = 0 '保存完毕后清0以备下次保存时可再用
    Select Case EditType
    Case cprET_病历文件定义
        gstrSQL = "Zl_病历文件结构_Commit(" & Me.EPRFileInfo.ID & ")"
    Case cprET_全文示范编辑
        gstrSQL = "Zl_病历范文内容_Commit(" & Me.EPRDemoInfo.ID & ")"
    Case cprET_单病历编辑
        gstrSQL = "Zl_电子病历记录_Update(" & Me.EPRPatiRecInfo.ID & "," & _
            Me.EPRPatiRecInfo.病人来源 & "," & Me.EPRPatiRecInfo.病人ID & "," & Me.EPRPatiRecInfo.主页ID & "," & _
            IIf(Me.EPRPatiRecInfo.婴儿 = 0, "NULL", Me.EPRPatiRecInfo.婴儿) & "," & _
            IIf(Me.EPRPatiRecInfo.科室ID = 0, "NULL", Me.EPRPatiRecInfo.科室ID) & "," & Me.EPRFileInfo.ID & "," & _
            IIf(Me.EPRPatiRecInfo.医嘱id = 0, "NULL", Me.EPRPatiRecInfo.医嘱id) & ",To_Date('" & Format(Me.EPRPatiRecInfo.创建时间, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"
    Case cprET_单病历审核
        gstrSQL = "Zl_电子病历记录_Update(" & Me.EPRPatiRecInfo.ID & "," & _
            Me.EPRPatiRecInfo.病人来源 & "," & Me.EPRPatiRecInfo.病人ID & "," & Me.EPRPatiRecInfo.主页ID & "," & _
            IIf(Me.EPRPatiRecInfo.婴儿 = 0, "NULL", Me.EPRPatiRecInfo.婴儿) & "," & _
            IIf(Me.EPRPatiRecInfo.科室ID = 0, "NULL", Me.EPRPatiRecInfo.科室ID) & "," & Me.EPRFileInfo.ID & ")"
    End Select
    ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
    ArraySQL(UBound(ArraySQL)) = gstrSQL
    
    '=========================================================================================
    '保存RTFText的Sql
    '=========================================================================================
    'RTF中图片的删除
    blnReadOnly = edtThis.ReadOnly
    blnForce = edtThis.ForceEdit
    edtThis.Freeze
    edtThis.InProcessing = True
    edtThis.ForceEdit = True
    edtThis.ReadOnly = False
    Call PreSavingRTFText(edtThis)      '保存RTFText的预处理
    edtThis.Tag = "SaveEPRDoc"
    
    lS = edtThis.Selection.StartPos
    lE = edtThis.Selection.EndPos
    
    Dim ChildPath As String, strFile As String
    ChildPath = IIf(Me.EPRPatiRecInfo.ID <> 0, Me.EPRPatiRecInfo.ID, GUID)
    strFile = App.Path & "\" & ChildPath & "\TMP.rtf"
    If Not gobjFSO.FolderExists(App.Path & "\" & ChildPath) Then Call gobjFSO.CreateFolder(App.Path & "\" & ChildPath)
    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True     '保存为临时文件
    edtThis.SaveDoc strFile
    
    strFileContent = GetTextContent(edtThis.Text)
    Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat
    '恢复图片和表格内容
    For i = 1 To Me.Pictures.Count
        bFinded = FindKey(edtThis, "P", Me.Pictures(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then '还原图片
            Me.Pictures(i).DeleteFromEditor edtThis
            Me.Pictures(i).InsertIntoEditor edtThis, -1, True
        End If
    Next
    For i = 1 To Me.Tables.Count
        bFinded = FindKey(edtThis, "T", Me.Tables(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            '还原表格
            Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt
            Set FontFmt = edtThis.Range(lKSE, lKES).Font.GetFontFmt
            
            Me.Tables(i).是否换行 = False
            edtThis.Range(lKSS, lKEE).Text = ""
            Me.Tables(i).InsertIntoEditor edtThis, lKSS, , , True
            
            edtThis.Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            edtThis.Range(lKSE, lKES).Font.SetFontFmt FontFmt
            edtThis.Range(lKSS, lKEE).Font.Protected = True
        End If
    Next
    '恢复选中位置
    edtThis.Range(lS, lE).Selected

    strZipFile = zlFileZip(strFile)
    If strZipFile = "" Then
        Err.Raise vbObjectError + 100, , "文件保存失败，请联系管理员检查：" & vbCrLf & vbCrLf & _
            "1、压缩组件是否存在并有效?" & vbCrLf & "2、Windows操作系统登录用户是否具备 导航台所在目录及其子目录的<写入>、<修改>权限?"
    End If
    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True
    
    If gobjFSO.FileExists(strZipFile) Then
        Select Case EditType
        Case cprET_病历文件定义
            blnSave = zlBlobSql(1, EPRFileInfo.ID, strZipFile, ArraySQL)
        Case cprET_全文示范编辑
            If Me.EPRDemoInfo.性质 = 0 Then
                blnSave = zlBlobSql(3, EPRDemoInfo.ID, strZipFile, ArraySQL)
            End If
        Case cprET_单病历编辑, cprET_单病历审核
            blnSave = zlBlobSql(5, EPRPatiRecInfo.ID, strZipFile, ArraySQL)
            Call zlClobSql(EPRPatiRecInfo.ID, strFileContent, ArraySQL)
        End Select
        gobjFSO.DeleteFile strZipFile, True      '删除临时文件
    End If
    If blnSave = False Then GoTo errHand
    If gobjFSO.FolderExists(App.Path & "\" & ChildPath) Then Call gobjFSO.DeleteFolder(App.Path & "\" & ChildPath, True) '删除临时目录

    '=========================================================================================
    '保存内容图片
    '=========================================================================================
    For i = 1 To Me.Pictures.Count
        '恢复图片显示
        blnSave = Me.Pictures(i).SaveOrigPic(Me.EditType, ArraySQL)
        If blnSave = False Then GoTo errHand
    Next
    
    '=========================================================================================
    '保存表格图片
    '=========================================================================================
    For i = 1 To Me.Tables.Count
        '恢复表格显示
        For j = 1 To Me.Tables(i).Pictures.Count
            blnSave = Me.Tables(i).Pictures(j).SaveOrigPic(Me.EditType, ArraySQL)
            If blnSave = False Then GoTo errHand
        Next
    Next
    edtThis.ReadOnly = blnReadOnly
    edtThis.ForceEdit = blnForce
    edtThis.UnFreeze
    edtThis.Tag = ""

    '#########################################################################################
    '启动事务
    '================================================================================================   进度条设置
    Dim lngMaxValue As Long         '进度条长度
    For i = 1 To Me.Tables.Count
        lngMaxValue = lngMaxValue + Me.Tables(i).Pictures.Count
    Next
    lngMaxValue = lngMaxValue + UBound(ArraySQL) + 1
    edtThis.ProgressValue = 0
    edtThis.ProgressVisible = True
    '================================================================================================   进度条

    gcnOracle.BeginTrans
    blnTran = True
    For i = 1 To UBound(ArraySQL)
        gstrSQL = ArraySQL(i)
        If Trim(gstrSQL) <> "" Then
            Call zlDatabase.ExecuteProcedure(gstrSQL, "cEPRCompends")
        End If
        edtThis.ProgressValue = (edtThis.ProgressValue + i) / lngMaxValue
    Next
    '=================================================================================================
    Dim bCancel As Boolean
    RaiseEvent BeforeSaved(Me.EPRPatiRecInfo.ID, bCancel)
    If bCancel = True Then GoTo errHand '被取消
    gcnOracle.CommitTrans
    blnTran = False
    
    '同时要更新电子病历记录的“保存时间”和其他状态
    Call UpdateDocInfo(edtThis)
    '更新页眉页脚中的完成时间等要素
    Me.EPRFileInfo.ReadPageHead edtThis
    Me.EPRFileInfo.ReadPageFoot edtThis
    Call GetReplacedHeadFootString(edtThis)
    
    '#########################################################################################
    edtThis.Modified = False
    Screen.MousePointer = 0
    edtThis.ProgressVisible = False
    edtThis.InProcessing = False
    
    If Me.EditMode = cprEM_新增 Then
        Me.EditMode = cprEM_修改
        Me.EPRPatiRecInfo.病历种类 = Me.EPRFileInfo.种类
        Me.EPRPatiRecInfo.病历名称 = Me.EPRFileInfo.名称
    End If
    
    On Error Resume Next
    mfrmParent.RefreshList
    
    SaveEPRDoc = True
    RaiseEvent AfterSaved(Me.EPRPatiRecInfo.ID)
    
    If edtThis.Parent.Name = Me.frmEditor.Name Then
        Call Me.frmEditor.SetStateInfo
    End If
    edtThis.ReadOnly = (Me.EditType = cprET_单病历编辑 And Me.Signs.Count > 0)
    
    Exit Function
errHand:
    mIntTextIDPos = 0 '保存完毕后清0以备下次保存时可再用
    edtThis.Tag = ""
    edtThis.UnFreeze
    If blnTran Then
        gcnOracle.RollbackTrans
        MsgBox "保存数据失败！请截图并联系管理员反馈。" & IIf(Err.Number <> 0, Err.Description, ""), vbInformation, gstrSysName
        Err.Clear
        GoTo exitFun
    End If
    
    If ErrCenter() = 1 Then
        Resume
    End If
    
exitFun:
    Call SaveErrLog
    SaveEPRDoc = False
    edtThis.ProgressVisible = False
    edtThis.InProcessing = False
    Screen.MousePointer = 0
End Function

'################################################################################################################
'## 功能：  将病历文件导出到HTML文档中
'##
'## 参数：  edtThis     :   RTB编辑器控件
'##         strFileName :   HTML文件名（全路径）
'##
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Public Function ExportToHTML(edtThis As Object, strFileName As String) As Boolean
    Dim strHead As String, strEnd As String
        
    On Error GoTo LL
    'HTML头
    strHead = "<html><head><title>" & Me.EPRFileInfo.名称 & "</title></head>" & vbCrLf & _
        "<div align=""center""><object classid=""clsid:3B7C8860-D78F-101B-B9B5-04021C009402"" id=""edtThis"" width=""" & _
        (edtThis.PaperWidth - edtThis.MarginLeft - edtThis.MarginRight + 300) / Screen.TwipsPerPixelX & _
        """ height=""100%"">" & vbCrLf & _
        "<param name=""TextRTF"" value="""
    'HTML尾
    strEnd = """>" & vbCrLf & _
        "</param>" & vbCrLf & _
        "<param name=""BorderStyle"" value=""0""/>" & vbCrLf & _
        "<param name=""Appearance"" value=""0""/>" & vbCrLf & _
        "<param name=""ScrollBars"" value=""3""/>" & vbCrLf & _
        "<param name=""DisableNoScroll"" value=""True""/>" & vbCrLf & _
        "</OBJECT></div>" & vbCrLf & _
        "</html>"
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Dim fs As FileSystemObject, f As TextStream
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(strFileName, ForWriting, TristateUseDefault)
    f.Write strHead & Replace(edtThis.TextRTF, """", "'") & strEnd
    f.Close
    ExportToHTML = True
    Exit Function
LL:
    ExportToHTML = False
End Function

'################################################################################################################
'## 功能：  将病历文件导出到XML文档中
'##
'## 参数：  edtThis     :   RTB编辑器控件
'##         strFileName :   XML文件名（全路径）
'##
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Public Function ExportToXMLFile(ByRef edtThis As Object, ByVal strFileName As String) As Boolean
    Dim i As Long, j As Long, k As Long
    Dim oDoc As DOMDocument             'xml文档
    Dim oRoot  As IXMLDOMElement        '根节点
    Dim oNode As IXMLDOMNode            '父节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim oSubNode2 As IXMLDOMNode        '节点
    Dim oSubNode3 As IXMLDOMNode        '节点
    Dim oSubNode4 As IXMLDOMNode        '节点
    Dim oSubNode5 As IXMLDOMNode        '节点
    Dim oStream As New ADODB.Stream     '流对象
    Dim strPath As String               '临时文件目录
    Dim strPic As String                '临时图片文件
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    
    'XML文档
    Set oDoc = New DOMDocument
    '注释
    oDoc.appendChild oDoc.createComment(gstrSysName & "  " & _
        "操作员:" & gstrUserName & "，部门:" & gstrDeptName & "，时间:" & _
        Format(Now(), "YYYY年MM月DD日"))
    '根节点
    Set oRoot = oDoc.createElement("Document")
    Set oDoc.documentElement = oRoot    '设置为根节点
    Call oRoot.setAttribute("EditType", Me.EditType)
    Call oRoot.setAttribute("EditMode", Me.EditMode)
    
    'EPRFileInfo
    Set oNode = CreateNode(1, oRoot, "EPRFileInfo", NODE_ELEMENT, "")
        CreateNode 2, oNode, "ID", , Me.EPRFileInfo.ID      '子节点
        CreateNode 2, oNode, "种类", , Me.EPRFileInfo.种类
        CreateNode 2, oNode, "编号", , Me.EPRFileInfo.编号
        CreateNode 2, oNode, "名称", , Me.EPRFileInfo.名称
        CreateNode 2, oNode, "说明", , Me.EPRFileInfo.说明
        CreateNode 2, oNode, "页面", , Me.EPRFileInfo.页面
        CreateNode 2, oNode, "保留", , Me.EPRFileInfo.保留
        CreateNode 2, oNode, "通用", , Me.EPRFileInfo.通用
        CreateNode 2, oNode, "报表", , Me.EPRFileInfo.报表
        CreateNode 2, oNode, "格式", , Me.EPRFileInfo.格式
        CreateNode 2, oNode, "页眉", , Me.EPRFileInfo.页眉
        CreateNode 2, oNode, "页脚", , Me.EPRFileInfo.页脚
    'EPRDemoInfo
    Set oNode = CreateNode(1, oRoot, "EPRDemoInfo", NODE_ELEMENT, "")
        CreateNode 2, oNode, "ID", , Me.EPRDemoInfo.ID
        CreateNode 2, oNode, "文件ID", , Me.EPRDemoInfo.文件ID
        CreateNode 2, oNode, "编号", , Me.EPRDemoInfo.编号
        CreateNode 2, oNode, "名称", , Me.EPRDemoInfo.名称
        CreateNode 2, oNode, "性质", , Me.EPRDemoInfo.性质
        CreateNode 2, oNode, "通用级", , Me.EPRDemoInfo.通用级
        CreateNode 2, oNode, "科室ID", , Me.EPRDemoInfo.科室ID
        CreateNode 2, oNode, "人员ID", , Me.EPRDemoInfo.人员ID
    'EPRPatiRecInfo
    Set oNode = CreateNode(1, oRoot, "EPRPatiRecInfo", NODE_ELEMENT, "")
        CreateNode 2, oNode, "ID", , Me.EPRPatiRecInfo.ID
        CreateNode 2, oNode, "病人来源", , Me.EPRPatiRecInfo.病人来源
        CreateNode 2, oNode, "病人ID", , Me.EPRPatiRecInfo.病人ID
        CreateNode 2, oNode, "主页ID", , Me.EPRPatiRecInfo.主页ID
        CreateNode 2, oNode, "婴儿", , Me.EPRPatiRecInfo.婴儿
        CreateNode 2, oNode, "科室ID", , Me.EPRPatiRecInfo.科室ID
        CreateNode 2, oNode, "病历种类", , Me.EPRPatiRecInfo.病历种类
        CreateNode 2, oNode, "文件ID", , Me.EPRPatiRecInfo.文件ID
        CreateNode 2, oNode, "病历名称", , Me.EPRPatiRecInfo.病历名称
        CreateNode 2, oNode, "最后版本", , Me.EPRPatiRecInfo.最后版本
        CreateNode 2, oNode, "完成时间", , Me.EPRPatiRecInfo.完成时间
        CreateNode 2, oNode, "签名级别", , Me.EPRPatiRecInfo.签名级别
        CreateNode 2, oNode, "归档人", , Me.EPRPatiRecInfo.归档人
        CreateNode 2, oNode, "归档日期", , Me.EPRPatiRecInfo.归档日期
        CreateNode 2, oNode, "创建人", , Me.EPRPatiRecInfo.创建人
        CreateNode 2, oNode, "创建时间", , Me.EPRPatiRecInfo.创建时间
        CreateNode 2, oNode, "保存人", , Me.EPRPatiRecInfo.保存人
        CreateNode 2, oNode, "保存时间", , Me.EPRPatiRecInfo.保存时间
        CreateNode 2, oNode, "申请id", , Me.EPRPatiRecInfo.医嘱id
    'Compends
    Set oNode = CreateNode(1, oRoot, "Compends", NODE_ELEMENT, "")
    For i = 1 To Me.Compends.Count
        Set oSubNode1 = CreateNode(2, oNode, "Compend", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "Key", , Me.Compends(i).Key
            CreateNode 3, oSubNode1, "ID", , Me.Compends(i).ID
            CreateNode 3, oSubNode1, "文件ID", , Me.Compends(i).文件ID
            CreateNode 3, oSubNode1, "开始版", , Me.Compends(i).开始版
            CreateNode 3, oSubNode1, "终止版", , Me.Compends(i).终止版
            CreateNode 3, oSubNode1, "父ID", , Me.Compends(i).父ID
            CreateNode 3, oSubNode1, "父Key", , Me.Compends(i).父Key
            CreateNode 3, oSubNode1, "对象序号", , Me.Compends(i).对象序号
            CreateNode 3, oSubNode1, "保留对象", , Me.Compends(i).保留对象
            CreateNode 3, oSubNode1, "名称", , Me.Compends(i).名称
            CreateNode 3, oSubNode1, "说明", , Me.Compends(i).说明
            CreateNode 3, oSubNode1, "预制提纲ID", , Me.Compends(i).预制提纲ID
            CreateNode 3, oSubNode1, "定义提纲ID", , Me.Compends(i).定义提纲ID
            CreateNode 3, oSubNode1, "复用提纲", , Me.Compends(i).复用提纲
            CreateNode 3, oSubNode1, "使用时机", , Me.Compends(i).使用时机
            CreateNode 3, oSubNode1, "Level", , Me.Compends(i).Level
            CreateNode 3, oSubNode1, "内部序号", , Me.Compends(i).内部序号
    Next
    'Signs
    Set oNode = CreateNode(1, oRoot, "Signs", NODE_ELEMENT, "")
    For i = 1 To Me.Signs.Count
        Set oSubNode1 = CreateNode(2, oNode, "Sign", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "Key", , Me.Signs(i).Key
            CreateNode 3, oSubNode1, "ID", , Me.Signs(i).ID
            CreateNode 3, oSubNode1, "文件ID", , Me.Signs(i).文件ID
            CreateNode 3, oSubNode1, "父ID", , Me.Signs(i).父ID
            CreateNode 3, oSubNode1, "开始版", , Me.Signs(i).开始版
            CreateNode 3, oSubNode1, "终止版", , Me.Signs(i).终止版
            CreateNode 3, oSubNode1, "对象序号", , Me.Signs(i).对象序号
            CreateNode 3, oSubNode1, "姓名", , Me.Signs(i).姓名
            CreateNode 3, oSubNode1, "签名信息", , Me.Signs(i).签名信息
            CreateNode 3, oSubNode1, "前置文字", , Me.Signs(i).前置文字
            CreateNode 3, oSubNode1, "签名级别", , Me.Signs(i).签名级别
            CreateNode 3, oSubNode1, "对象属性", , Me.Signs(i).对象属性
    Next
    'Diagnosises
    Set oNode = CreateNode(1, oRoot, "Diagnosises", NODE_ELEMENT, "")
    For i = 1 To Me.Diagnosises.Count
        Set oSubNode1 = CreateNode(2, oNode, "Diagnosis", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "Key", , Me.Diagnosises(i).Key
            CreateNode 3, oSubNode1, "ID", , Me.Diagnosises(i).ID
            CreateNode 3, oSubNode1, "文件ID", , Me.Diagnosises(i).文件ID
            CreateNode 3, oSubNode1, "父ID", , Me.Diagnosises(i).父ID
            CreateNode 3, oSubNode1, "开始版", , Me.Diagnosises(i).开始版
            CreateNode 3, oSubNode1, "终止版", , Me.Diagnosises(i).终止版
            CreateNode 3, oSubNode1, "对象序号", , Me.Diagnosises(i).对象序号
            CreateNode 3, oSubNode1, "类型", , Me.Diagnosises(i).类型
            CreateNode 3, oSubNode1, "中医", , Me.Diagnosises(i).中医
            CreateNode 3, oSubNode1, "疾病id", , Me.Diagnosises(i).疾病id
            CreateNode 3, oSubNode1, "诊断id", , Me.Diagnosises(i).诊断id
            CreateNode 3, oSubNode1, "证候id", , Me.Diagnosises(i).证候id
            CreateNode 3, oSubNode1, "描述", , Me.Diagnosises(i).描述
            CreateNode 3, oSubNode1, "疑诊", , Me.Diagnosises(i).疑诊
            CreateNode 3, oSubNode1, "日期", , Me.Diagnosises(i).日期
    Next
    'Elements
    Set oNode = CreateNode(1, oRoot, "Elements", NODE_ELEMENT, "")
    For i = 1 To Me.Elements.Count
        Set oSubNode1 = CreateNode(2, oNode, "Element", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "Key", , Me.Elements(i).Key
            CreateNode 3, oSubNode1, "ID", , Me.Elements(i).ID
            CreateNode 3, oSubNode1, "文件ID", , Me.Elements(i).文件ID
            CreateNode 3, oSubNode1, "开始版", , Me.Elements(i).开始版
            CreateNode 3, oSubNode1, "终止版", , Me.Elements(i).终止版
            CreateNode 3, oSubNode1, "父ID", , Me.Elements(i).父ID
            CreateNode 3, oSubNode1, "对象序号", , Me.Elements(i).对象序号
            CreateNode 3, oSubNode1, "保留对象", , Me.Elements(i).保留对象
            CreateNode 3, oSubNode1, "内容文本", , Me.Elements(i).内容文本
            CreateNode 3, oSubNode1, "是否换行", , Me.Elements(i).是否换行
            CreateNode 3, oSubNode1, "诊治要素ID", , Me.Elements(i).诊治要素ID
            CreateNode 3, oSubNode1, "替换域", , Me.Elements(i).替换域
            CreateNode 3, oSubNode1, "要素名称", , Me.Elements(i).要素名称
            CreateNode 3, oSubNode1, "要素类型", , Me.Elements(i).要素类型
            CreateNode 3, oSubNode1, "要素长度", , Me.Elements(i).要素长度
            CreateNode 3, oSubNode1, "要素小数", , Me.Elements(i).要素小数
            CreateNode 3, oSubNode1, "要素单位", , Me.Elements(i).要素单位
            CreateNode 3, oSubNode1, "要素表示", , Me.Elements(i).要素表示
            CreateNode 3, oSubNode1, "输入形态", , Me.Elements(i).输入形态
            CreateNode 3, oSubNode1, "要素值域", , Me.Elements(i).要素值域
            CreateNode 3, oSubNode1, "对象属性", , Me.Elements(i).对象属性
    Next
    'Pictures
    Set oNode = CreateNode(1, oRoot, "Pictures", NODE_ELEMENT, "")
    For i = 1 To Me.Pictures.Count
        Set oSubNode1 = CreateNode(2, oNode, "Picture", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "Key", , Me.Pictures(i).Key
            CreateNode 3, oSubNode1, "ID", , Me.Pictures(i).ID
            CreateNode 3, oSubNode1, "文件ID", , Me.Pictures(i).文件ID
            CreateNode 3, oSubNode1, "开始版", , Me.Pictures(i).开始版
            CreateNode 3, oSubNode1, "终止版", , Me.Pictures(i).终止版
            CreateNode 3, oSubNode1, "父ID", , Me.Pictures(i).父ID
            CreateNode 3, oSubNode1, "对象序号", , Me.Pictures(i).对象序号
            CreateNode 3, oSubNode1, "保留对象", , Me.Pictures(i).保留对象
            CreateNode 3, oSubNode1, "是否换行", , Me.Pictures(i).是否换行
            CreateNode 3, oSubNode1, "Row", , Me.Pictures(i).Row
            CreateNode 3, oSubNode1, "Col", , Me.Pictures(i).Col
            CreateNode 3, oSubNode1, "PictureType", , Me.Pictures(i).PictureType
            CreateNode 3, oSubNode1, "MarkedPicID", , Me.Pictures(i).MarkedPicID
            CreateNode 3, oSubNode1, "OrigWidth", , Me.Pictures(i).OrigWidth
            CreateNode 3, oSubNode1, "OrigHeight", , Me.Pictures(i).OrigHeight
            CreateNode 3, oSubNode1, "Width", , Me.Pictures(i).Width
            CreateNode 3, oSubNode1, "Height", , Me.Pictures(i).Height
            CreateNode 3, oSubNode1, "ZoomFactor", , Me.Pictures(i).ZoomFactor
            CreateNode 3, oSubNode1, "Modified", , Me.Pictures(i).Modified
            '存储图片对象
            oStream.Type = adTypeBinary
            oStream.Open
            strPic = strPath & "\XMLPIC" & App.hInstance & ".jpg"
            SavePicture Me.Pictures(i).OrigPic, strPic
            oStream.LoadFromFile strPic
            Set oSubNode2 = oDoc.createElement("OrigPic")
            oSubNode2.datatype = "bin.base64"
            oSubNode2.nodeTypedValue = oStream.Read
            oSubNode1.appendChild oSubNode2
            oStream.Close
            '删除临时文件
            If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
            'PicMarks
            Set oSubNode2 = CreateNode(3, oSubNode1, "PicMarks", NODE_ELEMENT, "")
            For j = 1 To Me.Pictures(i).PicMarks.Count
                Set oSubNode3 = CreateNode(4, oSubNode2, "PicMark", NODE_ELEMENT, "")
                    CreateNode 5, oSubNode3, "Key", , Me.Pictures(i).PicMarks(j).Key
                    CreateNode 5, oSubNode3, "类型", , Me.Pictures(i).PicMarks(j).类型
                    CreateNode 5, oSubNode3, "内容", , Me.Pictures(i).PicMarks(j).内容
                    CreateNode 5, oSubNode3, "点集", , Me.Pictures(i).PicMarks(j).点集
                    CreateNode 5, oSubNode3, "X1", , Me.Pictures(i).PicMarks(j).X1
                    CreateNode 5, oSubNode3, "Y1", , Me.Pictures(i).PicMarks(j).Y1
                    CreateNode 5, oSubNode3, "X2", , Me.Pictures(i).PicMarks(j).X2
                    CreateNode 5, oSubNode3, "Y2", , Me.Pictures(i).PicMarks(j).Y2
                    CreateNode 5, oSubNode3, "填充色", , Me.Pictures(i).PicMarks(j).填充色
                    CreateNode 5, oSubNode3, "填充方式", , Me.Pictures(i).PicMarks(j).填充方式
                    CreateNode 5, oSubNode3, "线条色", , Me.Pictures(i).PicMarks(j).线条色
                    CreateNode 5, oSubNode3, "字体色", , Me.Pictures(i).PicMarks(j).字体色
                    CreateNode 5, oSubNode3, "线型", , Me.Pictures(i).PicMarks(j).线型
                    CreateNode 5, oSubNode3, "线宽", , Me.Pictures(i).PicMarks(j).线宽
                    CreateNode 5, oSubNode3, "字体", , Me.Pictures(i).PicMarks(j).字体
            Next
    Next
    'Tables
    Set oNode = CreateNode(1, oRoot, "Tables", NODE_ELEMENT, "")
    For i = 1 To Me.Tables.Count
        Set oSubNode1 = CreateNode(2, oNode, "Table", NODE_ELEMENT, "")
        CreateNode 3, oSubNode1, "Key", , Me.Tables(i).Key
        CreateNode 3, oSubNode1, "ID", , Me.Tables(i).ID
        CreateNode 3, oSubNode1, "文件ID", , Me.Tables(i).文件ID
        CreateNode 3, oSubNode1, "开始版", , Me.Tables(i).开始版
        CreateNode 3, oSubNode1, "终止版", , Me.Tables(i).终止版
        CreateNode 3, oSubNode1, "父ID", , Me.Tables(i).父ID
        CreateNode 3, oSubNode1, "对象序号", , Me.Tables(i).对象序号
        CreateNode 3, oSubNode1, "保留对象", , Me.Tables(i).保留对象
        CreateNode 3, oSubNode1, "是否换行", , Me.Tables(i).是否换行
        CreateNode 3, oSubNode1, "预制提纲ID", , Me.Tables(i).预制提纲ID
        CreateNode 3, oSubNode1, "Rows", , Me.Tables(i).Rows
        CreateNode 3, oSubNode1, "Cols", , Me.Tables(i).Cols
        CreateNode 3, oSubNode1, "对象属性", , Me.Tables(i).对象属性
'        CreateNode 3, oSubNode1, "FixedRows", , Me.Tables(i).FixedRows
'        CreateNode 3, oSubNode1, "FixedCols", , Me.Tables(i).FixedCols
'        CreateNode 3, oSubNode1, "ShowGrid", , Me.Tables(i).ShowGrid
        CreateNode 3, oSubNode1, "Key", , Me.Tables(i).Key
        'Table -> Cells
        Set oSubNode2 = CreateNode(3, oSubNode1, "Cells", NODE_ELEMENT, "")
        For j = 1 To Me.Tables(i).Cells.Count
            Set oSubNode3 = CreateNode(4, oSubNode2, "Cell", NODE_ELEMENT, "")
                CreateNode 5, oSubNode3, "Key", , Me.Tables(i).Cells(j).Key
                CreateNode 5, oSubNode3, "ID", , Me.Tables(i).Cells(j).ID
                CreateNode 5, oSubNode3, "文件ID", , Me.Tables(i).Cells(j).文件ID
                CreateNode 5, oSubNode3, "父ID", , Me.Tables(i).Cells(j).父ID
                CreateNode 5, oSubNode3, "对象序号", , Me.Tables(i).Cells(j).对象序号
                CreateNode 5, oSubNode3, "内容文本", , Me.Tables(i).Cells(j).内容文本
                CreateNode 5, oSubNode3, "对象属性", , Me.Tables(i).Cells(j).对象属性
        Next
        'Table -> Pictures
        Set oSubNode2 = CreateNode(3, oSubNode1, "Pictures", NODE_ELEMENT, "")
        For j = 1 To Me.Tables(i).Pictures.Count
            Set oSubNode3 = CreateNode(4, oSubNode2, "Picture", NODE_ELEMENT, "")
                CreateNode 5, oSubNode3, "Key", , Me.Tables(i).Pictures(j).Key
                CreateNode 5, oSubNode3, "ID", , Me.Tables(i).Pictures(j).ID
                CreateNode 5, oSubNode3, "文件ID", , Me.Tables(i).Pictures(j).文件ID
                CreateNode 5, oSubNode3, "开始版", , Me.Tables(i).Pictures(j).开始版
                CreateNode 5, oSubNode3, "终止版", , Me.Tables(i).Pictures(j).终止版
                CreateNode 5, oSubNode3, "父ID", , Me.Tables(i).Pictures(j).父ID
                CreateNode 5, oSubNode3, "对象序号", , Me.Tables(i).Pictures(j).对象序号
                CreateNode 5, oSubNode3, "保留对象", , Me.Tables(i).Pictures(j).保留对象
                CreateNode 5, oSubNode3, "是否换行", , Me.Tables(i).Pictures(j).是否换行
                CreateNode 5, oSubNode3, "Row", , Me.Tables(i).Pictures(j).Row
                CreateNode 5, oSubNode3, "Col", , Me.Tables(i).Pictures(j).Col
                CreateNode 5, oSubNode3, "PictureType", , Me.Tables(i).Pictures(j).PictureType
                CreateNode 5, oSubNode3, "MarkedPicID", , Me.Tables(i).Pictures(j).MarkedPicID
                CreateNode 5, oSubNode3, "OrigWidth", , Me.Tables(i).Pictures(j).OrigWidth
                CreateNode 5, oSubNode3, "OrigHeight", , Me.Tables(i).Pictures(j).OrigHeight
                CreateNode 5, oSubNode3, "Width", , Me.Tables(i).Pictures(j).Width
                CreateNode 5, oSubNode3, "Height", , Me.Tables(i).Pictures(j).Height
                CreateNode 5, oSubNode3, "ZoomFactor", , Me.Tables(i).Pictures(j).ZoomFactor
                CreateNode 5, oSubNode3, "Modified", , Me.Tables(i).Pictures(j).Modified
                '存储图片对象
                oStream.Type = adTypeBinary
                oStream.Open
                strPic = strPath & "\XMLPIC" & App.hInstance & ".jpg"
                SavePicture Me.Tables(i).Pictures(j).OrigPic, strPic
                oStream.LoadFromFile strPic
                Set oSubNode4 = oDoc.createElement("OrigPic")
                oSubNode4.datatype = "bin.base64"
                oSubNode4.nodeTypedValue = oStream.Read
                oSubNode3.appendChild oSubNode4
                oStream.Close
                '删除临时文件
                If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
                
                'PicMarks
                Set oSubNode4 = CreateNode(5, oSubNode3, "PicMarks", NODE_ELEMENT, "")
                For k = 1 To Me.Tables(i).Pictures(j).PicMarks.Count
                    Set oSubNode5 = CreateNode(6, oSubNode4, "PicMark", NODE_ELEMENT, "")
                        CreateNode 6, oSubNode5, "Key", , Me.Tables(i).Pictures(j).PicMarks(k).Key
                        CreateNode 6, oSubNode5, "类型", , Me.Tables(i).Pictures(j).PicMarks(k).类型
                        CreateNode 6, oSubNode5, "内容", , Me.Tables(i).Pictures(j).PicMarks(k).内容
                        CreateNode 6, oSubNode5, "点集", , Me.Tables(i).Pictures(j).PicMarks(k).点集
                        CreateNode 6, oSubNode5, "X1", , Me.Tables(i).Pictures(j).PicMarks(k).X1
                        CreateNode 6, oSubNode5, "Y1", , Me.Tables(i).Pictures(j).PicMarks(k).Y1
                        CreateNode 6, oSubNode5, "X2", , Me.Tables(i).Pictures(j).PicMarks(k).X2
                        CreateNode 6, oSubNode5, "Y2", , Me.Tables(i).Pictures(j).PicMarks(k).Y2
                        CreateNode 6, oSubNode5, "填充色", , Me.Tables(i).Pictures(j).PicMarks(k).填充色
                        CreateNode 6, oSubNode5, "填充方式", , Me.Tables(i).Pictures(j).PicMarks(k).填充方式
                        CreateNode 6, oSubNode5, "线条色", , Me.Tables(i).Pictures(j).PicMarks(k).线条色
                        CreateNode 6, oSubNode5, "字体色", , Me.Tables(i).Pictures(j).PicMarks(k).字体色
                        CreateNode 6, oSubNode5, "线型", , Me.Tables(i).Pictures(j).PicMarks(k).线型
                        CreateNode 6, oSubNode5, "线宽", , Me.Tables(i).Pictures(j).PicMarks(k).线宽
                        CreateNode 6, oSubNode5, "字体", , Me.Tables(i).Pictures(j).PicMarks(k).字体
                Next
        Next
        'Table -> Elements
        Set oSubNode2 = CreateNode(3, oSubNode1, "Elements", NODE_ELEMENT, "")
        For j = 1 To Me.Tables(i).Elements.Count
            Set oSubNode3 = CreateNode(4, oSubNode2, "Element", NODE_ELEMENT, "")
                CreateNode 5, oSubNode3, "Key", , Me.Tables(i).Elements(j).Key
                CreateNode 5, oSubNode3, "ID", , Me.Tables(i).Elements(j).ID
                CreateNode 5, oSubNode3, "文件ID", , Me.Tables(i).Elements(j).文件ID
                CreateNode 5, oSubNode3, "开始版", , Me.Tables(i).Elements(j).开始版
                CreateNode 5, oSubNode3, "终止版", , Me.Tables(i).Elements(j).终止版
                CreateNode 5, oSubNode3, "父ID", , Me.Tables(i).Elements(j).父ID
                CreateNode 5, oSubNode3, "对象序号", , Me.Tables(i).Elements(j).对象序号
                CreateNode 5, oSubNode3, "保留对象", , Me.Tables(i).Elements(j).保留对象
                CreateNode 5, oSubNode3, "内容文本", , Me.Tables(i).Elements(j).内容文本
                CreateNode 5, oSubNode3, "是否换行", , Me.Tables(i).Elements(j).是否换行
                CreateNode 5, oSubNode3, "诊治要素ID", , Me.Tables(i).Elements(j).诊治要素ID
                CreateNode 5, oSubNode3, "替换域", , Me.Tables(i).Elements(j).替换域
                CreateNode 5, oSubNode3, "要素名称", , Me.Tables(i).Elements(j).要素名称
                CreateNode 5, oSubNode3, "要素类型", , Me.Tables(i).Elements(j).要素类型
                CreateNode 5, oSubNode3, "要素长度", , Me.Tables(i).Elements(j).要素长度
                CreateNode 5, oSubNode3, "要素小数", , Me.Tables(i).Elements(j).要素小数
                CreateNode 5, oSubNode3, "要素单位", , Me.Tables(i).Elements(j).要素单位
                CreateNode 5, oSubNode3, "要素表示", , Me.Tables(i).Elements(j).要素表示
                CreateNode 5, oSubNode3, "输入形态", , Me.Tables(i).Elements(j).输入形态
                CreateNode 5, oSubNode3, "要素值域", , Me.Tables(i).Elements(j).要素值域
                CreateNode 5, oSubNode3, "对象属性", , Me.Tables(i).Elements(j).对象属性
        Next
        '附加SQL语句暂时不需要，只拷贝纯文本过去！
'        'Table -> DicSpecSQLs
'        Set oSubNode2 = CreateNode(3, oSubNode1, "DicSpecSQLs", NODE_ELEMENT, "")
'        For j = 1 To Me.Tables(i).DicSpecSQLs.Count
'            CreateNode 4, oSubNode2, "SQL", NODE_ELEMENT, Me.Tables(i).DicSpecSQLs("K" & j)
'        Next
    Next
    
    'RTF文本
    Set oNode = CreateNode(1, oRoot, "Content", NODE_ELEMENT, "")
    Set oSubNode1 = CreateNode(2, oNode, "RTF", NODE_ELEMENT, "")
    CreateNode 3, oSubNode1, "RTFText", NODE_CDATA_SECTION, Replace(edtThis.TextRTF, "]]>", "]] >")
    
    '版本信息
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = oDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call oDoc.insertBefore(pi, oDoc.childNodes(0))
    '直接保存成文件即可
    oDoc.Save strFileName
    
    Set oDoc = Nothing
    ExportToXMLFile = True
    Exit Function
LL:
    ExportToXMLFile = False
End Function

'################################################################################################################
'## 功能：  只获取XML文件中的RTFText，用于比较操作
'##
'## 参数：  strFileName :   XML文件名
'##
'## 返回：  返回RTFText。
'################################################################################################################
Public Function GetRTFOfXMLFile(ByVal strFileName As String) As String
    Dim oDoc As DOMDocument             'xml文档
    Dim oRoot  As IXMLDOMElement        '根节点
    Dim oNode As IXMLDOMNode            '父节点
    Dim oSubNode1 As IXMLDOMNode        '子节点

    Set oDoc = New DOMDocument
    oDoc.Load strFileName
    '如果不包含任何元素，则退出
    If oDoc.documentElement Is Nothing Then
        Exit Function
    End If
    '读取文件结构
    Set oRoot = oDoc.selectSingleNode("Document")       'oRoot置为根节点
    'RTF文本
    Set oNode = oRoot.selectSingleNode("Content")
    Set oSubNode1 = oNode.selectSingleNode("RTF")
    GetRTFOfXMLFile = oSubNode1.Text
End Function

'################################################################################################################
'## 功能：  从XML文件导入病历文件
'##
'## 参数：  edtThis     :   RTB编辑器控件
'##         strFileName :   XML文件名（全路径）
'##         blnPrompt   :   是否提示导入覆盖，默认为True
'##         blnForUndoRedo : 是否用于Undo/Redo，默认为False
'##
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Public Function ImportFromXMLFile(ByRef edtThis As Object, _
    ByVal strFileName As String, _
    Optional blnPrompt As Boolean = True, _
    Optional blnForUndoRedo As Boolean = False) As Boolean
    
    Dim i As Long, j As Long, k As Long, lngSelStart As Long, lngSelEnd As Long
    Dim eEditType As EditTypeEnum
    Dim lKey As Long, lKey2 As Long, lKey3 As Long
    Dim oDoc As DOMDocument             'xml文档
    Dim oRoot  As IXMLDOMElement        '根节点
    Dim oNode As IXMLDOMNode            '父节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim oSubNode2 As IXMLDOMNode        '节点
    Dim oSubNode3 As IXMLDOMNode        '节点
    Dim oSubNode4 As IXMLDOMNode        '节点
    Dim oSubNode5 As IXMLDOMNode        '节点
    Dim oStream As New ADODB.Stream     '流对象
    Dim strPath As String               '临时文件目录
    Dim strPic As String                '临时图片文件名
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))

    Set oDoc = New DOMDocument
    oDoc.Load strFileName
    '如果不包含任何元素，则退出
    If oDoc.documentElement Is Nothing Then
        Exit Function
    End If
    If blnPrompt Then
        If MsgBox("注意：导入文件后原有内容将不可恢复，是否继续覆盖当前文件？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then
            Exit Function
        End If
    End If
    '读取文件结构
    Set oRoot = oDoc.selectSingleNode("Document")       'oRoot置为根节点
    If oRoot Is Nothing Then MsgBox "该XML文件不是正确的病历导出文件！", vbInformation, gstrSysName: Exit Function
    If oRoot.selectNodes("File").Length = 1 Then
         Set oRoot = oRoot.selectSingleNode("File")
    ElseIf oRoot.selectNodes("File").Length > 1 Then
        MsgBox "该XML文件包含多个病历文件，此处不能导入！", vbInformation, gstrSysName: Exit Function
    End If
    '获取基础信息
    On Error Resume Next
    eEditType = Val(oRoot.getAttributeNode("EditType").Text)
    
    Screen.MousePointer = vbHourglass
    edtThis.Tag = "禁止同步"
    edtThis.InProcessing = True
    Set Me.Compends = New cEPRCompends
    Set Me.Pictures = New cEPRPictures
    Set Me.Tables = New cEPRTables
    Set Me.Elements = New cEPRElements
    Set Me.Signs = New cEPRSigns
    Set Me.Diagnosises = New cEPRDiagnosises
    
    'Compends:
    Set oNode = oRoot.selectSingleNode("Compends")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "Key", 0)
        If lKey > 0 Then
            Me.Compends.Add lKey
            Me.Compends("K" & lKey).ID = GetNodeValue(oSubNode1, "ID", 0)
            Me.Compends("K" & lKey).文件ID = GetNodeValue(oSubNode1, "文件ID", 0)
            Me.Compends("K" & lKey).开始版 = GetNodeValue(oSubNode1, "开始版", 1)
            Me.Compends("K" & lKey).终止版 = GetNodeValue(oSubNode1, "终止版", 0)
            Me.Compends("K" & lKey).父ID = GetNodeValue(oSubNode1, "父ID", 0)
            Me.Compends("K" & lKey).父Key = GetNodeValue(oSubNode1, "父Key", 0)
            Me.Compends("K" & lKey).对象序号 = GetNodeValue(oSubNode1, "对象序号", 0)
            Me.Compends("K" & lKey).保留对象 = GetNodeValue(oSubNode1, "保留对象", 0)
            Me.Compends("K" & lKey).名称 = GetNodeValue(oSubNode1, "名称", "")
            Me.Compends("K" & lKey).说明 = GetNodeValue(oSubNode1, "说明", "")
            Me.Compends("K" & lKey).预制提纲ID = GetNodeValue(oSubNode1, "预制提纲ID", 0)
            Me.Compends("K" & lKey).定义提纲ID = GetNodeValue(oSubNode1, "定义提纲ID", 0)
            Me.Compends("K" & lKey).复用提纲 = GetNodeValue(oSubNode1, "复用提纲", 0)
            Me.Compends("K" & lKey).使用时机 = GetNodeValue(oSubNode1, "使用时机", 0)
            Me.Compends("K" & lKey).Level = GetNodeValue(oSubNode1, "Level", 0)
            Me.Compends("K" & lKey).内部序号 = GetNodeValue(oSubNode1, "内部序号", 0)
        End If
    Next
    'Signs:
    Set oNode = oRoot.selectSingleNode("Signs")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "Key", 0)
        If lKey > 0 Then
            Me.Signs.Add lKey
            Me.Signs("K" & lKey).Key = GetNodeValue(oSubNode1, "Key", 0)
            Me.Signs("K" & lKey).ID = GetNodeValue(oSubNode1, "ID", 0)
            Me.Signs("K" & lKey).文件ID = GetNodeValue(oSubNode1, "文件ID", 0)
            Me.Signs("K" & lKey).父ID = GetNodeValue(oSubNode1, "父ID", 0)
            Me.Signs("K" & lKey).开始版 = GetNodeValue(oSubNode1, "开始版", 1)
            Me.Signs("K" & lKey).终止版 = GetNodeValue(oSubNode1, "终止版", 0)
            Me.Signs("K" & lKey).对象序号 = GetNodeValue(oSubNode1, "对象序号", 0)
            Me.Signs("K" & lKey).姓名 = GetNodeValue(oSubNode1, "姓名", "")
            Me.Signs("K" & lKey).签名信息 = GetNodeValue(oSubNode1, "签名信息", "")
            Me.Signs("K" & lKey).前置文字 = GetNodeValue(oSubNode1, "前置文字", "")
            Me.Signs("K" & lKey).签名级别 = GetNodeValue(oSubNode1, "签名级别", 0)
            Me.Signs("K" & lKey).对象属性 = GetNodeValue(oSubNode1, "对象属性", "")
        End If
    Next
    'Diagnosises:
    Set oNode = oRoot.selectSingleNode("Diagnosises")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "Key", 0)
        If lKey > 0 Then
            Me.Diagnosises.Add lKey
            Me.Diagnosises("K" & lKey).Key = GetNodeValue(oSubNode1, "Key", 0)
            Me.Diagnosises("K" & lKey).ID = GetNodeValue(oSubNode1, "ID", 0)
            Me.Diagnosises("K" & lKey).文件ID = GetNodeValue(oSubNode1, "文件ID", 0)
            Me.Diagnosises("K" & lKey).父ID = GetNodeValue(oSubNode1, "父ID", 0)
            Me.Diagnosises("K" & lKey).开始版 = GetNodeValue(oSubNode1, "开始版", 1)
            Me.Diagnosises("K" & lKey).终止版 = GetNodeValue(oSubNode1, "终止版", 0)
            Me.Diagnosises("K" & lKey).对象序号 = GetNodeValue(oSubNode1, "对象序号", 0)
            Me.Diagnosises("K" & lKey).类型 = GetNodeValue(oSubNode1, "类型", 0)
            Me.Diagnosises("K" & lKey).中医 = GetNodeValue(oSubNode1, "中医", 0)
            Me.Diagnosises("K" & lKey).疾病id = GetNodeValue(oSubNode1, "疾病id", 0)
            Me.Diagnosises("K" & lKey).诊断id = GetNodeValue(oSubNode1, "诊断id", 0)
            Me.Diagnosises("K" & lKey).证候id = GetNodeValue(oSubNode1, "证候id", 0)
            Me.Diagnosises("K" & lKey).描述 = GetNodeValue(oSubNode1, "描述", "")
            Me.Diagnosises("K" & lKey).疑诊 = GetNodeValue(oSubNode1, "疑诊", 0)
            Me.Diagnosises("K" & lKey).日期 = GetNodeValue(oSubNode1, "日期", "")
        End If
    Next
    'Elements:
    Set oNode = oRoot.selectSingleNode("Elements")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "Key", 0)
        If lKey > 0 Then
            Me.Elements.Add lKey
            Me.Elements("K" & lKey).ID = GetNodeValue(oSubNode1, "ID", 0)
            Me.Elements("K" & lKey).文件ID = GetNodeValue(oSubNode1, "文件ID", 0)
            Me.Elements("K" & lKey).开始版 = GetNodeValue(oSubNode1, "开始版", 1)
            Me.Elements("K" & lKey).终止版 = GetNodeValue(oSubNode1, "终止版", 0)
            Me.Elements("K" & lKey).父ID = GetNodeValue(oSubNode1, "父ID", 0)
            Me.Elements("K" & lKey).对象序号 = GetNodeValue(oSubNode1, "对象序号", 0)
            Me.Elements("K" & lKey).保留对象 = GetNodeValue(oSubNode1, "保留对象", 0)
            Me.Elements("K" & lKey).内容文本 = GetNodeValue(oSubNode1, "内容文本", "")
            Me.Elements("K" & lKey).是否换行 = GetNodeValue(oSubNode1, "是否换行", 0)
            Me.Elements("K" & lKey).诊治要素ID = GetNodeValue(oSubNode1, "诊治要素ID", 0)
            Me.Elements("K" & lKey).替换域 = GetNodeValue(oSubNode1, "替换域", 0)
            Me.Elements("K" & lKey).要素名称 = GetNodeValue(oSubNode1, "要素名称", "")
            Me.Elements("K" & lKey).要素类型 = GetNodeValue(oSubNode1, "要素类型", 0)
            Me.Elements("K" & lKey).要素长度 = GetNodeValue(oSubNode1, "要素长度", 0)
            Me.Elements("K" & lKey).要素小数 = GetNodeValue(oSubNode1, "要素小数", 0)
            Me.Elements("K" & lKey).要素单位 = GetNodeValue(oSubNode1, "要素单位", "")
            Me.Elements("K" & lKey).要素表示 = GetNodeValue(oSubNode1, "要素表示", 0)
            Me.Elements("K" & lKey).输入形态 = GetNodeValue(oSubNode1, "输入形态", 0)
            Me.Elements("K" & lKey).要素值域 = GetNodeValue(oSubNode1, "要素值域", "")
            Me.Elements("K" & lKey).对象属性 = GetNodeValue(oSubNode1, "对象属性", "")
            '诊治要素的有效性验证（如果无效，那么改为临时要素）
            If Me.Elements("K" & lKey).CheckValid = False Then Me.Elements("K" & lKey).诊治要素ID = 0
        End If
    Next
    'Pictures:
    Set oNode = oRoot.selectSingleNode("Pictures")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "Key", 0)
        If lKey > 0 Then
            Me.Pictures.Add lKey
            Me.Pictures("K" & lKey).ID = GetNodeValue(oSubNode1, "ID", 0)
            Me.Pictures("K" & lKey).文件ID = GetNodeValue(oSubNode1, "文件ID", 0)
            Me.Pictures("K" & lKey).开始版 = GetNodeValue(oSubNode1, "开始版", 1)
            Me.Pictures("K" & lKey).终止版 = GetNodeValue(oSubNode1, "终止版", 0)
            Me.Pictures("K" & lKey).父ID = GetNodeValue(oSubNode1, "父ID", 0)
            Me.Pictures("K" & lKey).对象序号 = GetNodeValue(oSubNode1, "对象序号", 0)
            Me.Pictures("K" & lKey).保留对象 = GetNodeValue(oSubNode1, "保留对象", 0)
            Me.Pictures("K" & lKey).是否换行 = GetNodeValue(oSubNode1, "是否换行", 0)
            Me.Pictures("K" & lKey).Row = GetNodeValue(oSubNode1, "Row", 0)
            Me.Pictures("K" & lKey).Col = GetNodeValue(oSubNode1, "Col", 0)
            Me.Pictures("K" & lKey).PictureType = GetNodeValue(oSubNode1, "PictureType", 0)
            Me.Pictures("K" & lKey).MarkedPicID = GetNodeValue(oSubNode1, "MarkedPicID", 0)
            Me.Pictures("K" & lKey).OrigWidth = GetNodeValue(oSubNode1, "OrigWidth", 0)
            Me.Pictures("K" & lKey).OrigHeight = GetNodeValue(oSubNode1, "OrigHeight", 0)
            Me.Pictures("K" & lKey).Width = GetNodeValue(oSubNode1, "Width", 0)
            Me.Pictures("K" & lKey).Height = GetNodeValue(oSubNode1, "Height", 0)
            Me.Pictures("K" & lKey).ZoomFactor = GetNodeValue(oSubNode1, "ZoomFactor", 0)
            Me.Pictures("K" & lKey).Modified = GetNodeValue(oSubNode1, "Modified", 0)
            
            Set oSubNode2 = oSubNode1.selectSingleNode("OrigPic")
            oStream.Type = adTypeBinary
            oStream.Open
            oStream.Write oSubNode2.nodeTypedValue
            strPic = strPath & "\XML2JPG" & App.hInstance & ".JPG"
            oStream.SaveToFile strPic, adSaveCreateOverWrite
            oStream.Close
            Set Me.Pictures("K" & lKey).OrigPic = LoadPicture(strPic)
            '删除临时文件
            If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
            'PicMarks
            Set Me.Pictures("K" & lKey).PicMarks = New cPicMarks
            Set oSubNode2 = oSubNode1.selectSingleNode("PicMarks")
            For Each oSubNode3 In oSubNode2.childNodes
                lKey2 = Me.Pictures("K" & lKey).PicMarks.Add()
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).类型 = GetNodeValue(oSubNode3, "类型", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).内容 = GetNodeValue(oSubNode3, "内容", "")
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).点集 = GetNodeValue(oSubNode3, "点集", "")
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).X1 = GetNodeValue(oSubNode3, "X1", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).Y1 = GetNodeValue(oSubNode3, "Y1", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).X2 = GetNodeValue(oSubNode3, "X2", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).Y2 = GetNodeValue(oSubNode3, "Y2", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).填充色 = GetNodeValue(oSubNode3, "填充色", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).填充方式 = GetNodeValue(oSubNode3, "填充方式", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).线条色 = GetNodeValue(oSubNode3, "线条色", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).字体色 = GetNodeValue(oSubNode3, "字体色", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).线型 = GetNodeValue(oSubNode3, "线型", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).线宽 = GetNodeValue(oSubNode3, "线宽", 0)
                Me.Pictures("K" & lKey).PicMarks("K" & lKey2).字体.Name = GetNodeValue(oSubNode3, "字体", "宋体")
            Next
        End If
    Next
    'Tables:
    Set oNode = oRoot.selectSingleNode("Tables")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "Key", 0)
        If lKey > 0 Then
            Me.Tables.Add lKey
            Me.Tables("K" & lKey).ID = GetNodeValue(oSubNode1, "ID", 0)
            Me.Tables("K" & lKey).文件ID = GetNodeValue(oSubNode1, "文件ID", 0)
            Me.Tables("K" & lKey).开始版 = GetNodeValue(oSubNode1, "开始版", 1)
            Me.Tables("K" & lKey).终止版 = GetNodeValue(oSubNode1, "终止版", 0)
            Me.Tables("K" & lKey).父ID = GetNodeValue(oSubNode1, "父ID", 0)
            Me.Tables("K" & lKey).对象序号 = GetNodeValue(oSubNode1, "对象序号", 0)
            Me.Tables("K" & lKey).保留对象 = GetNodeValue(oSubNode1, "保留对象", 0)
            Me.Tables("K" & lKey).是否换行 = GetNodeValue(oSubNode1, "是否换行", 0)
            Me.Tables("K" & lKey).预制提纲ID = GetNodeValue(oSubNode1, "预制提纲ID", 0)
            Me.Tables("K" & lKey).对象属性 = GetNodeValue(oSubNode1, "对象属性", "")
            Me.Tables("K" & lKey).Rows = GetNodeValue(oSubNode1, "Rows", 0)
            Me.Tables("K" & lKey).Cols = GetNodeValue(oSubNode1, "Cols", 0)
            Me.Tables("K" & lKey).Key = GetNodeValue(oSubNode1, "Key", 0)
            'Cells
            Set oSubNode2 = oSubNode1.selectSingleNode("Cells")
            For Each oSubNode3 In oSubNode2.childNodes
                lKey2 = GetNodeValue(oSubNode3, "Key", 0)
                Me.Tables("K" & lKey).Cells.Add lKey2
                Me.Tables("K" & lKey).Cells("K" & lKey2).ID = GetNodeValue(oSubNode3, "ID", 0)
                Me.Tables("K" & lKey).Cells("K" & lKey2).文件ID = GetNodeValue(oSubNode3, "文件ID", 0)
                Me.Tables("K" & lKey).Cells("K" & lKey2).父ID = GetNodeValue(oSubNode3, "父ID", 0)
                Me.Tables("K" & lKey).Cells("K" & lKey2).对象序号 = GetNodeValue(oSubNode3, "对象序号", 0)
                Me.Tables("K" & lKey).Cells("K" & lKey2).内容文本 = GetNodeValue(oSubNode3, "内容文本", 0)
                Me.Tables("K" & lKey).Cells("K" & lKey2).对象属性 = GetNodeValue(oSubNode3, "对象属性", 0)
            Next
            'Table -> Pictures
            Set oSubNode2 = oSubNode1.selectSingleNode("Pictures")
            For Each oSubNode3 In oSubNode2.childNodes
                lKey2 = GetNodeValue(oSubNode3, "Key", 0)
                If lKey2 > 0 Then
                    Me.Tables("K" & lKey).Pictures.Add lKey2
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).ID = GetNodeValue(oSubNode3, "ID", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).文件ID = GetNodeValue(oSubNode3, "文件ID", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).开始版 = GetNodeValue(oSubNode3, "开始版", 1)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).终止版 = GetNodeValue(oSubNode3, "终止版", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).父ID = GetNodeValue(oSubNode3, "父ID", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).对象序号 = GetNodeValue(oSubNode3, "对象序号", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).保留对象 = GetNodeValue(oSubNode3, "保留对象", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).是否换行 = GetNodeValue(oSubNode3, "是否换行", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).Row = GetNodeValue(oSubNode3, "Row", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).Col = GetNodeValue(oSubNode3, "Col", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).PictureType = GetNodeValue(oSubNode3, "PictureType", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).MarkedPicID = GetNodeValue(oSubNode3, "MarkedPicID", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).OrigWidth = GetNodeValue(oSubNode3, "OrigWidth", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).OrigHeight = GetNodeValue(oSubNode3, "OrigHeight", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).Width = GetNodeValue(oSubNode3, "Width", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).Height = GetNodeValue(oSubNode3, "Height", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).ZoomFactor = GetNodeValue(oSubNode3, "ZoomFactor", 0)
                    Me.Tables("K" & lKey).Pictures("K" & lKey2).Modified = GetNodeValue(oSubNode3, "Modified", 0)
                    
                    Set oSubNode4 = oSubNode3.selectSingleNode("OrigPic")
                    oStream.Type = adTypeBinary
                    oStream.Open
                    oStream.Write oSubNode4.nodeTypedValue
                    strPic = strPath & "\XML2JPG" & App.hInstance & ".JPG"
                    oStream.SaveToFile strPic, adSaveCreateOverWrite
                    oStream.Close
                    Set Me.Tables("K" & lKey).Pictures("K" & lKey2).OrigPic = LoadPicture(strPic)
                    '删除临时文件
                    If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
                    
                    'PicMarks
                    Set Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks = New cPicMarks
                    Set oSubNode4 = oSubNode3.selectSingleNode("PicMarks")
                    For Each oSubNode5 In oSubNode4.childNodes
                        lKey3 = Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks.Add()
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).类型 = GetNodeValue(oSubNode5, "类型", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).内容 = GetNodeValue(oSubNode5, "内容", "")
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).点集 = GetNodeValue(oSubNode5, "点集", "")
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).X1 = GetNodeValue(oSubNode5, "X1", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).Y1 = GetNodeValue(oSubNode5, "Y1", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).X2 = GetNodeValue(oSubNode5, "X2", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).Y2 = GetNodeValue(oSubNode5, "Y2", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).填充色 = GetNodeValue(oSubNode5, "填充色", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).填充方式 = GetNodeValue(oSubNode5, "填充方式", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).线条色 = GetNodeValue(oSubNode5, "线条色", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).字体色 = GetNodeValue(oSubNode5, "字体色", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).线型 = GetNodeValue(oSubNode5, "线型", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).线宽 = GetNodeValue(oSubNode5, "线宽", 0)
                        Me.Tables("K" & lKey).Pictures("K" & lKey2).PicMarks("K" & lKey3).字体.Name = GetNodeValue(oSubNode5, "字体", "宋体")
                    Next
                End If
            Next
            'Table -> Elements
            Set oSubNode2 = oSubNode1.selectSingleNode("Elements")
            For Each oSubNode3 In oSubNode2.childNodes
                lKey2 = GetNodeValue(oSubNode3, "Key", 0)
                If lKey2 > 0 Then
                    Me.Tables("K" & lKey).Elements.Add lKey2
                    Me.Tables("K" & lKey).Elements("K" & lKey2).ID = GetNodeValue(oSubNode3, "ID", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).文件ID = GetNodeValue(oSubNode3, "文件ID", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).开始版 = GetNodeValue(oSubNode3, "开始版", 1)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).终止版 = GetNodeValue(oSubNode3, "终止版", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).父ID = GetNodeValue(oSubNode3, "父ID", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).对象序号 = GetNodeValue(oSubNode3, "对象序号", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).保留对象 = GetNodeValue(oSubNode3, "保留对象", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).内容文本 = GetNodeValue(oSubNode3, "内容文本", "")
                    Me.Tables("K" & lKey).Elements("K" & lKey2).是否换行 = GetNodeValue(oSubNode3, "是否换行", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).诊治要素ID = GetNodeValue(oSubNode3, "诊治要素ID", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).替换域 = GetNodeValue(oSubNode3, "替换域", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).要素名称 = GetNodeValue(oSubNode3, "要素名称", "")
                    Me.Tables("K" & lKey).Elements("K" & lKey2).要素类型 = GetNodeValue(oSubNode3, "要素类型", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).要素长度 = GetNodeValue(oSubNode3, "要素长度", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).要素小数 = GetNodeValue(oSubNode3, "要素小数", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).要素单位 = GetNodeValue(oSubNode3, "要素单位", "")
                    Me.Tables("K" & lKey).Elements("K" & lKey2).要素表示 = GetNodeValue(oSubNode3, "要素表示", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).输入形态 = GetNodeValue(oSubNode3, "输入形态", 0)
                    Me.Tables("K" & lKey).Elements("K" & lKey2).要素值域 = GetNodeValue(oSubNode3, "要素值域", "")
                    Me.Tables("K" & lKey).Elements("K" & lKey2).对象属性 = GetNodeValue(oSubNode3, "对象属性", "")
                End If
            Next
            '修正表格宽度
            For i = 1 To Me.Tables("K" & lKey).Cols
                If i = 1 Then
                    Me.Tables("K" & lKey).ColWidthString = Me.Tables("K" & lKey).Cell(1, i).Width
                Else
                    Me.Tables("K" & lKey).ColWidthString = Me.Tables("K" & lKey).ColWidthString & "|" & Me.Tables("K" & lKey).Cell(1, i).Width
                End If
            Next
        End If
    Next
    'RTF文本
    Set oNode = oRoot.selectSingleNode("Content")
    Set oSubNode1 = oNode.selectSingleNode("RTF")
    edtThis.ForceEdit = True
    edtThis.Freeze
    edtThis.TextRTF = oSubNode1.Text
    edtThis.LockAllOLEObjectSize
    
    If blnForUndoRedo = False Then
        '作后期处理
        AfterOpeningRTFText edtThis
        '清除所有的签名对象
        For i = Me.Signs.Count To 1 Step -1
            lKey = Me.Signs(i).Key
            Me.Signs(i).DeleteFromEditor edtThis
            Me.Signs.Remove "K" & lKey
        Next
        
        '如果调入的是审核后的文本，那么变为清洁模式！
        If eEditType = cprET_单病历审核 Then
            edtThis.SelectAll
            edtThis.AuditMode = True
            edtThis.AcceptAuditText
            edtThis.AuditMode = False
        End If
        Call ClearAllIDs        '必须清除所有ID，否则保存时会影响到XML源文件的病历文件内容
    End If
    
    Screen.MousePointer = 0
    edtThis.Tag = ""
    edtThis.Range(0, 0).Selected
    edtThis.InProcessing = False
    edtThis.Modified = True
    edtThis.ForceEdit = False
    edtThis.UnFreeze
    
    Set frmEditor.Document = Me
    ImportFromXMLFile = True
End Function

'################################################################################################################
'## 功能：  获取一个节点的值
'##
'## 参数：  CurNode         :   当前节点对象
'##         SubNodeName     :   子节点名称
'##         DefaultValue    :   默认值
'################################################################################################################
Private Function GetNodeValue(ByVal CurNode As IXMLDOMNode, _
    ByVal SubNodeName As String, _
    Optional ByVal DefaultValue As String = "") As String
    
    On Error Resume Next
    Dim NodeTMP As IXMLDOMNode
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
    
    If InStr(GetNodeValue, vbCr) > 0 And InStr(GetNodeValue, vbLf) = 0 Then '只有回车符无换行符
        GetNodeValue = Replace(GetNodeValue, vbCr, vbCrLf)
    ElseIf InStr(GetNodeValue, vbLf) > 0 And InStr(GetNodeValue, vbCr) = 0 Then '只有换行符无回车符
        GetNodeValue = Replace(GetNodeValue, vbLf, vbCrLf)
    End If
End Function

'################################################################################################################
'## 功能：  创建一个XML节点并赋值
'##
'## 参数：  TabNumber   :   缩进层次数（表示有多少个Tab制表符，便于阅读）
'##         Parent      :   父节点
'##         Node_Type   :   节点类型（目前支持 NODE_ELEMENT 、NODE_CDATA_SECTION 、NODE_COMMENT 、NODE_ATTRIBUTE等）
'##         Node_Name   :   节点名称
'##         Node_Value  :   节点文本
'################################################################################################################
Private Function CreateNode(ByVal TabNumber As Integer, _
    ByVal Parent As IXMLDOMNode, _
    Optional ByVal node_name As String, _
    Optional ByVal Node_Type As tagDOMNodeType = NODE_ELEMENT, _
    Optional ByVal Node_Value As String = "")
    Dim New_Node As IXMLDOMNode
    
    '字符缩进值设置（不影响数据），只影响阅读美观度
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))   '创建文本节点
    '创建新节点
    Set New_Node = Parent.ownerDocument.CreateNode(Node_Type, node_name, "")
    '设置文本值
    New_Node.Text = Node_Value
    '添加到父节点
    Parent.appendChild New_Node
    '添加末尾回车（不影响数据），只影响阅读美观度
    'Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf)   '创建文本节点
    Set CreateNode = New_Node
End Function

Public Function GetDelCharColor(ByRef COLOR As OLE_COLOR) As OLE_COLOR
    '根据Color和当前版本的颜色分量表获取删除字符的颜色
    Dim r As Long, g As Long, b As Long
    If COLOR = tomAutoColor Or COLOR = tomUndefined Then COLOR = vbBlack
    r = 255
    g = rgbGreen(COLOR)
    b = GetColorVectorB(Me.目标版本)
    GetDelCharColor = RGB(r, g, b)
End Function

Public Function GetNewCharColor(ByRef COLOR As OLE_COLOR) As OLE_COLOR
    '根据Color和当前版本的颜色分量表获取删除字符的颜色
    Dim r As Long, g As Long, b As Long
    If COLOR = tomAutoColor Or COLOR = tomUndefined Then COLOR = vbBlack
    r = 255
    g = GetColorVectorG(Me.目标版本)
    b = rgbBlue(COLOR)
    GetNewCharColor = RGB(r, g, b)
End Function

Public Function IsDelCharColor(ByVal COLOR As OLE_COLOR) As Boolean
    '是否是本版本的删除文本色
    IsDelCharColor = (rgbBlue(COLOR) = GetColorVectorB(Me.目标版本))
End Function

Public Function IsNewCharColor(ByVal COLOR As OLE_COLOR) As Boolean
    '是否是本版本的新增文本色
    IsNewCharColor = (rgbGreen(COLOR) = GetColorVectorG(Me.目标版本))
End Function

Public Function 目标版本() As Long
    If Me.EditType <> cprET_单病历审核 Then
        目标版本 = 1
    Else
        目标版本 = Me.EPRPatiRecInfo.最后版本 + IIf(Me.EPRPatiRecInfo.签名级别 = cprSL_空白, 0, 1)
    End If
End Function
Private Function GetTextContent(ByVal strContent As String) As String
'将文本内容中包含的关键字去掉并返回
Dim i As Integer, lKey As Long, strReplace
    For i = 1 To Me.Tables.Count
        lKey = Me.Tables(i).Key
        strReplace = "TS(" & Format(lKey, "00000000") & ",1,0)" & "□"
        strContent = Replace(strContent, strReplace, "")
        
        strReplace = "TE(" & Format(lKey, "00000000") & ",1,0)"
        strContent = Replace(strContent, strReplace, "")
    Next
    
    For i = 1 To Me.Pictures.Count
        lKey = Me.Pictures(i).Key
        strReplace = "PS(" & Format(lKey, "00000000") & ",1,0)" & "□"
        strContent = Replace(strContent, strReplace, "")
        
        strReplace = "PE(" & Format(lKey, "00000000") & ",1,0)"
        strContent = Replace(strContent, strReplace, "")
    Next
    
    For i = 1 To Me.Compends.Count
        lKey = Me.Compends(i).Key
        strReplace = "OS(" & Format(lKey, "00000000") & "," & IIf(Me.Compends(i).保留对象, "1", "0") & ",0)"
        strContent = Replace(strContent, strReplace, "【" & Me.Compends(i).名称 & "】")
        
        strReplace = "OE(" & Format(lKey, "00000000") & "," & IIf(Me.Compends(i).保留对象, "1", "0") & ",0)"
        strContent = Replace(strContent, strReplace, "")
    Next
    
    For i = 1 To Me.Elements.Count
        lKey = Me.Elements(i).Key
        strReplace = "ES(" & Format(lKey, "00000000") & "," & IIf(Me.Elements(i).保留对象, "1", "0") & ",0)"
        strContent = Replace(strContent, strReplace, "")
        
        strReplace = "EE(" & Format(lKey, "00000000") & "," & IIf(Me.Elements(i).保留对象, "1", "0") & ",0)"
        strContent = Replace(strContent, strReplace, "")
    Next
    
    For i = 1 To Me.Diagnosises.Count
        lKey = Me.Diagnosises(i).Key
        strReplace = "DS(" & Format(lKey, "00000000") & ",0,0)"
        strContent = Replace(strContent, strReplace, "")
        
        strReplace = "DE(" & Format(lKey, "00000000") & ",0,0)"
        strContent = Replace(strContent, strReplace, "")
    Next
    
    For i = 1 To Me.Signs.Count
        lKey = Me.Signs(i).Key
        strReplace = "SS(" & Format(lKey, "00000000") & ",0,0)"
        strContent = Replace(strContent, strReplace, "")
        
        strReplace = "SE(" & Format(lKey, "00000000") & ",0,0)"
        strContent = Replace(strContent, strReplace, "")
    Next
    strContent = Replace(strContent, "'", """")
    strContent = Replace(strContent, ",", "，")
    strContent = Replace(strContent, "(", "（")
    strContent = Replace(strContent, ")", "）")
    GetTextContent = strContent
End Function
