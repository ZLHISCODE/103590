VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTableEPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum TabEM
    TabEm_新增 = 0              '新增
    TabEm_修改 = 1              '修改
End Enum

'当前编辑器的编辑类型
Public Enum TabET
    TabET_病历文件定义 = 0      '病历文件定义
    TabET_全文示范编辑 = 1      '全文示范编辑
    TabET_单病历编辑 = 2        '单病历编辑
    TabET_单病历审核 = 3        '单病历审核
End Enum

'文件种类  1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-诊断文书;6-知情文件;7-诊疗报告;8-诊疗申请
Public Enum EPRDocType
    Tab门诊病历 = 1
    Tab住院病历 = 2
    Tab护理记录 = 3
    Tab护理病历 = 4
    Tab诊断文书 = 5
    Tab知情文件 = 6
    Tab诊疗报告 = 7             '诊疗单据：报告
    Tab诊疗申请 = 8             '诊疗单据：申请
End Enum
'病人来源
Public Enum PatiFrom
    TabPF_门诊 = 1              '1-门诊；
    TabPF_住院 = 2              '2-住院；
    TabPF_外来 = 3              '3-外来；
    TabPF_体检 = 4              '4-体检
End Enum

'签名状态
Public Enum EPRSignLevel
    TabSL_空白 = 0              '未签名
    TabSL_经治 = 1              '经治医师签名
    TabSL_主治 = 2              '主治医师签名
    TabSL_主任 = 3              '主任医师签名
    TabSL_正高 = 4              '正高：签名级别不包含，只表示人员居右正高职称，以便区别副主任医师
End Enum
Public 签名级别 As EPRSignLevel             '用户签名级别，判断用户是否有权限审签

Public Cells As cTabCells                      '表格集合
Public Pictures As cTabPictures                 '图片集合
Public Texts As cTabTexts
Public Elements As cTabElements                 '诊治要素集合
Public Signs As cTabSigns                       '签名集合
Public PicMarks As cTabPicMarks                    '图片标记
Public EPRFileInfo As cTabFileDefineInfo        '病历文件定义信息
Public EPRDemoInfo As cTabDemoInfo              '病历示范信息
Public EPRPatiRecInfo As cTabPatiRecordInfo     '病人电子病历信息
Public mfrmParent As Object                         '父窗体
Public frmEditor As Object                          '主编辑窗体
Private WithEvents mfrmPrint As frmPrintPreview
Attribute mfrmPrint.VB_VarHelpID = -1
Private mfrmDockView As frmDockView

Public EM As TabEM                     '编辑模式（新增、修改）
Public ET As TabET                     '编辑类型（病历文件定义、病历示范编辑、单病历编辑、单病历审核）
Public mblnMove As Boolean                          '数据是否转移
Public mReadOnly As Byte                         '是否只读,0-正常,1-签名后点修改,2-主界面打开查阅或查阅历次签名版本
Public mMaxNo As Long                            '最大对象序号
Public Sub InitTableEPR(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strDbOwner As String)
    'InitCommon cnMain '正式用时取掉本句,
    
    
    Set gcnOracle = cnMain
    glngSys = lngSys
    gstrDbOwner = strDbOwner
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("产品名称"), Default:="中联")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    Set gAllFont = GetAllFonts
    
    Call GetUserInfo
    InitPreDefinedKeys
End Sub
Public Sub InitOpenEPR(ByVal frmParent As Object, ByVal EditMode As TabEM, ByVal editType As TabET, ByVal lngFileID As Long, _
    Optional ByVal blnShow As Boolean = True, Optional ByVal ReadOnly As Byte, Optional ByVal PatiFrom As PatiFrom, _
    Optional ByVal lngPatiID As Long, Optional ByVal lngPageID As Long, Optional ByVal intBabyOrder As Integer, _
    Optional ByVal intCurDeptID As Long, Optional ByVal lngAdviceID As Long, Optional ByVal strModelPrivate As String, _
    Optional ByVal blnMoved As Boolean, Optional ByVal blnCanPrint As Boolean = True, Optional ByVal bytEsign As Byte, _
	Optional ByVal intStyle As Integer)
'## 参数：  frmParent       :父窗体
'##         eEdtMode    :当前编辑模式（新增、修改）
'##         eEdtType    :当前编辑方式（文件定义、示范编辑、单病历编辑、单病历审核）
'##         lngFileID   :文件ID（根据编辑方式的不同，可以表示文件定义ID、范文ID或者病人病历ID）
'##
'##         (以下为可选参数，用于病人电子病历编辑和审核):
'##         blnShow     :显示编辑器,并由编辑器回调本类中相关函数读取数据,默认是,否的情况主要用在存范文
'##         ReadOnly    :是否只读,可能的情况为0-正常,1-签名后点修改,2-主界面打开查阅或查阅历次签名版本
'##         PatiFrom    :病人来源
'##         lngPatiID   :病人ID
'##         lngPageID  :住院次数，主页ID；门诊则为挂号ID
'##         intBabyOrder:婴儿序号
'##         intCurDeptID:当前病人病历填写科室ID
'##         lngAdviceID      :医嘱ID，用于特殊病历
'##         strModelPrivate :调用模块权限
'##         blnCanPrint     :是否允许预览、打印
Dim rs As New ADODB.Recordset, lngTmp As Long
    On Error GoTo errHand
    zlCommFun.ShowFlash "正要打开文件，请稍等！", frmParent
    
    gbytEsign = bytEsign
    Me.EPRPatiRecInfo.医嘱id = lngAdviceID
    Me.EM = EditMode      '新增、修改；
    Me.ET = editType      '病历文件定义、全文示范编辑、单病历编辑、单病历审核
    Me.mblnMove = blnMoved
    Me.mReadOnly = ReadOnly
    
    Select Case editType
    Case TabET_病历文件定义
        Call EPRFileInfo.GetFileDefineInfo(lngFileID)
    Case TabET_全文示范编辑
        Call EPRDemoInfo.GetDemoInfo(lngFileID)
        Call EPRFileInfo.GetFileDefineInfo(EPRDemoInfo.文件ID)
    Case TabET_单病历编辑, TabET_单病历审核
        If EditMode = TabEm_修改 Or editType = TabET_单病历审核 Then
            Call EPRPatiRecInfo.GetPatiRecordInfo(lngFileID, blnMoved)
            Call EPRFileInfo.GetFileDefineInfo(EPRPatiRecInfo.文件ID)
        Else                '新增电子病历（第1版本）
            Call EPRFileInfo.GetFileDefineInfo(lngFileID)
            With EPRPatiRecInfo
                .ID = 0                         '新增记录，ID为0
                .病人来源 = PatiFrom
                .病人ID = lngPatiID
                .主页ID = lngPageID
                .婴儿 = intBabyOrder
                .科室ID = intCurDeptID
                .病历种类 = EPRFileInfo.种类    '病历种类
                .文件ID = EPRFileInfo.ID        '文件ID
                .病历名称 = EPRFileInfo.名称    '病历名称
                .创建人 = UserInfo.姓名
                .保存人 = UserInfo.姓名
                .签名级别 = TabSL_空白          '当前签名级别为空白(0)，表示未签名
                .最后版本 = 1
            End With
        End If
        Me.签名级别 = GetUserSignLevel(UserInfo.ID, UserInfo.姓名, lngPatiID, lngPageID)       '用户的签名级别
        If Me.EPRPatiRecInfo.病历种类 = Tab护理病历 And Me.签名级别 = TabSL_主治 Then Me.签名级别 = TabSL_主任 '对于护理病历，只有2级签名（护士、护士长）
    End Select

    If blnShow Then
        Set mfrmParent = frmParent
        frmEditor.ShowMe mfrmParent, Me, strModelPrivate, blnMoved, blnCanPrint, intStyle '显示编辑器
    End If
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub InitEmptyStructure()
Dim lngRow As Long, lngRows As Long, lngCol As Long, lngCols As Long, lngWidth As Long, lngHeight As Long
    On Error GoTo errHand
    '清空各种类
    Call Me.Cells.RemoveAll
    Call Me.Elements.RemoveAll
    Call Me.Pictures.RemoveAll
    Call Me.Signs.RemoveAll
    Call Me.Texts.RemoveAll
    Call Me.PicMarks.RemoveAll
    
    '初始化单元格
    With Me.Cells
        .Rows = 30: .Cols = 8: lngRows = 30: lngCols = 8
        For lngRow = 1 To lngRows
            For lngCol = 1 To lngCols
                .Add lngRow, lngCol
                With .Cell(lngRow, lngCol)
                    .Width = 1400
                    .Height = IIf(lngRow = 1, 1000, 480)
                    .FontName = "宋体"
                    .FontSize = IIf(lngRow = 1, 15, 9)
                    .FontBold = IIf(lngRow = 1, True, False)
                    .FontItalic = False
                    .FontUnderline = False
                    .FontStrikeout = False
                    .FontColor = vbBlack
                    .HAlignment = F1HAlignCenter
                    .VAlignment = F1VAlignCenter
                    .CellLineTop = IIf(lngRow = 1, F1BorderNone, F1BorderThin)
                    .CellLineBottom = F1BorderThin
                    .CellLineLeft = IIf(lngRow = 1, F1BorderNone, F1BorderThin)
                    .CellLineRight = IIf(lngRow = 1, F1BorderNone, F1BorderThin)
                    .CellLineTopColor = vbBlack
                    .CellLineBottomColor = vbBlack
                    .CellLineLeftColor = vbBlack
                    .CellLineRightColor = vbBlack
                    .Merge = IIf(lngRow = 1, True, False)                  '是否合并
                    .MergeRange = IIf(lngRow = 1 And lngCol = 1, "1,1;1," & lngCols, lngRow & "," & lngCol)
                    .对象序号 = .Index
                    .对象类型 = 0
                    .保留对象 = IIf(lngRow = 1, True, False)
                    .对象属性 = ""
                    .内容行次 = 0
                    .内容文本 = IIf(lngRow = 1, Me.EPRFileInfo.名称, "")
                    .开始版 = 1
                    .终止版 = 0
                End With
            Next
        Next
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function AnalyseFileStrucuture(ByVal strFile As String, Optional blnImport As Boolean) As Boolean
'功能：将XML文件解悉到CELL类中
Dim objFileXML As New MSXML2.DOMDocument, ElementList As MSXML2.IXMLDOMNodeList
Dim oSubNodes As MSXML2.IXMLDOMNodeList, oSubNode As MSXML2.IXMLDOMNode
Dim oStream As New ADODB.Stream, strFileTmp As String, i As Integer, lngKey As Long              ' 流对象           临时文件
Dim lngRows As Long, lngCols As Long, lngRow As Long, lngCol As Long, lngWidth As Long, lngHeight As Long
Dim lngTemp As Long, strHeight As String, strWidth As String, strKey As String
    
    On Error GoTo errHand
    
    If objFileXML.Load(strFile) = False Then
        If blnImport Then
            MsgBox "XML文档格式不正常,不能导入,请检查!", vbInformation, gstrSysName
            Exit Function
        Else
            GoTo errNew        '读出XML文档，出错则按空文档创建
        End If
    End If
    
    '清空各种类
    Call Me.Cells.RemoveAll
    Call Me.Elements.RemoveAll
    Call Me.Pictures.RemoveAll
    Call Me.Signs.RemoveAll
    Call Me.Texts.RemoveAll
    Call Me.PicMarks.RemoveAll

        
    If blnImport Then   '导入XML时解释文档页面
        With Me.EPRFileInfo
            .PaperKind = GetElemnetValue(objFileXML, "纸张", "纸型")
            .PaperWidth = GetElemnetValue(objFileXML, "纸张", "宽度") ') '210</宽度>
            .PaperHeight = GetElemnetValue(objFileXML, "纸张", "高度") ') ' >294</高度>
            .MarginTop = GetElemnetValue(objFileXML, "纸张", "上边距") ') ' >3</上边距>
            .MarginBottom = GetElemnetValue(objFileXML, "纸张", "下边距") ') ' >3</下边距>
            .MarginLeft = GetElemnetValue(objFileXML, "纸张", "左边距") ') ' >2.5</左边距>
            .MarginRight = GetElemnetValue(objFileXML, "纸张", "右边距") ') ' >2.5</右边距>
            .PaperOrient = GetElemnetValue(objFileXML, "纸张", "纸向") ')  '>纵向</纸向>
            '页眉
            .HeadConText = Replace(GetElemnetValue(objFileXML, "纸张", "页眉"), "\r\n", vbCrLf) ') ' String          '内容(要素用[]包括),页眉内容
            .HeadFontName = UsableFont(GetElemnetValue(objFileXML, "纸张", "页眉字体"))  ') ' String
            .HeadFontSize = GetElemnetValue(objFileXML, "纸张", "页眉字号") ') ' Long
            .HeadFontBold = GetElemnetValue(objFileXML, "纸张", "页眉粗体") ') ' Boolean
            .HeadFontItalic = GetElemnetValue(objFileXML, "纸张", "页眉斜体") ') ' Boolean
            .HeadFontUnderline = GetElemnetValue(objFileXML, "纸张", "页眉下划线") ') ' Boolean
            .HeadFontStrikethrough = GetElemnetValue(objFileXML, "纸张", "页眉删除线") ') ' Boolean
            .HeadFontColor = GetElemnetValue(objFileXML, "纸张", "页眉字色") ') ' OLE_COLOR
            .HeadMargin = GetElemnetValue(objFileXML, "纸张", "页眉边距")  ') ' 边距
            ''页脚                  '
            .FootConText = Replace(GetElemnetValue(objFileXML, "纸张", "页脚"), "\r\n", vbCrLf) ') ' String               '内容(要素用[]包括),页脚内容
            .FootFontName = UsableFont(GetElemnetValue(objFileXML, "纸张", "页脚字体")) ') ' String
            .FootFontSize = GetElemnetValue(objFileXML, "纸张", "页脚字号") ') ' Long
            .FootFontBold = GetElemnetValue(objFileXML, "纸张", "页脚粗体") ') ' Boolean
            .FootFontItalic = GetElemnetValue(objFileXML, "纸张", "页脚斜体") ') ' Boolean
            .FootFontUnderline = GetElemnetValue(objFileXML, "纸张", "页脚下划线") ') ' Boolean
            .FootFontStrikethrough = GetElemnetValue(objFileXML, "纸张", "页脚删除线") ') ' Boolean
            .FootFontColor = GetElemnetValue(objFileXML, "纸张", "页脚字色") ') ' OLE_COLOR
            .FootMargin = GetElemnetValue(objFileXML, "纸张", "页脚边距")  ') ' 边距
            
            Set oSubNode = objFileXML.selectSingleNode("页眉图片")
            If Not oSubNode Is Nothing Then
                oStream.Type = adTypeBinary
                oStream.Open
                oStream.Write oSubNode.nodeTypedValue
                If oStream.Size > 0 Then
                    strFileTmp = App.Path & "\Tmp.Jpg"
                    oStream.SaveToFile strFileTmp, adSaveCreateOverWrite
                    Set .HeadPic = LoadPicture(strFileTmp)
                    If gobjFSO.FileExists(strFileTmp) Then gobjFSO.DeleteFile strFileTmp, True '删除临时文件
                End If
                oStream.Close
            End If
        End With
    End If
    
    '解悉属性
    lngRows = Val(GetElemnetValue(objFileXML, "属性", "行数"))   '读出行数
    lngCols = Val(GetElemnetValue(objFileXML, "属性", "列数"))  '读出列数
    strHeight = GetElemnetValue(objFileXML, "属性", "行高") '读出行高串
    strWidth = GetElemnetValue(objFileXML, "属性", "列宽") '读出列宽串
    Me.Cells.Rows = lngRows: Me.Cells.Cols = lngCols
    For lngRow = 1 To lngRows                                   '循环创建单元格
        lngHeight = Val(Split(strHeight, "|")(lngRow))
        For lngCol = 1 To lngCols
            lngWidth = Val(Split(strWidth, "|")(lngCol))
            strKey = Me.Cells.Add(lngRow, lngCol)
            Me.Cells(strKey).Height = lngHeight
            Me.Cells(strKey).Width = lngWidth
        Next
    Next
    
    '解悉单元格格式
    Dim strRange As String, blnMerge As Boolean, strTmp As String
    Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, l As Long, j As Long
    Set ElementList = objFileXML.getElementsByTagName("Cell") '取出单元格格式
    For lngTemp = 1 To ElementList.Length
        strRange = ElementList.Item(lngTemp - 1).selectSingleNode("区域").Text
        If UBound(Split(strRange, ";")) = 0 Then                                            '非合并区域
            lngRow = Split(strRange, ",")(0)
            lngCol = Split(strRange, ",")(1)
        Else
            lngRow = Split(Split(strRange, ";")(0), ",")(0)
            lngCol = Split(Split(strRange, ";")(0), ",")(1)
        End If
        With Me.Cells.Cell(lngRow, lngCol)
            strTmp = ElementList.Item(lngTemp - 1).selectSingleNode("边框").Text
            .CellLineTop = Split(strTmp, "|")(0)             '<上边框>1</上边框><!--0-无边框 1-有边框-->
            .CellLineBottom = Split(strTmp, "|")(1)            '<下边框>1</下边框>
            .CellLineLeft = Split(strTmp, "|")(2)             '<左边框>1</左边框>
            .CellLineRight = Split(strTmp, "|")(3)           '<右边框>1</右边框>
            .CellLineTopColor = Split(strTmp, "|")(4)           '<上边框颜色>vbBlack</上边框颜色>
            .CellLineBottomColor = Split(strTmp, "|")(5)            '<下边框颜色>vbBlack</下边框颜色>
            .CellLineLeftColor = Split(strTmp, "|")(6)            '左边框颜色>vbBlack</左边框颜色>
            .CellLineRightColor = Split(strTmp, "|")(7)           '<右边框颜色>vbBlack</右边框颜色>
            .Merge = ElementList.Item(lngTemp - 1).selectSingleNode("合并").Text                   '是否合并
            .MergeRange = strRange
            
            If IIf(.Merge, InStr(strRange, ";") > 0, True) Then                                     '合并单元格首个和非合并单元格
                '字体|字号|粗体|斜体|下划线|删除线|字体颜色
                strTmp = ElementList.Item(lngTemp - 1).selectSingleNode("字体").Text
                .FontName = UsableFont(Split(strTmp, "|")(0))                '字体>宋体</字体>
                .FontSize = Split(strTmp, "|")(1)                  '<字号>9</字号>
                .FontBold = Split(strTmp, "|")(2)                  '<粗体>False</粗体>
                .FontItalic = Split(strTmp, "|")(3)               '<斜体>False</斜体>
                .FontUnderline = Split(strTmp, "|")(4)             '<下划线>False</下划线>
                .FontStrikeout = Split(strTmp, "|")(5)            '<删除线>False</删除线>
                .FontColor = Split(strTmp, "|")(6)                   '<字体颜色>vbblack</字体颜色>
                .HAlignment = Split(strTmp, "|")(7)            '<横向对齐>F1HAlignCenter</横向对齐>
                .VAlignment = Split(strTmp, "|")(8)             '<纵向对齐>F1VAlignCenter</纵向对齐>
                
                If blnImport Then   '导入XML时
                    Set oSubNodes = ElementList.Item(lngTemp - 1).selectNodes("内容")
                    .文件ID = Me.EPRFileInfo.ID
                    .对象类型 = oSubNodes.Item(0).selectSingleNode("对象类型").Text
                    .保留对象 = oSubNodes.Item(0).selectSingleNode("保留对象").Text
                    .对象属性 = oSubNodes.Item(0).selectSingleNode("对象属性").Text
                    .内容行次 = 0
                    .内容文本 = Replace(oSubNodes.Item(0).selectSingleNode("内容文本").Text, "\r\n", vbCrLf)
                    If InStr(.对象属性, "|") > 0 Then
                        .对象属性 = Split(.对象属性, "|")(1)
                    Else
                        .对象属性 = ""
                    End If
                    
                    Select Case .对象类型
                        Case cprCTFixtext, cprCTText    '固定文本与普通文本单元格
                        Case cprCTElement               '单要素单元格
                            .ElementKey = Me.Elements.Add
                            With Me.Elements("K" & .ElementKey)
                                .区域 = lngRow & "|" & lngCol
                                .内容文本 = oSubNodes.Item(0).selectSingleNode("要素文本").Text
                                .诊治要素ID = oSubNodes.Item(0).selectSingleNode("诊治要素ID").Text
                                .替换域 = oSubNodes.Item(0).selectSingleNode("替换域").Text
                                .要素名称 = oSubNodes.Item(0).selectSingleNode("要素名称").Text
                                .要素类型 = oSubNodes.Item(0).selectSingleNode("要素类型").Text
                                .要素长度 = oSubNodes.Item(0).selectSingleNode("要素长度").Text
                                .要素小数 = oSubNodes.Item(0).selectSingleNode("要素小数").Text
                                .要素单位 = oSubNodes.Item(0).selectSingleNode("要素单位").Text
                                .要素表示 = oSubNodes.Item(0).selectSingleNode("要素表示").Text
                                .输入形态 = oSubNodes.Item(0).selectSingleNode("输入形态").Text
                                .要素值域 = oSubNodes.Item(0).selectSingleNode("要素值域").Text
                                .保留对象 = oSubNodes.Item(0).selectSingleNode("保留对象").Text
                                .自动转文本 = oSubNodes.Item(0).selectSingleNode("自动转文本").Text
                                .必填 = oSubNodes.Item(0).selectSingleNode("必填").Text
                                If .替换域 = 1 Then .内容文本 = "" '自动替换要素需要导入后重新提取
                                If Not .CheckValid Then .诊治要素ID = 0
                            End With
                        Case cprCTTextElement           '混合单元格
                            For i = 1 To oSubNodes.Length - 1
                                Set oSubNode = oSubNodes.Item(i)
                                If oSubNode.selectSingleNode("对象类型").Text = cprCTText Then
                                    lngKey = Me.Texts.Add
                                    .TextKey = .TextKey & "|" & lngKey
                                    With Me.Texts("K" & lngKey)
'                                        .ID = oSubNode.selectSingleNode("ID").Text
'                                        .父ID = oSubNode.selectSingleNode("父ID").Text
                                        .区域 = lngRow & "|" & lngCol
                                        .保留对象 = oSubNode.selectSingleNode("保留对象").Text
                                        .内容行次 = oSubNode.selectSingleNode("文本行次").Text
                                        .内容文本 = Replace(oSubNode.selectSingleNode("文本内容").Text, "\r\n", vbCrLf)
                                        If .内容行次 = 0 Then .内容行次 = 1
                                    End With
                                Else
                                    lngKey = Me.Elements.Add
                                    .ElementKey = .ElementKey & "|" & lngKey
                                    With Me.Elements("K" & lngKey)
                                        .区域 = lngRow & "|" & lngCol
                                        .内容文本 = Replace(oSubNode.selectSingleNode("要素文本").Text, "\r\n", vbCrLf)
                                        .内容行次 = oSubNode.selectSingleNode("内容行次").Text
                                        .诊治要素ID = oSubNode.selectSingleNode("诊治要素ID").Text
                                        .替换域 = oSubNode.selectSingleNode("替换域").Text
                                        .要素名称 = oSubNode.selectSingleNode("要素名称").Text
                                        .要素类型 = oSubNode.selectSingleNode("要素类型").Text
                                        .要素长度 = oSubNode.selectSingleNode("要素长度").Text
                                        .要素小数 = oSubNode.selectSingleNode("要素小数").Text
                                        .要素单位 = oSubNode.selectSingleNode("要素单位").Text
                                        .要素表示 = oSubNode.selectSingleNode("要素表示").Text
                                        .输入形态 = oSubNode.selectSingleNode("输入形态").Text
                                        .要素值域 = oSubNode.selectSingleNode("要素值域").Text
                                        .保留对象 = oSubNode.selectSingleNode("保留对象").Text
                                        .自动转文本 = oSubNode.selectSingleNode("自动转文本").Text
                                        .必填 = oSubNode.selectSingleNode("必填").Text
                                        If Not .CheckValid Then .诊治要素ID = 0
                                        If .内容行次 = 0 Then .内容行次 = 1
                                    End With
                                End If
                            Next
                        Case cprCTPicture, cprCTReportPic              '参考图
                            .PictureKey = Me.Pictures.Add
                            .内容文本 = IIf(.对象类型 = cprCTPicture, "参考图", "报告图")
                            Set oSubNode = oSubNodes.Item(0).selectSingleNode(IIf(.对象类型 = cprCTPicture, "参考图", "报告图"))
                            If Not oSubNode Is Nothing Then
                                oStream.Type = adTypeBinary
                                oStream.Open
                                oStream.Write oSubNode.nodeTypedValue
                                If oStream.Size > 0 Then
                                    strFileTmp = App.Path & "\Tmp.Jpg"
                                    oStream.SaveToFile strFileTmp, adSaveCreateOverWrite
                                    Set Me.Pictures("K" & .PictureKey).OrigPic = LoadPicture(strFileTmp)
                                    Me.Pictures("K" & .PictureKey).对象属性 = .对象属性
                                    If gobjFSO.FileExists(strFileTmp) Then gobjFSO.DeleteFile strFileTmp, True '删除临时文件
                                End If
                                oStream.Close
                            End If
                            '标记
                            For i = 1 To oSubNodes.Length - 1  '对本单元格中所有标记进行提取
                                Set oSubNode = oSubNodes.Item(i)
                                lngKey = Me.PicMarks.Add
                                .PicMarkKey = .PicMarkKey & "|" & lngKey
                                With Me.PicMarks("K" & lngKey)
'                                    .父ID = oSubNode.selectSingleNode("父ID").Text
                                    .对象属性 = oSubNode.selectSingleNode("对象属性").Text
                                    .保留对象 = oSubNode.selectSingleNode("保留对象").Text
                                End With
                            Next
                        Case cprCTSign, cprCTRowSign, cprCTColSign                '签名
                            .SignKey = Me.Signs.Add
                            If Nvl(.内容文本) <> "" Then '签名在设计时仅为占位,无实际信息
                                Me.Signs("K" & .SignKey).姓名 = Nvl(.内容文本)
                                With Me.Signs("K" & .SignKey)
                                    .签名信息 = oSubNodes.Item(0).selectSingleNode("签名信息").Text
                                    .前置文字 = oSubNodes.Item(0).selectSingleNode("要素名称").Text
                                    .签名级别 = oSubNodes.Item(0).selectSingleNode("要素表示").Text
                                    .时间戳 = oSubNodes.Item(0).selectSingleNode("要素单位").Text
                                    
                                    .签名方式 = Split(oSubNodes.Item(0).selectSingleNode("对象属性").Text, "|")(0)
                                    .签名规则 = Split(oSubNodes.Item(0).selectSingleNode("对象属性").Text, "|")(1)
                                    .证书ID = Split(oSubNodes.Item(0).selectSingleNode("对象属性").Text, "|")(2)
                                    .显示手签 = Split(oSubNodes.Item(0).selectSingleNode("对象属性").Text, "|")(3)
                                    .签名时间 = Split(oSubNodes.Item(0).selectSingleNode("对象属性").Text, "|")(4)
                                    .显示时间 = Split(oSubNodes.Item(0).selectSingleNode("对象属性").Text, "|")(5)
                                End With
                            End If
                    End Select
                    If InStr(strRange, ";") > 0 Then
                        lsRow = Split(Split(strRange, ";")(0), ",")(0): lsCol = Split(Split(strRange, ";")(0), ",")(1)
                        leRow = Split(Split(strRange, ";")(1), ",")(0): leCol = Split(Split(strRange, ";")(1), ",")(1)
                        For l = lsRow To leRow
                            For j = lsCol To leCol
                                Me.Cells.Cell(l, j).对象类型 = .对象类型
                            Next
                        Next
                    End If
                End If
            End If
        End With
    Next
    
    Set ElementList = Nothing
    Set objFileXML = Nothing
    AnalyseFileStrucuture = True
    Exit Function
        
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
errNew:
    MsgBox "解析文件定义结构失败，将初始化空文档", vbInformation, gstrSysName
End Function
Public Function ReadFileStructure() As Boolean
'功能:读取初始文件结构定义,病历定义新增,病历定义修改,病历范文新增,病历编辑新增时调用
'返回:成功读取文件返回真；未能读取文件可能是定义丢失或病历定义新增,调用初始化空文档函数
'IntAction  0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
Dim strTemp As String, strZipFile As String, IntAction As Integer
    On Error GoTo errHand
    If Me.EM = TabEm_新增 Then
        strZipFile = zlBlobRead(1, EPRFileInfo.ID)
    Else
        Select Case Me.ET
            Case TabET_病历文件定义
                strZipFile = zlBlobRead(1, EPRFileInfo.ID)
            Case TabET_全文示范编辑
                strZipFile = zlBlobRead(3, EPRDemoInfo.ID)
                If strZipFile = "" Then
                    strZipFile = zlBlobRead(1, EPRFileInfo.ID)
                End If
            Case TabET_单病历编辑, TabET_单病历审核
                strZipFile = zlBlobRead(5, EPRPatiRecInfo.ID, "", mblnMove)
        End Select
    End If
    
    If strZipFile = "" Then '可能是病历文件定义时新增，或其它情况文件丢失
        Exit Function
    End If
    
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        gobjFSO.DeleteFile strZipFile, True
        If strTemp = "" Then Exit Function
        If Not gobjFSO.FileExists(strTemp) Then
            Exit Function
        End If
        
        If Not AnalyseFileStrucuture(strTemp) Then '解悉XML文件
            gobjFSO.DeleteFile strTemp, True
            Exit Function
        Else
            gobjFSO.DeleteFile strTemp, True
        End If
    Else
        Exit Function
    End If
    
    ReadFileStructure = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub ReadFileContent(ByVal blnMoved As Boolean)
'功能：读取文件内容
'参数：blnMoved 数据是否被转储
Dim rs As New ADODB.Recordset, strTableName As String, lngFileID As Long
    On Error GoTo errHand
    
    If Me.EM = TabEm_新增 Then
        strTableName = "病历文件结构"
        lngFileID = Me.EPRFileInfo.ID
    Else
        Select Case Me.ET
        Case TabET_病历文件定义
            strTableName = "病历文件结构"
            lngFileID = Me.EPRFileInfo.ID           '取文件定义ID
        Case TabET_全文示范编辑
            strTableName = "病历范文内容"
            lngFileID = Me.EPRDemoInfo.ID
            gstrSQL = "Select Count(ID) 数量 From 病历范文内容 Where 文件ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, App.ProductName, lngFileID)
            If rs.EOF Or rs!数量 = 0 Then
                strTableName = "病历文件结构"
                lngFileID = Me.EPRFileInfo.ID
            End If
        Case TabET_单病历编辑, TabET_单病历审核
            strTableName = IIf(blnMoved, "H电子病历内容", "电子病历内容")
            lngFileID = Me.EPRPatiRecInfo.ID
        End Select
    End If
    
    If InStr(strTableName, "电子病历内容") > 0 Then
        gstrSQL = "Select *" & vbNewLine & _
                    "From ( --提取行列签" & vbNewLine & _
                    "       Select *" & vbNewLine & _
                    "       From " & strTableName & "" & vbNewLine & _
                    "       Where 文件id = [1] And 对象类型 In (7, 8)" & vbNewLine & _
                    "       Union" & vbNewLine
                    
        If mReadOnly <> 2 Then
            gstrSQL = gstrSQL & "--除签名以外的单元格最新版" & vbNewLine & _
                            "       Select *" & vbNewLine & _
                            "       From " & strTableName & "" & vbNewLine & _
                            "       Where 文件id = [1] And 对象类型 <> 6 And 终止版 = 0" & vbNewLine & _
                            "       Union" & vbNewLine
        End If
        
        gstrSQL = gstrSQL & "       --指定版本" & vbNewLine & _
                    "       Select a.*" & vbNewLine & _
                    "       From " & strTableName & " A," & vbNewLine & _
                    "            (Select Substr(对象属性, 1, Instr(对象属性, '|') - 1) 行列, Max(终止版) 终止版" & vbNewLine & _
                    "              From " & strTableName & "" & vbNewLine & _
                    "              Where 文件id = [1] And 终止版 <= [2]" & IIf(mReadOnly <> 2, " And 对象类型 = 6", "") & vbNewLine & _
                    "              Group By Substr(对象属性, 1, Instr(对象属性, '|') - 1)) B" & vbNewLine & _
                    "       Where a.文件id = [1] And Instr(a.对象属性, b.行列) > 0 And a.终止版 = b.终止版)" & vbNewLine & _
                    "Order By 对象序号, 内容行次"
         '图片及图标，只有一个版本，终止版=0;文本、要素、混辑每签名后修改则终止上一版，新增终止版=0;行、列签只有一行记录，没签名时终止版=0,签名时为当前最后版
         '同一单元格重复签名为初始版＝最后版＝终止版;A单元格签:开始版=1,终止版=1，B单元格签:开始版=1,终止版=2,A再签开始版=2,终止版=3
    Else
        gstrSQL = "Select * From " & strTableName & " Where 文件id = [1] Order By 对象序号, 内容行次"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", lngFileID, EPRPatiRecInfo.最后版本)
    
    
    Dim lngRow As Long, lngCol As Long, strRange As String, lngKey As Long, lngTmp As Long '读取各类对象内容
    Dim lsRow As Long, lsCol As Long, leRow As Long, leCol As Long, l As Long, j As Long, strMarkKeys As String
    Do Until rs.EOF
        lngKey = 0
        strRange = Split(rs!对象属性, "|")(0)
        lngRow = Split(strRange, ",")(0): lngCol = Split(strRange, ",")(1)
        With Me.Cells.Cell(lngRow, lngCol)
            Select Case rs!对象类型
                Case cprCTFixtext          '0-固定文本(不可编辑)
                        .ID = rs!ID
                        .文件ID = rs!文件ID
                        .对象类型 = rs!对象类型
                        .保留对象 = True
                        .对象属性 = Split(rs!对象属性, "|")(1)
                        .内容行次 = 0
                        .内容文本 = Nvl(rs!内容文本)
                        .对象序号 = Nvl(rs!对象序号, 0)
                        If InStr(strTableName, "电子病历内容") > 0 Then
                            .开始版 = Nvl(rs!开始版, 1)
                            .终止版 = Nvl(rs!终止版, 0)
                        End If
                Case cprCTText            '1-文本型(可编辑多行文本)
                        If Nvl(rs!父ID, 0) <> 0 Then '混合单元格的文本子项
                            lngKey = Me.Texts.Add
                            .TextKey = .TextKey & "|" & lngKey
                            With Me.Texts("K" & lngKey)
                                .ID = rs!ID
                                .父ID = rs!父ID
                                .区域 = lngRow & "|" & lngCol
                                .保留对象 = Nvl(rs!保留对象, 0) = 1
                                .内容文本 = Nvl(rs!内容文本)
                                .内容行次 = Nvl(rs!内容行次, 1)
                                If InStr(strTableName, "电子病历内容") > 0 Then
                                    .开始版 = Nvl(rs!开始版, 1)
                                    .终止版 = Nvl(rs!终止版, 0)
                                End If
                            End With
                        Else
                            .ID = rs!ID
                            .文件ID = rs!文件ID
                            .对象类型 = rs!对象类型
                            .保留对象 = False
                            .对象属性 = Split(rs!对象属性, "|")(1)
                            .内容行次 = 0
                            .内容文本 = Nvl(rs!内容文本)
                            .对象序号 = Nvl(rs!对象序号, 0)
                            If InStr(strTableName, "电子病历内容") > 0 Then
                                .开始版 = Nvl(rs!开始版, 1)
                                .终止版 = Nvl(rs!终止版, 0)
                            End If
                        End If
                Case cprCTElement          '2-单要素
                    lngKey = Me.Elements.Add
                    If Nvl(rs!父ID, 0) <> 0 Then '混合单元格的要素子项
                        .ElementKey = .ElementKey & "|" & lngKey
                        With Me.Elements("K" & lngKey)
                            .ID = rs!ID
                            .父ID = rs!父ID
                            .区域 = lngRow & "|" & lngCol
                            .内容文本 = Nvl(rs!内容文本)
                            .内容行次 = Nvl(rs!内容行次, 1)
                            .诊治要素ID = Nvl(rs!诊治要素ID, 0)
                            .替换域 = rs!替换域
                            .要素名称 = rs!要素名称
                            .要素类型 = rs!要素类型
                            .要素长度 = rs!要素长度
                            .要素小数 = rs!要素小数
                            .要素单位 = Nvl(rs!要素单位)
                            .要素表示 = rs!要素表示
                            .输入形态 = Nvl(rs!输入形态)
                            .要素值域 = Nvl(rs!要素值域)
                            .保留对象 = Nvl(rs!保留对象, 0) = 1
                            .自动转文本 = Split(rs!对象属性, "|")(1)
                            .必填 = Split(rs!对象属性, "|")(2)
                            .对象属性 = Nvl(rs!对象属性, 0)
                            If InStr(strTableName, "电子病历内容") > 0 Then
                                .开始版 = Nvl(rs!开始版, 1)
                                .终止版 = Nvl(rs!终止版, 0)
                            End If
                            If .替换域 = 1 And Me.ET = TabET_全文示范编辑 Then .内容文本 = "" '导入范文时自动替换项目需重新得取
                        End With
                    Else
                        .ID = rs!ID
                        .文件ID = rs!文件ID
                        .对象类型 = rs!对象类型
                        .保留对象 = Nvl(rs!保留对象, 0) = 1
                        .内容行次 = 0
                        .内容文本 = Nvl(rs!内容文本)
                        .对象序号 = Nvl(rs!对象序号, 0)
                        .ElementKey = lngKey
                        If InStr(strTableName, "电子病历内容") > 0 Then
                            .开始版 = Nvl(rs!开始版, 1)
                            .终止版 = Nvl(rs!终止版, 0)
                        End If
                        With Me.Elements("K" & lngKey)
                            .ID = rs!ID
                            .父ID = 0
                            .区域 = lngRow & "|" & lngCol
                            .内容文本 = Nvl(rs!内容文本)
                            .诊治要素ID = Nvl(rs!诊治要素ID, 0)
                            .替换域 = rs!替换域
                            .要素名称 = rs!要素名称
                            .要素类型 = rs!要素类型
                            .要素长度 = rs!要素长度
                            .要素小数 = rs!要素小数
                            .要素单位 = Nvl(rs!要素单位)
                            .要素表示 = rs!要素表示
                            .输入形态 = Nvl(rs!输入形态)
                            .要素值域 = Nvl(rs!要素值域)
                            .保留对象 = Nvl(rs!保留对象, 0) = 1
                            .自动转文本 = Split(rs!对象属性, "|")(1)
                            .必填 = Split(rs!对象属性, "|")(2)
                            .对象属性 = Nvl(rs!对象属性, 0)
                            If InStr(strTableName, "电子病历内容") > 0 Then
                                .开始版 = Nvl(rs!开始版, 1)
                                .终止版 = Nvl(rs!终止版, 0)
                            End If
                        End With
                        If Me.Elements("K" & lngKey).替换域 = 1 And Me.ET = TabET_全文示范编辑 Then '导入范文时自动替换项目需重新得取
                            Me.Elements("K" & lngKey).内容文本 = ""
                            .内容文本 = ""
                        End If
                    End If
                Case cprCTTextElement       '3-文本与多要素混合编辑
                    .ID = rs!ID
                    .文件ID = rs!文件ID
                    .对象类型 = rs!对象类型
                    .保留对象 = Nvl(rs!保留对象, 0) = 1
                    .对象属性 = Split(rs!对象属性, "|")(1)
                    .内容行次 = 0
                    .内容文本 = Nvl(rs!内容文本)
                    .对象序号 = Nvl(rs!对象序号, 0)
                    If InStr(strTableName, "电子病历内容") > 0 Then
                        .开始版 = Nvl(rs!开始版, 1)
                        .终止版 = Nvl(rs!终止版, 0)
                    End If
                Case cprCTPicture, cprCTReportPic         '4-参考图'5-报告图
                    lngKey = Me.Pictures.Add
                    .ID = rs!ID
                    .文件ID = rs!文件ID
                    .对象类型 = rs!对象类型
                    .保留对象 = Nvl(rs!保留对象, 0) = 1
                    .内容行次 = 0
                    .内容文本 = ""
                    .对象序号 = Nvl(rs!对象序号, 0)
                    .对象属性 = Split(rs!对象属性, "|")(1)
                    .PictureKey = lngKey
                    If InStr(strTableName, "电子病历内容") > 0 Then
                        .开始版 = Nvl(rs!开始版, 1)
                        .终止版 = Nvl(rs!终止版, 0)
                    End If
                    strMarkKeys = ""
                    With Me.Pictures("K" & lngKey)
                        .PicID = rs!ID
                        Set .OrigPic = .GetPictureFromDB(Me.EM, Me.ET, blnMoved, Me.PicMarks, strMarkKeys)
                    End With
                    .PicMarkKey = strMarkKeys
                Case cprCTSign, cprCTRowSign, cprCTColSign         '6-签名7-行控签名8-列控签名
                    lngKey = Me.Signs.Add
                    .ID = rs!ID
                    .文件ID = rs!文件ID
                    .对象类型 = rs!对象类型
                    .保留对象 = Nvl(rs!保留对象, 0) = 1
                    .内容行次 = 0
                    .对象属性 = Split(rs!对象属性, "|")(1)
                    .对象序号 = Nvl(rs!对象序号, 0)
                    .内容文本 = Nvl(rs!内容文本)
                    .SignKey = lngKey
                    If InStr(strTableName, "电子病历内容") > 0 Then
                        .开始版 = Nvl(rs!开始版, 1)
                        .终止版 = Nvl(rs!终止版, 0)
                        If Me.ET = TabET_单病历编辑 And Nvl(rs!终止版, 0) <> 0 Then mReadOnly = 1 '已经签名后再次以修改方式进入
                    End If
                    With Me.Signs("K" & lngKey)
                        If Nvl(rs!内容文本) <> "" Then '签名在设计时仅为占位,无实际信息
                            .姓名 = Nvl(rs!内容文本)
                            .签名信息 = Nvl(rs!要素值域)
                            .前置文字 = Nvl(rs!要素名称)
                            .签名级别 = Nvl(rs!要素表示, 0)
                            .时间戳 = Nvl(rs!要素单位)
                            .签名方式 = Split(Split(rs!对象属性, "|")(1), ";")(0)
                            .签名规则 = Split(Split(rs!对象属性, "|")(1), ";")(1)
                            .证书ID = Split(Split(rs!对象属性, "|")(1), ";")(2)
                            .显示手签 = Split(Split(rs!对象属性, "|")(1), ";")(3)
                            .签名时间 = Split(Split(rs!对象属性, "|")(1), ";")(4)
                            .显示时间 = Split(Split(rs!对象属性, "|")(1), ";")(5)
                        End If
                    End With
            End Select
            If InStr(.MergeRange, ";") > 0 Then
                lsRow = Split(Split(.MergeRange, ";")(0), ",")(0): lsCol = Split(Split(.MergeRange, ";")(0), ",")(1)
                leRow = Split(Split(.MergeRange, ";")(1), ",")(0): leCol = Split(Split(.MergeRange, ";")(1), ",")(1)
                For l = lsRow To leRow
                    For j = lsCol To leCol
                        Me.Cells.Cell(l, j).对象类型 = .对象类型
                    Next
                Next
            End If
        End With
        rs.MoveNext
    Loop
    
    gstrSQL = "Select Max(对象序号) 对象序号 " & vbNewLine & _
            "From " & strTableName & vbNewLine & _
            "Where 文件id = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取最大序号", lngFileID)
    mMaxNo = rs!对象序号
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ClearClsID()
'功能清空ID，由 病人病历新增时调用
Dim l As Long

    For l = 1 To Me.Cells.Count '包括固定文本,可编辑文本,PIC,签名,
        Me.Cells.Item(l).ID = 0
        Me.Cells.Item(l).开始版 = 1
    Next

    For l = 1 To Me.Texts.Count
        Me.Texts.Item(l).ID = 0
        Me.Texts.Item(l).开始版 = 1
    Next
    For l = 1 To Me.Elements.Count
        Me.Elements.Item(l).ID = 0
        Me.Elements.Item(l).开始版 = 1
    Next
    
End Sub

Public Function SaveDoc(arrSQL As Variant) As Boolean
'功能：保存文档
Dim rs As New ADODB.Recordset, i As Integer
        On Error GoTo errHand
        
        If Me.ET = TabET_单病历编辑 Or Me.ET = TabET_单病历审核 Then '判断在编辑期间该病历文件是否已经被修改过！如果是，则不允许保存！
            Dim strDTOld As String, strDTNew As String
            strDTOld = Format(Me.EPRPatiRecInfo.保存时间, "YYYYMMDDHHmmss")
            gstrSQL = "select 保存时间 from 电子病历记录 where ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", Me.EPRPatiRecInfo.ID)
            If Not rs.EOF Then
                strDTNew = Format(Nvl(rs("保存时间"), 0), "YYYYMMDDHHmmss")
                If strDTNew <> strDTOld Then
                    MsgBox "对不起，同一份病历文件在其他地方已经被修改过。" & vbCrLf & _
                        "无法保存该病历文件，需要重新进入！", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If


        Select Case ET '先更新所有对象相关信息（同步其“文件ID”值）
        Case TabET_病历文件定义
            For i = 1 To Me.Cells.Count: Me.Cells(i).文件ID = Me.EPRFileInfo.ID: Next
        Case TabET_全文示范编辑
            If EPRDemoInfo.ID = 0 Then Me.EPRDemoInfo.ID = zlDatabase.GetNextId("病历范文目录")
            For i = 1 To Me.Cells.Count: Me.Cells(i).文件ID = Me.EPRDemoInfo.ID: Next
        Case TabET_单病历编辑, TabET_单病历审核
            If Me.EM = TabEm_新增 Then Call ClearClsID '新增时ID需要清空'因为ID是病历文件结构的ID
            If Me.EPRPatiRecInfo.ID = 0 Then Me.EPRPatiRecInfo.ID = zlDatabase.GetNextId("电子病历记录")
            For i = 1 To Me.Cells.Count: Me.Cells(i).文件ID = Me.EPRPatiRecInfo.ID: Next
        End Select
        Call CheckValidIDs 'ID号有效性检测！

        If Not SaveConText(arrSQL) Then MsgBox "生成文档内容出错，不能保存", vbInformation, gstrSysName: Exit Function '保存文档内容
        If Not SaveStructFile(arrSQL) Then MsgBox "生成文档结构出错，不能保存", vbInformation, gstrSysName: Exit Function '保存XML
        
        Select Case Me.ET
            Case TabET_病历文件定义  '保存文档页面设置,仅有定义时对页面的更改保存
                If Not EPRFileInfo.GetSaveSQL(arrSQL) Then Exit Function
            Case TabET_全文示范编辑
                If Not EPRDemoInfo.GetSaveSQL(arrSQL) Then Exit Function
            Case TabET_单病历编辑, TabET_单病历审核
                If Not EPRPatiRecInfo.GetSaveSQL(arrSQL) Then Exit Function
        End Select
        
        SaveDoc = True
        
        Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub CheckValidIDs()
Dim l As Long, lCount As Long, lKey As Long
    lCount = Me.Cells.Count
    For l = 1 To lCount
        With Me.Cells(l)
            If Me.ET = TabET_病历文件定义 Or Me.ET = TabET_全文示范编辑 Then .对象序号 = l '对象序号重新分配
            Select Case .对象类型
                Case cprCTFixtext, cprCTText, cprCTPicture, cprCTReportPic, cprCTRowSign, cprCTColSign, cprCTSign, cprCTElement
                    .CheckValidID ET
                Case cprCTTextElement
                    .CheckValidID ET
                    For lKey = 1 To UBound(Split(.TextKey, "|"))
                        Me.Texts("K" & Split(.TextKey, "|")(lKey)).CheckValidID ET
                    Next
                    For lKey = 1 To UBound(Split(.ElementKey, "|"))
                        Me.Elements("K" & Split(.ElementKey, "|")(lKey)).CheckValidID ET
                    Next
            End Select
        End With
    Next
End Sub
Private Function SaveConText(arrSQL As Variant) As Boolean
Dim lngRow As Long, lngCol As Long, lngCell As Long, lngTmp As Long, lngKey As Long, lngCount As Long
    On Error GoTo errHand
    lngCount = Me.Cells.Count
    For lngCell = 1 To lngCount
        With Me.Cells(lngCell)
            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then '只有合并单元格首个或非合并单元格才刷新
                Select Case .对象类型
                    Case cprCTFixtext, cprCTText    '固定文本与普通文本单元格,混合区域的文本不用此方法
                        If Not .GetSaveSQL(ET, arrSQL) Then Exit Function
                    Case cprCTElement               '要素单元格
                        If Not Me.Elements("K" & .ElementKey).GetSaveSQL(ET, Me.Cells(.Key), arrSQL) Then Exit Function
                    Case cprCTTextElement           '混合单元格
                        If Not .GetSaveSQL(ET, arrSQL) Then Exit Function
                        For lngTmp = 1 To UBound(Split(.TextKey, "|")) '对本单元格中所有文本进行保存
                            Me.Texts("K" & Split(.TextKey, "|")(lngTmp)).父ID = .ID
                            If Not Me.Texts("K" & Split(.TextKey, "|")(lngTmp)).GetSaveSQL(ET, Me.Cells(.Key), arrSQL) Then Exit Function
                        Next
                        For lngTmp = 1 To UBound(Split(.ElementKey, "|"))
                            Me.Elements("K" & Split(.ElementKey, "|")(lngTmp)).父ID = .ID
                            If Not Me.Elements("K" & Split(.ElementKey, "|")(lngTmp)).GetSaveSQL(ET, Me.Cells(.Key), arrSQL) Then Exit Function
                        Next
                    Case cprCTPicture, cprCTReportPic           '参考图'报告图
                        If Not Me.Pictures("K" & .PictureKey).GetSaveSQL(ET, Me.Cells(.Key), Me.PicMarks, arrSQL) Then Exit Function  '保存图片,同时保存标记图
                    Case cprCTSign                          '签名
                        If Not Me.Signs("K" & .SignKey).GetSaveSQL(ET, Me.Cells(.Key), arrSQL) Then Exit Function
                    Case cprCTRowSign, cprCTColSign    '行控签名'列控签名
                        If .终止版 = 0 Or .终止版 = 1 Or .终止版 = Me.EPRPatiRecInfo.最后版本 + 1 Then
                            If Not Me.Signs("K" & .SignKey).GetSaveSQL(ET, Me.Cells(.Key), arrSQL) Then Exit Function
                        End If
                End Select
            End If
        End With
    Next
    
    SaveConText = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveStructFile(arrSQL As Variant) As Boolean
'生成保存文档结构SQL
Dim strZipFile As String, strXmlFile As String, blnSave As Boolean
    On Error GoTo errHand
    If Not BuildXmlFile(strXmlFile) Then Exit Function
    
    If gobjFSO.FileExists(strXmlFile) Then
        strZipFile = zlFileZip(strXmlFile)
        gobjFSO.DeleteFile strXmlFile, True
    End If
    
    If strZipFile = "" Then Exit Function
    If gobjFSO.FileExists(strZipFile) Then
        Select Case ET
        Case TabET_病历文件定义
            blnSave = zlBlobSql(1, EPRFileInfo.ID, strZipFile, arrSQL)
        Case TabET_全文示范编辑
            blnSave = zlBlobSql(3, EPRDemoInfo.ID, strZipFile, arrSQL)
        Case TabET_单病历编辑, TabET_单病历审核
            blnSave = zlBlobSql(5, EPRPatiRecInfo.ID, strZipFile, arrSQL)
        End Select
        gobjFSO.DeleteFile strZipFile, True      '删除临时文件
    End If
    SaveStructFile = blnSave
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function BuildXmlFile(strFile As String, Optional ByVal blnExport As Boolean, Optional ByVal blnSignExp As Boolean) As Boolean
'功能:生成XML文件并返回文件全路径
'说明：blnExport＝True 导出XML ＝False 表示只导出XML用于存数据库不导出页面信息,blnSignExp签名时导出用于数字验证的原文
Dim objDoc As New MSXML2.DOMDocument, oRoot  As IXMLDOMElement        '根节点
Dim oParNode As IXMLDOMNode, oSubNode As IXMLDOMNode, oChildNode As IXMLDOMNode   '父节点   '子节点
Dim strTmp As String, oStream As New ADODB.Stream      '流对象
Dim l As Long, i As Integer, lngCount As Long
Dim strRow As String, strCol As String, lngRow As Long, lngCol As Long, strHeight As String, strWidth As String
    On Error GoTo errHand
    Call XmlSetComment(objDoc, "操作员:" & UserInfo.姓名 & ",部门:" & UserInfo.部门ID & ",时间:" & Format(Now(), "YYYY年MM月DD日"))
    Set oRoot = XmlSetRoot(objDoc, "Table") '设置根节点
    
    If blnExport Then '导出XML调用
        Set oParNode = XmlCreateNode(1, oRoot, "纸张") '记录页面格式
        With Me.EPRFileInfo
            Call XmlCreateNode(2, oParNode, "纸型", .PaperKind)
            Call XmlCreateNode(2, oParNode, "宽度", .PaperWidth) '210</宽度>
            Call XmlCreateNode(2, oParNode, "高度", .PaperHeight) ' >294</高度>
            Call XmlCreateNode(2, oParNode, "上边距", .MarginTop) ' >3</上边距>
            Call XmlCreateNode(2, oParNode, "下边距", .MarginBottom) ' >3</下边距>
            Call XmlCreateNode(2, oParNode, "左边距", .MarginLeft) ' >2.5</左边距>
            Call XmlCreateNode(2, oParNode, "右边距", .MarginRight) ' >2.5</右边距>
            Call XmlCreateNode(2, oParNode, "纸向", .PaperOrient)  '>纵向</纸向>
            '页眉
            Call XmlCreateNode(2, oParNode, "页眉", Replace(.HeadConText, vbCrLf, "\r\n")) ' String          '内容(要素用[]包括),页眉内容
            Call XmlCreateNode(2, oParNode, "页眉字体", .HeadFontName) ' String
            Call XmlCreateNode(2, oParNode, "页眉字号", .HeadFontSize) ' Long
            Call XmlCreateNode(2, oParNode, "页眉粗体", .HeadFontBold) ' Boolean
            Call XmlCreateNode(2, oParNode, "页眉斜体", .HeadFontItalic) ' Boolean
            Call XmlCreateNode(2, oParNode, "页眉下划线", .HeadFontUnderline) ' Boolean
            Call XmlCreateNode(2, oParNode, "页眉删除线", .HeadFontStrikethrough) ' Boolean
            Call XmlCreateNode(2, oParNode, "页眉字色", .HeadFontColor) ' OLE_COLOR
            Call XmlCreateNode(2, oParNode, "页眉边距", .HeadMargin) ' 页眉边距
            
            ''页脚                  '
            Call XmlCreateNode(2, oParNode, "页脚", Replace(.FootConText, vbCrLf, "\r\n")) ' String               '内容(要素用[]包括),页脚内容
            Call XmlCreateNode(2, oParNode, "页脚字体", .FootFontName) ' String
            Call XmlCreateNode(2, oParNode, "页脚字号", .FootFontSize) ' Long
            Call XmlCreateNode(2, oParNode, "页脚粗体", .FootFontBold) ' Boolean
            Call XmlCreateNode(2, oParNode, "页脚斜体", .FootFontItalic) ' Boolean
            Call XmlCreateNode(2, oParNode, "页脚下划线", .FootFontUnderline) ' Boolean
            Call XmlCreateNode(2, oParNode, "页脚删除线", .FootFontStrikethrough) ' Boolean
            Call XmlCreateNode(2, oParNode, "页脚字色", .FootFontColor) ' OLE_COLOR
            Call XmlCreateNode(2, oParNode, "页脚边距", .FootMargin) ' 页脚边距
            If .HeadPic.Handle <> 0 Then
                '存储图片对象
                oStream.Type = adTypeBinary
                oStream.Open
                strTmp = App.Path & "\Tmp.Jpg"
                If gobjFSO.FileExists(strTmp) Then gobjFSO.DeleteFile strTmp, True
                SavePicture .HeadPic, strTmp
                oStream.LoadFromFile strTmp
                Set oSubNode = objDoc.createElement("页眉图片")
                oSubNode.dataType = "bin.base64"
                oSubNode.nodeTypedValue = oStream.Read
                oParNode.appendChild oSubNode
                oStream.Close
                '删除临时文件
                If gobjFSO.FileExists(strTmp) Then gobjFSO.DeleteFile strTmp, True
            End If
        End With
    End If
    
    Set oParNode = XmlCreateNode(1, oRoot, "属性")  '记录行数，列数，行高，列宽
    Call XmlCreateNode(2, oParNode, "行数", Me.Cells.Rows)
    Call XmlCreateNode(2, oParNode, "列数", Me.Cells.Cols)
    lngCount = Me.Cells.Count: lngRow = 0: lngCol = 0: strRow = "": strCol = "": strHeight = "": strWidth = ""
    For l = 1 To Me.Cells.Rows
        strHeight = strHeight & "|" & Me.Cells.Cell(l, 1).Height
    Next
    Call XmlCreateNode(2, oParNode, "行高", strHeight)
    
    For l = 1 To Me.Cells.Cols
        strWidth = strWidth & "|" & Me.Cells.Cell(1, l).Width
    Next
    Call XmlCreateNode(2, oParNode, "列宽", strWidth)
    
    For l = 1 To lngCount
        With Me.Cells.Item(l)
        If blnSignExp And (.对象类型 = cprCTSign Or .对象类型 = cprCTRowSign Or .对象类型 = cprCTColSign) Then
             '签名时提取原文进行数字验证时不导签名单元格
        Else
            Set oParNode = XmlCreateNode(1, oRoot, "Cell")
            Call XmlCreateNode(2, oParNode, "区域", .MergeRange) '<!--合并区域左上方单元格(横向,纵向);合并区域右下方单元格(横向,纵向)-->
            Call XmlCreateNode(2, oParNode, "合并", .Merge)
            Call XmlCreateNode(2, oParNode, "边框", .CellLineTop & "|" & .CellLineBottom & "|" & .CellLineLeft & "|" & .CellLineRight & "|" & .CellLineTopColor & "|" & .CellLineBottomColor & "|" & .CellLineLeftColor & "|" & .CellLineRightColor) '边框样式 四位|分隔,上下左右 <!--0-无边框 1-有边框,2-粗线,4-虚线--> 边框颜色  四位|分隔,上下左右
            If IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then '合并单元格的首个单元格和非合并单元格
                Call XmlCreateNode(2, oParNode, "字体", .FontName & "|" & .FontSize & "|" & .FontBold & "|" & .FontItalic & "|" & .FontUnderline & "|" & .FontStrikeout & "|" & .FontColor & "|" & .HAlignment & "|" & .VAlignment) ' 宋体</字体|字号|粗体|斜体|下划线|删除线|字体颜色|横向对齐|纵向对齐
                If blnExport Then '导出XML调用,写内容,单元格混合区域,图片时,生成多个"内容"节点,第0个节点为Cell内容
                    Set oSubNode = XmlCreateNode(2, oParNode, "内容")
                    Call XmlCreateNode(3, oSubNode, "ID", .ID)
                    Call XmlCreateNode(3, oSubNode, "对象类型", .对象类型)
                    Call XmlCreateNode(3, oSubNode, "保留对象", .保留对象)
                    Call XmlCreateNode(3, oSubNode, "对象属性", .对象属性)
                    Call XmlCreateNode(3, oSubNode, "内容文本", Replace(.内容文本, vbCrLf, "\r\n"))
                    
                    Select Case .对象类型
                        Case cprCTFixtext, cprCTText    '固定文本与普通文本单元格
                        Case cprCTElement               '单要素单元格
                            With Me.Elements("K" & .ElementKey)
                                Call XmlCreateNode(3, oSubNode, "要素文本", Replace(.内容文本, vbCrLf, "\r\n"))
                                Call XmlCreateNode(3, oSubNode, "诊治要素ID", .诊治要素ID)
                                Call XmlCreateNode(3, oSubNode, "替换域", .替换域)
                                Call XmlCreateNode(3, oSubNode, "要素名称", .要素名称)
                                Call XmlCreateNode(3, oSubNode, "要素类型", .要素类型)
                                Call XmlCreateNode(3, oSubNode, "要素长度", .要素长度)
                                Call XmlCreateNode(3, oSubNode, "要素小数", .要素小数)
                                Call XmlCreateNode(3, oSubNode, "要素单位", .要素单位)
                                Call XmlCreateNode(3, oSubNode, "要素表示", .要素表示)
                                Call XmlCreateNode(3, oSubNode, "输入形态", .输入形态)
                                Call XmlCreateNode(3, oSubNode, "要素值域", .要素值域)
                                Call XmlCreateNode(3, oSubNode, "自动转文本", Split(.对象属性, "|")(0))
                                Call XmlCreateNode(3, oSubNode, "必填", Split(.对象属性, "|")(1))
                            End With
                        Case cprCTTextElement           '混合单元格
                            For i = 1 To UBound(Split(.TextKey, "|")) '对本单元格中所有文本进行保存
                                Set oSubNode = XmlCreateNode(2, oParNode, "内容")
                                Call XmlCreateNode(3, oSubNode, "父ID", .ID)
                                Call XmlCreateNode(3, oSubNode, "对象类型", cprCTText)
                                Call XmlCreateNode(3, oSubNode, "对象序号", .对象序号)
                                With Me.Texts("K" & Split(.TextKey, "|")(i))
                                    Call XmlCreateNode(3, oSubNode, "ID", .ID)
                                    Call XmlCreateNode(3, oSubNode, "保留对象", .保留对象)
                                    Call XmlCreateNode(3, oSubNode, "文本行次", .内容行次)
                                    Call XmlCreateNode(3, oSubNode, "文本内容", Replace(.内容文本, vbCrLf, "\r\n"))
                                End With
                            Next
                            For i = 1 To UBound(Split(.ElementKey, "|"))
                                Set oSubNode = XmlCreateNode(2, oParNode, "内容")
                                Call XmlCreateNode(3, oSubNode, "父ID", .ID)
                                Call XmlCreateNode(3, oSubNode, "对象类型", cprCTTextElement)
                                Call XmlCreateNode(3, oSubNode, "对象序号", .对象序号)
                                With Me.Elements("K" & Split(.ElementKey, "|")(i))
                                    Call XmlCreateNode(3, oSubNode, "ID", .ID)
                                    Call XmlCreateNode(3, oSubNode, "要素保留", .保留对象)
                                    Call XmlCreateNode(3, oSubNode, "要素文本", Replace(.内容文本, vbCrLf, "\r\n"))
                                    Call XmlCreateNode(3, oSubNode, "内容行次", .内容行次)
                                    Call XmlCreateNode(3, oSubNode, "诊治要素ID", .诊治要素ID)
                                    Call XmlCreateNode(3, oSubNode, "替换域", .替换域)
                                    Call XmlCreateNode(3, oSubNode, "要素名称", .要素名称)
                                    Call XmlCreateNode(3, oSubNode, "要素类型", .要素类型)
                                    Call XmlCreateNode(3, oSubNode, "要素长度", .要素长度)
                                    Call XmlCreateNode(3, oSubNode, "要素小数", .要素小数)
                                    Call XmlCreateNode(3, oSubNode, "要素单位", .要素单位)
                                    Call XmlCreateNode(3, oSubNode, "要素表示", .要素表示)
                                    Call XmlCreateNode(3, oSubNode, "输入形态", .输入形态)
                                    Call XmlCreateNode(3, oSubNode, "要素值域", .要素值域)
                                    Call XmlCreateNode(3, oSubNode, "保留对象", .保留对象)
                                    Call XmlCreateNode(3, oSubNode, "自动转文本", Split(.对象属性, "|")(0))
                                    Call XmlCreateNode(3, oSubNode, "必填", Split(.对象属性, "|")(1))
                                End With
                            Next
                        Case cprCTPicture, cprCTReportPic              '参考图
                            Set oChildNode = oSubNode.selectSingleNode("对象属性")
                            oChildNode.Text = Me.Pictures("K" & .PictureKey).对象属性
                            If Me.Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                                '存储图片对象
                                oStream.Type = adTypeBinary
                                oStream.Open
                                strTmp = App.Path & "\Tmp.Jpg"
                                If gobjFSO.FileExists(strTmp) Then
                                    gobjFSO.DeleteFile strTmp, True
                                End If
                                SavePicture Me.Pictures("K" & .PictureKey).OrigPic, strTmp
                                oStream.LoadFromFile strTmp
                                Set oChildNode = objDoc.createElement(IIf(.对象类型 = cprCTPicture, "参考图", "报告图"))
                                oChildNode.dataType = "bin.base64"
                                oChildNode.nodeTypedValue = oStream.Read
                                oSubNode.appendChild oChildNode
                                oStream.Close
                                '删除临时文件
                                If gobjFSO.FileExists(strTmp) Then gobjFSO.DeleteFile strTmp, True
                            End If
                            '标记
                            For i = 1 To UBound(Split(.PicMarkKey, "|"))  '对本单元格中所有标记进行保存
                                Set oSubNode = XmlCreateNode(2, oParNode, "内容")
                                Call XmlCreateNode(3, oSubNode, "父ID", .ID)
                                Call XmlCreateNode(3, oSubNode, "对象类型", cprCTPicMark)
                                With Me.PicMarks("K" & Split(.PicMarkKey, "|")(i))
                                    Call XmlCreateNode(3, oSubNode, "保留对象", .保留对象)
                                    Call XmlCreateNode(3, oSubNode, "对象属性", .对象属性)
                                End With
                            Next
                        Case cprCTSign, cprCTRowSign, cprCTColSign                '签名
                            Set oChildNode = oSubNode.selectSingleNode("对象属性")
                            oChildNode.Text = Me.Signs("K" & .SignKey).对象属性
                            With Me.Signs("K" & .SignKey)
                                Call XmlCreateNode(3, oSubNode, "要素名称", .前置文字)
                                Call XmlCreateNode(3, oSubNode, "要素单位", .时间戳)
                                Call XmlCreateNode(3, oSubNode, "要素表示", .签名级别)
                                Call XmlCreateNode(3, oSubNode, "签名信息", .签名信息)
                            End With
                    End Select
                End If
            End If
        End If
        End With
    Next
    
    Call XmlSetVersion(objDoc) '写版本信息
    
    If strFile = "" Then
        strFile = App.Path & "\TMP.XML"
    End If
    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True    '保存为临时文件
    objDoc.Save strFile
    
    BuildXmlFile = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetHeadFootContent(ByVal bType As Byte) As String
'功能:提取页眉或页脚,替换掉其中除页数,页码以外的所有要素
'参数:bType=0 页眉,bType=1 页脚
Dim strContent As String, strTmp As String
    On Error GoTo errHand
    If bType = 0 Then
        strContent = EPRFileInfo.页眉
    ElseIf bType = 1 Then
        strContent = EPRFileInfo.页脚
    End If
    
    If InStr(strContent, "[文件名]") > 0 Then
        strContent = Replace(strContent, "[文件名]", EPRFileInfo.名称)
    End If
    
    If InStr(strContent, "[打印日期]") > 0 Then
        strContent = Replace(strContent, "[打印日期]", Format(zlDatabase.Currentdate, "yyyy年mm月dd日"))
    End If
            
    If InStr(strContent, "[打印时间]") > 0 Then
        strContent = Replace(strContent, "[打印时间]", Format(zlDatabase.Currentdate, "HH:mm:ss"))
    End If
    
    If InStr(strContent, "[单位名称]") > 0 Then
        strContent = Replace(strContent, "[单位名称]", zl9ComLib.GetUnitName)
    End If
    
    If InStr(strContent, "[病历名称]") > 0 Then
        strContent = Replace(strContent, "[病历名称]", EPRFileInfo.名称)
    End If
    
    If InStr(strContent, "[姓名]") > 0 Then
        strContent = Replace(strContent, "[姓名]", ElementContent("姓名"))
    End If
    
    If InStr(strContent, "[性别]") > 0 Then
        strContent = Replace(strContent, "[性别]", ElementContent("性别"))
    End If
    
    If InStr(strContent, "[年龄]") > 0 Then
        strContent = Replace(strContent, "[年龄]", ElementContent("年龄"))
    End If
    
    If InStr(strContent, "[标识号]") > 0 Then
        strContent = Replace(strContent, "[标识号]", IIf(EPRPatiRecInfo.病人来源 <> TabPF_住院, "[门诊号]", "[住院号]"))
    End If
    
    If InStr(strContent, "[门诊号]") > 0 Then
        strContent = Replace(strContent, "[门诊号]", ElementContent("门诊号"))
    End If
    
    If InStr(strContent, "[住院号]") > 0 Then
        strContent = Replace(strContent, "[住院号]", ElementContent("住院号"))
    End If
    
    If InStr(strContent, "[入院科室]") > 0 Then
        strContent = Replace(strContent, "[入院科室]", ElementContent("入院科室"))
    End If
    
    If InStr(strContent, "[入院病区]") > 0 Then
        strContent = Replace(strContent, "[入院病区]", ElementContent("入院病区"))
    End If
    
    If InStr(strContent, "[当前科室]") > 0 Then
        strContent = Replace(strContent, "[当前科室]", ElementContent("当前科室"))
    End If
    
    If InStr(strContent, "[当前床号]") > 0 Then
        strContent = Replace(strContent, "[当前床号]", ElementContent("当前床号"))
    End If
    
    If InStr(strContent, "[入院日期]") > 0 Then
        strContent = Replace(strContent, "[入院日期]", ElementContent("入院日期"))
    End If
    
    If InStr(strContent, "[出院日期]") > 0 Then
        strContent = Replace(strContent, "[出院日期]", ElementContent("出院日期"))
    End If
    
    If InStr(strContent, "[住院次数]") > 0 Then
        strContent = Replace(strContent, "[住院次数]", ElementContent("住院次数"))
    End If
    
    If InStr(strContent, "[住院医师]") > 0 Then
        strContent = Replace(strContent, "[住院医师]", ElementContent("住院医师"))
    End If

    If InStr(strContent, "[责任护士]") > 0 Then
        strContent = Replace(strContent, "[责任护士]", ElementContent("责任护士"))
    End If
    GetHeadFootContent = strContent
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function ElementContent(ByVal elName As String) As String
'跟据要素名称取值，先遍历要素集合，没找到才取数据库,适用于更新页眉页脚中的要素
Dim l As Long, strTmp As String
    With Elements
        For l = 1 To .Count
            With .Item(l)
                If .要素名称 = elName Then
                    strTmp = .内容文本 & .要素单位
                    Exit For
                End If
            End With
        Next
        If l > .Count Then
            strTmp = GetReplaceEleValue(elName, EPRPatiRecInfo.病人ID, EPRPatiRecInfo.主页ID, EPRPatiRecInfo.病人来源, EPRPatiRecInfo.医嘱id)
        End If
    End With
    
    ElementContent = strTmp
End Function
Public Function CellContent(ByVal strCellKey As String) As String
'功能：提取单元格文字内容，主要用于打印
'参数：据在单元格的 Cell.Key
Dim strTmp As String
    With Me.Cells(strCellKey)
        Select Case .对象类型
        Case cprCTFixtext, cprCTText   '0-固定文本(不可编辑)1-文本型(可编辑多行文本)
            strTmp = .内容文本
        Case cprCTElement    '2-单要素
            If ET = TabET_病历文件定义 Or ET = TabET_全文示范编辑 Then
                If .ElementKey <> "" Then
                    If Elements("K" & .ElementKey).输入形态 = 1 Then
                        strTmp = Elements("K" & .ElementKey).内容文本 & Elements("K" & .ElementKey).要素单位
                    Else
                        strTmp = "" & Elements("K" & .ElementKey).要素单位
                    End If
                End If
            Else
                If .内容文本 = "" Then
                    If Elements("K" & .ElementKey).替换域 = 1 Then  '自动替换要素
                        strTmp = GetReplaceEleValue(Elements("K" & .ElementKey).要素名称, EPRPatiRecInfo.病人ID, EPRPatiRecInfo.主页ID, EPRPatiRecInfo.病人来源, EPRPatiRecInfo.医嘱id)
                        If strTmp = "" And Not Elements("K" & .ElementKey).自动转文本 Then '没取到值，是否自动转换成文本(空)
                            strTmp = "" & Elements("K" & .ElementKey).要素单位
                        Else
                            strTmp = strTmp & Elements("K" & .ElementKey).要素单位
                        End If
                    Else
                        strTmp = Elements("K" & .ElementKey).内容文本 & Elements("K" & .ElementKey).要素单位
                    End If
                Else
                    strTmp = .内容文本 & Elements("K" & .ElementKey).要素单位
                End If
            End If
        Case cprCTTextElement '3-文本与多要素混合编辑
            strTmp = CellTextELement(strCellKey) '在打印已经将内容更新到内容文本中
        Case cprCTReportPic, cprCTPicture    '5-报告图
            strTmp = IIf(.对象类型 = cprCTPicture, "参考图", "报告图")
        Case cprCTSign         '6-签名'签名在设计时仅为占位,无实际信息；没有签名时终止版=0；普通签名后审核时不显示，以便再次签名；行控/列控签名后审核时要显示
            If ET = TabET_单病历编辑 Or ET = TabET_单病历审核 Then
                Select Case mReadOnly 'mReadOnly 0-正常,1-签名后点修改,2-主界面打开查阅或查阅历次签名版本
                    Case 0
                        If .终止版 <> 0 Then
                            With Signs("K" & .SignKey)
                                strTmp = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
                                strTmp = strTmp & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
                            End With
                        Else
                            strTmp = ""
                        End If
                    Case 1, 2
                        If .终止版 = 0 Then
                            strTmp = ""
                        Else
                            With Signs("K" & .SignKey)
                                strTmp = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
                                strTmp = strTmp & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
                            End With
                        End If
                End Select
            Else
                strTmp = ""
            End If
        Case cprCTRowSign, cprCTColSign '7-行控签名 '8-列控签名
            If ET = TabET_单病历编辑 Or ET = TabET_单病历审核 Then
                If .终止版 <> 0 Then
                    With Signs("K" & .SignKey)
                        strTmp = .前置文字 & .姓名 & IIf(.显示手签, "，手签：_____________", "")
                        strTmp = strTmp & IIf(Trim(.显示时间) = "", "", "，" & Format(.签名时间, .显示时间))
                    End With
                Else
                    strTmp = ""
                End If
            Else
                strTmp = ""
            End If
        End Select
    End With
    CellContent = strTmp
End Function
Private Function CellTextELement(ByVal strCellKey As String) As String
'功能：跟据Text Element填写F1Main中的单元格及类的内容文本
Dim i As Long, lCount As Long, strTmp As String, ltCount As Long, leCount As Long, cleTmp As cTabElement, strAEl As String
    On Error GoTo errHand
    With Me.Cells(strCellKey)
        ltCount = UBound(Split(.TextKey, "|")): If ltCount < 0 Then ltCount = 0
        leCount = UBound(Split(.ElementKey, "|")): If leCount < 0 Then leCount = 0
        lCount = ltCount + leCount
        For i = 1 To lCount
            Set cleTmp = .clElement(Me.Elements, i)
            If cleTmp Is Nothing Then '该次序为文本
                strTmp = strTmp & ToVarchar(.clText(Me.Texts, i).内容文本, 4000)
            Else
                With Me.Elements("K" & cleTmp.Key)
                    If .替换域 = 1 And (ET = TabET_单病历编辑 Or ET = TabET_单病历审核) Then
                        If Trim(.内容文本) = "" Then
                            strAEl = GetReplaceEleValue(.要素名称, Me.EPRPatiRecInfo.病人ID, Me.EPRPatiRecInfo.主页ID, Me.EPRPatiRecInfo.病人来源, Me.EPRPatiRecInfo.医嘱id)
                            strTmp = strTmp & strAEl
                        Else
                            strTmp = strTmp & .内容文本 & .要素单位
                        End If
                    Else
                        If .内容文本 <> "" Then
                            strTmp = strTmp & .内容文本 & .要素单位
                        End If
                    End If
                End With
            End If
        Next
        CellTextELement = strTmp
    End With
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function CellPicture(ByVal picTmpBox As Object, ByVal strCellKey As String, ByVal W As Long, ByVal H As Long) As StdPicture
'功能：提取单元格图片，主要用于打印
'参数：picTmpBox 临时图片控件，主要用于加载临时PictureBox,strCellKey 据在单元格的 Cell.Key; W,H为目标区域宽高
Dim picTmp As New StdPicture, i As Integer
    If ChkControl(picTmpBox) Then
        With Cells(strCellKey)
            If .PictureKey <> "" Then
                picTmpBox.Height = H
                picTmpBox.Width = W
                If Pictures("K" & .PictureKey).OrigPic.Handle <> 0 Then
                    Set picTmpBox.Picture = Pictures("K" & .PictureKey).OrigPic
                    picTmpBox.PaintPicture picTmpBox.Picture, 0, 0, picTmpBox.Width, picTmpBox.Height
                    If .PicMarkKey <> "" Then '有标记图先绘标记
                        For i = 1 To UBound(Split(.PicMarkKey, "|"))
                            ShowPicMark picTmpBox, PicMarks("K" & Split(.PicMarkKey, "|")(i))
                        Next
                    End If
                    Set picTmp = picTmpBox.Image
                End If
            End If
        End With
    End If
    Set CellPicture = picTmp
End Function

Public Sub PrintDoc(ByVal frmPar As Object, ByVal blnPreview As Boolean, Optional ByVal strPrintDeviceName As String)
'功能：调用打印
    On Error GoTo errHand
    If mfrmPrint Is Nothing Then Set mfrmPrint = New frmPrintPreview
    If mfrmPrint.Visible Then Exit Sub
    mfrmPrint.OutPut frmPar, Me, blnPreview, strPrintDeviceName
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Sub zlPrintDoc(ByVal frmPar As Object, ByVal blnPreview As Boolean, Optional ByVal strPrintDeviceName As String)
'功能：由外部程序调用，调用之前需调用InitOpenEPR
    Call ReadFileStructure
    Call ReadFileContent(mblnMove)
    zlCommFun.StopFlash
    Call PrintDoc(frmPar, blnPreview, strPrintDeviceName)
End Sub
Public Sub frmEditorClosed(ByVal lngOrderId As Long)
    If EPRPatiRecInfo.病历种类 = Tab诊疗报告 And (ET = TabET_单病历编辑 Or ET = TabET_单病历审核) Then '可能会被医技工作站调用
        If mfrmParent.Caption = "诊疗报告管理" Then
            Call mfrmParent.EditorClosed(lngOrderId)
        End If
    End If
End Sub
Public Function zlExportXML(ByVal strFile As String) As Boolean
'功能：导出XML文件
'说明：调用之前需要调用InitOpenEPR
    Call ReadFileStructure
    Call ReadFileContent(mblnMove)
    zlExportXML = BuildXmlFile(strFile, True)
    zlCommFun.StopFlash
End Function
Public Function zlGetForm() As Object
'功能：外部嵌套预览子窗体
    Set zlGetForm = mfrmDockView
End Function
Public Sub zlRefreshDockfrm()
'功能：外部调用刷新嵌套预览子窗体
'说明：调用之前需要调用InitOpenEPR
    mfrmDockView.zlRefresh Me
    zlCommFun.StopFlash
End Sub
Public Sub zlRefreshPacsPic(ByVal mlngOrderId As Long)
'功能：由调用程序主动刷新PACS报告图口
Dim frmThis As Form
    For Each frmThis In Forms
        If frmThis.Name = "frmMain" Then
            If frmThis.Document.医嘱id = mlngOrderId Then
                frmThis.zlRefreshPacsPic
            End If
        End If
    Next
End Sub
Public Function Showfrm(ByVal lFileId As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bPatiType As Byte, ByVal lngDeptId As Long) As Boolean
'功能：用于显示已经打开的编辑器
Dim frmThis As Form
    For Each frmThis In Forms
        If frmThis.Name = "frmMain" Then
            With frmThis.Document
                If IIf(.EM = TabEm_新增, .EPRFileInfo.ID, .EPRPatiRecInfo.ID) = lFileId And .EPRPatiRecInfo.病人ID = lngPatiID _
                    And .EPRPatiRecInfo.病人来源 = bPatiType And .EPRPatiRecInfo.主页ID = lngPageID _
                    And .EPRPatiRecInfo.科室ID = lngDeptId Then
                    frmThis.WindowState = vbMaximized
                    frmThis.Show
                    Showfrm = True: Exit For
                End If
            End With
        End If
    Next
End Function

Private Sub Class_Initialize()
    Set Cells = New cTabCells '表格集合
    Set Pictures = New cTabPictures '图片集合
    Set Texts = New cTabTexts      '文本集合,用于混合编辑区域
    Set Elements = New cTabElements '诊治要素集合
    Set Signs = New cTabSigns '签名集合
    Set PicMarks = New cTabPicMarks
    Set EPRFileInfo = New cTabFileDefineInfo '病历文件定义信息
    Set EPRDemoInfo = New cTabDemoInfo '病历示范信息
    Set EPRPatiRecInfo = New cTabPatiRecordInfo '病人电子病历信息
    Set mfrmDockView = New frmDockView          '预览窗体
    Set mfrmParent = Nothing
    Set frmEditor = New frmMain
    Set gAllFont = Nothing
    Set gobjFSO = Nothing
End Sub

Private Sub Class_Terminate()
Dim frmThis As Form
    On Error Resume Next
    Unload frmEditor
    Set frmEditor = Nothing
    If Not mfrmPrint Is Nothing Then Unload mfrmPrint
    Set mfrmPrint = Nothing
    If Not mfrmDockView Is Nothing Then Unload mfrmDockView
    Set mfrmDockView = Nothing
    
    For Each frmThis In Forms
        If frmThis.Name = "frmMain" Then Unload frmThis
    Next
    
    Set Cells = Nothing
    Set Pictures = Nothing
    Set Texts = Nothing
    Set Elements = Nothing
    Set Signs = Nothing
    Set PicMarks = Nothing
    Set EPRFileInfo = Nothing
    Set EPRDemoInfo = Nothing
    Set EPRPatiRecInfo = Nothing
    Set mfrmParent = Nothing
End Sub
Private Sub InitPreDefinedKeys()
    gKeyWords(1).KeyStart = "OS"            '提纲
    gKeyWords(1).KeyEnd = "OE"              '
    gKeyWords(2).KeyStart = "PS"            '图片
    gKeyWords(2).KeyEnd = "PE"              '
    gKeyWords(3).KeyStart = "ES"            '要素
    gKeyWords(3).KeyEnd = "EE"              '
    gKeyWords(4).KeyStart = "TS"            '表格
    gKeyWords(4).KeyEnd = "TE"              '
    gKeyWords(5).KeyStart = "SS"            '签名
    gKeyWords(5).KeyEnd = "SE"              '
    gKeyWords(6).KeyStart = "DS"            '诊断
    gKeyWords(6).KeyEnd = "DE"              '
End Sub

Private Sub mfrmPrint_PrintEpr()
'功能：打印标记，调用上级窗口刷新事件，打印事件，上级窗口是以每次NEW方式实例化的类,不能以事件返回给上级窗口
    On Error GoTo errHand
    If EPRPatiRecInfo.ID <> 0 Then
        gstrSQL = "Zl_电子病历打印_Insert(" & EPRPatiRecInfo.ID & "," & EPRPatiRecInfo.病历种类 & "," & EPRPatiRecInfo.病人ID & "," & EPRPatiRecInfo.主页ID & ",'" & UserInfo.姓名 & "')"
        zlDatabase.ExecuteProcedure gstrSQL, App.EXEName
    End If
    On Error Resume Next
    Call mfrmParent.RefreshList
    If EPRPatiRecInfo.ID <> 0 Then
        If EPRPatiRecInfo.病历种类 = Tab诊疗报告 Then Call mfrmParent.Event_AfterPrinted(EPRPatiRecInfo.ID) '报告需要打印事件
    End If
    Err.Clear
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
