Attribute VB_Name = "mdlExcel"
Option Explicit

Public gcnExcel As New ADODB.Connection '公共连接

Public Enum Excel_Col
    'excel格式文件的字段顺序
    影响类型 = 0
    发布版本
    是否有HTML文档
    问题编号
    登记模块
    登记用户
    用户需求
    修改说明
    问题风险
    风险评估说明
    影响模块
    是否需要培训
    备注
    相关问题
    阅读记录
    对用户影响评估
    操作培训情况
End Enum

Public Enum mCol
    阅读 = 0: 风险: 版本: 分类: 编号: 模块: 影响模块: 影响类型: 风险评估: 用户: 需求: 说明: 关联问题: 备注: 培训: 影响评估: 连接: 修改
End Enum
Public fntUnderLine  As StdFont '超连接的字体


Public Function OpenExcelFile(ByVal strFilename As String) As String
    '功能：打开Excel格式文件
    '入参：strFileName
    '出参：Sheet列表，以|分隔
    
    Dim BiaoMing As Variant
    Dim TableName As String
    Dim strSheet As String
    On Error GoTo errHandle
    OpenExcelFile = ""

    If gcnExcel.State = 1 Then     '如果以连接过，则关闭，初始化下次事务
        gcnExcel.Close
    End If
    
    gcnExcel.ConnectionString = "Provider=microsoft.jet.oledb.4.0;data source=" & strFilename & ";" & _
                              "Extended Properties=Excel 8.0;" & _
                              "Persist Security Info=False"
    gcnExcel.Open
    Set BiaoMing = gcnExcel.OpenSchema(adSchemaTables)     '创建数据库记录集
    
    TableName = "": strSheet = ""
    Do Until BiaoMing.EOF
        If BiaoMing("table_name") <> TableName Then   '列出所有表
            TableName = BiaoMing("table_name")
            If Right(TableName, 1) = "$" Then
                strSheet = strSheet & "|" & TableName
            End If
        End If
        BiaoMing.MoveNext
    Loop
    
    Set BiaoMing = Nothing
    If strSheet <> "" Then
        OpenExcelFile = Mid(strSheet, 2)
    End If
    Exit Function
errHandle:
    OpenExcelFile = ""
    MsgBox Err.Number & " " & Err.Description, vbQuestion, "升级阅读器"
    
End Function

Public Function OpenExcelSheet(ByVal strSheetName As String) As ADODB.Recordset
    '打开一个Sheet
    '入参: Sheet名
    '出参: ADO记录集
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSheet As String
    On Error GoTo errHandle
    
    If strSheetName = "" Then Exit Function
    
    strSheet = strSheetName
    If Right(strSheet, 1) <> "$" Then
        strSheet = strSheet & "$"
    End If
    
    rsTmp.Open strSheetName, gcnExcel, adOpenDynamic, adLockPessimistic, adCmdTableDirect
    If Not rsTmp.EOF Then
        Set OpenExcelSheet = rsTmp
    End If

    Exit Function
errHandle:
    If Err.Number = -2147217865 Then Exit Function
    MsgBox Err.Number & " " & Err.Description, vbQuestion, "升级阅读器"
End Function


Public Sub initRptList(ByRef objRpt As ReportControl, ByRef objImg As ImageList, ByVal txtFont As StdFont, ByVal blnEdit As Boolean)
    '初始化report控件
    
    Dim rptCol As ReportColumn

    Dim TextFont As StdFont
    '初始化列表
    
    With objRpt
        .SetImageList objImg
        
        '.AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '必须在列设置之前设置，才能生效
        '已读 = 0: 风险: 版本: 分类: 编号: 模块: 影响模块: 风险评估: 用户:需求: 说明: 关联问题: 备注:影响评估: 培训:  连接
        Set rptCol = .Columns.Add(mCol.阅读, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        rptCol.Icon = ICON_Mail
        
        Set rptCol = .Columns.Add(mCol.风险, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        rptCol.Icon = ICON_Importance
        
        Set rptCol = .Columns.Add(mCol.版本, "版本", 60, True): rptCol.Editable = False: rptCol.Groupable = True
        
        
        Set rptCol = .Columns.Add(mCol.分类, "问题类型", 50, True): rptCol.Editable = False: rptCol.Groupable = True
        
        
        Set rptCol = .Columns.Add(mCol.编号, "编号", 60, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        
        
        Set rptCol = .Columns.Add(mCol.模块, "模块", 120, True): rptCol.Editable = False: rptCol.Groupable = True
        
        
        Set rptCol = .Columns.Add(mCol.影响模块, "影响模块", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.影响类型, "影响类型", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.风险评估, "风险评估", 10, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.用户, "登记用户", 80, True): rptCol.Editable = False: rptCol.Groupable = True
        
        
        Set rptCol = .Columns.Add(mCol.需求, "用户需求", 10, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
        
        Set rptCol = .Columns.Add(mCol.说明, "修改说明", 10, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
        
        Set rptCol = .Columns.Add(mCol.关联问题, "关联问题", 80, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.备注, "备注", 80, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.培训, "培训情况", 70, False): rptCol.Editable = False: rptCol.Groupable = False
        'rptCol.Icon = ICON_FlagTrain
        
        Set rptCol = .Columns.Add(mCol.影响评估, "影响评估", 70, True): rptCol.Editable = blnEdit: rptCol.Groupable = False
        With rptCol.EditOptions
            .Constraints.Add "", 未填写
            .Constraints.Add "正面作用", 正面作用
            .Constraints.Add "负面作用", 负面作用
            .Constraints.Add "无影响", 无影响
            .ConstraintEdit = True
            .AddComboButton
        End With
        
        Set rptCol = .Columns.Add(mCol.连接, "连接", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.修改, "修改", 0, False): rptCol.Editable = True: rptCol.Groupable = False: rptCol.Visible = False
        
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .ShowGroupBox = True
        With .PaintManager
            
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 2
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            '.VerticalGridStyle = xtpGridSolid
            '正文字体
            Set TextFont = txtFont
            'TextFont.Size = 12
            Set .TextFont = TextFont
            Set .CaptionFont = TextFont
            
            '超连接字体
            Set fntUnderLine = .TextFont
            fntUnderLine.Underline = True
                        
        End With
        .PreviewMode = False
        .AllowEdit = True
        .EditOnClick = True
        .FocusSubItems = True
    
        '加入分组
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mCol.分类)
        .GroupsOrder(0).SortAscending = True
        .Columns.Find(mCol.分类).Visible = False
        .Populate
    End With
End Sub
