VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawLisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'常量定义
'######################################################################################################################
Private Type Property
    字体名称 As String
    字体大小 As Byte
    字体粗体 As Boolean
    字体斜体 As Boolean
    字体下划线 As Boolean
    字体前景色 As Long
End Type

Private mProperty As Property

Private mfrmMain As Object

Private mobjElement             As Object
Private mstrRegisterPath        As String
Private mclsLisReportLib        As clsLisReportLib
Private mclsLisPackage          As clsLisPackage

Public Event DataFormProgress(ByVal strProgressInfo As String)
Public Event AfterPrinted()

'第二层，获取原始数据并进行格式化
'######################################################################################################################
Public Function InitData(Optional ByVal strElement As String, Optional ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrRegisterPath = strRegisterPath
    
'    Call gobjComLib.InitCommon(gcnOracle)
    
    '创建格式化程序对象
    '------------------------------------------------------------------------------------------------------------------
    If strElement <> "" Then
        Set mobjElement = Nothing
        
        err = 0
        On Error Resume Next
        Set mobjElement = CreateObject(strElement)
        On Error GoTo 0
        
        If mobjElement Is Nothing Then
            ShowSimpleMsg "您指定的格式程序(" & strElement & ")不存在！"
            Exit Function
        End If
    End If
    
    '
    '------------------------------------------------------------------------------------------------------------------
    Set gobjDraw = frmPubResource.picDraw
    
    Set grsData = New ADODB.Recordset
    With grsData
        .Fields.Append "类别", adVarChar, 30
        .Fields.Append "页号", adInteger
        .Fields.Append "序号", adVarChar, 30
        .Fields.Append "调试", adInteger
        .Fields.Append "对象", adVarChar, 20
        .Fields.Append "内容", adVarChar, 4000
        .Fields.Append "X0", adVarChar, 20
        .Fields.Append "Y0", adVarChar, 20
        .Fields.Append "X1", adVarChar, 20
        .Fields.Append "Y1", adVarChar, 20
        .Fields.Append "B0", adVarChar, 20
        .Fields.Append "R0", adVarChar, 20
        .Fields.Append "字体", adVarChar, 20
        .Fields.Append "前景色", adVarChar, 20
        .Fields.Append "背景色", adVarChar, 20
        .Fields.Append "大小", adInteger
        .Fields.Append "粗体", adTinyInt
        .Fields.Append "斜体", adTinyInt
        .Fields.Append "下划线", adTinyInt
        .Fields.Append "横向对齐", adTinyInt
        .Fields.Append "纵向对齐", adTinyInt
        .Fields.Append "自动换行", adTinyInt
        .Fields.Append "线条宽度", adTinyInt
        .Fields.Append "线条类型", adTinyInt
        .Fields.Append "行数", adBigInt
        .Fields.Append "自动适应", adTinyInt
        .Fields.Append "旋转角度", adTinyInt
        .Fields.Append "标记", adVarChar, 30
        .Open
    End With
                                    
    '------------------------------------------------------------------------------------------------------------------
    Set grsPage = New ADODB.Recordset
    With grsPage
        .Fields.Append "总页", adInteger
        .Fields.Append "页号", adInteger
        .Fields.Append "虚拟总页", adInteger
        .Fields.Append "虚拟页号", adInteger
        .Fields.Append "页码计算", adTinyInt
        .Fields.Append "显示页眉", adTinyInt
        .Fields.Append "显示页脚", adTinyInt
        .Fields.Append "显示内容", adVarChar, 50
        .Open
    End With
    
    glngVirtualPages = 0
    
    '------------------------------------------------------------------------------------------------------------------
    Set grsList = New ADODB.Recordset
    With grsList
        .Fields.Append "目录性质", adTinyInt
        .Fields.Append "目录页号", adInteger
        .Fields.Append "目录名称", adVarChar, 200
        .Fields.Append "目录级数", adTinyInt
        .Open
    End With
    
    Set grsTempFile = New ADODB.Recordset
    With grsTempFile
        .Fields.Append "文件", adVarChar, 100
        .Open
    End With

    If Not mobjElement Is Nothing Then
        InitData = mobjElement.InitElement(gobjPaper, mstrRegisterPath)
    End If
        
End Function

Public Function PrintSet(frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    If Not mobjElement Is Nothing Then
        PrintSet = mobjElement.FormatSetup(frmMain, mstrRegisterPath)
    End If
    
End Function

Public Function ExportReport(ByRef frmReport As Object, Optional ByVal bytMode As Byte = 1, Optional ByVal strCaption As String = "", Optional ByVal strFile As String, Optional ByVal strPassWord As String, Optional ByVal strPage As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：bytMode=1，预览;=2,打印
    '返回：
    '******************************************************************************************************************
    
    If Not grsData Is Nothing Then
        If grsData.State = adStateOpen Then
            
            If frmReport.InitReport(grsData, grsPage, mstrRegisterPath) Then
                Call frmReport.ExportReport(bytMode, strCaption, strFile, strPassWord, strPage)
            End If
            
        End If
    End If
    
    ExportReport = True
    
End Function

Public Function BeginGetData() As Boolean
    '******************************************************************************************************************
    '功能：格式化数据
    '参数：
    '返回：
    '******************************************************************************************************************
    If Not mobjElement Is Nothing Then
        
        gobjRect.Page = 0
        
        BeginGetData = True
    End If
    
End Function

Public Function GetHeadData() As Boolean
    '******************************************************************************************************************
    '功能：格式化数据
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand

    If Not mobjElement Is Nothing Then

        GetHeadData = mobjElement.FormatHeadData
        
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetFootData() As Boolean
    '******************************************************************************************************************
    '功能：格式化数据
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand

    If Not mobjElement Is Nothing Then

        GetFootData = mobjElement.FormatFootData
        
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function AnalyseProperty(ByVal strProperty As String) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intPos As Integer
    Dim strTmp As String
    Dim aryProperty As Variant
    
    
    If strProperty = "" Then Exit Function
    aryProperty = Split(strProperty, "'")
    For intLoop = 0 To UBound(aryProperty)
        strTmp = CStr(aryProperty(intLoop))
        intPos = InStr(strTmp, "=")
        If intPos > 0 Then
            Select Case Trim(Left(strTmp, intPos - 1))
            Case "字体名称"
                mProperty.字体名称 = Trim(Mid(strTmp, intPos + 1))
            Case "字体大小"
                mProperty.字体大小 = Val(Mid(strTmp, intPos + 1))
            Case "字体前景色"
                mProperty.字体前景色 = Val(Mid(strTmp, intPos + 1))
            Case "字体粗体"
                mProperty.字体粗体 = (UCase(Trim(Mid(strTmp, intPos + 1))) = "TRUE")
            Case "字体斜体"
                mProperty.字体斜体 = (UCase(Trim(Mid(strTmp, intPos + 1))) = "TRUE")
            Case "字体下划线"
                mProperty.字体下划线 = (UCase(Trim(Mid(strTmp, intPos + 1))) = "TRUE")
            End Select
        End If
    Next
    
    AnalyseProperty = True
End Function

Public Function GetPageHeadData(ByVal strElement As String, Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：格式化数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objElement As Object
    Dim objRect As USERRECT
    Dim objFont As USERFONT
    Dim objSvrFont As USERFONT
    
    On Error GoTo errHand

    If Not mobjElement Is Nothing Then

        objSvrFont = mclsLisReportLib.Font
        objRect = mclsLisReportLib.RECT
        
        objFont.Name = mProperty.字体名称
        objFont.Size = 9
        objFont.Bold = False
        objFont.Italic = False
        objFont.Underline = False
        objFont.BackColor = USERCOLOR.白色
        objFont.ForeColor = USERCOLOR.黑色
        objFont.AutoFit = False
        objFont.HAlignment = 1
        objFont.VAlignment = 2
        objFont.MuliLine = False
        objFont.Rows = 1
        objFont.WordWarp = False

        mProperty.字体名称 = "宋体"
        mProperty.字体粗体 = False
        mProperty.字体斜体 = False
        mProperty.字体大小 = 9
        mProperty.字体下划线 = False
        mProperty.字体前景色 = 0
    
        Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "报告属性"))

        objFont.Name = mProperty.字体名称
        objFont.Size = mProperty.字体大小
        objFont.Bold = mProperty.字体粗体
        objFont.Italic = mProperty.字体斜体
        objFont.Underline = mProperty.字体下划线
        objFont.ForeColor = mProperty.字体前景色
        
        mobjElement.CurrentRect = objRect
        mobjElement.CurrentFont = objFont
        
        Call mobjElement.SetElementParameter(rsParameter)
        GetPageHeadData = mobjElement.GetPageHeadData
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPageFootData(ByVal strElement As String, Optional ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：格式化数据
    '参数：
    '返回：
    '******************************************************************************************************************
'    Dim objElement As Object
    Dim objRect As USERRECT
    Dim objFont As USERFONT
    Dim objSvrFont As USERFONT
    
    On Error GoTo errHand

'
    If Not mobjElement Is Nothing Then

        objSvrFont = mclsLisReportLib.Font
        objRect = mclsLisReportLib.RECT
        
        objFont.Name = mProperty.字体名称
        objFont.Size = 9
        objFont.Bold = False
        objFont.Italic = False
        objFont.Underline = False
        objFont.BackColor = USERCOLOR.白色
        objFont.ForeColor = USERCOLOR.黑色
        objFont.AutoFit = False
        objFont.HAlignment = 1
        objFont.VAlignment = 2
        objFont.MuliLine = False
        objFont.Rows = 1
        objFont.WordWarp = False

        mProperty.字体名称 = "宋体"
        mProperty.字体粗体 = False
        mProperty.字体斜体 = False
        mProperty.字体大小 = 9
        mProperty.字体下划线 = False
        mProperty.字体前景色 = 0
    
        Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "报告属性"))

        objFont.Name = mProperty.字体名称
        objFont.Size = mProperty.字体大小
        objFont.Bold = mProperty.字体粗体
        objFont.Italic = mProperty.字体斜体
        objFont.Underline = mProperty.字体下划线
        objFont.ForeColor = mProperty.字体前景色
        
        mobjElement.CurrentRect = objRect
        mobjElement.CurrentFont = objFont
        
        Call mobjElement.SetElementParameter(rsParameter)
        GetPageFootData = mobjElement.GetPageFootData()
                
        RaiseEvent DataFormProgress("aa")
        
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetElementData(ByVal strElement As String, Optional ByVal rsParameter As ADODB.Recordset, Optional ByVal blnNewPage As Boolean, Optional ByVal bytList As Byte = 0) As Boolean
    '******************************************************************************************************************
    '功能：格式化数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objElement As Object
    Dim objRect As USERRECT
    Dim objFont As USERFONT
    Dim objSvrFont As USERFONT
    
    On Error GoTo errHand

    '创建格式化程序对象
    '------------------------------------------------------------------------------------------------------------------
    If strElement <> "" Then
        Set objElement = Nothing
        err = 0
        
        On Error Resume Next
        Set objElement = CreateObject(strElement)
        On Error GoTo 0
        
        If objElement Is Nothing Then
            ShowSimpleMsg "您指定的格式程序(" & strElement & ")不存在！"
            Exit Function
        End If
    End If
    
    If Not objElement Is Nothing Then
        If objElement.InitElement(2, gcnOracle) Then
            
            objSvrFont = mclsLisReportLib.Font
            objRect = mclsLisReportLib.RECT
            
            objFont.Name = mProperty.字体名称
            objFont.Size = 9
            objFont.Bold = False
            objFont.Italic = False
            objFont.Underline = False
            objFont.BackColor = USERCOLOR.白色
            objFont.ForeColor = USERCOLOR.黑色
            objFont.AutoFit = False
            objFont.HAlignment = 1
            objFont.VAlignment = 2
            objFont.MuliLine = False
            objFont.Rows = 1
            objFont.WordWarp = False

            mProperty.字体名称 = "宋体"
            mProperty.字体粗体 = False
            mProperty.字体斜体 = False
            mProperty.字体大小 = 9
            mProperty.字体下划线 = False
            mProperty.字体前景色 = 0
        
            Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "报告属性"))
            Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "组件属性"))
            Call AnalyseProperty(mclsLisPackage.GetParameter(rsParameter, "内容属性"))
    
            objFont.Name = mProperty.字体名称
            objFont.Size = mProperty.字体大小
            objFont.Bold = mProperty.字体粗体
            objFont.Italic = mProperty.字体斜体
            objFont.Underline = mProperty.字体下划线
            objFont.ForeColor = mProperty.字体前景色
            
            If blnNewPage Then
                objRect.Page = objRect.Page + 1
                Call InsertPage(objRect.Page, 1, "", True, True)
                objRect.X0 = mclsLisReportLib.Paper.BorderLeft
                objRect.Y0 = mclsLisReportLib.Paper.BorderTop + mclsLisReportLib.Paper.PageHead
                objRect.Y1 = 0
            Else
                objRect.X0 = mclsLisReportLib.Paper.BorderLeft
                objRect.X1 = mclsLisReportLib.Paper.Width - mclsLisReportLib.Paper.BorderRight
            End If
            
            objElement.CurrentRect = objRect
            objElement.CurrentFont = objFont
            
            Call objElement.SetElementParameter(rsParameter)
            GetElementData = objElement.GetElementData()
                
            objRect = objElement.CurrentRect
            
            objRect.Y0 = objRect.Y1 + mclsLisReportLib.GetTwipsY(0.2)
            objRect.Y1 = 0
            mclsLisReportLib.RECT = objRect
            mclsLisReportLib.Font = objSvrFont
        End If
        
        Set objElement = Nothing
        
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function EndGetData() As Boolean
    '******************************************************************************************************************
    '功能：格式化数据
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If Not mobjElement Is Nothing Then

        If CalcPagesData(grsPage) = False Then GoTo EndHand
        
        EndGetData = True
    End If
    
    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    
EndHand:
    
End Function

Public Function FormatData() As Boolean
    '******************************************************************************************************************
    '功能：格式化数据
    '参数：
    '返回：
    '******************************************************************************************************************
    
    On Error GoTo errHand

    If Not mobjElement Is Nothing Then
        
        gobjCommFun.ShowFlash "正在组织数据，请稍候...."
        
        gobjRect.Page = 0
        If mobjElement.FormatHeadData = False Then GoTo EndHand
        If mobjElement.FormatFootData = False Then GoTo EndHand
        If mobjElement.FormatPageData = False Then GoTo EndHand
        If CalcPagesData(grsPage) = False Then GoTo EndHand
        
        gobjCommFun.StopFlash
        
        FormatData = True
        
    End If
    
    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    gobjCommFun.StopFlash
    
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    
EndHand:
    gobjCommFun.StopFlash
End Function

Private Function CalcPagesData(ByRef rsPage As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：重新计算虚拟页号、虚拟总页、总页
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim int虚拟页号 As Integer
    Dim int虚拟总页 As Integer
    Dim int总页 As Integer

    On Error GoTo errHand

    '------------------------------------------------------------------------------------------------------------------
    rsPage.Filter = "页码计算=1"
    int虚拟总页 = rsPage.RecordCount

    rsPage.Filter = ""
    int总页 = rsPage.RecordCount

    If rsPage.RecordCount > 0 Then
        rsPage.MoveFirst
        Do While Not rsPage.EOF
            rsPage("总页").Value = int总页
            rsPage("虚拟总页").Value = int虚拟总页

            If Val(rsPage("页码计算").Value) = 1 Then
                int虚拟页号 = int虚拟页号 + 1
                rsPage("虚拟页号").Value = int虚拟页号
            End If
            rsPage.MoveNext
        Loop
        rsPage.MoveFirst
    End If
    
    
    '同时更新目录页号为虚拟页号
    
    CalcPagesData = True

    Exit Function

    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function AdjustListPage(ByVal intFromPage As Integer, ByVal intToPage As Integer) As Boolean
    '******************************************************************************************************************
    '功能：将第intFromPage到所有的页移到第intToPage页之后
    '参数：
    '返回：
    '算法：将ingPage后的页号和总页号分别加上目录总页数
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intListPage As Integer

    On Error GoTo errHand
    
    If grsList.RecordCount > 0 And grsPage.RecordCount > 0 And intFromPage > 0 Then  'And intToPage > 0 允许直接放在第一页
        
        '先将目录页号变为负数
        '------------------------------------------------------
        grsPage.Filter = ""
        grsPage.Filter = "页号>=" & intFromPage
        If grsPage.RecordCount = 0 Then Exit Function
        intListPage = grsPage.RecordCount
        grsPage.MoveFirst
        Do While Not grsPage.EOF
            grsPage("页号").Value = 0 - grsPage("页号").Value
            grsPage.MoveNext
        Loop
        
        grsData.Filter = ""
        grsData.Filter = "页号>=" & intFromPage
        If grsData.RecordCount > 0 Then
            grsData.MoveFirst
            Do While Not grsData.EOF
                grsData("页号").Value = 0 - grsData("页号").Value
                grsData.MoveNext
            Loop
        End If
        
        '向后移动
        '------------------------------------------------------
        grsPage.Filter = ""
        grsPage.Filter = "页号>" & intToPage
        If grsPage.RecordCount > 0 Then
            grsPage.MoveFirst
            Do While Not grsPage.EOF
                grsPage("页号").Value = grsPage("页号").Value + intListPage
                grsPage.MoveNext
            Loop
        End If
        
        grsData.Filter = ""
        grsData.Filter = "页号>" & intToPage
        If grsData.RecordCount > 0 Then
            grsData.MoveFirst
            Do While Not grsData.EOF
                grsData("页号").Value = grsData("页号").Value + intListPage
                grsData.MoveNext
            Loop
        End If
        
        '向前移动
        '------------------------------------------------------
        grsPage.Filter = ""
        grsPage.Filter = "页号<0"
        grsPage.Sort = "页号 Desc"
        If grsPage.RecordCount > 0 Then
            grsPage.MoveFirst
            Do While Not grsPage.EOF
                intLoop = intLoop + 1
                intListPage = grsPage("页号").Value
                grsPage("页号").Value = intToPage + intLoop
                grsData.Filter = ""
                grsData.Filter = "页号=" & intListPage
                If grsData.RecordCount > 0 Then
                    grsData.MoveFirst
                    Do While Not grsData.EOF
                        grsData("页号").Value = intToPage + intLoop
                        grsData.MoveNext
                    Loop
                End If
                
                grsPage.MoveNext
            Loop
        End If

        grsPage.Filter = ""
        grsPage.Sort = "页号"
        
    End If
    
    AdjustListPage = True

    Exit Function

    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function AppendList(ByVal strListName As String, ByVal bytList As Byte, ByVal intPage As Integer) As Boolean
    '******************************************************************************************************************
    '功能：添加到目录索引
    '参数：
    '返回：
    '******************************************************************************************************************
    
    AppendList = AppendListData(strListName, bytList, intPage)
    
End Function

Public Property Get GetLastStartPage() As Integer
    '
    GetLastStartPage = gintStartPage
    
End Property

Private Sub Class_Initialize()
    Set grsData = New ADODB.Recordset
    Set grsPage = New ADODB.Recordset
    Set grsList = New ADODB.Recordset
    Set mclsLisReportLib = New clsLisReportLib
    Set mclsLisPackage = New clsLisPackage
    
    ParamInfo.系统名称 = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    ParamInfo.产品名称 = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")

    mProperty.字体名称 = "宋体"
    mProperty.字体粗体 = False
    mProperty.字体斜体 = False
    mProperty.字体大小 = 9
    mProperty.字体下划线 = False
    mProperty.字体前景色 = 0
    
End Sub

Public Sub DestoryDrawReport()
    
    Dim frmThis As Form
    
    On Error Resume Next
    
    Set grsData = Nothing
    Set grsPage = Nothing
    Set grsList = Nothing
    Set mclsLisReportLib = Nothing
    Set mclsLisPackage = Nothing
    Set gcnOracle = Nothing
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    
End Sub

Private Sub Class_Terminate()
    Set grsData = Nothing
    Set grsPage = Nothing
    Set grsList = Nothing
    Set mclsLisReportLib = Nothing
    Set gcnOracle = Nothing
    
    '删除临时文件
    If Not (grsTempFile Is Nothing) Then
        If grsTempFile.RecordCount > 0 Then
            grsTempFile.MoveFirst
            Do While Not grsTempFile.EOF
                If grsTempFile("文件").Value <> "" Then
                    If Dir(grsTempFile("文件").Value) <> "" Then
                        On Error Resume Next
                        Kill grsTempFile("文件").Value
                        On Error GoTo 0
                    End If
                End If
                grsTempFile.MoveNext
            Loop
        End If
    End If
    
    Set grsTempFile = Nothing
End Sub

Private Sub mfrmListReportView_AfterPrinted()
    RaiseEvent AfterPrinted
End Sub
