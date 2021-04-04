Attribute VB_Name = "mdSysConfig"
Option Explicit
'--------------------------------------------------------
'功  能：系统参数设置
'编制人：黄捷
'编制日期：2004.6.12
'过程函数清单：
'        subGetWWWLToVal()：            读取“预设窗宽窗位表”的内容到系统参数
'        subGetLayoutToVar（）：        从数据库中获取预设屏幕布局，填写到系统变量中。读取“预设屏幕布局”表。
'        subSaveScreenLayout（）：      将修改过的屏幕布局保存到系统参数和数据库中，将系统变量的内容保存到"预设屏幕布局"表中。
'        subGetMouseUsageToVar（）：    从数据库中读取鼠标用法设置的值到系统变量，读取“鼠标按钮分配”表的内容到系统变量
'        subGetInfoLabelToVar（）：     从数据库获取信息标注位置设置数据到系统变量，读取“图像信息表”的内容到系统变量
'        subGetDBDicomPrintToVar（）：  从数据库获取打印机的参数，填写到系统变量集合里面，读取“DICOM打印机设置”表。
'        subGetInterfaceParaToVar（）： 从数据库读取“影像界面参数表”的内容，并将其保存到系统参数中。
'        LoadBarSetup():                读取数据库上次保存的工具栏设置
'修改记录：
'    2005.06.29     黄捷    将参数定义移动到mdlSystemCortrol模块，并修改“影像界面参数表”等数据库表。
'-------------------------------------------------------

Public Sub subGetFilterToVal()
'------------------------------------------------
'功能：读取“影像滤镜模板”的内容到系统参数
'参数：无
'返回：无，直接修改系统参数
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    strSQL = "Select Id,影像类型,滤镜名称,增强强度增加,增强强度减少,增强幅度增加,增强幅度减少, " _
        & " 平滑增加,平滑减少 From 影像滤镜模板 order by 影像类型, ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取影像滤镜")
    
    '初始化预设滤镜系统变量
    ReDim aPresetFilter(Val(rsTemp.RecordCount)) As TPresetFilter
    
    '读取预设的滤镜设置
    i = 0
    With rsTemp
        While Not .EOF
            aPresetFilter(i).lngID = rsTemp!Id
            aPresetFilter(i).strname = rsTemp!滤镜名称
            aPresetFilter(i).strModality = rsTemp!影像类型
            aPresetFilter(i).intUnSharpEnhancementUp = Nvl(rsTemp!增强强度增加, 0)
            aPresetFilter(i).intUnSharpEnhancementDown = Nvl(rsTemp!增强强度减少, 0)
            aPresetFilter(i).intUnSharpLengthUp = Nvl(rsTemp!增强幅度增加, 0)
            aPresetFilter(i).intUnSharpLengthDown = Nvl(rsTemp!增强幅度减少, 0)
            aPresetFilter(i).intFilterLengthUp = Nvl(rsTemp!平滑增加, 0)
            aPresetFilter(i).intFilterLengthDown = Nvl(rsTemp!平滑减少, 0)
            i = i + 1
            .MoveNext
        Wend
    End With
    Exit Sub
err:
    If ErrCenter() = 1 Then
            Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subGetWWWLToVal()
'------------------------------------------------
'功能：读取“预设窗宽窗位表”的内容到系统参数
'参数：无
'返回：无，直接修改系统参数
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim strModality As String           '保存当前的影像类别
    Dim lngModalityCount As Long
    Dim blnUseDefaultSet As Boolean     '是否使用默认设置
    
    strModality = ""        '初始化当前影像类别
    blnUseDefaultSet = False
    
    '设置影像类型的数量
    If blLocalRun = True Then
        strSQL = "SELECT COUNT(影像类型) as iCount FROM (SELECT DISTINCT 影像类型 FROM 影像预设窗宽窗位)"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT COUNT( Distinct 影像类型) as iCount FROM 影像预设窗宽窗位 where 人员id =[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, glngUserID)
        If rsTemp!iCount = 0 Then
            blnUseDefaultSet = True
            strSQL = "SELECT COUNT( Distinct 影像类型) as iCount FROM 影像预设窗宽窗位 where 人员id =[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, CLng(0))
        End If
    End If
    lngModalityCount = rsTemp!iCount
    
    ''''''''初始化[窗宽窗位设置]的系统变量''''''''''''''''''''''''
    ReDim aPresetWinWL(3 To 12, lngModalityCount) As TPresetWinWL        ''保存预设窗宽窗位的数组，
                     ''允许的快捷键值为F3--F12，对应于数组的下标,2为自动窗宽窗位
    
    '将数据库内容保存到系统公共变量中
    '保存窗宽窗位的设置
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        strSQL = "SELECT ID,影像类型,快捷键,窗口名称,窗口英文名,窗宽,窗位,是否默认 FROM 影像预设窗宽窗位 ORDER BY 影像类型,快捷键"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT ID,影像类型,快捷键,窗口名称,窗口英文名,窗宽,窗位,是否默认 FROM 影像预设窗宽窗位 " & _
                 " where 人员id =[1] ORDER BY 影像类型,快捷键"
        If blnUseDefaultSet = True Then
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, CLng(0))
        Else
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, glngUserID)
        End If
    End If
    i = 0
    With rsTemp
        Do While Not .EOF
            If strModality = "" Or strModality <> !影像类型 Then
                strModality = !影像类型
                i = i + 1
                aPresetWinWL(3, i).strModality = strModality
            End If
            aPresetWinWL(!快捷键, i).bInUse = True
            aPresetWinWL(!快捷键, i).intDefault = !是否默认
            aPresetWinWL(!快捷键, i).strModality = strModality
            aPresetWinWL(!快捷键, i).strWinWLCName = !窗口名称
            aPresetWinWL(!快捷键, i).strWinWLEName = !窗口英文名
            aPresetWinWL(!快捷键, i).lngWinWidth = !窗宽
            aPresetWinWL(!快捷键, i).lngWinLevel = !窗位
            aPresetWinWL(!快捷键, i).lngID = !Id
            .MoveNext
        Loop
    End With
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subGetLayoutToVar(lngUserID As Long)
'------------------------------------------------
'功能：从数据库中获取预设屏幕布局，填写到系统变量中。读取“预设屏幕布局”表。
'参数：无
'返回：无
'上级函数或过程：
'下级函数或过程：
'引用的外部参数：
'编制人：黄捷
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim intModalityCount As Integer
    Dim rsTmp As New ADODB.Recordset
    
    
    '设置影像类型的数量
    If blLocalRun = True Then
        strSQL = "SELECT COUNT(影像类型) as iCount FROM (SELECT DISTINCT 影像类型 FROM 影像屏幕布局)"
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT COUNT(影像类型) as iCount FROM 影像屏幕布局 where 人员ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
    End If
    intModalityCount = rsTmp!iCount
    
    ReDim aModifiedPresetLayout(intModalityCount) As TModifiedPresetLayout
    ReDim aPresetLayout(intModalityCount) As TModifiedPresetLayout         ''保存预设屏幕布局的数组
    
    '将数据库内容保存到系统公共变量中
    If blLocalRun = True Then
        strSQL = "SELECT 影像类型,自动序列布局,自动图像布局,序列行数,序列列数,图像行数,图像列数" & _
                ",自动反白,显示病人信息,选择定位线,选择序列同步,插值模式,图像排序 FROM 影像屏幕布局 ORDER BY 影像类型"
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT 影像类型,自动序列布局,自动图像布局,序列行数,序列列数,图像行数,图像列数" & _
                ",自动反白,显示病人信息,选择定位线,选择序列同步,插值模式,图像排序 FROM 影像屏幕布局 Where 人员ID = [1] ORDER BY 影像类型"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
    End If
    i = 1
    With rsTmp
        While Not .EOF
               '将数据库内容保存到系统公共变量中
               aPresetLayout(i).strModality = !影像类型
               aPresetLayout(i).bSeriesAutoFormat = IIf(IsNull(!自动序列布局), 0, !自动序列布局)
               aPresetLayout(i).lngSeriesColumns = IIf(IsNull(!序列列数), 2, !序列列数)
               aPresetLayout(i).lngSeriesRows = IIf(IsNull(!序列行数), 1, !序列行数)
               aPresetLayout(i).bImageAutoFormat = IIf(IsNull(!自动图像布局), 0, !自动图像布局)
               aPresetLayout(i).lngImageColumns = IIf(IsNull(!图像列数), 2, !图像列数)
               aPresetLayout(i).lngImageRows = IIf(IsNull(!图像行数), 1, !图像行数)
               aPresetLayout(i).bInvert = IIf(IsNull(!自动反白), 0, !自动反白)
               aPresetLayout(i).bShowPatientInfo = IIf(IsNull(!显示病人信息), 0, !显示病人信息)
               aPresetLayout(i).bAutoSelectReferenceLine = IIf(IsNull(!选择定位线), 0, !选择定位线)
               aPresetLayout(i).bAutoSelectSeriesSyn = IIf(IsNull(!选择序列同步), 0, !选择序列同步)
               aPresetLayout(i).lngInterpolationMode = IIf(IsNull(!插值模式), 0, !插值模式)
               aPresetLayout(i).lngImageSort = IIf(IsNull(!图像排序), 0, !图像排序)
               i = i + 1
               .MoveNext
        Wend
    End With
End Sub

Public Sub subGetImageShutterToVar(lngUserID As Long)
    '------------------------------------------------
'功能：从数据库中获取预设图像消隐，填写到系统变量中。读取“图像消隐表”。
'参数：无
'返回：无
'上级函数或过程：
'下级函数或过程：
'引用的外部参数：
'编制人：黄捷
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim intModalityCount As Integer
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errh
    
    If blLocalRun = True Then
        strSQL = "select count(影像类型) as iCount from 影像图像消隐表 "
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "select count(影像类型) as iCount from 影像图像消隐表 where 人员ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
    End If
    
    intModalityCount = rsTmp!iCount
    
    ReDim aModifiedImageShutter(intModalityCount) As TImageShutter
    ReDim aImageShutter(intModalityCount) As TImageShutter          ''保存图像消隐的数组
    
    If blLocalRun = True Then
        strSQL = "SELECT 影像类型,消隐类型,圆心X,圆心Y,圆形半径,矩形左边界,矩形右边界,矩形上边界" & _
                ",矩形下边界,多边形顶点,消隐颜色 FROM 影像图像消隐表  ORDER BY 影像类型"
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        '将数据库内容保存到系统公共变量中
        strSQL = "SELECT 影像类型,消隐类型,圆心X,圆心Y,圆形半径,矩形左边界,矩形右边界,矩形上边界" & _
                ",矩形下边界,多边形顶点,消隐颜色 FROM 影像图像消隐表  where 人员id = [1] ORDER BY 影像类型"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
    End If
    i = 1
    With rsTmp
        While Not .EOF
               '将数据库内容保存到系统公共变量中
               aImageShutter(i).strModality = !影像类型
               aImageShutter(i).intShutterType = IIf(IsNull(!消隐类型), 0, !消隐类型)
               aImageShutter(i).intCenterX = IIf(IsNull(!圆心X), 0, !圆心X)
               aImageShutter(i).intCenterY = IIf(IsNull(!圆心Y), 0, !圆心Y)
               aImageShutter(i).intRadius = IIf(IsNull(!圆形半径), 0, !圆形半径)
               aImageShutter(i).intRectLeft = IIf(IsNull(!矩形左边界), 0, !矩形左边界)
               aImageShutter(i).intRectRight = IIf(IsNull(!矩形右边界), 0, !矩形右边界)
               aImageShutter(i).intRectUpper = IIf(IsNull(!矩形上边界), 0, !矩形上边界)
               aImageShutter(i).intRectLower = IIf(IsNull(!矩形下边界), 0, !矩形下边界)
               aImageShutter(i).strVertices = IIf(Not IsNull(!多边形顶点), !多边形顶点, "")
               aImageShutter(i).lngColor = IIf(IsNull(!消隐颜色), 0, !消隐颜色)
               i = i + 1
               .MoveNext
        Wend
    End With
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subSaveImgShutter()
'------------------------------------------------
'功能：将修改过的图像消隐参数保存到系统参数和数据库中，将变量的内容保存到"图像消隐表"表中。
'参数：无
'返回：
'上级函数或过程：
'下级函数或过程：
'引用的外部参数：
'编制人：黄捷
'------------------------------------------------

    Dim i As Integer
    Dim strSQL As String
    Dim intModalityCount As Integer
    intModalityCount = UBound(aModifiedImageShutter)
    
    On Error GoTo errh
    
    For i = 1 To intModalityCount
        If aModifiedImageShutter(i).bModified Then
            '将设置的结果保存到系统变量
            aImageShutter(i).intCenterX = aModifiedImageShutter(i).intCenterX
            aImageShutter(i).intCenterY = aModifiedImageShutter(i).intCenterY
            aImageShutter(i).intRadius = aModifiedImageShutter(i).intRadius
            aImageShutter(i).intRectLeft = aModifiedImageShutter(i).intRectLeft
            aImageShutter(i).intRectLower = aModifiedImageShutter(i).intRectLower
            aImageShutter(i).intRectRight = aModifiedImageShutter(i).intRectRight
            aImageShutter(i).intRectUpper = aModifiedImageShutter(i).intRectUpper
            aImageShutter(i).intShutterType = aModifiedImageShutter(i).intShutterType
            aImageShutter(i).lngColor = aModifiedImageShutter(i).lngColor
            aImageShutter(i).strModality = aModifiedImageShutter(i).strModality
            aImageShutter(i).strVertices = aModifiedImageShutter(i).strVertices
            
            
            If blLocalRun = True Then
                '保存改变了的图像消隐设置到数据库，从修改记录中进行保存
                strSQL = "UPDATE 影像图像消隐表 SET 消隐类型 = " & aImageShutter(i).intShutterType _
                         & " , 圆心X = " & aImageShutter(i).intCenterX _
                         & " , 圆心Y = " & aImageShutter(i).intCenterY _
                         & " , 圆形半径 = " & aImageShutter(i).intRadius _
                         & " , 矩形左边界 = " & aImageShutter(i).intRectLeft _
                         & " , 矩形右边界 = " & aImageShutter(i).intRectRight _
                         & " , 矩形上边界 = " & aImageShutter(i).intRectUpper _
                         & " , 矩形下边界 = " & aImageShutter(i).intRectLower _
                         & " , 多边形顶点 = '" & aImageShutter(i).strVertices & "'" _
                         & " , 消隐颜色 = " & aImageShutter(i).lngColor _
                         & " where 影像类型 = '" & aImageShutter(i).strModality & "'"
                cnAccess.Execute strSQL, , adCmdText
            Else
                strSQL = "ZL_影像图像消隐表_UPDATE(" & glngUserID & ",'" & aImageShutter(i).strModality & "','" & _
                aImageShutter(i).intShutterType & "'," & aImageShutter(i).intCenterX & "," & aImageShutter(i).intCenterY & _
                "," & aImageShutter(i).intRadius & "," & aImageShutter(i).intRectLeft & "," & aImageShutter(i).intRectRight & _
                "," & aImageShutter(i).intRectUpper & "," & aImageShutter(i).intRectLower & ",'" & aImageShutter(i).strVertices & _
                "'," & aImageShutter(i).lngColor & ")"
                zlDatabase.ExecuteProcedure strSQL, App.ProductName
            End If
        End If
    Next
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subSaveScreenLayout()
'------------------------------------------------
'功能：将修改过的屏幕布局保存到系统参数和数据库中，将系统变量的内容保存到"预设屏幕布局"表中。
'参数：无
'返回：
'上级函数或过程：
'下级函数或过程：
'引用的外部参数：
'编制人：黄捷
'------------------------------------------------

    Dim i As Integer
    Dim strSQL As String
    Dim intModalityCount As Integer
    intModalityCount = UBound(aModifiedPresetLayout)
    
    On Error GoTo errh
    
    For i = 1 To intModalityCount
'        If aModifiedPresetLayout(i).bModified Then
            '将设置的结果保存到系统变量
            aPresetLayout(i).bSeriesAutoFormat = aModifiedPresetLayout(i).bSeriesAutoFormat
            aPresetLayout(i).lngSeriesColumns = aModifiedPresetLayout(i).lngSeriesColumns
            aPresetLayout(i).lngSeriesRows = aModifiedPresetLayout(i).lngSeriesRows
            aPresetLayout(i).bImageAutoFormat = aModifiedPresetLayout(i).bImageAutoFormat
            aPresetLayout(i).lngImageColumns = aModifiedPresetLayout(i).lngImageColumns
            aPresetLayout(i).lngImageRows = aModifiedPresetLayout(i).lngImageRows
            aPresetLayout(i).bInvert = aModifiedPresetLayout(i).bInvert
            aPresetLayout(i).bShowPatientInfo = aModifiedPresetLayout(i).bShowPatientInfo
            aPresetLayout(i).bAutoSelectReferenceLine = aModifiedPresetLayout(i).bAutoSelectReferenceLine
            aPresetLayout(i).bAutoSelectSeriesSyn = aModifiedPresetLayout(i).bAutoSelectSeriesSyn
            aPresetLayout(i).lngInterpolationMode = aModifiedPresetLayout(i).lngInterpolationMode
            aPresetLayout(i).lngImageSort = aModifiedPresetLayout(i).lngImageSort
            
            If blLocalRun = True Then
                '保存改变了的布局到数据库，从修改记录中进行保存
                strSQL = "UPDATE 影像屏幕布局 SET 自动序列布局=" & _
                         IIf(aModifiedPresetLayout(i).bSeriesAutoFormat, 1, 0) & _
                         ",自动图像布局 = " & IIf(aModifiedPresetLayout(i).bImageAutoFormat, 1, 0) & _
                         ",序列行数 = " & aModifiedPresetLayout(i).lngSeriesRows & ",序列列数 = " & _
                         aModifiedPresetLayout(i).lngSeriesColumns & ",图像行数 = " & _
                         aModifiedPresetLayout(i).lngImageRows & ",图像列数 = " & _
                         aModifiedPresetLayout(i).lngImageColumns & ",自动反白= " & _
                         aModifiedPresetLayout(i).bInvert & ",显示病人信息= " & _
                         aModifiedPresetLayout(i).bShowPatientInfo & ",选择定位线 = " & _
                         aModifiedPresetLayout(i).bAutoSelectReferenceLine & ",选择序列同步=" & _
                         aModifiedPresetLayout(i).bAutoSelectSeriesSyn & ",插值模式=" & _
                         aModifiedPresetLayout(i).lngInterpolationMode & ",图像排序=" & _
                         aModifiedPresetLayout(i).lngImageSort & " WHERE 影像类型='" & _
                         aModifiedPresetLayout(i).strModality & "'"
                cnAccess.Execute strSQL, , adCmdText
            Else
                strSQL = "ZL_影像屏幕布局_UPDATE(" & glngUserID & ",'" & aModifiedPresetLayout(i).strModality & "'," & _
                IIf(aModifiedPresetLayout(i).bSeriesAutoFormat, 1, 0) & "," & IIf(aModifiedPresetLayout(i).bImageAutoFormat, 1, 0) & _
                "," & aModifiedPresetLayout(i).lngSeriesRows & "," & aModifiedPresetLayout(i).lngSeriesColumns & _
                "," & aModifiedPresetLayout(i).lngImageRows & "," & aModifiedPresetLayout(i).lngImageColumns & "," & _
                CInt(aModifiedPresetLayout(i).bInvert) & "," & CInt(aModifiedPresetLayout(i).bShowPatientInfo) & _
                "," & CInt(aModifiedPresetLayout(i).bAutoSelectReferenceLine) & "," & CInt(aModifiedPresetLayout(i).bAutoSelectSeriesSyn) & _
                "," & aModifiedPresetLayout(i).lngInterpolationMode & "," & aModifiedPresetLayout(i).lngImageSort & ")"
                zlDatabase.ExecuteProcedure strSQL, App.ProductName
                
            End If
'        End If
    Next
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub


Public Sub subGetMouseUsageToVar(lngUserID As Long)
'------------------------------------------------
'功能：从数据库中读取鼠标用法设置的值到系统变量，读取“鼠标按钮分配”表的内容到系统变量
'参数：无
'返回：无
'上级函数或过程：frmViewer.Form_Load
'下级函数或过程：无
'引用的外部参数：cMouseUsage
'编制人：黄捷
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim clsOneMouseUsage As clsMouseUsage
    Dim iDrawLabel As Integer
    Dim strField As Variant
    
    On Error GoTo errh
    
    For i = 1 To cMouseUsage.Count
        cMouseUsage.Remove 1
    Next
    
    If blLocalRun = True Then
        strSQL = "select ID,直线,矩形,椭圆,箭头,多边形,多边线,角度,文字,穿梭定位,窗宽窗位,漫游,缩放,裁剪_标注调整," & _
                 " 自适应调窗,三维鼠标,画标注 from 影像鼠标按钮分配 "
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "select 人员ID,直线,矩形,椭圆,箭头,多边形,多边线,角度,文字,穿梭定位,窗宽窗位,漫游,缩放,裁剪_标注调整," & _
                 " 自适应调窗,三维鼠标,画标注 from 影像鼠标按钮分配 where 人员id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngUserID)
        If rsTmp.EOF = True Then
            strSQL = "select 人员ID,直线,矩形,椭圆,箭头,多边形,多边线,角度,文字,穿梭定位,窗宽窗位,漫游,缩放,裁剪_标注调整," & _
                 " 自适应调窗,三维鼠标,画标注 from 影像鼠标按钮分配 where 人员id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, CLng(0))
            
        End If
    End If
    
    iDrawLabel = -1
    If rsTmp.EOF = True Then Exit Sub
    For i = 1 To rsTmp.Fields.Count - 1
        '将数据库内容保存到系统公共变量中
        Set clsOneMouseUsage = New clsMouseUsage
        strField = Split(rsTmp(i).Value, ",")
        clsOneMouseUsage.lngFuncNo = strField(0)
        clsOneMouseUsage.lngMouseKey = strField(1)
        clsOneMouseUsage.lngShift = strField(2)
        clsOneMouseUsage.bSelected = strField(3)
        clsOneMouseUsage.strProgramName = strField(4)
        clsOneMouseUsage.ButtomID = strField(5)
        clsOneMouseUsage.strShowName = rsTmp(i).Name
        
        cMouseUsage.Add clsOneMouseUsage, CStr(clsOneMouseUsage.lngFuncNo)
        If clsOneMouseUsage.lngFuncNo = lngDrawLabelFuncNo Then
             iDrawLabel = cMouseUsage.Count
        End If
    Next
    
    '填写鼠标当前选中状态，到画标注按钮中
    If cMouseUsage(CStr(lngDrawLabelCurrent)).bSelected Then
        cMouseUsage(iDrawLabel).bSelected = True
    End If
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subGetInfoLabelToVar()
'------------------------------------------------
'功能：从数据库获取信息标注位置设置数据到系统变量，读取“图像信息表”的内容到系统变量
'参数：无
'返回：无
'上级函数或过程：frmViewer.Form_Load
'下级函数或过程：无
'引用的外部参数：aInfoLabelLocate
'编制人：黄捷
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    
    '设置信息标注的数量
    If blLocalRun = True Then
        strSQL = "SELECT COUNT(id) as iCount FROM 影像图像信息表 WHERE 常用=-1"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT COUNT(id) as iCount FROM 影像图像信息表 WHERE 常用=-1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    End If
    lngInfoLabelCount = rsTemp!iCount
    
    ''''''''初始化[信息标注设置]的系统变量
    ReDim aInfoLabelLocate(lngInfoLabelCount) As TInfoLabelLocate  ''保存信息标注的位置
    
    On Error GoTo errh
    '从数据库获取信息标注位置
    If blLocalRun = True Then
        strSQL = "SELECT id,开始地址,结束地址,英文简称,中文简称,被选用,位置,角内序号,可导出 FROM 影像图像信息表 WHERE 常用=-1"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT id,开始地址,结束地址,英文简称,中文简称,被选用,位置,角内序号,可导出 FROM 影像图像信息表 WHERE 常用=-1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    End If
    i = 1
    With rsTemp
        .MoveFirst
        While Not .EOF
            aInfoLabelLocate(i).lngID = !Id
            aInfoLabelLocate(i).bUsed = IIf(IsNull(!被选用), False, IIf(!被选用 = -1, True, False))
            aInfoLabelLocate(i).strGroup = !开始地址
            aInfoLabelLocate(i).strElement = !结束地址
            aInfoLabelLocate(i).lngLocation = IIf(IsNull(!位置) = True, 0, !位置)
            aInfoLabelLocate(i).lngOrder = IIf(IsNull(!角内序号), 0, !角内序号)
            aInfoLabelLocate(i).strCName = IIf(IsNull(!中文简称), "", !中文简称)
            aInfoLabelLocate(i).strEName = IIf(IsNull(!英文简称), "", !英文简称)
            aInfoLabelLocate(i).blnIsExport = IIf(IsNull(!可导出), 0, !可导出)
            i = i + 1
            .MoveNext
        Wend
    End With
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subGetDBDicomPrintToVar()
'------------------------------------------------
'功能：从数据库获取打印机的参数，填写到系统变量集合里面，读取“DICOM打印机设置”表。
'参数：无
'返回：无
'上级函数或过程：frmViewer.Form_Load
'下级函数或过程：无
'引用的外部参数：cDICOMPrinter
'编制人：黄捷
'------------------------------------------------
    '初始化集合
    Dim i As Integer
    Dim clsOnePrinter As clsDicomPrint
    
    For i = 1 To cDICOMPrinter.Count
        cDICOMPrinter.Remove (1)
    Next
    '从数据库读取信息
    Dim strSQL As String
    
    On Error GoTo errh
    
    cstrPrintAE = GetSetting("ZLSOFT", "公共模块\zlPacsCore", "本机AE", "ZLPACS")
    blnPrintOkEcho = GetSetting("ZLSOFT", "公共模块\zlPacsCore", "打印成功后提示", "False")
     
    If blLocalRun = True Then
        strSQL = "SELECT ID,打印机名,IP地址,端口号,AE名称,打印格式,优先级,打印份数,介质,方向," & _
                 "胶片规格,选用片盒,分辨率,放大模式,平滑模式,修整,最小密度,最大密度,空白密度," & _
                 "边框密度,极性,图像位数,用户AE名称 FROM 影像打印机设置"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "SELECT ID,打印机名,IP地址,端口号,AE名称,打印格式,优先级,打印份数,介质,方向," & _
                 "胶片规格,选用片盒,分辨率,放大模式,平滑模式,修整,最小密度,最大密度,空白密度," & _
                 "边框密度,极性,图像位数,用户AE名称,图像边框宽度,图片分辨率 FROM 影像打印机设置"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    End If
    '将数据库内容填写到系统变量集合里面
    With rsTemp
        While Not .EOF
            Set clsOnePrinter = New clsDicomPrint
            clsOnePrinter.lngID = !Id
            clsOnePrinter.lngCopies = !打印份数
            clsOnePrinter.lngPort = !端口号
            clsOnePrinter.strAETitle = !AE名称
            clsOnePrinter.strBorderDensity = IIf(IsNull(!边框密度), "", !边框密度)
            clsOnePrinter.strEmptyDensity = IIf(IsNull(!空白密度), "", !空白密度)
            clsOnePrinter.strFilmBox = IIf(IsNull(!选用片盒), "", !选用片盒)
            clsOnePrinter.strFilmSize = IIf(IsNull(!胶片规格), "", !胶片规格)
            clsOnePrinter.strFormat = IIf(IsNull(!打印格式), "", !打印格式)
            clsOnePrinter.strIPAddress = IIf(IsNull(!IP地址), "", !IP地址)
            clsOnePrinter.strMagnification = IIf(IsNull(!放大模式), "", !放大模式)
            clsOnePrinter.strMaxDensity = IIf(IsNull(!最大密度), "", !最大密度)
            clsOnePrinter.strMedium = IIf(IsNull(!介质), "", !介质)
            clsOnePrinter.strMinDensity = IIf(IsNull(!最小密度), "", !最小密度)
            clsOnePrinter.strname = IIf(IsNull(!打印机名), "", !打印机名)
            clsOnePrinter.strOrientation = IIf(IsNull(!方向), "", !方向)
            clsOnePrinter.strPolarity = IIf(IsNull(!极性), "", !极性)
            clsOnePrinter.strPriority = IIf(IsNull(!优先级), "", !优先级)
            clsOnePrinter.strResolution = IIf(IsNull(!分辨率), "", !分辨率)
            clsOnePrinter.strSmooth = IIf(IsNull(!平滑模式), "", !平滑模式)
            clsOnePrinter.strTrim = IIf(IsNull(!修整), "", !修整)
            clsOnePrinter.lngBitDepth = IIf(IsNull(!图像位数), 8, !图像位数)
            clsOnePrinter.strSCUAETitle = IIf(IsNull(!用户AE名称), cstrPrintAE, !用户AE名称)
            clsOnePrinter.lngImageBorderWidth = Val(Nvl(!图像边框宽度, 1))
            If clsOnePrinter.lngImageBorderWidth < 1 Or clsOnePrinter.lngImageBorderWidth > 99 Then
                clsOnePrinter.lngImageBorderWidth = 1
            End If
            clsOnePrinter.intImageResolution = Val(Nvl(!图片分辨率, 300))
            If clsOnePrinter.intImageResolution < 10 Or clsOnePrinter.intImageResolution > 999 Then
                clsOnePrinter.intImageResolution = 300
            End If
            cDICOMPrinter.Add clsOnePrinter, clsOnePrinter.strname
            .MoveNext
        Wend
    End With
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subGetInterfaceParaToVar(Optional blDefaultVal As Long)
'------------------------------------------------
'功能：从数据库读取“影像界面参数表”的内容，并将其保存到系统参数中。
'参数：     blDefaultVal 是否取缺省值 =True 取缺省值 =False 取正常值
'返回：无
'------------------------------------------------
    Dim strDefaultVal As String     ''提取字段名
    Dim StrTmp As String
    
    '初始化系统变量
    lngSelectedImageBorderColor = 1 ''选中图像边框颜色
    lngCurrentImageBorderColor = 1  ''选中图像边框颜色
    lngCurrentSeriesBorderColor = 1 ''选中序列边框颜色
    lngSelectImageForeColour = 1    ''选中图像标识填充色
    lngPeriodColor = 1              ''选择句柄颜色
    lngReferenceLineColor = 1       ''定位线颜色
    lngViewerBackColor = 1          ''Viewer背景颜色
    lngProgramBackColor = 1         ''程序背景颜色

    lngSelectedImageBorderLineStyle = 0 ''选中图像边框线形
    lngSelectedImageBorderLineWidth = 0 ''选中图像边框线宽度
    lngCurrentImageBorderLineStyle = 0 ''当前图像边框线形
    lngCurrentImageBorderLineWidth = 0 ''当前图像边框线宽度
    lngImageIdentifierSize = 0         ''图像选择标记大小
    intPeriodSize = 0                  ''选择句柄大小
    lngReferenceLineStyle = 0          ''定位线线形
    lngReferenceLineSpacing = 1        ''定位线间距

    intSpaceSize = 0                          ''序列之间的间隔宽度、高度
    intMaxAreaX = 0                           ''横向最多可划分的区域
    intMaxAreaY = 0                           ''纵向最多可划分的区域
    lngCellSpacing = 0                        ''图像间距
    blnDsipSpilthBorder = False               ''多余边框是否显示
    blnDockMiniImage = False                  ''缩略图停靠于菜单下
    blnShowMiniImageInfo = True               ''缩略图中是否显示图像信息
    blnShowMPRLine = True                     ''MPR显示辅助线，默认是True
    blnSquareFrame = True                     ''正方形框选
    blnShowPrintTag = False                   ''是否显示胶片打印标记
    blnPrintFilmBeep = False                  ''胶片打印时是否提示声音，包括添加胶片，打印
    
    '填充各种颜色
    lngLabelColor = 1             ''标注显示色，白色
    lngLabelSelectedColor = 1     ''标注选中色，红色
    lngRulerLeftColor = 1         ''标尺颜色
    
    lngLabelLineStyleNorm = 0      ''线型
    lngLabelLineWidthNorm = 1      ''线宽
    lngLabelFontSize = 16           ''字体大小
    '标注设置
    lngWinWidthLevelLocation = 1    '' 窗宽窗位位置
    '关联文字的显示设置
    bROIArea = False      ''显示面积
    bROIMean = False      ''显示平均值
    bROIStandardDeviation = False  ''显示均方差
    bROILength = False    ''显示周长
    bROIMax = False       ''显示最大值
    bROIMin = False       ''显示最小值
    bROITextChinese = False                    ''测量使用中文
    intTextoOffX = 0                           ''标注文字的偏移量
    intTextoOffY = 0                           ''标注文字的偏移量
    blnLabelTextScaleFontSize = False          ''标注文字大小是否随着图像一起缩放
    '体位标记设置
    blnAnatomicMarkersLeft = False     ''是否显示左边体位标记
    blnAnatomicMarkersRight = False      ''是否显示右边体位标记
    blnAnatomicMarkersTop = False       ''是否显示上边体位标记
    blnAnatomicMarkersBottom = False     ''是否显示下边体位标记
    blnChinaMark = False                   ''是否采用汉字显示体位标记
    '标尺设置
    blnRulerDsipLeft = False             ''是否显示左边标尺
    blnRulerDsipRight = False           ''是否显示右边标尺
    blnRulerDsipTop = False             ''是否显示上边标尺
    blnRulerDsipBottom = False          ''是否显示下边标尺
    intRulerLeft = 0                     ''标尺左边距
    intRulerTop = 0                       ''标尺上边距
    intRulerWidth = 0                   ''标尺宽度
    intRulerHeight = 0                   ''标尺高度
    intRulerLineWidth = 0               ''标尺线宽
    '工具栏设置
    intToolBarIconSize = 32             ''工具栏图标大小
    intToolBarPosition = 1              ''工具栏位置
    blToolBarHide = True                ''工具栏显示
    
    '血管狭窄测量
    intStandardThreshold = 50           ''正常血管测量的阈值
    intNarrowThreshold = 50             ''狭窄血管测量的阈值
    intVasEdgeWidth = 10                   ''血管狭窄测量中显示血管壁短直线的宽度
    
    '鼠标设置
    lngStackStep = 10                   ''鼠标穿梭步长
    lngCruiseStep = 10                  ''鼠标漫游步长
    lngWidthLevelStep = 10              ''鼠标调窗步长
    lngZoomStep = 10                    ''鼠标缩放步长
    intMouseWheelRoll = 0               ''鼠标滚轮
    
    
    '读取数据库的值
    Dim strSQL As String
    If blLocalRun = True Then
        strSQL = "select * from 影像界面参数表"
        Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
    Else
        strSQL = "select * from 影像界面参数表 where 人员id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, blDefaultVal)
        
        If rsTemp.EOF = True Then
            strSQL = "Select * from 影像界面参数表 where 人员id = 0 "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
        End If
    End If
    
    lngSelectedImageBorderColor = rsTemp("正常图像边框颜色")                                ''选中图像边框颜色
    lngCurrentImageBorderColor = rsTemp("选中图像边框颜色")                                 ''选中图像边框颜色
    lngCurrentSeriesBorderColor = rsTemp("选中序列边框颜色")                                ''当前（未选中）序列边框颜色
    lngSelectImageForeColour = rsTemp("图像标记颜色")                                       ''选中图像标识填充色
    lngPeriodColor = rsTemp("标注选择句柄颜色")                                             ''选择句柄颜色
    lngReferenceLineColor = rsTemp("定位线颜色")                                            ''定位线颜色
    lngViewerBackColor = rsTemp("背景颜色")                                                 ''Viewer背景颜色
    lngProgramBackColor = rsTemp("程序背景颜色")                                            ''程序背景颜色
    lngSelectedImageBorderLineStyle = rsTemp("正常图像边框线型")                            ''选中图像边框线形
    lngSelectedImageBorderLineWidth = rsTemp("正常图像边框线宽")                            ''选中图像边框线宽度
    lngCurrentImageBorderLineStyle = rsTemp("选中图像边框线型")                             ''当前图像边框线型
    lngCurrentImageBorderLineWidth = rsTemp("选中图像边框线宽")                             ''当前图像边框线宽
    lngImageIdentifierSize = rsTemp("图像标记大小")                                         ''图像选择标记大小
    intPeriodSize = rsTemp("标注选择句柄大小")                                              ''选择句柄大小
    lngReferenceLineStyle = rsTemp("定位线线型")                                            ''定位线线形
    lngReferenceLineSpacing = rsTemp("定位线间距")                                          ''定位线间距
    intSpaceSize = rsTemp("序列间间隔")                                                     ''序列之间的间隔宽度、高度
    intMaxAreaX = rsTemp("横向最大序列")                                                    ''横向最多可划分的区域
    intMaxAreaY = rsTemp("纵向最大序列")                                                    ''纵向最多可划分的区域
    If intMaxAreaX < 1 Or intMaxAreaX > 8 Then intMaxAreaX = 8
    If intMaxAreaY < 1 Or intMaxAreaY > 8 Then intMaxAreaY = 8
    lngCellSpacing = rsTemp("图像间距")                                                     ''图像间距
    blnDsipSpilthBorder = IIf(rsTemp("显示多余边框") = -1, True, False)                     ''多余边框是否显示
    bShowFilmConfig = IIf(rsTemp("直接照相") = -1, True, False)                             ''是否直接照相，不显示胶片设置窗口
    intStatusBarFontSize = rsTemp("状态栏字体大小")                                         ''状态态字体大小
    blnShowPrintTag = IIf(rsTemp("显示打印标记") = -1, True, False)                         ''是否显示胶片打印标记
    '获取鼠标信息
    lngStackStep = rsTemp("鼠标穿梭步长")
    lngCruiseStep = rsTemp("鼠标漫游步长")
    lngWidthLevelStep = rsTemp("鼠标调窗步长")
    lngZoomStep = rsTemp("鼠标缩放步长")
    intMouseWheelRoll = Nvl(rsTemp("鼠标滚轮操作"), 0)
    '从数据库获取病人信息标注的显示设置
    lngPatientInfoInvisibleSize = rsTemp("病人信息显示最小值")
    lngpatientInfoColor = rsTemp("病人信息颜色")
    blnpatientInfoScaleFontSize = IIf(rsTemp("病人信息随图像缩放") = -1, True, False)
    
    StrTmp = rsTemp("病人信息字体")
    If UBound(Split(StrTmp, "|")) = 3 Then
        '拆解字体信息“字体名称|字号|粗体|斜体”
        strPatientInfoFontName = Split(StrTmp, "|")(0)
        lngPatientInfoFontSize = Val(Split(StrTmp, "|")(1))
        blnPatientInfoFontBold = IIf(Split(StrTmp, "|")(2) = 1, True, False)
        blnPatientInfoFontItalic = IIf(Split(StrTmp, "|")(3) = 1, True, False)
    Else
        '向前兼容，病人信息字体字段原来直接保存的是字体大小
        lngPatientInfoFontSize = Val(StrTmp)
        blnPatientInfoFontBold = False
        blnPatientInfoFontItalic = False
        strPatientInfoFontName = "宋体"
    End If
    
    lngPatientInfoTitle = rsTemp("病人信息题头")
    '填充各种颜色
    lngLabelColor = rsTemp("标注正常颜色")                                                  ''标注显示色，白色
    lngLabelSelectedColor = rsTemp("标注选中颜色")                                          ''标注选中色，红色
    lngRulerLeftColor = rsTemp("标尺颜色")                                                  ''标尺颜色
    '标注设置
    lngWinWidthLevelLocation = rsTemp("窗宽窗位位置")
    lngLabelLineStyleNorm = rsTemp("标注正常线型")
    lngLabelLineWidthNorm = rsTemp("标注正常线宽")
    lngLabelFontSize = rsTemp("标注文字大小")
    '关联文字的显示设置
    bROIArea = IIf(rsTemp("测量显示面积") = -1, True, False)                                ''显示面积
    bROIMean = IIf(rsTemp("测量显示平均值") = -1, True, False)                              ''显示平均值
    bROIStandardDeviation = IIf(rsTemp("测量显示均方差") = -1, True, False)                 ''显示均方差
    bROILength = IIf(rsTemp("测量显示周长") = -1, True, False)                              ''显示周长
    bROIMax = IIf(rsTemp("测量显示最大值") = -1, True, False)                               ''显示最大值
    bROIMin = IIf(rsTemp("测量显示最小值") = -1, True, False)                               ''显示最小值
    bROITextChinese = IIf(rsTemp("测量显示中文") = -1, True, False)                         ''测量的结果是否使用中文
    intTextoOffX = rsTemp("文字X方向偏移")                                                  ''标注文字的偏移量
    intTextoOffY = rsTemp("文字Y方向偏移")                                                  ''标注文字的偏移量
    blnLabelTextScaleFontSize = IIf(rsTemp("文字随图像缩放") = -1, True, False)             ''标注文字大小是否随着
    '体位标记设置
    blnAnatomicMarkersLeft = IIf(Mid(rsTemp("显示体位标记"), 1, 1) = 1, True, False)        ''是否显示左边体位标记
    blnAnatomicMarkersRight = IIf(Mid(rsTemp("显示体位标记"), 3, 1) = 1, True, False)       ''是否显示右边体位标记
    blnAnatomicMarkersTop = IIf(Mid(rsTemp("显示体位标记"), 2, 1) = 1, True, False)         ''是否显示上边体位标记
    blnAnatomicMarkersBottom = IIf(Mid(rsTemp("显示体位标记"), 4, 1) = 1, True, False)      ''是否显示下边体位标记
    ''是否采用汉字显示体位标记
    blnChinaMark = IIf(rsTemp("中文体位标记") = -1, True, False)
    '标尺设置
    blnRulerDsipLeft = IIf(Mid(rsTemp("显示标尺"), 1, 1) = 1, True, False)                  ''是否显示左边标尺
    blnRulerDsipRight = IIf(Mid(rsTemp("显示标尺"), 3, 1) = 1, True, False)                 ''是否显示右边标尺
    blnRulerDsipTop = IIf(Mid(rsTemp("显示标尺"), 2, 1) = 1, True, False)                   ''是否显示上边标尺
    blnRulerDsipBottom = IIf(Mid(rsTemp("显示标尺"), 4, 1) = 1, True, False)                ''是否显示下边标尺
    intRulerLeft = rsTemp("标尺左右边距")                                                   ''标尺左边距
    intRulerTop = rsTemp("标尺上下边距")                                                    ''标尺上边距
    intRulerWidth = rsTemp("标尺宽度")                                                      ''标尺宽度
    intRulerHeight = rsTemp("标尺高度")                                                     ''标尺高度
    intRulerLineWidth = rsTemp("标尺线宽")                                                  ''标尺线宽
    '工具栏设置
    intToolBarIconSize = rsTemp("工具栏图标大小")
    intToolBarPosition = rsTemp("工具栏位置")
    blToolBarHide = IIf(rsTemp("工具栏显示") = -1, True, False)
    '血管狭窄测量
    intStandardThreshold = rsTemp("正常血管阈值")
    intNarrowThreshold = rsTemp("狭窄血管阈值")
    intVasEdgeWidth = rsTemp("血管壁宽度")
    
    '缩略图是否停靠于菜单下面
    blnDockMiniImage = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\frmSysConfig", "缩略图停靠于菜单下", False)
    blnShowMiniImageInfo = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\frmSysConfig", "缩略图中显示图像信息", True)
    blnSquareFrame = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\frmSysConfig", "框选报告图", True)
    blnShowMPRLine = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\frmSysConfig", "MPR显示辅助线", True)
    blnPrintFilmBeep = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\frmSysConfig", "胶片打印提示声音", False)
    
     '启用FTP文件大小对比
    gblnCompareSize = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", 1)) <> 0, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", IIf(gblnCompareSize, 1, 0))
End Sub


Public Sub LoadBarSetup(f As frmViewer)
'------------------------------------------------
'功能：读取数据库上次保存的工具栏设置
'参数：f--主窗体对象
'返回：无
'编制人：曾超
'------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    blfrmRefresh = False
       
    Select Case intToolBarIconSize
        Case 16
            BarterIco f.ImgList16
            CreateMenu f.ComToolBar, 16, 16
            f.ComToolBar.AddImageList f.ImgList16
        Case 24
            BarterIco f.ImgList24
            CreateMenu f.ComToolBar, 24, 24
            f.ComToolBar.AddImageList f.ImgList24
        Case 32
            BarterIco f.ImgList32
            CreateMenu f.ComToolBar, 32, 32
            f.ComToolBar.AddImageList f.ImgList32
    End Select
    
    f.ComToolBar.Item(ToolBar_Main).Position = intToolBarPosition
    
    ArrayToolBar f.ComToolBar, f.top, f.left, f.height, f.width

    For i = 2 To 8
        f.ComToolBar.Item(i).Visible = blToolBarHide
    Next
    f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_ToolBar_Hide, , True).Checked = Not blToolBarHide
    blfrmRefresh = True
End Sub

Public Function CreateUserWWWL(lngUserID As Long) As Boolean
'是否需要创建用户的窗宽窗位设置
'参数： lngUserID --- 用户ID
'返回值：True ---创建成功；False --- 创建失败，或不需要创建

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "select count(*) as Count from 影像预设窗宽窗位 where 人员id =[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否创建用户窗口设置", lngUserID)
    If rsTemp!Count = 0 Then
        strSQL = "Zl_影像预设窗宽窗位_Create(" & lngUserID & ")"
        zlDatabase.ExecuteProcedure strSQL, "创建用户窗口设置"
        CreateUserWWWL = True
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
End Function

Public Sub subSaveInterfaceParaIntoDB()
'------------------------------------------------
'功能：将当前的系统参数值，保存到“影像界面参数表”。
'参数：无
'返回：无
'------------------------------------------------
    Dim strAnatomicMarkers As String        '保存临时的体位标注
    Dim strRulerDsip As String              '保存临时的标尺标注
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    ''保存界面参数的时候，根据登录状态 blLocalRun 来判断，是保存到本机的MDB数据库，还是保存到联网的ORACLE数据库。
    
    '首先判断当前用户是否曾经保存过界面参数，如果是第一次保存界面参数，则首先插入一条用户自己的界面参数记录
    If blLocalRun = True Then
        strSQL = "select * from 影像界面参数表 "
        Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
        If rsTmp.EOF = True Then
            '--影像界面参数表
            strSQL = "insert into 影像界面参数表 (ID,正常图像边框颜色,正常图像边框线型,正常图像边框线宽,选中图像边框颜色,选中序列边框颜色,选中图像边框线型,选中图像边框线宽,图像标记颜色,图像标记大小,标注选择句柄颜色,标注选择句柄大小,定位线颜色,定位线线型,定位线间距,序列间间隔,横向最大序列,纵向最大序列,图像间距,显示多余边框,背景颜色,程序背景颜色,标注正常颜色,标注正常线型,标注正常线宽,标注选中颜色,标注选中线型,标注选中线宽,标注文字大小,测量显示面积,测量显示平均值,测量显示均方差,测量显示中文,文字X方向偏移,文字Y方向偏移,文字随图像缩放,显示体位标记,中文体位标记,显示标尺,标尺左右边距,标尺上下边距,标尺宽度,标尺高度,标尺线宽,标尺颜色,窗宽窗位位置,鼠标穿梭步长,鼠标漫游步长,鼠标调窗步长,鼠标缩放步长,病人信息上下边距,病人信息左右边距,病人信息颜色,病人信息显示最小值,病人信息随图像缩放,病人信息字体,病人信息题头,直接照相,工具栏图标大小,工具栏位置,工具栏显示,状态栏字体大小,正常血管阈值,狭窄血管阈值,血管壁宽度,测量显示周长)" & _
                     "VALUES (0,16777215,0,1,16777215,16777088,0,1,16777215,10,16777215,8,16777215,3,7,50,8,8,4,0,986895,131586,16777215,0,4,16777215,0,0,12,1,1,-1,-1,10,8,0,1010,1,1000,36,210,30,600,3,16777215,2,8,10,10,5,4,1,16777215,200,0,10,1,1,24,1,-1,9,51,50,10,-1);"
            cnAccess.Execute strSQL
            strSQL = "select * from 影像界面参数表 "
            Set rsTmp = cnAccess.Execute(strSQL, , adCmdText)
        End If
    Else
        strSQL = "select * from 影像界面参数表 where 人员id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取影像界面参数", glngUserID)
        If rsTmp.EOF = True Then
            strSQL = "select * from 影像界面参数表 where 人员id = 0 "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取影像界面参数")
            If rsTmp.EOF <> True Then
                strSQL = "ZL_影像界面参数表_INSERT(" & glngUserID
                For i = 1 To rsTmp.Fields.Count - 1
                    strSQL = strSQL & "," & rsTmp(i).Value
                Next
                strSQL = strSQL & ")"
                zlDatabase.ExecuteProcedure strSQL, "读取影像界面参数"
            End If
        End If
    End If
    
    '保存界面参数
    If blLocalRun = True Then
        strSQL = "update 影像界面参数表 set "
        ''选中图像边框颜色
        strSQL = strSQL & "正常图像边框颜色 = '" & lngSelectedImageBorderColor & "',"
        ''选中图像边框线形
        strSQL = strSQL & "正常图像边框线型 = '" & lngSelectedImageBorderLineStyle & "',"
        ''选中图像边框线宽度
        strSQL = strSQL & "正常图像边框线宽 = '" & lngSelectedImageBorderLineWidth & "',"
        ''选中图像边框颜色，就是当前颜色
        strSQL = strSQL & "选中图像边框颜色 = '" & lngCurrentImageBorderColor & "',"
        ''当前（未选中）序列边框颜色
        strSQL = strSQL & "选中序列边框颜色 = '" & lngCurrentSeriesBorderColor & "',"
        ''当前图像边框线形
        strSQL = strSQL & "选中图像边框线型 = '" & lngCurrentImageBorderLineStyle & "',"
        ''当前图像边框线宽度
        strSQL = strSQL & "选中图像边框线宽 = '" & lngCurrentImageBorderLineWidth & "',"
        ''选中图像标识填充色
        strSQL = strSQL & "图像标记颜色 = '" & lngSelectImageForeColour & "',"
        ''图像选择标记大小
        strSQL = strSQL & "图像标记大小 = '" & lngImageIdentifierSize & "',"
        ''选择句柄颜色
        strSQL = strSQL & "标注选择句柄颜色 = '" & lngPeriodColor & "',"
        ''选择句柄大小
        strSQL = strSQL & "标注选择句柄大小 = '" & intPeriodSize & "',"
        ''定位线颜色
        strSQL = strSQL & "定位线颜色 = '" & lngReferenceLineColor & "',"
        ''定位线线形
        strSQL = strSQL & "定位线线型 = '" & lngReferenceLineStyle & "',"
        ''定位线间距
        strSQL = strSQL & "定位线间距 = '" & lngReferenceLineSpacing & "',"
        ''序列之间的间隔宽度、高度
        strSQL = strSQL & "序列间间隔 = '" & intSpaceSize & "',"
        ''横向最多可划分的区域
        strSQL = strSQL & "横向最大序列 = '" & intMaxAreaX & "',"
        ''纵向最多可划分的区域
        strSQL = strSQL & "纵向最大序列 = '" & intMaxAreaY & "',"
        ''图像间距
        strSQL = strSQL & "图像间距 = '" & lngCellSpacing & "',"
        ''多余边框是否显示
        strSQL = strSQL & "显示多余边框 = '" & CInt(blnDsipSpilthBorder) & "',"
        ''Viewer背景颜色
        strSQL = strSQL & "背景颜色 = '" & lngViewerBackColor & "',"
        ''程序背景颜色
        strSQL = strSQL & "程序背景颜色 = '" & lngProgramBackColor & "',"
        ''标注显示色，白色
        strSQL = strSQL & "标注正常颜色 = '" & lngLabelColor & "',"
        ''标注正常线型
        strSQL = strSQL & "标注正常线型 = '" & lngLabelLineStyleNorm & "',"
        ''标注正常线宽
        strSQL = strSQL & "标注正常线宽 = '" & lngLabelLineWidthNorm & "',"
        ''标注选中色，红色
        strSQL = strSQL & "标注选中颜色 = '" & lngLabelSelectedColor & "',"
        ''标注文字大小
        strSQL = strSQL & "标注文字大小 = '" & lngLabelFontSize & "',"
        ''显示面积
        strSQL = strSQL & "测量显示面积 = '" & CInt(bROIArea) & "',"
        ''显示平均值
        strSQL = strSQL & "测量显示平均值 = '" & CInt(bROIMean) & "',"
        ''显示均方差
        strSQL = strSQL & "测量显示均方差 = '" & CInt(bROIStandardDeviation) & "',"
        ''测量结果信息是否使用中文
        strSQL = strSQL & "测量显示中文 = '" & CInt(bROITextChinese) & "',"
        ''标注文字的偏移量
        strSQL = strSQL & "文字X方向偏移 = '" & intTextoOffX & "',"
        ''标注文字的偏移量
        strSQL = strSQL & "文字Y方向偏移 = '" & intTextoOffY & "',"
        ''标注文字大小是否随着图像一起缩放
        strSQL = strSQL & "文字随图像缩放 = '" & CInt(blnLabelTextScaleFontSize) & "',"
        ''体位标注
        strAnatomicMarkers = IIf(blnAnatomicMarkersLeft, 1, 0) & IIf(blnAnatomicMarkersTop, 1, 0) _
                             & IIf(blnAnatomicMarkersRight, 1, 0) & IIf(blnAnatomicMarkersBottom, 1, 0)
        strSQL = strSQL & "显示体位标记 = '" & strAnatomicMarkers & "',"
        ''是否采用汉字显示体位标记
        strSQL = strSQL & "中文体位标记 = '" & CInt(blnChinaMark) & "',"
        ''显示标尺
        strRulerDsip = IIf(blnRulerDsipLeft, 1, 0) & IIf(blnRulerDsipTop, 1, 0) _
                       & IIf(blnRulerDsipRight, 1, 0) & IIf(blnRulerDsipBottom, 1, 0)
        strSQL = strSQL & "显示标尺 = '" & strRulerDsip & "',"
        ''标尺左边距
        strSQL = strSQL & "标尺左右边距 = '" & intRulerLeft & "',"
        ''标尺上边距
        strSQL = strSQL & "标尺上下边距 = '" & intRulerTop & "',"
        ''标尺宽度
        strSQL = strSQL & "标尺宽度 = '" & intRulerWidth & "',"
        ''标尺高度
        strSQL = strSQL & "标尺高度 = '" & intRulerHeight & "',"
        ''标尺线宽
        strSQL = strSQL & "标尺线宽 = '" & intRulerLineWidth & "',"
        ''标尺颜色
        strSQL = strSQL & "标尺颜色 = '" & lngRulerLeftColor & "',"
        ''窗宽窗位位置
        strSQL = strSQL & "窗宽窗位位置 = '" & lngWinWidthLevelLocation & "',"
        ''鼠标穿梭步长
        strSQL = strSQL & "鼠标穿梭步长 = '" & lngStackStep & "',"
        ''鼠标漫游步长
        strSQL = strSQL & "鼠标漫游步长 = '" & lngCruiseStep & "',"
        ''鼠标调窗步长
        strSQL = strSQL & "鼠标调窗步长 = '" & lngWidthLevelStep & "',"
        ''鼠标缩放步长
        strSQL = strSQL & "鼠标缩放步长 = '" & lngZoomStep & "',"
        ''鼠标滚轮操作
        strSQL = strSQL & "鼠标滚轮操作 = '" & intMouseWheelRoll & "',"
        ''病人信息上下边距
        strSQL = strSQL & "病人信息上下边距 = '0',"
        ''病人信息左右边距
        strSQL = strSQL & "病人信息左右边距 = '0',"
        ''病人信息颜色
        strSQL = strSQL & "病人信息颜色 = '" & lngpatientInfoColor & "',"
        ''病人信息显示最小值
        strSQL = strSQL & "病人信息显示最小值 = '" & lngPatientInfoInvisibleSize & "',"
        ''病人信息随图像缩放
        strSQL = strSQL & "病人信息随图像缩放 = '" & CInt(blnpatientInfoScaleFontSize) & "',"
        ''病人信息字体
        strSQL = strSQL & "病人信息字体 = '" & lngPatientInfoFontSize & "',"
        ''病人信息题头
        strSQL = strSQL & "病人信息题头 = '" & lngPatientInfoTitle & "',"
        ''是否直接照相，不显示胶片设置窗口
        strSQL = strSQL & "直接照相 = '" & CInt(bShowFilmConfig) & "',"
        strSQL = strSQL & "工具栏图标大小 = '" & intToolBarIconSize & "',"
        strSQL = strSQL & "工具栏位置 = '" & intToolBarPosition & "',"
        strSQL = strSQL & "工具栏显示 = '" & CInt(blToolBarHide) & "',"
        ''状态栏字体大小
        strSQL = strSQL & "状态栏字体大小 = '" & intStatusBarFontSize & "',"
        strSQL = strSQL & "正常血管阈值 = '" & intStandardThreshold & "',"
        strSQL = strSQL & "狭窄血管阈值 = '" & intNarrowThreshold & "',"
        strSQL = strSQL & "血管壁宽度 = '" & intVasEdgeWidth & "',"
        ''显示周长
        strSQL = strSQL & "测量显示周长 = '" & CInt(bROILength) & "' where id = 0"
        cnAccess.Execute strSQL, adCmdText
    Else
        strSQL = "ZL_影像界面参数表_UPDATE('" & glngUserID & "','"
        ''选中图像边框颜色
        strSQL = strSQL & lngSelectedImageBorderColor & "','"
        ''选中图像边框线形
        strSQL = strSQL & lngSelectedImageBorderLineStyle & "','"
        ''选中图像边框线宽度
        strSQL = strSQL & lngSelectedImageBorderLineWidth & "','"
        ''选中图像边框颜色，就是当前颜色
        strSQL = strSQL & lngCurrentImageBorderColor & "','"
        ''当前（未选中）序列边框颜色
        strSQL = strSQL & lngCurrentSeriesBorderColor & "','"
        ''当前图像边框线形
        strSQL = strSQL & lngCurrentImageBorderLineStyle & "','"
        ''当前图像边框线宽度
        strSQL = strSQL & lngCurrentImageBorderLineWidth & "','"
        ''选中图像标识填充色
        strSQL = strSQL & lngSelectImageForeColour & "','"
        ''图像选择标记大小
        strSQL = strSQL & lngImageIdentifierSize & "','"
        ''选择句柄颜色
        strSQL = strSQL & lngPeriodColor & "','"
        ''选择句柄大小
        strSQL = strSQL & intPeriodSize & "','"
        ''定位线颜色
        strSQL = strSQL & lngReferenceLineColor & "','"
        ''定位线线形
        strSQL = strSQL & lngReferenceLineStyle & "','"
        ''定位线间距
        strSQL = strSQL & lngReferenceLineSpacing & "','"
        ''序列之间的间隔宽度、高度
        strSQL = strSQL & intSpaceSize & "','"
        ''横向最多可划分的区域
        strSQL = strSQL & intMaxAreaX & "','"
        ''纵向最多可划分的区域
        strSQL = strSQL & intMaxAreaY & "','"
        ''图像间距
        strSQL = strSQL & lngCellSpacing & "','"
        ''多余边框是否显示
        strSQL = strSQL & CInt(blnDsipSpilthBorder) & "','"
        ''Viewer背景颜色
        strSQL = strSQL & lngViewerBackColor & "','"
        ''程序背景颜色
        strSQL = strSQL & lngProgramBackColor & "','"
        ''标注显示色，白色
        strSQL = strSQL & lngLabelColor & "','"
        ''标注正常线型
        strSQL = strSQL & lngLabelLineStyleNorm & "','"
        ''标注正常线宽
        strSQL = strSQL & lngLabelLineWidthNorm & "','"
        ''标注选中色，红色
        strSQL = strSQL & lngLabelSelectedColor & "','"
        ''标注文字大小
        strSQL = strSQL & lngLabelFontSize & "','"
        ''显示面积
        strSQL = strSQL & CInt(bROIArea) & "','"
        ''显示平均值
        strSQL = strSQL & CInt(bROIMean) & "','"
        ''显示均方差
        strSQL = strSQL & CInt(bROIStandardDeviation) & "','"
        ''测量结果信息是否使用中文
        strSQL = strSQL & CInt(bROITextChinese) & "','"
        ''标注文字的偏移量X
        strSQL = strSQL & intTextoOffX & "','"
        ''标注文字的偏移量Y
        strSQL = strSQL & intTextoOffY & "','"
        ''标注文字大小是否随着图像一起缩放
        strSQL = strSQL & CInt(blnLabelTextScaleFontSize) & "','"
        ''体位标注
        strAnatomicMarkers = IIf(blnAnatomicMarkersLeft, 1, 0) & IIf(blnAnatomicMarkersTop, 1, 0) _
                             & IIf(blnAnatomicMarkersRight, 1, 0) & IIf(blnAnatomicMarkersBottom, 1, 0)
        strSQL = strSQL & strAnatomicMarkers & "','"
        ''是否采用汉字显示体位标记
        strSQL = strSQL & CInt(blnChinaMark) & "','"
        ''显示标尺
        strRulerDsip = IIf(blnRulerDsipLeft, 1, 0) & IIf(blnRulerDsipTop, 1, 0) _
                       & IIf(blnRulerDsipRight, 1, 0) & IIf(blnRulerDsipBottom, 1, 0)
        strSQL = strSQL & strRulerDsip & "','"
        ''标尺左边距
        strSQL = strSQL & intRulerLeft & "','"
        ''标尺上边距
        strSQL = strSQL & intRulerTop & "','"
        ''标尺宽度
        strSQL = strSQL & intRulerWidth & "','"
        ''标尺高度
        strSQL = strSQL & intRulerHeight & "','"
        ''标尺线宽
        strSQL = strSQL & intRulerLineWidth & "','"
        ''标尺颜色
        strSQL = strSQL & lngRulerLeftColor & "','"
        ''窗宽窗位位置
        strSQL = strSQL & lngWinWidthLevelLocation & "','"
        ''鼠标穿梭步长
        strSQL = strSQL & lngStackStep & "','"
        ''鼠标漫游步长
        strSQL = strSQL & lngCruiseStep & "','"
        ''鼠标调窗步长
        strSQL = strSQL & lngWidthLevelStep & "','"
        ''鼠标缩放步长
        strSQL = strSQL & lngZoomStep & "','"
        '病人信息上下边距
        strSQL = strSQL & "0','"
        '病人信息左右边距
        strSQL = strSQL & "0','"
        '病人信息颜色
        strSQL = strSQL & lngpatientInfoColor & "','"
        '病人信息显示最小值
        strSQL = strSQL & lngPatientInfoInvisibleSize & "','"
        '病人信息随图像缩放
        strSQL = strSQL & CInt(blnpatientInfoScaleFontSize) & "','"
        '病人信息字体,字体信息组织方法“字体名称|字号|粗体|斜体”
        strSQL = strSQL & strPatientInfoFontName & "|" & lngPatientInfoFontSize & "|" & IIf(blnPatientInfoFontBold, 1, 0) & "|" & IIf(blnPatientInfoFontItalic, 1, 0) & "','"
        ''病人信息题头
        strSQL = strSQL & lngPatientInfoTitle & "','"
        ''是否直接照相，不显示胶片设置窗口
        strSQL = strSQL & CInt(bShowFilmConfig) & "','"
        ''工具栏图标大小
        strSQL = strSQL & intToolBarIconSize & "','"
        ''工具栏位置
        strSQL = strSQL & intToolBarPosition & "','"
        ''工具栏隐藏
        strSQL = strSQL & CInt(blToolBarHide) & "','"
        ''状态栏字体大小
        strSQL = strSQL & intStatusBarFontSize & "','"
        ''血管狭窄测量，正常血管阈值
        strSQL = strSQL & intStandardThreshold & "','"
        ''血管狭窄测量，狭窄血管阈值
        strSQL = strSQL & intNarrowThreshold & "','"
        ''血管狭窄测量，血管壁宽度
        strSQL = strSQL & intVasEdgeWidth & "','"
        ''显示周长
        strSQL = strSQL & CInt(bROILength) & "','"
        ''显示最大值
        strSQL = strSQL & CInt(bROIMax) & "','"
        ''显示最小值
        strSQL = strSQL & CInt(bROIMin) & "','"
        ''鼠标滚轮操作
        strSQL = strSQL & CInt(intMouseWheelRoll) & "','"
        ''是否显示胶片打印标记
        strSQL = strSQL & CInt(blnShowPrintTag) & "')"
        
        zlDatabase.ExecuteProcedure strSQL, "保存影像界面参数"
    End If
    Exit Sub
err:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subSaveParameters()
'------------------------------------------------
'功能：保存参数表的参数
'参数：无
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    Call zlDatabase.SetPara("鼠标滚轮拖动操作", intMouseWheelDrag, glngSys, 1289)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub subGetParameters()
'------------------------------------------------
'功能：读取参数表的参数
'参数：无
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    intMouseWheelDrag = Val(zlDatabase.GetPara("鼠标滚轮拖动操作", glngSys, 1289, 0))
    If intMouseWheelDrag < 0 Or intMouseWheelDrag > 2 Then intMouseWheelDrag = 0
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub
