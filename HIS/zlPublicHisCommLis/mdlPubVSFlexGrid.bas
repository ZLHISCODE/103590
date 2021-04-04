Attribute VB_Name = "mdlPubVSFlexGrid"
Option Explicit
Public Type VsfRowCol
    lngRow As Long
    lngCol As Long
End Type

Public Sub vfgSetting(ByVal LngStyle As Long, ByRef objVfg As VSFlexGrid, Optional ByVal strTtile As String, Optional VsfImg As ImageList, Optional ByVal lngFontSize As Long)
    'lngStyle＝0 默认设置，统一Vfg表格的外观
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'objVfg:    要初始化的控件
    'VsfImg:    ImageList图标集控件对象

    Dim arrHead As Variant, i As Long, strHead As String
    If strTtile = "" Then
        strHead = "第1列,900,1;第2列,900,1;第3列,900,1"
    Else
        strHead = strTtile
    End If
    arrHead = Split(strHead, ";")
    
    
    With objVfg
        '1.边框
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .GridLines = flexGridFlat
        .GridColorFixed = flexGridFlat
        
        '2.颜色
        .BackColor = vbWindowBackground '窗口背景
        .BackColorAlternate = vbWindowBackground
        .BackColorBkg = vbWindowBackground
        .BackColorFixed = vbButtonFace '按钮表面
        .BackColorFrozen = &H0&         '黑
        .FloodColor = &HC0&             '红
        .BackColorSel = &HFFEBD7        '浅绿
        .ForeColor = vbWindowText       '窗口文本
        .ForeColorFixed = vbButtonText  '按钮文本
        .ForeColorFrozen = &H0&         '黑
        .ForeColorSel = vbWindowText
        
        .GridColor = vbApplicationWorkspace '应用程序工作区
        .GridColorFixed = vbApplicationWorkspace
        .SheetBorder = vbWindowBackground
        .TreeColor = vbButtonShadow         '按钮阴影
        
        '3.初始化行列

        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            If CheckImgListKey(VsfImg, .TextMatrix(.FixedRows - 1, .FixedCols + i)) = True Then
                .Row = .FixedRows - 1
                .Col = .FixedCols + i
                .CellPicture = VsfImg.ListImages(Split(arrHead(i), ",")(0)).ExtractIcon
                '有图标时不显示文字
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = ""
            End If
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        
        '固定行文字居中
        If .FixedRows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .RowHeight(0) = 300
        .RowHeightMin = 300
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
        
        
        '4.其他属性
        .SelectionMode = flexSelectionByRow     '整行选择
        .ExplorerBar = flexExNone               '点标题栏不响应（排序及移动列）操作
        .AllowUserResizing = flexResizeColumns  '可调整列宽
        .Editable = flexEDNone                  '只读
        
         '5.字体
        If lngFontSize > 0 Then
            .FontSize = lngFontSize
        End If
        
    End With
    
End Sub

Public Function vfgLoadFromRecord(ByRef objVfg As VSFlexGrid, _
                                  ByRef rsTmp As ADODB.Recordset, _
                                  ByRef strErr As String, _
                                  Optional objImgList As ImageList) As Boolean
          '将记录集数据装入vfg控件
          'objVfg : vfg控件
          'rsTmp  : 装入控件的记录集
          'strErr :提示信息
          Dim i As Integer, strTitle As String
          
          '标题
1         On Error GoTo vfgLoadFromRecord_Error

2         For i = 0 To rsTmp.Fields.Count - 1
3             strTitle = strTitle & ";" & rsTmp.Fields(i).Name & ",0," & flexAlignLeftCenter
4         Next
5         If strTitle <> "" Then strTitle = Mid(strTitle, 2)
          
6         Call vfgSetting(0, objVfg, strTitle, objImgList)
          
          '处理数据
7         With objVfg
8             .Rows = .FixedRows + 1
9             .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
              'Set .DataSource = rsTmp 直接设数据源，则原来设置的格式标题等格式丢失，需手工添加数据
10            Do Until rsTmp.EOF
11                For i = 0 To rsTmp.Fields.Count - 1
12                    .TextMatrix(.Rows - 1, i) = CStr("" & rsTmp.Fields(i).value)
13                    If Not objImgList Is Nothing Then
14                        If CheckImgListKey(objImgList, rsTmp.Fields(i).Name) = True And CheckImgListKey(objImgList, rsTmp.Fields(i).value & "") = True Then
15                            .Row = .Rows - 1
16                            .Col = i
17                            .CellPicture = objImgList.ListImages(rsTmp.Fields(i).value).ExtractIcon
18                        End If
19                    End If
20                Next
21                .Rows = .Rows + 1
22                rsTmp.MoveNext
23            Loop
24            If .Rows > .FixedRows + 1 Then .Rows = .Rows - 1
25        End With
26        vfgLoadFromRecord = True
          
27        Exit Function
vfgLoadFromRecord_Error:
28        strErr = Err.Number & " " & Err.Description
29        Call WriteErrLog("zlPublicHisCommLis", "mdlPubVSFlexGrid", "执行(vfgLoadFromRecord)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
30        Err.Clear
End Function
Public Function CheckImgListKey(Vfgimg As ImageList, strKey As String) As Boolean
    '功能           检查关键字是否在图像列表中存在，如果存在返回为真
    '参数
    '               Vfgimg 传入的图像对象
    '               strKey 要检查当前传入的Key是否存在
    '返回           有返回真，没有返回假
    Dim intloop As Integer
    On Error Resume Next
    If Vfgimg Is Nothing Then Exit Function
    With Vfgimg
        For intloop = 1 To .ListImages.Count
            If .ListImages(intloop).Key = strKey Then
                CheckImgListKey = True
                Exit Function
            End If
        Next
    End With
End Function

Public Function vfgFindRowSel(ByRef objVfg As VSFlexGrid, strField As String, FindstrValue, Optional strErr As String) As Long
    '功能       查找指定字段和查找的值匹配，查找到并选择
    '参数
    '           objVfg      VSF对象
    '           strField    字段
    '           FindstrValue    查找的值
    Dim lngLoop As Long
    On Error Resume Next
    vfgFindRowSel = -1
    With objVfg
        For lngLoop = 1 To .Rows - 1
            If .TextMatrix(lngLoop, .ColIndex(strField)) = FindstrValue Then
                .Row = lngLoop
                vfgFindRowSel = lngLoop
                Exit Function
            End If
        Next
    End With

End Function
Public Function vfgFindRowSelA(ByRef objVfg As VSFlexGrid, strField As String, FindstrValue, Optional strErr As String) As Long
    '功能       查找指定字段和查找的值匹配，查找到并选择
    '参数
    '           objVfg      VSF对象
    '           strField    字段
    '           FindstrValue    查找的值
    Dim lngLoop As Long
    On Error Resume Next
    vfgFindRowSelA = -1
    With objVfg
        For lngLoop = 1 To .Rows - 1
            If .TextMatrix(lngLoop, .ColIndex(strField)) = FindstrValue Then
'                .Row = lngLoop
                vfgFindRowSelA = lngLoop
                Exit Function
            End If
        Next
    End With

End Function
Public Function vfgFindRowCheck(ByRef objVfg As VSFlexGrid, strField As String, FindstrValue As String, Optional lngRow As Long, Optional lngCol As Long) As Boolean
    '功能       检查是否有复复的值
    '参数
    '           objVfg      VSF对象
    '           strField    字段
    '           FindstrValue    查找的值
    '返回       查找有一样的值为真 否则为假
    Dim lngLoop As Long
    On Error Resume Next
    With objVfg
        For lngLoop = 1 To .Rows - 1
            If .TextMatrix(lngLoop, .ColIndex(strField)) = FindstrValue Then
                If lngLoop = lngRow And .ColIndex("strField") = lngCol Then
                Else
                    vfgFindRowCheck = True
                End If
                Exit Function
            End If
        Next
    End With
End Function
Public Function VsfColAllSelAllcls(objVSF As VSFlexGrid, intCol As Integer, Optional intSel As Integer, Optional strErr As String) As Boolean
          '功能               全选或全清选择框
          '参数               intSel 0=安批一行进行判断 1=全部选中 2=全部不选中

          Dim intRow As Integer
          
1         On Error GoTo VsfColAllSelAllcls_Error

2         With objVSF
3             If intSel = 0 Then
4                 If .Rows = 1 Then Exit Function
5                 intSel = .Cell(flexcpChecked, 1, intCol, 1, intCol)
6                 If intSel = 1 Then
7                     intSel = 2
8                 Else
9                     intSel = 1
10                End If
11            End If
12            For intRow = 1 To .Rows - 1
13                .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = intSel
14            Next
15        End With
16        VsfColAllSelAllcls = True

17        Exit Function
VsfColAllSelAllcls_Error:
18        strErr = "出错函数(vfgFindRowSel),出错信息:" & Err.Number & " " & Err.Description
19        Call WriteErrLog("zlPublicHisCommLis", "mdlPubVSFlexGrid", "执行(VsfColAllSelAllcls)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
20        Err.Clear
End Function

'其他未分类的控件公共函数

Public Sub SplitWE(LeftControls As Collection, _
                   SplitControl As Control, _
                   RightControls As Collection, _
                   X As Single, Optional minWidth As Single)
    '左右分隔MouseDown事件处理函数
    'LeftControls:左边控件集
    'RightControls:右边控件集
    'splitcontrol :分隔条控件
    'X            :偏移量，从Splitcontrol的MouseMove事件中传入
    'minWidth     :最小宽度
    Dim objControl As Control, sinWidth As Single
    On Error GoTo errH
    If LeftControls.Count < 0 Or RightControls.Count < 0 Then Exit Sub
    sinWidth = 3000
    If minWidth > 0 Then sinWidth = minWidth
    If LeftControls.Item(1).Width + X < sinWidth Or RightControls.Item(1).Width - X < sinWidth Then Exit Sub
    
    For Each objControl In LeftControls
        objControl.Width = objControl.Width + X
    Next
    
    SplitControl.Left = SplitControl.Left + X
            
    For Each objControl In RightControls
        objControl.Left = SplitControl.Left + SplitControl.Width
        objControl.Width = objControl.Width - X
    Next
    Exit Sub
errH:
    Exit Sub
End Sub

Public Sub SplitNS(TopControls As Collection, _
                   SplitControl As Control, _
                   ButtonControls As Collection, _
                   Y As Single, _
                   Optional minHight As Single)
    '上下分隔MouseDown事件处理函数
    'TopControls:上边控件集
    'ButtonControls:下边控件集
    'splitcontrol :分隔条控件
    'Y            :偏移量，从Splitcontrol的MouseMove事件中传入
    Dim objControl As Control, sigHight As Single
    On Error GoTo errH
    If TopControls.Count < 0 Or ButtonControls.Count < 0 Then Exit Sub
    sigHight = 3000
    If minHight > 0 Then sigHight = minHight
    If TopControls.Item(1).Height + Y < sigHight Or ButtonControls.Item(1).Height - Y < sigHight Then Exit Sub
    
    For Each objControl In TopControls
        objControl.Height = objControl.Height + Y
    Next
    
    SplitControl.Top = SplitControl.Top + Y
            
    For Each objControl In ButtonControls
        objControl.Top = SplitControl.Top + SplitControl.Height
        objControl.Height = objControl.Height - Y
    Next
    Exit Sub
errH:
    Exit Sub
End Sub

