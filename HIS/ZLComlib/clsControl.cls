VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsComboBox"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"clsGrid"
Option Explicit
'---------------------------------------------------------------------
'模块名称：clsControl
'包含对象：操作控件的函数
'命名规则：控件简名+操作类型 + 操作属性
'说明    ：控件简名按编码规范中的书写；操作类型一般为动词；操作属性为控件属性，对于涉及多个或不好确定的，可以用便于理解的词汇替代。
'例      ：LvwSortColumn，LvwFlatColumnHeader，FrmCheckInput
'--------------------------------------------------------------------------
'改变背景色，TvwSetBackColor
Private Const TVM_SETBKCOLOR = 4381&
Private Const TVM_GETBKCOLOR = 4383&
Private Const TVS_HASLINES = 2&
'控件坐标位置获取转换
Private Const EM_EXGETSEL = (&H400 + 52)
Private Const EM_POSFROMCHAR = &HD6
'结构声明
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Enum mTextAlign
    taLeftAlign = 0
    taCenterAlign = 1
    taRightAlign = 2
End Enum

Public Enum mTextType
    m文本式 = 0
    m数字式 = 1
    m金额式 = 2
    m负金额式 = 3
End Enum

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private mvarclsComboBox As clsComboBox

Public Property Get Cbo() As clsComboBox
    If mvarclsComboBox Is Nothing Then
        Set mvarclsComboBox = New clsComboBox
    End If
    Set Cbo = mvarclsComboBox
End Property

Public Sub CboSetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'功能：设置Combo控件的Index值
'为一个Combo控件选择列表项，但又不触发其Click事件
    Const CB_SETCURSEL = &H14E
    
    SendMessage hWnd_combo, CB_SETCURSEL, lngIndex, 0
End Sub

Public Sub CboSetWidth(ByVal hWnd_combo As Long, ByVal lngWidth As Long)
'功能：设置Combo控件下拉列表的宽度
'此处的宽度是批下拉列表的宽度，并且是以TWIP为单位
    Const CB_SETDROPPEDWIDTH As Long = &H160

    SendMessage hWnd_combo, CB_SETDROPPEDWIDTH, lngWidth / Screen.TwipsPerPixelX, 0
End Sub

Public Sub CboSetHeight(cboControl As Object, ByVal lngHeight As Long)
'功能：设置Combo控件的下拉列表的高度
'此处的宽度是批下拉列表的高度，并且是以TWIP为单位
    SetWindowPos cboControl.hWnd, 0, 0, 0, cboControl.Width / Screen.TwipsPerPixelX, lngHeight / Screen.TwipsPerPixelY, SWP_NOMOVE
End Sub

Public Sub CboSetText(cboControl As Object, ByVal strText As String, Optional ByVal blnAfter As Boolean = True, Optional strSplit As String = "-")
'功能：根据文本串更新Combo控件的当前值
'参数：cboControl  准备设置的ComboBox控件
'      strText     输入的文本串
'      blnAfter    表示在分隔符之前或之后取值。如果没有分隔符，则取之后
'      strSplit    分隔符，通常为-
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cboControl.ListCount - 1
        strTemp = cboControl.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            '直接返回整个字符串
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                '圆点之前
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '已经找到
        cboControl.ListIndex = lngCount
    Else
        If blnAfter = True Then
            '这才是实际内容，如果为前则只是编码
            If strText <> "" Then
                cboControl.AddItem strText
                cboControl.ListIndex = cboControl.NewIndex
            End If
        End If
    End If
End Sub

Public Sub CboAppendText(cboControl As Object, KeyAscii As Integer)
'功能：对ComboBox实现输入过程中自动完成的功能
'说明：在Combox.KeyPress事件中调用
    Dim strInput As String
    Dim lngStart As Long
    Dim lngIndex As Long
    Const CB_FINDSTRING = &H14C
    
    If cboControl.Style <> 0 Then Exit Sub
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    strInput = Chr(KeyAscii): KeyAscii = 0

    With cboControl
        '记录上次的插入点位置
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)

        '接着得到用户击键完成后文本框中出现的内容
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)

        '根据假想的内容得到可能的列表项
        lngIndex = SendMessage(cboControl.hWnd, CB_FINDSTRING, -1, ByVal strInput)
        If lngIndex >= 0 Then
            '.ListIndex = lngIndex
            .Text = .List(lngIndex)
            .SelStart = Len(strInput)
            .SelLength = 100
        Else
            .Text = strInput
            .SelStart = lngStart
        End If
    End With
End Sub

Public Function CboLocate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
'建议弃用，使用Cbo.SeekIndex代替
'blnItem:True-表示根据ItemData的值定位下拉框;False-表示根据文本的内容定位下拉框
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If blnItem Then
            If cboObj.ItemData(lngLocate) = Val(strValue) Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboObj.List(lngLocate), InStr(1, cboObj.List(lngLocate), "-") + 1) = strValue Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Public Sub CboAddData(objSource As Object, ByVal rsDataSource As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
    '-----------------------------------------------------------------------------------
    '功能:装载数据入指定的组合下拉框或网格中的下拉框中
    '参数:objSource   要装载记录集的下拉框控件
    '     rsDataSource     记录集数据,只能有两个数据项,第一个是下拉数据,第二个是对应的关键字,如id
    '     blnClear    装载时是否清楚原有的下拉数据,缺省为True
    '返回:
    '-----------------------------------------------------------------------------------
    If blnClear = True Then objSource.Clear
    
    If rsDataSource.BOF = False Then
        rsDataSource.MoveFirst
        While Not rsDataSource.EOF
            objSource.AddItem rsDataSource.Fields(0).value
            objSource.ItemData(objSource.NewIndex) = Val(rsDataSource.Fields(1).value)
            rsDataSource.MoveNext
        Wend
        rsDataSource.MoveFirst
    End If
End Sub

Public Function CboMatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：匹配的索引(含不匹配的索引)
'说明：只返回索引,但未设置索引。请将该函数在KeyPress事件中调用。
    Static lngPreTime As Single, lngPreHwnd As Long, strFind As String
    Dim sngTime As Single, lngR As Long
    Dim lngCount As Long, strText As String * 255
    Dim lngIdx As Long, strTmp As String, i As Long
    
    If lngHwnd <> 0 Then
        If KeyAscii = 13 Then '回车不处理,直接返回当前索引
            CboMatchIndex = SendMessage(lngHwnd, CB_GETCURSEL, 0, 0)
            Exit Function
        Else
            '弹出列表内容
            If SendMessage(lngHwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 Then
                Call gobjComLib.zlCommFun.PressKey(vbKeyF4)
            End If
        End If
    End If
    
    If lngPreHwnd <> lngHwnd Or lngHwnd = 0 Then
        lngPreTime = Empty: strFind = Empty
    End If
    lngPreHwnd = lngHwnd
    
    If lngHwnd = 0 Then
        CboMatchIndex = -1: Exit Function
    End If
    
    sngTime = timer
    If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
        strFind = ""
    End If
    strFind = strFind & Chr(KeyAscii)
    lngPreTime = timer
    KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        
    '用API从开头进行查找
    lngIdx = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
    
    '无匹配再从中间按简码匹配
    If lngIdx = -1 Then
        lngCount = SendMessage(lngHwnd, CB_GETCOUNT, 0, 0)
        For i = 0 To lngCount - 1
            strText = String(SendMessage(lngHwnd, CB_GETLBTEXTLEN, i, 0), Chr(32)) & Chr(0)
            Call SendMessage(lngHwnd, CB_GETLBTEXT, i, ByVal strText)
            
            If InStr(strText, Chr(0)) > 0 Then
                strTmp = Mid(strText, 1, InStr(strText, Chr(0)) - 1)
            Else
                strTmp = strText
            End If
            strTmp = gobjComLib.zlCommFun.GetNeedName(strTmp)
            If gobjComLib.zlCommFun.SpellCode(strTmp) Like UCase(strFind) & "*" Then
                lngIdx = i: Exit For
            End If
        Next
    End If
    
    If lngIdx = -1 Then Beep
    CboMatchIndex = lngIdx
End Function

Public Function CbsInit(ByRef cbsMain As Object, Optional ByVal blnEnableCustomization As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：初始化CommandBar控件，使菜单风格一致
    '参数：
    '返回：
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = blnEnableCustomization
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization blnEnableCustomization
    Set cbsMain.Icons = gobjComLib.zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = True
    CbsInit = True
End Function

Public Function FormCheckInput(objForm As Object) As Boolean
'功能：检查指定窗体中是否有非法字符输入及输入长度是否超长
'参数：objForm=要检查的窗体
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled And Not obj.Locked Then
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                Case "ComboBox"
                    If obj.Style = 0 Then strText = obj.Text
                End Select
                If InStr(strText, "'") > 0 _
                    Or InStr(strText, ",") > 0 _
                    Or InStr(strText, ";") > 0 _
                    Or InStr(strText, "|") > 0 _
                    Or InStr(strText, "~") > 0 _
                    Or InStr(strText, "^") > 0 Then
                    MsgBox "输入数据中包含非法字符！", vbInformation, gstrSysName
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    FormCheckInput = True
    If Err.Number <> 0 Then Err.Clear
End Function

Public Sub FormSetAlpha(ByVal lngfrmHwnd As Long, ByVal lngAlpha As Long)
'功能：设置窗口透明度,该方法只适用于窗口对象
'参数：lngfrmHwnd:窗体句柄
'          lngAlpha:透明参数取值，0-255
    Dim tmpLog As Long
    
    If lngfrmHwnd = 0 Then Exit Sub
    If lngAlpha >= 0 And lngAlpha <= 255 Then
        tmpLog = GetWindowLong(lngfrmHwnd, GWL_EXSTYLE) '窗口属性
        Call SetWindowLong(lngfrmHwnd, GWL_EXSTYLE, tmpLog Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(lngfrmHwnd, 0, lngAlpha, LWA_ALPHA)
    End If
End Sub

Public Sub FormLock(Optional ByVal lngHwnd As Long)
'功能：锁定对象区域不刷新。或者解除锁定
'参数：lngHwnd=0-解除锁定,<>0要锁定对象的Hwnd
    LockWindowUpdate (lngHwnd)
End Sub

Public Function LvwSelectColumns(objSet As Object, ByVal strColumn As String, Optional ByVal blnInit As Boolean = False) As Boolean
'功能:对列表控件的列进行设置
'参数:
'   objSet：要设置的对象,目前只支持ListView，以后再加上FlexGrid,DataGrid。
'   strColumn；列串。格式是"列名,列宽,对齐数值,列特性;列名,列宽,对齐数值,列特性"    注意列之间是用分号
'      比如 "名称,2000,0,1;编码,800,0,0;简码,1440,0,0"
'      对ListView而言：列特性为1表示该列不可删除，列特性为0表示该列可以删除
'      对FlexGridView而言：列特性还要表示是否属于固定列，以便不能和其它列进行顺序调整
'   blnInit：True,不显示选择窗口，直接初始化
    Dim varColumns As Variant, varColumn As Variant
    Dim lngCol As Long

    If blnInit Then
        varColumns = Split(strColumn, ";")
        Select Case TypeName(objSet)
            Case "ListView"
                With objSet.ColumnHeaders
                    .Clear
                    For lngCol = LBound(varColumns) To UBound(varColumns)
                        varColumn = Split(varColumns(lngCol), ",")
                        .Add , "_" & varColumn(0), varColumn(0), varColumn(1), varColumn(2)
                    Next
                End With
            Case "MSHFlexGrid"
            Case "DataGrid"
        End Select
    Else
        LvwSelectColumns = frmSelectColumn.设置列(objSet, strColumn)
    End If
End Function

Public Sub LvwFlatColumnHeader(ByVal lvw As Object)
'功能：使用ListView的列标题成为平面
    Const strHeaderClass As String = "msvb_lib_header"
    Const HDS_BUTTONS   As Long = 2
    
    Dim lngChild As Long, lngLen As Long, lngStyle As Long
    Dim strName As String * 255

    
    lngChild = GetWindow(lvw.hWnd, GW_CHILD)
    Do While lngChild <> 0
        lngLen = GetClassName(lngChild, strName, 255)
        If lngLen > 0 Then
            If Mid(strName, 1, lngLen) = strHeaderClass Then
                lngStyle = GetWindowLong(lngChild, GWL_STYLE)
                lngStyle = lngStyle And (Not HDS_BUTTONS)
                SetWindowLong lngChild, GWL_STYLE, lngStyle
                Exit Sub
            End If
        End If
        lngChild = GetWindow(lngChild, GW_HWNDNEXT)
    Loop

End Sub

Public Sub LvwSortColumn(lvw As Object, lngColumn As Long)
'功能：根据ListView列头的点击进行排序
'参数：lvw            ListView控件
'      lngColumn      当前点击的列头序号，即ColumnHeader.Index值
    If lvw.Sorted = False Then lvw.Sorted = True

    If lvw.SortKey = lngColumn - 1 Then '仍是刚才那列
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvw.SortKey = lngColumn - 1  '排序关键字是从0开始的
        lvw.SortOrder = lvwAscending
    End If
    
    If Not lvw.SelectedItem Is Nothing Then
        lvw.SelectedItem.EnsureVisible
    End If
End Sub

Public Function LvwSaveItem(lvwObj As Object) As String
    '-----------------------------------------------------------------------------------
    '功能:保存ListView中的选中项
    '参数:lvwobj      要保存的ListView控件
    '返回:选中项的关键字
    '-----------------------------------------------------------------------------------
    If Not (lvwObj.SelectedItem Is Nothing) Then LvwSaveItem = lvwObj.SelectedItem.Key
End Function

Public Sub LvwRestoreItem(lvwObj As Object, svrKey As String)
    '-----------------------------------------------------------------------------------
    '功能:恢复ListView中的选中项
    '参数:lvwobj      要恢复的ListView控件
    '     svrKey      要恢复的ListView项的关键字
    '返回:
    '-----------------------------------------------------------------------------------
    On Error GoTo EndP
    If lvwObj.ListItems.count > 0 Then
        If Not (lvwObj.ListItems(svrKey) Is Nothing) Then
            lvwObj.ListItems(svrKey).Selected = True
            lvwObj.ListItems(svrKey).EnsureVisible
        End If
    End If
    Exit Sub
EndP:
    Err.Clear
    If lvwObj.ListItems.count > 0 Then lvwObj.ListItems(1).Selected = True
End Sub

Public Sub LvwSetColWidth(objLvw As Object, Optional blnHideNullCol As Boolean)
'功能：根据ListView中当前的内容自动调整列为最小匹配宽度,并保持至少可以显示列头文字的宽度
'参数：objLvw=要调整的ListView对象
'      blnHideNullCol=是否隐藏没有任何数据的列
    Dim i As Integer, lngW As Long
    For i = 1 To objLvw.ColumnHeaders.count
        SendMessage objLvw.hWnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If blnHideNullCol Then If objLvw.ColumnHeaders(i).Width < 200 Then objLvw.ColumnHeaders(i).Width = 0
        If objLvw.ColumnHeaders(i).Width < (gobjComLib.zlCommFun.ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * 90 And objLvw.ColumnHeaders(i).Width <> 0 Then
            objLvw.ColumnHeaders(i).Width = (gobjComLib.zlCommFun.ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * 90
        End If
    Next
End Sub

Public Sub TvwSetBackColor(ByVal objtvw As Object, ByVal lngColor As Long)
'功能：改变TreeView控件的背景色
   Dim lngStyle As Long
   Dim nodTemp As Node
   
    Call SendMessageLong(objtvw.hWnd, TVM_SETBKCOLOR, 0, lngColor)   'Change the background color
    ' Now reset the style so that the tree lines appear properly
    lngStyle = GetWindowLong(objtvw.hWnd, GWL_STYLE)
    Call SetWindowLong(objtvw.hWnd, GWL_STYLE, lngStyle - TVS_HASLINES)
    Call SetWindowLong(objtvw.hWnd, GWL_STYLE, lngStyle)
    For Each nodTemp In objtvw.Nodes
        nodTemp.BackColor = lngColor
    Next
    objtvw.Refresh
End Sub

Public Sub MshSetColWidth(objMsh As Object, frmParent As Object)
'功能：自动调整MSHFlexGrid表格的各列宽度,以最小匹配为准,并保持列头文字可以全部显示
'参数：msh=要调整列宽的MSHFlexGrid控件
'      frmParent=父窗体,用于计算文字宽度
    Dim arrWidth() As Long, blnRedraw As Boolean
    Dim lngRows As Long
    Dim i As Long, j As Long
    
    ReDim arrWidth(objMsh.Cols - 1)
    
    blnRedraw = objMsh.Redraw
    lngRows = objMsh.Rows - 1
    If lngRows > 200 Then lngRows = 200
    objMsh.Redraw = False
    For i = 0 To objMsh.Cols - 1
        For j = IIf(objMsh.FixedRows = 0, 0, objMsh.FixedRows - 1) To lngRows
            If Not objMsh.MergeRow(j) Then
                If frmParent.TextWidth(objMsh.TextMatrix(j, i)) + 90 > arrWidth(i) Then
                    arrWidth(i) = frmParent.TextWidth(objMsh.TextMatrix(j, i)) + 90
                End If
            End If
        Next
    Next
    For i = 0 To objMsh.Cols - 1
        objMsh.ColWidth(i) = IIf(arrWidth(i) > 3000, 3000, arrWidth(i))
    Next
    objMsh.Redraw = blnRedraw
End Sub

Public Sub MshColumnSort(ByVal mshGrd As Object, ByVal IntSorCol As Integer, _
    ByRef intPreCol As Integer, ByRef intPreSort As Integer, _
    Optional blnNum As Boolean = False, Optional intMachingCol As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对列进行排序
    '参数:
    '       mshGrd-MSHFlexGrid网络控件
    '       intPreCol-前一次单据头的排序列
    '       intPreSort-前一次单据头的排序
    '       blnNum-是否为数量列
    '       intMachingCol-匹配列,主要用于排序后重新定位到指定的列
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    Const flexSortNumericAscending  As Long = 3
    Const flexSortNumericDescending  As Long = 4
    Const flexSortStringNoCaseAscending As Long = 5
    Const flexSortStringNoCaseDescending As Long = 6
    
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strTmp As String
    
    Err = 0
    On Error GoTo Errhand:
    With mshGrd
        If IntSorCol > .Cols - 1 Or IntSorCol < 0 Then Exit Sub
        If .Rows > 1 Then
            .Redraw = False
            intCol = IntSorCol   '.MouseCol
            .Col = intCol
            .ColSel = intCol
            If intMachingCol > .Cols - 1 Or intMachingCol < 0 Then
                strTmp = .TextMatrix(.Row, 0)
            Else
                strTmp = .TextMatrix(.Row, intMachingCol)
            End If
            If blnNum Then
                    If intCol = intPreCol And intPreSort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       intPreSort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       intPreSort = flexSortNumericDescending
                    End If
            Else
                    If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       intPreSort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       intPreSort = flexSortStringNoCaseDescending
                    End If
            End If
            
            intPreCol = intCol
            If intMachingCol > .Cols - 1 Or intMachingCol < 0 Then
                .Row = MshGrdFindRow(mshGrd, strTmp, 0)
            Else
                .Row = MshGrdFindRow(mshGrd, strTmp, intMachingCol)
            End If
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
    Exit Sub
Errhand:
    Err.Clear
    mshGrd.Redraw = True
End Sub

Public Function MshGrdFindRow(ByVal mshGrd As Object, _
        ByVal strTmp As Variant, ByVal intCol As Integer) As Integer
    '----------------------------------------------------------------------------
    '功能:寻找与某一单元值相等的行
    '参数:
    '       mshGrd-网格控件
    '       strTmp-匹配串
    '       intCol-匹配列
    '返回:查找到的行,如未找到,返位第一行
    '----------------------------------------------------------------------------
    
    Dim i As Integer
    Err = 0
    On Error GoTo Errhand:
    
    With mshGrd
        For i = 1 To .Rows - 1
            If IsDate(strTmp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(strTmp, "yyyy-mm-dd") Then
                  MshGrdFindRow = i
                  Exit Function
               End If
            ElseIf IsNumeric(strTmp) Then
                If Val(.TextMatrix(i, intCol)) = Val(strTmp) Then
                  MshGrdFindRow = i
                  Exit Function
                End If
            Else
                If .TextMatrix(i, intCol) = strTmp Then
                  MshGrdFindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    MshGrdFindRow = 1
    Exit Function
Errhand:
    Err.Clear
    MshGrdFindRow = 1
    Exit Function
End Function

Public Sub MshClearData(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '-----------------------------------------------------------------------------------
    '功能: 清除指定网格的指定行指定列的数据
    '参数: obj=要操作的网格控件
    '      intRow=要清除的行号
    '      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    '返回：
    '-----------------------------------------------------------------------------------
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Public Sub MshSetFormat(ByVal msfObj As Object, ByVal strColumns As String, ByVal strCaption As String, _
Optional ByVal rsData As ADODB.Recordset, Optional ByVal blnData As Boolean = False, Optional ByVal blnRestore As Boolean = False)
    Dim intCol As Integer, intCols As Integer
    '功能：设置表格控件的列头，或绑定数据到表格控件
    '各参数含义：msfObj是表格对象；strColumns是列名格式串（列名|宽度|列对齐方式）；blnData说明是设置格式还是绑定数据；blnRestore说明是否恢复原始宽度
    '列名格式：NO|1000|4,日期|1600|4,检查单位|2000|4,检查人|800|4,简要记录|2000|4,结论|1000|4
    Dim arrTmp As Variant
    
    '设置表格控件格式
    intCols = UBound(Split(strColumns, ",")) + 1
    arrTmp = Split(strColumns, ",")
    
    With msfObj
        .Redraw = False
        
        '设置单据头，单据体的格式
        If Not blnData Then
            .Clear
            .Rows = 2
            .Cols = intCols
            
            For intCol = 0 To .Cols - 1
                .TextMatrix(0, intCol) = Split(arrTmp(intCol), "|")(0)
                .ColWidth(intCol) = Split(arrTmp(intCol), "|")(1)
                .ColAlignment(intCol) = Split(arrTmp(intCol), "|")(2)
                
                If TypeName(msfObj) <> "BillEdit" Then .ColAlignmentFixed(intCol) = 4
            Next
            If blnRestore Then Call gobjComLib.RestoreFlexState(msfObj, strCaption)
        Else
            If Not rsData.EOF Then
                Set .DataSource = rsData
                For intCol = 0 To .Cols - 1
                    .ColAlignment(intCol) = Split(arrTmp(intCol), "|")(2)
                Next
            End If
        End If
        
        
        .Row = 1
        .Col = 0
        If TypeName(msfObj) <> "BillEdit" Then
            .RowSel = 1
            .ColSel = .Cols - 1
        End If
        
        .Redraw = True
    End With
End Sub

Public Sub VSFSetFontSize(ByRef vsf As Object, ByVal lngFontSize As Long, Optional ByVal lngCol As Long)
'功能：设置vsfflexgrid字体的大小，并自动调整列宽和行高
'参数：lngFontSize常用设置：9号即小五，12号即小四
'      lngCol,如果要设置自动调整行高，则传入行高依据的列，要求AutoSizeMode = flexAutoSizeRowHeight,WordWrap =True
    Dim i As Long, lngRate As Double, lngTmp As Long
    
    If lngFontSize < 5 Or lngFontSize > 50 Then Exit Sub
    With vsf
        lngRate = lngFontSize / .FontSize
        lngTmp = .Redraw
        
        .Redraw = flexRDNone
        .FontSize = lngFontSize
        
        For i = 0 To .Cols - 1
           If .ColWidth(i) > 0 Then
             .ColWidth(i) = .ColWidth(i) * lngRate
           End If
        Next
        
        If .AutoSizeMode = flexAutoSizeRowHeight And .WordWrap And lngCol > 0 Then
            .AutoSize lngCol
        Else
            For i = 0 To .Rows - 1
                .RowHeight(i) = .RowHeight(i) * lngRate
            Next
        End If
        .Redraw = lngTmp
    End With
End Sub

Public Function OneCharWidth(ByVal ftCur As StdFont) As Long
'功能：获取特定字体下的一个半角英文字符的宽度
'参数：ftCur=字体环境
'返回：单个字符宽度
    Dim objTmp As frmPubTmp
    
    '非加粗的宋体，9号与12号字体，系统内置字体宽度，主要因为这两个是常用字体
    If ftCur.Name = "宋体" And (ftCur.Size = 9 Or ftCur.Size = 12) And Not ftCur.Bold Then
        If ftCur.Size = 9 Then
            OneCharWidth = 90
        Else
            OneCharWidth = 120
        End If
    Else '其他字体宽度通过特殊方法获取
        Set objTmp = New frmPubTmp
        Set objTmp.picDraw.Font = ftCur '设置临时界面字体
        OneCharWidth = objTmp.picDraw.TextWidth("A") '获取一个字符宽
    End If
End Function

Public Function OneCharHeight(ByVal ftCur As StdFont) As Long
'功能：获取特定字体下的一个半角英文字符的高度
'参数：ftCur=字体环境
'返回：单个字符高度
    Dim objTmp As frmPubTmp
    
    '非加粗的宋体，9号与12号字体，系统内置字体宽度，主要因为这两个是常用字体
    If ftCur.Name = "宋体" And (ftCur.Size = 9 Or ftCur.Size = 12) And Not ftCur.Bold Then
        If ftCur.Size = 9 Then
            OneCharHeight = 180
        Else
            OneCharHeight = 240
        End If
    Else '其他字体宽度通过特殊方法获取
        Set objTmp = New frmPubTmp
        Set objTmp.picDraw.Font = ftCur '设置临时界面字体
        OneCharHeight = frmPubTmp.picDraw.TextHeight("A")  '获取一个字符宽
    End If
End Function

Public Sub PicShowFlat(objPic As Object, Optional IntStyle As Integer = -1, Optional strName As String = "", Optional intAlign As mTextAlign, Optional blnFontBold As Boolean)
'功能：将PictureBox模拟成凹下或凸起的面板
'参数：'intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
'      intAlign=如果要显示文本,则指定对齐方式
    
    Dim vRect As RECT, lngTmp As Long
    
    With objPic
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If IntStyle <> 0 Then
            vRect.Left = .ScaleLeft
            vRect.Top = .ScaleTop
            vRect.Right = .ScaleWidth
            vRect.Bottom = .ScaleHeight
            Select Case IntStyle
                Case 1
                    DrawEdge .hDC, vRect, CLng(BDR_RAISEDINNER Or BF_SOFT), BF_RECT
                Case 2
                    DrawEdge .hDC, vRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hDC, vRect, CLng(BDR_SUNKENOUTER Or BF_SOFT), BF_RECT
                Case -2
                    DrawEdge .hDC, vRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If intAlign = taCenterAlign Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2 '中间对齐
            ElseIf intAlign = taRightAlign Then
                .CurrentX = .ScaleWidth - .TextWidth(strName) - 2 '右边对齐
            Else
                .CurrentX = 2 '左边对齐
            End If
            .FontBold = blnFontBold
            objPic.Print strName
        End If
    End With
End Sub

Public Sub PicDrawPicture(picBox As Object, objPic As StdPicture, ByVal W As Long, ByVal H As Long)
    '******************************************************************************************************************
    '功能：在PictureBox中央按适当比例画一幅图
    '参数：W,H=要作图的尺寸
    '******************************************************************************************************************
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    If W <= picBox.ScaleWidth And H <= picBox.ScaleHeight Then
        lngW = W: lngH = H
    Else
        sngW = W / picBox.ScaleWidth
        sngH = H / picBox.ScaleHeight
        If sngW > sngH Then
            lngW = W / sngW: lngH = H / sngW
        Else
            lngW = W / sngH: lngH = H / sngH
        End If
    End If
    
    picBox.Cls
    On Error Resume Next
    picBox.PaintPicture objPic, (picBox.ScaleWidth - lngW) / 2, (picBox.ScaleHeight - lngH) / 2, lngW, lngH
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub PicDrawLine(picBox As Object, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1)
    '功能：在(X1,Y1),(X2,Y2)之间使用ForeColor色画一直线
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    
    lngSaveForeColor = picBox.ForeColor
    bytSaveLineWidth = picBox.DrawWidth
    picBox.ForeColor = ForeColor
    picBox.DrawStyle = DrawStyle
    picBox.DrawWidth = LineWidth
    picBox.Line (X2, Y2)-(X1, Y1)
    picBox.ForeColor = lngSaveForeColor
    picBox.DrawWidth = bytSaveLineWidth
End Sub

Public Function RPTCopyToVSF(rptList As Object, Optional ByRef vsfInput As Object) As Object
'功能: 将ReportControl全部组强制展开,并将数据复制到vsfInput或另一个公共表格中 , 便于Excel输出
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    Dim vsfReturn As VSFlexGrid
    Dim lngCol As Long, lngRow As Long
    Dim objTmp As frmPubTmp
    Dim intTiers As Integer, rptParent As ReportRow, rptChild As ReportRow
    
    On Error GoTo errH:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    If rptList.Rows.count < 1 Then Exit Function
    If vsfInput Is Nothing Then
        Set objTmp = New frmPubTmp
        Set vsfReturn = objTmp.vsfTmp
    End If
    With IIf(vsfInput Is Nothing, vsfReturn, vsfInput)
        .Clear
         .Rows = 0: .Cols = 0
        .Rows = 1: .FixedRows = 1
        .MergeCells = flexMergeFree
        '标题行复制
        For Each rptCol In rptList.GroupsOrder
            .Cols = .Cols + 1: lngCol = .Cols - 1
            .TextMatrix(0, lngCol) = rptCol.Caption: .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
                Case xtpAlignmentLeft
                    .FixedAlignment(lngCol) = flexAlignLeftCenter: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter
                    .FixedAlignment(lngCol) = flexAlignCenterCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight
                    .FixedAlignment(lngCol) = flexAlignRightCenter: .ColAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
'            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            If rptCol.Width < 20 * IIf(rptList.GroupsOrder.count = 0, 1, rptList.GroupsOrder.count) Then
                .ColWidth(lngCol) = 0
            Else
                .ColWidth(lngCol) = rptCol.Width * Screen.TwipsPerPixelX
            End If
            .MergeCol(lngCol) = True
        Next
        
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .Cols = .Cols + 1: lngCol = .Cols - 1
                .TextMatrix(0, lngCol) = rptCol.Caption: .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                    Case xtpAlignmentLeft
                        .FixedAlignment(lngCol) = flexAlignLeftCenter: .ColAlignment(lngCol) = flexAlignLeftCenter
                    Case xtpAlignmentCenter
                        .FixedAlignment(lngCol) = flexAlignCenterCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                    Case xtpAlignmentRight
                        .FixedAlignment(lngCol) = flexAlignRightCenter: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * Screen.TwipsPerPixelX
                End If
            End If
        Next
        
        '数据行复制
        lngRow = 0
        For Each rptRow In rptList.Rows
            .Rows = .Rows + 1: lngRow = .Rows - 1: .RowHeight(lngRow) = 280
            If rptRow.GroupRow Then
                intTiers = 0: Set rptParent = rptRow
                Do While Not (rptParent.ParentRow Is Nothing)
                    intTiers = intTiers + 1: Set rptParent = rptParent.ParentRow
                Loop
                Set rptChild = rptRow.Childs(0)
                Do While rptChild.GroupRow
                    Set rptChild = rptChild.Childs(0)
                Loop
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = String(intTiers, "　") & rptList.GroupsOrder(intTiers).Caption & ": "
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & rptChild.Record(rptList.GroupsOrder(intTiers).ItemIndex).value
                Next
                .MergeRow(lngRow) = True
            Else
                For lngCol = 0 To .Cols - 1
                    If rptList.Columns(.ColData(lngCol)).TreeColumn Then
                        intTiers = 0: Set rptParent = rptRow
                        Do While Not (rptParent.ParentRow Is Nothing)
                            intTiers = intTiers + 1: Set rptParent = rptParent.ParentRow
                        Loop
                        .TextMatrix(lngRow, lngCol) = String(intTiers, "　") & rptRow.Record(.ColData(lngCol)).value
                        .MergeRow(lngRow) = True
                    Else
                        .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).value
                    End If
                    .Cell(flexcpAlignment, lngRow, lngCol, lngRow) = .ColAlignment(lngCol)
                Next
            End If
        Next
        For lngCol = 0 To .Cols - 1
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
        Next
'        .RowHeight(-1) = .RowHeightMin
    End With
    Set RPTCopyToVSF = IIf(vsfInput Is Nothing, vsfReturn, vsfInput)
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
    Set RPTCopyToVSF = Nothing
End Function

Public Function RPTFindCol(ByRef rpt As Object, ByVal strColumn As String) As Long
'功能：根据列名返回列序号，没有找到时返回-1
    Dim arrTmp As Variant, i As Long
    
    RPTFindCol = -1
    For i = 0 To rpt.Columns.count - 1
        If rpt.Columns(i).Caption = strColumn Then RPTFindCol = i: Exit For
    Next
End Function

Public Sub RTFSetFontSize(ByRef objRTF As Object, ByVal lngFontSize As Long)
'功能：对RichTextBox进行字体设置
'参数：objRTF RichTextBox对象
'      bytSize 0-小字体,1-大字体
    With objRTF
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = lngFontSize
        .SelLength = 0
    End With
End Sub

Public Sub StaShowPercent(sngPercent As Single, objPanel As Object, objForm As Object)
'功能：在状态条上根据百分比显示当前处理进度(█)
'参数：sngPercent=当前进度值,为0-1之间的一个小数
'      objPanel=状态条中用于显示进度的格子
'      objForm=用于计算文字宽度的窗体
    Dim intAll As Integer
    intAll = objPanel.Width / objForm.TextWidth("█") - 4
    objPanel.Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
End Sub

Public Sub TxtGetCaretPos(ByVal lngTxtHwnd As Long, lngLineNo As Long, lngColNo As Long)
'功能：获取光标在文本框中的行列位置
'参数：lngTxtHwnd=要获取光标位置的文本框句柄
'返回：lngLineNo=光标所在行号,以1开始
'      lngColNo=光标所在列号,以0开始
    Dim i As Long, j As Long, K As Long
    Dim lParam As Long, wParam As Long

    '首先向文本框传递EM_GETSEL消息以获取从起始位置到
    '光标所在位置的字符数
    i = SendMessage(lngTxtHwnd, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16
    
    '再向文本框传递EM_LINEFROMCHAR消息根据获得的字符
    '数确定光标以获取所在行数
    lngLineNo = SendMessage(lngTxtHwnd, EM_LINEFROMCHAR, j, 0) '
    lngLineNo = lngLineNo + 1
    
    '向文本框传递EM_LINEINDEX消息以获取所在列数
    K = SendMessage(lngTxtHwnd, EM_LINEINDEX, -1, 0)
    lngColNo = j - K + 1
End Sub

Public Sub TxtGetCursorPos(objMain As Object, objTxt As Object, lngX As Long, lngY As Long)
'功能：获取文本框中当前光标相对于该框的坐标
'参数：objMain=用于计算文字尺寸的窗体
'      objTXT=要获取光标坐标的文本框控件
'返回： lngX,lngY=光标坐标(Twip)
    Dim lngRow As Long, lngCol As Long
    Dim lngFirst As Long
    Dim lngRH As Long, lngCW As Long
    
    objMain.Font.Name = objTxt.Font.Name
    objMain.Font.Size = objTxt.Font.Size
    objMain.Font.Bold = objTxt.Font.Bold
    objMain.Font.Italic = objTxt.Font.Italic
    objMain.Font.Underline = objTxt.Font.Underline
    objMain.Font.Strikethrough = objTxt.Font.Strikethrough
    lngRH = objMain.TextHeight("A")
    lngCW = objMain.TextWidth("A") '是以半角计算列号
    
    lngFirst = SendMessage(objTxt.hWnd, EM_GETFIRSTVISIBLELINE, lngRow, lngCol) + 1 '以0行开始
    Call TxtGetCaretPos(objTxt.hWnd, lngRow, lngCol) '以1行,0列开始
    lngRow = lngRow - lngFirst + 1 '行偏移
    
    lngY = lngRow * lngRH
    lngX = lngCol * lngCW
End Sub

Public Sub TxtCheckKeyPress(ByVal objCtl As Object, KeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '功能:只能输入数字和回车及退格
    '参数:
    '   objctl:限制控件
    '   Keyascii:
    '           Keyascii:8 (退格)
    '   TextType:(0-文本式;1-数字式;2-金额式)
    '返回:一个KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    
    If TextType = m文本式 Then
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        Select Case KeyAscii
        Case vbKeyReturn       '回车
        Case 8                 '退格
        Case Asc(".")
            If TextType = m金额式 Or TextType = m负金额式 Then
                If InStr(objCtl, ".") <> 0 Then     '只能存在一个小数点
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
        Case Asc("-")          '负数
            '70771:刘尔旋,2014-03-06,金额式可以输入负数的问题
            If TextType <> m负金额式 Then KeyAscii = 0: Exit Sub
            If Trim(objCtl.Text) = "" Then Exit Sub
            If objCtl.SelStart <> 0 Then KeyAscii = 0: Exit Sub      '光标不存第一位,不能输入负数
            If InStr(1, objCtl.Text, "-") <> 0 Then   '只能存在一个负数
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
        End Select
    End If
    Exit Sub
Errhand:
    Err.Clear
    KeyAscii = 0
End Sub

Public Sub TxtSelAll(objTxt As Object)
'功能：将编辑框的的文本全部选中
'参数：objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性
    
    If Trim(objTxt.Text) = "" Then Exit Sub
    
    If TypeName(objTxt) = "TextBox" Or TypeName(objTxt) = "ComboBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        
        If TypeName(objTxt) = "TextBox" Then
            If objTxt.MultiLine Then
                SendMessage objTxt.hWnd, WM_VSCROLL, SB_TOP, 0
            End If
        End If
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function TxtCheckInput(objTxt As Object, Optional ByVal strName As String, Optional ByVal intLen As Integer = -1, Optional blnAllowNULL As Boolean) As Boolean
'功能：检查工本框的真实长度是否在指定限制长度内
    Dim intCHeckLen As Integer
    
    If intLen = -1 Then
        intCHeckLen = objTxt.MaxLength
    ElseIf intLen >= 0 Then
        intCHeckLen = intLen
    End If

    If objTxt.Text = "" And Not blnAllowNULL Then
        MsgBox strName & "不允许为空！", vbInformation, gstrSysName
        objTxt.SetFocus: Exit Function
    ElseIf intCHeckLen <> 0 Then
        If gobjComLib.zlCommFun.ActualLen(objTxt.Text) > intCHeckLen Then
            MsgBox strName & "输入过长，只允许输入 " & intCHeckLen & " 个字符或 " & intCHeckLen \ 2 & " 个汉字！", vbInformation, gstrSysName
            objTxt.SetFocus: Exit Function
        End If
    End If
    TxtCheckInput = True
End Function

Public Sub FormSetCaption(ByVal objForm As Variant, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：objForm=传窗体对象，可以传窗体句柄（仅隐藏时blnCaption=false)
'         blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    Dim lngHwnd As Long
    If IsObject(objForm) Then
        lngHwnd = objForm.hWnd
    Else
        lngHwnd = objForm
    End If
    
    Call GetWindowRect(lngHwnd, vRect)
    lngStyle = GetWindowLong(lngHwnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong lngHwnd, GWL_STYLE, lngStyle
    SetWindowPos lngHwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function GetInputItemControlFrm(ByVal strControlPara As String) As Object
'入参:strControlPara:输入项目,是否禁用,是否必输项,光标是否跳入|....
'      是否禁用:1-禁用;0-不禁用
'      是否必输项:1-必输；0-可选输入
'      光标是否跳入:1-光标跳入;0-光标不跳入
'返回: From对象
    Call frmControlPara.zlOpenControlFrm(strControlPara)
    Set GetInputItemControlFrm = frmControlPara
End Function

Public Function IlsAddColorIcon(ByVal strKey As String, ByVal Color As OLE_COLOR, ByRef ils As Object) As Boolean
    '******************************************************************************************************************
    '功能：产生颜色块，并添加到ImageList对象中
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim ctlPictureBox As VB.PictureBox
    Dim objItem As Object
    Dim objTmp As frmPubTmp
    
    On Error Resume Next
    Set objItem = Nothing
    If strKey = "" Then strKey = "K" & Color
    Set objItem = ils.ListImages(strKey)
    Err.Clear
    On Error GoTo errH
    If objItem Is Nothing Then
        Set objTmp = New frmPubTmp
        Set ctlPictureBox = objTmp.Controls.Add("VB.PictureBox", "ctlPictureBox1")
        Dim ListImage As ListImage
        Set ListImage = ils.ListImages("User")
        
        ctlPictureBox.AutoRedraw = True
        ctlPictureBox.AutoSize = True
        ctlPictureBox.BackColor = ils.MaskColor
        ctlPictureBox.Picture = ListImage.ExtractIcon
        
        If Color = vbWhite Then Color = RGB(254, 254, 254)
        ctlPictureBox.Cls
        ctlPictureBox.Line (30, 30)-(ctlPictureBox.Width - 105, ctlPictureBox.Height - 105), Color, BF
        ctlPictureBox.Refresh
    
        On Error Resume Next
        ils.ListImages.Remove ils.ListImages(strKey).Index
        On Error GoTo errH
        ils.ListImages.Add , strKey, ctlPictureBox.Image
        ils.ListImages(strKey).Tag = CStr(Color)
        Set ctlPictureBox = Nothing
    End If
    IlsAddColorIcon = True
    Exit Function
errH:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IlsCheckIcon(ilsInput As Object, strKey As String) As Boolean
    '功能           检查关键字是否在图像列表ImageList中存在，如果存在返回为真
    '参数
    '               ilsInput 传入的图像对象
    '               strKey 要检查当前传入的Key是否存在
    '返回           有返回真，没有返回假
    Dim intLoop As Integer
    On Error Resume Next
    If ilsInput Is Nothing Then Exit Function
    With ilsInput
        For intLoop = 1 To .ListImages.count
            If .ListImages(intLoop).Key = strKey Then
                IlsCheckIcon = True
                Exit Function
            End If
        Next
    End With
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function GetCaretPos(ByVal lngHwnd As Long) As POINTAPI
'功能：返回编辑控件中当前光标的坐标
'参数：lngHwnd=Edit控件的句柄
'返回：坐标值，基于Edit控件,以Twip为单位
'      如果坐标在控件范围之外，则返回(-1,-1)坐标
    Dim lngPos As Long
    Dim vSel As CHARRANGE
    Dim vPos As POINTAPI
    Dim vRect As RECT
    
    SendMessage lngHwnd, EM_EXGETSEL, 0, vSel
    lngPos = SendMessage(lngHwnd, EM_POSFROMCHAR, vSel.cpMin, 0)
    
    vPos.X = lngPos Mod 2 ^ 16
    vPos.Y = lngPos \ 2 ^ 16
    
    '超范围判断
    GetWindowRect lngHwnd, vRect
    If vPos.X >= 0 And vPos.X <= vRect.Right - vRect.Left + 1 _
        And vPos.Y >= 0 And vPos.Y <= vRect.Bottom - vRect.Top + 1 Then
        vPos.X = vPos.X * Screen.TwipsPerPixelX
        vPos.Y = vPos.Y * Screen.TwipsPerPixelY
    Else
        vPos.X = -1: vPos.Y = -1
    End If
    
    GetCaretPos = vPos
End Function

Public Function GetControlRect(ByVal lngHwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'功能：获取指定控件在屏幕中的位置(Twip/Pixel)
'返回：blnTwip=True-返回Twip单位，False-返回像素单位
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function GetClientPoint(ByVal lngHwnd As Long) As POINTAPI
'获取当前指针对应在控件中的位置
    Dim pRet As POINTAPI
    Dim lngReturn As Long
    
    pRet = GetCursorPosition()
    lngReturn = ScreenToClient(lngHwnd, pRet)
    pRet.X = frmPubTmp.ScaleX(pRet.X, vbPixels, vbTwips)
    pRet.Y = frmPubTmp.ScaleY(pRet.Y, vbPixels, vbTwips)
    GetClientPoint = pRet
End Function

Public Function GetCursorPosition() As POINTAPI
'获取鼠标位置
    Dim pRet As POINTAPI
    Dim lngReturn As Long
    lngReturn = GetCursorPos(pRet)
    GetCursorPosition = pRet
End Function

Public Function SetPubFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'功能：设置窗体及所有控件的字体大小
'参数：frmMe=需要设置字体的窗体对象
'      bytSize:设置为9号字体,0:设置为9号字体,1,设置为12号字体
'      strOther:不进行字体设置的控件父容器的集合,格式为：容器名字1,容器名字2,容器名字3,....
'说明：如果存在未列出的其他控件或自定义控件,需要用特定方法指定字体大小及相关处理的，需另外单独设置

    Dim objCtrol As Control, objrptCol As ReportColumn
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "ReportControl", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKind", "PatiIdentify", "VSFlexGrid"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '对于CommandBars用户自定义控件读取objCtrol.Container会出错
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.Name
            Err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Height = frmMe.TextHeight("字") + 20
                        'Label宽度需要自行调整
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("字体" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                        objCtrol.Height = frmMe.TextHeight("字") + IIf(bytSize = 0, 100, 120)
                Case "TextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("字") + 60
                Case "MaskEdBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("字") + 90
                Case "ReportControl"
                        lngOldSize = objCtrol.PaintManager.TextFont.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        Set CtlFont = objCtrol.PaintManager.TextFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.TextFont = CtlFont
                        For Each objrptCol In objCtrol.Columns
                            objrptCol.Width = objrptCol.Width * dblRate
                        Next
                        objCtrol.Redraw
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '控件初始加载时CtlFont为nothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        lngOldSize = objCtrol.FontSize
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                Case "IDKind"
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "PatiIdentify"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        objCtrol.IDKindFont.Size = lngFontSize
                        objCtrol.objIDKind.Refrash
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "VSFlexGrid"
                    Call VSFSetFontSize(objCtrol, lngFontSize)
                Case "RichTextBox"
                    Call RTFSetFontSize(objCtrol, lngFontSize)
            End Select
        End If
    Next
End Function

Public Sub SetPubCtrlPos(ByVal blnVertical As Boolean, ByVal intAligType As Integer, ParamArray arrControls() As Variant)
'功能:对同一行的控件进行位置设置
'参数：
'blnvertical  true ,垂直方向设置控件位置，false,水平方向设置控件位置
'blnvertical=false :intAligType=-1,顶端对齐，0-中间对齐，1-底端对齐,blnvertical=true,intAligType=-1,左对齐，0-水平中心对齐，1-右对齐
'   arrControls格式为控件1,间距1,控件2,间距2,控件3,...
    Dim i As Long
    Dim lngPos As Long '第一个控件的某一位置
    Dim dblRate As Double
    If UBound(arrControls) = -1 Then Exit Sub
    If blnVertical Then
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Left
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Left + 0.5 * arrControls(0).Width
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Left + arrControls(0).Width
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Top = arrControls(i - 2).Top + arrControls(i - 2).Height + arrControls(i - 1)
                arrControls(i).Left = lngPos - arrControls(i).Width * dblRate
            End If
        Next
    Else
        Select Case intAligType
            Case -1
                lngPos = arrControls(0).Top
                dblRate = 0
            Case 0
                lngPos = arrControls(0).Top + 0.5 * arrControls(0).Height
                dblRate = 0.5
            Case 1
                lngPos = arrControls(0).Top + arrControls(0).Height
                dblRate = 1
        End Select
        
        For i = 0 To UBound(arrControls)
            If i > 0 And i Mod 2 = 0 Then
                arrControls(i).Left = arrControls(i - 2).Left + arrControls(i - 2).Width + arrControls(i - 1)
                arrControls(i).Top = lngPos - arrControls(i).Height * dblRate
            End If
        Next
    End If
End Sub

Public Sub ControlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '功能:判断控件是否可获得焦点
    '返回:初如成功,返回true,否则返回False
    '------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
Errhand:
    If gobjComLib.ErrCenter = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Sub ExchangeLocate(objA As Object, objB As Object)
'功能：交换交换两个控件位置
    Dim X1 As Long, Y1 As Long, w1 As Long, t1 As Integer
    Dim X2 As Long, Y2 As Long, w2 As Long, t2 As Integer
    Dim obj1 As Object, obj2 As Object
    
    X1 = objA.Left
    Y1 = objA.Top
    w1 = objA.Width
    t1 = objA.TabIndex
    Set obj1 = objA.Container

    X2 = objB.Left
    Y2 = objB.Top
    w2 = objB.Width
    t2 = objB.TabIndex
    Set obj2 = objB.Container
    
    Set objB.Container = obj1
    If TypeName(objB) = "Label" Then
        objB.Left = X1 + w1 - objB.Width
    Else
        objB.Left = X1
        objB.Width = w1
    End If
    objB.Top = Y1
    objB.TabIndex = t1
    
    Set objA.Container = obj2
    If TypeName(objA) = "Label" Then
        objA.Left = X2 + w2 - objA.Width
    Else
        objA.Left = X2
        objA.Width = w2
    End If
    objA.Top = Y2
    objA.TabIndex = t2
End Sub

Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'功能：判断当前屏幕鼠标是否在指定窗口的显示区域内
    Dim vRect As RECT, vPos As POINTAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Function GetNextControl(ByVal intTab As Integer, ByVal frmForm As Object, Optional ByVal strSkip As String) As Object
'功能：获取下一个光标顺序的控件
'参数：strSkip=要跳过光标的控件名,以";"间隔
    Dim objNext As Object, i As Long
    
    '先找比当前控件TabIndex大的
    For i = 0 To frmForm.Controls.count - 1
        If InStr("TextBox,ComboBox,VSFlexGrid", TypeName(frmForm.Controls(i))) > 0 Then
            If frmForm.Controls(i).Enabled And frmForm.Controls(i).Visible And frmForm.Controls(i).TabStop _
                And (InStr(";" & strSkip & ";", ";" & frmForm.Controls(i).Name & ";") = 0 Or strSkip = "") Then
                If frmForm.Controls(i).TabIndex > intTab Then
                    If objNext Is Nothing Then
                        Set objNext = frmForm.Controls(i)
                    ElseIf frmForm.Controls(i).TabIndex < objNext.TabIndex Then
                        Set objNext = frmForm.Controls(i)
                    End If
                End If
            End If
        End If
    Next
    If objNext Is Nothing Then
        '没有则找比当前控件TabIndex小的
        For i = 0 To frmForm.Controls.count - 1
            If InStr("TextBox,ComboBox,VSFlexGrid", TypeName(frmForm.Controls(i))) > 0 Then
                If frmForm.Controls(i).Enabled And frmForm.Controls(i).Visible And frmForm.Controls(i).TabStop _
                    And (InStr(";" & strSkip & ";", ";" & frmForm.Controls(i).Name & ";") = 0 Or strSkip = "") Then
                    If frmForm.Controls(i).TabIndex < intTab Then
                        If objNext Is Nothing Then
                            Set objNext = frmForm.Controls(i)
                        ElseIf frmForm.Controls(i).TabIndex < objNext.TabIndex Then
                            Set objNext = frmForm.Controls(i)
                        End If
                    End If
                End If
            End If
        Next
    End If
    Set GetNextControl = objNext
End Function

Public Sub LoadOldData(ByVal strOld As String, objTxt As Object, objCbo As Object)
'功能:将数据库中保存的年龄按规范的格式加载到界面,不规范的原样显示
'参数：strOld =年龄字符串
'          objTxt=显示年龄的文本框
'          objCbo=年龄单位下拉列表
'      intIndex=加载年龄的控件索引值值

    Dim strTmp As String, lngIdx As Long
    Dim arrTmp As Variant
    Dim i As Long
    
    If Trim(strOld) = "" Then Exit Sub
    '获取所有年龄单位
    For i = 0 To objCbo.ListCount
        strTmp = strTmp & "," & gobjComLib.zlCommFun.GetNeedName(objCbo.List(i))
    Next
    If strTmp = "" Then
        objTxt.Text = strOld
        objCbo.Visible = False
        Exit Sub
    End If
    strTmp = Mid(strTmp, 2)
    arrTmp = Split(strTmp, ",")
    strTmp = strOld
    For i = LBound(arrTmp) To UBound(arrTmp)
        If InStr(strOld, arrTmp(i)) > 0 Then
            If InStr(strOld, arrTmp(i)) + Len(arrTmp(i)) - 1 = Len(strOld) Then
                strTmp = Mid(strOld, 1, InStr(strOld, arrTmp(i)) - 1)
                lngIdx = i
            End If
            Exit For
        End If
    Next
    
    'IsNumeric("")=False,因此连接上字符串1
    If Not IsNumeric(strTmp & "1") Then
        lngIdx = -1
        strTmp = strOld
    End If
    
    objTxt.Text = strTmp
    If objCbo.ListCount > 0 Then Call gobjComLib.zlControl.CboSetIndex(objCbo.hWnd, lngIdx)
    If lngIdx = -1 Then
        objCbo.Visible = False
    Else
        If objCbo.Visible = False Then objCbo.Visible = True
    End If
End Sub

Private Sub Class_Terminate()
    Set mvarclsComboBox = Nothing
End Sub
