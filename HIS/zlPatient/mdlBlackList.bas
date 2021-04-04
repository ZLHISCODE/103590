Attribute VB_Name = "mdlBlackList"
Option Explicit

Public Const G_AlternateColor As Long = 16772055   '行交替色
Public Const G_LostFocusColor As Long = &HE0E0E0   '失去焦点时的网格背景色
Public Enum gEM_BlackListFun
    Em_Pane_FunFace = 1
    Em_Pane_Face = 2
    Em_Pane_Type = 11 '不良记录分类
    Em_Pane_Reason = 12 '常用的不良记录原因
    Em_Pane_Record = 13 '不良记录管理
End Enum

Public Function zlGetFirstCommandBar(ByRef objControls As CommandBarControls) As Long
    '功能：获取工具栏打印预览按钮后的第一个按钮的index
    Dim objControl As CommandBarControl, idx As Long
    
    For Each objControl In objControls
        If objControl.ID = conMenu_File_Preview Then
            idx = objControl.Index + 1
        End If
    Next
    zlGetFirstCommandBar = idx
End Function

Public Function zlGetPopupCommandBar(frmMain As Form, cbsMain As CommandBars, _
    Optional ByVal lngControlPopupID As Long = conMenu_EditPopup) As CommandBar
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:构建弹出菜单
    '返回:返回弹出菜单对象
    '编制:刘兴洪
    '日期:2018-11-08 11:21:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup, cbCommandBar As CommandBar
    Dim cbrControl As CommandBarControl, cbrControlNew As CommandBarControl
    Dim i As Integer
    
    On Error GoTo errHandle
      
    Set objPopup = cbsMain.FindControl(xtpControlPopup, lngControlPopupID, , True)
    If objPopup Is Nothing Then Exit Function
    Set cbCommandBar = cbsMain.Add("Popup", xtpBarPopup) '弹出菜单
    If cbCommandBar Is Nothing Then Exit Function
    
    For i = 1 To objPopup.CommandBar.Controls.Count
        Set cbrControl = objPopup.CommandBar.Controls(i)
        Call frmMain.zlUpdateCommandBars(cbrControl) '判断是否可见，因为第一次时菜单还没有执行Update
        If cbrControl.Visible Then
            Set cbrControlNew = cbCommandBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
            cbrControlNew.BeginGroup = cbrControl.BeginGroup
            cbrControlNew.Enabled = cbrControl.Enabled
        End If
    Next
    Set zlGetPopupCommandBar = cbCommandBar
    
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetCommbarFromName(ByVal objThis As CommandBars, ByVal strName As String, Optional intIndex_Out As Integer) As CommandBar
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据菜单名称，获取指定的菜单对象
    '入参:strName-名称
    '出参:intIndex_Out-返回的索引
    '返回:成功返回CommandBar,否则返回False
    '编制:刘兴洪
    '日期:2018-11-15 15:13:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    
    For i = 1 To objThis.Count
        If objThis(i).Title = strName Then
        
            Set GetCommbarFromName = objThis(i)
            intIndex_Out = i: Exit Function
        End If
    Next
    intIndex_Out = 0
    Set GetCommbarFromName = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub SetReportControlBackColorAlternate(rptData As ReportControl, Optional CustomColor As OLE_COLOR = -1)
    '设置ReportControl的行交替色
    Dim i As Long, objItem As ReportRecordItem
    Dim lngRowCount As Long '组内行号
    
    On Error Resume Next
    For i = 0 To rptData.Rows.Count - 1
        If rptData.Rows(i).GroupRow Then
            lngRowCount = 0
        Else
            For Each objItem In rptData.Rows(i).Record
                If lngRowCount Mod 2 = 0 Then
                    objItem.BackColor = rptData.PaintManager.BackColor
                Else
                    objItem.BackColor = IIf(CustomColor = -1, G_AlternateColor, CustomColor)
                End If
            Next
            lngRowCount = lngRowCount + 1
        End If
    Next
End Sub
Public Function zlGetVsfGrid(rptData As ReportControl, ByRef vsGrid As VSFlexGrid, Optional ByVal strHiddenCols As String) As Boolean
    '功能:将ReportControl转换为VSFlexGrid
    '入参:
    '   strHiddenCols 隐藏列索引(索引从0开始)，格式：列1,列2,列3,...
    
    Dim i As Long, j As Long, lngRowIndex As Long
    Dim varData As Variant
    
    Err = 0: On Error GoTo ErrHandler
    With vsGrid
        .Clear
        .Cols = rptData.Columns.Count
        .Rows = rptData.Records.Count + 1
        .FixedAlignment(-1) = flexAlignCenterCenter
        
        '标题行
        For i = 0 To rptData.Columns.Count - 1
            .TextMatrix(0, i) = rptData.Columns(i).Caption
            .ColWidth(i) = rptData.Columns(i).Width * 16
            .ColAlignment(i) = Choose(rptData.Columns(i).Alignment + 1, 1, 4, 7)
        Next
        '隐藏列
        If strHiddenCols <> "" Then
            varData = Split(strHiddenCols, ",")
            For i = 0 To UBound(varData)
                .ColWidth(Val(varData(i))) = 0
            Next
        End If
        
        '数据行
        lngRowIndex = 1
        For i = 0 To rptData.Rows.Count - 1
            If rptData.Rows(i).GroupRow = False Then
                For j = 0 To rptData.Columns.Count - 1
                    .TextMatrix(lngRowIndex, j) = rptData.Rows(i).Record(j).Value
                Next
                lngRowIndex = lngRowIndex + 1
            End If
        Next
    End With
    zlGetVsfGrid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln负数检查 As Boolean = True, Optional bln零检查 As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     bln负数检查     是否进行负数检查
    '     bln零检查         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    zlDblIsValid = zlCommFun.DblIsValid(strInput, intMax, bln负数检查, bln零检查, hWnd, str项目)
End Function

