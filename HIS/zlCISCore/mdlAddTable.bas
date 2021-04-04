Attribute VB_Name = "mdlAddTable"
Option Explicit
Public Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private clsComLib As New zl9ComLib.clsComLib
Private clsDatabase As New zl9ComLib.clsDatabase

Public Sub AddObject(theTable As TTF160Ctl.F1Book, ByVal Row As Long, ByVal Col As Long, VisItemID As String, Optional ByVal Editable As Boolean = True, Optional ByVal DefaultValue As String = "", Optional objParent As Object)
'theTable的父窗体要加载VisItem控件数组
    Dim objID As Long
    Dim X1 As Long, Y1 As Long, iWidth As Long, iHeight As Long, iShown As Integer
    Dim sItemID As String
    Dim rsTmp As New ADODB.Recordset, aValues() As String, i As Long, iValueNum As Long
    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim frmParent As Object, aVisItemInfo() As String
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    Dim iTableIndex As Integer
    Dim iItemLen As Integer, sItemFormat As String, sDefaultValue As String
    
    clsDatabase.OpenRecordset rsTmp, "Select * From 诊治所见项目 Where ID='" + VisItemID + "'", ""
    If rsTmp.EOF Then Exit Sub
    
    On Error Resume Next
    iCurrRow = theTable.Row: iCurrCol = theTable.Col
    iStartRow = theTable.SelStartRow: iEndRow = theTable.SelEndRow
    iStartCol = theTable.SelStartCol: iEndCol = theTable.SelEndCol
        
    theTable.SetSelection iStartRow, iStartCol, iStartRow, iStartCol
    theTable.SetActiveCell Row, Col
    
    If objParent Is Nothing Then
        Set frmParent = theTable.Parent '父窗体
    Else
        Set frmParent = objParent
    End If
    Load frmParent.VisItem(frmParent.VisItem.UBound + 1)
    With frmParent.VisItem(frmParent.VisItem.UBound)
        sDefaultValue = IIf(IsNull(rsTmp("初始值")), "", rsTmp("初始值"))
        Select Case rsTmp("编码")
            Case "101001" '年(YYYY)
                iItemLen = 4: sItemFormat = "YYYY"
            Case "101002" '月(YYYY-MM)
                iItemLen = 7: sItemFormat = "YYYY-MM"
            Case "101003" '日(YYYY-MM-DD)
                iItemLen = 10: sItemFormat = "YYYY-MM-DD"
            Case "101004" '时间(YYYY-MM-DD HH24:MI:SS)
                iItemLen = 19: sItemFormat = "YYYY-MM-DD HH:MM:SS"
            Case "101005" '时间(HH24:MI:SS)
                iItemLen = 8: sItemFormat = "HH:MM:SS"
            Case "101006" '时间(HH24:MI)
                iItemLen = 5: sItemFormat = "HH:MM"
            Case "101008" '当前日期
                iItemLen = 10: sItemFormat = "YYYY-MM-DD"
                sDefaultValue = Format(Date, sItemFormat)
            Case "101009" '当前时间
                iItemLen = 19: sItemFormat = "YYYY-MM-DD HH:MM:SS"
                sDefaultValue = Format(zlDatabase.Currentdate, sItemFormat)
            Case Else
                iItemLen = IIf(IsNull(rsTmp("长度")), 10, rsTmp("长度"))
                sItemFormat = ""
        End Select
        .Init "", "", rsTmp("表示法"), rsTmp("类型"), iItemLen, _
            IIf(IsNull(rsTmp("小数")), 0, rsTmp("小数")), _
            IIf(IsNull(rsTmp("数值域")), "", rsTmp("数值域")), _
            IIf(Len(DefaultValue) = 0, sDefaultValue, DefaultValue), _
            rsTmp("ID"), IIf(IsNull(rsTmp("替换域")), "", IIf(rsTmp("替换域") = 1, rsTmp("中文名"), "")), , , sItemFormat
        .Enabled = Editable Or Not theTable.EnableProtection
                
        '获取单元格（包括合并单元）位置、大小
        Set objRect = theTable.RangeToTwipsEx(theTable.SelStartRow, theTable.SelStartCol, theTable.SelEndRow, theTable.SelEndCol)
        .Left = objRect.Left + theTable.Left + 30
        .Top = objRect.Top + theTable.Top + 30
        .Width = objRect.Width - 30
        .Height = objRect.Height - 30
        If objRect.Width - 30 < .Width Then
            theTable.ColWidthTwips(theTable.SelStartCol) = _
                theTable.ColWidthTwips(theTable.SelStartCol) + .Width - (objRect.Width - 30)
        End If
        If objRect.Height - 30 < .Height Then
            theTable.RowHeight(theTable.SelStartRow) = _
                theTable.RowHeight(theTable.SelStartRow) + .Height - (objRect.Height - 30)
        End If
        '记录此所见项对应的单元格坐标
        iTableIndex = -1: iTableIndex = theTable.Index
        .Tag = theTable.SelStartRow & "," & theTable.SelStartCol & "," & iTableIndex
        
        Set .Container = theTable.Container
        .ZOrder 0
        If .Left < theTable.Left Or .Left + .Width > theTable.Left + theTable.Width Or _
            .Top < theTable.Top Or .Top + .Height > theTable.Top + theTable.Height Then
            .Visible = False
        Else
            .Visible = True
        End If
    End With
    With theTable
        Set objCellFormat = .GetCellFormat
        '卸载单元以前关联的所见项
        If Len(objCellFormat.ValidationText) > 0 Then
            aVisItemInfo = Split(objCellFormat.ValidationText, ",")
            Unload frmParent.VisItem(aVisItemInfo(1))
        End If
        objCellFormat.ValidationText = VisItemID & "," & (frmParent.VisItem.UBound)
        .SetCellFormat objCellFormat
        .Text = ""
    End With
        
    theTable.SetSelection iStartRow, iStartCol, iEndRow, iEndCol
    theTable.SetActiveCell iCurrRow, iCurrCol
End Sub

Public Sub RemoveObject(theTable As TTF160Ctl.F1Book, ByVal Row As Long, ByVal Col As Long, Optional objParent As Object)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim frmParent As Object, aVisItemInfo() As String
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    
    On Error Resume Next
    iCurrRow = theTable.Row: iCurrCol = theTable.Col
    iStartRow = theTable.SelStartRow: iEndRow = theTable.SelEndRow
    iStartCol = theTable.SelStartCol: iEndCol = theTable.SelEndCol
        
    theTable.SetSelection iStartRow, iStartCol, iStartRow, iStartCol
    theTable.SetActiveCell Row, Col
    
    If objParent Is Nothing Then
        Set frmParent = theTable.Parent '父窗体
    Else
        Set frmParent = objParent
    End If
    With theTable
        Set objCellFormat = .GetCellFormat
        '卸载单元以前关联的所见项
        If Len(objCellFormat.ValidationText) > 0 Then
            aVisItemInfo = Split(objCellFormat.ValidationText, ",")
            Unload frmParent.VisItem(aVisItemInfo(1))
        End If
        objCellFormat.ValidationText = ""
        .SetCellFormat objCellFormat
        
        .Text = ""
    End With
        
    theTable.SetSelection iStartRow, iStartCol, iEndRow, iEndCol
    theTable.SetActiveCell iCurrRow, iCurrCol
End Sub
'处理表格的滚动等事件
Public Sub Proc_Table_TopLeftChanged(theTable As TTF160Ctl.F1Book, Optional objParent As Object)
    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    Dim tmpCtrl As Control, aCellRC() As String
    Dim bValidCtrl As Boolean
    Dim frmParent As Object
        
    On Error Resume Next
    If objParent Is Nothing Then
        Set frmParent = theTable.Parent '父窗体
    Else
        Set frmParent = objParent
    End If
    With theTable
        iCurrRow = .Row: iCurrCol = .Col
        iStartRow = .SelStartRow: iEndRow = .SelEndRow
        iStartCol = .SelStartCol: iEndCol = .SelEndCol

        .SetSelection iStartRow, iStartCol, iStartRow, iStartCol
        For Each tmpCtrl In frmParent.Controls
            bValidCtrl = True
            If Not (tmpCtrl.Name = "VisItem" And Len(tmpCtrl.Tag) > 0 And tmpCtrl.Container.hwnd = .Container.hwnd) Then bValidCtrl = False
            
            If bValidCtrl Then
                aCellRC = Split(tmpCtrl.Tag, ",")
                .SetActiveCell aCellRC(0), aCellRC(1)
    
                tmpCtrl.Visible = False
                '单元可见
                If .RangeShown(.SelStartRow, .SelStartCol, .SelEndRow, .SelEndCol) = 1 Then
                    Set objRect = .RangeToTwipsEx(.SelStartRow, .SelStartCol, .SelEndRow, .SelEndCol)
        
                    tmpCtrl.Left = objRect.Left + .Left + 30
                    tmpCtrl.Top = objRect.Top + .Top + 30
                    tmpCtrl.Width = objRect.Width - 30
                    tmpCtrl.Height = objRect.Height - 30
                    If objRect.Width - 30 < tmpCtrl.Width Then
                        .ColWidthTwips(.SelStartCol) = _
                            .ColWidthTwips(.SelStartCol) + tmpCtrl.Width - (objRect.Width - 30)
                    End If
                    If objRect.Height - 30 < tmpCtrl.Height Then
                        .RowHeight(.SelStartRow) = _
                            .RowHeight(.SelStartRow) + tmpCtrl.Height - (objRect.Height - 30)
                    End If
                    tmpCtrl.Visible = True
                End If
            End If
        Next
        .SetSelection iStartRow, iStartCol, iEndRow, iEndCol
        .SetActiveCell iCurrRow, iCurrCol
    End With
End Sub
'重新刷新表内所见项
Public Sub RefreshObject(theTable As TTF160Ctl.F1Book, Optional objParent As Object, Optional ByVal HasVisItem As Boolean = True, Optional objProgBar As ProgressBar)
    Dim iDecPos As Integer
    Dim objCellFormat As TTF160Ctl.F1CellFormat, objRect As TTF160Ctl.F1Rect
    Dim iCurrRow As Integer, iCurrCol As Integer
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    Dim tmpCtrl As Control, aCellRC() As String, iRow As Integer, iCol As Integer, aVisItemInfo() As String
    Dim frmParent As Object
    
    On Error Resume Next
    If objParent Is Nothing Then
        Set frmParent = theTable.Parent '父窗体
    Else
        Set frmParent = objParent
    End If
    iCurrRow = theTable.Row: iCurrCol = theTable.Col
    iStartRow = theTable.SelStartRow: iEndRow = theTable.SelEndRow
    iStartCol = theTable.SelStartCol: iEndCol = theTable.SelEndCol

    theTable.SetSelection iStartRow, iStartCol, iStartRow, iStartCol
    For Each tmpCtrl In frmParent.Controls
        If tmpCtrl.Name = "VisItem" Then
            If tmpCtrl.Container.hwnd = theTable.Container.hwnd Then tmpCtrl.Visible = False
        End If
    Next
        
    objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = theTable.MaxRow * theTable.MaxCol
    For iRow = 1 To theTable.MaxRow
        For iCol = 1 To theTable.MaxCol
            theTable.SetActiveCell iRow, iCol

            Set objCellFormat = theTable.GetCellFormat
            If Len(objCellFormat.ValidationText) > 0 And iRow = theTable.SelStartRow And iCol = theTable.SelStartCol Then
                aVisItemInfo = Split(objCellFormat.ValidationText, ",")
                
                objCellFormat.ValidationText = ""
                theTable.SetCellFormat objCellFormat
                
                If Not HasVisItem Then
                    AddObject theTable, iRow, iCol, CLng(aVisItemInfo(0)), False, theTable.TextRC(iRow, iCol), frmParent
                Else
                    AddObject theTable, iRow, iCol, CLng(aVisItemInfo(0)), False, frmParent.VisItem(aVisItemInfo(1)).Value, frmParent
                End If
                With frmParent.VisItem(frmParent.VisItem.UBound)
'                    Set .Container = theTable.Container
                    .Visible = True: .Enabled = False
                End With
            End If
                
            objProgBar.Value = (iRow - 1) * theTable.MaxCol + iCol
        Next iCol
    Next iRow
    For Each tmpCtrl In frmParent.Controls
        If tmpCtrl.Name = "VisItem" Then
            If tmpCtrl.Container.hwnd = theTable.Container.hwnd And Not tmpCtrl.Visible Then Unload tmpCtrl
        End If
    Next
    theTable.SetSelection iStartRow, iStartCol, iEndRow, iEndCol
    theTable.SetActiveCell iCurrRow, iCurrCol
End Sub

Public Sub SaveTable(theTable As TTF160Ctl.F1Book, Optional ByVal Seq As Integer = 1, Optional objParent As Object, Optional objProgBar As ProgressBar)
    Dim ElementID As Long, ItemID As String
    Dim MergeNO As String, Locked As String, bLocked As Boolean, bHide As Boolean
    Dim i As Long, j As Long
    Dim OldRow As Long, OldCol As Long, aTmp() As String
    Dim frmParent As Object
    
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    Dim aVisItemInfo() As String
    Dim strCellText As String
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    
    On Error Resume Next
    If objParent Is Nothing Then
        Set frmParent = theTable.Parent '父窗体
    Else
        Set frmParent = objParent
    End If
    With theTable
        '新增附加表元素
        aTmp = Split(theTable.Tag, ";")
        ElementID = CLng(aTmp(0))
        
        '保存报表属性
        clsDatabase.ExecuteProcedure "ZL_附加表_SAVE(" & ElementID & "," & Seq & "," & .FixedRows & "," & .FixedCols & "," & .MaxRow & "," & .MaxCol & ")", ""
        
        OldRow = .Row: OldCol = .Col
        iStartRow = .SelStartRow: iEndRow = .SelEndRow
        iStartCol = .SelStartCol: iEndCol = .SelEndCol
        
        .SetSelection iStartRow, iStartCol, iStartRow, iStartCol
        
        objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = .MaxRow * .MaxCol
        For i = 1 To .MaxRow
            For j = 1 To .MaxCol
                .SetActiveCell i, j
                
                Set objCellFormat = .GetCellFormat
                If Len(objCellFormat.ValidationText) > 0 And i = .SelStartRow And j = .SelStartCol Then
                    aVisItemInfo = Split(objCellFormat.ValidationText, ",")
                    ItemID = aVisItemInfo(0)
                    
                    strCellText = frmParent.VisItem(aVisItemInfo(1)).Value
                Else
                    ItemID = ""
                    
                    strCellText = .EntryRC(i, j)
                End If
                
                If .SelStartRow <> .SelEndRow Or .SelStartCol <> .SelEndCol Then
                    MergeNO = Mid(CStr(10000 + .SelStartRow), 2) & Mid(CStr(10000 + .SelStartCol), 2) & Mid(CStr(10000 + .SelEndRow), 2) & Mid(CStr(10000 + .SelEndCol), 2)
                Else
                    MergeNO = 0
                End If
                
                .GetProtection bLocked, bHide
                Locked = IIf(bLocked, "1", "0")
                clsDatabase.ExecuteProcedure "ZL_附加表单元_SAVE(" & ElementID & "," & Seq & "," & i & "," & j & "," & _
                    .ColWidthTwips(j) & "," & .RowHeight(i) & "," & .HAlign & "," & _
                    MergeNO & "," & Locked & ",'" & ItemID & "','" & Replace(Format(strCellText, IIf(.NumberFormat Like "?.*", .NumberFormat, "")), "'", "''") & "')", ""
                    
                objProgBar.Value = (i - 1) * .MaxCol + j
            Next j
        Next i
        .SetSelection iStartRow, iStartCol, iEndRow, iEndCol
        .Row = OldRow: .Col = OldCol
    End With
End Sub

Public Sub SaveTable_Patient(ElementID As String, theTable As TTF160Ctl.F1Book, cnOracle As ADODB.Connection, Optional ByVal Seq As Integer = 1, Optional objParent As Object, Optional objProgBar As ProgressBar)
    Dim ItemID As String
    Dim MergeNO As String, Locked As String, bLocked As Boolean, bHide As Boolean
    Dim i As Long, j As Long
    Dim OldRow As Long, OldCol As Long, aTmp() As String
    Dim frmParent As Object
    
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    Dim aVisItemInfo() As String
    Dim strCellText As String
    Dim iStartRow As Long, iEndRow As Long, iStartCol As Long, iEndCol As Long
    
    On Error Resume Next
    If objParent Is Nothing Then
        Set frmParent = theTable.Parent '父窗体
    Else
        Set frmParent = objParent
    End If
    With theTable
        '保存报表属性
        cnOracle.Execute "ZL_病人病历附加表_SAVE(" & ElementID & "," & Seq & "," & .FixedRows & "," & .FixedCols & "," & .MaxRow & "," & .MaxCol & ")", , adCmdStoredProc
        
        OldRow = .Row: OldCol = .Col
        iStartRow = .SelStartRow: iEndRow = .SelEndRow
        iStartCol = .SelStartCol: iEndCol = .SelEndCol
        
        .SetSelection iStartRow, iStartCol, iStartRow, iStartCol
        
        objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = .MaxRow * .MaxCol
        For i = 1 To .MaxRow
            For j = 1 To .MaxCol
                .SetActiveCell i, j
                
                Set objCellFormat = .GetCellFormat
                If Len(objCellFormat.ValidationText) > 0 And i = .SelStartRow And j = .SelStartCol Then
                    aVisItemInfo = Split(objCellFormat.ValidationText, ",")
                    ItemID = aVisItemInfo(0)
                    
                    strCellText = frmParent.VisItem(aVisItemInfo(1)).Value
                Else
                    ItemID = ""
                    
                    strCellText = .EntryRC(i, j)
                End If
                
                If .SelStartRow <> .SelEndRow Or .SelStartCol <> .SelEndCol Then
                    MergeNO = Mid(CStr(10000 + .SelStartRow), 2) & Mid(CStr(10000 + .SelStartCol), 2) & Mid(CStr(10000 + .SelEndRow), 2) & Mid(CStr(10000 + .SelEndCol), 2)
                Else
                    MergeNO = 0
                End If
                
                .GetProtection bLocked, bHide
                Locked = IIf(bLocked, "1", "0")
                cnOracle.Execute "ZL_病人病历附加表单元_SAVE(" & ElementID & "," & Seq & "," & i & "," & j & "," & _
                    .ColWidthTwips(j) & "," & .RowHeight(i) & "," & .HAlign & "," & _
                    MergeNO & "," & Locked & ",'" & ItemID & "','" & Replace(Format(strCellText, IIf(.NumberFormat Like "?.*", .NumberFormat, "")), "'", "''") & "')", , adCmdStoredProc
                    
                objProgBar.Value = (i - 1) * .MaxCol + j
            Next j
        Next i
        .SetSelection iStartRow, iStartCol, iEndRow, iEndCol
        .Row = OldRow: .Col = OldCol
    End With
End Sub

Public Sub ReadTable(theTable As TTF160Ctl.F1Book, ByVal ElementID As Long, Optional ByVal Seq As Integer = 1, Optional objProgBar As ProgressBar)
    Dim ItemID As String
    Dim MergeNO As String, Locked As String, bLocked As Boolean, bHide As Boolean
    Dim i As Long, j As Long
    Dim OldRow As Long, OldCol As Long
    Dim rsTmp As New ADODB.Recordset
    Dim cellFormat As TTF160Ctl.F1CellFormat
    Dim iDecPos As Integer
    
    On Error Resume Next
    With theTable
        '读取报表属性
        .Tag = ElementID
        clsDatabase.OpenRecordset rsTmp, "Select * From 病历所见单 Where 元素ID=" & ElementID & " And (控件号=-" & Seq & " Or 控件号 Is Null) And 控件类=3", ""
        If rsTmp.EOF Then Exit Sub
        
        .FixedRows = rsTmp("固定行")
        .FixedCols = rsTmp("固定列")
        .MaxRow = rsTmp("行")
        .MaxCol = rsTmp("列")
        
        clsDatabase.OpenRecordset rsTmp, "Select * From 病历所见单 Where 元素ID=" & ElementID & " And (控件号=-" & Seq & " Or 控件号=0) And 控件类 is Null", ""
        OldRow = .Row: OldCol = .Col
        
        objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = rsTmp.RecordCount
        Do While Not rsTmp.EOF
            i = Abs(rsTmp("行")): j = Abs(rsTmp("列"))
            .SetActiveCell i, j
            
            bLocked = IIf(rsTmp("不可写") = 1, True, False)
            .SetProtection bLocked, False
            .ColWidthTwips(j) = rsTmp("宽")
            .RowHeight(i) = rsTmp("高")
            .HAlign = rsTmp("对齐")
            If IsNumeric(.Text) Then
                iDecPos = InStr(.Text, ".")
                If iDecPos > 0 And iDecPos < Len(.Text) Then
                    .NumberFormat = "#." + String(Len(.Text) - iDecPos, "0")
                Else
                    .NumberFormat = "General"
                End If
            Else
                .NumberFormat = "General"
            End If
            
            MergeNO = rsTmp("合并号")
            If MergeNO <> "0" Then
                MergeNO = "0000" & MergeNO
                MergeNO = Mid(MergeNO, Len(MergeNO) - 15, 16)
                If i = CLng(Mid(MergeNO, 1, 4)) And j = CLng(Mid(MergeNO, 5, 4)) Then
                    .SetSelection i, j, CLng(Mid(MergeNO, 9, 4)), CLng(Mid(MergeNO, 13, 4))
                    Set cellFormat = .GetCellFormat
                    cellFormat.MergeCells = True
                    .SetCellFormat cellFormat
                End If
            End If
            
            If Not IsNull(rsTmp("所见项ID")) Then
                ItemID = rsTmp("所见项ID")
                AddObject theTable, i, j, ItemID, IIf(rsTmp("不可写") = 0, True, False), IIf(IsNull(rsTmp("缺省内容")), "", rsTmp("缺省内容"))
            Else
                .Text = rsTmp("缺省内容")
            End If
                    
            objProgBar.Value = rsTmp.AbsolutePosition
            
            rsTmp.MoveNext
        Loop
        .Row = OldRow: .Col = OldCol
    End With
End Sub

Public Sub ReadTable_Patient(theTable As TTF160Ctl.F1Book, ByVal ElementID As Long, Optional ByVal Seq As Integer = 1, Optional objProgBar As ProgressBar)
    Dim ItemID As String
    Dim MergeNO As String, Locked As String, bLocked As Boolean, bHide As Boolean
    Dim i As Long, j As Long
    Dim OldRow As Long, OldCol As Long
    Dim rsTmp As New ADODB.Recordset
    Dim cellFormat As TTF160Ctl.F1CellFormat
    Dim iDecPos As Integer
    
    On Error Resume Next
    With theTable
        '读取报表属性
        clsDatabase.OpenRecordset rsTmp, "Select * From 病人病历所见单 Where 病历ID=" & ElementID & " And 控件号=-" & Seq & " And 控件类=3", ""
        If rsTmp.EOF Then Exit Sub
        
        .FixedRows = rsTmp("固定行")
        .FixedCols = rsTmp("固定列")
        .MaxRow = rsTmp("行")
        .MaxCol = rsTmp("列")
        
        clsDatabase.OpenRecordset rsTmp, "Select * From 病人病历所见单 Where 病历ID=" & ElementID & " And 控件号=-" & Seq & " And 控件类 is Null", ""
        OldRow = .Row: OldCol = .Col
        
        objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = rsTmp.RecordCount
        Do While Not rsTmp.EOF
            i = Abs(rsTmp("行")): j = Abs(rsTmp("列"))
            .SetActiveCell i, j
            
            bLocked = IIf(rsTmp("不可写") = 1, True, False)
            .SetProtection bLocked, False
            .ColWidthTwips(j) = rsTmp("宽")
            .RowHeight(i) = rsTmp("高")
            .HAlign = rsTmp("对齐")
            If IsNumeric(.Text) Then
                iDecPos = InStr(.Text, ".")
                If iDecPos > 0 And iDecPos < Len(.Text) Then
                    .NumberFormat = "#." + String(Len(.Text) - iDecPos, "0")
                Else
                    .NumberFormat = "General"
                End If
            Else
                .NumberFormat = "General"
            End If
            
            MergeNO = rsTmp("合并号")
            If MergeNO <> "0" Then
                MergeNO = "0000" & MergeNO
                MergeNO = Mid(MergeNO, Len(MergeNO) - 15, 16)
                If i = CLng(Mid(MergeNO, 1, 4)) And j = CLng(Mid(MergeNO, 5, 4)) Then
                    .SetSelection i, j, CLng(Mid(MergeNO, 9, 4)), CLng(Mid(MergeNO, 13, 4))
                    Set cellFormat = .GetCellFormat
                    cellFormat.MergeCells = True
                    .SetCellFormat cellFormat
                End If
            End If
            
            If Not IsNull(rsTmp("所见项ID")) Then
                ItemID = rsTmp("所见项ID")
                AddObject theTable, i, j, ItemID, IIf(rsTmp("不可写") = 0, True, False), IIf(IsNull(rsTmp("所见内容")), "", rsTmp("所见内容"))
            Else
                .Text = rsTmp("所见内容")
            End If
                    
            objProgBar.Value = rsTmp.AbsolutePosition
            
            rsTmp.MoveNext
        Loop
        .Row = OldRow: .Col = OldCol
    End With
End Sub

Public Sub ClearAllObject(theTable As TTF160Ctl.F1Book, Optional objParent As Object)
    Dim tmpCtrl As Control
    Dim frmParent As Object
    
    
    On Error Resume Next
'    Dim tmpID As Long
'
'    On Error GoTo EndSearch
'    tmpID = theTable.ObjFirstID
'    Do While True
'        theTable.ObjSetSelection tmpID
'        theTable.EditClear F1ClearAll
'
'        tmpID = theTable.ObjNextID(tmpID)
'    Loop
'EndSearch:
    If objParent Is Nothing Then
        Set frmParent = theTable.Parent '父窗体
    Else
        Set frmParent = objParent
    End If
    For Each tmpCtrl In frmParent.Controls
        If tmpCtrl.Name = "VisItem" Then
            If tmpCtrl.Container.hwnd = theTable.Container.hwnd Then Unload tmpCtrl
        End If
    Next
End Sub

Public Sub InitTable(theTable As TTF160Ctl.F1Book)
    Dim cellFormat  As TTF160Ctl.F1CellFormat
    With theTable
        .MaxRow = 10: .MaxCol = 20: .FixedRows = 0: .FixedCols = 0
        
        .SetColWidthTwips 1, .MaxCol, 961, True
        .SetRowHeight 1, .MaxRow, 255, True
        
        .SetSelection 1, 1, .MaxRow, .MaxCol
        .HAlign = F1HAlignGeneral
        .EditClear F1ClearAll
        .WordWrap = True
        Set cellFormat = .GetCellFormat
        cellFormat.ProtectionLocked = False
        cellFormat.MergeCells = False
        .SetCellFormat cellFormat
        .SetSelection 1, 1, 1, 1
        
        ClearAllObject theTable
    End With
End Sub

Public Sub SaveForm(theForm As Form, ContainerName As String, ByVal VisFormID As String, Optional objProgBar As ProgressBar)
    On Error GoTo DBError
    
    SaveFormData theForm, ContainerName, VisFormID, objProgBar
    
    Exit Sub
DBError:
    If clsComLib.ErrCenter() = 1 Then Resume
    clsComLib.SaveErrLog
End Sub

Private Sub SaveFormData(theForm As Form, ContainerName As String, ByVal VisFormID As String, Optional objProgBar As ProgressBar)
    Dim tmpCtrl As Control, ValidCtrl As Boolean
    Dim Seq As Long, aTmp() As String
    Dim i As Integer

    gcnOracle.BeginTrans
    On Error GoTo DBError
    
    gcnOracle.Execute "Delete From 病历所见单 Where 元素ID=" & VisFormID
        
    If Not objProgBar Is Nothing Then objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = theForm.Controls.Count
    i = 0
    For Each tmpCtrl In theForm.Controls
        ValidCtrl = True
        On Error Resume Next
        If UCase(tmpCtrl.Container.Name) <> UCase(ContainerName) Or Not tmpCtrl.Visible Then ValidCtrl = False
        i = i + 1
        On Error GoTo DBError
        If ValidCtrl Then
            Seq = tmpCtrl.TabIndex + 1
            Select Case UCase(tmpCtrl.Name)
                Case "TEXT1"
                    gcnOracle.Execute "ZL_所见单_SAVE(" & VisFormID & "," & Seq & ",'1','" + Replace(tmpCtrl.Text, "'", "''") + "'," & _
                    tmpCtrl.Top & "," & tmpCtrl.Left & "," & tmpCtrl.Width & "," & tmpCtrl.Height & "," & tmpCtrl.Alignment & "," & _
                    0 & ",'',0,'','','')", , adCmdStoredProc
                Case "LINE1"
                    gcnOracle.Execute "ZL_所见单_SAVE(" & VisFormID & "," & Seq & ",'9',''," & _
                    tmpCtrl.Top & "," & tmpCtrl.Left & "," & tmpCtrl.Width & "," & tmpCtrl.Height & ",0," & _
                    0 & ",'',0,'','','')", , adCmdStoredProc
                Case "VISITEM1" '项目ID
                    gcnOracle.Execute "ZL_所见单_SAVE(" & VisFormID & "," & Seq & ",'2','" + Replace(tmpCtrl.Title, "'", "''") + "'," & _
                    tmpCtrl.Top & "," & tmpCtrl.Left & "," & tmpCtrl.Width & "," & tmpCtrl.Height & ",0," & _
                    IIf(tmpCtrl.Enabled, 0, 1) & ",'" & tmpCtrl.ID & "'," & IIf(tmpCtrl.AllowMask, 1, 0) & ",'" & tmpCtrl.ItemType & "','" + tmpCtrl.Unit + "','" + Replace(tmpCtrl.Value, "'", "''") + "')", , adCmdStoredProc
                Case "FRATABLE" '元素ID
                    gcnOracle.Execute "ZL_所见单_SAVE(" & VisFormID & "," & Seq & ",'3',''," & _
                    tmpCtrl.Top & "," & tmpCtrl.Left & "," & tmpCtrl.Width & "," & tmpCtrl.Height & ",0," & _
                    0 & ",'',0,'','','" + VisFormID + "')", , adCmdStoredProc
                    
                    theForm.F1Book1(tmpCtrl.Index).Tag = VisFormID
                    SaveTable theForm.F1Book1(tmpCtrl.Index), Seq
            End Select
            
        End If
        If Not objProgBar Is Nothing Then objProgBar.Value = i
    Next
    gcnOracle.CommitTrans
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Err.Raise Err.Number, "保存所见单"
End Sub

Public Sub ReadForm(theForm As Object, ContainerName As String, ByVal VisFormID As String, Optional FormWidth As Long, Optional FormHeight As Long, Optional objProgBar As ProgressBar)

    Dim tmpCtrl As Control, ValidCtrl As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim Seq As Integer
    Dim iItemLen As Integer, sItemFormat As String, sDefaultValue As String

    On Error GoTo DBError
    If Len(VisFormID) = 0 Then Exit Sub
    
    clsDatabase.OpenRecordset rsTmp, "Select a.*,b.表示法,b.类型,b.长度,b.小数,b.数值域,b.替换域,b.中文名,b.编码 As 所见项编码 From 病历所见单 a,诊治所见项目 b Where a.元素ID=" & VisFormID & " And a.控件号>=0 " + _
        "And a.所见项ID=b.ID(+)", "查询所见单项目"
    Seq = 0
        
    If Not objProgBar Is Nothing Then objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = IIf(rsTmp.EOF, 1, rsTmp.RecordCount)
    Do While Not rsTmp.EOF
        Select Case rsTmp("控件类")
            Case 1
                Load theForm.Text1(theForm.Text1.Count)
                With theForm.Text1(theForm.Text1.Count - 1)
                    .Text = rsTmp("标题")
                    .Top = rsTmp("行"): .Left = rsTmp("列"): .Width = rsTmp("宽"): .Height = rsTmp("高")
                    .Alignment = rsTmp("对齐")
                    .Visible = True
                End With
            Case 9
                Load theForm.Line1(theForm.Line1.Count)
                With theForm.Line1(theForm.Line1.Count - 1)
                    .Top = rsTmp("行"): .Left = rsTmp("列"): .Width = rsTmp("宽"): .Height = rsTmp("高")
                    .Visible = True
                End With
            Case 2
                If Not IsNull(rsTmp("表示法")) Then
                    Load theForm.VisItem1(theForm.VisItem1.Count)
                    With theForm.VisItem1(theForm.VisItem1.Count - 1)
                        sDefaultValue = IIf(IsNull(rsTmp("缺省内容")), "", rsTmp("缺省内容"))
                        Select Case rsTmp("所见项编码")
                            Case "101001" '年(YYYY)
                                iItemLen = 4: sItemFormat = "YYYY"
                            Case "101002" '月(YYYY-MM)
                                iItemLen = 7: sItemFormat = "YYYY-MM"
                            Case "101003" '日(YYYY-MM-DD)
                                iItemLen = 10: sItemFormat = "YYYY-MM-DD"
                            Case "101004" '时间(YYYY-MM-DD HH24:MI:SS)
                                iItemLen = 19: sItemFormat = "YYYY-MM-DD HH:MM:SS"
                            Case "101005" '时间(HH24:MI:SS)
                                iItemLen = 8: sItemFormat = "HH:MM:SS"
                            Case "101006" '时间(HH24:MI)
                                iItemLen = 5: sItemFormat = "HH:MM"
                            Case "101008" '当前日期
                                iItemLen = 10: sItemFormat = "YYYY-MM-DD"
                                sDefaultValue = Format(Date, sItemFormat)
                            Case "101009" '当前时间
                                iItemLen = 19: sItemFormat = "YYYY-MM-DD HH:MM:SS"
                                sDefaultValue = Format(zlDatabase.Currentdate, sItemFormat)
                            Case Else
                                iItemLen = IIf(IsNull(rsTmp("长度")), 10, rsTmp("长度"))
                                sItemFormat = ""
                        End Select
                        .Init IIf(IsNull(rsTmp("标题")), "", rsTmp("标题")), IIf(IsNull(rsTmp("计量单位")), "", rsTmp("计量单位")), rsTmp("表示法"), rsTmp("类型"), iItemLen, IIf(IsNull(rsTmp("小数")), 0, rsTmp("小数")), IIf(IsNull(rsTmp("数值域")), "", rsTmp("数值域")), sDefaultValue, rsTmp("所见项ID"), IIf(IsNull(rsTmp("替换域")), "", IIf(rsTmp("替换域") = 1, rsTmp("中文名"), "")), , , sItemFormat
                        .Left = rsTmp("列"): .Top = rsTmp("行")
                        .Enabled = IIf(rsTmp("不可写") = 0, True, False)
                        .AllowMask = IIf(IsNull(rsTmp("可屏蔽")), False, IIf(rsTmp("可屏蔽") = 0, False, True))
                        .Width = rsTmp("宽"): .Height = rsTmp("高")
                        .TabIndex = Seq: Seq = Seq + 1
                        .Visible = True
                    End With
                End If
            Case 3
                Load theForm.F1Book1(theForm.F1Book1.Count)
                InitTable theForm.F1Book1(theForm.F1Book1.Count - 1)
                
                Load theForm.fraTable(theForm.fraTable.Count)
                Set theForm.F1Book1(theForm.F1Book1.Count - 1).Container = theForm.fraTable(theForm.fraTable.Count - 1)
                With theForm.fraTable(theForm.fraTable.Count - 1)
                    .Top = rsTmp("行"): .Left = rsTmp("列"): .Width = rsTmp("宽"): .Height = rsTmp("高")
                    .TabIndex = Seq: Seq = Seq + 1
                    .Visible = True
                End With
                With theForm.F1Book1(theForm.F1Book1.Count - 1)
                    .Left = 0: .Top = 0
                    .Width = theForm.fraTable(theForm.fraTable.Count - 1).Width
                    .Height = theForm.fraTable(theForm.fraTable.Count - 1).Height
                    .Visible = True
                End With
                ReadTable theForm.F1Book1(theForm.F1Book1.Count - 1), VisFormID, rsTmp("控件号")
        End Select
                    
        If Not objProgBar Is Nothing Then objProgBar.Value = rsTmp.AbsolutePosition
            
        rsTmp.MoveNext
    Loop
    Exit Sub
DBError:
    If clsComLib.ErrCenter() = 1 Then Resume
    clsComLib.SaveErrLog
End Sub
