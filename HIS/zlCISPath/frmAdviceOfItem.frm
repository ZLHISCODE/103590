VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceOfItem 
   BorderStyle     =   0  'None
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   13995
      _cx             =   24686
      _cy             =   7011
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceOfItem.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmAdviceOfItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum CONST_COL
    col期效 = 0
    col内容 = 1
    col总量 = 2
    col单量 = 3
    col频次 = 4
    col用法 = 5
    col嘱托 = 6
    col执行科室 = 7
    col执行性质 = 8
    colID = 9
    col相关ID = 10
    col诊疗项目ID = 11
    col诊疗类别 = 12
    col收费细目ID = 13
    col标本部位 = 14
    col检查方法 = 15
    col执行时间 = 16
    col_开始执行时间 = 17
    col_终止时间 = 18
End Enum

Public Sub ShowAdvice(ByVal bytUseType As Byte, Optional ByVal strSQL As String, Optional ByVal lng路径执行ID As Long, Optional ByVal str医嘱IDs As String)
'功能：路径项目定义时，或路径表中选中一行路径项目时，显示对应的医嘱清单
'参数：
'      bytUseType：     0-路径项目定义时显示医嘱,1-病人路径执行项目（路径表）显示医嘱清单,2-添加或修改路径外项目显示医嘱
'      strSQL：         bytUseType=0时传入，医嘱清单数据源,传入空时清除表格内容
'      lng路径执行ID：  bytUseType=1时传入，病人路径执行项目的ID
'      str医嘱IDs：     bytUseType=2时传入，当前添加的医嘱ID串
    Dim rsTmp As ADODB.Recordset
    Dim blnClear As Boolean
    
    If bytUseType = 0 Then
        If strSQL = "" Then blnClear = True
    ElseIf bytUseType = 1 Then
        If lng路径执行ID = 0 Then blnClear = True
    ElseIf bytUseType = 2 Then
        If str医嘱IDs = "" Then blnClear = True
    End If
    If blnClear Then
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1 '加一空白行
        Exit Sub
    End If
        
    If bytUseType <> 0 Then
        If bytUseType = 1 Then
            strSQL = "Select A.* From 病人医嘱记录 A,病人路径医嘱 B Where B.路径执行ID = [1] And A.ID = B.病人医嘱ID"
        Else
            strSQL = "Select * From 病人医嘱记录 a,Table(f_Num2list([2])) b Where a.ID = b.Column_value"
        End If
    End If
    
    '传入SQL：左边NULL字段右边(+)CBO下面读不出来
    strSQL = "Select " & IIf(bytUseType = 2, "/*+ rule*/", "") & "A.ID,A.相关ID,A.序号," & IIf(bytUseType = 0, "A.期效", "A.医嘱期效") & " as 期效,A.诊疗项目ID,A.医嘱内容," & _
        " A.单次用量,A.执行频次,A.医生嘱托,Nvl(C.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'<院外执行>')) as 执行科室," & _
        " A.执行性质, " & IIf(bytUseType = 0, "A.时间方案", "A.执行时间方案") & " as 时间方案,Nvl(B.类别,'*') as 诊疗类别,Nvl(D.名称||Decode(D.规格,NULL,NULL,' '||D.规格),B.名称) as 名称," & _
        " B.计算单位,A.标本部位,A.检查方法,A.总给予量,D.计算单位 as 总量单位,D.ID as 收费细目ID," & _
        " Nvl(B.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) As 撤档时间" & _
        IIf(bytUseType = 0, "", ",To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 开始时间,To_Char(A.执行终止时间,'YYYY-MM-DD HH24:MI') as 终止时间") & _
        IIf(bytUseType = 1, " ,a.医嘱状态", "") & _
        " From (" & strSQL & ") A,诊疗项目目录 B,部门表 C,收费项目目录 D" & _
        " Where Nvl(A.诊疗项目ID,-1)=B.ID(+) And Nvl(A.执行科室ID,-1)=C.ID(+) And Nvl(A.收费细目ID,-1)=D.ID(+)" & _
        " Order by A.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ShowAdvice", lng路径执行ID, str医嘱IDs)
    Call LoadAdvice(rsTmp, bytUseType)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetAdviceTitle(Optional ByVal lngRows As Long = 5) As String
'功能：获取医嘱多行医嘱内容的组合字符串(最多lngRows行)
    Dim strItem As String, i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                If UBound(Split(strItem, "、")) + 1 > lngRows Then
                    strItem = strItem & "......"
                    Exit For
                Else
                    strItem = strItem & "、" & .TextMatrix(i, col内容)
                End If
            End If
        Next
    End With
    GetAdviceTitle = Mid(strItem, 2)
End Function

Private Sub LoadAdvice(ByRef rsTmp As ADODB.Recordset, ByVal bytUseType As Byte)
'功能：显示路径项目对应的医嘱内容
    Dim strTmp As String
    Dim str中药 As String, str煎法 As String
    Dim str麻醉 As String, str标本 As String
    Dim strFilter As String
    Dim i As Long, j As Long
    
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '清除表格内容
        .Rows = .FixedRows + rsTmp.RecordCount
        If bytUseType = 0 Then      '项目医嘱定义
            .ColHidden(col_开始执行时间) = True
            .ColHidden(col_终止时间) = True
        ElseIf bytUseType = 2 Then  '添加路径外项目
            .ColHidden(col_终止时间) = True
        End If
        
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, col期效) = IIf(Nvl(rsTmp!期效, 0) = 0, "长期", "临时")
            .TextMatrix(i, col内容) = Nvl(rsTmp!医嘱内容, Nvl(rsTmp!名称))
            If bytUseType = 1 Then
                If rsTmp!医嘱状态 = 4 Then
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '灰色
                    .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                End If
            End If
            
            .TextMatrix(i, col标本部位) = Nvl(rsTmp!标本部位) '检验标本
            .TextMatrix(i, col检查方法) = Nvl(rsTmp!检查方法)
            .TextMatrix(i, col单量) = FormatEx(Nvl(rsTmp!单次用量), 4)
            If Not IsNull(rsTmp!单次用量) Then
                If rsTmp!诊疗类别 = "4" Then
                    .TextMatrix(i, col单量) = .TextMatrix(i, col单量) & Nvl(rsTmp!总量单位)
                Else
                    .TextMatrix(i, col单量) = .TextMatrix(i, col单量) & Nvl(rsTmp!计算单位)
                End If
            End If
            If .TextMatrix(i, col期效) = "临时" Then
                If Not IsNull(rsTmp!总给予量) Then
                    .TextMatrix(i, col总量) = FormatEx(Nvl(rsTmp!总给予量), 4)
                    If Not IsNull(rsTmp!总量单位) Then
                        .TextMatrix(i, col总量) = .TextMatrix(i, col总量) & Nvl(rsTmp!总量单位)
                    ElseIf InStr(",4,5,6,7,", rsTmp!诊疗类别) = 0 Then
                        .TextMatrix(i, col总量) = .TextMatrix(i, col总量) & Nvl(rsTmp!计算单位)
                    End If
                End If
            End If
            .TextMatrix(i, col频次) = Nvl(rsTmp!执行频次)
            .TextMatrix(i, col嘱托) = Nvl(rsTmp!医生嘱托)
            .TextMatrix(i, col执行时间) = Nvl(rsTmp!时间方案)
            .TextMatrix(i, col执行科室) = Nvl(rsTmp!执行科室)
            .Cell(flexcpData, i, col执行性质) = Nvl(rsTmp!执行性质, 0)
            .TextMatrix(i, colID) = rsTmp!ID
            .TextMatrix(i, col相关ID) = "" & rsTmp!相关ID
            .TextMatrix(i, col诊疗项目ID) = "" & rsTmp!诊疗项目ID
            .TextMatrix(i, col收费细目ID) = "" & rsTmp!收费细目ID
            .TextMatrix(i, col诊疗类别) = rsTmp!诊疗类别
            If Format(rsTmp!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF&
            End If
            
            If bytUseType <> 0 Then
                .TextMatrix(i, col_开始执行时间) = "" & rsTmp!开始时间
                .TextMatrix(i, col_终止时间) = "" & rsTmp!终止时间
            End If
            rsTmp.MoveNext
        Next
        
        '再处理一些附加行的隐藏,及相关内容的显示
        For i = 1 To .Rows - 1
            '给药途径
            If .TextMatrix(i, col诊疗类别) = "E" And Val(.TextMatrix(i, col相关ID)) = 0 _
                And Val(.TextMatrix(i - 1, col相关ID)) = Val(.TextMatrix(i, colID)) _
                And InStr(",5,6,", .TextMatrix(i - 1, col诊疗类别)) > 0 Then
                .RowHidden(i) = True
                '显示给药途径
                For j = i - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, col相关ID)) = Val(.TextMatrix(i, colID)) Then
                        .TextMatrix(j, col用法) = .TextMatrix(i, col内容)
                                                    
                        '显示成药的执行性质
                        If Val(.Cell(flexcpData, j, col执行性质)) = 5 And Val(.Cell(flexcpData, i, col执行性质)) <> 5 Then
                            .TextMatrix(j, col执行性质) = "自备药"
                        ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                            .TextMatrix(j, col执行性质) = "离院带药"
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
            
            '输血途径
            If .TextMatrix(i, col诊疗类别) = "E" And .TextMatrix(i - 1, col诊疗类别) = "K" _
                And Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(i - 1, colID)) Then
                .RowHidden(i) = True
                .TextMatrix(i - 1, col用法) = .TextMatrix(i, col内容)
                .TextMatrix(i - 1, col内容) = .TextMatrix(i - 1, col内容) & "(" & .TextMatrix(i, col内容) & ")"
            End If
            
            '中药配方和检验组合
            If .TextMatrix(i, col诊疗类别) = "E" And Val(.TextMatrix(i, col相关ID)) = 0 _
                And Val(.TextMatrix(i - 1, col相关ID)) = Val(.TextMatrix(i, colID)) _
                And InStr(",7,E,C,", .TextMatrix(i - 1, col诊疗类别)) > 0 Then
                
                str中药 = "": str煎法 = "": str标本 = "": strTmp = ""
                j = .FindRow(CStr(Val(.TextMatrix(i, colID))), , col相关ID)
                
                '中药及检验的执行科室
                .TextMatrix(i, col执行科室) = .TextMatrix(j, col执行科室)
                
                '显示中药配方执行性质:以药品为准判断
                If .TextMatrix(i - 1, col诊疗类别) <> "C" Then
                    If Val(.Cell(flexcpData, j, col执行性质)) = 5 And Val(.Cell(flexcpData, i, col执行性质)) <> 5 Then
                        .TextMatrix(i, col执行性质) = "自备药"
                    ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                        .TextMatrix(i, col执行性质) = "离院带药"
                    End If
                End If

                For j = j To i - 1
                    .RowHidden(j) = j <> i
                    If .TextMatrix(j, col诊疗类别) = "7" Then
                        str中药 = str中药 & "," & RTrim(.TextMatrix(j, col内容) & _
                            " " & .TextMatrix(j, col单量) & " " & .TextMatrix(j, col嘱托))
                    ElseIf .TextMatrix(j, col诊疗类别) = "C" Then
                        strTmp = strTmp & "," & .TextMatrix(j, col内容)
                        str标本 = .TextMatrix(j, col标本部位) '取第一个检验项目的标本
                    ElseIf .TextMatrix(j, col诊疗类别) = "E" And Val(.TextMatrix(j, col相关ID)) <> 0 Then
                        str煎法 = .TextMatrix(j, col内容)
                    End If
                Next
                
                .TextMatrix(i, col用法) = .TextMatrix(i, col内容) '显示中药用法或检验采集方法
                
                If .TextMatrix(i - 1, col诊疗类别) = "C" Then
                    .TextMatrix(i, col内容) = Mid(strTmp, 2) & IIf(str标本 <> "", "(" & str标本 & ")", "")
                Else
                    .TextMatrix(i, col内容) = "中药配方," & .TextMatrix(i, col频次) & "," & _
                        str煎法 & "," & .TextMatrix(i, col内容) & ":" & Mid(str中药, 2)
                    .TextMatrix(i, col总量) = .TextMatrix(i, col总量) & "付"
                End If
            End If
            
            '检查组合
            If .TextMatrix(i, col诊疗类别) = "D" And Val(.TextMatrix(i, col相关ID)) = 0 Then
                str标本 = "": str煎法 = "": strTmp = ""
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, col相关ID)) = Val(.TextMatrix(i, colID)) Then
                        .RowHidden(j) = True
                        If .TextMatrix(j, col标本部位) <> "" _
                            And Val(.TextMatrix(j, col诊疗项目ID)) = Val(.TextMatrix(i, col诊疗项目ID)) Then '相同的项目ID才是新方式
                            If .TextMatrix(j, col标本部位) <> strTmp And strTmp <> "" Then
                                str标本 = str标本 & "," & strTmp & IIf(str煎法 <> "", "(" & Mid(str煎法, 2) & ")", "")
                                str煎法 = ""
                            End If
                            If .TextMatrix(j, col检查方法) <> "" Then
                                str煎法 = str煎法 & "," & .TextMatrix(j, col检查方法)
                            End If
                            
                            strTmp = .TextMatrix(j, col标本部位)
                        End If
                    Else
                        Exit For
                    End If
                Next
                If strTmp <> "" Then
                    str标本 = str标本 & "," & strTmp & IIf(str煎法 <> "", "(" & Mid(str煎法, 2) & ")", "")
                End If
                If str标本 <> "" Then '以前的检查方式时不显示详细医嘱内容
                    .TextMatrix(i, col内容) = .TextMatrix(i, col内容) & ":" & Mid(str标本, 2)
                End If
            End If
            
            '手术项目
            If .TextMatrix(i, col诊疗类别) = "F" And Val(.TextMatrix(i, col相关ID)) = 0 Then
                strTmp = "": str麻醉 = ""
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, col相关ID)) = Val(.TextMatrix(i, colID)) Then
                        .RowHidden(j) = True
                        If .TextMatrix(j, col诊疗类别) = "F" Then
                            strTmp = strTmp & "," & .TextMatrix(j, col内容)
                        ElseIf .TextMatrix(j, col诊疗类别) = "G" Then
                            str麻醉 = .TextMatrix(j, col内容)
                        End If
                    Else
                        Exit For
                    End If
                Next
                If strTmp <> "" Or str麻醉 <> "" Then
                    If str麻醉 <> "" Then
                        .TextMatrix(i, col内容) = "在 " & str麻醉 & " 下行 " & .TextMatrix(i, col内容)
                    Else
                        .TextMatrix(i, col内容) = "行 " & .TextMatrix(i, col内容)
                    End If
                    If strTmp <> "" Then
                        .TextMatrix(i, col内容) = .TextMatrix(i, col内容) & " 及 " & Mid(strTmp, 2)
                    End If
                End If
            End If
        Next
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows: .Col = .FixedCols
            .AutoSize col内容
        Else
            .Rows = .FixedRows + 1
        End If
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub Form_Resize()
    vsAdvice.Top = Me.ScaleTop + 60
    vsAdvice.Left = Me.ScaleLeft
    vsAdvice.Height = Me.ScaleHeight - 60
    vsAdvice.Width = Me.ScaleWidth
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsAdvice.FixedRows And NewCol >= vsAdvice.FixedCols Then
        If NewRow <> OldRow Then
            vsAdvice.ForeColorSel = vsAdvice.CellForeColor
        End If
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col内容 Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '擦除一并给药相关行列的边线及内容
        lngLeft = col期效: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col频次: lngRight = col用法
        End If
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col执行时间: lngRight = col_终止时间
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) And Me.ActiveControl Is vsAdvice Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    
    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub
