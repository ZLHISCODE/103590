VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSchemeSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "成套方案选择"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   Icon            =   "frmSchemeSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCommand 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   5685
      Width           =   9390
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7695
         TabIndex        =   2
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6585
         TabIndex        =   1
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&R)"
         Height          =   350
         Left            =   1755
         TabIndex        =   4
         ToolTipText     =   "Ctrl+R"
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   645
         TabIndex        =   3
         ToolTipText     =   "Ctrl+A"
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9420
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   9420
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "成套方案名称"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   210
         TabIndex        =   6
         Top             =   75
         Width           =   1080
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsScheme 
      Align           =   1  'Align Top
      Height          =   5385
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   9420
      _cx             =   16616
      _cy             =   9499
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSchemeSelect.frx":058A
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
      Editable        =   2
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
      FrozenCols      =   1
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
Attribute VB_Name = "frmSchemeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng成套ID As Long 'IN
Private mint来源 As Integer 'IN:1-门诊,2-住院
Private mstr序号 As String 'Out
Private Enum COL成套方案
    col选择 = 0
    col期效 = 1
    col内容 = 2
    col总量 = 3
    col总量单位 = 4
    col单量 = 5
    col单量单位 = 6
    col频次 = 7
    col用法 = 8
    col嘱托 = 9
    col执行时间 = 10
    col执行科室 = 11
    col执行性质 = 12
    col序号 = 13
    col相关 = 14
    col项目ID = 15
    col类别 = 16
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng成套ID As Long, ByVal int来源 As Integer) As String
'返回：选择的项目序号
'     "+序号1,序号2,...":表示包括这些序号
'     "-序号1,序号2,...":表示排开这些序号
'     "*"表示选择所有,""表示取消操合
    mstr序号 = ""
    mlng成套ID = lng成套ID
    mint来源 = int来源
    Me.Show 1, frmParent
    ShowMe = mstr序号
End Function

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsScheme
        If .TextMatrix(lngRow, col类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub cmdAll_Click()
    Dim i As Long
    With vsScheme
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col序号)) <> 0 And RowCanSelect(i) = 0 Then
                .TextMatrix(i, col选择) = -1
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    With vsScheme
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, col选择) = 0
        Next
    End With
End Sub

Private Sub cmdOK_Click()
    Dim strSel As String, strUnSel As String, i As Long
    
    With vsScheme
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, col选择)) <> 0 Then
                strSel = strSel & "," & Val(.TextMatrix(i, col序号))
            Else
                strUnSel = strUnSel & "," & Val(.TextMatrix(i, col序号))
            End If
        Next
        strSel = Mid(strSel, 2)
        strUnSel = Mid(strUnSel, 2)

        If strSel = "" Then
            MsgBox "请从成套方案中选择需要的项目内容。", vbInformation, gstrSysName
            vsScheme.SetFocus: Exit Sub
        End If
        If strUnSel = "" Then
            mstr序号 = "*"
        Else
            If UBound(Split(strSel, ",")) > UBound(Split(strUnSel, ",")) Then
                mstr序号 = "-" & strUnSel
            Else
                mstr序号 = "+" & strSel
            End If
        End If
    End With
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    lblTitle.Caption = Get项目名称(mlng成套ID)
    Call ShowScheme
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsScheme.Height = Me.ScaleHeight - picTitle.Height - fraCommand.Height
    fraCommand.Left = 0
    fraCommand.Top = vsScheme.Top + vsScheme.Height
    fraCommand.Width = Me.ScaleWidth
    
    If Me.ScaleWidth - cmdCancel.Width - cmdAll.Left - cmdOK.Width < cmdClear.Left + cmdClear.Width + 300 Then
        cmdOK.Left = cmdClear.Left + cmdClear.Width + 300
        cmdCancel.Left = cmdOK.Left + cmdOK.Width
    Else
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdAll.Left
        cmdOK.Left = cmdCancel.Left - cmdOK.Width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vsScheme_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsScheme.FixedRows And NewCol >= vsScheme.FixedCols Then
        If NewRow <> OldRow Then
            vsScheme.ForeColorSel = vsScheme.CellForeColor
        End If
    End If
End Sub

Private Sub vsScheme_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col内容 Then
        vsScheme.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsScheme.TextMatrix(vsScheme.FixedRows - 1, Col) & "A")
        If vsScheme.ColWidth(Col) < lngW Then
            vsScheme.ColWidth(Col) = lngW
        ElseIf vsScheme.ColWidth(Col) > vsScheme.Width * 0.5 Then
            vsScheme.ColWidth(Col) = vsScheme.Width * 0.5
        End If
    End If
End Sub

Private Sub vsScheme_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col选择 Then Cancel = True
End Sub

Private Sub vsScheme_DblClick()
    Call vsScheme_KeyPress(32)
End Sub

Private Sub vsScheme_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    If Col <> col选择 Then
        Cancel = True
    ElseIf Val(vsScheme.TextMatrix(vsScheme.Row, col序号)) = 0 Then
        Cancel = True
    Else
        i = RowCanSelect(Row)
        If i > 0 Then
            Cancel = True
            MsgBox "因为""" & vsScheme.TextMatrix(i, col内容) & """已撤档或服务对象不匹配，该医嘱不能被选择。", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub vsScheme_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsScheme
        '擦除一并给药相关行列的边线及内容
        lngLeft = col期效: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col频次: lngRight = col用法
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
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
        If Between(Row, .Row, .RowSel) And Me.ActiveControl Is vsScheme Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsScheme_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col选择 Then Call RowSelectSame(Row)
End Sub

Private Sub vsScheme_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsScheme
        If KeyAscii = 32 Then
            If .Col <> col选择 Then
                KeyAscii = 0
                If Val(.TextMatrix(.Row, col序号)) = 0 Then Exit Sub
                
                i = RowCanSelect(.Row)
                If i > 0 And Val(.TextMatrix(.Row, col选择)) = 0 Then
                    MsgBox "因为""" & .TextMatrix(i, col内容) & """已撤档或服务对象不匹配，该医嘱不能被选择。", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                .TextMatrix(.Row, col选择) = IIF(Val(.TextMatrix(.Row, col选择)) = 0, -1, 0)
                Call RowSelectSame(.Row)
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long)
'功能：根据指定行(可能为任意行)的选择状态,将相关医嘱一并选择
    Dim i As Long
    
    With vsScheme
        If Val(.TextMatrix(lngRow, col相关)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) _
                    Or Val(.TextMatrix(i, col序号)) = Val(.TextMatrix(lngRow, col相关)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) _
                    Or Val(.TextMatrix(i, col序号)) = Val(.TextMatrix(lngRow, col相关)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col序号)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col序号)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Function RowCanSelect(ByVal lngRow As Long) As Long
'功能：判断指定行的(相关)医嘱可否选择
'返回：如果可以选择，返回0,否则返回行号
    Dim i As Long
    
    With vsScheme
        If .RowData(lngRow) = 1 Then RowCanSelect = lngRow: Exit Function
        
        If Val(.TextMatrix(lngRow, col相关)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) _
                    Or Val(.TextMatrix(i, col序号)) = Val(.TextMatrix(lngRow, col相关)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) _
                    Or Val(.TextMatrix(i, col序号)) = Val(.TextMatrix(lngRow, col相关)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col序号)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col序号)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Function ShowScheme() As Boolean
'功能：读取并显示数据库中的成套方案内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim str中药 As String, str煎法 As String
    Dim str标本 As String, str麻醉 As String
    Dim i As Long, j As Long, str来源 As String
    
    str来源 = IIF(mint来源 = 1, "门诊", "住院")
    strSQL = "Select A.序号,A.相关序号,A.期效,A.诊疗项目ID,A.总给予量,A.单次用量," & _
        " A.执行频次,A.医生嘱托,Nvl(C.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'<院外执行>')) as 执行科室," & _
        " A.执行性质,A.时间方案,B.类别,B.名称,B.计算单位,A.标本部位," & _
        " B.服务对象,B.撤档时间,E.服务对象 as 收费服务,E.撤档时间 as 收费撤档," & _
        " D." & str来源 & "单位 as 包装单位,D." & str来源 & "包装 as 包装系数" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,部门表 C,药品规格 D,收费项目目录 E" & _
        " Where A.诊疗项目ID=B.ID And A.执行科室ID=C.ID(+)" & _
        " And A.收费细目ID=D.药品ID(+) And A.收费细目ID=E.ID(+)" & _
        IIF(mint来源 = 1, " And A.期效=1", "") & " And A.诊疗组合ID=[1]" & _
        " Order by A.序号"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng成套ID)
    With vsScheme
        .Redraw = flexRDNone
        .Rows = .FixedRows '清除表格内容
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col选择) = -1
                .TextMatrix(i, col期效) = IIF(Nvl(rsTmp!期效, 0) = 0, "长期", "临时")
                .TextMatrix(i, col内容) = rsTmp!名称
                .Cell(flexcpData, i, col内容) = Nvl(rsTmp!标本部位) '检验标本
                
                '总量
                If InStr(",5,6,", rsTmp!类别) > 0 Then
                    '成药临嘱有总量,以零售单位存放,包装单位显示
                    If Not IsNull(rsTmp!总给予量) And Not IsNull(rsTmp!包装系数) Then
                        .TextMatrix(i, col总量) = FormatEx(rsTmp!总给予量 / rsTmp!包装系数, 5)
                    End If
                    If Nvl(rsTmp!期效, 0) = 1 Then
                        .TextMatrix(i, col总量单位) = Nvl(rsTmp!包装单位)
                    End If
                Else
                    '其它情况有中药和其它临嘱
                    If Not IsNull(rsTmp!总给予量) Then
                        .TextMatrix(i, col总量) = rsTmp!总给予量
                    End If
                    If rsTmp!类别 = "E" And Nvl(rsTmp!相关序号, 0) = 0 _
                        And Val(.TextMatrix(i - 1, col相关)) = rsTmp!序号 _
                        And InStr(",7,E,", .TextMatrix(i - 1, col类别)) > 0 Then
                        .TextMatrix(i, col总量单位) = "付" '中药配方总量单位为"付"
                    ElseIf Nvl(rsTmp!期效, 0) = 1 Then
                        .TextMatrix(i, col总量单位) = Nvl(rsTmp!计算单位)
                    End If
                End If
                                
                '单量
                .TextMatrix(i, col单量) = FormatEx(Nvl(rsTmp!单次用量), 4)
                If Not IsNull(rsTmp!单次用量) Then
                    .TextMatrix(i, col单量单位) = Nvl(rsTmp!计算单位)
                End If
                
                .TextMatrix(i, col频次) = Nvl(rsTmp!执行频次)
                .TextMatrix(i, col嘱托) = Nvl(rsTmp!医生嘱托)
                .TextMatrix(i, col执行时间) = Nvl(rsTmp!时间方案)
                .TextMatrix(i, col执行科室) = Nvl(rsTmp!执行科室)
                .Cell(flexcpData, i, col执行性质) = Nvl(rsTmp!执行性质, 0)
                .TextMatrix(i, col序号) = rsTmp!序号
                .TextMatrix(i, col相关) = Nvl(rsTmp!相关序号)
                .TextMatrix(i, col项目ID) = rsTmp!诊疗项目ID
                .TextMatrix(i, col类别) = rsTmp!类别
                
                '标记包含得有撤档或不服务的项目
                If Not (IsNull(rsTmp!撤档时间) Or Format(Nvl(rsTmp!撤档时间), "yyyy-MM-dd") = "3000-01-01") Then
                    .RowData(i) = 1
                ElseIf Not (Nvl(rsTmp!服务对象, 0) = 3 Or Nvl(rsTmp!服务对象, 0) = mint来源) Then
                    .RowData(i) = 1
                ElseIf Not IsNull(rsTmp!包装单位) Then
                    '对药品,同时要判断到收费项目目录
                    If Not (IsNull(rsTmp!收费撤档) Or Format(Nvl(rsTmp!收费撤档), "yyyy-MM-dd") = "3000-01-01") Then
                        .RowData(i) = 1
                    ElseIf Not (Nvl(rsTmp!收费服务, 0) = 3 Or Nvl(rsTmp!收费服务, 0) = mint来源) Then
                        .RowData(i) = 1
                    End If
                End If
                
                rsTmp.MoveNext
            Next
            
            '再处理一些附加行的隐藏,及相关内容的显示
            For i = 1 To .Rows - 1
                '给药途径
                If .TextMatrix(i, col类别) = "E" And Val(.TextMatrix(i, col相关)) = 0 _
                    And Val(.TextMatrix(i - 1, col相关)) = Val(.TextMatrix(i, col序号)) _
                    And InStr(",5,6,", .TextMatrix(i - 1, col类别)) > 0 Then
                    .RowHidden(i) = True
                    '显示给药途径
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关)) = Val(.TextMatrix(i, col序号)) Then
                            .TextMatrix(j, col用法) = .TextMatrix(i, col内容)
                            
                            '显示成药执行性质
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
                
                '中药配方和检验组合
                If .TextMatrix(i, col类别) = "E" And Val(.TextMatrix(i, col相关)) = 0 _
                    And Val(.TextMatrix(i - 1, col相关)) = Val(.TextMatrix(i, col序号)) _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col类别)) > 0 Then
                    
                    str中药 = "": str煎法 = "": str标本 = "": strTmp = ""
                    j = .FindRow(CStr(Val(.TextMatrix(i, col序号))), , col相关)
                    
                    '中药及检验的执行科室
                    .TextMatrix(i, col执行科室) = .TextMatrix(j, col执行科室)
                    
                    '显示中药配方执行性质:以药品为准判断
                    If .TextMatrix(i - 1, col类别) <> "C" Then
                        If Val(.Cell(flexcpData, j, col执行性质)) = 5 And Val(.Cell(flexcpData, i, col执行性质)) <> 5 Then
                            .TextMatrix(i, col执行性质) = "自备药"
                        ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                            .TextMatrix(i, col执行性质) = "离院带药"
                        End If
                    End If
                    
                    For j = j To i - 1
                        .RowHidden(j) = j <> i
                        If .TextMatrix(j, col类别) = "7" Then
                            str中药 = str中药 & "," & RTrim(.TextMatrix(j, col内容) & _
                                " " & .TextMatrix(j, col单量) & .TextMatrix(j, col单量单位) & _
                                " " & .TextMatrix(j, col嘱托))
                        ElseIf .TextMatrix(j, col类别) = "C" Then
                            strTmp = strTmp & "," & .TextMatrix(j, col内容)
                            str标本 = .Cell(flexcpData, j, col内容) '取第一个检验项目的标本
                        ElseIf .TextMatrix(j, col类别) = "E" And Val(.TextMatrix(j, col相关)) <> 0 Then
                            str煎法 = .TextMatrix(j, col内容)
                        End If
                    Next
                    
                    .TextMatrix(i, col用法) = .TextMatrix(i, col内容) '显示中药用法或检验采集方法
                    
                    If .TextMatrix(i - 1, col类别) = "C" Then
                        .TextMatrix(i, col内容) = Mid(strTmp, 2) & IIF(str标本 <> "", "(" & str标本 & ")", "")
                    Else
                        .TextMatrix(i, col内容) = "中药配方," & .TextMatrix(i, col频次) & "," & _
                            str煎法 & "," & .TextMatrix(i, col内容) & ":" & Mid(str中药, 2)
                    End If
                End If
                
                '检查组合
                If .TextMatrix(i, col类别) = "D" And Val(.TextMatrix(i, col相关)) = 0 Then
                    strTmp = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col相关)) = Val(.TextMatrix(i, col序号)) Then
                            .RowHidden(j) = True
                            strTmp = strTmp & "," & .Cell(flexcpData, j, col内容)
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Then
                        .TextMatrix(i, col内容) = .TextMatrix(i, col内容) & "(" & Mid(strTmp, 2) & ")"
                    End If
                End If
                
                '手术项目
                If .TextMatrix(i, col类别) = "F" And Val(.TextMatrix(i, col相关)) = 0 Then
                    strTmp = "": str麻醉 = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col相关)) = Val(.TextMatrix(i, col序号)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, col类别) = "F" Then
                                strTmp = strTmp & "," & .TextMatrix(j, col内容)
                            ElseIf .TextMatrix(j, col类别) = "G" Then
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
            
            '作了标记的行的相关行一并标记,并取消选择
            For i = 1 To .Rows - 1
                If .RowData(i) = 1 Then
                    .TextMatrix(i, col选择) = 0
                    Call RowSelectSame(i)
                End If
            Next
        End If
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col内容
        .Redraw = flexRDDirect
    End With
    ShowScheme = True
    Exit Function
errH:
    vsScheme.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
