VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSchemeSelect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
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
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "注意：浅灰色背景的行是没有库存的药品或卫材。"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   4995
         TabIndex        =   8
         Top             =   75
         Visible         =   0   'False
         Width           =   3960
      End
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
      Cols            =   28
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
Private mint来源 As Integer 'IN:1-门诊,2-住院,3-门诊和住院
Private mstr序号 As String 'Out
Private Enum COL成套方案
    col选择 = 0
    col期效 = 1
    col内容 = 2
    col总量 = 3
    col总量单位 = 4
    col单量 = 5
    col单量单位 = 6
    col天数 = 7
    col频次 = 8
    col用法 = 9
    col嘱托 = 10
    col执行时间 = 11
    col执行科室 = 12
    col执行性质 = 13
    col序号 = 14
    col相关 = 15
    col项目ID = 16
    col类别 = 17
    col标本部位 = 18
    col检查方法 = 19
    col是否适用 = 20
    col提示 = 21
    col毒理分类 = 22
    col价值分类 = 23
    col执行标记 = 24
    col性别 = 25
    col单独应用 = 26
    col操作类型 = 27
End Enum
Private mstr性别 As String
Private mlng病人科室id As Long
Private mbln麻醉类权限 As Boolean
Private mbln毒性类权限 As Boolean
Private mbln精神类权限 As Boolean
Private mbln贵重类权限 As Boolean
Private mlng病人性质 As Long '0-普通住院病人,1-门诊留观病人,2-住院留观病人

Public Function ShowMe(frmParent As Object, ByVal lng成套ID As Long, ByVal int来源 As Integer, Optional ByVal lng病人科室ID As Long, _
        Optional ByVal str性别 As String, Optional ByVal lng病人性质 As Long) As String
'返回：选择的项目序号
'     "+序号1,序号2,...":表示包括这些序号
'     "-序号1,序号2,...":表示排开这些序号
'     "*"表示选择所有,""表示取消操合
    mstr序号 = ""
    mlng成套ID = lng成套ID
    mint来源 = int来源
    mlng病人科室id = lng病人科室ID
    mstr性别 = str性别
    mlng病人性质 = lng病人性质
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

Private Sub cmdALL_Click()
    Dim i As Long
    Dim lngEnd As Long
    
    With vsScheme
        For i = .FixedRows To .Rows - 1
            If CheckCanSelGroup(i, False) Then
                '以前的检查医嘱不允许保存为成套方案
                If .TextMatrix(i, col类别) = "D" Then
                    If Val(.TextMatrix(i, col相关)) = 0 Then
                        If Not CheckIsOldAdvice(i) Then
                            Call SelGroup(i, 1, lngEnd)
                        End If
                    Else
                        '主项行已处理
                    End If
                Else
                    Call SelGroup(i, 1, lngEnd)
                End If
            End If
            If i < lngEnd Then i = lngEnd
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
        Call cmdALL_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    lblTitle.Caption = Sys.RowValue("诊疗项目目录", mlng成套ID, "名称")
    lblRemark.Visible = mlng病人科室id <> 0
    '执行天数
    If mint来源 = 3 Then
        vsScheme.ColHidden(col天数) = True
    Else
        vsScheme.ColHidden(col天数) = Val(zlDatabase.GetPara("医嘱执行天数", glngSys, IIF(mint来源 = 1, p门诊医嘱下达, p住院医嘱下达))) = 0
    End If
    
    If mint来源 = 1 Then
        mbln麻醉类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达麻醉药嘱;") = 0
        mbln毒性类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达毒性药嘱;") = 0
        mbln精神类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达精神药嘱;") = 0
        mbln贵重类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达贵重药嘱;") = 0
    ElseIf mint来源 = 2 Then
        mbln麻醉类权限 = InStr(GetTsPrivs(p住院医嘱下达), ";下达麻醉药嘱;") = 0
        mbln毒性类权限 = InStr(GetTsPrivs(p住院医嘱下达), ";下达毒性药嘱;") = 0
        mbln精神类权限 = InStr(GetTsPrivs(p住院医嘱下达), ";下达精神药嘱;") = 0
        mbln贵重类权限 = InStr(GetTsPrivs(p住院医嘱下达), ";下达贵重药嘱;") = 0
    ElseIf mint来源 = 3 Then
        mbln麻醉类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达麻醉药嘱;") = 0 And InStr(GetTsPrivs(p住院医嘱下达), ";下达麻醉药嘱;") = 0
        mbln毒性类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达毒性药嘱;") = 0 And InStr(GetTsPrivs(p住院医嘱下达), ";下达毒性药嘱;") = 0
        mbln精神类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达精神药嘱;") = 0 And InStr(GetTsPrivs(p住院医嘱下达), ";下达精神药嘱;") = 0
        mbln贵重类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达贵重药嘱;") = 0 And InStr(GetTsPrivs(p住院医嘱下达), ";下达贵重药嘱;") = 0
    End If
    
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
    With vsScheme
        If Col <> col选择 Then
            Cancel = True
        ElseIf Val(.TextMatrix(.Row, col序号)) = 0 Then
            Cancel = True
        Else
            '以前的检查医嘱不允许选择
            If CheckIsOldAdvice(Row) Then
                MsgBox "该检查医嘱是系统升级以前下达的，与现有方式不兼容，不能保存为成套方案。", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If .TextMatrix(Row, col选择) <> 0 Then
                Call SelGroup(Row, 0)
            Else
                If CheckCanSelGroup(Row, True) Then
                    Call SelGroup(Row, -1)
                End If
            End If
            '已经进行判断后选择，不需触发AfterEdit事件
            Cancel = True
        End If
    End With
End Sub

Private Function CheckIsOldAdvice(ByVal lngRow As Long) As Boolean
'功能：检查指定行的检查医嘱是否老方式
'参数：lngRow=检查医嘱可见行
    Dim lngIdx As Long

    With vsScheme
        If .TextMatrix(lngRow, col类别) = "D" Then
            lngIdx = .FindRow(CStr(.TextMatrix(lngRow, col序号)), lngRow + 1, col相关)
            If lngIdx = -1 Then
                'CheckIsOldAdvice = True '以前的单部位检查
            ElseIf Val(.TextMatrix(lngIdx, col项目ID)) <> Val(.TextMatrix(lngRow, col项目ID)) Then
                CheckIsOldAdvice = True '以前的多部位项目检查
            End If
        End If
    End With
End Function

Private Sub vsScheme_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsScheme
        '擦除一并给药相关行列的边线及内容
        lngLeft = col期效: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col天数: lngRight = col用法
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
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsScheme_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsScheme
        If KeyAscii = 32 Then
            If .Col <> col选择 Then
                KeyAscii = 0
                If .TextMatrix(.Row, col选择) = 0 Then
                    If CheckCanSelGroup(.Row, True) Then
                        Call SelGroup(.Row, -1)
                    End If
                Else
                    Call SelGroup(.Row, 0)
                End If
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


Private Function ShowScheme() As Boolean
'功能：读取并显示数据库中的成套方案内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim str中药 As String, str煎法 As String
    Dim str标本 As String, str麻醉 As String
    Dim i As Long, j As Long, lngEnd As Long
    Dim strDepartments As String
    Dim lngSel As Long
    Dim str服务对象 As String
    
    If mlng病人性质 = 1 Then
        str服务对象 = ",1,2,"
    Else
        str服务对象 = "," & mint来源 & ","
    End If

    '门诊不支持自由医嘱调入
    strSQL = "Select (Select Count(1) From 诊疗适用科室 Where 项目ID=b.ID) as 适用科室数,g.科室id as 适用科室ID,A.序号,A.相关序号,A.期效,A.诊疗项目ID,A.医嘱内容,A.总给予量,A.单次用量,A.天数," & _
             " A.执行频次,A.医生嘱托,Nvl(C.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'-')) as 执行科室,Nvl(b.适用性别, 0) As 性别," & _
             " A.执行性质,A.执行标记,A.时间方案,Nvl(B.类别,'*') as 类别,Nvl(E.名称||Decode(E.规格,NULL,NULL,' '||E.规格),B.名称) as 名称,B.计算单位," & _
             " E.计算单位 as 散装单位,A.标本部位,A.检查方法,B.服务对象,B.单独应用,B.操作类型,B.撤档时间,E.服务对象 as 收费服务,E.撤档时间 as 收费撤档,E.ID as 收费项目ID,Nvl(f.跟踪在用,0) as 跟踪在用," & _
             Decode(mint来源, 1, "D.门诊包装 as 包装系数,D.门诊单位 as 包装单位", _
                    2, "D.住院包装 as 包装系数,D.住院单位 as 包装单位", 3, "1 as 包装系数,E.计算单位 as 包装单位") & _
                    ",(Select f_List2str(Cast(Collect(j.名称) As t_Strlist))" & vbNewLine & _
                    "         From 诊疗项目组合 H, 诊疗项目目录 J, 收费项目目录 k" & vbNewLine & _
                    "         Where h.诊疗项目id = j.Id And k.id(+)=h.收费细目ID And a.诊疗组合id = h.诊疗组合id And a.序号 = h.相关序号 And NVL(k.撤档时间,j.撤档时间) <> To_Date('3000/1/1', 'yyyy/mm/dd') And" & vbNewLine & _
                    "               (j.类别 in ('C', '7') Or j.类别 = 'E' And Nvl(j.执行分类,0) = 0 And j.操作类型 = '3')) As 提示 ,h.毒理分类,h.价值分类 " & _
                    " From 诊疗项目组合 A,诊疗项目目录 B,部门表 C,药品规格 D,收费项目目录 E,材料特性 F,诊疗适用科室 G,药品特性 H" & _
                    " Where A.诊疗项目ID=B.ID" & IIF(mint来源 = 1, "", "(+)") & " And A.执行科室ID=C.ID(+) And e.id=f.材料id(+)  And h.药名ID(+)=b.ID " & _
                    " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
                    " And A.收费细目ID=D.药品ID(+) And A.收费细目ID=E.ID(+) And b.id=g.项目id(+) And g.科室ID(+)=[2] " & _
                    IIF(mint来源 = 1, " And A.期效=1", "") & " And A.诊疗组合ID=[1] Order by A.序号"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng成套ID, mlng病人科室id)
    With vsScheme
        .Redraw = flexRDNone
        .Rows = .FixedRows    '清除表格内容
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            lngSel = IIF(Val(Sys.RowValue("诊疗项目目录", mlng成套ID, "执行分类") & "") = 1, -1, 0)
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col选择) = lngSel
                .TextMatrix(i, col期效) = IIF(NVL(rsTmp!期效, 0) = 0, "长期", "临时")
                .TextMatrix(i, col内容) = NVL(rsTmp!医嘱内容, NVL(rsTmp!名称))
                .Cell(flexcpData, i, col内容) = .TextMatrix(i, col内容)
                .TextMatrix(i, col标本部位) = NVL(rsTmp!标本部位)
                .Cell(flexcpData, i, col标本部位) = .TextMatrix(i, col标本部位)
                .TextMatrix(i, col检查方法) = NVL(rsTmp!检查方法)
                .Cell(flexcpData, i, col检查方法) = .TextMatrix(i, col检查方法)
                .RowData(i) = 0
                '总量
                If InStr(",5,6,", rsTmp!类别) > 0 Then
                    '成药临嘱有总量,以零售单位存放,包装单位显示
                    If Not IsNull(rsTmp!总给予量) And Not IsNull(rsTmp!包装系数) Then
                        .TextMatrix(i, col总量) = FormatEx(rsTmp!总给予量 / rsTmp!包装系数, 5)
                    End If
                    If NVL(rsTmp!期效, 0) = 1 Then
                        .TextMatrix(i, col总量单位) = NVL(rsTmp!包装单位)
                    End If
                Else
                    '其它情况有中药和其它临嘱
                    If Not IsNull(rsTmp!总给予量) Then
                        .TextMatrix(i, col总量) = rsTmp!总给予量
                    End If
                    If rsTmp!类别 = "E" And NVL(rsTmp!相关序号, 0) = 0 _
                       And Val(.TextMatrix(i - 1, col相关)) = rsTmp!序号 _
                       And InStr(",7,E,", .TextMatrix(i - 1, col类别)) > 0 Then
                        .TextMatrix(i, col总量单位) = "付"    '中药配方总量单位为"付"
                    ElseIf NVL(rsTmp!期效, 0) = 1 Then
                        If rsTmp!类别 = "4" Then
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!散装单位)
                        Else
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!计算单位)
                        End If
                    End If
                End If

                '单量
                .TextMatrix(i, col单量) = FormatEx(NVL(rsTmp!单次用量), 4)
                If Not IsNull(rsTmp!单次用量) Then
                    If rsTmp!类别 = "4" Then
                        .TextMatrix(i, col单量单位) = NVL(rsTmp!散装单位)
                    Else
                        .TextMatrix(i, col单量单位) = NVL(rsTmp!计算单位)
                    End If
                End If
                .TextMatrix(i, col天数) = NVL(rsTmp!天数)
                .TextMatrix(i, col频次) = NVL(rsTmp!执行频次)
                .TextMatrix(i, col嘱托) = NVL(rsTmp!医生嘱托)
                .TextMatrix(i, col执行时间) = NVL(rsTmp!时间方案)
                .TextMatrix(i, col执行科室) = NVL(rsTmp!执行科室)
                .Cell(flexcpData, i, col执行性质) = NVL(rsTmp!执行性质, 0)
                .TextMatrix(i, col序号) = rsTmp!序号
                .TextMatrix(i, col相关) = NVL(rsTmp!相关序号)
                .TextMatrix(i, col项目ID) = NVL(rsTmp!诊疗项目ID)
                .TextMatrix(i, col类别) = rsTmp!类别
                .TextMatrix(i, col毒理分类) = NVL(rsTmp!毒理分类)
                .TextMatrix(i, col价值分类) = NVL(rsTmp!价值分类)
                .TextMatrix(i, col执行标记) = rsTmp!执行标记 & ""
                .TextMatrix(i, col性别) = Decode(Val(rsTmp!性别), 0, "未知", 1, "男", 2, "女")
                .TextMatrix(i, col单独应用) = rsTmp!单独应用 & ""
                .TextMatrix(i, col操作类型) = rsTmp!操作类型 & ""
                
                
                '判断非院外执行药品和跟踪在用卫材是否有库存
                If mlng病人科室id <> 0 And InStr(",4,5,6,7,", rsTmp!类别 & "") > 0 Then
                    strDepartments = ""
                    If Val(rsTmp!执行性质 & "") <> 5 And InStr(",5,6,7,", rsTmp!类别 & "") > 0 And Val(rsTmp!收费项目ID & "") <> 0 Then
                        strDepartments = Get可用药房IDs(rsTmp!类别 & "", NVL(rsTmp!诊疗项目ID), Val(rsTmp!收费项目ID & ""), mlng病人科室id, mint来源)
                    ElseIf Val(rsTmp!跟踪在用) = 1 And rsTmp!类别 & "" = "4" Then
                        strDepartments = Get可用发料部门IDs(Val(rsTmp!收费项目ID & ""), mlng病人科室id, mint来源)
                    End If
                    '判断库存是否大于总量
                    If strDepartments <> "" Then
                        If GetStock(Val(rsTmp!收费项目ID & ""), , mint来源, strDepartments, CDbl(Val(.TextMatrix(i, col总量)))) = 0 Then
                            .TextMatrix(i, col选择) = 0
                            .Cell(flexcpBackColor, i, 0, i, col是否适用) = &H8000000F
                            If InStr(",5,6,7,", rsTmp!类别 & "") > 0 Then
                                .Cell(flexcpData, i, col是否适用) = 1
                            End If
                        End If
                    Else
                        .TextMatrix(i, col选择) = 0
                        .Cell(flexcpBackColor, i, 0, i, col是否适用) = &H8000000F
                    End If
                End If
                '如果不属于指定的适用科室，则设置为深灰色
                If Val(rsTmp!适用科室数 & "") > 0 And rsTmp!适用科室ID & "" = "" Then
                    .TextMatrix(i, col选择) = 0
                    .TextMatrix(i, col是否适用) = "1"
                    '                    .Cell(flexcpBackColor, i, 0, i, col是否适用) = &HC0C0C0
                End If
                '如果提示不为空，则是中药配方中有停用中药或煎法
                .TextMatrix(i, col提示) = rsTmp!提示 & ""
                If rsTmp!提示 & "" <> "" Then
                    .TextMatrix(i, col选择) = 0
                End If
                '检查权限
                If mbln麻醉类权限 And .TextMatrix(i, col毒理分类) = "麻醉药" Or _
                   mbln毒性类权限 And .TextMatrix(i, col毒理分类) = "毒性药" Or _
                   mbln精神类权限 And (.TextMatrix(i, col毒理分类) = "精神I类") Or _
                   mbln贵重类权限 And (.TextMatrix(i, col价值分类) = "贵重" Or .TextMatrix(i, col价值分类) = "昂贵") Then
                    .TextMatrix(i, col选择) = 0
                End If

                '输血医嘱检查，必须中级及以上专业技术职务的医师才允许下达
                If .TextMatrix(i, col类别) = "K" And gbln输血申请中级以上 Then
                    If UserInfo.专业技术职务 <> "主治医师" And UserInfo.专业技术职务 <> "主任医师" And UserInfo.专业技术职务 <> "副主任医师" Then
                        .TextMatrix(i, col选择) = 0
                    End If
                End If

                '标记包含得有撤档或不服务的项目
                If Not IsNull(rsTmp!诊疗项目ID) Then
                    If Not (IsNull(rsTmp!撤档时间) Or Format(NVL(rsTmp!撤档时间), "yyyy-MM-dd") = "3000-01-01") Then
                        .RowData(i) = 1
                    ElseIf Not (NVL(rsTmp!服务对象, 0) = 3 Or InStr(str服务对象, "," & NVL(rsTmp!服务对象, 0) & ",") > 0) Or mint来源 = 3 Then
                        .RowData(i) = 1
                    ElseIf Not IsNull(rsTmp!包装单位) Or rsTmp!类别 & "" = "4" Then
                        '对药品,同时要判断到收费项目目录
                        If Not (IsNull(rsTmp!收费撤档) Or Format(NVL(rsTmp!收费撤档), "yyyy-MM-dd") = "3000-01-01") Then
                            .RowData(i) = 1
                        ElseIf Not (NVL(rsTmp!收费服务, 0) = 3 Or NVL(rsTmp!收费服务, 0) = mint来源) Or mint来源 = 3 Then
                            .RowData(i) = 1
                        End If
                    End If
                End If

                rsTmp.MoveNext
            Next

            '如果一组药品中有一个为不选择(没有库存)，则把同组的其他药品和给药途径也设置为不选择。
            For i = 1 To .Rows - 1
                If .TextMatrix(i, col选择) = 0 Then
                    For j = 1 To .Rows - 1
                        '药品
                        If .TextMatrix(i, col相关) <> "" And (.TextMatrix(j, col相关) = .TextMatrix(i, col相关) Or .TextMatrix(j, col序号) = .TextMatrix(i, col相关)) Or _
                           .TextMatrix(i, col相关) = "" And .TextMatrix(i, col序号) = .TextMatrix(j, col相关) Then
                           If .TextMatrix(i, col类别) = "7" And .TextMatrix(j, col相关) = "" Then
                                .Cell(flexcpBackColor, j, 0, j, col是否适用) = &H8000000F
                           End If
                           .TextMatrix(j, col选择) = 0
                        End If
                    Next
                End If
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
                            .TextMatrix(j, col用法) = .TextMatrix(i, col内容) & .TextMatrix(i, col嘱托)

                            '显示成药执行性质
                            If Val(.Cell(flexcpData, j, col执行性质)) = 5 And Val(.Cell(flexcpData, i, col执行性质)) <> 5 Then
                                .TextMatrix(j, col执行性质) = IIF(Val(.TextMatrix(j, col执行标记)) = 2, "不取药", "自备药")
                            ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                                .TextMatrix(j, col执行性质) = "离院带药"
                            Else
                                .TextMatrix(j, col执行性质) = IIF(Val(.TextMatrix(j, col执行标记)) = 1, "自取药", "正常")
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If

                '输血途径
                If .TextMatrix(i, col类别) = "E" And .TextMatrix(i - 1, col类别) = "K" _
                   And Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(i - 1, col序号)) Then
                    .RowHidden(i) = True
                    .TextMatrix(i - 1, col用法) = .TextMatrix(i, col内容)
                    .TextMatrix(i - 1, col内容) = .TextMatrix(i - 1, col内容) & "(" & .TextMatrix(i, col内容) & ")"
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
                            .TextMatrix(i, col执行性质) = IIF(Val(.TextMatrix(j, col执行标记)) = 2, "不取药", "自备药")
                        ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                            .TextMatrix(i, col执行性质) = "离院带药"
                        Else
                            .TextMatrix(i, col执行性质) = IIF(Val(.TextMatrix(j, col执行标记)) = 1, "自取药", "正常")
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
                            str标本 = .TextMatrix(j, col标本部位)    '取第一个检验项目的标本
                        ElseIf .TextMatrix(j, col类别) = "E" And Val(.TextMatrix(j, col相关)) <> 0 Then
                            str煎法 = .TextMatrix(j, col内容)
                        End If
                    Next

                    .TextMatrix(i, col用法) = .TextMatrix(i, col内容)    '显示中药用法或检验采集方法

                    If .TextMatrix(i - 1, col类别) = "C" Then
                        .TextMatrix(i, col内容) = Mid(strTmp, 2) & IIF(str标本 <> "", "(" & str标本 & ")", "")
                    Else
                        .TextMatrix(i, col内容) = "中药配方," & .TextMatrix(i, col频次) & "," & _
                                                str煎法 & "," & .TextMatrix(i, col内容) & ":" & Mid(str中药, 2)
                    End If
                End If

                '检查组合
                If .TextMatrix(i, col类别) = "D" And Val(.TextMatrix(i, col相关)) = 0 Then
                    str标本 = "": str煎法 = "": strTmp = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col相关)) = Val(.TextMatrix(i, col序号)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, col标本部位) <> "" _
                               And Val(.TextMatrix(j, col项目ID)) = Val(.TextMatrix(i, col项目ID)) Then    '相同的项目ID才是新方式
                                If .TextMatrix(j, col标本部位) <> strTmp And strTmp <> "" Then
                                    str标本 = str标本 & "," & strTmp & IIF(str煎法 <> "", "(" & Mid(str煎法, 2) & ")", "")
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
                        str标本 = str标本 & "," & strTmp & IIF(str煎法 <> "", "(" & Mid(str煎法, 2) & ")", "")
                    End If
                    If str标本 <> "" Then
                        .TextMatrix(i, col内容) = .TextMatrix(i, col内容) & ":" & Mid(str标本, 2)
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
                    Call SelGroup(i, 0, lngEnd)
                End If
                If i < lngEnd Then i = lngEnd
            Next
        End If
        .ColHidden(col期效) = mint来源 = 1
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

Private Function CheckCanSelRow(ByVal lngRow As Long) As String
'功能:验证指定行是否可以选择
    Dim lngCol As Long
    Dim strContent As String
    
    With vsScheme
        If .TextMatrix(lngRow, col类别) = "D" Then
            strContent = "[" & Trim(.Cell(flexcpData, lngRow, col标本部位)) & "]" & Trim(.Cell(flexcpData, lngRow, col检查方法))
            If strContent <> "[]" Then
                strContent = Chr(34) & strContent & Chr(34)
            Else
                strContent = Chr(34) & .Cell(flexcpData, lngRow, col内容) & Chr(34)
            End If
        Else
            strContent = Chr(34) & .Cell(flexcpData, lngRow, col内容) & Chr(34)
        End If
        If .RowData(lngRow) = 1 Then
            CheckCanSelRow = strContent & "(已撤档或不服务当前科室)": Exit Function
        End If
        
        If .TextMatrix(lngRow, col是否适用) = "1" Then
            CheckCanSelRow = strContent & "(不适用于当前科室)": Exit Function
        End If
        
        If InStr("未知" & mstr性别, .TextMatrix(lngRow, col性别)) = 0 Then
            CheckCanSelRow = strContent & "(不适用于当前病人性别)": Exit Function
        End If
        
        If mbln麻醉类权限 And .TextMatrix(lngRow, col毒理分类) = "麻醉药" Then
            CheckCanSelRow = strContent & "(无麻醉类药品权限)": Exit Function
        End If
        
        If mbln毒性类权限 And .TextMatrix(lngRow, col毒理分类) = "毒性药" Then
            CheckCanSelRow = strContent & "(无毒性药品权限)": Exit Function
        End If
        
        If mbln精神类权限 And (.TextMatrix(lngRow, col毒理分类) = "精神I类") Then
            CheckCanSelRow = strContent & "(无精神类药品权限)": Exit Function
        End If
        
        If mbln贵重类权限 And (.TextMatrix(lngRow, col价值分类) = "贵重" Or .TextMatrix(lngRow, col价值分类) = "昂贵") Then
            CheckCanSelRow = strContent & "(无贵重类药品权限)": Exit Function
        End If
        
        '输血医嘱检查，必须中级及以上专业技术职务的医师才允许下达
        If .TextMatrix(lngRow, col类别) = "K" And gbln输血申请中级以上 Then
            If UserInfo.专业技术职务 <> "主治医师" And UserInfo.专业技术职务 <> "主任医师" And UserInfo.专业技术职务 <> "副主任医师" Then
                CheckCanSelRow = Trim(.Cell(flexcpText, lngRow, col内容) & "") & "(无中级及以上专业技术职务)": Exit Function
            End If
        End If
    End With
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, Optional lngRow相关 As Long)
'功能:获取一组医嘱的起止位置，同时获取父医嘱行号
'参数:
'   lngRow 当前行
'返回:
'   lngBegin 起始行
'   lngEnd 终止行
'   lngRow相关 父医嘱行

    Dim i As Long, lng相关 As Long

    With vsScheme
        If .TextMatrix(lngRow, col类别) = "" Then '自由录入
            lngRow相关 = lngRow: lngBegin = lngRow: lngEnd = lngRow
            Exit Sub
        End If
        '获取相关序号
        If Val(.TextMatrix(lngRow, col相关)) <> 0 Then
            lng相关 = Val(.TextMatrix(lngRow, col相关))
            lngRow相关 = .FindRow(lng相关, , col序号, , True)
            If lngRow相关 = -1 Then
                lngRow相关 = lngRow
            End If
        Else
            lng相关 = Val(.TextMatrix(lngRow, col序号)): lngRow相关 = lngRow
        End If
        
        lngBegin = lngRow相关: lngEnd = lngRow相关
        
        For i = lngRow相关 - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, col相关)) = lng相关 Then
                lngBegin = i
            Else
                Exit For
            End If
        Next

        For i = lngRow相关 + 1 To .Rows - 1
            If Val(.TextMatrix(i, col相关)) = lng相关 Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Function CheckCanSelGroup(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = True) As Boolean
'功能：判断本组医嘱是否可以选择
'参数：
'   lngRow 当前行
'   blnAsk 是否进行提示或询问（全选时，不进行询问)
    Dim i As Long, strResult As String
    Dim lngBegin As Long, lngEnd As Long, lngRow相关 As Long
    Dim bln配方 As Boolean, bln检验 As Boolean, blnCanSel As Boolean
    Dim strMsg As String
    Dim blnMedicineAdvice As Boolean
    
    With vsScheme
        '获取本组医嘱信息
        Call GetRowScope(lngRow, lngBegin, lngEnd, lngRow相关)
        
         '检查是否有医嘱的诊疗项目未勾选“可以单独应用”，未勾选的不允许复制。
        If lngBegin = lngEnd Then
            If Val(.TextMatrix(lngRow, col单独应用)) = 0 And Val(.TextMatrix(lngRow, col项目ID)) <> 0 Then
                If blnAsk Then
                    MsgBox "医嘱“" & .TextMatrix(lngRow, col内容) & "”对应的诊疗项目不能单独应用，不可以被选择。如有疑问，请联系管理员！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        Else
            For i = lngBegin To lngEnd
                If InStr(",5,6,7,", .TextMatrix(i, col类别)) > 0 Then
                    blnMedicineAdvice = True
                End If
            Next
            If Not blnMedicineAdvice Then
                For i = lngBegin To lngEnd
                    If Not (.TextMatrix(i, col类别) = "G" Or (.TextMatrix(i, col类别) = "E" And InStr(",2,3,4,6,7,8,", .TextMatrix(i, col操作类型)) > 0)) Then
                        If Val(.TextMatrix(i, col单独应用)) = 0 And Val(.TextMatrix(i, col项目ID)) <> 0 Then
                            If blnAsk Then
                                MsgBox "医嘱“" & .TextMatrix(i, col内容) & "”对应的诊疗项目不能单独应用，不可以被选择。如有疑问，请联系管理员！", vbInformation, gstrSysName
                            End If
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
        
        '在启用参数 指定药房时限制库存 的情况下，不允许下达库存不足的医嘱
        If gblnStock Then
            For i = lngBegin To lngEnd
                If Val(.Cell(flexcpData, i, col是否适用)) = 1 Then
                    If Val(.TextMatrix(lngBegin, col类别)) = 7 Then
                        strMsg = strMsg & "," & .TextMatrix(i, col内容)
                    Else
                        If blnAsk Then
                            MsgBox "该药品库存不足,系统限制了不允许下达库存不足的药品，不能被选择！", vbInformation, gstrSysName
                        End If
                        Exit Function
                    End If
                End If
            Next
        End If
        If strMsg <> "" Then
            MsgBox "该配方中存在库存不足的药品(" & Mid(strMsg, 2) & ")。", vbInformation, gstrSysName
        End If
            
        strMsg = CheckCanSelRow(lngRow相关)
        If strMsg <> "" Then '父医嘱或单条医嘱检查
            If blnAsk Then
                MsgBox "该医嘱中" & vbNewLine & strMsg & vbNewLine & "无效,不能被选择", vbInformation, gstrSysName
            End If
            Exit Function
        Else
            If lngBegin <> lngEnd Then
                If .TextMatrix(lngRow相关, col类别) = "E" Then
                    If lngRow相关 - 2 >= lngBegin Then
                        If .TextMatrix(lngRow相关 - 2, col类别) = "7" And .TextMatrix(lngRow相关 - 1, col类别) = "E" Then '中药配方的剪法检查
                            strMsg = CheckCanSelRow(lngRow相关 - 1)
                            If strMsg <> "" Then
                                If blnAsk Then
                                    MsgBox "该中药配方中煎法:" & vbNewLine & strMsg & vbNewLine & "无效,不能被选择", vbInformation, gstrSysName
                                End If
                                Exit Function
                            Else
                                bln配方 = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        strMsg = ""
        '子医嘱全部检查
        If lngBegin <> lngEnd Then
            For i = lngBegin To lngEnd
                If .TextMatrix(lngRow相关, col类别) = "F" Then blnCanSel = True '手术医嘱父医嘱可用就可选
                If Not (i = lngRow相关 Or bln配方 And i = lngRow相关 - 1) Then
                    strResult = CheckCanSelRow(i)
                    If .TextMatrix(i, col类别) = "C" Then bln检验 = True
                    If strResult <> "" Then
                        strMsg = IIF(strMsg = "", "", strMsg & "、" & vbNewLine) & strResult
                    Else
                        If bln配方 Then  '中药配方含一味中药可用就可选（煎法以及用法前面已经判断），其余类型只要一个子医嘱可用就可选
                            If .TextMatrix(i, col类别) = "7" Then
                                blnCanSel = True
                            End If
                        Else
                            blnCanSel = True
                        End If
                    End If
                End If
            Next
        Else
            blnCanSel = True '单条医嘱检查在父医嘱时已经检查
        End If
        
        If Not blnCanSel Then strMsg = ""
        '中药配方未提取的药品信息
        If .TextMatrix(lngRow相关, col提示) <> "" Then
            If (bln检验 Or bln配方) And blnCanSel Then
                If bln配方 Then strMsg = strMsg & IIF(strMsg <> "", "、" & vbNewLine, "") & vsScheme.TextMatrix(lngRow相关, col提示) & "(已停用或没有可用规格)"
            ElseIf bln配方 Then
                blnCanSel = False
                strMsg = "该中药配方中所有中药已经被停用或没有可用规格,不能被选择"
            ElseIf bln检验 Then
                blnCanSel = False
                strMsg = "该检验组合中所有检验项目已经被停用,不能被选择"
            End If
        End If
        If Not blnCanSel And strMsg = "" Then strMsg = "该医嘱中不存在有效项目,不能被选择"

        If blnCanSel Then
            If strMsg <> "" Then
                If blnAsk Then
                    If MsgBox(IIF(InStr(1, strMsg, "、") > 0, "该医嘱中:" & vbNewLine & strMsg & vbNewLine & "无效,这些项目", "该医嘱中:" & vbNewLine & strMsg & vbNewLine & "无效,该项目") & "不会被选择,是否选择该医嘱？", _
                        vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        CheckCanSelGroup = True
                    End If
                End If
            Else
                CheckCanSelGroup = True
            End If
        Else
            If blnAsk Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End With
End Function

Private Sub SelGroup(ByVal lngRow As Long, ByVal int选择 As Integer, Optional ByRef lngEnd As Long)
'功能:根据情况选择该组医嘱
'参数：
'   lngRow 当前行
'   lngEnd 本组医嘱最后一行
'   int选择 选择结果 -1,检查选择（可选的选择，不可选不选择),0不选择,1，全选不检查
    Dim lngBegin As Long
    Dim i As Long
    
    With vsScheme
    
        '获取本组医嘱信息
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        '选择或取消选择
        If int选择 = -1 Then 'checkCanSelGroup(i,true)=true后调用
            For i = lngBegin To lngEnd
                If CheckCanSelRow(i) = "" Then
                    .TextMatrix(i, col选择) = int选择
                Else
                    .TextMatrix(i, col选择) = 0
                End If
            Next
        Else 'checkCanSelGroup(i,false)=true 后调用或取消选择后使用
            int选择 = int选择 * -1
            For i = lngBegin To lngEnd
                .TextMatrix(i, col选择) = int选择
            Next
        End If
    End With
End Sub

