VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDrugListEditEx 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "中药配方编辑"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4860
   Icon            =   "frmDrugListEditEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   4860
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboData 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   615
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   6000
      Width           =   2000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000004&
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3960
      Picture         =   "frmDrugListEditEx.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "取消(Esc)"
      Top             =   6000
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   2880
      Picture         =   "frmDrugListEditEx.frx":6DDC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "确认(F2)"
      Top             =   6000
      Width           =   885
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      _cx             =   8546
      _cy             =   10504
      Appearance      =   0
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
      MouseIcon       =   "frmDrugListEditEx.frx":7366
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   10000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugListEditEx.frx":7C40
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox pictmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1920
         ScaleHeight     =   240
         ScaleWidth      =   480
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Label lblData 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "煎法"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   6060
      Width           =   390
   End
End
Attribute VB_Name = "frmDrugListEditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum COLDurg
    COL_诊疗项目ID = 1
    COL_收费细目ID = 2
    col_中药 = 3
    COL_单量 = 4
    COL_单位 = 5
    COL_脚注 = 6
End Enum

Private Const GRD_UNEDITCELL_COLOR = &H8000000B  '未编辑的单元格颜色：灰蓝色


Private mblnReturn As Boolean
Private mblnOK As Boolean
Private mstrLike As String
Private mint简码 As Integer
Private mlng煎法ID As Long
Private mstrData As String  '格式:[配方数据]中药名称<Data>诊疗项目ID<Data>收费细目ID<Data>单量<Data>脚注<Data>单位


Public Function ShowEdit(frmParent As Object, ByRef strData As String, ByRef lng煎法 As Long) As Boolean
'功能：用药清单中药编辑器
'参数：vsTmp 传递用药清单界面的表格控件

    On Error Resume Next
    mlng煎法ID = 0
    mstrData = ""
    mblnOK = False

    mstrData = strData
    mlng煎法ID = lng煎法
    
    Me.Show 1, frmParent
    strData = mstrData
    lng煎法 = mlng煎法ID
    ShowEdit = mblnOK
    On Error GoTo 0
End Function

Private Sub LoadData()
    Dim arrTime As Variant, arrTmp As Variant
    Dim i As Long
    With vsAdvice
        If mstrData = "" Then
            .Rows = 1
            .Rows = vsAdvice.Rows + 1
        Else
             .Redraw = flexRDNone
             .Rows = .FixedRows
             arrTime = Split(mstrData, "[配方数据]")
            For i = 1 To UBound(arrTime)
                .Rows = .Rows + 1
                arrTmp = Split(arrTime(i), "<Data>")
                .TextMatrix(.Rows - 1, col_中药) = arrTmp(0)
                .TextMatrix(.Rows - 1, COL_诊疗项目ID) = arrTmp(1)
                .TextMatrix(.Rows - 1, COL_收费细目ID) = arrTmp(2)
                .TextMatrix(.Rows - 1, COL_单量) = arrTmp(3)
                .TextMatrix(.Rows - 1, COL_脚注) = arrTmp(4)
                .TextMatrix(.Rows - 1, COL_单位) = arrTmp(5)
            Next
             .Redraw = flexRDDirect
        End If
        .Row = .Rows - 1: .Col = col_中药
        .ShowCell .Rows - 1, col_中药
        .Cell(flexcpBackColor, .FixedRows, COL_单位, .Rows - 1, COL_单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
    End With
End Sub


Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long

    strHead = "诊疗项目ID;收费细目ID;中药,2000,1;单量,850,4;单位,850,4;脚注,950,1"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionFree
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .BackColorSel = &H404040

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        .Editable = flexEDKbdMouse
    End With
End Sub

Public Function AdviceCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsAdvice
        If .ColHidden(lngCol) Then Exit Function
        '必须先输入用药内容
        If lngCol = COL_单位 Then Exit Function
        If .TextMatrix(lngRow, col_中药) = "" Then
            If lngCol > col_中药 Then Exit Function
        End If
    End With
    AdviceCellEditable = True
End Function


Private Sub EnterNextCellAdvice()
    Dim i As Long, j As Long

    With vsAdvice
        '从下一单元开始循环搜索
        If .Row < .FixedRows Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, COL_单位, .Rows - 1, COL_单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .ShowCell .Rows - 1, col_中药
        End If
        For i = .Row To .Rows - 1
            For j = IIF(i = .Row, .Col + 1, col_中药) To COL_脚注
                If AdviceCellEditable(i, j) Then Exit For
            Next
            If j <= COL_脚注 Then Exit For
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > COL_脚注 And .TextMatrix(.Rows - 1, col_中药) <> "" Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, COL_单位, .Rows - 1, COL_单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .ShowCell .Rows - 1, col_中药
        End If
    End With
End Sub

Private Function checkDrug()
    Dim i As Long, j As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col_中药) <> "" Then
                '当配方不为空时检查中药煎法
                If cboData.ListIndex = -1 Then
                    MsgBox "请确定中药配方的煎法。", vbInformation, gstrSysName
                    cboData.SetFocus: Exit Function
                End If
                
                
                If Val(.TextMatrix(i, COL_单量)) <= 0 Then
                    MsgBox "中药配方的单量为必填项,请录入。", vbInformation, gstrSysName
                    .SetFocus
                    .Row = i: .Col = COL_单量: Call vsAdvice.ShowCell(.Row, .Col)
                    Exit Function
                End If
                
                If i <> .Rows - 1 Then  '检查是否存在相同中药配方
                    For j = .Rows - 1 To i + 1 Step -1
                        If .TextMatrix(j, col_中药) <> "" Then
                            If .TextMatrix(j, col_中药) & "|" & .TextMatrix(j, COL_诊疗项目ID) & "|" & .TextMatrix(j, COL_收费细目ID) = .TextMatrix(i, col_中药) & "|" & .TextMatrix(i, COL_诊疗项目ID) & "|" & .TextMatrix(i, COL_收费细目ID) Then
                                .SetFocus
                                MsgBox "发现两条重复的用药清单,请检查。", vbInformation, gstrSysName
                                .Row = j: .Col = col_中药: Call vsAdvice.ShowCell(.Row, .Col)
                                Exit Function
                            End If
                        End If
                    Next
                End If
            End If
        Next
        checkDrug = True
    End With
End Function


Private Sub cmdOK_Click()
    Dim i As Long
    Dim strTmp As String
    With vsAdvice
        If .Rows <= 2 And .TextMatrix(.Rows - 1, col_中药) = "" And mstrData <> "" Then
           If MsgBox("请确认是否清空中药配方数据？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
           End If
           mstrData = ""
           mlng煎法ID = 0
           mblnOK = True
           Unload Me
        Else
            If checkDrug Then
                mlng煎法ID = Val(cboData.ItemData(cboData.ListIndex))
                For i = 1 To vsAdvice.Rows - 1
                    If .TextMatrix(i, col_中药) <> "" Then
                        strTmp = strTmp & "[配方数据]" & .TextMatrix(i, col_中药) & "<Data>" & Val(.TextMatrix(i, COL_诊疗项目ID)) & "<Data>" & Val(.TextMatrix(i, COL_收费细目ID)) & "<Data>" & FormatEx(NVL(.TextMatrix(i, COL_单量)), 5) & "<Data>" & .TextMatrix(i, COL_脚注) & "<Data>" & .TextMatrix(i, COL_单位)
                    End If
                Next
                mstrData = strTmp
                mblnOK = True
                Unload Me
            End If
        End If
    End With
End Sub


Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    If (Not AdviceCellEditable(NewRow, NewCol)) Then
        vsAdvice.FocusRect = flexFocusLight
    Else
        vsAdvice.FocusRect = flexFocusSolid
    End If
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsAdvice
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not AdviceCellEditable(Row, Col) Then
        Cancel = True
    End If
End Sub


Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    With vsAdvice
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            mblnReturn = True
            Call EnterNextCellAdvice
        Else
            If .ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsAdvice
        If Not KeyAscii = vbKeyReturn Then
            If Col = COL_单量 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            End If
            mblnReturn = False
        Else
            mblnReturn = True
        End If
    End With
End Sub


Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsAdvice
        If KeyCode = vbKeyF4 Then
            If .ComboList = "..." Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Row < 1 Then Exit Sub
            If MsgBox("确实要删除该行中药配方吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                .RemoveItem .Row
                If .Rows = 1 Then
                    .Rows = .Rows + 1
                    .Cell(flexcpBackColor, .FixedRows, COL_单位, .Rows - 1, COL_单位) = GRD_UNEDITCELL_COLOR      '灰蓝色
                    .Row = .Rows - 1: .Col = col_中药
                    .ShowCell .Rows - 1, col_中药
                End If
            Else
                Exit Sub
            End If
                
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsAdvice_KeyPress(KeyCode)
        End If
    End With
End Sub



Private Sub vsAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'vsAdvice_ValidateEdit事件
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strLike As String
    Dim strInput As String
    Dim lngMax As Long

    On Error GoTo errH
   With vsAdvice
        strLike = mstrLike
        If Len(.EditText) < 2 Then strLike = "" '优化
        Select Case Col
            Case col_中药
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, Row, col_中药)
                    If mblnReturn Then Call EnterNextCellAdvice
                ElseIf .EditText = .Cell(flexcpData, Row, col_中药) Then
                    If mblnReturn Then Call EnterNextCellAdvice
                Else
                    strInput = " And (A.编码 Like [1] And E.码类=[3]" & _
                        " Or E.名称 Like [2] And E.码类=[3] Or E.简码 Like [2] And E.码类 IN([3],3))"
                
                    If IsNumeric(.EditText) Then
                        '1X.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                        If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.编码 Like [1] And E.码类=[3] Or E.简码 Like [2] And E.码类=3)"
                    ElseIf zlCommFun.IsCharAlpha(.EditText) Then
                        'X1.输入全是字母时只匹配简码
                        If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And E.简码 Like [2] And E.码类=[3]"
                    ElseIf zlCommFun.IsCharChinese(.EditText) Then
                        '包含汉字,则只匹配名称a
                        strInput = " And E.名称 Like [2] And E.码类=[3]"
                    End If
                    
                    strInput = IIF(.EditText = "*", "", strInput)
                    strSQL = "Select distinct a.Id, b.Id As 收费细目id, a.名称, b.规格, a.计算单位" & _
                    " From 诊疗项目目录 A, 收费项目目录 B, 药品规格 C, 药品特性 D,诊疗项目别名 E " & _
                    " Where c.药品id= b.Id(+) And a.Id =c.药名id(+) And c.药名id = d.药名id(+) And A.ID=E.诊疗项目ID(+) And a.类别 ='7' and (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & strInput
                    
                    vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药品目录", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, vsAdvice.RowHeight(Row), blnCancel, False, True, UCase(.EditText) & "%", strLike & UCase(.EditText) & "%", mint简码 + 1)
                    
                    If rsTmp Is Nothing Then
                        If Not blnCancel Then
                            MsgBox "未找到可用的中草药，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .EditText = .Cell(flexcpData, Row, col_中药)
                        End If
                        Exit Sub
                    Else
                        .EditText = rsTmp!名称 & ""
                        .TextMatrix(Row, col_中药) = rsTmp!名称 & ""
                        .Cell(flexcpData, Row, col_中药) = .TextMatrix(Row, col_中药)
                        .TextMatrix(Row, COL_诊疗项目ID) = Val(rsTmp!ID & "")
                        .TextMatrix(Row, COL_收费细目ID) = Val(rsTmp!收费细目ID & "")
                        .TextMatrix(Row, COL_单位) = IIF(rsTmp!计算单位 & "" = "", "g", rsTmp!计算单位 & "")
                    End If
                End If
            Case COL_单量
                lngMax = 10
            Case COL_脚注
                lngMax = 100
        End Select
        
        If LenB(StrConv(.EditText, vbFromUnicode)) > lngMax And lngMax <> 0 Then
            MsgBox "不能超过" & lngMax & "个字符的长度。", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        
        mblnReturn = False
   End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitAdviceTable
    Call Get煎法
    Call LoadData
    '输入匹配
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    '简码匹配方式：0-拼音,1-五笔
    mint简码 = Val(zlDatabase.GetPara("简码方式"))
End Sub



Private Sub Get煎法()
     '中药煎法
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    On Error GoTo errH
    strSQL = "Select A.ID,A.编码,A.名称 From 诊疗项目目录 A" & _
        " Where A.类别='E' And A.操作类型='3' And A.服务对象 IN(1,2,3)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
    If rsTmp.EOF Then
        MsgBox "未找到有效的中药煎法，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    For i = 1 To rsTmp.RecordCount
        cboData.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboData.ItemData(cboData.NewIndex) = rsTmp!ID
        If rsTmp!ID = mlng煎法ID Then
            Call Cbo.SetIndex(cboData.hwnd, cboData.NewIndex)
        End If
        rsTmp.MoveNext
    Next
    
    If cboData.ListCount = 1 And cboData.ListIndex = -1 Then Call Cbo.SetIndex(cboData.hwnd, 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    vsAdvice.SetFocus
End Sub


Private Sub cboData_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboData.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = Cbo.MatchIndex(cboData.hwnd, KeyAscii)
        If lngIdx = -1 And cboData.ListCount > 0 Then lngIdx = 0
        cboData.ListIndex = lngIdx
    End If
End Sub
