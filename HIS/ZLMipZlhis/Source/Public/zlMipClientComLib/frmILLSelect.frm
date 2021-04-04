VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmILLSelect1 
   AutoRedraw      =   -1  'True
   Caption         =   "疾病选择器"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmILLSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBottom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   45
      TabIndex        =   13
      Top             =   4890
      Width           =   8880
      Begin VB.CommandButton cmdUnUse 
         Caption         =   "取消常用(&U)"
         Height          =   350
         Left            =   3405
         TabIndex        =   9
         Top             =   120
         Width           =   1230
      End
      Begin VB.ComboBox cbo常用 
         Height          =   300
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   150
         Width           =   1590
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6255
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7350
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCommon 
         Caption         =   "设为常用(&M)"
         Height          =   350
         Left            =   255
         TabIndex        =   7
         Top             =   120
         Width           =   1230
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   4245
      Left            =   3315
      TabIndex        =   4
      Top             =   615
      Width           =   5745
      _cx             =   10134
      _cy             =   7488
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmILLSelect.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
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
   Begin MSComctlLib.ImageList iimg16 
      Left            =   1125
      Top             =   3405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":0618
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":0BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":114C
            Key             =   "wubi"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":16E6
            Key             =   "spell"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTop 
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   -75
      Width           =   9070
      Begin VB.TextBox txtLocate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3840
         TabIndex        =   14
         ToolTipText     =   "查找下一个按F3或回车，定位输入框按F4"
         Top             =   225
         Width           =   1665
      End
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   6765
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   2160
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   1005
         TabIndex        =   1
         Top             =   225
         Width           =   2160
      End
      Begin VB.Image imgCodeType 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   5550
         Top             =   250
         Width           =   240
      End
      Begin VB.Label lblLocate 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Left            =   3360
         TabIndex        =   15
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl类别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编码类别"
         Height          =   180
         Left            =   5970
         TabIndex        =   12
         Top             =   285
         Width           =   720
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "对应科室"
         Height          =   180
         Left            =   210
         TabIndex        =   0
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.Frame fraLR 
      BorderStyle     =   0  'None
      Height          =   4245
      Left            =   3225
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   615
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvwTree_s 
      Height          =   4245
      Left            =   15
      TabIndex        =   3
      Top             =   630
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   7488
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "iimg16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmILLSelect1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mstr类别 As String
Private mlng病人科室ID As Long
Private mstr性别 As String
Private mblnMultiSel As Boolean
Private mblnICD10 As Boolean

Private mrsList As ADODB.Recordset

Private mblnOK As Boolean
Private mstrLike As String
Private mlngPreDept As Long
Private mintPreClass As Integer
Private mstrPreNode As String
Private mint简码 As Integer
Private mbln简码修改 As Boolean

Public Function ShowMe(frmParent As Object, ByVal str类别 As String, ByVal lng病人科室ID As Long, _
    Optional ByVal str性别 As String, Optional ByVal blnMultiSel As Boolean, Optional ByVal blnICD10 As Boolean = True) As ADODB.Recordset
    mstr类别 = str类别
    mlng病人科室ID = lng病人科室ID
    mstr性别 = str性别
    mblnMultiSel = blnMultiSel
    mblnICD10 = blnICD10
    Debug.Print mstr类别
    Set mfrmParent = frmParent
    Me.Show 1, frmParent
    
    If mblnOK Then Set ShowMe = mrsList
End Function

Private Sub cbo常用_Click()
    Call SetControlEnabled
End Sub

Private Sub cbo科室_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, blnDo As Boolean, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo科室.ListIndex = -1 Then Exit Sub
    If cbo科室.ItemData(cbo科室.ListIndex) = mlngPreDept And cbo科室.ItemData(cbo科室.ListIndex) <> -1 Then Exit Sub
    
    blnDo = True
    If cbo科室.ItemData(cbo科室.ListIndex) = -1 Then
        '选择其他科室
        strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " Order by A.编码"
        vRect = GetControlRect(cbo科室.hwnd)
        Set rsTmp = gobjComLib.zlDatabase.ShowSelect(Me, strSQL, 0, IIf(mblnICD10, "选择疾病", "选择诊断"), , , , , , True, vRect.Left, vRect.Top, cbo科室.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo科室, rsTmp!id)
            '不另触发Click事件,在本事件结束时一并处理
            If intIdx <> -1 Then
                Call gobjComLib.zlControl.CboSetIndex(cbo科室.hwnd, intIdx)
            Else
                cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo科室.ListCount - 1
                cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!id
                Call gobjComLib.zlControl.CboSetIndex(cbo科室.hwnd, cbo科室.NewIndex)
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的科室(不引发Click)
            intIdx = SeekCboIndex(cbo科室, mlngPreDept)
            Call gobjComLib.zlControl.CboSetIndex(cbo科室.hwnd, intIdx)
            blnDo = False
        End If
    End If
    mlngPreDept = cbo科室.ItemData(cbo科室.ListIndex)
    
    '读取数据
    If blnDo Then
        Call SetControlEnabled
        Call FillTreeData
    End If
End Sub

Private Sub cbo科室_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(cbo科室)
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo科室.ListIndex = -1 Then
            Call cbo科室_Validate(blnCancel)
        End If
        If Not blnCancel Then
            Call cbo科室_Validate(False)
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo科室_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long
    Dim vRect As RECT, blnCancel As Boolean
    Dim strInput As String, i As Long
    
    If cbo科室.ListIndex <> -1 Then Exit Sub '已选中,不用处理
    If cbo科室.Text = "" Then Cancel = True: Exit Sub '无输入
    
    On Error GoTo errH
    
    strInput = UCase(gobjComLib.zlCommFun.GetNeedName(cbo科室.Text))
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"
    
    vRect = GetControlRect(cbo科室.hwnd)
    Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(mblnICD10, "疾病选择", "诊断选择"), False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo科室.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cbo科室, rsTmp!id)
        If intIdx <> -1 Then
            cbo科室.ListIndex = intIdx
        Else
            cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo科室.ListCount - 1
            cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!id
            cbo科室.ListIndex = cbo科室.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的科室。", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cbo类别_Click()
    If mintPreClass = cbo类别.ListIndex Then Exit Sub
    mintPreClass = cbo类别.ListIndex
    
    Call FillTreeData
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCommon_Click()
    Dim arrSQL As Variant, i As Integer
    
    If cbo常用.ListIndex = -1 Then
        MsgBox "请指定当前" & IIf(mblnICD10, "疾病", "诊断") & "的常用科室。", vbInformation, gstrSysName
        cbo常用.SetFocus: Exit Sub
    End If
    If cbo常用.ItemData(cbo常用.ListIndex) = cbo科室.ItemData(cbo科室.ListIndex) Then
        MsgBox "该" & IIf(mblnICD10, "疾病", "诊断") & "已经设置为""" & cbo常用.Text & """的常用" & IIf(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
        cbo常用.SetFocus: Exit Sub
    End If
    
    arrSQL = Array()
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And .RowData(i) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mblnICD10 Then
                        arrSQL(UBound(arrSQL)) = "zl_疾病编码科室_Insert(" & .RowData(i) & "," & cbo常用.ItemData(cbo常用.ListIndex) & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "zl_疾病诊断科室_Insert(" & .RowData(i) & "," & cbo常用.ItemData(cbo常用.ListIndex) & ")"
                    End If
                End If
            Next
        End If
        If UBound(arrSQL) = -1 Then
            If .RowData(.Row) = 0 Then
                MsgBox "没有选择" & IIf(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
                Exit Sub
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mblnICD10 Then
                arrSQL(UBound(arrSQL)) = "zl_疾病编码科室_Insert(" & .RowData(.Row) & "," & cbo常用.ItemData(cbo常用.ListIndex) & ")"
            Else
                arrSQL(UBound(arrSQL)) = "zl_疾病诊断科室_Insert(" & .RowData(.Row) & "," & cbo常用.ItemData(cbo常用.ListIndex) & ")"
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call gobjComLib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
        
    MsgBox "已设置。", vbInformation, gstrSysName
    vsList.SetFocus
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim strFilter As String
    Dim i As Long
    
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 Then
                    strFilter = strFilter & " Or 项目ID=" & .RowData(i)
                End If
            Next
            strFilter = Mid(strFilter, 5)
        End If
        If strFilter = "" Then
            If .RowData(.Row) = 0 Then
                MsgBox "没有选择" & IIf(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
                Exit Sub
            End If
            strFilter = "项目ID=" & .RowData(.Row)
        End If
        
        mrsList.Filter = strFilter
        If mrsList.EOF Then
            MsgBox "没有选择" & IIf(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdUnUse_Click()
    Dim arrSQL As Variant, i As Integer
    
    If MsgBox("确实要将选择的" & IIf(mblnICD10, "疾病", "诊断") & "从" & gobjComLib.zlCommFun.GetNeedName(cbo常用.Text) & "中取消吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    arrSQL = Array()
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And .RowData(i) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mblnICD10 Then
                        arrSQL(UBound(arrSQL)) = "Zl_疾病编码科室_Delete(" & .RowData(i) & "," & cbo常用.ItemData(cbo常用.ListIndex) & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_疾病诊断科室_Delete(" & .RowData(i) & "," & cbo常用.ItemData(cbo常用.ListIndex) & ")"
                    End If
                End If
            Next
        End If
        If UBound(arrSQL) = -1 Then
            If .RowData(.Row) = 0 Then
                MsgBox "没有选择" & IIf(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
                Exit Sub
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mblnICD10 Then
                arrSQL(UBound(arrSQL)) = "Zl_疾病编码科室_Delete(" & .RowData(.Row) & "," & cbo常用.ItemData(cbo常用.ListIndex) & ")"
            Else
                arrSQL(UBound(arrSQL)) = "Zl_疾病诊断科室_Delete(" & .RowData(.Row) & "," & cbo常用.ItemData(cbo常用.ListIndex) & ")"
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call gobjComLib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    mstrPreNode = ""
    Call tvwTree_s_NodeClick(tvwTree_s.SelectedItem)
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub InitListTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    If mblnICD10 Then
        strHead = ",255,4;编码,1000,1;附码,550,1;名称,2500,1;说明,1500,1;分类ID,0,1;简码,0,1"
    Else
        strHead = ",255,4;编码,1000,1;名称,2500,1;说明,1500,1;编者,850,1;分类ID,0,1;简码,0,1"
    End If
    arrHead = Split(strHead, ";")
    With vsList
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(.FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim blnDept As Boolean, blnHave As Boolean
        
    Call InitListTable
    Call gobjComLib.RestoreWinState(Me, App.ProductName, mfrmParent.Name & IIf(mblnICD10, 1, 0))
    
    mblnOK = False
    mlngPreDept = -1
    mintPreClass = -1
    mstrPreNode = ""
    Set mrsList = Nothing
    mstrLike = IIf(Val(gobjComLib.zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    
    If Not mblnICD10 Then Me.Caption = "诊断选择器"
    
    On Error GoTo errH
    
    '检查是否有对应科室
    If mblnICD10 Then
        If mstr类别 = "" Then
            strSQL = "Select A.* From 疾病编码科室 A,部门人员 B,上机人员表 C" & _
                " Where A.科室ID=B.部门ID And B.人员ID=C.人员ID And C.用户名=User And Rownum=1"
        Else
            strSQL = "Select A.* From 疾病编码目录 I,疾病编码科室 A,部门人员 B,上机人员表 C" & _
                " Where I.ID=A.疾病ID And A.科室ID=B.部门ID And B.人员ID=C.人员ID" & _
                " And (I.撤档时间 is Null Or I.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.用户名=User And Instr([1],I.类别)>0 And Rownum=1"
        End If
    Else
        If mstr类别 = "" Then mstr类别 = "1,2"
        strSQL = "Select A.* From 疾病诊断目录 I,疾病诊断科室 A,部门人员 B,上机人员表 C" & _
            " Where I.ID=A.诊断ID And A.科室ID=B.部门ID And B.人员ID=C.人员ID" & _
            " And C.用户名=User And Instr([1],I.类别)>0 And Rownum=1"
    End If
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr类别)
    If Not rsTmp.EOF Then blnDept = True
    
    '显示当前人员科室
    strSQL = "Select A.ID,A.编码,A.简码,A.名称,Max(Nvl(C.缺省,0)) as 缺省" & _
        " From 部门表 A,部门性质说明 B,部门人员 C,上机人员表 D" & _
        " Where A.ID=B.部门ID And B.工作性质 IN('临床','检查','检验','手术','治疗','营养')" & _
        " And A.ID=C.部门ID And C.人员ID=D.人员ID And D.用户名=User" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " Group by A.ID,A.编码,A.简码,A.名称" & _
        " Order by A.编码"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo科室.AddItem IIf(mblnICD10, "所有疾病", "所有诊断")
    Do While Not rsTmp.EOF
        blnHave = True
        cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!id
        If blnDept Then
            If rsTmp!id = mlng病人科室ID Then
                Call gobjComLib.zlControl.CboSetIndex(cbo科室.hwnd, cbo科室.NewIndex)
            ElseIf cbo科室.ListIndex = -1 And rsTmp!缺省 = 1 Then
                Call gobjComLib.zlControl.CboSetIndex(cbo科室.hwnd, cbo科室.NewIndex)
            End If
        End If
        
        cbo常用.AddItem rsTmp!名称
        cbo常用.ItemData(cbo常用.NewIndex) = rsTmp!id
        If rsTmp!id = mlng病人科室ID Then
            cbo常用.ListIndex = cbo常用.NewIndex
        ElseIf cbo常用.ListIndex = -1 And rsTmp!缺省 = 1 Then
            cbo常用.ListIndex = cbo常用.NewIndex
        End If
        
        rsTmp.MoveNext
    Loop
    cbo科室.AddItem "<其他科室...>"
    cbo科室.ItemData(cbo科室.NewIndex) = -1
    
    If cbo科室.ListIndex = -1 Then
        If Not blnDept Or Not blnHave Then
            '无任何疾病对应科室设置时,或者人员无对应科室时，缺省显示所有疾病
            Call gobjComLib.zlControl.CboSetIndex(cbo科室.hwnd, 0)
        Else
            Call gobjComLib.zlControl.CboSetIndex(cbo科室.hwnd, 1)
        End If
    End If
    If cbo常用.ListCount > 0 And cbo常用.ListIndex = -1 Then
        cbo常用.ListIndex = 0
    End If
    
    '显示疾病编码类别
    If mblnICD10 Then
        If mstr类别 = "" Then
            strSQL = "Select 编码,类别,是否分类 From 疾病编码类别 Order by 优先级"
        Else
            strSQL = "Select 编码,类别,是否分类 From 疾病编码类别 Where Instr([1],编码)>0 Order by 优先级"
        End If
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr类别)
        Do While Not rsTmp.EOF
            cbo类别.AddItem rsTmp!编码 & ". " & rsTmp!类别
            cbo类别.ItemData(cbo类别.NewIndex) = NVL(rsTmp!是否分类, 0)
            rsTmp.MoveNext
        Loop
        Call gobjComLib.zlControl.CboSetIndex(cbo类别.hwnd, 0)
        If cbo类别.ListCount = 1 Then cbo类别.Locked = True
    Else
        lbl类别.Visible = False
        cbo类别.Visible = False
    End If
    
    mint简码 = Val(gobjComLib.zlDatabase.GetPara("简码方式"))
    mbln简码修改 = Val(gobjComLib.zlDatabase.GetPara("简码匹配方式切换")) = 1
    If mint简码 = 1 Then
        imgCodeType.Picture = iimg16.ListImages("wubi").Picture
        imgCodeType.Tag = "wubi"
    Else
        imgCodeType.Picture = iimg16.ListImages("spell").Picture
        imgCodeType.Tag = "spell"
    End If
    
    '缺省读取数据
    Call FillTreeData
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraTop.Left = 0
    fraTop.Top = -75
    fraTop.Width = Me.ScaleWidth
    
    If fraTop.Width - cbo类别.Width - 200 > 4135 Then
        cbo类别.Left = fraTop.Width - cbo类别.Width - 200
        lbl类别.Left = cbo类别.Left - lbl类别.Width - 75
    End If
    
    fraBottom.Left = 0
    fraBottom.Top = Me.ScaleHeight - fraBottom.Height
    fraBottom.Width = Me.ScaleWidth
    
    If fraBottom.Width - cmdCancel.Width - 550 > 7000 Then
        cmdCancel.Left = fraBottom.Width - cmdCancel.Width - 800
        cmdOK.Left = cmdCancel.Left - cmdOK.Width
    End If
    
    tvwTree_s.Left = 0
    tvwTree_s.Top = fraTop.Top + fraTop.Height + 15
    tvwTree_s.Height = Me.ScaleHeight - tvwTree_s.Top - fraBottom.Height
    
    fraLR.Top = tvwTree_s.Top
    fraLR.Left = tvwTree_s.Left + tvwTree_s.Width
    fraLR.Height = tvwTree_s.Height
    
    vsList.Top = tvwTree_s.Top
    vsList.Left = IIf(tvwTree_s.Visible, fraLR.Left + fraLR.Width, 0)
    vsList.Width = Me.ScaleWidth - vsList.Left
    vsList.Height = tvwTree_s.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gobjComLib.SaveWinState(Me, App.ProductName, mfrmParent.Name & IIf(mblnICD10, 1, 0))
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvwTree_s.Width + X < 1000 Or vsList.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvwTree_s.Width = tvwTree_s.Width + X
        vsList.Left = vsList.Left + X
        vsList.Width = vsList.Width - X
    End If
End Sub

Private Sub FillTreeData()
'功能：读取疾病分类数据，可能是科室对应疾病只对应的分类
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objNode As Node
    
    '清除数据
    Set mrsList = Nothing
    tvwTree_s.Nodes.Clear
    vsList.Rows = vsList.FixedRows
    vsList.Rows = vsList.FixedRows + 1
    
    'ICD-10类别是否有分类
    If mblnICD10 Then
        If cbo类别.ItemData(cbo类别.ListIndex) = 0 Then
            tvwTree_s.Visible = False
            fraLR.Visible = False
        Else
            tvwTree_s.Visible = True
            fraLR.Visible = True
        End If
        Call Form_Resize
    End If
    
    Screen.MousePointer = 11
    Me.Refresh
    
    On Error GoTo errH
    
    If mblnICD10 Then
        If cbo类别.ItemData(cbo类别.ListIndex) <> 0 Then '为0表示该种疾病没有分类
            If cbo科室.ItemData(cbo科室.ListIndex) = 0 Then
                strSQL = "Select ID,上级ID,序号,名称 From 疾病编码分类 Where 类别=[1]" & _
                    " Start With 上级ID is Null Connect by Prior ID=上级ID Order by Level,序号"
            Else
                strSQL = _
                    " Select Distinct B.分类id From 疾病编码科室 A, 疾病编码目录 B" & _
                    " Where A.疾病id = B.ID And A.科室id = [2]" & _
                    " And (B.撤档时间 is Null Or B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
                strSQL = _
                    "Select Max(Level) as 级ID, ID, 上级id, 序号, 名称" & vbNewLine & _
                    "From 疾病编码分类 Where 类别=[1]" & vbNewLine & _
                    "Start With ID In (" & strSQL & ")" & vbNewLine & _
                    "Connect By Prior 上级id = ID" & vbNewLine & _
                    "Group By ID, 上级ID, 序号, 名称" & vbNewLine & _
                    "Order By 级ID Desc"
                strSQL = "Select ID, 上级id, 序号, 名称 From (" & strSQL & ")"
            End If
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Left(cbo类别.Text, 1), cbo科室.ItemData(cbo科室.ListIndex))
            Do Until rsTmp.EOF
                If IsNull(rsTmp!上级ID) Then
                    Set objNode = tvwTree_s.Nodes.Add(, , "_" & rsTmp!id, "【" & rsTmp!序号 & "】" & Trim(rsTmp!名称), 1, 2)
                Else
                    Set objNode = tvwTree_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!id, "【" & rsTmp!序号 & "】" & Trim(rsTmp!名称), 1, 2)
                End If
                rsTmp.MoveNext
            Loop
        End If
    Else
        If cbo科室.ItemData(cbo科室.ListIndex) = 0 Then
            strSQL = "Select ID,上级ID,编码,名称 From 疾病诊断分类 Where Instr([1],类别)>0" & _
                " Start With 上级ID is Null Connect by Prior ID=上级ID Order by Level,编码"
        Else
            strSQL = _
                " Select Distinct C.分类ID From 疾病诊断科室 A, 疾病诊断目录 B,疾病诊断属类 C" & _
                " Where A.诊断ID = B.ID And B.ID=C.诊断ID And A.科室ID = [2]"
            strSQL = _
                "Select Max(Level) as 级ID, ID, 上级id, 编码, 名称" & vbNewLine & _
                "From 疾病诊断分类 Where Instr([1],类别)>0" & vbNewLine & _
                "Start With ID In (" & strSQL & ")" & vbNewLine & _
                "Connect By Prior 上级id = ID" & vbNewLine & _
                "Group By ID, 上级ID, 编码, 名称" & vbNewLine & _
                "Order By 级ID Desc"
            strSQL = "Select ID, 上级id, 编码, 名称 From (" & strSQL & ")"
        End If
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr类别, cbo科室.ItemData(cbo科室.ListIndex))
        Do Until rsTmp.EOF
            If IsNull(rsTmp!上级ID) Then
                Set objNode = tvwTree_s.Nodes.Add(, , "_" & rsTmp!id, "[" & rsTmp!编码 & "]" & Trim(rsTmp!名称), 1, 2)
            Else
                Set objNode = tvwTree_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!id, "[" & rsTmp!编码 & "]" & Trim(rsTmp!名称), 1, 2)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    If tvwTree_s.Nodes.count > 0 Then
        tvwTree_s.Nodes(1).Selected = True
        tvwTree_s.Nodes(1).Expanded = True
        tvwTree_s.Nodes(1).EnsureVisible
    End If
    
    Screen.MousePointer = 0
    Call FillListData
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub FillListData()
    Dim strSQL As String
    Dim str性别 As String
    Dim lng分类ID As Long
    Dim i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    vsList.Rows = vsList.FixedRows
    vsList.Rows = vsList.FixedRows + 1
    vsList.ColHidden(0) = Not mblnMultiSel
    
    If mblnICD10 Then
        If mstr性别 Like "*男*" Then
            str性别 = "男"
        ElseIf mstr性别 Like "*女*" Then
            str性别 = "女"
        End If
        
        If cbo科室.ItemData(cbo科室.ListIndex) <> 0 Then
            strSQL = _
                " Select A.ID as 项目ID,A.编码,A.序号,A.附码,A.名称,A.说明, a.分类ID, a.简码" & _
                " From 疾病编码目录 A,疾病编码科室 B" & _
                " Where A.ID=B.疾病ID And B.科室ID=[1] And A.类别=[2]" & _
                " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
        Else
            strSQL = "Select A.ID as 项目ID,A.编码,A.序号,A.附码,A.名称,A.说明, a.分类ID, a.简码 From 疾病编码目录 A" & _
                " Where A.类别=[2] And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
        End If
        If str性别 <> "" Then strSQL = strSQL & " And (A.性别限制=[4] Or A.性别限制 is Null)"
        
        If cbo类别.ItemData(cbo类别.ListIndex) <> 0 Then '为0表示该种疾病没有分类
            If tvwTree_s.SelectedItem Is Nothing Then
                vsList.Row = 1: Screen.MousePointer = 0: Exit Sub
            End If
            
            lng分类ID = Val(Mid(tvwTree_s.SelectedItem.Key, 2))
            strSQL = strSQL & " And (A.分类ID=[3] Or A.分类ID In(SELECT a.Id " & _
               "FROM 疾病编码分类 a, 疾病编码分类 b " & _
               "WHERE a.类别 = [2] AND (a.上级id = b.Id OR b.上级id IS NULL) AND a.类别 = b.类别 AND b.Id = [3]))"
        End If
        strSQL = strSQL & " Order by A.编码,A.序号"
        
        
        Set mrsList = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo科室.ItemData(cbo科室.ListIndex), Left(cbo类别.Text, 1), lng分类ID, str性别)
        If Not mrsList.EOF Then
            With vsList
                .Redraw = flexRDNone
                .Rows = .FixedRows + mrsList.RecordCount
                For i = 1 To mrsList.RecordCount
                    .RowData(i) = Val(mrsList!项目ID)
                    .TextMatrix(i, 0) = 0
                    .TextMatrix(i, 1) = NVL(mrsList!编码)
                    .Cell(flexcpData, i, 1) = CStr(NVL(mrsList!编码))
                    If NVL(mrsList!编码) = .Cell(flexcpData, i - 1, 1) Then
                        If Not IsNull(mrsList!序号) Then
                            .TextMatrix(i, 1) = .TextMatrix(i, 1) & "." & mrsList!序号
                            If .TextMatrix(i - 1, 1) = .Cell(flexcpData, i - 1, 1) And mrsList!序号 = 2 Then
                                .TextMatrix(i - 1, 1) = .TextMatrix(i - 1, 1) & ".1"
                            End If
                        End If
                    End If
                    
                    .TextMatrix(i, 2) = NVL(mrsList!附码)
                    .TextMatrix(i, 3) = NVL(mrsList!名称)
                    .TextMatrix(i, 4) = NVL(mrsList!说明)
                    .TextMatrix(i, 5) = NVL(mrsList!分类ID)
                    .TextMatrix(i, 6) = NVL(mrsList!简码)
                    mrsList.MoveNext
                Next
                .Redraw = flexRDDirect
            End With
        End If
    Else
        If tvwTree_s.SelectedItem Is Nothing Then
            vsList.Row = 1: Screen.MousePointer = 0: Exit Sub
        End If
        lng分类ID = Val(Mid(tvwTree_s.SelectedItem.Key, 2))
        
        If cbo科室.ItemData(cbo科室.ListIndex) <> 0 Then
            'strSQL = _
                " Select A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                " From 疾病诊断目录 A,疾病诊断科室 B,疾病诊断属类 C" & _
                " Where A.ID=B.诊断ID And A.ID=C.诊断ID And B.科室ID=[1] And Instr([2],A.类别)>0 And C.分类ID=[3]" & _
                " Order by A.编码"
            strSQL = "SELECT a.Id AS 项目id, a.编码, a.名称, a.说明, a.编者, c.分类ID, '' as 简码 " & vbNewLine & _
                    "FROM 疾病诊断目录 a, 疾病诊断科室 b, 疾病诊断属类 c" & vbNewLine & _
                    "WHERE a.Id = b.诊断id AND a.Id = c.诊断id AND b.科室id = [1] AND Instr([2], a.类别) > 0 AND" & vbNewLine & _
                    "      c.分类id IN ((SELECT Id FROM 疾病诊断分类 WHERE Instr([2], 类别) > 0 AND Id = [3] OR 上级id = [3]))" & vbNewLine & _
                    "ORDER BY a.编码"

        Else
            'strSQL = "Select A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                " From 疾病诊断目录 A,疾病诊断属类 B" & _
                " Where Instr([2],A.类别)>0 And A.ID=B.诊断ID " & _
                " And B.分类ID=[3]" & _
                " Order by A.编码"
            strSQL = "SELECT a.Id AS 项目id, a.编码, a.名称, a.说明, a.编者, b.分类ID, '' as 简码" & vbNewLine & _
                    "FROM 疾病诊断目录 a, 疾病诊断属类 b" & vbNewLine & _
                    "WHERE Instr([2], a.类别) > 0 AND a.Id = b.诊断id AND" & vbNewLine & _
                    "      b.分类id IN (SELECT Id FROM 疾病诊断分类 WHERE Instr([2], 类别) > 0 AND Id = [3] OR 上级id = [3])" & vbNewLine & _
                    "ORDER BY a.编码"
        End If
        Set mrsList = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo科室.ItemData(cbo科室.ListIndex), mstr类别, lng分类ID)
        If Not mrsList.EOF Then
            With vsList
                .Redraw = flexRDNone
                .Rows = .FixedRows + mrsList.RecordCount
                For i = 1 To mrsList.RecordCount
                    .RowData(i) = Val(mrsList!项目ID)
                    .TextMatrix(i, 0) = 0
                    .TextMatrix(i, 1) = NVL(mrsList!编码)
                    .TextMatrix(i, 2) = NVL(mrsList!名称)
                    .TextMatrix(i, 3) = NVL(mrsList!说明)
                    .TextMatrix(i, 4) = NVL(mrsList!编者)
                    .TextMatrix(i, 5) = NVL(mrsList!分类ID)
                    .TextMatrix(i, 6) = NVL(mrsList!简码)
                    mrsList.MoveNext
                Next
                .Redraw = flexRDDirect
            End With
        End If
    End If
    
    vsList.Row = 1: vsList.Col = 1
    Screen.MousePointer = 0
    
    Call vsList_AfterRowColChange(-1, -1, vsList.Row, vsList.Col)
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub imgCodeType_Click()
    If Not mbln简码修改 Then Exit Sub
    If imgCodeType.Tag = "spell" Then
        Call gobjComLib.zlDatabase.SetPara("简码方式", 1)
        mint简码 = 1
        imgCodeType.Picture = iimg16.ListImages("wubi").Picture
        imgCodeType.Tag = "wubi"
    Else
        Call gobjComLib.zlDatabase.SetPara("简码方式", 0)
        mint简码 = 0
        imgCodeType.Picture = iimg16.ListImages("spell").Picture
        imgCodeType.Tag = "spell"
    End If
End Sub

Private Sub tvwTree_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrPreNode = Node.Key Then Exit Sub
    mstrPreNode = Node.Key
    Call FillListData
End Sub

Private Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Sub txtLocate_GotFocus()
    gobjComLib.zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    Dim strSQL As String, str性别 As String
    Dim strInput As String
    Dim rsTmp As ADODB.Recordset
    Dim vRect As RECT
    Dim blnCancle As Boolean
    
    If KeyAscii = vbKeyReturn Then
        On Error GoTo errH
        strInput = UCase(Trim(txtLocate.Text))
        
        If Not mblnICD10 Then
            '诊断目录
            If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                strSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
            ElseIf gobjComLib.zlCommFun.IsCharAlpha(strInput) Then
                strSQL = "B.简码 Like [2] Or B.名称 Like [2]"
            Else
                strSQL = "A.编码 Like [1] Or B.名称 Like [2]"
            End If
            strSQL = _
                " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者,D.分类ID" & _
                " From 疾病诊断目录 A,疾病诊断别名 B,疾病诊断科室 C,疾病诊断属类 D" & _
                " Where  A.ID=C.诊断ID(+) And A.ID=B.诊断ID AND a.Id = D.诊断id " & _
                IIf(Val(cbo科室.ItemData(cbo科室.ListIndex)) <> 0, " And C.科室ID=[3]", "") & _
                " And B.码类=[5] and instr([6],A.类别)>0 And (" & strSQL & ")" & _
                " Order by A.编码"
        Else
            If mstr性别 Like "*男*" Then
                str性别 = "男"
            ElseIf mstr性别 Like "*女*" Then
                str性别 = "女"
            End If
            If gobjComLib.zlCommFun.IsCharChinese(strInput) Then
                strSQL = "A.名称 Like [2]" '输入汉字时只匹配名称
            ElseIf gobjComLib.zlCommFun.IsCharAlpha(strInput) Then
                strSQL = "A.名称 Like [2] Or " & IIf(mint简码 = 0, "a.简码", "a.五笔码") & " Like [2]"
            Else
                strSQL = "A.编码 Like [1] Or A.名称 Like [2]"
            End If
            strSQL = _
                " Select A.ID,A.ID as 项目ID,A.编码,A.附码,A.名称," & IIf(mint简码 = 0, "a.简码", "a.五笔码 as 简码") & ",A.说明,A.分类ID" & _
                " From 疾病编码目录 A,疾病编码科室 B Where A.ID=B.疾病ID(+) " & _
                IIf(Val(cbo科室.ItemData(cbo科室.ListIndex)) <> 0, " And B.科室ID=[3]", "") & _
                " And (" & strSQL & ") And a.类别=[6]" & _
                IIf(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is NULL)", "") & _
                " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order by A.编码"
        End If
        vRect = GetControlRect(txtLocate.hwnd)
        
        Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(Not mblnICD10, "疾病诊断", "疾病编码"), False, "", "", False, False, True, _
            vRect.Left, vRect.Bottom, 0, blnCancle, False, True, strInput & "%", mstrLike & strInput & "%", Val(cbo科室.ItemData(cbo科室.ListIndex)), str性别, mint简码 + 1, IIf(mblnICD10, Left(cbo类别.Text, 1), mstr类别))

        '检查诊断输入方式
        If blnCancle Then Exit Sub
        If rsTmp Is Nothing Then
            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
        Else
            '定位
            If txtLocate.Tag <> txtLocate.Text Then
                lblLocate.Tag = ""
                txtLocate.Tag = txtLocate.Text
            End If
            
            lngStart = Val("" & lblLocate.Tag) + 1
            If lngStart >= vsList.Rows Then lngStart = 1
            '确定左边树节点
            tvwTree_s.Nodes("_" & rsTmp!分类ID).Selected = True
            tvwTree_s_NodeClick tvwTree_s.Nodes("_" & rsTmp!分类ID)
            '确定 VSLIST 项目
            For i = lngStart To vsList.Rows - 1
                If Val(vsList.RowData(i) & "") = Val(rsTmp!id & "") Then
                    vsList.Row = i
                    vsList.TopRow = i
                    lblLocate.Tag = i
                    vsList.SetFocus
                    Exit For
                End If
            Next
        End If
    End If
    Exit Sub
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnEnabled As Boolean, i As Integer
    
    Call SetControlEnabled
    
    '在有数据的情况下，只能取消自已所属科室的常用疾病
    If vsList.RowData(vsList.Row) <> 0 Then
        blnEnabled = True
    End If
    cmdUnUse.Enabled = blnEnabled
End Sub

Private Sub vsList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then
        If Val(vsList.TextMatrix(Row, 0)) <> 0 Then
            vsList.Cell(flexcpBackColor, Row, 0, Row, vsList.Cols - 1) = &HC0FFFF
        Else
            vsList.Cell(flexcpBackColor, Row, 0, Row, vsList.Cols - 1) = vsList.BackColor
        End If
    End If
End Sub

Private Sub vsList_DblClick()
    If vsList.MouseRow >= vsList.FixedRows Then
        Call cmdOK_Click
    End If
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    ElseIf KeyAscii = 32 Then
        If mblnMultiSel And vsList.Col > 0 And vsList.RowData(vsList.Row) <> 0 Then
            vsList.TextMatrix(vsList.Row, 0) = IIf(Val(vsList.TextMatrix(vsList.Row, 0)) = 0, 1, 0)
        End If
    End If
End Sub

Private Sub SetControlEnabled()
    Dim blnEnabled As Boolean
    
    '设为常用的可用性
    blnEnabled = True
    If cbo常用.ListIndex = -1 Then
        blnEnabled = False
    ElseIf cbo科室.ListIndex <> -1 And cbo常用.ListIndex <> -1 Then
        If cbo科室.ItemData(cbo科室.ListIndex) = cbo常用.ItemData(cbo常用.ListIndex) Then
            blnEnabled = False
        End If
    End If
    If blnEnabled Then
        If vsList.Row >= vsList.FixedRows Then
            blnEnabled = vsList.RowData(vsList.Row) <> 0
        End If
    End If
    cmdCommon.Enabled = blnEnabled
    
    '确定按钮的可用性
    blnEnabled = True
    If vsList.Row >= vsList.FixedRows Then
        blnEnabled = vsList.RowData(vsList.Row) <> 0
    Else
        blnEnabled = False
    End If
    cmdOK.Enabled = blnEnabled
End Sub

Private Sub vsList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsList.RowData(Row) = 0 Then
        Cancel = True
    ElseIf Col <> 0 Then
        Cancel = True
    End If
End Sub
