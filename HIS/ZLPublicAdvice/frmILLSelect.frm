VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmILLSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "疾病选择器"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmILLSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Begin VB.CommandButton cmdCommon 
         Caption         =   "个人常用(&P)"
         Height          =   350
         Index           =   1
         Left            =   100
         TabIndex        =   16
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdUnUse 
         Caption         =   "取消常用(&U)"
         Height          =   350
         Left            =   4485
         TabIndex        =   9
         Top             =   135
         Width           =   1230
      End
      Begin VB.ComboBox cbo常用 
         Height          =   300
         Left            =   2610
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   165
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
         Caption         =   "科室常用(&M)"
         Height          =   350
         Index           =   0
         Left            =   1335
         TabIndex        =   7
         Top             =   135
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
      Cols            =   10
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
            Picture         =   "frmILLSelect.frx":06A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":0C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":11D4
            Key             =   "wubi"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":176E
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
Attribute VB_Name = "frmILLSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const S_SUB As String = ",28," '扩展码
'主干码章节
Private Const S_MAIN As String = ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,27," ' 常规 普通诊断选择范围
Private Const S_SSZD As String = ",23," '损伤中毒
Private Const S_BLZD As String = ",2,"
Private Const S_ZYZD As String = ",26,"

'入口参数
Private mfrmParent As Object
Private mstr类别 As String
Private mlng病人科室ID As Long
Private mstr性别 As String
Private mblnMultiSel As Boolean
Private mblnICD10 As Boolean
Private mbln病案系统 As Boolean


Private mrsList As ADODB.Recordset
Private mblnOK As Boolean
Private mstrLike As String
Private mlngPreDept As Long
Private mintPreClass As Integer
Private mstrPreNode As String
Private mint简码 As Integer
Private mbln简码修改 As Boolean
Private mstrSel编码 As String
Private mlngUserID As Long '医生id/操作员id
Private mInt适用范围 As Integer    '1-门诊病人;2-住院病人;0-全院

Private mlngICD11 As Long '-1-判断系统参数，0-非ICD11,1-按ICD11录入，
Private mbln主 As Boolean '是否是录入主干码，true 录入主干码，false 扩展码

Private mlngDiagType As Long '诊断录入类型，
        '说明：1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码;8-术前诊断;9-术后诊断;10-并发症;11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断;22-影像学诊断
        '目前只分了两类：7 和 非7，

Private mstr章节 As String '当前过滤的编码范围
Private mbytModel As Byte  '=1 临床路径管理病种设置调用

Private mblnParICD11 As Boolean

Private Enum DiseaseCols
    ColSel = 0
    Col编码 = 1
    col附码 = 2 '只有疾病编码有
    Col名称 = 3
    col说明 = 4
    col编者 = 5 '只有诊断编码有
    Col分类ID = 6
    Col简码 = 7
    Col诊断id = 8 '疾病编码时使用，疾病对应的诊断
    Col疾病Id = 9 '诊断编码时使用,诊断对应的疾病
End Enum

Public Function ShowMe(frmParent As Object, ByVal str类别 As String, ByVal lng病人科室ID As Long, Optional ByVal str性别 As String, _
    Optional ByVal blnMultiSel As Boolean, Optional ByVal blnICD10 As Boolean = True, Optional ByVal strSel编码 As String, Optional ByVal lngSys As Long = 100, Optional ByVal intPatiType As Integer, _
    Optional ByVal lngICD11 As Long, Optional ByVal bln主 As Boolean, Optional ByVal lngDiagType As Long, Optional ByVal bytModel As Byte) As ADODB.Recordset
    
    mstr类别 = str类别
    mlng病人科室ID = lng病人科室ID
    mstr性别 = str性别
    mblnMultiSel = blnMultiSel
    mblnICD10 = blnICD10
    mstrSel编码 = strSel编码
    mbln病案系统 = (lngSys \ 100 = 3)
    mlngICD11 = lngICD11
    mbln主 = bln主
    mlngDiagType = lngDiagType
    Set mfrmParent = frmParent
    mInt适用范围 = intPatiType
    mbytModel = bytModel
    Me.Show 1, frmParent
    
    If mblnOK Then Set ShowMe = mrsList
End Function

Private Sub cbo常用_Click()
    Call SetControlEnabled
End Sub

Private Sub cbo科室_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim intIdx As Integer, blnDo As Boolean
    Dim vRect As Variant, blnCancel As Boolean
        
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
        vRect = gobjComlib.zlControl.GetControlRect(cbo科室.hWnd)
        Set rsTmp = gobjComlib.zlDatabase.ShowSelect(Me, strSQL, 0, IIF(mblnICD10, "选择疾病", "选择诊断"), , , , , , True, vRect.Left, vRect.Top, cbo科室.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = gobjComlib.cbo.FindIndex(cbo科室, rsTmp!ID)
            '不另触发Click事件,在本事件结束时一并处理
            If intIdx <> -1 Then
                Call gobjComlib.zlControl.CboSetIndex(cbo科室.hWnd, intIdx)
            Else
                cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo科室.ListCount - 1
                cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
                Call gobjComlib.zlControl.CboSetIndex(cbo科室.hWnd, cbo科室.NewIndex)
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的科室(不引发Click)
            intIdx = gobjComlib.cbo.FindIndex(cbo科室, mlngPreDept)
            Call gobjComlib.zlControl.CboSetIndex(cbo科室.hWnd, intIdx)
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
    Call gobjComlib.zlControl.TxtSelAll(cbo科室)
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
            Call gobjComlib.ZLCommFun.PressKey(vbKeyTab)
        End If
        '病案系统没用找到，则
        If cbo科室.ListIndex = -1 And mbln病案系统 Then cbo科室.ListIndex = 0
    Else
        If mbln病案系统 Then KeyAscii = 0
    End If
End Sub

Private Sub cbo科室_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long
    Dim vRect As Variant, blnCancel As Boolean
    Dim strInput As String
    
    If cbo科室.ListIndex <> -1 Then Exit Sub '已选中,不用处理
    If cbo科室.Text = "" Then Cancel = True: Exit Sub '无输入
    
    On Error GoTo errH
    
    strInput = UCase(gobjComlib.ZLCommFun.GetNeedName(cbo科室.Text))
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"
    
    vRect = gobjComlib.zlControl.GetControlRect(cbo科室.hWnd)
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(mblnICD10, "疾病选择", "诊断选择"), False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo科室.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = gobjComlib.cbo.FindIndex(cbo科室, rsTmp!ID)
        If intIdx <> -1 Then
            cbo科室.ListIndex = intIdx
        Else
            cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo科室.ListCount - 1
            cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cbo类别_Click()
    If mintPreClass = cbo类别.ListIndex Then Exit Sub
    mintPreClass = cbo类别.ListIndex
    
    Call FillTreeData
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCommon_Click(Index As Integer)
    Dim arrSQL As Variant, i As Long
    Dim strPar As String
    
    If Index = 0 Then '科室常用
        If cbo常用.ListIndex = -1 Then
            MsgBox "请指定当前" & IIF(mblnICD10, "疾病", "诊断") & "的常用科室。", vbInformation, gstrSysName
            cbo常用.SetFocus: Exit Sub
        End If
        If cbo常用.ItemData(cbo常用.ListIndex) = cbo科室.ItemData(cbo科室.ListIndex) Then
            MsgBox "该" & IIF(mblnICD10, "疾病", "诊断") & "已经设置为""" & cbo常用.Text & """的常用" & IIF(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
            cbo常用.SetFocus: Exit Sub
        End If
        strPar = cbo常用.ItemData(cbo常用.ListIndex)
    ElseIf Index = 1 Then '个人常用
        If mlngUserID = cbo科室.ItemData(cbo科室.ListIndex) Then
            MsgBox "该" & IIF(mblnICD10, "疾病", "诊断") & "已经设置为个人的常用" & IIF(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
            cbo常用.SetFocus: Exit Sub
        End If
        strPar = "Null," & mlngUserID
    End If
    
    arrSQL = Array()
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And .RowData(i) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mblnICD10 Then
                        arrSQL(UBound(arrSQL)) = "zl_疾病编码科室_Insert(" & .RowData(i) & "," & strPar & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "zl_疾病诊断科室_Insert(" & .RowData(i) & "," & strPar & ")"
                    End If
                End If
            Next
        End If
        If UBound(arrSQL) = -1 Then
            If .RowData(.Row) = 0 Then
                MsgBox "没有选择" & IIF(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
                Exit Sub
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mblnICD10 Then
                arrSQL(UBound(arrSQL)) = "zl_疾病编码科室_Insert(" & .RowData(.Row) & "," & strPar & ")"
            Else
                arrSQL(UBound(arrSQL)) = "zl_疾病诊断科室_Insert(" & .RowData(.Row) & "," & strPar & ")"
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call gobjComlib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
        
    MsgBox "已设置。", vbInformation, gstrSysName
    vsList.SetFocus
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim strFilter As String
    Dim i As Long
    
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And Val(.RowData(i)) <> 0 Then
                    strFilter = strFilter & " Or 项目ID=" & .RowData(i)
                End If
            Next
            strFilter = Mid(strFilter, 5)
        End If
        If strFilter = "" Then
            If Val(.RowData(.Row)) = 0 Then
                MsgBox "没有选择" & IIF(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
                Exit Sub
            End If
            strFilter = "项目ID=" & .RowData(.Row)
        End If
        
        mrsList.Filter = strFilter
        If mrsList.EOF Then
            MsgBox "没有选择" & IIF(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdUnUse_Click()
    Dim arrSQL As Variant, i As Long
    Dim strPar As String
    Dim strTmp As String
    
    If cbo科室.List(cbo科室.ListIndex) = IIF(mblnICD10, "所有疾病", "所有诊断") Then '删全部
        strPar = cbo常用.ItemData(cbo常用.ListIndex) & "," & mlngUserID
        strTmp = "个人常用和" & gobjComlib.ZLCommFun.GetNeedName(cbo常用.Text)
    ElseIf cbo科室.List(cbo科室.ListIndex) = "个人常用" Then '删个人常用
        strPar = "Null," & mlngUserID
        strTmp = "个人常用"
    Else '删科室常用
        strPar = cbo常用.ItemData(cbo常用.ListIndex)
        strTmp = gobjComlib.ZLCommFun.GetNeedName(cbo常用.Text)
    End If
    
    If MsgBox("确实要将选择的" & IIF(mblnICD10, "疾病", "诊断") & "从" & strTmp & "中取消吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    arrSQL = Array()
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And .RowData(i) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mblnICD10 Then
                        arrSQL(UBound(arrSQL)) = "Zl_疾病编码科室_Delete(" & .RowData(i) & "," & strPar & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_疾病诊断科室_Delete(" & .RowData(i) & "," & strPar & ")"
                    End If
                End If
            Next
        End If
        If UBound(arrSQL) = -1 Then
            If .RowData(.Row) = 0 Then
                MsgBox "没有选择" & IIF(mblnICD10, "疾病", "诊断") & "。", vbInformation, gstrSysName
                Exit Sub
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mblnICD10 Then
                arrSQL(UBound(arrSQL)) = "Zl_疾病编码科室_Delete(" & .RowData(.Row) & "," & strPar & ")"
            Else
                arrSQL(UBound(arrSQL)) = "Zl_疾病诊断科室_Delete(" & .RowData(.Row) & "," & strPar & ")"
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call gobjComlib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    mstrPreNode = ""
    Call tvwTree_s_NodeClick(tvwTree_s.SelectedItem)
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    Dim blnDoc As Boolean

    mstr章节 = ""
    If mlngICD11 = 1 Then
        mblnICD10 = True
        mstr类别 = "E"
        '确定编码的范围
        If mbytModel = 1 Then
            If mlng病人科室ID > 0 Then
                If gobjComlib.sys.DeptHaveProperty(mlng病人科室ID, "中医科") Then
                    mstr章节 = S_MAIN & "," & S_ZYZD
                Else
                    mstr章节 = S_MAIN
                End If
            Else
                mstr章节 = S_MAIN & S_ZYZD
            End If
        Else
            If mlngDiagType = 7 Then
                mstr章节 = S_SSZD
            ElseIf mlngDiagType = 6 Then
                mstr章节 = S_BLZD
            ElseIf mlngDiagType = 11 Or mlngDiagType = 12 Or mlngDiagType = 13 Then
                mstr章节 = S_ZYZD
            Else
                mstr章节 = S_MAIN
            End If
        
            If Not mbln主 Then
                If mlngDiagType = 11 Or mlngDiagType = 12 Or mlngDiagType = 13 Then
                    mstr章节 = S_ZYZD
                Else
                    mstr章节 = S_SUB
                End If
            End If
        End If
    End If
    
    '界面设置
    With vsList
        If mlngICD11 = 1 Then
            .ColHidden(col附码) = True
        Else
            .ColHidden(col附码) = Not mblnICD10
        End If
        .ColHidden(col编者) = mblnICD10
        .Rows = 1: .Rows = .FixedRows + 1
    End With
    If Not mblnICD10 Then Me.Caption = "诊断选择器"
    Call gobjComlib.RestoreWinState(Me, App.ProductName, mfrmParent.Name & IIF(mblnICD10, 1, 0))
    
    If mbln病案系统 Then
        '病案系统不显示科室项目
        lbl科室.Caption = "疾病选择"
        cmdCommon(0).Visible = False: cmdCommon(0).Enabled = False
        cbo常用.Visible = False: cbo常用.Enabled = False
        cmdUnUse.Left = cmdCommon(0).Left
    End If
    cbo科室.AddItem IIF(mblnICD10, "所有疾病", "所有诊断")
    cbo科室.AddItem "个人常用"
    
    
    mblnOK = False
    mlngPreDept = -1
    mintPreClass = -1
    mstrPreNode = ""
    Set mrsList = Nothing
    mstrLike = IIF(Val(gobjComlib.zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    
    On Error GoTo errH
    Call gobjComlib.zlDatabase.GetUserInfo
    mlngUserID = UserInfo.ID
    cbo科室.ItemData(cbo科室.NewIndex) = mlngUserID '个人常用项目
    
    '病案系统不关心其他科室
    If Not mbln病案系统 Then
        '检查是否对应操作员
        If mlngUserID <> 0 Then
            strSQL = "select * from " & IIF(mblnICD10, "疾病编码科室", "疾病诊断科室") & " where 人员id=[1] and Rownum<2"
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUserID)
            If Not rsTmp.EOF Then blnDoc = True
        End If
        
        '检查是否有对应科室
        If Not blnDoc Then
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
                    " And (I.撤档时间 is Null Or I.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And C.用户名=User And Instr([1],I.类别)>0 And Rownum=1"
            End If
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr类别)
            If Not rsTmp.EOF Then blnDept = True
        End If
        
        '显示当前人员科室
        strSQL = "Select A.ID,A.编码,A.简码,A.名称,Max(Nvl(C.缺省,0)) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C,上机人员表 D" & _
            " Where A.ID=B.部门ID And B.工作性质 IN('临床','检查','检验','手术','治疗','营养')" & _
            " And A.ID=C.部门ID And C.人员ID=D.人员ID And D.用户名=User" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " Group by A.ID,A.编码,A.简码,A.名称" & _
            " Order by A.编码"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        If blnDoc Then Call gobjComlib.zlControl.CboSetIndex(cbo科室.hWnd, cbo科室.NewIndex)
        
        Do While Not rsTmp.EOF
            blnHave = True
            cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
            If blnDept Then
                If rsTmp!ID = mlng病人科室ID Then
                    Call gobjComlib.zlControl.CboSetIndex(cbo科室.hWnd, cbo科室.NewIndex)
                ElseIf cbo科室.ListIndex = -1 And rsTmp!缺省 = 1 Then
                    Call gobjComlib.zlControl.CboSetIndex(cbo科室.hWnd, cbo科室.NewIndex)
                End If
            End If
            
            cbo常用.AddItem rsTmp!名称
            cbo常用.ItemData(cbo常用.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng病人科室ID Then
                cbo常用.ListIndex = cbo常用.NewIndex
            ElseIf cbo常用.ListIndex = -1 And rsTmp!缺省 = 1 Then
                cbo常用.ListIndex = cbo常用.NewIndex
            End If
            
            rsTmp.MoveNext
        Loop
        cbo科室.AddItem "<其他科室...>"
        cbo科室.ItemData(cbo科室.NewIndex) = -1
        
        If cbo常用.ListCount > 0 And cbo常用.ListIndex = -1 Then
            cbo常用.ListIndex = 0
        End If
    End If
    
    If cbo科室.ListIndex = -1 Then
        If Not blnDept Or Not blnHave Or Not blnDoc Then
            '无任何疾病对应科室设置时,或者人员无对应科室时，缺省显示所有疾病
            Call gobjComlib.zlControl.CboSetIndex(cbo科室.hWnd, 0) '病案系统设置为所有疾病
        Else
            Call gobjComlib.zlControl.CboSetIndex(cbo科室.hWnd, 1)
        End If
    End If

    '显示疾病编码类别
    If mblnICD10 Then
        If mstr类别 = "" Then
            strSQL = "Select 编码,类别,是否分类 From 疾病编码类别 Order by 优先级"
        Else
            strSQL = "Select 编码,类别,是否分类 From 疾病编码类别 Where Instr([1],编码)>0 Order by 优先级"
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr类别)
        Do While Not rsTmp.EOF
            cbo类别.AddItem rsTmp!编码 & ". " & rsTmp!类别
            cbo类别.ItemData(cbo类别.NewIndex) = Nvl(rsTmp!是否分类, 0)
            If mstr类别 <> "" Then
                If rsTmp!编码 & "" = Mid(mstr类别, 1, 1) Then
                    Call gobjComlib.zlControl.CboSetIndex(cbo类别.hWnd, cbo类别.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Loop
        If mstr类别 = "" Then Call gobjComlib.zlControl.CboSetIndex(cbo类别.hWnd, 0)
        If cbo类别.ListCount = 1 Then cbo类别.Locked = True
    Else
        lbl类别.Visible = False
        cbo类别.Visible = False
    End If
    
    mint简码 = Val(gobjComlib.zlDatabase.GetPara("简码方式"))
    mbln简码修改 = Val(gobjComlib.zlDatabase.GetPara("简码匹配方式切换")) = 1
    
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
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
    vsList.Left = IIF(tvwTree_s.Visible, fraLR.Left + fraLR.Width, 0)
    vsList.Width = Me.ScaleWidth - vsList.Left
    vsList.Height = tvwTree_s.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gobjComlib.SaveWinState(Me, App.ProductName, mfrmParent.Name & IIF(mblnICD10, 1, 0))
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
    Dim strICD11 As String
    Dim str序号 As String
    Dim lng疾病序号 As Long, lng证候序号 As Long
    
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
            If mlngICD11 = 1 Then
                If mstr章节 <> "" Then
                    strICD11 = " And instr('" & mstr章节 & "',','||A.章节||',')>0"
                    If InStr(mstr章节, S_ZYZD) > 0 Then
                        strSQL = "Select 序号 From 疾病编码分类 Where 章节 = '26' And 名称 = '传统医学疾病（TM1）' And 编码 = 'L1-SA0'"
                        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "疾病编码分类")
                        If Not rsTmp.EOF Then
                            lng疾病序号 = Val("" & rsTmp!序号)
                        End If
                        
                        strSQL = "Select 序号 From 疾病编码分类 Where 章节 = '26' And 名称 = '传统医学证候（TM1）' And 编码 = 'L1-SE7'"
                        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "疾病编码分类")
                        If Not rsTmp.EOF Then
                            lng证候序号 = Val("" & rsTmp!序号)
                        End If
                        If mbln主 Then
                            strICD11 = strICD11 & IIF(lng证候序号 <> 0, " And a.序号 Not In (Select e.序号 From 疾病编码分类 e Where e.章节 = '26' And e.序号 >=" & lng证候序号 & ")", "")
                        Else
                            strICD11 = strICD11 & IIF(lng证候序号 <> 0 And lng疾病序号 <> 0, " And a.序号 Not In (Select e.序号 From 疾病编码分类 e Where e.章节 = '26' And (e.序号 >=" & lng疾病序号 & " And e.序号 <" & lng证候序号 & "))", "")
                        End If
                    End If
                End If
            End If
            
            If cbo科室.ItemData(cbo科室.ListIndex) = 0 Then
                strSQL = "Select a.ID,a.上级ID,a.序号,decode(a.类别,'E', a.编码||' ',null) || a.名称 as 名称 From 疾病编码分类 a Where a.类别=[1]" & _
                    " And (a.撤档时间 is Null Or a.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) " & strICD11 & vbNewLine & _
                    " Start With a.上级ID is Null Connect by Prior a.ID=a.上级ID Order by Level,a.序号"
            Else
                strSQL = _
                    " Select Distinct B.分类id From 疾病编码科室 A, 疾病编码目录 B Where A.疾病id = B.ID" & _
                    IIF(cbo科室.List(cbo科室.ListIndex) = "个人常用", " And A.人员id = [3]", " And A.科室id = [2]") & _
                    IIF(mInt适用范围 = 0, "", " And (Nvl(B.适用范围,0) = 0 Or B.适用范围 = " & CStr(mInt适用范围) & ") ") & _
                    " And (B.撤档时间 is Null Or B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
                strSQL = _
                    "Select Max(Level) as 级ID, a.ID, a.上级id, a.序号, decode(a.类别,'E', a.编码||' ',null) || a.名称 as 名称 " & vbNewLine & _
                    "From 疾病编码分类 a Where a.类别=[1] " & strICD11 & "  And (a.撤档时间 is Null Or a.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                    "Start With a.ID In (" & strSQL & ")" & vbNewLine & _
                    "Connect By Prior a.上级id = a.ID" & vbNewLine & _
                    "Group By a.ID, a.上级ID, a.序号, a.名称,a.类别,a.编码" & vbNewLine & _
                    "Order By 级ID Desc"
                strSQL = "Select a.ID, a.上级id, a.序号, a.名称 From (" & strSQL & ") a"
            End If
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Left(cbo类别.Text, 1), cbo科室.ItemData(cbo科室.ListIndex), mlngUserID)
            Do Until rsTmp.EOF
                If "E" = Left(cbo类别.Text, 1) Then
                    str序号 = ""
                Else
                    str序号 = "【" & rsTmp!序号 & "】"
                End If
                If IsNull(rsTmp!上级ID) Then
                    Set objNode = tvwTree_s.Nodes.Add(, , "_" & rsTmp!ID, str序号 & Trim(rsTmp!名称), 1, 2)
                Else
                    Set objNode = tvwTree_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, str序号 & Trim(rsTmp!名称), 1, 2)
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
                " Where A.诊断ID = B.ID And B.ID=C.诊断ID" & _
                " And (B.撤档时间 is Null Or B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) " & _
                IIF(mInt适用范围 = 0, "", " And (Nvl(B.适用范围,0) = 0 Or B.适用范围 = " & CStr(mInt适用范围) & ") ") & _
                IIF(cbo科室.List(cbo科室.ListIndex) = "个人常用", " And A.人员id = [3]", " And A.科室id = [2]")
            strSQL = _
                "Select Max(Level) as 级ID, ID, 上级id, 编码, 名称" & vbNewLine & _
                "From 疾病诊断分类 Where Instr([1],类别)>0" & vbNewLine & _
                "Start With ID In (" & strSQL & ")" & vbNewLine & _
                "Connect By Prior 上级id = ID" & vbNewLine & _
                "Group By ID, 上级ID, 编码, 名称" & vbNewLine & _
                "Order By 级ID Desc"
            strSQL = "Select ID, 上级id, 编码, 名称 From (" & strSQL & ")"
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr类别, cbo科室.ItemData(cbo科室.ListIndex), mlngUserID)
        Do Until rsTmp.EOF
            If IsNull(rsTmp!上级ID) Then
                Set objNode = tvwTree_s.Nodes.Add(, , "_" & rsTmp!ID, "[" & rsTmp!编码 & "]" & Trim(rsTmp!名称), 1, 2)
            Else
                Set objNode = tvwTree_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, "[" & rsTmp!编码 & "]" & Trim(rsTmp!名称), 1, 2)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    If tvwTree_s.Nodes.Count > 0 Then
        tvwTree_s.Nodes(1).Selected = True
        tvwTree_s.Nodes(1).Expanded = True
        tvwTree_s.Nodes(1).EnsureVisible
    End If
    
    Screen.MousePointer = 0
    Call FillListData
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub FillListData()
    Dim strSQL As String, strSQLTmp As String
    Dim str性别 As String
    Dim lng分类ID As Long, str类别 As String
    Dim i As Long
    Dim str编码 As String
    Dim strICD11 As String
    Dim lng疾病序号 As Long, lng证候序号 As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    vsList.Rows = vsList.FixedRows
    vsList.Rows = vsList.FixedRows + 1
    vsList.ColHidden(0) = Not mblnMultiSel
    
    If mstr性别 Like "*男*" Then
        str性别 = "男"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "女"
    End If
    
    If mlngICD11 = 1 Then
        If mstr章节 <> "" Then
            strICD11 = " And instr('" & mstr章节 & "',','||A.章节||',')>0"
            If InStr(mstr章节, S_ZYZD) > 0 Then
                strSQL = "Select 序号 From 疾病编码分类 Where 章节 = '26' And 名称 = '传统医学疾病（TM1）' And 编码 = 'L1-SA0'"
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "疾病编码分类")
                If Not rsTmp.EOF Then
                    lng疾病序号 = Val("" & rsTmp!序号)
                End If
                
                strSQL = "Select 序号 From 疾病编码分类 Where 章节 = '26' And 名称 = '传统医学证候（TM1）' And 编码 = 'L1-SE7'"
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "疾病编码分类")
                If Not rsTmp.EOF Then
                    lng证候序号 = Val("" & rsTmp!序号)
                End If
                If mbln主 Then
                    strICD11 = strICD11 & IIF(lng证候序号 <> 0, " And a.序号 Not In (Select e.序号 From 疾病编码分类 e Where e.章节 = '26' And e.序号 >=" & lng证候序号 & ")", "")
                Else
                    strICD11 = strICD11 & IIF(lng证候序号 <> 0 And lng疾病序号 <> 0, " And a.序号 Not In (Select e.序号 From 疾病编码分类 e Where e.章节 = '26' And (e.序号 >=" & lng疾病序号 & " And e.序号 <" & lng证候序号 & "))", "")
                End If
            End If
        End If
    End If
        
    If mblnICD10 Then
        If cbo类别.ItemData(cbo类别.ListIndex) <> 0 Then '为0表示该种疾病没有分类
            If tvwTree_s.SelectedItem Is Nothing Then
                vsList.Row = 1: Screen.MousePointer = 0: Exit Sub
            End If
            lng分类ID = Val(Mid(tvwTree_s.SelectedItem.Key, 2))
            strSQLTmp = " And (A.分类id = [3] " & strICD11 & " Or" & vbNewLine & _
                        "      A.分类id In (Select A.Id" & vbNewLine & _
                        "                  From 疾病编码分类 A, 疾病编码分类 B" & vbNewLine & _
                        "                  Where A.类别 = [2] " & strICD11 & " And (A.上级id = B.Id Or B.上级id Is Null) And A.类别 = B.类别 And B.Id = [3]))"
        End If
    Else
        If tvwTree_s.SelectedItem Is Nothing Then
            vsList.Row = 1: Screen.MousePointer = 0: Exit Sub
        End If
        lng分类ID = Val(Mid(tvwTree_s.SelectedItem.Key, 2))
    End If
    
    If cbo科室.ItemData(cbo科室.ListIndex) <> 0 Then
        If mblnICD10 Then '按疾病编码输入
            If mbln病案系统 Then
                strSQL = "Select A.Id As 项目id, A.编码, A.序号, A.附码, Null 附码ID, Null 附码名称, A.名称, A.说明, Null 编者, A.分类id, A.简码, A.疗效限制, A.分娩, C.是否病人,A.编码 疾病编码, A.Id 疾病id,A.类别 疾病类别, Null 诊断id" & vbNewLine & _
                        "From 疾病编码目录 A, 疾病编码科室 B, 疾病编码分类 C " & vbNewLine & _
                        "Where A.Id = B.疾病id And A.类别 = [2] And A.分类id = C.Id(+)" & IIF(cbo科室.List(cbo科室.ListIndex) = "个人常用", " And b.人员id = [5]", " ") & vbNewLine & _
                        "  And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIF(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is Null) ", " ") & IIF(mInt适用范围 = 0, "", " And (Nvl(A.适用范围,0) = 0 Or A.适用范围 = " & CStr(mInt适用范围) & ") ") & strSQLTmp
            Else
                strSQL = "Select A.Id As 项目id, A.编码, A.序号, A.附码, Null 附码ID, Null 附码名称, A.名称, A.说明, Null 编者, A.分类id, A.简码, A.疗效限制, A.分娩, C.是否病人,A.编码 疾病编码, A.Id 疾病id,A.类别 疾病类别,Max(D.诊断id) 诊断id" & vbNewLine & _
                        "From 疾病编码目录 A, 疾病编码科室 B, 疾病编码分类 C, 疾病诊断对照 D" & vbNewLine & _
                        "Where A.Id = B.疾病id And A.类别 = [2] And A.分类id = C.Id(+) And A.Id = D.疾病id(+) And" & vbNewLine & _
                        IIF(cbo科室.List(cbo科室.ListIndex) = "个人常用", " b.人员id = [5] And ", "  b.科室id = [1] And ") & _
                        " (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIF(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is Null) ", " ") & strSQLTmp & vbNewLine & _
                        IIF(mInt适用范围 = 0, "", " And (Nvl(A.适用范围,0) = 0 Or A.适用范围 = " & CStr(mInt适用范围) & ") ") & _
                        "Group By A.Id, A.编码, A.序号, A.附码, A.名称, A.说明, A.分类id, A.简码, A.疗效限制, A.分娩, C.是否病人,A.类别"
            End If
        Else '按诊断输入
            strSQL = "Select A.Id As 项目id, A.编码, Null 序号, Null 附码,Null 附码ID, Null 附码名称, A.名称, A.说明, A.编者, C.分类id, '' As 简码, 0 疗效限制, 0 分娩, 0 是否病人, Max(D.疾病id) 疾病id," & vbNewLine & _
                    "       A.Id 诊断id" & vbNewLine & _
                    "From 疾病诊断目录 A, 疾病诊断科室 B, 疾病诊断属类 C, 疾病诊断对照 D" & vbNewLine & _
                    "Where A.Id = B.诊断id And A.Id = D.疾病id(+) And A.Id = C.诊断id And Instr([2], A.类别) > 0 " & IIF(cbo科室.List(cbo科室.ListIndex) = "个人常用", " And b.人员id = [5]", " And b.科室id = [1]") & vbNewLine & _
                    " And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) " & _
                    "     And C.分类id In ((Select ID From 疾病诊断分类 Where Instr([2], 类别) > 0 And ID = [3] Or 上级id = [3]))" & IIF(mInt适用范围 = 0, "", " And (Nvl(A.适用范围,0) = 0 Or A.适用范围 = " & CStr(mInt适用范围) & ") ") & vbNewLine & _
                    "Group By A.Id, A.编码,A.名称, A.说明, A.编者, C.分类id"
            '读取对应的疾病编码,附码
            strSQL = "Select A.项目id, A.编码, B.序号, B.附码, Null  附码ID, Null 附码名称, A.名称, A.说明, Null 编者, A.分类id, A.简码, A.疗效限制, A.分娩, A.是否病人,B.编码 疾病编码, B.Id 疾病id,B.类别 疾病类别,A.诊断id" & vbNewLine & _
                            "From (" & strSQL & ") A,疾病编码目录 B " & vbNewLine & _
                            "Where A.疾病id=B.ID(+) "
        End If
    Else
        If mblnICD10 Then '按疾病编码输入
            If mbln病案系统 Then
                strSQL = "Select A.Id As 项目id, A.编码, A.序号, A.附码,Null 附码ID, Null 附码名称, A.名称, A.说明, Null 编者, A.分类id, A.简码,  A.疗效限制, A.分娩, C.是否病人,A.编码 疾病编码, A.Id 疾病id,A.类别 疾病类别, Null 诊断id" & vbNewLine & _
                    "From 疾病编码目录 A, 疾病编码分类 C" & vbNewLine & _
                    "Where A.类别 = [2] And A.分类id = C.Id(+)  And" & vbNewLine & _
                    "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIF(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is Null) ", " ") & IIF(mInt适用范围 = 0, "", " And (Nvl(A.适用范围,0) = 0 Or A.适用范围 = " & CStr(mInt适用范围) & ") ") & strSQLTmp
            Else
                strSQL = "Select A.Id As 项目id, A.编码, A.序号, A.附码,Null 附码ID, Null 附码名称, A.名称, A.说明, Null 编者, A.分类id, A.简码,  A.疗效限制, A.分娩, C.是否病人,A.编码 疾病编码, A.Id 疾病id,A.类别 疾病类别," & vbNewLine & _
                        "       Max(B.诊断id) 诊断id" & vbNewLine & _
                        "From 疾病编码目录 A, 疾病诊断对照 B, 疾病编码分类 C" & vbNewLine & _
                        "Where A.类别 = [2] And A.Id = B.疾病id(+) And A.分类id = C.Id(+)  And" & vbNewLine & _
                        "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIF(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is Null) ", " ") & strSQLTmp & vbNewLine & _
                        IIF(mInt适用范围 = 0, "", " And (Nvl(A.适用范围,0) = 0 Or A.适用范围 = " & CStr(mInt适用范围) & ") ") & _
                        "Group By A.Id, A.编码, A.序号, A.附码, A.名称, A.说明, A.分类id, A.简码, A.疗效限制, A.分娩,A.类别, C.是否病人"
            End If
        Else '按诊断输入
            strSQL = "Select A.Id As 项目id, A.编码, Null 序号, Null 附码,Null 附码ID, Null 附码名称, A.名称, A.说明, A.编者, B.分类ID, '' As 简码, 0 疗效限制, 0 分娩, 0 是否病人," & vbNewLine & _
                    "       Max(D.疾病id) 疾病id, A.Id 诊断id" & vbNewLine & _
                    "From 疾病诊断目录 A, 疾病诊断属类 B, 疾病诊断对照 D" & vbNewLine & _
                    "Where Instr([2], A.类别) > 0 And A.Id = B.诊断id And A.Id = D.疾病id(+) And" & vbNewLine & _
                    "  (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And " & _
                    "      B.分类id In (Select ID From 疾病诊断分类 Where Instr([2], 类别) > 0 And ID = [3] Or 上级id = [3])" & vbNewLine & _
                    IIF(mInt适用范围 = 0, "", " And (Nvl(A.适用范围,0) = 0 Or A.适用范围 = " & CStr(mInt适用范围) & ") ") & _
                    "Group By A.Id, A.编码, A.名称, A.说明, A.编者, B.分类ID"
            '读取对应的疾病编码,附码
            strSQL = "Select A.项目id, A.编码, B.序号, B.附码, Null  附码ID, Null 附码名称, A.名称, A.说明, Null 编者, A.分类id, A.简码, A.疗效限制, A.分娩, A.是否病人,B.编码 疾病编码, B.Id 疾病id,B.类别 疾病类别,A.诊断id" & vbNewLine & _
                            "From (" & strSQL & ") A,疾病编码目录 B " & vbNewLine & _
                            "Where A.疾病id=B.ID(+) "
        End If
    End If
    If mblnICD10 Then
        str类别 = Left(cbo类别.Text, 1)
    Else
        str类别 = mstr类别
    End If
    strSQL = strSQL & " Order by A.编码" & IIF(mblnICD10, ",A.序号", "")
    Set mrsList = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo科室.ItemData(cbo科室.ListIndex), str类别, lng分类ID, str性别, mlngUserID)
    
    If Not mrsList.EOF Then
        With vsList
            .Redraw = flexRDNone
            .Rows = .FixedRows + mrsList.RecordCount
            For i = 1 To mrsList.RecordCount
                .RowData(i) = Val(mrsList!项目ID & "")
                str编码 = mrsList!编码 & ""
                .TextMatrix(i, Col编码) = str编码
                .TextMatrix(i, Col名称) = mrsList!名称 & ""
                .TextMatrix(i, Col分类ID) = mrsList!分类ID & ""
                .TextMatrix(i, Col简码) = mrsList!简码 & ""
                .TextMatrix(i, col说明) = mrsList!说明 & ""
                .TextMatrix(i, col附码) = mrsList!附码 & ""
                .TextMatrix(i, col编者) = mrsList!编者 & ""
                .TextMatrix(i, Col诊断id) = mrsList!诊断id & ""
                .TextMatrix(i, Col疾病Id) = mrsList!疾病Id & ""
                .Cell(flexcpData, i, Col编码) = CStr(str编码)
                If mstrSel编码 <> "" Then
                    If InStr(mstrSel编码, "," & str编码 & ",") > 0 Then
                        .TextMatrix(i, ColSel) = 1
                    Else
                        .TextMatrix(i, ColSel) = 0
                    End If
                Else
                    .TextMatrix(i, ColSel) = 0
                End If
                
                If mblnICD10 Then
                    If str编码 = .Cell(flexcpData, i - 1, Col编码) Then
                        If Not IsNull(mrsList!序号) Then
                            .TextMatrix(i, Col编码) = .TextMatrix(i, Col编码) & "." & mrsList!序号
                            If .TextMatrix(i - 1, Col编码) = .Cell(flexcpData, i - 1, Col编码) And mrsList!序号 = 2 Then
                                .TextMatrix(i - 1, Col编码) = .TextMatrix(i - 1, Col编码) & ".1"
                            End If
                        End If
                    End If
                End If
                
                mrsList.MoveNext
            Next
            .Redraw = flexRDDirect
        End With
    End If

    vsList.Row = 1: vsList.Col = 1
    Screen.MousePointer = 0
    Call vsList_AfterRowColChange(-1, -1, vsList.Row, vsList.Col)
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub imgCodeType_Click()
    If Not mbln简码修改 Then Exit Sub
    If imgCodeType.Tag = "spell" Then
        Call gobjComlib.zlDatabase.SetPara("简码方式", 1)
        mint简码 = 1
        imgCodeType.Picture = iimg16.ListImages("wubi").Picture
        imgCodeType.Tag = "wubi"
    Else
        Call gobjComlib.zlDatabase.SetPara("简码方式", 0)
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

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIF(IsNull(varValue), DefaultValue, varValue)
End Function

Private Sub txtLocate_GotFocus()
    gobjComlib.zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    Dim strSQL As String, str性别 As String
    Dim strInput As String
    Dim rsTmp As ADODB.Recordset
    Dim vRect As Variant
    Dim blnCancle As Boolean
    Dim strICD11 As String
    Dim lng疾病序号 As Long, lng证候序号 As Long
    
    If KeyAscii = vbKeyReturn Then
        On Error GoTo errH
        strInput = UCase(Trim(txtLocate.Text))
        
        If Not mblnICD10 Then
            '诊断目录
            If gobjComlib.ZLCommFun.IsCharChinese(strInput) Then
                strSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
            ElseIf gobjComlib.ZLCommFun.IsCharAlpha(strInput) Then
                strSQL = "B.简码 Like [2] Or B.名称 Like [2]"
            Else
                strSQL = "A.编码 Like [1] Or B.名称 Like [2]"
            End If
            strSQL = _
                " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者,D.分类ID" & _
                " From 疾病诊断目录 A,疾病诊断别名 B,疾病诊断科室 C,疾病诊断属类 D" & _
                " Where  A.ID=C.诊断ID(+) And A.ID=B.诊断ID AND a.Id = D.诊断id " & _
                " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) " & _
                IIF(Val(cbo科室.ItemData(cbo科室.ListIndex)) <> 0, " And C.科室ID=[3]", "") & _
                " And B.码类=[5] and instr([6],A.类别)>0 And (" & strSQL & ")" & _
                " Order by A.编码"
        Else
            
            If mlngICD11 = 1 Then
                If mstr章节 <> "" Then
                    strICD11 = " And instr('" & mstr章节 & "',','||A.章节||',')>0"
                    If InStr(mstr章节, S_ZYZD) > 0 Then
                        strSQL = "Select 序号 From 疾病编码分类 Where 章节 = '26' And 名称 = '传统医学疾病（TM1）' And 编码 = 'L1-SA0'"
                        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "疾病编码分类")
                        If Not rsTmp.EOF Then
                            lng疾病序号 = Val("" & rsTmp!序号)
                        End If
                        
                        strSQL = "Select 序号 From 疾病编码分类 Where 章节 = '26' And 名称 = '传统医学证候（TM1）' And 编码 = 'L1-SE7'"
                        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "疾病编码分类")
                        If Not rsTmp.EOF Then
                            lng证候序号 = Val("" & rsTmp!序号)
                        End If
                        If mbln主 Then
                            strICD11 = strICD11 & IIF(lng证候序号 <> 0, " And a.分类ID Not In (Select e.ID From 疾病编码分类 e where e.章节='26' And e.序号>=" & lng证候序号 & ")", "")
                        Else
                            strICD11 = strICD11 & IIF(lng证候序号 <> 0 And lng疾病序号 <> 0, " And a.分类ID Not In (Select e.ID From 疾病编码分类 e Where e.章节 = '26' And (e.序号 >=" & lng疾病序号 & " And e.序号 <" & lng证候序号 & "))", "")
                        End If
                    End If
                End If
            End If
    
            If mstr性别 Like "*男*" Then
                str性别 = "男"
            ElseIf mstr性别 Like "*女*" Then
                str性别 = "女"
            End If
            If gobjComlib.ZLCommFun.IsCharChinese(strInput) Then
                strSQL = "A.名称 Like [2]" '输入汉字时只匹配名称
            ElseIf gobjComlib.ZLCommFun.IsCharAlpha(strInput) Then
                strSQL = "A.名称 Like [2] Or " & IIF(mint简码 = 0, "a.简码", "a.五笔码") & " Like [2]"
            Else
                strSQL = "A.编码 Like [1] Or A.名称 Like [2]"
            End If
            strSQL = _
                " Select A.ID,A.ID as 项目ID,A.编码,A.附码,A.名称," & IIF(mint简码 = 0, "a.简码", "a.五笔码 as 简码") & ",A.说明,A.分类ID" & _
                " From 疾病编码目录 A,疾病编码科室 B Where A.ID=B.疾病ID(+) " & _
                IIF(Val(cbo科室.ItemData(cbo科室.ListIndex)) <> 0, " And B.科室ID=[3]", "") & _
                " And (" & strSQL & ") And a.类别=[6]" & strICD11 & _
                IIF(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is NULL)", "") & _
                " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order by A.编码"
        End If
        vRect = gobjComlib.zlControl.GetControlRect(txtLocate.hWnd)
        
        Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(Not mblnICD10, "疾病诊断", "疾病编码"), False, "", "", False, False, True, _
            vRect.Left, vRect.Bottom, 0, blnCancle, False, True, strInput & "%", mstrLike & strInput & "%", Val(cbo科室.ItemData(cbo科室.ListIndex)), str性别, mint简码 + 1, IIF(mblnICD10, Left(cbo类别.Text, 1), mstr类别))

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
            If tvwTree_s.Visible Then
                tvwTree_s.Nodes("_" & rsTmp!分类ID).Selected = True
                tvwTree_s_NodeClick tvwTree_s.Nodes("_" & rsTmp!分类ID)
            End If
            '确定 VSLIST 项目
            For i = lngStart To vsList.Rows - 1
                If Val(vsList.RowData(i) & "") = Val(rsTmp!ID & "") Then
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
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnEnabled As Boolean
    
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
        vsList.TextMatrix(vsList.RowSel, ColSel) = 1
        Call cmdOK_Click
    End If
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    ElseIf KeyAscii = 32 Then
        If mblnMultiSel And vsList.Col > 0 And vsList.RowData(vsList.Row) <> 0 Then
            vsList.TextMatrix(vsList.Row, 0) = IIF(Val(vsList.TextMatrix(vsList.Row, 0)) = 0, 1, 0)
        End If
    End If
End Sub

Private Sub SetControlEnabled()
    Dim blnEnabled As Boolean
    Dim bln个人常用 As Boolean
    
    '设为常用的可用性
    blnEnabled = True: bln个人常用 = True
    If cbo常用.ListIndex = -1 Then
        blnEnabled = False
    ElseIf cbo科室.ListIndex <> -1 And cbo常用.ListIndex <> -1 Then
        If cbo科室.ItemData(cbo科室.ListIndex) = cbo常用.ItemData(cbo常用.ListIndex) Then
            blnEnabled = False
        End If
        If cbo科室.List(cbo科室.ListIndex) = "个人常用" Then
            bln个人常用 = False
        End If
    End If
    If blnEnabled Or bln个人常用 Then
        If vsList.Row >= vsList.FixedRows Then
            blnEnabled = IIF(blnEnabled, vsList.RowData(vsList.Row) <> 0, blnEnabled)
            bln个人常用 = IIF(bln个人常用, vsList.RowData(vsList.Row) <> 0, bln个人常用)
        End If
    End If
    
    cmdCommon(0).Enabled = blnEnabled ' 科室常用
    cmdCommon(1).Enabled = bln个人常用 ' 个人常用
    
    '确定按钮的可用性
    blnEnabled = True
    If vsList.Row >= vsList.FixedRows Then
        blnEnabled = vsList.RowData(vsList.Row) <> 0
    Else
        blnEnabled = False
    End If
    cmdOK.Enabled = blnEnabled
End Sub

Private Sub vsList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    With vsList
        lngRow = .MouseRow
        If lngRow >= .FixedRows Then
            Call gobjComlib.ZLCommFun.ShowTipInfo(.hWnd, .TextMatrix(lngRow, col说明), True)     '路径外项目的添加原因
        Else
            Call gobjComlib.ZLCommFun.ShowTipInfo(.hWnd, "")
        End If
    End With
End Sub

Private Sub vsList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsList.RowData(Row) = 0 Then
        Cancel = True
    ElseIf Col <> 0 Then
        Cancel = True
    End If
End Sub

