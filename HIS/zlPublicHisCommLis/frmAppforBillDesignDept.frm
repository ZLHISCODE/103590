VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppforBillDesignDept 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "设置申请单执行小组"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2910
      TabIndex        =   2
      Top             =   2970
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4560
      TabIndex        =   1
      Top             =   2970
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFList 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      _cx             =   11192
      _cy             =   4948
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483635
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   2
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   0
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
End
Attribute VB_Name = "frmAppforBillDesignDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrDept As String
Dim mlngID As Long
Dim mblnfrmIfShow As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData = True Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If mblnfrmIfShow = False Then
        mblnfrmIfShow = True
        Call ReadData
    End If
End Sub

Private Sub Form_Load()
1         On Error GoTo Form_Load_Error

2         With Me.VSFList
3             .Rows = 2
4             .Cols = 4
5             .FixedRows = 1
6             .ColKey(0) = "编码": .ColWidth(.ColIndex("编码")) = 1500: .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("编码")) = "编码"
7                 .Cell(flexcpAlignment, 0, .ColIndex("编码"), 0, .ColIndex("编码")) = flexAlignCenterCenter
8             .ColKey(1) = "小组": .ColWidth(.ColIndex("小组")) = 2000: .ColAlignment(.ColIndex("小组")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("小组")) = "小组"
9                 .Cell(flexcpAlignment, 0, .ColIndex("小组"), 0, .ColIndex("小组")) = flexAlignCenterCenter
10            .ColKey(2) = "HIS部门编码": .ColWidth(.ColIndex("HIS部门编码")) = 2000: .ColAlignment(.ColIndex("HIS部门编码")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("HIS部门编码")) = "HIS部门编码"
11                .Cell(flexcpAlignment, 0, .ColIndex("HIS部门编码"), 0, .ColIndex("HIS部门编码")) = flexAlignCenterCenter
12            .ColKey(3) = "默认": .ColWidth(.ColIndex("默认")) = 600: .ColAlignment(.ColIndex("默认")) = flexAlignCenterCenter: .TextMatrix(0, .ColIndex("默认")) = "默认"
13        End With


14        Exit Sub
Form_Load_Error:
15        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(Form_Load)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
16        Err.Clear
End Sub
Public Function ShowMe(objFrm As Object, lngID As Long) As String
    mlngID = lngID
    Me.Show vbModal, objFrm
End Function
Private Sub ReadData()
          '功能           读入数据
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItem As String
          
1         On Error GoTo ReadData_Error

2         If gUserInfo.NodeNo <> "-" Then
3             strSQL = " select 编码,名称 小组,HIS部门编码 from 检验小组记录 where 站点=[1] or 站点 is null"
4         Else
5             strSQL = " select 编码,名称 小组,HIS部门编码 from 检验小组记录 "
6         End If
7         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "小组", gUserInfo.NodeNo)
8         With Me.VSFList
9             .Rows = 1
10            Do Until rsTmp.EOF
11                .Rows = .Rows + 1
12                .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTmp("编码") & ""
13                .TextMatrix(.Rows - 1, .ColIndex("小组")) = rsTmp("小组") & ""
14                .TextMatrix(.Rows - 1, .ColIndex("HIS部门编码")) = rsTmp("HIS部门编码") & ""
15                .Cell(flexcpChecked, .Rows - 1, .ColIndex("编码"), .Rows - 1, .ColIndex("编码")) = 2
16                .Cell(flexcpChecked, .Rows - 1, .ColIndex("默认"), .Rows - 1, .ColIndex("默认")) = 2
17                .Cell(flexcpPictureAlignment, .Rows - 1, .ColIndex("默认"), .Rows - 1, .ColIndex("默认")) = flexAlignCenterCenter
18                rsTmp.MoveNext
19            Loop
20        End With
          
21        strSQL = " select 执行小组,默认执行小组 from 检验申请单 where id = [1] "
22        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "申请单", mlngID)
23        If rsTmp.RecordCount > 0 Then
24            Call SetSel(rsTmp("执行小组") & "")
25            Call SetDefault(rsTmp("默认执行小组") & "")
26        End If


27        Exit Sub
ReadData_Error:
28        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(ReadData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
29        Err.Clear
          
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnfrmIfShow = False
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim intRow As Integer
          Dim intCol As Integer
1         On Error GoTo VSFList_MouseDown_Error

2         With Me.VSFList
3             If .MouseRow >= 0 And .MouseCol >= 0 Then
4                 intRow = .MouseRow
5                 intCol = .MouseCol
6                 If intCol = .ColIndex("编码") Then
7                     If .TextMatrix(intRow, intCol) = "" Then Exit Sub
8                     If .Cell(flexcpChecked, intRow, intCol) = 1 Then
9                         .Cell(flexcpChecked, intRow, intCol) = 2
10                        .Cell(flexcpChecked, intRow, .ColIndex("默认")) = 2
11                    Else
12                        .Cell(flexcpChecked, intRow, intCol) = 1
13                    End If
14                End If
15                If intCol = .ColIndex("默认") Then
16                    If .Cell(flexcpChecked, intRow, .ColIndex("编码")) = 1 Then
17                        ClsDefault
18                        If .Cell(flexcpChecked, intRow, intCol) = 1 Then
19                            .Cell(flexcpChecked, intRow, intCol) = 2
20                        Else
21                            .Cell(flexcpChecked, intRow, intCol) = 1
22                        End If
23                    End If
24                End If
25            End If
26        End With


27        Exit Sub
VSFList_MouseDown_Error:
28        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(VSFList_MouseDown)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
29        Err.Clear
End Sub

Private Function GetSel() As String
          Dim intRow As Integer
          Dim intCol As Integer
1         On Error GoTo GetSel_Error

2         With Me.VSFList
3             For intRow = 0 To .Rows - 1
4                 If .Cell(flexcpChecked, intRow, .ColIndex("编码"), intRow, .ColIndex("编码")) = 1 Then
5                     GetSel = GetSel & "," & .TextMatrix(intRow, .ColIndex("编码"))
6                 End If
7             Next
8         End With
9         If GetSel <> "" Then
10            GetSel = Mid(GetSel, 2)
11        End If


12        Exit Function
GetSel_Error:
13        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(GetSel)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
14        Err.Clear
End Function

Private Function SetSel(strItem As String) As String
          Dim intRow As Integer
          Dim intCol As Integer
1         On Error GoTo SetSel_Error

2         With Me.VSFList
3             For intRow = 0 To .Rows - 1
4                 If InStr("," & strItem & ",", "," & .TextMatrix(intRow, .ColIndex("编码")) & ",") > 0 Then
5                     .Cell(flexcpChecked, intRow, .ColIndex("编码"), intRow, .ColIndex("编码")) = 1
6                 End If
7             Next
8         End With


9         Exit Function
SetSel_Error:
10        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(SetSel)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
End Function
Private Sub ClsDefault()
          '功能           清空默认
          Dim intRow As Integer
1         On Error GoTo ClsDefault_Error

2         With Me.VSFList
3             For intRow = 1 To .Rows - 1
4                 .Cell(flexcpChecked, intRow, .ColIndex("默认"), intRow, .ColIndex("默认")) = 2
5             Next
6         End With


7         Exit Sub
ClsDefault_Error:
8         Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(ClsDefault)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
9         Err.Clear
End Sub
Private Function GetDefault() As String
          '功能           读到默认值
          Dim intRow As Integer
1         On Error GoTo GetDefault_Error

2         With Me.VSFList
3             For intRow = 0 To .Rows - 1
4                 If .Cell(flexcpChecked, intRow, .ColIndex("默认"), intRow, .ColIndex("默认")) = 1 Then
5                     GetDefault = .TextMatrix(intRow, .ColIndex("编码"))
6                 End If
7             Next
8         End With


9         Exit Function
GetDefault_Error:
10        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(GetDefault)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
End Function
Private Sub SetDefault(strNO As String)
          '功能       设置默认值
          Dim intRow As Integer
1         On Error GoTo SetDefault_Error

2         With Me.VSFList
3             For intRow = 0 To .Rows - 1
4                 If InStr("," & strNO & ",", "," & .TextMatrix(intRow, .ColIndex("编码")) & ",") > 0 Then
5                     .Cell(flexcpChecked, intRow, .ColIndex("默认"), intRow, .ColIndex("默认")) = 1
6                 End If
7             Next
8         End With


9         Exit Sub
SetDefault_Error:
10        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(SetDefault)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
End Sub


Private Function SaveData() As Boolean
          '功能       保存执行小组
          Dim strSQL As String
          Dim strItem As String
          Dim strDefault As String

1         On Error GoTo SaveData_Error

2         strItem = GetSel()
3         strDefault = GetDefault()
4         strSQL = "Zl_申请单执行小组_Edit(" & mlngID & ",'" & strItem & "','" & strDefault & "')"
5         ComExecuteProc Sel_Lis_DB, strSQL, "保存申请小组"
6         SaveDBLog 18, 6, 0, "编辑", "编辑申请单执行小组:" & strItem & ",默认小组:" & strDefault, 1012, "申请单设置"
7         mstrDept = strItem
8         SaveData = True


9         Exit Function
SaveData_Error:
10        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignDept", "执行(SaveData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
          
End Function
