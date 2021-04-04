VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmDeskMedi 
   Caption         =   "配液台药品对照"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6210
   Icon            =   "frmDeskMedi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6210
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4800
      TabIndex        =   3
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增(&A)"
      Height          =   350
      Left            =   1920
      TabIndex        =   2
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "删除(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   1
      Top             =   4800
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5760
      _cx             =   10160
      _cy             =   8017
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16771280
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDeskMedi.frx":6852
      ScrollTrack     =   0   'False
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
End
Attribute VB_Name = "frmDeskMedi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlng部门id As Long
Private mstr配液台 As String

Private Sub InitVSF()
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    strsql = "select A.药品id,A.配药台id, B.名称 配药台名称,C.名称 药品名称,c.编码 from 配液台药品对照 A,配液台 B,收费项目目录 C  where A.配药台id=B.id and A.药品id =C.id and A.部门id=[1]"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "InitVsf", mlng部门id)
    
    With Me.vsfPlan
        .rows = 1
        .ColComboList(.ColIndex("配药台号")) = mstr配液台
'        .ColComboList(.ColIndex("药品名称")) = "..."
        
        If Not rstemp.EOF Then
            Do While Not rstemp.EOF
                .rows = .rows + 1
                .TextMatrix(.rows - 1, .ColIndex("序号")) = .rows - 1
                .TextMatrix(.rows - 1, .ColIndex("药品id")) = rstemp!药品ID
                .TextMatrix(.rows - 1, .ColIndex("配药台id")) = rstemp!配药台id
                .TextMatrix(.rows - 1, .ColIndex("配药台号")) = rstemp!配药台名称
                .TextMatrix(.rows - 1, .ColIndex("药品名称")) = "(" & rstemp!编码 & ")" & rstemp!药品名称
                
                rstemp.MoveNext
            Loop
        Else
            .rows = 2
            .TextMatrix(.rows - 1, .ColIndex("序号")) = .rows - 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub ShowMe(ByVal lng部门ID As Long, ByVal frmObject As Object)
    Dim strsql As String
    Dim rstemp As Recordset
    
    mlng部门id = lng部门ID
    strsql = "select id,名称 from 配液台 where 部门id=[1]"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "", mlng部门id)
    
    If rstemp.EOF Then
        MsgBox "请先进行配液台进行设置！", vbInformation, gstrSysName
        Exit Sub
    Else
        Do While Not rstemp.EOF
            mstr配液台 = mstr配液台 & "#" & rstemp!Id & ";" & rstemp!名称 & "|"
            rstemp.MoveNext
        Loop
    End If
    
    Me.Show 1, frmObject
End Sub

Private Sub cmdAdd_Click()
     With vsfPlan
        If .TextMatrix(.Row, .ColIndex("配药台号")) = "" Or .TextMatrix(.Row, .ColIndex("配药台id")) = "" Or .TextMatrix(.Row, .ColIndex("药品id")) = "" Or .TextMatrix(.Row, .ColIndex("药品名称")) = "" Then
            MsgBox "请将当前行信息编辑好再进行新增！", vbInformation, gstrSysName
            Exit Sub
        End If
        .rows = .rows + 1
        .Row = .rows - 1
        .TextMatrix(.Row, .ColIndex("序号")) = .rows - 1
    End With
End Sub

Private Sub CmdCancle_Click()
    Dim strsql As String
    
    On Error GoTo errHandle
    
    If vsfPlan.Row = 0 Then Exit Sub
    
    strsql = "Zl_配液台药品对照_删除("
    strsql = strsql & Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("药品id"))) & ","
    strsql = strsql & Val(vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("配药台id"))) & ")"
    
    Call zldatabase.ExecuteProcedure(strsql, "CmdCancle_Click")
    vsfPlan.RemoveItem (vsfPlan.Row)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSave_Click()
    Dim strsql As String
    Dim arrSql As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    Dim strMsg As String
    
    arrSql = Array()
    On Error GoTo errHandle
    With vsfPlan
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("配药台号")) <> "" And .TextMatrix(i, .ColIndex("配药台id")) <> "" And .TextMatrix(i, .ColIndex("药品id")) <> "" And .TextMatrix(i, .ColIndex("药品名称")) <> "" Then
                strsql = "Zl_配液台药品对照_设置("
                strsql = strsql & "" & .TextMatrix(i, .ColIndex("药品id")) & ","
                strsql = strsql & .TextMatrix(i, .ColIndex("配药台id")) & ","
                strsql = strsql & mlng部门id & ","
                strsql = strsql & i - 1 & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strsql
            Else
                strMsg = "有数据尚未编辑完整，是否继续？"
            End If
        Next
        
        If strMsg <> "" Then
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        gcnOracle.BeginTrans
        blnBeginTrans = True
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "CmdSave_Click")
        Next
        gcnOracle.CommitTrans
        blnBeginTrans = False
        
        Unload Me
    End With
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call InitVSF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr配液台 = ""
End Sub

Private Sub vsfPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfPlan.TextMatrix(0, Col) = "序号" Then Cancel = True
End Sub

Private Sub vsfPlan_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    Me.vsfPlan.TextMatrix(Row, vsfPlan.ColIndex("配药台id")) = vsfPlan.ComboData(vsfPlan.ComboIndex)
End Sub

Private Sub vsfPlan_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    vsfPlan.TextMatrix(Row, vsfPlan.ColIndex("药品id")) = ""
    If KeyCode = 13 Then
        
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(2, "静脉配置中心", mlng部门id, mlng部门id)
        End If
        
        If vsfPlan.EditText = "" Then
            Set RecReturn = frmSelector.ShowMe(Me, 0, 1, , , , mlng部门id, , , 0, True, True, True, , , mstrPrivs)
        Else
            Set RecReturn = frmSelector.ShowMe(Me, 1, 1, vsfPlan.EditText, Me.vsfPlan.Left, Me.vsfPlan.Top, mlng部门id, , , 0, True, True, True, , , mstrPrivs)
        End If
        
        If Not RecReturn.EOF Then
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.Col) = "(" & RecReturn!药品编码 & ")" & RecReturn!通用名
            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("药品id")) = RecReturn!药品ID
        End If
    End If
End Sub

'Private Sub vsfPlan_KeyPress(KeyAscii As Integer)
'    vsfPlan.TextMatrix(Row, vsfPlan.ColIndex("药品id")) = ""
'    If KeyAscii = 13 Then
'        If grsMaster.State = adStateClosed Then
'            Call SetSelectorRS(2, "静脉配置中心", mlng部门id, mlng部门id)
'        End If
'
'        If vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.Col) = "" Then
'            Set RecReturn = frmSelector.ShowMe(Me, 0, 1, , , , mlng部门id, , , 0, True, True, True, , , mstrPrivs)
'        Else
'            Set RecReturn = frmSelector.ShowMe(Me, 1, 1, vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.Col), Me.vsfPlan.Left, Me.vsfPlan.Top, mlng部门id, , , 0, True, True, True, , , mstrPrivs)
'        End If
'
'        If Not RecReturn.EOF Then
'            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.Col) = "(" & RecReturn!药品编码 & ")" & RecReturn!通用名
'            vsfPlan.TextMatrix(vsfPlan.Row, vsfPlan.ColIndex("药品id")) = RecReturn!药品ID
'        End If
'    End If
'End Sub


