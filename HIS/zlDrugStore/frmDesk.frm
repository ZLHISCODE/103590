VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmDesk 
   Caption         =   "工作台设置"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4845
   Icon            =   "frmDesk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4845
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增(&A)"
      Height          =   350
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDesk 
      Height          =   2745
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3120
      _cx             =   5503
      _cy             =   4842
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDesk.frx":6852
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
Attribute VB_Name = "frmDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlng部门id As Long

Private Sub cmdAdd_Click()
    With vsfDesk
        If .TextMatrix(.Row, .ColIndex("配药台")) = "" Then
            MsgBox "请将当前配药台编辑好在进行新增！", vbInformation, gstrSysName
            Exit Sub
        End If
        .rows = .rows + 1
        .Row = .rows - 1
        .TextMatrix(.Row, .ColIndex("序号")) = .rows - 1
    End With
End Sub

Private Sub CmdCancle_Click()
    Unload Me
End Sub

Private Sub InitVSF()
    Dim strsql As String
    Dim rstemp As Recordset

    On Error GoTo errHandle

    strsql = "select id,名称 from 配液台 where 部门id=[1]"
    Set rstemp = zldatabase.OpenSQLRecord(strsql, "", mlng部门id)
    
    If Not rstemp.EOF Then
        With vsfDesk
            .rows = 1
            Do While Not rstemp.EOF
                .rows = .rows + 1
                .TextMatrix(.rows - 1, .ColIndex("序号")) = .rows - 1
                .TextMatrix(.rows - 1, .ColIndex("配药台")) = rstemp!名称
                .TextMatrix(.rows - 1, .ColIndex("id")) = rstemp!Id
                rstemp.MoveNext
            Loop
        End With
    Else
        Me.vsfDesk.rows = 2
        vsfDesk.TextMatrix(vsfDesk.Row, vsfDesk.ColIndex("序号")) = vsfDesk.rows - 1
    End If

    Exit Sub

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowMe(ByVal lng部门ID As Long, ByVal frmobjct As Object)
    mlng部门id = lng部门ID
    Me.Show 1, frmobjct
End Sub

Private Sub cmdDel_Click()
    Dim strsql As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    If vsfDesk.Row = 0 Then Exit Sub
    If vsfDesk.TextMatrix(vsfDesk.Row, vsfDesk.ColIndex("id")) <> "" Then
        strsql = "Zl_配液台_删除(" & vsfDesk.TextMatrix(vsfDesk.Row, vsfDesk.ColIndex("id")) & ")"
        
        Call zldatabase.ExecuteProcedure(strsql, "cmdDel_Click")
    End If
    
    vsfDesk.RemoveItem (vsfDesk.Row)
    
    For i = 1 To vsfDesk.rows - 1
        vsfDesk.TextMatrix(i, vsfDesk.ColIndex("序号")) = i
    Next
    Exit Sub
errHandle:
        If ErrCenter() Then
            Resume
        End If
        Call SaveErrLog
End Sub

Private Sub cmdSave_Click()
    Dim strsql As String
    Dim arrSql As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    Dim strMsg As String
    
    arrSql = Array()
    On Error GoTo errHandle
        For i = 1 To vsfDesk.rows - 1
            If vsfDesk.TextMatrix(i, vsfDesk.ColIndex("配药台")) <> "" Then
                strsql = "Zl_配液台_设置("
                strsql = strsql & mlng部门id & ","
                strsql = strsql & "'" & vsfDesk.TextMatrix(i, vsfDesk.ColIndex("配药台")) & "',"
                strsql = strsql & Val(vsfDesk.TextMatrix(i, vsfDesk.ColIndex("id"))) & ")"
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
        
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Me.vsfDesk.EditMaxLength = 20
    Call InitVSF
End Sub

Private Sub vsfDesk_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Me.vsfDesk.TextMatrix(0, Col) <> "配药台" Then Cancel = True
    
End Sub
