VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "排班设置"
   ClientHeight    =   6684
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11280
   Icon            =   "frmPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6684
   ScaleWidth      =   11280
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd生成 
      Caption         =   "生成新的排班(&D)"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Width           =   1932
   End
   Begin VB.Frame fraCon 
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton cmdCur 
         Caption         =   "当天(&C)"
         Height          =   350
         Left            =   2760
         TabIndex        =   8
         ToolTipText     =   "热键：F2"
         Top             =   215
         Width           =   1215
      End
      Begin VB.CommandButton cmdRe 
         Caption         =   "刷新&R)"
         Height          =   350
         Left            =   9360
         TabIndex        =   7
         ToolTipText     =   "热键：F2"
         Top             =   215
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "下一天(&N)"
         Height          =   350
         Left            =   7160
         TabIndex        =   6
         ToolTipText     =   "热键：F2"
         Top             =   215
         Width           =   1215
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "上一天(&L)"
         Height          =   350
         Left            =   4960
         TabIndex        =   5
         ToolTipText     =   "热键：F2"
         Top             =   215
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker Dtp开始时间 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         _ExtentX        =   3620
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   112066563
         CurrentDate     =   39998
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   2
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8400
      TabIndex        =   1
      Top             =   6240
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
      Height          =   5145
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10800
      _cx             =   19050
      _cy             =   9075
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
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPlan.frx":6852
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
      Editable        =   0
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
Attribute VB_Name = "frmPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng部门id As Long
Private mstr批次 As String
Private mdateCur As Date
Private mrs配液台 As Recordset
Private mrs批次 As Recordset
Private mrs人员 As New Recordset
Private mbln保存 As Boolean

Private mintRow As Integer
Private mintCol As Integer

Private Sub CheckDate(ByVal Row As Long, ByVal Col As Long)
    With Me.vsfPlan
        If Col = .ColIndex("摆药人") Or Col = .ColIndex("配液人") Or Col = .ColIndex("核对人") Or Col = .ColIndex("复核人") Or Col = .ColIndex("审核人") Then
            If Not .TextMatrix(Row, Col) = "" Then
                mrs人员.Filter = "姓名 = '" & .TextMatrix(Row, Col) & "'"
                If mrs人员.RecordCount = 0 Then
                    mrs人员.Filter = "简码 = '" & UCase(.TextMatrix(Row, Col)) & "'"
                    If mrs人员.RecordCount = 0 Then
                        MsgBox "未匹配到相关人员,请重新输入"
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    Else
                        .TextMatrix(Row, Col) = mrs人员!姓名
                        If Col <> .Cols - 1 Then .Col = .Col + 1
                    End If
                Else
                    If Col <> .Cols - 1 Then .Col = .Col + 1
                End If
            End If
        End If
    End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub InitCom()
    Dim strPreson As String
    Dim strsql As String
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    mdateCur = zldatabase.Currentdate
    Dtp开始时间.Value = mdateCur
    
    strsql = "Select a.Id, a.姓名, a.简码" & vbNewLine & _
            "From 人员表 A, 部门人员 B" & vbNewLine & _
            "Where a.Id = b.人员id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And b.部门id =[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "initVSF", mlng部门id)
    
    Set mrs人员 = rsTemp
    
    Do While Not rsTemp.EOF
        strPreson = strPreson & IIf(strPreson = "", "|", "") & rsTemp!姓名 & "|"
        rsTemp.MoveNext
    Loop
    
    With vsfPlan
        .ColComboList(.ColIndex("摆药人")) = strPreson
        .ColComboList(.ColIndex("配液人")) = strPreson
        .ColComboList(.ColIndex("核对人")) = strPreson
        .ColComboList(.ColIndex("复核人")) = strPreson
        .ColComboList(.ColIndex("审核人")) = strPreson
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub InitVSF()
    Dim rsTemp As Recordset
    Dim strsql As String
    Dim i As Integer
    Dim j As Integer
    Dim count As Integer
 
    strsql = "select A.配药台id,B.名称,A.批次,A.审核人,A.摆药人,A.核对人,A.配液人,A.复核人 from 配液工作安排 A,配液台 B where  A.配药台id=B.id and  A.部门id= B.部门id and A.部门id=[1] and A.日期=[2]"
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "initVSF", mlng部门id, CDate(Format(Dtp开始时间.Value, "Short Date")))
    
    With Me.vsfPlan
        
        If rsTemp.EOF Then
            Me.vsfPlan.Editable = flexEDKbdMouse
            .rows = (mrs配液台.RecordCount * mrs批次.RecordCount) + 1
            mrs配液台.MoveFirst
            For i = 1 To mrs配液台.RecordCount
                mrs批次.MoveFirst
                For j = 1 To mrs批次.RecordCount
                    count = count + 1
                    .TextMatrix(count, .ColIndex("配液台号")) = mrs配液台!名称
                    .TextMatrix(count, .ColIndex("配液台id")) = mrs配液台!Id
                    .TextMatrix(count, .ColIndex("配药批次")) = mrs批次!批次
                    .TextMatrix(count, .ColIndex("摆药人")) = ""
                    .TextMatrix(count, .ColIndex("配液人")) = ""
                    .TextMatrix(count, .ColIndex("核对人")) = ""
                    .TextMatrix(count, .ColIndex("复核人")) = ""
                    .TextMatrix(count, .ColIndex("审核人")) = ""
                    mrs批次.MoveNext
                Next
                mrs配液台.MoveNext
            Next
            
            Exit Sub
        End If
        Me.vsfPlan.Editable = flexEDNone
        Do While Not rsTemp.EOF
            i = i + 1
            .rows = i + 1
            .TextMatrix(i, .ColIndex("配液台号")) = rsTemp!名称
            .TextMatrix(i, .ColIndex("配液台id")) = rsTemp!配药台id
            .TextMatrix(i, .ColIndex("配药批次")) = rsTemp!批次
            .TextMatrix(i, .ColIndex("摆药人")) = NVL(rsTemp!摆药人)
            .TextMatrix(i, .ColIndex("配液人")) = NVL(rsTemp!配液人)
            .TextMatrix(i, .ColIndex("核对人")) = NVL(rsTemp!核对人)
            .TextMatrix(i, .ColIndex("复核人")) = NVL(rsTemp!复核人)
            .TextMatrix(i, .ColIndex("审核人")) = NVL(rsTemp!审核人)

            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowMe(ByVal lng部门id As Long, ByVal frmParent As Object)
    
    Dim rsTemp As Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    strsql = "select id,名称 from 配液台 where 部门id=[1]"
    Set mrs配液台 = zldatabase.OpenSQLRecord(strsql, "initVSF", lng部门id)
    
    If mrs配液台.EOF Then
        MsgBox "请先进行配液台进行设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strsql = "select 批次 from 配药工作批次 where 配置中心id=[1]"
    Set mrs批次 = zldatabase.OpenSQLRecord(strsql, "initVSF", lng部门id)
    
    If mrs批次.EOF Then
        MsgBox "请先进行工作批次进行设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mlng部门id = lng部门id
    Me.Show 1, frmParent
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCur_Click()
    Dtp开始时间.Value = mdateCur
    Call cmdRe_Click
End Sub

Private Sub cmdLast_Click()
    Dtp开始时间.Value = Dtp开始时间.Value - 1
    Call cmdRe_Click
End Sub

Private Sub cmdNext_Click()
    Dtp开始时间.Value = Dtp开始时间.Value + 1
    Call cmdRe_Click
End Sub

Private Sub cmdRe_Click()
    If mbln保存 = False Then
        If MsgBox("当前数据还未保存，是否继续当前操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    Call InitVSF
    
    Me.vsfPlan.Editable = (Me.Dtp开始时间.Value >= mdateCur)
    cmdSave.Enabled = (Me.Dtp开始时间.Value >= mdateCur)
End Sub

Private Sub cmdSave_Click()
    Dim strsql As String
    Dim i  As Integer
    Dim arrSql As Variant
    Dim j As Integer
    
    arrSql = Array()
    With Me.vsfPlan
        For i = 1 To .rows - 1
            
            If Val(.TextMatrix(i, .ColIndex("配液台id"))) = 0 Then
                Exit Sub
            End If
            strsql = "Zl_配液工作安排_设置("
            strsql = strsql & mlng部门id
            strsql = strsql & ",to_date('" & Format(Dtp开始时间.Value, "Short Date") & "' ,'yyyy-mm-dd')"
            strsql = strsql & "," & Val(.TextMatrix(i, .ColIndex("配液台id")))
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("配药批次")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("审核人")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("摆药人")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("核对人")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("配液人")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("复核人")) & "'"
            strsql = strsql & "," & i
            strsql = strsql & ")"

            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strsql
        Next
    End With
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "CmdSave_Click")
    Next
    gcnOracle.CommitTrans
    mbln保存 = True
    If MsgBox("保存成功，是否继续排班？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
        Unload Me
    End If
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd生成_Click()
    Call cmdRe_Click
    frmPlanCopy.ShowCard Me, Dtp开始时间.Value, mlng部门id
End Sub

Private Sub Form_Load()
    mbln保存 = True
    Call InitCom
End Sub

Private Sub vsfPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call CheckDate(Row, Col)
End Sub

Private Sub vsfPlan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    
    With Me.vsfPlan
        If KeyAscii = 13 Then
            If Col = .ColIndex("配液台号") Or Col = .ColIndex("配液台id") Or Col = .ColIndex("配药批次") Then
                .Col = Col + 1
            Else
                Call CheckDate(Row, Col)
            End If
        End If
    End With
End Sub
