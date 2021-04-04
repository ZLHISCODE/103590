VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmExtraFeemove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医嘱附费转移"
   ClientHeight    =   3504
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6228
   Icon            =   "frmExtraFeemove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmExtraFeemove.frx":058A
   ScaleHeight     =   3504
   ScaleWidth      =   6228
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   760
      Left            =   0
      ScaleHeight     =   756
      ScaleWidth      =   6228
      TabIndex        =   4
      Top             =   0
      Width           =   6225
      Begin VB.Image imgInfo 
         Height          =   576
         Left            =   120
         Picture         =   "frmExtraFeemove.frx":0B14
         Top             =   0
         Width           =   576
      End
      Begin VB.Label lblNote 
         BackColor       =   &H80000005&
         Caption         =   "当前选择的费用单据:XXXXYYYY,将转移关联到下列指定的医嘱上,请选择一行待关联的医嘱。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   5
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   612
      ScaleWidth      =   6228
      TabIndex        =   0
      Top             =   2892
      Width           =   6225
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   3950
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5070
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Bindings        =   "frmExtraFeemove.frx":2656
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      Top             =   795
      Width           =   6195
      _cx             =   10927
      _cy             =   3625
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmExtraFeemove.frx":266A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
Attribute VB_Name = "frmExtraFeemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng医嘱ID As Long      '附费关联的医嘱ID(主ID,附费可能是关联到以该ID为相关ID的部位或方法的医嘱ID上的)
Private mstrNO As String        '附费单据号
Private mint记录性质 As Integer '附费记录性质
Private mint费用性质 As Integer '1=门诊费用(包括门诊记帐)，2-住院费用
Private mblnOK As Boolean

Private Const col_NO = 0
Private Const col_医嘱内容 = 1
Private Const col_发送时间 = 2


Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页Id As Long, ByVal str挂号单 As String, _
    ByVal lng医嘱ID As Long, ByVal str诊疗类别 As String, ByVal lng执行部门id As Long, _
    ByVal strNO As String, ByVal int记录性质 As Integer, ByVal int费用性质 As Integer)
    
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    '1.本次挂号或本次住院的同一病人的医嘱
    '2.与当前附费关联医嘱相同类别和执行科室的医嘱
    '3.如果以前是关联到检查方法或部位上的，转移后关联到新医嘱的主记录上
    strSQL = "Select b.No, b.医嘱id, b.发送号, To_Char(b.发送时间,'YYYY-MM-DD HH24:MI') as 发送时间, a.医嘱内容" & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱发送 B" & vbNewLine & _
            "Where a.Id = b.医嘱id And a.相关id Is Null And a.诊疗类别 = [5] And b.执行部门id = [6]" & vbNewLine & _
            "      And a.id <> [4] And a.病人id = [1]" & _
            IIf(str挂号单 <> "", " And a.挂号单 = [2]", " And a.主页ID = [3]") & " Order by NO"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "附费转移", lng病人ID, str挂号单, lng主页Id, lng医嘱ID, str诊疗类别, lng执行部门id)
    If rsTmp.RecordCount = 0 Then
        MsgBox "该病人在本科室没有相同类别的医嘱，不能进行附费转移。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mlng医嘱ID = lng医嘱ID
    mstrNO = strNO
    mint记录性质 = int记录性质
    mint费用性质 = int费用性质
    
    
    lblNote.Caption = Replace(lblNote.Caption, "XXXXYYYY", strNO)
    Call LoadList(rsTmp)
    
    mblnOK = False
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    
    ShowMe = mblnOK
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub LoadList(ByRef rsList As ADODB.Recordset)
'功能：加载医嘱发送清单
    Dim i As Long
    
    With vsList
        .Rows = .FixedRows
        .Rows = .FixedRows + rsList.RecordCount
        
        For i = 1 To rsList.RecordCount
            .TextMatrix(i, col_NO) = rsList!NO
                        
            .TextMatrix(i, col_医嘱内容) = rsList!医嘱内容
            .Cell(flexcpData, i, col_医嘱内容) = Val(rsList!医嘱ID)
            
            .TextMatrix(i, col_发送时间) = rsList!发送时间
            .Cell(flexcpData, i, col_发送时间) = Val(rsList!发送号)
            
            rsList.MoveNext
        Next
        If .Rows = .FixedRows + 1 Then
            .Row = .Rows - 1
        Else
            .Row = 0 '缺省不选择任何一行
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim strSQL As String
    With vsList
        If .Row <= .FixedRows - 1 Then
            MsgBox "请选择一行待关联的医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你确定要将医嘱附费" & mstrNO & "关联到" & vbCrLf & "医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """(" & _
                    .TextMatrix(.Row, col_NO) & ")上吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        lng医嘱ID = Val(.Cell(flexcpData, .Row, col_医嘱内容))
        lng发送号 = Val(.Cell(flexcpData, .Row, col_发送时间))
        
        strSQL = "Zl_病人医嘱附费_Move(" & mint记录性质 & ",'" & mstrNO & "'," & mint费用性质 & "," & lng医嘱ID & "," & lng发送号 & ")"
        Call gobjDatabase.ExecuteProcedure(strSQL, "附费转移")
    End With
    
    mblnOK = True
    Unload Me
End Sub
