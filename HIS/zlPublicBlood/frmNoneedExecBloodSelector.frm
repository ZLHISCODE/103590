VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmNoneedExecBloodSelector 
   Caption         =   "无需执行血液选择"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5850
   Icon            =   "frmNoneedExecBloodSelector.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   5850
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdEdit 
      Caption         =   "填写(&E)"
      Height          =   855
      Left            =   5040
      TabIndex        =   6
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   330
      Left            =   4800
      TabIndex        =   5
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&S)"
      Default         =   -1  'True
      Height          =   330
      Left            =   3720
      TabIndex        =   4
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtReson 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3360
      Width           =   4860
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExec 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5610
      _cx             =   9895
      _cy             =   4154
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
      BackColorSel    =   16444122
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmNoneedExecBloodSelector.frx":6852
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "统一填写已勾选血液的更改原因："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3135
      Width           =   2700
   End
   Begin VB.Label lbl 
      Caption         =   "输注过程中如果出现输血反应,可将该医嘱尚未执行完成的血液更改为无需执行状态"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3690
   End
End
Attribute VB_Name = "frmNoneedExecBloodSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnShow As Boolean
Private mblnFinish As Boolean
Private mclsVsf As clsVsf
Private mlng医嘱ID As Long
Private mrsTmp As New ADODB.Recordset

Public Function ShowMe(ByVal frmParent As Object, ByVal lng医嘱ID As Long, Optional blnFinish As Boolean) As Boolean
    
    mblnOK = False
    mblnFinish = False

    mlng医嘱ID = lng医嘱ID
    mblnShow = False
    On Error Resume Next
    Me.Show 1, frmParent
    If Err <> 0 Then Err.Clear
    blnFinish = mblnFinish
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    Dim i As Integer
    
    With vsExec
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked Then
                .TextMatrix(i, .ColIndex("执行摘要")) = Trim(txtReson.Text)
            End If
        Next
    End With
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim mrsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim blnUnload As Boolean
    
    On Error GoTo ErrHand
    '初始化表格
    Call InitTable
    
    '加载血液执行情况
    Call LoadExecVsf
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitTable()
'表格初始化
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsExec, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTBoolean, "", "选择", False)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "ID", , True, False, False, True)  '收发ID
        Call .AppendColumn("状态", 800, flexAlignLeftCenter, flexDTString, , "血液状态") '接收执行状态
        Call .AppendColumn("血袋编号", 1200, flexAlignLeftCenter, flexDTString, , "血袋编号", , False, False)
        Call .AppendColumn("血液名称", 1300, flexAlignLeftCenter, flexDTString, , "血液名称", , False, False)
        Call .AppendColumn("更改原因", 1300, flexAlignLeftCenter, flexDTString, , "执行摘要", , False, False)
        .AppendRows = False
    End With
        vsExec.ExplorerBar = flexExNone
End Sub

Private Sub LoadExecVsf()
    '功能：加载血液执行情况
    Dim i As Integer, intRow As Integer
    Dim arrName, arrKey, arrColWidth
    Dim strSQL As String
    

    On Error GoTo ErrHand
    strSQL = "SELECT a.Id, decode(h.执行状态,4,1,0) 选择, a.血袋编号, h.执行状态, h.接收状态, h.执行摘要," & vbNewLine & _
                "       Decode(Nvl(h.执行状态, 0), 0, '已接收', 4, '无需执行') 血液状态, e.名称 AS 血液名称" & vbNewLine & _
                "FROM 收费项目目录 e, 血液品种 k, 血液规格 l, 血液收发记录 a, 血液发送记录 h, 血液配血记录 b" & vbNewLine & _
                "WHERE e.Id = a.血液id AND k.品种id = l.品种id AND l.规格id = a.血液id AND a.Id = h.收发id AND h.配发id = b.Id AND h.接收状态 = 1 AND" & vbNewLine & _
                "      h.执行状态 IN (0, 4) AND Nvl(a.发血状态, 0) = 2 AND b.申请id = [1]" & vbNewLine & _
                "ORDER BY h.执行分组, a.配血日期, a.序号"
    Set mrsTmp = gobjDatabase.OpenSQLRecord(strSQL, "提取可无需执行的血液", mlng医嘱ID)
    If mrsTmp.RecordCount = 0 Then
        Call MsgBox("无已接收或无需执行的血液！", vbInformation, "中联软件")
        Unload Me
        Exit Sub
    End If
    Call mclsVsf.LoadGrid(mrsTmp, "", True)
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean, strSQL As String
    Dim arrSQL, i As Integer
    Dim intStatus As Integer
    Dim rsData As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    With vsExec
    For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("选择")) = 1 And .TextMatrix(i, .ColIndex("执行摘要")) = "" Then
            MsgBox "第" & i & "行血液未填写更改原因，请填写后再进行保存！", vbInformation, "中联软件"
            Exit Function
        End If
    Next
    arrSQL = Array()
        For i = 1 To .Rows - 1
            mrsTmp.Filter = "ID = " & .TextMatrix(i, .ColIndex("ID"))
            intStatus = mrsTmp!执行状态
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = vbChecked And .TextMatrix(i, .ColIndex("血液状态")) <> "无需执行" Then
                intStatus = 4
            ElseIf .Cell(flexcpChecked, i, .ColIndex("选择")) = 2 And .TextMatrix(i, .ColIndex("血液状态")) = "无需执行" Then
                strSQL = _
                    " SELECT 1 FROM 血液执行记录 WHERE 收发id = [1] AND ROWNUM < 2"
                Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "血液执行记录提取", .TextMatrix(i, .ColIndex("ID")))
                If rsData.RecordCount > 0 Then
                    intStatus = 1
                Else
                    intStatus = 0
                End If
            End If
            If mrsTmp.RecordCount <> 0 Then
                If intStatus <> mrsTmp!执行状态 Or .TextMatrix(i, .ColIndex("执行摘要")) <> mrsTmp!执行摘要 & "" Then
                    strSQL = "Zl_血液发送记录_无需执行(" & .TextMatrix(i, .ColIndex("ID")) & "," & intStatus & "," & "'" & _
                                .TextMatrix(i, .ColIndex("执行摘要")) & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                End If
            End If
            mrsTmp.Filter = ""
        Next
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            If CStr(arrSQL(i)) <> "" Then
                Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            End If
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End With
    SaveData = True
    Exit Function
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub txtReson_KeyPress(KeyAscii As Integer)
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsExec_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsExec.ColIndex("选择") And Col <> vsExec.ColIndex("执行摘要") Then Cancel = True
    If (Col = vsExec.ColIndex("执行摘要") And vsExec.Cell(flexcpChecked, Row, vsExec.ColIndex("选择")) = 2) Then Cancel = True
End Sub

Private Sub vsExec_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(1, "'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsExec_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsExec.ColIndex("选择") And vsExec.Cell(flexcpChecked, Row, vsExec.ColIndex("选择")) = vbChecked Then
        vsExec.TextMatrix(Row, vsExec.ColIndex("执行摘要")) = ""
    End If
End Sub
