VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCommProc 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsfCustomProc 
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _cx             =   4260
      _cy             =   2143
      Appearance      =   1
      BorderStyle     =   0
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   30
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
Attribute VB_Name = "frmCommProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnStartUp As Boolean
Public Event AfterSelect(ByVal strProc As String)
Private mlngCx As Long
Private mlngCy As Long
Private mfrmCommProcCode As frmCommProcCode

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    Call RefreshData
End Sub

Public Function ShowMe(ByVal objMain As Object, ByVal Cx As Long, ByVal Cy As Long)
    Me.Left = Cx
    Me.Top = Cy
    Me.Show 1, objMain
End Function

Private Sub Form_Load()
    mblnStartUp = True
    Call InitData
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print X
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsfCustomProc.Move 15, 15, Me.ScaleWidth - 30, Me.ScaleHeight - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmCommProcCode Is Nothing Then
        Unload mfrmCommProcCode
    End If
End Sub

Private Sub vsfCustomProc_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsfCustomProc.ColIndex("操作") Then Cancel = True
End Sub

Private Sub vsfCustomProc_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfCustomProc
        Select Case Col
        Case .ColIndex("操作")
            Dim rsData As ADODB.Recordset
            Dim strSQL As String
            Dim strName As String
            If mfrmCommProcCode Is Nothing Then
                Set mfrmCommProcCode = New frmCommProcCode
            End If
            strName = .TextMatrix(Row, .ColIndex("名称"))
            strSQL = "Select Text From All_Source Where (Owner = 'ZLTOOLS' Or Owner In (Select 所有者 From Zlsystems Where 编号=100)) And Type = 'FUNCTION' And Name = [1] Order By Line"
            Set rsData = OpenSQLRecord(strSQL, Me.Caption, UCase(strName))
            If rsData.BOF = False Then
                mfrmCommProcCode.ShowMe Me, rsData
            Else
                MsgBox "该函数没有找到代码！", vbInformation + vbOKOnly, "提示"
            End If
        End Select
    End With
End Sub

Private Sub vsfCustomProc_DblClick()
    If vsfCustomProc.Col = vsfCustomProc.ColIndex("操作") Then Exit Sub
    RaiseEvent AfterSelect(vsfCustomProc.TextMatrix(vsfCustomProc.Row, vsfCustomProc.ColIndex("名称")))
    Unload Me
End Sub

Private Sub vsfCustomProc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        RaiseEvent AfterSelect(vsfCustomProc.TextMatrix(vsfCustomProc.Row, vsfCustomProc.ColIndex("名称")))
        Unload Me
    Case vbKeyEscape
        Unload Me
    End Select
End Sub

Private Function InitData()
    On Error GoTo ErrHand
    With vsfCustomProc
        .Rows = 2
        .Cols = 3
        .ColWidth(0) = 1350
        .ExtendLastCol = True
        .AllowUserResizing = flexResizeColumns
        .TextMatrix(0, 0) = "名称"
        .TextMatrix(0, 1) = "说明"
        .TextMatrix(0, 2) = "操作"
        .ColKey(0) = "名称"
        .ColKey(1) = "说明"
        .ColKey(2) = "操作"
        .ColWidth(.ColIndex("名称")) = 1200
        .ColWidth(.ColIndex("说明")) = 4800
        .BackColorSel = &H8000000D
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("说明")
        .WordWrap = True
        .Editable = flexEDKbdMouse
         .ColComboList(.ColIndex("操作")) = "..."
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RefreshData()
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    strSQL = "Select 名称, 说明 From zlUsualFunc"
    Set rsData = OpenSQLRecord(strSQL, Me.Caption)
    If rsData.BOF = False Then
        With vsfCustomProc
            .Rows = 2
            For i = 1 To rsData.RecordCount
                If i + 1 > .Rows Then .Rows = .Rows + 1
                .TextMatrix(i, .ColIndex("名称")) = Nvl(rsData("名称").Value, "")
                .TextMatrix(i, .ColIndex("说明")) = Nvl(rsData("说明").Value, "")
                .TextMatrix(i, .ColIndex("操作")) = "查看代码"
                rsData.MoveNext
            Next
            .Row = 1
            .BackColorSel = &H8000000D
            .AutoSizeMode = flexAutoSizeRowHeight
            .AutoSize .ColIndex("说明")
            .WordWrap = True
        End With
    End If
End Function
