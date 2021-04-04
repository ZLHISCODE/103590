VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDiseaseReportRela 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病报告对应信息设置"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmDiseaseReportRela.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4425
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2820
      TabIndex        =   2
      Top             =   4515
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1725
      TabIndex        =   1
      Top             =   4515
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRela 
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4200
      _cx             =   7408
      _cy             =   6456
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDiseaseReportRela.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   Begin VB.Label Label1 
      Caption         =   "设置疾病申报信息项目与疾病报告中的诊治要素之间的对应关系，以便接收时从报告中自动提取这些项目信息。"
      Height          =   525
      Left            =   165
      TabIndex        =   3
      Top             =   105
      Width           =   4185
   End
End
Attribute VB_Name = "frmDiseaseReportRela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strRela As String, i As Long
    
    With vsRela
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                If Trim(.TextMatrix(i, 1)) = "" Or .TextMatrix(i, 1) = "请输入对应临时诊治要素的名称" Then
                    .Row = i: .Col = 1: Call .ShowCell(i, 1)
                    MsgBox "请输入""" & .TextMatrix(i, 0) & """对应的临时诊治要素的名称。", vbInformation, gstrSysName
                    vsRela.SetFocus: Exit Sub
                Else
                    strRela = strRela & "|" & .TextMatrix(i, 0) & "," & .TextMatrix(i, 1)
                End If
            End If
        Next
    End With
    
    If strRela <> "" Then
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure("Zl_疾病申报对应_Update('" & Mid(strRela, 2) & "')", Me.Caption)
    End If
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    strSQL = "Select 申报项目,对应要素 From 疾病申报对应"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With vsRela
        .Row = 1: .Col = 1
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                .Cell(flexcpForeColor, i, 0, i, 1) = &HC0C0C0
            Else
                .RowData(i) = 1
                If .Row = 1 Then
                    .Row = i: Call .ShowCell(i, 1)
                End If
                
                rsTmp.Filter = "申报项目='" & .TextMatrix(i, 0) & "'"
                If Not rsTmp.EOF Then .TextMatrix(i, 1) = NVL(rsTmp!对应要素)
                
                If .TextMatrix(i, 1) = "" Then
                    .TextMatrix(i, 1) = "请输入对应临时诊治要素的名称"
                End If
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsRela_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsRela
        If .RowData(NewRow) = 1 And NewCol = 1 Then
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
        End If
    End With
End Sub

Private Sub vsRela_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsRela_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsRela.EditSelStart = 0
    vsRela.EditSelLength = zlCommFun.ActualLen(vsRela.EditText)
End Sub

Private Sub vsRela_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRela
        If Not (.RowData(Row) = 1 And Col = 1) Then
            Cancel = True
        End If
    End With
End Sub
