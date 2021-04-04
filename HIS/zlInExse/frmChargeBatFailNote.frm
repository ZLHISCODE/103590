VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmChargeBatFailNote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "未记帐病人列表"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   Icon            =   "frmChargeBatFailNote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfPati 
      Height          =   5520
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   7710
      _cx             =   13600
      _cy             =   9737
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmChargeBatFailNote.frx":076A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   0   'False
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
Attribute VB_Name = "frmChargeBatFailNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPatis As String

Public Sub ShowMe(frmMain As Object, ByVal strPatis As String)
    mstrPatis = strPatis
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
    Dim i As Integer, strSQL As String
    Dim arrPati() As String, rsPati As ADODB.Recordset
    On Error GoTo errH
    vsfPati.Clear 1
    vsfPati.Rows = 1
    arrPati = Split(mstrPatis, ",")
    For i = 0 To UBound(arrPati)
        With vsfPati
            strSQL = "Select 姓名,住院号,年龄,当前床号,性别 From 病人信息 Where 病人ID = [1]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Split(arrPati(i), "|")(0)))
            If Not rsPati.EOF Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("姓名")) = NVL(rsPati!姓名)
                .TextMatrix(.Rows - 1, .ColIndex("住院号")) = NVL(rsPati!住院号)
                .TextMatrix(.Rows - 1, .ColIndex("年龄")) = NVL(rsPati!年龄)
                .TextMatrix(.Rows - 1, .ColIndex("性别")) = NVL(rsPati!性别)
                .TextMatrix(.Rows - 1, .ColIndex("床号")) = NVL(rsPati!当前床号)
                .TextMatrix(.Rows - 1, .ColIndex("未记帐原因")) = Split(arrPati(i), "|")(1)
            End If
        End With
    Next i
    vsfPati.AutoSize 0, vsfPati.Cols - 1
    
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub
