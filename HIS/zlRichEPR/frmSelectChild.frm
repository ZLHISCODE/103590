VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSelectChild 
   BorderStyle     =   0  'None
   ClientHeight    =   3180
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2325
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   4305
      _cx             =   7594
      _cy             =   4101
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   8388608
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   -2147483624
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSelectChild.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "È¡Ïû(&C)"
      Height          =   350
      Left            =   315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Width           =   1100
   End
End
Attribute VB_Name = "frmSelectChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public blnOK As Boolean
Public strReturn As String

Private blnFirst As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If blnFirst = False Then Exit Sub
    blnFirst = False
    If vfgThis.Visible And vfgThis.Enabled Then vfgThis.SetFocus
End Sub

Private Sub Form_Load()
    blnOK = False
    blnFirst = True
    vfgThis.Visible = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With vfgThis
        .Left = Screen.TwipsPerPixelX
        .Top = Screen.TwipsPerPixelY
        .Width = Me.ScaleWidth - Screen.TwipsPerPixelX * 2
        .Height = Me.ScaleHeight - Screen.TwipsPerPixelY * 2
    End With
    If Me.Top + Me.Height > Screen.Height - 800 Then Me.Top = Me.Top - Me.Height - 200
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Me.Left - Me.Width
End Sub

Public Function ShowSelectChild(frmMain As Object, ByVal X As Single, ByVal Y As Single, ByVal cx As Single, ByVal cy As Single, Rs As ADODB.Recordset, aryColWidth As String, Optional ByVal blnAutoWidth As Boolean = False) As String
    On Error GoTo ErrHandle
    
    Dim i As Long, j As Long
    Dim lngWidth As Long
    
    Screen.MousePointer = vbHourglass
    ShowSelectChild = ""
    With frmSelectChild
        Set .vfgThis.DataSource = Rs
        Set .Font = .vfgThis.Font
        If blnAutoWidth Then
            For i = 0 To .vfgThis.Cols - 1
                If Val(Split(aryColWidth, ";")(i)) > 0 Then
                    For j = 0 To vfgThis.Rows - 1
                        If .vfgThis.ColWidth(i) < Me.TextWidth(.vfgThis.TextMatrix(j, i)) Then
                            .vfgThis.ColWidth(i) = Me.TextWidth(.vfgThis.TextMatrix(j, i))
                        End If
                    Next
                Else
                    .vfgThis.ColWidth(i) = 0
                End If
            Next
        Else
            For i = 0 To .vfgThis.Cols - 1
                .vfgThis.ColWidth(i) = Val(Split(aryColWidth, ";")(i))
            Next
        End If
        .vfgThis.Left = Screen.TwipsPerPixelX * 2
        .vfgThis.Top = Screen.TwipsPerPixelY * 2
        .vfgThis.Width = cx - Screen.TwipsPerPixelX * 2
        .vfgThis.Height = cy - Screen.TwipsPerPixelY * 2
        .Left = X
        .Top = Y
        .Width = cx
        .Height = cy
        Screen.MousePointer = vbDefault
        If IsNull(frmMain) Or frmMain Is Nothing Then
            .Show 1
        Else
            .Show 1, frmMain
        End If
        If .blnOK Then ShowSelectChild = .strReturn
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Screen.MousePointer = vbDefault
End Function

Private Sub vfgThis_DblClick()
    Dim i As Long
    
    If vfgThis.Row < 1 Then Unload Me: Exit Sub
    strReturn = ""
    For i = 0 To vfgThis.Cols - 1
        strReturn = strReturn & ";" & vfgThis.TextMatrix(vfgThis.Row, i)
    Next
    If Len(strReturn) > 0 Then strReturn = Mid(strReturn, 2)
    blnOK = True
    Unload Me
End Sub

Private Sub vfgThis_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    vfgThis_DblClick
End Sub
