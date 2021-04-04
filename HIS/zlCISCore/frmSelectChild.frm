VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelectChild 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4230
   ClientLeft      =   2760
   ClientTop       =   3465
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf 
      Height          =   2940
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5186
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorSel    =   8388608
      BackColorBkg    =   -2147483624
      GridColor       =   -2147483632
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "È¡Ïû(&C)"
      Height          =   350
      Left            =   3600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3615
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
Private v_SaveColor As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If blnFirst = False Then Exit Sub
    blnFirst = False
    msf_EnterCell
    msf.SetFocus
End Sub

Private Sub Form_Load()
    blnOK = False
    blnFirst = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With msf
        .Left = Screen.TwipsPerPixelX
        .Top = Screen.TwipsPerPixelY
        .Width = Me.ScaleWidth - Screen.TwipsPerPixelX * 2
        .Height = Me.ScaleHeight - Screen.TwipsPerPixelY * 2
    End With
End Sub

Public Function ShowSelectChild(frmMain As Object, ByVal X As Single, ByVal Y As Single, ByVal CX As Single, ByVal CY As Single, rs As ADODB.Recordset, aryColWidth As String, Optional ByVal blnAutoWidth As Boolean = False) As String
On Error GoTo ErrHandle
    Dim i As Long, j As Long
    Dim lngWidth As Long
    
    Screen.MousePointer = vbHourglass
    ShowSelectChild = ""
    With frmSelectChild
        Set .msf.DataSource = rs
        Set .Font = .msf.Font
        If blnAutoWidth Then
            For i = 0 To .msf.Cols - 1
                If Val(Split(aryColWidth, ";")(i)) > 0 Then
                    For j = 0 To msf.Rows - 1
                        If .msf.ColWidth(i) < Me.TextWidth(.msf.TextMatrix(j, i)) Then
                            .msf.ColWidth(i) = Me.TextWidth(.msf.TextMatrix(j, i))
                        End If
                    Next
                Else
                    .msf.ColWidth(i) = 0
                End If
            Next
        Else
            For i = 0 To .msf.Cols - 1
                .msf.ColWidth(i) = Val(Split(aryColWidth, ";")(i))
            Next
        End If
        .msf.Left = Screen.TwipsPerPixelX * 2
        .msf.Top = Screen.TwipsPerPixelY * 2
        .msf.Width = CX - Screen.TwipsPerPixelX * 2
        .msf.Height = CY - Screen.TwipsPerPixelY * 2
        .Left = X
        .Top = Y
        .Width = CX
        .Height = CY
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

Private Sub msf_EnterCell()
    v_SaveColor = msf.CellForeColor
    SelectRow msf
End Sub

Private Sub msf_LeaveCell()
    UnSelectRow msf, v_SaveColor
End Sub

Private Sub msf_DblClick()
    Dim i As Long
    
    If msf.Row < 1 Then Exit Sub
    strReturn = ""
    For i = 0 To msf.Cols - 1
        strReturn = strReturn & ";" & msf.TextMatrix(msf.Row, i)
    Next
    If Len(strReturn) > 0 Then strReturn = Mid(strReturn, 2)
    blnOK = True
    Unload Me
End Sub

Private Sub msf_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    msf_DblClick
End Sub

