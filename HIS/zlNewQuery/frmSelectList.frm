VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelectList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   2715
   ClientTop       =   3420
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf 
      Height          =   2730
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   4815
      _Version        =   393216
      Rows            =   6
      Cols            =   10
      RowHeightMin    =   300
      BackColorFixed  =   -2147483648
      BackColorSel    =   -2147483647
      BackColorBkg    =   -2147483634
      GridColor       =   -2147483636
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmSelectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrColAlign As String
Private mstrColWidth As String
Private mrsData As New ADODB.Recordset
Private mblnFirst As Boolean
Public mlngPos As Long

Private Sub Form_Activate()
    Dim i As Long
    
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    
    mlngPos = 0
    If mrsData.State <> adStateOpen Then Unload Me: Exit Sub
    
    If Not mrsData.EOF Then Set msf.DataSource = mrsData
    
    If mstrColWidth <> "" Then
        For i = 0 To UBound(Split(mstrColWidth, ";"))
            msf.ColWidth(i + 1) = Val(Split(mstrColWidth, ";")(i))
        Next
    End If
    
    If mstrColAlign <> "" Then
        For i = 0 To UBound(Split(mstrColAlign, ";"))
            msf.ColAlignment(i + 1) = Val(Split(mstrColAlign, ";")(i))
        Next
    End If
    
    For i = 0 To msf.Cols - 1
        msf.ColAlignmentFixed(i) = 4
    Next
    
    msf.ColWidth(0) = 350
    For i = 1 To msf.Rows - 1
        msf.TextMatrix(i, 0) = i
    Next
End Sub

Private Sub MSF_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        If Abs(Timer - sngTim) > 0.5 Then
            strIdx = ""
        End If
        sngTim = Timer
        strIdx = strIdx & Chr(KeyAscii)
        KeyAscii = 0
        
        If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
        
        If msf.Rows - 1 >= CInt(strIdx) And CInt(strIdx) > 0 Then
            msf.Row = CInt(strIdx)
            msf.Col = 1: msf.ColSel = msf.Cols - 1
            If msf.Row - msf.Height / msf.RowHeight(0) \ 2 >= 1 Then
                msf.TopRow = msf.Row - msf.Height / msf.RowHeight(0) \ 2
            Else
                msf.TopRow = 1
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With msf
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub msf_DblClick()
    mlngPos = msf.Row
    Unload Me
End Sub

Private Sub msf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then msf_DblClick
    If KeyCode = 27 Then
        mlngPos = -1
        Unload Me
    End If
End Sub

Public Function ShowSelectList(frmMain As Object, X As Single, Y As Single, W As Single, H As Single, _
    rsData As ADODB.Recordset, Optional strColWidth As String = "", Optional strColAlign As String) As Long
   
    Set mrsData = rsData
    mstrColWidth = strColWidth
    mstrColAlign = strColAlign
    
    With frmSelectList
        .Left = X
        .Top = Y
        .Width = W
        .Height = H
        .Show vbModal, frmMain
        ShowSelectList = .mlngPos
    End With
End Function
