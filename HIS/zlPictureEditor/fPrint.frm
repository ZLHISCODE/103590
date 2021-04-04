VERSION 5.00
Begin VB.Form fPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   195
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3300
      TabIndex        =   13
      Top             =   2430
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "打印(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   2430
      Width           =   1050
   End
   Begin VB.TextBox txtCopies 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   3750
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   907
      Width           =   390
   End
   Begin VB.VScrollBar sbCopies 
      Height          =   300
      Left            =   4155
      Max             =   1
      Min             =   99
      TabIndex        =   6
      Top             =   907
      Value           =   1
      Width           =   210
   End
   Begin VB.ComboBox cbPrinters 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   375
      Width           =   4170
   End
   Begin VB.CheckBox chkCenter 
      Caption         =   "图片居中(&R)"
      Height          =   240
      Left            =   960
      TabIndex        =   10
      Top             =   1777
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox chkFitMode 
      Caption         =   "适合页面(&F)"
      Height          =   240
      Left            =   960
      TabIndex        =   11
      Top             =   2025
      Width           =   1530
   End
   Begin VB.ComboBox cbQuality 
      Height          =   315
      ItemData        =   "fPrint.frx":000C
      Left            =   960
      List            =   "fPrint.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   900
      Width           =   1215
   End
   Begin VB.ComboBox cbOrientation 
      Height          =   315
      ItemData        =   "fPrint.frx":003A
      Left            =   960
      List            =   "fPrint.frx":0044
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1320
      Width           =   1410
   End
   Begin VB.Label lblCopies 
      Caption         =   "份数(&F):"
      Height          =   210
      Left            =   3030
      TabIndex        =   4
      Top             =   945
      Width           =   675
   End
   Begin VB.Label lblSelect 
      Caption         =   "打印机(&P):"
      Height          =   210
      Left            =   195
      TabIndex        =   0
      Top             =   150
      Width           =   900
   End
   Begin VB.Label lblQuality 
      Caption         =   "质量(&Q):"
      Height          =   285
      Left            =   195
      TabIndex        =   2
      Top             =   915
      Width           =   705
   End
   Begin VB.Label lblOrientation 
      Caption         =   "方向(&O):"
      Height          =   285
      Left            =   195
      TabIndex        =   7
      Top             =   1335
      Width           =   705
   End
   Begin VB.Label lblAdjust 
      Caption         =   "调节(&A):"
      Height          =   285
      Left            =   195
      TabIndex        =   9
      Top             =   1755
      Width           =   705
   End
End
Attribute VB_Name = "fPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' gfPrint form
' Last revision: 2003.11.02
'================================================

Option Explicit

Public Enum pOrientationConstants
    [poPortrait] = 1
    [poLandScape]
End Enum

Private Sub Form_Load()
    With cbQuality
        .Clear
        .AddItem "草稿"
        .AddItem "低"
        .AddItem "中"
        .AddItem "高"
    End With
    With cbOrientation
        .Clear
        .AddItem "纵向"
        .AddItem "横向"
    End With
    
    Dim lIdx As Long
    
    '-- Get available printers
    For lIdx = 0 To Printers.Count - 1
        Call cbPrinters.AddItem(Printers(lIdx).DeviceName)
    Next lIdx

    '-- Show curr. printer
    cbPrinters = Printer.DeviceName

    '-- Set print quality (Medium)
    cbQuality.ListIndex = 2

    '-- Set pre-orientation
    With gfrmMain.Canvas.DIB
        If (.Width < .Height) Then
            cbOrientation.ListIndex = 0
          Else
            cbOrientation.ListIndex = 1
        End If
    End With
End Sub

Private Sub Form_Paint()
    Line (0, 185)-(ScaleWidth, 185), vb3DShadow
    Line (0, 186)-(ScaleWidth, 186), vb3DHighlight
End Sub

Private Sub cbPrinters_Click()

  Dim iPrn As Printer
    
    '-- Select printer
    For Each iPrn In Printers
        If (iPrn.DeviceName = cbPrinters) Then
            Set Printer = iPrn
        End If
    Next iPrn
End Sub

Private Sub sbCopies_Change()
    txtCopies = sbCopies
End Sub

Private Sub sbCopies_GotFocus()
    txtCopies.BackColor = vbWindowBackground
End Sub

Private Sub sbCopies_LostFocus()
    txtCopies.BackColor = vbButtonFace
End Sub

'//

Private Sub cmdOK_Click()

    On Error Resume Next
    
    '-- Set copies
    Printer.Copies = sbCopies
    If (Err = 0) Then
        '-- Set print quality
        Printer.PrintQuality = -(cbQuality.ListIndex + 1)
        If (Err = 0) Then
            '-- Print...
            If (BestFitPrint(gfrmMain.Canvas.DIB, cbOrientation.ListIndex + 1, CBool(chkCenter), CBool(chkFitMode))) Then
                Call MsgBox("无法打印", vbExclamation, "zlPictureEditor")
            End If
        End If
    End If
    
    On Error GoTo 0
    Call Unload(Me)
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Function BestFitPrint(DIB As cDIB, ByVal Orientation As pOrientationConstants, Optional ByVal Center As Boolean = 0, Optional ByVal FitToPage As Boolean = 0) As Boolean

  Dim e As Long

  Dim ofx As Long, ofy As Long
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long

    On Error Resume Next

    '-- Initialize printer
    Printer.Print vbNullString
    e = e Or Err

    '-- Set orientation
    Printer.Orientation = Orientation

    '-- Scale mode = [Pixels]
    Printer.ScaleMode = vbPixels

    '-- Fit info...
    Call DIB.GetBestFitInfo(Printer.ScaleWidth, Printer.ScaleHeight, bfx, bfy, bfW, bfH, True)
    '-- No fit info
    ofx = (Printer.ScaleWidth - DIB.Width) \ 2
    ofy = (Printer.ScaleHeight - DIB.Height) \ 2

    '-- Force fit ?
    If (DIB.Width > bfW Or DIB.Height > bfH) Then
        FitToPage = True
    End If
    '-- Center ?
    If (Center = 0) Then
        If (FitToPage) Then
            bfx = 0
            bfy = 0
          Else
            ofx = 0
            ofy = 0
        End If
    End If

    '-- Printing...
    If (FitToPage) Then
        Call DIB.Stretch(Printer.hdc, bfx, bfy, bfW, bfH, 0, 0, DIB.Width, DIB.Height)
      Else
        Call DIB.Stretch(Printer.hdc, ofx, ofy, DIB.Width, DIB.Height, 0, 0, DIB.Width, DIB.Height)
    End If
    Call Printer.EndDoc
    e = e Or Err

    '-- Success
    BestFitPrint = (e = 0)
    On Error GoTo 0
End Function
