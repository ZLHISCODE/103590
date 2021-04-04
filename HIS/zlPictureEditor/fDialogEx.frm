VERSION 5.00
Begin VB.Form fDialogEx 
   BorderStyle     =   0  'None
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   163
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPreview 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   2445
      Begin VB.Frame fraJPEGOptions 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   0
         TabIndex        =   6
         Top             =   2655
         Width           =   2415
         Begin VB.TextBox txtJPEGQuality 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "80"
            Top             =   97
            Width           =   495
         End
         Begin VB.Label lblQuality 
            Caption         =   "JPEG 质量(0~100):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   1425
         End
      End
      Begin VB.CheckBox chkFitMode 
         Alignment       =   1  'Right Justify
         Caption         =   "适合大小"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1275
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2385
         Width           =   1035
      End
      Begin zlPictureEditor.ucCanvas Preview 
         Height          =   2310
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   4075
      End
      Begin VB.Label lblSize 
         Caption         =   "尺寸:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   2415
         Width           =   2235
      End
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "预览"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblWait 
      Alignment       =   1  'Right Justify
      Caption         =   "稍候..."
      Height          =   180
      Left            =   1455
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "fDialogEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' gfDialogEx form
' Last revision: 2003.11.02
'================================================

Option Explicit

Private Sub Form_Load()
    With gfrmMain
        '-- Get last status
        chkPreview = IIf(.DialogPreview, 1, 0)
        chkFitMode = IIf(.DialogFitMode, 1, 0)
        Preview.FitMode = .DialogFitMode
        txtJPEGQuality = .DialogJPEGquality
    End With
End Sub

Private Sub chkPreview_Click()
    If (chkPreview = 0) Then
        Call Preview.DIB.Destroy
        Call Preview.Resize
        lblSize = "大小:"
    End If
    gfrmMain.DialogPreview = CBool(chkPreview)
End Sub

Private Sub chkFitMode_Click()
    Preview.FitMode = CBool(chkFitMode)
    Call Preview.Resize
    gfrmMain.DialogFitMode = CBool(chkFitMode)
End Sub

Private Sub txtJPEGQuality_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub
Private Sub txtJPEGQuality_Change()
    With txtJPEGQuality
        If (.Text = vbNullString) Then
            .Text = "0"
            .SelStart = 0
            .SelLength = .MaxLength
        End If
        If (.Text > 100) Then
            .Text = 100
            .SelStart = .MaxLength
        End If
        gfrmMain.DialogJPEGquality = .Text
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Preview.DIB.Destroy
End Sub

