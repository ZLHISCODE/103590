VERSION 5.00
Begin VB.Form frmË¢¿¨ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ë¢¿¨"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4050
   Icon            =   "frmË¢¿¨.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin VB.CommandButton cmd¶Á¿¨ 
      Caption         =   "¶Á¿¨"
      Height          =   350
      Left            =   3300
      TabIndex        =   5
      Top             =   585
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdÆô¶¯ 
      Caption         =   "Æô¶¯"
      Height          =   350
      Left            =   3300
      TabIndex        =   7
      ToolTipText     =   "Æô¶¯ÃÜÂë¼üÅÌ"
      Top             =   1050
      Width           =   675
   End
   Begin VB.OptionButton opt¿¨Àà±ð 
      Caption         =   "´Å¿¨"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.OptionButton opt¿¨Àà±ð 
      Caption         =   "IC¿¨"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   1
      Left            =   1350
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   945
   End
   Begin VB.OptionButton opt¿¨Àà±ð 
      Caption         =   "Éí·ÝÖ¤ºÅ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Index           =   2
      Left            =   2460
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1365
   End
   Begin VB.TextBox txtÃÜÂë 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1050
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "È¡Ïû(&C)"
      Height          =   350
      Left            =   2790
      TabIndex        =   10
      Top             =   1620
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "È·¶¨(&O)"
      Height          =   350
      Left            =   1560
      TabIndex        =   9
      Top             =   1620
      Width           =   1100
   End
   Begin VB.TextBox txt¿¨ºÅ 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÃÜÂë"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   495
      TabIndex        =   6
      Top             =   1110
      Width           =   510
   End
   Begin VB.Label lbl¿¨ºÅ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "¿¨ºÅ"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   495
      TabIndex        =   3
      Top             =   660
      Width           =   510
   End
End
Attribute VB_Name = "frmË¢¿¨"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCard As String
Private mstrPass As String
Private mblnInit As Boolean

Public Function ShowME() As String
    mstrCard = ""
    mstrPass = ""
    mblnInit = False
    Me.Show 1
    ShowME = mstrCard & "|" & mstrPass
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txt¿¨ºÅ.Text) = "" Then
        MsgBox "ÇëË¢¿¨£¡", vbInformation, gstrSysName
        If txt¿¨ºÅ.Enabled = True Then txt¿¨ºÅ.SetFocus
        Exit Sub
    End If
    
    If gintType = 3 Then gstrIDNO = txt¿¨ºÅ.Text
    mstrCard = txt¿¨ºÅ.Text
    mstrPass = txtÃÜÂë.Text
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnInit Then Exit Sub
    mblnInit = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnInit = False
    Select Case gintType
        Case 1
            opt¿¨Àà±ð.Item(0).Value = True
        Case 2
            opt¿¨Àà±ð.Item(1).Value = True
        Case Else
            opt¿¨Àà±ð.Item(2).Value = True
            txt¿¨ºÅ.Text = gstrIDNO
    End Select
    gstrIDNO = ""
End Sub

Private Sub opt¿¨Àà±ð_Click(Index As Integer)
    gintType = Index + 1
    txt¿¨ºÅ.Enabled = (Index <> 1)
    cmd¶Á¿¨.Visible = (Index = 1)
    cmdÆô¶¯.Visible = (Index <> 1)
    Select Case Index
    Case 0
        lbl¿¨ºÅ.Caption = "¿¨ºÅ"
    Case 1
        lbl¿¨ºÅ.Caption = "IC¿¨ºÅ"
    Case 2
        lbl¿¨ºÅ.Caption = "Éí·ÝÖ¤ºÅ"
    End Select
    If Index <> 1 And mblnInit = True And txt¿¨ºÅ.Enabled = True Then txt¿¨ºÅ.SetFocus
End Sub

Private Sub txt¿¨ºÅ_GotFocus()
    If gblnLED Then
        zl9LedVoice.Speak "#5"
    End If
End Sub

Private Sub txtÃÜÂë_GotFocus()
    If gblnLED Then
        zl9LedVoice.Speak "#0"
    End If
End Sub

Private Sub cmd¶Á¿¨_Click()
    Dim intPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String
    Dim lngHandle As Long
    On Error GoTo errHand
    
    strPort = GetSetting("ZLSOFT", "¹«¹²Ä£¿é\¹óÑôÊÐÒ½±£", "¶Ë¿Ú", "COM1")
    If strPort = "USB" Then
        intPort = 100
    Else
        intPort = Right(strPort, 1)
    End If
    
    '´ò¿ª¶Á¿¨Æ÷
    STRERR = Space(2000)
    If SGZ_IFD_Open(intPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '¶ÁÈ¡PSAMÐ¾Æ¬ºÅÂë
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)
    
    '¶ÁÈ¡Éç±£¿¨ºÅ
    STRERR = Space(2000)
    gstrSNO = Space(2000)
    strPin = "000000"
    strAddr = "MF|EF05|07|$MF|EF06|01|$"
    If SGZ_ICC_ReadCardInfo(lngHandle, intCardType, strPin, strAddr, gstrSNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrSNO = TruncZero(gstrSNO)
    gstrIDNO = Split(gstrSNO, "|")(5)
    gstrSNO = Split(gstrSNO, "|")(2)
    txt¿¨ºÅ.Text = gstrSNO
    
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, intPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    txtÃÜÂë.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    Call cmdOK_Click
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Sub

Private Sub cmdÆô¶¯_Click()
        Dim intPort As Integer, intCardType As Integer
    Dim strPort As String
    Dim STRERR As String, strPin As String, strPass As String, strAddr As String
    Dim lngHandle As Long
    On Error GoTo errHand
    
    strPort = GetSetting("ZLSOFT", "¹«¹²Ä£¿é\¹óÑôÊÐÒ½±£", "¶Ë¿Ú", "COM1")
    If strPort = "USB" Then
        intPort = 100
    Else
        intPort = Right(strPort, 1)
    End If
    
    '´ò¿ª¶Á¿¨Æ÷
    STRERR = Space(2000)
    If SGZ_IFD_Open(intPort, lngHandle, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '¶ÁÈ¡PSAMÐ¾Æ¬ºÅÂë
    STRERR = Space(2000)
    gstrPSAMNO = Space(2000)
    If SGZ_SAM_ReadNmuber(lngHandle, gstrPSAMNO, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    gstrPSAMNO = TruncZero(gstrPSAMNO)
    
    STRERR = Space(2000)
    strPass = Space(2000)
    If SGZ_IFD_GetPIN(lngHandle, intPort, strPass, STRERR) <> 0 Then
        MsgBox STRERR, vbInformation, gstrSysName
        GoTo Exith
    End If
    strPass = TruncZero(strPass)
    txtÃÜÂë.Text = strPass
    
Exith:
    STRERR = Space(2000)
    If lngHandle > 0 Then Call SGZ_IFD_Close(lngHandle, STRERR)
    Call cmdOK_Click
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo Exith
End Sub
