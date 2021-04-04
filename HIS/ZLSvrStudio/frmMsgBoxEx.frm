VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMsgBoxEx 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5310
   LinkTopic       =   "frmMsgBoxEx"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList igl48 
      Left            =   -240
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBoxEx.frx":0000
            Key             =   "INFO"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMsgBoxEx.frx":1B52
            Key             =   "QUESTION"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkIgnorePrompt 
      Caption         =   "不再提示(&S)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtPrompt 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmMsgBoxEx.frx":36A4
      Top             =   240
      Width           =   4155
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   120
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdFuncs 
      Height          =   390
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   1110
   End
End
Attribute VB_Name = "frmMsgBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Const MLNG_MIN_WIDTH                    As Long = 15 * 300
'Private Const MLNG_MAX_WIDTH                    As Long = 15 * 480
'Private Const MLNG_MIN_HEIGHT                   As Long = 15 * 150
'Private Const MLNG_MAX_HEIGHT                   As Long = 15 * 320
    
Private menmResult As enum_MsgboxButtonResult                       '返回

Public Enum enum_MsgboxButtonResult
    mbrOK = 0
    mbrNo = 1
    mbrYes = 2
End Enum
Public Enum enum_MsgboxButton
    mbnOKOnly = 0
    mbnYesNo = 1
End Enum
Public Enum enum_MsgboxStyle
    mslInformation = 1
    mslQuestion = 2
End Enum

Public Function ShowMe(ByVal FormOwner As Form, ByVal Prompt As String _
    , Optional ByVal Button As enum_MsgboxButton = mbnOKOnly _
    , Optional ByVal Style As enum_MsgboxStyle = mslInformation _
    , Optional ByVal ShowIgnore As Boolean = False _
    , Optional ByVal Title _
    , Optional ByVal HelpFile _
    , Optional ByVal Content) As enum_MsgboxButtonResult
'功能：入口方法
    
    txtPrompt.Text = Prompt
'    menmButton = enum_MsgboxButton
'    menmStyle = enum_MsgboxStyle
    chkIgnorePrompt.Visible = ShowIgnore
    Caption = Title
    picInfo.Picture = igl48.ListImages(mslInformation).Picture
    
    Call InitButton(Button)
    
    Me.Show vbModal
    ShowMe = menmResult
End Function

Private Sub Form_Load()
    '
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '...
End Sub

Private Sub cmdFuncs_Click(Index As Integer)
    Select Case Index
        Case 0
            menmResult = mbrOK
            Me.Hide
    End Select
End Sub

Private Sub InitButton(ByVal enmButton As enum_MsgboxButton)
    Select Case enmButton
        Case mbnOKOnly
            With cmdFuncs(0)
                .Caption = "确定"
                .Cancel = True
                .Default = True
                .Left = (Me.ScaleWidth - .Width) \ 2
            End With
        Case mbnYesNo
            '略...
    End Select
End Sub

