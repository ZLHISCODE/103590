VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmQCShowInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "对话框标题"
   ClientHeight    =   4650
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8415
   Icon            =   "frmQCShowInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin RichTextLib.RichTextBox rtxInfo 
      Height          =   3870
      Left            =   75
      TabIndex        =   1
      Top             =   105
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   6826
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmQCShowInfo.frx":000C
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   4140
      Width           =   1215
   End
End
Attribute VB_Name = "frmQCShowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mstrCaption  As String
Private mstrInfo As String

Public Function ShowMe(ByVal strCaption As String, ByVal strInfo As String, ByVal frmMain As Form)
    mstrCaption = strCaption
    mstrInfo = strInfo
    Me.Show vbModal, frmMain
End Function

Private Sub Form_Load()
    Me.Caption = mstrCaption
    Me.rtxInfo.Text = mstrInfo
    
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
