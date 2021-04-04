VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInputStatus 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1170
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmInputStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra 
      Height          =   1215
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   5655
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         Height          =   350
         Left            =   4350
         TabIndex        =   4
         Top             =   780
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   525
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已完成 10%"
         Height          =   180
         Left            =   120
         MousePointer    =   11  'Hourglass
         TabIndex        =   3
         Top             =   885
         Width           =   900
      End
      Begin VB.Label lblTitle 
         Caption         =   "正在输出到打印机："
         Height          =   210
         Left            =   120
         MousePointer    =   11  'Hourglass
         TabIndex        =   2
         Top             =   180
         Width           =   4770
      End
   End
End
Attribute VB_Name = "frmInputStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancel As Boolean

Public Property Let Describle(ByVal strTmp As String)
    lblTitle.Caption = strTmp
End Property

Public Property Let Value(ByVal sglTmp As Single)
    ProgressBar1.Value = sglTmp
    lblNote.Caption = "已完成 " & sglTmp & "%"
End Property

Public Property Get State() As Boolean
    State = mblnCancel
End Property

Private Sub cmdCancel_Click()
    mblnCancel = True
End Sub

Private Sub Form_Load()
    mblnCancel = False
End Sub
