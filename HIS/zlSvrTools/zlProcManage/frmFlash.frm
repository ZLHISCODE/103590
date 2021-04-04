VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFlash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5070
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2445
      ScaleHeight     =   315
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picDo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   255
      ScaleHeight     =   195
      ScaleWidth      =   4590
      TabIndex        =   2
      Top             =   510
      Visible         =   0   'False
      Width           =   4590
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   4560
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   180
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   4560
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblDo 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   135
         Left            =   30
         TabIndex        =   3
         Tag             =   "¨€"
         Top             =   30
         Width           =   4500
      End
   End
   Begin MSComCtl2.Animation avi 
      Height          =   675
      Left            =   195
      TabIndex        =   0
      Top             =   75
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1191
      _Version        =   393216
      FullWidth       =   50
      FullHeight      =   45
   End
   Begin VB.Label lblPer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   4725
      TabIndex        =   4
      Top             =   300
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   210
      Left            =   1140
      TabIndex        =   1
      Top             =   285
      Width           =   90
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
