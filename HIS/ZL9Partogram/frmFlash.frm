VERSION 5.00
Begin VB.Form frmFlash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5070
   ControlBox      =   0   'False
   Icon            =   "frmFlash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1035
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1035
      Width           =   870
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   90
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1035
      Width           =   870
   End
   Begin VB.PictureBox picDo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   255
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   306
      TabIndex        =   1
      Top             =   465
      Width           =   4590
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   304
         Y1              =   12
         Y2              =   12
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   12
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   304
         X2              =   304
         Y1              =   0
         Y2              =   12
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   304
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblDo 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   135
         Left            =   30
         TabIndex        =   2
         Tag             =   "¨€"
         Top             =   30
         Width           =   4500
      End
   End
   Begin VB.Label lblPer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   4725
      TabIndex        =   3
      Top             =   255
      Width           =   90
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   210
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   90
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
