VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   0  'None
   Caption         =   "纸张"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "设置自定义纸张使用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   2565
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

