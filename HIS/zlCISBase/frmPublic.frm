VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{30676308-C43A-4FA0-8830-D0C65683F1E7}#1.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmPublic 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.ComboBox cmbFont 
      Height          =   300
      Left            =   135
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   765
      Width           =   2040
   End
   Begin zlRichEditor.Editor edtPublic 
      Height          =   690
      Left            =   540
      TabIndex        =   1
      Top             =   1260
      Visible         =   0   'False
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1217
   End
   Begin VB.Image imgErrPic 
      Height          =   240
      Left            =   135
      Picture         =   "frmPublic.frx":0000
      Top             =   1665
      Width           =   210
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPublic.frx":0302
   End
End
Attribute VB_Name = "frmPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Elements As New cEPRElements     '临时要素集合，用于文档间的复制粘贴

Private Sub Form_Load()
    Call GetAllFonts
End Sub

Private Sub GetAllFonts()
    '字体列表
    Dim sFont As String
    Dim i As Long
    
    If Printers.Count = 0 Then
        For i = 0 To Screen.FontCount - 1
           sFont = Screen.Fonts(i)
           cmbFont.AddItem sFont
        Next i
    Else
        For i = 0 To Printer.FontCount - 1
           sFont = Printer.Fonts(i)
           cmbFont.AddItem sFont
        Next i
    End If
End Sub
