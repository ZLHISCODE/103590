VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "用户变动过程管理"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   17115
   Icon            =   "frmMain.frx":0000
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   10335
      Left            =   0
      ScaleHeight     =   10305
      ScaleWidth      =   2550
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2580
      Begin XtremeSuiteControls.ShortcutBar sbFunc 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
         _Version        =   589884
         _ExtentX        =   2143
         _ExtentY        =   6376
         _StockProps     =   64
      End
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   4200
      Top             =   1560
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMain.frx":6852
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InitSbfunc()
    
    With sbFunc
        .VisualTheme = xtpShortcutThemeOfficeXP
        .AddItem 1, "变动过程管理", frmItem.hwnd
        Call sbFunc.Icons.AddIcons(imgMain.Icons)
        .ExpandedLinesCount = .ItemCount
        .Selected = .FindItem(1)
    End With
    
End Sub

Private Sub MDIForm_Load()
    InitSbfunc
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    picFunc.Height = Me.ScaleHeight
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frmChild As Form
    
    For Each frmChild In Forms
        Unload frmChild
    Next
End Sub

Private Sub picFunc_Resize()
    sbFunc.Move 0, 0, picFunc.ScaleWidth, picFunc.ScaleHeight
End Sub
Private Sub sbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub
