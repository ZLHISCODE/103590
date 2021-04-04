VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8310
   StartUpPosition =   3  '窗口缺省
   Begin XtremeSuiteControls.TabControl TabWindow 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _Version        =   589884
      _ExtentX        =   14631
      _ExtentY        =   11456
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowMe(ByVal lngHandle As Long)
    InitFaceScheme lngHandle
    
    Me.Show
End Sub

Private Sub InitFaceScheme(ByVal lngHandle As Long)
    Dim objItem As XtremeSuiteControls.TabControlItem
    
    Set objItem = TabWindow.InsertItem(1, "报告文本", lngHandle, 0)
    objItem.Selected = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    TabWindow.Left = 0
    TabWindow.Top = 0
    TabWindow.Width = Me.ScaleWidth
    TabWindow.Height = Me.ScaleHeight
End Sub
