VERSION 5.00
Begin VB.Form frmSelColor 
   BorderStyle     =   0  'None
   Caption         =   "颜色选择"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraColor 
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   2295
      Begin zl9CISBase.ColorPicker ColorFloodColor 
         Height          =   2190
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   3863
      End
   End
End
Attribute VB_Name = "frmSelColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public lngColor As Long
Public lblRow As Long
Public lblCol As Long

Private Sub ColorFloodColor_GotFocus()
    ColorFloodColor.Tag = "Focused"
End Sub

Private Sub ColorFloodColor_pOK()
    Dim lngSelFloodColor As Long
    lngColor = IIf(ColorFloodColor.Color = tomAutoColor, ColorFloodColor.AutoColor, ColorFloodColor.Color)
    
    frmMiningVessels.vfgList.Cell(flexcpFloodPercent, lblRow, lblCol) = 100
    If lngColor = 0 Then lngColor = -214748363
    frmMiningVessels.vfgList.Cell(flexcpFloodColor, lblRow, lblCol) = lngColor

    SendKeys "{ESCAPE}"
End Sub

Private Sub Form_Deactivate()
    frmMiningVessels.vfgList.Cell(flexcpFloodColor, lblRow, lblCol) = lngColor
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' esc clears, then quits
    If KeyAscii = 27 Then Tag = "": Unload Me
    ' enter quits
    If KeyAscii = 13 Then Unload Me
End Sub


