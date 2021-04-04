VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.2#0"; "DicomObjects.ocx"
Begin VB.Form frmImgShow 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin DicomObjects.DicomViewer Viewer 
      Height          =   3825
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   5235
      _Version        =   262146
      _ExtentX        =   9234
      _ExtentY        =   6747
      _StockProps     =   35
   End
End
Attribute VB_Name = "frmImgShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    With Me.Viewer
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.Height
        If .Images.count > 0 Then
            .Width = .Images(1).SizeX * Screen.TwipsPerPixelX
            .Height = .Images(1).SizeY * Screen.TwipsPerPixelY
            Me.Width = .Width
            Me.Height = .Height
        End If
    End With
End Sub
Public Sub ShowMe(img As DicomImage, ObjFrm As Object, Left As Long, Top As Long)
    Me.Viewer.Images.Clear
    Me.Viewer.Images.Add img
    Me.Left = Left
    Me.Top = Top
    Me.Show , ObjFrm
End Sub

Public Sub HideMe()
    Unload Me
End Sub
