VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "采集图像浏览"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13245
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8880
   ScaleWidth      =   13245
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picCapture 
      Height          =   2895
      Left            =   3120
      ScaleHeight     =   2835
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

  Call Me.Refresh
  Call Me.PaintPicture(picCapture.Picture, 0, 0, Me.Width, Me.Height)
  
End Sub

Private Sub Form_GotFocus()

  Call Me.Refresh
  Call Me.PaintPicture(picCapture.Picture, 0, 0, Me.Width, Me.Height)
  
End Sub


'...
Private Sub Form_Paint()
  'Call Me.Refresh
  'Call Me.PaintPicture(picCapture.Picture, 0, 0, Me.Width, Me.Height)
End Sub

'窗体大小改变事件
Private Sub Form_Resize()

  Call Me.Refresh
  Call Me.PaintPicture(picCapture.Picture, 0, 0, Me.Width, Me.Height)
  
End Sub

Public Sub ShowCaptureImgFromMemory(pic As IPictureDisp)
  Set picCapture.Picture = pic
  If picCapture.Picture.Handle = 0 Then Exit Sub
    
  
  Me.Width = picCapture.ScaleX(picCapture.Picture.Width, vbHimetric, vbTwips) + 100
  Me.Height = picCapture.ScaleY(picCapture.Picture.Height, vbHimetric, vbTwips) + 400
          
     
  Call Me.PaintPicture(picCapture.Picture, 0, 0, Me.Width, Me.Height)
  
  Call SavePicture(pic, App.Path & "\jpgFormat.jpg")
    
  Call Me.Show(1)
End Sub

'显示采集的图像
Public Sub ShowCaptureImg(ByVal imgFile As String)
  If Dir(imgFile) = "" Then Exit Sub
  
  Set picCapture.Picture = LoadPicture(imgFile)
  If picCapture.Picture.Handle = 0 Then Exit Sub
    
  
  Me.Width = picCapture.ScaleX(picCapture.Picture.Width, vbHimetric, vbTwips) + 100
  Me.Height = picCapture.ScaleY(picCapture.Picture.Height, vbHimetric, vbTwips) + 400
          
     
  Call Me.PaintPicture(picCapture.Picture, 0, 0, Me.Width, Me.Height)
    
  Call Me.Show(1)
End Sub


Public Sub ShowCaptureImgFromClipBoard()
    

  
  Set picCapture.Picture = Clipboard.GetData(2)
  
  Me.Width = picCapture.ScaleX(picCapture.Picture.Width, vbHimetric, vbTwips) + 100
  Me.Height = picCapture.ScaleY(picCapture.Picture.Height, vbHimetric, vbTwips) + 400
          
     
  Call Me.PaintPicture(picCapture.Picture, 0, 0, Me.Width, Me.Height)
    
  Call Me.Show(1)
End Sub
