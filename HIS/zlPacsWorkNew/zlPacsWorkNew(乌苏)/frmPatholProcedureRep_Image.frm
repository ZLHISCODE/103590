VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPatholProcedureRep_Image 
   Caption         =   "图像选择"
   ClientHeight    =   6870
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   8895
   Icon            =   "frmPatholProcedureRep_Image.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8895
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picImages 
      Height          =   5895
      Left            =   240
      ScaleHeight     =   5835
      ScaleWidth      =   8475
      TabIndex        =   2
      Top             =   120
      Width           =   8535
      Begin DicomObjects.DicomViewer dvImages 
         Height          =   5805
         Left            =   0
         TabIndex        =   3
         Tag             =   "0"
         Top             =   0
         Visible         =   0   'False
         Width           =   8490
         _Version        =   262147
         _ExtentX        =   14975
         _ExtentY        =   10239
         _StockProps     =   35
         BackColor       =   4210752
         AutoDisplay     =   0   'False
      End
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   6000
      TabIndex        =   1
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&S)"
      Height          =   400
      Left            =   7320
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatholProcedureRep_Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public Sub ShowImageWindow(ByVal lngAdviceID As Long, ByVal blnMoved As Boolean, owner As Form)

    
    Call GetAllImages(Me, dvImages, blnMoved, 1, lngAdviceID, "", 100, 20)
    
    Call Me.Show(1, owner)
End Sub


Private Sub AdjustFace()
    picImages.Left = 120
    picImages.Top = 120
    picImages.Width = Me.Width - 360
    picImages.Height = Me.Height - cmdCancel.Height - 840
    
    dvImages.Left = 0
    dvImages.Top = 0
    dvImages.Width = picImages.Width
    dvImages.Height = picImages.Height
    
    cmdCancel.Left = Me.Width - cmdCancel.Width - 240
    cmdCancel.Top = picImages.Top + picImages.Height + 120
    
    cmdSure.Left = cmdCancel.Left - cmdSure.Width - 120
    cmdSure.Top = cmdCancel.Top
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub
