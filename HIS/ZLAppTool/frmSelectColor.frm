VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSelectColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "颜色选择"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog cdg 
      Left            =   2100
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   0
      Top             =   270
      Width           =   180
   End
   Begin MSComctlLib.Toolbar tlbOther 
      Height          =   540
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   953
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   480
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tlbColor 
      Height          =   540
      Left            =   390
      TabIndex        =   2
      Top             =   390
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSelectColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mlngColor As Long

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadColor
End Sub

Private Sub LoadColor()
    Dim lngRGB As Long
    Dim i As Integer
    Dim lngColor(1 To 40, 1 To 3) As Long
    
    lngColor(1, 1) = 0: lngColor(1, 2) = 0: lngColor(1, 3) = 0
    lngColor(2, 1) = 150: lngColor(2, 2) = 50: lngColor(2, 3) = 0
    lngColor(3, 1) = 63: lngColor(3, 2) = 63: lngColor(3, 3) = 0
    lngColor(4, 1) = 0: lngColor(4, 2) = 63: lngColor(4, 3) = 0
    lngColor(5, 1) = 0: lngColor(5, 2) = 63: lngColor(5, 3) = 100
    lngColor(6, 1) = 0: lngColor(6, 2) = 0: lngColor(6, 3) = 128
    lngColor(7, 1) = 63: lngColor(7, 2) = 63: lngColor(7, 3) = 175
    lngColor(8, 1) = 63: lngColor(8, 2) = 63: lngColor(8, 3) = 63
    lngColor(9, 1) = 128: lngColor(9, 2) = 0: lngColor(9, 3) = 0
    lngColor(10, 1) = 255: lngColor(10, 2) = 100: lngColor(10, 3) = 0
    lngColor(11, 1) = 128: lngColor(11, 2) = 128: lngColor(11, 3) = 0
    lngColor(12, 1) = 0: lngColor(12, 2) = 128: lngColor(12, 3) = 0
    lngColor(13, 1) = 0: lngColor(13, 2) = 128: lngColor(13, 3) = 128
    lngColor(14, 1) = 0: lngColor(14, 2) = 0: lngColor(14, 3) = 255
    lngColor(15, 1) = 125: lngColor(15, 2) = 125: lngColor(15, 3) = 175
    lngColor(16, 1) = 125: lngColor(16, 2) = 125: lngColor(16, 3) = 125
    lngColor(17, 1) = 255: lngColor(17, 2) = 0: lngColor(17, 3) = 0
    lngColor(18, 1) = 255: lngColor(18, 2) = 150: lngColor(18, 3) = 0
    lngColor(19, 1) = 150: lngColor(19, 2) = 200: lngColor(19, 3) = 0
    lngColor(20, 1) = 50: lngColor(20, 2) = 155: lngColor(20, 3) = 100
    lngColor(21, 1) = 50: lngColor(21, 2) = 200: lngColor(21, 3) = 200
    lngColor(22, 1) = 50: lngColor(22, 2) = 100: lngColor(22, 3) = 255
    lngColor(23, 1) = 125: lngColor(23, 2) = 0: lngColor(23, 3) = 125
    lngColor(24, 1) = 156: lngColor(24, 2) = 156: lngColor(24, 3) = 156
    lngColor(25, 1) = 255: lngColor(25, 2) = 0: lngColor(25, 3) = 255
    lngColor(26, 1) = 255: lngColor(26, 2) = 200: lngColor(26, 3) = 0
    lngColor(27, 1) = 255: lngColor(27, 2) = 255: lngColor(27, 3) = 0
    lngColor(28, 1) = 0: lngColor(28, 2) = 255: lngColor(28, 3) = 0
    lngColor(29, 1) = 0: lngColor(29, 2) = 255: lngColor(29, 3) = 255
    lngColor(30, 1) = 0: lngColor(30, 2) = 200: lngColor(30, 3) = 255
    lngColor(31, 1) = 156: lngColor(31, 2) = 50: lngColor(31, 3) = 100
    lngColor(32, 1) = 198: lngColor(32, 2) = 198: lngColor(32, 3) = 198
    lngColor(33, 1) = 255: lngColor(33, 2) = 150: lngColor(33, 3) = 200
    lngColor(34, 1) = 255: lngColor(34, 2) = 200: lngColor(34, 3) = 150
    lngColor(35, 1) = 255: lngColor(35, 2) = 255: lngColor(35, 3) = 156
    lngColor(36, 1) = 200: lngColor(36, 2) = 255: lngColor(36, 3) = 200
    lngColor(37, 1) = 200: lngColor(37, 2) = 255: lngColor(37, 3) = 255
    lngColor(38, 1) = 150: lngColor(38, 2) = 200: lngColor(38, 3) = 255
    lngColor(39, 1) = 200: lngColor(39, 2) = 155: lngColor(39, 3) = 255
    lngColor(40, 1) = 255: lngColor(40, 2) = 255: lngColor(40, 3) = 255
    
    For i = 1 To 40
        lngRGB = RGB(lngColor(i, 1), lngColor(i, 2), lngColor(i, 3))
        picDraw.BackColor = lngRGB
        Rectangle picDraw.hdc, 0, 0, 12, 12
        ilsColor.ListImages.Add , "C" & lngRGB, picDraw.Image
    Next
    Set tlbColor.ImageList = ilsColor
    For i = 1 To 40
        tlbColor.Buttons.Add , ilsColor.ListImages(i).Key, , tbrCheck, ilsColor.ListImages(i).Key
    Next
    tlbColor.Buttons.Add , , , tbrSeparator
    tlbOther.Buttons.Add , "Other", "     获取其它颜色            ", tbrDefault
    
    '这些位置不是算出来的，是反复试出来的
    tlbColor.Left = 90
    tlbColor.Top = 90
    tlbOther.Top = tlbColor.Top + tlbColor.ButtonHeight * 6 - 120
    tlbOther.Left = tlbColor.Left
    
    Width = tlbColor.Width + 180
    Height = tlbOther.Top + tlbOther.Height + 240
    picDraw.Visible = False
End Sub

Public Function GetColor(lngColor As Long, frmParent As Form, ByVal sngLeft As Single, ByVal sngTop As Single) As Boolean
    
    frmSelectColor.Caption = "选择颜色"
    mblnOK = False
    mlngColor = lngColor
    
    frmSelectColor.Left = sngLeft - frmSelectColor.Width
    frmSelectColor.Top = sngTop + 600
    
    On Error Resume Next
    tlbColor.Buttons("C" & lngColor).Value = tbrPressed
    
    frmSelectColor.Show vbModal, frmParent
    
    GetColor = mblnOK
    If mblnOK = True Then lngColor = mlngColor
    
End Function

Private Sub tlbColor_ButtonClick(ByVal Button As MSComctlLib.Button)
    mblnOK = True
    mlngColor = Mid(Button.Key, 2)
    Unload Me
End Sub

Private Sub tlbOther_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next

    cdg.Color = mlngColor
    cdg.CancelError = True
    cdg.flags = cdlCCFullOpen Or cdlCCRGBInit
    cdg.ShowColor
    
    If Err <> 0 Then
        mblnOK = False
        Err.Clear
    Else
        mblnOK = True
        mlngColor = cdg.Color
    End If
    Unload Me
End Sub
