VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.UserControl ctlPicture 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  '透明
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   6360
   Begin VB.PictureBox picBack 
      Height          =   2745
      Left            =   60
      ScaleHeight     =   2685
      ScaleWidth      =   3945
      TabIndex        =   0
      Top             =   105
      Width           =   4005
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   1665
         ScaleHeight     =   1935
         ScaleWidth      =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1680
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf 
         Height          =   480
         Left            =   105
         TabIndex        =   2
         Top             =   150
         Visible         =   0   'False
         Width           =   615
         _cx             =   1085
         _cy             =   847
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   0   'False
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ExactFit"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
      End
   End
End
Attribute VB_Name = "ctlPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarBorder As Integer                  '边框类型:0:平面;-1:按下
Private mvarAutoSize As Boolean

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PlayPaint()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pic_Paint()
    Call picBack_Paint
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picBack_Paint()
    Call RaisEffect(picBack, mvarBorder)
'    DoEvents
    RaiseEvent PlayPaint
End Sub

Private Sub swf_GotFocus()
    On Error Resume Next
    
    UserControl.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mvarBorder = -1
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mvarBorder = PropBag.ReadProperty("Border", -1)
    mvarAutoSize = PropBag.ReadProperty("AutoSize", False)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    Call ResizeControl(picBack, 0, 0, UserControl.Width, UserControl.Height)
    If mvarBorder = 0 Then
        Call ResizeControl(pic, 0, 0, UserControl.Width, UserControl.Height)
    Else
        Call ResizeControl(pic, 30, 30, UserControl.Width - 60, UserControl.Height - 60)
    End If
End Sub

Public Sub ShowPictureByFieldNew(ByVal lngNo As Long, ByVal W As Single, ByVal H As Single, ByVal bytSwf As Byte)
    Dim strFile As String
    
    picBack.Tag = FileName
    
    Select Case bytSwf
    Case 2
        pic.Visible = False
        swf.Visible = True
        strFile = ReadFlashByFieldNew(lngNo)
        Call PlayFlash(picBack, swf, strFile, W, H)
        On Error Resume Next
        Kill strFile
    Case Else
        pic.Visible = True
        swf.Visible = False
        Call DrawPicture(pic, ReadPicByFieldNew(lngNo), W, H)
    End Select
    
End Sub

Public Sub ShowPictureByFile(ByVal FileName As String, Optional ByVal FactSize As Boolean = False, Optional ByVal W As Single, Optional ByVal H As Single)
    If FileName <> "" And Dir(FileName) <> "" Then
        If Right(FileName, 3) = "pic" Then
            '显示图片格式的图形
            pic.Visible = True
            swf.Visible = False
            pic.Cls
            
            If FactSize Then
                UserControl.Width = W
                UserControl.Height = H
            End If
            
            On Error Resume Next '徐强增加
            
            Call DrawPicture(pic, VB.LoadPicture(FileName), picBack.Width, picBack.Height)
        Else
            '播放SWF格式的电影
            pic.Visible = False
            swf.Visible = True
            
            If FactSize Then
                UserControl.Width = W
                UserControl.Height = H
            End If
            Call PlayFlash(picBack, swf, FileName, picBack.Width, picBack.Height)
        End If
        
    End If
    
End Sub

Public Sub ShowByIPictureDisp(ByVal objStd As IPictureDisp, ByVal W As Single, ByVal H As Single)
    Call DrawPicture(pic, objStd, W, H)
End Sub

Public Property Let Border(ByVal vData As Integer)
    mvarBorder = vData
    Call RaisEffect(picBack, vData)
    PropertyChanged "Border"
End Property

Public Property Get Border() As Integer
    Border = mvarBorder
End Property

Public Property Let AutoSize(ByVal vData As Boolean)
    mvarAutoSize = vData
    PropertyChanged "AutoSize"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = mvarAutoSize
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Border", mvarBorder, -1)
    Call PropBag.WriteProperty("AutoSize", mvarAutoSize, False)
End Sub

Public Sub Cls()
    Set pic.Picture = Nothing
    pic.Cls
    swf.Visible = False
    swf.Movie = "-1"
    swf.Playing = False
    swf.StopPlay
End Sub
