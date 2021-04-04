VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "电影播放"
   ClientHeight    =   1275
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   6750
   FillStyle       =   6  'Cross
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   1275
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtEnd 
      Height          =   300
      Left            =   5370
      TabIndex        =   12
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txtStart 
      Height          =   300
      Left            =   5370
      TabIndex        =   10
      Top             =   390
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Height          =   420
      Left            =   1170
      Picture         =   "frmCine.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   708
      Width           =   900
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Left            =   1860
      Top             =   84
   End
   Begin VB.CommandButton cmdAlong 
      Height          =   420
      Left            =   2070
      Picture         =   "frmCine.frx":041E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   708
      Width           =   900
   End
   Begin VB.CommandButton cmdBackwards 
      Height          =   420
      Left            =   270
      Picture         =   "frmCine.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   708
      Width           =   900
   End
   Begin MSComctlLib.Slider sldCine 
      Height          =   324
      Left            =   252
      TabIndex        =   1
      Top             =   360
      Width           =   2736
      _ExtentX        =   4815
      _ExtentY        =   582
      _Version        =   393216
      Min             =   5
      Max             =   104
      SelStart        =   95
      TickStyle       =   3
      Value           =   95
   End
   Begin VB.Frame famPlayMode 
      Caption         =   "播放模式"
      Height          =   1050
      Left            =   3252
      TabIndex        =   0
      Top             =   120
      Width           =   1344
      Begin VB.OptionButton OptShuffle 
         Caption         =   "钟摆"
         Height          =   216
         Left            =   204
         TabIndex        =   3
         Top             =   672
         Width           =   696
      End
      Begin VB.OptionButton OptLoop 
         Caption         =   "循环"
         Height          =   180
         Left            =   204
         TabIndex        =   2
         Top             =   324
         Value           =   -1  'True
         Width           =   708
      End
   End
   Begin VB.Label lblCurrentNo 
      AutoSize        =   -1  'True
      Caption         =   "当前图像号："
      Height          =   180
      Left            =   4770
      TabIndex        =   13
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "终点："
      Height          =   180
      Left            =   4770
      TabIndex        =   11
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "起点："
      Height          =   180
      Left            =   4770
      TabIndex        =   9
      Top             =   450
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "加快"
      Height          =   180
      Left            =   2568
      TabIndex        =   7
      Top             =   96
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "减慢"
      Height          =   180
      Left            =   252
      TabIndex        =   6
      Top             =   96
      Width           =   360
   End
End
Attribute VB_Name = "frmCine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'使用跟鼠标穿梭相同的方法来播放电影图片
Public f As frmViewer
Dim blnisBack As Boolean            '图像播放的方向，True--向后播放；False--向前播放。

Private Sub cmdAlong_Click()
    blnisBack = False       '向前播放图像
    Timer.Enabled = True
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
End Sub

Private Sub cmdBackwards_Click()
    blnisBack = True        '向后播放图像
    Timer.Enabled = True
    Me.txtStart.Enabled = False
    Me.txtEnd.Enabled = False
End Sub

Private Sub cmdStop_Click()
    Timer.Enabled = False
    Me.txtStart.Enabled = True
    Me.txtEnd.Enabled = True
End Sub

Private Sub Form_Load()
    
    Dim thisViewer As DicomViewer
    If f.intSelectedSerial = 0 Then Exit Sub
    
    Set thisViewer = f.Viewer(f.intSelectedSerial)
    '计算播放图像在Viewer中的偏移量
    f.intStackOffset = f.SelectedImageIndex - thisViewer.CurrentIndex    '记录当前图像跟CurrentIndex之间的距离
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If f.SelectedImage.FrameCount > 1 Then   ''''多帧图像处理
        f.intStackCurrentlyImage = f.SelectedImage.Frame    '记录当前祯数或当前图像号
        f.blnStackisFrame = True            '记录采用多帧播放或是单祯循环播放
        Me.txtEnd.Text = thisViewer.Images(f.SelectedImageIndex).FrameCount
    Else                                    '单帧图像处理
        f.blnStackisFrame = False           '记录采用多帧播放或是单祯循环播放
        '记录穿梭前Viewer的CurrentIndex和当前图像
        Set f.SelectedLabel = Nothing
        f.intStackCurrentlyImage = thisViewer.CurrentIndex   '记录当前祯数或当前图像号
        Set f.objStackOldImage = thisViewer.Images(f.SelectedImageIndex)
        f.intStackIndex = thisViewer.Images(f.SelectedImageIndex).Tag
        Me.txtEnd.Text = ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count
    End If
        
    Me.txtStart.Text = 1
    Timer.Interval = (105 - sldCine) * 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '恢复电影播放前的窗口状态
    Timer.Enabled = False
    Dim j As Integer
    If f.intSelectedSerial = 0 Then Exit Sub
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With f.Viewer(f.intSelectedSerial)
    
        If f.blnStackisFrame Then    ''''多帧图像处理
            j = f.SelectedImage.Frame - f.intStackOffset
            f.SelectedImage.Frame = f.intStackCurrentlyImage
        Else
            '调用函数结束穿梭
            subStackEnd f.Viewer(f.intSelectedSerial), f
            j = f.intStackIndex - f.intStackOffset
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If j > ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count - .MultiColumns * .MultiRows + 1 Then
            j = ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count - .MultiColumns * .MultiRows + 1
        End If
        If j < 1 Then j = 1
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If f.VScro(f.intSelectedSerial).Visible Then f.VScro(f.intSelectedSerial).Value = j
    End With
End Sub

Private Sub sldCine_Click()
    Timer.Interval = (105 - sldCine) * 5
End Sub

Private Sub sldCine_GotFocus()
    Me.cmdStop.SetFocus
End Sub

Private Sub Timer_Timer()
    '定时器，处理图像的自动播放
    Dim objTempImage As DicomImage
    Dim thisViewer As DicomViewer
    
    Set thisViewer = f.Viewer(f.intSelectedSerial)
    If f.blnStackisFrame Then    ''''多帧图像处理
        If Not blnisBack Then   '向前播放图像
            If thisViewer.Images(f.SelectedImageIndex).Frame >= Val(Me.txtEnd.Text) Then
                If OptShuffle.Value Then
                   thisViewer.Images(f.SelectedImageIndex).Frame = Val(Me.txtEnd.Text) - 1
                   blnisBack = True
                Else
                    thisViewer.Images(f.SelectedImageIndex).Frame = Val(Me.txtStart.Text)
                End If
            Else
                thisViewer.Images(f.SelectedImageIndex).Frame = thisViewer.Images(f.SelectedImageIndex).Frame + 1
            End If
        Else                    '向后播放图像
            If thisViewer.Images(f.SelectedImageIndex).Frame <= Val(Me.txtStart.Text) Then
                If OptShuffle.Value Then
                   thisViewer.Images(f.SelectedImageIndex).Frame = Val(Me.txtStart.Text) + 1
                   blnisBack = False
                Else
                    thisViewer.Images(f.SelectedImageIndex).Frame = Val(Me.txtEnd.Text)
                End If
            Else
                thisViewer.Images(f.SelectedImageIndex).Frame = thisViewer.Images(f.SelectedImageIndex).Frame - 1
            End If
        End If
        Me.lblCurrentNo.Caption = "当前图像号：" & thisViewer.Images(f.SelectedImageIndex).Frame
    Else        '单帧图像处理
        '计算新的位置
        If Not blnisBack Then       '向前播放图像
            If f.intStackIndex >= Val(Me.txtEnd.Text) Then
                If OptShuffle.Value Then
                   f.intStackIndex = f.intStackIndex - 1
                   blnisBack = True
                Else
                    f.intStackIndex = Val(Me.txtStart.Text)
                End If
            Else
                f.intStackIndex = f.intStackIndex + 1
            End If
        Else                        '向后播放图像
            If f.intStackIndex <= Val(Me.txtStart.Text) Then
                If OptShuffle.Value Then
                   f.intStackIndex = Val(Me.txtStart.Text) + 1
                   blnisBack = False
                Else
                    f.intStackIndex = Val(Me.txtEnd.Text)
                End If
            Else
                 f.intStackIndex = f.intStackIndex - 1
            End If
        End If
        
        '把指定位置的图像添加到Viewer中
        Set objTempImage = funLoadAImage(f.intSelectedSerial, f.intStackIndex, 1)
        If Not objTempImage Is Nothing Then
            Call subInitAImage(objTempImage, f.intSelectedSerial, thisViewer)
            
            thisViewer.Images.Add objTempImage
            thisViewer.Images.Move thisViewer.Images.Count, f.SelectedImageIndex
            thisViewer.Images.Remove f.SelectedImageIndex + 1
            thisViewer.CurrentIndex = f.intStackCurrentlyImage
            
            Me.lblCurrentNo.Caption = "当前图像号：" & f.intStackIndex
        End If
    End If
End Sub

Private Sub txtEnd_LostFocus()
    Dim iImage As Integer
    Dim bError As Boolean
    
    iImage = Val(f.MSFViewer.TextMatrix(f.intSelectedSerial, 3))
    If Val(Me.txtEnd.Text) < Val(Me.txtStart.Text) Then bError = True
    If f.blnStackisFrame Then        '处理多帧图像
        If Val(Me.txtEnd.Text) > f.Viewer(f.intSelectedSerial).Images(iImage).FrameCount Then bError = True
    Else
        If Val(Me.txtEnd.Text) > ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count Then bError = True
    End If
    If bError Then
        MsgBox "终止值要大于开始值，且小于图像数量。", vbExclamation, gstrSysName
        If f.blnStackisFrame Then       '处理多帧图像
            Me.txtEnd.Text = f.Viewer(f.intSelectedSerial).Images(iImage).FrameCount
        Else
            Me.txtEnd.Text = ZLShowSeriesInfos(f.intSelectedSerial).ImageInfos.Count
        End If
    End If
End Sub

Private Sub txtStart_LostFocus()
    If Val(Me.txtStart.Text) > Val(Me.txtEnd.Text) Or Val(Me.txtStart.Text) < 1 Then
        MsgBox "开始值要小于最大值，且大于1。", vbInformation, gstrSysName
        Me.txtStart.Text = 1
    End If
End Sub
