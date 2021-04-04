VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDSAConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "数字减影设置"
   ClientHeight    =   1410
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   330
      Left            =   4092
      TabIndex        =   6
      Top             =   828
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   330
      Left            =   4092
      TabIndex        =   5
      Top             =   312
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   828
      Left            =   360
      TabIndex        =   0
      Top             =   264
      Width           =   3396
      Begin VB.TextBox txtFrame 
         Height          =   300
         Left            =   1992
         TabIndex        =   2
         Top             =   276
         Width           =   552
      End
      Begin MSComCtl2.UpDown UpFrame 
         Height          =   300
         Left            =   2568
         TabIndex        =   3
         Top             =   276
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "帧"
         Height          =   180
         Left            =   2904
         TabIndex        =   4
         Top             =   336
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "基准图像为本序列的第"
         Height          =   180
         Left            =   156
         TabIndex        =   1
         Top             =   336
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmDSAConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintMaxFrame As Integer
Private mthisForm As frmViewer

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '对图像做出剪影
    Dim imgsTmp As New DicomImages
    Dim intViewerIndex As Integer
    Dim ww As Long
    Dim wl As Long
    
    '首先判断输入是否合法
    If Val(txtFrame.Text) < 1 Then
        MsgBox "输入图像帧数不能小于1", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtFrame.Text) > mintMaxFrame Then
        MsgBox "输入图像帧数不能大于图像的最大帧数:" & mintMaxFrame, vbInformation, gstrSysName
        Exit Sub
    End If
    
    '把剪影的图像作为临时图像添加到Viewer中
    '如果横向显示的Viewer个数少于2，则设置横向显示2列
    If mthisForm.intCountX < 2 Then
        mthisForm.intCountX = 2
        Call subChangeSeriesLayout(mthisForm)
    End If
    
    '把要剪影的图像作为临时图像，添加到Viewer中
    imgsTmp.Add mthisForm.viewer(mthisForm.intSelectedSerial).Images(mthisForm.SelectedImageIndex)
    intViewerIndex = funShowTempImages(mthisForm, imgsTmp, 0)
    
    '做出剪影
    mthisForm.viewer(intViewerIndex).Images(1).Mask = 1
    mthisForm.viewer(intViewerIndex).Images(1).MaskFrame = Val(txtFrame.Text)
    
    '调整窗宽窗位
    If funAutoWinWL(imgsTmp(1), 0, 0, imgsTmp(1).sizex, imgsTmp(1).sizey, ww, wl) Then
        mthisForm.viewer(intViewerIndex).Images(1).width = ww
        mthisForm.viewer(intViewerIndex).Images(1).Level = wl
    End If
    
    '显示作为蒙片的下一个图像
    If Val(txtFrame.Text) = mthisForm.viewer(intViewerIndex).Images(1).FrameCount Then
        mthisForm.viewer(intViewerIndex).Images(1).Frame = 1
    Else
        mthisForm.viewer(intViewerIndex).Images(1).Frame = Val(txtFrame.Text) + 1
    End If
    
    '提示剪影完成，并说明查看方法
    MsgBox "请使用电影播放功能查看减影效果!", vbInformation, gstrSysName
    
    Unload Me
End Sub

Private Sub txtFrame_Change()
    If Val(txtFrame) < 1 Then
        MsgBox "输入图像帧数不能小于1", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(txtFrame) > mintMaxFrame Then
        MsgBox "输入图像帧数不能大于图像的最大帧数:" & mintMaxFrame, vbInformation, gstrSysName
        Exit Sub
    End If
    
    UpFrame.Value = Val(txtFrame)

End Sub

Private Sub UpFrame_Change()
    txtFrame = UpFrame.Value
End Sub

Public Sub zlShowMe(intMaxFrame As Integer, intCurrentFrame As Integer, thisForm As frmViewer)
    mintMaxFrame = intMaxFrame
    UpFrame.Max = mintMaxFrame
    UpFrame.Min = 1
    UpFrame.Value = intCurrentFrame
    txtFrame = intCurrentFrame
    Set mthisForm = thisForm
    Me.Show 1, mthisForm
End Sub
