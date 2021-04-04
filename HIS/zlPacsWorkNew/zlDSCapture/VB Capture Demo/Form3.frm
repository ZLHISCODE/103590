VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5670
   ScaleWidth      =   7725
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCutRate 
      Caption         =   "裁剪比率"
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdIsVfwDevice 
      Caption         =   "检测是否为VFW设备"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdVideoSize 
      Caption         =   "设置分辨率"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox cboVideoSize 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1680
      Width           =   6015
   End
   Begin VB.CommandButton cmdCfgDevice 
      Caption         =   "设置设备"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetName 
      Caption         =   "当前设备名称"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox cboColorDepth 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   6015
   End
   Begin VB.ComboBox cboEncoder 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   6015
   End
   Begin VB.ComboBox cboDevices 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label labVideoSize 
      Caption         =   "采集分辨率："
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label labColorDepth 
      Caption         =   "颜色深度："
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label labEncoder 
      Caption         =   "编码器名称："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label labDeviceName 
      Caption         =   "设备名称："
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objParameterCfg As DSCapParameterEnum
Private objCapture As DSCapture
Dim curParameter As TCaptureParameter



Public Sub ShowCaptureParameterConfig(captureobj As DSCapture)
'  If (objParameterCfg = Null) Then
'    MsgBox "Nul"
'  End If
'
  Set objParameterCfg = New DSCapParameterEnum
  Set objCapture = captureobj
  
  Call objCapture.GetCaptureParameter(curParameter)
  
  Call LoadDeviceData
  Call LoadEncoderData
  Call LoadColorDepthData
  Call LoadVideoSizeData
  
  
  Call Me.Show(1)
End Sub


Private Sub LoadDeviceData()
  Dim i As Integer
  Dim iDeviceCount As Long
  Dim sDeviceName As String
  Dim sErrMsg As String
  
  sErrMsg = objParameterCfg.GetDeviceCount(iDeviceCount)
  
  If Trim(sErrMsg <> "") Then
    MsgBox sErrMsg
    Exit Sub
  End If
  
  cboDevices.Clear
  For i = 0 To iDeviceCount - 1
    Call objParameterCfg.GetDeviceName(i, sDeviceName)
    Call cboDevices.AddItem(sDeviceName)
  Next i
  
  If cboDevices.ListCount > 0 Then
    cboDevices.ListIndex = 0
  End If
End Sub


Private Sub LoadEncoderData()
  Dim i As Integer
  Dim iDeviceCount As Long
  Dim sDeviceName As String
  Dim sErrMsg As String
  
  sErrMsg = objParameterCfg.GetEncoderCount(iDeviceCount)
  
  If Trim(sErrMsg <> "") Then
    MsgBox sErrMsg
    Exit Sub
  End If
  
  cboEncoder.Clear
  For i = 0 To iDeviceCount - 1
    Call objParameterCfg.GetEncoderName(i, sDeviceName)
    Call cboEncoder.AddItem(sDeviceName)
  Next i
  
  If cboEncoder.ListCount > 0 Then
    cboEncoder.ListIndex = 0
  End If
End Sub


Private Sub LoadColorDepthData()
  Dim i As Integer
  Dim iColorDepthCount As Long
  Dim sColorDepth As Long
  Dim sErrMsg As String
  
  sErrMsg = objParameterCfg.GetVideoColorDepthCount(iColorDepthCount)
  
  If Trim(sErrMsg <> "") Then
    MsgBox sErrMsg
    Exit Sub
  End If
  
  cboColorDepth.Clear
  For i = 0 To iColorDepthCount - 1
    Call objParameterCfg.GetVideoColorDepth(i, sColorDepth)
    Call cboColorDepth.AddItem(sColorDepth)
  Next i
  
  If cboColorDepth.ListCount > 0 Then
    cboColorDepth.ListIndex = 0
  End If
End Sub


Private Sub LoadVideoSizeData()
  Dim i As Integer
  Dim iVideoSizeCount As Long
  Dim sVideoSize As String
  Dim sErrMsg As String
  
  sErrMsg = objParameterCfg.GetVideoSizeCount(iVideoSizeCount)
  
  If Trim(sErrMsg <> "") Then
    MsgBox sErrMsg
    Exit Sub
  End If
  
  cboVideoSize.Clear
  For i = 0 To iVideoSizeCount - 1
    Call objParameterCfg.GetVideoSizeName(i, sVideoSize)
    Call cboVideoSize.AddItem(sVideoSize)
  Next i
  
  If cboVideoSize.ListCount > 0 Then
    cboVideoSize.ListIndex = 0
  End If
End Sub


Private Sub cboDevices_Click()
'  Dim sErrMsg As String
'
'  If cboDevices.Text <> "" Then
'    sErrMsg = objParameterCfg.SetCaptureDevice(cboDevices.Text)
'
'    If sErrMsg <> "" Then
'      MsgBox sErrMsg
'    End If
'  End If
End Sub

Private Sub cmdCfgDevice_Click()

  curParameter.CaptureDeviceName = cboDevices.Text
  Call objCapture.SetCaptureParameter(curParameter)
        
  objCapture.RePreview
   
End Sub

Private Sub cmdCutRate_Click()
  Call MsgBox("LeftRate:" & curParameter.leftRate & vbCrLf & "TopRate:" & curParameter.topRate)
End Sub

Private Sub cmdGetName_Click()
  MsgBox curParameter.CaptureDeviceName
End Sub

Private Sub cmdIsVfwDevice_Click()
  MsgBox objParameterCfg.CheckIsVfwDevice(cboDevices.Text)
End Sub

Private Sub cmdSave_Click()
  Dim sErrMsg As String
  
  objCapture.ParameterCfgFileName = "c:\CaptureParameter.ini"
  
  sErrMsg = objCapture.SaveParameterToFile()
  If sErrMsg <> "" Then
    MsgBox sErrMsg
  End If
End Sub

Private Sub cmdVideoSize_Click()

  curParameter.videoSize = cboVideoSize.Text
  Call objCapture.SetCaptureParameter(curParameter)
        
  objCapture.RePreview
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set objParameterCfg = Nothing
End Sub
