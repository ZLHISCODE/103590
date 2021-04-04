VERSION 5.00
Object = "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}#1.0#0"; "shell32.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Begin VB.Form frmVideoSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "采集参数设置"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   Icon            =   "frmVideoSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin ScanLibCtl.ImgScan imageScannerConfig 
      Left            =   5400
      Top             =   4080
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.Frame fmeSuggestionMode 
      Caption         =   "采集提示方式"
      Height          =   615
      Left            =   240
      TabIndex        =   22
      Top             =   1080
      Width           =   5415
      Begin VB.CheckBox chkCaptureSound 
         Caption         =   "采集声音提示"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkCaptureWindow 
         Caption         =   "采集弹窗提示"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenDir 
      Left            =   2040
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "扫描参数设置"
      Height          =   1455
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Width           =   5415
      Begin VB.CommandButton cmdImageCompressConfig 
         Caption         =   "压缩设置(&P)"
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelectScanDevice 
         Caption         =   "设备选择(&D)"
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox tbxTempDir 
         Height          =   390
         Left            =   1800
         TabIndex        =   18
         Text            =   "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA"
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton cmdDirSelect 
         Caption         =   "…"
         Height          =   375
         Left            =   4920
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.Label labTempDir 
         Caption         =   "扫描设备临时目录："
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdParameterCfg 
      Caption         =   "视频设置(&V)"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame fraVideoDriverType 
      Caption         =   "视频驱动类型设置"
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   5415
      Begin VB.OptionButton optDriver 
         Caption         =   "TWAIN 驱动"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optDriver 
         Caption         =   "VFW 驱动"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optDriver 
         Caption         =   "WDM 驱动"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   5640
      Width           =   1100
   End
   Begin VB.Frame FraGather 
      Caption         =   "脚踏采集方式设置"
      Height          =   2055
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   5415
      Begin VB.ComboBox cboZoom 
         Height          =   300
         ItemData        =   "frmVideoSetup.frx":000C
         Left            =   3360
         List            =   "frmVideoSetup.frx":001C
         TabIndex        =   10
         Text            =   "1"
         Top             =   1650
         Width           =   1575
      End
      Begin VB.CheckBox chkShowImage 
         Caption         =   "鼠标移动时显示大图,放大倍数为："
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   3255
      End
      Begin VB.ComboBox cboCommCapType 
         Height          =   300
         ItemData        =   "frmVideoSetup.frx":002E
         Left            =   1440
         List            =   "frmVideoSetup.frx":003B
         TabIndex        =   1
         Top             =   740
         Width           =   3495
      End
      Begin VB.ComboBox cboPort 
         Height          =   300
         ItemData        =   "frmVideoSetup.frx":005D
         Left            =   1440
         List            =   "frmVideoSetup.frx":007C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   3495
      End
      Begin VB.TextBox txtComInterval 
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Text            =   "1"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "脚踏采集方式"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "脚踏端口(&P)"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   300
         Width           =   990
      End
      Begin VB.Label Label10 
         Caption         =   "脚踏时间间隔                                         秒"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1230
         Width           =   5055
      End
   End
   Begin Shell32Ctl.ShellFolderViewOC ShellFolderViewOC 
      Left            =   1560
      OleObjectBlob   =   "frmVideoSetup.frx":00B4
      Top             =   5640
   End
End
Attribute VB_Name = "frmVideoSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strRegName As String     '记录注册表的中间路径
Public frmParent As frmWork_Video  'frmVideoCapture     '父窗口
Public mstrPrivs As String
Public mlngModul As Long
Public mlngCurDepartId As Long

Private DX7 As New DirectX7
Private DxInput As DirectInput
Private DiDevEnum As DirectInputEnumDevices


Private mVideoCapture As clsVideoCapture





'modify by tjh at 2010-01-21
Public Sub ShowParameterConfig(ByRef videoCapture As clsVideoCapture, ByRef owner As Object)
  Set mVideoCapture = videoCapture
  
  Call LoadDriverType
  
  Call Me.Show(0, owner)
End Sub


'modify by tjh at 2010-01-21
'读取当前使用的驱动类型
Private Sub LoadDriverType()
  If mVideoCapture Is Nothing Then Exit Sub
  
  Select Case mVideoCapture.VideoDriverType
    Case vdtTWAIN
      optDriver(2).value = True
    Case vdtVFW
      optDriver(1).value = True
    Case vdtWDM
      optDriver(0).value = True
  End Select
  

End Sub

Private Sub cboCommCapType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


'Private Sub cboDrivers_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
'End Sub

Private Sub cboPort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


''''''''''''''''''''''''''''''''''
'选择扫描设备的临时图像存储目录
''''''''''''''''''''''''''''''''''
Private Sub cmdDirSelect_Click()
  Dim shl As Object
  Set shl = CreateObject("Shell.application")
  
  On Error GoTo final
  
    Dim fd As Object
    Set fd = shl.BrowseForFolder(0, "扫描设备临时目录选择", 0, "\")
  
    If Not fd Is Nothing Then
      tbxTempDir.Text = fd.Self.Path
    End If
final:
  Set shl = Nothing
  Set fd = Nothing
  
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
'显示压缩设置
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdImageCompressConfig_Click()
  On Error GoTo errHandle
    Call imageScannerConfig.ShowScanPreferences
  Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
  On Error GoTo errHandle
    Call SavePara
    
    Call frmParent.zlInitModule(mlngModul, mstrPrivs, mlngCurDepartId)
    
    Unload Me
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdParameterCfg_Click()
  On Error GoTo errHandle
    Call mVideoCapture.ShowCaptureParameterCfgDialog
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
'扫描设备选择
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSelectScanDevice_Click()
  On Error GoTo errHandle
    Call imageScannerConfig.ShowSelectScanner
  Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
'  SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
  
  Call InitPara
End Sub

Private Sub InitPara()
    Dim strExeRoom As String
    Dim strDeviceNO As String, iPortNumber As Integer
    Dim i  As Integer
    Dim iCapType As Integer
    Dim strTmp() As String
    On Error GoTo ErrorHand
    
    With cboPort
        .Clear
        .AddItem "无"
        .AddItem "COM1"
        .AddItem "COM2"
        .AddItem "COM3"
        .AddItem "COM4"
        .AddItem "COM5"
        .AddItem "COM6"
        .AddItem "COM7"
        .AddItem "COM8"
    End With
    
    Set DxInput = DX7.DirectInputCreate()
    Set DiDevEnum = DxInput.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
    For i = 1 To DiDevEnum.GetCount
        cboPort.AddItem DiDevEnum.GetItem(i).GetInstanceName
    Next
    
    If IsNumeric(zlDatabase.GetPara("脚踏端口", glngSys, mlngModul, "1")) Then
        iPortNumber = Val(zlDatabase.GetPara("脚踏端口", glngSys, mlngModul, "1", Array(cboPort), InStr(mstrPrivs, "采集参数设置") > 0))
        cboPort.ListIndex = iPortNumber
    Else
        SeekIndex cboPort, zlDatabase.GetPara("脚踏端口", glngSys, mlngModul, "", Array(cboPort), InStr(mstrPrivs, "采集参数设置") > 0)
    End If
        
    
    iCapType = Val(zlDatabase.GetPara("脚踏采集方式", glngSys, mlngModul, "1", Array(cboCommCapType), InStr(mstrPrivs, "采集参数设置") > 0))
    
    If iCapType = 0 Then
        cboCommCapType.ListIndex = 0
    ElseIf iCapType = 1 Then
        cboCommCapType.ListIndex = 1
    Else
        cboCommCapType.ListIndex = 2
    End If
    
    Dim strRegPath As String
    
    strRegPath = "公共模块\" & App.ProductName & "\frmVideoCapture"
    
    tbxTempDir.Text = GetSetting("ZLSOFT", strRegPath, "扫描设备临时目录", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
    'tbxTempDir.Text = zlDatabase.GetPara("扫描设备临时目录", glngSys, mlngModul, "0", Array(tbxTempDir), InStr(mstrPrivs, "采集参数设置") > 0)
    
    txtComInterval.Text = zlDatabase.GetPara("脚踏时间间隔", glngSys, mlngModul, "1", Array(txtComInterval), InStr(mstrPrivs, "采集参数设置") > 0)
    chkShowImage.value = zlDatabase.GetPara("鼠标移动时显示大图", glngSys, mlngModul, "0", Array(chkShowImage), InStr(mstrPrivs, "采集参数设置") > 0)
    cboZoom.Text = zlDatabase.GetPara("采集大图放大倍数", glngSys, mlngModul, "1", Array(cboZoom), InStr(mstrPrivs, "采集参数设置") > 0)
    
    chkCaptureWindow.value = zlDatabase.GetPara("采集后弹窗提示", glngSys, mlngModul, "0", Array(chkCaptureWindow), InStr(mstrPrivs, "采集参数设置") > 0)
    chkCaptureSound.value = zlDatabase.GetPara("采集后声音提示", glngSys, mlngModul, "0", Array(chkCaptureSound), InStr(mstrPrivs, "采集参数设置") > 0)
    
    If Val(cboZoom.Text) = 0 Then cboZoom.Text = 1
    
    cmdOK.Enabled = InStr(mstrPrivs, "采集参数设置") > 0
    cmdSelectScanDevice.Enabled = InStr(mstrPrivs, "采集参数设置") > 0
    cmdImageCompressConfig.Enabled = InStr(mstrPrivs, "采集参数设置") > 0
    
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SavePara()
On Error GoTo errHand
    
    Dim strRegPath As String
    
    strRegPath = "公共模块\" & App.ProductName & "\frmVideoCapture"
    
    '9以下是COM口,0表示不使用外部设备
    If cboPort.ListIndex = 0 Then
        Call zlDatabase.SetPara("脚踏端口", "无", glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    ElseIf cboPort.ListIndex < 9 Then
        Call zlDatabase.SetPara("脚踏端口", cboPort.ListIndex, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    Else
        Call zlDatabase.SetPara("脚踏端口", cboPort.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    End If

'    modify by tjh at 2010-01-21
'    If Me.cboDrivers.ListCount > 0 Then
'        Call zlDatabase.SetPara("Drivers", cboDrivers.ListIndex, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
'    End If

    '保存视频驱动类型，目前只有两种驱动类型
    If optDriver(0).value Then Call zlDatabase.SetPara("视频驱动类型", 0, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    If optDriver(1).value Then Call zlDatabase.SetPara("视频驱动类型", 1, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    If optDriver(2).value Then Call zlDatabase.SetPara("视频驱动类型", 2, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    
    Call zlDatabase.SetPara("采集后弹窗提示", chkCaptureWindow.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    Call zlDatabase.SetPara("采集后声音提示", chkCaptureSound.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    
    Call zlDatabase.SetPara("脚踏采集方式", cboCommCapType.ListIndex, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    Call zlDatabase.SetPara("脚踏时间间隔", IIf(Val(txtComInterval.Text) = 0, 1, Val(txtComInterval.Text)), glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    Call zlDatabase.SetPara("鼠标移动时显示大图", chkShowImage.value, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    Call zlDatabase.SetPara("采集大图放大倍数", IIf(Val(cboZoom.Text) = 0, 1, Val(cboZoom.Text)), glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    'Call zlDatabase.SetPara("扫描设备临时目录", tbxTempDir.Text, glngSys, mlngModul, InStr(mstrPrivs, ";参数设置;") > 0)
    Call SaveSetting("ZLSOFT", strRegPath, "扫描设备临时目录", tbxTempDir.Text)

    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub optDriver_Click(Index As Integer)
  On Error GoTo errHandle
    Select Case Index
        Case 0
            If mVideoCapture.VideoDriverType = vdtWDM Then Exit Sub
      
            'mVideoCapture.VideoDriverType = vdtWDM
            Call frmParent.UpdateCaptureDirver(vdtWDM)
        Case 1
            If mVideoCapture.VideoDriverType = vdtVFW Then Exit Sub
          
            'mVideoCapture.VideoDriverType = vdtVFW
            Call frmParent.UpdateCaptureDirver(vdtVFW)
        Case 2
            If mVideoCapture.VideoDriverType = vdtTWAIN Then Exit Sub
      
            'mVideoCapture.VideoDriverType = vdtTWAIN
            Call frmParent.UpdateCaptureDirver(vdtTWAIN)
    End Select
  
    Call mVideoCapture.StopPreview
  
    '如果为TWAIN的方式，则不进行视频的刷新
    If mVideoCapture.VideoDriverType <> vdtTWAIN Then
        Call mVideoCapture.StartPreview
  
        Call mVideoCapture.RefreshVideoWindow
    Else
        '不做任何操作...
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub txtComInterval_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function
