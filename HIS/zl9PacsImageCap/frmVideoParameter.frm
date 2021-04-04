VERSION 5.00
Begin VB.Form frmVideoParameter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVideoParameter.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   450
      Left            =   4320
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ��(&S)"
      Height          =   450
      Left            =   2880
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame fraCaptureParameterCfg 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "��ʾ����(&D)"
         Height          =   450
         Left            =   3600
         TabIndex        =   14
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Frame fraVideoShowWay 
         Caption         =   "��Ƶ��ʾ��ʽ"
         Height          =   2175
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2295
         Begin VB.OptionButton otpVideoShowWay 
            Caption         =   "����Ӧ��Ƶ��С"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   1800
            Width           =   1815
         End
         Begin VB.OptionButton otpVideoShowWay 
            Caption         =   "���ü���Χ��ʾ"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1815
         End
         Begin VB.OptionButton otpVideoShowWay 
            Caption         =   "����Ƶ������ʾ"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton otpVideoShowWay 
            Caption         =   "������������ʾ"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton otpVideoShowWay 
            Caption         =   "��ԭʼ��С��ʾ"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdCompressor 
         Caption         =   "ѹ������(&P)"
         Height          =   450
         Left            =   3600
         TabIndex        =   5
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdVideoSource 
         Caption         =   "��ƵԴ����(&R)"
         Height          =   450
         Left            =   3600
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdVideoFormat 
         Caption         =   "��ʽ����(&F)"
         Height          =   450
         Left            =   3600
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cboDevices 
         Height          =   330
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3930
      End
      Begin VB.Label labDevice 
         Caption         =   "�ɼ��豸"
         Height          =   255
         Left            =   255
         TabIndex        =   1
         Top             =   285
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmVideoParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'vfw����������
Public Enum VfwParameterCfgItem
  vpiShowWay = 1
  vpiVideoSource = 2
  vpiVideoFormat = 4
  vpiVideoCompressor = 8
  vpiVideoDisplay = 16
End Enum



Private WithEvents mVfwCaptureObj As clsVfwCapture  'vfw�ɼ�����
Attribute mVfwCaptureObj.VB_VarHelpID = -1

Private mlngDisplayWidth As Long    '��Ƶ��ʾ���
Private mlngDisplayHeight As Long   '��Ƶ��ʾ�߶�

Private mblnAllowChangeDevice As Boolean '�Ƿ����иı��豸


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'����: ��ʾvfw��������
'
'����˵��-----
'
'captureObj: vfw��Ƶ����
'iWidth: ��Ƶ��ʾ���ڿ��
'iHeight: ��Ƶ��ʾ���ڸ߶�
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowVfwParameter(ByRef captureObj As clsVfwCapture, _
                            ByVal iWidth As Integer, _
                            ByVal iHeight As Integer, _
                            objOwner As Object, _
                            Optional lngHideItem As Long = 0)

  Set mVfwCaptureObj = captureObj

  mlngDisplayWidth = iWidth
  mlngDisplayHeight = iHeight

  Call LoadCaptureDevice
  Call LoadVideoShowWay
  Call HideCfgItem(lngHideItem)

  Call Me.Show(1, objOwner)
End Sub


Private Sub HideCfgItem(ByVal lngCfgItem As Long)
  If (lngCfgItem And vpiShowWay) > 0 Then
    fraVideoShowWay.Enabled = False
    
    otpVideoShowWay(0).Enabled = False
    otpVideoShowWay(1).Enabled = False
    otpVideoShowWay(2).Enabled = False
    otpVideoShowWay(3).Enabled = False
    otpVideoShowWay(4).Enabled = False
  End If
  
  If (lngCfgItem And vpiVideoSource) > 0 Then
    cmdVideoSource.Enabled = False
  End If
  
  If (lngCfgItem And vpiVideoFormat) > 0 Then
    cmdVideoFormat.Enabled = False
  End If
  
  If (lngCfgItem And vpiVideoCompressor) > 0 Then
    cmdCompressor.Enabled = False
  End If
  
  If (lngCfgItem And vpiVideoDisplay) > 0 Then
    cmdDisplay.Enabled = False
  End If
  
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ȡ�òɼ��豸����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCaptureDeviceList() As String
    '��ȡ�����б�
  Const MAXVID_DRIVERS As Long = 9
  Const CAP_STRING_MAX As Long = 128

  Dim lngIndex As Long
  Dim strDevice As String
  Dim strVersion As String
  Dim strTmp As String

  strDevice = String$(CAP_STRING_MAX, 0)
  strVersion = String$(CAP_STRING_MAX, 0)

  For lngIndex = 0 To MAXVID_DRIVERS - 1
    If capGetDriverDescription(lngIndex, strDevice, CAP_STRING_MAX, strVersion, CAP_STRING_MAX) <> 0 Then

       strTmp = Left(strDevice, InStr(strDevice, vbNullChar) - 1) & "(" & Left$(strVersion, InStr(strVersion, vbNullChar) - 1) & ")"

       If Len(Trim(GetCaptureDeviceList)) > 0 Then
          GetCaptureDeviceList = GetCaptureDeviceList & ";"
       End If

       GetCaptureDeviceList = GetCaptureDeviceList & strTmp
    End If
  Next

End Function

'��ȡ��Ƶ����ʾ��ʽ
Private Sub LoadVideoShowWay()
  Dim parameter As clsVfwParameterCfg
  
  Set parameter = mVfwCaptureObj.GetCaptureParameter()
  
  otpVideoShowWay(parameter.VideoShowWay).value = True
End Sub

Private Sub LoadCaptureDevice()
  Dim i As Integer
  Dim parameter As clsVfwParameterCfg
  Dim strDevices() As String
  
  strDevices = Split(GetCaptureDeviceList(), ";")
  
  mblnAllowChangeDevice = False
  
  '��ȡ�豸����
  Me.cboDevices.Clear
  For i = 0 To UBound(strDevices)
    Call Me.cboDevices.AddItem(strDevices(i))
  Next
  
  If cboDevices.ListCount > 0 Then
    Set parameter = mVfwCaptureObj.GetCaptureParameter()
    
    If parameter.CaptureDeviceIndex < 0 Then
      cboDevices.ListIndex = 0
    Else
      If parameter.CaptureDeviceIndex < cboDevices.ListCount Then cboDevices.ListIndex = parameter.CaptureDeviceIndex
    End If
    
  End If

  mblnAllowChangeDevice = True
End Sub

Private Sub cboDevices_Click()
  On Error GoTo errProcess
    Dim parameter As clsVfwParameterCfg
  
    If Not mblnAllowChangeDevice Then
      Exit Sub
    End If
  
    '���µ�ǰ�ɼ��豸
    Set parameter = mVfwCaptureObj.GetCaptureParameter
  
    parameter.CaptureDeviceIndex = cboDevices.ListIndex
  
    Call mVfwCaptureObj.SetCaptureParameter(parameter)
  
    Call mVfwCaptureObj.RefreshParameter
  
    Call mVfwCaptureObj.UpdateCaptureWindowPos(mlngDisplayWidth, mlngDisplayHeight)
    
    Exit Sub
errProcess:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub

'vfwѹ������
Private Sub cmdCompressor_Click()
  On Error GoTo errProcess
    Call mVfwCaptureObj.ShowCaptureCompressionDialog
    
    Exit Sub
errProcess:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub

'vfw��ʾ����
Private Sub cmdDisplay_Click()
  On Error GoTo errProcess
    Call mVfwCaptureObj.ShowCaptureVideoDisplayDialog
    
    Exit Sub
errProcess:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub

'ȷ���¼�
Private Sub cmdSure_Click()
  Call mVfwCaptureObj.SaveVfwCaptureParameterToFile
  Call Unload(Me)
End Sub

'vfw��ʽ����
Private Sub cmdVideoFormat_Click()
  On Error GoTo errProcess
    Call mVfwCaptureObj.ShowCaptureVideoFormatDialog
    Call mVfwCaptureObj.UpdateCaptureWindowPos(mlngDisplayWidth, mlngDisplayHeight)
    
    Exit Sub
errProcess:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub

'vfw��ƵԴ����
Private Sub cmdVideoSource_Click()
  On Error GoTo errProcess
    Call mVfwCaptureObj.ShowCaptureVideoSourceDialog
    Exit Sub
errProcess:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
    err.Clear
End Sub

'����vfw��������
Private Sub cmdCancel_Click()
  Call mVfwCaptureObj.ReadVfwCaptureParameterFromFile
  Call mVfwCaptureObj.RefreshParameter
  Call mVfwCaptureObj.UpdateCaptureWindowPos(mlngDisplayWidth, mlngDisplayHeight)
  
  Call Unload(Me)
End Sub

Private Sub Form_Load()
  SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
End Sub

Private Sub mVfwCaptureObj_OnVideoWindowChange(ByVal lngWidth As Long, ByVal lngHeight As Long, blnIsChangeSize As Boolean)
  mlngDisplayHeight = lngHeight
  mlngDisplayWidth = lngWidth
End Sub


Private Sub otpVideoShowWay_Click(Index As Integer)
  Dim parameter As clsVfwParameterCfg
  
  Set parameter = mVfwCaptureObj.GetCaptureParameter
  
  parameter.VideoShowWay = Index
  
  Call mVfwCaptureObj.SetCaptureParameter(parameter)
  Call mVfwCaptureObj.UpdateCaptureWindowPos(mlngDisplayWidth, mlngDisplayHeight)
End Sub
