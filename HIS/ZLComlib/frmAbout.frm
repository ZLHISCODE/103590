VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于..."
   ClientHeight    =   5310
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6300
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "s"
   Begin MSComctlLib.ListView lvw 
      Height          =   1515
      Left            =   225
      TabIndex        =   2
      Tag             =   "s"
      Top             =   2235
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   2672
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编号"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "标题"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "说明"
         Object.Width           =   5768
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4575
      TabIndex        =   0
      Top             =   4680
      Width           =   1485
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "系统信息(&S)..."
      Height          =   350
      Left            =   4575
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2610
      Top             =   2550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":0E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   5
      X1              =   -75
      X2              =   7281
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   -75
      X2              =   7506
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label lblCompile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "编译信息："
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   60
      TabIndex        =   14
      Top             =   1155
      Width           =   6180
   End
   Begin VB.Label lblGrant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   1500
      TabIndex        =   13
      Top             =   1545
      Width           =   90
   End
   Begin VB.Label lbl开发商 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   180
      Left            =   1320
      TabIndex        =   12
      Top             =   4080
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "产品开发商："
      Height          =   180
      Left            =   225
      TabIndex        =   11
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label lbl技术支持商 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   180
      Left            =   1320
      TabIndex        =   10
      Top             =   3825
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "技术支持商："
      Height          =   180
      Left            =   225
      TabIndex        =   9
      Top             =   3825
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   7271
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   0
      X2              =   7286
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "使用权已授予："
      Height          =   180
      Left            =   225
      TabIndex        =   8
      Top             =   1545
      Width           =   1260
   End
   Begin VB.Image imgLogo 
      Height          =   720
      Left            =   285
      Picture         =   "frmAbout.frx":0F24
      Stretch         =   -1  'True
      Top             =   165
      Width           =   720
   End
   Begin VB.Label lblFunc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "包含以下功能："
      Height          =   180
      Left            =   225
      TabIndex        =   7
      Top             =   2025
      Width           =   1260
   End
   Begin VB.Label lblRegVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZLHIS+ v10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2160
      TabIndex        =   6
      Top             =   585
      Width           =   1650
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "中联医院信息系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   1290
      TabIndex        =   5
      Top             =   165
      Width           =   2520
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      Caption         =   "版权所有(C) 中联信息产业公司"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3540
      TabIndex        =   4
      Top             =   2010
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -90
      X2              =   7196
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   -75
      X2              =   7196
      Y1              =   4395
      Y2              =   4395
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":1DEE
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   210
      TabIndex        =   3
      Top             =   4575
      Width           =   4305
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 注册键安全选项...
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySounda" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function messagebeep Lib "User32.dll" (ByVal wtype As Integer) As Integer
Private intMouse As Integer
Public btname As Object, yy As Boolean
Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' 试图从注册表得到系统信息程序路径\名称...
    If gobjComLib.OS.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWaRE\Microsoft\Shared Tools\MSINFO", "PaTH", SysInfoPath) Then
    ' 试图从注册表得到系统信息程序路径...
    ElseIf gobjComLib.OS.GetRegValue("HKEY_LOCAL_MACHINE\SOFTWaRE\Microsoft\Shared Tools Location", "MSINFO", SysInfoPath) Then
        ' 验证已知 32 位文件版本的存在
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        ' 错误 - 文件未找到...
        Else
            GoTo SysInfoErr
        End If
    ' 错误 - 注册项未找到...
    Else
        GoTo SysInfoErr
    End If
    Call Shell(SysInfoPath, vbNormalFocus)
    Exit Sub
SysInfoErr:
    MsgBox "此时系统信息无效", vbOKOnly, gstrSysName
End Sub

Private Sub Form_Activate()
    lvw.SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With Me.lblCopyRight
        If X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height Then
            intMouse = intMouse + 1
            If intMouse = 9 Then .Visible = True
        Else
            intMouse = 0
            .Visible = False
        End If
    End With
End Sub

Private Sub lvw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lvw.HitTest(X, Y) Is Nothing Then
        lvw.ToolTipText = lvw.HitTest(X, Y).SubItems(2)
    Else
        lvw.ToolTipText = ""
    End If
End Sub
