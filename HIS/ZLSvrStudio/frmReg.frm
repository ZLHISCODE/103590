VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "请重新注册"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7245
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   5730
      TabIndex        =   3
      Top             =   465
      Width           =   1100
   End
   Begin VB.CommandButton cmdRegist 
      Caption         =   "重新注册(&R)…"
      Height          =   350
      Left            =   255
      TabIndex        =   1
      Top             =   465
      Width           =   1440
   End
   Begin MSComctlLib.ProgressBar pgbRegist 
      Height          =   165
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog DlgMain 
      Left            =   4455
      Top             =   615
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblRegist 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "正在注册，请稍等..."
      Height          =   210
      Left            =   1800
      TabIndex        =   2
      Top             =   570
      Visible         =   0   'False
      Width           =   1785
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean

Public Function ReReg() As Boolean
    mblnOK = False
    Me.Show vbModal
    ReReg = mblnOK
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRegist_Click()
    Dim strFile As String
    Dim strRegError As String
       
    With Me.DlgMain
        .FileName = ""
        .DialogTitle = "选择注册授权文件"
        .Filter = "(注册授权文件)|*.zcr"
        .ShowOpen
        If .FileName = "" Then Exit Sub
        strFile = .FileName
    End With
    
    Me.cmdRegist.Enabled = False
    Me.cmdExit.Enabled = False
    err = 0: On Error GoTo errHand
    
    lblRegist.Visible = True
    Me.MousePointer = vbHourglass
    
    If gobjRegister.zlRegBuild(strFile, pgbRegist) = False Then
        lblRegist.Visible = False
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    lblRegist.Visible = False
    Me.MousePointer = vbDefault
     
    Me.cmdRegist.Enabled = True
    Me.cmdExit.Enabled = True
    
    strRegError = gobjRegister.zlRegCheck(True)
    If strRegError = "" Then
        gcnOracle.Execute "call zltools.p_Reg_Apply()", , adCmdText
        strRegError = gobjRegister.zlRegCheck(False) '再次调用验证
        If strRegError = "" Then
            mblnOK = True
            MsgBox "注册授权信息已经应用！", vbInformation, gstrSysName
            Unload Me
        Else
            MsgBox strRegError & vbCrLf & "请检查zlRegAudit和zlRegFile表的[项目]字段长度，必须断掉所有客户端才能修正结构。", vbExclamation, gstrSysName
            mblnOK = False
        End If
    Else
        MsgBox strRegError & vbNewLine & "注册信息不正确，请重新注册！", vbExclamation, gstrSysName
        mblnOK = False
    End If
    Exit Sub
errHand:
    MsgBox "注册授权文件时出现错误，请检查！" & vbNewLine & err.Description, vbExclamation, Me.Caption
End Sub

