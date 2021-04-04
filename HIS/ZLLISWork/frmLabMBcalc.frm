VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmLabMBcalc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "批量调整OD值"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabMBcalc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -60
      TabIndex        =   5
      Top             =   930
      Width           =   4695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   3300
      TabIndex        =   4
      Top             =   1140
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Top             =   1140
      Width           =   1095
   End
   Begin VB.TextBox txt公式 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1545
      TabIndex        =   0
      Top             =   420
      Width           =   2820
   End
   Begin MSScriptControlCtl.ScriptControl Calc 
      Left            =   30
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Label lbl说明 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公式例子:""R-0.1""其中R表示OD值"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   150
      Width           =   2535
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入调整公式"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1260
   End
End
Attribute VB_Name = "frmLabMBcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrCalcl As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Public Function ShowMe(objfrm As Object) As String
    '功能       返回调整公式
    Me.Show vbModal, objfrm
    ShowMe = mstrCalcl
End Function

Private Sub cmdOK_Click()
    Dim str公式 As String
    str公式 = Me.txt公式.Text
    str公式 = Replace(UCase(str公式), "R", "1")
    If Trim(str公式) = "" Then
        MsgBox "请输入公式!", vbInformation, Me.Caption
        Me.txt公式.SetFocus
        Exit Sub
    End If
    If Trim(Calc.Eval(str公式)) = "" Then
        MsgBox "公式不正确，请检查!", vbInformation, Me.Caption
        Me.txt公式.SetFocus
        Exit Sub
    End If
    mstrCalcl = Me.txt公式
    Unload Me
End Sub

Private Sub txt公式_GotFocus()
    Me.txt公式.SelStart = 0
    Me.txt公式.SelLength = Len(Me.txt公式)
End Sub
