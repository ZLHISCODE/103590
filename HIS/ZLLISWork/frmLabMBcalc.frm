VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmLabMBcalc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������ODֵ"
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
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -60
      TabIndex        =   5
      Top             =   930
      Width           =   4695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   3300
      TabIndex        =   4
      Top             =   1140
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   345
      Left            =   2040
      TabIndex        =   3
      Top             =   1140
      Width           =   1095
   End
   Begin VB.TextBox txt��ʽ 
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
   Begin VB.Label lbl˵�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʽ����:""R-0.1""����R��ʾODֵ"
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
      Caption         =   "�����������ʽ"
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
    '����       ���ص�����ʽ
    Me.Show vbModal, objfrm
    ShowMe = mstrCalcl
End Function

Private Sub cmdOK_Click()
    Dim str��ʽ As String
    str��ʽ = Me.txt��ʽ.Text
    str��ʽ = Replace(UCase(str��ʽ), "R", "1")
    If Trim(str��ʽ) = "" Then
        MsgBox "�����빫ʽ!", vbInformation, Me.Caption
        Me.txt��ʽ.SetFocus
        Exit Sub
    End If
    If Trim(Calc.Eval(str��ʽ)) = "" Then
        MsgBox "��ʽ����ȷ������!", vbInformation, Me.Caption
        Me.txt��ʽ.SetFocus
        Exit Sub
    End If
    mstrCalcl = Me.txt��ʽ
    Unload Me
End Sub

Private Sub txt��ʽ_GotFocus()
    Me.txt��ʽ.SelStart = 0
    Me.txt��ʽ.SelLength = Len(Me.txt��ʽ)
End Sub
