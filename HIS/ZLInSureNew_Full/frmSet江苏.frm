VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3465
      TabIndex        =   5
      Top             =   1395
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2340
      TabIndex        =   4
      Top             =   1395
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   3
      Top             =   1200
      Width           =   4680
   End
   Begin VB.OptionButton optReadCard 
      Caption         =   "ʹ�ö�����"
      Height          =   180
      Index           =   1
      Left            =   885
      TabIndex        =   2
      Top             =   810
      Width           =   2250
   End
   Begin VB.OptionButton optReadCard 
      Caption         =   "�ֹ��������֤��"
      Height          =   180
      Index           =   0
      Left            =   885
      TabIndex        =   1
      Top             =   540
      Value           =   -1  'True
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ�񱾻��Ƿ�ʹ�ö�������"
      Height          =   180
      Left            =   330
      TabIndex        =   0
      Top             =   255
      Width           =   2340
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowMe(intInsure As Long) As Boolean
    Me.Show vbModal
    ShowMe = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strtemp As String
    If optReadCard(0).Value = True Then
        strtemp = "0"
    Else
        strtemp = "1"
    End If
    SaveSetting "ZLSOFT", "ҽ����Ϣ", "ReadCard", strtemp
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strtemp As String
    strtemp = GetSetting(appName:="ZLSOFT", Section:="ҽ����Ϣ", Key:="ReadCard", Default:="0")
    optReadCard(CInt(strtemp)).Value = True
End Sub
