VERSION 5.00
Begin VB.Form frmWordCharCfgV2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ôʾ�����"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7470
   Icon            =   "frmWordCharCfgV2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   415
      Left            =   6120
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   415
      Left            =   4800
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtWordChar 
      Height          =   5655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmWordCharCfgV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsOk As Boolean
Private mlngModule As Long

 
Public Function zlShowWordCharCfg(ByVal lngModuleNo As Long, Owner As Object) As Boolean
    mblnIsOk = False
    
    mlngModule = lngModuleNo
    Show 1, Owner
    
    zlShowWordCharCfg = mblnIsOk
End Function

Private Sub cmdCancel_Click()
    mblnIsOk = False
    Unload Me
End Sub

Private Sub cmdSure_Click()
    If LenB(StrConv(txtWordChar.Text, vbFromUnicode)) >= 2000 Then
        MsgBoxD Me, "���ݳ��ȳ��� 2000�����������á�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    mblnIsOk = True
    
    Call zlDatabase.SetPara("���泣�ôʾ�", txtWordChar.Text, glngSys, mlngModule)
    
    Unload Me
End Sub

Private Sub Form_Load()
    txtWordChar.Text = zlDatabase.GetPara("���泣�ôʾ�", glngSys, mlngModule)
End Sub
 
