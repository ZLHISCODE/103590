VERSION 5.00
Begin VB.Form frm�������ղ��� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������ղ���"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frm�������ղ���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk��ӡ 
      Caption         =   "���պ��ӡ�����嵥(B)"
      Height          =   300
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1650
      TabIndex        =   3
      Top             =   2175
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2835
      TabIndex        =   2
      Top             =   2175
      Width           =   1100
   End
   Begin VB.Frame fraSplit 
      Height          =   75
      Index           =   0
      Left            =   15
      TabIndex        =   1
      Top             =   1905
      Width           =   4245
   End
   Begin VB.Frame fraSplit 
      Height          =   75
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   4245
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frm�������ղ���.frx":030A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "  ��ѡ�������صĲ������չ������ݵ���ز�����"
      Height          =   405
      Left            =   765
      TabIndex        =   4
      Top             =   135
      Width           =   3105
   End
End
Attribute VB_Name = "frm�������ղ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long
Private mstrPrivs As String
Private mblnHavePriv As Boolean

 
 
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '����:�����ݿⱣ���������
    '����:����ɹ�����True,���򷵻�False
    '����:���˺�
    '����:2007/12/19
    '------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    If mlngModule = 201 Then
        
    Else
        Call zlDatabase.SetPara("��ӡ�����嵥", IIf(chk��ӡ.Value = 1, "1", "0"), glngSys, mlngModule)
    End If
    SaveSet = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdOK_Click()
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Public Sub ��������(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '------------------------------------------------------------------------------------
    '����:�����������
    '����:
    '����:
    '����:���˺�
    '�޸�:2007/12/21
    '------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnHavePriv = IsHavePrivs(mstrPrivs, "��������")
    
    If mlngModule = 201 Then
        chk��ӡ.Visible = False
    Else
        chk��ӡ.Value = IIf(Val(zlDatabase.GetPara("��ӡ�����嵥", glngSys, mlngModule, , Array(chk��ӡ), mblnHavePriv)) = 1, 1, 0)
        chk��ӡ.Visible = True
    End If
    
    frm�������ղ���.Show 1, frmMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'
'    If gbln���� = False Or gSystemPara.bln���� = False Then
'        ChkAuto.Visible = False
'        cmdOK.Top = chk����.Top + chk����.Height + 100
'        cmdCancel.Top = cmdOK.Top
'        Me.Height = cmdOK.Top + cmdOK.Height + 600
'    End If
'
'End Sub


