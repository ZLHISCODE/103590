VERSION 5.00
Begin VB.Form frmIdentify����ũ��_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "frmIdentify����ũ��_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   780
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ComboBox cbo�Ա� 
      Height          =   300
      ItemData        =   "frmIdentify����ũ��_����.frx":000C
      Left            =   780
      List            =   "frmIdentify����ũ��_����.frx":0019
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   690
      TabIndex        =   4
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   1875
      TabIndex        =   5
      Top             =   1920
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   105
      Left            =   -60
      TabIndex        =   6
      Top             =   1800
      Width           =   3555
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   780
      TabIndex        =   1
      Top             =   330
      Width           =   2295
   End
   Begin VB.Label lb���� 
      Caption         =   "����"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "�Ա�"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   390
      Width           =   360
   End
End
Attribute VB_Name = "frmIdentify����ũ��_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytType As Byte, mstrOther As String, mstrPatient As String

Public Function GetPatient(bytType As Byte) As String
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mstrPatient = ""
    mstrOther = ""
    
    mbytType = bytType
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cbo�Ա�_Change()
 If cbo�Ա�.ListIndex = -1 Then cbo�Ա�.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngSequence As Long
    
    If Trim(txt����.Text) = "" Then
        MsgBox "�����벡��������", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    '������Ϣ��δ������ֻ������Ա���������Ϊҽ�����뿨��
    lngSequence = zlDatabase.GetNextID("��Ա��")
   
    mstrOther = "": mstrPatient = ""
    
    mstrPatient = lngSequence & ";"                                     '0 ����
    mstrPatient = mstrPatient & lngSequence & ";"                       '1 ҽ���ʺ�
    mstrPatient = mstrPatient & ";"                                     '2 ����
    mstrPatient = mstrPatient & Me.txt����.Text & ";"                   '3 ����
    mstrPatient = mstrPatient & Me.cbo�Ա�.Text & ";"                   '4 �Ա�
    mstrPatient = mstrPatient & IIf(Trim(txt����.Text) = "", Format(zlDatabase.Currentdate, "yyyy-mm-dd"), Get��������("", Val(txt����.Text))) & ";"                                    '5 ��������
    mstrPatient = mstrPatient & ";"                                     '6 ���֤
    mstrPatient = mstrPatient & ";"                                     '7 ��λ����/����
    
    mstrOther = mstrOther & ";"                                         '8 ҽ����������(����)
    mstrOther = mstrOther & ";"                                         '9 ˳���
    mstrOther = mstrOther & ";"                                         '10 ���
    mstrOther = mstrOther & ";"                                         '11 ���
    mstrOther = mstrOther & ";"                                         '12 ��ǰ״̬
    mstrOther = mstrOther & ";"                                         '13 ����ID
    mstrOther = mstrOther & ";"
    mstrOther = mstrOther & ";"                                         '15 ����֤��
    mstrOther = mstrOther & ";"                                         '16 �����
    mstrOther = mstrOther & ";"                                         '17 �Ҷȼ�
    mstrOther = mstrOther & ";"                                         '18 �ʻ������ۼ�
    mstrOther = mstrOther & ";"                                         '19 �ʻ�֧���ۼ�
    mstrOther = mstrOther & ";"                                         '20 ����ͳ���ۼ�
    mstrOther = mstrOther & ";"                                         '21 ͳ�ﱨ���ۼ�
    mstrOther = mstrOther & ";"                                         '22 סԺ�����ۼ�
    mstrOther = mstrOther & ";"                                         '23 �������
    mstrOther = mstrOther & ";"                                         '24 ��������
    mstrOther = mstrOther & ";"                                         '25 �����ۼ�
    mstrOther = mstrOther & ";"                                         '26 ����ͳ���޶�
    Me.Hide
End Sub

Private Sub com�Ա�_Change()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        Call SendKeys("{Tab}")
    End If
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
