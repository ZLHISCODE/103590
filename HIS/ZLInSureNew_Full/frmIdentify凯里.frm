VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtCheckPass 
      Height          =   285
      Left            =   4695
      MaxLength       =   10
      TabIndex        =   36
      Top             =   225
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   34
      Top             =   240
      Width           =   1755
   End
   Begin VB.Frame fraInfo 
      Caption         =   "������Ϣ"
      Height          =   3255
      Left            =   68
      TabIndex        =   4
      Top             =   720
      Width           =   7035
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   4170
         TabIndex        =   18
         Tag             =   "30"
         Top             =   2820
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   4170
         TabIndex        =   17
         Tag             =   "20+22+24+26+28"
         Top             =   2385
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   4170
         TabIndex        =   16
         Tag             =   "18"
         Top             =   1965
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   4170
         TabIndex        =   15
         Tag             =   "11"
         Top             =   1530
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   4170
         TabIndex        =   14
         Tag             =   "5"
         Top             =   1095
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4170
         TabIndex        =   13
         Tag             =   "3"
         Top             =   675
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4170
         TabIndex        =   12
         Tag             =   "1"
         Top             =   240
         Width           =   2715
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   1290
         TabIndex        =   11
         Tag             =   "19"
         Top             =   2820
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   1290
         TabIndex        =   10
         Tag             =   "21+23+25+27"
         Top             =   2385
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   1290
         TabIndex        =   9
         Tag             =   "17"
         Top             =   1965
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1290
         TabIndex        =   8
         Tag             =   "6"
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1290
         TabIndex        =   7
         Tag             =   "4"
         Top             =   1095
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1290
         TabIndex        =   6
         Tag             =   "2"
         Top             =   675
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1290
         TabIndex        =   5
         Tag             =   "0"
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   13
         Left            =   3375
         TabIndex        =   32
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "ͳ�ﱨ���ۼ�"
         Height          =   180
         Index           =   11
         Left            =   3015
         TabIndex        =   31
         Top             =   2460
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����֧���ۼ�"
         Height          =   180
         Index           =   9
         Left            =   3015
         TabIndex        =   30
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���ⲡ����"
         Height          =   180
         Index           =   7
         Left            =   3195
         TabIndex        =   29
         Top             =   1605
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   4
         Left            =   3735
         TabIndex        =   28
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���"
         Height          =   180
         Index           =   3
         Left            =   3375
         TabIndex        =   27
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��ᱣ�Ϻ�"
         Height          =   180
         Index           =   1
         Left            =   3195
         TabIndex        =   26
         Top             =   315
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "סԺ�����ۼ�"
         Height          =   180
         Index           =   12
         Left            =   135
         TabIndex        =   25
         Top             =   2895
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����ͳ���ۼ�"
         Height          =   180
         Index           =   10
         Left            =   135
         TabIndex        =   24
         Top             =   2460
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���������ۼ�"
         Height          =   180
         Index           =   8
         Left            =   135
         TabIndex        =   23
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���֤����"
         Height          =   180
         Index           =   6
         Left            =   315
         TabIndex        =   22
         Top             =   1605
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   5
         Left            =   495
         TabIndex        =   21
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   855
         TabIndex        =   20
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "IC����"
         Height          =   180
         Index           =   0
         Left            =   675
         TabIndex        =   19
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "����(&R)"
      Height          =   400
      Left            =   83
      TabIndex        =   3
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   4853
      TabIndex        =   2
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   5993
      TabIndex        =   1
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmdChangPass 
      Caption         =   "������(&E)"
      Height          =   400
      Left            =   1238
      TabIndex        =   0
      Top             =   4065
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "ȷ�����룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   15
      Left            =   3615
      TabIndex        =   35
      Top             =   270
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmIdentify����.frx":0000
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "���룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   14
      Left            =   795
      TabIndex        =   33
      Top             =   285
      Width           =   675
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPass As String
Public mstrPatient As String, mstrOther As String

Public Function GetPatient(bytType As Byte) As String
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    Me.Show vbModal
    gstrҽ���� = txtInfo(1).Text
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrOther = "": mstrPatient = ""
    Me.Hide
End Sub

Private Sub cmdChangPass_Click()
    If txtCheckPass.Visible Then
        If txtCheckPass.Text = txtPassword.Text And txtPassword.Text <> "" Then
            glngReturn = ChangePass(mstrPass & "|" & txtCheckPass.Text)
            If glngReturn <> 0 Then
                MsgBox "�޸�����ʧ��", vbInformation, "�޸�����"
            Else
                MsgBox "�޸�����ɹ�", vbInformation, "�޸�����"
            End If
            cmdChangPass.Caption = "������(&E)"
            txtCheckPass.Visible = False
            txtPassword.Text = ""
            txtPassword.SetFocus
        ElseIf txtCheckPass.Text = txtPassword.Text Then
            MsgBox "��������벻��Ϊ��", vbInformation, "�޸�����"
            txtPassword.SetFocus
        ElseIf txtCheckPass.Text <> txtPassword.Text Then
            MsgBox "������������벻ͬ������������", vbInformation, "�޸�����"
            txtPassword.SetFocus
        End If
    Else
        txtCheckPass.Text = ""
        txtCheckPass.Visible = True
        txtPassword.SetFocus
        cmdChangPass.Caption = "�޸�(&E)"
    End If
End Sub

Private Sub cmdOK_Click()
    Dim cur֧���ۼ� As Currency, cur�����ۼ� As Currency
    If txtInfo(0).Text = "" Then
        MsgBox "���Ƚ��ж�������", vbInformation, "�����֤"
        Exit Sub
    End If
    mstrOther = "": mstrPatient = ""

    mstrPatient = txtInfo(0).Text & ";"                                 '0 ����
    mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '1 ҽ���ʺ�
    mstrPatient = mstrPatient & mstrPass & ";"                          '2 ����
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '3 ����
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '4 �Ա�
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '5 ��������
    mstrPatient = mstrPatient & txtInfo(6).Text & ";"                   '6 ���֤
    mstrPatient = mstrPatient & ";"                                     '7 ��λ����

    mstrOther = mstrOther & ";"                                         '8 ҽ����������(����)
    mstrOther = mstrOther & ";"                                         '9 ˳���
    mstrOther = mstrOther & ";"                                         '10 ���
    mstrOther = mstrOther & txtInfo(13).Text & ";"                      '11 ���
    mstrOther = mstrOther & ";"                                         '12 ��ǰ״̬
    mstrOther = mstrOther & ";"                                         '13 ����ID
    mstrOther = mstrOther & IIf(txtInfo(5).Text = "��ְ" Or txtInfo(5).Text = "��ְ������", "1", "2") & ";"
    mstrOther = mstrOther & ";"                                         '14 ����֤��
    mstrOther = mstrOther & ";"                                         '16 �����
    mstrOther = mstrOther & ";"                                         '17 �Ҷȼ�
    mstrOther = mstrOther & txtInfo(8).Text & ";"                       '18 �ʻ������ۼ�
    mstrOther = mstrOther & txtInfo(9).Text & ";"                       '19 �ʻ�֧���ۼ�
    mstrOther = mstrOther & txtInfo(10).Text & ";"                      '20 ����ͳ���ۼ�
    mstrOther = mstrOther & txtInfo(11).Text & ";"                      '21 ͳ�ﱨ���ۼ�
    mstrOther = mstrOther & txtInfo(12).Text & ";"                      '22 סԺ�����ۼ�
    mstrOther = mstrOther & ";"                                         '23 �������
    mstrOther = mstrOther & ";"                                         '24 ��������
    mstrOther = mstrOther & ";"                                         '25 �����ۼ�
    mstrOther = mstrOther & ";"                                         '26 ����ͳ���޶�
    Me.Hide
End Sub

Private Sub cmdRead_Click()
    Dim strPara() As String
    If txtPassword = "" Then
        MsgBox "����������", vbInformation, "�����֤"
        txtPassword.SetFocus
    End If
    mstrPass = txtPassword
    gstrReturn = ""
    glngReturn = GetPersonInfo(mstrPass, gstrReturn)
    If glngReturn <> 0 Then
        mstrPass = ""
        txtPassword.SetFocus
        MsgBox "����", vbInformation, "�����֤"
        Exit Sub
    End If
    strPara = Split(gstrReturn, "|")
    txtInfo(0).Text = strPara(0)
    txtInfo(1).Text = strPara(1)
    txtInfo(2).Text = strPara(2)
    txtInfo(3).Text = IIf(strPara(3) = "1", "��", "Ů")
    txtInfo(4).Text = strPara(4)
    '01.��ְ,02.����,03.����,04.�Ϻ��,05.��ְ������,06.���ݶ�����,07.���ݶ�����
    Select Case strPara(5)
        Case "01"
            txtInfo(5).Text = "��ְ"
        Case "02"
            txtInfo(5).Text = "����"
        Case "03"
            txtInfo(5).Text = "����"
        Case "04"
            txtInfo(5).Text = "�Ϻ��"
        Case "05"
            txtInfo(5).Text = "��ְ������"
        Case "06"
            txtInfo(5).Text = "���ݶ�����"
        Case "07"
            txtInfo(5).Text = "���ݶ�����"
    End Select
    txtInfo(6).Text = strPara(6)
    txtInfo(7).Text = strPara(10)
    txtInfo(8).Text = strPara(16)
    txtInfo(9).Text = strPara(17)
    txtInfo(10).Text = cNumber(strPara(19)) + cNumber(strPara(21)) + cNumber(strPara(23)) + cNumber(strPara(25)) + cNumber(strPara(27))
    txtInfo(11).Text = cNumber(strPara(20)) + cNumber(strPara(22)) + cNumber(strPara(24)) + cNumber(strPara(26)) + cNumber(strPara(28))
    txtInfo(12).Text = strPara(18)
    txtInfo(13).Text = strPara(29)
    txtPassword.Text = ""
    txtPassword.SetFocus
    cmdChangPass.Visible = True
End Sub

Private Sub txtCheckPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdChangPass_Click
    End If
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCheckPass.Visible = True Then
            txtCheckPass.SetFocus
        Else
            cmdRead_Click
        End If
    End If
End Sub
