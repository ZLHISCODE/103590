VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ݱ�ʶ"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdChangPass 
      Caption         =   "������(&E)"
      Height          =   400
      Left            =   1230
      TabIndex        =   37
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   6255
      TabIndex        =   17
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   5115
      TabIndex        =   16
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "����(&R)"
      Height          =   400
      Left            =   75
      TabIndex        =   15
      Top             =   3825
      Width           =   1100
   End
   Begin VB.Frame fraInfo 
      Caption         =   "������Ϣ"
      Height          =   3675
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   7335
      Begin VB.CommandButton CmdSel 
         Caption         =   "��"
         Height          =   300
         Left            =   6930
         TabIndex        =   36
         Top             =   3255
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   15
         Left            =   4455
         TabIndex        =   35
         Top             =   3255
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1635
         TabIndex        =   25
         Top             =   240
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1635
         TabIndex        =   24
         Top             =   675
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1635
         TabIndex        =   23
         Top             =   1095
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1635
         TabIndex        =   22
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   1635
         TabIndex        =   21
         Top             =   1965
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   1635
         TabIndex        =   20
         Top             =   2385
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   1635
         TabIndex        =   19
         Top             =   2820
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   14
         Left            =   1635
         TabIndex        =   18
         Top             =   3255
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4455
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4455
         TabIndex        =   6
         Top             =   675
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   4455
         TabIndex        =   5
         Top             =   1095
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   4455
         TabIndex        =   4
         Top             =   1530
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   4455
         TabIndex        =   3
         Top             =   1965
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   4455
         TabIndex        =   2
         Top             =   2385
         Width           =   2775
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   4995
         TabIndex        =   1
         Top             =   2820
         Width           =   2235
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����ѡ��"
         Height          =   180
         Index           =   15
         Left            =   3660
         TabIndex        =   34
         Top             =   3330
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "ҽ����������"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   33
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "IC����"
         Height          =   180
         Index           =   2
         Left            =   1020
         TabIndex        =   32
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   5
         Left            =   1200
         TabIndex        =   31
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��λ����"
         Height          =   180
         Index           =   6
         Left            =   840
         TabIndex        =   30
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   8
         Left            =   840
         TabIndex        =   29
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "IC�����"
         Height          =   180
         Index           =   10
         Left            =   840
         TabIndex        =   28
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����ҽ������޶�"
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   27
         Top             =   2895
         Width           =   1440
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��ҽ���ۼ�֧��"
         Height          =   180
         Index           =   14
         Left            =   120
         TabIndex        =   26
         Top             =   3330
         Width           =   1440
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�����ʺ�"
         Height          =   180
         Index           =   1
         Left            =   3660
         TabIndex        =   14
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���֤����"
         Height          =   180
         Index           =   3
         Left            =   3480
         TabIndex        =   13
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   4
         Left            =   4020
         TabIndex        =   12
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��λ����"
         Height          =   180
         Index           =   7
         Left            =   3660
         TabIndex        =   11
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���"
         Height          =   180
         Index           =   9
         Left            =   3660
         TabIndex        =   10
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�𸶱�׼"
         Height          =   180
         Index           =   11
         Left            =   3660
         TabIndex        =   9
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����ҽ���ۼ�֧��"
         Height          =   180
         Index           =   13
         Left            =   3480
         TabIndex        =   8
         Top             =   2895
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean, mbytType As Byte
Public mstrPatient As String, mstrOther As String, mstr������ As String, mstr������� As String, mstr�𸶱�׼ As String
 
Public Function GetPatient(bytType As Byte, str������� As String) As String
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mbytType = bytType
    mstr������� = str�������
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = ""
    mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdChangPass_Click()
    initType
    mblnReturn = fl_changePassword(gstrҽ����������, gstrҽԺ����, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
    Else
        MsgBox "�����޸ĳɹ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdOK_Click()
    Dim cur֧���ۼ� As Currency, cur�����ۼ� As Currency
    With gstrOutPara
        '֧���ۼ� = ����ҽ���ۼ�֧�� + ��ҽ���ۼ�֧��
        cur֧���ۼ� = IIf(txtInfo(13).Text = "", 0, CCur(txtInfo(13).Text)) + IIf(txtInfo(14).Text = "", 0, CCur(txtInfo(14).Text))
        '�����ۼ� = IC����� + ֧���ۼ�
        cur�����ۼ� = IIf(txtInfo(10).Text = "", 0, CCur(txtInfo(10).Text)) + cur֧���ۼ�
        mstrOther = "": mstrPatient = ""
        
        mstrPatient = txtInfo(2).Text & ";"                                 '0 ����
        mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '1 ҽ���ʺ�
        mstrPatient = mstrPatient & ";"                                     '2 ����
        mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '3 ����
        mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '4 �Ա�
        mstrPatient = mstrPatient & txtInfo(8).Text & ";"                   '5 ��������
        mstrPatient = mstrPatient & txtInfo(5).Text & ";"                   '6 ���֤
        mstrPatient = mstrPatient & txtInfo(7).Text & "(" & txtInfo(6).Text & ");"                   '7 ��λ����/����
        
        mstrOther = mstrOther & txtInfo(0).Text & ";"                       '8 ҽ����������(����)
        mstrOther = mstrOther & ";"                                         '9 ˳���
        mstrOther = mstrOther & ";"                                         '10 ���
        mstrOther = mstrOther & txtInfo(10).Text & ";"                      '11 ���
        mstrOther = mstrOther & ";"                                         '12 ��ǰ״̬
        mstrOther = mstrOther & ";"                                         '13 ����ID
        mstrOther = mstrOther & IIf(txtInfo(9).Text = "��ְ", "1", IIf(txtInfo(9).Text = "����", "2", "3")) & ";"
        mstrOther = mstrOther & CLng(mstr�������) + 1 & ";"                '14 ����֤��
        mstrOther = mstrOther & ";"                                         '16 �����
        mstrOther = mstrOther & ";"                                         '17 �Ҷȼ�
        mstrOther = mstrOther & cur�����ۼ� & ";"                           '18 �ʻ������ۼ�
        mstrOther = mstrOther & cur֧���ۼ� & ";"                           '19 �ʻ�֧���ۼ�
        mstrOther = mstrOther & ";"                                         '20 ����ͳ���ۼ�
        mstrOther = mstrOther & ";"                                         '21 ͳ�ﱨ���ۼ�
        mstrOther = mstrOther & ";"                                         '22 סԺ�����ۼ�
        mstrOther = mstrOther & ";"                                         '23 �������
        mstrOther = mstrOther & txtInfo(11).Text & ";"                      '24 ��������
        mstrOther = mstrOther & ";"                                         '25 �����ۼ�
        mstrOther = mstrOther & ";"                                         '26 ����ͳ���޶�
    End With
    initType
    If mbytType = 0 Then      '���������סԺ���ȡ������
'    mblnReturn = fl_dall(gstrҽ����������, gstrҽԺ����, "2003121000031", gstrOutPara)
        mblnReturn = fl_reg(gstrҽ����������, gstrҽԺ����, 0, UserInfo.����, Format(zlDatabase.Currentdate, "yyyy-MM-dd"), "0", gstrOutPara)
        TrimType
        If mblnReturn = False Then
            MsgBox gstrOutPara.errtext, vbInformation, Me.Caption
            Exit Sub
        End If
        TrimType
        mstr������ = gstrOutPara.out1
    End If
    mstr�𸶱�׼ = txtInfo(11).Text
    Me.Hide
End Sub

Private Sub cmdRead_Click()
    cmdOK.SetFocus
    initType
    mblnReturn = fl_getybjgbm(gstrOutPara)
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, Me.Caption
        Exit Sub
    End If
    gstrҽ���������� = Trim(gstrOutPara.out1)
    initType
    mblnReturn = fl_readicxx(gstrҽ����������, gstrҽԺ����, "0", gstrOutPara)
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, Me.Caption
        Exit Sub
    End If
    TrimType
    With gstrOutPara
        txtInfo(0).Text = .out1
        txtInfo(1).Text = .out2
        txtInfo(2).Text = .out3
        txtInfo(3).Text = .out5
        txtInfo(4).Text = IIf(.out6 = "0", "��", "Ů")
        txtInfo(5).Text = .out4
        txtInfo(6).Text = .out7
        txtInfo(7).Text = .out8
        txtInfo(8).Text = .out9
        'Modified by zyb 2004-10-09
        'txtInfo(9).Text = IIf(.out10 = "21", "��ְ", IIf(.out10 = "22", "����", "�¸�"))
        txtInfo(9).Text = IIf(.out10 = "11", "��ְ", IIf(.out10 = "21", "����", "�¸�"))
        txtInfo(10).Text = .out11
        txtInfo(11).Text = .out12
        txtInfo(12).Text = .out13
        txtInfo(13).Text = .out14
        txtInfo(14).Text = .out15
    End With
End Sub

