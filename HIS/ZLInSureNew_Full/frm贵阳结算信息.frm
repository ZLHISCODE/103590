VERSION 5.00
Begin VB.Form frm����������Ϣ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   Icon            =   "frm����������Ϣ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5220
      TabIndex        =   36
      Top             =   4260
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -120
      TabIndex        =   37
      Top             =   4050
      Width           =   7125
   End
   Begin VB.TextBox txt�����ֱ��� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   35
      Top             =   3480
      Width           =   1485
   End
   Begin VB.TextBox txt���㷽ʽ 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   33
      Top             =   3090
      Width           =   1485
   End
   Begin VB.TextBox txt�����ʻ���� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   31
      Top             =   2700
      Width           =   1485
   End
   Begin VB.TextBox txtҽ�Ʋ���֧�� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   29
      Top             =   2310
      Width           =   1485
   End
   Begin VB.TextBox txt�����ܷ��� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1140
      Width           =   1485
   End
   Begin VB.TextBox txt���޶��Ը� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   27
      Top             =   1920
      Width           =   1485
   End
   Begin VB.TextBox txt�����ʻ�֧�� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   25
      Top             =   1530
      Width           =   1485
   End
   Begin VB.TextBox txt���ͳ���Ը� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   23
      Top             =   1140
      Width           =   1485
   End
   Begin VB.TextBox txt���ͳ��֧�� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   21
      Top             =   750
      Width           =   1485
   End
   Begin VB.TextBox txt����ͳ���Ը� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   19
      Top             =   360
      Width           =   1485
   End
   Begin VB.TextBox txt����ͳ��֧�� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   17
      Top             =   3480
      Width           =   1485
   End
   Begin VB.TextBox txt�������� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   15
      Top             =   3090
      Width           =   1485
   End
   Begin VB.TextBox txt�������� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   13
      Top             =   2700
      Width           =   1485
   End
   Begin VB.TextBox txt������ 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   11
      Top             =   2310
      Width           =   1485
   End
   Begin VB.TextBox txt�ҹ��Ը� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   1485
   End
   Begin VB.TextBox txtȫ�Է� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   7
      Top             =   1530
      Width           =   1485
   End
   Begin VB.TextBox txtҽԺ�ܷ��� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   750
      Width           =   1485
   End
   Begin VB.TextBox txtҽ���ܷ��� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   1485
   End
   Begin VB.Label lbl�����ֱ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ֱ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3960
      TabIndex        =   34
      Top             =   3540
      Width           =   900
   End
   Begin VB.Label lbl���㷽ʽ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���㷽ʽ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4140
      TabIndex        =   32
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lbl�����ʻ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ʻ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   30
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label lblҽ�Ʋ���֧�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ�Ʋ���֧��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   28
      Top             =   2370
      Width           =   1080
   End
   Begin VB.Label lbl�����ܷ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ܷ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lbl���޶��Ը� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���޶��Ը�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3960
      TabIndex        =   26
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label lbl�����ʻ�֧�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ʻ�֧��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   24
      Top             =   1590
      Width           =   1080
   End
   Begin VB.Label lbl���ͳ���Ը� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ͳ���Ը�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   22
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lbl���ͳ��֧�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ͳ��֧��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   20
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lbl����ͳ���Ը� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ͳ���Ը�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   18
      Top             =   420
      Width           =   1080
   End
   Begin VB.Label lbl����ͳ��֧�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ͳ��֧��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   16
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   14
      Top             =   3150
      Width           =   900
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   12
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label lbl������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   660
      TabIndex        =   10
      Top             =   2370
      Width           =   720
   End
   Begin VB.Label lbl�ҹ��Ը� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ҹ��Ը�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   660
      TabIndex        =   8
      Top             =   1980
      Width           =   720
   End
   Begin VB.Label lblȫ�Է� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ȫ�Է�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   840
      TabIndex        =   6
      Top             =   1590
      Width           =   540
   End
   Begin VB.Label lblҽԺ�ܷ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽԺ�ܷ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   810
      Width           =   900
   End
   Begin VB.Label lblҽ���ܷ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ���ܷ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm����������Ϣ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txt��������.Text = Val(Format(GetElemnetValue("STARTFEE"), "#0.00;-#0.00;0;"))
    txt���޶��Ը�.Text = Val(Format(GetElemnetValue("FEEOVER"), "#0.00;-#0.00;0;"))
    txt���ͳ��֧��.Text = Val(Format(GetElemnetValue("FUND2PAY"), "#0.00;-#0.00;0;"))
    txt���ͳ���Ը�.Text = Val(Format(GetElemnetValue("FUND2SELF"), "#0.00;-#0.00;0;"))
    txt�����ʻ����.Text = Val(Format(GetElemnetValue("ACCTBALANCE"), "#0.00;-#0.00;0;"))
    txt�����ʻ�֧��.Text = Val(Format(GetElemnetValue("ACCTPAY"), "#0.00;-#0.00;0;"))
    txt�ҹ��Ը�.Text = Val(Format(GetElemnetValue("FEESELF"), "#0.00;-#0.00;0;"))
    txt����ͳ��֧��.Text = Val(Format(GetElemnetValue("FUND1PAY"), "#0.00;-#0.00;0;"))
    txt����ͳ���Ը�.Text = Val(Format(GetElemnetValue("FUND1SELF"), "#0.00;-#0.00;0;"))
    txt�����ܷ���.Text = Val(Format(GetElemnetValue("CALFEEALL"), "#0.00;-#0.00;0;"))
    txt��������.Text = Val(Format(GetElemnetValue("ENTERSTARTFEE"), "#0.00;-#0.00;0;"))
    txtȫ�Է�.Text = Val(Format(GetElemnetValue("FEEOUT"), "#0.00;-#0.00;0;"))
    txtҽ���ܷ���.Text = Val(Format(GetElemnetValue("FEEALL"), "#0.00;-#0.00;0;"))
    txtҽ�Ʋ���֧��.Text = Val(Format(GetElemnetValue("FUND3PAY"), "#0.00;-#0.00;0;"))
    txtҽԺ�ܷ���.Text = Val(Format(GetElemnetValue("HOSPFEEALL"), "#0.00;-#0.00;0;"))
    txt������.Text = Val(Format(GetElemnetValue("ALLOWFUND"), "#0.00;-#0.00;0;"))
    
    txt�����ֱ���.Text = GetElemnetValue("SINGLEILLNESSCODE")
    txt���㷽ʽ.Text = GetElemnetValue("RECKONINGTYPE")

End Sub
