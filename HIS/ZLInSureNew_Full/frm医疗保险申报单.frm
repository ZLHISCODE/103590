VERSION 5.00
Begin VB.Form frmҽ�Ʊ����걨�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�Ʊ����걨��"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   Icon            =   "frmҽ�Ʊ����걨��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "���ɽ���"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   5
      Left            =   4950
      TabIndex        =   55
      Top             =   4080
      Width           =   4665
      Begin VB.TextBox txtҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   3270
         TabIndex        =   65
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt�����ʻ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   59
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   57
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txtͳ����� 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   3270
         TabIndex        =   61
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt���ͳ�� 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   63
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ʋ���"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   2490
         TabIndex        =   64
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl�����ʻ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʻ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   270
         TabIndex        =   58
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   270
         TabIndex        =   56
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblͳ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ�����"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   2475
         TabIndex        =   60
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl���ͳ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ͳ��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   255
         TabIndex        =   62
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�հ���סԺ"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   4
      Left            =   180
      TabIndex        =   24
      Top             =   4080
      Width           =   4635
      Begin VB.TextBox txtסԺ���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   1020
         TabIndex        =   28
         Top             =   690
         Width           =   585
      End
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   1020
         TabIndex        =   26
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt�����ʻ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   4
         Left            =   1020
         TabIndex        =   30
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txtҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   4
         Left            =   3270
         TabIndex        =   32
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblסԺ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl�����ʻ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʻ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   29
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ʋ���"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   4
         Left            =   2460
         TabIndex        =   31
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��֢סԺ"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   3
      Left            =   4950
      TabIndex        =   44
      Top             =   2370
      Width           =   4665
      Begin VB.TextBox txt���ͳ�� 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   52
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txtͳ����� 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   3270
         TabIndex        =   50
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   46
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt�����ʻ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   48
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txtҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   3270
         TabIndex        =   54
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl���ͳ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ͳ��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   255
         TabIndex        =   51
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblͳ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ�����"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   2475
         TabIndex        =   49
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   45
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl�����ʻ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʻ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   47
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ʋ���"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   2490
         TabIndex        =   53
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������סԺ"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   2
      Left            =   4950
      TabIndex        =   33
      Top             =   720
      Width           =   4635
      Begin VB.TextBox txtҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   3240
         TabIndex        =   43
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt�����ʻ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   37
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   35
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txtͳ����� 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   3240
         TabIndex        =   39
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt���ͳ�� 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   41
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ʋ���"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   2460
         TabIndex        =   42
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl�����ʻ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʻ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblͳ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ�����"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   2445
         TabIndex        =   38
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl���ͳ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ͳ��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   40
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   1
      Left            =   180
      TabIndex        =   13
      Top             =   2370
      Width           =   4635
      Begin VB.TextBox txt���ͳ�� 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   21
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txtͳ����� 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   3225
         TabIndex        =   19
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   15
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt�����ʻ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   17
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txtҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   3225
         TabIndex        =   23
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl���ͳ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ͳ��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   20
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblͳ����� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ�����"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   2430
         TabIndex        =   18
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl�����ʻ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʻ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ʋ���"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   2430
         TabIndex        =   22
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ͨ����"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Width           =   4635
      Begin VB.TextBox txtҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   12
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt�����ʻ� 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   10
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt�����˴� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   8
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblҽ�Ʋ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ʋ���"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl�����ʻ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʻ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl�����˴� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����˴�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.CommandButton cmd�걨 
      Caption         =   "�걨(&O)"
      Height          =   350
      Left            =   6360
      TabIndex        =   5
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&D)"
      Height          =   350
      Left            =   5220
      TabIndex        =   4
      Top             =   210
      Width           =   1100
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      Left            =   3390
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1665
   End
   Begin VB.ComboBox cbo�ں� 
      Height          =   300
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1665
   End
   Begin VB.Label lbl������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2610
      TabIndex        =   2
      Top             =   300
      Width           =   720
   End
   Begin VB.Label lbl�ں� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ں�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frmҽ�Ʊ����걨��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngID As Long              '0-����;�����ʾ����
Private mblnOK As Boolean           '�༭�ɹ�

Private Enum ����
    ��ͨ����
    ��������
    ������סԺ
    ��֢סԺ
    �հ���סԺ
    ���ɽ���
End Enum
'1��ҽ�Ʊ����걨�嵥�У������˴���ָ����ͨ������˴Σ����ɽ�������˴���ָ��ͨ������ѡ��ĵ����ֵĲ�������
'   a�������ߣ�����=1������֢������=2�������հ��ɣ�����=4�������ɣ�����=6��
'   b����ͨ����ѡ���˵����ֵľ����������

Public Function ShowME(ByVal lngID As Long) As Boolean
    mblnOK = False
    mlngID = lngID
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub cmdȡ��_Click()
    Dim int������� As Integer
    Dim lng������ As Long, lngסԺ���� As Long
    Dim str�ں� As String, str��ʼ���� As String, str�������� As String, str���ڽ������� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If mlngID <> 0 Then
        '����ģʽ
        Unload Me
        Exit Sub
    End If
    
    '���
    Call ClearCons
    
    str�ں� = Me.cbo�ں�.Text
    int������� = Me.cbo�������.ItemData(Me.cbo�������.ListIndex)
    str��ʼ���� = Mid(str�ں�, 1, 4) & "-" & Mid(str�ں�, 5, 2) & "-01 00:00:00"
    gstrSQL = " SELECT last_day(to_date('" & Mid(str��ʼ����, 1, 10) & "','yyyy-MM-dd')) from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�¶����һ��")
    str�������� = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd") & " 23:59:59"
    str���ڽ������� = Format(DateAdd("d", -1, str��ʼ����), "yyyy-MM-dd")
    
    '�����趨������ȡ��
    '����֢������Ǳ������
    '1����ͨ�������ֻ�շѣ��˴�=1�������շ����˷ѣ��˴�=1������������˷ѣ���˲����ǣ�
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS �����˴�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'�����ʻ�',NVL(C.��Ԥ��,0),0)),0) AS �����ʻ�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ�Ʋ���',NVL(C.��Ԥ��,0),0)),0) AS ҽ�Ʋ��� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B,����Ԥ����¼ C " & _
             " WHERE A.��¼ID=B.����ID AND A.ҽ�����='11' AND A.����=1 AND B.�����ֱ���_���� IS NULL " & _
             " AND A.��¼ID=C.����ID And A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ͨ����", int�������, TYPE_������, CDate(str��ʼ����), CDate(str��������))
    Me.txt�����˴�(��ͨ����).Text = Format(rsTemp!�����˴�, "#0;-#0; ;")
    Me.txt�����ʻ�(��ͨ����).Text = Format(rsTemp!�����ʻ�, "#0.00;-#0.00; ;")
    Me.txtҽ�Ʋ���(��ͨ����).Text = Format(rsTemp!ҽ�Ʋ���, "#0.00;-#0.00; ;")
    
    '2����������
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT B.ҽ����) AS �����˴�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'�����ʻ�',NVL(C.��Ԥ��,0),0)),0) AS �����ʻ�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ������',NVL(C.��Ԥ��,0),0)),0) AS ͳ�����, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'��ͳ��',NVL(C.��Ԥ��,0),0)),0) AS ���ͳ��, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ�Ʋ���',NVL(C.��Ԥ��,0),0)),0) AS ҽ�Ʋ��� " & _
             " FROM ���ս����¼ A,�����ʻ� B,����Ԥ����¼ C " & _
             " WHERE A.����ID=B.����ID And A.��¼ID=C.����ID AND A.ҽ�����='18' AND A.����=1 " & _
             " And A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������", int�������, TYPE_������, CDate(str��ʼ����), CDate(str��������))
    Me.txt�����˴�(��������).Text = Format(rsTemp!�����˴�, "#0;-#0; ;")
    Me.txt�����ʻ�(��������).Text = Format(rsTemp!�����ʻ�, "#0.00;-#0.00; ;")
    Me.txtͳ�����(��������).Text = Format(rsTemp!ͳ�����, "#0.00;-#0.00; ;")
    Me.txt���ͳ��(��������).Text = Format(rsTemp!���ͳ��, "#0.00;-#0.00; ;")
    Me.txtҽ�Ʋ���(��������).Text = Format(rsTemp!ҽ�Ʋ���, "#0.00;-#0.00; ;")
    
    '3��������סԺ
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS סԺ�˴�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'�����ʻ�',NVL(C.��Ԥ��,0),0)),0) AS �����ʻ�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ������',NVL(C.��Ԥ��,0),0)),0) AS ͳ�����, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'��ͳ��',NVL(C.��Ԥ��,0),0)),0) AS ���ͳ��, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ�Ʋ���',NVL(C.��Ԥ��,0),0)),0) AS ҽ�Ʋ��� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B,����Ԥ����¼ C " & _
             " WHERE A.��¼ID=B.����ID AND A.��¼ID=C.����ID AND A.����=2 " & _
             " AND B.���㷽ʽ=1 And A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������סԺ", int�������, TYPE_������, CDate(str��ʼ����), CDate(str��������))
    Me.txt�����˴�(������סԺ).Text = Format(rsTemp!סԺ�˴�, "#0;-#0; ;")
    Me.txt�����ʻ�(������סԺ).Text = Format(rsTemp!�����ʻ�, "#0.00;-#0.00; ;")
    Me.txtͳ�����(������סԺ).Text = Format(rsTemp!ͳ�����, "#0.00;-#0.00; ;")
    Me.txt���ͳ��(������סԺ).Text = Format(rsTemp!���ͳ��, "#0.00;-#0.00; ;")
    Me.txtҽ�Ʋ���(������סԺ).Text = Format(rsTemp!ҽ�Ʋ���, "#0.00;-#0.00; ;")
    
    '4����֢סԺ
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS סԺ�˴�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'�����ʻ�',NVL(C.��Ԥ��,0),0)),0) AS �����ʻ�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ������',NVL(C.��Ԥ��,0),0)),0) AS ͳ�����, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'��ͳ��',NVL(C.��Ԥ��,0),0)),0) AS ���ͳ��, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ�Ʋ���',NVL(C.��Ԥ��,0),0)),0) AS ҽ�Ʋ��� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B,����Ԥ����¼ C " & _
             " WHERE A.��¼ID=B.����ID AND A.��¼ID=C.����ID AND A.����=2 " & _
             " AND B.���㷽ʽ=2 And A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��֢סԺ", int�������, TYPE_������, CDate(str��ʼ����), CDate(str��������))
    Me.txt�����˴�(��֢סԺ).Text = Format(rsTemp!סԺ�˴�, "#0;-#0; ;")
    Me.txt�����ʻ�(��֢סԺ).Text = Format(rsTemp!�����ʻ�, "#0.00;-#0.00; ;")
    Me.txtͳ�����(��֢סԺ).Text = Format(rsTemp!ͳ�����, "#0.00;-#0.00; ;")
    Me.txt���ͳ��(��֢סԺ).Text = Format(rsTemp!���ͳ��, "#0.00;-#0.00; ;")
    Me.txtҽ�Ʋ���(��֢סԺ).Text = Format(rsTemp!ҽ�Ʋ���, "#0.00;-#0.00; ;")
    
    '5���հ���סԺ
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS סԺ�˴�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'�����ʻ�',NVL(C.��Ԥ��,0),0)),0) AS �����ʻ�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ�Ʋ���',NVL(C.��Ԥ��,0),0)),0) AS ҽ�Ʋ��� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B,����Ԥ����¼ C " & _
             " WHERE A.��¼ID=B.����ID AND A.��¼ID=C.����ID AND A.����=2 " & _
             " AND B.���㷽ʽ=4 And A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�հ���סԺ", int�������, TYPE_������, CDate(str��ʼ����), CDate(str��������))
    Me.txt�����˴�(�հ���סԺ).Text = Format(rsTemp!סԺ�˴�, "#0;-#0; ;")
    Me.txt�����ʻ�(�հ���סԺ).Text = Format(rsTemp!�����ʻ�, "#0.00;-#0.00; ;")
    Me.txtҽ�Ʋ���(�հ���סԺ).Text = Format(rsTemp!ҽ�Ʋ���, "#0.00;-#0.00; ;")
    '����סԺ������ԭ����������룺8��31����Ժδ��Ժ�ģ�סԺ����Ϊ���죻�����31���룬��Ժ����סԺ����Ϊ1�죩��
    '1�����ں�����:�������һ���ȥ��Ժʱ��
    '2���ò��˵��ڳ�Ժ:��Ժʱ����������һ��
    '3���������Ժ��:��Ժʱ���ȥ��Ժʱ��
    '4��������δ��Ժ:�������һ���ȥ�������һ��
    gstrSQL = " SELECT DISTINCT" & _
             "      A.������ˮ��,C.��Ժ����,C.��Ժ����,TO_CHAR(C.��Ժ����,'YYYYMM') AS ��Ժ�ں� " & _
             "  FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B,������ҳ C  " & _
             "  WHERE A.��¼ID=B.����ID AND A.����ID=C.����ID And A.��ҳID=C.��ҳID AND A.����=2  " & _
             "  AND B.���㷽ʽ=4 And A.����֢=[1] And A.����=[2]" & _
             "  AND A.����ʱ�� BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�հ���סԺ����", int�������, TYPE_������, CDate(str��ʼ����), CDate(str��������))
    With rsTemp
        Do While Not .EOF
            If !��Ժ�ں� <> str�ں� Then
                '������ǰ���Ժ
                If Not IsNull(!��Ժ����) Then
                    '���ڳ�Ժ
                    lngסԺ���� = DateDiff("d", str���ڽ�������, !��Ժ����)
                Else
                    lngסԺ���� = DateDiff("d", str���ڽ�������, str��������)
                End If
            Else
                '������Ժ
                If Not IsNull(!��Ժ����) Then
                    '���ڳ�Ժ
                    lngסԺ���� = DateDiff("d", !��Ժ����, !��Ժ����)
                Else
                    lngסԺ���� = DateDiff("d", !��Ժ����, str��������)
                End If
            End If
            If lngסԺ���� = 0 And Not IsNull(!��Ժ����) Then lngסԺ���� = 1
            lng������ = lng������ + lngסԺ����
        Loop
    End With
    Me.txtסԺ����(�հ���סԺ).Text = Format(lng������, "#0;-#0; ;")
    
    '6�����ɽ��㣨��ͨ������ѡ���˵����ֵģ�����סԺ���������㷽ʽ=6�ģ�
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS �����˴�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'�����ʻ�',NVL(C.��Ԥ��,0),0)),0) AS �����ʻ�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ������',NVL(C.��Ԥ��,0),0)),0) AS ͳ�����, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'��ͳ��',NVL(C.��Ԥ��,0),0)),0) AS ���ͳ��, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ�Ʋ���',NVL(C.��Ԥ��,0),0)),0) AS ҽ�Ʋ��� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B,����Ԥ����¼ C " & _
             " WHERE A.��¼ID=B.����ID AND A.ҽ�����='11' AND A.����=1 AND B.�����ֱ���_���� IS Not NULL " & _
             " AND A.��¼ID=C.����ID And A.����֢=" & int������� & " And A.����=" & TYPE_������ & _
             " AND A.����ʱ�� BETWEEN TO_DATE('" & str��ʼ���� & "','YYYY-MM-DD HH24:MI:SS') " & _
             " AND TO_DATE('" & str�������� & "','YYYY-MM-DD HH24:MI:SS') "
    gstrSQL = gstrSQL & _
             "UNION " & _
             "SELECT  " & _
             "        COUNT(DISTINCT A.������ˮ��) AS �����˴�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'�����ʻ�',NVL(C.��Ԥ��,0),0)),0) AS �����ʻ�, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ������',NVL(C.��Ԥ��,0),0)),0) AS ͳ�����, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'��ͳ��',NVL(C.��Ԥ��,0),0)),0) AS ���ͳ��, " & _
             "        NVL(SUM(DECODE(C.���㷽ʽ,'ҽ�Ʋ���',NVL(C.��Ԥ��,0),0)),0) AS ҽ�Ʋ��� " & _
             " FROM ���ս����¼ A,ZLGYYB.���㸽����Ϣ B,����Ԥ����¼ C " & _
             " WHERE A.��¼ID=B.����ID AND A.��¼ID=C.����ID AND A.����=2 " & _
             " AND B.���㷽ʽ=6 And A.����֢=[1] And A.����=[2]" & _
             " AND A.����ʱ�� BETWEEN [3] AND [4]"
    gstrSQL = " SELECT SUM(�����˴�) AS �����˴�,SUM(�����ʻ�) AS �����ʻ�,SUM(ͳ�����) AS ͳ�����," & _
              "       SUM(���ͳ��) AS ���ͳ��,SUM(ҽ�Ʋ���) AS ҽ�Ʋ���" & _
              " FROM (" & gstrSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ɽ���", int�������, TYPE_������, CDate(str��ʼ����), CDate(str��������))
    Me.txt�����˴�(���ɽ���).Text = Format(rsTemp!�����˴�, "#0;-#0; ;")
    Me.txt�����ʻ�(���ɽ���).Text = Format(rsTemp!�����ʻ�, "#0.00;-#0.00; ;")
    Me.txtͳ�����(���ɽ���).Text = Format(rsTemp!ͳ�����, "#0.00;-#0.00; ;")
    Me.txt���ͳ��(���ɽ���).Text = Format(rsTemp!���ͳ��, "#0.00;-#0.00; ;")
    Me.txtҽ�Ʋ���(���ɽ���).Text = Format(rsTemp!ҽ�Ʋ���, "#0.00;-#0.00; ;")
    
    Me.Tag = 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ClearCons
End Sub

Private Sub cmd�걨_Click()
    Dim str��ˮ�� As String
    On Error GoTo errHand
    
    If Val(Me.Tag) = 0 Then
        MsgBox "��ָ��������㡰ȡ������ť��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnGYYB.BeginTrans
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    'סԺ�������ֻҪ������˱��룬��ʽ����ʱ��Ҫ����ſ����ݼ�����
    Call InsertChild(mdomInput.documentElement, "PERIOD", cbo�ں�.Text)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", cbo�������.ItemData(cbo�������.ListIndex))
    Call InsertChild(mdomInput.documentElement, "MZPSNS", Val(txt�����˴�(��ͨ����).Text))                ' ��������˴�
    Call InsertChild(mdomInput.documentElement, "MZACCT", Val(txt�����ʻ�(��ͨ����).Text))
    Call InsertChild(mdomInput.documentElement, "MZFUND3", Val(txtҽ�Ʋ���(��ͨ����).Text))
    Call InsertChild(mdomInput.documentElement, "TMPSNS", Val(txt�����˴�(��������).Text))
    Call InsertChild(mdomInput.documentElement, "TMACCT", Val(txt�����ʻ�(��������).Text))
    Call InsertChild(mdomInput.documentElement, "TMFUND1", Val(txtͳ�����(��������).Text))
    Call InsertChild(mdomInput.documentElement, "TMFUND2", Val(txt���ͳ��(��������).Text))
    Call InsertChild(mdomInput.documentElement, "TMFUND3", Val(txtҽ�Ʋ���(��������).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1PSNS", Val(txt�����˴�(������סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1ACCT", Val(txt�����ʻ�(������סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1FUND1", Val(txtͳ�����(������סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1FUND2", Val(txt���ͳ��(������סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1FUND3", Val(txtҽ�Ʋ���(������סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2PSNS", Val(txt�����˴�(��֢סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2ACCT", Val(txt�����ʻ�(��֢סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2FUND1", Val(txtͳ�����(��֢סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2FUND2", Val(txt���ͳ��(��֢סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2FUND3", Val(txtҽ�Ʋ���(��֢סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY3PSNS", Val(txt�����˴�(�հ���סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY3DAYS", Val(txtסԺ����(�հ���סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY3ACCT", Val(txt�����ʻ�(�հ���סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY3FUND3", Val(txtҽ�Ʋ���(�հ���סԺ).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4PSNS", Val(txt�����˴�(���ɽ���).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4ACCT", Val(txt�����ʻ�(���ɽ���).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4FUND1", Val(txtͳ�����(���ɽ���).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4FUND2", Val(txt���ͳ��(���ɽ���).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4FUND3", Val(txtҽ�Ʋ���(���ɽ���).Text))
    '���ýӿ�
    If CommRecServer("APPRECM") = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    str��ˮ�� = GetElemnetValue("APPNO")
    
    '��������
    mlngID = GetNextID("���㵥", gcnGYYB)
    gstrSQL = "ZL_���㵥_INSERT(" & mlngID & ",0,'" & Me.cbo�ں�.Text & "'," & Me.cbo�������.ItemData(cbo�������.ListIndex) & "," & _
        "'" & Me.cbo�������.Text & "','" & gstrUserName & "',sysdate,'" & str��ˮ�� & "',NULL)"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    gstrSQL = "ZL_����ҽ��������ϸ_INSERT(" & mlngID & "," & Val(txt�����˴�(��ͨ����).Text) & "," & Val(txt�����ʻ�(��ͨ����).Text) & "," & Val(txtҽ�Ʋ���(��ͨ����).Text) & "," & _
            Val(txt�����˴�(��������).Text) & "," & Val(txt�����ʻ�(��������).Text) & "," & Val(txtͳ�����(��������).Text) & "," & Val(txt���ͳ��(��������).Text) & "," & Val(txtҽ�Ʋ���(��������).Text) & "," & _
            Val(txt�����˴�(������סԺ).Text) & "," & Val(txt�����ʻ�(������סԺ).Text) & "," & Val(txtͳ�����(������סԺ).Text) & "," & Val(txt���ͳ��(������סԺ).Text) & "," & Val(txtҽ�Ʋ���(������סԺ).Text) & "," & _
            Val(txt�����˴�(��֢סԺ).Text) & "," & Val(txt�����ʻ�(��֢סԺ).Text) & "," & Val(txtͳ�����(��֢סԺ).Text) & "," & Val(txt���ͳ��(��֢סԺ).Text) & "," & Val(txtҽ�Ʋ���(��֢סԺ).Text) & "," & _
            Val(txt�����˴�(�հ���סԺ).Text) & "," & Val(txtסԺ����(�հ���סԺ).Text) & "," & Val(txt�����ʻ�(�հ���סԺ).Text) & "," & Val(txtҽ�Ʋ���(�հ���סԺ).Text) & "," & _
            Val(txt�����˴�(���ɽ���).Text) & "," & Val(txt�����ʻ�(���ɽ���).Text) & "," & Val(txtͳ�����(���ɽ���).Text) & "," & Val(txt���ͳ��(���ɽ���).Text) & "," & Val(txtҽ�Ʋ���(���ɽ���).Text) & ")"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    gcnGYYB.CommitTrans
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnGYYB.RollbackTrans
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim str���� As String, str���� As String
    Dim rsData As New ADODB.Recordset
    
    If mlngID = 0 Then
        With cbo�������
            .Clear
            .AddItem "��ҵְ������ҽ�Ʊ���"
            .ItemData(.NewIndex) = 1
            .AddItem "��ҵ����ҽ�Ʊ���"
            .ItemData(.NewIndex) = 2
            .AddItem "������ҵ��λҽ�Ʊ���"
            .ItemData(.NewIndex) = 3
            .AddItem "����"
            .ItemData(.NewIndex) = 6
            .ListIndex = 0
        End With
        
        'ȱʡֻװ�����¡����¹��걨
        curDate = zlDatabase.Currentdate()
        str���� = Format(DateAdd("m", -1, curDate), "yyyyMM")
        str���� = Format(curDate, "yyyyMM")
        With cbo�ں�
            .Clear
            .AddItem str����
            .AddItem str����
            .ListIndex = 0
        End With
        Exit Sub
    End If
    
    '��ȡ�걨������
    gstrSQL = "SELECT  " & _
             "        A.ID, A.�ں�, A.�������, A.����Ա, A.���� ,B.�����˴�, B.��������ʻ�, B.����ҽ�Ʋ���, B.���������˴�, B.������������ʻ�, B.�����������ͳ��, B.����������ͳ��,  " & _
             "        B.��������ҽ�Ʋ���, B.������סԺ�˴�, B.������סԺ�����ʻ�, B.������סԺ����ͳ��, B.������סԺ���ͳ��, B.������סԺҽ�Ʋ���,  " & _
             "        B.��֢סԺ�˴�, B.��֢סԺ�����ʻ�, B.��֢סԺ����ͳ��, B.��֢סԺ���ͳ��, B.��֢סԺҽ�Ʋ���, B.�հ���סԺ�˴�, B.�հ���סԺ����,  " & _
             "        B.�հ���סԺ�����ʻ�, �հ���סԺҽ�Ʋ���, B.���ɽ����˴�, B.���ɽ�������ʻ�, B.���ɽ������ͳ��, B.���ɽ�����ͳ��, B.���ɽ���ҽ�Ʋ���, A.������ˮ��, A.������� " & _
             " FROM ���㵥 A, ����ҽ��������ϸ B " & _
             " WHERE A.ID=B.���㵥ID AND A.ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�걨������", mlngID)
    
    '����
    With rsData
        Me.cbo�������.AddItem !�������
        Me.cbo�������.ListIndex = 0
        Me.cbo�ں�.AddItem !�ں�
        Me.cbo�ں�.ListIndex = 0
        
        Me.txt�����˴�(��ͨ����).Text = Format(Nvl(!�����˴�, 0), "#0;-#0; ;")
        Me.txt�����ʻ�(��ͨ����).Text = Format(Nvl(!��������ʻ�, 0), "#0.00;-#0.00; ;")
        Me.txtҽ�Ʋ���(��ͨ����).Text = Format(Nvl(!����ҽ�Ʋ���, 0), "#0.00;-#0.00; ;")
        
        Me.txt�����˴�(��������).Text = Format(Nvl(!���������˴�, 0), "#0;-#0; ;")
        Me.txt�����ʻ�(��������).Text = Format(Nvl(!������������ʻ�, 0), "#0.00;-#0.00; ;")
        Me.txtͳ�����(��������).Text = Format(Nvl(!�����������ͳ��, 0), "#0.00;-#0.00; ;")
        Me.txt���ͳ��(��������).Text = Format(Nvl(!����������ͳ��, 0), "#0.00;-#0.00; ;")
        Me.txtҽ�Ʋ���(��������).Text = Format(Nvl(!��������ҽ�Ʋ���, 0), "#0.00;-#0.00; ;")
        
        Me.txt�����˴�(������סԺ).Text = Format(Nvl(!������סԺ�˴�, 0), "#0;-#0; ;")
        Me.txt�����ʻ�(������סԺ).Text = Format(Nvl(!������סԺ�����ʻ�, 0), "#0.00;-#0.00; ;")
        Me.txtͳ�����(������סԺ).Text = Format(Nvl(!������סԺ����ͳ��, 0), "#0.00;-#0.00; ;")
        Me.txt���ͳ��(������סԺ).Text = Format(Nvl(!������סԺ���ͳ��, 0), "#0.00;-#0.00; ;")
        Me.txtҽ�Ʋ���(������סԺ).Text = Format(Nvl(!������סԺҽ�Ʋ���, 0), "#0.00;-#0.00; ;")
        
        Me.txt�����˴�(��֢סԺ).Text = Format(Nvl(!��֢סԺ�˴�, 0), "#0;-#0; ;")
        Me.txt�����ʻ�(��֢סԺ).Text = Format(Nvl(!��֢סԺ�����ʻ�, 0), "#0.00;-#0.00; ;")
        Me.txtͳ�����(��֢סԺ).Text = Format(Nvl(!��֢סԺ����ͳ��, 0), "#0.00;-#0.00; ;")
        Me.txt���ͳ��(��֢סԺ).Text = Format(Nvl(!��֢סԺ���ͳ��, 0), "#0.00;-#0.00; ;")
        Me.txtҽ�Ʋ���(��֢סԺ).Text = Format(Nvl(!��֢סԺҽ�Ʋ���, 0), "#0.00;-#0.00; ;")
        
        Me.txt�����˴�(�հ���סԺ).Text = Format(Nvl(!�հ���סԺ�˴�, 0), "#0;-#0; ;")
        Me.txtסԺ����(�հ���סԺ).Text = Format(Nvl(!�հ���סԺ����, 0), "#0;-#0; ;")
        Me.txt�����ʻ�(�հ���סԺ).Text = Format(Nvl(!�հ���סԺ�����ʻ�, 0), "#0.00;-#0.00; ;")
        Me.txtҽ�Ʋ���(�հ���סԺ).Text = Format(Nvl(!�հ���סԺҽ�Ʋ���, 0), "#0.00;-#0.00; ;")
        
        Me.txt�����˴�(���ɽ���).Text = Format(Nvl(!���ɽ����˴�, 0), "#0;-#0; ;")
        Me.txt�����ʻ�(���ɽ���).Text = Format(Nvl(!���ɽ�������ʻ�, 0), "#0.00;-#0.00; ;")
        Me.txtͳ�����(���ɽ���).Text = Format(Nvl(!���ɽ������ͳ��, 0), "#0.00;-#0.00; ;")
        Me.txt���ͳ��(���ɽ���).Text = Format(Nvl(!���ɽ�����ͳ��, 0), "#0.00;-#0.00; ;")
        Me.txtҽ�Ʋ���(���ɽ���).Text = Format(Nvl(!���ɽ���ҽ�Ʋ���, 0), "#0.00;-#0.00; ;")
    End With
    
    '���ÿؼ�״̬
    Me.cbo�������.Enabled = False
    Me.cbo�ں�.Enabled = False
    
    cmd�걨.Visible = False
    cmdȡ��.Caption = "�˳�(&X)"
End Sub

Private Sub ClearCons()
    Me.Tag = ""
    Me.txt�����˴�(��ͨ����).Text = ""
    Me.txt�����ʻ�(��ͨ����).Text = ""
    Me.txtҽ�Ʋ���(��ͨ����).Text = ""
    
    Me.txt�����˴�(��������).Text = ""
    Me.txt�����ʻ�(��������).Text = ""
    Me.txtͳ�����(��������).Text = ""
    Me.txt���ͳ��(��������).Text = ""
    Me.txtҽ�Ʋ���(��������).Text = ""
    
    Me.txt�����˴�(������סԺ).Text = ""
    Me.txt�����ʻ�(������סԺ).Text = ""
    Me.txtͳ�����(������סԺ).Text = ""
    Me.txt���ͳ��(������סԺ).Text = ""
    Me.txtҽ�Ʋ���(������סԺ).Text = ""
    
    Me.txt�����˴�(��֢סԺ).Text = ""
    Me.txt�����ʻ�(��֢סԺ).Text = ""
    Me.txtͳ�����(��֢סԺ).Text = ""
    Me.txt���ͳ��(��֢סԺ).Text = ""
    Me.txtҽ�Ʋ���(��֢סԺ).Text = ""
    
    Me.txt�����˴�(�հ���סԺ).Text = ""
    Me.txtסԺ����(�հ���סԺ).Text = ""
    Me.txt�����ʻ�(�հ���סԺ).Text = ""
    Me.txtҽ�Ʋ���(�հ���סԺ).Text = ""
    
    Me.txt�����˴�(���ɽ���).Text = ""
    Me.txt�����ʻ�(���ɽ���).Text = ""
    Me.txtͳ�����(���ɽ���).Text = ""
    Me.txt���ͳ��(���ɽ���).Text = ""
    Me.txtҽ�Ʋ���(���ɽ���).Text = ""
End Sub
