VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4860
   Icon            =   "frmSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3540
      TabIndex        =   8
      Top             =   1890
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   7
      Top             =   1380
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ʱ���趨"
      Height          =   1605
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3075
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   300
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   92078083
         CurrentDate     =   40540
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ��1 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   690
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   92078083
         CurrentDate     =   40540.0833333333
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   92078083
         CurrentDate     =   40540.1666666667
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������ʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   675
         TabIndex        =   5
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ӡ������ʼʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   315
         TabIndex        =   3
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label lbl��ʼʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ǩ��ʼʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   315
         TabIndex        =   1
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Label Label3 
      Caption         =   "    ������Ǩ���߸���ָ���Ŀ�ʼʱ�����������Ǩ��������ӡ������ʼʱ��ʱֹͣ��Ȼ����д�ӡ������������������ʱ��Ϊֹ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "��ʼʱ��", Format(Me.dtp��ʼʱ��.Value, "HH:mm")
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "��ʼʱ��1", Format(Me.dtp��ʼʱ��1.Value, "HH:mm")
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "����ʱ��", Format(Me.dtp����ʱ��.Value, "HH:mm")
    
    Unload Me
End Sub

Private Sub Form_Load()
    dtp��ʼʱ��.Value = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "��ʼʱ��", "00:00")
    dtp��ʼʱ��1.Value = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "��ʼʱ��1", "02:00")
    dtp����ʱ��.Value = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "����ʱ��", "04:00")
End Sub
