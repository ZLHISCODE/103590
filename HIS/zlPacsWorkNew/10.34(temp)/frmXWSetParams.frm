VERSION 5.00
Begin VB.Form frmXWSetParams 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PACS��������"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmXWSetParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtImageShare 
      Height          =   375
      Left            =   6495
      TabIndex        =   29
      Text            =   "DCMSHARE"
      Top             =   3540
      Width           =   2145
   End
   Begin VB.TextBox txtWebServerIP 
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   2955
      Width           =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����(&T)"
      Height          =   400
      Left            =   7575
      TabIndex        =   26
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "��¼��־"
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   4275
      Width           =   3975
   End
   Begin VB.Frame Frame4 
      Caption         =   "PACS �û�����"
      Height          =   5175
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame6 
         Caption         =   "���̿�¼"
         Height          =   1455
         Left            =   240
         TabIndex        =   20
         Top             =   3480
         Width           =   3975
         Begin VB.TextBox txtDVDBurnPswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   22
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtDVDBurnUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   21
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label10 
            Caption         =   "����"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "�û���"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "����ͼ��"
         Height          =   1455
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   3975
         Begin VB.TextBox txtSendImageUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   17
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtSendImagePswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   16
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "�û���"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "����"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   900
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ɾ��ͼ��"
         Height          =   1455
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3975
         Begin VB.TextBox txtDelImagePswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   12
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtDelImageUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "����"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "�û���"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   420
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PACS ���ݿ������"
      Height          =   2055
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   3975
      Begin VB.TextBox txtDBServerPswd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtDBServerUser 
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtDBServerIP 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "�û���"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "������"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   7200
      TabIndex        =   1
      Top             =   4920
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   5280
      TabIndex        =   0
      Top             =   4920
      Width           =   1000
   End
   Begin VB.Label Label11 
      Caption         =   "��ʷͼ����Ŀ¼"
      Height          =   255
      Left            =   4680
      TabIndex        =   30
      Top             =   3600
      Width           =   1725
   End
   Begin VB.Label Label6 
      Caption         =   "WEB������IP"
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   3015
      Width           =   1380
   End
End
Attribute VB_Name = "frmXWSetParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function zlShowMe(frmParent As Form) As Long
'------------------------------------------------
'���ܣ�������PACS�Ĳ������ô���
'���أ�
'------------------------------------------------
    On Error GoTo err
    
    Call fillParams
    Me.Show 1, frmParent
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveParams
    Unload Me
End Sub

Private Sub fillParams()
'------------------------------------------------
'���ܣ��������PACS�Ĳ���
'���أ�
'------------------------------------------------
    On Error GoTo err
    
    '������ORACLE ģ������л�ȡ���������ݿ������IP��ַ���û���������
    txtDBServerIP = zlDatabase.GetPara("XW���ݿ������IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDBServerUser = zlDatabase.GetPara("XW���ݿ�������û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDBServerPswd = zlDatabase.GetPara("XW���ݿ����������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtWebServerIP = zlDatabase.GetPara("XWWEB������IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtDelImageUser = zlDatabase.GetPara("XWɾ��ͼ���û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDelImagePswd = zlDatabase.GetPara("XWɾ��ͼ������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtSendImageUser = zlDatabase.GetPara("XW����ͼ���û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtSendImagePswd = zlDatabase.GetPara("XW����ͼ������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtDVDBurnUser = zlDatabase.GetPara("XW���̿�¼�û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDVDBurnPswd = zlDatabase.GetPara("XW���̿�¼����", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtImageShare = zlDatabase.GetPara("XW��ʷͼ����Ŀ¼", glngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    
    chkLog.value = IIf(Val(zlDatabase.GetPara("XW��¼�ӿ���־", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1, 1, 0)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SaveParams()
'------------------------------------------------
'���ܣ���������PACS�Ĳ���
'���أ�
'------------------------------------------------
    On Error GoTo err
    
    '������PACS�Ĳ������ñ��浽����ORACLE ģ�������
    Call zlDatabase.SetPara("XW���ݿ������IP", txtDBServerIP.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW���ݿ�������û���", txtDBServerUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW���ݿ����������", txtDBServerPswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XWWEB������IP", txtWebServerIP.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XWɾ��ͼ���û���", txtDelImageUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XWɾ��ͼ������", txtDelImagePswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW����ͼ���û���", txtSendImageUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW����ͼ������", txtSendImagePswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW���̿�¼�û���", txtDVDBurnUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW���̿�¼����", txtDVDBurnPswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW��ʷͼ����Ŀ¼", txtImageShare.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW��¼�ӿ���־", chkLog.value, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo errHandle
    Call XWTestDBConnection(txtDBServerIP.Text, txtDBServerUser.Text, txtDBServerPswd.Text)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

