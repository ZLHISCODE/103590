VERSION 5.00
Begin VB.Form frmXWSetParams 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PACS��������"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11040
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
   ScaleHeight     =   6225
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtStaticImage 
      Height          =   375
      Left            =   1680
      TabIndex        =   39
      Text            =   "http://127.0.0.1:8080/KeyImage.aspx?colid0=22&colvalue0=[@STU_NO]"
      Top             =   4440
      Width           =   9120
   End
   Begin VB.ComboBox cbo3DViewType 
      Height          =   360
      ItemData        =   "frmXWSetParams.frx":038A
      Left            =   6360
      List            =   "frmXWSetParams.frx":0394
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   5640
      Width           =   1812
   End
   Begin VB.TextBox txtWebServerPath 
      Height          =   375
      Left            =   1680
      TabIndex        =   34
      Text            =   "http://127.0.0.1:8080/TakeImage.aspx?colid0=22&colvalue0=[@STU_NO]"
      Top             =   3960
      Width           =   9120
   End
   Begin VB.TextBox txtSeriesSchemeNo 
      Height          =   375
      Left            =   3480
      TabIndex        =   33
      Text            =   "2"
      Top             =   5640
      Width           =   372
   End
   Begin VB.TextBox txtStudySchemeNo 
      Height          =   375
      Left            =   1680
      TabIndex        =   31
      Text            =   "1"
      Top             =   5640
      Width           =   372
   End
   Begin VB.TextBox txtXWOracleOwner 
      Height          =   375
      Left            =   1680
      TabIndex        =   28
      Text            =   "zlhis"
      Top             =   5040
      Width           =   2160
   End
   Begin VB.TextBox txtImageShare 
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Text            =   "DCMSHARE"
      Top             =   5040
      Width           =   1788
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "��¼��־"
      Height          =   255
      Left            =   9120
      TabIndex        =   25
      Top             =   5085
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "PACS �û�����"
      Height          =   2295
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   10575
      Begin VB.Frame Frame6 
         Caption         =   "���̿�¼"
         Height          =   1455
         Left            =   7080
         TabIndex        =   20
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtDVDBurnPswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   22
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtDVDBurnUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   21
            Top             =   360
            Width           =   2055
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
         Left            =   3600
         TabIndex        =   15
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtSendImageUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   17
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtSendImagePswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   16
            Top             =   840
            Width           =   2055
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
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtDelImagePswd 
            Height          =   375
            Left            =   1080
            TabIndex        =   12
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtDelImageUser 
            Height          =   375
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   2055
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
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   10575
      Begin VB.CommandButton Command1 
         Caption         =   "����(&T)"
         Height          =   400
         Left            =   9360
         TabIndex        =   36
         Top             =   460
         Width           =   1000
      End
      Begin VB.TextBox txtDBServerPswd 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   7320
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtDBServerUser 
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtDBServerIP 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         Height          =   255
         Left            =   6720
         TabIndex        =   7
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "�û���"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   540
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
      Left            =   9800
      TabIndex        =   1
      Top             =   5640
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   8520
      TabIndex        =   0
      Top             =   5640
      Width           =   1000
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "�ؼ�ͼ���ַ"
      Height          =   240
      Left            =   120
      TabIndex        =   40
      Top             =   4485
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "3D��Ƭ����"
      Height          =   240
      Left            =   5040
      TabIndex        =   37
      Top             =   5685
      Width           =   1200
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "WEB��Ƭ��ַ"
      Height          =   240
      Left            =   240
      TabIndex        =   35
      Top             =   4000
      Width           =   1320
   End
   Begin VB.Label Label14 
      Caption         =   "���з�����"
      Height          =   255
      Left            =   2280
      TabIndex        =   32
      Top             =   5685
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "��鷽����"
      Height          =   240
      Left            =   360
      TabIndex        =   30
      Top             =   5685
      Width           =   1200
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "�ӿڰ�ӵ����"
      Height          =   240
      Left            =   105
      TabIndex        =   29
      Top             =   5100
      Width           =   1440
   End
   Begin VB.Label Label11 
      Caption         =   "��ʷͼ����Ŀ¼"
      Height          =   255
      Left            =   4320
      TabIndex        =   27
      Top             =   5100
      Width           =   1965
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

Private Sub CmdOK_Click()
    Call SaveParams
    Unload Me
End Sub

Private Sub fillParams()
'------------------------------------------------
'���ܣ��������PACS�Ĳ���
'���أ�
'------------------------------------------------
    Dim i As Integer
    Dim str3DViewType As String
    
    On Error GoTo err
    
    '������ORACLE ģ������л�ȡ���������ݿ������IP��ַ���û���������
    txtDBServerIP = zlDatabase.GetPara("XW���ݿ������IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDBServerUser = zlDatabase.GetPara("XW���ݿ�������û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDBServerPswd = zlDatabase.GetPara("XW���ݿ����������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")

    txtXWOracleOwner = zlDatabase.GetPara("XWOracleӵ����", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    'txtWebServerIP = zlDatabase.GetPara("XWWEB������IP", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtWebServerPath = zlDatabase.GetPara("XWWEB��Ƭ��ַ", glngSys, G_LNG_XWPACSVIEW_MODULE, "http://127.0.0.1:8080/TakeImage.aspx?colid0=22&colvalue0=[@STU_NO]")
    txtStaticImage = zlDatabase.GetPara("XW�ؼ�ͼ���ַ", glngSys, G_LNG_XWPACSVIEW_MODULE, "http://127.0.0.1:8080/KeyImage.aspx?colid0=22&colvalue0=[@STU_NO]")
    
    str3DViewType = zlDatabase.GetPara("XW3D��Ƭ����", glngSys, G_LNG_XWPACSVIEW_MODULE, "Study3D")
    For i = 0 To cbo3DViewType.ListCount - 1
        If cbo3DViewType.list(i) = str3DViewType Then
            cbo3DViewType.ListIndex = i
            Exit For
        End If
    Next
    If cbo3DViewType.ListCount > 0 Then If cbo3DViewType.ListIndex < 0 Then cbo3DViewType.ListIndex = 0
    
    txtDelImageUser = zlDatabase.GetPara("XWɾ��ͼ���û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDelImagePswd = zlDatabase.GetPara("XWɾ��ͼ������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtSendImageUser = zlDatabase.GetPara("XW����ͼ���û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtSendImagePswd = zlDatabase.GetPara("XW����ͼ������", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtDVDBurnUser = zlDatabase.GetPara("XW���̿�¼�û���", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    txtDVDBurnPswd = zlDatabase.GetPara("XW���̿�¼����", glngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    txtImageShare = zlDatabase.GetPara("XW��ʷͼ����Ŀ¼", glngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    
    chkLog.value = IIf(Val(zlDatabase.GetPara("XW��¼�ӿ���־", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1, 1, 0)
    
    txtStudySchemeNo.Text = zlDatabase.GetPara("XW��鷽����", glngSys, G_LNG_XWPACSVIEW_MODULE, "1")
    txtSeriesSchemeNo.Text = zlDatabase.GetPara("XW���з�����", glngSys, G_LNG_XWPACSVIEW_MODULE, "2")
    
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
    
    Call zlDatabase.SetPara("XWOracleӵ����", txtXWOracleOwner.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    'Call zlDatabase.SetPara("XWWEB������IP", txtWebServerIP.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XWWEB��Ƭ��ַ", txtWebServerPath.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW�ؼ�ͼ���ַ", txtStaticImage.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW3D��Ƭ����", cbo3DViewType.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XWɾ��ͼ���û���", txtDelImageUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XWɾ��ͼ������", txtDelImagePswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW����ͼ���û���", txtSendImageUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW����ͼ������", txtSendImagePswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW���̿�¼�û���", txtDVDBurnUser.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW���̿�¼����", txtDVDBurnPswd.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW��ʷͼ����Ŀ¼", txtImageShare.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW��¼�ӿ���־", chkLog.value, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Call zlDatabase.SetPara("XW��鷽����", txtStudySchemeNo.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    Call zlDatabase.SetPara("XW���з�����", txtSeriesSchemeNo.Text, glngSys, G_LNG_XWPACSVIEW_MODULE)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo ErrHandle
    Call XWTestDBConnection(txtDBServerIP.Text, txtDBServerUser.Text, txtDBServerPswd.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

