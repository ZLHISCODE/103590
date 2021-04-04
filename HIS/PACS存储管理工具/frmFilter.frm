VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����(ָ���ֶ��鵵��Χ)"
   ClientHeight    =   2550
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   6870
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   5430
      TabIndex        =   16
      Top             =   2100
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   300
      Left            =   4020
      TabIndex        =   15
      Top             =   2100
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.ComboBox cobDevice 
         Height          =   300
         ItemData        =   "frmFilter.frx":000C
         Left            =   1110
         List            =   "frmFilter.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   2085
      End
      Begin VB.TextBox txtFStudy 
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   660
         Width           =   2085
      End
      Begin VB.TextBox txtEStudy 
         Height          =   315
         Left            =   4410
         TabIndex        =   3
         Top             =   660
         Width           =   2100
      End
      Begin VB.ComboBox cobArchiveState 
         Height          =   300
         ItemData        =   "frmFilter.frx":0010
         Left            =   1110
         List            =   "frmFilter.frx":001D
         TabIndex        =   2
         Text            =   "δ�鵵"
         Top             =   1410
         Width           =   2085
      End
      Begin VB.ComboBox cobStorageDevice 
         Height          =   300
         ItemData        =   "frmFilter.frx":0045
         Left            =   4410
         List            =   "frmFilter.frx":0047
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2100
      End
      Begin MSComCtl2.DTPicker DTPETime 
         Height          =   315
         Left            =   4410
         TabIndex        =   5
         Top             =   1035
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
         Format          =   16646147
         CurrentDate     =   38169
      End
      Begin MSComCtl2.DTPicker DTPFTime 
         Height          =   315
         Left            =   1110
         TabIndex        =   6
         Top             =   1035
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd hh:mm:ss"
         Format          =   16646147
         CurrentDate     =   38169
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��"
         Height          =   195
         Left            =   315
         TabIndex        =   14
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   3615
         TabIndex        =   13
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ӱ������"
         Height          =   180
         Left            =   330
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ����"
         Height          =   180
         Left            =   165
         TabIndex        =   11
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   3435
         TabIndex        =   10
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�鵵״̬"
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�洢�豸"
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IFFilter As Boolean                         '�Ƿ����
Private Sub CmdCancel_Click()
    IFFilter = False
    Unload Me
End Sub
Private Sub CmdOK_Click()
    SaveSetting "ZLSOFT", "����ģ��\�鵵����\����", "Ӱ������", cobDevice.Text
    SaveSetting "ZLSOFT", "����ģ��\�鵵����\����", "�����豸", cobStorageDevice.ItemData(cobStorageDevice.ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\�鵵����\����", "�����豸��", cobStorageDevice.Text
    SaveSetting "ZLSOFT", "����ģ��\�鵵����\����", "��ʼ����", txtFStudy.Text
    SaveSetting "ZLSOFT", "����ģ��\�鵵����\����", "��������", txtEStudy.Text
    SaveSetting "ZLSOFT", "����ģ��\�鵵����\����", "��ʼʱ��", DTPFTime
    SaveSetting "ZLSOFT", "����ģ��\�鵵����\����", "����ʱ��", DTPETime
    SaveSetting "ZLSOFT", "����ģ��\�鵵����\����", "�鵵״̬", cobArchiveState.Text
    IFFilter = True
    Unload Me
End Sub
Private Sub cobArchiveState_Click()
    If Me.cobArchiveState.Text = "�ѹ鵵δɾ��" Then
        Me.cobStorageDevice.ListIndex = 0
        Me.cobStorageDevice.Enabled = False
    Else
        Me.cobStorageDevice.Enabled = True
    End If
End Sub
Private Sub Form_Load()
    Dim strSQL As String
    Dim tmpset As ADODB.Recordset
    strSQL = "select /*+ Rule */ distinct Ӱ����� from Ӱ������Ŀ where ������ĿID>0"
    cobDevice.Clear
    Set tmpset = gcnOracle.Execute(strSQL)
    cobDevice.AddItem "��������"
    Do While Not tmpset.EOF
        cobDevice.AddItem tmpset!Ӱ�����
        tmpset.MoveNext
    Loop
    cobDevice.Text = "��������"
    
    Me.cobStorageDevice.AddItem "�����豸"
    Me.cobStorageDevice.ItemData(Me.cobStorageDevice.NewIndex) = cAllStorageDevice
    strSQL = "select * from Ӱ���豸Ŀ¼ where ����=1"
    Set tmpset = gcnOracle.Execute(strSQL)
    Do While Not tmpset.EOF
        Me.cobStorageDevice.AddItem tmpset!�豸��
        Me.cobStorageDevice.ItemData(Me.cobStorageDevice.NewIndex) = tmpset!�豸��
        tmpset.MoveNext
    Loop
    Me.cobStorageDevice.ListIndex = 0
    
    '��ע����ж�ȡ�ϴα����ֵ
    cobDevice = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "Ӱ������", "��������")
    cobStorageDevice = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "�����豸��", cobStorageDevice.Text)
    txtFStudy.Text = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "��ʼ����", "")
    txtEStudy.Text = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "��������", "")
    DTPFTime = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "��ʼʱ��", zlDatabase.Currentdate - 90)
    DTPETime = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "����ʱ��", zlDatabase.Currentdate - 30)
    cobArchiveState.Text = GetSetting("ZLSOFT", "����ģ��\�鵵����\����", "�鵵״̬", "δ�鵵")
    
End Sub
Public Function ShowMe(frmObj As Object) As Boolean
    Me.Show vbModal, frmObj
    ShowMe = IFFilter
End Function
