VERSION 5.00
Begin VB.Form frmCreateAdvicedSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼�߼�����"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "frmCreateAdvicedSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdDefault 
      Caption         =   "Ĭ��(&D)"
      Height          =   350
      Left            =   120
      TabIndex        =   12
      Top             =   2460
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6450
      TabIndex        =   11
      Top             =   2460
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4740
      TabIndex        =   10
      Top             =   2460
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "д��ѡ��"
      Height          =   2265
      Left            =   4260
      TabIndex        =   1
      Top             =   60
      Width           =   3915
      Begin VB.CheckBox ChkWriterAutoVerify 
         Caption         =   "�Զ�У���ļ�����"
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   1710
         Width           =   3195
      End
      Begin VB.CheckBox ChkWriterBufferProof 
         Caption         =   "������У��"
         CausesValidation=   0   'False
         Height          =   345
         Left            =   180
         TabIndex        =   8
         Top             =   1290
         Width           =   2235
      End
      Begin VB.CheckBox ChkWriterTestWriter 
         Caption         =   "����д��(DVD��ʽ��Ч)"
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   960
         Width           =   3195
      End
      Begin VB.CheckBox ChkWriterCloseDisk 
         Caption         =   "��������(������д��)"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   600
         Width           =   2805
      End
      Begin VB.CheckBox ChkWriterCheckImage 
         Caption         =   "��ʹ�ø��ٻ���д��(��CD_RW)"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   3525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ѡ��"
      Height          =   2265
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4065
      Begin VB.CheckBox ChkDataHighCompatibilityMode 
         Caption         =   "�߼�����DVD(д���ļ���СҪ�ﵽ1GB)"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   990
         Width           =   3585
      End
      Begin VB.CheckBox ChkDataCDRWMode 
         Caption         =   "CDR/Wд��ʱʹ��ģʽ"
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   630
         Width           =   3555
      End
      Begin VB.CheckBox ChkDataUseJoliet 
         Caption         =   "ʹ��Joliet(ʹ�ļ������ɴ�64���ַ�)"
         Height          =   345
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Width           =   3705
      End
   End
End
Attribute VB_Name = "frmCreateAdvicedSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDefault_Click()
    '����
    ChkDataUseJoliet.Value = 1
    ChkDataCDRWMode.Value = 0
    ChkDataHighCompatibilityMode.Value = 0
    'д��
    ChkWriterCheckImage.Value = 0
    ChkWriterCloseDisk.Value = 1
    ChkWriterTestWriter.Value = 1
    ChkWriterBufferProof.Value = 1
    ChkWriterAutoVerify.Value = 1
End Sub

Private Sub cmdOK_Click()
    SaveOrLoadSetup 1
    Unload Me
End Sub



Private Sub Form_Load()
    SaveOrLoadSetup 2
End Sub

Sub SaveOrLoadSetup(SaveOrLoad As Integer)
    '����򱣴����
    'SaveOrLoad = 1 ���� = 2 ����
    
    Dim intUseJoliet As Integer
    Dim intCDRWMode  As Integer
    Dim blHighCompatibilityMode  As Boolean
    Dim blCheckImage As Boolean
    Dim blCloseDisk As Boolean
    Dim blTestWriter As Boolean
    Dim blBufferProof As Boolean
    Dim blAutoVerify As Boolean
    
    '����
    If SaveOrLoad = 1 Then
        intUseJoliet = IIf((ChkDataUseJoliet.Value = vbChecked), vtyISO9660_JOLIET, vtyISO9660_ONLY)
        intCDRWMode = IIf((ChkDataCDRWMode.Value = vbChecked), wtpDataMode2_XA, wtpDataMode1)
        blHighCompatibilityMode = (ChkDataHighCompatibilityMode.Value = vbChecked)
        blCheckImage = (ChkWriterCheckImage.Value = vbChecked)
        blCloseDisk = (ChkWriterCloseDisk.Value = vbChecked)
        blTestWriter = (ChkWriterTestWriter.Value = vbChecked)
        blBufferProof = (ChkWriterBufferProof.Value = vbChecked)
        blAutoVerify = (ChkWriterAutoVerify.Value = vbChecked)
    
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "ʹ��Joliet", intUseJoliet
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "ʹ��CDRWģʽ", intCDRWMode
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "�߼���DVDģʽ", blHighCompatibilityMode
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "��ʹ�ø��ٻ���", blCheckImage
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "�رչ���", blCloseDisk
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "����д��", blTestWriter
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "����У��", blBufferProof
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "�Զ�����У��", blAutoVerify
    Else
        ChkDataUseJoliet.Value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "ʹ��Joliet", 1)
        intCDRWMode = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "ʹ��CDRWģʽ", 1)
        ChkDataCDRWMode.Value = IIf(intCDRWMode = 2, 1, 0)
        blHighCompatibilityMode = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "�߼���DVDģʽ", 0)
        ChkDataHighCompatibilityMode.Value = IIf(blHighCompatibilityMode, 1, 0)
        blCheckImage = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "��ʹ�ø��ٻ���", 0)
        ChkWriterCheckImage.Value = IIf(blCheckImage, 1, 0)
        blCloseDisk = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "�رչ���", 1)
        ChkWriterCloseDisk.Value = IIf(blCloseDisk, 1, 0)
        blTestWriter = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "����д��", 1)
        ChkWriterTestWriter.Value = IIf(blTestWriter, 1, 0)
        blBufferProof = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "����У��", 1)
        ChkWriterBufferProof.Value = IIf(blBufferProof, 1, 0)
        blAutoVerify = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��¼����", "�Զ�����У��", 1)
        ChkWriterAutoVerify.Value = IIf(blAutoVerify, 1, 0)
    End If
    
End Sub
