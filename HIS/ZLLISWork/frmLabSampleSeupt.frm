VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmLabSampleSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8595
   Icon            =   "frmLabSampleSeupt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   6810
      TabIndex        =   10
      Top             =   3630
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   5220
      TabIndex        =   9
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   1365
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   3525
      Begin VB.CheckBox chkFindMove 
         Caption         =   "���ҵ����˺󽹵��ƶ�����������"
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   3135
      End
      Begin VB.CheckBox ChkContinuous 
         Caption         =   "��������������"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   900
         Width           =   2595
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ʹ�õ��Թ�"
      Height          =   3375
      Left            =   3630
      TabIndex        =   1
      Top             =   90
      Width           =   4905
      Begin XtremeReportControl.ReportControl rptCuvette 
         Height          =   2985
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4665
         _Version        =   589884
         _ExtentX        =   8229
         _ExtentY        =   5265
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         SkipGroupsFocus =   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   1905
      Left            =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   3525
      Begin VB.CheckBox chkBackBill 
         Caption         =   "����ɲɼ����ӡ��ִ��"
         Height          =   225
         Left            =   300
         TabIndex        =   4
         Top             =   1440
         Width           =   2325
      End
      Begin VB.CheckBox chkComPlete 
         Caption         =   "���ɻ��������־Ϊ�Ѳɼ�"
         Height          =   225
         Left            =   300
         TabIndex        =   3
         Top             =   900
         Width           =   2835
      End
      Begin VB.CheckBox ChkBarCodePrint 
         Caption         =   "���ɻ��������ӡ����"
         Height          =   225
         Left            =   300
         TabIndex        =   2
         Top             =   390
         Width           =   2715
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   120
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleSeupt.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleSeupt.frx":0078
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleSeupt.frx":0612
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleSeupt.frx":0BAC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLabSampleSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCuvette                               '�Թ�
    ѡ��
    ����
    ����
    ��Ӽ�
    ��Ѫ��
    ���
    ��ɫ
End Enum

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer
    Dim Record As ReportRecord
    Dim Column As ReportColumn
    
    rptCuvette.SetImageList ImgList
    With Me.rptCuvette.Columns
        Set Column = .Add(mCuvette.ѡ��, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mCuvette.����, "����", 55, True)
        Set Column = .Add(mCuvette.����, "����", 80, True)
        Set Column = .Add(mCuvette.��Ӽ�, "��Ӽ�", 90, True)
        Set Column = .Add(mCuvette.��Ѫ��, "��Ѫ��", 60, True)
        Set Column = .Add(mCuvette.���, "���", 60, True)
        Set Column = .Add(mCuvette.��ɫ, "", 18, True): Column.Icon = 3
    End With
    
    gstrSql = "Select ����,����,����,��Ӽ�,��Ѫ��,���,��ɫ From ��Ѫ������"
    zlDatabase.OpenRecordset rsTmp, gstrSql, gstrSysName
    Do While Not rsTmp.EOF
        Set Record = Me.rptCuvette.Records.Add
        For intLoop = 0 To Me.rptCuvette.Columns.Count
            Record.AddItem ""
        Next
        
        Record(mCuvette.ѡ��).HasCheckbox = True
        Record(mCuvette.ѡ��).Checked = True
        Record(mCuvette.����).Value = Nvl(rsTmp("����"))
        Record(mCuvette.����).Value = Nvl(rsTmp("����"))
        Record(mCuvette.��Ӽ�).Value = Nvl(rsTmp("��Ӽ�"))
        Record(mCuvette.��Ѫ��).Value = Nvl(rsTmp("��Ѫ��"))
        Record(mCuvette.���).Value = Nvl(rsTmp("���"))
        Record(mCuvette.��ɫ).BackColor = Nvl(rsTmp("��ɫ"))
        
        For intLoop = 0 To Me.rptCuvette.Columns.Count
            Record(intLoop).ForeColor = Nvl(rsTmp("��ɫ"))
        Next
        
        rsTmp.MoveNext
    Loop
    Me.rptCuvette.Populate
End Sub
