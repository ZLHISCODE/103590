VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmWorkFlow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   Icon            =   "frmWorkFlow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame framWorkFlow 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   10575
      Begin VB.Frame frmResultInput 
         Height          =   435
         Left            =   1185
         TabIndex        =   45
         Top             =   6060
         Width           =   5490
         Begin VB.OptionButton optResultInput 
            Caption         =   "�����ӡǰ"
            Height          =   240
            Index           =   2
            Left            =   4050
            TabIndex        =   56
            Top             =   150
            Width           =   1290
         End
         Begin VB.OptionButton optResultInput 
            Caption         =   "���ǩ����"
            Height          =   240
            Index           =   1
            Left            =   2625
            TabIndex        =   55
            Top             =   150
            Width           =   1290
         End
         Begin VB.OptionButton optResultInput 
            Caption         =   "���ǩ����"
            Height          =   240
            Index           =   0
            Left            =   1290
            TabIndex        =   54
            Top             =   150
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.Label lblImageQuality 
            Caption         =   "¼��ʱ����"
            Height          =   180
            Left            =   210
            TabIndex        =   46
            Top             =   165
            Width           =   1035
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1650
         Left            =   120
         TabIndex        =   47
         Top             =   4680
         Width           =   7695
         Begin VB.TextBox txtImageLevel 
            Height          =   270
            Left            =   4290
            TabIndex        =   68
            Text            =   "��,��"
            ToolTipText     =   "��������Ӱ�������ĵǼǣ�����ĸ��ȼ�"
            Top             =   960
            Width           =   1035
         End
         Begin VB.TextBox txtReportLevel 
            Height          =   270
            Left            =   4290
            TabIndex        =   67
            Text            =   "��,��"
            Top             =   360
            Width           =   1035
         End
         Begin VB.CheckBox chkImageLevel 
            Caption         =   "Ӱ�������ȼ�"
            Height          =   180
            Left            =   2880
            TabIndex        =   66
            Top             =   990
            Width           =   1410
         End
         Begin VB.CheckBox chkReportLevel 
            Caption         =   "���������ȼ�"
            Height          =   180
            Left            =   2880
            TabIndex        =   65
            Top             =   375
            Width           =   1410
         End
         Begin VB.CheckBox chkConformDetermine 
            Caption         =   "��������ж�"
            Height          =   180
            Left            =   5700
            TabIndex        =   53
            ToolTipText     =   "�������������ܺͲ˵�"
            Top             =   990
            Width           =   1455
         End
         Begin VB.CheckBox chkCriticalValues 
            Caption         =   "Σ������ж�"
            Height          =   180
            Left            =   5700
            TabIndex        =   52
            ToolTipText     =   "����Σ��������ܺͲ˵�"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame Frame5 
            Height          =   1005
            Left            =   120
            TabIndex        =   48
            Top             =   270
            Width           =   2655
            Begin VB.CheckBox chkDefaultPosi 
               Caption         =   "��Ͻ��Ĭ������"
               Height          =   300
               Left            =   240
               TabIndex        =   51
               ToolTipText     =   "����������ѡ�񴰿ڣ�Ĭ��ѡ�����ԡ�"
               Top             =   300
               Width           =   1815
            End
            Begin VB.CheckBox chkReportAfterResult 
               Caption         =   "���������Ϊ����"
               Height          =   180
               Left            =   240
               TabIndex        =   50
               ToolTipText     =   "��д����ʱ��û��¼����ϣ���Ĭ�ϼ�¼Ϊ���ԡ�"
               Top             =   720
               Width           =   1740
            End
            Begin VB.CheckBox chkIgnorePosi 
               Caption         =   "���Խ����������"
               Height          =   180
               Left            =   240
               TabIndex        =   49
               ToolTipText     =   "����¼�ʹ��������ԡ�"
               Top             =   0
               Width           =   1920
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ƴ����"
         Height          =   2070
         Left            =   8040
         TabIndex        =   36
         Top             =   4280
         Width           =   2415
         Begin VB.OptionButton optCapital 
            Caption         =   "��д"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   42
            ToolTipText     =   "ѡ���ƴ������ʾȫΪ��д��ĸ��"
            Top             =   260
            Width           =   735
         End
         Begin VB.OptionButton optCapital 
            Caption         =   "Сд"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   41
            ToolTipText     =   "ѡ���ƴ������ʾȫΪСд��ĸ��"
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton optCapital 
            Caption         =   "����ĸ��д"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   40
            ToolTipText     =   "ѡ���ƴ��������ĸ��д��"
            Top             =   960
            Width           =   1215
         End
         Begin VB.Frame Frame9 
            Caption         =   "���"
            Height          =   540
            Left            =   120
            TabIndex        =   37
            Top             =   1365
            Width           =   2175
            Begin VB.OptionButton optSplitter 
               Caption         =   "��"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   39
               ToolTipText     =   "ƴ����֮���޼����"
               Top             =   200
               Width           =   495
            End
            Begin VB.OptionButton optSplitter 
               Caption         =   "�ո�"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   38
               ToolTipText     =   "ƴ����֮��ʹ�ÿո�Ϊ�������"
               Top             =   200
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "��������"
         Height          =   2100
         Left            =   120
         TabIndex        =   29
         Top             =   2600
         Width           =   7695
         Begin VB.CheckBox chkAutoInc 
            Caption         =   "�Զ���������"
            Height          =   180
            Left            =   5160
            TabIndex        =   57
            Top             =   375
            Width           =   1635
         End
         Begin VB.Frame Frame11 
            Height          =   1695
            Left            =   5040
            TabIndex        =   72
            Top             =   240
            Width           =   2415
            Begin VB.OptionButton OptBuildcode 
               Caption         =   "��ͬ�������Զ�����"
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   74
               ToolTipText     =   "�����Լ�����Ϊ�������Զ�������"
               Top             =   600
               Value           =   -1  'True
               Width           =   2130
            End
            Begin VB.OptionButton OptBuildcode 
               Caption         =   "���������Զ�����"
               Height          =   210
               Index           =   1
               Left            =   120
               TabIndex        =   73
               ToolTipText     =   "�����Կ���Ϊ�������Զ�������"
               Top             =   1170
               Width           =   1740
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "����һ����"
            Height          =   1770
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2295
            Begin VB.Frame Frame10 
               Height          =   735
               Left            =   375
               TabIndex        =   61
               Top             =   930
               Width           =   1695
               Begin VB.OptionButton OptUnicode 
                  Caption         =   "������ͳһ"
                  Height          =   210
                  Index           =   1
                  Left            =   75
                  TabIndex        =   63
                  ToolTipText     =   "������ͬ�����ּ��Ų��䡣"
                  Top             =   390
                  Width           =   1290
               End
               Begin VB.OptionButton OptUnicode 
                  Caption         =   "��������ͳһ"
                  Height          =   210
                  Index           =   0
                  Left            =   75
                  TabIndex        =   62
                  ToolTipText     =   "��������ͬ�����ּ��Ų��䡣"
                  Top             =   165
                  Width           =   1590
               End
            End
            Begin VB.OptionButton OptCode 
               Caption         =   "���߼��ű��ֲ���"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   35
               ToolTipText     =   "ͬһ�����ߣ�����ʱ���ּ��Ų��䡣"
               Top             =   660
               Width           =   1935
            End
            Begin VB.OptionButton OptCode 
               Caption         =   "ÿ�μ�����¼���"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   34
               ToolTipText     =   "����ʱ�����µļ��š�"
               Top             =   345
               Value           =   -1  'True
               Width           =   1920
            End
         End
         Begin VB.CheckBox chkCanOverWrite 
            Caption         =   "��������ظ�"
            Height          =   300
            Left            =   2760
            TabIndex        =   32
            ToolTipText     =   "����Ǽǲ��˵ļ��ų����ظ���"
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox chkChangeNO 
            Caption         =   "�����ֹ���������"
            Height          =   180
            Left            =   2760
            TabIndex        =   31
            ToolTipText     =   "�������ʵ����Ҫ�ֶ��޸ļ��š�"
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkCheckMaxNo 
            Caption         =   "��ȡʵ��������"
            Height          =   300
            Left            =   2760
            TabIndex        =   30
            ToolTipText     =   "��ʵ��������Ϊ����˳���ţ�����ѡ�����Ե�ǰ���õ�������˳���š�"
            Top             =   1680
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "�ȼ��󱨵���ͼ��ƥ��"
         Height          =   1605
         Left            =   8025
         TabIndex        =   25
         Top             =   2600
         Width           =   2415
         Begin VB.OptionButton optMatch 
            Caption         =   "����/סԺ��"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   28
            ToolTipText     =   "����ʱͨ������/סԺ�ź�ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "����"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   27
            ToolTipText     =   "����ʱͨ�����ź�ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "ҽ��ID"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   26
            ToolTipText     =   "����ʱͨ��ҽ��ID��ͼ����Ϣ����ƥ�䣬������Ӱ��ҽ��վ��"
            Top             =   780
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "��������"
         Height          =   2500
         Left            =   8040
         TabIndex        =   22
         Top             =   0
         Width           =   2415
         Begin VB.CheckBox chkSwitchUser 
            Caption         =   "�����л��û�"
            Height          =   180
            Left            =   240
            TabIndex        =   64
            ToolTipText     =   "�����л��û����ܣ����Խ����û��л���������Ӱ����վ��"
            Top             =   697
            Width           =   1455
         End
         Begin VB.Frame Frame2 
            Height          =   780
            Left            =   120
            TabIndex        =   58
            ToolTipText     =   "ѡ��ɼ�ͼ���ɨ�����뵥��ʹ�õĴ洢�豸��"
            Top             =   1560
            Width           =   2175
            Begin VB.ComboBox cboSaveDevice 
               Height          =   300
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   315
               Width           =   1965
            End
            Begin VB.CheckBox chkPetitionCapture 
               Caption         =   "�������뵥ɨ��"
               Height          =   180
               Left            =   135
               TabIndex        =   59
               ToolTipText     =   "������˺󣬸ü���Զ���ɡ�"
               Top             =   0
               Value           =   1  'Checked
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkUseReferencePatient 
            Caption         =   "���ù�������"
            Height          =   180
            Left            =   240
            TabIndex        =   24
            ToolTipText     =   "֧�ֶ����������ͬһ��������Ϣ��"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkChangeUser 
            Caption         =   "���ý����û�"
            Height          =   180
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "������û����ܣ����Խ������ҽ���ͱ���ҽ����������Ӱ��ɼ�վ��"
            Top             =   315
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����������"
         Height          =   2505
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   7695
         Begin VB.CheckBox chkReagent 
            Caption         =   "��ʾ��Ӱ��"
            Height          =   180
            Left            =   2280
            TabIndex        =   78
            ToolTipText     =   "�ڵǼǱ���������ʾ��Ӱ��һ�������վ����ʾ"
            Top             =   2200
            Width           =   1935
         End
         Begin VB.CheckBox chkAddons 
            Caption         =   "��ʾ��������"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            ToolTipText     =   "�ڵǼǱ���������ʾ��������һ��"
            Top             =   2160
            Width           =   1935
         End
         Begin VB.CheckBox chkAutoSendWorkList 
            Caption         =   "����ʱ�Զ�����WorkList"
            Height          =   180
            Left            =   2280
            TabIndex        =   76
            Top             =   1920
            Value           =   1  'Checked
            Width           =   2532
         End
         Begin VB.CheckBox chkCompletePrint 
            Caption         =   "�����ֱ�Ӵ�ӡ"
            Height          =   180
            Left            =   240
            TabIndex        =   75
            ToolTipText     =   "����ǩ����ֱ�Ӵ�ӡ���棬���������°汨���ĵ��༭����"
            Top             =   1920
            Width           =   1680
         End
         Begin VB.TextBox txtViewHistoryImageDays 
            Height          =   270
            Left            =   6960
            MaxLength       =   2
            TabIndex        =   70
            Text            =   "1"
            Top             =   1878
            Width           =   465
         End
         Begin VB.CheckBox chkCanViewImage 
            Caption         =   "��ͼ��ҽ��վ���ɹ�Ƭ"
            Height          =   180
            Left            =   2280
            TabIndex        =   69
            ToolTipText     =   "�ɼ�ͼ�����û�м����ɵ�����£�ҽ��վҲ�ɽ��й�Ƭ��"
            Top             =   1611
            Width           =   2160
         End
         Begin VB.TextBox txtRefreshInterval 
            Height          =   270
            Left            =   4080
            TabIndex        =   44
            Text            =   "0"
            Top             =   960
            Width           =   390
         End
         Begin VB.TextBox TxtLike 
            Enabled         =   0   'False
            Height          =   270
            Left            =   4200
            TabIndex        =   43
            ToolTipText     =   "0������ʱ������,ģ���������в���"
            Top             =   600
            Width           =   270
         End
         Begin VB.CheckBox ChkFinishCommit 
            Caption         =   "�ޱ�����ɺ�ֱ�����"
            Height          =   180
            Left            =   2280
            TabIndex        =   20
            ToolTipText     =   "����ޱ�����ɺ󣬸ü���Զ���ɡ�"
            Top             =   1302
            Width           =   2160
         End
         Begin VB.CheckBox chkPrintCommit 
            Caption         =   "��ӡ��ֱ�����"
            Height          =   180
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "��ӡ����󣬸ü���Զ���ɡ�"
            Top             =   684
            Width           =   1815
         End
         Begin VB.CheckBox ChkCompleteCommit 
            Caption         =   "��˺�ֱ�����"
            Height          =   180
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "������˺󣬸ü���Զ���ɡ�"
            Top             =   993
            Width           =   1935
         End
         Begin VB.CheckBox chkSample 
            Caption         =   "����ǼǺ�ֱ�ӱ���"
            Height          =   180
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "�Ǽ��뱨��ͬʱ���С�"
            Top             =   1611
            Width           =   1935
         End
         Begin VB.TextBox TxtĬ������ 
            Height          =   270
            Left            =   6720
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "2"
            Top             =   1491
            Width           =   705
         End
         Begin VB.CheckBox chkReportAfterImging 
            Caption         =   "��ͼ�����д����"
            Height          =   180
            Left            =   240
            TabIndex        =   15
            ToolTipText     =   "����ɼ�ͼ�����ܱ�дӰ�񱨸档"
            Top             =   375
            Width           =   2040
         End
         Begin VB.CheckBox chkPrintNeedComplete 
            Caption         =   "ƽ��������˲��ܴ򱨸�"
            Height          =   180
            Left            =   5040
            TabIndex        =   14
            ToolTipText     =   "ƽ������뾭����˺���ܴ�ӡ���档"
            Top             =   762
            Width           =   2505
         End
         Begin VB.CheckBox chkTechReportSame 
            Caption         =   "ֻ����д�Լ����ı���"
            Height          =   180
            Left            =   5040
            TabIndex        =   13
            ToolTipText     =   "ֻ���Լ��ɼ�ͼ��ļ�飬������д���档"
            Top             =   375
            Width           =   2295
         End
         Begin VB.CheckBox chkWriteCapDoctor 
            Caption         =   "�ɼ�ͼ����Ϊ��鼼ʦ"
            Height          =   180
            Left            =   5040
            TabIndex        =   12
            ToolTipText     =   "�ɼ�ͼ��֮���Զ�����ǰ�û���¼�ɼ�鼼ʦ��"
            Top             =   1149
            Width           =   2400
         End
         Begin VB.CheckBox chkLocalizerBackward 
            Caption         =   "��λƬ����"
            Height          =   180
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "����λƬ�ŵ����һ��������ʾ��"
            Top             =   1302
            Width           =   1320
         End
         Begin VB.CheckBox chkRefreshInterval 
            Caption         =   "�����Զ�ˢ�¼��      ��"
            Height          =   180
            Left            =   2280
            TabIndex        =   10
            ToolTipText     =   "���˼���б����N���Զ�ˢ�¡�"
            Top             =   993
            Width           =   2500
         End
         Begin VB.CheckBox ChkLike 
            Caption         =   "�Ǽ�ʱ����ģ������    ��"
            Height          =   195
            Left            =   2280
            TabIndex        =   9
            ToolTipText     =   "�Ǽ�ʱ֧�ֶ���������ģ�����ң����Բ��ҵ�N���ڵ���Ϣ��"
            Top             =   684
            Width           =   2500
         End
         Begin VB.CheckBox chkAllPatientIsOutside 
            Caption         =   "���еǼǲ��˱��Ϊ����"
            Height          =   180
            Left            =   2280
            TabIndex        =   8
            ToolTipText     =   "���ڸù���վ�еǼǵĲ��˾����Ϊ�������ˡ�"
            Top             =   375
            Width           =   2295
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Զ�����ʷͼ������"
            Height          =   180
            Left            =   5040
            TabIndex        =   71
            ToolTipText     =   "�����ǰ���û��ͼ�����Զ���ָ��ʱ����ڵ���ʷͼ��"
            Top             =   1923
            Width           =   1800
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ĭ�ϼ�¼��ѯ����"
            Height          =   180
            Left            =   5040
            TabIndex        =   21
            ToolTipText     =   "����б���Ĭ����ʾ��Ӧ�����ڵļ���¼��"
            Top             =   1536
            Width           =   1440
         End
      End
   End
   Begin VB.ComboBox cmbDept 
      Height          =   300
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   75
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9465
      TabIndex        =   3
      Top             =   7640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8265
      TabIndex        =   2
      Top             =   7640
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7650
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl TabWindow 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      _Version        =   589884
      _ExtentX        =   18865
      _ExtentY        =   12303
      _StockProps     =   64
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Ӱ�����"
      Height          =   180
      Left            =   165
      TabIndex        =   5
      Top             =   135
      Width           =   735
   End
End
Attribute VB_Name = "frmWorkFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String         '��ģ���Ȩ��
Public mlng����ID As Long 'IN:��ǰִ�п���ID
Private mlngCur����ID As Long       '��ǰ����ID
Private mstrCur���� As String      '��ǰ���� ����-����
Private mstrCanUse���� As String    '��ǰ���ÿ���  ID_����-����
Private mobjfrmTabPass As New FrmReqInput     '��꾭������
Private mobjfrmEnableCtr As New FrmReqInput  '�������������
Private mobjFrmReportSetup As New frmReportSetup '��������
Private mobjFrmStudyListCfg As New frmStudyListCfg '����б�����
Private mobjfrmTechnicGroupCfg As New frmTechnicQueueCfg 'ҽ��ִ�м��������


Private Sub chkAutoInc_Click()
On Error Resume Next
    If chkAutoInc.value = 0 Then
        OptBuildcode(0).Enabled = False
        OptBuildcode(1).Enabled = False
        
        chkChangeNO.value = 1
        chkChangeNO.Enabled = False
        
        chkCheckMaxNo.value = 0
        chkCheckMaxNo.Enabled = False
    Else
        OptBuildcode(0).Enabled = True
        OptBuildcode(1).Enabled = True
        
        chkChangeNO.Enabled = True
        chkCheckMaxNo.Enabled = True
    End If
err.Clear
End Sub


Private Sub chkImageLevel_Click()
    txtImageLevel.Enabled = chkImageLevel.value = 1
End Sub

Private Sub ChkLike_Click()
    TxtLike.Enabled = IIf(ChkLike.value, True, False)
End Sub

Private Sub chkPetitionCapture_Click()
    cboSaveDevice.Enabled = IIf(chkPetitionCapture.value, True, False)
End Sub

Private Sub chkRefreshInterval_Click()
    txtRefreshInterval.Enabled = IIf(chkRefreshInterval.value, True, False)
End Sub

Private Sub chkReportAfterResult_Click()
    If chkReportAfterResult.value = vbChecked Then
        chkIgnorePosi.Enabled = False
        chkIgnorePosi.value = vbUnchecked
    Else
        chkIgnorePosi.Enabled = True
    End If
End Sub


Private Sub chkReportLevel_Click()
    txtReportLevel.Enabled = chkReportLevel.value = 1
End Sub

Private Sub cmbDept_Click()
    mlng����ID = cmbDept.ItemData(cmbDept.ListIndex)
    If TabWindow.ItemCount = IIf(InStr(GetPrivFunc(glngSys, 1160), "����") > 0, 7, 6) Then '�ж�tab������Ŀ����Ϊ��ȷ����װ����tab֮��Ŵ������е����
        'ˢ�¹������̲�������
        Call frmWorkFlowRefresh
        'ˢ��ִ�м����
        Call frmTechRoomRefresh
        'ˢ���������ý���
        Call frmReqInputRefresh(0)
        '���������
        Call frmReqInputRefresh(1)
        'ˢ�±�������
        Call frmReportRefresh
        'ˢ����ɫ����
        Call frmStudyListCfgRefresh
        'ˢ���Ŷӽк�����
        RefreshTechnicRoomGroupCfg
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub CmdOK_Click()

    Dim intTxtLen As Integer
    
    If txtImageLevel.Enabled Then
        '������״̬�µ� �����滻��Ӣ��״̬
        txtImageLevel.Text = Replace(txtImageLevel.Text, "��", ",")
        
        intTxtLen = Len(txtImageLevel.Text) - Len(Replace(txtImageLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBoxD Me, "Ӱ��ȼ�����Ϊ2�֣����Ϊ4�֣���������д��", vbOKOnly, "��ʾ��Ϣ"
            txtImageLevel.Text = Nvl(GetDeptPara(mlng����ID, "Ӱ�������ȼ�", "��,��"))
            txtImageLevel.SetFocus
            Exit Sub
        End If
    End If
    
    
    If txtReportLevel.Enabled Then
        '������״̬�µ� �����滻��Ӣ��״̬
        txtReportLevel.Text = Replace(txtReportLevel.Text, "��", ",")
        
        intTxtLen = Len(txtReportLevel.Text) - Len(Replace(txtReportLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBoxD Me, "����ȼ�����Ϊ2�֣����Ϊ4�֣���������д��", vbOKOnly, "��ʾ��Ϣ"
            txtReportLevel.Text = Nvl(GetDeptPara(mlng����ID, "���������ȼ�", "��,��"))
            txtReportLevel.SetFocus
            Exit Sub
        End If
    End If
    

    Call SaveWorkFlow
    Call mobjfrmTabPass.zlSave
    Call mobjfrmEnableCtr.zlSave
    Call mobjFrmReportSetup.zlSave
    Call mobjFrmStudyListCfg.zlSave
    Call mobjfrmTechnicGroupCfg.zlSave
    
    Unload Me
End Sub

Private Sub Form_Load()
    '��ʼ��ģ�鼶����
    mstrPrivs = gstrPrivs
    mlng����ID = 0
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""
    
    mobjfrmTabPass.mintType = 0
    mobjfrmEnableCtr.mintType = 1
    
    'û�ж�Ӧ�Ŀ��ң����˳�
    If InitDepts = False Then
        Unload Me
        Exit Sub
    End If
    
    'װ���Ӵ���
    Call InitFaceScheme
    
    '��ʼ���Ӵ���
    'ˢ�¹������̲�������
    Call frmWorkFlowRefresh
    'ˢ��ִ�м����
    Call frmTechRoomRefresh
    'ˢ���������ý���
    Call frmReqInputRefresh(0)
    '���������
    Call frmReqInputRefresh(1)
    'ˢ�±�������
    Call frmReportRefresh
    'ˢ�¼���б�����
    Call frmStudyListCfgRefresh
    'ˢ���Ŷӽк�����
    Call RefreshTechnicRoomGroupCfg
End Sub

Private Sub Form_Resize()
    TabWindow.Left = 1
    TabWindow.Top = 480
    TabWindow.Width = Me.ScaleWidth
    TabWindow.Height = Me.ScaleHeight - 480
End Sub

Private Sub InitFaceScheme()
    Dim Item As TabControlItem
    
    mobjfrmTabPass.mlngDeptId = mlng����ID
    mobjfrmEnableCtr.mlngDeptId = mlng����ID
    frmTechnicRoom.mlngdept = mlng����ID
    
    With TabWindow
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 1, "����������", framWorkFlow.hWnd, 0
        .InsertItem 2, "ִ�м�����", frmTechnicRoom.hWnd, 0
        
        '��1160��Ȩ��ʱ�������������
        If InStr(GetPrivFunc(glngSys, 1160), "����") > 0 Then
            .InsertItem 3, "�����Ŷ�����", mobjfrmTechnicGroupCfg.hWnd, 0
        End If
        
        .InsertItem 4, "���뾭������", mobjfrmTabPass.hWnd, 0
        .InsertItem 5, "�����¼����", mobjfrmEnableCtr.hWnd, 0
        .InsertItem 6, "PACS��������", mobjFrmReportSetup.hWnd, 0
        .InsertItem 7, "����б�����", mobjFrmStudyListCfg.hWnd, 0
        
        framWorkFlow.BorderStyle = 0
        .Item(0).Selected = True
    End With
    framWorkFlow.Width = Me.ScaleWidth
    framWorkFlow.Height = Me.ScaleHeight
    frmTechnicRoom.Width = Me.ScaleWidth
    frmTechnicRoom.Height = Me.ScaleHeight
    mobjfrmTabPass.Width = Me.ScaleWidth
    mobjfrmTabPass.Height = Me.ScaleHeight
    mobjfrmEnableCtr.Width = Me.ScaleWidth
    mobjfrmEnableCtr.Height = Me.ScaleHeight
    mobjFrmReportSetup.Width = Me.ScaleWidth
    mobjFrmReportSetup.Height = Me.ScaleHeight
    mobjFrmStudyListCfg.Width = Me.ScaleWidth
    mobjFrmStudyListCfg.Height = Me.ScaleHeight
    mobjfrmTechnicGroupCfg.Width = Me.ScaleWidth
    mobjfrmTechnicGroupCfg.Height = Me.ScaleHeight
End Sub

Private Sub frmTechRoomRefresh()
    'ˢ��ִ�м�ҳ��
    frmTechnicRoom.mlngdept = mlng����ID
    frmTechnicRoom.zlRoomRef
End Sub

Private Sub frmReqInputRefresh(ByVal intType As Integer)
    If intType = 0 Then
        mobjfrmTabPass.mlngDeptId = mlng����ID
        mobjfrmTabPass.zlRefresh
    ElseIf intType = 1 Then
        mobjfrmEnableCtr.mlngDeptId = mlng����ID
        mobjfrmEnableCtr.zlRefresh
    End If
End Sub

Private Sub frmStudyListCfgRefresh()
    Call mobjFrmStudyListCfg.zlRefresh(mlng����ID)
End Sub


Private Sub RefreshTechnicRoomGroupCfg()
'ˢ��ִ�м��������
    Call mobjfrmTechnicGroupCfg.zlRefresh(mlng����ID)
End Sub


Private Sub frmWorkFlowRefresh()
    Dim rsTemp As ADODB.Recordset
    Dim lngHintType As Long
        
    '��ʼ��Ĭ��ֵ,Ӧ����һ��ͳһ�ĵط�����Ĭ��ֵ������������ʾ�����ն�ȡ
    chkIgnorePosi.value = 0     '���Խ��������
    chkReportAfterResult.value = 0 '��Ӱ�����Ϊ����
    ChkFinishCommit.value = 0   '�ޱ�����ɺ�ֱ�����
    chkReportAfterImging.value = 0  '��ͼ�񲻿ɱ༭����
    chkLocalizerBackward.value = 0  '��λƬ����
    chkChangeUser.value = 0         '�������û�
    chkSwitchUser.value = 0         '�����л��û�
    chkTechReportSame.value = 0     'ֻ����д�Լ����ı���
    chkWriteCapDoctor.value = 0     '�ɼ�ͼ����Ϊ��鼼ʦ
    ChkCompleteCommit.value = 0     '��˺�ֱ�����
    optMatch(0).value = True        'ƥ�����ݿ���Ŀ
    
    ChkLike.value = 0               '���õǼ�ʱ����ģ������
    TxtLike.Text = 0                '�Ǽ�ʱ����ģ����������
    TxtĬ������.Text = 2            'Ĭ�Ϲ�������
    txtViewHistoryImageDays.Text = 1 'Ĭ���Զ�����ʷͼ������
    chkRefreshInterval.value = 0    '���ò����б��Զ�ˢ��
    txtRefreshInterval.Text = 0     'Ĭ�ϲ����б��Զ�ˢ�¼��Ϊ0�룬��ˢ��
    cboSaveDevice.Clear                 '�洢�豸
    chkPrintCommit.value = 0        '��ӡ��ֱ�����
    chkCompletePrint.value = 0      '�����ֱ�Ӵ�ӡ
    chkUseReferencePatient.value = 0  'Ĭ�ϲ����ù�������
    optCapital(0).value = True      'Ĭ��ƴ��ʹ�ô�д
    optCapital(1).value = True      'Ĭ��ƴ������ÿո�
    chkCheckMaxNo.value = 1         'Ĭ����ȡʵ��������
    chkDefaultPosi.value = 0        '��Ͻ��Ĭ������Ϊδ��ѡ
    chkConformDetermine.value = 1       '��������ж�Ĭ��Ϊѡ��
    chkCriticalValues.value = 1      'Σ������ж�Ĭ��Ϊѡ��
    txtImageLevel.Text = "��,��"     'Ĭ��Ӱ�������ȼ�
    txtReportLevel.Text = "��,��"    'Ĭ�ϱ��������ȼ�
    chkPetitionCapture.value = 1     'Ĭ�Ϲ�ѡ�������뵥ɨ��
    chkAddons.value = 1              '�ڵǼǴ�����ʾ��������
    chkReagent.value = 1             '�ڵǼǴ�����ʾ��Ӱ��
    
    On Error GoTo err
    
    lngHintType = Val(GetDeptPara(mlng����ID, "��Ͻ����ʾ����", 0))
    optResultInput(lngHintType).value = True
    
    chkIgnorePosi.value = Val(GetDeptPara(mlng����ID, "���Խ��������", 0)) '��һ��ʹ��ʱ��Ҫ���¶�ȡ
    chkDefaultPosi.value = Val(GetDeptPara(mlng����ID, "��Ͻ��Ĭ������", 0))  '��ȡĬ�����Բ���
    chkReportAfterResult.value = Val(GetDeptPara(mlng����ID, "��Ӱ�����Ϊ����", 0))
    
    chkCriticalValues.value = Val(GetDeptPara(mlng����ID, "Σ������ж�", 0))    '��ȡΣ������ж�
    chkConformDetermine.value = Val(GetDeptPara(mlng����ID, "��������ж�", 0))    '��ȡ��������ж�
    
    chkImageLevel.value = Val(GetDeptPara(mlng����ID, "Ӱ�������ж�", 0))   '��ȡӰ�������ж�
    txtImageLevel.Text = Nvl(GetDeptPara(mlng����ID, "Ӱ�������ȼ�", "��,��"))  '��ȡӰ�������ȼ�
    txtImageLevel.Enabled = chkImageLevel.value = 1
    
    chkReportLevel.value = Val(GetDeptPara(mlng����ID, "���������ж�", 0)) '��ȡ���������ж�
    txtReportLevel.Text = Nvl(GetDeptPara(mlng����ID, "���������ȼ�", "��,��"))  '��ȡ���������ȼ�
    txtReportLevel.Enabled = chkReportLevel.value = 1
    
    chkPetitionCapture.value = Val(GetDeptPara(mlng����ID, "�������뵥ɨ��", 1))    '��ȡ�������뵥ɨ�����

    ChkFinishCommit.value = Val(GetDeptPara(mlng����ID, "�ޱ�����ɺ�ֱ�����", 0))
    chkCanViewImage.value = Val(GetDeptPara(mlng����ID, "��ͼ��ҽ��վ���ɹ�Ƭ", 0))
    chkReportAfterImging.value = Val(GetDeptPara(mlng����ID, "��ͼ�����д����", 0))
    chkCanOverWrite.value = Val(GetDeptPara(mlng����ID, "��������ظ�", 0))
    chkCheckMaxNo.value = Val(GetDeptPara(mlng����ID, "��ȡʵ��������", 1))
    chkChangeNO.value = Val(GetDeptPara(mlng����ID, "�ֹ���������", 0))
    chkLocalizerBackward.value = Val(GetDeptPara(mlng����ID, "��λƬ����", 0))
    chkChangeUser.value = Val(GetDeptPara(mlng����ID, "�������û�", 0))
    chkSwitchUser.value = Val(GetDeptPara(mlng����ID, "�����л��û�", 0))
    chkTechReportSame.value = Val(GetDeptPara(mlng����ID, "ֻ����д�Լ����ı���", 0))
    chkWriteCapDoctor.value = Val(GetDeptPara(mlng����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", 0))
    ChkCompleteCommit.value = Val(GetDeptPara(mlng����ID, "��˺�ֱ�����", 0))
    chkPrintCommit.value = Val(GetDeptPara(mlng����ID, "��ӡ��ֱ�����", 0))
    chkCompletePrint.value = Val(GetDeptPara(mlng����ID, "�����ֱ�Ӵ�ӡ", 0))
    
    TxtLike.Text = Val(GetDeptPara(mlng����ID, "�Ǽ�ʱ����ģ����������", 0))
    chkSample.value = Val(GetDeptPara(mlng����ID, "�ǼǺ�ֱ�Ӽ��", 0))
    ChkLike.value = IIf(Val(TxtLike.Text) <> 0, 1, 0)
    chkAllPatientIsOutside.value = Val(GetDeptPara(mlng����ID, "���еǼǲ��˱��Ϊ����", 0))
    
    TxtĬ������.Text = Val(GetDeptPara(mlng����ID, "Ĭ�Ϲ�������", 2))
    
    If Val(TxtĬ������.Text) > 15 Or Val(TxtĬ������.Text) <= 0 Then
        TxtĬ������.Text = 2
    End If
    
    txtViewHistoryImageDays.Text = Val(GetDeptPara(mlng����ID, "�Զ�����ʷͼ������", 1))
    If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
        txtViewHistoryImageDays.Text = 1
    End If
    
    txtRefreshInterval.Text = Val(GetDeptPara(mlng����ID, "�Զ�ˢ�¼��", 0))
    chkRefreshInterval.value = IIf(Val(txtRefreshInterval.Text) <> 0, 1, 0)
    optMatch(Val(GetDeptPara(mlng����ID, "ƥ�����ݿ���Ŀ", 0))).value = True
    
    OptBuildcode(Val(GetDeptPara(mlng����ID, "�������ɷ�ʽ", 0))).value = True
    chkAutoInc.value = Val(GetDeptPara(mlng����ID, "�Զ���������"))
    chkAutoSendWorkList.value = Val(GetDeptPara(mlng����ID, "����ʱ�Զ�����WorkList", "1"))
    chkAddons.value = Val(GetDeptPara(mlng����ID, "��ʾ��������", "1"))
    chkReagent.value = Val(GetDeptPara(mlng����ID, "��ʾ��Ӱ��", "1"))
    
    If chkAutoInc.value = 0 Then
        OptBuildcode(0).Enabled = False
        OptBuildcode(1).Enabled = False
        
        chkChangeNO.value = 1
        chkChangeNO.Enabled = False
        
        chkCheckMaxNo.value = 0
        chkCheckMaxNo.Enabled = False
    Else
        OptBuildcode(0).Enabled = True
        OptBuildcode(1).Enabled = True
        
        chkChangeNO.Enabled = True
        chkCheckMaxNo.Enabled = True
    End If
    
    OptCode(Val(GetDeptPara(mlng����ID, "���߼��ű��ֲ���", 0))).value = True
    If OptCode(1).value = True Then
        OptUnicode(0).Enabled = True
        OptUnicode(1).Enabled = True
        OptUnicode(Val(GetDeptPara(mlng����ID, "���ű��ֲ������", 0))).value = True
    Else
        OptUnicode(0).Enabled = False: OptUnicode(0).value = False
        OptUnicode(1).Enabled = False: OptUnicode(1).value = False
    End If
    
    chkUseReferencePatient.value = Val(GetDeptPara(mlng����ID, "������������", 0))
    chkPrintNeedComplete.value = Val(GetDeptPara(mlng����ID, "ƽ������˲��ܴ򱨸�", 0))
    
    'ƴ��������
    optCapital(Val(GetDeptPara(mlng����ID, "ƴ������Сд", 0))).value = True
    optSplitter(Val(GetDeptPara(mlng����ID, "ƴ�����ָ���", 0))).value = True
    
    
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and NVL(״̬,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTemp.EOF Then
        MsgBoxD Me, "δ����Ӱ��洢�豸���뵽Ӱ���豸Ŀ¼�����ã�", vbInformation, gstrSysName
        Exit Sub
    Else
        cboSaveDevice.AddItem ""
        
        Do While Not rsTemp.EOF
            cboSaveDevice.AddItem rsTemp!�豸�� & "-" & Nvl(rsTemp!�豸��)
            
            If GetDeptPara(mlng����ID, "�洢�豸��", "") = rsTemp!�豸�� Then
                cboSaveDevice.ListIndex = cboSaveDevice.NewIndex
            End If
            
            rsTemp.MoveNext
        Loop
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub SaveWorkFlow()
    On Error GoTo errHand

    SetDeptPara mlng����ID, "�������뵥ɨ��", chkPetitionCapture.value        '�������뵥ɨ�� ��������
    
    SetDeptPara mlng����ID, "��������ж�", chkConformDetermine.value         '��������ж� ��������
    SetDeptPara mlng����ID, "Σ������ж�", chkCriticalValues.value           'Σ������ж� ��������
    
    SetDeptPara mlng����ID, "���Խ��������", chkIgnorePosi.value
    SetDeptPara mlng����ID, "��Ӱ�����Ϊ����", chkReportAfterResult.value
    SetDeptPara mlng����ID, "��Ͻ��Ĭ������", chkDefaultPosi.value   '��Ͻ��Ĭ������ ��������
    
    SetDeptPara mlng����ID, "Ӱ�������ж�", chkImageLevel.value           'Ӱ�������ж� ��������
    SetDeptPara mlng����ID, "Ӱ�������ȼ�", txtImageLevel.Text            'ͼ�������ȼ� ��������
    SetDeptPara mlng����ID, "���������ж�", chkReportLevel.value          '���������ж� ��������
    SetDeptPara mlng����ID, "���������ȼ�", txtReportLevel.Text           '���������ȼ� ��������
    
    SetDeptPara mlng����ID, "��Ͻ����ʾ����", IIf(optResultInput(0).value = True, 0, IIf(optResultInput(1).value = True, 1, 2))
    
    SetDeptPara mlng����ID, "�ޱ�����ɺ�ֱ�����", ChkFinishCommit.value
    SetDeptPara mlng����ID, "��ͼ��ҽ��վ���ɹ�Ƭ", chkCanViewImage.value     '��ͼ��ҽ��վ���ɹ�Ƭ
    SetDeptPara mlng����ID, "��ͼ�����д����", chkReportAfterImging.value
    SetDeptPara mlng����ID, "���߼��ű��ֲ���", IIf(OptCode(1).value, 1, 0)
    SetDeptPara mlng����ID, "���ű��ֲ������", IIf(OptUnicode(1).value, 1, 0)
    SetDeptPara mlng����ID, "�������ɷ�ʽ", IIf(OptBuildcode(1).value, 1, 0)
    SetDeptPara mlng����ID, "�Զ���������", chkAutoInc.value
    SetDeptPara mlng����ID, "�ֹ���������", chkChangeNO.value
    SetDeptPara mlng����ID, "��������ظ�", chkCanOverWrite.value
    SetDeptPara mlng����ID, "��ȡʵ��������", chkCheckMaxNo.value
    SetDeptPara mlng����ID, "��λƬ����", chkLocalizerBackward.value
    SetDeptPara mlng����ID, "�������û�", chkChangeUser.value
    SetDeptPara mlng����ID, "�����л��û�", chkSwitchUser.value
    SetDeptPara mlng����ID, "ֻ����д�Լ����ı���", chkTechReportSame.value
    SetDeptPara mlng����ID, "�ɼ�ͼ����Ϊ��鼼ʦ", chkWriteCapDoctor.value
    SetDeptPara mlng����ID, "��˺�ֱ�����", ChkCompleteCommit.value
    SetDeptPara mlng����ID, "��ӡ��ֱ�����", chkPrintCommit.value
    SetDeptPara mlng����ID, "�����ֱ�Ӵ�ӡ", chkCompletePrint.value
    SetDeptPara mlng����ID, "�ǼǺ�ֱ�Ӽ��", chkSample.value
    SetDeptPara mlng����ID, "ƥ�����ݿ���Ŀ", IIf(optMatch(0).value, 0, IIf(optMatch(1), 1, 2))
    
    SetDeptPara mlng����ID, "�Ǽ�ʱ����ģ����������", IIf(ChkLike.value = 1, Abs(Val(TxtLike.Text)), 0)
    SetDeptPara mlng����ID, "���еǼǲ��˱��Ϊ����", chkAllPatientIsOutside
    
    If Val(TxtĬ������.Text) > 15 Or Val(TxtĬ������.Text) <= 0 Then
        TxtĬ������.Text = 2
    End If
    SetDeptPara mlng����ID, "Ĭ�Ϲ�������", Val(TxtĬ������.Text)
    
    If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
        txtViewHistoryImageDays.Text = 1
    End If
    SetDeptPara mlng����ID, "�Զ�����ʷͼ������", Val(txtViewHistoryImageDays.Text)
    
    
    SetDeptPara mlng����ID, "������������", chkUseReferencePatient.value
    SetDeptPara mlng����ID, "ƽ������˲��ܴ򱨸�", chkPrintNeedComplete.value
    
    SetDeptPara mlng����ID, "ƴ������Сд", IIf(optCapital(0).value, 0, IIf(optCapital(1), 1, 2))
    SetDeptPara mlng����ID, "ƴ�����ָ���", IIf(optSplitter(0).value, 0, 1)
    
    If cboSaveDevice.Text <> "" Then
        SetDeptPara mlng����ID, "�洢�豸��", Split(cboSaveDevice.Text, "-")(0)
    Else
        SetDeptPara mlng����ID, "�洢�豸��", ""
    End If
    
    If Abs(Val(txtRefreshInterval.Text)) = 0 Or Abs(Val(txtRefreshInterval.Text)) > 65 Then
        txtRefreshInterval.Text = 10
    End If
    SetDeptPara mlng����ID, "�Զ�ˢ�¼��", IIf(chkRefreshInterval.value = 1, Abs(Val(txtRefreshInterval.Text)), 0)
    SetDeptPara mlng����ID, "����ʱ�Զ�����WorkList", chkAutoSendWorkList.value
    SetDeptPara mlng����ID, "��ʾ��������", chkAddons.value
    SetDeptPara mlng����ID, "��ʾ��Ӱ��", chkReagent.value
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Private Function InitDepts() As Boolean
'���ܣ���ʼ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String, str��Դ As String
    Dim strDepartment() As String
    Dim intCurDept As Integer
    
    On Error GoTo errH
    
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.�������� IN('���')  Order by A.����"
    Else
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.�������� IN('���')  Order by A.����"
    End If
     
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
            rsTmp.MoveNext
        Loop
        
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���,ȡ���������ҵ�һ��
            rsTmp.MoveNext
        Loop
        mstrCanUse���� = Mid(mstrCanUse����, 2)
        If InStr(mstrPrivs, "���п���") > 0 And mlngCur����ID = 0 Then
            mlngCur����ID = Split(Split(mstrCanUse����, "|")(0), "_")(0)
            mstrCur���� = Split(Split(mstrCanUse����, "|")(0), "_")(1)
        End If
        
        If mlngCur����ID = 0 And InStr(mstrPrivs, "���п���") <= 0 Then 'û�����п��Ҳ���Ȩ��,���Ҳ����߿��Ҳ����ڼ�������
            MsgBoxD Me, "û�з�������������,����ʹ��ҽ������վ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        '���cmbDept
        cmbDept.Clear
        intCurDept = -1
        strDepartment = Split(mstrCanUse����, "|")
        For i = 0 To UBound(strDepartment)
            cmbDept.AddItem Split(strDepartment(i), "_")(1)
            cmbDept.ItemData(cmbDept.ListCount - 1) = Split(strDepartment(i), "_")(0)
            If Split(strDepartment(i), "_")(0) = mlngCur����ID Then
                intCurDept = i
            End If
        Next i
        If intCurDept <> -1 Then
            cmbDept.ListIndex = intCurDept
        Else
            cmbDept.ListIndex = 0
        End If
        mlng����ID = cmbDept.ItemData(cmbDept.ListIndex)
        InitDepts = True
    End If
    
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTechnicRoom
    Unload mobjfrmEnableCtr
    Unload mobjfrmTabPass
    Unload mobjFrmReportSetup
    Unload mobjFrmStudyListCfg
    Unload mobjfrmTechnicGroupCfg
End Sub


Private Sub OptCode_Click(Index As Integer)
    OptUnicode(0).Enabled = Index = 1
    OptUnicode(1).Enabled = Index = 1
End Sub

Private Sub frmReportRefresh()
    mobjFrmReportSetup.zlRefresh (mlng����ID)
End Sub


Private Sub txtViewHistoryImageDays_Change()
    If Val(txtViewHistoryImageDays.Text) > 15 Or Val(txtViewHistoryImageDays.Text) <= 0 Then
        MsgBoxD Me, "�Զ�����ʷͼ����������Ϊ1�죬���Ϊ15�죬��������д��", vbOKOnly, "��ʾ��Ϣ"
    End If
End Sub

Private Sub TxtĬ������_Change()
    If Val(TxtĬ������.Text) > 15 Or Val(TxtĬ������.Text) <= 0 Then
        MsgBoxD Me, "Ĭ����������Ϊ1�죬���Ϊ15�죬��������д��", vbOKOnly, "��ʾ��Ϣ"
    End If
End Sub
