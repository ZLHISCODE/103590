VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDataMove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ת�ƹ���"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "frmDataMove.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5000
      Index           =   0
      Left            =   0
      TabIndex        =   30
      Tag             =   "����ת��"
      Top             =   885
      Width           =   9405
      Begin VB.TextBox txtDatePre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F4E4&
         ForeColor       =   &H00000000&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   7740
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   86
         Text            =   "2011-01-01"
         Top             =   2775
         Width           =   1020
      End
      Begin VB.CommandButton cmdDateThis 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   8805
         Picture         =   "frmDataMove.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2400
         Width           =   280
      End
      Begin VB.TextBox txtDateThis 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   280
         IMEMode         =   3  'DISABLE
         Left            =   7740
         MaxLength       =   10
         TabIndex        =   84
         Text            =   "2012-01-01"
         Top             =   2400
         Width           =   1020
      End
      Begin MSComCtl2.MonthView monSel 
         Height          =   2460
         Left            =   3000
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   4339
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollRate      =   1
         StartOfWeek     =   39583745
         TitleBackColor  =   8421504
         TitleForeColor  =   16777215
         CurrentDate     =   38003
         MaxDate         =   73415
         MinDate         =   -18260
      End
      Begin VB.TextBox txtDateLast 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   280
         IMEMode         =   3  'DISABLE
         Left            =   7740
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "2015-01-01"
         Top             =   1995
         Width           =   1020
      End
      Begin VB.Frame framode 
         Caption         =   "ת������"
         Height          =   4425
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   6135
         Begin VB.PictureBox picBakspace 
            BorderStyle     =   0  'None
            Height          =   1425
            Left            =   240
            ScaleHeight     =   1425
            ScaleWidth      =   3210
            TabIndex        =   87
            Top             =   2880
            Width           =   3210
            Begin VB.ComboBox cboBakspace 
               Height          =   300
               Index           =   0
               Left            =   1180
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   300
               Width           =   1920
            End
            Begin VB.ComboBox cboBakspace 
               Height          =   300
               Index           =   1
               Left            =   1180
               Style           =   2  'Dropdown List
               TabIndex        =   89
               Top             =   660
               Width           =   1920
            End
            Begin VB.ComboBox cboBakspace 
               Height          =   300
               Index           =   2
               Left            =   1180
               Style           =   2  'Dropdown List
               TabIndex        =   88
               Top             =   960
               Width           =   1920
            End
            Begin VB.Label lblBakSpace 
               Alignment       =   1  'Right Justify
               Caption         =   "��׼��"
               Height          =   300
               Index           =   0
               Left            =   360
               TabIndex        =   94
               Top             =   330
               Width           =   645
            End
            Begin VB.Label lblBakSpace 
               AutoSize        =   -1  'True
               Caption         =   "��ʷ��ռ�"
               Height          =   180
               Index           =   3
               Left            =   135
               TabIndex        =   93
               Top             =   0
               Width           =   900
            End
            Begin VB.Label lblBakSpace 
               Alignment       =   1  'Right Justify
               Caption         =   "���ϵͳ"
               Height          =   300
               Index           =   1
               Left            =   240
               TabIndex        =   92
               Top             =   720
               Width           =   765
            End
            Begin VB.Label lblBakSpace 
               Alignment       =   1  'Right Justify
               Caption         =   "����ϵͳ"
               Height          =   300
               Index           =   2
               Left            =   240
               TabIndex        =   91
               Top             =   1080
               Width           =   765
            End
         End
         Begin VB.CheckBox chkBakTbsDisable 
            Caption         =   "������ʷ���Լ��������"
            Height          =   180
            Left            =   1440
            TabIndex        =   80
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.TextBox txtSplit 
            Alignment       =   1  'Right Justify
            Height          =   280
            Left            =   1440
            TabIndex        =   6
            Text            =   "30"
            Top             =   2280
            Width           =   375
         End
         Begin VB.CheckBox chkjob 
            Caption         =   "���õ�ǰϵͳ�����ߵĺ�̨��ҵ"
            Height          =   180
            Left            =   1440
            TabIndex        =   3
            Top             =   1080
            Width           =   2895
         End
         Begin VB.CheckBox chkTrigger 
            Caption         =   "����ת�����ϵĴ�����"
            Height          =   180
            Left            =   1440
            TabIndex        =   4
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton optmode 
            Caption         =   "����ģʽ(�����ж�ҵ�񣬿����ڿͻ�������ʹ�õ�����½��С�)"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   0
            Top             =   285
            Value           =   -1  'True
            Width           =   5775
         End
         Begin VB.OptionButton optmode 
            Caption         =   "����ģʽ(��Ҫ�ж�ҵ��Ҫ�������пͻ���ͣ�õ�����½��С�)"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   1
            Top             =   600
            Width           =   5775
         End
         Begin VB.Label Label13 
            Caption         =   "����ת���ڼ�"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   1065
            Width           =   1455
         End
         Begin VB.Label lblSplit 
            Caption         =   "ÿ��ת��     �������"
            Height          =   255
            Left            =   680
            TabIndex        =   5
            Top             =   2325
            Width           =   2175
         End
      End
      Begin VB.Timer TIMStatus 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   6000
         Top             =   0
      End
      Begin VB.CommandButton cmdPrompt 
         Caption         =   "�鿴ת����֪"
         Height          =   350
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   1335
      End
      Begin VB.CheckBox chkAffirm 
         Caption         =   "������ϸ�Ķ�ת����֪�����������ص�׼���͵�����"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   0
         Width           =   4695
      End
      Begin VB.CommandButton cmdDateLast 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   8805
         Picture         =   "frmDataMove.frx":0680
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1995
         Width           =   280
      End
      Begin VB.CommandButton cmdMoveMark 
         Caption         =   "���ת��"
         Height          =   350
         Left            =   6555
         TabIndex        =   12
         Top             =   4530
         Width           =   1100
      End
      Begin VB.CommandButton cmdMoveOut 
         Caption         =   "ת��(&M)"
         Height          =   350
         Left            =   7995
         TabIndex        =   14
         Top             =   4530
         Width           =   1100
      End
      Begin VB.TextBox txtPrompt 
         Height          =   2175
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   56
         Text            =   "frmDataMove.frx":0776
         Top             =   360
         Visible         =   0   'False
         Width           =   6315
      End
      Begin VB.Label lblDateLast 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ս�ֹ����"
         Height          =   180
         Left            =   6600
         TabIndex        =   83
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblDatePre 
         Caption         =   "�ϴν�ֹ����"
         Height          =   180
         Left            =   6600
         TabIndex        =   11
         Top             =   2820
         Width           =   1080
      End
      Begin VB.Label lblDateThis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ν�ֹ����"
         Height          =   180
         Left            =   6600
         TabIndex        =   9
         Top             =   2445
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "��� 2011-01-01 �� 2011-02-01 ֮�������ʱ�����ж�,��������ת����Щ���ݺ����ִ���µĲ�����"
         ForeColor       =   &H00C00000&
         Height          =   1305
         Left            =   6360
         TabIndex        =   7
         Top             =   600
         Width           =   2865
      End
   End
   Begin VB.Frame fraFunc 
      BorderStyle     =   0  'None
      Height          =   5000
      Index           =   4
      Left            =   0
      TabIndex        =   57
      Tag             =   "ת����"
      Top             =   840
      Width           =   9375
      Begin VB.CommandButton cmdRebIndexOther 
         Caption         =   "�ؽ���������"
         Height          =   350
         Left            =   3360
         TabIndex        =   63
         ToolTipText     =   $"frmDataMove.frx":14B7
         Top             =   2055
         Width           =   1425
      End
      Begin VB.CommandButton cmdShrink 
         Caption         =   "���������ļ�"
         Height          =   350
         Left            =   1800
         TabIndex        =   81
         ToolTipText     =   "������ZLΪǰ׺�ı�ռ�����������ļ���һ��Ӧ����������ؽ�����������ִ�в����ͷſ��пռ䣨�������ļ�β�������µ����ݣ���"
         Top             =   3840
         Width           =   1425
      End
      Begin VB.OptionButton optmode_Index 
         Caption         =   "�����ؽ�(��ͣҵ��)"
         Height          =   180
         Index           =   1
         Left            =   4560
         TabIndex        =   78
         Top             =   1230
         Width           =   2055
      End
      Begin VB.OptionButton optmode_Index 
         Caption         =   "�����ؽ�(�ǳ���ʱ)"
         Height          =   180
         Index           =   0
         Left            =   1560
         TabIndex        =   77
         Top             =   1230
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Frame fraRebuild 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   1440
         TabIndex        =   73
         Top             =   1605
         Width           =   5295
         Begin VB.OptionButton optRebScope_Manual 
            Caption         =   "ȫ��"
            Height          =   375
            Index           =   2
            Left            =   4320
            TabIndex        =   76
            Top             =   0
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optRebScope_Manual 
            Caption         =   "���ú����ࡢҽ����"
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   75
            Top             =   0
            Width           =   2055
         End
         Begin VB.OptionButton optRebScope_Manual 
            Caption         =   "���ú�����"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdRebJobTrigger 
         Caption         =   "�ָ������õĺ�̨��ҵ�ʹ�����"
         Height          =   350
         Left            =   6480
         TabIndex        =   68
         ToolTipText     =   "��������ת������ȫ����ɺ�ִ��"
         Top             =   4440
         Width           =   2745
      End
      Begin VB.CommandButton cmdRebOnline 
         Caption         =   "�ָ����߿ռ䱻���õ�Լ��������"
         Height          =   350
         Left            =   3360
         TabIndex        =   67
         ToolTipText     =   "���ת������Ϊ����ת��������������ؽ����ȽϺ�ʱ������Ӱ��ҵ�����У���������������ؽ�"
         Top             =   4440
         Width           =   2985
      End
      Begin VB.CommandButton cmdMoveTable 
         Caption         =   "����ת����"
         Height          =   350
         Left            =   240
         TabIndex        =   66
         ToolTipText     =   "��������ʷ����ת����ִ��Move������Ȼ��ָ�ʧЧ������"
         Top             =   3840
         Width           =   1425
      End
      Begin VB.CommandButton cmdRebBakSpace 
         Caption         =   "�ָ���ʷ�ռ䱻���õ�Լ��������"
         Height          =   350
         Left            =   240
         TabIndex        =   65
         ToolTipText     =   "����ȫ��ת��������ɺ�ִ�У��Ա���ʷ�ռ�Ĳ�ѯҵ���ܹ���������(�̶����������ؽ�ģʽ)"
         Top             =   4440
         Width           =   2985
      End
      Begin VB.TextBox txtParallel 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   2880
         TabIndex        =   64
         Text            =   "12"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdRebIndexForTag 
         Caption         =   "�ؽ����ת�����������"
         Height          =   350
         Left            =   240
         TabIndex        =   62
         Top             =   2055
         Width           =   2985
      End
      Begin VB.Frame fraMove 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   240
         TabIndex        =   58
         Top             =   3480
         Width           =   4305
         Begin VB.OptionButton optMove 
            Caption         =   "���ú�����"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   60
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optMove 
            Caption         =   "ȫ��"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   59
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "������Χ(��׼��)"
            Height          =   255
            Left            =   0
            TabIndex        =   61
            Top             =   30
            Width           =   1575
         End
      End
      Begin VB.Label lblPrompt 
         ForeColor       =   &H00C00000&
         Height          =   645
         Left            =   3360
         TabIndex        =   82
         Top             =   3645
         Width           =   5775
      End
      Begin VB.Line Line3 
         X1              =   1440
         X2              =   6600
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label lblReIndexMode 
         Caption         =   "�����ؽ���ʽ"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label lblReIndexScope 
         Caption         =   "�����ؽ���Χ"
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblParallel 
         Caption         =   "�����ؽ��ͱ����������Ĳ��ж�"
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   285
         Width           =   6975
      End
      Begin VB.Label lblReIndex 
         Caption         =   "ת�����ݺ���������Ƭ�Ƚ϶࣬����Ӱ�������ѯ��ת�����ݵ�SQL���ܣ�Ҳ��Ӱ������ҵ���е���ز�ѯ���ܣ������ؽ�������"
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label lblMove 
         Caption         =   $"frmDataMove.frx":154E
         Height          =   615
         Left            =   240
         TabIndex        =   69
         Top             =   2760
         Width           =   6855
      End
   End
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4830
      Index           =   1
      Left            =   0
      TabIndex        =   31
      Tag             =   "��ѡ����"
      Top             =   960
      Width           =   9285
      Begin VB.CommandButton cmdMoveIn 
         Caption         =   "���(&I)"
         Height          =   350
         Left            =   6960
         TabIndex        =   26
         Top             =   3840
         Width           =   1100
      End
      Begin VB.TextBox txtPati 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5520
         TabIndex        =   25
         ToolTipText     =   "ͨ�����·�ʽ���룺ֱ��ˢ����-����ID,*�����,.�Һŵ�,+סԺ��,����"
         Top             =   2895
         Width           =   2535
      End
      Begin VB.ComboBox cboPatiType 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2685
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2895
         Width           =   1770
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5520
         TabIndex        =   20
         ToolTipText     =   "�����������ĵ��ݺ�"
         Top             =   1545
         Width           =   2535
      End
      Begin VB.ComboBox cboBillType 
         Height          =   300
         Left            =   2685
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1545
         Width           =   1770
      End
      Begin VB.OptionButton optInType 
         Caption         =   "��ĳ���˳�ѡ����(������������������ݺ�δ���ʷ���)"
         Height          =   195
         Index           =   1
         Left            =   1605
         TabIndex        =   21
         Top             =   2535
         Width           =   5100
      End
      Begin VB.OptionButton optInType 
         Caption         =   "�����ݺų�ѡ����"
         Height          =   195
         Index           =   0
         Left            =   1605
         TabIndex        =   13
         Top             =   1185
         Value           =   -1  'True
         Width           =   5100
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4980
         TabIndex        =   24
         Top             =   2955
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2160
         TabIndex        =   22
         Top             =   2955
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         Height          =   180
         Left            =   4800
         TabIndex        =   19
         Top             =   1605
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2160
         TabIndex        =   17
         Top             =   1605
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "    ͨ���Ѿ�ת���������ǲ��ٲ����ģ�ֻ�ܲ�ѯ������һЩ���������£����Գ�ѡĳЩ�ض������ݷ����������ݱ��Ա�ʵʩ��Ҫ�Ĳ�����"
         Height          =   540
         Left            =   1680
         TabIndex        =   35
         Top             =   405
         Width           =   6195
      End
   End
   Begin VB.Frame fraFunc 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4935
      Index           =   3
      Left            =   0
      TabIndex        =   51
      Tag             =   "ת����־"
      Top             =   960
      Width           =   9255
      Begin VSFlex8Ctl.VSFlexGrid vsflog 
         Height          =   4800
         Left            =   120
         TabIndex        =   52
         Top             =   0
         Width           =   9015
         _cx             =   1995324957
         _cy             =   1995317523
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDataMove.frx":1632
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4890
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Tag             =   "δת��ѯ"
      Top             =   915
      Width           =   9165
      Begin VB.CommandButton cmdData 
         Caption         =   "�����������(&P)"
         Height          =   350
         Index           =   4
         Left            =   7050
         TabIndex        =   53
         Top             =   4080
         Width           =   1620
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "ֱ���շ�����(&A)"
         Height          =   350
         Index           =   0
         Left            =   7050
         TabIndex        =   41
         Top             =   1065
         Width           =   1620
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "���ʷ�������(&B)"
         Height          =   345
         Index           =   1
         Left            =   7050
         TabIndex        =   40
         Top             =   1815
         Width           =   1620
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "�����������(&L)"
         Height          =   350
         Index           =   2
         Left            =   7050
         TabIndex        =   39
         Top             =   2565
         Width           =   1620
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "סԺ��������(&P)"
         Height          =   350
         Index           =   3
         Left            =   7050
         TabIndex        =   38
         Top             =   3330
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2940
         TabIndex        =   28
         Top             =   615
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   39452675
         CurrentDate     =   38471
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1245
         TabIndex        =   27
         Top             =   600
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   39452675
         CurrentDate     =   38471
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   480
         X2              =   6375
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label lblData 
         Caption         =   "��ѯ�޷�ת������������ݼ�����ԭ��"
         Height          =   210
         Index           =   4
         Left            =   465
         TabIndex        =   54
         Top             =   4170
         Width           =   4575
      End
      Begin VB.Label lblData 
         Caption         =   "��ѯ�޷�ת�����ѽ������סԺ���ʣ��Զ����ʵĵ������ݼ�����ԭ��"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   44
         Top             =   1890
         Width           =   6015
      End
      Begin VB.Label lblData 
         Caption         =   "��ѯ�޷�ת��ҽ�����ݵ�������ﲡ����Ϣ������ԭ��"
         Height          =   225
         Index           =   2
         Left            =   480
         TabIndex        =   43
         Top             =   2625
         Width           =   4575
      End
      Begin VB.Label lblData 
         Caption         =   "��ѯ�޷�ת��ҽ�����ݵ�סԺ������Ϣ������ԭ��"
         Height          =   210
         Index           =   3
         Left            =   480
         TabIndex        =   42
         Top             =   3405
         Width           =   4575
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   480
         X2              =   6360
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   480
         X2              =   6360
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   480
         X2              =   6360
         Y1              =   2865
         Y2              =   2865
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   480
         X2              =   6360
         Y1              =   3630
         Y2              =   3630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯʱ��                ��"
         Height          =   180
         Left            =   495
         TabIndex        =   37
         Top             =   675
         Width           =   2340
      End
      Begin VB.Label lblData 
         Caption         =   "��ѯ�޷�ת��������Һţ��շѵĵ������ݼ�����ԭ��"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   45
         Top             =   1155
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "��������ʵĲ�ѯʱ�䷶Χ������ʱ��Ӧ����ת�Ƶ�ʱ�䷶Χ�ڣ�����ʱ������ÿ���Ӱ������޷�ת����ԭ��"
         Height          =   405
         Left            =   135
         TabIndex        =   36
         Top             =   240
         Width           =   9240
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9390
      TabIndex        =   47
      Top             =   5925
      Width           =   9390
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   6600
         TabIndex        =   49
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&C)"
         Height          =   350
         Left            =   8040
         TabIndex        =   48
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.TabStrip tabFunc 
         Height          =   345
         Left            =   120
         TabIndex        =   50
         Tag             =   "ת�Ʋ�ѯ"
         Top             =   165
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   609
         TabWidthStyle   =   2
         MultiRow        =   -1  'True
         Style           =   2
         TabFixedWidth   =   2027
         TabFixedHeight  =   616
         Placement       =   1
         TabMinWidth     =   0
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "����ת��(&1)"
               Key             =   "����ת��"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��ѡ����(&2)"
               Key             =   "��ѡ����"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "δת��ѯ(&3)"
               Key             =   "δת��ѯ"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ת����־(&4)"
               Key             =   "ת����־"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ת����(&5)"
               Key             =   "ת����"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   0
         X2              =   10500
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   10500
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9390
      TabIndex        =   29
      Top             =   0
      Width           =   9390
      Begin MSComctlLib.ImageList img48 
         Left            =   -375
         Top             =   -330
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":1750
               Key             =   "����ת��"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":262A
               Key             =   "��ѡ����"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":3504
               Key             =   "δת��ѯ"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":43DE
               Key             =   "ת����־"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDataMove.frx":E473
               Key             =   "ת����"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ת��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1425
         TabIndex        =   34
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    Ϊ����ϵͳ��Ч���С����ٱ����������������ؽ�������ͳ����Ϣ�ռ������߿ռ�ά����ʱ�䣬���鶨�ڽ���ʷ����ת�Ƶ���ʷ�ռ��С�"
         Height          =   360
         Left            =   1425
         TabIndex        =   33
         Top             =   390
         Width           =   8025
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   240
         Picture         =   "frmDataMove.frx":11905
         Top             =   60
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   11040
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   11040
         Y1              =   840
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frmDataMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mrsMovelog As ADODB.Recordset   '����ת������ʱ�����ת�������ν���ת������
Private mdatBegin As Date
Private mstrPeisPrivs As String         '�����ϵͳ������ת��Ȩ��
Private mlngPeisSys As Long             '�����ϵͳ���
Private mlngOperSys As Long             '������ϵͳ���
Private mblnDBA As Boolean
Private mlngMinDays As Long, mlngMaxDays As Long
Private mblnOffLineMoved As Boolean            '�Ƿ�ִ����ת����������û�лָ����߿ռ��Լ��������

Private Sub cboBakspace_Click(Index As Integer)
    If Index = 0 And Me.Visible Then
        Dim strText As String
        Dim i As Long, j As Long
        
        strText = cboBakspace(Index).Text
        For i = 1 To 2
            For j = 0 To cboBakspace(i).ListCount
                If cboBakspace(i).List(j) = strText Then
                    cboBakspace(i).ListIndex = j
                    Exit For
                End If
            Next
        Next
    End If
End Sub

Private Sub cboBillType_Click()
    txtNO.Text = ""
End Sub

Private Sub cboPatiType_Click()
    txtPati.Text = ""
    txtPati.Tag = ""
    cboPatiType.Tag = ""
    
    Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
    Case 0, 1
        txtPati.ToolTipText = "ͨ�����·�ʽ���룺ֱ��ˢ����-����ID,*�����,.�Һŵ�,+סԺ��,����"
        lblPatient.Caption = "����"
    Case 2
        txtPati.ToolTipText = "ͨ�����·�ʽ���룺ֱ��ˢ����-����ID,*�����,+������,����"
        lblPatient.Caption = "����"
    Case 3
        txtPati.ToolTipText = "ͨ�����·�ʽ���룺-����ID,����"
        lblPatient.Caption = "����"
    End Select
    
End Sub

Private Sub chkAffirm_Click()
    If txtPrompt.Visible Then txtPrompt.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdData_Click(Index As Integer)
    If dtpBegin.value > dtpEnd.value Then
        MsgBox "��ʼʱ��ӦС�ڽ���ʱ�䡣", vbInformation, gstrSysName
        dtpBegin.SetFocus: Exit Sub
    End If
    
    If MsgBox("���ָ��ʱ���е�δת�����ݽ϶࣬��ѯ������Ҫ�ϳ�ʱ�䡣" & vbCrLf & "����ִ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call frmDataMoveQuery.ShowMe(Index, dtpBegin, dtpEnd, Split(cmdData(Index).Caption, "(")(0), lblData(Index).Caption, Me)
End Sub

Private Sub cmdDateThis_Click()
'���ܣ�������ѡ����
        If IsDate(txtDateThis.Text) Then monSel.value = CDate(txtDateThis.Text)
        
        monSel.Tag = "txtDateThis"
        monSel.Left = Me.ScaleLeft + Me.ScaleWidth - monSel.Width - 120
        monSel.Top = txtDateThis.Top + txtDateThis.Height + 30
        monSel.ZOrder
        monSel.Visible = True
        monSel.SetFocus
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdMoveIn_Click()
    Dim lng����ID As Long, str���� As String
    Dim blnMoved As Boolean
    
    If optInType(0).value Then
        If txtNO.Text = "" Then
            MsgBox "�����뵥�ݺš�", vbInformation, gstrSysName
            txtNO.SetFocus: Exit Sub
        End If
        
        '   "1-�շѵ���","2-���ʵ���","3-�Զ�����","4-�Һŵ���","5-���￨","6-Ԥ������","7-���ʵ���"
        '��鵥���Ƿ��Ѿ�ת��
        Select Case cboBillType.ItemData(cboBillType.ListIndex)
        Case 2 '���ʵ���(���ܴ���������ʺ�סԺ���ʵ����,������Ҫ����������)
            blnMoved = MovedByNO(txtNO.Text, "���˷��ü�¼", "��¼����=[2] ")
        Case 3, 5           ',�Զ�����,���￨
            blnMoved = MovedByNO(txtNO.Text, "סԺ���ü�¼", "��¼����=[2] ")
        Case 1, 4   '�շѵ���,,�Һŵ���
            blnMoved = MovedByNO(txtNO.Text, "������ü�¼", "��¼����=[2] ")
        Case 6 'Ԥ������
            blnMoved = MovedByNO(txtNO.Text, "����Ԥ����¼", "��¼����=1")
        Case 7 '���ʵ���
            blnMoved = MovedByNO(txtNO.Text, "���˽��ʼ�¼")
        Case 8  '�������
            blnMoved = MovedByPeis(1, txtNO.Text)
        End Select
        
        If Not blnMoved Then
            MsgBox Replace(Mid(cboBillType.Text, 3), "����", "") & "���� " & txtNO.Text & " û��ת����", vbInformation, gstrSysName
            txtNO.SetFocus: Exit Sub
        End If
        
        If MsgBox("���ڽ���" & Replace(Mid(cboBillType.Text, 3), "����", "") & "���� " & txtNO.Text & " �����ݳ���������ݿ⣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    ElseIf optInType(1).value Then
        If txtPati.Tag = "" Then
            MsgBox "�����벡�ˡ�", vbInformation, gstrSysName
            txtPati.SetFocus: Exit Sub
        End If
        
        Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
        Case 0, 1
            lng����ID = Val(Split(txtPati.Tag, ",")(0))
            str���� = CStr(Split(txtPati.Tag, ",")(1))
            
            Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
            Case 0          '���ﲡ��(����id,�Һŵ���)
                blnMoved = MovedByNO(str����, "���˹Һż�¼")
            Case 1          'סԺ����(����id,��ҳid)
                blnMoved = MovedByPati(lng����ID, Val(str����))
            End Select
            
        Case 2         '�ܼ���Ա(����id)
            lng����ID = Val(txtPati.Tag)
            blnMoved = MovedByPeis(2, Val(txtPati.Tag))
        Case 3         '�ܼ�����(����id)
            lng����ID = Val(txtPati.Tag)
            blnMoved = MovedByPeis(3, Val(txtPati.Tag))
            
        End Select
        
        If Not blnMoved Then
            MsgBox Mid(cboPatiType.Text, 3) & " " & txtPati.Text & "���������û��ת����", vbInformation, gstrSysName
            txtPati.SetFocus: Exit Sub
        End If
        
        If MsgBox("���ڽ���" & Mid(cboPatiType.Text, 3) & " " & txtPati.Text & " ������������ݺ�δ���ʷ��ó���������ݿ⣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
   
    End If
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    If optInType(0).value Then
        Select Case cboBillType.ItemData(cboBillType.ListIndex)
        Case 1, 2, 3, 4, 5, 6, 7
            gstrSQL = "Zl_Retu_Exes('" & txtNO.Text & "'," & cboBillType.ItemData(cboBillType.ListIndex) & ")"
        Case 8
            gstrSQL = "zl_Return_Peis(3,'" & txtNO.Text & "')"
        End Select
        
    ElseIf optInType(1).value Then
        Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
        Case 0, 1
            gstrSQL = "Zl_Retu_Clinic(" & lng����ID & ",'" & str���� & "'," & cboPatiType.ItemData(cboPatiType.ListIndex) & ")"
        Case 2
            gstrSQL = "zl_Return_Peis(1,'" & lng����ID & "')"
        Case 3
            gstrSQL = "zl_Return_Peis(2,'" & lng����ID & "')"
        End Select
    Else
        gstrSQL = "Zl_Retu_Clinic(0,'" & str���� & "',2)"
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    MsgBox "���ݳ�ѡ������ִ����ɡ�", vbInformation, gstrSysName
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RefreshMove() As Boolean
'���ܣ�ˢ��ת����������Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strMsg As String, strTagStartDate As String
    Dim datCurr As Date, blnFirst As Boolean, blnWaitMove As Boolean, blnWaitTag As Boolean
    Dim blnDo As Boolean
    Dim lngTmpSysNO As Long, lngDays As Long
        
    On Error GoTo errH
    
   
     '���봰��ʱ����ȱʡ���ж�
    If Me.Visible = False Then
        If mblnDBA Then
            gstrSQL = "Select Value From V$parameter Where Name = 'cpu_count'"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys)
            If rsTmp.EOF Then
                txtParallel.Text = "0"
                txtParallel.Enabled = False
                lblParallel.Caption = "DDL�������ж�        (δ�ҵ����жȲ���cpu_count)"
            Else
                txtParallel.Tag = "" & rsTmp!value
                If Val(rsTmp!value) < 3 Then
                    txtParallel.Text = "0"
                    txtParallel.Enabled = False
                    lblParallel.Caption = "DDL�������ж�        (cpu����С��3�����ز��ò���ִ��)"
                    
                ElseIf Val(rsTmp!value) < 13 Then
                    txtParallel.Text = Val(rsTmp!value) \ 2 'һ��ȡ��
                Else
                    txtParallel.Text = "12"  '��ʹcpu�㹻�����Կ��������ڴ������ܣ����жȲ���Խ��Խ��
                End If
            End If
        Else
            txtParallel.Text = "0"
        End If
                
        
        '��ʷ��ռ�ĳ�ʼ��
        gstrSQL = "Select ϵͳ, ���, ����, ������, ��ǰ From Zlbakspaces "
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        If rsTmp.EOF Then
            MsgBox "ϵͳ��δ������ʷ��ռ䣬�����ܲ���ʹ�á�", vbExclamation, gstrSysName
            Exit Function
        End If
        For i = 0 To 2
            cboBakspace(i).Clear
            lngTmpSysNO = Decode(i, 0, glngSys, 1, mlngPeisSys, 2, mlngOperSys)
            If lngTmpSysNO > 0 Then
                rsTmp.Filter = "ϵͳ=" & lngTmpSysNO
                rsTmp.Sort = "���"
                Do While Not rsTmp.EOF
                    cboBakspace(i).AddItem NVL(rsTmp!����)
                    cboBakspace(i).ItemData(cboBakspace(i).NewIndex) = Val(NVL(rsTmp!���))
                    If NVL(rsTmp!��ǰ, 0) = 0 Then cboBakspace(i).ListIndex = cboBakspace(i).NewIndex
                    rsTmp.MoveNext
                Loop
                If cboBakspace(i).ListCount > 0 And cboBakspace(i).ListIndex < 0 Then cboBakspace(i).ListIndex = 0
            End If
            cboBakspace(i).Visible = cboBakspace(i).ListCount > 1
            lblBakSpace(i).Visible = cboBakspace(i).ListCount > 1
            '������ʾ����λ��
            If (Not cboBakspace(i).ListCount > 1) And i < 2 Then
                If i = 0 Then
                    cboBakspace(2).Top = cboBakspace(1).Top
                    lblBakSpace(2).Top = lblBakSpace(1).Top
                End If
                cboBakspace(i + 1).Top = cboBakspace(i).Top
                lblBakSpace(i + 1).Top = lblBakSpace(i).Top
            End If
        Next i
        picBakspace.Visible = False
        For i = 0 To 2
            If cboBakspace(i).ListCount > 1 Then
                picBakspace.Visible = True
                Exit For
            End If
        Next i
    End If
        
    gstrSQL = "Select �ϴ�����,������������ From zlDataMove Where ϵͳ=[1] And ���=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys)
    If rsTmp.EOF Then
        MsgBox "ϵͳ��δ������Ч������ת�ƶ��壬�����ܲ���ʹ�á�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If IsNull(rsTmp!������������) = False Then
        txtDateLast.Text = Format(rsTmp!������������, "yyyy-mm-dd")
        txtDateLast.Enabled = False
    Else
        txtDateLast.Enabled = True
    End If
    cmdDateLast.Enabled = txtDateLast.Enabled
    
    gstrSQL = "Select ����,����,��ֹʱ��,��ת��,��ǿ�ʼʱ��,��ǽ���ʱ��,ת����ʼʱ��,ת������ʱ��,�ؽ�����ʱ��" & _
            " From zlDataMovelog Where ϵͳ = [1] Order by ����"
    Set mrsMovelog = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys)
    If mrsMovelog.RecordCount > 0 Then
        '������ڱ��ת������ģ���Ը�ʱ������±��ת��������ת����ᵼ�����ݲ�һ�¡�
        mrsMovelog.Filter = "��ת��=2"
        If mrsMovelog.RecordCount > 0 Then
            mdatBegin = mrsMovelog!��ֹʱ��
            blnWaitTag = True
        Else
            mrsMovelog.Filter = "��ת��=1"
            If mrsMovelog.RecordCount > 0 Then
                mrsMovelog.MoveLast
                mdatBegin = mrsMovelog!��ֹʱ��
                blnWaitMove = True
            End If
            mrsMovelog.Filter = ""
            mrsMovelog.MoveFirst
        End If
    End If
    
    
    '���ת����û�и����ϴ�����
    If IsNull(rsTmp!�ϴ�����) Or blnWaitTag Then
        If Not blnWaitMove And Not blnWaitTag Then blnFirst = True
        
        If blnWaitTag Then
            'ȡ��һ�α��ת����ת����ʱ��
            gstrSQL = "Select ��ֹʱ�� as �ϴ����� From zlDataMovelog Where ϵͳ = [1] And Nvl(��ת��,0)<>2 Order by ��ֹʱ�� Desc "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys)
            blnDo = rsTmp.RecordCount = 0
        Else
            blnDo = True
        End If
        
        If blnDo Then
            gstrSQL = "Select Min(�Ǽ�ʱ��) �ϴ����� From (Select Min(�Ǽ�ʱ��) �Ǽ�ʱ�� From ������ü�¼ Union All Select Min(�Ǽ�ʱ��) From ���˹Һż�¼)"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
            
            If IsNull(rsTmp!�ϴ�����) Then
                MsgBox "��ǰϵͳû�з�������ҺŻ��շ�ҵ�����ݣ������ܲ���ʹ�á�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    strTagStartDate = Format(rsTmp!�ϴ�����, "yyyy-MM-dd")
    txtDatePre.Text = strTagStartDate
    
    If blnWaitTag Then
        lblDateThis.Caption = "�����������"
    ElseIf blnWaitMove Then
        lblDateThis.Caption = "����ת������"
    Else
        mdatBegin = rsTmp!�ϴ�����
        lblDateThis.Caption = "���ν�ֹ����"
    End If
   
    datCurr = zlDatabase.Currentdate
    
    
    If blnWaitTag Then
        mlngMaxDays = DateDiff("d", rsTmp!�ϴ�����, datCurr)
    Else
        mlngMaxDays = DateDiff("d", mdatBegin, datCurr)
    End If
    
    mlngMinDays = 365   '���ٱ���һ�������
    If mlngMinDays > mlngMaxDays Then mlngMinDays = mlngMinDays - 1
        
    If blnWaitTag Or blnWaitMove Then
        cmdMoveOut.Enabled = blnWaitMove
        cmdMoveMark.Enabled = blnWaitTag
        
        txtDateThis.Text = Format(mdatBegin, "yyyy-MM-dd")
        txtDateThis.Enabled = False
    Else
        cmdMoveOut.Enabled = True
        cmdMoveMark.Enabled = True
        
        txtDateThis.Enabled = True
    
        'ȱʡ�����ս�ֹ����Ϊ������������
        If txtDateLast.Enabled Then
            gstrSQL = "Select Trunc(add_months(Sysdate,-24*3) ,'yyyy') As year_firstday From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            
            lngDays = DateDiff("d", Format(rsTmp!year_firstday, "yyyy-mm-dd"), datCurr)
        
            If mlngMaxDays < lngDays Then lngDays = mlngMaxDays - 365
            If lngDays < 0 Then lngDays = mlngMaxDays - 180
            If lngDays < 0 Then lngDays = 0
            
            txtDateLast.Text = Format(datCurr - lngDays, "yyyy-mm-dd")
        End If
        
        'ȱʡһ��תһ��
        txtDateThis.Text = Format(datCurr - mlngMaxDays + 365, "yyyy-mm-dd")
        If CDate(txtDateThis.Text) > CDate(txtDateLast.Text) Then txtDateThis.Text = txtDateLast.Text
    End If
    cmdDateThis.Enabled = txtDateThis.Enabled
    
            
    If blnFirst Then
        strMsg = "�Ӵ��ڹҺŻ��շ����ݵ� " & Format(mdatBegin, "yyyy-MM-dd") & " ��ʼת������"
        dtpEnd.MaxDate = Int(DateAdd("d", -90, datCurr) - 1)
    Else
        If blnWaitMove Then
            strMsg = "�Ѿ������ " & strTagStartDate & " �� " & Format(mdatBegin, "yyyy-MM-dd") & " ֮�������" & vbCrLf & "����ת����Щ���ݺ����ִ���µ�ת��������"
        ElseIf blnWaitTag Then
            strMsg = "��� " & strTagStartDate & " �� " & Format(mdatBegin, "yyyy-MM-dd") & " ֮�������ʱ�����ж�" & vbCrLf & "��������ת����Щ���ݺ����ִ���µĲ�����"
        Else
            strMsg = "�ϴ��Ѿ�ת���� " & Format(mdatBegin, "yyyy-MM-dd") & " ��ǰ������"
        End If
        dtpEnd.MaxDate = Int(mdatBegin - 1)
    End If
    lblStatus.Caption = strMsg
    
    
    '����δת��ѯ
    dtpBegin.MaxDate = dtpEnd.MaxDate
    If Not Visible Then
        dtpEnd.value = dtpEnd.MaxDate
        dtpBegin.value = DateAdd("d", -30, dtpEnd.value)
    End If
    For i = 0 To cmdData.UBound
        cmdData(i).Enabled = Not blnFirst
    Next
    dtpBegin.Enabled = Not blnFirst
    dtpEnd.Enabled = Not blnFirst
        
    RefreshMove = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAffirm() As Boolean
    If chkAffirm.value = 0 Then
        MsgBox "��ȷ��������ϸ�Ķ����������ʷ����ת��ǰ����е�׼���͵�����", vbInformation, gstrSysName
        chkAffirm.SetFocus
    Else
        CheckAffirm = True
    End If
End Function

Private Sub cmdMoveMark_Click()
'���ܣ�ִ�б��ת��
    Dim datCurr As Date, datBegin As Date, strTime As String, lngTotaltime As Long
    Dim lngBeginDays As Long, i As Long, lngEndDays As Long, lngCurrDays As Long, bytSpeedMode As Byte
    Dim lngSplit As Long, lngAddDay As Long
    Dim rsTmp As ADODB.Recordset
    
    Dim strBakUser As String, strPeisBakUser As String, strOperBakUser As String, blnNoData As Boolean
    
    If CheckAffirm = False Then Exit Sub
    If Not CheckDate(1) Then Exit Sub
          
    If Not IsNumeric(txtSplit.Text) Then
        MsgBox "��������Ч�ļ��������", vbInformation, gstrSysName
        txtSplit.SetFocus: Exit Sub
    End If
    lngSplit = Val(txtSplit.Text)
    
    If CheckData = False Then Exit Sub
    
    
    If MsgBox("������ת�������ݽ϶࣬������Ҫ�ϳ�ʱ�䡣" & vbCrLf & "��ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
    On Error GoTo errH
    
    blnNoData = cboBakspace(0).ListCount = 0
    If blnNoData = False Then
        gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys, cboBakspace(0).ItemData(cboBakspace(0).ListIndex))
        blnNoData = rsTmp.RecordCount = 0
    End If
    If blnNoData = True Then
        MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ��ʷ�ռ䣡", vbInformation, gstrSysName
        Exit Sub
    End If
    strBakUser = rsTmp!������
    
    
     '�����ϵͳ���ж�
    If mlngPeisSys > 0 Then
        blnNoData = cboBakspace(1).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPeisSys, cboBakspace(1).ItemData(cboBakspace(1).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
        If blnNoData = True Then
            MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ�����ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
            Exit Sub
        End If
        strPeisBakUser = rsTmp!������
    End If
    
    '������ϵͳ���ж�
    If mlngOperSys > 0 Then
       blnNoData = cboBakspace(2).ListCount = 0
       If blnNoData = False Then
           gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
           Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOperSys, cboBakspace(2).ItemData(cboBakspace(2).ListIndex))
           blnNoData = rsTmp.RecordCount = 0
       End If
        If blnNoData = True Then
           MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ������ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
           Exit Sub
       End If
       strOperBakUser = rsTmp!������
    End If
    
    
    datBegin = zlDatabase.Currentdate
    
    lngEndDays = DateDiff("d", CDate(txtDateThis.Text), datBegin)
    lngBeginDays = mlngMaxDays
    bytSpeedMode = IIF(optmode(0).value, 0, 1)
    
    
    If (lngBeginDays - lngEndDays) Mod lngSplit = 0 Then
        lngTotaltime = (lngBeginDays - lngEndDays) \ lngSplit
    Else
        lngTotaltime = (lngBeginDays - lngEndDays) \ lngSplit + 1
    End If
    
    Screen.MousePointer = 11
    TIMStatus.Enabled = True    '���ö�ʱˢ������ʾ����
    TIMStatus.Tag = "���ת��"
    SetCommandEnable False
    
    If txtDateLast.Enabled Then
        gcnOracle.Execute "Update zlDataMove Set ������������ = to_date('" & txtDateLast.Text & "','yyyy-mm-dd')"
    End If
    
    For i = 1 To lngTotaltime
        datCurr = zlDatabase.Currentdate
        lngAddDay = DateDiff("d", datBegin, datCurr)    'ת���ڼ���ܿ���
        
        lngBeginDays = lngBeginDays - lngSplit
        lngCurrDays = IIF(lngBeginDays > lngEndDays, lngBeginDays, lngEndDays) + lngAddDay
        
        lblStatus.Caption = "���ڱ��" & Format(DateAdd("d", -lngCurrDays, datCurr), "yyyy-MM-dd") & "ǰ������(" & i & "/" & lngTotaltime & ")�������ĵȴ� �� ��"
        lblStatus.Refresh
        gstrSQL = "zl1_DataMoveOut1(" & lngCurrDays & ",1," & i & "," & lngTotaltime & "," & bytSpeedMode & "," & chkTrigger.value & "," & chkjob.value & "," & _
                     "0,'" & strBakUser & "','" & strPeisBakUser & "','" & strOperBakUser & "')"
        DoEvents
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    TIMStatus.Enabled = False
    TIMStatus.Tag = ""
    SetCommandEnable True
    
    strTime = GetTimeString(datBegin, zlDatabase.Currentdate)
    Screen.MousePointer = 0
    MsgBox "���ת��ִ����ɣ����ι���ʱ��" & strTime & "��", vbInformation, gstrSysName
    
    Call RefreshMove
    
    Exit Sub
    
errH:
    TIMStatus.Enabled = False
    TIMStatus.Tag = ""
    SetCommandEnable True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
'        Screen.MousePointer = 11
'        Resume
    End If
    Call SaveErrLog
    Call RefreshMove
End Sub

Private Function CheckPrivilegeOfTrigger(ByRef cnThis As ADODB.Connection) As Boolean
'���ܣ���鵱ǰ���Ӷ�����û����Ƿ��д�����������Ȩ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select 1 From User_Sys_Privs Where Privilege in ('CREATE TRIGGER','CREATE ANY TRIGGER')"
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
    
    CheckPrivilegeOfTrigger = rsTmp.RecordCount > 0
End Function


Private Function GrantPrivilegeOfTrigger() As Boolean
'���ܣ���Ӧ��ϵͳ�������ߣ���鲢ִ�д�������������Ȩ
    Dim rsTmp As ADODB.Recordset
    Dim cnDBA As ADODB.Connection
    Dim strOwner As String
    
        
    If CheckPrivilegeOfTrigger(gcnOracle) = False Then
        Call zlDatabase.UserIdentify(Me, "ʹ��DBA�û���Ӧ��ϵͳ�����������贴����������Ȩ�ޡ�", glngSys, 0, "system", cnDBA, True)
        If cnDBA Is Nothing Then
            MsgBox "�û���¼ʧ�ܣ�������ǰ������", vbInformation, gstrSysName
            Exit Function
        End If
        
        'ȡӦ��ϵͳ��������(��������ϵͳ�Ǳ�׼�����ϵͳ������ͬ��������)
        gstrSQL = "Select Trunc(��� / 100) as ���, ������ From zlSystems Where Trunc(��� / 100) = 1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rsTmp.RecordCount = 0 Then
            MsgBox "��ZLSYSTEMS��ȡϵͳ��Ϣʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        strOwner = rsTmp!������
        
        cnDBA.Execute "grant create trigger to " & strOwner
        cnDBA.Close
    End If
    
    GrantPrivilegeOfTrigger = True
End Function


Private Sub cmdMoveOut_Click()
'���ܣ�ִ��ת��
    Dim datCurr As Date, datBegin As Date, strTime As String, lngTotaltime As Long
    Dim strMsg As String
    Dim i As Long, lngAddDay As Long
    Dim strBakUser As String, strPeisBakUser As String, strOperBakUser As String
    Dim cnBakDB As ADODB.Connection, cnPeisBakDB As New ADODB.Connection, cnOperBakDB As New ADODB.Connection
        
    Dim rsTmp As ADODB.Recordset
    Dim blnRollBack As Boolean, lngTag As Long
    Dim lngBeginDays As Long, lngEndDays As Long, lngCurrDays As Long, bytSpeedMode As Byte
    Dim lngSplit As Long
    Dim blnNoData As Boolean
    
    On Error GoTo errH
    If CheckAffirm = False Then Exit Sub
    
    
    If Not IsNumeric(txtSplit.Text) Then
        MsgBox "��������Ч�ļ��������", vbInformation, gstrSysName
        txtSplit.SetFocus: Exit Sub
    End If
    lngSplit = Val(txtSplit.Text)
    
    bytSpeedMode = IIF(optmode(0).value, 0, 1)
    mrsMovelog.Filter = "��ת��=1"
    lngTag = mrsMovelog.RecordCount
    
    '���ѱ��ת��������ִ��ת������ʱ���ü��������
    If lngTag = 0 Then
        If Not CheckDate(0) Then Exit Sub
    End If
    
    If CheckData = False Then Exit Sub
    
        
    If bytSpeedMode = 1 Then
        strMsg = "��ѡ��������ģʽ������ת���ڼ佫����ò���������Լ�����⽫���±�ϵͳ�����пͻ��˲����á�" & vbCrLf & _
                "��ȷ��Ҫ������"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If lngTag > 0 Then
            strMsg = "��ת�������ѱ�ǵ�����," & vbCrLf & "������ݽ϶࣬������Ҫ�ϳ�ʱ�䡣" & vbCrLf & "��ȷ��Ҫ������"
        Else
            strMsg = "���ת�����ݽ϶࣬������Ҫ�ϳ�ʱ�䡣" & vbCrLf & "��ȷ��Ҫ������"
        End If
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    
    
    '������ʷ�ռ��������Լ�������߿�Ĵ�����zl1_DataMoveOut1�н��У�,��ʹ����ģʽҲ������ã���߲������ܣ��Լ��������������־
    blnNoData = cboBakspace(0).ListCount = 0
    If blnNoData = False Then
        gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys, cboBakspace(0).ItemData(cboBakspace(0).ListIndex))
        blnNoData = rsTmp.RecordCount = 0
    End If
    If blnNoData = True Then
        MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ��ʷ�ռ䣡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strBakUser = rsTmp!������
    
    If chkBakTbsDisable.value = 1 Then
        Call zlDatabase.UserIdentify(Me, "��ʷ�ռ��û���֤", glngSys, 0, strBakUser, cnBakDB, True)
        If cnBakDB Is Nothing Then
            MsgBox "ת��ǰ��Ҫ�Ƚ�����ʷ�ռ��Լ����������������������ʷ�ռ䣡", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�����ϵͳ���ж�
    If mlngPeisSys > 0 Then
        blnNoData = cboBakspace(1).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPeisSys, cboBakspace(1).ItemData(cboBakspace(1).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
        If blnNoData = True Then
            MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ�����ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strPeisBakUser = rsTmp!������
        
        If chkBakTbsDisable.value = 1 Then
            If strBakUser = strPeisBakUser Then
                Set cnPeisBakDB = cnBakDB
            Else
                Call zlDatabase.UserIdentify(Me, "�����ϵͳ��ʷ�ռ��û���֤", mlngPeisSys, 0, strPeisBakUser, cnPeisBakDB, True)
                If cnPeisBakDB Is Nothing Then
                    MsgBox "ת��ǰ��Ҫ�Ƚ��������ϵͳ��ʷ�ռ��Լ�������������������������ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '������ϵͳ���ж�
    If mlngOperSys > 0 Then
        blnNoData = cboBakspace(2).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOperSys, cboBakspace(2).ItemData(cboBakspace(2).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
         If blnNoData = True Then
            MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ������ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
            Exit Sub
        End If
        strOperBakUser = rsTmp!������
        
        If chkBakTbsDisable.value = 1 Then
            If strBakUser = strOperBakUser Then
                Set cnOperBakDB = cnBakDB
            Else
                Call zlDatabase.UserIdentify(Me, "������ϵͳ��ʷ�ռ��û���֤", mlngOperSys, 0, strOperBakUser, cnOperBakDB, True)
                If cnOperBakDB Is Nothing Then
                    MsgBox "ת��ǰ��Ҫ�Ƚ���������ϵͳ��ʷ�ռ��Լ��������������������������ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
         
    '��鲢����ϵͳ�������û�������ɾ����������Ȩ�ޣ��Ա����߿����Լ��ʱ��Ϊ����ɾ��������õ���������ʱ������
    '������ͨ����̬SQL���������������ԣ���ʹϵͳ��������DBA��ɫ��Ҳ��Ҫ��ʽ��Ȩ����ʹ�����Ľ�ɫRESOURCE�д�����������Ȩ�ޣ�
    If lngTag = 0 And bytSpeedMode = 0 Then
        If GrantPrivilegeOfTrigger = False Then Exit Sub
    End If
    
    Screen.MousePointer = 11
    datBegin = zlDatabase.Currentdate
           
    If chkBakTbsDisable.value = 1 Then
        lblStatus.Caption = "���ڽ�����ʷ�ռ��Լ���������������ĵȴ� �� ��"
        Call SetConstraintStatus(glngSys, cnBakDB, False)
        If mlngPeisSys > 0 Then Call SetConstraintStatus(mlngPeisSys, cnPeisBakDB, False)
        If mlngOperSys > 0 Then Call SetConstraintStatus(mlngOperSys, cnOperBakDB, False)
        
        Call SetIndexStatus(glngSys, cnBakDB, False)
        If mlngPeisSys > 0 Then Call SetIndexStatus(mlngPeisSys, cnPeisBakDB, False)
        If mlngOperSys > 0 Then Call SetIndexStatus(mlngOperSys, cnOperBakDB, False)
    End If

        
    TIMStatus.Enabled = True    '���ö�ʱˢ������ʾʱ��
    TIMStatus.Tag = "ת��"
    SetCommandEnable False

    
    blnRollBack = True
    If lngTag > 0 Then  'a.���ݱ��ת������ת��
        For i = 1 To mrsMovelog.RecordCount
            lblStatus.Caption = "����ת��" & Format(mrsMovelog!��ֹʱ��, "yyyy-MM-dd") & "֮ǰ������(" & i & "/" & mrsMovelog.RecordCount & ")�������ĵȴ� �� ��"
            Me.Refresh
            gstrSQL = "zl1_DataMoveOut1(0,2," & i & "," & lngTag & "," & bytSpeedMode & "," & chkTrigger.value & "," & chkjob.value & ",0,'" & strBakUser & "','" & strPeisBakUser & "','" & strOperBakUser & "')"
            DoEvents
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
    Else                'b.ֱ��ת��
        lngEndDays = DateDiff("d", CDate(txtDateThis.Text), datBegin)
        lngBeginDays = mlngMaxDays
        
        If (lngBeginDays - lngEndDays) Mod lngSplit = 0 Then
            lngTotaltime = (lngBeginDays - lngEndDays) \ lngSplit
        Else
            lngTotaltime = (lngBeginDays - lngEndDays) \ lngSplit + 1
        End If
        
        For i = 1 To lngTotaltime
            datCurr = zlDatabase.Currentdate
            lngAddDay = DateDiff("d", datBegin, datCurr)    'ת���ڼ���ܿ���
            
            lngBeginDays = lngBeginDays - lngSplit
            lngCurrDays = IIF(lngBeginDays > lngEndDays, lngBeginDays, lngEndDays) + lngAddDay
            
            
            lblStatus.Caption = "����ת��" & Format(DateAdd("d", -lngCurrDays, datCurr), "yyyy-MM-dd") & "ǰ������(" & i & "/" & lngTotaltime & ")�������ĵȴ� �� ��"
            lblStatus.Refresh
            gstrSQL = "zl1_DataMoveOut1(" & lngCurrDays & ",0," & i & "," & lngTotaltime & "," & bytSpeedMode & "," & chkTrigger.value & "," & chkjob.value & ",0,'" & strBakUser & "','" & strPeisBakUser & "','" & strOperBakUser & "')"
            DoEvents
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
    End If
    blnRollBack = False
    TIMStatus.Enabled = False
    TIMStatus.Tag = ""
    SetCommandEnable True
    mblnOffLineMoved = True
               
    If chkBakTbsDisable.value = 1 Then
        cnBakDB.Close
        If mlngPeisSys > 0 And cnPeisBakDB.State = adStateOpen Then cnPeisBakDB.Close
        If mlngOperSys > 0 And cnOperBakDB.State = adStateOpen Then cnOperBakDB.Close
    End If
    
    strTime = GetTimeString(datBegin, zlDatabase.Currentdate)
    Screen.MousePointer = 0
    MsgBox "����ת��ִ����ɣ����ι���ʱ��" & strTime & "��", vbInformation, gstrSysName
    
    Call RefreshMove
    
    Exit Sub
    
    
errH:
    TIMStatus.Enabled = False
    TIMStatus.Tag = ""
    SetCommandEnable True
    If blnRollBack And chkBakTbsDisable.value = 1 Then
        cnBakDB.Close
        If mlngPeisSys > 0 And cnPeisBakDB.State = adStateOpen Then cnPeisBakDB.Close
        If mlngOperSys > 0 And cnOperBakDB.State = adStateOpen Then cnOperBakDB.Close
    End If
    
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        'Screen.MousePointer = 11
        'Resume
    End If
    Call SaveErrLog
    Call RefreshMove
End Sub

Private Function GetTimeString(ByVal datBegin As Date, ByVal datEnd As Date) As String
'���ܣ���ȡ����ʱ��ֵ��ĸ�ʽ�ַ���
'   datBegin=��ʼʱ��
'   datEnd=��ֹʱ��
    Dim intH As Integer, intM As Integer, intS As Integer
    Dim datTmp As Date

    intH = DateDiff("h", datBegin, datEnd)
    datTmp = DateAdd("h", intH, datBegin)
    intM = DateDiff("n", datTmp, datEnd)
    datTmp = DateAdd("n", intM, datTmp)
    intS = DateDiff("s", datTmp, datEnd)
    
    If intS < 0 Then
        intM = intM - 1
        intS = 60 + intS
    End If
    
    If intM < 0 Then
        intH = intH - 1
        intM = 60 + intM
    End If
    GetTimeString = IIF(intH <> 0, intH & "Сʱ", "") & IIF(intM <> 0, intM & "��", "") & intS & "��"
End Function

Private Sub cmdPrompt_Click()
    If txtPrompt.Visible = False Then
        txtPrompt.Top = cmdPrompt.Top + cmdPrompt.Height + 30
        txtPrompt.Left = cmdPrompt.Left
        txtPrompt.Height = Me.Height - (fraFunc(0).Top + cmdPrompt.Height + 120) - (PicBottom.Height + 120) - 240 '(������)
        txtPrompt.Width = Me.Width - 120 - 240
        
        txtPrompt.ZOrder
        chkAffirm.value = 0
        txtPrompt.Visible = True
    End If
End Sub

Private Sub cmdRebBakSpace_Click()
'���ܣ��ָ���ʷ�ռ䱻���õ�Լ��������
    Dim strBakUser As String
    Dim cnBakDB As ADODB.Connection
    Dim strPeisBakUser As String
    Dim cnPeisBakDB As New ADODB.Connection
    Dim strOperBakUser As String
    Dim cnOperBakDB As New ADODB.Connection
    Dim rsTmp As ADODB.Recordset
    Dim strParallel As String, strTime As String
    Dim datCurr  As Date
    Dim blnNoData As Boolean
    
    
    If MsgBox("�ò����ǳ���ʱ����ȷ��Ҫ�ָ���ʷ�ռ䱻���õ�Լ����������", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo errH
    strParallel = Val(txtParallel.Text)
    blnNoData = cboBakspace(0).ListCount = 0
    If blnNoData = False Then
        gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngSys, cboBakspace(0).ItemData(cboBakspace(0).ListIndex))
        blnNoData = rsTmp.RecordCount = 0
    End If
    If blnNoData = True Then
        MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ��ʷ�ռ䣡", vbInformation, gstrSysName
        Exit Sub
    End If
    strBakUser = rsTmp!������
    Call zlDatabase.UserIdentify(Me, "��ʷ�ռ��û���֤", glngSys, 0, strBakUser, cnBakDB, True)
    If cnBakDB Is Nothing Then
        MsgBox "����ģʽ�ָ���ʷ�ռ��Լ����������������������ʷ�ռ䣡", vbInformation, gstrSysName
        Exit Sub
    Else
        If Val(strParallel) > 1 Then
            cnBakDB.Execute "ALTER Session FORCE PARALLEL DDL PARALLEL " & strParallel
        Else
            cnBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
        End If
    End If
    
    '�����ϵͳ���ж�
    If mlngPeisSys > 0 Then
        blnNoData = cboBakspace(1).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPeisSys, cboBakspace(1).ItemData(cboBakspace(1).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
        If blnNoData = True Then
            MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ�����ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strPeisBakUser = rsTmp!������
        
        If strBakUser = strPeisBakUser Then
            Set cnPeisBakDB = cnBakDB
        Else
            Call zlDatabase.UserIdentify(Me, "�����ϵͳ��ʷ�ռ��û���֤", mlngPeisSys, 0, strPeisBakUser, cnPeisBakDB, True)
            If cnPeisBakDB Is Nothing Then
                MsgBox "����ģʽ�ָ������ϵͳ��ʷ�ռ��Լ�������������������������ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
                Exit Sub
            Else
                If Val(strParallel) > 1 Then
                    cnPeisBakDB.Execute "ALTER Session FORCE PARALLEL DDL PARALLEL " & strParallel
                Else
                    cnPeisBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
                End If
            End If
        End If
    End If
    
    '������ϵͳ���ж�
    If mlngOperSys > 0 Then
        blnNoData = cboBakspace(2).ListCount = 0
        If blnNoData = False Then
            gstrSQL = "Select ������ From zlBakSpaces Where ϵͳ = [1] And ��� = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOperSys, cboBakspace(2).ItemData(cboBakspace(2).ListIndex))
            blnNoData = rsTmp.RecordCount = 0
        End If
        If blnNoData = True Then
            MsgBox "�ڱ�zlBakSpaces��δ�ҵ���ǰ������ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strOperBakUser = rsTmp!������
        
        If strBakUser = strOperBakUser Then
            Set cnOperBakDB = cnBakDB
        Else
            Call zlDatabase.UserIdentify(Me, "������ϵͳ��ʷ�ռ��û���֤", mlngOperSys, 0, strOperBakUser, cnOperBakDB, True)
            If cnOperBakDB Is Nothing Then
                MsgBox "����ģʽ�ָ�������ϵͳ��ʷ�ռ��Լ��������������������������ϵͳ��ʷ�ռ䣡", vbInformation, gstrSysName
                Exit Sub
            Else
                If Val(strParallel) > 1 Then
                    cnPeisBakDB.Execute "ALTER Session FORCE PARALLEL DDL PARALLEL " & strParallel
                Else
                    cnPeisBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
                End If
            End If
        End If
    End If
    
    
    lblPrompt.Caption = "���ڻָ���ʷ�ռ䱻���õ�Լ���������������ĵȴ� �� ��"
    Me.Refresh
    cmdRebBakSpace.Enabled = False
    Me.Enabled = False  'Ϊ��doevents
    Screen.MousePointer = 11
    datCurr = zlDatabase.Currentdate
    
    '������ʷ�ռ��������Լ������������������ֹ����ΨһԼ���ֶ��������ֶ���ͬ��˳��ͬ����������
    Call SetIndexStatus(glngSys, cnBakDB, True)
    If mlngPeisSys > 0 Then Call SetIndexStatus(mlngPeisSys, cnPeisBakDB, True)
    If mlngOperSys > 0 Then Call SetIndexStatus(mlngOperSys, cnOperBakDB, True)
    
    Call SetConstraintStatus(glngSys, cnBakDB, True)
    If mlngPeisSys > 0 Then Call SetConstraintStatus(mlngPeisSys, cnPeisBakDB, True)
    If mlngOperSys > 0 Then Call SetConstraintStatus(mlngOperSys, cnOperBakDB, True)
    
    
    'ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������)
    'ȡ��֮ǰ���õ�ǿ�Ʋ���DDL
    If Val(strParallel) > 1 Then
        Call SetNOParallel(cnBakDB, 0)
        cnBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
        
        If mlngPeisSys > 0 Then
            Call SetNOParallel(cnPeisBakDB, 0)
            cnPeisBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
        End If
        If mlngOperSys > 0 Then
            Call SetNOParallel(cnOperBakDB, 0)
            cnOperBakDB.Execute "ALTER Session DISABLE PARALLEL DDL"
        End If
    End If
    
    cnBakDB.Close
    If mlngPeisSys > 0 And cnPeisBakDB.State = adStateOpen Then cnPeisBakDB.Close
    If mlngOperSys > 0 And cnOperBakDB.State = adStateOpen Then cnOperBakDB.Close
    
    Screen.MousePointer = 0
    lblPrompt.Caption = ""
    Me.Enabled = True
    cmdRebBakSpace.Enabled = True
    
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "�ָ�������ɣ�����ʱ��" & strTime & "��", vbInformation, gstrSysName
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    Me.Enabled = True
    cmdRebBakSpace.Enabled = True
    If ErrCenter() = 1 Then
        'Resume
    End If
    lblPrompt.Caption = ""
End Sub

Private Sub cmdRebIndexForTag_Click()
'���ܣ��ؽ����ת����ѯ���������
    Dim bytRebScope As Byte
    Dim bytSpeedMode As Byte, datCurr As Date
    Dim strParallel As String, strTime As String
    Dim i As Double
            
    If MsgBox("�ò����ǳ���ʱ����ȷ��Ҫ�ؽ������ת����ѯ�������" & IIF(optRebScope_Manual(0).value, optRebScope_Manual(0).Caption, _
        IIF(optRebScope_Manual(1).value, optRebScope_Manual(1).Caption, optRebScope_Manual(2).Caption)) & "������", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
            
    bytSpeedMode = IIF(optmode_Index(0).value, 0, 1)
    bytRebScope = IIF(optRebScope_Manual(0).value, 0, IIF(optRebScope_Manual(1).value, 1, 2))
    strParallel = Val(txtParallel.Text)
        
    Screen.MousePointer = 11
    datCurr = zlDatabase.Currentdate
    cmdRebIndexForTag.Enabled = False
    
    On Error GoTo errH
    lblPrompt.Caption = "�����ؽ������ת����ѯ���衱������"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 6,1," & strParallel & "," & bytRebScope & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ��������")
    
    If mlngPeisSys > 0 Then
        lblPrompt.Caption = "�����ؽ����ϵͳ�����ת����ѯ���衱������"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 6,1," & strParallel & "," & bytRebScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ��������")
    End If
    
    If mlngOperSys > 0 Then
        lblPrompt.Caption = "�����ؽ�����ϵͳ�����ת����ѯ���衱������"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 6,1," & strParallel & "," & bytRebScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ��������")
    End If
        
    cmdRebIndexForTag.Enabled = True
    lblPrompt.Caption = ""
    Screen.MousePointer = 0
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "�ؽ�������ɣ�����ʱ��" & strTime & "��", vbInformation, gstrSysName
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    cmdRebIndexForTag.Enabled = True
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRebIndexOther_Click()
'���ܣ��ؽ�������ʷ����ת�����ϣ����˱��ת�����������������������
'       ����ת�����ֻ�ȫ�����ݺ��ջ���Щ��������ɾ�����ݵĿ��пռ䣬���������ɾ�����ݵ�Ч��
    Dim bytRebScope As Byte
    Dim bytSpeedMode As Byte, datCurr As Date
    Dim strParallel As String, strTime As String
    Dim i As Double
            
    If MsgBox("�ò����ǳ���ʱ����ȷ��Ҫ�ؽ������ת����ѯ���������������������", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
            
    bytSpeedMode = IIF(optmode_Index(0).value, 0, 1)
    bytRebScope = IIF(optRebScope_Manual(0).value, 0, IIF(optRebScope_Manual(1).value, 1, 2))
    strParallel = Val(txtParallel.Text)
        
    Screen.MousePointer = 11
    datCurr = zlDatabase.Currentdate
    cmdRebIndexOther.Enabled = False
    
    On Error GoTo errH
    lblPrompt.Caption = "�����ؽ������ת����ѯ���衱���������������"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 8,1," & strParallel & "," & bytRebScope & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ���������")
    
    If mlngPeisSys > 0 Then
        lblPrompt.Caption = "�����ؽ����ϵͳ�����ת����ѯ���衱���������������"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 8,1," & strParallel & "," & bytRebScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ���������")
    End If
    
    If mlngOperSys > 0 Then
        lblPrompt.Caption = "�����ؽ�����ϵͳ�����ת����ѯ���衱���������������"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 8,1," & strParallel & "," & bytRebScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ���������")
    End If
        
    cmdRebIndexOther.Enabled = True
    lblPrompt.Caption = ""
    Screen.MousePointer = 0
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "�ؽ�������ɣ�����ʱ��" & strTime & "��", vbInformation, gstrSysName
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    cmdRebIndexOther.Enabled = True
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRebJobTrigger_Click()
'���ܣ��ָ�ת��ǰ���õĺ�̨��ҵ�ʹ�����
    On Error GoTo errH
    
    gstrSQL = "Zl1_Datamove_Reb(100, 0, 1, 1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ָ�������")
    
    gstrSQL = "Zl1_Datamove_Reb(100, 0, 2, 1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ָ��Զ���ҵ")
       
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdRebOnline_Click()
'���ܣ��ָ����߿ռ��Լ��������
    Dim bytSpeedMode As Byte, datCurr As Date
    Dim strParallel As String, strTime As String
        
    If MsgBox("�ò����ǳ���ʱ����ȷ��Ҫ�ָ����߿ռ䱻���õ�Լ����������", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    bytSpeedMode = IIF(optmode(0).value, 0, 1)  '���ת��������ѡ����ģʽ�����ܲ�׼ȷ(��Ϊû�м�¼�ϴ�ת����ģʽ�����ת����ģʽ���ܲ�һ��)
    strParallel = Val(txtParallel.Text)
    
    Screen.MousePointer = 11
    cmdRebOnline.Enabled = False
    datCurr = zlDatabase.Currentdate
    
    On Error GoTo errH
    
    lblPrompt.Caption = "�����ؽ�����Ψһ��������"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 4,1," & strParallel & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ�����")
        
    lblPrompt.Caption = "���ڻָ���������Ψһ���������"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 3,1," & strParallel & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ�Լ��")
    
    mblnOffLineMoved = False
        
    If mlngPeisSys > 0 Then
        lblPrompt.Caption = "�����ؽ����ϵͳ�ġ���Ψһ��������"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 4,1," & strParallel & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ�����")
        
        
        lblPrompt.Caption = "���ڻָ����ϵͳ�ġ�������Ψһ���������"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 3,1," & strParallel & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ�Լ��")
    End If
     
        
    If mlngOperSys > 0 Then
        lblPrompt.Caption = "�����ؽ�����ϵͳ�ġ���Ψһ��������"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngOperSys & ", " & bytSpeedMode & ", 4,1," & strParallel & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ�����")
        
        
        lblPrompt.Caption = "���ڻָ�����ϵͳ�ġ�������Ψһ���������"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngOperSys & ", " & bytSpeedMode & ", 3,1," & strParallel & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ؽ�Լ��")
    End If
        
    cmdRebOnline.Enabled = True
    lblPrompt.Caption = ""
    Screen.MousePointer = 0
    
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "�ָ�������ɣ�����ʱ��" & strTime & "��", vbInformation, gstrSysName
        
    Exit Sub
errH:
    Screen.MousePointer = 0
    cmdRebOnline.Enabled = True
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdShrink_Click()
'���ܣ����������ļ�
    Dim strErr As String, strSQL As String
    Dim rsTmp As ADODB.Recordset, rsSize As ADODB.Recordset
    Dim cnsys As ADODB.Connection, lngBlockSize As Long, lngSumSize As Long
    Dim cmdTmp As New ADODB.Command
        
    On Error GoTo errH
    
    '���������ļ���Ҫ����DBA���ִ�У�
    If mblnDBA = False Then
        Call zlDatabase.UserIdentify(Me, "�����ļ�����", glngSys, 0, "sys", cnsys, True)
        If cnsys Is Nothing Then
            MsgBox "�����ļ�����Ҫ����sys�û����ӣ������Ը��û����ӣ�������������ʽ����������", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        Set cnsys = gcnOracle
    End If
    
    cmdShrink.Enabled = False
    Screen.MousePointer = 11
    
    gstrSQL = "select value from v$parameter where name = 'db_block_size'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open gstrSQL, cnsys, adOpenKeyset, adLockReadOnly
    lngBlockSize = Val("" & rsTmp!value)
        
    lblPrompt.Caption = "���ڲ�ѯ�������������ļ���"
    Me.Refresh
    gstrSQL = "Select File_Name,'alter database datafile ''' || Trim(File_Name) || ''' resize ' || Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024) || 'm' Cmd" & vbNewLine & _
            "From Dba_Data_Files A, (Select File_Id, Max(Block_Id + Blocks ) Hwm From Dba_Extents Group By File_Id) B" & vbNewLine & _
            "Where a.File_Id = b.File_Id(+) And a.Tablespace_Name Like 'ZL%' And" & vbNewLine & _
            "      Ceil(Blocks * " & lngBlockSize & " / 1024 / 1024) - Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024) > 0"
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open gstrSQL, cnsys, adOpenKeyset, adLockReadOnly
    
    Screen.MousePointer = 0
    lblPrompt.Caption = ""
    If rsTmp.RecordCount = 0 Then
        Call MsgBox("û��Ҫ���������ļ���", vbInformation, gstrSysName)
        
        cmdShrink.Enabled = True
        Exit Sub
    Else
        Set cmdTmp.ActiveConnection = cnsys
        cmdTmp.CommandType = adCmdText
        
        If MsgBox("����" & rsTmp.RecordCount & "���������������ļ�����ȷ��Ҫ������Щ�����ļ���", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            If mblnDBA = False Then cnsys.Close
            
            cmdShrink.Enabled = True
            Exit Sub
        End If
    End If
    
    '��¼����ǰ���ܴ�С
    strSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files Where Tablespace_Name Like 'ZL%'"
    Set rsSize = New ADODB.Recordset
    rsSize.Open strSQL, cnsys, adOpenKeyset, adLockReadOnly
    lngSumSize = rsSize!Mb_Size
    
    On Error Resume Next
    Screen.MousePointer = 11
    strErr = ""
    While Not rsTmp.EOF
        lblPrompt.Caption = "����������" & rsTmp!File_Name
        Me.Refresh
        DoEvents
        gstrSQL = rsTmp!cmd
        cmdTmp.CommandText = gstrSQL
        cmdTmp.Execute
        If Err.Number <> 0 Then
            strErr = strErr & vbCrLf & rsTmp!cmd & "������" & Err.Description
            Err.Clear
        End If
        
        rsTmp.MoveNext
    Wend
    
    
    '��¼��������ܴ�С
    strSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files Where Tablespace_Name Like 'ZL%'"
    Set rsSize = New ADODB.Recordset
    rsSize.Open strSQL, cnsys, adOpenKeyset, adLockReadOnly
    lngSumSize = lngSumSize - rsSize!Mb_Size
    
    If mblnDBA = False Then cnsys.Close
    
    cmdShrink.Enabled = True
    Screen.MousePointer = 0
    lblPrompt.Caption = ""
        
    If strErr <> "" Then
        MsgBox "������Ϣ��" & strErr, vbInformation, gstrSysName
    Else
        MsgBox "������ɣ���������" & lngSumSize & "M�Ŀռ䡣", vbInformation, gstrSysName
    End If
        
    Exit Sub
errH:
    cmdShrink.Enabled = True
    Screen.MousePointer = 0
    
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdMoveTable_Click()
'���ܣ�����ת�������ָ������õ������������������ļ�
'      Move��������������ģʽ����
    Dim strMoveScope As String
    Dim bytSpeedMode As Byte, datCurr As Date
    Dim strParallel As String, strTime As String, strErr As String
    Dim rsTmp As ADODB.Recordset
    
    
    If MsgBox("�ò����ǳ���ʱ����Ҫ�ж�ҵ��Ҫ�������пͻ���ͣ�õ�����½��С�" & vbCrLf & _
        "��ȷ��Ҫ����" & IIF(optMove(0).value, optMove(0).Caption, optMove(1).Caption) & "��ʷת������", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    strMoveScope = IIF(optMove(0).value, 0, 1)
    bytSpeedMode = IIF(optmode(0).value, 0, 1) '���ת��������ѡ����ģʽ�����ܲ�׼ȷ(��Ϊû�м�¼�ϴ�ת����ģʽ�����ת����ģʽ���ܲ�һ��)
    strParallel = Val(txtParallel.Text)
        
    Screen.MousePointer = 11
    datCurr = zlDatabase.Currentdate
    cmdMoveTable.Enabled = False
    
    On Error GoTo errH
    '��move�����ָ������õ�����
    lblPrompt.Caption = "����������ʷת����"
    Me.Refresh
    gstrSQL = "Zl1_Datamove_Reb(" & glngSys & ", " & bytSpeedMode & ", 7,1," & strParallel & "," & strMoveScope & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ת����")
    
    If mlngPeisSys > 0 Then
        lblPrompt.Caption = "�����������ϵͳ����ʷת����"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngPeisSys & ", " & bytSpeedMode & ", 7,1," & strParallel & "," & strMoveScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ת����")
    End If
        
    If mlngOperSys > 0 Then
        lblPrompt.Caption = "������������ϵͳ����ʷת����"
        Me.Refresh
        gstrSQL = "Zl1_Datamove_Reb(" & mlngOperSys & ", " & bytSpeedMode & ", 7,1," & strParallel & "," & strMoveScope & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ת����")
    End If
   
    lblPrompt.Caption = ""
    cmdMoveTable.Enabled = True
    Screen.MousePointer = 0
    strTime = GetTimeString(datCurr, zlDatabase.Currentdate)
    MsgBox "����������ɣ�����ʱ��" & strTime & "��", vbInformation, gstrSysName
           
            
    Exit Sub
errH:
    cmdMoveTable.Enabled = True
    Screen.MousePointer = 0
    
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then cmdHelp_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    mblnOffLineMoved = False
    mblnDBA = False
    'Dba_Role_Privs���ڰ�װ�ʹ����û�ʱ�Զ���������Ȩ��
    gstrSQL = "Select Nvl(Count(*), 0) cnt From Sys.Dba_Role_Privs Where Grantee = User And Granted_Role = 'DBA'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        mblnDBA = rsTmp!cnt > 0
    End If
        
    mstrPrivs = gstrPrivs
        
    If InStr(mstrPrivs, "����ת��") = 0 Then
        tabFunc.Tabs.Remove "����ת��"
    End If
    If InStr(mstrPrivs, "���ݳ�ѡ") = 0 Then
        tabFunc.Tabs.Remove "��ѡ����"
    End If
    For i = 1 To tabFunc.Tabs.Count
        tabFunc.Tabs(i).Caption = tabFunc.Tabs(i).Key & "(&" & i & ")"
    Next
    
    mstrPeisPrivs = ""
    mlngPeisSys = 0
    gstrSQL = "Select ��� From zlSystems Where ��� Like '21%'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.BOF = False Then
        mlngPeisSys = rsTmp("���").value
        mstrPeisPrivs = ";" & GetPrivFunc(mlngPeisSys, 2139) & ";"
    End If
    
    cmdData(4).Visible = (InStr(mstrPeisPrivs, "δת���ݲ�ѯ") > 0)
    lblData(4).Visible = (InStr(mstrPeisPrivs, "δת���ݲ�ѯ") > 0)
    Line2(4).Visible = (InStr(mstrPeisPrivs, "δת���ݲ�ѯ") > 0)
    
    mlngOperSys = 0
    gstrSQL = "Select ��� From zlSystems Where ��� Like '24%'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTmp.BOF = False Then
        mlngOperSys = rsTmp("���").value
    End If
    
    
    If Not RefreshMove Then
        Unload Me: Exit Sub
    End If
    
    
    If InStr(mstrPrivs, "���ݳ�ѡ") > 0 Then
        cboBillType.AddItem "1-�շѵ���"
        cboBillType.ItemData(cboBillType.NewIndex) = 1
        
        cboBillType.AddItem "2-���ʵ���"
        cboBillType.ItemData(cboBillType.NewIndex) = 2
        
        cboBillType.AddItem "3-�Զ�����"
        cboBillType.ItemData(cboBillType.NewIndex) = 3
        
        cboBillType.AddItem "4-�Һŵ���"
        cboBillType.ItemData(cboBillType.NewIndex) = 4
        
        cboBillType.AddItem "5-���￨"
        cboBillType.ItemData(cboBillType.NewIndex) = 5
        
        cboBillType.AddItem "6-Ԥ������"
        cboBillType.ItemData(cboBillType.NewIndex) = 6
        
        cboBillType.AddItem "7-���ʵ���"
        cboBillType.ItemData(cboBillType.NewIndex) = 7
        
        If InStr(mstrPeisPrivs, "���ݳ�ѡ") > 0 Then
            cboBillType.AddItem "8-�������"
            cboBillType.ItemData(cboBillType.NewIndex) = 8
        End If
                
        cboBillType.ListIndex = 0
        
        cboPatiType.AddItem "1-���ﲡ��"
        cboPatiType.ItemData(cboPatiType.NewIndex) = 0
        cboPatiType.AddItem "2-סԺ����"
        cboPatiType.ItemData(cboPatiType.NewIndex) = 1
        If InStr(mstrPeisPrivs, "���ݳ�ѡ") > 0 Then
            cboPatiType.AddItem "3-�ܼ���Ա"
            cboPatiType.ItemData(cboPatiType.NewIndex) = 2
            cboPatiType.AddItem "4-�ܼ�����"
            cboPatiType.ItemData(cboPatiType.NewIndex) = 3
        End If
        
        cboPatiType.ListIndex = 0
    End If
    
    Call InitLogTable
    
    Call tabFunc_Click
    Exit Sub
errH:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Function CheckDate(ByVal bytMode As Byte) As Boolean
'���ܣ����ת�����ڵ���Ч��
'������bytMode=0-ת��,1-���ת��
'      �Ա��ת����ת���ı������������������ƣ�
'      ��Ϊ�˱��⽫���ڵ����ݱ��ת�����ֽ���ҵ����˴���Ȼ���ѱ�ǵ�����ִ��ת����Ӱ��ҵ���ٴδ������ȷ�ԡ�
'      ��Ϊ�漰�ķ�Χ̫�㣬Ӧ�ó�����û�ж��ѱ��ת�������ݽ��в������ƣ���������ͨ����С����ʱ��Ϊ3���������
    Dim lngLimitDays As Long, lngDays As Long
    Dim dateCur As Date
    
    If txtDateLast.Enabled Then
        If IsNull(txtDateLast.Text) Then
            MsgBox "��������Ч�����ڡ�", vbInformation, gstrSysName
            txtDateLast.SetFocus: Exit Function
        ElseIf IsDate(txtDateLast.Text) = False Then
            MsgBox "��������Ч�����ڡ�", vbInformation, gstrSysName
            txtDateLast.SetFocus: Exit Function
        End If
    End If
    
    If IsNull(txtDateThis.Text) Then
        MsgBox "��������Ч�����ڡ�", vbInformation, gstrSysName
        txtDateThis.SetFocus: Exit Function
    ElseIf IsDate(txtDateThis.Text) = False Then
        MsgBox "��������Ч�����ڡ�", vbInformation, gstrSysName
        txtDateThis.SetFocus: Exit Function
    End If
    
    If CDate(txtDateLast.Text) < CDate(txtDateThis.Text) Then
        MsgBox "���ս�ֹ���ڲ���С�ڱ��ν�ֹ����", vbInformation, gstrSysName
        txtDateThis.SetFocus: Exit Function
    End If
    
    If CDate(txtDateThis.Text) <= CDate(txtDatePre.Text) Then
        MsgBox "���ν�ֹ����Ӧ�����ϴ�ת������", vbInformation, gstrSysName
        txtDateThis.SetFocus: Exit Function
    End If
    
    dateCur = zlDatabase.Currentdate
    lngDays = DateDiff("d", CDate(txtDateThis.Text), dateCur)
        
    If lngDays < mlngMinDays Then
        MsgBox "���ν�ֹ���ڲ���С����С���� " & Format(dateCur - mlngMinDays, "yyyy-mm-dd"), vbInformation, gstrSysName
        If txtDateThis.Enabled Then txtDateThis.SetFocus
        Exit Function
    End If
    
    If lngDays > mlngMaxDays Then
        MsgBox "���ν�ֹ���ڲ��ܴ���������� " & Format(dateCur - mlngMaxDays, "yyyy-mm-dd"), vbInformation, gstrSysName
        If txtDateThis.Enabled Then txtDateThis.SetFocus
        Exit Function
    End If
    
    
    lngLimitDays = IIF(bytMode = 0, 365, 365 * 2)   '���ת��̫����ʱ�䣬�����ʵ��ת�������׵�����Щ�����ڱ�Ǻ󱻸ı�
    If lngDays < lngLimitDays Then
        MsgBox IIF(bytMode = 0, "ת��������", "���ת��������") & "Ҫ�����߿���뱣������" & lngLimitDays & " ������ݡ�" & vbCrLf & _
                "�����������㣬���ܽ���" & IIF(bytMode = 0, "ת��������", "���ת��������"), vbInformation, gstrSysName
        Exit Function
    End If
    CheckDate = True
End Function

Private Function CheckData() As Boolean
'���ܣ�������ص������߼����
    Dim strMsg As String, i As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    CheckData = True
    
    '1.��������Ĳ��ж�
    strMsg = ""
    strSQL = "Select Index_Name, Degree" & vbNewLine & _
            "From All_Indexes" & vbNewLine & _
            "Where Degree Not In ('0', '1') And Owner = Zl_Owner And Table_Name In (Select ���� From zlBakTables)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        strMsg = strMsg & "," & rsTmp!Index_Name & "(" & Trim(rsTmp!degree) & ")"
        rsTmp.MoveNext
    Next
    If strMsg <> "" Then
        If MsgBox("�������������˲��жȣ�" & vbCrLf & Mid(strMsg, 2) & _
            vbCrLf & vbCrLf & "�������ܵ���ִ�мƻ��������󣬽�������Ӱ����ת�����������ܣ�ǿ�ҽ���ȡ����Щ�����Ĳ��ж����ԡ�" & _
            vbCrLf & "��ȷ��Ҫ������", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckData = False
            Exit Function
        End If
    End If
    
    '2.����Ĳ��ж�
    strMsg = ""
    strSQL = "Select Table_Name, Degree" & vbNewLine & _
            "From All_Tables" & vbNewLine & _
            "Where Degree != ('         1') And Owner = Zl_Owner And Table_Name In (Select ���� From zlBakTables)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        strMsg = strMsg & "," & rsTmp!Table_name & "(" & Trim(rsTmp!degree) & ")"
        rsTmp.MoveNext
    Next
    If strMsg <> "" Then
        If MsgBox("���±������˲��жȣ�" & vbCrLf & Mid(strMsg, 2) & _
            vbCrLf & vbCrLf & "�������ܵ���ִ�мƻ��������󣬽�������Ӱ����ת�����������ܣ�ǿ�ҽ���ȡ����Щ��Ĳ��ж����ԡ�" & _
            vbCrLf & "��ȷ��Ҫ������", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckData = False
            Exit Function
        End If
    End If
    
    '3.��������洢��ռ�Ĺ淶��
    strMsg = ""
    strSQL = "Select a.Index_Name, a.Tablespace_Name" & vbNewLine & _
            "From All_Indexes A" & vbNewLine & _
            "Where a.Owner = Zl_Owner And a.Tablespace_Name Not Like 'ZL%INDEX%' And" & vbNewLine & _
            "      a.Table_Name In (Select ���� From Zltools.Zlbaktables)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        strMsg = strMsg & "," & rsTmp!Index_Name & "(" & Trim(rsTmp!Tablespace_Name) & ")"
        rsTmp.MoveNext
    Next
    If strMsg <> "" Then
        If MsgBox("��������û�а��淶�洢��ZL%INDEX��ռ䣺" & vbCrLf & Mid(strMsg, 2) & _
            vbCrLf & vbCrLf & "�⽫������Ӱ��ת�����������ܣ���������Nologging���ԣ���ǿ�ҽ��������ؽ���������ȷ�ı�ռ䡣" & _
            vbCrLf & "��ȷ��Ҫ������", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckData = False
            Exit Function
        End If
    End If
    
    
    '4.���¼��ֻ�е�ǰ�û���dba��ɫʱ�Ž���
    If mblnDBA Then
        '���delete��updateʱ�����������õ�����
        strSQL = "Select Value From V$parameter Where Name = 'skip_unusable_indexes'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rsTmp!value = "FALSE" Then
            If MsgBox("��ʷ����ת����Ҫ��Oracle��ʼ������skip_unusable_indexes�޸�ΪTRUE�����������У��Ƿ�����ò�����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                CheckData = False
                Exit Function
            Else
                gstrSQL = "alter system set skip_unusable_indexes=true"
                Call gcnOracle.Execute(gstrSQL)
            End If
        End If
    End If
    
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnOffLineMoved And optmode(1).value Then
        If MsgBox("����ģʽת����ʷ���ݺ�û�лָ����߿ռ��Լ���������������¿ͻ��˵�ҵ���޷�����ʹ�ã������´ν����ģ��ʱҲ��ǳ�������ȷ��Ҫ������", _
            vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsMovelog = Nothing
End Sub

Private Sub cmdDateLast_Click()
'���ܣ�������ѡ����
        If IsDate(txtDateLast.Text) Then monSel.value = CDate(txtDateLast.Text)
        
        monSel.Tag = "txtDateLast"
        monSel.Left = Me.ScaleLeft + Me.ScaleWidth - monSel.Width - 120
        monSel.Top = txtDateLast.Top + txtDateLast.Height + 30
        monSel.ZOrder
        monSel.Visible = True
        monSel.SetFocus
End Sub


Private Sub monSel_LostFocus()
    monSel.Visible = False
End Sub

Private Sub optInType_Click(Index As Integer)
    cboBillType.Enabled = Index = 0
    txtNO.Enabled = Index = 0
    cboPatiType.Enabled = Index = 1
    txtPati.Enabled = Index = 1
    
    cboBillType.BackColor = IIF(cboBillType.Enabled, txtDateThis.BackColor, Me.BackColor)
    txtNO.BackColor = IIF(txtNO.Enabled, txtDateThis.BackColor, Me.BackColor)
    cboPatiType.BackColor = IIF(cboPatiType.Enabled, txtDateThis.BackColor, Me.BackColor)
    txtPati.BackColor = IIF(txtPati.Enabled, txtDateThis.BackColor, Me.BackColor)
    
    If cboBillType.Enabled Then cboBillType.SetFocus
    If cboPatiType.Enabled Then cboPatiType.SetFocus
    
End Sub

Private Sub optMode_Click(Index As Integer)
    If Index = 1 Then chkBakTbsDisable.value = 1    '����ģʽʱ�̶�����
    chkBakTbsDisable.Enabled = Index = 0
    chkjob.value = Index
    chkTrigger.value = Index
End Sub

Private Sub optmode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    If Index = 0 Then
        strTip = "Ϊ��������ܣ�ת���ڼ佫�������������������" & vbCrLf & _
                "1.����ת�����������������(���磺ҩƷǩ����ϸ_FK_�շ�ID)" & vbCrLf & _
                "2.���÷�ת��������,��ת��������ϵ�����(���磺����ҽ���Ƽ�_IX_�շ�ϸĿID)" & vbCrLf & _
                "3.����ɾ�������ͣ���ڼ���Զ���������������������ҵ�����ʱ�Զ�ɾ���ӱ����ݡ�"
    Else
        strTip = "Ϊ��������ܣ�ת���ڼ���˽�������Լ������������Ҫ���ã�" & vbCrLf & _
                "1.��ʷ���ݿռ������Լ��������;" & vbCrLf & _
                "2.ת�����������Ψһ��(������,���������ת����ѯ���������)" & vbCrLf & _
                "3.ת����������������������ת����ѯ�����������"
    End If

    Call zlCommFun.ShowTipInfo(optmode(Index).hwnd, strTip, True)
End Sub

Private Sub txtDateLast_GotFocus()
    Call zlControl.TxtSelAll(txtDateLast)
End Sub
Private Sub txtDateThis_GotFocus()
    Call zlControl.TxtSelAll(txtDateThis)
End Sub

Private Sub txtDateLast_KeyPress(KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txtDateThis_KeyPress(KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDateLast_Validate(Cancel As Boolean)
    If IsNull(txtDateLast.Text) Then
        Cancel = True
        Exit Sub
    ElseIf IsDate(txtDateLast.Text) = False Then
        Cancel = True
        Exit Sub
    Else
        If CheckLessBegin(txtDateLast) Or CheckLessThis(CDate(txtDateLast.Text), CDate(txtDateThis.Text)) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtDateThis_Validate(Cancel As Boolean)
    If IsNull(txtDateThis.Text) Then
        Cancel = True
        Exit Sub
    ElseIf IsDate(txtDateThis.Text) = False Then
        Cancel = True
        Exit Sub
    Else
        If CheckLessBegin(txtDateThis) Or CheckLessThis(CDate(txtDateLast.Text), CDate(txtDateThis.Text)) Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub


Private Sub monSel_DateDblClick(ByVal DateDblClicked As Date)
    monSel.Visible = False
    
    If monSel.Tag = "txtDateLast" Then
        If CheckLessBegin(txtDateLast, monSel.value) = False And CheckLessThis(monSel.value, CDate(txtDateThis.Text)) = False Then
            txtDateLast.Text = Format(monSel.value, "YYYY-MM-DD")
        End If
        
        If txtDateLast.Enabled And txtDateLast.Visible Then txtDateLast.SetFocus
    Else
        If CheckLessBegin(txtDateThis, monSel.value) = False And CheckLessThis(CDate(txtDateLast.Text), monSel.value) = False Then
            txtDateThis.Text = Format(monSel.value, "YYYY-MM-DD")
        End If
        
        If txtDateThis.Enabled And txtDateThis.Visible Then txtDateThis.SetFocus
    End If
End Sub

Private Function CheckLessBegin(objText As TextBox, Optional ByVal dateTemp As Date) As Boolean
'���ܣ����ָ���ؼ��������Ƿ�С���ϴ���ֹ����
        
    If dateTemp = CDate(0) Then dateTemp = CDate(objText.Text)
        
    If dateTemp < mdatBegin Then
        Call FS.ShowTipInfo(objText.hwnd, "����С���ϴ�ת������ֹ����:" & Format(mdatBegin, "YYYY-MM-DD"))
        CheckLessBegin = True
    Else
        Call FS.ShowTipInfo(objText.hwnd, "")
    End If
End Function

Private Function CheckLessThis(dateLast As Date, dateThis As Date) As Boolean
'���ܣ�������ս�ֹ�����Ƿ�С�ڱ��ν�ֹ����
        
    If dateLast < dateThis Then
        Call FS.ShowTipInfo(txtDateThis.hwnd, "���ս�ֹ���ڲ���С�ڱ��ν�ֹ����")
        CheckLessThis = True
    Else
        Call FS.ShowTipInfo(txtDateThis.hwnd, "")
    End If
End Function

Private Sub txtParallel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "��������Լ���ؽ�������������ʱ���õĲ��жȡ� " & vbCrLf & _
            "�����CPU�������洢�豸���������ָ����ͨ�����ڵ�һ�洢�豸���������ޣ�����Խ��Խ�á�"
    Call zlCommFun.ShowTipInfo(txtParallel.hwnd, strTip, True)
End Sub

Private Sub txtSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "ת��������Խ�࣬��ѯ�����ݷ�Χ��Խ������������ڴ�ҲԽ�󣬲����ɾ���ļ�¼��ҲԽ�࣬��Undo��Temp��ռ估��־����������Խ��" & vbCrLf & _
            "�����ҵ��������������ȷ������ͬʱ����ò�ͬ������"
    Call zlCommFun.ShowTipInfo(txtSplit.hwnd, strTip, True)
End Sub


Private Sub tabFunc_Click()
    Dim i As Long
    
    For i = 0 To fraFunc.UBound
        If fraFunc(i).Tag = tabFunc.SelectedItem.Key Then
            fraFunc(i).Visible = True
        Else
            fraFunc(i).Visible = False
        End If
    Next
    Set imgInfo.Picture = img48.ListImages(tabFunc.SelectedItem.Key).Picture
    
    If tabFunc.SelectedItem.Key = "����ת��" Then
        lblInfo.Caption = "����ת��"
        lblNote.Caption = "    Ϊ����ϵͳ��Ч���С����ٱ����������������ؽ�������ͳ����Ϣ�ռ������߿ռ�ά����ʱ�䣬���鶨�ڽ���ʷ����ת�Ƶ���ʷ�ռ��С�"
        If Visible And txtDateThis.Enabled Then txtDateThis.SetFocus
    ElseIf tabFunc.SelectedItem.Key = "��ѡ����" Then
        lblInfo.Caption = "��ѡ����"
        lblNote.Caption = "    ��ѡĳЩ��������ݷ����������ݱ��Ա�ʵʩ��Ҫ�Ĳ���"
        If Visible Then
            If optInType(0).value Then
                optInType(0).SetFocus
            Else
                optInType(1).SetFocus
            End If
        End If
    ElseIf tabFunc.SelectedItem.Key = "δת��ѯ" Then
        lblInfo.Caption = "�޷�ת�Ƶ�����ԭ���ѯ"
        lblNote.Caption = "    �оٷ���ת��ʱ����������δת�������ݼ�¼�Ͳ���ת�Ƶ�ԭ��"
        If Visible And dtpBegin.Enabled Then
            dtpBegin.SetFocus
        End If
        If Not dtpBegin.Enabled Then
            MsgBox "���ڻ�δִ�й�����ת�ƣ����ܶ��޷�ת�Ƶ�����ԭ����в�ѯ��", vbInformation, gstrSysName
        End If
    ElseIf tabFunc.SelectedItem.Key = "ת����־" Then
        lblInfo.Caption = "ת�Ʋ�����־"
        lblNote.Caption = "    �鿴ÿ��ת��������ʱ��Σ��Լ�ת�������ĺ�ʱ(��λ������)"
    
        Call RefreshMoveLog
    
    ElseIf tabFunc.SelectedItem.Key = "ת����" Then
        lblInfo.Caption = "ת����"
        lblNote.Caption = "    ת��ȫ����ɺ�,��Ҫ�˹�ִ�еĲ������ָ���ʷ�ռ䱻���õ�Լ�����������ָ����߿ռ䱻���õ�Լ�����������ָ�ת��ǰ���õĺ�̨��ҵ�ʹ�����"
    
        If txtParallel.Enabled And txtParallel.Visible Then txtParallel.SetFocus
    End If
End Sub

Private Sub TIMStatus_Timer()
'ˢ�½���
    Dim strStatus As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select To_Char(��ֹʱ��, 'yyyy-mm-dd') ��ֹʱ��, ��ǰ����" & vbNewLine & _
            "From Zldatamovelog" & vbNewLine & _
            "Where ���� = (Select Max(����) From Zldatamovelog)"
    On Error Resume Next
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        If TIMStatus.Tag = "ת��" Then
            strStatus = "����ת��" & rsTmp!��ֹʱ�� & "֮ǰ�����ݣ���ǰ���ȣ�" & rsTmp!��ǰ����
        ElseIf TIMStatus.Tag = "���ת��" Then
            strStatus = "���ڱ��" & rsTmp!��ֹʱ�� & "֮ǰ�����ݣ���ǰ���ȣ�" & rsTmp!��ǰ����
        Else
            strStatus = "��ǰ���ȣ�" & rsTmp!��ǰ����
        End If
        lblStatus.Caption = strStatus
        lblStatus.Refresh
    End If
    If Err.Number > 0 Then
        lblStatus.Caption = "ˢ�½��ȳ���:" & Err.Description
        Err.Clear
    End If
End Sub

Private Sub txtNO_GotFocus()
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    Select Case cboBillType.ItemData(cboBillType.ListIndex)
    Case 1 '�շѵ���
        txtNO.Text = GetFullNO(txtNO.Text, 13)
    Case 2, 3 '���ʵ���,�Զ�����
        txtNO.Text = GetFullNO(txtNO.Text, 14)
    Case 4 '�Һŵ���
        txtNO.Text = GetFullNO(txtNO.Text, 12)
    Case 5 '���￨
        txtNO.Text = GetFullNO(txtNO.Text, 16)
    Case 6 'Ԥ������
        txtNO.Text = GetFullNO(txtNO.Text, 11)
    Case 7 '���ʵ���
        txtNO.Text = GetFullNO(txtNO.Text, 15)
    Case 8 '������
        txtNO.Text = GetFullNO(txtNO.Text, 78)
    End Select
End Sub

Private Sub txtParallel_GotFocus()
    Call zlControl.TxtSelAll(txtParallel)
End Sub

Private Sub txtParallel_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtParallel_Validate(Cancel As Boolean)
    If Val(txtParallel.Tag) <> 0 Then
        If Val(txtParallel.Text) > Val(txtParallel.Tag) Then
            MsgBox "���жȲ��ܳ���cpu����" & txtParallel.Tag, vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub txtPati_Change()
    If txtPati.Text = "" Then
        txtPati.Tag = ""
        cboPatiType.Tag = ""
    End If
End Sub

Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    
    If Left(txtPati.Text, 1) = "." Then
        If txtPati.SelLength = 0 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
        
    If KeyAscii = 13 And txtPati.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPati.Text = txtPati.Text & Chr(KeyAscii)
            txtPati.SelStart = Len(txtPati.Text)
        End If
        KeyAscii = 0
            
        gstrSQL = Trim(txtPati.Text)
        
        Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
        Case 0, 1
            
            If Left(gstrSQL, 1) = "-" And IsNumeric(Mid(gstrSQL, 2)) Then '����ID
                gstrSQL = " And A.����ID=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "*" And IsNumeric(Mid(gstrSQL, 2)) Then '�����
                gstrSQL = " And A.�����=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "+" And IsNumeric(Mid(gstrSQL, 2)) Then 'סԺ��
                gstrSQL = " And A.סԺ��=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "." Then '�Һŵ�
                If cboPatiType.ListIndex = 0 Then
                    gstrSQL = " And B.NO='" & Mid(gstrSQL, 2) & "'"
                Else
                    gstrSQL = " And A.����ID=-1"
                End If
            Else
                gstrSQL = " And A.���� Like '" & gstrSQL & "%' And Rownum<=100"
            End If
        
            
            Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
            Case 0
                gstrSQL = _
                    " Select Rownum as ID,A.����ID,A.�����,A.����,B.NO as �Һŵ�," & _
                    " To_Char(B.�Ǽ�ʱ��,'YYYY-MM-DD') as ����ʱ��,C.���� as �������,B.ִ���� as ҽ��" & _
                    " From ������Ϣ A,H���˹Һż�¼ B,���ű� C" & _
                    " Where '%'='%' And A.����ID=B.����ID And B.ִ�в���ID=C.ID" & gstrSQL & _
                    " Order by B.�Ǽ�ʱ�� Desc"
            Case 1
                gstrSQL = _
                    " Select Rownum as ID,A.����ID,A.סԺ��,A.����,B.��ҳID as סԺ����," & _
                    " C.���� as סԺ����,To_Char(B.��Ժ����,'YYYY-MM-DD')||'��'||To_Char(B.��Ժ����,'YYYY-MM-DD') as סԺ�ڼ�" & _
                    " From ������Ϣ A,������ҳ B,���ű� C" & _
                    " Where '%'='%' And B.����ת��=1 And A.����ID=B.����ID And B.��Ժ����ID=C.ID" & gstrSQL & _
                    " Order by B.��Ժ���� Desc"
            End Select
        Case 2
            If Left(gstrSQL, 1) = "-" And IsNumeric(Mid(gstrSQL, 2)) Then '����ID
                gstrSQL = " And A.����ID=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "*" And IsNumeric(Mid(gstrSQL, 2)) Then '�����
                gstrSQL = " And A.�����=" & Val(Mid(gstrSQL, 2))
            ElseIf Left(gstrSQL, 1) = "+" And IsNumeric(Mid(gstrSQL, 2)) Then 'סԺ��
                gstrSQL = " And A.������='" & Trim(Mid(gstrSQL, 2)) & "'"
            Else
                gstrSQL = " And A.���� Like '" & gstrSQL & "%' And Rownum<=100"
            End If
            gstrSQL = _
                " Select Rownum as ID,A.����ID,A.�����,A.����,A.������" & _
                " From ������Ϣ A,�����ԱĿ¼ B" & _
                " Where '%'='%' And A.����ID=B.����ID " & gstrSQL & _
                " Order by B.����ʱ�� Desc"
        Case 3
            
            If Left(gstrSQL, 1) = "-" And IsNumeric(Mid(gstrSQL, 2)) Then '����id
                gstrSQL = " And A.ID=" & Val(Mid(gstrSQL, 2))
            Else
                gstrSQL = " And A.���� Like '%" & gstrSQL & "%' And Rownum<=100"
            End If
            
            gstrSQL = "Select A.ID,A.����,A.����,A.˵�� From �������Ŀ¼ A Where '%'='%' " & gstrSQL & " Order by A.����ʱ�� Desc"
                    
        End Select
        
        'gstrSQL = Replace(gstrSQL, "H���˹Һż�¼", "���˹Һż�¼")
        'gstrSQL = Replace(gstrSQL, "B.����ת��=1", "Nvl(B.����ת��,0)=0")
        
        vPoint = zlControl.GetCoordPos(txtPati.Container.hwnd, txtPati.Left, txtPati.Top)
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "���ﲡ��", , , , , , True, vPoint.X, vPoint.Y, txtPati.Height, blnCancel)
        If Not rsTmp Is Nothing Then
            Select Case cboPatiType.ItemData(cboPatiType.ListIndex)
            Case 0
                txtPati.Tag = rsTmp!����ID & "," & rsTmp!�Һŵ�
                txtPati.Text = rsTmp!���� & "," & rsTmp!����ʱ�� & "�վ���"
            Case 1
                txtPati.Tag = rsTmp!����ID & "," & rsTmp!סԺ����
                txtPati.Text = rsTmp!���� & ",��" & rsTmp!סԺ���� & "��סԺ"
            Case 2
                txtPati.Tag = rsTmp("����ID").value
                txtPati.Text = rsTmp("����").value & "," & rsTmp("������").value
            Case 3
                txtPati.Tag = rsTmp("ID").value
                txtPati.Text = rsTmp("����").value
            End Select
            
            cboPatiType.Tag = txtPati.Text
            Call zlControl.TxtSelAll(txtPati)
        Else
            If Not blnCancel Then
                MsgBox "û���ҵ����������Ĳ��ˡ�", vbInformation, gstrSysName
            End If
            txtPati.Text = "": txtPati.Tag = "": cboPatiType.Tag = ""
        End If
    ElseIf KeyAscii = 13 And txtPati.Text = "" Then
        KeyAscii = 0
    Else
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPati_Validate(Cancel As Boolean)
    If txtPati.Text <> cboPatiType.Tag Then
        Call txtPati_KeyPress(13)
    End If
End Sub

Private Sub txtSplit_GotFocus()
     Call zlControl.TxtSelAll(txtSplit)
End Sub

Private Sub txtSplit_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ���(���ò���)��
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim intType As Integer, curDate As Date
    
    If strNo = "" Then Exit Function
    
    If Len(strNo) >= 8 Then
        GetFullNO = Right(strNo, 8)
        Exit Function
    ElseIf Len(strNo) = 7 Then
        GetFullNO = zlStr.PrefixNO & strNo
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = zlStr.PrefixNO & Format(Right(strNo, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNo
    
    gstrSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intNum)
    
    If Not rsTmp.EOF Then
        intType = NVL(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        gstrSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = zlStr.PrefixNO & gstrSQL & Format(Right(strNo, 4), "0000")
    Else
        '������
        GetFullNO = zlStr.PrefixNO & Format(Right(strNo, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByNO(ByVal strNo As String, ByVal strTable As String, Optional ByVal strWhere As String) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��
'������vDate=ʱ����ʱ��εĿ�ʼʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    If strTable = "���˷��ü�¼" Then
        gstrSQL = "" & _
        "   Select NO From H������ü�¼ Where NO=[1] And Rownum<2" & IIF(strWhere <> "", " And " & strWhere, "") & _
        "   Union ALL " & _
        "   Select NO From HסԺ���ü�¼ Where NO=[1] And Rownum<2" & IIF(strWhere <> "", " And " & strWhere, "")
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, cboBillType.ListIndex + 1)
    Else
        gstrSQL = "Select NO From H" & strTable & " Where NO=[1] And Rownum<2" & IIF(strWhere <> "", " And " & strWhere, "")
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, cboBillType.ListIndex + 1)
    End If
    If Not rsTmp.EOF Then
        MovedByNO = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByPati(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��ж�ָ�����˵�סԺ�����Ƿ��Ѿ�ת��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    gstrSQL = "Select ����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        MovedByPati = NVL(rsTmp!����ת��, 0) = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MovedByPeis(ByVal bytMode As Byte, ParamArray varParam() As Variant) As Boolean
    '���ܣ��ж�ָ���������������Ƿ��Ѿ�ת��
    '������bytMode  �жϷ�ʽ
    '       varParam  ����
    '���أ����������ת��������True,���򷵻�False
        
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    Select Case bytMode
    Case 1          '��������
        strSQL = "Select 1 From H��������¼ Where ������=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(varParam(0)))
    Case 2          '���ܼ���Ա
        strSQL = "Select 1 From H���������Ա Where ����id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(varParam(0)))
    Case 3          '���ܼ�����
        strSQL = "Select 1 From H��������¼ Where �������id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(varParam(0)))
    End Select
    
    MovedByPeis = (rsTmp.BOF = False)
    
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetIndexStatus(ByVal lngSys As Long, ByVal cnThis As ADODB.Connection, ByVal blnEnable As Boolean)
'����:���û��������������ú������ʷ�ռ�����ݲ����ٶ�
'     ����ʱ���ù���ִ��Ҫ����SetConstraintStatus������������Ψһ���ֶδ�����Ч��������������,ORA-14063
'����:lngSys-ϵͳ���
'     cnThis-���Ӷ���
'     blnEnable-���������ԣ�true-�������� false -��������

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
    Dim strErr As String, i As Long

    '���ڹ����Ż��ӿ�SQLִ��
    If blnEnable Then
        strSQL = "Select /*+ rule*/" & vbNewLine & _
                " 'alter index ' || a.Index_Name || ' Rebuild Nologging' Sql,a.Index_Name" & vbNewLine & _
                "From User_Indexes A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Table_Name = t.���� And t.ϵͳ = " & lngSys & " And t.ֱ��ת�� = 1 And a.Status = 'UNUSABLE' And a.Index_Type = 'NORMAL' And" & vbNewLine & _
                "      Not Exists" & vbNewLine & _
                " (Select 1 From User_Constraints C Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U')) Order by a.Index_Name"
    Else
        strSQL = "Select /*+ rule*/" & vbNewLine & _
                " 'alter index ' || a.Index_Name || ' unusable' Sql,a.Index_Name" & vbNewLine & _
                "From User_Indexes A, Zltools.Zlbaktables T" & vbNewLine & _
                "Where a.Table_Name = t.���� And t.ϵͳ = " & lngSys & " And t.ֱ��ת�� = 1 And a.Status = 'VALID' And a.Index_Type = 'NORMAL' And Not Exists" & vbNewLine & _
                " (Select 1 From User_Constraints C Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U')) Order by a.Index_Name"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
       
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    On Error Resume Next
    While Not rsTmp.EOF
        i = i + 1
        If blnEnable Then
            DoEvents  'Ϊ��ˢ����ʾ����
            lblPrompt.Caption = "�������ã�" & rsTmp!Index_Name & "(" & i & "/" & rsTmp.RecordCount & ")"
            lblPrompt.Refresh
        Else
            lblStatus.Caption = "���ڽ��ã�" & rsTmp!Index_Name & "(" & i & "/" & rsTmp.RecordCount & ")"
        End If
        
        strSQL = rsTmp!SQL
        cmdTmp.CommandText = strSQL
        cmdTmp.Execute
        
        If Err.Number <> 0 And blnEnable Then
            '�������������ʹ�ã���ֻ�������ؽ����Ƚ���
            If InStr(Err.Description, "ORA-00054") > 0 Then
                Err.Clear
                strSQL = Replace(rsTmp!SQL, "Rebuild", "Rebuild Online")
                cmdTmp.CommandText = strSQL
                cmdTmp.Execute
            End If
            
            If Err.Number <> 0 Then
                strErr = strErr & vbCrLf & strSQL & " : " & Err.Description
                Err.Clear
            End If
        End If
        
        rsTmp.MoveNext
    Wend
    
    If strErr <> "" Then
        If Len(strErr) > 1000 Then strErr = Mid(strErr, 1, 1000) & "......"
        Call MsgBox(IIF(blnEnable, "����", "����") & "��������ʱ��������" & strErr, vbInformation, "����״̬����")
    End If
End Sub

Private Sub SetConstraintStatus(ByVal lngSys As Long, ByVal cnThis As ADODB.Connection, ByVal blnEnable As Boolean)
'����:���û����õ�Լ�������ú������ʷ�ռ�����ݲ����ٶ�
'     ����������Ψһ�����ɾ����Ӧ������
'����:lngSys-ϵͳ���
'     cnThis-���Ӷ���
'     blnEnable=true-����Լ��,false-����Լ��

    Dim strSQL As String, strErr As String, i As Long, strTbs As String
    Dim rsTmp As ADODB.Recordset, rsTbs As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    '��ʷ�ռ�û�����������Լ�������ԣ�����ȫ��������Ψһ��
    If blnEnable Then
        '���ؽ�������������Լ�����Ա��ؽ�����ʱ���ò���ִ������ʱ�䣬��������Լ��ʱҲ���Բ���novalidate��ʽ
         strSQL = "Select d.Table_Name, d.Constraint_Name, f_List2str(Cast(Collect(d.Column_Name Order By d.Position) As t_Strlist)) Colstr" & vbNewLine & _
                    "From User_Cons_Columns D," & vbNewLine & _
                    "     (Select a.Table_Name, a.Constraint_Name" & vbNewLine & _
                    "       From User_Constraints A, zlBakTables T" & vbNewLine & _
                    "       Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = " & lngSys & " And a.Status = 'DISABLED' And" & vbNewLine & _
                    "             a.Constraint_Type In ('P', 'U')) A" & vbNewLine & _
                    "Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name" & vbNewLine & _
                    "Group By d.Table_Name, d.Constraint_Name" & vbNewLine & _
                    "Order By Constraint_Name"
    Else
        strSQL = "Select " & vbNewLine & _
                " 'ALTER TABLE ' || a.Table_Name || ' disable constraint ' || a.Constraint_Name || Decode(a.Constraint_Type,'P',' Cascade drop index','U',' Cascade drop index','') Sql,a.Constraint_Name" & vbNewLine & _
                "From User_Constraints A, Zltools.Zlbaktables T, User_Tables b" & vbNewLine & _
                "Where a.Table_Name = t.���� And t.ϵͳ = " & lngSys & " And t.ֱ��ת�� = 1 And a.Status = 'ENABLED' And a.Table_Name = b.Table_Name And b.Iot_Type Is Null  Order by Constraint_Name"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
    
    If blnEnable Then
        '����ʹ�ú�IDX�ؼ��ֵ�������ռ�
        strSQL = "Select Tablespace_Name" & vbNewLine & _
                "From (Select Tablespace_Name" & vbNewLine & _
                "       From User_Indexes" & vbNewLine & _
                "       Where Rownum < 2" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Tablespace_Name" & vbNewLine & _
                "       From User_Indexes" & vbNewLine & _
                "       Where Tablespace_Name Like '%IDX%' And Rownum < 2)" & vbNewLine & _
                "Order By 1 Desc"
        
        Set rsTbs = New ADODB.Recordset
        rsTbs.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
        If rsTbs.RecordCount > 0 Then
            strTbs = " Tablespace " & rsTbs!Tablespace_Name
        End If
    End If
        
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText
    
    On Error Resume Next
    While Not rsTmp.EOF
        i = i + 1
        If blnEnable Then
            DoEvents  'Ϊ��ˢ����ʾ����
            lblPrompt.Caption = "�������ã�" & rsTmp!Constraint_Name & "(" & i & "/" & rsTmp.RecordCount & ")"
            lblPrompt.Refresh
        Else
            lblStatus.Caption = "���ڽ��ã�" & rsTmp!Constraint_Name & "(" & i & "/" & rsTmp.RecordCount & ")"
        End If
        
        If blnEnable Then
            '����������Ψһ��ʱ�������Ǳ�ɾ���˵ģ���������Ҫ��Create
            strSQL = "Create Unique Index " & rsTmp!Constraint_Name & " On " & rsTmp!Table_name & "(" & rsTmp!Colstr & ") Nologging" & strTbs
            cmdTmp.CommandText = strSQL
            cmdTmp.Execute
            
            If Err.Number <> 0 Then
                strErr = strErr & vbCrLf & strSQL & " : " & Err.Description
                Err.Clear
            End If
        
            '���Զ�����Լ���������Ĺ���
            strSQL = "Alter Table " & rsTmp!Table_name & " Enable Novalidate Constraint " & rsTmp!Constraint_Name
            cmdTmp.CommandText = strSQL
            cmdTmp.Execute
            
            If Err.Number <> 0 Then
                strErr = strErr & vbCrLf & strSQL & " : " & Err.Description
                Err.Clear
            End If
        Else
            strSQL = rsTmp!SQL
            cmdTmp.CommandText = strSQL
            cmdTmp.Execute
            
            If Err.Number <> 0 Then
                strErr = strErr & vbCrLf & strSQL & " : " & Err.Description
                Err.Clear
            End If
        End If
        
        rsTmp.MoveNext
    Wend
    
    If strErr <> "" Then
        If Len(strErr) > 1000 Then strErr = Mid(strErr, 1, 1000) & "......"
        Call MsgBox(IIF(blnEnable, "����", "����") & "����Լ��ʱ��������" & strErr, vbInformation, "Լ��״̬����")
    End If
End Sub

Private Sub SetNOParallel(ByVal cnThis As ADODB.Connection, ByVal bytType As Byte)
'���ܣ�����ִ�к���Զ�Ϊ�����������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������)
'������bytType��0=������1=��

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim cmdTmp As New ADODB.Command
        
    If bytType = 0 Then
        strSQL = "Select Index_Name From User_Indexes Where Degree Not In ('1', '0')"
    Else
        strSQL = "Select Table_name From User_Tables Where Degree !=('         1')"
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.Open strSQL, cnThis, adOpenKeyset, adLockReadOnly
        
    Set cmdTmp.ActiveConnection = cnThis
    cmdTmp.CommandType = adCmdText

    While Not rsTmp.EOF
        If bytType = 0 Then
            strSQL = "alter index " & rsTmp!Index_Name & " noparallel"
        Else
            strSQL = "alter table " & rsTmp!Table_name & " noparallel"
        End If
        cmdTmp.CommandText = strSQL
        cmdTmp.Execute
        
        rsTmp.MoveNext
    Wend
End Sub

Private Sub InitLogTable()
'���ܣ���ʼ�����
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "����,450,7;����,450,7;���ݿ�ʼ����,1400,0;���ݽ�������,1400,0;�ܺ�ʱ,850,7;��Ǻ�ʱ,850,7;ת����ʱ,850,7;�ؽ���ʱ,850,7"
    arrHead = Split(strHead, ";")
    
    With vsflog
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            .ColKey(.FixedCols + i) = Split(arrHead(i), ",")(0)
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub RefreshMoveLog()
'���ܣ�ˢ��ת����־
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, blnDo As Boolean
    Dim DatStart As Date, lngPre���� As Long, lng���� As Long

    On Error GoTo errH
    vsflog.Rows = vsflog.FixedRows
    
    If glngSys \ 100 = 1 Then
        gstrSQL = "Select Min(�Ǽ�ʱ��) �ϴ����� From (Select Min(�Ǽ�ʱ��) �Ǽ�ʱ�� From ������ü�¼ Union All Select Min(�Ǽ�ʱ��) From ���˹Һż�¼)"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        If IsNull(rsTmp!�ϴ�����) Then
            'δ������ҵ������
            Exit Sub
        Else
            DatStart = rsTmp!�ϴ�����
        End If
    Else
        gstrSQL = "Select To_Date('2001-01-01', 'yyyy-mm-dd') as �ϴ����� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
    End If
 
    strSQL = "Select ����, ��ֹʱ��, Nvl(��Ǻ�ʱ, 0) + Nvl(ת����ʱ, 0) + Nvl(�ؽ���ʱ, 0) As �ܺ�ʱ, ��Ǻ�ʱ, ת����ʱ, �ؽ���ʱ" & vbNewLine & _
            "From (Select ����, ����, ��ֹʱ��, Round(To_Number(��ǽ���ʱ�� - ��ǿ�ʼʱ��) * 24 * 60) As ��Ǻ�ʱ," & vbNewLine & _
            "              Round(To_Number(ת������ʱ�� - Nvl(ת����ʼʱ��, ��ǽ���ʱ��)) * 24 * 60) As ת����ʱ," & vbNewLine & _
            "              Round(To_Number(�ؽ�����ʱ�� - ת������ʱ��) * 24 * 60) As �ؽ���ʱ" & vbNewLine & _
            "       From Zldatamovelog" & vbNewLine & _
            "       Where ϵͳ = [1])" & vbNewLine & _
            "Order By ����"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngSys)

    With vsflog
        .redraw = False
        .Rows = .FixedRows + rsTmp.RecordCount
        .MergeCells = flexMergeFree
        .MergeCol(.ColIndex("����")) = True
        
        For i = .FixedRows To .Rows - 1
        
            If lngPre���� = rsTmp!���� Then
                lng���� = lng���� + 1
            Else
                lng���� = 1
                lngPre���� = rsTmp!����
            End If
            
            .TextMatrix(i, .ColIndex("����")) = rsTmp!����
            .TextMatrix(i, .ColIndex("����")) = lng����
                        
            .TextMatrix(i, .ColIndex("���ݿ�ʼ����")) = Format(DatStart, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("���ݽ�������")) = Format(rsTmp!��ֹʱ��, "yyyy-MM-dd")
            DatStart = rsTmp!��ֹʱ��
            
            
            .TextMatrix(i, .ColIndex("�ܺ�ʱ")) = "" & rsTmp!�ܺ�ʱ
            .TextMatrix(i, .ColIndex("��Ǻ�ʱ")) = "" & rsTmp!��Ǻ�ʱ
            .TextMatrix(i, .ColIndex("ת����ʱ")) = "" & rsTmp!ת����ʱ
            .TextMatrix(i, .ColIndex("�ؽ���ʱ")) = "" & rsTmp!�ؽ���ʱ
            
            
            rsTmp.MoveNext
        Next
        
        .redraw = True
        If .Rows > .FixedRows Then
            .Row = .Rows - 1
            .TopRow = .Row
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetCommandEnable(ByVal blnEnable As Boolean)
'���ܣ��ں�ʱ�����ڼ���ý�����Ҫ���ܵ����ť
    
    cmdMoveMark.Enabled = blnEnable
    cmdMoveOut.Enabled = blnEnable
    
    cmdCancel.Enabled = blnEnable
End Sub




