VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInsSymbol 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "frmInsSymbol.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSpot 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2040
      Left            =   255
      ScaleHeight     =   2040
      ScaleWidth      =   6585
      TabIndex        =   47
      Top             =   1005
      Width           =   6585
      Begin VB.Label lblPot 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2910
         TabIndex        =   56
         Top             =   1230
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblPot 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   2910
         TabIndex        =   55
         Top             =   420
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblPot 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3375
         TabIndex        =   54
         Top             =   810
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblPot 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   2520
         TabIndex        =   53
         Top             =   810
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblPot 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1155
         TabIndex        =   52
         Top             =   435
         Width           =   330
      End
      Begin VB.Label lblPot 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   480
         TabIndex        =   51
         Top             =   435
         Width           =   330
      End
      Begin VB.Label lblPot 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1155
         TabIndex        =   50
         Top             =   1110
         Width           =   330
      End
      Begin VB.Label lblPot 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   480
         TabIndex        =   49
         Top             =   1110
         Width           =   330
      End
      Begin VB.Line Line8 
         Visible         =   0   'False
         X1              =   2520
         X2              =   3675
         Y1              =   1560
         Y2              =   405
      End
      Begin VB.Line Line7 
         Visible         =   0   'False
         X1              =   2535
         X2              =   3645
         Y1              =   435
         Y2              =   1545
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   1764
         X2              =   194
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   960
         X2              =   960
         Y1              =   155
         Y2              =   1680
      End
   End
   Begin VB.PictureBox picFormat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   90
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   48
      Top             =   3360
      Width           =   2460
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5670
      TabIndex        =   3
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4230
      TabIndex        =   2
      Top             =   4065
      Width           =   1100
   End
   Begin VB.TextBox txtChar 
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   75
      TabIndex        =   1
      Top             =   3345
      Width           =   7230
   End
   Begin VB.PictureBox picCard 
      BackColor       =   &H00FFFFFF&
      Height          =   2130
      Index           =   2
      Left            =   240
      ScaleHeight     =   2070
      ScaleWidth      =   6300
      TabIndex        =   38
      Top             =   990
      Width           =   6360
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   765
         TabIndex        =   39
         Top             =   750
         Width           =   915
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   1890
         TabIndex        =   40
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   1890
         TabIndex        =   41
         Top             =   960
         Width           =   1170
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   3
         Left            =   3225
         TabIndex        =   42
         Top             =   750
         Width           =   2220
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�վ�����(��ĩ��ͣ������)"
         Height          =   180
         Index           =   3
         Left            =   3330
         TabIndex        =   46
         Top             =   510
         Width           =   2160
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����������"
         Height          =   180
         Index           =   2
         Left            =   2010
         TabIndex        =   45
         Tag             =   "�����������"
         Top             =   1290
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ���о�����"
         Height          =   180
         Index           =   1
         Left            =   1965
         TabIndex        =   44
         Top             =   315
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   0
         Left            =   840
         TabIndex        =   43
         Top             =   525
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   1815
         X2              =   3135
         Y1              =   885
         Y2              =   885
      End
   End
   Begin VB.PictureBox picCard 
      BackColor       =   &H80000005&
      Height          =   2130
      Index           =   1
      Left            =   240
      ScaleHeight     =   2070
      ScaleWidth      =   6300
      TabIndex        =   20
      Tag             =   "������ע"
      Top             =   990
      Width           =   6360
      Begin VB.Frame fraLineRYV 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   2475
         TabIndex        =   22
         Top             =   225
         Width           =   30
      End
      Begin VB.Frame fraLineRYH 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   30
         Left            =   435
         TabIndex        =   21
         Top             =   1515
         Width           =   4065
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRY 
         Height          =   675
         Left            =   435
         TabIndex        =   23
         Top             =   1185
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   1191
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorBkg    =   16777215
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "�ڶ���ĥ��"
         Height          =   945
         Index           =   4
         Left            =   4005
         TabIndex        =   32
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "��һ��ĥ��"
         Height          =   945
         Index           =   3
         Left            =   3660
         TabIndex        =   31
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "    �����"
         Height          =   945
         Index           =   2
         Left            =   3330
         TabIndex        =   30
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmInsSymbol.frx":000C
         Height          =   945
         Index           =   1
         Left            =   2985
         TabIndex        =   29
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmInsSymbol.frx":0020
         Height          =   945
         Index           =   0
         Left            =   2670
         TabIndex        =   28
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblRYUp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   2295
         TabIndex        =   27
         Top             =   45
         Width           =   360
      End
      Begin VB.Label lblRYDn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   2295
         TabIndex        =   26
         Top             =   1905
         Width           =   360
      End
      Begin VB.Label lblRYRight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   4590
         TabIndex        =   25
         Top             =   1440
         Width           =   180
      End
      Begin VB.Label lblRYLeft 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   210
         TabIndex        =   24
         Top             =   1440
         Width           =   180
      End
   End
   Begin VB.PictureBox picCard 
      BackColor       =   &H80000005&
      Height          =   2130
      Index           =   0
      Left            =   240
      ScaleHeight     =   2070
      ScaleWidth      =   6300
      TabIndex        =   4
      Tag             =   $"frmInsSymbol.frx":0032
      Top             =   990
      Width           =   6360
      Begin VB.Frame fraLineHYH 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   30
         Left            =   405
         TabIndex        =   6
         Top             =   1500
         Width           =   5505
      End
      Begin VB.Frame fraLineHYV 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   3090
         TabIndex        =   5
         Top             =   210
         Width           =   30
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHY 
         Height          =   675
         Left            =   405
         TabIndex        =   7
         Top             =   1170
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   1191
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorBkg    =   16777215
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
      End
      Begin VB.Label lblHYLeft 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   195
         TabIndex        =   19
         Top             =   1425
         Width           =   180
      End
      Begin VB.Label lblHYRight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   5970
         TabIndex        =   18
         Top             =   1425
         Width           =   180
      End
      Begin VB.Label lblHYDn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   2910
         TabIndex        =   17
         Top             =   1890
         Width           =   360
      End
      Begin VB.Label lblHYUp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   2910
         TabIndex        =   16
         Top             =   45
         Width           =   360
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "    ������"
         Height          =   930
         Index           =   0
         Left            =   3255
         TabIndex        =   15
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "    ������"
         Height          =   930
         Index           =   1
         Left            =   3600
         TabIndex        =   14
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "      ����"
         Height          =   930
         Index           =   2
         Left            =   3945
         TabIndex        =   13
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��һǰĥ��"
         Height          =   930
         Index           =   3
         Left            =   4275
         TabIndex        =   12
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "�ڶ�ǰĥ��"
         Height          =   930
         Index           =   4
         Left            =   4620
         TabIndex        =   11
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "  ��һĥ��"
         Height          =   930
         Index           =   5
         Left            =   4965
         TabIndex        =   10
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "  �ڶ�ĥ��"
         Height          =   930
         Index           =   6
         Left            =   5310
         TabIndex        =   9
         Top             =   255
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "  ����ĥ��"
         Height          =   930
         Index           =   7
         Left            =   5655
         TabIndex        =   8
         Top             =   255
         Width           =   165
      End
   End
   Begin VB.PictureBox picFree 
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   240
      ScaleHeight     =   2130
      ScaleWidth      =   6360
      TabIndex        =   34
      Top             =   990
      Width           =   6360
      Begin VB.ComboBox cboGroup 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   0
         Width           =   3615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgFree 
         Height          =   1785
         Left            =   0
         TabIndex        =   37
         Top             =   345
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   3149
         _Version        =   393216
         Rows            =   1
         Cols            =   15
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         Caption         =   "�ַ��Ӽ�(&K)"
         Height          =   180
         Left            =   0
         TabIndex        =   35
         Top             =   60
         Width           =   990
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgChar 
      Height          =   2130
      Left            =   240
      TabIndex        =   33
      Top             =   990
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   3757
      _Version        =   393216
      Rows            =   6
      Cols            =   15
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin MSComctlLib.TabStrip tabCard 
      Height          =   3180
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   5609
      MultiRow        =   -1  'True
      TabFixedWidth   =   2646
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   $"frmInsSymbol.frx":003F
            Key             =   "������ע"
            Object.Tag             =   "������ע"
            Object.ToolTipText     =   "������ע"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������ע(&Y)"
            Key             =   "������ע"
            Object.Tag             =   "������ע"
            Object.ToolTipText     =   "������ע"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������(&P)"
            Key             =   "������"
            Object.Tag             =   "������"
            Object.ToolTipText     =   "������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��λ����(&U)"
            Key             =   "��λ����"
            Object.Tag             =   "��λ����"
            Object.ToolTipText     =   "��λ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�������(&N)"
            Key             =   "�������"
            Object.Tag             =   "�������"
            Object.ToolTipText     =   "�������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ѧ����(&M)"
            Key             =   "��ѧ����"
            Object.Tag             =   "��ѧ����"
            Object.ToolTipText     =   "��ѧ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�������(&S)"
            Key             =   "�������"
            Object.Tag             =   "�������"
            Object.ToolTipText     =   "�������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����ѡ��(&F)"
            Key             =   "����ѡ��"
            Object.Tag             =   "����ѡ��"
            Object.ToolTipText     =   "����ѡ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�¾�ʷ(&J)"
            Key             =   "�¾�ʷ"
            Object.Tag             =   "�¾�ʷ"
            Object.ToolTipText     =   "�¾�ʷ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "̥��λ��(&T)"
            Key             =   "̥��λ��"
            Object.Tag             =   "̥��λ��"
            Object.ToolTipText     =   "̥��λ��"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInsSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�¾�ʷ������ʾ
Private Const YJ���� = "��������������������"
Private Const YJ��ĸ = "���˪�����ū۫񬩬�"
Private Const YJ����1 = _
        "�ͪϪѪӪժת٪۪ݪ�" & _
        "�����������" & _
        "��������������������" & _
        "��������������������" & _
        "�ǫɫ˫ͫϫѫӫի׫�" & _
        "�ݫ߫��������" & _
        "�������������������" & _
        "��������������������" & _
        "���ìŬǬɬˬͬϬѬ�"
Private Const YJ����2 = _
        "�����������ªĪƪȪ�" & _
        "�ΪЪҪԪ֪تڪܪު�" & _
        "������������" & _
        "��������������������" & _
        "�����������������«�" & _
        "�ȫʫ̫ΫЫҫԫ֫ث�" & _
        "�ޫ���������" & _
        "��������������������" & _
        "��������������������" & _
        "�¬ĬƬȬʬ̬άЬҬ�"
        
'������ע�ַ�
Private Const RY���� = "��������������������������������������������������"
Private Const RYС���� = "����������"
Private Const RYС��ĸ = "����������"
Private Const RY����� = "����������"
Private Const RY���ĸ = "����������"
Private Const RY����� = "����������"
Private Const RY���ĸ = "����������"
Private Const RY�ҷ��� = "����������"
Private Const RY�ҷ�ĸ = "����������"
'������ע�ַ�
Private Const HY���� = "��������������������������������������������������������������������������������������������������������������������������������"
Private Const HYС���� = "����������������"
Private Const HYС��ĸ = "����������������"
Private Const HY����� = "����������������"
Private Const HY���ĸ = "����������������"
Private Const HY����� = "����������������"
Private Const HY���ĸ = "����������������"
Private Const HY�ҷ��� = "����������������"
Private Const HY�ҷ�ĸ = "����������������"

'Word�������
Private Const CON������ As String = "�����������������U���E��F�����������o�p�q�r�s�t�u���C�򡪦����n������������񡲡���㡾��������硴����塸����顺�����v�w�x�y�z�{�������������A�@"
Private Const CON��λ���� As String = "����磤����꣥����H�����멈�T�L�M�N�Q�O�J�K�P����"
Private Const CON������� As String = "����������������������������������������������������������������������������¢âĢŢƢǢȢɢʢˢ̢͢΢ϢТѢҢӢԢբ֢עآ٢ڢۢܢݢޢߢ�������������"
Private Const CON��ѧ���� As String = "�֡ԡ٣��ܡݣ����ڡۡˡ��������£��ҡӡءޡġšơǡȡɡʡߡ�͡ΡϡСѡաס̨Q�R�P�ԩ������������N�S�S�R"
Private Const CON������� As String = "�����������졨������������������������I�G�����ߩh�i�l�m�j�k�|�}�~��ᨒ�ѡ��������I�J�L�K�ΨO���ܨM��"
Private Const CONҽѧ���� As String = "������������������"

'���ݱ�ע��ɫ
Private Const M_FLAGCOLOR = &HC0E0FF

'�ڲ�����
Dim blnOK As Boolean
Private mlFontSize As Long
Private mstrInfor As String
Private mblnReturnStr As Boolean
Private mobjPic As StdPicture



Private Sub cboGroup_Click()
Dim intStart As Integer, i As Integer
    If Me.cboGroup.Visible = False Then Exit Sub
    If Me.ActiveControl.Name <> Me.cboGroup.Name Then Exit Sub
    
    intStart = 0
    For i = 0 To Me.cboGroup.ListIndex - 1
        intStart = intStart + Me.cboGroup.ItemData(i)
    Next
    
    With Me.mfgFree
        .Row = intStart \ .Cols
        .Col = intStart Mod .Cols
        .TopRow = .Row
        If .Visible Then .SetFocus
    End With
End Sub
Private Sub cmdCancel_Click()
    blnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    blnOK = True
    If txtChar.Visible Then mstrInfor = Trim(txtChar.Text)
    Unload Me
End Sub

Private Sub Form_Activate()
    Call tabCard_Click
End Sub

Private Sub lblPot_Click(Index As Integer)
    If Index >= 4 Then
        lblPot(0) = "��": lblPot(1) = "��": lblPot(2) = "��": lblPot(3) = "��"
    Else
        lblPot(4) = "��": lblPot(5) = "��": lblPot(6) = "��": lblPot(7) = "��"
    End If
    
    If lblPot(Index).Caption = "��" Then
       lblPot(Index).Caption = "��"
    Else
       lblPot(Index).Caption = "��"
    End If
    
    If picSpot.Visible Then
        Call MakeSpotPic
    End If
End Sub

Private Sub mfgChar_DblClick()
    With Me.mfgChar
        If Trim(.Text) = "" Then Exit Sub
        Me.txtChar.Text = Me.txtChar.Text + .Text
        Me.txtChar.SelStart = Len(Me.txtChar.Text)
    End With
End Sub

Private Sub mfgChar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then Call mfgChar_DblClick
End Sub

Private Sub mfgFree_DblClick()
    With Me.mfgFree
        If Trim(.Text) = "" Then Exit Sub
        Me.txtChar.Text = Me.txtChar.Text + .Text
        Me.txtChar.SelStart = Len(Me.txtChar.Text)
    End With
End Sub

Private Sub mfgFree_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then Call mfgFree_DblClick
End Sub

Private Sub mfgFree_RowColChange()
Dim intPoint As Integer, intStart As Integer
Dim i As Integer, j As Integer
    With Me.mfgFree
        intPoint = .Cols * .Row + .Col + 1
    End With
    intStart = 0
    For i = 0 To Me.cboGroup.ListCount - 1
        intStart = intStart + Me.cboGroup.ItemData(i)
        If intPoint <= intStart Then Me.cboGroup.ListIndex = i: Exit Sub
    Next
End Sub

Private Sub mshHY_Click()
    If mshHY.CellBackColor = vbWhite Then
        mshHY.CellBackColor = M_FLAGCOLOR
    Else
        mshHY.CellBackColor = vbWhite
    End If
    If mblnReturnStr Then
        txtChar.Text = MakeToothString(mshHY, 8)
        If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
    Else
        Call MakeToothPic(mshHY, 8)
    End If
End Sub

Private Sub mshHY_EnterCell()
    mshHY.CellFontBold = True
    mshHY.CellFontUnderline = True
    mshHY.CellForeColor = vbBlue
End Sub

Private Sub mshHY_GotFocus()
    mshHY_EnterCell
End Sub

Private Sub mshHY_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then mshHY_Click
End Sub

Private Sub mshHY_LeaveCell()
    mshHY.CellFontBold = False
    mshHY.CellFontUnderline = False
    mshHY.CellForeColor = mshHY.ForeColor
End Sub

Private Sub mshHY_LostFocus()
    mshHY_LeaveCell
End Sub

Private Sub mshRY_Click()
    If mshRY.CellBackColor = vbWhite Then
        mshRY.CellBackColor = M_FLAGCOLOR
    Else
        mshRY.CellBackColor = vbWhite
    End If
    If mblnReturnStr Then
        txtChar.Text = MakeToothString(mshRY, 5)
        If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
    Else
        Call MakeToothPic(mshRY, 5)
    End If
End Sub

Private Sub mshRY_EnterCell()
    mshRY.CellFontBold = True
    mshRY.CellFontUnderline = True
    mshRY.CellForeColor = vbBlue
End Sub

Private Sub mshRY_GotFocus()
    mshRY_EnterCell
End Sub

Private Sub mshRY_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then mshRY_Click
End Sub

Private Sub mshRY_LeaveCell()
    mshRY.CellFontBold = False
    mshRY.CellFontUnderline = False
    mshRY.CellForeColor = mshRY.ForeColor
End Sub

Private Sub mshRY_LostFocus()
    mshRY_LeaveCell
End Sub
Private Sub tabCard_Click()
Dim strTemp As String
Dim introw As Integer, intCol As Integer
Dim i As Integer, j As Integer

    Me.txtChar.Visible = False
    Me.picFormat.Visible = False
    Me.picCard(0).Visible = False
    Me.picCard(1).Visible = False
    Me.picCard(2).Visible = False
    Me.mfgChar.Visible = False
    Me.picFree.Visible = False
    Me.picSpot.Visible = False
    Me.txtChar.Text = ""
    Set Me.picFormat.Picture = Nothing
    Select Case Me.tabCard.SelectedItem.Key
    Case "������ע"
        Me.picCard(0).Visible = True
        If mblnReturnStr Then
            Me.txtChar.Visible = True
            Me.txtChar.Text = ""
        Else
            Me.picFormat.Visible = True
            If mstrInfor <> "" Then
                If Split(mstrInfor, "|")(0) <> 2 Then '�Ѿ���ֵʱ���п�Ƭ�л�
                    mstrInfor = ""
                Else    '�༭����
                    strTemp = Split(mstrInfor, "|")(1)
                    If strTemp <> "" Then
                        For i = 0 To 7
                            If InStr(strTemp, mshHY.TextMatrix(0, i)) > 0 Then
                                mshHY.Row = 0: mshHY.Col = i
                                mshHY.CellBackColor = M_FLAGCOLOR
                            End If
                        Next
                    End If
                    strTemp = Split(mstrInfor, "|")(2)
                    If strTemp <> "" Then
                        For i = 8 To 15
                            If InStr(strTemp, mshHY.TextMatrix(0, i)) > 0 Then
                                mshHY.Row = 0: mshHY.Col = i
                                mshHY.CellBackColor = M_FLAGCOLOR
                            End If
                        Next
                    End If
                    strTemp = Split(mstrInfor, "|")(3)
                    If strTemp <> "" Then
                        For i = 0 To 7
                            If InStr(strTemp, mshHY.TextMatrix(1, i)) > 0 Then
                                mshHY.Row = 1: mshHY.Col = i
                                mshHY.CellBackColor = M_FLAGCOLOR
                            End If
                        Next
                    End If
                    strTemp = Split(mstrInfor, "|")(4)
                    If strTemp <> "" Then
                        For i = 8 To 15
                            If InStr(strTemp, mshHY.TextMatrix(1, i)) > 0 Then
                                mshHY.Row = 1: mshHY.Col = i
                                mshHY.CellBackColor = M_FLAGCOLOR
                            End If
                        Next
                    End If
                    Call MakeToothPic(mshHY, 8)
                End If
            End If
        End If
    Case "������ע"
        Me.picCard(1).Visible = True
        If mblnReturnStr Then
            Me.txtChar.Visible = True
            Me.txtChar.Text = ""
        Else
            Me.picFormat.Visible = True
            If mstrInfor <> "" Then
                If Split(mstrInfor, "|")(0) <> 3 Then '�Ѿ���ֵʱ���п�Ƭ�л�
                    mstrInfor = ""
                Else    '�༭����
                    strTemp = Split(mstrInfor, "|")(1)
                    If strTemp <> "" Then
                        For i = 0 To 4
                            If InStr(strTemp, mshRY.TextMatrix(0, i)) > 0 Then
                                mshRY.Row = 0: mshRY.Col = i
                                mshRY.CellBackColor = M_FLAGCOLOR
                            End If
                        Next
                    End If
                    strTemp = Split(mstrInfor, "|")(2)
                    If strTemp <> "" Then
                        For i = 5 To 9
                            If InStr(strTemp, mshRY.TextMatrix(0, i)) > 0 Then
                                mshRY.Row = 0: mshRY.Col = i
                                mshRY.CellBackColor = M_FLAGCOLOR
                            End If
                        Next
                    End If
                    strTemp = Split(mstrInfor, "|")(3)
                    If strTemp <> "" Then
                        For i = 0 To 4
                            If InStr(strTemp, mshRY.TextMatrix(1, i)) > 0 Then
                                mshRY.Row = 1: mshRY.Col = i
                                mshRY.CellBackColor = M_FLAGCOLOR
                            End If
                        Next
                    End If
                    strTemp = Split(mstrInfor, "|")(4)
                    If strTemp <> "" Then
                        For i = 5 To 9
                            If InStr(strTemp, mshRY.TextMatrix(1, i)) > 0 Then
                                mshRY.Row = 1: mshRY.Col = i
                                mshRY.CellBackColor = M_FLAGCOLOR
                            End If
                        Next
                    End If
                    Call MakeToothPic(mshRY, 5)
                End If
            End If
        End If
    Case "������", "��λ����", "�������", "��ѧ����", "�������"
        Me.txtChar.Visible = True
        Me.mfgChar.Visible = True
        Set mobjPic = Nothing
        Select Case Me.tabCard.SelectedItem.Key
        Case "������"
            strTemp = CON������
        Case "��λ����"
            strTemp = CON��λ����
        Case "�������"
            strTemp = CON�������
        Case "��ѧ����"
            strTemp = CON��ѧ����
        Case "�������"
            strTemp = CON������� + CONҽѧ����
        End Select
        With Me.mfgChar
            .Clear
            For i = 0 To Len(strTemp) - 1
                introw = i \ .Cols: intCol = i Mod .Cols
                .TextMatrix(introw, intCol) = Mid(strTemp, i + 1, 1)
            Next
            If .Visible Then .SetFocus
        End With
    Case "����ѡ��"
        Me.txtChar.Visible = True
        Me.picFree.Visible = True
        Set mobjPic = Nothing
        If mfgFree.Visible Then Me.mfgFree.SetFocus
    Case "�¾�ʷ"
        Me.picCard(2).Visible = True
        If mblnReturnStr Then
            Me.txtChar.Visible = True
            Me.txtChar.Text = ""
        Else
            Me.picFormat.Visible = True
            If mstrInfor <> "" Then
               If Split(mstrInfor, "|")(0) <> 1 Then '�Ѿ���ֵʱ���п�Ƭ�л�
                   mstrInfor = ""
               Else    '�༭����
                   txtYJ(0).Text = Split(mstrInfor, "|")(1)
                   txtYJ(1).Text = Split(mstrInfor, "|")(2)
                   txtYJ(2).Text = Split(mstrInfor, "|")(3)
                   txtYJ(3).Text = Split(mstrInfor, "|")(4)
                   Call MakeYJPic
               End If
            End If
        End If
    Case "̥��λ��"
        Me.picFormat.Visible = True
        Me.picSpot.Visible = True
        If mstrInfor <> "" Then
           If Split(mstrInfor, "|")(0) <> 4 Then '�Ѿ���ֵʱ���п�Ƭ�л�
               mstrInfor = ""
           Else    '�༭����
               If Split(mstrInfor, "|")(1) = 1 Then lblPot(0) = "��"
               If Split(mstrInfor, "|")(1) = 2 Then lblPot(4) = "��"
               If Split(mstrInfor, "|")(2) = 1 Then lblPot(1) = "��"
               If Split(mstrInfor, "|")(2) = 2 Then lblPot(5) = "��"
               If Split(mstrInfor, "|")(3) = 1 Then lblPot(2) = "��"
               If Split(mstrInfor, "|")(3) = 2 Then lblPot(6) = "��"
               If Split(mstrInfor, "|")(4) = 1 Then lblPot(3) = "��"
               If Split(mstrInfor, "|")(4) = 2 Then lblPot(7) = "��"
               Call MakeSpotPic
           End If
        End If
    End Select
End Sub

Private Sub txtChar_Change()
    Me.cmdOK.Enabled = Me.txtChar.Text <> ""
End Sub

Private Sub txtChar_KeyPress(KeyAscii As Integer)
    If InStr("'%?&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Function MakeToothPic(objMSH As MSHFlexGrid, bytCount As Byte) As StdPicture
'���ܣ����ݺ�����ע��������ʾ������ע��ͼƬ
'��ʽΪ������|���ݡ��¾�ʷ 1|ǰ�|����|��ĸ|���|�ֺ�; ���� 2(����)/3(����)|����|����|����|����|�ֺ�; ̥��λ�� 4|�Ϸ�|�·�|��|�ҷ�|�ֺ�
Dim introw As Integer, intCol As Integer, i As Integer
Dim a As String, b As String, C As String, D As String 'A=����,B=����,C=����,D=����

    '��ABCD�ĸ�����ı�ע���,�����Ŀ�ʼ��ݺ�,��"37"
    
    
    objMSH.Redraw = False
    introw = objMSH.Row: intCol = objMSH.Col
    
    objMSH.Row = 0
    For i = 0 To bytCount - 1
        objMSH.Col = i
        If objMSH.CellBackColor = M_FLAGCOLOR Then a = a & objMSH.TextMatrix(0, i)
    Next
    For i = bytCount To bytCount * 2 - 1
        objMSH.Col = i
        If objMSH.CellBackColor = M_FLAGCOLOR Then b = b & objMSH.TextMatrix(0, i)
    Next
    
    objMSH.Row = 1
    For i = 0 To bytCount - 1
        objMSH.Col = i
        If objMSH.CellBackColor = M_FLAGCOLOR Then C = C & objMSH.TextMatrix(1, i)
    Next
    For i = bytCount To bytCount * 2 - 1
        objMSH.Col = i
        If objMSH.CellBackColor = M_FLAGCOLOR Then D = D & objMSH.TextMatrix(1, i)
    Next
    
    objMSH.Row = introw: objMSH.Col = intCol
    objMSH.Redraw = True
    
    '���ݲ�ͬ�ĸ��������������ע
Dim r As RECT, pt As POINTAPI
Dim lAW As Long, lBW As Long, lCW As Long, lDW As Long
Dim lAH As Long, lBH As Long, lCH As Long, lDH As Long
    On Error Resume Next
    
    Set picFormat.Picture = Nothing: picFormat.Cls: picFormat.Width = "2400"
    picFormat.Font.SIZE = 8: picFormat.Refresh
    If a = "" And b = "" And C = "" And D = "" Then cmdOK.Enabled = False: Exit Function
    '����������
    lAW = picFormat.TextWidth(a):   lAH = picFormat.TextHeight(a):      lBW = picFormat.TextWidth(b):       lBH = picFormat.TextHeight(b)
    lCW = picFormat.TextWidth(C):   lCH = picFormat.TextHeight(C):      lDW = picFormat.TextWidth(D):       lDH = picFormat.TextHeight(D)
    
    If a <> "" And b = "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        picFormat.Width = picFormat.ScaleX(lAW + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lAH + 1, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Bottom = r.Top + lAH: r.Left = 2: r.Right = r.Left + lAW
        DrawTextEx picFormat.hDC, a, -1, r, DT_CENTER, ByVal 0&         'д��
        MoveToEx picFormat.hDC, 4, lAH, pt  '����
        LineTo picFormat.hDC, lAW + 4, lAH
        MoveToEx picFormat.hDC, lAW + 4, 2, pt  '����
        LineTo picFormat.hDC, lAW + 4, lAH
    ElseIf a = "" And b <> "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        picFormat.Width = picFormat.ScaleX(lBW + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lBH + 1, vbPixels, vbTwips)
        picFormat.Refresh: picFormat.AutoRedraw = True
        
        r.Top = 0: r.Bottom = r.Top + lBH: r.Left = 5: r.Right = r.Left + lBW
        DrawTextEx picFormat.hDC, b, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, lBH, pt
        LineTo picFormat.hDC, lBW + 5, lBH
        MoveToEx picFormat.hDC, 2, 2, pt
        LineTo picFormat.hDC, 2, lBH
    ElseIf a = "" And b = "" And C <> "" And D = "" Then
        'ֻ�����±�ע
        picFormat.Width = picFormat.ScaleX(lCW + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lCH, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 2: r.Bottom = r.Top + lCH: r.Left = 2: r.Right = r.Left + lCW
        DrawTextEx picFormat.hDC, C, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, lCW + 5, 1
        MoveToEx picFormat.hDC, lCW + 4, 1, pt
        LineTo picFormat.hDC, lCW + 4, lCH + 4
    ElseIf a = "" And b = "" And C = "" And D <> "" Then
        'ֻ�����±�ע
        picFormat.Width = picFormat.ScaleX(lDW + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lDH, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 2: r.Bottom = r.Top + lDH: r.Left = 5: r.Right = r.Left + lDW
        DrawTextEx picFormat.hDC, D, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, lDW + 5, 1
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, 2, lDH + 4
    ElseIf a <> "" And b <> "" And C = "" And D = "" Then
        'ֻ���������б�ע
        picFormat.Width = picFormat.ScaleX(lAW + lBW + 9, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lAH + 1, vbPixels, vbTwips)
        picFormat.Refresh
        
         r.Bottom = r.Top + lAH: r.Left = 2: r.Right = r.Left + lAW
        DrawTextEx picFormat.hDC, a, -1, r, DT_CENTER, ByVal 0&  'д��
         r.Bottom = r.Top + lAH: r.Left = r.Right + 5: r.Right = r.Left + lBW
        DrawTextEx picFormat.hDC, b, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, lAH, pt
        LineTo picFormat.hDC, lAW + lBW + 7, lAH
        MoveToEx picFormat.hDC, lAW + 4, 2, pt
        LineTo picFormat.hDC, lAW + 4, lAH
    ElseIf a = "" And b = "" And C <> "" And D <> "" Then
        'ֻ���������б�ע
        picFormat.Width = picFormat.ScaleX(lCW + lDW + 9, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lCH + 1, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 2: r.Bottom = r.Top + lCH: r.Left = 2: r.Right = r.Left + lCW
        DrawTextEx picFormat.hDC, C, -1, r, DT_CENTER, ByVal 0&
        r.Top = 2: r.Bottom = r.Top + lCH: r.Left = r.Right + 5: r.Right = r.Left + lDW
        DrawTextEx picFormat.hDC, D, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, lCW + lDW + 7, 1
        MoveToEx picFormat.hDC, lCW + 4, 2, pt
        LineTo picFormat.hDC, lCW + 4, lCH + 3
    ElseIf a <> "" And b = "" And C <> "" And D = "" Then
        'ֻ�����������б�ע
        picFormat.Width = picFormat.ScaleX(IIf(lAW > lCW, lAW, lCW) + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lAH + lCH - 2, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 0: r.Bottom = r.Top + lAH: r.Left = 2: r.Right = r.Left + IIf(lAW > lCW, lAW, lCW)
        DrawTextEx picFormat.hDC, a, -1, r, DT_CENTER, ByVal 0&
        r.Top = r.Bottom: r.Bottom = r.Top + lCH: r.Left = 2: r.Right = r.Left + IIf(lAW > lCW, lAW, lCW)
        DrawTextEx picFormat.hDC, C, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 2, lAH - 1, pt
        LineTo picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, lAH - 1
        MoveToEx picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, 2, pt
        LineTo picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, lAH + lCH + 7
    ElseIf a = "" And b <> "" And C = "" And D <> "" Then
        'ֻ�����������б�ע
        picFormat.Width = picFormat.ScaleX(IIf(lBW > lDW, lBW, lDW) + 7, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(lBH + lDH - 2, vbPixels, vbTwips)
        picFormat.Refresh
        
        r.Top = 0: r.Bottom = r.Top + lBH: r.Left = 3: r.Right = r.Left + IIf(lBW > lDW, lBW, lDW)
        DrawTextEx picFormat.hDC, b, -1, r, DT_CENTER, ByVal 0&
        r.Top = r.Bottom: r.Bottom = r.Top + lDH: r.Left = 3: r.Right = r.Left + IIf(lBW > lDW, lBW, lDW)
        DrawTextEx picFormat.hDC, D, -1, r, DT_CENTER, ByVal 0&
        MoveToEx picFormat.hDC, 3, lBH - 1, pt
        LineTo picFormat.hDC, IIf(lBW > lDW, lBW, lDW) + 4, lBH - 1
        MoveToEx picFormat.hDC, 2, 1, pt
        LineTo picFormat.hDC, 2, lAH + lCH + 6
    Else
        '�������Ҷ��б�ע
        picFormat.Width = picFormat.ScaleX(IIf(lAW > lCW, lAW, lCW) + IIf(lBW > lDW, lBW, lDW) + 9, vbPixels, vbTwips)
        picFormat.Height = picFormat.ScaleY(IIf(lAH > lBH, lAH, lBH) + IIf(lCH > lDH, lCH, lDH) - 2, vbPixels, vbTwips)
        picFormat.Refresh
        
        If a <> "" Then
            r.Bottom = lAH: r.Left = 2: r.Right = r.Left + IIf(lAW > lCW, lAW, lCW)
            DrawTextEx picFormat.hDC, a, -1, r, DT_CENTER, ByVal 0&
        End If
        If b <> "" Then
          r.Bottom = r.Top + lBH: r.Left = IIf(lAW > lCW, lAW, lCW) + 7: r.Right = r.Left + IIf(lBW > lDW, lBW, lDW)
            DrawTextEx picFormat.hDC, b, -1, r, DT_CENTER, ByVal 0&
        End If
        If C <> "" Then
            r.Top = IIf(lAH > lBH, lAH, lBH): r.Bottom = r.Top + lCH: r.Left = 2: r.Right = r.Left + IIf(lAW > lCW, lAW, lCW)
            DrawTextEx picFormat.hDC, C, -1, r, DT_CENTER, ByVal 0&
        End If
        If D <> "" Then
            r.Top = IIf(lAH > lBH, lAH, lBH): r.Bottom = r.Top + lDH: r.Left = IIf(lAW > lCW, lAW, lCW) + 7: r.Right = r.Left + IIf(lBW > lDW, lBW, lDW)
            DrawTextEx picFormat.hDC, D, -1, r, DT_CENTER, ByVal 0&
        End If
        
        MoveToEx picFormat.hDC, 2, IIf(lAH > lBH, lAH, lBH) - 1, pt
        LineTo picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + IIf(lBW > lDW, lBW, lDW) + 7, IIf(lAH > lBH, lAH, lBH) - 1
        MoveToEx picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, 2, pt
        LineTo picFormat.hDC, IIf(lAW > lCW, lAW, lCW) + 4, IIf(lAH > lBH, lAH, lBH) + IIf(lCH > lDH, lCH, lDH)
    End If
    cmdOK.Enabled = True
    Set picFormat.Picture = picFormat.Image
    Set MakeToothPic = picFormat.Image
    Set mobjPic = picFormat.Image
    mstrInfor = IIf(bytCount = 8, 2, 3) & "|" & a & "|" & b & "|" & C & "|" & D & "|" & mlFontSize
End Function
Private Function MakeToothString(objMSH As MSHFlexGrid, bytCount As Byte) As String
    '���ܣ����ݺ�����ע��������ʾ������ע�������ַ�����
    '������objMSH=������������ע���
    '      bytCount=����������
Dim byt���� As Byte, byt��ĸ As Byte, strTemp As String
Dim introw As Integer, intCol As Integer
Dim i As Integer, j As Integer
Dim a As String, b As String, C As String, D As String 'A=����,B=����,C=����,D=����
Dim YC���� As String
Dim YCС���� As String, YCС��ĸ As String
Dim YC����� As String, YC���ĸ As String
Dim YC����� As String, YC���ĸ As String
Dim YC�ҷ��� As String, YC�ҷ�ĸ As String
        
    strTemp = ""
    If objMSH.Name = "mshHY" Then
        YC���� = HY����
        YCС���� = HYС����: YCС��ĸ = HYС��ĸ
        YC����� = HY�����: YC���ĸ = HY���ĸ
        YC����� = HY�����: YC���ĸ = HY���ĸ
        YC�ҷ��� = HY�ҷ���: YC�ҷ�ĸ = HY�ҷ�ĸ
    Else
        YC���� = RY����
        YCС���� = RYС����: YCС��ĸ = RYС��ĸ
        YC����� = RY�����: YC���ĸ = RY���ĸ
        YC����� = RY�����: YC���ĸ = RY���ĸ
        YC�ҷ��� = RY�ҷ���: YC�ҷ�ĸ = RY�ҷ�ĸ
    End If
            
    '��ABCD�ĸ�����ı�ע���,�����Ŀ�ʼ��ݺ�,��"37"
    objMSH.Redraw = False
    introw = objMSH.Row: intCol = objMSH.Col
    
    objMSH.Row = 0
    For i = bytCount To 1 Step -1
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then a = a & bytCount + 1 - i
    Next
    For i = bytCount + 1 To bytCount * 2
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then b = b & i - bytCount
    Next
    
    objMSH.Row = 1
    For i = bytCount To 1 Step -1
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then C = C & bytCount + 1 - i
    Next
    For i = bytCount + 1 To bytCount * 2
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then D = D & i - bytCount
    Next
    
    objMSH.Row = introw: objMSH.Col = intCol
    objMSH.Redraw = True
    
    '���ݲ�ͬ�ĸ��������������ע�����ַ���
    If a <> "" And b = "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        For i = Len(a) To 1 Step -1
            If i = 1 Then
                strTemp = strTemp & Mid(YC�����, CByte(Mid(a, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC�����, CByte(Mid(a, i, 1)), 1)
            End If
        Next
    ElseIf a = "" And b <> "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        For i = 1 To Len(b)
            If i = 1 Then
                strTemp = strTemp & Mid(YC�ҷ���, CByte(Mid(b, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC�����, CByte(Mid(b, i, 1)), 1)
            End If
        Next
    ElseIf a = "" And b = "" And C <> "" And D = "" Then
        'ֻ�����±�ע
        For i = Len(C) To 1 Step -1
            If i = 1 Then
                strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
            End If
        Next
    ElseIf a = "" And b = "" And C = "" And D <> "" Then
        'ֻ�����±�ע
        For i = 1 To Len(D)
            If i = 1 Then
                strTemp = strTemp & Mid(YC�ҷ�ĸ, CByte(Mid(D, i, 1)), 1)
            Else
                strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(D, i, 1)), 1)
            End If
        Next
    ElseIf a <> "" And b <> "" And C = "" And D = "" Then
        'ֻ���������б�ע
        For i = Len(a) To 1 Step -1
            strTemp = strTemp & Mid(YC�����, CByte(Mid(a, i, 1)), 1)
        Next
        strTemp = strTemp & "��"
        For i = 1 To Len(b)
            strTemp = strTemp & Mid(YC�����, CByte(Mid(b, i, 1)), 1)
        Next
    ElseIf a = "" And b = "" And C <> "" And D <> "" Then
        'ֻ���������б�ע
        For i = Len(C) To 1 Step -1
            strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
        Next
        strTemp = strTemp & "��"
        For i = 1 To Len(D)
            strTemp = strTemp & Mid(YC���ĸ, CByte(Mid(D, i, 1)), 1)
        Next
    ElseIf a <> "" And b = "" And C = "" And D <> "" Then
        'ֻ�����������б�ע
        For i = Len(a) To 1 Step -1
            strTemp = strTemp & Mid(YCС����, CByte(Mid(a, i, 1)), 1)
        Next
        strTemp = strTemp & "��"
        For i = 1 To Len(D)
            strTemp = strTemp & Mid(YCС��ĸ, CByte(Mid(D, i, 1)), 1)
        Next
    ElseIf a = "" And b <> "" And C <> "" And D = "" Then
        'ֻ�����������б�ע
        For i = Len(C) To 1 Step -1
            strTemp = strTemp & Mid(YCС��ĸ, CByte(Mid(C, i, 1)), 1)
        Next
        strTemp = strTemp & "��"
        For i = 1 To Len(b)
            strTemp = strTemp & Mid(YCС����, CByte(Mid(b, i, 1)), 1)
        Next
    ElseIf Not (a = "" And b = "" And C = "" And D = "") Then
        '���¶��б�ע
        If a = "" And C = "" Then strTemp = "��"
        
        '����߷�����
        i = 1: j = 1 'i��ӦA,j��ӦC
        Do While i <= Len(a) Or j <= Len(C)
            byt���� = 0: byt��ĸ = 0
            If i <= Len(a) Then byt���� = Mid(a, i, 1)
            If j <= Len(C) Then byt��ĸ = Mid(C, j, 1)
            '���ݷ��ӷ�ĸ��һ�������������
            If byt���� <> 0 And byt��ĸ <> 0 Then
                strTemp = strTemp & Mid(YC����, (byt��ĸ - 1) * bytCount + byt����, 1)
            ElseIf byt���� <> 0 And byt��ĸ = 0 Then
                strTemp = strTemp & Mid(YCС����, byt����, 1)
            ElseIf byt���� = 0 And byt��ĸ <> 0 Then
                strTemp = strTemp & Mid(YCС��ĸ, byt��ĸ, 1)
            End If
            i = i + 1: j = j + 1
        Loop
        strTemp = StrReverse(strTemp)
        
        '���ӷ�
        If (a <> "" Or C <> "") And (b <> "" Or D <> "") Then
            strTemp = strTemp & "��"
        ElseIf b = "" And D = "" Then
            strTemp = strTemp & "��"
        End If
        
        '���ұ߷�����
        i = 1: j = 1 'i��ӦB,j��ӦD
        Do While i <= Len(b) Or j <= Len(D)
            byt���� = 0: byt��ĸ = 0
            If i <= Len(b) Then byt���� = Mid(b, i, 1)
            If j <= Len(D) Then byt��ĸ = Mid(D, j, 1)
            '���ݷ��ӷ�ĸ��һ�������������
            If byt���� <> 0 And byt��ĸ <> 0 Then
                strTemp = strTemp & Mid(YC����, (byt��ĸ - 1) * bytCount + byt����, 1)
            ElseIf byt���� <> 0 And byt��ĸ = 0 Then
                strTemp = strTemp & Mid(YCС����, byt����, 1)
            ElseIf byt���� = 0 And byt��ĸ <> 0 Then
                strTemp = strTemp & Mid(YCС��ĸ, byt��ĸ, 1)
            End If
            i = i + 1: j = j + 1
        Loop
    End If
    MakeToothString = strTemp
End Function

Public Function ShowMe(ByVal bytSex As Byte, ByVal blnReturnStr As Boolean, strInfor As String, objPic As StdPicture, lFontSize As Long) As String
    '���ܣ���ʾ���Ի���
    '������bytSex =1 �У�=2 Ů ��=0 δ֪
    '   strInfor �༭ͼƬʱ�����������Ϣ���༭���ش�
    '            ��ʽΪ������|���ݡ��¾�ʷ 1|ǰ�|����|��ĸ|���|�ֺ�; ���� 2(����)/3(����)|����|����|����|����|�ֺ�; ̥��λ�� 4|�Ϸ�|�·�|��|�ҷ�|�ֺ�
    '   objPic   �༭�������ɵ�ͼƬ�ش�
Dim intLoop As Integer
Dim introw As Integer, intCol As Integer
Dim i As Integer, j As Integer
    If lFontSize < 8 Then lFontSize = 8 'Ϊ�˱�֤ͼƬ������������岻��С����������
    mlFontSize = lFontSize * 0.9:         mstrInfor = strInfor:   mblnReturnStr = blnReturnStr
    Set mobjPic = objPic
    '������ע
    mshHY.Rows = 2: mshHY.Cols = 16
    mshHY.Height = mshHY.RowHeightMin * mshHY.Rows - 30
    mshHY.Width = mshHY.RowHeightMin * mshHY.Cols - 90
    mshHY.Left = (mshHY.Container.Width - mshHY.Width) / 2
    For i = 0 To mshHY.Cols - 1
        mshHY.ColWidth(i) = mshHY.RowHeight(0)
        mshHY.ColAlignment(i) = 4
        If i + 1 <= 8 Then
            mshHY.TextMatrix(0, i) = 8 - ((i + 1) Mod 9) + 1
            mshHY.TextMatrix(1, i) = 8 - ((i + 1) Mod 9) + 1
        Else
            mshHY.TextMatrix(0, i) = (i - 7) Mod 9
            mshHY.TextMatrix(1, i) = (i - 7) Mod 9
        End If
    Next
    fraLineHYH.Left = mshHY.Left
    fraLineHYH.Top = mshHY.Top + (mshHY.Height - fraLineHYH.Height) / 2
    fraLineHYH.Width = mshHY.Width
    fraLineHYV.Left = mshHY.Left + (mshHY.Width - fraLineHYV.Width) / 2
    
    For i = 0 To 7
        lblHY(i).Left = fraLineHYV.Left + (mshHY.ColWidth(0) - lblHY(i).Width) / 2 + mshHY.ColWidth(0) * i
    Next
    lblHYLeft.Top = fraLineHYH.Top - lblHYLeft.Height / 2
    lblHYLeft.Left = fraLineHYH.Left - lblHYLeft.Width - 60
    lblHYRight.Top = lblHYLeft.Top
    lblHYRight.Left = fraLineHYH.Left + fraLineHYH.Width + 60
    lblHYUp.Left = fraLineHYV.Left - lblHYUp.Width / 2
    lblHYUp.Top = fraLineHYV.Top - lblHYUp.Height - 30
    lblHYDn.Left = lblHYUp.Left
    lblHYDn.Top = mshHY.Top + mshHY.Height + 60
    mshHY.Row = 0: mshHY.Col = 8
    
    '������ע
    mshRY.Rows = 2: mshRY.Cols = 10
    mshRY.Height = mshRY.RowHeightMin * mshRY.Rows - 30
    mshRY.Width = mshRY.RowHeightMin * mshRY.Cols - 90
    mshRY.Left = (mshRY.Container.Width - mshRY.Width) / 2
    
    mshRY.TextMatrix(0, 0) = "��"
    mshRY.TextMatrix(0, 1) = "��"
    mshRY.TextMatrix(0, 2) = "��"
    mshRY.TextMatrix(0, 3) = "��"
    mshRY.TextMatrix(0, 4) = "��"
    For i = 0 To mshRY.Cols - 1
        mshRY.ColWidth(i) = mshRY.RowHeight(0)
        mshRY.ColAlignment(i) = 4
        
        If i >= 5 Then mshRY.TextMatrix(0, i) = mshRY.TextMatrix(0, mshRY.Cols - i - 1)
        mshRY.TextMatrix(1, i) = mshRY.TextMatrix(0, i)
    Next
    
    fraLineRYH.Left = mshRY.Left
    fraLineRYH.Top = mshRY.Top + (mshRY.Height - fraLineRYH.Height) / 2
    fraLineRYH.Width = mshRY.Width
    fraLineRYV.Left = mshRY.Left + (mshRY.Width - fraLineRYV.Width) / 2
    
    For i = 0 To 4
        lblRY(i).Left = fraLineRYV.Left + (mshRY.ColWidth(0) - lblRY(i).Width) / 2 + mshRY.ColWidth(0) * i
    Next
    lblRYLeft.Top = fraLineRYH.Top - lblRYLeft.Height / 2
    lblRYLeft.Left = fraLineRYH.Left - lblRYLeft.Width - 60
    lblRYRight.Top = lblRYLeft.Top
    lblRYRight.Left = fraLineRYH.Left + fraLineRYH.Width + 60
    lblRYUp.Left = fraLineRYV.Left - lblRYUp.Width / 2
    lblRYUp.Top = fraLineRYV.Top - lblRYUp.Height - 30
    lblRYDn.Left = lblRYUp.Left
    lblRYDn.Top = mshRY.Top + mshRY.Height + 60
    mshRY.Row = 0: mshRY.Col = 5
    
    'Word�����������
    With Me.mfgChar
        For i = 0 To .Rows - 1
            .RowHeight(i) = (.Height - 90) / .Rows
        Next
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 150) / .Cols
            .ColAlignment(i) = 4
        Next
    End With
    
    '���б�׼�ַ�
    Dim aryFree(28, 1) As String
    aryFree(0, 0) = "����������": aryFree(0, 1) = " !" & Chr(34) & "#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    aryFree(1, 0) = "������-1������": aryFree(1, 1) = "��졧��������������������������¨�������������������������������������"
    aryFree(2, 0) = "������������": aryFree(2, 1) = "����"
    aryFree(3, 0) = "���������ַ�": aryFree(3, 1) = "�����@�A�B"
    aryFree(4, 0) = "����ϣ����": aryFree(4, 1) = "���������������������������������������������������¦æĦŦƦǦȦɦʦ˦̦ͦΦϦЦѦҦӦԦզ֦צ�"
    aryFree(5, 0) = "�������": aryFree(5, 1) = "�������������������������������������������������������������������ѧҧӧԧէ֧ا٧ڧۧܧݧާߧ�������������������"
    aryFree(6, 0) = "������": aryFree(6, 1) = "�\�C���D�����������E������F��"
    aryFree(7, 0) = "���ҷ���": aryFree(7, 1) = "�"
    aryFree(8, 0) = "������ĸ�ķ���": aryFree(8, 1) = "��G�H��Y"
    aryFree(9, 0) = "������ʽ": aryFree(9, 1) = "�����������������������������������������"
    aryFree(10, 0) = "��ͷ": aryFree(10, 1) = "���������I�J�K�L"
    aryFree(11, 0) = "��ѧ�����": aryFree(11, 1) = "�ʡǡƨM�̡ءިN�ϨO�Ρġšɡȡҡӡ�ߡáˡס֡ըP�١ԡܡݨR�ڡۨ��ѡͨS"
    aryFree(12, 0) = "���Ӽ����÷���": aryFree(12, 1) = "��"
    aryFree(13, 0) = "�����ŵ���ĸ����": aryFree(13, 1) = "�٢ڢۢܢݢޢߢ���ŢƢǢȢɢʢˢ̢͢΢ϢТѢҢӢԢբ֢עآ����������������������������������¢â�"
    aryFree(14, 0) = "�Ʊ��": aryFree(14, 1) = "�������������������������������������������������������������©éĩũƩǩȩɩʩ˩̩ͩΩϩЩѩҩөԩթ֩שة٩ک۩ܩݩީߩ����������������T�U�V�W�X�Y�Z�[�\�]�^�_�`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�v�w"
    aryFree(15, 0) = "����Ԫ��": aryFree(15, 1) = "�x�y�z�{�|�}�~����������������������"
    aryFree(16, 0) = "����ͼ�η�": aryFree(16, 1) = "������������������񨍨�����"
    aryFree(17, 0) = "���Ӷ�����(ʾ�����)": aryFree(17, 1) = "�����"
    aryFree(18, 0) = "CJK���źͱ��": aryFree(18, 1) = "���������e���������������������������������������@�A�B�C�D�E�F�G�H"
    aryFree(19, 0) = "ƽ����": aryFree(19, 1) = "�������������������������������������������������������������������¤äĤŤƤǤȤɤʤˤ̤ͤΤϤФѤҤӤԤդ֤פؤ٤ڤۤܤݤޤߤ��������������������a�b�f�g"
    aryFree(20, 0) = "Ƭ����": aryFree(20, 1) = "�������������������������������������������������������������������¥åĥťƥǥȥɥʥ˥̥ͥΥϥХѥҥӥԥե֥ץ٥ڥۥܥݥޥߥ��������������������������`�c�d"
    aryFree(21, 0) = "ע��": aryFree(21, 1) = "�ŨƨǨȨɨʨ˨̨ͨΨϨШѨҨӨԨը֨רب٨ڨۨܨݨިߨ����������"
    aryFree(22, 0) = "�����ŵ�CJK��ĸ���·�": aryFree(22, 1) = "�����������Z�I"
    aryFree(23, 0) = "CJK�����ַ�": aryFree(23, 1) = "�J�K�L�M�N�O�P�Q�R�S�T"
    aryFree(24, 0) = "CJK������ʽ": aryFree(24, 1) = "�U����������������������h�i�j�k�l�m�n"
    aryFree(25, 0) = "Сд����": aryFree(25, 1) = "�o�p�q�r�s�t�u�v�w�x�y�z�{�|�}�~������������������"
    aryFree(26, 0) = "���м�ȫ���ַ�": aryFree(26, 1) = "��" & Chr(-23646) & "���磥���������������������������������������������������������£ãģţƣǣȣɣʣˣ̣ͣΣϣУѣңӣԣգ֣ףأ٣ڣۣܣݣޣߣ��������������������������������������������V���W��"
    aryFree(27, 0) = "�����ַ�": aryFree(27, 1) = "�ͪϪѪӪժת٪۪ݪߪ�������������������������������������������������ëǫɫ˫ͫϫѫӫի׫٫ݫ߫�������������������������������������������������ìŬǬɬˬͬϬѬӪ����������������������˪�����ū۫񬩬������������ªĪƪȪʪΪЪҪԪ֪تڪܪު�������������������������������������������������«īȫʫ̫ΫЫҫԫ֫ثګޫ�������������������������������������������������¬ĬƬȬʬ̬άЬҬ�"

    With Me.mfgFree
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 150 - 200) / .Cols
            .ColAlignment(i) = 4
        Next
        .RowHeight(0) = (.Height - 90) / 5
    End With
    
    introw = 0: intCol = 0
    cboGroup.Clear
    For i = 0 To UBound(aryFree) - 1
        Me.cboGroup.AddItem aryFree(i, 0)
        Me.cboGroup.ItemData(Me.cboGroup.NewIndex) = Len(aryFree(i, 1))
        For j = 0 To Len(aryFree(i, 1)) - 1
            Me.mfgFree.TextMatrix(introw, intCol) = Mid(aryFree(i, 1), j + 1, 1)
            intCol = intCol + 1
            If intCol = Me.mfgFree.Cols Then
                introw = introw + 1: intCol = 0
                If introw >= Me.mfgFree.Rows - 1 Then
                    Me.mfgFree.Rows = Me.mfgFree.Rows + 1
                    Me.mfgFree.RowHeight(Me.mfgFree.Rows - 1) = Me.mfgFree.RowHeight(0)
                End If
            End If
        Next
    Next
    Me.cboGroup.ListIndex = 0
    
    
    If bytSex = 1 Or mblnReturnStr Then
        If bytSex = 1 Then '����ʱ�����¾�ʷ
            For intLoop = 1 To Me.tabCard.Tabs.Count
                If Me.tabCard.Tabs(intLoop).Key = "�¾�ʷ" Then
                    Me.tabCard.Tabs.Remove "�¾�ʷ"
                    Exit For
                End If
            Next
        End If
        
        For intLoop = 1 To Me.tabCard.Tabs.Count '���Ի�ֻ֧���ַ�����ʱ����̥��λ��
            If Me.tabCard.Tabs(intLoop).Key = "̥��λ��" Then
                Me.tabCard.Tabs.Remove "̥��λ��"
                Exit For
            End If
        Next
    End If
    
    
    If mstrInfor <> "" Then
        Select Case Split(mstrInfor, "|")(0)
            Case 1
                Me.tabCard.Tabs.Clear
                Me.tabCard.Tabs.Add 1, "�¾�ʷ", "�¾�ʷ(&J)"
                Me.tabCard.Tabs("�¾�ʷ").Tag = "�¾�ʷ"
                Me.tabCard.Tabs("�¾�ʷ").ToolTipText = "�¾�ʷ"
            Case 2
                Me.tabCard.Tabs.Clear
                Me.tabCard.Tabs.Add 1, "������ע", "������ע(&G)"
                Me.tabCard.Tabs("������ע").Tag = "������ע"
                Me.tabCard.Tabs("������ע").ToolTipText = "������ע"
            Case 3
                Me.tabCard.Tabs.Clear
                Me.tabCard.Tabs.Add 1, "������ע", "������ע(&Y)"
                Me.tabCard.Tabs("������ע").Tag = "������ע"
                Me.tabCard.Tabs("������ע").ToolTipText = "������ע"
            Case 4
                Me.tabCard.Tabs.Clear
                Me.tabCard.Tabs.Add 1, "̥��λ��", "̥��λ��(&T)"
                Me.tabCard.Tabs("̥��λ��").Tag = "̥��λ��"
                Me.tabCard.Tabs("̥��λ��").ToolTipText = "̥��λ��"
        End Select
        Me.tabCard.Tabs(1).Selected = True
    End If
    Call tabCard_Click
    
    blnOK = False
    Me.Show vbModal
    If blnOK = False Then Exit Function
    strInfor = mstrInfor
    Set objPic = mobjPic
    ShowMe = mstrInfor
End Function

Private Sub txtYJ_Change(Index As Integer)
    If Visible Then
        If mblnReturnStr Then
            txtChar.Text = MakeYJString
            If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
        Else
            Call MakeYJPic
        End If
    End If
End Sub
Private Sub txtYJ_DblClick(Index As Integer)
    txtYJ_Change Index
End Sub

Private Sub txtYJ_GotFocus(Index As Integer)
    If mblnReturnStr Then
        txtYJ(Index).IMEMode = 3
    Else
        txtYJ(Index).IMEMode = 0
    End If
End Sub

Private Sub txtYJ_KeyPress(Index As Integer, KeyAscii As Integer)
    If mblnReturnStr Then
        If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        If InStr("|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub
Private Function MakeYJPic() As StdPicture
'��ʽΪ������|���ݡ��¾�ʷ 1|ǰ�|����|��ĸ|���|�ֺ�; ���� 2(����)/3(����)|����|����|����|����|�ֺ�; ̥��λ�� 4|�Ϸ�|�·�|��|�ҷ�|�ֺ�
Dim strB As String, strU As String, strD As String, strA As String, r As RECT, lPW As Long, lPH As Long, pt As POINTAPI
Dim lBW As Long, lBH As Long, lUW As Long, lUH As Long, lDW As Long, lDH As Long, lAW As Long, lAH As Long
    
    Set mobjPic = Nothing:                          mstrInfor = ""
    strB = txtYJ(0).Text:   strU = txtYJ(1).Text:   strD = txtYJ(2).Text:   strA = txtYJ(3).Text
    If strB <> "" And strU <> "" And strD <> "" And strA <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    
    Set picFormat.Picture = Nothing:                picFormat.Cls: picFormat.Width = "2400"
    picFormat.FontSize = 8:       picFormat.Refresh
    
    lBW = picFormat.TextWidth(strB): lBH = picFormat.TextHeight(strB): lUW = picFormat.TextWidth(strU): lUH = picFormat.TextHeight(strU)
    lDW = picFormat.TextWidth(strD): lDH = picFormat.TextHeight(strB): lAW = picFormat.TextWidth(strA): lAH = picFormat.TextHeight(strA)
    lPW = lBW + IIf(lUW > lDW, lUW, lDW) + lAW + 8
    lPH = IIf(lBH > 0, lBH, IIf(lUH > 0, lUH, IIf(lDH > 0, lDH, IIf(lAH > 0, lAH, 30)))) * 2 - 4
    picFormat.Width = picFormat.ScaleX(lPW, vbPixels, vbTwips)
    picFormat.Height = picFormat.ScaleY(lPH, vbPixels, vbTwips)
    picFormat.Refresh
    
    If strB <> "" Then
        r.Top = (lPH - lBH + 2) / 2: r.Bottom = r.Top + lBH: r.Left = 2: r.Right = r.Left + lBW
        DrawTextEx picFormat.hDC, strB, -1, r, DT_CENTER, ByVal 0&
    End If
    
    If strU <> "" Then
        r.Top = 0: r.Bottom = r.Top + lUH: r.Left = lBW + 4: r.Right = r.Left + IIf(lUW > lDW, lUW, lDW)
        DrawTextEx picFormat.hDC, strU, -1, r, DT_CENTER, ByVal 0&
    End If
    
    If strD <> "" Then
        r.Top = IIf(lUH > lDH, lUH, lDH) - 2: r.Bottom = r.Top + lDH: r.Left = lBW + 4: r.Right = r.Left + IIf(lUW > lDW, lUW, lDW)
        DrawTextEx picFormat.hDC, strD, -1, r, DT_CENTER, ByVal 0&
    End If
    
    If strA <> "" Then
        r.Top = (lPH - lAH + 2) / 2: r.Bottom = r.Top + lAH: r.Left = lBW + IIf(lUW > lDW, lUW, lDW) + 6: r.Right = r.Left + lAW
        DrawTextEx picFormat.hDC, strA, -1, r, DT_CENTER, ByVal 0&
    End If
    
    MoveToEx picFormat.hDC, lBW + 4, (lPH + 1) / 2, pt
    LineTo picFormat.hDC, lBW + IIf(lUW > lDW, lUW, lDW) + 4, (lPH + 1) / 2
    
    Set picFormat.Picture = picFormat.Image
    Set MakeYJPic = picFormat.Image
    Set mobjPic = picFormat.Image
    mstrInfor = "1|" & strB & "|" & strU & "|" & strD & "|" & strA & "|" & mlFontSize
End Function

Private Function MakeYJString() As String
'���ܣ������¾�ʷ��д���������������ַ���ע��
    Dim str���� As String, str��ĸ As String
    Dim strtmp As String
    
    If Not (IsNumeric(txtYJ(1).Text) And IsNumeric(txtYJ(2).Text)) Then Exit Function
    
    '��������֣��������Ҷ���
    '------------------------
    str���� = Right(Format(Int(txtYJ(1).Text), "00"), 2)
    str��ĸ = Right(Format(Int(txtYJ(2).Text), "00"), 2)
    
    '��10λ���ַ�
    If Val(Left(str��ĸ, 1)) <> 0 Or Val(Left(str����, 1)) <> 0 Then
        If Val(Left(str��ĸ, 1)) <> 0 And Val(Left(str����, 1)) <> 0 Then
            strtmp = Mid(YJ����1, (Val(Left(str��ĸ, 1)) - 1) * 10 + Val(Left(str����, 1)) + 1, 1)
        ElseIf Val(Left(str����, 1)) = 0 Then
            strtmp = Mid(YJ��ĸ, Val(Left(str��ĸ, 1)) + 1, 1)
        ElseIf Val(Left(str��ĸ, 1)) = 0 Then
            strtmp = Mid(YJ����, Val(Left(str����, 1)) + 1, 1)
        End If
    End If
        
    '���λ���ַ�
    strtmp = strtmp & Mid(YJ����2, Val(Right(str��ĸ, 1)) * 10 + Val(Right(str����, 1)) + 1, 1)
        
    '��������ַ�
    If IsNumeric(txtYJ(0).Text) Then
        strtmp = txtYJ(0).Text & strtmp
    End If
    If IsNumeric(txtYJ(3).Text) Or IsDate(txtYJ(3).Text) Then
        strtmp = strtmp & txtYJ(3).Text
    End If
    
    MakeYJString = strtmp
End Function
Private Function MakeSpotPic() As StdPicture
'�� ��
'���ܣ�����ѡ������̥��λ��ͼƬ,��������Ӧ��Ϣ
'��ʽΪ������|���ݡ��¾�ʷ 1|ǰ�|����|��ĸ|���|�ֺ�; ���� 2(����)/3(����)|����|����|����|����|�ֺ�; ̥��λ�� 4|�Ϸ�|�·�|��|�ҷ�|�ֺ�
Dim lPW As Long, lPH As Long, r As RECT, pt As POINTAPI, intType As Integer, lsw As Long, lsh As Long
    Set mobjPic = Nothing:                          mstrInfor = ""
    Set picFormat.Picture = Nothing:                picFormat.Cls: picFormat.Width = "2400"
    picFormat.FontSize = 8:       picFormat.Refresh
    lsw = picFormat.TextWidth("��"): lsh = picFormat.TextHeight("��")
    If lblPot(0) = "��" Or lblPot(1) = "��" Or lblPot(2) = "��" Or lblPot(3) = "��" Then
        lPW = lsw * 2 + 3
        lPH = lsh * 2
        intType = 1
        cmdOK.Enabled = True
    ElseIf lblPot(4) = "��" Or lblPot(5) = "��" Or lblPot(6) = "��" Or lblPot(7) = "��" Then
        lPW = lsw * 3 - 8
        lPH = lsh * 3 - 10
        intType = 2
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
        Exit Function
    End If
    picFormat.Width = picFormat.ScaleX(lPW, vbPixels, vbTwips)
    picFormat.Height = picFormat.ScaleY(lPH, vbPixels, vbTwips)
    picFormat.Refresh
    
Dim ba As Byte, bb As Byte, bc As Byte, bd As Byte, be As Byte, bf As Byte, bg As Byte, bh As Byte
    If lblPot(0) = "��" Then
        r.Top = 0: r.Bottom = r.Top + lsh: r.Left = 1: r.Right = r.Left + lsw: ba = 1
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(1) = "��" Then
        r.Top = 0: r.Bottom = r.Top + lsh: r.Left = lsw + 4: r.Right = r.Left + lsw: bb = 1
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(2) = "��" Then
        r.Top = lsh: r.Bottom = r.Top + lsh: r.Left = 1: r.Right = r.Left + lsw: bc = 1
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(3) = "��" Then
        r.Top = lsh: r.Bottom = r.Top + lsh: r.Left = lsw + 4: r.Right = r.Left + lsw: bd = 1
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(4) = "��" Then
        r.Top = -1: r.Bottom = r.Top + lsh: r.Left = lsw - 4: r.Right = r.Left + lsw: be = 2
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If lblPot(5) = "��" Then
        r.Top = lPH - lsh + 2: r.Bottom = r.Top + lsh: r.Left = lsw - 3: r.Right = r.Left + lsw: bf = 2
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    If lblPot(6) = "��" Then
        r.Top = lsh - 4: r.Bottom = r.Top + lsh: r.Left = -1: r.Right = r.Left + lsw: bg = 2
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    If lblPot(7) = "��" Then
        r.Top = lsh - 4: r.Bottom = r.Top + lsh: r.Left = lPW - lsw + 2: r.Right = r.Left + lsw: bh = 2
        DrawTextEx picFormat.hDC, "��", -1, r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, ByVal 0&
    End If
    
    If intType = 1 Then
        MoveToEx picFormat.hDC, 0, lsh - 1, pt
        LineTo picFormat.hDC, lPW - 1, lsh - 1
        MoveToEx picFormat.hDC, lsw + 2, 0, pt
        LineTo picFormat.hDC, lsw + 2, lPH - 1
        mstrInfor = "4|" & ba & "|" & bb & "|" & bc & "|" & bd & "|" & mlFontSize
    ElseIf intType = 2 Then
        MoveToEx picFormat.hDC, 1, 2, pt
        LineTo picFormat.hDC, lPW - 1, lPH - 1
        MoveToEx picFormat.hDC, 1, lPH - 1, pt
        LineTo picFormat.hDC, lPW - 1, 1
        mstrInfor = "4|" & be & "|" & bf & "|" & bg & "|" & bh & "|" & mlFontSize
    End If
    
    Set picFormat.Picture = picFormat.Image
    Set MakeSpotPic = picFormat.Image
    Set mobjPic = picFormat.Image
End Function

