VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageCourse 
   AutoRedraw      =   -1  'True
   Caption         =   "�����������"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9135
   Icon            =   "frmManageCourse.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picCard_s 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   5850
      MouseIcon       =   "frmManageCourse.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmManageCourse.frx":045C
      ScaleHeight     =   4110
      ScaleWidth      =   2805
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1700
      Width           =   2835
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   41
         Top             =   1725
         Width           =   105
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   240
         TabIndex        =   40
         Top             =   1725
         Width           =   420
      End
      Begin VB.Line Line30 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5030
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line29 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5045
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   255
         TabIndex        =   35
         Top             =   3870
         Width           =   420
      End
      Begin VB.Label lbl��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   34
         Top             =   3870
         Width           =   105
      End
      Begin VB.Line Line28 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5045
         Y1              =   3825
         Y2              =   3825
      End
      Begin VB.Line Line27 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5060
         Y1              =   3810
         Y2              =   3810
      End
      Begin VB.Label lblҽ�Ƹ��ʽ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   33
         Top             =   3510
         Width           =   105
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   255
         TabIndex        =   32
         Top             =   3510
         Width           =   420
      End
      Begin VB.Line Line26 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5060
         Y1              =   3435
         Y2              =   3435
      End
      Begin VB.Line Line25 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5045
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Label lblҽ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   765
         TabIndex        =   29
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   45
         TabIndex        =   28
         Top             =   1080
         Width           =   630
      End
      Begin VB.Line Line24 
         BorderColor     =   &H80000015&
         X1              =   -45
         X2              =   5045
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Line Line23 
         BorderColor     =   &H80000014&
         X1              =   -30
         X2              =   5045
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   26
         Top             =   390
         Width           =   105
      End
      Begin VB.Line Line22 
         BorderColor     =   &H80000015&
         X1              =   690
         X2              =   690
         Y1              =   330
         Y2              =   645
      End
      Begin VB.Line Line21 
         BorderColor     =   &H80000014&
         X1              =   705
         X2              =   705
         Y1              =   330
         Y2              =   660
      End
      Begin VB.Line Line20 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5000
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Line Line19 
         BorderColor     =   &H80000014&
         X1              =   15
         X2              =   5000
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line18 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   5000
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line17 
         BorderColor     =   &H80000014&
         X1              =   15
         X2              =   5000
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Line Line16 
         BorderColor     =   &H80000014&
         X1              =   -75
         X2              =   5000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line15 
         BorderColor     =   &H80000014&
         X1              =   -30
         X2              =   5000
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000014&
         X1              =   -75
         X2              =   5000
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line13 
         BorderColor     =   &H80000014&
         X1              =   -45
         X2              =   5000
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000014&
         X1              =   1440
         X2              =   1440
         Y1              =   660
         Y2              =   1005
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000014&
         X1              =   1980
         X2              =   1980
         Y1              =   660
         Y2              =   1005
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000014&
         X1              =   705
         X2              =   705
         Y1              =   1005
         Y2              =   4100
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000015&
         X1              =   690
         X2              =   690
         Y1              =   990
         Y2              =   4100
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000015&
         X1              =   1965
         X2              =   1965
         Y1              =   645
         Y2              =   990
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000015&
         X1              =   1425
         X2              =   1425
         Y1              =   645
         Y2              =   990
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000015&
         X1              =   -60
         X2              =   5000
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000015&
         X1              =   -90
         X2              =   5000
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000015&
         X1              =   -45
         X2              =   5000
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         X1              =   -90
         X2              =   5000
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         X1              =   0
         X2              =   5000
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   -15
         X2              =   5000
         Y1              =   645
         Y2              =   645
      End
      Begin VB.Label lblҽ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   25
         Top             =   3135
         Width           =   105
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   255
         TabIndex        =   24
         Top             =   3135
         Width           =   420
      End
      Begin VB.Label lbl����ȼ� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   23
         Top             =   2775
         Width           =   105
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   2775
         Width           =   420
      End
      Begin VB.Label lbl��Ժʱ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   21
         Top             =   2415
         Width           =   105
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   240
         TabIndex        =   20
         Top             =   2415
         Width           =   420
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   19
         Top             =   2085
         Width           =   105
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   240
         TabIndex        =   18
         Top             =   2055
         Width           =   420
      End
      Begin VB.Label lblסԺ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   795
         TabIndex        =   17
         Top             =   1395
         Width           =   105
      End
      Begin VB.Label lbl��ʶ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   30
         TabIndex        =   16
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2055
         TabIndex        =   15
         Top             =   735
         Width           =   420
      End
      Begin VB.Label lbl�Ա� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1485
         TabIndex        =   14
         Top             =   735
         Width           =   420
      End
      Begin VB.Label lbl���� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         TabIndex        =   13
         Top             =   735
         Width           =   1275
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   450
         TabIndex        =   12
         Top             =   60
         Width           =   570
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "�ȼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   390
         Width           =   420
      End
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3240
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   34
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1A4C
            Key             =   "M"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":2326
            Key             =   "M_Change"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":2C00
            Key             =   "KM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":34DA
            Key             =   "KM_Change"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":3DB4
            Key             =   "F"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":468E
            Key             =   "F_Change"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":4F68
            Key             =   "KF"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":5842
            Key             =   "KF_Change"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":611C
            Key             =   "O"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":69F6
            Key             =   "O_Change"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":72D0
            Key             =   "KO"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":7BAA
            Key             =   "K0_Change"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":8484
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":879E
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":8AB8
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":8DD2
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":90EC
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":9406
            Key             =   "KHolding"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":9CE0
            Key             =   "Change"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":9FFA
            Key             =   "KChange"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":A8D4
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":B1AE
            Key             =   "KOut"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":BA88
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":BDA2
            Key             =   "KFamily"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":C67C
            Key             =   "Limit"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":D4C6
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":D62C
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":D792
            Key             =   "MASK_�Ӵ�"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":DAAC
            Key             =   "MASK_�Ǳ�"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":DDC6
            Key             =   "MASK_����"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":E0E0
            Key             =   "MASK_����_�Ӵ�"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":E3FA
            Key             =   "MASK_����_�Ǳ�"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":E714
            Key             =   "U"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":F566
            Key             =   "KU"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3240
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":103B8
            Key             =   "M"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":10C92
            Key             =   "M_Change"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1156C
            Key             =   "KM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":11E46
            Key             =   "KM_Change"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":12720
            Key             =   "F"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":12FFA
            Key             =   "F_Change"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":138D4
            Key             =   "KF"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":141AE
            Key             =   "KF_Change"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":14A88
            Key             =   "O"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":15362
            Key             =   "O_Change"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":15C3C
            Key             =   "KO"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":16516
            Key             =   "KO_Change"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":16DF0
            Key             =   "M_Empty"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1710A
            Key             =   "F_Empty"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":17424
            Key             =   "Empty"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1773E
            Key             =   "Remedy"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":17A58
            Key             =   "Holding"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":17D72
            Key             =   "KHolding"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1864C
            Key             =   "Change"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":18966
            Key             =   "KChange"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":19240
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":19B1A
            Key             =   "KOut"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1A3F4
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1A70E
            Key             =   "KFamily"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1AFE8
            Key             =   "MASK_�Ӵ�"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B142
            Key             =   "MASK_�Ǳ�"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B29C
            Key             =   "MASK_����"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B3F6
            Key             =   "MASK_����_�Ӵ�"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B550
            Key             =   "MASK_����_�Ǳ�"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1B6AA
            Key             =   "U"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1C4FC
            Key             =   "KU"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timSize 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8445
      Top             =   5775
   End
   Begin MSComctlLib.Toolbar tbrFilter 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   30
      Top             =   780
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   609
      ButtonWidth     =   1984
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgFilter"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "������Ժ"
            Key             =   "curDay"
            Object.ToolTipText     =   "ֻ��ʾ������Ժ�Ĳ���(F7)"
            Object.Tag             =   "������Ժ"
            ImageKey        =   "UnCheck_"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9135
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   7635
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "����"
      Child2          =   "cboUnit"
      MinWidth2       =   2205
      MinHeight2      =   300
      Width2          =   1215
      NewRow2         =   0   'False
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   6840
         TabIndex        =   5
         Text            =   "cboUnit"
         Top             =   240
         Width           =   2205
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   165
         TabIndex        =   8
         Top             =   30
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ס"
               Key             =   "In"
               Description     =   "��ס"
               Object.ToolTipText     =   "��ס"
               Object.Tag             =   "��ס"
               ImageKey        =   "In"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ת��"
               Key             =   "Change"
               Description     =   "ת��"
               Object.ToolTipText     =   "ת��"
               Object.Tag             =   "ת��"
               ImageKey        =   "Change"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Move"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Move"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ժ"
               Key             =   "Out"
               Description     =   "��Ժ"
               Object.ToolTipText     =   "��Ժ"
               Object.Tag             =   "��Ժ"
               ImageKey        =   "Out"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Adjust"
               Description     =   "����"
               Object.ToolTipText     =   "�������˵���ݻ���Ժ��Ϣ"
               Object.Tag             =   "����"
               ImageKey        =   "Adjust"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Adjust_"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Undo"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Undo"
               Style           =   5
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�б�"
               Key             =   "View"
               Description     =   "�б�"
               Object.ToolTipText     =   "��λ�б���ʾ��ʽ"
               Object.Tag             =   "�б�"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Icon"
                     Object.Tag             =   "��ͼ��(&G)"
                     Text            =   "��ͼ��(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Object.Tag             =   "Сͼ��(&M)"
                     Text            =   "Сͼ��(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Object.Tag             =   "�б�(&L)"
                     Text            =   "�б�(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Object.Tag             =   "��ϸ����(&D)"
                     Text            =   "��ϸ����(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   6195
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageCourse.frx":1D34E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9419
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "������ɫ"
            TextSave        =   "������ɫ"
            Key             =   "PatiColor"
            Object.Tag             =   "PatiColor"
            Object.ToolTipText     =   "������ɫ˵��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picVsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   5580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4905
      ScaleWidth      =   45
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1125
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwOut_s 
      Height          =   2415
      Left            =   5670
      TabIndex        =   3
      Tag             =   "�ɱ仯��"
      Top             =   3735
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   4260
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwFamily_s 
      Height          =   2190
      Left            =   5655
      TabIndex        =   1
      Tag             =   "�ɱ仯��"
      Top             =   1275
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   3863
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwIn_s 
      Height          =   1410
      Left            =   75
      TabIndex        =   2
      Tag             =   "�ɱ仯��"
      Top             =   4740
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   2487
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwBeds_s 
      Height          =   3210
      Left            =   60
      TabIndex        =   0
      Tag             =   "�ɱ仯��"
      Top             =   1275
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5662
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   60
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1DBE2
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1DDFC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E016
            Key             =   "In"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E230
            Key             =   "Change"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E44A
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E664
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1E87E
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1EA98
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1ECB2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1EECC
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F0E6
            Key             =   "Adjust"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F300
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F51A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F734
            Key             =   "In"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1F94E
            Key             =   "Change"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1FB68
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1FD82
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":1FF9C
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":201B6
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":203D0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":205EA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20804
            Key             =   "Adjust"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFilter 
      Left            =   855
      Top             =   1590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20A1E
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20B78
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20CD2
            Key             =   "UnCheck_"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCourse.frx":20E2C
            Key             =   "Check_"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicOut 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5655
      MousePointer    =   7  'Size N S
      ScaleHeight     =   225
      ScaleWidth      =   3450
      TabIndex        =   36
      Top             =   3495
      Width           =   3450
      Begin VB.CheckBox chk���� 
         BackColor       =   &H00808080&
         Caption         =   "�ѽ���"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   1350
         TabIndex        =   38
         Top             =   20
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chk���� 
         BackColor       =   &H00808080&
         Caption         =   "δ����"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   2355
         TabIndex        =   37
         Top             =   20
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.Label lblOut 
         BackColor       =   &H00808080&
         Caption         =   " ��Ժ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   190
         Left            =   0
         TabIndex        =   39
         Top             =   20
         Width           =   945
      End
   End
   Begin VB.Label lblBed 
      BackColor       =   &H00808080&
      Caption         =   " ��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   75
      TabIndex        =   31
      Top             =   1035
      Width           =   5475
   End
   Begin VB.Label lblIn 
      BackColor       =   &H00808080&
      Caption         =   " ����ס����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   75
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      Top             =   4515
      Width           =   5460
   End
   Begin VB.Label lblFamily 
      BackColor       =   &H00808080&
      Caption         =   " ��ͥ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   5655
      TabIndex        =   10
      Top             =   1035
      Width           =   3450
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintMed 
         Caption         =   "��ӡ����(&M)"
      End
      Begin VB.Menu mnuFilePrintCard 
         Caption         =   "��ӡ��ͷ��(&C)"
      End
      Begin VB.Menu mnuFile_PrintWristlet 
         Caption         =   "��ӡ���(&W)"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_In 
         Caption         =   "��ס(&I)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEdit_Change 
         Caption         =   "ת��(&C)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit_ChangeUnit 
         Caption         =   "ת����(&T)"
      End
      Begin VB.Menu mnuEdit_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_ChangeGroup 
         Caption         =   "תҽ��С��(&G)"
      End
      Begin VB.Menu mnuEdit_6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Move 
         Caption         =   "����(&M)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEdit_Swap 
         Caption         =   "��λ�Ի�(&S)"
      End
      Begin VB.Menu mnuEdit_AddBeds 
         Caption         =   "����(&B)"
      End
      Begin VB.Menu mnuEdit_7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Out 
         Caption         =   "��Ժ(&O)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEdit_PreOut 
         Caption         =   "Ԥ��Ժ(&P)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEdit_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_ModifOut 
         Caption         =   "�޸ĳ�Ժʱ��(&E)"
      End
      Begin VB.Menu mnuEdit_OutAndModi 
         Caption         =   "��Ժ��������Ժ(&J)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditToInPati 
         Caption         =   "תΪסԺ����(&K)"
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Level 
         Caption         =   "���Ĵ�λ�ȼ�(&B)"
      End
      Begin VB.Menu mnuEdit_Nurse 
         Caption         =   "���Ļ���ȼ�(&N)"
      End
      Begin VB.Menu mnuEdit_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Adjust 
         Caption         =   "����סԺ��Ϣ(&F)"
      End
      Begin VB.Menu mnuEdit_BabyReg 
         Caption         =   "�������Ǽ�(&Y)"
      End
      Begin VB.Menu mnuEdit_Memo 
         Caption         =   "���˱�ע��Ϣ(&Z)"
      End
      Begin VB.Menu mnuEdit_Recalc 
         Caption         =   "���ѱ��������(&R)"
      End
      Begin VB.Menu mnuEdit_Disease 
         Caption         =   "ҽ������ѡ��(&D)"
      End
      Begin VB.Menu mnuEdit_Adjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Undo 
         Caption         =   "����(&U)"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "��ѯ(&Q)"
      Begin VB.Menu mnuQuery_Log 
         Caption         =   "���˱䶯��¼(&C)"
      End
      Begin VB.Menu mnuQuery_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQueryInfo 
         Caption         =   "������Ϣ(&I)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolUnit 
            Caption         =   "����ѡ��(&U)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_Card 
         Caption         =   "��λ��(&C)"
         Checked         =   -1  'True
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuView_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColSel 
         Caption         =   "ѡ����(&C)"
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "��ͼ��(&G)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuView_ListView 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFindNext 
         Caption         =   "������һ��(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuViewreFlash 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmManageCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'����
Private Const COLOR_FOCUS = &H966334   '&HC0844E
Private Const COLOR_LOST = &H808080   '&H966334
Private Const COL_BEDS = "����,1170,0,1;����,959,0,1;�Ա�,650,2,0;����,585,2,0;��λ�ȼ�,975,0,2;" & "����,959,0,2;����ID,750,0,0;סԺ��,799,0,2;��ǰ����,929,2,0;��Ժʱ��,1620,2,2;" & "����ȼ�,1000,0,2;סԺҽʦ,1000,0,0;�����,799,2,0;�Ա����,1000,2,0;��λ����,1000,2,0;����,0,0,1;���￨��,0,0,1;���֤��,0,0,1;IC����,0,0,1;��������,1000,0,2"
Private Const COL_FAMILY = "����,1000,0,1;�Ա�,650,2,0;����,650,2,0;����ID,799,0,0;" & "סԺ��,799,0,2;��ǰ����,1000,0,2;��ǰ����,1000,2,0;��Ժʱ��,1635,2,2;����ȼ�,1000,2,2;סԺҽʦ,1000,0,0;���￨��,0,0,1;���֤��,0,0,1;IC����,0,0,1;��������,1000,0,2"
Private Const COL_IN = "����,1000,0,1;�Ա�,555,2,0;����,650,2,0;����ID,799,0,0;" & "סԺ��,799,0,2;�ѱ�,799,2,0;��ǰ����,1000,0,0;��ǰ����,1000,0,2;ת�����,1000,0,2;" & "��Ժʱ��,1635,2,2;��ǰ����,615,0,0;����ȼ�,1440,0,2;���￨��,0,0,1;���֤��,0,0,1;IC����,0,0,1;��������,1000,0,2"
Private Const COL_OUT = "����,959,0,1;�Ա�,650,2,0;����,650,2,0;����ID,799,0,0;" & "סԺ��,799,0,2;��Ժ��ʽ,1000,2,0;��Ժʱ��,1665,2,2;��Ժʱ��,1635,2,0;��Ժ����,1000,0,2;" & "��Ժ����,929,0,2;��Ժ����,929,2,0;����ȼ�,1000,2,0;�ѱ�,650,2,0;���￨��,0,0,1;���֤��,0,0,1;IC����,0,0,1;��������,1000,0,2;�������,0,0,1"

Private mblnUnload As Boolean
Private mlngPreX As Long, mlngPreY As Long
Private mblnMax As Boolean
Private mblnDropIn As Boolean, mblnDropOut As Boolean '�����б�ߴ���λ
Private mblnDownIn As Boolean, mblnDownOut As Boolean, mblnDownVsc As Boolean '��С����
Private mblnBeds As Boolean, mblnFamily As Boolean, mblnIn As Boolean, mblnOut As Boolean '��Ŀ���
Private mlngUnit As Long
Public mstrPrivs As String
Private mlngModul As Long
'ͳ������
Private mintBeds_A As Integer, mintChange_A As Integer, mintHolding As Integer
Private mintBeds_B As Integer, mintChange_B As Integer
Private mintIn As Integer, mintChange_C As Integer
Private mintOut As Long
'���ݶ���:���б������
Public mobjLVW As ListView '��ǰ��б�
Public mrsBeds As ADODB.Recordset '��λӳ���(��λ������)
Public mrsFamily As ADODB.Recordset '��ͥ��������
Public mrsIn As ADODB.Recordset '��Ʋ���(����Ժ�ǼǺ�ת�Ʋ���)
Public mrsOut As ADODB.Recordset '��Ժ����
'���ݿ�¡,���ڸ�������
Public mrsCBeds As ADODB.Recordset
Public mrsCFamily As ADODB.Recordset
Public mrsCIn As ADODB.Recordset
Public mrsCOut As ADODB.Recordset
'��λ��ʽ
Private mstrSeekKey As String, mstrSeekValue As String
'�������ͼ���ɫ
Private mstrPatiTypeColor As String
Private mstrDeptName As String
Private mlng��ǰ����id As Long
Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cboUnit_Click()
    If cboUnit.ItemData(cboUnit.ListIndex) = mlngUnit Then Exit Sub
    mlngUnit = cboUnit.ItemData(cboUnit.ListIndex)
    Call LoadList(True, True, True, True, True)
End Sub
'����28811 by lesfeng 2010-03-30
Private Sub cboUnit_GotFocus()
    With cboUnit
        .SelStart = 0
        .SelLength = Len(.Text)
        mstrDeptName = .Text
    End With
End Sub
'����28811 by lesfeng 2010-03-30
Private Sub cboUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String
    Dim strSQL As String, intIdx As Long, i As Long
    Dim lngUnit As Long
    
    If KeyCode = vbKeyReturn Then
        If Trim(cboUnit.Text) = "" Then
            If cboUnit.Enabled Then cboUnit.SetFocus
            Exit Sub
        End If
         Set rsTmp = InputGetDept(cboUnit, blnCancel)
        If Not rsTmp Is Nothing Then
            intIdx = cbo.FindIndex(cboUnit, rsTmp!ID)
            If intIdx <> -1 Then
                cboUnit.ListIndex = intIdx
            End If
        Else
            If cboUnit.ListIndex = -1 And cboUnit.ListCount = 0 Then
            Else
                If Not blnCancel Then
                    MsgBox "δ�ҵ���Ӧ�Ĳ�����", vbInformation, gstrSysName
                    cboUnit.SetFocus
                    cboUnit.SelStart = 0
                    cboUnit.Text = mstrDeptName
                    cboUnit.SelLength = Len(cboUnit.Text)
                End If
            End If
        End If
    End If
End Sub
'����28811 by lesfeng 2010-03-30
Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cbr_Resize()
    Form_Resize
End Sub
Private Sub chk����_Click(Index As Integer)
    If chk����(0).Value = 0 And chk����(1).Value = 0 Then
        chk����((Index + 1) Mod 2).Value = 1
    End If
    Call picOut_Click
    LoadList False, False, False, True
End Sub

Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    mnuView_Card.Checked = picCard_s.Visible
    If mobjLVW Is Nothing Then
        lvwBeds_s.SetFocus
    Else
        If mobjLVW.Visible And mobjLVW.Enabled Then mobjLVW.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long, k As Long
    
    If KeyCode = vbKeyF7 Then
        Call tbrFilter_ButtonClick(tbrFilter.Buttons("curDay"))
    ElseIf Shift = vbAltMask And InStr("0123456789", Chr(KeyCode)) > 0 Then
        j = IIf(KeyCode = vbKey0, 10, Val(Chr(KeyCode)))
        For i = 1 To tbrFilter.Buttons.Count
            If tbrFilter.Buttons(i).Key Like "Nurse*" Then
                k = k + 1
                If k = j Then
                    Call tbrFilter_ButtonClick(tbrFilter.Buttons(i))
                    Exit For
                End If
            End If
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim X As Long, Y As Long, i As Integer
    Dim strLoc As String, blnCard As Boolean
    
    RestoreWinState Me, App.ProductName
    picCard_s.width = 2835
    picCard_s.Height = 3865
    
    If lvwBeds_s.ColumnHeaders.Count = 0 Then Call zlControl.LvwSelectColumns(lvwBeds_s, COL_BEDS, True)
    If lvwFamily_s.ColumnHeaders.Count = 0 Then Call zlControl.LvwSelectColumns(lvwFamily_s, COL_FAMILY, True)
    If lvwIn_s.ColumnHeaders.Count = 0 Then Call zlControl.LvwSelectColumns(lvwIn_s, COL_IN, True)
    If lvwOut_s.ColumnHeaders.Count = 0 Then Call zlControl.LvwSelectColumns(lvwOut_s, COL_OUT, True)
    
    '����λ������
    For Y = 0 To picCard_s.Height / Screen.TwipsPerPixelY Step 87
        For X = 0 To picCard_s.width / Screen.TwipsPerPixelX Step 50
            BitBlt picCard_s.hDC, X, Y, 50, 87, picCard_s.hDC, 0, 0, SRCCOPY
        Next
    Next
    zlControl.PicShowFlat picCard_s, 1
    
    '��������ͼ��
    Call MakeBedIcon
    
    mblnUnload = False
    mblnBeds = False: mblnFamily = False: mblnIn = False: mblnOut = False
    mlngUnit = 0
   
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    'Ȩ������
    If InStr(mstrPrivs, "���˳�Ժ") = 0 Then
        mnuEdit_Out.Visible = False
        tbr.Buttons("Out").Visible = False
    End If
    If InStr(mstrPrivs, "����ת��") = 0 Then
        mnuEdit_Change.Visible = False
        tbr.Buttons("Change").Visible = False
    End If
    '����Ȩ�޿���
    If InStr(mstrPrivs, "����") = 0 Then
        mnuEdit_Move.Visible = False
        mnuEdit_Swap.Visible = False
        mnuEdit_AddBeds.Visible = False
        tbr.Buttons("Move").Visible = False
        mnuEdit_7.Visible = False
    End If
    
    If InStr(mstrPrivs, "ת����") = 0 Then
        mnuEdit_ChangeUnit.Visible = False
        tbr.Buttons("Change").Visible = mnuEdit_Change.Visible
    End If
    
    If InStr(mstrPrivs, "����Ԥ��Ժ") = 0 Then
        mnuEdit_PreOut.Visible = False
    End If
    If InStr(mstrPrivs, "����������Ϣ") = 0 Then
        mnuEdit_Adjust.Visible = False
        'mnuEdit_Adjust_.Visible = False
        tbr.Buttons("Adjust").Visible = False
        'tbr.Buttons("Adjust_").Visible = False
    End If
    If InStr(mstrPrivs, "�������Ǽ�") = 0 Then
        mnuEdit_BabyReg.Visible = False
    End If
    If InStr(mstrPrivs, "�������") = 0 Then
        mnuEdit_Recalc.Visible = False
    End If
    '����27392 by lesfeng 2010-01-14
    If InStr(mstrPrivs, "������Ժʱ��") = 0 Then
        mnuEdit_ModifOut.Visible = False
    End If
    '����27866 by lesfeng 2010-02-05
    If (InStr(mstrPrivs, "���˳�Ժ") = 0 Or InStr(mstrPrivs, "������Ժʱ��") = 0) Then
        mnuEdit_OutAndModi.Visible = False
    End If
    If Not (mnuEdit_ModifOut.Visible Or mnuEdit_OutAndModi.Visible) Then
        mnuEdit_4.Visible = False
    End If
    
    If InStr(mstrPrivs, "������λ�ȼ�") = 0 Then
        mnuEdit_Level.Visible = False
    End If
    
'    If InStr(mstrPrivs, "���˱�ע�༭") = 0 Then
'        mnuEdit_Memo.Visible = False
'    End If

    If InStr(mstrPrivs, "��������ȼ�") = 0 Then
        mnuEdit_Nurse.Visible = False
    End If
    
    If InStr(mstrPrivs, "סԺ����תסԺ") = 0 Then
        mnuEditToInPati.Visible = False
        mnuEdit_2.Visible = False
    End If
                
    Call InitPatiType
    
    If Val(zlDatabase.GetPara("������Ժ", glngSys, mlngModul, 0)) = 0 Then
        tbrFilter.Buttons("curDay").Image = "UnCheck_"
    Else
        tbrFilter.Buttons("curDay").Image = "Check_"
    End If
    For i = 1 To tbrFilter.Buttons.Count
        If tbrFilter.Buttons(i).Key Like "Nurse*" Then
            If Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "����ȼ�" & Replace(tbrFilter.Buttons(i).Key, "Nurse", ""), 1)) <> 0 Then
                tbrFilter.Buttons(i).Image = "Check"
            Else
                tbrFilter.Buttons(i).Image = "UnCheck"
            End If
        End If
    Next
            
    '��ʼסԺ����
    If Not InitUnits Then mblnUnload = True: Exit Sub
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, Me.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�
    Dim staH As Long '״̬��ռ�ø߶�
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    If mblnMax Then
        lvwFamily_s.Height = Me.ScaleHeight / 3
        lvwIn_s.Height = Me.ScaleHeight / 4
        
        lvwFamily_s.width = Me.ScaleWidth * 0.35
        lvwOut_s.width = Me.ScaleWidth * 0.35
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0) + tbrFilter.Height
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    With lblBed
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop + cbrH + 15
        .width = Me.ScaleWidth - lvwFamily_s.width - picVsc.width
    End With
    With lvwBeds_s
        .Left = lblBed.Left
        .Top = lblBed.Top + lblBed.Height + 15
        .width = lblBed.width
        .Height = Me.ScaleHeight - lvwIn_s.Height - lblIn.Height - lblBed.Height - cbrH - staH - 60
    End With
    With lblIn
        .Top = lvwBeds_s.Top + lvwBeds_s.Height + 15
        .Left = lblBed.Left
        .width = lblBed.width
    End With
    With lvwIn_s
        .Top = lblIn.Top + lblIn.Height + 15
        .Left = lblIn.Left
        .width = lvwBeds_s.width
    End With
    With picVsc
        .Top = lblBed.Top
        .Left = lblBed.Left + lblBed.width
        .Height = Me.ScaleHeight - cbrH - staH
    End With
    With lblFamily
        .Top = lblBed.Top
        .Left = picVsc.Left + picVsc.width
        .width = lvwFamily_s.width
    End With
    With lvwFamily_s
        .Top = lblFamily.Top + lblFamily.Height + 15
        .Left = lblFamily.Left
    End With
    With PicOut
        .Left = lblFamily.Left
        .Top = lvwFamily_s.Top + lvwFamily_s.Height + 15
        .width = lvwFamily_s.width
    End With
    With lvwOut_s
        .Left = PicOut.Left
        .Top = PicOut.Top + PicOut.Height + 15
        .width = lvwFamily_s.width
        .Height = Me.ScaleHeight - lblFamily.Height - PicOut.Height - lvwFamily_s.Height - cbrH - staH - 60
    End With
    Me.Refresh
    
    If WindowState = 0 Or WindowState = 2 Then
        timSize.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    Set mrsBeds = Nothing
    Set mrsFamily = Nothing
    Set mrsIn = Nothing
    Set mrsOut = Nothing
    
    Set mrsCBeds = Nothing
    Set mrsCFamily = Nothing
    Set mrsCIn = Nothing
    Set mrsCOut = Nothing
    
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
    
    mstrSeekKey = "": mstrSeekValue = ""
    
    SaveWinState Me, App.ProductName
    
    zlDatabase.SetPara "������Ժ", IIf(tbrFilter.Buttons("curDay").Image = "UnCheck_", 0, 1), glngSys, mlngModul
    For i = 1 To tbrFilter.Buttons.Count
        If tbrFilter.Buttons(i).Key Like "Nurse*" Then
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "����ȼ�" & Replace(tbrFilter.Buttons(i).Key, "Nurse", ""), IIf(tbrFilter.Buttons(i).Image = "Check", 1, 0)
        End If
    Next
End Sub

Private Sub lblBed_Click()
    lvwBeds_s.SetFocus
End Sub

Private Sub lblFamily_Click()
    lvwFamily_s.SetFocus
End Sub

Private Sub lblIn_Click()
    If mblnDropIn Then
        If lvwBeds_s.Height >= lvwIn_s.Height Then
            lvwBeds_s.SetFocus
        Else
            lvwIn_s.SetFocus
        End If
    Else
        lvwIn_s.SetFocus
    End If
End Sub

Private Sub lblIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreY = Y: mblnDownIn = True: mblnDropIn = False
End Sub

Private Sub lblIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnDownIn Then
        If lvwIn_s.Height - (Y - mlngPreY) < 600 Or lvwBeds_s.Height + Y - mlngPreY < 600 Then Exit Sub
        lblIn.Top = lblIn.Top + Y - mlngPreY
        lvwBeds_s.Height = lvwBeds_s.Height + Y - mlngPreY
        lvwIn_s.Top = lvwIn_s.Top + Y - mlngPreY
        lvwIn_s.Height = lvwIn_s.Height - (Y - mlngPreY)
        Me.Refresh
        mblnDropIn = True
    End If
End Sub

Private Sub lblIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDownIn = False
End Sub

Private Sub lblOut_Click()
 Call picOut_Click
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    Dim strValue As String, strDepts As String
    Dim lngInTime As Long, lngDept As Long, lngUnit As Long, strCurDate As String
    Dim lngPatID As Long, lngPageID As Long
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnExit As Boolean
    
    On Error GoTo ErrHand
    
    If UCase(strMsgItemIdentity) = "ZLHIS_PATIENT_001" Then
        If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
        strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
        strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
        If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
        '��鲡��
        If mclsXML.GetSingleNodeValue("in_dept_id", strValue, xsNumber) = False Then Exit Sub
        lngDept = Val(strValue)
        strValue = "": Call mclsXML.GetSingleNodeValue("in_area_id", strValue, xsNumber)
        If Val(strValue) = 0 Then
            strValue = ""
            strSQL = "Select ����ID From �������Ҷ�Ӧ where ����ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngDept)
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!����ID
            rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        End If
        If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
        
        '�����Ժ�����Ƿ��ڴ���Ʋ��˿�����
        strDepts = zlDatabase.GetPara("����Ʋ��˿���", glngSys, mlngModul, "")
        If strDepts <> "" Then
            strDepts = "," & strDepts & ","
            If InStr(1, strDepts, "," & lngDept & ",") = 0 Then Exit Sub
        End If
        '�����Ժʱ���Ƿ�����Ժ�Ǽ�������
        strValue = "": Call mclsXML.GetSingleNodeValue("in_date", strValue, xsString)
        If IsDate(strValue) Then
            lngInTime = Val(zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, 3))
            strCurDate = zlDatabase.Currentdate
            If lngInTime <> 0 Then
                strCurDate = Format(DateAdd("D", -1 * lngInTime, CDate(strCurDate)), "YYYY-MM-DD HH:mm:ss")
            Else
                strCurDate = Format(strCurDate, "YYYY-MM-DD")
            End If
            If Format(strValue, "YYYY-MM-DD HH:mm:ss") < Format(strCurDate, "YYYY-MM-DD HH:mm:ss") Then Exit Sub
        End If
        '��ȡ������Ϣ
        strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsNumber)
        mclsXML.CloseXMLDocument
        mrsIn.Filter = "����ID=" & lngPatID
        If mrsIn.EOF = True Then
            Call LoadList(False, False, True, False)
            If strValue <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "���µǼǵĲ���:" & strValue, "�����������")
            End If
        End If
    ElseIf UCase(strMsgItemIdentity) = "ZLHIS_PATIENT_003" Then
        If mclsXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
        strValue = "": Call mclsXML.GetSingleNodeValue("send_program", strValue, xsString)
        If strValue <> "" And Val(strValue) = Me.hWnd Then Exit Sub
        strValue = "": Call mclsXML.GetSingleNodeValue("patient_id", strValue, xsNumber): lngPatID = Val(strValue)
        strValue = "": Call mclsXML.GetSingleNodeValue("page_id", strValue, xsNumber): lngPageID = Val(strValue)
        If lngPatID = 0 Or lngPageID = 0 Then Exit Sub
        '��鲡��
        strValue = "": Call mclsXML.GetSingleNodeValue("change_dept_id", strValue, xsNumber)
        lngDept = Val(strValue)
        strValue = "": Call mclsXML.GetSingleNodeValue("change_area_id", strValue, xsNumber)
        lngUnit = Val(strValue)
        
        If lngDept = 0 Then Exit Sub
        
        If lngUnit = 0 Then
            strValue = ""
            strSQL = "Select ����ID From �������Ҷ�Ӧ where ����ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngDept)
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!����ID
            rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        Else
            strValue = lngUnit
        End If
        If InStr(1, "," & strValue & ",", "," & cboUnit.ItemData(cboUnit.ListIndex) & ",") = 0 Then Exit Sub
        
        '�����Ժ�����Ƿ��ڴ���Ʋ��˿�����
        strDepts = zlDatabase.GetPara("����Ʋ��˿���", glngSys, mlngModul, "")
        If strDepts <> "" Then
            strDepts = "," & strDepts & ","
            If InStr(1, strDepts, "," & lngDept & ",") = 0 Then Exit Sub
        End If
        
        '��ȡ������Ϣ
        strValue = "": Call mclsXML.GetSingleNodeValue("patient_name", strValue, xsNumber)
        mclsXML.CloseXMLDocument
        mrsIn.Filter = "����ID=" & lngPatID
        If mrsIn.EOF = True Then
            Call LoadList(True, True, True, False)
            If strValue <> "" Then
                Call mclsMipModule.ShowMessage(strMsgItemIdentity, "����ת��Ĳ���:" & strValue, "�����������")
            End If
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEdit_ChangeGroup_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    End If
    If ExecPatiChange(EFun.Eתҽ��С��, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID) Then
        Call LoadList(mobjLVW Is lvwBeds_s, mobjLVW Is lvwFamily_s, False, False)
    End If
End Sub

Private Sub mnuEdit_ChangeUnit_Click()
    Call ChangeUnit
End Sub

Private Sub mnuEdit_InUnit_Click()
    Dim strBeds As String, byt��Ʒ�ʽ As Byte, lng��λ����ID As Long
    Dim lng����ID As Long, lng��ҳID As Long
    
    If lvwBeds_s.SelectedItem Is Nothing Then
        strBeds = ""
    ElseIf lvwBeds_s.SelectedItem.Tag <> "�մ�" Then
        If mrsBeds!����ID = mrsIn!����ID Then '����ԭס��
            strBeds = Trim(Mid(lvwBeds_s.SelectedItem.Key, 2))
            lng��λ����ID = Val("" & mrsIn!��ס����id)
        Else
            strBeds = ""
        End If
    ElseIf Not (mrsBeds!�Ա���� = "���޴�" Or (mrsBeds!�Ա���� = "�д�" And "" & mrsIn!�Ա� = "��") _
        Or (mrsBeds!�Ա���� = "Ů��" And "" & mrsIn!�Ա� = "Ů")) Then
        strBeds = ""
    Else
        strBeds = Trim(Mid(lvwBeds_s.SelectedItem.Key, 2))
        lng��λ����ID = Val("" & mrsBeds!����ID)
    End If
    byt��Ʒ�ʽ = Val(lvwIn_s.SelectedItem.Tag)
    lng����ID = mrsIn!����ID
    lng��ҳID = mrsIn!��ҳID
    
    Call ExecPatiChange(EFun.E�벡��, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID, strBeds, lng��λ����ID)
    
    '��ƺ�λ
    If gblnOK Then
        If strBeds <> "" Then
            If InStr(strBeds, ",") Then
                strBeds = Split(strBeds, ",")(0)
                lvwBeds_s.ListItems("_" & strBeds).Selected = True
            Else
                lvwBeds_s.ListItems("_" & strBeds).Selected = True
            End If
        End If
        Call LoadList(True, True, True, False)
    End If
End Sub

Private Sub mnuEdit_Memo_Click()
    Dim lng����ID As Long, lng��ҳID As Long, strBeds As String
    
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    ElseIf mobjLVW Is lvwOut_s Then
        lng����ID = mrsOut!����ID
        lng��ҳID = mrsOut!��ҳID
    ElseIf mobjLVW Is lvwIn_s Then
        lng����ID = mrsIn!����ID
        lng��ҳID = mrsIn!��ҳID
    End If
    
    Call ExecPatiChange(EFun.E���˱�ע�༭, Me, mstrPrivs, lng����ID, lng��ҳID)
'
'    If gblnOK Then Call LoadList(True, True, True, False)
End Sub

'����27392 by lesfeng 2010-01-14
Private Sub mnuEdit_ModifOut_Click()
    '���ܣ��޸Ĳ��˳�Ժʱ��
    Dim lng����ID As Long, lng��ҳID As Long, str���� As String
    
    If mobjLVW Is lvwOut_s Then
        lng����ID = mrsOut!����ID
        lng��ҳID = mrsOut!��ҳID
        str���� = mrsOut!����
        Call ExecPatiChange(EFun.E�޸ĳ�Ժʱ��, Me, mstrPrivs, lng����ID, lng��ҳID)
        If gblnOK Then Call LoadList(False, False, False, True)
    End If
End Sub
'����27866 by lesfeng 2010-02-05
Private Sub mnuEdit_OutAndModi_Click()
    frmOutAndModi.Show 1, Me
End Sub

Private Sub mnuEdit_Swap_Click()
    Call SwapBeds
End Sub

Private Sub mnuFile_PrintWristlet_Click()
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    '49854:������,2013-10-31,���������ӡ(�ų���Ժ����)
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    ElseIf mobjLVW Is lvwIn_s Then
        lng����ID = mrsIn!����ID
        lng��ҳID = mrsIn!��ҳID
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, 2)
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    frmSetCourse.mlngModul = mlngModul
    frmSetCourse.mstrPrivs = mstrPrivs
    frmSetCourse.Show 1, Me
    If gblnOK Then
        LoadList False, False, True, True
    End If
End Sub

Private Sub mnuViewFind_Click()
    Dim intIdx As Long
    With frmFindCourse
        .mstrSeekKey = mstrSeekKey
        .mstrSeekValue = mstrSeekValue
        .Show 1, Me
        If .mblnOk Then
            mstrSeekKey = .mstrSeekKey
            mstrSeekValue = .mstrSeekValue
            mlng��ǰ����id = .mlng����id
            ' ����30040 by lesfeng 2010-05-18
            If mstrSeekKey <> "����" Then
                If mlng��ǰ����id <> mlngUnit Then
                    If InStr(mstrPrivs, "���в���") <> 0 Then
                        intIdx = cbo.FindIndex(cboUnit, mlng��ǰ����id)
                        If intIdx <> -1 Then
                            cboUnit.ListIndex = intIdx
                        End If
'                        mlngUnit = mlng��ǰ����id
'                        Call LoadList
                    Else
                        MsgBox "��û�С����в�����Ȩ�ޣ����ܲ��ҵ� " & mstrSeekKey & "=" & mstrSeekValue & " �Ĳ��ˣ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
            Call SeekPati(True)
        End If
    End With
End Sub
Private Sub SeekPati(ByVal blnFirst As Boolean)
Dim lvwRow As Integer, lvwfor As Integer, intStart As Integer
Dim lvwTemp As ListView, lviTemp As ListItem, intColKey As Integer, lvwValue As String
    If mstrSeekValue = "" Then Exit Sub
    
reFind:
    lvwfor = lvwfor + 1
    If lvwfor > 4 Then
        MsgBox "û����Ҫ���ҵ� " & mstrSeekKey & "=" & mstrSeekValue & " �Ĳ��ˣ�", vbInformation, gstrSysName
        mstrSeekKey = "": mstrSeekValue = ""
        Exit Sub
    End If
    
    '���õ�ǰ���ҵ��б�
    If lvwfor = 1 Then
        Set lvwTemp = mobjLVW
    Else '��ǰ�б���û�ҵ��������б��л�
        Select Case lvwTemp.Name
            Case "lvwIn_s"
                Set lvwTemp = lvwBeds_s
            Case "lvwBeds_s"
                Set lvwTemp = lvwFamily_s
            Case "lvwFamily_s"
                Set lvwTemp = lvwOut_s
            Case "lvwOut_s"
                Set lvwTemp = lvwIn_s
        End Select
    End If
    
    '�ڵ�ǰ�б�������
    With lvwTemp
        intStart = 1
        If Not blnFirst And lvwfor = 1 Then '����ǵ�һ�β���,��ʼλ��
            If .ListItems.Count > 0 Then intStart = .SelectedItem.Index + 1
        End If
        
        intColKey = GetColNum(lvwTemp, mstrSeekKey) 'ȡ����
        If intColKey <> 0 Or mstrSeekKey = "ҽ����" Then
            For lvwRow = intStart To .ListItems.Count
                If mstrSeekKey = "ҽ����" Then '��ҽ���Ų��ҵ����⴦��
                    If .Name = "lvwBeds_s" Then 'ȡ������ID
                        lvwValue = Trim(.ListItems(lvwRow).SubItems(GetColNum(lvwTemp, "����ID") - 1))
                    Else
                        lvwValue = Trim(Split(.ListItems(lvwRow).Key, "_")(1))
                    End If
                    
                    If lvwValue <> "" Then 'ȡҽ����
                        lvwValue = GetInsureInfo(CLng(lvwValue))
                        If InStr(lvwValue, ";") > 0 Then
                            lvwValue = Trim(Split(lvwValue, ";")(1))
                        Else
                            lvwValue = ""
                        End If
                    End If
                Else
                    If intColKey = 1 Then 'ȡ�������ж�Ӧֵ
                        lvwValue = Trim(.ListItems(lvwRow).Text)
                    Else
                        lvwValue = Trim(.ListItems(lvwRow).SubItems(intColKey - 1))
                    End If
                End If
                
                If mstrSeekValue = lvwValue Then '��ͬ��λ���˳�
                    Set lviTemp = .ListItems(lvwRow)
                    Select Case .Name
                        Case "lvwIn_s"
                            lvwIn_s.ListItems(lviTemp.Key).Selected = True
                            lvwIn_s.SelectedItem.EnsureVisible
                            Call lvwIn_s_ItemClick(lviTemp)
                            lvwIn_s.SetFocus
                            Call lvwIn_s_GotFocus
                        Case "lvwBeds_s"
                            lvwBeds_s.ListItems(lviTemp.Key).Selected = True
                            lvwBeds_s.SelectedItem.EnsureVisible
                            Call lvwBeds_s_ItemClick(lviTemp)
                            lvwBeds_s.SetFocus
                            Call lvwBeds_s_GotFocus
                        Case "lvwFamily_s"
                            lvwFamily_s.ListItems(lviTemp.Key).Selected = True
                            lvwFamily_s.SelectedItem.EnsureVisible
                            Call lvwFamily_s_ItemClick(lviTemp)
                            lvwFamily_s.SetFocus
                            Call lvwFamily_s_GotFocus
                        Case "lvwOut_s"
                            lvwOut_s.ListItems(lviTemp.Key).Selected = True
                            lvwOut_s.SelectedItem.EnsureVisible
                            Call lvwOut_s_ItemClick(lviTemp)
                            lvwOut_s.SetFocus
                            Call lvwOut_s_GotFocus
                    End Select
                    Exit Sub
                End If
            Next
        End If
    End With
    GoTo reFind
End Sub
Private Sub mnuViewFindNext_Click()
    Call SeekPati(False)
End Sub

Private Sub picOut_Click()
    If mblnDropOut Then
        If lvwFamily_s.Height >= lvwOut_s.Height Then
            lvwFamily_s.SetFocus
        Else
            lvwOut_s.SetFocus
        End If
    Else
        lvwOut_s.SetFocus
    End If
End Sub

Private Sub picOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreY = Y: mblnDownOut = True: mblnDropOut = False
End Sub

Private Sub picOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnDownOut Then
        If lvwOut_s.Height - (Y - mlngPreY) < 600 Or lvwFamily_s.Height + Y - mlngPreY < 600 Then Exit Sub
        PicOut.Top = PicOut.Top + Y - mlngPreY
        lvwFamily_s.Height = lvwFamily_s.Height + Y - mlngPreY
        lvwOut_s.Top = lvwOut_s.Top + Y - mlngPreY
        lvwOut_s.Height = lvwOut_s.Height - (Y - mlngPreY)
        Me.Refresh
        mblnDropOut = True
    End If
End Sub

Private Sub picOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDownOut = False
End Sub

Private Sub lvwBeds_s_DblClick()
    If mblnBeds Then
        If lvwBeds_s.SelectedItem.Tag = "ռ��" Then
            mnuQueryInfo_Click
        End If
    End If
End Sub

Private Sub lvwBeds_s_GotFocus()
    Call ClearCard
    If Not lvwBeds_s.SelectedItem Is Nothing Then Call lvwBeds_s_ItemClick(lvwBeds_s.SelectedItem)
    Set mobjLVW = lvwBeds_s
    Call SetFocusColor
    
    Call SetMenu
End Sub

Private Sub ClearCard()
    lbl����.Caption = "����:"
    lbl����.Caption = "����"
    lbl�Ա�.Caption = "�Ա�"
    lbl����.Caption = "����"
    lbl��ʶ.Caption = "סԺ��"
    lblסԺ��.Caption = ""
    lblҽ����.Caption = ""
    lblҽ����.ForeColor = Me.ForeColor
    lbl����.Caption = ""
    lbl��������.Caption = ""
    lbl��Ժʱ��.Caption = ""
    lbl����ȼ�.Caption = ""
    lblҽ��.Caption = ""
    lblLevel.Caption = ""
    lblҽ�Ƹ��ʽ.Caption = ""
    lbl���.Caption = ""
End Sub

Private Sub lvwBeds_s_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strInfo As String
    If Item Is Nothing Then Exit Sub
    mrsBeds.Filter = "����='" & Mid(Item.Key, 2) & "'"

    mblnBeds = True
    If Item.Tag = "ռ��" And Not IsNull(mrsBeds!����ID) Then
        If Nvl(mrsBeds!�����) <> "" Then
            lbl����.Caption = "����:" & mrsBeds!���� & "(" & Nvl(mrsBeds!�����) & ")"
        Else
            lbl����.Caption = "����:" & mrsBeds!����
        End If
        If Not IsNull(mrsBeds!����) Then strInfo = GetInsureInfo(mrsBeds!����ID)
        If strInfo <> "" Then
            lblҽ����.Caption = Split(strInfo, ";")(1)
            lblҽ����.ForeColor = vbRed
        Else
            lblҽ����.Caption = "��ҽ������"
            lblҽ����.ForeColor = Me.ForeColor
        End If
        
        lbl����.Caption = mrsBeds!����
        lbl�Ա�.Caption = IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�)
        lbl����.Caption = IIf(IsNull(mrsBeds!����), "", mrsBeds!����)
        
        If Not IsNull(mrsBeds!סԺ��) Then
            lbl��ʶ.Caption = "סԺ��"
            lblסԺ��.Caption = mrsBeds!סԺ��
        Else
            lbl��ʶ.Caption = "����ID"
            lblסԺ��.Caption = mrsBeds!����ID
        End If
        
        lbl��������.Caption = Nvl(mrsBeds!��������, "��ͨ����")
        lbl����.Caption = IIf(IsNull(mrsBeds!��ǰ����), "", mrsBeds!��ǰ����)
        lbl��Ժʱ��.Caption = Format(mrsBeds!��Ժʱ��, "yyyy-MM-dd HH:mm")
        lbl����ȼ�.Caption = IIf(IsNull(mrsBeds!����ȼ�), "", mrsBeds!����ȼ�)
        lblҽ��.Caption = IIf(IsNull(mrsBeds!סԺҽʦ), "", mrsBeds!סԺҽʦ)
        lblLevel.Caption = IIf(IsNull(mrsBeds!��λ�ȼ�), "", mrsBeds!��λ�ȼ�)
        lblҽ�Ƹ��ʽ.Caption = "" & mrsBeds!ҽ�Ƹ��ʽ
        lbl���.Caption = GetDiagnostic(mrsBeds!����ID, Val("" & mrsBeds!��ҳID))
        
        stbThis.Panels(2).Text = "סԺ��:" & IIf(IsNull(mrsBeds!סԺ��), "", mrsBeds!סԺ��) & " ��Ժ:" & mrsBeds!��Ժʱ�� & _
            " ����:" & IIf(IsNull(mrsBeds!����ȼ�), "", mrsBeds!����ȼ�) & " " & _
            " ����:" & mrsBeds!���� & " �ȼ�:" & mrsBeds!��λ�ȼ�
    Else
        ClearCard
        If Nvl(mrsBeds!�����) <> "" Then
            lbl����.Caption = "����:" & mrsBeds!���� & "(" & Nvl(mrsBeds!�����) & ")"
        Else
            lbl����.Caption = "����:" & mrsBeds!����
        End If
        lblLevel.Caption = IIf(IsNull(mrsBeds!��λ�ȼ�), "", mrsBeds!��λ�ȼ�)
        stbThis.Panels(2).Text = "����:" & mrsBeds!���� & " �ȼ�:" & mrsBeds!��λ�ȼ�
    End If
    
    Call SetMenu
End Sub

Private Function GetDiagnostic(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetDiagnosticInfo(lng����ID, lng��ҳID, "1,2,3", 2)
    If Not rsTmp Is Nothing Then
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "�������=3"
            If rsTmp.RecordCount > 0 Then
                GetDiagnostic = "" & rsTmp!�������
            Else
                rsTmp.Filter = "�������=2"
                If rsTmp.RecordCount > 0 Then
                    GetDiagnostic = "" & rsTmp!�������
                Else
                    rsTmp.Filter = "�������=1"
                    If rsTmp.RecordCount > 0 Then GetDiagnostic = "" & rsTmp!�������
                End If
            End If
        End If
    End If
End Function




Private Sub lvwBeds_s_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not lvwBeds_s.SelectedItem Is Nothing And KeyCode = vbKeyReturn Then
        mblnBeds = True: Call lvwBeds_s_DblClick
    End If
End Sub

Private Sub lvwBeds_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnBeds = False
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        If lvwBeds_s.HitTest(X, Y) Is Nothing Then
            stbThis.Panels(2) = "���������� " & mintBeds_A & " ��,����ռ�� " & mintHolding & " ��,�մ� " & mintBeds_A - mintHolding & " ��,ת�Ʋ��� " & mintChange_A & " ��"
        End If
    End If
End Sub

Private Sub lvwFamily_s_DblClick()
    If mblnFamily Then mnuQueryInfo_Click
End Sub

Private Sub lvwFamily_s_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If InStr(1, mstrPrivs, "��ͥ����") = 0 Then
        Static objIconFam As IPictureDisp
        If Not Source Is lvwIn_s Then
            If State = 0 Then
                Set objIconFam = Source.DragIcon
            ElseIf State = 2 Then
                Set Source.DragIcon = img32.ListImages("Limit").Picture
            ElseIf State = 1 Then
                Set Source.DragIcon = objIconFam
            End If
        End If
    End If
End Sub

Private Sub lvwFamily_s_GotFocus()
    Call ClearCard
    If Not lvwFamily_s.SelectedItem Is Nothing Then Call lvwFamily_s_ItemClick(lvwFamily_s.SelectedItem)
    Set mobjLVW = lvwFamily_s
    Call SetFocusColor
    
    Call SetMenu
End Sub

Private Sub lvwFamily_s_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strInfo As String
    If Item Is Nothing Then Exit Sub
    mrsFamily.Filter = "����ID=" & Mid(Item.Key, 2)
    
    mblnFamily = True
    lbl����.Caption = "����:��"
    lblLevel.Caption = "��ͥ����"
    
    strInfo = GetInsureInfo(mrsFamily!����ID)
    If strInfo <> "" Then
        lblҽ����.Caption = Split(strInfo, ";")(1)
        lblҽ����.ForeColor = vbRed
    Else
        lblҽ����.Caption = "��ҽ������"
        lblҽ����.ForeColor = Me.ForeColor
    End If
    
    lbl����.Caption = mrsFamily!����
    lbl�Ա�.Caption = IIf(IsNull(mrsFamily!�Ա�), "", mrsFamily!�Ա�)
    lbl����.Caption = IIf(IsNull(mrsFamily!����), "", mrsFamily!����)
    
    If Not IsNull(mrsFamily!סԺ��) Then
        lbl��ʶ.Caption = "סԺ��"
        lblסԺ��.Caption = mrsFamily!סԺ��
    Else
        lbl��ʶ.Caption = "����ID"
        lblסԺ��.Caption = mrsFamily!����ID
    End If
    
    lbl����.Caption = IIf(IsNull(mrsFamily!��ǰ����), "", mrsFamily!��ǰ����)
    lbl��Ժʱ��.Caption = Format(mrsFamily!��Ժʱ��, "yyyy-MM-dd HH:mm")
    lbl����ȼ�.Caption = IIf(IsNull(mrsFamily!����ȼ�), "", mrsFamily!����ȼ�)
    lblҽ��.Caption = IIf(IsNull(mrsFamily!סԺҽʦ), "", mrsFamily!סԺҽʦ)
    lblҽ�Ƹ��ʽ.Caption = "" & mrsFamily!ҽ�Ƹ��ʽ
    lbl���.Caption = GetDiagnostic(mrsFamily!����ID, Val("" & mrsFamily!��ҳID))
    
    stbThis.Panels(2).Text = "סԺ��:" & IIf(IsNull(mrsFamily!סԺ��), "", mrsFamily!סԺ��) & " ����:" & IIf(IsNull(mrsFamily!����ȼ�), "", mrsFamily!����ȼ�) & " ����:" & mrsFamily!��ǰ����
    
    Call SetMenu
End Sub

Private Sub lvwFamily_s_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not lvwFamily_s.SelectedItem Is Nothing And KeyCode = vbKeyReturn Then
        mblnFamily = True: Call lvwFamily_s_DblClick
    End If
End Sub

Private Sub lvwFamily_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnFamily = False
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        If lvwFamily_s.HitTest(X, Y) Is Nothing Then
            stbThis.Panels(2) = "��ͥ������ " & mintBeds_B & " ��,ת�Ʋ��� " & mintChange_B & " ��"
        End If
    End If
End Sub

Private Sub lvwIn_s_DblClick()
    If mblnIn Then mnuQueryInfo_Click
End Sub

Private Sub lvwIn_s_GotFocus()
    If Not lvwIn_s.SelectedItem Is Nothing Then Call lvwIn_s_ItemClick(lvwIn_s.SelectedItem)
    Set mobjLVW = lvwIn_s
    Call SetFocusColor
    
    Call SetMenu
End Sub

Private Sub lvwIn_s_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item Is Nothing Then Exit Sub
    mrsIn.Filter = "����ID=" & Mid(Item.Key, 2)
    mblnIn = True
    
    stbThis.Panels(2).Text = "סԺ��:" & IIf(IsNull(mrsIn!סԺ��), "", mrsIn!סԺ��) & " ��ǰ����:" & mrsIn!��ǰ���� & _
        IIf(IsNull(mrsIn!ת�����), "", " ת�����:" & mrsIn!ת�����) & _
        " ��Ժʱ��:" & mrsIn!��Ժʱ��
    
    Call SetMenu
End Sub

Private Sub lvwIn_s_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not lvwIn_s.SelectedItem Is Nothing And KeyCode = vbKeyReturn Then
        mblnIn = True: Call lvwIn_s_DblClick
    End If
End Sub

Private Sub lvwIn_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnIn = False
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        If lvwIn_s.HitTest(X, Y) Is Nothing Then
            stbThis.Panels(2) = "��Ʋ��˹� " & mintIn & " ��,����Ժ " & mintIn - mintChange_C & " ��,����ת�� " & mintChange_C & " ��"
        End If
    End If
End Sub

Private Sub lvwOut_s_DblClick()
    If mblnOut Then mnuQueryInfo_Click
End Sub

Private Sub lvwOut_s_GotFocus()
    If Not lvwOut_s.SelectedItem Is Nothing Then Call lvwOut_s_ItemClick(lvwOut_s.SelectedItem)
    Set mobjLVW = lvwOut_s
    Call SetFocusColor
    
    Call SetMenu
End Sub

Private Sub lvwOut_s_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item Is Nothing Then Exit Sub
    '����28365 by lesfeng 2010-03-04 ���˶�γ�Ժʱ��û��������ҳid ����
    mrsOut.Filter = "����ID=" & Split(Item.Key, "_")(1) & " and ��ҳid=" & Split(Item.Key, "_")(2)
    Call ClearCard
    mblnOut = True
        
    stbThis.Panels(2).Text = "סԺ��:" & IIf(IsNull(mrsOut!סԺ��), "", mrsOut!סԺ��) & " ��Ժ����:" & mrsOut!��Ժ���� & " ��Ժ����:" & mrsOut!��Ժ���� & " ��Ժʱ��:" & mrsOut!��Ժʱ��
    
    Call SetMenu
End Sub

Private Sub lvwOut_s_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not lvwOut_s.SelectedItem Is Nothing And KeyCode = vbKeyReturn Then
        mblnOut = True: Call lvwOut_s_DblClick
    End If
End Sub

Private Sub lvwOut_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnOut = False
    If Button = 2 Then
        PopupMenu mnuEdit, 2
    Else
        If lvwOut_s.HitTest(X, Y) Is Nothing Then
            stbThis.Panels(2) = "��Ժ���˹� " & mintOut & " ��"
        End If
    End If
End Sub

Private Sub mnuEdit_AddBeds_Click()
    If Not mobjLVW Is lvwBeds_s Then Exit Sub
            
    Call ChangeBeds(1)
End Sub

Private Sub mnuEdit_Adjust_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    
    If mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    ElseIf mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    End If
    Call ExecPatiChange(EFun.E����������Ϣ, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID)
    
    If gblnOK Then Call LoadList(True, True, False, False)
End Sub

Private Sub mnuEdit_BabyReg_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    
    If mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    ElseIf mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    End If
    Call ExecPatiChange(EFun.E�������Ǽ�, Me, mstrPrivs, lng����ID, lng��ҳID)
End Sub

Private Sub mnuEdit_Disease_Click()
    Dim lng����ID As Long, lng��ҳID As Long, int���� As Integer
    
    If mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
        int���� = Nvl(mrsFamily!����, 0)
    ElseIf mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
        int���� = Nvl(mrsBeds!����, 0)
    ElseIf mobjLVW Is lvwOut_s Then
        lng����ID = mrsOut!����ID
        lng��ҳID = mrsOut!��ҳID
        int���� = Nvl(mrsOut!����, 0)
    End If
    Call ExecPatiChange(EFun.Eҽ������ѡ��, Me, mstrPrivs, lng����ID, lng��ҳID, int����)
End Sub

Private Sub mnuEdit_Level_Click()
    
    Call ExecPatiChange(EFun.E���Ĵ�λ�ȼ�, Me, mstrPrivs, mrsBeds!����ID, mrsBeds!��ҳID, mrsBeds!����)
    
    If gblnOK Then Call LoadList(True, False, False, False)
End Sub

Private Sub mnuEdit_Change_Click()
    Dim lng����ID As Long, lng��ҳID As Long, strBeds As String
    
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    Else
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    End If
    
    Call ExecPatiChange(EFun.Eת��, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID)
    
    If gblnOK Then Call LoadList(True, True, True, False)
End Sub

Private Sub mnuEdit_In_Click()
    Dim strBeds As String, byt��Ʒ�ʽ As Byte, lng��λ����ID As Long
    Dim lng����ID As Long, lng��ҳID As Long
    
    If lvwBeds_s.SelectedItem Is Nothing Then
        strBeds = ""
    ElseIf lvwBeds_s.SelectedItem.Tag <> "�մ�" Then
        If mrsBeds!����ID = mrsIn!����ID Then '����ԭס��
            strBeds = Trim(Mid(lvwBeds_s.SelectedItem.Key, 2))
            lng��λ����ID = Val("" & mrsIn!��ס����id)
        Else
            strBeds = ""
        End If
    ElseIf Not (mrsBeds!�Ա���� = "���޴�" Or (mrsBeds!�Ա���� = "�д�" And "" & mrsIn!�Ա� = "��") _
        Or (mrsBeds!�Ա���� = "Ů��" And "" & mrsIn!�Ա� = "Ů")) Then
        strBeds = ""
    Else
        strBeds = Trim(Mid(lvwBeds_s.SelectedItem.Key, 2))
        lng��λ����ID = Val("" & mrsBeds!����ID)
    End If
    byt��Ʒ�ʽ = Val(lvwIn_s.SelectedItem.Tag)
    lng����ID = mrsIn!����ID
    lng��ҳID = mrsIn!��ҳID
    
    If byt��Ʒ�ʽ <> 2 Then
        Call ExecPatiChange(EFun.E���, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID, strBeds, lng��λ����ID, byt��Ʒ�ʽ)
    ElseIf byt��Ʒ�ʽ = 2 Then
        Call ExecPatiChange(EFun.E�벡��, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID, strBeds, lng��λ����ID)
    End If
    '��ƺ�λ
    If gblnOK Then
        If strBeds <> "" Then
            If InStr(strBeds, ",") Then
                strBeds = Split(strBeds, ",")(0)
                lvwBeds_s.ListItems("_" & strBeds).Selected = True
            Else
                lvwBeds_s.ListItems("_" & strBeds).Selected = True
            End If
        End If
        Call LoadList(True, True, True, False, True)
    End If
End Sub

Private Sub mnuEdit_Move_Click()
    Call ChangeBeds(0)
End Sub

Private Sub ChangeBeds(ByVal bytFun As Byte, Optional ByVal strĿ�괲�� As String)
'����:bytFun:0-����,1-����
    Dim lng����ID As Long, lng��ҳID As Long
    
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Or bytFun = 1 Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    Else
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    End If
        
    Call ExecPatiChange(EFun.E����, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID, bytFun, strĿ�괲��, "")
        
    If gblnOK Then
        If strĿ�괲�� <> "" Then
            On Error Resume Next '����Ŀ�겡�������ڱ�����û�У����ܴ���
            If InStr(strĿ�괲��, ",") > 0 Then '��������
                strĿ�괲�� = Split(strĿ�괲��, ",")(0)
                lvwBeds_s.ListItems("_" & strĿ�괲��).Selected = True
            Else
                lvwBeds_s.ListItems("_" & strĿ�괲��).Selected = True
            End If
            Err.Clear
        End If
    
        Call LoadList(True, True, False, False)
    End If
End Sub

Private Sub ChangeUnit()
    Dim lng����ID As Long, lng��ҳID As Long
    
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    Else
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    End If
        
    Call ExecPatiChange(EFun.Eת����, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID)
        
    If gblnOK Then
'
'            On Error Resume Next '����Ŀ�겡�������ڱ�����û�У����ܴ���
'            If InStr(strĿ�괲��, ",") > 0 Then '��������
'                strĿ�괲�� = Split(strĿ�괲��, ",")(0)
'                lvwBeds_s.ListItems("_" & strĿ�괲��).Selected = True
'            Else
'                lvwBeds_s.ListItems("_" & strĿ�괲��).Selected = True
'            End If
'            Err.Clear
'        End If
'
        Call LoadList(True, True, True, True)
    End If
End Sub

Private Sub SwapBeds(Optional ByVal strĿ�괲�� As String)
'###########################################################################################################
'## ���ܣ�ͬ�������˴�λ�Ի�
'## ���������˴�λ�Ի���Ŀ�괲�ţ���ѡ
'##
'###########################################################################################################
    
    Dim lng����ID As String, lng��ҳID As String, str���� As String
    Dim rsTmp As ADODB.Recordset
    
    lng����ID = mrsBeds!����ID
    lng��ҳID = mrsBeds!��ҳID
    str���� = mrsBeds!����
    
    If Trim(strĿ�괲��) <> "" Then
        '�϶�����Ŀ�괲λ���뵱ǰ��λ����ͬʱ��ִ���κβ���
        If str���� = strĿ�괲�� Then Exit Sub
        'Ŀ�괲λ������Ϣ�뵱ǰ��λ������Ϣ��ͬ��ִ���κβ���(��������)
        If lng����ID = mrsCBeds!����ID And lng��ҳID = mrsCBeds!��ҳID Then Exit Sub
    End If
    
    If ExecPatiChange(EFun.E��λ�Ի�, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID, str����, strĿ�괲��) Then
        If strĿ�괲�� <> "" Then
            On Error Resume Next '����Ŀ�겡�������ڱ�����û�У����ܴ���
            If InStr(strĿ�괲��, ",") > 0 Then '��������
                strĿ�괲�� = Split(strĿ�괲��, ",")(0)
                lvwBeds_s.ListItems("_" & strĿ�괲��).Selected = True
            Else
                lvwBeds_s.ListItems("_" & strĿ�괲��).Selected = True
            End If
            Err.Clear
        End If
    
        Call LoadList(True, True, False, False)
    End If
End Sub

Private Sub mnuEdit_Nurse_Click()
    On Error Resume Next
    Err.Clear
    
    frmNurse.mblnBed = (mobjLVW Is lvwBeds_s)
    frmNurse.Show 1, Me
    If gblnOK Then Call LoadList(True, True, False, False, True)
End Sub

Private Sub mnuEdit_Out_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    
    If mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    Else
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    End If
    
    Call ExecPatiChange(EFun.E��Ժ, Me, mstrPrivs, lng����ID, lng��ҳID)
    
    If gblnOK Then Call LoadList(True, True, False, True, True)
End Sub

Private Sub mnuEdit_PreOut_Click()
'���ܣ�����Ԥ��Ժ
    Dim lng����ID As Long, lng��ҳID As Long, str���� As String
    Dim blnTrue As Boolean
    
    If mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
        str���� = mrsFamily!����
    Else
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
        str���� = mrsBeds!����
    End If
    '--55791:������,2012-11-13,���ϳ�Ժҽ�����ܳ�����Ժ
    On Error Resume Next
    blnTrue = frmPreOut.ShowMe(Me, lng����ID, lng��ҳID, str����, mstrPrivs)
    
    If blnTrue = True Then
        Call LoadList(True, True, False, False)
    End If
End Sub

Private Sub mnuEdit_Recalc_Click()
    Dim lng����ID As Long, lng��ҳID As Long, str���� As String
    Dim rsTmp As ADODB.Recordset
    
    If mobjLVW Is lvwBeds_s Then
        Set rsTmp = mrsBeds
    ElseIf mobjLVW Is lvwFamily_s Then
        Set rsTmp = mrsFamily
    ElseIf mobjLVW Is lvwOut_s Then
        Set rsTmp = mrsOut
    ElseIf mobjLVW Is lvwIn_s Then
        Set rsTmp = mrsIn
    Else
        Exit Sub
    End If
    
    lng����ID = rsTmp!����ID
    lng��ҳID = rsTmp!��ҳID
    str���� = rsTmp!����
        
    gblnOK = False
    Call ExecPatiChange(EFun.E�������, Me, mstrPrivs, lng����ID, lng��ҳID, str����)
    
    If gblnOK Then stbThis.Panels(2).Text = "������������ɹ����!"
End Sub

Private Sub mnuEdit_Undo_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    Dim int���� As Integer
    
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
        int���� = Nvl(mrsBeds!����, 0)
    ElseIf mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
        int���� = Nvl(mrsFamily!����, 0)
    ElseIf mobjLVW Is lvwOut_s Then
        lng����ID = mrsOut!����ID
        lng��ҳID = mrsOut!��ҳID
        int���� = Nvl(mrsOut!����, 0)
    ElseIf mobjLVW Is lvwIn_s Then
        Exit Sub
    End If
    
    gblnOK = False
    Call ExecPatiChange(EFun.E����, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID, int����, CStr(tbr.Buttons("Undo").ButtonMenus(1).Text))
    
    If gblnOK Then Call LoadList(True, True, True, True, True)
End Sub

Private Sub mnuEditToInPati_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strסԺ�� As String, str���� As String
        
    If MsgBox("ȷʵҪ����סԺ���۲���תΪסԺ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
        
        strסԺ�� = IIf(IsNull(mrsBeds!סԺ��), "", mrsBeds!סԺ��)
        str���� = mrsBeds!����
    ElseIf mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
        
        strסԺ�� = IIf(IsNull(mrsFamily!סԺ��), "", mrsFamily!סԺ��)
        str���� = mrsFamily!����
    End If
    gblnOK = False
    Call ExecPatiChange(EFun.EתΪסԺ, Me, mstrPrivs, lng����ID, lng��ҳID, strסԺ��, str����)
    
    If gblnOK Then Call LoadList(True, True, False, False)
End Sub

Private Sub mnuFilePrintCard_Click()
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    If mobjLVW.SelectedItem Is Nothing Then Exit Sub
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_2", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_2", Me, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, 2)
    End If
End Sub

Private Sub mnuFilePrintMed_Click()
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    ElseIf mobjLVW Is lvwIn_s Then
        lng����ID = mrsIn!����ID
        lng��ҳID = mrsIn!��ҳID
    ElseIf mobjLVW Is lvwOut_s Then
        lng����ID = mrsOut!����ID
        lng��ҳID = mrsOut!��ҳID
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_1", Me) Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_1", Me, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, 2)
    End If
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub
Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuQuery_Log_Click()
    If mobjLVW Is lvwBeds_s Then
        frmHistory.mlng����ID = mrsBeds!����ID
        frmHistory.mlng��ҳID = mrsBeds!��ҳID
    ElseIf mobjLVW Is lvwOut_s Then
        frmHistory.mlng����ID = mrsOut!����ID
        frmHistory.mlng��ҳID = mrsOut!��ҳID
    ElseIf mobjLVW Is lvwFamily_s Then
        frmHistory.mlng����ID = mrsFamily!����ID
        frmHistory.mlng��ҳID = mrsFamily!��ҳID
    Else
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear
    
    frmHistory.Show 1, Me
End Sub


Private Sub mnuQueryInfo_Click()
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    
    If mobjLVW Is lvwBeds_s Then
        lng����ID = mrsBeds!����ID
        lng��ҳID = mrsBeds!��ҳID
    ElseIf mobjLVW Is lvwFamily_s Then
        lng����ID = mrsFamily!����ID
        lng��ҳID = mrsFamily!��ҳID
    ElseIf mobjLVW Is lvwIn_s Then
        lng����ID = mrsIn!����ID
        lng��ҳID = mrsIn!��ҳID
    Else
        lng����ID = mrsOut!����ID
        lng��ҳID = mrsOut!��ҳID
    End If
    
    On Error Resume Next
    Err.Clear
    
    If CreatePublicPatient() Then
        Call gobjPublicPatient.ReadPatiDegreeCard(Me, lng����ID, lng��ҳID)
    End If
    
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng����ID As Long
    
    If lvwBeds_s Is Me.ActiveControl And Not lvwBeds_s.SelectedItem Is Nothing Then
        lng����ID = Val("" & mrsBeds!����ID)
    ElseIf lvwFamily_s Is Me.ActiveControl And Not lvwFamily_s.SelectedItem Is Nothing Then
        lng����ID = Val("" & mrsFamily!����ID)
    ElseIf lvwOut_s Is Me.ActiveControl And Not lvwOut_s.SelectedItem Is Nothing Then
        lng����ID = Val("" & mrsOut!����ID)
    ElseIf lvwIn_s Is Me.ActiveControl And Not lvwIn_s.SelectedItem Is Nothing Then
        lng����ID = Val("" & mrsIn!����ID)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & mlngUnit, "����ID=" & lng����ID)
End Sub

Private Sub mnuView_Card_Click()
    mnuView_Card.Checked = Not mnuView_Card.Checked
    picCard_s.Visible = mnuView_Card.Checked
    If picCard_s.Visible Then
        If picCard_s.Left <= -picCard_s.width Or picCard_s.Top <= -picCard_s.Height Then
            With picCard_s
                .Left = picVsc.Left - (picCard_s.width - picVsc.width) / 2
                .Top = lvwBeds_s.Top + (lvwBeds_s.Height - picCard_s.Height) / 2
            End With
        End If
    End If
End Sub

Private Sub mnuView_ListView_Click(Index As Integer)
    Call SetView(CByte(Index))
End Sub

Private Sub mnuViewColSel_Click()
    Select Case mobjLVW.Name
        Case "lvwBeds_s"
            If zlControl.LvwSelectColumns(lvwBeds_s, COL_BEDS) Then
                LoadList True, False, False, False
            End If
        Case "lvwFamily_s"
            If zlControl.LvwSelectColumns(lvwFamily_s, COL_FAMILY) Then
                LoadList False, True, False, False
            End If
        Case "lvwIn_s"
            If zlControl.LvwSelectColumns(lvwIn_s, COL_IN) Then
                LoadList False, False, True, False
            End If
        Case "lvwOut_s"
            If zlControl.LvwSelectColumns(lvwOut_s, COL_OUT) Then
                LoadList False, False, False, True
            End If
    End Select
End Sub

Private Sub mnuViewreFlash_Click()
    Call LoadList
    Me.Refresh
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mnuViewToolUnit_Click()
    mnuViewToolUnit.Checked = Not mnuViewToolUnit.Checked
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = False
    cbr.Bands(2).Visible = Not cbr.Bands(2).Visible
    If mnuViewToolButton.Checked Then cbr.Bands(1).Visible = True
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Bands(1).Visible = Not cbr.Bands(1).Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = cbr.Bands(2).Visible Or cbr.Bands(1).Visible
    Form_Resize
End Sub

Private Sub picCard_s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Set picCard_s.MouseIcon = img32.ListImages("Down").Picture
        Call MoveObj(picCard_s.hWnd)
        Set picCard_s.MouseIcon = img32.ListImages("Up").Picture
        mobjLVW.SetFocus
    ElseIf Button = 2 Then
        PopupMenu mnuEdit, 2
    End If
End Sub

Private Sub picCard_s_OLECompleteDrag(Effect As Long)
    mobjLVW.SetFocus
End Sub
Private Sub picVsc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreX = X: mblnDownVsc = True
End Sub

Private Sub picVsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnDownVsc Then
        If lvwBeds_s.width + X - mlngPreX < 1500 Or lvwFamily_s.width - (X - mlngPreX) < 1000 Then Exit Sub
        picVsc.Left = picVsc.Left + X - mlngPreX
        lblBed.width = lblBed.width + X - mlngPreX
        lvwBeds_s.width = lvwBeds_s.width + X - mlngPreX
        lblIn.width = lblIn.width + X - mlngPreX
        lvwIn_s.width = lvwIn_s.width + X - mlngPreX
        lblFamily.Left = lblFamily.Left + X - mlngPreX
        lblFamily.width = lblFamily.width - (X - mlngPreX)
        lvwFamily_s.Left = lvwFamily_s.Left + X - mlngPreX
        lvwFamily_s.width = lvwFamily_s.width - (X - mlngPreX)
        PicOut.Left = PicOut.Left + X - mlngPreX
        PicOut.width = PicOut.width - (X - mlngPreX)
        lvwOut_s.Left = lvwOut_s.Left + X - mlngPreX
        lvwOut_s.width = lvwOut_s.width - (X - mlngPreX)
        Me.Refresh
    End If
End Sub

Private Sub picVsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDownVsc = False
    mobjLVW.SetFocus
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PatiColor" Then
        zlDatabase.ShowPatiColorTip Me
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "View"
            Call SetView((mobjLVW.View + 1) Mod 4)
        Case "In"
            mnuEdit_In_Click
        Case "Change"
            mnuEdit_Change_Click
        Case "Out"
            mnuEdit_Out_Click
        Case "Move"
            mnuEdit_Move_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Adjust"
            mnuEdit_Adjust_Click
        Case "Undo"
            If mnuEdit_Undo.Enabled And mnuEdit_Undo.Visible Then mnuEdit_Undo_Click
    End Select
End Sub

Private Sub SetView(bytStyle As Byte)
'���ܣ������б���ʾ��ʽ
'������bytstyle=0-��ͼ��,1-Сͼ��,2-�б�,3-��ϸ����
    mnuView_ListView(0).Checked = False
    mnuView_ListView(1).Checked = False
    mnuView_ListView(2).Checked = False
    mnuView_ListView(3).Checked = False
    mnuView_ListView(bytStyle).Checked = True
    mobjLVW.View = bytStyle
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Icon"
            Call SetView(0)
        Case "Small"
            Call SetView(1)
        Case "List"
            Call SetView(2)
        Case "Detail"
            Call SetView(3)
        Case Else
            mnuEdit_Undo_Click
    End Select
End Sub

Private Sub lvwBeds_s_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwBeds_s.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwBeds_s.SortOrder = lvwDescending
    Else
        lvwBeds_s.SortOrder = lvwAscending
    End If
    lvwBeds_s.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwBeds_s.SelectedItem Is Nothing Then lvwBeds_s.SelectedItem.EnsureVisible
End Sub

Private Sub lvwFamily_s_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwFamily_s.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwFamily_s.SortOrder = lvwDescending
    Else
        lvwFamily_s.SortOrder = lvwAscending
    End If
    lvwFamily_s.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwFamily_s.SelectedItem Is Nothing Then lvwFamily_s.SelectedItem.EnsureVisible
End Sub

Private Sub lvwIn_s_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwIn_s.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwIn_s.SortOrder = lvwDescending
    Else
        lvwIn_s.SortOrder = lvwAscending
    End If
    lvwIn_s.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwIn_s.SelectedItem Is Nothing Then lvwIn_s.SelectedItem.EnsureVisible
End Sub

Private Sub lvwOut_s_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwOut_s.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwOut_s.SortOrder = lvwDescending
    Else
        lvwOut_s.SortOrder = lvwAscending
    End If
    lvwOut_s.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwOut_s.SelectedItem Is Nothing Then lvwOut_s.SelectedItem.EnsureVisible
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub SetFocusColor()
    Dim i As Integer
    
    '���õ�ǰ�б�ͻ����ʾ
    lblBed.BackColor = COLOR_LOST
    lblFamily.BackColor = COLOR_LOST
    lblIn.BackColor = COLOR_LOST
    PicOut.BackColor = COLOR_LOST
        chk����(0).BackColor = COLOR_LOST
        chk����(1).BackColor = COLOR_LOST
        lblOut.BackColor = COLOR_LOST
    Select Case mobjLVW.Name
        Case "lvwBeds_s"
            lblBed.BackColor = COLOR_FOCUS
        Case "lvwFamily_s"
            lblFamily.BackColor = COLOR_FOCUS
        Case "lvwIn_s"
            lblIn.BackColor = COLOR_FOCUS
        Case "lvwOut_s"
            PicOut.BackColor = COLOR_FOCUS
            chk����(0).BackColor = COLOR_FOCUS
            chk����(1).BackColor = COLOR_FOCUS
            lblOut.BackColor = COLOR_FOCUS
    End Select
    '��ȡ��ǰ�б���ʾ��ʽ
    mnuView_ListView(0).Checked = False
    mnuView_ListView(1).Checked = False
    mnuView_ListView(2).Checked = False
    mnuView_ListView(3).Checked = False
    mnuView_ListView(mobjLVW.View).Checked = True
    
    If Not mobjLVW.SelectedItem Is Nothing Then mobjLVW.SelectedItem.EnsureVisible
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ����
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, blnLimitUnit As Boolean, strUnitIDs As String
    
    On Error GoTo errH
    
    blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
    If blnLimitUnit Then
        strUnitIDs = "," & GetUserUnits(False) & ","
    Else
        strUnitIDs = "," & GetUserUnits(True) & ","
    End If
    'by lesfeng 2010-01-12 �����Ż�
    'Ŀǰ��������۲���
    gstrSQL = _
        " Select A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where B.����ID = A.ID" & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.������� IN(1,2,3) And B.��������='����'" & _
        IIf(blnLimitUnit, " And instr([1],',' || A.ID || ',')>0 ", "") & _
        " And (A.վ��=[2] Or A.վ�� is Null)" & _
        " Order by A.����"
        '
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strUnitIDs, gstrNodeNo)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = UserInfo.����ID And cboUnit.ListIndex = -1 Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
        If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then cboUnit.ListIndex = 0  '����Click�¼�
    ElseIf InStr(";" & mstrPrivs, "���в���") > 0 Then
        MsgBox "û�����ò���,�뵽���Ź��������ù�������Ϊ����Ĳ��ţ�", vbExclamation, gstrSysName
        Exit Function
    Else
        MsgBox "��û�� [���в���] ��Ȩ��,���������ڲ��Ų��ǲ��������ڲ�����", vbExclamation, gstrSysName
        Exit Function
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadBedsMap(lngUnitID As Long) As Boolean
'���ܣ���ȡָ�������Ĵ�λӳ���(����λ��Ϣ�����������Ϣ��������Ժ��Ϣ),����ʾ���б���
    Dim i As Integer, j As Integer, strIcon As String
    Dim objItem As ListItem, blnChange As Boolean
    Dim bytLen As Byte, strChange As String
    Dim strTmp As String
    Dim k As Integer
    Dim strTemp As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '��������
    strTmp = ""
    gstrSQL = ""
    For i = 1 To tbrFilter.Buttons.Count
        If tbrFilter.Buttons(i).Key Like "Nurse*" And tbrFilter.Buttons(i).Image = "Check" Then
            strTmp = strTmp & "," & Val(Replace(tbrFilter.Buttons(i).Key, "Nurse", ""))
        End If
    Next
    strTmp = strTmp & ","
    gstrSQL = " And (instr([2],',' || B.����ȼ�ID || ',')>0  Or B.����ȼ�ID is NULL)"
    
    If tbrFilter.Buttons("curDay").Image = "Check_" Then
        gstrSQL = gstrSQL & " And B.��Ժ���� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60"
    End If
    
    gstrSQL = _
        "Select A.����, A.����id, A.����, A.�����, A.�Ա����, A.��λ����, A.��λ�ȼ�id, A.��λ�ȼ�, A.״̬, A.����, B.��ҳid," & vbNewLine & _
        "       Nvl(B.״̬, 0) As ����״̬, B.��ǰ����id, B.��ǰ����, B.����id, B.סԺ��, B.����, B.�Ա�, B.����, B.ҽ�Ƹ��ʽ," & vbNewLine & _
        "       B.��ͬ��λid, B.��ǰ����, To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.����ȼ�id, B.����ȼ�," & vbNewLine & _
        "       B.סԺҽʦ, B.��������, B.����,B.��ǰ���� as ��Ҫ����,B.���￨��,B.���֤��,B.IC����,B.��������" & vbNewLine & _
        "From (Select A.����,A.˳���, A.����id, Nvl(C.����, Decode(A.����, 1, '<���ò���>', Null)) As ����, A.�����, A.�Ա����," & vbNewLine & _
        "              A.��λ����, A.�ȼ�id As ��λ�ȼ�id, B.���� As ��λ�ȼ�, A.״̬, A.����" & vbNewLine & _
        "       From ��λ״����¼ A, �շ���ĿĿ¼ B, ���ű� C" & vbNewLine & _
        "       Where A.����id = C.ID(+) And A.�ȼ�id = B.ID(+) And A.����id = [1]) A," & vbNewLine & _
        "     (Select Distinct B.��ҳid, B.״̬, B.��Ժ����id As ��ǰ����id, E.���� As ��ǰ����, A.����id, B.סԺ��, C.����, NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�," & vbNewLine & _
        "              NVL(B.����,A.����) ����, A.ҽ�Ƹ��ʽ, A.��ͬ��λid, B.��ǰ����, B.��Ժ����, B.����ȼ�id, D.���� As ����ȼ�, B.סԺҽʦ," & vbNewLine & _
        "              B.��������, B.����,A.��ǰ����,A.���￨��,A.���֤��,A.IC����,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
        "       From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, �շ���ĿĿ¼ D, ���ű� E, ��λ״����¼ F" & vbNewLine & _
        "       Where B.����id = A.����id And C.����id = B.����id And C.��ҳid = B.��ҳid And B.����ȼ�id = D.ID(+) " & gstrSQL & " And" & vbNewLine & _
        "             B.��Ժ����id = E.ID And B.��Ժ���� Is Null And Nvl(B.��ҳid, 0) <> 0 And Nvl(B.״̬, 0) In (0, 2, 3) And" & vbNewLine & _
        "             C.��ʼʱ�� Is Not Null And C.��ֹʱ�� Is Null And C.���� Is Not Null And F.����id = B.����id And" & vbNewLine & _
        "             F.����id = [1] And F.����id Is Not Null) B" & vbNewLine & _
        "Where A.���� = B.����(+)" & vbNewLine & _
        "Order By A.˳���,LPad(A.����, 10, ' ')"
    Set mrsBeds = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUnitID, strTmp)
    Set mrsCBeds = mrsBeds.Clone
    
    mintBeds_A = 0: mintHolding = 0: mintChange_A = 0
    
    With mrsBeds
        If Not .EOF Then
            bytLen = GetMaxBedLen(lngUnitID)
            For i = 1 To .RecordCount
                If Not (!״̬ = "ռ��" And IsNull(!����ID)) Then
                    blnChange = False
                    If !����״̬ = 2 Then 'ת�Ʋ���
                        blnChange = True
                        If Not IsNull(!����ID) Then
                            If InStr(strChange & ",", "," & !����ID & ",") = 0 Then
                                strChange = strChange & "," & !����ID
                            End If
                        End If
                    End If
'
'                    If Not (IsNull(!����id) And IsNull(!��ҳid)) Then
'                        gstrSQL = "Select 1 From ���˱䶯��¼ Where ����id=[1] And ��ҳid=[2] And ��ʼʱ�� Is Null And ��ֹʱ�� Is Null And ��ʼԭ��=15 "
'                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(!����id), Val(!��ҳid))
'
'                        If rsTmp.RecordCount > 0 Then
'                            blnChange = True
'                        End If
'                    End If
                    
                    If blnChange Then
                        strIcon = "Change"
                    Else
                        Select Case !״̬
                            Case "�մ�"
                                If !�Ա���� = "�д�" Then
                                    strIcon = "M_Empty"
                                ElseIf !�Ա���� = "Ů��" Then
                                    strIcon = "F_Empty"
                                Else
                                    strIcon = "Empty"
                                End If
                            Case "����"
                                strIcon = "Remedy"
                            Case "ռ��"
                                If !����״̬ = 3 Then
                                    strIcon = "Out" 'Ԥ��Ժ����
                                Else
                                    strIcon = "Holding"
                                End If
                                mintHolding = mintHolding + 1
                        End Select
                    End If
                    
                    '���۲���ͼ��
                    If IIf(IsNull(!��������), 0, !��������) <> 0 Then strIcon = "K" & strIcon
                    
                    '�Ӵ��ķǱ�ͼ��
                    If IIf(IsNull(!��λ����), "", !��λ����) = "�Ӵ�" Then
                        strIcon = "�Ӵ�_" & strIcon
                    ElseIf IIf(IsNull(!��λ����), "", !��λ����) = "�Ǳ�" Then
                        strIcon = "�Ǳ�_" & strIcon
                    End If
                    '���ò�����ʾ
                    If IIf(IsNull(!����), 0, !����) = 1 Then
                        strIcon = "����_" & strIcon
                    End If
                    '����29710 by lesfeng 2010-05-12 ǿ�ƽ������ûص�һ�У���Ϊ��������ʱ����
                    strTemp = lvwBeds_s.ColumnHeaders(1).Text
                    For k = 1 To lvwBeds_s.ColumnHeaders.Count
                        If lvwBeds_s.ColumnHeaders(k).Text = "����" Then Exit For
                    Next
                    If k <> 1 Then
                        lvwBeds_s.ColumnHeaders(1).Text = "����"
                        lvwBeds_s.ColumnHeaders(k).Text = strTemp
                        lvwBeds_s.ColumnHeaders(1).Key = "_����1"
                        lvwBeds_s.ColumnHeaders(k).Key = "_" & strTemp
                        lvwBeds_s.ColumnHeaders(1).Key = "_����"
                    End If
                    
                    '�Դ�λΪ��λ��ʾ,�Դ���Ϊ�ؼ���
                    Set objItem = lvwBeds_s.ListItems.Add(, "_" & !����, Space(bytLen - Len(!����)) & !���� & IIf(IsNull(!����), "", ":" & !����), strIcon, strIcon)
                   
                    objItem.ForeColor = GetPatiColor(Nvl(mrsBeds!��������, "��ͨ����"))
                    
                    For j = 2 To lvwBeds_s.ColumnHeaders.Count
                        objItem.SubItems(j - 1) = IIf(IsNull(mrsBeds.Fields(lvwBeds_s.ColumnHeaders(j).Text).Value), "", mrsBeds.Fields(lvwBeds_s.ColumnHeaders(j).Text).Value)
                        objItem.ListSubItems(j - 1).ForeColor = objItem.ForeColor
                    Next
                    objItem.Tag = !״̬ '��Tag��־��λ״̬
                    
                    mintBeds_A = mintBeds_A + 1
                End If
                .MoveNext
            Next
            mintChange_A = UBound(Split(Mid(strChange, 2), ",")) + 1
            
            If Not lvwBeds_s.SelectedItem Is Nothing Then
                lvwBeds_s.ListItems(1).Selected = True
                lvwBeds_s.SelectedItem.EnsureVisible
            End If
        End If
    End With
    ReadBedsMap = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReadNure(lngUnitID As Long) As Boolean
    Dim strTmp As String
    Dim rsNure As New ADODB.Recordset
    Dim i As Integer, lngLen As Integer
    Dim objButton As Button
    
    On Error GoTo errH
    strTmp = "Select c.Id, c.����, Count(ID) As ����" & vbNewLine & _
        "From ��Ժ���� A, ������ҳ B, �շ���ĿĿ¼ C" & vbNewLine & _
        "Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.����ȼ�id = c.Id And a.����id =[1] And b.״̬ In (0, 2, 3)" & vbNewLine & _
        "Group By c.Id, c.����"
        
    Set rsNure = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, lngUnitID)
    tbrFilter.Buttons.Clear
    Set objButton = tbrFilter.Buttons.Add(, "curDay", "������Ժ", , "UnCheck_")
    If rsNure.RecordCount <> 0 Then
        For i = 1 To rsNure.RecordCount
            If LenB(rsNure!����) > lngLen Then lngLen = LenB(rsNure!����)
            rsNure.MoveNext
        Next
        rsNure.MoveFirst
    End If
    With rsNure
        If Not .EOF Then
            For i = 1 To .RecordCount
                Set objButton = tbrFilter.Buttons.Add(, "Nurse" & !ID, GetLenText(!����, lngLen) & "(" & !���� & ")", , "Check")
                 If i <= 10 Then
                    objButton.ToolTipText = !���� & "����(ALT + " & i Mod 10 & ")"
                End If
                .MoveNext
            Next
        End If
    End With
    tbrFilter.Buttons(1).Caption = GetLenText(tbrFilter.Buttons(1).Caption, lngLen)
    ReadNure = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ReadFamily(lngUnitID As Long) As Boolean
'���ܣ���ȡָ�������ļ�ͥ�������˲���ʾ���б���
'˵������ͥ��������Ϊ��,������ס�˵�
    Dim i As Integer, j As Integer, objItem As ListItem
    Dim strChange As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    '��������
    strTmp = ""
    gstrSQL = ""
    For i = 1 To tbrFilter.Buttons.Count
        If tbrFilter.Buttons(i).Key Like "Nurse*" And tbrFilter.Buttons(i).Image = "Check" Then
            strTmp = strTmp & "," & Val(Replace(tbrFilter.Buttons(i).Key, "Nurse", ""))
        End If
    Next
    strTmp = strTmp & ","
    gstrSQL = " And (instr([2],','|| B.����ȼ�ID || ',')>0 Or B.����ȼ�ID is NULL)"
    
    If tbrFilter.Buttons("curDay").Image = "Check_" Then
        gstrSQL = gstrSQL & " And B.��Ժ���� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60"
    End If
    
    '58842,������,2013-02-25,��Ժ���˶�ȡ(����Ժ�����ж�ȡ)
    gstrSQL = _
       "Select Nvl(B.״̬, 0) As ����״̬, B.��Ժ����id As ��ǰ����id, E.���� As ��ǰ����, A.����id, B.סԺ��, NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�," & vbNewLine & _
        "       NVL(B.����,A.����) ����, A.ҽ�Ƹ��ʽ, A.��ͬ��λid, B.��ҳid, B.��ǰ����," & vbNewLine & _
        "       To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.����ȼ�id, D.���� As ����ȼ�, B.סԺҽʦ," & vbNewLine & _
        "       B.��������, B.����,A.���￨��,A.���֤��,A.IC����,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
        "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, �շ���ĿĿ¼ D, ���ű� E,��Ժ���� F" & vbNewLine & _
        "Where B.����id = A.����id And F.����ID=A.����ID And C.����id = B.����id And C.��ҳid = B.��ҳid And B.����ȼ�id = D.ID(+) And" & vbNewLine & _
        "      B.��Ժ����id = E.ID And Nvl(B.��ҳid, 0) <> 0 And Nvl(B.״̬, 0) In (0, 2, 3) And" & vbNewLine & _
        "      C.��ʼʱ�� Is Not Null And C.��ֹʱ�� Is Null And C.���� Is Null And B.��ǰ����id+0 = F.����ID And F.����ID=[1] And B.��Ժ���� Is Null" & gstrSQL & vbNewLine & _
        "Order By B.��Ժ���� Desc, B.סԺ�� Desc"

    Set mrsFamily = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUnitID, strTmp)
    Set mrsCFamily = mrsFamily.Clone
    
    mintBeds_B = 0: mintChange_B = 0
    
    With mrsFamily
        If .RecordCount <> 0 Then
            For i = 1 To .RecordCount
                '�Բ���Ϊ��λ��ʾ,�Բ���IDΪ�ؼ���
                If !����״̬ = 2 Then
                    'ת�Ʋ���
                    If Nvl(!��������, 0) <> 0 Then
                        '���۲���ͼ��
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !����ID, !����, "KChange", "KChange")
                    Else
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !����ID, !����, "Change", "Change")
                    End If
                    If Not IsNull(!����ID) Then
                        If InStr(strChange & ",", "," & !����ID & ",") = 0 Then
                            strChange = strChange & "," & !����ID
                        End If
                    End If
                ElseIf !����״̬ = 3 Then
                    'Ԥ��Ժ����
                    If Nvl(!��������, 0) <> 0 Then
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !����ID, !����, "KOut", "KOut")
                    Else
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !����ID, !����, "Out", "Out")
                    End If
                Else
                    If Nvl(!��������, 0) <> 0 Then
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !����ID, !����, "KFamily", "KFamily")
                    Else
                        Set objItem = lvwFamily_s.ListItems.Add(, "_" & !����ID, !����, "Family", "Family")
                    End If
                End If
                
                objItem.ForeColor = GetPatiColor(Nvl(mrsFamily!��������, "��ͨ����"))
                For j = 2 To lvwFamily_s.ColumnHeaders.Count
                    objItem.SubItems(j - 1) = IIf(IsNull(mrsFamily.Fields(lvwFamily_s.ColumnHeaders(j).Text).Value), "", mrsFamily.Fields(lvwFamily_s.ColumnHeaders(j).Text).Value)
                    objItem.ListSubItems(j - 1).ForeColor = objItem.ForeColor
                Next
                mintBeds_B = mintBeds_B + 1
                
                .MoveNext
            Next
            mintChange_B = UBound(Split(Mid(strChange, 2), ",")) + 1

            If Not lvwFamily_s.SelectedItem Is Nothing Then
                lvwFamily_s.ListItems(1).Selected = True
                lvwFamily_s.SelectedItem.EnsureVisible
            End If
        End If
    End With
    ReadFamily = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadIn(lngUnitID As Long) As Boolean
'���ܣ���ȡָ�������Ǽ�Ϊ��ǰ��������δ�Ǽǲ������Ǽǿ������ڵ�ǰ�����Ĵ���Ʋ���,����Ժ�Ǽǲ��˺�ת�Ʋ���,����ʾ���б���
    Dim objItem As ListItem, i As Integer, j As Integer
    Dim strSex As String, strpar1 As String, strpar2 As String
    Dim strDepts As String, lngInTime As Long
    
    On Error GoTo errH
    strDepts = zlDatabase.GetPara("����Ʋ��˿���", glngSys, mlngModul, "")
    If strDepts <> "" Then
        strDepts = "," & strDepts & ","
        strpar1 = " And Instr([2],',' || B.��Ժ����id || ',')>0 "
        strpar2 = " And Instr([2],',' || C.����ID || ',')>0 "
    End If
    lngInTime = Val(zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, 3))
    strpar1 = strpar1 & " And B.��Ժ����>=" & IIf(lngInTime <> 0, "Sysdate-[3]", "trunc(sysdate)")
    
    '��Ժ����(״̬=1),ʹ�õ�ǰ����ID,��Ժ����ID�Ա�ʹ����������ʹ����Ժ����ID,��Ժ����ID
    '����29002 by lesfeng 2010-04-09 ԭAnd C.����id = H.����id ��Ϊ And C.����id+0 = H.����id
    '58842,������,2013-02-25,��Ժ���˶�ȡ(����Ժ�����ж�ȡ)
    gstrSQL = _
        "Select 0 As ��Ʊ�־, A.����id, B.סԺ��, NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, NVL(B.����,A.����) ����, B.�ѱ�, B.��ҳid, B.��ǰ����id, E.���� As ��ǰ����," & vbNewLine & _
        "       B.��Ժ����id, F.���� As ��ǰ����, To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.��ǰ����," & vbNewLine & _
        "       B.����ȼ�id, D.���� As ����ȼ�, B.��Ժ����id As ��ס����id, F.���� As ת�����, B.���λ�ʿ, B.����ҽʦ," & vbNewLine & _
        "       B.סԺҽʦ, B.��������, B.����,A.���￨��,A.���֤��,A.IC����,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
        "From ������Ϣ A, ������ҳ B, �շ���ĿĿ¼ D, ���ű� E, ���ű� F" & vbNewLine & _
        "Where B.����id = A.����id And B.����ȼ�id = D.ID(+) And B.��ǰ����id = E.ID(+) And B.��Ժ����id = F.ID And" & vbNewLine & _
        "      B.��Ժ���� Is Null And Nvl(B.��ҳid, 0) <> 0 And B.״̬ = 1 And" & vbNewLine & _
        "      (B.��ǰ����ID+0 = [1] Or B.��ǰ����ID Is Null And Exists(Select 1 From �������Ҷ�Ӧ C Where B.��Ժ����id = C.����id And C.����id = [1]))" & strpar1
    '84937:������,�����Ż�
    'ת�Ʋ���(���ڿ�ʼʱ��Ϊ�յ���Ʊ䶯)
    gstrSQL = gstrSQL & vbNewLine & " Union All " & vbNewLine & _
        "Select 1 As ��Ʊ�־, A.����id, B.סԺ��, NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, NVL(B.����,A.����) ����, B.�ѱ�, B.��ҳid, B.��ǰ����id, E.���� As ��ǰ����," & vbNewLine & _
        "       B.��Ժ����id, F.���� As ��ǰ����, To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.��ǰ����," & vbNewLine & _
        "       B.����ȼ�id, D.���� As ����ȼ�, C.����id As ��ס����id, G.���� As ת�����, B.���λ�ʿ, B.����ҽʦ, B.סԺҽʦ," & vbNewLine & _
        "       B.��������, B.����,A.���￨��,A.���֤��,A.IC����,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
        "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, �շ���ĿĿ¼ D, ���ű� E, ���ű� F, ���ű� G, �������Ҷ�Ӧ H" & vbNewLine & _
        "Where A.��Ժ=1 And B.����id = A.����id And B.��ҳID=A.��ҳID And C.����id = B.����id And C.��ҳid = B.��ҳid And B.����ȼ�id = D.ID(+) And" & vbNewLine & _
        "      B.��ǰ����id+0 = E.ID And B.��Ժ����id+0 = F.ID And Nvl(B.��ҳid, 0) <> 0 And C.��ʼԭ�� = 3 And C.��ʼʱ�� Is Null And" & vbNewLine & _
        "      C.��ֹʱ�� Is Null And B.״̬ = 2 And C.����id = G.ID And C.����id+0 = H.����id And H.����id = [1] " & strpar2
        
    'ת��������(���ڿ�ʼʱ��Ϊ�յ��벡���䶯)
    gstrSQL = gstrSQL & vbNewLine & " Union All " & vbNewLine & _
        "Select 2 As ��Ʊ�־, A.����id, B.סԺ��, NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, NVL(B.����,A.����) ����, B.�ѱ�, B.��ҳid, B.��ǰ����id, E.���� As ��ǰ����," & vbNewLine & _
        "       B.��Ժ����id, F.���� As ��ǰ����, To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.��ǰ����," & vbNewLine & _
        "       B.����ȼ�id, D.���� As ����ȼ�, C.����id As ��ס����id, G.���� As ת�����, B.���λ�ʿ, B.����ҽʦ, B.סԺҽʦ," & vbNewLine & _
        "       B.��������, B.����,A.���￨��,A.���֤��,A.IC����,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
        "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, �շ���ĿĿ¼ D, ���ű� E, ���ű� F, ���ű� G, �������Ҷ�Ӧ H" & vbNewLine & _
        "Where A.��Ժ=1 And B.����id = A.����id And B.��ҳID=A.��ҳID And C.����id = B.����id And C.��ҳid = B.��ҳid And B.����ȼ�id = D.ID(+) And" & vbNewLine & _
        "      B.��ǰ����id+0 = E.ID And B.��Ժ����id+0 = F.ID And Nvl(B.��ҳid, 0) <> 0 And C.��ʼԭ�� = 15 And C.��ʼʱ�� Is Null And" & vbNewLine & _
        "      C.��ֹʱ�� Is Null And B.״̬ = 2 And C.����id = G.ID And  C.����id+0 = H.����id And C.����id+0 = H.����id And H.����id = [1] " & strpar2 & vbNewLine & _
        "Order By ��Ʊ�־ Desc, ��Ժʱ�� Desc, סԺ�� Desc"
    Set mrsIn = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUnitID, strDepts, lngInTime)
    Set mrsCIn = mrsIn.Clone
    
    mintIn = 0: mintChange_C = 0
    
    With mrsIn
        If mrsIn.RecordCount <> 0 Then
            For i = 1 To .RecordCount
                If IsNull(!�Ա�) Then
                    strSex = "O"
                Else
                    If InStr(!�Ա�, "��") > 0 Then
                        strSex = "M"
                    ElseIf InStr(!�Ա�, "Ů") > 0 Then
                        strSex = "F"
                    Else
                        strSex = "O"
                    End If
                End If
                
                '���۲���ͼ��
                If IIf(IsNull(!��������), 0, !��������) <> 0 Then strSex = "K" & strSex
                
                '�Բ���IDΪ�ؼ���
                If !��Ʊ�־ = 0 Then
                    Set objItem = lvwIn_s.ListItems.Add(, "_" & !����ID, !����, strSex, strSex)
                ElseIf !��Ʊ�־ = 1 Then
                    Set objItem = lvwIn_s.ListItems.Add(, "_" & !����ID, !����, strSex & "_Change", strSex & "_Change")
                    mintChange_C = mintChange_C + 1
                Else
                    Set objItem = lvwIn_s.ListItems.Add(, "_" & !����ID, !����, strSex & "_ChangeUnit", strSex & "_ChangeUnit")
                    'mintChange_C = mintChange_C + 1
                End If
                
                objItem.ForeColor = GetPatiColor(Nvl(mrsIn!��������, "��ͨ����"))
                For j = 2 To lvwIn_s.ColumnHeaders.Count
                    If Not (!��Ʊ�־ = 0 And lvwIn_s.ColumnHeaders(j).Text = "ת�����") Then
                        objItem.SubItems(j - 1) = IIf(IsNull(mrsIn.Fields(lvwIn_s.ColumnHeaders(j).Text).Value), "", mrsIn.Fields(lvwIn_s.ColumnHeaders(j).Text).Value)
                        objItem.ListSubItems(j - 1).ForeColor = objItem.ForeColor
                    End If
                Next
                objItem.Tag = !��Ʊ�־ '��Tag��־��Ʋ������
                
                mintIn = mintIn + 1
                
                .MoveNext
            Next
            
            If Not lvwIn_s.SelectedItem Is Nothing Then
                lvwIn_s.ListItems(1).Selected = True
                lvwIn_s.SelectedItem.EnsureVisible
            End If
        End If
    End With
    ReadIn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadOut(lngUnitID As Long) As Boolean
'���ܣ���ȡָ��������Ժ���˲���ʾ���б���
    Dim i As Integer, j As Integer, strSex As String
    Dim objItem As ListItem, int���� As Integer
    Dim lngOutTime As Long, strסԺ�� As String
    
    '��Ժ������ʾ����
    lngOutTime = Val(zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, "30"))

    '����δ�����
    If chk����(0).Value = 1 And chk����(1).Value = 1 Then
        int���� = 0               '����ʾ
    ElseIf chk����(0).Value = 0 And chk����(1).Value = 1 Then
        int���� = 1               'ֻ��ʾδ�����
    ElseIf chk����(0).Value = 1 And chk����(1).Value = 0 Then
        int���� = 2              'ֻ��ʾ�ѽ����
    End If
    
    '50323,������,2012-08-14,�жϲ����Ƿ��Ѿ����壬Ӧ���ж�ĳ��סԺ�Ƿ����δ����ķ��á�
    gstrSQL = " And B.��ǰ����ID+0=[1] And B.��Ժ����>=" & IIf(lngOutTime <> 0, "Sysdate-[2]", "trunc(Sysdate)")
    
    'ע��ԭ�д���
'    gstrSQL = _
'        "Select A.����id,NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�,A.סԺ����,B.סԺ��, NVL(B.����,A.����) ����, B.�ѱ�, B.��ҳid,B.��ǰ����ID," & vbNewLine & _
'        "       To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��," & vbNewLine & _
'        "       To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, D.���� As ��Ժ����,B.��Ժ����ID, B.��Ժ����, B.��ǰ���� As ��Ժ����," & vbNewLine & _
'        "       C.���� As ����ȼ�, B.��Ժ��ʽ, B.��������, B.����, Decode(Nvl(E.�������, 0), 0, '��', Null) As ����" & vbNewLine & _
'        "       ,A.���￨��,A.���֤��,A.IC����,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
'        "From ������Ϣ A, ������ҳ B, �շ���ĿĿ¼ C, ���ű� D, (select ����ID,����,Nvl(sum(Ԥ�����),0) Ԥ�����,Nvl(sum(�������),0) ������� from ������� group by ����ID,����) E" & vbNewLine & _
'        "Where B.����id = A.����id And B.��Ժ����id = D.ID And B.��Ժ���� Is Not Null And" & vbNewLine & _
'        "      Nvl(B.��ҳid, 0) <> 0 And B.����ȼ�id = C.ID(+) And A.����id = E.����id(+) And E.����(+) = 1" & gstrSQL
    '����µ�sql
    '84946��������,SQL�Ż�(���Բ���δ����õĲ�ѯ�����Ӳ�ѯ)
    gstrSQL = _
        " Select a.����id, Nvl(b.����, a.����) ����, Nvl(b.�Ա�, a.�Ա�) �Ա�, a.סԺ����, b.סԺ��, Nvl(b.����, a.����) ����, b.�ѱ�, b.��ҳid, b.��ǰ����id," & vbNewLine & _
        "       To_Char(b.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, To_Char(b.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, d.���� As ��Ժ����," & vbNewLine & _
        "       b.��Ժ����id, b.��Ժ����, b.��ǰ���� As ��Ժ����, c.���� As ����ȼ�, b.��Ժ��ʽ, b.��������, b.����," & vbNewLine & _
        "       (Select Decode(Nvl(Sum(���), 0), 0, '��', Null)" & vbNewLine & _
        "         From ����δ�����" & vbNewLine & _
        "         Where ��Դ;�� = 2 And ����id = b.����id And ��ҳid = b.��ҳid" & vbNewLine & _
        "          ) ����, a.���￨��, a.���֤��, a.Ic����, Nvl(b.��������, Decode(b.����, Null, '��ͨ����', 'ҽ������')) ��������," & vbNewLine & _
        "       a.��ҳid �������" & vbNewLine & _
        " From ������Ϣ a, ������ҳ b, �շ���ĿĿ¼ c, ���ű� d" & vbNewLine & _
        " Where b.����id = a.����id And b.��Ժ����id = d.Id And b.��Ժ���� Is Not Null And Nvl(b.��ҳid, 0) <> 0 And b.����ȼ�id = c.Id(+)" & gstrSQL


    gstrSQL = "Select /*+ rule*/ ����id,����,�Ա�,סԺ����,סԺ��, ����, �ѱ�, ��ҳid,��ǰ����ID,��Ժʱ��," & _
                "��Ժʱ�� , ��Ժ����, ��Ժ����ID, ��Ժ����, ��Ժ����, ����ȼ�, ��Ժ��ʽ, ��������, ����, ����" & _
                ",���￨��,���֤��,IC����,��������,�������  From (" & gstrSQL & ") Where 1=1" & _
                IIf(int���� = 0, "", IIf(int���� = 1, " And ���� is NULL", " And ���� is Not NULL")) & _
             " Order by ��Ժʱ�� Desc,סԺ�� Desc"
    
    On Error GoTo errH
    Set mrsOut = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUnitID, lngOutTime)
    Set mrsCOut = mrsOut.Clone
        
    mintOut = 0
    With mrsOut
        If mrsOut.RecordCount <> 0 Then
            For i = 1 To .RecordCount
                If IsNull(!�Ա�) Then
                    strSex = "O"
                Else
                    If InStr(!�Ա�, "��") > 0 Then
                        strSex = "M"
                    ElseIf InStr(!�Ա�, "Ů") > 0 Then
                        strSex = "F"
                    Else
                        strSex = "O"
                    End If
                End If
                
                '���۲���ͼ��
                If IIf(IsNull(!��������), 0, !��������) <> 0 Then strSex = "K" & strSex
                
                '�Բ���ID ��ҳΪ�ؼ���
                Set objItem = lvwOut_s.ListItems.Add(, "_" & !����ID & "_" & !��ҳID, !����, strSex, strSex)
                
                objItem.ForeColor = GetPatiColor(Nvl(mrsOut!��������, "��ͨ����"))
                For j = 2 To lvwOut_s.ColumnHeaders.Count
                    objItem.SubItems(j - 1) = IIf(IsNull(mrsOut.Fields(lvwOut_s.ColumnHeaders(j).Text).Value), "", mrsOut.Fields(lvwOut_s.ColumnHeaders(j).Text).Value)
                    objItem.ListSubItems(j - 1).ForeColor = objItem.ForeColor
                Next

                mintOut = mintOut + 1
                
                .MoveNext
            Next
        End If
    End With
    ReadOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwIn_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lvwIn_s.SelectedItem Is Nothing Then Exit Sub
    If Button = 1 And mblnIn Then
        Set lvwIn_s.DragIcon = lvwIn_s.SelectedItem.CreateDragImage
        lvwIn_s.Drag 1
    End If
End Sub

Private Sub lvwOut_s_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Static objIcon As IPictureDisp
    If Source Is lvwIn_s Or InStr(mstrPrivs, "���˳�Ժ") = 0 Then   '��Ʋ��˲����ϵ���Ժ�б�
        If State = 0 Then
            Set objIcon = Source.DragIcon
        ElseIf State = 2 Then
            Set Source.DragIcon = img32.ListImages("Limit").Picture
        ElseIf State = 1 Then
            Set Source.DragIcon = objIcon
        End If
    End If
End Sub

Private Sub lvwOut_s_DragDrop(Source As Control, X As Single, Y As Single)
    If (Source Is lvwBeds_s Or Source Is lvwFamily_s) And InStr(mstrPrivs, "���˳�Ժ") > 0 Then
        '���˳�Ժ����
        mnuEdit_Out_Click
    End If
End Sub

Private Sub lvwIn_s_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Static objIcon As IPictureDisp
    '�κζ�����������Ʋ����б�
    If Not Source Is lvwIn_s Then
        If State = 0 Then
            Set objIcon = Source.DragIcon
        ElseIf State = 2 Then
            Set Source.DragIcon = img32.ListImages("Limit").Picture
        ElseIf State = 1 Then
            Set Source.DragIcon = objIcon
        End If
    End If
End Sub

Private Sub lvwBeds_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnBeds Then
        '�մ�������,ת�Ʋ��˲��ܳ�Ժ�򻻴�
        If lvwBeds_s.SelectedItem.Tag <> "ռ��" Or lvwBeds_s.SelectedItem.Icon Like "*Change" Then Exit Sub
        If IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "��" Then
            Set lvwBeds_s.DragIcon = img32.ListImages("M").Picture
        ElseIf IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "Ů" Then
            Set lvwBeds_s.DragIcon = img32.ListImages("F").Picture
        Else
            Set lvwBeds_s.DragIcon = img32.ListImages("O").Picture
        End If
        lvwBeds_s.Drag 1
    End If
End Sub

Private Sub lvwBeds_s_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Dim objOver As ListItem
    
    If Source Is lvwIn_s Then
        Set objOver = lvwBeds_s.HitTest(X, Y)
        If Not objOver Is Nothing Then
            mrsCBeds.Filter = "����='" & Mid(objOver.Key, 2) & "'"
            
            If objOver.Tag = "�մ�" And (mrsCBeds!�Ա���� = "���޴�" Or _
            (mrsCBeds!�Ա���� = "�д�" And IIf(IsNull(mrsIn!�Ա�), "", mrsIn!�Ա�) = "��") _
            Or (mrsCBeds!�Ա���� = "Ů��" And IIf(IsNull(mrsIn!�Ա�), "", mrsIn!�Ա�) = "Ů")) Then
                Set lvwBeds_s.DropHighlight = objOver
                lvwBeds_s.DropHighlight.EnsureVisible
            ElseIf mrsCBeds!����ID = mrsIn!����ID _
                And mrsCBeds!���� = 1 And objOver.Tag <> "�մ�" Then '��ǰ��λ�����ô��������ǲ���ԭס��λ��ԭס��λ������ǰ������������Ϊ����Ʋ���ֻ������Ŀ����Һ͹�������
                Set lvwBeds_s.DropHighlight = objOver
                lvwBeds_s.DropHighlight.EnsureVisible
            End If
        Else
            Set lvwBeds_s.DropHighlight = Nothing
        End If
    ElseIf Source Is lvwBeds_s Then
        Set objOver = lvwBeds_s.HitTest(X, Y)
        If Not objOver Is Nothing And InStr(mstrPrivs, "����") <> 0 Then
            mrsCBeds.Filter = "����='" & Mid(objOver.Key, 2) & "'"
            'objOver.Tag = "�մ�" And
            If mrsBeds!���� = 1 Then
                If (mrsCBeds!����ID = mrsBeds!����ID Or IsNull(mrsCBeds!����ID) Or mrsCBeds!���� = 1) _
                    And Nvl(mrsBeds!����״̬, 0) <> 3 And Nvl(mrsCBeds!����״̬, 0) <> 3 And Nvl(mrsCBeds!����״̬, 0) <> 2 Then

                    If mrsBeds!�Ա���� = "���޴�" Then
                        If Not (mrsCBeds!�Ա���� = "���޴�" _
                                Or (mrsCBeds!�Ա���� = "�д�" And IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "��") _
                                Or (mrsCBeds!�Ա���� = "Ů��" And IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "Ů")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    ElseIf mrsBeds!�Ա���� = "�д�" Then
                        If Not ((mrsCBeds!�Ա���� = "���޴�" And mrsCBeds!�Ա� = "��") _
                                Or (mrsCBeds!�Ա���� = "�д�" And IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "��")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    ElseIf mrsBeds!�Ա���� = "Ů��" Then
                        If Not ((mrsCBeds!�Ա���� = "���޴�" And mrsCBeds!�Ա� = "Ů") _
                                Or (mrsCBeds!�Ա���� = "Ů��" And IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "Ů")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    End If
                    
                    Set lvwBeds_s.DropHighlight = objOver
                    lvwBeds_s.DropHighlight.EnsureVisible
                    
                End If
            Else
                If (mrsCBeds!����ID = mrsBeds!����ID Or IsNull(mrsCBeds!����ID) Or (mrsCBeds!���� = 1 And mrsCBeds!����ID = mrsBeds!����ID)) _
                    And Nvl(mrsBeds!����״̬, 0) <> 3 And Nvl(mrsCBeds!����״̬, 0) <> 3 And Nvl(mrsCBeds!����״̬, 0) <> 2 Then
                    
                    If mrsBeds!�Ա���� = "���޴�" Then
                        If Not (mrsCBeds!�Ա���� = "���޴�" _
                                Or (mrsCBeds!�Ա���� = "�д�" And IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "��") _
                                Or (mrsCBeds!�Ա���� = "Ů��" And IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "Ů")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    ElseIf mrsBeds!�Ա���� = "�д�" Then
                        If Not (mrsCBeds!�Ա���� = "���޴�" _
                                Or (mrsCBeds!�Ա���� = "�д�" And IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "��")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    ElseIf mrsBeds!�Ա���� = "Ů��" Then
                        If Not (mrsCBeds!�Ա���� = "���޴�" _
                                Or (mrsCBeds!�Ա���� = "Ů��" And IIf(IsNull(mrsBeds!�Ա�), "", mrsBeds!�Ա�) = "Ů")) Then
                           Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
                        End If
                    End If
                    
                    Set lvwBeds_s.DropHighlight = objOver
                    lvwBeds_s.DropHighlight.EnsureVisible
                
                End If
            End If
        Else
            Set lvwBeds_s.DropHighlight = Nothing
        End If
    ElseIf Source Is lvwFamily_s Then
        Set objOver = lvwBeds_s.HitTest(X, Y)
        If Not objOver Is Nothing Then
            mrsCBeds.Filter = "����='" & Mid(objOver.Key, 2) & "'"
        
            If objOver.Tag = "�մ�" And (mrsCBeds!����ID = mrsFamily!��ǰ����id Or IsNull(mrsCBeds!����ID)) _
                And Nvl(mrsFamily!����״̬, 0) <> 3 And (mrsCBeds!�Ա���� = "���޴�" _
                Or (mrsCBeds!�Ա���� = "�д�" And IIf(IsNull(mrsFamily!�Ա�), "", mrsFamily!�Ա�) = "��") _
                Or (mrsCBeds!�Ա���� = "Ů��" And IIf(IsNull(mrsFamily!�Ա�), "", mrsFamily!�Ա�) = "Ů")) Then
                Set lvwBeds_s.DropHighlight = objOver
                lvwBeds_s.DropHighlight.EnsureVisible
            End If
        Else
            Set lvwBeds_s.DropHighlight = Nothing
        End If
    End If
    If State = 1 Then Set lvwBeds_s.DropHighlight = Nothing
End Sub

Private Sub lvwBeds_s_DragDrop(Source As Control, X As Single, Y As Single)
    Dim strĿ�괲�� As String
    If Source Is lvwIn_s And Not lvwBeds_s.DropHighlight Is Nothing Then
        Set lvwBeds_s.SelectedItem = lvwBeds_s.DropHighlight
        Set lvwBeds_s.DropHighlight = Nothing
        '��������ס����(˫����ѡ����)
        
        Call lvwBeds_s_ItemClick(lvwBeds_s.SelectedItem)
        
        If mrsIn!��Ʊ�־ = 2 Then
            Call mnuEdit_InUnit_Click
        Else
            Call mnuEdit_In_Click
        End If
        
    ElseIf (Source Is lvwFamily_s Or Source Is lvwBeds_s) And Not lvwBeds_s.DropHighlight Is Nothing Then
        '���˻�������
        If InStr(mstrPrivs, "����") = 0 Then Set lvwBeds_s.DropHighlight = Nothing: Exit Sub
        strĿ�괲�� = Trim(Mid(lvwBeds_s.DropHighlight.Key, 2))
        
        mrsCBeds.Filter = "����='" & strĿ�괲�� & "'"
        
        If Nvl(mrsCBeds!����ID, 0) = 0 Then
            Set lvwBeds_s.DropHighlight = Nothing
            Call ChangeBeds(0, strĿ�괲��)
        Else
            Set lvwBeds_s.DropHighlight = Nothing
            Call SwapBeds(strĿ�괲��)
        End If
    End If
End Sub

Private Sub lvwFamily_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mblnFamily Then
        If lvwFamily_s.SelectedItem Is Nothing Then Exit Sub
        If lvwFamily_s.SelectedItem.Icon Like "*Change" Then Exit Sub
        
        If IIf(IsNull(mrsFamily!�Ա�), "", mrsFamily!�Ա�) = "��" Then
            Set lvwFamily_s.DragIcon = img32.ListImages("M").Picture
        ElseIf IIf(IsNull(mrsFamily!�Ա�), "", mrsFamily!�Ա�) = "Ů" Then
            Set lvwFamily_s.DragIcon = img32.ListImages("F").Picture
        Else
            Set lvwFamily_s.DragIcon = img32.ListImages("O").Picture
        End If
        lvwFamily_s.Drag 1
    End If
End Sub

Private Sub lvwFamily_s_DragDrop(Source As Control, X As Single, Y As Single)
    If Source Is lvwIn_s Then
        '������ƴ���(��ͥ����)
        Dim byt��Ʒ�ʽ As Byte, lng����ID As Long, lng��ҳID As Long
        
        byt��Ʒ�ʽ = Val(lvwIn_s.SelectedItem.Tag)
        lng����ID = mrsIn!����ID
        lng��ҳID = mrsIn!��ҳID
        Call ExecPatiChange(EFun.E���, Me, mstrPrivs, mlngUnit, lng����ID, lng��ҳID, "��ͥ����", 0, byt��Ʒ�ʽ)
        
        If gblnOK Then Call LoadList(True, True, True, False)
        
    ElseIf Source Is lvwBeds_s And InStr(1, mstrPrivs, "��ͥ����") > 0 And InStr(mstrPrivs, "����") Then
        If Nvl(mrsBeds!����״̬, 0) <> 3 Then
            '���˻�������
            
            Call ChangeBeds(0, "��ͥ����")
        Else
            MsgBox "Ԥ��Ժ���˲��ܽ��л�������!"
        End If
    End If
End Sub

Private Sub mnuFile_Excel_Click()
    If mobjLVW.ListItems.Count > 100 Then
        If MsgBox("�����Excel�����ݹ���,�⽫�ķ����ʱ��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrintLvw
    Dim bytR As Byte
    
    On Error GoTo errH
    
    '��ͷ
    Select Case mobjLVW.Name
        Case "lvwBeds_s"
            objOut.Title.Text = "��λӳ���"
        Case "lvwFamily_s"
            objOut.Title.Text = "��ͥ������"
        Case "lvwIn_s"
            objOut.Title.Text = "��Ʋ��˱�"
        Case "lvwOut_s"
            objOut.Title.Text = "��Ժ���˱�"
    End Select
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    objOut.UnderAppItems.Add "����:" & zlCommFun.GetNeedName(cboUnit.Text)
    objOut.BelowAppItems.Add "��ӡ�ˣ�" & UserInfo.����
    objOut.BelowAppItems.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    
    '����
    Set objOut.Body.objData = mobjLVW
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrViewLvw objOut, bytR
    Else
        zlPrintOrViewLvw objOut, bytStyle
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub LoadList(Optional mblnBeds As Boolean = True, Optional mblnFamily As Boolean = True, _
    Optional mblnIn As Boolean = True, Optional mblnOut As Boolean = True, Optional mblnReadNure As Boolean)
'���ܣ�ˢ�½����б�����
'������ȱʡˢ�������б�,���Էֱ�ָ��
    Dim strBeds As String, strFamily As String
    Dim strIn As String, strOut As String, lngUnit As Long
    Dim objFind As ListItem
    
    '��¼ԭλ�ã���������б�
    If mblnBeds Then
        If Not lvwBeds_s.SelectedItem Is Nothing Then strBeds = lvwBeds_s.SelectedItem.Key
        lvwBeds_s.ListItems.Clear
    End If
    If mblnFamily Then
        If Not lvwFamily_s.SelectedItem Is Nothing Then strFamily = lvwFamily_s.SelectedItem.Key
        lvwFamily_s.ListItems.Clear
    End If
    If mblnIn Then
        If Not lvwIn_s.SelectedItem Is Nothing Then strIn = lvwIn_s.SelectedItem.Key
        lvwIn_s.ListItems.Clear
    End If
    If mblnOut Then
        If Not lvwOut_s.SelectedItem Is Nothing Then strOut = lvwOut_s.SelectedItem.Key
        lvwOut_s.ListItems.Clear
    End If
    
    If mblnReadNure = True Then
        '��ס��������ס����������ȼ�����������ȼ�����Ժʱˢ�»���ȼ�
        If Not ReadNure(mlngUnit) Then Exit Sub
    End If
    'ˢ�´�λӳ���
    If mblnBeds Then Call ReadBedsMap(mlngUnit)
    
    'ˢ�¼�ͥ������
    If mblnFamily Then Call ReadFamily(mlngUnit)
    
    'ˢ����Ʋ��˱�
    If mblnIn Then Call ReadIn(mlngUnit)
        
    'ˢ�³�Ժ���˱�
    If mblnOut Then Call ReadOut(mlngUnit)
    
    '�Զ���λ����ǰλ��
    On Error Resume Next
    
    If strBeds <> "" And lvwBeds_s.ListItems.Count > 0 And mblnBeds Then
        Set objFind = lvwBeds_s.ListItems(strBeds)
        If Err.Number = 0 Then
            objFind.Selected = True
            objFind.EnsureVisible
        Else
            Err.Clear
        End If
    End If
    If strFamily <> "" And lvwFamily_s.ListItems.Count > 0 And mblnFamily Then
        Set objFind = lvwFamily_s.ListItems(strFamily)
        If Err.Number = 0 Then
            objFind.Selected = True
            objFind.EnsureVisible
        Else
            Err.Clear
        End If
    End If
    If strIn <> "" And lvwIn_s.ListItems.Count > 0 And mblnIn Then
        Set objFind = lvwIn_s.ListItems(strIn)
        If Err.Number = 0 Then
            objFind.Selected = True
            objFind.EnsureVisible
        Else
            Err.Clear
        End If
    End If
    If strOut <> "" And lvwOut_s.ListItems.Count > 0 And mblnOut Then
        Set objFind = lvwOut_s.ListItems(strOut)
        If Err.Number = 0 Then
            objFind.Selected = True
            objFind.EnsureVisible
        Else
            Err.Clear
        End If
    End If
    
    '83992:ÿ�θ���ѡ���λ��¼���е�������
    If Not lvwBeds_s.SelectedItem Is Nothing Then mrsBeds.Filter = "����='" & Mid(lvwBeds_s.SelectedItem.Key, 2) & "'"
    If Not lvwFamily_s.SelectedItem Is Nothing Then mrsFamily.Filter = "����ID=" & Mid(lvwFamily_s.SelectedItem.Key, 2)
    If Not lvwIn_s.SelectedItem Is Nothing Then mrsIn.Filter = "����ID=" & Mid(lvwIn_s.SelectedItem.Key, 2)
    If Not lvwOut_s.SelectedItem Is Nothing Then mrsOut.Filter = "����ID=" & Split(lvwOut_s.SelectedItem.Key, "_")(1) & " and ��ҳid=" & Split(lvwOut_s.SelectedItem.Key, "_")(2)
    
    If mobjLVW Is lvwBeds_s Then
        Call lvwBeds_s_ItemClick(lvwBeds_s.SelectedItem)
    ElseIf mobjLVW Is lvwFamily_s Then
        Call lvwFamily_s_ItemClick(lvwFamily_s.SelectedItem)
    ElseIf mobjLVW Is lvwIn_s Then
        Call lvwIn_s_ItemClick(lvwIn_s.SelectedItem)
    ElseIf mobjLVW Is lvwOut_s Then
        Call lvwOut_s_ItemClick(lvwOut_s.SelectedItem)
    Else
        If Not lvwBeds_s.SelectedItem Is Nothing Then Call lvwBeds_s_ItemClick(lvwBeds_s.SelectedItem)
    End If
    If Me.Visible Then mobjLVW.SetFocus
End Sub

Private Sub SetMenu()
'���ܣ����ݵ�ǰѡ���б��λ��������Ӧ�˵����ܵ�״̬��
    Dim lng����ID As Long, lng��ҳID As Long, lng���� As Long, blnDo As Boolean
    Dim rsTmp As ADODB.Recordset
    
    'סԺ���첡��תΪסԺ����
    blnDo = True
    If mobjLVW Is lvwIn_s Or mobjLVW Is lvwOut_s Then
        blnDo = False
    Else
        If mobjLVW Is lvwBeds_s Then
            If lvwBeds_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            End If
        Else
            If lvwFamily_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwFamily_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            End If
        End If
        If blnDo Then
            If mobjLVW Is lvwBeds_s Then
                blnDo = IIf(IsNull(mrsBeds!��������), 0, mrsBeds!��������) = 2
            ElseIf mobjLVW Is lvwFamily_s Then
                blnDo = IIf(IsNull(mrsFamily!��������), 0, mrsFamily!��������) = 2
            End If
        End If
    End If
    mnuEditToInPati.Enabled = blnDo
    
    '�������(���ʱѡ�񴲺�,�������ݽ����������)
    blnDo = True
    If Not mobjLVW Is lvwIn_s Then
        blnDo = False
        mnuEdit_In.Enabled = blnDo
    Else
        If lvwIn_s.SelectedItem Is Nothing Then
            blnDo = False
            mnuEdit_In.Enabled = blnDo
        Else
            mnuEdit_In.Enabled = blnDo
        End If
    End If
    tbr.Buttons("In").Enabled = blnDo
    
    '���˱�ע�༭
    blnDo = True
    If mobjLVW Is lvwBeds_s Then
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwFamily_s Then
        If lvwFamily_s.SelectedItem Is Nothing Then blnDo = False
    ElseIf mobjLVW Is lvwOut_s Then
        If lvwOut_s.SelectedItem Is Nothing Then blnDo = False
    ElseIf mobjLVW Is lvwIn_s Then
        If lvwIn_s.SelectedItem Is Nothing Then blnDo = False
    End If
    mnuEdit_Memo.Enabled = blnDo
    
    '����ת�ơ����˻�����Ԥ��Ժ������ȼ���������Ϣ���������Ǽ�(ת�ƺ�Ԥ��Ժ״̬����)
    blnDo = True
    If mobjLVW Is lvwIn_s Or mobjLVW Is lvwOut_s Then
        blnDo = False
    Else
        If mobjLVW Is lvwBeds_s Then
            If lvwBeds_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Icon Like "*Out" Then
                blnDo = False
            End If
        Else
            If lvwFamily_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwFamily_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            ElseIf lvwFamily_s.SelectedItem.Icon Like "*Out" Then
                blnDo = False
            End If
        End If
    End If
    mnuEdit_Change.Enabled = blnDo
    mnuEdit_ChangeUnit.Enabled = blnDo
    tbr.Buttons("Change").Enabled = blnDo
    
    mnuEdit_ChangeGroup.Enabled = blnDo

    mnuEdit_Move.Enabled = blnDo
    mnuEdit_Swap.Enabled = blnDo
    tbr.Buttons("Move").Enabled = blnDo
    mnuEdit_AddBeds.Enabled = blnDo
    
    mnuEdit_PreOut.Enabled = blnDo
    mnuEdit_Nurse.Enabled = blnDo
    
    mnuEdit_Adjust.Enabled = blnDo
    tbr.Buttons("Adjust").Enabled = blnDo
    
    '�������Ǽ�(���Ʋ��˲�����)
    If blnDo Then
        If mobjLVW Is lvwBeds_s Then
            blnDo = is����(mrsBeds!��ǰ����id, Nothing)
            blnDo = blnDo And mrsBeds!�Ա� = "Ů"
        ElseIf mobjLVW Is lvwFamily_s Then
            blnDo = is����(mrsFamily!��ǰ����id, Nothing)
            blnDo = blnDo And mrsFamily!�Ա� = "Ů"
        End If
    End If
    mnuEdit_BabyReg.Enabled = blnDo
    
    '�������,�޸ĳ�Ժʱ��
    blnDo = True
    If Not mobjLVW Is Nothing Then
        If mobjLVW.SelectedItem Is Nothing Then
            blnDo = False
        Else
            If mobjLVW Is lvwBeds_s Then
                If lvwBeds_s.SelectedItem.Tag <> "ռ��" Then blnDo = False
                Set rsTmp = mrsBeds
            ElseIf mobjLVW Is lvwFamily_s Then
                Set rsTmp = mrsFamily
            ElseIf mobjLVW Is lvwOut_s Then
                If Nvl(mrsOut!��ҳID, 0) <> Nvl(mrsOut!�������, 0) Then blnDo = False
                Set rsTmp = mrsOut
            ElseIf mobjLVW Is lvwIn_s Then
                Set rsTmp = mrsIn
            End If
            If Nvl(rsTmp!����, 0) <> 0 Then blnDo = False
        End If
    Else
        blnDo = False
    End If
    mnuEdit_Recalc.Enabled = blnDo
    '����27392 by lesfeng 2010-01-14
    If mobjLVW Is lvwOut_s Then
        mnuEdit_ModifOut.Enabled = blnDo
    Else
        mnuEdit_ModifOut.Enabled = False
    End If
    
    '���˳�Ժ(ת��״̬����)
    blnDo = True
    If mobjLVW Is lvwIn_s Or mobjLVW Is lvwOut_s Then
        blnDo = False
    Else
        If mobjLVW Is lvwBeds_s Then
            If lvwBeds_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            End If
        Else
            If lvwFamily_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwFamily_s.SelectedItem.Icon Like "*Change" Then
                blnDo = False
            End If
        End If
    End If
    mnuEdit_Out.Enabled = blnDo
    tbr.Buttons("Out").Enabled = blnDo
    
    '��λ�ȼ�(ת��״̬��Ԥ��Ժ����)
    blnDo = True
    If Not mobjLVW Is lvwBeds_s Then
        blnDo = False
    Else
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Icon Like "*Change" Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Icon Like "*Out" Then
            blnDo = False
        End If
    End If
    mnuEdit_Level.Enabled = blnDo
    
    'ת�Ƽ�¼����λ��¼�������¼
    blnDo = True
    If mobjLVW Is lvwIn_s Then
        blnDo = False
    Else
        If mobjLVW Is lvwFamily_s Then
            If lvwFamily_s.SelectedItem Is Nothing Then
                blnDo = False
            End If
        ElseIf mobjLVW Is lvwOut_s Then
            If lvwOut_s.SelectedItem Is Nothing Then
                blnDo = False
            End If
        Else
            If lvwBeds_s.SelectedItem Is Nothing Then
                blnDo = False
            ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
                blnDo = False
            End If
        End If
    End If
    mnuQuery_Log.Enabled = blnDo

    '��ӡ������������Ϣ��������Ϣ(�������Ͳ��˾���)
    blnDo = True
    If mobjLVW Is lvwBeds_s Then
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwFamily_s Then
        If lvwFamily_s.SelectedItem Is Nothing Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwOut_s Then
        If lvwOut_s.SelectedItem Is Nothing Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwIn_s Then
        If lvwIn_s.SelectedItem Is Nothing Then
            blnDo = False
        End If
    End If
    mnuFilePrintMed.Enabled = blnDo
    mnuQueryInfo.Enabled = blnDo
    
    '49854:������,2013-10-31,���������ӡ
    '��ӡ�������(��Ժ���˲���)
    mnuFile_PrintWristlet.Visible = (InStr(mstrPrivs, ";�����ӡ;") > 0)
    mnuFile_PrintWristlet.Enabled = mnuFile_PrintWristlet.Visible And blnDo And (Not mobjLVW Is lvwOut_s)
    
    '����ѡ��
    blnDo = True
    If mobjLVW Is lvwIn_s Then
        blnDo = False
    ElseIf mobjLVW Is lvwOut_s Then
        If lvwOut_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf Nvl(mrsOut!����, 0) = 0 Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwBeds_s Then
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
            blnDo = False
        ElseIf Nvl(mrsBeds!����, 0) = 0 Then
            blnDo = False
        End If
    ElseIf mobjLVW Is lvwFamily_s Then
        If lvwFamily_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf Nvl(mrsFamily!����, 0) = 0 Then
            blnDo = False
        End If
    End If
    mnuEdit_Disease.Enabled = blnDo
    
    '����״̬����
    blnDo = True
    If mobjLVW Is lvwBeds_s Then
        If lvwBeds_s.SelectedItem Is Nothing Then
            blnDo = False
        ElseIf lvwBeds_s.SelectedItem.Tag <> "ռ��" Then
            blnDo = False
        Else
            lng����ID = mrsBeds!����ID
            lng��ҳID = mrsBeds!��ҳID
        End If
    ElseIf mobjLVW Is lvwFamily_s Then
        If lvwFamily_s.SelectedItem Is Nothing Then
            blnDo = False
        Else
            lng����ID = mrsFamily!����ID
            lng��ҳID = mrsFamily!��ҳID
        End If
    ElseIf mobjLVW Is lvwOut_s Then
        If lvwOut_s.SelectedItem Is Nothing Then
            blnDo = False
        Else
            lng����ID = mrsOut!����ID
            lng��ҳID = mrsOut!��ҳID
            If Nvl(mrsOut!��ҳID, 0) <> Nvl(mrsOut!�������, 0) Then blnDo = False
        End If
    ElseIf mobjLVW Is lvwIn_s Then
        blnDo = False
    End If
    tbr.Buttons("Undo").ButtonMenus.Clear
    mnuEdit_Undo.Caption = "����(&U)"
    tbr.Buttons("Undo").Enabled = False
    mnuEdit_Undo.Enabled = False
    If lng����ID > 0 And lng��ҳID > 0 And blnDo Then Call SetUndoLog(lng����ID, lng��ҳID)
    
    '����
    If lvwBeds_s.ListItems.Count > 0 Or lvwIn_s.ListItems.Count > 0 Or lvwFamily_s.ListItems.Count > 0 Or lvwOut_s.ListItems.Count > 0 Then
        mnuViewFind.Enabled = True
        mnuViewFindNext.Enabled = (mstrSeekKey = "����" Or mstrSeekKey = "סԺ��")
    Else
        mnuViewFind.Enabled = False
        mnuViewFindNext.Enabled = False
    End If
    
End Sub

Private Sub SetUndoLog(lng����ID As Long, lng��ҳID As Long)
'���ܣ����ݲ���������ʾ���˿ɳ�������
'˵����1.���øú���֮ǰ���ù��ܵĳ�ʼ״̬
'      2.���ܳ�����Ժ(��Ժͬʱ��Ƶĳ�������Ժ״̬)
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, blnExist As Boolean
    Dim objMenu As ButtonMenu
    
    Set rsTmp = GetPatiLog(lng����ID, lng��ҳID)
    If rsTmp Is Nothing Then Exit Sub
    
    mnuEdit_Undo.Enabled = True
    tbr.Buttons("Undo").Enabled = True
        
    '����ǳ�Ժ,����һ��������Ժ
    If Not IsNull(rsTmp!��ֹʱ��) And rsTmp!��ֹԭ�� = 1 Then
        Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "��Ժ")
        If tbr.Buttons("Undo").ButtonMenus.Count > 1 Then objMenu.Enabled = False
        mnuEdit_Undo.Caption = "������Ժ(&U)"
        blnExist = True
        
        If InStr(";" & mstrPrivs & ";", ";������Ժ;") = 0 Then
            objMenu.Enabled = False
            mnuEdit_Undo.Enabled = False
        End If
    End If
    '����28386 by lesfeng 2010-03-06 ������ʼԭ��Ϊ2\3����Ʒֱ�Ϊ��Ժ���\ת�����
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!��ʼʱ��) And rsTmp!��ʼԭ�� = 3 Then
            Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "ת��")
        ElseIf IsNull(rsTmp!��ʼʱ��) And rsTmp!��ʼԭ�� = 15 Then
            Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "ת����")
        Else
            Select Case rsTmp!��ʼԭ��
                Case 1 '��Ժ
                    '��lvwIN�еĲ��˵�ǰ�ɳ�����Ϊ��Ժ�䶯��һ������Ժͬʱ���,����Ϊ�������
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "��Ժ��ס")
                    '���ǵ�ǰ�ɳ�������Ժ�䶯,��Ϊ��������Ժ�Ǽ�
                    If tbr.Buttons("Undo").ButtonMenus.Count > 1 Then objMenu.Visible = False
                Case 2 '��Ժ���
'                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "���")
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "��ס")
                Case 3 'ת�����
'                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "���")
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "ת����ס")
                Case 4
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "����")
                Case 5
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "��λ�ȼ��䶯")
                Case 6
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "����ȼ��䶯")
                Case 7
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "����ҽʦ�ı�")
                Case 8
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "���λ�ʿ�ı�")
                Case 9
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "תΪסԺ����")
                Case 10
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "Ԥ��Ժ")
                Case 11
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "����ҽʦ�䶯")
                Case 12
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "����ҽʦ�䶯")
                Case 13
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "�����䶯")
                Case 14
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "תҽ��С��")
                Case 15
                    Set objMenu = tbr.Buttons("Undo").ButtonMenus.Add(, , "ת������ס")
            End Select
        End If
        If tbr.Buttons("Undo").ButtonMenus.Count > 1 Then objMenu.Enabled = False
        If Not blnExist And i = 1 Then mnuEdit_Undo.Caption = "����" & objMenu.Text & "(&U)"
        
        If InStr(mstrPrivs, "�������") = 0 And (objMenu.Text = "��ס" Or objMenu.Text = "��Ժ��ס" Or objMenu.Text = "ת����ס" Or objMenu.Text = "ת������ס") Then
            objMenu.Enabled = False
        End If
        
        If InStr(mstrPrivs, "����") = 0 And (objMenu.Text = "����") Then
            objMenu.Enabled = False
        End If
        
        If InStr(mstrPrivs, "סԺ����תסԺ") = 0 And objMenu.Text = "תΪסԺ����" Then
            objMenu.Enabled = False
        End If
        
        If InStr(mstrPrivs, "����Ԥ��Ժ") = 0 And objMenu.Text = "Ԥ��Ժ" Then
            objMenu.Enabled = False
        End If
        
        rsTmp.MoveNext
    Next
    
    If tbr.Buttons("Undo").ButtonMenus(1).Enabled = False Then '���,��Not��ʽ��Ч
        mnuEdit_Undo.Enabled = False
    End If
End Sub
Private Function InitPatiType() As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    mstrPatiTypeColor = ""
    gstrSQL = "select ����,��ɫ from ��������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������")
    Do Until rsTemp.EOF
        mstrPatiTypeColor = mstrPatiTypeColor & rsTemp!���� & "," & Nvl(rsTemp!��ɫ, 0) & "|"
        rsTemp.MoveNext
    Loop
    If Len(mstrPatiTypeColor) > 0 Then
        mstrPatiTypeColor = Mid(mstrPatiTypeColor, 1, Len(mstrPatiTypeColor) - 1)
    Else
        mstrPatiTypeColor = "��ͨ����,0|ҽ������,255"
    End If
    InitPatiType = True
    Exit Function
errH:
    mstrPatiTypeColor = "��ͨ����,0|ҽ������,255"
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetPatiColor(ByVal strPatiType) As Long
Dim arrType As Variant, i As Integer
    arrType = Split(mstrPatiTypeColor, "|")
    For i = LBound(arrType) To UBound(arrType)
        If Split(arrType(i), ",")(0) = strPatiType Then
            GetPatiColor = Split(arrType(i), ",")(1)
            Exit Function
        End If
    Next
End Function
Private Function GetLenText(ByVal strText As String, ByVal lngLen As Long) As String
'����������������ո�ָ������
    Dim i As Long
    
    i = zlCommFun.ActualLen(strText)
    If i < lngLen Then
        i = lngLen - i
    Else
        i = i - lngLen
    End If
    GetLenText = strText & Space(i)
End Function

Private Sub MakeBedIcon()
    Dim i As Integer, k As Integer
    
    k = img32.ListImages.Count
    For i = 13 To 22
        img32.ListImages.Add , "�Ӵ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_�Ӵ�", i)
        img32.ListImages.Add , "�Ǳ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_�Ǳ�", i)
        img32.ListImages.Add , "����_" & img32.ListImages(i).Key, img32.Overlay("MASK_����", i)
        img32.ListImages.Add , "����_�Ӵ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_����_�Ӵ�", i)
        img32.ListImages.Add , "����_�Ǳ�_" & img32.ListImages(i).Key, img32.Overlay("MASK_����_�Ǳ�", i)
    Next
    
    img32.ListImages.Add , "M_ChangeUnit", img32.Overlay("M", "U")
    img32.ListImages.Add , "KM_ChangeUnit", img32.Overlay("KM", "KU")
    img32.ListImages.Add , "F_ChangeUnit", img32.Overlay("F", "U")
    img32.ListImages.Add , "FM_ChangeUnit", img32.Overlay("KF", "KU")

    k = img16.ListImages.Count
    For i = 13 To 22
        img16.ListImages.Add , "�Ӵ�_" & img16.ListImages(i).Key, img16.Overlay("MASK_�Ӵ�", i)
        img16.ListImages.Add , "�Ǳ�_" & img16.ListImages(i).Key, img16.Overlay("MASK_�Ǳ�", i)
        img16.ListImages.Add , "����_" & img16.ListImages(i).Key, img16.Overlay("MASK_����", i)
        img16.ListImages.Add , "����_�Ӵ�_" & img16.ListImages(i).Key, img16.Overlay("MASK_����_�Ӵ�", i)
        img16.ListImages.Add , "����_�Ǳ�_" & img16.ListImages(i).Key, img16.Overlay("MASK_����_�Ǳ�", i)
    Next
    
    img16.ListImages.Add , "M_ChangeUnit", img16.Overlay("M", "U")
    img16.ListImages.Add , "KM_ChangeUnit", img16.Overlay("KM", "KU")
    img16.ListImages.Add , "F_ChangeUnit", img16.Overlay("F", "U")
    img16.ListImages.Add , "FM_ChangeUnit", img16.Overlay("KF", "KU")
End Sub

Private Sub tbrFilter_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim blnCheck As Boolean, i As Long
    
    If Button.Key = "curDay" Then
        Button.Image = IIf(Button.Image = "UnCheck_", "Check_", "UnCheck_")
        Call LoadList(True, True, False, False)
    ElseIf Button.Key Like "Nurse*" Then
        '��׼ȫ�����
        blnCheck = False
        For i = 1 To tbrFilter.Buttons.Count
            If tbrFilter.Buttons(i).Key Like "Nurse*" _
                And tbrFilter.Buttons(i).Key <> Button.Key Then
                If tbrFilter.Buttons(i).Image = "Check" Then
                    blnCheck = True: Exit For
                End If
            End If
        Next
        If blnCheck Then
            Button.Image = IIf(Button.Image = "UnCheck", "Check", "UnCheck")
            Call LoadList(True, True, False, False)
        Else
            Button.Image = "Check"
            Exit Sub
        End If
    End If
End Sub

Private Sub tbrFilter_Change()
    Caption = Timer
End Sub

Private Sub timSize_Timer()
    Call Form_Resize
    timSize.Enabled = False
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub
'����28811 by lesfeng 2010-03-30
Private Function InputGetDept(ByRef cboToDept As ComboBox, ByRef blnCancel As Boolean) As ADODB.Recordset
    'ѡ��������
'    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim lngHeigth As Long
    Dim strInput As String
    Dim strInputN As String
    Dim strno As String
    Dim blnLimitUnit As Boolean
    Dim strUnitIDs As String
    
    On Error GoTo errH
    
    cboToDept.Text = Replace(UCase(cboToDept.Text), "'", "")
    strInput = UCase(cboToDept.Text)
    strInputN = gstrLike & strInput & "%"
    strno = strInput & "%"
    
    If zlCommFun.IsCharChinese(strInput) Or InStr(1, strInput, "-", 0) <> 0 Then
        strSQL = strSQL & " And (A.���� Like [3] or A.����||'-'||A.���� Like [3])" '���뺺��ʱֻƥ������
    Else
        strSQL = strSQL & " And (A.���� Like [4] Or A.���� Like [3] Or A.���� Like [3])"
    End If
    
    blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
    If blnLimitUnit Then
        strUnitIDs = "," & GetUserUnits(False) & ","
    Else
        strUnitIDs = "," & GetUserUnits(True) & ","
    End If
    'Ŀǰ��������۲���
    strSQL = _
        " Select A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where B.����ID = A.ID" & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.������� IN(1,2,3) And B.��������='����'" & _
        IIf(blnLimitUnit, " And instr([1],',' || A.ID || ',')>0 ", "") & _
        " And (A.վ��=[2] Or A.վ�� is Null) " & strSQL & _
        " Order by A.����"
        '
    vRect = zlControl.GetControlRect(cboToDept.hWnd)
    lngHeigth = cboToDept.Height

    Set InputGetDept = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ѡ��", False, cboToDept.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strUnitIDs, gstrNodeNo, strInputN, strno)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


