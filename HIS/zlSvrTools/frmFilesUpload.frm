VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFilesUpload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ļ��ϴ�"
   ClientHeight    =   9888
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   15912
   Icon            =   "frmFilesUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10296.87
   ScaleMode       =   0  'User
   ScaleWidth      =   15915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picHelp 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   0
      ScaleHeight     =   996
      ScaleWidth      =   15912
      TabIndex        =   18
      Top             =   0
      Width           =   15915
      Begin VB.PictureBox picState 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   0
         Left            =   1290
         ScaleHeight     =   780
         ScaleWidth      =   8628
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   8625
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��װ·��ȱʧ���ļ��İ�װ·��Ϊ�գ����޸�Ϊ��Ч�ļ�"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   5
            Left            =   15
            TabIndex        =   31
            Top             =   540
            Width           =   4500
         End
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ļ�ȱʧ����ȷ�ϸ��ļ�·�����ļ�����"
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   4
            Left            =   15
            TabIndex        =   30
            Top             =   0
            Width           =   3600
         End
         Begin VB.Label lblEXP 
            BackStyle       =   0  'Transparent
            Caption         =   "��׼����ȱʧ����׼�ļ��ڵ�ǰ�����嵥��ȱʧ�����������ļ��嵥���޸���ǰ�����ļ��嵥"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   15
            TabIndex        =   29
            Top             =   270
            Width           =   8505
         End
      End
      Begin VB.PictureBox picState 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   1
         Left            =   1290
         ScaleHeight     =   780
         ScaleWidth      =   8628
         TabIndex        =   32
         Top             =   135
         Visible         =   0   'False
         Width           =   8625
         Begin VB.Label lblEXP 
            BackStyle       =   0  'Transparent
            Caption         =   "�����ļ����׼�ļ��嵥���ļ���ƥ�䣬���ܴ��ڷ��գ�����ϸ��龯���ļ���֤������"
            ForeColor       =   &H00007FFF&
            Height          =   225
            Index           =   8
            Left            =   15
            TabIndex        =   35
            Top             =   270
            Width           =   8505
         End
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���棺�����ļ����׼�ļ���������ȷ���ļ�����"
            ForeColor       =   &H00007FFF&
            Height          =   180
            Index           =   7
            Left            =   15
            TabIndex        =   34
            Top             =   0
            Width           =   3960
         End
         Begin VB.Label lblEXP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ȷ�������ļ���ȷ�ԣ�������ܻ����һЩ����"
            ForeColor       =   &H00007FFF&
            Height          =   180
            Index           =   6
            Left            =   30
            TabIndex        =   33
            Top             =   540
            Width           =   3960
         End
      End
      Begin VB.Frame fraBounds 
         Height          =   1170
         Index           =   1
         Left            =   10020
         TabIndex        =   27
         Top             =   -135
         Width           =   30
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ҫ���̣�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   10230
         TabIndex        =   26
         Top             =   405
         Width           =   900
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ��ļ�"
         Height          =   180
         Index           =   2
         Left            =   14700
         TabIndex        =   25
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ռ��ļ�"
         Height          =   180
         Index           =   1
         Left            =   13005
         TabIndex        =   24
         Top             =   75
         Width           =   720
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ļ�"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   11250
         TabIndex        =   23
         Top             =   75
         Width           =   720
      End
      Begin VB.Image imgPro 
         Height          =   540
         Index           =   4
         Left            =   13872
         Picture         =   "frmFilesUpload.frx":6852
         Top             =   288
         Width           =   540
      End
      Begin VB.Image imgPro 
         Height          =   540
         Index           =   3
         Left            =   12156
         Picture         =   "frmFilesUpload.frx":807E
         Top             =   288
         Width           =   540
      End
      Begin VB.Image imgPro 
         Height          =   576
         Index           =   2
         Left            =   14688
         Picture         =   "frmFilesUpload.frx":98AA
         Top             =   276
         Width           =   576
      End
      Begin VB.Image imgPro 
         Height          =   576
         Index           =   1
         Left            =   12996
         Picture         =   "frmFilesUpload.frx":B3EC
         Top             =   276
         Width           =   576
      End
      Begin VB.Image imgPro 
         Height          =   576
         Index           =   0
         Left            =   11232
         Picture         =   "frmFilesUpload.frx":CF2E
         Top             =   252
         Width           =   576
      End
      Begin VB.Image imgCaption 
         Height          =   576
         Left            =   300
         Picture         =   "frmFilesUpload.frx":EA70
         Top             =   120
         Width           =   576
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "�ռ��ļ�����Ҫ����ǰ�������ϴ����ļ�ѹ������ʱ�ļ��У��������ж��Ѿ�ѹ�������ļ����ӿ��ռ�����"
         Height          =   225
         Index           =   1
         Left            =   1305
         TabIndex        =   21
         Top             =   405
         Width           =   8505
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ļ�����Ҫ����ļ��Ƿ��쳣(����ȱʧ����׼����ȱʧ��)������ļ��Ƿ���Ҫ�ϴ�"
         Height          =   180
         Index           =   0
         Left            =   1305
         TabIndex        =   20
         Top             =   135
         Width           =   7020
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ��ļ������������Ѿ����õ��ļ����ϴ���ǰ�����ļ�"
         Height          =   180
         Index           =   2
         Left            =   1300
         TabIndex        =   19
         Top             =   675
         Width           =   4500
      End
   End
   Begin VB.Frame fraBounds 
      Height          =   30
      Index           =   0
      Left            =   -1695
      TabIndex        =   22
      Top             =   1005
      Width           =   17730
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   13815
      Top             =   1530
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   1048
      ImageHeight     =   27
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":105B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":2519C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":39D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":4E970
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":6355A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":78144
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":8CD2E
            Key             =   "�ļ�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":8E880
            Key             =   "����ļ�"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":903D2
            Key             =   "�ռ��ļ�"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":91F24
            Key             =   "�ϴ��ļ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":93A76
            Key             =   "�ļ��쳣"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":955C8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilesUpload.frx":9711A
            Key             =   "�쳣"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInformation 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   324
      Left            =   24
      ScaleHeight     =   405.238
      ScaleMode       =   0  'User
      ScaleWidth      =   15847.55
      TabIndex        =   7
      Top             =   9360
      Width           =   15840
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ϴ��ļ���"
         Height          =   180
         Index           =   5
         Left            =   13470
         TabIndex        =   13
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��龯���ļ���"
         Height          =   180
         Index           =   4
         Left            =   10890
         TabIndex        =   12
         Top             =   75
         Width           =   1260
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ϴ����ļ���"
         Height          =   180
         Index           =   3
         Left            =   8565
         TabIndex        =   11
         Top             =   90
         Width           =   1260
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ϴ��ļ���"
         Height          =   180
         Index           =   2
         Left            =   5685
         TabIndex        =   10
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "״̬�쳣�ļ���"
         Height          =   180
         Index           =   1
         Left            =   3045
         TabIndex        =   9
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label lblInformation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ϴ��ļ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   465
         TabIndex        =   8
         Top             =   90
         Width           =   1260
      End
      Begin VB.Image imgInformation 
         Height          =   324
         Left            =   12
         Picture         =   "frmFilesUpload.frx":98C6C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15684
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   7185
      Left            =   60
      TabIndex        =   3
      Top             =   2055
      Visible         =   0   'False
      Width           =   15690
      _cx             =   27675
      _cy             =   12674
      Appearance      =   0
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
      BackColorBkg    =   -2147483638
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.CommandButton cmdUpload 
      Caption         =   "�ϴ��µ��ļ�(&Q)"
      Enabled         =   0   'False
      Height          =   288
      Left            =   12480
      TabIndex        =   2
      Top             =   1155
      Width           =   1500
   End
   Begin VB.CommandButton cmdAllUpLoad 
      Caption         =   "�ϴ������ļ�(&T)"
      Height          =   288
      Left            =   14205
      TabIndex        =   14
      Top             =   1155
      Width           =   1500
   End
   Begin VB.CommandButton cmdMD5Check 
      Caption         =   "���¼��(&D)"
      Height          =   288
      Left            =   11055
      TabIndex        =   5
      Top             =   1155
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar pgbThis 
      Height          =   330
      Left            =   495
      TabIndex        =   4
      Top             =   1560
      Width           =   9615
      _ExtentX        =   16955
      _ExtentY        =   572
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&Q)"
      Height          =   288
      Left            =   8085
      TabIndex        =   1
      Top             =   1155
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSever 
      Height          =   1155
      Left            =   14490
      TabIndex        =   6
      Top             =   1515
      Visible         =   0   'False
      Width           =   1305
      _cx             =   2302
      _cy             =   2037
      Appearance      =   0
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   16
      Top             =   1215
      Width           =   120
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   1245
      Width           =   90
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   1620
      TabIndex        =   0
      Top             =   1260
      Width           =   90
   End
End
Attribute VB_Name = "frmFilesUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'StateDisplay-״̬��ʾ
Private Const SDP_׼������ = "׼������"
Private Const SDP_׼���ϴ� = "׼���ϴ�"
Private Const SDP_������� = "�������"
Private Const SDP_״̬�쳣 = "״̬�쳣"
Private Const SDP_�Ѿ��ϴ� = "�Ѿ��ϴ�"
Private Const SDP_�ϴ�ʧ�� = "�ϴ�ʧ��"
Private Const SDP_�ռ���� = "�ռ����"
Private Const SDP_�����ռ� = "�����ռ�"

'StateColor-״̬��ɫ
Private Const SC_��ɫ = vbRed
Private Const SC_��ɫ = 2188065  'RGB(33, 99, 33)
Private Const SC_��ɫ = 9109504  'RGB(0, 0, 139)
Private Const SC_��ɫ = 32767      'RGB(255, 127, 0)

'�ļ�״̬
Private Enum FilesState
    FS_Ĭ������ = 0 '�����ϴ�
    FS_״̬�쳣 = 1 '�쳣�������ϴ�
    FS_������� = 2 '�����ϴ�
    FS_������¾����ļ� = 3
    FS_׼���ϴ������ļ� = 4
    FS_�Ѿ��ϴ� = 5 '�ϴ��ɹ�
    FS_�ϴ�ʧ�� = -1 '�ϴ�ʧ��
End Enum

Private Enum UpLoadCol
    Col_��� = 0
    Col_�ļ� = 1
    Col_״̬ = 2 '״ֵ̬ 0-���� 1-��������ȱʧ(�����ļ�������) 2-�����ļ������� 3-������� 4-���浫�����ϴ� 5-�Ѿ��ϴ� 6-�ϴ�ʧ�� 7 -���浫�����ϴ�
    Col_���� = 3
    Col_��ǰ�汾 = 4
    Col_��׼�汾 = 5
    Col_��װ·�� = 6
    Col_ϵͳ = 7
    Col_�޸����� = 8
    Col_ҵ�񲿼� = 9
    Col_�ļ�˵�� = 10
    Col_��ǰmd5 = 11
    Col_��׼md5 = 12
    Col_����md5 = 13
    Col_�ļ���ַ = 14
    Col_�ռ���ַ = 15  '�ռ����ļ���ַ
    Col_�ռ��ļ� = 16 '�ռ����ļ�����
    Col_�ļ����� = 17
    Col_���� = 18
End Enum

Private Enum UpSeverCol
    Col_��� = 0
    Col_���� = 1
    Col_��ַ = 2
    Col_�û��� = 3
    Col_���� = 4
    Col_�˿� = 5
    Col_�ϴ�״̬ = 6
    Col_���������� = 7
End Enum

Private Enum UploadResult
    Res_δ�ϴ� = 0
    Res_�ϴ��ɹ� = 1
    Res_�ϴ�ʧ�� = 2
    Res_δ֪���� = 3
End Enum

Private Enum lblItemNum
    LN_�����ϴ��ļ� = 0
    LN_״̬�쳣�ļ� = 1
    LN_���ϴ��ļ� = 2
    LN_���ϴ����ļ� = 3
    LN_��龯���ļ� = 4
    LN_���ϴ��ļ� = 5
'    LN_��������ļ� = 5
    LN_lbl���� = 6
End Enum

Private Enum pbgstatus
    Sta_�ٷֱ� = 0
    Sta_��ǰ���� = 1
    Sta_�������� = 2
    Sta_״̬���� = 3
End Enum

Private mrsTemp As New ADODB.Recordset
Private strParstWarning As String 'ȱʧ��������
Private mstrSQL As String
Private mintUpFilesCount As Integer '�ϴ��ļ�����
Private mblnCheckMD5Tag As Boolean '���ڼ��MD5��־
Private mblnUploadTag As Boolean '�����ϴ���־
Private mblnAllUploadTag As Boolean '����ȫ���ϴ���־
Private mblnAllUpload As Boolean

Private mobjFile As New FileSystemObject
Private mstrScratchFilePath As String '��ʱ�ļ�Ŀ¼
Private mblnUpLoadSuccess As Boolean '�ϴ��ɹ�������MD5��־
Private mstrSuccessUploadSever As String '�ɹ��ϴ�������
Private mcllPath As Collection '��װ·��ת��ʵ��·������

Private mlngAbnormal  As Long '�쳣
Private mlngCorrect As Long '����
Private mlngUnchanged As Long '�������
Private mlngWarning As Long  '����
Private mlngUpload As Long
Private mlngNewFile As Long

Public Function ShowMe() As Boolean
    '��ʱĿ¼��ʼ��
    If gblnInIDE Then
        mstrScratchFilePath = "C:\APPSOFT\ZLUPTMP"
    Else
        mstrScratchFilePath = App.Path & "\ZLUPTMP"
    End If
    Me.Show 1, frmMDIMain
End Function

Private Sub cmdAllUpLoad_Click()
    Dim strSQL As String
    
    strSQL = "update zlfilesupgrade set MD5 = Null"
    gcnOracle.Execute strSQL
    '������ԭ
    mblnAllUpload = True
    Call AllUploadRestore
'    Call cmdMD5Check_Click
    Call cmdUpload_Click
    mblnAllUpload = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMD5Check_Click()
    '����MD5�ļ�·�����
    Set mcllPath = CheckAndAdjustFolder()
    ShowStatus "", "", "��ʼ���", 0
    lblInformation_Click (LN_�����ϴ��ļ�)
    Call DataLoad
    Call DataCheck
    Call FilesMD5Check
    '�����ɺ��Զ��л�����
    If Val(Split(lblInformation(LN_״̬�쳣�ļ�), "��")(1)) > 0 Then
        lblInformation_Click (LN_״̬�쳣�ļ�)
        ShowStatus "", "", "�����ɣ������쳣�ļ�", 0, , False
    ElseIf Val(Split(lblInformation(LN_��龯���ļ�), "��")(1)) > 0 Then
        lblInformation_Click (LN_��龯���ļ�)
        ShowStatus "", "", "�����ɣ����龯���ļ�", 0, , False
    ElseIf Val(Split(lblInformation(LN_���ϴ����ļ�), "��")(1)) > 0 Then
        lblInformation_Click (LN_���ϴ����ļ�)
        ShowStatus "", "", "�����ɿ����ϴ��µ��ļ�", 0, , False
    Else
        lblInformation_Click (LN_���ϴ��ļ�)
        ShowStatus "", "", "�����ɿ����ϴ��ļ�", 0, , False
    End If
End Sub

Private Sub cmdUpload_Click()
    Dim strNumber As String
    Dim strSeverType As String
    Dim strServerAddress As String
    Dim strUser As String
    Dim strPassword As String
    Dim strPort As String
    Dim strBatch As String
    Dim i As Integer
    Dim strErrInfor As String '������Ϣ
    On Error GoTo errH
    
    If mblnAllUpload Then
        Call lblInformation_Click(LN_���ϴ��ļ�)
    Else
        Call lblInformation_Click(LN_���ϴ����ļ�)
    End If
    Call ControlVisible(False)
    '�ռ��ļ���ѹ������ʱ�ļ���
    If FilesCollections() = True Then
        '���7z.exe����ϵͳ����
        Call fun_KillProcess(PROAPPCTION)
        '�����ռ��ļ���������
'        Call FloderToClipBoard(mstrScratchFilePath)
    End If
    
    '��ȡ�������б�
    Call SeverDataLoad
    imgCaption.Picture = imgList.ListImages("�ϴ��ļ�").Picture
    lblEXP(2).ForeColor = vbBlue
    With vsfSever
        If .Rows < 1 Then MsgBox "����������һ���ϴ�������!", vbDefaultButton1 + vbInformation, gstrSysName: Exit Sub
        For i = 1 To .Rows - 1
            strNumber = .TextMatrix(i, Col_���): strSeverType = .TextMatrix(i, Col_����)
            strServerAddress = .TextMatrix(i, Col_��ַ): strUser = .TextMatrix(i, Col_�û���)
            strPassword = Decipher(Trim(.TextMatrix(i, Col_����))): strPort = .TextMatrix(i, Col_�˿�)
            
            Select Case strSeverType
                Case "0" '����
                    If CopyFileToShareServer(strNumber, strServerAddress, strUser, strPassword, strErrInfor) = True Then
                        .TextMatrix(i, Col_�ϴ�״̬) = str(Res_�ϴ��ɹ�)
                    Else
                        If strErrInfor <> "" Then
                            MsgBox strErrInfor, vbDefaultButton1 + vbInformation, gstrSysName
                            .TextMatrix(i, Col_�ϴ�״̬) = str(Res_δ֪����)
                       Else
                            .TextMatrix(i, Col_�ϴ�״̬) = str(Res_�ϴ�ʧ��)
                       End If
                    End If
                Case "1" 'FTP
                    If CopyFileToFTPServer(strNumber, strServerAddress, strUser, strPassword, strPort, strErrInfor) = True Then
                        .TextMatrix(i, Col_�ϴ�״̬) = str(Res_�ϴ��ɹ�)
                    Else
                        If strErrInfor <> "" Then
                            MsgBox strErrInfor, vbDefaultButton1 + vbInformation, gstrSysName
                            .TextMatrix(i, Col_�ϴ�״̬) = str(Res_δ֪����)
                       Else
                            .TextMatrix(i, Col_�ϴ�״̬) = str(Res_�ϴ�ʧ��)
                       End If
                    End If
            End Select
        Next
        If UpdateMD5 = True Then '����MD5
            strBatch = BatchLoad
            strBatch = Trim(str(Nvl(strBatch) + 1))
            BatchUpdate strBatch
        End If
        imgCaption.Picture = imgList.ListImages("�ļ�").Picture
        lblEXP(2).ForeColor = vbBlack
        Call ControlVisible(True)
    End With
    cmdUpload.Enabled = False
    Call UpLoadFilesCount
    Call lblInformation_Click(LN_���ϴ��ļ�)
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    Call ControlVisible(True)
    If False Then
        Resume
    End If
End Sub

Private Sub Command1_Click()
    Dim Data1 As String
    Dim Data2 As String
    Dim strFilePath As String
    Dim i As Long
    
    With vsfMain
        i = 2
        strFilePath = .TextMatrix(i, Col_�ļ���ַ)
        Data1 = Format(FileDateTime(strFilePath), "yyyy-MM-DD hh:mm:ss")
        Data2 = .TextMatrix(i, Col_�޸�����)
        Data2 = "2000-04-22 20:10:56"
        Call CompareDate(Data1, Data2)
    End With
End Sub

Private Sub Form_Activate()
    Dim lngWidth As Long
    Dim i As Integer
    
    '�ؼ���ͼ��ʼ��
    pgbThis.Move 100, 600 + picHelp.Height, Me.ScaleWidth - 200, 330
    vsfMain.Move 100, 1000 + picHelp.Height, Me.ScaleWidth - 200, Me.ScaleHeight - 1490 - picHelp.Height

    picInformation.Move 100, 9800, Me.ScaleWidth - 200, 400
    imgInformation.Move 0, 0, picInformation.ScaleWidth, picInformation.ScaleHeight
    imgInformation.Picture = imgList.ListImages.Item(LN_�����ϴ��ļ� + 1).Picture

    lngWidth = picInformation.Width / 6
    For i = 0 To LN_lbl���� - 1
        lblInformation.Item(i).Move (i * lngWidth) + ((lngWidth) - lblInformation.Item(i).Width) / 2, (picInformation.ScaleHeight - lblInformation.Item(i).Height) / 2
    Next
    lblInformation.Item(LN_�����ϴ��ļ�).ForeColor = vbBlack
    lblInformation.Item(LN_�����ϴ��ļ�).FontBold = True
    lblInformation.Item(LN_״̬�쳣�ļ�).ForeColor = vbRed
    lblInformation.Item(LN_���ϴ��ļ�).ForeColor = SC_��ɫ
    lblInformation.Item(LN_��龯���ļ�).ForeColor = SC_��ɫ
    lblInformation.Item(LN_���ϴ��ļ�).ForeColor = SC_��ɫ
    lblInformation.Item(LN_���ϴ����ļ�).ForeColor = SC_��ɫ
'    lblInformation.Item(LN_��������ļ�).ForeColor = SC_��ɫ
    vsfMain.Visible = True
    Call cmdMD5Check_Click
End Sub

Private Sub Form_Load()
'   Call cmdMD5Check_Click
End Sub

Private Sub Form_Resize()
'    vsfMain.Move 50, 1200, Me.ScaleWidth - 100, Me.ScaleHeight - 1000
End Sub

Private Function GetSystemName(ByVal strNum As String) As String
'����ϵͳ��ţ���ö�Ӧϵͳ���ƣ���δ�ҵ�
On err GoTo errH
    Select Case strNum
        Case "1", "100"
            GetSystemName = "ҽԺϵͳ��׼��"
        Case "2", "200"
            GetSystemName = "���¹���ϵͳ"
        Case "3", "300"
            GetSystemName = "��������ϵͳ"
        Case "4", "400"
            GetSystemName = "���ʹ�Ӧϵͳ"
        Case "5", "500"
            GetSystemName = "�������ϵͳ"
        Case "6", "600"
            GetSystemName = "�豸����ϵͳ"
        Case "7", "700"
            GetSystemName = "�ɱ�Ч�����ϵͳ"
        Case "21", "2100"
            GetSystemName = "������ϵͳ"
        Case "22", "2200"
            GetSystemName = "Ѫ�����ϵͳ"
        Case "23", "2300"
            GetSystemName = "Ժ�й���ϵͳ"
        Case "24", "2400"
            GetSystemName = "�������ϵͳ"
        Case "25", "2500"
            GetSystemName = "�ٴ��������ϵͳ"
        Case "26", "2600"
            GetSystemName = "������������ϵͳ"
    End Select
    Exit Function

errH:
    If False Then
        Resume
    End If
End Function

'���ݶ�ȡ�������ͼ����
Public Sub DataLoad()
    Dim i, j As Long
    Dim strSQL       As String
    Dim strTemp As String
    Dim arrSys As Variant
    On Error GoTo errH

    With vsfMain
        .Redraw = flexRDNone
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cols = Col_����
'        Exit Sub
        .TextMatrix(0, Col_���) = ""
        .Cell(flexcpAlignment, 0, Col_���) = flexAlignCenterCenter
        .ColWidth(Col_���) = 400
        
        .TextMatrix(0, Col_״̬) = "״̬"
        .Cell(flexcpAlignment, 0, Col_״̬) = flexAlignCenterCenter
        .ColWidth(Col_״̬) = 1000
        
        .TextMatrix(0, Col_����) = "�ļ���龯��"
        .Cell(flexcpAlignment, 0, Col_����) = flexAlignCenterCenter
        .ColWidth(Col_����) = 4200
        
        .TextMatrix(0, Col_�ļ�) = "�ļ�"
        .Cell(flexcpAlignment, 0, Col_�ļ�) = flexAlignCenterCenter
        .ColWidth(Col_�ļ�) = 2400
        
        .TextMatrix(0, Col_��ǰ�汾) = "��ǰ�汾"
        .Cell(flexcpAlignment, 0, Col_��ǰ�汾) = flexAlignCenterCenter
        .ColWidth(Col_��ǰ�汾) = 900
        
        .TextMatrix(0, Col_��׼�汾) = "��׼�汾"
        .Cell(flexcpAlignment, 0, Col_��׼�汾) = flexAlignCenterCenter
        .ColWidth(Col_��׼�汾) = 900
        
        .TextMatrix(0, Col_��װ·��) = "��װ·��"
        .Cell(flexcpAlignment, 0, Col_��װ·��) = flexAlignCenterCenter
        .ColWidth(Col_��װ·��) = 1800
        
        .TextMatrix(0, Col_ϵͳ) = "ϵͳ"
        .Cell(flexcpAlignment, 0, Col_ϵͳ) = flexAlignCenterCenter
        .ColWidth(Col_ϵͳ) = 1000
        .ColHidden(Col_ϵͳ) = True
        
        .TextMatrix(0, Col_�޸�����) = "�޸�����"
        .Cell(flexcpAlignment, 0, Col_�޸�����) = flexAlignCenterCenter
        .ColWidth(Col_�޸�����) = 1800

        .TextMatrix(0, Col_ҵ�񲿼�) = "ҵ�񲿼�"
'        .Cell(flexcpAlignment, 0, Col_ҵ�񲿼�) = flexAlignCenterCenter
        .ColWidth(Col_ҵ�񲿼�) = 1000
        .ColHidden(Col_ҵ�񲿼�) = True
        
        .TextMatrix(0, Col_�ļ�˵��) = "�ļ�˵��"
        .Cell(flexcpAlignment, 0, Col_�ļ�˵��) = flexAlignCenterCenter
        .ColWidth(Col_�ļ�˵��) = 400
        
        .TextMatrix(0, Col_��ǰmd5) = "��ǰmd5"
        .ColWidth(Col_��ǰmd5) = 10
        .ColHidden(Col_��ǰmd5) = True
        
        .TextMatrix(0, Col_��׼md5) = "��ǰmd5"
        .ColWidth(Col_��׼md5) = 10
        .ColHidden(Col_��׼md5) = True

        .TextMatrix(0, Col_����md5) = "����md5"
        .ColWidth(Col_����md5) = 10
        .ColHidden(Col_����md5) = True
        
        .TextMatrix(0, Col_�ļ���ַ) = "�ļ���ַ"
        .ColWidth(Col_�ļ���ַ) = 10
        .ColHidden(Col_�ļ���ַ) = True
        
        .TextMatrix(0, Col_�ռ���ַ) = "�ռ��ļ���ַ"
        .ColWidth(Col_�ռ���ַ) = 10
        .ColHidden(Col_�ռ���ַ) = True
        
        .TextMatrix(0, Col_�ռ��ļ�) = "�ռ��ļ�����"
        .ColWidth(Col_�ռ��ļ�) = 10
        .ColHidden(Col_�ռ��ļ�) = True
        
        .TextMatrix(0, Col_�ļ�����) = "�ռ��ļ�����"
        .ColWidth(Col_�ļ�����) = 10
        .ColHidden(Col_�ļ�����) = True

'        strSQL = "Select A.�ļ��� As �ļ�, A.�汾�� As �汾��,b.�汾�� as ��׼�汾, A.��װ·�� As ��װ·��, A.����ϵͳ As ϵͳ,A.�޸����� As �޸�����, A.ҵ�񲿼� As ҵ�񲿼�, A.�ļ�˵�� As �ļ�˵�� " & _
'                      "From zlFilesUpgrade A Left Join Zlfiles B " & _
'                      "On A.�ļ��� = B.���� " & _
'                      "order by �ļ�"
        strSQL = "Select Nvl(a.�ļ���, b.����) As ��������, a.�ļ��汾�� As ��ǰ�汾, b.�汾�� As ��׼�汾, a.��װ·�� As ��װ·��, a.����ϵͳ As ϵͳ, a.�޸����� As �޸�����, " & _
                      "a.ҵ�񲿼� As ҵ�񲿼�, a.�ļ�˵�� As �ļ�˵��, a.Md5 As ��ǰmd5, b.��׼md5, Decode(a.�ļ���, Null, 1, Decode(b.����, Null, 2)) As ����, Decode(a.�ļ�����,null,b.�ļ�����,a.�ļ�����) as �ļ����� " & _
                      "From zlFilesUpgrade A Full Join Zlfiles B " & _
                      "On A.�ļ��� = B.���� " & _
                      "order by ��������"
        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)

        '��������
        .Rows = mrsTemp.RecordCount + 1
        i = 1
        Do Until mrsTemp.EOF
            .TextMatrix(i, Col_���) = i
            .Cell(flexcpAlignment, i, Col_���) = flexAlignCenterCenter
            
            .TextMatrix(i, Col_״̬) = ""  '״̬
            .Cell(flexcpAlignment, i, Col_״̬) = flexAlignCenterCenter
  
            .TextMatrix(i, Col_����) = ""
            .Cell(flexcpData, i, Col_����) = Trim(Nvl(mrsTemp.Fields("����"), ""))
            .Cell(flexcpAlignment, i, Col_����) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_�ļ�) = Nvl(mrsTemp.Fields("��������"), "")
            .Cell(flexcpAlignment, i, Col_�ļ�) = flexAlignLeftCenter
            
            
            strTemp = Nvl(mrsTemp.Fields("��ǰ�汾"), "")
'            .Cell(flexcpData, i, Col_��ǰ�汾) = strTemp 'δת���汾��
'            strTemp = GetFileVision(strTemp)
            .TextMatrix(i, Col_��ǰ�汾) = strTemp  'ת����汾��
            .Cell(flexcpAlignment, i, Col_��ǰ�汾) = flexAlignLeftCenter
            
            strTemp = Nvl(mrsTemp.Fields("��׼�汾"), "")
'            .Cell(flexcpData, i, Col_��׼�汾) = strTemp
'            strTemp = GetFileVision(strTemp)
            .TextMatrix(i, Col_��׼�汾) = strTemp
            .Cell(flexcpAlignment, i, Col_��׼�汾) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_��װ·��) = Nvl(mrsTemp.Fields("��װ·��"), "")
            .Cell(flexcpAlignment, i, Col_��װ·��) = flexAlignLeftCenter

            strTemp = Nvl(mrsTemp.Fields("ϵͳ"), "")

            If Trim(strTemp) <> "" Then
                arrSys = Split(Trim(strTemp), ",")
                strTemp = ""
                For j = 0 To UBound(arrSys)
                    If GetSystemName(arrSys(j)) <> "" Then strTemp = strTemp & "��" & GetSystemName(arrSys(j))
                Next
                strTemp = Mid(strTemp, 2)
            Else
                strTemp = ""
            End If
            .TextMatrix(i, Col_ϵͳ) = strTemp
            .Cell(flexcpAlignment, i, Col_ϵͳ) = flexAlignLeftCenter

            .TextMatrix(i, Col_�޸�����) = Nvl(mrsTemp.Fields("�޸�����"), "")
            .Cell(flexcpAlignment, i, Col_�޸�����) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_ҵ�񲿼�) = Nvl(mrsTemp.Fields("ҵ�񲿼�"), "")
            .Cell(flexcpAlignment, i, Col_ҵ�񲿼�) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_�ļ�˵��) = Nvl(mrsTemp.Fields("�ļ�˵��"), "")
            .Cell(flexcpAlignment, i, Col_�ļ�˵��) = flexAlignLeftCenter
            
            .TextMatrix(i, Col_��ǰmd5) = Trim(Nvl(mrsTemp.Fields("��ǰmd5"), ""))

            .TextMatrix(i, Col_��׼md5) = Trim(Nvl(mrsTemp.Fields("��׼md5"), ""))
            
            .TextMatrix(i, Col_�ļ�����) = Trim(Nvl(mrsTemp.Fields("�ļ�����"), ""))
            
            mrsTemp.MoveNext
            i = i + 1
        Loop
        
        'ѡ�п���
        .FocusRect = flexFocusSolid
        '���һ���Զ��п�
        .ExtendLastCol = True
        '�����������
        .ScrollTrack = True
        '�Զ�����
        .WordWrap = True
        '�и�����
        .RowHeightMin = 300
        .RowHeightMax = 300
        '���������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

'��ȡ��Ҫ�ϴ��ķ�����
Public Sub SeverDataLoad()
    Dim i, j As Long
    Dim strSQL       As String
    Dim strTemp As String
    On Error GoTo errH

    With vsfSever
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cols = Col_����������
'        Exit Sub
        .TextMatrix(0, Col_���) = "���������"
        .Cell(flexcpAlignment, 0, Col_���) = flexAlignCenterCenter
        .ColWidth(Col_���) = 800
        
        .TextMatrix(0, Col_����) = "����������"
        .Cell(flexcpAlignment, 0, Col_����) = flexAlignCenterCenter
        .ColWidth(Col_����) = 800
        
        .TextMatrix(0, Col_��ַ) = "��������ַ"
        .Cell(flexcpAlignment, 0, Col_��ַ) = flexAlignCenterCenter
        .ColWidth(Col_��ַ) = 3000
        
        .TextMatrix(0, Col_�û���) = "�û���"
        .Cell(flexcpAlignment, 0, Col_�û���) = flexAlignCenterCenter
        .ColWidth(Col_�û���) = 1500
        
        .TextMatrix(0, Col_����) = "����"
        .Cell(flexcpAlignment, 0, Col_����) = flexAlignCenterCenter
        .ColWidth(Col_����) = 1500
        
        .TextMatrix(0, Col_�˿�) = "�˿�"
        .Cell(flexcpAlignment, 0, Col_�˿�) = flexAlignCenterCenter
        .ColWidth(Col_�˿�) = 1600
        
        .TextMatrix(0, Col_�ϴ�״̬) = "״̬"
        .Cell(flexcpAlignment, 0, Col_�ϴ�״̬) = flexAlignCenterCenter
        .ColWidth(Col_�ϴ�״̬) = 1600

        strSQL = "Select ��� As ���, ���� As ����, λ�� As ��ַ, �û��� As �û���, ���� As ����, �˿� As �˿� " & _
                      "From Zlupgradeserver " & _
                      "Where �Ƿ����� = 1 " & _
                      "order by ���"
        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)

        '��������
        .Rows = mrsTemp.RecordCount + 1
        i = 1
        Do Until mrsTemp.EOF
            .TextMatrix(i, Col_���) = Trim(Nvl(mrsTemp.Fields("���"), ""))

            .TextMatrix(i, Col_����) = Trim(Nvl(mrsTemp.Fields("����"), ""))
    
            .TextMatrix(i, Col_��ַ) = Trim(Nvl(mrsTemp.Fields("��ַ"), ""))
            
            .TextMatrix(i, Col_�û���) = Trim(Nvl(mrsTemp.Fields("�û���"), ""))

            .TextMatrix(i, Col_����) = Trim(Nvl(mrsTemp.Fields("����"), ""))

            .TextMatrix(i, Col_�˿�) = Trim(Nvl(mrsTemp.Fields("�˿�"), ""))
            
            .TextMatrix(i, Col_�ϴ�״̬) = Trim(str(Res_δ�ϴ�))
            
            mrsTemp.MoveNext
            i = i + 1
        Loop
        
    End With
    Exit Sub
errH:
    If False Then
        Resume
    End If
End Sub

'MD5���
Public Function FilesMD5Check()
    Dim strMD5 As String '�����ļ�MD5
    Dim strMD5Upgrade As String '�����ļ�MD5(��ǰ)
    Dim strMD5Standard As String '��׼�ļ�MD5
    Dim lngPercent As Long '�������ٷֱ�
    Dim intPercent As Integer '������ʾ�ٷֱ�
    Dim i As Long
    Dim strSQL As String
    Dim strTemp As String
    
    On Error Resume Next
    Call ControlVisible(False)
    imgCaption.Picture = imgList.ListImages("����ļ�").Picture
    lblEXP(0).ForeColor = vbBlue
    With vsfMain
        If .Rows < .FixedRows Then Exit Function
        Call ShowStatus("���ڼ��", "", "", 0, .Rows - 1)
        mlngCorrect = 0: mlngUnchanged = 0: mlngWarning = 0: mlngNewFile = 0

        For i = .FixedRows To .Rows - 1
'            .Row = i
'            If .Rows - (i + 14) > 0 Then '����ѡ�������м�λ��
'                .ShowCell i + 14, Col_�ļ�
'            Else
'                .ShowCell i, Col_�ļ�
'            End If
            ShowCenterRow i
            
            Call ShowStatus("���ڼ��", .TextMatrix(i, Col_�ļ�), "", i)
            If .Cell(flexcpData, i, Col_״̬) = "0" Then
                strMD5Standard = .TextMatrix(i, Col_��׼md5)
                strMD5Upgrade = .TextMatrix(i, Col_��ǰmd5)
                DoEvents '��ֹ���濨��
                strMD5 = FileMD5(Trim(UCase(.TextMatrix(i, Col_�ļ���ַ))))
                .TextMatrix(i, Col_����md5) = strMD5
                
                If strMD5Upgrade = strMD5 Then '����������MD5��ͬ����Ҫ����
                    If strMD5 = strMD5Standard Then
                        Call FileStateSet(FS_�������, i, "�ļ������ڲ���,����Ҫ����")
                    Else
                        If strMD5Standard = "" Then '��׼���������ڸ��ļ�
                            strTemp = "���棺��׼�ļ��嵥(zlFiles)�в����ڸ��ļ�"
                            If .TextMatrix(i, Col_�ļ�����) = "4" Then strTemp = strTemp & "(��������)"
                        Else
                            strTemp = "���棺�����ļ����׼�ļ���������ȷ���ļ�����"
                        End If
                        Call FileStateSet(FS_������¾����ļ�, i, strTemp)
                        mlngWarning = mlngWarning + 1
                    End If
                    mlngUnchanged = mlngUnchanged + 1
                    mlngCorrect = mlngCorrect + 1
                Else '����������MD5����ͬ��Ҫ����
                    If strMD5 <> strMD5Standard Then '����MD5���׼MD5��ͬ
                        If strMD5Standard = "" Then '��׼���������ڸ��ļ�
                            strTemp = "���棺��׼�ļ��嵥(zlFiles)�в����ڸ��ļ�"
                            If .TextMatrix(i, Col_�ļ�����) = "4" Then strTemp = strTemp & "(��������)"
                        Else
                            strTemp = "���棺�����ļ����׼�ļ���������ȷ���ļ�����"
                        End If
                        Call FileStateSet(FS_׼���ϴ������ļ�, i, strTemp)
                        mlngWarning = mlngWarning + 1
                        mlngCorrect = mlngCorrect + 1
                        mlngNewFile = mlngNewFile + 1
                    Else
                        Call FileStateSet(FS_Ĭ������, i)
                        mlngCorrect = mlngCorrect + 1
                        mlngNewFile = mlngNewFile + 1
                    End If
                End If
            End If
            lblInformation.Item(LN_���ϴ��ļ�).Caption = Split(lblInformation.Item(LN_���ϴ��ļ�).Caption, "��")(0) & "��" & str(mlngCorrect)
            lblInformation.Item(LN_��龯���ļ�).Caption = Split(lblInformation.Item(LN_��龯���ļ�).Caption, "��")(0) & "��" & str(mlngWarning)
            lblInformation.Item(LN_���ϴ����ļ�).Caption = Split(lblInformation.Item(LN_���ϴ����ļ�).Caption, "��")(0) & "��" & str(mlngNewFile)
'            lblInformation.Item(LN_��������ļ�).Caption = Split(lblInformation.Item(LN_��������ļ�).Caption, "��")(0) & "��" & str(mlngUnchanged)
        Next
        .Row = 1
        .ShowCell 1, Col_�ļ�
        Call ShowStatus("������", "", "", pgbThis.Max)
        imgCaption.Picture = imgList.ListImages("�ļ�").Picture
        lblEXP(0).ForeColor = vbBlack
        Call ControlVisible(True)
        Call UpLoadFilesCount
    End With
End Function

'���ݼ�飬��������ͼ����������״̬
'״ֵ̬ 0-���� 1-��׼����ȱʧ(�����ļ�������) 2-�����ļ������� 3-������� 4-���浫�����ϴ� 5-�Ѿ��ϴ�
Private Sub DataCheck()
    Dim objFile As New FileSystemObject
    Dim strFileName As String
    Dim strTemp As String
    Dim strStateContent As String
    Dim i As Long
    Dim lngAbnormal As Long '�쳣�ļ�
    On Error GoTo errH
    '�ļ����ڼ�飬�����ʾ
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        For i = .FixedRows To .Rows - 1
            strTemp = .Cell(flexcpData, i, Col_����)
            strStateContent = ""
            If strTemp <> "1" Then 'Ϊ1˵����׼����ȱʧ
                If Trim(.TextMatrix(i, Col_��װ·��)) <> "" Then
                    '��װ·��ת����ʵ��·��
                    strFileName = mcllPath("K_" & UCase(.TextMatrix(i, Col_��װ·��))) & "\" & UCase(.TextMatrix(i, Col_�ļ�))
                    .TextMatrix(i, Col_�ļ���ַ) = UCase(Trim(strFileName))
                    '�����ļ��治����
                    If objFile.FileExists(strFileName) = False Then
                        strStateContent = "�����ļ�ȱʧ"
                        If .TextMatrix(i, Col_�ļ�����) = "4" Then strStateContent = strStateContent & "(��������)"
                        Call FileStateSet(FS_״̬�쳣, i, strStateContent)
                    Else
                        Call FileStateSet(FS_Ĭ������, i, strStateContent, "׼������", vbBlack)
                    End If
                Else
                    strStateContent = "��װ·��ȱʧ"
                    If .TextMatrix(i, Col_�ļ�����) = "4" Then strStateContent = strStateContent & "(��������)"
                    Call FileStateSet(FS_״̬�쳣, i, strStateContent)
                End If
            Else
                strStateContent = "�����ļ��嵥(zlFilesUpgrade)ȱʧ���ļ�"
                If .TextMatrix(i, Col_�ļ�����) = "4" Then strStateContent = strStateContent & "(��������)"
                Call FileStateSet(FS_״̬�쳣, i, strStateContent)
            End If
        Next
    End With
    Call UpLoadFilesCount
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Function CheckFTPServer(ByVal strIp As String, ByVal strUser As String, ByVal strPass As String, ByVal strPort As String) As Boolean
    '-----------------------------------------------------------------------------
    '����:��鵱ǰ��FTP�������Ƿ���ȷ
    '����:��ǰ���ļ��������ĸ�����ȷ,����true,���򷵻�False
    '����:����ԭ
    '����:2016/07/05
    'strIp - FTP��ַ
    'strUser - �û���
    'strPass - ����
    'strPort - �˿�
    '-----------------------------------------------------------------------------
    On Error GoTo errHand:
    
    If strIp = "" Or strUser = "" Or strPass = "" Or strPort = "" Then
        CheckFTPServer = False
        Exit Function
    End If
    
    If IsFtpServer(Trim(strIp), Trim(strUser), Trim(strPass), Trim(strPort)) Then
        CheckFTPServer = True
    Else
        CheckFTPServer = False
        MsgBox "������������������������FTP����������!", vbInformation + vbDefaultButton1, gstrSysName
    End If
    Exit Function
    
errHand:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function CheckShareServer(ByVal strAddress As String, ByVal strUser As String, ByVal strPass As String) As Boolean
    '-----------------------------------------------------------------------------
    '����:��鵱ǰ���ļ��������Ƿ���ȷ
    '����:��ǰ���ļ��������ĸ�����ȷ,����true,���򷵻�False
    '����:����ԭ
    '����:2016/07/05
    'strAddress - ��ַ
    'strUser - �û�
    'strPass - ����
    '-----------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT

    On Error GoTo errHand:
    
    If strAddress = "" Or strUser = "" Or strPass = "" Then
        CheckShareServer = False
        Exit Function
    End If
    
    If FindFile(Trim(strAddress)) = False Then
        If IsNetServer(Trim(strAddress), Trim(strUser), Trim(strPass)) = False Then
            MsgBox "�����������������������鹲����������ã�", vbInformation + vbDefaultButton1, gstrSysName
            CheckShareServer = False
            Exit Function
        End If
    End If
    Call CancelNetServer(Trim(strAddress))
    CheckShareServer = True
    Exit Function
errHand:
    If err Then
        MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function FindFile(ByVal strFileName As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--����:����ָ�����ļ����ļ��Ƿ����
    '--����: ������ڴ��ļ�ΪTrue,����ΪFlase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

'��ȡ�汾��ֱ����ʾֵ
Private Function GetFileVision(ByVal strVision As String) As String
    Dim lng�汾�� As Variant
    Dim str�汾�� As String
    If Len(strVision) > 0 Then
        lng�汾�� = strVision
        str�汾�� = Int(lng�汾�� / 10 ^ 8)
        If Len(lng�汾��) > 9 Then
            lng�汾�� = Right(lng�汾��, 9) Mod (10 ^ 8)
        Else
            lng�汾�� = lng�汾�� Mod (10 ^ 8)
        End If
        
        str�汾�� = str�汾�� & "." & Int(lng�汾�� / 10 ^ 4)
        lng�汾�� = lng�汾�� Mod 10 ^ 4
        str�汾�� = str�汾�� & "." & lng�汾��
        GetFileVision = str�汾��
    End If
End Function

Private Sub UpLoadFilesCount()
    Dim i As Long
    Dim lngAbnormal As Long '�쳣
    Dim lngCorrect As Long '����
    Dim lngUnchanged As Long '�������
    Dim lngWarning As Long  '����
    Dim lngUpload As Long '�Ѿ��ϴ�
    Dim lngNewFile As Long '�Ѿ��ϴ�
    
    With vsfMain
        If .Rows < 1 Then Exit Sub
'        lblInformation.Item(Num_�ļ�����).Caption = str(.Rows - 1)
        lngAbnormal = 0: lngCorrect = 0: lngUnchanged = 0: lngWarning = 0: lngUpload = 0: lngNewFile = 0
        
        For i = 1 To .Rows - 1
            Select Case .Cell(flexcpData, i, Col_״̬)
                Case FS_Ĭ������ '����(��Ҫ)�ϴ�������
                    lngNewFile = lngNewFile + 1
                    lngCorrect = lngCorrect + 1
                Case FS_״̬�쳣 '�����ļ�ȱʧ���쳣
                    lngAbnormal = lngAbnormal + 1
                Case FS_�������
                    lngUnchanged = lngUnchanged + 1
                    lngCorrect = lngCorrect + 1
                Case FS_������¾����ļ� '������£�����
                    lngUnchanged = lngUnchanged + 1
                    lngWarning = lngWarning + 1
                    lngCorrect = lngCorrect + 1
                Case FS_׼���ϴ������ļ� '���棬���ϴ�
                    lngNewFile = lngNewFile + 1
                    lngCorrect = lngCorrect + 1
                    lngWarning = lngWarning + 1
                Case FS_�Ѿ��ϴ�  '�Ѿ��ϴ�
                    lngUpload = lngUpload + 1
                    lngCorrect = lngCorrect + 1
            End Select
        Next
'        If mblnCheckMD5Flag = True Then
            lblInformation.Item(LN_�����ϴ��ļ�).Caption = Split(lblInformation.Item(LN_�����ϴ��ļ�).Caption, "��")(0) & "��" & str(.Rows - 1)
            lblInformation.Item(LN_״̬�쳣�ļ�).Caption = Split(lblInformation.Item(LN_״̬�쳣�ļ�).Caption, "��")(0) & "��" & str(lngAbnormal)
            lblInformation.Item(LN_���ϴ��ļ�).Caption = Split(lblInformation.Item(LN_���ϴ��ļ�).Caption, "��")(0) & "��" & str(lngUpload)
            lblInformation.Item(LN_���ϴ��ļ�).Caption = Split(lblInformation.Item(LN_���ϴ��ļ�).Caption, "��")(0) & "��" & str(lngCorrect)
            lblInformation.Item(LN_��龯���ļ�).Caption = Split(lblInformation.Item(LN_��龯���ļ�).Caption, "��")(0) & "��" & str(lngWarning)
            lblInformation.Item(LN_���ϴ����ļ�).Caption = Split(lblInformation.Item(LN_���ϴ����ļ�).Caption, "��")(0) & "��" & str(lngNewFile)
'            lblInformation.Item(LN_��������ļ�).Caption = Split(lblInformation.Item(LN_��������ļ�).Caption, "��")(0) & "��" & str(lngUnchanged)
'        Else
'            lblInformation.Item(Num_�ļ�����).Caption = str(.Rows - 1)
'            lblInformation.Item(Num_״̬�쳣).Caption = str(lngAbnormal)
'            lblInformation.Item(Num_�Ѿ��ϴ�).Caption = "0"
'            lblInformation.Item(Num_��Ҫ�ϴ�).Caption = "δ���"
'            lblInformation.Item(Num_MD5����).Caption = "δ���"
'            lblInformation.Item(Num_�����ϴ�).Caption = "δ���"
'        End If
    End With

    If lngNewFile = 0 Then cmdUpload.Enabled = False
End Sub

Private Sub ControlVisible(blnVisible As Boolean) '���水�������Կ���
'   cmdExit.Enabled = blnVisible
    cmdMD5Check.Enabled = blnVisible
    cmdUpload.Enabled = blnVisible
    cmdAllUpLoad.Enabled = blnVisible
    picInformation.Enabled = blnVisible
    mblnCheckMD5Tag = IIf(blnVisible = False, True, False)
    mblnUploadTag = IIf(blnVisible = False, True, False)
    mblnAllUploadTag = IIf(blnVisible = False, True, False)
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    If mblnCheckMD5Tag = True Then Cancel = 1
    If mblnUploadTag = True Then Cancel = 1
    If mblnAllUploadTag = True Then Cancel = 1
End Sub

Private Sub lblInformation_Change(Index As Integer)
    Dim lngWidth As Long
    lngWidth = picInformation.Width / 6
    lblInformation.Item(Index).Move (Index * lngWidth) + ((lngWidth) - lblInformation.Item(Index).Width) / 2, (picInformation.ScaleHeight - lblInformation.Item(Index).Height) / 2
End Sub

Private Sub lblInformation_Click(Index As Integer)
'��ԭ����״̬
    Dim i As Integer
    Dim vPic As Variant
    
    For i = 0 To LN_lbl���� - 1
        lblInformation.Item(i).FontBold = False
        lblInformation_Change i
    Next
    
    lblInformation.Item(Index).FontBold = True
    lblInformation_Change Index
'    picInformation.Picture = imgList.ListImages(Index + 1).Picture
    imgInformation.Picture = imgList.ListImages(Index + 1).Picture
    
    '�쳣�ļ���ʾ
    For Each vPic In picState
        vPic.Visible = False
    Next
    Select Case Index
    Case LN_״̬�쳣�ļ�
        imgCaption.Picture = imgList.ListImages("�쳣").Picture
        picState(0).Visible = True
    Case LN_��龯���ļ�
        imgCaption.Picture = imgList.ListImages("����").Picture
        picState(1).Visible = True
    Case Else
        imgCaption.Picture = imgList.ListImages("�ļ�").Picture
    End Select

    Call DataFilter(Index)
End Sub

'���ݹ���
Private Sub DataFilter(intFiler As Integer)
'״ֵ̬ 0-���� 1-��������ȱʧ(�����ļ��ز�����) 2-�����ļ������� 3-������� 4-���浫�����ϴ� 5-�Ѿ��ϴ�
'strFiler ����ֵ�����״̬
'"-1"-��ʾ�������ݡ�"0"-�����ϴ���"1"-״̬�쳣��"2"-״̬�쳣��"3"-������¡�"4"-�����ҿ��ϴ� ��"5"-�Ѿ��ϴ�

    Dim i As Long
    Dim strData As String
    
    With vsfMain
        If .Rows < 1 Then Exit Sub
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            strData = Trim(.Cell(flexcpData, i, Col_״̬))
            Select Case intFiler
            Case LN_�����ϴ��ļ�
                If .RowHidden(i) = True Then .RowHidden(i) = False
            Case LN_���ϴ��ļ�
                If strData <> FS_״̬�쳣 Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            Case LN_״̬�쳣�ļ�
                If strData = FS_״̬�쳣 Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            Case LN_��龯���ļ�
                If strData = FS_׼���ϴ������ļ� Or strData = FS_������¾����ļ� Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            Case LN_���ϴ��ļ�
                If strData = FS_�Ѿ��ϴ� Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            Case LN_���ϴ����ļ�
                If strData = FS_Ĭ������ Or strData = FS_׼���ϴ������ļ� Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
'            Case LN_��������ļ�
'                If strData = FS_������� Or strData = FS_������¾����ļ� Then
'                    .RowHidden(i) = False
'                Else
'                    .RowHidden(i) = True
'                End If
            End Select
        Next
        '��λ
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False Then
                .ShowCell i, Col_�ļ�
                .Row = i
                Exit For
            End If
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Function CopyFileToShareServer(ByVal strNumber As String, ByVal strServerAddress As String, Optional ByVal strUser As String, Optional ByVal strPassword As String, Optional ByRef strErrInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------
    '����:�����ļ���ָ���ķ�����
    '����:strNumber-�ļ����������
    '     strSourcePath-Դ�ļ�Ŀ¼(�ռ��ļ���Ŀ¼)
    '     strServerAddress-�������Ĺ���Ŀ¼
    '     strUser-���ʵ��û���
    '     strPassword-����
    '����:strErrInfor-���صĴ�����Ϣ
    '����:�����ɹ�,����true,���򷵻�False
    '����:����ԭ
    '����:2016/07/22
    '---------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim i As Long
    Dim strSQL As String
    Dim strTemp As String
    Dim strFilePath As String   '�����ļ���ַ
    Dim strFileSource As String '������ԭ�ļ���ַ
    Dim BlnState As Boolean
    '�������
    Dim blnUpLoadFail As Boolean '�ϴ�ʧ�ܲ������MD5
    
    '1.���������Ƿ���ͨ
    If CheckShareServer(strServerAddress, strUser, strPassword) = False Then Exit Function '�������������У��
'    MsgBox strNumber & " �ŷ������� " & """" & strServerAddress & """" & " �������ӳɹ�", vbDefaultButton1 + vbInformation, gstrSysName
                    
    If objFile.FolderExists(mstrScratchFilePath) = False Then
        strErrInfor = "Դ�ļ�Ŀ¼:" & mstrScratchFilePath & "������,����!"
        Exit Function
    End If
    
    err = 0: On Error GoTo errHand:
    
    With vsfMain
        If .Rows < 1 Then MsgBox "�ļ��б�Ϊ�գ����飡", vbDefaultButton1 + vbInformation, gstrSysName: Exit Function
        '�����ʼ��
        Call ShowStatus("", "", "", 0, .Rows - 1)
        For i = 1 To .Rows - 1
            ShowCenterRow i
            strTemp = UCase(Trim(.TextMatrix(i, Col_�ռ��ļ�)))
            strFilePath = strServerAddress & "\" & strTemp
            strFileSource = .TextMatrix(i, Col_�ռ���ַ)
            Call ShowStatus("�����ϴ�", strTemp & " �� " & strNumber & " �ŷ����� ", "", i)
            strTemp = .Cell(flexcpData, i, Col_״̬)
            If mblnAllUpload = True Then
                BlnState = (strTemp <> FS_״̬�쳣) 'ȫ���ϴ���ֻҪ��״̬�쳣�����������ϴ�
            Else
                BlnState = (strTemp = FS_Ĭ������ Or strTemp = FS_׼���ϴ������ļ� Or strTemp = FS_�Ѿ��ϴ�) '�����ϴ�״̬Ϊ�����;�����ļ������ϴ�
            End If
            If BlnState Then  '״̬Ϊ�����;�����ļ������ϴ�
                err = 0: On Error Resume Next
                .TextMatrix(i, Col_״̬) = "�����ϴ�"
'                Call FileStateSet(FS_Ĭ������, i, , "�����ϴ�")
                DoEvents '��ֹ���濨��
                objFile.CopyFile strFileSource, strFilePath, True
                If err <> 0 Then
                    If MsgBox("Դ�ļ���" & strFileSource & vbCrLf & " ���ܿ�����Ŀ���ļ���" & vbCrLf & strFilePath & vbCrLf & "��,�Ƿ������" & vbNewLine & err.Description, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    Call FileStateSet(FS_�ϴ�ʧ��, i, "�����ļ����Ժ������ϴ�")
                    blnUpLoadFail = True '�ϴ�ʧ�ܲ��ܸ���MD5
                Else
                    Call FileStateSet(FS_�Ѿ��ϴ�, i, strNumber & " �ŷ��������ϴ�")
                End If
            End If
        Next
        .Row = 1
        .ShowCell 1, Col_�ļ�
        Call ShowStatus("", "", strNumber & " �ŷ������ϴ����", pgbThis.Max)
    End With
    
    If blnUpLoadFail = False Then
        mblnUpLoadSuccess = True
        mstrSuccessUploadSever = mstrSuccessUploadSever & strNumber & "��," '����MD5
        CopyFileToShareServer = True
    Else
        CopyFileToShareServer = False
    End If
    Exit Function
errHand:
    strErrInfor = "�����ϴ����̳��ִ���:" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description
    If False Then
        Resume
    End If
End Function

Private Function CopyFileToFTPServer(ByVal strNumber As String, ByVal strServerAddress As String, Optional ByVal strUser As String, Optional ByVal strPassword As String, Optional ByVal strPort As String, Optional ByRef strErrInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------
    '����:�����ļ���ָ���ķ�����
    '����:strNumber-�ļ����������
    '     strSourcePath-Դ�ļ�Ŀ¼
    '     strServerAddress-�������Ĺ���Ŀ¼
    '     strUser-���ʵ��û���
    '     strPassword-����
    '     strPort-�˿�
    '����:strErrInfor-���صĴ�����Ϣ
    '����:�����ɹ�,����true,���򷵻�False
    '����:����ԭ
    '����:2016/07/22
    '---------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strTemp As String
    Dim strFileName As String   '�����ļ�����
    Dim strFileSource As String '�������ļ�Դ��ַ
    Dim BlnState As Boolean
    
    Dim i As Long
    Dim strSQL As String
    Dim blnUpLoadFail As Boolean
    
    If CheckFTPServer(strServerAddress, strUser, strPassword, strPort) = False Then Exit Function 'FTP����������У��
'    MsgBox strNumber & " �ŷ������� " & """" & strServerAddress & """" & " �������ӳɹ�", vbDefaultButton1 + vbInformation, gstrSysName

    err = 0: On Error GoTo errHand:
    
    With vsfMain
        If .Rows < 1 Then MsgBox "�ļ��б�Ϊ�գ����飡", vbDefaultButton1 + vbInformation, gstrSysName: Exit Function
        
        Call ShowStatus("", "", "��ʼ�ϴ�", 0, .Rows - 1)
        For i = 1 To .Rows - 1
            ShowCenterRow i
            '���ݳ�ʼ��
            strFileName = .TextMatrix(i, Col_�ռ��ļ�)
            strFileSource = .TextMatrix(i, Col_�ռ���ַ)
            Call ShowStatus("�����ϴ�", strFileName & " �� " & strNumber & " �ŷ����� ", "", i)
            strTemp = .Cell(flexcpData, i, Col_״̬)
            
            If mblnAllUpload = True Then
                BlnState = (strTemp <> FS_״̬�쳣) 'ȫ���ϴ���ֻҪ��״̬�쳣�����������ϴ�
            Else
                BlnState = (strTemp = FS_Ĭ������ Or strTemp = FS_׼���ϴ������ļ� Or strTemp = FS_�Ѿ��ϴ�) '�����ϴ�״̬Ϊ�����;�����ļ������ϴ�
            End If
            
            If BlnState Then
                err = 0: On Error Resume Next
                '�ļ�����,�����ж�

'                If UCase(Nvl(strFileName, "")) <> UCase("zlHisCrust.exe") And UCase(Nvl(strFileName, "")) <> UCase("7z.exe") And UCase(Nvl(strFileName, "")) <> UCase("7z.dll") And UCase(Nvl(strFileName, "")) <> UCase("aamd532.dll") And UCase(Nvl(strFileName, "")) <> UCase("zlRunas.exe") And UCase(Nvl(strFileName, "")) <> UCase("RegCom.dll") Then
'                    strFileName = strFileName & ".7Z"
'                End If
                .TextMatrix(i, Col_״̬) = "�����ϴ�"
                DoEvents '��ֹ���濨��
                If FtpupFile(strFileSource, strFileName) = False Then
'                    If MsgBox("Դ�ļ���" & strFileSource & vbCrLf & " ���ܿ�����Ŀ���ļ���" & vbCrLf & strFileName & vbCrLf & "��,�Ƿ������" & vbNewLine & err.Description, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    Call FileStateSet(FS_�ϴ�ʧ��, i, "�����ļ����Ժ������ϴ�")
                    blnUpLoadFail = True '�ϴ�ʧ�ܲ��ܸ���MD5
                Else
                    Call FileStateSet(FS_�Ѿ��ϴ�, i, strNumber & " �ŷ��������ϴ�")
                End If
            End If
        Next
        .Row = 1
        .ShowCell 1, Col_�ļ�
        Call ShowStatus("", "", strNumber & " �ŷ������ϴ����", pgbThis.Max)
    End With
    
    If blnUpLoadFail = False Then
        mblnUpLoadSuccess = True
        mstrSuccessUploadSever = mstrSuccessUploadSever & strNumber & "��," '����MD5
        CopyFileToFTPServer = True
    Else
        CopyFileToFTPServer = False
    End If
    Exit Function
errHand:
    strErrInfor = "FTP�ϴ����̳���:" & vbCrLf & "�����:" & err.Number & vbCrLf & "��������:" & err.Description
    If False Then
        Resume
    End If
End Function

Private Function UpdateMD5() As Boolean
    If mblnUpLoadSuccess = False Then UpdateMD5 = False: Exit Function
    Dim i As Long
    Dim lngPercent As Long
    Dim intPercent As Integer
    Dim strSQL As String
    Dim BlnState As String
    mstrSuccessUploadSever = Mid(mstrSuccessUploadSever, 1, Len(mstrSuccessUploadSever) - 1) & "�������ϴ����"
    With vsfMain
        mlngUpload = 0
        Call ShowStatus("", "", "׼������", i, .Rows - 1)
        For i = 1 To .Rows - 1
            ShowCenterRow i
            If mblnAllUpload = True Then
                BlnState = (.Cell(flexcpData, i, Col_״̬) <> FS_״̬�쳣) 'ȫ���ϴ���ֻҪ��״̬�쳣�����������ϴ�
            Else
                BlnState = (.Cell(flexcpData, i, Col_״̬) = FS_Ĭ������ Or .Cell(flexcpData, i, Col_״̬) = FS_׼���ϴ������ļ� Or .Cell(flexcpData, i, Col_״̬) = FS_�Ѿ��ϴ�) '�����ϴ�״̬Ϊ�����;�����ļ������ϴ�
            End If
            If BlnState Then
                DoEvents
                Call ShowStatus("���ڸ���", .TextMatrix(i, Col_�ļ�) & " �� MD5��", "", i)
                strSQL = "update zlfilesupgrade set MD5 = '" & .TextMatrix(i, Col_����md5) & "' where  upper(�ļ���) = '" & UCase(.TextMatrix(i, Col_�ļ�)) & "'"
                gcnOracle.Execute strSQL
                Call FileStateSet(FS_�Ѿ��ϴ�, i, mstrSuccessUploadSever, "�ϴ����")
                mlngUpload = mlngUpload + 1
            Else
                Call ShowStatus("��������", .TextMatrix(i, Col_�ļ�), "", i)
            End If
            lblInformation.Item(LN_���ϴ��ļ�).Caption = Split(lblInformation.Item(LN_���ϴ��ļ�).Caption, "��")(0) & "��" & str(mlngUpload)
        Next
        
        .Row = 1
        .ShowCell 1, Col_�ļ�
        Call ShowStatus("", "", mstrSuccessUploadSever, pgbThis.Max, , False)
        mstrSuccessUploadSever = ""
        mblnUpLoadSuccess = False
        UpdateMD5 = True
    End With
End Function

Private Function BatchLoad() As String
'��ȡ�ϴ�����
    Dim strSQL As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "select ���� as ���� from ZLReginfo where ��Ŀ = '���������ļ�����'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    If rsTemp.EOF Then
        strSQL = "insert into zltools.ZLReginfo(��Ŀ,����) select '���������ļ�����','0' from dual where not Exists (select 1 from zltools.ZLReginfo where ��Ŀ ='���������ļ�����')"
        gcnOracle.Execute strSQL
        BatchLoad = "0"
    Else
        BatchLoad = rsTemp.Fields("����")
    End If
    Exit Function
errH:
    MsgBox "���ζ�ȡ����"
End Function

Private Function BatchUpdate(strBath As String) As Boolean
'�����ϴ�����
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    
    strSQL = "update ZLReginfo set ���� = '" & Trim(strBath) & "' where ��Ŀ = '���������ļ�����'"
    gcnOracle.Execute strSQL
    
    With vsfSever
        If .Rows < 1 Then BatchUpdate = False: Exit Function
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, Col_�ϴ�״̬)) = Trim(str(Res_�ϴ��ɹ�)) Then
                strSQL = "update ZLUpgradeServer set ���� = " & strBath & " where ��� = " & Trim(.TextMatrix(i, Col_���))
                gcnOracle.Execute strSQL
            End If
        Next
    End With

    BatchUpdate = True
    
    Exit Function
errH:
    BatchUpdate = False
    MsgBox "���θ��´���"
End Function


Private Function funCanWrite(strWritePath As String) As Boolean
'�ж�Զ��Ŀ¼�Ƿ����дȨ��
    Dim strDest     As String
    Dim objFile As New FileSystemObject
    On Error GoTo errH
            strDest = strWritePath & "\tmp.txt"
            objFile.CreateTextFile strDest
            objFile.DeleteFile strDest, True
            funCanWrite = True
    Exit Function
errH:
    funCanWrite = False
End Function

'Public Function ISCopyFile(ByVal strSourceFile As String, ByVal strTarGetFile As String) As Boolean
'     '---------------------------------------------------------------------------------------------------------------
'    '
'    '����:�ж��Ƿ���Ҫ�����ļ�(�Ƚϰ汾��,�޸�ʱ��)
'    '�����:
'    '   strSourceFile:Դ�ļ�
'    '   strTargetFile:Ŀ���ļ�
'    '����:��Ҫ�����򷵻�true,���򷵻�false
'    '---------------------------------------------------------------------------------------------------------------
'    Dim strSource As String, strTarget As String
'
'    ISCopyFile = False
'    err = 0: On Error Resume Next
'    If FindFile(strTarGetFile) = False Then
'        'û�з����ļ����򷵻�true
'        ISCopyFile = True
'        Exit Function
'    End If
'
'    '�Ƚ��ļ��汾��
'    strTarget = GetCommpentVersion(strTarGetFile)
'    strSource = GetCommpentVersion(strSourceFile)
'    If RtnVerNum(strTarget) < RtnVerNum(strSource) Then
'        ISCopyFile = True
'        Exit Function
'    End If
'
'    '�Ƚ��ļ�������޸�ʱ��
'    strTarget = Format(FileDateTime(strTarGetFile), "yyyy-MM-DD hh:mm:ss")
'    strSource = Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
'    If strTarget < strSource Then
'        ISCopyFile = True
'        Exit Function
'    End If
'End Function

Private Function FilesCollections() As Boolean
    Dim i As Long
    Dim lngPercent As Long '�������ٷֱ�
    Dim intPercent As Integer '������ʾ�ٷֱ�
    Dim blnCollect As Boolean '�ռ�״̬ 0-���ռ� 1-�ռ�
    Dim BlnState As Boolean
    Dim strTemp As String
    
    'ѹ�����
    Dim strCurFileDirectory As String 'Ŀ���ļ���
    Dim strCompTxt  As String 'ѹ���ű�
    Dim strSourcePath   As String 'ѹ��Դ�ļ�·��
    Dim strDescPath     As String 'ѹ��Ŀ���ļ�·��

    Dim objFile As New FileSystemObject
    Dim str7zFile   As String
    Dim driver As Drive
        
    '���ݿ��ļ���ֵ
    Dim strFileName As String
    Dim strFilePath As String
    Dim strFileMD5 As String '�����ļ�MD5ֵ
        
    '��ⲿ���Ƿ����ռ�
    Dim strEditDate As String '�����ļ��޸�ʱ��
    Dim strEditDateNow As String '��ǰ�ļ��޸�ʱ��
    Dim strVersion As String '���ذ汾��
    Dim strVersionNow As String '��ǰ�汾��

    err = 0: On Error GoTo errHand:
    strCurFileDirectory = Trim(mstrScratchFilePath)
    FilesCollections = False
        
    '���ʣ��ռ�
    For Each driver In objFile.Drives
        If driver.IsReady Then
            If driver.DriveLetter = "C" Then
                If driver.FreeSpace < 204800000 Then 'С��200M
                    MsgBox "��ʱ�ռ�Ŀ¼û���㹻�Ŀռ�!", vbInformation, gstrSysName
                    Exit Function
                End If
                Exit For
            End If
        End If
    Next driver

    If FindFile(strCurFileDirectory) = False Then
        On Error Resume Next
        Call mobjFile.CreateFolder(strCurFileDirectory)
        If mobjFile.FolderExists(strCurFileDirectory) = False Then
            MsgBox "��ʱ�ռ�Ŀ¼���ܴ���,����!" & vbCrLf & strCurFileDirectory, vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
    End If

    '���7z·��
    If Init7Z = False Then Exit Function
       
    '�������ʱ�ռ��ļ�Ŀ¼�е���������
    err = 0: On Error Resume Next
    
'    If MsgBox("�ϴ�ǰ��Ҫ���ռ������ļ����Ƿ�Ҫ����ռ��ļ��У�ȫ���ļ������ռ���", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
'        objFile.DeleteFolder strCurFileDirectory & "\*", True
'        objFile.DeleteFile strCurFileDirectory & "\*.*", True
'    End If
    
'    MsgBox "�ռ������л��п��٣������ĵȴ���", vbDefaultButton1 + vbInformation, gstrSysName
    imgCaption.Picture = imgList.ListImages("�ռ��ļ�").Picture
    lblEXP(1).ForeColor = vbBlue
    With vsfMain
        If .Rows < 1 Then MsgBox "�ļ��б�Ϊ�գ�����", vbDefaultButton1 + vbQuestion, gstrSysName: Exit Function

        Call ShowStatus("", "", "׼���ռ�", 0, .Rows - 1, True)
        For i = 1 To .Rows - 1
            ShowCenterRow i
            DoEvents
            '���½���״̬��������������ʾ��
            Call ShowStatus("�����ռ�", strFileName & " �� " & strCurFileDirectory, "", i)
            '��ʼ����ǰ�����ݡ����ơ���ַ�����桢�޸�����
            strFileName = UCase(.TextMatrix(i, Col_�ļ�))
            strFilePath = .TextMatrix(i, Col_�ļ���ַ)
            strFileMD5 = .TextMatrix(i, Col_����md5)
            strVersion = GetDealVersion(strFilePath)  '�����ļ��汾
            strVersionNow = .TextMatrix(i, Col_��ǰ�汾)  '��ǰ�ļ��汾
            strEditDate = Format(FileDateTime(strFilePath), "yyyy-MM-DD hh:mm:ss") '�����ļ��޸�ʱ��
            strEditDateNow = Format(.TextMatrix(i, Col_�޸�����), "yyyy-MM-DD hh:mm:ss") '��ǰ�ļ��޸�ʱ��
            blnCollect = False
            If mblnAllUpload = True Then
                BlnState = (.Cell(flexcpData, i, Col_״̬) <> FS_״̬�쳣)
            Else
                BlnState = (.Cell(flexcpData, i, Col_״̬) = FS_Ĭ������ Or .Cell(flexcpData, i, Col_״̬) = FS_׼���ϴ������ļ�)
            End If
            If BlnState Then
                '״̬��Ϊ�쳣�Ŀ����ռ�
                '7z����ѹ����5���ļ�����Ҫѹ�� ���⴦��
                If InStr(";ZLHISCRUST.EXE;7Z.EXE;7Z.DLL;AAMD532.DLL;ZLRUNAS.EXE;REGCOM.DLL;GACUTIL.EXE;GACUTIL.EXE.CONFIG;", ";" & UCase(Nvl(strFileName, "")) & ";") > 0 Then
                    strDescPath = strCurFileDirectory & "\" & UCase(strFileName)
                    '�ļ����� �Ұ汾����ͬ ���޸�������ͬ ����Ҫ����ѹ��
                    If objFile.FileExists(strDescPath) = True And strVersion = strVersionNow And CompareDate(strEditDate, strEditDateNow) = 0 Then
                        blnCollect = False
                    Else
                        blnCollect = True
                    End If
                    
                    If blnCollect = True Then
                        DoEvents '��ֹ���濨��
                        Call objFile.CopyFile(strFilePath, strDescPath, True)
                        Call SaveCollectFilesInformation(i)

                        Call FileStateSet(FS_Ĭ������, i, "�ռ����")
                    Else
                        Call FileStateSet(FS_Ĭ������, i, "�����ռ�")
                    End If
                    '�洢ѹ���ļ���ַ��ѹ�����ļ�����
                    .TextMatrix(i, Col_�ռ���ַ) = strDescPath
                    .TextMatrix(i, Col_�ռ��ļ�) = strFileName
                Else
                    strDescPath = strCurFileDirectory & "\" & GetCompressName(Nvl(strFileName, ""))
                    If objFile.FileExists(strDescPath) = True And strVersion = strVersionNow And CompareDate(strEditDate, strEditDateNow) = 0 Then
                        blnCollect = False
                    Else
                        blnCollect = True
                    End If
                    
                    If blnCollect = True Then
                        strCompTxt = CompressionCmd(strDescPath, strFilePath, COMPRESSIONRATE)
                        If strCompTxt <> "" Then
                            DoEvents '��ֹ���濨��
                            Call GetCmdTxt(strCompTxt)
                            Call SaveCollectFilesInformation(i)
                            
                            Call FileStateSet(FS_Ĭ������, i, "�ռ����")
                        Else
                            Call FileStateSet(FS_״̬�쳣, i, "�ļ��ռ�ʧ��,������")
                        End If
                    Else
                        .TextMatrix(i, Col_����) = "�����ռ�"
                    End If
                    '�洢ѹ���ļ���ַ��ѹ�����ļ�����
                    .TextMatrix(i, Col_�ռ���ַ) = strDescPath
                    .TextMatrix(i, Col_�ռ��ļ�) = GetCompressName(Nvl(strFileName, ""))
                End If
            End If
        Next
        Call ShowStatus("", "", "�ļ��ռ����", pgbThis.Max)
        .Row = 1
        .ShowCell 1, Col_�ļ�
    End With
    imgCaption.Picture = imgList.ListImages("�ļ�").Picture
    lblEXP(1).ForeColor = vbBlack
    FilesCollections = True
    Exit Function
errHand:
    MsgBox "ѹ���ռ����̳��ִ���:" & vbCrLf & err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Function

Private Function GetCompressName(ByVal strFileName As String) As String
'����ת��Ϊ7z��׺��ѹ����ʽ����
    On Error GoTo errH
    GetCompressName = strFileName & ".7z"
    Exit Function
errH:
    If err Then
         MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function SaveCollectFilesInformation(lngRow As Long) As Boolean
'�����ϴ������ļ���Ϣ
'lngRow - �����������ÿһ�д���һ������
    Dim strFilesPath As String '�ļ�·��
    Dim strFileName As String '�ļ���
    Dim strMD5      As String 'MD5
    Dim strEditDate As String '�޸�����
    Dim strVision   As String   '�汾��
    Dim strSQL As String
    On Error GoTo errH
    
    With vsfMain
        strFilesPath = .TextMatrix(lngRow, Col_�ļ���ַ)
        strFileName = .TextMatrix(lngRow, Col_�ļ�)
'        strMD5 = .TextMatrix(lngRow, Col_����md5)
        strEditDate = Format(FileDateTime(strFilesPath), "yyyy-MM-DD hh:mm:ss")
        strVision = GetDealVersion(strFilesPath)
'        strVision = GetCommpentVersion(strFilesPath)
'        strVision = GetTransVersion(strVision)
        
        If strFileName <> "" Then
            If InStr(";ZLHISCRUST.EXE;7Z.EXE;7Z.DLL;AAMD532.DLL;ZLRUNAS.EXE;REGCOM.DLL;GACUTIL.EXE;GACUTIL.EXE.CONFIG;", ";" & UCase(Nvl(strFileName, "")) & ";") > 0 Then
                strSQL = "update zlfilesupgrade set �汾��='1000350040' ,�ļ��汾��='" & strVision & "',�޸�����='" & strEditDate & "' where upper(�ļ���)='" & UCase(strFileName) & "'"
            Else
                strSQL = "update zlfilesupgrade set �ļ��汾��='" & strVision & "',�޸�����='" & strEditDate & "' where upper(�ļ���)='" & UCase(strFileName) & "'"
            End If
            gcnOracle.Execute strSQL
            SaveCollectFilesInformation = True
        End If
    End With
    Exit Function
errH:
    If err Then
         MsgBox err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function GetTransVersion(ByVal strVersion As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:���ת����İ汾��
    '���:strVersion
    '����:strVersion ת������"."�İ汾��
    '����:�ɹ�,���ذ汾��,���򷵻ؿ�
    '����:����ԭ
    '����:2016-08-03 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim lngVision   As Double '�汾��
    Dim strTmp    As Variant
    
        If strVersion <> "" Then
            strTmp = Split(strVersion, ".")
            lngVision = strTmp(0) * 10 ^ 8 + strTmp(1) * 10 ^ 4 + strTmp(2)
            strVersion = lngVision
        End If
    GetTransVersion = strVersion
End Function

Private Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ؼ��İ汾��
    '���:
    '����:
    '����:�ɹ�,���ذ汾��,���򷵻ؿ�
    '����:���˺�
    '����:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    err = 0: On Error Resume Next
    '��ȡ�ļ��汾��
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
'    If Trim(strVer) <> "" Then
'        varVersion = Split(strVer, ".")
'        If UBound(varVersion) > 2 Then
'            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
'        ElseIf UBound(varVersion) = 2 Then
'            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
'        End If
'    End If
    GetCommpentVersion = strVer
End Function

Private Function CompareDate(Data1 As String, Data2 As String) As Integer
'���ڱȽ� ��ʽ����Ϊ "yyyy-MM-DD hh:mm:ss"
'Data1>Data2 ���� 1
'Data1<Data2 ���� -1
'Data1=Data2 ���� 0
'���� ���� 3
    If Data1 = "" Or Data2 = "" Then CompareDate = 3: Exit Function
    
    If DateDiff("s", Data1, Data2) < 0 Then
        CompareDate = 1 'data1��
    ElseIf DateDiff("s", Data1, Data2) > 0 Then
        CompareDate = 2 'data2��
    Else
        If Trim(Data1) = Trim(Data2) Then
            CompareDate = 0 '���
        Else
            CompareDate = 3 '����
        End If
    End If
End Function

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Me.Refresh
    Select Case Item.Index
        Case 0
        Case 1
        Case 2
        Case 3
        Case 4
        Case 5
        Case 6
    End Select
End Sub

Private Sub ShowStatus(strOperation As String, strContent As String, strCondition As String, lngPgbValue As Long, Optional lngPgbMax As Long = 0, Optional blnShowPgbbar As Boolean = True)
    '�״������˽�����max��ֻ�ô��������Value����
    'strCondition��Ϊ�գ���������ʾstrOperation��strContent����
    '���������ValueΪ-1ʱ�����ȰٷֱȲ�����ʾ
    'blnShowPgbbar�������غ���ʾ������ true-���� false-������
    On Error Resume Next
    Dim intPercent As Integer
    If lngPgbMax < 0 Then Exit Sub

    If lngPgbValue = -1 Then
        lblstatus(Sta_�ٷֱ�).Visible = False
    Else
        lblstatus(Sta_�ٷֱ�).Visible = True
    End If
    
    If lngPgbMax = 0 Then
        If pgbThis.Max <> 0 Then
            intPercent = lngPgbValue / pgbThis.Max * 100
        Else
            Exit Sub
        End If
    Else
        pgbThis.Max = lngPgbMax
        intPercent = lngPgbValue / pgbThis.Max * 100
    End If
    
    If strCondition <> "" Then
        lblstatus(Sta_״̬����).Caption = strCondition
        lblstatus(Sta_�ٷֱ�).Caption = ""
        lblstatus(Sta_��ǰ����).Caption = ""
        lblstatus(Sta_��������).Caption = ""
    Else
        lblstatus(Sta_״̬����).Caption = ""
        lblstatus(Sta_�ٷֱ�).Caption = intPercent & "%"
        lblstatus(Sta_��ǰ����).Caption = strOperation & ""
        lblstatus(Sta_��������).Caption = strContent & ""
        pgbThis.value = lngPgbValue
    End If
    
    If blnShowPgbbar Then
        pgbThis.Visible = True
        vsfMain.Move 100, 1000 + picHelp.Height, Me.ScaleWidth - 185, Me.ScaleHeight - 1490 - picHelp.Height
    Else
        pgbThis.Visible = False
        vsfMain.Move 100, pgbThis.Top, Me.ScaleWidth - 185, Me.ScaleHeight - 1800 - pgbThis.Height
    End If
End Sub


Private Sub FileStateSet(emuFileState As FilesState, lngRow As Long, Optional strStateContent As String = "NULL", Optional strStateDisplay As String = "", Optional lngStateColor As Long = -1)
'    Ĭ��״̬��ɫ����
    Select Case emuFileState
        Case FS_Ĭ������
            If lngStateColor = -1 Then lngStateColor = SC_��ɫ
            If strStateDisplay = "" Then strStateDisplay = SDP_׼���ϴ�
        Case FS_״̬�쳣
            If lngStateColor = -1 Then lngStateColor = SC_��ɫ
            If strStateDisplay = "" Then strStateDisplay = SDP_״̬�쳣
        Case FS_�������
            If lngStateColor = -1 Then lngStateColor = SC_��ɫ
            If strStateDisplay = "" Then strStateDisplay = SDP_�������
        Case FS_������¾����ļ�
            If lngStateColor = -1 Then lngStateColor = SC_��ɫ
            If strStateDisplay = "" Then strStateDisplay = SDP_�������
        Case FS_׼���ϴ������ļ�
            If lngStateColor = -1 Then lngStateColor = SC_��ɫ
            If strStateDisplay = "" Then strStateDisplay = SDP_׼���ϴ�
        Case FS_�Ѿ��ϴ�
            If lngStateColor = -1 Then lngStateColor = SC_��ɫ
            If strStateDisplay = "" Then strStateDisplay = SDP_�Ѿ��ϴ�
        Case FS_�ϴ�ʧ��
            If lngStateColor = -1 Then lngStateColor = SC_��ɫ
            If strStateDisplay = "" Then strStateDisplay = SDP_�ϴ�ʧ��
        Case Else
            If lngStateColor = -1 Then lngStateColor = vbBlack
            If strStateDisplay = "" Then strStateDisplay = "TEST"
    End Select
    
    With vsfMain
        .Cell(flexcpData, lngRow, Col_״̬) = emuFileState
        .TextMatrix(lngRow, Col_״̬) = strStateDisplay
        If strStateContent <> "NULL" Then
            .TextMatrix(lngRow, Col_����) = strStateContent
        End If
        If FS_״̬�쳣 = emuFileState Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, Col_���� - 1) = lngStateColor
        Else
            .Cell(flexcpForeColor, lngRow, Col_״̬, lngRow, Col_����) = lngStateColor
        End If
    End With
End Sub

Private Sub ShowCenterRow(lngRow As Long)

    Dim i As Integer
    Dim intLocation As Integer
    Dim lngShowRow As Integer
    With vsfMain
        If .RowHidden(lngRow) = True Then Exit Sub
        intLocation = 14
        lngShowRow = lngRow
        i = 0
        Do Until i >= intLocation
            If .Rows - (lngShowRow + intLocation) <= 0 Then
                .Row = lngRow
                .ShowCell lngRow, Col_�ļ�
                Exit Sub
            End If
            lngShowRow = lngShowRow + 1
            If .RowHidden(lngShowRow) = False Then i = i + 1
        Loop
        .Row = lngRow
        .ShowCell lngShowRow, Col_�ļ�
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub AllUploadRestore()
    Dim i As Long
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, Col_״̬) <> FS_״̬�쳣 Then
                Select Case .Cell(flexcpData, i, Col_״̬)
                Case FS_Ĭ������
                    Call FileStateSet(FS_Ĭ������, i, "")
                Case FS_�������
                    Call FileStateSet(FS_Ĭ������, i, "")
                Case FS_������¾����ļ�
                    Call FileStateSet(FS_׼���ϴ������ļ�, i)
                Case FS_׼���ϴ������ļ�
                    Call FileStateSet(FS_׼���ϴ������ļ�, i)
                Case Else
                    Call FileStateSet(FS_Ĭ������, i, "")
                End Select
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    Call UpLoadFilesCount
End Sub
'Private Sub FloderToClipBoard(ByVal strSourceFloder As String)
'    '������ʱ�ռ��ļ�Ŀ¼���ļ�����������ȥ
'    Dim strFile() As String
'    Dim strSourceFile As String
'    Dim strTemp As String
'    Dim i As Integer
'    strSourceFile = strSourceFloder & "\"
'    Erase strFile
'
'
'    If mobjFile.FolderExists(strSourceFile) Then
'        With FileList
'            .Refresh
'            .Path = strSourceFile
'            .FileName = "*.*"
'
'            For i = 0 To .ListCount - 1
'                ReDim Preserve strFile(i)
'                strTemp = strSourceFile & .List(i)
'                strFile(i) = strTemp
'            Next
'
'            If .ListCount <> 0 Then
'                Call clipCopyFiles(strFile)
'            End If
'        End With
'    End If
'End Sub
