VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmInAdviceSend 
   AutoRedraw      =   -1  'True
   Caption         =   "סԺ��������"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "frmInAdviceSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9615
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6615
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   3
      Top             =   6210
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6150
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInAdviceSend.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceSend.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceSend.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   900
      BandCount       =   2
      FixedOrder      =   -1  'True
      BandBorders     =   0   'False
      _CBWidth        =   9615
      _CBHeight       =   510
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinWidth1       =   2895
      MinHeight1      =   450
      Width1          =   2895
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "tbrSys"
      MinHeight2      =   450
      Width2          =   9195
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   450
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   794
         ButtonWidth     =   2514
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���͵�סԺ"
               Key             =   "���͵�סԺ"
               Description     =   "���͵�סԺ"
               Object.ToolTipText     =   "���͵�סԺ(Ctrl+1)"
               Object.Tag             =   "���͵�סԺ"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���͵�����"
               Key             =   "���͵�����"
               Description     =   "���͵�����"
               Object.ToolTipText     =   "���͵�����(Ctrl+2)"
               Object.Tag             =   "���͵�����"
               ImageKey        =   "����"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrSys 
         Height          =   450
         Left            =   3120
         TabIndex        =   6
         Top             =   30
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   794
         ButtonWidth     =   1561
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ѡ��"
               Key             =   "ѡ��"
               Description     =   "ѡ��"
               Object.ToolTipText     =   "��������ѡ��(F12)"
               Object.Tag             =   "ѡ��"
               ImageKey        =   "ѡ��"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫѡ"
               Key             =   "ȫѡ"
               Description     =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Ctrl+A)"
               Object.Tag             =   "ȫѡ"
               ImageKey        =   "ȫѡ"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Ctrl+R)"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "ȫ��"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����(F1)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�(ALT+X)"
               Object.Tag             =   "�˳�"
               ImageKey        =   "�˳�"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      Top             =   4605
      Width           =   9495
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   8
      Top             =   525
      Width           =   9435
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   60
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   825
      Width           =   9540
      _cx             =   16828
      _cy             =   6641
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
      BackColorSel    =   16764057
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInAdviceSend.frx":1A92
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList img16 
         Left            =   3435
         Top             =   1905
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceSend.frx":1B2D
               Key             =   "T"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceSend.frx":20C7
               Key             =   "F"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInAdviceSend.frx":2661
               Key             =   "ǩ��"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1470
      Left            =   0
      TabIndex        =   1
      Top             =   4665
      Width           =   9525
      _cx             =   16801
      _cy             =   2593
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin MSComctlLib.ImageList imgColor 
      Left            =   360
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":29B3
            Key             =   "ȫѡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":2BCD
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":2DE7
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3001
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":321B
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3435
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   960
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":364F
            Key             =   "ȫѡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3869
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3A83
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3C9D
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":3EB7
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInAdviceSend.frx":40D1
            Key             =   "ѡ��"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInAdviceSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String 'IN
Private mlng����ID As Long 'IN
Private mlng��ҳID As Long 'IN
Private mlngǰ��ID As Long 'IN
Private mblnSend As Boolean 'OUT:�Ƿ�ɹ����͹���
Private mblnRefresh As Boolean 'OUT'�Ƿ���Ҫˢ��������

Private mcolStock As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mrsPati As ADODB.Recordset '����������Ϣ
Private mrsPrice As ADODB.Recordset '�����Ƽ۹�ϵ
Private mrsBill As ADODB.Recordset
Private mstrLike As String
Private mblnFirst As Boolean
Private mint���� As Integer

'----------------------------------------------
Private Const COL_ѡ�� = 0
Private Const COL_Ӥ�� = 1
Private Const COL_ҽ������ = 2
Private Const COL_���� = 3
Private Const COL_������λ = 4
Private Const COL_���� = 5
Private Const COL_������λ = 6
Private Const COL_��� = 7
Private Const COL_Ƶ�� = 8
Private Const COL_�÷� = 9
Private Const COL_ҽ������ = 10
Private Const COL_ִ��ʱ�� = 11
Private Const COL_ִ�п��� = 12
Private Const COL_ִ������ = 13
Private Const COL_ID = 14 '������
Private Const COL_���ID = 15
Private Const COL_ҽ��״̬ = 16
Private Const COL_���˿���ID = 17
Private Const COL_��������ID = 18
Private Const COL_����ҽ�� = 19
Private Const COL_����ʱ�� = 20
Private Const COL_������� = 21
Private Const COL_������ĿID = 22
Private Const COL_�Ƽ����� = 23
Private Const COL_ִ������ID = 24
Private Const COL_ִ�п���ID = 25
Private Const COL_�������� = 26
Private Const COL_ҩƷID = 27
Private Const COL_����ϵ�� = 28
Private Const COL_סԺ��װ = 29
Private Const COL_סԺ��λ = 30
Private Const COL_�ɷ���� = 31
Private Const COL_��� = 32
Private Const COL_���� = 33
Private Const COL_�ֽ�ʱ�� = 34
Private Const COL_�״�ʱ�� = 35
Private Const COL_ĩ��ʱ�� = 36
Private Const COL_ǩ��ID = 37

'-------------------------------------------------
Private Const COLP_�к� = 0
Private Const COLP_�շ�ϸĿID = 1
Private Const COLP_�̶� = 2
Private Const COLP_��� = 3
Private Const COLP_�Ƽ�ҽ�� = 4 '�ɼ���
Private Const COLP_��� = 5
Private Const COLP_�շ���Ŀ = 6
Private Const COLP_�Ƽ����� = 7
Private Const COLP_���� = 8
Private Const COLP_��λ = 9
Private Const COLP_���� = 10
Private Const COLP_Ӧ�ս�� = 11
Private Const COLP_ʵ�ս�� = 12
Private Const COLP_ִ�п��� = 13
Private Const COLP_�������� = 14
Private Const COLP_���� = 15
Private Const COLP_�շ���� = 16
Private Const COLP_ִ�п���ID = 17
Private Const COLP_�������� = 18

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.Value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.Value = vNewValue
        txtPer.Text = CInt(psb.Value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal lngǰ��ID As Long, blnRefresh As Boolean) As Boolean
'���ܣ�����ҽ��
'������blnRefresh=�Ƿ�ˢ������������
    mstrPrivs = strPrivs
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlngǰ��ID = lngǰ��ID
    
    On Error Resume Next
    Me.Show 1, frmParent
    Err.Clear: On Error GoTo 0
    blnRefresh = mblnRefresh
    ShowMe = mblnSend
End Function

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Dim str���s As String
    
    If mblnFirst Then
        mblnFirst = False
        
        '��ȡ�����嵥
        Me.Refresh
        str���s = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "סԺ�����������", "")
        If Not LoadAdviceSend(str���s) Then Unload Me: Exit Sub
    End If
End Sub

Private Function GetPatiInfo() As Boolean
'���ܣ���ȡ������Ϣ
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1]" & _
        " Union ALL" & _
        " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
    strSQL = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSQL & ") Group by ����ID"
    
    strSQL = "Select A.סԺ��,A.����,A.�Ա�,A.����,B.��Ժ���� as ����," & _
        " B.��ǰ����ID,B.��Ժ����ID,B.�ѱ�,B.ҽ�Ƹ��ʽ,B.����,C.ʣ���," & _
        " B.״̬,D.���� as ������,Decode(D.����,'1',1,Decode(Nvl(B.����,0),0,0,1)) as ҽ��," & _
        " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
        " From ������Ϣ A,������ҳ B,(" & strSQL & ") C,ҽ�Ƹ��ʽ D" & _
        " Where A.����ID=B.����ID And A.����ID=C.����ID(+)" & _
        " And B.ҽ�Ƹ��ʽ=D.����(+) And B.����ID=[1] And B.��ҳID=[2]"
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    lblPati.Caption = _
        "סԺ��:" & Nvl(mrsPati!סԺ��) & "������:" & mrsPati!���� & "���Ա�:" & Nvl(mrsPati!�Ա�) & "������:" & Nvl(mrsPati!����) & _
        "������:" & Nvl(mrsPati!����) & "���ѱ�:" & Nvl(mrsPati!�ѱ�) & "��ҽ�Ƹ��ʽ:" & Nvl(mrsPati!ҽ�Ƹ��ʽ) & _
        "��ʣ���:" & Format(Nvl(mrsPati!ʣ���, 0), "0.00")
    
    '���ղ����ú�ɫ��ʾ
    If Not IsNull(mrsPati!����) Then lblPati.ForeColor = vbRed
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("����"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("�˳�"))
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("ȫѡ"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbrSys_ButtonClick(tbrSys.Buttons("ȫ��"))
    ElseIf KeyCode = vbKey1 And Shift = vbCtrlMask Then
        If tbrMain.Buttons("����Ϊ�շѵ�").Visible Then
            Call tbrMain_ButtonClick(tbrMain.Buttons("����Ϊ�շѵ�"))
        End If
    ElseIf KeyCode = vbKey2 And Shift = vbCtrlMask Then
        If tbrMain.Buttons("����Ϊ���ʵ�").Visible Then
            Call tbrMain_ButtonClick(tbrMain.Buttons("����Ϊ���ʵ�"))
        End If
    ElseIf KeyCode = vbKeyF7 Then '�л����뷨
        If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
            If stbThis.Panels("WB").Bevel = sbrRaised Then
                Call stbThis_PanelClick(stbThis.Panels("WB"))
            Else
                Call stbThis_PanelClick(stbThis.Panels("PY"))
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
        
    If InStr(mstrPrivs, "�����������") = 0 Then
        tbrMain.Buttons("���͵�����").Visible = False
        tbrMain.Buttons("���͵�סԺ").Caption = "����"
        tbrMain.Buttons("���͵�סԺ").Tag = "����"
        tbrMain.Buttons("���͵�סԺ").ToolTipText = "����(Ctrl+1)"
        cbr.Bands(1).MinWidth = cbr.Bands(1).MinWidth / 3
        cbr.Bands(1).Width = cbr.Bands(1).MinWidth
    End If
    Call InitAdviceTable
    Call InitPriceTable
    Call RestoreWinState(Me, App.ProductName)
    
    mstrLike = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    mint���� = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0)) '����ƥ�䷽ʽ��0-ƴ��,1-���
    Select Case mint����
        Case 0
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrRaised
        Case 1
            stbThis.Panels("PY").Bevel = sbrRaised
            stbThis.Panels("WB").Bevel = sbrInset
        Case Else
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrInset
    End Select
   
    mblnSend = False
    mblnRefresh = False
    mblnFirst = True
    
    '�����ⷿҩƷ�����鷽ʽ
    Set mcolStock = InitStockCheck(2, True)
    
    '��ʾ������Ϣ
    If Not GetPatiInfo Then Unload Me: Exit Sub
End Sub

Private Function GetStockCheck(ByVal lng�ⷿID As Long) As Integer
'���ܣ���ȡָ���ⷿ�ĳ������鷽ʽ
    Dim intStyle As Integer
    On Error Resume Next
    intStyle = mcolStock("_" & lng�ⷿID)
    Err.Clear: On Error GoTo 0
    GetStockCheck = intStyle
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    fraInfo.Top = cbr.Height
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraInfo.Top + fraInfo.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraInfo.Height - vsPrice.Height - fraUD.Height - cbr.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = 0
    fraUD.Width = Me.ScaleWidth
    
    vsPrice.Left = 0
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = Me.ScaleWidth
    
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '�ͷ�˽�м�IN����
    mstrPrivs = ""
    mlng��ҳID = 0
    mlng����ID = 0
    Set mrsPati = Nothing
    Set mrsPrice = Nothing
    Set mrsBill = Nothing
    Set mcolStock = Nothing
    
    gbln�Ӱ�Ӽ� = False
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsPrice.Height - y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + y
        vsAdvice.Height = vsAdvice.Height + y
        vsPrice.Top = vsPrice.Top + y
        vsPrice.Height = vsPrice.Height - y
        Me.Refresh
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", _
            IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
        mint���� = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0)) '����ƥ�䷽ʽ��0-ƴ��,1-���
    End If
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lng���ͺ� As Long, strMsg As String, i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                Exit For
            End If
        Next
        If i > .Rows - 1 Then
            MsgBox "��ǰû�п��Է��͵�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    If Button.Key = "���͵�סԺ" Then
        strMsg = "����ҽ�����͵ķ��ý�����ΪסԺ���ʵ��ݣ�ȷʵҪ������ѡ���ҽ����"
    ElseIf Button.Key = "���͵�����" Then
        strMsg = "����ҽ�����͵ķ��ý�����Ϊ�����շѵ��ݣ�ȷʵҪ������ѡ���ҽ����"
    End If
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    
    lng���ͺ� = SendAdvice(IIF(Button.Key = "���͵�����", True, False))
    If lng���ͺ� <> 0 Then
        mblnSend = True
        '��ӡ���Ƶ���
        Call frmSendBillPrint.ShowMe(lng���ͺ�, 2, Me, mlngǰ��ID)
        
        '���ȫ���������,���˳�
        If vsAdvice.Rows = 2 Then
            If Val(vsAdvice.TextMatrix(1, COL_ID)) = 0 Then
                Unload Me: Exit Sub
            End If
        End If
        Call GetPatiInfo
    End If
End Sub

Private Sub tbrSys_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Long
    
    Select Case Button.Key
        Case "ȫѡ"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = img16.ListImages("T").Picture
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "ȫ��"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "ѡ��"
            With frmInAdviceSendCond
                .Show 1, Me
                If .mblnOK Then
                    Call LoadAdviceSend(.mstr���s)
                End If
            End With
        Case "����"
            ShowHelp App.ProductName, Me.Hwnd, Me.Name
        Case "�˳�"
            Unload Me
    End Select
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, _
    Optional rsUpload As ADODB.Recordset, Optional strҽ��IDs As String)
'���ܣ����ݿɼ��е�ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsAdvice
        If lngCol = COL_ѡ�� Then
            For i = lngRow + 1 To .Rows - 1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            'ȡ��ѡ��ʱ
            If Not (.Cell(flexcpData, lngRow, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, lngRow, COL_ѡ��) Is Nothing) Then
                i = IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_���ID)))
                '1.�����Ӧ�ķ��ü����ͼ�¼��д
                If Not rsSQL Is Nothing Then
                    rsSQL.Filter = "ҽ��ID=" & i
                    Do While Not rsSQL.EOF
                        rsSQL.Delete
                        rsSQL.Update
                        rsSQL.MoveNext
                    Loop
                    rsSQL.Filter = 0 '��ΪҪʹ��BookMark����˻ָ�
                End If
                '2.�����Ӧ�ķ��ͼƼ������ۼ�
                If Not rsTotal Is Nothing Then
                    rsTotal.Filter = "ҽ��ID=" & i
                    Do While Not rsTotal.EOF
                        rsTotal.Delete
                        rsTotal.Update
                        rsTotal.MoveNext
                    Loop
                End If
                '3.�����Ӧ��ҽ���ϴ����ݺ�
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "ҽ��ID=" & i
                    Do While Not rsUpload.EOF
                        rsUpload.Delete
                        rsUpload.Update
                        rsUpload.MoveNext
                    Loop
                End If
                '4.��������͵�ǩ��ҽ����ID
                If strҽ��IDs <> "" Then
                    strҽ��IDs = strҽ��IDs & ","
                    strҽ��IDs = Replace(strҽ��IDs, "," & i & ",", ",")
                    If strҽ��IDs <> "" Then
                        strҽ��IDs = Left(strҽ��IDs, Len(strҽ��IDs) - 1)
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function GetVisibleRow(ByVal lngRow As Long, Optional ByVal blnFirst As Boolean) As Long
'���ܣ�����ָ��ҽ���У����ظ�ҽ���пɼ�����
    Dim lng��ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
        'һ����ҩ�Ķ�λ����һҩƷ��
        If blnFirst Then
            If .TextMatrix(lngRow, COL_�������) = "E" And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 _
                And Val(.TextMatrix(lngRow, COL_���ID)) = 0 And Val(.TextMatrix(lngRow, COL_ID)) = Val(.TextMatrix(lngRow - 1, COL_���ID)) Then
                i = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_���ID)
                If i <> -1 Then GetVisibleRow = i: Exit Function
            End If
        End If
        
        lng��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
    End With
End Function

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        If OldRow <> NewRow And .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                Call ShowAdvicePrice(NewRow)
                
                'ȱʡѡ��Ƽ�ҽ��(�������)
                Call ShowDefaultRow
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_ѡ�� + 1 - .FixedCols Then
            .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = COL_ҽ������ Then
            .AutoSize COL_ҽ������
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_ѡ�� Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_ѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_Ƶ��: lngRight = COL_�÷�
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: Exit For
                End If
            Next
            If i > .Rows - 1 Then .Row = .FixedRows
            Call .ShowCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_ѡ�� Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_ѡ��) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_ѡ��) Is Nothing Then
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = img16.ListImages("T").Picture
                Else
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        End If
    End With
End Sub

Private Sub ShowDefaultRow()
'���ܣ����ڿ��ԼƼ۵�ҽ��,ȱʡ����һ�в�����ȱʡ�Ƽ�ҽ��
'˵����ComboList="#ҽ��ID1;�Ƽ�ҽ��1|#ҽ��ID2;�Ƽ�ҽ��2|..."
'      ���ڵ�һ����ʾ�Ƽ۱�ͻس�������ʱ����
    Dim arrCombo As Variant, lngRow As Long, i As Long
    Dim lngҽ��ID As Long, lng�к� As Long, str�Ƽ�ҽ�� As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    With vsPrice
        If .ColData(COLP_�Ƽ�ҽ��) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_�Ƽ�ҽ��), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_�к�)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_�շ�ϸĿID)) <> 0 Then
                '��һ����ʾʱȱʡ����һ��
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '���ǵ�һ����ʾʱȱʡ�Ƽ�ҽ������һ����ͬ
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_�̶�)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_�к�)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lngҽ��ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str�Ƽ�ҽ�� = Replace(arrCombo(i), "#" & lngҽ��ID & ";", "")
                lng�к� = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
                If blnHave Then
                    If lng�к� = Val(.TextMatrix(lngRow - 1, COLP_�к�)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            'ģ��ѡ������Ƽ�ҽ��
            .TextMatrix(lngRow, COLP_�к�) = lng�к�
            .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = str�Ƽ�ҽ��
            .Cell(flexcpData, lngRow, COLP_�Ƽ�ҽ��) = .TextMatrix(lngRow, COLP_�Ƽ�ҽ��)
            
            'ֻ��һ���Ƽ�ҽ��ʱ����ͣ��
            If UBound(arrCombo) = 0 Then
                .Col = COLP_�շ���Ŀ
            Else
                .Col = COLP_�Ƽ�ҽ��
            End If
        End If
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngԭ��ID As Long, lngҽ��ID As Long
    Dim lng�շ�ϸĿID As Long, i As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_�Ƽ�ҽ�� Then
            '�������ComboData,TextMatrixȡֵ��ΪComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lngҽ��ID = .ComboData
                lngԭ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
                lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                                
                '���üƼ�ҽ���Ƿ�������ͬ�շ�ϸĿ
                If lng�շ�ϸĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """�Ѿ��������շ���Ŀ""" & .TextMatrix(Row, COLP_�շ���Ŀ) & """��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                'ԭ����ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                If lngԭ��ID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '���������˵ļƼ�ҽ������
                i = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
                .TextMatrix(Row, COLP_�к�) = i
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                If lng�շ�ϸĿID <> 0 Then
                    '��ѡ���ҽ���Ƿ��д�������޸ĺ����Ŀ�Ƿ����
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ����=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_����) = IIF(blnHaveSub, "��", "")
                
                    '���»����Ӽ�¼������
                    If lngԭ��ID = 0 Then
                        mrsPrice.AddNew '����
                    Else '����
                        mrsPrice.Filter = "ҽ��ID=" & lngԭ��ID & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                    End If
                    mrsPrice!ҽ��ID = lngҽ��ID
                    If Val(vsAdvice.TextMatrix(i, COL_���ID)) <> 0 Then
                        mrsPrice!���ID = vsAdvice.TextMatrix(i, COL_���ID)
                    Else
                        mrsPrice!���ID = Null
                    End If
                    If lngԭ��ID = 0 Then
                        mrsPrice!�շ�ϸĿID = lng�շ�ϸĿID
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_�Ƽ�����))
                        mrsPrice!���� = Val(.TextMatrix(Row, COLP_����))
                        mrsPrice!�̶� = 0
                    End If
                    mrsPrice!���� = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
            End If
        ElseIf Col = COLP_�շ���Ŀ Or Col = COLP_ִ�п��� Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        ElseIf Col = COLP_�Ƽ����� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        ElseIf Col = COLP_���� Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, COLP_Ӧ�ս��), .Cell(flexcpData, Row, COLP_ʵ�ս��), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), "0.00000")
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '���¼�¼��
            lngҽ��ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_�к�)), COL_ID))
            lng�շ�ϸĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
            If lngҽ��ID <> 0 And lng�շ�ϸĿID <> 0 Then
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lng�շ�ϸĿID
                mrsPrice!���� = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    '���ݿɷ�༭����
    If Not CellEditable(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_�Ƽ�ҽ�� Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_�շ���Ŀ Or NewCol = COLP_ִ�п��� Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
        
    '��ʾҩƷ���
    If NewRow <> OldRow Then
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_�к�))
            If lngRow <> 0 And .TextMatrix(NewRow, COLP_�շ����) <> "" Then
                If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_�շ����)) > 0 _
                    Or .TextMatrix(NewRow, COLP_�շ����) = "4" And Val(.TextMatrix(NewRow, COLP_��������)) = 1 Then
                    '��ʾҩƷ���������ĵĿ��
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                        stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, COL_ҽ������) & "," & vsAdvice.TextMatrix(lngRow, COL_ִ�п���) & "���ÿ��:" & FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_���)), 5) & vsAdvice.TextMatrix(lngRow, COL_סԺ��λ)
                    Else
                        'ͬһ������ȡ:ҩƷ��סԺ��λ,���İ��ۼ۵�λ
                        stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "," & .TextMatrix(NewRow, COLP_ִ�п���) & "���ÿ��:" & _
                            FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_�շ�ϸĿID)), Val(.TextMatrix(NewRow, COLP_ִ�п���ID))), 5) & .TextMatrix(NewRow, COLP_��λ)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵�������ص��кŷ�Χ��������ҩ;�����к�
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub InitAdviceTable()
'���ܣ���ʼ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;" & _
        "Ӥ��,550,1;ҽ������,3000,1;����,600,7;��λ,450,1;����,600,7;��λ,450,1;���,850,7;" & _
        "Ƶ��,1000,1;�÷�,1000,1;ҽ������,1500,1;ִ��ʱ��,1000,1;ִ�п���,850,1;ִ������,850,1;" & _
        "ID;���ID;ҽ��״̬;���˿���ID;��������ID;����ҽ��;����ʱ��;�������;������ĿID;�Ƽ�����;ִ������ID;" & _
        "ִ�п���ID;��������;ҩƷID;����ϵ��;סԺ��װ;סԺ��λ;�ɷ����;���;����;�ֽ�ʱ��;�״�ʱ��;ĩ��ʱ��;ǩ����"
    arrHead = Split(strHead, ";")
    With vsAdvice
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
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "�к�;�շ�ϸĿID;�̶�;���;�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2000,1;�Ƽ�����,900,7;" & _
        "����,800,7;��λ,500,1;����,1000,7;Ӧ�ս��,1050,7;ʵ�ս��,1050,7;ִ�п���,1000,1;��������,850,1;" & _
        "����,450,4;�շ����;ִ�п���ID;��������"
    arrHead = Split(strHead, ";")
    With vsPrice
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
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub DeleteCurRow(ByVal lngRow As Long, Optional ByVal blnDelCur As Boolean = True)
'���ܣ��ڴ���������嵥�Ĺ�����ɾ������������(��ҩ�ƻ��ҩ)
'������blnDelCur=�Ƿ�ɾ����ǰ��
    Dim lngҽ��ID As Long, lng���ID As Long, i As Long
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
        lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
                
        'ɾ����ǰ��
        If blnDelCur Then .RemoveItem lngRow
        
        'ɾ�������
        If lng���ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = lng���ID _
                    Or Val(.TextMatrix(i, COL_ID)) = lng���ID Then
                    .RemoveItem i
                End If
            Next
        Else
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Sub InitPriceRecordset()
'���ܣ���ʼ��ҽ���Ƽۼ�¼��
    Set mrsPrice = New ADODB.Recordset
    
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
    mrsPrice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "�շ����", adVarChar, 1
    mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt
    mrsPrice.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble
    mrsPrice.Fields.Append "����", adDouble, , adFldIsNullable '��ۼ۸�
    mrsPrice.Fields.Append "����", adInteger '�����Ƿ��������
    mrsPrice.Fields.Append "����", adInteger
    mrsPrice.Fields.Append "�̶�", adInteger
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset)
'��ʼ����¼��
    'SQL��¼��
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "����", adInteger '1-�Ƽ�,2-ǩ��,3-У��,4-����,5-����,6-����
    rsSQL.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsSQL.Fields.Append "��ĿID", adBigInt '�շ�ϸĿID
    rsSQL.Fields.Append "���", adBigInt '��������
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '�Ƽ������ۼƼ�¼��
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsTotal.Fields.Append "��ĿID", adBigInt
    rsTotal.Fields.Append "�ⷿID", adBigInt
    rsTotal.Fields.Append "����", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    'ҽ���ϴ����ʵ�
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsUpload.Fields.Append "NO", adVarChar, 10
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
End Sub

Private Function LoadAdvicePrice(ByVal lngRow As Long, rsSend As ADODB.Recordset, cur��� As Currency) As Boolean
'���ܣ���ȡָ��ҽ��(����ǰ��)�ļƼ۹�ϵ����ʱ��¼��,������ȱʡ���ͽ��(���ѱ����)
'���أ�cur���=�������ҽ�����ͽ��(��ҩ���δ��,��Ҫ����۸�����)
    Dim rsTmp As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim strSQL As String, blnDo As Boolean, i As Long
    Dim dbl���� As Double, dbl���� As Double
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim bln�������� As Boolean, lng��ĿID As Long
    Dim lng������ID As Long, blnHaveSub As Boolean
    Dim lngִ�п���ID As Long
    
    On Error GoTo errH
    
    cur��� = 0
    With vsAdvice
        If InStr(",5,6,7,", rsSend!�������) > 0 Then
            '��ΪԺ��ִ��(�Ա�ҩ),ҩƷ������Ϊ����,�ҹ̶������Ƽ�
            If Nvl(rsSend!ִ������, 0) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!ҽ��ID = rsSend!ID
                mrsPrice!���ID = rsSend!���ID
                mrsPrice!�շ���� = rsSend!�������
                mrsPrice!�շ�ϸĿID = rsSend!ҩƷID
                mrsPrice!ִ�п���ID = rsSend!ִ�п���ID
                mrsPrice!���� = 1 'ҩƷ�̶�Ϊ1
                mrsPrice!���� = 0 'ҩƷ�̶�
                mrsPrice!�̶� = 1 'ҩƷ�̶�
                mrsPrice!���� = 0
                                
                '���͵���������
                If rsSend!������� = "7" Then
                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                    If Nvl(rsSend!�ɷ����, 0) = 0 Then
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����)) / Nvl(rsSend!����ϵ��, 1)
                    Else
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) _
                            * IntEx(Val(.TextMatrix(lngRow, COL_����)) / Nvl(rsSend!����ϵ��, 1) / Nvl(rsSend!סԺ��װ, 1)) * Nvl(rsSend!סԺ��װ, 1)
                    End If
                Else
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Nvl(rsSend!סԺ��װ, 1)
                End If
                dbl���� = Format(dbl����, "0.00000")
                                
                '��¼�ۼ۵���
                If Nvl(rsSend!�Ƿ���, 0) = 0 Then
                    mrsPrice!���� = Format(CalcPrice(rsSend!ҩƷID, , , True), "0.00000")
                Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                    mrsPrice!���� = Format(CalcDrugPrice(rsSend!ҩƷID, Nvl(rsSend!ִ�п���ID, 0), dbl����, , True), "0.00000")
                End If
                mrsPrice.Update
                                
                '����ҽ�����ͽ��(���ѱ���۵�ʵ�ս��)
                If Not IsNull(mrsPati!�ѱ�) Then
                    If Nvl(rsSend!�Ƿ���, 0) = 0 Then
                        cur��� = Format(CalcPrice(rsSend!ҩƷID, mrsPati!�ѱ�, dbl����, , Nvl(rsSend!ִ�п���ID, 0)), gstrDec)
                    Else
                        cur��� = Format(CalcDrugPrice(rsSend!ҩƷID, Nvl(rsSend!ִ�п���ID, 0), dbl����, mrsPati!�ѱ�), "0.00000")
                    End If
                Else
                    If gbln�Ӱ�Ӽ� Then
                        '����Ӱ�Ӽ�
                        If Nvl(rsSend!�Ƿ���, 0) = 0 Then
                            dbl���� = Format(CalcPrice(rsSend!ҩƷID), "0.00000")
                        Else '���ۼۼ���ҩƷʱ��,�Ա�ҩʱ�޶�Ӧҩ��
                            dbl���� = Format(CalcDrugPrice(rsSend!ҩƷID, Nvl(rsSend!ִ�п���ID, 0), dbl����), "0.00000")
                        End If
                        cur��� = Format(mrsPrice!���� * dbl���� * dbl����, gstrDec)
                    Else
                        cur��� = Format(mrsPrice!���� * dbl���� * mrsPrice!����, gstrDec)
                    End If
                End If
            End If
        Else
            'ȡ�����շѹ�ϵ�еĶ���(����ʱ�Ŷ��Ƽ�):�����Ƽ�,��Ϊ������Ժ��ִ��
            If Nvl(rsSend!�Ƽ�����, 0) = 0 And InStr(",0,5,", Nvl(rsSend!ִ������, 0)) = 0 Then
                dbl���� = Format(Val(.TextMatrix(lngRow, COL_����)), "0.00000")
                bln�������� = (rsSend!������� = "F" And Not IsNull(rsSend!���ID))
                
                '�ȶ�ȡ���еļƼ�
                strSQL = IIF(bln��������, "*Nvl(B.�����շ���,100)/100", "")
                strSQL = _
                    " Select C.���,A.�շ�ϸĿID as �շ���ĿID,A.���� as �շ�����,Nvl(E.���ж���,0) as ���ж���," & _
                    " B.������ĿID,C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,Decode(C.�Ƿ���,1,A.����,B.�ּ�)" & strSQL & " as ����," & _
                    " C.�Ƿ���,Nvl(A.����,0) as ����,D.��������,Nvl(A.ִ�п���ID,[3]) as ִ�п���ID,C.���ηѱ�" & _
                    " From ����ҽ���Ƽ� A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D,�����շѹ�ϵ E" & _
                    " Where A.ҽ��ID=[1] And E.������ĿID(+)=[2] And A.�շ�ϸĿID=E.�շ���ĿID(+)" & _
                    " And A.�շ�ϸĿID=B.�շ�ϸĿID And A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=D.����ID(+)" & _
                    " And C.������� IN(2,3) And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " Order by ����,A.�շ�ϸĿID"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!ID), Val(rsSend!������ĿID), Val(Nvl(rsSend!ִ�п���ID, 0)))
                
                'û�����ȡĬ�ϵļƼ�
                If rsTmp.EOF Then
                    strSQL = IIF(bln��������, "*Nvl(B.�����շ���,100)/100", "")
                    strSQL = _
                        " Select C.���,A.�շ���ĿID,A.�շ�����,A.���ж���,B.������ĿID," & _
                        " C.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,Decode(C.�Ƿ���,1,NULL,B.�ּ�)" & strSQL & " as ����," & _
                        " C.�Ƿ���,Nvl(A.������Ŀ,0) as ����,D.��������,[2] as ִ�п���ID,C.���ηѱ�" & _
                        " From �����շѹ�ϵ A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,�������� D" & _
                        " Where A.������ĿID=[1]" & _
                        " And A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And A.�շ���ĿID=D.����ID(+)" & _
                        " And C.������� IN(2,3) And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                        " Order by ����,A.�շ���ĿID"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!������ĿID), Val(Nvl(rsSend!ִ�п���ID, 0)))
                End If
                
                'ȷ���Ƽ�֮���Ƿ���������Լ���������ID
                If Not rsTmp.EOF And gbln��������ۿ� Then
                    Do While Not rsTmp.EOF
                        If Nvl(rsTmp!����, 0) = 0 Then
                            'SQL����������ǰ��,ֻȡ����Ŀ�ĵ�һ������
                            If lng������ID = 0 Then lng������ID = rsTmp!������ĿID
                        ElseIf Nvl(rsTmp!����, 0) = 1 Then
                            blnHaveSub = True: Exit Do
                        End If
                        rsTmp.MoveNext
                    Loop
                    rsTmp.MoveFirst
                End If
                
                Do While True
                    blnDo = False
                    If rsTmp.EOF Then
                        If lng��ĿID <> 0 Then blnDo = True
                    Else
                        If rsTmp!�շ���ĿID <> lng��ĿID And lng��ĿID <> 0 Then blnDo = True
                    End If
                    If blnDo Then
                        If Not IsNull(mrsPrice!����) Then
                            mrsPrice!���� = Format(mrsPrice!����, "0.00000")
                        End If
                        mrsPrice.Update
                        
                        'ҽ�����ͽ��
                        cur��� = cur��� + Format(curʵ��, gstrDec)
                    End If
                    If rsTmp.EOF Then Exit Do
                    
                    '------------------------------------
                    If rsTmp!�շ���ĿID <> lng��ĿID Then
                        curʵ�� = 0
                        mrsPrice.AddNew
                        mrsPrice!ҽ��ID = rsSend!ID
                        mrsPrice!���ID = rsSend!���ID
                        mrsPrice!�շ���� = rsTmp!���
                        mrsPrice!�շ�ϸĿID = rsTmp!�շ���ĿID
                        mrsPrice!���� = Nvl(rsTmp!�շ�����, 0)
                        mrsPrice!���� = Nvl(rsTmp!��������, 0)
                        mrsPrice!�̶� = Nvl(rsTmp!���ж���, 0)
                        mrsPrice!���� = Nvl(rsTmp!����, 0)
                        
                        'ִ�п���:��ҩ��ҩƷ���������ĵ�ר��ȡ
                        lngִ�п���ID = Nvl(rsTmp!ִ�п���ID, 0)
                        If rsTmp!��� = "4" And Nvl(rsTmp!��������, 0) = 1 Or InStr(",5,6,7,", rsTmp!���) > 0 Then
                            lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!�շ���ĿID, 4, Nvl(rsSend!���˿���ID, 0), 0, 2, lngִ�п���ID)
                        End If
                        If lngִ�п���ID <> 0 Then
                            mrsPrice!ִ�п���ID = lngִ�п���ID
                        Else
                            mrsPrice!ִ�п���ID = Null
                        End If
                    End If
                    lng��ĿID = rsTmp!�շ���ĿID
                    
                    '���㵥�ۺ�ʵ��
                    If Nvl(rsTmp!�Ƿ���, 0) = 0 Then '�̶��۸�
                        mrsPrice!���� = Nvl(mrsPrice!����, 0) + rsTmp!����
                        
                        curӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(rsTmp!����, "0.00000")
                        
                        '����Ӱ�Ӽ�
                        If gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                            curӦ�� = curӦ�� * (1 + Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                        End If
                        
                        curӦ�� = Format(curӦ��, gstrDec)
                        
                        If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And Nvl(rsTmp!���ηѱ�, 0) = 0 Then
                            curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ�, rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1, Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                        Else
                            curʵ�� = curʵ�� + curӦ��
                        End If
                    ElseIf InStr(",5,6,7,", rsTmp!���) > 0 Then
                        '��ҩ��ҩƷ�Ƽ۰�ʱ�ۼ���(��һ������),���������Ҫ��ҽ������
                        mrsPrice!���� = CalcDrugPrice(rsTmp!�շ���ĿID, Nvl(mrsPrice!ִ�п���ID, 0), dbl���� * Nvl(rsTmp!�շ�����, 0), , True)
                        
                        curӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(mrsPrice!����, "0.00000")
                        
                        '����Ӱ�Ӽ�
                        If gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                            curӦ�� = curӦ�� * (1 + Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                        End If

                        curӦ�� = Format(curӦ��, gstrDec)
                        
                        If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And Nvl(rsTmp!���ηѱ�, 0) = 0 Then
                            curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ�, rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1, Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                        Else
                            curʵ�� = curʵ�� + curӦ��
                        End If
                    ElseIf rsTmp!��� = "4" And Nvl(rsTmp!��������, 0) = 1 Then
                        '�������õ�ʱ�����ĺ�ҩƷһ������
                        mrsPrice!���� = CalcDrugPrice(rsTmp!�շ���ĿID, Nvl(mrsPrice!ִ�п���ID, 0), dbl���� * Nvl(rsTmp!�շ�����, 0), , True)
                        
                        curӦ�� = Format(mrsPrice!���� * dbl����, "0.00000") * Format(mrsPrice!����, "0.00000")
                        
                        '����Ӱ�Ӽ�
                        If gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                            curӦ�� = curӦ�� * (1 + Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                        End If

                        curӦ�� = Format(curӦ��, gstrDec)
                        
                        If Not IsNull(mrsPati!�ѱ�) And Not (gbln��������ۿ� And blnHaveSub) And Nvl(rsTmp!���ηѱ�, 0) = 0 Then
                            curʵ�� = curʵ�� + Format(ActualMoney(mrsPati!�ѱ�, rsTmp!������ĿID, curӦ��, rsTmp!�շ���ĿID, lngִ�п���ID, _
                                mrsPrice!���� * dbl����, IIF(gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1, Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                        Else
                            curʵ�� = curʵ�� + curӦ��
                        End If
                    End If
                    
                    rsTmp.MoveNext
                Loop
                
                '������Ŀ���ܼ����ۿ�
                If gbln��������ۿ� And blnHaveSub And lng������ID <> 0 Then
                    cur��� = Format(ActualMoney(Nvl(mrsPati!�ѱ�), lng������ID, cur���), gstrDec)
                End If
            End If
        End If
    End With
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetComboList(ByVal lngRow As Long) As String
'���ܣ����ݵ�ǰҽ���л�ȡ��ѡ��ļƼ�ҽ������
'������lngRow=�ɼ���(ҩ�ƻ��ҩ)
'˵����ע�������Ǹ��ݾ���ҽ����ȡ,��סԺ��ͬ
    Dim strCombo As String
    Dim strTmp As String, lngTmp As Long
    Dim i As Long, j As Long
    
    With vsAdvice
        If .Cell(flexcpData, lngRow, COL_ID) = 3 Then
            '��ҩ�÷�����ҩ�÷�,��ҩ�巨
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            For i = lngTmp To lngRow
                If InStr(",2,3,", CLng(.Cell(flexcpData, i, COL_ID))) > 0 Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        For j = 1 To mrsPrice.RecordCount
                            If Nvl(mrsPrice!�̶�, 0) = 0 Then
                                If .Cell(flexcpData, i, COL_ID) = 2 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�巨-" & .Cell(flexcpData, i, COL_ҽ������)
                                ElseIf .Cell(flexcpData, i, COL_ID) = 3 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";��ҩ�÷�-" & .Cell(flexcpData, i, COL_ҽ������)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    End If
                End If
            Next
        ElseIf .Cell(flexcpData, lngRow, COL_ID) = 4 Then
            '�ɼ�������
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            For i = lngTmp To lngRow
                If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                    For j = 1 To mrsPrice.RecordCount
                        If Nvl(mrsPrice!�̶�, 0) = 0 Then
                            If .TextMatrix(i, COL_�������) = "C" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";������Ŀ-" & .Cell(flexcpData, i, COL_ҽ������)
                            ElseIf .TextMatrix(i, COL_�������) = "E" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";�ɼ�����-" & .Cell(flexcpData, i, COL_ҽ������)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                        End If
                        mrsPrice.MoveNext
                    Next
                End If
            Next
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '���г�ҩ����ҩ;��
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) <> Val(.TextMatrix(lngRow, COL_���ID)) Then
                lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), lngRow + 1, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_ִ������ID))) = 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                    For j = 1 To mrsPrice.RecordCount
                        If Nvl(mrsPrice!�̶�, 0) = 0 Then
                            strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";��ҩ;��-" & .Cell(flexcpData, lngTmp, COL_ҽ������)
                            Exit For
                        End If
                        mrsPrice.MoveNext
                    Next
                End If
            End If
        Else
            'һ��������������ҽ��
            For i = lngRow To .Rows - 1
                If i = lngRow Or Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        For j = 1 To mrsPrice.RecordCount
                            If Nvl(mrsPrice!�̶�, 0) = 0 Then
                                If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, COL_ҽ������)
                                ElseIf .TextMatrix(i, COL_�������) = "G" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";��������-" & .Cell(flexcpData, i, COL_ҽ������)
                                ElseIf .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";��鲿λ-" & .Cell(flexcpData, i, COL_ҽ������)
                                Else
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_�������) & "ҽ��-" & .Cell(flexcpData, i, COL_ҽ������)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetComboList = Mid(strCombo, 2)
End Function

Private Function ShowAdvicePrice(ByVal lngRow As Long) As Boolean
'���ܣ�����ҽ���Ƽ۹�ϵ�����㲢��ʾָ��ҽ���ķ���(����ҽ�������ܶ���)
'������lngRow=�ɼ���(ҩ�ƻ��ҩ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTopRow As Long, lngLeftCol As Long
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnFirst As Boolean, str�Ƽ�ҽ�� As String
    Dim str��λ As String, dbl���� As Double
    Dim bln�������� As Boolean, strCombo As String, str�к� As String
    Dim dbl���� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim dbl��ǰ���� As Double, cur��ǰӦ�� As Currency, cur��ǰʵ�� As Currency
    Dim lng�к� As Long, cur�ϼ� As Currency
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
        
    On Error GoTo errH
    
    '���ڻ��ܼ����ۿ۵���ʱ��¼��
    rsMain.Fields.Append "ҽ���к�", adBigInt
    rsMain.Fields.Append "�����к�", adBigInt
    rsMain.Fields.Append "������ID", adBigInt
    rsMain.Fields.Append "ҽ���ϼ�", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        blnFirst = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnFirst = False 'һ����ҩ���Ƿ��һҩƷ��
            End If
        End If
        
        If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            If blnFirst Then
                mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                    " Or ҽ��ID=" & Val(.TextMatrix(lngRow, COL_���ID))
            Else
                mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
        Else
            mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                " Or ���ID=" & Val(.TextMatrix(lngRow, COL_ID))
        End If
        
        For i = 1 To mrsPrice.RecordCount
            '�Ƽ�ҽ��
            bln�������� = False
            lng�к� = .FindRow(CStr(mrsPrice!ҽ��ID), , COL_ID)
            If InStr(",5,6,7", .TextMatrix(lng�к�, COL_�������)) > 0 Then
                str�Ƽ�ҽ�� = "ҩƷҽ��-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 1 Then
                str�Ƽ�ҽ�� = "��ҩ;��-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 2 Then
                str�Ƽ�ҽ�� = "��ҩ�巨-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 3 Then
                str�Ƽ�ҽ�� = "��ҩ�÷�-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            ElseIf CLng(.Cell(flexcpData, lng�к�, COL_ID)) = 4 Then
                str�Ƽ�ҽ�� = "�ɼ�����-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "C" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "������Ŀ-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "F" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                bln�������� = True
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "G" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "��������-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            ElseIf .TextMatrix(lng�к�, COL_�������) = "D" And Val(.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                str�Ƽ�ҽ�� = "��鲿λ-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            Else
                str�Ƽ�ҽ�� = .Cell(flexcpData, lng�к�, COL_�������) & "ҽ��-" & .Cell(flexcpData, lng�к�, COL_ҽ������)
            End If
            str�Ƽ�ҽ�� = Replace(str�Ƽ�ҽ��, "'", "''")
            
            '����:ҩƷ��סԺ��λ������,��������������
            If InStr(",5,6,", .TextMatrix(lng�к�, COL_�������)) > 0 Then
                dbl���� = Val(.TextMatrix(lng�к�, COL_����))
            ElseIf .TextMatrix(lng�к�, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Val(.TextMatrix(lng�к�, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����)) * Val(.TextMatrix(lng�к�, COL_����)) _
                        / Val(.TextMatrix(lng�к�, COL_����ϵ��)) / Val(.TextMatrix(lng�к�, COL_סԺ��װ))
                Else
                    dbl���� = Val(.TextMatrix(lng�к�, COL_����)) _
                        * IntEx(Val(.TextMatrix(lng�к�, COL_����)) / Val(.TextMatrix(lng�к�, COL_����ϵ��)) / Val(.TextMatrix(lng�к�, COL_סԺ��װ)))
                End If
            Else
                dbl���� = Val(.TextMatrix(lng�к�, COL_����))
            End If
            dbl���� = Format(dbl���� * Nvl(mrsPrice!����, 0), "0.00000")
                        
            '���SQL
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & i & " as ���," & mrsPrice!ҽ��ID & " as ҽ��ID,ID as �շ�ϸĿID," & _
                Nvl(mrsPrice!�̶�, 0) & " as �̶�,'" & str�Ƽ�ҽ�� & "' as �Ƽ�ҽ��,���,����,����,���," & _
                "���㵥λ as ��λ," & Nvl(mrsPrice!����, 0) & " as �Ƽ�����," & dbl���� & " as ����," & _
                Format(Nvl(mrsPrice!����, 0), "0.00000") & " as ����,��������," & lng�к� & " as �к�," & _
                " �Ƿ���,�Ӱ�Ӽ�," & IIF(bln��������, 1, 0) & " as ��������," & mrsPrice!���� & " as ����," & _
                Nvl(mrsPrice!ִ�п���ID, 0) & " as ִ�п���ID,���ηѱ� From �շ���ĿĿ¼ Where ID=" & mrsPrice!�շ�ϸĿID
            mrsPrice.MoveNext
        Next
    End With
    
    With vsPrice
        lngPreRow = .Row: lngPreCol = .Col
        lngTopRow = .TopRow: lngLeftCol = .LeftCol
        .Editable = flexEDNone
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        '��Ҫ�Ƽ۵�ҽ��ѡ��
        '���ݴ�����ҽ��ȡ�ɼƼ�ҽ��(���ܴ�mrsPriceȡ,��Ϊ�������շѹ�ϵ����ɾ��,����Ҳ�����ڼƼ���ȫ��ɾ��)
        strCombo = GetComboList(lngRow)
        If strCombo <> "" Then
            .ColData(COLP_�Ƽ�ҽ��) = strCombo
            .Editable = flexEDKbdMouse '����ѡ������Ա༭
        Else
            .ColData(COLP_�Ƽ�ҽ��) = ""
        End If
        
        '��ʾ���еļƼ���Ŀ
        If strSQL <> "" Then
            strSQL = "Select A.�к�,A.�շ�ϸĿID,A.�̶�,A.����,A.�Ƽ�ҽ��,A.���,C.���� as �������,A.ִ�п���ID,G.���� as ִ�п���," & _
                " Nvl(E.����,A.����)||Decode(A.����,NULL,NULL,'('||A.����||')')||Decode(A.���,NULL,NULL,' '||A.���) as ����," & _
                " A.��λ,A.�Ƽ�����,A.����,D.סԺ��װ,D.סԺ��λ,Decode(A.�Ƿ���,1,A.����,B.�ּ�) as ����,F.��������," & _
                " A.��������,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.ԭ��,B.�ּ�,A.��������,B.�����շ���,B.������ĿID" & _
                " From (" & strSQL & ") A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,�շ���Ŀ���� E,�������� F,���ű� G" & _
                " Where A.�շ�ϸĿID=B.�շ�ϸĿID And A.���=C.���� And A.�շ�ϸĿID=D.ҩƷID(+)" & _
                " And A.�շ�ϸĿID=F.����ID(+) And A.ִ�п���ID=G.ID(+)" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIF(gbln��Ʒ��, 3, 1) & _
                " Order by A.���"
                '��Ϊ������ǵ��ñ�����ˢ��,Ҫ���ֶ�̬��¼���м�¼˳��
                'Ҫ��֤��������ǰ��,LoadAdvicePriceʱ������������ǰ�棬���ұ༭��ֻ���ܼ��˴���
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'û��
            
            If Not rsTmp.EOF And gbln��������ۿ� Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str�к� <> rsTmp!�к� & "_" & rsTmp!�շ�ϸĿID Then
                    If str�к� <> "" Then
                        If Not (Val(.TextMatrix(.Rows - 1, COLP_���)) = 1 And dbl���� = 0) Then
                            .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, "0.00000")
                            .Cell(flexcpData, .Rows - 1, COLP_����) = .TextMatrix(.Rows - 1, COLP_����) '��¼���ڻָ�����
                            .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                            .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                        End If
                        cur�ϼ� = cur�ϼ� + Format(curʵ��, gstrDec)
                    End If
                    str�к� = rsTmp!�к� & "_" & rsTmp!�շ�ϸĿID
                    dbl���� = 0: curӦ�� = 0: curʵ�� = 0
                    .Rows = .Rows + 1
                    
                    '��ʶ�̶�����Ϊ��ɫ
                    If rsTmp!�̶� <> 0 Then
                        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HE0E0E0
                    End If

                    .TextMatrix(.Rows - 1, COLP_�к�) = rsTmp!�к�
                    .TextMatrix(.Rows - 1, COLP_�շ�ϸĿID) = rsTmp!�շ�ϸĿID
                    .TextMatrix(.Rows - 1, COLP_�̶�) = rsTmp!�̶�
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��) = rsTmp!�Ƽ�ҽ��
                    .TextMatrix(.Rows - 1, COLP_���) = rsTmp!�������
                    .TextMatrix(.Rows - 1, COLP_�շ����) = rsTmp!���
                    .TextMatrix(.Rows - 1, COLP_�շ���Ŀ) = rsTmp!����
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0) '�������
                    
                    dbl���� = Nvl(rsTmp!����, 0) '�ۼ��������ں��水�ɱ����ۼ���
                    If InStr(",5,6,7,", rsTmp!���) > 0 Then 'סԺ��װ
                        .TextMatrix(.Rows - 1, COLP_��λ) = Nvl(rsTmp!סԺ��λ)
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(Nvl(rsTmp!����, 0), 5)
                            dbl���� = dbl���� * Nvl(rsTmp!סԺ��װ, 1)
                        Else
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,���ת��Ϊҩ����λ��ʾʱ���������㴦��
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(Nvl(rsTmp!����, 0) / Nvl(rsTmp!סԺ��װ, 1), 5)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_��λ) = Nvl(rsTmp!��λ)
                        .TextMatrix(.Rows - 1, COLP_����) = FormatEx(Nvl(rsTmp!����, 0), 5)
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_ִ�п���) = Nvl(rsTmp!ִ�п���)
                    .TextMatrix(.Rows - 1, COLP_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID, 0)
                    .TextMatrix(.Rows - 1, COLP_��������) = Nvl(rsTmp!��������)
                    .TextMatrix(.Rows - 1, COLP_����) = IIF(Nvl(rsTmp!����, 0) = 0, "", "��")
                    .TextMatrix(.Rows - 1, COLP_��������) = Nvl(rsTmp!��������, 0)
                    
                    '��¼��������ָ�
                    .Cell(flexcpData, .Rows - 1, COLP_�Ƽ�ҽ��) = .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��)
                    .Cell(flexcpData, .Rows - 1, COLP_�շ���Ŀ) = .TextMatrix(.Rows - 1, COLP_�շ���Ŀ)
                    .Cell(flexcpData, .Rows - 1, COLP_�Ƽ�����) = .TextMatrix(.Rows - 1, COLP_�Ƽ�����)
                    .Cell(flexcpData, .Rows - 1, COLP_ִ�п���) = .TextMatrix(.Rows - 1, COLP_ִ�п���)
                    
                    '��¼�����������Ϣ���Ա����
                    If gbln��������ۿ� And rsTmp!���� = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!�к� & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!�к� & ",") = 0 Then
                            rsClone.Filter = "�к�=" & rsTmp!�к� & " And ����=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!ҽ���к� = rsTmp!�к�
                                rsMain!�����к� = .Rows - 1
                                rsMain!������ID = rsTmp!������ĿID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!�к�
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!�к�
                            End If
                        End If
                    End If
                    
                    '��ҩ��ҩƷ����������:��ʹ�̶�Ҳ�����޸�ִ�п���
                    If InStr(",5,6,7,", rsTmp!���) > 0 _
                        Or rsTmp!��� = "4" And Nvl(rsTmp!��������, 0) = 1 Then
                        .Editable = flexEDKbdMouse
                    End If
                End If
                
                '���ۼ��㴦��
                If InStr(",5,6,7,", rsTmp!���) > 0 Then
                    If Nvl(rsTmp!�Ƿ���, 0) = 0 Then
                        dbl��ǰ���� = Nvl(rsTmp!����, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, Nvl(rsTmp!ִ�п���ID, 0), Format(Nvl(rsTmp!����, 0) * Nvl(rsTmp!סԺ��װ, 1), "0.00000"), , True)
                        Else
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, Nvl(rsTmp!ִ�п���ID, 0), Format(Nvl(rsTmp!����, 0), "0.00000"), , True)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                        dbl��ǰ���� = Format(dbl��ǰ���� * Nvl(rsTmp!סԺ��װ, 1), "0.00000")
                        cur��ǰӦ�� = Format(Nvl(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    Else
                        cur��ǰӦ�� = Format(Nvl(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                        dbl��ǰ���� = Format(dbl��ǰ���� * Nvl(rsTmp!סԺ��װ, 1), "0.00000")
                    End If
                ElseIf rsTmp!��� = "4" And Nvl(rsTmp!��������, 0) = 1 And Nvl(rsTmp!�Ƿ���, 0) = 1 Then
                    '�������õ�ʱ�����ĺ�ҩƷһ������
                    dbl��ǰ���� = CalcDrugPrice(rsTmp!�շ�ϸĿID, Nvl(rsTmp!ִ�п���ID, 0), Format(Nvl(rsTmp!����, 0), "0.00000"), , True)
                    cur��ǰӦ�� = Format(Nvl(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                Else
                    dbl��ǰ���� = Format(Nvl(rsTmp!����, 0), "0.00000") '�������Ϊ��������û������
                    cur��ǰӦ�� = Format(Nvl(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    If Nvl(rsTmp!�Ƿ���, 0) = 1 Then '��¼��ҩ��۷�Χ
                        .TextMatrix(.Rows - 1, COLP_���) = 1
                        .Cell(flexcpData, .Rows - 1, COLP_Ӧ�ս��) = CCur(Nvl(rsTmp!ԭ��, 0))
                        .Cell(flexcpData, .Rows - 1, COLP_ʵ�ս��) = CCur(Nvl(rsTmp!�ּ�, 0))
                        .Editable = flexEDKbdMouse '��ҩƷ���,��ʹ�̶�Ҳ���Զ���
                    End If
                End If
                'Ӧ��
                If rsTmp!�������� = 1 Then
                    cur��ǰӦ�� = cur��ǰӦ�� * Nvl(rsTmp!�����շ���, 100) / 100
                End If
                '����Ӱ�Ӽ�
                If gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                    cur��ǰӦ�� = cur��ǰӦ�� * (1 + Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100)
                End If
                cur��ǰӦ�� = Format(cur��ǰӦ��, gstrDec)
                
                'ʵ��
                If gbln��������ۿ� And (rsTmp!���� = 1 Or InStr(strHaveSub & ",", "," & rsTmp!�к� & ",") > 0) Then
                    cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                    '�ۼ�ҽ���ϼ��������ۿ�
                    rsMain.Filter = "ҽ���к�=" & rsTmp!�к�
                    rsMain!ҽ���ϼ� = Nvl(rsMain!ҽ���ϼ�, 0) + cur��ǰʵ��
                    rsMain.Update
                ElseIf Nvl(rsTmp!���ηѱ�, 0) = 0 And Not IsNull(mrsPati!�ѱ�) Then
                    cur��ǰʵ�� = Format(ActualMoney(mrsPati!�ѱ�, rsTmp!������ĿID, cur��ǰӦ��, rsTmp!�շ�ϸĿID, Nvl(rsTmp!ִ�п���ID, 0), _
                        dbl����, IIF(gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1, Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                Else
                    cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                End If
                
                dbl���� = dbl���� + dbl��ǰ����
                curӦ�� = curӦ�� + cur��ǰӦ��
                curʵ�� = curʵ�� + cur��ǰʵ��
                
                rsTmp.MoveNext
            Next
            If str�к� <> "" Then
                If Not (Val(.TextMatrix(.Rows - 1, COLP_���)) = 1 And dbl���� = 0) Then
                    .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, "0.00000")
                    .Cell(flexcpData, .Rows - 1, COLP_����) = .TextMatrix(.Rows - 1, COLP_����) '��¼���ڻָ�����
                    .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                    .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                End If
                cur�ϼ� = cur�ϼ� + Format(curʵ��, gstrDec)
            End If
        End If
        
        '���ܼ����ۿ�
        If gbln��������ۿ� And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                cur��ǰʵ�� = Format(ActualMoney(Nvl(mrsPati!�ѱ�), rsMain!������ID, rsMain!ҽ���ϼ�), gstrDec)
                cur�ϼ� = cur�ϼ� - Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��))
                .TextMatrix(rsMain!�����к�, COLP_ʵ�ս��) = Format(Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��)) + (cur��ǰʵ�� - rsMain!ҽ���ϼ�), gstrDec)
                cur�ϼ� = cur�ϼ� + Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��))
                rsMain.MoveNext
            Loop
        End If
        
        '------------------------------------------------
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        '��λȱʡ��Ԫ
        If lngPreRow >= .FixedRows And lngPreRow <= .Rows - 1 Then
            .Row = lngPreRow
        Else
            .Row = .FixedRows
        End If
        If lngPreCol >= COLP_�Ƽ�ҽ�� And lngPreCol <= .Cols - 1 Then
            .Col = lngPreCol
        Else
            .Col = COLP_�Ƽ�ҽ��
        End If
        '��λ�������λ��
        If lngTopRow >= .FixedRows And lngTopRow <= .Rows - 1 Then
            .TopRow = lngTopRow
        End If
        If lngLeftCol >= COLP_�Ƽ�ҽ�� And lngLeftCol <= .Cols - 1 Then
            .LeftCol = lngLeftCol
        End If
        .Redraw = flexRDDirect
    End With
    
    '���»�����ʾ�ɼ��еķ���ҽ�����
    vsAdvice.TextMatrix(lngRow, COL_���) = Format(cur�ϼ�, gstrDec)
    ShowAdvicePrice = True
    
    Call ShowSendTotal
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long, Optional bln�Ǳ��� As Boolean) As Boolean
'���ܣ��жϼ۱��е�Ԫ���Ƿ���Ա༭
    Dim lng�к� As Long
    
    With vsPrice
        bln�Ǳ��� = False
        CellEditable = .Editable
        lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
        If lngCol = COLP_ִ�п��� Then
            '�������õ�����,��ҩ��ҩƷ�Ƽ۵�ִ�п��ҿ����޸�
            If Not (.TextMatrix(lngRow, COLP_�շ����) = "4" And Val(.TextMatrix(lngRow, COLP_��������)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_�շ����)) > 0 And InStr(",5,6,7,", vsAdvice.TextMatrix(lng�к�, COL_�������)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Or .TextMatrix(lngRow, COLP_�к�) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_�̶�)) <> 0 Then
            '�̶������н������޸ı��
            If Not (Val(.TextMatrix(lngRow, COLP_���)) = 1 And lngCol = COLP_����) Then
                CellEditable = False
            End If
        Else
            If lngCol = COLP_���� Then
                If Val(.TextMatrix(lngRow, COLP_���)) <> 1 Then
                    CellEditable = False
                Else
                    '�Ǳ���ִ�еı����Ŀ�������۸�
                    If lng�к� <> 0 Then
                        If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                            bln�Ǳ��� = True: CellEditable = False
                        End If
                    End If
                End If
            ElseIf lngCol <> COLP_�Ƽ�ҽ�� And lngCol <> COLP_�Ƽ����� And lngCol <> COLP_�շ���Ŀ Then
                CellEditable = False
            End If
        End If
    End With
End Function

Private Function LoadAdviceSend(Optional ByVal str���s As String) As Boolean
'���ܣ�����������ȡ����ʾҪ���͵�ҩƷҽ���嵥
'˵����ע��CellData�д�ŵ��и�������
'   RowData��0-δ���͵�,-1-�ѳɹ����͵�
'   COL_ѡ��0-������ѡ���,1-��ֹ�ı�ѡ��״̬��
'   COL_ID��1-��ҩ;����2-��ҩ�巨��3-��ҩ�÷���4-�ɼ�����
'   COL_Ӥ�������Ӥ�����
'   COL_������𣺴������������ƣ�������ʾ�Ƽ�ҽ��
'   COL_ҽ�����ݣ����������Ŀ���ƻ�걾��λ��������ʾ�Ƽ�ҽ��
'   COL_�ֽ�ʱ�䣺��ŷ��õķ���ʱ��(�޷ֽ�ʱ��ʱ)
'   COL_Ƶ�ʣ�1-"һ����"����
'   COL_��ԭʼ�Ľ��������ۼ���ʾ��
    Dim rsSend As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long, strTmp As String
    Dim lngRow As Long, lngDel��ID As Long
    Dim bln����ʱ�� As Boolean, lng���� As Long, lng��С���� As Long
    Dim str�ֽ�ʱ�� As String, dbl���� As Double, cur��� As Currency
    
    Dim vMsg As VbMsgBoxResult, blnʱ����ʾ As Boolean
    Dim bln�����ʾ As Boolean, blnĬ�Ϸ��� As Boolean
    Dim str�÷� As String, i As Long, j As Long
        
    Screen.MousePointer = 11
    
    stbThis.Panels(3).Text = "": Call Form_Resize
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    vsAdvice.Rows = vsAdvice.FixedRows '��ɾ���й���
    
    vsAdvice.ColHidden(COL_Ӥ��) = True
    Me.Refresh
    
    Call InitPriceRecordset '�Ƽ۹�ϵ��
    
    '��ȡ�����嵥:�¿�����У��ÿ��ҽ����¼(ҩƷ�ͷ�ҩƷ),����ҽ��Ϊ����
    '----------------------------------------------------------------------------------------------------------
    '����������ȼ�������ҽ��������,�������ȶ�ȡ����(��ҩ;��,�÷�,�巨,�ɼ�����)
    strSQL = _
        " Select A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���,A.ҽ��״̬," & _
        " A.�������,F.���� as �������,A.������ĿID,B.���� as ������Ŀ,A.�շ�ϸĿID as ҩƷID,A.Ӥ��," & _
        " A.ҽ������,A.�걾��λ,A.����,A.�ܸ�����,D.סԺ��λ,A.��������,B.���㵥λ,D.����ϵ��,D.סԺ��װ," & _
        " A.��ʼִ��ʱ��,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ҽ������,A.ִ��ʱ�䷽��," & _
        " A.���˿���ID,A.��������ID,A.����ҽ��,A.����ʱ��,A.�Ƽ�����,A.ִ������,A.ִ�п���ID,E.���� as ִ�п���," & _
        " B.��������,D.�ɷ����,D.ҩ������,C.�Ƿ���,C.����ʱ��,C.�������,S.ǩ��ID" & _
        " From ����ҽ����¼ A,������ĿĿ¼ B,�շ���ĿĿ¼ C,ҩƷ��� D,���ű� E,������Ŀ��� F,����ҽ��״̬ S,����ҽ����¼ X" & _
        " Where A.����ID=[1] And A.��ҳID=[2] And Nvl(A.ǰ��ID,0)=[3] And A.ID=S.ҽ��ID And S.��������=1" & _
        " And A.ҽ��״̬ IN(1,3,5) And A.ҽ����Ч=1 And A.���ID=X.ID(+) And B.���=F.����" & _
        " And A.������ĿID=B.ID And A.�շ�ϸĿID=C.ID(+) And A.�շ�ϸĿID=D.ҩƷID(+)" & _
        " And A.ִ�п���ID=E.ID(+) And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & _
        " And Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1)=[4]" & _
        " And Exists(Select M.���� From ��Ա�� M,ִҵ��� N Where M.����=Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1) And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ'))" & _
        " And Not(A.�������='H' And B.��������='1') And Not(A.�������='Z' And B.��������='4')" & _
        " Order by A.Ӥ��,���,��ID,A.���"
    
    On Error GoTo errH
    Set rsSend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mlngǰ��ID, UserInfo.����)
    
    '���㲢��ʾ�����嵥
    '----------------------------------------------------------------------------------------------------------
    If Not rsSend.EOF Then
        With vsAdvice
            blnʱ����ʾ = True: bln�����ʾ = True: blnĬ�Ϸ��� = True
            .Redraw = flexRDNone
            For i = 1 To rsSend.RecordCount
                'һ��ҽ���е�һ�����ܷ���,�����鲻�ܷ���
                If lngDel��ID <> 0 Then
                    If Nvl(rsSend!���ID, rsSend!ID) = lngDel��ID Then
                        GoTo NextLoop
                    Else
                        lngDel��ID = 0
                    End If
                End If
                
                '��鲻�����͵��������
                'һ��ҽ������һ��ҽ����,��������в����
                If str���s <> "" And lngDel��ID = 0 Then
                    If rsSend!������� = "7" Then
                        '��ҩ�䷽
                        If InStr(str���s, "'8'") = 0 Then lngDel��ID = rsSend!���ID
                    ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                        '������ҩ(��������ҩ���г�ҩ���һ����ҩ�����)
                        If InStr(str���s, "'" & rsSend!������� & "'") = 0 Then
                            lngDel��ID = rsSend!���ID
                            'ɾ���ѿ��ܼ��������һ����ҩ��,��ǰ����δ���벻ɾ��
                            Call DeleteCurRow(lngRow, False)
                            lng��С���� = 0
                        End If
                    ElseIf rsSend!������� = "D" Then
                        '������(������ļ��)
                        If InStr(str���s, "'D'") = 0 Then lngDel��ID = rsSend!ID
                    ElseIf rsSend!������� = "F" Then
                        '�������(�����������)
                        If InStr(str���s, "'F'") = 0 Then lngDel��ID = rsSend!ID
                    ElseIf rsSend!������� = "C" Then
                        '�������(������ļ���)
                        If InStr(str���s, "'C'") = 0 Then lngDel��ID = Nvl(rsSend!���ID, rsSend!ID)
                    ElseIf IsNull(rsSend!���ID) And rsSend!ID <> Val(.TextMatrix(.Rows - 1, COL_���ID)) Then
                        '����������Ŀ
                        If InStr(str���s, "'" & rsSend!������� & "'") = 0 Then lngDel��ID = rsSend!ID
                    End If
                    If lngDel��ID <> 0 Then GoTo NextLoop
                End If
                                                
                '���뵱ǰ��
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .Cell(flexcpPictureAlignment, lngRow, COL_ѡ��) = 4
                Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("T").Picture
                
                '���������
                If rsSend!������� = "7" Then
                    .RowHidden(lngRow) = True '�в�ҩ
                ElseIf rsSend!������� = "E" Then
                    If Not IsNull(rsSend!���ID) Then
                        .RowHidden(lngRow) = True
                        .Cell(flexcpData, lngRow, COL_ID) = 2 '��ҩ�巨
                    ElseIf Val(.TextMatrix(lngRow - 1, COL_���ID)) = rsSend!ID Then
                        If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_ID) = 1 '��ҩ;��
                        ElseIf .TextMatrix(lngRow - 1, COL_�������) = "C" Then
                            .Cell(flexcpData, lngRow, COL_ID) = 4 '�ɼ�����
                        Else
                            .Cell(flexcpData, lngRow, COL_ID) = 3 '��ҩ�÷�
                        End If
                    End If
                ElseIf InStr(",5,6,", rsSend!�������) = 0 And Not IsNull(rsSend!���ID) Then
                    '��������,��������,��鲿λ,һ���ɼ��ļ�����Ŀ
                    .RowHidden(lngRow) = True
                End If
                
                '�ſ�һ��Ķ���(������ҩ;��,��ҩ�巨,�÷�,�ɼ�����)
                If Nvl(rsSend!ִ������, 0) = 0 Then
                    If InStr(",1,2,3,4,", CLng(.Cell(flexcpData, lngRow, COL_ID))) = 0 _
                        And InStr(",5,6,7,", rsSend!�������) = 0 Then
                        Call .RemoveItem(lngRow): GoTo NextLoop
                    End If
                End If
                
                'һ���и�ֵ
                '---------------------------------------------------------------
                .Cell(flexcpData, lngRow, COL_Ӥ��) = CLng(Nvl(rsSend!Ӥ��, 0))
                If Nvl(rsSend!Ӥ��, 0) = 0 Then
                    .TextMatrix(lngRow, COL_Ӥ��) = "����"
                Else
                    .TextMatrix(lngRow, COL_Ӥ��) = "Ӥ��" & rsSend!Ӥ��
                    .ColHidden(COL_Ӥ��) = False '��Ӥ��ҽ��ʱ����ʾ
                End If
                
                .TextMatrix(lngRow, COL_ID) = rsSend!ID
                .TextMatrix(lngRow, COL_���ID) = Nvl(rsSend!���ID)
                .TextMatrix(lngRow, COL_ҽ��״̬) = rsSend!ҽ��״̬
                .TextMatrix(lngRow, COL_�������) = rsSend!�������
                .TextMatrix(lngRow, COL_������ĿID) = rsSend!������ĿID
                .TextMatrix(lngRow, COL_ҽ������) = Nvl(rsSend!ҽ������)
                
                '����ǩ����ʶ
                .TextMatrix(lngRow, COL_ǩ��ID) = Nvl(rsSend!ǩ��ID)
                If Val(.TextMatrix(lngRow, COL_ǩ��ID)) <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, COL_ҽ������) = img16.ListImages("ǩ��").Picture
                End If
                
                '������ʾ�Ƽ�ҽ��
                .Cell(flexcpData, lngRow, COL_�������) = CStr(Nvl(rsSend!�������))
                If Not IsNull(rsSend!���ID) And rsSend!������� = "D" Then
                    .Cell(flexcpData, lngRow, COL_ҽ������) = CStr(Nvl(rsSend!�걾��λ))
                Else
                    .Cell(flexcpData, lngRow, COL_ҽ������) = CStr(Nvl(rsSend!������Ŀ))
                End If
                
                .TextMatrix(lngRow, COL_ҽ������) = Nvl(rsSend!ҽ������)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = Nvl(rsSend!ִ��ʱ�䷽��)
                .TextMatrix(lngRow, COL_Ƶ��) = Nvl(rsSend!ִ��Ƶ��)
                
                .TextMatrix(lngRow, COL_���˿���ID) = Nvl(rsSend!���˿���ID)
                .TextMatrix(lngRow, COL_��������ID) = Nvl(rsSend!��������ID)
                .TextMatrix(lngRow, COL_����ҽ��) = Nvl(rsSend!����ҽ��)
                .TextMatrix(lngRow, COL_����ʱ��) = Format(Nvl(rsSend!����ʱ��), "yyyy-MM-dd HH:mm:ss")
                                
                '�ɼ���������ʾ������Ŀ��ִ�п���
                If Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                    .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngRow - 1, COL_ִ�п���)
                Else
                    .TextMatrix(lngRow, COL_ִ�п���) = Nvl(rsSend!ִ�п���)
                End If
                .TextMatrix(lngRow, COL_ִ�п���ID) = Nvl(rsSend!ִ�п���ID)
                
                .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsSend!�Ƽ�����, 0)
                .TextMatrix(lngRow, COL_ִ������ID) = Nvl(rsSend!ִ������, 0)
                .TextMatrix(lngRow, COL_��������) = Nvl(rsSend!��������)
                                
                'ҩƷ�����Ϣ
                If InStr(",5,6,7", rsSend!�������) > 0 Then
                    'ҩƷ��Ӧ�Ĺ���ѳ�����������(������Ŀ����Ҳ������ͬ����,Ŀǰ��δ����)
                    If Format(Nvl(rsSend!����ʱ��, "3000-01-01"), "yyyy-MM-dd") <> "3000-01-01" Or InStr(",2,3,", Nvl(rsSend!�������, 0)) = 0 Then
                        If rsSend!������� = "7" Then
                            strTmp = "���в�ҩ��Ӧ����ҩ�䷽�޷����ͣ�" & vbCrLf & vbCrLf & "����" & Nvl(rsSend!ҽ������)
                        Else
                            strTmp = "��ҩƷ(��һ����ҩ������ҩƷ)�޷����ͣ�" & vbCrLf & vbCrLf & "����" & Nvl(rsSend!ҽ������)
                        End If
                        strTmp = strTmp & vbCrLf & vbCrLf & "û�з�����Ч��ҩƷ�����Ϣ����ҩƷ�����Ѿ���ͣ�û�������סԺ���ˡ�"
                        strTmp = strTmp & vbCrLf & "���ȵ�ҩƷĿ¼�����д�����[ȷ��]������������ҽ����"
                        
                        .Redraw = flexRDDirect
                        Call .ShowCell(lngRow, COL_ѡ��)
                        Screen.MousePointer = 0
                        MsgBox strTmp, vbInformation, gstrSysName
                        
                        'ɾ����ǰ��(�������),��������һҽ��
                        Screen.MousePointer = 11
                        lngDel��ID = Nvl(rsSend!���ID, rsSend!ID)
                        Call DeleteCurRow(lngRow)
                        .Refresh: .Redraw = flexRDNone
                        lng��С���� = 0: GoTo NextLoop
                    End If
                
                    .TextMatrix(lngRow, COL_ҩƷID) = rsSend!ҩƷID
                    .TextMatrix(lngRow, COL_����ϵ��) = Nvl(rsSend!����ϵ��, 1)
                    .TextMatrix(lngRow, COL_סԺ��װ) = Nvl(rsSend!סԺ��װ, 1)
                    .TextMatrix(lngRow, COL_סԺ��λ) = Nvl(rsSend!סԺ��λ)
                    .TextMatrix(lngRow, COL_�ɷ����) = Nvl(rsSend!�ɷ����, 0)
                    .TextMatrix(lngRow, COL_���) = GetStock(rsSend!ҩƷID, Nvl(rsSend!ִ�п���ID, 0), 2) '��סԺ��װ
                End If
                                                                        
                '���㷢�ʹ�����ִ�еķֽ�ʱ���
                '---------------------------------------------------------------
                If rsSend!������� = "7" Then
                    .TextMatrix(lngRow, COL_����) = rsSend!�ܸ�����
                    If Not IsNull(rsSend!ִ��ʱ�䷽��) Then
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Calc�����ֽ�ʱ��(rsSend!�ܸ�����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", rsSend!ִ��ʱ�䷽��, rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(0), "MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(rsSend!�ܸ����� - 1), "MM-dd HH:mm")
                    Else
                        '�޷ֽ�ʱ��(��������δ����ִ��ʱ����޷��ֽ�)
                        '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = Nvl(rsSend!��������) '����
                    .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = rsSend!�ܸ����� '����
                    .TextMatrix(lngRow, COL_������λ) = "��"
                ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                    '����������ҩ����
                    If Nvl(rsSend!����, 0) <> 0 And Not IsNull(rsSend!ִ��Ƶ��) Then
                        'һ��Ƶ�����ڵĴ���
                        If rsSend!�����λ = "��" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / 7))
                        ElseIf rsSend!�����λ = "��" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��))
                        ElseIf rsSend!�����λ = "Сʱ" Then
                            lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��) * 24)
                        End If
                    Else
                        '�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,����һ��Ƶ�����ڵĴ�������
                        If Nvl(rsSend!�ɷ����, 0) = 0 And Nvl(rsSend!��������, 0) <> 0 Then
                            lng���� = IntEx(rsSend!�ܸ����� * rsSend!����ϵ�� / rsSend!��������)
                        Else
                            lng���� = Nvl(rsSend!Ƶ�ʴ���, 0)
                        End If
                    End If
                    If Not IsNull(rsSend!ִ��ʱ�䷽��) Then
                        str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", rsSend!ִ��ʱ�䷽��, rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                        If str�ֽ�ʱ�� <> "" Then
                            .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "MM-dd HH:mm")
                        End If
                    Else
                        '�޷ֽ�ʱ��(��������δ����ִ��ʱ����޷��ֽ�)
                        '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_����) = lng����
                    .TextMatrix(lngRow, COL_����) = FormatEx(Nvl(rsSend!��������), 5)
                    .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                    .TextMatrix(lngRow, COL_����) = FormatEx(rsSend!�ܸ����� / rsSend!סԺ��װ, 5) '��סԺ��λ��ʾ
                    .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!סԺ��λ)
                    
                    If lng���� < lng��С���� Or lng��С���� = 0 Then lng��С���� = lng����
                ElseIf rsSend!������� = "E" And CLng(.Cell(flexcpData, lngRow, COL_ID)) <> 0 Then
                    '��ҩ;��,��ҩ�巨,��ҩ�÷�,�ɼ�����
                    'һ����ҩ�İ���С��������(Ӱ���ҩ;���Ʒ�)
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '��ҩ;��
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_����)) > lng��С���� Then
                                    .TextMatrix(j, COL_����) = lng��С����
                                    If .TextMatrix(j, COL_�ֽ�ʱ��) <> "" Then
                                        .TextMatrix(j, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(j, COL_�ֽ�ʱ��))
                                        .TextMatrix(j, COL_�״�ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(0), "MM-dd HH:mm")
                                        .TextMatrix(j, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "MM-dd HH:mm")
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng��С���� = 0
                    End If
                    
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����) '���������
                    .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                    If .Cell(flexcpData, lngRow, COL_ID) = 3 Then '��ҩ�÷�
                        .TextMatrix(lngRow, COL_������λ) = "��"
                    Else
                        .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                    End If
                    
                    .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                    .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                    .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                    .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                Else
                    '������ҩ����:�ɼ�����������ķ�֧����������
                    If IsNull(rsSend!���ID) Or (Not IsNull(rsSend!���ID) And rsSend!������� = "C") Then '��Ҫҽ��,�����������
                        dbl���� = Nvl(rsSend!�ܸ�����, 1)
                        lng���� = IntEx(dbl���� / Nvl(rsSend!��������, 1))
                        
                        If IsNull(rsSend!ִ��ʱ�䷽��) And (Nvl(rsSend!Ƶ�ʴ���, 0) = 0 Or Nvl(rsSend!Ƶ�ʼ��, 0) = 0 Or IsNull(rsSend!�����λ)) Then
                            'ִ��Ƶ��Ϊ"һ����"����Ŀ
                            str�ֽ�ʱ�� = "" '����Ҫ
                            .Cell(flexcpData, lngRow, COL_Ƶ��) = 1
                        Else
                            'ִ��Ƶ��Ϊ"��ѡƵ��"����Ŀ:��ҽ��ʱӦ����������
                            If Not IsNull(rsSend!ִ��ʱ�䷽��) Then
                                str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", rsSend!ִ��ʱ�䷽��, rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                            Else
                                str�ֽ�ʱ�� = "" '����Ҳ��δ����ִ��ʱ��,�޷��ֽ�
                            End If
                        End If
                        .TextMatrix(lngRow, COL_����) = lng����
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                        If str�ֽ�ʱ�� <> "" Then
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "MM-dd HH:mm")
                        Else
                            '��¼���÷���ʱ��(���޷ֽ�ʱ��ʱ),��ҽ���Ŀ�ʼִ��ʱ��
                            .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = CStr(Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss"))
                        End If
                        
                        .TextMatrix(lngRow, COL_����) = FormatEx(Nvl(rsSend!��������), 5)
                        If Not IsNull(rsSend!��������) Then
                            .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                        End If
                        .TextMatrix(lngRow, COL_����) = FormatEx(dbl����, 5)
                        .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                    Else
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                        .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                        .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                        .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                        .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                    End If
                End If
                
                '������Ŀ���ͽ��
                cur��� = 0
                If Not LoadAdvicePrice(lngRow, rsSend, cur���) Then
                    lngDel��ID = Nvl(rsSend!���ID, rsSend!ID)
                    Call DeleteCurRow(lngRow)
                    lng��С���� = 0: GoTo NextLoop
                End If
                .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
                .Cell(flexcpData, lngRow, COL_���) = CCur(.TextMatrix(lngRow, COL_���))
                
                '�����ʱ��һЩ�����ۼ���ʾ���,��ҩ;��,�÷�,ִ�п���,ִ������
                '---------------------------------------------------------------
                If rsSend!������� = "E" And InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_ID))) > 0 Then '��ҩ;������ҩ�÷�
                    cur��� = 0
                    lngTmp = .FindRow(CStr(rsSend!ID), , COL_���ID)
                    
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '��ҩ;��
                        'һ����ҩʱ,��ҩ;���Ľ���ۼ���ʾ�ڵ�һ����ҩ��
                        .TextMatrix(lngTmp, COL_���) = Format(Val(.TextMatrix(lngTmp, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                        '��ʾ��ҩ;��,ִ������
                        For j = lngTmp To lngRow - 1
                            strTmp = ""
                            If Val(.TextMatrix(j, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                                strTmp = "�Ա�ҩ"
                            ElseIf Val(.TextMatrix(j, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                                strTmp = "��Ժ��ҩ"
                            End If
                            .TextMatrix(j, COL_ִ������) = strTmp
                            .TextMatrix(j, COL_�÷�) = rsSend!������Ŀ
                        Next
                    Else
                        'ҩƷ��ִ������
                        strTmp = ""
                        If Val(.TextMatrix(lngTmp, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                            strTmp = "�Ա�ҩ"
                        ElseIf Val(.TextMatrix(lngTmp, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                            strTmp = "��Ժ��ҩ"
                        End If
                    
                        '��ҩ�÷�,�巨
                        str�÷� = rsSend!������Ŀ
                        If Val(.Cell(flexcpData, lngRow - 1, COL_ID)) = 2 Then
                            str�÷� = str�÷� & "|" & GetItemField("������ĿĿ¼", Val(.TextMatrix(lngRow - 1, COL_������ĿID)), "����")
                        End If
                        For j = lngTmp To lngRow
                            .TextMatrix(j, COL_�÷�) = str�÷� '������д�շ���¼
                            cur��� = cur��� + Val(.TextMatrix(j, COL_���))
                        Next
                        .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
                        '��ʾִ������
                        .TextMatrix(lngRow, COL_ִ������) = strTmp
                        '��ʾ�䷽ִ�п���
                        .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngTmp, COL_ִ�п���)
                    End If
                    
                    'ʹ���ҽ��ѡ��״̬��ͬ(��Ϊ����ԭ�򣻷�ҩҽ������)
                    For j = lngTmp To lngRow
                        If .Cell(flexcpData, j, COL_ѡ��) <> 0 Then
                            Call RowSelectSame(j, COL_ѡ��)
                            Exit For 'һ����ֹ,ȫ����ֹ
                        End If
                    Next
                    If j > lngRow Then
                        For j = lngRow To lngTmp Step -1
                            If InStr(",5,6,7,", .TextMatrix(j, COL_�������)) > 0 Then
                                If .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                                    Call RowSelectSame(j, COL_ѡ��)
                                    Exit For '���ѡ,ȫ����ѡ
                                End If
                            End If
                        Next
                    End If
                ElseIf InStr(",5,6,7,", rsSend!�������) = 0 Then
                    If Not IsNull(rsSend!���ID) And rsSend!������� <> "C" Then
                        '������ҩҽ��
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_ID)) = rsSend!���ID Then
                                .TextMatrix(j, COL_���) = Format(Val(.TextMatrix(j, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                                Exit For
                            End If
                        Next
                    ElseIf Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                        '����걾�ɼ�����Ϊ��ʾ��
                        .TextMatrix(lngRow, COL_�÷�) = rsSend!������Ŀ
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                .TextMatrix(lngRow, COL_���) = Format(Val(.TextMatrix(lngRow, COL_���)) + Val(.TextMatrix(j, COL_���)), gstrDec)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If

                'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ),�Ա�ҩ�����
                '---------------------------------------------------------------
                If InStr(",5,6,7,", rsSend!�������) > 0 And Nvl(rsSend!ִ������, 0) <> 5 Then
                    Call CheckStock(lngRow, rsSend, bln�����ʾ, blnʱ����ʾ, blnĬ�Ϸ���)
                End If
                
NextLoop:       '---------------------------------------------------------------
                Progress = i / rsSend.RecordCount * 100
                txtPer.Text = CInt(psb.Value) & "%"
                txtPer.Refresh
                rsSend.MoveNext
            Next
        End With
    End If
    With vsAdvice
        .AutoSize COL_ҽ������
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        '����ǩ��ͼ�����
        .Cell(flexcpPictureAlignment, .FixedRows, COL_ҽ������, .Rows - 1, COL_ҽ������) = 0
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    LoadAdviceSend = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone: Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Sub CheckStock(ByVal lngRow As Long, rsSend As ADODB.Recordset, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean)
'���ܣ����ݿ���������鷢��ҩƷ�Ŀ��
'������lngRow=ҽ���к�,rsSend=��ǰ����ҽ����Ϣ
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = GetStockCheck(Val(.TextMatrix(lngRow, COL_ִ�п���ID)))
        bln���� = Nvl(rsSend!ҩ������, 0) = 1
        blnʱ�� = Nvl(rsSend!�Ƿ���, 0) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = .TextMatrix(lngRow, COL_סԺ��λ) '������ʾ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����:סԺ��װ
            If .TextMatrix(lngRow, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Val(.TextMatrix(lngRow, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����))
                    dbl���� = dbl���� / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ))
                Else
                    dbl���� = IntEx(Val(.TextMatrix(lngRow, COL_����)) / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ)))
                    dbl���� = dbl���� * Val(.TextMatrix(lngRow, COL_����))
                End If
            Else
                dbl���� = Val(.TextMatrix(lngRow, COL_����))
            End If
            
            '��ǰ���ÿ��:סԺ��װ,��ȥǰ����ͬҩƷҪ���͵Ŀ��
            For i = lngRow - 1 To .FixedRows Step -1
                blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                If blnDo Then
                    blnDo = Val(.TextMatrix(i, COL_ҩƷID)) = Val(.TextMatrix(lngRow, COL_ҩƷID)) _
                        And Val(.TextMatrix(i, COL_ִ�п���ID)) = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                End If
                If blnDo Then
                    blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                End If
                If blnDo Then
                    If .TextMatrix(i, COL_�������) = "7" Then
                        '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                        If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                            dbl�ѷ���� = dbl�ѷ���� + _
                                Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))
                        Else
                            dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                        End If
                    Else
                        dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                    End If
                End If
            Next
            dbl���ÿ�� = Val(.TextMatrix(lngRow, COL_���))
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        strTmp = "ҩ��������ʱ��ҩƷ""" & .TextMatrix(lngRow, COL_ҽ������) & """��治�㣺" & vbCrLf & vbCrLf & _
                            .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                            IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                            "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                    Else
                        strTmp = """" & .TextMatrix(lngRow, COL_ҽ������) & """��治�㣺" & vbCrLf & vbCrLf & _
                            .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                            IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                            "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҩƷ��"
                    End If
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                        End If
                    End If
                    
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function CheckPriceStock(ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, _
    rsTotal As ADODB.Recordset, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean) As Boolean
'���ܣ����͹�����ʱ���Է�ҩ��ҩƷ���������õ����ļƼ۽��п����(�ۼƼ��)
'������lngRow=ҽ���к�
'      dbl����=�Ѽ���õļƼ�����(�ۼ۵�λ)
'      rsTotal=��ǰ����ǰ�����ۼƷ��͵ļƼ�ҩƷ����������(�ۼ۵�λ)
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = GetStockCheck(lng�ⷿID)
        bln���� = Nvl(rsPrice!����, 0) = 1
        blnʱ�� = Nvl(rsPrice!�Ƿ���, 0) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = Nvl(rsPrice!סԺ��λ, Nvl(rsPrice!���㵥λ)) '������ʾ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����������:סԺ��װ
            dbl���� = Format(dbl���� / Nvl(rsPrice!סԺ��װ, 1), "0.00000")
            
            '��ǰ���ÿ��:סԺ��װ,��ȥǰ����ͬҩƷҽ��Ҫ���͵Ŀ��
            If InStr(",5,6,7,", rsPrice!���) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_ҩƷID)) = rsPrice!ID And Val(.TextMatrix(i, COL_ִ�п���ID)) = lng�ⷿID
                    End If
                    If blnDo Then
                        blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                    End If
                    If blnDo Then
                        If .TextMatrix(i, COL_�������) = "7" Then
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                dbl�ѷ���� = dbl�ѷ���� + _
                                    Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                    / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))
                            Else
                                dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                    * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                            End If
                        Else
                            dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                        End If
                    End If
                Next
            End If
            '�Ƽ۲���Ҫ���͵��ۼ�����
            rsTotal.Filter = "��ĿID=" & rsPrice!ID & " And �ⷿID=" & lng�ⷿID
            Do While Not rsTotal.EOF
                dbl�ѷ���� = dbl�ѷ���� + Format(rsTotal!���� / Nvl(rsPrice!סԺ��װ, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl���ÿ�� = Format(GetStock(rsPrice!ID, lng�ⷿID, 2), "0.00000")
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        strTmp = "ҽ��""" & .TextMatrix(lngRow, COL_ҽ������) & """�ķ�����ʱ�ۼƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                            vbCrLf & vbCrLf & Get��������(lng�ⷿID) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                            IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                    Else
                        strTmp = "ҽ��""" & .TextMatrix(lngRow, COL_ҽ������) & """�ļƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                            vbCrLf & vbCrLf & Get��������(lng�ⷿID) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                            IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҽ����"
                    End If
                    
                    .Redraw = flexRDDirect
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                End If
            End If
        End If
        
        '���δ��ʾ��Ҫ����,�����ۼƷ�������
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_���ID))
            Else
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
            End If
            rsTotal!��ĿID = rsPrice!ID
            rsTotal!�ⷿID = lng�ⷿID
            rsTotal!���� = dbl����
            rsTotal.Update
        End If
    End With
End Function

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�к� As Long, i As Long
    Dim str��ĿIDs As String, blnCancel As Boolean
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim vPoint As POINTAPI
    
    With vsPrice
        lng�к� = Val(.TextMatrix(Row, COLP_�к�))
        If Col = COLP_�շ���Ŀ Then
            '����ѡ�����е���Ŀ
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_�к�)) = lng�к� And lng�к� <> 0 And i <> Row Then
                    str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                End If
            Next
            str��ĿIDs = Mid(str��ĿIDs, 2)
            
            strSQL = _
                " Select Distinct 0 as ĩ��,To_Number('999999999'||����) as ID,-NULL as �ϼ�ID," & _
                " CHR(13)||���� as ����,Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',7,'��������') as ����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ˵��,NULL as �۸�," & _
                " -NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7)"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,-ID as ID,Nvl(-�ϼ�ID,To_Number('999999999'||����)) as �ϼ�ID,����,����," & _
                " NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ˵��,NULL as �۸�," & _
                " -NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From ���Ʒ���Ŀ¼ Where ���� in (1,2,3,7)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������," & _
                " NULL as ˵��,NULL as �۸�,-NULL as ԭ��ID,-NULL as �ּ�ID,-NULL as �Ƿ���ID,Null as ���ID,-Null as ��������ID" & _
                " From �շѷ���Ŀ¼ Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL " & _
                " Select ĩ��,ID,�ϼ�ID,����,����,��λ,���,����,���,��������,˵��," & _
                " Decode(Nvl(�Ƿ���,0),1,Decode(Instr('567',���ID),0,Sum(ԭ��)||'-'||Sum(�ּ�),'ʱ��'),Sum(�ּ�)) as �۸�," & _
                " Sum(ԭ��) as ԭ��ID,Sum(�ּ�) as �ּ�ID,�Ƿ��� as �Ƿ���ID,���ID,��������ID" & _
                " From (" & _
                " Select Distinct 1 as ĩ��,A.ID,Decode(Instr('567',A.���),0,A.����ID,-E.����ID) as �ϼ�ID,A.����,A.����," & _
                " A.���㵥λ as ��λ,A.���,A.����,C.���� as ���,A.��������,A.˵��,B.ԭ��,B.�ּ�,A.�Ƿ���," & _
                " A.��� as ���ID,-Null as ��������ID" & _
                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,������ĿĿ¼ E" & _
                " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.��� Not IN('4','J','1') And A.���=C.���� And A.ID=D.ҩƷID(+) And D.ҩ��ID=E.ID(+)"
            If DeptExist("���ϲ���", 2) Then
                strSQL = strSQL & " Union ALL " & _
                    " Select Distinct 1 as ĩ��,A.ID,-E.����ID as �ϼ�ID,A.����,A.����," & _
                    " A.���㵥λ as ��λ,A.���,A.����,C.���� as ���,A.��������,A.˵��," & _
                    " B.ԭ��,B.�ּ�,A.�Ƿ���,A.��� as ���ID,D.�������� as ��������ID" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ��� C,�������� D,������ĿĿ¼ E" & _
                    " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                    " And A.���='4' And A.���=C.���� And A.ID=D.����ID And D.����ID=E.ID"
            End If
            strSQL = strSQL & " ) Group by ĩ��,ID,�ϼ�ID,���,����,����,��λ,���,����,��������,˵��,�Ƿ���,���ID,��������ID"
            
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "�շ���Ŀ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str��ĿIDs & ",")
            If Not rsTmp Is Nothing Then
                '�Ǳ���ִ�е�ҽ����������������Ŀ
                If lng�к� <> 0 Then
                    If Nvl(rsTmp!�Ƿ���ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!���ID) > 0 Or rsTmp!���ID = "4" And Nvl(rsTmp!��������ID, 0) = 1) Then
                        If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                            MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ""" & rsTmp!���� & """���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
                            .SetFocus: Exit Sub
                        End If
                    End If
                End If
                
                'ҽ��������
                If CheckItemInsure(rsTmp) Then
                    .SetFocus: Exit Sub
                End If
                
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                Call SetItemInput(Row, rsTmp, lngҽ��ID, lngԭ��ĿID)
                If lng�к� <> 0 Then
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û�п��õ��շ���Ŀ�����ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_ִ�п��� Then
            vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_�շ����) = "4" Then
                '�������õ�����
                strSQL = _
                    " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                    " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                    " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                    " And B.������� IN(2,3) And B.����ID=C.ID" & _
                    " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And (A.������Դ is NULL Or A.������Դ=2)" & _
                    " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                    " And A.�շ�ϸĿID=[1]" & _
                    " Order by B.�������,C.����"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                'ҩƷ
                'ҩƷ��ϵͳָ���Ĵ���ҩ������
                If Not Check�ϰల��(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And A.�շ�ϸĿID=[1]" & _
                        " Order by B.�������,C.����"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                    Decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!����
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!ִ�п���ID = rsTmp!ID
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset) As Boolean
'���ܣ��������(ѡ��)�Ƽ���Ŀ�Ƿ�ҽ������
'���أ����δ���룬������ʾѡ�񲻼������򷵻��档
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int���� As Integer
    
    If gintҽ������ = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then int���� = Nvl(rsTmp!����, 0)
    If int���� <> 0 Then
        If Not ItemExistInsure(rsInput!ID, int����) Then
            If gintҽ������ = 1 Then
                If MsgBox("��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckItemInsure = True
                End If
            ElseIf gintҽ������ = 2 Then
                MsgBox "��Ŀ""" & rsInput!���� & """û�����ö�Ӧ�ı�����Ŀ��", vbInformation, gstrSysName
                CheckItemInsure = True
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditable(.Row, .Col) And .Col = COLP_�Ƽ�ҽ�� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_�̶�)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_�к�)) <> 0 And Val(.TextMatrix(.Row, COLP_�շ�ϸĿID)) <> 0 Then
                    'ҽ������д�������Ҫ����һ��(�����ǹ̶����ɶ���)
                    mrsPrice.Filter = "ҽ��ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_�к�)), COL_ID)) & " And ����=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_����) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_�Ƽ�ҽ��) & """����Ҫ����һ�������Ƽ���Ŀ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                
                    If MsgBox("ȷʵҪɾ����ǰ�Ƽ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "ҽ��ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_�к�)), COL_ID)) & " And �շ�ϸĿID=" & Val(.TextMatrix(.Row, COLP_�շ�ϸĿID))
                    mrsPrice.Delete
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_�Ƽ�ҽ��
                End If
                
                Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
            If CellEditable(.Row, .Col) And (.Col = COLP_�շ���Ŀ Or .Col = COLP_ִ�п���) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ���λ���۱�����һ����������ĵ�Ԫ��
    Dim i As Long, j As Long
    
    With vsPrice
        '��ǰ��Ԫ�����δ��������,���˳�
        If CellEditable(lngRow, lngCol) Then
            If lngCol = COLP_���� And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '����һ��Ԫ��ʼѭ������
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_�Ƽ�ҽ��) To .Cols - 1
                If CellEditable(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '��ǰ�����û���ҵ���һ���ɱ༭��Ԫ,�������Ƽ�ҽ��,������һ����
            If CStr(.ColData(COLP_�Ƽ�ҽ��)) <> "" Then
                '��ǰ��δ��������,��λ����������Ԫ
                If .TextMatrix(lngRow, COLP_�Ƽ�ҽ��) = "" Then
                    .Col = COLP_�Ƽ�ҽ��
                ElseIf .TextMatrix(lngRow, COLP_�Ƽ�����) = "" Then
                    .Col = COLP_�Ƽ�����
                ElseIf .TextMatrix(lngRow, COLP_�շ���Ŀ) = "" Then
                    .Col = COLP_�շ���Ŀ
                ElseIf Val(.TextMatrix(lngRow, COLP_���)) = 1 _
                    And Val(.TextMatrix(lngRow, COLP_����)) = 0 _
                    And CellEditable(lngRow, COLP_����) Then
                    .Col = COLP_����
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_�Ƽ�ҽ��
                    
                    'ȱʡѡ��Ƽ�ҽ��(�������)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '���ɱ༭ʱ���ⶨһ��
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng�к� As Long, i As Long
    Dim str��ĿIDs As String
    Dim lngҽ��ID As Long, lngԭ��ĿID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim StrInput As String, strMatch As String
    Dim vPoint As POINTAPI
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            lng�к� = Val(.TextMatrix(Row, COLP_�к�))
            If Col = COLP_�Ƽ�ҽ�� Then
                '����ʱ�س�
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '��ȻEnterNextCell����Ҫ�˳�
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_�Ƽ����� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�Ƽ�����������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_���� Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "�շѵ���������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '��������뷶Χ
                strTmp = CheckScope(.Cell(flexcpData, Row, COLP_Ӧ�ս��), .Cell(flexcpData, Row, COLP_ʵ�ս��), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, "0.00000")
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '���¼�¼��
                lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                    mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
                    mrsPrice!���� = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_�շ���Ŀ And .EditText <> "" Then
                '����ѡ�����е���Ŀ
                For i = .FixedRows To .Rows - 1
                    If Val(vsAdvice.TextMatrix(Val(.TextMatrix(i, COLP_�к�)), COL_ID)) = Val(vsAdvice.TextMatrix(lng�к�, COL_ID)) _
                        And Val(vsAdvice.TextMatrix(lng�к�, COL_ID)) <> 0 And i <> Row Then
                        str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COLP_�շ�ϸĿID))
                    End If
                Next
                str��ĿIDs = Mid(str��ĿIDs, 2)
                
                '��ͬ������ƥ�䷽ʽ
                StrInput = UCase(.EditText)
                strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=[3] Or C.���� Like [2] And C.���� IN([3],3))"
                If IsNumeric(StrInput) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.���� Like [1] And C.����=[3] Or C.���� Like [2] And C.����=3)"
                ElseIf zlCommFun.IsCharAlpha(StrInput) Then         '01,11.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.���� Like [2] And C.����=[3]"
                ElseIf zlCommFun.IsCharChinese(StrInput) Then
                    strMatch = " And C.���� Like [2] And C.����=[3]"
                End If
                
                strSQL = ""
                If Not DeptExist("���ϲ���", 2) Then strSQL = " And A.���<>'4'"
                strSQL = _
                    " Select A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����,A.��������,A.˵��," & _
                    " Decode(Nvl(A.�Ƿ���,0),1,Decode(Instr('567',A.���ID),0,Sum(A.ԭ��)||'-'||Sum(A.�ּ�),'ʱ��'),Sum(A.�ּ�)) as �۸�," & _
                    " Sum(A.ԭ��) as ԭ��ID,Sum(A.�ּ�) as �ּ�ID,A.�Ƿ��� as �Ƿ���ID,A.���ID,B.�������� as ��������ID" & _
                    " From (" & _
                    " Select Distinct 1 as ĩ��,A.ID,A.��� as ���ID,D.���� as ���,A.����,A.����,A.���㵥λ as ��λ," & _
                    " A.���,A.����,A.��������,A.˵��,B.ԭ��,B.�ּ�,A.�Ƿ���" & _
                    " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,�շ���Ŀ���� C,�շ���Ŀ��� D" & _
                    " Where A.ID=B.�շ�ϸĿID And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And A.������� IN(2,3)" & IIF(str��ĿIDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.�շ�ϸĿID And A.���=D.���� And A.��� Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,�������� B" & _
                    " Where A.ID=B.����ID(+)" & _
                    " Group by A.ĩ��,A.ID,A.���,A.����,A.����,A.��λ,A.���,A.����,A.��������,A.˵��,A.�Ƿ���,A.���ID,B.��������" & _
                    " Order by A.���,A.����"
                vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�շ���Ŀ", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    StrInput & "%", mstrLike & StrInput & "%", mint���� + 1, "," & str��ĿIDs & ",")
                If Not rsTmp Is Nothing Then
                    '�Ǳ���ִ�е�ҽ����������������Ŀ
                    If lng�к� <> 0 Then
                        If Nvl(rsTmp!�Ƿ���ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!���ID) > 0 Or rsTmp!���ID = "4" And Nvl(rsTmp!��������ID, 0) = 1) Then
                            If Not Check����ִ��(Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))) Then
                                MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ""" & rsTmp!���� & """���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                                .SetFocus: Exit Sub
                            End If
                        End If
                    End If
                
                    'ҽ��������
                    If CheckItemInsure(rsTmp) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                        .SetFocus: Exit Sub
                    End If
                    
                    lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    Call SetItemInput(Row, rsTmp, lngҽ��ID, lngԭ��ĿID)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    If lng�к� <> 0 Then
                        Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õ��շ���Ŀ��", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            ElseIf Col = COLP_ִ�п��� And .EditText <> "" Then 'ִ�п���
                vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_�շ����) = "4" Then
                    '�������õ�����
                    strSQL = _
                        " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                        " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
                        " And B.������� IN(2,3) And B.����ID=C.ID" & _
                        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                        " And (A.������Դ is NULL Or A.������Դ=2)" & _
                        " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                        " And A.�շ�ϸĿID=[1] And (C.���� Like [3] Or C.���� Like [4] Or C.���� Like [4])" & _
                        " Order by B.�������,C.����"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���ϲ���", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_�շ����)) > 0 Then
                    'ҩƷ��ϵͳָ���Ĵ���ҩ������
                    If Not Check�ϰల��(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(2,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And (A.������Դ is NULL Or A.������Դ=2)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.����,C.����,C.����,B.������� as ��ΧID" & _
                            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
                            " Where A.ִ�п���ID+0=B.����ID And B.��������=[3]" & _
                            " And B.������� IN(2,3) And B.����ID=C.ID" & _
                            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                            " And D.����ID=C.ID And D.����=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
                            " And (A.������Դ is NULL Or A.������Դ=2)" & _
                            " And (A.��������ID is NULL Or A.��������ID=[2])" & _
                            " And A.�շ�ϸĿID=[1] And (C.���� Like [4] Or C.���� Like [5] Or C.���� Like [5])" & _
                            " Order by B.�������,C.����"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩ��", False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_�շ�ϸĿID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���˿���ID)), _
                        Decode(.TextMatrix(Row, COLP_�շ����), "5", "��ҩ��", "6", "��ҩ��", "7", "��ҩ��"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_ִ�п���ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!����
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                    
                    '���¼�¼��
                    lngҽ��ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ID))
                    lngԭ��ĿID = Val(.TextMatrix(Row, COLP_�շ�ϸĿID))
                    If lngҽ��ID <> 0 And lngԭ��ĿID <> 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
                        mrsPrice!ִ�п���ID = rsTmp!ID
                        mrsPrice.Update
                        Call ShowAdvicePrice(vsAdvice.Row) '���¼�����ʾ
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_�Ƽ����� Or Col = COLP_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lngҽ��ID As Long, ByVal lngԭ��ĿID As Long)
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim lng�к� As Long, dbl���� As Double
    Dim blnHaveSub As Boolean
    
    With vsPrice
        '��¼������
        '�������:����ʱ��ʾ�����������Ŀ,Ҳ���Դ���Ϊδ���Ƽ�ҽ��������������Ŀ
        .TextMatrix(lngRow, COLP_���) = rsInput!���
        .TextMatrix(lngRow, COLP_�շ����) = rsInput!���ID
        .TextMatrix(lngRow, COLP_�շ�ϸĿID) = rsInput!ID
        .TextMatrix(lngRow, COLP_�շ���Ŀ) = rsInput!����
        If Not IsNull(rsInput!����) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & "(" & rsInput!���� & ")"
        End If
        If Not IsNull(rsInput!���) Then
            .TextMatrix(lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ) & " " & rsInput!���
        End If
        .TextMatrix(lngRow, COLP_��λ) = Nvl(rsInput!��λ) '�������۵�λ(������ҩ��ҩƷ�Ƽ�)
        .TextMatrix(lngRow, COLP_�Ƽ�����) = 1 'ȱʡ��ԼƼ�1,ҩƷΪ��1�����۵�λ
        
        'ִ�п���
        lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
        If lng�к� <> 0 Then
            lngִ�п���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_ִ�п���ID))
            '��ҩ��ҩƷ�͸������õ�����ר����ִ�п���
            If rsInput!���ID = "4" And Nvl(rsInput!��������ID, 0) = 1 Or InStr(",5,6,7,", rsInput!���ID) > 0 Then
                lng���˿���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���˿���ID))
                lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsInput!���ID, rsInput!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID)
            End If
        End If
        .TextMatrix(lngRow, COLP_ִ�п���) = Get��������(lngִ�п���ID)
        .TextMatrix(lngRow, COLP_ִ�п���ID) = lngִ�п���ID
        
        '���ۼ��㴦��:ҩ����ҩƷ�Ƽ۲����������ﴦ��
        If InStr(",5,6,7,", rsInput!���ID) > 0 Then
            If Nvl(rsInput!�Ƿ���ID, 0) = 0 Then
                dbl���� = Nvl(rsInput!�ּ�ID, 0)
            ElseIf lng�к� <> 0 Then
                '��ÿ��ȱʡһ�����۵�λ,��ǰ�������μ���
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, Val(vsAdvice.TextMatrix(lng�к�, COL_����)), , True)
            End If
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, "0.00000")
                        
            'ʱ��ҩƷ������۸�
            .TextMatrix(lngRow, COLP_���) = 0
            .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
            .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
        ElseIf rsInput!���ID = "4" And Nvl(rsInput!��������ID, 0) = 1 And Nvl(rsInput!�Ƿ���ID, 0) = 1 Then
            '�������õ�ʱ�����ĺ�ҩƷһ������
            dbl���� = 0
            If lng�к� <> 0 Then
                dbl���� = CalcDrugPrice(rsInput!ID, lngִ�п���ID, Val(vsAdvice.TextMatrix(lng�к�, COL_����)), , True)
            End If
            .TextMatrix(lngRow, COLP_���) = 0
            .TextMatrix(lngRow, COLP_����) = Format(dbl����, "0.00000")
            .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
            .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
        Else
            If Nvl(rsInput!�Ƿ���ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_���) = 0
                .TextMatrix(lngRow, COLP_����) = Format(Nvl(rsInput!�ּ�ID, 0), "0.00000")
                .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = 0
                .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = 0
            Else
                .TextMatrix(lngRow, COLP_���) = 1
                .TextMatrix(lngRow, COLP_����) = ""
                .Cell(flexcpData, lngRow, COLP_Ӧ�ս��) = Nvl(rsInput!ԭ��ID, 0)
                .Cell(flexcpData, lngRow, COLP_ʵ�ս��) = Nvl(rsInput!�ּ�ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_��������) = Nvl(rsInput!��������)
        .TextMatrix(lngRow, COLP_�̶�) = 0
        
        '��������ָ�
        .Cell(flexcpData, lngRow, COLP_�շ���Ŀ) = .TextMatrix(lngRow, COLP_�շ���Ŀ)
        .Cell(flexcpData, lngRow, COLP_�Ƽ�����) = .TextMatrix(lngRow, COLP_�Ƽ�����)
        .Cell(flexcpData, lngRow, COLP_����) = .TextMatrix(lngRow, COLP_����)
        .Cell(flexcpData, lngRow, COLP_ִ�п���) = .TextMatrix(lngRow, COLP_ִ�п���)
        
        '��¼������
        If lngҽ��ID <> 0 Then
            If lngԭ��ĿID = 0 Then
                '��ǰҽ���Ƿ��д��������������Ŀ�Ƿ����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And ����=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_����) = IIF(blnHaveSub, "��", "")
            
                mrsPrice.AddNew '����
            Else '����
                mrsPrice.Filter = "ҽ��ID=" & lngҽ��ID & " And �շ�ϸĿID=" & lngԭ��ĿID
            End If
            If lngԭ��ĿID = 0 Then
                mrsPrice!ҽ��ID = lngҽ��ID
                lng�к� = Val(.TextMatrix(lngRow, COLP_�к�))
                If Val(vsAdvice.TextMatrix(lng�к�, COL_���ID)) <> 0 Then
                    mrsPrice!���ID = Val(vsAdvice.TextMatrix(lng�к�, COL_���ID))
                Else
                    mrsPrice!���ID = Null
                End If
                mrsPrice!���� = IIF(blnHaveSub, 1, 0)
            End If
            mrsPrice!�շ���� = rsInput!���ID
            mrsPrice!�շ�ϸĿID = rsInput!ID
            If lngִ�п���ID <> 0 Then
                mrsPrice!ִ�п���ID = lngִ�п���ID
            Else
                mrsPrice!ִ�п���ID = Null
            End If
            mrsPrice!���� = Nvl(rsInput!��������ID, 0)
            mrsPrice!���� = 1
            mrsPrice!���� = Val(.TextMatrix(lngRow, COLP_����))
            mrsPrice!�̶� = 0
            mrsPrice.Update
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlCommFun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim bln�Ǳ��� As Boolean
    
    If Not CellEditable(Row, Col, bln�Ǳ���) Then
        '�Ǳ���ִ�еı����Ŀ�������۸�
        If bln�Ǳ��� Then
            MsgBox "��ҽ���Ǳ���ִ�У�������Ա����Ŀ���ۡ��üƼ���Ŀ��Ҫ�ֹ��Ƽۡ�", vbInformation, gstrSysName
        End If
        Cancel = True
    Else
        If Col = COLP_�Ƽ����� Or Col = COLP_���� Or Col = COLP_ִ�п��� Then
            '������ȷ���շ���Ŀ
            If vsPrice.TextMatrix(Row, COLP_�շ���Ŀ) = "" Then Cancel = True
        End If
        If Col = COLP_���� Then
            '������ǰ������ȷ���Ƽ�ҽ��,�Ծ����Ƿ��������(����ִ��)
            If vsPrice.TextMatrix(Row, COLP_�Ƽ�ҽ��) = "" Then Cancel = True
        End If
    End If
    
    If Col = COLP_�Ƽ����� Or Col = COLP_���� Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Sub InitBillSet()
'���ܣ���ʼ��ҽ�����ʵ������ɼ�¼��
    Set mrsBill = New ADODB.Recordset
    
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 8
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng������� As Long, lng������� As Long, bln������ As Boolean)
'���ܣ���ȡ��ǰ���õ��ݵ�NO�����
'������lng�������=���ü�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'      lng�������=���ͼ�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'˵����strKey=���ݼ��ʵ������ɹ��򶨵�Ψһ�ؼ���
'1.������ҩ��"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
'2.һ���䷽�е����в�ҩ����һ���������ݺ�
'3.����ҽ�����ҩ�ֺŹ�����ͬ��
'4.������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷�)
'5.��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
'6.һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        mrsBill!NO = zlDatabase.GetNextNO(IIF(bln������, 13, 14))
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill.Update
    Else
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        mrsBill.Update
    End If
    strNO = mrsBill!NO
    If lng������� <> -1 Then lng������� = mrsBill!�������
    If lng������� <> -1 Then lng������� = mrsBill!�������
End Sub

Private Sub DeleteSendRow()
'���ܣ���������ҽ���嵥���ѷ��ͳɹ��ĵ���ɾ��
    Dim i As Long, blnDel As Boolean
    
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) = -1 Then .RemoveItem i: blnDel = True
        Next
        .Redraw = flexRDDirect
        
        If blnDel Then
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: .Col = COL_ѡ��
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            
            vsPrice.Rows = vsPrice.FixedRows
            vsPrice.Rows = vsPrice.FixedRows + 1
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function Getʵ�ս��(ByVal strSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strSQL = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    strSQL = Left(strSQL, InStr(strSQL, strMatch) - 1)
    Getʵ�ս�� = CCur(strSQL)
End Function

Private Function Setʵ�ս��(ByVal strSQL As String, ByVal cur��� As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Setʵ�ս�� = strLeft & cur��� & strRight
End Function

Private Function CheckSignSend() As Boolean
'���ܣ����һ��ǩ����ҽ���Ƿ�һ���͵�
'˵��������ֻ����¿���ҽ������У�Ե�ҽ�����Ͳ�����Ӱ��(��ͬ������û��У��)
    Dim colǩ��ID As New Collection, strǩ��ID As String
    Dim lngǩ��ID As Long, strTmp As String
    Dim i As Long, j As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                '�ռ���ǩ��ҽ���ķ���״̬
                lngǩ��ID = Val(.TextMatrix(i, COL_ǩ��ID))
                If lngǩ��ID <> 0 Then
                    If InStr(strǩ��ID & ",", "," & lngǩ��ID & ",") > 0 Then
                        strTmp = Split(colǩ��ID("_" & lngǩ��ID), "=")(1)
                        If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                            If InStr(strTmp, "1") = 0 Then
                                colǩ��ID.Remove "_" & lngǩ��ID
                                colǩ��ID.Add lngǩ��ID & "=" & strTmp & "1", "_" & lngǩ��ID
                            End If
                        Else
                            If InStr(strTmp, "0") = 0 Then
                                colǩ��ID.Remove "_" & lngǩ��ID
                                colǩ��ID.Add lngǩ��ID & "=" & strTmp & "0", "_" & lngǩ��ID
                            End If
                        End If
                    Else
                        strǩ��ID = strǩ��ID & "," & lngǩ��ID
                        If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                            colǩ��ID.Add lngǩ��ID & "=1", "_" & lngǩ��ID
                        Else
                            colǩ��ID.Add lngǩ��ID & "=0", "_" & lngǩ��ID
                        End If
                    End If
                End If
            End If
        Next
            
        '���ǩ�����(һ��ǩ����ҽ������һ����)
        strTmp = ""
        For i = 1 To colǩ��ID.Count
            lngǩ��ID = Split(colǩ��ID(i), "=")(0)
            strǩ��ID = Split(colǩ��ID(i), "=")(1)
            If Not (strǩ��ID = "1" Or strǩ��ID = "0") Then
                '���ǩ�������ݲ���"��Ҫ���ͻ򶼲�����"�����
                j = .FindRow(CStr(lngǩ��ID), , COL_ǩ��ID)
                Do While j <> -1
                    If Not .RowHidden(j) Then
                        If .Cell(flexcpData, j, COL_ѡ��) = 1 Or .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                            strTmp = strTmp & vbCrLf & "��" & .TextMatrix(j, COL_ҽ������)
                        End If
                    End If
                    j = .FindRow(CStr(lngǩ��ID), j + 1, COL_ǩ��ID)
                Loop
                Exit For '��ֻ��ʾ��һ��
            End If
        Next
    End With
    
    If strTmp <> "" Then
        MsgBox "����ҽ������������Ҫ���͵�ҽ��һ��ǩ��������ǰ����Ϊ�����ͣ�" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "һ��ǩ����ҽ������һ���ͣ���������ҽ���ķ���״̬��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckSignSend = True
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng��ĿID As Long, ByVal lngCol As Long)
'���ܣ���λ������ʾָ��ҽ����ָ���Ƽ���
'������lngRow=ҽ���к�
'      lng��ĿID=�Ƽ���ĿID
'      lngCol=�Ƽ۱����ʾ��
    Dim k As Long
    
    With vsAdvice
        .Col = COL_ҽ������ '�������Զ�ShowPrice,mrsPrice�����仯
        If Not .RowHidden(lngRow) Then
            .Row = lngRow
        Else
            If InStr(",F,D,G,C,", .TextMatrix(lngRow, COL_�������)) > 0 And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                '��������,��������,��鲿λ,���������Ŀ
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), , COL_ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 1 Then
                '��ҩ;��
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_���ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 2 Then
                '��ҩ�巨
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_���ID))), lngRow + 1, COL_ID)
            End If
        End If
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_�к�)) = lngRow _
                And Val(vsPrice.TextMatrix(k, COLP_�շ�ϸĿID)) = lng��ĿID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function SendAdvice(ByVal bln������ As Boolean) As Long
'���ܣ�����ҽ������(��������м��ʱ���)
'˵����������˷����ύ
'���أ�����ɹ��򷵻ط��ͺ�
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsUpload As ADODB.Recordset
    Dim rsMoney As New ADODB.Recordset

    Dim i As Long, j As Long
    Dim strSQL As String, curDate As Date
    Dim blnTran As Boolean, blnBool As Boolean, strTmp As String
    Dim strWarn As String, intWarn As Integer, str��� As String, str������� As String
    
    Dim lng���ͺ� As Long, int�Ʒ�״̬ As Integer, int���� As Integer, strNO As String
    Dim lngϸĿID As Long, lng������� As Long, lng���ø��� As Long, lng������� As Long
    Dim int���� As Integer, dbl���� As Double, cur�ϼ� As Currency
    Dim dbl���� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim str�ֽ�ʱ�� As String, str�״�ʱ�� As String, strĩ��ʱ�� As String
    Dim int�䷽�� As Integer, strNOKey As String, str�Զ����� As String
    Dim str����ʱ�� As String, str�Ǽ�ʱ�� As String
    Dim dbl�������� As Double, blnFirst As Boolean '�䷽�����ֺŹؼ���
    Dim lngҩƷ���ID As Long, lng�������ID As Long
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim bln��Ժ��ҩ As Boolean, blnVarZero As Boolean
    Dim bln�������� As Boolean, blnHaveSub As Boolean
    Dim int����� As Integer, var������ As Variant
    Dim lng������ID As Long, strʵ�� As String
    Dim curҽ���ϼ� As Currency
    
    Dim blnҩƷʱ����ʾ As Boolean, blnҩƷ�����ʾ As Boolean, blnҩƷĬ�Ϸ��� As Boolean
    Dim bln����ʱ����ʾ As Boolean, bln���Ŀ����ʾ As Boolean, bln����Ĭ�Ϸ��� As Boolean
    Dim bln������Ŀ�� As Boolean, lng���մ���ID As Long, curͳ���� As Currency, str���ձ��� As String, str�������� As String
    
    Dim rsAudit As ADODB.Recordset
    Dim strAudit As String
    
    '����ǩ��
    Dim lng��ID As Long, strҽ��IDs As String, strSource As String
    Dim intRule As Integer, strSign As String
    Dim lng֤��ID As Long, lngǩ��ID As Long
    
    On Error GoTo errH
    
    '���һ��ǩ����ҽ���Ƿ�һ����
    If Not CheckSignSend Then Exit Function
    
    With vsAdvice
        '�ȼ�鲢��ʾ����ҽ��:3-ת��,5-��Ժ,6-תԺ,11-����
        strTmp = ""
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                If .TextMatrix(i, COL_�������) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                    strTmp = strTmp & vbCrLf & mrsPati!���� & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & "��" & .TextMatrix(i, COL_ҽ������)
                End If
            End If
        Next
        If strTmp <> "" Then
            If MsgBox("Ҫ���͵�ҽ���а�����������ҽ����" & vbCrLf & strTmp & vbCrLf & vbCrLf & "ȷʵҪ���͵�ǰѡ���ҽ����", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End With
    
    '��ȡ��ǰ���˵�������Ŀ�嵥
    strAudit = ""
    If Not IsNull(mrsPati!����) Then
        Set rsAudit = GetAuditRecord(mlng����ID, mlng��ҳID)
    Else
        Set rsAudit = Nothing '��NothingΪ��־�ò��˲���Ҫ�ж�
    End If
    
    '��ȡҩƷ/����������
    lngҩƷ���ID = ExistIOClass(IIF(bln������, 8, 9))
    lng�������ID = ExistIOClass(IIF(bln������, 40, 41))
    
    Screen.MousePointer = 11
    
    blnҩƷʱ����ʾ = True: blnҩƷ�����ʾ = True: blnҩƷĬ�Ϸ��� = True
    bln����ʱ����ʾ = True: bln���Ŀ����ʾ = True: bln����Ĭ�Ϸ��� = True
    
    Call InitBillSet
    Call InitRecordSet(rsSQL, rsTotal, rsUpload)
    lng���ͺ� = zlDatabase.GetNextNO(10)
    
    '���ʱ�䷢�͹�����δ����ֹͣʱ��,Ϊ������У��ʱ���ظ�(ȡ��Sysdate)
    curDate = zlDatabase.Currentdate
    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ,�벡���޹�
    int�䷽�� = 1 '��ʾ���͵ĵڼ����䷽,���ڷֵ��ݺ�
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '����ҽ����3-ת��;5-��Ժ;6-תԺ,11-����
                If .TextMatrix(i, COL_�������) = "Z" Then
                    'ת��,��Ժ,תԺ,����ҽ������ʱ������Ҫ��������״̬
                    If .Cell(flexcpData, i, COL_Ӥ��) = 0 Then
                        If InStr(",3,5,6,11,", .TextMatrix(i, COL_��������)) > 0 And Nvl(mrsPati!״̬, 0) <> 0 Then
                            MsgBox "����""" & mrsPati!���� & """��ǰ����""" & Decode(Nvl(mrsPati!״̬, 0), 1, "�ȴ����", 2, "����ת��", 3, "��Ԥ��Ժ") & """״̬��" & _
                                "���ܷ���""" & .TextMatrix(i, COL_ҽ������) & """ҽ����", vbInformation, gstrSysName
                            Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                            GoTo NextAdvice
                        End If
                    End If
                    
                    '�����ת�ơ���Ժ��תԺҽ��,��鲡���Ƿ���δִ�е�ҽ����Ŀ��δ��ҩƷ
                    If InStr(",3,5,6,", .TextMatrix(i, COL_��������)) > 0 Then
                        strTmp = ExistWaitExe(mlng����ID, mlng��ҳID, .Cell(flexcpData, i, COL_Ӥ��))
                        If strTmp <> "" Then
                            Call .ShowCell(i, COL_ҽ������): .Refresh
                            If MsgBox("���ֲ���""" & mrsPati!���� & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & """������δִ����ɵ����ݣ�" & _
                                vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "ȷʵҪ����""" & .TextMatrix(i, COL_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                                GoTo NextAdvice
                            End If
                        End If
                        strTmp = ExistWaitDrug(mlng����ID, mlng��ҳID, .Cell(flexcpData, i, COL_Ӥ��))
                        If strTmp <> "" Then
                            Call .ShowCell(i, COL_ҽ������): .Refresh
                            If MsgBox("���ֲ���""" & mrsPati!���� & IIF(.Cell(flexcpData, i, COL_Ӥ��) <> 0, "(Ӥ��" & .Cell(flexcpData, i, COL_Ӥ��) & ")", "") & """" & _
                                strTmp & vbCrLf & vbCrLf & "ȷʵҪ����""" & .TextMatrix(i, COL_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                                GoTo NextAdvice
                            End If
                        End If
                    End If
                End If
            
                '�������ݺŷ���ؼ���
                '-----------------------------------------------------------------------------------------
                If InStr(",5,6,M,", .TextMatrix(i, COL_�������)) > 0 Then
                    '������ҩ�Ͳ��ϰ�"����(����ID,�Һŵ�)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                    strNOKey = "��ҩ����_" & mlng����ID & "_" & mlng��ҳID & "_" & _
                        Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                        .TextMatrix(i, COL_����ҽ��) & "_" & Val(.TextMatrix(i, COL_ִ�п���ID))
                    '�ٰ�Ҫ��ӡ�����Ƶ��ݷֺ�
                    strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), 2)
                ElseIf .TextMatrix(i, COL_�������) = "7" Then
                    'һ���䷽�е����в�ҩ����һ���������ݺ�
                    strNOKey = "��ҩ�䷽_" & mlng����ID & "_" & mlng��ҳID & "_" & int�䷽��
                ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And .TextMatrix(i, COL_�������) = "C" Then
                    'һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
                    strNOKey = "һ���ɼ�_" & Val(.TextMatrix(i, COL_���ID))
                ElseIf Val(.TextMatrix(i, COL_���ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_�������)) > 0 Then
                    '��鲿λ�͸�����������Ҫҽ��������ͬ���ݺţ�����������䵥���ĵ��ݺš�
                    strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_���ID))
                Else
                    '������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷�)
                    strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_ID))
                End If
                
                '�Ƿ���Ժ��ҩ
                bln��Ժ��ҩ = False
                If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                    If .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ" Then bln��Ժ��ҩ = True
                ElseIf .TextMatrix(i, COL_�������) = "7" Then
                    j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                    If j <> -1 Then
                        If .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ" Then bln��Ժ��ҩ = True
                    End If
                End If
                
                '����ҽ�����ʷ���:�����¼۸����
                '-----------------------------------------------------------------------------------------
                strSQL = "": lngϸĿID = 0
                If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                    'ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
                    If Val(.TextMatrix(i, COL_ִ������ID)) <> 5 Then
                        strSQL = _
                            " Select A.ID,A.���,D.���� as �������,RTrim(A.����||' '||A.���) as ����," & _
                            " A.���㵥λ,A.�Ƿ���,A.���ηѱ�,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,100 as �����շ���," & _
                            " Y.סԺ��λ,Y.סԺ��װ,Y.����ϵ��,Y.ҩ������ as ����,0 as ��������,B.������ĿID," & _
                            " C.�վݷ�Ŀ,1 as ����,B.�ּ� as ����,[2] as ִ�п���ID,0 as ����,I.Ҫ������" & _
                            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,ҩƷ��� Y,����֧����Ŀ I" & _
                            " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.���=D.����" & _
                            " And A.ID=Y.ҩƷID(+) And A.ID=[1] And A.ID=I.�շ�ϸĿID(+) And I.����(+)=[3]" & _
                            " And ((Sysdate Between B.ִ������ and B.��ֹ����) Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                            " Order by A.����"
                    End If
                Else
                    '��ɾ��ԭ��ҩҽ���ļƼ�(Ӧ��û��)
                    rsSQL.AddNew
                    rsSQL!���� = 1: rsSQL!��ĿID = 0: rsSQL!��� = i
                    rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    rsSQL!SQL = "ZL_����ҽ���Ƽ�_Delete(" & Val(.TextMatrix(i, COL_ID)) & ",1)"
                    rsSQL.Update
                    
                    '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�е�ҽ������ȡ
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        mrsPrice.Filter = "ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                blnVarZero = False '�Ƿ����Ҽ۸�Ϊ��
                                If Nvl(mrsPrice!����, 0) = 0 Then
                                    blnVarZero = ItemIsVarPrice(mrsPrice!�շ�ϸĿID)
                                End If
                                If Not blnVarZero Then
                                    If Nvl(mrsPrice!����, 0) <> 0 Then '��������Ϊ0���Զ����˵�
                                        '��д����ҽ���Ƽ�:ֻ�����ҩ��ҩƷ���������ļƼ�
                                        If InStr(",5,6,7,", mrsPrice!�շ����) > 0 _
                                            Or mrsPrice!�շ���� = "4" And Nvl(mrsPrice!����, 0) = 1 Then
                                            lngִ�п���ID = Nvl(mrsPrice!ִ�п���ID, 0)
                                            
                                            '���ı�������ִ�п���
                                            If lngִ�п���ID = 0 And mrsPrice!�շ���� = "4" Then
                                                Call SeekPriceRow(i, mrsPrice!�շ�ϸĿID, COLP_ִ�п���)
                                                Screen.MousePointer = 0
                                                MsgBox "����""" & vsPrice.TextMatrix(vsPrice.Row, COLP_�շ���Ŀ) & """û��ȷ��ִ�п��ң����ֹ�������ȷ��ִ�п��ҡ�" & vbCrLf & _
                                                    "�������ȷ����ȷ��ִ�п��ң��뵽""����Ŀ¼����""�м��洢�ⷿ�����Ƿ���ȷ��", vbInformation, gstrSysName
                                                vsPrice.SetFocus: GoTo FuncEnd
                                            End If
                                        Else
                                            lngִ�п���ID = 0
                                        End If
                                        rsSQL.AddNew
                                        rsSQL!���� = 1: rsSQL!��ĿID = mrsPrice!�շ�ϸĿID: rsSQL!��� = i
                                        rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                        rsSQL!SQL = "ZL_����ҽ���Ƽ�_INSERT(" & _
                                            mrsPrice!ҽ��ID & "," & mrsPrice!�շ�ϸĿID & "," & _
                                            Nvl(mrsPrice!����, 0) & "," & Nvl(mrsPrice!����, 0) & "," & _
                                            Nvl(mrsPrice!����, 0) & "," & ZVal(lngִ�п���ID) & ")"
                                        rsSQL.Update
                                        
                                        '��ʱ����ҽ���Ƽ۱�
                                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                                            "Select " & mrsPrice!�շ�ϸĿID & " as �շ�ϸĿID," & _
                                            Nvl(mrsPrice!ִ�п���ID, 0) & " as ִ�п���ID," & _
                                            Nvl(mrsPrice!����, 0) & " as ����," & _
                                            Format(Nvl(mrsPrice!����, 0), "0.00000") & " as ����," & _
                                            Nvl(mrsPrice!����, 0) & " as ���� From Dual"
                                    End If
                                Else 'If Check����ִ��(Val(.TextMatrix(i, COL_ִ�п���ID))) Then
                                    '����Ϊ��,�����Ǳ��δ����(�����۸񲻿���Ϊ0)
                                    '����ִ�е���Ҫ����,�������ֹ��Ƽ�
                                    Call SeekPriceRow(i, mrsPrice!�շ�ϸĿID, COLP_����)
                                    Screen.MousePointer = 0
                                    MsgBox "����Ϊ��۵��շ���Ŀȷ��һ���շѼ۸�", vbInformation, gstrSysName
                                    vsPrice.SetFocus: GoTo FuncEnd
                                End If
                                mrsPrice.MoveNext
                            Next
                        End If
                    End If
                    
                    If strSQL <> "" Then
                        strSQL = _
                            " Select A.ID,A.���,D.���� as �������,A.����,A.���㵥λ,A.�Ƿ���," & _
                            " A.���ηѱ�,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,B.�����շ���,Y.סԺ��λ,Y.סԺ��װ,Y.����ϵ��," & _
                            " Decode(A.���,'4',E.���÷���,Y.ҩ������) as ����,E.��������,B.������ĿID," & _
                            " C.�վݷ�Ŀ,X.����,Decode(A.�Ƿ���,1,X.����,B.�ּ�) as ����,X.ִ�п���ID,X.����,I.Ҫ������" & _
                            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,�������� E,(" & strSQL & ") X,ҩƷ��� Y,����֧����Ŀ I" & _
                            " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.ID=E.����ID(+)" & _
                            " And A.���=D.���� And X.�շ�ϸĿID=A.ID And A.ID=Y.ҩƷID(+)" & _
                            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                            " And A.ID=I.�շ�ϸĿID(+) And I.����(+)=[3]" & _
                            " Order by X.����,A.ID"
                            'һ��Ҫ����������ǰ��,�Ա��ڼ�����ڷ��ü�¼�б������ӹ�ϵ
                    End If
                End If
                                
                '�����ۿ۱�����ʼ
                blnHaveSub = False
                var������ = Empty: int����� = 0
                curҽ���ϼ� = 0: lng������ID = 0
                
                int�Ʒ�״̬ = IIF(Val(.TextMatrix(i, COL_�Ƽ�����)) = 1, -1, 0) '����Ʒѻ�δ�Ʒ�
                If strSQL <> "" Then
                    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_ҩƷID)), Val(.TextMatrix(i, COL_ִ�п���ID)), Val(Nvl(mrsPati!����, 0)))
                    If Not rsMoney.EOF Then
                        int�Ʒ�״̬ = 1 '�ѼƷ�
                                                
                        'ȷ���Ƿ����ӹ�ϵ:��ʹ�������ۿ�,ҲҪ��¼
                        rsMoney.Filter = "����=1"
                        If Not rsMoney.EOF Then blnHaveSub = True
                        rsMoney.Filter = 0
                    End If
                    
                    '����������Ŀ���ķ�����ϸ
                    bln�������� = .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) <> 0
                    For j = 1 To rsMoney.RecordCount
                        '����Ƿ���Ҫ���Ѿ�����
                        If Nvl(rsMoney!Ҫ������, 0) = 1 And Not rsAudit Is Nothing Then
                            rsAudit.Filter = "��ĿID=" & rsMoney!ID
                            If rsAudit.EOF Then
                                If UBound(Split(strAudit, vbCrLf)) < 10 Then
                                    If InStr(strAudit, "��" & rsMoney!����) = 0 Then
                                        strAudit = strAudit & vbCrLf & "��" & rsMoney!����
                                    End If
                                ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                                    strAudit = strAudit & vbCrLf & "�� ��"
                                End If
                            End If
                        End If
                    
                        'ִ�п���ID
                        lngִ�п���ID = Nvl(rsMoney!ִ�п���ID, 0)
                        '��ԭֵ������ȡ��Ч�ķ�ҩ��ҩƷ���������ĵ�ִ�п���
                        If rsMoney!��� = "4" And Nvl(rsMoney!��������, 0) = 1 _
                            Or InStr(",5,6,7", rsMoney!���) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_�������)) = 0 Then
                            lng���˿���ID = Val(.TextMatrix(i, COL_���˿���ID))
                            lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, 0, rsMoney!���, rsMoney!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID)
                        End If
                        If InStr(",5,6,7", rsMoney!���) > 0 Then
                            If lngҩƷ���ID = 0 Then
                                MsgBox "����ȷ��ҩƷ�������ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                        
                            If InStr(",5,6,7", .TextMatrix(i, COL_�������)) > 0 Then
                                If .TextMatrix(i, COL_�������) = "7" Then
                                    int���� = Val(.TextMatrix(i, COL_����))
                                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                    If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                        dbl���� = Val(.TextMatrix(i, COL_����)) / Nvl(rsMoney!����ϵ��, 1)
                                    Else
                                        dbl���� = IntEx(Val(.TextMatrix(i, COL_����)) / Nvl(rsMoney!����ϵ��, 1) / Nvl(rsMoney!סԺ��װ, 1)) * Nvl(rsMoney!סԺ��װ, 1)
                                    End If
                                Else
                                    int���� = 1
                                    dbl���� = Val(.TextMatrix(i, COL_����)) * Nvl(rsMoney!סԺ��װ, 1)
                                End If
                            Else
                                int���� = 1
                                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,��˲��������㴦��
                                dbl���� = Val(.TextMatrix(i, COL_����)) * Nvl(rsMoney!����, 0)
                            End If
                            dbl���� = Format(dbl����, "0.00000")
                            
                            If Nvl(rsMoney!�Ƿ���, 0) = 1 Then
                                dbl���� = Format(CalcDrugPrice(rsMoney!ID, lngִ�п���ID, int���� * dbl����, , True), "0.00000")
                            Else
                                dbl���� = Format(Nvl(rsMoney!����, 0), "0.00000")
                            End If
                        ElseIf rsMoney!��� = "4" And Nvl(rsMoney!��������, 0) = 1 Then
                            '�����������������
                            If lng�������ID = 0 Then
                                Screen.MousePointer = 0
                                MsgBox "����ȷ���������ϵ��ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                            
                            int���� = 1
                            dbl���� = Format(Val(.TextMatrix(i, COL_����)) * Nvl(rsMoney!����, 0), "0.00000")
                            
                            'ȷ��ʱ�����ļ۸�
                            If Nvl(rsMoney!�Ƿ���, 0) = 1 Then
                                dbl���� = Format(CalcDrugPrice(rsMoney!ID, lngִ�п���ID, dbl����, , True), "0.00000")
                            Else
                                dbl���� = Format(Nvl(rsMoney!����, 0), "0.00000")
                            End If
                        Else
                            int���� = 1
                            dbl���� = Format(Val(.TextMatrix(i, COL_����)) * Nvl(rsMoney!����, 0), "0.00000")
                            dbl���� = Format(Nvl(rsMoney!����, 0), "0.00000")
                        End If
                        
                        '��ҩ��ҩƷ���������ĵĿ����
                        If rsMoney!��� = "4" And Nvl(rsMoney!��������, 0) = 1 _
                            Or InStr(",5,6,7", rsMoney!���) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_�������)) = 0 Then
                            If GetStockCheck(lngִ�п���ID) <> 0 Or Nvl(rsMoney!�Ƿ���, 0) = 1 Or Nvl(rsMoney!����, 0) = 1 Then
                                If rsMoney!��� = "4" Then
                                    blnBool = CheckPriceStock(i, rsMoney, lngִ�п���ID, int���� * dbl����, rsTotal, bln���Ŀ����ʾ, bln����ʱ����ʾ, bln����Ĭ�Ϸ���)
                                Else
                                    blnBool = CheckPriceStock(i, rsMoney, lngִ�п���ID, int���� * dbl����, rsTotal, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���)
                                End If
                                If blnBool Then
                                    Call RowSelectSame(i, COL_ѡ��, rsSQL, rsTotal, rsUpload, strҽ��IDs)
                                    GoTo NextAdvice
                                End If
                            End If
                        End If
                            
                        '���ͽ��
                        curӦ�� = int���� * dbl���� * dbl����
                        If bln�������� Then
                            curӦ�� = curӦ�� * Nvl(rsMoney!�����շ���, 100) / 100
                        End If
                        
                        '����Ӱ�Ӽ�
                        If gbln�Ӱ�Ӽ� And Nvl(rsMoney!�Ӱ�Ӽ�, 0) = 1 Then
                            curӦ�� = curӦ�� * (1 + Nvl(rsMoney!�Ӱ�Ӽ���, 0) / 100)
                        End If
                        
                        curӦ�� = Format(curӦ��, gstrDec)
                        
                        '��������ۿۺϼ�
                        If gbln��������ۿ� And blnHaveSub Then
                            curʵ�� = curӦ��
                            curҽ���ϼ� = curҽ���ϼ� + curʵ��
                        ElseIf Nvl(rsMoney!���ηѱ�, 0) = 0 Then
                            curʵ�� = Format(ActualMoney(Nvl(mrsPati!�ѱ�), rsMoney!������ĿID, curӦ��, rsMoney!ID, lngִ�п���ID, int���� * dbl����, _
                                IIF(gbln�Ӱ�Ӽ� And Nvl(rsMoney!�Ӱ�Ӽ�, 0) = 1, Nvl(rsMoney!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                        Else
                            curʵ�� = curӦ��
                        End If
                        
                        'ҽ������ֶ�
                        bln������Ŀ�� = False: lng���մ���ID = 0: curͳ���� = 0: str���ձ��� = "": str�������� = ""
                        If Not IsNull(mrsPati!����) Then
                            strTmp = gclsInsure.GetItemInsure(mlng����ID, rsMoney!ID, curʵ��, False, mrsPati!����)
                            If strTmp <> "" Then
                                bln������Ŀ�� = Val(Split(strTmp, ";")(0)) <> 0
                                lng���մ���ID = Val(Split(strTmp, ";")(1))
                                curͳ���� = Format(Val(Split(strTmp, ";")(2)), gstrDec)
                                str���ձ��� = CStr(Split(strTmp, ";")(3))
                                If UBound(Split(strTmp, ";")) >= 5 Then
                                    If Split(strTmp, ";")(5) <> "" Then
                                        str�������� = Split(strTmp, ";")(5)
                                    End If
                                End If
                            End If
                        End If
                        
                        '�ռ����ʱ������
                        cur�ϼ� = cur�ϼ� + curʵ��
                        If InStr(str���, rsMoney!���) = 0 Then
                            str��� = str��� & rsMoney!���
                            str������� = str������� & "," & rsMoney!�������
                        End If
                        
                        'NO,���
                        Call GetCurBillSet(strNOKey, strNO, lng�������, -1, bln������)
                        rsSQL.AddNew: blnBool = False
                        If rsMoney!ID <> lngϸĿID Then
                            lng���ø��� = lng�������
                            '���ӹ�ϵʱ����¼������Ϣ
                            If rsMoney!���� = 0 And blnHaveSub Then
                                int����� = lng�������
                                lng������ID = rsMoney!������ĿID
                                var������ = rsSQL.Bookmark
                                blnBool = True
                            End If
                        End If
                        lngϸĿID = rsMoney!ID
                        
                        '�����ۿ�ʱ���������ʵ�ս�������⴦��
                        If gbln��������ۿ� And blnHaveSub And blnBool Then
                            strʵ�� = Chr(0) & Chr(1) & "Begin" & curʵ�� & "End" & Chr(0) & Chr(1)
                        Else
                            strʵ�� = curʵ��
                        End If
                        
                        '����ʱ��
                        If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                            str����ʱ�� = "To_Date('" & Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            str����ʱ�� = "To_Date('" & .Cell(flexcpData, i, COL_�ֽ�ʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                        
                        '��Ϊ���ڲ��Ƽ۵�ҽ������������,���Դ���ļƼ����Զ�Ϊ(0-�����Ƽ�)
                        rsSQL!���� = 5: rsSQL!��ĿID = rsMoney!ID: rsSQL!��� = i
                        rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                        If bln������ Then
                            '��δȡ��ҩ����
                            rsSQL!SQL = "ZL_���ﻮ�ۼ�¼_INSERT(" & _
                                "'" & strNO & "'," & lng������� & "," & mlng����ID & "," & ZVal(mlng��ҳID) & "," & _
                                ZVal(Nvl(mrsPati!סԺ��, 0)) & ",'" & Nvl(mrsPati!����) & "','" & Nvl(mrsPati!����) & "'," & _
                                "'" & Nvl(mrsPati!�Ա�) & "','" & Nvl(mrsPati!����) & "'," & _
                                "'" & Nvl(mrsPati!�ѱ�) & "',NULL," & ZVal(Nvl(mrsPati!��ǰ����ID, 0)) & "," & _
                                ZVal(.TextMatrix(i, COL_���˿���ID)) & "," & ZVal(.TextMatrix(i, COL_��������ID)) & "," & _
                                "'" & .TextMatrix(i, COL_����ҽ��) & "'," & IIF(rsMoney!���� = 1, ZVal(int�����), "NULL") & "," & _
                                rsMoney!ID & ",'" & rsMoney!��� & "','" & Nvl(rsMoney!���㵥λ) & "',NULL," & _
                                int���� & "," & dbl���� & "," & IIF(bln��������, 1, 0) & "," & ZVal(lngִ�п���ID) & "," & _
                                IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsMoney!������ĿID & "," & _
                                "'" & Nvl(rsMoney!�վݷ�Ŀ) & "'," & dbl���� & "," & curӦ�� & "," & strʵ�� & "," & _
                                str����ʱ�� & ",To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "'ҽ������','" & UserInfo.���� & "'," & IIF(rsMoney!��� = "4", lng�������ID, lngҩƷ���ID) & "," & _
                                "'" & .TextMatrix(i, COL_ҽ������) & "'," & Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                                ZVal(.TextMatrix(i, COL_����)) & ",'" & .TextMatrix(i, COL_�÷�) & "',1," & _
                                IIF(bln��Ժ��ҩ, 3, Val(.TextMatrix(i, COL_�Ƽ�����))) & ",2)"
                        Else
                            '�Ƿ񻮼۷���
                            If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                                int���� = IIF(InStr(gstr���ͻ��۵�, "5") > 0, 1, 0)
                            Else
                                int���� = IIF(InStr(gstr���ͻ��۵�, .TextMatrix(i, COL_�������)) > 0, 1, 0)
                            End If
                            
                            '�Ǽ�ʱ��
                            If int���� = 1 Then '��ǻ��۵�ʱ�������ֿ�
                                str�Ǽ�ʱ�� = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str�Ǽ�ʱ�� = "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            
                            '�ռ�ҽ���ϴ����ݺ�:mrsBill�еĲ�һ�������˷���
                            If int���� = 0 Then
                                rsUpload.Filter = "NO='" & strNO & "'"
                                If rsUpload.EOF Then
                                    rsUpload.AddNew
                                    rsUpload!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                    rsUpload!NO = strNO
                                    rsUpload.Update
                                End If
                            End If
                            
                            rsSQL!SQL = "ZL_סԺ���ʼ�¼_Insert(" & _
                                "'" & strNO & "'," & lng������� & "," & mlng����ID & "," & ZVal(mlng��ҳID) & "," & _
                                ZVal(Nvl(mrsPati!סԺ��)) & ",'" & Nvl(mrsPati!����) & "'," & _
                                "'" & Nvl(mrsPati!�Ա�) & "','" & Nvl(mrsPati!����) & "'," & _
                                "'" & Nvl(mrsPati!����) & "','" & Nvl(mrsPati!�ѱ�) & "'," & _
                                ZVal(Nvl(mrsPati!��ǰ����ID, 0)) & "," & ZVal(.TextMatrix(i, COL_���˿���ID)) & ",0," & _
                                Val(.Cell(flexcpData, i, COL_Ӥ��)) & "," & _
                                ZVal(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                                IIF(rsMoney!���� = 1, ZVal(int�����), "NULL") & "," & rsMoney!ID & "," & _
                                "'" & rsMoney!��� & "','" & Nvl(rsMoney!���㵥λ) & "'," & _
                                IIF(bln������Ŀ��, 1, 0) & "," & ZVal(lng���մ���ID) & ",'" & str���ձ��� & "'," & _
                                int���� & "," & dbl���� & ",NULL," & ZVal(lngִ�п���ID) & "," & _
                                IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsMoney!������ĿID & "," & _
                                "'" & Nvl(rsMoney!�վݷ�Ŀ) & "'," & dbl���� & "," & curӦ�� & "," & strʵ�� & "," & _
                                curͳ���� & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                                "'ҽ������'," & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "',0," & _
                                IIF(rsMoney!��� = "4", lng�������ID, lngҩƷ���ID) & "," & _
                                "NULL,'" & .TextMatrix(i, COL_ҽ������) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                                "'" & .TextMatrix(i, COL_Ƶ��) & "'," & ZVal(.TextMatrix(i, COL_����)) & "," & _
                                "'" & .TextMatrix(i, COL_�÷�) & "',1," & _
                                IIF(bln��Ժ��ҩ, 3, Val(.TextMatrix(i, COL_�Ƽ�����))) & ",Null,'" & str�������� & "')"
                        End If
                        rsSQL.Update
                        
                        '��¼�Զ����ϵ�SQL
                        If gblnסԺ�Զ����� And Not bln������ And int���� = 0 And lngִ�п���ID <> 0 And rsMoney!��� = "4" And Nvl(rsMoney!��������, 0) = 1 Then
                            If InStr(str�Զ����� & ";", ";" & strNO & "," & lngִ�п���ID & ";") = 0 Then
                                rsSQL.AddNew
                                rsSQL!���� = 6
                                rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                rsSQL!��ĿID = 0
                                rsSQL!��� = i
                                rsSQL!SQL = "zl_�����շ���¼_��������(" & lngִ�п���ID & ",25,'" & strNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
                                rsSQL.Update
                                str�Զ����� = str�Զ����� & ";" & strNO & "," & lngִ�п���ID
                            End If
                        End If
                        
                        rsMoney.MoveNext
                    Next
                End If
                
                '��ҽ�������л����ۿ۴���
                If gbln��������ۿ� And blnHaveSub And var������ <> Empty And lng������ID <> 0 Then
                    rsSQL.Bookmark = var������
                    curʵ�� = Format(ActualMoney(Nvl(mrsPati!�ѱ�), lng������ID, curҽ���ϼ�), gstrDec)
                    curʵ�� = curʵ�� - curҽ���ϼ� '���۲��
                    curʵ�� = Getʵ�ս��(rsSQL!SQL) + curʵ��
                    rsSQL!SQL = Setʵ�ս��(rsSQL!SQL, curʵ��)
                    rsSQL.Update
                End If
                
                'ѡ��Ҫ���͵�ҽ���Զ�����У��(ʵ�ʿ�����Ϊ����������)
                If Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    rsSQL.AddNew
                    rsSQL!���� = 3: rsSQL!��ĿID = 0: rsSQL!��� = i
                    rsSQL!ҽ��ID = Val(.TextMatrix(i, COL_ID))
                    rsSQL!SQL = "ZL_����ҽ����¼_У��(" & Val(.TextMatrix(i, COL_ID)) & ",3," & _
                        "To_Date('" & Format(.TextMatrix(i, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),0)"
                End If
                
                
                '����ҽ�����ͼ�¼
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_ִ������ID)) <> 0 Then '����������(��ҩ;�����䷽�巨���÷�,�ɼ���������Ϊ)
                    '�����˳�Ժ,תԺ,����ҽ��
                    If .TextMatrix(i, COL_�������) = "Z" _
                        And InStr(",5,6,11,", Val(.TextMatrix(i, COL_��������))) > 0 Then
                        mblnRefresh = True
                    End If
                    
                    'һ��Ҫ��������NO
                    Call GetCurBillSet(strNOKey, strNO, -1, lng�������, bln������)
                                                            
                    '�Ƿ�һ��ҽ���ĵ�һҽ����:ҩ�Ƶĵ�һҩƷ��Ϊ��һҽ����
                    blnFirst = False
                    If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = True
                        End If
                    ElseIf .TextMatrix(i, COL_�������) = "C" And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = True '��������еĵ�һ������
                        End If
                    ElseIf InStr(",1,2,3,4,", CLng(.Cell(flexcpData, i, COL_ID))) = 0 Then '�ſ��ɼ�����
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            blnFirst = True
                        End If
                    End If
                                        
                    '��������:ҩƷΪ������λ������,����Ϊ����
                    If .TextMatrix(i, COL_�������) = "7" Then
                        dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����))
                    ElseIf InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_סԺ��װ)) * Val(.TextMatrix(i, COL_����ϵ��))
                    Else
                        dbl�������� = Val(.TextMatrix(i, COL_����))
                    End If
                    dbl�������� = Format(dbl��������, "0.00000")
                                                            
                    '��ĩʱ��
                    str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                    If str�ֽ�ʱ�� <> "" Then
                        str�״�ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        strĩ��ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(Val(.TextMatrix(i, COL_����)) - 1) & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        '�޷��ֽ��Ϊ"һ����"����
                        str�״�ʱ�� = "NULL"
                        strĩ��ʱ�� = "NULL"
                    End If

                    rsSQL.AddNew
                    rsSQL!���� = 4: rsSQL!��ĿID = 0: rsSQL!��� = i
                    rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                    
                    rsSQL!SQL = "ZL_����ҽ������_Insert(" & _
                        Val(.TextMatrix(i, COL_ID)) & "," & lng���ͺ� & "," & IIF(bln������, 1, 2) & ",'" & strNO & "'," & _
                        lng������� & "," & dbl�������� & "," & str�״�ʱ�� & "," & strĩ��ʱ�� & "," & _
                        "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "0," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & int�Ʒ�״̬ & "," & IIF(blnFirst, 1, 0) & ")"
                    rsSQL.Update
                    
                    'Ҫ���͵���δǩ�����¿�ҽ��ID(��ID,һ���еĶ���Ҳ�ᱻǩ��)
                    If Val(.TextMatrix(i, COL_ǩ��ID)) = 0 And Val(.TextMatrix(i, COL_ҽ��״̬)) = 1 Then
                        If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                            lng��ID = Val(.TextMatrix(i, COL_���ID))
                        Else
                            lng��ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If InStr(strҽ��IDs & ",", "," & lng��ID & ",") = 0 Then
                            strҽ��IDs = strҽ��IDs & "," & lng��ID
                        End If
                    End If
                End If
                
                '������ҩ�䷽��
                If .Cell(flexcpData, i, COL_ID) = 3 Then '��ҩ�÷�
                    int�䷽�� = int�䷽�� + 1
                End If
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
            txtPer.Text = CInt(psb.Value) & "%"
            txtPer.Refresh
        Next
                
        '��ʾδ�����Ŀ
        If strAudit <> "" Then
            MsgBox "����""" & mrsPati!���� & """���·�����Ŀ��û�о�����������Ӧ��ҽ�����ܷ��ͣ�" & vbCrLf & strAudit, vbInformation, gstrSysName
            GoTo errH
        End If
        
        '�Զ����е���ǩ��(δǩ������)
        '-----------------------------------------------------------------------------------------
        If Not gobjESign Is Nothing And Mid(gstrESign, 2, 1) = "1" And strҽ��IDs <> "" Then
            strҽ��IDs = Mid(strҽ��IDs, 2) '��������ID,����Ϊ��ϸ��ID
            intRule = ReadAdviceSignSource(1, mlng����ID, mlng��ҳID, strҽ��IDs, 0, False, strSource, mlngǰ��ID)
            If intRule = 0 Then GoTo FuncEnd
            If strSource = "" Then
                Screen.MousePointer = 0
                MsgBox "���ܶ�ȡҪǩ����ҽ��Դ�ġ�", vbInformation, gstrSysName
                GoTo FuncEnd
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID)
            If strSign = "" Then GoTo FuncEnd
            lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
            rsSQL.AddNew
            rsSQL!���� = 2: rsSQL!ҽ��ID = 0: rsSQL!��ĿID = 0: rsSQL!��� = 0
            rsSQL!SQL = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strҽ��IDs & "')"
            rsSQL.Update
        End If
        
        '�ύ��������
        '-----------------------------------------------------------------------------------------
        If Not CompletePatiSend(bln������, rsSQL, rsUpload, cur�ϼ�, str���, str�������, strWarn, intWarn, blnTran) Then GoTo errH
    End With
    SendAdvice = lng���ͺ�
FuncEnd:
    'ɾ�������ѳɹ����͵���
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If Err.Number <> 0 Then
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Function CompletePatiSend(ByVal bln������ As Boolean, rsSQL As ADODB.Recordset, _
    rsUpload As ADODB.Recordset, ByVal cur�ϼ� As Currency, ByVal str��� As String, ByVal str������� As String, _
    strWarn As String, intWarn As Integer, blnTran As Boolean) As Boolean
'���ܣ��ύһ�����˵�ҽ����������,����֮ǰ������ʱ���
'������
'      rsSQL=��������Ҫִ�е�SQL
'      rsUpload=����ҽ���ϴ��ļ��ʵ��ݺ�
'      cur�ϼ�=���˱���Ҫ����ҽ���ļ��ʽ��ϼ�,���ڼ��ʱ���
'      str���=���˱��η��ͼ��ʷ��õ��շ����,���ڼ��ʱ���
'      str���=���˱��η��ͼ��ʷ��õ��շ��������,���ڼ��ʱ���
'      strWarn(I/O)=���ڼ�¼��ǰ�����ѱ������
'      intWarn(I/O)=���ڼ�¼���η��ͱ�����ʾʱ��ѡ����
'˵�����������,���ڵ��ú����д���,blnTran�����Ƿ�����������
    Dim rsWarn As New ADODB.Recordset
    Dim strSQL As String, intR As Integer
    Dim cur���� As Currency, i As Long
    Dim arrNOs() As String, strMsg As String
    
    '���˷��ñ���
    If Not bln������ And cur�ϼ� > 0 Then
        strSQL = "Select Nvl(���ò���,1) as ���ò���,Nvl(��������,1) as ��������," & _
            " ����ֵ,������־1,������־2,������־3 From ���ʱ�����" & _
            " Where ����ID=[1] And Nvl(���ò���,1)=[2]"
        Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsPati!��ǰ����ID), IIF(Nvl(mrsPati!ҽ��, 0) = 1, 2, 1))
        If Not rsWarn.EOF Then
            If rsWarn!�������� = 2 Then cur���� = GetPatiDayMoney(mlng����ID)
            str������� = Mid(str�������, 2)
            For i = 1 To Len(str���)
                intR = BillingWarn(Me, mstrPrivs, rsWarn, mrsPati!����, Nvl(mrsPati!ʣ���, 0), cur����, cur�ϼ�, Nvl(mrsPati!������, 0), Mid(str���, i, 1), Split(str�������, ",")(i - 1), strWarn, intWarn, Nvl(mrsPati!ҽ��, 0) = 1)
                If InStr(",2,3,", intR) > 0 Then Exit For
            Next
        End If
    End If
    
    If InStr(",2,3,", intR) = 0 Then
        'ִ��˳��:1-�Ƽ�,2-ǩ��,3-У��,4-����,5-����,6-����
        '1.�Է��ü�¼���շ�ϸĿID�������
        rsSQL.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst
        rsSQL.Sort = "����,��ĿID,���"
        rsUpload.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst
        
        gcnOracle.BeginTrans: blnTran = True
        Do While Not rsSQL.EOF
            Call zlDatabase.ExecuteProcedure(rsSQL!SQL, Me.Caption)
            rsSQL.MoveNext
        Loop
            
        'ҽ�������ϴ�
        If Not IsNull(mrsPati!����) Then
            If gclsInsure.GetCapability(supportҽ���ϴ�, , mrsPati!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , mrsPati!����) Then
                Do While Not rsUpload.EOF
                    strMsg = "" '��Ϊ����һ��NO�ڿ϶�Ϊһ�����˵�,��������˲������Բ���
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , mrsPati!����) Then
                        'δ�ύǰ�ϴ�ʧ����ع�����ֹ����
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName 'ÿ����ʾ
                        Else
                            MsgBox mrsPati!���� & "�ķ����ϴ�ʧ�ܣ����Ͳ���������ֹ��", vbExclamation, gstrSysName
                        End If
                        Exit Function
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName 'ÿ����ʾ
                    End If
                    rsUpload.MoveNext
                Loop
            End If
        End If
        gcnOracle.CommitTrans: blnTran = False
        
        'ҽ�������ϴ�
        If Not IsNull(mrsPati!����) Then
            If gclsInsure.GetCapability(supportҽ���ϴ�, , mrsPati!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, , mrsPati!����) Then
                Do While Not rsUpload.EOF
                    strMsg = ""
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , mrsPati!����) Then
                        '�ύ���ϴ�ʧ��,����ʾ
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                        Else
                            MsgBox mrsPati!���� & "�ļ��ʵ�""" & rsUpload!NO & """�ϴ�ʧ�ܣ�HIS���������ύ����ȷ���������͡�", vbExclamation, gstrSysName
                        End If
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                    End If
                    rsUpload.MoveNext
                Loop
            End If
        End If
            
        '�ύ�ɹ�,������ҽ���б��Ϊ��ɾ��
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    .RowData(i) = -1
                End If
            Next
        End With
    End If
    CompletePatiSend = True
End Function

Private Sub ShowSendTotal()
'���ܣ����ݵ�ǰѡ��Ҫ���͵�ҽ�������㲢��ʾ���͵�ҽ���ϼ�
    Dim cur��� As Currency, curҩƷ��� As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '�ɼ��еĽ��:��һ��Ļ��ܽ��
                If Not .RowHidden(i) Then
                    cur��� = cur��� + Val(.TextMatrix(i, COL_���))
                End If
                'ҩƷ�Ľ��,ȡԭʼ���
                If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                    curҩƷ��� = curҩƷ��� + Val(.Cell(flexcpData, i, COL_���))
                End If
            End If
        Next
    End With
    stbThis.Panels(3).Text = "���:" & FormatEx(cur���, gbytDec) & "(ҩ" & FormatEx(curҩƷ���, gbytDec) & ")"
    Call Form_Resize
End Sub
