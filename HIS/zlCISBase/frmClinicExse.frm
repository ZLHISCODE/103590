VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmClinicExse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����շѶ���"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   Icon            =   "frmClinicExse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "�����˳�(&E)"
      Height          =   350
      Left            =   11280
      TabIndex        =   28
      Top             =   6465
      Width           =   1215
   End
   Begin VB.CommandButton cmdAutoGet 
      Caption         =   "����ƥ��(&A)"
      Height          =   350
      Left            =   10560
      TabIndex        =   10
      ToolTipText     =   "��ʾ������������Ŀ�����Զ�����,ƥ��ģʽ��ϵͳѡ��--ʹ��ϰ�߾���"
      Top             =   1065
      Width           =   1290
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�(&R)"
      Height          =   350
      Left            =   12000
      Picture         =   "frmClinicExse.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1065
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2715
      Left            =   240
      TabIndex        =   7
      Top             =   7470
      Visible         =   0   'False
      Width           =   6200
      _ExtentX        =   10927
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.OptionButton optPreproty 
      Caption         =   "�ֹ��Ƽ�(&2)"
      Height          =   180
      Index           =   2
      Left            =   11985
      TabIndex        =   22
      Top             =   810
      Width           =   1305
   End
   Begin VB.OptionButton optPreproty 
      Caption         =   "���Ƽ�(&1)"
      Height          =   180
      Index           =   1
      Left            =   10845
      TabIndex        =   21
      Top             =   810
      Width           =   1110
   End
   Begin VB.OptionButton optPreproty 
      Caption         =   "�����Ƽ�(&0)"
      Height          =   180
      Index           =   0
      Left            =   9405
      TabIndex        =   20
      Top             =   810
      Value           =   -1  'True
      Width           =   1410
   End
   Begin VB.Frame fraDept 
      Caption         =   "�����趨"
      Height          =   4725
      Left            =   180
      TabIndex        =   16
      Top             =   1605
      Width           =   2475
      Begin VB.CommandButton cmdCopy 
         Height          =   315
         Left            =   1995
         Picture         =   "frmClinicExse.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "���ƿ���"
         Top             =   300
         Width           =   345
      End
      Begin VB.CommandButton cmdDeptDel 
         Height          =   315
         Left            =   1620
         Picture         =   "frmClinicExse.frx":6F26
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "ɾ������"
         Top             =   300
         Width           =   345
      End
      Begin VB.TextBox txtDept 
         Height          =   350
         Left            =   105
         TabIndex        =   18
         Top             =   300
         Width           =   1455
      End
      Begin VB.ListBox lstDept 
         Height          =   3840
         Left            =   120
         TabIndex        =   17
         Top             =   690
         Width           =   2250
      End
   End
   Begin VB.Frame fraTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   7920
      TabIndex        =   9
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1215
      MaxLength       =   50
      TabIndex        =   2
      Top             =   750
      Width           =   5895
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "&P"
      Height          =   300
      Left            =   7125
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   750
      Width           =   255
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4395
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicExse.frx":D778
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicExse.frx":DD12
            Key             =   "ExseUse"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�������һ��(&S)"
      Height          =   350
      Left            =   9600
      TabIndex        =   4
      Top             =   6465
      Width           =   1575
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   300
      Picture         =   "frmClinicExse.frx":E2AC
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6465
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   12600
      TabIndex        =   5
      Top             =   6465
      Width           =   1100
   End
   Begin TabDlg.SSTab stbExse 
      Height          =   4740
      Left            =   2880
      TabIndex        =   11
      Top             =   1560
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   8361
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "�շ���Ŀ(E)"
      TabPicture(0)   =   "frmClinicExse.frx":E3F6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "msfExse"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��鲿λ(J)"
      TabPicture(1)   =   "frmClinicExse.frx":E412
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vfgExse"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "���Ի����м���(&B)"
      TabPicture(2)   =   "frmClinicExse.frx":E42E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vfg����"
      Tab(2).ControlCount=   1
      Begin ZL9BillEdit.BillEdit msfExse 
         Height          =   4125
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "��ʾ����Del����ɾ��һ��"
         Top             =   480
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7276
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgExse 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   10455
         _cx             =   18441
         _cy             =   7435
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
         AutoResize      =   0   'False
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg���� 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   10455
         _cx             =   18441
         _cy             =   7435
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
         AutoResize      =   0   'False
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.TabStrip tabDept 
      Height          =   5205
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   9181
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���п���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "סԺ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMessage 
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   4200
      TabIndex        =   27
      Tag             =   "��ʾ:"
      Top             =   6525
      Width           =   5175
   End
   Begin VB.Label lblInfo 
      Caption         =   "ҽ�����ͺ�"
      Height          =   255
      Left            =   8280
      TabIndex        =   26
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label txtTotal 
      Height          =   180
      Left            =   2175
      TabIndex        =   25
      Top             =   6525
      Width           =   1815
   End
   Begin VB.Label lbltotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϼƣ�"
      Height          =   180
      Left            =   1665
      TabIndex        =   24
      Top             =   6525
      Width           =   540
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "������Ŀ(&I)"
      Height          =   180
      Left            =   195
      TabIndex        =   1
      Top             =   810
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��ѡ����Ŀ��ָ�����Ӧ�Ĺ̶���Ŀ���շ���Ŀ���Ա�ϵͳ�ܸ���������Ŀ��Ӧ�շ����ݣ����в���ҽ�����Զ��Ʒѡ�"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   840
      TabIndex        =   0
      Top             =   270
      Width           =   10275
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   195
      Picture         =   "frmClinicExse.frx":E44A
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frmClinicExse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1����ǰ״̬����me.cmdClose.tag���棬�ֱ�Ϊ"�޸�"��"����"�����ϼ�����ͨ��ShowMe��������
'   2��ָ����Ŀ����me.lblItem.tag���棬���ϼ�����ͨ��ShowMe�������룬���Դ��ݣ�Ҳ���Բ�����
'��ѡ���շ���Ŀ��
'   1�����������Ϊ�Һš���λ�������ǹ̶���Ŀ���շ���Ŀ
'   2������ҩ�Ƶ��շ�ͨ������Ӧ��������ڲ�����ҩƷ��Ϊ��Ӧ�շ���Ŀ
'---------------------------------------------------
Private strInputed As String
Public rsSelect As New ADODB.Recordset                          '�շ���Ŀѡ��������ʱ�õ�
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer
Dim mrsBwff As ADODB.Recordset '��¼��λ����,��������ĺϷ��Լ��
Dim mstrType As String          '��ǰ������Ŀ���
Dim mstrOper As String          '��ǰ������Ŀ�Ĳ�������
Dim mlngFlag As Long            '��ǰ������Ŀ��ִ�б��
Dim mlngClient As Long          '��ǰ������Ŀ�ķ������
Private mstr��� As String      '��¼��ѡ������
Private mstrIDS As String       '����֮��id��
Private mlngItem As Long        '��ǰ��λ��

Dim mlngLastSource As Long      '�ϴ�ѡ���ҳ���Ӧ�Ĳ�����Դ
Dim mlngLastDeptID As Long      '�ϴ�ѡ��Ŀ���ID
Private mDelDeptList As String  'ɾ�������б�����ʱҪ�ȴ��������������
Private mlngCodeType As Long '0-ƴ����1-���

Private Enum ExseCol
    ��� = 0
    ��Ŀ�� = 1
    ��� = 2
    ��λ = 3
    ��ǰ�� = 4
    ��Ӧ�� = 5
    �̶� = 6
    ���� = 7
    �շѷ�ʽ = 8
    
    ���� = 9
End Enum
'�����б�
Private mDept1() As String   '�������� �����б�
Private mDept2() As String   '����סԺ �����б�
Private mDept3() As String   '������� �����б�
'��ͨ�շѶ���
Private mGen0()  As String '����ȫԺ ��ͨ�շ���Ŀ
Private mGen1()  As String '�������� ��ͨ�շ���Ŀ
Private mGen2()  As String '����סԺ ��ͨ�շ���Ŀ
Private mGen3()  As String '������� ��ͨ�շ���Ŀ
'��λ�շѶ���
Private mPlace0() As String '����ȫԺ ����λ�շ���Ŀ
Private mPlace1() As String '�������� ����λ�շ���Ŀ
Private mPlace2() As String '����סԺ ����λ�շ���Ŀ
Private mPlace3() As String '������� ����λ�շ���Ŀ
'�����շѶ���
Private mAppend0() As String '����ȫԺ �����շ���Ŀ
Private mAppend1() As String '�������� �����շ���Ŀ
Private mAppend2() As String '����סԺ �����շ���Ŀ
Private mAppend3() As String '������� �����շ���Ŀ

'  ������Ŀid_In In �����շѹ�ϵ.������Ŀid%Type,
'  �Ƽ�����_In   ������ĿĿ¼.�Ƽ�����%Type,
'  �շ�����_In   In Varchar2, --��"|"�ָ��������շ����ݣ�ÿ����¼��"�շ���ĿID^����^�̶�^����^����^��λ^��鷽��^�շѷ�ʽ"��֯
'  �Ƿ�ɾ��_In   Number := 1,
'  ���ÿ���id_In �����շѹ�ϵ.���ÿ���id%Type := Null,
'  ������Դ_In   �����շѹ�ϵ.������Դ%Type := 0

Private Sub IniItemList()
    With Me.msfExse
        .Active = True
        .ClearBill
        .MsfObj.FixedCols = 1
        .Rows = 2
        .Cols = ExseCol.����
        .TextMatrix(0, ExseCol.���) = ""
        .TextMatrix(0, ExseCol.��Ŀ��) = "��Ŀ��"
        .TextMatrix(0, ExseCol.���) = "���"
        .TextMatrix(0, ExseCol.��λ) = "��λ"
        .TextMatrix(0, ExseCol.��ǰ��) = "��ǰ��"
        .TextMatrix(0, ExseCol.��Ӧ��) = "��Ӧ��"
        .TextMatrix(0, ExseCol.�̶�) = "�̶�"
        .TextMatrix(0, ExseCol.����) = "����"
        .TextMatrix(0, ExseCol.�շѷ�ʽ) = "�շѷ�ʽ"
        .colData(ExseCol.���) = 5
        .colData(ExseCol.��Ŀ��) = 1
        .colData(ExseCol.���) = 5
        .colData(ExseCol.��λ) = 5
        .colData(ExseCol.��ǰ��) = 5
        .colData(ExseCol.��Ӧ��) = 4
        .colData(ExseCol.�̶�) = -1
        .colData(ExseCol.����) = -1
        .colData(ExseCol.�շѷ�ʽ) = 3
        .ColWidth(ExseCol.���) = 250
        .ColWidth(ExseCol.��Ŀ��) = 2800
        .ColWidth(ExseCol.���) = 1000
        .ColWidth(ExseCol.��λ) = 600
        .ColWidth(ExseCol.��ǰ��) = 800
        .ColWidth(ExseCol.��Ӧ��) = 1000
        .ColWidth(ExseCol.�̶�) = 500
        .ColWidth(ExseCol.����) = 500
        .ColWidth(ExseCol.�շѷ�ʽ) = 3100
        
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
        .ColAlignment(ExseCol.�̶�) = 4
        .ColAlignment(ExseCol.����) = 4
        .ColAlignment(ExseCol.�շѷ�ʽ) = 1
        
        .Clear
        .AddItem ("0-������ȡ")
        
        '�����ࡢ������(�ɼ���ʽ)����Ŀ
        If mstrType = "C" Or (mstrType = "E" And mstrOper = "6") Then
            .AddItem ("1-�����Թܷ���")
        End If
        .AddItem ("2-һ�η���ֻ��ȡһ��")
        .AddItem ("3-����ֻ��ȡһ��")
        .AddItem ("4-����δִ����ȡһ��")
        .AddItem ("5-����ֻ��ȡһ�Σ��ų�������Ŀ")
        .AddItem ("6-����δִ����ȡһ�Σ��ų�������Ŀ")
        .AddItem ("7-ÿ���״β���ȡ")
        .AddItem ("9-�Զ���")
    End With
    
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal blnEdit As Boolean, Optional ByVal lng��Ŀid As Long, Optional ByVal strIDS As String)
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    mstrIDS = strIDS
    If mstrIDS = "" Then Me.cmdSave.Visible = False
    Me.cmdClose.Tag = IIf(blnEdit, "�޸�", "����")
    If Me.cmdClose.Tag = "����" Then
        Me.msfExse.Active = False
        Me.cmdSave.Visible = False
        
        Me.cmdRestore.Visible = False
        Me.cmdAutoGet.Visible = False
        
        txtDept.Enabled = False
        cmdDeptDel.Enabled = False
        cmdCopy.Enabled = False
        cmdItem.Enabled = False
        txtItem.Enabled = False
    Else
        Me.msfExse.Active = True
    End If
    Me.lblItem.Tag = lng��Ŀid
    
    '�õ�����
    GetOneRec
    
    If mstrOper = "����" And mstr��� = "D" Then
        stbExse.TabVisible(1) = False
    End If
    
    Set mrsBwff = zlDatabase.OpenSQLRecord("Select a.���� as ��λ,a.���� From ���Ƽ�鲿λ a, ������ĿĿ¼ b Where a.���� = b.�������� And b.Id = [1]", Me.Caption, lng��Ŀid)
    
    Me.Show 1, frmParent

End Sub

Private Sub GetOneRec()
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.����,I.����,I.���㵥λ,nvl(I.�Ƽ�����,0) as �Ƽ�����,I.���,I.��������,I.ִ�б��,I.������� " & _
            " from ������ĿĿ¼ I" & _
            " where I.���>='A' and I.ID=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblItem.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag
        Else
            mstr��� = !���
            Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !���� & "]" & !����: Me.txtItem.Text = Me.txtItem.Tag
            Me.optPreproty(!�Ƽ�����).Value = True
            mstrType = Trim("" & !���)
            mstrOper = IIf(IsNull(!��������), "", !��������)
            mlngFlag = Val("" & !ִ�б��)
            mlngClient = Val("" & !�������)
            Call zlExseRef(Me.lblItem.Tag)
        End If
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdCopy_Click()
    Dim lngDeptID As Long, lngSource As Long
    If Val(Me.lblItem.Tag) <= 0 Then Exit Sub
    If lstDept.ListIndex <> -1 Then
        lngDeptID = lstDept.ItemData(lstDept.ListIndex)
        lngSource = GetCurrSource
        Call DeptCopy(lngSource, lngDeptID)
    Else
        MsgBox "����ѡ��һ�����ң�", vbInformation
    End If
    
End Sub

Private Sub cmdDeptDel_Click()
    Dim lngDeptID As Long, lngSource As Long
    
    If Val(Me.lblItem.Tag) <= 0 Then Exit Sub
    If lstDept.ListIndex <> -1 Then
        lngDeptID = lstDept.ItemData(lstDept.ListIndex)
        lngSource = GetCurrSource
        If delDept(lngSource, lngDeptID) Then
            '���浽ɾ�������б�����ʱ��ɾ�����������Ŀ��Ҷ���
            If InStr(mDelDeptList & ",", "," & lngSource & "|" & lngDeptID & ",") <= 0 Then
                mDelDeptList = mDelDeptList & "," & lngSource & "|" & lngDeptID
            End If
            lstDept.RemoveItem lstDept.ListIndex
        End If
        'Call zlExseRef(Me.lblItem.Tag)
        Call msfExseRef(lngDeptID, -1)
    Else
        MsgBox "����ѡ��һ�����ң�", vbInformation
    End If
End Sub

Private Sub cmdRestore_Click()
    Call zlExseRef(Me.lblItem.Tag)
    strInputed = ""
End Sub

Private Sub cmdSave_Click()
    Dim lngId As Long
    Dim lngCount As Long
    SaveData
    
    lngId = Split(mstrIDS, ",")(mlngItem)
    Me.lblItem.Tag = lngId
    
    '�õ�����
    GetOneRec
    Me.txtItem.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
    mlngItem = mlngItem + 1
    If mlngItem = UBound(Split(mstrIDS, ",")) Then Me.cmdSave.Enabled = False
End Sub
Private Sub cmdSaveExit_Click()
    SaveData
    Unload Me
End Sub
Private Sub SaveData()
    Dim blnErr As Boolean
    Dim i As Integer
    Dim varDelList As Variant
    Dim NullGen(0) As String, NullPlan(0) As String, NullAppend(0) As String
    On Error GoTo hErr
    If Val(Me.lblItem.Tag) = 0 Then lblMessage.Caption = lblMessage.Tag & "δ��ȷָ��������Ŀ��": Me.txtItem.SetFocus: Exit Sub
    
    '��ɾ���б��еĿ�������
    If mDelDeptList <> "" Then
        If Left(mDelDeptList, 1) = "," Then mDelDeptList = Mid(mDelDeptList, 2)
        varDelList = Split(mDelDeptList, ",")
        For i = LBound(varDelList) To UBound(varDelList)
            Call SaveArryData(CLng(Split(varDelList(i), "|")(0)), CLng(Split(varDelList(i), "|")(1)), NullGen, NullPlan, NullAppend)
        Next
        mDelDeptList = ""
    End If
    
    Call lstDeptSelect(1) '���浱ǰ�����ϵ����ݵ�����
    If mGen0(UBound(mGen0)) <> "" Then
        blnErr = CheckArrData(mGen0)
        If blnErr = False Then Exit Sub
    End If
    If mGen1(UBound(mGen1)) <> "" Then
        blnErr = CheckArrData(mGen1)
        If blnErr = False Then Exit Sub
    End If
    If mGen2(UBound(mGen2)) <> "" Then
        blnErr = CheckArrData(mGen2)
        If blnErr = False Then Exit Sub
    End If
    If mGen3(UBound(mGen3)) <> "" Then
        blnErr = CheckArrData(mGen3)
        If blnErr = False Then Exit Sub
    End If
    Call SaveArryData(0, 0, mGen0, mPlace0, mAppend0)
    
    For i = LBound(mDept1) To UBound(mDept1)
        If mDept1(i) <> "" Then
            Call SaveArryData(1, Val(Split(mDept1(i), "|")(0)), mGen1, mPlace1, mAppend1)
        End If
    Next
    
    For i = LBound(mDept2) To UBound(mDept2)
        If mDept2(i) <> "" Then
            Call SaveArryData(2, Val(Split(mDept2(i), "|")(0)), mGen2, mPlace2, mAppend2)
        End If
    Next
    
    For i = LBound(mDept3) To UBound(mDept3)
        If mDept3(i) <> "" Then
            Call SaveArryData(3, Val(Split(mDept3(i), "|")(0)), mGen3, mPlace3, mAppend3)
        End If
    Next
    
    lblMessage.Caption = lblMessage.Tag & Mid(Me.txtItem.Text, 1, 18) & " �շѶ��ձ���ɹ���"
    Call zlExseRef(Me.lblItem.Tag)
    Me.txtItem.SetFocus
    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdItem_Click()
    Dim rsTmp As ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand

    gstrSql = "Select Distinct 0 As ĩ��,id,�ϼ�ID,����,����,'' As ���㵥λ,0 As �Ƽ�����,'' As ���, '' As ��������,0 as ִ�б��, 0 as ������� " & _
        " From ���Ʒ���Ŀ¼ Start With id In (Select ����ID" & _
            " from ������ĿĿ¼ I" & _
            " where I.���>='A'" & _
            " And (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))) Connect By Prior �ϼ�id=id" & _
        " Union All"
    gstrSql = gstrSql & " Select 1 As ĩ��,I.ID,����ID As �ϼ�id,I.����,I.����,I.���㵥λ,nvl(I.�Ƽ�����,0) as �Ƽ�����, ���, ��������, ִ�б��, ������� " & _
            " from ������ĿĿ¼ I" & _
            " where I.���>='A'" & _
            " And (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) Order By ����"
    Set rsTmp = zlDatabase.ShowSelect(Me, gstrSql, 2, "������Ŀ", , , , , True)
    If Not rsTmp Is Nothing Then
        mstr��� = rsTmp!���
        Me.lblItem.Tag = rsTmp("ID")
        Me.txtItem.Tag = "[" & rsTmp("����") & "]" & rsTmp("����")
        Me.txtItem.Text = Me.txtItem.Tag
        Me.optPreproty(rsTmp("�Ƽ�����")).Value = True
        mstrType = Trim("" & rsTmp("���"))
        mstrOper = IIf(IsNull(rsTmp("��������")), "", rsTmp("��������"))
        mlngFlag = Val("" & rsTmp!ִ�б��)
        mlngClient = Val("" & rsTmp!�������)
        Call zlExseRef(Me.lblItem.Tag)
        Me.txtItem.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        If Me.lvwItems.Tag = Me.txtItem.Name Then
            Me.txtItem.SetFocus
        Else
            Me.msfExse.SetFocus
        End If

    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mlngCodeType = zlDatabase.GetPara("���뷽ʽ")
    Call IniItemList
    mDelDeptList = "" '���ɾ�������б�
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 2000
        .Add , "���", "���", 1200
        .Add , "����", "����", 1000
        .Add , "���㵥λ", "��λ", 800
        .Add , "����", "����", 1000
        .Add , "�ۼ�", "�ۼ�", 1000
        .Add , "���", "���", 0
        .Add , "��������", "��������", 0
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngItem = 0
End Sub

Private Sub lstDept_Click()
    Dim lngDeptID As Long
    '��ʾ��ǰ���ҵķ��ö���
    If lstDept.ListIndex >= 0 Then
        lngDeptID = lstDept.ItemData(lstDept.ListIndex)
        Call msfExseRef(lngDeptID, 1)
        txtDept.Text = lstDept.List(lstDept.ListIndex)
    Else
        'δѡ �п��ң������ʾ
        Call msfExseRef(-1, 1)
        txtDept.Text = ""
    End If
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim dblCurrJe As Double
    
    On Error GoTo ErrHandle
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        If .Tag = Me.txtItem.Name Then
            If Me.lblItem.Tag <> Mid(.SelectedItem.Key, 2) Then
                Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
                Me.txtItem.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
                Me.txtItem.Text = Me.txtItem.Tag
                Me.optPreproty(Val(.SelectedItem.Tag)).Value = True
                mstrType = .SelectedItem.SubItems(.ColumnHeaders("���").Index - 1)
                mstrOper = .SelectedItem.SubItems(.ColumnHeaders("��������").Index - 1)
                Call zlExseRef(Me.lblItem.Tag)
            End If
            Me.txtItem.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            dblCurrJe = Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��))
            
            Me.msfExse.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
            Me.msfExse.RowData(Me.msfExse.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ŀ��) = Me.msfExse.Text
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.���) = .SelectedItem.SubItems(.ColumnHeaders("���").Index - 1)
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��λ) = .SelectedItem.SubItems(.ColumnHeaders("���㵥λ").Index - 1)
            If Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��)) = 0 Then
                Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��) = "1"
                Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.�̶�) = "��"
            End If
            
            gstrSql = "select decode(I.�Ƿ���,1,'���',to_char(P.�۸�)) As �۸�" & _
                    " from (select �Ƿ��� from �շ���ĿĿ¼ where id=[1]) I," & _
                    "      (Select sum(�ּ�) As �۸�" & _
                    "      From �շѼ�Ŀ  Where �۸�ȼ� Is Null and �շ�ϸĿid=[1] and ִ������<=Sysdate And (��ֹ���� Is Null Or ��ֹ����>=Sysdate)) P"
                    
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.msfExse.RowData(Me.msfExse.Row)), gstrPriceClass)
            
            With rsTemp
                If .RecordCount > 0 Then
                    Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��) = IIf(IsNull(!�۸�), "", !�۸�)
                Else
                    Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��) = ""
                End If
            
                txtTotal = Val(txtTotal) - dblCurrJe + Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��))
                txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
            End With
            Me.msfExse.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfExse_AfterAddRow(Row As Long)
    With Me.msfExse
        If .Rows > 2 Then
            .TextMatrix(1, ExseCol.���) = 1
        End If
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, ExseCol.���) = intCount
            If .Rows > 2 Then
                .TextMatrix(intCount, ExseCol.����) = .TextMatrix(intCount - 1, ExseCol.����)
            End If
        Next
    End With
End Sub

Private Sub msfExse_AfterDeleteRow()
    With Me.msfExse
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, ExseCol.���) = intCount
        Next
    End With
End Sub

Private Sub msfExse_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    txtTotal = Val(txtTotal) - Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��))
    txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
End Sub

Private Sub msfExse_CommandClick()
    Dim rsTmp As ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim dblCurrJe As Double
    err = 0: On Error GoTo ErrHand
    Dim strKey As String
    Dim str������� As String

    '����ѡ����
    strKey = tabDept.SelectedItem.Key
    
    If strKey = "ALL" Or strKey = "TJ" Then
        str������� = "1,2,3"
    ElseIf strKey = "MZ" Then
        str������� = "1,3"
    ElseIf strKey = "ZY" Then
        str������� = "2,3"
    End If
    
    frmClinicExseSelect.ShowMe Me, str�������
    Set rsTmp = rsSelect
    
    If Not rsTmp Is Nothing And rsTmp.State = 1 Then
        Me.msfExse.Text = "[" & rsTmp("����") & "]" & rsTmp("����")
        Me.msfExse.RowData(Me.msfExse.Row) = rsTmp("ID")
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ŀ��) = Me.msfExse.Text
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.���) = IIf(IsNull(rsTmp("���")), "", rsTmp("���"))
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��λ) = IIf(IsNull(rsTmp("���㵥λ")), "", rsTmp("���㵥λ"))
        If Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��)) = 0 Then
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��) = "1"
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.�̶�) = "��"
        End If
        
        gstrSql = "select decode(I.�Ƿ���,1,'���',to_char(P.�۸�)) As �۸�" & _
                " from (select �Ƿ��� from �շ���ĿĿ¼ where id=[1]) I," & _
                "      (Select sum(�ּ�) As �۸�" & _
                "      From �շѼ�Ŀ " & _
                "      Where �շ�ϸĿid=[1]  And �۸�ȼ� Is NULL  and ִ������<=Sysdate And (��ֹ���� Is Null Or ��ֹ����>=Sysdate)) P"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.msfExse.RowData(Me.msfExse.Row)), gstrPriceClass)
        
        With rsTemp
            If .RecordCount > 0 Then
                Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��) = IIf(IsNull(!�۸�), "", !�۸�)
            Else
                Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��) = ""
            End If
        
            txtTotal = Val(txtTotal) - dblCurrJe + Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��))
            txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
        End With
        Me.msfExse.SetFocus
        Call zlCommFun.PressKey(vbKeyRight)
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfExse_DblClick(Cancel As Boolean)
    If msfExse.Col = ExseCol.�̶� Then
        With msfExse
            If .TextMatrix(.Row, ExseCol.��Ӧ��) <> "" And IsNumeric(.TextMatrix(.Row, ExseCol.��Ӧ��)) Then
                If Int(.TextMatrix(.Row, ExseCol.��Ӧ��)) = 0 Then
                    Cancel = True
                    .TextMatrix(msfExse.Row, ExseCol.�̶�) = ""
                    lblMessage.Caption = lblMessage.Tag & "��Ӧ������Ϊ0ʱֻ����Ϊ�ǹ̶���."
                End If
            End If
        End With
    End If
End Sub

Private Sub msfExse_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfExse.TextMatrix(Row, Col)
End Sub

Private Sub msfExse_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfExse_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim dblCurrJe As Double
    Dim str������� As String
    
    If KeyCode <> vbKeyReturn Then
        If msfExse.Col = ExseCol.�̶� And KeyCode = vbKeySpace Then  '���ù̶���ʱ���
            With msfExse
                 If Int(Nvl(.TextMatrix(.Row, ExseCol.��Ӧ��), 0)) = 0 Then
                        Cancel = True
                        .Text = ""
                         .TextMatrix(msfExse.Row, ExseCol.�̶�) = ""
                         lblMessage.Caption = lblMessage.Tag & "��Ӧ��Ϊ0ʱ��Ӧ����Ϊ�̶���."
                 End If
            End With
        End If
        Exit Sub
    End If
    
    lblMessage.Caption = ""
    
    With Me.msfExse
        If .Active = False Then Exit Sub
        If .Col = ExseCol.��ǰ�� And .TxtVisible Then
            .Text = Format(.Text, "0.00000"): .TextMatrix(.Row, ExseCol.��ǰ��) = .Text
        End If
        If .Col <> ExseCol.��Ŀ�� Then
            '��Ӧ��Ϊ��ʱ,��������Ϊ�̶���
            If .Col = ExseCol.��Ӧ�� Then
                If Not IsNumeric(Nvl(.Text, "X")) Then
                    lblMessage.Caption = lblMessage.Tag & "��Ӧ������Ϊ�գ�����Ҫ������Ϊ������."
                    .TxtSetFocus
                    Exit Sub
                End If
                If Int(.Text) = 0 And .TextMatrix(.Row, ExseCol.�̶�) = "��" Then
                    .TextMatrix(.Row, ExseCol.�̶�) = ""
                    lblMessage.Caption = lblMessage.Tag & "��Ӧ������Ϊ0ʱ���Զ�����Ϊ�ǹ̶���."
                End If
            End If
            Exit Sub
        End If
        If .TxtVisible = False Then
            If .TextMatrix(.Row, ExseCol.��Ŀ��) = "" Then Exit Sub
            strTemp = Trim(.TextMatrix(.Row, ExseCol.��Ŀ��))
        Else
            If Trim(.Text) = "" Then Exit Sub
            strTemp = Trim(.Text)
        End If
    End With
    If Trim(strTemp) = Trim(strInputed) Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
    strTemp = UCase(strTemp)
    If tabDept.SelectedItem.Key = "ALL" Or tabDept.SelectedItem.Key = "TJ" Then
        str������� = " And (I.�������=1 or I.�������=2 or I.�������=3) "
    ElseIf tabDept.SelectedItem.Key = "MZ" Then
        str������� = " And (I.�������=1 or I.�������=3) "
    ElseIf tabDept.SelectedItem.Key = "ZY" Then
        str������� = " And (I.�������=2 or I.�������=3) "
    End If
    gstrSql = "Select c.Id, c.����, c.����, c.���, c.����, c.���㵥λ," & vbNewLine & _
            "       Decode(Nvl(c.�Ƿ���, 0)," & vbNewLine & _
            "               0," & vbNewLine & _
            "               Ltrim(Rtrim(To_Char(Nvl(d.�ּ�, 0), '9999999990.0000')))," & vbNewLine & _
            "               Decode(Instr('4567', c.���), 0, Ltrim(Rtrim(To_Char(d.ȱʡ�۸�, 0), '9999999990.0000')), 'ʱ��')) as �ۼ�" & vbNewLine & _
            "From (Select Distinct (a.Id), a.����, a.����, a.���, a.����, a.���㵥λ, a.�Ƿ���, a.���" & vbNewLine & _
            "       From �շ���ĿĿ¼ a, �շ���Ŀ���� b" & vbNewLine & _
            "       Where a.Id = b.�շ�ϸĿid And a.��� Not In ('1', 'J') And" & vbNewLine & _
            "             (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "             (a.������� = 1 Or a.������� = 2 Or a.������� = 3) And" & vbNewLine & _
            "             (a.���� Like [1] Or b.���� Like [2] Or b.���� Like [2]) And b.���� = [4]) c," & vbNewLine & _
            "     �շѼ�Ŀ d" & vbNewLine & _
            "Where c.Id = d.�շ�ϸĿid(+) And d.ִ������ <= Sysdate And (d.��ֹ���� > Sysdate Or d.��ֹ���� Is Null) And d.�۸�ȼ� Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%", gstrPriceClass, mlngCodeType + 1)
    
    If rsTemp.BOF Or rsTemp.EOF Then
        Me.msfExse.Text = ""
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ŀ��) = ""
        lblMessage.Caption = lblMessage.Tag & "δ�ҵ�ָ���շ���Ŀ��"
        Exit Sub
    End If
    
    If rsTemp.RecordCount = 1 Then
        Me.msfExse.Text = "[" & rsTemp!���� & "]" & rsTemp!����
        Me.msfExse.RowData(Me.msfExse.Row) = rsTemp!ID
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ŀ��) = Me.msfExse.Text
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��λ) = IIf(IsNull(rsTemp!���㵥λ), "", rsTemp!���㵥λ)
        If Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��)) = 0 Then
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��) = "1"
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.�̶�) = "��"
        End If
       
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��) = IIf(IsNull(rsTemp!�ۼ�), "", rsTemp!�ۼ�)
        
        txtTotal = Val(txtTotal) - dblCurrJe + Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��ǰ��)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.��Ӧ��))
        txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
        Exit Sub
    End If
    Me.lvwItems.ListItems.Clear
    Do While Not rsTemp.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTemp!ID, rsTemp!����)
        objItem.Icon = "ExseUse": objItem.SmallIcon = "ExseUse"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = rsTemp!����
        objItem.SubItems(Me.lvwItems.ColumnHeaders("���㵥λ").Index - 1) = IIf(IsNull(rsTemp!���㵥λ), "", rsTemp!���㵥λ)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("�ۼ�").Index - 1) = IIf(IsNull(rsTemp!�ۼ�), "", rsTemp!�ۼ�)
        rsTemp.MoveNext
    Loop
    Me.lvwItems.ListItems(1).Selected = True
  
    With Me.lvwItems
        .Tag = Me.msfExse.Name
        .Left = stbExse.Left + Me.msfExse.Left + 300
        .Top = stbExse.Top + Me.msfExse.Top + msfExse.RowHeight(msfExse.Row) * (msfExse.Row + 1)
        '.Height = .ListItems(1).Height * (.ListItems.Count + 1)
        If .Top > Me.msfExse.Top + Me.msfExse.Height Then
            .Top = Me.msfExse.Top + Me.msfExse.Height
        End If
        .Height = Me.Height - .Top - .ListItems(1).Top * 2
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdAutoGet_Click()
'���ܣ�����������Ŀ���ƣ��Զ��ҳ���Ӧ���շ���Ŀ
Dim strTemp As String
Dim n As Integer

strTemp = Mid(txtItem.Text, InStr(1, txtItem.Text, "]") + 1)
With msfExse
    If .Col <> ExseCol.��Ŀ�� Then .Col = ExseCol.��Ŀ��
    .SetFocus
    If .TextMatrix(.Rows - 1, ExseCol.��Ŀ��) <> "" Then
        .Rows = .Rows + 1
    End If
    .TextMatrix(.Rows - 1, ExseCol.��Ŀ��) = strTemp
    .Row = .Rows - 1
    strInputed = ""
    
    '����Ƿ��ظ�
    If .Rows > 2 Then
        For n = 1 To .Rows - 2
            If strTemp = Trim(Mid(.TextMatrix(n, ExseCol.��Ŀ��), InStr(1, .TextMatrix(n, ExseCol.��Ŀ��), "]") + 1)) Then
                .Rows = .Rows - 1
                Exit Sub
            End If
        Next
    End If
End With
Call msfExse_KeyDown(vbKeyReturn, 0, False)

End Sub

Private Sub msfExse_LeaveCell(Row As Long, Col As Long)
    Select Case Col
        Case ExseCol.��Ӧ��
            txtTotal = Val(txtTotal) + Val(Me.msfExse.TextMatrix(Row, ExseCol.��ǰ��)) * (Val(Me.msfExse.TextMatrix(Row, ExseCol.��Ӧ��)) - Val(strInputed))
            txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
        Case ExseCol.�շѷ�ʽ
            msfExse.TextMatrix(Row, Col) = msfExse.CboText
    End Select
End Sub

Private Sub optPreproty_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub stbExse_Click(PreviousTab As Integer)
    msfExse.Visible = stbExse.Tab = 0
    vfgExse.Visible = stbExse.Tab = 1
    vfg����.Visible = stbExse.Tab = 2
End Sub

Private Sub tabDept_Click()
    Call ResizeTabDept
    Call lstDeptSelect(1)
End Sub

Private Sub txtDept_GotFocus()
    Call zlControl.TxtSelAll(txtDept)
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call DeptSelect(txtDept.Text)
    End If
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtItem.Text))
    If strTemp = "" Then Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.����,I.����,I.���㵥λ,nvl(I.�Ƽ�����,0) as �Ƽ�����,I.���,I.�������� " & _
            " from ������ĿĿ¼ I,������Ŀ���� N" & _
            " where I.ID=N.������ĿID and I.���>='A'" & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.���� like [1] or N.���� like [2] or N.���� like [2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
    With rsTemp
        If .BOF Or .EOF Then
            lblMessage.Caption = lblMessage.Tag & "δ�ҵ�ָ����������Ŀ��������ָ��"
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag: Me.txtItem.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblItem.Tag <> !ID Then
                Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !���� & "]" & !����: Me.txtItem.Text = Me.txtItem.Tag
                Me.optPreproty(!�Ƽ�����).Value = True
                mstrType = !���
                mstrOper = IIf(IsNull(!��������), "", !��������)
                Call zlExseRef(Me.lblItem.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            'objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("���㵥λ").Index - 1) = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = !���
            objItem.SubItems(Me.lvwItems.ColumnHeaders("��������").Index - 1) = IIf(IsNull(!��������), "", !��������)
            objItem.Tag = !�Ƽ�����
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtItem.Name
        .Left = Me.txtItem.Left
        .Top = Me.txtItem.Top + Me.txtItem.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_LostFocus()
    Me.txtItem.Text = Me.txtItem.Tag
End Sub

Private Sub zlExseRef(lngItemID As Long)
    '--------------------------------------------------------
    '���ܣ�ˢ����ʾ������Ŀ��Ӧ���շ���Ŀ
    '��Σ�lngItemId-ָ����������Ŀid
    '--------------------------------------------------------
    Dim dblTotal As Double
    Dim n As Integer
    
    err = 0: On Error GoTo ErrHand
    
    '���ݷ��������ʾ��ѡ�����

    If mlngClient = 0 Then
        '0(��)-��ֱ��Ӧ���ڲ���,1-����,2-סԺ,3-�����סԺ(ȫԺ),4-���
        MsgBox "����Ŀ��ֱ��Ӧ���ڲ��ˣ�"
        Exit Sub
    ElseIf mlngClient = 1 Then
        tabDept.Tabs.Clear
        tabDept.Tabs.Add 1, "ALL", "���п���"
        tabDept.Tabs.Add 2, "MZ", "�������"
    ElseIf mlngClient = 2 Then
        tabDept.Tabs.Clear
        tabDept.Tabs.Add 1, "ALL", "���п���"
        tabDept.Tabs.Add 2, "ZY", "סԺ����"
    ElseIf mlngClient = 3 Then
        tabDept.Tabs.Clear
        tabDept.Tabs.Add 1, "ALL", "���п���"
        tabDept.Tabs.Add 2, "MZ", "�������"
        tabDept.Tabs.Add 3, "ZY", "סԺ����"
    ElseIf mlngClient = 4 Then
        tabDept.Tabs.Clear
        tabDept.Tabs.Add 1, "ALL", "���п���"
        tabDept.Tabs.Add 2, "TJ", "������"
    End If
    
    If tabDept.SelectedItem Is Nothing Then
        tabDept.SelectedItem = tabDept.Tabs("ALL")
    End If
    
    Call ResizeTabDept
    
    '��ȡ��������
    Call ReadClinicData(lngItemID, mstrType, mlngFlag)
    '����ѡ��Ŀ�����ʾ����
    Call lstDeptSelect(0)
    
    If mstrOper = "����" And mstr��� = "D" Then
        stbExse.TabVisible(1) = False
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lstDeptSelect(ByVal lngCache As Long)
        '����ѡ��Ŀ�����ʾ����
        '
        ' lngCache : 0-��ʼ���� <>0 �����л�����<>0ʱ����Ҫ����ԭ�����ϵ����ݡ�
        '
        Dim strDept() As String '����Ҫ��ʾ�Ŀ���
        Dim curDept() As String '��ʾǰҪ����Ŀ���
        Dim lngListIndex As Long
        Dim lngRow As Long
        Dim lngSource As Long '��ǰ������Դ
        On Error GoTo hErr

    
100     ReDim strDept(0) As String
102     ReDim curDept(0) As String
104     If lngCache <> 0 Then
            '�������
106         If lstDept.ListCount > 0 Then
108             For lngRow = 0 To lstDept.ListCount - 1
110                 If curDept(UBound(curDept)) <> "" Then ReDim Preserve curDept(UBound(curDept) + 1)
112                 curDept(UBound(curDept)) = lstDept.ItemData(lngRow) & "|" & _
                                               Replace(Split(lstDept.List(lngRow), "(")(1), ")", "") & "|" & _
                                               Split(lstDept.List(lngRow), "(")(0)
                Next
            End If
            
114         If mlngLastSource = 1 Then
116             mDept1 = curDept
118         ElseIf mlngLastSource = 2 Then
120             mDept2 = curDept
122         ElseIf mlngLastSource = 3 Then
124             mDept3 = curDept
            End If
        End If
        '���ݵ�ǰѡ��ҳ�棬ѡ������
        lngSource = GetCurrSource
126     If lngSource = 0 Then
            '��ʾ ��ǰҳ������
128         Call msfExseRef(0, lngCache)
            Exit Sub
130     ElseIf lngSource = 1 Then
132         strDept = mDept1
134     ElseIf lngSource = 2 Then
136         strDept = mDept2
138     ElseIf lngSource = 3 Then
140         strDept = mDept3
        End If
142     txtDept.Text = ""
        '��ӿ���
144     lngListIndex = -1
146     lstDept.Clear
148     For lngRow = LBound(strDept) To UBound(strDept)
150         If strDept(lngRow) <> "" Then
                'strDept��ʽ�� ���ÿ���ID | ���ұ���  |  ��������
152             lstDept.AddItem CStr(Split(strDept(lngRow), "|")(2)) & "(" & Split(strDept(lngRow), "|")(1) & ")"
154             lstDept.ItemData(lstDept.NewIndex) = Val(Split(strDept(lngRow), "|")(0))
156             If mlngLastDeptID <> 0 And lstDept.ItemData(lstDept.NewIndex) = mlngLastDeptID Then
158                 lngListIndex = lstDept.NewIndex
                End If
            End If
        Next
160     If lngListIndex <> -1 Then
162         lstDept.ListIndex = lngListIndex
164     ElseIf lstDept.ListCount > 0 Then
166        lstDept.ListIndex = 0
        Else
            '�޿��ң���տؼ�����ʾ�Ķ�������
168         Call msfExseRef(-1, lngCache)
        End If
        Exit Sub
hErr:
170     MsgBox "lstDeptSelect��" & CStr(Erl()) & "�У�" & err.Description
    '    If ErrCenter = 1 Then
    '        Resume
    '    End If
End Sub

Private Sub ReadClinicData(ByVal lngItemID As Long, strType As String, lngFlag As Long)
        '��ȡ������Ŀ���ݣ������浽���ر�����
        'lngItemID:������Ŀ.ID
        'strType  :������Ŀ.���
        'lngFlag  :������Ŀ.ִ�б��
    
        Dim strSql As String, rsCharge As ADODB.Recordset
        Dim strTmp As String, strTmpDept As String, strDeptList As String
        Dim str�շѷ�ʽ As String
100     err = 0: On Error GoTo ErrHand
    
102     ReDim mDept1(0) As String: ReDim mDept2(0) As String: ReDim mDept3(0) As String
104     ReDim mGen0(0) As String: ReDim mGen1(0) As String: ReDim mGen2(0) As String: ReDim mGen3(0) As String
106     ReDim mPlace0(0) As String: ReDim mPlace1(0) As String: ReDim mPlace2(0) As String: ReDim mPlace3(0) As String
108     ReDim mAppend0(0) As String: ReDim mAppend1(0) As String: ReDim mAppend2(0) As String: ReDim mAppend3(0) As String
    
        '��ͨ���շѶ���
110     strSql = "Select i.Id, '[' || i.���� || ']' || i.���� As ����, i.���, i.���㵥λ, Decode(i.�Ƿ���, 1, '���', To_Char(Sum(p.�ּ�))) As �۸�," & vbNewLine & _
            "       Nvl(r.�շ�����, 0) As ����, Nvl(r.���ж���, 0) As �̶�, Nvl(r.������Ŀ, 0) As ����, r.��������, Nvl(r.�շѷ�ʽ, 0) As �շѷ�ʽ," & vbNewLine & _
            "       To_Number(r.������Դ) As ������Դ, r.���ÿ���id, b.���� As ���ұ���, b.���� As ��������" & vbNewLine & _
            "From �����շѹ�ϵ R, �շ���ĿĿ¼ I, �շѼ�Ŀ P, ���ű� B" & vbNewLine & _
            "Where r.�շ���Ŀid = i.Id And i.Id = p.�շ�ϸĿid(+) And (r.�������� = 0 Or r.�������� Is Null) And r.��鲿λ Is Null " & _
            "       And r.���ÿ���id = b.Id(+) And r.������Ŀid = [1] " & vbNewLine & _
            "       And p.ִ������ <= Sysdate And (p.��ֹ���� Is Null Or p.��ֹ���� >= Sysdate)  And P.�۸�ȼ� Is Null " & _
            "Group By i.Id, i.����, i.����, i.���, i.���㵥λ, i.�Ƿ���, r.�շ�����, r.���ж���, r.������Ŀ, r.��������, r.�շѷ�ʽ, r.������Դ, r.���ÿ���id, b.����, b.���� " & _
            "Order By r.������Դ, b.����, Nvl(r.������Ŀ, 0)"

112     Set rsCharge = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngItemID, gstrPriceClass)
114     strDeptList = ""
116     Do Until rsCharge.EOF
118         Select Case Val("" & rsCharge!�շѷ�ʽ)
            Case 0: str�շѷ�ʽ = "0-������ȡ"
120         Case 1: str�շѷ�ʽ = "1-�����Թܷ���"
122         Case 2: str�շѷ�ʽ = "2-һ�η���ֻ��ȡһ��"
124         Case 3: str�շѷ�ʽ = "3-����ֻ��ȡһ��"
126         Case 4: str�շѷ�ʽ = "4-����δִ����ȡһ��"
128         Case 5: str�շѷ�ʽ = "5-����ֻ��ȡһ�Σ��ų�������Ŀ"
130         Case 6: str�շѷ�ʽ = "6-����δִ����ȡһ�Σ��ų�������Ŀ"
            Case 7: str�շѷ�ʽ = "7-ÿ���״β���ȡ"
            Case 9: str�շѷ�ʽ = "9-�Զ���"
132         Case Else
134             str�շѷ�ʽ = ""
            End Select
        
136         strTmp = "" & rsCharge!ID & "|" & "" & rsCharge!���� & "|" & rsCharge!��� & "|" & rsCharge!���㵥λ & "|" & _
                     rsCharge!�۸� & "|" & rsCharge!���� & "|" & IIf(rsCharge!�̶� = 0, "", "��") & "|" & _
                     IIf(rsCharge!���� = 0, "", "��") & "|" & str�շѷ�ʽ & "|" & _
                     Val("" & rsCharge!������Դ) & "|" & Val("" & rsCharge!���ÿ���id)
            
138         If Val("" & rsCharge!���ÿ���id) <> 0 And InStr(strDeptList, "," & Val("" & rsCharge!���ÿ���id) & ":" & rsCharge!������Դ & ",") <= 0 Then
140             strTmpDept = Val("" & rsCharge!���ÿ���id) & "|" & rsCharge!���ұ��� & "|" & rsCharge!��������
142             strDeptList = strDeptList & "," & Val("" & rsCharge!���ÿ���id) & ":" & rsCharge!������Դ & ","
            Else
144             strTmpDept = ""
            End If
            
146         If Val("" & rsCharge!������Դ) = 0 Then
                'ȫԺ
148             If mGen0(UBound(mGen0)) <> "" Then ReDim Preserve mGen0(UBound(mGen0) + 1)
150             mGen0(UBound(mGen0)) = strTmp
152         ElseIf Val("" & rsCharge!������Դ) = 1 Then
                '����
154             If mGen1(UBound(mGen1)) <> "" Then ReDim Preserve mGen1(UBound(mGen1) + 1)
156             mGen1(UBound(mGen1)) = strTmp
            
158             If strTmpDept <> "" Then
160                 If mDept1(UBound(mDept1)) <> "" Then ReDim Preserve mDept1(UBound(mDept1) + 1)
162                 mDept1(UBound(mDept1)) = strTmpDept
                End If
164         ElseIf Val("" & rsCharge!������Դ) = 2 Then
                'סԺ
166             If mGen2(UBound(mGen2)) <> "" Then ReDim Preserve mGen2(UBound(mGen2) + 1)
168             mGen2(UBound(mGen2)) = strTmp
170             If strTmpDept <> "" Then
172                 If mDept2(UBound(mDept2)) <> "" Then ReDim Preserve mDept2(UBound(mDept2) + 1)
174                 mDept2(UBound(mDept2)) = strTmpDept
                End If
176         ElseIf Val("" & rsCharge!������Դ) = 3 Then
                '���
178             If mGen3(UBound(mGen3)) <> "" Then ReDim Preserve mGen3(UBound(mGen3) + 1)
180             mGen3(UBound(mGen3)) = strTmp
182             If strTmpDept <> "" Then
184                 If mDept3(UBound(mDept3)) <> "" Then ReDim Preserve mDept3(UBound(mDept3) + 1)
186                 mDept3(UBound(mDept3)) = strTmpDept
                End If
            End If
188         rsCharge.MoveNext
        Loop

        '�в�λ���շѶ���
    
190     If strType = "D" Then
192         strSql = "Select /*+ RULE */" & vbNewLine & _
                    " i.Id As �շ�id, '[' || i.���� || ']' || i.���� As ��Ŀ��, i.���㵥λ As ��λ," & vbNewLine & _
                    " Decode(i.�Ƿ���, 1, '���', To_Char(p.�۸�)) As �۸�, Nvl(r.�շ�����, 0) As ����, Nvl(r.���ж���, 0) As �̶�,Nvl(r.�շѷ�ʽ,0) as �շѷ�ʽ," & vbNewLine & _
                    " Nvl(r.������Ŀ, 0) As ����, r.�������� As ����, d.����, r.��鲿λ As ��λ, r.��鷽�� As ����, r.������Դ, r.���ÿ���id, b.���� as ���ұ���, b.����  as ��������" & vbNewLine & _
                    "From �շ���ĿĿ¼ i," & vbNewLine & _
                    "        (   Select p.�շ�ϸĿid, Sum(p.�ּ�) As �۸�" & vbNewLine & _
                    "            From �շѼ�Ŀ p" & vbNewLine & _
                    "            Where p.ִ������ <= Sysdate And (p.��ֹ���� Is Null Or p.��ֹ���� >= Sysdate)  And p.�۸�ȼ� Is Null " & _
                    "            Group By p.�շ�ϸĿid) p, ������ĿĿ¼ c, ���Ƽ�鲿λ d, �����շѹ�ϵ r, ���ű� B" & vbNewLine & _
                    "Where r.�շ���Ŀid = i.Id And i.Id = p.�շ�ϸĿid(+) And c.�������� = d.���� And r.��鲿λ = d.���� And" & vbNewLine & _
                    "      r.���ÿ���id=b.id(+) And r.������Ŀid = c.Id And r.��鲿λ Is Not Null And r.������Ŀid = [1]" & vbNewLine & _
                    "Order By d.����, r.��鲿λ, r.��鷽��"
         
194         Set rsCharge = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngItemID, gstrPriceClass)
196         Do Until rsCharge.EOF
198             strTmp = "" & rsCharge!���� & "|" & rsCharge!��λ & "|" & rsCharge!���� & "|" & rsCharge!��Ŀ�� & "|" & _
                        rsCharge!��λ & "|" & rsCharge!�۸� & "|" & rsCharge!���� & "|" & _
                        IIf("" & rsCharge!�̶� = "1", "��", "") & "|" & rsCharge!�շѷ�ʽ & "|" & rsCharge!���� & "|" & rsCharge!�շ�ID & "|" & _
                        Val("" & rsCharge!������Դ) & "|" & Val("" & rsCharge!���ÿ���id)
            
            
200             If Val("" & rsCharge!���ÿ���id) <> 0 And InStr(strDeptList, "," & Val("" & rsCharge!���ÿ���id) & ":" & rsCharge!������Դ & ",") <= 0 Then
202                 strTmpDept = Val("" & rsCharge!���ÿ���id) & "|" & rsCharge!���ұ��� & "|" & rsCharge!��������
204                 strDeptList = strDeptList & "," & Val("" & rsCharge!���ÿ���id) & ":" & rsCharge!������Դ & ","
                Else
206                 strTmpDept = ""
                End If
            
208             If Val("" & rsCharge!������Դ) = 0 Then
210                 If mPlace0(UBound(mPlace0)) <> "" Then ReDim Preserve mPlace0(UBound(mPlace0) + 1)
212                 mPlace0(UBound(mPlace0)) = strTmp

214             ElseIf Val("" & rsCharge!������Դ) = 1 Then
216                 If mPlace1(UBound(mPlace1)) <> "" Then ReDim Preserve mPlace1(UBound(mPlace1) + 1)
218                 mPlace1(UBound(mPlace1)) = strTmp
220                 If strTmpDept <> "" Then
222                     If mDept1(UBound(mDept1)) <> "" Then ReDim Preserve mDept1(UBound(mDept1) + 1)
224                     mDept1(UBound(mDept1)) = strTmpDept
                    End If
226             ElseIf Val("" & rsCharge!������Դ) = 2 Then
228                 If mPlace2(UBound(mPlace2)) <> "" Then ReDim Preserve mPlace2(UBound(mPlace2) + 1)
230                 mPlace2(UBound(mPlace2)) = strTmp
232                 If strTmpDept <> "" Then
234                     If mDept2(UBound(mDept2)) <> "" Then ReDim Preserve mDept2(UBound(mDept2) + 1)
236                     mDept2(UBound(mDept2)) = strTmpDept
                    End If
238             ElseIf Val("" & rsCharge!������Դ) = 3 Then
240                 If mPlace3(UBound(mPlace3)) <> "" Then ReDim Preserve mPlace3(UBound(mPlace3) + 1)
242                 mPlace3(UBound(mPlace3)) = strTmp
244                 If strTmpDept <> "" Then
246                     If mDept3(UBound(mDept3)) <> "" Then ReDim Preserve mDept3(UBound(mDept3) + 1)
248                     mDept3(UBound(mDept3)) = strTmpDept
                    End If
                End If
250             rsCharge.MoveNext
            Loop
        
252         If lngFlag = 1 Then
                '���ӵ��շѶ���
254             strSql = "Select /*+ RULE */" & vbNewLine & _
                        " i.Id as �շ�ID, '[' || i.���� || ']' || i.���� As ��Ŀ��, i.���㵥λ as ��λ, Decode(i.�Ƿ���, 1, '���', To_Char(p.�۸�)) As �۸�," & vbNewLine & _
                        " Nvl(r.�շ�����, 0) As ����, Nvl(r.���ж���, 0) As �̶�, Nvl(r.������Ŀ, 0) As ����, r.�������� as ����, r.��鲿λ as ��λ, r.��鷽�� as ����, r.������Դ, r.���ÿ���id, b.���� as ���ұ���, b.����  as ��������" & vbNewLine & _
                        "From �����շѹ�ϵ r, �շ���ĿĿ¼ i," & vbNewLine & _
                        "        (Select p.�շ�ϸĿid, Sum(p.�ּ�) As �۸�" & vbNewLine & _
                        "            From �շѼ�Ŀ p" & vbNewLine & _
                        "            Where p.ִ������ <= Sysdate And (p.��ֹ���� Is Null Or p.��ֹ���� >= Sysdate) And p.�۸�ȼ� Is Null  " & _
                        "            Group By p.�շ�ϸĿid) p, ���ű� B" & vbNewLine & _
                        "Where r.�շ���Ŀid = i.Id And i.Id = p.�շ�ϸĿid(+) And r.��������=1 And r.���ÿ���id=b.id(+) And r.������Ŀid = [1]"
        
256             Set rsCharge = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngItemID, gstrPriceClass)
258             Do Until rsCharge.EOF
260                 strTmp = "" & rsCharge!��Ŀ�� & "|" & rsCharge!��λ & "|" & rsCharge!�۸� & "|" & rsCharge!���� & "|" & _
                             IIf("" & rsCharge!�̶� = "1", "��", "") & "|" & rsCharge!���� & "|" & rsCharge!�շ�ID & "|" & _
                             Val("" & rsCharge!������Դ) & "|" & Val("" & rsCharge!���ÿ���id)
                         
262                 If Val("" & rsCharge!���ÿ���id) <> 0 And InStr(strDeptList, "," & Val("" & rsCharge!���ÿ���id) & ":" & rsCharge!������Դ & ",") <= 0 Then
264                     strTmpDept = Val("" & rsCharge!���ÿ���id) & "|" & rsCharge!���ұ��� & "|" & rsCharge!��������
266                     strDeptList = strDeptList & "," & Val("" & rsCharge!���ÿ���id) & ":" & rsCharge!������Դ & ","
                    Else
268                     strTmpDept = ""
                    End If
270                 If Val("" & rsCharge!������Դ) = 0 Then
272                     If mAppend0(UBound(mAppend0)) <> "" Then ReDim Preserve mAppend0(UBound(mAppend0) + 1)
274                     mAppend0(UBound(mAppend0)) = strTmp
276                 ElseIf Val("" & rsCharge!������Դ) = 1 Then
278                     If mAppend1(UBound(mAppend1)) <> "" Then ReDim Preserve mAppend1(UBound(mAppend1) + 1)
280                     mAppend1(UBound(mAppend1)) = strTmp
282                     If strTmpDept <> "" Then
284                         If mDept1(UBound(mDept1)) <> "" Then ReDim Preserve mDept1(UBound(mDept1) + 1)
286                         mDept1(UBound(mDept1)) = strTmpDept
                        End If
288                 ElseIf Val("" & rsCharge!������Դ) = 2 Then
290                     If mAppend2(UBound(mAppend2)) <> "" Then ReDim Preserve mAppend2(UBound(mAppend2) + 1)
292                     mAppend2(UBound(mAppend2)) = strTmp
294                     If strTmpDept <> "" Then
296                         If mDept2(UBound(mDept2)) <> "" Then ReDim Preserve mDept2(UBound(mDept2) + 1)
298                         mDept2(UBound(mDept2)) = strTmpDept
                        End If
300                 ElseIf Val("" & rsCharge!������Դ) = 3 Then
302                     If mAppend3(UBound(mAppend3)) <> "" Then ReDim Preserve mAppend3(UBound(mAppend3) + 1)
304                     mAppend3(UBound(mAppend3)) = strTmp
306                     If strTmpDept <> "" Then
308                         If mDept3(UBound(mDept3)) <> "" Then ReDim Preserve mDept3(UBound(mDept3) + 1)
310                         mDept3(UBound(mDept3)) = strTmpDept
                        End If
                    End If
312                 rsCharge.MoveNext
                Loop
            End If
        End If
 

        Exit Sub
ErrHand:
314     MsgBox "ReadClinicData��" & CStr(Erl()) & "�У�" & err.Description
    '    If ErrCenter() = 1 Then
    '    Resume
    '    End If
    '    Call SaveErrLog
End Sub

Private Sub msfExseRef(ByVal lngDeptID As Long, ByVal lngCacheData As Long)
        '������ʾ��������
        'lngDeptID : ����ID����������ʾ������ҵ�����
        'lngCacheData: �Ƿ񻺴�ԭ����
    
        Dim dblTotal As Double
        Dim lngRow As Long
        Dim strGen() As String '��ͨ
        Dim strPlace() As String '��λ
        Dim strAppend() As String '����
        Dim lngListIndex As Long
    
100     err = 0: On Error GoTo ErrHand
102     If lngCacheData = 1 And mlngLastDeptID <> -1 Then
104         Call CacheData(mlngLastSource, mlngLastDeptID)
        End If
    
106     mlngLastSource = GetCurrSource
108     mlngLastDeptID = lngDeptID
    
        '���ݵ�ǰѡ��ҳ�棬ѡ������
110     ReDim strGen(0) As String
112     ReDim strPlace(0) As String
114     ReDim strAppend(0) As String
    
116     If mlngLastSource = 0 Then
118         strGen = mGen0
120         strPlace = mPlace0
122         strAppend = mAppend0
124     ElseIf mlngLastSource = 1 Then
126         strGen = mGen1
128         strPlace = mPlace1
130         strAppend = mAppend1
132     ElseIf mlngLastSource = 2 Then
134         strGen = mGen2
136         strPlace = mPlace2
138         strAppend = mAppend2
140     ElseIf mlngLastSource = 3 Then
142         strGen = mGen3
144         strPlace = mPlace3
146         strAppend = mAppend3
        End If
    
        '�����Ŀ�ſ����ü��շ�
    
    
148     stbExse.TabVisible(1) = False
150     stbExse.TabVisible(2) = False
    
152     Call IniItemList
154     If lngDeptID = -1 Or Me.cmdClose.Tag = "����" Then
156         Me.msfExse.Active = False  '��ǰҳ�治�����п��ң������޿��ң����ܱ༭
        Else
158         Me.msfExse.Active = True
        End If
160     dblTotal = 0
162     For lngRow = LBound(strGen) To UBound(strGen)
164         If strGen(lngRow) <> "" Then
166             With Me.msfExse
                    'strGen��ʽΪ�� id|����|���|���㵥λ|�۸�|����|�̶�|����|�շѷ�ʽ|������Դ|���ÿ���id
168                 If lngDeptID = 0 Or lngDeptID = Val(Split(strGen(lngRow), "|")(10)) Then
170                     If .RowData(.Rows - 1) <> 0 Then .Rows = .Rows + 1
172                     .TextMatrix(.Rows - 1, ExseCol.���) = .Rows - 1
174                     .RowData(.Rows - 1) = Split(strGen(lngRow), "|")(0) 'ID
176                     .TextMatrix(.Rows - 1, ExseCol.��Ŀ��) = Split(strGen(lngRow), "|")(1)              '����
178                     .TextMatrix(.Rows - 1, ExseCol.���) = Split(strGen(lngRow), "|")(2)                '���
180                     .TextMatrix(.Rows - 1, ExseCol.��λ) = Split(strGen(lngRow), "|")(3)                '��λ
182                     .TextMatrix(.Rows - 1, ExseCol.��ǰ��) = Split(strGen(lngRow), "|")(4)              '�۸�
184                     .TextMatrix(.Rows - 1, ExseCol.��Ӧ��) = FormatEx(Split(strGen(lngRow), "|")(5), 5) '����
186                     .TextMatrix(.Rows - 1, ExseCol.�̶�) = Split(strGen(lngRow), "|")(6)                '�̶�
188                     .TextMatrix(.Rows - 1, ExseCol.����) = Split(strGen(lngRow), "|")(7)                '����
190                     .TextMatrix(.Rows - 1, ExseCol.�շѷ�ʽ) = Split(strGen(lngRow), "|")(8)

                    
192                     dblTotal = dblTotal + Val(Split(strGen(lngRow), "|")(4)) * Val(Split(strGen(lngRow), "|")(5))
                    End If
                End With
            End If
        Next
194     txtTotal = IIf(dblTotal = 0, "", Format(dblTotal, "0.0000"))
    
196     If mstrType = "D" Then
198         stbExse.TabVisible(1) = True
            '��ʼ�����
200         Call vfgExseRef(strPlace, lngDeptID)
    
202         If mlngFlag = 1 Then
204             stbExse.TabVisible(2) = True
206             Call vfg����Ref(strAppend, lngDeptID)
            End If
        End If
        
        If mstrOper = "����" And mstr��� = "D" Then
            stbExse.TabVisible(1) = False
        End If
        
        Exit Sub
ErrHand:
208     MsgBox "msfExseRef��" & CStr(Erl()) & "�У�" & err.Description
    '    If ErrCenter() = 1 Then Resume
    '    Call SaveErrLog
End Sub

Private Sub vfgExseRef(ByRef strPla() As String, ByVal lngDeptID As Long)
        Dim strSql As String, dblTotal As Double
        Dim lngRow As Long
100     err = 0: On Error GoTo ErrHand
102     dblTotal = Val(txtTotal)
104     If lngDeptID = -1 Or Me.cmdClose.Tag = "����" Then
106         Me.vfgExse.Enabled = False '��ǰҳ�治�����п��ң������޿��ң����ܱ༭
        Else
108         Me.vfgExse.Enabled = True
        End If
110     With vfgExse
            '��ʼ�����
112         .Clear
114         .FixedCols = 0: .FixedRows = 1
116         .Rows = 1: .Cols = 11
        
118         .MergeRow(0) = True
120         .MergeCellsFixed = flexMergeRestrictColumns
    '
122         .MergeCol(0) = True ': .MergeCol(1) = True
124         .MergeCells = flexMergeRestrictColumns
        
126         .RowHeightMin = 300
        
128         .TextMatrix(0, 0) = "��λ": .TextMatrix(0, 1) = "��λ": .TextMatrix(0, 2) = "����": .TextMatrix(0, 3) = "��Ŀ��"
130         .TextMatrix(0, 4) = "��λ": .TextMatrix(0, 5) = "�۸�": .TextMatrix(0, 6) = "����"
132         .TextMatrix(0, 7) = "�̶�": .TextMatrix(0, 8) = "�շѷ�ʽ": .TextMatrix(0, 9) = "����": .TextMatrix(0, 10) = "�շ�ID":
134         .ColKey(0) = "����": .ColKey(1) = "��λ": .ColKey(2) = "����": .ColKey(3) = "��Ŀ��"
136         .ColKey(4) = "��λ": .ColKey(5) = "�۸�"
138         .ColKey(6) = "����": .ColKey(7) = "�̶�": .ColKey(8) = "�շѷ�ʽ": .ColKey(9) = "����": .ColKey(10) = "�շ�id"
        
140         .ColHidden(.ColIndex("����")) = False
142         .ColHidden(.ColIndex("��λ")) = False: .ColHidden(.ColIndex("����")) = False: .ColHidden(.ColIndex("��Ŀ��")) = False
144         .ColHidden(.ColIndex("��λ")) = False: .ColHidden(.ColIndex("�۸�")) = False: .ColHidden(.ColIndex("����")) = False
146         .ColHidden(.ColIndex("�̶�")) = False: .ColHidden(.ColIndex("����")) = True: .ColHidden(.ColIndex("�շ�id")) = True
            .ColHidden(.ColIndex("�շѷ�ʽ")) = False
148         .ColWidth(.ColIndex("����")) = 900
150         .ColWidth(.ColIndex("��λ")) = 1000: .ColWidth(.ColIndex("����")) = 1000: .ColWidth(.ColIndex("��Ŀ��")) = 2200
152         .ColWidth(.ColIndex("��λ")) = 450: .ColWidth(.ColIndex("�۸�")) = 800: .ColWidth(.ColIndex("����")) = 800
154         .ColWidth(.ColIndex("�̶�")) = 450: .ColWidth(.ColIndex("����")) = 0: .ColWidth(.ColIndex("�շ�id")) = 0
            .ColWidth(.ColIndex("�շѷ�ʽ")) = 1400
156         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
158         .WordWrap = True
160         .AutoResize = True
        
162         .ColComboList(.ColIndex("��λ")) = "..."
164         .ColComboList(.ColIndex("����")) = "..."
166         .ColComboList(.ColIndex("��Ŀ��")) = "..."
        
168         .Editable = flexEDKbdMouse
        
        
170         For lngRow = LBound(strPla) To UBound(strPla)
172             If strPla(lngRow) <> "" Then
174                 If lngDeptID = 0 Or lngDeptID = Val(Split(strPla(lngRow), "|")(12)) Then
176                     .Rows = .Rows + 1
                        'strPla��ʽ������|��λ|����|��Ŀ��|��λ|�۸�|����|�շѷ�ʽ|�̶�|����|�շ�ID|������Դ|���ÿ���id
178                     .TextMatrix(.Rows - 1, .ColIndex("����")) = Split(strPla(lngRow), "|")(0)
180                     .TextMatrix(.Rows - 1, .ColIndex("��λ")) = Split(strPla(lngRow), "|")(1)
182                     .TextMatrix(.Rows - 1, .ColIndex("����")) = Split(strPla(lngRow), "|")(2)
184                     .TextMatrix(.Rows - 1, .ColIndex("��Ŀ��")) = Split(strPla(lngRow), "|")(3)
186                     .TextMatrix(.Rows - 1, .ColIndex("��λ")) = Split(strPla(lngRow), "|")(4)
                    
188                     .TextMatrix(.Rows - 1, .ColIndex("�۸�")) = Format(Val(Split(strPla(lngRow), "|")(5)), "0.00")
190                     .TextMatrix(.Rows - 1, .ColIndex("����")) = Val(Split(strPla(lngRow), "|")(6))
                    
192                     .TextMatrix(.Rows - 1, .ColIndex("�̶�")) = Split(strPla(lngRow), "|")(7)
                        .TextMatrix(.Rows - 1, .ColIndex("�շѷ�ʽ")) = IIf(0 = Val(Split(strPla(lngRow), "|")(8)), "0-������ȡ", "9-�Զ���")
194                     .TextMatrix(.Rows - 1, .ColIndex("����")) = Val(Split(strPla(lngRow), "|")(9))
196                     .TextMatrix(.Rows - 1, .ColIndex("�շ�ID")) = Val(Split(strPla(lngRow), "|")(10))
                    
198                     dblTotal = dblTotal + Val(Split(strPla(lngRow), "|")(5)) * Val(Split(strPla(lngRow), "|")(6))
                    End If
                End If
            Next
200         If .Rows < 2 Then .Rows = .Rows + 1
202         .AutoSizeMode = flexAutoSizeRowHeight
204         .AutoSize .ColIndex("����"), .ColIndex("����")
        End With
    
206     txtTotal = IIf(Val(dblTotal) = 0, "", Format(dblTotal, "0.00"))
        Exit Sub
ErrHand:
208     MsgBox "vfgExseRef��" & CStr(Erl()) & "�У�" & err.Description
    '    If ErrCenter() = 1 Then Resume
    '    Call SaveErrLog
End Sub

Private Sub vfgExse_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If InStr("," & vfgExse.ColIndex("��Ŀ��") & ",", "," & Col & ",") > 0 Then
        Call vfgExse_CellButtonClick(Row, Col)
    End If
End Sub

Private Sub vfgExse_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vfgExse
        If NewCol = .ColIndex("�շѷ�ʽ") Then
            .ComboList = "0-������ȡ|9-�Զ���"
        Else
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vfgExse_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgExse
        If InStr("," & .ColIndex("��λ") & "," & .ColIndex("����") & "," & .ColIndex("��Ŀ��") & "," & .ColIndex("����") & "," & .ColIndex("�շѷ�ʽ") & ",", "," & Col & ",") <= 0 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vfgExse_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim pt As POINTAPI, strReturn As String
    Dim strSql As String, str��λ As String, bytType As Byte
    Dim rsTmp As ADODB.Recordset, varReturn As Variant, lngRow As Long, curOld��� As Currency
    
    On Error GoTo ErrHandle
    With vfgExse
        'ȡ�õ�ǰ�е�λ��
        pt.x = .ColPos(Col) \ Screen.TwipsPerPixelX
        pt.y = (.RowPos(Row) + .RowHeight(Row)) \ Screen.TwipsPerPixelY
        ClientToScreen .hWnd, pt
        
        Select Case Col
            Case .ColIndex("��λ")
'                strSQL = "Select Distinct '��ѡ' As ����, ��λ" & vbNewLine & _
'                        "From ������Ŀ��λ a" & vbNewLine & _
'                        "Where ��Ŀid = [1]" & vbNewLine & _
'                        "Union All" & vbNewLine & _
'                        "Select '��ѡ' As ����, ���� As ��λ" & vbNewLine & _
'                        "From ���Ƽ�鲿λ" & vbNewLine & _
'                        "Where ���� = (Select �������� From ������ĿĿ¼ Where Id = [1]) And" & vbNewLine & _
'                        "           ���� Not In (Select ��λ From ������Ŀ��λ Where ��Ŀid = [1])"
                strSql = "Select Distinct Decode(C.��λ,Null,'��ѡ','��ѡ') As ����,A.���� as ���, A.���� As ��λ" & vbNewLine & _
                        "From ���Ƽ�鲿λ A,������ĿĿ¼ B,(Select ����,��λ From ������Ŀ��λ C Where ��ĿID=[1]) C" & vbNewLine & _
                        "Where A.���� = B.�������� And B.Id=[1] And A.����=C.����(+) And A.����=C.��λ(+) Order by ����,���"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(lblItem.Tag))
                bytType = 1
            Case .ColIndex("����")
                str��λ = vfgExse.TextMatrix(Row, .ColIndex("��λ"))
                If str��λ <> "" Then
                    strSql = "Select '��ѡ' As ����, ����" & vbNewLine & _
                            "From ������Ŀ��λ" & vbNewLine & _
                            "Where ��λ = [2] And ��Ŀid = [1]" & vbNewLine & _
                            "Union All" & vbNewLine & _
                            "Select '��ѡ' As ����, ���� As ��λ" & vbNewLine & _
                            "From ���Ƽ�鲿λ" & vbNewLine & _
                            "Where ���� = [2] And ���� = (Select �������� From ������ĿĿ¼ Where Id = [1])"
                Else
                    MsgBox "��ѡ��λ", vbInformation, gstrSysName
                    vfgExse.Select Row, .ColIndex("��λ")
                    Exit Sub
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(lblItem.Tag), str��λ)
                bytType = 2
            Case .ColIndex("��Ŀ��")
                strSql = "select distinct I.ID,Rpad('['||I.����||']'||I.����||' '||I.���,60) as ����,I.���㵥λ as ��λ" & _
                        " from �շ���ĿĿ¼ I,�շ���Ŀ���� N" & _
                        " where I.ID=N.�շ�ϸĿid and I.��� not in ('1','J')" & _
                        "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                        "       and (I.���� like [1] " & _
                        "           or N.���� like [2] " & _
                        "           or N.���� like [2])"
                
                strTemp = UCase(.TextMatrix(.Row, .Col))
                If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
                bytType = 3
        End Select
        If rsTmp.BOF Or rsTmp.EOF Then
            If bytType = 1 Or bytType = 2 Then
                Me.vfgExse.TextMatrix(Row, Col) = ""
                lblMessage.Caption = lblMessage.Tag & "δ�ҵ���Ŀ��"
            Else
                Me.vfgExse.TextMatrix(Row, Col) = ""
                lblMessage.Caption = lblMessage.Tag & "δ�ҵ�ָ���շ���Ŀ��"
            End If
            Exit Sub
        End If
        
        frmClinicExseVsSelect.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
        
        Call frmClinicExseVsSelect.ShowSelect(bytType, rsTmp, strReturn)
        
        If InStr(strReturn, "|") > 0 Then
            varReturn = Split(strReturn, "|")
            If bytType = 1 Then
                '��λ
                .TextMatrix(Row, Col - 1) = varReturn(1)
                .TextMatrix(Row, Col) = varReturn(2)
                .TextMatrix(Row, Col + 1) = ""
                .Select Row, Col + 1
            ElseIf bytType = 2 Then
                '����
                .TextMatrix(Row, Col) = varReturn(1)
                '11295 Ҫ��һ���������Զ�Ӧ����շ���Ŀ
'                For lngRow = .FixedRows To .Rows - 1
'                    If lngRow <> Row And .TextMatrix(lngRow, Col) = varReturn(1) And _
'                       .TextMatrix(lngRow, .ColIndex("��λ")) = .TextMatrix(Row, .ColIndex("��λ")) Then
'                        lblMessage.Caption = lblMessage.Tag & "ÿ�ַ���ֻ�ܶ�Ӧһ���շ���Ŀ��"
'                        .TextMatrix(Row, Col) = ""
'                    End If
'                Next
                .Select Row, Col + 1
            Else
                '��Ŀ
                .TextMatrix(Row, .ColIndex("�շ�ID")) = Val("" & varReturn(0))
                .TextMatrix(Row, .ColIndex("��Ŀ��")) = "" & varReturn(1)
                .TextMatrix(Row, .ColIndex("��λ")) = "" & varReturn(2)
                
                strSql = "select decode(I.�Ƿ���,1,'���',to_char(P.�۸�)) As �۸�" & _
                " from (select �Ƿ��� from �շ���ĿĿ¼ where id=[1]) I," & _
                "      (Select sum(�ּ�) As �۸�" & _
                "      From �շѼ�Ŀ " & _
                "      Where �շ�ϸĿid=[1] " & _
                "           and ִ������<=Sysdate And (��ֹ���� Is Null Or ��ֹ����>=Sysdate)  And �۸�ȼ� Is Null " & ") P"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val("" & varReturn(0)), gstrPriceClass)
                
                curOld��� = Val(.TextMatrix(Row, .ColIndex("�۸�"))) * Val(.TextMatrix(Row, .ColIndex("����")))
                
                If rsTmp.RecordCount > 0 Then
                    If rsTmp.RecordCount > 1 Then
                        lblMessage.Caption = lblMessage.Tag & "����ļ۸���ڶ��������Ŀ������ѡ��"
                        .TextMatrix(Row, .ColIndex("�۸�")) = ""
                    Else
                        .TextMatrix(Row, .ColIndex("�۸�")) = IIf("" & rsTmp.Fields("�۸�") <> "���", Format(Val("" & rsTmp.Fields("�۸�")), "0.00"), "���")
                    End If
                Else
                    .TextMatrix(Row, .ColIndex("�۸�")) = ""
                End If
                
                txtTotal = Val(txtTotal) - curOld��� + Val(.TextMatrix(Row, .ColIndex("����"))) * Val(.TextMatrix(Row, .ColIndex("�۸�")))
                If Val(txtTotal) = 0 Then
                    txtTotal = ""
                Else
                    txtTotal = Format(txtTotal, "0.0000")
                End If
                .Select Row, .ColIndex("����")
            End If
            .AutoSize .ColIndex("����"), .ColIndex("����")
            
        End If
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgExse_DblClick()
    If vfgExse.Col = vfgExse.ColIndex("�̶�") Then
        Call vfgExse_KeyDown(vbKeySpace, 0)
    End If
End Sub

Private Sub vfgExse_EnterCell()
    Dim blnOk As Boolean
    
    With vfgExse
        If InStr("," & .ColIndex("��λ") & "," & .ColIndex("����") & "," & .ColIndex("��Ŀ��") & "," & .ColIndex("����") & ",", "," & .Col & ",") > 0 Then
            blnOk = True
        End If
    End With
    On Error Resume Next
    If blnOk And vfgExse.Row > 0 Then
        Call vfgExse.CellBorder(vfgExse.GridColor, 1, 1, 2, 2, 0, 0)
    End If
End Sub

Private Sub vfgExse_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancle As Boolean
    
    On Error Resume Next
    
    If KeyCode = vbKeyDelete Then
        With vfgExse
        
            txtTotal = Val(txtTotal) - Val(.TextMatrix(.Row, .ColIndex("����"))) * Val(.TextMatrix(.Row, .ColIndex("�۸�")))
            If Val(txtTotal) = 0 Then
                txtTotal = ""
            Else
                txtTotal = Format(txtTotal, "0.0000")
            End If
            If .Row > .FixedRows And .Row <= .Rows - 1 Then
                vfgExse.RemoveItem vfgExse.Row
            ElseIf .Row = .FixedRows Then
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            End If
        End With
    ElseIf KeyCode = vbKeyReturn Then
        
        With vfgExse
            If .EditText = "" Then
                KeyCode = 0
                If .Col = .ColIndex("�շѷ�ʽ") Then
                    If .Row = vfgExse.Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Select .Row + 1, .ColIndex("��λ")
                Else
                    If .Col < .Cols Then
                        If .Col = .ColIndex("��Ŀ��") Then
                            .Select .Row, .ColIndex("����")
                        Else
                            .Select .Row, .Col + 1
                        End If
                    End If
                End If
            End If
        End With
    ElseIf KeyCode = vbKeySpace Then
        With vfgExse
        If .Col = .ColIndex("�̶�") Then
            If .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "��"
            Else
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
        End With
    ElseIf KeyCode = vbKeyEscape Then
        With vfgExse
            If .Col = .ColIndex("��Ŀ��") Then
                If .ColComboList(.Col) <> "" Then
                    KeyCode = 0
                End If
            End If
        End With
    ElseIf InStr("," & vfgExse.ColIndex("��Ŀ��") & ",", "," & vfgExse.Col & ",") > 0 Then
        If vfgExse.ColComboList(vfgExse.Col) <> "" Then
            vfgExse.Tag = vfgExse.TextMatrix(vfgExse.Row, vfgExse.Col)
            vfgExse.ColComboList(vfgExse.Col) = ""
        End If
    End If
End Sub

Private Sub vfgExse_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Integer, blnCancle As Boolean
    If InStr("," & vfgExse.ColIndex("��Ŀ��") & ",", "," & Col & ",") > 0 And KeyCode = vbKeyReturn Then
        vfgExse.ColComboList(vfgExse.Col) = "..."
        
    ElseIf vfgExse.ColIndex("����") = vfgExse.Col And KeyCode <> vbKeyReturn Then
        If InStr("01234567890.", Chr(KeyCode)) <= 0 Then
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyEscape Then
        vfgExse.TextMatrix(Row, Col) = vfgExse.Tag
        vfgExse.ColComboList(vfgExse.Col) = "..."
'    ElseIf vbKeyReturn Then
'        Call vfgExse_CellButtonClick(Row, Col)
    End If
End Sub

Private Sub vfgExse_LeaveCell()
    Dim blnOk As Boolean
    
    With vfgExse
        If InStr("," & .ColIndex("��λ") & "," & .ColIndex("����") & "," & .ColIndex("��Ŀ��") & "," & .ColIndex("����") & ",", "," & .Col & ",") > 0 Then
            blnOk = True
        End If
    End With
    If InStr("," & vfgExse.ColIndex("��λ") & "," & vfgExse.ColIndex("����") & "," & vfgExse.ColIndex("��Ŀ��") & ",", "," & vfgExse.Col & ",") > 0 Then
        vfgExse.ColComboList(vfgExse.Col) = "..."
    End If
    On Error Resume Next
    If blnOk And vfgExse.Row > 0 Then
        Call vfgExse.CellBorder(vfgExse.GridColor, 0, 0, 0, 0, 0, 0)
    End If
    
End Sub

Private Sub vfgExse_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str���� As String
    
    If Col = vfgExse.ColIndex("����") Then
        With vfgExse
            txtTotal = Val(txtTotal) - Val(.TextMatrix(Row, Col)) * Val(.TextMatrix(Row, .ColIndex("�۸�"))) _
                       + Val(.EditText) * Val(.TextMatrix(Row, .ColIndex("�۸�")))
        End With
        If Val(txtTotal) = 0 Then
            txtTotal = ""
        Else
            txtTotal = Format(txtTotal, "0.0000")
        End If
    End If
    
    If vfgExse.ColIndex("��λ") = Col Then
        If vfgExse.TextMatrix(Row, Col) <> "" Then
            mrsBwff.Filter = " ��λ='" & vfgExse.EditText & "'"
            If mrsBwff.RecordCount <= 0 Then
                lblMessage.Caption = lblMessage.Tag & "��λ�������顣"
                Cancel = True
            End If
        End If
    End If
    
    If vfgExse.ColIndex("����") = Col Then
        If vfgExse.TextMatrix(Row, Col) <> "" Then
            mrsBwff.Filter = " ��λ='" & vfgExse.TextMatrix(Row, vfgExse.ColIndex("��λ")) & "'"
            If mrsBwff.RecordCount <= 0 Then
                lblMessage.Caption = lblMessage.Tag & "��λ���󣬼�顣"
                Cancel = True
             Else
                str���� = mrsBwff.Fields("����")
                str���� = Replace(str����, vbTab, "|")
                str���� = Replace(str����, ",", "|")
                str���� = Replace(str����, ";", "|")
                str���� = Replace(str����, "0", "")
                str���� = Replace(str����, "1", "")
                
                If InStr("|" & str���� & "|", "|" & vfgExse.EditText & "|") <= 0 Then
                    lblMessage.Caption = lblMessage.Tag & "�����������顣"
                    Cancel = True
                End If
            End If
        End If
    End If
End Sub

'----------------- ����
Private Sub vfg����Ref(ByRef strAppend() As String, ByVal lngDeptID As Long)
        Dim strSql As String, dblTotal As Double
        Dim lngRow As Long
100     err = 0: On Error GoTo ErrHand
102     If lngDeptID = -1 Or Me.cmdClose.Tag = "����" Then
104         Me.vfgExse.Enabled = False '��ǰҳ�治�����п��ң������޿��ң����ܱ༭
        Else
106         Me.vfgExse.Enabled = True
        End If
108     dblTotal = Val(txtTotal)
110     With vfg����
            '��ʼ�����
112         .Clear
114         .FixedCols = 0: .FixedRows = 1
116         .Rows = 1: .Cols = 7
        
118         .RowHeightMin = 300
        
120         .TextMatrix(0, 0) = "��Ŀ��": .TextMatrix(0, 1) = "��λ": .TextMatrix(0, 2) = "�۸�": .TextMatrix(0, 3) = "����"
122         .TextMatrix(0, 4) = "�̶�": .TextMatrix(0, 5) = "����": .TextMatrix(0, 6) = "�շ�ID"
124         .ColKey(0) = "��Ŀ��": .ColKey(1) = "��λ": .ColKey(2) = "�۸�"
126         .ColKey(3) = "����": .ColKey(4) = "�̶�": .ColKey(5) = "����": .ColKey(6) = "�շ�id"
        
128         .ColHidden(.ColIndex("��Ŀ��")) = False
130         .ColHidden(.ColIndex("��λ")) = False: .ColHidden(.ColIndex("�۸�")) = False: .ColHidden(.ColIndex("����")) = False
132         .ColHidden(.ColIndex("�̶�")) = False: .ColHidden(.ColIndex("����")) = True: .ColHidden(.ColIndex("�շ�id")) = True
        
134         .ColWidth(.ColIndex("��Ŀ��")) = 5000
136         .ColWidth(.ColIndex("��λ")) = 450: .ColWidth(.ColIndex("�۸�")) = 800: .ColWidth(.ColIndex("����")) = 800
138         .ColWidth(.ColIndex("�̶�")) = 450: .ColWidth(.ColIndex("����")) = 0: .ColWidth(.ColIndex("�շ�id")) = 0
        
140         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
142         .WordWrap = True
144         .AutoResize = True
        
146         .ColComboList(.ColIndex("��Ŀ��")) = "..."
        
148         .Editable = flexEDKbdMouse
        
150         For lngRow = LBound(strAppend) To UBound(strAppend)
152             If strAppend(lngRow) <> "" Then
154                 If lngDeptID = 0 Or lngDeptID = Val(Split(strAppend(lngRow), "|")(8)) Then
156                     .Rows = .Rows + 1
                        'strAppend��ʽ����Ŀ��|��λ|�۸�|����|�̶�|����|�շ�ID|������Դ|���ÿ���id
158                     .TextMatrix(.Rows - 1, .ColIndex("��Ŀ��")) = Split(strAppend(lngRow), "|")(0)
160                     .TextMatrix(.Rows - 1, .ColIndex("��λ")) = Split(strAppend(lngRow), "|")(1)
                    
162                     .TextMatrix(.Rows - 1, .ColIndex("�۸�")) = Format(Val(Split(strAppend(lngRow), "|")(2)), "0.00")
164                     .TextMatrix(.Rows - 1, .ColIndex("����")) = Val(Split(strAppend(lngRow), "|")(3))
                    
166                     .TextMatrix(.Rows - 1, .ColIndex("�̶�")) = Split(strAppend(lngRow), "|")(4)
168                     .TextMatrix(.Rows - 1, .ColIndex("����")) = Val(Split(strAppend(lngRow), "|")(5))
170                     .TextMatrix(.Rows - 1, .ColIndex("�շ�ID")) = Val(Split(strAppend(lngRow), "|")(6))
                    
172                     dblTotal = dblTotal + Val(Split(strAppend(lngRow), "|")(2)) * Val(Split(strAppend(lngRow), "|")(3))
                    End If
                End If
            Next
174         If .Rows < 2 Then .Rows = .Rows + 1
176         .AutoSizeMode = flexAutoSizeRowHeight
178         .AutoSize .ColIndex("��Ŀ��")
        End With
180         txtTotal = IIf(Val(dblTotal) = 0, "", Format(dblTotal, "0.00"))
        Exit Sub
ErrHand:
182     MsgBox "vfg����Ref��" & CStr(Erl()) & "�У�" & err.Description
    '    If ErrCenter() = 1 Then Resume
    '    Call SaveErrLog
End Sub

Private Sub vfg����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vfg����.ColIndex("��Ŀ��") = Col Then
        Call vfg����_CellButtonClick(Row, Col)
    End If
End Sub

Private Sub vfg����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfg����
        If InStr("," & .ColIndex("��Ŀ��") & "," & .ColIndex("����") & ",", "," & Col & ",") <= 0 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vfg����_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim pt As POINTAPI, strReturn As String
    Dim strSql As String, curOld��� As Currency
    Dim rsTmp As ADODB.Recordset, varReturn As Variant, lngRow As Long
    
    On Error GoTo ErrHandle
    With vfg����
        'ȡ�õ�ǰ�е�λ��
        pt.x = .ColPos(Col) \ Screen.TwipsPerPixelX
        pt.y = (.RowPos(Row) + .RowHeight(Row)) \ Screen.TwipsPerPixelY
        ClientToScreen .hWnd, pt
        
        If Col = .ColIndex("��Ŀ��") Then
            strSql = "select distinct I.ID,Rpad('['||I.����||']'||I.����||' '||I.���,60) as ����,I.���㵥λ as ��λ" & _
                    " from �շ���ĿĿ¼ I,�շ���Ŀ���� N" & _
                    " where I.ID=N.�շ�ϸĿid and I.��� not in ('1','J')" & _
                    "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                    "       and (I.���� like [1] " & _
                    "           or N.���� like [2] " & _
                    "           or N.���� like [2])"
            strTemp = UCase(.TextMatrix(.Row, .Col))
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        End If
        
        If rsTmp.BOF Or rsTmp.EOF Then
            .TextMatrix(Row, Col) = ""
            lblMessage.Caption = lblMessage.Tag & "δ�ҵ�ָ���շ���Ŀ��"
            Exit Sub
        End If
        
        frmClinicExseVsSelect.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
        
        Call frmClinicExseVsSelect.ShowSelect(3, rsTmp, strReturn)
        
        If InStr(strReturn, "|") > 0 Then
            varReturn = Split(strReturn, "|")
            '��Ŀ
            .TextMatrix(Row, .ColIndex("�շ�ID")) = Val("" & varReturn(0))
            .TextMatrix(Row, .ColIndex("��Ŀ��")) = "" & varReturn(1)
            .TextMatrix(Row, .ColIndex("��λ")) = "" & varReturn(2)
            
            strSql = "select decode(I.�Ƿ���,1,'���',to_char(P.�۸�)) As �۸�" & _
            " from (select �Ƿ��� from �շ���ĿĿ¼ where id=[1]) I," & _
            "      (Select sum(�ּ�) As �۸�" & _
            "      From �շѼ�Ŀ " & _
            "      Where �շ�ϸĿid=[1] " & _
            "           and ִ������<=Sysdate And (��ֹ���� Is Null Or ��ֹ����>=Sysdate)  And �۸�ȼ� Is Null " & ") P"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val("" & varReturn(0)), gstrPriceClass)
            curOld��� = Val(.TextMatrix(Row, .ColIndex("�۸�"))) * Val(.TextMatrix(Row, .ColIndex("����")))
            If rsTmp.RecordCount > 0 Then
                If rsTmp.RecordCount > 1 Then
                    lblMessage.Caption = lblMessage.Tag & "����ļ۸���ڶ��������Ŀ������ѡ��"
                    .TextMatrix(Row, .ColIndex("�۸�")) = ""
                Else
                    .TextMatrix(Row, .ColIndex("�۸�")) = IIf("" & rsTmp.Fields("�۸�") <> "���", Format(Val("" & rsTmp.Fields("�۸�")), "0.00"), "���")
                End If
            Else
                .TextMatrix(Row, .ColIndex("�۸�")) = ""
            End If
            
            txtTotal = Val(txtTotal) - curOld��� + Val(.TextMatrix(Row, .ColIndex("����"))) * Val(.TextMatrix(Row, .ColIndex("�۸�")))
            If Val(txtTotal) = 0 Then
                txtTotal = ""
            Else
                txtTotal = Format(txtTotal, "0.0000")
            End If
            
            .Select Row, .ColIndex("����")
        End If
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfg����_DblClick()
    If vfg����.Col = vfg����.ColIndex("�̶�") Then
        Call vfg����_KeyDown(vbKeySpace, 0)
    End If
End Sub

Private Sub vfg����_EnterCell()
    Dim blnOk As Boolean
    
    With vfg����
        If InStr("," & .ColIndex("��Ŀ��") & "," & .ColIndex("����") & ",", "," & .Col & ",") > 0 Then
            blnOk = True
        End If
    End With
    On Error Resume Next
    If blnOk And vfg����.Row > 0 Then
        Call vfg����.CellBorder(vfg����.GridColor, 1, 1, 2, 2, 0, 0)
    End If
End Sub

Private Sub vfg����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancle As Boolean
    If KeyCode = vbKeyDelete Then
        With vfg����
        
            txtTotal = Val(txtTotal) - Val(.TextMatrix(.Row, .ColIndex("����"))) * Val(.TextMatrix(.Row, .ColIndex("�۸�")))
            If Val(txtTotal) = 0 Then
                txtTotal = ""
            Else
                txtTotal = Format(txtTotal, "0.0000")
            End If
            If .Row > 1 And .Row < .Rows - 1 Then
                .RemoveItem .Row
            Else
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            End If
        End With
    ElseIf KeyCode = vbKeyReturn Then
        
        With vfg����
            If .EditText = "" Then
                KeyCode = 0
                If .Col = .ColIndex("�̶�") Then
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Select .Row + 1, .ColIndex("��Ŀ��")
                Else
                    If .Col < .Cols Then
                        If .Col = .ColIndex("��Ŀ��") Then
                            .Select .Row, .ColIndex("����")
                        Else
                            .Select .Row, .Col + 1
                        End If
                    End If
                End If
            End If
        End With
    ElseIf KeyCode = vbKeySpace Then
        With vfg����
        If .Col = .ColIndex("�̶�") Then
            If .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "��"
            Else
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
        End With
    ElseIf KeyCode = vbKeyEscape Then
        With vfg����
            If .Col = .ColIndex("��Ŀ��") Then
                If .ColComboList(.Col) <> "" Then
                    KeyCode = 0
                End If
            End If
        End With
    ElseIf vfg����.ColIndex("��Ŀ��") = vfg����.Col Then
        If vfg����.ColComboList(vfg����.Col) <> "" Then
            vfg����.ColComboList(vfg����.Col) = ""
        End If
    End If
End Sub

Private Sub vfg����_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Integer, blnCancle As Boolean
    If vfg����.ColIndex("��Ŀ��") = Col And KeyCode = vbKeyReturn Then
        vfg����.ColComboList(vfg����.Col) = "..."
    
    ElseIf vfg����.ColIndex("����") = vfg����.Col And KeyCode <> vbKeyReturn Then
        If InStr("01234567890.", Chr(KeyCode)) <= 0 Then
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyEscape Then
        vfg����.TextMatrix(Row, Col) = vfg����.Tag
        vfg����.ColComboList(vfg����.Col) = "..."
    End If
End Sub

Private Sub vfg����_LeaveCell()
    Dim blnOk As Boolean
    
    With vfg����
        If InStr("," & .ColIndex("��Ŀ��") & "," & .ColIndex("����") & ",", "," & .Col & ",") > 0 Then
            blnOk = True
        End If
    End With
    If vfg����.ColIndex("��Ŀ��") = vfg����.Col Then
        vfg����.ColComboList(vfg����.Col) = "..."
    End If
    On Error Resume Next
    If blnOk And vfg����.Row > 0 Then
        Call vfg����.CellBorder(vfg����.GridColor, 0, 0, 0, 0, 0, 0)
    End If
    
End Sub


Private Sub vfg����_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str���� As String
    
    If Col = vfg����.ColIndex("����") Then
        With vfg����
            txtTotal = Val(txtTotal) - Val(.TextMatrix(Row, Col)) * Val(.TextMatrix(Row, .ColIndex("�۸�"))) _
                       + Val(.EditText) * Val(.TextMatrix(Row, .ColIndex("�۸�")))
        End With
        If Val(txtTotal) = 0 Then
            txtTotal = ""
        Else
            txtTotal = Format(txtTotal, "0.0000")
        End If
    End If
    
End Sub

Private Sub ResizeTabDept()
    '����stbExse�ؼ��Ĵ�С
    On Error Resume Next
    With tabDept
        If .SelectedItem.Index = 1 Then
            '����ʾ����
            fraDept.Visible = False
            stbExse.Left = .Left + 90
            stbExse.Top = .Top + 400
            
            stbExse.Width = .Width - 180
            stbExse.Height = .Height - 500
        Else
            fraDept.Visible = True
            
            fraDept.Left = .Left + 90
            fraDept.Top = .Top + 400
            fraDept.Height = .Height - 500
            
            stbExse.Left = fraDept.Left + fraDept.Width + 45
            stbExse.Width = .Width - fraDept.Width - 230
            stbExse.Top = fraDept.Top
            stbExse.Height = fraDept.Height
        End If
        ResizeStbExse
    End With
End Sub

Private Sub ResizeStbExse()
    On Error Resume Next
    With stbExse
            msfExse.Left = 90
            msfExse.Top = 325
            msfExse.Width = .Width - 180
            msfExse.Height = .Height - 410
            
            vfgExse.Left = 90
            vfgExse.Top = 325
            vfgExse.Width = .Width - 180
            vfgExse.Height = .Height - 410
            
            vfg����.Left = 90
            vfg����.Top = 325
            vfg����.Width = .Width - 180
            vfg����.Height = .Height - 410
    End With
End Sub

Private Sub DeptSelect(ByVal strInput As String)
    'ѡ�����
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strDeptList As String, strReturn As String
    Dim strInquiry As String, i As Integer, lngSource As Long
    On Error GoTo hErr
        lngSource = GetCurrSource
        If Trim(strInput) <> "" Then
            strInquiry = gstrMatch & UCase(strInput) & "%"
        End If
        If lstDept.ListCount > 0 Then
            For i = 0 To lstDept.ListCount - 1
                strDeptList = strDeptList & "," & lstDept.ItemData(i)
            Next
        End If
        If lngSource = 1 Or lngSource = 3 Then
            '�����������
            strSql = "Select Distinct a.����, a.����, a.ID" & vbNewLine & _
                    "From ���ű� A, ��������˵�� D" & vbNewLine & _
                    "Where a.Id = d.����id And (d.������� = 1 Or d.������� = 3) and d.�������� in ('�ٴ�','���','����','����','����','����','Ӫ��') " & vbNewLine & _
                    " And (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                    IIf(strDeptList = "", "", "   And Instr([1], ',' || a.Id || ',') <= 0") & vbNewLine & _
                    IIf(strInquiry = "", "", " And (a.���� Like [2] Or a.���� Like [2] Or a.���� Like [2]) ") & _
                    "Order By ����, ����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptList & ",", strInquiry)
        ElseIf lngSource = 2 Then
            'סԺ����
            strSql = "Select Distinct a.����, a.����, a.ID" & vbNewLine & _
                    "From ���ű� A, ��������˵�� D" & vbNewLine & _
                    "Where a.Id = d.����id And (d.������� = 2 Or d.������� = 3) And d.�������� in ('����','���','����','����','����','����','Ӫ��') " & vbNewLine & _
                    " And (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                    IIf(strDeptList = "", "", "   And Instr([1], ',' || a.Id || ',') <= 0") & vbNewLine & _
                    IIf(strInquiry = "", "", " And (a.���� Like [2] Or a.���� Like [2] Or a.���� Like [2]) ") & _
                    "Order By ����, ����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptList & ",", strInquiry)
        End If
        
        strReturn = ""
        If Not rsTmp.EOF Then
            If rsTmp.RecordCount > 0 Then
                strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "����,1200,0,2;����,1800,0,2;ID,0,1,2", "ѡ�����", True, , , 3000 + 2000)
            Else
                strReturn = rsTmp!ID
            End If
        End If
        
        If strReturn <> "" Then
            txtDept = Split(strReturn, ",")(1) & "(" & Split(strReturn, ",")(0) & ")"
            
            lstDept.AddItem Split(strReturn, ",")(1) & "(" & Split(strReturn, ",")(0) & ")"
            lstDept.ItemData(lstDept.NewIndex) = Split(strReturn, ",")(2)
            lstDept.ListIndex = lstDept.NewIndex
            txtDept.Text = ""
        Else
            If rsTmp.RecordCount = 0 Then
                MsgBox "û���ҵ����õĿ��ҡ�", vbInformation, Me.Caption
                Call zlControl.TxtSelAll(txtDept)
            End If
        End If

    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function CacheData(ByVal lngSource As Long, ByVal lngDeptID As Long) As Boolean
        '���浱ǰ�����ϵ����ݵ�����
        Dim blnErr As Boolean
        Dim strGen() As String
        Dim strPlan() As String
        Dim strAppend() As String
    
100     ReDim strGen(0) As String
102     ReDim strPlan(0) As String
104     ReDim strAppend(0) As String
        On Error GoTo hErr
    
106     CacheData = False

108     With Me.msfExse
110         blnErr = False
112         For intCount = 1 To .Rows - 1
114             If Trim(.TextMatrix(intCount, ExseCol.��Ŀ��)) <> "" And .RowData(intCount) <> 0 Then
116                 If Not IsNumeric(Nvl(.TextMatrix(intCount, ExseCol.��Ӧ��), "X")) Then
118                     lblMessage.Caption = lblMessage.Tag & intCount & IIf(.TextMatrix(intCount, ExseCol.��Ӧ��) = "", "�еĶ�Ӧ������Ϊ��", "�в���Ϊ��������.")
120                     blnErr = True
                    End If
                
                    '������0.000��
122                 If Int(.TextMatrix(intCount, ExseCol.��Ӧ��)) = 0 And .TextMatrix(intCount, ExseCol.�̶�) = "��" Then
124                     .TextMatrix(intCount, ExseCol.�̶�) = ""
126                     lblMessage.Caption = lblMessage.Tag & intCount & "�еĶ�Ӧ��Ϊ0,ӦΪ�ǹ̶���,���Զ�����."
                    End If
            
128                 If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
130                     lblMessage.Caption = lblMessage.Tag & intCount & "���շ���Ŀ��ǰ����շ���Ŀ���ظ���"
132                     blnErr = True
                    End If
                
134                 If Not blnErr Then
                        '�޴�����У��ż���
136                     If strGen(UBound(strGen)) <> "" Then ReDim Preserve strGen(UBound(strGen) + 1)
138                     strGen(UBound(strGen)) = .RowData(intCount)
140                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.��Ŀ��)
142                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.���)
144                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.��λ)
146                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.��ǰ��)
148                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.��Ӧ��)
150                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.�̶�)
152                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.����)
154                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.�շѷ�ʽ)
156                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & lngSource  '������Դ
158                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & lngDeptID  '���ÿ���id
                    End If
                End If
            Next
        End With
    
        '�����Ŀ ���Ӹ��ӷ���
160     If stbExse.TabVisible(1) Then
162         With Me.vfgExse
164             For intCount = .FixedRows To .Rows - 1
166                 If Val(.TextMatrix(intCount, .ColIndex("�շ�ID"))) > 0 And Val(.TextMatrix(intCount, .ColIndex("����"))) > 0 _
                      And .TextMatrix(intCount, .ColIndex("��λ")) <> "" And .TextMatrix(intCount, .ColIndex("����")) <> "" Then
168                       If strPlan(UBound(strPlan)) <> "" Then ReDim Preserve strPlan(UBound(strPlan) + 1)
170                       strPlan(UBound(strPlan)) = .TextMatrix(intCount, .ColIndex("����"))
172                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("��λ"))
174                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("����"))
176                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("��Ŀ��"))
178                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("��λ"))
180                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("�۸�"))
182                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("����"))
184                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("�̶�"))
                          strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("�շѷ�ʽ"))
186                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("����"))
188                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("�շ�ID"))
190                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & lngSource  '������Դ
192                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & lngDeptID  '���ÿ���id
                    End If
                Next
            End With
        End If
    
194     If stbExse.TabVisible(2) Then
196         With Me.vfg����
198             For intCount = .FixedRows To .Rows - 1
200                 If Val(.TextMatrix(intCount, .ColIndex("�շ�ID"))) > 0 And Val(.TextMatrix(intCount, .ColIndex("����"))) > 0 Then
202                     If strAppend(UBound(strAppend)) <> "" Then ReDim Preserve strAppend(UBound(strAppend) + 1)
                        'Append��ʽ����Ŀ��|��λ|�۸�|����|�̶�|����|�շ�ID|������Դ|���ÿ���id
204                     strAppend(UBound(strAppend)) = .TextMatrix(intCount, .ColIndex("��Ŀ��"))
206                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("��λ"))
208                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("�۸�"))
210                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("����"))
212                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("�̶�"))
214                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("����"))
216                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("�շ�ID"))
218                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & lngSource  '������Դ
220                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & lngDeptID  '���ÿ���id
                    End If
                Next
            End With
        End If
        'ɾ��ԭ�����ݣ��������ڵ����ݼ��뵽����
222     Select Case lngSource
        Case 0 '����
            '�����ֿ���
224         mGen0 = strGen
226         mPlace0 = strPlan
228         mAppend0 = strAppend
230     Case 1 '����
232         Call UpdateArray(mGen1, strGen, 10, lngDeptID)
234         Call UpdateArray(mPlace1, strPlan, 12, lngDeptID)
236         Call UpdateArray(mAppend1, strAppend, 8, lngDeptID)
238     Case 2 'סԺ
240         Call UpdateArray(mGen2, strGen, 10, lngDeptID)
242         Call UpdateArray(mPlace2, strPlan, 12, lngDeptID)
244         Call UpdateArray(mAppend2, strAppend, 8, lngDeptID)
246     Case 3 '���
248         Call UpdateArray(mGen3, strGen, 10, lngDeptID)
250         Call UpdateArray(mPlace3, strPlan, 12, lngDeptID)
252         Call UpdateArray(mAppend3, strAppend, 8, lngDeptID)
        End Select

254     CacheData = True
        Exit Function
hErr:
256     MsgBox "CacheData��" & CStr(Erl()) & "�У�" & err.Description
End Function

Private Sub UpdateArray(ByRef ArryA() As String, ByRef ArryB() As String, ByVal lngSub As Long, ByVal lngDeptKey As Long)
        '��B��������ݣ����µ�A�����С�
        'A�������� �շѶ��ջ���
        'B�������� ��ǰ�����ϵ��շѶ��ա�
        'lngSub: ����ID���ڵ��±�
        'lngDeptKey: ����ID
        On Error GoTo hErr
    
100     For intCount = LBound(ArryA) To UBound(ArryA)
102         If ArryA(intCount) <> "" Then
104             If Split(ArryA(intCount), "|")(lngSub) <> lngDeptKey Then
106                 If ArryB(UBound(ArryB)) <> "" Then ReDim Preserve ArryB(UBound(ArryB) + 1)
108                 ArryB(UBound(ArryB)) = ArryA(intCount)
                End If
            End If
        Next
110     ArryA = ArryB
        Exit Sub
hErr:
112     MsgBox "UpdateArryay��" & CStr(Erl()) & "�У�" & err.Description
End Sub

Private Function CheckArrData(ByRef ArryA() As String) As Boolean
    '��黺��������Ƿ�������
    
    CheckArrData = False
    If Val(Me.lblItem.Tag) = 0 Then lblMessage.Caption = lblMessage.Tag & "δ��ȷָ��������Ŀ��": Me.txtItem.SetFocus: Exit Function
    
    'У�����������ȫ��Ϊ����(�൱�ڲ����ײ�)����������ڴ����ֻ�����ұ�����һ������Ҹ��������Ϊ�̶���Ŀ(����ɾ��)��
    Dim bln���ڴ��� As Boolean
    Dim int������ As Integer
    Dim int���������� As Integer
    Dim intRows As Integer
    Dim rs As New ADODB.Recordset
    'Gen��ʽΪ�� id|����|���|���㵥λ|�۸�|����|�̶�|����|�շѷ�ʽ|������Դ|���ÿ���id
    For intCount = LBound(ArryA) To UBound(ArryA)
        If ArryA(intCount) <> "" Then
            If Split(ArryA(intCount), "|")(7) = "��" Then
                bln���ڴ��� = True
                Exit For
            End If
        End If
    Next
    If bln���ڴ��� Then
        For intCount = LBound(ArryA) To UBound(ArryA)
            If Split(ArryA(intCount), "|")(7) <> "��" Then
                int���������� = intCount
                int������ = int������ + 1
                If int������ > 1 Then
                    lblMessage.Caption = "��ʾ��ֻ������һ�����"
                    Exit Function
                End If
            End If
        Next
        If int������ = 1 Then
            If Split(ArryA(int����������), "|")(6) <> "��" Then
                lblMessage.Caption = "��ʾ����" & int���������� & "�����������Ϊ�̶���Ŀ��"
                Exit Function
            End If
        End If
        If int������ = 0 Then
            lblMessage.Caption = "��ʾ������Ҫ��һ�����"
            Exit Function
        End If
    End If
 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '�������ļ۸��Ƿ���ڶ��������Ŀ�����������ʾ�����ܱ���
    If bln���ڴ��� Then
        gstrSql = "Select Id From �շѼ�Ŀ Where �շ�ϸĿid=[1] And ִ������ <= SYSDATE AND (��ֹ���� > SYSDATE OR ��ֹ���� IS NULL)    And �۸�ȼ� Is Null "
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Split(ArryA(int����������), "|")(0)), gstrPriceClass)
        If rs.RecordCount > 1 Then
            lblMessage.Caption = "��ʾ������ļ۸���ڶ��������Ŀ�����ܱ��档"
            Exit Function
        End If
        rs.Close
    End If
    CheckArrData = True
        
End Function

Private Function SaveArryData(ByVal lngSource As Long, lngDeptID As Long, ArryGen() As String, ArryPlan() As String, ArryAppend() As String) As Boolean
    '���滺���е����ݡ�
    'lngSource: ������Դ
    'lngDeptID: ����ID
    'arrdept :  ���ÿ�������
    'arrGen  :  ��ͨ��������
    'arrPlan :  ��λ��������
    'arrAppen:  ���Ӷ�������
    
    Dim strItemList As String
    Dim lngCount As Long ' �ܸ���
    Dim lngLoop As Long, lngEndloop As Long
    Dim varItem As Variant, strItem As String, blnBeginTrans As Boolean, i As Integer
 
    strTemp = "": strItemList = ""
    
    For lngCount = LBound(ArryGen) To UBound(ArryGen)
        If ArryGen(lngCount) <> "" Then
            'Gen��ʽΪ�� id|����|���|���㵥λ|�۸�|����|�̶�|����|�շѷ�ʽ|������Դ|���ÿ���id
            If lngSource = Val(Split(ArryGen(lngCount), "|")(9)) And lngDeptID = Val(Split(ArryGen(lngCount), "|")(10)) Then
                If Trim(Split(ArryGen(lngCount), "|")(1)) <> "" And Split(ArryGen(lngCount), "|")(0) <> 0 Then
                    strItemList = strItemList & "|" & Split(ArryGen(lngCount), "|")(0) & "^" & _
                              Val(Split(ArryGen(lngCount), "|")(5)) & "^" & _
                              IIf(Trim(Split(ArryGen(lngCount), "|")(6)) = "", 0, 1) & "^" & _
                              IIf(Trim(Split(ArryGen(lngCount), "|")(7)) = "", 0, 1) & "^0^^ " & _
                              Val(Mid(Split(ArryGen(lngCount), "|")(8), 1, 1)) & ""
                End If
            End If
        End If
    Next
    
    '�����Ŀ ���Ӹ��ӷ���
    
        'Pla��ʽ������|��λ|����|��Ŀ��|��λ|�۸�|����|�̶�|�շѷ�ʽ|����|�շ�ID|������Դ|���ÿ���id
    For lngCount = LBound(ArryPlan) To UBound(ArryPlan)
        If ArryPlan(lngCount) <> "" Then
            If lngSource = Val(Split(ArryPlan(lngCount), "|")(11)) And lngDeptID = Val(Split(ArryPlan(lngCount), "|")(12)) Then
                If Val(Split(ArryPlan(lngCount), "|")(10)) > 0 And Val(Split(ArryPlan(lngCount), "|")(6)) > 0 _
                  And Split(ArryPlan(lngCount), "|")(1) <> "" And Split(ArryPlan(lngCount), "|")(2) <> "" Then
                  
                    strItemList = strItemList & "|" & Val(Split(ArryPlan(lngCount), "|")(10)) & "^" & _
                              Val(Split(ArryPlan(lngCount), "|")(6)) & "^" & _
                             IIf(Split(ArryPlan(lngCount), "|")(7) = "��", 1, 0) & "^0^0^" & _
                             Split(ArryPlan(lngCount), "|")(1) & "^" & _
                             Split(ArryPlan(lngCount), "|")(2) & "^" & Val(Split(ArryPlan(lngCount), "|")(8))
                End If
            End If
        End If
    Next
    
     
    
    'Append��ʽ����Ŀ��|��λ|�۸�|����|�̶�|����|�շ�ID|������Դ|���ÿ���id
    For lngCount = LBound(ArryAppend) To UBound(ArryAppend)
        If ArryAppend(lngCount) <> "" Then
            If lngSource = Val(Split(ArryAppend(lngCount), "|")(7)) And lngDeptID = Val(Split(ArryAppend(lngCount), "|")(8)) Then
                If Val(Split(ArryAppend(lngCount), "|")(6)) > 0 And Val(Split(ArryAppend(lngCount), "|")(3)) > 0 Then
                    strItemList = strItemList & "|" & Val(Split(ArryAppend(lngCount), "|")(6)) & "^" & _
                             Val(Split(ArryAppend(lngCount), "|")(3)) & "^" & _
                             IIf(Split(ArryAppend(lngCount), "|")(4) = "��", 1, 0) & "^0^1^^0"
                End If
            End If
        End If
    Next
        
    
    If strItemList <> "" Then strItemList = Mid(strItemList, 2)
    
    varItem = Split(strItemList, "|")
    lngCount = UBound(varItem)
    lngEndloop = 0
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    For lngLoop = 0 To lngCount
        
        strItem = strItem & "|" & varItem(lngLoop)
        If i = 40 Then
            strItem = Mid(strItem, 2)
            If Me.optPreproty(2).Value Then
                gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",2,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
            ElseIf Me.optPreproty(1).Value Then
                gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",1,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
            Else
                gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",0,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
            End If
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            i = 0: strItem = ""
            lngEndloop = lngEndloop + 1
        End If
        i = i + 1
    Next
    
    If Left(strItem, 1) = "|" Then
        strItem = Mid(strItem, 2)
        If Me.optPreproty(2).Value Then
            gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",2,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        ElseIf Me.optPreproty(1).Value Then
            gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",1,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        Else
            gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",0,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        End If
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If
    
    If lngLoop = 0 Then '11303 ����ȫ��ɾ�����յ��շ���Ŀ
        If Me.optPreproty(2).Value Then
            gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",2,'',1," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        ElseIf Me.optPreproty(1).Value Then
            gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",1,'',1," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        Else
            gstrSql = "zl_�����շ�_UPDATE(" & Val(Me.lblItem.Tag) & ",0,'',1," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        End If
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    Exit Function

ErrHand:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function delDept(ByVal lngSource As Long, ByVal lngDeptID As Long) As Boolean
    '�ӻ�����ɾ�����Ҽ���Ӧ���շѶ��ա�
    Dim i As Integer, curDept() As String
    ReDim curDept(0) As String
    On Error GoTo hErr
    delDept = False
    If lngSource = 1 Then
        Call DelDeptCharge(lngDeptID, mDept1, mGen1, mPlace1, mAppend1)
    ElseIf lngSource = 2 Then
        Call DelDeptCharge(lngDeptID, mDept2, mGen2, mPlace2, mAppend2)
    ElseIf lngSource = 3 Then
        Call DelDeptCharge(lngDeptID, mDept3, mGen3, mPlace3, mAppend3)
    End If
    delDept = True
    Exit Function
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DelDeptCharge(ByVal lngDeptID As Long, arrDept() As String, arrGen() As String, arrPlan() As String, arrAppend() As String)
    '�ӻ���������ɾ��ָ�����ҵĶ���
    Dim i As Integer, curDept() As String
    Dim curGen() As String, curPlan() As String, curAppend() As String
    
    ReDim curDept(0) As String
    ReDim curGen(0) As String
    ReDim curPlan(0) As String
    ReDim curAppend(0) As String
    
    For i = LBound(arrDept) To UBound(arrDept)
        If arrDept(i) <> "" Then
        If Val(Split(arrDept(i), "|")(0)) <> lngDeptID Then
            If curDept(UBound(curDept)) <> "" Then ReDim Preserve curDept(UBound(curDept) + 1)
            curDept(UBound(curDept)) = arrDept(i)
        End If
        End If
    Next
    arrDept = curDept
            
    For i = LBound(arrGen) To UBound(arrGen)
        If arrGen(i) <> "" Then
        If Val(Split(arrGen(i), "|")(10)) <> lngDeptID Then
            If curGen(UBound(curGen)) <> "" Then ReDim Preserve curGen(UBound(curGen) + 1)
            curGen(UBound(curGen)) = arrGen(i)
        End If
        End If
    Next
    arrGen = curGen
    
    For i = LBound(arrPlan) To UBound(arrPlan)
        If arrPlan(i) <> "" Then
        If Val(Split(arrPlan(i), "|")(11)) <> lngDeptID Then
            If curPlan(UBound(curPlan)) <> "" Then ReDim Preserve curPlan(UBound(curPlan) + 1)
            curPlan(UBound(curPlan)) = arrPlan(i)
        End If
        End If
    Next
    arrPlan = curPlan
    
    For i = LBound(arrAppend) To UBound(arrAppend)
        If arrAppend(i) <> "" Then
        If Val(Split(arrAppend(i), "|")(8)) <> lngDeptID Then
            If curAppend(UBound(curAppend)) <> "" Then ReDim Preserve curAppend(UBound(curAppend) + 1)
            curAppend(UBound(curAppend)) = arrAppend(i)
        End If
        End If
    Next
    arrAppend = curAppend
    
End Sub

Private Sub DeptCopy(ByVal lngSource As Long, ByVal lngOldDeptID As Long)
    '���Ƶ�ǰѡ�п��ҵ���Ŀ���յ���������
    'lngSource   :������Դ
    'lngOldDeptID :����ѡ�еĿ���ID
    
'���ܣ��๦��ѡ����
'������
'     frmParent=��ʾ�ĸ�����
'     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
'     bytStyle=ѡ�������
'       Ϊ0ʱ:�б���:ID,��
'       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
'       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
'     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
'     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
'     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
'             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
'             bytStyle=1ʱ,�����Ǳ��������
'     strNote=ѡ������˵������
'     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
'     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
'     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
'     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
'     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
'     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
'     blnMulti=�Ƿ������ѡ
'     arrInput=��Ӧ�ĸ���SQL����ֵ,��˳����,����Ϊ��ȷ����
'���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    Dim strSql As String
    Dim rsDept As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strInfo As String
    Dim strGen() As String, strPlan() As String, strAppend() As String, strDept() As String
    Dim strDeptList As String
    ReDim strDept(0) As String
    ReDim strGen(0) As String
    ReDim strPlan(0) As String
    ReDim strAppend(0) As String
    Dim varDept As Variant, strLine As String, strReturn As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    If lstDept.ListCount > 0 Then
        For i = 0 To lstDept.ListCount - 1
            strDeptList = strDeptList & "," & lstDept.ItemData(i)
        Next
    End If
    
    If lngSource = 1 Or lngSource = 3 Then  '1����� 3����죻
        strSql = "Select Distinct a.����, a.����, a.ID" & vbNewLine & _
                "From ���ű� A, ��������˵�� D" & vbNewLine & _
                "Where a.Id = d.����id And (d.������� = 1 Or d.������� = 3) and d.�������� in ('�ٴ�','���','����','����','����','����','Ӫ��') " & vbNewLine & _
                " And (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                IIf(strDeptList = "", "", "   And Instr([1], ',' || a.Id || ',') <= 0") & vbNewLine & _
                "Order By ����, ����"
    Else    '2��סԺ��
        strSql = "Select Distinct a.����, a.����, a.ID" & vbNewLine & _
                "From ���ű� A, ��������˵�� D" & vbNewLine & _
                "Where a.Id = d.����id And  (d.������� = 2 Or d.������� = 3) And d.�������� in ('����','���','����','����','����','����','Ӫ��') " & vbNewLine & _
                " And (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                IIf(strDeptList = "", "", "   And Instr([1], ',' || a.Id || ',') <= 0") & vbNewLine & _
                "Order By ����, ����"
    End If
    Set rsDept = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptList & ",")
    strReturn = frmSelCur.ShowCurrSel(Me, rsDept, "����,1200,0,2;����,1800,0,2;ID,0,1,2", "ѡ����", True, , , 5000, True)
    
    If strReturn = "" Then Exit Sub
    varDept = Split(strReturn, "|")
    strInfo = ""
    
    For i = LBound(varDept) To UBound(varDept)
        '
        strLine = varDept(i)
        If UBound(Split(strLine, ",")) = 2 Then
            '�����Ƿ������˶��գ�û�в��ܸ���
            strSql = "Select �շ���ĿID From �����շѹ�ϵ Where ������Դ=[3] And ���ÿ���ID=[1] and ������ĿID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(Split(strLine, ",")(2)), CLng(Me.lblItem.Tag), lngSource)
            If rsTmp.EOF Then
                'ȡ��ǰҳ�棬��ǰ���ҵķ��ö���
                If lngSource = 1 Then
                    Call arrCopyCharge(lngOldDeptID, 10, CLng(Split(strLine, ",")(2)), mGen1)
                    Call arrCopyCharge(lngOldDeptID, 11, CLng(Split(strLine, ",")(2)), mPlace1)
                    Call arrCopyCharge(lngOldDeptID, 8, CLng(Split(strLine, ",")(2)), mAppend1)
                    Call arrCopyCharge(lngOldDeptID, 0, CLng(Split(strLine, ",")(2)), mDept1)
                ElseIf lngSource = 2 Then
                    Call arrCopyCharge(lngOldDeptID, 10, CLng(Split(strLine, ",")(2)), mGen2)
                    Call arrCopyCharge(lngOldDeptID, 11, CLng(Split(strLine, ",")(2)), mPlace2)
                    Call arrCopyCharge(lngOldDeptID, 8, CLng(Split(strLine, ",")(2)), mAppend2)
                    Call arrCopyCharge(lngOldDeptID, 0, CLng(Split(strLine, ",")(2)), mDept2)
                ElseIf lngSource = 3 Then
                    Call arrCopyCharge(lngOldDeptID, 10, CLng(Split(strLine, ",")(2)), mGen3)
                    Call arrCopyCharge(lngOldDeptID, 11, CLng(Split(strLine, ",")(2)), mPlace3)
                    Call arrCopyCharge(lngOldDeptID, 8, CLng(Split(strLine, ",")(2)), mAppend3)
                    Call arrCopyCharge(lngOldDeptID, 0, CLng(Split(strLine, ",")(2)), mDept3)
                End If
                lstDept.AddItem "" & Split(strLine, ",")(1) & "(" & Split(strLine, ",")(0) & ")"
                lstDept.ItemData(lstDept.NewIndex) = Val(Split(strLine, ",")(2))
                '��������
                'Call SaveArryData(lngSource, CLng("" & rsDept!ID), strGen, strPlan, strAppend)
            Else
               strInfo = IIf(strInfo = "", "", vbNewLine) & "" & Split(strLine, ",")(0) & " " & Split(strLine, ",")(1) & " �ÿ����Ѿ��趨�˷��ã�"
            End If
        End If

    Next
    If strInfo <> "" Then
        MsgBox strInfo
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub arrCopyCharge(ByVal lngOldDeptID As Long, ByVal lngDeptIndex As Long, ByVal lngNewDeptID As Long, arrA() As String)
    '����arrA�е�ָ�����ҵĶ�����ϸ��arrA��
    'lngDeptID : ����ID
    'lngDeptIndex :����ID��arrA�е��±�
    'arrA:����A��������ϸ
    
    Dim arrB() As String
    Dim lngRow As Long, i As Integer, varTmp As Variant, strTmp As String
    ReDim arrB(0) As String
    For lngRow = LBound(arrA) To UBound(arrA)
        If arrA(lngRow) <> "" Then
            If Split(arrA(lngRow), "|")(lngDeptIndex) = lngOldDeptID Then
                strTmp = ""
                varTmp = Split(arrA(lngRow), "|")
                For i = LBound(varTmp) To UBound(varTmp)
                    If i = lngDeptIndex Then
                        strTmp = strTmp & "|" & lngNewDeptID
                    Else
                        strTmp = strTmp & "|" & varTmp(i)
                    End If
                Next
                If strTmp <> "" Then
                    strTmp = Mid(strTmp, 2)
                If arrB(UBound(arrB)) <> "" Then ReDim Preserve arrB(UBound(arrB) + 1)
                    arrB(UBound(arrB)) = strTmp
                End If
            End If
        End If
    Next
    
    Dim blnAdd As Boolean
    '��arrB�ӵ�ArrA��
    For lngRow = LBound(arrB) To UBound(arrB)
        strTmp = arrB(lngRow)
        blnAdd = True
        For i = LBound(arrA) To UBound(arrA)
            If strTmp = arrA(i) Then
                blnAdd = False
                Exit For
            End If
        Next
        
        If blnAdd Then
            If arrA(UBound(arrA)) <> "" Then ReDim Preserve arrA(UBound(arrA) + 1)
            arrA(UBound(arrA)) = strTmp
        End If
    Next
End Sub

Private Function GetCurrSource() As Long
    '����ȡ��ǰѡ��ҳ���ȡ������Դ
    If Me.tabDept.SelectedItem.Caption = "���п���" Then
        GetCurrSource = 0
    ElseIf Me.tabDept.SelectedItem.Caption = "�������" Then
        GetCurrSource = 1
    ElseIf Me.tabDept.SelectedItem.Caption = "סԺ����" Then
        GetCurrSource = 2
    ElseIf Me.tabDept.SelectedItem.Caption = "������" Then
        GetCurrSource = 3
    End If
End Function
