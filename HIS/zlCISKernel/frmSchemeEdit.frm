VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchemeEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "���׷���"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   Icon            =   "frmSchemeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4065
      Left            =   60
      TabIndex        =   0
      Top             =   555
      Width           =   10770
      _cx             =   18997
      _cy             =   7170
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   18
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSchemeEdit.frx":058A
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fraAdvice 
      Height          =   2070
      Left            =   45
      TabIndex        =   19
      Top             =   4680
      Width           =   10800
      Begin VB.CommandButton cmd����֤�� 
         Caption         =   "��"
         Height          =   285
         Left            =   10285
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   203
         Width           =   285
      End
      Begin VB.CheckBox chkMedicineVariety 
         Caption         =   "��Ʒ������ҽ��"
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   195
         Width           =   1575
      End
      Begin VB.TextBox txt���� 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2380
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1635
         Visible         =   0   'False
         Width           =   360
      End
      Begin MSComctlLib.Toolbar tbrFree 
         Height          =   330
         Left            =   300
         TabIndex        =   33
         Top             =   810
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "����¼��ҽ��(F3)"
               ImageIndex      =   1
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbo����ִ�� 
         Height          =   300
         Left            =   6255
         TabIndex        =   18
         Text            =   "cbo����ִ��"
         Top             =   1635
         Width           =   1725
      End
      Begin VB.CommandButton cmdƵ�� 
         Height          =   240
         Left            =   4860
         Picture         =   "frmSchemeEdit.frx":065F
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(F4)"
         Top             =   1305
         Width           =   270
      End
      Begin VB.TextBox txtƵ�� 
         Height          =   300
         Left            =   3495
         TabIndex        =   8
         Top             =   1275
         Width           =   1665
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1635
         Width           =   1380
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1635
         Width           =   1530
      End
      Begin VB.CommandButton cmd�÷� 
         Height          =   240
         Left            =   2445
         Picture         =   "frmSchemeEdit.frx":0755
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(F4)"
         Top             =   1305
         Width           =   270
      End
      Begin VB.TextBox txt�÷� 
         Height          =   300
         Left            =   930
         TabIndex        =   6
         Top             =   1275
         Width           =   1815
      End
      Begin VB.ComboBox cbo��Ч 
         Height          =   300
         ItemData        =   "frmSchemeEdit.frx":084B
         Left            =   930
         List            =   "frmSchemeEdit.frx":0855
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   195
         Width           =   2160
      End
      Begin VB.CommandButton cmdExt 
         Height          =   285
         Left            =   4890
         Picture         =   "frmSchemeEdit.frx":0869
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "�༭(F4)"
         Top             =   600
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   285
         Left            =   4890
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   900
         Width           =   285
      End
      Begin VB.ComboBox cboִ�п��� 
         Height          =   300
         Left            =   6255
         TabIndex        =   16
         Text            =   "cboִ�п���"
         Top             =   1275
         Width           =   1725
      End
      Begin VB.TextBox txtҽ������ 
         Height          =   675
         Left            =   930
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "�� ~ ���л���ݸ������"
         Top             =   555
         Width           =   3945
      End
      Begin VB.ComboBox cboִ��ʱ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6255
         TabIndex        =   15
         Top             =   915
         Width           =   4350
      End
      Begin VB.ComboBox cboִ������ 
         Height          =   300
         ItemData        =   "frmSchemeEdit.frx":095F
         Left            =   8805
         List            =   "frmSchemeEdit.frx":096C
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1275
         Width           =   1800
      End
      Begin VB.ComboBox cboҽ������ 
         Height          =   300
         Left            =   6255
         TabIndex        =   14
         Top             =   555
         Width           =   4350
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6255
         TabIndex        =   37
         Top             =   195
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox txt����֤�� 
         Height          =   300
         Left            =   6255
         MaxLength       =   100
         TabIndex        =   13
         Top             =   195
         Width           =   4335
      End
      Begin VB.Label lbl���ٵ�λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��/����"
         Height          =   180
         Left            =   7320
         TabIndex        =   38
         Top             =   255
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lbl����֤�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����֤��"
         Height          =   180
         Left            =   5490
         TabIndex        =   36
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         Height          =   180
         Left            =   2190
         TabIndex        =   34
         Top             =   1695
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl����ִ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ִ��"
         Height          =   180
         Left            =   5490
         TabIndex        =   32
         Top             =   1695
         Width           =   720
      End
      Begin VB.Label lblƵ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ƶ��"
         Height          =   180
         Left            =   3105
         TabIndex        =   26
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lbl������λ 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ"
         Height          =   180
         Left            =   4905
         TabIndex        =   22
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3105
         TabIndex        =   21
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label lbl������λ 
         BackStyle       =   0  'Transparent
         Caption         =   "��λ"
         Height          =   180
         Left            =   2490
         TabIndex        =   24
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   525
         TabIndex        =   23
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label lbl��Ч 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����Ч"
         Height          =   180
         Left            =   165
         TabIndex        =   31
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lblҽ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   5490
         TabIndex        =   30
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblִ�п��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         Height          =   180
         Left            =   5490
         TabIndex        =   28
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl�÷� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�÷�"
         Height          =   180
         Left            =   525
         TabIndex        =   25
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lblҽ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   165
         TabIndex        =   20
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblִ��ʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ��ʱ��"
         Height          =   180
         Left            =   5490
         TabIndex        =   27
         Top             =   975
         Width           =   720
      End
      Begin VB.Label lblִ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ������"
         Height          =   180
         Left            =   8055
         TabIndex        =   29
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   5640
         TabIndex        =   39
         Top             =   255
         Visible         =   0   'False
         Width           =   570
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   435
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSchemeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOK As Boolean
        
'��ڲ���
Private mint��Χ As Integer '1-����ʹ��,2-סԺʹ��,3-�����סԺ������ʹ��
Private mrsScheme As ADODB.Recordset '���"������Ŀ���"��ͬ�ṹ�Ķ�̬��¼��
Private mbln��ʾȱʡ�� As Boolean        '�ٴ�·����Ŀ�������"ѡ��ʹ��"ʱ��ʾȱʡ��
Private mbyt���� As Byte            'byt����=1-�ٴ�·����Ŀ������ã�0-���׷�������,2-�ٴ�·����Ŀ��������

Private mstr���Ʒ��� As String
Private mstr�������� As String
Private mstrִ�з��� As String

'�������
Private mobjVBA As Object
Private mobjScript As clsScript
Private mrsDefine As ADODB.Recordset
Private mlngNextID As Long
Private mblnView As Boolean
Private msng���� As Single
Private mstrʹ�ÿ��� As String

'���ز���
Private mint���� As Integer
Private mstrLike As String
Private mblnһ���� As Boolean '����ȱʡΪһ����
Private mblnNewLIS As Boolean

'�¼�״̬���Ʊ���
Private mblnNoSave As Boolean
Private mblnRowMerge As Boolean
Private mblnRunFirst As Boolean
Private mblnRowChange As Boolean

'����������
Private Const conMenu_New = 100
Private Const conMenu_Insert = 101
Private Const conMenu_Delete = 102
Private Const conMenu_Merge = 104
Private Const conMenu_Import = 105
Private Const conMenu_Save = 107
Private Const conMenu_Exit = 111
Private Const conMenu_MoveDown = 203
Private Const conMenu_MoveUp = 204

'ִ��ʱ��ʾ��
Private Const COL_����ִ�� = _
    "ÿ������ 1/8-3/8-5/8 �� 1/8:00-3/8:00-5/8:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ������һ��8:00,��������8:00,�������8:00�⼸��ʱ��ִ��"
Private Const COL_����ִ�� = _
    "ÿ������ 8-12-16 �� 8:00-12:00-16:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ��8:00,12:00,16:00�⼸��ʱ��ִ��" & vbCrLf & _
    "����һ�� 1/8 �� 1/8:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ�����еĵ�1��8:00���ʱ��ִ��"
Private Const COL_��ʱִ�� = _
    "ÿСʱ���� 1:20-1:40" & vbCrLf & _
        vbTab & "��ʾ��ÿСʱ�ڵ�20��40����������ʱ��ִ��" & vbCrLf & _
    "��Сʱһ�� 2:30 �� 1:30 �� 1:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ��Сʱ�ڵĵ�2�ĸ�Сʱ��30�������ʱ��ִ��" & vbCrLf & _
        vbTab & "������ÿ��Сʱ�ڵĵ�1�ĸ�Сʱ��30�������ʱ��ִ��" & vbCrLf & _
        vbTab & "������ÿ��Сʱ�ڵĵ�1�ĸ�Сʱ���ʱ��ִ��"

Private Enum mvCol
    '�ɼ�������
    col_��ѡ = 0
    col_ȱʡ = 1
    COL_��Ч = 2
    col_ҽ������ = 3
    COL_���� = 4
    COL_������λ = 5
    COL_���� = 6
    COL_������λ = 7
    COL_���� = 8
    COL_Ƶ�� = 9
    COL_�÷� = 10
    COL_ҽ������ = 11
    COL_ִ��ʱ�� = 12
    
    '����������
    COL_���ID = 13
    COL_��� = 14
    COL_��� = 15
    COL_������ĿID = 16
    COL_���� = 17
    COL_�걾��λ = 18
    COL_��鷽�� = 19
        COL_��ҩ��̬ = 19 '0=ɢװ��1=��ҩ��Ƭ��2=����
    COL_�շ�ϸĿID = 20
    COL_Ƶ�ʴ��� = 21
    COL_Ƶ�ʼ�� = 22
    COL_�����λ = 23
    COL_ִ�п���ID = 24
    COL_ִ������ = 25 '����ҽ����¼.ִ������=������ĿĿ¼.ִ�п���
    COL_ִ�б�� = 26
    
    COL_���㷽ʽ = 27 '������ĿĿ¼.���㷽ʽ
    COL_Ƶ������ = 28 '������ĿĿ¼.ִ��Ƶ��
    COL_�������� = 29 '������ĿĿ¼.��������
    COL_�ɷ���� = 30 '�������ڴ���Ƿ��������
        COL_�������� = 30
    COL_����ϵ�� = 31
    COL_��װ��λ = 32
    COL_��װϵ�� = 33
    COL_������� = 34
    COL_ҩƷ���� = 35
    COL_�䷽ID = 36
    COL_�ٴ��Թ�ҩ = 37
    COL_�����ĿID = 38
    COL_����֤�� = 39
    COL_�����ȼ� = 40 '����ҩ��ȼ�:0-�ǿ���ҩ,1-�����Ƽ�,2-���Ƽ�,3-����ʹ�ü�
    COL_�Ƿ�ͣ�� = 41 '=1��ʶ��ͣ�ã�=0��NULL��ʶδͣ��
    COL_ִ�з��� = 42 '0-�����������,1-��Һ��,2-ע����,3-Ƥ��,4-�ڷ�
End Enum

Public Function ShowMe(frmParent As Object, ByVal int��Χ As Long, Optional rsScheme As ADODB.Recordset, _
    Optional ByVal blnView As Boolean, Optional ByVal bln��ʾȱʡ�� As Boolean, Optional ByVal strʹ�ÿ��� As String, Optional ByVal byt���� As Byte, _
    Optional ByVal str���Ʒ��� As String, Optional ByVal str�������� As String, Optional ByVal strִ�з��� As String) As ADODB.Recordset
'���أ����"������Ŀ���"��ͬ�ṹ�Ķ�̬��¼��,���ȡ���򷵻�Nothing
'������byt����=1-�ٴ�·����Ŀ������ã�0-���׷�������,2-�ٴ�·����Ŀ��������
'   str���Ʒ���:byt����=2ʱ����
'   str��������:byt����=2ʱ����
'   strִ�з���:byt����=2ʱ����

    mint��Χ = int��Χ
    mbln��ʾȱʡ�� = bln��ʾȱʡ��
    Set mrsScheme = rsScheme
    mblnView = blnView
    mstrʹ�ÿ��� = strʹ�ÿ���
    mbyt���� = byt����
    
    mstr���Ʒ��� = str���Ʒ���
    mstr�������� = str��������
    mstrִ�з��� = strִ�з���
   
    On Error Resume Next
    Me.Show 1, frmParent
    
    If mblnOK Then
        Set ShowMe = mrsScheme
    End If
    Set mrsScheme = Nothing
    
End Function

Private Sub InitCommandBar()
'���ܣ���ʼ��������
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = frmIcons.imgMain.Icons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_New, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Insert, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Merge, "һ����ҩ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_MoveUp, "����")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_MoveDown, "����")
        Set objControl = .Add(xtpControlButton, conMenu_Import, "����")
        objControl.IconId = conMenu_Insert
        objControl.BeginGroup = True
        objControl.ToolTipText = "�Ӳ���ҽ������"
        Set objControl = .Add(xtpControlButton, conMenu_Save, "����")
        objControl.ToolTipText = "ȷ�ϱ��沢�˳�"
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "�˳�"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_New
        .Add FCONTROL, vbKeyI, conMenu_Insert
        .Add FCONTROL, vbKeyK, conMenu_Merge
        .Add FCONTROL, vbKeyT, conMenu_Import
        .Add FCONTROL, vbKeyS, conMenu_Save
        .Add FALT, vbKeyX, conMenu_Exit
    End With
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ��������ݣ����ڴ�����Ի����ûָ�֮ǰ
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
    
    strHead = _
        "��ѡ,450,4;ȱʡ,450,4;��Ч,500,4;ҽ������,3500,1;����,600,7;��λ,450,1;����,600,7;��λ,450,1;����,450,1;Ƶ��,1200,1;�÷�,1200,1;" & _
        "ҽ������,1000,1;ִ��ʱ��;���ID;���;���;������ĿID;����;�걾��λ;��鷽��;�շ�ϸĿID;Ƶ�ʴ���;Ƶ�ʼ��;�����λ;ִ�п���ID;" & _
        "ִ������;ִ�б��;���㷽ʽ;Ƶ������;��������;�ɷ����;����ϵ��;��װ��λ;��װϵ��;�������;ҩƷ����;�䷽ID;�ٴ��Թ�ҩ;�����ĿID;" & _
        "����֤��,1000,1;�����ȼ�;�Ƿ�ͣ��;ִ�з���"
        
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        If mbln��ʾȱʡ�� = False Then
            .ColHidden(col_ȱʡ) = True
            .ColHidden(col_��ѡ) = True
        Else
            .ColDataType(col_ȱʡ) = flexDTBoolean
            .ColDataType(col_��ѡ) = flexDTBoolean
            .ColHidden(col_ȱʡ) = False
            .ColHidden(col_��ѡ) = False
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub cbo����_Change()
     cbo����.Tag = "1"
End Sub

Private Sub cbo����_Click()
    cbo����.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo����_GotFocus()
    zlControl.TxtSelAll cbo����
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cbo����_Validate(False)
    ElseIf InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cbo����.Text) > 10 Then
        MsgBox "�����������ݹ��������������Ƿ���ȷ��", vbInformation, gstrSysName
        Call cbo����_GotFocus: Cancel = True: Exit Sub
    End If
    '��������
    Call AdviceChange
End Sub

Private Sub cbo����ִ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSql As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo����ִ��.ListIndex = -1 Then Exit Sub
    
    If cbo����ִ��.ItemData(cbo����ִ��.ListIndex) = -1 Then
        strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([1],3)") & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " Order by A.����"
        vRect = zlControl.GetControlRect(cbo����ִ��.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl����ִ��.Caption, False, "", "", False, False, True, vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, mint��Χ)
        If Not rsTmp Is Nothing Then
            intIdx = Cbo.FindIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then
                cbo����ִ��.ListIndex = intIdx
            Else
                cbo����ִ��.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ִ��.ListCount - 1
                cbo����ִ��.ItemData(cbo����ִ��.NewIndex) = rsTmp!ID
                cbo����ִ��.ListIndex = cbo����ִ��.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
            '�ָ������еĿ���(������Click)
            intIdx = Cbo.FindIndex(cbo����ִ��, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�п���ID)))
            Call Cbo.SetIndex(cbo����ִ��.Hwnd, intIdx)
        End If
    Else
        cbo����ִ��.Tag = "1"
        lngRow = vsAdvice.Row
        
        '���¸����˵�ִ�п���ҽ������
       Call AdviceChange
    End If
End Sub

Private Sub cbo����ִ��_GotFocus()
    Call zlControl.TxtSelAll(cbo����ִ��)
End Sub

Private Sub cbo����ִ��_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo����ִ��.ListIndex = -1 Then
            Call cbo����ִ��_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo����ִ��_Validate(False)
        End If
    End If
End Sub

Private Sub cbo����ִ��_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo����ִ��.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cbo����ִ��.Text = "" Then '������
        cbo����ִ��.Tag = "1"
        Call AdviceChange
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '�Ƿ���������ѡ�����
    blnLimit = True
    If cbo����ִ��.ListCount > 0 Then
        If cbo����ִ��.ItemData(cbo����ִ��.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    strInput = UCase(zlCommFun.GetNeedName(cbo����ִ��.Text))
    strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([3],3)") & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
        " Order by A.����"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint��Χ)
        For i = 1 To rsTmp.RecordCount
            intIdx = Cbo.FindIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then cbo����ִ��.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cbo����ִ��.ListIndex = -1 Then
            MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = zlControl.GetControlRect(cbo����ִ��.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl����ִ��.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = Cbo.FindIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then
                cbo����ִ��.ListIndex = intIdx
            Else
                cbo����ִ��.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ִ��.ListCount - 1
                cbo����ִ��.ItemData(cbo����ִ��.NewIndex) = rsTmp!ID
                cbo����ִ��.ListIndex = cbo����ִ��.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo��Ч_Click()
'���ܣ�������Ŀ��Чʱ,��յ�ǰ�е�����
    Dim lngRow As Long, i As Long
    
    With vsAdvice
        lngRow = .Row
        If .RowData(lngRow) = 0 Then Exit Sub
        
        If zlCommFun.GetNeedName(cbo��Ч.Text) = .TextMatrix(lngRow, COL_��Ч) Then Exit Sub
        
        '����¼��ҽ��ֱ�Ӹ�����Ч
        If Val(.TextMatrix(lngRow, COL_������ĿID)) = 0 Then
            .TextMatrix(lngRow, COL_��Ч) = zlCommFun.GetNeedName(cbo��Ч.Text)
            mblnNoSave = True: Exit Sub
        End If
        
        If CanAlterType(lngRow) Then
            Call AdviceAlterType(lngRow)
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, col_ҽ������)
        Else
            'һ����ҩ��ĳһ����׼��(��Ϊ���ԭ��),��ǰ�����ݲ������
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                If RowInһ����ҩ(lngRow) Then
                    MsgBox "һ����ҩ��ҩƷ�д���δ������´��ҩƷ�����ܸ���Ϊ������", vbInformation, gstrSysName
                    Call Cbo.SetIndex(cbo��Ч.Hwnd, IIF(.TextMatrix(lngRow, COL_��Ч) = "����", 0, 1))
                    Exit Sub
                End If
            End If
        
            If MsgBox("����ҽ����Ч����Ҫ��������ҽ������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Call Cbo.SetIndex(cbo��Ч.Hwnd, IIF(.TextMatrix(lngRow, COL_��Ч) = "����", 0, 1))
                Exit Sub
            End If
            
            '���ҽ��������
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '����ҩ���г�ҩ:ֻ�����ǵ�����ҩ��,ɾ����ҩ;����,�������ǰ��
                i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow, True)
            ElseIf InStr(",D,F,K,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '��������Ŀ��������Ŀ����Ѫҽ��
                'ɾ����λ�С�����������(��������,������Ŀ)����Ѫ;��
                Call Delete���������Ѫ(lngRow)
                '�����ǰ��
                Call DeleteRow(lngRow, True)
            ElseIf RowIn�䷽��(lngRow) Then
                '��ҩ�䷽��˳��(���)Ҫ������ϸ����
                'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
                lngRow = Delete��ҩ�䷽(lngRow)
                '�����ǰ��(��ҩ�÷���)
                Call DeleteRow(lngRow, True)
            Else
                '������Ŀֱ�������ǰ������
                Call DeleteRow(lngRow, True)
            End If
            
            '���½�����
            i = cbo��Ч.ListIndex '������ǰѡ�����Ч
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, col_ҽ������)
            cbo��Ч.ListIndex = i '������Ҫ�ټ��������ÿ�ʼʱ��ֵ
        End If
    End With
End Sub

Private Sub cbo��Ч_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo��Ч.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = Cbo.MatchIndex(cbo��Ч.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo��Ч.ListCount > 0 Then lngIdx = 0
        cbo��Ч.ListIndex = lngIdx
    End If
End Sub

Private Sub Set�÷�Input(rsInput As ADODB.Recordset, ByVal int���� As Integer)
'���ܣ������ҩ;������ҩ�÷������
'������rsInput=�����ѡ��ķ��ؼ�¼
'      int����=2-��ҩ;��,4-��ҩ�÷�
'˵���������ѡƵ��,����ϸ�ҩ;���������ִ��ʱ�䷽���ı仯
    Dim rsTmp As New ADODB.Recordset
    Dim blnValid As Boolean, strSql As String, i As Long
    Dim strƵ�� As String, intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim vMsg As VbMsgBoxResult, strMsg As String
    
    On Error GoTo errH
    cmd�÷�.Tag = rsInput!ID
    txt�÷�.Text = rsInput!����
    txt�÷�.Tag = "1"
    
    With vsAdvice
       
        If int���� = 2 Then
            If NVL(rsInput!ִ�з���ID, 0) <> 1 And cbo����.Text <> "" Then
                '����Һ���������
                cbo����.Text = ""
                cbo����.Tag = "1"
            End If
        End If
        '���»�ȡ���õ�ȱʡʱ�䷽��
        If cboִ��ʱ��.Enabled Then '"��ѡƵ��"��ҩƷʱ
            Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(.Row), .TextMatrix(.Row, COL_Ƶ��), rsInput!ID)
            If cboִ��ʱ��.ListCount > 0 Then
                Call Cbo.SetIndex(cboִ��ʱ��.Hwnd, 0)
                cboִ��ʱ��.Tag = "1"
            Else
                '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                If cboִ��ʱ��.Text <> "" Then
                    blnValid = ExeTimeValid(cboִ��ʱ��.Text, Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), .TextMatrix(.Row, COL_�����λ))
                    If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                        cboִ��ʱ��.Text = ""
                        cboִ��ʱ��.Tag = "1"
                    End If
                End If
            End If
        End If
        
        '���������÷�������ȱʡ����
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            If Val(.TextMatrix(.Row, COL_�շ�ϸĿID)) <> 0 Then
                strSql = "Select Ƶ��,С������,���˼���,ҽ������,�Ƴ�" & _
                    " From ҩƷ�÷����� Where  ҩƷID=[1] And �÷�ID=[2] And ����=1"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(.Row, COL_�շ�ϸĿID)), Val(rsInput!ID))
            Else
                strSql = "Select Ƶ��,С������,���˼���,ҽ������,�Ƴ� From �����÷����� Where ����>0 And ��ĿID=[1] And �÷�ID=[2]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(.Row, COL_������ĿID)), Val(rsInput!ID))
            End If
            
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!Ƶ��) And Val(.TextMatrix(.Row, COL_Ƶ������)) <> 1 Then '��Ϊһ����ʱ����
                    Call GetƵ����Ϣ_����(rsTmp!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                    txtƵ��.Text = strƵ��
                    cmdƵ��.Tag = strƵ��
                    txtƵ��.Tag = "1"
                End If
                
                '�����µ�Ƶ����������ִ��ʱ��
                If cboִ��ʱ��.Enabled Then
                    Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(.Row), strƵ��, rsInput!ID)
                    If cboִ��ʱ��.ListCount > 0 Then
                        Call Cbo.SetIndex(cboִ��ʱ��.Hwnd, 0)
                        cboִ��ʱ��.Tag = "1"
                    Else
                        '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                        If cboִ��ʱ��.Text <> "" Then
                            blnValid = ExeTimeValid(cboִ��ʱ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                            If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                                cboִ��ʱ��.Text = ""
                                cboִ��ʱ��.Tag = "1"
                            End If
                        End If
                    End If
                End If

                'ҩƷ����
                If NVL(rsTmp!���˼���, 0) <> 0 Then
                    txt����.Text = FormatEx(rsTmp!���˼���, 5)
                    txt����.Tag = "1"
                End If
                
                'ҽ������
                If Not IsNull(rsTmp!ҽ������) Then
                    cboҽ������.Text = rsTmp!ҽ������
                    cboҽ������.Tag = "1"
                End If
            End If
        End If
    End With
    
    '����ǰҽ����ҩ;��/�巨�ı仯
    Call AdviceChange
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetƵ��Input(rsInput As ADODB.Recordset, ByVal int��Χ As Integer, ByVal int��ĿƵ�� As Integer)
'���ܣ�����ִ��Ƶ�ʺ����
'������rsInput=�����ѡ��ķ��ؼ�¼
'      int��Χ=1-��ҽ;2-��ҽ;-1-һ����;-2-������
'      int��ĿƵ��=��Ŀ�����ִ��Ƶ������
'˵��������÷��������ִ��ʱ�䷽���ı仯
    Dim lng�÷�ID As Long, blnValid As Boolean
    Dim strԭִ��ʱ�� As String, i As Long
    Dim sng���� As Single
    
    strԭִ��ʱ�� = cboִ��ʱ��.Text
    With vsAdvice
        '����ҽ����ִ��Ƶ�ʺ���ִ��һ�¡�
        .TextMatrix(.Row, COL_Ƶ������) = decode(int��Χ, 1, 0, 2, 0, -1, 1, -2, 2, -3, 1, -5, 1)
        If RowIn������(.Row) Or int��Χ = -3 Or int��Χ = -5 Then   'ͬ����ֵ,��Ϊ�����Լ�����Ŀ��ִ���������ж�
            For i = .Row - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(.Row) Then
                    .TextMatrix(i, COL_Ƶ������) = .TextMatrix(.Row, COL_Ƶ������)
                Else
                    Exit For
                End If
            Next
        End If
        cmdƵ��.Tag = rsInput!����
        txtƵ��.Text = rsInput!����
        txtƵ��.Tag = "1"
        
        '����������ҩƷ�����Ŀ�����
        If cbo��Ч.ListIndex = 1 And InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            If Val(.TextMatrix(.Row, COL_Ƶ������)) = 1 Then
                If txt����.Enabled Then SetDayState -1, -1
            Else
                If Not txt����.Enabled Then SetDayState 1, 1
            End If
        End If
        
        '�����������Ŀ�����:����"�ƴ�"��ѡƵ�ʵ�����Ϊһ���Ժ���������(��ҩƷ��)
        If cbo��Ч.ListIndex = 1 And InStr(",5,6,", .TextMatrix(.Row, COL_���)) = 0 And Not RowIn�䷽��(.Row) Then
            If Val(.TextMatrix(.Row, COL_���㷽ʽ)) = 3 And int��ĿƵ�� = 0 Then
                If txt����.Enabled And Val(.TextMatrix(.Row, COL_Ƶ������)) = 1 Then
                    SetItemEditable , -1
                    txt����.Text = "1"
                ElseIf Not txt����.Enabled And Val(.TextMatrix(.Row, COL_Ƶ������)) = 0 Then
                    SetItemEditable , 1
                End If
                lbl������λ.Caption = .TextMatrix(.Row, COL_������λ)
            End If
        End If
        
        '������ִ��ʱ��Ŀ�����(������ѡƵ����Ŀ������һ����֮���л�,������Ƶ���л�)
        If int��ĿƵ�� = 0 And decode(int��Χ, 1, 0, 2, 0, -1, 1, -2, 2, -3, 1, -5, 1) <> 1 Then
            If Not cboִ��ʱ��.Enabled Then SetItemEditable , , , , 1
        Else
            If cboִ��ʱ��.Enabled Then SetItemEditable , , , , -1
        End If
        If cboִ��ʱ��.Enabled Then '"��ѡƵ��"��ҩƷʱ
            If rsInput!�����λ & "" = "����" Then
                cboִ��ʱ��.Text = ""
            Else
                '�������ִ��ʱ�䷽���ı仯
                If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                    '���Ҹ�ҩ;����Ӧ����
                    lng�÷�ID = .FindRow(CLng(.TextMatrix(.Row, COL_���ID)), .Row + 1)
                    If lng�÷�ID <> -1 Then 'δ�ҵ���ҩ;�������,Ӧ�ò�����
                        lng�÷�ID = .TextMatrix(lng�÷�ID, COL_������ĿID)
                    Else
                        lng�÷�ID = 0
                    End If
                ElseIf RowIn�䷽��(.Row) Then
                    '�õ���Ӧ����ҩ�÷�ID
                    lng�÷�ID = Val(.TextMatrix(.Row, COL_������ĿID))
                End If
                
                Call Getʱ�䷽��(cboִ��ʱ��, int��Χ, txtƵ��.Text, lng�÷�ID)
                'ȡ�µ�Ƶ�ʵ�Ĭ��ִ��ʱ��
                If cboִ��ʱ��.ListCount > 0 Then
                    Call Cbo.SetIndex(cboִ��ʱ��.Hwnd, 0)
                    cboִ��ʱ��.Tag = "1"
                Else
                    '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                    If cboִ��ʱ��.Text <> "" Then
                        blnValid = ExeTimeValid(cboִ��ʱ��.Text, Val(rsInput!Ƶ�ʴ��� & ""), Val(rsInput!Ƶ�ʼ�� & ""), rsInput!�����λ & "")
                        If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                            cboִ��ʱ��.Text = ""
                            cboִ��ʱ��.Tag = "1"
                        End If
                    End If
                End If
            End If
            
            '���¼�������
            If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 _
                And .TextMatrix(.Row, COL_��Ч) = "����" And Val(.TextMatrix(.Row, COL_Ƶ������)) <> 1 Then
                sng���� = Val(txt����.Text)
                If sng���� = 0 Then sng���� = 1
                
                If txtƵ��.Text <> "" And Val(txt����.Text) <> 0 _
                    And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                    And Val(.TextMatrix(.Row, COL_��װϵ��)) <> 0 Then
                    
                    txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                        Val(txt����.Text), sng����, rsInput!Ƶ�ʴ���, _
                        rsInput!Ƶ�ʼ��, rsInput!�����λ & "", cboִ��ʱ��.Text, _
                        Val(.TextMatrix(.Row, COL_����ϵ��)), _
                        Val(.TextMatrix(.Row, COL_��װϵ��)), _
                        Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                    txt����.Tag = "1"
                End If
            End If
        End If
        If rsInput!�����λ & "" = "����" Then
            If cboִ��ʱ��.Enabled Then SetItemEditable , , , , -1
        End If
    End With
    
    '����Ƿ�仯
    If cboִ��ʱ��.Text <> strԭִ��ʱ�� Then cboִ��ʱ��.Tag = "1"
    
    '����ǰҽ��ִ��Ƶ�ʵı仯
    Call AdviceChange
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = lngTop
    vsAdvice.Height = lngBottom - lngTop - (fraAdvice.Height - 80)
    vsAdvice.Width = lngRight - lngLeft
    
    fraAdvice.Left = lngLeft
    fraAdvice.Top = vsAdvice.Top + vsAdvice.Height - 80
    fraAdvice.Width = lngRight - lngLeft
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    If vsAdvice.Redraw = flexRDNone Then Exit Sub
    
    If mbyt���� = 2 Then
        If Control.ID = conMenu_Delete Or Control.ID = conMenu_Save Or Control.ID = conMenu_Exit Then
            Control.Visible = True
            If Control.ID = conMenu_Save Then Control.Enabled = mblnNoSave
        Else
            Control.Visible = False
        End If
        Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_New
            If mblnView Then Control.Visible = False
        Case conMenu_Insert
            If mblnView Then
                Control.Visible = False
            Else
                blnEnabled = True
                If Not fraAdvice.Enabled Then
                    If InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) > 0 _
                        And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID)) = Val(vsAdvice.TextMatrix(vsAdvice.Row - 1, COL_���ID)) Then
                        blnEnabled = False
                    End If
                End If
                Control.Enabled = blnEnabled
            End If
        Case conMenu_Delete
            If mblnView Then
                Control.Visible = False
            Else
                With vsAdvice
                    blnEnabled = True
                    If .RowData(.Row) <> 0 Then
                        If Not fraAdvice.Enabled Then blnEnabled = False
                    End If
                    Control.Enabled = blnEnabled
                End With
            End If
        Case conMenu_Merge
            If mblnView Then
                Control.Visible = False
            Else
                Control.Checked = mblnRowMerge
                blnEnabled = True
                If Not fraAdvice.Enabled Then blnEnabled = False
                Control.Enabled = blnEnabled
            End If
        Case conMenu_Import
            If mblnView Then Control.Visible = False
        Case conMenu_Save
            If mblnView Then
                Control.Visible = False
            Else
                Control.Enabled = mblnNoSave
            End If
    End Select
    
End Sub

Private Sub chkMedicineVariety_Click()
    'ȡ����Ʒ������
    If chkMedicineVariety.Tag = "" And Trim(txtҽ������.Text) <> "" Then
        If MsgBox("��ȷ��Ҫ�����ǰҽ����������������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            txtҽ������.Text = ""
            If txtҽ������.Enabled Then txtҽ������.SetFocus
        End If
    End If
End Sub

Private Sub chkMedicineVariety_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub

Private Sub cmdƵ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str��Χ As String, intƵ�� As Integer, vRect As RECT
    Dim lng������ĿID As Long, lngFind As Long
        
    With vsAdvice
        If cbo��Ч.ListIndex = 1 Then
            intƵ�� = Get��ĿƵ��(.Row)
            If Not RowIn�䷽��(.Row) And intƵ�� = 0 Then
                str��Χ = "1,-1" '��������Ϊһ����
            Else
                str��Χ = GetƵ�ʷ�Χ(.Row)
            End If
        Else
            str��Χ = GetƵ�ʷ�Χ(.Row)
            intƵ�� = decode(str��Χ, "1", 0, "2", 0, "-1", 1, "-2", 2, "-3", 1, "-5", 1)
        End If
        
        '��ѡ��Ƶ�ʵĳ���Ƶ��
        lng������ĿID = Val(.TextMatrix(.Row, COL_������ĿID))
        If RowIn������(.Row) Then
            lngFind = .FindRow(CStr(.RowData(.Row)), .FixedRows, COL_���ID)
            If lngFind <> -1 Then
                lng������ĿID = Val(.TextMatrix(lngFind, COL_������ĿID))
            End If
        End If
        strSql = ""
        If InStr("," & str��Χ & ",", ",1,") > 0 Then
            strSql = " And (Exists(Select 1 From �����÷����� Where ��ĿID=[2] And �÷�ID is NULL And Ƶ��=A.���� And A.���÷�Χ=1)" & _
                " Or (Select Count(*) From �����÷����� Where ��ĿID=[2] And �÷�ID is NULL And Ƶ�� Is Not NULL)<=1)"
        End If
        strSql = _
            " Select Rownum as ID,A.����,A.����,A.����," & _
            " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.���÷�Χ as ��ΧID" & _
            " From ����Ƶ����Ŀ A" & _
            " Where (Instr([1],','||A.���÷�Χ||',')>0  Or a.���÷�Χ=[3])" & strSql & _
            " Order by A.���÷�Χ,A.����"
        vRect = zlControl.GetControlRect(txtƵ��.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "����Ƶ��", False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, False, True, "," & str��Χ & ",", lng������ĿID, IIF(cbo��Ч.ListIndex = 1, -5, -3))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�п��õ�����Ƶ����Ŀ�����ȵ�ҽ��Ƶ�ʹ��������á�", vbInformation, gstrSysName
            End If
            txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
            Call zlControl.TxtSelAll(txtƵ��)
            txtƵ��.SetFocus: Exit Sub
        End If
        Call SetƵ��Input(rsTmp, rsTmp!��ΧID, intƵ��)
        txtƵ��.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub cmd����֤��_Click()
    Dim strSql As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vPoint As PointAPI
    
    strSql = _
            " Select ID,ID as ��ĿID,����,����,����," & IIF(mint���� = 0, "����", "����� as ����") & ",˵��" & _
            " From ��������Ŀ¼" & _
            " Where ���='Z' " & _
            " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by ����"
                
    vPoint = zlControl.GetCoordPos(txt����֤��.Hwnd, 0, 0)
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "��ҽ֤��", False, "", "", False, False, True, _
        vPoint.X, vPoint.Y, txt����֤��.Height, blnCancel, False, True)
    If Not blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
        '���������뷽ʽ
        If rsTmp Is Nothing Then
            MsgBox "û���ҵ���ҽ֤�򼲲���", vbInformation, gstrSysName
        Else
            txt����֤��.Text = rsTmp!���� & ""
            txt����֤��.Tag = rsTmp!ID & ""
        End If
    End If
    '��������
    Call AdviceChange
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    
    If mblnNoSave Then
        strMsg = "��ǰ���׷������ݱ༭����δ���棬ȷʵҪ�˳���"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
End Sub

Private Sub lbl����_Click()
    Call Load��Һ����(cbo����, lbl���ٵ�λ, True)
    cbo����.Tag = "1"
    Call AdviceChange
End Sub

Private Sub tbrFree_ButtonClick(ByVal Button As MSComctlLib.Button)
    'ǿ��ʱ�����������
    If Button.value = 0 Then
        If vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If MsgBox("ȡ������¼��״̬�������¼���ҽ�����ݣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Button.value = 1
                Call zlControl.TxtSelAll(txtҽ������)
                txtҽ������.SetFocus: Exit Sub
            End If
            Call DeleteRow(vsAdvice.Row, True)
            mblnNoSave = True
            Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        End If
    End If
    
    txtҽ������.Text = ""
    txtҽ������.SetFocus
End Sub

Private Sub txtƵ��_GotFocus()
    Call zlControl.TxtSelAll(txtƵ��)
End Sub

Private Sub txtƵ��_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str��Χ As String, intƵ�� As Integer, vRect As RECT
    Dim lng������ĿID As Long, lngFind As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If cmdƵ��.Tag <> "" And txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��) And txtƵ��.Text <> "" Then
                Call SeekNextControl
            ElseIf txtƵ��.Text = "" Then
                If cmdƵ��.Enabled And cmdƵ��.Visible Then cmdƵ��_Click
            Else
                If cbo��Ч.ListIndex = 1 Then
                    intƵ�� = Get��ĿƵ��(.Row)
                    If Not RowIn�䷽��(.Row) And intƵ�� = 0 Then
                        str��Χ = "1,-1" '��������Ϊһ����
                    Else
                        str��Χ = GetƵ�ʷ�Χ(.Row)
                    End If
                Else
                    str��Χ = GetƵ�ʷ�Χ(.Row)
                    intƵ�� = intƵ�� = decode(str��Χ, "1", 0, "2", 0, "-1", 1, "-2", 2, "-3", 1, "-5", 1)
                End If
                
                '��ѡ��Ƶ�ʵĳ���Ƶ��
                lng������ĿID = Val(.TextMatrix(.Row, COL_������ĿID))
                If RowIn������(.Row) Then
                    lngFind = .FindRow(CStr(.RowData(.Row)), .FixedRows, COL_���ID)
                    If lngFind <> -1 Then
                        lng������ĿID = Val(.TextMatrix(lngFind, COL_������ĿID))
                    End If
                End If
                strSql = ""
                If InStr("," & str��Χ & ",", ",1,") > 0 Then
                    strSql = " And (Exists(Select 1 From �����÷����� Where ��ĿID=[4] And �÷�ID is NULL And Ƶ��=A.���� And A.���÷�Χ=1)" & _
                        " Or (Select Count(*) From �����÷����� Where ��ĿID=[4] And �÷�ID is NULL And Ƶ�� Is Not NULL)<=1)"
                End If
                strSql = _
                    " Select Rownum as ID,A.����,A.����,A.����," & _
                    " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.���÷�Χ as ��ΧID" & _
                    " From ����Ƶ����Ŀ A" & _
                    " Where (Instr([3],','||A.���÷�Χ||',')>0   Or a.���÷�Χ=[5])" & strSql & _
                    " And (A.���� Like [1] Or Upper(A.����) Like [2]" & _
                    " Or Upper(A.����) Like [2] Or Upper(A.Ӣ������) Like [2])" & _
                    " Order by A.���÷�Χ,A.����"
                vRect = zlControl.GetControlRect(txtƵ��.Hwnd)
                Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "����Ƶ��", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, False, True, UCase(txtƵ��.Text) & "%", _
                    mstrLike & UCase(txtƵ��.Text) & "%", "," & str��Χ & ",", lng������ĿID, IIF(cbo��Ч.ListIndex = 1, -5, -3))
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ�������Ƶ����Ŀ��", vbInformation, gstrSysName
                    End If
                    txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
                    Call zlControl.TxtSelAll(txtƵ��)
                    txtƵ��.SetFocus: Exit Sub
                End If
                Call SetƵ��Input(rsTmp, rsTmp!��ΧID, intƵ��)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Function GetBaseRow(ByVal lngRow As Long) As Long
'���ܣ��ɵ�ǰ�ɼ��л�ȡ����Ŀ����
    If RowIn�䷽��(lngRow) Then
        '��ȡ��ҩ�䷽��һζ��ҩ��
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
    ElseIf RowIn������(lngRow) Then
        '��ȡһ�������ĵ�һ����Ŀ��
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
    Else
        GetBaseRow = lngRow
    End If
End Function

Private Function Get��ĿƵ��(ByVal lngRow As Long) As Integer
'���ܣ���ȡָ����Ŀ��ԭʼִ��Ƶ������
'������lngRow=��ǰ�ɼ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    lngRow = GetBaseRow(lngRow)
    strSql = "Select ִ��Ƶ�� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID)))
    If Not rsTmp.EOF Then Get��ĿƵ�� = NVL(rsTmp!ִ��Ƶ��, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmd�÷�_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int���� As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim lng��Ŀid As Long
    Dim strWhere As String

    With vsAdvice
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            int���� = 2 '��ҩ;��
        ElseIf RowIn������(vsAdvice.Row) Then
            int���� = 6 '�ɼ�����
        ElseIf .TextMatrix(.Row, COL_���) = "K" Then
            If gblnѪ��ϵͳ = True Then
                If Val(.TextMatrix(.Row, COL_��鷽��)) = 0 Then
                    int���� = 9 '�ɼ���Ѫ;��
                Else
                    int���� = 8 '��Ѫ;��
                    strWhere = " And nvl(A.ִ�з���,0)=1 "
                End If
            Else
                int���� = 8 '��Ѫ;��
            End If
        Else
            int���� = 4 '��ҩ�÷�
        End If
        lng��Ŀid = Val(.TextMatrix(.Row, COL_������ĿID))
        If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
            If Val(.TextMatrix(.Row, COL_�շ�ϸĿID)) = 0 Then
                strSql = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[2] And ����>0)" & _
                    " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                        " Where A.�÷�ID=B.ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([3],3)") & " And A.��ĿID=[2] And A.����>0)<=1)"
            Else
                lng��Ŀid = Val(.TextMatrix(.Row, COL_�շ�ϸĿID))
                strSql = " And (A.ID IN (Select �÷�ID From ҩƷ�÷����� Where ҩƷID=[2] And ����=1)" & _
                    " Or (Select Count(A.�÷�ID) From ҩƷ�÷����� A,������ĿĿ¼ B" & _
                        " Where A.�÷�ID=B.ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([3],3)") & " And A.ҩƷID=[2] And A.����=1)<=1)"
            End If
        End If
        strSql = "Select Distinct A.ID,A.����,A.����,C.���� as ����,A.ִ�з��� as ִ�з���ID " & _
            " From ������Ŀ���� B,������ĿĿ¼ A,���Ʒ���Ŀ¼ C" & _
            " Where A.ID=B.������ĿID And A.����ID=C.ID(+)" & _
            " And A.���='E' And A.��������=[1] And " & IIF(mint��Χ = 3, "Nvl(A.�������,0)<>0", "A.������� IN([3],3)") & strWhere & strSQL & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " Order by A.����"
        vRect = zlControl.GetControlRect(txt�÷�.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl�÷�.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, CStr(int����), lng��Ŀid, mint��Χ)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�п��õ�" & lbl�÷�.Caption & "�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
            Call zlControl.TxtSelAll(txt�÷�)
            txt�÷�.SetFocus: Exit Sub
        End If
        
        '��һ����ҩ������ҩƷ�Ŀ��ø�ҩ;�����м��
        If int���� = 2 Then
            Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                If i <> .Row And .RowData(i) <> 0 Then
                    If Not Check�����÷�(rsTmp!ID, Val(.TextMatrix(i, COL_������ĿID)), mint��Χ) Then
                        .Refresh
                        MsgBox """" & rsTmp!���� & """���������뵱ǰҩƷһ����ҩ��""" & .TextMatrix(i, col_ҽ������) & """��", vbInformation, gstrSysName
                        .Refresh
                        txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
                        Call zlControl.TxtSelAll(txt�÷�)
                        txt�÷�.SetFocus: Exit Sub
                    End If
                End If
            Next
        End If
        
        Call Set�÷�Input(rsTmp, int����)
        txt�÷�.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub txt����֤��_GotFocus()
    zlControl.TxtSelAll txt����֤��
End Sub

Private Sub txt����֤��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub



Private Sub txt����֤��_Validate(Cancel As Boolean)
    Dim strSql As String, rsTmp As Recordset
    Dim strInput As String, blnCancel As Boolean, vPoint As PointAPI
    
    strInput = UCase(txt����֤��.Text)
    If strInput = "" Then Exit Sub
    If zlCommFun.IsCharChinese(strInput) Then
        strSql = "���� Like [2]" '���뺺��ʱֻƥ������
    Else
        strSql = "���� Like [1] Or ���� Like [2] Or " & IIF(mint���� = 0, "����", "�����") & " Like [2]"
    End If
    strSql = _
            " Select ID,ID as ��ĿID,����,����,����," & IIF(mint���� = 0, "����", "����� as ����") & ",˵��" & _
            " From ��������Ŀ¼" & _
            " Where ���='Z' And (" & strSql & ")" & _
            " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by ����"
                
    vPoint = zlControl.GetCoordPos(txt����֤��.Hwnd, 0, 0)
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, "��ҽ֤��", False, "", "", False, False, True, _
        vPoint.X, vPoint.Y, txt����֤��.Height, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
        Cancel = True
    Else
        '���������뷽ʽ
        If rsTmp Is Nothing Then
            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
            Cancel = True
        Else
            txt����֤��.Text = rsTmp!���� & ""
            txt����֤��.Tag = rsTmp!ID & ""
        End If
    End If
    '��������
    Call AdviceChange
End Sub

Private Sub txt����_Change()
    txt����.Tag = "1"
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        'Ϊ����
        If (IsNumeric(txt����.Text) Or txt����.Text = "") _
            And (IsNumeric(txt����.Text) Or txt����.Text = "") Then
            If SeekNextControl Then Call txt����_Validate(False)
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim sng���� As Single, i As Long
    Dim strSame As String, strMsg As String
        
    If txt����.Text <> "" Then
        With vsAdvice
            If Val(txt����.Text) = 0 Then
                txt����.Text = 1: txt����.Tag = "1"
            End If
            
            '����������Ҫһ��Ƶ��ͬ�ڵ�����
            If Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) <> 0 Then
                If .TextMatrix(.Row, COL_�����λ) = "��" Then
                    sng���� = 7
                ElseIf .TextMatrix(.Row, COL_�����λ) = "��" Then
                    sng���� = Val(.TextMatrix(.Row, COL_Ƶ�ʼ��))
                ElseIf .TextMatrix(.Row, COL_�����λ) = "Сʱ" Then
                    sng���� = Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) \ 24
                ElseIf .TextMatrix(.Row, COL_�����λ) = "����" Then
                    sng���� = Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) \ (24 * 60)
                End If
                If Val(txt����.Text) < sng���� Then
                    If MsgBox("��""" & .TextMatrix(.Row, COL_Ƶ��) & """ִ��ʱ��������Ҫ " & sng���� & " �����ҩ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt����_GotFocus: Exit Sub
                    End If
                End If
            End If
    
            '���¼�������
            If .TextMatrix(.Row, COL_Ƶ��) <> "" _
                And Val(.TextMatrix(.Row, COL_����)) <> 0 _
                And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                And Val(.TextMatrix(.Row, COL_��װϵ��)) <> 0 Then
                
                txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                    Val(.TextMatrix(.Row, COL_����)), Val(txt����.Text), _
                    Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), _
                    .TextMatrix(.Row, COL_�����λ), .TextMatrix(.Row, COL_ִ��ʱ��), _
                    Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_��װϵ��)), _
                    Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                txt����.Tag = "1"
            End If
        End With
        
        'ÿ��������������Ϊ�´ε�ȱʡ
        If txt����.Tag = "1" Then
            msng���� = Val(txt����.Text)
        End If
    End If
    
    Call AdviceChange
End Sub

Private Sub txt�÷�_GotFocus()
    Call zlControl.TxtSelAll(txt�÷�)
End Sub

Private Sub txt�÷�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int���� As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long
    Dim strLike As String, i As Long
    Dim lng��Ŀid As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�)) And txt�÷�.Text <> "" Then
                Call SeekNextControl
            ElseIf txt�÷�.Text = "" Then
                If cmd�÷�.Enabled And cmd�÷�.Visible Then cmd�÷�_Click
            Else
                If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                    int���� = 2 '��ҩ;��
                ElseIf RowIn������(vsAdvice.Row) Then
                    int���� = 6 '�ɼ�����
                ElseIf .TextMatrix(.Row, COL_���) = "K" Then
                    int���� = 8 '��Ѫ;��
                Else
                    int���� = 4 '��ҩ�÷�
                End If
                lng��Ŀid = Val(.TextMatrix(.Row, COL_������ĿID))
                If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
                    If Val(.TextMatrix(.Row, COL_�շ�ϸĿID)) = 0 Then
                        strSql = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[4] And ����>0)" & _
                            " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                                " Where A.�÷�ID=B.ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([6],3)") & " And A.��ĿID=[4] And A.����>0)<=1)"
                    Else
                        lng��Ŀid = Val(.TextMatrix(.Row, COL_�շ�ϸĿID))
                        strSql = " And (A.ID IN(Select �÷�ID From ҩƷ�÷����� Where ҩƷID=[4] And ����=1)" & _
                            " Or (Select Count(A.�÷�ID) From ҩƷ�÷����� A,������ĿĿ¼ B" & _
                                " Where A.�÷�ID=B.ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([6],3)") & " And A.ҩƷID=[4] And A.����=1)<=1)"
                    End If
                End If
                
                '�Ż�
                strLike = mstrLike
                If Len(txt�÷�.Text) < 2 Then strLike = ""
                
                strSql = "Select Distinct A.ID,A.����,A.����,A.ִ�з��� as ִ�з���ID " & _
                    " From ������ĿĿ¼ A,������Ŀ���� B" & _
                    " Where A.ID=B.������ĿID" & _
                    " And A.���='E' And A.��������=[3] And " & IIF(mint��Χ = 3, "Nvl(A.�������,0)<>0", "A.������� IN([6],3)") & strSql & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2])" & _
                    decode(mint����, 0, " And B.���� IN([5],3)", 1, " And B.���� IN([5],3)", "") & _
                    " Order by A.����"
                vRect = zlControl.GetControlRect(txt�÷�.Hwnd)
                Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lbl�÷�.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, UCase(txt�÷�.Text) & "%", _
                    strLike & UCase(txt�÷�.Text) & "%", CStr(int����), lng��Ŀid, mint���� + 1, mint��Χ)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ���" & lbl�÷�.Caption & "��", vbInformation, gstrSysName
                    End If
                    txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
                    Call zlControl.TxtSelAll(txt�÷�)
                    txt�÷�.SetFocus: Exit Sub
                End If
                
                '��һ����ҩ������ҩƷ�Ŀ��ø�ҩ;�����м��
                If int���� = 2 Then
                    Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
                    For i = lngBegin To lngEnd
                        If i <> .Row And .RowData(i) <> 0 Then
                            If Not Check�����÷�(rsTmp!ID, Val(.TextMatrix(i, COL_������ĿID)), mint��Χ) Then
                                .Refresh
                                MsgBox """" & rsTmp!���� & """���������뵱ǰҩƷһ����ҩ��""" & .TextMatrix(i, col_ҽ������) & """��", vbInformation, gstrSysName
                                .Refresh
                                txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
                                Call zlControl.TxtSelAll(txt�÷�)
                                txt�÷�.SetFocus: Exit Sub
                            End If
                        End If
                    Next
                End If
                
                Call Set�÷�Input(rsTmp, int����)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub txt�÷�_Validate(Cancel As Boolean)
    With vsAdvice
        '�ָ���Ϊ�����
        If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Text <> IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�)) Then
            txt�÷�.Text = IIF(cbo����.Text <> "", Replace(.TextMatrix(.Row, COL_�÷�), cbo����.Text & lbl���ٵ�λ.Caption, ""), .TextMatrix(.Row, COL_�÷�))
        End If
    End With
End Sub

Private Sub txtƵ��_Validate(Cancel As Boolean)
    With vsAdvice
        '�ָ���Ϊ�����
        If cmdƵ��.Tag <> "" And txtƵ��.Text <> .TextMatrix(.Row, COL_Ƶ��) Then
            txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
        End If
    End With
End Sub

Private Sub cboִ�п���_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSql As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
    Dim lng�������� As Long, strҩ��IDs As String
    Dim lngBegin As Long, lngEnd As Long
    
    If cboִ�п���.ListIndex = -1 Then Exit Sub
    
    If cboִ�п���.ItemData(cboִ�п���.ListIndex) = -1 Then
        strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([1],3)") & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " Order by A.����"
        vRect = zlControl.GetControlRect(cboִ�п���.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lblִ�п���.Caption, False, "", "", False, False, True, vRect.Left, vRect.Top, cboִ�п���.Height, blnCancel, False, True, mint��Χ)
        If Not rsTmp Is Nothing Then
            intIdx = Cbo.FindIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then
                cboִ�п���.ListIndex = intIdx
            Else
                cboִ�п���.AddItem rsTmp!���� & "-" & rsTmp!����, cboִ�п���.ListCount - 1
                cboִ�п���.ItemData(cboִ�п���.NewIndex) = rsTmp!ID
                cboִ�п���.ListIndex = cboִ�п���.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
            '�ָ������еĿ���(������Click)
            intIdx = Cbo.FindIndex(cboִ�п���, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�п���ID)))
            Call Cbo.SetIndex(cboִ�п���.Hwnd, intIdx)
        End If
    Else
        lngRow = vsAdvice.Row
        
        '���һ����ҩ����������
        With vsAdvice
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 And RowInһ����ҩ(lngRow) Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow, COL_���ID)), lngBegin, lngEnd)
                
                '��ǰ������ͨҩ���������������ĸ�Ϊ��������
                If sys.DeptHaveProperty(cboִ�п���.ItemData(cboִ�п���.ListIndex), "��������") Then
                    lng�������� = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                End If
                '��ǰ�����������Ļ��Ϊ��ͨҩ��
                If lng�������� = 0 Then
                    For i = lngBegin To lngEnd
                        If i <> lngRow And Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                            '�Ա�ҩ������
                            If Not (Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(i, COL_ִ������)) = 5) Then
                                If sys.DeptHaveProperty(Val(.TextMatrix(i, COL_ִ�п���ID)), "��������") Then
                                    lng�������� = Val(.TextMatrix(i, COL_ִ�п���ID)): Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                '�������������ҩƷ��ִ�п�����ͬ�����洢�趨
                If lng�������� <> 0 Then
                    For i = lngBegin To lngEnd
                        If i <> lngRow And Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                            '�Ա�ҩ������
                            If Not (Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(i, COL_ִ������)) = 5) Then
                                strҩ��IDs = Get����ҩ��IDs(.TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), Val(.TextMatrix(i, COL_�շ�ϸĿID)), 0, mint��Χ)
                                If InStr("," & strҩ��IDs & ",", "," & cboִ�п���.ItemData(cboִ�п���.ListIndex) & ",") = 0 Then
                                    MsgBox "һ����ҩ��ҩƷ�У�""" & .TextMatrix(i, col_ҽ������) & """��""" & zlCommFun.GetNeedName(cboִ�п���.Text) & """��û�д洢��", vbInformation, gstrSysName
                                    '�ָ������еĿ���(������Click)
                                    intIdx = Cbo.FindIndex(cboִ�п���, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�п���ID)))
                                    Call Cbo.SetIndex(cboִ�п���.Hwnd, intIdx)
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End With
        
        cboִ�п���.Tag = "1"
        
        '���¸����˵�ִ�п���ҽ������
        Call AdviceChange
    End If
End Sub

Private Sub cboִ�п���_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboִ�п���.ListIndex = -1 Then
            Call cboִ�п���_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cboִ�п���_Validate(False)
        End If
    End If
End Sub

Private Sub cboִ�п���_GotFocus()
    Call zlControl.TxtSelAll(cboִ�п���)
End Sub

Private Sub cboִ�п���_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If cboִ�п���.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cboִ�п���.Text = "" Then '������
        cboִ�п���.Tag = "1"
        Call AdviceChange
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '�Ƿ���������ѡ�����
    blnLimit = True
    If cboִ�п���.ListCount > 0 Then
        If cboִ�п���.ItemData(cboִ�п���.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    strInput = UCase(zlCommFun.GetNeedName(cboִ�п���.Text))
    strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([3],3)") & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
        " Order by A.����"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint��Χ)
        For i = 1 To rsTmp.RecordCount
            intIdx = Cbo.FindIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then cboִ�п���.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cboִ�п���.ListIndex = -1 Then
            MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = zlControl.GetControlRect(cboִ�п���.Hwnd)
        Set rsTmp = zldatabase.ShowSQLSelect(Me, strSql, 0, lblִ�п���.Caption, False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", mint��Χ)
        If Not rsTmp Is Nothing Then
            intIdx = Cbo.FindIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then
                cboִ�п���.ListIndex = intIdx
            Else
                cboִ�п���.AddItem rsTmp!���� & "-" & rsTmp!����, cboִ�п���.ListCount - 1
                cboִ�п���.ItemData(cboִ�п���.NewIndex) = rsTmp!ID
                cboִ�п���.ListIndex = cboִ�п���.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboִ��ʱ��_Change()
    cboִ��ʱ��.Tag = "1"
End Sub

Private Sub cboִ��ʱ��_Click()
    'cboִ��ʱ��_Change
    '��������
    cboִ��ʱ��.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cboִ��ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cboִ��ʱ��_Validate(False)
    Else
        If InStr("0123456789:-/" & Chr(8) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cboִ��ʱ��_Validate(Cancel As Boolean)
    Dim blnValid As Boolean, lngRow As Long, strTmp As String
    
    lngRow = vsAdvice.Row
        
    With vsAdvice
        If cboִ��ʱ��.Text <> "" Then
            '��鳤��
            If Len(cboִ��ʱ��.Text) > 50 Then
                MsgBox "�������ݲ��ܳ��� 50 ���ַ���", vbInformation, gstrSysName
                Call cboִ��ʱ��_GotFocus
                Cancel = True: Exit Sub
            End If
            '���Ϸ���
            If .RowData(lngRow) <> 0 Then
                blnValid = ExeTimeValid(cboִ��ʱ��.Text, Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)), Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)), .TextMatrix(lngRow, COL_�����λ))
                If Not blnValid Then
                    If .TextMatrix(lngRow, COL_�����λ) = "��" Then
                        strTmp = COL_����ִ��
                    ElseIf .TextMatrix(lngRow, COL_�����λ) = "��" Then
                        strTmp = COL_����ִ��
                    ElseIf .TextMatrix(lngRow, COL_�����λ) = "Сʱ" Then
                        strTmp = COL_��ʱִ��
                    End If
                    MsgBox "�����ִ��ʱ�䷽����ʽ����ȷ�����顣" & vbCrLf & vbCrLf & "����" & vbCrLf & strTmp, vbInformation, gstrSysName
                    Call cboִ��ʱ��_GotFocus
                    Cancel = True: Exit Sub
                End If
            End If
        End If
    End With
    
    '��������
    Call AdviceChange
End Sub

Private Sub cboִ������_Click()
    cboִ������.Tag = "1"
    '��������
    Call AdviceChange
End Sub

Private Sub cboִ������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboִ������.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = Cbo.MatchIndex(cboִ������.Hwnd, KeyAscii)
        If lngIdx = -1 And cboִ������.ListCount > 0 Then lngIdx = 0
        cboִ������.ListIndex = lngIdx
    End If
End Sub

Private Sub cmdExt_Click()
'���ܣ��޸�����ҽ������������
    Dim rsCurr As New ADODB.Recordset
    Dim strExtData As String, strTmp As String
    Dim lngRow As Long, lngFirstRow As Long
    Dim lng������ĿID As Long, lng�÷�ID As Long, strȱʡ As String
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim lngBegin As Long, lngEnd As Long, i As Long, blnRefresh As Boolean
    Dim lng�䷽ID As Long
    Dim intType As Integer, lng��Ŀid As Long, blnOK As Boolean
    Dim t_Pati As TYPE_PatiInfoEx
    
    lngRow = vsAdvice.Row
        
    If vsAdvice.TextMatrix(lngRow, COL_���) = "D" Then
        strExtData = Get��鲿λ����(lngRow)
        If strExtData = "" Then
            MsgBox "�ü��ҽ����ϵͳ������ǰ�´�ģ������з�ʽ�����ݡ��������´�ü��ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        intType = 0
    ElseIf vsAdvice.TextMatrix(lngRow, COL_���) = "F" Then
        strExtData = Get��������IDs(lngRow)
        intType = 1
    ElseIf RowIn�䷽��(lngRow) Then
        strExtData = Get��ҩ�䷽IDs(lngRow)
        intType = 2
    ElseIf RowIn������(lngRow) Then
        strExtData = Get�������IDs(lngRow)
        intType = 4
    Else
        Exit Sub '������ǰ�ļ�����Ŀ
    End If
    
    If intType = 4 Then
        lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
        lng��Ŀid = Val(vsAdvice.TextMatrix(lngFirstRow, COL_������ĿID))
    Else
        lng��Ŀid = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
    End If

    On Error Resume Next
    If intType = 2 Then
        blnOK = frmAdviceFormula.ShowMe(Me, Nothing, txtҽ������.Hwnd, t_Pati, 3, IIF(mbyt���� <> 2, 0, 3), cbo��Ч.ListIndex, mint��Χ, , lng��Ŀid, strExtData)
    Else
        blnOK = frmSchemeEditEx.ShowMe(Me, txtҽ������.Hwnd, intType, cbo��Ч.ListIndex, mint��Χ, mblnNewLIS, False, lng��Ŀid, strExtData)
    End If
    On Error GoTo 0
    
    '���������������
    If blnOK Then
        strȱʡ = vsAdvice.TextMatrix(lngRow, col_ȱʡ)
        
        If vsAdvice.TextMatrix(lngRow, COL_���) = "D" Then
            '������
            Call AdviceSet������(lngRow, strExtData)
            vsAdvice.TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
            txtҽ������.Text = vsAdvice.TextMatrix(lngRow, col_ҽ������)
        ElseIf vsAdvice.TextMatrix(lngRow, COL_���) = "F" Then
            'һ������
            Call AdviceSet�������(lngRow, strExtData)
            vsAdvice.TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
            txtҽ������.Text = vsAdvice.TextMatrix(lngRow, col_ҽ������)
            
            'ˢ�´������������ִ�п���
            blnRefresh = True
        ElseIf RowIn������(lngRow) Then
            '�������
            lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
            lng�÷�ID = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
            
            '�Ȼ�ȡ��ǰ�Ѿ����ú�ֵ
            rsCurr.Fields.Append "ҽ��ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʴ���", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʼ��", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "�����λ", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "����", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "ִ��ʱ��", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "ҽ������", adVarChar, 100, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
                        
            '�ɼ�������ִ�п��ҿ����������Ŀ��ͬ
            If Val(vsAdvice.TextMatrix(lngFirstRow, COL_ִ�п���ID)) <> 0 Then
                rsCurr!ִ�п���ID = Val(vsAdvice.TextMatrix(lngFirstRow, COL_ִ�п���ID))
            End If
            If Val(vsAdvice.TextMatrix(lngRow, COL_����)) <> 0 Then
                rsCurr!���� = Val(vsAdvice.TextMatrix(lngRow, COL_����))
            End If
            rsCurr!ִ��ʱ�� = vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��)
            rsCurr!Ƶ�� = vsAdvice.TextMatrix(lngRow, COL_Ƶ��)
            rsCurr!Ƶ�ʴ��� = Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���))
            rsCurr!Ƶ�ʼ�� = Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��))
            rsCurr!�����λ = vsAdvice.TextMatrix(lngRow, COL_�����λ)
            rsCurr!ҽ������ = vsAdvice.TextMatrix(lngRow, COL_ҽ������)
            rsCurr!ҽ��ID = vsAdvice.RowData(lngRow)
            rsCurr.Update
            
            '��ȫ�������øü������
            '------------------------
            'ɾ��������Ŀ��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete�������(lngRow)
            '�����ǰ��(�ɼ�������)
            Call DeleteRow(lngRow, True, False)
            '���²���:����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet�������(lngRow, lng�÷�ID, strExtData, rsCurr)
            
            'ǿ����ʾ��ǰҽ����Ƭ
            blnRefresh = True
        ElseIf RowIn�䷽��(lngRow) Then
            '��ҩ�䷽
            lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
            lng������ĿID = Val(vsAdvice.TextMatrix(lngFirstRow, COL_������ĿID))
            lng�÷�ID = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
            
            '�Ȼ�ȡ��ǰ�Ѿ����ú�ֵ
            rsCurr.Fields.Append "ҽ��ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "ִ������", adVarChar, 10, adFldIsNullable
            rsCurr.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʴ���", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʼ��", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "�����λ", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "����", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "ִ��ʱ��", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "ҽ������", adVarChar, 100, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
            
            rsCurr!ִ������ = zlCommFun.GetNeedName(cboִ������.Text) '����,�Ա�ҩ,��Ժ��ҩ
             'ȡ�䷽����ѡ���ҩ��
            rsCurr!ִ�п���ID = Val(Split(strExtData, "|")(4))
            rsCurr!Ƶ�� = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
            rsCurr!Ƶ�ʴ��� = Val(vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���))
            rsCurr!Ƶ�ʼ�� = Val(vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��))
            rsCurr!�����λ = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
            If Val(vsAdvice.TextMatrix(lngFirstRow, COL_����)) <> 0 Then
                rsCurr!���� = Val(vsAdvice.TextMatrix(lngFirstRow, COL_����))
            End If
            rsCurr!ִ��ʱ�� = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
            rsCurr!ҽ������ = vsAdvice.TextMatrix(lngRow, COL_ҽ������)
            rsCurr!ҽ��ID = vsAdvice.RowData(lngRow)
            
            rsCurr.Update
            
            '��ȫ�������ø���ҩ�䷽��
            '------------------------
            'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete��ҩ�䷽(lngRow)
            '�����ǰ�÷����䷽ID��Ϊ�գ������䷽ID
            lng�䷽ID = Val(vsAdvice.TextMatrix(lngRow, COL_�䷽ID))
            '�����ǰ��(��ҩ�÷���)
            Call DeleteRow(lngRow, True, False)
            '�����䷽:����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet��ҩ�䷽(lng������ĿID, lngRow, lng�÷�ID, strExtData, rsCurr, lng�䷽ID)
            
            blnRefresh = True
        End If
        
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            vsAdvice.TextMatrix(i, col_ȱʡ) = strȱʡ
        Next
    
        'ˢ��ҽ����Ƭ
        If blnRefresh Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        
        mblnNoSave = True '���Ϊδ����
    End If
    
    Call vsAdvice.AutoSize(col_ҽ������)
    
    txtҽ������.SetFocus
End Sub

Private Sub ClinicSelecter(Optional ByVal lng����ID As Long)
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = frmClinicSelect.ShowSelect(Me, -1, 0, 0, cbo��Ч.ListIndex, "", , , mint��Χ, lng����ID, , , , mstrʹ�ÿ���, , mstr���Ʒ���, mstr��������, mstrִ�з���)
    If rsTmp Is Nothing Then 'ȡ����������
        zlControl.TxtSelAll txtҽ������
        txtҽ������.SetFocus: Exit Sub
    End If
        
    '����ѡ����Ŀ����ȱʡҽ����Ϣ
    If AdviceInput(rsTmp, vsAdvice.Row) Then
        '��ʾ��ȱʡ���õ�ֵ
        Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�ٴ��Թ�ҩ)) = 1 Then
            cboִ������.Tag = "1"
            Call AdviceChange
        End If
        txtҽ������.SetFocus '�����ȶ�λ
        Call SeekNextControl
    Else
        '�ָ�ԭֵ(AdviceInput�����п��ܴ�����һ��)
        txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������)
        txtҽ������.SetFocus
    End If
End Sub

Private Sub cmdSel_Click()
    ClinicSelecter
End Sub

Private Sub Form_Activate()
    If mblnRunFirst Then
        mblnRunFirst = False
        If cbo��Ч.Visible And cbo��Ч.Enabled Then
            cbo��Ч.SetFocus
        ElseIf txtҽ������.Enabled Then
            txtҽ������.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            If tbrFree.Visible And tbrFree.Enabled And tbrFree.Buttons(1).Enabled And tbrFree.Buttons(1).Visible Then
                tbrFree.Buttons(1).value = IIF(tbrFree.Buttons(1).value = 1, 0, 1)
                Call tbrFree_ButtonClick(tbrFree.Buttons(1))
            End If
        Case vbKeyF4
            If Me.ActiveControl Is txt�÷� Then
                If cmd�÷�.Visible And cmd�÷�.Enabled Then cmd�÷�_Click
            ElseIf Me.ActiveControl Is txtƵ�� Then
                If cmdƵ��.Visible And cmdƵ��.Enabled Then cmdƵ��_Click
            End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeySpace Then
        If Me.ActiveControl Is txt�÷� Then
            KeyAscii = 0
            If cmd�÷�.Visible And cmd�÷�.Enabled Then cmd�÷�_Click
        ElseIf Me.ActiveControl Is txtƵ�� Then
            KeyAscii = 0
            If cmdƵ��.Visible And cmdƵ��.Enabled Then cmdƵ��_Click
        ElseIf Me.ActiveControl Is cbo���� Then
            KeyAscii = 0
            If cbo����.Visible And cbo����.Enabled Then zlCommFun.PressKey (vbKeyF4)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lngRow As Long
    Dim strErr As String
    Dim strPre As String
    
    If gobjLIS Is Nothing Then
        '������ٴ�·������ף�����δ����LIS���������ȴ���
        Call InitObjLis(IIF(mint��Χ = 1, p����ҽ��վ, pסԺҽ��վ))
        If gobjLIS Is Nothing Then
            mblnNewLIS = False
        Else
            On Error Resume Next
            mblnNewLIS = gobjLIS.GetApplicationFormShowType
            err.Clear: On Error GoTo 0
        End If
    Else
        On Error Resume Next
        mblnNewLIS = gobjLIS.GetApplicationFormShowType
        err.Clear: On Error GoTo 0
    End If
    Call InitCommandBar
    Call InitAdviceTable
    Call RestoreWinState(Me, App.ProductName)
    
    Call Cbo.SetListHeight(cbo����, Me.Height)
    Call Cbo.SetListHeight(cboִ�п���, Me.Height)
    Call Cbo.SetListWidth(cboִ�п���.Hwnd, cboִ�п���.Width * 1.3)
    
    'ͼ��
    tbrFree.HotImageList = frmIcons.img24
    tbrFree.ImageList = frmIcons.img24
    tbrFree.Buttons(1).Image = 1
    
    tbrFree.Top = 810 '��ʼλ��
    tbrFree.Visible = Not (mint��Χ = 1)  '�������Ŀǰ��֧������¼��ҽ��
    
    If mbyt���� = 0 Then
        Me.Caption = "����ҽ��"
    ElseIf mbyt���� = 1 Then
        Me.Caption = "·��ҽ��"
    ElseIf mbyt���� = 2 Then
        Me.Caption = "�滻ҽ��"
        tbrFree.Visible = False '��������,��ֹ����¼��
    End If
    
    mblnOK = False
    mblnNoSave = False
    mblnRowMerge = False
    mblnRunFirst = True
    mblnRowChange = True
    mlngNextID = 0
        
    '����ƥ��
    mstrLike = IIF(Val(zldatabase.GetPara("����ƥ��")) = 0, "%", "")
    '����ƥ�䷽ʽ��0-ƴ��,1-���
    mint���� = Val(zldatabase.GetPara("���뷽ʽ"))
    
    '����ȱʡһ����
    mblnһ���� = Val(zldatabase.GetPara("����ȱʡһ����", glngSys, pסԺҽ���´�)) <> 0
    
    If mbyt���� <> 2 Then
        '���õ���
        strPre = cbo����.Text '����󱣳�ԭ��ֵ
        Call Load��Һ����(cbo����, lbl���ٵ�λ, False)
        cbo����.Text = strPre
    End If
    
    If Not mblnView Then
        '��������
        Call ReadEnjoin
        
        'ҽ�����ݶ���
        If CreateScript(mobjVBA, mobjScript) Then
            Set mrsDefine = InitAdviceDefine
        End If
    End If
    
    If mint��Χ = 1 Then
        lbl��Ч.Enabled = False
        cbo��Ч.ListIndex = 1
        cbo��Ч.Enabled = False
    End If
    
    '��ȡ����ʾ��������
    If Not mrsScheme Is Nothing Then
        Call LoadAdvice(0, vsAdvice.FixedRows)
    End If
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
End Sub

Private Sub InitSchemeRecordset()
    Set mrsScheme = New ADODB.Recordset
    mrsScheme.Fields.Append "�Ƿ�ȱʡ", adSmallInt
    mrsScheme.Fields.Append "�Ƿ�ѡ", adSmallInt
    mrsScheme.Fields.Append "���", adBigInt
    mrsScheme.Fields.Append "������", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "��Ч", adSmallInt
    mrsScheme.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    mrsScheme.Fields.Append "����", adSingle, , adFldIsNullable
    mrsScheme.Fields.Append "��������", adSingle, , adFldIsNullable
    mrsScheme.Fields.Append "�ܸ�����", adSingle, , adFldIsNullable
    mrsScheme.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    mrsScheme.Fields.Append "ִ��Ƶ��", adVarChar, 100, adFldIsNullable
    mrsScheme.Fields.Append "Ƶ�ʴ���", adSmallInt, , adFldIsNullable
    mrsScheme.Fields.Append "Ƶ�ʼ��", adSmallInt, , adFldIsNullable
    mrsScheme.Fields.Append "�����λ", adVarChar, 10, adFldIsNullable
    mrsScheme.Fields.Append "ʱ�䷽��", adVarChar, 100, adFldIsNullable
    mrsScheme.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "ִ������", adSmallInt
    mrsScheme.Fields.Append "�걾��λ", adVarChar, 100, adFldIsNullable
    mrsScheme.Fields.Append "��鷽��", adVarChar, 100, adFldIsNullable
    mrsScheme.Fields.Append "�䷽ID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "�����ĿID", adBigInt, , adFldIsNullable
    mrsScheme.Fields.Append "ִ�б��", adSingle, , adFldIsNullable
    If mbyt���� = 1 Then
        mrsScheme.Fields.Append "���", adVarChar, 1, adFldIsNullable
        mrsScheme.Fields.Append "��������", adVarChar, 20, adFldIsNullable
    End If
    mrsScheme.CursorLocation = adUseClient
    mrsScheme.LockType = adLockOptimistic
    mrsScheme.CursorType = adOpenStatic
    mrsScheme.Open
End Sub

Private Function ReadEnjoin() As Boolean
'���ܣ���ȡ�����볣������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strPre As String
        
    On Error GoTo errH
    
    strPre = cboҽ������.Text '����󱣳�ԭ��ֵ
    cboҽ������.Clear
    
    strSql = _
        " Select ���� From �������� Where ���� is Not Null And ��Ա=[1]" & _
        " Union" & _
        " Select ���� From �������� Where ���� is Not Null And ��Ա is Null" & _
        " Order by ����"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.����)
    Do While Not rsTmp.EOF
        AddComboItem cboҽ������.Hwnd, CB_ADDSTRING, 0, rsTmp!����
        rsTmp.MoveNext
    Loop
    cboҽ������.Text = strPre
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    msng���� = 0
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mrsDefine = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function RowCanMerge(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional strMsg As String) As Boolean
'���ܣ��ж������Ƿ����һ����ҩ
'������lngRow1=ǰ��һ���Ѿ������ҩƷ��
'      lngRow2=��ǰ��(�������δ����)
'���أ���������ԣ���strMsg������ʾ��Ϣ
    Dim lngFind As Long
    Dim lng�������� As Long
    Dim strҩ��IDs As String
    
    With vsAdvice
        strMsg = ""
        If Not Between(lngRow1, .FixedRows, .Rows - 1) Then Exit Function
        If Not Between(lngRow2, .FixedRows, .Rows - 1) Then Exit Function
        If .RowHidden(lngRow1) Or .RowHidden(lngRow2) Then Exit Function
        If .RowData(lngRow1) = 0 Then Exit Function
        
        If .RowData(lngRow2) = 0 Then
            '����ȫ��Ϊ��ҩ�������ͬ
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_���)) = 0 Then
                strMsg = "һ����ҩ��ҩƷ���붼Ϊ����ҩ��Ϊ�г�ҩ��"
                Exit Function
            End If
        ElseIf .RowData(lngRow2) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_���)) = 0 _
                Or InStr(",5,6,", .TextMatrix(lngRow2, COL_���)) = 0 Then
                strMsg = "һ����ҩ��ҩƷ���붼Ϊ����ҩ��Ϊ�г�ҩ��"
                Exit Function
            End If
            
            '��Ч������ͬ
            If .TextMatrix(lngRow1, COL_��Ч) <> .TextMatrix(lngRow2, COL_��Ч) Then
                strMsg = "һ����ҩ��ҩƷҽ����Ч������ͬ��"
                Exit Function
            End If
            
            'һ����ҩ(ǰ��ҩƷ)�ĸ�ҩ;���Ƿ������ڵ�ǰҩƷ
            lngFind = .FindRow(CLng(.TextMatrix(lngRow1, COL_���ID)), lngRow1 + 1)
            If lngFind <> -1 Then
                If Not Check�����÷�(Val(.TextMatrix(lngFind, COL_������ĿID)), Val(.TextMatrix(lngRow2, COL_������ĿID)), mint��Χ) Then
                    strMsg = """" & .TextMatrix(lngRow2, col_ҽ������) & """����ʹ��""" & .TextMatrix(lngFind, col_ҽ������) & """��ҩ;����" & _
                    vbCrLf & "������""" & .TextMatrix(lngRow1, col_ҽ������) & """����Ϊһ����ҩ��"
                    Exit Function
                End If
            End If
            
            '���������������ģ��Ƿ񶼿��Դ洢���Ա�ҩ������
            If Not (Val(.TextMatrix(lngRow1, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow1, COL_ִ������)) = 5) Then
                If sys.DeptHaveProperty(Val(.TextMatrix(lngRow1, COL_ִ�п���ID)), "��������") Then
                    lng�������� = Val(.TextMatrix(lngRow1, COL_ִ�п���ID))
                End If
            End If
            If lng�������� = 0 Then
                If Not (Val(.TextMatrix(lngRow2, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow2, COL_ִ������)) = 5) Then
                    If sys.DeptHaveProperty(Val(.TextMatrix(lngRow2, COL_ִ�п���ID)), "��������") Then
                        lng�������� = Val(.TextMatrix(lngRow2, COL_ִ�п���ID))
                    End If
                End If
            End If
            If lng�������� <> 0 Then
                If Not (Val(.TextMatrix(lngRow1, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow1, COL_ִ������)) = 5) Then
                    strҩ��IDs = Get����ҩ��IDs(.TextMatrix(lngRow1, COL_���), Val(.TextMatrix(lngRow1, COL_������ĿID)), Val(.TextMatrix(lngRow1, COL_�շ�ϸĿID)), 0, mint��Χ)
                    If InStr("," & strҩ��IDs & ",", "," & lng�������� & ",") = 0 Then
                        strMsg = "ҩƷ""" & .TextMatrix(lngRow1, col_ҽ������) & """����������""" & sys.RowValue("���ű�", lng��������, "����") & """û�д洢��"
                        Exit Function
                    End If
                End If
                If Not (Val(.TextMatrix(lngRow2, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow2, COL_ִ������)) = 5) Then
                    strҩ��IDs = Get����ҩ��IDs(.TextMatrix(lngRow2, COL_���), Val(.TextMatrix(lngRow2, COL_������ĿID)), Val(.TextMatrix(lngRow2, COL_�շ�ϸĿID)), 0, mint��Χ)
                    If InStr("," & strҩ��IDs & ",", "," & lng�������� & ",") = 0 Then
                        strMsg = "ҩƷ""" & .TextMatrix(lngRow2, col_ҽ������) & """����������""" & sys.RowValue("���ű�", lng��������, "����") & """û�д洢��"
                        Exit Function
                    End If
                End If
            End If
        End If
    End With
    RowCanMerge = True
End Function

Private Sub MoveCurrRow(ByVal lngRow As Long, ByVal lngWay As Long)
'���ܣ�����ǰ�����ƻ�����һ��
'������lngRow=��ǰ��
'      lngWay=1����һ��,-1����һ��(�൱����һ������һ��)
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngUpBegin As Long, lngUpEnd As Long
    Dim lngDownBegin As Long, lngDownEnd As Long
    Dim i As Long, j As Long
    Dim lngMoveRows As Long, blnRedraw As Boolean
   
    With vsAdvice
        If .RowData(lngRow) = 0 Then Exit Sub   '�հ����ų�
        '��ǰ�п�����һ����ҩ�м����
        Call GetRowScope(lngRow, lngBegin, lngEnd)
                
        If lngWay = 1 Then
            lngPreRow = GetPreRow(lngBegin)
            If lngPreRow = -1 Then Exit Sub
          
            lngDownBegin = lngBegin
            lngDownEnd = lngEnd
            Call GetRowScope(lngPreRow, lngUpBegin, lngUpEnd)
            lngMoveRows = lngDownBegin - lngUpBegin
        Else
            lngNextRow = GetNextRow(lngEnd)
            If lngNextRow = -1 Then Exit Sub
            
            lngUpBegin = lngBegin
            lngUpEnd = lngEnd
            Call GetRowScope(lngNextRow, lngDownBegin, lngDownEnd)
            lngMoveRows = lngDownEnd - lngUpEnd
        End If
        
        blnRedraw = .Redraw
        .Redraw = False
        
        j = 0
        For i = lngDownBegin To lngDownEnd
            .RowPosition(i) = lngUpBegin + j
            j = j + 1
        Next
               
        mblnRowChange = False
        lngRow = lngRow - lngWay * lngMoveRows
        .Row = lngRow
        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col): .TopRow = .Row
        mblnRowChange = True
         
        mblnNoSave = True '���Ϊδ����
        .Redraw = blnRedraw
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngҽ��ID As Long, lng���ID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngPreRow As Long, strMsg As String
    Dim str��Ч As String, lng������ĿID As Long
    Dim lngTmp As Long, i As Long, j As Long
    
    If Not Control.Visible Then Exit Sub 'Visible=Falseʱͨ���ȼ���ȻҲ��ִ��
    
    Call AdviceChange 'ǿ�Ƹ���ҽ������
    
    With vsAdvice
        Select Case Control.ID
            Case conMenu_MoveUp
                Call MoveCurrRow(.Row, 1)
            Case conMenu_MoveDown
                Call MoveCurrRow(.Row, -1)
            Case conMenu_New
                If .RowData(.Row) = 0 Then
                ElseIf .RowData(.Rows - 1) = 0 Then
                    .Row = .Rows - 1
                Else
                    '��ɾ���м����Ŀ���
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                End If
                
                Call .ShowCell(.Row, .Col)
                If Visible Then
                    If cbo��Ч.Visible And cbo��Ч.Enabled Then
                        cbo��Ч.SetFocus
                    ElseIf txtҽ������.Enabled Then
                        txtҽ������.SetFocus
                    End If
                End If
            Case conMenu_Insert
                If .RowData(.Row) = 0 Then
                    MsgBox "��ǰ�������ݣ������ڵ�ǰ��¼����Чҽ����", vbInformation, gstrSysName
                    Exit Sub
                End If
                            
                lngPreRow = GetPreRow(.Row)
                            
                '�������Զ���Ϊһ����ҩ:������һ����ҩ���м����
                If lngPreRow <> -1 Then
                    If Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) _
                        And Val(.TextMatrix(lngPreRow, COL_���ID)) <> 0 And InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                        
                        lng���ID = Val(.TextMatrix(lngPreRow, COL_���ID))
                    End If
                End If
                
                '��ɾ���м����Ŀ���
                mblnRowChange = False
                lngҽ��ID = .RowData(.Row)
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                .Row = .FindRow(lngҽ��ID)
                mblnRowChange = True
                            
                '��ǰ��֮ǰ��������
                '--------------------------------------------------------------
                If RowIn�䷽��(.Row) Or RowIn������(.Row) Then
                    '��ҩ�䷽�������������ǰ���������
                    lngBegin = .FindRow(CStr(.RowData(.Row)), , COL_���ID)
                Else
                    lngBegin = .Row
                End If
                
                mblnRowChange = False
                .AddItem "", lngBegin
                .Row = lngBegin
                .Col = .FixedCols
                mblnRowChange = True
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
                Call .ShowCell(.Row, .Col)
                
                If cbo��Ч.Visible And cbo��Ч.Enabled Then
                    cbo��Ч.SetFocus
                ElseIf txtҽ������.Enabled Then
                    txtҽ������.SetFocus
                End If
            Case conMenu_Merge 'һ����ҩ
                If Not Control.Checked Then '�밴��
                    lngBegin = GetPreRow(.Row)
                    'ǰ��û����
                    If lngBegin = -1 Then
                        MsgBox "ǰ��û�п���һ����ҩ��ҽ���С�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '���в���������
                    If Not RowCanMerge(lngBegin, .Row, strMsg) Then
                        MsgBox strMsg, vbInformation, gstrSysName
                        Exit Sub
                    End If
                    If .RowData(.Row) = 0 Then
                        '��ǰ����δ�������ݵ����
                        cbo��Ч.ListIndex = IIF(.TextMatrix(lngBegin, COL_��Ч) = "����", 1, 0)
                        mblnRowMerge = True: cbsMain.RecalcLayout '*������
                        txtҽ������.SetFocus: Exit Sub
                    Else
                        'Ҫ�ѵ�ǰ����ǰ����һ��һ����ҩ
                        Call MergeRow(lngBegin, .Row, False)
                    End If
                Else '�뵯��
                    If .RowData(.Row) = 0 Then
                        '�Ƿ�ǰ����δ�������ݵ����
                        If Not RowInһ����ҩ(.Row) Then
                            mblnRowMerge = False '*������
                            cbsMain.RecalcLayout
                        End If
                        Exit Sub
                    Else
                        '��ǰ����һ����ҩ�е���
                        Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
                                                
                        '����ʾ
                        If Not (.Row = lngEnd And lngEnd - lngBegin > 1) Then
                            '����һ����ҩȡ��Ϊ������ҩ
                            If MsgBox("Ҫ������һ����ҩ��ҩƷȫ��ȡ��Ϊ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                        'ɾ���м�Ŀ���
                        lngTmp = .RowData(.Row)
                        For i = lngEnd To lngBegin Step -1
                            If .RowData(i) = 0 Then
                                .RemoveItem i
                                lngEnd = lngEnd - 1
                            End If
                        Next
                        .Row = .FindRow(lngTmp, lngBegin)
                        
                        If .Row = lngEnd And lngEnd - lngBegin > 1 Then
                            '��һ����ҩ�з������
                            Call SplitRow(.Row)
                        Else
                            'ȡ��һ����ҩ
                            lngTmp = .RowData(.Row) '��¼���ڻָ��ж�λ
                            Call AdviceSet������ҩ(lngBegin, lngEnd)
                            .Row = .FindRow(lngTmp)
                        End If
                    End If
                End If
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
            Case conMenu_Delete
                If .RowSel <> .Row Then
                    MsgBox "һ��ֻ��ɾ��һ��ҽ������ѡ��Ҫɾ����ҽ���С�", vbInformation, gstrSysName
                    Exit Sub
                End If
                If .RowData(.Row) <> 0 Then
                    If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """��", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                'ɾ����ǰ��
                Call AdviceDelete(.Row)
                .SetFocus
            Case conMenu_Import '����ҽ��
                strMsg = frmSchemeImport.ShowMe(Me, mint��Χ, lngTmp)
                If strMsg <> "" And lngTmp <> 0 Then
                    Call LoadAdvice(0, 0, strMsg, lngTmp)
                    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
                    mblnNoSave = True
                End If
            Case conMenu_Save '����ҽ��
                If Not CheckAdvice Then Exit Sub '����д����˹�궨λ
                If Not SaveAdvice Then .SetFocus: Exit Sub
                Unload Me
            Case conMenu_Exit
                Unload Me
        End Select
    End With
End Sub

Private Sub Getһ����ҩ��Χ(ByVal lng���ID As Long, lngBegin As Long, lngEnd As Long)
'���ܣ�������صĸ�ҩ;��ҽ��ID,ȷ��һ����ҩ��һ��ҩƷ����ֹ�к�
'˵�����м���ܰ����п���
    Dim i As Long
    lngBegin = vsAdvice.FindRow(CStr(lng���ID), , COL_���ID)
    For i = lngBegin To vsAdvice.Rows - 1
        If Not vsAdvice.RowHidden(i) And vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lng���ID Then
                lngEnd = i
            Else
                Exit For
            End If
        End If
    Next
End Sub

Private Sub txt����_Change()
    txt����.Tag = "1"
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt����.Text) Or txt����.Text = "" Then
            If SeekNextControl Then Call txt����_Validate(False)
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim strMsg As String, dbl���� As Double, sng���� As Single
    
    With vsAdvice
        If Val(txt����.Text) = 0 Then txt����.Text = ""
        If Not IsNumeric(txt����.Text) Then
            If txt����.Text <> "" Then
                Cancel = True: txt����_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 And .TextMatrix(.Row, COL_��Ч) = "����" Then
'                '�ָ���Ϊ�����
'                If IsNumeric(.TextMatrix(.Row, COL_����)) Then
'                    txt����.Text = .TextMatrix(.Row, COL_����)
'                End If
            End If
        ElseIf CDbl(txt����.Text) <= 0 Then
            Cancel = True: txt����_GotFocus: Exit Sub
        ElseIf CDbl(txt����.Text) > LONG_MAX Then
            Cancel = True: txt����_GotFocus: Exit Sub
        Else
            '�����Ϸ��Լ��
            If txt����.Text <> "" And InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 And Val(.TextMatrix(.Row, COL_�շ�ϸĿID)) <> 0 Then
                dbl���� = IIF(Val(.TextMatrix(.Row, COL_����)) = 0, 1, Val(.TextMatrix(.Row, COL_����))) * _
                    Val(.TextMatrix(.Row, COL_��װϵ��)) * Val(.TextMatrix(.Row, COL_����ϵ��)) / Val(txt����.Text)
                If dbl���� > 200 Then
                    If MsgBox("��ҩƷ��ÿ�� " & FormatEx(txt����.Text, 5) & .TextMatrix(.Row, COL_������λ) & " ʹ�ã�" & _
                        IIF(Val(.TextMatrix(.Row, COL_����)) = 0, "ÿ", Val(.TextMatrix(.Row, COL_����))) & _
                        .TextMatrix(.Row, COL_��װ��λ) & "����ʹ�� " & FormatEx(dbl����, 5) & " �Ρ�" & _
                        vbCrLf & vbCrLf & "��ȷ�ϵ���������ȷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt����_GotFocus: Exit Sub
                    End If
                End If
            End If
            
            txt����.Text = FormatEx(txt����.Text, 5)
            
            '���¼���ҩƷ����(�����뵥��ʱ)
            If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 And .TextMatrix(.Row, COL_��Ч) = "����" Then
                If .TextMatrix(.Row, COL_Ƶ��) <> "" And Val(.TextMatrix(.Row, COL_Ƶ������)) <> 1 _
                    And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 And Val(.TextMatrix(.Row, COL_��װϵ��)) <> 0 Then
                    
                    sng���� = Val(.TextMatrix(.Row, COL_����))
                    If sng���� = 0 Then sng���� = 1
                    
                    txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                        Val(txt����.Text), sng����, _
                        Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), _
                        .TextMatrix(.Row, COL_�����λ), .TextMatrix(.Row, COL_ִ��ʱ��), _
                        Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_��װϵ��)), _
                        Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                    txt����.Tag = "1"
                End If
            End If
        End If
        
        '��������
        Call AdviceChange
    End With
End Sub

Private Sub cboҽ������_Change()
    cboҽ������.Tag = "1"
End Sub

Private Sub cboҽ������_Click()
    cboҽ������.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cboҽ������_GotFocus()
    zlControl.TxtSelAll cboҽ������
End Sub

Private Sub cboҽ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cboҽ������_Validate(False)
    Else
        Call Cbo.AppendText(cboҽ������, KeyAscii)
    End If
End Sub

Private Sub cboҽ������_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cboҽ������.Text) > 100 Then
        MsgBox "�������ݲ������� 50 �����ֻ� 100 ���ַ���", vbInformation, gstrSysName
        cboҽ������_GotFocus
        Cancel = True: Exit Sub
    End If
    
    '��������
    Call AdviceChange
End Sub

Private Sub txtҽ������_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txtҽ������_GotFocus()
    Call zlControl.TxtSelAll(txtҽ������)
End Sub

Private Sub txtҽ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txtҽ������)
    End If
End Sub

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim bytƥ�� As Byte
    If KeyAscii = 13 Then
        
        KeyAscii = 0
        If txtҽ������.Text = "" Then Exit Sub
        If txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������) Then
            Call SeekNextControl
            Exit Sub
        End If
        
        If tbrFree.Buttons(1).value = 0 Then
            Set rsTmp = frmClinicSelect.ShowSelect(Me, -1, 0, 0, cbo��Ч.ListIndex, "", txtҽ������.Text, txtҽ������, mint��Χ, , , , , mstrʹ�ÿ���, bytƥ��, mstr���Ʒ���, mstr��������, mstrִ�з���)
            If rsTmp Is Nothing Then 'ȡ����������
                '�ָ�ԭֵ
                txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������)
                zlControl.TxtSelAll txtҽ������
                txtҽ������.SetFocus: Exit Sub
            ElseIf bytƥ�� = 1 Then
                Call Cbo.SetIndex(cbo��Ч.Hwnd, IIF(cbo��Ч.ListIndex = 0, 1, 0))
            End If
            '����Ŀ��¼��
            '������Ŀ�����������ҩ,���ܰ������ҽ��
            
            '����ѡ����Ŀ����ȱʡҽ����Ϣ
            Me.Refresh
            If AdviceInput(rsTmp, vsAdvice.Row) Then
                '��ʾ��ȱʡ���õ�ֵ
                Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
                If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�ٴ��Թ�ҩ)) = 1 Then
                    cboִ������.Tag = "1"
                    Call AdviceChange
                End If
                Call SeekNextControl
            Else
                '�ָ�ԭֵ
                txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������)
                zlControl.TxtSelAll txtҽ������
                txtҽ������.SetFocus: Exit Sub
        End If
        ElseIf tbrFree.Buttons(1).value = 1 Then
            If txtҽ������.Text <> "" Then
                If zlCommFun.ActualLen(txtҽ������.Text) > txtҽ������.MaxLength Then
                    MsgBox "�������ݲ������� " & txtҽ������.MaxLength \ 2 & " �����ֻ� " & txtҽ������.MaxLength & " ���ַ���", vbInformation, gstrSysName
                    Call txtҽ������_GotFocus: Exit Sub
                End If
                Call AdviceInputFree(vsAdvice.Row)
                Call SeekNextControl
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdSel.Visible And cmdSel.Enabled Then Call cmdSel_Click
    End If
End Sub

Private Sub cboִ��ʱ��_GotFocus()
    zlControl.TxtSelAll cboִ��ʱ��
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    If tbrFree.Buttons(1).value = 0 Then
        '�ָ���Ϊ�ĸı�
        If txtҽ������.Text <> vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������) Then
            txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������)
        End If
    ElseIf tbrFree.Buttons(1).value = 1 Then
        If vsAdvice.RowData(vsAdvice.Row) <> 0 And txtҽ������.Text = "" Then
            '��Ϊ����¼��,�����Զ��ָ�
            txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������)
            Exit Sub
        End If
        
        If txtҽ������.Text <> "" Then
            If zlCommFun.ActualLen(txtҽ������.Text) > txtҽ������.MaxLength Then
                MsgBox "�������ݲ������� " & txtҽ������.MaxLength \ 2 & " �����ֻ� " & txtҽ������.MaxLength & " ���ַ���", vbInformation, gstrSysName
                Call txtҽ������_GotFocus: Cancel = True: Exit Sub
            End If
            Call AdviceInputFree(vsAdvice.Row)
        End If
    End If
End Sub

Private Sub txt����_Change()
    txt����.Tag = "1"
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt����.Text) Or txt����.Text = "" Then
            If SeekNextControl Then Call txt����_Validate(False)
        End If
    Else
        If RowIn�䷽��(vsAdvice.Row) Then
            strMask = "0123456789" '��ҩ�䷽ֻ����������
        Else
            strMask = "0123456789."
        End If
        If InStr(strMask & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim strMsg As String, sng���� As Single
    Dim dbl���� As Double, bln�䷽�� As Boolean
    
    With vsAdvice
        If Val(txt����.Text) = 0 Then txt����.Text = ""
        If Not IsNumeric(txt����.Text) Then
            If txt����.Text <> "" Then
                Cancel = True: txt����_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 Then
'                '�ָ���Ϊ�����
'                If IsNumeric(.TextMatrix(.Row, COL_����)) Then
'                    txt����.Text = .TextMatrix(.Row, COL_����)
'                End If
            End If
        ElseIf CDbl(txt����.Text) <= 0 Then
            Cancel = True: txt����_GotFocus: Exit Sub
        ElseIf CDbl(txt����.Text) > LONG_MAX Then
            Cancel = True: txt����_GotFocus: Exit Sub
        Else
            txt����.Text = FormatEx(txt����.Text, 5)
        End If
        
        bln�䷽�� = RowIn�䷽��(.Row)
        
        If IsNumeric(txt����.Text) Then
            If bln�䷽�� Then
                txt����.Text = CInt(txt����.Text)
            End If
        End If
        
        '�����������
        If txt����.Text <> "" And InStr(",4,5,6,", .TextMatrix(.Row, COL_���)) > 0 And .TextMatrix(.Row, COL_��Ч) = "����" Then
            If .TextMatrix(.Row, COL_Ƶ��) <> "" _
                And Val(.TextMatrix(.Row, COL_����)) <> 0 _
                And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                And Val(.TextMatrix(.Row, COL_��װϵ��)) <> 0 Then
                
                If Val(.TextMatrix(.Row, COL_Ƶ������)) = 1 Then
                    dbl���� = FormatEx(CalcȱʡҩƷ����( _
                        Val(.TextMatrix(.Row, COL_����)), 1, 1, 1, "��", "", _
                        Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_��װϵ��)), _
                        Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                Else
                    sng���� = Val(.TextMatrix(.Row, COL_����))
                    If sng���� = 0 Then sng���� = 1
                    
                    dbl���� = FormatEx(CalcȱʡҩƷ����( _
                        Val(.TextMatrix(.Row, COL_����)), sng����, _
                        Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), _
                        .TextMatrix(.Row, COL_�����λ), .TextMatrix(.Row, COL_ִ��ʱ��), _
                        Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_��װϵ��)), _
                        Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                End If
                If Val(txt����.Text) < dbl���� Then
                    If MsgBox(.TextMatrix(.Row, COL_����) & "��ÿ�� " & .TextMatrix(.Row, COL_����) & .TextMatrix(.Row, COL_������λ) & "," & _
                        .TextMatrix(.Row, COL_Ƶ��) & IIF(Val(.TextMatrix(.Row, COL_Ƶ������)) <> 1 _
                            And Val(.TextMatrix(.Row, COL_����)) > 0 And .TextMatrix(.Row, COL_���) <> "4", ",��ҩ " & sng���� & " ��", "") & _
                        "ִ��ʱ,������Ҫ " & FormatEx(dbl����, 5) & .TextMatrix(.Row, COL_������λ) & ",Ҫ������", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt����_GotFocus: Exit Sub
                    End If
                End If
            End If
        End If
        
        '��������
        Call AdviceChange
    End With
End Sub

Private Sub ClearAdviceCard()
'���ܣ����ҽ����ʾ��Ƭ��ص�����
'������bln��ʼʱ��=�Ƿ������ʼʱ��
    Call SetCardEditable(True)
    
    txtҽ������.Text = ""
    cboҽ������.Text = ""
    cboִ�п���.Clear
    cbo����ִ��.Clear
    
    cmdExt.Enabled = False
    Call SetDayState(-1, -1)
    Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1)
End Sub

Private Sub SetCardEditable(ByVal Editable As Boolean)
'���ܣ�����ɫ��ʶ��ǰҽ���Ƿ���Ա༭
    Dim obj As Object
    
    For Each obj In Controls
        If InStr("Label;TextBox;ComboBox;CheckBox", TypeName(obj)) > 0 Then
            If Not obj.Container Is Nothing Then
                If obj.Container Is fraAdvice Then
                    If Editable Then
                        obj.ForeColor = Me.ForeColor
                    Else
                        obj.ForeColor = &H808080
                    End If
                End If
            End If
        End If
    Next
    fraAdvice.Enabled = Editable
    cmdSel.Enabled = fraAdvice.Enabled
End Sub

Private Sub SetDayState(Optional ByVal intVisible As Integer, Optional ByVal intEnabled As Integer)
'���ܣ�����ִ���������úͻ��״̬
'������0-���ֲ���,-1-��ֹ,1-����
    If intEnabled = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
    ElseIf intEnabled = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If
    
    If intVisible = -1 Then
        lbl����.Visible = False
        txt����.Visible = False
        txt����.Text = ""
        
        lbl����.Left = lbl�÷�.Left + lbl�÷�.Width - lbl����.Width
        txt����.Left = txt�÷�.Left
        txt����.Width = txt�÷�.Width - cmd�÷�.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        lbl����.Left = lblƵ��.Left + lblƵ��.Width - lbl����.Width
        txt����.Left = txtƵ��.Left
        txt����.Width = txtƵ��.Width - cmdƵ��.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        txt����.TabIndex = cmdƵ��.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
    ElseIf intVisible = 1 Then
        lbl����.Visible = True
        txt����.Visible = True
        
        lbl����.Left = lbl�÷�.Left + lbl�÷�.Width - lbl����.Width
        txt����.Left = txt�÷�.Left
        txt����.Width = txt�÷�.Width - txt����.Width - Me.TextWidth("������!") - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        lbl����.Left = lblƵ��.Left + lblƵ��.Width - lbl����.Width
        txt����.Left = txtƵ��.Left
        txt����.Width = txtƵ��.Width - cmdƵ��.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        txt����.TabIndex = cmdƵ��.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
    End If
End Sub

Private Function GetƵ�ʷ�Χ(ByVal lngRow As Long) As Integer
    Dim lngFind As Long
    
    With vsAdvice
        If RowIn�䷽��(lngRow) Then
            GetƵ�ʷ�Χ = 2 '��ҽ
        Else
            If RowIn������(lngRow) Then '�Լ�����Ŀ��Ϊ׼
                lngFind = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                If lngFind <> -1 Then lngRow = lngFind
            End If
            If Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Then
                GetƵ�ʷ�Χ = 1 '��ѡƵ�ʵ���Ŀʹ����ҽƵ����Ŀ
            ElseIf Val(.TextMatrix(lngRow, COL_Ƶ������)) = 1 Then
                GetƵ�ʷ�Χ = -1 'һ����
            ElseIf Val(.TextMatrix(lngRow, COL_Ƶ������)) = 2 Then
                GetƵ�ʷ�Χ = -2 '������
            End If
        End If
    End With
End Function

Private Function SeekVisibleRow() As Boolean
'���ܣ���ǰ��Ϊ������ʱ����λ���������Ŀɼ���
    Dim lngRow As Long
    
    With vsAdvice
        If Not .RowHidden(.Row) Then Exit Function
        If InStr(",F,G,C,D,E,", .TextMatrix(.Row, COL_���)) > 0 And Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_���ID))))
        ElseIf .TextMatrix(.Row, COL_���) = "7" Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_���ID))))
        ElseIf .TextMatrix(.Row, COL_���) = "E" And Val(.TextMatrix(.Row, COL_���ID)) = 0 Then
            lngRow = .Row - 1
        End If
        If lngRow <> -1 Then
            If .RowData(lngRow) <> 0 Then
                .Row = lngRow: SeekVisibleRow = True
            End If
        End If
    End With
End Function

Private Sub SetCboִ������(ByVal bln���� As Boolean, ByVal bln���Ա�ҩ As Boolean, ByVal bln�ٴ��Թ�ҩ As Boolean)
    cboִ������.Clear
    
    If bln�ٴ��Թ�ҩ Then
        cboִ������.AddItem "1-�Ա�ҩ"
    Else
        If bln���� Then
            cboִ������.AddItem "0-����"
            If bln���Ա�ҩ Then cboִ������.AddItem "1-�Ա�ҩ"
            cboִ������.AddItem "2-��Ժ��ҩ"
            cboִ������.AddItem "3-��ȡҩ"
            cboִ������.AddItem "4-��ȡҩ"
        Else
            cboִ������.AddItem "0-����"
            If bln���Ա�ҩ Then cboִ������.AddItem "1-�Ա�ҩ"
            cboִ������.AddItem "4-��ȡҩ"
        End If
    End If
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'����һ��ҽ���е�"ȱʡ"��״̬

    If Col = col_ȱʡ Or Col = col_��ѡ Then
        Dim i As Long, lng��ID As Long, lngThis��ID As Long
        Dim lngBegin As Long, lngEnd As Long
        
        With vsAdvice
            
            'һ����ҩ��һ�����ã���������ҳ���ʼ��
            If Not RowInһ����ҩ(Row) Then
                Call GetRowScope(Row, lngBegin, lngEnd)
            Else
                Call Getһ����ҩ��Χ(Val(.TextMatrix(Row, COL_���ID)), lngBegin, lngEnd)
            End If
            
            For i = lngBegin To lngEnd
                If i <> Row Then
                    .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                End If
                If Col = col_��ѡ And .TextMatrix(Row, Col) = -1 And mbln��ʾȱʡ�� Then
                    .TextMatrix(i, col_ȱʡ) = 0
                End If
                If Col = col_ȱʡ And .TextMatrix(Row, Col) = -1 And mbln��ʾȱʡ�� Then
                    .TextMatrix(i, col_��ѡ) = 0
                End If
            Next
            mblnNoSave = True
        End With
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'���ܣ����иı�ʱ�����¿�Ƭ����
    Dim rsItem As New ADODB.Recordset
    Dim strSql As String, lngRow As Long
    Dim lng�÷�ID As Long, blnEditable As Boolean
    Dim lngҩƷID As Long, lngBaseRow As Long    '��ҩ�䷽�ĵ�һζ���ҩ��
    Dim dblPrice As Double, strTmp As String, i As Long

    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, col_ҽ������)
    End If

    If NewRow = OldRow Then Exit Sub
    If Not mblnRowChange Then Exit Sub
    If SeekVisibleRow Then Exit Sub

    lngRow = NewRow

    '��ǰ���ǿ���ʱ�����ǰһ����һ����ҩ�У���ȱʡ���¡�һ������ť
    If vsAdvice.RowData(lngRow) = 0 Then
        i = GetPreRow(lngRow)
        If i = -1 Then
            mblnRowMerge = False
        Else
            mblnRowMerge = RowInһ����ҩ(i)
        End If
    Else
        mblnRowMerge = RowInһ����ҩ(lngRow)
    End If
    cbsMain.RecalcLayout    '*��ʱˢ��

    Me.Refresh
    zlControl.FormLock Me.Hwnd

    On Error GoTo errH
    chkMedicineVariety.Visible = True

    With vsAdvice
        If .RowData(lngRow) = 0 Then
            '��Ч�������Ƭ����
            Call ClearAdviceCard

            'ȱʡΪ������¼��
            tbrFree.Buttons(1).value = 0
            tbrFree.Buttons(1).Enabled = Not RowInһ����ҩ(lngRow)
            tbrFree.Buttons(1).Image = IIF(tbrFree.Buttons(1).Enabled, 1, 2)

            'ȱʡ��Ч������һ�е���ʾ
            i = GetPreRow(lngRow)
            If i = -1 Or Not Visible Then
                Call Cbo.SetIndex(cbo��Ч.Hwnd, 1)    'ȱʡΪ����
            Else
                Call Cbo.SetIndex(cbo��Ч.Hwnd, IIF(.TextMatrix(i, COL_��Ч) = "����", 0, 1))
            End If
        ElseIf Val(.TextMatrix(lngRow, COL_������ĿID)) = 0 Then
            '����¼��ҽ��
            blnEditable = Not mblnView
            Call SetCardEditable(blnEditable)

            tbrFree.Buttons(1).value = 1
            tbrFree.Buttons(1).Enabled = blnEditable
            tbrFree.Buttons(1).Image = IIF(blnEditable, 1, 2)
            cmdExt.Enabled = False
            cmdSel.Enabled = False
            chkMedicineVariety.Visible = False

            '�������������
            Call SetDayState(-1, -1)
            SetItemEditable -1, -1, -1, -1, -1, , -1, -1, -1

            '��ʾ��ǰҽ����Ƭ����
            '--------------------------------------------------------------------------------------------
            Call Cbo.SetIndex(cbo��Ч.Hwnd, IIF(.TextMatrix(lngRow, COL_��Ч) = "����", 0, 1))

            'ҽ������
            txtҽ������.Text = .TextMatrix(lngRow, col_ҽ������)

            'ҽ������
            cboҽ������.Text = .TextMatrix(lngRow, COL_ҽ������)

            '��ѡִ�п���
            SetItemEditable , , , , , 1
            Call Get����ִ�п���(cboִ�п���, "*", 0, 0, 4, Val(.TextMatrix(lngRow, COL_ִ�п���ID)), cbo��Ч.ListIndex, mint��Χ)
        Else
            '��Ƭ�༭����У�Ե�ҽ�������޸�,��¼ҽ��ʱ���ܸ��ķǲ�¼������
            blnEditable = Not mblnView
            Call SetCardEditable(blnEditable)

            '����������Ŀ�����ɱ�Ϊ����¼��
            tbrFree.Buttons(1).value = 0
            tbrFree.Buttons(1).Enabled = False
            tbrFree.Buttons(1).Image = 2


            '��ȡ������Ŀ������Ϣ
            '---------------------
            chkMedicineVariety.Tag = "�����"
            If InStr("4,5,6", Val(.TextMatrix(lngRow, COL_���))) > 0 Then
                lngҩƷID = Val(.TextMatrix(lngRow, COL_�շ�ϸĿID))
                chkMedicineVariety.value = IIF(lngҩƷID = 0, 1, 0)
            Else
                chkMedicineVariety.Visible = False
            End If
            chkMedicineVariety.Tag = ""

            If RowIn�䷽��(lngRow) Then
                txt����.MaxLength = 3
                '��ȡ��ҩ�䷽��һζ��ҩ��
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                lngҩƷID = Val(.TextMatrix(lngBaseRow, COL_�շ�ϸĿID))
            ElseIf RowIn������(lngRow) Then
                '��ȡһ�������ĵ�һ����Ŀ��
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                txt����.MaxLength = txt����.MaxLength
            Else
                lngBaseRow = lngRow
                txt����.MaxLength = txt����.MaxLength
            End If
            Set rsItem = Get������Ŀ��¼(Val(.TextMatrix(lngBaseRow, COL_������ĿID)))

            '��չ��ť����״̬(������,�������,����,��ҩ�䷽)
            cmdExt.Enabled = InStr(",7,C,F,D,", rsItem!���) > 0

            '��ʾ��ǰҽ����Ƭ����
            '--------------------------------------------------------------------------------------------
            Call Cbo.SetIndex(cbo��Ч.Hwnd, IIF(.TextMatrix(lngRow, COL_��Ч) = "����", 0, 1))
            'ҽ������
            txtҽ������.Text = .TextMatrix(lngRow, col_ҽ������)

            '����
            '----------------------
            If rsItem!��� = "7" Then    '��ҩ�䷽(�в�ҩ)��Ȼ�е���,������������д
                SetItemEditable -1
            ElseIf cbo��Ч.ListIndex = 0 Then
                '��������ҩ���ʱ,������Ŀ����¼��
                If InStr(",1,2,", NVL(rsItem!���㷽ʽ, 0)) > 0 Or InStr(",5,6,", rsItem!���) > 0 Then
                    SetItemEditable 1
                    txt����.Text = .TextMatrix(lngRow, COL_����)
                    lbl������λ.Caption = .TextMatrix(lngRow, COL_������λ)
                Else
                    SetItemEditable -1
                End If
            ElseIf cbo��Ч.ListIndex = 1 Then
                '����:��ҩ���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����¼��(ע������ԭʼƵ��,��ǰ��������һ����)
                If (NVL(rsItem!ִ��Ƶ��, 0) = 0 And InStr(",1,2,", NVL(rsItem!���㷽ʽ, 0)) > 0) _
                   Or InStr(",5,6,", rsItem!���) > 0 Then
                    SetItemEditable 1
                    txt����.Text = .TextMatrix(lngRow, COL_����)
                    lbl������λ.Caption = .TextMatrix(lngRow, COL_������λ)
                Else
                    SetItemEditable -1
                End If
            End If

            '��������ҩ���г�ҩ������ʹ�ã����ڼ�������
            'һ�㣺������ҩƷ(����ҩ)���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����ʹ���������Զ���������
            blnEditable = False
            If cbo��Ч.ListIndex = 1 And InStr(",5,6,", rsItem!���) > 0 Then
                If Val(.TextMatrix(lngRow, COL_Ƶ������)) <> 1 Then blnEditable = True
            End If
            If blnEditable Then
                SetDayState 1, 1
            Else
                SetDayState -1, -1
            End If
            txt����.Text = Val(.TextMatrix(lngRow, COL_����))
            If Val(txt����.Text) = 0 Then txt����.Text = ""

            '����
            '--------------------
            If rsItem!��� = "7" Then
                '��ҩ�䷽(�в�ҩ)��дΪ����
                If cbo��Ч.ListIndex = 1 Then
                    SetItemEditable , 1
                Else
                    SetItemEditable , -1    '�䷽�����������������������������������(�µ�Ϊ�������ܵ������̶�Ϊ1����������)
                End If
                lbl������λ.Caption = "��"
                txt����.Text = .TextMatrix(lngRow, COL_����)    '����

            ElseIf cbo��Ч.ListIndex = 1 Then
                '��������Ҫ��д����:��������������Ϊ׼
                If rsItem!��� = "Z" And NVL(rsItem!��������) <> "0" Then
                    SetItemEditable , -1    '����ҽ���������޸�����(�̶�Ϊ1��)
                ElseIf InStr(",5,6,", rsItem!���) = 0 And NVL(rsItem!���㷽ʽ, 0) = 3 _
                       And (NVL(rsItem!ִ��Ƶ��, 0) = 1 Or Val(.TextMatrix(lngRow, COL_Ƶ������)) = 1) Then
                    SetItemEditable , -1    '��ҩƷһ���Լƴ���Ŀ����������(ԭʼƵ��Ϊһ���Ի�ǰ����Ϊһ����)
                Else
                    SetItemEditable , 1
                End If
                lbl������λ.Caption = .TextMatrix(lngRow, COL_������λ)
                txt����.Text = .TextMatrix(lngRow, COL_����)
            Else
                '����������������д����
                SetItemEditable , -1
            End If

            '��ҩ;������ҩ�÷�
            '--------------
            If InStr(",5,6,", rsItem!���) > 0 Then
                SetItemEditable , , 1
                lbl�÷�.Caption = "��ҩ;��"
                '���Ҹ�ҩ;����Ӧ����:���ҵ�Rowdata(Variant)����ҪתΪLong��,���ܾ�ȷƥ��
                lng�÷�ID = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                lng�÷�ID = Val(.TextMatrix(lng�÷�ID, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt�÷�.Text = sys.RowValue("������ĿĿ¼", lng�÷�ID, "����")
            ElseIf rsItem!��� = "K" Then
                '��Ѫҽ����Ҫ������ǰû����Ѫ;�������
                lng�÷�ID = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
                If lng�÷�ID <> -1 Then
                    SetItemEditable , , 1
                    If Val(.TextMatrix(lngRow, COL_��鷽��)) = 0 And gblnѪ��ϵͳ = True Then
                        lbl�÷�.Caption = "�ɼ�����"
                    Else
                        lbl�÷�.Caption = "��Ѫ;��"
                    End If
                    lng�÷�ID = Val(.TextMatrix(lng�÷�ID, COL_������ĿID))
                    cmd�÷�.Tag = lng�÷�ID
                    txt�÷�.Text = sys.RowValue("������ĿĿ¼", lng�÷�ID, "����")
                Else
                    SetItemEditable , , -1
                End If
            ElseIf rsItem!��� = "7" Then
                SetItemEditable , , 1
                lbl�÷�.Caption = "��ҩ�÷�"

                '��ҩ�䷽��ʾ�о�����ҩ�÷���
                lng�÷�ID = Val(.TextMatrix(lngRow, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt����֤��.Text = .TextMatrix(lngRow, COL_����֤��)
                txt�÷�.Text = sys.RowValue("������ĿĿ¼", lng�÷�ID, "����")
            ElseIf RowIn������(lngRow) Then    '��������ж�,������ǰ�ļ���
                '�������
                SetItemEditable , , 1
                lbl�÷�.Caption = "�ɼ�����"

                '���������ʾ�о��ǲɼ�������
                lng�÷�ID = Val(.TextMatrix(lngRow, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt�÷�.Text = sys.RowValue("������ĿĿ¼", lng�÷�ID, "����")
            Else
                SetItemEditable , , -1
            End If

            If rsItem!��� = "7" And mbyt���� = 1 Then
                SetItemEditable , , , , , , , , 1
            Else
                SetItemEditable , , , , , , , , -1
            End If
            
            '���٣���Һ���ҩ;����ҩƷ��������
            If InStr(",5,6,", rsItem!���) > 0 And mbyt���� <> 2 Then
                i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                If Val(.TextMatrix(i, COL_ִ�з���)) = 1 Then
                    SetItemEditable , , , , , , , , , 1
                    If InStr(.TextMatrix(i, COL_ҽ������), "��/����") > 0 Then
                        lbl���ٵ�λ.Caption = "��/����"
                    ElseIf InStr(.TextMatrix(i, COL_ҽ������), "����/Сʱ") > 0 Then
                        lbl���ٵ�λ.Caption = "����/Сʱ"
                    End If
                    Call Load��Һ����(cbo����, lbl���ٵ�λ, False)
                    cbo����.Text = Replace(.TextMatrix(i, COL_ҽ������), lbl���ٵ�λ.Caption, "")
                Else
                    SetItemEditable , , , , , , , , , -1
                End If
            Else
                SetItemEditable , , , , , , , , , -1
            End If
       
            
            If mbyt���� <> 2 Then
                'Ƶ�ʣ�������ѡ��(������������ָ��ʹ��)
                If True Then
                    SetItemEditable , , , 1
                    cmdƵ��.Tag = .TextMatrix(lngRow, COL_Ƶ��)
                    txtƵ��.Text = .TextMatrix(lngRow, COL_Ƶ��)
                Else
                    SetItemEditable , , , -1
                End If
    
                'ִ��ʱ�䣺"��ѡƵ��"��ҩƷ(��ǰδ������Ϊһ����)����"����"���ִ�е�
                If NVL(rsItem!ִ��Ƶ��, 0) = 0 And Val(.TextMatrix(lngBaseRow, COL_Ƶ������)) <> 1 And .TextMatrix(lngRow, COL_�����λ) <> "����" Then
                    SetItemEditable , , , , 1
                    Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(lngRow), .TextMatrix(lngRow, COL_Ƶ��), lng�÷�ID)
                    cboִ��ʱ��.Text = .TextMatrix(lngRow, COL_ִ��ʱ��)
                Else
                    SetItemEditable , , , , -1
                End If
    
                'ҽ������
                cboҽ������.Text = .TextMatrix(lngRow, COL_ҽ������)
    
                'ִ������:����Ŀǰ����ʹ��"�Ա�ҩ"
                If InStr(",5,6,7,", rsItem!���) > 0 Then
                    '������Թ�ҩ��̶�ѡ���Ա�ҩ
                    If Val(.TextMatrix(lngRow, COL_�ٴ��Թ�ҩ)) = 1 Then
                        strTmp = "�Ա�ҩ"
                    Else
                        If rsItem!��� = "7" Then
                            '������ҩ�䷽,����������Ŀ���������Ƽ���������,�������÷��ͼ巨һ��ΪԺ��ִ��,һ����Ϊ
                            If Val(.TextMatrix(lngBaseRow, COL_ִ������)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������)) <> 5 Then
    
                                strTmp = IIF(Val(.TextMatrix(lngBaseRow, COL_ִ�б��)) = 2, "��ȡҩ", "�Ա�ҩ")
    
                            ElseIf Val(.TextMatrix(lngBaseRow, COL_ִ������)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                                strTmp = "��Ժ��ҩ"
                            Else
                                strTmp = IIF(Val(.TextMatrix(lngBaseRow, COL_ִ�б��)) = 0, "����", "��ȡҩ")
                            End If
                        Else
                            i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                            If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                If Val(.TextMatrix(lngRow, COL_ִ�б��)) = 2 Then
                                    strTmp = "��ȡҩ"
                                Else
                                    strTmp = "�Ա�ҩ"
                                End If
                            ElseIf Val(.TextMatrix(lngRow, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                strTmp = "��Ժ��ҩ"
                            Else
                                strTmp = IIF(Val(.TextMatrix(lngRow, COL_ִ�б��)) = 0, "����", "��ȡҩ")
                            End If
                        End If
                    End If
    
                    Call SetCboִ������(cbo��Ч.ListIndex = 1, gbln����ҩ��ʹ���Ա�ҩ Or Not gblnKSSStrict Or Val(.TextMatrix(lngRow, COL_�����ȼ�)) = 0, Val(.TextMatrix(lngRow, COL_�ٴ��Թ�ҩ)) = 1)
                    SetItemEditable , , , , , , 1
                    Call Cbo.SetIndex(cboִ������.Hwnd, Cbo.FindIndex(cboִ������, strTmp, True))
                Else
                    SetItemEditable , , , , , , -1
                End If
    
                lblִ�п���.Caption = "ִ�п���"
                'ִ�п���
                If rsItem!��� = "Z" And NVL(rsItem!��������, 0) = 3 Then
                    'ת��ҽ�����ٴ�����
                    SetItemEditable , , , , , 1
                    lblִ�п���.Caption = "ת�����"
                    Call Get�ٴ�����(mint��Χ, 0, Val(.TextMatrix(lngRow, COL_ִ�п���ID)), cboִ�п���, True)
                ElseIf rsItem!��� = "Z" And NVL(rsItem!��������, 0) = 7 Then
                    '����ҽ�����ٴ�����
                    SetItemEditable , , , , , 1
                    lblִ�п���.Caption = "�������"
                    Call Get�ٴ�����(mint��Χ, 0, Val(.TextMatrix(lngRow, COL_ִ�п���ID)), cboִ�п���)
                Else
                    '��ҩƷ����ҩƷ��Ϊ׼��ʾ,��������Լ�����ĿΪ׼��ʾ
                    i = lngRow
                    If rsItem!��� = "7" Then
                        i = lngBaseRow
                    ElseIf RowIn������(lngRow) Then    '��������ж�,������ǰ�ļ���
                        i = lngBaseRow
                    End If
    
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                        '�Ƕ�����Ժ��ִ��ʱ����ʾ�Ϳ���ѡ��(����ҩƷ)
                        SetItemEditable , , , , , 1
                        Call Get����ִ�п���(cboִ�п���, rsItem!���, rsItem!ID, lngҩƷID, NVL(rsItem!ִ�п���, 0), Val(.TextMatrix(i, COL_ִ�п���ID)), cbo��Ч.ListIndex, mint��Χ)
    
                        '��ɢװ��̬��ֻ�������䷽����ѡҩ��
                        If rsItem!��� = "7" Then
                            If Val(.TextMatrix(lngRow, COL_��ҩ��̬)) <> 0 Then
                                cboִ�п���.Enabled = False
                                cboִ�п���.BackColor = Me.BackColor
                            End If
                        End If
    
                    ElseIf InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                        SetItemEditable , , , , , -1
                        If Val(.TextMatrix(i, COL_ִ������)) = 0 Then
                            cboִ�п���.AddItem "<��ִ�ж���>"
                        Else
                            cboִ�п���.AddItem "-"
                        End If
                        Call Cbo.SetIndex(cboִ�п���.Hwnd, 0)
                    End If
                    If InStr("5,6,7", rsItem!���) > 0 Then lblִ�п���.Caption = "��ҩҩ��"
                End If
    
                '����ִ��:ָ��ҩ;��,��ҩ�÷�,��������,�ɼ���ʽ��ִ�п���
                If Should����ִ��(lngRow, i, strTmp) Then
                    SetItemEditable , , , , , , , 1
                    Call Get����ִ�п���(cbo����ִ��, .TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), lngҩƷID, Val(.TextMatrix(i, COL_ִ������)), Val(.TextMatrix(i, COL_ִ�п���ID)), cbo��Ч.ListIndex, mint��Χ)
                Else
                    SetItemEditable , , , , , , , -1
                    If i <> -1 Then
                        If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                            If Val(.TextMatrix(i, COL_ִ������)) = 0 Then
                                cbo����ִ��.AddItem "<��ִ�ж���>"
                            ElseIf Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                cbo����ִ��.AddItem "-"
                            End If
                            Call Cbo.SetIndex(cbo����ִ��.Hwnd, 0)
                        End If
                    End If
                End If
                lbl����ִ��.Caption = strTmp
            Else
                SetItemEditable , , , 1, -1, -1, -1, -1, -1
            End If
        End If
    End With

    '����༭��־
    Call ClearItemTag

    cbsMain.RecalcLayout    '��ʱˢ��,��Lock�ɲ�Ҫ
    zlControl.FormLock 0
    Exit Sub
errH:
    zlControl.FormLock 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Should����ִ��(ByVal lngRow As Long, lngRow2 As Long, strִ�п��� As String) As Boolean
'���ܣ��ж�ָ����ҽ����(�ɼ���)�Ƿ�������ø��ӵ�ִ�п���
'������lngRow2=���ظ����е�ҽ���к�
'      strִ�п���=����ִ�п�������
    Dim i As Long
    
    lngRow2 = -1
    strִ�п��� = "����ִ��"
    With vsAdvice
        If lngRow = 0 Or .RowData(lngRow) = 0 Then Exit Function

        If RowIn�䷽��(lngRow) Then
            '��ҩ�÷�
            lngRow2 = lngRow
            strִ�п��� = "�÷�ִ��"
            Should����ִ�� = True
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
            '��ҩ;��
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
            strִ�п��� = "��ҩִ��"
            Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "F" Then
            '��������
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
            strִ�п��� = "����ִ��"
            If lngRow2 <> -1 Then Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "K" Then
            '��Ѫ;��
            If Val(.TextMatrix(lngRow, COL_��鷽��)) = 0 And gblnѪ��ϵͳ = True Then
                strִ�п��� = "�ɼ�ִ��"
            Else
                strִ�п��� = "��Ѫִ��"
            End If
            lngRow2 = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
            If lngRow2 <> -1 Then Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "E" _
            And .TextMatrix(lngRow - 1, COL_���) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
            '�ɼ���ʽ
            lngRow2 = lngRow
            strִ�п��� = "�ɼ�ִ��"
            Should����ִ�� = True
        End If
        
        '������Ժ��ִ��
        If Should����ִ�� Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_ִ������))) > 0 Then
                Should����ִ�� = False
            End If
        End If
    End With
End Function


Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(0, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
        
        If Col = col_ҽ������ Then Call vsAdvice.AutoSize(col_ҽ������)
    End If
End Sub

Private Sub vsAdvice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> col_ȱʡ And Col <> col_��ѡ)
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        End If
    End If
End Sub

Private Function RowIsLastVisible(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����һ�ɼ���
    Dim i As Long
    
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            RowIsLastVisible = lngRow = i
        End If
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If RowIsLastVisible(Row) Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            lngLeft = COL_��Ч: lngRight = COL_��Ч
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_����: lngRight = COL_�÷�
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowInһ����ҩ(Row) Then Exit Sub
            If .RowData(Row) = 0 Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(Row - 1, COL_���ID)), lngBegin, lngEnd)
            Else
                Call Getһ����ҩ��Χ(Val(.TextMatrix(Row, COL_���ID)), lngBegin, lngEnd)
            End If
            
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
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        'ִ��Executeʱ,�����ť"��ǰ"��������ʵ��ִ�У������Ƿ�ɼ�
        cbsMain.FindControl(, conMenu_Delete, True, True).Execute
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim objEdit As Object
    
    If KeyAscii = 13 Then
        '��λ����Ӧ�ı༭�ؼ�
        KeyAscii = 0
        Select Case vsAdvice.Col
            Case COL_��Ч
                Set objEdit = cbo��Ч
            Case col_ҽ������
                Set objEdit = txtҽ������
            Case COL_����
                Set objEdit = txt����
            Case COL_����
                Set objEdit = txt����
            Case COL_�÷�
                Set objEdit = txt�÷�
            Case COL_Ƶ��
                Set objEdit = txtƵ��
            Case COL_ִ��ʱ��
                Set objEdit = cboִ��ʱ��
            Case COL_ִ�п���ID
                Set objEdit = cboִ�п���
            Case COL_ҽ������
                Set objEdit = cboҽ������
        End Select
        If Not objEdit Is Nothing Then
            If objEdit.Enabled And objEdit.Visible Then objEdit.SetFocus
        End If
    End If
End Sub

Private Sub ClearItemTag()
'���ܣ�����ؼ��༭��־
    txt����.Tag = ""
    txt����.Tag = ""
    txt����.Tag = ""
    txt�÷�.Tag = ""
    txtƵ��.Tag = ""
    cboִ��ʱ��.Tag = ""
    cboҽ������.Tag = ""
    cboִ�п���.Tag = ""
    cboִ������.Tag = ""
    cbo����ִ��.Tag = ""
    txt����֤��.Tag = ""
    cbo����.Tag = ""
End Sub

Private Sub SetItemEditable(Optional int���� As Integer, Optional int���� As Integer, _
    Optional int�÷� As Integer, Optional intƵ�� As Integer, _
    Optional intִ��ʱ�� As Integer, Optional intִ�п��� As Integer, _
    Optional intִ������ As Integer, Optional int����ִ�� As Integer, _
    Optional int����֤�� As Integer, Optional int���� As Integer)
'���ܣ�����ָ���༭��Ŀ���״̬
'������0-���ֲ���,-1-��ֹ,1-����,2-����
'˵������ֹʱ,ͬʱ�������Ŀ����(����ȫ��)

    '��������Ϊ��ֹʱ,����������ı�,�Ӷ���������Validate�¼�,�����Ƚ�ֹ����˳��
    If int���� = -1 Then txt����.TabStop = False
    If int���� = -1 Then txt����.TabStop = False
    If int�÷� = -1 Then txt�÷�.TabStop = False
    If intƵ�� = -1 Then txtƵ��.TabStop = False
    If intִ��ʱ�� = -1 Then cboִ��ʱ��.TabStop = False
    If intִ�п��� = -1 Then cboִ�п���.TabStop = False
    If intִ������ = -1 Then cboִ������.TabStop = False
    If int����ִ�� = -1 Then cbo����ִ��.TabStop = False
    If int����֤�� = -1 Then txt����֤��.TabStop = False
    
    If int���� = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
        lbl������λ.Caption = "" '"��λ"
    ElseIf int���� = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If

    If int���� = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
        lbl������λ.Caption = "" '"��λ"
    ElseIf int���� = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If
    
    If int�÷� = -1 Then
        txt�÷�.Enabled = False
        txt�÷�.BackColor = Me.BackColor
        txt�÷�.Text = ""
        cmd�÷�.Enabled = False
        lbl�÷�.Caption = "�÷�"
    ElseIf int�÷� = 1 Then
        txt�÷�.TabStop = True
        txt�÷�.Enabled = True
        cmd�÷�.Enabled = True
        txt�÷�.BackColor = vsAdvice.BackColor
    End If

    If intƵ�� = -1 Then
        txtƵ��.Enabled = False
        cmdƵ��.Enabled = False
        txtƵ��.BackColor = Me.BackColor
        txtƵ��.Text = ""
    ElseIf intƵ�� = 1 Then
        txtƵ��.TabStop = True
        txtƵ��.Enabled = True
        cmdƵ��.Enabled = True
        txtƵ��.BackColor = vsAdvice.BackColor
    End If

    If intִ��ʱ�� = -1 Then
        cboִ��ʱ��.Text = ""
        cboִ��ʱ��.Enabled = False
        cboִ��ʱ��.BackColor = Me.BackColor
        cboִ��ʱ��.Clear
    ElseIf intִ��ʱ�� = 1 Then
        cboִ��ʱ��.TabStop = True
        cboִ��ʱ��.Enabled = True
        cboִ��ʱ��.BackColor = vsAdvice.BackColor
    End If

    If intִ�п��� = -1 Then
        lblִ�п���.Caption = "ִ�п���"
        cboִ�п���.Enabled = False
        cboִ�п���.BackColor = Me.BackColor
        cboִ�п���.Clear
    ElseIf intִ�п��� = 1 Then
        lblִ�п���.Caption = "ִ�п���"
        cboִ�п���.TabStop = True
        cboִ�п���.Enabled = True
        cboִ�п���.BackColor = vsAdvice.BackColor
    End If

    If intִ������ = -1 Then
        cboִ������.Enabled = False
        cboִ������.BackColor = Me.BackColor
        Call Cbo.SetIndex(cboִ������.Hwnd, -1) '�����
    ElseIf intִ������ = 1 Then
        cboִ������.TabStop = True
        cboִ������.Enabled = True
        cboִ������.BackColor = vsAdvice.BackColor
    End If
    
    If int����ִ�� = -1 Then
        lbl����ִ��.Caption = "����ִ��"
        cbo����ִ��.Enabled = False
        cbo����ִ��.BackColor = Me.BackColor
        cbo����ִ��.Clear
    ElseIf int����ִ�� = 1 Then
        lbl����ִ��.Caption = "����ִ��"
        cbo����ִ��.TabStop = True
        cbo����ִ��.Enabled = True
        cbo����ִ��.BackColor = vsAdvice.BackColor
    End If
    
    If int����֤�� = -1 Then
        lbl����֤��.Visible = False
        txt����֤��.Visible = False
        cmd����֤��.Visible = False
    ElseIf int����֤�� = 1 Then
        lbl����֤��.Visible = True
        txt����֤��.Visible = True
        cmd����֤��.Visible = True
        txt����֤��.TabStop = True
    End If
    
    If int���� = -1 Then
        cbo����.Text = ""
        lbl����.Visible = False
        cbo����.Visible = False
        lbl���ٵ�λ.Visible = False
    ElseIf int���� = 1 Then
        lbl����.Visible = True
        cbo����.Visible = True
        lbl���ٵ�λ.Visible = True
    End If
End Sub

Private Function GetPreRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ�����Ч�ɼ���
'���أ�����Ч��ʱ,����-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        
        End If
    Next
    GetPreRow = lngTmp
End Function

Private Function GetNextRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ�����Ч�ɼ���
'���أ�����Ч��ʱ,����-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetNextRow = lngTmp
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'���ܣ���ȡ��ID��ͬ��һ��ҽ���кŷ�Χ(ע�⿼��һ����ҩ�еĿ���)
    Dim lngS��ID As Long, lngO��ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, .RowData(lngRow), Val(.TextMatrix(lngRow, COL_���ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_���ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '��������
                If lngO��ID = lngS��ID Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_���ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '��������
                If lngO��ID = lngS��ID Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function GetNextID() As Long
'���ܣ�ģ���ȡ��һ��ID
    mlngNextID = mlngNextID + 1
    GetNextID = mlngNextID
End Function

Private Function GetCurRow���(lngRow As Long) As Long
'���ܣ���ȡָ���п��õĵ����
'������lngRow=Ҫȡ��ŵ���
    Dim lng��� As Long, i As Long
    Dim lng���1 As Long, lng���2 As Long
            
    'ȡ֮�����һ����Ч���,ֱ��ʹ��
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                lng��� = Val(vsAdvice.TextMatrix(i, COL_���))
                Exit For
            End If
        End If
    Next
    If lng��� = 0 Then
        '����û��,��ȡ֮ǰ��������+1
        For i = lngRow - 1 To vsAdvice.FixedRows Step -1
            If vsAdvice.RowData(i) <> 0 Then
                If IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                    lng��� = Val(vsAdvice.TextMatrix(i, COL_���))
                    Exit For
                End If
            End If
        Next
        If lng��� <> 0 Then lng��� = lng��� + 1
    End If
    If lng��� = 0 Then lng��� = 1
    GetCurRow��� = lng���
End Function

Private Sub AdviceSetҽ�����(lngRow As Long, intStep As Integer)
'���ܣ�����ǰ����ҽ����¼�����ǰ�ƻ����
'������lngRow=��ʼ������,intStep=��������,��1��-1
    Dim i As Long
    
    For i = lngRow To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                vsAdvice.TextMatrix(i, COL_���) = Val(vsAdvice.TextMatrix(i, COL_���)) + intStep
            End If
        End If
    Next
End Sub

Private Sub AdviceDelete(ByVal lngRow As Long)
'���ܣ�ָ����ҽ��ɾ������
    Dim lngBegin As Long, lngEnd As Long
    Dim lng���ID As Long, blnGroup As Boolean
    Dim lngҽ��ID As Long, i As Integer
    
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
    
    If vsAdvice.RowData(lngRow) <> 0 Then
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
            lngҽ��ID = vsAdvice.RowData(lngRow)
            lng���ID = Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
            blnGroup = RowInһ����ҩ(lngRow)
            If blnGroup Then
                '��ɾ��һ����ҩ�еĿ���(һ��Ҫɾ)
                Call Getһ����ҩ��Χ(lng���ID, lngBegin, lngEnd)
                For i = lngEnd To lngBegin Step -1 '���뷴��
                    If vsAdvice.RowData(i) = 0 Then Call DeleteRow(i)
                Next
                
                'ɾ��֮��ǰ�кſ��ܱ���
                lngRow = vsAdvice.FindRow(lngҽ��ID, lngBegin)
                
                'һ����ҩֻɾ����ǰ��
                Call DeleteRow(lngRow)
            Else
                '�����ĳ�ҩ��ɾ����ҩ;���м���ǰ��
                i = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow)
            End If
        ElseIf InStr(",D,F,K,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
            Call Delete���������Ѫ(lngRow)
            Call DeleteRow(lngRow)
        ElseIf RowIn�䷽��(lngRow) Then
            'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete��ҩ�䷽(lngRow)
            'ɾ����ǰ��(��ҩ�÷���)
            Call DeleteRow(lngRow)
        ElseIf RowIn������(lngRow) Then
            lngRow = Delete�������(lngRow)
            Call DeleteRow(lngRow)
        Else
            Call DeleteRow(lngRow)
        End If
        
        mblnNoSave = True '���Ϊδ����
    Else
        '����ֱ��ɾ��
        Call DeleteRow(lngRow)
    End If
    
    '���¶�λ��
    If vsAdvice.RowHidden(vsAdvice.Row) Then
        i = GetPreRow(vsAdvice.Row)
        If i = -1 Then i = GetNextRow(vsAdvice.Row)
        If i <> -1 Then vsAdvice.Row = i
    End If
    
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    mblnRowChange = True
    vsAdvice.Redraw = flexRDDirect
    Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
End Sub

Private Sub DeleteRow(ByVal lngRow As Long, Optional ByVal blnClear As Boolean, Optional blnDelID As Boolean = True)
'���ܣ�ɾ������е�һ��,�����ı䵱ǰ��
'������blnClear=�Ƿ�������������,��ɾ��
'      blnDelID=�Ƿ��¼Ҫɾ����ҽ��ID
    Dim lngCol As Long, blnDraw As Boolean, blnChange As Boolean
    
    With vsAdvice
        lngCol = .Col
        blnDraw = .Redraw
        blnChange = mblnRowChange
        
        mblnRowChange = False
        .Redraw = flexRDNone
        
        If .RowData(lngRow) <> 0 Then
            '�������
            Call AdviceSetҽ�����(lngRow + 1, -1)
        End If
            
        '���Ϊ��1�ҽ�ʣ��1������,����
        If Not (lngRow = .FixedRows And .Rows = .FixedRows + 1) And Not blnClear Then
            .RemoveItem lngRow
        Else
            '�����������
            .RowData(lngRow) = Empty
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "" '����
            .Cell(flexcpData, lngRow, 0, lngRow, .Cols - 1) = Empty '����
            .Cell(flexcpFontBold, lngRow, .FixedCols, lngRow, .Cols - 1) = False '����
            .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = .ForeColor '����ɫ
            If .FixedCols > 0 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .FixedCols - 1) = .ForeColorFixed '�̶�������ɫ
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .FixedCols - 1) = .BackColorFixed '�̶��б���ɫ
            End If
            Set .Cell(flexcpPicture, lngRow, 0, lngRow, .Cols - 1) = Nothing '��ԪͼƬ
            
            '��Ԫ��߿�
            .Select lngRow, .FixedCols, lngRow, COL_ִ��ʱ��
            .CellBorder vbRed, 0, 0, 0, 0, 0, 0
        End If
        
        .Col = lngCol '��Ϊ��ɾ����,���Ե��ó���϶����ж�λ,���Բ��ػָ���
        .Redraw = blnDraw
        mblnRowChange = blnChange
    End With
End Sub

Private Sub Delete���������Ѫ(ByVal lngRow As Long)
'���ܣ�1.ɾ����������Ŀ�Ĳ�λ��
'      2.ɾ��������Ŀ�ĸ��������м�������Ŀ��
'      3.ɾ����Ѫ��Ŀ����Ѫ;����
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_���ID) '��һ����,�����ò���
    If i <> -1 Then
        lngBegin = i
        For i = lngBegin To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = vsAdvice.RowData(lngRow) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        For i = lngEnd To lngBegin Step -1
            Call DeleteRow(i)
        Next
    End If
End Sub

Private Function Delete��ҩ�䷽(ByVal lngRow As Long) As Long
'���ܣ�ɾ����ҩ�䷽�����ζҩ���巨��
'������lngRow=��ҩ�䷽�÷���(�ɼ�)
'���أ�ɾ��֮�����¶�λ�ĵ�ǰ��(��ҩ�÷���)
    Dim lngBegin As Long, lngEnd As Long
    Dim lngҽ��ID As Long, i As Long
    
    lngҽ��ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '��Ϊ����ǰ��ɾ��,��Ҫ���¶�λ����ҩ�÷���
    i = vsAdvice.FindRow(lngҽ��ID)
    vsAdvice.Row = i '�������Ҳ���
    
    mblnRowChange = True
    
    Delete��ҩ�䷽ = vsAdvice.Row
End Function

Private Function Delete�������(ByVal lngRow As Long) As Long
'���ܣ�ɾ��һ���ɼ��Ķ��������Ŀ��
'������lngRow=�ɼ�������(�ɼ�)
'���أ�ɾ��֮�����¶�λ�ĵ�ǰ��(�ɼ�������)
    Dim lngBegin As Long, lngEnd As Long
    Dim lngҽ��ID As Long, i As Long
    
    lngҽ��ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '��Ϊ����ǰ��ɾ��,��Ҫ���¶�λ���ɼ�������
    i = vsAdvice.FindRow(lngҽ��ID)
    vsAdvice.Row = i '�������Ҳ���
    
    mblnRowChange = True
    
    Delete������� = vsAdvice.Row
End Function

Private Function Get��鲿λ����(ByVal lngRow As Long) As String
'���ܣ���ȡָ���еļ�鲿λ������
'������lngRow=���ҽ���Ŀɼ���
'���أ�"��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
'      ������ϵļ����Ϸ�ʽ����������ǰ�ĵ���λ��飬�򷵻ؿ��Ա����ʶ��
    Dim str��λ As String, str��λLast As String
    Dim str���� As String, i As Long
    
    With vsAdvice
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If Val(.TextMatrix(i, COL_������ĿID)) <> Val(.TextMatrix(lngRow, COL_������ĿID)) Then Exit Function '�ϵķ�ʽ
                
                If .TextMatrix(i, COL_�걾��λ) <> "" Then
                    If .TextMatrix(i, COL_�걾��λ) <> str��λLast And str��λLast <> "" Then
                        str��λ = str��λ & "|" & str��λLast & IIF(str���� <> "", ";" & Mid(str����, 2), "")
                        str���� = ""
                    End If
                    If .TextMatrix(i, COL_��鷽��) <> "" Then
                        str���� = str���� & "," & .TextMatrix(i, COL_��鷽��)
                    End If
                    
                    str��λLast = .TextMatrix(i, COL_�걾��λ)
                End If
            Else
                Exit For
            End If
        Next
        If str��λLast <> "" Then
            str��λ = str��λ & "|" & str��λLast & IIF(str���� <> "", ";" & Mid(str����, 2), "")
        End If
        Get��鲿λ���� = Mid(str��λ, 2) & vbTab & 0
    End With
End Function

Private Function Get��������IDs(ByVal lngRow As Long) As String
'���ܣ���ȡָ�������еĸ���������������ĿID��
'���أ�"����ID1,����ID2,...;����ID",���п���û�и�������������
    Dim strTmp As String, lng����ID As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_���ID)
    If i <> -1 Then
        For i = i To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = vsAdvice.RowData(lngRow) Then
                If vsAdvice.TextMatrix(i, COL_���) = "G" Then
                    lng����ID = Val(vsAdvice.TextMatrix(i, COL_������ĿID))
                Else
                    strTmp = strTmp & "," & Val(vsAdvice.TextMatrix(i, COL_������ĿID))
                End If
            Else
                Exit For
            End If
        Next
    End If
    Get��������IDs = Mid(strTmp, 2) & ";" & IIF(lng����ID = 0, "", lng����ID)
End Function

Private Function Get��ҩ�䷽IDs(ByVal lngRow As Long) As String
'���ܣ���ȡ��ҩ�䷽�����ζҩ���巨ID��
'���أ�"��ҩ���ID1,����1,��ע1;��ҩ���ID2,����2,��ע2;...|�巨ID|��ҩ��̬|����|ҩ��ID"
    Dim lng�巨ID As Long, str��ҩIDs As String, i As Long, lng��̬ As Long
    Dim lng���� As Long, lngҩ��ID As Long
    Dim strTmp As String
    
    With vsAdvice
        lng��̬ = Val(.TextMatrix(lngRow, COL_��ҩ��̬))    '�÷���
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_���) = "E" Then
                    lng�巨ID = Val(.TextMatrix(i, COL_������ĿID))
                    strTmp = .TextMatrix(i, COL_�걾��λ) '������ҩ�� ����
                ElseIf .TextMatrix(i, COL_���) = "7" Then
                    str��ҩIDs = Val(.TextMatrix(i, COL_�շ�ϸĿID)) & "," & _
                        .TextMatrix(i, COL_����) & "," & .TextMatrix(i, COL_ҽ������) & _
                        ";" & str��ҩIDs
                    If lngҩ��ID = 0 Then
                        lngҩ��ID = Val(.TextMatrix(i, COL_ִ�п���ID))
                        lng���� = Val(.TextMatrix(i, COL_����))
                    End If
                End If
            Else
                Exit For
            End If
        Next
        Get��ҩ�䷽IDs = Mid(str��ҩIDs, 1, Len(str��ҩIDs) - 1) & "|" & lng�巨ID & "|" & lng��̬ & "|" & lng���� & "|" & lngҩ��ID & "|" & strTmp
    End With
End Function

Private Function Get�������IDs(ByVal lngRow As Long) As String
'���ܣ���ȡһ���ɼ��ļ��������ĿID���걾
'���أ�"'      �������="��ĿID1,��ĿID2,...;����걾" ������°�LIS��ģʽ���ǣ�"��ĿID1|ָ��1|ָ��2...,��ĿID2|ָ��1|ָ��2...,...;����걾""
    Dim str��ĿIDs As String, str�걾 As String, i As Long
    Dim j As Long
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If Val(.TextMatrix(i, COL_�����ĿID)) = 0 And mblnNewLIS Then
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_���ID)) = .RowData(lngRow) Then
                            If Val(.TextMatrix(j, COL_�����ĿID)) = Val(.TextMatrix(i, COL_������ĿID)) And Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                                str��ĿIDs = "|" & Val(.TextMatrix(j, COL_������ĿID)) & str��ĿIDs
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    str��ĿIDs = "," & Val(.TextMatrix(i, COL_������ĿID)) & str��ĿIDs
                Else
                    If Not mblnNewLIS Then
                        str��ĿIDs = "," & Val(.TextMatrix(i, COL_������ĿID)) & str��ĿIDs
                    End If
                End If
                str�걾 = .TextMatrix(i, COL_�걾��λ)
            Else
                Exit For
            End If
        Next
    End With
    Get�������IDs = Right(str��ĿIDs, Len(str��ĿIDs) - 1) & ";" & str�걾
End Function

Private Function RowIn������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����ڼ�������е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_���) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
            '�ɼ�������
            If .TextMatrix(lngRow - 1, COL_���) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
                RowIn������ = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_���) = "C" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '������Ŀ��
            RowIn������ = True: Exit Function
        End If
    End With
End Function

Private Function RowIn�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ�������ҩ�䷽�е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_���) = "E" Then
            If Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
                '�÷���
                If Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_���) = "E" Then
                    RowIn�䷽�� = True: Exit Function
                End If
            Else
                '�巨��
                If .TextMatrix(lngRow - 1, COL_���) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    RowIn�䷽�� = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_���) = "7" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '��ҩ��
            RowIn�䷽�� = True: Exit Function
        End If
    End With
End Function

Private Function RowInһ����ҩ(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��
'������lngRow=�ɼ�����,�����ǿ���
'˵����һ����ҩ�ķ�Χ�п��ܴ��ڿ���
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lng���ID As Long, blnGroup As Boolean, i As Long
    
    lngPreRow = GetPreRow(lngRow)
    lngNextRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            If lngPreRow <> -1 And lngNextRow <> -1 Then
                If Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(lngNextRow, COL_���ID)) _
                    And Val(.TextMatrix(lngPreRow, COL_���ID)) <> 0 _
                    And InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                    And InStr(",5,6,", .TextMatrix(lngNextRow, COL_���)) > 0 Then
                    blnGroup = True
                End If
            End If
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 _
            And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            
            lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
            If lngPreRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                    And Val(.TextMatrix(lngPreRow, COL_���ID)) = lng���ID Then blnGroup = True
            End If
            If Not blnGroup And lngNextRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngNextRow, COL_���)) > 0 _
                    And Val(.TextMatrix(lngNextRow, COL_���ID)) = lng���ID Then blnGroup = True
            End If
        End If
    End With
    RowInһ����ҩ = blnGroup
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset, ByVal lngRow As Long) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ��ҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��,lngRow=��ǰ������
'���أ�����¼���Ƿ���Ч
    Dim intType As Integer
    Dim str���� As String, blnGroup As Boolean
    Dim lng�÷�ID As Long, lngGroupRow As Long
    Dim lngPreRow As Long, lngNextRow As Long
    Dim strExtData As String, strAppend As String
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim lngBegin As Long, lngEnd As Long
    Dim blnOK As Boolean
    Dim lngҩƷID As Long
    Dim t_Pati As TYPE_PatiInfoEx
    Dim bln��Ѫ As Boolean '�Ƿ�Ϊ��Ѫҽ�� ��Ѫ=0����Ѫ=1,����K���ҽ���е� ��鷽��  �ֶ�;��Ѫ-�ɼ���ʽ / ��Ѫ-��Ѫ;��
    Dim strWhere As String
    
    On Error GoTo errH
        
    lngPreRow = GetPreRow(lngRow) 'ȡ��һ��Ч��,ĳЩ����ȱʡ����һ����ͬ
    lngNextRow = GetNextRow(lngRow) 'ȡ��һ��Ч��
    
    '��Ŀ�����������뼰����Ϸ��Լ��
    '---------------------------------------------------------------------------------------------------------------
    txtҽ������.Text = rsInput!���� '��ʱ��ʾ
    
    With vsAdvice
        '������Ŀ���ɼ������ж�
        If rsInput!���ID = "C" Then
            '����������ȡһ��ȱʡ�Ĳɼ�����,ͬʱ�ж��Ƿ��вɼ���������
            lng�÷�ID = Getȱʡ�÷�ID(6, mint��Χ)
            If lng�÷�ID = 0 Then
                .Refresh
                MsgBox "û�п��õı걾�ɼ�����,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            'ȱʡ����һ����ͬ
            If lngPreRow <> -1 Then
                If RowIn������(lngPreRow) Then
                    If Val(.TextMatrix(lngPreRow, COL_�Ƿ�ͣ��)) = 0 Then lng�÷�ID = Val(.TextMatrix(lngPreRow, COL_������ĿID))
                End If
            End If
        End If
        
        '��Ѫҽ������Ѫ;���ж�
        If rsInput!���ID = "K" Then
            If gblnѪ��ϵͳ Then
                vMsg = frmMsgBox.ShowMsgBox("��ѡ����Ѫҽ�����͡�", Me, , 2)
                If vMsg = vbNo Then
                    bln��Ѫ = True
                ElseIf vMsg = vbCancel Then
                    Exit Function
                End If
            Else
                bln��Ѫ = True
            End If
            '����������ȡһ��ȱʡ����Ѫ;��
            strWhere = ""
            If bln��Ѫ = False And gblnѪ��ϵͳ = True Then
                strWhere = " And NVL(ִ�з���,0)=1 "
            End If
            lng�÷�ID = Getȱʡ�÷�ID(IIF(bln��Ѫ And gblnѪ��ϵͳ, 9, 8), mint��Χ, strWhere)
            
            If lng�÷�ID = 0 Then
                .Refresh
                 MsgBox "û�п��õ���Ѫ" & IIF(bln��Ѫ And gblnѪ��ϵͳ, "�ɼ�����", ";��") & ",���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            'ȱʡ����һ����ͬ
            If lngPreRow <> -1 Then
                If .TextMatrix(lngPreRow, COL_���) = "K" And Val(.TextMatrix(lngPreRow, COL_��鷽��)) = IIF(bln��Ѫ, "0", "1") Then
                    i = .FindRow(CStr(.RowData(lngPreRow)), lngPreRow + 1, COL_���ID)
                    If i <> -1 Then
                        If Val(.TextMatrix(i, COL_�Ƿ�ͣ��)) = 0 Then lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID))
                    End If
                End If
            End If
        End If
        
        '��ҩ�䷽����������ҩ�÷��ж�
        If InStr(",7,8,", rsInput!���ID) > 0 Then
            If rsInput!���ID = "8" Then
                If GetGroupCount(rsInput!������ĿID, mint��Χ, False) = 0 Then
                    .Refresh
                    MsgBox """" & rsInput!���� & """��һ����ҩ�䷽����û��������Ч�������ҩ��" & vbCrLf & "���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                    .Refresh: Exit Function
                End If
                
                '����ҩ��Ч����ʾ
                strMsg = GetGroupNone(rsInput!������ĿID, mint��Χ)
                If strMsg <> "" Then
                    .Refresh
                    MsgBox "�䷽""" & rsInput!���� & """������ҩƷ�ѳ�����������ƥ�䣺" & _
                        vbCrLf & vbCrLf & vbTab & strMsg & vbCrLf & vbCrLf & "��ЩҩƷ������������䷽�С�", vbInformation, gstrSysName
                    .Refresh
                End If
            End If
        
            '����������ȡһ��ȱʡ����ҩ�÷�,ͬʱ�ж��Ƿ�����ҩ�÷�����
            lng�÷�ID = Getȱʡ�÷�ID(4, mint��Χ)
            If lng�÷�ID = 0 Then
                .Refresh
                MsgBox "û�п��õ���ҩ��(��)��,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '��ҩ�÷�ȱʡ����һ����ͬ
            If RowIn�䷽��(lngPreRow) Then
                If Val(.TextMatrix(lngPreRow, COL_�Ƿ�ͣ��)) = 0 Then lng�÷�ID = Val(.TextMatrix(lngPreRow, COL_������ĿID))
            End If
        End If
        
        '������ҩ����ҩ;���ж�
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            '��ҩ;��ȱʡ����һ������ͬ���͵���ͬ
            If lngPreRow <> -1 And Not IsNull(rsInput!ҩƷ����) Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 And .TextMatrix(lngPreRow, COL_ҩƷ����) = NVL(rsInput!ҩƷ����) Then
                    i = .FindRow(CLng(.TextMatrix(lngPreRow, COL_���ID)), lngPreRow + 1)
                    If i <> -1 Then
                        If Val(.TextMatrix(i, COL_�Ƿ�ͣ��)) = 0 Then lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID))
                    End If
                End If
            End If
        End If
        
        '������ҩ��һ����ҩ���ж�
        blnGroup = RowInһ����ҩ(lngRow)
        If blnGroup Then
            If rsInput!���ID = "9" Then
                .Refresh
                MsgBox "������һ����ҩ��ҩƷ��ֱ��������׷�����", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            If .RowData(lngRow) = 0 Then
                'һ����ҩ�еĴ�������У�ֻ�в�����һ����ҩ���м�,�����Զ���Ϊһ����ҩ
                lngGroupRow = lngPreRow
            Else
                'һ����ҩ�е�ҩƷ�У������ǵ�һ�л����һ��'ȡ��ǰ�е���һ�У������ڲ�������ҽ��ʱ��ѡ������Ŀ����ʱ����ǰ�е����ݱ�ɾ�������������޷�ȡ�����е�ֵ
                If lngPreRow = -1 Then
                    lngGroupRow = vsAdvice.FindRow(.TextMatrix(lngRow, COL_���ID), lngRow + 1, COL_���ID)
                Else
                    If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                        And Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                        lngGroupRow = lngPreRow
                    Else
                        lngGroupRow = lngNextRow
                    End If
                End If
            End If
            
            'һ����ҩ��,�����Ч������ͬ
            If decode(rsInput!���ID, "5", "Y", "6", "Y", "N") <> decode(.TextMatrix(lngGroupRow, COL_���), "5", "Y", "6", "Y", "N") Then
                .Refresh
                MsgBox "����һ����ҩ��ҩƷ���붼Ϊ����ҩ���г�ҩ��", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            If zlCommFun.GetNeedName(cbo��Ч.Text) <> .TextMatrix(lngGroupRow, COL_��Ч) Then
                .Refresh
                MsgBox "����һ����ҩ��ҩƷ���붼Ϊ""" & .TextMatrix(lngGroupRow, COL_��Ч) & """��", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            i = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_���ID)), lngGroupRow + 1)
            lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID)) 'һ����ҩ�ĸ�ҩ;����ͬ
            
            '���һ����ҩ�ĵĸ�ҩ;���Ƿ��ʺ��ڵ�ǰ����ҩƷ(��һ����ҩ��ȱʡ�÷������뺯���������жϴ���)
            If Not Check�����÷�(lng�÷�ID, rsInput!������ĿID, mint��Χ) Then
                .Refresh
                MsgBox "һ���ĸ�ҩ;��Ϊ""" & .TextMatrix(i, col_ҽ������) & """���������ڵ�ǰ����ҩƷ��", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
        End If
    
        '������Ŀ
        If rsInput!���ID = "9" Then
            If GetGroupCount(rsInput!������ĿID, mint��Χ) = 0 Then
                .Refresh
                MsgBox """" & rsInput!���� & """��һ�����׷�������û��������Ч�������Ŀ��" & vbCrLf & "���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            strExtData = frmSchemeSelect.ShowMe(Me, rsInput!������ĿID, mint��Χ)
            If strExtData = "" Then .Refresh: Exit Function
        End If
    
        '��Ҫ����������ݵ�һЩ��Ŀ
        '---------------------------------------------------------------------------------------------------------------
        intType = -1
        If rsInput!���ID <> "9" Then strExtData = ""
        If rsInput!���ID = "D" Then
            '�����Ŀ����Ҫ��չ�༭�ˣ�������ǰ���е���λ��Ŀ
            intType = 0
        ElseIf rsInput!���ID = "F" Then
            '��������Ҫ����������Ŀ������ѡ�񸽼�����
            intType = 1
        ElseIf InStr(",7,8,", rsInput!���ID) > 0 Then
            '��ҩ�䷽(��ζ��ҩ���䷽����)
            intType = 2
        ElseIf rsInput!���ID = "C" Then
            '����һ���ɼ��Ķ��������Ŀ������걾
            intType = 4
            strExtData = rsInput!������ĿID & ";" & NVL(rsInput!���) '��Ŀ;�걾
        End If
        If intType <> -1 Then
            If intType = 2 Then
                lngҩƷID = Val("" & rsInput!�շ�ϸĿID)   'һ���䷽ʱΪ��
            End If
            On Error Resume Next
            If intType = 2 Then
                blnOK = frmAdviceFormula.ShowMe(Me, Nothing, txtҽ������.Hwnd, t_Pati, 3, IIF(mbyt���� <> 2, 0, 3), cbo��Ч.ListIndex, mint��Χ, _
                            , rsInput!������ĿID, strExtData, , lngҩƷID)
            Else
                blnOK = frmSchemeEditEx.ShowMe(Me, txtҽ������.Hwnd, intType, cbo��Ч.ListIndex, mint��Χ, mblnNewLIS, True, rsInput!������ĿID, strExtData)
            End If
            On Error GoTo errH
            
            If Not blnOK Then Exit Function
        End If
    
        '�޸�������Ŀʱ,��ɾ����ǰҽ��������
        '---------------------------------------------------------------------------------------------------------------
        If .RowData(lngRow) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '����ҩ���г�ҩ
                If Not blnGroup Then
                    '������ҩɾ����ҩ;����,�������ǰ��
                    i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    Call DeleteRow(i)
                    Call DeleteRow(lngRow, True)
                Else
                    'һ���ҩʱ,ֻ�����ǰ��
                    Call DeleteRow(lngRow, True)
                End If
            ElseIf InStr(",D,F,K,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '��������Ŀ��������Ŀ����Ѫҽ��
                'ɾ����λ�С�����������(��������,������Ŀ)����Ѫ;��
                Call Delete���������Ѫ(lngRow)
                '�����ǰ��
                Call DeleteRow(lngRow, True)
            ElseIf RowIn�䷽��(lngRow) Then
                '��ҩ�䷽��˳��(���)Ҫ������ϸ����
                'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
                lngRow = Delete��ҩ�䷽(lngRow)
                '�����ǰ��(��ҩ�÷���)
                Call DeleteRow(lngRow, True)
            ElseIf RowIn������(lngRow) Then
                'ɾ��������Ŀ��:ɾ��֮�����¶�λ�ĵ�ǰ��
                lngRow = Delete�������(lngRow)
                '�����ǰ��(�ɼ�������)
                Call DeleteRow(lngRow, True)
            Else
                '������Ŀֱ�������ǰ������
                Call DeleteRow(lngRow, True)
            End If
        End If
        
        '��ǰ������ҽ��
        '---------------------------------------------------------------------------------------------------------------
        If InStr(",7,8,", rsInput!���ID) > 0 Then
            '��ҩ�䷽(��ζ��ҩ���䷽����):����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet��ҩ�䷽(rsInput!������ĿID, lngRow, lng�÷�ID, strExtData)
        ElseIf rsInput!���ID = "9" Then
            '����ҽ����Ҫ�ֽ�Ϊ�����Ŀ����
            Call LoadAdvice(rsInput!������ĿID, lngRow, strExtData)
        ElseIf rsInput!���ID = "C" Then
            '�������
            lngRow = AdviceSet�������(lngRow, lng�÷�ID, strExtData)
        Else
            '�С�����ҩ�����ģ����(���)������(���)����Ѫ��������������Ŀ
            Call AdviceSet������Ŀ(rsInput, lngRow, lng�÷�ID, lngGroupRow, strExtData, bln��Ѫ)
            
            '�Զ�����һ����ҩ
            If InStr(",5,6,", rsInput!���ID) > 0 Then
                If Not RowInһ����ҩ(lngRow) Then
                    If mblnRowMerge Then
                        '�ֹ�ʹһ����ҩ
                        Call MergeRow(lngPreRow, lngRow) '����������ʾ��ǰ�е�����,������ǿ��RowChange
                    ElseIf lngPreRow <> -1 Then
                        '�Զ�ʹһ����ҩ
                        Set objControl = cbsMain.FindControl(, conMenu_Merge, , True)
                        If objControl.Checked = True Then
                            If .TextMatrix(lngPreRow, COL_���) = rsInput!���ID Then
                                If RowInһ����ҩ(lngPreRow) And RowCanMerge(lngPreRow, lngRow) And GetNextRow(lngRow) = -1 Then
                                    mblnRowMerge = True: cbsMain.RecalcLayout '*��ʱˢ��
                                    Call MergeRow(lngPreRow, lngRow, False)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If i <> lngRow Then vsAdvice.TextMatrix(i, col_ȱʡ) = vsAdvice.TextMatrix(lngRow, col_ȱʡ)
        Next
        
        '�����Զ������и�
        Call .AutoSize(col_ҽ������)
    End With
    mblnNoSave = True '���Ϊδ����
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub MergeRow(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional ByVal blnCheck As Boolean = True)
'���ܣ�����������Ϊһ����ҩ
'������lngRow1=ǰ����,���ܱ����Ѿ�����һ����ҩ
'      lngRow2=��ǰ��
'˵����������ɺ�,����Զ�λ��ԭlngRow2�ĵ�ǰ��
    Dim lngBegin As Long, lngEnd As Long
    Dim blnDo As Boolean, lngTmp As Long
    
    With vsAdvice
        If blnCheck Then
            blnDo = RowCanMerge(lngRow1, lngRow2)
        Else
            blnDo = True
        End If
        If blnDo Then
            mblnRowChange = False: .Redraw = flexRDNone
            lngTmp = .RowData(lngRow2) '��¼���ٶ�λ����ǰ��
            '��ȡ��֮ǰ��һ����ҩ
            If RowInһ����ҩ(lngRow1) Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow1, COL_���ID)), lngBegin, lngEnd)
                Call AdviceSet������ҩ(lngBegin, lngEnd)
                lngRow1 = lngBegin
                lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            End If
            Call AdviceSetһ����ҩ(lngRow1, lngRow2)
            lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            .Row = lngRow2
            mblnRowChange = True: .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub SplitRow(ByVal lngRow As Long)
'���ܣ���ָ���д�һ����ҩ�ж�������(����һ����ҩ�������ٰ�������)
'������lngRow=��ǰ��,��Ϊһ����ҩ�е����һҩƷ��
'˵����������ɺ�,����Զ�λ��ԭlngRow�ĵ�ǰ��
    Dim lngBegin As Long, lngEnd As Long, lngTmp As Long
    
    With vsAdvice
        mblnRowChange = False: .Redraw = flexRDNone
        lngTmp = .RowData(lngRow) '��¼���ڻָ���λ��ǰ��
        Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow, COL_���ID)), lngBegin, lngEnd)
        
        '��ȡ��������һ����ҩ
        Call AdviceSet������ҩ(lngBegin, lngEnd)
        
        '�����ó�����������Ϊһ����ҩ
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        lngEnd = GetPreRow(lngRow)
        Call AdviceSetһ����ҩ(lngBegin, lngEnd)
        
        '�ָ���ǰ��
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        .Row = lngRow
        mblnRowChange = True: .Redraw = flexRDDirect
    End With
End Sub

Private Function GetTableFromRecordSet() As String
'���ܣ����ݴ���ļ�¼������һ�������
    Dim strSql As String, i As Long
    Dim strValue As String, strFiled As String
    Dim blnHave As Boolean
    Dim blnHave��ѡ As Boolean
    Dim lng��� As Long
    
    For i = 0 To mrsScheme.Fields.Count - 1
        If mrsScheme.Fields(i).Name = "�Ƿ�ȱʡ" Then blnHave = True
        If mrsScheme.Fields(i).Name = "�Ƿ�ѡ" Then blnHave��ѡ = True
    Next
    
    If mrsScheme.RecordCount > 0 Then
        mrsScheme.MoveFirst
        Do While Not mrsScheme.EOF
            '���ﲻ��������¼��ҽ��
            If Not (mint��Χ = 1 And IsNull(mrsScheme!������ĿID)) Then
                lng��� = lng��� + 1
                strFiled = lng��� & IIF(strSql = "", " as ˳��", "")
                If Not blnHave Then
                    strFiled = strFiled & ",1" & IIF(strSql = "", " as �Ƿ�ȱʡ", "")
                End If
                If Not blnHave��ѡ Then
                    strFiled = strFiled & ",1" & IIF(strSql = "", " as �Ƿ�ѡ", "")
                End If
                
                For i = 0 To mrsScheme.Fields.Count - 1
                    If IsNull(mrsScheme.Fields(i).value) Then
                        Select Case mrsScheme.Fields(i).Type
                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                strValue = "-Null"
                            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                strValue = "Null+Sysdate"
                            Case Else
                                strValue = "Null"
                        End Select
                    Else
                        Select Case mrsScheme.Fields(i).Type
                            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                                strValue = "'" & Replace(Replace(mrsScheme.Fields(i).value, "[", "("), "]", ")") & "'"
                            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                                strValue = mrsScheme.Fields(i).value
                            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                                strValue = "To_Date('" & Format(mrsScheme.Fields(i).value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End Select
                    End If
                    
                    If strSql = "" Then
                        strFiled = strFiled & "," & strValue & " as " & mrsScheme.Fields(i).Name '���мӱ���
                    Else
                        strFiled = strFiled & "," & strValue
                    End If
                Next
                
                strSql = strSql & " Union ALL Select " & strFiled & " From Dual"
            End If
            mrsScheme.MoveNext
        Loop
        mrsScheme.MoveFirst
        If strSql <> "" Then
            GetTableFromRecordSet = "(" & Mid(strSql, 12) & ")"
        End If
    End If
End Function

Private Function GetTableFromAdvice(ByVal str��IDs As String, ByVal lng����ID As Long) As String
'���ܣ�����ҽ���Ĳ��˼���ID���������һ��"������Ŀ���"��
'ע�⣺����SQL�У�����ID������˳����[3],��ID����˳��Ϊ[4]
    Dim strSql As String
    
    '���ﲻ֧������¼��ҽ��
    strSql = "Select /*+ Rule*/ ��� as ˳��,1 as �Ƿ�ȱʡ,0 as �Ƿ�ѡ,ID as ���,���ID as ������,ҽ����Ч as ��Ч,A.������ĿID,A.ҽ������,A.����,A.��������,A.�ܸ�����,A.ҽ������," & _
        " A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ִ�п���ID,A.ִ��ʱ�䷽�� as ʱ�䷽��,A.ִ������,A.ִ�б��,A.�շ�ϸĿID,A.�걾��λ,A.��鷽��,A.�䷽ID,a.�����ĿID" & _
        " From ����ҽ����¼ A,Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) B" & _
        " Where Nvl(A.���ID,A.ID)=B.Column_Value And A.����ID=[3]" & IIF(mint��Χ = 1, " And A.������ĿID is Not NULL", "")
    
    GetTableFromAdvice = "(" & strSql & ")"
End Function

Private Sub LoadAdvice(ByVal lng����ID As Long, ByVal lngRow As Long, Optional ByVal str��� As String, Optional ByVal lng����ID As Long)
'���ܣ����������Ŀ(����һ����ҩ,������,��������,��ҩ�䷽)
'������lng����ID=Ϊ0ʱ��ʾ�Ӵ���ļ�¼���л��ߴӲ���ҽ����¼��ȡ
'      lngRow=�յ�������(�����ǲ��������,����λ��һ����ҩ�м�)�����Ϊ0��ʾ�����ǰ����
'      str���=Ҫ��ȡ�ĳ��׷������ݵ���ϸ��ţ�����ҽ����¼����ID
'      lng����ID=����Ϊ0ʱ����ʾͨ��"str���"Ϊ��ID����ȡ����ҽ����¼
    Dim rsItems As New ADODB.Recordset
    Dim rs��� As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    
    Dim lngCurRow As Long, intCount As Integer, lng��� As Long
    Dim bln��ҩ;�� As Boolean, bln�ɼ����� As Boolean, bln��Ѫ;�� As Boolean
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim bln��ҩ�÷� As Boolean, bln��ҩ�巨 As Boolean, bln�䷽ As Boolean
    Dim lng���� As Long, vBookMark As Variant, strҩ��IDs As String
    Dim lng���ID As Long, strSQL��� As String, str��¼ As String
    Dim intƵ������ As Integer, str���÷�Χ As String, strƵ�� As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Me.Refresh
    
    '��ҽ��ȱʡ������
    If msng���� = 0 Then msng���� = 1

    If lng����ID <> 0 Then
        '���ﲻ֧������¼��ҽ��������ѡ������������
        str��¼ = "(Select 1 as ˳��,1 as �Ƿ�ȱʡ,0 as �Ƿ�ѡ,A.* From ������Ŀ��� A Where A.�������ID=[1])"
        
        If str��� <> "" Then
            If Left(str���, 1) = "+" Then
                strSQL��� = " And Instr([2],','||A.���||',')>0"
            ElseIf Left(str���, 1) = "-" Then
                strSQL��� = " And Instr([2],','||A.���||',')=0"
            End If
        End If
    ElseIf lng����ID <> 0 Then
        str��¼ = GetTableFromAdvice(str���, lng����ID)
    Else
        str��¼ = GetTableFromRecordSet
        If str��¼ = "" Then
            '���ﲻ֧������¼��ҽ��
            str��¼ = "(Select 1 as ˳��,1 as �Ƿ�ȱʡ,0 as �Ƿ�ѡ,A.* From ������Ŀ��� A Where A.�������ID=[1]" & IIF(mint��Χ = 1, " And A.������ĿID is Not NULL", "") & ")"
        End If
    End If
    
    'ҩƷ�����Ϣ:��Ȼ�����շ�ϸĿID,����������û��,��ǰ������Ҳû��
    strSql = "Select A.���,B.ҩ��ID,B.ҩƷID,B.����ϵ��,B." & IIF(mint��Χ = 1, "����", "סԺ") & "�ɷ���� As �ɷ����,C.����,Nvl(D.����,C.����) as ����,C.���,C.����," & _
        decode(mint��Χ, 1, "B.�����װ as ��װϵ��,B.���ﵥλ as ��װ��λ", 2, "B.סԺ��װ as ��װϵ��,B.סԺ��λ as ��װ��λ", "C.���㵥λ as ��װ��λ,1 as ��װϵ��") & _
        " From " & str��¼ & " A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
        " Where A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID" & strSQL��� & _
        " And C.ID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=[5]" & _
        " Order by A.˳��,A.���,C.����"
    Set rs��� = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",", lng����ID, str���, IIF(gbytҩƷ������ʾ = 0, 1, 3))
    
    '������Ϣ
    strSql = "Select A.���,B.����ID,B.��������,C.����,C.���㵥λ" & _
        " From ������Ŀ��� A,�������� B,�շ���ĿĿ¼ C" & _
        " Where A.�շ�ϸĿID=B.����ID And B.����ID=C.ID" & _
        " And A.�������ID=[1]" & strSQL��� & _
        " Order by A.���,C.����"
    Set rs���� = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",")
    
    '��������к�Ӧ����ҽ���༭ʱ�Ĵ���һ��
    strSql = "Select A.�Ƿ�ȱʡ,A.�Ƿ�ѡ,A.��Ч,A.���,A.������,A.������ĿID,A.�շ�ϸĿID,A.ҽ������,A.����,A.�ܸ�����,A.��������," & _
        " A.ҽ������,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ִ�п���ID,B.���,B.����,B.���㵥λ,Decode(B.���,'D',A.�걾��λ," & _
        " Nvl(A.�걾��λ,B.�걾��λ)) as �걾��λ,A.��鷽��,A.ʱ�䷽��,Nvl(A.ִ������,B.ִ�п���) as ִ������," & _
        " Nvl(A.ִ�б��,0) as ִ�б��,B.��������,B.���㷽ʽ,B.ִ��Ƶ��,C.�������,C.������,C.ҩƷ����,C.Ʒ��ҽ��,A.�䷽ID," & _
        " c.�ٴ��Թ�ҩ,a.�����ĿID,d.���� As ����֤��,b.����ʱ��,b.ִ�з��� " & _
        " From " & str��¼ & " A,������ĿĿ¼ B,ҩƷ���� C,��������Ŀ¼ D" & _
        " Where Nvl(A.������ĿID,0)=B.ID(+) And Nvl(A.������ĿID,0)=C.ҩ��ID(+) And Nvl(a.�����ĿID,0)=d.ID(+) " & strSQL��� & _
        " Order by A.˳��,A.���"
    Set rsItems = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",", lng����ID, str���)
    With vsAdvice
        mblnRowChange = False
        .Redraw = flexRDNone
        If lngRow = 0 And lng����ID = 0 Then
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
            lngRow = .FixedRows
        End If
        If lng����ID <> 0 Then
            '�ҵ�������
            For i = 1 To .Rows - 1
                If .TextMatrix(i, col_ҽ������) = "" Then
                    If i = .Rows - 1 Then
                        lngRow = i
                    Else
                        .RemoveItem i
                        .Rows = .Rows + 1
                        lngRow = .Rows - 1
                    End If
                    Exit For
                End If
            Next
            If lngRow = 0 Then
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
        End If
        intCount = 0 '�Ѿ����õ�����
        lng��� = GetCurRow���(lngRow) '��ʼ���
        
        For i = 1 To rsItems.RecordCount
            lngCurRow = lngRow + intCount
            If lngCurRow > lngRow Then .AddItem "", lngCurRow
             
            '��¼���ID
            .RowData(lngCurRow) = -1 * rsItems!���
            If Not IsNull(rsItems!������) Then
                .TextMatrix(lngCurRow, COL_���ID) = -1 * rsItems!������
            End If
            
            .TextMatrix(lngCurRow, col_ȱʡ) = IIF(rsItems!�Ƿ�ȱʡ = 1, -1, 0)
            .TextMatrix(lngCurRow, COL_���) = lng��� + intCount
            .TextMatrix(lngCurRow, COL_��Ч) = IIF(NVL(rsItems!��Ч, 0) = 0, "����", "����")
            .TextMatrix(lngCurRow, COL_���) = NVL(rsItems!���, "*") '����¼��ҽ��������
            
            .TextMatrix(lngCurRow, COL_������ĿID) = NVL(rsItems!������ĿID)
            .TextMatrix(lngCurRow, COL_����) = NVL(rsItems!����)
            .TextMatrix(lngCurRow, COL_�걾��λ) = NVL(rsItems!�걾��λ)
            .TextMatrix(lngCurRow, COL_��鷽��) = NVL(rsItems!��鷽��)

            '����
            .TextMatrix(lngCurRow, COL_���㷽ʽ) = NVL(rsItems!���㷽ʽ, 0)
            .TextMatrix(lngCurRow, COL_��������) = NVL(rsItems!��������)
            .TextMatrix(lngCurRow, COL_�������) = NVL(rsItems!�������)
            .TextMatrix(lngCurRow, COL_ҩƷ����) = NVL(rsItems!ҩƷ����)
            .TextMatrix(lngCurRow, col_��ѡ) = IIF(rsItems!�Ƿ�ѡ = 1, -1, 0)
            .TextMatrix(lngCurRow, COL_�䷽ID) = NVL(rsItems!�䷽ID)
            .TextMatrix(lngCurRow, COL_�ٴ��Թ�ҩ) = rsItems!�ٴ��Թ�ҩ & ""
            .TextMatrix(lngCurRow, COL_�����ĿID) = "" & rsItems!�����ĿID
            If .TextMatrix(lngCurRow, COL_�����ĿID) <> "" Then
                If .TextMatrix(lngCurRow, COL_���) = "E" And .TextMatrix(lngCurRow, COL_��������) = "4" Then
                    .TextMatrix(lngCurRow, COL_����֤��) = rsItems!����֤�� & ""
                End If
            End If
            .TextMatrix(lngCurRow, COL_�����ȼ�) = Val("" & rsItems!������)
            .TextMatrix(lngCurRow, COL_ִ�б��) = Val("" & rsItems!ִ�б��)
            
            If Format(NVL(rsItems!����ʱ��, "3000/1/1"), "yyyy-MM-dd") <> "3000-01-01" Then
                .TextMatrix(lngCurRow, COL_�Ƿ�ͣ��) = 1
            End If
            .TextMatrix(lngCurRow, COL_ִ�з���) = rsItems!ִ�з��� & ""
            
            'ҩƷ�����Ϣ:�в�ҩ�϶���,��ҩ�������������λ�Զ�ƥ��
            lng���� = 0: vBookMark = 0
            '�ٴ�·������ͳ������������
            If NVL(rsItems!���) = "7" Or (InStr(",5,6,", NVL(rsItems!���, "*")) > 0 _
                And (NVL(rsItems!��Ч, 0) = 1 Or gblnҩƷ�������ҽ�� And NVL(rsItems!Ʒ��ҽ��, 0) = 0)) Then
                If Not IsNull(rsItems!�շ�ϸĿID) Then
                    rs���.Filter = "ҩƷID=" & rsItems!�շ�ϸĿID
                Else
                    rs���.Filter = "ҩ��ID=0"
                End If
                If Not rs���.EOF Then
                    If IsNull(rsItems!�շ�ϸĿID) Then
                        'ȡ����ϵ��Ϊ��������С����������һ�����
                        If CInt(NVL(rsItems!��������, 0)) <> 0 Then
                            Do While Not rs���.EOF
                                If rs���!����ϵ�� / rsItems!�������� = Int(rs���!����ϵ�� / rsItems!��������) Then
                                    If rs���!����ϵ�� / rsItems!�������� < lng���� Or lng���� = 0 Then
                                        vBookMark = rs���.Bookmark
                                        lng���� = rs���!����ϵ�� / rsItems!��������
                                    End If
                                End If
                                rs���.MoveNext
                            Loop
                            If vBookMark <> 0 Then rs���.Bookmark = vBookMark
                        End If
                        If rs���.EOF Then rs���.MoveFirst
                    End If
                    .TextMatrix(lngCurRow, COL_����) = NVL(rs���!����)
                    .TextMatrix(lngCurRow, COL_�շ�ϸĿID) = rs���!ҩƷID
                    .TextMatrix(lngCurRow, COL_����ϵ��) = NVL(rs���!����ϵ��)
                    .TextMatrix(lngCurRow, COL_��װϵ��) = NVL(rs���!��װϵ��)
                    .TextMatrix(lngCurRow, COL_��װ��λ) = NVL(rs���!��װ��λ)
                    .TextMatrix(lngCurRow, COL_�ɷ����) = NVL(rs���!�ɷ����, 0)
                End If
            ElseIf NVL(rsItems!���) = "4" Then
                rs����.Filter = "����ID=" & NVL(rsItems!�շ�ϸĿID, 0)
                If Not rs����.EOF Then
                    .TextMatrix(lngCurRow, COL_����) = NVL(rs����!����)
                    .TextMatrix(lngCurRow, COL_��װ��λ) = NVL(rs����!���㵥λ) 'ɢװ��λ
                    .TextMatrix(lngCurRow, COL_��������) = NVL(rs����!��������, 0)
                End If
                .TextMatrix(lngCurRow, COL_����ϵ��) = 1
                .TextMatrix(lngCurRow, COL_��װϵ��) = 1
                .TextMatrix(lngCurRow, COL_�շ�ϸĿID) = NVL(rsItems!�շ�ϸĿID, 0)
            End If
                                
            '�ж��Ƿ��ض���
            bln��ҩ;�� = False: bln�ɼ����� = False: bln��Ѫ;�� = False
            bln��ҩ�÷� = False: bln��ҩ�巨 = False: bln�䷽ = False
            If rsItems!��� = "E" Then
                If IsNull(rsItems!������) Then
                    If Val(.TextMatrix(lngCurRow - 1, COL_���ID)) = .RowData(lngCurRow) Then
                        If InStr(",5,6,", .TextMatrix(lngCurRow - 1, COL_���)) > 0 Then
                            bln��ҩ;�� = True
                        ElseIf .TextMatrix(lngCurRow - 1, COL_���) = "C" Then
                            bln�ɼ����� = True
                        Else
                            bln��ҩ�÷� = True
                        End If
                    End If
                ElseIf .TextMatrix(lngCurRow - 1, COL_���) = "K" And .RowData(lngCurRow - 1) = Val(.TextMatrix(lngCurRow, COL_���ID)) Then
                    bln��Ѫ;�� = True
                Else
                    bln��ҩ�巨 = True
                End If
            End If
            If rsItems!��� = "7" Or bln��ҩ�巨 Or bln��ҩ�÷� Then bln�䷽ = True
            
            'Ƶ������
            If bln�ɼ����� Then
                '�ɼ������Լ�����Ŀ��Ϊ׼
                j = .FindRow(CStr(.RowData(lngCurRow)), , COL_���ID)
                intƵ������ = .TextMatrix(j, COL_Ƶ������)
            Else
                intƵ������ = NVL(rsItems!ִ��Ƶ��, 0)
            End If
            If bln�䷽ Then
                str���÷�Χ = 2 '��ҩ�䷽(�����巨,�÷�)����ҽ
            ElseIf intƵ������ = 1 Then
                str���÷�Χ = -1 'һ����
            ElseIf intƵ������ = 2 Then
                str���÷�Χ = -2 '������
            ElseIf intƵ������ = 0 Then '��ѡƵ��
                If NVL(rsItems!��Ч, 0) = 1 Then
                    str���÷�Χ = "1,-1" '��������Ϊһ����(�����Ʋ���Ψһ����)
                Else
                    str���÷�Χ = 1
                End If
            End If
            If rsItems!ִ��Ƶ�� & "" = "��Ҫʱ" Then
                str���÷�Χ = -3
            ElseIf rsItems!ִ��Ƶ�� & "" = "��Ҫʱ" Then
                str���÷�Χ = -5
            End If
            
            'Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ
            .TextMatrix(lngCurRow, COL_Ƶ������) = intƵ������
            If Not IsNull(rsItems!ִ��Ƶ��) Then
                If CheckƵ�ʿ���(NVL(rsItems!������ĿID, 0), Val(str���÷�Χ), NVL(rsItems!ִ��Ƶ��)) Then 'Val(str���÷�Χ)
                    If GetƵ����Ϣ_����(rsItems!ִ��Ƶ��, 0, 0, "", str���÷�Χ) Then
                        .TextMatrix(lngCurRow, COL_Ƶ��) = rsItems!ִ��Ƶ��
                        .TextMatrix(lngCurRow, COL_Ƶ�ʴ���) = NVL(rsItems!Ƶ�ʴ���, 0)
                        .TextMatrix(lngCurRow, COL_Ƶ�ʼ��) = NVL(rsItems!Ƶ�ʼ��, 0)
                        .TextMatrix(lngCurRow, COL_�����λ) = NVL(rsItems!�����λ)
                        
                        '������ѡƵ�ʿ�������Ϊ��һ����
                        If NVL(rsItems!��Ч, 0) = 1 And intƵ������ = 0 And NVL(rsItems!Ƶ�ʴ���, 0) = 0 And NVL(rsItems!Ƶ�ʼ��, 0) = 0 Then
                            .TextMatrix(lngCurRow, COL_Ƶ������) = 1
                        End If
                    End If
                End If
            End If
            If .TextMatrix(lngCurRow, COL_Ƶ��) = "" And Not IsNull(rsItems!������ĿID) Then 'ȡȱʡ��
                If NVL(rsItems!��Ч, 0) = 1 And intƵ������ = 0 Then
                    If mblnһ���� Then '����ȱʡΪһ����
                        str���÷�Χ = -1
                        .TextMatrix(lngCurRow, COL_Ƶ������) = 1
                    Else
                        str���÷�Χ = 1
                    End If
                End If
                Call GetȱʡƵ��(NVL(rsItems!������ĿID, 0), str���÷�Χ, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                .TextMatrix(lngCurRow, COL_Ƶ��) = strƵ��
                .TextMatrix(lngCurRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngCurRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngCurRow, COL_�����λ) = str�����λ
            End If
            
            '����
            .TextMatrix(lngCurRow, COL_����) = NVL(rsItems!����)
            If InStr(",5,6,", NVL(rsItems!���, "*")) > 0 And NVL(rsItems!����, 0) > 0 Then
                msng���� = rsItems!���� '�������Ϊȱʡ
            End If
            
            '����
            .TextMatrix(lngCurRow, COL_����) = FormatEx(NVL(rsItems!��������), 5)
            If NVL(rsItems!���) = "4" Then
                .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow, COL_��װ��λ) 'ɢװ��λ
            ElseIf bln��ҩ�÷� Then
                .TextMatrix(lngCurRow, COL_������λ) = ""
            ElseIf NVL(rsItems!��Ч, 0) = 0 Then
                If InStr(",5,6,7,", NVL(rsItems!���, "*")) > 0 Or InStr(",1,2,", NVL(rsItems!���㷽ʽ, 0)) > 0 Then
                    .TextMatrix(lngCurRow, COL_������λ) = NVL(rsItems!���㵥λ)
                End If
            Else
                If InStr(",5,6,7,", NVL(rsItems!���, "*")) > 0 Or (intƵ������ = 0 And InStr(",1,2,", NVL(rsItems!���㷽ʽ, 0)) > 0) Then
                    .TextMatrix(lngCurRow, COL_������λ) = NVL(rsItems!���㵥λ)
                End If
            End If
            
            '����
            If InStr(",5,6,", NVL(rsItems!���, "*")) > 0 Then
                '��ҩ����������,�����۵�λ���,��װ��λ��ʾ
                If Not IsNull(rsItems!�ܸ�����) And Val(.TextMatrix(lngCurRow, COL_��װϵ��)) <> 0 Then
                    .TextMatrix(lngCurRow, COL_����) = FormatEx(rsItems!�ܸ����� / Val(.TextMatrix(lngCurRow, COL_��װϵ��)), 5)
                End If
                If NVL(rsItems!��Ч, 0) = 1 Then
                    .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow, COL_��װ��λ)
                End If
            Else
                '�����������ҩ����������
                If Not IsNull(rsItems!�ܸ�����) Then
                    .TextMatrix(lngCurRow, COL_����) = rsItems!�ܸ�����
                End If
                If bln�䷽ Then
                    .TextMatrix(lngCurRow, COL_������λ) = "��" '��ҩ�䷽������λΪ"��"
                ElseIf NVL(rsItems!��Ч, 0) = 1 Then
                    If NVL(rsItems!���) = "4" Then
                        .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow, COL_��װ��λ) 'ɢװ��λ
                    Else
                        .TextMatrix(lngCurRow, COL_������λ) = NVL(rsItems!���㵥λ)
                    End If
                End If
            End If
            
            'ִ��ʱ��
            If .TextMatrix(lngCurRow, COL_Ƶ��) <> "" And Val(.TextMatrix(lngCurRow, COL_Ƶ������)) <> 1 Then
                If Not IsNull(rsItems!ʱ�䷽��) Then
                    If ExeTimeValid(rsItems!ʱ�䷽��, Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                        Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), .TextMatrix(lngCurRow, COL_�����λ)) Then
                        .TextMatrix(lngCurRow, COL_ִ��ʱ��) = rsItems!ʱ�䷽��
                    End If
                End If
            End If
            
            '�÷�����ʾ
            If bln�ɼ����� Then
                .TextMatrix(lngCurRow, COL_�÷�) = rsItems!����
            ElseIf bln��ҩ;�� Or bln��ҩ�÷� Then
                '��ҩ����ҩ�䷽���÷�,ִ��ʱ��
                If bln��ҩ�÷� Then
                    .TextMatrix(lngCurRow, COL_�÷�) = rsItems!����
                End If
                For j = lngCurRow - 1 To lngRow Step -1
                    If Val(.TextMatrix(j, COL_���ID)) = .RowData(lngCurRow) Then
                        If bln��ҩ;�� Then
                            .TextMatrix(j, COL_�÷�) = rsItems!���� & rsItems!ҽ������  '����
                        End If
                        .TextMatrix(j, COL_ִ��ʱ��) = .TextMatrix(lngCurRow, COL_ִ��ʱ��)
                    Else
                        Exit For
                    End If
                Next
            ElseIf bln��Ѫ;�� Then
                .TextMatrix(lngCurRow - 1, COL_�÷�) = rsItems!����
            End If
                                
            'ִ������
            If InStr(",5,6,7,", NVL(rsItems!���, "*")) > 0 Then
                If NVL(rsItems!ִ������, 0) = 5 Then
                    .TextMatrix(lngCurRow, COL_ִ������) = 5
                Else
                    .TextMatrix(lngCurRow, COL_ִ������) = 4
                End If
            ElseIf NVL(rsItems!���) = "4" Then
                .TextMatrix(lngCurRow, COL_ִ������) = 4
            Else
                .TextMatrix(lngCurRow, COL_ִ������) = NVL(rsItems!ִ������, 0)
            End If
            
            'ִ�п���ID:Ϊ0-����,5-Ժ��ִ��ʱȡ��Ϊ0
            If InStr(",0,5,", Val(.TextMatrix(lngCurRow, COL_ִ������))) = 0 And NVL(rsItems!ִ�п���ID, 0) <> 0 Then
                If InStr(",5,6,7,", NVL(rsItems!���, "*")) > 0 Then
                    strҩ��IDs = Get����ҩ��IDs(rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 0, mint��Χ)
                    If InStr("," & strҩ��IDs & ",", "," & rsItems!ִ�п���ID & ",") > 0 Then
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = NVL(rsItems!ִ�п���ID, 0)
                    End If
                ElseIf NVL(rsItems!���) = "4" Then
                    strҩ��IDs = Get���÷��ϲ���IDs(Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 0, mint��Χ, rsItems!������ĿID)
                    If InStr("," & strҩ��IDs & ",", "," & rsItems!ִ�п���ID & ",") > 0 Then
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = NVL(rsItems!ִ�п���ID, 0)
                    End If
                Else
                    .TextMatrix(lngCurRow, COL_ִ�п���ID) = NVL(rsItems!ִ�п���ID, 0)
                End If
            End If
                        
            'ҽ������
            .TextMatrix(lngCurRow, COL_ҽ������) = NVL(rsItems!ҽ������)
            
            '----------------------
            '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
            If InStr(",5,6,", .TextMatrix(lngCurRow, COL_���)) > 0 And .TextMatrix(lngCurRow, COL_�������) <> "" Then
                If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(lngCurRow, COL_�������)) > 0 Then
                    .Cell(flexcpFontBold, lngCurRow, col_ҽ������) = True
                End If
            End If
            
            '����һЩ������
            If (InStr(",F,G,D,7,E,C,", NVL(rsItems!���, "*")) > 0 And Not IsNull(rsItems!������)) Or bln��ҩ;�� Then
                .RowHidden(lngCurRow) = True
            End If
            
            'ҽ������
            If Not .RowHidden(lngCurRow) Then
                If IsNull(rsItems!������ĿID) Then
                    .TextMatrix(lngCurRow, col_ҽ������) = rsItems!ҽ������ '����¼��ҽ��
                ElseIf InStr(",F,D,", NVL(rsItems!���, "*")) > 0 And IsNull(rsItems!������) Then
                    .TextMatrix(lngCurRow, col_ҽ������) = rsItems!���� '��ʱ
                Else
                    .TextMatrix(lngCurRow, col_ҽ������) = AdviceTextMake(lngCurRow)
                End If
            Else
                .TextMatrix(lngCurRow, col_ҽ������) = rsItems!����
            End If
            
            '----------------------
            intCount = intCount + 1
            rsItems.MoveNext
        Next
        
        '--------------------------------------------------
        If intCount > 0 Then
            '��ȡ����������ҽ������
            For i = lngRow To lngCurRow
                If InStr(",F,D,", .TextMatrix(i, COL_���)) > 0 And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    .TextMatrix(i, col_ҽ������) = AdviceTextMake(i)
                End If
            Next
            
            '������Ӱ���е����
            Call AdviceSetҽ�����(lngCurRow + 1, intCount)
            
            '������ʵ��ҽ��ID
            For i = lngRow To lngCurRow
                lng���ID = .RowData(i)
                .RowData(i) = GetNextID
                For j = i - 1 To lngRow Step -1
                    If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                        .TextMatrix(j, COL_���ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
                For j = i + 1 To lngCurRow
                    If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                        .TextMatrix(j, COL_���ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
            Next
        End If
        
        '--------------------------------------------------
        If .RowHidden(lngRow) Then 'Ѱ�ҿɼ���(���䷽�ͼ���֮��)
            For i = lngRow + 1 To .Rows - 1
                If Not .RowHidden(i) And .RowData(i) <> 0 Then
                    lngRow = i: Exit For
                End If
            Next
        End If
        
        '�̶���ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
        Call .AutoSize(col_ҽ������)
        .Row = lngRow: .Col = col_ҽ������
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        mblnRowChange = True
    End With
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function AdviceSet��ҩ�䷽(lng������ĿID As Long, ByVal lngRow As Long, ByVal lng�÷�ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset, Optional ByVal lng�䷽ID As Long) As Long
'���ܣ�(����)������ҩ�䷽��ȱʡҽ������
'������lng������ĿID=�������ҩ�䷽ID��ζ��ҩID
'      lngRow=��ǰ������
'      lng�÷�ID=ȱʡ��ҩ�÷�ID
'      strExtData=�����䷽���ζҩ���巨����:���ID1,����,��ע;���ID2,����,��ע...|��ҩ�巨|��ҩ��̬|����|ҩ��ID|����"
'      rsCurr=������޸����䷽���ݺ����,�����Ҫ���ֵ�һЩ��ǰֵ
'���أ���������ҩ�䷽�ĵ�ǰ��ʾ�к�
    Dim rsItems As New ADODB.Recordset '��ҩ��ϸ��Ϣ
    Dim rsUse As New ADODB.Recordset '��ҩ�÷���Ϣ
    Dim rs�巨 As New ADODB.Recordset '��ҩ�巨��Ŀ��Ϣ
    Dim rs�÷� As New ADODB.Recordset '��ҩ�÷���Ŀ��Ϣ
    Dim arr��ҩs As Variant, str��ҩIDs As String, lng���ID As Long
    Dim lngCopyRow As Long 'ȱʡ������
    Dim lngDrugRow As Long '���ȱʡ����������ҩ�䷽,��Ϊ���䷽�ĵ�һ����ҩ��
    Dim lngFirstRow As Long '��ǰ�䷽�ĵ�һ����ҩ��
    Dim strSql As String, i As Long
    
    Dim strƵ�� As String, intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim lng�巨ID As Long, int�Ƴ� As Integer
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim lng��̬ As Long
    Dim str���� As String
        
    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngDrugRow = -1
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    If lngCopyRow <> -1 Then
        If RowIn�䷽��(lngCopyRow) Then
            '�����һ��Ч������ҩ�䷽��,��ȡ���ĵ�һ��ҩ��
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngCopyRow)), , COL_���ID)
        End If
    End If
    
    '��ȡ������ݿ���Ϣ
    '------------------
    arr��ҩs = Split(Split(strExtData, "|")(0), ";")
    For i = 0 To UBound(arr��ҩs)
        str��ҩIDs = str��ҩIDs & "," & CStr(Split(arr��ҩs(i), ",")(0))
    Next
    str��ҩIDs = Mid(str��ҩIDs, 2)
    lng�巨ID = Val(Split(strExtData, "|")(1))
    lng��̬ = Val(Split(strExtData, "|")(2))
    str���� = Split(strExtData, "|")(5)
    
    '�䷽�÷���Ϣ:ֱ�������䷽ʱ���п�����,���뵥ζ��ҩ��
    strSql = "Select A.�÷�ID,A.Ƶ��,A.�Ƴ�,A.ҽ������" & _
        " From �����÷����� A,������ĿĿ¼ B" & _
        " Where A.�÷�ID=B.ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & _
        " And Nvl(A.����,0)=0 And A.��ĿID=[1] And (b.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or b.����ʱ�� is NULL)"
    Set rsUse = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng������ĿID, mint��Χ)
    If Not rsUse.EOF Then lng�÷�ID = rsUse!�÷�ID 'ȱʡ���õ���ҩ�䷽�÷�����
    
    '�䷽���ζ��ҩ��Ϣ:��ҩ�޹�����,��Ӧ�ĵĹ���¼һ������ֻ��һ��
    strSql = "Select A.�������,A.վ��,A.���,A.����ID,A.ID,A.����,A.����,A.�걾��λ,A.���㵥λ,A.���㷽ʽ,A.ִ��Ƶ��," & _
        "A.�����Ա�,A.����Ӧ��,A.�����Ŀ,A.��������,A.ִ�а���,A.ִ�п���,A.�������,A.�Ƽ�����,A.�ο�Ŀ¼ID,A.��ԱID,A.����ʱ��,A.����ʱ��,A.¼������,A.�Թܱ���,A.ִ�з���,A.ִ�б��," & _
        "B.ҩƷID,B.����ϵ��,B." & IIF(mint��Χ = 1, "����", "סԺ") & "�ɷ���� As �ɷ����," & _
        decode(mint��Χ, 1, "B.�����װ as ��װϵ��,B.���ﵥλ as ��װ��λ", 2, "B.סԺ��װ as ��װϵ��,B.סԺ��λ as ��װ��λ", "C.���㵥λ as ��װ��λ,1 as ��װϵ��") & _
        " From ������ĿĿ¼ A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
        " Where A.ID=B.ҩ��ID And B.ҩƷID=C.ID And B.ҩƷID IN(Select Column_Value From Table(f_Num2list([1])))"
    Set rsItems = zldatabase.OpenSQLRecord(strSql, Me.Caption, str��ҩIDs) 'In
        
    '�䷽�巨��Ŀ��Ϣ
    Set rs�巨 = Get������Ŀ��¼(lng�巨ID)
    
    '�䷽�÷���Ŀ��Ϣ
    Set rs�÷� = Get������Ŀ��¼(lng�÷�ID)
    
    
    '�����䷽���ζ��ҩ��:�����û�����˳��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    mblnRowChange = False
    
    '��ҩ�÷���ҽ��ID,ID˳������Ų�һ��һ��
    If Not rsCurr Is Nothing Then
        '�޸����䷽�е�����,�÷��б��Ϊ�޸�,ҽ��ID����
        lng���ID = rsCurr!ҽ��ID
    Else
        '���������ҩ�䷽
        lng���ID = GetNextID
    End If
    
    For i = 0 To UBound(arr��ҩs)
        rsItems.Filter = "ҩƷID=" & CStr(Split(arr��ҩs(i), ",")(0)) 'Ӧ�ÿ϶���
        
        vsAdvice.AddItem "", lngRow
        
        vsAdvice.RowHidden(lngRow) = True
        vsAdvice.RowData(lngRow) = GetNextID
        vsAdvice.TextMatrix(lngRow, COL_���ID) = lng���ID '��Ӧ���������ҩ�÷���
        vsAdvice.TextMatrix(lngRow, COL_��Ч) = zlCommFun.GetNeedName(cbo��Ч.Text)
        vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1) '�������
        
        vsAdvice.TextMatrix(lngRow, COL_���) = rsItems!���
        vsAdvice.TextMatrix(lngRow, col_ҽ������) = rsItems!����
        vsAdvice.TextMatrix(lngRow, COL_������ĿID) = rsItems!ID
        vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rsItems!���㷽ʽ, 0)
        vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = NVL(rsItems!ִ��Ƶ��, 0)
        vsAdvice.TextMatrix(lngRow, COL_��������) = NVL(rsItems!��������)
        
        vsAdvice.TextMatrix(lngRow, COL_����) = FormatEx(Val(Split(arr��ҩs(i), ",")(1)), 5) '��ζҩ�ĵ�������
        vsAdvice.TextMatrix(lngRow, COL_������λ) = NVL(rsItems!���㵥λ)
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = CStr(Split(arr��ҩs(i), ",")(2)) '��ζҩ�Ľ�ע
        
        '�����Ϣ:��ҩ�����ڹ�����,һ����
        vsAdvice.TextMatrix(lngRow, COL_�շ�ϸĿID) = rsItems!ҩƷID
        vsAdvice.TextMatrix(lngRow, COL_����ϵ��) = rsItems!����ϵ��
        vsAdvice.TextMatrix(lngRow, COL_��װ��λ) = rsItems!��װ��λ
        vsAdvice.TextMatrix(lngRow, COL_��װϵ��) = rsItems!��װϵ��
        vsAdvice.TextMatrix(lngRow, COL_�ɷ����) = NVL(rsItems!�ɷ����, 0) '����ҩʵ��������
        
        If lngFirstRow <> 0 Then
            '����һ�������õ������ҩ��ͬ
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ������)
            vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ�п���ID)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
            vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
            vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
            vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
        ElseIf Not rsCurr Is Nothing Then
            '�޸����䷽���ݺ���������,�����뵱ǰ��ֵ
            
            'ִ������:�޸�ʱ���ݵ�ǰ�������þ���
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = decode(NVL(rsCurr!ִ������), "�Ա�ҩ", 5, 4)
            'ִ�п���
            vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = NVL(rsCurr!ִ�п���ID)
            
            vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = NVL(rsCurr!Ƶ��)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = NVL(rsCurr!Ƶ�ʴ���)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = NVL(rsCurr!Ƶ�ʼ��)
            vsAdvice.TextMatrix(lngRow, COL_�����λ) = NVL(rsCurr!�����λ)
            vsAdvice.TextMatrix(lngRow, COL_����) = NVL(rsCurr!����)
            vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = NVL(rsCurr!ִ��ʱ��)
        Else
            'ִ������:��ҩ�䷽�����ҩ��ͬ,ȱʡ=4-ָ������
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = 4
                        
            'ִ�п���(�����䷽����ѡ��)
            vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Val(Split(strExtData, "|")(4))
                        
            'ִ��Ƶ��
            '�����÷��������õ�����
            If Not rsUse.EOF Then
                If Not IsNull(rsUse!Ƶ��) Then
                    Call GetƵ����Ϣ_����(rsUse!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                    vsAdvice.TextMatrix(lngRow, COL_�����λ) = str�����λ
                End If
            End If
            '��ȱʡ����һ����ͬ
            If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = "" And lngDrugRow <> -1 Then
                If vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��) <> "" Then
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ�ʴ���)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ�ʼ��)
                    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngDrugRow, COL_�����λ)
                End If
            End If
            '��ȡȱʡֵ
            If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = "" Then
                Call GetȱʡƵ��(NVL(rsItems!ID, 0), 2, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                vsAdvice.TextMatrix(lngRow, COL_�����λ) = str�����λ
            End If
            
            '����(����):��������Ҫ,��ɢװ��̬��ȷ������
            If Val(Split(strExtData, "|")(3)) > 1 Or lng��̬ <> 0 Then
                vsAdvice.TextMatrix(lngRow, COL_����) = Val(Split(strExtData, "|")(3))
            Else
                If vsAdvice.TextMatrix(lngRow, COL_��Ч) = "����" And vsAdvice.TextMatrix(lngRow, COL_Ƶ��) <> "" Then
                    int�Ƴ� = 1
                    If Not rsUse.EOF Then int�Ƴ� = NVL(rsUse!�Ƴ�, 1)
                    '�䷽����
                    vsAdvice.TextMatrix(lngRow, COL_����) = CalcȱʡҩƷ����(1, int�Ƴ�, _
                            Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���)), _
                            Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��)), _
                            vsAdvice.TextMatrix(lngRow, COL_�����λ))
                End If
            End If
            
            'ִ��ʱ��
            If lngDrugRow <> -1 Then 'ȱʡ����һ����ͬ
                If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��) Then
                    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngDrugRow, COL_ִ��ʱ��)
                End If
            End If
            If vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then 'ȱʡʱ�䷽��
                vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(2, vsAdvice.TextMatrix(lngRow, COL_Ƶ��), lng�÷�ID)
            End If
        End If
        
        '---------------------------------------
        If lngFirstRow = 0 Then lngFirstRow = lngRow '����ҩ�䷽�ĵ�һ�������ҩ��
        lngRow = lngRow + 1 '���ֵ�ǰ������λ��
    Next
    
    '������ҩ�䷽�巨��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.AddItem "", lngRow
    vsAdvice.RowHidden(lngRow) = True
    vsAdvice.RowData(lngRow) = GetNextID
    vsAdvice.TextMatrix(lngRow, COL_���ID) = lng���ID
    vsAdvice.TextMatrix(lngRow, COL_��Ч) = vsAdvice.TextMatrix(lngFirstRow, COL_��Ч)
    vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
    Call AdviceSetҽ�����(lngRow + 1, 1) '�������
    vsAdvice.TextMatrix(lngRow, COL_���) = rs�巨!���
    vsAdvice.TextMatrix(lngRow, COL_������ĿID) = lng�巨ID
    vsAdvice.TextMatrix(lngRow, COL_�걾��λ) = str����
    vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rs�巨!���㷽ʽ, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������) = NVL(rs�巨!��������)
    
    '!��ҩ�巨��Ҳ�����ҩ�ĸ���
    vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
    
    vsAdvice.TextMatrix(lngRow, col_ҽ������) = rs�巨!����
    
    vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ������) '��ҩƷ��Ϊ׼
    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
    
    'ִ������:ȱʡ������Ŀ����(������ΪԺ��ִ��),�޸�ʱ���ݵ�ǰ��������
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = NVL(rs�巨!ִ�п���, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = decode(NVL(rsCurr!ִ������), "��Ժ��ҩ", 5, NVL(rs�巨!ִ�п���, 0))
    End If
    
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_ִ������))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(rs�巨!���, lng�巨ID, 0, NVL(rs�巨!ִ�п���, 0), cbo��Ч.ListIndex, mint��Χ)
    End If
    
    '���ֵ�ǰ������λ��
    lngRow = lngRow + 1
    
    '������ҩ�䷽�÷���:��ҩ�䷽����ʾ��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.RowData(lngRow) = lng���ID
    
    If Get������Ŀ��¼(lng������ĿID)!��� & "" = "8" Then
        vsAdvice.TextMatrix(lngRow, COL_�䷽ID) = lng������ĿID
    End If
    If lng�䷽ID <> 0 Then
        vsAdvice.TextMatrix(lngRow, COL_�䷽ID) = lng�䷽ID
    End If
    vsAdvice.TextMatrix(lngRow, COL_��Ч) = vsAdvice.TextMatrix(lngFirstRow, COL_��Ч)
    vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
    Call AdviceSetҽ�����(lngRow + 1, 1) '�������
    vsAdvice.TextMatrix(lngRow, COL_���) = rs�÷�!���
    vsAdvice.TextMatrix(lngRow, COL_������ĿID) = lng�÷�ID
    vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rs�÷�!���㷽ʽ, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������) = NVL(rs�÷�!��������)
    
    '!��ҩ�÷���Ҳ�����ҩ�ĸ���
    vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
    vsAdvice.TextMatrix(lngRow, COL_������λ) = "��"
    
    vsAdvice.TextMatrix(lngRow, COL_����) = rs�÷�!����
    vsAdvice.TextMatrix(lngRow, COL_�÷�) = rs�÷�!����
    vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ������)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
    
    'ִ������:ȱʡ������Ŀ����(������ΪԺ��ִ��),�޸�ʱ���ݵ�ǰ��������
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = NVL(rs�÷�!ִ�п���, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = decode(NVL(rsCurr!ִ������), "��Ժ��ҩ", 5, NVL(rs�÷�!ִ�п���, 0))
    End If
    
    '��ҩ�÷����δ����ִ�п���,��ȱʡΪ�������ڿ���
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_ִ������))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(rs�÷�!���, lng�÷�ID, 0, NVL(rs�÷�!ִ�п���, 0), cbo��Ч.ListIndex, mint��Χ)
    End If
    
    If Not rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = NVL(rsCurr!ҽ������)
    ElseIf Not rsUse.EOF Then
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = NVL(rsUse!ҽ������)
    End If
    
    '��ҩ��̬(����AdviceTextMake��)
    vsAdvice.TextMatrix(lngRow, COL_��ҩ��̬) = lng��̬
    
    '��ҩ�䷽ҽ������
    vsAdvice.TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
    
    '-------------------
    vsAdvice.Row = lngRow
    mblnRowChange = True
        
    AdviceSet��ҩ�䷽ = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet�������(ByVal lngRow As Long, ByVal lng�ɼ�����ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset) As Long
'���ܣ����������ļ���(���)
'������rsItems=�����ѡ�񷵻صļ�¼��
'      lngRow=��ǰ������
'      lng�ɼ�����ID=ȱʡ�Ĳɼ�����
'      strExtData=���:"'      �������="��ĿID1,��ĿID2,...;����걾" ������°�LIS��ģʽ���ǣ�"��ĿID1|ָ��1|ָ��2...,��ĿID2|ָ��1|ָ��2...,...;����걾"
'      rsCurr=�޸ļ�����Ŀʱ��
'���أ�����֮��ĵ�ǰ��ʾ�к�
    Dim rsMore As New ADODB.Recordset '�ɼ�������Ϣ
    Dim rsItems As New ADODB.Recordset '������Ŀ��Ϣ
    Dim arrItems As Variant, strItems As String
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim strƵ�� As String, intƵ�ʴ��� As Integer
    Dim intƵ�ʼ�� As Integer, str�����λ As String
    Dim lng���ID As Long, strҽ������ As String
    Dim lngCopyRow As Long, lngFirstRow As Long
    Dim strSql As String, i As Long
    Dim rsLIS As New ADODB.Recordset
    Dim strTmp As String
    Dim Y As Long
    Dim blnLis As Boolean
    
    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    
    '������Ŀ��Ϣ
    '----------------------------------------------------------------------------
    '����������Ŀ��Ϣ:������˳��
    arrItems = Split(Split(strExtData, ";")(0), ",")
    For i = UBound(arrItems) To 0 Step -1
        If mblnNewLIS Then
            strTmp = arrItems(i)
            If InStr(strTmp, "|") > 0 Then
                For Y = 0 To UBound(Split(strTmp, "|"))
                    strItems = strItems & "," & Val(Split(strTmp, "|")(Y))
                    If Y > 0 Then
                        strSql = strSql & " Union All " & " Select '" & Val(Split(strTmp, "|")(Y)) & "' as ����,'" & Val(Split(strTmp, "|")(0)) & "' as ���� From Dual "
                    End If
                Next
            Else
                strItems = strItems & "," & Val(strTmp)
            End If
        Else
            strItems = strItems & "," & Val(arrItems(i))
        End If
    Next
    Set rsItems = Get������Ŀ��¼(0, Mid(strItems, 2))
    If strSql <> "" Then
        Set rsLIS = zldatabase.OpenSQLRecord(Mid(strSql, 11), Me.Caption)
        blnLis = True
    End If
    
    'ȡĳ��������Ŀ�Ĳɼ�����
    strSql = "Select A.��ĿID,Nvl(A.����,0) as ���,A.�÷�ID" & _
        " From �����÷����� A,������ĿĿ¼ B" & _
        " Where A.�÷�ID=B.ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & _
        " And A.��ĿID IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
        " And (b.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or b.����ʱ�� is NULL)" & _
        " Order by A.��ĿID,Nvl(A.����,0)"
    Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, Mid(strItems, 2), mint��Χ)
    If Not rsMore.EOF Then
        If rsCurr Is Nothing Or lng�ɼ�����ID = 0 Then
            lng�ɼ�����ID = rsMore!�÷�ID '�޸�ʱ����
        End If
    End If
    Set rsMore = Get������Ŀ��¼(lng�ɼ�����ID)
    
    mblnRowChange = False
    
    '���ø��м�����Ŀ
    '----------------------------------------------------------------------------
    '�ɼ�����ҽ��ID,ID˳������Ų�һ��һ��
    If Not rsCurr Is Nothing Then
        '�޸��˼�������е�����,�ɼ������б��Ϊ�޸�,ҽ��ID����
        lng���ID = rsCurr!ҽ��ID
    Else
        '���������ҩ�䷽
        lng���ID = GetNextID
    End If
    
    With vsAdvice
        For i = 1 To rsItems.RecordCount
            .AddItem "", lngRow
            
            .RowHidden(lngRow) = True
            .RowData(lngRow) = GetNextID
            .TextMatrix(lngRow, COL_���ID) = lng���ID '��Ӧ���ɼ�������
            .TextMatrix(lngRow, COL_��Ч) = zlCommFun.GetNeedName(cbo��Ч.Text)
            
            .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
            Call AdviceSetҽ�����(lngRow + 1, 1) '�������
            
            .TextMatrix(lngRow, COL_���) = rsItems!���
            .TextMatrix(lngRow, col_ҽ������) = rsItems!����
            .TextMatrix(lngRow, COL_������ĿID) = rsItems!ID
            .TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rsItems!���㷽ʽ, 0)
            If .TextMatrix(lngRow, COL_��Ч) = "����" And NVL(rsItems!ִ��Ƶ��, 0) = 0 And mblnһ���� Then
                .TextMatrix(lngRow, COL_Ƶ������) = 1 '��ѡ��Ƶ�ʵ�����ȱʡΪһ����
            Else
                .TextMatrix(lngRow, COL_Ƶ������) = NVL(rsItems!ִ��Ƶ��, 0)
            End If
            .TextMatrix(lngRow, COL_��������) = NVL(rsItems!��������)
            .TextMatrix(lngRow, COL_ִ������) = NVL(rsItems!ִ�п���, 0)
            '����걾
            .TextMatrix(lngRow, COL_�걾��λ) = Split(strExtData, ";")(1)
            If mblnNewLIS And rsItems!ID & "" <> "" And blnLis Then
                rsLIS.Filter = "����=" & rsItems!ID
                If rsLIS.EOF = False Then
                    .TextMatrix(lngRow, COL_�����ĿID) = rsLIS!���� & ""
                End If
            End If
            
            '��������һ���ɼ��ļ�����Ŀ��ͬ
            If lngFirstRow <> 0 Then
                .TextMatrix(lngRow, COL_����) = .TextMatrix(lngFirstRow, COL_����)
                
                'һ���ɼ��ļ�����ĿӦ����ͬ
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = .TextMatrix(lngFirstRow, COL_ִ�п���ID)
                End If
                .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngFirstRow, COL_Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngFirstRow, COL_�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngFirstRow, COL_ִ��ʱ��)
            ElseIf Not rsCurr Is Nothing Then
                If cbo��Ч.ListIndex = 1 Then
                    .TextMatrix(lngRow, COL_����) = NVL(rsCurr!����, 1)
                End If
                
                'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    If NVL(rsCurr!ִ�п���ID, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = rsCurr!ִ�п���ID
                    Else
                        .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(rsItems!���, rsItems!ID, 0, NVL(rsItems!ִ�п���, 0), cbo��Ч.ListIndex, mint��Χ)
                    End If
                End If
                
                'ִ��Ƶ��
                .TextMatrix(lngRow, COL_Ƶ��) = NVL(rsCurr!Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = NVL(rsCurr!Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = NVL(rsCurr!Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = NVL(rsCurr!�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = NVL(rsCurr!ִ��ʱ��)
            Else
                If cbo��Ч.ListIndex = 1 Then
                    .TextMatrix(lngRow, COL_����) = 1
                End If
                
                'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    '֮ǰҪ�����������ID
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(rsItems!���, rsItems!ID, 0, NVL(rsItems!ִ�п���, 0), cbo��Ч.ListIndex, mint��Χ)
                End If
                
                'ִ��Ƶ��
                Call GetȱʡƵ��(NVL(rsItems!ID, 0), GetƵ�ʷ�Χ(lngRow), strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngRow, COL_�����λ) = str�����λ
            
                'ִ��ʱ��:"��ѡƵ��"(ҩƷ�ǿ�ѡƵ��,����������Ϊһ����)
                If Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Then
                    If lngCopyRow <> -1 Then '����һ����ͬ
                        If .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��) Then
                            .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngCopyRow, COL_ִ��ʱ��)
                        End If
                    End If
                    If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then  'ȱʡʱ�䷽��
                        .TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(1, .TextMatrix(lngRow, COL_Ƶ��))
                    End If
                End If
            End If
            
            strҽ������ = strҽ������ & "," & rsItems!���� 'ҽ������
            If lngFirstRow = 0 Then lngFirstRow = lngRow '��һ��Ŀ��
            lngRow = lngRow + 1 '���ֵ�ǰ������λ��
            
            rsItems.MoveNext
        Next
        
        '���ñ걾�Ĳɼ�����
        '----------------------------------------------------------------------------
        rsItems.MoveFirst
        .RowData(lngRow) = lng���ID
        
        .TextMatrix(lngRow, COL_��Ч) = zlCommFun.GetNeedName(cbo��Ч.Text)
        
        .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1) '�������
        
        .TextMatrix(lngRow, COL_���) = rsMore!���
        .TextMatrix(lngRow, COL_����) = rsMore!����
        .TextMatrix(lngRow, COL_�÷�) = rsMore!����
        .TextMatrix(lngRow, COL_������ĿID) = rsMore!ID
        .TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rsMore!���㷽ʽ, 0)
        .TextMatrix(lngRow, COL_��������) = NVL(rsMore!��������)
        .TextMatrix(lngRow, COL_�걾��λ) = .TextMatrix(lngFirstRow, COL_�걾��λ)
        
        '����Ϊ������Ŀ��,�������Ŀ��ͬ
        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngFirstRow, COL_����)
        If cbo��Ч.ListIndex = 1 Then
            .TextMatrix(lngRow, COL_������λ) = NVL(rsMore!���㵥λ)
        End If
        
        'ִ��Ƶ��
        .TextMatrix(lngRow, COL_Ƶ������) = .TextMatrix(lngFirstRow, COL_Ƶ������) '�Լ����Ϊ׼
        .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngFirstRow, COL_Ƶ��)
        .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
        .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
        .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngFirstRow, COL_�����λ)
        .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngFirstRow, COL_ִ��ʱ��)
        .TextMatrix(lngRow, COL_ִ������) = NVL(rsMore!ִ�п���, 0)
        
        'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
        If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
            .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(rsMore!���, rsMore!ID, 0, NVL(rsMore!ִ�п���, 0), cbo��Ч.ListIndex, mint��Χ)
        End If
        
        If Not rsCurr Is Nothing Then
            .TextMatrix(lngRow, COL_ҽ������) = NVL(rsCurr!ҽ������)
        End If
        
        'ҽ������:����1,����2(�걾 �ɼ�����)
        .TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
        
        .Row = lngRow
    End With
    mblnRowChange = True
    AdviceSet������� = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceSet������Ŀ(rsInput As ADODB.Recordset, ByVal lngRow As Long, ByVal lng��ҩ;��ID As Long, ByVal lngGroupRow As Long, ByVal strExtData As String, Optional ByVal bln��Ѫ As Boolean = True)
'���ܣ���������(����)���С�����ҩ�����(���)������(���)�����ģ���Ѫ��������������Ŀ��ȱʡҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��
'      lngRow=��ǰ������
'      lng��ҩ;��ID=ȱʡ��ҩ;��ID,��һ����ҩʱ�ĸ�ҩ;��ID
'      lngGroupRow=��һ����ҩ��һ���ҩ�в����µĳ�ҩ��ʱ,��Ӧһ����ҩ��һ���к�
'      strExtData=���:������鲿λ��������Ϣ,����:���������������������Ϣ,�����޸�������
'      bln��Ѫ ��ǰ����Ѫҽ��Ϊ��Ѫҽ�����������ΪK��������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim rsMore As New ADODB.Recordset '������Ŀ��ϸ��Ϣ
    Dim strSql As String, lngCopyRow As Long
    Dim lngTmp As Long, i As Long
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim strҩ��IDs As String, sng���� As Single
    Dim strƵ�� As String, intƵ�ʴ��� As Integer
    Dim intƵ�ʼ�� As Integer, str�����λ As String
    Dim lng�շ���ĿID As Long, blnƷ�� As Boolean
        
    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
            
    With vsAdvice
        '��ʼ����ҽ��ȱʡ����
        .RowData(lngRow) = GetNextID
        .TextMatrix(lngRow, COL_��Ч) = zlCommFun.GetNeedName(cbo��Ч.Text)
        
        '���:��������,��ǰ��ռ������ź�,�������������
        .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1)
        
        .TextMatrix(lngRow, COL_���) = rsInput!���ID
        .TextMatrix(lngRow, COL_����) = rsInput!���� '�����ƿ����Ǳ���
        .TextMatrix(lngRow, COL_������ĿID) = rsInput!������ĿID
        
        'ҩƷ����
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            strSql = "Select �������,ҩƷ����,Ʒ��ҽ��,�ٴ��Թ�ҩ,������ From ҩƷ���� Where ҩ��ID=[1]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(rsInput!������ĿID))
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COL_�������) = NVL(rsTmp!�������)
                .TextMatrix(lngRow, COL_ҩƷ����) = NVL(rsTmp!ҩƷ����)
                .TextMatrix(lngRow, COL_�ٴ��Թ�ҩ) = rsTmp!�ٴ��Թ�ҩ & ""
                .TextMatrix(lngRow, COL_�����ȼ�) = Val("" & rsTmp!������)
                If chkMedicineVariety.value = 1 Then
                    blnƷ�� = True
                Else
                    '�Ƿ���ҩƷ�̶���Ʒ���´�
                    blnƷ�� = NVL(rsTmp!Ʒ��ҽ��, 0) <> 0 And cbo��Ч.ListIndex = 0
                End If
            End If
        End If
        
        If NVL(rsInput!���ID) = "4" And mbyt���� = 1 Then
            If chkMedicineVariety.value = 1 Then
                blnƷ�� = True
            Else
                blnƷ�� = False
            End If
        End If
        
        '�Ƿ���ҩƷ�̶���Ʒ���´�
        lng�շ���ĿID = NVL(rsInput!�շ�ϸĿID, 0)
        If blnƷ�� Then lng�շ���ĿID = 0
        
        'ҩƷ�����ĵĹ����Ϣ
        .TextMatrix(lngRow, COL_�շ�ϸĿID) = lng�շ���ĿID
        If lng�շ���ĿID <> 0 Then
            If InStr(",5,6,", rsInput!���ID) > 0 Then
                strSql = "Select Nvl(C.����,A.����) as ����,B.����ϵ��,B." & IIF(mint��Χ = 1, "����", "סԺ") & "�ɷ���� As �ɷ����," & _
                    decode(mint��Χ, 1, "B.�����װ as ��װϵ��,B.���ﵥλ as ��װ��λ", 2, "B.סԺ��װ as ��װϵ��,B.סԺ��λ as ��װ��λ", "A.���㵥λ as ��װ��λ,1 as ��װϵ��") & _
                    " From �շ���ĿĿ¼ A,ҩƷ��� B,�շ���Ŀ���� C" & _
                    " Where A.ID=B.ҩƷID And A.ID=[1]" & _
                    " And A.ID=C.�շ�ϸĿID(+) And C.����(+)=1 And C.����(+)=[2]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng�շ���ĿID, IIF(gbytҩƷ������ʾ = 0, 1, 3))
                .TextMatrix(lngRow, COL_����) = rsTmp!���� '������������ʽ�������
                .TextMatrix(lngRow, COL_����ϵ��) = rsTmp!����ϵ��
                .TextMatrix(lngRow, COL_��װ��λ) = rsTmp!��װ��λ
                .TextMatrix(lngRow, COL_��װϵ��) = rsTmp!��װϵ��
                .TextMatrix(lngRow, COL_�ɷ����) = NVL(rsTmp!�ɷ����, 0)
            ElseIf rsInput!���ID = "4" Then
                strSql = "Select A.��������,B.����,B.���㵥λ From �������� A,�շ���ĿĿ¼ B Where A.����ID=B.ID And A.����ID=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng�շ���ĿID)
                .TextMatrix(lngRow, COL_����) = rsTmp!���� '������������ʽ�������
                .TextMatrix(lngRow, COL_����ϵ��) = 1
                .TextMatrix(lngRow, COL_��װϵ��) = 1
                .TextMatrix(lngRow, COL_��װ��λ) = NVL(rsTmp!���㵥λ) 'ɢװ��λ
                .TextMatrix(lngRow, COL_��������) = NVL(rsTmp!��������, 0)
            End If
        End If
        
        '��ȡ����������Ŀ��Ϣ
        '----------------------------------------------------------------------------
        If InStr(",5,6,", rsInput!���ID) > 0 And lng�շ���ĿID <> 0 Then
            strSql = "Select a.�÷�id,a.Ƶ��,a.���˼���,a.С������,a.ҽ������,a.�Ƴ�,c.ҩ��id as ��ĿID " & _
                " From ҩƷ�÷����� A,������ĿĿ¼ B,ҩƷ��� C " & _
                " Where A.�÷�ID=B.ID and a.ҩƷID=c.ҩƷid And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & _
                " And A.ҩƷID=[3] And A.����=1"
            strSql = "Select A.*,1 as ����,B.�÷�ID," & _
                " B.Ƶ��,B.���˼���,B.С������,B.ҽ������,B.�Ƴ�" & _
                " From ������ĿĿ¼ A,(" & strSql & ") B" & _
                " Where A.ID=b.��Ŀid(+) And A.ID=[1]"
        Else
            strSql = "Select A.*" & _
                " From �����÷����� A,������ĿĿ¼ B" & _
                " Where A.�÷�ID=B.ID And (Nvl(A.����,0)=0 Or " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & ")" & _
                " And A.��ĿID=[1]"
            strSql = "Select A.*,Nvl(B.����,0) as ����,B.�÷�ID," & _
                " B.Ƶ��,B.���˼���,B.С������,B.ҽ������,B.�Ƴ�" & _
                " From ������ĿĿ¼ A,(" & strSql & ") B" & _
                " Where A.ID=B.��ĿID(+) And A.ID=[1]" & _
                " Order by ����"
        End If
        
        Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(rsInput!������ĿID), mint��Χ, lng�շ���ĿID)
        
        If lng�շ���ĿID = 0 Then '������������ʽ��������
            .TextMatrix(lngRow, COL_����) = rsMore!����
        End If
        
        '������λ
        If rsInput!���ID = "4" Then
            .TextMatrix(lngRow, COL_������λ) = .TextMatrix(lngRow, COL_��װ��λ) 'ɢװ��λ
        Else
            If cbo��Ч.ListIndex = 0 Then
                If InStr(",5,6,", rsInput!���ID) > 0 Or InStr(",1,2,", NVL(rsMore!���㷽ʽ, 0)) > 0 Then
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsMore!���㵥λ) 'ҩƷΪ������λ
                End If
            Else
                If InStr(",5,6,", rsInput!���ID) > 0 Or (NVL(rsMore!ִ��Ƶ��, 0) = 0 And InStr(",1,2,", NVL(rsMore!���㷽ʽ, 0)) > 0) Then
                    .TextMatrix(lngRow, COL_������λ) = NVL(rsMore!���㵥λ) 'ҩƷΪ������λ
                End If
            End If
        End If
        
        If cbo��Ч.ListIndex = 1 Then
            If InStr(",5,6,", rsInput!���ID) > 0 Then
                '�С�����ҩ������������λ���ǰ�װ��λ
                .TextMatrix(lngRow, COL_������λ) = .TextMatrix(lngRow, COL_��װ��λ)
            ElseIf rsInput!���ID = "4" Then
                .TextMatrix(lngRow, COL_������λ) = .TextMatrix(lngRow, COL_��װ��λ) 'ɢװ��λ
            Else
                '��������Ҫ��������
                '���Ϊһ���Ի�ƴ�����ȱʡ����Ϊ1
                If NVL(rsMore!ִ��Ƶ��, 0) = 1 Or NVL(rsMore!���㷽ʽ, 0) = 3 Then
                    .TextMatrix(lngRow, COL_����) = 1
                End If
                .TextMatrix(lngRow, COL_������λ) = NVL(rsMore!���㵥λ)
            End If
        End If
        
        .TextMatrix(lngRow, COL_���㷽ʽ) = NVL(rsMore!���㷽ʽ, 0)
        If .TextMatrix(lngRow, COL_��Ч) = "����" And NVL(rsMore!ִ��Ƶ��, 0) = 0 And mblnһ���� Then
            .TextMatrix(lngRow, COL_Ƶ������) = 1 '��ѡ��Ƶ�ʵ�����ȱʡΪһ����
        Else
            .TextMatrix(lngRow, COL_Ƶ������) = NVL(rsMore!ִ��Ƶ��, 0)
        End If
        .TextMatrix(lngRow, COL_��������) = NVL(rsMore!��������)
        
        '�걾��λ
        If InStr(",4,5,6,", rsInput!���ID) > 0 Then
            .TextMatrix(lngRow, COL_�걾��λ) = rsInput!���� '��¼ҩƷ����������ʱѡ������
        ElseIf rsInput!���ID <> "D" Then
            .TextMatrix(lngRow, COL_�걾��λ) = NVL(rsMore!�걾��λ)
        End If
        
        'ִ������:������Ŀʱ������Ŀ����,ҩƷ������=4-ָ������,һ����ҩ����ͬ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            If lngGroupRow <> 0 Then
                .TextMatrix(lngRow, COL_ִ������) = .TextMatrix(lngGroupRow, COL_ִ������)
            Else
                .TextMatrix(lngRow, COL_ִ������) = 4
            End If
        ElseIf rsInput!���ID = "4" Then
            .TextMatrix(lngRow, COL_ִ������) = 4
        Else
            .TextMatrix(lngRow, COL_ִ������) = NVL(rsMore!ִ�п���, 0)
        End If
        
        'ִ�п���:ҩƷȱʡ����һ����ͬ,һ����ҩ����ͬ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            If lngGroupRow <> 0 Then
                strҩ��IDs = Get����ҩ��IDs(rsInput!���ID, rsInput!������ĿID, lng�շ���ĿID, 0, mint��Χ)
                If InStr("," & strҩ��IDs & ",", "," & .TextMatrix(lngGroupRow, COL_ִ�п���ID) & ",") > 0 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = .TextMatrix(lngGroupRow, COL_ִ�п���ID)
                End If
            ElseIf lngCopyRow <> -1 Then
                If rsInput!���ID = .TextMatrix(lngCopyRow, COL_���) Then
                    strҩ��IDs = Get����ҩ��IDs(rsInput!���ID, rsInput!������ĿID, lng�շ���ĿID, 0, mint��Χ)
                    If InStr("," & strҩ��IDs & ",", "," & .TextMatrix(lngCopyRow, COL_ִ�п���ID) & ",") > 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = .TextMatrix(lngCopyRow, COL_ִ�п���ID)
                    End If
                End If
            End If
        End If
        If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
            If rsInput!���ID = "Z" And (NVL(rsMore!��������, 0) = 3 Or NVL(rsMore!��������, 0) = 2 Or NVL(rsMore!��������, 0) = 1) Then
                'ת��,��Ժ������ҽ����ȱʡִ�п���Ϊ��
            ElseIf rsInput!���ID = "Z" And NVL(rsMore!��������, 0) = 7 Then
                '����ҽ��
            ElseIf InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                'ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                '��Ҫ�����������ID
                .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(rsInput!���ID, rsInput!������ĿID, lng�շ���ĿID, NVL(rsMore!ִ�п���, 0), cbo��Ч.ListIndex, mint��Χ)
            End If
        End If
        
        'ִ��Ƶ��:��ѡƵ��,һ���Ի������
        If True Then 'If Nvl(rsMore!ִ��Ƶ��, 0) = 0 Then
            'ȱʡ����һ��������ͬ
            If lngCopyRow <> -1 Then
                If .TextMatrix(lngRow, COL_��Ч) = .TextMatrix(lngCopyRow, COL_��Ч) And GetƵ�ʷ�Χ(lngRow) = GetƵ�ʷ�Χ(lngCopyRow) Then
                    If .TextMatrix(lngCopyRow, COL_Ƶ��) <> "" _
                        And Not (.TextMatrix(lngRow, COL_���) = "7" And Not RowIn�䷽��(lngCopyRow)) _
                        And Not (.TextMatrix(lngRow, COL_���) <> "7" And RowIn�䷽��(lngCopyRow)) _
                        And CheckƵ�ʿ���(NVL(rsInput!������ĿID, 0), GetƵ�ʷ�Χ(lngRow), .TextMatrix(lngCopyRow, COL_Ƶ��)) Then
                        .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��)
                        .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngCopyRow, COL_Ƶ�ʴ���)
                        .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngCopyRow, COL_Ƶ�ʼ��)
                        .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngCopyRow, COL_�����λ)
                    End If
                End If
            End If
            '��ȡȱʡƵ��
            If .TextMatrix(lngRow, COL_Ƶ��) = "" Then
                Call GetȱʡƵ��(NVL(rsInput!������ĿID, 0), GetƵ�ʷ�Χ(lngRow), strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngRow, COL_�����λ) = str�����λ
            End If
        End If
        
        '�У�����ҩ��һЩȱʡ��Ϣ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            'ִ��Ƶ��
            If lngGroupRow <> 0 Then
                'һ����ҩ����ͬ
                .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngGroupRow, COL_Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngGroupRow, COL_Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngGroupRow, COL_Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngGroupRow, COL_�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngGroupRow, COL_ִ��ʱ��)
                'Ƶ������ҲҪ��ͬ,����ǿ������Ϊһ����
                .TextMatrix(lngRow, COL_Ƶ������) = .TextMatrix(lngGroupRow, COL_Ƶ������)
            End If
            
            'ȷ��������ҩ������
            '1.����Ϊһ��Ƶ����������
            '2-���Ƴ���Ϊ�Ƴ�����(Ӧ����һ��Ƶ����������)
            If cbo��Ч.ListIndex = 1 Then
                sng���� = msng����

                If .TextMatrix(lngRow, COL_�����λ) = "��" Then
                    If 7 > sng���� Then sng���� = 7
                ElseIf .TextMatrix(lngRow, COL_�����λ) = "��" Then
                    If Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) > sng���� Then
                        sng���� = Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��))
                    End If
                ElseIf .TextMatrix(lngRow, COL_�����λ) = "Сʱ" Then
                    If Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) \ 24 > sng���� Then
                        sng���� = Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) \ 24
                    End If
                ElseIf .TextMatrix(lngRow, COL_�����λ) = "����" Then
                    If sng���� = 0 Then sng���� = 1
                End If
                If sng���� = 0 Then sng���� = 1
            End If

            rsMore.Filter = "����>0" 'ȡ��һ�ָ�ҩ;����Ϊȱʡ����
            If Not rsMore.EOF Then
                '����һ����ҩʱ,���õ�ȱʡ�÷�Ƶ������
                If lngGroupRow = 0 Then
                    If Not IsNull(rsMore!�÷�ID) Then lng��ҩ;��ID = rsMore!�÷�ID
                    If Not IsNull(rsMore!Ƶ��) And Val(.TextMatrix(lngRow, COL_Ƶ������)) <> 1 Then 'ȱʡΪһ��������
                        Call GetƵ����Ϣ_����(rsMore!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                        .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                        .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                        .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                        .TextMatrix(lngRow, COL_�����λ) = str�����λ
                    End If
                End If
                
                'ҽ������
                .TextMatrix(lngRow, COL_ҽ������) = NVL(rsMore!ҽ������) 'һ��Ϊ��ҩ;����˵��
                
                'ҩƷ����
                If NVL(rsMore!���˼���, 0) <> 0 Then
                    .TextMatrix(lngRow, COL_����) = FormatEx(rsMore!���˼���, 5)
                End If
                If Val(.TextMatrix(lngRow, COL_����)) = 0 Then .TextMatrix(lngRow, COL_����) = ""
                
                'ҩƷ��������:��װ��λ
                If cbo��Ч.ListIndex = 1 Then
                    If NVL(rsMore!�Ƴ�, 1) > sng���� Then sng���� = NVL(rsMore!�Ƴ�, 1)
                    If .TextMatrix(lngRow, COL_Ƶ��) <> "" And Val(.TextMatrix(lngRow, COL_����)) <> 0 _
                        And Val(.TextMatrix(lngRow, COL_����ϵ��)) <> 0 And Val(.TextMatrix(lngRow, COL_��װϵ��)) <> 0 Then
                        If Val(.TextMatrix(lngRow, COL_Ƶ������)) = 1 Then '����ҩƷ����ȱʡΪһ����
                            '�����Ƴ����Ϊ��������ҩ������
                            .TextMatrix(lngRow, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                    Val(.TextMatrix(lngRow, COL_����)), 1, 1, 1, "��", "", _
                                    Val(.TextMatrix(lngRow, COL_����ϵ��)), _
                                    Val(.TextMatrix(lngRow, COL_��װϵ��)), _
                                    Val(.TextMatrix(lngRow, COL_�ɷ����))), 5)
                        Else
                            '�����Ƴ����Ϊ��������ҩ������
                            .TextMatrix(lngRow, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                    Val(.TextMatrix(lngRow, COL_����)), sng����, _
                                    Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)), _
                                    Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)), _
                                    .TextMatrix(lngRow, COL_�����λ), _
                                    .TextMatrix(lngRow, COL_ִ��ʱ��), _
                                    Val(.TextMatrix(lngRow, COL_����ϵ��)), _
                                    Val(.TextMatrix(lngRow, COL_��װϵ��)), _
                                    Val(.TextMatrix(lngRow, COL_�ɷ����))), 5)
                        End If
                    End If
                End If
            End If
            
            '��¼ȱʡ����
            If cbo��Ч.ListIndex = 1 And Val(.TextMatrix(lngRow, COL_Ƶ������)) <> 1 Then
                .TextMatrix(lngRow, COL_����) = IIF(sng���� = 0, "", sng����)
            End If
        End If
        
        If rsMore.Filter <> 0 Then rsMore.Filter = 0
        
        'ִ��ʱ��:"��ѡƵ��"(ҩƷ�ǿ�ѡƵ��,����������Ϊһ����)
        If Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Then
            If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then
                If lngCopyRow <> -1 Then '����һ����ͬ
                    If .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��) Then
                        .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngCopyRow, COL_ִ��ʱ��)
                    End If
                End If
                If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then  'ȱʡʱ�䷽��
                    .TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(1, .TextMatrix(lngRow, COL_Ƶ��), lng��ҩ;��ID)
                End If
            End If
        End If
        
        '�����д������֮��������,�����ҽ������
        '-------------------------------------------------------------------------
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            '����һ����ҩ;����Ŀ,���������
            If lng��ҩ;��ID <> 0 Then
                .TextMatrix(lngRow, COL_�÷�) = sys.RowValue("������ĿĿ¼", lng��ҩ;��ID, "����")
            End If
            If lngGroupRow <> 0 Then
                'һ����ҩ�Ĺ�����ͬ�ĸ�ҩ;����
                lngTmp = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_���ID)), lngGroupRow + 1)
                If lngTmp > lngRow Then
                    .TextMatrix(lngRow, COL_���ID) = .TextMatrix(lngGroupRow, COL_���ID)
                Else
                    '��������ǽ�Ϊ��ʹ��һ����ҩ����ͬ����
                    .TextMatrix(lngRow, COL_���ID) = AdviceSet��ҩ;��(lngRow, lng��ҩ;��ID)
                End If
            Else '���������ĳ�ҩ���������ĸ�ҩ;����
                .TextMatrix(lngRow, COL_���ID) = AdviceSet��ҩ;��(lngRow, lng��ҩ;��ID)
            End If
            
            '���龫����ɫ��ʶ
            If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(lngRow, COL_�������)) > 0 _
                And .TextMatrix(lngRow, COL_�������) <> "" Then
                .Cell(flexcpFontBold, lngRow, col_ҽ������) = True
            End If
        ElseIf rsInput!���ID = "D" And strExtData <> "" Then
            '������ϲ�λ��
            Call AdviceSet������(lngRow, strExtData)
        ElseIf rsInput!���ID = "F" And strExtData <> "" Then
            '�����ĸ���������������Ŀ��
            Call AdviceSet�������(lngRow, strExtData)
        ElseIf rsInput!���ID = "K" Then
            '��Ѫ��;����
            If lng��ҩ;��ID <> 0 Then
                If gblnѪ��ϵͳ = True Then
                    strSQL = "Select a.����,a.��������,a.ִ�з��� From ������ĿĿ¼ A where a.id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ҩ;��ID)
                    .TextMatrix(lngRow, COL_�÷�) = rsTmp!���� & ""
                
                    If Val(rsTmp!�������� & "") = 8 And Val(rsTmp!ִ�з��� & "") = 1 Then '����Ǳ༭���������뵥ʱ��Ҫ����һ��
                        .TextMatrix(lngRow, COL_��鷽��) = 1
                    Else
                        .TextMatrix(lngRow, COL_��鷽��) = ""
                    End If
                Else
                    .TextMatrix(lngRow, COL_�÷�) = Sys.RowValue("������ĿĿ¼", lng��ҩ;��ID, "����")
                End If
                Call AdviceSet��Ѫ;��(lngRow, lng��ҩ;��ID)
            End If
        End If
        
        'ҽ������
        .TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AdviceInputFree(ByVal lngRow As Long)
'���ܣ�����������������ҽ��
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim lngCopyRow As Long
    
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) <> 0 Then
            If txtҽ������.Text <> .TextMatrix(lngRow, col_ҽ������) Then
                .TextMatrix(lngRow, col_ҽ������) = txtҽ������.Text
                mblnNoSave = True '���Ϊδ����
            End If
        Else
            .RowData(lngRow) = GetNextID
            .TextMatrix(lngRow, COL_��Ч) = zlCommFun.GetNeedName(cbo��Ч.Text)
            
            '���:��������,��ǰ��ռ������ź�,�������������
            .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
            Call AdviceSetҽ�����(lngRow + 1, 1)
                            
            .TextMatrix(lngRow, col_ҽ������) = txtҽ������.Text
            .TextMatrix(lngRow, COL_���) = "*" '������,Ϊ��������Ҫ
            .TextMatrix(lngRow, COL_������ĿID) = 0
            
            .TextMatrix(lngRow, COL_ִ������) = 4 '����ѡִ�п��Ҵ���ȱʡΪ��
            .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID("*", 0, 0, 4, cbo��Ч.ListIndex, mint��Χ)
            mblnNoSave = True '���Ϊδ����
            
            Call vsAdvice_AfterRowColChange(-1, -1, lngRow, .Col)
        End If
    End With
End Sub

Private Sub AdviceSet������(ByVal lngRow As Long, ByVal strExData As String)
'���ܣ���������ָ����������Ŀ�Ĳ�λ������,�����������������Ŀ���޸Ĳ�λ����
'������lngRow=��ǰ������
'      strExData=������鲿λ��������Ϣ,��ʽΪ:"��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
    Dim arrItems As Variant, arrMethod As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim str��鲿λ As String
    
    'ɾ�����еļ�鲿λ������
    Call Delete���������Ѫ(lngRow)
    
    '���¼��벿λ������
    If strExData <> "" Then
        arrItems = Split(Split(strExData, vbTab)(0), "|")
        For i = 0 To UBound(arrItems)
            str��鲿λ = Split(arrItems(i), ";")(0)
            arrMethod = Split(Split(arrItems(i), ";")(1), ",")
            For j = 0 To UBound(arrMethod)
                k = k + 1
                With vsAdvice
                    .AddItem "", lngRow + k
                    .RowHidden(lngRow + k) = True
                    
                    .RowData(lngRow + k) = GetNextID
                    .TextMatrix(lngRow + k, COL_���ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + k, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + k
                    .TextMatrix(lngRow + k, COL_��Ч) = .TextMatrix(lngRow, COL_��Ч)
                    
                    .TextMatrix(lngRow + k, COL_���) = .TextMatrix(lngRow, COL_���)
                    .TextMatrix(lngRow + k, COL_������ĿID) = .TextMatrix(lngRow, COL_������ĿID) 'Ϊͬһ�������Ŀ
                    
                    .TextMatrix(lngRow + k, COL_���㷽ʽ) = .TextMatrix(lngRow, COL_���㷽ʽ)
                    .TextMatrix(lngRow + k, COL_Ƶ������) = .TextMatrix(lngRow, COL_Ƶ������)
                    .TextMatrix(lngRow + k, COL_��������) = .TextMatrix(lngRow, COL_��������)
                    
                    .TextMatrix(lngRow + k, col_ҽ������) = .TextMatrix(lngRow, COL_����) '��¼Ϊ�����Ŀ����
                    .TextMatrix(lngRow + k, COL_�걾��λ) = str��鲿λ
                    .TextMatrix(lngRow + k, COL_��鷽��) = arrMethod(j)
                    
                    .TextMatrix(lngRow + k, COL_����) = .TextMatrix(lngRow, COL_����)
                    .TextMatrix(lngRow + k, COL_����) = .TextMatrix(lngRow, COL_����)
                    
                    .TextMatrix(lngRow + k, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                    .TextMatrix(lngRow + k, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                    .TextMatrix(lngRow + k, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                    .TextMatrix(lngRow + k, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                    .TextMatrix(lngRow + k, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                    
                    .TextMatrix(lngRow + k, COL_ִ������) = .TextMatrix(lngRow, COL_ִ������)
                    .TextMatrix(lngRow + k, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                End With
            Next
        Next
                
        '��������ҽ�������
        Call AdviceSetҽ�����(lngRow + k + 1, k)
    End If
End Sub

Private Sub AdviceSet�������(ByVal lngRow As Long, ByVal strDataIDs As String)
'���ܣ���������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������lngRow=��ǰ������
'      strDataIDs=��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    'ɾ�����еĸ���������������Ŀ��
    Call Delete���������Ѫ(lngRow)
    
    '���¼��븽�������м�������Ŀ��
    strDataIDs = Trim(Replace(strDataIDs, ";", ","))
    If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
    If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    
    If strDataIDs <> "" Then
        Set rsTmp = Get������Ŀ��¼(0, strDataIDs)
        If Not rsTmp.EOF Then
            arrIDs = Split(strDataIDs, ",")
            For i = 0 To UBound(arrIDs) '���û�������Ŀ˳��
                rsTmp.Filter = "ID=" & CStr(arrIDs(i)) '������EOF
                
                With vsAdvice
                    .AddItem "", lngRow + i + 1
                    .RowHidden(lngRow + i + 1) = True
                    
                    .RowData(lngRow + i + 1) = GetNextID
                    .TextMatrix(lngRow + i + 1, COL_���ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + i + 1, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + i + 1
                    .TextMatrix(lngRow + i + 1, COL_��Ч) = .TextMatrix(lngRow, COL_��Ч)
                    
                    .TextMatrix(lngRow + i + 1, COL_���) = rsTmp!���
                    .TextMatrix(lngRow + i + 1, COL_������ĿID) = rsTmp!ID
                    .TextMatrix(lngRow + i + 1, COL_���㷽ʽ) = NVL(rsTmp!���㷽ʽ, 0)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ������) = NVL(rsTmp!ִ��Ƶ��, 0)
                    .TextMatrix(lngRow + i + 1, COL_��������) = NVL(rsTmp!��������)
                    
                    .TextMatrix(lngRow + i + 1, COL_�걾��λ) = NVL(rsTmp!�걾��λ)
                    .TextMatrix(lngRow + i + 1, col_ҽ������) = rsTmp!����
                    
                    .TextMatrix(lngRow + i + 1, COL_����) = .TextMatrix(lngRow, COL_����)
                    .TextMatrix(lngRow + i + 1, COL_����) = .TextMatrix(lngRow, COL_����)
                    
                    .TextMatrix(lngRow + i + 1, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                    .TextMatrix(lngRow + i + 1, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                    
                    'ִ������:������Ŀ��������
                    .TextMatrix(lngRow + i + 1, COL_ִ������) = NVL(rsTmp!ִ�п���, 0)
                    
                    '������Ժ��ִ����ִ�п���,����������ִ�п���
                    '���򲻹���ִ�п�������,һ���������Ӧ����ͬ
                    If InStr(",0,5,", NVL(rsTmp!ִ�п���, 0)) > 0 Then
                        .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = 0
                    Else
                        If rsTmp!��� = "G" Then
                            .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = Get����ִ�п���ID(rsTmp!���, rsTmp!ID, 0, NVL(rsTmp!ִ�п���, 0), IIF(.TextMatrix(lngRow, COL_��Ч) = "����", 0, 1), mint��Χ)
                        Else
                            .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                        End If
                    End If
                End With
            Next
                
            '�������
            Call AdviceSetҽ�����(lngRow + UBound(arrIDs) + 2, UBound(arrIDs) + 1)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSet��ҩ;��(ByVal lngRow As Long, ByVal lng��ҩ;��ID As Long, Optional strִ������ As String, Optional ByVal str���� As String) As Long
'���ܣ�Ϊ¼����У�����ҩ���ö�Ӧ�ĸ�ҩ;����(�������޸�)
'������lngRow=Ҫ�����ҩ;����ҩƷ��
'      lng��ҩ;��ID=��ҩ;��ID
'      strִ������=�޸ĸ�ҩ;��ʱ,��ǰ�������õ�ִ������
'      str����=�޸ĸ�ҩ;��ʱ,��ǰ�������õĵ���
'���أ������õĸ�ҩ;���е�ҽ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    Set rsTmp = Get������Ŀ��¼(lng��ҩ;��ID)
    If rsTmp.EOF Then lng��ҩ;��ID = 0 'û�����ݣ��������Ա��ֹ�ϵ
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then 'δ����"���ID"ʱ
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        Else
            '�޸�ҽ��������ʱ�������ø�ҩ;������(���Ǹ���������Ŀ)
            blnNew = False
            lngNewRow = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
        End If
        
        '��Ч���ݣ�����,�շ�ϸĿID,����ϵ��,��װ��λ,��װϵ��,�걾��λ,ҽ������,����,����,�÷�
        If blnNew Then
            .RowData(lngNewRow) = GetNextID
            .TextMatrix(lngNewRow, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + 1
            .TextMatrix(lngNewRow, col_ȱʡ) = .TextMatrix(lngRow, col_ȱʡ)
        End If
        
        .TextMatrix(lngNewRow, COL_��Ч) = .TextMatrix(lngRow, COL_��Ч)
        
        .TextMatrix(lngNewRow, COL_���) = "E" '��ҩ;����������
        .TextMatrix(lngNewRow, COL_������ĿID) = lng��ҩ;��ID
        
        '���û��ȷ����ҩ;������ʱ�����õ�����
        If Not rsTmp.EOF Then
            .TextMatrix(lngNewRow, COL_���㷽ʽ) = NVL(rsTmp!���㷽ʽ, 0)
            .TextMatrix(lngNewRow, COL_��������) = NVL(rsTmp!��������)
            .TextMatrix(lngNewRow, COL_ִ�з���) = NVL(rsTmp!ִ�з���, 0)
            .TextMatrix(lngNewRow, col_ҽ������) = rsTmp!����
            
            '����
            If str���� <> "" Then
                .TextMatrix(lngNewRow, COL_ҽ������) = str����
            End If
            'ִ������:ȱʡ������Ŀ����,�޸�ʱ���ݵ�ǰ��������
            If strִ������ = "" Then
                .TextMatrix(lngNewRow, COL_ִ������) = NVL(rsTmp!ִ�п���, 0)
            Else
                .TextMatrix(lngNewRow, COL_ִ������) = decode(strִ������, "��Ժ��ҩ", 5, NVL(rsTmp!ִ�п���, 0))
            End If
            
            If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_ִ������))) = 0 Then
                .TextMatrix(lngNewRow, COL_ִ�п���ID) = Get����ִ�п���ID("E", lng��ҩ;��ID, 0, NVL(rsTmp!ִ�п���, 0), IIF(.TextMatrix(lngRow, COL_��Ч) = "����", 0, 1), mint��Χ)
            Else
                .TextMatrix(lngNewRow, COL_ִ�п���ID) = 0
            End If
        End If
        
        .TextMatrix(lngNewRow, COL_Ƶ������) = .TextMatrix(lngRow, COL_Ƶ������) '��ҩƷ��Ϊ׼
        .TextMatrix(lngNewRow, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
        .TextMatrix(lngNewRow, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
        .TextMatrix(lngNewRow, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
        .TextMatrix(lngNewRow, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
        .TextMatrix(lngNewRow, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
        
        '����������
        If blnNew Then Call AdviceSetҽ�����(lngNewRow + 1, 1)
        
        AdviceSet��ҩ;�� = .RowData(lngNewRow)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet��Ѫ;��(ByVal lngRow As Long, ByVal lng��Ѫ;��ID As Long) As Long
'���ܣ�Ϊ¼����У�����ҩ���ö�Ӧ�ĸ�ҩ;����(�������޸�)
'������lngRow=Ҫ������Ѫ;������Ѫҽ����
'      lng��Ѫ;��ID=��Ѫ;��ID
'���أ������õ���Ѫ;���е�ҽ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    Set rsTmp = Get������Ŀ��¼(lng��Ѫ;��ID)
    
    With vsAdvice
        lngNewRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
        If lngNewRow = -1 Then '��δ������Ѫ;��ʱ
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        End If
        
        '��Ч���ݣ�����,�շ�ϸĿID,����ϵ��,��װ��λ,��װϵ��,�걾��λ,ҽ������,����,����,�÷�
        If blnNew Then
            .RowData(lngNewRow) = GetNextID
            .TextMatrix(lngNewRow, COL_���ID) = .RowData(lngRow)
            .TextMatrix(lngNewRow, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + 1
            .TextMatrix(lngNewRow, col_ȱʡ) = .TextMatrix(lngRow, col_ȱʡ)
        End If
        
        .TextMatrix(lngNewRow, COL_��Ч) = .TextMatrix(lngRow, COL_��Ч)
        
        .TextMatrix(lngNewRow, COL_���) = "E" '��Ѫ;����������
        .TextMatrix(lngNewRow, COL_������ĿID) = lng��Ѫ;��ID
        
        .TextMatrix(lngNewRow, COL_���㷽ʽ) = NVL(rsTmp!���㷽ʽ, 0)
        .TextMatrix(lngNewRow, COL_��������) = NVL(rsTmp!��������)
        .TextMatrix(lngNewRow, col_ҽ������) = rsTmp!����
        .TextMatrix(lngNewRow, COL_ִ������) = NVL(rsTmp!ִ�п���, 0)
        
        If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_ִ������))) = 0 Then
            .TextMatrix(lngNewRow, COL_ִ�п���ID) = Get����ִ�п���ID("E", lng��Ѫ;��ID, 0, NVL(rsTmp!ִ�п���, 0), IIF(.TextMatrix(lngRow, COL_��Ч) = "����", 0, 1), mint��Χ)
        Else
            .TextMatrix(lngNewRow, COL_ִ�п���ID) = 0
        End If
        
        .TextMatrix(lngNewRow, COL_Ƶ������) = .TextMatrix(lngRow, COL_Ƶ������) '��ҩƷ��Ϊ׼
        .TextMatrix(lngNewRow, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
        .TextMatrix(lngNewRow, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
        .TextMatrix(lngNewRow, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
        .TextMatrix(lngNewRow, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
        .TextMatrix(lngNewRow, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
        
        '����������
        If blnNew Then Call AdviceSetҽ�����(lngNewRow + 1, 1)
        
        AdviceSet��Ѫ;�� = .RowData(lngNewRow)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceChange()
'���ܣ����ݵ�ǰҽ����Ƭ�е����ݣ����µ�ǰҽ������
'˵��������ListIndex=-1����Ӧҽ�����������ݵģ�����ԭ���ݲ�����
    Dim lngRow As Long, lngBeginRow As Long
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim blnCurDo As Boolean, blnTmp As Boolean
    Dim lngTmp As Long, strTmp As String
    Dim lngִ�п���ID As Long, lng��������ID As Long
    Dim blnReInRow As Boolean, i As Long, j As Long

    With vsAdvice
        lngRow = .Row

        If .RowData(lngRow) = 0 Then Call ClearItemTag: Exit Sub    '����༭��־

        If RowIn�䷽��(lngRow) Then
            '��ҩ�䷽
            lngBeginRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            For i = lngBeginRow To lngRow
                '�޸Ĵ����䷽������������(�����巨���÷�)
                If txt����.Enabled And IsNumeric(txt����.Text) And txt����.Tag <> "" Then
                    .TextMatrix(i, COL_����) = FormatEx(Val(txt����.Text), 5)
                    blnCurDo = True
                End If
                If txtƵ��.Enabled And cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                    .TextMatrix(i, COL_Ƶ��) = txtƵ��.Text
                    Call GetƵ����Ϣ_����(txtƵ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ, 2)    '��ҽ��Χ
                    .TextMatrix(i, COL_Ƶ�ʴ���) = intƵ�ʴ���
                    .TextMatrix(i, COL_Ƶ�ʼ��) = intƵ�ʼ��
                    .TextMatrix(i, COL_�����λ) = str�����λ
                    blnCurDo = True
                End If
                If cboִ��ʱ��.Tag <> "" Then
                    .TextMatrix(i, COL_ִ��ʱ��) = cboִ��ʱ��.Text
                    blnCurDo = True
                End If

                '����֤��
                If txt����֤��.Tag <> "" Then
                    .TextMatrix(i, COL_�����ĿID) = txt����֤��.Tag
                    .TextMatrix(i, COL_����֤��) = txt����֤��.Text
                    blnCurDo = True
                End If

                If .TextMatrix(i, COL_���) = "7" Then
                    '���ĵ��������ҩ��ִ�п���(�÷��巨�ĸĲ���)
                    If cboִ�п���.Tag <> "" Then
                        If cboִ�п���.ListIndex <> -1 Then
                            .TextMatrix(i, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                        Else
                            .TextMatrix(i, COL_ִ�п���ID) = ""
                        End If
                        blnCurDo = True
                    End If

                    'ִ������:�䷽��������ɵ���ҩ��ͬ
                    If cboִ������.Tag <> "" Then
                        .TextMatrix(i, COL_ִ������) = decode(zlCommFun.GetNeedName(cboִ������.Text), "�Ա�ҩ", 5, "��ȡҩ", 5, 4)
                        .TextMatrix(i, COL_ִ�б��) = decode(zlCommFun.GetNeedName(cboִ������.Text), "��ȡҩ", 1, "��ȡҩ", 2, 0)
                        If Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                            .TextMatrix(i, COL_ִ�п���ID) = 0
                        ElseIf Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                            '�ָ�ȱʡִ�п���,ȱʡ��ǰ����ͬ
                            If i = lngBeginRow Then
                                For j = i - 1 To .FixedRows Step -1
                                    If .TextMatrix(j, COL_���) = "7" And Val(.TextMatrix(j, COL_ִ�п���ID)) <> 0 Then
                                        .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(j, COL_ִ�п���ID)
                                        Exit For
                                    End If
                                Next
                                If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = Get����ִ�п���ID(.TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), Val(.TextMatrix(i, COL_�շ�ϸĿID)), 4, cbo��Ч.ListIndex, mint��Χ)
                                End If
                            Else
                                .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngBeginRow, COL_ִ�п���ID)
                            End If
                        End If
                        blnReInRow = True    '����ִ�п��ұ༭�Ա仯
                        blnCurDo = True
                    End If
                End If

                '�޸�ʱ�Զ����²�������
                blnTmp = False
                If cboҽ������.Tag <> "" Or cboִ������.Tag <> "" _
                   Or (Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "") Then
                    blnTmp = True
                End If

                If .TextMatrix(i, COL_���) = "E" And i <> lngRow Then lngTmp = i    '�巨�к�

                '---------------
                If blnCurDo Then mblnNoSave = True    '���Ϊδ����
            Next

            '�漰��ҩ�÷��е�����:ֱ�Ӹ��ĵ�ǰ�е�����(�巨�����䷽�༭�в��ܸ�)
            '-----------------------------------------------------------
            blnCurDo = False

            'ҽ������:�Ƿ�����ҩ�÷���(��ʾ��)�е�
            If cboҽ������.Tag <> "" Then
                .TextMatrix(lngRow, COL_ҽ������) = cboҽ������.Text
                blnCurDo = True
            End If

            '��ҩ�÷�
            If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                .TextMatrix(lngRow, COL_������ĿID) = Val(cmd�÷�.Tag)
                .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text

                'ͬʱ����ִ������
                i = NVL(sys.RowValue("������ĿĿ¼", Val(cmd�÷�.Tag), "ִ�п���"), 0)
                .TextMatrix(lngRow, COL_ִ������) = decode(zlCommFun.GetNeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID("E", Val(cmd�÷�.Tag), 0, Val(.TextMatrix(lngRow, COL_ִ������)), cbo��Ч.ListIndex, mint��Χ)
                End If

                blnReInRow = True    '��Ҫˢ����ҩ�÷�ִ�п���
                blnCurDo = True
            End If

            '�÷��ͼ巨��ִ������
            If cboִ������.Tag <> "" Then
                '�÷�
                i = NVL(sys.RowValue("������ĿĿ¼", Val(.TextMatrix(lngRow, COL_������ĿID)), "ִ�п���"), 0)
                .TextMatrix(lngRow, COL_ִ������) = decode(zlCommFun.GetNeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(.TextMatrix(lngRow, COL_���), Val(.TextMatrix(lngRow, COL_������ĿID)), 0, Val(.TextMatrix(lngRow, COL_ִ������)), cbo��Ч.ListIndex, mint��Χ)
                End If

                '�巨
                i = NVL(sys.RowValue("������ĿĿ¼", Val(.TextMatrix(lngTmp, COL_������ĿID)), "ִ�п���"), 0)
                .TextMatrix(lngTmp, COL_ִ������) = decode(zlCommFun.GetNeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngTmp, COL_ִ������)) = 5 Then
                    .TextMatrix(lngTmp, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngTmp, COL_ִ�п���ID) = Get����ִ�п���ID(.TextMatrix(lngTmp, COL_���), Val(.TextMatrix(lngTmp, COL_������ĿID)), 0, Val(.TextMatrix(lngTmp, COL_ִ������)), cbo��Ч.ListIndex, mint��Χ)
                End If

                mblnNoSave = True    '���Ϊδ����

                blnCurDo = True
            End If

            '��ҩ�÷�ִ�п���:���䷽��ǰ��ʾ�е�ִ�п���
            If cbo����ִ��.Tag <> "" Then
                If cbo����ִ��.ListIndex <> -1 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                Else
                    .TextMatrix(lngRow, COL_ִ�п���ID) = ""
                End If
                blnCurDo = True
            End If

            '---------------
            If blnCurDo Then mblnNoSave = True    '���Ϊδ����
        Else    '����������Ŀ
            If txt����.Enabled And (IsNumeric(txt����.Text) Or txt����.Text = "") And txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = FormatEx(txt����.Text, 5)
                blnCurDo = True
            End If

            If txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = txt����.Text
                blnCurDo = True
            End If

            If txt����.Enabled And (IsNumeric(txt����.Text) Or txt����.Text = "") And txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = FormatEx(txt����.Text, 5)
                blnCurDo = True
            End If

            If txtƵ��.Enabled And cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                'Ƶ�������Ѿ�������ʱȷ��(����������һ����֮���л�)
                .TextMatrix(lngRow, COL_Ƶ��) = txtƵ��.Text
                Call GetƵ����Ϣ_����(txtƵ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ, GetƵ�ʷ�Χ(lngRow))
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngRow, COL_�����λ) = str�����λ
                blnCurDo = True
            End If

            If cboִ��ʱ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_ִ��ʱ��) = cboִ��ʱ��.Text
                blnCurDo = True
            End If
            If cboҽ������.Tag <> "" Then
                .TextMatrix(lngRow, COL_ҽ������) = cboҽ������.Text
                blnCurDo = True
            End If

            If cboִ�п���.Tag <> "" Then
                If Not RowIn������(lngRow) Then    '�ɼ�������ִ�п��Ҳ�ͬ
                    If cboִ�п���.ListIndex <> -1 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                    Else
                        .TextMatrix(lngRow, COL_ִ�п���ID) = ""
                    End If
                End If
                blnCurDo = True
            End If
            
            '���٣���ҺҩƷ
            If cbo����.Tag <> "" Then
                lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                If lngTmp <> -1 Then
                    If cbo����.Text <> "" Then
                        .TextMatrix(lngTmp, COL_ҽ������) = cbo����.Text & lbl���ٵ�λ.Caption
                    Else
                        .TextMatrix(lngTmp, COL_ҽ������) = ""
                    End If
                    blnCurDo = True
                End If
                If cbo����.Text <> "" Then
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text & cbo����.Text & lbl���ٵ�λ.Caption
                Else
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                End If
            End If
            
            '����ִ�п��ң���ҩ;��,��������,�ɼ�����
            If cbo����ִ��.Tag <> "" Then
                lngTmp = -1
                If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                ElseIf .TextMatrix(lngRow, COL_���) = "F" Then
                    For i = lngRow + 1 To .Rows - 1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If .TextMatrix(i, COL_���) = "G" Then
                                lngTmp = i: Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf .TextMatrix(lngRow, COL_���) = "E" _
                       And .TextMatrix(lngRow - 1, COL_���) = "C" _
                       And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow
                ElseIf .TextMatrix(lngRow, COL_���) = "K" _
                    And .TextMatrix(lngRow + 1, COL_���) = "E" _
                    And Val(.TextMatrix(lngRow + 1, COL_���ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow + 1
                End If

                'ֻ���¶�Ӧ��,��Ӱ��������
                If lngTmp <> -1 Then
                    If cbo����ִ��.ListIndex <> -1 Then
                        .TextMatrix(lngTmp, COL_ִ�п���ID) = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                    Else
                        .TextMatrix(lngTmp, COL_ִ�п���ID) = ""
                    End If
                    mblnNoSave = True    '���Ϊδ����
                End If
            End If

            'ִ������,��ҩ;��:Ϊ���¿���ʱ��(������ҩ;����ͬ������),���ж��Ƿ�ı�
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                If cboִ������.Tag <> "" Then blnCurDo = True
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then blnCurDo = True
            End If

            '�޸�ʱ�Զ����²�������
            blnTmp = False
            If cboִ������.Tag <> "" Or (Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "") Then
                blnReInRow = True    '��Ҫˢ�¸�ҩ;��,�ɼ���ʽ��ִ�п���
                blnTmp = True
            End If

            '������Ҫͬ������Ĺ�����
            '----------------------------------------------------------------
            If RowIn������(lngRow) Then
                '�ɼ�����
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                    .TextMatrix(lngRow, COL_������ĿID) = Val(cmd�÷�.Tag)
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                    .TextMatrix(lngRow, COL_����) = txt�÷�.Text

                    'ͬʱ����ִ������
                    .TextMatrix(lngRow, COL_ִ������) = NVL(sys.RowValue("������ĿĿ¼", Val(cmd�÷�.Tag), "ִ�п���"), 0)
                    If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID("E", Val(cmd�÷�.Tag), 0, Val(.TextMatrix(lngRow, COL_ִ������)), cbo��Ч.ListIndex, mint��Χ)
                    Else
                        .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                    End If
                    
                    blnCurDo = True
                End If

                '����һ���ɼ��ĸ���������Ŀ
                If blnCurDo Then
                    For i = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                            End If
                            If txtƵ��.Tag <> "" Then
                                .TextMatrix(i, COL_Ƶ������) = .TextMatrix(lngRow, COL_Ƶ������)
                                .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                            End If
                            If cboִ�п���.Tag <> "" Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Or cboִ�п���.ListIndex = -1 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = 0
                                Else
                                    .TextMatrix(i, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                                End If
                                
                            End If
                            If cboִ��ʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                                
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '�С�����ҩ�����ҩ;����һ����ҩ�����

                'ִ������
                If cboִ������.Tag <> "" Then
                    .TextMatrix(lngRow, COL_ִ������) = decode(zlCommFun.GetNeedName(cboִ������.Text), "�Ա�ҩ", 5, "��ȡҩ", 5, 4)
                    .TextMatrix(lngRow, COL_ִ�б��) = decode(zlCommFun.GetNeedName(cboִ������.Text), "��ȡҩ", 1, "��ȡҩ", 2, 0)
                    If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                    ElseIf Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                        '�ָ�ȱʡҩ��,ȱʡ��ǰ��ĳ�ҩ��ͬ
                        strTmp = Get����ҩ��IDs(.TextMatrix(lngRow, COL_���), Val(.TextMatrix(lngRow, COL_������ĿID)), Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), 0, mint��Χ)
                        For i = lngRow - 1 To .FixedRows Step -1
                            '����ҩ���г�ҩ��ҩ�����ܲ�ͬ,�������Ҫ��ͬ
                            If .TextMatrix(i, COL_���) = .TextMatrix(lngRow, COL_���) And Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 Then
                                If InStr("," & strTmp & ",", "," & Val(.TextMatrix(i, COL_ִ�п���ID)) & ",") > 0 Then
                                    .TextMatrix(lngRow, COL_ִ�п���ID) = Val(.TextMatrix(i, COL_ִ�п���ID))
                                    Exit For
                                End If
                            End If
                        Next
                        If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                            .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(.TextMatrix(lngRow, COL_���), Val(.TextMatrix(lngRow, COL_������ĿID)), Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), 4, cbo��Ч.ListIndex, mint��Χ)
                        End If
                    End If

                    cboִ�п���.Tag = "1"    '����ִ�п���һ����ҩ��Ҫͬ����
                    blnReInRow = True    '����ִ�п��ұ༭�Ա仯
                End If

                '��ҩ;�����������������ͬ������
                strTmp = ""
                If Trim(cbo����.Text) <> "" Then
                    strTmp = cbo����.Text & lbl���ٵ�λ.Caption
                End If
                
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text & strTmp
                    Call AdviceSet��ҩ;��(lngRow, Val(cmd�÷�.Tag), zlCommFun.GetNeedName(cboִ������.Text), strTmp)
                ElseIf blnCurDo Then    'cboִ������.Tag <> "" Then
                    '���ִ�����ʸ�����,��Ҫǿ���޸Ķ�Ӧ�ĸ�ҩ;����ִ�����ʺ�ִ�п���
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    Call AdviceSet��ҩ;��(lngRow, Val(.TextMatrix(lngTmp, COL_������ĿID)), zlCommFun.GetNeedName(cboִ������.Text), strTmp)
                End If

                'һ����ҩ:�������ҩ;��,ǰ���ѵ�������
                If blnCurDo Then
                    lngBeginRow = .FindRow(.TextMatrix(lngRow, COL_���ID), , COL_���ID)
                    For i = lngBeginRow To .Rows - 1
                        If i <> lngRow And .RowData(i) <> 0 Then    '���������м��п���
                            If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                                If txt�÷�.Tag <> "" Then
                                    .TextMatrix(i, COL_�÷�) = .TextMatrix(lngRow, COL_�÷�)
                                    '�����ٲ��һ����ҩ����������
                                    If cbo����.Tag <> "" Then
                                        .TextMatrix(i, COL_�÷�) = txt�÷�.Text & strTmp
                                    End If
                                End If
                                
                                
                                If txtƵ��.Tag <> "" Then
                                    .TextMatrix(i, COL_Ƶ������) = .TextMatrix(lngRow, COL_Ƶ������)    '��Ҫͬ������,��Ϊ����������һ����֮���л�
                                    .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                    .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                    .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                    .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                                End If

                                'һ����ҩ��,������ͬ�仯,�������¼���
                                If txt����.Tag <> "" Then
                                    .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                    If txt����.Text <> "" And .TextMatrix(i, COL_Ƶ��) <> "" _
                                       And Val(.TextMatrix(i, COL_Ƶ������)) <> 1 And Val(.TextMatrix(i, COL_����)) <> 0 _
                                       And Val(.TextMatrix(i, COL_����ϵ��)) <> 0 And Val(.TextMatrix(i, COL_��װϵ��)) <> 0 Then

                                        .TextMatrix(i, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                                                          Val(.TextMatrix(i, COL_����)), Val(.TextMatrix(i, COL_����)), _
                                                                          Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), _
                                                                          .TextMatrix(i, COL_�����λ), .TextMatrix(i, COL_ִ��ʱ��), _
                                                                          Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_��װϵ��)), _
                                                                          Val(.TextMatrix(i, COL_�ɷ����))), 5)
                                    End If
                                End If

                                If cboִ��ʱ��.Tag <> "" Then
                                    .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                                End If

                                'ִ�����ʡ�ִ�б��:��Ժ��ҩ����ȡҩ��һ����ҩ����һ�£������ɵ�������
                                If cboִ������.Tag <> "" And zlCommFun.GetNeedName(cboִ������.Text) = "��Ժ��ҩ" Or zlCommFun.GetNeedName(cboִ������.Text) = "��ȡҩ" Then
                                    .TextMatrix(i, COL_ִ������) = .TextMatrix(lngRow, COL_ִ������)
                                    .TextMatrix(i, COL_ִ�б��) = .TextMatrix(lngRow, COL_ִ�б��)
                                    '���Ա�ҩת����ʱ��Ҫ��������ִ�п���
                                    If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                                        .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                    End If
                                End If

                                'ִ�п���:ִ�п���(ҩ��)���Բ�ͬ,��������������
                                If cboִ�п���.Tag <> "" Then
                                    '�����и�Ϊ�Ա�ҩ����ĳ��Ϊ�Ա�ҩ�����������
                                    If Not (Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(lngRow, COL_ִ������)) = 5) _
                                       And Not (Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(.TextMatrix(i, COL_ִ������)) = 5) Then
                                        If sys.DeptHaveProperty(Val(.TextMatrix(lngRow, COL_ִ�п���ID)), "��������") Then
                                            '������ҩƷ����ͨҩ���������������ĸ�Ϊ�µ���������,�����ҩ����Ϊ����������
                                            .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                        ElseIf sys.DeptHaveProperty(Val(.TextMatrix(i, COL_ִ�п���ID)), "��������") Then
                                            '������ҩƷ���������ĸĳ���ͨҩ��,�����ҩ����Ϊ����ͨҩ��
                                            .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                            
                                        End If
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    Next
                End If
            ElseIf .TextMatrix(lngRow, COL_���) = "K" Then
                '��Ѫҽ���Ĵ���
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                    Call AdviceSet��Ѫ;��(lngRow, Val(cmd�÷�.Tag))
                ElseIf blnCurDo Then
                    lngTmp = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
                    If lngTmp <> -1 Then
                        Call AdviceSet��Ѫ;��(lngRow, Val(.TextMatrix(lngTmp, COL_������ĿID)))
                    End If
                End If
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_���)) > 0 And blnCurDo Then
                '��������Ŀ�л�����������
                lngBeginRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
                If lngBeginRow <> -1 Then
                    For i = lngBeginRow To .Rows - 1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                            End If
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                            End If

                            If cboִ��ʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                            End If
                            If txtƵ��.Tag <> "" Then
                                .TextMatrix(i, COL_Ƶ������) = .TextMatrix(lngRow, COL_Ƶ������)
                                .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                            End If
                            If cboִ�п���.Tag <> "" Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = 0
                                ElseIf .TextMatrix(i, COL_���) <> "G" Then    '���������ִ�п���Ϊ����
                                    .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If

            If blnCurDo Then mblnNoSave = True    '���Ϊδ����
        End If

        '����ҽ������
        If AdviceTextChange(lngRow) Then
            .TextMatrix(lngRow, col_ҽ������) = AdviceTextMake(lngRow)
            txtҽ������.Text = .TextMatrix(lngRow, col_ҽ������)
        End If
    End With

    '����༭��־
    Call ClearItemTag

    'ĳЩ�������Ҫ�������ÿ�Ƭ����Ŀ�༭��(���޸���ִ������ʱ)
    If blnReInRow Then
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub AdviceSetһ����ҩ(ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ���ѡ��Χ�ڵ�ҩƷ����Ϊһ����ҩ
'��������ֹ�к�,�м䲻��������,���������һ��ҩƷ�ĸ�ҩ;����
'˵�����Ե�һ��ҩƷ�ĸ�ҩ;��Ϊ׼,��λ�÷������һ��ҩƷ֮��
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lngRow1 As Long, lngRow2 As Long
    Dim lng���ID As Long, i As Long
    Dim strStart As String, lng�������� As Long
    
    lngRow1 = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngBegin, COL_���ID)), lngBegin + 1) '��һ��ҩ;����
    lngRow2 = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngEnd, COL_���ID)), lngEnd + 1) '����ҩ;����
    
    'ɾ����ҩ;����֮ǰ��¼ִ������,�Ա�������ж�
    For i = lngRow2 To lngRow1 Step -1
        If vsAdvice.RowHidden(i) Then
            vsAdvice.Cell(flexcpData, i - 1, COL_ִ������) = Val(vsAdvice.TextMatrix(i, COL_ִ������))
        End If
    Next
    
    '���Ƶ�һ�еĸ�ҩ;�������һ�еĸ�ҩ;��
    For i = vsAdvice.FixedCols To vsAdvice.Cols - 1
        If i <> COL_���ID And i <> COL_��� Then
            vsAdvice.TextMatrix(lngRow2, i) = vsAdvice.TextMatrix(lngRow1, i)
        End If
    Next
    lng���ID = vsAdvice.RowData(lngRow2)
    
    varTmp1 = mblnRowChange: varTmp2 = vsAdvice.Redraw
    mblnRowChange = False: vsAdvice.Redraw = flexRDNone
    
    'ɾ�������һ�и�ҩ;�����������ҩ;��
    For i = lngEnd To lngBegin Step -1
        If vsAdvice.RowHidden(i) Then
            Call DeleteRow(i)
        Else
            vsAdvice.TextMatrix(i, COL_���ID) = lng���ID
        End If
    Next
    
    '�к��ѱ��
    lngRow1 = lngBegin '��ʼһ����ҩ��
    
    '����һ����ҩ�����е���ͬ��Ϣ
    For i = lngRow1 + 1 To vsAdvice.Rows - 1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lng���ID Then
            lngRow2 = i '��¼�µĽ����к�
            
            'һ����ҩ�Ĳ�����Ϣ��ͬ
            vsAdvice.TextMatrix(i, col_ȱʡ) = vsAdvice.TextMatrix(lngRow1, col_ȱʡ)
            vsAdvice.TextMatrix(i, COL_����) = vsAdvice.TextMatrix(lngRow1, COL_����)
            vsAdvice.TextMatrix(i, COL_�÷�) = vsAdvice.TextMatrix(lngRow1, COL_�÷�)
            
            vsAdvice.TextMatrix(i, COL_Ƶ������) = vsAdvice.TextMatrix(lngRow1, COL_Ƶ������)
            vsAdvice.TextMatrix(i, COL_Ƶ��) = vsAdvice.TextMatrix(lngRow1, COL_Ƶ��)
            vsAdvice.TextMatrix(i, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngRow1, COL_Ƶ�ʴ���)
            vsAdvice.TextMatrix(i, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngRow1, COL_Ƶ�ʼ��)
            vsAdvice.TextMatrix(i, COL_�����λ) = vsAdvice.TextMatrix(lngRow1, COL_�����λ)
            vsAdvice.TextMatrix(i, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngRow1, COL_ִ��ʱ��)
            
            '��Ժ��ҩһ����ͬ
            If Val(vsAdvice.TextMatrix(lngRow1, COL_ִ������)) <> 5 And Val(vsAdvice.Cell(flexcpData, lngRow1, COL_ִ������)) = 5 Then
                '��һ������Ժ��ҩ,ȫ������Ϊ��Ժ��ҩ
                vsAdvice.TextMatrix(i, COL_ִ������) = vsAdvice.TextMatrix(lngRow1, COL_ִ������)
                If Val(vsAdvice.TextMatrix(i, COL_ִ�п���ID)) = 0 Then 'ִ�п��ҿ��Բ�ͬ,û��ʱ��ȱʡ��ͬ
                    vsAdvice.TextMatrix(i, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngRow1, COL_ִ�п���ID)
                End If
            ElseIf Val(vsAdvice.TextMatrix(i, COL_ִ������)) <> 5 And Val(vsAdvice.Cell(flexcpData, i, COL_ִ������)) = 5 Then
                '��ǰ������Ժ��ҩ,������Ϊ���һ����ͬ
                vsAdvice.TextMatrix(i, COL_ִ������) = vsAdvice.TextMatrix(lngRow1, COL_ִ������)
                If Val(vsAdvice.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                    vsAdvice.TextMatrix(i, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngRow1, COL_ִ�п���ID)
                End If
            Else
                '���򱣳ֲ���
            End If
        Else
            Exit For
        End If
    Next
    
    '�����ЩҩƷ���Ƿ��������������ҩ�ģ��Ե�һ��Ϊ׼
    For i = lngRow1 To vsAdvice.Rows - 1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lng���ID Then
            '�Ա�ҩ�����������
            If Not (Val(vsAdvice.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(vsAdvice.TextMatrix(i, COL_ִ������)) = 5) Then
                If sys.DeptHaveProperty(Val(vsAdvice.TextMatrix(i, COL_ִ�п���ID)), "��������") Then
                    lng�������� = Val(vsAdvice.TextMatrix(i, COL_ִ�п���ID)): Exit For
                End If
            End If
        Else
            Exit For
        End If
    Next
    '��������һ����ͬ
    If lng�������� <> 0 Then
        For i = lngRow1 To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lng���ID Then
                '�Ա�ҩ�����������
                If Not (Val(vsAdvice.TextMatrix(i, COL_ִ�п���ID)) = 0 And Val(vsAdvice.TextMatrix(i, COL_ִ������)) = 5) Then
                    vsAdvice.TextMatrix(i, COL_ִ�п���ID) = lng��������
                End If
            Else
                Exit For
            End If
        Next
    End If
    
    mblnRowChange = varTmp1: vsAdvice.Redraw = varTmp2
    mblnNoSave = True '���Ϊδ����
End Sub

Private Sub AdviceSet������ҩ(ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ�ȡ��һ��ҩƷ��һ����ҩ
'��������ֹ�к�,�м䲻��������,���������һ��ҩƷ�ĸ�ҩ;����
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lng��ҩ;��ID As Long, i As Long
    Dim intִ������ As Integer, strִ������ As String, str���� As String
    Dim lngRow As Long, curDate As Date
    
    With vsAdvice
        varTmp1 = mblnRowChange: varTmp2 = .Redraw
        mblnRowChange = False: .Redraw = flexRDNone
        
        'һ����ҩ;��
        lngRow = .FindRow(CLng(.TextMatrix(lngEnd, COL_���ID)), lngEnd + 1)
        lng��ҩ;��ID = Val(.TextMatrix(lngRow, COL_������ĿID))
        intִ������ = Val(.TextMatrix(lngRow, COL_ִ������))
        str���� = .TextMatrix(lngRow, COL_ҽ������)
        
        For i = lngEnd - 1 To lngBegin Step -1 '���뷴��
            '���ø�ҩ;����
            If Val(.TextMatrix(i, COL_ִ������)) = 5 And intִ������ <> 5 Then
                strִ������ = "�Ա�ҩ"
            ElseIf Val(.TextMatrix(i, COL_ִ������)) <> 5 And intִ������ = 5 Then
                strִ������ = "��Ժ��ҩ"
            Else
                strִ������ = ""
            End If
            .TextMatrix(i, COL_���ID) = "" '���������Ϊ��־
            .TextMatrix(i, COL_���ID) = AdviceSet��ҩ;��(i, lng��ҩ;��ID, strִ������, str����)
        Next
        
        mblnRowChange = varTmp1: .Redraw = varTmp2
        mblnNoSave = True '���Ϊδ����
    End With
End Sub

Private Function SaveAdvice() As Boolean
'���ܣ����浱ǰ���˵�ҽ����¼
    Dim dbl���� As Double, lng���ID
    Dim i As Long, j As Long
    
    With vsAdvice
        mlngNextID = 0
        Call InitSchemeRecordset

        '�������Ϊ˳������
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                .RowData(i) = -1 * .RowData(i)
                If Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                    .TextMatrix(i, COL_���ID) = -1 * Val(.TextMatrix(i, COL_���ID))
                End If
            End If
        Next
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                lng���ID = .RowData(i)
                .RowData(i) = GetNextID
                For j = i - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                        .TextMatrix(j, COL_���ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                        .TextMatrix(j, COL_���ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
            End If
        Next
        
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                '����ת��
                dbl���� = 0
                If Val(.TextMatrix(i, COL_����)) <> 0 Then
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                        '��ҩת�������۵�λ
                        dbl���� = Format(Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_��װϵ��)), "0.00000")
                    Else
                        '��ҩ�䷽�������ҩ��������,��ת��
                        dbl���� = Val(.TextMatrix(i, COL_����))
                    End If
                End If
                
                mrsScheme.AddNew
                mrsScheme!��� = Val(.RowData(i))
                mrsScheme!������ = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Null, Val(.TextMatrix(i, COL_���ID)))
                mrsScheme!��Ч = IIF(.TextMatrix(i, COL_��Ч) = "����", 0, 1)
                mrsScheme!������ĿID = IIF(Val(.TextMatrix(i, COL_������ĿID)) = 0, Null, Val(.TextMatrix(i, COL_������ĿID)))
                mrsScheme!�շ�ϸĿID = IIF(Val(.TextMatrix(i, COL_�շ�ϸĿID)) = 0, Null, Val(.TextMatrix(i, COL_�շ�ϸĿID)))
                mrsScheme!ҽ������ = IIF(.TextMatrix(i, col_ҽ������) = "", Null, .TextMatrix(i, col_ҽ������))
                mrsScheme!���� = IIF(Val(.TextMatrix(i, COL_����)) = 0, Null, Val(.TextMatrix(i, COL_����)))
                mrsScheme!�������� = IIF(Val(.TextMatrix(i, COL_����)) = 0, Null, Val(.TextMatrix(i, COL_����)))
                mrsScheme!�ܸ����� = IIF(dbl���� = 0, Null, dbl����)
                mrsScheme!ҽ������ = IIF(.TextMatrix(i, COL_ҽ������) = "", Null, .TextMatrix(i, COL_ҽ������))
                mrsScheme!ִ��Ƶ�� = IIF(.TextMatrix(i, COL_Ƶ��) = "", Null, .TextMatrix(i, COL_Ƶ��))
                mrsScheme!Ƶ�ʴ��� = Val(.TextMatrix(i, COL_Ƶ�ʴ���))
                mrsScheme!Ƶ�ʼ�� = Val(.TextMatrix(i, COL_Ƶ�ʼ��))
                mrsScheme!�����λ = IIF(.TextMatrix(i, COL_�����λ) = "", Null, .TextMatrix(i, COL_�����λ))
                mrsScheme!ʱ�䷽�� = IIF(.TextMatrix(i, COL_ִ��ʱ��) = "", Null, .TextMatrix(i, COL_ִ��ʱ��))
                mrsScheme!ִ�п���ID = IIF(Val(.TextMatrix(i, COL_ִ�п���ID)) = 0, Null, Val(.TextMatrix(i, COL_ִ�п���ID)))
                mrsScheme!ִ������ = Val(.TextMatrix(i, COL_ִ������))
                mrsScheme!�걾��λ = IIF(.TextMatrix(i, COL_�걾��λ) = "", Null, .TextMatrix(i, COL_�걾��λ))
                mrsScheme!��鷽�� = IIF(.TextMatrix(i, COL_��鷽��) = "", Null, .TextMatrix(i, COL_��鷽��))
                mrsScheme!�Ƿ�ȱʡ = IIF(Val(.TextMatrix(i, col_ȱʡ)) = -1, 1, 0)
                mrsScheme!�Ƿ�ѡ = IIF(Val(.TextMatrix(i, col_��ѡ)) = -1, 1, 0)
                mrsScheme!�䷽ID = .TextMatrix(i, COL_�䷽ID)
                mrsScheme!�����ĿID = .TextMatrix(i, COL_�����ĿID)
                mrsScheme!ִ�б�� = Val(.TextMatrix(i, COL_ִ�б��))
                If mbyt���� = 1 Then
                    mrsScheme!��� = .TextMatrix(i, COL_���)
                    mrsScheme!�������� = .TextMatrix(i, COL_��������)
                End If
                mrsScheme.Update
            End If
        Next
        
        If mrsScheme.RecordCount > 0 Then mrsScheme.MoveFirst
    End With
    
    mblnNoSave = False
    SaveAdvice = True
    mblnOK = True
End Function

Private Function CheckAdvice() As Boolean
'���ܣ���鵱ǰ����(Ӥ��)��ҽ�������Ƿ�Ϸ�
'˵��������в��Ϸ��ĵط����ڱ���������ʾ����λ
    Dim blnValid As Boolean
    Dim bln�䷽�� As Boolean, bln������ As Boolean
    Dim dbl���� As Double, strMsg As String
    Dim blnSkipTotal As Boolean, lngRow As Long, i As Long, j As Long
    Dim vMsg As VbMsgBoxResult, sng���� As Single
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        'Ϊ�������⣬��һ��ҽ���ġ�ȱʡ��������Ϊ��ͬ��ֵ
        For i = .FixedRows To .Rows - 1
            Call GetRowScope(i, lngBegin, lngEnd)
            For j = lngBegin To lngEnd
                If .TextMatrix(j, col_ȱʡ) <> .TextMatrix(i, col_ȱʡ) Then
                    .TextMatrix(j, col_ȱʡ) = .TextMatrix(i, col_ȱʡ)
                End If
            Next
            i = lngEnd + 1
        Next
    
    
        For i = .FixedRows To .Rows - 1
            '��������Ϸ��Լ��
            If .RowData(i) <> 0 And Not .RowHidden(i) Then
                bln�䷽�� = RowIn�䷽��(i)
                bln������ = RowIn������(i)
                lngRow = i
                If bln�䷽�� Then '�õ��䷽�ĵ�һҩƷ��
                    lngRow = .FindRow(CStr(.RowData(i)), , COL_���ID)
                ElseIf bln������ Then '�õ�����ҽ����
                    lngRow = .FindRow(CStr(.RowData(i)), , COL_���ID)
                End If
                
                If Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                    '��������ж�(�ٴ�·�����壬���׷������壬�����Ȳ�ȷ�����ֻȷ����Ʒ��)
'                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
'                        If .TextMatrix(i, COL_��Ч) = "����" And Val(.TextMatrix(i, COL_�շ�ϸĿID)) = 0 Then
'                            strMsg = "û�ж�Ӧ��ҩƷ�����Ϣ��"
'                            .Col = COL_ҽ������: Exit For
'                        End If
'                    End If
                    
                    '����¼��Ϸ���
                    If .TextMatrix(i, COL_����) <> "" Then
                        If .TextMatrix(i, COL_��Ч) = "����" Then
                            '��������ҩ���ʱ,������Ŀ��Ҫ¼��
                            If InStr(",1,2,", Val(.TextMatrix(i, COL_���㷽ʽ))) > 0 Or InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                                If Not IsNumeric(.TextMatrix(i, COL_����)) Or Val(.TextMatrix(i, COL_����)) <= 0 Then
                                    strMsg = "û��¼����ȷ�ĵ���������"
                                    .Col = COL_����: Exit For
                                End If
                            End If
                        Else
                            '����:��ҩ���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����¼��(Ҳ�ɲ�¼)
                            If Val(.TextMatrix(i, COL_Ƶ������)) = 0 And InStr(",1,2,", Val(.TextMatrix(i, COL_���㷽ʽ))) > 0 Then
                                If .TextMatrix(i, COL_����) <> "" Then
                                    If Not IsNumeric(.TextMatrix(i, COL_����)) Or Val(.TextMatrix(i, COL_����)) <= 0 Then
                                        strMsg = "û��¼����ȷ�ĵ���������"
                                        .Col = COL_����: Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '����¼��Ϸ���:�䷽,����(ҩƷ������)
                    If .TextMatrix(i, COL_����) <> "" Then
                        If .TextMatrix(i, COL_��Ч) = "����" Then
                            If Not IsNumeric(.TextMatrix(i, COL_����)) Or Val(.TextMatrix(i, COL_����)) <= 0 Then
                                If bln�䷽�� Then
                                    strMsg = "û��¼����ȷ����ҩ�䷽������"
                                ElseIf InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                                    strMsg = "û��¼����ȷ��ҩƷ�ܸ�������"
                                Else
                                    strMsg = "û��¼����ȷ��������"
                                End If
                                .Col = COL_����: Exit For
                            End If
                        End If
                    End If
                End If
                
                '�����������޸ĵ���
                '---------------------------------------------------
                If Val(.TextMatrix(i, COL_������ĿID)) = 0 Then
                    If .TextMatrix(i, col_ҽ������) = "" Then
                        strMsg = "û��¼��ҽ�����ݡ�"
                        .Col = COL_�÷�: Exit For
                    End If
                Else
                    '��ҩ;������ҩ�÷����ɼ��������ü��
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(i + 1) And Val(.TextMatrix(i + 1, COL_������ĿID)) = 0 Then
                            strMsg = "û�����ö�Ӧ�ĸ�ҩ;����"
                            .Col = COL_�÷�: Exit For
                        End If
                    End If
                    If .TextMatrix(i, COL_���) = "E" And Val(.TextMatrix(i, COL_������ĿID)) = 0 Then
                        If .RowData(i) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                            If InStr(",7,E,", .TextMatrix(i - 1, COL_���)) > 0 Then
                                strMsg = "��ҩ�䷽û�����ö�Ӧ���÷���"
                            ElseIf .TextMatrix(i - 1, COL_���) = "C" Then
                                strMsg = "û�����ö�Ӧ�ı걾�ɼ�������"
                            End If
                            .Col = COL_�÷�: Exit For
                        End If
                    End If
                    
                    '�����������:����Ҫ����һ��Ƶ�����ڵ�����
                    If Val(.TextMatrix(i, COL_����)) <> 0 And .TextMatrix(i, COL_��Ч) = "����" And (InStr(",4,5,6,", .TextMatrix(i, COL_���)) > 0 Or bln�䷽��) Then
                        If Not blnSkipTotal And .TextMatrix(i, COL_Ƶ��) <> "" Then
                            strMsg = ""
                            If bln�䷽�� Then '�ж�
                                dbl���� = CalcȱʡҩƷ����(1, 1, Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ))
                                If Val(.TextMatrix(i, COL_����)) < dbl���� Then
                                    strMsg = .TextMatrix(i, col_ҽ������) & vbCrLf & vbCrLf & _
                                        "�ڰ�""" & .TextMatrix(i, COL_Ƶ��) & """ִ��ʱ,������Ҫ " & dbl���� & "����"
                                End If
                            ElseIf Val(.TextMatrix(i, COL_����ϵ��)) <> 0 And Val(.TextMatrix(i, COL_����)) <> 0 Then
                                If Val(.TextMatrix(i, COL_Ƶ������)) = 1 Then '������ҩ����Ϊһ����
                                    dbl���� = CalcȱʡҩƷ����(Val(.TextMatrix(i, COL_����)), 1, 1, 1, "��", "", Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_��װϵ��)), Val(.TextMatrix(i, COL_�ɷ����)))
                                Else
                                    sng���� = Val(.TextMatrix(i, COL_����))
                                    If sng���� = 0 Then sng���� = 1
                                    dbl���� = CalcȱʡҩƷ����(Val(.TextMatrix(i, COL_����)), sng����, Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ), .TextMatrix(i, COL_ִ��ʱ��), Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_��װϵ��)), Val(.TextMatrix(i, COL_�ɷ����)))
                                End If
                                If Val(.TextMatrix(i, COL_����)) < dbl���� Then
                                    strMsg = .TextMatrix(i, col_ҽ������) & vbCrLf & vbCrLf & _
                                        "�ڰ�ÿ�� " & .TextMatrix(i, COL_����) & .TextMatrix(i, COL_������λ) & "," & .TextMatrix(i, COL_Ƶ��) & _
                                        IIF(Val(.TextMatrix(i, COL_Ƶ������)) <> 1 And Val(.TextMatrix(i, COL_����)) > 0 And .TextMatrix(i, COL_���) <> "4", ",��ҩ " & sng���� & " ��", "") & _
                                        "ִ��ʱ,������Ҫ " & dbl���� & .TextMatrix(i, COL_������λ) & "��"
                                End If
                            End If
                            If strMsg <> "" Then '��ʾ
                                .Row = i: .Col = COL_����: Call .ShowCell(.Row, .Col)
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^Ҫ������", Me)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                                    Exit Function
                                ElseIf vMsg = vbIgnore Then
                                    blnSkipTotal = True
                                End If
                            End If
                        End If
                    End If
                    
                    'ִ��ʱ��Ϸ��Լ��
                    If .TextMatrix(i, COL_ִ��ʱ��) <> "" And .TextMatrix(i, COL_Ƶ��) <> "" Then
                        blnValid = ExeTimeValid(.TextMatrix(i, COL_ִ��ʱ��), Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ))
                        If Not blnValid Then
                            If .TextMatrix(i, COL_�����λ) = "��" Then
                                strMsg = COL_����ִ��
                            ElseIf .TextMatrix(i, COL_�����λ) = "��" Then
                                strMsg = COL_����ִ��
                            ElseIf .TextMatrix(i, COL_�����λ) = "Сʱ" Then
                                strMsg = COL_��ʱִ��
                            End If
                            strMsg = "¼���ִ��ʱ�䷽����ʽ����ȷ�����顣" & vbCrLf & vbCrLf & "����" & vbCrLf & strMsg
                            .Col = COL_ִ��ʱ��: Exit For
                        End If
                    End If
                End If
            End If
        Next
        
        '--------------------------------------------------------------------------
        '�м��˳��Ĵ�����ʾ
        If i <= .Rows - 1 Then
            .Row = i: Call .ShowCell(.Row, .Col)
            If strMsg <> "" Then
                If bln�䷽�� Then
                    strMsg = "����ҩ�䷽" & strMsg
                Else
                    strMsg = """" & .TextMatrix(i, col_ҽ������) & """" & strMsg
                End If
                MsgBox strMsg, vbInformation, gstrSysName
                .Refresh
            End If
            If .Col = col_ҽ������ Then
                If txtҽ������.Enabled Then txtҽ������.SetFocus
            Else
                Call vsAdvice_KeyPress(13)
            End If
            Exit Function
        End If
        
        'û������
        If lngRow = 0 Then
            MsgBox "���׷�����û�����ݣ�����¼����׷������ݣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckAdvice = True
End Function

Private Function SeekNextControl() As Boolean
'���ܣ���λ����һ������Ŀؼ���,��������������Ƿ��Զ�����һ��ҽ��
'���أ����ͨ��SetFocusǿ�ƶ�λ��,�򷵻�True
    Dim objActive As Object, objNext As Object
    Dim blnDo As Boolean, i As Long
    Dim strSkip As String
    
    Set objActive = Me.ActiveControl
    
    If Not objActive Is Nothing Then
        If TypeName(objActive) = "TextBox" Or TypeName(objActive) = "ComboBox" Then
            If objActive.Container Is fraAdvice Then
                strSkip = GetInputSkip(vsAdvice.Row)
                Set objNext = zlControl.GetNextControl(objActive.TabIndex, Me, strSkip)
                If Not objNext Is Nothing Then
                    If objNext Is vsAdvice Then
                        For i = vsAdvice.Row + 1 To vsAdvice.Rows - 1
                            If Not vsAdvice.RowHidden(i) Then
                                Call AdviceChange 'ǿ�Ƹ���ҽ������
                                vsAdvice.Row = i
                                Call zlCommFun.PressKey(vbKeyTab)
                                blnDo = vsAdvice.RowData(i) <> 0 '��������������༭
                                Exit For
                            End If
                        Next
                        If i > vsAdvice.Rows - 1 Then
                            blnDo = True
                            If mbyt���� = 2 Then Exit Function '�����滻ʱֻ��������һ���滻����Ŀ��
                            cbsMain.FindControl(, conMenu_New, True, True).Execute
                        End If
                    ElseIf strSkip <> "" And InStr(";" & strSkip & ";", objNext.Name) = 0 Then
                        blnDo = True: objNext.SetFocus
                    End If
                End If
            End If
        End If
    End If
    If Not blnDo Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        SeekNextControl = True
    End If
End Function

Private Function GetInputSkip(ByVal lngRow As Long) As String
'���ܣ���ȡ����ҽ�������У��س����Ӧ�����Ŀؼ�
    Dim strSkip As String, lngFind As Long
    
    With vsAdvice
        'һ����ҩ�е�ҩƷ����ʱӦ����������
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 And .RowData(lngRow) <> 0 Then
            If Val(.TextMatrix(lngRow, COL_���ID)) = Val(.TextMatrix(lngRow - 1, COL_���ID)) Then
                '��ҩ;��,����ִ��
                If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                    lngFind = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    If lngFind <> -1 Then
                        If Val(.TextMatrix(lngFind, COL_������ĿID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.txt�÷�.Name
                        End If
                        If Val(.TextMatrix(lngFind, COL_ִ�п���ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.cbo����ִ��.Name
                        End If
                    End If
                End If
                'Ƶ��
                If .TextMatrix(lngRow, COL_Ƶ��) <> "" Then strSkip = strSkip & ";" & Me.txtƵ��.Name
                'ִ��ʱ��
                If .TextMatrix(lngRow, COL_ִ��ʱ��) <> "" Then strSkip = strSkip & ";" & Me.cboִ��ʱ��.Name
            End If
        End If
    End With
    GetInputSkip = Mid(strSkip, 2)
End Function

Private Function AdviceTextChange(ByVal lngRow As Long) As Boolean
'���ܣ���ҽ����Ƭ�������ݱ仯ʱ���ж�ҽ�������ı��Ƿ�Ӧ��������֯
    Dim str��� As String, strText As String, blnDefine As Boolean
    
    With vsAdvice
        'ȷ��ҽ�����
        str��� = .TextMatrix(lngRow, COL_���)
        If str��� = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then '��ҩ�䷽��һ�����
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            If lngRow <> -1 Then str��� = .TextMatrix(lngRow, COL_���)
        End If
        If str��� = "7" Then str��� = "8"
                
        'ȷ���Ƿ���
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "�������='" & str��� & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(NVL(mrsDefine!ҽ������)) = "" Then
                blnDefine = False
            End If
        End If
        If blnDefine Then strText = mrsDefine!ҽ������
        
        '������ݱ䶯
        If blnDefine Then '�����ֶβ��ݻ���Թ�������Ĳ���
            If cboҽ������.Tag <> "" And InStr(strText, "[ҽ������]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                If InStr(strText, "[����Ƶ��]") > 0 Or InStr(strText, "[Ӣ��Ƶ��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
            If cboִ��ʱ��.Tag <> "" And InStr(strText, "[ִ��ʱ��]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If (IsNumeric(txt����.Text) Or txt����.Text = "") And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If IsNumeric(txt����.Text) And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
        End If
        
        Select Case str��� '��ͬ�������
        Case "5", "6" '������ҩ
            If Not blnDefine Then
                
            Else
                '[������][ͨ����][��Ʒ��][Ӣ����][���][����]��������޸�����ҩƷʱ�仯
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[��ҩ;��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "8" '��ҩ�䷽
            If Not blnDefine Then
                If IsNumeric(txt����.Text) And txt����.Tag <> "" Then AdviceTextChange = True: Exit Function
                If cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then AdviceTextChange = True: Exit Function
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[�䷽���][�巨]��������޸������䷽ʱ�仯
                If IsNumeric(txt����.Text) And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[�÷�]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "C" '����
            If Not blnDefine Then
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[������Ŀ][����걾]��������޸�������Ŀʱ�仯
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[�ɼ�����]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "D" '���
            If Not blnDefine Then
                
            Else
                '[�����Ŀ][��鲿λ]��������޸�������Ŀʱ�仯
            End If
        Case "F" '����
            If Not blnDefine Then
            Else
                '[��Ҫ����][��������][������]��������޸�������Ŀʱ�仯
            End If
        Case "K" '��Ѫ
            If Not blnDefine Then
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[��Ѫ;��]
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[��Ѫ;��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case Else '����
            If Not blnDefine Then
                
            Else
                '[������Ŀ]��������޸�������Ŀʱ�仯
            End If
        End Select
    End With
End Function

Private Function AdviceTextMake(ByVal lngRow As Long) As String
'���ܣ���ȡҽ�������ı�
'������lngRow=����ҽ�����ݵĿɼ���
    Dim rsTmp As New ADODB.Recordset
    Dim blnDefine As Boolean, str��� As String
    Dim strText As String, strSql As String
    Dim strField As String, intƵ�ʷ�Χ As Integer
    Dim i As Long, k As Long
    
    Dim str��ҩ As String, str�巨 As String, str��̬ As String
    Dim str���� As String, str���� As String
    Dim str���� As String, str�걾 As String
    Dim str��λ As String, str��λLast As String, str���� As String
    Dim dbl���� As Double
    Dim blnDo As Boolean
    Dim str��ҩ���� As String
    
    On Error GoTo errH
    
    With vsAdvice
        'ȷ��ҽ�����
        str��� = .TextMatrix(lngRow, COL_���)
        If str��� = "E" Then '��ҩ�䷽��һ�����
            k = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            If k <> -1 Then str��� = .TextMatrix(k, COL_���)
        End If
        If str��� = "7" Then str��� = "8"
                
        'ȷ���Ƿ���
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "�������='" & str��� & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(NVL(mrsDefine!ҽ������)) = "" Then
                blnDefine = False
            End If
        End If
        
ReDoDefault: '���ڰ����幫ʽ����ʧ�ܣ����°�ȱʡ���������֯
        strText = ""
        If blnDefine Then strText = mrsDefine!ҽ������
        
        '����ҽ������
        Select Case str���
        Case "C" '����-------------------------------------------------------------
            str���� = "": str�걾 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_�����ĿID)) = 0 And mblnNewLIS Or Not mblnNewLIS Then
                        str���� = .TextMatrix(i, col_ҽ������) & "," & str����
                    End If
                    str�걾 = .TextMatrix(i, COL_�걾��λ)
                Else
                    Exit For
                End If
            Next
            If str���� = "" Then '�ϵķ�ʽ
                str���� = .TextMatrix(lngRow, COL_����)
            Else
                str���� = Left(str����, Len(str����) - 1)
            End If
            
            If Not blnDefine Then
                strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
            Else
                If InStr(strText, "[������Ŀ]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[������Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[����걾]") > 0 Then
                    strField = str�걾
                    strText = Replace(strText, "[����걾]", """" & strField & """")
                End If
                If InStr(strText, "[�ɼ�����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[�ɼ�����]", """" & strField & """")
                End If
            End If
        Case "D" '���-------------------------------------------------------------
            str��λ = "": str���� = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_�걾��λ) <> "" Then
                        If .TextMatrix(i, COL_�걾��λ) <> str��λLast And str��λLast <> "" Then
                            str��λ = str��λ & "," & str��λLast & IIF(str���� <> "", "(" & Mid(str����, 2) & ")", "")
                            str���� = ""
                        End If
                        If .TextMatrix(i, COL_��鷽��) <> "" Then
                            str���� = str���� & "," & .TextMatrix(i, COL_��鷽��)
                        End If
                        
                        str��λLast = .TextMatrix(i, COL_�걾��λ)
                    End If
                Else
                    Exit For
                End If
            Next
            If str��λLast <> "" Then
                str��λ = str��λ & "," & str��λLast & IIF(str���� <> "", "(" & Mid(str����, 2) & ")", "")
            End If
            str��λ = Mid(str��λ, 2) '��������Ŀ�Ĳ�λ
            
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_����) & IIF(str��λ <> "", ":" & str��λ, "")
            Else
                If InStr(strText, "[�����Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[�����Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[��鲿λ]") > 0 Then
                    strField = str��λ
                    strText = Replace(strText, "[��鲿λ]", """" & strField & """")
                End If
            End If
        Case "F" '����-------------------------------------------------------------
            str���� = "": str���� = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "G" Then
                        str���� = .TextMatrix(i, col_ҽ������)
                    Else
                        str���� = str���� & "," & .TextMatrix(i, col_ҽ������)
                    End If
                Else
                    Exit For
                End If
            Next
            str���� = Mid(str����, 2)
            
            If Not blnDefine Then
                strText = ""
                If str���� <> "" Then
                    strText = strText & IIF(str���� <> "", " �� " & str���� & " ���� ", " �� ")
                End If
                strText = strText & .TextMatrix(lngRow, COL_����)
                If str���� <> "" Then
                    strText = strText & " �� " & str����
                End If
            Else
                If InStr(strText, "[��Ҫ����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[��Ҫ����]", """" & strField & """")
                End If
                If InStr(strText, "[��������]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[��������]", """" & strField & """")
                End If
                If InStr(strText, "[������]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[������]", """" & strField & """")
                End If
            End If
        Case "8" '��ҩ�䷽---------------------------------------------------------
            str��ҩ = "": str�巨 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" Then
                        dbl���� = dbl���� + Val(.TextMatrix(i, COL_����))
                        If Val(.TextMatrix(lngRow, COL_��ҩ��̬)) = 0 Then
                            blnDo = .TextMatrix(i, COL_�շ�ϸĿID) <> .TextMatrix(i - 1, COL_�շ�ϸĿID)
                        Else
                            blnDo = .TextMatrix(i, COL_������ĿID) <> .TextMatrix(i - 1, COL_������ĿID)
                        End If
                        
                        If blnDo Then
                            str��ҩ���� = .TextMatrix(i, col_ҽ������)
                            
                            If Val(.TextMatrix(lngRow, COL_��ҩ��̬)) = 0 Then
                                strSql = "Select ��� as ���� From �շ���ĿĿ¼ Where ID=[1] And Exists(Select 1 From ҩƷ��� Where ҩƷID<>[1] And ҩ��ID=[2])"
                                Set rsTmp = New ADODB.Recordset '���Filter
                                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(i, COL_�շ�ϸĿID)), Val(.TextMatrix(i, COL_������ĿID)))
                                If rsTmp.RecordCount > 0 Then
                                    If Not IsNull(rsTmp!����) Then str��ҩ���� = str��ҩ���� & "(" & rsTmp!���� & ")"
                                End If
                            End If
                        
                            str��ҩ = RTrim(str��ҩ���� & _
                                " " & FormatEx(dbl����, 5) & .TextMatrix(i, COL_������λ) & _
                                " " & .TextMatrix(i, COL_ҽ������)) & "," & str��ҩ
                            dbl���� = 0
                        End If
                    ElseIf .TextMatrix(i, COL_���) = "E" Then
                        str�巨 = .TextMatrix(i, col_ҽ������) & .TextMatrix(i, COL_�걾��λ)
                    End If
                Else
                    Exit For
                End If
            Next
            If str��ҩ <> "" Then
                str��ҩ = Mid(str��ҩ, 1, Len(str��ҩ) - 1)
            End If
            If Not blnDefine Or .TextMatrix(lngRow, COL_��Ч) = "����" Then
                If .TextMatrix(lngRow, COL_��ҩ��̬) = "1" Then
                    str��̬ = "[��Ƭ]"
                ElseIf .TextMatrix(lngRow, COL_��ҩ��̬) = "2" Then
                    str��̬ = "[����]"
                End If
                '���ֺ���˿ո����ı����л��Զ�����
                If .TextMatrix(lngRow, COL_��Ч) = "����" Then
                    '�����䷽���ݸ������ô������ù̶�����
                    strText = "��ҩ�䷽" & str��̬ & "," & _
                        .TextMatrix(lngRow, COL_Ƶ��) & "," & str�巨 & "," & _
                        .TextMatrix(lngRow, COL_�÷�) & ":" & str��ҩ
                Else
                    strText = "��ҩ" & str��̬ & .TextMatrix(lngRow, COL_����) & "��," & _
                        .TextMatrix(lngRow, COL_Ƶ��) & "," & str�巨 & "," & _
                        .TextMatrix(lngRow, COL_�÷�) & ":" & str��ҩ
                End If
            Else
                If InStr(strText, "[����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[����]", """" & strField & """")
                End If
                If InStr(strText, "[�䷽���]") > 0 Then
                    strField = str��ҩ
                    strText = Replace(strText, "[�䷽���]", """" & strField & """")
                End If
                If InStr(strText, "[�÷�]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[�÷�]", """" & strField & """")
                End If
                If InStr(strText, "[�巨]") > 0 Then
                    strField = str�巨
                    strText = Replace(strText, "[�巨]", """" & strField & """")
                End If
            End If
        Case "4" '����------------------------------------------------------------
            If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                strSql = "Select ����,���,���� From �շ���ĿĿ¼ Where ID=[1]"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)))
            ElseIf blnDefine Then
                strSql = "Select ����,NULL As ���,NULL As ���� From ������ĿĿ¼ Where ID=[1]"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_������ĿID)))
            End If
                If Not blnDefine Then
                    strText = .TextMatrix(lngRow, COL_����)
                    If Not IsNull(rsTmp!���) Then
                        strText = strText & " " & rsTmp!���
                    End If
                Else
                    If InStr(strText, "[��������]") > 0 Then
                        strField = rsTmp!����
                        strText = Replace(strText, "[��������]", """" & strField & """")
                    End If
                    If InStr(strText, "[���]") > 0 Then
                        strField = NVL(rsTmp!���)
                        strText = Replace(strText, "[���]", """" & strField & """")
                    End If
                    If InStr(strText, "[����]") > 0 Then
                        strField = NVL(rsTmp!����)
                        strText = Replace(strText, "[����]", """" & strField & """")
                    End If
                End If
        Case "5", "6" '����ҩ���г�ҩ---------------------------------------------
            If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                '����:0-����,1-Ӣ����,3-��Ʒ��
                strSql = "Select Nvl(B.����,A.����) as ����,A.���,A.����,B.����" & _
                    " From �շ���ĿĿ¼ A,�շ���Ŀ���� B Where A.ID=B.�շ�ϸĿID(+) And A.ID=[1] Order by B.����,B.����"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)))
            ElseIf blnDefine Then
                '����:0-����,1-Ӣ����
                strSql = "Select Nvl(B.����,A.����) as ����,Null as ���,Null as ����,B.����" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B Where A.ID=B.������ĿID(+) And A.ID=[1] Order by B.����,B.����"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_������ĿID)))
            End If
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_�걾��λ)
                If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                    If strText = "" Then
                        If gbytҩƷ������ʾ <> 0 Then rsTmp.Filter = "����=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strText = rsTmp!����
                    End If
                    If Not IsNull(rsTmp!����) Then
                        strText = strText & "(" & rsTmp!���� & ")"
                    End If
                    If Not IsNull(rsTmp!���) Then
                        strText = strText & " " & rsTmp!���
                    End If
                Else
                    If strText = "" Then
                        strText = .TextMatrix(lngRow, COL_����)
                    End If
                End If
            Else
                If InStr(strText, "[������]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�걾��λ)
                    If strField = "" Then
                        If gbytҩƷ������ʾ <> 0 Then rsTmp.Filter = "����=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strField = rsTmp!����
                    End If
                    strText = Replace(strText, "[������]", """" & strField & """")
                End If
                If InStr(strText, "[ͨ����]") > 0 Then
                    rsTmp.Filter = 0
                    strField = rsTmp!����
                    strText = Replace(strText, "[ͨ����]", """" & strField & """")
                End If
                If InStr(strText, "[��Ʒ��]") > 0 Then
                    rsTmp.Filter = "����=3"
                    If rsTmp.EOF Then
                        strField = ""
                    Else
                        strField = rsTmp!����
                    End If
                    strText = Replace(strText, "[��Ʒ��]", """" & strField & """")
                End If
                If InStr(strText, "[Ӣ����]") > 0 Then
                    rsTmp.Filter = "����=2"
                    If rsTmp.EOF Then
                        strField = ""
                    Else
                        strField = rsTmp!����
                    End If
                    strText = Replace(strText, "[Ӣ����]", """" & strField & """")
                End If
                If InStr(strText, "[���]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = NVL(rsTmp!���)
                    strText = Replace(strText, "[���]", """" & strField & """")
                End If
                If InStr(strText, "[����]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = NVL(rsTmp!����)
                    strText = Replace(strText, "[����]", """" & strField & """")
                End If
                If InStr(strText, "[��ҩ;��]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[��ҩ;��]", """" & strField & """")
                End If
            End If
        Case "K" '��Ѫҽ��
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_����)
                If .TextMatrix(lngRow, COL_�÷�) <> "" Then
                    strText = strText & "(" & .TextMatrix(lngRow, COL_�÷�) & ")"
                End If
            Else
                If InStr(strText, "[������Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[������Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[��Ѫ��Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[��Ѫ��Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[��Ѫ;��]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[��Ѫ;��]", """" & strField & """")
                End If
            End If
        Case Else '�����������-----------------------------------------------------
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_����)
            Else
                If InStr(strText, "[������Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[������Ŀ]", """" & strField & """")
                End If
            End If
            '����ҽ��������ʾ
            If .TextMatrix(lngRow, COL_���) = "Z" And (Val(.TextMatrix(lngRow, COL_��������)) = 4 Or Val(.TextMatrix(lngRow, COL_��������)) = 14) Then
                strText = "������" & strText & "������"
            End If
            'ת��ҽ��������ʾ
            If .TextMatrix(lngRow, COL_���) = "Z" And Val(.TextMatrix(lngRow, COL_��������)) = 3 Then
                strText = "������" & strText & "������"
            End If
        End Select
        
        '�����ֶλ���Թ���������ֶ�-------------------------------------------
        If blnDefine Then
            If InStr(strText, "[ҽ������]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_ҽ������)
                strText = Replace(strText, "[ҽ������]", """" & strField & """")
            End If
            If InStr(strText, "[����Ƶ��]") > 0 Then
                strField = .TextMatrix(lngRow, COL_Ƶ��)
                strText = Replace(strText, "[����Ƶ��]", """" & strField & """")
            End If
            If InStr(strText, "[Ӣ��Ƶ��]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_Ƶ��) <> "" Then
                    intƵ�ʷ�Χ = GetƵ�ʷ�Χ(lngRow)
                    strSql = "Select Ӣ������ From ����Ƶ����Ŀ Where ����=[1] And ���÷�Χ=[2]"
                    Set rsTmp = New ADODB.Recordset '���Filter
                    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(lngRow, COL_Ƶ��), intƵ�ʷ�Χ)
                    If Not rsTmp.EOF Then strField = NVL(rsTmp!Ӣ������)
                End If
                strText = Replace(strText, "[Ӣ��Ƶ��]", """" & strField & """")
            End If
            If InStr(strText, "[����]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_����) <> "" Then
                    strField = .TextMatrix(lngRow, COL_����) & .TextMatrix(lngRow, COL_������λ)
                End If
                strText = Replace(strText, "[����]", """" & strField & """")
            End If
            If InStr(strText, "[����]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_����) <> "" Then
                    strField = .TextMatrix(lngRow, COL_����) & .TextMatrix(lngRow, COL_������λ)
                End If
                strText = Replace(strText, "[����]", """" & strField & """")
            End If
            If InStr(strText, "[ִ��ʱ��]") > 0 Then
                strField = .TextMatrix(lngRow, COL_ִ��ʱ��)
                strText = Replace(strText, "[ִ��ʱ��]", """" & strField & """")
            End If
        End If
                
        '����ҽ������
        If blnDefine Then
            On Error Resume Next
            strText = mobjVBA.Eval(strText)
            If mobjVBA.Error.Number <> 0 Then
                err.Clear: On Error GoTo errH
                blnDefine = False: GoTo ReDoDefault
            End If
        End If
    End With
    AdviceTextMake = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CanAlterType(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ����ҽ���Ƿ�����л���Ч
'������lngRow=�ɼ���ҽ����
'˵���������л���Ч��������
'   1.�ɳ�����ִ��Ƶ��=0(��ѡƵ��),2(������)
'   2.��������ִ��Ƶ��=0(��ѡƵ��),1(һ����);ҩƷ����ָ���˹��
    Dim rsMore As New ADODB.Recordset
    Dim strSql As String, strType As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            CanAlterType = True: Exit Function
        ElseIf Val(.TextMatrix(lngRow, COL_������ĿID)) = 0 Then
            '��������Ŀ����л�
            CanAlterType = True: Exit Function
        ElseIf RowIn�䷽��(lngRow) Then
            '��ҩ�䷽�̶������л�
            CanAlterType = True: Exit Function
        ElseIf RowIn������(lngRow) Then
            '�����Լ�����Ϊ׼�ж�
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            If lngRow = -1 Then Exit Function
        End If
    
        strType = IIF(.TextMatrix(lngRow, COL_��Ч) = "����", "����", "����")
        
        '��ԭʼƵ��Ϊ׼�ж�:��Ϊ��ѡ��Ƶ�ʵĿ�����ȱ��һ����
        strSql = "Select ִ��Ƶ�� From ������ĿĿ¼ Where ID=[1]"
        On Error GoTo errH
        Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(lngRow, COL_������ĿID)))
        
        If strType = "����" Then
            If InStr(",0,2,", NVL(rsMore!ִ��Ƶ��, 0)) = 0 Then Exit Function
        Else
            If InStr(",0,1,", NVL(rsMore!ִ��Ƶ��, 0)) = 0 Then Exit Function
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                Call GetRowScope(lngRow, lngBegin, lngEnd)
                For i = lngBegin To lngEnd
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                        If Val(.TextMatrix(i, COL_�շ�ϸĿID)) = 0 Then Exit Function
                    End If
                Next
            End If
        End If
    End With
    CanAlterType = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceAlterType(ByVal lngRow As Long)
'���ܣ��ھ����������ݵ�����£��л�ָ����ҽ������Ч(����<->��ʱ)
'������lngRow=�ɼ���ҽ����
'˵����ִ�иú���ʱӦ��֤����CanAlterType�����������ж�
    Dim rsMore As New ADODB.Recordset
    Dim strType As String, strSql As String
    Dim intƵ������ As Integer, sng���� As Single
    Dim strƵ�� As String, intƵ�ʴ��� As Integer
    Dim intƵ�ʼ�� As Integer, str�����λ As String
    Dim lng�÷�ID As Long, blnToNormal As Boolean
    Dim lngBegin As Long, lngEnd As Long
    Dim lngCopyRow As Long, i As Long
    
    On Error GoTo errH
    With vsAdvice
        '����Ҫת��Ϊ����Ч
        strType = IIF(.TextMatrix(lngRow, COL_��Ч) = "����", "����", "����")
        
        If Val(.TextMatrix(lngRow, COL_������ĿID)) <> 0 Then
            'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
            lngCopyRow = GetPreRow(lngRow)
            If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
            
            '��ȡһ��ҽ���Ĳ����з�Χ
            Call GetRowScope(lngRow, lngBegin, lngEnd)
        End If
        
        '��Բ�ͬ����ҽ������ת��-----------------------------------------
        If Val(.TextMatrix(lngRow, COL_������ĿID)) = 0 Then
            '����¼���ҽ��ֱ�Ӵ���
            .TextMatrix(lngRow, COL_��Ч) = strType
        ElseIf RowIn�䷽��(lngRow) Then '��ҩ�䷽
            'ҩƷ��������Ϊ��Ժ��ҩ
            If strType = "����" And .TextMatrix(lngEnd, COL_���) = "E" _
                And .RowData(lngEnd) = Val(.TextMatrix(lngBegin, COL_���ID)) Then
                If Val(.TextMatrix(lngBegin, COL_ִ������)) <> 5 And Val(.TextMatrix(lngEnd, COL_ִ������)) = 5 Then
                    lng�÷�ID = Val(.TextMatrix(lngEnd, COL_������ĿID))
                    blnToNormal = True '��ʾ��ҩִ��Ӧ�ָ�������ֵ
                End If
            End If
            
            For i = lngBegin To lngEnd
                '��Чֵ
                .TextMatrix(i, COL_��Ч) = strType

                '����
                If strType = "����" Then
                    .TextMatrix(i, COL_����) = ""
                End If
                '����ҽ��Ƶ��
                If .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ" Then
                    .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ"
                    txtƵ��.Text = "��Ҫʱ"
                    cmdƵ��.Tag = "��Ҫʱ"
                ElseIf .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ" Then
                    .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ"
                    txtƵ��.Text = "��Ҫʱ"
                    cmdƵ��.Tag = "��Ҫʱ"
                End If
                
                'ִ������:ҩƷ��������Ϊ"��Ժ��ҩ"
                If i = lngEnd And blnToNormal Then
                    strSql = "Select ִ�п��� From ������ĿĿ¼ Where ID=[1]"
                    Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng�÷�ID)
                    
                    .TextMatrix(i, COL_ִ������) = NVL(rsMore!ִ�п���, 0)
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                        .TextMatrix(i, COL_ִ�п���ID) = Get����ִ�п���ID("E", lng�÷�ID, 0, NVL(rsMore!ִ�п���, 0), IIF(strType = "����", 0, 1), mint��Χ)
                    Else
                        .TextMatrix(i, COL_ִ�п���ID) = 0
                    End If
                End If
            Next
        Else '���������Ŀ,����ҩƷ,����,���(���),����(���)�������������봦������ͬ,���һ����
            '��ȡ��ҩ;��ID
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 _
                And .TextMatrix(lngEnd, COL_���) = "E" And .RowData(lngEnd) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                lng�÷�ID = Val(.TextMatrix(lngEnd, COL_������ĿID))
                
                'ҩƷ��������Ϊ��Ժ��ҩ
                If strType = "����" Then
                    If Val(.TextMatrix(lngRow, COL_ִ������)) <> 5 And Val(.TextMatrix(lngEnd, COL_ִ������)) = 5 Then
                        blnToNormal = True '��ʾ��ҩִ��Ӧ�ָ�������ֵ
                    End If
                End If
            End If
            
            '------------------------------------------------------------------------------------------------------
            'ͬʱ����һ��ҽ���������
            For i = lngBegin To lngEnd
                '��Чֵ
                .TextMatrix(i, COL_��Ч) = strType
                
                '�ɳ���ҩ���л�Ϊ��ʱҩ��
                If .Cell(flexcpData, i, COL_�ɷ����) = 1 Then
                    .Cell(flexcpData, i, COL_�ɷ����) = Empty '����Ӧ��ԭ����ķ��㷽ʽ
                End If
                
                '��ȡ��ǰ��Ŀ�ĸ�����Ϣ
                If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 And i = lngBegin Then
                    '��һҩƷ�в�ȡ��Щ��Ϣ
                    strSql = "Select ��ĿID,Ƶ��,�Ƴ� From �����÷����� Where Nvl(����,0)>0 And ��ĿID=[1] And �÷�ID=[2]"
                    strSql = "Select A.ִ�п���,A.ִ��Ƶ��,A.���㷽ʽ,A.���㵥λ,B.Ƶ��,B.�Ƴ�" & _
                        " From ������ĿĿ¼ A,(" & strSql & ") B Where A.ID=B.��ĿID(+) And A.ID=[1]"
                Else
                    strSql = "Select ִ�п���,ִ��Ƶ��,���㷽ʽ,���㵥λ,Null as Ƶ��,Null as �Ƴ� From ������ĿĿ¼ Where ID=[1]"
                End If
                Set rsMore = zldatabase.OpenSQLRecord(strSql, Me.Caption, Val(.TextMatrix(i, COL_������ĿID)), lng�÷�ID)
                If Not rsMore.EOF Then '��ҩ;��û��ָ�������
                    '����(��λ)
                    If strType = "����" Then
                        If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                            '�С�����ҩ������������λ���ǰ�װ��λ
                            .TextMatrix(i, COL_������λ) = .TextMatrix(i, COL_��װ��λ)
                        ElseIf .TextMatrix(i, COL_���) = "4" Then
                            .TextMatrix(i, COL_������λ) = .TextMatrix(i, COL_��װ��λ) 'ɢװ��λ
                        Else
                            '��������Ҫ��������
                            .TextMatrix(i, COL_������λ) = NVL(rsMore!���㵥λ)
                            
                            '���Ϊһ���Ի�ƴ�����ȱʡ����Ϊ1
                            If i = lngBegin Then
                                If NVL(rsMore!ִ��Ƶ��, 0) = 1 Or NVL(rsMore!���㷽ʽ, 0) = 3 Then
                                    .TextMatrix(i, COL_����) = 1
                                End If
                            ElseIf Not (lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID))) Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngBegin, COL_����)
                            End If
                        End If
                    Else
                        .TextMatrix(i, COL_����) = ""
                        .TextMatrix(i, COL_������λ) = ""
                    End If
                    
                    '����ҽ��Ƶ��
                    If .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ" Then
                        .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ"
                        txtƵ��.Text = "��Ҫʱ"
                        cmdƵ��.Tag = "��Ҫʱ"
                    ElseIf .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ" Then
                        .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ"
                        txtƵ��.Text = "��Ҫʱ"
                        cmdƵ��.Tag = "��Ҫʱ"
                    Else
                
                        'Ƶ������,ִ��Ƶ��,ִ��ʱ��
                        If i = lngBegin Then '�Ե�һ��Ϊ׼
                            intƵ������ = Val(.TextMatrix(i, COL_Ƶ������))
                            If strType = "����" And NVL(rsMore!ִ��Ƶ��, 0) = 0 And mblnһ���� Then
                                .TextMatrix(i, COL_Ƶ������) = 1 '��ѡ��Ƶ�ʵ�����ȱʡΪһ����
                            Else
                                .TextMatrix(i, COL_Ƶ������) = NVL(rsMore!ִ��Ƶ��, 0)
                            End If
            
                            'ִ��Ƶ��:�����÷�Χ�����仯ʱ
                            If Val(.TextMatrix(i, COL_Ƶ������)) <> intƵ������ Then
                                '���Ϊ����ȡ
                                .TextMatrix(i, COL_Ƶ��) = ""
                                .TextMatrix(i, COL_ִ��ʱ��) = ""
                                
                                'ҩƷ���õ�ȱʡƵ������
                                If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 _
                                    And Not IsNull(rsMore!Ƶ��) And Val(.TextMatrix(i, COL_Ƶ������)) <> 1 Then
                                    Call GetƵ����Ϣ_����(rsMore!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                                    .TextMatrix(i, COL_Ƶ��) = strƵ��
                                    .TextMatrix(i, COL_Ƶ�ʴ���) = intƵ�ʴ���
                                    .TextMatrix(i, COL_Ƶ�ʼ��) = intƵ�ʼ��
                                    .TextMatrix(i, COL_�����λ) = str�����λ
                                End If
                                'ȱʡ����һ��������ͬ
                                If .TextMatrix(i, COL_Ƶ��) = "" And lngCopyRow <> -1 Then
                                    If .TextMatrix(i, COL_��Ч) = .TextMatrix(lngCopyRow, COL_��Ч) _
                                        And Val(.TextMatrix(i, COL_Ƶ������)) = Val(.TextMatrix(lngCopyRow, COL_Ƶ������)) Then
                                        If .TextMatrix(lngCopyRow, COL_Ƶ��) <> "" _
                                            And Not (.TextMatrix(i, COL_���) = "7" And Not RowIn�䷽��(lngCopyRow)) _
                                            And Not (.TextMatrix(i, COL_���) <> "7" And RowIn�䷽��(lngCopyRow)) _
                                            And CheckƵ�ʿ���(Val(.TextMatrix(i, COL_������ĿID)), GetƵ�ʷ�Χ(i), .TextMatrix(lngCopyRow, COL_Ƶ��)) Then
                                            .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��)
                                            .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngCopyRow, COL_Ƶ�ʴ���)
                                            .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngCopyRow, COL_Ƶ�ʼ��)
                                            .TextMatrix(i, COL_�����λ) = .TextMatrix(lngCopyRow, COL_�����λ)
                                        End If
                                    End If
                                End If
                                '��ȡȱʡƵ��
                                If .TextMatrix(i, COL_Ƶ��) = "" Then
                                    Call GetȱʡƵ��(Val(.TextMatrix(i, COL_������ĿID)), GetƵ�ʷ�Χ(i), strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                                    .TextMatrix(i, COL_Ƶ��) = strƵ��
                                    .TextMatrix(i, COL_Ƶ�ʴ���) = intƵ�ʴ���
                                    .TextMatrix(i, COL_Ƶ�ʼ��) = intƵ�ʼ��
                                    .TextMatrix(i, COL_�����λ) = str�����λ
                                End If
                                
                                'ִ��ʱ��:��ѡƵ�ʵ���Ŀ
                                If Val(.TextMatrix(i, COL_Ƶ������)) = 0 Then
                                    If lngCopyRow <> -1 Then '����һ����ͬ
                                        If .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��) Then
                                            .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngCopyRow, COL_ִ��ʱ��)
                                        End If
                                    End If
                                    If .TextMatrix(i, COL_ִ��ʱ��) = "" Then  'ȱʡʱ�䷽��
                                        .TextMatrix(i, COL_ִ��ʱ��) = Getȱʡʱ��(1, .TextMatrix(i, COL_Ƶ��), lng�÷�ID)
                                    End If
                                End If
                            End If
                        Else
                            .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngBegin, COL_Ƶ��)
                            .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngBegin, COL_Ƶ�ʴ���)
                            .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngBegin, COL_Ƶ�ʼ��)
                            .TextMatrix(i, COL_�����λ) = .TextMatrix(lngBegin, COL_�����λ)
                            .TextMatrix(i, COL_Ƶ������) = .TextMatrix(lngBegin, COL_Ƶ������)
                            .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngBegin, COL_ִ��ʱ��)
                        End If
                    End If
                    
                    'ҩƷ��������������
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 And strType = "����" Then
                        'ȷ��������ҩ������
                        '1.����Ϊһ��Ƶ����������
                        '2-���Ƴ���Ϊ�Ƴ�����(Ӧ����һ��Ƶ����������)
                        If i = lngBegin Then '�Ե�һ��Ϊ׼
                            sng���� = Val(.TextMatrix(i, COL_����)) '�����ǰ���򱣳�
                            If sng���� = 0 Then sng���� = msng����
                            
                            If .TextMatrix(i, COL_�����λ) = "��" Then
                                If 7 > sng���� Then sng���� = 7
                            ElseIf .TextMatrix(i, COL_�����λ) = "��" Then
                                If Val(.TextMatrix(i, COL_Ƶ�ʼ��)) > sng���� Then
                                    sng���� = Val(.TextMatrix(i, COL_Ƶ�ʼ��))
                                End If
                            ElseIf .TextMatrix(i, COL_�����λ) = "Сʱ" Then
                                If Val(.TextMatrix(i, COL_Ƶ�ʼ��)) \ 24 > sng���� Then
                                    sng���� = Val(.TextMatrix(i, COL_Ƶ�ʼ��)) \ 24
                                End If
                            ElseIf .TextMatrix(i, COL_�����λ) = "����" Then
                                If sng���� = 0 Then sng���� = 1
                            End If

                            If NVL(rsMore!�Ƴ�, 1) > sng���� Then sng���� = NVL(rsMore!�Ƴ�, 1)
                            If sng���� = 0 Then sng���� = 1
                        End If
                        
                        '����
                        If Val(.TextMatrix(i, COL_Ƶ������)) <> 1 Then
                            .TextMatrix(i, COL_����) = IIF(sng���� = 0, "", sng����)
                        End If
                        
                        '����
                        If .TextMatrix(i, COL_Ƶ��) <> "" And Val(.TextMatrix(i, COL_����)) <> 0 _
                            And Val(.TextMatrix(i, COL_����ϵ��)) <> 0 And Val(.TextMatrix(i, COL_��װϵ��)) <> 0 Then
                            If Val(.TextMatrix(i, COL_Ƶ������)) = 1 Then '����ҩƷ����ȱʡΪһ����
                                '�����Ƴ����Ϊ��������ҩ������
                                .TextMatrix(i, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                        Val(.TextMatrix(i, COL_����)), 1, 1, 1, "��", "", Val(.TextMatrix(i, COL_����ϵ��)), _
                                        Val(.TextMatrix(i, COL_��װϵ��)), Val(.TextMatrix(i, COL_�ɷ����))), 5)
                            Else
                                '�����Ƴ����Ϊ��������ҩ������
                                .TextMatrix(i, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                        Val(.TextMatrix(i, COL_����)), sng����, Val(.TextMatrix(i, COL_Ƶ�ʴ���)), _
                                        Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ), _
                                        .TextMatrix(i, COL_ִ��ʱ��), Val(.TextMatrix(i, COL_����ϵ��)), _
                                        Val(.TextMatrix(i, COL_��װϵ��)), Val(.TextMatrix(i, COL_�ɷ����))), 5)
                            End If
                        End If
                    End If
                    
                    'ִ������:ҩƷ��������Ϊ"��Ժ��ҩ"
                    If i = lngEnd And blnToNormal Then
                        .TextMatrix(i, COL_ִ������) = NVL(rsMore!ִ�п���, 0)
                        If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                            .TextMatrix(i, COL_ִ�п���ID) = Get����ִ�п���ID("E", lng�÷�ID, 0, NVL(rsMore!ִ�п���, 0), IIF(strType = "����", 0, 1), mint��Χ)
                        Else
                            .TextMatrix(i, COL_ִ�п���ID) = 0
                        End If
                    End If
                End If
            Next
        End If
    End With
    
    mblnNoSave = True '���Ϊδ����
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get����ִ�п���(objCbo As Object, ByVal str��� As String, ByVal lng��Ŀid As Long, ByVal lngҩƷID As Long, _
    ByVal intִ�п��� As Integer, ByVal lng��ǰִ��ID As Long, ByVal int��Ч As Integer, ByVal int��Χ As Integer) As Boolean
'���ܣ�����������Ŀִ�п�����Ϣ���ؿ��õ�ִ�п�����ָ����������
'������intִ�п���=��Ŀִ�п��ұ�־
'      lng��ǰִ��ID=ҽ����ǰ��ִ�п���ID
'      int��Ч=0-����,1-����,��������Ҫ�ж��ϰ�ʱ��
'      int��Χ=1-����,2-סԺ,3-�����סԺ
'˵�����Է�ҩҽ��,��ǰ��ִ�п��ҿ�����ǿ��ѡ�������,��Ҫ��ʾ��ѡ�����;��ѡ���������һ��������ѡ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strҩ�� As String
    Dim bln��� As Boolean, i As Long
    
    If str��� = "4" Then
        strSql = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.�������" & _
            " From " & IIF(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            " And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & " And B.����ID=C.ID " & IIF(lngҩƷID <> 0, " And A.�շ�ϸĿID=[3]", " And A.������ĿID=[4]") & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " Order by B.�������,C.����"
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        bln��� = ((int��Ч = 1 Or gblnҩƷ�������ҽ��) And lngҩƷID <> 0) Or lngҩƷID <> 0
        
        'ϵͳ����ָ��ҩƷִ�п���,������ȡ���п�ѡ�Ĺ���ѡ��
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
        End If
            
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        strSql = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.�������" & _
            " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
            " And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & " And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            IIF(bln���, " And A.�շ�ϸĿID=[3]", " And A.������ĿID=[4]") & _
            " Order by B.�������,C.����"
    Else
        Select Case intִ�п���
            Case 0, 5 '0-��ִ�еĶ���,5-Ժ��ִ��
                Get����ִ�п��� = True: Exit Function
            Case 1, 2, 3, 6 '1-�������ڿ���/2-�������ڲ���/3-����Ա���ڿ���/6-���������ڿ���
                strSql = _
                    " Select ID,����,����,���� From ���ű� Where ID=[5]" & _
                    " Union " & _
                    " Select Distinct A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,������Ա B,��������˵�� C" & _
                    " Where A.ID=B.����ID And A.ID=C.����ID" & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And " & IIF(mint��Χ = 3, "Nvl(C.�������,0)<>0", "C.������� IN([2],3)") & " And B.��ԱID=[6]" & _
                    " Order by ����"
            Case 4 '4-ָ������
                strSql = _
                    " Select Distinct A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,����ִ�п��� B,��������˵�� C" & _
                    " Where A.ID=B.ִ�п���ID And A.ID=C.����ID" & _
                    " And " & IIF(mint��Χ = 3, "Nvl(C.�������,0)<>0", "C.������� IN([2],3)") & " And B.������ĿID=[4]" & _
                    " Union Select ID,����,����,���� From ���ű� Where ID=[5]" & _
                    " Order by ����"
        End Select
    End If
        
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", strҩ��, int��Χ, lngҩƷID, lng��Ŀid, lng��ǰִ��ID, UserInfo.����ID)
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        'ʹ��API���ټ���,��Ȼ�����е���
        AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, rsTmp!���� & "-" & rsTmp!����
        SetComboData objCbo.Hwnd, CB_SETITEMDATA, i - 1, CLng(rsTmp!ID)
        If lng��ǰִ��ID = rsTmp!ID Then
            Call Cbo.SetIndex(objCbo.Hwnd, i - 1)
        End If
        rsTmp.MoveNext
    Next
    
    '����ҩ��������ҽ������ѡ��
    If InStr(",4,5,6,7,", str���) = 0 And objCbo.ListCount = 0 Then
        AddComboItem objCbo.Hwnd, CB_ADDSTRING, 0, "[����...]"
        SetComboData objCbo.Hwnd, CB_SETITEMDATA, objCbo.ListCount - 1, -1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get����ִ�п���ID(ByVal str��� As String, ByVal lng��Ŀid As Long, ByVal lngҩƷID As Long, _
    ByVal intִ�п��� As Integer, ByVal int��Ч As Integer, ByVal int��Χ As Integer) As Long
'���ܣ�����������Ŀִ�п�����Ϣ����ȱʡ��ִ�п���ID
'������lngҩƷID=ҩƷID,ȷ�������ʱҪ��
'      intִ�п���=��Ŀִ�п��ұ�־
'      int��Ч=0-����,1-����,��������Ҫ�ж��ϰ�ʱ��
'      int��Χ=1-����,2-סԺ,3-�����סԺ
'      blnByȱʡ=��ȡȱʡҩ��ʱ�����������ָ�����Ƿ񰴱���ȱʡָ����ҩ������û���򲻷���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim strҩ�� As String, lngҩ�� As Long
    Dim bln��� As Boolean
    
    On Error GoTo errH
    
    If str��� = "4" Then
        lngҩ�� = Val(zldatabase.GetPara(decode(int��Χ, 1, "����", 2, "סԺ", "") & "ȱʡ���ϲ���", glngSys, decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0)))
        strSql = _
            " Select Distinct B.�������,C.����,A.ִ�п���ID" & _
            " From " & IIF(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
            " And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([1],3)") & " And B.����ID=C.ID " & IIF(lngҩƷID <> 0, " And A.�շ�ϸĿID=[2]", " And A.������ĿID=[3]") & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " Order by B.�������,C.����"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", int��Χ, lngҩƷID, lng��Ŀid)
        If Not rsTmp.EOF Then
            If rsTmp.RecordCount = 1 Then
                Get����ִ�п���ID = rsTmp!ִ�п���ID
            ElseIf lngҩ�� <> 0 Then
                rsTmp.Filter = "ִ�п���ID=" & lngҩ��
                If Not rsTmp.EOF Then Get����ִ�п���ID = rsTmp!ִ�п���ID
            End If
        End If
    ElseIf InStr(",5,6,7,", str���) > 0 Then
        bln��� = ((int��Ч = 1 Or gblnҩƷ�������ҽ��) And lngҩƷID <> 0) Or lngҩƷID <> 0
        
        If str��� = "5" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zldatabase.GetPara(decode(int��Χ, 1, "����", 2, "סԺ", "") & "ȱʡ��ҩ��", glngSys, decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0)))
        ElseIf str��� = "6" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zldatabase.GetPara(decode(int��Χ, 1, "����", 2, "סԺ", "") & "ȱʡ��ҩ��", glngSys, decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0)))
        ElseIf str��� = "7" Then
            strҩ�� = "��ҩ��"
            lngҩ�� = Val(zldatabase.GetPara(decode(int��Χ, 1, "����", 2, "סԺ", "") & "ȱʡ��ҩ��", glngSys, decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0)))
        End If
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        strSql = _
            " Select Distinct B.�������,C.����,A.ִ�п���ID" & _
            " From " & IIF(bln���, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
            " And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & " And B.����ID=C.ID" & _
            " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
             IIF(bln���, " And A.�շ�ϸĿID=[3]", " And A.������ĿID=[4]") & _
            " Order by B.�������,C.����"
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", strҩ��, int��Χ, lngҩƷID, lng��Ŀid)
        If Not rsTmp.EOF Then
            If rsTmp.RecordCount = 1 Then
                Get����ִ�п���ID = rsTmp!ִ�п���ID
            ElseIf lngҩ�� <> 0 Then
                rsTmp.Filter = "ִ�п���ID=" & lngҩ��
                If Not rsTmp.EOF Then Get����ִ�п���ID = rsTmp!ִ�п���ID
            End If
        End If
    Else
        Select Case intִ�п���
            Case 0, 5 '0-��ִ�еĶ���/5-Ժ��ִ��
                Exit Function
            Case 1, 2, 3, 6 '1-�������ڿ���/2-�������ڲ���/3-����Ա���ڿ���/6-���������ڿ���
                Get����ִ�п���ID = UserInfo.����ID
            Case 4 '4-ָ������
                strSql = "Select Distinct A.ִ�п���ID From ����ִ�п��� A,��������˵�� B" & _
                    " Where A.ִ�п���ID=B.����ID And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & " And A.������ĿID=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng��Ŀid, int��Χ)
                If Not rsTmp.EOF Then
                    If rsTmp.RecordCount = 1 Then
                        Get����ִ�п���ID = rsTmp!ִ�п���ID
                    End If
                End If
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get����ҩ��IDs(ByVal str��� As String, ByVal lng��Ŀid As Long, _
    ByVal lngҩƷID As Long, ByVal lng����id As Long, Optional ByVal int��Χ As Integer = 2) As String
'���ܣ���ȡҩƷ����Ч����ִ�п���ID��,�����ж�ȱʡִ�п���
'������lng����ID=���˿���ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strҩ�� As String, str����ҩ�� As String
    Dim strҩ��IDs As String
    
    'ϵͳ����ָ��ҩƷִ�п���,������ȡ���п�ѡ�Ĺ���ѡ��
    If str��� = "5" Then
        strҩ�� = "��ҩ��"
        str����ҩ�� = zldatabase.GetPara(decode(int��Χ, 1, "����", 2, "סԺ", "") & "������ҩ��", glngSys, decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0))
    ElseIf str��� = "6" Then
        strҩ�� = "��ҩ��"
        str����ҩ�� = zldatabase.GetPara(decode(int��Χ, 1, "����", 2, "סԺ", "") & "���ó�ҩ��", glngSys, decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0))
    ElseIf str��� = "7" Then
        strҩ�� = "��ҩ��"
        str����ҩ�� = zldatabase.GetPara(decode(int��Χ, 1, "����", 2, "סԺ", "") & "������ҩ��", glngSys, decode(int��Χ, 1, p����ҽ���´�, 2, pסԺҽ���´�, 0))
    End If
        
    'ҩƷ��ϵͳָ���Ĵ���ҩ������
    strSql = _
        " Select Distinct C.ID" & _
        " From " & IIF(lngҩƷID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
        " Where A.ִ�п���ID+0=B.����ID And B.��������=[1]" & _
        " And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([2],3)") & " And B.����ID=C.ID" & _
        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
        IIF(int��Χ <> 3, " And (A.������Դ is NULL Or A.������Դ=[2])", "") & _
        IIF(lng����id <> 0, " And (A.��������ID is NULL Or A.��������ID=[3])", "") & _
        IIF(lngҩƷID <> 0, " And A.�շ�ϸĿID=[4]", " And A.������ĿID=[5]")
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", strҩ��, int��Χ, lng����id, lngҩƷID, lng��Ŀid)
    Do While Not rsTmp.EOF
        If str����ҩ�� = "" Then
            strҩ��IDs = strҩ��IDs & "," & rsTmp!ID
        ElseIf InStr("," & str����ҩ�� & ",", "," & rsTmp!ID & ",") > 0 Then
            strҩ��IDs = strҩ��IDs & "," & rsTmp!ID
        End If
        rsTmp.MoveNext
    Loop
    Get����ҩ��IDs = Mid(strҩ��IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get���÷��ϲ���IDs(ByVal lng����ID As Long, ByVal lng����id As Long, Optional ByVal int��Χ As Integer = 2, Optional ByVal lng��Ŀid As Long) As String
'���ܣ���ȡ���ĵ���Ч����ִ�п���ID��,�����ж�ȱʡִ�п���
'������lng����ID=���˿���ID
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, str���ϲ���IDs As String
    
    strSql = _
        " Select Distinct C.ID" & _
        " From " & IIF(lng����ID <> 0, "�շ�ִ�п���", "����ִ�п���") & " A,��������˵�� B,���ű� C" & _
        " Where A.ִ�п���ID+0=B.����ID And B.��������='���ϲ���'" & _
        " And " & IIF(mint��Χ = 3, "Nvl(B.�������,0)<>0", "B.������� IN([1],3)") & " And B.����ID=C.ID " & IIF(lng����ID <> 0, " And A.�շ�ϸĿID=[3]", " And A.������ĿID=[4]") & _
        " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
        IIF(int��Χ <> 3, " And (A.������Դ is NULL Or A.������Դ=[1])", "") & _
        IIF(lng����id <> 0, " And (A.��������ID is NULL Or A.��������ID=[2])", "")
    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "mdlCISKernel", int��Χ, lng����id, lng����ID, lng��Ŀid)
    Do While Not rsTmp.EOF
        str���ϲ���IDs = str���ϲ���IDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    Get���÷��ϲ���IDs = Mid(str���ϲ���IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
