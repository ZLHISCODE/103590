VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "*\A..\idking\zlIDKind.vbp"
Begin VB.Form frmPacsQuery 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer TimFlicker 
      Interval        =   500
      Left            =   120
      Top             =   1560
   End
   Begin VB.PictureBox PicLine 
      BorderStyle     =   0  'None
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   5775
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6000
      Width           =   5775
   End
   Begin VB.TextBox txtDetail 
      BackColor       =   &H00FDD6C6&
      Height          =   615
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "frmPacsQuery.frx":0000
      Top             =   7560
      Width           =   5775
   End
   Begin VB.PictureBox picListRowInfo 
      Height          =   615
      Left            =   720
      ScaleHeight     =   555
      ScaleWidth      =   5715
      TabIndex        =   10
      Top             =   6840
      Width           =   5775
      Begin VB.Label labPatientInfoName 
         AutoSize        =   -1  'True
         Caption         =   "??? ? ???"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
      Begin VB.Image imgState 
         Height          =   375
         Index           =   0
         Left            =   4920
         Top             =   120
         Width           =   495
      End
      Begin VB.Label labPatientInfoNo 
         AutoSize        =   -1  'True
         Caption         =   "���ţ�99999999 ��ʶ�ţ�12345678"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   11
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.PictureBox picHistory 
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   5715
      TabIndex        =   7
      Top             =   6240
      Width           =   5775
      Begin VB.ComboBox cboHistory 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ʷ��飺"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.PictureBox picVsf 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   720
      ScaleHeight     =   3975
      ScaleWidth      =   5775
      TabIndex        =   4
      Top             =   1920
      Width           =   5775
      Begin VB.PictureBox picGroup 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   5115
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Bindings        =   "frmPacsQuery.frx":000F
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5175
         _cx             =   9128
         _cy             =   5318
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
         FixedCols       =   1
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
         OwnerDraw       =   4
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
   Begin VB.PictureBox picFilter 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   720
      ScaleHeight     =   615
      ScaleWidth      =   5775
      TabIndex        =   3
      Top             =   1200
      Width           =   5775
      Begin XtremeCommandBars.CommandBars cbrFilter 
         Left            =   0
         Top             =   120
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picSearch 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   5775
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      Begin zlIDKind.PatiIdentify patiSearch 
         Bindings        =   "frmPacsQuery.frx":0023
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPacsQuery.frx":0037
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         DefaultCardType =   "���￨"
         IDKindWidth     =   1200
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
      Begin XtremeCommandBars.CommandBars cbrBaseFilter 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picTag2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin XtremeSuiteControls.TabControl tabQuery 
      Bindings        =   "frmPacsQuery.frx":00EA
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   0
      Width           =   4485
      _Version        =   589884
      _ExtentX        =   7911
      _ExtentY        =   873
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":00FE
            Key             =   "��ѡ����"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":0698
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":0C32
            Key             =   "��λ"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":0FC4
            Key             =   "����"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":1356
            Key             =   "��ѡ����"
            Object.Tag             =   "90002"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsQuery.frx":1A68
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90003"
         EndProperty
      EndProperty
   End
   Begin VB.Label labHint 
      AutoSize        =   -1  'True
      Caption         =   "----"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   15
      Top             =   8400
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmPacsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const C_LAYOUT_BASEHEIGHTOFDETAILINFO As Long = 1200 ' ��ϸ��Ϣ��׼�߶�1200
Private Const C_LAYOUT_LISTLEFT As Long = 30 ' �б������ƿճ���� 30
Private Const C_ICON_FIND As Long = 4
Private Const C_ICON_LOCATE As Long = 3
Private Const C_ICON_MENUCHOOSE As Long = 90001
Private Const C_ICON_MENUNOCHOOSE As Long = 90000
            
'ĳЩ���ܱ����������
Private mblnRelatingPatient  As Boolean '�Ƿ����ù�������
Private mblnAssignment As Boolean
Private mblSearching As Boolean 'DataSource�����У��ᴥ�� selchange Ҫ�������selchange ����



'������Ϣ/ϵͳ���ñ�������
Private mcnOracle As ADODB.Connection
Private mlngUserId As Long
Private mlngModule As Long              'ģ���
Private mstrCurRoom As String          '����ID
Private mlngSys As Long
Private mstrDBUser As String                 '��ǰ���ݿ��û�
Private mbytFontSize As Byte                '�ֺ�
Private mfrmParent As Object

'��ǰ��ѯ���

Private mTPatientBaseInfo As TPatientBaseInfo '������Ϣ���������½ǲ�����Ϣ����ʷ��飩
Private mrsData As ADODB.Recordset  '���ݿ��ѯ���ļ�¼�����������κ��޸ģ�
Private mrsDataShow As ADODB.Recordset 'mrsData����һЩת����ļ�¼��
Private mDTStart As Date
Private mDTEnd As Date
Private mPatiName  'pati�ؼ�������(����pati�ؼ��Ĳ�ѯ)
Private mTqueryType As TqueryType  '��ѯ���ͣ����ڲ�ѯ������ֵ��LSQ���Ż�����
Private mintShowType As Integer '��ʾ����0��pacsMain    1������
Private mlngSortCol As Long               '��ǰ�����������
Private mintSortOrder As Integer         '��ǰ��������ķ�ʽ

Private WithEvents mobjSqlParse As clsSqlParse  '���ڿ��ٹ��˲���ֵ�Ļ�ȡ
Attribute mobjSqlParse.VB_VarHelpID = -1

''������Ϣ
Private mstrListKeyCol As String '�б�ؼ���  ����"ҽ��ID"
Private mstrCachePath As String

Private mTColSort As TColSort   '������Ϣ

Private mTPatiIdentifyInfo As TPatiIdentifyInfo

Private mPicDictionary As Scripting.Dictionary    'ͼ�껺��
Private mTQuickFilterState As TQuickFilterState   '���ڿ��ٹ��˴���ѡ��״̬���棩
Private mlngSchemeNo As Long '��ǰʹ�õķ�����
Private mColCfgInfo() As Integer    '��������Ϣ��ֻҪ�б�ı���˳���Ӧ�ø���������������ڿ��ٸ��ݵ�ǰ�б�������ҵ���Ӧ�������ã�

Private mSqlScheme As clsSqlScheme '��ǰ����
Private mstrSchemeCfg As TSchemeCfg '�������ò�������ѯ���������ٹ��ˡ���ͷ�ȣ�

Private WithEvents mObjQuery As clsPacsQuery
Attribute mObjQuery.VB_VarHelpID = -1


''��ǰ����Ϣ
Private mTStudyInfo As TStudyInfo '�����Ϣ
Private mlngAdviceID As Long '��ǰҽ��ID(ѡ��ҽ��ID)

''''''''''''''''''''��������ؼ�������չʾ��ر���
Private mDataGrid As VSFlexGrid
Private mlngMove As Long '���沼��λ�����
Private mTLayout As TLayout

'��Ҫ����
Private mobjSquareCard As Object

'''''''''''''�¼�
Public Event OnListRowSelClear() '�б�ѡ����Ŀ���ã����羭�����ٹ��˺����б���ʾ���ݣ���ʱ��Ҫͬ������һЩ״̬��
Public Event OnColStatistics(ByVal strStatisticsInfo As String)   '������ͳ��
Public Event OnDblClick() '˫��
Public Event OnRefreshSelectTab(ByVal lngAdviveID As Long)  '����ԭ���Ĺ���
Public Event OnSelectScheme(ByVal strName As String)
Public Event OnSelChange()
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Type TPatiIdentifyInfo
    blFind As Boolean '�Ƿ����   true:  ����    false:��λ
    blHaveLoad As Boolean '�Ƿ��Ѿ����ع�
    blIsFinding As Boolean '���ڽ��в��ҹ���
    blShowPatiIdentify As Boolean '�Ƿ���ʾˢ���ؼ�
    
    strFindItems As String '������Ŀ��
    strLocateItems As String '��λ��Ŀ��
    strFindItem As String '��ǰ������Ŀ
    strLocateItem As String '��ǰ��λ��Ŀ
    strDefault As String 'Ĭ�ϲ�����
    
    
End Type

'���沼�����
Private Type TLayout
    blShowTimeSelect As Boolean '�Ƿ���ʾʱ��ѡ��˵�
    blShowBaseFilter As Boolean '�������ˣ�ʱ�䣬Pati�ؼ���
    blShowQuickFilter As Boolean '���ٹ���
    blShowHistory As Boolean '��ʷ���������
End Type


'����ɫ��Ϣ
Private Type TColSort
    LngSchemeNo As Long '��ǰ������
    dictSortInfo As Dictionary
End Type

'����ɫ��Ϣ
Private Type TRowColorInfo
    LngSchemeNo As Long '��ǰ������
    intRowColorIndex As Long '�漰����ɫ����
    blHaveRowColor As Boolean
End Type

'��˸��ʱ��Ϣ
Private Type TFlickerInfo
    LngSchemeNo As Long '��ǰ������
    strName As String '��˸�ֶ��� �磺 "������"
    strInfo As String '��ϸ��Ϣ ��"�ѵǼ�,����ʱ��,30|�ѱ���,����ʱ��,40|"
End Type

'��ͳ����Ϣ
Private Type TColTotalInfo
    strName As String
    intCount As Integer
End Type

Private Type TPatientBaseInfo
    lngAdviceID As Long 'ҽ��ID
    lngPatientID As Long '����ID
    lngLinkId As Long        ' ����ID
    lngBaby As Long  ' Ӥ��
    lngMarkNum As Long '��־��
    
    strName As String '����
    strAge As String '����
    strSex As String '�Ա�
End Type
                
Private Type TSchemeCfg
    strSearchCfg  As String '���ٲ�ѯ����
    strFilterCfg As String  '���ٹ�������
    strListCfg As String  '�б�����(˳�򡢿�ȡ��Ƿ�ɼ�)
    strListCfgDefault As String '�б��ʼ����(˳�򡢿�ȡ��Ƿ�ɼ�)
    strListCfgDefaultColOrder As String '�б��ʼ����(ֻ����˳��)
End Type

Private Type TQuickFilterCmdItem
    intItemIndex As Integer '��ţ�1,2,3,4...,99��
    blChoose As Boolean '�Ƿ�ѡ��
    strName As String '��Ŀ����
    strFilterValue As String '��Ŀ�������� "a,b,c,d"
End Type
    
Private Type TQuickFilterCmdState
'�������˶��� ���� "Ӱ�����"->"��λ"
'��Ӱ�������˵ intRelation=1 strRelationName="��λ"
'�Բ�λ��˵ intRelation =2
    cmdItem() As TQuickFilterCmdItem '����Ŀ��Ϣ
    
    intMenuIndex As Integer '���
    intItemCount As Integer '����Ŀ��������һ�����ٹ��˲˵������Ŀ�ѡ��Ŀ��
    intRelation As Integer ' 1:������ǰ��    2:�����к���
    lngID As Long '���˵�ID
    
    strName As String '��������(�ֶ�����)
    strRelationName As String '������������
    strRelationChooseMenu As String '��̬���˵�ѡ�ò˵�"ͷ;��֫;���" Ҫʹ������
    strRelationValueForVBSFilter As String '��Ϻõ�����VBS���˵�����"1,2,3,a,b,c,d,e",�ڲ˵�ѡ����ı�ʱ�仯��ֱ������VBS����
    strCustomScript As String
    strMenuSQL As String
    
    blSimpleFilter As Boolean '�Ƿ�򵥹��ˣ��ǣ���ʾ���ֶ��ǲ�����˵��ֶ�   ����ʾ�ֶβ�������ˣ��˵�Category���Բ������ ��
    blSingleChoose As Boolean '�������Ƿ�ѡ��Ĭ�϶�ѡ

End Type

Private Type TQuickFilterState
    intQuickFilterMenuCount As Integer '������Ŀ����
    TCmdState() As TQuickFilterCmdState '����Ŀ��Ϣ
End Type

Private Enum TqueryType
    ����һ�� = 0
    ���� = 1
    ˢ�� = 2
    ���� = 3
End Enum

Property Get rsDataShow() As ADODB.Recordset
    If Not mrsDataShow Is Nothing Then Set rsDataShow = mrsDataShow
End Property

Property Get rsData() As ADODB.Recordset
    If Not mrsData Is Nothing Then Set rsData = mrsData
End Property

Property Get objSqlScheme() As clsSqlScheme
    Set objSqlScheme = mSqlScheme
End Property

Property Get objQuery() As clsPacsQuery
    Set objQuery = mObjQuery
End Property

Property Get DataGrid() As VSFlexGrid
    Set DataGrid = mDataGrid
End Property

Public Sub SetVars(ByVal strVarName As String, ByVal Value As Variant)
'��Ҫ�ı����Ͳ�����ֵ

    Select Case strVarName
        Case varName_���ݿ�����
            Set mcnOracle = Value
        Case varName_ģ���
            mlngModule = Value
        Case varName_�û�ID
            mlngUserId = Value
        Case varName_����ID
            mstrCurRoom = Value
        Case varName_��ѯ����ID
            mlngSchemeNo = Value
        Case varName_��ѯ��������
            mintShowType = Value
        Case varName_�б�ؼ���
            mstrListKeyCol = Value
        Case varName_ϵͳ��
            mlngSys = Value
        Case varName_���ݿ��û���
            mstrDBUser = Value
        Case varName_�ֺ�
            mbytFontSize = Value
        Case varName_������
            Set mfrmParent = Value
        Case varName_�Ƿ����ù�������
            mblnRelatingPatient = Value
        Case Else
            MsgBox "[SetVars]" & vbLf & "����:[" & strVarName & "]������", vbOKOnly, "�쳣"
    End Select

End Sub

'Public Function SetVars( _
'    ByVal cnOracle As ADODB.Connection, _
'    Optional ByVal lngModule As Long = 0, _
'    Optional ByVal lngUserId As Long = 0, _
'    Optional ByVal strCurRoom As String = "0", _
'    Optional ByVal lngSchemeId As Long = 0, _
'    Optional ByVal intShowType As Integer = 0, _
'    Optional ByVal lngSys As Long = 0, _
'    Optional ByVal strDBUser As String = 0, _
'    Optional ByVal bytFontSize As String = 0)
''��Ҫ�ı����Ͳ�����ֵ
'
'    Set mcnOracle = cnOracle
'    mlngModule = lngModule
'    mlngUserId = lngUserId
'    mstrCurRoom = strCurRoom
'    mlngSchemeNo = lngSchemeId
'    mintShowType = intShowType
'    mstrListKeyCol = "ҽ��ID"
'
'    mlngSys = lngSys
'    mstrDBUser = strDBUser
'    mbytFontSize = bytFontSize
'End Function

Public Function init() As Boolean
'����Ӧ�õ��õĺ���
'��ѯ�����ʼ��
On Error GoTo errHandle
    Dim i As Integer
    Dim intIndex As Integer
    
    Set mObjQuery = New clsPacsQuery
    
    mObjQuery.init mcnOracle, mlngUserId, mstrCachePath
    mObjQuery.LoadQueryScheme mlngModule
    

    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    
    mobjSquareCard.zlInitComponents Me, mlngModule, mlngSys, mstrDBUser, gcnOracle
    patiSearch.zlInit Me, mlngSys, mlngModule, gcnOracle, mstrDBUser, mobjSquareCard, InitCardType("����;")
    
    With tabQuery
        .RemoveAll
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll

        For i = 1 To mObjQuery.SchemeCount
            If mObjQuery.SchemeInfo(i).IsDefault Or mObjQuery.SchemeInfo(i).IsOften Then
                Call .InsertItem(i, mObjQuery.SchemeInfo(i).Name, picTag2.hwnd, 0)
                .Item(.ItemCount - 1).Tag = mObjQuery.SchemeInfo(i).SchemeId
                If mObjQuery.SchemeInfo(i).IsDefault Then
                    intIndex = .ItemCount - 1
                End If
            End If
        Next i

        If .ItemCount >= 1 Then
            If intIndex > 0 Then
                .Item(intIndex).Selected = True
            Else
                .Item(0).Selected = True
            End If
        Else
            Call HaveNoScheme
        End If

    End With
    
    Call ReSetFormFontSize
    
    Exit Function
errHandle:
    Err.Raise -1, "frmPacsQuery", "[init]" & vbCrLf & Err.Description
End Function

Private Sub HaveNoScheme()
'û���κη�����һЩ�ؼ����ɼ���һЩ�ؼ�enable=false
On Error GoTo errHandle
    picSearch.Visible = False
    picFilter.Visible = False
    picVsf.Visible = False
    picHistory.Visible = False
    picListRowInfo.Visible = False
    txtDetail.Visible = False
    
    labHint = "δ�ҵ���Ч�Ĳ�ѯ�������������÷���"
    labHint.Visible = True
    Call labHint.Move(0.5 * (Me.Width - labHint.Width), 0.5 * (Me.Height - labHint.Height))
errHandle:
End Sub

Private Sub cboHistory_DropDown()
On Error GoTo errHandle
    Call SendMessage(cboHistory.hwnd, &H160, 500, 0)
errHandle:
End Sub

Private Sub cboHistory_Click()
On Error GoTo errHandle
    Dim lngAdviceID As Long
    
    If cboHistory.ListCount <= 1 Then Exit Sub
    If cboHistory.Tag = "" Then Exit Sub '��ʱ cboHistory ��Ŀδ������ɣ���listindex ��ֵ����
    
    lngAdviceID = cboHistory.ItemData(cboHistory.ListIndex)
    
    If lngAdviceID = mTStudyInfo.lngAdviceID Then
        Call vsfList_SelChange
        Exit Sub  '�����뵱ǰѡ��ҽ��ID��ͬʱ���ɱ���������
    End If
    
    '���ݲ����ṩ��ͬѡ�
    RaiseEvent OnRefreshSelectTab(lngAdviceID)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbrBaseFilter_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim objControl As CommandBarControl
    Dim objCboControl As CommandBarComboBox
    Dim objfrmTimeSet As frmTimeSet
    Dim DTStartNew As Date
    Dim DTEndNew As Date
    
    Select Case control.Id
        Case conMenu_PacsQuery_TimeLab
        Case conMenu_PacsQuery_TimeCbo

            Set objCboControl = cbrBaseFilter.FindControl(xtpControlComboBox, conMenu_PacsQuery_TimeCbo)
            
            If objCboControl.Text = "�Զ���" Then   '�Զ��嵯��ʱ��ѡ��
                '�����Զ���ʱ�䴦��
                If objfrmTimeSet Is Nothing Then Set objfrmTimeSet = New frmTimeSet
                
                Call objfrmTimeSet.zlShowMe(mfrmParent, mDTStart, mDTEnd, mSqlScheme.dateRange)
                Call objfrmTimeSet.GetTimeSet(DTStartNew, DTEndNew)
                
                mDTStart = DTStartNew
                mDTEnd = DTEndNew
                
                objCboControl.ToolTipText = "�Զ���ʱ�䷶Χ:" & DTStartNew & "��" & DTEndNew
                
                mstrSchemeCfg.strSearchCfg = Split(mstrSchemeCfg.strSearchCfg, ",")(0) & "," & mDTStart & "," & _
                mDTEnd & "," & Split(mstrSchemeCfg.strSearchCfg, ",")(3)
            Else
                objCboControl.ToolTipText = "ʱ��ѡ��"
            End If
            
            If objCboControl.Text <> "������" Then  '��������������±����ʱ��ѡ��
                mstrSchemeCfg.strSearchCfg = objCboControl.Text & "," & Split(mstrSchemeCfg.strSearchCfg, ",")(1) & "," & _
                Split(mstrSchemeCfg.strSearchCfg, ",")(2) & "," & Split(mstrSchemeCfg.strSearchCfg, ",")(3)
            End If
            
        
        Case conMenu_PacsQuery_FindWay
            Dim blFindWayOld As Boolean
            
            blFindWayOld = mTPatiIdentifyInfo.blFind

            mTPatiIdentifyInfo.blFind = Not mTPatiIdentifyInfo.blFind
            
            If blFindWayOld <> mTPatiIdentifyInfo.blFind Then
                Call DoPatiIdentify
            End If
            
            Call SaveLocalPara_PatiIdentify
        Case conMenu_PacsQuery_PatiControl
        Case conMenu_PacsQuery_Do
            If mTPatiIdentifyInfo.blFind Then
                Call ExecuteQuery("����")
            Else
                Call SeekNextPati(patiSearch.Tag <> patiSearch.Text, patiSearch.GetCurCard.����, patiSearch.Text)
            End If
    End Select
    Exit Sub
errHandle:
    MsgBox "[cbrBaseFilter_Execute]" & vbCrLf & Err.Description, vbOKOnly, "�쳣"
End Sub

Private Sub cbrBaseFilter_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCboControl As CommandBarComboBox
    Dim objCusControl As CommandBarControlCustom
    
    Select Case control.Id
        Case conMenu_PacsQuery_TimeLab
        Case conMenu_PacsQuery_TimeCbo
        Case conMenu_PacsQuery_FindWay
            control.IconId = IIf(mTPatiIdentifyInfo.blFind, C_ICON_FIND, C_ICON_LOCATE)
        Case conMenu_PacsQuery_PatiControl
        Case conMenu_PacsQuery_Do
            control.Caption = IIf(mTPatiIdentifyInfo.blFind, "ִ�в���", "ִ�ж�λ")
        Case Else
    End Select
    
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[cbrBaseFilter_Update]" & vbCrLf & Err.Description
End Sub

Private Sub Form_Initialize()
    Set mDataGrid = vsfList
    Set mPicDictionary = New Dictionary
    
    mDataGrid.AllowUserResizing = flexResizeColumns '���ڸı��п�
    mDataGrid.ExplorerBar = 7 '������ͷ�϶�������
    mDataGrid.SelectionMode = flexSelectionListBox '����ѡ������
    mDataGrid.AllowSelection = False '����ѡ������
    mDataGrid.ScrollTrack = True '��������ʱ����
    mDataGrid.FixedCols = 1
    mDataGrid.BackColorSel = &HFEE0E2      '&HFECFD2

    picHistory.BorderStyle = 0
    picListRowInfo.BorderStyle = 0
    
End Sub

Private Sub Label1_Click()
On Error GoTo errH
'    'LSQ ���Թ���
'    Dim t1 As Long
'    Dim t2 As Long
'    Dim strTMp As String
'    Dim i As Long
'    Dim j As Long
'
'
'
'    Debug.Print ""
'    Debug.Print ""
'
   MsgBox "vsfList.MouseRow:" & vsfList.MouseRow
    MsgBox "pati����2:" & patiSearch.Name
'patiSearch.Text = ""  '�л�Itemʱ��Ҫ����������
'    mPatiName = objCard.����
    Exit Sub
errH:
    MsgBox "���Թ���" & Err.Description
End Sub

Private Sub mobjSqlParse_OnGetParameterValue(ByVal strParName As String, Value As Variant)
'��ȡ���ٹ��˵Ĳ���ֵ
On Error GoTo errHandle
    Dim i As Integer
    Dim strValue As String
    Dim strValueAll As String
    Dim j As Integer
    Dim blChooseOne As Boolean
    
    blChooseOne = False
    
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).strName = strParName Then
            Exit For
        End If
    Next
    
    For j = 1 To mTQuickFilterState.TCmdState(i).intItemCount
        If mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then
            strValue = IIf(Len(strValue) = 0, strValue, strValue & ",")
            strValue = strValue & mTQuickFilterState.TCmdState(i).cmdItem(j).strName
            blChooseOne = True
        End If
        strValueAll = IIf(Len(strValueAll) = 0, strValueAll, strValueAll & ",")
        strValueAll = strValueAll & mTQuickFilterState.TCmdState(i).cmdItem(j).strName
    Next
    
    If Not blChooseOne Then
        Value = strValueAll
    Else
        Value = strValue
    End If
    
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[mobjSqlParse_OnGetParameterValue]" & vbCrLf & Err.Description
End Sub

Private Sub patiSearch_ItemClick(Index As Integer, objCard As zlIDKind.Card)
On Error GoTo errHandle
    
    If mblnAssignment Then Exit Sub
    patiSearch.Text = ""  '�л�Itemʱ��Ҫ����������
    mPatiName = objCard.����
    
    If mTPatiIdentifyInfo.blFind Then
        mTPatiIdentifyInfo.strFindItem = mPatiName
    Else
        mTPatiIdentifyInfo.strLocateItem = mPatiName
    End If
    
    Call SaveLocalPara_PatiIdentify
    Exit Sub
errHandle:
    MsgBox "[patiSearch_ItemClick]" & vbCrLf & Err.Description, vbOKOnly, "�쳣"
End Sub


Public Sub StartReadCard()
On Error GoTo errHandle
'��ʼ����
    Dim lngPatientID As Long
    Dim strCurCardName As String

    If mTPatiIdentifyInfo.blFind Then
        strCurCardName = mTPatiIdentifyInfo.strFindItem
    Else
        strCurCardName = mTPatiIdentifyInfo.strLocateItem
    End If

    If patiSearch.GetCurCard.�ӿ���� > 0 Then
        Call mobjSquareCard.zlGetPatiID(patiSearch.GetCurCard.�ӿ����, patiSearch.Text, , lngPatientID)

        Call OnFilterRead(strCurCardName, patiSearch.Text, IIf(lngPatientID > 0, lngPatientID, ""))
    Else
        Call OnFilterRead(strCurCardName, patiSearch.Text, "")
    End If
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[StartReadCard]" & vbCrLf & Err.Description
End Sub

Private Sub OnFilterRead(ByVal strCardName As String, ByVal strFilter As String, ByVal strPatientId As String)
'��ʼ��������
On Error GoTo errHandle

    If mTPatiIdentifyInfo.blFind Then
        '����
        mTPatiIdentifyInfo.blIsFinding = True
        Call ExecuteQuery("����")
        mTPatiIdentifyInfo.blIsFinding = False
        
'        If mrsData.RecordCount < 1 Then
'            Call MsgBoxD(Me, "δ�ҵ��κ�����,��ע��ʱ�䷶Χ�Ƿ���ȷ" & vbCrLf & "  �������:" & strCardName & vbCrLf & "  ��������:" & strFilter, vbOKOnly, "��ʾ")
'        Else
'            If vsfList.Rows <= 1 Then
'                Call MsgBoxD(Me, "��ѯ�����ݵ�δ��ʾ���б���,��ע����ٹ�������", vbOKOnly, "��ʾ")
'            End If
'        End If
    Else
        '��λ
        Call SeekNextPati(patiSearch.Tag <> patiSearch.Text, patiSearch.GetCurCard.����, patiSearch.Text)
    End If

    Call patiSearch.SetFocus
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[OnFilterRead]" & vbCrLf & Err.Description
End Sub


Private Sub patiSearch_KeyPress(KeyAscii As Integer)
'¼���¼�
On Error GoTo errHandle
    Dim blnCard As Boolean
    Dim lngPatientID As Long

    If KeyAscii = 13 Then
        Call StartReadCard

        Exit Sub
    End If

'    If patiSearch.GetCurCard.�Ƿ�ˢ�� Then
'        blnCard = patiSearch.zlIsBrushCard(patiSearch.objTxtInput, KeyAscii)
'
'        If blnCard And Len(patiSearch.Text) = patiSearch.GetCardNoLen - 1 And KeyAscii <> 8 Then  'ˢ����ϴ���
'            patiSearch.Text = patiSearch.Text & Chr(KeyAscii)
'
'            KeyAscii = 0
'
'            If patiSearch.GetCurCard.�ӿ���� > 0 Then
'                Call mobjSquareCard.zlGetPatiID(patiSearch.GetCurCard.�ӿ����, patiSearch.Text, , lngPatientID)
'
'                Call OnFilterRead(patiSearch.GetCurCard.����, patiSearch.Text, IIf(lngPatientID > 0, lngPatientID, ""))
'            Else
'                Call OnFilterRead(patiSearch.GetCurCard.����, patiSearch.Text, "")
'            End If
'        End If
'    End If
    
Exit Sub
errHandle:
    MsgBox "[cbrBaseFilter_Execute]" & vbCrLf & Err.Description, vbOKOnly, "�쳣"
End Sub

Private Sub PicLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���·���ϸ��Ϣ�߶ȿ��Ըı�
On Error GoTo errHandle
    
    If Button = 1 Then
        '��ֵ�ﵽһ����Χ���˳�����
        
        If PicLine.Top + Y < Me.Top + 3000 Or PicLine.Top + Y > Me.Height - picHistory.Height - picListRowInfo.Height Then
            Exit Sub
        End If

        picVsf.Height = picVsf.Height + Y
        PicLine.Top = PicLine.Top + Y
        picHistory.Top = picHistory.Top + Y
        picListRowInfo.Top = picListRowInfo.Top + Y
        txtDetail.Top = txtDetail.Top + Y
        txtDetail.Height = txtDetail.Height - Y
        mlngMove = txtDetail.Height - C_LAYOUT_BASEHEIGHTOFDETAILINFO
        
    End If
    
errHandle:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errH
    SaveLocalPara

    If mlngSchemeNo > 0 And mintShowType = 0 Then Call SaveShemeCustomCfg(mlngSchemeNo)
    
    Set mObjQuery = Nothing
    Set mrsData = Nothing
    Set mobjSqlParse = Nothing
    Set mTColSort.dictSortInfo = Nothing
    Set mPicDictionary = Nothing
    Set mDataGrid = Nothing
    Set mfrmParent = Nothing
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[Form_Unload]" & vbCrLf & Err.Description
End Sub

Private Sub mObjQuery_OnGetParameterValue(ByVal strParName As String, Value As Variant)
On Error GoTo errH
    Dim strTest As String
    Dim strValue As String
    

    Select Case strParName
        Case "ϵͳ.����ID"
            Value = mstrCurRoom
'            Value = "123,123,23,24,25,26"
        Case "ϵͳ.ҽ��ID"
            If mTqueryType = TqueryType.����һ�� Then
            ElseIf mTqueryType = TqueryType.ˢ�� Then
                Value = ""
            Else
                Value = ""
            End If
        Case "ϵͳ.��ʼ����"
            If mTqueryType = TqueryType.����һ�� Then
                Value = ""
            ElseIf mTqueryType = TqueryType.ˢ�� Then
            Else
            End If
        Case "ϵͳ.��������"
            If mTqueryType = TqueryType.����һ�� Then
                Value = ""
            ElseIf mTqueryType = TqueryType.ˢ�� Then
            Else
            End If
        Case Else
            If mTqueryType = ���� Then
                If mPatiName = strParName Then
                    strValue = patiSearch.Text
                    Value = IIf(IsNumeric(strValue), Val(strValue), strValue)
                End If
            End If
            
    End Select
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[mObjQuery_OnGetParameterValue]" & vbCrLf & Err.Description
End Sub

Private Sub picHistory_Resize()
On Error Resume Next
    Dim lngLeft As Long
    
    Label1.Move 120, 80
    lngLeft = Label1.Left + Label1.Width + 120
    cboHistory.Move lngLeft, 30, picHistory.Width - lngLeft - 60
    
End Sub

Private Sub picListRowInfo_Resize()
On Error Resume Next
    Dim i As Integer, j As Integer
    Dim lngLeft As Long
    
    labPatientInfoName.Move C_LAYOUT_LISTLEFT, C_LAYOUT_LISTLEFT, labPatientInfoName.Width, labPatientInfoName.Height
    labPatientInfoNo.Move labPatientInfoName.Left + labPatientInfoName.Width + 2 * C_LAYOUT_LISTLEFT, C_LAYOUT_LISTLEFT, labPatientInfoNo.Width, labPatientInfoNo.Height

    For i = 0 To imgState.Count - 1
        '��������λ��
        lngLeft = picListRowInfo.Width
        For j = 0 To i
            lngLeft = lngLeft - imgState(i).Width
        Next
        Call imgState(i).Move(lngLeft, 0)
    Next
    
End Sub

Private Sub picVsf_Resize()
On Error Resume Next
    Call vsfList.Move(0, 0, picVsf.Width, picVsf.Height)
End Sub

Private Sub tabQuery_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errH
'�����̴�����������Ⱥ�˳�򣬸Ķ������
'˳�� �� ������������·����ţ����ز��� �������б����沼��ˢ��
    Dim i As Long
    
    
    Call SaveLocalPara
    
    If mlngSchemeNo > 0 And mintShowType = 0 Then Call SaveShemeCustomCfg(mlngSchemeNo)
    
    '������ʼ��
    mstrSchemeCfg.strListCfgDefault = ""
    mstrSchemeCfg.strListCfgDefaultColOrder = ""
    mTPatiIdentifyInfo.blHaveLoad = False
    mTPatiIdentifyInfo.strFindItems = ""
    mTPatiIdentifyInfo.strLocateItems = ""
    mTLayout.blShowBaseFilter = False
    mTLayout.blShowTimeSelect = False
    mTPatiIdentifyInfo.blShowPatiIdentify = False
    
    
    
    ReDim mColCfgInfo(0)
    
    mlngSchemeNo = Item.Tag
    
    Call GetLocalPara
    
    Call mObjQuery.ChangeCurScheme(mlngSchemeNo)
    Set mSqlScheme = mObjQuery.GetSqlScheme(mlngSchemeNo)
    
    Call GetSchemePara
    

    Call LoadShemeCustomCfg(mlngSchemeNo)
    Call RefreshQueryWindow(mlngSchemeNo)
    
    
    
    Call ReSetFormFontSize(mbytFontSize)
    
    Call Form_Resize
    
    RaiseEvent OnSelectScheme(Item.Caption)
    Call ExecuteQuery("ˢ��", 1)
    
    Exit Sub
errH:
    MsgBox Err.Description & "tabQuery_SelectedChanged"
End Sub

Private Sub TimFlicker_Timer()
On Error GoTo errH
'   ��ʱ��˸�Ĵ���
    Dim i As Integer, j As Integer
    Dim lngCol As Long, lngColContrast As Long
    Dim strTmp As String
    Dim lngStateColor As Long, lngNextStateColor As Long, lngPreStateColor As Long
    Dim objRowRelation As New clsScRowRelation
    
    Static intSta As Integer
    Static TPFlickerInfo As TFlickerInfo '��ʱ��˸����
    
    '������һ�μ���ʱ��ȡ��ʱ��˸�����Ϣ
    If TPFlickerInfo.LngSchemeNo <> mlngSchemeNo Then
        TPFlickerInfo.strName = ""
        TPFlickerInfo.strInfo = ""
    
        If mSqlScheme Is Nothing Then Exit Sub
        TPFlickerInfo.LngSchemeNo = mlngSchemeNo
        
        For i = 1 To mSqlScheme.ShowCfgCount
            For j = 1 To mSqlScheme.ShowCfg(i).RowRelationCount
                Set objRowRelation = mSqlScheme.ShowCfg(i).RowRelation(j)
                
                If objRowRelation.FlickerTimeOut > 0 Then
                    TPFlickerInfo.strName = mSqlScheme.ShowCfg(i).Name
                    TPFlickerInfo.strInfo = TPFlickerInfo.strInfo & objRowRelation.TiggerData & "," & objRowRelation.TimeOutReferCol & "," & objRowRelation.FlickerTimeOut & "|"

                End If
            Next
        Next
        
        intSta = 0
        Exit Sub
        
    End If
    
    intSta = intSta + 1
    If intSta = 4 Then intSta = 1

    lngCol = vsfList.ColIndex(TPFlickerInfo.strName)
    If vsfList.TopRow = vsfList.BottomRow Then Exit Sub
    For i = vsfList.TopRow To vsfList.BottomRow   '�����ɼ���  For 1
        For j = 0 To UBound(Split(TPFlickerInfo.strInfo, "|")) - 1 '�ж��Ƿ����㳬ʱ���� For 2
            strTmp = Split(TPFlickerInfo.strInfo, "|")(j)
            If Split(strTmp, ",")(0) = vsfList.TextMatrix(i, lngCol) Then
                lngColContrast = vsfList.ColIndex(Split(strTmp, ",")(1))
                
                If IsDate(vsfList.TextMatrix(i, lngColContrast)) Then
                
                    If DateDiff("N", vsfList.TextMatrix(i, lngColContrast), Now) >= Val(Split(strTmp, ",")(2)) Then    '���������õĳ�ʱʱ��
                    
                        '���Ȳ�����˸����
                        lngStateColor = vsfList.Cell(flexcpBackColor, i, 0)
                        lngNextStateColor = RGB(200, 0, 0)
                        lngPreStateColor = RGB(0, 0, 0)
    
                        If intSta = 1 Then
                            vsfList.Cell(flexcpBackColor, i, 0) = lngPreStateColor
                        ElseIf intSta = 2 Then
                            vsfList.Cell(flexcpBackColor, i, 0) = lngStateColor
                        Else
                            vsfList.Cell(flexcpBackColor, i, 0) = lngNextStateColor
                        End If
                    End If
                End If
                
                Exit For   '�����㳬ʱ���� �˳�For 2
            End If
        Next
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[TimFlicher_Timer]" & vbCrLf & Err.Description
End Sub

Private Sub vsfList_AfterSort(ByVal Col As Long, Order As Integer)
'�������Ҫ�����б�ɼ����й�������
On Error GoTo errH
    Dim RowIndex As Long
    
    mlngSortCol = Col
    mintSortOrder = Order
    
    If vsfList.TopRow = vsfList.BottomRow Then Exit Sub
    For RowIndex = vsfList.TopRow To vsfList.BottomRow
        Call RefreshRowRelation(RowIndex)
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[vsfList_AfterSort]" & vbCrLf & Err.Description
End Sub

Private Sub vsfList_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
On Error GoTo errH
'��ȡ֮����б���ʾ��Χ�����ڿ���ִ��RefreshRowRelation
    Dim lngHeight As Long
    Dim RowIndex As Long
    Dim LngListBottom As Long
    
    
'    Debug.Print "��ǰ������" & mlngSchemeNo
'    Debug.Print "��ǰ������" & mSqlScheme.SchemeId
'    Debug.Print "��ǰ��������" & mSqlScheme.SchemeName
'    Debug.Print "��ǰ����1��" & mSqlScheme
'    Debug.Print "��ǰ��ػ���"
'    Debug.Print "ʵ������1��" & mSqlScheme.SchemeName
    
    
    lngHeight = vsfList.BottomRow - vsfList.TopRow
    
    LngListBottom = NewTopRow + lngHeight
    If LngListBottom > vsfList.Rows - 1 Then LngListBottom = vsfList.Rows - 1
    
    For RowIndex = NewTopRow To LngListBottom
        Call RefreshRowRelation(RowIndex)
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[vsfList_BeforeScroll]" & vbCrLf & Err.Description
End Sub

Private Sub vsfList_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo errH
    Dim lngOrder As Long
    
    If Col <> vsfList.ColIndex(GetColSort(vsfList.ColKey(Col))) Then
        
        '������Ҳ���������ַ��������
        If mintSortOrder = 2 Or mintSortOrder = 4 Or mintSortOrder = 6 Or mintSortOrder = 8 Then
            lngOrder = 3
        Else
            lngOrder = 4
        End If
        
        'ʹ������������Ҫ�������򣬺�����Ҫ��Order ����Ϊ0 ����ִ���Դ�������
        Call SetOrder(vsfList.ColIndex(GetColSort(vsfList.ColKey(Col))), lngOrder)
        
        mlngSortCol = vsfList.ColIndex(GetColSort(vsfList.ColKey(Col)))
        mintSortOrder = lngOrder
    
        Order = 0
    End If
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[vsfList_BeforeSort]" & vbCrLf & Err.Description
End Sub

Private Sub vsfList_DblClick()
    RaiseEvent OnDblClick
End Sub

Private Sub vsfList_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim rc As Rect
    
    If Col = 0 And Row > 0 Then
        rc.Bottom = Bottom
        rc.Left = Left
        rc.Right = Right
        rc.Top = Top
        
        Call DrawText(hDC, Row, Len("" & Row), rc, 0)
         
    End If
End Sub

Private Sub vsfList_SelChange()
'�б�ѡ���иı�
'1 ����ҽ��ID��ѯ������Ϣ���ұ��浽�б���
'2 ˢ���б��·��ؼ���ʾ����
'3 ���µ�ǰ�б�ṹ����
On Error GoTo errH
    Dim intCol As Integer
    Dim lngAdviceID As Long
    Static lngListSelectRow As Long

    If mblSearching Then Exit Sub 'datasource ������������������
    
    If mSqlScheme Is Nothing Then Exit Sub

    If vsfList.MouseRow < 1 And vsfList.Row < 1 Then Exit Sub

        
    If vsfList.MouseRow > 0 Then
        '�ֶ��������selchange
        If lngListSelectRow = vsfList.MouseRow Then
            Exit Sub
        Else
            lngListSelectRow = vsfList.MouseRow
        End If
    Else
        '���ˡ�ˢ�µȲ�������selchange
        lngListSelectRow = vsfList.RowSel
    End If

    '״̬ͼ
    Call DoStateImage(lngListSelectRow)
    intCol = vsfList.ColIndex(mstrListKeyCol)
    If intCol = -1 Then Exit Sub

    mlngAdviceID = Val(vsfList.TextMatrix(lngListSelectRow, intCol))
    mTStudyInfo.lngAdviceID = mlngAdviceID

'    �����Ƿ��Ѿ��м�������ж��Ƿ���Ҫ���²�ѯ
    If IsEmpty(vsfList.Cell(flexcpData, lngListSelectRow)) Then
        Call GetStudyInfo
        vsfList.Cell(flexcpData, lngListSelectRow) = mTStudyInfo
    Else
        mTStudyInfo = vsfList.Cell(flexcpData, lngListSelectRow)
        mTStudyInfo.lngAdviceID = mlngAdviceID
    End If
    RaiseEvent OnSelChange

    Call FillHistoryStudy
    Call FillCurAdviceTxtInfor
    Call FillCurAdviceAppend(lngListSelectRow)
    Call SetSelectRowFont
    
    Exit Sub
errH:
    MsgBox "[vsfList_SelChange]" & vbCrLf & Err.Description, vbOKOnly, "�쳣"
End Sub

Private Sub SetSelectRowFont()
'ѡ������������Ӵ֣�����ȡ������Ӵ�
On Error GoTo errH
    
    With vsfList
        
        If .RowSel < 0 Then Exit Sub
    
        If .Cols > 2 And .Rows > 1 Then
            .Cell(flexcpFontBold, .TopRow, 1, .BottomRow, .Cols - 1) = False
            .Cell(flexcpFontBold, .RowSel, 1, .RowSel, .Cols - 1) = True
        End If
    End With
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[SetSelectRowFont]" & vbCrLf & Err.Description
End Sub

Private Sub GetRGB(ByVal lngColor As Long, lngR As Long, lngG As Long, lngB As Long)
On Error GoTo errH
    Dim lngMinVal As Long
    Dim lngMaxVal As Long
    
    lngMinVal = 80
    lngMaxVal = 225
    
    lngR = lngColor Mod 256
    
    If lngR <= lngMinVal Then
        lngR = lngMinVal
    ElseIf lngR > lngMaxVal Then
        lngR = lngMaxVal
    End If
    
    lngG = (Fix(lngColor \ 256)) Mod 256
 
    If lngG <= lngMinVal Then
        lngG = lngMinVal
    ElseIf lngG > lngMaxVal Then
        lngG = lngMaxVal
    End If
    
    lngB = Fix(lngColor \ 256 \ 256)
 
    If lngB <= lngMinVal Then
        lngB = lngMinVal
    ElseIf lngB > lngMaxVal Then
        lngB = lngMaxVal
    End If
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[GetRGB]" & vbCrLf & Err.Description
End Sub

Private Sub FillCurAdviceAppend(ByVal lngListSelectRow As Long, Optional blIsClear As Boolean = False)
'������½���ϸ��Ϣ����Ҫ����һЩ���������ж���Ϣ��ʾ������
On Error GoTo errHandle
    Dim i As Integer
      
    txtDetail = ""
    If blIsClear Then Exit Sub
    
    If vsfList.Rows = 1 Or vsfList.Cols < 2 Then Exit Sub
    For i = 2 To vsfList.Cols
        txtDetail = txtDetail & vsfList.TextMatrix(0, i - 1) & ":  " & LTrim(vsfList.TextMatrix(lngListSelectRow, i - 1)) & vbNewLine
    Next

    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[FillCurAdviceAppend]" & vbCrLf & Err.Description
End Sub

Private Sub GetStudyInfo()
'��ȡ���˻�����Ϣ
On Error GoTo errHandle

    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    
    If mlngModule <> G_LNG_PATHSTATION_MODULE Then
        strSql = "select A.ID ҽ��ID,A.����ID,A.����ʱ��,A.ҽ������,A.����,A.�Ա�,A.����,B.ִ��״̬,B.ִ�й���,C.����ID,C.����" & _
               " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C" & _
               " Where A.ID = [1] And A.���id Is Null And B.ҽ��ID=A.ID " & _
               " AND A.ID=C.ҽ��ID(+)"
    Else
        strSql = "select A.ID ҽ��ID,A.����ID,A.����ʱ��,A.ҽ������,A.����,A.�Ա�,A.����,B.ִ��״̬,B.ִ�й���,C.����ID,D.����� " & _
               " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C,��������Ϣ D" & _
               " Where A.ID = [1] And A.���id Is Null And B.ҽ��ID=A.ID " & _
               " AND A.ID=C.ҽ��ID(+) and C.ҽ��ID=D.ҽ��ID(+)"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯPacs���˻�����Ϣ", mTStudyInfo.lngAdviceID)

    With mTStudyInfo
        .strPatientAge = NVL(rsTemp!����)
        .strPatientName = NVL(rsTemp!����)
        .strPatientSex = NVL(rsTemp!�Ա�)
        .strStudyNum = NVL(rsTemp(GetStudyNumberDisplayName))
        .lngLinkId = NVL(rsTemp!����ID, 0)
        .lngPatId = rsTemp!����ID
    End With

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function GetStudyNumberDisplayName() As String
'��ȡ��������ʾ����
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHSTATION_MODULE, "�����", "����")
End Function

Private Sub FillHistoryStudy()
'�����ʷ����¼
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    
    If mTStudyInfo.lngAdviceID = 0 Then
        cboHistory.Clear
        Exit Sub
    End If

    cboHistory.Tag = "" 'cboHistory������������"������Ŀ"ʱ��������"cboHistory"����
    
    If mlngModule <> G_LNG_PATHSTATION_MODULE Then
        strSql = "select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
               " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C" & _
               " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID " & _
               " AND A.ID=C.ҽ��ID And Instr([2],A.ִ�п���id ) >0"
    Else
        strSql = "select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
               " From ����ҽ����¼ A,����ҽ������ B,��������Ϣ C" & _
               " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID " & _
               " AND A.ID=C.ҽ��ID And Instr([2],A.ִ�п���id ) >0 "
    End If
              
    '���ù������ˣ��Ų�ѯ����ID
    If mblnRelatingPatient = True And mTStudyInfo.lngLinkId <> 0 Then
        If mTStudyInfo.lngLinkId <> 0 Then
            If mlngModule <> G_LNG_PATHSTATION_MODULE Then
                strSql = strSql & " union select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                    " From ����ҽ����¼ A " & _
                    " Where A.id in (select ҽ��ID from Ӱ�����¼ Where ����ID =[3]) "
            Else
                strSql = strSql & " union select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                    " From ����ҽ����¼ A, ��������Ϣ B " & _
                    " Where  A.id in (select ҽ��ID from Ӱ�����¼ Where ����ID =[3]) and a.id=b.ҽ��ID "
            End If
        End If
    End If
    
    strTemp = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
    strTemp = Replace(strTemp, "����ҽ������", "H����ҽ������")
    strTemp = Replace(strTemp, "Ӱ�����¼", "HӰ�����¼")
    strSql = strSql & vbNewLine & " Union ALL " & vbNewLine & strTemp
    strSql = "select * From (" & vbNewLine & strSql & vbNewLine & ") Order By ����ʱ�� Asc"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", mTStudyInfo.lngPatId, mstrCurRoom, mTStudyInfo.lngLinkId)

    cboHistory.Clear
    
    If rsTemp.RecordCount > 50 Then
        If MsgBox("��⵽����ҽ������ʷ��¼����50��ѡ[��]�������أ�ѡ[��]�����������ʷ���", vbYesNo + vbDefaultButton2, "��ʾ") = vbNo Then Exit Sub
    End If
    
    Do Until rsTemp.EOF
        If rsTemp!ҽ��ID = mTStudyInfo.lngAdviceID Then
            '��ǰ
            cboHistory.AddItem "���" & rsTemp.AbsolutePosition & "��/��" & rsTemp.RecordCount & "��(" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & ")  " & _
            Trim(rsTemp!ҽ������)
        Else
            cboHistory.AddItem "  ��" & rsTemp.AbsolutePosition & "��/��" & rsTemp.RecordCount & "��(" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & ")  " & _
            Trim(rsTemp!ҽ������)
        End If
        
        cboHistory.ItemData(cboHistory.NewIndex) = rsTemp!ҽ��ID
       
        If rsTemp!ҽ��ID = mTStudyInfo.lngAdviceID Then cboHistory.ListIndex = cboHistory.NewIndex
        
        rsTemp.MoveNext
    Loop
    
    If cboHistory.ListCount > 1 Then
        cboHistory.ForeColor = &HC0&
    Else
        cboHistory.ForeColor = &H80000008
    End If
    
    cboHistory.Tag = "��ɼ���"

Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub FillCurAdviceTxtInfor()
'������Ϸ����˻�����Ϣ
On Error GoTo errHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intChargeState As Integer
    Dim intColIndex As Integer
    Dim blnQueryMoneyState As Boolean

    If mTStudyInfo.lngAdviceID <= 0 Then
        labPatientInfoName = "����:  �Ա�:  ����:"
        labPatientInfoNo = "[" & GetStudyNumberDisplayName & ":--- ]"
        Call picListRowInfo_Resize
        Exit Sub
    End If
    
    labPatientInfoName = mTStudyInfo.strPatientName & " " & mTStudyInfo.strPatientSex & " " & mTStudyInfo.strPatientAge

    If mTStudyInfo.lngAdviceID > 0 Then
        labPatientInfoNo.Caption = "[" & GetStudyNumberDisplayName & ":" & IIf(mTStudyInfo.strStudyNum <> "-1", mTStudyInfo.strStudyNum, "--- ") & "]"


'lsq ������Ӥ�����˵Ĵ���
'            If mcurAdviceInf.lngBaby <> 0 Then
'
'                strSql = "select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
'                        "From ������������¼ A, ������Ϣ B" & vbNewLine & _
'                        "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"
'
'                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡӤ����Ϣ", mcurAdviceInf.lngPatId, mcurAdviceInf.lngPageID, mcurAdviceInf.lngBaby)
'
'                If Not rsTemp.EOF Then
'                    labPatientInfoName.Caption = "����:" & NVL(rsTemp!Ӥ������) & "  �Ա�:" & NVL(rsTemp!Ӥ���Ա�) & _
'                                        "  ����:" & NVL(rsTemp!����ʱ��)
'                End If
'            End If

    Else
        labPatientInfoNo.Caption = "[" & GetStudyNumberDisplayName & ":--- ]"
    End If
    
    Call picListRowInfo_Resize
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Resize()
    If mbytFontSize > 0 Then
        Call AdjustFace(mbytFontSize)
    Else
        Call AdjustFace(9)
    End If
End Sub

Private Sub LoadShemeCustomCfg(ByVal LngSchemeNo As Long)
'�����û�ID/����ID���� ���Ի�����,��������Ƿ����Ҫ�󣬲��������Ӧ�����Զ�����Ϊ��
On Error GoTo errH
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strTmp As String
    
    strSql = "select ��������,��������,�б����� from Ӱ���ѯ����  where �û�ID=[1] and ��ѯ����ID =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ز�ѯ���Ի�����", mlngUserId, LngSchemeNo)

    '��ʼ����ѯ����
    
    On Error Resume Next
    
    If rsTemp.RecordCount = 1 Then
        mstrSchemeCfg.strSearchCfg = Split(rsTemp!��������, "%")(0)
        mstrSchemeCfg.strFilterCfg = Split(rsTemp!��������, "%")(1)
        mstrSchemeCfg.strListCfg = rsTemp!�б�����
    End If
    
    '���strSearchCfg
    If UBound(Split(mstrSchemeCfg.strSearchCfg, ",")) < 3 Then mstrSchemeCfg.strSearchCfg = "����," & Date & "," & Date & ","
    
    On Error GoTo errH
    
    mDTStart = CDate(Split(mstrSchemeCfg.strSearchCfg, ",")(1))
    mDTEnd = CDate(Split(mstrSchemeCfg.strSearchCfg, ",")(2))
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[LoadShemeCustomCfg]" & vbCrLf & Err.Description
    
End Sub

Private Function RefreshQueryWindow(ByVal LngSchemeNo As Long) As Boolean
'���ݷ����ı���ٹ��˽���,���ؿ��ٹ��˲˵������ݸ��Ի��������ز˵�ѡ����
On Error GoTo errH
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    
    Dim i As Integer, j As Integer, lngID As Long
    Dim intMenuCount As Integer '���ٹ��˲˵���
    Dim intItemCount As Integer '���ٹ��˲˵�����Ŀ��
    Dim lngCount As Long
    
    Dim strSql As String
    Dim strTemp As String
    Dim strItems As String '�������ݿ��ѯ�����Ŀ��ٲ�ѯ��Ŀ
    Dim strName As String, strValue As String, strItemValue As String, strValueTmp As String
    
    Dim blNeedCreat As Boolean '�������˹�������Ҫ����
    Dim blDynamicFilter As Boolean '�Ƿ�̬����
    
    RefreshQueryWindow = False
    blNeedCreat = True
    
    '''''''������п��ٹ��˲˵�
    Call LockWindowUpdate(Me.hwnd)
    For lngCount = cbrFilter.Count To 2 Step -1
        cbrFilter(lngCount).Delete
    Next
    
    '�ж����Ѿ������������ض��˵�˵���Ѿ������� blNeedCreat ����ΪT
    Set objControl = cbrBaseFilter.FindControl(xtpControlLabel, conMenu_PacsQuery_TimeLab)
    If Not objControl Is Nothing Then
        blNeedCreat = False
    Else
        blNeedCreat = True
    End If
    
    Call LockWindowUpdate(0)
    
    '''''''�����������˲˵�
    Call InitCbrBaseFilter(LngSchemeNo, blNeedCreat)
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrFilter.VisualTheme = xtpThemeOfficeXP
    With cbrFilter.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    cbrFilter.AddImageList img16 '��VB.ImageList��Tag��ID���й���
    cbrFilter.EnableCustomization False
    cbrFilter.ActiveMenuBar.Visible = False
    
    Set objBar = cbrFilter.Add("���ٹ���", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False

    ReDim mTQuickFilterState.TCmdState(mSqlScheme.FilterCfgCount)
    mTQuickFilterState.intQuickFilterMenuCount = 0


    '�������ٹ��˲˵� ��Ϊ�������ͣ�
    '1:  ��ͨ���ٹ��ˣ���ѡ��̶���  �����ã��ѵǼ�;�ѱ���;�ѱ���
    '2:  ��ͨ���ٹ��ˣ���ѡ��ͨ����ѯ�õ��� �����ã� select distinct ���� as Ӱ����� from Ӱ�������
    '3:  �Զ�����ٹ��ˣ���ѡ��ͨ��ǰ��Ŀ��ٹ���ѡ������õ����������ã�[Ӱ�����]#(�����ǿ��ٹ��˿�ѡ���Select��ѯ���)

    For i = 1 To mSqlScheme.FilterCfgCount
        strTemp = mSqlScheme.FilterCfg(i).DataFrom
        
        If InStr(UCase(strTemp), "SELECT") > 0 And Len(mSqlScheme.FilterCfg(i).CustomScript) = 0 Then
        '��������ΪSQL��䣬Ŀǰ����������"SELECT"
            strSql = mSqlScheme.FilterCfg(i).DataFrom
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ٹ��˻�ȡ��Ŀ")
    
            strItems = ""
            strTemp = mSqlScheme.FilterCfg(i).Name
            If rsTemp.RecordCount > 0 Then
                While rsTemp.EOF = False
                    If strItems = "" Then
                        strItems = strItems & rsTemp(strTemp)
                    Else
                        strItems = strItems & ";" & rsTemp(strTemp)
                    End If
                    rsTemp.MoveNext
                Wend
            End If
            
            Call cbrListAdd(objBar.Controls, mSqlScheme.FilterCfg(i).Name, strItems, i, , , , mSqlScheme.FilterCfg(i).SelectWay = swSingle)
        
        ElseIf InStr(UCase(strTemp), "SELECT") = 0 And Len(mSqlScheme.FilterCfg(i).CustomScript) = 0 Then
        '��������Ϊ��ǰ���ã���";"��������
            Call cbrListAdd(objBar.Controls, mSqlScheme.FilterCfg(i).Name, mSqlScheme.FilterCfg(i).DataFrom, i, , , , mSqlScheme.FilterCfg(i).SelectWay = swSingle)
        ElseIf Len(mSqlScheme.FilterCfg(i).CustomScript) > 0 Then
        '��������Ϊǰ����ٹ��������Ľ��
            strTemp = Split(mSqlScheme.FilterCfg(i).DataFrom, "#")(0)
            strTemp = Replace(strTemp, "[", "")
            strTemp = Replace(strTemp, "]", "")
            strTemp = Trim(strTemp)
            mTQuickFilterState.TCmdState(i).strMenuSQL = Split(mSqlScheme.FilterCfg(i).DataFrom, "#")(1)
            
            Call cbrListAdd(objBar.Controls, mSqlScheme.FilterCfg(i).Name, "", i, True, strTemp, mSqlScheme.FilterCfg(i).CustomScript, mSqlScheme.FilterCfg(i).SelectWay = swSingle)
        End If
    Next
    
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next

    cbrFilter.RecalcLayout
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���¹��ܣ����ݲ����ָ���ǰ���ٹ��˲˵���ѡ�����
    '���ݲ������ÿ��ٹ���ѡ�����
    'strValue ���еĿ��ٹ�����Ϣ
    'strValueTmp һ��˵�����Ϣ
    'strItemValue һ��˵�����Ϣ

    If mintShowType = 1 Then Exit Function
    strValue = mstrSchemeCfg.strFilterCfg
    If Len(strValue) = 0 Then Exit Function
    
    intMenuCount = mTQuickFilterState.intQuickFilterMenuCount
    
    '����ÿ�����ٹ��˲˵�
    For i = 1 To intMenuCount
    
        intItemCount = mTQuickFilterState.TCmdState(i).intItemCount
        lngID = 100 * i
        Set objControl = cbrFilter.FindControl(, lngID)
        strName = objControl.Parameter
        
        If UBound(Split(strValue, "|")) <> intMenuCount Then  '����Ĳ˵����������ݿⱣ���������ͬ
            On Error Resume Next
        Else
            On Error GoTo errH
        End If
        
        '�ҵ���Ӧ�Ŀ��ٹ��˲˵�,���ݲ˵����ƻ�ȡ����,��Ҫ�ж��Ƿ�̬����
        For j = 0 To UBound(Split(strValue, "|")) - 1
            strValueTmp = Split(strValue, "|")(j) 'һ��˵�����Ϣ
            
            blDynamicFilter = False
            
            If Split(strValueTmp, ",")(0) = strName Then
                blDynamicFilter = (Split(strValueTmp, ",")(1) = "1")
                strValueTmp = Split(strValueTmp, ",")(2)
                
                '���Ƕ�̬���� ��ȡѡ�в˵���
                If blDynamicFilter Then mTQuickFilterState.TCmdState(i).strRelationChooseMenu = strValueTmp
                Exit For
            End If
        Next
        
        If UBound(Split(strValueTmp, ";")) + 1 <> intItemCount Then '����Ĳ˵����������ݿⱣ���������ͬ
            On Error Resume Next
        Else
            On Error GoTo errH
        End If
        
        '������ٹ��˲˵�����ѡ�����
        If Not blDynamicFilter Then
            '�Ƕ�̬���˸��� 0,1�ж�
            For j = 1 To intItemCount
                mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose = IIf(Val(Split(strValueTmp, ";")(j - 1)) = 1, True, False)
            Next
        End If
        
    Next
    
    Call RefreshCbrQuickFilterALL
    
    RefreshQueryWindow = True
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[RefreshQueryWindow]" & vbCrLf & Err.Description
End Function

Private Function DoStateImage(ByVal lngRow As Long) As Boolean
'����״̬ͼ
On Error GoTo errH
    Dim i As Integer, j As Integer, k1 As Integer, k2 As Integer
    Dim objClsRelation As New clsScRowRelation
    Dim intImgCount As Integer
    Dim lngLeft As Long
    
    '�������״̬ͼ
    For i = imgState.Count - 1 To 0 Step -1
        imgState(i).Visible = False
    Next
    intImgCount = 0
    
    If mSqlScheme.ShowCfgCount < 1 Then Exit Function
    
    With vsfList
        
        For i = 1 To mSqlScheme.ShowCfgCount 'i ��������ʾ����
            If mSqlScheme.ShowCfg(i).RowRelationCount > 0 Then
                
                For j = 1 To mSqlScheme.ShowCfg(i).RowRelationCount 'j�����й���
                    If Len(mSqlScheme.ShowCfg(i).RowRelation(j).Icon) > 0 Then '�����ж��Ƿ���������ʾͼ��
                        If .Cell(flexcpText, lngRow, .ColIndex(mSqlScheme.ShowCfg(i).Name)) = mSqlScheme.ShowCfg(i).RowRelation(j).TiggerData Then '�ж��Ƿ���ϴ�������
                            '���״̬ͼ
                            If intImgCount = 0 Then
                                Set imgState(0).Picture = GetIcon(mSqlScheme.ShowCfg(i).RowRelation(j).Icon)
                                Call imgState(0).Move(labPatientInfoNo.Left + labPatientInfoNo.Width, 0)
                                imgState(0).Visible = True
                                
                                intImgCount = 1
                            Else
                                If imgState.Count <= intImgCount Then Load imgState(intImgCount)

                                Set imgState(intImgCount).Picture = GetIcon(mSqlScheme.ShowCfg(i).RowRelation(j).Icon)
                            
                                '��������λ��
                                lngLeft = picListRowInfo.Width
                                lngLeft = intImgCount * imgState(0).Width
                                Call imgState(intImgCount).Move(lngLeft, 0)
                                imgState(intImgCount).Visible = True
                                
                                intImgCount = intImgCount + 1
                            End If
                            
                        End If
                    End If
                    
                Next  ' for j
            End If
        Next 'for i
    End With
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[DoStateImage]" & vbCrLf & Err.Description
End Function

Private Sub InitCbrBaseFilter(ByVal LngSchemeNo As Long, ByVal blNeedCreat As Boolean)
'��ʼ���������˿ؼ���blNeedCreat�Ƿ���Ҫ�������״μ�����Ҫ���л�����ֻ��Ҫ���¿ؼ�
On Error GoTo errH
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim objPopbar As CommandBarPopup, objCusControl As CommandBarControlCustom
    Dim objCboControl As CommandBarComboBox
    Dim i As Integer
    Dim strSql As String
    
    Call LoadPatiIdentifyInfo
    
    If blNeedCreat Then
        '�½��������Ĵ���
        CommandBarsGlobalSettings.App = App
        CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
        CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
        cbrBaseFilter.VisualTheme = xtpThemeOfficeXP
        With cbrBaseFilter.Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '����VisualTheme����Ч
            .UseDisabledIcons = True
            .LargeIcons = False
            .SetIconSize False, 16, 16
            .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
        End With
        cbrBaseFilter.AddImageList img16 '��VB.ImageList��Tag��ID���й���
        cbrBaseFilter.EnableCustomization False
        cbrBaseFilter.ActiveMenuBar.Visible = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set objBar = cbrBaseFilter.Add("��������", xtpBarTop)
        objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        objBar.ContextMenuPresent = False
    
    ''''''''''''''''''''''''''''''''''''''''''ʱ��˵��Ĵ���
        If mTLayout.blShowTimeSelect Then
            '���ҷ�ʽ
            Set objControl = objBar.Controls.Add(xtpControlLabel, conMenu_PacsQuery_TimeLab, "ʱ�䷶Χ��")
            
            Set objCboControl = objBar.Controls.Add(xtpControlComboBox, conMenu_PacsQuery_TimeCbo, "ʱ��ѡ��")
            Call objCboControl.AddItem("����")
            Call objCboControl.AddItem("����")
            Call objCboControl.AddItem("һ��")
            Call objCboControl.AddItem("�����")
            Call objCboControl.AddItem("һ����")
            Call objCboControl.AddItem("������")
            Call objCboControl.AddItem("����")
            If mSqlScheme.dateRange = 0 Then Call objCboControl.AddItem("������")
            Call objCboControl.AddItem("�Զ���")
            Call SeekIndexSimple(objCboControl, Split(mstrSchemeCfg.strSearchCfg, ",")(0), False)
            
        End If
        
        Set objControl = objBar.Controls.Add(xtpControlButton, conMenu_PacsQuery_FindWay, "��ѯ")
        objControl.Style = xtpButtonIcon
        objControl.IconId = IIf(mTPatiIdentifyInfo.blFind, C_ICON_FIND, C_ICON_LOCATE)
        
        
    
        Set objCusControl = objBar.Controls.Add(xtpControlCustom, conMenu_PacsQuery_PatiControl, "����ֵ")
        objCusControl.Handle = patiSearch.hwnd
        objControl.Visible = mTPatiIdentifyInfo.blShowPatiIdentify
        
        If mTPatiIdentifyInfo.blShowPatiIdentify Then
        '����pati�ؼ�
            Call DoPatiIdentify
        End If
        
        Set objControl = objBar.Controls.Add(xtpControlButton, conMenu_PacsQuery_Do, "ִ��")
        objControl.Style = xtpButtonCaption
    Else
    '���¹������Ĵ���
        
        Set objControl = cbrBaseFilter.FindControl(xtpControlLabel, conMenu_PacsQuery_TimeLab)
        objControl.Visible = mTLayout.blShowTimeSelect
        
        Set objCboControl = cbrBaseFilter.FindControl(xtpControlComboBox, conMenu_PacsQuery_TimeCbo)
        objCboControl.Clear
        Call objCboControl.AddItem("����")
        Call objCboControl.AddItem("����")
        Call objCboControl.AddItem("һ��")
        Call objCboControl.AddItem("�����")
        Call objCboControl.AddItem("һ����")
        Call objCboControl.AddItem("������")
        Call objCboControl.AddItem("����")
        If mSqlScheme.dateRange = 0 Then Call objCboControl.AddItem("������")
        Call objCboControl.AddItem("�Զ���")
        Call SeekIndexSimple(objCboControl, Split(mstrSchemeCfg.strSearchCfg, ",")(0), False)
        objCboControl.Visible = mTLayout.blShowTimeSelect
        
        Set objControl = cbrBaseFilter.FindControl(xtpControlButton, conMenu_PacsQuery_FindWay)
        objControl.Visible = mTPatiIdentifyInfo.blShowPatiIdentify
            
        'ע��ʹ���������� xtpControlButton
        Set objControl = cbrBaseFilter.FindControl(xtpControlButton, conMenu_PacsQuery_PatiControl)
        objControl.Visible = mTPatiIdentifyInfo.blShowPatiIdentify
        
        If mTPatiIdentifyInfo.blShowPatiIdentify Then
        '����pati�ؼ�
            Call DoPatiIdentify
        End If
        
        Set objControl = cbrBaseFilter.FindControl(xtpControlButton, conMenu_PacsQuery_Do)
        objControl.Visible = mTPatiIdentifyInfo.blShowPatiIdentify
    End If
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[InitCbrBaseFilter]" & vbCrLf & Err.Description
End Sub

Private Sub SaveShemeCustomCfg(ByVal LngSchemeNo As Long)
'������Ի�����
On Error GoTo errH
    Dim strSql As String
    Dim objCboControl As CommandBarComboBox
    
    strSql = "Zl_Ӱ���ѯ_���Ի�����(" & mlngUserId & "," & LngSchemeNo & ",'" & mstrSchemeCfg.strSearchCfg & "%" & mstrSchemeCfg.strFilterCfg & "','" & mstrSchemeCfg.strListCfg & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "������Ի�����")
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[SaveShemeCustomCfg]" & vbCrLf & Err.Description
End Sub

Private Sub cbrFilter_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
'���ٹ��� ִ��
On Error GoTo errHandle
    Dim i As Integer
    Dim strTemp As String
    Dim lngAdviceID As Long
    Dim intIndex As Integer '�˵����
    Dim intItemIndex As Integer '��Ŀ���
    Dim intTMP As Integer
    
    intTMP = control.Id
    If (intTMP Mod 100) <> 0 Then
        intIndex = Int(intTMP / 100) + 1
        intItemIndex = (intTMP Mod 100)
        
        '����ѡ��������»����е�ֵ
        If mTQuickFilterState.TCmdState(intIndex).blSingleChoose Then
            For i = 1 To mTQuickFilterState.TCmdState(intIndex).intItemCount
                mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose = False
            Next
            mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose = True
        Else
            mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose = Not mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�Զ�����ٹ��˴���ʼ
'        'ˢ�¶�̬�˵�����
        If mTQuickFilterState.TCmdState(intIndex).intRelation = 1 Then
        '���ǹ�������ǰ�ߣ���Ҫ���º�����ʾ��
            Call RefreshCbrQuickFilter(intIndex, False)
            
        ElseIf mTQuickFilterState.TCmdState(intIndex).intRelation = 2 Then
        '���ǹ������˺��ߣ���Ҫ����ѡ����
            control.Checked = mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose
            
            If control.Checked Then
                If InStr(";" & mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & ";", ";" & control.Parameter & ";") = 0 Then
                    mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu = mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & control.Parameter & ";"
                End If
            Else
                If InStr(";" & mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu & ";", ";" & control.Parameter & ";") > 0 Then
                    mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu = Replace(mTQuickFilterState.TCmdState(intIndex).strRelationChooseMenu, control.Parameter & ";", "")
                End If
            End If
            
            Call GetQuickFilterSQLPar(intIndex)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�Զ�����ٹ��˴������
        
        Call SaveFilterCfg

        ''''''''''''''������ٹ��˺�����ִ�й��˲��ҳ��ֵ��б���
        If Not mrsData Is Nothing Then
            Set mrsDataShow = GetFilterFromQuickFilter
            
            lngAdviceID = GetSelectRowAdviceID
            
            mblSearching = True
            Set vsfList.DataSource = mrsDataShow
            mblSearching = False
            
            '��ͳ��
            Call ColStatistics(mrsDataShow)
            
            '��˳������
            Call LoadListHeadCfg(mlngAdviceID)
            
            '�й���
            If vsfList.TopRow <> vsfList.BottomRow Then
                For i = vsfList.TopRow To vsfList.BottomRow
                    Call RefreshRowRelation(i)
                Next
            End If
            
            '��������
            Call ResetSort(mlngSortCol, mintSortOrder)
        Else
            Call ColStatistics(mrsData)
        End If
        
    End If
    
    Exit Sub
errHandle:
    MsgBox "[cbrFilter_Execute]" & vbCrLf & Err.Description, vbOKOnly, "�쳣"
End Sub

Private Sub cbrFilter_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
'On Error GoTo errHandle ID9400 �±�Խ�� ԭ��δ֪
On Error Resume Next
    Dim i As Integer
    Dim strTemp As String
    Dim intTMP As Integer
    Dim intIndex As Integer
    Dim intItemIndex As Integer
    
    Static blRun As Boolean
    
    If blRun Then Exit Sub
    blRun = True
    strTemp = ""
    intTMP = control.Id
    
    intItemIndex = (intTMP Mod 100)

    If (intTMP Mod 100) = 0 Then
        intIndex = Int(intTMP / 100)
        '��������Ŀѡ����������˵����ƺ�ʹ�õ�ͼ��
        For i = 1 To mTQuickFilterState.TCmdState(intIndex).intItemCount
            
            If mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose Then
                If Len(strTemp) = 0 Then
                    strTemp = mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName
                Else
                    strTemp = strTemp & "," & mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName
                End If
            End If
            
        Next

        If Len(strTemp) = 0 Then
            control.ToolTipText = "����[" & control.Parameter & "]���й���"
            control.Caption = control.Parameter
            control.IconId = C_ICON_MENUNOCHOOSE
        Else
            control.ToolTipText = "��ʾ[" & control.Parameter & "]Ϊ[" & strTemp & "]�ļ��"
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
            control.IconId = C_ICON_MENUCHOOSE
        End If
    
        If mTQuickFilterState.TCmdState(intIndex).intItemCount = 0 Then
            control.ToolTipText = "����[" & control.Parameter & "]���й���"
            control.Caption = control.Parameter
            control.IconId = C_ICON_MENUCHOOSE
            control.Enabled = True
        Else
            control.Enabled = True
        End If
        
    Else
        '�ı�ͼ��

        intIndex = Int(intTMP / 100) + 1
        
        If mTQuickFilterState.TCmdState(intIndex).intRelation <> 2 Then
            control.IconId = IIf(mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose, C_ICON_MENUCHOOSE, C_ICON_MENUNOCHOOSE)
        Else
            If mTQuickFilterState.TCmdState(intIndex).intItemCount < 1 Then
                control.Enabled = False
            End If
            
            control.IconId = IIf(mTQuickFilterState.TCmdState(intIndex).cmdItem(intItemIndex).blChoose, C_ICON_MENUCHOOSE, C_ICON_MENUNOCHOOSE)
        End If
    End If
    
    blRun = False
    
    Exit Sub
'errHandle:
'    Err.Raise -1, "frmPacsQuery", "[cbrFilter_Update]" & vbCrLf & Err.Description
End Sub

Private Sub RefreshCbrQuickFilterALL()
'ˢ�����ж�̬���ٹ���ѡ�������Ҳ���ǳ�ʼ��
On Error GoTo errH
    Dim i As Integer
    Dim intNeedDoCount As Integer
    Dim intNeedDoIndex() As Integer
    
    intNeedDoCount = 0
    ReDim intNeedDoIndex(0)
    
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).intRelation = 2 Then
            
            intNeedDoCount = intNeedDoCount + 1
            ReDim Preserve intNeedDoIndex(intNeedDoCount)
            intNeedDoIndex(intNeedDoCount) = i
        End If
    Next
    
    For i = intNeedDoCount To 1 Step -1
        Call RefreshCbrQuickFilter(intNeedDoIndex(i), True)
    Next
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[RefreshCbrQuickFilterALL]" & vbCrLf & Err.Description
End Sub

Private Sub RefreshCbrQuickFilter(ByVal lngIndex As Long, ByVal blInit As Boolean)
'�����Զ��弰���ٹ��˲˵��������� �˵�ID�������˵�ѡ����"123,456,789"������ʽ
'���ݴ����˵���Ϣ�����Զ���˵���Ϣ��ɾ��ԭ�����Զ���˵����Ȼ���������ɡ�
'blInit �Ƿ��ʼ�� �ǣ�lngIndex��ʾ��Ҫ�ı�Ĳ˵�   �񣺱�ʾ�����˵�ǰ��
On Error GoTo errH
    Dim lngIndexRelationMenu ' �Զ���˵�����
    Dim lngRelationID As Long '�Զ���˵�ID
    Dim strRelationName As String '�Զ���˵�����
    Dim i As Long, j As Long
    Dim ObjPopMenu As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    Dim rsTemp As Recordset
    Dim blHaveSameMenu As Boolean
    Dim strMenuName As String
    Dim intMenuCount As Integer
    Dim strSql As String, strTmp As String
    
    '''��ȡ��Ҫ�ı�Ĳ˵��ڻ����е�ID��ʼ
    If Not blInit Then
        strRelationName = mTQuickFilterState.TCmdState(lngIndex).strRelationName
        
        'Ѱ�ҹ����˵�ID �� ����
        For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
            If mTQuickFilterState.TCmdState(i).strName = strRelationName Then
                lngRelationID = mTQuickFilterState.TCmdState(i).lngID
                lngIndexRelationMenu = i
                Exit For
            End If
        Next
    Else
        lngRelationID = mTQuickFilterState.TCmdState(lngIndex).lngID
        lngIndexRelationMenu = lngIndex
    End If
    '''��ȡ��Ҫ�ı�Ĳ˵��ڻ����е�ID����

    '''������в˵����ʼ
    Call LockWindowUpdate(Me.hwnd)
    
    Set ObjPopMenu = cbrFilter.FindControl(, lngRelationID)
    
    If Not ObjPopMenu Is Nothing Then
        For i = 1 To ObjPopMenu.CommandBar.Controls.Count
            ObjPopMenu.CommandBar.Controls(1).Delete
        Next
    End If
    '''������в˵��������
    
    Call LockWindowUpdate(0)
    
    '�����˵�
    strSql = mTQuickFilterState.TCmdState(lngIndexRelationMenu).strMenuSQL
            
    If mobjSqlParse Is Nothing Then Set mobjSqlParse = New clsSqlParse
    Call mobjSqlParse.init(strSql)
    strSql = mobjSqlParse.GetQuerySql
    
    Set rsTemp = ExecuteCore(strSql, "���ٹ��˻�ȡ��Ŀ", mobjSqlParse.ParValues)
    
    ''''�ж��Ƿ��Ӳ˵���ʾ�����ǹ�������,�� rsTemp.Fields.Count ��ֵ�й�
    ''�ǣ����� Ӱ�����-��λ����     ���ڲ�λ������˵����ʾ���Ƿ��飬������˵��ǲ�λ��rsTemp.Fields.Count Ӧ���� 2
    ''��rsTemp.Fields.Count Ӧ����1 �˵���ʾ�����ݾ��ǲ�����˵�����
    If rsTemp.RecordCount = 0 Then
        ReDim mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(0)
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount = 0
        ObjPopMenu.Enabled = False
        Exit Sub
    End If
    
    '��ȡѡ��˵����ã����ڻָ�ѡ�������
    strTmp = mTQuickFilterState.TCmdState(lngIndexRelationMenu).strRelationChooseMenu
    
    If rsTemp.Fields.Count = 1 Then
    '''1 ��ʾ�������ݾ�����ʾ����
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).blSimpleFilter = True
        ReDim mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(rsTemp.RecordCount)
        
        For i = 1 To rsTemp.RecordCount
            strMenuName = rsTemp.Fields(0).Value
            Set cbrControl = ObjPopMenu.CommandBar.Controls.Add(xtpControlButton, lngRelationID - 100 + i, strMenuName)
            
            cbrControl.Parameter = strMenuName
            
            mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(i).blChoose = IIf(InStr(";" & strTmp & ";", ";" & strMenuName & ";") > 0, True, False)
            mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(i).intItemIndex = i
            mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(i).strName = strMenuName
            cbrControl.CloseSubMenuOnClick = False
            If i <> rsTemp.RecordCount Then rsTemp.MoveNext
        Next
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount = rsTemp.RecordCount
    Else
    '''rsTemp.Fields.Count <> 1(=2) ��ʾ��һ���ֶ�����ʾ���� �ڶ����ֶ���ʵ�ʹ�������
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).blSimpleFilter = False
        rsTemp.MoveFirst

        intMenuCount = 1
        
        ReDim mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(0)
        mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount = 0
               
        For i = 1 To rsTemp.RecordCount
            strMenuName = rsTemp.Fields(0).Value
            blHaveSameMenu = False
            
            '�����ж��Ƿ����ظ��ķ��飬���У��Ѻ�����ǰ�߲�ͬ�Ĳ�λ�ӵ�ǰ��Ĳ˵���
            If mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount > 0 Then
                For j = 1 To mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount
                    If rsTemp.Fields(0).Value = mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(j).strName Then
    '                        '���ظ�����
                        Set cbrControl = cbrFilter.FindControl(, 100 * (lngIndex - 1) + j, , True)
    
                        '����ط������򵥵Ĵ���ֻ��֮ǰû�еĲ�λ����Ҫ��������
                        cbrControl.Category = cbrControl.Category & CbrFilterDeal(cbrControl.Category, rsTemp.Fields(1).Value)
                        cbrControl.Category = Replace(cbrControl.Category, ",,", ",")
                        mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(j).strFilterValue = cbrControl.Category
    
                        '���Ӻ��ظ�����Ŀ
                        blHaveSameMenu = True
                    End If
                Next
            End If

            If Not blHaveSameMenu Then
            'û���ظ���ֱ������
                ReDim Preserve mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount + 1)

                Set cbrControl = ObjPopMenu.CommandBar.Controls.Add(xtpControlButton, lngRelationID - 100 + intMenuCount, strMenuName)

                cbrControl.Parameter = strMenuName
                cbrControl.Category = rsTemp.Fields(1).Value
                
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(intMenuCount).blChoose = IIf(InStr(";" & strTmp & ";", ";" & strMenuName & ";") > 0, True, False)
                
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(intMenuCount).intItemIndex = i
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(intMenuCount).strName = strMenuName
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).cmdItem(intMenuCount).strFilterValue = cbrControl.Category
                mTQuickFilterState.TCmdState(lngIndexRelationMenu).intItemCount = intMenuCount

                intMenuCount = intMenuCount + 1

                cbrControl.CloseSubMenuOnClick = False

            End If

             If i <> rsTemp.RecordCount Then rsTemp.MoveNext
        Next
               
        Call GetQuickFilterSQLPar(lngIndex)
    End If
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[RefreshCbrQuickFilter]" & vbCrLf & Err.Description
End Sub

Private Sub cbrListAdd(ByVal Mycontrol As CommandBarControls, ByVal strName As String, ByVal strItems As String, ByVal intIndex As Integer, _
Optional blDynamic As Boolean = False, Optional strRelationName As String = "", Optional strCustomScript As String = "", Optional blSingleChoose As Boolean = False)
'ÿ�����˲˵���ռ��100��ID ����1�� 1~99  2��100~199
'��cbrList�����ӿ��ٹ��˲˵�
'strName���˵��� ��:�������
'strItems���˵���Ŀ�����磺 �ѵǼ�;�ѱ���;�Ѽ�� ��";"�ֿ����Ƕ�̬���ٹ��� �ɴ˻�ȡ�˵�����
'intIndex:�ڼ���������
'blDynamic�Ƿ�̬���ٹ���  ���ǣ���Ҫ�ڹ����˵��н���һЩ����
' strRelationName ������Ŀ����
'rsData: �˵���¼��

On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim i As Integer
    Dim intCount As Integer
    Dim rsTemp As Recordset
    
    '���ٹ��˲˵��� 1
    mTQuickFilterState.intQuickFilterMenuCount = mTQuickFilterState.intQuickFilterMenuCount + 1
    
    intCount = UBound(Split(strItems, ";")) + 1
    
    Set objControl = Mycontrol.Add(xtpControlButtonPopup, 100 * intIndex, strName)
    objControl.ToolTipText = "����" & strName & "����"
    
    objControl.Parameter = strName
    mTQuickFilterState.TCmdState(intIndex).intMenuIndex = mTQuickFilterState.intQuickFilterMenuCount
    mTQuickFilterState.TCmdState(intIndex).intItemCount = intCount
    mTQuickFilterState.TCmdState(intIndex).strName = strName
    mTQuickFilterState.TCmdState(intIndex).lngID = 100 * intIndex
    mTQuickFilterState.TCmdState(intIndex).blSingleChoose = blSingleChoose
    
    '100 * (intIndex - 1) + i�����������ٹ������ID
    If blDynamic Then
        '����������������2
        mTQuickFilterState.TCmdState(intIndex).intRelation = 2
        mTQuickFilterState.TCmdState(intIndex).strCustomScript = strCustomScript
        mTQuickFilterState.TCmdState(intIndex).strRelationName = strRelationName
        '��������Ѱ�ұ���������
        For i = 1 To intIndex - 1
            If mTQuickFilterState.TCmdState(i).strName = strRelationName Then
                
                '������������������1
                mTQuickFilterState.TCmdState(i).intRelation = 1
                '�������������ù�����������
                mTQuickFilterState.TCmdState(i).strRelationName = strName
                '�ҵ����˳�
                Exit For
            End If
        Next
        
    Else
        ReDim mTQuickFilterState.TCmdState(intIndex).cmdItem(intCount)
        '���Զ�����ٹ���
        For i = 1 To intCount
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, 100 * (intIndex - 1) + i, Split(strItems, ";")(i - 1))
            cbrPopControl.Parameter = Split(strItems, ";")(i - 1)
            
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).blChoose = False
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).intItemIndex = i
            mTQuickFilterState.TCmdState(intIndex).cmdItem(i).strName = Split(strItems, ";")(i - 1)
            cbrPopControl.CloseSubMenuOnClick = False
        Next
    
    End If
    
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[cbrListAdd]" & vbCrLf & Err.Description
End Sub

Private Sub SaveFilterCfg()
'������ٹ��˲���
On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim strName As String
    Dim strValue As String
    Dim lngID As Long
    Dim i As Integer
    Dim j As Integer
    Dim intMenuCount As Integer '�˵���
    Dim intItemCount As Integer '�˵��Ӳ˵���
    Dim strValueAll As String
    
    strValueAll = ""
    
    intMenuCount = mTQuickFilterState.intQuickFilterMenuCount
    For i = 1 To intMenuCount
        intItemCount = mTQuickFilterState.TCmdState(i).intItemCount
        
        lngID = 100 * i
        Set objControl = cbrFilter.FindControl(, lngID)
        strName = objControl.Parameter
        strValue = ""
        
        If mTQuickFilterState.TCmdState(i).intRelation <> 2 Then
            '�Ƕ�̬���ٹ��˵ı��棬��˳����0/1��ʾ�Ƿ�ѡ��
            If intItemCount > 0 Then
                For j = 1 To intItemCount
                    
                    If Len(strValue) = 0 Then
                        strValue = strValue & IIf(mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose, "1", "0")
                    Else
                        strValue = strValue & ";" & IIf(mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose, "1", "0")
                    End If
                Next
            End If
            'ע��˴���"0"������ʾ�Ƕ�̬���ٹ���
            strValue = strName & ",0," & strValue
            strValueAll = strValueAll & strValue & "|"
        Else
            '��̬���ٹ��˵ı��棬����Ĳ˵�����
            If intItemCount > 0 Then
                For j = 1 To intItemCount
                    If mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then strValue = strValue & mTQuickFilterState.TCmdState(i).cmdItem(j).strName & ";"
                Next
            End If
            'ע��˴���"1"������ʾ��̬���ٹ���
            strValue = strName & ",1," & strValue
            strValueAll = strValueAll & strValue & "|"
        End If

    Next
    
    mstrSchemeCfg.strFilterCfg = strValueAll
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[SaveFilterCfg]" & vbCrLf & Err.Description
End Sub
Private Function GetFilterFromQuickFilter() As Recordset
'��ȡ���ٹ������������ҽ��й���,���ع��˺�ļ�¼��
'��mrsData ���������к󲻻�ı�mrsData,���ؾ������˺��ȫ�¼�¼��
'��ÿ��ٹ�������
On Error GoTo errH
    Dim objControl As CommandBarControl
    Dim strFilter As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim blChooseOne As Boolean 'ĳ���˵��Ƿ��й����ѡ�У���û���൱�ڲ�����
    Dim strFilterTmp As String
    Dim blIsStr As Boolean '���ٹ��������Ƿ����ַ��������Ƿ����� ��' �������
    Dim strFilterField As String
    Dim intCustomeFilter As Integer   '�Ƿ����Զ�����ٹ���
    
    intCustomeFilter = 0
    strFilter = ""
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        blChooseOne = False
        strFilterTmp = ""
        Set objControl = cbrFilter.FindControl(, i * 100)
        blIsStr = True
        
        For j = 0 To mrsData.Fields.Count - 1
            
            If objControl.Parameter = mrsData.Fields(j).Name Then
                If mrsData.Fields(j).Type = adVarNumeric Or mrsData.Fields(j).Type = adNumeric Then
                    blIsStr = False
                    Exit For
                End If
            End If
        Next
        
        '���Ƚ��й̶����ٹ��˵Ĵ���
        If mTQuickFilterState.TCmdState(i).intRelation <> 2 Then
            For j = 1 To mTQuickFilterState.TCmdState(i).intItemCount
                strFilterField = mTQuickFilterState.TCmdState(i).cmdItem(j).strName
            
                If InStr(strFilterField, "-") > 0 Then strFilterField = Mid(strFilterField, 1, InStr(strFilterField, "-") - 1)
    
                If mTQuickFilterState.TCmdState(i).intRelation <> 2 Then
                    If Not mTQuickFilterState.TCmdState(i).cmdItem(j).blChoose Then
                    'δ��ѡ��
                        If Not blIsStr Then
                            If Len(strFilterTmp) = 0 Then
                                strFilterTmp = strFilterTmp & objControl.Parameter & " <> " & strFilterField & " "
                            Else
                                strFilterTmp = strFilterTmp & " and " & objControl.Parameter & " <> " & strFilterField & " "
                            End If
                        Else
                            If Len(strFilterTmp) = 0 Then
                                strFilterTmp = strFilterTmp & objControl.Parameter & " <> '" & strFilterField & "' "
                            Else
                                strFilterTmp = strFilterTmp & " and " & objControl.Parameter & " <> '" & strFilterField & "' "
                            End If
                        End If
                    Else
                        blChooseOne = True
                    End If
                    
                End If
            Next
        Else
            intCustomeFilter = intCustomeFilter + 1
        End If
        'ֻ�б�ѡ���˹�����ż��뵽����������
        If blChooseOne And Len(strFilterTmp) > 0 Then
            If Len(strFilter) = 0 Then
                strFilter = strFilter & strFilterTmp
            Else
                strFilter = strFilter & " and " & strFilterTmp
            End If
        End If

    Next
    mrsData.Filter = strFilter

    'û���Զ�����ٹ��˿��������˳�
    If intCustomeFilter = 0 Then
        Set GetFilterFromQuickFilter = CopyRecordSet(mrsData)
        Exit Function
    End If
    
    Dim strVBS As String
    Dim rstVBS As Recordset
    Dim rsTmp() As Recordset
    Dim objGlobal As clsGlobal

    Set objGlobal = New clsGlobal
    '����ÿ���Զ�����ٹ�������,���ݲ˵�ѡ�����ִ��VBS�ű�
    ReDim rsTmp(intCustomeFilter)
    j = 0
    
    For i = 1 To mTQuickFilterState.intQuickFilterMenuCount
        If mTQuickFilterState.TCmdState(i).intRelation = 2 Then
            j = j + 1
            If j = 1 Then
                Set rsTmp(0) = CopyRecordSet(mrsData)
            Else
                Set rsTmp(j - 1) = CopyRecordSet(rsTmp(j - 2))
            End If

            Call objGlobal.ExecuteScript(mTQuickFilterState.TCmdState(i).strCustomScript, rsTmp(j - 1), mTQuickFilterState.TCmdState(i).strRelationValueForVBSFilter)

        End If
    Next
    
    Set GetFilterFromQuickFilter = rsTmp(intCustomeFilter - 1)
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetFilterFromQuickFilter]" & vbCrLf & Err.Description
End Function

Private Function GetListHeadString() As String
'�õ���������: ����,���,�Ƿ���ʾ  ����  "���,1000,1|ִ�й���,2000,0|"
On Error GoTo errH
    Dim i As Integer
    Dim strValue As String
    Dim strTemp As String
    
    strTemp = ""
    
    For i = 1 To vsfList.Cols - 1
        If Len(strTemp) > 0 Then
            strTemp = strTemp & "|"
        End If
        
        strTemp = strTemp & vsfList.TextMatrix(0, i)
        strTemp = strTemp & "," & vsfList.ColWidth(i)
        strTemp = strTemp & "," & IIf(vsfList.ColHidden(i), "0", "1")
        
    Next

    GetListHeadString = strTemp
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetListHeadString]" & vbCrLf & Err.Description
End Function

Private Sub SaveListHeadCfg()
'���������ͷ����
On Error GoTo errH
    Dim strValue As String
    
    mstrSchemeCfg.strListCfg = GetListHeadString()
    Call DoBeforeFilterData(True)
    
    Exit Sub
errH:
    MsgBox "SaveListHeadCfg" & Err.Description
End Sub

Private Function LoadListHeadCfg(ByVal lngAdviceIDOld As Long) As Boolean
'LoadListHeadCfg��֮ǰ��ѡ�м��ҽ��ID����Ϊ0��ֱ��ѡ�е�һ��,��Ϊ -1 ����Ҫ�ı�ѡ��
'���ݸ��Ի�����ˢ���б�����(���򡢿�ȡ��Ƿ���ʾ)
'ע��i��j�ĳ�ʼֵ

'�ú�����Ҫ����֮ǰѡ��״̬�ı�������֮ǰѡ���˼�飬���κβ�����Ӧ�ñ���֮ǰ�ļ�飬��֮ǰ�ļ���Ѿ���ʧ����ѡ�е�һ����
On Error GoTo errH
    Dim strTmp As String
    Dim strValue As String
    Dim strColName As String
    Dim intWidth As Integer
    Dim lngAdviceIDNew As Long
    
    Dim blShow As Boolean
    Dim blMatch As Boolean '�Ѿ�������б��Ƿ������� ƥ�䣬����ƥ�䣨�����Ѿ�ʹ�ã������޸������á����������ֶΣ�
    Dim i As Integer
    Dim j As Integer
    
    
    blMatch = True
'    If mSqlScheme.ShowCfgCount < 1 Then Exit Function
    If mintShowType = 1 Then Exit Function
    
    strValue = mstrSchemeCfg.strListCfg
    If Len(strValue) = 0 Then Exit Function
    
    '�жϱ�����б������Ƿ��뵱ǰ�б�ƥ�䣨�б��ֶΣ�����������Ϊ������������ܵ��¾����ò�����
    If UBound(Split(strValue, "|")) <> vsfList.Cols - 2 Then blMatch = False
    For i = 1 To vsfList.Cols - 1
        If InStr(strValue, vsfList.TextMatrix(0, i)) = 0 Then blMatch = False
    Next
    
    If blMatch Then
    '"LoadListHeadCfgƥ��"
        For i = 1 To vsfList.Cols - 1

            strTmp = Split(strValue, "|")(i - 1)
            strColName = Split(strTmp, ",")(0)
            intWidth = Val(Split(strTmp, ",")(1))
            blShow = Val(Split(strTmp, ",")(2))

            If vsfList.TextMatrix(0, i) <> strColName Then
                For j = 1 To vsfList.Cols - 1
                    If vsfList.TextMatrix(0, j) = strColName Then

                        vsfList.ColPosition(j) = i

                        Exit For
                    End If
                Next
            End If

            vsfList.ColWidth(i) = intWidth
            vsfList.ColHidden(i) = Not blShow

        Next
    Else
        '"LoadListHeadCfg��ƥ��"
        '�б�������ó�ʼ��
        On Error Resume Next
        strValue = mstrSchemeCfg.strListCfgDefault
        
        For i = 1 To vsfList.Cols - 1
    
            strTmp = Split(strValue, "|")(i - 1)
            strColName = Split(strTmp, ",")(0)
            intWidth = Val(Split(strTmp, ",")(1))
            blShow = Val(Split(strTmp, ",")(2))
    
            If vsfList.TextMatrix(0, i) <> strColName Then
                For j = 1 To vsfList.Cols - 1
                    If vsfList.TextMatrix(0, j) = strColName Then
    
                        vsfList.ColPosition(j) = i
    
                        Exit For
                    End If
                Next
            End If
    
            vsfList.ColWidth(i) = intWidth
            vsfList.ColHidden(i) = Not blShow
    
        Next
        On Error GoTo errH
    End If

    Call DoBeforeFilterData(True)
     
    '��������к��޸ĵ�һ�еĿ��
    If vsfList.Rows < 11 Then
        vsfList.ColWidth(0) = TextWidth("XX")
    ElseIf 10 < vsfList.Rows And vsfList.Rows < 101 Then
        vsfList.ColWidth(0) = TextWidth("XXX")
    ElseIf 100 < vsfList.Rows And vsfList.Rows < 1001 Then
        vsfList.ColWidth(0) = TextWidth("XXXX")
    Else
        vsfList.ColWidth(0) = TextWidth("XXXXX")
    End If
    
    If lngAdviceIDOld = -1 Then Exit Function
    
    If vsfList.Rows > 1 Then
        
        '֮ǰ�ļ���б�ѡ���е�ҽ��ID
        If lngAdviceIDOld = 0 Then
            GoTo ClearListFace
        Else
            lngAdviceIDNew = vsfList.FindRow(lngAdviceIDOld, 1, vsfList.ColIndex(mstrListKeyCol), False, False)
            
            If lngAdviceIDNew > 0 Then
                vsfList.Row = lngAdviceIDNew
    
                If vsfList.TopRow > vsfList.Row Then vsfList.TopRow = vsfList.Row
                
                If vsfList.BottomRow - 1 < vsfList.Row Then
                    vsfList.TopRow = vsfList.TopRow + (vsfList.Row - vsfList.BottomRow) + 1
                End If
            Else
                GoTo ClearListFace
            End If
        
        End If
    Else
        GoTo ClearListFace
    End If
    Exit Function
    
ClearListFace:
'ClearListFace ����б�ѡ���κμ�飬�ұ�TabҲ��Ҫ����Ϊδѡ�����״̬
    mlngAdviceID = 0
    For i = imgState.Count - 1 To 0 Step -1
        imgState(i).Visible = False
    Next

    GetNullStudyInfo
    cboHistory.Clear
    Call FillCurAdviceTxtInfor
    Call FillCurAdviceAppend(0, True)
    
    RaiseEvent OnListRowSelClear
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[LoadListHeadCfg]" & vbCrLf & Err.Description
End Function


Private Sub vsfList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    SaveListHeadCfg
End Sub

Private Sub vsfList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    SaveListHeadCfg
End Sub
Public Function ExecuteQuery(ByVal strExecuteType As String, Optional ByVal LngSetRow As Long = 0) As Boolean
'ִ�й��ˡ�ˢ�¡����ҹ���
On Error GoTo errH
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim objCboControl As CommandBarComboBox
    Dim strTmp As String
    Dim i As Integer, j As Integer
    Dim lngAdviceID As Long
    
    Dim bllimitTime As Boolean '�Ƿ�����ʱ��  ѡ�������ơ�ΪFalse  ����Ϊ  true
    
    bllimitTime = True
    
    Set objCboControl = cbrBaseFilter.FindControl(xtpControlComboBox, conMenu_PacsQuery_TimeCbo)
    If objCboControl.Text = "�Զ���" Then
        dtStart = mDTStart
        dtEnd = mDTEnd
    Else
        dtEnd = zlDatabase.Currentdate()
        Select Case objCboControl.Text
            Case "����"
                dtStart = DateAdd("d", -1, dtEnd)
            Case "����"
                dtStart = DateAdd("d", -3, dtEnd)
            Case "һ��"
                dtStart = DateAdd("ww", -1, dtEnd)
            Case "�����"
                dtStart = DateAdd("ww", -2, dtEnd)
            Case "һ����"
                dtStart = DateAdd("m", -1, dtEnd)
            Case "������"
                dtStart = DateAdd("m", -3, dtEnd)
            Case "����"
                dtStart = DateAdd("m", -6, dtEnd)
            Case "������"
                bllimitTime = False
'                '��������ôȡʱ�䣿 �ܴ��ʱ�䣿100�ꣿ
'                dtStart = DateAdd("yyyy", -50, dtEnd)
        End Select
    End If

    If bllimitTime Then
        Call mObjQuery.SetFilterValue("ϵͳ.��ʼ����", dtStart)
        Call mObjQuery.SetFilterValue("ϵͳ.��������", dtEnd)
'    Else
'        Call mObjQuery.SetFilterValue("ϵͳ.��ʼ����", Null)
'        Call mObjQuery.SetFilterValue("ϵͳ.��������", Null)
    End If

    If strExecuteType = "����" Then
        mTqueryType = ����
        Set mrsData = mObjQuery.ExecuteWithFilter(dtStart, dtEnd, Me)
    ElseIf strExecuteType = "ˢ��" Then
        mTqueryType = ˢ��
        Set mrsData = mObjQuery.Execute(dtStart, dtEnd, False)
    ElseIf strExecuteType = "����" Then
        mTqueryType = ����
        Set mrsData = mObjQuery.Execute(dtStart, dtEnd, False)
    End If
    
    Call DoBeforeFilterData(False)
    
    If Not mrsData Is Nothing Then
        '��ȡ�����ֶ���Ϣ
        If Len(mstrSchemeCfg.strListCfgDefaultColOrder) = 0 Then
            For i = 0 To mrsData.Fields.Count - 1
                mstrSchemeCfg.strListCfgDefaultColOrder = mstrSchemeCfg.strListCfgDefaultColOrder & mrsData.Fields(i).Name & "|"
            Next
        End If
             
        If mrsData.RecordCount > 0 Then

            Set mrsDataShow = GetFilterFromQuickFilter
            Set mrsDataShow = mObjQuery.DataConvert(mrsDataShow, mlngSchemeNo)
        
            If Not mrsDataShow Is Nothing Then
                
                lngAdviceID = GetSelectRowAdviceID
                mblSearching = True
                Set vsfList.DataSource = mrsDataShow
                mblSearching = False
                
                If vsfList.TopRow <> vsfList.BottomRow Then
                    For i = vsfList.TopRow To vsfList.BottomRow
                        Call RefreshRowRelation(i)
                    Next
                End If
                
                '��ͳ��
                Call ColStatistics(mrsDataShow)
                
                Call DoListCfg
                
                Call LoadListHeadCfg(mlngAdviceID)
            Else
                Call ColStatistics(mrsData)
                Call LoadListHeadCfg(0)
            End If
    
            Call ResetSort(mlngSortCol, mintSortOrder)
        Else
            lngAdviceID = GetSelectRowAdviceID
            mblSearching = True
            Set vsfList.DataSource = mrsData
            mblSearching = False
            '��ͳ��
            Call ColStatistics(mrsData)
            Call LoadListHeadCfg(0)
        End If
    End If

    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[ExecuteQuery]" & vbCrLf & Err.Description
End Function

Public Function ExecuteWithLink(ByVal strSql As String) As Boolean
'�ղع���
On Error GoTo errH
    Set mrsData = mObjQuery.ExecuteWithLink(strSql)
    
    If Not mrsData Is Nothing Then
        If mrsData.RecordCount > 0 Then
            Set mrsDataShow = GetFilterFromQuickFilter
            Set mrsDataShow = mObjQuery.DataConvert(mrsDataShow, mlngSchemeNo)
            If Not mrsDataShow Is Nothing Then
                mblSearching = True
                Set vsfList.DataSource = mrsDataShow
                mblSearching = False
                '��ͳ��
                Call ColStatistics(mrsDataShow)
                Call LoadListHeadCfg(0)
            Else
                Call LoadListHeadCfg(0)
            End If
        End If
    Else
        Call ColStatistics(mrsData)
        Call LoadListHeadCfg(0)
    End If
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[ExecuteWithLink]" & vbCrLf & Err.Description
End Function

Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH
    Dim strCfgNew As String
    Dim blRaiseEvent As Boolean '��������pacsMain�����Ҽ��˵�
    
    blRaiseEvent = True
    
    If mintShowType = 1 Then Exit Sub
    If vsfList.MouseRow = 0 And Button = 2 Then
        '��ó�ʼ���ô� �� ���Ի����ô�
        Call frmVsfColsList.ShowVsfColsListWindow(mstrSchemeCfg.strListCfgDefault, mstrSchemeCfg.strListCfg, mfrmParent)
        strCfgNew = frmVsfColsList.GetListCfg
        
        If Len(strCfgNew) > 0 Then
            mstrSchemeCfg.strListCfg = strCfgNew
            Call LoadListHeadCfg(-1)
        End If
        
        '�˴���Ҫ�����������֤�й�������ȷ
        Call DoBeforeFilterData(True)
        
        blRaiseEvent = False
    End If
    

    If blRaiseEvent Then RaiseEvent OnMouseUp(Button, Shift, X, Y)

    Exit Sub
errH:
    MsgBox "[vsfList_MouseUp]" & vbCrLf & Err.Description, vbOKOnly, "�쳣"
End Sub

Private Function InitCardType(ByVal strCardNames As String) As String
'��ָ����ʽ��ʼ��������
On Error GoTo errH
    Dim i As Integer
    Dim aryKindInfo() As String
    Dim strKinds As String
    
    aryKindInfo = Split(strCardNames, ";")
    
    strKinds = ""
    For i = 0 To UBound(aryKindInfo) - 1
        If strKinds <> "" Then strKinds = strKinds & ";"
        strKinds = strKinds & aryKindInfo(i) & "|" & aryKindInfo(i) & "|-1"
    Next i
    
    InitCardType = strKinds & ";"
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[InitCardType]" & vbCrLf & Err.Description
End Function

Private Function GetFilter() As String
'��ָ����ʽ��ʼ��������
'�������ָ�ʽ   ��    "����|ʱ��|ϵͳʱ��|����|���"
On Error GoTo errH
    Dim i As Integer
    GetFilter = ""
    If mSqlScheme Is Nothing Then Exit Function
    
    For i = 1 To mSqlScheme.SerachCfgCount
        
        If mSqlScheme.SerachCfg(i).InputType = itPopup Or mSqlScheme.SerachCfg(i).InputType = itBoth Then
            GetFilter = GetFilter & mSqlScheme.SerachCfg(i).Name & ";"
        End If
    Next

    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetFilter]" & vbCrLf & Err.Description
End Function

Private Sub SeekNextPati(ByVal blnFirst As Boolean, ByVal strName As String, _
    ByVal strFilter As String, Optional blnIsReSeek As Boolean = False)
On Error GoTo errH
'------------------------------------------------
'���ܣ��ڲ����б��ж�λָ���ļ�¼
'������ blnFirst -- �Ƿ��һ�β���
'���أ��ޣ�ֱ���ڲ����б��ж�λ
'------------------------------------------------
    Dim i As Long
    Dim intB As Integer
    Dim lngEndRow As Long
    Dim lngSelRow As Long
    Dim strTemp As String
    Dim lngRowIndex As Long

    
    '���û�м�¼�����˳�
    If mDataGrid.Rows - 1 <= 0 Then Exit Sub

    intB = 0
    lngRowIndex = -1

    If Not blnFirst Then
        intB = mDataGrid.Row + 1
        If intB >= mDataGrid.Rows Then intB = 1
    End If
'
    lngSelRow = mDataGrid.Row
    lngEndRow = mDataGrid.Rows - 1

continue1:
   '�����ֶ�����λ�Ĵ���
    If mDataGrid.ColIndex(strName) > 0 Then
        lngRowIndex = mDataGrid.FindRow(strFilter, intB, mDataGrid.ColIndex(strName), False, False)
    End If

    If lngRowIndex > 0 Then
        patiSearch.Tag = patiSearch.Text
        mDataGrid.Row = lngRowIndex

        If mDataGrid.TopRow > mDataGrid.Row Then mDataGrid.TopRow = mDataGrid.Row
        If mDataGrid.BottomRow - 1 < mDataGrid.Row Then
            mDataGrid.TopRow = mDataGrid.TopRow + (mDataGrid.Row - mDataGrid.BottomRow) + 1
        End If
    Else
        If intB > 1 Then
            intB = 0
            GoTo continue1:
        End If
        
    End If

    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[SeekNextPati]" & vbCrLf & Err.Description
End Sub

Public Function UpdateRow(ByVal blIsAdd As Boolean, ByVal lngAdviceID As Long) As Boolean
'����ҽ��ID,���²�ѯһ�����ݲ���ˢ���б���һ��
'strCol: ������ VarValueֵ  blIsAdd �Ƿ�������  ҽ��ID,lngAdviceIDֻ����������Ҫ
'�����ж��б��Ƿ�����˱����µ��У�����������Ҫˢ������ʾ
'���¼�¼����Ӧ������
On Error GoTo errH
    Dim rsTemp As ADODB.Recordset
    Dim rsTempShow As ADODB.Recordset
    Dim strColName As String
    Dim i As Integer
    Dim lngAdviceIDOld As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strTmp As String
    
    UpdateRow = False
    If mObjQuery Is Nothing Then Exit Function

    With mObjQuery
'        mTqueryType = ����һ��
        Set rsTemp = .ExecuteWithAttach("[ϵͳ.ҽ��ID]", lngAdviceID)
        
        Set rsTempShow = CopyRecordSet(rsTemp)
        Set rsTempShow = mObjQuery.DataConvert(rsTempShow, mlngSchemeNo)
        Call rsTemp.MoveFirst
        
        If Not blIsAdd Then

            mrsData.MoveFirst

            While Not mrsData.EOF
                If Val(mrsData!ҽ��ID) = lngAdviceID Then

                    For i = 0 To rsTemp.Fields.Count - 1
                        mrsData.Fields(i).Value = rsTemp.Fields(i).Value
                    Next
                    
                    mrsDataShow.MoveFirst
                    While Not mrsDataShow.EOF
                        If Val(mrsDataShow!ҽ��ID) = lngAdviceID Then

                            For i = 0 To rsTempShow.Fields.Count - 1
                                mrsDataShow.Fields(i).Value = rsTempShow.Fields(i).Value
                            Next

                            GoTo Refresh
                        End If
                        mrsDataShow.MoveNext
                    Wend

                End If
                mrsData.MoveNext
            Wend
  
        Else
            'lsq����֤��Ч��
            mrsData.AddNew

            For i = 0 To rsTemp.Fields.Count - 1
                mrsData.Fields(i) = rsTemp.Fields(i)
            Next

        End If

    End With

Refresh:
    '����б���ʾ����
    If blIsAdd Then
        'ֱ��ˢ�������б�
        If Not mrsData Is Nothing Then
            If mrsData.RecordCount > 0 Then
                Set mrsDataShow = GetFilterFromQuickFilter
                Set mrsDataShow = mObjQuery.DataConvert(mrsDataShow, mlngSchemeNo)
                
                If Not mrsDataShow Is Nothing Then
                    lngAdviceIDOld = GetSelectRowAdviceID
                    mblSearching = True
                    Set vsfList.DataSource = mrsDataShow
                    mblSearching = False
                    '��ͳ��
                    Call ColStatistics(mrsDataShow)
                    
                    lngRow = vsfList.FindRow(lngAdviceID, 1, vsfList.ColIndex(mstrListKeyCol))
                    If lngRow = -1 Then lngRow = 0
                    Call LoadListHeadCfg(lngRow)
                Else
                    Call LoadListHeadCfg(0)
                End If

            End If
        End If
    Else
        lngRow = vsfList.FindRow(lngAdviceID, 1, vsfList.ColIndex(mstrListKeyCol))
        If lngRow = -1 Then
            UpdateRow = True
            Exit Function
        End If
                    
        '����һ�е�����
        For i = 1 To vsfList.Cols - 1
            strColName = vsfList.TextMatrix(0, i)
            vsfList.TextMatrix(lngRow, i) = NVL(rsTempShow.Fields(strColName).Value)
        Next

        Call RefreshRowRelation(lngRow)
    End If
    
    UpdateRow = True
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[UpdateRow]" & vbCrLf & Err.Description
End Function

Private Function GetQuickFilterSQLPar(ByVal lngIndex As Long) As String
'��ȡ������Զ�����ٹ�������������"ͷ,˫��,��״��"���֣����ڹ���
'������ ������Ϣindex���˵�ID
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngID As Long, lngIDEnd As Long
    
    Dim strTmp As String
    Dim cbrPopControl As CommandBarControl
    Dim objControl As CommandBarControl
    Dim blSimpleFilter As Boolean
    Dim blChooseOne As Boolean
    Dim strAry() As String
    
    blSimpleFilter = mTQuickFilterState.TCmdState(lngIndex).blSimpleFilter
    lngID = 100 * lngIndex
    lngIDEnd = mTQuickFilterState.TCmdState(lngIndex).intItemCount
    blChooseOne = False
    
    For i = 1 To lngIDEnd
        Set objControl = cbrFilter.FindControl(, lngID - 100 + i, , True)
        
        If mTQuickFilterState.TCmdState(lngIndex).cmdItem(i).blChoose Then
            blChooseOne = True
            If blSimpleFilter Then
                If Len(strTmp) > 0 Then strTmp = strTmp & ","
                strTmp = strTmp & objControl.Caption
            Else
                If Len(strTmp) > 0 Then strTmp = strTmp & ","
                strTmp = strTmp & objControl.Category
            End If
        End If
        
    Next
    
    If blChooseOne = False Then
        For i = 1 To lngIDEnd
            Set objControl = cbrFilter.FindControl(, lngID - 100 + i, , True)
            If blSimpleFilter Then
                If Len(strTmp) > 0 Then strTmp = strTmp & ","
                strTmp = strTmp & objControl.Caption
            Else
                If Len(strTmp) > 0 Then strTmp = strTmp & ","
                strTmp = strTmp & objControl.Category
            End If
        Next
    End If
    
    strAry = Split(strTmp, ",")
    strTmp = ""
    For i = 0 To UBound(strAry)
        For j = i + 1 To UBound(strAry)
            If strAry(i) = strAry(j) Then strAry(j) = ""
        Next
        
        If strAry(i) <> "" Then
            If Len(strTmp) > 0 Then strTmp = strTmp & ","
            strTmp = strTmp & strAry(i)
        End If
    Next
    
    strTmp = Replace(strTmp, ",,,,,,", ",")
    strTmp = Replace(strTmp, ",,,,,", ",")
    strTmp = Replace(strTmp, ",,,,", ",")
    strTmp = Replace(strTmp, ",,,", ",")
    strTmp = Replace(strTmp, ",,", ",")
    
    mTQuickFilterState.TCmdState(lngIndex).strRelationValueForVBSFilter = Trim(strTmp)

    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetQuickFilterSQLPar]" & vbCrLf & Err.Description
End Function

Private Function CbrFilterDeal(ByVal strFilterOld As String, ByVal strFilterNew As String) As String
On Error GoTo errH
    Dim i As Integer
    Dim strNew() As String
    Dim strOld As String
    
    strOld = "," & strFilterOld & ","
    
    strNew = Split(strFilterNew, ",")
    
    For i = 0 To UBound(strNew)
        If InStr(strOld, "," & strNew(i) & ",") = 0 Then
            If Len(CbrFilterDeal) > 0 Then CbrFilterDeal = CbrFilterDeal & ","
            CbrFilterDeal = CbrFilterDeal & strNew(i)
        End If
    Next
    
    CbrFilterDeal = "," & CbrFilterDeal
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[CbrFilterDeal]" & vbCrLf & Err.Description
End Function

Private Function DoBeforeFilterData(ByVal blOnlyChangeCol As Boolean) As Boolean
'��ѯ֮�������ͷ�ṹ�ı��DataSource֮ǰ�Ĵ���
'blChangeCol �Ƿ�ֻ�Ǹı�Col��ȡ�˳�����ֲ��������ǣ�����Ҫ���list
On Error GoTo errH
    Dim i As Integer, j As Integer
    
    If mrsData Is Nothing Or mSqlScheme Is Nothing Then Exit Function
    
    If blOnlyChangeCol Then
        With vsfList
            ReDim mColCfgInfo(.Cols - 1)
            
            For i = 1 To .Cols - 1
                .ColKey(i) = .TextMatrix(0, i)
                
                For j = 1 To mSqlScheme.ShowCfgCount
                    mColCfgInfo(i) = 0
                    If mSqlScheme.ShowCfg(j).Name = .ColKey(i) Then
                        mColCfgInfo(i) = j
                        Exit For
                    End If
                Next
            Next
        End With
    Else
        ReDim mColCfgInfo(mrsData.Fields.Count)
        
        With vsfList
            If Not blOnlyChangeCol Then
                .Clear
                .Cols = mrsData.Fields.Count + 1
            End If
            
            For i = 1 To mrsData.Fields.Count
                .ColKey(i) = mrsData.Fields(i - 1).Name
                
                For j = 1 To mSqlScheme.ShowCfgCount
                    mColCfgInfo(i) = 0
                    If mSqlScheme.ShowCfg(j).Name = .ColKey(i) Then
                        mColCfgInfo(i) = j
                        Exit For
                    End If
                Next
            Next
        End With
    End If
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[DoBeforeFilterData]" & vbCrLf & Err.Description
End Function

Private Function GetDefaultColCfg() As Boolean
'��ȡ��ʼ��������Ϣ�����ƣ��Ƿ����أ����ڻָ���״̬
On Error GoTo errH
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim strCol() As String
    Dim picLoad As StdPicture
    Dim blHaveCfg As Boolean

    If mSqlScheme Is Nothing Or Len(mstrSchemeCfg.strListCfgDefault) > 0 Then Exit Function

    '���һ��û������
    strCol = Split(mstrSchemeCfg.strListCfgDefaultColOrder, "|")
    If mSqlScheme.ShowCfgCount > 0 Then
        For i = 0 To UBound(strCol) - 1   'for 1
            For j = 1 To mSqlScheme.ShowCfgCount   'for 2
                blHaveCfg = False
                If strCol(i) = mSqlScheme.ShowCfg(j).Name Then
                    blHaveCfg = True
    '                    ������
                    If mSqlScheme.ShowCfg(j).HiddenCol Then
                        strCol(i) = strCol(i) & "," & TextWidth(strCol(i)) & "," & "0"
                    Else
                        strCol(i) = strCol(i) & "," & TextWidth(strCol(i)) & "," & "1"
                    End If
                    
                    Exit For  '����for 2
                End If
                            
            Next ' for 2
            
            '��ִ�е�����˵��û�ж�Ӧ������������
            If Not blHaveCfg Then
                If Len(strCol(i)) > 0 Then strCol(i) = strCol(i) & "," & TextWidth(strCol(i)) & "," & "1"
            End If
    
        Next ' for 1
    Else
        For i = 0 To UBound(strCol) - 1   'for 1
            strCol(i) = strCol(i) & "," & TextWidth(strCol(i)) & "," & "1"
        Next ' for 1
    End If
    

    For i = 0 To UBound(strCol) - 1
        If Len(mstrSchemeCfg.strListCfgDefault) > 0 Then mstrSchemeCfg.strListCfgDefault = mstrSchemeCfg.strListCfgDefault & "|"
        mstrSchemeCfg.strListCfgDefault = mstrSchemeCfg.strListCfgDefault & strCol(i)
    Next
        
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetDefaultColCfg]" & vbCrLf & Err.Description
End Function


Private Function DoListCfg() As Boolean
'�����б�����
'��ʱlist���Ѿ������ݣ���Ҫ������ͷ���ã�˳��ı�Ȳ���
'���������ã���������ʾ��Ϣ��ͼ�꣬�Ƿ����صȣ�
On Error GoTo errH
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim strCol() As String
    Dim picLoad As StdPicture
    Dim ObjScShowCfg As New clsScShowCfg
    Dim blHaveCfg As Boolean

    If mSqlScheme Is Nothing Then Exit Function
    
    If Len(mstrSchemeCfg.strListCfgDefault) = 0 Then
        Call GetDefaultColCfg
    End If
        
    strCol = Split(mstrSchemeCfg.strListCfgDefaultColOrder, "|")
    
    'i:��ʼ�������ֶ����  j:����ʾ�������
    With vsfList
        For i = 1 To UBound(strCol) - 1   'for 1
            
            '�������ý���һЩ����
            If mColCfgInfo(i) > 0 Then
                Set ObjScShowCfg = mSqlScheme.ShowCfg(mColCfgInfo(i))
                           
                If Len(ObjScShowCfg.Icon) > 0 Then
                    '��ͼ�����
                    Set picLoad = GetIcon(ObjScShowCfg.Icon)
                    Set .Cell(flexcpPicture, 0, i) = GetIcon(ObjScShowCfg.Icon)
        '
        '                '��Ҫ��Ч����СͼƬ�ķ�ʽ[LSQB2]
        ''                If imgList16.ImageHeight > .RowHeight(0) Then .RowHeight(0) = imgList16.ImageHeight
        '
        '                If picLoad.Height > .RowHeight(0) Then
        '                    .RowHeight(0) = picLoad.Height
        '                End If
                End If
                
                '����������ʾ ͳһ���� ��ǰ�����ӿո�
                If ObjScShowCfg.HiddenData Then
                
                    For j = 1 To .Rows - 1
                        .Cell(flexcpText, j, i) = "                                " & .Cell(flexcpText, j, i)
                        .Cell(flexcpAlignment, j, i) = flexAlignLeftTop
                    
                    Next
                End If
                
            End If
          
        Next ' for 1
    End With
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[DoListCfg]" & vbCrLf & Err.Description
End Function

Private Function GetIcon(ByVal strID As String) As StdPicture
'ͨ��ID��ȡͼ�꣬�ж��ֵ����Ƿ��Ѿ����ڸ�ͼ�꣬�����ڣ�ʹ���ֵ��ж����������ڡ�����mObjQuery.GetIconRes������ӵ��ֵ�
On Error GoTo errH
    Dim stdPic As StdPicture
    
    If mPicDictionary.Exists(strID) Then
        Set GetIcon = mPicDictionary.Item(strID)
    Else
        Set stdPic = mObjQuery.GetIconRes(strID)
        Call mPicDictionary.Add(strID, stdPic)
        Set GetIcon = stdPic
    End If
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetIcon]" & vbCrLf & Err.Description
End Function

Private Sub ResetSort(ByVal lngCol As Long, ByVal lngWay As Long)
'��������
On Error GoTo errH
    Dim RowIndex As Long
    
    If vsfList.Rows <= 1 Then Exit Sub
    
    If vsfList.Col <> vsfList.ColIndex(GetColSort(vsfList.ColKey(lngCol))) Then
        vsfList.Col = vsfList.ColIndex(GetColSort(vsfList.ColKey(lngCol)))
        
        '����  ��  ż�� ���������������
        If lngWay = 2 Or lngWay = 4 Or lngWay = 6 Or lngWay = 8 Then
            vsfList.Sort = 4
        Else
            vsfList.Sort = 3
        End If
    Else
        vsfList.Col = lngCol
        vsfList.Sort = lngWay
    End If
    
    If vsfList.TopRow = vsfList.BottomRow Then Exit Sub
    For RowIndex = vsfList.TopRow To vsfList.BottomRow
        Call RefreshRowRelation(RowIndex)
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[ResetSort]" & vbCrLf & Err.Description
End Sub

Private Sub GetSchemePara()
'��ȡ����������������Ӱ�������ʾ�����ã�
    mTLayout.blShowHistory = mSqlScheme.ShowHistory
    mTLayout.blShowQuickFilter = mSqlScheme.FilterCfgCount > 0
End Sub

Private Sub GetLocalPara()
'��ȡ���ز�����ע��������
    mlngSortCol = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "������", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "������", 0))
    mlngMove = Val(GetSetting("ZLSOFT", "˽��ģ��\" & mstrDBUser & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "�б�����Ϣ�߶�����", 0))
    
    mTPatiIdentifyInfo.strFindItem = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "������Ŀ")
    mTPatiIdentifyInfo.strLocateItem = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "��λ��Ŀ")
    
    mTPatiIdentifyInfo.blFind = (GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "�Ƿ����", "1") = "1")
End Sub

Private Sub SaveLocalPara()
'���ñ��ز���
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & mstrDBUser & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "�б�����Ϣ�߶�����", mlngMove)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "������", mlngSortCol)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "������", mintSortOrder)
    
    Call SaveLocalPara_PatiIdentify
End Sub

Private Sub SaveLocalPara_PatiIdentify()
'����Pati�ؼ���ز���
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "������Ŀ", mTPatiIdentifyInfo.strFindItem)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "��λ��Ŀ", mTPatiIdentifyInfo.strLocateItem)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name & "\" & mlngSchemeNo & "\", "�Ƿ����", IIf(mTPatiIdentifyInfo.blFind, 1, 0))
End Sub



Private Sub LoadPatiIdentifyInfo()
'��һ�μ��ط�������Ҫ ��ȡ����/��λѡ��������ֱ��ʹ��
'���ҹ��ܣ�Ĭ������+��������
'��λ���ܣ��б�����
'ͬʱ�ж��Ƿ���Ҫ��ʾPati�ؼ�
On Error GoTo errH
    Dim i As Integer
    Dim blHaveFind As Boolean '�Ƿ��в�����Ŀ
    Dim blHaveLocate As Boolean '�Ƿ��ж�λ��Ŀ
    Dim strSql As String
    
    blHaveFind = False
    blHaveLocate = False
    
    '����δ���ع����ȼ���
    If Not mTPatiIdentifyInfo.blHaveLoad Then
        
        For i = 1 To mSqlScheme.SerachCfgCount
        
            If mSqlScheme.SerachCfg(i).InputType = itFast Or mSqlScheme.SerachCfg(i).InputType = itBoth Then
                mTPatiIdentifyInfo.strFindItems = mTPatiIdentifyInfo.strFindItems & mSqlScheme.SerachCfg(i).Name & ";"
                blHaveFind = True
            End If
            
        Next
        
        If mTPatiIdentifyInfo.strFindItems = "" Then mTPatiIdentifyInfo.strFindItems = "����;"
        
        For i = 1 To mSqlScheme.ShowCfgCount
            If mSqlScheme.ShowCfg(i).UseListLocate Then
                mTPatiIdentifyInfo.strLocateItems = mTPatiIdentifyInfo.strLocateItems & mSqlScheme.ShowCfg(i).Name & ";"
                blHaveLocate = True
            End If
        Next
        
        If mTPatiIdentifyInfo.strLocateItems = "" Then mTPatiIdentifyInfo.strLocateItems = "����;"
        
        mTPatiIdentifyInfo.blHaveLoad = True
    
        If blHaveFind Or blHaveLocate Then mTPatiIdentifyInfo.blShowPatiIdentify = True
        
        '�ж��Ƿ���ʱ�䷶Χ����
        strSql = mSqlScheme.GetScheme
        If InStr(strSql, "[ϵͳ.��ʼ����]") > 0 And InStr(strSql, "[ϵͳ.��������]") Then
            mTLayout.blShowTimeSelect = True
        Else
            mTLayout.blShowTimeSelect = False
        End If
        
        mTLayout.blShowBaseFilter = mTPatiIdentifyInfo.blShowPatiIdentify Or mTLayout.blShowTimeSelect
    End If

    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[LoadPatiIdentifyInfo]" & vbCrLf & Err.Description
End Sub

Public Sub RefreshRowRelation(ByVal lngRow As Long)
'����ˢ��ĳһ�У�ͬʱ�����й�������
On Error GoTo errH
    Dim Value As String
    Dim i As Integer
    
    For i = 1 To vsfList.Cols - 1
        Value = vsfList.TextMatrix(lngRow, i)
        Call RowRelationConvert(lngRow, i, Value)
    Next
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[RefreshRowRelation]" & vbCrLf & Err.Description
    
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    mbytFontSize = bytFontSize
End Sub

Public Sub ReSetFormFontSize(Optional bytFontSize As Byte = 0)
On Error Resume Next
    
    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType As String
    
    If bytFontSize > 0 Then mbytFontSize = bytFontSize
    
    Me.FontSize = mbytFontSize
    Set CtlFont = New StdFont
    strFontType = IIf(IsUseClearType = True, "΢���ź�", "����")
    CtlFont.Name = strFontType
    CtlFont.Size = mbytFontSize
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") 'ҳ��ؼ�
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            If objCtrl.Name <> "lblCash" Then
                objCtrl.Font.Name = strFontType
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("��") + 60
            End If
        Case UCase("vsFlexGrid")
            objCtrl.Font = CtlFont
        Case UCase("ComboBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("�޹�" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("��") * 1.5
        Case UCase("textBox")
          objCtrl.FontName = strFontType
          objCtrl.FontSize = mbytFontSize
        Case UCase("ReportControl")
            
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            
            Set objCtrl.Options.Font = CtlFont
            
        Case UCase("TabControl")
            Set objCtrl.PaintManager.Font = CtlFont
            
        Case UCase("CommandButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            
        Case UCase("PatiIdentify")
            objCtrl.CardNoShowFont.Size = mbytFontSize
            objCtrl.Font.Size = mbytFontSize
            objCtrl.IDKindFont.Size = mbytFontSize
            If mbytFontSize = 9 Then
                objCtrl.Height = 330
            ElseIf mbytFontSize = 12 Then
                objCtrl.Height = 360
            ElseIf mbytFontSize = 15 Then
                objCtrl.Height = 390
            End If
            objCtrl.Refrash
        
        End Select
    Next
    
    Call AdjustFace(mbytFontSize)
    
End Sub

Private Sub AdjustFace(ByVal bytFontSize As Byte)
'�ֺ� Ŀǰ����վ֧��9,12,15����
On Error Resume Next
    Dim lngHeightҳǩ As Long
    Dim lngHeight�������� As Long
    Dim lngHeight���ٹ��� As Long
    Dim lngHeight�б� As Long
    Dim lngHeight��ʷ��� As Long
    Dim lngHeight������Ϣ As Long
    Dim lngHeight��ϸ��Ϣ As Long
    Dim lngHeight�ָ��� As Long
    
    '����� 10000  1000�Ǵ�Ź涨�ķָ�����Ч�ƶ���Χ
    If mlngMove > 10000 Then mlngMove = 10000
    If mlngMove < -1000 Then mlngMove = -1000
    
    lngHeight�ָ��� = 50
    
    If bytFontSize = 9 Then
        lngHeightҳǩ = 300
        
        lngHeight�������� = IIf(mTLayout.blShowBaseFilter, 350, 0)
        lngHeight���ٹ��� = IIf(mTLayout.blShowQuickFilter, 400, 0)
        
        lngHeight��ʷ��� = IIf(Label1.Height > cboHistory.Height, Label1.Height, cboHistory.Height) + 60
        lngHeight��ʷ��� = IIf(mTLayout.blShowHistory, lngHeight��ʷ���, 0)
        
        lngHeight������Ϣ = labPatientInfoName.Height + 90
    
        lngHeight��ϸ��Ϣ = C_LAYOUT_BASEHEIGHTOFDETAILINFO + mlngMove
        
        lngHeight�б� = Me.ScaleHeight - lngHeightҳǩ - lngHeight�������� - lngHeight���ٹ��� - lngHeight��ʷ��� - lngHeight������Ϣ - lngHeight��ϸ��Ϣ
    
        Call tabQuery.Move(C_LAYOUT_LISTLEFT, 0, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeightҳǩ)
    
        Call picSearch.Move(C_LAYOUT_LISTLEFT, tabQuery.Top + tabQuery.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��������)
    
        Call picFilter.Move(C_LAYOUT_LISTLEFT, picSearch.Top + picSearch.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight���ٹ���)
    
        Call picVsf.Move(C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight�б�)
        
        Call PicLine.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight�ָ���)
    
        Call picHistory.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��ʷ���)
    
        Call picListRowInfo.Move(C_LAYOUT_LISTLEFT, picHistory.Top + picHistory.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight������Ϣ)
        Call txtDetail.Move(C_LAYOUT_LISTLEFT, picListRowInfo.Top + picListRowInfo.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��ϸ��Ϣ)
    
    ElseIf bytFontSize = 12 Then
        lngHeightҳǩ = 360
        
        lngHeight�������� = IIf(mTLayout.blShowBaseFilter, 390, 0)
        lngHeight���ٹ��� = IIf(mTLayout.blShowQuickFilter, 420, 0)
        
        lngHeight��ʷ��� = IIf(Label1.Height > cboHistory.Height, Label1.Height, cboHistory.Height) + 60
        lngHeight��ʷ��� = IIf(mTLayout.blShowHistory, lngHeight��ʷ���, 0)
        
        lngHeight������Ϣ = labPatientInfoName.Height + 90
    
        lngHeight��ϸ��Ϣ = C_LAYOUT_BASEHEIGHTOFDETAILINFO + mlngMove
        
        lngHeight�б� = Me.ScaleHeight - lngHeightҳǩ - lngHeight�������� - lngHeight���ٹ��� - lngHeight��ʷ��� - lngHeight������Ϣ - lngHeight��ϸ��Ϣ
    
        Call tabQuery.Move(C_LAYOUT_LISTLEFT, 0, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeightҳǩ)
    
        Call picSearch.Move(C_LAYOUT_LISTLEFT, tabQuery.Top + tabQuery.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��������)
    
        Call picFilter.Move(C_LAYOUT_LISTLEFT, picSearch.Top + picSearch.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight���ٹ���)
    
        Call picVsf.Move(C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight�б�)
        
        Call PicLine.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight�ָ���)
    
        Call picHistory.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��ʷ���)
    
        Call picListRowInfo.Move(C_LAYOUT_LISTLEFT, picHistory.Top + picHistory.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight������Ϣ)
        Call txtDetail.Move(C_LAYOUT_LISTLEFT, picListRowInfo.Top + picListRowInfo.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��ϸ��Ϣ)
    
    ElseIf bytFontSize = 15 Then
        lngHeightҳǩ = 420
        
        lngHeight�������� = IIf(mTLayout.blShowBaseFilter, 430, 0)
        lngHeight���ٹ��� = IIf(mTLayout.blShowQuickFilter, 440, 0)
        
        lngHeight��ʷ��� = IIf(Label1.Height > cboHistory.Height, Label1.Height, cboHistory.Height) + 60
        lngHeight��ʷ��� = IIf(mTLayout.blShowHistory, lngHeight��ʷ���, 0)
        
        lngHeight������Ϣ = labPatientInfoName.Height + 90
    
        lngHeight��ϸ��Ϣ = C_LAYOUT_BASEHEIGHTOFDETAILINFO + mlngMove
        
        lngHeight�б� = Me.ScaleHeight - lngHeightҳǩ - lngHeight�������� - lngHeight���ٹ��� - lngHeight��ʷ��� - lngHeight������Ϣ - lngHeight��ϸ��Ϣ
    
        Call tabQuery.Move(C_LAYOUT_LISTLEFT, 0, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeightҳǩ)
    
        Call picSearch.Move(C_LAYOUT_LISTLEFT, tabQuery.Top + tabQuery.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��������)
    
        Call picFilter.Move(C_LAYOUT_LISTLEFT, picSearch.Top + picSearch.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight���ٹ���)
    
        Call picVsf.Move(C_LAYOUT_LISTLEFT, picFilter.Top + picFilter.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight�б�)
        
        Call PicLine.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight�ָ���)
    
        Call picHistory.Move(C_LAYOUT_LISTLEFT, picVsf.Top + picVsf.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��ʷ���)
    
        Call picListRowInfo.Move(C_LAYOUT_LISTLEFT, picHistory.Top + picHistory.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight������Ϣ)
        Call txtDetail.Move(C_LAYOUT_LISTLEFT, picListRowInfo.Top + picListRowInfo.Height, Me.Width - 2 * C_LAYOUT_LISTLEFT, lngHeight��ϸ��Ϣ)
    End If
    
End Sub

Private Sub ColStatistics(rsData As Recordset)
'��ͳ�ƴ���
On Error GoTo errHandle
    Dim i As Long, j As Long, k As Long
    Dim lngColICount As Long '��ͳ�Ƶ��������
    Dim lngColIndex() As Integer
    
    Dim strColName As String '"��Ҫ��������е���" ���ƣ� "Ӱ������;������;�Ƿ�����"
    Dim strColNameAll As String
    Dim strStateBarInfo As String '����ͨ���¼����ݵ���ͳ����Ϣ
    Dim strInfoTmp As String
    
    Dim objTColTotalInfo() As TColTotalInfo
    Dim DictColTotal() As Dictionary
    
    strColNameAll = ""
    strColName = ""
    lngColICount = 0
    
    If mSqlScheme Is Nothing Or rsData Is Nothing Then
        RaiseEvent OnColStatistics("")
        Exit Sub
    End If
    
    If rsData.RecordCount = 0 Then
        RaiseEvent OnColStatistics("")
        Exit Sub
    End If
     
    With mSqlScheme
        For i = 1 To .ShowCfgCount
            If .ShowCfg(i).IsTotal Then
                lngColICount = lngColICount + 1
                strColNameAll = strColNameAll & .ShowCfg(i).Name & ";"
                ReDim Preserve DictColTotal(lngColICount)
                Set DictColTotal(lngColICount) = New Dictionary
            End If
        Next
    End With
        
    With rsData
        
        strStateBarInfo = ""
        strColName = ""
        
        For k = 1 To lngColICount
        
            '��ȡ�ֶ���
            strColName = Split(strColNameAll, ";")(k - 1)

            .MoveFirst
            While Not .EOF
                If Not IsNull(.Fields(strColName).Value) Then
                    If Len(Trim(.Fields(strColName).Value)) > 0 Then
                        If DictColTotal(k).Exists(.Fields(strColName).Value) Then
                             DictColTotal(k).Item(.Fields(strColName).Value) = DictColTotal(k).Item(.Fields(strColName).Value) + 1
                        Else
                            Call DictColTotal(k).Add(.Fields(strColName).Value, 1)
                        End If
                        
                    End If
                End If
            .MoveNext
            Wend
            
            strInfoTmp = "[" & strColName & "]:"
            For j = 1 To DictColTotal(k).Count
                strInfoTmp = strInfoTmp & DictColTotal(k).Keys(j - 1) & ":" & DictColTotal(k).Item(DictColTotal(k).Keys(j - 1)) & " "
            Next
            
            If Len(strStateBarInfo) > 0 Then strStateBarInfo = strStateBarInfo & "|"
            strStateBarInfo = strStateBarInfo & strInfoTmp
            
            RaiseEvent OnColStatistics(strStateBarInfo)
                
        Next
   
        .MoveFirst
    End With
    
    RaiseEvent OnColStatistics(strStateBarInfo)
    Exit Sub
errHandle:
    Err.Raise -1, "frmPacsQuery", "[ColStatistics]" & vbCrLf & Err.Description
End Sub

Private Sub RowRelationConvert(ByVal Row As Long, ByVal Col As Long, Value As String)
'�й������� ��Ҫ���� ��ɫ��ͼ�� ÿ��ֻ��Ҫ����һ������ɫ
On Error GoTo errH
    Dim i As Integer, j As Integer, k As Integer
    Dim lngRelationIndex As Long
    Dim objClsRelation As New clsScRowRelation
    Dim blContinue As Boolean '�Ƿ����
    Dim lngColColor As Long '����ɫ��
    Dim strColorColValue As String '����ɫ������ �� "�ѵǼ�"
    Static TpRowColorInfo As TRowColorInfo
       
    lngRelationIndex = mColCfgInfo(Col)
    If UBound(mColCfgInfo) < Col - 1 Then Exit Sub
    
    If lngRelationIndex < 1 Then
        If Col <> 1 Then
            Exit Sub
        End If
    Else
        If mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount < 1 Then
            If Col <> 1 Then
                Exit Sub
            End If
        End If
    End If

    If Col = 1 Then
        '�״μ��ط������ȡ����ɫ�����
        
        If TpRowColorInfo.LngSchemeNo <> mlngSchemeNo Then
            TpRowColorInfo.blHaveRowColor = False
            TpRowColorInfo.LngSchemeNo = mlngSchemeNo
            
            For i = 1 To mSqlScheme.ShowCfgCount
            
                For j = 1 To mSqlScheme.ShowCfg(i).RowRelationCount
                    If mSqlScheme.ShowCfg(i).RowRelation(j).RowBackColor > 0 Or mSqlScheme.ShowCfg(i).RowRelation(j).RowFontColor > 0 Then
                        TpRowColorInfo.intRowColorIndex = i
                        TpRowColorInfo.blHaveRowColor = True
                        Exit For
                    End If
                Next
                
            Next
    
        End If
        
        If TpRowColorInfo.blHaveRowColor Then
        '��������ɫ����Ҫ��������Ĵ���
            If TpRowColorInfo.intRowColorIndex > 0 Then
                With vsfList
                    '���ȴ����й�������ɫ
                    For i = 1 To mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).RowRelationCount
                        Set objClsRelation = mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).RowRelation(i)
    
                        lngColColor = .ColIndex(mSqlScheme.ShowCfg(TpRowColorInfo.intRowColorIndex).Name)
                        strColorColValue = .TextMatrix(Row, lngColColor)
                        If strColorColValue = objClsRelation.TiggerData Then
                            '�б���ɫ
                            If objClsRelation.RowBackColor > 0 Then .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = objClsRelation.RowBackColor
                            '��ǰ��ɫ
                            If objClsRelation.RowFontColor > 0 Then .Cell(flexcpForeColor, Row, 1, Row, .Cols - 1) = objClsRelation.RowFontColor
                        End If
                    Next
                End With
                
                mrsDataShow.MoveFirst
            End If
        End If
        
    End If
    
    'lngRelationIndex <1 ˵������Ҫ����Ĵ���
    If lngRelationIndex < 1 Then Exit Sub

    With vsfList
        '���ȴ����й���
        If mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount > 0 Then
            For i = 1 To mSqlScheme.ShowCfg(lngRelationIndex).RowRelationCount
                Set objClsRelation = mSqlScheme.ShowCfg(lngRelationIndex).RowRelation(i)

                If Value = objClsRelation.TiggerData Then
                 'ͼ����ʾ�� ָ������ʾ
                    If Val(objClsRelation.Icon) > 0 Then
                        'ͼ����ʾ��
                        If Len(objClsRelation.IconPerformCol) > 0 Then
                            Set .Cell(flexcpPicture, Row, .ColIndex(objClsRelation.IconPerformCol)) = GetIcon(objClsRelation.Icon)
                        Else
                        'ͼ����ʾ
                            Set .Cell(flexcpPicture, Row, Col) = GetIcon(objClsRelation.Icon)
                        End If
                    End If

                    If Len(objClsRelation.ColorPerformCol) > 0 Then
                        If objClsRelation.CellBackColor > 0 Or objClsRelation.CellFontColor > 0 Then
                            For k = 1 To .Cols - 1
                                If .Cell(flexcpText, 0, k) = objClsRelation.ColorPerformCol Then
                                    '��ɫ��ʾ��
                                    If objClsRelation.RowBackColor > 0 Then .Cell(flexcpBackColor, Row, k) = objClsRelation.CellBackColor
                                    '��ɫ��ʾ��
                                    If objClsRelation.RowFontColor > 0 Then .Cell(flexcpForeColor, Row, k) = objClsRelation.CellFontColor
        
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        '��ǰcell����ɫ
                        If objClsRelation.CellBackColor > 0 Then .Cell(flexcpBackColor, Row, Col) = objClsRelation.CellBackColor
                        '��ǰcellǰ��ɫ
                        If objClsRelation.CellFontColor > 0 Then .Cell(flexcpForeColor, Row, Col) = objClsRelation.CellFontColor
                    End If

                End If
            Next
        End If
    End With

    Exit Sub

errH:
    Err.Raise -1, "frmPacsQuery", "[RowRelationConvert]" & vbCrLf & Err.Description
'    MsgBox Err.Description
End Sub


Private Function GetColSort(ByVal strColName As String) As String
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
'��ȡ������
    GetColSort = strColName
    
    If mTColSort.LngSchemeNo <> mlngSchemeNo Then
        '���Ȼ�ȡһ��������Ϣ
        mTColSort.LngSchemeNo = mlngSchemeNo
        Set mTColSort.dictSortInfo = New Dictionary
        
        If mSqlScheme Is Nothing Then Exit Function
        
        For i = 1 To mSqlScheme.ShowCfgCount
            If Len(mSqlScheme.ShowCfg(i).SortContrastCol) > 0 Then
                Call mTColSort.dictSortInfo.Add(mSqlScheme.ShowCfg(i).Name, mSqlScheme.ShowCfg(i).SortContrastCol)
            End If
        Next
    
    End If
    
    If mTColSort.dictSortInfo.Exists(strColName) Then
        GetColSort = mTColSort.dictSortInfo.Item(strColName)
    End If
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetColSort]" & vbCrLf & Err.Description
End Function

Public Function SetOrder(ByVal lngCurSortCol As Long, ByVal lngCurOrder As Long) As Long
'������������ؼ��Դ������򣨲ο�vsflexgrid������demo��
On Error GoTo errH
    SetOrder = lngCurOrder
    
     'û������ʱ�˳�����
    If vsfList.Rows = 1 Then Exit Function
    
    With vsfList
        Dim R&, c&, RS&, cs&
        .GetSelection R, c, RS, cs
        .Redraw = flexRDNone
    
        ' apply sort to non-empty range
        Dim Row%
        
        For Row = .Rows - 1 To .FixedRows Step -1
            '��������Ϊ��ʱ������������
            If Len(.TextMatrix(Row, lngCurSortCol)) Or Not Trim(.TextMatrix(Row, .ColIndex(mstrListKeyCol))) = "" Then Exit For
        Next
        
        If Row > .FixedRows Then
            .Select .FixedRows, lngCurSortCol, Row, lngCurSortCol
            .Sort = lngCurOrder
        End If
        
        ' restore selection
        .Select R, c, RS, cs
        .Redraw = flexRDDirect
        
        ' cancel default sort
        SetOrder = 0
    End With
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[SetOrder]" & vbCrLf & Err.Description
End Function

Private Sub GetNullStudyInfo()
    With mTStudyInfo
        .strPatientAge = ""
        .strPatientName = ""
        .strPatientSex = ""
        .strStudyNum = ""
        .lngLinkId = 0
        .lngPatId = 0
        .lngAdviceID = 0
    End With
End Sub

Private Function GetSelectRowAdviceID() As Long
'���ݵ�ǰ��ѡ���л�ȡҽ��ID
On Error GoTo errH
    GetSelectRowAdviceID = 0
    If vsfList.Rows < 1 Or vsfList.RowSel < 1 Then Exit Function
    
    GetSelectRowAdviceID = Val(vsfList.TextMatrix(vsfList.RowSel, vsfList.ColIndex(mstrListKeyCol)))
    
    Exit Function
errH:
    Err.Raise -1, "frmPacsQuery", "[GetSelectRowAdviceID]" & vbCrLf & Err.Description
End Function

Private Sub DoPatiIdentify()
'����Pati�ؼ�������Ŀ�͵�ǰѡ����Ŀ
On Error GoTo errH
    mblnAssignment = True
    If mTPatiIdentifyInfo.blFind Then
        patiSearch.IDKindStr = InitCardType(mTPatiIdentifyInfo.strFindItems)
        If mTPatiIdentifyInfo.strFindItem <> "" Then
            patiSearch.IDKindIDX = patiSearch.GetKindIndex(mTPatiIdentifyInfo.strFindItem)
        Else
            mTPatiIdentifyInfo.strFindItem = Split(mTPatiIdentifyInfo.strFindItems, ";")(0)
            patiSearch.IDKindIDX = patiSearch.GetKindIndex(mTPatiIdentifyInfo.strFindItem)
        End If
    Else
        patiSearch.IDKindStr = InitCardType(mTPatiIdentifyInfo.strLocateItems)
        If mTPatiIdentifyInfo.strLocateItems <> "" Then
            patiSearch.IDKindIDX = patiSearch.GetKindIndex(mTPatiIdentifyInfo.strLocateItem)
        Else
            mTPatiIdentifyInfo.strLocateItem = Split(mTPatiIdentifyInfo.strLocateItems, ";")(0)
            patiSearch.IDKindIDX = patiSearch.GetKindIndex(mTPatiIdentifyInfo.strLocateItem)
        End If
    End If
    mblnAssignment = False
    Exit Sub
errH:
    Err.Raise -1, "frmPacsQuery", "[GetSelectRowAdviceID]" & vbCrLf & Err.Description
End Sub

