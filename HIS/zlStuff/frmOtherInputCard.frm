VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmOtherInputCard 
   Caption         =   "����������ⵥ"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmOtherInputCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   2175
      Left            =   3030
      TabIndex        =   30
      Top             =   5730
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   29
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   28
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   13
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9945
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   6
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin VSFlex8Ctl.VSFlexGrid mshBill 
         Height          =   2730
         Left            =   225
         TabIndex        =   4
         Top             =   1035
         Width           =   11085
         _cx             =   19553
         _cy             =   4815
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483628
         GridColor       =   12632256
         GridColorFixed  =   -2147483630
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   32
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOtherInputCard.frx":014A
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
         ExplorerBar     =   7
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
         Begin VB.Image imgLeft 
            Height          =   240
            Left            =   30
            Picture         =   "frmOtherInputCard.frx":058D
            Top             =   45
            Width           =   240
         End
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   27
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6510
         TabIndex        =   23
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9210
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   21
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   20
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   19
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "��������������ⵥ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   17
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   16
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   5925
         TabIndex        =   15
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   8400
         TabIndex        =   14
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&T)"
         Height          =   180
         Left            =   8040
         TabIndex        =   2
         Top             =   660
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0B17
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0D31
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0F4B
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1165
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":137F
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1599
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":17B3
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":19CD
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1BE7
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1E01
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":201B
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":2235
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":244F
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":2669
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":2883
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":2A9D
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOtherInputCard.frx":2CB7
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherInputCard.frx":354B
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherInputCard.frx":3A4D
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin VB.Label lblCode 
      Caption         =   "����"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmOtherInputCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln��ǿ�ƿ���ָ���۸� As Boolean
Private mblnʱ������ֱ��ȷ���ۼ� As Boolean '�⹺���ʱ,ʱ������ֱ��ȷ���ۼ�

Private mbln��������    As Boolean          '����ʱ���ݺ��ۼ�1
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mbln�Ӽ��� As Boolean               'ʱ�������Ƿ��������Ӽ���
Private mdbl�Ӽ��� As Double
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintErrInfor As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼

Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��

Private mrsInOutType As Recordset           '������
Dim mstrPrivs As String                     'Ȩ��
Private mbln�ֶμӳ��� As Boolean   '�Էֶμӳ���Ϊ����
'���˺�:2007/06/10:����10813
Private mstrTime_Start As String            '���뵥�ݱ༭�ĵ���ʱ�� ,��Ҫ�ж��Ƿ񵥾ݱ����˸��Ĺ�,����༭��,���ܽ������
Private mstrTime_End As String
Private Const mlngModule = 1714
Private mbln�ⷿ  As Boolean    '�ÿⷿ�Ƿ�Ϊ���Ŀ�!
Private mblnSort As Boolean     '��������,����������¼�
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴

Private mbln�����������Ų��ؿ��� As Boolean  '�Ƿ�������������Ų����Ƿ�¼��

Private Const mstrCaption As String = "����������ⵥ"

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraMaxFmt As g_FmtString

'----------------------------------------------------------------------------------------------------------
'=========================================================================================
Private mblnFirst As Boolean    '��һ������ʱ
'=========================================================================================



'�������������
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    Dim strSql As String
    
    GetDepend = False
    
    On Error GoTo ErrHandle
    strSql = "" & _
        "   SELECT B.Id,b.���� " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID   AND A.���� = 32 "

    zlDatabase.OpenRecordset rsTemp, strSql, "��������������-��ȡ������"
    If rsTemp.EOF Then
        MsgBox "û����������������������������������������ã�", vbInformation + vbOKOnly, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    Set mrsInOutType = rsTemp
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
    Optional int��¼״̬ As Integer = 1, Optional ByVal strPrivs As String, Optional blnSuccess As Boolean = False)
   '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʾ��༭��Ƭ,��Ψһ���
    '--�����:
    '--������:
    '--��  ��:blnSuccess
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = blnSuccess
    mblnChange = False
    mintErrInfor = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    
    Call GetRegInFor(g˽��ģ��, "��������������", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)
    
    If mint�༭״̬ = 1 Then
        mblnEdit = True
        txtNO.Locked = True
        txtNO.TabStop = True
        txtNO = mstr���ݺ�
        txtNO.Tag = txtNO.Text
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint�༭״̬ = 6 Then
        mblnEdit = False
        CmdSave.Caption = "����(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub
 

Private Sub cboStock_Click()
    Call ��ǰ��Ϊ�ⷿ
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
End Sub
Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, mshBill.ColIndex("����ID")) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���Ӧ���ĵĵ�λ��" & vbCrLf & "��Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '�������ĵ�λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.Rows = 2: mshBill.Cell(flexcpData, 1, mshBill.Cols - 1) = ""
                    mshBill.Clear 1
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
    End With
End Sub
Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    With mshBill
        .SetFocus
        .Row = 1
        .Col = .ColIndex("������Ϣ")
    End With
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("����ID"))) <> 0 Then
                .TextMatrix(intRow, .ColIndex("��������")) = Format(0, mFMT.FM_����)
                .TextMatrix(intRow, .ColIndex("�ɹ����")) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, .ColIndex("�ۼ۽��")) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, .ColIndex("���")) = Format(0, mFMT.FM_���)
                '���˺�:���ۼ۴���
                .TextMatrix(intRow, .ColIndex("���۽��")) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, .ColIndex("���۲��")) = Format(0, mFMT.FM_���)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    With mshBill
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("����ID"))) <> 0 Then
                .TextMatrix(intRow, .ColIndex("��������")) = Format(Val(.TextMatrix(intRow, .ColIndex("����"))), mFMT.FM_����)
                .TextMatrix(intRow, .ColIndex("�ɹ����")) = Format(Val(.TextMatrix(intRow, .ColIndex("����"))) * Val(.TextMatrix(intRow, .ColIndex("�ɹ���"))), mFMT.FM_���)
                .TextMatrix(intRow, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(intRow, .ColIndex("����"))) * Val(.TextMatrix(intRow, .ColIndex("�ۼ�"))), mFMT.FM_���)
                .TextMatrix(intRow, .ColIndex("���")) = Format(Val(.TextMatrix(intRow, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(intRow, .ColIndex("�ɹ����"))), mFMT.FM_���)
                '���˺�:���ۼ۴���,��Ҫ��ȷ��ʱ�۶�������
                Call �������ۼۼ����۲��(intRow, False)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()
    Dim lngRow As Integer
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindVsRowNew mshBill, mshBill.ColIndex("������Ϣ"), txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnChange = False
    Select Case mintErrInfor
        Case 1
            '����
        Case 2
            If mint�༭״̬ = 6 Then
                MsgBox "�õ�����û�п��Գ��������ģ����飡", vbOKOnly, gstrSysName
            Else
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
            End If
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram stbThis, gSystem_Para.int���뷽ʽ
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindVsRowNew mshBill, mshBill.ColIndex("������Ϣ"), txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    
    '�����������ݼ�
    Call SetSortRecord
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        If Not ��鵥��(17, txtNO.Tag) Then Exit Sub
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
        
        '���˺�:2007/06/10:����10813
        mstrTime_End = GetBillInfo(17, txtNO.Tag)
        If mstrTime_End = "" Then
            MsgBox "ע��:" & vbCrLf & "  �õ����Ѿ�����������Աɾ��,���ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("ע��:" & vbCrLf & "  �õ����Ѿ�����������Ա�༭�����ܼ���!" & vbCrLf & "  �Ƿ�����ˢ�µ���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call initCard
            End If
            Exit Sub
        End If
        If SaveCheck = True Then
            strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then '����
        '������Ƿ����
        If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
            MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
            txtժҪ.SetFocus
            Exit Sub
        End If
        
        If SaveStrike Then Unload Me
        Exit Sub
    End If
            
    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
        
        strReg = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
        If Val(strReg) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                printbill
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    txtNO.Text = ""
    mblnSave = False
    mblnEdit = True
    mshBill.Rows = 2: mshBill.Clear 1
    
    Call RefreshRowNO(mshBill, mshBill.ColIndex("�к�"), 1)
    SetEdit
    
    txtժҪ.Text = ""
    cboType.SetFocus
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
End Sub

Private Sub Form_Load()
    Dim strReg As String

    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    mintUnit = Val(strReg)
 

    mblnFirst = True
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(2, g_�ۼ�)
    End With
    

    mintBatchNoLen = GetBatchNoLen()
    
    mbln�Ӽ��� = Get�Ӽ���
    mbln�ֶμӳ��� = IS�ֶμӳ���()
    mbln��ǿ�ƿ���ָ���۸� = ISCHECK��ǿ�ƿ���ָ���۸�()
    mblnʱ������ֱ��ȷ���ۼ� = isʱ������ֱ��ȷ���ۼ�()
    
    mbln�����������Ų��ؿ��� = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO.Text
    With cboType
        .Clear
        Do While Not mrsInOutType.EOF
            .AddItem mrsInOutType.Fields(1)
            .ItemData(.NewIndex) = mrsInOutType.Fields(0)
            mrsInOutType.MoveNext
        Loop
        .ListIndex = 0
    End With
      
    Call initCard
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(.ColIndex("�ɹ���")) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(.ColIndex("�ɹ����")) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(.ColIndex("���")) = IIf(mblnCostView = True, 900, 0)
    End With
        
    mshBill_LostFocus
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    
    On Error GoTo ErrHandle

    '�ⷿ
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    
    strCompare = Mid(strOrder, 1, 1)
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    '��ʼ������ؼ�
    Call initGrid
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        Case 2, 3, 4, 6
            If mint�༭״̬ = 4 Then
                gstrSQL = "" & _
                    "   Select b.id,b.����  " & _
                    "   From ҩƷ�շ���¼ a,���ű� b " & _
                    "   Where a.�ⷿid=b.id and A.���� = 17 and a.no=[1]"
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
                
                If rsTemp.EOF Then: mintErrInfor = 2: Exit Sub
                
                With cboStock
                    .AddItem rsTemp!����: .ItemData(.NewIndex) = rsTemp!Id: .ListIndex = 0
                End With
                rsTemp.Close
            End If
            
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "c.���㵥λ AS ��λ ,c.���㵥λ AS ���۵�λ,(A.��д���� ) AS ����,b.ָ�������� as ָ�������� , a.�ɱ��� as �ɱ���  ,  1 as ����ϵ��,"
                Case Else
                    strUnitQuantity = "B.��װ��λ AS ��λ,c.���㵥λ AS ���۵�λ,(A.��д���� / B.����ϵ��) AS ����,b.ָ��������*B.����ϵ�� as ָ�������� , a.�ɱ���*B.����ϵ�� as �ɱ��� ,B.����ϵ�� as ����ϵ��,"
            End Select
            
            If mint�༭״̬ <> 6 Then
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.ҩƷid as ����id,A.���,('[' || c.���� || ']' || c.����) AS ������Ϣ, " & _
                    "               zlSpellCode(c.����) ����,c.���,c.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,to_char(a.��������,'yyyy-mm-dd') ��������," & _
                    "               b.���Ч��,A.Ч��,a.�������,a.���Ч�� as ���ʧЧ��,a.��Ʒ����,b.һ���Բ���,nvl(b.�Ƿ��������,0) as �������,b.�ⷿ����,b.���Ч��," & strUnitQuantity & _
                    "               A.�ɱ����,A.���ۼ�,to_number(nvl(to_char(a.�÷�," & gOraFmt_Max.FM_��� & " ),0), " & gOraFmt_Max.FM_��� & ") as ���۲��, " & _
                    "               A.���۽��,A.���,b.ָ�������/100 as ָ�������,nvl(b.�ӳ���,0)/100 as �ӳ���,c.�Ƿ���,b.���÷���, " & _
                    "               a.ժҪ,������,��������,�����,�������,a.�ⷿid,g.���� as ����,a.������id " & _
                    "           FROM ҩƷ�շ���¼ A,�������� b,�շ���ĿĿ¼ c,���ű� g " & _
                    "           Where A.ҩƷid = B.����id and a.ҩƷid=c.id and a.�ⷿid=g.id " & _
                    "                   AND A.��¼״̬ =[2]" & _
                    "                   AND A.���� = 17 AND A.No =[1] )" & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Else
                gstrSQL = "" & _
                    "   Select * " & _
                    "   From (  SELECT distinct a.ҩƷid as ����id,A.���,('[' || c.���� || ']' || c.����) AS ������Ϣ, " & _
                    "                   zlSpellCode(c.����) ����,c.���,c.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,to_char(a.��������,'yyyy-mm-dd') ��������," & _
                    "                   b.���Ч��,A.Ч��,a.�������,a.���Ч�� as ���ʧЧ��,b.һ���Բ���,nvl(b.�Ƿ��������,0) as �������,b.�ⷿ����,b.���Ч��," & strUnitQuantity & _
                    "                   A.�ɱ����,A.���ۼ�,A.���۲��," & _
                    "                   A.���۽��,A.���,b.ָ�������/100 as ָ�������,nvl(b.�ӳ���,0)/100 as �ӳ���,c.�Ƿ���,b.���÷���,A.��д���� as ��ʵ����, " & _
                    "                   a.�ⷿid,g.���� as ����,a.������id,a.��Ʒ���� " & _
                    "           FROM (  Select min(id) as id, sum(��д����) as ��д����,sum(�ɱ����) as �ɱ����, " & _
                    "                       ҩƷid,���,����,��׼�ĺ�, ����,��������,Ч��,�������,���Ч��,����,�ɱ���," & _
                    "                       ���ۼ�,sum(���۽��) as ���۽��,Sum(���) as ���,Sum(to_number(nvl(to_char(x.�÷�," & gOraFmt_Max.FM_��� & " ),0), " & gOraFmt_Max.FM_��� & ")) as ���۲��," & _
                    "                       �ⷿID,������ID,��Ʒ����" & _
                    "                   From ҩƷ�շ���¼ x " & _
                    "                   WHERE NO=[1] AND ����=17  " & _
                    "                   group by ҩƷID,���,����,��׼�ĺ�, ����,��������,Ч��,�������,���Ч��,����,�ɱ���,���ۼ�,�ⷿID,������ID,��Ʒ����" & _
                    "                   having sum(��д����)<>0 " & _
                    "                 ) A,�������� b,�շ���ĿĿ¼ c,���ű� g " & _
                    "           Where A.ҩƷid = B.����id and a.ҩƷid=c.id and a.�ⷿid=g.id ) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, mint��¼״̬)
            
            If rsTemp.EOF Then: mintErrInfor = 2: Exit Sub
            
            '���˺�:2007/06/10:����10813
            mstrTime_Start = GetBillInfo(17, mstr���ݺ�)
            
            Select Case mint�༭״̬
                Case 2, 6
                    Txt������ = UserInfo.�û���
                    Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    If mint�༭״̬ = 2 Then
                        Txt����� = ""
                        Txt������� = ""
                    Else
                        Txt����� = UserInfo.�û���
                        Txt������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    End If
                Case Else
                    Txt������ = rsTemp!������
                    Txt�������� = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
                    Txt����� = IIf(IsNull(rsTemp!�����), "", rsTemp!�����)
                    Txt������� = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd hh:mm:ss"))
            End Select
            
            If mint�༭״̬ <> 6 Then
                txtժҪ.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
            Else
                txtժҪ.Text = GetժҪ(mstr���ݺ�)
            End If
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintErrInfor = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsTemp!������ID Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
            End With
            intRow = 0
            With mshBill
                .Clear 1
                .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
                Do While Not rsTemp.EOF
                    intRow = intRow + 1
                    .TextMatrix(intRow, .ColIndex("����ID")) = zlStr.NVL(rsTemp!����ID)
                    .TextMatrix(intRow, .ColIndex("������Ϣ")) = zlStr.NVL(rsTemp!������Ϣ)
                    .TextMatrix(intRow, .ColIndex("���")) = zlStr.NVL(rsTemp!���)
                    .TextMatrix(intRow, .ColIndex("���")) = zlStr.NVL(rsTemp!���)
                    .TextMatrix(intRow, .ColIndex("����")) = zlStr.NVL(rsTemp!����)
                    .TextMatrix(intRow, .ColIndex("��׼�ĺ�")) = zlStr.NVL(rsTemp!��׼�ĺ�)
                    .TextMatrix(intRow, .ColIndex("��λ")) = zlStr.NVL(rsTemp!��λ)
                    .TextMatrix(intRow, .ColIndex("����")) = zlStr.NVL(rsTemp!����)
                    .TextMatrix(intRow, .ColIndex("��������")) = zlStr.NVL(rsTemp!��������)
                    .TextMatrix(intRow, .ColIndex("Ч��")) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
                    .TextMatrix(intRow, .ColIndex("һ���Բ���")) = Val(zlStr.NVL(rsTemp!һ���Բ���))
                    .TextMatrix(intRow, .ColIndex("�������")) = Val(zlStr.NVL(rsTemp!�������))
                    .TextMatrix(intRow, .ColIndex("���Ч��")) = zlStr.NVL(rsTemp!���Ч��)
                    .TextMatrix(intRow, .ColIndex("�������")) = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd"))
                    .TextMatrix(intRow, .ColIndex("���ʧЧ��")) = IIf(IsNull(rsTemp!���ʧЧ��), "", Format(rsTemp!���ʧЧ��, "yyyy-mm-dd"))
                    .TextMatrix(intRow, .ColIndex("����")) = Format(rsTemp!����, mFMT.FM_����)
                    .TextMatrix(intRow, .ColIndex("��Ʒ����")) = zlStr.NVL(rsTemp!��Ʒ����)
                    If rsTemp!���� <> 0 Then
                        .TextMatrix(intRow, .ColIndex("�ɹ���")) = Format(rsTemp!�ɱ���� / rsTemp!����, mFMT.FM_�ɱ���)
                    Else
                        .TextMatrix(intRow, .ColIndex("�ɹ���")) = "0.00"
                    End If
   
                    '���˺�:���ۼ۴���:���ۼ�-->���ۼ�;���۽��-->���۽��;���-->���۲��;��;-->�ⷿ��λ���
                    ' ���۽�������������ۼۣ�
                    ' ���۲��"���ۼ۽����۽�������ⵥλ����Ľ��Ͱ����۵�λ����Ľ��Ĳ�ֵ��
                    .TextMatrix(intRow, .ColIndex("���ۼ�")) = Format(Val(zlStr.NVL(rsTemp!���ۼ�)), mFMT.FM_ɢװ���ۼ�)          'If Val(.TextMatrix(.row, .colindex("���ۼ�"))) = 0 Then
                    
                    '�����ۼ�
                    .TextMatrix(intRow, .ColIndex("�ۼ�")) = Format((Val(zlStr.NVL(rsTemp!���۽��)) - Val(zlStr.NVL(rsTemp!���۲��))) / Val(zlStr.NVL(rsTemp!����)), mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, .ColIndex("���۵�λ")) = zlStr.NVL(rsTemp!���۵�λ)
                    
                    If mint�༭״̬ = 6 Then
                        '����û����صĲ��
                        .TextMatrix(intRow, .ColIndex("���۲��")) = ""
                        .TextMatrix(intRow, .ColIndex("���۽��")) = ""
                        .TextMatrix(intRow, .ColIndex("���")) = ""
                        .TextMatrix(intRow, .ColIndex("�ۼ۽��")) = ""
                        .TextMatrix(intRow, .ColIndex("�ɹ����")) = ""
                    Else
                        .TextMatrix(intRow, .ColIndex("���۲��")) = Format(Val(zlStr.NVL(rsTemp!���)), mFMT.FM_���)
                        .TextMatrix(intRow, .ColIndex("���۽��")) = Format(Val(zlStr.NVL(rsTemp!���۽��)), mFMT.FM_���)
                        '�����ۼۼ��ۼ۽��
                        .TextMatrix(intRow, .ColIndex("���")) = Format(Val(zlStr.NVL(rsTemp!���)) - Val(zlStr.NVL(rsTemp!���۲��)), mFMT.FM_���)
                        .TextMatrix(intRow, .ColIndex("�ۼ۽��")) = Format(Val(zlStr.NVL(rsTemp!���۽��)) - Val(zlStr.NVL(rsTemp!���۲��)), mFMT.FM_���)
                        .TextMatrix(intRow, .ColIndex("�ɹ����")) = Format(Val(zlStr.NVL(rsTemp!�ɱ����)), mFMT.FM_���)
                    End If
                    .TextMatrix(intRow, .ColIndex("ԭ����")) = IIf(IsNull(rsTemp!ԭ����), "!", rsTemp!ԭ����)
                    
                    '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                    .TextMatrix(intRow, .ColIndex("ԭ����")) = IIf(IsNull(rsTemp!���Ч��), "0", rsTemp!���Ч��) & "||" & rsTemp!�ӳ��� & "||" & IIf(IsNull(rsTemp!�Ƿ���), 0, rsTemp!�Ƿ���) & "||" & IIf(IsNull(rsTemp!���÷���), 0, rsTemp!���÷���) & "||" & zlStr.NVL(rsTemp!�ⷿ����, 0)
                    .TextMatrix(intRow, .ColIndex("����ϵ��")) = rsTemp!����ϵ��
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, .ColIndex("��������")) = Format(0, mFMT.FM_����)
                        .TextMatrix(intRow, .ColIndex("��ʵ����")) = zlStr.NVL(rsTemp!��ʵ����)
                    End If
                    rsTemp.MoveNext
                Loop
            End With
            rsTemp.Close
    End Select
    SetEdit         '���ñ༭����
    Call RefreshRowNO(mshBill, mshBill.ColIndex("�к�"), 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetժҪ(ByVal strNo As String) As String
    '��ȡ�µ�ժҪ
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
         '����(ȡ���һ�γ�����ժҪ)
    gstrSQL = "Select ժҪ From ҩƷ�շ���¼ Where ����=17 And No=[1] and (��¼״̬ =1 or mod(��¼״̬,3)=0) Order By ������� Desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡժҪ��Ϣ", strNo)
    
    If Not rsTemp.EOF Then
        GetժҪ = zlStr.NVL(rsTemp!ժҪ)
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub imgLeft_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(mshBill.hwnd)
    lngLeft = vRect.Left + imgLeft.Left
    lngTop = vRect.Top + imgLeft.Height
    Call frmVsColSel.ShowColSet(Me, mstrCaption, mshBill, lngLeft, lngTop, imgLeft.Height)
    zl_vsGrid_Para_Save mlngModule, mshBill, mstrCaption, "��ͷ��Ϣ", True
End Sub

Private Sub mshBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        If mblnSort = True Then Exit Sub
        Call zl_VsGridRowChange(mshBill, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub mshBill_AfterSort(ByVal Col As Long, Order As Integer)
    With mshBill
    End With
End Sub

Private Sub mshBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long, arrSplit As Variant
    With mshBill
        If mblnEdit = False Then
            If mint�༭״̬ = 6 Then
                If .ColIndex("��������") = Col Then: Exit Sub
                Cancel = True: Exit Sub
            End If
            Cancel = True
        End If
        Select Case Col
        Case .ColIndex("������Ϣ"), .ColIndex("����")
        Case .ColIndex("�������")
            '��һ���Բ��ϲ��ܱ༭�������
            If Val(.TextMatrix(Row, .ColIndex("һ���Բ���"))) <> "1" Then Cancel = True: Exit Sub
        Case .ColIndex("��Ʒ����")
            If Val(.TextMatrix(Row, .ColIndex("�������"))) <> 1 Then Cancel = True: Exit Sub
        Case .ColIndex("��������"), .ColIndex("��׼�ĺ�")
        
        Case .ColIndex("����"), .ColIndex("�ɹ���"), .ColIndex("�ɹ����")
        Case .ColIndex("����")
            If Val(.TextMatrix(Row, .ColIndex("����ID"))) <= 0 Then '���н�ֹ����
                Cancel = True
            End If
        Case .ColIndex("Ч��")
             '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
             If .TextMatrix(Row, .ColIndex("ԭ����")) <> "" Then
                arrSplit = Split(.TextMatrix(Row, .ColIndex("ԭ����")), "||")
                If Val(arrSplit(4)) = 0 Then Cancel = True: Exit Sub    '�ǿⷿ����,���ܱ༭Ч��
             Else
                Cancel = True
             End If
        Case .ColIndex("�ۼ�")
            '�����ʱ�����ģ������������ۼ�,�Ҳ���"ʱ������ֱ��ȷ���ۼ���Ч
             If .TextMatrix(Row, .ColIndex("ԭ����")) <> "" Then
                arrSplit = Split(.TextMatrix(Row, .ColIndex("ԭ����")), "||")
                If Not (Val(arrSplit(2)) = 1 And mblnʱ������ֱ��ȷ���ۼ�) Then Cancel = True: Exit Sub     '�ǿⷿ����,���ܱ༭Ч��
             Else
                Cancel = True
             End If
        Case .ColIndex("���ۼ�")
             If .TextMatrix(Row, .ColIndex("ԭ����")) <> "" Then
                arrSplit = Split(.TextMatrix(Row, .ColIndex("ԭ����")), "||")
                If Val(arrSplit(2)) = 1 And (IIf(mbln�ⷿ, Val(arrSplit(4)) = 1, Val(arrSplit(3)) = 1)) Then
                   'ʵ�������ҿⷿ������
                    Exit Sub
                End If
             End If
            Cancel = True: Exit Sub
        Case Else: Cancel = True
        End Select
    End With
End Sub
Private Sub SetInputFormat(ByVal intRow As Integer)
    Dim arrSplit As Variant
    If mblnEdit = False Then Exit Sub
    
    With mshBill
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("����")) = "0||" & IIf(Val(.TextMatrix(intRow, .ColIndex("����ID"))) > 0, 0, 2) '& IIf(.TextMatrix(intRow, .ColIndex("ԭ����")) = "!", 0, 2)
        .ColData(.ColIndex("�������")) = "0||" & IIf(Val(.TextMatrix(intRow, .ColIndex("һ���Բ���"))) = 1, 0, 2)
        .ColData(.ColIndex("Ч��")) = "0||2"
        .ColData(.ColIndex("�ۼ�")) = "0||2"
        .ColData(.ColIndex("���ۼ�")) = "0||2"
        .ColData(.ColIndex("��Ʒ����")) = "0||" & IIf(.TextMatrix(intRow, .ColIndex("�������")) = "1", 0, 2)
        If .TextMatrix(intRow, .ColIndex("ԭ����")) <> "" Then
            '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
            arrSplit = Split(.TextMatrix(intRow, .ColIndex("ԭ����")), "||")
            .ColData(.ColIndex("Ч��")) = "0||" & IIf(Val(arrSplit(4)) = 1, 0, 2)
            '�����ʱ�����ģ������������ۼ�
            .ColData(.ColIndex("�ۼ�")) = "0||" & IIf(Val(arrSplit(2)) = 1 And mblnʱ������ֱ��ȷ���ۼ�, 0, 2)
            If Val(arrSplit(2)) = 1 And arrSplit(4) = 1 Then
                'ʵ���ҷ�����
                .ColData(.ColIndex("���ۼ�")) = "0||0"
            End If
        End If
    End With
End Sub
Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            cboStock.Enabled = False
            cboType.Enabled = False
            txtժҪ.Enabled = True
            If mint�༭״̬ = 6 Then .Editable = flexEDKbdMouse
            
            If mint�༭״̬ <> 6 Then
                txtժҪ.Enabled = False
            End If
        Else
            cboStock.Enabled = True

            cboType.Enabled = True
            txtժҪ.Enabled = True
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub initGrid()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ���Ĭ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-12-02 11:39:14
    '-----------------------------------------------------------------------------------------------------------

    With mshBill
        
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, mshBill, mstrCaption, "��ͷ��Ϣ", True, True
        
        .FixedCols = 1
        If mintUnit = 0 Then
            .ColHidden(.ColIndex("���ۼ�")) = True
            .ColHidden(.ColIndex("���۵�λ")) = True
            .ColHidden(.ColIndex("���۽��")) = True
            .ColHidden(.ColIndex("���۲��")) = True
            .ColData(.ColIndex("���ۼ�")) = -1
            .ColData(.ColIndex("���۵�λ")) = -1
            .ColData(.ColIndex("���۽��")) = -1
            .ColData(.ColIndex("���۲��")) = -1
        End If
        '���س�����
        .ColHidden(.ColIndex("��������")) = IIf(mint�༭״̬ = 6, False, True)
        If .ColWidth(.ColIndex("��������")) = 0 And .ColHidden(.ColIndex("��������")) = False Then .ColWidth(.ColIndex("��������")) = 800
        .ColHidden(.ColIndex("����ID")) = True
        .ColHidden(.ColIndex("���")) = True
        .ColHidden(.ColIndex("��ʵ����")) = True
        .ColHidden(.ColIndex("һ���Բ���")) = True
        .ColHidden(.ColIndex("�������")) = True
        .ColHidden(.ColIndex("���Ч��")) = True
        .ColHidden(.ColIndex("ԭ����")) = True
        .ColHidden(.ColIndex("ԭ����")) = True
        .ColHidden(.ColIndex("����ϵ��")) = True
        If mblnCostView = False Then
            .ColHidden(.ColIndex("�ɹ���")) = True
            .ColHidden(.ColIndex("�ɹ����")) = True
            .ColHidden(.ColIndex("���")) = True
        Else
            .ColHidden(.ColIndex("�ɹ���")) = False
            .ColHidden(.ColIndex("�ɹ����")) = False
            .ColHidden(.ColIndex("���")) = False
        End If
        .ColData(.ColIndex("��������")) = "-1|0"
        .ColData(.ColIndex("������Ϣ")) = "1|0"
        .ColData(.ColIndex("����")) = "1|0"
        If mblnCostView = False Then
            .ColData(.ColIndex("�ɹ���")) = "-1|1"
            .ColData(.ColIndex("�ɹ����")) = "-1|1"
            .ColData(.ColIndex("���")) = "-1|1"
        Else
            .ColData(.ColIndex("�ɹ���")) = "1|0"
            .ColData(.ColIndex("���")) = "0||2"
        End If
        .ColData(.ColIndex("�ۼ�")) = "1|0"
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("����ID")) = -1
        .ColData(.ColIndex("���")) = -1
        .ColData(.ColIndex("��ʵ����")) = -1
        .ColData(.ColIndex("һ���Բ���")) = -1
        .ColData(.ColIndex("�������")) = -1
        .ColData(.ColIndex("���Ч��")) = -1
        .ColData(.ColIndex("ԭ����")) = -1
        .ColData(.ColIndex("ԭ����")) = -1
        .ColData(.ColIndex("����ϵ��")) = -1
        
        .ColData(.ColIndex("���")) = "0||2"
        .ColData(.ColIndex("��λ")) = "0||2"
        .ColData(.ColIndex("���ʧЧ��")) = "0||2"
        .ColData(.ColIndex("�ۼ۽��")) = "0||2"
        .ColData(.ColIndex("���۵�λ")) = "0||2"
        .ColData(.ColIndex("���۽��")) = "0||2"
        .ColData(.ColIndex("���۲��")) = "0||2"
        
        If gblnCode Then
            .ColData(.ColIndex("��Ʒ����")) = "0||2"
        Else
            .ColData(.ColIndex("��Ʒ����")) = "-1||1"
            .ColHidden(.ColIndex("��Ʒ����")) = True
        End If

    End With
    txtժҪ.MaxLength = sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With
    
    
    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cboType.Left = mshBill.Left + mshBill.Width - cboType.Width
    
    LblType.Left = cboType.Left - LblType.Width - 100
    
    
    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With
    
    With Lbl��������
        .Top = Lbl������.Top
        .Left = Txt������.Left + Txt������.Width + 250
    End With
    
    With Txt��������
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    With Txt�������
        .Top = Lbl������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl�������
        .Top = Lbl������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With
    
    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = Lbl�������.Left - 200 - .Width
    End With
    
    With Lbl�����
        .Top = Lbl������.Top
        .Left = Txt�����.Left - 100 - .Width
    End With
    
    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With
        
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnCostView = False Then
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        zl_vsGrid_Para_Save mlngModule, mshBill, mstrCaption, "��ͷ��Ϣ", True, True
        
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    zl_vsGrid_Para_Save mlngModule, mshBill, mstrCaption, "��ͷ��Ϣ", True, True
End Sub
Private Function SaveCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���������ⵥ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-12-02 11:40:30
    '-----------------------------------------------------------------------------------------------------------
    mblnSave = False: SaveCheck = False
    
    gstrSQL = "zl_�����������_Verify('" & txtNO.Tag & "','" & UserInfo.�û��� & "')"
    
    On Error GoTo ErrHandle
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
    SaveCheck = True: mblnSave = True: mblnSuccess = True: mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub mshBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '-----------------------------------------------------------------------------------------------------------
    '����:������صĸ�ʽ
    '����:
    '����:���˺�
    '����:2008-12-02 11:43:38
    '-----------------------------------------------------------------------------------------------------------
    Dim arrSplit As Variant, str���� As String, strxq As String
    With mshBill
        Select Case Col
        Case .ColIndex("������Ϣ")
            .ColComboList(Col) = "..."
        Case .ColIndex("����")
        Case .ColIndex("�ɹ���"), .ColIndex("�ɹ����")
            Call �������ۼۼ����۲��(Row, True)
              ��ʾ�ϼƽ��
        Case .ColIndex("�ۼ�")
            Call �������ۼۼ����۲��(Row, True)
              ��ʾ�ϼƽ��
        Case .ColIndex("���ۼ�")
            Call �������ۼۼ����۲��(Row, False)
              ��ʾ�ϼƽ��
        Case .ColIndex("����"), .ColIndex("��������")
            Call �������ۼۼ����۲��(Row, True)
              ��ʾ�ϼƽ��
        Case .ColIndex("��Ʒ����")
            .TextMatrix(Row, Col) = UCase(.TextMatrix(Row, Col))
        Case .ColIndex("����")
                If Trim(.TextMatrix(Row, .ColIndex("����"))) = "" Or IsNumeric(.TextMatrix(Row, .ColIndex("����"))) = False Then
                    If Not IsDate(Trim(.TextMatrix(Row, .ColIndex("��������")))) Then
                        str���� = ""
                    Else
                        str���� = Format(.TextMatrix(Row, .ColIndex("��������")), "yyyymmdd")
                    End If
                Else
                    str���� = Trim(.TextMatrix(Row, .ColIndex("����")))
                End If
                If str���� <> "" Then
                    '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                    arrSplit = Split(.TextMatrix(Row, .ColIndex("ԭ����")) & "||||||||||", "||")
                    If IsNumeric(str����) And Val(arrSplit(0)) <> 0 Then
                        strxq = UCase(str����)
                        If Trim(.TextMatrix(Row, .ColIndex("��������"))) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq, True)
                                If strxq = "" Then Exit Sub
                                .TextMatrix(Row, .ColIndex("��������")) = Format(strxq, "yyyy-mm-dd")
                                'Call CheckLapse(.TextMatrix(row, .ColIndex("Ч��")))
                            End If
                        End If
                    End If
                End If
        Case .ColIndex("��������")
'                If Trim(.TextMatrix(row, .ColIndex("����"))) = "" Or IsNumeric(.TextMatrix(row, .ColIndex("����"))) = False Then
                    If Not IsDate(Trim(.TextMatrix(Row, .ColIndex("��������")))) Then
                        str���� = ""
                    Else
                        str���� = Format(.TextMatrix(Row, .ColIndex("��������")), "yyyymmdd")
                    End If
'                Else
'                    str���� = Trim(.TextMatrix(row, .ColIndex("����")))
'                End If
                
                If str���� <> "" Then
                    '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                    arrSplit = Split(.TextMatrix(Row, .ColIndex("ԭ����")) & "||||||||||", "||")
                    If IsNumeric(str����) And Val(arrSplit(0)) <> 0 Then
                        strxq = UCase(str����)
                        If Trim(.TextMatrix(Row, .ColIndex("Ч��"))) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq, True)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(Row, .ColIndex("Ч��")) = Format(DateAdd("M", Val(arrSplit(0)), strxq), "yyyy-mm-dd")
                                Call CheckLapse(.TextMatrix(Row, .ColIndex("Ч��")))
                            End If
                        End If
                    End If
                End If
        End Select
    End With
End Sub
Private Sub AfterAddRow(Row As Long)
    '�����к�
    Call RefreshRowNO(mshBill, mshBill.ColIndex("�к�"), Row)
End Sub

Private Sub BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If Val(.TextMatrix(Row, .ColIndex("����ID"))) <> 0 Then
            If MsgBox("���Ƿ����Ҫɾ����������Ϊ��" & .TextMatrix(.Row, .ColIndex("������Ϣ")) & "���ļ�¼��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    End With
End Sub
Private Sub mshBill_BeforeSort(ByVal Col As Long, Order As Integer)
    mblnSort = True
    Call zl_VsGridBeforeSort(mshBill, Col, Order, mshBill.ColIndex("�к�"))
    With mshBill
        .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColorBkg
        Call zl_VsGridRowChange(mshBill, .FixedRows, .Row, .FixedCols, .Col)
        If InStr(1, "12", mint�༭״̬) > 0 Then Call RefreshRowNO(mshBill, .ColIndex("�к�"), 1)
    End With
    mblnSort = False
End Sub

Private Sub mshBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '����:��ťѡ��
    '����:
    '--------------------------------------------------------------------------
    Dim lngRow As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    With mshBill
        Select Case Col
        Case .ColIndex("������Ϣ")
            If Select������Ϣ("") = False Then Exit Sub
            Call zlVsMoveGridCell(mshBill, .ColIndex("������Ϣ"), , IIf(mint�༭״̬ = 1 Or mint�༭״̬ = 2, True, False), lngRow)
            
        Case .ColIndex("����")
            If SelectAndNotAddItem(Me, mshBill, "", "����������", "����������ѡ����", True, True, , zl_��ȡվ������(True)) = True Then
                Call zlVsMoveGridCell(mshBill, .ColIndex("������Ϣ"), , IIf(mint�༭״̬ = 1 Or mint�༭״̬ = 2, True, False), lngRow)
            End If
            
            If .TextMatrix(.Row, .ColIndex("����")) <> "" Then
                gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, .ColIndex("����")), .TextMatrix(.Row, .ColIndex("����ID")))
                If rsTemp.RecordCount Then
                    .TextMatrix(.Row, .ColIndex("��׼�ĺ�")) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
                Else
                    .TextMatrix(.Row, .ColIndex("��׼�ĺ�")) = ""
                End If
            End If
        Case .ColIndex("Ч��")
            If SelDate(Col) = False Then Exit Sub
        Case .ColIndex("��������")
            If SelDate(Col) = False Then Exit Sub
        Case .ColIndex("�������")
            If SelDate(Col) = False Then Exit Sub
        End Select
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_ChangeEdit()
    mblnChange = True
End Sub
 
Private Sub mshBill_GotFocus()
    Call zl_VsGridGotFocus(mshBill)
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    
    With mshBill
        If KeyCode <> vbKeyReturn And KeyCode <> vbKeyReturn _
            And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                mshBill_CellButtonClick .Row, .Col
            Else
            
            Select Case .Col
            Case .ColIndex("������Ϣ"), .ColIndex("����"), .ColIndex("Ч��"), .ColIndex("��������"), .ColIndex("�������")
                .ColComboList(.Col) = ""
            Case Else
            End Select
            End If
        End If
 
        If KeyCode = vbKeyDelete Then
            blnCancel = False
            'ɾ����ǰ
            Call BeforeDeleteRow(.Row, blnCancel)
            If blnCancel = True Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            'ɾ���к�
            Call AfterDeleteRow
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        If Val(.TextMatrix(.Row, .ColIndex("����ID"))) = 0 And .Col = .ColIndex("������Ϣ") Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(mshBill, .ColIndex("������Ϣ"), , IIf(mint�༭״̬ = 1 Or mint�༭״̬ = 2, True, False), lngRow)
        If lngRow >= 0 Then
            Call AfterAddRow(lngRow)
        End If
    End With
End Sub

Private Sub mshBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long
    Dim rsProvider As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshBill
        Select Case Col
        Case .ColIndex("������Ϣ")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If Select������Ϣ(strKey) = False Then
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
            .EditText = .TextMatrix(Row, Col)
        Case .ColIndex("����")
            strKey = Trim(.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If SelectAndNotAddItem(Me, mshBill, strKey, "����������", "����������ѡ����", True, True, , zl_��ȡվ������(True)) = True Then
                gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, .ColIndex("����")), .TextMatrix(.Row, .ColIndex("����ID")))
                If rsProvider.RecordCount > 0 Then
                    .TextMatrix(.Row, .ColIndex("��׼�ĺ�")) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
                Else
                    .TextMatrix(.Row, .ColIndex("��׼�ĺ�")) = ""
                End If
            Else
                .EditText = ""
                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
        Case Else
        
        End Select
        Call zlVsMoveGridCell(mshBill, .ColIndex("������Ϣ"), -1, True, lngRow)
        If lngRow >= 0 Then AfterAddRow lngRow
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub mshBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        Select Case .Col
            Case .ColIndex("������Ϣ")
                VsFlxGridCheckKeyPress mshBill, Row, Col, KeyAscii, m�ı�ʽ
            Case .ColIndex("����"), .ColIndex("����")
                VsFlxGridCheckKeyPress mshBill, Row, Col, KeyAscii, m�ı�ʽ
            Case .ColIndex("Ч��"), .ColIndex("��������"), .ColIndex("�������")
                VsFlxGridCheckKeyPress mshBill, Row, Col, KeyAscii, m�ı�ʽ
            Case .ColIndex("�ɹ���"), .ColIndex("�ɹ����"), .ColIndex("�ۼ�"), _
                 .ColIndex("����"), .ColIndex("���ۼ�"), .ColIndex("��������")
                VsFlxGridCheckKeyPress mshBill, Row, Col, KeyAscii, m���ʽ
                strKey = .EditText
                If strKey = "" Then
                    strKey = .TextMatrix(.Row, .Col)
                End If
                Select Case .Col
                    Case .ColIndex("����"), .ColIndex("��������")
                        intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.����С��, g_С��λ��.obj_ɢװС��.����С��)
                    Case .ColIndex("�ɹ���")
                       intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.�ɱ���С��, g_С��λ��.obj_ɢװС��.�ɱ���С��)
                    Case .ColIndex("�ɹ����")
                        intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.���С��, g_С��λ��.obj_ɢװС��.���С��)
                    Case .ColIndex("���ۼ�"), .ColIndex("�ۼ�")
                        intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.���ۼ�С��, g_С��λ��.obj_ɢװС��.���ۼ�С��)
                End Select
                If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                    Exit Sub
                End If
                If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                    If .EditSelLength = Len(strKey) Then Exit Sub
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Case .ColIndex("��Ʒ����")
                Select Case KeyAscii
                    Case vbKeyBack, vbKeyEscape, 3, 22
                        Exit Sub
                    Case vbKeyReturn
'                        Call OS.PressKey(vbKeyTab)
                        Exit Sub
                    Case Else
                        '����¼�����ֺ���ĸ
                        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Then Exit Sub
                End Select
                KeyAscii = 0
        End Select
    End With
End Sub
Private Sub mshBill_LeaveCell()
    If mblnSort Then Exit Sub
    OS.OpenIme False
End Sub
Private Sub mshBill_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        '���õ�Ԫ��ı༭����
        With mshBill
           Select Case .Col
               Case .ColIndex("������Ϣ")
                   .EditMaxLength = 40
               Case .ColIndex("����")
                   .EditMaxLength = 30
               Case .ColIndex("����")
                   .EditMaxLength = mintBatchNoLen
              Case .ColIndex("Ч��")
                   .EditMaxLength = 10
               Case .ColIndex("��������")
                   .EditMaxLength = 10
               Case .ColIndex("�������")
                    .EditMaxLength = 10
               Case .ColIndex("�ɹ���"), .ColIndex("�ɹ����"), .ColIndex("�ۼ�"), .ColIndex("���ۼ�"), .ColIndex("����"), .ColIndex("��������")
                   .EditMaxLength = 16
           End Select
    End With
End Sub

Private Sub mshbill_EnterCell()
    If mblnSort = True Then Exit Sub
    '�������޸ĲŴ�������
    If mint�༭״̬ <> 1 And mint�༭״̬ <> 2 Then Exit Sub
    With mshBill
        SetInputFormat .Row
        OS.OpenIme (False)
        Select Case .Col
        Case .ColIndex("������Ϣ")
             .ColComboList(.Col) = "..."
            'ֻ������id�в���ʾ�ϼ���Ϣ�Ϳ����
            Call ��ʾ�ϼƽ��
            Call ��ʾ�����
        Case .ColIndex("Ч��"), .ColIndex("�������"), .ColIndex("��������")
            .ColComboList(.Col) = "..."
            If .ColIndex("Ч��") = .Col Then
                If Trim(.TextMatrix(.Row, .Col)) <> "" Then Exit Sub
                Dim str�������� As String, strxq As String
                If Not IsDate(.TextMatrix(.Row, .ColIndex("��������"))) Then
                    str�������� = ""
                Else
                    str�������� = Format(.TextMatrix(.Row, .ColIndex("��������")), "yyyymmdd")
                End If
                
                If str�������� <> "" And Trim(.TextMatrix(.Row, .ColIndex("ԭ����"))) <> "" Then
                    '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����

                    If IsNumeric(str��������) And Split(.TextMatrix(.Row, .ColIndex("ԭ����")), "||")(0) <> "0" Then
                        strxq = UCase(str��������)
'                            If Trim(.TextMatrix(.Row, mColЧ��)) = "" Then
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq, True)
                                If strxq = "" Then Exit Sub
                                .TextMatrix(.Row, .ColIndex("Ч��")) = Format(DateAdd("M", Split(.TextMatrix(.Row, .ColIndex("ԭ����")), "||")(0), strxq), "yyyy-mm-dd")
                                Call CheckLapse(.TextMatrix(.Row, .ColIndex("Ч��")))
                            End If
'                            End If
                    End If
                End If
             End If
        Case .ColIndex("����")
            OS.OpenIme (True)
             .ColComboList(.Col) = "..."
        End Select
    End With
End Sub
 
 '������������Ϣ��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, ByVal str����id As String, ByVal str��� As String, _
    ByVal str���� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, _
    ByVal numָ�������� As Double, ByVal strԭ���� As String, _
    ByVal intԭЧ�� As Integer, dbl����ϵ�� As Double, _
    ByVal int�Ƿ��� As Integer, ByVal int���÷��� As Integer, ByVal dblָ������� As Double, ByVal str��׼�ĺ� As String) As Boolean
    Dim sng�ֶ��ۼ� As Double
    Dim intCount As Integer, intCol As Integer, lngDepartid As Long
    Dim rsTemp As New ADODB.Recordset
    Dim dbl�ɱ���  As Double, dbl�ӳ��� As Double, int�ⷿ���� As Integer
    Dim strɢװ��λ As String
    Dim lngRow As Long
    
    On Error GoTo ErrHandle

    SetColValue = False
    lngDepartid = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    gstrSQL = "SELECT a.�ӳ��� from �������� a where a.����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ӳ���", lng����ID)
    dbl�ӳ��� = NVL(rsTemp!�ӳ���, 0) / 100
        
    gstrSQL = "SELECT nvl(A.����,0) ����,A.���Ч��,A.һ���Բ���,A.�ɱ���,A.�ⷿ����,A.ע��֤��,B.���㵥λ ɢװ��λ,Nvl(A.�Ƿ��������,0) As ������� " & _
              "From �������� A, �շ���ĿĿ¼ B Where a.����ID=b.id and  A.����id=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
    
    int�ⷿ���� = Val(zlStr.NVL(rsTemp!�ⷿ����))
    dbl�ɱ��� = zlStr.NVL(rsTemp!�ɱ���, 0)
    strɢװ��λ = zlStr.NVL(rsTemp!ɢװ��λ)
    
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> .ColIndex("�к�") Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, .ColIndex("�к�")) = intRow
        .TextMatrix(intRow, .ColIndex("����ID")) = lng����ID
        If Trim(.EditText) <> "" Then .EditText = str����id
        .TextMatrix(intRow, .ColIndex("������Ϣ")) = str����id
        .TextMatrix(intRow, .ColIndex("���")) = str���
        .TextMatrix(intRow, .ColIndex("һ���Բ���")) = zlStr.NVL(rsTemp!һ���Բ���)
        .TextMatrix(intRow, .ColIndex("�������")) = zlStr.NVL(rsTemp!�������)
        .TextMatrix(intRow, .ColIndex("���Ч��")) = zlStr.NVL(rsTemp!���Ч��)
        .TextMatrix(intRow, .ColIndex("����")) = IIf(IsNull(str����), "", str����)
        .TextMatrix(intRow, .ColIndex("��׼�ĺ�")) = IIf(IsNull(str��׼�ĺ�), "", str��׼�ĺ�)
        .TextMatrix(intRow, .ColIndex("��λ")) = str��λ
        .TextMatrix(intRow, .ColIndex("�ۼ�")) = Format(num�ۼ� * dbl����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, .ColIndex("ԭ����")) = IIf(IsNull(strԭ����), "", strԭ����)
        
        '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
        .TextMatrix(intRow, .ColIndex("ԭ����")) = IIf(IsNull(intԭЧ��), "0", intԭЧ��) & "||" & dbl�ӳ��� & "||" & int�Ƿ��� & "||" & int���÷��� & "||" & int�ⷿ����
        .TextMatrix(intRow, .ColIndex("�ɹ���")) = Format(numָ�������� * dbl����ϵ��, mFMT.FM_�ɱ���)
        .TextMatrix(intRow, .ColIndex("����ϵ��")) = dbl����ϵ��
        
        SetInputFormat intRow
        
        '˵�����������ַ�������Ͳ����������Ŀ������������ٶȡ�
        '�������Բ�����Щ��ֱ���õ�һ��SQL���ʵ�֣��������������ľͶ������ݿ���ɨ��һ�Ρ�
        If Val(int�ⷿ����) > 0 Then
'            If mintUnit = 1 Then
                gstrSQL = "" & _
                    "   Select �ϴβɹ���,�ϴβ���,�ϴ��������� " & _
                    "   From ҩƷ��� " & _
                    "   Where ����=1 and �ⷿid=[3] and ҩƷid=" & lng����ID & _
                    "        and nvl(����,0) =( select max(nvl(����,0)) " & _
                    "                           from ҩƷ��� " & _
                    "                           where ����=1 and �ⷿid=[1] and ҩƷid=[2] )"
'            Else
'            End If
        Else
            gstrSQL = "" & _
                "   Select �ϴβɹ���,�ϴβ���,�ϴ��������� " & _
                "   From ҩƷ��� " & _
                "   Where ����=1 and �ⷿid=[1] and ҩƷid=[2]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "--ȡ�ϴβɹ���", lngDepartid, lng����ID, lngDepartid)
        
        If Not rsTemp.EOF Then
            If .TextMatrix(intRow, .ColIndex("����")) = "" Then
                .TextMatrix(intRow, .ColIndex("����")) = IIf(IsNull(rsTemp.Fields(1)), "", rsTemp.Fields(1))
            End If
            .TextMatrix(intRow, .ColIndex("�ɹ���")) = Format(IIf(IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0)) * dbl����ϵ�� = 0, .TextMatrix(intRow, .ColIndex("�ɹ���")), IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0)) * dbl����ϵ��), mFMT.FM_�ɱ���)
            If IsNull(rsTemp!�ϴ���������) Then
                .TextMatrix(intRow, .ColIndex("��������")) = ""
            Else
                .TextMatrix(intRow, .ColIndex("��������")) = Format(rsTemp!�ϴ���������, "yyyy-mm-dd")
            End If
            
        Else
            If dbl�ɱ��� <> 0 Then .TextMatrix(intRow, .ColIndex("�ɹ���")) = Format(dbl�ɱ��� * dbl����ϵ��, mFMT.FM_�ɱ���)
        End If
        
        If .TextMatrix(intRow, .ColIndex("����")) <> "" Then
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, .ColIndex("����")), lng����ID)
            If rsTemp.RecordCount > 0 Then
               .TextMatrix(intRow, .ColIndex("��׼�ĺ�")) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
            End If
        End If
        
        'ʱ�۲��ϴ���
        If int�Ƿ��� = 1 Then
            .TextMatrix(intRow, .ColIndex("�ۼ�")) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                             ʱ�۲������ۼ�(lng����ID, Val(.TextMatrix(intRow, .ColIndex("�ɹ���"))), 0, -1, sng�ֶ��ۼ�)) _
                                                             , mFMT.FM_���ۼ�)
        End If
        .TextMatrix(intRow, .ColIndex("���۵�λ")) = strɢװ��λ
        '���˺�:���ۼ۴���
        Call �������ۼۼ����۲��(intRow)
    End With
    Call ��ʾ�����
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub mshBill_LostFocus()
    OS.OpenIme False
     Call zl_VsGridLOSTFOCUS(mshBill)
End Sub
Private Sub mshBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, strTemp As String, arrSplit As Variant
    Dim dbl�ӳ��� As Double, sng�ֶ��ۼ� As Double, dbl�ɹ��� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim dbl�ɹ��޼� As Double
    
    '������֤
    On Error GoTo ErrHandle
    With mshBill
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
          Case .ColIndex("�������")
                '�д���
                If strKey = "" Then Exit Sub
                strKey = zlCheckIsDate(strKey, .ColKey(Col))
                If strKey = "" Then Cancel = True: Exit Sub
                
                If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(Row, .ColIndex("���Ч��"))), CDate(strKey)), "yyyy-mm-dd") Then
                    If MsgBox("�������Ѿ��������ʧЧ��(" & Format(DateAdd("m", Val(.TextMatrix(Row, .ColIndex("���Ч��"))), CDate(strKey)), "yyyy-mm-dd") & "),�Ƿ�Ҫ�������!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                        Cancel = True
                        Exit Sub
                    End If
                End If
                '����ʧЧ��
                .TextMatrix(Row, .ColIndex("���ʧЧ��")) = Format(DateAdd("m", Val(.TextMatrix(Row, .ColIndex("���Ч��"))), CDate(strKey)), "yyyy-mm-dd")
                .EditText = strKey
           Case .ColIndex("��������"), .ColIndex("Ч��")
                '�д���
                If strKey = "" Then Exit Sub
                strKey = zlCheckIsDate(strKey, .ColKey(Col))
                If strKey = "" Then Cancel = True: Exit Sub
                .EditText = strKey
            Case .ColIndex("����")
                '����Ҳ�����Ӧ�Ĳ��أ�����������Ϊ����
                If strKey = "" Then Exit Sub
                If zlCommFun.StrIsValid(strKey, .EditMaxLength, , .ColKey(Col)) = False Then
                    Cancel = True
                End If
            Case .ColIndex("����")
                '����Ҳ�����Ӧ�Ĳ��أ�����������Ϊ����
                If strKey = "" Then Exit Sub
                If zlCommFun.StrIsValid(strKey, .EditMaxLength, , .ColKey(Col)) = False Then
                    Cancel = True
                End If
            Case .ColIndex("�ɹ���")
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    '���۸��ܴ�����ָ��������
                    gstrSQL = "Select nvl(a.ָ��������,0) as ָ��������, b.�ּ�" & vbNewLine & _
                                " From �������� A, �շѼ�Ŀ B" & vbNewLine & _
                                " Where a.����id = b.�շ�ϸĿid And Sysdate Between b.ִ������ And b.��ֹ���� And a.����id = [1]" & _
                                GetPriceClassString("B")
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯָ��������", Val(.TextMatrix(.Row, .ColIndex("����ID"))))
                    
                    dbl�ɹ��޼� = Format(rsTemp!ָ�������� * Val(.TextMatrix(Row, .ColIndex("����ϵ��"))), mFMT.FM_�ɱ���)
                    If mbln��ǿ�ƿ���ָ���۸� = False Then
                        If dbl�ɹ��޼� < Val(Format(Val(strKey), mFMT.FM_�ɱ���)) Then
                            MsgBox "��ǰ�۸������ָ��������" & dbl�ɹ��޼� & "��", vbInformation, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    .EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
                    
                    If .TextMatrix(Row, .ColIndex("ԭ����")) <> "" Then
                        '��ʱ�����ĵĴ���
                         '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                        arrSplit = Split(.TextMatrix(Row, .ColIndex("ԭ����")), "||")
                         If arrSplit(2) = 1 Then
                            'ʵ�����Ĵ���
                             .EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
                            If mbln�Ӽ��� And mbln�ֶμӳ��� = False Then
                                If Show�ӳ���(Col) = False Then Cancel = True: Exit Sub
                            Else
                                If mbln�ֶμӳ��� Then
                                    dbl�ӳ��� = 0 'Get�ֶμӳ���(Val(strkey)) / 100
                                    If Get�ֶμӳ��ۼ�(Val(strKey), Val(.TextMatrix(Row, .ColIndex("����ϵ��"))), mstrCaption, sng�ֶ��ۼ�) = False Then
                                        Cancel = True
                                        Exit Sub
                                    End If
                                    .TextMatrix(Row, .ColIndex("�ۼ�")) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                                          ʱ�۲������ۼ�(Val(.TextMatrix(Row, .ColIndex("����ID"))), Val(strKey), dbl�ӳ���, -1, sng�ֶ��ۼ�)) _
                                                                          , mFMT.FM_���ۼ�)
                                Else
                                    '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                                    dbl�ӳ��� = Val(arrSplit(1))
                                    .TextMatrix(Row, .ColIndex("�ۼ�")) = Format(У�����ۼ�(strKey * (1 + dbl�ӳ���) + _
                                                                          ʱ�۲������ۼ�(Val(.TextMatrix(Row, .ColIndex("����ID"))), strKey, dbl�ӳ���)) _
                                                                          , mFMT.FM_���ۼ�)
                                End If
                                If .TextMatrix(Row, .ColIndex("����")) <> "" Then
                                    .TextMatrix(Row, .ColIndex("�ۼ۽��")) = Format(.TextMatrix(Row, .ColIndex("����")) * .TextMatrix(Row, .ColIndex("�ۼ�")), mFMT.FM_���)
                                End If
                            End If
                         Else
                            '���ۼ��ɱ��۴������ۼ���ʾ
                            If Val(Format(rsTemp!�ּ�, mFMT.FM_���ۼ�)) < Val(Format(Val(strKey), mFMT.FM_�ɱ���)) Then
                                MsgBox "��ǰ�۸�������ۼۣ�", vbInformation, gstrSysName
                            End If
                         End If
                    End If
                End If
                '���ý��
                If strKey <> "" And strKey <> .TextMatrix(Row, .ColIndex("�ɹ���")) And .TextMatrix(Row, .ColIndex("����")) <> "" Then
                    .TextMatrix(Row, .ColIndex("�ɹ����")) = Format(.TextMatrix(Row, .ColIndex("����")) * strKey, mFMT.FM_���)
                    .TextMatrix(Row, .ColIndex("���")) = Format(IIf(.TextMatrix(Row, .ColIndex("�ۼ۽��")) = "", 0, .TextMatrix(Row, .ColIndex("�ۼ۽��"))) - IIf(.TextMatrix(Row, .ColIndex("�ɹ����")) = "", 0, .TextMatrix(Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
                End If
       
                
            Case .ColIndex("�ɹ����")
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(Row, .ColIndex("�ɹ����")) Then
                     If .TextMatrix(Row, .ColIndex("����")) <> "" Then
                            If mbln�Ӽ��� Then
                                'ȡ�øı�ɹ����ǰ�ļӼ���
                                mdbl�Ӽ��� = 15
                                If Val(.TextMatrix(Row, .ColIndex("�ۼ�"))) <> 0 And Val(.TextMatrix(Row, .ColIndex("�ɹ���"))) <> 0 Then
                                    mdbl�Ӽ��� = ����ӳ���(Val(.TextMatrix(Row, .ColIndex("����ID"))), Val(.TextMatrix(Row, .ColIndex("�ۼ�"))), Val(.TextMatrix(Row, .ColIndex("�ɹ���"))))
                                End If
                            End If
                            .TextMatrix(Row, .ColIndex("�ɹ���")) = Format(strKey / .TextMatrix(Row, .ColIndex("����")), mFMT.FM_�ɱ���)
                            
                            '��ʱ�����ĵĴ���
                            If .TextMatrix(Row, .ColIndex("ԭ����")) <> "" Then
                                '���¼������ۼۡ����
                                arrSplit = Split(.TextMatrix(Row, .ColIndex("ԭ����")), "||")
                                '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                                If Val(arrSplit(2)) = 1 Then
                                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                                    If mbln�Ӽ��� And mbln�ֶμӳ��� = False Then
                                        .TextMatrix(Row, .ColIndex("�ۼ�")) = Format(У�����ۼ�(Val(.TextMatrix(Row, .ColIndex("�ɹ���"))) * (1 + (mdbl�Ӽ��� / 100)) + _
                                                                              ʱ�۲������ۼ�(Val(.TextMatrix(Row, .ColIndex("����ID"))), Val(.TextMatrix(Row, .ColIndex("�ɹ���"))), (mdbl�Ӽ��� / 100))) _
                                                                              , mFMT.FM_���ۼ�)
                                        .TextMatrix(Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(Row, .ColIndex("�ۼ�"))) * Val(.TextMatrix(Row, .ColIndex("����"))), mFMT.FM_���)
                                        .TextMatrix(Row, .ColIndex("���")) = Format(IIf(.TextMatrix(Row, .ColIndex("�ۼ۽��")) = "", 0, .TextMatrix(Row, .ColIndex("�ۼ۽��"))) - IIf(.TextMatrix(Row, .ColIndex("�ɹ����")) = "", 0, .TextMatrix(Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
                                    Else
                                        Dim sng�ɹ��� As Double
                                        sng�ɹ��� = Val(.TextMatrix(Row, .ColIndex("�ɹ���")))
    
                                        If mbln�ֶμӳ��� Then
                                            dbl�ӳ��� = 0 ' Get�ֶμӳ���(Val(.TextMatrix(row, .colindex("�ɹ���")))) / 100
                                            If Get�ֶμӳ��ۼ�(sng�ɹ���, Val(.TextMatrix(Row, .ColIndex("����ϵ��"))), mstrCaption, sng�ֶ��ۼ�) = False Then
                                                Cancel = True
                                                Exit Sub
                                            End If
                                            .TextMatrix(Row, .ColIndex("�ۼ�")) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                                                  ʱ�۲������ۼ�(Val(.TextMatrix(Row, .ColIndex("����ID"))), sng�ɹ���, dbl�ӳ���, -1, sng�ֶ��ۼ�)) _
                                                                                  , mFMT.FM_���ۼ�)
                                        Else
                                            '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                                            dbl�ӳ��� = Val(arrSplit(1))
                                            .TextMatrix(Row, .ColIndex("�ۼ�")) = Format(У�����ۼ�(.TextMatrix(Row, .ColIndex("�ɹ���")) * (1 + dbl�ӳ���) + _
                                                                                  ʱ�۲������ۼ�(Val(.TextMatrix(Row, .ColIndex("����ID"))), Val(.TextMatrix(Row, .ColIndex("�ɹ���"))), dbl�ӳ���)) _
                                                                                  , mFMT.FM_���ۼ�)
                                        End If
                                        .TextMatrix(Row, .ColIndex("�ۼ۽��")) = Format(.TextMatrix(Row, .ColIndex("����")) * .TextMatrix(Row, .ColIndex("�ۼ�")), mFMT.FM_���)
                                    End If
                                End If
                            End If
                            .TextMatrix(Row, .ColIndex("���")) = Format(Val(.TextMatrix(Row, .ColIndex("�ۼ۽��"))) - Val(strKey), mFMT.FM_���)
                            .EditText = Format(strKey, mFMT.FM_���)
                            .TextMatrix(Row, .ColIndex("�ɹ����")) = .EditText
                    End If
                End If
 
            Case .ColIndex("����")
            
                If zlCommFun.DblIsValid(strKey, 16, True, True, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                strKey = Format(strKey, mFMT.FM_����)
                .EditText = strKey
                If .TextMatrix(Row, .ColIndex("�ɹ���")) <> "" Then
                    .TextMatrix(Row, .ColIndex("�ɹ����")) = Format(Val(.TextMatrix(Row, .ColIndex("�ɹ���"))) * Val(strKey), mFMT.FM_���)
                    'ʱ�����ĵĴ���
                    If .TextMatrix(Row, .ColIndex("ԭ����")) <> "" Then
                        arrSplit = Split(.TextMatrix(Row, .ColIndex("ԭ����")), "||")
                        '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                        If Val(arrSplit(2)) = 1 Then
                            '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                            If mbln�Ӽ��� Then
                                mdbl�Ӽ��� = Round(arrSplit(1) * 100, 2)
                                If Val(.TextMatrix(Row, .ColIndex("�ۼ�"))) <> 0 And Val(.TextMatrix(Row, .ColIndex("�ɹ���"))) <> 0 Then
                                    mdbl�Ӽ��� = ����ӳ���(Val(.TextMatrix(Row, .ColIndex("����ID"))), Val(.TextMatrix(Row, .ColIndex("�ۼ�"))), Val(.TextMatrix(Row, .ColIndex("�ɹ���"))))
                                End If
                                .TextMatrix(Row, .ColIndex("�ۼ�")) = Format(У�����ۼ�(Val(.TextMatrix(Row, .ColIndex("�ɹ���"))) * (1 + (mdbl�Ӽ��� / 100)) + _
                                                                      ʱ�۲������ۼ�(Val(.TextMatrix(Row, .ColIndex("����ID"))), Val(.TextMatrix(Row, .ColIndex("�ɹ���"))), (mdbl�Ӽ��� / 100))) _
                                                                      , mFMT.FM_���ۼ�)
                                .TextMatrix(Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(Row, .ColIndex("�ۼ�"))) * strKey, mFMT.FM_���)
                                .TextMatrix(Row, .ColIndex("���")) = Format(IIf(.TextMatrix(Row, .ColIndex("�ۼ۽��")) = "", 0, .TextMatrix(Row, .ColIndex("�ۼ۽��"))) - IIf(.TextMatrix(Row, .ColIndex("�ɹ����")) = "", 0, .TextMatrix(Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
                            Else
                                If mbln�ֶμӳ��� Then
                                    dbl�ӳ��� = 0 ' Get�ֶμӳ���(Val(.TextMatrix(row, .colindex("�ɹ���")))) / 100
                                    If Get�ֶμӳ��ۼ�(Val(.TextMatrix(Row, .ColIndex("�ɹ���"))), Val(.TextMatrix(Row, .ColIndex("����ϵ��"))), mstrCaption, sng�ֶ��ۼ�) = False Then
                                        Cancel = True
                                        Exit Sub
                                    End If
                                    .TextMatrix(Row, .ColIndex("�ۼ�")) = Format(У�����ۼ�(sng�ֶ��ۼ� + _
                                                                          ʱ�۲������ۼ�(Val(.TextMatrix(Row, .ColIndex("����ID"))), Val(.TextMatrix(Row, .ColIndex("�ɹ���"))), dbl�ӳ���, -1, sng�ֶ��ۼ�)) _
                                                                          , mFMT.FM_���ۼ�)
                                Else
                                    dbl�ӳ��� = Split(.TextMatrix(Row, .ColIndex("ԭ����")), "||")(1)
                                    .TextMatrix(Row, .ColIndex("�ۼ�")) = Format(У�����ۼ�(.TextMatrix(Row, .ColIndex("�ɹ���")) * (1 + dbl�ӳ���) + _
                                                                          ʱ�۲������ۼ�(Val(.TextMatrix(Row, .ColIndex("����ID"))), Val(.TextMatrix(Row, .ColIndex("�ɹ���"))), dbl�ӳ���)) _
                                                                          , mFMT.FM_���ۼ�)
                                End If
                            End If
                        End If
                    End If
                    If .TextMatrix(Row, .ColIndex("�ۼ�")) <> "" Then
                        .TextMatrix(Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(Row, .ColIndex("�ۼ�"))) * Val(strKey), mFMT.FM_���)
                    End If
                    .TextMatrix(Row, .ColIndex("���")) = Format(Val(.TextMatrix(Row, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
                End If
            Case .ColIndex("��������")
                
                If strKey = "" Then
                    MsgBox "���������������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    
                    If Val(strKey) > Val(.TextMatrix(Row, .ColIndex("����"))) Then
                        MsgBox "�����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    
                    strKey = Format(Val(strKey), mFMT.FM_����)
                    .Text = Val(strKey)
                    If .TextMatrix(Row, .ColIndex("�ɹ���")) <> "" Then
                        .TextMatrix(Row, .ColIndex("�ɹ����")) = Format(Val(.TextMatrix(Row, .ColIndex("�ɹ���"))) * Val(strKey), mFMT.FM_���)
                    End If
                    If .TextMatrix(Row, .ColIndex("�ۼ�")) <> "" Then
                        .TextMatrix(Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(Row, .ColIndex("�ۼ�"))) * Val(strKey), mFMT.FM_���)
                    End If
                    .TextMatrix(Row, .ColIndex("���")) = Format(Val(.TextMatrix(Row, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
                End If

            Case .ColIndex("�ۼ�")
                '�������:
                ' 1.�ۼ۲��ܴ���ָ�����ۼ�(���ݲ���:��ǿ�ƿ���ָ���۸����)
                ' 2.����˽�������ۼ�

                If Val(.TextMatrix(Row, .ColIndex("����ID"))) = 0 Then Exit Sub
                
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                
                If strKey <> "" Then
                    
                    If mbln��ǿ�ƿ���ָ���۸� = False Then
                        '�ж���������ۼ���ָ�����ۼ�
                        gstrSQL = "Select ָ�����ۼ� From �������� Where ����ID=[1] "
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[��ȡָ�����ۼ�]", Val(.TextMatrix(Row, .ColIndex("����ID"))))
                        Dim dblָ�����ۼ� As Double
                        dblָ�����ۼ� = Val(zlStr.NVL(rsTemp!ָ�����ۼ�))
                        dblָ�����ۼ� = Val(Format(dblָ�����ۼ� * Val(.TextMatrix(Row, .ColIndex("����ϵ��"))), mFMT.FM_���ۼ�))
                        If Val(Format(Val(strKey), mFMT.FM_���ۼ�)) > dblָ�����ۼ� Then
                            ShowMsgBox "�ۼ۲��ܴ���ָ�����ۼۣ�ָ�����ۼۣ���" & dblָ�����ۼ� & "��"
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    If Val(strKey) < Val(.TextMatrix(Row, .ColIndex("�ɹ���"))) Then
                        If MsgBox("ע�⣺" & vbCrLf & "     �ۼ�(��" & Format(Val(strKey), mFMT.FM_���ۼ�) & " С����" & vbCrLf & "     �ɹ��ۣ���" & Format(Val(.TextMatrix(Row, .ColIndex("�ɹ���"))), mFMT.FM_�ɱ���) & "��,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                strKey = Format(Val(strKey), mFMT.FM_���ۼ�)
                .EditText = strKey
                '������
                .TextMatrix(Row, .ColIndex("�ۼ۽��")) = Format(Val(strKey) * Val(.TextMatrix(Row, .ColIndex("����"))), mFMT.FM_���)
                .TextMatrix(Row, .ColIndex("���")) = Format(Val(.TextMatrix(Row, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
        Case .ColIndex("���ۼ�")
                '�������:
                ' 1.�ۼ۲��ܴ���ָ�����ۼ�(���ݲ���:��ǿ�ƿ���ָ���۸����)
                ' 2.����˽�������ۼ�
                If Val(.TextMatrix(Row, .ColIndex("����ID"))) = 0 Then Exit Sub
                
                If zlCommFun.DblIsValid(strKey, 16, True, False, 0, .ColKey(Col)) = False Then
                   Cancel = True: Exit Sub
                End If
                If strKey <> "" Then
                    If mbln��ǿ�ƿ���ָ���۸� = False Then
                        '�ж���������ۼ���ָ�����ۼ�
                        gstrSQL = "Select ָ�����ۼ� From �������� Where ����ID=[1] "
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[��ȡָ�����ۼ�]", Val(.TextMatrix(Row, .ColIndex("����ID"))))
                        dblָ�����ۼ� = Val(zlStr.NVL(rsTemp!ָ�����ۼ�))
                        dblָ�����ۼ� = Val(Format(dblָ�����ۼ�, mFMT.FM_ɢװ���ۼ�))
                        If Val(Format(Val(strKey), mFMT.FM_ɢװ���ۼ�)) > dblָ�����ۼ� Then
                            ShowMsgBox "���ۼ۲��ܴ���ָ�����ۼۣ�ָ�����ۼۣ���" & dblָ�����ۼ� & "��"
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    If Val(.TextMatrix(Row, .ColIndex("����ϵ��"))) = 0 Then
                        dbl�ɹ��� = Val(.TextMatrix(Row, .ColIndex("�ɹ���")))
                    Else
                        dbl�ɹ��� = Val(.TextMatrix(Row, .ColIndex("�ɹ���"))) / Val(.TextMatrix(Row, .ColIndex("����ϵ��")))
                    End If
                    
                    If Val(strKey) < dbl�ɹ��� Then
                        If MsgBox("ע�⣺" & vbCrLf & "     ���ۼ�(��" & Format(Val(strKey), mFMT.FM_ɢװ���ۼ�) & " С����" & vbCrLf & "     ����ۣ���" & Format(dbl�ɹ���, mFMT.FM_�ɱ���) & "��,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    strKey = Format(Val(strKey), mFMT.FM_ɢװ���ۼ�)
                    .EditText = strKey
                    .TextMatrix(.Row, .Col) = strKey
                    '���˺�:���ۼ۴���
                    Call �������ۼۼ����۲��(.Row, False)
                    If strKey <> "" Then
                        .TextMatrix(.Row, .ColIndex("�ۼ�")) = Format(Val(strKey) * Val(.TextMatrix(.Row, .ColIndex("����ϵ��"))), mFMT.FM_���ۼ�)
                        .TextMatrix(.Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ۼ�"))) * Val(.TextMatrix(.Row, .ColIndex("����"))), mFMT.FM_���)
                        .TextMatrix(.Row, .ColIndex("���")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(.Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
                    End If
                    ��ʾ�ϼƽ��
                Else
                    strKey = Format(Val(strKey), mFMT.FM_ɢװ���ۼ�)
                    .EditText = strKey
                End If
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Sub msh����_DblClick()
    msh����_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    
    With mshBill
        If KeyCode = vbKeyEscape Then
            msh����.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .ColIndex("����")) = msh����.TextMatrix(msh����.Row, 2)
            
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, .ColIndex("����")), .ColIndex("����ID"))
            If rsTemp.RecordCount Then
                .TextMatrix(.Row, .ColIndex("��׼�ĺ�")) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
            Else
                .TextMatrix(.Row, .ColIndex("��׼�ĺ�")) = ""
            End If
            
            msh����.Visible = False
            .Col = .ColIndex("����")
            .SetFocus
        End If
    End With
End Sub

Private Sub msh����_LostFocus()
    If msh����.Visible Then
        msh����.Visible = False
    End If
End Sub
Private Function ��ǰ��Ϊ�ⷿ() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ�ⷿ��Ϊ�ⷿ
    '���:
    '����:
    '����:����true��ʾ��Ϊ�ⷿ,����Ϊ(���ϲ��Ż��Ƽ���)
    '����:���˺�
    '����:2008-12-03 11:23:18
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From ��������˵�� " & _
        "   WHERE ((�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')) " & _
        "        AND ����id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
    If rsTemp.Fields(0) > 0 Then
        ��ǰ��Ϊ�ⷿ = False
        mbln�ⷿ = False
    Else
        ��ǰ��Ϊ�ⷿ = True
        mbln�ⷿ = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��֤���ݵĺϷ���
    '���:
    '����:
    '����:��ѹ������,����true,���򷵻�False
    '����:���˺�
    '����:2008-12-03 09:49:18
    '-----------------------------------------------------------------------------------------------------------
    Dim intLop As Integer, rsTemp As New Recordset, blnStock As Boolean, arrSplit As Variant
    ValidData = False
    blnStock = ��ǰ��Ϊ�ⷿ()
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "���ݺŲ���Ϊ��"
            Exit Function
        End If
        
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
            Exit Function
        End If
        
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "���ݺų���,���������" & CInt(txtNO.MaxLength / 2) & "�����֣���ò�Ҫ���֣���" & txtNO.MaxLength & "���ַ�!"
            txtNO.SetFocus
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, .ColIndex("����ID")) <> "" Then        '�����з�����
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, .ColIndex("������Ϣ"))) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, .ColIndex("����")))) = "" Then
                        MsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("����")
                        Exit Function
                    End If
'
 
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, .ColIndex("����")))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "��" & intLop & "�����ĵ����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("����")
                        Exit Function
                    End If
                    
                    If Len(Trim(.TextMatrix(intLop, .ColIndex("��Ʒ����")))) > 50 Then
                        MsgBox "��" & intLop & "�����ĵ���Ʒ���볬��,���������50���ַ�!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("��Ʒ����")
                        Exit Function
                    End If
                    
                    If blnStock = True Then
                        '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                        arrSplit = Split(.TextMatrix(intLop, .ColIndex("ԭ����")) & "||||||||||", "||")
                        If Val(arrSplit(0)) <> 0 Then
                            If .TextMatrix(intLop, .ColIndex("����")) = "" Or .TextMatrix(intLop, .ColIndex("Ч��")) = "" Then
                                MsgBox "��" & intLop & "�е�������Ч������,����������ż�Ч����Ϣ�������뵥���У�", vbInformation, gstrSysName
                                mshBill.SetFocus: .Row = intLop: .TopRow = intLop
                                If .TextMatrix(intLop, .ColIndex("����")) = "" Then
                                    .Col = .ColIndex("����")
                                Else
                                    .Col = .ColIndex("Ч��")
                                End If
                                Exit Function
                            End If
                        End If
                        
                        If Val(arrSplit(4)) <> 0 Then '�ⷿ����
                            If mbln�����������Ų��ؿ��� = True Then
                                If .TextMatrix(intLop, .ColIndex("����")) = "" Or .TextMatrix(intLop, .ColIndex("����")) = "" Then
                                    MsgBox "��" & intLop & "�е������Ƿ�������,������Ĳ��غ�����" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                                    mshBill.SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    If .TextMatrix(intLop, .ColIndex("����")) = "" Then
                                        .Col = .ColIndex("����")
                                    Else
                                        .Col = .ColIndex("����")
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If

                    Else '�����ǡ����ϲ��š�
                        '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                        arrSplit = Split(.TextMatrix(intLop, .ColIndex("ԭ����")) & "||||||||||", "||")
                        If Val(arrSplit(3)) <> 0 Then '���÷���
                            '�洢��ʽֵ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                            If arrSplit(0) <> "0" Then
                                If .TextMatrix(intLop, .ColIndex("����")) = "" Or .TextMatrix(intLop, .ColIndex("Ч��")) = "" Then
                                    MsgBox "��" & intLop & "�е�����������Ч�ڲ���,����������ż�Ч��" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                                    mshBill.SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    If .TextMatrix(intLop, .ColIndex("����")) = "" Then
                                        .Col = .ColIndex("����")
                                    Else
                                        .Col = .ColIndex("Ч��")
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    
                        '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
                        arrSplit = Split(.TextMatrix(intLop, .ColIndex("ԭ����")) & "||||||||||", "||")
                        If Val(arrSplit(3)) <> 0 Then '���÷���
                            If mbln�����������Ų��ؿ��� = True Then
                                If .TextMatrix(intLop, .ColIndex("����")) = "" Or .TextMatrix(intLop, .ColIndex("����")) = "" Then
                                    MsgBox "��" & intLop & "�е������Ƿ�������,������Ĳ��غ�����" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                                    mshBill.SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    If .TextMatrix(intLop, .ColIndex("����")) = "" Then
                                        .Col = .ColIndex("����")
                                    Else
                                        .Col = .ColIndex("����")
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, .ColIndex("�ɹ���"))) > 9999999999# Then
                        MsgBox "  ��" & intLop & "�����ĵĲɹ��۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        .Row = intLop: .TopRow = intLop: .Col = .ColIndex("�ɹ���")
                        mshBill.SetFocus
                        Exit Function
                    End If
                    
                    If zlCommFun.DblIsValid(.TextMatrix(intLop, .ColIndex("���ۼ�")), 16, False, False, , "��" & intLop & "���������ϵ����ۼ�") = False Then
                        mshBill.SetFocus
                        .Row = intLop: .TopRow = intLop: .Col = .ColIndex("���ۼ�")
                        Exit Function
                    End If
                    If zlCommFun.DblIsValid(.TextMatrix(intLop, .ColIndex("���۽��")), 16, False, False, , "��" & intLop & "���������ϵ����۽��") = False Then
                        mshBill.SetFocus
                        .Row = intLop: .TopRow = intLop: .Col = .ColIndex("���ۼ�")
                        Exit Function
                    End If
                    If zlCommFun.DblIsValid(.TextMatrix(intLop, .ColIndex("���۲��")), 16, False, False, , "��" & intLop & "���������ϵ����۲��") = False Then
                        mshBill.SetFocus
                        .Row = intLop: .TopRow = intLop: .Col = .ColIndex("���ۼ�")
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, .ColIndex("����"))) > 9999999999# Then
                        MsgBox "��" & intLop & "�����ĵ��������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("����")
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, .ColIndex("�ɹ����"))) > 9999999999999# Then
                        MsgBox "��" & intLop & "�����ĵĲɹ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("�ɹ����")
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, .ColIndex("�ۼ۽��"))) > 9999999999999# Then
                        MsgBox "��" & intLop & "�����ĵ��ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = .ColIndex("����")
                        Exit Function
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
End Function


Private Function SaveCard() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������ⵥ��Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-12-03 10:00:08
    '-----------------------------------------------------------------------------------------------------------

    Dim chrNo As Variant, cllPro As New Collection
    Dim lng��� As Long, lng�ⷿid As Long, lng������ID As Long, lng����ID As Long, intRow As Integer
    Dim dbl���� As Double, dbl�ɹ��� As Double, dbl�ɹ���� As Double
    Dim dbl���۽�� As Double, dbl��� As Double, str���۲�� As String, dbl���ۼ� As Double
    Dim strժҪ As String, str������ As String, str�������� As String, str�������� As String
    Dim str����� As String, str������� As String, str���Ч�� As String
    Dim str���� As String, str���� As String, strЧ�� As String, str��Ʒ���� As String
    Dim str��׼�ĺ� As String
    Dim n As Long
    
    
    SaveCard = False
    With mshBill
        chrNo = Trim(txtNO)
        lng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
        If mint�༭״̬ = 1 Then   'mbln�������� Or
            If chrNo <> "" Then
                If CheckNOExists(70, chrNo) Then Exit Function
            End If
            If chrNo = "" Then chrNo = sys.GetNextNo(70, lng�ⷿid)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        lng������ID = cboType.ItemData(cboType.ListIndex)
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str����� = Txt�����
        
        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_�����������_Delete('" & mstr���ݺ� & "')"
            AddArray cllPro, gstrSQL
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("����ID")) <> "" Then
                lng����ID = Val(.TextMatrix(intRow, .ColIndex("����ID")))
                str���� = .TextMatrix(intRow, .ColIndex("����"))
                str��׼�ĺ� = .TextMatrix(intRow, .ColIndex("��׼�ĺ�"))
                str���� = .TextMatrix(intRow, .ColIndex("����"))
                strЧ�� = IIf(.TextMatrix(intRow, .ColIndex("Ч��")) = "", "", .TextMatrix(intRow, .ColIndex("Ч��")))
                
                dbl���� = Round(Val(.TextMatrix(intRow, .ColIndex("����"))) * Val(.TextMatrix(intRow, .ColIndex("����ϵ��"))), g_С��λ��.obj_���С��.����С��)
                dbl�ɹ��� = Round(Val(.TextMatrix(intRow, .ColIndex("�ɹ���"))) / Val(.TextMatrix(intRow, .ColIndex("����ϵ��"))), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɹ���� = Round(Val(.TextMatrix(intRow, .ColIndex("�ɹ����"))), g_С��λ��.obj_���С��.���С��)
                
                
   
                '���˺�:���ۼ۴���
                '���ݿ��е�:��� = ���۽�� - ������
                '���ݿ��е�:�÷� = ���۽��-�ۼ۽������۲��-���(�ⷿ��λ�Ĳ��)

                dbl���ۼ� = Round(Val(.TextMatrix(intRow, .ColIndex("���ۼ�"))), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, .ColIndex("���۽��"))), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl��� = Round(Val(.TextMatrix(intRow, .ColIndex("���۲��"))), g_С��λ��.obj_���С��.���ۼ�С��)
                str���۲�� = Round(Val(.TextMatrix(intRow, .ColIndex("���۲��"))) - Val(.TextMatrix(intRow, .ColIndex("���"))), g_С��λ��.obj_���С��.���ۼ�С��)
'                dbl�ۼ� = Round(Val(.TextMatrix(intRow, .ColIndex("�ۼ�"))) / Val(.TextMatrix(intRow, .ColIndex("����ϵ��"))), g_С��λ��.obj_ɢװС��.���ۼ�С��)
'                dbl���۽�� = Round(Val(.TextMatrix(intRow, .ColIndex("�ۼ۽��"))), g_С��λ��.obj_ɢװС��.���С��)
'                dbl��� = Round(Val(.TextMatrix(intRow, .ColIndex("���"))), g_С��λ��.obj_ɢװС��.���С��)
                
                str�������� = Trim(IIf(.TextMatrix(intRow, .ColIndex("��������")) = "", "", .TextMatrix(intRow, .ColIndex("��������"))))
                str������� = Trim(IIf(.TextMatrix(intRow, .ColIndex("�������")) = "", "", .TextMatrix(intRow, .ColIndex("�������"))))
                str���Ч�� = Trim(IIf(.TextMatrix(intRow, .ColIndex("���ʧЧ��")) = "", "", .TextMatrix(intRow, .ColIndex("���ʧЧ��"))))
                If gblnCode = True Then str��Ʒ���� = Trim(IIf(.TextMatrix(intRow, .ColIndex("��Ʒ����")) = "", "", .TextMatrix(intRow, .ColIndex("��Ʒ����"))))
                
                lng��� = intRow
                
                'Zl_�����������_Insert
                gstrSQL = "zl_�����������_INSERT("
                '  No_In         In ҩƷ�շ���¼.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                gstrSQL = gstrSQL & "" & lng�ⷿid & ","
                '  ������id_In In ҩƷ�շ���¼.������id%Type,
                gstrSQL = gstrSQL & "" & lng������ID & ","
                '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type,
                gstrSQL = gstrSQL & "" & dbl���� & ","
                '  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type,
                gstrSQL = gstrSQL & "" & dbl�ɹ��� & ","
                '  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type,
                gstrSQL = gstrSQL & "" & dbl�ɹ���� & ","
                '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type,
                gstrSQL = gstrSQL & "" & dbl���ۼ� & ","
                '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type,
                gstrSQL = gstrSQL & "" & dbl���۽�� & ","
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & dbl��� & ","
                '  ���۲��_In   In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & str���۲�� & ","
                '  ������_In     In ҩƷ�շ���¼.������%Type,
                gstrSQL = gstrSQL & "'" & str������ & "',"
                '  ��������_In   In ҩƷ�շ���¼.��������%Type,
                gstrSQL = gstrSQL & "to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS'),"
                '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                gstrSQL = gstrSQL & "'" & strժҪ & "',"
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str�������� = "", "Null", "to_date('" & Format(str��������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
                gstrSQL = gstrSQL & "" & IIf(str������� = "", "Null", "to_date('" & Format(str�������, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null
                gstrSQL = gstrSQL & "" & IIf(str���Ч�� = "", "Null", "to_date('" & Format(str���Ч��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ","
                '  ��Ʒ����_In   In ҩƷ�շ���¼.��Ʒ����%Type := Null
                gstrSQL = gstrSQL & "'" & str��Ʒ���� & "',"
                '  ��׼�ĺ�_In   In ҩƷ�շ���¼.��׼�ĺ�%Type := Null
                gstrSQL = gstrSQL & IIf(str��׼�ĺ� = "", "NULL", "'" & str��׼�ĺ� & "'")
                gstrSQL = gstrSQL & ")"
                AddArray cllPro, gstrSQL
            End If
            
            recSort.MoveNext
        Next
    End With
    
    err = 0: On Error GoTo ErrHandle:
    ExecuteProcedureArrAy cllPro, mstrCaption, True
    If Not ��鵥��(17, txtNO.Tag) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    gcnOracle.CommitTrans
    
    mblnSave = True: mblnSuccess = True: mblnChange = False: SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������ⵥ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-12-03 10:44:27
    '-----------------------------------------------------------------------------------------------------------

    Dim cllProc As New Collection, blnȫ�� As Boolean
    Dim lng�д� As Integer, lngԭ��¼״̬ As Integer, lng��� As Integer, intRow As Integer
    Dim lng����ID As Long, dbl�������� As Double, dblʵ������ As Double, dbl���ۼ� As Double
    Dim chrNo As String, str������ As String, str�������� As String
    Dim int����� As Integer, lng�ⷿid As Long, lng���� As Long
    Dim strժҪ As String
    Dim n As Long
    
    SaveStrike = False
    With mshBill
        '����������������С����
        lng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
        int����� = Get������(cboStock.ItemData(cboStock.ListIndex))
        chrNo = Trim(txtNO.Tag)
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("��������"))) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, .ColIndex("����"))), Val(.TextMatrix(intRow, .ColIndex("��������")))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
                If int����� <> 0 Then
                    dbl�������� = Round(Val(.TextMatrix(intRow, .ColIndex("��������"))) * Val(.TextMatrix(intRow, .ColIndex("����ϵ��"))), g_С��λ��.obj_ɢװС��.����С��)
                    dblʵ������ = Round(Val(.TextMatrix(intRow, .ColIndex("����"))) * Val(.TextMatrix(intRow, .ColIndex("����ϵ��"))), g_С��λ��.obj_ɢװС��.����С��)
                    blnȫ�� = (dbl�������� = dblʵ������)
                    If blnȫ�� Then
                        dbl�������� = Val(.TextMatrix(intRow, .ColIndex("��ʵ����")))
                    End If
                    lng���� = ȡ��������(17, chrNo, Val(.TextMatrix(intRow, .ColIndex("����ID"))), Val(.TextMatrix(intRow, .ColIndex("���"))))
                    If Check��������(lng�ⷿid, Val(.TextMatrix(intRow, .ColIndex("����ID"))), lng����, dbl��������, int�����) = False Then Exit Function
                End If
            End If
        Next
    
        str������ = UserInfo.�û���
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        lngԭ��¼״̬ = mint��¼״̬
        
        lng�д� = 0
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("����ID")) <> "" And Val(.TextMatrix(intRow, .ColIndex("��������"))) <> 0 Then
                lng�д� = lng�д� + 1
                
                lng����ID = .TextMatrix(intRow, .ColIndex("����ID"))
                dbl�������� = Round(Val(.TextMatrix(intRow, .ColIndex("��������"))) * Val(.TextMatrix(intRow, .ColIndex("����ϵ��"))), g_С��λ��.obj_ɢװС��.����С��)
                dblʵ������ = Round(Val(.TextMatrix(intRow, .ColIndex("����"))) * Val(.TextMatrix(intRow, .ColIndex("����ϵ��"))), g_С��λ��.obj_ɢװС��.����С��)
                strժҪ = txtժҪ.Text
                
                blnȫ�� = (dbl�������� = dblʵ������)
                lng��� = Val(.TextMatrix(intRow, .ColIndex("���")))
                'Zl_�����������_Strike
                gstrSQL = "Zl_�����������_Strike("
                '  �д�_In       In Integer,
                gstrSQL = gstrSQL & "" & lng�д� & ","
                '  ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
                gstrSQL = gstrSQL & "" & lngԭ��¼״̬ & ","
                '  No_In         In ҩƷ�շ���¼.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
                gstrSQL = gstrSQL & "" & dbl�������� & ","
                '  ������_In     In ҩƷ�շ���¼.������%Type,
                gstrSQL = gstrSQL & "'" & str������ & "',"
                '  ��������_In   In ҩƷ�շ���¼.��������%Type,
                gstrSQL = gstrSQL & "to_date('" & Format(str��������, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS'),"
                '  ȫ������_In   In ҩƷ�շ���¼.ʵ������%Type := 0 --1-ȫ������,0-���ֳ���
                gstrSQL = gstrSQL & "" & IIf(blnȫ��, 1, 0)
                gstrSQL = gstrSQL & ",'" & strժҪ & "')"
                AddArray cllProc, gstrSQL
            End If
            
            recSort.MoveNext
        Next
        If lng�д� = 0 Then
            MsgBox "û��ѡ��һ�����������������ܳ��������飡", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With
    err = 0: On Error GoTo ErrHandle
    ExecuteProcedureArrAy cllProc, mstrCaption
    mblnSave = True: mblnSuccess = True: mblnChange = False
    
    SaveStrike = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, .ColIndex("�ɹ����")))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, .ColIndex("�ۼ۽��")))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & Format(curTotal, mFMT.FM_���)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���)
    lblDifference.Caption = "��ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
    
    
End Sub

Private Sub ��ʾ�����()
    Dim recTmp As New ADODB.Recordset
    Dim dbl���� As Double
    Dim str��λ As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo ErrHandle
    If mshBill.TextMatrix(mshBill.Row, mshBill.ColIndex("������Ϣ")) = "" Then
        stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, mshBill.ColIndex("����ID")) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, mshBill.ColIndex("����ID"))
    Select Case mintUnit
        Case 0
            strQuantity = "a.��������"
        Case Else
            strQuantity = "a.��������/b.����ϵ�� "
    End Select
        
    gstrSQL = "" & _
        "   Select b.����ID," & IIf(mintUnit = 0, "M.���㵥λ", "b.��װ��λ") & " as ��λ, Sum(" & strQuantity & ") as ���� " & _
        "   From ҩƷ��� a,�������� b,�շ���ĿĿ¼ M " & _
        "   Where a.����=1 and a.ҩƷid=b.����id and a.ҩƷid=M.id and a.��������<>0 And " & _
        "           a.�ⷿID=[1]" & _
        "           and b.����ID=[2]" & _
        "   Group by b.����ID," & IIf(mintUnit = 0, "m.���㵥λ", "b.��װ��λ")
    Set recTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ʾ�����", cboStock.ItemData(cboStock.ListIndex), intID)
        
    With recTmp
        If .EOF Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        dbl���� = IIf(IsNull(!����), 0, !����)
        
        stbThis.Panels(2).Text = "�����ĵ�ǰ�����Ϊ[" & Format(dbl����, mFMT.FM_����) & "]" & zlStr.NVL(!��λ)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    ImeLanguage True
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    ImeLanguage False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


'��ӡ����
Private Sub printbill()
    Dim strUnit As String
    Dim int��λϵ�� As Integer
    Dim strNo As String
    
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1714", mint��¼״̬, mintUnit, 1714, "����������ⵥ", strNo
End Sub

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
        
    zlDatabase.OpenRecordset rsBatchNolen, gstrSQL, "ȡ�ֶγ���"
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function ʱ�۲������ۼ�(ByVal lng����ID As Long, ByVal sin�ɹ��� As Double, ByVal sin�ӳ��� As Double, _
    Optional LngLastRow As Long = -1, Optional sng�ۼ� As Double = -99999999) As Double
    '------------------------------------------------------------------------------------------------------
    '����:����ָ���۸���۱ȼ����ʱ�۲��ϵĲ���������
    '���:lng����ID-����ID
    '     sin�ɹ���-�ɹ��۸�
    '     sin�ӳ���-�ӳ���(�������0,ͬʱ�ִ���dbl���ۼ�,�򽫰���������ۼ۽��м���)
    '     LngLastRow-���ݵ��к�
    '     sng�ۼ�-��������ۼ�
    '����:
    '����:���ۼ۵��������
    '�޸���:���˺�
    '�޸�ʱ��:2007/2/25
    '------------------------------------------------------------------------------------------------------
    'ʱ�۲������ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
    '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    
    Dim sin���ۼ� As Double, sinָ�����ۼ� As Double, sin��������� As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ָ�����ۼ�,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", lng����ID)
    
    If rsTemp.EOF Then Exit Function
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sin��������� = rsTemp!���������
    
    ʱ�۲������ۼ� = 0
    If sin��������� = 100 Then Exit Function
    If sinָ�����ۼ� = 0 Then Exit Function
    If LngLastRow = -1 Then LngLastRow = mshBill.Row
    
    sin���ۼ� = sin�ɹ��� * (1 + sin�ӳ���)
    If sin���ۼ� / Val(mshBill.TextMatrix(LngLastRow, mshBill.ColIndex("����ϵ��"))) >= sinָ�����ۼ� Then Exit Function
    sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(LngLastRow, mshBill.ColIndex("����ϵ��")))
    ʱ�۲������ۼ� = (sinָ�����ۼ� - sin���ۼ�) * (1 - sin��������� / 100)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ����ӳ���(ByVal lng����ID As Long, ByVal sin���ۼ� As Double, ByVal sin�ɱ��� As Double) As Double
    Dim sinָ�����ۼ� As Double, sin��������� As Double
    Dim rsTemp As New ADODB.Recordset
    '�������ۼ۷���ɱ���,����ʱ�۲��Ϲ�ʽ�ı仯,����ԭ������ӳ��ʵĹ�ʽ��Ч,�����¼���
    'ԭ��ʽ:(���ۼ�/�ɱ���-1)*100
    '�ֹ�ʽ������:�������ۼ��ǰ��ӳ����������,�ټ������������ǲ��ֽ��,���ʵ�ʰ��ӳ�����������ۼ�=ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
    '������ԭ��ʽ���ʵ�ʵļӳ���
    ����ӳ��� = 0.15
    
    On Error GoTo ErrHandle
    gstrSQL = "Select A.ָ�����ۼ�,Nvl(A.���������,100) ���������,Nvl(B.�Ƿ���,0) ʱ�� From �������� A,�շ���ĿĿ¼ B Where A.����id=B.id and A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", lng����ID)
    
    If rsTemp.EOF Then Exit Function
    
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sin��������� = rsTemp!���������
    If rsTemp!ʱ�� = 0 Then Exit Function
'   If mbln�ֶμӳ��� Then
'            ����ӳ��� = Get�ֶμӳ���(sin�ɱ���)
'   Else
        'ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
        sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(mshBill.Row, mshBill.ColIndex("����ϵ��")))
        If sin��������� <> 100 And sin��������� > 0 Then
            sin���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - sin���ۼ�) / sin��������� * 100
        Else
            sin���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - sin���ۼ�)
        End If
        ����ӳ��� = (sin���ۼ� / sin�ɱ��� - 1) * 100
'    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function У�����ۼ�(ByVal sin���ۼ� As Double, Optional LngLastRow As Long = -1) As Double
    '�õ�����ǰ��λϵ�����������ָ�����ۼۣ����ʱ������ǿ�ƿ���ָ���ۼ�����������ۼ۴���ָ�����ۼۣ���ָ�����ۼ�Ϊ׼
    Dim sinָ�����ۼ� As Double
    Dim rsTemp As New ADODB.Recordset
    If LngLastRow = -1 Then LngLastRow = mshBill.Row
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ָ�����ۼ�,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����ۼ�", Val(mshBill.TextMatrix(LngLastRow, mshBill.ColIndex("����ID"))))
    
    If rsTemp.EOF Then Exit Function
    sinָ�����ۼ� = zlStr.NVL(rsTemp!ָ�����ۼ�, mshBill.ColIndex("����ID"))
    sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(LngLastRow, mshBill.ColIndex("����ϵ��")))
    If sinָ�����ۼ� = 0 Then sinָ�����ۼ� = sin���ۼ�
    У�����ۼ� = IIf(sin���ۼ� > sinָ�����ۼ� And Not mbln��ǿ�ƿ���ָ���۸�, sinָ�����ۼ�, sin���ۼ�)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function �������ۼۼ����۲��(ByVal lngRow As Long, Optional bln���ۼ� As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݿⷿ��λ����ɢװ��λ�����ۼۼ����
    '���:lngRow -ָ���������
    '     bln���ۼ�-���ۼ�Ϊ�ۼ�
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-28 12:09:04
    '-----------------------------------------------------------------------------------------------------------
    Dim dbl����ϵ�� As Double, arrSplit As Variant
    Dim dbl���� As Double
    
    With mshBill
        dbl����ϵ�� = Val(.TextMatrix(lngRow, mshBill.ColIndex("����ϵ��")))
        dbl���� = IIf(mint�༭״̬ = 6, Val(.TextMatrix(lngRow, .ColIndex("��������"))), Val(.TextMatrix(lngRow, .ColIndex("����"))))
        If dbl���� = 0 Or Val(.TextMatrix(lngRow, .ColIndex("����ID"))) = 0 Then
            .TextMatrix(lngRow, .ColIndex("���۽��")) = 0
            .TextMatrix(lngRow, .ColIndex("���۲��")) = 0
            .TextMatrix(lngRow, .ColIndex("���")) = 0
            .TextMatrix(lngRow, .ColIndex("�ۼ۽��")) = 0
            .TextMatrix(lngRow, .ColIndex("���ۼ�")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ۼ�"))) / IIf(dbl����ϵ�� = 0, 1, dbl����ϵ��), mFMT.FM_ɢװ���ۼ�)
            Exit Function
        End If
        '�洢��ʽ:���Ч��||ָ�������||�Ƿ���||���÷���||�ⷿ����
        If .TextMatrix(lngRow, .ColIndex("ԭ����")) <> "" Then
           arrSplit = Split(.TextMatrix(lngRow, .ColIndex("ԭ����")), "||")
           If Val(arrSplit(2)) = 1 And (IIf(mbln�ⷿ, arrSplit(4) = 1, arrSplit(3) = 1)) Then
                'ʵ������
                '���˺�:���ۼ۴���
                If bln���ۼ� Then
                    .TextMatrix(lngRow, .ColIndex("���ۼ�")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ۼ�"))) / dbl����ϵ��, mFMT.FM_ɢװ���ۼ�)
                End If
                If Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("���ۼ�")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ۼ�"))) / dbl����ϵ��, mFMT.FM_ɢװ���ۼ�)
                End If
                .TextMatrix(lngRow, .ColIndex("���۽��")) = Format(Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) * (dbl���� * dbl����ϵ��), mFMT.FM_���)
                '���۲��=���۽��-������
                .TextMatrix(lngRow, .ColIndex("���۲��")) = Format(Val(.TextMatrix(lngRow, .ColIndex("���۽��"))) - Val(.TextMatrix(lngRow, .ColIndex("�ɹ����"))), mFMT.FM_���)
           Else '����
                .TextMatrix(lngRow, .ColIndex("���ۼ�")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ۼ�"))) / dbl����ϵ��, mFMT.FM_ɢװ���ۼ�)
                .TextMatrix(lngRow, .ColIndex("���۽��")) = Format(Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) * (dbl���� * dbl����ϵ��), mFMT.FM_���)
                '���۲��=���۽��-������
                .TextMatrix(lngRow, .ColIndex("���۲��")) = Format(Val(.TextMatrix(lngRow, .ColIndex("���۽��"))) - Val(.TextMatrix(lngRow, .ColIndex("�ɹ����"))), mFMT.FM_���)
           End If
        Else
                .TextMatrix(lngRow, .ColIndex("���ۼ�")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ۼ�"))) / dbl����ϵ��, mFMT.FM_ɢװ���ۼ�)
                .TextMatrix(lngRow, .ColIndex("���۽��")) = Format(Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) * (dbl���� * dbl����ϵ��), mFMT.FM_���)
                '���۲��=���۽��-������
                .TextMatrix(lngRow, .ColIndex("���۲��")) = Format(Val(.TextMatrix(lngRow, .ColIndex("���۽��"))) - Val(.TextMatrix(lngRow, .ColIndex("�ɹ����"))), mFMT.FM_���)
        End If
    End With
    �������ۼۼ����۲�� = True
End Function

Private Sub AfterDeleteRow()
    'ɾ���к�
    
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mshBill.ColIndex("�к�"), mshBill.Row)
End Sub
Private Function Select������Ϣ(Optional strSearch As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������Ϣѡ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-12-02 11:50:35
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, lng�ⷿid As Long
    Dim sngLeft As Single, sngTop As Single
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    With mshBill
        lng�ⷿid = cboStock.ItemData(cboStock.ListIndex)
        If strSearch = "" Then
            Set rsTemp = Frm����ѡ����.ShowMe(Me, 1, , lng�ⷿid, lng�ⷿid, , , , , , , , , , , 1714, , mstrPrivs, , False)
        Else
            Call CalcPosition(sngLeft, sngTop, mshBill)
            Set rsTemp = FrmMulitSel.ShowSelect(Me, 1, lng�ⷿid, lng�ⷿid, lng�ⷿid, strSearch, sngLeft, sngTop, mshBill.CellWidth, mshBill.CellHeight, , , , , , , , , , 1714, , mstrPrivs, , False)
        End If
        
        If rsTemp.RecordCount <= 0 Then Exit Function
        rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            SetColValue .Row, rsTemp!����ID, _
                "[" & rsTemp!���� & "]" & rsTemp!����, IIf(IsNull(rsTemp!���), "", rsTemp!���), _
                IIf(IsNull(rsTemp!����), "", rsTemp!����), IIf(mintUnit = 0, rsTemp!ɢװ��λ, rsTemp!��װ��λ), _
                IIf(IsNull(rsTemp!�ۼ�), 0, rsTemp!�ۼ�), rsTemp!ָ�������� / IIf(mintUnit = 0, 1, rsTemp!����ϵ��), _
                IIf(IsNull(rsTemp!����), "!", rsTemp!����), rsTemp!���Ч��, IIf(mintUnit = 0, 1, rsTemp!����ϵ��), _
                rsTemp!ʱ��, rsTemp!���÷���, rsTemp!ָ������� / 100, IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
            
            
            If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
            .Row = .Row + 1
            
            rsTemp.MoveNext
        Next
        
        mshBill.Row = int�����
        
'        If rsTemp.RecordCount = 1 Then
'            SetColValue .Row, rsTemp!����ID, _
'                "[" & rsTemp!���� & "]" & rsTemp!����, IIf(IsNull(rsTemp!���), "", rsTemp!���), _
'                IIf(IsNull(rsTemp!����), "", rsTemp!����), IIf(mintUnit = 0, rsTemp!ɢװ��λ, rsTemp!��װ��λ), _
'                IIf(IsNull(rsTemp!�ۼ�), 0, rsTemp!�ۼ�), rsTemp!ָ�������� / IIf(mintUnit = 0, 1, rsTemp!����ϵ��), _
'                IIf(IsNull(rsTemp!����), "!", rsTemp!����), rsTemp!���Ч��, IIf(mintUnit = 0, 1, rsTemp!����ϵ��), _
'                rsTemp!ʱ��, rsTemp!���÷���, rsTemp!ָ������� / 100, IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
'        End If
        rsTemp.Close
        Select������Ϣ = True
    End With
    Call ��ʾ�����
End Function
 
Private Function SelDate(ByVal intCol As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-07 11:59:54
    '-----------------------------------------------------------------------------------------------------------
    Dim strDate As String, blnreturn As Boolean
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim strMaxDate As String
    Dim lngRow As Long
    With mshBill
        strDate = .TextMatrix(.Row, intCol)
        If strDate = "" Then strDate = Format(sys.Currentdate, "yyyy-mm-dd")
        lngH = .CellHeight
        Call CalcPosition(sngX, sngY, mshBill)
        If intCol = .ColIndex("Ч��") Then strMaxDate = "3000-01-01"
        If intCol = .ColIndex("���ʧЧ��") Then strMaxDate = "3000-01-01"
        If intCol = .ColIndex("��������") Then strMaxDate = Format(sys.Currentdate, "yyyy-mm-dd")
    End With
    blnreturn = frmDateSel.SelectDate(Me, sngX, sngY, lngH, strDate, , strMaxDate)
    If blnreturn = False Then Exit Function
    With mshBill
        .TextMatrix(.Row, intCol) = strDate
    End With
    zlVsMoveGridCell mshBill, mshBill.ColIndex("������Ϣ"), 0, True, lngRow
    SelDate = True
End Function

Private Function Show�ӳ���(ByVal intCol As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʾ�ӳ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-07 11:59:54
    '-----------------------------------------------------------------------------------------------------------
    Dim blnreturn As Boolean, dbl���ۼ� As Double, dbl����� As Double, dbl�ӳ��� As Double, lng����ID As Long
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim dblԭ�ӳ��� As Double
    
    dbl�ӳ��� = 15
    With mshBill
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        If lng����ID = 0 Then Exit Function
        
        dbl���ۼ� = Val(.TextMatrix(.Row, .ColIndex("�ۼ�"))) '
        If intCol = .ColIndex("�ɹ���") Then
            dbl����� = Val(.EditText)
        Else
            dbl����� = Val(.TextMatrix(.Row, .ColIndex("�ɹ���")))
        End If
        If dbl���ۼ� <> 0 And dbl����� <> 0 Then
            dbl�ӳ��� = Format(����ӳ���(lng����ID, dbl���ۼ�, dbl�����), "####0.0000000;-###0.0000000;0;0")
        End If
        lngH = .CellHeight
        Call CalcPosition(sngX, sngY, mshBill)
    End With
    dblԭ�ӳ��� = dbl�ӳ���
    
    blnreturn = frm����Set.ShowCalc(Me, sngX, sngY, lngH, lng����ID, mintUnit, dbl���ۼ�, dbl�����, dbl�ӳ���, mbln��ǿ�ƿ���ָ���۸�)
    With mshBill
        If blnreturn = False Then
            mdbl�Ӽ��� = dblԭ�ӳ���
            '���¼������ۼۡ����
            .TextMatrix(.Row, .ColIndex("�ۼ�")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ɹ���"))) * (1 + (mdbl�Ӽ��� / 100)), mFMT.FM_���ۼ�)
            .TextMatrix(.Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ۼ�"))) * Val(.TextMatrix(.Row, .ColIndex("����"))), mFMT.FM_���)
            .TextMatrix(.Row, .ColIndex("���")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(.Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
            Exit Function
        End If
        .TextMatrix(.Row, .ColIndex("�ۼ�")) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
        .TextMatrix(.Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ۼ�"))) * Val(.TextMatrix(.Row, .ColIndex("����"))), mFMT.FM_���)
        .TextMatrix(.Row, .ColIndex("���")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(.Row, .ColIndex("�ɹ����"))), mFMT.FM_���)
        mdbl�Ӽ��� = dbl�ӳ���
    End With
    Show�ӳ��� = True
    'debug.Print "aaa"
End Function

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.Rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.Rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(mshBill.TextMatrix(n, mshBill.ColIndex("���"))) = 0, n, Val(mshBill.TextMatrix(n, mshBill.ColIndex("���"))))
                !ҩƷid = Val(mshBill.TextMatrix(n, mshBill.ColIndex("����id")))
                
                .Update
            End If
        Next
        
    End With
End Sub
