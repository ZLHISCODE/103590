VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSelfMakeCard 
   AutoRedraw      =   -1  'True
   Caption         =   "����������ⵥ"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmSelfMakeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   5970
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   10
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   7
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   8
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   11655
      TabIndex        =   12
      Top             =   0
      Width           =   11715
      Begin VSFlex8Ctl.VSFlexGrid vs��ɲ��� 
         Height          =   2220
         Left            =   210
         TabIndex        =   30
         Top             =   2610
         Width           =   11145
         _cx             =   19659
         _cy             =   3916
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSelfMakeCard.frx":014A
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
         ExplorerBar     =   5
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
      Begin VB.TextBox txtNo 
         Height          =   300
         Left            =   9945
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   29
         Top             =   180
         Width           =   1410
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDrug 
         Height          =   2235
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   3942
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   1230
         Left            =   195
         TabIndex        =   4
         Top             =   945
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2170
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   6
         Top             =   4920
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
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   26
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:asdfasdfasdfsadfsadfsdfasdfsadfasdfsdf"
         Height          =   180
         Left            =   2040
         TabIndex        =   25
         Top             =   2280
         Width           =   4590
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7230
         TabIndex        =   22
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   21
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   20
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   19
         Top             =   5280
         Width           =   915
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
         TabIndex        =   18
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
         Top             =   4995
         Width           =   645
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "��������������ⵥ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   5340
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   15
         Top             =   5340
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   6645
         TabIndex        =   14
         Top             =   5340
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   8520
         TabIndex        =   13
         Top             =   5340
         Width           =   720
      End
      Begin VB.Label LblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƽ���(&T)"
         Height          =   180
         Left            =   8220
         TabIndex        =   2
         Top             =   660
         Width           =   810
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
            Picture         =   "frmSelfMakeCard.frx":02AA
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":04C4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":06DE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":08F8
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0B12
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0D2C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0F46
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1160
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
            Picture         =   "frmSelfMakeCard.frx":137A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1594
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":17AE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":19C8
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1BE2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1DFC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":2016
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":2230
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
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
            Picture         =   "frmSelfMakeCard.frx":244A
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
            Picture         =   "frmSelfMakeCard.frx":2CDE
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSelfMakeCard.frx":31E0
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
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmSelfMakeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln�������� As Boolean
Private mblnFirst As Boolean
Private mintUnit  As Integer                '0-ɢװ��λ,1-��װ��λ
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mint����� As Integer             '��ʾ���ĳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Dim mstrPrivs As String                     'Ȩ��
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
'���˺�:2007/06/10:����10813
Private mstrTime_Start As String            '���뵥�ݱ༭�ĵ���ʱ�� ,��Ҫ�ж��Ƿ񵥾ݱ����˸��Ĺ�,����༭��,���ܽ������
Private mstrTime_End As String
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴
Private Const mstrCaption As String = "����������ⵥ"
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
Private Const mlngModule = 1713

'----------------------------------------------------------------------------------------------------------

Private mcolUseCount As Collection

'=========================================================================================

Private Const mconIntCol���� As Integer = 1
Private Const mconIntCol��� As Integer = 2
Private Const mconIntColԭ���� As Integer = 3
Private Const mconIntCol����ϵ�� As Integer = 4
Private Const mconIntCol��λ As Integer = 5
Private Const mconIntCol���� As Integer = 6
Private Const mconIntColЧ�� As Integer = 7
Private Const mconIntColһ���Բ��� As Integer = 8
Private Const mconIntCol���Ч�� As Integer = 9
Private Const mconIntCol�������   As Integer = 10
Private Const mconIntCol���ʧЧ�� As Integer = 11

Private Const mconIntCol���� As Integer = 12
Private Const mconIntCol�ɹ��� As Integer = 13
Private Const mconIntCol�ɹ���� As Integer = 14
Private Const mconIntCol�ۼ� As Integer = 15
Private Const mconIntCol�ۼ۽�� As Integer = 16
Private Const mconintCol��� As Integer = 17


Private Const mconIntColS As Integer = 18       '������
'=========================================================================================
'�������������
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    Dim strSQL As String
    GetDepend = False
    
    On Error GoTo ErrHandle
    strSQL = "" & _
        "   SELECT B.Id,b.���� " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID " & _
        "       AND A.���� = 31  and b.ϵ��=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption & "-��������")
    If rsTemp.EOF Then
        MsgBox "û����������������������������������������ã�", vbInformation + vbOKOnly, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    strSQL = "" & _
        "   SELECT B.Id,b.���� " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID " & _
        "           AND A.���� = 31  and b.ϵ��=-1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption & "-�������")
    If rsTemp.EOF Then
        MsgBox "û�����������������ĳ����������������������ã�", vbInformation + vbOKOnly, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    strSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� " & _
        "           AND b.���� ='K'" & _
        "           AND a.id = c.����id " & _
        "           AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption & "-����Ƽ���")
    If rsTemp.EOF Then
        MsgBox "����������û������Ϊ�Ƽ��ҵĲ���,��鿴���Ź���", vbInformation, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    strSQL = " SELECT a.���Ʋ���id FROM ���Ʋ��Ϲ��� a, �������� b Where a.���Ʋ���id = b.����id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, mstrCaption & "-��ȡ�������Ĺ���")
    If rsTemp.EOF Then
        MsgBox "û��һ�־���ԭ��������ɵ���������,��鿴��������Ŀ¼����", vbInformation, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
        Optional int��¼״̬ As Integer = 1, Optional strPrivs As String, _
        Optional blnSuccess As Boolean = False)
        
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
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain

    Call GetRegInFor(g˽��ģ��, "��������������", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)
    
    If mint�༭״̬ = 1 Then
        mblnEdit = True
        txtNo = mstr���ݺ�
        txtNo.Tag = txtNo
        txtNo.Locked = True
        txtNo.TabStop = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
        txtNo.Locked = True
        txtNo.TabStop = True
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
    End If
    
    If Not GetDepend Then Exit Sub
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub
Private Sub cboStock_Click()
    mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
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
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���Ӧ���ĵĵ�λ��" & vbCrLf & "��Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '�������ĵ�λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                            
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
        .Col = mconIntCol����
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRownew mshBill, mconIntCol����, txtCode.Text, True
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
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            '�����ѱ�ɾ��
            MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
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
        FindRownew mshBill, mconIntCol����, txtCode.Text, False
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
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
        
    If mint�༭״̬ = 3 Then        '���
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub

        '���˺�:2007/06/10:����10813
        mstrTime_End = GetBillInfo(16, txtNo.Tag)
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
    
    If txtNo.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
    
    mblnSave = False
'    mblnEdit = True
    mshBill.ClearBill
    vs��ɲ���.Clear (1)
    vs��ɲ���.Rows = 2
    Call ��ʾ�ϼƽ��
'    SetEdit
    txtժҪ.Text = ""
    cboType.SetFocus
    mblnChange = False
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    mblnFirst = True
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    mintUnit = Val(strReg)
    
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
        
    
    mintBatchNoLen = GetBatchNoLen()
    
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo.Text
    With cboType
    
        gstrSQL = "" & _
            "   SELECT DISTINCT a.id, a.���� " & _
            "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
            "   Where c.�������� = b.���� " & _
            "           AND b.���� ='K'" & _
            "           AND a.id = c.����id " & _
            "           AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption)
        
        If rsTemp.EOF Then Exit Sub
        
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp.Fields(1)
            .ItemData(.NewIndex) = rsTemp.Fields(0)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    Call initCard
    
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(mconIntCol�ɹ���) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mconIntCol�ɹ����) = IIf(mblnCostView = True, 1200, 0)
        .ColWidth(mconintCol���) = IIf(mblnCostView = True, 1200, 0)
    End With
    With vs��ɲ���
        .ColWidth(.ColIndex("�ɱ���")) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(.ColIndex("�ɱ����")) = IIf(mblnCostView = True, 1200, 0)
        .ColWidth(.ColIndex("���")) = IIf(mblnCostView = True, 1200, 0)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim str����ϵ�� As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim strReg As String
    
    On Error GoTo ErrHandle
    strReg = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    
    strOrder = strReg
    
    '�ⷿ
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
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
                
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "select b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id and A.���� = 16 and a.no=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
                
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsTemp!����
                    .ItemData(.NewIndex) = rsTemp!Id
                    .ListIndex = 0
                End With
                rsTemp.Close
            End If
            
            
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "c.���㵥λ AS ��λ,(A.��д����) AS ����,1 as ����ϵ��,"
                    str����ϵ�� = "1"
                Case Else
                    strUnitQuantity = "B.��װ��λ AS ��λ,(A.��д���� / B.����ϵ��) AS ����,B.����ϵ�� as ����ϵ��, "
                    str����ϵ�� = "B.����ϵ��"
            End Select
            
            gstrSQL = "" & _
                "   SELECT * " & _
                "   FROM (  SELECT DISTINCT ���,a.ҩƷid as ����id, ('[' || c.���� || ']' || c.����) AS ������Ϣ,c.���,a.����, a.����, a.Ч��," & _
                "                   zlSpellCode(c.����) ����," & strUnitQuantity & _
                "                   (a.�ɱ���*" & str����ϵ�� & ") AS �ɱ���," & _
                "                   a.�ɱ���� ,(a.���ۼ�*" & str����ϵ�� & ") AS ���ۼ�,a.���۽�� AS ���۽��," & _
                "                   a.��� AS ���,a.������,a.��������,a.�����,a.�������,a.ժҪ,c.���� as ԭ����,b.���Ч��,b.һ���Բ���,b.���Ч��," & _
                "                   a.�������,a.���Ч�� as ���ʧЧ��,a.�Է�����id,c.�Ƿ���,b.ָ�������/100 as ָ�������,b.���÷��� " & _
                "           FROM ҩƷ�շ���¼ a, �������� b,�շ���ĿĿ¼ c" & _
                "           Where a.ҩƷid = b.����id and a.ҩƷid=c.id " & _
                "                   AND a.��¼״̬ = [2]" & _
                "                   AND a.���� = 16 AND ���ϵ��=1 " & _
                "                   AND a.no =[1]  " & _
                "           )" & _
                " ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, mint��¼״̬)
                
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            '���˺�:2007/06/10:����10813
            mstrTime_Start = GetBillInfo(16, mstr���ݺ�)
            
            Txt������ = rsTemp!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û���
            End If
            
            Txt�������� = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
            
            Txt����� = IIf(IsNull(rsTemp!�����), "", rsTemp!�����)
            Txt������� = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd hh:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
            
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsTemp!�Է�����id Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
            End With
            
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = rsTemp.AbsolutePosition
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp!����ID
                    .TextMatrix(intRow, mconIntCol����) = rsTemp!������Ϣ
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)

                    .TextMatrix(intRow, mconIntCol��λ) = rsTemp!��λ
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntColһ���Բ���) = zlStr.Nvl(rsTemp!һ���Բ���)
                    .TextMatrix(intRow, mconIntCol���Ч��) = zlStr.Nvl(rsTemp!���Ч��)
                    .TextMatrix(intRow, mconIntCol�������) = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol���ʧЧ��) = IIf(IsNull(rsTemp!���ʧЧ��), "", Format(rsTemp!���ʧЧ��, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol����) = Format(zlStr.Nvl(rsTemp!����, 0), mFMT.FM_����)
                    .TextMatrix(intRow, mconIntCol�ɹ���) = Format(zlStr.Nvl(rsTemp!�ɱ���, 0), mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mconIntCol�ɹ����) = Format(zlStr.Nvl(rsTemp!�ɱ����, 0), mFMT.FM_���)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = Format(zlStr.Nvl(rsTemp!���ۼ�, 0), mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = Format(zlStr.Nvl(rsTemp!���۽��, 0), mFMT.FM_���)
                    .TextMatrix(intRow, mconintCol���) = Format(zlStr.Nvl(rsTemp!���, 0), mFMT.FM_���)
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsTemp!���Ч��), "0", rsTemp!���Ч��) & "||" & rsTemp!ָ������� & "||" & rsTemp!�Ƿ��� & "||" & rsTemp!���÷���
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsTemp!����ϵ��
                    rsTemp.MoveNext
                Loop
                
                Dim dblCostPrice As Double
                If .TextMatrix(1, 0) <> "" Then
                    Call Set��ɲ���(Val(.TextMatrix(1, 0)), Val(.TextMatrix(1, mconIntCol����) * .TextMatrix(1, mconIntCol����ϵ��)), False, dblCostPrice)
                End If
                
            End With
            rsTemp.Close
                 
    End Select
    SetEdit         '���ñ༭����
    Call ��ʾ�ϼƽ��
    
    If mint�༭״̬ = 2 And mint����� <> 0 Then
        SetUseCountCol
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'�����޸�ǰԭ��ҩ��ʹ���������Ա������޸Ĺ����жԿ���������жϸ�׼ȷ
Private Sub SetUseCountCol()
    Dim rsTemp As New Recordset
    Dim numUsedCount As Double
    Dim vardrug As Variant
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select ҩƷid as ����id,��д����,����id,���� " & _
        "   From ҩƷ�շ���¼ " & _
        "   Where no=[1] and ����=16 and ��¼״̬=1 and ���ϵ��=-1 "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
    
    If rsTemp.EOF Then Exit Sub
    
    Set mcolUseCount = New Collection
    With mcolUseCount
        Do While Not rsTemp.EOF
            numUsedCount = 0
            For Each vardrug In mcolUseCount
                If vardrug(0) = zlStr.Nvl(rsTemp!����ID) & "!" & zlStr.Nvl(rsTemp!����ID) & "!" & Val(zlStr.Nvl(rsTemp!����)) Then
                    numUsedCount = vardrug(1)
                    .Remove vardrug(0)
                    Exit For
                End If
            Next
            .Add Array(zlStr.Nvl(rsTemp!����ID) & "!" & zlStr.Nvl(rsTemp!����ID) & "!" & Val(zlStr.Nvl(rsTemp!����)), Val(rsTemp!��д����)), zlStr.Nvl(rsTemp!����ID) & "!" & zlStr.Nvl(rsTemp!����ID) & "!" & Val(zlStr.Nvl(rsTemp!����))
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            cboStock.Enabled = False
            cboType.Enabled = False
            txtժҪ.Enabled = False
        Else
            .ColData(0) = 5
            .ColData(mconIntCol����) = 1
            .ColData(mconIntCol���) = 5
            
            .ColData(mconIntCol��λ) = 5
            .ColData(mconIntCol����) = 4
            .ColData(mconIntColЧ��) = 5
            .ColData(mconIntCol�������) = 2
            .ColData(mconIntCol���Ч��) = 5
            .ColData(mconIntCol����) = 4
            .ColData(mconIntCol�ɹ���) = 5
            .ColData(mconIntCol�ɹ����) = 5
            .ColData(mconIntCol�ۼ�) = 5
            .ColData(mconIntCol�ۼ۽��) = 5
            .ColData(mconintCol���) = 5
            
            
            .ColData(mconIntColԭ����) = 5
            .ColData(mconIntCol����ϵ��) = 5
            
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntCol���) = flexAlignLeftCenter
            
            .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol����) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
            .ColAlignment(mconintCol���) = flexAlignRightCenter
            
            cboStock.Enabled = True
            cboType.Enabled = True
            txtժҪ.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol����) = "���������"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = "ʧЧ��"

        .TextMatrix(0, mconIntColһ���Բ���) = "һ���Բ���"
        .TextMatrix(0, mconIntCol���Ч��) = "���Ч��"
        .TextMatrix(0, mconIntCol�������) = "�������"
        .TextMatrix(0, mconIntCol���ʧЧ��) = "���ʧЧ��"
                
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        
        .TextMatrix(0, mconIntColԭ����) = "ԭЧ��"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        
        
        .TextMatrix(1, 0) = ""
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol����) = 2000
        .ColWidth(mconIntCol���) = 900
        
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        
        
        .ColWidth(mconIntColһ���Բ���) = 0
        .ColWidth(mconIntCol���Ч��) = 0
        .ColWidth(mconIntCol�������) = 1000
        .ColWidth(mconIntCol���ʧЧ��) = 1000
                
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol�ɹ���) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mconIntCol�ɹ����) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mconIntCol�ۼ�) = 900
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconintCol���) = IIf(mblnCostView = False, 0, 800)
        
        
        
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol����) = 1
        .ColData(mconIntCol���) = 5
        
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 4
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntColһ���Բ���) = 5
        .ColData(mconIntCol���Ч��) = 5
        .ColData(mconIntCol�������) = 2
        .ColData(mconIntCol���ʧЧ��) = 5
        
        .ColData(mconIntCol����) = 4
        .ColData(mconIntCol�ɹ���) = 5
        .ColData(mconIntCol�ɹ����) = 5
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 0
        
        
        .ColData(mconIntColԭ����) = 5
        .ColData(mconIntCol����ϵ��) = 5
        
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntColһ���Բ���) = flexAlignCenterCenter
        .ColAlignment(mconIntCol���Ч��) = flexAlignCenterCenter
        .ColAlignment(mconIntCol�������) = flexAlignCenterCenter
        .ColAlignment(mconIntCol���ʧЧ��) = flexAlignCenterCenter
        
             
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol����
        .LocateCol = mconIntCol����
    End With
    
    With vs��ɲ���
        .RowHeight(0) = .RowHeight(0) * 2
        .ExplorerBar = .ExplorerBar + &H1000&
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
    With txtNo
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
    
    With vs��ɲ���
        .Left = mshBill.Left
        .Width = mshBill.Width
        .Top = txtժҪ.Top - 60 - .Height
    End With
            
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = vs��ɲ���.Top - 60 - .Height
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
    
    If mshDrug.Visible Then
        mshDrug.Visible = False
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    
End Sub

Private Function SaveCheck() As Boolean
    mblnSave = False
    SaveCheck = False
    gstrSQL = "zl_���Ʋ������_verify('" & txtNo.Tag & "','" & UserInfo.�û��� & "')"
    On Error GoTo ErrHandle
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub mshBill_AfterDeleteRow()
    With mshBill
        If .Row > 1 Then
            .Row = .Row - 1
        Else
            .Row = 1
        End If
        If .TextMatrix(.Row, 0) = "" Then
            vs��ɲ���.Clear (1)
        Else
            Dim dblCostPrice As Double
            Call Set��ɲ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol����ϵ��)), False, dblCostPrice)
        End If
        
    End With
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ���������ģ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim rsTemp As New Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim intStockID As Long
    Dim strUnitQuantity As String
    
    On Error GoTo ErrHandle
    Select Case mintUnit
        Case 0
            strUnitQuantity = "D.���㵥λ AS ��λ, (to_char(s.������� ," & mOraFMT.FM_���� & ")) AS ����,1 as ����ϵ��," _
                & "to_char(p.�ۼ�," & mOraFMT.FM_���ۼ� & ") as �ۼ�,"
        Case Else
            strUnitQuantity = "d.��װ��λ AS ��λ, (to_char(s.������� / d.����ϵ��," & mOraFMT.FM_���� & ")) AS ����,d.����ϵ�� as ����ϵ��," _
                & "to_char(p.�ۼ�*d.����ϵ��," & mOraFMT.FM_���ۼ� & ") as �ۼ�, "
    End Select
        
    intStockID = cboStock.ItemData(cboStock.ListIndex)
    
    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50
    
    
    '��������
    gstrSQL = "" & _
        "   SELECT  D.����, D.����,D.���, d.����,d.����id as ҩƷid, " & strUnitQuantity & "  s.�����, d.���Ч��,d.�Ƿ���,d.ָ�������/100 as ָ�������,d.���÷���,e.�ⷿ��λ " & _
        "   FROM  ( SELECT DISTINCT L.����, L.����,L.���, L.����, d.����id, L.���㵥λ,NVL (TO_CHAR (d.���Ч��, '9999990'), 0) ���Ч��," & _
        "                   d.��װ��λ,TO_CHAR (d.����ϵ��, " & GFM_XS & ") ����ϵ��,l.�Ƿ���,d.ָ�������,d.���÷��� " & _
        "           FROM ���Ʋ��Ϲ��� f,  �������� d,�շ���ĿĿ¼ L,�շ�ִ�п��� R" & _
        "           Where f.���Ʋ���id = d.����id  and F.���Ʋ���id=L.id And (L.վ��=[2] or L.վ�� is null) AND nvl(d.���Ʋ���,0)=1 and f.���Ʋ���ID=R.�շ�ϸĿid and R.ִ�п���ID=[1]" & _
        "                   AND (   EXISTS (SELECT 1 From ��������˵�� WHERE  �������� In ('���Ŀ�','�Ƽ���', '����ⷿ') AND ����id = [1]) " & _
        "                           OR L.������� =(SELECT distinct '1' From ��������˵�� WHERE �������� LIKE '���ϲ���' AND ����id =[1] AND ������� IN (1, 3)) " & _
        "                           OR L.������� =(SELECT distinct '2' From ��������˵�� WHERE �������� LIKE '���ϲ���' AND ����id =[1] AND ������� IN (2, 3))) " & _
        "                   AND (L.����ʱ�� IS NULL OR TO_CHAR (L.����ʱ��, 'yyyy-MM-dd') = '3000-01-01') " & _
        "           ) d,"
        
    '���շѼ�Ŀ����Ҫ���ۼ�
    gstrSQL = gstrSQL & _
        "   (   SELECT �շ�ϸĿid, TO_CHAR (�ּ�, '999999999990.9999') �ۼ� " & _
        "       From �շѼ�Ŀ " & _
        "       WHERE ((SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (    SYSDATE >= ִ������ AND ��ֹ���� IS NULL)) " & _
        GetPriceClassString("") & _
        "   ) p,"
             
    
    '��ҩƷ���
    gstrSQL = gstrSQL & _
        "   (   SELECT ҩƷid, TO_CHAR (SUM (��������), " & mOraFMT.FM_���� & ") ��������,TO_CHAR (SUM (ʵ������)," & mOraFMT.FM_���� & ") �������,TO_CHAR (SUM (ʵ�ʽ��), " & mOraFMT.FM_��� & ") ����� " & _
        "       From ҩƷ��� " & _
        "       Where �ⷿid =[1]  and ����=1 " & _
        "       GROUP BY ҩƷid) s, "
    
    gstrSQL = gstrSQL & _
        "   (   Select ����ID,�ⷿID,�ⷿ��λ From ���ϴ����޶� " & _
        "       Where �ⷿID=[1]) E ,�շ���ĿĿ¼ M"
        
    '��������
    gstrSQL = gstrSQL & _
        "   Where d.����id = p.�շ�ϸĿid  And D.����id=M.id And (M.վ��=[2] or M.վ�� is null) And M.�Ƿ���<>1 AND d.����id = s.ҩƷid (+) and D.����ID=E.����ID(+)" & _
        "   ORDER BY d.����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, intStockID, gstrNodeNo)
           
    If rsTemp.EOF Then
        ShowMsgBox "���������Ʋ���,����δ���ô洢�ⷿ��δ�������Ʋ���,����[����Ŀ¼����]������!"
        Exit Sub
    End If
    
    Set mshDrug.Recordset = rsTemp
    rsTemp.Close
    Call SetDrugWidth(sngLeft, sngTop)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��������ѡ�����Ŀ�ȼ��������
Private Sub SetDrugWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    With mshDrug
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
        If RestoreFlexState(mshDrug, mstrCaption) = False Then
            .ColWidth(0) = 1000
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 1000
            
            .ColWidth(4) = 1000
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 0
            
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 0
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(.Cols - 1) = 1500
        End If
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
        .ColAlignment(12) = flexAlignRightCenter
        
        .SetFocus
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol���� Or .Col = mconIntCol�ɹ��� Or .Col = mconIntCol�ۼ� Or .Col = mconIntCol�ɹ���� Or .Col = mconIntCol�ۼ۽�� Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconIntCol����
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.����С��, g_С��λ��.obj_ɢװС��.����С��)
                Case mconIntCol�ɹ���, mconIntCol�ۼ�
                   intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.�ɱ���С��, g_С��λ��.obj_ɢװС��.�ɱ���С��)
                Case mconIntCol�ɹ����, mconIntCol�ۼ۽��
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.���С��, g_С��λ��.obj_ɢװС��.���С��)
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                KeyAscii = 0
                Exit Sub
            End If
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            SetInputFormat .Row
            Dim dblCostPrice As Double
            
            If .TextMatrix(.Row, 0) <> "" Then
                Call Set��ɲ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol����ϵ��)), False, dblCostPrice)
            Else
                vs��ɲ���.Rows = 2
                vs��ɲ���.Clear (1)
            End If
                
        End If
        
        Select Case .Col
            Case mconIntCol����
                .TxtCheck = False
                .MaxLength = 80
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                            
            Case mconIntCol����
                .TxtCheck = False
                '.TextMask = "1234567890"
                .MaxLength = mintBatchNoLen
            
            Case mconIntColЧ��
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol����)) And .TextMatrix(.Row, mconIntColԭ����) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0) <> "0" Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                                 Call CheckLapse(.TextMatrix(.Row, mconIntColЧ��))
                            End If
                        End If
                    End If
                End If
            Case mconIntCol�ɹ���
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol�ɹ����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim strSerach As String
    Dim strLike As String
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        With mshBill
            .Text = Trim(.Text)
            strKey = Trim(.Text)
            
            If Mid(strKey, 1, 1) = "[" Then
                If InStr(2, strKey, "]") <> 0 Then
                    strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
                Else
                    strKey = Mid(strKey, 2)
                End If
            End If
            Select Case .Col
                
                Case mconIntCol����
                    If strKey <> "" Then
                        Dim rsTemp As New Recordset
                        Dim sngLeft As Single
                        Dim sngTop As Single
                        Dim intStockID As Long
                        
                        Select Case mintUnit
                            Case 0
                                strUnitQuantity = "d.���㵥λ AS ��λ, (to_char(s.������� ," & mOraFMT.FM_���� & ")) AS ����,1 as ����ϵ��," & _
                                    "   to_char(p.�ۼ�," & mOraFMT.FM_���ۼ� & ") as �ۼ�,"
                            Case 1
                                strUnitQuantity = "d.��װ��λ AS ��λ, (to_char(s.������� / d.����ϵ��," & mOraFMT.FM_���� & ")) AS ����,d.����ϵ�� as ����ϵ��," _
                                    & "to_char(p.�ۼ�*d.����ϵ��," & mOraFMT.FM_���ۼ� & ") as �ۼ�, "
                        End Select
                            
                        intStockID = cboStock.ItemData(cboStock.ListIndex)
                        
                        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50
                        
                        strSerach = " And (A.���� Like [2] OR B.���� Like [2] OR B.���� LIKE [2])"
                        
                        If IsNumeric(strKey) Then                         '���������,��ֻȡ����
                            If Mid(gSystem_Para.Para_���뷽ʽ, 1, 1) = "1" Then strSerach = " And (A.���� Like [2])"
                            strLike = "" & GetMatchingSting(UCase(strKey)) & ""
                        ElseIf zlStr.IsCharAlpha(strKey) Then          '����ȫ����ĸʱֻƥ�����
                            If Mid(gSystem_Para.Para_���뷽ʽ, 2, 1) = "1" Then strSerach = " And B.���� Like [2] "
                            strLike = "" & GetMatchingSting(UCase(strKey)) & ""
                        ElseIf zlStr.IsCharChinese(strKey) Then
                            strSerach = " And B.���� Like [2] "
                            strLike = "" & GetMatchingSting(strKey) & ""
                        End If
                        
                        
                        '��������
                          gstrSQL = "" & _
                              "   SELECT  d.����, d.����,d.���, d.����,d.����id , " & strUnitQuantity & "  s.�����, d.���Ч��,d.�Ƿ���,d.ָ�������/100 as ָ�������,d.���÷���,e.�ⷿ��λ " & _
                              "   FROM  ( SELECT DISTINCT l.����, l.����,l.���, L.����, d.����id, l.���㵥λ,NVL (TO_CHAR (d.���Ч��, '9999990'), 0) ���Ч��," & _
                              "                   d.��װ��λ,TO_CHAR (d.����ϵ��, " & GFM_XS & ") ����ϵ��,l.�Ƿ���,d.ָ�������,d.���÷��� " & _
                              "           FROM ���Ʋ��Ϲ��� f,  �������� d, " & _
                              "                    (  Select A.ID,A.����,A.����,A.���,A.����,A.���㵥λ,A.�������,a.�Ƿ���  " & _
                              "                       From �շ���ĿĿ¼ A,�շ���Ŀ���� B  " & _
                              "                       Where A.ID=B.�շ�ϸĿID And (A.վ��=[4] or A.վ�� is null) " & _
                              "                           And B.����=[5] " & _
                              "                           AND A.��� ='4' And (A.����ʱ�� is null Or A.����ʱ��>=[3]) " & strSerach & ")  L," & _
                              "                 �շ�ִ�п��� R" & _
                              "           Where f.���Ʋ���id = d.����id and f.���Ʋ���id=l.id AND nvl(d.���Ʋ���,0)=1 and F.���Ʋ���ID=R.�շ�ϸĿID and R.ִ�п���ID=[1]" & _
                              "                   AND (   EXISTS (SELECT 1 From ��������˵�� WHERE  �������� In ('���Ŀ�','�Ƽ���', '����ⷿ') AND ����id = [1]) " & _
                              "                           OR L.������� =(SELECT distinct '1' From ��������˵�� WHERE �������� LIKE '���ϲ���' AND ����id =[1] AND ������� IN (1, 3)) " & _
                              "                           OR l.������� =(SELECT distinct '2' From ��������˵�� WHERE �������� LIKE '���ϲ���' AND ����id =[1] AND ������� IN (2, 3))) " & _
                              "                     " & _
                              "           ) d,"
                              
                          
                          '���շѼ�Ŀ����Ҫ���ۼ�
                          gstrSQL = gstrSQL & _
                              "   (   SELECT �շ�ϸĿid, TO_CHAR (�ּ�, '999999999990.9999') �ۼ� " & _
                              "       From �շѼ�Ŀ " & _
                              "       WHERE ((SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (    SYSDATE >= ִ������ AND ��ֹ���� IS NULL)) " & _
                              GetPriceClassString("") & _
                              "   ) p,"
                                   
                          
                          '��ҩƷ���
                          gstrSQL = gstrSQL & _
                              "   (   SELECT ҩƷid, TO_CHAR (SUM (��������), " & mOraFMT.FM_���� & ") ��������,TO_CHAR (SUM (ʵ������)," & mOraFMT.FM_���� & ") �������,TO_CHAR (SUM (ʵ�ʽ��), " & mOraFMT.FM_��� & ") ����� " & _
                              "       From ҩƷ��� " & _
                              "       Where �ⷿid =[1]  and ����=1 " & _
                              "       GROUP BY ҩƷid) s, "
                          
                          gstrSQL = gstrSQL & _
                              "   (   Select ����ID,�ⷿID,�ⷿ��λ From ���ϴ����޶� " & _
                              "       Where �ⷿID=[1]) E,�շ���ĿĿ¼ M"
                              
                          '��������
                          gstrSQL = gstrSQL & _
                              "   Where d.����id = p.�շ�ϸĿid  And D.����id=M.id And (M.վ��=[4] or M.վ�� is null) And M.�Ƿ���<>1 AND d.����id  = s.ҩƷid (+) and D.����id =E.����ID(+)" & _
                              "   ORDER BY d.����"
                              '(Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1])

                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, intStockID, strLike, CDate("3000-01-01"), gstrNodeNo, IIf(gSystem_Para.int���뷽ʽ = 1, 2, 1))
                                                  
                        
                        If rsTemp.EOF Then
                            MsgBox "û��ƥ����������ģ�", vbInformation + vbOKOnly, gstrSysName
                            rsTemp.Close
                            Cancel = True
                            Exit Sub
                        ElseIf rsTemp.RecordCount = 1 Then
                            If SetColValue(.Row, rsTemp!����ID, "[" & rsTemp!���� & "]" & rsTemp!����, IIf(IsNull(rsTemp!���), "", rsTemp!���), _
                               rsTemp!��λ, _
                               IIf(IsNull(rsTemp!�ۼ�), 0, rsTemp!�ۼ�), _
                               IIf(IsNull(rsTemp!���Ч��), "0", rsTemp!���Ч��), rsTemp!����ϵ��, rsTemp!�Ƿ���, rsTemp!ָ�������, rsTemp!���÷���) = False Then
                               rsTemp.Close
                               Cancel = True
                               Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            rsTemp.Close
                        Else
                            Set mshDrug.Recordset = rsTemp
                            rsTemp.Close
                            Call SetDrugWidth(sngLeft, sngTop)
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    Call ��ʾ�����
                    'End If
                
                Case mconIntCol����
                    '�޴���
                    If strKey = "" Then
                        If .TxtVisible = True Then
                            .TextMatrix(.Row, mconIntCol����) = ""
                        End If
                        If .ColData(mconIntColЧ��) = 2 Then
                            .Col = mconIntColЧ��
                        Else
                            .Col = mconIntCol����
                        End If
                        
                        
                        Cancel = True
                        Exit Sub
                    End If
                    
                    If zlCommFun.ActualLen(strKey) > mintBatchNoLen Then
                        MsgBox "���ų��Ȳ��ܳ���" & mintBatchNoLen & "λ��" & Int(mintBatchNoLen / 2) & "������! ,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                Case mconIntColЧ��
                    '�д���
                    If strKey <> "" Then
                        If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                            strKey = TranNumToDate(strKey)
                            If strKey = "" Then
                                MsgBox "ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            .Text = strKey
                            Exit Sub
                        End If
                        If Not IsDate(strKey) Then
                            MsgBox "ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                    ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntColЧ��) Then
                    
                        If .TxtVisible = True Then
                            .Text = " "
                            Exit Sub
                        End If
                        
                        Exit Sub
                    End If
                    
                Case mconIntCol�ɹ���
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "�ɹ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    
                    '���ý��
                    If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol�ɹ���) And .TextMatrix(.Row, mconIntCol����) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɹ����) = Format(.TextMatrix(.Row, mconIntCol����) * strKey, mFMT.FM_���)
                        .TextMatrix(.Row, mconintCol���) = Format(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɹ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɹ����)), mFMT.FM_���)
                    End If
                    
                    ��ʾ�ϼƽ��
                Case mconIntCol�ɹ����
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "�ɹ�������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    
                    If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol�ɹ����) Then
                        If .TextMatrix(.Row, mconIntCol����) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ɹ���) = Format(strKey / .TextMatrix(.Row, mconIntCol����), mFMT.FM_�ɱ���)
                        End If
                        
                        .TextMatrix(.Row, mconintCol���) = Format(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - strKey, mFMT.FM_���)
                        .TextMatrix(.Row, mconIntCol�ɹ����) = Format(strKey, mFMT.FM_���)
                    End If
                    ��ʾ�ϼƽ��
            Case mconIntCol�������
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "������ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol���Ч��)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("�������Ѿ��������ʧЧ��(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol���Ч��)), CDate(strKey)), "yyyy-mm-dd") & "),�Ƿ�Ҫ�������!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    '����ʧЧ��
                    .TextMatrix(.Row, mconIntCol���ʧЧ��) = Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol���Ч��)), CDate(strKey)), "yyyy-mm-dd")
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol�������) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
                
                Case mconIntCol����
                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                        MsgBox "�����������룡", vbOKOnly + vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "��������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If strKey <> "" Then
                        If Val(strKey) = 0 Then
                            MsgBox "�������������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If Abs(Val(strKey)) < 0.001 Then
                            MsgBox "�����ı������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        Dim dblCostPrice As Double
                        If Val(strKey) >= 10 ^ 11 - 1 Then
                            MsgBox "��������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If .TextMatrix(.Row, 0) = "" Then Exit Sub
                        
                        'ȡ������ĵ�����,�������������ĵĲɹ��� ��
                        If Set��ɲ���(Val(.TextMatrix(.Row, 0)), Val(strKey) * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), True, dblCostPrice) = False Then
                            Cancel = True
                            Exit Sub
                        End If
                        .TextMatrix(.Row, mconIntCol�ɹ���) = Format(dblCostPrice * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), mFMT.FM_�ɱ���)
                                
                        strKey = Format(strKey, mFMT.FM_����)
                        .Text = strKey
                        If .TextMatrix(.Row, mconIntCol�ɹ���) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ɹ����) = Format(.TextMatrix(.Row, mconIntCol�ɹ���) * strKey, mFMT.FM_���)
                        End If
                        If Val(.TextMatrix(.Row, mconIntCol�ɹ����)) >= 10 ^ 14 - 1 Then
                            MsgBox "�ɹ�������С��" & (10 ^ 14 - 1) & ",��������������!", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                            .TextMatrix(.Row, mconIntCol�ۼ�) = Format(.TextMatrix(.Row, mconIntCol�ɹ���) / (1 - Split(.TextMatrix(.Row, mconIntColԭ����), "||")(1)), mFMT.FM_���ۼ�)
                        End If
                            
                        If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ۼ۽��) = Format(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mFMT.FM_���)
                              
                        End If
                        If Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) >= 10 ^ 14 - 1 Then
                            MsgBox "�ۼ۽�����С��" & (10 ^ 14 - 1) & ",��������������!", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .TextMatrix(.Row, mconintCol���) = Format(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɹ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɹ����)), mFMT.FM_���)
                        
                    End If
                    ��ʾ�ϼƽ��
                
            End Select
        End With
    ElseIf KeyCode = vbKeyDown And Shift = vbAltMask Then
        mshbill_CommandClick
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'����������Ŀ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, ByVal str���� As String, _
    ByVal str��� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, _
    ByVal intԭЧ�� As Integer, ByVal num����ϵ�� As Double, _
    ByVal int�Ƿ��� As Integer, ByVal dblָ������� As Double, ByVal int���÷��� As Integer) As Boolean
    
    Dim intCount As Integer
    Dim rsStructure As New Recordset
    Dim intCol As Integer
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select һ���Բ���,���Ч�� from �������� where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
    
    SetColValue = False
    With mshBill
        For intCol = 0 To .Cols - 1
            '.TextMatrix(intRow, intCol) = ""
            '2010-5-5 ������ʱ������ֵ
            If mconIntCol���� <> intCol Or Trim(.TextMatrix(intRow, mconIntCol����)) = "" Then
                .TextMatrix(intRow, intCol) = ""
            End If
        Next
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng����ID Then
                    Call MsgBox("�����������Ѿ����ڣ���ϲ��������ӣ�", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntColһ���Բ���) = zlStr.Nvl(rsTemp!һ���Բ���)
        .TextMatrix(intRow, mconIntCol���Ч��) = zlStr.Nvl(rsTemp!���Ч��)
        
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol�ۼ�) = Format(num�ۼ�, mFMT.FM_���ۼ�)
        
        .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(intԭЧ��), "0", intԭЧ��) & "||" & dblָ������� & "||" & int�Ƿ��� & "||" & int���÷���
        
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        SetInputFormat intRow
        
        If Set��ɲ���(lng����ID, 0, True, 0) = False Then
            For intCol = 0 To .Cols - 1
                .TextMatrix(intRow, intCol) = ""
            Next
            Exit Function
        End If
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
'
'Private Function SetStructure(ByVal int����id As Long) As Boolean
'    Dim rsTemp As New Recordset
'
'    SetStructure = False
'    vs��ɲ���.Redraw = False
'
'    If mint�༭״̬ <> 4 Then
'        gstrSQL = "" & _
'            "   SELECT DISTINCT b.����id, b.����,b.���� AS ��Ʒ����, b.���, c.�ϴβ���, b.���㵥λ as ��λ, c.ʵ�ʲ��,c.ʵ�ʽ��, d.�ۼ�, " & _
'            "             (a.���� / a.��ĸ) AS ���,c.��������, b.ָ�������,b.�Ƿ���,b.���÷��� " & _
'            "   FROM ���Ʋ��Ϲ��� a,(select b.����id,a.����,a.����, a.���,b.ָ�������,a.�Ƿ���,a.���㵥λ,b.���÷��� from �շ���ĿĿ¼ a,�������� b where a.id=b.����id and nvl(a.�Ƿ���,0)=0) b," & _
'            "        (SELECT ҩƷid, ʵ�ʲ��,ʵ�ʽ��, �ϴβ���,�������� From ҩƷ��� WHERE �ⷿid =[2] and ����=1) c," & _
'            "        (SELECT �շ�ϸĿid, TO_CHAR (�ּ�," & mOraFMT.FM_���ۼ� & ") �ۼ� From �շѼ�Ŀ WHERE ( (SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (    SYSDATE >= ִ������ AND ��ֹ���� IS NULL) )) d " & _
'            "   Where a.ԭ�ϲ���id = b.����ID " & _
'            "AND a.ԭ�ϲ���id = d.�շ�ϸĿid " & _
'            "AND a.ԭ�ϲ���id = c.ҩƷid (+) " & _
'            "AND a.���Ʋ���id =[1]"
'
'        gstrSQL = gstrSQL & " union " & _
'             "  SELECT DISTINCT b.����id, b.����,b.���� AS ��Ʒ����, b.���, c.�ϴβ���, b.���㵥λ as ��λ, c.ʵ�ʲ��,c.ʵ�ʽ��,TO_CHAR ((c.ʵ�ʽ��/c.ʵ������), " & mOraFMT.FM_���ۼ� & ")  as �ۼ�, " & _
'             "          (a.���� / a.��ĸ) AS ���,c.��������,b.ָ�������,b.�Ƿ���,b.���÷��� " & _
'             "  FROM ���Ʋ��Ϲ��� a,(select b.����id,a.����,a.����,a.���,b.ָ�������,a.�Ƿ���,a.���㵥λ,b.���÷��� from �շ���ĿĿ¼ a,�������� b where a.id=b.����id and nvl(a.�Ƿ���,0)=1) b," & _
'             "      (SELECT ҩƷid, ʵ�ʲ��,ʵ�ʽ��, �ϴβ���,��������,ʵ������ From ҩƷ��� WHERE �ⷿid =[2]  and ����=1 and ʵ������>0 ) c " & _
'             "  Where a.ԭ�ϲ���id = b.����ID " & _
'             "          AND a.ԭ�ϲ���id = c.ҩƷid  " & _
'             "          AND a.���Ʋ���id =[1]"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, int����id, cboType.ItemData(cboType.ListIndex))
'
'
'        If rsTemp.EOF Then
'            vs��ɲ���.Redraw = True
'            Exit Function
'        End If
'
'        With vs��ɲ���
'            .ClearBill
'            Do While Not rsTemp.EOF
'                If rsTemp!���÷��� = 1 Then
'                    MsgBox "���������һ�����÷������ģ�����ǰ�汾��֧�����÷�����������ģ����飡", vbInformation + vbOKOnly, gstrSysName
'                    vs��ɲ���.Redraw = True
'                    Exit Function
'                End If
'
'                .TextMatrix(.Row, mconIntCol������) = "[" & rsTemp!���� & "]" & rsTemp!��Ʒ����
'                .TextMatrix(.Row, mconIntCol�����) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
'                .TextMatrix(.Row, mconIntCol������) = IIf(IsNull(rsTemp!�ϴβ���), "", rsTemp!�ϴβ���)
'                .TextMatrix(.Row, mconIntCol����λ) = rsTemp!��λ
'                .TextMatrix(.Row, mconIntCol���ۼ�) = Format(rsTemp!�ۼ�, mFMT.FM_���ۼ�)
'                .TextMatrix(.Row, mconIntCol����������) = Format(IIf(IsNull(rsTemp!��������), "0", rsTemp!��������), mFMT.FM_����)
'                .TextMatrix(.Row, mconIntCol���������) = rsTemp!���
'                .TextMatrix(.Row, mconintcol��ָ�������) = rsTemp!ָ������� & "||" & IIf(IsNull(rsTemp!�Ƿ���), 0, rsTemp!�Ƿ���) & "||" & IIf(IsNull(rsTemp!���÷���), 0, rsTemp!���÷���)
'                .TextMatrix(.Row, mconintcol��ʵ�ʲ��) = IIf(IsNull(rsTemp!ʵ�ʲ��), "0", rsTemp!ʵ�ʲ��)
'                .TextMatrix(.Row, mconintcol��ʵ�ʽ��) = IIf(IsNull(rsTemp!ʵ�ʽ��), "0", rsTemp!ʵ�ʽ��)
'                .TextMatrix(.Row, mconintcol������id) = rsTemp!����ID
'
'
'                If .Row = .Rows - 1 Then
'                    .Rows = .Rows + 1
'                End If
'                .Row = .Row + 1
'                rsTemp.MoveNext
'            Loop
'        End With
'    Else            '�鿴
'        gstrSQL = "" & _
'            "   SELECT DISTINCT a.����id, c.����,c.���� AS ��Ʒ����,b.һ���Բ���,b.���Ч��, c.���," & _
'            "           a.����, c.���㵥λ as ��λ,a.ʵ������,a.�ɱ���,a.�ɱ����,a.���ۼ�,a.���۽��,a.��� " & _
'            "   FROM (  Select ҩƷid as ����id,����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,��� " & _
'            "           From ҩƷ�շ���¼ " & _
'            "           Where   no=[1] and ����=16 and ��¼״̬=[2]" & _
'            "                   and ���ϵ��=-1 and ����=[4] AND ����id =[3]) a," & _
'            "       �������� b,�շ���ĿĿ¼ c " & _
'            "Where a.����id = b.����ID and a.����id=c.id "
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, txtNo.Tag, mint��¼״̬, int����id, mshBill.Row)
'
'        If rsTemp.EOF Then
'            vs��ɲ���.Redraw = True
'            Exit Function
'        End If
'        With vs��ɲ���
'            .ClearBill
'            Do While Not rsTemp.EOF
'                .TextMatrix(.Row, mconIntCol������) = "[" & rsTemp!���� & "]" & rsTemp!��Ʒ����
'                .TextMatrix(.Row, mconIntCol�����) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
'                .TextMatrix(.Row, mconIntCol������) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
'                .TextMatrix(.Row, mconIntCol����λ) = rsTemp!��λ
'                .TextMatrix(.Row, mconIntCol������) = Format(rsTemp!ʵ������, mFMT.FM_����)
'                .TextMatrix(.Row, mconIntCol���ɹ���) = Format(rsTemp!�ɱ���, mFMT.FM_�ɱ���)
'                .TextMatrix(.Row, mconIntCol���ɹ����) = Format(IIf(IsNull(rsTemp!�ɱ����), 0, rsTemp!�ɱ����), mFMT.FM_���)
'                .TextMatrix(.Row, mconIntCol���ۼ�) = Format(rsTemp!���ۼ�, mFMT.FM_���ۼ�)
'                .TextMatrix(.Row, mconIntCol���ۼ۽��) = Format(IIf(IsNull(rsTemp!���۽��), 0, rsTemp!���۽��), mFMT.FM_���)
'                .TextMatrix(.Row, mconintCol�����) = Format(IIf(IsNull(rsTemp!���), 0, rsTemp!���), mFMT.FM_���)
'                .TextMatrix(.Row, mconintcol������id) = rsTemp!����ID
'
'                If .Row = .Rows - 1 Then
'                    .Rows = .Rows + 1
'                End If
'                .Row = .Row + 1
'                rsTemp.MoveNext
'            Loop
'
'        End With
'        rsTemp.Close
'        vs��ɲ���.Redraw = True
'        Exit Function
'    End If
'    rsTemp.Close
'    SetStructure = True
'    vs��ɲ���.Redraw = True
'    Exit Function
'errHandle:
'    vs��ɲ���.Redraw = True
'    Exit Function
'
'End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    If mblnEdit = False Then Exit Sub
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then
            .ColData(mconIntColЧ��) = 5
            Exit Sub
        End If
        
        If .TextMatrix(intRow, mconIntColһ���Բ���) = "1" Then
            .ColData(mconIntCol�������) = 2
            .ColData(mconIntCol���ʧЧ��) = 5
        Else
            .ColData(mconIntCol�������) = 5              '��ֹ
            .ColData(mconIntCol���ʧЧ��) = 5
        End If
        
         
        If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
            If Split(.TextMatrix(intRow, mconIntColԭ����), "||")(0) = "0" Then
                .ColData(mconIntColЧ��) = 5
            Else
                .ColData(mconIntColЧ��) = 2                '���������
            End If
        Else
            .ColData(mconIntColЧ��) = 5
        End If
    End With
End Sub


Private Sub mshDrug_DblClick()
    mshDrug_KeyPress 13
    
End Sub

Private Sub mshDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshDrug
        If KeyCode = vbKeyRight Then
            If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
                
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If .LeftCol <> 0 Then
                .LeftCol = .LeftCol - 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyHome Then
            If .LeftCol <> 0 Then
                .LeftCol = 0
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyEnd Then
            For i = .Cols - 1 To 0 Step -1
                sngWidth = sngWidth + .ColWidth(i)
                If sngWidth > .Width Then
                    .LeftCol = i + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub mshDrug_KeyPress(KeyAscii As Integer)
    With mshDrug
        If KeyAscii = 13 Then
            If Not SetColValue(mshBill.Row, .TextMatrix(.Row, 4), "[" & .TextMatrix(.Row, 0) & "]" & .TextMatrix(.Row, 1), _
                 .TextMatrix(.Row, 2), .TextMatrix(.Row, 5), Val(.TextMatrix(.Row, 8)), _
                 IIf(IsNull(.TextMatrix(.Row, 10)), "0", .TextMatrix(.Row, 10)), .TextMatrix(.Row, 7), Val(.TextMatrix(.Row, 11)), Val(.TextMatrix(.Row, 12)), Val(.TextMatrix(.Row, 13))) Then
                mshBill.SetFocus
                mshBill.Col = mconIntCol����
                .Visible = False
                Exit Sub
            End If
            .Visible = False
            mshBill.Text = "[" & .TextMatrix(.Row, 2) & "]" & .TextMatrix(.Row, 4)
            
            mshBill.Col = mconIntCol����
            
            mshBill.SetFocus
        End If
    End With
                
            
End Sub

Private Sub mshDrug_LostFocus()
    SaveFlexState mshDrug, mstrCaption
    If mshDrug.Visible Then mshDrug.Visible = False
End Sub

Private Sub vs��ɲ���_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub
 

Private Sub r_DecideInput(strInput As String, Cancel As Boolean)

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

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    
    If txtNo.Locked = False Then
        If Trim(txtNo.Text) = "" Then
            ShowMsgBox "���ݺŲ���Ϊ��"
            Exit Function
        End If
        
        If InStr(1, txtNo.Text, "'") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
            Exit Function
        End If
        
        If LenB(StrConv(txtNo.Text, vbFromUnicode)) > txtNo.MaxLength Then
            ShowMsgBox "���ݺų���,���������" & CInt(txtNo.MaxLength / 2) & "�����֣���ò�Ҫ���֣���" & txtNo.MaxLength & "���ַ�!"
            txtNo.SetFocus
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol����)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol����))) = "" Then
                        MsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol����))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "��" & intLop & "�����ĵ����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
          
                    If Split(.TextMatrix(intLop, mconIntColԭ����), "||")(0) <> "0" Then
                        If .TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntColЧ��) = "" Then
                            MsgBox "��" & intLop & "�е�������Ч������,����������ż�Ч����Ϣ�������뵥���У�", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            If .TextMatrix(intLop, mconIntCol����) = "" Then
                                .Col = mconIntCol����
                            Else
                                .Col = mconIntColЧ��
                            End If
                            Exit Function
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol����)) > 9999999999# Then
                        MsgBox "��" & intLop & "�����ĵ��������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɹ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "�����ĵĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɹ����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "�����ĵ��ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    If Checkԭ�Ͽ��(Val(.TextMatrix(intLop, 0)), Val(.TextMatrix(intLop, mconIntCol����)) * Val(.TextMatrix(intLop, mconIntCol����ϵ��))) = False Then
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
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
    Dim chrNo As Variant
    Dim lng��� As Long, lngStockID As Long, lng��¼�� As Long, lng����ID As Long, lng�Ƽ���ID As Long
    Dim str���� As String, strЧ�� As String, str�������� As String, str������� As String, str���Ч�� As String
    Dim dbl�������� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double
    Dim dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim strժҪ As String, str������ As String
    Dim intRow As Integer, cllProc As Collection
    
    SaveCard = False
    Set cllProc = New Collection
    With mshBill
        chrNo = Trim(txtNo)
        '���ܼ�¼��
        lng��¼�� = 0
        For intRow = 1 To .Rows - 1
             If .TextMatrix(intRow, 0) <> "" Then
                    lng��¼�� = lng��¼�� + 1
             End If
        Next
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint�༭״̬ = 1 Then 'mbln�������� Or
            If chrNo <> "" Then
                If CheckNOExists(69, chrNo) Then Exit Function
            End If
        
            If chrNo = "" Then chrNo = sys.GetNextNo(69, lngStockID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNo.Tag = chrNo
        
        
        lng�Ƽ���ID = cboType.ItemData(cboType.ListIndex)
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_���Ʋ������_Delete('" & mstr���ݺ� & "')"
            AddArray cllProc, gstrSQL
        End If
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = .TextMatrix(intRow, 0)
                
                str���� = .TextMatrix(intRow, mconIntCol����)
                strЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                dbl�������� = Round(Val(.TextMatrix(intRow, mconIntCol����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / .TextMatrix(intRow, mconIntCol����ϵ��), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mconIntCol�ɹ����)), g_С��λ��.obj_���С��.���С��)
                dbl���ۼ� = Round(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / .TextMatrix(intRow, mconIntCol����ϵ��), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)), g_С��λ��.obj_���С��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mconintCol���)), g_С��λ��.obj_���С��.���С��)
                
                str������� = IIf(.TextMatrix(intRow, mconIntCol�������) = "", "", .TextMatrix(intRow, mconIntCol�������))
                str���Ч�� = IIf(.TextMatrix(intRow, mconIntCol���ʧЧ��) = "", "", .TextMatrix(intRow, mconIntCol���ʧЧ��))
                
                lng��� = intRow
                'Zl_���Ʋ������_Insert
                gstrSQL = "Zl_���Ʋ������_Insert("
                '  No_In         In ҩƷ�շ���¼.NO%Type,
                gstrSQL = gstrSQL & "'" & chrNo & "',"
                '  ���_In       In ҩƷ�շ���¼.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                gstrSQL = gstrSQL & "" & lngStockID & ","
                '  �Է�����id_In In ҩƷ�շ���¼.�Է�����id%Type,
                gstrSQL = gstrSQL & "" & lng�Ƽ���ID & ","
                '  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type,
                gstrSQL = gstrSQL & "" & dbl�������� & ","
                '  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type,
                gstrSQL = gstrSQL & "" & dbl���ۼ� & ","
                '  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type,
                gstrSQL = gstrSQL & "" & dbl���۽�� & ","
                '  ������_In     In ҩƷ�շ���¼.������%Type,
                gstrSQL = gstrSQL & "'" & str������ & "',"
                '  ����_In       In ҩƷ�շ���¼.����%Type := Null,
                gstrSQL = gstrSQL & "'" & str���� & "',"
                '  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                gstrSQL = gstrSQL & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & " ,"
                '  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
                gstrSQL = gstrSQL & IIf(str������� = "", "Null", "to_date('" & Format(str�������, "yyyy-MM-dd") & "','yyyy-mm-dd')") & " ,"
                '  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
                gstrSQL = gstrSQL & IIf(str���Ч�� = "", "Null", "to_date('" & Format(str���Ч��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & " ,"
                '  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                gstrSQL = gstrSQL & "'" & strժҪ & "',"
                '  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
                gstrSQL = gstrSQL & "to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS'),"
                '  ��¼��_In     In Integer := 0
                gstrSQL = gstrSQL & "" & lng��¼�� & ")"
                AddArray cllProc, gstrSQL
            End If
        Next
    End With
    
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllProc, mstrCaption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
ErrHand:
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
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol�ɹ����))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & Format(curTotal, mFMT.FM_���)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���)
    lblDifference.Caption = "��ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
End Sub

Private Sub ��ʾ�����()
    Dim rsTemp As New ADODB.Recordset
    Dim dbl���� As Double
    Dim str��λ As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo ErrHandle
    If mshBill.TextMatrix(mshBill.Row, mconIntCol����) = "" Then
        stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
    Select Case mintUnit
        Case 1
            strQuantity = "��������/����ϵ�� "
        Case 0
            strQuantity = "�������� "
    End Select
    
    gstrSQL = "" & _
        "   Select b.����ID , Sum(" & strQuantity & ") as ���� " & _
        "   From ҩƷ��� a,�������� b " & _
        "   Where a.����=1 and a.ҩƷid=b.����id and ��������<>0 And �ⷿID=[1]" & _
        "       and b.����ID=[2]" & _
        "   Group by b.����ID "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "--��ʾ�����", cboStock.ItemData(cboStock.ListIndex), intID)
    
    With rsTemp
        If .EOF Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        dbl���� = IIf(IsNull(!����), 0, !����)
        stbThis.Panels(2).Text = "�����ĵ�ǰ�����Ϊ[" & dbl���� & "]"
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ��ʾԭ�Ͽ����()
    Dim rsTemp As New ADODB.Recordset
    Dim dbl���� As Double, lng����ID As Long
    
    On Error GoTo ErrHandle
    With vs��ɲ���
        If .TextMatrix(.Row, .ColIndex("ԭ���ϱ��뼰����")) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("ԭ���ϱ��뼰����")))
    End With
    
    gstrSQL = "" & _
        "   Select b.ID as ����id, Sum(��������) as ����,b.���㵥λ as ��λ " & _
        "   From ҩƷ��� a,�շ���ĿĿ¼ b " & _
        "   Where a.����=1 and a.ҩƷid=b.id and ��������<>0 And �ⷿID=[1]" & _
        "       and b.ID=[2]" & _
        "   Group by b.ID,b.���㵥λ "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "---��ʾԭ�Ͽ����", cboType.ItemData(cboType.ListIndex), lng����ID)
    
    With rsTemp
        If .EOF Then
            stbThis.Panels(2).Text = "��ǰ�޿��"
            Exit Sub
        End If
        dbl���� = !����
        stbThis.Panels(2).Text = "�����ĵ�ǰ�����Ϊ[" & dbl���� & "]" & zlStr.Nvl(!��λ)
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
    Dim strNo As String
    strNo = txtNo.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1713", mint��¼״̬, mintUnit, 1713, "����������ⵥ", strNo
End Sub


'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, mstrCaption & "--ȡ�ֶγ���"
    GetBatchNoLen = rsTemp.Fields(0).DefinedSize
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Set��ɲ���(ByVal lng����ID As Long, Optional dbl�������� As Double = 0, Optional bln����� As Boolean = True, Optional ByRef dblOut�ɱ��� As Double = 0) As Boolean
    '------------------------------------------------------------------------------
    '����:������ص���ɲ���]
    '    int����id-����ID
    '    dbl��������-���Ʋ������������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/03/21
    '------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, rsSort As New ADODB.Recordset
    Dim lngRow As Long, lng�ⷿID As Long, arrtemp As Variant
    Dim dblʣ������ As Double, dbl��ǰ���� As Double, dbl�������� As Double
    Dim dbl��� As Double, dbl���� As Double, dbl�ɱ���� As Double, dblSum�ɱ���� As Double
    Dim blnContinue As Boolean '����
    Dim blnʵ�� As Boolean
    err = 0: On Error GoTo ErrHand:
    blnContinue = False
    Set��ɲ��� = False
    
    vs��ɲ���.Redraw = flexRDNone
    
    On Error GoTo ErrHand
    If mint�༭״̬ <> 4 Then
        gstrSQL = "" & _
        "   SELECT DISTINCT b.ID as ����id, b.����,b.���� AS ��Ʒ����, b.���, b.���㵥λ as ��λ,d.�ۼ�, " & _
        "             (a.���� / a.��ĸ) AS ���, C.ָ�������,B.�Ƿ���,C.���÷��� " & _
        "   FROM ���Ʋ��Ϲ��� a,�շ���ĿĿ¼ B,�������� C," & _
        "        (  SELECT �շ�ϸĿid,�ּ� as �ۼ� From �շѼ�Ŀ WHERE  ((SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (SYSDATE >= ִ������ AND ��ֹ���� IS NULL))" & _
        GetPriceClassString("") & ") d " & _
        "   Where a.ԭ�ϲ���id = b.ID And (B.վ��=[2] or B.վ�� is null) and A.ԭ�ϲ���id=c.����ID  AND a.ԭ�ϲ���id = d.�շ�ϸĿid(+)" & _
        "         AND a.���Ʋ���id =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID, gstrNodeNo)
        If rsTemp.EOF Then
            vs��ɲ���.Redraw = flexRDBuffered
            Exit Function
        End If

        lng�ⷿID = cboType.ItemData(cboType.ListIndex)
        
        With vs��ɲ���
            .Clear (1)
            .Rows = 2
            lngRow = 1
            Do While Not rsTemp.EOF
                    If mint�༭״̬ <> 1 Then
                        gstrSQL = "" & _
                            "   SELECT nvl(����,0) ����," & _
                            "          nvl(��������,0)  as ��������, nvl(ʵ������,0)  as ʵ������, " & _
                            "          nvl(ʵ�ʲ��,0)  as ʵ�ʲ��,nvl(ʵ�ʽ��,0) as ʵ�ʽ��,���ۼ�," & _
                            "         �ϴβ���,�ϴ�����,�ϴ���������, Ч�� ,nvl(��������,0) as ʵ�ʿ�������" & _
                            "   From ҩƷ��� " & _
                            "   WHERE ҩƷid=[1] and �ⷿid =[2]  and ����=1 " & _
                            "   Union ALL " & _
                            "   Select nvl(����,0) as ����,��д���� as ��������,0 as ʵ������,0 as ʵ�ʲ��,0 as ʵ�ʽ��,���ۼ�,����,����,��������,Ч��,0 as ʵ�ʿ������� " & _
                            "   From ҩƷ�շ���¼ " & _
                            "   where ����=16 and ҩƷid=[1] and NO=[3] and ���ϵ��=-1"
                        gstrSQL = "" & _
                            "   SELECT nvl(����,0) ����," & _
                            "           sum(nvl(��������,0)) as ��������,sum(nvl(ʵ������,0)) as ʵ������, " & _
                            "           Sum(nvl(ʵ�ʲ��,0)) as ʵ�ʲ��,sum(nvl(ʵ�ʽ��,0)) as ʵ�ʽ��, " & _
                            "           Sum(nvl(ʵ�ʿ�������,0)) as ʵ�ʿ�������,max(���ۼ�) as ���ۼ�," & _
                            "           max(�ϴβ���) as �ϴβ���,max(�ϴ�����) as �ϴ�����,max(�ϴ���������) as �ϴ���������,max(Ч��) as Ч�� " & _
                            "   From (" & gstrSQL & ") " & _
                            "   Group by nvl(����,0) " & _
                            "   Order by ����"
                    Else
                        gstrSQL = "" & _
                            "   SELECT nvl(����,0) ����," & _
                            "           sum(nvl(��������,0)) as ��������,sum(nvl(ʵ������,0)) as ʵ������, " & _
                            "           Sum(nvl(ʵ�ʲ��,0)) as ʵ�ʲ��,sum(nvl(ʵ�ʽ��,0)) as ʵ�ʽ��, " & _
                            "           sum(nvl(��������,0)) as ʵ�ʿ�������,Max(���ۼ�) as ���ۼ�," & _
                            "           max(�ϴβ���) as �ϴβ���,max(�ϴ�����) as �ϴ�����,max(�ϴ���������) as �ϴ���������,max(Ч��) as Ч�� " & _
                            "   From ҩƷ��� " & _
                            "   WHERE ҩƷid=[1] and �ⷿid =[2]  and ����=1 " & _
                            "   Group by nvl(����,0) " & _
                            "   Order by ����"
                    End If
                    
                    Set rsSort = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.Nvl(rsTemp!����ID)), lng�ⷿID, mstr���ݺ�)
                    dblʣ������ = dbl�������� * Val(zlStr.Nvl(rsTemp!���))
                    .TextMatrix(lngRow, .ColIndex("ԭ���ϱ��뼰����")) = "[" & zlStr.Nvl(rsTemp!����) & "]" & zlStr.Nvl(rsTemp!��Ʒ����)
                    .Cell(flexcpData, lngRow, .ColIndex("ԭ���ϱ��뼰����")) = zlStr.Nvl(rsTemp!����ID)
                    .TextMatrix(lngRow, .ColIndex("���")) = zlStr.Nvl(rsTemp!���)
                    .TextMatrix(lngRow, .ColIndex("��λ")) = zlStr.Nvl(rsTemp!��λ)
                    dblSum�ɱ���� = 0
                    .Cell(flexcpData, lngRow, .ColIndex("��λ")) = 0 & "," & 0 & "," & 0 & "," & 0 & "," & Val(zlStr.Nvl(rsTemp!ָ�������))
                    Do While Not rsSort.EOF
                        .TextMatrix(lngRow, .ColIndex("ԭ���ϱ��뼰����")) = "[" & zlStr.Nvl(rsTemp!����) & "]" & zlStr.Nvl(rsTemp!��Ʒ����)
                        .Cell(flexcpData, lngRow, .ColIndex("ԭ���ϱ��뼰����")) = zlStr.Nvl(rsTemp!����ID)
                        .TextMatrix(lngRow, .ColIndex("���")) = zlStr.Nvl(rsTemp!���)
                        .TextMatrix(lngRow, .ColIndex("��λ")) = zlStr.Nvl(rsTemp!��λ)
                        .Cell(flexcpData, lngRow, .ColIndex("��λ")) = zlStr.Nvl(rsSort!ʵ�ʿ�������) & "," & zlStr.Nvl(rsSort!ʵ������) & "," & zlStr.Nvl(rsSort!ʵ�ʲ��) & "," & zlStr.Nvl(rsSort!ʵ�ʽ��) & "," & Val(zlStr.Nvl(rsTemp!ָ�������))
                        .TextMatrix(lngRow, .ColIndex("����")) = zlStr.Nvl(rsSort!�ϴ�����)
                        
                        .Cell(flexcpData, lngRow, .ColIndex("����")) = zlStr.Nvl(rsSort!����)
                        
                        If Val(zlStr.Nvl(rsTemp!�Ƿ���)) = 0 Then
                            '����
                            .TextMatrix(lngRow, .ColIndex("�ۼ�")) = Format(Val(zlStr.Nvl(rsTemp!�ۼ�)), mFMT.FM_���ۼ�)
                            .Cell(flexcpData, lngRow, .ColIndex("�ۼ�")) = zlStr.Nvl(rsTemp!�ۼ�)
                        ElseIf Val(zlStr.Nvl(rsSort!ʵ������)) <> 0 Then
                            If Val(zlStr.Nvl(rsSort!���ۼ�)) <> 0 Then
                                .TextMatrix(lngRow, .ColIndex("�ۼ�")) = Format(Val(zlStr.Nvl(rsSort!���ۼ�)), mFMT.FM_���ۼ�)
                                .Cell(flexcpData, lngRow, .ColIndex("�ۼ�")) = Val(zlStr.Nvl(rsSort!���ۼ�))
                            Else
                                .TextMatrix(lngRow, .ColIndex("�ۼ�")) = Format(Val(zlStr.Nvl(rsSort!ʵ�ʽ��)) / Val(zlStr.Nvl(rsSort!ʵ������)), mFMT.FM_���ۼ�)
                                .Cell(flexcpData, lngRow, .ColIndex("�ۼ�")) = Val(zlStr.Nvl(rsSort!ʵ�ʽ��)) / Val(zlStr.Nvl(rsSort!ʵ������))
                            End If
                        Else
                            .TextMatrix(lngRow, .ColIndex("�ۼ�")) = ""
                            .Cell(flexcpData, lngRow, .ColIndex("�ۼ�")) = ""
                        End If
                        dbl�������� = Val(zlStr.Nvl(rsSort!��������))
                        If dbl�������� >= dblʣ������ Then
                            dbl��ǰ���� = dblʣ������
                        Else
                            dbl��ǰ���� = dbl��������
                        End If
                        
                        .TextMatrix(lngRow, .ColIndex("����")) = Format(dbl��ǰ����, mFMT.FM_����)
                        .Cell(flexcpData, lngRow, .ColIndex("����")) = dbl��ǰ����
                        .TextMatrix(lngRow, .ColIndex("�ۼ۽��")) = Format(dbl��ǰ���� * Val(.Cell(flexcpData, lngRow, .ColIndex("�ۼ�"))), mFMT.FM_���)
                
'                        Call ��֤�����ۼ���(lng�ⷿID, Val(zlStr.Nvl(rsTemp!����ID)), Val(.Cell(flexcpData, lngRow, .ColIndex("����"))), 1, Val(zlStr.Nvl(rsSort!ʵ�ʲ��)), Val(zlStr.Nvl(rsSort!ʵ�ʽ��)), Val(zlStr.Nvl(rsTemp!ָ�������)) / 100, dbl��ǰ����, Val(.TextMatrix(lngRow, .ColIndex("�ۼ۽��"))), dbl���, dbl����, dbl�ɱ����)
                        
'                        .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(dbl����, mFMT.FM_�ɱ���)
'                        .TextMatrix(lngRow, .ColIndex("�ɱ����")) = Format(dbl�ɱ����, mFMT.FM_���)
'                        .TextMatrix(lngRow, .ColIndex("���")) = Format(dbl���, mFMT.FM_���)
                        
                        .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(Get�ɱ���(Val(zlStr.Nvl(rsTemp!����ID)), lng�ⷿID, Val(.Cell(flexcpData, lngRow, .ColIndex("����")))), mFMT.FM_�ɱ���)
                        .TextMatrix(lngRow, .ColIndex("�ɱ����")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) * dbl��ǰ����, mFMT.FM_���)
                        .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(lngRow, .ColIndex("�ɱ����"))), mFMT.FM_���)
                        
                        dblʣ������ = dblʣ������ - Val(zlStr.Nvl(rsSort!��������))
                        If dblʣ������ <= 0 Then
                            Exit Do
                        End If
                        rsSort.MoveNext
                        If rsSort.EOF Then Exit Do
                        .Rows = .Rows + 1
                        lngRow = lngRow + 1
                    Loop
                    
                    If Round(dblʣ������, 7) > 0 Then
                         .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) + dblʣ������, mFMT.FM_����)
                         .Cell(flexcpData, lngRow, .ColIndex("����")) = Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) + dblʣ������
                         .TextMatrix(lngRow, .ColIndex("�ۼ۽��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) * Val(.TextMatrix(lngRow, .ColIndex("�ۼ�"))), mFMT.FM_���)
                         arrtemp = Split(.Cell(flexcpData, lngRow, .ColIndex("��λ")) & ",,,,,", ",")
                         ' NVL(rsSort!ʵ�ʿ�������) & "," & NVL(rsSort!ʵ������) & "," & NVL(rsSort!ʵ�ʲ��) & "," & NVL(rsSort!ʵ�ʽ��),ָ�������
'                        Call ��֤�����ۼ���(lng�ⷿID, Val(NVL(rsTemp!����ID)), Val(.Cell(flexcpData, lngRow, .ColIndex("����"))), 1, Val(ArrTemp(2)), Val(ArrTemp(3)), Val(ArrTemp(4)) / 100, Val(.Cell(flexcpData, lngRow, .ColIndex("����"))), Val(.TextMatrix(lngRow, .ColIndex("�ۼ۽��"))), dbl���, dbl����, dbl�ɱ����)
'                        .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(dbl����, mFMT.FM_�ɱ���)
'                        .TextMatrix(lngRow, .ColIndex("�ɱ����")) = Format(dbl�ɱ����, mFMT.FM_���)
'                        .TextMatrix(lngRow, .ColIndex("���")) = Format(dbl���, mFMT.FM_���)
                        
                        .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(Get�ɱ���(Val(zlStr.Nvl(rsTemp!����ID)), lng�ⷿID, Val(.Cell(flexcpData, lngRow, .ColIndex("����")))), mFMT.FM_�ɱ���)
                        .TextMatrix(lngRow, .ColIndex("�ɱ����")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) * Val(.Cell(flexcpData, lngRow, .ColIndex("����"))), mFMT.FM_���)
                        .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(.TextMatrix(lngRow, .ColIndex("�ۼ۽��"))) - Val(.TextMatrix(lngRow, .ColIndex("�ɱ����"))), mFMT.FM_���)
                        
                        blnʵ�� = Val(zlStr.Nvl(rsTemp!�Ƿ���)) = 1
                        
                        If bln����� Then
                            If mint����� = 0 Then
                                '�����
                                If blnʵ�� Or Val(zlStr.Nvl(rsTemp!���÷���)) = 1 Then
                                    vs��ɲ���.Redraw = flexRDBuffered
                                    MsgBox "���������ĵ�ԭ�����ġ�" & .TextMatrix(lngRow, .ColIndex("ԭ���ϱ��뼰����")) & "�����ÿ���������������ԭ�����ĵĿ�棡", vbInformation + vbOKOnly, gstrSysName
                                    Exit Function
                                End If
                            ElseIf mint����� = 1 Then
                                '��飬����
                                If blnʵ�� Or Val(zlStr.Nvl(rsTemp!���÷���)) = 1 Then
                                    vs��ɲ���.Redraw = flexRDBuffered
                                    MsgBox "���������ĵ�ԭ�����ġ�" & .TextMatrix(lngRow, .ColIndex("ԭ���ϱ��뼰����")) & "�����ÿ���������������ԭ�����ĵĿ�棡", vbInformation + vbOKOnly, gstrSysName
                                    Exit Function
                                ElseIf blnContinue = False Then
                                    If MsgBox("���������ĵ�ԭ�����ġ�" & .TextMatrix(lngRow, .ColIndex("ԭ���ϱ��뼰����")) & "�����ÿ�����������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        vs��ɲ���.Redraw = flexRDBuffered
                                        Exit Function
                                    End If
                                    blnContinue = True
                                End If
                            ElseIf mint����� = 2 Then
                                '��ֹ
                                vs��ɲ���.Redraw = flexRDBuffered
                                MsgBox "���������ĵ�ԭ�����ġ�" & .TextMatrix(lngRow, .ColIndex("ԭ���ϱ��뼰����")) & "�����ÿ���������������ԭ�����ĵĿ�棡", vbInformation + vbOKOnly, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                    .Rows = .Rows + 1
                    lngRow = lngRow + 1
                    rsTemp.MoveNext
                Loop
                dblSum�ɱ���� = 0
                '��ɱ���
                For lngRow = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ���ϱ��뼰����"))) <> 0 Then
                        If Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) <> 0 Then
                            dblSum�ɱ���� = dblSum�ɱ���� + Val(.TextMatrix(lngRow, .ColIndex("�ɱ����")))
                        End If
                    End If
                Next
            End With
            If dbl�������� <> 0 Then
                dblOut�ɱ��� = dblSum�ɱ���� / dbl��������
            Else
                dblOut�ɱ��� = 0
            End If
    Else            '�鿴
        gstrSQL = "" & _
            "   SELECT DISTINCT a.����id,a.����, c.����,c.���� AS ��Ʒ����,b.һ���Բ���,b.���Ч��, c.���," & _
            "           a.����, c.���㵥λ as ��λ,a.ʵ������,a.�ɱ���,a.�ɱ����,a.���ۼ�,a.���۽��,a.��� " & _
            "   FROM (  Select ҩƷid as ����id,����,����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,��� " & _
            "           From ҩƷ�շ���¼ " & _
            "           Where   no=[1] and ����=16 and ��¼״̬=[2]" & _
            "                   and ���ϵ��=-1 and ����=[4] AND ����id =[3]) a," & _
            "       �������� b,�շ���ĿĿ¼ c " & _
            "Where a.����id = b.����ID and a.����id=c.id And (C.վ��=[5] or C.վ�� is null) "
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, txtNo.Tag, mint��¼״̬, lng����ID, mshBill.Row, gstrNodeNo)
        
        If rsTemp.EOF Then
            vs��ɲ���.Redraw = flexRDBuffered
            Exit Function
        End If
        With vs��ɲ���
            .Clear (1)
            .Rows = 2
            lngRow = 1
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, .ColIndex("ԭ���ϱ��뼰����")) = "[" & zlStr.Nvl(rsTemp!����) & "]" & zlStr.Nvl(rsTemp!��Ʒ����)
                .Cell(flexcpData, .ColIndex("ԭ���ϱ��뼰����")) = zlStr.Nvl(rsTemp!����ID)
                .TextMatrix(lngRow, .ColIndex("���")) = zlStr.Nvl(rsTemp!���)
                .TextMatrix(lngRow, .ColIndex("��λ")) = zlStr.Nvl(rsTemp!��λ)
                .TextMatrix(lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
                .TextMatrix(lngRow, .ColIndex("����")) = Format(rsTemp!ʵ������, mFMT.FM_����)
                .TextMatrix(lngRow, .ColIndex("�ۼ�")) = Format(Val(zlStr.Nvl(rsTemp!���ۼ�)), mFMT.FM_���ۼ�)
                .TextMatrix(lngRow, .ColIndex("�ۼ۽��")) = Format(Val(zlStr.Nvl(rsTemp!���۽��)), mFMT.FM_���)
                .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���)), mFMT.FM_�ɱ���)
                .TextMatrix(lngRow, .ColIndex("�ɱ����")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ����)), mFMT.FM_���)
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(zlStr.Nvl(rsTemp!���)), mFMT.FM_���)
                .Rows = .Rows + 1
                lngRow = lngRow + 1
                rsTemp.MoveNext
            Loop
        End With
        rsTemp.Close
        vs��ɲ���.Redraw = flexRDBuffered
        Exit Function
    End If
    rsTemp.Close
    Set��ɲ��� = True
    vs��ɲ���.Redraw = flexRDBuffered
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    vs��ɲ���.Redraw = flexRDBuffered
    Exit Function
End Function

Private Sub vs��ɲ���_BeforeSort(ByVal Col As Long, Order As Integer)
    
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim intRow As Integer
    
    
    With vs��ɲ���
        '�Զ�������
        If .ExplorerBar > &H1000& Then Exit Sub
'
'        .GetSelection lngRow, lngCol, lngRows, lngCols
'        .Redraw = flexRDNone
'        'Ӧ�õ��ǿ���
'        For intRow = .Rows - 1 To .FixedRows Step -1
'            If Len(.TextMatrix(intRow, Col)) Then Exit For
'        Next
'
'        If intRow > .FixedRows Then
'            .Select .FixedRows, Col, intRow, Col
'            .Sort = Order
'        End If
'
'        ' �ָ�ѡ��
'        .Select lngRow, lngCol, lngRows, lngCols
'        .Redraw = flexRDDirect
        Order = 0
    End With
End Sub

Private Sub vs��ɲ���_EnterCell()
    Call ��ʾԭ�Ͽ����
End Sub
Private Function Checkԭ�Ͽ��(ByVal lng����ID As Long, ByVal dbl�������� As Double) As Boolean
    '------------------------------------------------------------------------------
    '����:���ԭ�Ͽ���Ƿ�Ϸ�
    '����:�п��,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/23
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, rsStorck As New ADODB.Recordset
    Dim dbl�������� As Double
    Dim lng�ⷿID As Long, blnʵ�� As Boolean, blnContinue As Boolean
    
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "" & _
    "   SELECT DISTINCT b.ID as ����id, b.����,b.���� AS ��Ʒ����, b.���, b.���㵥λ as ��λ, " & _
    "             (a.���� / a.��ĸ) AS ���,B.�Ƿ���,C.���÷��� " & _
    "   FROM ���Ʋ��Ϲ��� a,�շ���ĿĿ¼ B,�������� C" & _
    "   Where a.ԭ�ϲ���id = b.ID and A.ԭ�ϲ���id=c.����ID" & _
    "         AND a.���Ʋ���id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
    If rsTemp.EOF Then
        gstrSQL = "Select ����,���� From �շ���ĿĿ¼ where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
        If Not rsTemp.EOF Then
            ShowMsgBox "���Ʋ���:" & zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!����) & vbCrLf & " û����ص���ɲ���,����!"
        End If
        Exit Function
    End If
    
    lng�ⷿID = cboType.ItemData(cboType.ListIndex)
    
    If mint�༭״̬ = 2 Then
        gstrSQL = "" & _
            "   SELECT nvl(����,0) ����," & _
            "          nvl(��������,0)  as ��������, nvl(ʵ������,0)  as ʵ������, " & _
            "          nvl(ʵ�ʲ��,0)  as ʵ�ʲ��,nvl(ʵ�ʽ��,0) as ʵ�ʽ��, " & _
            "         �ϴβ���,�ϴ�����,�ϴ���������, Ч�� ,nvl(��������,0) as ʵ�ʿ�������" & _
            "   From ҩƷ��� " & _
            "   WHERE ҩƷid=[1] and �ⷿid =[2]  and ����=1 " & _
            "   Union ALL " & _
            "   Select nvl(����,0) as ����,��д���� as ��������,0 as ʵ������,0 as ʵ�ʲ��,0 as ʵ�ʽ��,����,����,��������,Ч��,0 as ʵ�ʿ������� " & _
            "   From ҩƷ�շ���¼ " & _
            "   where ����=16 and ҩƷid=[1] and NO=[3] and ���ϵ��=-1"
        gstrSQL = "" & _
            "   SELECT  sum(nvl(��������,0)) as �������� " & _
            "   From (" & gstrSQL & ") "
    Else
        gstrSQL = "" & _
            "   SELECT sum(nvl(��������,0)) as �������� " & _
            "   From ҩƷ��� " & _
            "   WHERE ҩƷid=[1] and �ⷿid =[2]  and ����=1 "
    End If
    
    Do While Not rsTemp.EOF
        Set rsStorck = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.Nvl(rsTemp!����ID)), lng�ⷿID, mstr���ݺ�)
        If rsStorck.EOF Then
            dbl�������� = 0
        Else
            dbl�������� = Val(zlStr.Nvl(rsStorck!��������))
        End If
        If Round(dbl��������, 7) < Round(dbl�������� * Val(zlStr.Nvl(rsTemp!���)), 7) Then
            blnʵ�� = Val(zlStr.Nvl(rsTemp!�Ƿ���)) = 1
            If mint����� = 0 Then
                '�����
                If blnʵ�� Or Val(zlStr.Nvl(rsTemp!���÷���)) = 1 Then
                    vs��ɲ���.Redraw = flexRDBuffered
                    MsgBox "���������ĵ�ԭ�����ġ�" & zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!��Ʒ����) & "�����ÿ���������������ԭ�����ĵĿ�棡", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint����� = 1 Then
                '��飬����
                If blnʵ�� Or Val(zlStr.Nvl(rsTemp!���÷���)) Then
                    MsgBox "���������ĵ�ԭ�����ġ�" & zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!��Ʒ����) & "�����ÿ���������������ԭ�����ĵĿ�棡", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                ElseIf blnContinue = False Then
                    If MsgBox("���������ĵ�ԭ�����ġ�" & zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!��Ʒ����) & "�����ÿ�����������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                    blnContinue = True
                End If
            ElseIf mint����� = 2 Then
                '��ֹ
                MsgBox "���������ĵ�ԭ�����ġ�" & zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!��Ʒ����) & "�����ÿ���������������ԭ�����ĵĿ�棡", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
        rsTemp.MoveNext
    Loop
    Checkԭ�Ͽ�� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

