VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceFormula 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtJL 
      Height          =   300
      Left            =   2955
      MaxLength       =   20
      TabIndex        =   24
      Top             =   3690
      Width           =   900
   End
   Begin VB.ComboBox cboData 
      Height          =   300
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.OptionButton optMode 
      Caption         =   "ɢװ(&0)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   97
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.OptionButton optMode 
      Caption         =   "��Ƭ(&1)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   1
      Left            =   1155
      TabIndex        =   19
      Top             =   97
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.OptionButton optMode 
      Caption         =   "����(&2)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   2
      Left            =   2310
      TabIndex        =   18
      Top             =   97
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   3210
      MousePointer    =   7  'Size N S
      TabIndex        =   17
      Top             =   3975
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   3210
      TabIndex        =   16
      Top             =   4245
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   3120
      TabIndex        =   15
      Top             =   3960
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   3870
      MousePointer    =   9  'Size W E
      TabIndex        =   14
      Top             =   3975
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   4
      Left            =   4185
      MousePointer    =   7  'Size N S
      TabIndex        =   13
      Top             =   4110
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fra��ҩ 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ComboBox cboҩ�� 
         Height          =   300
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   0
         Width           =   1920
      End
      Begin VB.TextBox txt���� 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "1"
         Top             =   0
         Width           =   450
      End
      Begin VB.Label lblҩ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ��"
         Height          =   240
         Left            =   1680
         TabIndex        =   12
         Top             =   60
         Width           =   405
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   720
         TabIndex        =   11
         Top             =   60
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   5895
      Picture         =   "frmAdviceFormula.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "ȷ��(F2)"
      Top             =   3720
      Width           =   450
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   6450
      Picture         =   "frmAdviceFormula.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "ȡ��(Esc)"
      Top             =   3720
      Width           =   450
   End
   Begin VB.CommandButton cmdInsert 
      Height          =   315
      Left            =   5400
      Picture         =   "frmAdviceFormula.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "����(&A)"
      Top             =   3720
      Visible         =   0   'False
      Width           =   450
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExt 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6900
      _cx             =   12171
      _cy             =   3254
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
      ForeColorSel    =   16777215
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
      Rows            =   7
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceFormula.frx":7366
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.CommandButton cmd 
         Caption         =   "��"
         Height          =   240
         Left            =   4920
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   720
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vs��ҩ��� 
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   6975
      _cx             =   12303
      _cy             =   2355
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
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceFormula.frx":7472
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Editable        =   2
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
      Begin VB.CommandButton cmd��̬ 
         Caption         =   "ɢװ(&D)"
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblJL 
      Caption         =   "����"
      Height          =   240
      Left            =   2565
      TabIndex        =   25
      Top             =   3705
      Width           =   405
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   3840
      X2              =   4515
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line lin 
      Index           =   1
      X1              =   3840
      X2              =   4515
      Y1              =   3750
      Y2              =   3750
   End
   Begin VB.Line lin 
      Index           =   2
      X1              =   3840
      X2              =   4515
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Line lin 
      Index           =   3
      X1              =   3840
      X2              =   4515
      Y1              =   3810
      Y2              =   3810
   End
   Begin VB.Line lin 
      Index           =   4
      X1              =   3840
      X2              =   4515
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line lin 
      Index           =   5
      X1              =   3840
      X2              =   4515
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line lin 
      Index           =   6
      X1              =   3840
      X2              =   4515
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Line lin 
      Index           =   7
      X1              =   3840
      X2              =   4515
      Y1              =   3930
      Y2              =   3930
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�巨"
      Height          =   180
      Left            =   105
      TabIndex        =   7
      Top             =   3780
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblNumZY 
      Caption         =   "��___ζ"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblZYStock 
      Caption         =   "��ҩ�����ʾ"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmAdviceFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================================================
'��ڲ�����
Private mclsInsure As Object '�ô���Ϊ�������������������в�ʹ��ҽ����Ҳ������
Private mlngHwnd As Long '���ڶ�λ�Ŀؼ����
Private mint��Ч As Integer
Private mstr�Ա� As String
Private mint�������� As Integer  '1-����,2-סԺ
Private mint������� As Integer '1-����,2-סԺ,3-�����סԺ
Private mbytUseType As Byte      '0=ҽ���´�,1-·����Ŀ��ҽ������,2-���·������Ŀ

'��:��������ĿID,��ҩ�䷽ʱΪ�䷽ID��ζ��ҩID
Private mlng��ĿID As Long


'��/��:���Ӷ�������,����ʱһ��Ϊ��
'      ��ҩ="��ҩID1,����1,��ע1;��ҩID2,����2,��ע2;...|�巨ID|��ҩ��̬|����|ҩ��ID|����"
Private mstrExtData As String
Private mstr�䷽��ϸ As String

'��:�����������Ҫ,��������ȡ��������
Private mlng����ID As Long
Private mvar����ID As Variant '��ҳID��Һŵ���
Private mintӤ�� As Integer
Private mint���� As Integer 'ҽ�����˵�����
Private mlng���˿���id As Long '����ȷ����ҩ�䷽��ȱʡҩ��
Private mlngPreRow��ҩ�� As Long '�ϻ���һ����ҩ�䷽��ҩ��
Private mlngҩƷID As Long       'ѡ����ѡ�е���ҩ���ID
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mlng�������� As Long

Private mint���� As Integer  '0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS),3-���׷���(·����Ŀ����)����
Private mblnҽ�� As Boolean '�Ƿ�ҽ���򹫷Ѳ���

'������ҽ���ӿ�GetItemInfo�����ص�ժҪ����Ҫ�����䷽��
Private mstrժҪ As String

'���ڲ�����
Private mblnOK As Boolean '��

Private mlng��ҩ�� As Long
Private mstr������ҩ�� As String

Private mblnFirst As Boolean
Private mblnReturn As Boolean '�Ƿ��˻س�ȷ��
Private mcol������� As Collection  '��ҩ��IDΪ���洢ҩƷ�Ĺ�����������1,����;���2,����|δ��������
Private mbytSize As Byte '�����С 0-С���壨9�ţ���1-�����壨12�ţ�
Private mintÿ��ζ�� As Integer '��ҩ�䷽ÿ����ҩ��ζ��
Private Enum E��ҩ���
    col��� = 0
    col���� = 1
    col������λ = 2
    col���� = 3
    col���� = 4
    colҩƷID = 5
End Enum
Private mblnChangeSel As Boolean
Private mstrPrivs As String             'Ȩ��
Private mfrmParent As Object
Private mblnSelf As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal objclsInsure As Object, ByVal lngHwnd As Long, ByRef t_Pati As TYPE_PatiInfoEx, ByVal int���� As Integer, _
             ByVal bytUseType As Byte, ByVal int��Ч As Integer, ByVal int������� As Integer, Optional ByVal int�������� As Integer, _
             Optional ByVal lng��Ŀid As Long, Optional ByRef strExtData As String, _
             Optional ByRef strժҪ As String, Optional ByVal lngҩƷID As Long, Optional ByVal lngPreRow��ҩ�� As Long, Optional ByVal strҩƷ�۸�ȼ� As String) As Boolean
'����:
'     frmParent         ������
'     objclsInsure      �ô���Ϊ�������������������в�ʹ��ҽ����Ҳ������,���Ҫ����ҽ������
'     lngHwnd           ���ڶ�λ�Ŀؼ����,�����øô���Ŀؼ�
'     t_Pati            ������Ϣ
'     int����           0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS),3-���׷���(·����Ŀ����)����
'     bytUseType        0=ҽ���´�,1-·����Ŀ��ҽ������,2-���·������Ŀ,3-·����Ŀ��������
'     int��Ч           ��Ҫ�����ҽ����Ч 0-������1-����
'     int�������       ��ҽ��Ҫ����Ĳ������� 1-����������ﲡ�ˣ���첡�ˣ��������˵�) 2-סԺ��ֻ��סԺ���ˣ�
'     int��������       ���øô���Ĺ���վ���� 1-����ҽ������վ 2-סԺҽ������վ(����ֻ�����ҩ�䷽)
'     lng��Ŀid         ��������ĿID , ��ҩ�䷽ʱΪ�䷽ID��ζ��ҩID
'     lngҩƷID         ѡ����ѡ�е���ҩ���ID
'     lngPreRow��ҩ��   Ĭ����ҩ�����ϻ���һ����ҩ�䷽��ҩ��
'���أ�
'     strExtData        ���Ӷ������� , ����ʱһ��Ϊ��
'                       ��ҩ = "��ҩID1,����1,��ע1;��ҩID2,����2,��ע2;...|�巨ID|��ҩ��̬|����|ҩ��ID"
'     strժҪ           ��ҽ���ӿ�GetItemInfo�����ص�ժҪ����Ҫ�����䷽�ġ�

    Set mfrmParent = frmParent
    Set mclsInsure = objclsInsure
    mlngHwnd = lngHwnd
    With t_Pati
        mblnҽ�� = .blnҽ��
        mint���� = .int����
        mintӤ�� = .intӤ��
        mlng����ID = .lng����ID
        mlng���˿���id = .lng���˿���ID
        mvar����ID = IIF(.str�Һŵ� = "", .lng��ҳID, .str�Һŵ�)
        mstr�Ա� = .str�Ա�
    End With
    mint���� = int����
    mbytUseType = bytUseType
    mint��Ч = int��Ч
    mint������� = int�������
    If mint���� <> 3 Then
        mint�������� = int��������
    Else
        mint�������� = IIF(int������� = 1, 1, 2)
    End If
    mlng��ĿID = lng��Ŀid
    mstrExtData = strExtData
    mstrժҪ = strժҪ
    mlngҩƷID = lngҩƷID
    mlngPreRow��ҩ�� = lngPreRow��ҩ��
    mstrҩƷ�۸�ȼ� = strҩƷ�۸�ȼ�
    mblnOK = False
    mlng�������� = 0 '�ڲ����и�ֵ
    
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    
    strExtData = mstrExtData
    strժҪ = mstrժҪ
    
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmd��̬_Click()
'���ܣ���ζҩ����ɢװ���䲻��ʱ������ɢװ���
    Dim lngҩ��ID As Long, dbl���� As Double, lngҩƷID As Long
    Dim strKey As String
    
    strKey = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
    If strKey <> "" Then lngҩ��ID = Val(Split(strKey, "_")(0))
    
    dbl���� = Val(vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + 1))
    lngҩƷID = Val(cmd��̬.Tag)    'ȱʡ���
        
    Call mcol�������.Remove("_" & strKey)
    mcol�������.Add lngҩƷID & "," & dbl����, "_" & strKey
    
    Call Show��ҩ���(lngҩ��ID, dbl����, 0)
    mblnChangeSel = True
    vsExt.SetFocus
    mblnChangeSel = False
End Sub

Private Sub Form_Resize()
    Dim lngAppend As Long
    Dim lngMinRows As Long
    Dim lngRows As Long, i As Long
    Dim lngHeight As Long, lngTotalHeight As Long
    
    On Error Resume Next
    
    fraBorder(0).Left = 0
    fraBorder(0).Top = 0
    fraBorder(0).Width = Me.ScaleWidth
    fraBorder(1).Top = fraBorder(0).Top + fraBorder(0).Height
    fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
    fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    fraBorder(2).Left = 0
    fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
    fraBorder(2).Width = Me.ScaleWidth
    fraBorder(3).Top = fraBorder(0).Top + fraBorder(0).Height
    fraBorder(3).Left = 0
    fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    
    vsExt.Left = fraBorder(3).Width
    vsExt.Top = fraBorder(0).Top + fraBorder(0).Height + fra��ҩ.Height
    vsExt.Width = Me.ScaleWidth - fraBorder(3).Width * 2

    fra��ҩ.Left = vsExt.Left
    fra��ҩ.Top = fraBorder(0).Top + fraBorder(0).Height
    fra��ҩ.Width = vsExt.Width
    
    If Me.Visible = False Then
        For i = 0 To optMode.Count - 1
            Set optMode(i).Container = fra��ҩ
            optMode(i).Top = lbl����.Top
        Next
        optMode(0).Left = 60
        optMode(1).Left = optMode(0).Left + optMode(0).Width
        optMode(2).Left = optMode(1).Left + optMode(1).Width
    
        lbl����.Left = optMode(2).Left + optMode(2).Width + 360
        txt����.Left = lbl����.Left + lbl����.Width
        lblҩ��.Left = txt����.Left + txt����.Width + 360
        cboҩ��.Left = lblҩ��.Left + lblҩ��.Width
    End If
    
    vsExt.Height = Me.ScaleHeight - fraBorder(2).Height * 2 - (cboData.Height + 150) - fra��ҩ.Height - vs��ҩ���.Height - IIF(lblZYStock.Visible, lblZYStock.Height + 60, 0)
    lngMinRows = 7
    With vsExt
        For i = .FixedRows To .Rows - 1
            If Replace(.Cell(flexcpText, i, 0, i, .Cols - 1), Chr(9), "") <> "" Then
                lngMinRows = i + .FixedRows
            Else
                Exit For
            End If
        Next
        lngRows = Int((vsExt.Height - vsExt.RowHeight(0) - 15) / (vsExt.RowHeight(1) + 15))
        If lngRows < lngMinRows Then lngRows = lngMinRows
        .Rows = lngRows
    End With
    Call SetSplitLine
    
    With vs��ҩ���
        .Top = vsExt.Top + vsExt.Height + 30
        .Left = vsExt.Left
        .Width = vsExt.Width
        .ColWidth(col����) = .Width - .ColWidth(col���) - .ColWidth(col������λ) - .ColWidth(col����) - .ColWidth(col����)
    End With

    
    
    '��ҩ��ʾ���
    lblZYStock.Top = vs��ҩ���.Top + vs��ҩ���.Height + 60
    lblZYStock.Left = vs��ҩ���.Left
    lblZYStock.Width = vs��ҩ���.Width
    cboData.Top = vs��ҩ���.Top + vs��ҩ���.Height + 60 + IIF(mint���� <> 3, lblZYStock.Height, 0) + 60
    lblData.Top = cboData.Top + (cboData.Height - lblData.Height) / 2
    cmdOK.Top = cboData.Top + (cboData.Height - cmdOK.Height) / 2
    cmdCancel.Top = cmdOK.Top
        

    lblData.Left = 200
    cboData.Left = lblData.Left + lblData.Width + fraBorder(3).Width
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdCancel.Height
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - fraBorder(1).Width * 3
        
    cboData.Width = cboҩ��.Width
    lblJL.Width = lblData.Width
    lblJL.Top = lblData.Top
    lblJL.Left = cboData.Left + cboData.Width + 200
    
    txtJL.Top = cboData.Top
    txtJL.Left = lblJL.Left + lblJL.Width
    
    cmdInsert.Top = cmdOK.Top
    cmdInsert.Left = cmdOK.Left - cmdInsert.Width - 100
    lblNumZY.Top = cmdOK.Top + 45
    lblNumZY.Left = cmdInsert.Left - lblNumZY.Width - 100
    
    txtJL.Width = lblNumZY.Left - txtJL.Left - 400
    
    Me.Refresh
End Sub

Private Sub Form_Activate()
    If mblnFirst And vsExt.TabStop And vsExt.Enabled And vsExt.Visible And Not Me.ActiveControl Is vsExt Then
        mblnFirst = False: vsExt.SetFocus '�������Ϊʲô�Զ���λ��rtfAppend����ȥ�ˡ�
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    ElseIf KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(",;|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '����������ָ�����������
    End If
End Sub

Private Sub Form_Load()
    Dim blnMulti As Boolean, vRect As RECT
    Dim str���� As String, i As Long, lngBaseHeight As Long
    
    Me.Height = 2325
    
    '�߿�����
    For i = 0 To fraBorder.UBound
        fraBorder(i).BackColor = vbButtonFace
    Next
    Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
    Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
    Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
    Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
    lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
    lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
    lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
    lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
    lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
    lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
    lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
    lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
    
    If mint������� = 0 Then mint������� = IIF(mint���� = 3, 3, 2) 'ȱʡΪסԺ,����ȱʡΪסԺ������
    mblnOK = False
    mblnFirst = True
    
    If mint���� = 0 Then
        If mint�������� = 1 Then
            mbytSize = zlDatabase.GetPara("����", glngSys, pm����ҽ��վ, "0")
        Else
            mbytSize = zlDatabase.GetPara("����", glngSys, pmסԺҽ��վ, "0")
        End If
    ElseIf mint���� = 1 Then
        mbytSize = zlDatabase.GetPara("����", glngSys, pmסԺ��ʿվ, "0")
    ElseIf mint���� = 2 Then
        mbytSize = zlDatabase.GetPara("����", glngSys, pmҽ������վ, "0")
    End If

    
    '��ʼ�������ʽ
    mintÿ��ζ�� = IIF(Val(zlDatabase.GetPara(213, glngSys)) = 4, 4, 3)
    mstrժҪ = ""
    Set mcol������� = New Collection
    mlng��ҩ�� = Val(zlDatabase.GetPara(IIF(mint�������� = 2, "סԺ", "����") & "ȱʡ��ҩ��", glngSys, IIF(mint�������� = 2, pmסԺҽ���´�, pm����ҽ���´�), , , , , mlng���˿���id))
    mstr������ҩ�� = zlDatabase.GetPara(IIF(mint�������� = 2, "סԺ", "����") & "������ҩ��", glngSys, IIF(mint�������� = 2, pmסԺҽ���´�, pm����ҽ���´�), , , , , mlng���˿���id)
    mstrPrivs = GetInsidePrivs(IIF(mint�������� = 1, pm����ҽ���´�, pmסԺҽ���´�))
    vs��ҩ���.Visible = True
    '��ʼ�������
    Call Grid.Init(vs��ҩ���, "���,1305,1;����,5070,1;������λ,900,1;����,900,1;����,900,1;ҩƷID")
    fra��ҩ.Visible = True
    lblData.Visible = True
    cboData.Visible = True
    lblData.Caption = "�巨"
    lblNumZY.Visible = True
    cmdInsert.Visible = mbytUseType <> 3
    If mint���� <> 3 Then
       lblZYStock.Visible = True
       Me.Height = Me.Height + lblZYStock.Height + 60
    End If
    If Not Init��ҩ�䷽ Then Unload Me: Exit Sub

    '��������
    Call zlControl.SetPubFontSize(Me, mbytSize)
    '�ָ����Ի�
    lngBaseHeight = Me.Height
    Call RestoreWinState(Me, App.ProductName, 2)
    
     '10.26.80���ӹ����ʾ����ǰ���Ի�����ĸ߶ȿ��ܲ���
    If Me.Height < lngBaseHeight Then
        Me.Height = lngBaseHeight
    End If
    
    '���嶨λ
    GetWindowRect mlngHwnd, vRect
    Me.Left = (vRect.Left - 1) * Screen.TwipsPerPixelX
    Me.Top = (vRect.Top - 1) * Screen.TwipsPerPixelY - Me.Height
    Call Form_Resize
    Call RefreshWeiNum
End Sub

Private Sub SetSplitLine()
'���ܣ�������ҩ�䷽������зָ���
    Dim lngRow As Long, lngCol As Long
    Dim i As Long
        
    vsExt.Redraw = False
    lngRow = vsExt.Row: lngCol = vsExt.Col
    mblnChangeSel = True
    For i = 1 To mintÿ��ζ��
        vsExt.Select vsExt.FixedRows, i * 4 - 1, vsExt.Rows - 1, i * 4 - 1
        vsExt.CellBorder &HC0C0C0, 0, 0, 1, 0, 0, 0
    Next

    vsExt.ColWidth(0) = ((vsExt.Width - 60) / mintÿ��ζ�� - 285) * 0.45 '��ζ��ҩ
    vsExt.ColWidth(1) = ((vsExt.Width - 60) / mintÿ��ζ�� - 285) * 0.22  '��ζ����
    vsExt.ColWidth(2) = 285 '��λ
    vsExt.ColWidth(3) = ((vsExt.Width - 60) / mintÿ��ζ�� - 285) * 0.33 '��ע
    For i = 4 To vsExt.Cols - 1
        vsExt.ColWidth(i) = vsExt.ColWidth(i - 4)
    Next
    
    vsExt.Row = lngRow: vsExt.Col = lngCol
    mblnChangeSel = False
    vsExt.Redraw = True
End Sub

Private Function Init��ҩ�䷽() As Boolean
'���ܣ���ʼ����ҩ�䷽����ʽ������
'������mstrExtData=����ÿζ��ҩ��Ϣ���巨��Ϣ�Ĵ�,Ϊ��ʱ��ʾ��������ҩ�䷽
    Dim rsTmp As New ADODB.Recordset
    Dim rsTmpCopy As New ADODB.Recordset
    
    Dim strSQL As String, i As Long, j As Long
    Dim lngRow As Long, lngCol As Long, blnDo As Boolean
    Dim str��ҩIDs As String, lng�巨ID As Long, lngFirstҩ��ID As Long, lngFirstҩƷID As Long
    Dim arr��ҩ As Variant, lng��̬ As Long, dbl���� As Double, str������� As String
    Dim lngCurҩ��ID As Long, lngNextҩ��ID As Long, lngҩƷID As Long
    Dim lngҩ��ID As Long, bln�䷽ As Boolean
    Dim strKey As String, blnSameɢװ��ҩ As Boolean
    Dim str���� As String
    
    mstr�䷽��ϸ = ""
    vsExt.Clear
    vsExt.Cols = mintÿ��ζ�� * 4: vsExt.Rows = 7
  
    vsExt.FixedCols = 0: vsExt.FixedRows = 1
    vsExt.ColAlignment(0) = 1 '��ζ��ҩ
    vsExt.ColAlignment(1) = 7 '��ζ����
    vsExt.ColAlignment(2) = 1 '��λ
    vsExt.ColAlignment(3) = 1 '��ע

    
    Me.Width = (Me.Width - Me.ScaleWidth) + IIF(mbytSize = 0, 2320, 2870) * mintÿ��ζ�� + 250
    Me.Height = Me.Height + vs��ҩ���.Height + fra��ҩ.Height + 600

    For i = 4 To vsExt.Cols - 1
        vsExt.ColAlignment(i) = vsExt.ColAlignment(i - 4)
    Next
    vsExt.MergeCells = flexMergeFixedOnly
    vsExt.MergeRow(0) = True
    vsExt.Cell(flexcpAlignment, 0, 0, 0, vsExt.Cols - 1) = 1
    vsExt.Cell(flexcpText, 0, 0, 0, vsExt.Cols - 1) = "����ѡ����ҩ��̬,Ȼ�����������в�ҩ,��ζ����,��ע����*��ѡȡ��ҩ���ע��"
    vsExt.GridColor = vsExt.BackColor
    vsExt.Editable = flexEDKbdMouse
       
    vs��ҩ���.TabIndex = vsExt.TabIndex + 1
    txtJL.TabIndex = cboData.TabIndex + 1
    cmdOK.TabIndex = txtJL.TabIndex + 1

    On Error GoTo errH
    txt����.Text = "1"
    txt����.Tag = "1"
    If mint��Ч = 0 Then    '�������丶��
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
    End If
    
    If mstrExtData <> "" Then '�޸�
        lng�巨ID = Val(Split(mstrExtData, "|")(1))
        lng��̬ = Val(Split(mstrExtData, "|")(2))
        txt����.Text = Val(Split(mstrExtData, "|")(3))
        txtJL.Text = Split(mstrExtData, "|")(5)
        arr��ҩ = Split(Split(mstrExtData, "|")(0), ";")
        lngҩƷID = Val(Split(arr��ҩ(0), ",")(0))
                
        For i = 0 To UBound(arr��ҩ)
            str��ҩIDs = str��ҩIDs & "," & CStr(Split(arr��ҩ(i), ",")(0))
        Next
        str��ҩIDs = Mid(str��ҩIDs, 2)
                
        strSQL = "Select/*+ Rule*/ a.ID,b.ҩƷID,a.����,a.���㵥λ,c.��� as ��� From ������ĿĿ¼ A,ҩƷ��� B,�շ���ĿĿ¼ C " & _
            "Where a.ID = b.ҩ��ID And b.ҩƷID = C.ID And b.ҩƷID IN(Select Column_Value From Table(f_Num2list([1]))) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��ҩIDs)
        Set rsTmpCopy = zlDatabase.CopyNewRec(rsTmp)
        
        
        If vsExt.Rows < -Int(rsTmp.RecordCount / (-1 * mintÿ��ζ��)) + 1 Then
            vsExt.Rows = -Int(rsTmp.RecordCount / (-1 * mintÿ��ζ��)) + 1
        End If
        lngRow = vsExt.FixedRows: lngCol = 0
        
        '�������ڵ����ݺʹ�����ʾ
        dbl���� = 0
        str������� = ""
        For i = 0 To UBound(arr��ҩ)
            blnDo = True
            dbl���� = dbl���� + Val(Split(arr��ҩ(i), ",")(1))
            str������� = str������� & ";" & Split(arr��ҩ(i), ",")(0) & "," & Split(arr��ҩ(i), ",")(1)
            If i < UBound(arr��ҩ) Then
                rsTmp.Filter = "ҩƷID=" & CStr(Split(arr��ҩ(i), ",")(0))
                lngCurҩ��ID = rsTmp!ID
                rsTmp.Filter = "ҩƷID=" & CStr(Split(arr��ҩ(i + 1), ",")(0))
                lngNextҩ��ID = rsTmp!ID
                
                If lngCurҩ��ID = lngNextҩ��ID And lng��̬ <> 0 Then
                    blnDo = False   '��ɢװ��ͬ��ҩ�Ĳ�ͬ���������ۼ�
                End If
            End If
            
            If blnDo Then
                rsTmp.Filter = "ҩƷID=" & CStr(Split(arr��ҩ(i), ",")(0))
                If Not rsTmp.EOF Then
                    str���� = rsTmp!����
                    If lng��̬ = 0 Then 'ɢװ
                        strKey = rsTmp!ID & "_" & rsTmp!ҩƷID
                        
                        rsTmpCopy.Filter = "ID=" & rsTmp!ID
                        If rsTmpCopy.RecordCount > 0 Then
                            If rsTmpCopy.RecordCount > 1 Then blnSameɢװ��ҩ = True
                            If Not IsNull(rsTmp!���) Then str���� = str����
                        End If
                    Else
                        strKey = "" & rsTmp!ID
                    End If
                    
                    str������� = Mid(str�������, 2)
                    mcol�������.Add str�������, "_" & strKey
                    str������� = ""
                
                    vsExt.TextMatrix(lngRow, lngCol) = str����
                    vsExt.TextMatrix(lngRow, lngCol + 1) = FormatEx(dbl����, 5): dbl���� = 0
                    vsExt.TextMatrix(lngRow, lngCol + 2) = NVL(rsTmp!���㵥λ)
                    vsExt.TextMatrix(lngRow, lngCol + 3) = CStr(Split(arr��ҩ(i), ",")(2))
                    
                    '���ڻָ���ʾ�ļ�¼
                    vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
                    vsExt.Cell(flexcpData, lngRow, lngCol + 1) = vsExt.TextMatrix(lngRow, lngCol + 1)
                    vsExt.Cell(flexcpData, lngRow, lngCol + 2) = strKey '��¼��ҩID
                    vsExt.Cell(flexcpData, lngRow, lngCol + 3) = vsExt.TextMatrix(lngRow, lngCol + 3)
                                    
                    '��һλ��
                    If lngCol + 4 > vsExt.Cols - 1 Then
                        lngRow = lngRow + 1: lngCol = 0
                    Else
                        lngCol = lngCol + 4
                    End If
                End If
            End If
        Next
    Else '����
        strSQL = "Select a.ID,a.���,a.����,a.���㵥λ From ������ĿĿ¼ a Where a.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID)
        If rsTmp!��� = "7" Then
            '�����˵�ζ�в�ҩ
            vsExt.TextMatrix(vsExt.FixedRows, 0) = rsTmp!����
            vsExt.TextMatrix(vsExt.FixedRows, 2) = NVL(rsTmp!���㵥λ)
            
            '���ڻָ���ʾ�ļ�¼
            vsExt.Cell(flexcpData, vsExt.FixedRows, 0) = vsExt.TextMatrix(vsExt.FixedRows, 0)
            lngFirstҩ��ID = CLng(rsTmp!ID)
            
            '������Ʒ���´�ʱ��ѡ�������ص�ҩƷIDΪ0
            If mlngҩƷID = 0 Then
                Set rsTmp = Get��ҩ���(lngFirstҩ��ID)
                If rsTmp.RecordCount > 0 Then
                    lngҩƷID = rsTmp!ҩƷID
                    
                    '���ֻ��һ�ֹ���һ����̬����ȡ�ù�����̬������ȱʡΪɢװ��̬
                    lng��̬ = Val("" & rsTmp!��ҩ��̬)
                    rsTmp.Filter = "��ҩ��̬<>" & lng��̬
                    If rsTmp.RecordCount > 1 Then lng��̬ = 0
                Else
                    MsgBox "δ�ҵ���ҩƷ�κο��õĹ����ѡ������ҩƷ", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                strSQL = "Select ��ҩ��̬ From ҩƷ��� Where ҩƷID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҩƷID)
                lng��̬ = Val("" & rsTmp!��ҩ��̬)
                lngҩƷID = mlngҩƷID
            End If
                        
            If lng��̬ = 0 Then 'ɢװ
                strKey = lngFirstҩ��ID & "_" & lngҩƷID
                mcol�������.Add lngҩƷID & ",0", "_" & strKey
            Else
                strKey = lngFirstҩ��ID
                mcol�������.Add "", "_" & strKey
            End If
            If lng��̬ = 0 Then 'ɢװ
                Set rsTmp = Get��ҩ���(lngFirstҩ��ID, lng��̬)
                If rsTmp.RecordCount > 1 Then
                    rsTmp.Filter = "ҩƷID =" & lngҩƷID
                    vsExt.TextMatrix(vsExt.FixedRows, 0) = rsTmp!����
                    vsExt.Cell(flexcpData, vsExt.FixedRows, 0) = vsExt.TextMatrix(vsExt.FixedRows, 0)
                End If
            End If
            vsExt.Cell(flexcpData, vsExt.FixedRows, 2) = strKey '��¼��ҩID
        Else
            '�������䷽��Ŀ
            strSQL = "Select A.ID,A.����,b.�շ�ϸĿid as ҩƷid,A.���㵥λ,B.��������,B.ҽ������,C.���" & _
                " From ������ĿĿ¼ A,������Ŀ��� B,�շ���ĿĿ¼ C" & _
                " Where A.ID=B.������ĿID And B.�������ID=[1] And c.Id(+) = b.�շ�ϸĿid" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) And A.������� IN([2],3) Order By B.���"
                
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID, mint�������)
            If rsTmp.EOF Then
                MsgBox "����ҩ�䷽��ǰ����Ч���䷽��ɣ����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                Exit Function
            End If
            
            bln�䷽ = True
            If vsExt.Rows < -Int(rsTmp.RecordCount / (-1 * mintÿ��ζ��)) + 1 Then
                vsExt.Rows = -Int(rsTmp.RecordCount / (-1 * mintÿ��ζ��)) + 1
            End If
            lngRow = vsExt.FixedRows: lngCol = 0
            
            '�������õ����ݵĴ�����ʾ
            For i = 1 To rsTmp.RecordCount
                vsExt.TextMatrix(lngRow, lngCol) = rsTmp!����
                vsExt.TextMatrix(lngRow, lngCol + 1) = NVL(rsTmp!��������)
                vsExt.TextMatrix(lngRow, lngCol + 2) = NVL(rsTmp!���㵥λ)
                vsExt.TextMatrix(lngRow, lngCol + 3) = NVL(rsTmp!ҽ������)
                
                '���ڻָ���ʾ�ļ�¼
                vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
                vsExt.Cell(flexcpData, lngRow, lngCol + 1) = vsExt.TextMatrix(lngRow, lngCol + 1)
                 '��¼��ҩID(����ɢװ����ҩƷID���������������Ϊ"ҩ��Id_ҩƷID")
                vsExt.Cell(flexcpData, lngRow, lngCol + 2) = CLng(rsTmp!ID) & IIF(NVL(rsTmp!ҩƷID) = "", "", "_" & rsTmp!ҩƷID)
                vsExt.Cell(flexcpData, lngRow, lngCol + 3) = vsExt.TextMatrix(lngRow, lngCol + 3)
                
                If i = 1 Then
                    lngFirstҩ��ID = CLng(rsTmp!ID)
                    lngFirstҩƷID = Val("" & rsTmp!ҩƷID)
                End If
                '��һλ��
                If lngCol + 4 > vsExt.Cols - 1 Then
                    lngRow = lngRow + 1: lngCol = 0
                Else
                    lngCol = lngCol + 4
                End If
                rsTmp.MoveNext
            Next
            
            '��ȡ�䷽��Ŀ��ȱʡ�巨
            strSQL = "Select �÷�ID From �����÷����� Where ����=1 And ��ĿID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID)
            If Not rsTmp.EOF Then lng�巨ID = rsTmp!�÷�ID
            
            '��ȡ��һζҩ��ȱʡ���(��������ȱʡ��̬�Ϳ���ҩ��)
            Set rsTmp = Get��ҩ���(lngFirstҩ��ID, , , True)
            If rsTmp.RecordCount > 0 Then
                If lngFirstҩƷID <> 0 Then
                    lngҩƷID = lngFirstҩƷID
                    rsTmp.Filter = "ҩƷid=" & lngFirstҩƷID
                    If Not rsTmp.EOF Then lng��̬ = Val("" & rsTmp!��ҩ��̬)
                Else
                    lngҩƷID = rsTmp!ҩƷID
                    
                    '���ֻ��һ�ֹ���һ����̬����ȡ�ù�����̬������ȱʡΪɢװ��̬
                    lng��̬ = Val("" & rsTmp!��ҩ��̬)
                    rsTmp.Filter = "��ҩ��̬<>" & lng��̬
                    If rsTmp.RecordCount > 1 Then lng��̬ = 0
                End If
            End If
        End If
    End If
    vsExt.ScrollBars = flexScrollBarNone
    
    If mint�������� = 2 Then
        strSQL = "select a.�������� from ������ҳ a where a.����id=[1] and a.��ҳid=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mvar����ID))
        If Not rsTmp.EOF Then mlng�������� = Val(rsTmp!�������� & "")
    End If
        
    '��ҩ�巨
    strSQL = "Select A.ID,A.����,A.���� From ������ĿĿ¼ A" & _
        " Where A.���='E' And A.��������='3'" & IIF(mlng�������� = 1, "", " And A.������� IN([1],3)") & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        IIF(mint���� <> 3 And mlng�������� <> 1, " And (Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And ����ID=[2])" & _
                            " Or Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID))", "") & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint�������, mlng���˿���id)
    If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
    If rsTmp.EOF Then
        MsgBox "δ�ҵ���Ч����ҩ�巨�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    
    For i = 1 To rsTmp.RecordCount
        cboData.AddItem rsTmp!���� & "-" & rsTmp!����
        cboData.ItemData(cboData.NewIndex) = rsTmp!ID
        If rsTmp!ID = lng�巨ID Then
            Call Cbo.SetIndex(cboData.hwnd, cboData.NewIndex)
        End If
        rsTmp.MoveNext
    Next
    If cboData.ListCount = 1 And cboData.ListIndex = -1 Then Call Cbo.SetIndex(cboData.hwnd, 0)
    
    '����ҩ��(��ָ��ҩ���޶����ʱ����ҩƷID��Ӱ��)
    Call Get��ҩ��(cboҩ��, lngҩƷID, mlng���˿���id, mint�������, mlng��ҩ��)
    If mstrExtData <> "" Then
        lngҩ��ID = Val(Split(mstrExtData, "|")(4))
    Else
        lngҩ��ID = IIF(mlngPreRow��ҩ�� = 0, mlng��ҩ��, mlngPreRow��ҩ��)
    End If
    Call Cbo.Locate(cboҩ��, lngҩ��ID, True)
    If cboҩ��.ListCount > 0 And cboҩ��.ListIndex = -1 Then Call Cbo.SetIndex(cboҩ��.hwnd, 0)
    
    
    '��ҩ��̬
    arr��ҩ = Split("ɢװ(&0),��Ƭ(&1),����(&2)", ",")
    For i = 0 To 2
        optMode(i).Visible = True
        optMode(i).Enabled = True
        optMode(i).Caption = arr��ҩ(i)
        optMode(i).Width = optMode(i).Width + IIF(i = 2, 450, 300)
        If i = lng��̬ Then optMode(i).value = True
    Next
    If blnSameɢװ��ҩ Then
        optMode(1).Enabled = False
        optMode(2).Enabled = False
    End If
    Call SetSameItem
    If mstrExtData = "" And bln�䷽ Then
        '���붨�Ƶġ���ҩ�䷽��ʱ��������Ԥ��
        For i = vsExt.FixedRows To vsExt.Rows - 1
            For j = 0 To vsExt.Cols - 1 Step 4
                strKey = vsExt.Cell(flexcpData, i, j + 2)
                
                lngCurҩ��ID = 0 '���lngCurҩ��ID
                
                If strKey <> "" Then lngCurҩ��ID = Val(Split(strKey, "_")(0))
                
                If lngCurҩ��ID <> 0 Then
                    dbl���� = Val(vsExt.TextMatrix(i, j + 1))
                    'û�����ù���ҩ����ȡ���
                    str������� = ""
                    If InStr(strKey, "_") = 0 Then
                        'ȱʡ���
                        Set rsTmp = Get��ҩ���(lngCurҩ��ID, lng��̬)
                                             
                        strKey = lngCurҩ��ID
                        If rsTmp.RecordCount > 0 Then
                            If lng��̬ = 0 Then 'ɢװ
                                strKey = lngCurҩ��ID & "_" & rsTmp!ҩƷID
                                vsExt.Cell(flexcpData, i, j + 2) = strKey
                            End If
                            str������� = rsTmp!ҩƷID & "," & 0
                        End If
                    ElseIf cboҩ��.ListIndex <> -1 Then
                        Set rsTmp = GetҩƷ���(Val(Split(strKey, "_")(1)), True)
                        If rsTmp.RecordCount > 0 Then str������� = Val(Split(strKey, "_")(1)) & "," & 0
                    End If
                    
                    mcol�������.Add str�������, "_" & strKey
                    
                    Call Split��ҩ���(lngCurҩ��ID, dbl����, strKey)
                    
                    If mcol�������("_" & strKey) = "" Or InStr(mcol�������("_" & strKey), "|") > 0 Then
                        vsExt.Cell(flexcpForeColor, i, j + 1) = vbRed
                    End If
                End If
                If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                    mstr�䷽��ϸ = mstr�䷽��ϸ & ";" & Replace(CStr(mcol�������("_" & strKey)), ";", "," & vsExt.TextMatrix(i, j + 3) & ";") & "," & vsExt.TextMatrix(i, j + 3)
                End If
            Next
        Next
    End If
    
    vsExt.Row = vsExt.FixedRows: vsExt.Col = 1
    Init��ҩ�䷽ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RefreshWeiNum()
'���ܣ�ȷ��ζ��
    Dim intNum As Integer
    Dim i As Long, j As Long
    
    intNum = 0
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = vsExt.FixedCols To vsExt.Cols - 1 Step 4
            If vsExt.TextMatrix(i, j) <> "" And vsExt.Cell(flexcpData, i, j) <> 0 Then intNum = intNum + 1
        Next
    Next
    lblNumZY = "�� " & intNum & " ζ"
End Sub


Private Function Get��ҩ���(ByVal lngҩ��ID As Long, Optional ByVal lng��̬ As Long = -1, Optional ByVal blnFirst As Boolean, Optional ByVal bln�䷽ As Boolean) As ADODB.Recordset
'���ܣ�������ҩ����ID��ȡ��ҩ���
'������bln�䷽ true=�¿�ʱ�����䷽
    Dim strSQL As String, lngҩ��ID As Long
    
    On Error GoTo errH
    If lng��̬ = 0 Then
        Set Get��ҩ��� = GetҩƷ���(lngҩ��ID)
    Else
        If mstr������ҩ�� <> "" Then
            If gblnStock And Not blnFirst Then
                If cboҩ��.ListIndex = -1 Then
                    lngҩ��ID = IIF(mlngPreRow��ҩ�� = 0, mlng��ҩ��, mlngPreRow��ҩ��)
                    If bln�䷽ And mlngPreRow��ҩ�� <> 0 Then lngҩ��ID = 0
                Else
                    lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
                End If
                '����ǳ��׵��ã�ҩ��ID=0�򲻼ӿ������
                If mint���� <> 3 Or lngҩ��ID <> 0 Then
                    strSQL = " And Exists(Select 1 From ҩƷ��� B" & _
                        " Where (Nvl(b.����, 0) = 0 Or b.Ч�� Is Null Or b.Ч��>Trunc(Sysdate))" & _
                        " And b.����=1 And a.ҩƷID=b.ҩƷID" & IIF(lngҩ��ID = 0, "", " And b.�ⷿID=[2]") & _
                        " And b.��������>0)"
                End If
            End If
        End If
    
        strSQL = "Select A.ҩƷID,A.��ҩ��̬,D.���� From ҩƷ��� A,�շ���ĿĿ¼ D Where A.ҩ��ID = [1] And A.ҩƷID = D.ID" & _
             IIF(lng��̬ = -1, "", " And A.��ҩ��̬ = [4]") & strSQL & _
             " And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([3],3)" & _
             " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null) Order by D.����"
        Set Get��ҩ��� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ҩ���", lngҩ��ID, lngҩ��ID, mint�������, lng��̬)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Get��ҩ��(objCbo As ComboBox, ByVal lngҩƷID As Long, ByVal lng���˿���ID As Long, _
    ByVal int��Χ As Integer, Optional ByVal lng��ǰҩ��ID As Long) As Boolean
'���ܣ���ȡ������ҩ���������ص������б���
'������
'      int��Χ=1-����,2-סԺ(ȱʡ)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim bytDay As Byte, bln�ϰల�� As Boolean
    Dim bln��� As Boolean, i As Long
    Dim strStock As String
        
   
    'ҩƷ�������
    If mstr������ҩ�� <> "" Then
        If gblnStock Then
            strStock = " And Exists(" & _
                " Select 1 From ҩƷ���" & _
                " Where (Nvl(����,0)=0 Or Ч�� Is Null Or Ч��>Trunc(Sysdate))" & _
                " And ����=1 And ҩƷID=[3] And �ⷿID=A.ִ�п���ID" & _
                " And ��������>0 And Instr('," & mstr������ҩ�� & ",',','||�ⷿID||',')>0)"
        Else
            strStock = " And Instr('," & mstr������ҩ�� & ",',','||A.ִ�п���ID||',')>0"
        End If
    End If
          
     'ҩƷ��ϵͳָ���Ĵ���ҩ������
    If mint���� <> 3 Then
        If int��Χ = 1 Then bln�ϰల�� = Check�ϰల��() 'סԺҽ������ҩ���ϰల��
    
        If bln�ϰల�� Then
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
        End If
    End If
    strSQL = _
         " Select Distinct C.ID,C.����,C.����,C.����,B.�������" & _
         " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & IIF(bln�ϰల��, ",���Ű��� D", "") & _
         " Where A.ִ�п���ID+0=B.����ID And B.��������='��ҩ��' And B.������� IN([1],3) And B.����ID=C.ID" & _
         " And (A.������Դ is NULL Or A.������Դ=[1]) " & IIF(mint���� <> 3, " And (A.��������ID is NULL Or A.��������ID=[2])", "") & _
         " And A.�շ�ϸĿID=[3]" & strStock & _
         " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL) And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
         IIF(bln�ϰల��, " And D.����ID=C.ID And D.����=[4] And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS')", "") & _
         " Order by B.�������,C.����"
     
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ҩ��", int��Χ, lng���˿���ID, lngҩƷID, bytDay)
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        objCbo.AddItem rsTmp!���� & "-" & rsTmp!����
        objCbo.ItemData(i - 1) = Val(rsTmp!ID)
        If lng��ǰҩ��ID = Val(rsTmp!ID) Then
            Call Cbo.SetIndex(objCbo.hwnd, i - 1)
        End If
        rsTmp.MoveNext
    Next
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetSameItem(Optional ByVal lng������ĿID As Long)
'���ܣ�ɢװ��ҩ�Ƿ�����ϴ��ڶ�����ʹ��,��������ʾ�������,����ȡ��
    Dim i As Long, j As Long, strKey As String
    Dim strItem As String, strTmp As String
    Dim arrTmp As Variant
    Dim rsTmp As Recordset, strSQL As String
    Dim str��ҩIDs As String
    Dim lngRow As Long, lngCol As Long
    Dim lngҩ��ID As Long, lngҩƷID As Long
    
    If Get��ҩ��̬ <> 0 Then Exit Sub
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            strKey = vsExt.Cell(flexcpData, i, j + 2)
            If strKey <> "" And InStr(strKey, "_") > 0 Then
                If lng������ĿID = 0 Then
                    strItem = strItem & "," & strKey & "|" & i & "|" & j
                    str��ҩIDs = str��ҩIDs & "," & Val(Mid(strKey, InStr(strKey, "_") + 1))
                ElseIf Val(Mid(strKey, 1, InStr(strKey, "_") - 1)) = lng������ĿID Then
                    strItem = strItem & "," & strKey & "|" & i & "|" & j
                    str��ҩIDs = str��ҩIDs & "," & Val(Mid(strKey, InStr(strKey, "_") + 1))
                End If
            End If
        Next
    Next
    strItem = Mid(strItem, 2)
    str��ҩIDs = Mid(str��ҩIDs, 2)
    If strItem = "" Then Exit Sub
    
    strSQL = "Select b.ҩ��id,b.ҩƷid,c.����,a.��� From �շ���ĿĿ¼ a,ҩƷ��� b,������ĿĿ¼ c" & _
        " Where a.id=b.ҩƷid and c.id=b.ҩ��id and a.ID IN (Select Column_Value From Table(f_Num2list([1])))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��ҩIDs)

    arrTmp = Split(strItem, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        strKey = Split(arrTmp(i), "|")(0)
        lngRow = Val(Split(arrTmp(i), "|")(1))
        lngCol = Val(Split(arrTmp(i), "|")(2))
        
        If InStr(strKey, "_") > 0 Then
            lngҩ��ID = Val(Split(strKey, "_")(0))
            lngҩƷID = Val(Split(strKey, "_")(1))
        Else
            lngҩ��ID = Val(strKey)
            lngҩƷID = 0
        End If
        
        rsTmp.Filter = "ҩ��id=" & lngҩ��ID
        If Not rsTmp.EOF Then
            vsExt.TextMatrix(lngRow, lngCol) = rsTmp!���� & ""
        End If
        
        If rsTmp.RecordCount > 1 Then
            rsTmp.Filter = "ҩ��id=" & lngҩ��ID & " and ҩƷID=" & lngҩƷID
            If Not rsTmp.EOF Then
                vsExt.TextMatrix(lngRow, lngCol) = rsTmp!���� & "(" & rsTmp!��� & ")"
            End If
        End If
        vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
    Next
End Sub

Private Function GetҩƷ���(ByVal lngҩ��ID As Long, Optional ByVal bln��� As Boolean, Optional ByVal lng��̬ As Long) As ADODB.Recordset
'���ܣ���ȡ��ǰҩƷ��ָ����̬�Ŀ��õĹ��
'������
'      bln��� ���ò�������Ϊtrueʱ��lngҩ��ID �͵��� ҩƷid ��ʹ��
'      int��̬=0-ɢװ;1-��Ƭ��2-����
    Dim lngҩ��ID As Long, strSQL As String
    
    If mstr������ҩ�� <> "" Then
        If gblnStock Then
            If cboҩ��.ListIndex <> -1 Then lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
            If mint���� <> 3 Or lngҩ��ID <> 0 Then
                strSQL = " And Exists(Select 1 From ҩƷ��� B" & _
                    " Where (Nvl(b.����, 0) = 0 Or b.Ч�� Is Null Or b.Ч��>Trunc(Sysdate))" & _
                    " And b.����=1 And a.ҩƷID=b.ҩƷID" & IIF(lngҩ��ID = 0, "", " And b.�ⷿID=[2]") & _
                    " And b.��������>0)"
            End If
        End If
    End If
    
    strSQL = "Select a.ҩ��id, a.ҩƷid, d.���, d.����, a.����ϵ��, d.����, d.����,A.��ҩ��̬,d.�Ƿ���" & vbNewLine & _
            "From ҩƷ��� A, �շ���ĿĿ¼ D" & vbNewLine & _
            "Where a.��ҩ��̬ = [4] And a.ҩƷID = d.ID" & strSQL & vbNewLine & _
            IIF(bln���, " And a.ҩƷid = [1]", " And a.ҩ��id = [1]") & vbNewLine & _
            " And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([3],3)" & _
            " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null) Order By D.����"
    On Error GoTo errH
    Set GetҩƷ��� = zlDatabase.OpenSQLRecord(strSQL, "����б�", lngҩ��ID, lngҩ��ID, mint�������, lng��̬)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Split��ҩ���(ByVal lngҩ��ID As Long, ByVal dbl���� As Double, ByVal strKey As String, Optional ByVal strҩƷIDs As String)
'���ܣ����������󣬽��й�������ķ���(�洢��mcol���������)
'������strKey=ɢװ��ҩ��ID_ҩƷID����ɢװ:ҩƷID
'      strҩƷIDs=��ǰʹ�õ�ҩƷID�ַ���
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str������� As String, lngҩ��ID As Long, lng��̬ As Long
    Dim lngҩƷID As Long
        
    If mblnSelf = True Then Exit Sub
    If InStr(strKey, "_") > 0 Then
        lng��̬ = 0
    Else
        lng��̬ = Get��ҩ��̬
    End If
    
    If lng��̬ = 0 Then
        'ɢװ��������ʱ��ȷ�����
        str������� = mcol�������("_" & strKey)
        If str������� <> "" Then str������� = Split(str�������, ",")(0) & "," & FormatEx(dbl����, 5)
    Else
        '2.������,ҩƷid,����;ҩƷid,����;...|ʣ������
        On Error GoTo errH
        If cboҩ��.ListIndex <> -1 Then
            lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
        End If
        '���û��ҩ��������������������·���
        If lngҩ��ID = 0 Then
            '��ֹ��ѭ��
            mblnSelf = True
            Call ReSet��ҩ���
            If cboҩ��.ListIndex = -1 And mint���� <> 3 Then
                lngҩƷID = 0
                strKey = vsExt.Cell(flexcpData, 1, 2)
                If InStr(strKey, "_") > 0 Then
                    lngҩƷID = Val(Mid(strKey, InStr(strKey, "_") + 1))
                Else
                    Set rsTmp = Get��ҩ���(Val(strKey), Get��ҩ��̬)
                    If rsTmp.RecordCount > 0 Then
                        lngҩƷID = Val(rsTmp!ҩƷID & "")
                    End If
                End If
                lngҩ��ID = IIF(mlngPreRow��ҩ�� = 0, mlng��ҩ��, mlngPreRow��ҩ��)
                Call Get��ҩ��(cboҩ��, lngҩƷID, mlng���˿���id, mint�������, lngҩ��ID)
                lngҩ��ID = 0
                If cboҩ��.ListIndex <> -1 Then
                    If cboҩ��.ListCount > 0 Then cboҩ��.ListIndex = 0
                    lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
                End If
            End If
            mblnSelf = False
        End If
        strSQL = "Select Zl_Dispensechspecs([1],[2],[3],[4],[5],Null,[6],[7]) as txt From dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������", lngҩ��ID, lng��̬, dbl����, Val(txt����.Text), lngҩ��ID, mint��������, strҩƷIDs)
        str������� = "" & rsTmp!txt
    End If
    
    Call mcol�������.Remove("_" & strKey)
    mcol�������.Add str�������, "_" & strKey
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check�ϰల��() As Boolean
'���ܣ������ҩ���Ƿ��������ϰల��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Static blnҩ��Load As Boolean
    Static blnҩ��Last As Boolean
    
    '�Ƿ��а���ֻ���ȡһ��
    If blnҩ��Load Then Check�ϰల�� = blnҩ��Last: Exit Function
     
    On Error GoTo errH
    strSQL = "Select Count(B.����ID) as NUM From ��������˵�� A,���Ű��� B" & _
            " Where A.����ID=B.����ID And A.�������� ='��ҩ��'"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "Check�ϰల��")
    If Not rsTmp.EOF Then
        Check�ϰల�� = NVL(rsTmp!Num, 0) > 0
    End If
    
    blnҩ��Load = True: blnҩ��Last = Check�ϰల��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReSet��ҩ���(Optional ByVal blnReset As Boolean = True)
'���ܣ�����������ҩ�Ĺ��(����������),��������ʾ��ǰ��ҩ����������������б�
'������blnReset=false ��������
    Dim i As Long, j As Long, lngҩ��ID As Long, dbl���� As Double, lngҩƷID As Long
    Dim lng��̬ As Long, rsTmp As ADODB.Recordset
    Dim strKey As String, strTmp As String
    
    lng��̬ = Get��ҩ��̬
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            strKey = vsExt.Cell(flexcpData, i, j + 2)
            If strKey <> "" Then lngҩ��ID = Val(Split(strKey, "_")(0))
            
            If strKey <> "" Then
                strTmp = GetҩƷ����(lngҩ��ID)
                If lng��̬ = 0 Then
                    If mcol�������("_" & strKey) <> "" Then lngҩƷID = Val(Split(mcol�������("_" & strKey), ";")(0))
                    mcol�������.Remove ("_" & strKey)  '��ѡ���ID
                    If InStr(strKey, "_") > 0 And blnReset = False Then
                        Set rsTmp = GetҩƷ���(Val(Mid(strKey, InStr(strKey, "_") + 1)), True) 'ȡԭ���
                        lngҩƷID = Val(Mid(strKey, InStr(strKey, "_") + 1))
                        strTmp = vsExt.TextMatrix(i, j)
                    Else
                        Set rsTmp = GetҩƷ���(lngҩ��ID)  'ȡȱʡ���
                    End If
                    rsTmp.Filter = "ҩƷid=" & lngҩƷID    '���ԭ�����ã��򱣳ֲ���
                    If rsTmp.RecordCount = 0 Then rsTmp.Filter = ""
                    
                    If rsTmp.RecordCount > 0 Then
                        strKey = rsTmp!ҩ��ID & "_" & rsTmp!ҩƷID
                        mcol�������.Add rsTmp!ҩƷID & ",0", "_" & strKey
                    Else
                        If lngҩƷID = 0 Then
                            strKey = lngҩ��ID
                        Else
                            strKey = lngҩ��ID & "_" & lngҩƷID
                        End If
                        mcol�������.Add "", "_" & strKey
                    End If
                Else
                    mcol�������.Remove ("_" & strKey) '��ǰ������ɢװ,Key������"ҩ��ID_ҩƷID"��Ϊ"ҩƷID"������Ҫ��ɾ��
                    strKey = lngҩ��ID
                    mcol�������.Add "", "_" & strKey
                End If
                vsExt.Cell(flexcpData, i, j + 2) = strKey
                vsExt.TextMatrix(i, j) = strTmp
                vsExt.Cell(flexcpData, i, j) = strTmp
                dbl���� = Val(vsExt.TextMatrix(i, j + 1))
                If dbl���� <> 0 Then Call Split��ҩ���(lngҩ��ID, dbl����, strKey)
                
                If mcol�������("_" & strKey) = "" Or InStr(mcol�������("_" & strKey), "|") > 0 Then
                    vsExt.Cell(flexcpForeColor, i, j + 1) = vbRed
                Else
                    vsExt.Cell(flexcpForeColor, i, j + 1) = vsExt.ForeColor
                End If
            End If
        Next
    Next
    strKey = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
    
    lngҩ��ID = 0 '���lngҩ��ID
    
    If strKey <> "" Then lngҩ��ID = Val(Split(strKey, "_")(0))
    Call SetSameItem
    
    If lngҩ��ID <> 0 Then
        dbl���� = Val(vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + 1))
        Call Show��ҩ���(lngҩ��ID, dbl����)
    End If
End Sub

Private Function Get��ҩ��̬() As Long
    Dim i As Long
    
    For i = 0 To optMode.UBound
        If optMode(i).value = True Then Exit For
    Next
    Get��ҩ��̬ = i
End Function

Private Function GetҩƷ����(ByVal lngҩ��ID As Long) As String
'���ܣ�����ҩƷ����
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ���� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lngҩ��ID)
    If Not rsTmp.EOF Then GetҩƷ���� = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Show��ҩ���(ByVal lngҩ��ID As Long, ByVal dbl���� As Double, Optional ByVal lng��̬ As Long = -1)
'���ܣ����ݵ�ǰ�к��У���ʾ��������ҩ����б�
'      �����ɢװ��̬������ؿ�ѡ��Ĺ�������б�

    Dim str������� As String, arrTmp As Variant, arrValue As Variant
    Dim i As Long, strҩƷIDs As String, lngColBegin As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strMsg As String, strKey As String
    
    lngColBegin = (vsExt.Col \ 4) * 4
    cmd��̬.Visible = False
        
    With vs��ҩ���
        .Rows = .FixedRows
        .ColComboList(col���) = ""
        '���ױ༭����ʾ���
        If mint���� <> 3 Then
            '��ʾ��ҩ���
            strSQL = "Select d.����, d.���, d.����, e.���� As ҩ��, d.���㵥λ,Sum(Nvl(m.��������, 0))  As ��������" & vbNewLine & _
                    "From ҩƷ��� M, ҩƷ��� A, �շ���ĿĿ¼ D, ���ű� E" & vbNewLine & _
                    "Where m.ҩƷid = d.Id And m.ҩƷid = a.ҩƷid And m.�ⷿid = e.Id And" & vbNewLine & _
                    "      (Nvl(m.����, 0) = 0 Or m.Ч�� Is Null Or m.Ч�� > Trunc(Sysdate)) And a.ҩ��id = [1] And m.�ⷿid = [2] And" & vbNewLine & _
                    "      (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null) And d.������� In ([3],'3') And" & vbNewLine & _
                    "      (d.վ�� = '" & gstrNodeNo & "' Or d.վ�� Is Null)" & vbNewLine & _
                    "Group By e.����, d.����, d.���, d.����, d.���㵥λ" & vbNewLine & _
                    "Having Sum(Nvl(m.��������, 0)) > 0" & vbNewLine & _
                    "Order By d.����"
            lblZYStock.Caption = ""
            If cboҩ��.ListIndex <> -1 Then
                If lngҩ��ID = 0 Then Exit Sub
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҩ��ID, cboҩ��.ItemData(cboҩ��.ListIndex), mint������� & "")
                If rsTmp.RecordCount > 0 Then
                    If InStr(mstrPrivs, "��ʾҩƷ���") = 0 Then
                        lblZYStock.Caption = "��棺�С�"
                    Else
                        Do While Not rsTmp.EOF
                            lblZYStock.Caption = lblZYStock.Caption & IIF(lblZYStock.Caption = "", "��棺", "    ") & rsTmp!��� & ":" & rsTmp!�������� & rsTmp!���㵥λ
                            rsTmp.MoveNext
                        Loop
                    End If
                Else
                    lblZYStock.Caption = "��棺�ޡ�"
                End If
            End If
        End If
        If lng��̬ = -1 Then lng��̬ = Get��ҩ��̬
        
        If dbl���� = 0 And lng��̬ <> 0 Or cboҩ��.ListIndex = -1 And mint���� <> 3 Then Exit Sub
        vsExt.Cell(flexcpForeColor, vsExt.Row, lngColBegin + 1) = vsExt.ForeColor
        
        strKey = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
        If strKey = "" And lngҩ��ID <> 0 Then strKey = lngҩ��ID
        If strKey <> "" Then str������� = Trim(mcol�������("_" & strKey))

        .Redraw = False
        If str������� = "" Then
            .Rows = .FixedRows + 1
            '���ܷ���ʱ���ؿ�,����:����Ϊ6��10�������,3�˵ķ���
            .MergeCells = flexMergeRestrictRows
            If lng��̬ = 0 Then
                strMsg = "��ҩƷû�п��õ�ɢװ��̬����ѡ������ҩƷ����̬��"
            Else
                strMsg = "�޷����������������ù����䣬�����������"
            End If
            
            .TextMatrix(.Rows - 1, col���) = strMsg
            .TextMatrix(.Rows - 1, col����) = strMsg
            .MergeRow(.Rows - 1) = True
            .TextMatrix(.Rows - 1, col����) = FormatEx(dbl����, 5)
            .Cell(flexcpForeColor, .Rows - 1, col����) = vbRed
            
            vsExt.Cell(flexcpForeColor, vsExt.Row, lngColBegin + 1) = vbRed
            If lng��̬ <> 0 Then
                'ɢװ��̬����Ƭ���Ҳ����ѡ����66074��
                Set rsTmp = GetҩƷ���(lngҩ��ID, , lng��̬)
                If rsTmp.RecordCount > 1 Then
                    strҩƷIDs = ""
                    For i = 1 To rsTmp.RecordCount
                        strҩƷIDs = strҩƷIDs & "|#" & rsTmp!ҩƷID & ";" & rsTmp!���� & "-" & rsTmp!���� & IIF(Not IsNull(rsTmp!���), "(" & rsTmp!��� & ")", "")
                        rsTmp.MoveNext
                    Next
                    .ColComboList(col���) = Mid(strҩƷIDs, 2)
                    rsTmp.MoveFirst
                    .RowData(.FixedRows) = rsTmp   'ֻ��һ��
                    .Cell(flexcpBackColor, .FixedRows, col���, .Rows - 1, col���) = &HF0F4E4
                End If
            End If
        Else
            arrTmp = Split(Split(str�������, "|")(0), ";")
            
            If InStr(str�������, "|") > 0 Then
                vs��ҩ���.Rows = vs��ҩ���.FixedRows + UBound(arrTmp) + 2
            Else
                vs��ҩ���.Rows = vs��ҩ���.FixedRows + UBound(arrTmp) + 1
            End If
            For i = 0 To UBound(arrTmp)
                arrValue = Split(arrTmp(i), ",")
                strҩƷIDs = strҩƷIDs & "," & Val(arrValue(0))
                .TextMatrix(.FixedRows + i, colҩƷID) = Val(arrValue(0))  '���ID
                .TextMatrix(.FixedRows + i, col����) = FormatEx(arrValue(1), 5)    '����
            Next
            strҩƷIDs = Mid(strҩƷIDs, 2)
            
            On Error GoTo errH
            '�������п���(�п��)��ɢװ����Ա����ѡ�������Ĺ��
             'ɢװ��̬����Ƭ���Ҳ����ѡ����66074��
            Set rsTmp = GetҩƷ���(lngҩ��ID, , lng��̬)
            For i = .FixedRows To .Rows - 1
                If InStr(str�������, "|") > 0 And i = .Rows - 1 Then
                '���һ����ʾδ��������
                    .MergeCells = flexMergeRestrictRows
                    strMsg = "�޷����������������ù����䣬�����������"
                    .TextMatrix(i, col���) = strMsg
                    .TextMatrix(i, col����) = strMsg
                    .MergeRow(i) = True
                    .Cell(flexcpForeColor, i, col����) = vbRed
                    .TextMatrix(i, col����) = FormatEx(Split(str�������, "|")(1), 5)
                    vsExt.Cell(flexcpForeColor, vsExt.Row, lngColBegin + 1) = vbRed
                Else
                    rsTmp.Filter = "ҩƷID = " & CStr(.TextMatrix(i, colҩƷID))
                    If rsTmp.RecordCount = 0 Then 'ɢװ����治��ʱ�������棩
                        strMsg = "��ǰҩ����治�㣬����û��ɢװ���"
                        .TextMatrix(.Rows - 1, col���) = strMsg
                        .TextMatrix(.Rows - 1, col����) = strMsg
                        .MergeRow(.Rows - 1) = True
                    
                        .Cell(flexcpForeColor, i, col����) = vbRed
                        vsExt.Cell(flexcpForeColor, vsExt.Row, lngColBegin + 1) = vbRed
                    Else
                        .TextMatrix(i, col���) = "" & rsTmp!���
                        .Cell(flexcpData, i, col���) = "" & rsTmp!��� '����ɢװ���ȡ������ѡ��ʱ�ָ�
                        .TextMatrix(i, col����) = "" & rsTmp!����
                        .TextMatrix(i, col������λ) = vsExt.TextMatrix(vsExt.Row, lngColBegin + 2)
                        
                        '��¼�ۼ۵���
                        If NVL(rsTmp!�Ƿ���, 0) = 0 Then
                            .TextMatrix(i, col����) = Format(CalcPrice(Val("" & rsTmp!ҩƷID), , , True, , , mstrҩƷ�۸�ȼ�), gstrDecPrice)
                        Else 'ʱ��
                            .TextMatrix(i, col����) = Format(CalcDrugPrice(Val("" & rsTmp!ҩƷID), cboҩ��.ItemData(cboҩ��.ListIndex), Val(.TextMatrix(i, col����)), , True, 1, mstrҩƷ�۸�ȼ�), gstrDecPrice)
                        End If
                    End If
                End If
            Next
            
            'ɢװ��̬����Ƭ���Ҳ����ѡ����66074��
            rsTmp.Filter = ""
            If rsTmp.RecordCount > 1 Then
                strҩƷIDs = ""
                For i = 1 To rsTmp.RecordCount
                    strҩƷIDs = strҩƷIDs & "|#" & rsTmp!ҩƷID & ";" & rsTmp!���� & "-" & rsTmp!���� & IIF(Not IsNull(rsTmp!���), "(" & rsTmp!��� & ")", "")
                    rsTmp.MoveNext
                Next
                .ColComboList(col���) = Mid(strҩƷIDs, 2)
                rsTmp.MoveFirst
                .RowData(.FixedRows) = rsTmp   'ֻ��һ��
                .Cell(flexcpBackColor, .FixedRows, col���, .Rows - 1, col���) = &HF0F4E4
            End If
        End If
        
        If lng��̬ <> 0 Then
            '��ɢװ��̬��δ������ʱ������Ϊɢװ
            If str������� = "" Or InStr(str�������, "|") > 0 Then
                Set rsTmp = GetҩƷ���(lngҩ��ID)
                If rsTmp.RecordCount > 0 Then
                    strMsg = "�޷����������������ù����䣬��������������ɢװ��"
                    .TextMatrix(.Rows - 1, col���) = strMsg
                    .TextMatrix(.Rows - 1, col����) = strMsg
                    .MergeRow(.Rows - 1) = True
                    
                    .Select .Rows - 1, col������λ
                    cmd��̬.Visible = True
                    cmd��̬.Tag = rsTmp!ҩƷID  'ȱʡ���
                    cmd��̬.Caption = "ɢװ(&D)"
                    cmd��̬.Top = vs��ҩ���.CellTop
                    cmd��̬.Left = vs��ҩ���.CellLeft
                    cmd��̬.Width = vs��ҩ���.CellWidth
                    cmd��̬.Height = vs��ҩ���.CellHeight
                End If
            End If
        End If
        
        .Redraw = True
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, 2)

    mlngHwnd = 0
    mblnҽ�� = False
    mint���� = 0
    mintӤ�� = 0
    mlng����ID = 0
    mlng���˿���id = 0
    mvar����ID = Empty
    mstr�Ա� = ""
    mint���� = 0
    mbytUseType = 0
    mint��Ч = 0
    mint������� = 0
    mint�������� = 0
    mlng��ĿID = 0
    mlngҩƷID = 0
    mlngPreRow��ҩ�� = 0
    Set mclsInsure = Nothing
    Set mcol������� = Nothing
    Set mfrmParent = Nothing
End Sub

Private Sub optMode_Click(Index As Integer)
    Dim lngҩ��ID As Long, lngҩ��ID As Long, str������� As String, lngҩƷID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strKey As String, bln����ҩ�� As Boolean
    
    If Not Me.Visible Then Exit Sub
    strKey = vsExt.Cell(flexcpData, vsExt.FixedRows, vsExt.FixedCols + 2)
    If strKey <> "" Then lngҩ��ID = Val(Split(strKey, "_")(0))
    
    If lngҩ��ID <> 0 And gblnStock Then     'ָ��ҩ��ʱ�޶����ʱ����̬�ı�󣬵�һζҩ��ȱʡ�����ܱ��ˣ�����ҩ���ͱ���
        
        str������� = mcol�������("_" & strKey)
        If str������� <> "" Then
            Set rsTmp = Get��ҩ���(lngҩ��ID, Index)
            If rsTmp.RecordCount > 0 Then
                lngҩƷID = Val(Split(str�������, ",")(0))
                If lngҩƷID <> Val(rsTmp!ҩƷID) Then
                    rsTmp.Filter = "ҩƷID=" & lngҩƷID
                    If rsTmp.EOF Then
                        rsTmp.Filter = 0
                        lngҩƷID = Val(rsTmp!ҩƷID)
                    End If
                    If cboҩ��.ListIndex = -1 Then
                        lngҩ��ID = IIF(mlngPreRow��ҩ�� = 0, mlng��ҩ��, mlngPreRow��ҩ��)
                    Else
                        lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
                    End If
                    'ȱʡҩ��Ҳ���ܱ��ˣ�Ҫ���·������������
                    Call Get��ҩ��(cboҩ��, lngҩƷID, mlng���˿���id, mint�������, lngҩ��ID)
                    bln����ҩ�� = True
                    If cboҩ��.ListIndex = -1 And cboҩ��.ListCount > 0 Then
                        Call Cbo.SetIndex(cboҩ��.hwnd, 0)
                    End If
                End If
            End If
        End If
    End If
    
    '��̬���ˣ�Ҫ���·����������
    Call ReSet��ҩ���
    If (Not bln����ҩ�� And mcol�������.Count = 1 Or cboҩ��.ListIndex = -1) And mint���� <> 3 Then
        lngҩƷID = 0
        strKey = vsExt.Cell(flexcpData, 1, 2)
        If InStr(strKey, "_") > 0 Then
            lngҩƷID = Val(Mid(strKey, InStr(strKey, "_") + 1))
        Else
            Set rsTmp = Get��ҩ���(Val(strKey), Get��ҩ��̬)
            If rsTmp.RecordCount > 0 Then
                lngҩƷID = Val(rsTmp!ҩƷID & "")
            End If
        End If
        lngҩ��ID = IIF(mlngPreRow��ҩ�� = 0, mlng��ҩ��, mlngPreRow��ҩ��)
        Call Get��ҩ��(cboҩ��, lngҩƷID, mlng���˿���id, mint�������, lngҩ��ID)
        If cboҩ��.ListIndex = -1 And cboҩ��.ListCount > 0 Then
            Call Cbo.SetIndex(cboҩ��.hwnd, 0)
        End If
    End If
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        vsExt.SetFocus
    End If
End Sub

Private Sub vsExt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ��ǻس�ȷ�����༭�Ĵ���(����Text:=EditText,��ValidateEdit�¼��л�û��)
    Dim strPrivs As String, i As Long
    Dim strKey As String, lngҩ��ID As Long
    
    If Not mblnReturn Then
        If Col Mod 4 = 0 Then '��ҩ
            vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
        ElseIf Col Mod 4 = 1 Then '��ζ����
            If Not IsNumeric(vsExt.TextMatrix(Row, Col)) _
                Or Val(vsExt.TextMatrix(Row, Col)) <= 0 _
                Or Val(vsExt.TextMatrix(Row, Col)) > LONG_MAX Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
            Else
                'С������(���ײ���)
                If mint���� <> 3 Then
                    If Val(vsExt.TextMatrix(Row, Col)) <> Int(Val(vsExt.TextMatrix(Row, Col))) Then
                        If mint�������� = 1 Then
                            strPrivs = GetInsidePrivs(pm����ҽ���´�)
                        ElseIf mint�������� = 2 Then
                            strPrivs = GetInsidePrivs(pmסԺҽ���´�)
                        End If
                        If InStr(strPrivs, "ҩƷС������") = 0 Then
                            vsExt.TextMatrix(Row, Col) = IntEx(Val(vsExt.TextMatrix(Row, Col)))
                        End If
                    End If
                End If
                vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
                
                strKey = vsExt.Cell(flexcpData, Row, (Col \ 4) * 4 + 2)
                lngҩ��ID = Val(Split(strKey, "_")(0))
                Call Split��ҩ���(lngҩ��ID, Val(vsExt.TextMatrix(Row, Col)), strKey)
                Call Show��ҩ���(lngҩ��ID, Val(vsExt.TextMatrix(Row, Col)))
            End If
        ElseIf Col Mod 4 = 3 Then '��ע
            If zlCommFun.ActualLen(vsExt.TextMatrix(Row, Col)) > 100 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
            Else
                vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
            End If
        End If
    End If
    'ȷ��ζ��
    Call RefreshWeiNum
End Sub

Private Sub vsExt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'����:��ʾѡ��ť,����֤��ǰ��Ԫ��ɼ�
    Dim strKey As String, lngҩ��ID As Long
    
    If mblnChangeSel = True Then Exit Sub
    '��֤��ǰ��Ԫ��ɼ�
    If NewRow >= vsExt.FixedRows And NewRow <= vsExt.Rows - 1 Then
        If vsExt.LeftCol >= vsExt.FixedCols And vsExt.LeftCol <= vsExt.Cols - 1 Then
            Call vsExt.ShowCell(NewRow, vsExt.LeftCol)
        End If
    End If

    '��ʾ��ҩ���
    If Me.Visible Then
        If OldRow <> NewRow Or (OldCol \ 4) <> (NewCol \ 4) Then   '���л򻻵���һҩƷ��
            lblZYStock.Caption = ""
            strKey = vsExt.Cell(flexcpData, NewRow, (NewCol \ 4) * 4 + 2)
            If strKey <> "" Then
                lngҩ��ID = Val(Split(strKey, "_")(0))
                Call Show��ҩ���(lngҩ��ID, Val(vsExt.TextMatrix(NewRow, (NewCol \ 4) * 4 + 1)))
            Else
                vs��ҩ���.Rows = vs��ҩ���.FixedRows
                cmd��̬.Visible = False
            End If
        End If

        If NewCol = (NewCol \ 4) * 4 And vsExt.TextMatrix(NewRow, NewCol) <> "" Then
            strKey = "�С�" & vsExt.TextMatrix(NewRow, (NewCol \ 4) * 4) & "��"
        Else
            strKey = "��"
        End If
        vsExt.ToolTipText = "��" & NewRow & strKey
    End If
End Sub

Private Sub vsExt_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
   
    If Button = 1 Then
        '��λ����겻�ɽ���
        If vsExt.MouseCol Mod 4 = 2 Then Cancel = True
        If mbytUseType = 3 And (vsExt.MouseCol >= 4 Or vsExt.MouseRow >= 2) Then Cancel = True  'mbytUseType = 3ʱ����ҩ¼��
    End If
End Sub

Private Sub vsExt_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)

    If NewCol Mod 4 = 2 Then '��λ�а������ɽ���
        Cancel = True
        If OldCol > NewCol Then '�����ƶ�ʱ����
            vsExt.Col = NewCol - 1
        Else
            vsExt.Col = NewCol + 1
        End If
        vsExt.Row = NewRow
    End If
    
    If mbytUseType = 3 And (NewCol >= 4 Or NewRow >= 2) Then Cancel = True  'mbytUseType = 3ʱ����ҩ¼��
End Sub

Private Sub vsExt_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If cmd.Visible Then cmd.Visible = False
End Sub

Private Sub vsExt_GotFocus()
    Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) 'ʹ��ť�ɼ�
End Sub

Private Sub vsExt_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�ɾ��������
    Dim i As Long, j As Long, k As Long, g As Long
    Dim intRow As Integer        '��Ч��
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strKey As String, lngҩƷID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim rsTmp As Recordset
    Dim lngҩ��ID As Long
    
    If KeyCode = vbKeyDelete Then
        If MsgBox("Ҫɾ��""" & vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '�����ǰζҩ��Ϣ
        
        strKey = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
        If strKey <> "" Then
            lngҩƷID = Val(Split(strKey, "_")(0))
            If InStr(mcol�������("_" & strKey), "|") > 0 Then
                vsExt.Select vsExt.Row, (vsExt.Col \ 4) * 4 + 1 '�����У�֮ǰ�Ǻ�ɫ
                vsExt.CellForeColor = vsExt.ForeColorSel
            End If
            mcol�������.Remove ("_" & strKey)
            If mint���� = 3 Then vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2) = ""
            Call Show��ҩ���(0, 0)
        End If
        
        For i = 0 To 3
            vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + i) = ""
            vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + i) = Empty
        Next
        '�����������ǰ��
        For i = vsExt.Row To vsExt.Rows - 1
            For j = 0 To vsExt.Cols - 1 Step 4
                If Not (i = vsExt.Row And j <= (vsExt.Col \ 4) * 4) Then
                    For k = 0 To 3
                        If j = 0 Then
                            vsExt.TextMatrix(i - 1, vsExt.Cols - (4 - k)) = vsExt.TextMatrix(i, j + k)
                            vsExt.Cell(flexcpData, i - 1, vsExt.Cols - (4 - k)) = vsExt.Cell(flexcpData, i, j + k)
                            vsExt.Cell(flexcpForeColor, i - 1, vsExt.Cols - (4 - k)) = vsExt.Cell(flexcpForeColor, i, j + k)
                        Else
                            vsExt.TextMatrix(i, j + k - 4) = vsExt.TextMatrix(i, j + k)
                            vsExt.Cell(flexcpData, i, j + k - 4) = vsExt.Cell(flexcpData, i, j + k)
                            vsExt.Cell(flexcpForeColor, i, j + k - 4) = vsExt.Cell(flexcpForeColor, i, j + k)
                        End If
                        vsExt.TextMatrix(i, j + k) = ""
                        vsExt.Cell(flexcpData, i, j + k) = Empty
                    Next
                End If
            Next
        Next
        'ɾ������Ŀ���(���ٱ������������ʾ������7)
        If vsExt.Rows > 7 Then
            For i = vsExt.Rows - 1 To 7 Step -1
                If Val(vsExt.Cell(flexcpData, i - 1, 2)) = 0 Then
                    vsExt.RemoveItem i
                End If
            Next
        End If
        
        If optMode(1).Enabled = False Then
            If Check��ͬɢװ��ҩ Then
                optMode(1).Enabled = False: optMode(2).Enabled = False
            Else
                optMode(1).Enabled = True: optMode(2).Enabled = True
            End If
        End If
        If vsExt.Row >= vsExt.FixedRows Then Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col)
        Call vsExt.ShowCell(vsExt.Row, vsExt.Col)
        Call RefreshWeiNum
        '����ҩ��
        If InStr(strKey, "_") > 0 Then
            Call SetSameItem(Val(Mid(strKey, 1, InStr(strKey, "_") - 1)))
        End If
        
        '���ɾ�����ǵ�һλ��ҩ�����¼���ҩ��(���ױ༭����)
        If mint���� <> 3 Then
            If cboҩ��.ListIndex = -1 And vsExt.Col \ 4 = 0 Then
                lngҩƷID = 0
                strKey = vsExt.Cell(flexcpData, 1, 2)
                If InStr(strKey, "_") > 0 Then
                    lngҩƷID = Val(Mid(strKey, InStr(strKey, "_") + 1))
                Else
                    Set rsTmp = Get��ҩ���(Val(strKey), Get��ҩ��̬)
                    If rsTmp.RecordCount > 0 Then
                        lngҩƷID = Val(rsTmp!ҩƷID & "")
                    End If
                End If
                lngҩ��ID = IIF(mlngPreRow��ҩ�� = 0, mlng��ҩ��, mlngPreRow��ҩ��)
                Call Get��ҩ��(cboҩ��, lngҩƷID, mlng���˿���id, mint�������, lngҩ��ID)
                Call ReSet��ҩ���(False)
            End If
        End If
    ElseIf KeyCode = vbKeyInsert Then
        If Val(vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)) <> 0 Then

            '����Ƿ��в������û����д
            If CheckIsNullZY(lngRow, lngCol) Then
                MsgBox "�����������ٲ������", vbInformation, Me.Caption
                Call vsExt.Select(lngRow, lngCol)
                Exit Sub
            End If
            '�ҵ���Ч��
            intRow = -1
            For i = 0 To vsExt.Rows - 1
                For j = 0 To vsExt.Cols - 1 Step 4
                    If vsExt.TextMatrix(i, j) = "" Then
                        intRow = i
                        Exit For
                    End If
                Next
                If intRow <> -1 Then Exit For
            Next
            '���û���ҵ���Ч����,˵���б��Ѿ����ˣ������һ��
            If intRow = -1 Then
                intRow = vsExt.Rows - 1
                vsExt.Rows = vsExt.Rows + 1
            End If
            '��������������
            For i = intRow To vsExt.Row Step -1
                For j = vsExt.Cols - 1 To 0 Step -4
                    If vsExt.TextMatrix(i, j - 3) <> "" Then
                        If Not (i = vsExt.Row And j <= (vsExt.Col \ 4) * 4) Then
                            For k = 0 To 3
                                If j = vsExt.Cols - 1 Then
                                    vsExt.TextMatrix(i + 1, k) = vsExt.TextMatrix(i, j + (k - 3))
                                    vsExt.Cell(flexcpData, i + 1, k) = vsExt.Cell(flexcpData, i, j + (k - 3))
                                    vsExt.Cell(flexcpForeColor, i + 1, k) = vsExt.Cell(flexcpForeColor, i, j + (k - 3))
                                Else
                                    vsExt.TextMatrix(i, j + k + 1) = vsExt.TextMatrix(i, j + (k - 3))
                                    vsExt.Cell(flexcpData, i, j + k + 1) = vsExt.Cell(flexcpData, i, j + (k - 3))
                                    vsExt.Cell(flexcpForeColor, i, j + k + 1) = vsExt.Cell(flexcpForeColor, i, j + (k - 3))
                                End If
                                vsExt.TextMatrix(i, j + (k - 3)) = ""
                                vsExt.Cell(flexcpData, i, j + (k - 3)) = Empty
                            Next
                        End If
                    End If
                Next
            Next
            For i = 0 To 3
                vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + i) = ""
                vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + i) = Empty
                vsExt.Cell(flexcpForeColor, vsExt.Row, (vsExt.Col \ 4) * 4 + i) = vsExt.ForeColor
            Next
            Call vsExt.ShowCell(vsExt.Row, vsExt.Col)
            Call Show��ҩ���(0, 0)
        End If
        Call RefreshWeiNum
    End If
End Sub

Private Sub vsExt_KeyPress(KeyAscii As Integer)
'���ܣ��Ǳ༭״̬ʱ���Զ��ƶ���Ԫ��
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        '��λ����һӦ���뵥Ԫ��
        If Val(vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)) = 0 Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        Else
            Call EnterNextCell(vsExt.Row, vsExt.Col)
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        cmd_Click 'ѡ��ζ�в�ҩ���ע
    End If
End Sub

Private Sub vsExt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'���ܣ���������ȷ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strҩƷ As String
    Dim strStock As String, blnCancel As Boolean, i As Long
    Dim vPoint As PointAPI, strLike As String
    Dim strSamples As String, strPrivs As String
    Dim strKey As String, lngҩ��ID As Long
    
    If KeyAscii = 13 Then
        mblnReturn = True '����ǰ��س�ȷ�ϱ༭
        KeyAscii = 0
        
        '�Ż�
        strLike = gstrLike
        If Len(vsExt.EditText) < 2 Then strLike = ""
        
        On Error GoTo errH
        '��ȡ�س���,�����MsgboxʹEdit���㶪ʧ,�����ɱ༭,�����ἤ��AfterEdit�¼�
        If Col Mod 4 = 0 Then '��ҩ
            Call Set��ҩInput(True)
            
            strKey = vsExt.Cell(flexcpData, Row, (Col \ 4) * 4 + 2)
            If strKey <> "" Then lngҩ��ID = Val(Split(strKey, "_")(0))
            Call Show��ҩ���(lngҩ��ID, Val(vsExt.TextMatrix(Row, Col)))
            Exit Sub
        ElseIf Col Mod 4 = 1 Then '����
            If Not IsNumeric(vsExt.EditText) Or Val(vsExt.EditText) <= 0 Or Val(vsExt.EditText) > LONG_MAX Then
                MsgBox "��ζ����������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            'С������(���ײ���)
            If mint���� <> 3 Then
                If Val(vsExt.EditText) <> Int(Val(vsExt.EditText)) Then
                    If mint�������� = 1 Then
                        strPrivs = GetInsidePrivs(pm����ҽ���´�)
                    ElseIf mint�������� = 2 Then
                        strPrivs = GetInsidePrivs(pmסԺҽ���´�)
                    End If
                    If InStr(strPrivs, "ҩƷС������") = 0 Then
                        vsExt.EditText = IntEx(Val(vsExt.EditText))
                    End If
                End If
            End If
            vsExt.TextMatrix(Row, Col) = vsExt.EditText
            
            strKey = vsExt.Cell(flexcpData, Row, (Col \ 4) * 4 + 2)
            lngҩ��ID = Val(Split(strKey, "_")(0))
            
            Call Split��ҩ���(lngҩ��ID, Val(vsExt.TextMatrix(Row, Col)), strKey)
            Call Show��ҩ���(lngҩ��ID, Val(vsExt.TextMatrix(Row, Col)))
        ElseIf Col Mod 4 = 3 Then '��ע
            If vsExt.EditText <> "" Then
                strSQL = "Select Rownum as ID,����,����,���� From ��ҩ�����ע" & _
                    " Where Upper(����) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2]" & _
                    " Order by ����"
                vPoint = zlControl.GetCoordPos(vsExt.hwnd, vsExt.CellLeft, vsExt.CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ע", False, "", "", False, False, True, vPoint.X, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                    UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%")
            End If
            If rsTmp Is Nothing Then
                If blnCancel Then
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '��ƥ�䵱��ֱ������
                If zlCommFun.ActualLen(vsExt.EditText) > 100 Then
                    MsgBox "��ע�������ݹ��������ֻ���� 50 �����ֻ� 100 ���ַ���", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                vsExt.TextMatrix(Row, Col) = vsExt.EditText
            Else
                vsExt.EditText = rsTmp!���� 'ֱ������ƥ��ʱ��Ҫ
                vsExt.TextMatrix(Row, Col) = rsTmp!����
            End If
        End If
        vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
        Call EnterNextCell(Row, Col)
    Else
        '��ζ����ֻ������������
        If Col Mod 4 = 1 Then
            If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsExt_LostFocus()
    If Not ActiveControl Is cmd Then cmd.Visible = False
End Sub

Private Sub vsExt_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsExt.EditSelStart = 0
    vsExt.EditSelLength = zlCommFun.ActualLen(vsExt.EditText)
End Sub

Private Sub vsExt_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ�����ĳЩ�в�����༭(���¼�����BeforeEdit,��EditText��ֵ֮ǰ)
    mblnReturn = False
        
    '������������
    If Not CellCanEdit(Row, Col) Then Cancel = True
    
    If Col Mod 4 = 1 Then
        vsExt.EditMaxLength = 8
    Else
        vsExt.EditMaxLength = 0
    End If
End Sub

Private Sub vs��ҩ���_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim lng��̬ As Long
    Dim strKey As String
    Dim strTmp As String
    Dim i As Long
    Dim str������� As String
    Dim strҩƷIDs As String
    Dim dbl���� As Double
    
    With vs��ҩ���
        If .Col = col��� Then
            If .ComboData = "" Then
            'û��ѡ��ʱ�ƿ�����
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
            Else
            'ɢװ��ҩ��ѡ����֮��
                Set rsTmp = .RowData(.FixedRows)
                rsTmp.Filter = "ҩƷID = " & CLng(.ComboData)
                
                lng��̬ = Get��ҩ��̬
                If lng��̬ = 0 Then
                    strKey = rsTmp!ҩ��ID & "_" & rsTmp!ҩƷID
                Else
                    strKey = rsTmp!ҩ��ID
                End If
                
                If lng��̬ = 0 Then
                    On Error Resume Next
                    strTmp = mcol�������("_" & strKey)
                    If err.Number = 0 Then
                        MsgBox "��ͬ����ҩƷ�Ѵ��ڣ���ѡ���������", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Exit Sub
                    Else
                        err.Clear
                    End If
                    On Error GoTo 0
                Else
                    For i = 1 To .Rows - 1
                        If i <> Row Then
                            If Val(.TextMatrix(i, colҩƷID)) = rsTmp!ҩƷID Then
                                 MsgBox "��ͬ����ҩƷ�Ѵ��ڣ���ѡ���������", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Exit Sub
                            End If
                        End If
                    Next
                End If
                
                .Cell(flexcpData, Row, Col) = CLng(.ComboData)
            
                strTmp = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
                mcol�������.Remove "_" & strTmp
                vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2) = strKey
                .TextMatrix(Row, col���) = "" & rsTmp!���
                .Cell(flexcpData, Row, Col) = "" & rsTmp!���   '���ڻָ�
                .TextMatrix(Row, col����) = "" & rsTmp!����
                .TextMatrix(Row, colҩƷID) = Val(rsTmp!ҩƷID & "")
                If NVL(rsTmp!�Ƿ���, 0) = 0 Then
                    .TextMatrix(Row, col����) = Format(CalcPrice(Val("" & rsTmp!ҩƷID), , , True, , , mstrҩƷ�۸�ȼ�), gstrDecPrice)
                Else 'ʱ��
                    .TextMatrix(Row, col����) = Format(CalcDrugPrice(Val("" & rsTmp!ҩƷID), cboҩ��.ItemData(cboҩ��.ListIndex), Val(.TextMatrix(Row, col����)), , True, 1, mstrҩƷ�۸�ȼ�), gstrDecPrice)
                End If
                
                For i = 1 To .Rows - 1
                    str������� = str������� & ";" & .TextMatrix(i, colҩƷID) & "," & .TextMatrix(i, col����)
                    strҩƷIDs = strҩƷIDs & "," & .TextMatrix(i, colҩƷID)
                    dbl���� = dbl���� + Val(.TextMatrix(i, col����))
                Next
                str������� = Mid(str�������, 2)
                strҩƷIDs = Mid(strҩƷIDs, 2)
                mcol�������.Add str�������, "_" & strKey
                If lng��̬ = 1 Or lng��̬ = 2 Then
                    '��Ƭ�������޸��˹�����ݵ�ǰ������·��䡣
                    Call Split��ҩ���(rsTmp!ҩ��ID, dbl����, strKey, strҩƷIDs)
                    vsExt_AfterRowColChange 0, 0, vsExt.Row, vsExt.Col
                End If
                
                '������λ����
                
                '�����������ƣ���Ϊ�޸Ĺ�񲻻�Ӱ�����ƣ�ֻ��ͬһ��ҩƷʹ���˶�����ģ�����ʾ����
            End If
        End If
    End With
End Sub

Private Sub vs��ҩ���_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewCol = -1 Or NewRow = -1 Then Exit Sub
    If NewCol = col��� And vs��ҩ���.ColComboList(col���) <> "" Then
        vs��ҩ���.FocusRect = flexFocusSolid
    Else
        vs��ҩ���.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vs��ҩ���_ChangeEdit()
    'Call vs��ҩ���_AfterEdit(vs��ҩ���.Row, vs��ҩ���.Col)
End Sub

Private Sub vs��ҩ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vs��ҩ���_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vs��ҩ���.ComboIndex <> -1 Then
            Call vs��ҩ���_KeyPress(13)
        End If
    End If
End Sub

Private Sub vs��ҩ���_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col��� And vs��ҩ���.ColComboList(col���) <> "") Then
        Cancel = True
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    '�������
    If Not IsNumeric(txt����.Text) Then
        MsgBox "������һ����Ч����ֵ��", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txt����)
        Cancel = True: Exit Sub
    End If
    If Val(txt����.Text) <> Int(txt����.Text) Then
        MsgBox "��ҩ����Ӧ����������ֵ��", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txt����)
        Cancel = True: Exit Sub
    End If
    If Val(txt����.Text) = 0 Then
        MsgBox "������һ������ĸ�����", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txt����)
        Cancel = True: Exit Sub
    End If
    
    If Val(txt����.Tag) <> Val(txt����.Text) Then
        txt����.Tag = Val(txt����.Text)
    End If
End Sub

Private Function Set��ҩInput(ByVal blnInputKey As Boolean) As Boolean
'���ܣ������������ݻ�*����ѡ��ť������������ѡ��ļ�¼��
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    Dim strMain As String, strPerson As String, strPrivs As String, strCode As String, strLike As String
    Dim str��� As String, str���� As String, strInput As String, str���� As String, str�洢�ⷿ As String
    Dim lng��̬ As Long, lngҩ��ID As Long, lngҩƷID As Long, dbl���� As Double
    Dim int�Ա� As Integer, strStock As String
    Dim vPoint As PointAPI, blnCancel As Boolean
    Dim strKey As String
    Dim blnFirst As Boolean '�б��еĵ�һζҩ
    Dim rs��� As ADODB.Recordset
    Dim strTsPrivs As String
    
    On Error GoTo errH
    
    If mint���� <> 3 Then
        If mstr�Ա� Like "*��*" Then
            int�Ա� = 1
        ElseIf mstr�Ա� Like "*Ů*" Then
            int�Ա� = 2
        End If
        
        '�б��еĵ�һζҩ
        If vsExt.Row = vsExt.FixedRows And vsExt.Col = vsExt.FixedCols Then
            If vsExt.TextMatrix(vsExt.FixedRows, vsExt.Col + 4) = "" Then blnFirst = True
        End If
    End If
    
    lng��̬ = Get��ҩ��̬
    If cboҩ��.ListIndex <> -1 Then lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
    
    '���,��ҩ��δָ��ʱ,����������¼
    If lngҩ��ID <> 0 Then
        strStock = _
            "Select ҩƷID,Sum(Nvl(��������,0)) as ��� From ҩƷ���" & _
            " Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))" & _
            " And ���� = 1 And �ⷿID=" & lngҩ��ID & _
            " Group by ҩƷID" & _
            " Having Sum(Nvl(��������,0))<>0"
            '��ҩ��һ��ֻ��һ���������˴��ɲ��ð󶨱���
    Else
        strStock = "Select NULL as ҩƷID,NULL as ��� From Dual"
    End If
        
    If mint���� <> 3 Then
    '�洢�ⷿ�������ò�����Դ
        If lng��̬ = 0 Then
            str�洢�ⷿ = " And Exists(select 1 from �շ�ִ�п��� f Where f.�շ�ϸĿid=d.id and (f.��������id Is Null Or f.��������id=[8]) And f.ִ�п���id=" & lngҩ��ID & ")"
        Else
            str�洢�ⷿ = " And Exists(select 1 from ����ִ�п��� f Where f.������Ŀid=a.id and (f.��������id Is Null Or f.��������id=[8]) And f.ִ�п���id=" & lngҩ��ID & ")"
        End If
        If blnFirst Then str�洢�ⷿ = ""
    End If
    
    '����ҩƷȨ��
    str���� = ""
    strPrivs = GetInsidePrivs(IIF(mint�������� = 1, pm����ҽ���´�, pmסԺҽ���´�))
    strTsPrivs = GetTsPrivs(IIF(mint�������� = 1, pm����ҽ���´�, pmסԺҽ���´�))
    
    If mint���� <> 3 Then
        If InStr(strTsPrivs, "�´�����ҩ��") = 0 Then
            str���� = str���� & " And E.�������<>'����ҩ'"
        End If
        If InStr(strTsPrivs, "�´ﶾ��ҩ��") = 0 Then
            str���� = str���� & " And E.�������<>'����ҩ'"
        End If
        If InStr(strTsPrivs, "�´ﾫ��ҩ��") = 0 Then
            str���� = str���� & " And E.������� Not IN('����I��')"
        End If
        If InStr(strTsPrivs, "�´����ҩ��") = 0 Then
            str���� = str���� & " And E.��ֵ���� Not IN('����','����')"
        End If
    End If
    str���� = " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) And A.������� IN([1],3) And A.���='7' And Nvl(A.ִ��Ƶ��,0) IN(0,[2]) " & _
            IIF(mint���� <> 3, " And Nvl(A.�����Ա�,0) IN(0,[3]) ", "")
        
    If lng��̬ = 0 Then
        str��� = " And Nvl(C.��ҩ��̬,0) = [4] And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([1],3)" & _
                " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)"
    Else
         str��� = " And Exists(Select 1 From ҩƷ��� F,�շ���ĿĿ¼ S Where F.ҩƷid=C.ҩƷid And F.ҩƷid=S.ID And Nvl(F.��ҩ��̬,0) = [4] And S.������� IN([1],3) " & _
                    "And (S.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or S.����ʱ�� IS NULL) And (S.վ��='" & gstrNodeNo & "' Or S.վ�� is Null) )"
    End If
     
    If gblnStock And (mlng��ҩ�� <> 0 And mint���� <> 3 Or lngҩ��ID <> 0 And mint���� = 3) Then
        str��� = str��� & " And X.���>0"
        If blnFirst Then 'ȥ��ҩ������������
            strStock = "Select ҩƷID,Sum(Nvl(��������,0)) as ��� From ҩƷ���" & _
                " Where (Nvl(����,0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))" & _
                " And ���� = 1 Group by ҩƷID Having Sum(Nvl(��������,0))<>0"
        End If
    End If
          
    If blnInputKey Then
        strCode = UCase(vsExt.EditText)
        strLike = gstrLike
                
        strInput = " And (A.���� Like [5] And B.����=[7]" & _
                    " Or B.���� Like [6] And B.����=[7] Or B.���� Like [6] And B.���� IN([7],3))"
        If IsNumeric(strCode) Then
            '1X.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
            If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.���� Like [5] And B.����=[7] Or B.���� Like [6] And B.����=3)"
        ElseIf zlCommFun.IsCharAlpha(strCode) Then
            'X1.����ȫ����ĸʱֻƥ�����
            If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.���� Like [6] And B.����=[7]"
        ElseIf zlCommFun.IsCharChinese(strCode) Then
            '��������,��ֻƥ������
            strInput = " And B.���� Like [6] And B.����=[7]"
        End If
                
        strSQL = "Select Distinct A.ID,A.����,A.����,A.���㵥λ" & _
            " From ������ĿĿ¼ A,������Ŀ���� B" & _
            " Where A.ID=B.������ĿID" & str���� & strInput
               
        strSQL = _
            " Select C.ҩƷID as ID,D.����,A.����,A.���㵥λ as ��λ,D.���,D.����," & _
            IIF(InStr(strPrivs, "��ʾҩƷ���") = 0, " Decode(Sign(Nvl(X.���,0)),1,'��','')", _
                " Decode(X.���,NULL,NULL,X.���/" & IIF(mint������� = 1, "C.�����װ||C.���ﵥλ)", "C.סԺ��װ||C.סԺ��λ)")) & _
            " as ���,d.�������� As ��������,E.����ְ�� as ����ְ��ID,C.ҩƷID,A.ID as ҩ��ID" & _
            " From ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,(" & strSQL & ") A,(" & strStock & ") X" & _
            " Where A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID And C.ҩƷID=X.ҩƷID(+)" & str���� & str��� & str�洢�ⷿ & _
            IIF(strLike = "", "", " And Rownum<=100") & _
            " Order by D.����"
            
        vPoint = zlControl.GetCoordPos(vsExt.hwnd, vsExt.CellLeft, vsExt.CellTop)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҩ", False, "", "", False, False, True, vPoint.X, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
            mint�������, IIF(mint��Ч = 0, 2, 1), int�Ա�, lng��̬, strCode & "%", _
            strLike & strCode & "%", gbytCode + 1, mlng���˿���id)
    Else
        strSQL = "Select 0 as ĩ��,-1 as ID,-NULL as �ϼ�ID,NULL as ����," & _
            " CHR(13)||'������ҩ' as ����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ����ְ��ID,NULL as ҩƷID,NULL as ҩ��ID From Dual" & _
            " Union ALL" & _
            " Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ��������,NULL as ����ְ��ID,NULL as ҩƷID,NULL as ҩ��ID" & _
            " From ���Ʒ���Ŀ¼ Where ����=3 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"

         strMain = "Select 1 as ĩ��,C.ҩƷID as ID,A.����ID as �ϼ�ID,D.����,A.����,A.���㵥λ as ��λ,D.���,D.����," & _
             IIF(InStr(strPrivs, "��ʾҩƷ���") = 0, " Decode(Sign(Nvl(X.���,0)),1,'��','')", _
                " Decode(X.���,NULL,NULL,X.���/" & IIF(mint������� = 1, "C.�����װ||C.���ﵥλ)", "C.סԺ��װ||C.סԺ��λ)")) & _
            " as ���,d.�������� As ��������,E.����ְ�� as ����ְ��ID,C.ҩƷID,A.ID as ҩ��ID" & _
            " From ������ĿĿ¼ A,ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,(" & strStock & ") X" & _
            " Where A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID And C.ҩƷID=X.ҩƷID(+)" & str���� & str���� & str��� & str�洢�ⷿ
            
        strSQL = strSQL & " Union ALL " & strMain
        strPerson = Replace(strMain, "ҩƷ���� E", "ҩƷ���� E,���Ƹ�����Ŀ T") & " And T.������ĿID=A.ID And T.��ԱID=[5]"
        strSQL = strSQL & " Union ALL " & strPerson
        
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "��ҩ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, True, _
            mint�������, IIF(mint��Ч = 0, 2, 1), int�Ա�, lng��̬, UserInfo.ID, 0, 0, mlng���˿���id)
    End If
    

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "δ�ҵ����õ���ҩ��Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
        End If
        If blnInputKey Then vsExt.TextMatrix(vsExt.Row, vsExt.Col) = CStr(vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col))
        Exit Function
    End If
                
    '����ظ�����
    If lng��̬ = 0 Then
        On Error Resume Next
        strKey = mcol�������("_" & rsTmp!ҩ��ID & "_" & rsTmp!ҩƷID)
        If err.Number = 0 Then
            MsgBox "��ζ��ҩ���䷽���Ѿ�¼�롣", vbInformation, gstrSysName
            If blnInputKey Then vsExt.TextMatrix(vsExt.Row, vsExt.Col) = CStr(vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col))
            Exit Function
        End If
        On Error GoTo 0: err.Clear
    Else
        If ItemExist(rsTmp!ҩ��ID, vsExt.Row, vsExt.Col) Then
            MsgBox "��ζ��ҩ���䷽���Ѿ�¼�롣", vbInformation, gstrSysName
            If blnInputKey Then vsExt.TextMatrix(vsExt.Row, vsExt.Col) = CStr(vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col))
            Exit Function
        End If
    End If
    
    '����ְ����
    If mint���� = 0 Then
        strSQL = CheckOneDuty(rsTmp!����, NVL(rsTmp!����ְ��ID), UserInfo.����, mblnҽ��)
        If strSQL <> "" Then
            MsgBox strSQL, vbInformation, gstrSysName
            If blnInputKey Then vsExt.TextMatrix(vsExt.Row, vsExt.Col) = CStr(vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col))
            Exit Function
        End If
    End If
    
    strKey = vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col + 2)
    lngҩƷID = -1  '��һζҩ����ɾ����
    If strKey <> "" Then
        If lng��̬ = 0 Then  '����ǵ�һζɢװҩ�������ˣ�����ҩ�����Ÿı�
            If vsExt.Row = vsExt.FixedRows And vsExt.Col = vsExt.FixedCols Then
                If mcol�������("_" & strKey) <> "" Then
                    lngҩƷID = Val(Split(mcol�������("_" & strKey), ",")(0))
                Else
                    lngҩƷID = 0
                End If
            End If
        End If
        mcol�������.Remove "_" & strKey
    End If
    
    '��ȡ����ֵ
    If blnInputKey Then vsExt.EditText = rsTmp!���� 'ֱ������ƥ��ʱ��Ҫ
    vsExt.TextMatrix(vsExt.Row, vsExt.Col) = rsTmp!����
    vsExt.TextMatrix(vsExt.Row, vsExt.Col + 2) = rsTmp!��λ
    vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = vsExt.TextMatrix(vsExt.Row, vsExt.Col)
    
    If lng��̬ = 0 Then
        strKey = rsTmp!ҩ��ID & "_" & rsTmp!ҩƷID
    Else
        strKey = "" & rsTmp!ҩ��ID '��¼��ҩID
    End If
    vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col + 2) = strKey
        
    If lng��̬ = 0 Then
        '�������������ҩƷ�������޸ģ����Բ��ܼ�����:If optMode(1).Enabled Then
        If Check��ͬɢװ��ҩ Then
            optMode(1).Enabled = False: optMode(2).Enabled = False
        ElseIf optMode(1).Enabled = False Then
            optMode(1).Enabled = True: optMode(2).Enabled = True
        End If
    End If
    
    If lng��̬ = 0 Then
        mcol�������.Add rsTmp!ҩƷID & ",0", "_" & strKey
        
        'ɢװ��̬����ʱ����ȷ�˹�������һζɢװҩƷ�Ĺ����ˣ��������ҩ��
        If lngҩƷID <> rsTmp!ҩƷID And vsExt.Row = vsExt.FixedRows And vsExt.Col = vsExt.FixedCols Then
            If cboҩ��.ListIndex <> -1 Then
                lngҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
            Else
                lngҩ��ID = IIF(mlngPreRow��ҩ�� = 0, mlng��ҩ��, mlngPreRow��ҩ��)
            End If
            Call Get��ҩ��(cboҩ��, Val(rsTmp!ҩƷID), mlng���˿���id, mint�������, lngҩ��ID)
            If cboҩ��.ListIndex = -1 And cboҩ��.ListCount > 0 Then Call Cbo.SetIndex(cboҩ��.hwnd, 0)
            
            If cboҩ��.ListIndex <> -1 Then
                i = cboҩ��.ItemData(cboҩ��.ListIndex)
            Else
                i = 0
            End If
            If lngҩ��ID <> i Then Call ReSet��ҩ���
        End If
    Else
        mcol�������.Add "", "_" & rsTmp!ҩ��ID
        If blnFirst Then
            Set rs��� = Get��ҩ���(rsTmp!ҩ��ID, lng��̬, blnFirst)
            If rs���.RecordCount > 0 Then
                rs���.Filter = "��ҩ��̬ = " & lng��̬
                If rs���.RecordCount > 0 Then
                    Call Get��ҩ��(cboҩ��, Val(rs���!ҩƷID), mlng���˿���id, mint�������, lngҩ��ID)
                    If cboҩ��.ListIndex = -1 And cboҩ��.ListCount > 0 Then Call Cbo.SetIndex(cboҩ��.hwnd, 0)
                Else
                    MsgBox "δ�ҵ���ҩƷ����Ҫ�����̬����ѡ��������̬", vbInformation, gstrSysName
                    vsExt.TextMatrix(vsExt.Row, vsExt.Col) = ""
                    Exit Function
                End If
            Else
                MsgBox "δ�ҵ���ҩƷ�κο��õĹ����ѡ������ҩƷ", vbInformation, gstrSysName
                vsExt.TextMatrix(vsExt.Row, vsExt.Col) = ""
                Exit Function
            End If
        End If
    End If
    
    '����������ʱ���޸�ҩ��
    dbl���� = Val(vsExt.TextMatrix(vsExt.Row, vsExt.Col + 1))
    If dbl���� <> 0 Or lng��̬ = 0 Then
        Call Split��ҩ���(Val(rsTmp!ҩ��ID), dbl����, strKey)
        Call Show��ҩ���(Val(rsTmp!ҩ��ID), dbl����)
    End If
    
    '��������
    If lng��̬ = 0 Then
        Call SetSameItem(Val(rsTmp!ҩ��ID))
    End If
    
    Call EnterNextCell(vsExt.Row, vsExt.Col)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboҩ��_Click()
    '���·����������
    If Me.Visible = False Then Exit Sub
    If cboҩ��.Tag <> "" And Val(cboҩ��.Tag) = cboҩ��.ListIndex Then Exit Sub
    cboҩ��.Tag = cboҩ��.ListIndex
    
    Call ReSet��ҩ���
End Sub

Private Sub cboҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboData_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboData.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = Cbo.MatchIndex(cboData.hwnd, KeyAscii)
        If lngIdx = -1 And cboData.ListCount > 0 Then lngIdx = 0
        cboData.ListIndex = lngIdx
    End If
End Sub

Private Sub txtJL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtJL_GotFocus()
    Call zlControl.TxtSelAll(txtJL)
End Sub

Private Sub cmdInsert_Click()
    Call vsExt_KeyDown(vbKeyInsert, 0)
End Sub

Private Sub cmdOK_Click()
    Dim str��ҩIDs As String, blnSkip As Boolean
    Dim strMsg As String, strTmp As String
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim strժҪ As String, str������� As String, lng��ҩ��̬ As Long
    Dim strKey As String, lngҩ��ID As Long
    Dim str�������� As String
    Dim blnMsg As Boolean
    
    Dim lngBegin As Long, lngEnd As Long
    Dim strAppend As String, strData As String

    blnSkip = False
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            strKey = vsExt.Cell(flexcpData, i, j + 2)
            If strKey <> "" Then
                str������� = CStr(mcol�������("_" & strKey))
                If str������� = "" Then
                    MsgBox "ҩƷ""" & vsExt.TextMatrix(i, j) & """δ�ҵ����ù��", vbInformation, gstrSysName
                    vsExt.Select i, j + 1
                    vsExt.SetFocus: Exit Sub
                End If
                If cboҩ��.ListIndex = -1 And mint���� <> 3 Then
                    MsgBox "��ѡ��һ����ҩҩ����", vbInformation, gstrSysName
                    cboҩ��.SetFocus: Exit Sub
                End If
                If InStr(str�������, "|") > 0 Then
                    MsgBox "��������������ʣ�࣬�����""" & vsExt.TextMatrix(i, j) & """��������ѡ����ɢװ�����档", vbInformation, gstrSysName
                    vsExt.Select i, j + 1
                    vsExt.SetFocus: Exit Sub
                End If
                If Val(vsExt.TextMatrix(i, j + 1)) = 0 Then
                    If Not blnSkip Then
                        If MsgBox("""" & vsExt.TextMatrix(i, j) & """û�����뵥ζ������Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            vsExt.Row = i: vsExt.Col = j + 1
                            Call vsExt.ShowCell(i, j + 1)
                            vsExt.SetFocus: Exit Sub
                        End If
                        blnSkip = True
                    End If
                End If
                If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                    '��ϳ����ָ�ʽ�����ID1,����,��ע;���ID3,����,��ע
                    strTmp = strTmp & ";" & Replace(CStr(mcol�������("_" & strKey)), ";", "," & vsExt.TextMatrix(i, j + 3) & ";") & "," & vsExt.TextMatrix(i, j + 3)
                    str��ҩIDs = str��ҩIDs & "," & Split(strKey, "_")(0)
                End If
            End If
        Next
    Next
    strTmp = Mid(strTmp, 2)
    str��ҩIDs = Mid(str��ҩIDs, 2)
    lng��ҩ��̬ = Get��ҩ��̬
    
    If strTmp = "" Then
        MsgBox "�����䷽����������һζ��ҩ��", vbInformation, gstrSysName
        vsExt.Row = vsExt.FixedRows: vsExt.Col = 0
        vsExt.SetFocus: Exit Sub
    End If
    If cboData.ListIndex = -1 Then
        MsgBox "��ȷ����ҩ�䷽�ļ巨��", vbInformation, gstrSysName
        cboData.SetFocus: Exit Sub
    End If
    If cboҩ��.ListIndex = -1 And lng��ҩ��̬ <> 0 Then
        MsgBox "��ȷ����ҩҩ����", vbInformation, gstrSysName
        cboҩ��.SetFocus: Exit Sub
    End If
    
    '����ְ����(���ײ��ã�
    If mint���� = 0 Then
        strSQL = "Select /*+ Rule*/ ҩ��ID,����ְ�� From ҩƷ���� Where ҩ��ID IN(Select Column_Value From Table(f_Num2list([1])))"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��ҩIDs)
        For i = vsExt.FixedRows To vsExt.Rows - 1
            For j = 0 To vsExt.Cols - 1 Step 4
                strKey = vsExt.Cell(flexcpData, i, j + 2)
                If strKey <> "" Then
                    lngҩ��ID = Val(Split(strKey, "_")(0))
                Else
                    lngҩ��ID = 0
                End If
                
                If lngҩ��ID <> 0 Then
                    If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                        rsTmp.Filter = "ҩ��ID=" & lngҩ��ID
                        If Not rsTmp.EOF Then
                            strMsg = CheckOneDuty(vsExt.TextMatrix(i, j), NVL(rsTmp!����ְ��), UserInfo.����, mblnҽ��)
                            If strMsg <> "" Then
                                vsExt.Row = i: vsExt.Col = j
                                Call vsExt.ShowCell(i, j)
                                MsgBox strMsg, vbInformation, gstrSysName
                                vsExt.SetFocus: Exit Sub
                            End If
                        End If
                    End If
                End If
            Next
        Next
    End If
    
    'ҩƷ���ɼ�飨���ײ��ã�
    If mint���� <> 3 Then
        If Not Check��ҩ����(str��ҩIDs) Then Exit Sub
    End If
    If cboҩ��.ListIndex <> -1 Then
        i = Val(cboҩ��.ItemData(cboҩ��.ListIndex))
    Else
        i = 0
    End If
    strTmp = strTmp & "|" & cboData.ItemData(cboData.ListIndex) & "|" & lng��ҩ��̬ & "|" & Val(txt����.Text) & "|" & i
    
    'ҽ����Ϣ��ʾ(���ײ��ã�
    If mint���� <> 3 Then
        If Not mclsInsure Is Nothing And mlng����ID <> 0 Then  '��ҩ�䷽
            'ҽ��������������ʱ����ʾ
            If UBound(Split(str��ҩIDs, ",")) = 0 Then
                strժҪ = mclsInsure.GetItemInfo(mint����, mlng����ID, Val(Split(mcol�������.Item(1), ",")(0)), "", 0, "", str��ҩIDs & "||" & mint��������)
            Else
                strժҪ = mclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", str��ҩIDs & "||" & mint��������)
            End If
        End If
    End If
    
    strTmp = strTmp & "|" & Trim(txtJL.Text)
    
    If InStr(";" & strTmp, mstr�䷽��ϸ) > 0 And mstr�䷽��ϸ <> "" Then
        strTmp = strTmp & "|1"
    Else
        strTmp = strTmp & "|0"
    End If
    
    mstrExtData = strTmp
    mstrժҪ = strժҪ
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check��ͬɢװ��ҩ() As Boolean
'���ܣ�����Ƿ������ͬɢװ����ҩƷ��
    Dim i As Long, j As Long, strKey As String
    Dim colTmp As New Collection
    
    On Error Resume Next
    With vsExt
        For i = .FixedRows To .Rows - 1
            For j = 0 To .Cols - 1 Step 4
                strKey = .Cell(flexcpData, i, j + 2)
                If strKey <> "" Then
                    colTmp.Add 1, "_" & Split(strKey, "_")(0)
                    If err.Number > 0 Then
                        err.Clear
                        Check��ͬɢװ��ҩ = True
                        Exit Function
                    End If
                End If
            Next
        Next
    End With
End Function

Private Function CheckIsNullZY(ByRef lngRow As Long, ByRef lngCol As Long) As Boolean
'���ܣ�����Ƿ����Ѿ�������δ��д��
    Dim blnChange As Boolean
    Dim i As Long, j As Long
    
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = vsExt.FixedCols To vsExt.Cols - 1 Step 4
            If vsExt.TextMatrix(i, j) = "" Then
                lngRow = i: lngCol = j
                blnChange = True
                Exit For
            End If
        Next
        If blnChange Then Exit For
    Next
    
    If j > vsExt.Cols - 1 Then j = j - 4
    If i > vsExt.Rows - 1 Then i = i - 1
    If j = vsExt.Cols - 4 Then
        If i = vsExt.Rows - 1 Then
            Exit Function
        ElseIf vsExt.TextMatrix(i + 1, vsExt.FixedCols) <> "" Then
            CheckIsNullZY = True
        End If
    Else
        If vsExt.TextMatrix(i, j + 4) <> "" Then
            CheckIsNullZY = True
        End If
    End If
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ�������һ����ҩ�䷽�����뵥Ԫ��

    '��ǰλ��δ������ҩ
    If Val(vsExt.Cell(flexcpData, lngRow, (lngCol \ 4) * 4 + 2)) = 0 Then Exit Sub
    
    '����δ����
    If lngCol Mod 4 = 1 And vsExt.TextMatrix(lngRow, lngCol) = "" Then Exit Sub
    
    If mbytUseType = 3 And (lngRow > 1 Or lngCol >= 3) Then
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
            
    If lngCol + 1 <= vsExt.Cols - 1 Then
        lngCol = lngCol + 1
    Else
        If lngRow + 1 > vsExt.Rows - 1 Then
            vsExt.AddItem "", vsExt.Rows
            Call SetSplitLine
        End If
        lngRow = lngRow + 1
        lngCol = vsExt.FixedCols
    End If
    
    vsExt.Row = lngRow: vsExt.Col = lngCol
End Sub

Private Sub cmd_Click()
'���ܣ�����Ŀѡ����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strSQLItem As String, i As Long
    Dim strStock As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    
    On Error GoTo errH
    
    If CellCanEdit(vsExt.Row, vsExt.Col) Then
        If vsExt.Col Mod 4 = 0 Then
            Call Set��ҩInput(False)
            
        ElseIf vsExt.Col Mod 4 = 3 Then
            'ѡ���ע
            strSQL = "Select Rownum as ID,����,����,���� From ��ҩ�����ע Order by ����"
            vPoint = zlControl.GetCoordPos(vsExt.hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "��ע", , , , , , True, vPoint.X, vPoint.Y, vsExt.CellHeight, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ����õļ����ע�����ȵ�����������������á�", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            
            '��ȡ����ֵ
            vsExt.TextMatrix(vsExt.Row, vsExt.Col) = rsTmp!����
            vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = vsExt.TextMatrix(vsExt.Row, vsExt.Col)
            
            Call EnterNextCell(vsExt.Row, vsExt.Col)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CellCanEdit(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ܣ�������ҩ�䷽ʱ,�ж�ָ���ĵ�Ԫ��ǰ�Ƿ���������
'˵�������䷽��������,���ǰһ��δ����,��ǰ����������
    '��λ����һ����ҩ���뵥Ԫ
    On Error Resume Next
    
    '�����ǰ��ֵ���������޸�
    If Val(vsExt.Cell(flexcpData, lngRow, (lngCol \ 4) * 4)) <> 0 Or vsExt.TextMatrix(lngRow, (lngCol \ 4) * 4) <> "" Then
        CellCanEdit = True
        Exit Function
    End If
    If lngCol = (lngCol \ 4) * 4 + 1 Then '����ҩ��,��������
        If Val(vsExt.Cell(flexcpData, lngRow, (lngCol \ 4) * 4 + 2)) = 0 Then
            CellCanEdit = False
            Exit Function
        End If
    End If
    
    lngCol = (lngCol \ 4) * 4
    If lngCol - 4 >= vsExt.FixedCols Then
        lngCol = lngCol - 4
    Else
        If lngRow - 1 >= vsExt.FixedRows Then
            lngRow = lngRow - 1
            lngCol = vsExt.Cols - 4
        Else
            CellCanEdit = True
            Exit Function
        End If
    End If
    CellCanEdit = Val(vsExt.Cell(flexcpData, lngRow, lngCol + 2)) <> 0
End Function

Private Function ItemExist(ByVal lng��ҩID As Long, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ܣ��ж���ҩ�䷽��������,ָ������ҩ�Ƿ��Ѿ�����
    Dim i As Long, j As Long
    
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            If Not (lngRow = i And (lngCol \ 4) * 4 = j) Then
                If Val(vsExt.Cell(flexcpData, i, j + 2)) = lng��ҩID Then
                    ItemExist = True
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Private Function Check��ҩ����(ByVal str��ҩIDs As String) As Boolean
'���ܣ����һ���䷽�е���ҩ�������
'������str��ҩIDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str���� As String, str���� As String, lng���� As Long
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ���ƻ�����Ŀ" & _
        " Where ��ĿID+0 IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) Group by ���� Having Count(*)>1"
    strSQL = "Select /*+ Rule*/ A.����,A.����,B.����" & _
        " From ���ƻ�����Ŀ A,������ĿĿ¼ B" & _
        " Where A.��ĿID=B.ID And A.���� IN(" & strSQL & ")" & _
        " And A.��ĿID+0 IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
        " Order by A.����,B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��ҩIDs)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!���� <> lng���� Then
                If rsTmp!���� = 1 Then
                    str���� = str���� & vbCrLf & "��"
                Else
                    str���� = str���� & vbCrLf & "��"
                End If
                lng���� = rsTmp!����
            End If
            If rsTmp!���� = 1 Then
                str���� = str���� & "��" & rsTmp!����
            Else
                str���� = str���� & "��" & rsTmp!����
            End If
            rsTmp.MoveNext
        Next
        If str���� <> "" Then
            MsgBox "��ǰ�䷽�з�������ҩƷ������ã�" & Replace(str����, "��", "�� "), vbInformation, gstrSysName
            Exit Function
        ElseIf str���� <> "" Then
            If MsgBox("��ǰ�䷽�з�������ҩƷ�������ã�" & Replace(str����, "��", "�� ") & vbCrLf & vbCrLf & "Ҫ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    Check��ҩ���� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 0 Then
            If Me.Height - Y < 2355 Or Me.Height - Y > 7200 Then Exit Sub
            Me.Top = Me.Top + Y
            Me.Height = Me.Height - Y
        ElseIf Index = 1 Then
            If Me.Width + X < 4140 Or Me.Width + X > 9600 Then Exit Sub
            Me.Width = Me.Width + X
        ElseIf Index = 4 Then
            If vsExt.Height + Y < 1000 Or vsExt.Height + Y > Me.Height * 0.7 Then Exit Sub
            vsExt.Height = vsExt.Height + Y
            Call Form_Resize
        End If
    End If
End Sub
