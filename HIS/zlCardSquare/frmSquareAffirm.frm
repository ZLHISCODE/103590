VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSquareAffirm 
   Caption         =   "�������ѽ���"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSquareAffirm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9930
   StartUpPosition =   1  '����������
   Begin MSCommLib.MSComm mscCom 
      Left            =   7290
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picSum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   135
      ScaleHeight     =   1845
      ScaleWidth      =   3060
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1665
      Width           =   3090
      Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption2 
         Height          =   420
         Left            =   15
         TabIndex        =   20
         Top             =   30
         Width           =   3045
         _Version        =   589884
         _ExtentX        =   5371
         _ExtentY        =   741
         _StockProps     =   6
         Caption         =   "�������Ѻϼ�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label lbl�Ը��ϼ� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2025
         TabIndex        =   19
         Top             =   795
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdPara 
      Caption         =   "��ӡ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8160
      TabIndex        =   17
      Top             =   2610
      Width           =   1680
   End
   Begin VB.PictureBox picFee 
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   75
      ScaleHeight     =   4260
      ScaleWidth      =   11445
      TabIndex        =   15
      Top             =   3630
      Width           =   11445
      Begin VB.Frame fraSplitBottom 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   0
         TabIndex        =   16
         Top             =   -75
         Width           =   11805
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFee 
         Height          =   2505
         Left            =   -15
         TabIndex        =   12
         Top             =   405
         Width           =   9855
         _cx             =   17383
         _cy             =   4419
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSquareAffirm.frx":0442
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "������ϸ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -15
         TabIndex        =   11
         Top             =   150
         Width           =   840
      End
   End
   Begin VB.PictureBox picPayMode 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   3300
      ScaleHeight     =   1875
      ScaleWidth      =   4575
      TabIndex        =   14
      Top             =   1650
      Width           =   4575
      Begin VB.TextBox txt��Ԥ�� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   1
         Top             =   210
         Width           =   2760
      End
      Begin VB.TextBox txt��� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   1485
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1335
         Width           =   2790
      End
      Begin VB.ComboBox cbo֧����ʽ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   765
         Width           =   2775
      End
      Begin VB.Label lblԤ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Ԥ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   255
         Width           =   1110
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   810
         TabIndex        =   4
         Top             =   1410
         Width           =   630
      End
      Begin VB.Label lbl֧����ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֧����ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   195
         TabIndex        =   2
         Top             =   870
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   0
      TabIndex        =   13
      Top             =   1500
      Width           =   8025
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8205
      TabIndex        =   7
      Top             =   990
      Width           =   1515
   End
   Begin VB.Frame fraSplitLeft 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   8025
      TabIndex        =   8
      Top             =   60
      Width           =   30
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8190
      TabIndex        =   6
      Top             =   375
      Width           =   1500
   End
   Begin VB.Label lbl������� 
      AutoSize        =   -1  'True
      Caption         =   "�������:3333.22"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4830
      TabIndex        =   26
      Top             =   1110
      Width           =   2580
   End
   Begin VB.Label lblʣ����� 
      AutoSize        =   -1  'True
      Caption         =   "ʣ����:3333.22"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   25
      Top             =   1110
      Width           =   2580
   End
   Begin VB.Label lbl�Ա� 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�:��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3045
      TabIndex        =   9
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   90
      TabIndex        =   24
      Top             =   210
      Width           =   1110
   End
   Begin VB.Label lblMZH 
      AutoSize        =   -1  'True
      Caption         =   "�����:99999"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4830
      TabIndex        =   23
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1215
      TabIndex        =   22
      Top             =   255
      Width           =   570
   End
   Begin VB.Label lbl������� 
      AutoSize        =   -1  'True
      Caption         =   "δ�����:3232.22"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4830
      TabIndex        =   21
      Top             =   690
      Width           =   2580
   End
   Begin VB.Label lblԤ����� 
      AutoSize        =   -1  'True
      Caption         =   "Ԥ�����:3232.22"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   10
      Top             =   690
      Width           =   2580
   End
End
Attribute VB_Name = "frmSquareAffirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'��α���
Private mbytBillType As Byte '0-�������շѻ���ʵ�,1-�շѼ�¼;2-���ʼ�¼
Private mlngModule As Long, mlngPatiID As Long
Private mstrNos As String, mstrҽ��IDs As String, mstrPrivs As String
Private mstrExpand As String
Private mlngCardTypeID As Long, mbln���ѿ� As Boolean
Private mstrPrintNO As String
Private mblnCliniqueRoomPay As Boolean  '���֧��
Private mobjDrugPacker As Object
Private mblnDrugPacker As Boolean
Private mobjDrugMachine As Object
Private mblnDrugMachine As Boolean
Private mblnʹ��Ԥ�� As Boolean '�Ƿ�����ʹ��Ԥ����,104381
'---------------------------------------------------------------------
'ģ�����
Private mlng����ID As Long, mblnOk As Boolean
Private mcolPayMode As Collection
Private mrsInfo As ADODB.Recordset
Private mblnFirst As Boolean
Private mlng���￨���� As Long
Private mlng���������� As Long
Private mlng�����ID As Long 'ͨ������ˢ���Ŀ����ID
Private mstr����IDs As String    '����ID,�ö��ŷ���,���صĽ���ID���
'---------------------------------------------------------------------
'ģ�����
Private mblnReadCard As Boolean  '���ڶ�ȡ����
Private mintFeePrecision  As Integer
Private mstrFeePrecisionFmt  As String
Private mbytFeeMoneyPrecision  As Byte
Private mstrFeeMoneyPrecisionFmt As String
Private mblnSeekName As Boolean    '�Ƿ�ͨ����������ģ������
Private mintNameDays As Integer  'ͨ������ģ����������
Private mblnBrushCardPass As Boolean         'ˢ��Ҫ����������
Private mdbl�ʻ���� As Double
Private mstrCardNo As String  '����ˢ���Ŀ���
Private mstr������� As String
Private Type Ty_Para
        int���Ʊ�ݸ�ʽ As Integer
        int�շ�Ʊ�ݸ�ʽ As Integer
        int��˴�ӡ��ʽ As Integer
        int�շѴ�ӡ��ʽ As Integer
        intҩƷ��λ As Integer
End Type
Private mintCurType As Integer '1-�����շ�;2-�������
Private mPara As Ty_Para
Private mbytAssign As Byte '��ҩ���ڶ�̬���䷽ʽ(0,1)
'---------------------------------------------------------------------
'�ӿ����
Private WithEvents mobjIDCard As zlIDCard.clsIDCard  '���֤�ӿ�
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard   'IC���ӿ�
Attribute mobjICCard.VB_VarHelpID = -1
Private mobjCardPay As Object    '���������ӿڻ�����ӿ�
Private mblnPassInputCardNo As Boolean  '�Ƿ��������뿨��
Private mblnDefaultPassInputCardNo As Boolean 'ȱʡˢ���Ƿ��������뿨��
Private mlngҽ�ƿ�����  As Long
Private mblnPayCardNoPass As Boolean
'----------------------------------------------------------------------------
Private Type TY_ChargeMoney
    dbl�������Ѻϼ� As Double
    dbl���γ�Ԥ��  As Double
    dbl��ǰδ�� As Double
    dblԤ����� As Double
    dbl������� As Double
    dbl����Ԥ�� As Double
End Type
Private mCurCarge As TY_ChargeMoney
'------------------------------------------------------------------------------------------
Private mobjCard As clsCards

'��֧�����
Private Type TY_PayMoney
    lngҽ�ƿ����ID As Long
    bln���ѿ� As Boolean
    str���㷽ʽ As String
    str���� As String
    strˢ������ As String
    strˢ������ As String
    str������ˮ�� As String
    str����˵�� As String
    bln���� As Boolean
    bln��������  As Boolean
    intҽ�ƿ����� As Integer
    bln֧Ʊ As Boolean
    bln���ƿ� As Boolean
    blnOneCard As Boolean '�Ƿ�һ��ͨ����
    int���� As Integer '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����;<0 ��ʾ������֧��
    strNO As String
    lngID As Long 'Ԥ��ID
    lng����ID As Long
    objCard As clsCard
End Type
Private mCurCardPay As TY_PayMoney '���ο�֧��
Private mcllSquareBalance As Collection '���ѿ�����


Private mblnOK_Click As Boolean  '�������ȷ��:59412
'----------------------------------------------------------------------------
Private mPatiCard As SquareCard 'ˢ�������
Private mstrPassWord As String
Private mobjPatiCardObject As clsCardObject
Private mrsFeeData As ADODB.Recordset   '��¼����ˢ�����ѵ�����
Private mfrMain As Object
Private mstr����IDs As String '���˼���ID,79868
Private mdblԤ��������鿨 As Double 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤

'ҩ�������ڿ���
Private mlng��ҩ�� As Long 'ָ������ҩ��,0Ϊ��̬����
Private mlng��ҩ�� As Long 'ָ������ҩ��,0Ϊ��̬����
Private mlng��ҩ�� As Long 'ָ���ĳ�ҩ��,0Ϊ��̬����
Private mlng���ϲ��� As Long 'ָ�������ķ��ϲ���,0Ϊ��̬����

Private mstr���� As String  'ָ������ҩ����ҩ����,��Ϊ��̬����
Private mstr�д� As String 'ָ������ҩ����ҩ����,��Ϊ��̬����
Private mstr�ɴ� As String  'ָ���ĳ�ҩ����ҩ����,��Ϊ��̬����
Private mstrPayDrugWins As String '��ҩ�����ַ�������ʽ��ִ�в���1;��ҩ����1|...

Private Function zlGetFeeData(ByVal lng����ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ȡ�ķ�������
    '����:��ȡ��������
    '����:���˺�
    '����:2011-09-14 20:09:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTableNos As String, strTableIDs As String
    Dim varPara() As Variant, strWhere As String, strSubTable As String
    Dim rsTemp As ADODB.Recordset
    Dim strSfTable As String, strJzTable As String
    
     On Error GoTo errHandle
    If lng����ID = 0 Then Exit Function
    ReDim Preserve varPara(0 To 1) As Variant
    
    varPara(0) = lng����ID: varPara(1) = mbytBillType
         
    
    If mstrҽ��IDs <> "" Then
          If zlGetSubTable(0, mstrҽ��IDs, strTableIDs, varPara(), 2) = False Then Exit Function
    End If
    If mstrNos <> "" Then
          If zlGetSubTable(1, mstrNos, strTableNos, varPara(), UBound(varPara) + 1) = False Then Exit Function
    End If
 
    If mstrҽ��IDs <> "" And mstrNos <> "" Then
        strSubTable = " With  ҽ��  As (" & strTableIDs & "),���� as (" & strTableNos & ")"
        'strSQL = strSubTable & strSQL & " And A.NO= C.Column_Value And A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where A.����ID=[1] and A.ҽ�����=P.ID)"
    ElseIf mstrҽ��IDs <> "" Then
        strSubTable = " With  ҽ��  As (" & strTableIDs & ") "
'        strSQL = strSubTable & strSQL & " And  A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where A.����ID=[1] and A.ҽ�����=P.ID)"
    ElseIf strTableNos <> "" Then
        strSubTable = " With   ���� as (" & strTableNos & ")"
        'strSQL = strSubTable & strSQL & " And A.NO= C.Column_Value  "
    End If
    '110421:���ϴ�,2017/6/23,����ִ��ʱӦʹ�ü۸񸸺Ŷ����Ǵ�������
    strSfTable = "": strJzTable = ""
    If mbytBillType <= 1 Then
        strSfTable = "" & _
        "Select /*+ rule */ decode(A.��¼����,1,'�շ�',2,'����',4,'�Һ�') as ���,A.��¼����,A.ִ�в���ID,A.��ҩ����,A.����ID, " & vbNewLine & _
        "       A.NO,nvl(A.�۸񸸺�,A.���) as ���,B.����||'-'||B.���� as ��Ŀ,B.���,nvl(A.����,1)*A.���� as ����, " & vbNewLine & _
        "       B.���㵥λ,A.�շ�ϸĿID,A.��׼����,A.Ӧ�ս��,A.ʵ�ս��,A.�շ����,A.�Ǽ�ʱ��,a.�����־" & vbNewLine & _
        "From ������ü�¼ A,�շ���ĿĿ¼ B" & IIf(mstrNos <> "", " ,���� C", "") & vbNewLine & _
        "Where A.�շ�ϸĿID=B.ID And A.��¼����=1  And A.����ID=[1] And  A.��¼״̬=0 "
        If mstrҽ��IDs <> "" And mstrNos <> "" Then
            '����:49593
            strSfTable = strSfTable & " And (A.NO= C.Column_Value  or  A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where J.����ID=[1] and J.ҽ�����=P.ID And J.��¼����=1  ))"
        ElseIf mstrҽ��IDs <> "" Then
            strSfTable = strSfTable & " And  A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where J.����ID=[1] and J.ҽ�����=P.ID And J.��¼����=1)"
        ElseIf strTableNos <> "" Then
            strSfTable = strSfTable & " And A.NO= C.Column_Value  "
        End If
    End If
    If mbytBillType = 2 Or mbytBillType = 0 Then
        strJzTable = "" & _
        "Select /*+ rule */ decode(A.��¼����,1,'�շ�',2,'����',4,'�Һ�') as ���,A.��¼����,A.ִ�в���ID,A.��ҩ����,A.����ID, " & vbNewLine & _
        "       A.NO,nvl(A.�۸񸸺�,A.���) as ���,B.����||'-'||B.���� as ��Ŀ,B.���,nvl(A.����,1)*A.���� as ����, " & vbNewLine & _
        "       B.���㵥λ,A.�շ�ϸĿID,A.��׼����,A.Ӧ�ս��,A.ʵ�ս��,A.�շ����,A.�Ǽ�ʱ��,a.�����־" & vbNewLine & _
        "From ������ü�¼ A,�շ���ĿĿ¼ B" & IIf(mstrNos <> "", " ,���� C", "") & vbNewLine & _
        "Where A.�շ�ϸĿID=B.ID And A.��¼����=2  And A.����ID=[1] And  A.��¼״̬=0 "
        If mstrҽ��IDs <> "" And mstrNos <> "" Then
            '����:49593
            'strJzTable = strJzTable & " And A.NO= C.Column_Value And A.ҽ�����=P.ID"
            strJzTable = strJzTable & " And (A.NO= C.Column_Value  or  A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where J.����ID=[1] and J.ҽ�����=P.ID And J.��¼����=2  ))"
        ElseIf mstrҽ��IDs <> "" Then
            strJzTable = strJzTable & " And   A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where J.����ID=[1] and J.ҽ�����=P.ID And J.��¼����=2  ) "
        ElseIf strTableNos <> "" Then
            strJzTable = strJzTable & " And A.NO= C.Column_Value "
        End If
        If strSfTable <> "" Then strJzTable = vbCrLf & " Union all   " & vbCrLf & strJzTable
    End If
    strSQL = strSubTable & vbCrLf & strSfTable & vbCrLf & strJzTable
    strSQL = "  Select /*+ rule */  * From (" & strSQL & ") Order by ��¼����,NO,���"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "��ȡ���˷�����Ϣ", varPara)
    Set zlGetFeeData = rsTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function LoadFeeData(ByVal intTYPE As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�������
    ' ����:intType-1-�����շ�;2-����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-15 14:33:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTableNos As String, strTableIDs As String
    Dim varPara() As Variant, strWhere As String, strSubTable As String
    Dim lngRow As Long
    Dim dblMoney As Double, i As Long
    mintCurType = intTYPE
    If mrsFeeData Is Nothing Then Exit Function
    mrsFeeData.Filter = "��¼����=" & intTYPE
    With vsFee
        .Clear 1
        .Rows = IIf(mrsFeeData.RecordCount = 0, 1, mrsFeeData.RecordCount) + 1
        lngRow = 1
        If mrsFeeData.RecordCount <> 0 Then mrsFeeData.MoveFirst
        Do While Not mrsFeeData.EOF
            .RowData(lngRow) = Val(Nvl(mrsFeeData!���))
            .TextMatrix(lngRow, .ColIndex("���")) = Nvl(mrsFeeData!���)
            .Cell(flexcpData, lngRow, .ColIndex("���")) = Val(Nvl(mrsFeeData!��¼����))
            .TextMatrix(lngRow, .ColIndex("���ݺ�")) = Nvl(mrsFeeData!NO)
            .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = Trim(Nvl(mrsFeeData!�շ����))
            .TextMatrix(lngRow, .ColIndex("��Ŀ")) = Nvl(mrsFeeData!��Ŀ)
            .TextMatrix(lngRow, .ColIndex("���")) = Nvl(mrsFeeData!���)
            .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(Val(Nvl(mrsFeeData!����)), 5)
            .TextMatrix(lngRow, .ColIndex("��λ")) = Nvl(mrsFeeData!���㵥λ)
            .TextMatrix(lngRow, .ColIndex("����")) = FormatEx(Val(Nvl(mrsFeeData!��׼����)), mintFeePrecision)
            .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = FormatEx(Val(Nvl(mrsFeeData!Ӧ�ս��)), mbytFeeMoneyPrecision)
            .TextMatrix(lngRow, .ColIndex("ʵ�ս��")) = FormatEx(Val(Nvl(mrsFeeData!ʵ�ս��)), mbytFeeMoneyPrecision)
            .Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��")) = Val(Nvl(mrsFeeData!ʵ�ս��))
            .TextMatrix(lngRow, .ColIndex("�����־")) = Val(Nvl(mrsFeeData!�����־))
            dblMoney = dblMoney + Val(Nvl(mrsFeeData!ʵ�ս��))
            lngRow = lngRow + 1
            mrsFeeData.MoveNext
        Loop
    End With
    mrsFeeData.Filter = 0
    dblMoney = RoundEx(dblMoney, 2)
    mCurCarge.dbl�������Ѻϼ� = dblMoney
    mCurCarge.dbl��ǰδ�� = dblMoney
    mCurCarge.dbl���γ�Ԥ�� = 0
    lbl�Ը��ϼ�.Caption = Format(dblMoney, "####0.00;-###0.00;;")
    lbl�Ը��ϼ�.Tag = dblMoney
    '���þ���Ľ�������
    LoadFeeData = True
End Function

Public Function zlSquareAffirm(ByVal frmMain As Object, _
    ByVal lngModule As Long, strPrivs As String, _
    Optional ByVal lngPatiID As Long = 0, _
    Optional ByVal lngCardTypeID As Long = 0, _
    Optional ByVal bln���ѿ� As Boolean = False, _
    Optional ByVal blnCliniqueRoomPay As Boolean = False, _
    Optional ByVal bytBillType As Byte, _
    Optional ByVal strNos As String = "", _
    Optional ByVal strҽ��IDs As String = "", _
    Optional ByRef strExpand As String = "", _
    Optional ByRef lng����ID As Long = 0, _
    Optional ByVal blnʹ��Ԥ�� As Boolean = True) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ����ȷ�Ͻӿ� , ��Ҫ��Ӧ���ڲ����ڸ����ѻ�����������ȷ��
    '���:frmMain-������ö���
    '       lngModule:���õ�ģ���
    '       strPrivs:Ȩ�޴�
    '       lngPatiID :����ID,���Բ���,�ڱ��ӿڴ�����ˢ��!
    '       lngCardTypeID   Long    In  �����ID(���ѿ�Ϊ���ѽӿ����):0Ϊ������;��ȷ�ϴ����д��� Ŀǰ , ֻ����Ԥ����ɿ���ʹ��,�����,֧����ʽȱʡΪ�÷�ʽ.
    '       bln���ѿ�   Boolean In  ȱʡΪFase,��ʾ�Ƿ����ѿ�����
    '       bytBillType:�������: 0-�������շѻ���ʵ�,1-�շѼ�¼;2-���ʼ�¼
    '       strNOs:��ʽΪ( ����1,����2),���BytBillType��������ʹ��.һ��ֻ��ʹ��һ������
    '                   ��:  A0001,A002,A003��;
    '       strҽ��IDs:��ʽΪ:ID1,ID2,...
    '       strCardNO-��������ˢ�Ŀ���
    '       blnCliniqueRoomPay-���֧��(���֧��������ˢ������),���֧��ʱ��ֻ����շ�����
    '       blnʹ��Ԥ��-�Ƿ�����ʹ��Ԥ����Ture������ʹ��Ԥ����Ҵ���Ԥ����ʱȱʡʹ��Ԥ���False��������ʹ��Ԥ�������Ҫ�����õ������ʻ�
    '����:
    '����:Boolean ����    �ɹ�,����true,����ķ���False
    '����:���˺�
    '����:2011-06-15 09:53:37
    '˵��:
    '      ���strNos��strҽ��IDs��û��,ֻ�Ƕ�ָ�����˵������շѻ��۵��շѺ�������ʻ��۽������.
    '      �������ID������,����Ҫ�ڴ������Ƚ���ˢ���ҵ����˺�,�ٽ�������ȷ��.
    '������:
    '    1.  ���;����;ҩ����.
    '    2.  ����������Ҫ��������ȷ�ϵĵط���Ӧ�õ��øýӿ�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strPayType As String
    On Error GoTo errHandle
    mblnDrugPacker = False:  mblnDrugMachine = False
    mstrExpand = strExpand
    mlngModule = lngModule: mlngPatiID = lngPatiID: mstrPrivs = strPrivs
    mstrNos = strNos: mstrҽ��IDs = strҽ��IDs: mlng����ID = 0: mblnOk = False
    mbytBillType = bytBillType: mlngCardTypeID = lngCardTypeID
    mblnCliniqueRoomPay = blnCliniqueRoomPay
    mblnʹ��Ԥ�� = blnʹ��Ԥ��
    '������֧��
    If CliniqueRoomPayValied = False Then Exit Function
    If blnʹ��Ԥ�� = False Then
        strPayType = GetAvailabilityCardType
        If strPayType = "" Then
            MsgBox "ע��:" & vbCrLf & "    ��ǰû�п��õ�֧����𣬲���֧����", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If mblnCliniqueRoomPay Then  '���֧����Ҫ������صĲ���
        Call InitPara
    End If
    Set mrsFeeData = zlGetFeeData(lngPatiID)
    If mrsFeeData Is Nothing Then Exit Function
    If mrsFeeData.State <> 1 Then Exit Function
    If mrsFeeData.RecordCount = 0 Then zlSquareAffirm = True: Exit Function
    
    '95366:���ϴ�,2016/4/19,��ȡҩƷ���õ��ð�ҩ��
    Call CreateDrugPacker
    
    Set mfrMain = frmMain
    If mblnCliniqueRoomPay Then
        mblnOk = False
        If ExecuteCliniqueRoomPay = False Then
            Exit Function
        End If
        lng����ID = mlng����ID
        mblnOk = True
        zlSquareAffirm = mblnOk
        Exit Function
    End If
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    lng����ID = mlng����ID
    zlSquareAffirm = mblnOk
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2011-06-20 09:29:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mrsInfo = New ADODB.Recordset
    mstr����IDs = ""
    lbl����.Caption = ""
    lbl�Ա�.Caption = "�Ա�:"
    lblԤ�����.Caption = "Ԥ�����:0.00"
    lbl�������.Caption = "δ�����:0.00"
    lblʣ�����.Caption = "ʣ����:0.00"
    lbl�������.Caption = "�������:0.00"
    lbl�������.Visible = False
    lbl�Ը��ϼ�.Caption = "0.00"
    txt��Ԥ��.Text = ""
    txt���.Text = ""
    vsFee.Clear 1: vsFee.Rows = 2
End Sub
Private Sub cbo֧����ʽ_Click()
    Dim i As Long, lngIndex As Long
    '���ʲ�����
    With mCurCardPay
        .lngҽ�ƿ����ID = 0
        .bln���ѿ� = False
        .str���㷽ʽ = ""
        .str���� = ""
        .strˢ������ = ""
        .strˢ������ = ""
        .lngID = 0
        .strNO = ""
        .str���� = ""
        .bln�������� = False
        .intҽ�ƿ����� = 0
        .bln���� = False
        .bln֧Ʊ = False
        .blnOneCard = False
        .bln���ƿ� = False
        .int���� = 0
     End With
    If mintCurType = 2 Then Exit Sub
    With cbo֧����ʽ
        If .ListIndex = -1 Then GoTo SetProperty:
        lngIndex = .ListIndex + 1
        mCurCardPay.int���� = .ItemData(.ListIndex)
        mCurCardPay.blnOneCard = .ItemData(.ListIndex) = 7
    End With
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not mcolPayMode Is Nothing Then
        With mCurCardPay
            .lngҽ�ƿ����ID = Val(mcolPayMode(lngIndex)(3))
            .bln���ѿ� = Val(mcolPayMode(lngIndex)(5)) = 1
            .str���㷽ʽ = Trim(mcolPayMode(lngIndex)(6))
            .str���� = Trim(mcolPayMode(lngIndex)(1))
            .bln���� = Val(mcolPayMode(lngIndex)(2)) = 0
            .bln���ƿ� = Val(mcolPayMode(lngIndex)(8)) = 1
         End With
     Else
            mCurCardPay.str���㷽ʽ = zlstr.NeedName(cbo֧����ʽ.Text)
     End If
    '����������
    Call CreatePayObject
SetProperty:
End Sub
Private Sub CreatePayObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����֧������ӿ�
    '����:���˺�
    '����:2011-06-22 13:15:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�����ID As Long, bln���ѿ� As Boolean, int�Զ���ȡ As Integer
    Dim strKey As String
    Dim i As Long
    Set mobjCardPay = Nothing:
    Err = 0: On Error Resume Next
    
    If zlGetCardObj(Me, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mobjPatiCardObject) = False Then
        Set mobjPatiCardObject = Nothing
        Set mobjCardPay = Nothing
        Exit Sub
    End If
    Set mobjCardPay = mobjPatiCardObject.CardObject
    If Err <> 0 Then
        MsgBox "δ�ҵ�" & mCurCardPay.str���� & "����Ӧ�Ĳ���,����", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If mobjCardPay Is Nothing Then Exit Sub
End Sub

Private Function GetSelectNOs(ByRef str������Դ As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡѡ��ĵ��ݺ�
    '����:���ݺ�,����֮���ö��ŷ���,��:A0001,A0002....
    '����:���˺�
    '����:2011-06-23 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strNos As String, strNO As String
    
    If mblnCliniqueRoomPay Then
        '���֧��,��Ҫ�ر���
        mrsFeeData.Filter = "��¼����=1"
        mrsFeeData.Sort = "NO"
        strNos = ""
        With mrsFeeData
            Do While Not .EOF
                strNO = Nvl(!NO)
                If strNO <> "" Then
                    If InStr(1, strNos & ",", "," & strNO & ",") = 0 Then
                        strNos = strNos & "," & strNO
                        If InStr(str������Դ, Decode(Val(Nvl(!�����־)), 4, 3, 2, 2, 1)) = 0 Then
                            str������Դ = str������Դ & "," & Decode(Val(Nvl(!�����־)), 4, 3, 2, 2, 1)
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
        If strNos <> "" Then strNos = Mid(strNos, 2)
        GetSelectNOs = strNos
        Exit Function
    End If
    
    With vsFee
        For i = 1 To .Rows - 1
            strNO = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
            If strNO <> "" Then
                If InStr(1, strNos & ",", "," & strNO & ",") = 0 Then
                    strNos = strNos & "," & strNO
                    If InStr(str������Դ, Decode(Val(.TextMatrix(i, .ColIndex("�����־"))), 4, 3, 2, 2, 1)) = 0 Then
                        str������Դ = str������Դ & "," & Decode(Val(.TextMatrix(i, .ColIndex("�����־"))), 4, 3, 2, 2, 1)
                    End If
                End If
            End If
        Next
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If str������Դ <> "" Then str������Դ = Mid(str������Դ, 2)
    GetSelectNOs = strNos
End Function

Private Function GetSelectNOsAndSerialNum(ByRef strNos As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡѡ��ĵ��ݺź͵����е����
    '����:���ݺ�,����֮���ö��ŷ���,��:A0001|0;1;2,A0002|1;2;3....
    '����:���˺�
    '����:2011-06-23 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strNO As String
    Dim str��� As String, strData As String
    Dim j As Long
    With vsFee
        strData = "": strNos = ""
        For i = 1 To .Rows - 1
            strNO = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
            If InStr(1, strNos & ",", "," & strNO & ",") = 0 Then
                    str��� = ""
                    For j = 1 To .Rows - 1
                        If strNO = Trim(.TextMatrix(j, .ColIndex("���ݺ�"))) Then
                            str��� = str��� & ";" & .RowData(j)
                        End If
                    Next
                    If str��� <> "" Then str��� = Mid(str���, 2)
                    strNos = strNos & "," & strNO
                    strData = strData & "," & strNO & "|" & str���
             End If
        Next
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If strData <> "" Then strData = Mid(strData, 2)
    GetSelectNOsAndSerialNum = strData
End Function
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݺϷ��Լ��
    '����:���ݺϷ�������true,���򷵻�False
    '����:���˺�
    '����:2011-06-22 15:28:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mrsInfo Is Nothing Then
        MsgBox "������Ϣ����ȷ��,����!", vbInformation + vbOKOnly, gstrSysName
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
        Exit Function
    End If
    
    If mrsInfo.State <> 1 Then
        MsgBox "������Ϣ����ȷ��,����!!", vbInformation + vbOKOnly, gstrSysName
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
        Exit Function
    End If
    If mintCurType <> 2 Then
        '79621:���ϴ�,2014/11/14,�Խ���ʽ������
        If RoundEx(Val(txt��Ԥ��.Text) + Val(txt���), 2) <> RoundEx(Val(lbl�Ը��ϼ�.Tag), 2) Then
            If Val(txt���) = 0 Then
                MsgBox "���˵�Ԥ�������,���ֵ!", vbInformation + vbOKOnly, gstrSysName
            Else
                MsgBox "����֧������ϼ��뱾����Ҫ֧���ĺϼƲ��ȣ����ֵ!", vbInformation + vbOKOnly, gstrSysName
            End If
            If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
            Exit Function
        End If
    End If
    If (Val(txt��Ԥ��.Text) > 0 Or mintCurType = 2) And Val(lblԤ���.Tag) = 0 Then
        '֤��û����֤������Ҫ����������֤
          If CheckPrepayMoneyIsValied = False Then Exit Function
    End If
    
    isValied = True
End Function
Private Function CheckPrepayMoneyIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ԥ�����������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-24 10:36:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If BrushcardStrikePrepay = False Then
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��
        Exit Function
    End If
    CheckPrepayMoneyIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub setControlMove()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����
    '����:���˺�
    '����:2011-08-12 10:43:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngTop As Single, sngSplitHeight As Single, blnԤ�� As Boolean
    Dim sngHeght As Single
    sngSplitHeight = 80
    
    blnԤ�� = mCurCarge.dbl����Ԥ�� <> 0 Or cbo֧����ʽ.ListCount = 0
    If mintCurType = 1 Then
        ' �շ�
        lblԤ���.Visible = blnԤ��: txt��Ԥ��.Visible = blnԤ��
        sngHeght = picPayMode.ScaleHeight
        sngHeght = sngHeght - IIf(blnԤ��, txt��Ԥ��.Height - sngSplitHeight, 0)
        If cbo֧����ʽ.ListCount = 0 Then
            sngTop = (sngHeght + sngSplitHeight) / 2
        Else
            sngHeght = sngHeght - cbo֧����ʽ.Height - sngSplitHeight
            sngHeght = sngHeght - txt���.Height
            sngTop = sngHeght / IIf(blnԤ��, 3, 2)
        End If
        If blnԤ�� Then
            txt��Ԥ��.Top = sngTop: sngTop = txt��Ԥ��.Top + txt��Ԥ��.Height + sngSplitHeight
        End If
        cbo֧����ʽ.Top = sngTop: sngTop = cbo֧����ʽ.Top + cbo֧����ʽ.Height + sngSplitHeight
        txt���.Top = sngTop
        lblԤ���.Top = txt��Ԥ��.Top + (txt��Ԥ��.Height - lblԤ���.Height) \ 2
        lbl֧����ʽ.Top = cbo֧����ʽ.Top + (cbo֧����ʽ.Height - lbl֧����ʽ.Height) \ 2
        lbl���.Top = txt���.Top + (txt���.Height - lbl���.Height) \ 2
        Exit Sub
    End If
    '����
    sngHeght = picPayMode.ScaleHeight
    sngHeght = sngHeght - txt��Ԥ��.Height
    sngTop = sngHeght / 2
    txt��Ԥ��.Top = sngTop
    lblԤ���.Top = txt��Ԥ��.Top + (txt��Ԥ��.Height - lblԤ���.Height) \ 2
    cbo֧����ʽ.Visible = False: lbl֧����ʽ.Visible = False
    txt���.Visible = False: lbl���.Visible = False
End Sub

Private Function BrushcardStrikePrepay() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤ˢ����Ԥ��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Val(lblԤ���.Tag) = 1 Then BrushcardStrikePrepay = True: Exit Function
    If Val(txt��Ԥ��) = 0 And mintCurType <> 2 Then BrushcardStrikePrepay = True: Exit Function
    If mintCurType <> 2 Then If CheckPrepayValied = False Then Exit Function
     'ˢ��ȷ��
    'frmParent As Object, ByVal lngSys As Long, _
    ByVal lng����ID As Long, ByVal cur��� As Currency, _
    Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0
    gblnNotCloseWindows = True
    If zlDatabase.PatiIdentify(Me, glngSys, mlngPatiID, Val(txt��Ԥ��), mlngModule, 1, mlngCardTypeID, IIf(-1 * mdblԤ��������鿨 >= Val(txt��Ԥ��), False, True), True, _
        mstr����IDs, (mdblԤ��������鿨 <> 0), (mdblԤ��������鿨 = 2)) Then
        gblnNotCloseWindows = False
        lblԤ���.Tag = "1"
        txt��Ԥ��.BackColor = Me.BackColor
        txt��Ԥ��.Tag = Val(txt��Ԥ��): txt��Ԥ��.Enabled = False
        Call cbo֧����ʽ_Click
        '59412
        If mblnOK_Click Then BrushcardStrikePrepay = True: Exit Function
        If RoundEx(txt��Ԥ��.Text, 5) = RoundEx(Val(lbl�Ը��ϼ�.Tag), 5) Or mintCurType = 2 Then
            '���ʱ,��������
            If zlExcuteAffirm = False Then
                lblԤ���.Tag = "": txt��Ԥ��.Enabled = True: txt��Ԥ��.BackColor = vbWhite
                txt��Ԥ��.Tag = ""
                If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
                zlControl.TxtSelAll txt��Ԥ��
                Exit Function
            End If
        Else
           If txt���.Enabled And txt���.Visible Then txt���.SetFocus
            zlControl.TxtSelAll txt���
        End If
        BrushcardStrikePrepay = True
        Exit Function
    Else
        lblԤ���.Tag = "": txt��Ԥ��.Enabled = True
        txt��Ԥ��.Tag = ""
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��
        gblnNotCloseWindows = False
        Call cbo֧����ʽ_Click
       Exit Function
    End If
    BrushcardStrikePrepay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo֧����ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function WriteInforToCard(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String, _
        ByVal objSquareCard As Object, ByVal lngCardTypeID As Long, _
        ByVal strNos As String) As Boolean
    '����:��������Ϣд�뿨��
    '��Σ�
    '    frmMain - ���ô���
    '    lngModul - ģ���
    '    strPrivs - Ȩ�޴�
    '    objSquareCard - ҽ�ƿ�����
    '    strNOs - ���ݺţ���ʽ��'A0001','A0002','A0003',...��A0001,A0002,A0003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng����ID As Long, lng������� As Long
    
    Err = 0: On Error GoTo errH:
    '����:56615
'    If InStr(strPrivs, ";������Ϣд��;") = 0 Then Exit Function
    
    strSQL = "Select Distinct A.����ID,B.�������" & _
        " From ������ü�¼ A,����Ԥ����¼ B,Table( f_Str2list([1])) J" & _
        " Where A.����ID=B.����ID And A.NO=J.Column_Value And  Nvl(A.���ӱ�־,0)<>9 And A.��¼���� = 1 " & _
        "       And A.��¼״̬ in(1,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ݽ������", Replace(strNos, "'", ""))
    If rsTemp.EOF Then Exit Function
    Do While Not rsTemp.EOF
        lng����ID = Val(Nvl(rsTemp!����ID))
        lng������� = Val(Nvl(rsTemp!�������))
        '���ý�����д���ӿ�
        If lng����ID <> 0 And lng������� <> 0 Then
            Call objSquareCard.zlMzInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng����ID, lng�������)
        End If
        rsTemp.MoveNext
    Loop
    
    WriteInforToCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlExcuteAffirm() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ������ȷ��
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-14 22:46:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '����У��
    If isValied = False Then
        Exit Function
    End If
    If SaveData = False Then mstrPrintNO = "": Exit Function
    '��ӡƱ��
    Call PrintBill
    
    '��ҽһ��ͨд����85950
    If mintCurType = 1 Then '���ﻮ���շ�
        Call WriteInforToCard(Me, mlngModule, mstrPrivs, mPatiCard.objSquareCard, 0, mstrPrintNO)
    End If
    Set mPatiCard.objSquareCard = Nothing
    
    If mbytBillType = 0 And mintCurType = 1 Then
        mintCurType = 2
        Call LoadFeeData(2)
        setControlMove
        If vsFee.TextMatrix(1, vsFee.ColIndex("���ݺ�")) = "" Then
            mblnOk = True
            Unload Me
        End If
        Exit Function
    End If
    mblnOk = True: Unload Me
    zlExcuteAffirm = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub PrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡƱ��
    '����:���˺�
    '����:2014-01-20 11:01:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean, strFormat As String
    Dim frmMain As Object
    If mblnCliniqueRoomPay Then
        Set frmMain = mfrMain
    Else
        Set frmMain = Me
    End If
    Select Case mbytBillType
    Case 1, 4, 5
        blnPrint = mPara.int�շѴ�ӡ��ʽ = 1
        If mPara.int�շѴ�ӡ��ʽ = 2 Then
            If MsgBox("���Ƿ����Ҫ��ӡ�嵥��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrint = True
            End If
        End If
        If blnPrint Then
            strFormat = IIf(mPara.int�շ�Ʊ�ݸ�ʽ = 0, "", "ReportFormat=" & mPara.int�շ�Ʊ�ݸ�ʽ)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", frmMain, "NO=" & mstrPrintNO, "ҩƷ��λ=" & mPara.intҩƷ��λ, "PrintEmpty=0", strFormat, 2)
        End If
    Case 2
        blnPrint = mPara.int��˴�ӡ��ʽ = 1
        If mPara.int��˴�ӡ��ʽ = 2 Then
            If MsgBox("���Ƿ����Ҫ��ӡ�嵥��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrint = True
            End If
        End If
        If blnPrint Then
            strFormat = IIf(mPara.int���Ʊ�ݸ�ʽ = 0, "", "ReportFormat=" & mPara.int���Ʊ�ݸ�ʽ)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", frmMain, "NO=" & mstrPrintNO, "ҩƷ��λ=" & mPara.intҩƷ��λ, "PrintEmpty=0", strFormat, 2)
        End If
    End Select
End Sub
Private Sub cmdOK_Click()
     mblnOK_Click = True
    Call zlExcuteAffirm
End Sub
Private Function VerifyFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��˷���
    '����:��˳ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2011-06-23 09:59:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, cllPro As New Collection
    Dim varData As Variant, i As Long, strNO As String, strNos As String
    Dim strNosData As String, varTemp As Variant, str��� As String
    strNosData = GetSelectNOsAndSerialNum(strNos)
     '���ʵĻ�,Ҫ���ñ���
    If Not zlAuditingWarn(mstrPrivs, strNos, Val(Nvl(mrsInfo!����ID))) Then Exit Function
    varData = Split(strNosData, ",")
    For i = 0 To UBound(varData)
        If varData(i) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            strNO = varTemp(0): str��� = Replace(varTemp(1), ";", ",")
            mstrPrintNO = mstrPrintNO & ",'" & strNO & "'"
            'No_In/����Ա���_In /����Ա����_In /���_In/���ʱ��_In
             strSQL = "zl_������ʼ�¼_Verify('" & strNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "','" & str��� & "')"
             AddArray cllPro, strSQL
        End If
    Next
    If mstrPrintNO <> "" Then mstrPrintNO = Mid(mstrPrintNO, 2)
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllPro, Me.Caption
    VerifyFee = True
    
    '110319
    If mblnDrugMachine Then
        '�����ʽ��1|����1,������1;����2,������2
        Dim strData As String, strReturn As String
        strData = "1|" & "9," & Replace(Replace(strNos, "'", ""), ",", ";9,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-��ҩ[�����סԺ������ϸ�ϴ�]"), strData, strReturn)
    End If
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    mstrPrintNO = ""
End Function
Private Function SaveCharge() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�
    '����:�շѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-23 11:38:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, lng����ID As Long, lng������� As Long
    Dim dbl���� As Double, dbl������ As Double, dbl�޶��� As Double, dblMoney As Double
    Dim dblThreeMoney As Double, dblTemp As Double, dbl��Ԥ�� As Double
    Dim strNos As String, strNO As String, str����ʱ�� As String, strSQL As String, str����IDs As String
    Dim str������ˮ�� As String, str����˵�� As String, strSwapExtendInfor As String
    Dim cllPro As New Collection
    Dim int������Դ As Integer, intIndex As Integer, strTemp As String
    Dim rsTemp As ADODB.Recordset, cllDept As Collection
    Dim str��ҩ���� As String
    Dim strReturn As String, strData As String
    Dim str������Դ As String
 
    Err = 0: On Error GoTo Errhand:
    int������Դ = IIf(Val(Nvl(mrsInfo!��Ժ)) = 1, 2, 1)
    lng����ID = Val(Nvl(mrsInfo!����ID))
    strNos = GetSelectNOs(str������Դ)
    mstrPrintNO = "'" & Replace(strNos, ",", "','") & "'"
    strSQL = "" & _
    "   Select   /*+ rule */ NO,Max(���ʽ) as ���ʽ, " & _
    "               Max(���˿���ID) as ���˿���ID,Max(��������ID) as ��������Id, " & _
    "               Max(��ҩ����) as ��ҩ����,max(�Ƿ���) as �Ƿ���," & _
    "               Sum(ʵ�ս��) as ���,sum(case when instr([2],','||A.�շ����||',')>0 then a.ʵ�ս�� else 0 end)  as �޶���" & _
    "   From ������ü�¼ A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) J  " & _
    "  Where  A.No=J.Column_value and ��¼״̬=0 and A.��¼����=1" & _
    "   Group by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�������ѽ���-��ȡ������Ϣ", strNos, mstr�������)
    
    With rsTemp
        Do While Not .EOF
            dbl������ = dbl������ + Val(Nvl(rsTemp!���))
            dbl�޶��� = dbl�޶��� + Val(Nvl(rsTemp!�޶���))
            .MoveNext
        Loop
    End With
    dblTemp = RoundEx(dbl������, 6)
    dbl������ = RoundEx(dbl������, 2)
    dblMoney = dbl������
    dblThreeMoney = dbl������
    dbl���� = dblTemp - dblMoney
    dbl�޶��� = RoundEx(dbl�޶���, 2)
    
        
    If mblnCliniqueRoomPay = False Then '�����֧��ʱ����Ҫ�����ص����ݺϷ���
        '79621:���ϴ�,2014/11/14,�Խ���ʽ������
        If dbl������ <> RoundEx(Val(lbl�Ը��ϼ�.Tag), 2) Then
            If MsgBox("ע��:" & vbCrLf & "    ����ѡ��Ļ��۵��ݵ�ʵ�ս���Ѿ������仯,�Ƿ�������ȡ��Ӧ���ݵķ���!", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                Set mrsFeeData = zlGetFeeData(Val(Nvl(mrsInfo!����ID)))
                Call LoadFeeData(mintCurType): Exit Function
            End If
        End If
        '79621:���ϴ�,2014/11/14,�Խ���ʽ������
        If RoundEx(Val(txt��Ԥ��.Text) + Val(txt���.Text), 2) <> dbl������ Then
            MsgBox "ע��:" & vbCrLf & "    ������ۿ����(Ԥ���+" & cbo֧����ʽ.Text & "֧�������ڱ���֧���ķ��úϼ�,����!", vbOKOnly + vbDefaultButton1 + vbInformation
            Exit Function
        End If
        If cbo֧����ʽ.ListIndex >= 0 Then
            intIndex = cbo֧����ʽ.ListIndex + 1
            If Trim(mcolPayMode(intIndex)(6)) = "" Then
                MsgBox "ע��:" & vbCrLf & "    " & Trim(mcolPayMode(intIndex)(1)) & "  δ���ý��㷽ʽ,����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        ElseIf Val(txt���.Text) <> 0 Then
            MsgBox "ע��:" & vbCrLf & "    δѡ��֧�����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        dbl��Ԥ�� = RoundEx(Val(txt��Ԥ��.Text), 2)
        dblMoney = RoundEx(Val(txt���.Text), 2)
        '79621:���ϴ�,2014/11/14,�Խ���ʽ������
        If RoundEx(Val(lbl�Ը��ϼ�.Tag) - Val(txt���.Text) - dbl�޶���, 2) > RoundEx(Val(txt��Ԥ��.Text), 2) And Val(txt��Ԥ��.Text) <> 0 Then
            MsgBox "ע��:" & vbCrLf & "    ��Ԥ���Ķ���������,���ֻ�ܿ�Ԥ���:" & Format(Val(lbl�Ը��ϼ�.Tag) - Val(txt���.Text) - dbl�޶���, "0.00"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If RoundEx(dblMoney, 7) <> RoundEx(Val(txt���.Text), 7) Then
            MsgBox "ע��:" & vbCrLf & "    " & Trim(mcolPayMode(intIndex)(1)) & "  ֧���ϼƲ���ȷ,����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        '79621:���ϴ�,2014/11/17,�Խ���ʽ������
        If RoundEx(Val(txt��Ԥ��.Text), 4) <> RoundEx(mCurCarge.dbl�������Ѻϼ�, 4) Then
           If BrushCardThreeSwapCheck(strNos, Val(txt���.Text), str������Դ, lng����ID) = False Then Exit Function
           dblThreeMoney = Val(txt���.Text)
        End If
    Else
        If BrushCardThreeSwapCheck(strNos, dblThreeMoney, str������Դ, lng����ID) = False Then Exit Function
    End If
    
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    lng������� = -1 * lng����ID
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    Set cllDept = New Collection
    With mrsFeeData
        strTemp = ""
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(",5,6,7,", Nvl(!�շ����)) > 0 And _
                InStr(strTemp, "," & Nvl(!�շ����) & "|" & Nvl(!ִ�в���ID) & ",") = 0 Then
                cllDept.Add Array(Nvl(!�շ����), Val(Nvl(!ִ�в���ID)), Nvl(!��ҩ����))
            End If
            .MoveNext
        Loop
        str��ҩ���� = GetPayDrugWindow(lng����ID, CDate(str����ʱ��), cllDept)
    End With
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
               
            '---------------------------------------------------------------
            'Zl_���˻����շ�_Insert
            strSQL = "Zl_���˻����շ�_Insert("
            '  No_In         ������ü�¼.NO%Type,
            strSQL = strSQL & "'" & Nvl(rsTemp!NO) & "',"
            '  ����id_In     ������ü�¼.����id%Type,
            strSQL = strSQL & "" & ZVal(lng����ID) & ","
            '  ������Դ_In   Number,
            strSQL = strSQL & "" & int������Դ & ","
            '  ���ʽ_In   ������ü�¼.���ʽ%Type,
            If Nvl(mrsInfo!���ʽ����) <> "" Then
               strSQL = strSQL & "'" & Nvl(mrsInfo!���ʽ����) & "',"
            Else
               strSQL = strSQL & "'" & Nvl(rsTemp!���ʽ) & "',"
            End If
            '  ����_In       ������ü�¼.����%Type,
            strSQL = strSQL & "'" & Nvl(mrsInfo!����) & "',"
            '  �Ա�_In       ������ü�¼.�Ա�%Type,
            strSQL = strSQL & "'" & Nvl(mrsInfo!�Ա�) & "',"
            '  ����_In       ������ü�¼.����%Type,
            strSQL = strSQL & "'" & Nvl(mrsInfo!����) & "',"
            '  ���˿���id_In ������ü�¼.���˿���id%Type,
            strSQL = strSQL & "" & IIf(Val(Nvl(rsTemp!���˿���ID)) = 0, "NULL", Val(Nvl(rsTemp!���˿���ID))) & ","
            '  ��������id_In ������ü�¼.��������id%Type,
            strSQL = strSQL & "" & IIf(Val(Nvl(rsTemp!��������ID)) = 0, "NULL", Val(Nvl(rsTemp!��������ID))) & ","
            '  ������_In     ������ü�¼.������%Type,
            strSQL = strSQL & "NULL,"    ' �����ڲ�����,����ԭ���Ĳ���
            '  ����id_In     ������ü�¼.����id%Type,
            strSQL = strSQL & "" & lng����ID & ","
            '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
            strSQL = strSQL & "to_date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),"
            '  ����Ա���_In ������ü�¼.����Ա���%Type,
            strSQL = strSQL & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ������ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ��ҩ����_In   ������ü�¼.��ҩ����%Type := Null,
            strSQL = strSQL & "'" & str��ҩ���� & "',"
            '  �Ƿ���_In   ������ü�¼.�Ƿ���%Type := 0,
            strSQL = strSQL & "" & Val(Nvl(rsTemp!�Ƿ���)) & ","
            '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
            strSQL = strSQL & "to_date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'))"
            zlAddArray cllPro, strSQL
            .MoveNext
        Loop
    End With
    str����IDs = lng����ID
    mCurCardPay.lng����ID = lng����ID
    
    'bytType-1-�����ӿ�֧��;2-���ѿ�֧��,0-����
    If mCurCardPay.bln���ѿ� And dblMoney <> 0 Then
        If SetCurBalanceSQL(2, lng����ID, dblMoney, dbl��Ԥ��, 0, 0, dbl����, cllPro) = False Then Exit Function
    ElseIf dbl��Ԥ�� = dblThreeMoney Then
        If SetCurBalanceSQL(0, lng����ID, 0, dbl��Ԥ��, 0, 0, dbl����, cllPro) = False Then Exit Function
    Else
        If SetCurBalanceSQL(1, lng����ID, dblMoney, dbl��Ԥ��, 0, 0, dbl����, cllPro) = False Then Exit Function
    End If
    
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If mblnCliniqueRoomPay = False Then
        If Val(txt���.Text) = 0 Or mCurCardPay.bln���ѿ� Then
            '����϶��ǳ�Ԥ������Ϊ���ѿ���ҽԺ�Ŀ��ʻ�
            gcnOracle.CommitTrans
            mstr����IDs = str����IDs
            mlng����ID = lng�������
            SaveCharge = True
            
            GoTo DoDrugPacker:
            Exit Function
        End If
    End If
    
    ' Public Function zlPaymentMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    '    ByVal strCardNo As String, ByVal strBalanceIDs As String,byval strPrepayNos as string , _
    '    ByVal dblMoney As Double, _
    '    ByRef strSwapGlideNO As String, _
    '    ByRef strSwapMemo As String, _
    '    Optional ByRef strSwapExtendInfor As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�ʻ��ۿ��
    '    '���:frmMain-���õ�������
    '    '        lngModule-����ģ���
    '    '        strBalanceIDs-����ID,����ö��ŷ���
    '    '       strCardNo-����
    '    '       dblMoney-֧�����
    '    '����:strSwapGlideNO-������ˮ��
    '    '       strSwapMemo-����˵��
    '    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    If mblnCliniqueRoomPay Then
        If mobjCardPay.zlPaymentMoney(mfrMain, mlngModule, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.strˢ������, str����IDs, "", dblThreeMoney, str������ˮ��, str����˵��, strSwapExtendInfor) = False Then
                gcnOracle.RollbackTrans: Exit Function
        End If
    Else
        If mobjCardPay.zlPaymentMoney(Me, mlngModule, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.strˢ������, str����IDs, "", dblThreeMoney, str������ˮ��, str����˵��, strSwapExtendInfor) = False Then
                gcnOracle.RollbackTrans: Exit Function
        End If
    End If
    
    Dim cllUpdate As New Collection, cllOthers As New Collection
    Call zlAddUpdateSwapSQL(False, lng����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, str������ˮ��, str����˵��, cllUpdate, 0, 1)
    Call zlAddThreeSwapSQLToCollection(False, lng����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapExtendInfor, cllOthers)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, False, True
    Err = 0: On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllOthers, Me.Caption
    SaveCharge = True
    
DoDrugPacker:
    '95366:���ϴ�,2016/4/19,��ȡҩƷ���õ��ð�ҩ��
    If mblnDrugMachine Then
        '�°淢ҩ��
        '�����ʽ��1|����1,������1;����2,������2;...
        strData = "1|" & "8," & Replace(Replace(strNos, "'", ""), ",", ";8,")
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-��ҩ[�����סԺ������ϸ�ϴ�]"), strData, strReturn)
    ElseIf mblnDrugPacker Then
        '��ʽ������1,������1|����2,������2|...
        strData = "8," & Replace(Replace(strNos, "'", ""), ",", "|8,")
        Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.���, UserInfo.����, 0, strData, strReturn)
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrOthers:
    gcnOracle.CommitTrans   '�ܱ������,������
    Call ErrCenter
    SaveCharge = True
End Function

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '����:���˺�
    '����:2011-06-22 16:01:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1-�շѼ�¼;2-���ʼ�¼��4-�Һż�¼;5-����;10-��Ԥ��
    
    Select Case mintCurType
    Case 1  '�շѻ��۴���
        If SaveCharge = False Then Exit Function
        SaveData = True:
        '��ӡ��ص�Ʊ��
    Case 2 '���ۼ������
        If VerifyFee = False Then Exit Function
        SaveData = True: Exit Function
'    Case 10 '��Ԥ����
'        If SavePrePayMoney = False Then Exit Function
'        SaveData = True
    End Select
End Function
Private Sub cmdPara_Click()
    If frmSquareAffirmParaSet.SetPara(Me) = False Then Exit Sub
    Call InitFactPara
End Sub
 

Private Sub Form_Activate()
    Dim intTYPE As Integer
    
    If mblnCliniqueRoomPay Then Exit Sub
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If GetPatient() = False Then Unload Me: Exit Sub
    '���ط���
    If mbytBillType = 0 Then
        mrsFeeData.Filter = "��¼����=1"
        If mrsFeeData.RecordCount = 0 Then
            intTYPE = 2
        Else
            intTYPE = 1
        End If
        mbytBillType = intTYPE
        mrsFeeData.Filter = 0
    Else
       intTYPE = mbytBillType
    End If
    Call LoadFeeData(intTYPE)
    If mblnʹ��Ԥ�� Then '������ʹ��Ԥ����ʱ������Ԥ��
        '����Ԥ��
        Call LoadԤ�����(mrsInfo!����ID)
    End If
    Call SetCtlEnable
    If mCurCarge.dbl����Ԥ�� = 0 Then
        If cbo֧����ʽ.Enabled And txt���.Enabled And txt���.Visible Then txt���.SetFocus
        '91315,��ǰԤ���Ϊ0����Ϊ����ʱ���Կ������ѳɹ�
        If txt���.Visible Then txt���.Text = FormatEx(Val(lbl�Ը��ϼ�.Tag), mbytFeeMoneyPrecision)
        zlControl.TxtSelAll txt���
    Else
        '79621:���ϴ�,2014/11/17,�Խ���ʽ������
        If RoundEx(mCurCarge.dbl����Ԥ��, 2) > RoundEx(Val(lbl�Ը��ϼ�.Tag), 2) Then
            txt��Ԥ��.Text = FormatEx(Val(lbl�Ը��ϼ�.Tag), mbytFeeMoneyPrecision)
        Else
            txt��Ԥ��.Text = FormatEx(mCurCarge.dbl����Ԥ��, mbytFeeMoneyPrecision)
        End If
        If Val(txt��Ԥ��.Text) <> 0 And txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��
    End If
    Call setControlMove
    '78773:���ϴ�,2014-10-29,LED��ʾһ��֧ͨ����Ϣ
    Call ShowLedInfor
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2
        If cmdOK.Enabled = False Then Exit Sub
        Call cmdOK_Click: Exit Sub
    Case vbKeyF4
        If Me.ActiveControl Is txt��� Then
            If cbo֧����ʽ.Enabled = False Then Exit Sub
            If Me.ActiveControl Is txt��� And txt���.Enabled = False Then Exit Sub
            If Shift = vbShiftMask Then
                If cbo֧����ʽ.ListIndex - 1 < 0 Then
                    cbo֧����ʽ.ListIndex = cbo֧����ʽ.ListCount - 1
                Else
                    cbo֧����ʽ.ListIndex = cbo֧����ʽ.ListIndex - 1
                End If
            Else
                If cbo֧����ʽ.ListIndex + 1 > cbo֧����ʽ.ListCount - 1 Then
                    cbo֧����ʽ.ListIndex = 0
                Else
                    cbo֧����ʽ.ListIndex = cbo֧����ʽ.ListIndex + 1
                End If
            End If
        End If
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '78773:���ϴ�,2014-10-29,LED��ʾһ��֧ͨ����Ϣ
    If gblnLED Then zl9LedVoice.DisplayPatient ""
    If Not mobjDrugPacker Is Nothing Then Set mobjDrugPacker = Nothing
    If Not mobjDrugMachine Is Nothing Then Set mobjDrugMachine = Nothing
End Sub

Private Sub picFee_Resize()
    Err = 0: On Error Resume Next
    With picFee
        vsFee.Left = .ScaleLeft
        vsFee.Height = .ScaleHeight - vsFee.Top
        vsFee.Width = .ScaleWidth - vsFee.Left
    End With
End Sub
Private Function LoadԤ�����(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ�����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-21 10:47:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    '79868,�����˼��������벡��ʣ���
    '��ü�¼�����ֻ��������һ���ǲ��˱��˵ģ�һ���ǲ��˼�����
    Set rsTemp = GetMoneyInfo(lng����ID, , , 1, , , True)
    Dim dbl������� As Double, dbl������� As Double, dbl������� As Double
    With mCurCarge
        .dblԤ����� = 0
        .dbl������� = 0
        Do While Not rsTemp.EOF
            .dblԤ����� = .dblԤ����� + Val(Nvl(rsTemp!Ԥ�����))
            .dbl������� = .dbl������� + Val(Nvl(rsTemp!�������))
            If Nvl(rsTemp!����, 0) = 0 Then
                dbl������� = Val(Nvl(rsTemp!Ԥ�����))
                dbl������� = Val(Nvl(rsTemp!�������))
            Else
                dbl������� = Val(Nvl(rsTemp!Ԥ�����)) - Val(Nvl(rsTemp!�������))
            End If
            rsTemp.MoveNext
        Loop
        .dbl����Ԥ�� = .dblԤ����� - .dbl�������
        If .dbl����Ԥ�� < 0 Then .dbl����Ԥ�� = 0
    End With
    lblԤ�����.Caption = "Ԥ�����:" & Format(dbl�������, "###0.00;-###0.00;0.00;0.00")
    lblԤ�����.Tag = mCurCarge.dblԤ�����
    lbl�������.Caption = "δ�����:" & Format(dbl�������, "###0.00;-###0.00;0.00;0.00")
    lbl�������.Tag = mCurCarge.dbl�������
    lblʣ�����.Caption = "ʣ����:" & Format(dbl������� - dbl�������, "###0.00;-###0.00;0.00;0.00")
    lblʣ�����.Tag = mCurCarge.dbl����Ԥ��
    lbl�������.Caption = "�������:" & Format(dbl�������, "###0.00;-###0.00;0.00;0.00")
    lbl�������.Visible = dbl������� <> 0
    LoadԤ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô����С
    '����:���˺�
    '����:2011-09-15 11:26:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��С����ߴ�
    With gWinRect
        .MaxW = Me.Width
        .MaxH = Screen.Height * Screen.TwipsPerPixelY
        .MinH = Me.Height
        .MinW = Me.Width
    End With
    glngOld = GetWindowLong(hWnd, GWL_WNDPROC)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SetWindowResizeWndMessage)
End Sub
Private Sub Form_Load()
    mblnFirst = True
    If mblnCliniqueRoomPay Then Exit Sub
    If Not IsDesinMode Then
         Call SetWindowsSize
    End If
    zlControl.CboSetWidth cbo֧����ʽ.hWnd, cbo֧����ʽ.Width * 2
    zlControl.PicShowFlat picSum, -1, , 1: zlControl.PicShowFlat picPayMode, -1, , 1
    '��ʼ������
    Call InitFace: Call SetCtlEnable
End Sub
Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-06-21 13:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKindStr As String, blnVisible As Boolean
    mblnOK_Click = False
    Set mPatiCard = New SquareCard
    
    '��������,���ӽ��㿨�Ľ���
    Err = 0: On Error Resume Next
    Set mPatiCard.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        MsgBox "���㿨����zl9CardSquare.clsCardSquare����ʧ�ܣ�", vbInformation, gstrSysName
        Err = 0: On Error GoTo 0: Exit Sub
    End If
    If mPatiCard.objSquareCard Is Nothing Then Exit Sub
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If mPatiCard.objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Sub
    End If
    
    Call InitPara: Call ClearData: Call Load֧����ʽ
End Sub
Private Sub InitFactPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ʊ��صĲ���
    '����:���˺�
    '����:2011-08-11 00:24:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With mPara
        .int�շ�Ʊ�ݸ�ʽ = Val(zlDatabase.GetPara("�շ��վݸ�ʽ", glngSys, 1151))
        .int�շѴ�ӡ��ʽ = Val(zlDatabase.GetPara("�շѴ�ӡ��ʽ", glngSys, 1151))
        .int���Ʊ�ݸ�ʽ = Val(zlDatabase.GetPara("����վݸ�ʽ", glngSys, 1151))
        .int��˴�ӡ��ʽ = Val(zlDatabase.GetPara("��˴�ӡ��ʽ", glngSys, 1151))
        .intҩƷ��λ = Val(zlDatabase.SetPara("ҩƷ��λ", glngSys, 1151))
    End With
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ֵ
    '����:���˺�
    '����:2011-06-20 16:48:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intStart As Integer
    Dim strValue As String
    
    Call InitFactPara
    '���ﲡ������ʱ��Ҫˢ����֤
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    mdblԤ��������鿨 = Val(Split(strValue, "|")(0))
    '���õ��۱���λ��
    mintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    mstrFeePrecisionFmt = "0." & String(mintFeePrecision, "0")
    '���ý��С����λ��
    mbytFeeMoneyPrecision = Val(zlDatabase.GetPara(9, glngSys, , 2))
    mstrFeeMoneyPrecisionFmt = "0." & String(mbytFeeMoneyPrecision, "0")
    mblnSeekName = zlDatabase.GetPara("����ģ������", glngSys, mlngModule) = "1"
    mintNameDays = Val(zlDatabase.GetPara("������������", glngSys, mlngModule))
    mbytAssign = Val(zlDatabase.GetPara(19, glngSys, , 0))

    If mblnCliniqueRoomPay Then
        'ҩ�������ڷ��䷽ʽ
        mstr�д� = zlDatabase.GetPara(49, glngSys, mlngModule)
        mstr���� = zlDatabase.GetPara(50, glngSys, mlngModule)
        mstr�ɴ� = zlDatabase.GetPara(51, glngSys, mlngModule)
        
        mlng��ҩ�� = Val(zlDatabase.GetPara(18, glngSys, mlngModule))
        mlng��ҩ�� = Val(zlDatabase.GetPara(19, glngSys, mlngModule))
        mlng��ҩ�� = Val(zlDatabase.GetPara(20, glngSys, mlngModule))
        mlng���ϲ��� = Val(zlDatabase.GetPara(21, glngSys, mlngModule))
    Else
        mstr���� = "": mstr�д� = "": mstr�ɴ� = ""
        mlng��ҩ�� = 0: mlng��ҩ�� = 0: mlng��ҩ�� = 0: mlng���ϲ��� = 0
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    fraSplitBottom.Width = Me.ScaleWidth + fraSplitBottom.Left
    picFee.Width = Me.ScaleWidth - picFee.Left * 2
    picFee.Height = Me.ScaleHeight - picFee.Top - 50
End Sub

Private Function GetPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=��ʾ�Ƿ���￨ˢ��
    '����:
    '����:���˶�ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-20 16:04:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    '��ȡ������Ϣ
    strSQL = "" & _
    "   Select Decode(Sign(A.����ʱ��-A.�Ǽ�ʱ��),0,1,0) as ����,A.����ID,A.��������," & _
    "               A.IC����,A.���￨��,A.�����,A.סԺ��,A.����, A.����֤��, " & _
    "               A.�Ա�,A.����, A.��������,A.�ѱ�,A.������,A.ҽ�Ƹ��ʽ,M.���� as ���ʽ����,A.��Ժ," & _
    "               decode(B1.��������,NULL,0,1,1,0) as ����,B1.��Ժ����,A.����,C.���� ��������" & _
    "   From ������Ϣ A,������ҳ B1,������� C ,ҽ�Ƹ��ʽ M" & _
    "   Where A.���� = C.���(+) And A.ҽ�Ƹ��ʽ=M.����(+) " & _
    "               And A.����ID=B1.����ID(+) And A.��ҳID=B1.��ҳID(+) " & _
    "               And A.ͣ��ʱ�� is NULL And A.����ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, "�������ѽ���-��ȡ������Ϣ", mlngPatiID)
    If mrsInfo.EOF Then GoTo NotFoundPati:
    
    If mblnCliniqueRoomPay = False Then
        lbl����.Caption = Nvl(mrsInfo!����)
        lbl�Ա�.Caption = "�Ա�:" & Nvl(mrsInfo!�Ա�)
        lblMZH.Caption = "�����:" & Nvl(mrsInfo!�����)
        '74309:���ϴ���2014-7-7������������ʾ��ɫ����
        Call SetPatiColor(lbl����, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!����), &HFF0000, vbRed))
    End If
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Set mrsInfo = New ADODB.Recordset
    Call SaveErrLog
    Exit Function
NotFoundPati:
    MsgBox "������Ϣδ�ҵ�,����!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    Set mrsInfo = New ADODB.Recordset
End Function
Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-06-21 11:08:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long
    Dim strPayType As String, varData As Variant, varTemp As Variant, i As Long
    j = 0
    '��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    strPayType = GetAvailabilityCardType: varData = Split(strPayType, ";")
    Set mcolPayMode = New Collection
    With cbo֧����ʽ
        .Clear: j = 0
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                mcolPayMode.Add varTemp, "K" & j
                cbo֧����ʽ.AddItem varTemp(1)
                cbo֧����ʽ.ItemData(cbo֧����ʽ.NewIndex) = Val(varTemp(2))
                j = j + 1
            End If
        Next
    End With
    If cbo֧����ʽ.ListCount > 0 And cbo֧����ʽ.ListIndex < 0 Then cbo֧����ʽ.ListIndex = 0
    
End Sub
Private Function CheckPayIsEnough(Optional blnYesNo As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ԥ�����Ƿ�֧��
    '����:���˺�
    '����:2011-06-21 11:29:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Val(lblʣ�����.Tag) < Val(lbl�Ը��ϼ�.Tag) Then
        If blnYesNo Then
            If cbo֧����ʽ.Enabled Then
                '����������֧��,���Բ�����
                CheckPayIsEnough = True: Exit Function
            End If
            If MsgBox("ע��:" & vbCrLf & "   Ԥ�������֧�����η���,���ֵ" & vbCrLf & "   ���β����Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                CheckPayIsEnough = True: Exit Function
            End If
            Exit Function
        Else
            '��Ҫ�ſ�����֧����ʽ,����Ƿ���
            '79621:���ϴ�,2014/11/14,�Խ���ʽ������
            If Val(Val(lblʣ�����.Tag) + Val(txt���.Text)) < Val(lbl�Ը��ϼ�.Tag) Then
                Call MsgBox("ע��:" & vbCrLf & "   Ԥ�������֧�����η���,���ֵ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
                Exit Function
            End If
        End If
    End If
    CheckPayIsEnough = True
End Function
Private Sub SetCtlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enable����
    '����:���˺�
    '����:2011-06-21 11:19:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cbo֧����ʽ.Enabled = cbo֧����ʽ.ListCount > 0
    txt���.Enabled = cbo֧����ʽ.Enabled And cbo֧����ʽ.ListCount > 0
    cbo֧����ʽ.Visible = cbo֧����ʽ.ListCount > 0
    lbl֧����ʽ.Visible = cbo֧����ʽ.ListCount > 0
    txt���.Visible = cbo֧����ʽ.ListCount > 0
    lbl���.Visible = cbo֧����ʽ.ListCount > 0
    txt��Ԥ��.Enabled = mCurCarge.dbl����Ԥ�� <> 0
End Sub
Private Sub txt��Ԥ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Val(txt��Ԥ��) = 0 Then
        If txt���.Enabled And txt���.Visible Then txt���.SetFocus
        Exit Sub
    End If
    mblnOK_Click = False
    If CheckPrepayMoneyIsValied = False Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txt��Ԥ��_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt��Ԥ��, KeyAscii, m���ʽ)
End Sub

Private Function CheckPrepayValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ԥ��������Ƿ���Ч
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-14 22:30:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mblnOk Then Exit Function
    If txt��Ԥ��.Text = "" Then
        txt��Ԥ��.Text = "0.00"
    ElseIf Not IsNumeric(txt��Ԥ��.Text) And txt��Ԥ��.Text <> "" Then
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Exit Function
    ElseIf Val(txt��Ԥ��.Text) < 0 Then
        MsgBox "Ԥ��������Ϊ����", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Exit Function
    ElseIf Val(txt��Ԥ��.Text) > 0 And mCurCarge.dbl��ǰδ�� < 0 Then
        MsgBox "��ǰӦ�����Ϊ��ʱ����ʹ��Ԥ��", vbInformation, gstrSysName
        txt��Ԥ��.Text = "0.00"
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��:   Exit Function
    '79621:���ϴ�,2014/11/14,�Խ���ʽ������
    ElseIf RoundEx(Val(txt��Ԥ��.Text), 2) > RoundEx(mCurCarge.dbl����Ԥ��, 2) Then
        MsgBox "Ԥ�������ܳ������˵�Ԥ�����:" & Format(mCurCarge.dbl����Ԥ��, "0.00") & " ��", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Exit Function
    ElseIf RoundEx(Val(txt��Ԥ��.Text), 2) > RoundEx(mCurCarge.dbl��ǰδ��, 2) And Val(txt��Ԥ��.Text) <> 0 Then
        MsgBox "Ԥ�������ܴ���Ӧ�����:" & Format(mCurCarge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Exit Function
    Else
        txt��Ԥ��.Text = Format(Val(txt��Ԥ��.Text), "0.00")
    End If
    CheckPrepayValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub txt��Ԥ��_Validate(Cancel As Boolean)
    If lblԤ���.Tag = "1" Or mlngPatiID = 0 Then Exit Sub
    If Val(txt��Ԥ��.Tag) = Val(txt��Ԥ��.Text) Then Exit Sub
    If CheckPrepayValied = False Then Cancel = True: Exit Sub
End Sub

Private Sub txt���_GotFocus()
    txt���.Text = Format(RoundEx(Val(lbl�Ը��ϼ�.Tag), 2) - RoundEx(Val(txt��Ԥ��), 2), "####0.00;-###0.00")
    If txt���.Text < 0 Then txt���.Text = ""
    zlControl.TxtSelAll txt���
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt��Ԥ��, KeyAscii, m���ʽ)
    picPayMode.Tag = ""
    If KeyAscii <> 13 Then Exit Sub
    mblnOK_Click = False
    If Val(txt���.Text) = 0 Then txt���.Text = "0.00"
    If txt���.Text <> "0.00" Then
        If RoundEx(mCurCarge.dbl�������Ѻϼ� - Val(txt��Ԥ��.Text) - Val(txt���.Text), 7) <> 0 Then
            MsgBox "���׽���������,����������(" & Format(RoundEx(mCurCarge.dbl�������Ѻϼ� - Val(txt��Ԥ��.Text), 7), "0.00") & ")��", vbInformation, gstrSysName
           If txt���.Enabled And txt���.Visible Then txt���.SetFocus
           zlControl.TxtSelAll txt���
            picPayMode.Tag = "1"
        End If
        Call cmdOK_Click
        Exit Sub
    End If
End Sub

Private Sub txt���_Validate(Cancel As Boolean)
    '79621:���ϴ�,2014/11/14,�Խ���ʽ������
    If RoundEx(Val(txt��Ԥ��) + Val(txt���.Text), 2) > RoundEx(Val(lbl�Ը��ϼ�.Tag), 2) Then
        If picPayMode.Tag <> "1" Then MsgBox "���뱾��֧����Ԥ�������" & cbo֧����ʽ & "֧���ĺϼƴ����˱��ν�����úϼ�,���ܼ���!", vbInformation + vbOKOnly, gstrSysName
        txt���.Text = Format(RoundEx(Val(lbl�Ը��ϼ�.Tag), 2) - RoundEx(Val(txt��Ԥ��), 2), "####0.00;-###0.00")
        If txt���.Text < 0 Then txt���.Text = ""
        If txt���.Enabled And txt���.Visible Then txt���.SetFocus
        zlControl.TxtSelAll txt���
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub lblԤ���_Change()
    lblԤ���.Tag = ""
End Sub
Private Function IsCheckThreeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������׽�������Ƿ�Ϸ�
    '����:�Ϸ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-15 00:03:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mblnOk Then Exit Function
    If Val(txt���) = 0 Then
        MsgBox "δ���뽻�׽��,����!", vbInformation + vbOKOnly
        If txt���.Enabled And txt���.Visible Then txt���.SetFocus
        zlControl.TxtSelAll txt���
         Exit Function
    End If
    If Not IsNumeric(txt���.Text) And txt���.Text <> "" Then
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        If txt���.Enabled And txt���.Visible Then txt���.SetFocus
        zlControl.TxtSelAll txt���: Exit Function
    ElseIf Val(txt���.Text) < 0 Then
        MsgBox "���׽���Ϊ����", vbInformation, gstrSysName
        If txt���.Enabled And txt���.Visible Then txt���.SetFocus
        zlControl.TxtSelAll txt���: Exit Function
    '79621:���ϴ�,2014/11/14,�Խ���ʽ������
    ElseIf RoundEx(Val(txt���.Text), 2) > RoundEx(mCurCarge.dbl��ǰδ��, 2) And Val(txt���.Text) <> 0 Then
        MsgBox "���׽��ܴ��ڱ���δ�����:" & Format(mCurCarge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
        If txt���.Enabled And txt���.Visible Then txt���.SetFocus
        zlControl.TxtSelAll txt���: Exit Function
    End If
    IsCheckThreeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlGetClassMoney(ByVal strNos As String, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    'ʵ�ս��:'����:50339
    strSQL = "" & _
    "   Select  /*+ rule */  A.�շ����,nvl(sum(ʵ�ս��) ,0) as  ���   " & _
    "   From ������ü�¼ A,Table(f_str2List([1])) B " & _
    "   Where A.NO=B.Column_value and A.��¼����=1 and A.��¼״̬=0 " & _
    "   Group by A.�շ����"
    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlBrush���ѿ�(ByVal dblMoney As Double, ByVal rsClassMoney As ADODB.Recordset, _
    ByVal str������Դ As String, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ѿ�ˢ��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-15 09:54:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As Collection
    Dim frmInput As New frmInputPass
    Set mcllSquareBalance = Nothing
    If mobjPatiCardObject Is Nothing Then
        MsgBox "��ǰ֧�����ӿ�����,����", vbOKOnly, gstrSysName
        Exit Function
    End If
    zlBrush���ѿ� = frmInput.zlBrushPay(Me, mlngModule, mobjPatiCardObject, rsClassMoney, _
        mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, Nvl(mrsInfo!����), Nvl(mrsInfo!�Ա�), _
        Nvl(mrsInfo!����), dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, , True, , False, cllBalance, _
        False, True, str������Դ, lng����ID)
    Set frmInput = Nothing
    Set mcllSquareBalance = cllBalance

End Function
Private Function BrushCardThreeSwapCheck(ByVal strNos As String, _
    ByVal dblMoney As Double, ByVal str������Դ As String, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����֤
    '���:strNos -����֧���ĵ��ݺ�
    '       dblMoney-֧�����ܽ��
    '����:����true,���򷵻�False
    '����:���˺�
    '����:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim frmMain As Object
    
    On Error GoTo errHandle
    If mintCurType = 2 Then BrushCardThreeSwapCheck = True: Exit Function
    If mCurCardPay.lngҽ�ƿ����ID = 0 Then BrushCardThreeSwapCheck = True: Exit Function
    If mblnCliniqueRoomPay = False Then
        If IsCheckThreeValied = False Then Exit Function
        Set frmMain = Me
    Else
        Set frmMain = mfrMain
    End If
    
    '����ˢ������
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln���ѿ� As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl��� As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String) As Boolean
    If mCurCardPay.bln���ѿ� And mCurCardPay.bln���ƿ� Then
        '����:50339
        If zlGetClassMoney(strNos, rsMoney) = False Then Exit Function
        '�϶��Ǵ������ƿ�
        If zlBrush���ѿ�(dblMoney, rsMoney, str������Դ, lng����ID) = False Then Exit Function
    Else
        '    zlBrushCard(frmMain As Object, _
        '        ByVal lngModule As Long, _
        '        ByVal lngCardTypeID As Long, _
        '        ByVal strPatiName As String, ByVal strSex As String, _
        '        ByVal strOld As String, ByVal dbl��� As Double, _
        '        Optional ByRef strCardNo As String, _
        '        Optional ByRef strPassWord As String
        If mobjCardPay.zlBrushCard(frmMain, mlngModule, mCurCardPay.lngҽ�ƿ����ID, _
         Nvl(mrsInfo!����), Nvl(mrsInfo!�Ա�), Nvl(mrsInfo!����), dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������) = False Then Exit Function
    End If
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If mobjCardPay.zlPaymentCheck(frmMain, mlngModule, mCurCardPay.lngҽ�ƿ����ID, _
          mCurCardPay.strˢ������, dblMoney, strNos, strXMLExpend) = False Then Exit Function
    BrushCardThreeSwapCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function interfacePayMoney(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�õ������ӿ�֧��(�������ѿ�,���п���������)
    '���:strCardNo-֧���Ŀ���
    '����:֧���ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2011-06-22 12:01:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim lng����ID As Long, strCardPass As String, lng�����ID As Long
    Dim dbl��� As Double
    
    On Error GoTo errHandle
    dbl��� = RoundEx(Val(txt���), 2)
    '79621:���ϴ�,2014/11/14,�Խ���ʽ������
    If RoundEx(Val(txt��Ԥ��) + dbl���, 2) > RoundEx(Val(lbl�Ը��ϼ�.Tag), 2) Then
        MsgBox "���뱾�οۿ�������˱��ν�����úϼ�,���ܼ���!", vbInformation + vbOKOnly, gstrSysName
        If txt���.Enabled And txt���.Visible Then txt���.SetFocus
        Exit Function
    End If
    If RoundEx(Val(txt��Ԥ��) + dbl���, 2) <> RoundEx(Val(lbl�Ը��ϼ�.Tag), 2) Then
        MsgBox "���뱾�οۿ���С���˱��ν�����úϼ�,���ܼ���!", vbInformation + vbOKOnly, gstrSysName
        If txt���.Enabled And txt���.Visible Then txt���.SetFocus
        Exit Function
    End If
    If mrsInfo Is Nothing Then
        MsgBox "�������벡��!", vbInformation + vbOKOnly, gstrSysName
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "�������벡��!", vbInformation + vbOKOnly, gstrSysName
        If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus
        Exit Function
    End If
    If Val(lblԤ���.Tag) = 0 And Val(txt��Ԥ��.Text) <> 0 Then
        'δ������֤,��Ҫ����ȷ����������
        MsgBox "ʹ��Ԥ���������ˢ��ȷ�����ѣ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    interfacePayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub setDefaultPrepayMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡԤ�����
    '����:���˺�
    '����:2011-08-13 17:21:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurCarge
         txt��Ԥ��.Text = "0.00"
        If .dbl����Ԥ�� <> 0 Then
            txt��Ԥ��.Text = Format(IIf(.dbl����Ԥ�� > .dbl��ǰδ��, .dbl��ǰδ��, .dbl����Ԥ��), "###0.00;###0.00;0.00;0.00")
        End If
    End With
End Sub
Private Function ExecuteCliniqueRoomPay() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���֧��
    '����:���֧���ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-01-14 17:28:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intTYPE As Integer, objSquareCard As Object
    On Error GoTo errHandle
    
    '��������,���ӽ��㿨�Ľ���
    Err = 0: On Error Resume Next
    Set objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        MsgBox "���㿨����zl9CardSquare.clsCardSquare����ʧ�ܣ�", vbInformation, gstrSysName
        Err = 0: On Error GoTo 0:      Exit Function
    End If
    If objSquareCard Is Nothing Then Exit Function
    '��װ�˽��㿨�Ĳ���
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '����:
    '����:   True:���óɹ�,False:����ʧ��
    '����:���˺�
    '����:2009-12-15 15:16:22
    'HIS����˵��.
    '   1.���������շ�ʱ���ñ��ӿ�
    '   2.����סԺ����ʱ���ñ��ӿ�
    '   3.����Ԥ����ʱ
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If objSquareCard.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False) = False Then
         '��ʼ�������ɹ�,����Ϊ�����ڴ���
         Exit Function
    End If
    
    '��ȡ������Ϣ
    If GetPatient = False Then Exit Function
     
    '��������
    If SaveCharge = False Then mstrPrintNO = "": Exit Function
    Call PrintBill
    
    '��ҽһ��ͨд����85950
    Call WriteInforToCard(Me, mlngModule, mstrPrivs, objSquareCard, 0, mstrPrintNO)
    Set objSquareCard = Nothing
    
    ExecuteCliniqueRoomPay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CreateLocalTypeObject(ByVal lngCardTypeID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����������
    '���:lngCardTypeID-�����ID
    '����:
    '����:�����ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-01-14 18:19:38
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objCard As clsCard, blnReturn As Boolean
    On Error GoTo errHandle
    '����ָ���Ķ���
    Set mobjCardPay = Nothing
    
    blnReturn = zlGetCardProperty(lngCardTypeID, False, objCard)
    If blnReturn = False Or objCard Is Nothing Then
        MsgBox "ע��:" & vbCrLf & _
                      "      ��ҽ�ƿ�����У�δ�ҵ�ָ���������ʻ���֧�ֵ���� " & vbCrLf & _
                      "���ܸ����δ����,�����ҽ�ƿ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.�Ƿ�����ʻ� = False Then
        MsgBox objCard.���� & "δ���������ʻ�,�����ҽ�ƿ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.���㷽ʽ = "" Then
        MsgBox objCard.���� & "δ���ý��㷽ʽ,�����ҽ�ƿ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If objCard.�ӿڳ����� = "" Then
        MsgBox objCard.���� & "δ���������ӿ���֧�ֵĲ���,�����ҽ�ƿ�����", vbInformation, gstrSysName
        Exit Function
    End If
    With mCurCardPay
       .lngҽ�ƿ����ID = objCard.�ӿ����
       .bln���ѿ� = objCard.���ѿ�
       .str���㷽ʽ = objCard.���㷽ʽ
       .str���� = objCard.����
       .strˢ������ = ""
       .strˢ������ = ""
       .lngID = 0
       .strNO = ""
       .bln�������� = False
       .intҽ�ƿ����� = 0
       .bln���� = False
       .bln֧Ʊ = False
       .blnOneCard = False
       .bln���ƿ� = False
       .int���� = 0
    End With
    Err = 0: On Error Resume Next
    
    If zlGetCardObj(Me, objCard.�ӿ����, objCard.���ѿ�, mobjPatiCardObject, , True) = False Then
        Set mobjPatiCardObject = Nothing
        Set mobjCardPay = Nothing
        Exit Function
    End If
    
    Set mobjCardPay = mobjPatiCardObject.CardObject
    If Err <> 0 Then
        MsgBox "δ�ҵ�" & mCurCardPay.str���� & "����Ӧ�Ĳ���,����", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjCardPay Is Nothing Then Exit Function
    CreateLocalTypeObject = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function CliniqueRoomPayValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���֧�����
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-01-17 16:36:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mblnCliniqueRoomPay = False Then CliniqueRoomPayValied = True: Exit Function
    If mbytBillType <> 1 Then   'ֻ����շѵ�
        MsgBox "ע��:" & vbCrLf & "    ���֧��ʱ����������Լ��ʵ������ʵĽ���֧����", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If mlngCardTypeID = 0 Then
        MsgBox "ע��:" & vbCrLf & "    ���֧��ʱҪ��ָ��һ�������ʻ�֧�����,����ϵͳ����Ա��ϵ��", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
 
    '���󴴽�ʧ�ܵ�,������֧��
    If Not CreateLocalTypeObject(mlngCardTypeID) Then Exit Function
    CliniqueRoomPayValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPayDrugWindow(ByVal lng����ID As Long, ByVal dt�շ�ʱ�� As Date, _
    ByVal cllDept As Collection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ����䷢ҩ����
    '���:lng����ID-����ID
    '     dt�շ�ʱ��-�շ�ʱ��
    '     cllDept-����ִ�в���:array(�շ����,ִ�в���ID,��ҩ����)
    '���أ���ҩ��������
    '���ƣ����ϴ�
    '���:strNO
    'ʱ�䣺2014-6-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��ҩ���� As String, strPayDrugWins As String
    Dim str���� As String, str�д� As String, str�ɴ� As String
    Dim i As Long, varData As Variant
    
    Err = 0: On Error GoTo Errhand:
    strPayDrugWins = ""
    For i = 1 To cllDept.count
        varData = cllDept(i)
        str��ҩ���� = varData(2)
        If str��ҩ���� = "" Then
            '�жϵ�ǰ�����Ƿ������ִͬ�в��ŵ�δ��ҩƷ���������򷵻�δ��ҩƷ�ķ�ҩ����
            str��ҩ���� = Getδ��ҩƷ��ҩ����(lng����ID, Val(varData(1)))
            If str��ҩ���� = "" Then str��ҩ���� = GetDrugWindow(Val(varData(1)), Trim(varData(0)))
            If str��ҩ���� = "" Then
                str��ҩ���� = Get��ҩ����(dt�շ�ʱ��, Val(varData(1)), Trim(varData(0)), str����, str�ɴ�, str�д�)
            End If
        End If
        If InStr(1, strPayDrugWins & ";", ";" & Val(varData(1)) & "|") = 0 Then
            strPayDrugWins = strPayDrugWins & ";" & Val(varData(1)) & "|" & str��ҩ����
        End If
    Next
    GetPayDrugWindow = strPayDrugWins
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function GetDrugWindow(ByVal lngҩ��ID As Long, ByVal str��� As String) As String
'���ܣ���ȡȱʡ�ķ�ҩ����,�������ָ����ȱʡ,����ָ��Ϊ׼,����,����ǻ��۵�,���Ե�һҩƷ�еĴ���Ϊ׼,��������������ͬҩƷ�Ĵ���Ϊ׼
'������intPage=��¼���ĵ��ݱ��
'˵������Ҫ���ڶ൥���շ�ʱ����ͬ����ҩƷ���ܶ�̬���䵽ͬһҩ�����������ǵĴ���ҲӦ��ͬ����ǿ��ָ���ĳ���
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim p As Integer, i As Integer, varData As Variant, varTemp As Variant
    Err = 0: On Error GoTo errH:
    GetDrugWindow = GetDefaultWindow(str���, lngҩ��ID)
    If GetDrugWindow = "" Then Exit Function
    strSQL = "Select ���� From ��ҩ���� Where �ϰ��=1 And ҩ��ID=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҩ��ID, GetDrugWindow)
    If rsTmp.EOF Then GetDrugWindow = ""
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Getδ��ҩƷ��ҩ����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As String
    '-------------------------------------------------------------------------
    '���ܣ��жϵ�ǰ�����Ƿ������ִͬ�в��ŵ�δ��ҩƷ���������򷵻�δ��ҩƷ�ķ�ҩ����
    '���أ���������ִͬ�в��ŵ�δ��ҩƷ���򷵻�δ��ҩƷ�ķ�ҩ���ڣ����򷵻ؿ�
    '���ƣ�Ƚ����
    '���ڣ�2014-04-09
    '���⣺71902
    '˵����
    '   ͬһ���˲��˲�ͬʱ��ζ��ŵ����շѣ�����ͬһ����ҩ���ڣ����㲡��ȡҩ
    '-------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select ��ҩ����" & vbNewLine & _
            "From δ��ҩƷ��¼" & vbNewLine & _
            "Where ���� = 8 And ��ҩ���� Is Not Null And ����id = [1] And �ⷿid = [2]" & vbNewLine & _
            "Order By ���շ� Desc, �������� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����δ��ҩƷ��ҩ����", lng����ID, lngִ�в���ID)
    
    If Not rsTemp.EOF Then
        Getδ��ҩƷ��ҩ���� = Nvl(rsTemp!��ҩ����)
    End If
    rsTemp.Close: Set rsTemp = Nothing
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get��ҩ����(ByVal Curdate As Date, ByVal lngҩ��ID As Long, ByVal str��� As String, _
    str���� As String, str�ɴ� As String, str�д� As String) As String
'���ܣ���ȡҩƷ��Ӧ�ķ�ҩ����
'������lngҩ��ID=ִ�в���ID,curDate=��ǰʱ��
'˵������ͬһ������ҩ���ķ�ҩ������ƽ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ָ��ʱ�̶�����(ָ����ָû�ж�Ӧҩ���ϰ�ʱָ��)
    Select Case str���
        Case "5"
            If str���� <> "" Then
                Get��ҩ���� = str����
            ElseIf mlng��ҩ�� > 0 Then
                Get��ҩ���� = GetDefaultWindow(str���, lngҩ��ID)
                str���� = Get��ҩ����
            End If
        Case "6"
            If str�ɴ� <> "" Then
                Get��ҩ���� = str�ɴ�
            ElseIf mlng��ҩ�� > 0 Then
                Get��ҩ���� = GetDefaultWindow(str���, lngҩ��ID)
                str�ɴ� = Get��ҩ����
            End If
        Case "7"
            If str�д� <> "" Then
                Get��ҩ���� = str�д�
            ElseIf mlng��ҩ�� > 0 Then
                Get��ҩ���� = GetDefaultWindow(str���, lngҩ��ID)
                str�д� = Get��ҩ����
            End If
    End Select
    
    
    If Get��ҩ���� <> "" Then
        strSQL = "Select ���� From ��ҩ���� Where �ϰ��=1 And ҩ��ID=[1] And ����=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", lngҩ��ID, Get��ҩ����)
        If rsTmp.EOF Then Get��ҩ���� = ""
        Exit Function
    End If
    
    '��̬�����ϰ�ķ�ר�Ҵ���,98876
    strSQL = "Select Zl_Get��ҩ����([1],[2],[3]) As ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ҩ����", lngҩ��ID, mbytAssign, Curdate)
    If Not rsTmp.EOF Then
        Get��ҩ���� = Nvl(rsTmp!����)
    End If
    
    If Get��ҩ���� <> "" Then
        Select Case str���
            Case "5"
                str���� = Get��ҩ����
            Case "6"
                str�ɴ� = Get��ҩ����
            Case "7"
                str�д� = Get��ҩ����
        End Select
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDefaultWindow(ByVal str��� As String, ByVal lngҩ��ID As Long) As String
'����:��ȡȱʡ��ҩ����������
    Dim strTmp As String, i As Long, arrTmp As Variant, arrWin As Variant
    
    Select Case str���
        Case "5"
            If InStr(mstr����, ":") > 0 Then '������û�д�ҩ��ID
                 strTmp = mstr����
            ElseIf mlng��ҩ�� > 0 And mstr���� <> "" Then
                strTmp = mlng��ҩ�� & ":" & mstr����
            End If
        Case "6"
            If InStr(mstr�ɴ�, ":") > 0 Then
                 strTmp = mstr�ɴ�
            ElseIf mlng��ҩ�� > 0 And mstr�ɴ� <> "" Then
                 strTmp = mlng��ҩ�� & ":" & mstr�ɴ�
            End If
        Case "7"
            If InStr(mstr�д�, ":") > 0 Then
                 strTmp = mstr�д�
            ElseIf mlng��ҩ�� > 0 And mstr�д� <> "" Then
                 strTmp = mlng��ҩ�� & ":" & mstr�д�
            End If
    End Select
    
    If strTmp <> "" Then
        arrTmp = Split(strTmp, ",")
        strTmp = ""
        For i = 0 To UBound(arrTmp)
            arrWin = Split(arrTmp(i), ":")
            Select Case str���
                Case "5"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "6"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
                Case "7"
                    If arrWin(0) = lngҩ��ID Then strTmp = arrWin(1): Exit For
            End Select
        Next
    End If
    GetDefaultWindow = strTmp
End Function
Private Function SetCurBalanceSQL(ByVal bytType As Byte, ByVal lng����ID As Long, _
    ByVal dblPayMoney As Double, ByVal dbl��Ԥ�� As Double, ByVal dbl�ɿ� As Double, ByVal dbl�Ҳ� As Double, _
    ByVal dbl�������� As Double, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ�ǰ�����SQL��cllpro����
    '���:  bytType-1-�����ӿ�֧��;2-���ѿ�֧��;0-����
    '       dblPayMoney-��ǰ֧�����
    '       dbl��Ԥ��-Ԥ����֧��
    '       dbl�ɿ�-Ͷ����Ч
    '       dbl�Ҳ�-Ͷ����Ч
    '       dblԤ���-�������Ԥ����Ļ�,����Ԥ������
    '       dbl��������-���β���������
    '����:cllPro-ִ�й���
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-15 15:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL  As String, str�շѽ��� As String
    Dim dblԤ��� As Double, lngCardTypeID As Long, j As Long
    
    
    On Error GoTo errHandle
    
    ' Zl_�����շѽ���_Modify
    strSQL = "Zl_�����շѽ���_Modify("
    '  ------------------------------------------------------------------------------------------------------------------------------
    '  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
    '  --��������_In:
    '  --   0-��ͨ�շѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����֧Ʊ��_In:������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����֧Ʊ��_In:������
    '  --   3-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."  ���ѿ�ID:Ϊ��ʱ,���ݿ����Զ���λ
    '  --     �ڳ�Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  -- ��Ԥ��_In: ���ڳ�Ԥ��ʱ,����
    '  -- �����_In:��������ʱ,����
    '  -- ��ɽ���_In:1-����շ�;0-δ����շ�
    '  ------------------------------------------------------------------------------------------------------------------------------
    ' bytType- 1-�����ӿ�֧��;2-���ѿ�֧��,3�ʻ�֧��
    Select Case bytType
    Case 1  '1-�����ӿ�֧��
        strSQL = strSQL & "1" & ","
        '"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        str�շѽ��� = mCurCardPay.str���㷽ʽ
        str�շѽ��� = str�շѽ��� & "|" & dblPayMoney
        str�շѽ��� = str�շѽ��� & "|" & " "
        str�շѽ��� = str�շѽ��� & "|" & " "
        lngCardTypeID = mCurCardPay.lngҽ�ƿ����ID
    Case 2 ' 2-���ѿ�֧��
        strSQL = strSQL & "3" & ","
        If mcllSquareBalance Is Nothing Then Exit Function
        If mcllSquareBalance.count = 0 Then Exit Function
        '�����ID|����|���ѿ�ID|���ѽ��||."
        '���ѿ�ID���Բ���,��Ϊ0ʱ,�Կ����Զ�����
        str�շѽ��� = ""
        For j = 1 To mcllSquareBalance.count
            ' array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
            str�շѽ��� = str�շѽ��� & "||" & Val(mcllSquareBalance(j)(0))
            str�շѽ��� = str�շѽ��� & "|" & mcllSquareBalance(j)(3)
            str�շѽ��� = str�շѽ��� & "|" & Val(mcllSquareBalance(j)(1))
            str�շѽ��� = str�շѽ��� & "|" & Val(mcllSquareBalance(j)(2))
        Next
        If str�շѽ��� <> "" Then str�շѽ��� = Mid(str�շѽ���, 3)
        lngCardTypeID = mCurCardPay.lngҽ�ƿ����ID
    Case Else
        strSQL = strSQL & "0" & ","
    End Select
    '    ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & lng����ID & ","
    '    ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & mCurCardPay.lng����ID & ","
    '    ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & IIf(str�շѽ��� = "", "NULL", "'" & str�շѽ��� & "'") & ","
    '    ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & IIf(dbl��Ԥ�� <> 0, dbl��Ԥ��, "NULL") & ","
    '    ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '    �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(lngCardTypeID = 0, "NULL", lngCardTypeID) & ","
    '    ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.strˢ������ <> "", "'" & mCurCardPay.strˢ������ & "'", "NULL") & ","
    '    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL,"
    '    �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & dbl�ɿ� & ","
    '    �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "" & dbl�Ҳ� & ","
    '    �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '    -- �����_In:��������ʱ,����
    strSQL = strSQL & "" & dbl�������� & ","
    '    ��ɽ���_In Number:=0
    '    -- ��ɽ���_In:1-����շ�;0-δ����շ�
    strSQL = strSQL & "1,"
    '  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '79868,Ƚ����,2015-06-10,ʹ�ò��˼���Ԥ��
    '  ��Ԥ������ids_In Varchar2:=Null
    strSQL = strSQL & "'" & lng����ID & "," & mstr����IDs & "')"
    zlAddArray cllPro, strSQL
    SetCurBalanceSQL = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ�Լ��������
    '����:���ϴ�
    '����:2014-10-29
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim strInfo As String, lngPatient As Long
    If gblnLED = False Then Exit Sub
    
    On Error GoTo Errhand
    zl9LedVoice.Reset mscCom
    strInfo = lbl����.Caption
    If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!�Ա� & " " & mrsInfo!����: lngPatient = Val("" & mrsInfo!����ID)
    zl9LedVoice.DisplayPatient strInfo, lngPatient
    '�����ܶ�:������Ҫ֧���Ľ�Ԥ�����:���˵�ǰ��Ԥ�����
    Call zl9LedVoice.DisplayBank( _
            "�����ܶ�:" & mCurCarge.dbl�������Ѻϼ� & "Ԫ" & _
            IIf(mCurCarge.dblԤ����� = 0, "", ",Ԥ�����:" & mCurCarge.dblԤ����� & "Ԫ"))
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CreateDrugPacker()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ҩ��(�Զ���ҩ��)
    '����:���˺�
    '����:2014-06-05 15:30:47
    '˵��:bug-51510
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnDrugPacker = False: mblnDrugMachine = False
    
    '0-�������շѻ���ʵ�,1-�շѼ�¼;2-���ʼ�¼
'    If mbytBillType = 2 Then Exit Sub
    
    If mblnDrugMachine Or mblnDrugPacker Then Exit Sub

    Err = 0: On Error Resume Next
    If Val(zlDatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 Then
        '�����½ӿ�
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    If mblnDrugMachine = False Then
        '�ɲ���
        Err = 0
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err = 0 Then mblnDrugPacker = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        'Ȩ�޼��
        strPrivs = GetPrivFunc(glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))
        If InStr(";" & strPrivs & ";", ";����;") >= 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, objComLib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    ElseIf mblnDrugPacker Then
        mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
    End If
End Sub
