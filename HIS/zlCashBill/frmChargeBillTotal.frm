VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmChargeBillTotal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picReturnBill 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   5370
      ScaleHeight     =   2685
      ScaleWidth      =   3630
      TabIndex        =   7
      Top             =   2730
      Width           =   3630
      Begin VSFlex8Ctl.VSFlexGrid vsReturnBill 
         Height          =   1800
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmChargeBillTotal.frx":0000
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
   End
   Begin VB.PictureBox picBillInfor 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   570
      ScaleHeight     =   1470
      ScaleWidth      =   3630
      TabIndex        =   1
      Top             =   4365
      Width           =   3630
      Begin VSFlex8Ctl.VSFlexGrid vsBill 
         Height          =   870
         Left            =   0
         TabIndex        =   5
         Top             =   90
         Width           =   1860
         _cx             =   3281
         _cy             =   1535
         Appearance      =   2
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillTotal.frx":007A
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
   End
   Begin VB.PictureBox picChargeInfor 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   450
      ScaleHeight     =   2685
      ScaleWidth      =   3630
      TabIndex        =   0
      Top             =   1170
      Width           =   3630
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1170
         TabIndex        =   4
         Top             =   2010
         Width           =   2280
      End
      Begin VSFlex8Ctl.VSFlexGrid vsChagre 
         Height          =   1800
         Left            =   315
         TabIndex        =   2
         Top             =   15
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeBillTotal.frx":00F4
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
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "�տ�ϼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   405
         TabIndex        =   3
         Top             =   2070
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRptPrint 
      Height          =   1560
      Left            =   765
      TabIndex        =   6
      Top             =   6105
      Visible         =   0   'False
      Width           =   3735
      _cx             =   6588
      _cy             =   2752
      Appearance      =   2
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmChargeBillTotal.frx":016E
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   45
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeBillTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long, mstrPrivs As String
Private Enum TotalType
    EM_�շ�Ա���� = 1
    EM_С���տ� = 2
    EM_С������ = 3
    EM_�����տ� = 4
    EM_�����տ�_���շ�Ա = 5
End Enum
'1-�շ�Ա���ʣ�2-С���տ�;3-С������;4-�����տ�(����շ�Ա��������տ�)������տ��ѯ;5-�����տ�(����Է��շ�Ա�տ�)��
Private mbytType As TotalType
Private mlngChargeRollingID As Long '����ID���տ�ID(����mbytType)������
Private mdtStartDate As Date, mdtendDate As Date '���ʵĿ�ʼʱ������ʱ��
Private mblnDel As Boolean '�Ƿ����ϼ�¼
Public mrsList As ADODB.Recordset '�տ���ؼ�¼
Public mrsListBill As ADODB.Recordset 'Ʊ����ؼ�¼
Private mrsBalance As ADODB.Recordset '����Ա���
Private mblnHideFilter As Boolean '�Ƿ����ع�������
Private mlngErrorRow As Long
Private mdblRemain As Double
Private mblnOlnyView As Boolean '���ܲ鿴����,���ܱ༭ʵ��Ʊ��
Private Enum mPaneIndex
    EM_PN_ChargeTotal = 260102  '�տ����
    EM_PN_BillTotal = 260103    'Ʊ�ݻ���
    EM_PN_BackFeeBill = 260104  '�˷�Ʊ��
    EM_PN_ReprintBill = 260105  '�ش�Ʊ��
End Enum
Private mlngCashRow As Long '�ֽ���ָ������
Private mbytFontSize As Byte
Private mstrPersonName  As String '��ǰ�շ���Ա
Private mstrRollingType As String  '�������

Public Sub ClearData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2013-09-12 11:09:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Call InitGrid
    mlngCashRow = 0
    txtTotal.Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
Public Function LoadChargeAndBillTotalData(ByVal frmMain As Object, _
      ByVal lngModule As Long, ByVal strPrivs As String, _
      ByVal bytType As Byte, ByVal lngChargeRollingID As Long, _
      Optional ByVal dtStartDate As Date, Optional ByVal dtEndDate As Date, _
      Optional blnOlnyView As Boolean = True, _
      Optional ByVal blnDel As Boolean = False, _
      Optional strPersonName As String = "", _
      Optional strRollingType As String) As Boolean
    '-------------------------------------------------------------------------------------------------
    '����:�շ�Ա���˽ӿ�
    '���:frmMain-���õ�������
    '    lngModule-ģ���
    '    strPrivs-Ȩ�޴�
    '����bytType:1-�շ�Ա���ʣ�2-С���տ�;3-С������;
    '            4-�����տ�(����շ�Ա��������տ�)������տ��ѯ;
    '            5-�����տ�(����Է��շ�Ա�տ�)��
    '    lngChargeRollingID -�շ�Ա������ID
    '    dtStartDate-��ѡ����,��ʼ����ʱ��,lngChargeRollIngID=0ʱ�����봫��
    '    dtEndDate-��ѡ��������������ʱ��,lngChargeRollIngID=0ʱ�����봫��
    '    blnOlnyView-���ܲ鿴(���ܱ���Ʊ�ݺ���)
    '    blnDel-�Ƿ����ϼ�¼
    '    strPersonName-ָ�����շ�Ա(Ϊ��ʱ,Ϊ��ǰ����Ա)
    '    strRollingType-�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨)
    '����:���ݼ��سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-08-13 10:31:00
    '-------------------------------------------------------------------------------------------------
    mbytType = bytType:  mlngChargeRollingID = lngChargeRollingID
    mdtStartDate = dtStartDate: mdtendDate = dtEndDate
    mlngMode = lngModule: mstrPrivs = strPrivs
    mblnDel = blnDel
    mblnOlnyView = blnOlnyView
    vsChagre.Editable = IIf(mblnOlnyView, flexEDNone, flexEDKbdMouse)
    mstrPersonName = IIf(strPersonName = "", UserInfo.����, strPersonName)
    mstrRollingType = strRollingType
    
    If Not mblnOlnyView Then
        If Not zlStr.IsHavePrivs(mstrPrivs, "����") Then vsChagre.Editable = flexEDNone
    End If
    LoadChargeAndBillTotalData = ReadChargeBillData
End Function
Private Function ReadChargeBillData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�տƱ�ݻ�������
    '����:���ݻ�ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-04 10:29:42
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mlngChargeRollingID = 0 And (mbytType = EM_�շ�Ա���� Or mbytType = EM_�����տ�_���շ�Ա) Then
         ReadChargeBillData = LoadPersonChargeAndBill         '�����շ�Ա���ʼ�¼�������շ�Ա
    Else
         ReadChargeBillData = LoadChargeAndBillAndTotal          '������ص��տ�ʹ�ü�Ʊ�ݻ���
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Public Function GetHandInFee() As Double
    Dim dblTotal As Double
    If (mlngCashRow = 0 And mlngErrorRow = 0) Or (mlngCashRow > vsChagre.Rows And mlngErrorRow > vsChagre.Rows) Then Exit Function
    With vsChagre
        dblTotal = Val(.TextMatrix(mlngErrorRow, .ColIndex("���"))) + Val(.TextMatrix(mlngCashRow, .ColIndex("���")))
    End With
    GetHandInFee = dblTotal
End Function
 
Private Function LoadPersonChargeAndBill() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�Ա�������ݻ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-04 11:28:56
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bytType As Byte, rsBalanceMode As ADODB.Recordset
    Dim strWithTable As String, lngRow As Long, lngNo As Long
    Dim str�ֽ� As String, str���㷽ʽ As String, strTemp As String, intƱ�� As Integer
    Dim dblToTotal As Double, i As Integer, dblInsure As Double, intInsureRow As Integer
    Dim blnTempDelete As Boolean, strRollingType As String, strWhere As String
    Dim strCardUseRecord As String
    On Error GoTo errHandle
    
    blnTempDelete = False
    If mstrPersonName = "" Or mstrPersonName = "-" Then
        'ֻ�������
        Call ClearData: LoadPersonChargeAndBill = True
        Exit Function
    End If
    
    bytType = 1: dblToTotal = 0
    Set rsBalanceMode = Get���㷽ʽ
    str�ֽ� = "�ֽ�"
    rsBalanceMode.Filter = "����=1"
    If Not rsBalanceMode.EOF Then
        str�ֽ� = rsBalanceMode!����
    End If
    
    Set mrsBalance = GetBalance(str�ֽ�)
    
    '1.�����տ����ݻ���
    'Ԥ������NULL,2-����,3-�շ�,4-�Һ�,5-���￨,6-����ҽ������
    'mstrRollingType:�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
    strRollingType = "": strWhere = ""
    strSQL = ""
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",1,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",3,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",4,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",5,") > 0 Or _
        InStr("," & mstrRollingType & ",", ",6,") > 0 Then
        If mstrRollingType <> "" Then
            strWhere = " And Instr([6],','|| Nvl(A.��������,0)||',')> 0"
        End If
        If InStr("," & mstrRollingType & ",", ",0,") > 0 Then
            strRollingType = ",2,3,4,5,6,"
        Else
            If InStr("," & mstrRollingType & ",", ",1,") > 0 Then
                strRollingType = "3,6,"
            End If
            If InStr("," & mstrRollingType & ",", ",3,") > 0 Then
                strRollingType = strRollingType & "2,"
            End If
            If InStr("," & mstrRollingType & ",", ",4,") > 0 Then
                strRollingType = strRollingType & "4,"
            End If
            If InStr("," & mstrRollingType & ",", ",5,") > 0 Then
                strRollingType = strRollingType & "5,"
            End If
            If InStr("," & mstrRollingType & ",", ",6,") > 0 Then
                strRollingType = strRollingType & "-,"
            End If
            If strRollingType <> "" Then strRollingType = "," & strRollingType
        End If

        strSQL = "" & _
        "   Select Decode(Nvl(a.��������, 0), 2, 2, 6, 9, 1) As ����, a.����id, Decode(Mod(a.��¼����, 10), 1, '[��Ԥ����]', a.���㷽ʽ) As ���㷽ʽ," & vbNewLine & _
        "              Sum(a.��Ԥ��) As ���, Sum(Decode(Mod(a.��¼����, 10), 1, 1, 0) * a.��Ԥ��) As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
        "       From ����Ԥ����¼ A" & vbNewLine & _
        "       Where a.����Ա���� || '' = [3] and a.��¼����<>1 " & strWhere & vbNewLine & _
        "       And a.�տ�ʱ�� Between [4] And [5] " & vbNewLine & _
        "       And Not Exists(Select 1 From ������ü�¼ B Where a.����id = b.����id And Nvl(b.����״̬, 0) = 1) " & vbNewLine & _
        "       And Not Exists(Select 1 From ���˽��ʼ�¼ B Where a.����id = b.Id And b.����״̬ Is Not Null)" & vbNewLine & _
        "       And Not Exists(Select 1 From ���ò����¼ B Where a.����id = b.����id And Nvl(b.����״̬, 0) >= 1)" & vbNewLine & _
        "       Group By Decode(Nvl(a.��������, 0), 2, 2, 6, 9, 1), a.����id, Decode(Mod(a.��¼����, 10), 1, '[��Ԥ����]', a.���㷽ʽ)" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",2,") > 0 Then 'Ԥ����
        strSQL = strSQL & _
        IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
        "       Select 3 As ����, ID As ����id, a.���㷽ʽ, a.���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
        "       From ����Ԥ����¼ A" & vbNewLine & _
        "       Where ��¼���� = 1 And ����Ա���� || '' = [3] And �տ�ʱ�� Between [4] And [5] And Nvl(��������,0) <> 12" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",21,") > 0 Then '����Ԥ����
        strSQL = strSQL & _
        IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
        "       Select 3 As ����, ID As ����id, a.���㷽ʽ, a.���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
        "       From ����Ԥ����¼ A" & vbNewLine & _
        "       Where ��¼���� = 1 And Nvl(Ԥ�����,0) = 1 And ����Ա���� || '' = [3] And �տ�ʱ�� Between [4] And [5] And Nvl(��������,0) <> 12" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",22,") > 0 Then 'סԺԤ����
        strSQL = strSQL & _
        IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
        "       Select 3 As ����, ID As ����id, a.���㷽ʽ, a.���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
        "       From ����Ԥ����¼ A" & vbNewLine & _
        "       Where ��¼���� = 1 And Nvl(Ԥ�����,0) = 2  And ����Ա���� || '' = [3] And �տ�ʱ�� Between [4] And [5] And Nvl(��������,0) <> 12" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",3,") > 0 Then '���ʲ���Ԥ��
        strSQL = strSQL & _
        IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
        "       Select 3 As ����, ID As ����id, a.���㷽ʽ, a.���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
        "       From ����Ԥ����¼ A" & vbNewLine & _
        "       Where ��¼���� = 1 And ����Ա���� || '' = [3] And �տ�ʱ�� Between [4] And [5] And Nvl(��������,0) = 12" & vbNewLine
    End If
    
    If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",6,") > 0 Then '���ѿ���ֵ
        strSQL = strSQL & _
            IIf(strSQL <> "", "Union All", "") & vbNewLine & _
            "Select 5 As ����, a.����Id As ����id, a.���㷽ʽ, a.ʵ�ս�� As ���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
            "From ���˿������¼ A, ���˿������¼ B" & vbNewLine & _
            "Where a.������� = b.�������(+) And a.���ѿ�id = b.���ѿ�id(+)  " & vbNewLine & _
            "      And (a.��¼���� = 2 Or a.��¼���� = 3 And b.��¼���� = 2) And b.��¼����(+) = 2 " & vbNewLine & _
            "      And a.Id <> b.Id(+) And a.����Ա���� || '' = [3] And a.�Ǽ�ʱ�� Between [4] And [5]" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select 6 As ����, a.����Id, a.���㷽ʽ, a.ʵ�ս�� As ���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
            "From ���˿������¼ A, ���˿������¼ B" & vbNewLine & _
            "Where a.������� = b.�������(+) And a.���ѿ�id = b.���ѿ�id(+)  " & vbNewLine & _
            "      And a.Id <> b.Id(+) And (a.��¼���� = 1 Or a.��¼���� = 3 And b.��¼���� = 1) And b.��¼����(+) = 1 " & vbNewLine & _
            "      And a.����Ա���� || '' = [3] And a.�Ǽ�ʱ�� Between [4] And [5]" & vbNewLine
    End If
    
    '���ݴ��
    strSQL = strSQL & _
    IIf(strSQL <> "", "UNION ALL ", "") & vbNewLine & _
    "       Select 4 As ����, a.Id As ����id, a.���㷽ʽ, Nvl(a.�����, 0) As ���, 0 As ��Ԥ��, Nvl(�����, 0) As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
    "       From ��Ա����¼ A" & vbNewLine & _
    "       Where a.����� || '' = [3] And a.ȡ��ʱ�� Is Null " & vbNewLine & _
    "             And a.���ʱ�� Between [4] And [5]" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select 4 As, a.Id As ����id, a.���㷽ʽ, -1 * Nvl(a.�����, 0) As ���, 0 As ��Ԥ��, 0 As ���ϼ�, Nvl(�����, 0) As ����ϼ�" & vbNewLine & _
    "       From ��Ա����¼ A" & vbNewLine & _
    "       Where a.����� || '' = [3] And a.���ʱ�� Between [4] And [5] And a.ȡ��ʱ�� Is Null" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select 7 As ����, a.Id As ����id, '�ֽ�' As ���㷽ʽ, a.���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�" & vbNewLine & _
    "       From ��Ա�ݴ��¼ A" & vbNewLine & _
    "       Where a.��¼���� = 2 And a.�ջ�ʱ�� Is Null And �տ�Ա || '' = [3] " & vbNewLine & _
    "             And a.�Ǽ�ʱ�� Between [4] And [5] " & vbNewLine


        
    strWithTable = "" & _
    " With �������� as (" & vbNewLine & _
    "       Select ����, ����id, ���㷽ʽ, Nvl(���, 0) As ���, Nvl(��Ԥ��, 0) As ��Ԥ��, Nvl(���ϼ�, 0) As ���ϼ�, Nvl(����ϼ�, 0) As ����ϼ�" & vbNewLine & _
    "       From ( " & strSQL & ") A" & vbNewLine & _
    "       Where Not Exists (Select y.��¼id From ��Ա�սɼ�¼ X, ��Ա�սɶ��� Y" & vbNewLine & _
    "                   Where y.��¼id = a.����id And x.����ʱ�� Is Null And x.�տ�Ա = [3] And x.Id = y.�ս�id And y.���� = a.����) " & vbNewLine & _
    "                  )" & vbNewLine
 
    
    strSQL = strWithTable & vbNewLine & _
        "   Select -1 as ����,0 as ����ID,���㷽ʽ,sum(nvl(���,0)) as ���, " & vbNewLine & _
        "           sum(nvl(��Ԥ��,0)) as ��Ԥ��,sum(nvl(���ϼ�,0)) as ���ϼ�,sum(nvl(����ϼ�,0)) as ����ϼ�   " & vbNewLine & _
        "   From �������� " & _
        "   Group by ���㷽ʽ " & _
        "   Union ALL" & _
        "   Select ����,����ID,'-' as ���㷽ʽ,0 as ���, " & vbNewLine & _
        "           0 as ��Ԥ��,0 as ���ϼ�,0 as ����ϼ�   " & vbNewLine & _
        "   From �������� " & vbNewLine & _
        "   Group by ����,����ID " & vbNewLine & _
        "   "
    strSQL = "" & _
        "   Select  ����,  nvl(����ID,0) as ����ID,���㷽ʽ,���,��Ԥ��, ���ϼ�, ����ϼ�   " & vbNewLine & _
        "   From (" & strSQL & ")" & vbNewLine & _
        "   Order by ����"
        
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID, bytType, mstrPersonName, mdtStartDate, mdtendDate, strRollingType)
    With vsChagre
        dblInsure = 0
        .Clear 1
        .Rows = 2: lngRow = 1
        mrsList.Filter = "����=-1 And ��� <> 0"
        mlngCashRow = 0
        
        Do While Not mrsList.EOF
            str���㷽ʽ = Nvl(mrsList!���㷽ʽ)
            If str���㷽ʽ = "[��Ԥ����]" Then
                lngNo = 1
                dblToTotal = dblToTotal + Val(Nvl(mrsList!���))
            Else
                rsBalanceMode.Filter = "����='" & str���㷽ʽ & "'"
                If Not rsBalanceMode.EOF Then
                    '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
                    Select Case Val(Nvl(rsBalanceMode!����))
                    Case 1 '�ֽ���㷽ʽ
                        lngNo = 10
                        mlngCashRow = lngRow
                    Case 2  '������ҽ������
                        lngNo = 11
                    Case 7 'һ��ͨ����
                        lngNo = 12
                    Case 8  '���㿨����
                        lngNo = 14
                    Case 3 '�����˻�
                        lngNo = 15
                    Case 4   '
                        lngNo = 16
                    Case 9
                        lngNo = 17
                        mlngErrorRow = lngRow
                    Case Else
                        lngNo = 18
                    End Select
                Else
                    lngNo = 13
                End If
                If lngNo = 15 Or lngNo = 16 Then dblInsure = dblInsure + Val(Nvl(mrsList!���))
                dblToTotal = dblToTotal + Val(Nvl(mrsList!���))
            End If
            
            If Not (lngNo = 15 Or lngNo = 16) Then
                .TextMatrix(lngRow, .ColIndex("���")) = lngNo
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = str���㷽ʽ
                If lngNo = 17 Then
                    .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsList!���)), "0.###########")
                Else
                    .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsList!���)), "#,###0.00;-#,###0.00;0.00;-0.00")
                End If
                .RowData(lngRow) = Val(Nvl(mrsList!���))
                If mlngChargeRollingID <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(mrsList!�����)
                End If
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            mrsList.MoveNext
        Loop
        mrsList.Filter = "����=-1"
        If mrsList.RecordCount <> 0 Then
            mrsList.MoveFirst
            Do While Not mrsList.EOF
                If Val(Nvl(mrsList!���ϼ�)) <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("���")) = 2
                    .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = "[����ϼ�]"
                    .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsList!���ϼ�)), "#,###0.00;-#,###0.00;0.00;-0.00")
                    .RowData(lngRow) = Val(Nvl(mrsList!���ϼ�))
                    .Rows = .Rows + 1
                    lngRow = lngRow + 1
                End If
                If Val(Nvl(mrsList!����ϼ�)) <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("���")) = 3
                    .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = "[����ϼ�]"
                    .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsList!����ϼ�)), "#,###0.00;-#,###0.00;0.00;-0.00")
                    .RowData(lngRow) = Val(Nvl(mrsList!����ϼ�))
                    .Rows = .Rows + 1
                    lngRow = lngRow + 1
                End If
                mrsList.MoveNext
            Loop
        End If
        
        If .Rows > 2 Then .Rows = .Rows - 1: blnTempDelete = True
        .Cell(flexcpSort, 1, .ColIndex("���"), .Rows - 1, .ColIndex("���")) = flexSortNumericAscending
        If blnTempDelete = True Then .Rows = .Rows + 1
        intInsureRow = lngRow
        .TextMatrix(intInsureRow, .ColIndex("���㷽ʽ")) = "ҽ�����"
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H80000016
        .IsSubtotal(intInsureRow) = True
        lngRow = lngRow + 1
        .Rows = .Rows + 1
        
        mrsList.Filter = ""
        mrsList.Filter = "����=-1 And ��� <> 0"
        Do While Not mrsList.EOF
            str���㷽ʽ = Nvl(mrsList!���㷽ʽ)
            If str���㷽ʽ = "[��Ԥ����]" Then
                lngNo = 1
            Else
                rsBalanceMode.Filter = "����='" & str���㷽ʽ & "'"
                If Not rsBalanceMode.EOF Then
                    '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
                    Select Case Val(Nvl(rsBalanceMode!����))
                    Case 1 '�ֽ���㷽ʽ
                        lngNo = 10
                        mlngCashRow = lngRow
                     Case 2  '������ҽ������
                        lngNo = 11
                    Case 7 'һ��ͨ����
                        lngNo = 12
                    Case 8  '���㿨����
                        lngNo = 14
                    Case 3 '�����˻�
                        lngNo = 15
                    Case 4 'ҽ��
                        lngNo = 16
                    Case 9
                        lngNo = 17
                        mlngErrorRow = lngRow
                    Case Else
                        lngNo = 18
                    End Select
                Else
                    lngNo = 13
                End If
            End If

            If lngNo = 15 Or lngNo = 16 Then
                .TextMatrix(lngRow, .ColIndex("���")) = lngNo
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = str���㷽ʽ
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsList!���)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(mrsList!���))
                .RowOutlineLevel(lngRow) = 1
                If mlngChargeRollingID <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(mrsList!�����)
                End If
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            mrsList.MoveNext
        Loop
        
        .TextMatrix(intInsureRow, .ColIndex("���")) = Format(dblInsure, "#,###0.00;-#,###0.00;0.00;-0.00")
        
        If .Rows > 2 Then .Rows = .Rows - 1
        
        If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = "ҽ�����" Then
            .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = ""
            .TextMatrix(.Rows - 1, .ColIndex("���")) = ""
            .IsSubtotal(.Rows - 1) = False
        End If
        .Outline (0)
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
'        dblToTotal = Round(dblToTotal, 2)
        If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = "" And .Rows > 2 Then .Rows = .Rows - 1
        txtTotal.Text = Format(dblToTotal, "0.#########") & "Ԫ"
        For i = 0 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���㷽ʽ")) = str�ֽ� Then mlngCashRow = i
        Next i
    End With
    
    '�ָ�������
    zl_vsGrid_Para_Restore mlngMode, vsChagre, Me.Name, "���㷽ʽ�б�", False
    
    '2.����Ʊ��ʹ�������Ϣ
    'Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
    'mstrRollingType:�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
    
    strRollingType = "": strWhere = ""
    If mstrRollingType <> "" Then
        If InStr("," & mstrRollingType & ",", ",0,") > 0 Then
            strWhere = ""
        Else
            '110414:���ϴ���2017/6/20��ҽ�ƿ�ʹ�����﷢Ʊ
            strWhere = ""
            If InStr("," & mstrRollingType & ",", ",1,") > 0 Then
                strWhere = " (Instr(',1,' , ','|| A.Ʊ�� || ',') > 0 And Not Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E Where D.ID=A.ID And D.��ӡID=E.ID And E.�������� In (3,4,5))) "
            End If
            If InStr("," & mstrRollingType & ",", ",2,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',2,' , ','|| A.Ʊ�� || ',') > 0 And Not Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E,����Ԥ����¼ F Where D.ID=A.ID And D.��ӡID=E.ID And E.��������=2 And E.NO=F.NO And F.��¼����=1 And Nvl(F.��������,0) = 12))"
            End If
            '����ţ�118482,����,2017/12/19/�������ΪԤ��ʱ����Ӧ���ж� ��������.������תסԺԤ�� ��סԺԤ��ת����֮��, Ʊ��Ӧ�����벻���
            If InStr("," & mstrRollingType & ",", ",21,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',2,' , ','|| A.Ʊ�� || ',') > 0 And Not Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E,����Ԥ����¼ F Where D.ID=A.ID And D.��ӡID=E.ID And E.��������=2 And E.NO=F.NO And F.��¼����=1 And Nvl(F.��������,0) = 12) And Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E,����Ԥ����¼ F Where D.ID=A.ID And D.��ӡID=E.ID And E.��������=2 And E.NO=F.NO And F.��¼����=1 And F.Ԥ�����=1 And F.���>0 ))"
            End If
            If InStr("," & mstrRollingType & ",", ",22,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',2,' , ','|| A.Ʊ�� || ',') > 0 And Not Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E,����Ԥ����¼ F Where D.ID=A.ID And D.��ӡID=E.ID And E.��������=2 And E.NO=F.NO And F.��¼����=1 And Nvl(F.��������,0) = 12) And Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E,����Ԥ����¼ F Where D.ID=A.ID And D.��ӡID=E.ID And E.��������=2 And E.NO=F.NO And F.��¼����=1 And F.Ԥ�����=2 And F.���>0 ))"
            End If
            If InStr("," & mstrRollingType & ",", ",3,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',1,2,3,' , ','|| A.Ʊ�� || ',') > 0 And Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E Where D.ID=A.ID And D.��ӡID=E.ID And E.�������� = 3) And Not Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E,����Ԥ����¼ F Where D.ID=A.ID And D.��ӡID=E.ID And E.��������=2 And E.NO=F.NO And F.��¼����=1 And Nvl(F.��������,0) <> 12))"
            End If
            If InStr("," & mstrRollingType & ",", ",4,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',1,4,' , ','|| A.Ʊ�� || ',') > 0 And Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E Where D.ID=A.ID And D.��ӡID=E.ID And E.��������=4))"
            End If
            If InStr("," & mstrRollingType & ",", ",5,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',1,5,' , ','|| A.Ʊ�� || ',') > 0 And Exists (Select 1 From Ʊ��ʹ����ϸ D,Ʊ�ݴ�ӡ���� E Where D.ID=A.ID And D.��ӡID=E.ID And E.��������=5))"
            End If
            If InStr("," & mstrRollingType & ",", ",6,") > 0 Then
                strWhere = strWhere & IIf(strWhere = "", "", " Or ") & " (Instr(',-,' , ','|| A.Ʊ�� || ',') > 0)"
            End If
        End If
        '���ѿ�����
        If InStr("," & mstrRollingType & ",", ",0,") > 0 Or InStr("," & mstrRollingType & ",", ",6,") > 0 Then
            strCardUseRecord = _
            "Union All" & vbNewLine & _
            "Select 6 As Ʊ��, a.ԭ��, a.����, a.���� As ����, Zl_Incstr(a.����) As ��һ����, a.ʹ��ʱ��, Null As ��ӡid, b.����" & vbNewLine & _
            "From ���ѿ�ʹ�ü�¼ A, ���ѿ����ü�¼ B" & vbNewLine & _
            "Where a.����id = b.Id(+) And a.ʹ���� || '' = [3] And a.ʹ��ʱ�� Between [4] And [5]" & vbNewLine & _
            "      And ((a.���� = 2" & vbNewLine & _
            "               And Not Exists (Select 1" & vbNewLine & _
            "                   From ��Ա�ս�Ʊ�� C, ��Ա�սɼ�¼ D" & vbNewLine & _
            "                   Where c.�ս�id = d.Id And Nvl(c.����, '-') = Nvl(b.����, '-') And c.���� In (2, 3)" & vbNewLine & _
            "                       And Length(c.��ʼƱ��) = Length(a.����) And a.���� Between c.��ʼƱ�� And c.��ֹƱ��" & vbNewLine & _
            "                       And c.Ʊ�� = 6 And d.����ʱ�� Is Null))" & vbNewLine & _
            "           Or (a.���� = 1" & vbNewLine & _
            "               And Not Exists (Select 1" & vbNewLine & _
            "                   From ��Ա�ս�Ʊ�� C, ��Ա�սɼ�¼ D" & vbNewLine & _
            "                   Where c.�ս�id = d.Id And Nvl(c.����, '-') = Nvl(b.����, '-') And c.���� = 1" & vbNewLine & _
            "                       And Length(c.��ʼƱ��) = Length(a.����) And a.���� Between c.��ʼƱ�� And c.��ֹƱ��" & vbNewLine & _
            "                       And c.Ʊ�� = 6 And d.����ʱ�� Is Null)))"
        End If
        
        If strWhere <> "" Then strWhere = "And (" & strWhere & ")"
        
        strWithTable = "" & _
        "    With Ʊ��ʹ�� As  ( " & _
        "           Select A.Ʊ��, A.ԭ��,A.����, A.����, Zl_Incstr(A.����) As ��һ����,A.ʹ��ʱ��,A.��ӡID,B.���� " & _
        "           From Ʊ��ʹ����ϸ A,Ʊ�����ü�¼ B " & _
        "           Where A.ʹ����|| '' = [3] And A.����id = B.id And A.ʹ��ʱ�� Between [4] and [5] " & _
        "                   And ((A.���� = 2 And Not Exists(Select 1 " & _
        "                           From ��Ա�ս�Ʊ�� C,��Ա�սɼ�¼ D,Ʊ��ʹ����ϸ E,Ʊ�����ü�¼ F " & _
        "                           Where c.�ս�id = d.Id And d.�տ�Ա = [3] And e.��ӡid = a.��ӡid And e.����id = f.Id " & _
        "                               And Nvl(f.����,'-') = Nvl(c.����,'-') And c.���� In (2, 3) " & _
        "                               And d.����ʱ�� Is Null And a.Ʊ�� = c.Ʊ�� And Length(c.��ʼƱ��) = Length(a.����) " & _
        "                               And a.���� Between c.��ʼƱ�� And c.��ֹƱ��)) " & _
        "                       Or (a.���� = 1 And Not Exists(Select 1 " & _
        "                           From ��Ա�ս�Ʊ�� E,��Ա�սɼ�¼ F,Ʊ��ʹ����ϸ G,Ʊ�����ü�¼ H " & _
        "                           Where e.�ս�ID=f.ID And f.�տ�Ա = [3] And g.��ӡid = a.��ӡid And g.����id = h.Id " & _
        "                               And Nvl(h.����,'-') = Nvl(e.����,'-') And e.���� = 1 And f.����ʱ�� Is Null And a.Ʊ��=e.Ʊ�� " & _
        "                               And Length(e.��ʼƱ��) = Length(a.����) And a.���� between e.��ʼƱ�� and e.��ֹƱ��)))" & _
                    strWhere & vbNewLine & _
                    strCardUseRecord & "), "
        
        strWithTable = strWithTable & _
        "           �ջ�Ʊ�� as (  " & _
        "               Select Distinct 1 as ����,y.Ʊ��,x.No, y.���� " & _
        "               From Ʊ�ݴ�ӡ���� X, Ʊ��ʹ�� Y " & _
        "               Where y.ԭ�� = 2 AND y.����=2 And x.Id = y.��ӡid AND Y.Ʊ��<>1 " & _
        "               Union all " & _
        "               Select Distinct 2 as ����,y.Ʊ��,x.No, y.���� " & _
        "               From Ʊ�ݴ�ӡ���� X, Ʊ��ʹ�� Y " & _
        "               Where y.ԭ�� = 4 AND y.����=2 And x.Id = y.��ӡid AND Y.Ʊ��<>1 ), " & _
        "           �ջ�Ʊ��_�շ� as ( " & _
        "               Select 1 as ����,Ʊ��,���� From Ʊ��ʹ��  where ԭ��=2 And ����=2 AND Ʊ��=1 " & _
        "               Union all " & _
        "               Select 2 as ����,Ʊ��,���� From Ʊ��ʹ��  where ԭ��=4 And ����=2 AND Ʊ��=1 ), " & _
        "           �ջ�Ʊ�ݽ�� as ( " & _
        "               Select  a.����, a.Ʊ��,A.����,Sum(C.���ʽ��) As ���ݽ��  " & _
        "               From �ջ�Ʊ��_�շ� A,Ʊ�ݴ�ӡ��ϸ B, ������ü�¼ C " & _
        "               Where a.Ʊ��=b.Ʊ�� And a.����=b.Ʊ�� and a.Ʊ��=1     " & _
        "                           And b.No =c.No And C.��¼���� = 1 And C.��¼״̬ In (3, 1) " & _
        "                           And Instr(',' || b.��� || ',', ',' || Nvl(c.�۸񸸺�, c.���) || ',') > 0 " & _
        "                Group by a.����,a.Ʊ��,A.���� "
        strWithTable = strWithTable & _
        "               Union all " & _
        "               Select a.����,A.Ʊ��,A.����,sum(C.��Ԥ��) as ���ݽ��  " & _
        "               From �ջ�Ʊ�� A,���˽��ʼ�¼ B,����Ԥ����¼ C " & _
        "               Where a.Ʊ��=3 And A.NO=B.NO And B.����״̬ Is Null and b.��¼״̬ in (1,3) and b.ID=C.����ID  " & _
        "               Group by A.����,a.Ʊ��,a.���� " & _
        "               Union all " & _
        "               Select a.����,A.Ʊ��,A.����,sum(B.���) " & _
        "               From �ջ�Ʊ�� A,����Ԥ����¼  B " & _
        "               Where a.Ʊ��=2 And a.no=b.No  and b.��¼����=1 and B.��¼״̬ in (1,3) " & _
        "               Group by A.����,a.Ʊ��,a.���� " & _
        "               Union all " & _
        "               Select a.����,A.Ʊ��,a.����,sum(b.���ʽ��) as ���ݽ��  " & _
        "               From �ջ�Ʊ�� A,������ü�¼ B " & _
        "               Where a.Ʊ��=4 And A.NO=B.NO and b.��¼����=4 and B.��¼״̬ in (1,3)  " & _
        "               Group by A.����,a.Ʊ��,a.���� " & _
        "               Union all " & _
        "               Select a.����,A.Ʊ��,a.����,sum(b.���ʽ��) as ���ݽ��  " & _
        "               From �ջ�Ʊ�� A,סԺ���ü�¼ B " & _
        "               Where a.Ʊ��=5 And A.NO=B.NO and b.��¼����=5 and B.��¼״̬ in (1,3) and Nvl(B.���ʷ���, 0) = 0 " & _
        "               Group by A.����,a.Ʊ��,a.���� "
        strWithTable = strWithTable & _
        "               Union all " & _
        "               Select a.����, a.Ʊ��, a.����, Sum(b.���ʽ��) As ���ݽ��" & vbNewLine & _
        "               From �ջ�Ʊ��_�շ� A, ������ü�¼ B, Ʊ��ʹ����ϸ C, Ʊ�ݴ�ӡ���� D" & vbNewLine & _
        "               Where a.Ʊ�� = 1 And a.���� = c.���� And c.���� = 2 And c.��ӡid = d.Id And d.�������� = 4 And d.No = b.No And b.��¼���� = 4 And b.��¼״̬ In (1, 3)" & vbNewLine & _
        "               Group By a.����, a.Ʊ��, a.���� " & _
        "               Union all " & _
        "               Select a.����, a.Ʊ��, a.����, Sum(b.���ʽ��) As ���ݽ��" & vbNewLine & _
        "               From �ջ�Ʊ��_�շ� A, סԺ���ü�¼ B, Ʊ��ʹ����ϸ C, Ʊ�ݴ�ӡ���� D" & vbNewLine & _
        "               Where a.Ʊ�� = 1 And a.���� = c.���� And c.���� = 2 And c.��ӡid = d.Id And d.�������� = 5 And d.No = b.No And b.��¼���� = 5 And b.��¼״̬ In (1, 3)" & vbNewLine & _
        "               Group By a.����, a.Ʊ��, a.���� ) "
    
        strSQL = "" & _
        "   Select /*+ Rule*/   Ʊ��,����,����,��ʼ���� as ��ʼƱ��,��ֹ���� as ��ֹƱ��,���,ʹ��ʱ��  as ����ʱ��,���� " & _
        "   FROM (   " & strWithTable & _
        "               Select 1 As ����, a.Ʊ��,a.����, a.���� As ��ʼ����, b.���� As ��ֹ����,count(*) as ����,null as ʹ��ʱ��,0 as ���, a.���� " & _
        "               From (Select Rownum As �к�, Ʊ��, ����, ���� " & _
        "                           From (Select Ʊ��, ����, ���� From Ʊ��ʹ�� where ԭ�� In (1,3,6)  Minus Select Ʊ��, ��һ����, ���� From Ʊ��ʹ�� where ԭ�� In (1,3,6))) A, " & _
        "                         (Select Rownum As �Ϻ�, Ʊ��, Zl_Incstr_Pre(����) As ����,���� " & _
        "                           From (Select Ʊ��, ��һ���� As ����,���� From Ʊ��ʹ��  Ʊ��ʹ�� where ԭ�� In (1,3,6)  Minus Select Ʊ��, ����,���� From Ʊ��ʹ��  Ʊ��ʹ�� where ԭ�� In (1,3,6))) B,"
        strSQL = strSQL & "" & _
        "                          ( Select distinct Ʊ��,���� from Ʊ��ʹ��) M " & _
        "               Where a.�к� = b.�Ϻ� And a.Ʊ�� = b.Ʊ�� and a.Ʊ��=M.Ʊ�� And M.���� between a.���� and b.���� And Nvl(a.����,0) = Nvl(b.����,0)   " & _
        "               Group by a.Ʊ��,a.����,b.����,a.���� " & _
        "               Union all " & _
        "               Select 2 As ����, a.Ʊ��,a.����, a.���� As ��ʼ����, b.���� As ��ֹ����,count(*) as ����,m.ʹ��ʱ�� as ʹ��ʱ��,sum(q.���ݽ��) as ���ݽ��, a.���� " & _
        "               From (Select Rownum As �к�, Ʊ��, ����, ���� " & _
        "               From (Select Ʊ��,ʹ��ʱ��, ����,���� From Ʊ��ʹ�� where ԭ��=2 And ����=2 Minus Select Ʊ��,ʹ��ʱ��, ��һ����,���� From Ʊ��ʹ�� where ԭ��=2 and ����=2)) A, " & _
        "                          (Select Rownum As �Ϻ�, Ʊ��, Zl_Incstr_Pre(����) As ����, ���� " & _
        "                           From (Select Ʊ��,ʹ��ʱ��, ��һ���� As ����,���� From Ʊ��ʹ��    where ԭ��=2 And ����=2 Minus Select Ʊ��,ʹ��ʱ��, ����,���� From Ʊ��ʹ��  where ԭ��=2 And ����=2)) B, " & _
        "                           (select  Ʊ��,����,Max(ʹ��ʱ��) as ʹ��ʱ�� From Ʊ��ʹ�� Where ԭ��=2 and ����=2 Group by Ʊ��,����)  M,�ջ�Ʊ�ݽ�� Q " & _
        "               Where a.�к� = b.�Ϻ� And a.Ʊ�� = b.Ʊ�� and a.Ʊ��=M.Ʊ�� And M.���� between a.���� and b.���� And Nvl(a.����,0) = Nvl(b.����,0)  " & _
        "                           and m.Ʊ��=Q.Ʊ��(+) and m.����=Q.����(+) AND q.����(+)=1  " & _
        "               group by a.Ʊ��,a.����,b.����,m.ʹ��ʱ��,a.���� " & _
        "               union all  " & _
        "               Select 3 As ����, a.Ʊ��,a.����, a.���� As ��ʼ����, b.���� As ��ֹ����,count(*) as ����,m.ʹ��ʱ�� as ʹ��ʱ��,sum(q.���ݽ��) as ���ݽ��, a.���� " & _
        "               From (  Select Rownum As �к�, Ʊ��, ����,���� " & _
        "                           From (Select Ʊ��,ʹ��ʱ��, ����,���� From Ʊ��ʹ�� where ԭ��=4 And ����=2 Minus Select Ʊ��,ʹ��ʱ��, ��һ����,���� From Ʊ��ʹ�� where ԭ��=4 and ����=2)) A, " & _
        "                           (  Select Rownum As �Ϻ�, Ʊ��, Zl_Incstr_Pre(����) As ����, ���� " & _
        "                               From (Select Ʊ��,ʹ��ʱ��, ��һ���� As ����,���� From Ʊ��ʹ��    where ԭ��=4 And ����=2 Minus Select Ʊ��,ʹ��ʱ��, ����,���� From Ʊ��ʹ��    where ԭ��=4 And ����=2)) B, " & _
        "                           (select  Ʊ��,����,Max(ʹ��ʱ��) as ʹ��ʱ�� From Ʊ��ʹ�� Where ԭ��=4 And ����=2 Group by Ʊ��,����)  M,�ջ�Ʊ�ݽ�� Q " & _
        "               Where a.�к� = b.�Ϻ� And a.Ʊ�� = b.Ʊ�� and a.Ʊ��=M.Ʊ�� And M.���� between a.���� and b.���� And Nvl(a.����,0) = Nvl(b.����,0)  " & _
        "                       and m.Ʊ��=Q.Ʊ��(+) and m.����=Q.����(+) AND q.����(+)=2 " & _
        "               group by a.Ʊ��,a.����,b.����,m.ʹ��ʱ��,a.���� ) " & _
        " ORDER BY Ʊ��,����,ʹ��ʱ�� ,��ʼ����"
    Else
        strSQL = "" & _
        "   Select 1 as Ʊ��, 1 as ����,0 as ����,''  as ��ʼƱ��, '' as ��ֹƱ��,0 as ���, sysdate as ����ʱ��, Null as ����" & _
        "   From dual " & _
        "   Where 1=2"
    End If
    
    Set mrsListBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID, bytType, mstrPersonName, mdtStartDate, mdtendDate)
    
    With vsReturnBill
        .Clear 1
        .Rows = 2: lngRow = 1
        mrsListBill.Filter = "����=2 Or ����=3"
        Do While Not mrsListBill.EOF
            If .TextMatrix(lngRow - 1, .ColIndex("���")) <> GetBillTypeName(mrsListBill!Ʊ��) Then
                .TextMatrix(lngRow, .ColIndex("���")) = GetBillTypeName(mrsListBill!Ʊ��)
                .IsSubtotal(lngRow) = True
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H80000016
                lngRow = lngRow + 1
                .Rows = .Rows + 1
            End If
            .RowOutlineLevel(lngRow) = 1
            .TextMatrix(lngRow, .ColIndex("���")) = GetBillTypeName(mrsListBill!Ʊ��)
            .TextMatrix(lngRow, .ColIndex("����")) = Decode(mrsListBill!����, _
                                                                2, Decode(mrsListBill!Ʊ��, 6, "�ջ�", "�˷�"), _
                                                                3, Decode(mrsListBill!Ʊ��, 6, "����", "�ش�"), "����")
            .TextMatrix(lngRow, .ColIndex("�ջ�ʱ��")) = Format(mrsListBill!����ʱ��, "yyyy-mm-dd HH:MM")
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsListBill!���)), "#,###0.00;-#,###0.00;0.00;-0.00")
            If InStr(";�շ�;���ѿ�;", ";" & .TextMatrix(lngRow, .ColIndex("���")) & ";") > 0 And Val(.TextMatrix(lngRow, .ColIndex("���"))) = 0 Then
                .TextMatrix(lngRow, .ColIndex("���")) = "-"
            End If
            If Nvl(mrsListBill!��ʼƱ��) = Nvl(mrsListBill!��ֹƱ��) Then
                .TextMatrix(lngRow, .ColIndex("Ʊ�ݺ�")) = Nvl(mrsListBill!��ʼƱ��)
            Else
                .TextMatrix(lngRow, .ColIndex("Ʊ�ݺ�")) = Nvl(mrsListBill!��ʼƱ��) & "-" & Nvl(mrsListBill!��ֹƱ��)
            End If
            .Rows = .Rows + 1: lngRow = lngRow + 1
            mrsListBill.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        If .TextMatrix(.Rows - 1, .ColIndex("���")) = "" Then
            .IsSubtotal(.Rows - 1) = False
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        '�ָ�������
    zl_vsGrid_Para_Restore mlngMode, vsReturnBill, Me.Name, "�ջ�Ʊ���б�", False
    End With
  
    With vsBill
        .Clear 1
        .Rows = 1: lngRow = 1: .Cols = 3
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        mrsListBill.Filter = 0:  strTemp = ""
        Do While Not mrsListBill.EOF
            intƱ�� = Val(Nvl(mrsListBill!Ʊ��))
            If InStr(1, strTemp & ",", "," & intƱ�� & ",") = 0 Then
                strTemp = strTemp & "," & intƱ��
                .Rows = .Rows + 2
                .TextMatrix(.Rows - 3, 0) = GetBillTypeName(mrsListBill!Ʊ��)
                .TextMatrix(.Rows - 2, 0) = GetBillTypeName(mrsListBill!Ʊ��)
                .TextMatrix(.Rows - 3, 1) = "����Ʊ��"
                .TextMatrix(.Rows - 2, 1) = "Ʊ�ݷ�Χ"
                .Cell(flexcpData, .Rows - 3, 0, .Rows - 2, 0) = intƱ��
            End If
            mrsListBill.MoveNext
        Loop
        Dim lngBillTotal(0 To 2) As Long
        lngRow = 0
        For lngRow = 0 To .Rows - 1 Step 2
             intƱ�� = Val(.Cell(flexcpData, lngRow, 0))
             mrsListBill.Filter = "Ʊ��=" & intƱ��
             lngBillTotal(0) = 0: lngBillTotal(1) = 0: lngBillTotal(2) = 0
             Do While Not mrsListBill.EOF
                Select Case Val(Nvl(mrsListBill!����))
                Case 1 '����Ʊ��ͳ��
                        lngBillTotal(0) = lngBillTotal(0) + Val(Nvl(mrsListBill!����))
                        .TextMatrix(lngRow + 1, 2) = Trim(.TextMatrix(lngRow + 1, 2)) & IIf(Trim(.TextMatrix(lngRow + 1, 2)) = "", "", ";") & _
                            IIf(Trim(.TextMatrix(lngRow, 2)) = "", "", ";") & _
                            IIf(Nvl(mrsListBill!��ʼƱ��) = Nvl(mrsListBill!��ֹƱ��), _
                                Nvl(mrsListBill!��ʼƱ��), _
                                Nvl(mrsListBill!��ʼƱ��) & "-" & Nvl(mrsListBill!��ֹƱ��))
                Case 2
                        lngBillTotal(1) = lngBillTotal(1) + Val(Nvl(mrsListBill!����))
                Case 3
                        lngBillTotal(2) = lngBillTotal(2) + Val(Nvl(mrsListBill!����))
                End Select
                mrsListBill.MoveNext
             Loop
             If intƱ�� <> 0 Then
                .TextMatrix(lngRow, 2) = "ʹ��:" & lngBillTotal(0) & "��; " & _
                    Decode(intƱ��, 6, "�ջ�(�˿������պͻ���):", "�ջ�(�˷Ѻ��ش�):") & lngBillTotal(1) + lngBillTotal(2) & "��"
            End If
        Next
        .Rows = .Rows - 1
    End With
    LoadPersonChargeAndBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Private Function GetBillTypeName(ByVal bytƱ�� As Byte) As String
    '����Ʊ�ֻ�ȡ�������
    On Error GoTo errHandle
    GetBillTypeName = Decode(bytƱ��, _
        1, "�շ�", 2, "Ԥ��", 3, "����", 4, "�Һ�", _
        5, "���￨", 6, "���ѿ�", "����")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function LoadChargeAndBillAndTotal() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�տƱ�ݻ�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-04 11:28:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, bytType As Byte, rsBalanceMode As ADODB.Recordset
    Dim strWithTable As String, lngRow As Long, lngNo As Long
    Dim str���㷽ʽ As String, strTemp As String, intƱ�� As Integer
    Dim strWhere As String, dblToTotal As Double, blnCYJ As Boolean
    Dim strTable As String, i As Integer, intInsureRow As Integer
    Dim objRecord As ReportRecord, dblInsure As Double
    Dim objItem As ReportRecordItem, rsList As ADODB.Recordset
    
    On Error GoTo errHandle
    '1.�����տ����ݻ���
    strSQL = "" & _
    "   Select decode(nvl(M.����,0),1,1,2,2,3,10,4,11,9,9,4) as ���,  " & _
    "           b.���㷽ʽ,b.���,b.�����,b.���," & _
    "           a.��Ԥ���� as ��Ԥ��,A.����ϼ� as ���ϼ�,A.����ϼ� " & _
    "   From ��Ա�սɼ�¼ A, ��Ա�ս���ϸ B,���㷽ʽ M" & _
    "   Where a.Id = b.�ս�id And a.ID=[1] and B.���㷽ʽ=M.����(+) and nvl(���,0)<>0 " & _
    "   Order by ���,���㷽ʽ"
    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID, bytType, mstrPersonName, mdtStartDate, mdtStartDate)
    
    strSQL = "Select a.��Ԥ���� As ��Ԥ��, a.����ϼ� As ���ϼ�, a.����ϼ� From ��Ա�սɼ�¼ A Where a.Id = [1]"
    Set rsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID)
    With vsChagre
        .Clear 1
        .Rows = 2: lngRow = 1: blnCYJ = False
        If rsList.RecordCount <> 0 Then
            If Val(Nvl(rsList!��Ԥ��)) <> 0 Then
                .TextMatrix(lngRow, .ColIndex("���")) = 1
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = "[��Ԥ����]"
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(rsList!��Ԥ��)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(rsList!��Ԥ��))
                dblToTotal = dblToTotal + Val(Nvl(rsList!��Ԥ��))
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            If Val(Nvl(rsList!���ϼ�)) <> 0 Then
                .TextMatrix(lngRow, .ColIndex("���")) = 2
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = "[����ϼ�]"
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(rsList!���ϼ�)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(rsList!���ϼ�))
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            If Val(Nvl(rsList!����ϼ�)) <> 0 Then
                .TextMatrix(lngRow, .ColIndex("���")) = 3
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = "[����ϼ�]"
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(rsList!����ϼ�)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(rsList!����ϼ�))
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
        End If
        mlngCashRow = 0
        mrsList.Filter = "���<>10 And ���<>11"
        Do While Not mrsList.EOF
            str���㷽ʽ = Nvl(mrsList!���㷽ʽ)
            dblToTotal = dblToTotal + Val(Nvl(mrsList!���))
            .TextMatrix(lngRow, .ColIndex("���")) = Val(Nvl(mrsList!���)) + 10
            If Val(Nvl(mrsList!���)) = 1 Then mlngCashRow = lngRow
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = str���㷽ʽ
            If Val(Nvl(mrsList!���)) = 9 Then
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsList!���)), "0.#########")
            Else
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsList!���)), "#,###0.00;-#,###0.00;0.00;-0.00")
            End If
            .RowData(lngRow) = Val(Nvl(mrsList!���))
            .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(mrsList!�����)
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            mrsList.MoveNext
        Loop
        mrsList.Filter = "���=10 Or ���=11"
        If mrsList.RecordCount <> 0 Then
            dblInsure = 0
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = "ҽ�����"
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H80000016
            intInsureRow = lngRow
            .IsSubtotal(intInsureRow) = True
            .Rows = .Rows + 1
            lngRow = lngRow + 1
            Do While Not mrsList.EOF
                str���㷽ʽ = Nvl(mrsList!���㷽ʽ)
                dblInsure = dblInsure + Val(Nvl(mrsList!���))
                dblToTotal = dblToTotal + Val(Nvl(mrsList!���))
                .TextMatrix(lngRow, .ColIndex("���")) = Val(Nvl(mrsList!���)) + 10
                .RowOutlineLevel(lngRow) = 1
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = str���㷽ʽ
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsList!���)), "#,###0.00;-#,###0.00;0.00;-0.00")
                .RowData(lngRow) = Val(Nvl(mrsList!���))
                .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(mrsList!�����)
                .Rows = .Rows + 1
                lngRow = lngRow + 1
                mrsList.MoveNext
            Loop
            .TextMatrix(intInsureRow, .ColIndex("���")) = Format(dblInsure, "#,###0.00;-#,###0.00;0.00;-0.00")
            .Outline (0)
        End If
        If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = "" And .Rows > 2 Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        txtTotal.Text = Format(dblToTotal, "0.#########") & "Ԫ"
    End With
    '�ָ�������
    zl_vsGrid_Para_Restore mlngMode, vsChagre, Me.Name, "���㷽ʽ�б�", False
    
    '2.����Ʊ��ʹ�������Ϣ
        '������ʷ����
        'Ʊ��,����,����,��ʼ���� as ��ʼƱ��,��ֹ���� as ��ֹƱ��,���,ʹ��ʱ��  as ����ʱ��
    strTable = ""
    If mblnDel Or mbytType = EM_�շ�Ա���� Then
        If mbytType = EM_�շ�Ա���� Then
            strWhere = " And  A.ID =[1] And a.��¼����=1"
        Else
            strWhere = " And  A.ID=C.��¼ID And C.����=8 And C.�ս�ID=[1] And a.��¼����=1"
            strTable = ",��Ա�սɶ��� C"
        End If
    Else
        If mbytType = EM_С���տ� Then
            strWhere = " And  A.С���տ�ID =[1] And a.��¼����=1"
        ElseIf mbytType = EM_С������ Then
            strWhere = " And  A.С������ID =[1] And a.��¼����=1"
        Else
            strWhere = " And  A.�����տ�ID =[1] And a.��¼����=1"
        End If
    End If
    
    strSQL = "" & _
    "   Select b.Ʊ��,b.����,b.Ʊ������ as ����,b.��ʼƱ��,b.��ֹƱ��,b.���,b.����ʱ�� " & _
    "   From ��Ա�սɼ�¼ a,��Ա�ս�Ʊ�� b " & strTable & _
    "   Where  a.id=b.�ս�id  " & strWhere & _
    "   Order by Ʊ��,����,���"
    Set mrsListBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID, bytType, mstrPersonName, mdtStartDate, mdtStartDate)
    
    With vsReturnBill
        .Clear 1
        .Rows = 2: lngRow = 1
        mrsListBill.Filter = "����=2 Or ����=3"
        Do While Not mrsListBill.EOF
            If .TextMatrix(lngRow - 1, .ColIndex("���")) <> GetBillTypeName(mrsListBill!Ʊ��) Then
                .TextMatrix(lngRow, .ColIndex("���")) = GetBillTypeName(mrsListBill!Ʊ��)
                .IsSubtotal(lngRow) = True
                .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = &H80000016
                lngRow = lngRow + 1
                .Rows = .Rows + 1
            End If
            .RowOutlineLevel(lngRow) = 1
            .TextMatrix(lngRow, .ColIndex("���")) = GetBillTypeName(mrsListBill!Ʊ��)
            .TextMatrix(lngRow, .ColIndex("����")) = Decode(mrsListBill!����, _
                                                                2, Decode(mrsListBill!Ʊ��, 6, "�ջ�", "�˷�"), _
                                                                3, Decode(mrsListBill!Ʊ��, 6, "����", "�ش�"), "����")
            .TextMatrix(lngRow, .ColIndex("�ջ�ʱ��")) = Format(mrsListBill!����ʱ��, "yyyy-mm-dd HH:MM")
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(mrsListBill!���)), "#,###0.00;-#,###0.00;0.00;-0.00")
            If InStr(";�շ�;���ѿ�;", ";" & .TextMatrix(lngRow, .ColIndex("���")) & ";") > 0 And Val(.TextMatrix(lngRow, .ColIndex("���"))) = 0 Then
                .TextMatrix(lngRow, .ColIndex("���")) = "-"
            End If
            If Nvl(mrsListBill!��ʼƱ��) = Nvl(mrsListBill!��ֹƱ��) Then
                .TextMatrix(lngRow, .ColIndex("Ʊ�ݺ�")) = Nvl(mrsListBill!��ʼƱ��)
            Else
                .TextMatrix(lngRow, .ColIndex("Ʊ�ݺ�")) = Nvl(mrsListBill!��ʼƱ��) & "-" & Nvl(mrsListBill!��ֹƱ��)
            End If
            .Rows = .Rows + 1: lngRow = lngRow + 1
            mrsListBill.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        If .TextMatrix(.Rows - 1, .ColIndex("���")) = "" Then
            .IsSubtotal(.Rows - 1) = False
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        '�ָ�������
        zl_vsGrid_Para_Restore mlngMode, vsReturnBill, Me.Name, "�ջ�Ʊ���б�", False
    End With
    With vsBill
        .Clear 1
        .Rows = 1: lngRow = 0: .Cols = 3
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        mrsListBill.Filter = 0:  strTemp = ""
        Do While Not mrsListBill.EOF
            intƱ�� = Val(Nvl(mrsListBill!Ʊ��))
            If InStr(1, strTemp & ",", "," & intƱ�� & ",") = 0 Then
                strTemp = strTemp & "," & intƱ��
                .Rows = .Rows + 2
                .TextMatrix(.Rows - 3, 0) = GetBillTypeName(mrsListBill!Ʊ��)
                .TextMatrix(.Rows - 2, 0) = GetBillTypeName(mrsListBill!Ʊ��)
                .TextMatrix(.Rows - 3, 1) = "����Ʊ��"
                .TextMatrix(.Rows - 2, 1) = "Ʊ�ݷ�Χ"
                .Cell(flexcpData, .Rows - 3, 0, .Rows - 2, 0) = intƱ��
            End If
            mrsListBill.MoveNext
        Loop
        Dim lngBillTotal(0 To 2) As Long
        lngRow = 0
        For lngRow = 0 To .Rows - 1 Step 2
            intƱ�� = Val(.Cell(flexcpData, lngRow, 0))
            mrsListBill.Filter = "Ʊ��=" & intƱ��
            lngBillTotal(0) = 0: lngBillTotal(1) = 0: lngBillTotal(2) = 0
            Do While Not mrsListBill.EOF
               Select Case Val(Nvl(mrsListBill!����))
               Case 1 '����Ʊ��ͳ��
                    lngBillTotal(0) = lngBillTotal(0) + Val(Nvl(mrsListBill!����))
                    .TextMatrix(lngRow + 1, 2) = Trim(.TextMatrix(lngRow + 1, 2)) & _
                        IIf(Trim(.TextMatrix(lngRow + 1, 2)) = "", "", ";") & _
                        IIf(Nvl(mrsListBill!��ʼƱ��) = Nvl(mrsListBill!��ֹƱ��), _
                            Nvl(mrsListBill!��ʼƱ��), _
                            Nvl(mrsListBill!��ʼƱ��) & "-" & Nvl(mrsListBill!��ֹƱ��))
               Case 2
                    lngBillTotal(1) = lngBillTotal(1) + Val(Nvl(mrsListBill!����))
               Case 3
                    lngBillTotal(2) = lngBillTotal(2) + Val(Nvl(mrsListBill!����))
               End Select
               mrsListBill.MoveNext
            Loop
            If intƱ�� <> 0 Then
                .TextMatrix(lngRow, 2) = "ʹ��:" & lngBillTotal(0) & "��; " & _
                    Decode(intƱ��, 6, "�ջ�(�˿������պͻ���):", "�ջ�(�˷Ѻ��ش�):") & lngBillTotal(1) + lngBillTotal(2) & "��"
            End If
        Next
        .Rows = .Rows - 1
    End With
    
    LoadChargeAndBillAndTotal = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
 
Public Sub SetFontSize(ByVal bytSize As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ9����)��1-���(12��);>1: Ϊָ�����ֺ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-03 18:05:20
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������С
    '����:���˺�
    '����:2013-09-03 18:04:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Me.FontSize = mbytFontSize
    Set dkpMan.PaintManager.CaptionFont = Me.Font
    Set vsBill.Font = Me.Font
    Set vsChagre.Font = Me.Font
    Set vsReturnBill.Font = Me.Font
    txtTotal.Font = Me.Font
 End Sub
 
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2013-09-03 15:28:24
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtTotal.Text = "": txtTotal.Locked = True: txtTotal.Enabled = False
    txtTotal.FontBold = True: mbytFontSize = 9
    Call InitPanel
    Call InitGrid
    Call ReSetFontSize
 End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim objReturnPane As Pane
    Dim objChargePane As Pane
    Dim lngFilterHeight As Long, lngBillHeight As Long
    Dim lngBalanceHeight As Long
    
    lngFilterHeight = 810 / Screen.TwipsPerPixelY
    lngBillHeight = 1275 / Screen.TwipsPerPixelY
    lngBalanceHeight = (Me.ScaleHeight - 1275) \ Screen.TwipsPerPixelY - 205
    With dkpMan
        Set objChargePane = .CreatePane(mPaneIndex.EM_PN_ChargeTotal, 400, lngBalanceHeight, DockBottomOf, Nothing)
        objChargePane.Title = "�տ���Ϣ"
        objChargePane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        objChargePane.Handle = picChargeInfor.hWnd
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_ChargeTotal, 400, lngBillHeight, DockBottomOf, objChargePane)
        objPane.MinTrackSize.Height = lngBillHeight * 0.5
        objPane.Title = "Ʊ��ʹ����Ϣ"
        objPane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        objPane.Handle = picBillInfor.hWnd
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_ReprintBill, 400, lngBalanceHeight, DockRightOf, objChargePane)
        objPane.Title = "�ջ�Ʊ����Ϣ"
        objPane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        objPane.Handle = picReturnBill.hWnd

'        Set objReturnPane = .CreatePane(mPaneIndex.EM_PN_BackFeeBill, 400, 400, DockRightOf, objChargePane)
'        objReturnPane.Title = "�˷��ջ�Ʊ��"
'        objReturnPane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
'        objReturnPane.Handle = picDelFeeBill.hwnd
'
'        Set objPane = .CreatePane(mPaneIndex.EM_PN_ReprintBill, 400, 165, DockBottomOf, objReturnPane)
'        objPane.Title = "�ش��ջ�Ʊ��"
'        objPane.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
'        objPane.Handle = picRePrintBill.hwnd
        
      '  .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function
Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ���Ϣ
    '����:
    '����:���˺�
    '����:2013-09-03 11:38:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objCol As ReportColumn
    
    '�ջ�Ʊ����Ϣ
    With vsReturnBill
        Set .Font = Me.Font
        .Clear 1
        .Cols = 5: .Rows = 2
        .OutlineBar = flexOutlineBarComplete
        .OutlineCol = 0
        .SubtotalPosition = flexSTAbove
        .FixedRows = 1
        .TextMatrix(0, 0) = "���"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "�ջ�ʱ��"
        .TextMatrix(0, 3) = "���"
        .TextMatrix(0, 4) = "Ʊ�ݺ�"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            If i = .ColIndex("���") Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        .ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsReturnBill, Me.Name, "�ջ�Ʊ���б�", False
    End With
    
    '�տ������Ϣ
    With vsChagre
        Set .Font = Me.Font
        .Clear 1
        .Cols = 4: .Rows = 2
        .OutlineBar = flexOutlineBarComplete
        .OutlineCol = 1
        .SubtotalPosition = flexSTAbove
        .FixedRows = 1
        .TextMatrix(0, 0) = "���"
        .TextMatrix(0, 1) = "���㷽ʽ"
        .TextMatrix(0, 2) = "���"
        .TextMatrix(0, 3) = "�������"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            If i = .ColIndex("���") Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .ColHidden(.ColIndex("���")) = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        .ExtendLastCol = True
        zl_vsGrid_Para_Restore mlngMode, vsChagre, Me.Name, "���㷽ʽ�б�", False
    End With
    '���ʹ����Ϣ
    With vsBill
        .Clear 1
        Set .Font = Me.Font
        .Cols = 3: .Rows = 1
        .FixedRows = 0: .FixedCols = 1
        .ColAlignment(2) = flexAlignLeftCenter
        .MergeCells = flexMergeFree
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        .ExtendLastCol = True
        For i = 0 To .Rows - 1
            .MergeRow(0) = True
        Next
    End With
 End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionAttaching Then Cancel = True
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call InitFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngCashRow = 0
End Sub

Private Sub picBillInfor_Resize()
    Err = 0: On Error Resume Next
    With picBillInfor
        vsBill.Left = .ScaleLeft
        vsBill.Top = .ScaleTop
        vsBill.Height = .ScaleHeight
        vsBill.Width = .ScaleWidth
    End With
End Sub
Private Sub picChargeInfor_Resize()
    Err = 0: On Error Resume Next
    With picChargeInfor
        vsChagre.Top = .ScaleTop
        vsChagre.Left = .ScaleLeft
        lblTotal.Left = .ScaleLeft
        txtTotal.Left = lblTotal.Left + lblTotal.Width + 10
        txtTotal.Width = .ScaleWidth - txtTotal.Left
        txtTotal.Top = .ScaleHeight - txtTotal.Height
        lblTotal.Top = txtTotal.Top + (txtTotal.Height - lblTotal.Height) \ 2
        vsChagre.Height = txtTotal.Top - vsChagre.Top - 50
        vsChagre.Width = .ScaleWidth
    End With
End Sub

Private Sub picReturnBill_Resize()
    Err = 0: On Error Resume Next
    With picReturnBill
        vsReturnBill.Left = .ScaleLeft
        vsReturnBill.Top = .ScaleTop
        vsReturnBill.Height = .ScaleHeight
        vsReturnBill.Width = .ScaleWidth
    End With
End Sub

 

Private Sub vsChagre_GotFocus()
    Call zl_VsGridGotFocus(vsChagre)
End Sub

Private Sub vsChagre_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsChagre, vsChagre.BackColor)
    'On Error Resume Next
    With vsChagre
        If .TextMatrix(.RowSel, .ColIndex("���㷽ʽ")) = "ҽ�����" Then .Cell(flexcpBackColor, .RowSel, 0, .RowSel, .Cols - 1) = &H80000016
    End With
End Sub
Private Sub vsChagre_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngMode, vsChagre, Me.Name, "���㷽ʽ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsChagre_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsChagre, OldRow, NewRow, OldCol, NewCol)
    With vsChagre
        If OldRow >= .Rows - 1 Then Exit Sub
        If .TextMatrix(OldRow, .ColIndex("���㷽ʽ")) = "ҽ�����" Then .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000016
    End With
End Sub
Private Sub vsChagre_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsChagre, Me.Name, "���㷽ʽ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub vsChagre_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnOlnyView Then Cancel = True: Exit Sub
    With vsChagre
        Select Case Col
        Case .ColIndex("�������")
            If .TextMatrix(Row, .ColIndex("���㷽ʽ")) Like "*��Ԥ��*" _
                Or .TextMatrix(Row, Col) Like "*����ϼ�*" _
                Or .TextMatrix(Row, Col) Like "*����ϼ�*" _
                Or .TextMatrix(Row, .ColIndex("���㷽ʽ")) = "ҽ�����" Then
                Cancel = True: Exit Sub
            End If
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub

Private Sub vsChagre_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlVsMoveGridCell(vsChagre, vsChagre.ColIndex("���㷽ʽ"), vsChagre.Cols - 1, False)
End Sub

Private Sub vsChagre_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlVsMoveGridCell(vsChagre, vsChagre.ColIndex("���㷽ʽ"), vsChagre.Cols - 1, False)
End Sub

Private Sub vsChagre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub
Private Sub vsChagre_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsChagre
        If Row <= 1 Then Exit Sub
            VsFlxGridCheckKeyPress vsChagre, Row, Col, KeyAscii, m�ı�ʽ
            If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0
    End With
End Sub
Private Sub vsChagre_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer
    '������֤
    With vsChagre
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("�������")
            If zlCommFun.ActualLen(strKey) > 10 Then
                MsgBox "������볬��,���ֻ������10���ַ���5������", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If InStr(1, strKey, "'") > 0 Or InStr(1, strKey, "|") > 0 Or InStr(1, strKey, ",") > 0 Then
                MsgBox "��������в��ܰ��������ַ�:',| ", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsReturnBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngMode, vsReturnBill, Me.Name, "�ջ�Ʊ���б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub vsReturnBill_GotFocus()
    Call zl_VsGridGotFocus(vsReturnBill)
End Sub
Private Sub vsReturnBill_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsReturnBill, vsReturnBill.BackColor)
    'On Error Resume Next
    With vsReturnBill
        If .IsSubtotal(.RowSel) = True Then .Cell(flexcpBackColor, .RowSel, 0, .RowSel, .Cols - 1) = &H80000016
    End With
End Sub

Private Sub vsReturnBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsReturnBill, OldRow, NewRow, OldCol, NewCol)
    With vsReturnBill
        If OldRow >= .Rows - 1 Then Exit Sub
        If .IsSubtotal(OldRow) = True Then .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &H80000016
    End With
End Sub

Private Function GetBalance(ByVal strCash As String) As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ֹʱ��ε����
    '���:strCash-�ֽ���㷽ʽ
    '����:�������,����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 10:40:24
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select ���㷽ʽ, Sum(���) As ��� " & _
    "   From (  Select ���㷽ʽ, ��� As ���  From ��Ա�ɿ����  " & _
    "                Where �տ�Ա =[1] and ����=1" & _
    "                Union All " & _
    "                Select a.���㷽ʽ, -1 * Sum(a.��Ԥ��) As ���" & vbNewLine & _
    "                From ����Ԥ����¼ A" & vbNewLine & _
    "                Where Nvl(a.У�Ա�־, 0) = 0 And ����Ա���� || '' = [1] And Mod(a.��¼����, 10) <> 1 And" & vbNewLine & _
    "                   a.�տ�ʱ�� > [2] And Not Exists" & vbNewLine & _
    "                   (Select 1 From ������ü�¼ B Where a.����id = b.����id And Nvl(b.����״̬, 0) = 1) And Not Exists" & vbNewLine & _
    "                   (Select 1 From ���˽��ʼ�¼ B Where a.����id = b.Id And Nvl(b.����״̬, 0) <> 0) And Not Exists" & vbNewLine & _
    "                   (Select 1 From ���ò����¼ B Where a.����id = b.����id And Nvl(b.����״̬, 0) >= 1)" & vbNewLine & _
    "                Group By a.���㷽ʽ" & _
    "                Union All " & _
    "                Select ���㷽ʽ, -1 * nvl(sum(���),0) As ��� " & _
    "                From ����Ԥ����¼ A " & _
    "               Where ��¼���� = 1 And ����Ա���� || '' =[1] And �տ�ʱ�� > [2]  " & _
    "               Group by ���㷽ʽ "
    
    strSQL = strSQL & _
    "               Union All " & _
    "               Select a.���㷽ʽ, -1 * Nvl(Sum(a.ʵ�ս��), 0) As ��� " & _
    "               From ���˿������¼ A " & _
    "               Where a.��¼���� In(1, 2, 3) And a.����Ա���� || '' =[1] And a.�Ǽ�ʱ�� > [2]  " & _
    "               Group By ���㷽ʽ " & _
    "               Union All " & _
    "               Select a.���㷽ʽ, -1 * Nvl(Sum(a.�����), 0) As ��� " & _
    "               From ��Ա����¼ A " & _
    "               Where a.����� || '' =[1] And a.ȡ��ʱ�� Is Null And a.���ʱ�� > [2]  " & _
    "               Group By ���㷽ʽ " & _
    "               Union All " & _
    "               Select a.���㷽ʽ, Nvl(Sum(a.�����), 0) As ��� " & _
    "               From ��Ա����¼ A" & _
    "               Where a.����� || '' =[1] And a.���ʱ�� > [2]  And a.ȡ��ʱ�� Is Null " & _
    "               Group By a.���㷽ʽ " & _
    "               Union All " & _
    "               Select '" & strCash & "' As ���㷽ʽ, -1 * a.���  From ��Ա�ݴ��¼ A " & _
    "               Where �տ�Ա || '' =[1] And �Ǽ�ʱ�� > [2] ) " & _
    " Group By ���㷽ʽ  Having Sum(���) <> 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrPersonName, mdtendDate)
    Set GetBalance = rsTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 09:48:31
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, i As Long, strCaption As String
    
    strCaption = IIf(mbytType = EM_�����տ�_���շ�Ա, "�տ�", "����")
    On Error GoTo errHandle
    If mrsList Is Nothing Then
        If MsgBox("��������ص�" & strCaption & "����,��������ȡ" & strCaption & "����,�Ƿ�������ȡ" & strCaption & "����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call LoadPersonChargeAndBill
        End If
        Exit Function
    End If
    mrsList.Filter = "����=-1"
    mrsListBill.Filter = ""
    If mrsList.RecordCount = 0 And mrsListBill.RecordCount = 0 Then
        If MsgBox("��������ص�" & strCaption & "����,��������ȡ" & strCaption & "����,�Ƿ�������ȡ" & strCaption & "����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call LoadPersonChargeAndBill
        End If
        Exit Function
    End If
    
    mrsList.Filter = "����>=1"
    If mrsList.RecordCount = 0 And mrsListBill.RecordCount = 0 Then
        If MsgBox("��������ص�" & strCaption & "����,��������ȡ" & strCaption & "����,�Ƿ�������ȡ" & strCaption & "����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call LoadPersonChargeAndBill
        End If
        Exit Function
    End If
    If CheckMzFeeChargeValied = False Then Exit Function
    With vsChagre
        For i = 1 To .Rows - 1
            strTemp = .TextMatrix(i, .ColIndex("�������"))
            If zlCommFun.ActualLen(strTemp) > 10 Then
                MsgBox "������볬��,���ֻ������10���ַ���5������", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("�������")
                If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                    .TopRow = .Row: .LeftCol = .Col
                End If
                If .Visible And .Enabled Then .SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") > 0 Or InStr(1, strTemp, "|") > 0 Or InStr(1, strTemp, ",") > 0 Then
                MsgBox "��������в��ܰ��������ַ�:',| ", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("�������")
                If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                    .TopRow = .Row: .LeftCol = .Col
                End If
                If .Visible And .Enabled Then .SetFocus
                Exit Function
            End If
        Next
    End With
    CheckValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Get�շѶ���(ByRef cllData As Collection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѶ���
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 10:14:21
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    
    On Error GoTo errHandle
    Set cllData = New Collection
    
    '�սɶ���
    mrsList.Filter = "����>=1 and ����ID<>0"
    With mrsList
        .Sort = "����,����id"
        If .RecordCount <> 0 Then .MoveFirst
        strTemp = ""
        Do While Not .EOF
            '����, ����id, '' As ���㷽ʽ, 0 As ���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�
            If strTemp <> "" And zlCommFun.ActualLen(strTemp & !���� & "," & !����id) >= 4000 Then
                '����1,��¼ID1|����2,��¼ID2|...|����n,��¼IDn
                strTemp = Mid(strTemp, 2)
                cllData.Add strTemp
                strTemp = ""
            End If
            strTemp = strTemp & "|" & !���� & "," & !����id
            .MoveNext
        Loop
        If strTemp <> "" Then
            '����1,��¼ID1|����2,��¼ID2|...|����n,��¼IDn
            strTemp = Mid(strTemp, 2)
            cllData.Add strTemp
            strTemp = ""
        End If
    End With
    If cllData.Count = 0 Then Exit Function
    Get�շѶ��� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Get�ս���ϸ(ByRef cllData As Collection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѶ���
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 10:14:21
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strHaveBalance As String
    Dim strBalance As String, i As Long
    
    On Error GoTo errHandle
    Set cllData = New Collection
    
    With vsChagre
        strBalance = ""
        For i = 1 To .Rows - 1
            '���㷽ʽ1,������1,�����1,���1|���㷽ʽ2,������2,�����2,���2|...
            strTemp = .TextMatrix(i, .ColIndex("���㷽ʽ"))
            If Not (strTemp Like "*��Ԥ��*" Or strTemp Like "*����ϼ�*" _
                        Or strTemp Like "*����ϼ�*" Or strTemp = "ҽ�����") And strTemp <> "" Then
                strHaveBalance = strHaveBalance & "," & strTemp
                If i = mlngCashRow Then
                    strTemp = strTemp & "," & Val(Replace(.RowData(i), ",", "")) - mdblRemain
                Else
                    strTemp = strTemp & "," & Val(Replace(.RowData(i), ",", ""))
                End If
                strTemp = strTemp & "," & Trim(.TextMatrix(i, .ColIndex("�������")))
                mrsBalance.Filter = "���㷽ʽ='" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & "'"
                If mrsBalance.EOF Then
                    strTemp = strTemp & "," & 0
                Else
                    strTemp = strTemp & "," & Val(Nvl(mrsBalance!���))
                End If
                If zlCommFun.ActualLen(strBalance & "|" & strTemp) > 4000 Then
                    strBalance = Mid(strBalance, 2)
                    cllData.Add strBalance
                    strBalance = ""
                End If
                strBalance = strBalance & "|" & strTemp
            End If
        Next
    End With
    mrsBalance.Filter = 0
    With mrsBalance
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strTemp = Nvl(!���㷽ʽ)
            If InStr(strHaveBalance & ",", "," & strTemp & ",") = 0 And strTemp <> "" Then
                strTemp = strTemp & "," & 0
                strTemp = strTemp & "," & ""
                strTemp = strTemp & "," & Val(Nvl(mrsBalance!���))
                If zlCommFun.ActualLen(strBalance & "|" & strTemp) > 4000 Then
                    strBalance = Mid(strBalance, 2)
                    cllData.Add strBalance
                    strBalance = ""
                End If
                strBalance = strBalance & "|" & strTemp
            End If
            .MoveNext
        Loop
    End With
    If strBalance <> "" Then
        strBalance = Mid(strBalance, 2)
        cllData.Add strBalance
    End If
    Get�ս���ϸ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Function Get�ս�Ʊ��(ByRef cllData As Collection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ս�Ʊ��
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 10:14:21
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strData As String, lngNo As Long
    Dim strPre As String
    On Error GoTo errHandle
    Set cllData = New Collection
    
    '�սɶ���
    mrsListBill.Filter = "����>=1"
    With mrsListBill
        If .RecordCount = 0 Then
            Get�ս�Ʊ�� = True
            Exit Function
        End If
        .Sort = "Ʊ��,����,��ʼƱ��"
        If .RecordCount <> 0 Then .MoveFirst
        strTemp = "": strPre = "": lngNo = 0
        strData = ""
        Do While Not .EOF
            'Ʊ��,����,����, ��ʼƱ��, ��ֹƱ��,���, ����ʱ��, ����
            If strPre <> Val(Nvl(!Ʊ��)) & "-" & Val(Nvl(!����)) Then
                 strPre = Val(Nvl(!Ʊ��)) & "-" & Val(Nvl(!����))
                 lngNo = 0
            End If
            lngNo = lngNo + 1
            strTemp = Val(Nvl(!Ʊ��))
            strTemp = strTemp & "," & Val(Nvl(!����))
            strTemp = strTemp & "," & lngNo
            strTemp = strTemp & "," & Val(Nvl(!����))
            strTemp = strTemp & "," & Nvl(!��ʼƱ��)
            strTemp = strTemp & "," & Nvl(!��ֹƱ��)
            strTemp = strTemp & "," & Val(Nvl(!���))
            strTemp = strTemp & "," & Format(!����ʱ��, "yyyy-mm-dd HH:MM:SS")
            strTemp = strTemp & "," & Nvl(!����)
            
            If strTemp <> "" And zlCommFun.ActualLen(strData & "|" & strTemp) >= 4000 Then
                'Ʊ��,����,���,Ʊ������,��ʼƱ��,��ֹƱ��,���,����ʱ��,����|Ʊ��,����,���,Ʊ������,��ʼƱ��,��ֹƱ��,���,����ʱ��,����|...
                strData = Mid(strData, 2)
                cllData.Add strData
                strData = ""
            End If
            strData = strData & "|" & strTemp
            .MoveNext
        Loop
        If strData <> "" Then
            '����1,��¼ID1|����2,��¼ID2|...|����n,��¼IDn
            strData = Mid(strData, 2)
            cllData.Add strData
            strData = ""
        End If
    End With
    
    Get�ս�Ʊ�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Public Function SaveData(ByVal strStartDate As String, ByVal strEndDate As String, _
    ByVal strMemo As String, ByVal lngDeptID As Long, _
    ByRef strNO As String, ByRef lngID As Long, _
    Optional ByVal dblRemain As Double = 0, Optional strRollingType As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '���:strStartDate-��ʼ����ʱ��
    '       strEndDate-��ֹ����ʱ��
    '       strMemo-��ע
    '       lngDeptID-�տ��ID
    '       strRollingType-�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨)
    '����:strNo-NO
    '        lngID-����ID
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-09 18:00:05
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String, cllData(0 To 3) As Collection
    Dim strTemp As String, i As Long, strDate As String, strSQL1 As String
    Dim dblԤ�� As Double, dbl���ϼ� As Double, dbl����ϼ� As Double
    Dim lng�տ�ID As Long, str�տ�NO As String, rsTemp As ADODB.Recordset
    Dim lng�鳤ID As Long
    
    On Error GoTo errHandle

    If CheckValied = False Then Exit Function
    
    If mbytType <> EM_�����տ�_���շ�Ա Then
        strSQL = _
            "Select b.��id" & vbNewLine & _
            "From �������鳤���� A, �ɿ��Ա��� B, ����ɿ���� C" & vbNewLine & _
            "Where a.��id = b.��id And b.��id = c.Id And b.��Աid = [1]" & vbNewLine & _
            "      And (c.ɾ������ > Sysdate Or c.ɾ������ Is Null)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
        If Not rsTemp.EOF Then
            lng�鳤ID = frmChargeBillSel.ShowMe(Me, Val(Nvl(rsTemp!��ID)))
        Else
            lng�鳤ID = 0
        End If
    End If
    
    mrsList.Filter = "����=-1"
    mdblRemain = dblRemain
    Do While Not mrsList.EOF
        dblԤ�� = dblԤ�� + Round(Val(Nvl(mrsList!��Ԥ��)), 2)
        dbl���ϼ� = dbl���ϼ� + Round(Val(Nvl(mrsList!���ϼ�)), 2)
        dbl����ϼ� = dbl����ϼ� + Round(Val(Nvl(mrsList!����ϼ�)), 2)
        mrsList.MoveNext
    Loop
    If Get�շѶ���(cllData(0)) = False And Get�ս���ϸ(cllData(1)) = False And Get�ս�Ʊ��(cllData(2)) = False Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Set cllPro = New Collection
    lngID = zlDatabase.GetNextId("��Ա�սɼ�¼")
    strNO = zlDatabase.GetNextNo(137)
    With vsChagre
        'Zl_�շ�Ա���ʼ�¼_Insert
        strSQL = "Zl_�շ�Ա���ʼ�¼_Insert("
        '  Id_In         In ��Ա�սɼ�¼.Id%Type,
        strSQL = strSQL & "" & lngID & ","
        '  No_In         In ��Ա�սɼ�¼.No%Type,
        strSQL = strSQL & "'" & strNO & "',"
        '  �տ�Ա_In     In ��Ա�սɼ�¼.�տ�Ա%Type,
        strSQL = strSQL & "'" & mstrPersonName & "',"
        '  �տ��id_In In ��Ա�սɼ�¼.�տ��id%Type,
        strSQL = strSQL & "Null,"
        '  �鳤id_In In     ��Ա��.id%Type,
        strSQL = strSQL & ZVal(lng�鳤ID) & ","
        '  ��Ԥ����_In   In ��Ա�սɼ�¼.��Ԥ����%Type,
        strSQL = strSQL & "" & dblԤ�� & ","
        '  ����ϼ�_In   In ��Ա�սɼ�¼.����ϼ�%Type,
        strSQL = strSQL & "" & dbl���ϼ� & ","
        '  ����ϼ�_In   In ��Ա�սɼ�¼.����ϼ�%Type,
        strSQL = strSQL & "" & dbl����ϼ� & ","
        '  ժҪ_In       In ��Ա�սɼ�¼.ժҪ%Type,
        strSQL = strSQL & IIf(strMemo = "", "NULL", "'" & strMemo & "'") & ","
        '  ��ʼʱ��_In   In ��Ա�սɼ�¼.��ʼʱ��%Type,
        strSQL = strSQL & "to_Date('" & strStartDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ��ֹʱ��_In   In ��Ա�սɼ�¼.��ֹʱ��%Type,
        strSQL = strSQL & "to_Date('" & strEndDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  �Ǽ���_In     In ��Ա�սɼ�¼.�Ǽ���%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  �Ǽ�ʱ��_In   In ��Ա�սɼ�¼.�Ǽ�ʱ��%Type,
        strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  �սɱ�־_In   In ��Ա�սɼ�¼.�սɱ�־%Type,
        strSQL = strSQL & "NULL,"
        If cllData(0).Count = 0 Then
            '  �շѶ���_In   In Varchar2
            strSQL = strSQL & "NULL,"
            '  �������_In   In Integer := 0,
            '   0-�������ʼ�¼�Ͷ���;1-ֻ�������
            strSQL = strSQL & "0,"
            '  �ݴ��_In     In ��Ա�ݴ��¼.���%Type := 0,
            strSQL = strSQL & "0,"
            '  ���_In       In Varvhar2(100)
            strSQL = strSQL & "" & IIf(strRollingType = "", "0", "'" & strRollingType & "'") & ")"
            cllPro.Add strSQL
        Else
            For i = 1 To cllData(0).Count
                '  �շѶ���_In   In Varchar2
                strSQL1 = strSQL & "'" & cllData(0)(i) & "',"
                '  �������_In   In Integer := 0,
                '   0-�������ʼ�¼�Ͷ���;1-ֻ�������
                strSQL1 = strSQL1 & "" & IIf(i = 1, "0", "1") & ","
                '  �ݴ��_In     In ��Ա�ݴ��¼.���%Type := 0,
                strSQL1 = strSQL1 & "" & IIf(i = 1, dblRemain, 0) & ","
                '  ���_In       In Varvhar2(100)
                strSQL1 = strSQL1 & "" & IIf(strRollingType = "", "0", "'" & strRollingType & "'") & ")"
                cllPro.Add strSQL1
            Next
        End If
        
        '�����ս���ϸ
        For i = 1 To cllData(1).Count
            'Zl_�շ�Ա������ϸ_Insert
            strSQL = "Zl_�շ�Ա������ϸ_Insert("
            '�ս�id_In   In ��Ա�ս���ϸ.�ս�id%Type,
            strSQL = strSQL & "" & lngID & ","
            '������Ϣ_In In Varchar2
            '       ������Ϣ_IN:���㷽ʽ1,������1,�����1,���1|���㷽ʽ2,������2,�����2,���2|...
            strSQL = strSQL & "'" & cllData(1)(i) & "')"
            cllPro.Add strSQL
        Next
                
        '�����ս�Ʊ��
        For i = 1 To cllData(2).Count
            'Zl_�շ�Ա������ϸ_Insert
            strSQL = "Zl_�շ�Ա����Ʊ��_Insert("
            '�ս�id_In   In ��Ա�ս���ϸ.�ս�id%Type,
            strSQL = strSQL & "" & lngID & ","
            ' Ʊ����Ϣ_In Varchar2
            '       ��ʽ:Ʊ��,����,���,Ʊ������,��ʼƱ��,��ֹƱ��,���,����ʱ��|Ʊ��,����,���,Ʊ������,��ʼƱ��,��ֹƱ��,���,����ʱ��|...
            '  --           Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
            '  --           ����:1-����Ʊ��;2-�˷��ջ�Ʊ��;3-�ش��ջ�Ʊ��
            '  --            ����ʱ��:yyyy-mm-dd hh24:mi:ss
            strSQL = strSQL & "'" & cllData(2)(i) & "')"
            cllPro.Add strSQL
        Next
    End With
    For i = 0 To 2
        Set cllData(i) = Nothing
    Next
    If mbytType = EM_�����տ�_���շ�Ա Then
        '����Ҫ�����շ�Ա���տ��¼
        lng�տ�ID = zlDatabase.GetNextId("��Ա�սɼ�¼")
        str�տ�NO = zlDatabase.GetNextNo(140)
        'Zl_���շ�Ա�տ��¼_Insert
        strSQL = "Zl_���շ�Ա�տ��¼_Insert("
        '  Id_In         In ��Ա�սɼ�¼.Id%Type,
        strSQL = strSQL & "" & lng�տ�ID & ","
        '  No_In         In ��Ա�սɼ�¼.No%Type,
        strSQL = strSQL & "'" & str�տ�NO & "',"
        '  �տ��id_In In ��Ա�սɼ�¼.�տ��id%Type,
        strSQL = strSQL & "" & "Null,"
        '  ժҪ_In       In ��Ա�սɼ�¼.ժҪ%Type,
        strSQL = strSQL & IIf(strMemo = "", "NULL", "'" & strMemo & "'") & ","
        '  �Ǽ���_In     In ��Ա�սɼ�¼.�Ǽ���%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  �Ǽ�ʱ��_In   In ��Ա�սɼ�¼.�Ǽ�ʱ��%Type,
        strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ����id_In     In ��Ա�սɼ�¼.Id%Type
        strSQL = strSQL & "" & lngID & ")"
        cllPro.Add strSQL
    End If
    'ִ�й���
    On Error GoTo ErrCommit:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrCommit:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function

Public Sub zlPrint(ByVal bytMode As Byte, _
    Optional strDeptName As String = "", Optional strMemo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '       strDeptName-�տ������(�շ�Ա����ʱת��)
    '       strMemo-��ע(�շ�Ա����ʱת��)
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo ErrHand:
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    If mbytType = EM_�շ�Ա���� Or mbytType = EM_�����տ�_���շ�Ա Then
        objPrint.Title.Text = gstr��λ���� & "�շ�Ա�տƱ�ݻ���"
        If mlngChargeRollingID = 0 Then
            Set objRow = New zlTabAppRow
            objRow.Add "�շ�Ա��" & mstrPersonName
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            'objRow.Add "�տ��:" & strDeptName
            objRow.Add IIf(mbytType <> EM_�����տ�_���շ�Ա, "����ʱ�䣺", "�տ�ʱ��") & Format(mdtStartDate, "yyyy-mm-dd HH:MM:SS") & "��" & Format(mdtendDate, "yyyy-mm-dd HH:MM:SS")
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            objRow.Add IIf(mbytType <> EM_�����տ�_���շ�Ա, "����˵����", "�տ�ʱ��") & strMemo
            objPrint.BelowAppRows.Add objRow
        Else
            strSQL = "" & _
            "   Select /*+ rule */a.Id,a.No ,a.�տ�Ա, a.��ʼʱ��, a.��ֹʱ��, a.�Ǽ�ʱ�� ,  " & _
            "         b.���� As �տ��, a.ժҪ,M.������ as ������  " & _
            "  From ��Ա�սɼ�¼ A, ���ű� B,����ɿ���� M" & _
            "  Where a.�տ��id = b.Id(+) and a.�ɿ���ID=M.ID(+) And a.ID=[1]  " & _
            "  Order by �Ǽ�ʱ��,NO desc"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID)
            If rsTemp.EOF Then Exit Sub
            Set objRow = New zlTabAppRow
            objRow.Add "�շ�Ա��" & rsTemp!�տ�Ա
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            objRow.Add "���ʵ��ţ�" & Nvl(rsTemp!NO)
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            objRow.Add "����ʱ�䣺" & Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            'objRow.Add "�տ��:" & rsTemp!�տ��
            objRow.Add "����ʱ�䣺" & Format(rsTemp!��ʼʱ��, "yyyy-mm-dd HH:MM:SS") & "��" & Format(rsTemp!��ֹʱ��, "yyyy-mm-dd HH:MM:SS")
            objPrint.UnderAppRows.Add objRow
            Set objRow = New zlTabAppRow
            objRow.Add "����˵��:" & rsTemp!ժҪ
            objPrint.BelowAppRows.Add objRow
        End If
    ElseIf mbytType = EM_С���տ� Or mbytType = EM_С������ Then
        strSQL = "" & _
        "   Select /*+ rule */a.Id,a.No , a.��ʼʱ��, a.��ֹʱ��, a.�Ǽ�ʱ�� ,  " & _
        "         b.���� As �տ��, a.ժҪ,M.������ as ������  " & _
        "  From ��Ա�սɼ�¼ A, ���ű� B,����ɿ���� M" & _
        "  Where a.�տ��id = b.Id(+) and a.�ɿ���ID=M.ID(+) And a.ID=[1]  " & _
        "  Order by �Ǽ�ʱ��,NO desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID)
        If rsTemp.EOF Then Exit Sub
        objPrint.Title.Text = gstr��λ���� & "�������տƱ�ݻ���"
        Set objRow = New zlTabAppRow
        objRow.Add "С�鸺���ˣ�" & UserInfo.����
        'objRow.Add "�տ��:" & Nvl(rsTemp!�տ��)
        objRow.Add "������:" & Nvl(rsTemp!������)
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add IIf(mbytType = EM_С������, "���ʵ���:", "�տ�ţ�") & Nvl(rsTemp!NO)
        If mbytType = EM_С������ Then
            objRow.Add "����ʱ�䣺" & Format(rsTemp!��ʼʱ��, "yyyy-mm-dd HH:MM:SS") & "��" & Format(rsTemp!��ֹʱ��, "yyyy-mm-dd HH:MM:SS")
        Else
            objRow.Add "�տ�ʱ�䣺" & Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
        End If
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add IIf(mbytType = EM_С������, "����˵��:", "�տ�˵��:") & Nvl(rsTemp!��ע)
        objPrint.BelowAppRows.Add objRow
    ElseIf mbytType = EM_�����տ� Then
        objPrint.Title.Text = gstr��λ���� & "�����տƱ�ݻ���"
       strSQL = "" & _
        "   Select /*+ rule */a.Id,a.No , a.��ʼʱ��, a.��ֹʱ��,  a.�Ǽ�ʱ�� ,  " & _
        "         b.���� As �տ��, a.ժҪ " & _
        "  From ��Ա�սɼ�¼ A, ���ű� B " & _
        "  Where a.�տ��id = b.Id(+) and a.�ɿ���ID=M.ID(+) And a.ID=[1]  " & _
        "  Order by �Ǽ�ʱ��,NO desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngChargeRollingID)
        If rsTemp.EOF Then Exit Sub
        objPrint.Title.Text = gstr��λ���� & "�������տƱ�ݻ���"
        Set objRow = New zlTabAppRow
        objRow.Add "С�鸺���ˣ�" & UserInfo.����
        'objRow.Add "�տ��:" & Nvl(rsTemp!�տ��)
        objRow.Add "������:" & Nvl(rsTemp!������)
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "�տ�ţ�" & Nvl(rsTemp!NO)
        objRow.Add "�տ�ʱ�䣺" & Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "�տ�˵��:" & Nvl(rsTemp!��ע)
        objPrint.BelowAppRows.Add objRow
    Else
        Exit Sub
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    '��װ����
    With vsRptPrint
        .Clear: .Redraw = flexRDNone
        .Rows = 1: .Cols = 5: lngRow = 0: .FixedRows = 0
        '��װ�տ��¼
        .TextMatrix(lngRow, 0) = "�տ���Ϣ"
        .TextMatrix(lngRow, 1) = "�տ���Ϣ"
        .TextMatrix(lngRow, 2) = "�տ���Ϣ"
        .TextMatrix(lngRow, 3) = "�տ���Ϣ"
        .TextMatrix(lngRow, 4) = "�տ���Ϣ"
        
        lngRow = lngRow + 1: .Rows = .Rows + 1
        .TextMatrix(lngRow, 0) = "���㷽ʽ"
        .TextMatrix(lngRow, 1) = "���㷽ʽ"
        .TextMatrix(lngRow, 2) = "���㷽ʽ"
        .TextMatrix(lngRow, 3) = "���"
        .TextMatrix(lngRow, 4) = "�����"
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = Me.BackColor
        .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = 4
        
        lngRow = lngRow + 1: .Rows = .Rows + 1
        blnFind = False
        For i = 1 To vsChagre.Rows - 1
            strTemp = Trim(vsChagre.TextMatrix(i, vsChagre.ColIndex("���㷽ʽ")))
            If strTemp <> "" And strTemp <> "ҽ�����" Then
                .TextMatrix(lngRow, 0) = strTemp
                .TextMatrix(lngRow, 1) = strTemp
                .TextMatrix(lngRow, 2) = strTemp
                .TextMatrix(lngRow, 3) = Trim(vsChagre.TextMatrix(i, vsChagre.ColIndex("���")))
                .TextMatrix(lngRow, 4) = Trim(vsChagre.TextMatrix(i, vsChagre.ColIndex("�������")))
                blnFind = True
                .Rows = .Rows + 1: lngRow = lngRow + 1
            End If
        Next
        '�ϼ���Ϣ
        .TextMatrix(lngRow, 0) = "�ϼ�"
        .Cell(flexcpAlignment, lngRow, 0, lngRow, 0) = 4
        .TextMatrix(lngRow, 1) = txtTotal.Text
        .TextMatrix(lngRow, 2) = txtTotal.Text
        .TextMatrix(lngRow, 3) = txtTotal.Text
        .TextMatrix(lngRow, 4) = txtTotal.Text
        lngRow = lngRow + 1: .Rows = .Rows + 1
        
        'Ʊ��ʹ����Ϣ
        .TextMatrix(lngRow, 0) = "Ʊ��ʹ����Ϣ"
        .TextMatrix(lngRow, 1) = "Ʊ��ʹ����Ϣ"
        .TextMatrix(lngRow, 2) = "Ʊ��ʹ����Ϣ"
        .TextMatrix(lngRow, 3) = "Ʊ��ʹ����Ϣ"
        .TextMatrix(lngRow, 4) = "Ʊ��ʹ����Ϣ"
        
        lngRow = lngRow + 1: .Rows = .Rows + 1
        blnFind = False
        For i = 0 To vsBill.Rows - 1
            strTemp = Trim(vsBill.TextMatrix(i, 0))
            If strTemp <> "" Then
                .TextMatrix(lngRow, 0) = strTemp
                .TextMatrix(lngRow, 1) = Trim(vsBill.TextMatrix(i, 1))
                .Cell(flexcpAlignment, lngRow, 0, lngRow, 1) = 4
                .TextMatrix(lngRow, 2) = Trim(vsBill.TextMatrix(i, 2))
                .TextMatrix(lngRow, 3) = Trim(vsBill.TextMatrix(i, 2))
                .TextMatrix(lngRow, 4) = Trim(vsBill.TextMatrix(i, 2))
                blnFind = True
                .Rows = .Rows + 1: lngRow = lngRow + 1
            End If
        Next
        If blnFind = False Then
            .TextMatrix(lngRow, 0) = Space(1)
            .TextMatrix(lngRow, 1) = Space(2)
            .TextMatrix(lngRow, 2) = Space(3)
            .TextMatrix(lngRow, 3) = Space(3)
            .TextMatrix(lngRow, 4) = Space(3)
            .Rows = .Rows + 1: lngRow = lngRow + 1
        End If
        blnFind = False
        '�˷ѻ�����Ϣ
        .TextMatrix(lngRow, 0) = "�˷ѻ���Ʊ��"
        .TextMatrix(lngRow, 1) = "�˷ѻ���Ʊ��"
        .TextMatrix(lngRow, 2) = "�˷ѻ���Ʊ��"
        .TextMatrix(lngRow, 3) = "�˷ѻ���Ʊ��"
        .TextMatrix(lngRow, 4) = "�˷ѻ���Ʊ��"
        lngRow = lngRow + 1: .Rows = .Rows + 1
        
        .TextMatrix(lngRow, 0) = "���"
        .TextMatrix(lngRow, 1) = "�˷�ʱ��"
        .TextMatrix(lngRow, 2) = "�˷�ʱ��"
        .TextMatrix(lngRow, 3) = "�˷ѽ��"
        .TextMatrix(lngRow, 4) = "Ʊ�ݺ�"
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = Me.BackColor
        .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = 4
        lngRow = lngRow + 1: .Rows = .Rows + 1
        
        For i = 1 To vsReturnBill.Rows - 1
            If vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("����")) = "�˷�" Then
                strTemp = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("���")))
                If strTemp <> "" Then
                    .TextMatrix(lngRow, 0) = strTemp
                    .Cell(flexcpAlignment, lngRow, 0, lngRow, 0) = 4
                    .TextMatrix(lngRow, 1) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("�ջ�ʱ��")))
                    .TextMatrix(lngRow, 2) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("�ջ�ʱ��")))
                    .TextMatrix(lngRow, 3) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("���")))
                    .TextMatrix(lngRow, 4) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("Ʊ�ݺ�")))
                    blnFind = True
                    .Rows = .Rows + 1: lngRow = lngRow + 1
                    blnFind = True
                End If
            End If
        Next
        If blnFind = False Then
            .TextMatrix(lngRow, 0) = Space(1)
            .TextMatrix(lngRow, 1) = Space(2)
            .TextMatrix(lngRow, 2) = Space(2)
            .TextMatrix(lngRow, 3) = Space(3)
            .TextMatrix(lngRow, 4) = Space(4)
            .Rows = .Rows + 1: lngRow = lngRow + 1
        End If
        blnFind = False
        '�ش������Ϣ
        .TextMatrix(lngRow, 0) = "�ش������Ϣ"
        .TextMatrix(lngRow, 1) = "�ش������Ϣ"
        .TextMatrix(lngRow, 2) = "�ش������Ϣ"
        .TextMatrix(lngRow, 3) = "�ش������Ϣ"
        .TextMatrix(lngRow, 4) = "�ش������Ϣ"
        lngRow = lngRow + 1: .Rows = .Rows + 1
       .TextMatrix(lngRow, 0) = "���"
        .TextMatrix(lngRow, 1) = "�ش�ʱ��"
        .TextMatrix(lngRow, 2) = "�ش�ʱ��"
        .TextMatrix(lngRow, 3) = "�ش���"
        .TextMatrix(lngRow, 4) = "Ʊ�ݺ�"
        .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = Me.BackColor
        .Cell(flexcpAlignment, lngRow, 0, lngRow, .Cols - 1) = 4
        lngRow = lngRow + 1: .Rows = .Rows + 1
        
        For i = 1 To vsReturnBill.Rows - 1
            If vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("����")) = "�ش�" Then
                strTemp = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("���")))
                If strTemp <> "" Then
                    .TextMatrix(lngRow, 0) = strTemp
                    .Cell(flexcpAlignment, lngRow, 0, lngRow, 0) = 4
                    .TextMatrix(lngRow, 1) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("�ջ�ʱ��")))
                    .TextMatrix(lngRow, 2) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("�ջ�ʱ��")))
                    .TextMatrix(lngRow, 3) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("���")))
                    .TextMatrix(lngRow, 4) = Trim(vsReturnBill.TextMatrix(i, vsReturnBill.ColIndex("Ʊ�ݺ�")))
                    .Rows = .Rows + 1: lngRow = lngRow + 1
                    blnFind = True
                End If
            End If
        Next
        If blnFind = False Then
            .TextMatrix(lngRow, 0) = Space(1)
            .TextMatrix(lngRow, 1) = Space(2)
            .TextMatrix(lngRow, 2) = Space(2)
            .TextMatrix(lngRow, 3) = Space(3)
            .TextMatrix(lngRow, 4) = Space(4)
            .Rows = .Rows + 1: lngRow = lngRow + 1
        End If
        .Rows = .Rows - 1
       ' .AutoSizeMode = flexAutoSizeColWidth
        '.AutoSize 0, .Cols - 1
        For i = 0 To .Rows - 1
            .MergeRow(i) = True
            .RowHeight(i) = 350
        Next
        For i = 0 To .Cols - 1
            If i = 0 Then .ColWidth(i) = 800
            If i = 1 Then .ColWidth(i) = 1000
            If i = 2 Then .ColWidth(i) = 800
            If i = 3 Then .ColWidth(i) = 1400
            If i = 4 Then .ColWidth(i) = 5000
            .MergeCol(i) = True
        Next
        .MergeCells = flexMergeRestrictRows
        .Redraw = flexRDDirect
    End With
    
    Set objPrint.Body = vsRptPrint
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Property Get GetCashMoney() As Double
    '��ȡ�ֽ���
    With vsChagre
        If mlngCashRow < 1 Or mlngCashRow > .Rows - 1 Then GetCashMoney = 0: Exit Property
        GetCashMoney = Val(Replace(.TextMatrix(mlngCashRow, .ColIndex("���")), ",", ""))
    End With
End Property

Private Function CheckMzFeeChargeValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������շѵķ��úϼ������ϼ��Ƿ�Ϸ�
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-22 10:27:31
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, cllData As Collection
    '       cllData -Array(����, NO, ��¼״̬, ������, ��Ԥ��)
    '                     ����=1(�������ȷ;2.�쳣����)
    On Error GoTo errHandle
    
    If mrsList Is Nothing Then Exit Function
    If mrsList.State <> 1 Then Exit Function
    mrsList.Filter = "����=1"
    If mrsList.RecordCount = 0 Then GoTo GoSucces:
    Set cllData = New Collection
    With mrsList
        .Sort = "����id": .MoveFirst
        strTemp = ""
        Do While Not .EOF
            '����, ����id, '' As ���㷽ʽ, 0 As ���, 0 As ��Ԥ��, 0 As ���ϼ�, 0 As ����ϼ�
            If strTemp <> "" And zlCommFun.ActualLen(strTemp & "," & !����id) >= 4000 Then
                '����ID1,����ID2,...
                strTemp = Mid(strTemp, 2)
                Call CheckMzFeeValied(strTemp, cllData)
                strTemp = ""
            End If
            strTemp = strTemp & "," & !����id
            .MoveNext
        Loop
        If strTemp <> "" Then
            '����ID1,����ID2,...
            strTemp = Mid(strTemp, 2)
            Call CheckMzFeeValied(strTemp, cllData)
            strTemp = ""
        End If
    End With
    'cllData:Array(����, NO, ��¼״̬, ������, ��Ԥ��)
    If cllData.Count <> 0 Then
        If frmErrInfor.ShowErrInfor(Me, cllData) = False Then Exit Function
    End If
GoSucces:
    mrsList.Filter = 0
    CheckMzFeeChargeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckMzFeeValied(ByVal strIDs As String, ByRef cllData As Collection) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ������ID�����������Ԥ���Ƿ�Ϸ�
    '���:strIDs-����IDs,��ʽΪ:����ID1,����ID2,...
    '����:cllData-array(����,NO,��¼״̬,������,��Ԥ��)
    '                     ����=1(�������ȷ;2.�쳣����)
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-22 10:57:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = " " & _
    "   Select * " & _
    "   From (With c_�쳣 As (Select /*+cardinality(j,10)*/ a.����id, Max(a.��¼״̬) As ��¼״̬, Sum(a.���ʽ��) As ������ " & _
    "                       From ������ü�¼ A, Table(f_Num2list([1])) J " & _
    "                       Where a.����id = j.Column_Value And MOD(a.��¼����,10) = 1 " & _
    "                       Group By a.����id " & _
    "                       Having Nvl(Sum(a.���ʽ��), 0) <> (Select Nvl(Sum(��Ԥ��), 0) " & _
    "                                                     From ����Ԥ����¼ M " & _
    "                                                     Where a.����id = m.����id)) " & _
    "          Select a.����id, b.No, Max(a.��¼״̬) As ��¼״̬, Max(a.������) As ������, Sum(m.��Ԥ��) As ��Ԥ�� " & _
    "          From c_�쳣 A, ������ü�¼ B," & _
    "               (Select B1.����id, Nvl(Sum(B1.��Ԥ��), 0) As ��Ԥ�� " & _
    "                 From ����Ԥ����¼ B1, c_�쳣 C " & _
    "                 Where B1.����id = c.����id " & _
    "                 Group By B1.����id) M " & _
    "          Where a.����id = m.����id(+) And a.����id = b.����id(+) " & _
    "          Group By a.����id, b.No) " & _
    "          Order By NO "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
    With rsTemp
        strSQL = ""
        Do While Not .EOF
            '����,NO,��¼״̬,������,��Ԥ��
            cllData.Add Array(1, Nvl(!NO), Val(Nvl(!��¼״̬)), Val(Nvl(!������)), Val(Nvl(!��Ԥ��)))
            .MoveNext
        Loop
    End With
    CheckMzFeeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
