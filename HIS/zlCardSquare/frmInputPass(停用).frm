VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmInputPass 
   BorderStyle     =   0  'None
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPati 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   15
      ScaleHeight     =   765
      ScaleWidth      =   7515
      TabIndex        =   15
      Top             =   0
      Width           =   7515
      Begin VB.Frame fraSplitPati 
         Height          =   90
         Left            =   0
         TabIndex        =   17
         Top             =   675
         Width           =   7515
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6270
         TabIndex        =   16
         Top             =   435
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   420
         Width           =   600
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��ˢ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   5745
         TabIndex        =   20
         Top             =   15
         Width           =   1635
      End
      Begin VB.Label lblMargin 
         BackStyle       =   0  'Transparent
         Caption         =   "δˢ����:3000"
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
         Height          =   360
         Left            =   2955
         TabIndex        =   19
         Top             =   405
         Width           =   1800
      End
      Begin VB.Label lblPatiName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:����"
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
         Left            =   270
         TabIndex        =   18
         Top             =   75
         Width           =   1155
      End
   End
   Begin VB.PictureBox picPassWord 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   0
      ScaleHeight     =   1785
      ScaleWidth      =   7515
      TabIndex        =   4
      Top             =   795
      Width           =   7515
      Begin VB.CommandButton cmdOK 
         Caption         =   "���(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   9
         Top             =   45
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   8
         Top             =   60
         Width           =   4350
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1455
         TabIndex        =   7
         Top             =   1170
         Width           =   4305
      End
      Begin VB.TextBox txtMoney 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1455
         TabIndex        =   5
         Top             =   645
         Width           =   1740
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   6
         Top             =   630
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   765
         TabIndex        =   14
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   735
         TabIndex        =   13
         Top             =   90
         Width           =   570
      End
      Begin VB.Label lblMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   735
         Width           =   1140
      End
      Begin VB.Label lblBalance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3495
         TabIndex        =   11
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblBalanceMoney 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   4035
         TabIndex        =   10
         Top             =   645
         Width           =   1740
      End
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   360
      Top             =   2235
   End
   Begin VB.PictureBox picBlance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   1710
      Left            =   0
      ScaleHeight     =   1680
      ScaleWidth      =   7470
      TabIndex        =   0
      Top             =   2610
      Visible         =   0   'False
      Width           =   7500
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8745
         TabIndex        =   1
         Top             =   60
         Width           =   1080
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBlance 
         Height          =   1545
         Left            =   0
         TabIndex        =   2
         Top             =   -15
         Width           =   7470
         _cx             =   13176
         _cy             =   2725
         Appearance      =   3
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInputPass.frx":0000
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
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   6555
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7785
      _Version        =   589884
      _ExtentX        =   13732
      _ExtentY        =   11562
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin VB.Shape shpRange 
      BorderStyle     =   6  'Inside Solid
      Height          =   330
      Left            =   7890
      Top             =   4995
      Width           =   210
   End
End
Attribute VB_Name = "frmInputPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'--��ڲ���
Private mdbl���������ܶ� As Double, mlngModule As Long
Private mlngCardTypeID As Long, mbln���ѿ� As Boolean
Private mstrOutCardNo As String, mstrOutPassWord As String
Private mbln�˷� As Boolean, mbln���� As Boolean
Private mstrPatiName As String, mstrSex As String, mstrOld As String
Private mblnShowclsPatientInfo As Boolean
Private mcurCardObject As clsCardObject
Private mstr������Դ As String, mlng����ID As Long
'---------------------------------------------------------------------
Private mstrCardNo As String, mstrPassWord As String
Private mstrOldCardNo As String '�ɵĿ���
Private mblnFirst As Boolean, mblnOk As Boolean
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mintPassInputCount As Integer   '��������ͳ��
Private mdbl�ʻ���� As Double, mstr������� As String
Private mrsClassMoney As ADODB.Recordset
Private msngOldX As Single, msngOldY As Single
Private mblnReadCard As Boolean
Private mblnPassInputCardNo As Boolean '�Ƿ��������뿨��
Private mobjKeyboard As Object '�����������
Private mdbl���οۿ�� As Double
Private mbln�����ֹ As Boolean
Private mlng���ѿ�ID As Long
Private mdbl��ˢ�ܶ� As Double
Private mblnתԤ�� As Boolean
Private mstrCurFeeType As String '��ǰ���շ����
Private mblnAllPay As Boolean
Private mblnPosPass As Boolean
'ˢ��������Ϣ,ȷ��ʱ����(ֻ�ޱ���ˢ��������);������˷�ʱ,�����˷ѵ�ԭʼ����
Private mVarData As Variant 'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mobjSquare As Object
'---------------------------------------------------------------------
Public Function zlBrushPay(frmMain As Object, _
    ByVal lngModule As Long, _
    ByVal objCardObject As clsCardObject, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strPatiName As String, ByVal strSex As String, _
    ByVal strOld As String, ByRef dbl����ˢ����� As Double, _
    Optional ByRef strCardNo As String, _
    Optional ByRef strPassWord As String, _
    Optional ByRef bln�˷� As Boolean = False, _
    Optional ByRef blnShowclsPatientInfo As Boolean = False, _
    Optional ByRef bln���� As Boolean = False, _
    Optional ByVal bln�����ֹ As Boolean = True, _
    Optional ByRef varBrushCardData As Variant = Nothing, _
    Optional ByVal blnתԤ�� As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal str������Դ As String, _
    Optional ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������
    '���:frmMain-���õ�������
    '        lngCardTypeID-�����ID(0-��ʾֻˢһ��ͨ)
    '        rsClassMoney-�շ����,ʵ�ս��
    '        blnShowclsPatientInfo-�Ƿ���ʾ������Ϣ
    '       dbl���-������Ҫ�ۿ�Ľ��
    '       bln�����ֹ-����ʱ,��ֹ��������,�����ʾ��ʾ�����֧��
    '       VarBrushCardData-Collection����,�Ѿ�ˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����,ʣ��δ�˽��(ֻ����˷�)))
    '       blnAllPay-�Ƿ����ȫ֧����true-����δ֧���겻����ɽ��㣬false-����ֻ֧�����ֲ�����
    '       str������Դ - ��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
    '       lng����ID - ����ID(ʹ�����ѿ�֧��ʱ����)
    '����:strCardNO-���ؿ���
    '        strPassWord-�������������
    '        bln����-�Ƿ񽫵�ǰ��ˢ����������
    '        dbl���-���ر��εĿۿ���
    '        str������� -�������(���ѿ�����)
    '        lng���ѿ�ID-���ѿ���Ϣ.ID(���ѿ�����)
    '       varBrushCardData-Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-10 12:54:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intMouse As Integer
    Screen.MousePointer = 0: mlngCardTypeID = lngCardTypeID
    mbln���ѿ� = bln���ѿ�: mdbl���������ܶ� = dbl����ˢ�����
    mblnתԤ�� = blnתԤ��: mblnAllPay = blnAllPay
    Set mrsClassMoney = rsClassMoney
    mblnShowclsPatientInfo = blnShowclsPatientInfo: mbln�˷� = bln�˷�
    mstrPatiName = strPatiName
    mstrSex = strSex: mstrOld = strOld: mbln���� = False
    If IsEmpty(varBrushCardData) Then
        Set mVarData = Nothing
    Else
        Err = 0: On Error Resume Next
        Set mVarData = varBrushCardData '�Ѿ�ˢ������
        If Err <> 0 Then
            Set mVarData = Nothing
        End If
    End If
    mstrOldCardNo = strCardNo: mbln�����ֹ = bln�����ֹ
    Set mcurCardObject = objCardObject
    mblnOk = False: intMouse = Screen.MousePointer
    strCardNo = ""
    mlngModule = lngModule
    mstr������Դ = str������Դ: mlng����ID = lng����ID
    
    On Error GoTo 0
    'IC������
    On Error Resume Next
    'Set mobjICCard = CreateObject("zlICCard.clsICCard")
    On Error GoTo 0
    Me.Show 1, frmMain
    zlBrushPay = mblnOk
    strCardNo = mstrCardNo: strPassWord = mstrPassWord
    dbl����ˢ����� = mdbl���οۿ��
    Set varBrushCardData = mVarData '����ˢ������
    Screen.MousePointer = intMouse
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = intMouse
End Function

Private Sub InitTaskPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��InitTaskPancel
    '����:���˺�
    '����:2011-06-30 18:20:30
    '����:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    If mblnShowclsPatientInfo Then
        Call wndTaskPanel.SetGroupInnerMargins(0, 2, 0, 0)
    Else
        picPati.Visible = False
    End If
    
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    Set tkpGroup = wndTaskPanel.Groups.Add(1, "��ˢ������������")
    If mblnShowclsPatientInfo Then
        Set Item = tkpGroup.Items.Add(101, "", xtpTaskItemTypeControl)
        Set Item.Control = picPati
        picPati.BackColor = Item.BackColor
        Call Item.SetMargins(0, -19, 0, IIf(mblnShowclsPatientInfo, -10, -4))
    End If
    Set Item = tkpGroup.Items.Add(102, "", xtpTaskItemTypeControl)
    Set Item.Control = picPassWord
    tkpGroup.CaptionVisible = False
    If mblnShowclsPatientInfo Then
        Call Item.SetMargins(0, 20, 0, 0)
    Else
        Call Item.SetMargins(0, -19, 0, -4)
    End If
    picPassWord.BackColor = Item.BackColor
    'picPati.BackColor = Item.BackColor
    fraSplitPati.BackColor = Item.BackColor
    chk����.BackColor = Item.BackColor
    tkpGroup.Expandable = False
    vsBlance.BackColor = Item.BackColor
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
End Sub
Private Sub AddWndTaskPancelExpend()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ����Ϣ�б����չ����
    '����:���˺�
    '����:2013-02-22 15:53:55
    '����:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    On Error GoTo errHandle
    
    Set tkpGroup = wndTaskPanel.Groups.Find(2)
    '���ڣ����˳�
    If Not tkpGroup Is Nothing Then Exit Sub
    picBlance.Visible = True
    Set tkpGroup = wndTaskPanel.Groups.Add(2, "��ǰˢ����Ϣ")
    Set Item = tkpGroup.Items.Add(201, "", xtpTaskItemTypeControl)
    Set Item.Control = picBlance
    'Call Item.SetMargins(0, -19, 0, IIf(mblnShowclsPatientInfo, -10, -4))
    tkpGroup.CaptionVisible = True
    tkpGroup.Expandable = True
    picBlance.BackColor = Item.BackColor
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
    '������ɺ�ȡ������ʾ
    cmdOK.Visible = True: cmdCancel.Visible = True
    Call SetWindowHeight
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetWindowHeight()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô���ĸ߶�
    '����:���˺�
    '����:2013-02-22 15:55:37
    '����:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngHeight As Single, sngSplit As Single
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    '����λ��,��On Error resume next ���ο��ܳ��ֵĴ���
    Err = 0: On Error Resume Next
    sngSplit = 700
    sngHeight = picPassWord.Height + sngSplit
    If mblnShowclsPatientInfo Then
        sngHeight = sngHeight + picPati.Height
    End If
    Set tkpGroup = wndTaskPanel.Groups.Find(2)
    If tkpGroup Is Nothing Then Me.Height = sngHeight: Exit Sub
    If tkpGroup.Expanded Then
        sngHeight = sngHeight + picBlance.Height + IIf(mblnShowclsPatientInfo = False, 200, 0)
    End If
    sngHeight = sngHeight + 550
    Me.Height = sngHeight
End Sub

Private Sub cmdCancel_Click()
    Dim cllBalance As Collection
    Set mVarData = cllBalance
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txt���� <> "" Then
        If Not zlSquareAffirm(True) Then Exit Sub
    End If
    Call SetReturnBrushCardInfor
    mblnOk = True
    Unload Me
End Sub
Private Function SetReturnBrushCardInfor() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���÷��ص�ˢ����Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-02-25 15:29:54
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBrushCardInfor As Collection, lngCardTypeID As Long, dblMoney As Double
    Dim lng���ѿ�ID As Long, strCardNo As String, strPassWord As String, str������� As String
    Dim int���� As Integer
    Dim i As Long
    On Error GoTo errHandle
    Set cllBrushCardInfor = New Collection
    With vsBlance
        mdbl���οۿ�� = 0
        For i = .Rows - 1 To 1 Step -1
            strCardNo = Trim(.Cell(flexcpData, i, .ColIndex("����")))
            If strCardNo <> "" And Val(.RowData(i)) = 0 Then
                lngCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID")))
                lng���ѿ�ID = Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                strPassWord = Trim(.TextMatrix(i, .ColIndex("����")))
                str������� = Trim(.TextMatrix(i, .ColIndex("�������")))
                dblMoney = Val(.TextMatrix(i, .ColIndex("ˢ�����")))
                int���� = Val(.TextMatrix(i, .ColIndex("����������ʾ")))
                '�������һ�ε�ˢ����Ϣ
                mlngCardTypeID = lngCardTypeID
                mstrCardNo = strCardNo
                mbln���ѿ� = True
                mlng���ѿ�ID = lng���ѿ�ID
                mstr������� = str�������
                mdbl���οۿ�� = mdbl���οۿ�� + dblMoney
                'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
                cllBrushCardInfor.Add Array(lngCardTypeID, lng���ѿ�ID, dblMoney, strCardNo, strPassWord, str�������, int����)
            End If
        Next
    End With
    Set mVarData = cllBrushCardInfor
    SetReturnBrushCardInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me: Exit Sub
    If KeyCode = vbKeyF2 And lbl����.BorderStyle = 1 Then ClickReadCard: Exit Sub
End Sub

Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؿؼ���Ϣ
    '����:���˺�
    '����:2013-02-22 16:02:31
    '����:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstr������� = "": mlng���ѿ�ID = 0
    mdbl���οۿ�� = mdbl���������ܶ�   '�ȳ�ʼ��
    mstrCardNo = "": mstrPassWord = "":   mstrCurFeeType = Get�շ��������_����
    picPati.Visible = mblnShowclsPatientInfo
    lblPatiName.Caption = "����:" & mstrPatiName
    lblSex.Caption = "�Ա�:" & mstrSex
    lblMoney.Caption = IIf(mbln�˷�, "�����˿�", "��������")
    lblMoney.ForeColor = IIf(mbln�˷�, vbRed, lbl����.ForeColor)
    lblMargin.Visible = Not mbln�˷�
    If Not mcurCardObject Is Nothing Then
        chk����.Visible = mbln�˷� And mcurCardObject.CardPreporty.�Ƿ����� = 1
        lblType.Caption = "��ˢ" & mcurCardObject.CardPreporty.����
    Else
        lblType.Caption = "": chk����.Visible = False
    End If
    If mlngCardTypeID = 0 Then
        txt����.Locked = True: lbl����.BorderStyle = 1
        lbl����.Enabled = False: txtPass.Enabled = False
    End If
    cmdOK.Enabled = False
    txtMoney.Enabled = False
    Call LoadBruhCardInfor
    Call ShowMoney
    
    '��ʼ������ :276-���ѿ�ˢ�������붨λ�������
    mblnPosPass = Val(zlDatabase.GetPara(Val("276-���ѿ�ˢ�������붨λ�������"), glngSys, , "1")) = 1
End Sub
Private Sub LoadBruhCardInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ѿ�ˢ������Ϣ
    '����:���˺�
    '����:2013-02-25 15:58:01
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As Collection, strCardNo As String
    Dim i As Long
    Dim lngRow As Long
    
    On Error GoTo errHandle
    mdbl��ˢ�ܶ� = 0
    With vsBlance
        .Rows = 2
        .Clear 1
        .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .RowData(1) = 0
        If IsEmpty(mVarData) Then Exit Sub
        If mVarData Is Nothing Then Exit Sub
        
        Err = 0: On Error Resume Next
        Set cllBalance = mVarData
        If Err <> 0 Then
            Err = 0: On Error GoTo 0: Exit Sub
        End If
        Err = 0: On Error GoTo errHandle:
        lngRow = 1
        For i = 1 To cllBalance.Count
            'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����,ʣ���˿���)
            If Val(cllBalance(i)(2)) <> 0 Then
                .RowData(lngRow) = 1
                .TextMatrix(lngRow, .ColIndex("�����ID")) = Val(cllBalance(i)(0))
                .TextMatrix(lngRow, .ColIndex("���ѿ�ID")) = Val(cllBalance(i)(1))
                .TextMatrix(lngRow, .ColIndex("ˢ�����")) = Format(Val(cllBalance(i)(2)), "0.00")
                mdbl��ˢ�ܶ� = mdbl��ˢ�ܶ� + Val(cllBalance(i)(2))
                .TextMatrix(lngRow, .ColIndex("����������ʾ")) = Val(cllBalance(i)(6))
                strCardNo = Trim(cllBalance(i)(3))
                If Val(cllBalance(i)(6)) = 1 Then
                    .TextMatrix(lngRow, .ColIndex("����")) = String(Len(strCardNo), "*")
                Else
                    .TextMatrix(lngRow, .ColIndex("����")) = strCardNo
                End If
                .Cell(flexcpData, lngRow, .ColIndex("����")) = strCardNo
                .TextMatrix(lngRow, .ColIndex("����")) = Trim(cllBalance(i)(4))
                .TextMatrix(lngRow, .ColIndex("�������")) = Trim(cllBalance(i)(5))
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
        Next
        If .Rows > 2 Then
            .Rows = .Rows - 1
        End If
    End With
    
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ShowMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ֧�����
    '����:���˺�
    '����:2012-02-24 14:19:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double '��ˢ��
    On Error GoTo errHandle
    dblMoney = GetBruhMoney
    If mdbl���������ܶ� <> dblMoney And mblnAllPay Then
        cmdOK.Visible = False
    End If
    txtMoney.Text = Format(mdbl���οۿ��, "0.00")
    dblMoney = mdbl���������ܶ� + mdbl��ˢ�ܶ� - mdbl���οۿ�� - dblMoney
    lblMargin.Caption = "ʣ��δˢ�����:" & Format(dblMoney, "0.00")
    lblMargin.AutoSize = True
    lblMargin.Visible = dblMoney <> 0
    lblBalanceMoney.Caption = Format(mdbl�ʻ����, "0.00")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    msngOldX = 0: msngOldY = 0
    Call InitBalanceGrid '���ˢ������
    If mlngCardTypeID = 0 Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    Call InitFace
    txtPass.PasswordChar = "*"
    Call CreateObjectKeyboard
    Call InitTaskPancel
    Err = 0: On Error Resume Next
    If Not mVarData Is Nothing Then
        If mVarData.Count <> 0 Then
            Call AddWndTaskPancelExpend
            Call ShowWndCaption
        End If
    End If
    
    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    
    Call SetWindowHeight
    mblnFirst = True
    'lbl�ʻ����.Caption = "�ʻ����:"
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With wndTaskPanel
        .Left = ScaleLeft: .Top = ScaleTop
        .Height = ScaleHeight: .Width = ScaleWidth
    End With
End Sub

Private Sub lbl����_Click()
    Dim strExpand As String, strCardNo As String, strOutXml As String
    If Not mcurCardObject.CardPreporty.�Ƿ�Ӵ�ʽ���� Then Exit Sub
'    If mobjICCard Is Nothing Then
'        Set mobjICCard = CreateObject("zlICCard.clsICCard")
'        Set mobjICCard.gcnOracle = gcnOracle
'    End If
    
'    If Not mobjICCard Is Nothing Then
'        txt����.Text = mobjICCard.Read_Card()
'        If txt����.Text <> "" Then
'            mblnICCard = True
'            Call CheckFreeCard(txt����.Text)
'        End If
'    End If
  
    If mcurCardObject.CardObject Is Nothing Then Exit Sub
    If gobjOneCardComLib.objThirdSwap.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    txt����.Text = Trim(strCardNo)
    If txt����.Text <> "" Then
        If Not CheckBrush���ѿ�(Trim(strCardNo)) Then
            Call txt����_GotFocus
            If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Exit Sub
        End If
    End If
    txt����.Tag = Trim(strCardNo)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt����.Text = Trim(strCardNo)
    If txt����.Text <> "" Then
        If Not CheckBrush���ѿ�(txt����.Text) Then
            Call txt����_GotFocus
            If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Exit Sub
        Else
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
        End If
    End If
    txt����.Tag = Trim(strCardNo)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub picBlance_Click()
    Debug.Print "DD"
End Sub

Private Sub picPassWord_Click()
    Debug.Print "DD"
End Sub

Private Sub picPassWord_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    msngOldX = X: msngOldY = Y
End Sub
Private Sub picPassWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button <> 1 Then Exit Sub
        Me.Left = Me.Left + Me.ScaleLeft - msngOldX + X
        Me.Top = Me.Top + Me.ScaleTop - msngOldY + Y
End Sub
 
Private Sub picPassWord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngOldX = 0: msngOldY = 0
End Sub
'--------------------------------------------------------

Private Function SetBrushObject() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-10 13:22:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Errhand
    tmrMain.Tag = "": mblnPassInputCardNo = False
    'һ��ͨ,ֱ���˳�
    If mlngCardTypeID = 0 Then SetBrushObject = True: Exit Function
    If mcurCardObject.CardObject Is Nothing Then
        MsgBox "ע��:" & vbCrLf & "   δ�ҵ���ص������ӿ�,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Not mcurCardObject.InitCompents Then
        If mcurCardObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") Then
              Exit Function
        End If
        mcurCardObject.InitCompents = True
    End If
    mblnPassInputCardNo = mcurCardObject.CardPreporty.�������Ĺ��� <> "" And mcurCardObject.CardPreporty.�������Ĺ��� <> "0"
    Me.Caption = mcurCardObject.CardPreporty.����
    With mcurCardObject.CardPreporty
        Me.txt����.MaxLength = .���ų���
        If .�Ƿ��Զ���ȡ = 1 Then
            tmrMain.Interval = IIf(.�Զ���ȡ��� = 0, 300, .�Զ���ȡ���)
            tmrMain.Tag = 1
        End If
    End With
    '֧��ˢ�������
    '85565,���ϴ�,2015/7/10:��������
    If mcurCardObject.CardPreporty.�Ƿ�Ӵ�ʽ���� Then
        lbl����.BorderStyle = 1
    Else
        lbl����.BorderStyle = 0
    End If
    txt����.Locked = Not (mcurCardObject.CardPreporty.�Ƿ�ˢ�� Or mcurCardObject.CardPreporty.�Ƿ�ɨ��)
    'If cmdRead.Visible = False Then txt����.Width = txtPass.Width
    SetBrushObject = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤���ݵĺϷ���
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-10 14:02:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Trim(txt����.Text) = "" Then
        MsgBox "����δ����,�����뿨�Ż�ˢ��!", vbInformation, gstrSysName
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Function
    End If
    
    If Trim(txt����.Tag) = "" Then
        MsgBox "��δ��֤��Ƭ,���ڿ��Ŵ�����س������鿨!", vbInformation, gstrSysName
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Function
    End If
    
    If Not mbln���ѿ� Then isValied = True: Exit Function
    
    '�����ƿ�,�����ò�������,������֧���ӿ��н�����֤����(һ����˵,���Ƿ�װ�˵�)
    If mcurCardObject.���ƿ� = False Then Exit Function
    
   ' If CheckBrush���ѿ�(Trim(txt����.Tag)) = False Then Exit Function
    
    If zlCommFun.zlStringEncode(txtPass.Text) <> txtPass.Tag Then
        MsgBox "�����������", vbExclamation, gstrSysName
        txtPass.Text = "": mintPassInputCount = mintPassInputCount - 1
        If mintPassInputCount > 2 Then Unload Me: Exit Function
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
        Exit Function
    End If
    If Not mbln�˷� Then
        '����������Ƿ�֧��
        If Round(mdbl�ʻ����, 6) < Round(mdbl���οۿ��, 6) Then
            MsgBox "��ǰ֧�����(" & Format(mdbl���������ܶ�, "0.00") & ")�������ʻ����(" & Format(mdbl�ʻ����, "0.00") & "),���ܼ���!", vbInformation, gstrSysName
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Exit Function
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ClickReadCard()
    If ReadCardNo = False Then
        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Sub
    End If
    If mlngCardTypeID = 0 Then Unload Me: Exit Sub
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    If SetBrushObject = False Then Unload Me: Exit Sub
    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    If lbl����.BorderStyle = 1 Then txt����.ToolTipText = "��F2��س����ж���"
'
'    If cmdRead.Visible = False Then
'        Me.Width = Me.Width - cmdRead.Width * 0.3
'        picPassWord.Width = picPassWord.Width - cmdRead.Width * 0.3
'     End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set mobjICCard = Nothing
    UnHookKBD
    Set mcurCardObject = Nothing
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    Set mobjCommEvents = Nothing
    Set mrsClassMoney = Nothing
End Sub

Private Sub picPassWord_Resize()
    Err = 0: On Error Resume Next
 
End Sub

Private Sub picPati_Resize()
    Err = 0: On Error Resume Next
    lblType.Left = picPati.ScaleWidth - lblType.Width - 20
End Sub
'------------------------------------------------------------------------------------------------------------
Private Sub picPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    msngOldX = X: msngOldY = Y
End Sub
Private Sub picPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button <> 1 Then Exit Sub
        Me.Left = Me.Left + Me.ScaleLeft - msngOldX + X
        Me.Top = Me.Top + Me.ScaleTop - msngOldY + Y
End Sub
 
Private Sub picPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngOldX = 0: msngOldY = 0
End Sub
'------------------------------------------------------------------------------------------------------------

Private Sub tmrMain_Timer()
    If mblnReadCard = False Then
        mblnReadCard = True
        Call ReadCardNo
        mblnReadCard = False
    End If
End Sub

Private Sub txtPass_LostFocus()
    Call ClosePassKeyboard(txtPass)
End Sub
Private Sub txt����_Change()
    txtPass.Enabled = txt����.Text <> ""
    txt����.Tag = "": mstrCardNo = "'"   ' lbl�ʻ����.Caption = "�ʻ����:"
 
    If Not txtPass.Enabled Then txtPass.Text = ""
    tmrMain.Enabled = Val(tmrMain.Tag) <> 0 And Trim(txt����.Text) = ""
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txt����.Text = "")
End Sub
Private Sub txt����_GotFocus()
    Dim strExpend As String
    
    On Error GoTo Errhand
    Call zlControl.TxtSelAll(txt����)
    If Not mobjICCard Is Nothing And mlngCardTypeID = 0 Then Call mobjICCard.SetEnabled(True)
    tmrMain.Enabled = Val(tmrMain.Tag) <> 0 And Trim(txt����.Text) = ""
    
    If mobjSquare Is Nothing Then
        Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
        '��ʼ����Ƶ������
        If Err <> 0 Then Exit Sub
        mobjSquare.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
        If mobjCommEvents Is Nothing Then Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    End If
    If mcurCardObject.CardPreporty.�Ƿ�ǽӴ�ʽ���� Then mobjSquare.SetEnabled True
    '85565:���ϴ�,2015/7/21,����ˢ���ӿ�
    Err = 0: On Error Resume Next
    
    If mcurCardObject.CardPreporty.�ӿ���� = 0 Or mcurCardObject.CardPreporty.�ӿڳ����� = "" Then Exit Sub
    If Not (mcurCardObject.CardPreporty.�Ƿ�ˢ�� Or mcurCardObject.CardPreporty.�Ƿ�ɨ��) Then Exit Sub
    
    Call mobjSquare.zlSetBrushCardObject(mcurCardObject.CardPreporty.�ӿ����, txt����, strExpend, _
                                        mcurCardObject.CardPreporty.���ѿ�)
    If mobjCommEvents Is Nothing Then Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    mobjSquare.zlInitEvents Me.hWnd, mobjCommEvents
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean, lng����ID As Long
    Dim strCardNo As String
    
    If KeyAscii = 13 And Trim(txt����.Text) = "" Then
        If lbl����.BorderStyle = 1 Then
            KeyAscii = 0
            txt����.PasswordChar = IIf(mblnPassInputCardNo, "*", "")
            Call ClickReadCard: Exit Sub
        ElseIf cmdOK.Visible And cmdOK.Enabled Then
            cmdOK.SetFocus: Exit Sub
        End If
    End If
    If txt����.Locked Or txt����.Enabled = False Then Exit Sub
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    blnCard = zlCommFun.InputIsCard(txt����, KeyAscii, mblnPassInputCardNo)
    txt����.PasswordChar = IIf(mblnPassInputCardNo, "*", "")
'
'    If lbl����.BorderStyle = 1 Then
'        'ֻ�ܶ��������ܽ�������
'        If KeyAscii <> 13 Then KeyAscii = 0: Exit Sub
'    End If
    
    If Not (blnCard And Len(txt����.Text) = txt����.MaxLength - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txt����.Text) <> "") Then
        '����:51570
        '����ˢ���ͻس�,���˳�
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        '��ȫˢ�����
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = timer
            If txt����.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txt����.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txt����.Text = Chr(KeyAscii)
                txt����.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
        Exit Sub
    End If
    
    If KeyAscii <> 13 Then
        txt����.Text = txt����.Text & Chr(KeyAscii)
        txt����.SelStart = Len(txt����.Text)
    End If
    KeyAscii = 0
    strCardNo = Trim(txt����.Text)
    '68927,������,2014-01-17,ˢ����ˢ��ĩβ���ܴ����лس��������
    EnableKBDHook
    If CheckBrush���ѿ�(strCardNo) = False Then
        txt����.Text = ""
        Exit Sub
    End If
    txt����.Text = strCardNo ' GetCardNODencode(strCardNo, mlngCardTypeID, mcurCardObject.CardPreporty.�������Ĺ���, mbln���ѿ�)
    txt����.Tag = strCardNo
    '����ˢ����,������ȵ����,��һλ���лس���,��Ҫȡ���ûس���,����ת����������ո��ַ�
    '���Լ�����Doevnts:63335
    DoEvents
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Public Function CheckBrush���ѿ�(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ���ѿ�
    '����:���˺�
    '����:2011-06-23 17:48:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strPassWord As String, dblʧЧ��� As Double
    Dim dblʵ��֧�� As Double, str������� As String, lng���ѿ�ID As Long
    Dim dbl��ˢ�ܶ� As Double, dbl��ǰ֧�� As Double
    
    If Not mbln���ѿ� Then CheckBrush���ѿ� = True: Exit Function
    dbl��ˢ�ܶ� = GetBruhMoney  '��ȡ�Ѿ�ˢ�����ܶ�
    dbl��ǰ֧�� = mdbl���������ܶ� + mdbl��ˢ�ܶ� - dbl��ˢ�ܶ�
    If CheckBrushSquareCard(mlngCardTypeID, strCardNo, mrsClassMoney, _
        dbl��ǰ֧��, strPassWord, mdbl�ʻ����, dblʧЧ���, dblʵ��֧��, _
        mbln�����ֹ, str�������, mlng���ѿ�ID, dbl��ˢ�ܶ�) = False Then Exit Function
    '������ˢ������
    mstr������� = str�������
    mdbl���οۿ�� = dblʵ��֧��:    txtPass.Tag = strPassWord
    Call ShowMoney
    CheckBrush���ѿ� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetBruhMoney() As Double
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�Ѿ�ˢ���Ľ��
    '����:�����Ѿ�ˢ���Ľ��
    '����:���˺�
    '����:2013-02-22 16:09:48
    '����:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    On Error GoTo errHandle
    '��ʾ��ˢ���
    With vsBlance
        dblMoney = 0
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, .ColIndex("����")) <> "" Then
                dblMoney = dblMoney + Val(.TextMatrix(i, .ColIndex("ˢ�����")))
            End If
        Next
    End With
    GetBruhMoney = dblMoney
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitBalanceGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '����:���˺�
    '����:2013-02-22 16:18:49
    '����:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
   On Error GoTo errHandle
    With vsBlance
        .Clear
        .Rows = 2: .Cols = 7
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "�����ID"
        .TextMatrix(0, 3) = "���ѿ�ID"
        .TextMatrix(0, 4) = "�������"
        .TextMatrix(0, 5) = "ˢ�����"
        .TextMatrix(0, 6) = "����������ʾ"
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            Select Case .ColKey(i)
            Case "�����ID", "���ѿ�ID", "����", "�������", "����������ʾ"
                .ColHidden(i) = True
            Case "ˢ�����"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
Private Function CheckTypeValied(ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
    ByVal str������� As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������շ����ĺϷ���
    '���:lngCardTypeID-�����ID
    '        strCardNo-����
    '       str�������-���Ƶ��շ����
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2013-02-25 11:07:06
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strName As String, j As Long
    Dim strTemp As String, varData As Variant

    On Error GoTo errHandle
    If mcurCardObject Is Nothing Then
        strName = "���ѿ�"
    Else
        strName = mcurCardObject.CardPreporty.����
    End If
    
    With vsBlance
        '�ȼ���Ƿ�Ϸ�,�Ϸ��ż���
        For i = 1 To .Rows - 1
            If Trim(.Cell(flexcpData, i, .ColIndex("����"))) <> "" Then
                If Trim(.Cell(flexcpData, i, .ColIndex("����"))) = strCardNo _
                    And Val(.TextMatrix(i, .ColIndex("�����ID"))) = lngCardTypeID Then
                    '���ſ��Ѿ�ˢ��,�����ٽ���ˢ����֤
                    MsgBox "����Ϊ" & strCardNo & " ��" & strName & ",�����Ѿ�ˢ������,������ˢ��!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                strTemp = Trim(.TextMatrix(i, .ColIndex("�������")))
                
                If (strTemp <> "" Or Trim(str�������) <> "") And strTemp <> str������� Then
                    '�����������Ƿ���ͬ
                    If strTemp <> "" Then
                        varData = Split(strTemp, ",")
                        For j = 0 To UBound(varData)
                            If Trim(varData(j)) <> "" Then
                                '��ǰ�շ����,�������������ʱ,�����
                                If InStr(1, "," & mstrCurFeeType & ",", "," & varData(j) & ",") > 0 Or mstrCurFeeType = "" Then
                                        If InStr(1, "," & str������� & ",", "," & varData(j) & ",") = 0 Then
                                            MsgBox "����Ϊ" & strCardNo & " ��" & strName & "������������Ѿ�ˢ�����������ͬ,���ܻ��ˢ��,��������:" & vbCrLf & "  ��ˢ���������:" & strTemp & vbCrLf & "  ��ǰˢ���������:" & str�������, vbInformation + vbOKOnly, gstrSysName
                                            Exit Function
                                        End If
                                End If
                            End If
                        Next
                    End If
                    If str������� <> "" Then
                        varData = Split(str�������, ",")
                        For j = 0 To UBound(varData)
                            If Trim(varData(j)) <> "" Then
                                '��ǰ�շ����,�������������ʱ,�����
                                If InStr(1, "," & mstrCurFeeType & ",", "," & varData(j) & ",") > 0 Or mstrCurFeeType = "" Then
                                    If InStr(1, "," & strTemp & ",", "," & varData(j) & ",") = 0 Then
                                        MsgBox "����Ϊ" & strCardNo & " ��" & strName & "������������Ѿ�ˢ�����������ͬ,���ܻ��ˢ��,��������:" & vbCrLf & "  ��ˢ���������:" & strTemp & vbCrLf & "  ��ǰˢ���������:" & str�������, vbInformation + vbOKOnly, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Next
    End With
    CheckTypeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function AddBrushCardInfor(ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, ByVal strPassWord As String, _
    ByVal str������� As String, ByVal lng���ѿ�ID As Long, _
    ByVal dblMoney As Double) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ˢ����Ϣ
    '���:lngCardTypeID-��ǰ�����ID
    '       strCardNo-��ǰ����
    '       strPassWord-��ǰ����
    '       str�������-��ǰ�����������
    '       lng���ѿ�ID-��ǰ���ѿ��Ŀ�ID
    '       dblMoney-����ˢ�����
    '����:����ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-02-22 16:16:01
    '����:57682
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strName As String
    Dim strTemp As String, varData As Variant, lngRow As Long
    On Error GoTo errHandle
    If mcurCardObject Is Nothing Then
        strName = "���ѿ�"
    Else
        strName = mcurCardObject.CardPreporty.����
    End If
    With vsBlance
        '�ȼ���Ƿ�Ϸ�,�Ϸ��ż���
        lngRow = 0
        For i = 1 To .Rows - 1
            If Trim(.Cell(flexcpData, i, .ColIndex("����"))) = "" Then
                    lngRow = i: Exit For
            End If
        Next
        If lngRow > .Rows - 1 Or lngRow = 0 Then
            .Rows = .Rows + 1
            lngRow = .Rows - 1
        End If
        .TextMatrix(lngRow, .ColIndex("����")) = IIf(mblnPassInputCardNo, String(Len(strCardNo), "*"), strCardNo)
        .Cell(flexcpData, lngRow, .ColIndex("����")) = strCardNo
        .TextMatrix(lngRow, .ColIndex("����")) = strPassWord
        .TextMatrix(lngRow, .ColIndex("�����ID")) = lngCardTypeID
        .TextMatrix(lngRow, .ColIndex("���ѿ�ID")) = lng���ѿ�ID
        .TextMatrix(lngRow, .ColIndex("ˢ�����")) = Format(dblMoney, "0.00")
        .TextMatrix(lngRow, .ColIndex("����������ʾ")) = IIf(mblnPassInputCardNo, 1, 0)
        .TextMatrix(lngRow, .ColIndex("�������")) = str�������
        .RowData(lngRow) = 0: .RowPosition(lngRow) = 1
        '77292,Ƚ����,2014-8-29,���ѿ��˿���֤ʱ,����������֣����˿��δȫ����֤������ͨ��
        If mbln���ѿ� And mbln�˷� And mbln���� = False And mdbl���οۿ�� <> 0 Then
            cmdOK.Enabled = False
        Else
            cmdOK.Enabled = True
        End If
    End With
    '��ʾ�ϼƽ��
    AddBrushCardInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ShowWndCaption()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ˢ���
    '����:���˺�
    '����:2013-02-22 17:09:56
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    On Error GoTo errHandle
    Set tkpGroup = wndTaskPanel.Groups.Find(2)
    '���ڣ����˳�
    If Not tkpGroup Is Nothing Then
        tkpGroup.Caption = "��ǰˢ����Ϣ(��ˢ:" & Format(GetBruhMoney, "0.00") & ")"
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt����.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt����.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPass.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPass_GotFocus()
    '89759:���ϴ�,2017/4/14,ˢ�����Ƿ����ȷ��ˢ����Ϣ
    If Not mblnPosPass Then
        If txtPass.Tag = "" And txt����.Text <> "" Then Call txtPass_KeyPress(13): Exit Sub
    End If
    EnableKBDHook
    Call zlControl.TxtSelAll(txtPass)
    Call OpenPassKeyboard(txtPass)
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlSquareAffirm
    ElseIf KeyAscii = 22 Then
        KeyAscii = 0 '������ճ��
    End If
End Sub
Private Function zlSquareAffirm(Optional blnOK As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȷ��
    '���:blnOk-����ɹ���ʱ����
    '����:���غϷ�����true,���򷵻�False
    '����:���˺�
    '����:2013-02-26 17:34:02
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblʣ��� As Double
    If Not isValied Then Exit Function
    mstrCardNo = txt����.Tag: mstrPassWord = Trim(txtPass.Text)
    dblʣ��� = mdbl���������ܶ� + mdbl��ˢ�ܶ� - Val(txtMoney.Text) - GetBruhMoney
    Call AddBrushCardInfor(mlngCardTypeID, mstrCardNo, mstrPassWord, mstr�������, mlng���ѿ�ID, Val(txtMoney.Text))
    
    If Round(dblʣ���, 6) = 0 Or mblnתԤ�� Then
        'ˢ�����,��������ˢ��
        mbln���� = False
        If chk����.Visible Then mbln���� = chk����.value = 1
        mblnOk = True
        zlSquareAffirm = True
        If blnOK Then Exit Function
        Call SetReturnBrushCardInfor
        Unload Me
        Exit Function
    End If
    
    Call AddWndTaskPancelExpend
    mdbl���οۿ�� = mdbl���������ܶ� + mdbl��ˢ�ܶ� - GetBruhMoney
    Call ShowMoney
    Call ShowWndCaption
    txtPass.Text = "": txt����.Text = ""
    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
    zlSquareAffirm = True
End Function
Private Function ReadCardNo() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����
    '����:�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2011-06-22 14:40:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strOutPatiXML As String, strCardNo As String
    On Error GoTo errHandle
    If mlngCardTypeID = 0 Then
        If mobjICCard Is Nothing Then Exit Function
         txt����.Text = mobjICCard.Read_Card()
         txt����.Tag = txt����.Text
         mstrCardNo = txt����.Text
         Exit Function
    End If
    
    ' ��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    'frmMain Object  In  ���õ�������
    'lngModule   Long    In  ���õ�ģ���
    'blnOlnyCardNO   boolean In  ������ȡ����
    'strExpand   String  In  ��չ����,��������Ϊ��
    'strOutCardNO    String  Out ����
    'strOutclsPatientInfoXml  XML Out ��strOutclsPatientInfoXml����˵��
    '����: blnOlnyCardNO=trueʱ,���ؿ�
    If gobjOneCardComLib.objThirdSwap.zlReadCard(Me, mlngModule, True, "", strCardNo, strOutPatiXML) = True Then
         txt����.Text = strCardNo ' GetCardNODencode(strCardNo, mlngCardTypeID, mcurCardObject.CardPreporty.�������Ĺ���, mbln���ѿ�)
         txt����.Tag = strCardNo
    End If
    
    'txt����.PasswordChar = ""
    '91140:���ϴ�,2015/11/30,���ѿ�ˢ��
    If mbln���ѿ� Then
        If CheckBrush���ѿ�(strCardNo) Then ReadCardNo = True
        Exit Function
    End If
    
    'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long,
       ' strCardTypeID as long ,ByVal strCardNo As String, strExpand As String, dblMoney As Double) As Boolean
    If mcurCardObject.CardObject.zlGetAccountMoney(Me, mlngModule, mlngCardTypeID, strCardNo, "", mdbl�ʻ����) Then
         lblBalanceMoney.Caption = Format(mdbl�ʻ����, "0.00")
          If mdbl�ʻ���� - mdbl���������ܶ� < 0 Then
                If mdbl�ʻ���� = 0 Then
                    Call MsgBox("�ÿ��Ѿ�û�п������,���ܼ���!", vbInformation + vbOKOnly, gstrSysName)
                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                    Exit Function
                Else
                    If mbln�����ֹ Then
                        Call MsgBox("���ʻ�����֧���������Ѷ�,���ܼ���!", vbInformation + vbOKOnly, gstrSysName)
                        If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                        Exit Function
                    End If
                    If MsgBox("���ʻ�����֧���������Ѷ�,�Ƿ����ʻ������Ϊ�������Ѷ�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                          If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                          Exit Function
                    End If
                    mdbl���οۿ�� = Round(mdbl�ʻ����, 2)
                End If
          End If
    End If
    
    Call ShowMoney
    ReadCardNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    Dim lngPreIDKind As Long
    If Me.ActiveControl Is txt���� Then
        txt����.Text = strNo
        If txt����.Text = "" Then Call mobjICCard.SetEnabled(False)
        mstrCardNo = strNo
        Unload Me: Exit Sub
    End If
End Sub

Public Function CheckBrushSquareCard(ByVal lngCardTypeID As Long, _
    ByVal strCardNo As String, _
    ByVal rsClassMoney As ADODB.Recordset, _
    ByVal dbl����֧���� As Double, _
    ByRef strPassWord As String, _
    ByRef dbl�ʻ���� As Double, _
    ByRef dblʧЧ��� As Double, ByRef dblʵ��֧�� As Double, _
    Optional bln�����ֹ As Boolean = True, _
    Optional ByRef str�������Out As String = "", _
    Optional ByRef lng���ѿ�ID As Long = 0, _
    Optional ByRef dbl��ˢ�ܶ� As Double) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ���ѿ�
    '���:lngCardTypeID-�����ID
    '       dbl��ˢ�ܶ�-�Ѿ�ˢ�����ܶ�
    '����: strPassWord-���ؽ��ܵ�����
    '         dbl�ʻ����-�ʻ����
    '         dblʧЧ���-ʧЧ���
    '         str�������Out-���Ƶ�ʹ�����
    '        dblʵ��֧��-ʵ��֧�����
    '����:ˢ���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-02-25 10:57:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, blnFind As Boolean
    Dim dbl��� As Double, str���� As String
    Dim strSQL As String, dbl�޶��� As Double, dbl�ϼ� As Double, strMsg As String
    Dim intIndex As Integer, rs�շ���� As ADODB.Recordset, str������� As String
    Dim varData As Variant, dblMoney As Double
    Dim strҵ�񳡺� As String, var������Դ As Variant
    
    dblʵ��֧�� = dbl����֧����
    
    '��|ȫ��|������־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If lngCardTypeID < 0 Then Exit Function
    
    If str���� = "" Then
        Set rsTemp = zlGet���ѿ��ӿ�
        rsTemp.Filter = "���=" & lngCardTypeID
        If rsTemp.EOF Then
            MsgBox "δ�ҵ���صĿ�����ӿ�", vbInformation, gstrSysName
            Exit Function
        End If
        str���� = NVL(rsTemp!����)
        rsTemp.Filter = 0
    End If
    
    strSQL = _
        "Select a.Id,a.������,a.����,a.���,a.�ɷ��ֵ,to_char(a.��Ч��,'yyyy-mm-dd hh24:mi:ss') as ��Ч��," & vbNewLine & _
        "       to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��," & vbNewLine & _
        "       decode(a.��ǰ״̬,2,'����',3,'�˿�','����') as ��ǰ״̬," & vbNewLine & _
        "       to_char(a.������," & gOraFmtString.FM_��� & ") as ������," & vbNewLine & _
        "       to_char(a.���۽��," & gOraFmtString.FM_��� & ") as ���۽��," & vbNewLine & _
        "       to_char(a.��ֵ�ۿ���," & gOraFmtString.FM_�ۿ��� & ") as ��ֵ�ۿ���," & vbNewLine & _
        "       to_char(a.���," & gOraFmtString.FM_��� & ") as ���," & vbNewLine & _
        "       to_char(a.ͣ������,'yyyy-mm-dd hh24:mi:ss') as ͣ������," & vbNewLine & _
        "       a.������� ,A.����, b.Ӧ�ó���, b.�Ƿ��ض�����, a.����ID" & vbNewLine & _
        "From ���ѿ���Ϣ A, ���ѿ����Ŀ¼ B" & vbNewLine & _
        "Where a.�ӿڱ�� = b.��� And A.���� = [1] and A.�ӿڱ��=[2]" & vbNewLine & _
        "      And ��� = (Select Max(���) From ���ѿ���Ϣ B Where ���� = A.���� and �ӿڱ��=A.�ӿڱ��)" & vbNewLine & _
        "Order by a.���"
    Err = 0: On Error GoTo Errhand:
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ѿ����", strCardNo, lngCardTypeID)
    If rsTemp.EOF Then
       ShowMsgbox "δ�ҵ���ص�" & str���� & "��Ϣ�����飡"
        Exit Function
    End If
    
    '��鵱ǰˢ���ĺϷ���
    '�Ƿ����
    If NVL(rsTemp!����ʱ��, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & str���� & "�Ѿ���" & NVL(rsTemp!��ǰ״̬) & "��������ˢ����"
        Exit Function
    End If
    
    '�Ƿ�ͣ��
    If NVL(rsTemp!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & str���� & "�Ѿ���ֹͣʹ�ã�������ˢ����"
        Exit Function
    End If
    
    If NVL(rsTemp!����) <> "" Then
        strPassWord = NVL(rsTemp!����)
    End If
    
    lng���ѿ�ID = Val(NVL(rsTemp!id))
    str�������Out = NVL(rsTemp!�������)
    dbl��� = Val(NVL(rsTemp!���))
    
    dblʧЧ��� = 0
    '���Ч��
    If NVL(rsTemp!��Ч��, "3000-01-01 00:00:00") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
        '������Ч��
        If Val(NVL(rsTemp!�ɷ��ֵ)) = 1 Then
            '������ֵ��,���ڵ�,�������ѿ�����,ֻ��������ֵ����
            dblʧЧ��� = zlGetʧЧ���(Val(NVL(rsTemp!id)))
            dbl��� = dbl��� - dblʧЧ���
            If dbl��� <= 0 Then dbl��� = 0
        ElseIf mbln�˷� = False Then
            '��������ֵ��,�����ٽ�������
            ShowMsgbox "����Ϊ" & strCardNo & "��" & str���� & "�Ѿ�ʧЧ��������ˢ����"
            Exit Function
        End If
    End If
    
    If mbln�˷� Then  '�˷�
        If Not mVarData Is Nothing Then
            dblMoney = 0
            blnFind = False
            For i = 1 To mVarData.Count
                'arrayarray(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����,ʣ��δ�˽��(ֻ����˷�))
                varData = mVarData(i)
                If Val(varData(0)) = lngCardTypeID And lng���ѿ�ID = varData(1) Then
                    blnFind = True
                    If UBound(varData) >= 7 Then
                        dblMoney = dblMoney + Val(varData(7))
                    Else
                        dblMoney = dblMoney + Val(varData(2))
                    End If
                End If
            Next
            If bln�����ֹ Then
                If Round(dblMoney, 6) < dbl����֧���� Then
                    ShowMsgbox "����Ϊ " & strCardNo & " ��" & str���� & "�˿������ʣ��δ�˽������˷ѣ�"
                    Exit Function
                End If
            Else
                If Round(dblMoney, 6) < dbl����֧���� Then dblʵ��֧�� = dblMoney
            End If
            '77292,Ƚ����,2014-8-29,��ǰ�������շ�ʱʹ�õĿ��б��У���ÿ���֤ʧ��
            If Not blnFind Then
                ShowMsgbox "����Ϊ " & strCardNo & " ��" & str���� & "���շ�ʱδʹ�ã������˷ѣ�"
                Exit Function
            End If
            '77292,Ƚ����,2014-8-29,ˢ���ʱ,�����ظ�ʹ��ͬһ�ſ�
            If Not CheckTypeValied(lngCardTypeID, strCardNo, "") Then Exit Function
            '78494,Ƚ����,2014-10-10,��ǰ��δ�˽��Ϊ��,����������
            If Round(dblMoney, 6) = 0 Then
                ShowMsgbox "����Ϊ " & strCardNo & " ��" & str���� & "δ�˽��Ϊ�㣬�������˷ѣ�"
                Exit Function
            End If
        End If
        GoTo EndNO:
    End If
    

    If mblnתԤ�� And NVL(rsTemp!�������) <> "" Then
        ShowMsgbox str���� & "�������Ƶ��շ���𣬲�����תԤ����"
        Exit Function
    End If
    
    strҵ�񳡺� = NVL(rsTemp!Ӧ�ó���) & "000"
    'Ӧ�ó��ϼ��
    '����λ������ɣ�ÿһλ1��ʾ����ʹ�ã�0��ʾ����ʹ�ã���һλ����ҵ�񣬵ڶ�λסԺҵ�񣬵���λ���ҵ��ȱʡΪ''111''
    var������Դ = Split(mstr������Դ, ",")
    For i = 0 To UBound(var������Դ)
        If InStr("123", Val(var������Դ(i))) > 0 Then
            If Val(Mid(strҵ�񳡺�, Val(var������Դ(i)), 1)) = 1 Then
                ShowMsgbox "����Ϊ" & strCardNo & "��" & str���� & _
                    "���" & Decode(Val(var������Դ(i)), 2, "סԺ", 3, "���", "����") & "���ò�����ʹ�ã�"
                Exit Function
            End If
        Else
            ShowMsgbox "����Ϊ" & strCardNo & "��" & str���� & "�ڵ�ǰҵ�񳡺ϲ�����ʹ�ã�"
            Exit Function
        End If
    Next
    If Val(NVL(rsTemp!�Ƿ��ض�����)) = 1 Then
        If mlng����ID <> Val(NVL(rsTemp!����ID)) Then
            ShowMsgbox "����Ϊ" & strCardNo & "��" & str���� & "ֻ������֧���ֿ��˱��˵ķ��ã�"
            Exit Function
        End If
    End If

    If dbl��� <= 0 Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & str���� & "�Ѿ�û����������ˢ�����ѣ�"
        Exit Function
    End If
    If dbl��� < dbl����֧���� Then
        If bln�����ֹ Then
            ShowMsgbox "" & str���� & "�����(" & Format(dbl���, "0.00") & ")����֧�����ν��(" & Format(dbl����֧����, "0.00") & ")��"
            Exit Function
        End If
        '�������Ϊ����֧����
        dblʵ��֧�� = dbl���
    End If
    'ˢ���ʱ,���ܴ��ڲ�ͬ����ˢ�����
    If Not CheckTypeValied(lngCardTypeID, strCardNo, NVL(rsTemp!�������)) Then Exit Function

    '�������ˢ����
    Set rs�շ���� = zlGet�շ����
    str������� = zlGet��ȡ�������FromNameToCode(NVL(rsTemp!�������))
    
    If rsClassMoney Is Nothing Then GoTo EndNO:
    If rsClassMoney.State <> 1 Then GoTo EndNO:
    
    With rsClassMoney
        dbl�ϼ� = 0: dbl�޶��� = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dbl�ϼ� = dbl�ϼ� + Val(NVL(!���))
            If InStr(1, str�������, "," & !�շ���� & ",") > 0 Then
                rs�շ����.Filter = "����='" & NVL(!�շ����) & "'"
                If Not rsTemp.EOF Then
                    strMsg = strMsg & vbCrLf & "" & rs�շ����!���� & ":" & Val(NVL(!���))
                End If
                dbl�޶��� = dbl�޶��� + Val(NVL(!���))
            End If
            .MoveNext
        Loop
        dbl�޶��� = Format(dbl�޶���, "0.00")
        dbl�ϼ� = Format(dbl�ϼ�, "0.00")
        If dbl�ϼ� - dbl�޶��� - dbl��ˢ�ܶ� >= dblʵ��֧�� Then GoTo EndNO:
        If dbl�ϼ� - dbl�޶��� - dbl��ˢ�ܶ� <= 0 Then
            If dbl�޶��� <> 0 Then
                ShowMsgbox "����Ϊ" & strCardNo & "��" & str���� & " " & vbCrLf & "���ڽ�����ƣ����β���֧��,�����������:" & vbCrLf & strMsg
            Else
                ShowMsgbox "�Ѿ�ˢ���������,������ˢ������!"
            End If
            Exit Function
        End If
        dblʵ��֧�� = dbl�ϼ� - dbl�޶��� - dbl��ˢ�ܶ�
    End With
EndNO:
    dbl�ʻ���� = dbl���
    CheckBrushSquareCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get�շ��������_����() As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ��������,������Ϊ��
    '����:�շ��������,������Ϊ��
    '����:���˺�
    '����:2013-03-07 10:30:39
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If mblnתԤ�� Then Exit Function
    If mrsClassMoney Is Nothing Then Exit Function
    If mrsClassMoney.State <> 1 Then Exit Function
    If mrsClassMoney.RecordCount = 0 Then Exit Function
    Set rsTemp = zlGet�շ����
    rsTemp.Filter = 0
    mrsClassMoney.Filter = 0
    With mrsClassMoney
        .MoveFirst
        Do While Not .EOF
            rsTemp.Find "����='" & NVL(!�շ����, "-") & "'", , adSearchForward, 1
            If Not rsTemp.EOF Then strTemp = strTemp & "," & rsTemp!����
            .MoveNext
        Loop
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    mrsClassMoney.Filter = 0
    Get�շ��������_���� = strTemp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 

End Function


