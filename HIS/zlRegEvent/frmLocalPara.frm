VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLocalPara 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   ControlBox      =   0   'False
   Icon            =   "frmLocalPara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin TabDlg.SSTab tbPage 
      Height          =   5385
      Left            =   45
      TabIndex        =   23
      Top             =   105
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9499
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frmLocalPara.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstDept"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDefaultSet"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkRigistHeadSort"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDeviceSetup"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPrintSet(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPrintSet(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPrintSet(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdPrintSet(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "����Ʊ��"
      TabPicture(1)   =   "frmLocalPara.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCards"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraTitle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "ԤԼ�Һŵ���ӡ����"
         Height          =   345
         Index           =   2
         Left            =   2340
         TabIndex        =   14
         Top             =   3840
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "�Һ�ƾ����ӡ����"
         Height          =   345
         Index           =   3
         Left            =   2340
         TabIndex        =   12
         Top             =   3405
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "���������ӡ����"
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   3405
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "�Һ�Ʊ�ݴ�ӡ����"
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   3840
         Width           =   1875
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   330
         Left            =   240
         TabIndex        =   15
         Top             =   4275
         Width           =   1425
      End
      Begin VB.CheckBox chkRigistHeadSort 
         Caption         =   "�ҺŰ��ű�����ͷ����"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2325
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ùҺ�Ʊ��"
         Height          =   1845
         Left            =   -74835
         TabIndex        =   19
         Top             =   525
         Width           =   6675
         Begin MSComctlLib.ListView lvwBill 
            Height          =   1455
            Left            =   150
            TabIndex        =   20
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483630
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "������"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "��������"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "���뷶Χ"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "ʣ��"
               Object.Width           =   1499
            EndProperty
         End
      End
      Begin VB.Frame fraDefaultSet 
         Caption         =   "ȱʡֵ"
         Height          =   2445
         Left            =   240
         TabIndex        =   25
         Top             =   825
         Width           =   4290
         Begin VB.ComboBox cboType 
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1590
            Width           =   2865
         End
         Begin VB.ComboBox cboDefaultSex 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1995
            Width           =   1260
         End
         Begin VB.ComboBox cboDefaultPayMode 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   405
            Width           =   2865
         End
         Begin VB.ComboBox cboDefaultFeeType 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   780
            Width           =   2865
         End
         Begin VB.ComboBox cboDefaultBalance 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1185
            Width           =   2865
         End
         Begin VB.Label lblDefaultPayCard 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   420
            TabIndex        =   7
            Top             =   1650
            Width           =   720
         End
         Begin VB.Label lblDefaultSex 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            Height          =   180
            Left            =   780
            TabIndex        =   9
            Top             =   2055
            Width           =   360
         End
         Begin VB.Label lblDefaultPayMode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ʽ"
            Height          =   180
            Left            =   420
            TabIndex        =   1
            Top             =   465
            Width           =   720
         End
         Begin VB.Label lblDefaultBalance 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���㷽ʽ"
            Height          =   180
            Left            =   420
            TabIndex        =   5
            Top             =   1245
            Width           =   720
         End
         Begin VB.Label lblDefaultFeeType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ѱ�"
            Height          =   180
            Left            =   780
            TabIndex        =   3
            Top             =   840
            Width           =   360
         End
      End
      Begin VB.ListBox lstDept 
         ForeColor       =   &H80000012&
         Height          =   4470
         Left            =   4785
         Style           =   1  'Checkbox
         TabIndex        =   17
         ToolTipText     =   "Ctrl+Aȫѡ,Ctrl+Cȫ��,���һ����δѡ���ʾ�����ƿ���"
         Top             =   690
         Width           =   2175
      End
      Begin VB.Frame fraCards 
         Caption         =   "���ع���ҽ�ƿ�"
         Height          =   2655
         Left            =   -74805
         TabIndex        =   21
         Top             =   2550
         Width           =   6660
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2190
            Left            =   195
            TabIndex        =   22
            Top             =   300
            Width           =   6405
            _cx             =   11298
            _cy             =   3863
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmLocalPara.frx":0044
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
            ExplorerBar     =   2
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һſ���"
         Height          =   180
         Left            =   4770
         TabIndex        =   16
         ToolTipText     =   "�趨�����ɹ���Щ���ҵĺ�"
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   330
      Left            =   7260
      TabIndex        =   18
      Top             =   5100
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   330
      Left            =   7260
      TabIndex        =   26
      Top             =   885
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   330
      Left            =   7260
      TabIndex        =   24
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "frmLocalPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrPrivs As String
Public mlngModul As Long
Private Sub chkDeptBespeakOneNum_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboDefaultBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboDefaultFeeType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboDefaultPayMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
 
Private Sub cboDefaultSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
 
Private Sub chkRigistHeadSort_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1111)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strTmp As String
    Dim blnHavePrivs As Boolean
    
    On Error GoTo Hd
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '���ݿ�洢��ģ�����
    '-------------------------------------------------------------------------------------------
    zlDatabase.SetPara "ȱʡ���ʽ", cboDefaultPayMode.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ�ѱ�", cboDefaultFeeType.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ���㷽ʽ", cboDefaultBalance.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ�Ա�", cboDefaultSex.Text, glngSys, mlngModul, blnHavePrivs
     '���� 43847
    zlDatabase.SetPara "������ͷ����", chkRigistHeadSort.Value, glngSys, mlngModul, blnHavePrivs
    
    strTmp = ""
    If lstDept.ListCount <> lstDept.SelCount Then
        For i = 0 To lstDept.ListCount - 1
            If lstDept.Selected(i) = True Then
                strTmp = strTmp & "," & lstDept.ItemData(i)
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End If
    zlDatabase.SetPara "�Һſ���", strTmp, glngSys, mlngModul, blnHavePrivs
    
    '���ùҺ�Ʊ������
    strTmp = "0"
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then strTmp = Mid(lvwBill.ListItems(i).Key, 2)
    Next
    zlDatabase.SetPara "���ùҺ�Ʊ������", strTmp, glngSys, mlngModul, blnHavePrivs
    
    Call SaveInvoice
    Call InitLocPar(mlngModul)
    gblnOk = True
    Unload Me
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Function LoadFactList(bytKind As Byte) As Boolean
'���ܣ���ȡ���ù��ùҺ�Ʊ�ݻ���￨����
'����:bytKind=4-�Һ�Ʊ��,5-���￨
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngTmp As Long
    Dim ObjItem As ListItem
    Dim blnBill As Boolean
    
    On Error GoTo errH
    lngTmp = zlDatabase.GetPara("���ùҺ�Ʊ������", glngSys, mlngModul, 0, Array(lvwBill), InStr(mstrPrivs, "��������") > 0)
    Set rsTmp = GetShareInvoiceGroupID(bytKind)
    
    For i = 1 To rsTmp.RecordCount
        Set ObjItem = lvwBill.ListItems.Add(, "_" & rsTmp!id, rsTmp!������)
        ObjItem.SubItems(1) = Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd")
        ObjItem.SubItems(2) = rsTmp!��ʼ���� & "," & rsTmp!��ֹ����
        ObjItem.SubItems(3) = rsTmp!ʣ������
        If rsTmp!id = lngTmp Then
            ObjItem.Checked = True
            ObjItem.Selected = True
            blnBill = True
        End If
        rsTmp.MoveNext
    Next
    
    If Not blnBill Then
        zlDatabase.SetPara IIf(bytKind = 4, "���ùҺ�Ʊ������", "���þ��￨����"), "0", glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    
    LoadFactList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdPrintSet_Click(index As Integer)
    On Error GoTo Hd
    Select Case index
    '���������ӡ
    Case 0:
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_2", Me)
    Case 1:
        '�Һ��շѴ�ӡ
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 2:
        'ԤԼ�ҺŴ�ӡ   '56274
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    Case 3:
        '68408,������,2013-12-11,�Һ�ƾ����ӡ
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me)
    Case Else:
    End Select
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Dim i As Integer
        If UCase(Chr(KeyCode)) = "A" Then
            For i = 0 To lstDept.ListCount - 1
                lstDept.Selected(i) = True
            Next
        ElseIf UCase(Chr(KeyCode)) = "C" Then
            For i = 0 To lstDept.ListCount - 1
                lstDept.Selected(i) = False
            Next
        End If
    End If
End Sub

Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:Ƚ����
    '����:2014-07-02
    '�����:74552
    '˵��:�ҺŹ���������Ĭ�Ͻ��㷽ʽʱ�����ѡ����㷽ʽ����Ϊ"7-һ��ͨ����"�Ľ��㷽ʽ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String

    strSQL = _
        " Select B.����,B.����,Nvl(B.ȱʡ��־,0) as ȱʡ,Nvl(B.����,1) as ����,Nvl(B.Ӧ����,0) as Ӧ����" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ" & _
        "   And(B.����<>7 Or B.����=7 And Exists(Select 1 From һ��ͨĿ¼ C Where C.���㷽ʽ=B.���� And C.����=1))" & _
        "   and B.����<>8 And Instr(',1,2,7,',','||B.����||',')>0" & _
        " Order by ����,lpad(����,3,' ')"
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "�Һ�")
    
    '��ȡ�������Ľ��㷽ʽ
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    
    varData = Split(strPayType, ";")
    With cboDefaultBalance
        .Clear
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then
                    blnFind = True: Exit For
                End If
            Next
                         
            If Not blnFind Then
                .AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
                .ItemData(.NewIndex) = 1
                If Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����) = gstr���㷽ʽ Then
                     .ItemData(.NewIndex) = 1
                     .ListIndex = .NewIndex
                End If
                If Val(Nvl(rsTemp!ȱʡ)) = 1 Then .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        
        '���ؽ��㷽ʽ����Ϊ��7-һ��ͨ���㡱��ҽ�ƿ����
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                If varTemp(1) = gstr���㷽ʽ Then
                     .ItemData(.NewIndex) = 1
                     .ListIndex = .NewIndex
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp           As New ADODB.Recordset
    Dim strSQL          As String
    Dim i               As Integer
    Dim str����ID       As String
    Dim strTmp          As String
    Dim blnParSet       As Boolean

    
    gblnOk = False
    
    blnParSet = InStr(mstrPrivs, "��������") > 0
    On Error GoTo errH
    'a.��ʼ����
    '----------------------------------------------------------------------------------------
    strSQL = "Select Distinct B.���� ||'-'|| B.���� as ����,B.ID From �ҺŰ��� A,���ű� B Where A.����ID=B.ID Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    zlcontrol.CboAddData lstDept, rsTmp, True
    
    strSQL = "Select 'ҽ�Ƹ��ʽ' ����,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ�Ƹ��ʽ" & _
            " Union All " & _
            " Select '�Ա�' ����,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա�" & _
            " Union All " & _
            " Select '�ѱ�' ����,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ�" & _
            " Where ����=1 And Nvl(���޳���,0)=0 And Nvl(�������,3) IN(1,3)" & _
            " Order by ����,����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    'ȱʡҽ�Ƹ��ʽ
    rsTmp.Filter = "����='ҽ�Ƹ��ʽ'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultPayMode.AddItem rsTmp!����
        If rsTmp!ȱʡ = 1 Then cboDefaultPayMode.ListIndex = cboDefaultPayMode.NewIndex
        rsTmp.MoveNext
    Next
     'ȱʡ�ѱ�    '���ǽ��޳������Ψһ����Ŀ(������ȱʡ�ѱ�),������Ч�ڼ估����
    rsTmp.Filter = "����='�ѱ�'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultFeeType.AddItem rsTmp!����
        If rsTmp!ȱʡ = 1 Then cboDefaultFeeType.ListIndex = cboDefaultFeeType.NewIndex
        rsTmp.MoveNext
    Next
    
    'ȱʡ�Ա�
    rsTmp.Filter = "����='�Ա�'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultSex.AddItem rsTmp!����
        If rsTmp!ȱʡ = 1 Then cboDefaultSex.ListIndex = cboDefaultSex.NewIndex
        rsTmp.MoveNext
    Next
    cboDefaultSex.AddItem "��"
    'ȱʡ���㷽ʽ
    Call Load֧����ʽ

    strTmp = zlDatabase.GetPara("ȱʡ���ʽ", glngSys, mlngModul, , Array(cboDefaultPayMode), blnParSet)
    zlcontrol.CboLocate cboDefaultPayMode, strTmp
    strTmp = zlDatabase.GetPara("ȱʡ�ѱ�", glngSys, mlngModul, , Array(cboDefaultFeeType), blnParSet)
    zlcontrol.CboLocate cboDefaultFeeType, strTmp
    strTmp = zlDatabase.GetPara("ȱʡ�Ա�", glngSys, mlngModul, , Array(cboDefaultSex), blnParSet)
    zlcontrol.CboLocate cboDefaultSex, strTmp
    If cboDefaultSex.ListIndex = -1 Or strTmp = "��" Then cboDefaultSex.ListIndex = cboDefaultSex.ListCount - 1
    strTmp = zlDatabase.GetPara("ȱʡ���㷽ʽ", glngSys, mlngModul, , Array(cboDefaultBalance), blnParSet)
    zlcontrol.CboLocate cboDefaultBalance, strTmp
 
    'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    chkRigistHeadSort.Value = IIf(zlDatabase.GetPara("������ͷ����", glngSys, mlngModul, , Array(chkRigistHeadSort), blnParSet) = "1", 1, 0)
    '��ȡ���õĹҺſ���
    str����ID = zlDatabase.GetPara("�Һſ���", glngSys, mlngModul, , Array(lstDept), blnParSet)
    If str����ID = "" Then
        For i = 0 To lstDept.ListCount - 1
            lstDept.Selected(i) = True
        Next
    Else
        For i = 0 To lstDept.ListCount - 1
            lstDept.Selected(i) = InStr(1, "," & str����ID & ",", "," & lstDept.ItemData(i) & ",") > 0
        Next
    End If
    If lstDept.ListCount > 0 Then lstDept.TopIndex = 0: lstDept.ListIndex = 0
    
    '��ȡ���ù��ùҺ�Ʊ������
    Call LoadFactList(4)
    
    '��ȡ���õľ��￨����
     Call InitShareInvoice
    If tbPage.TabVisible(0) Then tbPage.Tab = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lstDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Key <> Item.Key Then lvwBill.ListItems(i).Checked = False
    Next
    Item.Selected = True
End Sub
Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-07-06 18:41:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String, rsҽ�ƿ���� As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim strȱʡҽ�ƿ� As String, lngȱʡҽ�ƿ� As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '�ָ��п��
    lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, , , True, intType))
    '90875:���ϴ�,2016/11/8,ҽ�ƿ�֤������
    gstrSQL = "Select ID,����,����, nvl(�Ƿ�̶�,0) as �Ƿ�̶�  from ҽ�ƿ����  Where nvl(�Ƿ�����,0)=1 And nvl(�Ƿ�֤��,0)=0 "
    On Error GoTo Hd
    Set rsҽ�ƿ���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    rsҽ�ƿ����.Filter = "����='���￨' and �Ƿ�̶�=1"
    If rsҽ�ƿ����.EOF = False Then
        strȱʡҽ�ƿ� = rsҽ�ƿ����!����: lngȱʡҽ�ƿ� = Val(rsҽ�ƿ����!id)
    End If
    With rsҽ�ƿ����
        cboType.Clear
        rsҽ�ƿ����.Filter = 0
        If rsҽ�ƿ����.RecordCount <> 0 Then rsҽ�ƿ����.MoveFirst
        Do While Not .EOF
            cboType.AddItem Nvl(!����)
            cboType.ItemData(cboType.NewIndex) = Nvl(!id)
            If Nvl(!����) = "���￨" And cboType.ListIndex < 0 Then cboType.ListIndex = cboType.NewIndex
            If lngCardTypeID = Val(Nvl(!id)) Then
                cboType.ListIndex = cboType.NewIndex
            End If
            .MoveNext
        Loop
    End With
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "����ҽ��Ʊ���б�", False, False
    strShareInvoice = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModul, , , True)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,ҽ�ƿ����ID1|����IDn,ҽ�ƿ����IDn|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(5)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!id))
            If Val(Nvl(rsTemp!ʹ�����ID)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = strȱʡҽ�ƿ�
                .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = lngȱʡҽ�ƿ�
            Else
                rsҽ�ƿ����.Filter = "ID=" & Val(Nvl(rsTemp!ʹ�����ID))
                If Not rsҽ�ƿ����.EOF Then
                    .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = Nvl(rsҽ�ƿ����!����)
                Else
                    .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = Nvl(rsTemp!ʹ�����)
                End If
                .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = Val(Nvl(rsTemp!ʹ�����ID))
            End If
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ʊ��
    '����:���˺�
    '����:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long, lng�����ID As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '���湲��Ʊ��
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("ҽ�ƿ����")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "����ҽ�ƿ�����", strValue, glngSys, mlngModul, blnHavePrivs
    If cboType.ListIndex >= 0 Then
        lng�����ID = cboType.ItemData(cboType.ListIndex)
    End If
    Call zlDatabase.SetPara("ȱʡҽ�ƿ����", lng�����ID, glngSys, mlngModul, blnHavePrivs)
End Sub
 
