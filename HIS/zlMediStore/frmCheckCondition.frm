VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�̵���������"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "����(&1)"
      TabPicture(0)   =   "frmCheckCondition.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl�̵㷽ʽ"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl�ⷿ"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lvw����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkNoNum"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkNum"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Chk����"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Cbo�̵㷽ʽ"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbo�ⷿ"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chk�����̵�ʱ��"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "ҩƷ����(&2)"
      TabPicture(1)   =   "frmCheckCondition.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tvw����"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�ⷿ��λ(&3)"
      TabPicture(2)   =   "frmCheckCondition.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chk��λ"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "vsfStock"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CheckBox chk��λ 
         Caption         =   "����ʾ��ǰ�ⷿ�ѷ���Ļ�λ"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chk�����̵�ʱ�� 
         Caption         =   "ʼ���Ե�ǰ�����Ϊ��������"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   4800
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   420
         Width           =   3045
      End
      Begin VB.ComboBox Cbo�̵㷽ʽ 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3180
         Width           =   3045
      End
      Begin VB.CheckBox Chk���� 
         Appearance      =   0  'Flat
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3555
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   810
         Width           =   675
      End
      Begin VB.CheckBox chkNum 
         Caption         =   "���޿���¼ҩƷ"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   4020
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkNoNum 
         Caption         =   "���������������п������۵�ҩƷ"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   4425
         Width           =   3585
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1140
         TabIndex        =   7
         Top             =   3600
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   162136067
         CurrentDate     =   36901
      End
      Begin MSComctlLib.ListView Lvw���� 
         Height          =   2085
         Left            =   420
         TabIndex        =   10
         Top             =   1020
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   3678
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.TreeView tvw���� 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   8705
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStock 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   18
         Top             =   840
         Width           =   4575
         _cx             =   8070
         _cy             =   8070
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCheckCondition.frx":0054
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         Editable        =   1
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
      Begin VB.Label lbl�ⷿ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Lbl�̵㷽ʽ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʽ(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   13
         Top             =   3240
         Width           =   630
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��(&T)"
         Height          =   180
         Left            =   480
         TabIndex        =   12
         Top             =   3660
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   11
         Top             =   810
         Width           =   630
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5160
      TabIndex        =   2
      Top             =   4680
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5160
      TabIndex        =   1
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5160
      TabIndex        =   0
      Top             =   405
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5160
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCondition.frx":00C4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCheckCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mblnBootUp As Boolean
Private mstr���� As String
Private mlng�ⷿID As Long
Private mlng�̵㷽ʽ As Integer
Private mstr�̵�ʱ�� As String
Private mint���޿��ҩƷ As Integer
Private mstr�ⷿ��λ  As String
Private mbln��ҩ�ⷿ As Boolean                         '�������ҩ������ҩ�⣬��ñ���Ϊ��
Private mfrmMain As Form
Private mblnCheckNoNum  As Boolean
Private mstr����ID As String
Private mbln�����̵�ʱ�� As Boolean
Private mstrSelRows As String
Private mint�̵�ʱ�䷶Χ As Integer         '����ģ��������õ��̵�ʱ�䷶Χ

Private Sub CheckItem(ByVal str���� As String, ByVal intChecked As Integer)
    Dim lngRows As Long
    
    With vsfStock
        If str���� = "����" Then
            '��ǰ�Ǹ��ڵ㣬���пⷿ����Ϊȫѡ��ȫ��ѡ
            For lngRows = 2 To .rows - 1
                .Cell(flexcpChecked, lngRows, .ColIndex("ѡ��")) = intChecked
            Next
        Else
            '��ǰ���ӽڵ㣬��Ҫ�ݹ鴦���¼��ӽڵ�
            For lngRows = 2 To .rows - 1
                If .TextMatrix(lngRows, .ColIndex("�ϼ�")) = str���� Then
                    .Cell(flexcpChecked, lngRows, .ColIndex("ѡ��")) = intChecked
                    
                    Call CheckItem(.TextMatrix(lngRows, .ColIndex("����")), intChecked)
                End If
            Next
        End If
    End With
End Sub
Public Function GetCondition(FrmMain As Form, ByRef str���ͱ��� As String, _
    ByRef lng�ⷿID As Long, ByRef �̵㷽ʽ As Integer, ByRef str�̵�ʱ�� As String, _
    ByRef int���޿��ҩƷ As Integer, ByRef str�ⷿ��λ As String, ByRef bln���޿���н��ҩƷ As Boolean, _
    ByRef str����ID As String, ByRef bln�����̵�ʱ�� As Boolean) As Boolean
    
    mstr���� = ""
    mlng�ⷿID = 0
    mlng�̵㷽ʽ = 0
    mstr�̵�ʱ�� = ""
    mint���޿��ҩƷ = 0
    mstr�ⷿ��λ = "����"
    mblnSelect = False
    mblnCheckNoNum = False
    mbln�����̵�ʱ�� = False
    
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    
    str���ͱ��� = mstr����
    lng�ⷿID = mlng�ⷿID
    �̵㷽ʽ = mlng�̵㷽ʽ
    str�̵�ʱ�� = mstr�̵�ʱ��
    int���޿��ҩƷ = mint���޿��ҩƷ
    str�ⷿ��λ = mstr�ⷿ��λ
    bln���޿���н��ҩƷ = mblnCheckNoNum
    str����ID = mstr����ID
    bln�����̵�ʱ�� = mbln�����̵�ʱ��
    
End Function

Private Sub GetSubItem(ByVal str�ϼ����� As String, ByVal rsData As ADODB.Recordset)
    '�õݹ��㷨�����������Ŀ
    Dim rsClone As ADODB.Recordset
    
    Set rsClone = rsData.Clone
    
    rsClone.Filter = "�ϼ�='" & str�ϼ����� & "'"
    rsClone.Sort = "����"
    
    'û�ҵ���һ��ʱһ��Ҫ�˳�
    If rsClone.RecordCount = 0 Then Exit Sub
    
    With vsfStock
        .Redraw = flexRDNone
        
        Do While Not rsClone.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsClone!���� + 1
            .TextMatrix(.rows - 1, .ColIndex("�ϼ�")) = rsClone!�ϼ�
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsClone!����
            .TextMatrix(.rows - 1, .ColIndex("ѡ��")) = 0
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsClone!����
            
            '����һ������Ŀ
            Call GetSubItem(rsClone!����, rsData)
            
            rsClone.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub Load�ⷿ��λ()
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim intLevel As Integer
    
    On Error GoTo errHandle
    If chk��λ.Value = 1 Then
        gstrSQL = "Select a.����, a.���� From ҩƷ�ⷿ��λ A " & _
            " Where a.�ⷿid = [1] And Exists (Select 1 From ҩƷ��λ���� B Where b.�ⷿid = a.�ⷿid And b.��λid = a.Id) " & _
            " Order By ���� "
    Else
        gstrSQL = "Select ����, ���� From  ҩƷ�ⷿ��λ Where �ⷿid = [1] Order By ���� "
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҩƷ�ⷿ��λ", Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)))
    
    If rsData.RecordCount = 0 Then
        vsfStock.rows = 1
        Exit Sub
    End If
    
    With vsfStock
        .Redraw = flexRDNone
        .rows = 1
        
        Do While Not rsData.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("ѡ��")) = 0
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsData!����
            .TextMatrix(.rows - 1, .ColIndex("����")) = rsData!����
            
            rsData.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo�ⷿ_Click()
    Dim blnEXIST As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
    mbln��ҩ�ⷿ = False
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ��������˵�� " & _
             " Where �������� Like '��ҩ%' And ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��鲿������]", Me.cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))

    If Not rsTemp.EOF Then mbln��ҩ�ⷿ = True
    
    gstrSQL = "Select Distinct J.����,J.���� " & _
             " From ����ִ�п��� A,ҩƷ���� B,ҩƷ���� J " & _
             " Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.����" & _
             " And A.ִ�п���ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", Me.cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    Lvw����.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            If blnEXIST = False Then
                blnEXIST = (!���� = "����")
            End If
            Lvw����.ListItems.Add , "K" & !����, !����, , 1
            .MoveNext
        Loop
        If mbln��ҩ�ⷿ And blnEXIST = False Then
            Lvw����.ListItems.Add , "KK1", "����", , 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkNoNum_Click()
    chkNum.Enabled = chkNoNum.Value = 0 '����ѡ�����������������п������۵�ҩƷ��ʱ�������޿���¼ҩƷ���ſ���
End Sub

Private Sub chkNum_Click()
    chkNoNum.Enabled = chkNum.Value = 0 '����ѡ�����޿���¼ҩƷ��ʱ�������������������п������۵�ҩƷ���ſ���
End Sub

Private Sub chk��λ_Click()
    Load�ⷿ��λ
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intItem As Integer, intItems As Integer
    
    'ȡ�ü��ͣ���ѡ��������ȡҩƷ����Ҫ�ֹ�¼�룩
    mstr���� = ""
    
    If Chk����.Value = 1 Then
        mstr���� = "����"
    Else
        intItems = Me.Lvw����.ListItems.count
        For intItem = 1 To intItems
            If Lvw����.ListItems(intItem).Checked Then
                mstr���� = mstr���� & "," & Lvw����.ListItems(intItem).Text
            End If
        Next
    
        If mstr���� <> "" Then mstr���� = Mid(mstr����, 2)
    End If

    'ȡ��ҩƷ���ࣨ��ѡ�����ʾ���з��ࣩ
    mstr����ID = ""
    For intItem = 1 To tvw����.Nodes.count
        If tvw����.Nodes(intItem).Key = "Root" And tvw����.Nodes(intItem).Checked = True Then
            mstr����ID = ""
            Exit For
        ElseIf tvw����.Nodes(intItem).Key <> "Root" And _
            tvw����.Nodes(intItem).Key <> "_�г�ҩ" And _
            tvw����.Nodes(intItem).Key <> "_�в�ҩ" And _
            tvw����.Nodes(intItem).Key <> "_����ҩ" And _
            tvw����.Nodes(intItem).Checked Then
            mstr����ID = mstr����ID & "," & Mid(tvw����.Nodes(intItem).Key, 2)
        End If
    Next

    If mstr����ID <> "" Then
        mstr����ID = Mid(mstr����ID, 2)
    End If
    
    mlng�ⷿID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    mlng�̵㷽ʽ = Cbo�̵㷽ʽ.ItemData(Cbo�̵㷽ʽ.ListIndex)
    mstr�̵�ʱ�� = Format(dtpDate.Value, "yyyy-MM-dd hh:mm:ss")
    mblnSelect = True
    mint���޿��ҩƷ = chkNum.Value
    mbln�����̵�ʱ�� = (chk�����̵�ʱ��.Value = 1)
    
    'ȡ�ÿⷿ��λ����ѡ�ⷿ��ʾ�����Ǵ洢�ⷿ��
    mstr�ⷿ��λ = ""
    With vsfStock
        For intItem = 1 To .rows - 1
            If .Cell(flexcpChecked, intItem, .ColIndex("ѡ��")) = flexChecked Then
                mstr�ⷿ��λ = .TextMatrix(intItem, .ColIndex("����")) & "," & mstr�ⷿ��λ
            End If
        Next
    End With
    
'    If mstr�ⷿ��λ <> "" Then
'        mstr�ⷿ��λ = Mid(mstr�ⷿ��λ, 2)
'    End If
    
    mblnCheckNoNum = chkNoNum.Value
    
    frmNewCheckCard.txtStock.Caption = cbo�ⷿ.Text
    frmNewCheckCard.txtStock.Tag = mlng�ⷿID
    frmNewCheckCard.txtCheckDate = mstr�̵�ʱ��
'    frmCheckCard.CmdSave.Enabled = False
'    frmCheckCard.cmdCancel.Enabled = False
    
    Unload Me
End Sub



Private Sub Command1_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    If mblnBootUp = False Then
        Unload Me
        Exit Sub
    End If
    
    Call Load�ⷿ��λ
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim blnSelectStock As String
    Dim objnode As Node
    
    On Error GoTo errHandle
    blnSelectStock = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�̵����", "�ⷿ", "0")
    mint�̵�ʱ�䷶Χ = Val(zlDatabase.GetPara("�̵�ʱ�䷶Χ����", glngSys, 1307, 30))
    dtpDate.MinDate = CDate(Format(DateAdd("d", -mint�̵�ʱ�䷶Χ, Date), "yyyy-mm-dd") & " 00:00:00")
    'ҩƷ����Ȩ�޿���
    
    dtpDate.Value = Format(Sys.Currentdate, dtpDate.CustomFormat)
    dtpDate.MaxDate = dtpDate.Value
    
    mblnBootUp = False

    With Cbo�̵㷽ʽ
        .Clear
        .AddItem "ÿ��"
        .ItemData(.NewIndex) = 1
        .AddItem "ÿ��"
        .ItemData(.NewIndex) = 2
        .AddItem "ÿ��"
        .ItemData(.NewIndex) = 3
        .AddItem "ÿ����"
        .ItemData(.NewIndex) = 4
        .AddItem "�����̵㷽ʽ"
        .ItemData(.NewIndex) = 5
        .ListIndex = 0
    End With
    
    With mfrmMain.cboStock
        cbo�ⷿ.Clear
        For i = 0 To .ListCount - 1
            cbo�ⷿ.AddItem .List(i)
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = .ItemData(i)
        Next
        cbo�ⷿ.ListIndex = .ListIndex
    End With
        
    If zlStr.IsHavePrivs(gstrprivs, "���пⷿ") Then
        If blnSelectStock = "0" Then
            cbo�ⷿ.Enabled = False
        Else
            cbo�ⷿ.Enabled = True
        End If
    Else
        cbo�ⷿ.Enabled = False
    End If
    
    '�ⷿ��λ
    Load�ⷿ��λ
    
    'ҩƷ����
    gstrSQL = "Select Level as ��,ID,�ϼ�ID,����,DECODE(����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') As ���� " & _
        " From ���Ʒ���Ŀ¼" & _
        " Where ���� in (1,2,3)" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Order by Level,����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ����")

    Set objnode = tvw����.Nodes.Add(, , "Root", "������;", 1)
    Set objnode = tvw����.Nodes.Add("Root", 4, "_����ҩ", "����ҩ", 1)
    Set objnode = tvw����.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", 1)
    Set objnode = tvw����.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", 1)

    Do While Not rsTemp.EOF
        If rsTemp!�� = 1 Then
            Set objnode = tvw����.Nodes.Add("_" & rsTemp!����, 4, "_" & rsTemp!id, rsTemp!����, 1)
        Else
            Set objnode = tvw����.Nodes.Add("_" & rsTemp!�ϼ�ID, 4, "_" & rsTemp!id, rsTemp!����, 1)
        End If
        rsTemp.MoveNext
    Loop
    tvw����.Nodes("Root").Selected = True
    tvw����.Nodes("Root").Expanded = True
    
    mblnBootUp = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk����_Click()
    If Chk����.Value = 2 Then Exit Sub
    Call SetSelect(Lvw����, Chk����.Value)
End Sub

Private Sub Lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ItemCheck(Lvw����, Item, Chk����)
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem, ByVal chkObj As CheckBox)
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.count Then
            chkObj.Value = 1
        ElseIf intCount > 0 Then
            chkObj.Value = 2
        Else
            chkObj.Value = 0
        End If
    End With
End Sub

Private Sub tvw����_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode tvw����, Node, Node.Checked
End Sub
Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal objMyTreeView As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If objMyTreeView.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = objMyTreeView.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If objMyTreeView.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode objMyTreeView, Node, blnCheck
        End If
    End If
End Sub

Private Sub vsfStock_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    Dim lngRow As Long
    Dim strFlag As String

    With vsfStock
        If .MouseRow <= 0 Then Exit Sub
        If .MouseCol <> .ColIndex("ѡ��") Then Exit Sub
        
        mstrSelRows = ""

        For lngRow = 1 To .rows - 1
            If .IsSelected(lngRow) Then
                mstrSelRows = IIf(mstrSelRows = "", "", mstrSelRows & ",") & lngRow
            End If
        Next
    End With
End Sub


Private Sub vsfStock_Click()
    Dim IntCheck As Integer
    Dim lngRow As Long
    Dim lngRows As Long
    
    With vsfStock
        If .MouseRow <= 0 Then Exit Sub
        If .MouseCol <> .ColIndex("ѡ��") Then Exit Sub
        
        lngRow = .MouseRow
        
        IntCheck = .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"))

        If InStr(1, mstrSelRows, ",") > 0 Then
            '��ѡ����
            For lngRows = 1 To .rows - 1
                If InStr(1, "," & mstrSelRows & ",", "," & lngRows & ",") > 0 Then
                    .Cell(flexcpChecked, lngRows, .ColIndex("ѡ��")) = IntCheck
                End If
            Next
            
            mstrSelRows = ""
        Else
            '��ѡʱ����ǰ�ڵ�
            .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��")) = IntCheck
        End If
        
    End With
End Sub
