VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAuditItemShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���������Ŀ"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9900
   Icon            =   "frmAuditItemShow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9900
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   3690
      ScaleHeight     =   450
      ScaleWidth      =   5010
      TabIndex        =   13
      Top             =   4980
      Width           =   5010
      Begin VB.CommandButton cmdSearch 
         Height          =   360
         Left            =   2610
         Picture         =   "frmAuditItemShow.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   375
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   930
         TabIndex        =   17
         Top             =   60
         Width           =   1650
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "����(&A)"
         Height          =   360
         Left            =   0
         TabIndex        =   16
         Top             =   30
         Width           =   885
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   360
         Left            =   4995
         TabIndex        =   15
         Top             =   30
         Width           =   885
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��(&O)"
         Height          =   360
         Left            =   4050
         TabIndex        =   14
         Top             =   30
         Width           =   885
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3075
         TabIndex        =   19
         Top             =   135
         Width           =   90
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   0
      Left            =   4260
      ScaleHeight     =   2715
      ScaleWidth      =   5010
      TabIndex        =   11
      Top             =   1875
      Width           =   5010
      Begin VSFlex8Ctl.VSFlexGrid vsfAuditItem 
         Height          =   4695
         Left            =   105
         TabIndex        =   12
         Top             =   150
         Width           =   6270
         _cx             =   11060
         _cy             =   8281
         Appearance      =   2
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
         BackColorSel    =   16772055
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
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   5520
      Index           =   2
      Left            =   90
      ScaleHeight     =   5520
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   300
      Width           =   3015
      Begin VB.PictureBox pic������Ϣ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   45
         ScaleHeight     =   1695
         ScaleWidth      =   2790
         TabIndex        =   4
         Top             =   2565
         Width           =   2790
         Begin VB.PictureBox picFAXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2415
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   5
            Top             =   75
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   10
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label lbl�������� 
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   9
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label lbl����ʱ�� 
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��:"
            Height          =   195
            Left            =   225
            TabIndex        =   8
            Top             =   1365
            Width           =   2580
         End
         Begin VB.Label lbl�ܷ� 
            BackStyle       =   0  'Transparent
            Caption         =   "�ܷ�:"
            Height          =   195
            Left            =   225
            TabIndex        =   7
            Top             =   705
            Width           =   2580
         End
         Begin VB.Label lbl�ֶ��� 
            BackStyle       =   0  'Transparent
            Caption         =   "�ֶ���:"
            Height          =   195
            Left            =   225
            TabIndex        =   6
            Top             =   1035
            Width           =   2580
         End
      End
      Begin VB.PictureBox picTree 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   60
         ScaleHeight     =   1590
         ScaleWidth      =   2940
         TabIndex        =   1
         Top             =   240
         Width           =   2940
         Begin MSComctlLib.TreeView tvwAuditType 
            Height          =   1200
            Left            =   495
            TabIndex        =   2
            Top             =   420
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   2117
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            Appearance      =   0
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "����׼"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   3
            Top             =   90
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   990
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditItemShow.frx":D0A4
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditItemShow.frx":DEF6
            Key             =   "RootSel"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   7125
      Picture         =   "frmAuditItemShow.frx":E16A
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgOpen 
      Height          =   225
      Left            =   4245
      Picture         =   "frmAuditItemShow.frx":E1B9
      Top             =   7410
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   495
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAuditItemShow.frx":E20E
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin VB.Image imgBG 
      Height          =   1695
      Left            =   4635
      Picture         =   "frmAuditItemShow.frx":E222
      Top             =   7380
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   1770
      Picture         =   "frmAuditItemShow.frx":E3E0
      Top             =   7365
      Visible         =   0   'False
      Width           =   2790
   End
End
Attribute VB_Name = "frmAuditItemShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public mlngObject As Long                               '�������
Public mstrSummitID As String                              '�ύID
Public mRetMain As ADODB.Recordset                                  '�����¼��

Public mlngOk As Long
Private mstrSaveKey             As String               '������ϴεķ���ѡ��ؼ���
Private mintTypeID              As Integer              '�����������޸ġ�ɾ��ʱ��ID
Private mintItemID              As Integer              '��Ŀ�������޸ġ�ɾ��ʱ��ID
Private mRsAuditItem            As ADODB.Recordset      '���ݼ�
Private mlngCurFAID             As Long                 '��ǰ����ID
Private mblnProgUsed            As Boolean              '�����Ƿ���ʹ��
Private mblnCheckAll            As Boolean              '�Ƿ���ʾ�¼�
Private zlCheck                 As New clsCheck         '�����
Public mstr����  As String
Public mstr��ֵ  As String
Public mstr���� As String
Public mstrID As String
Private mlng����ID As Long

Private Const con_vsfField = "/*+ rule */ '' as ͼ��,a.id, a.����id,a.����,a.����,a.����,a.��ֵ,a.����,b.���� as ����,decode(a.���ö���,1,'סԺҽ��',2,'סԺ����',3,'������',4,'�����¼',5,'��ҳ��¼',6,'ҽ������',7,'����֤��',8,'֪���ļ�','δ����') as ���ö���,a.˵��,a.�������,���ö��� as ���ñ���,�ļ�ID,���û���"
Private Const conFieldFiles = "Select /*+ rule */ a.id as �ļ�ID,a.��� as �ļ�����,a.���� as �ļ�����,a.˵�� as �ļ�˵��" & vbCrLf & _
                         "from �����ļ��б� A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B " & vbCrLf & _
                         "where /*+ rule */a.id = b.COLUMN_VALUE And a.���� = [2]"


'���ڵ㶨λ
Dim nod                         As Node
Dim i                           As Long
Dim FirstKey                    As String
Dim v                           As Variant

Public Function zlInitData(ByVal RetMain As ADODB.Recordset, ByVal lngObject As Long, ByVal strSummitID As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Set mRetMain = RetMain
    mlngObject = lngObject
    mstrSummitID = strSummitID
End Function

Private Sub cmdAll_Click()
    Call DataFill
End Sub

Private Sub cmdCancel_Click()
    mlngOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    mlngOk = True
    With vsfAuditItem
        
        If .Row > 0 Then
            mstr���� = .TextMatrix(.Row, .ColIndex("����"))
            mstrID = .RowData(.Row)
            mstr���� = .TextMatrix(.Row, .ColIndex("����"))
            mstr��ֵ = .TextMatrix(.Row, .ColIndex("��ֵ"))
        End If
    End With
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    If txtSearch.Text = "" Then Exit Sub
    Call GetAuditItem(mlngObject, mstrSummitID, txtSearch.Text)
End Sub

'==============================================================================
'=���ܣ� ����ָ�
'==============================================================================
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo ErrH
    
    Select Case Item.ID
        Case 1
            Item.Handle = picPane(2).hWnd
        Case 2
            Item.Handle = picPane(0).hWnd
        Case 3
            Item.Handle = picPane(1).hWnd
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    On Error GoTo ErrH
    mblnCheckAll = True
    Call ExecuteCommand("��ʼ�ؼ�")
    Call ExecuteCommand("��������")
    picFAXX.Picture = imgClose.Picture
    Call RestoreWinState(Me, App.ProductName)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrH
    Call SetPaneRange(dkpMain, 1, 200, 100, 200, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 2, 150, 100, 150, Me.ScaleWidth)
    Call SetPaneRange(dkpMain, 3, 350, 30, 350, 30)
    
    dkpMain.RecalcLayout
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'******************************************************************************************************************
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
    
    On Error GoTo ErrH
    Dim strF As String
    Dim strTvwName As String
    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        Call InitControl
    Case "��ȡ���������Ŀ"
        Call DataAuditItem
    Case "��������"
        Call DataFill
    End Select
    ExecuteCommand = True
    
    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� �ؼ���ʼ��
'==============================================================================
Private Sub InitControl()
    
    On Error GoTo ErrH
    
    Call InitVsflexGrid
    Call InitCommandBar
    Call InitDockPannel
    Call InitTreeView
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ���򻮷�
'==============================================================================
Private Sub InitDockPannel()
    Dim objPane As Pane

    On Error GoTo ErrH
    
    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 200, 100, DockRightOf, Nothing)
    objPane.Title = "��ϸ"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 300, 30, DockBottomOf, Nothing)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption

    dkpMain.SetCommandBars cbsMain
    
    Call DockPannelInit(dkpMain)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ������ VsflexGrid
'==============================================================================
Private Sub InitVsflexGrid()
    Dim strField        As String
    Dim strFieldWidth   As String
    Dim varField        As Variant
    Dim varFieldWidth   As Variant
    Dim i               As Integer
    On Error GoTo ErrH
    vsfAuditItem.FocusRect = flexFocusNone
    vsfAuditItem.ExtendLastCol = True
    vsfAuditItem.ExplorerBar = flexExSortShowAndMove
    vsfAuditItem.AutoResize = False
    gstrSQL = "" & _
        "Select " & con_vsfField & vbCrLf & _
        "From �������Ŀ¼ a,(SELECT /*+ rule */ id,���� FROM ���������� START WITH id=[1] CONNECT BY PRIOR ID = �ϼ�ID)b " & vbCrLf & _
        "Where a.����id = b.ID and 1=0"
    Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, -1)
    Set vsfAuditItem.DataSource = mRsAuditItem
    With vsfAuditItem
        .ColWidth(0) = 250
        .MergeCol(.ColIndex("����id")) = True
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColWidth(.ColIndex("����")) = 3000
        .ColWidth(.ColIndex("ͼ��")) = 450
        .ColWidth(.ColIndex("���ö���")) = 2000
        .ColWidth(.ColIndex("��ֵ")) = 500
        .ColWidthMin = 450
        
'        .FrozenCols = 3
        If GetPersonSet Then
            'ʹ�ø��Ի����á����ѱ���ĸ�ʽ��
            strField = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrid", vsfAuditItem.Name & "����", "")
            strFieldWidth = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrid", vsfAuditItem.Name & "���", "")
            varField = Split(strField, ",")
            varFieldWidth = Split(strFieldWidth, ",")
            For i = 0 To UBound(varField)
                If varField(i) <> "" And Val(varFieldWidth(i)) <> 0 Then
                    .ColPosition(.ColIndex(varField(i))) = i
                    .ColWidth(i) = Val(varFieldWidth(i))
                End If
            Next
        End If
        .ColWidth(.ColIndex("ID")) = 0: .ColHidden(.ColIndex("ID")) = True
        .ColWidth(.ColIndex("����id")) = 0: .ColHidden(.ColIndex("����id")) = True
        .ColWidth(.ColIndex("���ñ���")) = 0: .ColHidden(.ColIndex("���ñ���")) = True
        .ColWidth(.ColIndex("�������")) = 0: .ColHidden(.ColIndex("�������")) = True
        .ColWidth(.ColIndex("�ļ�ID")) = 0: .ColHidden(.ColIndex("�ļ�ID")) = True
        .ColWidth(.ColIndex("���û���")) = 0: .ColHidden(.ColIndex("���û���")) = True
        .ColWidth(.ColIndex("����")) = 0: .ColHidden(.ColIndex("����")) = True
    End With
    DoEvents
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=���ܣ� ����������
'==============================================================================
Private Sub InitTreeView()
    Dim rsTree      As ADODB.Recordset
    Dim intStartid As Integer
    On Error GoTo ErrH

    'Tree�ĳ�ʼ��
    Set tvwAuditType.ImageList = GetImageList(16)
    tvwAuditType.Nodes.Clear
    
    gstrSQL = "Select ID,����,����ʱ�� From ������鷽�� where ����ʱ�� is not null"
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    Do Until rsTree.EOF
        If zlCommFun.NVL(rsTree!����ʱ��) <> "" Then
            intStartid = rsTree!ID
        End If
        Set nod = tvwAuditType.Nodes.Add(, , "Root" & rsTree!ID, zlCommFun.NVL(rsTree!����, "Ĭ�Ϸ���"), 20, 20)
        nod.Expanded = True
            
        rsTree.MoveNext
    Loop
    
'    '��Ӹ��ڵ�
'    Set nod = tvwAuditType.Nodes.Add(, , "Root", "����", 20, 20)
'    nod.Expanded = True

    gstrSQL = "SELECT /*+ rule */ id,�ϼ�ID,����ID,����,���� FROM ���������� where ����ID=" & intStartid & " START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID "
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    rsTree.Sort = "����"
    i = 1
    Do Until rsTree.EOF
        '����ӽڵ�
        Set nod = tvwAuditType.Nodes.Add(IIf("" & rsTree("�ϼ�ID") = "", "Root" & rsTree("����ID"), "A" & rsTree("�ϼ�ID")), tvwChild, "A" & rsTree("ID"), "��" + "" & rsTree("����") + "��" + "" & rsTree("����"), 23, 24)
        If i = 1 Then FirstKey = nod.Key
        If FirstKey = nod.Key Then i = 2
        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
        rsTree.MoveNext
    Loop
    FirstKey = "A" & mintTypeID
    For Each v In tvwAuditType.Nodes
        If v.Key = FirstKey Then
            '����ѡ��
            v.Selected = True
            v.EnsureVisible
        End If
    Next
    If tvwAuditType.SelectedItem Is Nothing Then
        tvwAuditType.Nodes("Root" & intStartid).Selected = True
        tvwAuditType.Nodes("Root" & intStartid).Bold = True
        tvwAuditType.Nodes("Root" & intStartid).Tag = 1
    End If
    DoEvents
    tvwAuditType_NodeClick tvwAuditType.SelectedItem
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Err.Clear
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ�˵�������
'==============================================================================
Private Sub InitCommandBar()

    
    On Error GoTo ErrH

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Visible = False
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub tvwAuditType_NodeClick(ByVal Node As MSComctlLib.Node)
        
    If mstrSaveKey = Node.Key Then Exit Sub
    If Left(Node.Key, 4) = "Root" Then
        vsfAuditItem.Rows = 1
        mstrSaveKey = Node.Key
        mlngCurFAID = Replace(mstrSaveKey, "Root", "")
        If Node.Tag = "1" Then
            mblnProgUsed = True
        Else
            mblnProgUsed = False
        End If
        Call DataUpdate
        Exit Sub
    End If
    mstrSaveKey = Node.Key
    
    Call ExecuteCommand("��ȡ���������Ŀ")
    
End Sub

'==============================================================================
'=���ܣ� ����ͳ��
'==============================================================================
Private Sub DataUpdate()
    Dim rs              As ADODB.Recordset
    Dim lng�ܷ�         As Double
    On Error GoTo ErrH
    gstrSQL = "Select ����,�ܷ�,�ֶ���,����ʱ��,ͣ��ʱ��,˵�� From ������鷽�� where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCurFAID)
    If Not rs.EOF Then
        lbl��������.Caption = rs("����")
        lbl�ֶ���.Caption = "�ֶ���:" & rs("�ֶ���")
        lbl����ʱ��.Caption = "����ʱ��:" & zlCommFun.NVL(rs("����ʱ��"))
        lbl�ܷ�.Caption = "�ܷ�:" & rs("�ܷ�")
        lng�ܷ� = rs("�ܷ�")
    Else
        lbl��������.Caption = ""
        lbl�ֶ���.Caption = ""
        lbl����ʱ��.Caption = ""
        lbl�ܷ�.Caption = ""
    End If
    
'''    gstrSQL = "select sum(��׼��ֵ) from �������ֱ�׼ where �ϼ�ID is null and ����ID = [1]"
'''    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
'''    If Not rs.EOF Then
'''        If Abs(lng�ܷ� - rs.Fields(0)) > 0.01 Then
'''            lbl�ܷ� = lbl�ܷ� + "����Ŀ������Ϊ:" & rs.Fields(0)
'''            lbl�ܷ�.ForeColor = vbRed
'''        Else
'''            lbl�ܷ�.ForeColor = vbBlack
'''        End If
'''    Else
'''        lbl�ܷ�.ForeColor = vbRed
'''    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=���ܣ� �������ݼ��� vsfAuditItem
'==============================================================================
Private Sub DataAuditItem(Optional strWhere As String)
    Dim strKey      As String
    Dim i           As Long
    Dim nTmpNode As Node
    Dim strWhere1 As String
    
    On Error GoTo ErrH
    If strWhere = "" Then
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            If tvwAuditType.SelectedItem.Tag = "1" Then
                mblnProgUsed = True
            Else
                mblnProgUsed = False
            End If
            Exit Sub
        End If
        
        
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            mlngCurFAID = Replace(tvwAuditType.SelectedItem.Key, "Root", "")
            If tvwAuditType.SelectedItem.Tag = "1" Then
                mblnProgUsed = True
            Else
                mblnProgUsed = False
            End If
        Else
            Set nTmpNode = tvwAuditType.SelectedItem
            While Not nTmpNode.Parent Is Nothing
                Set nTmpNode = nTmpNode.Parent
            Wend
            
            If InStrRev(nTmpNode.Key, "Root") > 0 Then
                mlngCurFAID = Replace(nTmpNode.Key, "Root", "")
                If nTmpNode.Tag = "1" Then
                    mblnProgUsed = True
                Else
                    mblnProgUsed = False
                End If
            End If
        End If
        
        
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            strKey = Mid(tvwAuditType.SelectedItem.Key, 5)
        Else
            strKey = Mid(tvwAuditType.SelectedItem.Key, 2)
        End If
        
        If mstrSummitID = "" Then
            strWhere1 = " And A.�ļ�ID is Null"
        Else
            strWhere1 = " And (A.�ļ�ID is null or instr(','|| A.�ļ�ID || ',' , ','|| '" & mstrSummitID & "' || ',')>0 )"
        End If
        
        If mblnCheckAll Then
            gstrSQL = "" & _
                    "Select " & con_vsfField & vbCrLf & _
                    "From �������Ŀ¼ a,(SELECT /*+ rule */ id,���� FROM ���������� START WITH id=[1] CONNECT BY PRIOR ID = �ϼ�ID)b " & vbCrLf & _
                    "Where a.����id = b.ID and a.���ö���=[2]" & strWhere1
        Else
            gstrSQL = "" & _
                    "Select " & con_vsfField & vbCrLf & _
                    "From �������Ŀ¼ a,���������� b" & vbCrLf & _
                    "Where a.����id = b.ID and a.����id=[1] and a.���ö���=[2]" & strWhere1
        End If
        Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, strKey, mlngObject)
    Else
        gstrSQL = "" & _
                "Select " & con_vsfField & vbCrLf & _
                "From �������Ŀ¼ a,���������� b" & vbCrLf & _
                "Where a.����id = b.ID and a.���ö���=[1] And" & vbCrLf & strWhere
        Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mlngObject)
    End If
    Set vsfAuditItem.DataSource = mRsAuditItem
       
    With vsfAuditItem
        If .Rows > 1 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpPictureAlignment, i, .ColIndex("ͼ��")) = flexPicAlignCenterCenter
                Select Case .Cell(flexcpText, i, .ColIndex("���ñ���"))
                    Case "1"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(2).Picture
                    Case "2"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(15).Picture
                    Case "3"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(16).Picture
                    Case "4"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(17).Picture
                    Case "5"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(18).Picture
                    Case "6"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(6).Picture
                    Case "7"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(3).Picture
                    Case "8"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(20).Picture
                End Select
            Next i
            .Row = 1
        End If
    End With
    Call DataUpdate
    Call vsfAuditItem_RowColChange
    lblInfo.Caption = "�ܹ�:" & mRsAuditItem.RecordCount & "������¼��"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataFill()
    Dim i           As Long
    
    On Error GoTo ErrH
    Set vsfAuditItem.DataSource = mRetMain
       
    With vsfAuditItem
        If .Rows > 1 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpPictureAlignment, i, .ColIndex("ͼ��")) = flexPicAlignCenterCenter
                Select Case .Cell(flexcpText, i, .ColIndex("���ñ���"))
                    Case "1"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(2).Picture
                    Case "2"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(15).Picture
                    Case "3"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(16).Picture
                    Case "4"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(17).Picture
                    Case "5"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(18).Picture
                    Case "6"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(6).Picture
                    Case "7"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(3).Picture
                    Case "8"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(20).Picture
                End Select
            Next i
            .Row = 1
        End If
    End With
    Call DataUpdate
    Call vsfAuditItem_RowColChange
    lblInfo.Caption = "�ܹ�:" & mRetMain.RecordCount & "������¼��"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsfAuditItem.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 2
'        tvwAuditType.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        
        On Error Resume Next
        pic������Ϣ.Move 0, picPane(2).ScaleHeight - pic������Ϣ.Height, picPane(2).ScaleWidth
        With picTree
            .Move 0, 0, pic������Ϣ.Width, picPane(2).Height - pic������Ϣ.Height
            .Cls
            .PaintPicture imgBGBlue.Picture, 0, 0, picTree.Width, 360, 0, 0, imgBGBlue.Width, 360
            .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, picTree.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, picTree.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picTree.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, 0, picTree.ScaleHeight - Screen.TwipsPerPixelY, picTree.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
        End With
        
        tvwAuditType.Move Screen.TwipsPerPixelX * 4, 390, Abs(picTree.ScaleWidth - 8 * Screen.TwipsPerPixelX), Abs(picTree.ScaleHeight - 390 - Screen.TwipsPerPixelY * 4)
        With pic������Ϣ
            .Cls
            .PaintPicture imgBGBlue.Picture, 0, 0, pic������Ϣ.Width, 360, 0, 0, imgBGBlue.Width, 360
            .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, pic������Ϣ.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, 0, pic������Ϣ.ScaleHeight - Screen.TwipsPerPixelY, pic������Ϣ.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
        End With
        picFAXX.Move pic������Ϣ.ScaleWidth - picFAXX.Width - 80
        Refresh
    Case 1
        cmdCancel.Move picPane(Index).Width - cmdCancel.Width - 60
        cmdOk.Move cmdCancel.Left - cmdOk.Width - 30
    End Select
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = 13 Then
       If txtSearch.Text = "" Then Exit Sub
        Call GetAuditItem(mlngObject, mstrSummitID, txtSearch.Text)
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub vsfAuditItem_DblClick()
    Call cmdOk_Click
End Sub

'==============================================================================
'=���ܣ����б任ʱ
'==============================================================================
Private Sub vsfAuditItem_RowColChange()
    Dim rsTemp          As ADODB.Recordset
    Dim varPos          As Variant
    On Error GoTo ErrH
    DoEvents
    If vsfAuditItem.Rows = 1 Then
        With frmAuditItemEdit
            .txtTypeID.Tag = "-1"
            .txtTypeID.Text = ""
            .txtName.Text = ""
            .txtCode.Text = ""
            .txtMnemonicCode.Text = ""
            .cboUsed.ListIndex = -1
            .cboLink.ListIndex = -1
            .txtDescription.Text = ""
            .txtAudit_NotCheck.Text = ""
            .txtNumValue = ""
            .CboPalValue.ListIndex = -1
            .blnProgUsed = False
            Set .vsfFiles.DataSource = Nothing
        End With
'        stbThis.Panels(2) = "��ǰ��ʾ�� 0 ����Ŀ��"
        frmAuditItemEdit.vsfFiles.Rows = 1
        Exit Sub
    End If
    If vsfAuditItem.ColIndex("ID") <= 0 Then Exit Sub
    If Val(vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))) <= 0 Then
        frmAuditItemEdit.vsfFiles.Rows = 1
        Exit Sub
    End If
    With frmAuditItemEdit
        
        .txtTypeID.Tag = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
        
        gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, CStr(Val("" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����ID")))))
        If Not zlCheck.Connection_ChkRsState(rsTemp) Then
            .txtTypeID.Tag = "" & rsTemp!ID
            .txtTypeID.Text = "[" + rsTemp!���� + "]" & rsTemp!����
        Else
            .txtTypeID.Tag = "-1"
            .txtTypeID.Text = "[ȫ��]����"
        End If
        
        .txtName.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����"))
        .txtCode.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����"))
        .txtMnemonicCode.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����"))
        .cboUsed.ListIndex = zlCheck.Cmb_EditIndex(.cboUsed, "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("���ñ���")))
        .cboLink.ListIndex = zlCheck.Cmb_EditIndex(.cboLink, "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("���û���")))
        .txtDescription.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("˵��"))
        .txtAudit_NotCheck.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("�������"))
        .txtFileID.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("�ļ�ID"))
        .txtNumValue.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("��ֵ"))
        .CboPalValue.ListIndex = IIf(vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����")) = "", 0, vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����")))
        
        .blnProgUsed = mblnProgUsed
        gstrSQL = conFieldFiles
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("�ļ�ID")), AuditFileTran(zlCheck.Cmb_ID(.cboUsed), 0))
        Set .vsfFiles.DataSource = rsTemp
'        '��ȡ �����ļ�����
'        If .txtFileID.Tag <> "" Then
'            gstrSQL = "select ���� from �����ļ��б� where ID = [1] "
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, Val(.txtFileID.Tag))
'            If Not zlCheck.Connection_ChkRsState(rsTemp) Then
'                .txtFileID.Text = "" & rsTemp.Fields!����
'            Else
'                .txtFileID.Text = ""
'            End If
'        Else
'            .txtFileID.Text = ""
'        End If
    End With
    mlng����ID = vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����ID"))
    
'    stbThis.Panels(2) = "��ǰ��ʾ�� " & vsfAuditItem.Rows - 1 & " ����Ŀ��"
    varPos = zlCheck.Connection_GetBookMark(mRsAuditItem, "ID=" & CStr("" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))))
    DoEvents
    If Not IsNull(varPos) Then
        If Val(varPos) > 0 Then mRsAuditItem.Bookmark = varPos
    End If
    
    Call TreeviewSelect(mlng����ID, tvwAuditType)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������Ϣ�رջ���ʾ
'==============================================================================
Private Sub picFAXX_Click()
    On Error GoTo ErrH
    
    If picFAXX.Tag = "" Then
        picFAXX.Tag = "Opened"
        picFAXX.Picture = imgOpen.Picture
        pic������Ϣ.Height = 340
    Else
        picFAXX.Tag = ""
        picFAXX.Picture = imgClose.Picture
        pic������Ϣ.Height = 1695
    End If
    picFAXX.Refresh
    Call picPane_Resize(2)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������Ϣ�����ɫ
'==============================================================================
Private Sub picFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    If X >= 0 And X <= picFAXX.ScaleWidth And Y >= 0 And Y <= picFAXX.ScaleHeight Then
        SetCapture picFAXX.hWnd
        '������룡����
        picFAXX.Line (0, 0)-(picFAXX.ScaleWidth - Screen.TwipsPerPixelX, picFAXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
    Else
        '����Ƴ�������
        picFAXX.Cls
        ReleaseCapture
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetAuditItem(intObject As Long, strFileID As String, Optional shortName As String = "")
Dim rsData As ADODB.Recordset, strSubid As String, strReturn As String
Dim i As Long
On Error GoTo ErrH
    If IsNumeric(strFileID) Then
        '����ļ�ID�������
        '���Ӳ�����¼ ���������ֱ��ȡ�ļ�ID����������ֱ�Ӱ����Ͷ�ȡ
        gstrSQL = "Select �ļ�ID From ���Ӳ�����¼ a Where a.ID = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFileID)
        If zlCheck.Connection_ChkRsState(rsData) Then
            strFileID = 0
        Else
            strFileID = "" & rsData.Fields!�ļ�ID
        End If
    Else
        If Not gobjEmr Is Nothing Then
            If InStr(strFileID, "|") > 0 Then
                strSubid = Split(strFileID, "|")(1)
                strFileID = Split(strFileID, "|")(0)
            End If
            gstrSQL = "Select RawtoHex(Antetype_Id) �ļ�ID From Bz_Doc_Tasks A Where Real_Doc_Id = Hextoraw(:docid)" & IIf(strSubid = "", "", " And Subdoc_Id =:subdocid")
            strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strFileID & "^" & DbType.T_String & "^docid" & IIf(strSubid = "", "", "|" & strSubid & "^" & DbType.T_String & "^subdocid"), rsData)
            If strReturn <> "" Then strFileID = 0
            If Not rsData Is Nothing Then
            If rsData.RecordCount > 0 Then
                strFileID = rsData!�ļ�ID
            End If
            End If
        End If
    End If
    If strFileID = "0" Then
        gstrSQL = "Select /*+ rule */ '' as ͼ��,A.ID, A.����ID,A.����, A.����, A.����, A.��ֵ,A.����,b.���� as ����,decode(a.���ö���,1,'סԺҽ��',2,'סԺ����',3,'������',4,'�����¼',5,'��ҳ��¼',6,'ҽ������',7,'����֤��',8,'֪���ļ�','δ����') as ���ö���,a.˵��,a.�������,���ö��� as ���ñ���,�ļ�ID,���û��� From �������Ŀ¼  A ,���������� B,������鷽�� C where  A.����ID =B.id And B.����ID =C.ID And C.����ʱ�� is Not Null And A.���ö��� = " & CStr(intObject)
    Else
        gstrSQL = "Select /*+ rule */ '' as ͼ��,A.ID, A.����ID,A.����, A.����, A.����, A.��ֵ,A.����,b.���� as ����,decode(a.���ö���,1,'סԺҽ��',2,'סԺ����',3,'������',4,'�����¼',5,'��ҳ��¼',6,'ҽ������',7,'����֤��',8,'֪���ļ�','δ����') as ���ö���,a.˵��,a.�������,���ö��� as ���ñ���,�ļ�ID,���û��� From �������Ŀ¼ A ,���������� B,������鷽�� C  where A.����ID =B.id And B.����ID =C.ID And C.����ʱ�� is Not Null And A.���ö��� = " & CStr(intObject) & " And (A.�ļ�ID is null or instr(','|| A.�ļ�ID || ',' , ','|| '" & strFileID & "' || ',')>0 )"
    End If
    If shortName <> "" Then
        shortName = UCase(shortName)
        gstrSQL = gstrSQL & vbCrLf & "And (A.���� like '%" & shortName & "%' or A.���� like '%" & shortName & "%' or A.���� like '%" & shortName & "%')"
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsData.RecordCount > 0 Then
        
        Set vsfAuditItem.DataSource = rsData
       
        With vsfAuditItem
            If .Rows > 1 Then
                For i = .FixedRows To .Rows - 1
                    .Cell(flexcpPictureAlignment, i, .ColIndex("ͼ��")) = flexPicAlignCenterCenter
                    Select Case .Cell(flexcpText, i, .ColIndex("���ñ���"))
                        Case "1"
                            .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(2).Picture
                        Case "2"
                            .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(15).Picture
                        Case "3"
                            .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(16).Picture
                        Case "4"
                            .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(17).Picture
                        Case "5"
                            .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(18).Picture
                        Case "6"
                            .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(6).Picture
                        Case "7"
                            .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(3).Picture
                        Case "8"
                            .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(20).Picture
                    End Select
                Next i
                .Row = 1
            End If
        End With
        Call DataUpdate
        Call vsfAuditItem_RowColChange
        lblInfo.Caption = "�ܹ�:" & rsData.RecordCount & "������¼��"
      
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
End Sub

Private Function TreeviewSelect(ByVal lng����ID As Long, tvwMain As TreeView)
    '����ѡ�����
    Dim i As Long
    Dim NodeMain As MSComctlLib.Node
    Dim nodeChild As MSComctlLib.Node
    Dim str����ID  As String
    str����ID = CStr(lng����ID)
    
    On Error Resume Next
    If lng����ID = 0 Then Exit Function
    If ObjPtr(tvwAuditType) > 0 Then
        Set NodeMain = tvwMain.Nodes.Item(1)
        If NodeMain.Key = "A" & str����ID Then
            NodeMain.Selected = True
            Exit Function
        End If
        
        
        If NodeMain.Children > 0 Then
            Set nodeChild = NodeMain.Child
            For i = 1 To NodeMain.Children
                If nodeChild.Key = "A" & str����ID Then
                    nodeChild.Selected = True
                    Exit For
                End If
                
                If Not nodeChild.Child Is Nothing Then
                   If SetSelectTvwChild("A" & str����ID, nodeChild) Then
                        Exit For
                   End If
                End If
                Set nodeChild = nodeChild.Next
                
            Next
        End If
    End If
End Function

Public Function SetSelectTvwChild(ByVal strTvwMain As String, tvwNode As Node) As Boolean
    Dim nodeChild As MSComctlLib.Node
    Dim i As Long
    On Error Resume Next
    
    
    If tvwNode.Key = strTvwMain Then
        tvwNode.Selected = True
        SetSelectTvwChild = True
        Exit Function
    End If
    
    If tvwNode.Children > 0 Then
        Set nodeChild = tvwNode.Child
        For i = 1 To tvwNode.Children
    
            If nodeChild.Key = strTvwMain Then
                nodeChild.Selected = True
                SetSelectTvwChild = True
                Exit Function
            End If
            
        
            If Not nodeChild.Child Is Nothing Then
               If SetSelectTvwChild(strTvwMain, nodeChild) Then
                    SetSelectTvwChild = True
                    Exit Function
               End If
            End If
            Set nodeChild = nodeChild.Next
            
        Next
    End If
    
End Function
