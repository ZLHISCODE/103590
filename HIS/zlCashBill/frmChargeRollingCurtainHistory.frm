VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmChargeRollingCurtainHistory 
   BorderStyle     =   0  'None
   Caption         =   "�շ�Ա��ʷ������Ϣ"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
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
   ScaleHeight     =   7365
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRollingCurtain 
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   285
      ScaleHeight     =   1980
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   1695
      Width           =   10170
      Begin VB.ComboBox cboDate 
         Height          =   330
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   127
         Width           =   1260
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "���¹�������(&R)"
         Height          =   350
         Left            =   7290
         TabIndex        =   5
         Top             =   105
         Width           =   1800
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtain 
         Height          =   1800
         Left            =   165
         TabIndex        =   1
         Top             =   555
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
         FormatString    =   $"frmChargeRollingCurtainHistory.frx":0000
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
         Begin VB.PictureBox picImgPlan 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   30
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   9
            Top             =   60
            Width           =   210
            Begin VB.Image imgColPlan 
               Height          =   195
               Left            =   0
               Picture         =   "frmChargeRollingCurtainHistory.frx":00B5
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   2325
         TabIndex        =   2
         Top             =   135
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   116654083
         CurrentDate     =   41520
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   4785
         TabIndex        =   3
         Top             =   135
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   116654083
         CurrentDate     =   41520
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
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
         Left            =   2355
         TabIndex        =   8
         Top             =   150
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblHistoryDate 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   180
         Width           =   840
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "��"
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
         Left            =   4485
         TabIndex        =   4
         Top             =   195
         Width           =   225
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeRollingCurtainHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModule As Long
Private Enum mPaneIndex
    EM_PN_Filter = 270101   '��������
    EM_PN_RollingList = 270101  '�����б�
    EM_PN_ChargeBillTotal = 270102   '�տƱ�ݻ���
End Enum
Private mblnNotBrush As Boolean '��ˢ������
Private mobjChargeBill As clsChargeBill
Private mlngRollingCurtainID As Long '����ID
Private mstrRollingCurtainNO As String   '���ʵ��ݺ�
Private mblnDel As Boolean   '�Ƿ�����������
Private mfrmMain As Object

Public Sub zlInitVar(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���
    '���:lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:���˺�
    '����:2013-09-09 14:41:46
    '˵��:���ش����,��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    mlngModule = lngModule: mstrPrivs = strPrivs
End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngHeight As Long
    
    With dkpMan
        'Set .ImageList = zlCommFun.GetPubIcons
        lngHeight = 3980 / Screen.TwipsPerPixelY
        Set objPane = .CreatePane(mPaneIndex.EM_PN_RollingList, 400, lngHeight, DockRightOf, Nothing)
        objPane.Title = "������Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picRollingCurtain.hWnd
        objPane.MinTrackSize.Height = lngHeight * 0.6
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_Filter, 400, 400, DockBottomOf, objPane)
        objPane.Title = "�տƱ�ݻ���": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = mobjChargeBill.GetChargeAndBillTotalForm.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function
Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ�
    '����:���˺�
    '����:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsRollingCurtain
        .Clear
        .Rows = 2: .Cols = 19: i = 1
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "���ʵ���": i = i + 1
        .TextMatrix(0, i) = "�������": i = i + 1
        .TextMatrix(0, i) = "��ʼʱ��": i = i + 1
        .TextMatrix(0, i) = "��ֹʱ��": i = i + 1
        .TextMatrix(0, i) = "������": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": i = i + 1
        .TextMatrix(0, i) = "�տ��": i = i + 1
        .TextMatrix(0, i) = "����˵��": i = i + 1
        .TextMatrix(0, i) = "��Ԥ����": i = i + 1
        .TextMatrix(0, i) = "����ϼ�": i = i + 1
        .TextMatrix(0, i) = "����ϼ�": i = i + 1
        .TextMatrix(0, i) = "С���տ���": i = i + 1
        .TextMatrix(0, i) = "С���տ�ʱ��": i = i + 1
        .TextMatrix(0, i) = "�����տ���": i = i + 1
        .TextMatrix(0, i) = "�����տ�ʱ��": i = i + 1
        .TextMatrix(0, i) = "������": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": i = i + 1
        .ColKey(0) = "����"
        .ColData(0) = "-1|1"
       For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) = "���ʵ���" Or .ColKey(i) = "��ʼʱ��" Or .ColKey(i) = "��ֹʱ��" Or _
               .ColKey(i) = "������" Or .ColKey(i) = "����ʱ��" Then
                .ColData(i) = "1|0"
            End If
            If .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Then .ColHidden(i) = True
            If .ColKey(i) = "�տ��" Then
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            End If
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "���ʵ���" Or .ColKey(i) = "�������" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*�ϼ�" Or .ColKey(i) = "��Ԥ����" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False
    End With
End Sub
Private Function LoadHistoryData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʷ��������
    '����:���ݼ��سɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 17:08:50
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, i As Long, blnDel As Boolean
    Dim dtStartDate As Date, dtEndDate As Date, strTemp As String
    Dim strValue As String
    On Error GoTo errHandle
    Call GetDateRange(dtStartDate, dtEndDate)
    If dtpEndDate - dtStartDate > 90 Then
        '���շ�Ա��Ϊͳ������
        strWhere = " And a.�տ�Ա = [1] And a.�Ǽ�ʱ��+0 Between [2] And [3] "
    Else
        '��ʱ����Ϊͳ������
        strWhere = " And a.�տ�Ա||'' = [1] And a.�Ǽ�ʱ�� Between [2] And [3] "
    End If
    
    strSQL = "" & _
    "   Select /*+ rule */a.Id,a.No As ���ʵ���, a.��ʼʱ��, a.��ֹʱ��, a.�Ǽ��� As ������, a.�Ǽ�ʱ�� As ����ʱ��,  " & _
    "         b.���� As �տ��, a.ժҪ As ����˵��, " & _
    "         ltrim(to_char(a.��Ԥ����,'9999999999990.00')) as ��Ԥ����, " & _
    "         ltrim(to_char(a.����ϼ�,'9999999999990.00')) as ����ϼ�, " & _
    "         ltrim(to_char(a.����ϼ�,'9999999999990.00')) as ����ϼ�, " & _
    "         a.С���տ���, To_Char(a.С���տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As С���տ�ʱ��,  " & _
    "         a.�����տ���,To_Char(a.�����տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �����տ�ʱ��,  " & _
    "         a.������, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, " & _
    "         a.�Ƿ�Һ�,a.�Ƿ���￨,a.�Ƿ����ѿ�,a.�Ƿ��շ�,a.Ԥ����� As �Ƿ�Ԥ��,a.�Ƿ���� " & _
    "  From ��Ա�սɼ�¼ A, ���ű� B " & _
    "  Where a.�տ��id = b.Id(+) And a.��¼���� = 1 " & strWhere & _
    "  Order by �Ǽ�ʱ�� desc,���ʵ��� desc,С���տ�ʱ�� desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, dtStartDate, dtEndDate)
    With vsRollingCurtain
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .FixedRows = 1
        Do While Not rsTemp.EOF
            strTemp = ""
            strValue = ""
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("���ʵ���")) = Nvl(rsTemp!���ʵ���)
            '0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�
            If Val(Nvl(rsTemp!�Ƿ�Һ�)) = 1 Then
                strTemp = ",�Һ�"
                strValue = ",4"
            End If
            If Val(Nvl(rsTemp!�Ƿ���￨)) = 1 Then
                strTemp = strTemp & ",���￨"
                strValue = strValue & ",5"
            End If
            If Val(Nvl(rsTemp!�Ƿ����ѿ�)) = 1 Then
                strTemp = strTemp & ",���ѿ�"
                strValue = strValue & ",6"
            End If
            If Val(Nvl(rsTemp!�Ƿ��շ�)) = 1 Then
                strTemp = strTemp & ",�շ�"
                strValue = strValue & ",1"
            End If
            If Val(Nvl(rsTemp!�Ƿ�Ԥ��)) = 1 Then
                strTemp = strTemp & ",Ԥ��"
                strValue = strValue & ",2"
            ElseIf Val(Nvl(rsTemp!�Ƿ�Ԥ��)) = 2 Then
                strTemp = strTemp & ",����Ԥ��"
                strValue = strValue & ",21"
            ElseIf Val(Nvl(rsTemp!�Ƿ�Ԥ��)) = 3 Then
                strTemp = strTemp & ",סԺԤ��"
                strValue = strValue & ",22"
            End If
            If Val(Nvl(rsTemp!�Ƿ����)) = 1 Then
                strTemp = strTemp & ",����"
                strValue = strValue & ",3"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("�������")) = Mid(strTemp, 2)
            .Cell(flexcpData, .Rows - 1, .ColIndex("�������")) = Mid(strValue, 2)
            .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = Format(Nvl(rsTemp!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = Format(Nvl(rsTemp!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Format(Nvl(rsTemp!����ʱ��), "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = Nvl(rsTemp!����˵��)
            .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = Nvl(rsTemp!��Ԥ����)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = Nvl(rsTemp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = Nvl(rsTemp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ���")) = Nvl(rsTemp!С���տ���)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ�ʱ��")) = Nvl(rsTemp!С���տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ���")) = Nvl(rsTemp!�����տ���)
            .TextMatrix(.Rows - 1, .ColIndex("�����տ�ʱ��")) = Nvl(rsTemp!�����տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Nvl(rsTemp!����ʱ��)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount <> 0 Then .Rows = .Rows - 1
'        If rsTemp.RecordCount <> 0 Then
'            Set .DataSource = rsTemp
'        End If
        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "���ʵ���" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*�ϼ�" Or .ColKey(i) = "��Ԥ����" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        For i = 1 To .Rows - 1
            blnDel = Trim(.TextMatrix(i, .ColIndex("����ʱ��"))) <> ""
            If blnDel Then
                '���ϼ�¼���ú�ɫ����
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
               
                .RowData(i) = 0
            ElseIf Trim(.TextMatrix(i, .ColIndex("С���տ�ʱ��"))) = "" _
                    And Trim(.TextMatrix(i, .ColIndex("�����տ�ʱ��"))) = "" Then
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H80000018
                    .RowData(i) = 1 '��ʾδ�տ�, &H80000018
            Else
                .RowData(i) = 0
            End If
        Next
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False
        If .Enabled And .Visible Then .SetFocus
    End With
    mblnNotBrush = False
    '������ϸ����
    Call LoadDetail
    LoadHistoryData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2013-09-11 17:46:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngRollingCurtainID = 0
    mstrRollingCurtainNO = ""
    mblnDel = False
    
    Call InitGrid '��ʼ������
    With cboDate
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 1: .ListIndex = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 2
        .AddItem "����"
        .ItemData(.NewIndex) = 3
        .AddItem "����"
        .ItemData(.NewIndex) = 4
        .AddItem "����"
        .ItemData(.NewIndex) = 5
        .AddItem "ָ��ʱ��"
        .ItemData(.NewIndex) = 9
    End With
    dtpEndDate.Value = zlDatabase.Currentdate
    dtpEndDate.MaxDate = Format(dtpEndDate.Value, "yyyy-mm-dd 23:59:59")
    dtpStartDate.Value = DateAdd("m", -1, dtpEndDate.Value)
    dtpStartDate.MaxDate = dtpEndDate.MaxDate
    Call SetCtrlVisible
End Sub
Private Sub SetCtrlVisible()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Visible����
    '����:���˺�
    '����:2013-09-11 18:21:29
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date, dtStartDate As Date
    dtpStartDate.Visible = False: dtpEndDate.Visible = False
    lblEndDate.Visible = False
    Select Case cboDate.ItemData(cboDate.ListIndex)
    Case 1 '����
        dtDate = zlDatabase.Currentdate
        lblRange.Caption = Format(dtDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 2 '����
        dtDate = DateAdd("d", -1, zlDatabase.Currentdate)
        lblRange.Caption = Format(dtDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 3 '����
        dtDate = zlDatabase.Currentdate
        dtStartDate = DateAdd("d", -1 * (Weekday(dtDate) - 2), dtDate)
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 4 '����
        dtDate = zlDatabase.Currentdate
        dtDate = DateAdd("d", -1 * (Weekday(dtDate) - 2), dtDate)
        dtStartDate = DateAdd("d", -7, dtDate)
        dtDate = DateAdd("d", 6, dtStartDate)
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 5 '����
        dtDate = zlDatabase.Currentdate
        dtStartDate = CDate(Format(dtDate, "yyyy") & "-" & Month(dtDate) & "-01")
        lblRange.Caption = Format(dtStartDate, "yyyy-mm-dd 00:00:00") & "��" & Format(dtDate, "yyyy-mm-dd 23:59:59")
        lblRange.Visible = True
    Case 9 'ָ������
        lblRange.Visible = False
        dtpStartDate.Visible = True: dtpEndDate.Visible = True
        lblEndDate.Visible = True
        If dtpStartDate.Enabled And dtpStartDate.Visible Then dtpStartDate.SetFocus
    End Select
End Sub
Private Function GetDateRange(ByRef dtStartDate As Date, ByRef dtEndDate As Date) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡʱ�䷶Χ
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 18:45:57
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant
    On Error GoTo errHandle
    Select Case cboDate.ItemData(cboDate.ListIndex)
    Case 9 'ָ������
        dtStartDate = dtpStartDate.Value
        dtEndDate = dtpEndDate.Value
    Case Else '1, 2, 3, 4, 5 '���� '����'���� '����'����
        varData = Split(lblRange.Caption, "��")
        dtStartDate = CDate(varData(0))
        dtEndDate = CDate(varData(1))
    End Select
    GetDateRange = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cboDate_Click()
    Call SetCtrlVisible
End Sub

Private Sub cmdRefresh_Click()
    Call mfrmMain.RefreshBasic
    Call LoadHistoryData
End Sub

Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then dtpStartDate.Value = dtpEndDate.Value
End Sub

Private Sub dtpStartDate_Change()
    If dtpStartDate.Value > dtpEndDate.Value Then dtpEndDate.Value = dtpStartDate.Value
End Sub

Private Sub Form_Load()
    Set mobjChargeBill = New clsChargeBill
    Call InitPanel
    Call InitFace
End Sub
Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
    Set mobjChargeBill = Nothing
End Sub
Private Sub picRollingCurtain_Resize()
        Err = 0: On Error Resume Next
        With picRollingCurtain
            vsRollingCurtain.Left = .ScaleLeft
            vsRollingCurtain.Top = cboDate.Top + cboDate.Height + 100
            vsRollingCurtain.Height = .ScaleHeight - vsRollingCurtain.Top - 50
            vsRollingCurtain.Width = .ScaleWidth
        End With
End Sub

Private Sub vsRollingCurtain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtain_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtain)
    vsRollingCurtain.Tag = "1"
End Sub
Private Sub vsRollingCurtain_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsRollingCurtain)
    vsRollingCurtain.Tag = "0"
End Sub
Private Sub vsRollingCurtain_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsRollingCurtain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsRollingCurtain, OldRow, NewRow, OldCol, NewCol)
    If OldRow >= 1 And OldRow <= vsRollingCurtain.Rows - 1 Then
        With vsRollingCurtain
            If Trim(.TextMatrix(OldRow, .ColIndex("С���տ�ʱ��"))) = "" And _
                Trim(.TextMatrix(OldRow, .ColIndex("�����տ�ʱ��"))) = "" And Trim(.TextMatrix(OldRow, .ColIndex("����ʱ��"))) = "" Then
              .Cell(flexcpBackColor, OldRow, 1, OldRow, .Cols - 1) = &H80000018
        End If
        End With
    End If
    If OldRow = NewRow Then Exit Sub
    '��ˢ�����ݣ��˳�
    If mblnNotBrush = True Then Exit Sub
    Call LoadDetail
End Sub
Private Sub vsRollingCurtain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Function LoadDetail() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϸ����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 11:17:09
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtStartDate As Date, dtEndDate As Date
    Dim lng����ID As Long, strNO As String, blnDel As Boolean
  
    With vsRollingCurtain
        If .Row >= 1 Then
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strNO = Val(.TextMatrix(.Row, .ColIndex("���ʵ���")))
            blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
        End If
    End With
    mlngRollingCurtainID = lng����ID
    mstrRollingCurtainNO = strNO
    mblnDel = blnDel
    
    On Error GoTo errHandle
    dtStartDate = CDate("1991-01-01"): dtEndDate = dtStartDate
    If lng����ID = 0 Then
        mobjChargeBill.ClearChargeAndBillTotalForm
    Else
        If mobjChargeBill.LoadChargeAndBillTotalData(Me, mlngModule, mstrPrivs, 1, lng����ID, dtStartDate, dtEndDate, True, blnDel) = False Then Exit Function
    End If
    LoadDetail = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Property Get GetChargeRollingCurtainID() As Long
    GetChargeRollingCurtainID = mlngRollingCurtainID
End Property
Public Property Get GetChargeRollingCurtainNO() As String
    GetChargeRollingCurtainNO = mstrRollingCurtainNO
End Property
Public Property Get GetChargeRollingCurtainDel() As Boolean
    GetChargeRollingCurtainDel = mblnDel
End Property
Private Function CheckCancelValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������ݵĺϷ���
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 15:44:57
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strNO As String, blnDel As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str������� As String, strWhere As String
    Dim strRollingType As String
    On Error GoTo errHandle
    With vsRollingCurtain
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Val(.TextMatrix(.Row, .ColIndex("���ʵ���")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
        str������� = "," & .Cell(flexcpData, .Row, .ColIndex("�������")) & ","
    End With
    
    If blnDel Then
        MsgBox "���ʵ���Ϊ:" & strNO & "�����ʼ�¼�Ѿ������ϣ�������������!", vbInformation + vbOKOnly, gstrSysName
        If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
        Exit Function
    End If
    
    '����Ƿ��Ѿ����տ�
    With vsRollingCurtain
        If .TextMatrix(.Row, .ColIndex("С���տ�ʱ��")) <> "" Then
            MsgBox "���ʵ���Ϊ:" & strNO & "�����ʼ�¼�Ѿ�С���տ����������!", vbInformation + vbOKOnly, gstrSysName
            If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
            Exit Function
        End If
        If .TextMatrix(.Row, .ColIndex("�����տ�ʱ��")) <> "" Then
            MsgBox "���ʵ���Ϊ:" & strNO & "�����ʼ�¼�Ѿ������տ����������!", vbInformation + vbOKOnly, gstrSysName
            If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
            Exit Function
        End If
    End With
    '����Ƿ����һ������
    If InStr(str�������, ",1,") > 0 Then
        strWhere = "�Ƿ��շ� = 1"
    End If
    If InStr(str�������, ",2,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "Ԥ����� = 1"
    End If
    If InStr(str�������, ",21,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "Ԥ����� = 2"
    End If
    If InStr(str�������, ",22,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "Ԥ����� = 3"
    End If
    If InStr(str�������, ",3,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "�Ƿ���� = 1"
    End If
    If InStr(str�������, ",4,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "�Ƿ�Һ� = 1"
    End If
    If InStr(str�������, ",5,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "�Ƿ���￨ = 1"
    End If
    If InStr(str�������, ",6,") > 0 Then
        strWhere = strWhere & IIf(strWhere = "", "", " Or ") & "�Ƿ����ѿ� = 1"
    End If
    If strWhere <> "" Then strWhere = " And (" & strWhere & ")"
    strSQL = "" & _
    "   Select Max(NO) as NO From ��Ա�սɼ�¼  " & _
    "   Where �Ǽ�ʱ��>(Select Max(�Ǽ�ʱ��) From ��Ա�սɼ�¼ where ID=[1] and ��¼����=1 ) " & _
    "               And ID+0 <> [1] and ��¼����=1  and �տ�Ա=[2]  And ����ʱ�� Is Null " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, UserInfo.����)
    
    If rsTemp.EOF = False Then
        If Nvl(rsTemp!NO) <> "" Then
        MsgBox "ע��: " & vbCrLf & _
                        "     ���ʵ���Ϊ:" & strNO & "�����ʼ�¼���������һ�ε����ʼ�¼," & vbCrLf & _
                       "Ϊ�˱�֤����������ȷ�����������һ�����ʼ�¼[" & Nvl(rsTemp!NO) & "]��ʼ����!", vbInformation + vbOKOnly, gstrSysName
        If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
        Exit Function
        End If
    End If
    CheckCancelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Public Function CancelData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ϵ�ǰ��������
    '����:���ϳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 15:44:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strNO As String, blnDel As Boolean
    Dim strDate As String, strSQL As String
    
    On Error GoTo errHandle
    With vsRollingCurtain
        If .Row < 1 Then Exit Function
        If .ColIndex("С���տ�ʱ��") < 0 _
            Or .ColIndex("���ʵ���") < 0 _
            Or .ColIndex("ID") < 0 _
            Or .ColIndex("����ʱ��") < 0 _
            Or .ColIndex("�����տ�ʱ��") < 0 Then Exit Function
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        strNO = Val(.TextMatrix(.Row, .ColIndex("���ʵ���")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
    End With
    If MsgBox("���Ƿ����Ҫ�����ʵ���Ϊ:" & strNO & "����������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    If CheckCancelValied = False Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    'Zl_�շ�Ա���ʼ�¼_Cancel
    strSQL = "Zl_�շ�Ա���ʼ�¼_Cancel("
    '  Id_In       In ��Ա�սɼ�¼.Id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ������_In   In ��Ա�սɼ�¼.������%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In In ��Ա�սɼ�¼.����ʱ��%Type
    strSQL = strSQL & "to_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    With vsRollingCurtain
        .TextMatrix(.Row, .ColIndex("������")) = UserInfo.����
        .TextMatrix(.Row, .ColIndex("����ʱ��")) = strDate
        .Cell(flexcpForeColor, .Row, 1, .Row, .Cols - 1) = vbRed
         mblnDel = True
    End With
    CancelData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
 
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б���Ϣ
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim i As Long, lngRow As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    If Val(vsRollingCurtain.Tag) = 0 Then
        '��ӡ�տƱ�ݻ���
        With vsRollingCurtain
            If .Row < 1 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("���ʵ���")) = "" Then Exit Sub
        End With
        Call mobjChargeBill.zlPrint(bytMode): Exit Sub
    End If
    '���������Ϣ
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr��λ���� & "�շ�Ա�������"
    Set objRow = New zlTabAppRow
    If lblRange.Visible Then
        objRow.Add "ʱ�䷶Χ��" & lblRange.Caption
    Else
        objRow.Add "ʱ�䷶Χ��" & Format(dtpStartDate, "yyyy-mm-dd HH:MM:SS") & "��" & Format(dtpEndDate, "yyyy-mm-dd HH:MM:SS")
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsRollingCurtain
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
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Public Sub RePrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش�ɿ���
    '����:���˺�
    '����:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String
    If Not (zlStr.IsHavePrivs(mstrPrivs, "�ɿ����ӡ") And zlStr.IsHavePrivs(mstrPrivs, "�ش�ɿ���")) Then Exit Sub
    With vsRollingCurtain
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("���ʵ���")))
    End With
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1506", Me, "NO=" & strNO, 2)
End Sub
Public Sub zlRefresh()
    '���½�������ˢ��
    Call cmdRefresh_Click
End Sub
Public Sub ShowChargeList(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ�տ�����
    '����:���˺�
    '����:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    
    With vsRollingCurtain
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("���ʵ���")))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
    End With
    Dim frmNew As frmChargeBillList
    Set frmNew = New frmChargeBillList
    Load frmNew
   Call frmNew.ShowMe(frmMain, mlngModule, mstrPrivs, 1, CStr(lng����ID), dtStartDate, dtEndDate, blnDel)
   If Not frmNew Is Nothing Then Unload frmNew
   Set frmNew = Nothing
    
End Sub
Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ��屨��
    '���:lngSys-ϵͳ��
    '        strRptCode-������
    '����:���˺�
    '����:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngDeptID As Long
    Dim lng����ID As Long, dtStartDate As Date, dtEndDate As Date, blnDel As Boolean
    Dim strNO As String
    With vsRollingCurtain
        If .Row < 1 Then Exit Sub
        strNO = Trim(.TextMatrix(.Row, .ColIndex("���ʵ���")))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        blnDel = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��"))) <> ""
    End With
    Call ReportOpen(gcnOracle, lngSys, strRptCode, frmMain, _
        "NO=" & strNO, _
        "����ID=" & lng����ID, _
        "���ϱ�־=" & IIf(blnDel, 1, 0))
End Sub
Public Sub zlDefaultSetFocus()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ����
    '����:���˺�
    '����:2013-10-16 14:23:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If cboDate.Enabled And cboDate.Visible Then
        cboDate.SetFocus
    ElseIf dtpStartDate.Enabled And dtpStartDate.Visible Then
        dtpStartDate.SetFocus
    ElseIf vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then
        vsRollingCurtain.SetFocus
    End If
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtain, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
