VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmEPRModelSaveAs 
   Caption         =   "���Ϊ����..."
   ClientHeight    =   7110
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   10755
   Icon            =   "frmEPRModelSaveAs.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10755
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   5865
      Index           =   2
      Left            =   6150
      ScaleHeight     =   5865
      ScaleWidth      =   3900
      TabIndex        =   3
      Top             =   480
      Width           =   3900
      Begin VB.Frame fra 
         Height          =   5145
         Index           =   0
         Left            =   195
         TabIndex        =   8
         Top             =   420
         Width           =   3480
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   18
            Top             =   1470
            Width           =   2940
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   17
            Top             =   435
            Width           =   2940
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   840
            TabIndex        =   16
            Top             =   780
            Width           =   2940
         End
         Begin VB.OptionButton opt��Χ 
            Caption         =   "&1)ȫԺͨ��"
            Height          =   180
            Index           =   0
            Left            =   825
            TabIndex        =   15
            Top             =   1905
            Width           =   1215
         End
         Begin VB.OptionButton opt��Χ 
            Caption         =   "&2)����ͨ��"
            Height          =   180
            Index           =   1
            Left            =   825
            TabIndex        =   14
            Top             =   2205
            Width           =   1215
         End
         Begin VB.OptionButton opt��Χ 
            Caption         =   "&3)����ʹ��"
            Height          =   180
            Index           =   2
            Left            =   825
            TabIndex        =   13
            Top             =   2505
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   4305
            Width           =   2940
         End
         Begin VB.TextBox txt 
            Height          =   735
            Index           =   3
            Left            =   840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   2775
            Width           =   2940
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   840
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1125
            Width           =   2940
         End
         Begin VB.CheckBox chkAdd 
            Caption         =   "����(&A)"
            Enabled         =   0   'False
            Height          =   240
            Left            =   855
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   150
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����(&F)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   26
            Top             =   1515
            Width           =   630
         End
         Begin VB.Label lbl��� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���(&B)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   495
            Width           =   630
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����(&N)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   825
            Width           =   630
         End
         Begin VB.Label lbl��Χ 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ��(&U)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   1920
            Width           =   630
         End
         Begin VB.Label lbl��Ա 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   840
            TabIndex        =   22
            Top             =   4650
            Width           =   2940
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����(&R)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   105
            TabIndex        =   21
            Top             =   4350
            Width           =   630
         End
         Begin VB.Label lbl˵�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "˵��(&M)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   105
            TabIndex        =   20
            Top             =   2820
            Width           =   630
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����(&D)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   1185
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   1
      Left            =   750
      ScaleHeight     =   1935
      ScaleWidth      =   3630
      TabIndex        =   2
      Top             =   3390
      Width           =   3630
      Begin VB.Frame fra 
         Height          =   1905
         Index           =   2
         Left            =   255
         TabIndex        =   6
         Top             =   -15
         Width           =   6585
         Begin VSFlex8Ctl.VSFlexGrid vfgTerm 
            Height          =   750
            Left            =   285
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   180
            Width           =   4080
            _cx             =   7197
            _cy             =   1323
            Appearance      =   2
            BorderStyle     =   0
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16777215
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   2
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
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
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3075
      Index           =   0
      Left            =   675
      ScaleHeight     =   3075
      ScaleWidth      =   3630
      TabIndex        =   1
      Top             =   105
      Width           =   3630
      Begin VB.Frame fra 
         Height          =   3885
         Index           =   1
         Left            =   315
         TabIndex        =   4
         Top             =   180
         Width           =   6795
         Begin XtremeReportControl.ReportControl rptList 
            Height          =   2970
            Left            =   75
            TabIndex        =   5
            Top             =   150
            Width           =   4515
            _Version        =   589884
            _ExtentX        =   7964
            _ExtentY        =   5239
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
         End
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":10BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":1458
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":1D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":20CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6735
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRModelSaveAs.frx":2466
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13917
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRModelSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������
'######################################################################################################################

Private Enum mLCol
    ͼ�� = 0: ����: ID: ����ID: ��ԱID: ����: ���: ����: ����: ˵��: ����: ��Ա
End Enum

Private mbytFromTab As Byte     '��Դ������:1-��������Ŀ¼,2-���Ӳ�����¼
Private mlngFromId As Long      '��Դ��¼id
Private mbytPower As Integer     '�û�Ȩ�޼���

Private mlngFileId As Long      '�ļ�ID
Private mlngDemoId As Long      '����ʾ��id
Private mblnOK As Boolean

Private mlngSelfId As Long      '��ǰ�û�����Աid
Private mstrSelfName As String  '��ǰ�û�����Ա����

'��ʱ����
Private lngCount As Long

Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mstrPrivs As String
Private rptCol As ReportColumn
Private rptRcd As ReportRecord
Private rptItem As ReportRecordItem
Private rptRow As ReportRow
Public Event SaveModels(ByRef lngDemoId As Long, ByRef blnOK As Boolean)


'�������Զ�����̻���
'######################################################################################################################
Public Function ShowMe(ByVal bytFromTab As Byte, ByVal lngFromId As Long) As Long
    '******************************************************************************************************************
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '������ bytFromTab-��Դ������,1-��������Ŀ¼,2-���Ӳ�����¼
    '       lngFromId-��Դ�ļ�¼id
    '       strCompends-�������id����δ����ʱ��ʾ���Ϊ���ģ��������Ƭ��
    '���أ�ȷ�������������޸ĵ�ID��ȡ������0
    '******************************************************************************************************************
    mbytFromTab = bytFromTab
    mlngFromId = lngFromId
    mlngDemoId = 0
    If ExecuteCommand("��ʼ�ؼ�") = False Then Unload Me: Exit Function
    If ExecuteCommand("��ע���") = False Then Unload Me: Exit Function
    If ExecuteCommand("��ʼ����") = False Then Unload Me: Exit Function

    Call ExecuteCommand("ˢ������")
    
    DataChanged = False
    
    'Ĭ��Ϊ����
    Call chkAdd_Click
        
    '��ʾ����
    
    Me.Show vbModal
    
    If mblnOK Then
        ShowMe = mlngDemoId
    Else
        ShowMe = 0
    End If

End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim lngTmp As Long
    Dim bytӦ�÷�Χ As Byte
    
    On Error GoTo errHand

    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        Call InitCommandBar
        
        '���
        With rptList
                
            Set rptCol = .Columns.Add(mLCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
            Set rptCol = .Columns.Add(mLCol.����, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
            Set rptCol = .Columns.Add(mLCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
            Set rptCol = .Columns.Add(mLCol.����ID, "����id", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
            Set rptCol = .Columns.Add(mLCol.��ԱID, "��Աid", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
            Set rptCol = .Columns.Add(mLCol.����, "����", 0, False): rptCol.Editable = False: rptCol.Groupable = False:  rptCol.Visible = False
            Set rptCol = .Columns.Add(mLCol.���, "���", 49, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
            Set rptCol = .Columns.Add(mLCol.����, "����", 100, True): rptCol.Editable = False: rptCol.Groupable = False
            Set rptCol = .Columns.Add(mLCol.����, "����", 60, False): rptCol.Editable = False: rptCol.Groupable = False
            Set rptCol = .Columns.Add(mLCol.˵��, "˵��", 200, True): rptCol.Editable = False: rptCol.Groupable = False
            Set rptCol = .Columns.Add(mLCol.����, "����", 70, True): rptCol.Editable = False: rptCol.Groupable = True
            Set rptCol = .Columns.Add(mLCol.��Ա, "������", 50, False): rptCol.Editable = False: rptCol.Groupable = True
            
            .SetImageList Me.imgList
            .AllowColumnRemove = False
            .MultipleSelection = False
            .ShowItemsInGroups = False
            With .PaintManager
                .ColumnStyle = xtpColumnFlat
                .GridLineColor = RGB(225, 225, 225)
                .NoGroupByText = "�϶��б��⵽����,�����з���..."
                .NoItemsText = "û�п���ʾ����Ŀ..."
                .VerticalGridStyle = xtpGridSolid
            End With
            .GroupsOrder.DeleteAll
            .GroupsOrder.Add .Columns.Find(mLCol.����)
            .GroupsOrder(0).SortAscending = True
            .SortOrder.Add .Columns.Find(mLCol.���)
        End With
        
        txt(0).MaxLength = GetMaxLength("��������Ŀ¼", "���")
        txt(1).MaxLength = GetMaxLength("��������Ŀ¼", "����")
        txt(2).MaxLength = GetMaxLength("��������Ŀ¼", "����")
        txt(3).MaxLength = GetMaxLength("��������Ŀ¼", "˵��")
        
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        mstrPrivs = GetPrivFunc(glngSys, 1070)
        
        If InStr(1, mstrPrivs, "ȫԺ��������") <> 0 Then
            mbytPower = 0
        ElseIf InStr(1, mstrPrivs, "���Ҳ�������") <> 0 Then
            mbytPower = 1
            opt��Χ(0).Enabled = False
        ElseIf InStr(1, mstrPrivs, "���˲�������") <> 0 Then
            mbytPower = 2
            opt��Χ(0).Enabled = False
            opt��Χ(1).Enabled = False
        Else
            mbytPower = -1
            MsgBox "�Բ����㲻�߱����ı༭Ȩ�ޣ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        'ѡ���ѡ���ȱʡ
        If opt��Χ(2).Enabled Then
            opt��Χ(2).Value = 1
        Else
            If opt��Χ(1).Enabled Then
                opt��Χ(2).Value = 1
            ElseIf opt��Χ(0).Enabled Then
                opt��Χ(0).Value = 1
            End If
        End If
    
        Me.Caption = "���Ϊ����..."
    
        '��ȡ�ļ�����id
        If mbytFromTab = 1 Then
            gstrSQL = "Select f.Id, f.���, f.���� From �����ļ��б� f, ��������Ŀ¼ s Where f.Id = s.�ļ�id And s.Id = [1]"
        Else
            gstrSQL = "Select f.Id, f.���, f.���� From �����ļ��б� f, ���Ӳ�����¼ s Where f.Id = s.�ļ�id And s.Id = [1]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFromId)
        If rsTemp.RecordCount <= 0 Then
            MsgBox "�ļ���δ������Ӧ�ļ����嶪ʧ��������淶�ģ�", vbExclamation, gstrSysName
            Exit Function
        End If
        
        Me.Caption = "����Ϊ��" & rsTemp!��� & "-" & rsTemp!���� & "���ķ���:"
        mlngFileId = rsTemp!ID
        
        '����������Ϣ
        gstrSQL = "Select Distinct D.ID, D.����, D.����, R.ȱʡ, R.��Աid, P.����" & vbNewLine & _
                "From ���ű� D, ������Ա R, ��Ա�� P, �ϻ���Ա�� U, ��������˵�� C," & vbNewLine & _
                "     (Select ����, ͨ�� From �����ļ��б� Where ID = [1]) L" & vbNewLine & _
                "Where D.ID = R.����id And R.��Աid = P.ID And P.ID = U.��Աid And U.�û��� = User And D.ID = C.����id And" & vbNewLine & _
                "      C.�������� In ('�ٴ�', '���', '����', '����', '����', '����', 'Ӫ��', '���') And" & vbNewLine & _
                "      (Nvl(L.ͨ��, 0) <> 2 Or L.���� = 7 Or" & vbNewLine & _
                "      L.���� <> 7 And L.ͨ�� = 2 And D.ID In (Select ����id From ����Ӧ�ÿ��� Where �ļ�id = [1]))" & vbNewLine & _
                "Order By R.ȱʡ Desc, D.����"
                
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
        With rsTemp
            If .RecordCount <= 0 Then
                MsgBox "��Ŀǰ�����ڸò���Ӧ�ÿ��ҷ�Χ�����ܹ����ģ�", vbExclamation, gstrSysName
                Exit Function
            End If
            Do While Not .EOF
                cbo(1).AddItem !���� & "-" & !����
                cbo(1).ItemData(cbo(1).NewIndex) = !ID
                If !ȱʡ = 1 Then cbo(1).ListIndex = cbo(1).NewIndex
                mlngSelfId = !��ԱID: mstrSelfName = !����
                .MoveNext
            Loop
            If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
        End With
        
        cbo(0).Clear
        cbo(0).AddItem ""
        gstrSQL = "Select a.���� From ��������Ŀ¼ a Where a.�ļ�id=[1] And a.���� Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
        If rsTemp.BOF = False Then
            Do While Not rsTemp.EOF
                cbo(0).AddItem rsTemp("����").Value
                rsTemp.MoveNext
            Loop
        End If
        cbo(0).ListIndex = 0

    '--------------------------------------------------------------------------------------------------------------
    Case "ˢ������"

        Call ExecuteCommand("��ȡ����")
        Call ExecuteCommand("��ȡ����")
            
    '--------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
        
        Dim objItem As ReportRecordItem
        
        gstrSQL = "Select L.ID, L.���, L.����, L.����, Nvl(L.����, 'δ����') As ����, L.����, L.˵��, L.ͨ�ü�," & vbNewLine & _
                    "       L.����id, L.��Աid, D.���� As ����, P.���� As ��Ա,Decode(L.����, Null, 1, 2) As ����" & vbNewLine & _
                    "From ��������Ŀ¼ L, ���ű� D, ��Ա�� P" & vbNewLine & _
                    "Where L.����id = D.ID And L.��Աid = P.ID And L.�ļ�id =[1] And Nvl(L.����, 0) =[2] And L.ͨ�ü� >=[3]" & vbNewLine & _
                    Decode(mbytPower, 0, "", 1, " And ����ID In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User)", 2, " And ��ԱID=[5] ") & vbNewLine & _
                    "Order By Decode(L.����, Null, 1, 2), L.����, L.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId, CInt(2), mbytPower, UserInfo.����ID, UserInfo.ID)
        
        rptList.Records.DeleteAll
        Do While Not rsTemp.EOF
            Set rptRcd = rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CInt(IIf(IsNull(rsTemp!ͨ�ü�), 0, rsTemp!ͨ�ü�))): rptItem.Icon = rptItem.Value
            Set rptItem = rptRcd.AddItem(CInt(Val("" & rsTemp!����))): rptItem.Icon = IIf(rptItem.Value = 0, 4, 5)
            rptRcd.AddItem CStr(rsTemp!ID)
            rptRcd.AddItem zlCommFun.Nvl(rsTemp!����ID, 0)
            rptRcd.AddItem zlCommFun.Nvl(rsTemp!��ԱID, 0)
            Set objItem = rptRcd.AddItem(Val(rsTemp!����) & CStr(rsTemp!����))
            objItem.Caption = CStr(rsTemp!����)
            rptRcd.AddItem CStr(rsTemp!���)
            rptRcd.AddItem CStr(rsTemp!����)
            rptRcd.AddItem CStr("" & rsTemp!����)
            rptRcd.AddItem CStr("" & rsTemp!˵��)
            rptRcd.AddItem CStr("" & rsTemp!����)
            rptRcd.AddItem CStr("" & rsTemp!��Ա)
            rsTemp.MoveNext
        Loop
        rptList.Populate
        
        If rptList.Rows.Count > 0 Then
            For Each rptRow In Me.rptList.Rows
                If Not (rptRow.Record Is Nothing) Then
                    If mlngDemoId = rptRow.Record(mLCol.ID).Value Then Set rptList.FocusedRow = rptRow: Exit For
                End If
            Next
            If rptList.FocusedRow Is Nothing Then Set rptList.FocusedRow = rptList.Rows(0)
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
    
        With vfgTerm
            .Clear
            .Rows = .FixedRows
            Set .Cell(flexcpPicture, .FixedRows - 1, 0) = imgList.ListImages(4).Picture
                        
            If Not (rptList.FocusedRow Is Nothing) Then
                If Not (rptList.FocusedRow.Record Is Nothing) Then
                    lngTmp = rptList.FocusedRow.Record.Item(mLCol.ID).Value
                End If
            End If
                        
            gstrSQL = "Select ���� As ������, ���� As ����ֵ" & vbNewLine & _
                    "From Table(Cast(f_Segment_������([1]) As " & gstrDbOwner & ".t_Dic_Rowset))" & vbNewLine & _
                    "Where ���� Is Not Null"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngTmp)
            
            If rsTemp.RecordCount <= 0 Then
                .TextMatrix(.FixedRows - 1, 0) = "��ʹ������������"
            Else
                .TextMatrix(.FixedRows - 1, 0) = "��������������ʱ����ʹ�ã�"
            End If
            
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = Space(2) & .Rows - 1 & ")" & rsTemp!������ & "Ϊ'" & Replace(rsTemp!����ֵ, vbTab, "'��'") & "'"
                rsTemp.MoveNext
            Loop

            .AutoSize 0
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "��ȡ��ϸ"
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
            
                
                '���ж��Ƿ���Ȩ���޸�
                bytӦ�÷�Χ = rptList.FocusedRow.Record.Item(mLCol.ͼ��).Value
                
                If opt��Χ(bytӦ�÷�Χ).Enabled = False Then
                    MsgBox "�Բ����㲻�ܸ��ġ�" & rptList.FocusedRow.Record.Item(mLCol.����).Value & "�����ģ�", vbInformation, gstrSysName
                    Exit Function
                End If
                
                lngTmp = rptList.FocusedRow.Record.Item(mLCol.ID).Value
                
                chkAdd.Tag = lngTmp
                
                chkAdd.Value = vbUnchecked
                chkAdd.Enabled = True
                
                txt(0).Text = rptList.FocusedRow.Record.Item(mLCol.���).Value
                txt(1).Text = rptList.FocusedRow.Record.Item(mLCol.����).Value
                txt(2).Text = rptList.FocusedRow.Record.Item(mLCol.����).Value
                txt(3).Text = rptList.FocusedRow.Record.Item(mLCol.˵��).Value
                cbo(0).Text = rptList.FocusedRow.Record.Item(mLCol.����).Caption
                
                opt��Χ(bytӦ�÷�Χ).Value = True
                                
                For lngCount = 0 To cbo(1).ListCount - 1
                    If cbo(1).ItemData(lngCount) = rptList.FocusedRow.Record.Item(mLCol.����ID).Value Then
                        cbo(1).ListIndex = lngCount
                        Exit For
                    End If
                Next
                
                lbl��Ա.Tag = rptList.FocusedRow.Record.Item(mLCol.��ԱID).Value
                lbl��Ա.Caption = rptList.FocusedRow.Record.Item(mLCol.��Ա).Value

            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case "ɾ������"
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
            
                strTmp = "���ɾ�����ġ�" & rptList.FocusedRow.Record.Item(mLCol.����).Value & "����"
                If MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    
                    lngTmp = rptList.FocusedRow.Record.Item(mLCol.ID).Value
                    
                    gstrSQL = "zl_��������Ŀ¼_delete('" & lngTmp & "')"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    
                    rptList.Records.RemoveAt rptList.FocusedRow.Record.Index
                    rptList.Populate
                    
                    If lngTmp = Val(chkAdd.Tag) And lngTmp > 0 Then
                        Call chkAdd_Click
                    End If
                    
                    Call ExecuteCommand("��ȡ����")
                End If
                
            End If
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case "У������"
                
        If Trim(txt(0).Text) = "" Then MsgBox "�������ţ�", vbInformation, gstrSysName: txt(0).SetFocus: Exit Function
        If Trim(txt(1).Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: txt(1).SetFocus: Exit Function
        If cbo(1).ListIndex = -1 Then MsgBox "��������ң�", vbInformation, gstrSysName: cbo(1).SetFocus: Exit Function
        
        If Val(chkAdd.Tag) > 0 Then
            
            gstrSQL = "Select * From ��������Ŀ¼ Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(chkAdd.Tag))
            If rsTemp.BOF = False Then
                
                Select Case zlCommFun.Nvl(rsTemp("ͨ�ü�").Value, 0)
                Case 0            'ȫԺͨ��
                    
                Case 1            '����ͨ��
                    '������
                    If zlCommFun.Nvl(rsTemp("����id").Value, 0) <> UserInfo.����ID Then
                        '��ֹ
                        Call MsgBox("����Ȩ�����Ѵ��ڵġ�" & zlCommFun.Nvl(rsTemp("����").Value) & "�����ģ�", vbInformation, gstrSysName)
                        
                        Exit Function
                    End If
                Case 2            '����ͨ��
                    '����
                    If zlCommFun.Nvl(rsTemp("��Աid").Value, 0) <> UserInfo.ID Then
                        '��ֹ
                        Call MsgBox("����Ȩ�����Ѵ��ڵġ�" & zlCommFun.Nvl(rsTemp("����").Value) & "�����ģ�", vbInformation, gstrSysName)
                                                
                        Exit Function
                    End If
                End Select
                
                If MsgBox("��ѡ���˸����Ѵ��ڵġ�" & zlCommFun.Nvl(rsTemp("����").Value) & "�����ģ��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            
            End If
            
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case "��������"
        Dim blnOK As Boolean
        If chkAdd.Value = vbChecked Then
            '��������/Ƭ��
            mlngDemoId = zlDatabase.GetNextId("��������Ŀ¼")
        Else
            '�޸�ԭ�еķ���/Ƭ��
            mlngDemoId = Val(chkAdd.Tag)
        End If
        
        bytӦ�÷�Χ = IIf(opt��Χ(0).Value, 0, IIf(opt��Χ(1).Value, 1, 2))
            
        gstrSQL = mlngDemoId & IIf(Me.chkAdd.Value = vbChecked, "," & mlngFileId, "")
        gstrSQL = gstrSQL & ",'" & Trim(Me.txt(0).Text) & "','" & Trim(Me.txt(1).Text) & "','" & Trim(Me.txt(2).Text) & "'"
        gstrSQL = gstrSQL & IIf(Me.chkAdd.Value = vbChecked, ",2", "") & ",'" & Replace(Trim(Me.txt(3).Text), Chr(vbKeyReturn), "") & "'"
        gstrSQL = gstrSQL & "," & bytӦ�÷�Χ

        gstrSQL = gstrSQL & "," & cbo(1).ItemData(cbo(1).ListIndex) & IIf(chkAdd.Value = vbChecked, "," & lbl��Ա.Tag, "") & ",'" & cbo(0).Text & "'"
        gstrSQL = IIf(chkAdd.Value = vbChecked, "Zl_��������Ŀ¼_Insert", "Zl_��������Ŀ¼_Update") & "(" & gstrSQL & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        RaiseEvent SaveModels(mlngDemoId, blnOK)
        
        If Not blnOK Then '�ύ���ɹ���ɾ��
            gstrSQL = "zl_��������Ŀ¼_delete('" & mlngDemoId & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End Select


    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.Options.LargeIcons = True
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        .Add 0, vbKeyF2, conMenu_Edit_Save                  '����
    End With

End Function

'����Ϊ�ؼ��¼�����
'######################################################################################################################

Private Sub cbo_Change(Index As Integer)
    
End Sub

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Save                        '����
        zlCommFun.ShowFlash "���ڱ������ݣ����Եȣ�", Me
        If ExecuteCommand("У������") And DataChanged Then
            If ExecuteCommand("��������") Then
                mblnOK = True
                Unload Me
            End If
        End If
        zlCommFun.StopFlash
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        Call ExecuteCommand("ɾ������")
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub


Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft - picPane(2).Width, lngBottom - lngTop - picPane(1).Height
    picPane(1).Move lngLeft, picPane(0).Top + picPane(0).Height, picPane(0).Width
    picPane(2).Move picPane(1).Left + picPane(1).Width, lngTop, picPane(2).Width, lngBottom - lngTop
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Save
        Control.Enabled = DataChanged
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete

        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
                Control.Enabled = rptList.FocusedRow.Record.Item(mLCol.ID).Value > 0
            Else
                Control.Enabled = False
            End If
        Else
            Control.Enabled = False
        End If
            

    End Select
End Sub

Private Sub chkAdd_Click()
    Dim rsTemp As New ADODB.Recordset
    
    If Me.chkAdd.Value <> vbChecked Then Exit Sub

    txt(0).Text = GetMax("��������Ŀ¼", "���", txt(0).MaxLength, " Where �ļ�id=" & mlngFileId)
    txt(1).Text = "�·���-" & Me.txt(0).Text
    txt(2).Text = Left(zlCommFun.SpellCode(txt(1).Text), 10)
    lbl��Ա.Tag = mlngSelfId: Me.lbl��Ա.Caption = mstrSelfName
        
    If txt(0).Visible Then txt(0).SetFocus
    
    Me.chkAdd.Enabled = False
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub opt��Χ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
       
    Select Case Index
    Case 0
        fra(1).Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        rptList.Move 15, 105, fra(1).Width - 45, fra(1).Height - 105 - 30
    Case 1
    
        fra(2).Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        vfgTerm.Move 15, 105, fra(2).Width - 45, fra(2).Height - 105 - 30
        
    Case 2
        fra(0).Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        txt(3).Move txt(3).Left, txt(3).Top, txt(3).Width, fra(0).Height - txt(3).Top - 810
        
        cbo(1).Move cbo(1).Left, txt(3).Top + txt(3).Height + 45
        lbl����.Top = cbo(1).Top + 45
        
        lbl��Ա.Move lbl��Ա.Left, cbo(1).Top + cbo(1).Height + 45
    End Select
    
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyDelete Then
        Call ExecuteCommand("ɾ������")
    End If

End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    Call ExecuteCommand("��ȡ��ϸ")
    
End Sub

Private Sub rptList_SelectionChanged()
    
    Call ExecuteCommand("��ȡ����")
    
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    DataChanged = True

End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 1, 3
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab

    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0

        Select Case Index
        Case 2
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 1
        zlCommFun.OpenIme False
        If InStr(txt(Index).Text, "'") = 0 Then txt(2).Text = zlGetSymbol(txt(Index).Text)

    Case 3
        zlCommFun.OpenIme False
        txt(Index) = Replace(Me.txt(Index).Text, Chr(vbKeyReturn), "")
    End Select

End Sub
Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    '����%����
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
