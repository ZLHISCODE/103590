VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareCardManager 
   Caption         =   "���ѿ�����"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "frmSquareCardManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picCardList 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   240
      ScaleHeight     =   2565
      ScaleWidth      =   11055
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Width           =   11055
      Begin VB.PictureBox picModify 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   4845
         ScaleHeight     =   465
         ScaleWidth      =   4650
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   -90
         Visible         =   0   'False
         Width           =   4650
         Begin VB.ComboBox cboCardType 
            Height          =   300
            Left            =   1605
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   145
            Width           =   1620
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "����޸�(&O)"
            Height          =   350
            Left            =   3330
            TabIndex        =   13
            Top             =   120
            Width           =   1230
         End
         Begin VB.CheckBox chkModify 
            Caption         =   "�޸Ŀ�����(&X)"
            Height          =   350
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Value           =   2  'Grayed
            Width           =   1425
         End
         Begin MSComCtl2.DTPicker dtpValidDate 
            Height          =   300
            Left            =   1605
            TabIndex        =   15
            Top             =   150
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   181469187
            CurrentDate     =   40156.0854282407
         End
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "����"
         Height          =   405
         Index           =   3
         Left            =   4050
         TabIndex        =   10
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "�˿�"
         Height          =   405
         Index           =   2
         Left            =   3210
         TabIndex        =   9
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "ʧЧ��"
         Height          =   405
         Index           =   1
         Left            =   2235
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "��Ч��"
         Height          =   405
         Index           =   0
         Left            =   1305
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCardList 
         Height          =   2055
         Left            =   105
         TabIndex        =   4
         Top             =   435
         Width           =   6825
         _cx             =   12039
         _cy             =   3625
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
         GridColor       =   -2147483633
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
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   24
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSquareCardManager.frx":6852
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picImgList 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   60
            Width           =   210
            Begin VB.Image imgCol 
               Height          =   195
               Left            =   0
               Picture         =   "frmSquareCardManager.frx":6B7F
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����Ϣ"
         Height          =   180
         Left            =   210
         TabIndex        =   6
         Top             =   90
         Width           =   900
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   4590
      ScaleHeight     =   1875
      ScaleWidth      =   5070
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5175
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8025
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSquareCardManager.frx":70CD
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10372
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "frmSquareCardManager.frx":7961
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   645
      Top             =   3500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareCardManager.frx":8CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareCardManager.frx":900F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   495
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSquareCardManager.frx":9363
      Left            =   1005
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSquareCardManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mblnFirst As Boolean, mstrTitle As String     '���ܱ���

Private Enum mPaneID
    Pane_Search = 1     '��������
    Pane_CardLists = 2  '���б�
    Pane_CardDetails = 3    '��ϸ�б�
End Enum
Private Enum mPgIndex
    Pg_��ֵ��¼ = 250101
    Pg_���ռ�¼ = 250102
    Pg_���Ѽ�¼ = 250103
    Pg_�������� = 250104
End Enum

Private mrs���ѿ� As ADODB.Recordset
Private Type Ty_CurrentCardType '��ǰ�������Ϣ
    lng��� As Long
    bln���� As Boolean
    bln�ϸ���� As Boolean
    bln�ض����� As Boolean
    bln������ As Boolean
    bln��������˿� As Boolean
    bln������ As Boolean
End Type
Private mTy_CurCardType As Ty_CurrentCardType

Private Enum m_CardStatus '���ѿ�״̬
    Normal = 1 '����
    Recycled = 2 '����
    Refunded = 3 '�˿�
    Stoped = 4 'ͣ��
    Invalid = 5 'ʧЧ
End Enum

Private Type Ty_CurrentCard '��ǰ����Ϣ
    blnHaveData As Boolean
    lng��ID As Long
    str���� As String
    str������ As String
    byt��״̬ As m_CardStatus
    bln������ As Boolean
    bln�����ֵ���� As Boolean
    str������ As String
    str�������� As String 'yyyy-MM-dd HH:mm:ss
    bln��ֵ�� As Boolean
    lng������� As Long
End Type
Private mTy_CurCard As Ty_CurrentCard

Private WithEvents mfrmFilter As frmSquareCardFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mfrmSquareCardCallBack As frmSquareCardCallBack
Private WithEvents mfrmSquareCardConsume As frmSquareCardConsume
Attribute mfrmSquareCardConsume.VB_VarHelpID = -1
Private WithEvents mfrmSquareCardInFull As frmSquareCardInFul
Attribute mfrmSquareCardInFull.VB_VarHelpID = -1
Private mcllSubFrm As New Collection

Private mArrFilter As Variant
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��
Private mblnPrinting As Boolean

Public Sub ShowList(ByVal lngModule As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,��ʾ��ص���Ŀ��������Ϣ
    '����:���˺�
    '����:2009-11-19 15:38:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrTitle = strTitle
    If Not zlCheckDepend Then Exit Sub            '���������Բ���
    
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        zlCommFun.ShowChildWindow Me.hWnd, frmMain
    End If
    Me.ZOrder 0
End Sub

Private Function zlCheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:���ݺϷ�,����true�����򷵻�False
    '����:���˺�
    '����:2009-11-19 15:37:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, strSQL As String
    
    On Error GoTo errHandle:
    Set rsTemp = Get���㷽ʽ("���ѿ�", "1,2,8")
    If rsTemp.EOF Then
        ShowMsgbox "���ѿ�����û�п��õĽ��㷽ʽ�����ȵ����㷽ʽ���������á�"
        Exit Function
    End If
    
    strSQL = "Select ����,����, ȱʡ���, ȱʡ�ۿ�, ȱʡ��־ From ���ѿ�����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "   û��������ص����ѿ����ͣ�������[�ֵ����]�����ã�"
        Exit Function
    End If
    Do While Not rsTemp.EOF
        cboCardType.AddItem NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
        rsTemp.MoveNext
    Loop
    If cboCardType.ListCount > 0 Then cboCardType.ListIndex = 0
    zlCheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-11-20 16:02:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strHead As String, varHead As Variant
    
    strHead = "��־,1,285|ID,1,0|����,1,945|������,4,780|��ǰ״̬,4,900|��ֵ��,4,600|��Ч��,4,1935|������,1,900|����ʱ��,4,1850|" & _
            "�쿨��,1,900|�쿨����,1,1600|������,1,900|����ʱ��,4,1850|�������,1,1860|��ֵ�ۿ���,7,990|���,7,720|���۽��,7,885|" & _
            "��ǰ���,7,1020|������,4,900|ͣ����,1,900|ͣ��ʱ��,4,1850|��ע,1,900|�������,1,0"
    varHead = Split(strHead, "|")
    With vsCardList
        .Cols = UBound(varHead) + 1
        For i = 0 To UBound(varHead)
            .TextMatrix(0, i) = Split(varHead(i), ",")(0)
            .ColKey(i) = Split(varHead(i), ",")(0)
            If .TextMatrix(0, i) = "��־" Then .TextMatrix(0, i) = ""
            .ColAlignment(i) = Split(varHead(i), ",")(1)
            .ColWidth(i) = Split(varHead(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("��־")) = "-1|1"
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("����")) = "1|0"
        .ColData(.ColIndex("��ǰ���")) = "1|0"
        .ColData(.ColIndex("�������")) = "-1|1"
    End With
End Sub

Private Function InitPage() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-11-19 15:15:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As TabControlItem
    
    Err = 0: On Error GoTo errHand:
    Set mfrmSquareCardInFull = New frmSquareCardInFul
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_��ֵ��¼, "��ֵ��Ϣ", mfrmSquareCardInFull.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_��ֵ��¼
    mcllSubFrm.Add mfrmSquareCardInFull, CStr(objItem.Tag)
    
    Set mfrmSquareCardCallBack = New frmSquareCardCallBack
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_���ռ�¼, "������Ϣ", mfrmSquareCardCallBack.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_���ռ�¼
    mcllSubFrm.Add mfrmSquareCardCallBack, CStr(objItem.Tag)


    Set mfrmSquareCardConsume = New frmSquareCardConsume
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_���Ѽ�¼, "������Ϣ", mfrmSquareCardConsume.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_���Ѽ�¼
    mcllSubFrm.Add mfrmSquareCardConsume, CStr(objItem.Tag)

     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    InitPage = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitPanel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-11-18 16:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHand:
    Set mfrmFilter = New frmSquareCardFilter
    Call mfrmFilter.Init����(mlngModule, mstrPrivs)
    mcllSubFrm.Add mfrmFilter, CStr(mPgIndex.Pg_��������)

    With dkpMan
        Set objPane = .CreatePane(mPaneID.Pane_Search, 260, 400, DockLeftOf, Nothing)
        objPane.Title = "��������"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable
        objPane.MinTrackSize.Width = 260: objPane.MaxTrackSize.Width = 260
        
        Set objPane = .CreatePane(mPaneID.Pane_CardLists, 400, 400, DockRightOf, objPane)
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        
        Set objPane = .CreatePane(mPaneID.Pane_CardDetails, 400, 400, DockBottomOf, objPane)
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        
        .SetCommandBars cbsThis
        .ImageList = imlPaneIcons
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    InitPanel = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '���:
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-18 16:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim intIndex As Integer
      
    Err = 0: On Error GoTo errHand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.id = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "�շ�����(&M)"): cbrControl.BeginGroup = True
        cbrControl.IconId = 227
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "��ӡ�ɿ(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.id = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        '83399:���ϴ�,2015/7/19,���ѿ��������
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���ӿ���(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸Ŀ���(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "����(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardModify, "�޸�(&M)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Cardtrade, "����(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardFill, "����(&F)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "�˿�(&B)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelBack, "ȡ���˿�(&K)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "����(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelCallBack, "ȡ������(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardStop, "ͣ��(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardResume, "����(&F)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "��ֵ(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "��ֵ����(&T)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBackMoney, "����˿�(&E)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord, "�޸�����(&G)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord_Force, "ǿ���޸�����(&O)")
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.id = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.id = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        
        .Add FCONTROL, Asc("S"), conMenu_Edit_CardPay
        .Add FCONTROL, Asc("M"), conMenu_Edit_CardModify
        .Add FCONTROL, Asc("C"), conMenu_Edit_CardInFull
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F11, conMenu_File_FeeCollect
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "����"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Cardtrade, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardFill, "����")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "�˿�"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "����")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardStop, "ͣ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardResume, "����")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "��ֵ"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "��ֵ����")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_FeeCollect, "�շ�����"): cbrControl.BeginGroup = True
        cbrControl.IconId = 227
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
        '���ѿ����
        Set cbrControl = .Add(xtpControlComboBox, conMenu_COMBOX_INTERFACE, "�����")
        cbrControl.Flags = xtpFlagRightAlign
        cbrControl.Width = 160
        cbrControl.Style = xtpComboLabel '��ʾ�ı���ǩ��ע�����������õ�ʱ���ų�
    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.type <> xtpControlComboBox And cbrControl.type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
 
    zlDefCommandBars = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetCurrentCardType(ByVal lng�ӿڱ�� As Long)
    '����:���õ�ǰ��������Ϣ
    Dim ty_Temp As Ty_CurrentCardType
    
    On Error GoTo errHandle
    mTy_CurCardType = ty_Temp '�Զ���Type��ʼ��
    
    If lng�ӿڱ�� = 0 Then Exit Sub
    If mrs���ѿ� Is Nothing Then Exit Sub
    
    mrs���ѿ�.Filter = "���=" & lng�ӿڱ��
    If mrs���ѿ�.RecordCount = 0 Then Exit Sub
    
    With mTy_CurCardType
        .lng��� = Val(NVL(mrs���ѿ�!���))
        .bln���� = Val(NVL(mrs���ѿ�!����)) = 1
        .bln�ϸ���� = Val(NVL(mrs���ѿ�!�Ƿ��ϸ����)) = 1
        .bln�ض����� = Val(NVL(mrs���ѿ�!�Ƿ��ض�����)) = 1
        .bln������ = Val(NVL(mrs���ѿ�!�Ƿ�������)) = 1
        .bln������ = Val(NVL(mrs���ѿ�!�Ƿ�������)) = 1
        .bln��������˿� = Val(NVL(mrs���ѿ�!�Ƿ���������˿�)) = 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Function RowHidden() As Boolean
    '��ǰ���Ƿ�ɼ�
    On Error GoTo errHandle
    If vsCardList.Row < 0 Then Exit Function
    RowHidden = vsCardList.RowHidden(vsCardList.Row)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnValidCardType As Boolean, blnValidCard As Boolean
    
    If Me.Visible = False Then Exit Sub

    Err = 0: On Error Resume Next
    blnValidCardType = mTy_CurCardType.lng��� > 0 And mTy_CurCardType.bln���� '�������õĿ����
    blnValidCard = mTy_CurCard.blnHaveData And RowHidden() = False _
                And (mTy_CurCard.byt��״̬ = m_CardStatus.Normal Or mTy_CurCard.byt��״̬ = m_CardStatus.Invalid) '��Ч��
    
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = mTy_CurCard.blnHaveData
    Case conMenu_File_PrintSingleBill '��ӡ�ɿ
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "���ѿ��շ��վ�")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard
    Case conMenu_File_FeeCollect '�շ�����
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs_RollingCurtain, "����")
        Control.Enabled = Control.Visible

    Case conMenu_Edit_NewItem '�������ѿ����
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify  '�޸����ѿ����
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "�޸�")
        Control.Enabled = Control.Visible And mTy_CurCardType.lng��� > 0
    Case conMenu_Edit_Delete 'ɾ�����ѿ����
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "ɾ��")
        Control.Enabled = Control.Visible And mTy_CurCardType.lng��� > 0
    
    Case conMenu_Edit_CardPay
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible And blnValidCardType
    Case conMenu_Edit_CardModify
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "�޸Ŀ���Ϣ")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard
    Case conMenu_Edit_Cardtrade
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����") And mTy_CurCardType.bln������
        Control.Enabled = Control.Visible And blnValidCardType
    Case conMenu_Edit_CardFill
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����") And mTy_CurCardType.bln�ض����� And mTy_CurCardType.bln������
        Control.Enabled = Control.Visible And blnValidCardType
    
    Case conMenu_Edit_CardBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "�˿�")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard And Not mTy_CurCard.bln������
    Case conMenu_Edit_CardCancelBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "�˿�")
        Control.Enabled = Control.Visible And blnValidCardType _
                        And mTy_CurCard.blnHaveData And mTy_CurCard.byt��״̬ = Refunded And RowHidden() = False _
                        And Not mTy_CurCard.bln������
    
    Case conMenu_Edit_CardCallBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible And blnValidCardType
    Case conMenu_Edit_CardCancelCallBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible And blnValidCardType _
                        And mTy_CurCard.blnHaveData And mTy_CurCard.byt��״̬ = m_CardStatus.Recycled And RowHidden() = False
    
    Case conMenu_Edit_CardStop
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "��Ƭͣ��")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard
    Case conMenu_Edit_CardResume
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "��Ƭ����")
        Control.Enabled = Control.Visible And blnValidCardType _
                        And mTy_CurCard.blnHaveData And mTy_CurCard.byt��״̬ = m_CardStatus.Stoped And RowHidden() = False
    
    Case conMenu_Edit_CardInFull
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "��ֵ")
        Control.Enabled = Control.Visible And blnValidCardType
    Case conMenu_Edit_CardInFullBack
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "���˳�ֵ")
        Control.Enabled = Control.Visible And blnValidCardType And blnValidCard And mTy_CurCard.bln�����ֵ����
    Case conMenu_Edit_CardBackMoney
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����˿�") And mTy_CurCardType.bln��������˿�
        Control.Enabled = Control.Visible And blnValidCardType

    Case conMenu_View_ToolBar_Button: Control.Checked = cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text: Control.Checked = Not (cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size: Control.Checked = cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_1" _
                            And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_2"
        End If
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim lngID As Long, lng��ֵID As Long
    Dim blnValidCard As Boolean
    
    blnValidCard = mTy_CurCard.blnHaveData And RowHidden() = False _
                And (mTy_CurCard.byt��״̬ = m_CardStatus.Normal Or mTy_CurCard.byt��״̬ = m_CardStatus.Invalid) '��Ч��
    
    On Error GoTo errHand
    Select Case Control.id
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_Parameter '��������
        Call frmSquareCardParaSet.ShowParaSet(Me, mlngModule, mstrPrivs)
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_PrintSingleBill '��ӡ�ɿ
        Call PrintReBill
    Case conMenu_File_FeeCollect '�շ�����
        Call zlExecuteChargeRollingCurtain(Me)
    
    Case conMenu_Edit_NewItem '���������
        If frmSquareSendCardTypeEdit.zlEditSendCard(Me, mlngModule, mstrPrivs, gSendCardEdit.Card_����) = False Then Exit Sub
        Call LoadCardTypeData
    Case conMenu_Edit_Modify  '�޸Ŀ����
        If frmSquareSendCardTypeEdit.zlEditSendCard(Me, mlngModule, mstrPrivs, gSendCardEdit.Card_�޸�, mTy_CurCardType.lng���) = False Then Exit Sub
        Call LoadCardTypeData
    Case conMenu_Edit_Delete     'ɾ�������
        If DeleteCardType() Then Call LoadCardTypeData
    
    Case conMenu_Edit_CardPay    '����
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_����, mTy_CurCardType.lng���) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardModify    '�޸�
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_�޸�, mTy_CurCardType.lng���, mTy_CurCard.lng��ID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_Cardtrade '����
        If blnValidCard Then lngID = mTy_CurCard.lng��ID
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_����, mTy_CurCardType.lng���, lngID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardFill '����
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_����, mTy_CurCardType.lng���) = False Then Exit Sub
        Call LoadDataToRpt
    
    Case conMenu_Edit_CardBack    '�˿�
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_�˿�, mTy_CurCardType.lng���, mTy_CurCard.lng��ID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCancelBack   'ȡ���˿�
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_ȡ���˿�, mTy_CurCardType.lng���, mTy_CurCard.lng��ID) = False Then Exit Sub
        Call LoadDataToRpt
    
    Case conMenu_Edit_CardCallBack    '����
        'ͣ�õġ����յĺ��������������
        If blnValidCard Or (mTy_CurCard.blnHaveData And RowHidden() = False And mTy_CurCard.byt��״̬ <> m_CardStatus.Stoped) Then
            lngID = mTy_CurCard.lng��ID
        End If
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_����, mTy_CurCardType.lng���, lngID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCancelCallBack  'ȡ������
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_ȡ������, mTy_CurCardType.lng���, mTy_CurCard.lng��ID) = False Then Exit Sub
        Call LoadDataToRpt
    
    Case conMenu_Edit_CardResume        '��Ƭ����
        If CardStopAndStart(False) Then Call LoadDataToRpt
    Case conMenu_Edit_CardStop        '��Ƭͣ��
        If CardStopAndStart(True) Then Call LoadDataToRpt
    
    Case conMenu_Edit_CardInFull    '��ֵ
        If blnValidCard And mTy_CurCard.bln��ֵ�� Then lngID = mTy_CurCard.lng��ID
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_��ֵ, mTy_CurCardType.lng���, lngID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardInFullBack    '��ֵ����
        lng��ֵID = mfrmSquareCardInFull.zlGet��ֵID()
        If lng��ֵID = 0 Then Exit Sub
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_��ֵ����, _
            mTy_CurCardType.lng���, mTy_CurCard.lng��ID, lng��ֵID) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardBackMoney '����˿�
        If frmSquareRefundBalance.ShowMe(Me, mlngModule, mstrPrivs, mTy_CurCardType.lng���) = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_ChangePassWord    '�޸�����
        Call frmModiCardPass.zlModifyPass(Me, mlngModule, mTy_CurCardType.lng���, True)
    Case conMenu_Edit_ChangePassWord_Force  'ǿ���޸�����
        Call frmModiCardPass.zlModifyPass(Me, mlngModule, mTy_CurCardType.lng���, False)
    Case conMenu_View_Refresh 'ˢ��
        Call LoadDataToRpt
    Case conMenu_COMBOX_INTERFACE '���ѡ�����
        If Val(Control.Category) = Control.ItemData(Control.ListIndex) Then Exit Sub
        Call SetCurrentCardType(Control.ItemData(Control.ListIndex))
        Call LoadDataToRpt
        Control.Category = Control.ItemData(Control.ListIndex)
        
    Case conMenu_View_StatusBar
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In cbsThis(2).Controls
            If cbrControl.type <> xtpControlComboBox And cbrControl.type <> xtpControlLabel Then
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            End If
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
        cbsThis.RecalcLayout
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zlOpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    zlControl.ControlSetFocus vsCardList
    Call vsCardList_GotFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strShow As String, i As Long
    
    On Error GoTo errHandle
    mblnFirst = True
    mstrPrivs = gstrPrivs
    Me.Caption = mstrTitle
    
    mstrPrivs_RollingCurtain = GetPrivFunc(glngSys, 1506)
    mTy_CurCardType.lng��� = Val(zlDatabase.GetPara("�ϴνӿں�", glngSys, mlngModule, 0))
    Set mrs���ѿ� = zlGet���ѿ��ӿ�(, False)
    
    strShow = Trim(zlDatabase.GetPara("����ʾ��ʽ", glngSys, mlngModule, "1011"))
    If Len(strShow) < 4 Then strShow = strShow & "1111"
    For i = 0 To 3
        chkStatus(i).value = IIf(Val(Mid(strShow, i + 1, 1)) = 1, vbChecked, vbUnchecked)
    Next
    dtpValidDate.MinDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    dtpValidDate.value = DateAdd("d", 1, dtpValidDate.MinDate)
    chkModify.value = vbUnchecked
    
    Call InitPanel
    Call InitPage
    Call zlDefCommandBars '��ʼ�˵���������
    Call InitVsGrid
    Set mArrFilter = mfrmFilter.GetFilterCon
    Call LoadCardTypeData
    
    RestoreWinState Me, App.ProductName, mstrTitle
    '2006-04-25:���˺�,ͳһ���ӱ�������ģ��Ĺ���
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTemp As String
    
    strTemp = ""
    For i = 0 To 3
        strTemp = strTemp & IIf(chkStatus(i).value = 1, 1, 0)
    Next
    zlDatabase.SetPara "����ʾ��ʽ", strTemp, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "�ϴνӿں�", mTy_CurCardType.lng���, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlstr.IsHavePrivs(mstrPrivs, "��������")
    SaveWinState Me, App.ProductName, mstrTitle
   
    '�ر��Ӵ���
    For i = 1 To mcllSubFrm.count
        If Not mcllSubFrm(i) Is Nothing Then Unload mcllSubFrm(i)
    Next
End Sub

Private Sub chkStatus_Click(Index As Integer)
    Call SetCardRowColHide
End Sub

Private Sub chkModify_Click()
    Call SetModifyEnabled
End Sub

Private Sub cmdModify_Click()
   Call SaveBatchUpdateCardInfor
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case mPaneID.Pane_Search    '������������
        Item.Handle = mfrmFilter.hWnd
    Case mPaneID.Pane_CardLists '���б�
        Item.Handle = picCardList.hWnd
    Case mPaneID.Pane_CardDetails   '��ϸ����Ϣ
        Item.Handle = picList.hWnd
    End Select
End Sub

Private Sub SetCardRowColHide(Optional lngLocalRow As Long = -1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����е���ʾ������
    '���:lngLocalRow -ָ����(-1����ȫ����������)
    '����:
    '����:
    '����:���˺�
    '����:2009-12-22 21:05:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngRows As Long, i As Long
    Dim lngCurRow As Long
    
    Err = 0: On Error GoTo errHand:
    With vsCardList
        i = 1: lngRows = .Rows - 1
        If lngLocalRow < 0 Then
            .Redraw = flexRDNone
        Else
            i = lngLocalRow: lngRows = lngLocalRow
        End If
        
        lngCurRow = -2
        For lngRow = i To lngRows
            .RowHidden(lngRow) = False
            Select Case Val(.Cell(flexcpData, lngRow, .ColIndex("��ǰ״̬")))
            Case m_CardStatus.Normal
                If chkStatus(0).value = 0 Then .RowHidden(lngRow) = True
            Case m_CardStatus.Recycled
                If chkStatus(3).value = 0 Then .RowHidden(lngRow) = True
            Case m_CardStatus.Refunded
                If chkStatus(2).value = 0 Then .RowHidden(lngRow) = True
            Case m_CardStatus.Invalid
                If chkStatus(1).value = 0 Then .RowHidden(lngRow) = True
            End Select
            If .RowHidden(lngRow) = False Then
                If lngCurRow < .Row Then lngCurRow = lngRow
            End If
        Next
        If lngLocalRow < 0 Then
            If lngCurRow > 0 Then
                If .Row > 0 Then
                    If .RowHidden(.Row) Then .Row = lngCurRow
                Else
                    .Row = lngCurRow
                End If
            Else
                .Row = -1
            End If
            .Redraw = flexRDBuffered
        End If
    End With
    Exit Sub
errHand:
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant)
    Set mArrFilter = arrFilter
    '���¼�������
    Call LoadDataToRpt
End Sub

Private Function LoadDataToRpt() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-11-19 15:43:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strSubWhere As String, lngRow As Long, lngPre����ID As Long
    Dim rsTemp As ADODB.Recordset, strCurDate As String, strSQL As String
    
    Err = 0: On Error GoTo errHand:
    If mTy_CurCardType.lng��� <= 0 Then Exit Function
    
    If mArrFilter("����ʱ��")(0) <> "1901-01-01" And mArrFilter("����ʱ��")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And (����ʱ�� Between [1] And [2] Or ����ʱ�� Between [3] And [4])"
    ElseIf mArrFilter("����ʱ��")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And ����ʱ�� Between [1] And [2]"
    ElseIf mArrFilter("����ʱ��")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And ����ʱ�� Between [3] And [4]"
    End If
    If mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        strSubWhere = strSubWhere & " And ���� Between [5] And [6]"
    ElseIf mArrFilter("���ŷ�Χ")(0) <> "" Then
        strWhere = strWhere & " And a.����=[5]"
    ElseIf mArrFilter("���ŷ�Χ")(1) <> "" Then
        strWhere = strWhere & " And a.����=[6]"
    End If
    If strSubWhere = "" Then
        '���û�нᶨʱ�䷶Χ,��ֻ�ܲ��ҵ�ǰ���쿨�˺ͷ�����
        If mArrFilter("�쿨��") <> "" Then strWhere = strWhere & " And a.�쿨�� like [7]"
        If mArrFilter("������") <> "" Then strWhere = strWhere & " And a.������ like [8]"
    Else
        If mArrFilter("�쿨��") <> "" Then strSubWhere = strSubWhere & " And �쿨�� like [7]"
        If mArrFilter("������") <> "" Then strSubWhere = strSubWhere & " And ������ like [8]"
    End If
    
    If Trim(mArrFilter("������")) <> "����" Then strWhere = strWhere & " And a.������ = [9]"
    
    If Val(mArrFilter("����ͣ�ÿ�")) = 1 Then
        strWhere = strWhere & " And a.��ǰ״̬ <= 9"   '��Ҫ�õ�����
    Else
        strWhere = strWhere & " And (a.ͣ������ Is Null Or a.ͣ������ >= To_Date('3000-01-01', 'yyyy-mm-dd'))"   '��Ҫ�õ�����
    End If
    
    strSQL = _
    "Select a.Id, a.������, a.����, a.���, a.�������, a.�ɷ��ֵ, a.��Ч��, a.����ʱ��, a.������, a.�쿨��," & vbNewLine & _
    "       a.������, a.����ʱ��, a.��ע, a.������, a.���۽��," & vbNewLine & _
    "       a.��ֵ�ۿ���, a.���, a.ͣ����, a.ͣ������, Decode(b.����, Null, '', b.���� || '-' || b.����) As �쿨����," & vbNewLine & _
    "       Nvl((Select 1 From ���˿������¼ Where ���ѿ�id = a.Id And ��¼���� = 4 And Rownum < 2), 0) As ����," & vbNewLine & _
    "       Case" & vbNewLine & _
    "          When a.��ǰ״̬ = 1 And Nvl(a.��Ч��, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 5" & vbNewLine & _
    "          When a.��ǰ״̬ = 1 And Nvl(a.ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')) <= Sysdate Then 4" & vbNewLine & _
    "          When a.��ǰ״̬ = 4 Then 2" & vbNewLine & _
    "          When a.��ǰ״̬ = 5 Then 4" & vbNewLine & _
    "          Else Nvl(a.��ǰ״̬, 1)" & vbNewLine & _
    "       End As ��ǰ״̬, a.�������" & vbNewLine
    If strSubWhere <> "" Then
        strSubWhere = Mid(strSubWhere, 6)
        strSQL = strSQL & _
        "From ���ѿ���Ϣ A, ���ű� B, " & vbNewLine & _
        "      (Select ����, Max(���) As ��� From ���ѿ���Ϣ Where " & strSubWhere & " And �ӿڱ�� = [10] Group By ����) C" & vbNewLine & _
        "Where a.�쿨����id = b.Id(+) And a.���� = c.���� And a.��� = c.��� And a.�ӿڱ�� = [10] " & strWhere
    Else
        strSQL = strSQL & _
        "From ���ѿ���Ϣ A,���ű� B" & _
        "Where a.�쿨����id = b.Id(+) and a.�ӿڱ��=[10] " & strWhere
    End If
    strSQL = strSQL & vbNewLine & _
            "Order By ����ʱ�� Desc,����"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1)), _
        CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1)), _
        CStr(mArrFilter("���ŷ�Χ")(0)), CStr(mArrFilter("���ŷ�Χ")(1)), _
        CStr(mArrFilter("�쿨��")), CStr(mArrFilter("������")), _
        CStr(mArrFilter("������")), mTy_CurCardType.lng���)
    
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With vsCardList
        If .Row > 0 Then lngPre����ID = Val(.Cell(flexcpData, .Row, .ColIndex("����")))
        
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("ID")) = NVL(rsTemp!id)
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��ֵ��")) = IIf(Val(NVL(rsTemp!�ɷ��ֵ)) = 1, "��", "")
            .TextMatrix(lngRow, .ColIndex("��Ч��")) = Format(rsTemp!��Ч��, "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("��Ч��"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("��Ч��")) = ""
            
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
            
            .TextMatrix(lngRow, .ColIndex("�쿨��")) = NVL(rsTemp!�쿨��)
            .TextMatrix(lngRow, .ColIndex("�쿨����")) = NVL(rsTemp!�쿨����)
            
            .TextMatrix(lngRow, .ColIndex("ͣ����")) = NVL(rsTemp!ͣ����)
            .TextMatrix(lngRow, .ColIndex("ͣ��ʱ��")) = Format(rsTemp!ͣ������, "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("ͣ��ʱ��"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("ͣ��ʱ��")) = ""
            
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("����ʱ��"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("����ʱ��")) = ""
            
            .TextMatrix(lngRow, .ColIndex("�������")) = NVL(rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("��ֵ�ۿ���")) = Format(rsTemp!��ֵ�ۿ���, "###0.00;-###0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("���")) = Format(rsTemp!������, "###0.00;-###0.00;;")
            .TextMatrix(lngRow, .ColIndex("���۽��")) = Format(rsTemp!���۽��, "###0.00;-###0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("��ǰ���")) = Format(rsTemp!���, "###0.00;-###0.00;;")
            .TextMatrix(lngRow, .ColIndex("������")) = IIf(Val(NVL(rsTemp!����)) = 1, "��", "")
            .TextMatrix(lngRow, .ColIndex("��ע")) = NVL(rsTemp!��ע)
            
            .TextMatrix(lngRow, .ColIndex("��ǰ״̬")) = _
                decode(Val(NVL(rsTemp!��ǰ״̬)), _
                            m_CardStatus.Recycled, "����", _
                            m_CardStatus.Refunded, "�˿�", _
                            m_CardStatus.Invalid, "ʧЧ", _
                            m_CardStatus.Stoped, "ͣ��", "��Ч")
            .Cell(flexcpData, lngRow, .ColIndex("��ǰ״̬")) = Val(NVL(rsTemp!��ǰ״̬))
            .TextMatrix(lngRow, .ColIndex("�������")) = Val(NVL(rsTemp!�������))
            
            If lngPre����ID = Val(NVL(rsTemp!id)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            
            '������ɫ��
            Call SetGridRowForeColor(lngRow)
            SetCardRowColHide lngRow
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, True
    
    Call vsCardList_AfterRowColChange(-1, 0, vsCardList.Row, 0)
    LoadDataToRpt = True
    Exit Function
errHand:
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetGridRowForeColor(ByVal lngRow As Long)
    '����״̬����������ɫ
    Dim lngColor As Long, int״̬ As Integer
    
    With vsCardList
        Select Case Val(.Cell(flexcpData, lngRow, .ColIndex("��ǰ״̬")))
        Case m_CardStatus.Stoped
            lngColor = vbRed
        Case m_CardStatus.Invalid
            lngColor = &HFF00FF
        Case m_CardStatus.Recycled, m_CardStatus.Refunded
            lngColor = vbBlue
        Case Else
            '1-��Ч, 2-����,3-�˿�,4-ʧЧ,8-ͣ��
            lngColor = &H80000008
        End Select
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = lngColor
    End With
End Sub

Private Sub mfrmSquareCardConsume_zlDblClick(ByVal lng����ID As Long, ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    If zlstr.IsHavePrivs(mstrPrivs, "������������ϸ��") = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_INSIDE_1503_1", Me, "������ID=" & lng����ID, 1)
End Sub

Private Sub mfrmSquareCardInFull_AfterRowChange(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    mTy_CurCard.bln�����ֵ���� = mfrmSquareCardInFull.zl�������
End Sub

Private Sub mfrmSquareCardInFull_zlPopupMenus(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    '�����˵�:��ֵ���
    Call ShowPopuMenus(1)
End Sub

Private Function CardStopAndStart(ByVal blnStop As Boolean) As Boolean
    '����:��Ƭͣ�û�����
    '���:blnStop-ͣ�ÿ�Ƭ
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mTy_CurCard.lng��ID <= 0 Then Exit Function
    
    If blnStop Then 'ͣ��
        If mTy_CurCard.byt��״̬ = m_CardStatus.Stoped Then Exit Function
        If MsgBox("��ȷ��Ҫ�Կ���Ϊ:��" & mTy_CurCard.str���� & "���ļ�¼����ͣ�ò�����" & vbCrLf & _
                    "   ���ǡ�:����ͣ�ò�����ͣ�ú�Ŀ�Ƭ�����ܽ���ˢ�����ѣ�Ҳ�����ٷ���" & vbCrLf & _
                    "   ����:��������ͣ�ò���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        If mTy_CurCard.byt��״̬ <> m_CardStatus.Stoped Then Exit Function
        '����ͣ�õĿ���������
        strSQL = "Select 1 From ���ѿ���Ϣ Where ID = [1] And ��ǰ״̬ = 5"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mTy_CurCard.lng��ID)
        If rsTemp.EOF = False Then
            ShowMsgbox "��ǰ��(����Ϊ:" & mTy_CurCard.str���� & ")Ϊ����ͣ�õĿ����������ã�"
            Exit Function
        End If
        If MsgBox("��ȷ��Ҫ�Կ���Ϊ:��" & mTy_CurCard.str���� & "���ļ�¼�������ò�����" & vbCrLf & _
                    "   ���ǡ�:�������ò��������ú�Ŀ�Ƭ���ܽ���ˢ�����ѻ���պ��ٷ���" & vbCrLf & _
                    "   ����:�������������ò���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If

    'Zl_���ѿ���Ϣ_Stopandstart
    strSQL = "Zl_���ѿ���Ϣ_Stopandstart("
    '  Id_In     In ���ѿ���Ϣ.Id%Type,
    strSQL = strSQL & "" & mTy_CurCard.lng��ID & ","
    '  ͣ����_In In ���ѿ���Ϣ.ͣ����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ͣ��_In Number:=0 --ͣ��_In 0-����,1-ͣ��
    strSQL = strSQL & "" & IIf(blnStop, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    CardStopAndStart = True
    Exit Function
ErrHandler:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Function SaveBatchUpdateCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������¿�Ƭ��Ϣ
    '����:���³ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-05 12:02:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strFieldValue As String, blnDate As Boolean
    Dim strSQL As String, strIDs As String, lngRow As Long
    Dim blnTrain As Boolean, cllPro As New Collection
    Dim blnYes As Boolean, strTemp As String
    
    On Error GoTo ErrHandler
    With vsCardList
        Select Case .Col
        Case .ColIndex("��Ч��")
           strFields = "��Ч��": blnDate = True
           If IsNull(dtpValidDate.value) Then
                strFieldValue = "3000-01-01"
           Else
                strFieldValue = Format(dtpValidDate.value, "yyyy-MM-dd")
           End If
        Case .ColIndex("������")
           strFields = "������": blnDate = False
           If cboCardType.ListIndex < 0 Then
                ShowMsgbox "������δѡ����ѡ�����ͣ�"
                Exit Function
           End If
           strFieldValue = zlstr.NeedName(cboCardType.Text)
        Case Else
           Exit Function
        End Select
        ShowMsgbox "��ȷ��Ҫ�����޸ĵ�ǰ���п��ġ�" & strFields & "����ֵ��", True, blnYes
        If blnYes = False Then Exit Function
    End With
    
    strIDs = ""
    With vsCardList
        For lngRow = 1 To .Rows - 1
            strTemp = Val(.TextMatrix(lngRow, .ColIndex("ID")))
            If zlCommFun.ActualLen(strIDs & "," & strTemp) >= 3980 Then
                ' Zl_���ѿ���Ϣ_Update_Batch
                strSQL = "Zl_���ѿ���Ϣ_Update_Batch("
                '  Ids_In    Varchar2,
                strSQL = strSQL & "'" & Mid(strIDs, 2) & "',"
                '  �ֶ�_In   Varchar2,
                strSQL = strSQL & "'" & strFields & "',"
                '  �ֶ�ֵ_In Varchar2
                strSQL = strSQL & "'" & strFieldValue & "')"
                AddArray cllPro, strSQL
                strIDs = ""
            End If
            If strTemp <> 0 And .RowHidden(lngRow) = False Then
                strIDs = strIDs & "," & strTemp
            End If
        Next
    End With
    If strIDs <> "" Then
        ' Zl_���ѿ���Ϣ_Update_Batch
        strSQL = "Zl_���ѿ���Ϣ_Update_Batch("
        '  Ids_In    Varchar2,
        strSQL = strSQL & "'" & Mid(strIDs, 2) & "',"
        '  �ֶ�_In   Varchar2,
        strSQL = strSQL & "'" & strFields & "',"
        '  �ֶ�ֵ_In Varchar2
        strSQL = strSQL & "'" & strFieldValue & "')"
        AddArray cllPro, strSQL
        strIDs = ""
    End If
    If cllPro.count = 0 Then Exit Function
    
    blnTrain = True
    ExecuteProcedureArrAy cllPro, Me.Caption
    blnTrain = False
    
    SaveBatchUpdateCardInfor = True
    ShowMsgbox "�޸ĳɹ���"
    
    'ˢ������
    Call LoadDataToRpt
    chkModify.value = Unchecked
    
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadCardTypeData()
    'ˢ�����ѿ����
    Dim intIndex  As Integer
    Dim cbrControl As CommandBarComboBox
    
    On Error GoTo errHand
    Set cbrControl = cbsThis(2).Controls.Find(xtpControlComboBox, conMenu_COMBOX_INTERFACE)
    If cbrControl Is Nothing Then Exit Sub
    
    cbrControl.Clear
    cbrControl.Category = ""
    
    Set grsStatic.rs���ѿ��ӿ� = Nothing
    Set mrs���ѿ� = zlGet���ѿ��ӿ�(, False)
    
    intIndex = 1
    With mrs���ѿ�
        Do While Not .EOF
            cbrControl.AddItem NVL(!���) & "-" & NVL(!����) & IIf(NVL(!����) = 1, "", "(ͣ��)")
            cbrControl.ItemData(intIndex) = Val(NVL(!���))
            If Val(NVL(!���)) = mTy_CurCardType.lng��� Then
               cbrControl.ListIndex = intIndex
            End If
            intIndex = intIndex + 1
            .MoveNext
        Loop
    End With
    If intIndex > 1 And cbrControl.ListIndex <= 0 Then cbrControl.ListIndex = 1
    
    If cbrControl.ListIndex > 0 Then
        Call SetCurrentCardType(cbrControl.ItemData(cbrControl.ListIndex))
        cbrControl.Category = cbrControl.ItemData(cbrControl.ListIndex)
    Else
        Call SetCurrentCardType(0)
    End If
    
    Call LoadDataToRpt
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function DeleteCardType() As Boolean
    'ɾ�����ѿ����
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHand
    If mTy_CurCardType.lng��� <= 0 Then Exit Function
    
    mrs���ѿ�.Filter = "��� = " & mTy_CurCardType.lng���
    If Val(NVL(mrs���ѿ�!ϵͳ)) = 1 Then
        ShowMsgbox "ϵͳ�̶������ɾ�������飡"
        Exit Function
    End If
    
    '����Ƿ���ڷ�����¼
    strSQL = "Select 1 From ���ѿ���Ϣ Where �ӿڱ��=[1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mTy_CurCardType.lng���)
    If Not rsTemp.EOF Then
        ShowMsgbox "����Ϊ��" & NVL(mrs���ѿ�!����) & "�������ѿ����ڷ�����¼������ɾ����"
        Exit Function
    End If
    
    If MsgBox("��ȷ��Ҫɾ������Ϊ��" & NVL(mrs���ѿ�!����) & "�������ѿ������", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
       
    'Zl_���ѿ����Ŀ¼_Delete(
    strSQL = "Zl_���ѿ����Ŀ¼_Delete("
    ' ���_In In ���ѿ����Ŀ¼.���%Type
    strSQL = strSQL & "" & mTy_CurCardType.lng��� & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    DeleteCardType = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetModiyCaption()
    With vsCardList
        Select Case .Col
        Case .ColIndex("��Ч��")
            chkModify.Caption = "�޸ġ���Ч�ڡ�"
        Case .ColIndex("������")
            chkModify.Caption = "�޸ġ������͡�"
        Case Else
            chkModify.Visible = False
        End Select
    End With
End Sub

Private Sub SetModifyEnabled()
    Dim blnEnabled As Boolean
    
    blnEnabled = (chkModify.value = vbChecked)
    With vsCardList
        cboCardType.Visible = False
        dtpValidDate.Visible = False
        chkModify.Visible = True
        cmdModify.Visible = blnEnabled
        Select Case .Col
        Case .ColIndex("��Ч��")
            dtpValidDate.Visible = blnEnabled
            picModify.Visible = True
        Case .ColIndex("������")
            cboCardType.Visible = blnEnabled
            picModify.Visible = True
        Case Else
            picModify.Visible = False
        End Select
    End With
End Sub

Private Sub SetModifyDefaultValue()
    With vsCardList
        If .Row <= 0 Then Exit Sub
        
        Select Case .Col
        Case .ColIndex("��Ч��")
            If Trim(.TextMatrix(.Row, .Col)) = "" Then
                 dtpValidDate.value = Null
            Else
                If CDate(.TextMatrix(.Row, .Col)) < dtpValidDate.MinDate Then
                    dtpValidDate.value = dtpValidDate.MinDate
                Else
                    dtpValidDate.value = CDate(.TextMatrix(.Row, .Col))
                End If
            End If
        Case .ColIndex("������")
            cbo.SeekIndex cboCardType, .TextMatrix(.Row, .Col)
        End Select
    End With
End Sub

Private Sub SetCurrentCard()
    '����:���õ�ǰѡ����Ϣ
    Dim ty_Temp As Ty_CurrentCard
    
    On Error GoTo ErrHandler
    mTy_CurCard = ty_Temp '�Զ���Type��ʼ��
    
    With vsCardList
        If .Rows < 2 Then Exit Sub
        If .Row < 1 Then Exit Sub
    
        mTy_CurCard.blnHaveData = .TextMatrix(1, .ColIndex("����")) <> ""
        mTy_CurCard.lng��ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        mTy_CurCard.byt��״̬ = Val(.Cell(flexcpData, .Row, .ColIndex("��ǰ״̬")))
        mTy_CurCard.str���� = .TextMatrix(.Row, .ColIndex("����"))
        mTy_CurCard.str������ = .TextMatrix(.Row, .ColIndex("������"))
        mTy_CurCard.bln������ = .TextMatrix(.Row, .ColIndex("������")) = "��"
        mTy_CurCard.str������ = .TextMatrix(.Row, .ColIndex("������"))
        mTy_CurCard.str�������� = .TextMatrix(.Row, .ColIndex("����ʱ��"))
        mTy_CurCard.bln��ֵ�� = .TextMatrix(.Row, .ColIndex("��ֵ��")) = "��"
        mTy_CurCard.lng������� = Val(.TextMatrix(.Row, .ColIndex("�������")))
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsCardList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error GoTo ErrHandler
    If mblnPrinting Then Exit Sub
    If OldRow <> NewRow Then
        zl_VsGridRowChange vsCardList, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
        
        zlCommFun.ShowFlash "����װ������,���Ժ�..."
        '�����е���Ϣ
        Call SetCurrentCard
        
        With mTy_CurCard
            Call mfrmSquareCardCallBack.zlReLoadData(mTy_CurCardType.lng���, .lng��ID, .str������, .str����)    '���ռ�¼
            Call mfrmSquareCardInFull.zlReLoadData(mTy_CurCardType.lng���, .lng��ID, .str������, .str����)     '��ֵ��¼
            Call mfrmSquareCardConsume.zlReLoadData(mTy_CurCardType.lng���, .lng��ID, .str������, .str����)      '���Ѽ�¼
        End With
        zlCommFun.StopFlash
    End If
    
    If mTy_CurCard.blnHaveData Then
        If zlstr.IsHavePrivs(mstrPrivs, "�޸Ŀ���Ϣ") Then
            If OldCol <> NewCol Then chkModify.value = vbUnchecked
            Call SetModiyCaption
            Call SetModifyEnabled
            Call SetModifyDefaultValue
        End If
    Else
        picModify.Visible = False
    End If
    Exit Sub
ErrHandler:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsCardList_DblClick()
    On Error GoTo ErrHandler
    '˫���鿴
    If vsCardList.MouseRow <= 0 Then Exit Sub
    If mTy_CurCard.lng��ID = 0 Then Exit Sub
    frmSquareSendCard.zlShowCard Me, mlngModule, mstrPrivs, gEd_��ѯ, mTy_CurCardType.lng���, mTy_CurCard.lng��ID
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsCardList_GotFocus()
    zl_VsGridGotFocus vsCardList, gSysColor.lngGridColorSel
End Sub

Private Sub vsCardList_LostFocus()
    zl_VsGridLostFocus vsCardList, gSysColor.lngGridColorLost
End Sub

Private Sub vsCardList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlstr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub vsCardList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlstr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long, vRect As RECT
    
    vRect = zlControl.GetControlRect(picImgList.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCardList, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlstr.IsHavePrivs(mstrPrivs, "��������")
End Sub
 
Private Sub picCardList_Resize()
    Err = 0: On Error Resume Next
    With picCardList
        vsCardList.Left = .ScaleLeft
        vsCardList.Width = .ScaleWidth
        vsCardList.Height = .ScaleHeight - vsCardList.Top
        picModify.Width = .ScaleWidth - picModify.Left - 50
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picModify_Click()
    Err = 0: On Error Resume Next
    With picModify
        cmdModify.Left = .ScaleWidth - cmdModify.Width - 50
        cboCardType.Width = cmdModify.Left - cboCardType.Left
        dtpValidDate.Width = cboCardType.Width
    End With
End Sub

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:
    '����:
    '����:���˺�
    '����:2009-11-20 15:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim blnCardList As Boolean
    Dim lngPreRow As Long, lngPreCol As Long
    
    blnCardList = Me.ActiveControl Is vsCardList
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "���ѿ����"
    
    If CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "����ʱ�䣺" & CStr(mArrFilter("����ʱ��")(0)) & "��" & CStr(mArrFilter("����ʱ��")(1))
    End If
    If CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "����ʱ�䣺" & CStr(mArrFilter("����ʱ��")(0)) & "��" & CStr(mArrFilter("����ʱ��")(1))
    End If
    If objRow.count > 1 Then
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
    End If
    If mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        objRow.Add "���ŷ�Χ��" & CStr(mArrFilter("���ŷ�Χ")(0)) & "��" & CStr(mArrFilter("���ŷ�Χ")(1))
    ElseIf mArrFilter("���ŷ�Χ")(0) = "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        objRow.Add "���ţ�" & CStr(mArrFilter("���ŷ�Χ")(1))
    ElseIf mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) = "" Then
        objRow.Add "���ţ�" & CStr(mArrFilter("���ŷ�Χ")(0))
    End If
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    
    If mArrFilter("�쿨��") <> "" Then objRow.Add "�쿨�ˣ�" & mArrFilter("�쿨��")
    If mArrFilter("������") <> "" Then objRow.Add "�����ˣ�" & mArrFilter("������")
    If mArrFilter("������") <> "" Then objRow.Add "�����ͣ�" & mArrFilter("������")
    If Val(mArrFilter("����ͣ�ÿ�")) = 1 Then objRow.Add "����ͣ�ÿ�"
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    mblnPrinting = True
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsCardList
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            If .ColHidden(i) Or i = 0 Then
                .Cell(flexcpData, 0, i) = .ColWidth(i)
                .ColWidth(i) = 0
            End If
        Next
        lngPreRow = .Row: lngPreCol = .Col
    End With
    
    Err = 0: On Error GoTo errHand:
    Set objPrint.Body = vsCardList
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    '�ָ�
    With vsCardList
        For i = 0 To .Cols - 1
            If .ColHidden(i) Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
                .Cell(flexcpData, 0, i) = ""
            End If
        Next
        .GridColor = &H8000000F
        .Row = lngPreRow: .Col = lngPreCol
        .Redraw = flexRDBuffered
    End With
    mblnPrinting = False
    Exit Sub
errHand:
    mblnPrinting = False
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PrintReBill()
    '����:�ش�Ʊ��
    Dim strTemp As String
    Dim blnYes As Boolean
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If strTemp = "����" Then
        If mTy_CurCard.lng��ID <= 0 Then
            ShowMsgbox "δѡ����Ч�����ѿ���¼��": Exit Sub
        End If
    Else
        lngID = mfrmSquareCardInFull.zlGet��ֵID
        If lngID <= 0 Then
            ShowMsgbox "δѡ����Ч�ĳ�ֵ��¼��":  Exit Sub
        End If
    End If
    
    If mTy_CurCard.bln��ֵ�� = False Then
        ShowMsgbox "��ȷ��Ҫ��ӡ�ɿ��", True, blnYes
        If blnYes = False Then Exit Sub
        strTemp = "����"
    Else
        strTemp = zlCommFun.ShowMsgbox("�ɿ��ӡ", "��ѡ����Ҫ��ӡ�Ľɿ", "����(&F),��ֵ(&I),ȡ��(&C)", Me, vbDefaultButton2)
    End If
    If strTemp = "ȡ��" Or strTemp = "" Then Exit Sub
    
    If strTemp = "����" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, _
            "�������=" & mTy_CurCard.lng�������, "�ɿ�=" & 0, "�Ҳ�=" & 0, "��ֵID=0", "ReportFormat=1", 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, _
            "��ֵID=" & lngID, "�ɿ�=" & 0, "�Ҳ�=" & 0, "�������=0", "ReportFormat=2", 2)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlOpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode������
    With mTy_CurCard
        If .lng��ID = 0 Then Exit Sub
        
        Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, _
            "���ѿ�ID=" & .lng��ID, "����=" & .str����, "������=" & .str������, "������=" & .str������, "��������=" & .str��������)
    End With
End Sub

Private Sub vsCardList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�Ҽ������˵�
    If Button <> vbRightButton Then Exit Sub
    Call ShowPopuMenus(0)
End Sub

Private Function ShowPopuMenus(ByVal bytMode As Byte) As Boolean
    '��ʾ�����˵�
    '��Σ�
    '   bytMode 0-������Ϣ�б�1-��ֵ��Ϣ�б�
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrControl As CommandBarControl
    
    Err = 0: On Error Resume Next
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(2)
    If cbrMenuBar.Visible = False Then Exit Function
    
    Set cbrPopupBar = cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Select Case cbrControl.id
        Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete
            '����ʾ�˵�
        Case Else
            If bytMode = 0 Or _
                (bytMode = 1 And (cbrControl.id = conMenu_Edit_CardInFull Or cbrControl.id = conMenu_Edit_CardInFullBack)) Then
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.id, cbrControl.Caption)
                cbrPopupItem.BeginGroup = cbrControl.BeginGroup
            End If
        End Select
    Next
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(3)
    If cbrMenuBar.Visible Then
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            Select Case cbrControl.id
            Case conMenu_View_Refresh
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.id, cbrControl.Caption)
                cbrPopupItem.BeginGroup = cbrControl.BeginGroup
            End Select
        Next
    End If
    
    cbrPopupBar.ShowPopup
End Function
