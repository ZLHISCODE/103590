VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExaminePathTable 
   AutoRedraw      =   -1  'True
   Caption         =   "�ٴ�·�����"
   ClientHeight    =   9495
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   14550
   Icon            =   "frmExaminePathTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14550
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   6000
      ScaleHeight     =   2385
      ScaleWidth      =   4680
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4680
      Begin VSFlex8Ctl.VSFlexGrid vsgIllness 
         Height          =   855
         Left            =   480
         TabIndex        =   9
         Top             =   1440
         Width           =   5535
         _cx             =   9763
         _cy             =   1508
         Appearance      =   0
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         FillStyle       =   1
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
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ÿ��ң�"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   13
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������������������������������������������������������"
         Height          =   180
         Index           =   1
         Left            =   330
         MouseIcon       =   "frmExaminePathTable.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   780
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ӧ���֣�"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   11
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lbl˵�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵������������������������������������������������������������������"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   6210
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   360
      ScaleHeight     =   6015
      ScaleWidth      =   4215
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
      Begin VB.PictureBox picDetail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1245
         Left            =   600
         ScaleHeight     =   1245
         ScaleWidth      =   2895
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3840
         Width           =   2895
         Begin XtremeReportControl.ReportControl rptLog 
            Height          =   1095
            Left            =   960
            TabIndex        =   7
            Top             =   480
            Width           =   3015
            _Version        =   589884
            _ExtentX        =   5318
            _ExtentY        =   1931
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.PictureBox picList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   2085
         Left            =   120
         ScaleHeight     =   2085
         ScaleWidth      =   3615
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   3615
         Begin XtremeReportControl.ReportControl rptPath 
            Height          =   1095
            Left            =   -120
            TabIndex        =   5
            Top             =   120
            Width           =   3015
            _Version        =   589884
            _ExtentX        =   5318
            _ExtentY        =   1931
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPath 
         Height          =   675
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1191
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2250
      ScaleHeight     =   600
      ScaleWidth      =   660
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   660
   End
   Begin MSComctlLib.ImageList ilsPic 
      Left            =   1245
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExaminePathTable.frx":06DC
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExaminePathTable.frx":0C76
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExaminePathTable.frx":1210
            Key             =   "branch"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExaminePathTable.frx":7A72
            Key             =   "Merge"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   9135
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmExaminePathTable.frx":E2D4
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22754
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
   Begin XtremeSuiteControls.TabControl tbcContent 
      Height          =   675
      Left            =   6840
      TabIndex        =   14
      Top             =   6360
      Width           =   1935
      _Version        =   589884
      _ExtentX        =   3413
      _ExtentY        =   1191
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   480
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmExaminePathTable.frx":EB66
      Left            =   1320
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmExaminePathTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mfrmDesign As frmPathDesign
Attribute mfrmDesign.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmPathDesign
Attribute mfrmContent.VB_VarHelpID = -1
Private WithEvents mfrmEdit As frmPathEdit
Attribute mfrmEdit.VB_VarHelpID = -1
Private mstrPrivs As String
Private mlngModul As Long
Private mlng·��ID As Long
Private mlng�汾�� As Long
Private mint���״̬ As Integer  '���״̬��0-�༭;1-�ύ���;2-ҩ�������ͨ��;3-ҩ���ƾܾ�ͨ����4-ҽ���ͨ����5-ҽ��ƾܾ�ͨ��

Private Enum E_STATUS
    E_�༭ = 0
    E_�ύ = 1
    E_ҩ��ͨ�� = 2
    E_ҩ���ܾ� = 3
    E_ͨ�� = 4
    E_�ܾ� = 5
End Enum

Private Enum COL_LIST
    COL_ID = 0
    COL_ͼ�� = 1
    COL_��֧ = 2
    COL_�к� = 3
    COL_���� = 4
    COL_���� = 5
    COL_���� = 6
    COL_���״̬
    COL_��������
    COL_���ò���
    COL_�����Ա�
    COL_��������
    COL_˵��
    COL_ͨ��
    COL_���°汾
    COL_�������
    COL_����     '1=�ϲ�·�� 0=��Ҫ·��
    COL_�汾��
End Enum

Private Enum COL_LIST_LOG
    LOG_���� = 0
    LOG_����˵��
    LOG_������Ա
    LOG_����ʱ��
End Enum


Private Enum CHK_INDEX
    CHK_�Ѿ���� = 0
    CHK_δ��� = 1
    CHK_��ͣ�� = 2
End Enum

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim lng·��ID As Long
    Dim blnTmp As Boolean
    Dim str���� As String
    Dim str���� As String
    Dim frmSub As Form
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
    Case conMenu_Edit_Audit     '���\ҽ������
        Call FuncVersionAudit(1)
    Case conMenu_Edit_Untread 'ȡ�� ���\ҽ������
        Call FuncVersionAudit(2)
    Case conMenu_Edit_Preferences 'ȫ·����Ŀ
        Call frmPathItemAll.ShowMe(Me, mstrPrivs, mlng·��ID, mlng�汾��, True)
    Case conMenu_View_Refresh    'ˢ��
        Call RefreshData
    Case conMenu_Edit_Modify  '�鿴Ŀ¼
        Call FuncPathView
    Case conMenu_File_Exit    '�˳�
        Unload Me
    End Select
End Sub

Public Sub ShowMe(frmParent As Object, ByVal lngMode As Long, ByVal strPrivs As String)
    mstrPrivs = strPrivs
    mlngModul = lngMode
    gbln˫��� = zlDatabase.GetPara("˫���ģʽ", glngSys, p�ٴ�·������) = 1
    '��˫���ģʽ
        gbln˫��� = False
        mstrPrivs = ";����;ȫԺ·��;���;"
    '�����ģʽ
'        gbln˫��� = True
'        mstrPrivs = ";����;ȫԺ·��;���;ҩ�������;"
    Me.Show 1, frmParent
End Sub

Private Sub FuncPathView()
    mfrmEdit.ShowEdit Me, mstrPrivs, mlng·��ID, , True
End Sub

Private Sub FuncVersionAudit(ByVal bytFunc As Byte)
'���ܣ����/ȡ����˵�ǰ�汾
'����
'   bytFunc 1=���;2=ȡ�����
'   1=ҽ������ -1=ҽ���ȡ����� 2=ҩ������� -2=ҩ����ȡ�����

    Dim strSql As String
    Dim intNum As Integer
    If bytFunc = 1 Then
        If gbln˫��� Then
            If mint���״̬ = E_�ύ Then
                intNum = 2
            ElseIf mint���״̬ = E_ҩ��ͨ�� Then
                intNum = 1
            End If
        Else
            intNum = 1
        End If
    ElseIf bytFunc = 2 Then
        If gbln˫��� Then
            If mint���״̬ = E_ͨ�� Or mint���״̬ = E_�ܾ� Then
                intNum = 5
            ElseIf mint���״̬ = E_ҩ��ͨ�� Or mint���״̬ = E_ҩ���ܾ� Then
                intNum = 6
            End If
        Else
            intNum = 5
        End If
    End If
    If intNum = 5 Or intNum = 6 Then
        If MsgBox("ȷʵҪȡ����˵�ǰ�汾���ٴ�·����", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        strSql = "Zl_�ٴ�·�����_Insert(" & intNum & "," & mlng·��ID & "," & mlng�汾�� & ",NULL,NULL," & IIf(gbln˫���, 1, 0) & ")"
        On Error GoTo errH
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        On Error GoTo 0
    ElseIf intNum = 1 Or intNum = 2 Then
        If frmPathAduit.ShowAudit(Me, mlng·��ID, mlng�汾��, intNum) = False Then Exit Sub
    End If
    
    Call RefreshData
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.staThis.Visible Then Bottom = Me.staThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    picLeft.Move lngLeft, lngTop, , lngBottom - lngTop
    picDetail.Move lngLeft, lngTop, , lngBottom - lngTop
    With Me.picInfo
        .Left = picLeft.Left + picLeft.Width + 60
        .Top = lngTop
        .Width = lngRight - .Left
        .Height = lngBottom - lngTop
    End With
    Call ResizeInfoPane
    With Me.tbcContent
        .Left = picInfo.Left
        .Top = picInfo.Top + picInfo.Height
        .Width = picInfo.Width
        .Height = lngBottom - .Top
    End With
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean

    Select Case Control.ID
    Case conMenu_Edit_Audit
        If (InStr(";" & mstrPrivs & ";", ";���;") = 0 And InStr(";" & mstrPrivs & ";", ";ҩ�������;") = 0 And gbln˫���) Or _
            (gbln˫��� = False And InStr(";" & mstrPrivs & ";", ";���;") = 0) Then
            Control.Visible = False
        Else
            blnEnabled = mlng·��ID <> 0 And mlng�汾�� > 0 And (mint���״̬ = E_ҩ��ͨ�� Or mint���״̬ = E_�ύ) And tbcPath.Selected.Tag = "�����"
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_Untread
        If (InStr(";" & mstrPrivs & ";", ";���;") = 0 And InStr(";" & mstrPrivs & ";", ";ҩ�������;") = 0 And gbln˫���) Or _
            (gbln˫��� = False And InStr(";" & mstrPrivs & ";", ";���;") = 0) Then
            Control.Visible = False
        Else
            blnEnabled = mlng·��ID <> 0 And mlng�汾�� > 0 And InStr(",2,3,4,5,", mint���״̬) > 0 And InStr(",���ͨ��,���δ��,", tbcPath.Selected.Tag) > 0
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_Modify
        Control.Enabled = mlng·��ID <> 0 And mlng�汾�� > 0
    Case conMenu_Edit_Preferences 'ȫ·����Ŀ
        Control.Enabled = mlng·��ID <> 0 And mlng�汾�� > 0
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picLeft.Hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picDetail.Hwnd
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As XtremeDockingPane.Pane
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    gbln˫��� = zlDatabase.GetPara("˫���ģʽ", glngSys, p�ٴ�·������) = 1
    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, False)

    Set mfrmEdit = New frmPathEdit
    Set mfrmDesign = New frmPathDesign
    Set mfrmContent = New frmPathDesign

    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    '�������
    '---------------------------------------------------
    Call InitReportColumnLog
    
    Call MainDefCommandBar
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 400, 300, DockLeftOf, Nothing)
    objPane.Title = "·���б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set objPane = Me.dkpMain.CreatePane(2, 400, 200, DockBottomOf, objPane)
    objPane.Title = "����б�"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
     
    'tbcPath ·���б�
    With Me.tbcPath
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "�����", picList.Hwnd, 0).Tag = "�����"
        .InsertItem(1, "���ͨ��", picList.Hwnd, 0).Tag = "���ͨ��"
        .InsertItem(2, "���δ��", picList.Hwnd, 0).Tag = "���δ��"
        
        .Item(2).Selected = True
        .Item(0).Selected = True
        '��λ·��ѡ�
        .Item(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(tbcPath), "tbcPath", 0)).Selected = True
    End With
        
    'TabControl
    '-----------------------------------------------------
    With Me.tbcContent
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        .InsertItem 0, "�ٴ�·����", mfrmContent.Hwnd, 0
    End With
 
    '��Ӧ����
    '---------------------------------------------------------
    Call InitVsgIllness
    
    Call RestoreWinState(Me, App.ProductName)
    Call RefreshData
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    Dim lngCount As Long
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True    '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ�����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "·����Ϣ")
        objControl.IconId = 3022
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Preferences, "ȫ·����Ŀ")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        objControl.BeginGroup = True
    End With
    
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh    'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help    '����
    End With

End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPath
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)��ItemIndex������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(COL_ID, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_ͼ��, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_��֧, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_�к�, "�к�", 35, True)
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_����, "����", 80, True)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "����", 50, True)
        objCol.Groupable = False
        Set objCol = .Columns.Add(COL_����, "����", 150, True)
        objCol.Groupable = False
        Set objCol = .Columns.Add(COL_���״̬, "��˽���", 120, True)
        Set objCol = .Columns.Add(COL_��������, "��������", 65, True)
        Set objCol = .Columns.Add(COL_���ò���, "���ò���", 55, True)
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_�����Ա�, "�����Ա�", 55, True)
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_��������, "��������", 55, True)
        Set objCol = .Columns.Add(COL_˵��, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_ͨ��, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_���°汾, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_�������, "�������", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_�汾��, "", 0, False)
        objCol.Visible = False

        For Each objCol In .Columns
            objCol.Editable = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ���ٴ�·��..."
            '.ShadeGroupHeadings = True
        End With

        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False    '������SelectionChanged�¼�
        .SetImageList Me.ilsPic
                
        If gbln˫��� And InStr(mstrPrivs, ";ҩ�������;") > 0 And InStr(mstrPrivs, ";���;") > 0 Then
            .GroupsOrder.Add .Columns(COL_�������)
            .GroupsOrder(0).SortAscending = True    '����֮��,��������в���ʾ,�����е������ǲ����
            .GroupsOrder.Add .Columns(COL_����)
            .GroupsOrder(0).SortAscending = True
            '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
            .SortOrder.Add .Columns(COL_�������)
            .SortOrder(0).SortAscending = True
            .SortOrder.Add .Columns(COL_����)
            .SortOrder(1).SortAscending = True
            .SortOrder.Add .Columns(COL_����)
            .SortOrder(2).SortAscending = True
        Else
            .GroupsOrder.Add .Columns(COL_����)
            .GroupsOrder(0).SortAscending = True    '����֮��,��������в���ʾ,�����е������ǲ����
    
            '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
            .SortOrder.Add .Columns(COL_����)
            .SortOrder(0).SortAscending = True
            .SortOrder.Add .Columns(COL_����)
            .SortOrder(1).SortAscending = True
        End If
    End With
End Sub

Private Sub InitReportColumnLog()
    Dim objCol As ReportColumn

    With rptLog
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)��ItemIndex������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(LOG_����, "��˽��", 100, True)
        Set objCol = .Columns.Add(LOG_����˵��, "���˵��", 200, True)
        Set objCol = .Columns.Add(LOG_������Ա, "�����", 80, True)
        Set objCol = .Columns.Add(LOG_����ʱ��, "���ʱ��", 140, True)

        For Each objCol In .Columns
            objCol.Editable = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ���������..."
        End With

        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False    '������SelectionChanged�¼�
        .SetImageList Me.ilsPic
    
        .SortOrder.Add .Columns(LOG_����ʱ��)
        .SortOrder(0).SortAscending = False
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Unload mfrmDesign
    Set mfrmDesign = Nothing
    
    Unload mfrmContent
    Set mfrmContent = Nothing
    
    Unload mfrmEdit
    Set mfrmEdit = Nothing
End Sub


Private Sub lbl����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 1 And lbl����(Index).Caption <> "" Then
        Me.lbl����(1).Font.Underline = True
        Me.lbl����(1).ForeColor = RGB(0, 0, 128)
    End If
End Sub

Private Sub mfrmDesign_DataChanged(ByVal ·��ID As Long)
'ˢ��·������Ϣ
    Call mfrmContent.zlRefresh(·��ID, mstrPrivs, lbl����(1).Caption, vsgIllness.Tag)
End Sub

Private Sub mfrmEdit_AfterSave(ByVal ���� As String, ByVal ���� As String)
    Call RefreshData(����, ����)
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl����(1).Font.Underline = False
    Me.lbl����(1).ForeColor = lbl����(0).ForeColor
    vsgIllness.FontUnderline = False
    vsgIllness.ForeColor = lbl����(0).ForeColor
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next

    lbl˵��.Width = picInfo.ScaleWidth - lbl˵��.Left * 2
    lbl����(1).Width = picInfo.ScaleWidth - lbl����(1).Left - lbl˵��.Left
    vsgIllness.Left = lbl����(0).Left
    vsgIllness.Width = picInfo.ScaleWidth - vsgIllness.Left - lbl˵��.Left
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With tbcPath
        .Top = 0
        .Left = 0
        .Height = picLeft.ScaleHeight
        .Width = picLeft.ScaleWidth
    End With
    picList.Width = picLeft.ScaleWidth
    rptLog.Move 0, 0, picLeft.ScaleWidth
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    rptLog.Left = 0
    rptLog.Top = 0
    rptLog.Width = picDetail.ScaleWidth
    rptLog.Height = picDetail.ScaleHeight
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    rptPath.Left = 0
    rptPath.Top = 0
    rptPath.Width = picList.ScaleWidth
    rptPath.Height = picList.ScaleHeight
End Sub

Private Sub rptPath_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow
    
    If KeyCode = vbKeyReturn And Shift = 0 Then
        If rptPath.SelectedRows.count > 0 Then
            If Not rptPath.SelectedRows(0).GroupRow Then
                Set objRow = rptPath.SelectedRows(0)
            End If
        End If
        If Not objRow Is Nothing Then
            Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
            If Not objControl Is Nothing Then objControl.Execute
        End If
    End If
End Sub

Private Sub rptPath_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objControl As CommandBarControl
    
    If Not Row.GroupRow Then
        Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
End Sub

Private Sub rptPath_SelectionChanged()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim arrStr As Variant
    Dim intRowNum As Integer: Dim intColNum As Integer
    Dim i As Long

    On Error GoTo errH

    If rptPath.SelectedRows.count = 0 Then
        Call ClearSubData
    ElseIf rptPath.SelectedRows(0).GroupRow Then
        Call ClearSubData
    Else
        With rptPath.SelectedRows(0)
            mlng·��ID = Val(.Record(COL_ID).Value)
            mlng�汾�� = Val(.Record(COL_�汾��).Value)
            mint���״̬ = Val(.Record(COL_���״̬).Value)
            
            '��Ӧ�������
            Call LoadAduit(mlng·��ID, mlng�汾��)
            
            lbl˵��.Caption = "˵����" & .Record(COL_˵��).Value
            
            '��Ӧ������Ϣ
            If .Record(COL_ͨ��).Value = 1 Then
                lbl����(1).Caption = "���ٴ�·������������סԺ�ٴ�����"
            Else
                strSql = "Select B.����,B.���� From �ٴ�·������ A,���ű� B Where A.����ID=B.ID And A.·��ID=[1] Order by B.����"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_ID).Value))
                strTmp = ""
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "," & rsTmp!���� & "-" & rsTmp!����
                    rsTmp.MoveNext
                Loop
                If strTmp <> "" Then
                    lbl����(1).Caption = Mid(strTmp, 2)
                Else
                    lbl����(1).Caption = "<���ٴ�·����δ���������õĿ���>"
                End If
            End If

            '��Ӧ������Ϣ
            vsgIllness.Clear
        
            strSql = "Select Decode(B.����,NULL,'['||C.����||']'||C.����,'['||B.����||']'||B.����) as ���� ,B.���� " & _
                     " From �ٴ�·������ A,��������Ŀ¼ B,�������Ŀ¼ C" & _
                     " Where A.����ID=B.ID(+) And A.���ID=C.ID(+) And A.·��ID=[1] and a.����=0" & _
                     " Order by B.����,C.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_ID).Value))
            strTmp = ""
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!����
                rsTmp.MoveNext
            Loop
            If strTmp <> "" Then
                With vsgIllness
                    arrStr = Split(Mid(strTmp, 2), ",")
                    .Cols = 3: .Rows = ((UBound(arrStr) + 1) + (.Cols - 1)) \ .Cols
                    .Tag = Mid(strTmp, 2)
                    For i = 0 To UBound(arrStr)
                        intRowNum = i \ .Cols
                        intColNum = i Mod .Cols
                        .TextMatrix(intRowNum, intColNum) = arrStr(i)
                    Next i
                End With
            Else
                vsgIllness.Rows = 1: vsgIllness.Cols = 1
                vsgIllness.TextMatrix(0, 0) = "<���ٴ�·����δ��������Ӧ�Ĳ���>"
            End If
            
            '·������Ϣ
            Call mfrmContent.zlRefresh(mlng·��ID, mstrPrivs, lbl����(1).Caption, vsgIllness.Tag, 2, mlng�汾��)
        End With
        
        Call Form_Resize
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ResizeInfoPane()
'���ܣ����ݵ�ǰ��Ϣ���ݣ�������Ϣ�������Ϣ����ߴ��λ��
'˵��������Label��AutoSize�����Զ�������ǩ�߶�
    lbl����(0).Top = lbl˵��.Top + lbl˵��.Height + Screen.TwipsPerPixelY * 6
    lbl����(1).Top = lbl����(0).Top + lbl����(0).Height + Screen.TwipsPerPixelY * 3
    lbl����(0).Top = lbl����(1).Top + lbl����(1).Height + Screen.TwipsPerPixelY * 6

    vsgIllness.Top = lbl����(0).Top + lbl����(0).Height + Screen.TwipsPerPixelY * 3
    '���ݶ�Ӧ����������̬��ʾ��Ӧ������Ϣ�������ʾ5��
    vsgIllness.Height = vsgIllness.RowHeightMin * IIf(vsgIllness.Rows > 5, 5, vsgIllness.Rows)
    vsgIllness.ColWidthMin = vsgIllness.Width / vsgIllness.Cols
    picInfo.Height = vsgIllness.Top + vsgIllness.Height + Screen.TwipsPerPixelY * 6
End Sub

Private Function RefreshData(Optional ByVal str���� As String, Optional ByVal str���� As String) As Boolean
'���ܣ����ݵ�ǰ���õ�������ȡ�ٴ�·��Ŀ¼����
'���������ڶ�λ
    Dim rsTmp       As ADODB.Recordset
    Dim strSql      As String
    Dim strFilter   As String
    
    Dim objRecord   As ReportRecord
    Dim objItem     As ReportRecordItem
    Dim objRow As ReportRow, i As Long
    Dim lngPreID As Long, lngPreIdx As Long
    Dim intTypeNum  As Integer
    Dim lngPathColor As Long                'δ���·��Ŀ¼ǰ����ɫֵ
    Dim lngStopColor As Long                '��ֹͣ·��Ŀ¼ǰ����ɫֵ
    Dim intType As Integer                 '���״̬
    Dim strTmp As String
    Dim strPrivs As String                  '11-1-ҩ�������;1-ҽ������
    Screen.MousePointer = 11
    
    On Error GoTo errH
    strPrivs = IIf(InStr(mstrPrivs, ";ҩ�������;") > 0, "1", "0") & IIf(InStr(mstrPrivs, ";���;") > 0, "1", "0")
    
    Select Case tbcPath.Selected.Tag
    Case "�����"
        If gbln˫��� Then
            If "11" = strPrivs Then
                strFilter = " And Instr(',1,2,',','||a.���״̬||',') > 0"
            ElseIf "10" = strPrivs Then
                strFilter = " And NVL(a.���״̬,0)=1"
            ElseIf "01" = strPrivs Then
                strFilter = " And NVL(a.���״̬,0)=2"
            Else
                strFilter = " And NVL(a.���״̬,0)<0"
            End If
        Else
            '����˫���ģʽ�л�����˫���ģʽ,���ݴ������״̬Ϊ2������
            strFilter = " And Instr(',1,2,',','||a.���״̬||',') > 0"
        End If
    Case "���ͨ��"
        '������ʷ����:���״̬Ϊ��,���ʱ�䲻Ϊ��
        If gbln˫��� Then
            If "11" = strPrivs Then
                strFilter = " And (Instr(',2,4,',','||a.���״̬||',') > 0 Or C.���ʱ�� Is not NULL)  "
            ElseIf "10" = strPrivs Then
                strFilter = " And NVL(a.���״̬,0)=2"
            ElseIf "01" = strPrivs Then
                strFilter = " And (NVL(a.���״̬,0)=4 Or C.���ʱ�� Is not NULL) "
            Else
                strFilter = " And NVL(a.���״̬,0)<0"
            End If
        Else
            If InStr(mstrPrivs, ";���;") > 0 Then
                strFilter = " And (NVL(a.���״̬,0)=4 OR C.���ʱ�� Is Not NULL) "
            Else
                strFilter = " And NVL(a.���״̬,0)<0"
            End If
        End If
    Case "���δ��"
        If gbln˫��� Then
            If "11" = strPrivs Then
                strFilter = " And Instr(',3,5,',','||a.���״̬||',') > 0"
            ElseIf "10" = strPrivs Then
                strFilter = " And NVL(a.���״̬,0)=3"
            ElseIf "01" = strPrivs Then
                strFilter = " And NVL(a.���״̬,0)=5"
            Else
                strFilter = " And NVL(a.���״̬,0)<0"
            End If
        Else
            '����˫���ģʽ�л�����˫���ģʽ,���ݴ������״̬Ϊ3������
            If InStr(mstrPrivs, ";���;") > 0 Then
                strFilter = " And Instr(',3,5,',','||a.���״̬||',') > 0"
            Else
                strFilter = " And NVL(a.���״̬,0)<0"
            End If
        End If
    End Select

    'SQL�в��������Ч��,ReportControl��������
    strSql = "Select Distinct a.Id, a.����, a.����, a.����, a.��������, a.���ò���, a.�����Ա�, a.��������, a.˵��, a.ͨ��, a.���°汾, NVL(a.���״̬,0) as ���״̬," & vbNewLine & _
             "                Decode(b.Id, Null, Null, 1) As �Ƿ���ڷ�֧, a.����,Decode(c.���ʱ��, Null, 0, 1) As �����,c.�汾�� " & vbNewLine & _
             "From �ٴ�·��Ŀ¼ A, �ٴ�·����֧ B, �ٴ�·���汾 C" & vbNewLine & _
             "Where a.Id = b.·��id(+) And a.Id = c.·��id(+) And  C.�汾��=(Select max(�汾��) from �ٴ�·���汾 D where A.id=d.·��ID(+)) And C.ͣ��ʱ�� Is NULL "
    strSql = strSql & strFilter
    If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
        'û��Ȩ��ʱ��ֻ�ܶ�ֻӦ���ڱ��Ƶ�·�����д���
        strSql = strSql & _
                 " And A.ͨ�� = 2 And Exists" & vbNewLine & _
                 "      (Select 1 From ������Ա C,�ٴ�·������ D  " & vbNewLine & _
                 "       Where C.��Աid = [1] And D.����id = C.����id And ·��id = A.ID)"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    '��¼����ѡ�еķ���
    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lngPreIdx = rptPath.SelectedRows(0).Index    '���ڿ������¶�λ
            lngPreID = rptPath.SelectedRows(0).Record(COL_ID).Value
        End If
    End If
    
    rptPath.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptPath.Records.Add()
        Set objItem = objRecord.AddItem(Val(rsTmp!ID))
        Set objItem = objRecord.AddItem("")
        If Val(rsTmp!���� & "") = 1 Then
            objItem.Icon = ilsPic.ListImages("Merge").Index - 1
        Else
            objItem.Icon = ilsPic.ListImages("Path").Index - 1
        End If
        Set objItem = objRecord.AddItem("")
        If NVL(rsTmp!�Ƿ���ڷ�֧, 0) <> 0 Then
            objItem.Icon = ilsPic.ListImages("branch").Index - 1
        End If
        Set objItem = objRecord.AddItem("")
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����, "<δָ������>")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        intType = Val(rsTmp!���״̬ & "")
        If tbcPath.Selected.Tag = "���ͨ��" Then
            If Val(rsTmp!�����) = 1 And intType = 0 Then '��ʷ���ݴ���
                intType = 4
            End If
        End If
        Set objItem = objRecord.AddItem(intType)
        Select Case tbcPath.Selected.Tag
        Case "�����"
            If gbln˫��� Then
                strTmp = IIf(intType = 1, "��ҩ�������", "��ҽ������")
            Else
                strTmp = "�����"
            End If
        Case "���ͨ��"
            If gbln˫��� Then
                strTmp = IIf(intType = 2, "ҩ�������ͨ��", "ҽ������ͨ��")
            Else
                strTmp = "���ͨ��"
            End If
        Case "���δ��"
            If gbln˫��� Then
                strTmp = IIf(intType = 3, "ҩ�������δ��", "ҽ������δ��")
            Else
                strTmp = "���δ��"
            End If
        End Select
        objItem.Caption = strTmp

        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!��������)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!���ò���)))
        Set objItem = objRecord.AddItem(CStr(Decode(NVL(rsTmp!�����Ա�, 0), 0, "", 1, "��", 2, "Ů")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!��������)))
        Set objItem = objRecord.AddItem(CStr("" & rsTmp!˵��))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!ͨ��, 1)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!���°汾, 0)))
        
        If gbln˫��� And InStr(mstrPrivs, ";ҩ�������;") > 0 And InStr(mstrPrivs, ";���;") > 0 Then
            If tbcPath.Selected.Tag = "�����" Then
                Set objItem = objRecord.AddItem(IIf(intType = 1, "ҩ����", "ҽ���"))
            ElseIf tbcPath.Selected.Tag = "���ͨ��" Then '2,4
                Set objItem = objRecord.AddItem(IIf(intType = 2, "ҩ����", "ҽ���"))
            ElseIf tbcPath.Selected.Tag = "���δ��" Then '3,5
                Set objItem = objRecord.AddItem(IIf(intType = 3, "ҩ����", "ҽ���"))
            End If
        Else
            Set objItem = objRecord.AddItem("")
        End If
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!����, 0)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!�汾��)))
        lngPathColor = IIf(Val(rsTmp!�����) = 1, vbBlack, &HFF&)
        For i = COL_�к� To COL_ͨ��
            If lngStopColor <> 0 Then
                objRecord.Item(i).ForeColor = lngStopColor
            ElseIf lngPathColor <> 0 Then
                objRecord.Item(i).ForeColor = lngPathColor
            End If
        Next
        rsTmp.MoveNext
    Loop

    rptPath.Populate

    '�����ж��ʱ����ʾ�к���
    If rptPath.Rows.count - rptPath.Records.count > 1 Then
        rptPath.Columns(COL_�к�).Visible = True
        rptPath.Columns(COL_�к�).SortAscending = True
    Else
        rptPath.Columns(COL_�к�).Visible = False
    End If

    '�кŸ�ֵ
    For i = 0 To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If .GroupRow Then intTypeNum = intTypeNum + 1
            If Not .GroupRow Then
                .Record(COL_�к�).Value = i - intTypeNum + 1
            End If
        End With
    Next

    If rptPath.Rows.count = 0 Then
        Call ClearSubData
    Else
        If str���� <> "" And str���� <> "" Then
            For i = 0 To rptPath.Rows.count - 1
                If Not rptPath.Rows(i).GroupRow Then
                    If rptPath.Rows(i).Record(COL_����).Value = str���� _
                       And rptPath.Rows(i).Record(COL_����).Value = str���� Then
                        Set objRow = rptPath.Rows(i): Exit For
                    End If
                End If
            Next
        Else
            If lngPreID <> 0 Then
                '�ȿ��ٶ�λ
                If lngPreIdx <= rptPath.Rows.count - 1 Then
                    If Not rptPath.Rows(lngPreIdx).GroupRow Then
                        If rptPath.Rows(lngPreIdx).Record(COL_ID).Value = lngPreID Then
                            Set objRow = rptPath.Rows(lngPreIdx)
                        End If
                    End If
                End If
                '�ٽ��в���
                If objRow Is Nothing Then
                    For i = 0 To rptPath.Rows.count - 1
                        If Not rptPath.Rows(i).GroupRow Then
                            If rptPath.Rows(i).Record(COL_ID).Value = lngPreID Then
                                Set objRow = rptPath.Rows(i): Exit For
                            End If
                        End If
                    Next
                End If
            End If
            'ȡ��һ���Ƿ�����
            If objRow Is Nothing Then
                For i = 0 To rptPath.Rows.count - 1
                    If Not rptPath.Rows(i).GroupRow Then Set objRow = rptPath.Rows(i): Exit For
                Next
            End If
        End If

        Set rptPath.FocusedRow = objRow    '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Me.staThis.Panels(2).Text = "���� " & rptPath.Records.count & " ���ٴ�·��"
    End If

    Screen.MousePointer = 0
    RefreshData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadAduit(ByVal lngPathID As Long, ByVal lngVersion As Long)

    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim objRecord   As ReportRecord
    Dim objItem     As ReportRecordItem
    If lngPathID = 0 Then
        rptLog.Records.DeleteAll
        rptLog.Populate
        Exit Sub
    End If
    If gbln˫��� Then
        strSql = "Select Decode(����״̬, 1, 'ҽ������ͨ��', 2, 'ҽ������δ��', 3, 'ҩ�������ͨ��', 4, 'ҩ�������δ��') As ����״̬, NVL(����˵��,'δ��д') As ����˵��, ������Ա, ����ʱ��" & vbNewLine & _
            "From �ٴ�·�����" & vbNewLine & _
            "Where ·��id = [1] And �汾�� = [2]" & vbNewLine & _
            "Order By ����ʱ�� Desc"
    Else
        strSql = "Select Decode(����״̬, 1, '���ͨ��', 2, '���δ��', 3, 'ҩ�������ͨ��', 4, 'ҩ�������δ��') As ����״̬, NVL(����˵��,'δ��д') As ����˵��, ������Ա, ����ʱ��" & vbNewLine & _
            "From �ٴ�·�����" & vbNewLine & _
            "Where ·��id = [1] And �汾�� = [2]" & vbNewLine & _
            "Order By ����ʱ�� Desc"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, lngVersion)
    
    rptLog.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptLog.Records.Add()
        Set objItem = objRecord.AddItem(rsTmp!����״̬ & "")
        Set objItem = objRecord.AddItem(rsTmp!����˵�� & "")
        Set objItem = objRecord.AddItem(rsTmp!������Ա & "")
        Set objItem = objRecord.AddItem(Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS"))
        rsTmp.MoveNext
    Loop

    rptLog.Populate
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearSubData()
    Dim i As Integer

    lbl˵��.Caption = "˵����"

    lbl����(1).Caption = ""

    vsgIllness.Rows = 0
    vsgIllness.Rows = 5

    Me.staThis.Panels(2).Text = ""
    mlng·��ID = 0
    mlng�汾�� = 0
    Call mfrmContent.zlRefresh(0, mstrPrivs, lbl����(1).Caption, vsgIllness.Tag, 2, 0)
    Call LoadAduit(0, 0)
    Call Form_Resize
    Call picList_Resize
End Sub

Private Sub tbcPath_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible Then
        Call RefreshData
    End If
End Sub

Private Sub InitVsgIllness()
'����:��ʼ����Ӧ����

    With vsgIllness
        .Cols = 3
        .Rows = 5
        .FixedCols = 0
        .FixedRows = 0
        .AllowSelection = False
        .BackColorBkg = vbWhite
        .RowHeightMin = 300
        .Appearance = flexXPThemes
        .BorderStyle = flexBorderNone
        .ScrollBars = flexScrollBarVertical
        .GridLines = flexGridNone
        .ColWidthMin = .Width / .Cols
    End With
End Sub

Private Sub vsgIllness_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsgIllness.FontUnderline = True
    vsgIllness.ForeColor = RGB(0, 0, 128)
    vsgIllness.ToolTipText = vsgIllness.Text
End Sub

