VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceQuery 
   Caption         =   "���˽��ʷ��ò�ѯ"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11715
   Icon            =   "frmBalanceQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11715
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4755
      ScaleHeight     =   495
      ScaleWidth      =   510
      TabIndex        =   6
      Top             =   90
      Width           =   510
      Begin VB.Label lblCancel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   5355
      ScaleHeight     =   2295
      ScaleWidth      =   3705
      TabIndex        =   4
      Top             =   3255
      Width           =   3705
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1845
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
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
   Begin XtremeSuiteControls.TabControl tabMain 
      Height          =   1320
      Left            =   1740
      TabIndex        =   3
      Top             =   3705
      Width           =   2775
      _Version        =   589884
      _ExtentX        =   4895
      _ExtentY        =   2328
      _StockProps     =   64
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1425
      ScaleHeight     =   420
      ScaleWidth      =   5550
      TabIndex        =   1
      Top             =   2565
      Width           =   5550
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   2460
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��������: XXX"
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
         Left            =   90
         TabIndex        =   2
         Top             =   105
         Width           =   1560
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      SimpleText      =   $"frmBalanceQuery.frx":058A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBalanceQuery.frx":05D1
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13018
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1260
      Top             =   1230
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBalanceQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gViewType
    g_Ed_���ʱ� = 0
    g_Ed_��ϸ�� = 1
    g_Ed_��Ŀ��ϸ = 2
    g_Ed_����� = 3
    g_Ed_���±� = 4
    g_Ed_��Ŀ�� = 5
    g_Ed_���յ��� = 6
    g_Ed_���շ��� = 7
End Enum

Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mcbrPopupMain As CommandBar, mcbrMenuView As CommandBarPopup, mcbrRefresh As CommandBarControl
Private mcbrCmb As CommandBarComboBox
Private mBalanceType As gBalanceBill, mlng����ID As Long, mstrPrivs As String, mlngModule As Long
Private mblnDateMoved As Boolean, mViewType As gViewType
Private mlng����ID As Long
Private mstrTime As String  '���˽��ʴ���(��ʼ="",����Ϊ"1,2,3...")
Private mdtBeginDate As Date       '���˽��ʵĿ�ʼʱ��,��ʼΪ'1900-01-01'
Private mdtEndDate As Date         '���˽��ʵĽ���ʱ��,��ʼΪ'3000-01-01'
Private mstrDeptIDs As String      '���˽��ʿ���ID��(��ʼ="",����Ϊ"0,1,2,3...",0��ʾ��������IDΪ��)
Private mstrClass As String       '��������=""-���з���(��δ����),"'����','����',..."
Private mstrChargeType As String      '�շ���� '34260
Private mstrBaby As String      '�Ƿ������Ӥ������(0-���з���,1-���˷���,2������-��mbytbaby-1��Ӥ������)
Private mstrItem As String      'Ҫ����վݷ�Ŀ
Private mbytKind As Byte       '0-����ͨ����,1-��������,2-��ͨ���ú�������
Private mblnCurBalanceOwnerFee As Boolean      '��ǰ�Ƿ����ڽᡰ�Էѷ��á�
Private mclsCon As clsBalanceCon


Public Function ShowMe(ByVal frmMain As Object, BalanceType As gBalanceBill, ByRef clsCon As clsBalanceCon, ByVal lng����ID As Long, ByVal lngModule As Long, ByVal strPrivs As String, Optional ByVal ViewType As gViewType = g_Ed_���ʱ�) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʷ��ò�ѯ�ĳ������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-30 10:38:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mclsCon = clsCon
    If Not mclsCon Is Nothing Then
        With mclsCon
            mstrTime = .strTime
            mlng����ID = .lng����ID
            mdtBeginDate = .dtBeginDate
            mdtEndDate = .dtEndDate
            mstrDeptIDs = .strDeptIDs
            mstrClass = .strClass
            mstrChargeType = .strChargeType
            mstrBaby = .strBaby
            mstrItem = .strItem
            mbytKind = .bytKind
            mblnCurBalanceOwnerFee = .blnCurBalanceOwnerFee
        End With
    End If
    mBalanceType = BalanceType
    mViewType = ViewType
    mlng����ID = lng����ID
    mstrPrivs = strPrivs
    mlngModule = lngModule
    Me.Show vbModal, frmMain
    ShowMe = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:������
    '����:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim intPara As Integer
    
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    '��ʼ������
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
        Set .Font = vsfMain.Font
    End With
    
    cbsThis.EnableCustomization False
    
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.ActiveMenuBar.ModifyStyle &H400000, 0 'ȥ���˵���ǰ׺
    
    '-----------------------------------------------------
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ModifyStyle &H400000, 0
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With
    
    For Each mcbrControl In mcbrToolBar.Controls
        If mcbrControl.ID <> conMenu_Edit_UserType Then
            mcbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    zlDefCommandBars = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:������
    '����:2013-09-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long, intActive As Integer
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String
    
    Select Case tabMain.Selected.Index
        Case 0
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "���˽��ʱ�"
        Case 1
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "���˽�����ϸ��"
        Case 2
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "���˽�����Ŀ��ϸ��"
        Case 3
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "���˽��ʷ����"
        Case 4
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "���˽��ʷ��±�"
        Case 5
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "���˽��ʷ�Ŀ��"
        Case 6
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "���˽������յ��ݱ�"
        Case 7
            With vsfMain
                If .Rows = 1 Then Exit Sub
                If .Rows = 2 And .TextMatrix(1, 1) = "" Then Exit Sub
            End With
            Set vsBill = vsfMain: strTittle = GetUnitName & "���˽������շ��ñ�"
    End Select
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTittle
    
    Set objRow = New zlTabAppRow
    objRow.Add lblInfo.Caption
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    If vsBill Is Nothing Then Exit Sub
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsBill
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Then
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsBill
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
    With vsBill
        For i = 0 To .Cols - 1
           If .ColHidden(i) = True Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000C
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Print
            Call zlRptPrint(1)
        Case conMenu_File_Preview
            Call zlRptPrint(2)
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Public Sub UnloadForm()
    Unload Me
End Sub

Private Sub Form_Load()
    stbThis.Panels(3).Text = UserInfo.����
    Call zlDefCommandBars
    Call SetTabControl
    Call InitInfo
    tabMain.Item(mViewType).Selected = True
    Call LoadCardData(tabMain.Selected.Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            'ȡ����ť
            Unload Me
    End Select
End Sub

Private Sub InitInfo()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If mlng����ID = 0 Then
        lblInfo.Caption = ""
        If mclsCon Is Nothing Then Exit Sub
        If mclsCon.lng����ID <> 0 Then
            strSQL = "Select ����,�Ա�,��������,����,�����,סԺ�� From ������Ϣ Where ����ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mclsCon.lng����ID)
            If rsTmp.EOF Then Exit Sub
            lblInfo.Caption = "��������:" & NVL(rsTmp!����) & "  �Ա�:" & NVL(rsTmp!�Ա�) & "   ��������:" & NVL(rsTmp!��������) & "   ����:" & NVL(rsTmp!����) & "   �����:" & NVL(rsTmp!�����) & "    סԺ��:" & NVL(rsTmp!סԺ��)
        End If
    Else
        lblInfo.Caption = ""
        strSQL = _
            " Select a.����, a.�Ա�, a.��������, a.����, a.�����, a.סԺ��" & vbNewLine & _
            " From ������Ϣ A, ����Ԥ����¼ B" & vbNewLine & _
            " Where b.����id = [1] And b.����id = a.����id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        If rsTmp.EOF Then Exit Sub
        lblInfo.Caption = "��������:" & NVL(rsTmp!����) & "  �Ա�:" & NVL(rsTmp!�Ա�) & "   ��������:" & NVL(rsTmp!��������) & "   ����:" & NVL(rsTmp!����) & "   �����:" & NVL(rsTmp!�����) & "    סԺ��:" & NVL(rsTmp!סԺ��)
    End If
End Sub

Private Sub SetTabControl()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����TAB�ؼ�
    '����:������
    '����:2013-09-04
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With tabMain
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        Set .PaintManager.Font = vsfMain.Font
        .InsertItem 1, "���ʱ�", picMain.hWnd, 0
        .InsertItem 2, "��ϸ��", picMain.hWnd, 0
        .InsertItem 3, "��Ŀ��ϸ", picMain.hWnd, 0
        .InsertItem 4, "�����", picMain.hWnd, 0
        .InsertItem 5, "���±�", picMain.hWnd, 0
        .InsertItem 6, "��Ŀ��", picMain.hWnd, 0
        .InsertItem 7, "���յ���", picMain.hWnd, 0
        .InsertItem 8, "���շ���", picMain.hWnd, 0
        .Item(0).Selected = True
        .PaintManager.BoldSelected = True
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.StaticFrame = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    With picInfo
        .Left = Left
        .Top = Top
        .Width = Right - Left
        Line1.X2 = .Left + .Width
    End With
    With tabMain
        .Left = picInfo.Left
        .Top = picInfo.Top + picInfo.Height + 15
        .Width = Right - Left
        .Height = Bottom - .Top
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picCancel.Left = Me.Width - picCancel.Width - 300
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    With vsfMain
        .Left = 0
        .Top = 0
        .Height = picMain.Height
        .Width = picMain.Width
    End With
End Sub

Private Sub tabMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call LoadCardData(Item.Index)
End Sub

Private Function LoadCardData(ByVal intIndex As Integer) As Boolean
'���ܣ����ݵ�ǰѡ��Ĳ��˷�����Ŀ��Ƭ����ȡ�����÷����嵥
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim strInfo As String, strPre As String
    Dim strMoney As String, strTmp As String, strTmpSql As String
    Dim arrTotal() As Currency
    Dim intCol As Integer, blnZero As Boolean
    Dim strCond As String, bytType As Byte '0-����;1-סԺ;2-�����סԺ
    Dim DateBegin As Date, DateEnd As Date
    Dim strTable As String, strTimeRange As String
    Dim strOriginal As String, strTmpOriginal As String
    On Error GoTo errH

    strPre = stbThis.Panels(2).Text
    stbThis.Panels(2).Text = "���ڶ�ȡ����,���Ժ� ����"
    Screen.MousePointer = 11
    vsfMain.Redraw = False
    Me.Refresh
    
    If mBalanceType = g_Ed_������� Or mBalanceType = g_Ed_סԺ���� Then
        blnZero = zlDatabase.GetPara("���������", glngSys, 1137) = "1"
        strCond = ""
        strCond = strCond & IIf(mstrTime = "", "", " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0")
        If mdtBeginDate <> CDate("0:00:00") Then
            strTimeRange = " And " & IIf(gint����ʱ�� = 0, "A.�Ǽ�ʱ��", "A.����ʱ��") & " Between [3] And [4]"
            DateBegin = CDate(Format(mdtBeginDate, "yyyy-MM-dd 00:00:00"))
            DateEnd = CDate(Format(mdtEndDate, "yyyy-MM-dd 23:59:59"))
        End If
        strCond = strCond & IIf(mstrDeptIDs = "", "", " And Instr([5],','||A.��������ID||',')>0")
        strCond = strCond & IIf(mstrBaby = "", "", " And Instr([6],','|| Nvl(A.Ӥ����,0) ||',')>0")
        strCond = strCond & IIf(mstrItem = "", "", " And Instr([7],','''||A.�վݷ�Ŀ||''',')>0")
        
        If mbytKind = 1 Then
            strCond = strCond & " And A.�����־=4"
        Else
            If InStr(mstrPrivs, ";סԺ���ý���;") = 0 Then strCond = strCond & " And A.�����־<>2"
            If InStr(mstrPrivs, ";������ý���;") = 0 Then strCond = strCond & " And A.�����־<>1"
            If mbytKind = 0 Then strCond = strCond & " And A.�����־<>4"
        End If
        
        bytType = IIf(mBalanceType = g_Ed_�������, 0, 1)
        
        strSQL = _
        " Select NO,Mod(��¼����,10) as ��¼����, Nvl(Sum(ʵ�ս��),0) as ʵ�ս��,Nvl(Sum(���ʽ��),0) as ���ʽ��,��� " & _
        " From סԺ���ü�¼ A" & _
        " Where ��¼״̬<>0 And ���ʷ���=1 " & strCond & _
        "       And ����ID=[1]" & _
        " Group by NO,Mod(��¼����,10),��� " & _
        IIf(blnZero, "", " Having   Nvl(Sum(ʵ�ս��),0)-Nvl(Sum(���ʽ��),0)<>0 ")
        
        strSQL = _
            " Select Mod(A.��¼����,10) as ��¼����,A.����ʱ��,Max(A.�Ǽ�ʱ��) As �Ǽ�ʱ��,A.NO,Decode(a.Ӥ����, 1, '��', Null) As Ӥ��,A.�շ����,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ," & _
            "        Sum(Decode(Floor(a.��¼���� / 10),0,a.����,0)) As ����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������,Sum(a.Ӧ�ս��) As Ӧ�ս��" & _
            " From סԺ���ü�¼ A,(" & strSQL & ") B" & _
            " Where A.NO=B.NO And A.����ID Is Not Null And Mod(A.��¼����,10)=B.��¼����" & _
            "       And A.��¼״̬<>0 And A.���ʷ���=1 And A.���=B.��� " & _
            "       And A.����ID+0=[1] And Not Exists (Select 1 From סԺ���ü�¼ C, ���˽��ʼ�¼ D Where c.No = a.No And Mod(c.��¼����,10) = Mod(a.��¼����,10) And c.��� = a.��� And c.����id = d.Id And Nvl(d.����״̬, 0) = 1) " & strCond & strTimeRange & _
            "" & _
            " Group by Mod(A.��¼����,10),A.����ʱ��,A.NO,A.�շ����,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID," & _
            "       A.�վݷ�Ŀ,A.��������ID,A.���㵥λ,Nvl(A.����,1),A.��׼����,A.����Ա����,A.��������,a.Ӥ���� " & _
            " Having    " & vbNewLine & _
            "        Sum(Nvl(a.ʵ�ս��, 0)) - Sum(Nvl(a.���ʽ��, 0)) <> 0 Or (Sum(Nvl(a.ʵ�ս��, 0)) = 0 And Sum(Nvl(a.Ӧ�ս��, 0)) = 0 And Sum(Nvl(a.���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or" & vbNewLine & _
            "             (Sum(Nvl(a.ʵ�ս��, 0)) = 0 And Sum(Nvl(a.Ӧ�ս��, 0)) <> 0 And Sum(Nvl(a.���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or" & vbNewLine & _
            "             Sum(Nvl(a.���ʽ��, 0)) = 0 And Sum(Nvl(a.Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0 " & _
            " Union All " & _
            " Select Mod(A.��¼����,10) as ��¼����,A.����ʱ��,Max(A.�Ǽ�ʱ��) As �Ǽ�ʱ��,A.NO,Decode(a.Ӥ����, 1, '��', Null) As Ӥ��,A.�շ����,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ," & _
            "        Sum(Decode(Floor(a.��¼���� / 10),0,a.����,0)) As ����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������,Sum(a.Ӧ�ս��) As Ӧ�ս��" & _
            " From סԺ���ü�¼ A,(" & strSQL & ") B" & _
            " Where A.NO=B.NO And Mod(A.��¼����,10)=B.��¼����" & _
            "       And A.��¼״̬<>0 And A.���ʷ���=1 And A.���=B.���" & _
            "       And A.����ID+0=[1] And A.����ID Is Null " & strCond & strTimeRange & _
            "" & _
            " Group by Mod(A.��¼����,10),A.����ʱ��,A.NO,A.�շ����,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID," & _
            "       A.�վݷ�Ŀ,A.��������ID,A.���㵥λ,Nvl(A.����,1),A.��׼����,A.����Ա����,A.��������,a.Ӥ���� "

        If mblnDateMoved Then
            strSQL = strSQL & " Union All " & Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼")
        End If
        
        Select Case bytType
        Case 0 '����
            strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
            
            strTmpSql = _
            " Select NO,Mod(��¼����,10) as ��¼����, Nvl(Sum(ʵ�ս��),0) as ʵ�ս��,Nvl(Sum(���ʽ��),0) as ���ʽ��" & _
            " From סԺ���ü�¼ A" & _
            " Where ��¼״̬<>0 And ���ʷ���=1 And Mod(��¼����,10)=5 And ��ҳID Is Null " & strCond & strTimeRange & _
            "       And ����ID=[1]" & _
            " Group by NO,Mod(��¼����,10) " & _
            " "
            
            strTmpSql = _
            " Select Mod(A.��¼����,10) as ��¼����,A.����ʱ��,Max(A.�Ǽ�ʱ��) As �Ǽ�ʱ��,A.NO,Decode(a.Ӥ����, 1, '��', Null) As Ӥ��,A.�շ����,A.�շ�ϸĿID,A.�վݷ�Ŀ,A.��������ID,A.���㵥λ," & _
            "        Sum(Decode(Floor(a.��¼���� / 10),0,a.����,0)) As ����,Nvl(A.����,1) as ����,A.��׼����,Sum(A.ʵ�ս��) As ʵ�ս��,Sum(A.���ʽ��) As ���ʽ��,A.����Ա����,A.��������,Sum(a.Ӧ�ս��) As Ӧ�ս��" & _
            " From סԺ���ü�¼ A,(" & strTmpSql & ") B" & _
            " Where A.NO=B.NO And Mod(A.��¼����,10)=B.��¼����" & _
            "       And A.��¼״̬<>0 And A.���ʷ���=1 And Mod(A.��¼����,10)=5 And A.��ҳID Is Null " & _
            "       And A.����ID+0=[1] " & strCond & strTimeRange & _
            " " & _
            " Group by Mod(A.��¼����,10),A.����ʱ��,A.�Ǽ�ʱ��,A.NO,A.�շ����,Nvl(A.�۸񸸺�,A.���),A.�շ�ϸĿID," & _
            "       A.�վݷ�Ŀ,A.��������ID,A.���㵥λ,A.����,Nvl(A.����,1),A.��׼����,A.����Ա����,A.��������,a.Ӥ���� "
                If mblnDateMoved Then
                    strTmpSql = strTmpSql & " Union All " & Replace(strTmpSql, "סԺ���ü�¼", "HסԺ���ü�¼")
                End If
                strTmpSql = Replace(strTmpSql, " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0", "")
                strSQL = strSQL & " Union All " & strTmpSql
        Case 1 'סԺ
        Case Else
            '�����סԺ
             strSQL = strSQL & " Union All " & Replace(Replace(strSQL, "סԺ���ü�¼", "������ü�¼"), " And Instr([2],','||Nvl(A.��ҳID,0)||',')>0", "")
        End Select
            
        strTable = "(" & strSQL & ") "
        
        'δ������嵥
        Select Case intIndex
            Case 0
                strSQL = _
                "Select To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') As ʱ��, a.No, d.���� As ��Ŀ, a.�վݷ�Ŀ As ��Ŀ, a.Ӥ�� As Ӥ��, Ltrim(To_Char(Nvl(a.ʵ�ս��,0) - Nvl(a.���ʽ��,0),'999999999" & gstrDec & "')) As δ����" & vbNewLine & _
                "From (" & strTable & vbNewLine & _
                "       ) A, �շ���ĿĿ¼ D" & vbNewLine & _
                "Where d.Id = a.�շ�ϸĿid " & _
                       IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(D.��������,'δ֪')||''',')>0") & _
                       IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(D.���,'��')||''',')>0") & _
                "Order By No,��Ŀ "
                strMoney = "4,4,1,1,1,7"
            Case 1 '��ϸ�嵥
                strSQL = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ����,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')|| To_Char(A.����,'999999990.9')||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Nvl(A.Ӧ�ս��,0),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Nvl(A.ʵ�ս��,0)-Nvl(A.���ʽ��,0),'999999999" & gstrDec & "')) as δ����,A.����Ա���� as ����Ա" & _
                " FROM " & strTable & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID " & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Order by ��������,���ݺ�,��Ŀ"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 2 '����Ŀ��ϸ
                strSQL = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ��������,Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ') ���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')|| To_Char(A.����,'999999990.9')||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Nvl(A.Ӧ�ս��,0),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Nvl(A.ʵ�ս��,0)-Nvl(A.���ʽ��,0),'999999999" & gstrDec & "')) as δ����," & _
                "       Nvl(A.��������,C.��������) as ����,A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��" & _
                " FROM " & strTable & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1)
                
               strSQL = strSQL & _
                " Union All" & _
                " SELECT NULL as ��������,NULL as ���ݺ�,NULL as ��������," & _
                "       Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ')||'ZZZZZ' as ���,NULL,to_char(sum(Nvl(A.����,1)*Nvl(A.����,1)), '999999990.9')||' '||A.���㵥λ as ����,NULL as ��׼����," & _
                "       Ltrim(To_Char(Sum(Nvl(A.Ӧ�ս��,0)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����," & _
                "       NULL as ����,NULL as ����Ա,NULL as �Ǽ�ʱ��" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.�շ�ϸĿID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID(+)" & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                "              And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Group by Nvl(D.����,C.����),C.���,A.���㵥λ" & _
                " Order by ��Ŀ,���,��������,���ݺ�"
                
                strMoney = "4,4,1,1,1,1,1,7,7,7,1,1,1"
            Case 3 '������ϸ
                strSQL = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ����,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||To_Char(A.����,'999999990.9')||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Nvl(A.Ӧ�ս��,0),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Nvl(A.ʵ�ս��,0)-Nvl(A.���ʽ��,0),'999999999" & gstrDec & "')) as δ����,A.����Ա���� as ����Ա " & _
                " FROM " & strTable & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID " & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Union All" & _
                " SELECT NULL as ��������,NULL as ���ݺ�,NULL as ����,NULL as ��Ŀ,Null as ���,A.�վݷ�Ŀ||'ZZZZZ' as ��Ŀ," & _
                "        NULL as ����,NULL as ��׼����," & _
                "        Ltrim(To_Char(Sum(Nvl(A.Ӧ�ս��,0)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����,NULL as ����Ա" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by A.�վݷ�Ŀ||'ZZZZZ'" & _
                " Order by ��Ŀ,��������,���ݺ�"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 4 '�����嵥
                strSQL = _
                " SELECT B.�ڼ�,A.�վݷ�Ŀ as ��Ŀ," & _
                "        Ltrim(To_Char(Sum(Nvl(A.Ӧ�ս��,0)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����" & _
                "        FROM " & strTable & " A,�ڼ�� B,�շ���ĿĿ¼ C" & _
                " Where A.�Ǽ�ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                "       And A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by B.�ڼ�,A.�վݷ�Ŀ" & _
                " Union All" & _
                " SELECT B.�ڼ�||'ZZZZZ',NULL as ��Ŀ," & _
                "        Ltrim(To_Char(Sum(Nvl(A.Ӧ�ս��,0)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����" & _
                " FROM " & strTable & " A,�ڼ�� B,�շ���ĿĿ¼ C" & _
                " Where A.�Ǽ�ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                "       And A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by B.�ڼ�||'ZZZZZ'" & _
                " Order by �ڼ�,��Ŀ"
                strMoney = "4,4,7,7"
                
            Case 5 '��Ŀ
                strSQL = _
                " SELECT A.�վݷ�Ŀ as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Nvl(A.Ӧ�ս��,0)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by A.�վݷ�Ŀ Order by ��Ŀ"
                strMoney = "4,7,7"
            Case 6 '���յ���
                strSQL = _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�,A.�վݷ�Ŀ as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Nvl(A.Ӧ�ս��,0)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����," & _
                "        A.����Ա���� as ����Ա,A.��¼����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by A.��¼����,TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.NO,A.�վݷ�Ŀ,A.����Ա����"
                strSQL = strSQL & _
                " Union All" & _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO||'ZZZZZ' as ���ݺ�,NULL as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����,NULL as ����Ա,A.��¼����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " " & _
                " Group by A.��¼����,TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.NO" & _
                " Union All" & _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD')||'ZZZZZ' as ��������,NULL as ���ݺ�,NULL as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����,NULL as ����Ա,-1" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " " & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD')" & _
                " Order by ��������,��¼���� desc,���ݺ�,������Ŀ"
                
                strMoney = "4,4,4,7,7,1"
            Case 7 '���շ���
                strSQL = _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.�վݷ�Ŀ as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Nvl(A.Ӧ�ս��,0)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.�վݷ�Ŀ" & _
                " Union All" & _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD')||'ZZZZZ' as ��������,NULL as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Nvl(A.Ӧ�ս��,0)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as δ����" & _
                " FROM " & strTable & " A,�շ���ĿĿ¼ C" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                        IIf(mstrClass = "", "", " And Instr([8],','''||Nvl(A.��������,Nvl(C.��������,'δ֪'))||''',')>0") & _
                        IIf(mstrChargeType = "", "", " And Instr([9],','''||Nvl(A.�շ����,Nvl(C.���,'��'))||''',')>0") & _
                " " & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD')" & _
                " Order by ��������,������Ŀ"
                strMoney = "4,4,7,7"
        End Select
                
        vsfMain.MergeCells = flexMergeFree
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, "," & mstrTime & ",", DateBegin, DateEnd, _
                    "," & mstrDeptIDs & ",", "," & mstrBaby & ",", "," & mstrItem & ",", "," & mstrClass & ",", "," & mstrChargeType & ",")
        If rsTmp.RecordCount > 0 Then
            Set vsfMain.DataSource = rsTmp
        Else
            Call Grid.BandRec(vsfMain, rsTmp)
        End If
        
        
        vsfMain.Tag = intIndex
        For i = 0 To vsfMain.Cols - 1
            vsfMain.MergeCol(i) = False
        Next
        
        '��ϼ�(С��)
        Select Case intIndex
            Case 0
                For i = 1 To vsfMain.Rows - 1
                    vsfMain.TextMatrix(i, 5) = Format(vsfMain.TextMatrix(i, 5), gstrDec)
                Next i
            Case 1, 3  '��ϸ�嵥��������ϸ
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 5) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 8))
                            If IsNumeric(vsfMain.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 9))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            strTmp = vsfMain.TextMatrix(i, 5)
                            For j = 0 To 7
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "С ��:" & Left(strTmp, Len(strTmp) - 5)
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            For j = 8 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 7
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 8) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 9) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 2 '����Ŀ��ϸ
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 4) Like "*ZZZZZ") Then
                            vsfMain.Cell(flexcpAlignment, i, 6) = 7
                            If IsNumeric(vsfMain.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 8))
                            If IsNumeric(vsfMain.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 9))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            strTmp = vsfMain.TextMatrix(i, 3)
                            For j = 0 To 5
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "С ��:" & strTmp
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            vsfMain.TextMatrix(i, 7) = " " '������
                            For j = 8 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                            vsfMain.Cell(flexcpAlignment, i, 6) = 7
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 7
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 8) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 9) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 4 '�����嵥
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 2))
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 3))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            For j = 0 To 1
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "С��:" & vsfMain.TextMatrix(i - 1, 0)
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            For j = 2 To vsfMain.Cols - 1
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 1
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 5 '�����嵥
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If IsNumeric(vsfMain.TextMatrix(i, 1)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 1))
                        If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 2))
                        vsfMain.MergeRow(i) = False
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.Col = 0: vsfMain.CellAlignment = 4
                    vsfMain.TextMatrix(vsfMain.Row, 0) = "�� ��"
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 1) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 6 '���յ���
                If rsTmp.RecordCount > 0 Then
                    vsfMain.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 1) Like "*ZZZZZ") And Not (vsfMain.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 3))
                            If IsNumeric(vsfMain.TextMatrix(i, 4)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 4))
                            vsfMain.MergeRow(i) = False
                        Else
                            If vsfMain.TextMatrix(i, 1) Like "*ZZZZZ" Then
                                vsfMain.Row = i
                                vsfMain.MergeRow(i) = True
                                For j = 1 To 2
                                    vsfMain.Col = j: vsfMain.CellAlignment = 4
                                    vsfMain.TextMatrix(i, j) = "С��:" & vsfMain.TextMatrix(i - 1, 1)
                                    vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                                Next
                                For j = 3 To vsfMain.Cols - 2
                                    vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                                Next
                            Else
                                vsfMain.Row = i
                                vsfMain.MergeRow(i) = True
                                For j = 0 To 2
                                    vsfMain.Col = j: vsfMain.CellAlignment = 4
                                    vsfMain.TextMatrix(i, j) = "С��:" & vsfMain.TextMatrix(i - 1, 0)
                                    vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                                Next
                                For j = 3 To vsfMain.Cols - 2
                                    vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                                Next
                            End If
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 2
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 4) = Format(arrTotal(1), " " & gstrDec)
                    
                    'ɾ��ֻ��һ�е��ݵ�С����
                    j = 0
                    For i = 1 To vsfMain.Rows - 1
                        If vsfMain.TextMatrix(i, 1) Like "*С��*" Then
                            If j = 1 Then vsfMain.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                    For i = 0 To vsfMain.Cols - 1
                        If vsfMain.TextMatrix(0, i) = "��¼����" Then vsfMain.ColHidden(i) = True
                    Next i
                End If
            Case 7 '���շ�Ŀ
                If rsTmp.RecordCount > 0 Then
                    vsfMain.MergeCol(0) = True
                    ReDim arrTotal(1)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 2))
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 3))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.MergeRow(i) = True
                            vsfMain.Row = i
                            vsfMain.Col = 1: vsfMain.CellAlignment = 4
                            vsfMain.TextMatrix(i, 0) = "С��:" & vsfMain.TextMatrix(i - 1, 0)
                            vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            vsfMain.TextMatrix(i, 1) = vsfMain.TextMatrix(i, 0)
                            For j = 2 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 1
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(1), " " & gstrDec)
                
                    'ɾ��ֻ��һ�з��õ�С����
                    j = 0
                    For i = 1 To vsfMain.Rows - 1
                        If vsfMain.TextMatrix(i, 1) Like "*С��*" Then
                            If j = 1 Then vsfMain.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
        End Select
    Else
        strSQL = "Select ����ʱ��,�Ǽ�ʱ��,NO,�վݷ�Ŀ,��������,����,����,���㵥λ,��׼����,���ʽ��,����Ա����,��������ID,�շ�ϸĿID,����ID From סԺ���ü�¼  where ����ID= [1]  Union ALL " & _
                 "Select ����ʱ��,�Ǽ�ʱ��,NO,�վݷ�Ŀ,��������,����,����,���㵥λ,��׼����,���ʽ��,����Ա����,��������ID,�շ�ϸĿID,����ID From ������ü�¼  where ����ID= [1]"
        
        If mblnDateMoved Then
            strSQL = Replace(Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
        End If
        strSQL = "(" & strSQL & ")"
        
        '��ȡ���ʵ�ʱ,����ʷ�����ϸ
        Select Case intIndex
            Case 0
                strSQL = _
                " Select Trunc(�Ǽ�ʱ��) As ����,A.NO as ���ݺ�,C.���� as ��Ŀ����," & _
                "       A.�վݷ�Ŀ as ��Ŀ," & _
                "       Null as Ӥ����," & _
                "       " & _
                "       " & _
                "       Ltrim(To_Char(A.���ʽ��,'999999999" & gstrDec & "')) as ���ʽ��" & _
                " From " & strSQL & " A,���ű� B,�շ���ĿĿ¼ C" & _
                " Where A.��������ID = B.ID(+) And A.�շ�ϸĿID=C.ID" & _
                "       " & _
                " Order by ���ݺ�,��Ŀ"
                strMoney = "4,4,1,1,1,7"
            Case 1 '��ϸ
                '��������,���ݺ�,����,��Ŀ,��Ŀ,����,����,Ӧ�ս��,���ʽ��,����Ա
                strSQL = _
                " Select To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       Nvl(B.����,'δ֪') as ����,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')|| To_Char(A.����,'999999990.9')||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(A.��׼����,'99999" & gstrFeePrecisionFmt & "')) as ����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(A.���ʽ��,'999999999" & gstrDec & "')) as ���ʽ��,A.����Ա���� as ����Ա" & _
                " From " & strSQL & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID = B.ID(+) And A.�շ�ϸĿID=C.ID" & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Order by ��������,���ݺ�,��Ŀ"
                
                '�嵥��ʽ����
               strMoney = "4,4,1,1,1,4,1,7,7,7,1"
            Case 2 '����Ŀ��ϸ
                '��������,���ݺ�,����,��Ŀ,���,��Ŀ,����,����,Ӧ�ս��,���ʽ��,����,����Ա
                strSQL = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ��������,Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ') as ���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||To_Char(A.����,'99999990.9')||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Nvl(A.���ʽ��,0),'999999999" & gstrDec & "')) as ���ʽ��," & _
                "       Nvl(A.��������,C.��������) as ����,A.����Ա���� as ����Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ��" & _
                " FROM " & strSQL & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID" & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Union All" & _
                " SELECT NULL as ��������,NULL as ���ݺ�,NULL as ��������,Nvl(D.����,C.����) as ��Ŀ,Nvl(C.���,' ')||'ZZZZZ' as ���," & _
                "        NULL as ��Ŀ,to_char(sum(Nvl(A.����,1)*Nvl(A.����,1)),'99999990.9')||' '||A.���㵥λ as ����,NULL as ��׼����," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��,NULL as ����,NULL as ����Ա,NULL as �Ǽ�ʱ��" & _
                " FROM " & strSQL & " A,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.�շ�ϸĿID=C.ID " & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Group by Nvl(D.����,C.����),C.���,A.���㵥λ" & _
                " Order by ��Ŀ,���,��������,���ݺ�"
                strMoney = "4,4,1,1,1,4,1,7,7,7,1,1,1"
            Case 3 '������ϸ
                strSQL = _
                " SELECT To_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�," & _
                "       B.���� as ����,Nvl(D.����,C.����) as ��Ŀ,C.���,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Decode(Nvl(A.����,1),1,'',0,'',A.����||' �� �� ')||To_Char(A.����,'999999990.9')||' '||A.���㵥λ as ����," & _
                "       Ltrim(To_Char(Nvl(A.��׼����,0),'999999999" & gstrFeePrecisionFmt & "')) as ��׼����," & _
                "       Ltrim(To_Char(Round(A.��׼����*A.����*Nvl(A.����,1),5),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Nvl(A.���ʽ��,0),'999999999" & gstrDec & "')) as ���ʽ��,A.����Ա���� as ����Ա " & _
                " FROM " & strSQL & " A,���ű� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID" & _
                "       And A.�շ�ϸĿID=D.�շ�ϸĿID(+) And ����(+)=1 And D.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                " Union All" & _
                " SELECT NULL as ��������,NULL as ���ݺ�,NULL as ����,NULL as ��Ŀ,Null as ���,A.�վݷ�Ŀ||'ZZZZZ' as ��Ŀ," & _
                "       NULL as ����,NULL as ��׼����," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��,NULL as ����Ա" & _
                " FROM " & strSQL & " A,���ű� B,�շ���ĿĿ¼ C" & _
                " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID " & _
                " Group by A.�վݷ�Ŀ||'ZZZZZ' " & _
                " Order by ��Ŀ,��������,���ݺ�"
                strMoney = "4,4,1,1,1,1,1,7,7,7,1"
            Case 4 '�����嵥
                strSQL = _
                " SELECT B.�ڼ�,A.�վݷ�Ŀ as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��" & _
                " FROM " & strSQL & " A,�ڼ�� B" & _
                " Where A.�Ǽ�ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                " Group by B.�ڼ�,A.�վݷ�Ŀ" & _
                " Union All" & _
                " SELECT B.�ڼ�||'ZZZZZ',NULL as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��" & _
                " FROM " & strSQL & " A,�ڼ�� B" & _
                " Where A.�Ǽ�ʱ�� Between Trunc(B.��ʼ����) and Trunc(B.��ֹ����)+1-1/24/60/60 " & _
                " Group by B.�ڼ�||'ZZZZZ'" & _
                " Order by �ڼ�,��Ŀ"
                strMoney = "4,4,7,7"
            Case 5 '�����嵥
                strSQL = _
                " SELECT A.�վݷ�Ŀ as ��Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��" & _
                " FROM " & strSQL & " A" & _
                " Group by A.�վݷ�Ŀ Order by ��Ŀ"
                strMoney = "4,7,7"
            Case 6 '���յ���
                strSQL = _
                    " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO as ���ݺ�,A.�վݷ�Ŀ as ������Ŀ," & _
                    "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��,A.����Ա���� as ����Ա " & _
                    " FROM " & strSQL & " A" & _
                    " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.NO,A.�վݷ�Ŀ,A.����Ա����" & _
                    " Union All" & _
                    " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.NO||'ZZZZZ' as ���ݺ�,NULL as ������Ŀ," & _
                    "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��, NULL as ����Ա  " & _
                    " FROM " & strSQL & " A" & _
                    " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.NO" & vbCrLf & _
                    " Union All" & _
                    " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,'ZZZZZAAAAA' as ���ݺ�,NULL as ������Ŀ," & _
                    "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                    "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��,NULL as ����Ա " & _
                    " FROM  " & strSQL & " A" & _
                    " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD')" & _
                    " Order by ��������,���ݺ�,������Ŀ"
                strMoney = "4,4,4,7,7,1"
            Case 7 '���շ�Ŀ
                strSQL = _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD') as ��������,A.�վݷ�Ŀ as ������Ŀ," & _
                "       Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "       Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��" & _
                " FROM " & strSQL & " A " & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD'),A.�վݷ�Ŀ" & _
                " Union All" & _
                " SELECT TO_Char(A.����ʱ��,'YYYY-MM-DD')||'ZZZZZ' as ��������,NULL as ������Ŀ," & _
                "        Ltrim(To_Char(Sum(Round(A.��׼����*A.����*Nvl(A.����,1),5)),'999999999" & gstrDec & "')) as Ӧ�ս��," & _
                "        Ltrim(To_Char(Sum(Nvl(A.���ʽ��,0)),'999999999" & gstrDec & "')) as ���ʽ��" & _
                " FROM " & strSQL & " A" & _
                " Group by TO_Char(A.����ʱ��,'YYYY-MM-DD')" & _
                " Order by ��������,������Ŀ"
                strMoney = "4,4,7,7"
        End Select
        
        vsfMain.MergeCells = flexMergeFree
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        If rsTmp.RecordCount > 0 Then
            Set vsfMain.DataSource = rsTmp
        Else
            Call Grid.BandRec(vsfMain, rsTmp)
        End If

        vsfMain.Tag = intIndex
        For i = 0 To vsfMain.Cols - 1
            vsfMain.MergeCol(i) = False
        Next

        Select Case intIndex
            Case 1, 3  '��ϸ�嵥��������ϸ
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 5) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 8))
                            If IsNumeric(vsfMain.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 9))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            strTmp = vsfMain.TextMatrix(i, 5)
                            For j = 0 To 7
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "С ��:" & Left(strTmp, Len(strTmp) - 5)
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            For j = 8 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 7
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 8) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 9) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 2 '����Ŀ��ϸ
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 4) Like "*ZZZZZ") Then
                            vsfMain.Cell(flexcpAlignment, i, 6) = 7
                            If IsNumeric(vsfMain.TextMatrix(i, 8)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 8))
                            If IsNumeric(vsfMain.TextMatrix(i, 9)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 9))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            strTmp = vsfMain.TextMatrix(i, 3)
                            For j = 0 To 5
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "С ��:" & strTmp
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            vsfMain.TextMatrix(i, 7) = " " '������
                            For j = 8 To vsfMain.Cols - 2
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                            vsfMain.Cell(flexcpAlignment, i, 6) = 7
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 7
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 8) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 9) = Format(arrTotal(1), " " & gstrDec)
                End If
             Case 4 '�����嵥
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 0) Like "*ZZZZZ") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 2))
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 3))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = True
                            For j = 0 To 1
                                vsfMain.Col = j: vsfMain.CellAlignment = 4
                                vsfMain.TextMatrix(i, j) = "С��:" & vsfMain.TextMatrix(i - 1, 0)
                                vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                            Next
                            For j = 2 To vsfMain.Cols - 1
                                vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                            Next
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 1
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(1), " " & gstrDec)
                End If
             Case 5 '�����嵥
                If rsTmp.RecordCount > 0 Then
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If IsNumeric(vsfMain.TextMatrix(i, 1)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 1))
                        If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 2))
                        vsfMain.MergeRow(i) = False
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.Col = 0: vsfMain.CellAlignment = 4
                    vsfMain.TextMatrix(vsfMain.Row, 0) = "�� ��"
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 1) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(1), " " & gstrDec)
                End If
            Case 6
                For i = 0 To vsfMain.Cols - 1
                    vsfMain.MergeCol(i) = False
                Next
                If rsTmp.RecordCount > 0 Then
                    vsfMain.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not (vsfMain.TextMatrix(i, 1) Like "*ZZZZZ") And Not (vsfMain.TextMatrix(i, 1) Like "*AAAAA") Then
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 3))
                            If IsNumeric(vsfMain.TextMatrix(i, 4)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 4))
                            vsfMain.MergeRow(i) = False
                        Else
                            If Not (vsfMain.TextMatrix(i, 1) Like "*AAAAA") Then
                                vsfMain.Row = i
                                vsfMain.MergeRow(i) = True
                                For j = 1 To 2
                                    vsfMain.Col = j: vsfMain.CellAlignment = 4
                                    vsfMain.TextMatrix(i, j) = "����С��:" & vsfMain.TextMatrix(i - 1, 1)
                                    vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                                Next
                                For j = 3 To vsfMain.Cols - 2
                                    vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                                Next
                            Else
                                vsfMain.Row = i
                                vsfMain.MergeRow(i) = True
                                For j = 1 To 2
                                    vsfMain.Col = j: vsfMain.CellAlignment = 4
                                    vsfMain.TextMatrix(i, j) = "��С��"
                                    vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                                Next
                                For j = 3 To vsfMain.Cols - 2
                                    vsfMain.TextMatrix(i, j) = Space(j Mod 2) & vsfMain.TextMatrix(i, j)
                                Next
                            End If
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 2
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 4) = Format(arrTotal(1), " " & gstrDec)
                    
                    'ɾ��ֻ��һ�е��ݵ�С����
                    j = 0
                    For i = 1 To vsfMain.Rows - 1
                        If vsfMain.TextMatrix(i, 1) Like "*С��*" Then
                            If j = 1 Then vsfMain.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
            Case 7
                For i = 0 To vsfMain.Cols - 1
                    vsfMain.MergeCol(i) = False
                Next
                If rsTmp.RecordCount > 0 Then
                    vsfMain.MergeCol(0) = True
                    ReDim arrTotal(3)
                    For i = 1 To vsfMain.Rows - 1
                        If Not vsfMain.TextMatrix(i, 0) Like "*ZZZZZ" Then
                            If IsNumeric(vsfMain.TextMatrix(i, 2)) Then arrTotal(0) = arrTotal(0) + CCur(vsfMain.TextMatrix(i, 2))
                            If IsNumeric(vsfMain.TextMatrix(i, 3)) Then arrTotal(1) = arrTotal(1) + CCur(vsfMain.TextMatrix(i, 3))
                            vsfMain.MergeRow(i) = False
                        Else
                            vsfMain.Row = i
                            vsfMain.MergeRow(i) = False
                            vsfMain.Col = 0: vsfMain.CellAlignment = 4
                            vsfMain.TextMatrix(i, 0) = Left(vsfMain.TextMatrix(i, 0), Len(vsfMain.TextMatrix(i, 0)) - 5)
                            vsfMain.TextMatrix(i, 1) = "��С��"
                            vsfMain.Cell(flexcpFontBold, i, 0, i, vsfMain.Cols - 1) = True
                        End If
                    Next
                    vsfMain.Rows = vsfMain.Rows + 1
                    vsfMain.Row = vsfMain.Rows - 1
                    vsfMain.MergeRow(vsfMain.Row) = True
                    For i = 0 To 1
                        vsfMain.Col = i: vsfMain.CellAlignment = 4
                        vsfMain.TextMatrix(vsfMain.Row, i) = "�� ��"
                    Next
                    vsfMain.Cell(flexcpFontBold, vsfMain.Row, 0, vsfMain.Row, vsfMain.Cols - 1) = True
                    vsfMain.TextMatrix(vsfMain.Row, 2) = Format(arrTotal(0), gstrDec)
                    vsfMain.TextMatrix(vsfMain.Row, 3) = Format(arrTotal(1), " " & gstrDec)
                    
                    'ɾ��ֻ��һ�е��ݵ�С����
                    j = 0
                    For i = 1 To vsfMain.Rows - 1
                        If vsfMain.TextMatrix(i, 1) Like "*С��*" Then
                            If j = 1 Then vsfMain.RowHeight(i) = 0
                            j = 0
                        Else
                            j = j + 1
                        End If
                    Next
                End If
        End Select
    End If
    
    '�ܵĸ�ʽ����
    If vsfMain.Rows = 1 Then vsfMain.Rows = 2
    
    For i = 0 To vsfMain.Cols - 1
        If vsfMain.TextMatrix(0, i) = "���ʽ��" Then intCol = i
        vsfMain.FixedAlignment(i) = 4
    Next
    
'    lblCancel.Visible = True
    picCancel.Visible = False
    vsfMain.RowHeight(0) = 350
    For i = 1 To vsfMain.Rows - 1
'        If Val(vsfMain.TextMatrix(i, intCol)) < 0 Then vsfMain.TextMatrix(i, intCol) = Format(-1 * vsfMain.TextMatrix(i, intCol), gstrDec): picCancel.Visible = True
        vsfMain.RowHeight(i) = 300
    Next
    
    '���ȡ��,����û�����ó�ʼ�п�,��ӡ���쳣
'    Call SetGridWidth(vsfMain, Me)
    
    '�и���¼������
    If intIndex = 6 And mBalanceType = g_Ed_������� And mBalanceType = g_Ed_סԺ���� Then
        vsfMain.ColWidth(vsfMain.Cols - 1) = 0
    End If
    
    For i = 0 To UBound(Split(strMoney, ","))
        vsfMain.ColAlignment(i) = Split(strMoney, ",")(i)
    Next
    
'    vsfMain.Row = 1: vsfMain.Col = 0
    
    stbThis.Panels(2).Text = strPre
    
    vsfMain.Redraw = True
    vsfMain.Refresh
    Screen.MousePointer = 0
    LoadCardData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    vsfMain.Redraw = True
    If ErrCenter() = 1 Then
        vsfMain.Redraw = False
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    stbThis.Panels(2).Text = strPre
End Function
