VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrugSendQuery 
   AutoRedraw      =   -1  'True
   Caption         =   "ҩ���շ���ѯ"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   Icon            =   "frmDrugSendQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6780
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picQuery 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   1200
      ScaleHeight     =   5535
      ScaleWidth      =   9015
      TabIndex        =   2
      Top             =   1080
      Width           =   9015
      Begin VSFlex8Ctl.VSFlexGrid vsQuery 
         Height          =   5355
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   9030
         _cx             =   15928
         _cy             =   9446
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
         BackColorSel    =   16764057
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   280
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDrugSendQuery.frx":000C
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         AutoSizeMouse   =   0   'False
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
         Begin VB.Frame fraColSel 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   195
            Begin VB.Image imgColSel 
               Height          =   195
               Left            =   0
               Picture         =   "frmDrugSendQuery.frx":00A7
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsColumn 
            Height          =   3270
            Left            =   4800
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   1470
            _cx             =   2593
            _cy             =   5768
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
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmDrugSendQuery.frx":05F5
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
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6420
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14102
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Text            =   "ͨ��"
            TextSave        =   "ͨ��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Text            =   "����"
            TextSave        =   "����"
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
   Begin XtremeSuiteControls.TabControl tbcQuery 
      Height          =   5715
      Left            =   825
      TabIndex        =   1
      Top             =   405
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   10081
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   165
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmDrugSendQuery.frx":0643
      Left            =   675
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDrugSendQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmCond As frmDrugSendQueryCond
Attribute mfrmCond.VB_VarHelpID = -1

Public mMainPrivs As String 'IN:���������������е�Ȩ��,ע����ڲ�ģ��Ȩ��
Public mlng����ID As Long 'IN:���ڼ�¼������Ĳ������ϴβ�ѯ����
Public mlng����ID As Long 'IN
Private mblnOnePati As Boolean 'IN��������ģʽ
Private mblnMoved As Boolean
Private mbytSize As Byte '������壺0-9�����壨С���壩��1-12�����壨�����壩
Private mstrNewHead  As String '��ҩ��ϸ�嵥��ѡ���к����ͷ��Ϣ
'��ҩ��ϸ�嵥ԭ������ͷ��Ϣ
Private Const mstrOldHead = "��Ч,850,1;״̬,850,1;ҩƷ��Ϣ,5000,1;����,850,1;����,850,1;����,1000,7;���,1000,7;����,850,1;Ƶ��,1000,1;�÷�,1000,1;����ʱ��,1530,1;������,750,1"

'��ѯ����
Private Type QUERY_COND
    Mode As Byte '0-��ҽ������ʱ��,1-��ҩ����ҩʱ��
    DateBegin As Date
    DateEnd As Date
    ��ҩDateB As Date
    ��ҩDateE As Date
    ��ҩ;�� As String
    NO As String
    ��ҩ�� As String
    ҩ��ID As Long
    ����IDs As String
    ����ID As Long
    ��ҩ����ID As Long
    ��Ч As Integer '2-ȫ��
    ״̬ As String
End Type
Private mvQuery As QUERY_COND

Public Sub ShowQuery(frmParent As Object, strPriv As String, lng����ID As Long, lng����ID As Long, ByVal blnOnePati As Boolean)
    mMainPrivs = strPriv
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mblnOnePati = blnOnePati
    Me.Show , frmParent
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_File_Print
        Call OutputList(1)
    Case conMenu_File_Preview
        Call OutputList(2)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit
        Unload Me
    Case conMenu_FontSet_FontSize_S 'С����
        If mbytSize <> 0 Then
            mbytSize = 0
            Call Grid.SetFontSize(vsColumn, IIF(mbytSize = 0, 9, 12))
            Call Grid.SetFontSize(vsQuery, IIF(mbytSize = 0, 9, 12))
        End If
    Case conMenu_FontSet_FontSize_L '������
        If mbytSize <> 1 Then
            mbytSize = 1
            Call Grid.SetFontSize(vsColumn, IIF(mbytSize = 0, 9, 12))
            Call Grid.SetFontSize(vsQuery, IIF(mbytSize = 0, 9, 12))
        End If
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    With Me.tbcQuery
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_FontSet_FontSize_S 'С����
            Control.Checked = Not (mbytSize = 1)
        Case conMenu_FontSet_FontSize_L '������
            Control.Checked = (mbytSize = 1)
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Item.Handle = mfrmCond.Hwnd
End Sub

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objItem As TabControlItem
    Dim objPane As Pane

    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_FontSet, "�������")
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_FontSet_FontSize_S, "С����(&S)", -1, False
            .Add xtpControlButton, conMenu_FontSet_FontSize_L, "������(&L)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_FontSet_FontSize_S
        .Add FALT, vbKeyL, conMenu_FontSet_FontSize_L
    End With

    '������----------------------------------------------
    Set mfrmCond = New frmDrugSendQueryCond
    Call mfrmCond.InitParameter(mMainPrivs, mlng����ID, mlng����ID, mblnOnePati)
    
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 250, 400, DockLeftOf, Nothing)
    objPane.Title = "��ѯ����"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    mstrNewHead = mstrOldHead
    'ҳ������----------------------------------------------
    
    
    
    With Me.tbcQuery
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
        End With
        Set objItem = .InsertItem(0, "��ҩ��ϸ�嵥", picQuery.Hwnd, 0): objItem.Color = tbcQuery.PaintManager.ColorSet.ButtonNormal
        Set objItem = .InsertItem(1, "��ҩ�����嵥", picQuery.Hwnd, 0): objItem.Color = tbcQuery.PaintManager.ColorSet.ButtonNormal
        Set objItem = .InsertItem(2, "��ҩ��ϸ�嵥", picQuery.Hwnd, 0): objItem.Color = &HC0C0FF
        Set objItem = .InsertItem(3, "��ҩ�����嵥", picQuery.Hwnd, 0): objItem.Color = &HC0C0FF
        Set objItem = .InsertItem(4, "���˻����嵥", picQuery.Hwnd, 0): objItem.Color = tbcQuery.PaintManager.ColorSet.ButtonNormal
        
        '��Ϊ����ͬ,���Ҫ�л��ص�1��;�����ݲ�Ӱ���ٶ�
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    '���ñ������
    mbytSize = Val(zlDatabase.GetPara("ҩ���շ���ѯ�������", glngSys, pסԺҽ������, "0"))
    Call Grid.SetFontSize(vsColumn, IIF(mbytSize = 0, 9, 12))
    Call Grid.SetFontSize(vsQuery, IIF(mbytSize = 0, 9, 12))
            
    Call RestoreWinState(Me, App.ProductName)
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mvQuery.DateBegin = Empty
    Call zlDatabase.SetPara("ҩ���շ���ѯ�������", mbytSize, glngSys, pסԺҽ������, InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��ѡ������;") > 0)
    If Not mfrmCond Is Nothing Then
        Unload mfrmCond
        Set mfrmCond = Nothing
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub InitQueryTable()
    Dim arrHead As Variant, strHead As String, i As Long
    
    If tbcQuery.Selected.Index = 0 Then
        strHead = mstrNewHead
    ElseIf tbcQuery.Selected.Index = 1 Then
        strHead = "����,1000,1;ҩƷ��Ϣ,5000,1;����,850,1;���,1000,7"
    ElseIf tbcQuery.Selected.Index = 2 Then
        strHead = "��Ч,500,1;ҩƷ��Ϣ,5000,1;����,850,1;����,850,1;����,1000,7;���,1000,7;����,850,1;Ƶ��,1000,1;�÷�,1000,1;����ʱ��,1530,1;������,750,1"
    ElseIf tbcQuery.Selected.Index = 3 Then
        strHead = "����,1000,1;ҩƷ��Ϣ,5000,1;����,850,1;���,1000,7"
    ElseIf tbcQuery.Selected.Index = 4 Then
        strHead = "����,1000,1;ҩƷ��Ϣ,5000,1;Ӧ����,850,1;��ҩ��,850,1;ʵ����,850,1;���,1000,7"
    End If
    arrHead = Split(strHead, ";")
    With vsQuery
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)
                If .ColWidth(.FixedCols + i) = 0 Then
                    .ColHidden(.FixedCols + i) = True
                End If
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Function LoadQueryData() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strKey As String, curTotal As Currency
    Dim strSQL As String, strSQLSend As String, strSQLDel As String
    Dim strCondSend As String, strCondDel As String
    Dim strSub1 As String, strSub2 As String
    Dim i As Long, j As Long
    Dim intBedLen As Integer
    
    If mvQuery.DateBegin = Empty Then
        vsQuery.Rows = vsQuery.FixedRows
        vsQuery.Rows = vsQuery.FixedRows + 1
        LoadQueryData = True: Exit Function
    End If
    
    Me.Refresh
    
    On Error GoTo errH
    
    '��ҩ����������SQL
    '------------------------------------------------------------------------------------------
    '�����޷�ȷ��,ֻ����ʱ����ȷ��
    mblnMoved = zlDatabase.DateMoved(mvQuery.DateBegin)
    
    If InStr(",0,1,4,", tbcQuery.Selected.Index) > 0 Then
        'ҩ��
        If mvQuery.ҩ��ID <> 0 Then
            strCondSend = strCondSend & " And A.�ⷿID+0=[1]"
        End If
        
        'ʱ��:Ҫ�Է���ʱ��,��Ϊ��ҩʱ���������ڲ�һ��
        If mvQuery.Mode = 0 Then
            strCondSend = strCondSend & " And (A.NO,A.����) IN (Select NO,Decode(��¼����,1,8,9) From ����ҽ������ Where ����ʱ�� Between [2] And [3])"
        Else
            strCondSend = strCondSend & " And A.������� Between [2] And [3]"
        End If
        
        'NO
        If mvQuery.NO <> "" Then strCondSend = strCondSend & " And A.NO=[4]"
        '��ҩ��
        If mvQuery.��ҩ�� <> "" Then strCondSend = strCondSend & " And A.���ܷ�ҩ��=[12] "
        '��Ч
        If mvQuery.��Ч = 0 Or mvQuery.��Ч = 1 Then
            strCondSend = strCondSend & " And Nvl(Substr(A.����,1,1),0)=[5]"
        End If
        
        '��ҩ״̬
        If Val(Mid(mvQuery.״̬, 1, 1)) = 1 And Val(Mid(mvQuery.״̬, 2, 1)) = 1 Then
            strCondSend = strCondSend & " And (Mod(A.��¼״̬,3)=1 And A.����� is Null Or A.����� is Not Null)"
        ElseIf Val(Mid(mvQuery.״̬, 1, 1)) = 1 Then
            strCondSend = strCondSend & " And Mod(A.��¼״̬,3)=1 And A.����� is Null"
        ElseIf Val(Mid(mvQuery.״̬, 2, 1)) = 1 Then
            strCondSend = strCondSend & " And A.����� is Not Null"
        End If
        
        '��ҩ;��
        If mvQuery.��ҩ;�� <> "" Then
            strCondSend = strCondSend & " And A.�÷� IN(Select Column_Value From Table(Cast(f_Str2list([8]) As zlTools.t_Strlist)))"
        End If
        
        '��������SQL
        strSQLSend = "Select Decode(A.�����,NULL,0,1) as ״̬,A.NO,A.���," & _
            " Sum(A.��д����*A.����) as ����,Sum(A.���۽��) as ���" & _
            " From ҩƷ�շ���¼ A" & _
            " Where A.���� IN(9,10)" & IIF(mvQuery.��ҩ����ID = 0, "", " And a.�Է�����ID=[13]") & strCondSend & _
            " Group by Decode(A.�����,NULL,0,1),A.NO,A.���" & _
            " Having Nvl(Sum(A.��д����),0)<>0 Or Nvl(Sum(A.���۽��),0)<>0"
        strSQLSend = "Select B.״̬,C.NO,D.���,C.�ⷿID,C.�Է�����ID,D.����,D.��ʶ�� as סԺ��,D.����," & _
            " C.ҩƷID,B.����,D.��׼���� as ����,B.���,Decode(Nvl(Substr(C.����,1,1),0),0,'����','����') as ��Ч," & _
            " C.����,C.Ƶ��,C.�÷�,A.����ʱ�� as ʱ��,A.������ as ��Ա,C.����,d.����id,d.��ҳid" & _
            " From ����ҽ������ A,(" & strSQLSend & ") B,ҩƷ�շ���¼ C,סԺ���ü�¼ D" & _
            " Where B.NO=C.NO And B.���=C.��� And (C.��¼״̬=1 Or Mod(C.��¼״̬,3)=0)" & _
             IIF(mvQuery.��ҩ����ID = 0, "", " And c.�Է�����ID=[13]") & _
            " And C.����ID=D.ID And A.NO=D.NO And A.ҽ��ID=D.ҽ����� And A.��¼����=2" & _
            IIF(mvQuery.����ID <> 0, " And D.���˲���ID+0=[6]", "") & _
            IIF(mvQuery.����IDs <> "", " And D.����ID+0 IN(Select Column_Value From Table(Cast(f_Num2list([7]) As zlTools.t_Numlist)))", "")
        If mblnMoved Then
            strSub1 = strSQLSend
            strSub1 = Replace(strSub1, "סԺ���ü�¼", "HסԺ���ü�¼")
            strSub1 = Replace(strSub1, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
            
            strSub2 = strSQLSend
            strSub2 = Replace(strSub2, "����ҽ������", "H����ҽ������")
            strSub2 = Replace(strSub2, "סԺ���ü�¼", "HסԺ���ü�¼")
            strSub2 = Replace(strSub2, "ҩƷ�շ���¼", "HҩƷ�շ���¼")

            strSQLSend = strSQLSend & " Union ALL " & strSub1 & " Union ALL " & strSub2
        End If
        '��������ϸ�ϲ�
        strSQLSend = "Select ״̬,NO,���,�ⷿID,�Է�����ID,����,סԺ��,����,ҩƷID,����," & _
            " Sum(����) as ����,Sum(���) as ���,��Ч,����,Ƶ��,�÷�,ʱ��,��Ա,����,����id,��ҳid" & _
            " From (" & strSQLSend & ")" & _
            " Group by ״̬,NO,���,�ⷿID,�Է�����ID,����,סԺ��,����,ҩƷID,����,��Ч,����,Ƶ��,�÷�,ʱ��,��Ա,����,����id,��ҳid"
    End If
    
    '��ҩ����������SQL
    '------------------------------------------------------------------------------------------
    If InStr(",2,3,4,", tbcQuery.Selected.Index) > 0 Then '��ҩ����
        'ҩ��
        strCondDel = strCondDel & " And A.��˲���ID=[1]"
        
        '��ҩ����ʱ��
        If mvQuery.��ҩDateB <> Empty And mvQuery.��ҩDateE <> Empty Then
            strCondDel = strCondDel & " And A.����ʱ�� Between [9] And [10]"
        Else
            strCondDel = strCondDel & " And A.����ʱ�� Between Sysdate-1 And Sysdate"
        End If
        
        'NO
        If mvQuery.NO <> "" Then
            strCondDel = strCondDel & " And B.NO=[4]"
        End If
        '��ҩ��
        If mvQuery.��ҩ�� <> "" Then strCondDel = strCondDel & " And D.���ܷ�ҩ��=[12] "
        '��Ч
        If mvQuery.��Ч = 0 Or mvQuery.��Ч = 1 Then
            strCondDel = strCondDel & " And Nvl(C.ҽ����Ч,0)=[5]"
        End If
        
        '����
        strCondDel = strCondDel & " And A.���벿��ID=[6]"
        
        '����ID
        If mvQuery.����IDs <> "" Then
            strCondDel = strCondDel & " And B.����ID+0 IN(Select Column_Value From Table(Cast(f_Num2list([7]) As zlTools.t_Numlist)))"
        End If
        
        '��ҩ;��
        If mvQuery.��ҩ;�� <> "" Then
            strCondDel = strCondDel & " And D.�÷� IN(Select Column_Value From Table(Cast(f_Str2list([8]) As zlTools.t_Strlist)))"
        End If
        
        '��������SQL����������ϸ�ϲ�
        strSQLDel = "Select Distinct -1 as ״̬,D.NO,B.���,D.�ⷿID,D.�Է�����ID,B.����,B.��ʶ�� as סԺ��,B.����," & _
            " D.ҩƷID,A.����,B.��׼���� as ����,A.����*B.��׼���� as ���,Decode(Nvl(C.ҽ����Ч,0),0,'����','����') as ��Ч," & _
            " D.����,D.Ƶ��,D.�÷�,A.����ʱ�� as ʱ��,A.������ as ��Ա,D.����,b.����id,b.��ҳid" & _
            " From ���˷������� A,סԺ���ü�¼ B,����ҽ����¼ C,ҩƷ�շ���¼ D" & _
            " Where Nvl(A.״̬,0)=0 And A.����ID=B.ID And B.�շ���� IN('5','6','7')" & _
            IIF(mvQuery.��ҩ����ID = 0, "", " And d.�Է�����ID=[13]") & _
            " And B.ҽ�����=C.ID And B.ID=D.����ID And (D.��¼״̬=1 Or Mod(D.��¼״̬,3)=0)" & strCondDel
        If mblnMoved Then
            strSub1 = strSQLDel
            strSub1 = Replace(strSub1, "סԺ���ü�¼", "HסԺ���ü�¼")
            strSub1 = Replace(strSub1, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
            
            strSub2 = strSQLDel
            strSub2 = Replace(strSub2, "����ҽ����¼", "H����ҽ����¼")
            strSub2 = Replace(strSub2, "סԺ���ü�¼", "HסԺ���ü�¼")
            strSub2 = Replace(strSub2, "ҩƷ�շ���¼", "HҩƷ�շ���¼")

            strSQLDel = strSQLDel & " Union ALL " & strSub1 & " Union ALL " & strSub2
        End If
    End If
    
    '������ͬ�Ĳ�ѯSQL
    '------------------------------------------------------------------------------------------
    If tbcQuery.Selected.Index = 0 Or tbcQuery.Selected.Index = 2 Then '��ҩ��ϸ����ҩ��ϸ
        intBedLen = GetMaxBedLen(mvQuery.����ID, False)
        strSQL = IIF(tbcQuery.Selected.Index = 0, strSQLSend, strSQLDel)
        strSQL = _
            " Select /*+ Rule*/ A.״̬,A.NO,A.���,I.���� as ҩ��,H.���� as ��������,A.����,e.סԺ��,LPAD(e.��Ժ����," & intBedLen & ",' ') as ����," & _
            " Nvl(X.����,F.����)||Decode(F.����,NULL,NULL,'('||F.����||')')||Decode(F.���,NULL,NULL,' '||F.���) as ҩƷ��Ϣ," & _
            " A.����/Nvl(E.סԺ��װ,1) as ����,E.סԺ��λ,A.����*Nvl(E.סԺ��װ,1) as ����," & _
            " A.���,A.��Ч,A.����,G.���㵥λ as ������λ,A.Ƶ��,A.�÷�,A.ʱ��,A.��Ա,A.����,g.���" & _
            " From (" & strSQL & ") A,ҩƷ��� E,�շ���ĿĿ¼ F,������ĿĿ¼ G,���ű� H,���ű� I,�շ���Ŀ���� X,������ҳ E" & _
            " Where A.ҩƷID=E.ҩƷID And A.ҩƷID=F.ID And E.ҩ��ID=G.ID" & _
            " And A.�Է�����ID=H.ID And A.�ⷿID=I.ID And a.����id=e.����id and a.��ҳid=e.��ҳid" & _
            " And F.ID=X.�շ�ϸĿID(+) And X.����(+)=1 And X.����(+)=[11]" & _
            " Order by ����,A.NO,A.���"
    ElseIf tbcQuery.Selected.Index = 1 Or tbcQuery.Selected.Index = 3 Then '��ҩ���ܡ���ҩ����
        strSQL = IIF(tbcQuery.Selected.Index = 1, strSQLSend, strSQLDel)
        strSQL = "Select B.ҩƷID,C.���� as ҩƷ����,C.����," & _
            " C.����,C.���,B.סԺ��λ,Sum(A.����/Nvl(B.סԺ��װ,1)) as ����,Sum(A.���) as ���" & _
            " From (" & strSQL & ") A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
            " Where A.ҩƷID=B.ҩƷID And A.ҩƷID=C.ID" & _
            " Group by B.ҩƷID,C.����,C.����,C.����,C.���,B.סԺ��λ" & _
            " Having Sum(A.����/Nvl(B.סԺ��װ,1))<>0 Or Sum(A.���)<>0"
    
        strSQL = "Select /*+ Rule*/ A.ҩƷ����," & _
            " Nvl(B.����,A.����)||Decode(A.����,NULL,NULL,'('||A.����||')')||Decode(A.���,NULL,NULL,' '||A.���) as ҩƷ��Ϣ," & _
            " A.סԺ��λ,A.����,A.���" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B" & _
            " Where A.ҩƷID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[11]" & _
            " Order by A.ҩƷ����"
    ElseIf tbcQuery.Selected.Index = 4 Then '���˻���
        strSQL = "Select ҩƷID,���� as Ӧ����,0 as ��ҩ��,��� From (" & strSQLSend & ")" & _
            " Union ALL Select ҩƷID,0 as Ӧ����,���� as ��ҩ��,-1*��� From (" & strSQLDel & ")"
            
        strSQL = "Select B.ҩƷID,C.���� as ҩƷ����,C.����,C.����,C.���,B.סԺ��λ," & _
            " Sum(A.Ӧ����/Nvl(B.סԺ��װ,1)) as Ӧ����,Sum(A.��ҩ��/Nvl(B.סԺ��װ,1)) as ��ҩ��," & _
            " (Sum(A.Ӧ����)-Sum(A.��ҩ��))/Nvl(B.סԺ��װ,1) as ʵ����,Sum(A.���) as ���" & _
            " From (" & strSQL & ") A,ҩƷ��� B,�շ���ĿĿ¼ C" & _
            " Where A.ҩƷID=B.ҩƷID And A.ҩƷID=C.ID" & _
            " Group by B.ҩƷID,C.����,C.����,C.����,C.���,B.סԺ��λ,Nvl(B.סԺ��װ,1)" & _
            " Having Sum(A.Ӧ����/Nvl(B.סԺ��װ,1))<>0 Or Sum(A.��ҩ��/Nvl(B.סԺ��װ,1))<>0 Or Sum(A.���)<>0"
    
        strSQL = "Select /*+ Rule*/ A.ҩƷ����," & _
            " Nvl(B.����,A.����)||Decode(A.����,NULL,NULL,'('||A.����||')')||Decode(A.���,NULL,NULL,' '||A.���) as ҩƷ��Ϣ," & _
            " A.סԺ��λ,A.Ӧ����,A.��ҩ��,A.ʵ����,A.���" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B" & _
            " Where A.ҩƷID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[11]" & _
            " Order by A.ҩƷ����"
    End If
    
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    With mvQuery
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .ҩ��ID, .DateBegin, .DateEnd, _
            .NO, .��Ч, .����ID, .����IDs, .��ҩ;��, .��ҩDateB, .��ҩDateE, IIF(gbytҩƷ������ʾ = 0, 1, 3), Val(.��ҩ��), Val(.��ҩ����ID))
    End With
    
    If Not rsTmp.EOF Then
        With vsQuery
            .Redraw = flexRDNone
            .Rows = .FixedRows
            If tbcQuery.Selected.Index = 0 Or tbcQuery.Selected.Index = 2 Then '��ҩ��ϸ����ҩ��ϸ
                For i = 1 To rsTmp.RecordCount
                    If strKey <> rsTmp!NO Then
                        If strKey <> "" Then
                            j = .FindRow(CStr(strKey))
                            If j <> -1 Then
                                .Cell(flexcpText, j, 0, j, .Cols - 1) = .TextMatrix(j, 0) & Format(curTotal, gstrDec)
                                .Cell(flexcpBackColor, j, 0, j, .Cols - 1) = &HEFF0EF   '����ͷ
                            End If
                        End If
                        
                        .AddItem ""
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = CStr(rsTmp!NO)
                        .TextMatrix(.Rows - 1, 0) = " ���ݺ�:��" & rsTmp!NO & "��������:" & rsTmp!�������� & _
                            "������:��" & rsTmp!���� & "����סԺ��:" & Nvl(rsTmp!סԺ��) & _
                            "������:" & Nvl(rsTmp!����) & "�����:"
                        curTotal = 0
                    End If
                    
                    .AddItem ""
                    
                    If tbcQuery.Selected.Index = 0 Then
                        .TextMatrix(.Rows - 1, 0) = Nvl(rsTmp!��Ч)
                        .TextMatrix(.Rows - 1, 1) = IIF(Nvl(rsTmp!״̬, 0) = 0, "δ��ҩ", "�ѷ�ҩ")
                        .TextMatrix(.Rows - 1, 2) = rsTmp!ҩƷ��Ϣ
                        If rsTmp!��� & "" = "7" Then
                            .TextMatrix(.Rows - 1, 3) = rsTmp!���� & ""
                        End If
                        .TextMatrix(.Rows - 1, 4) = FormatEx(rsTmp!����, 5) & Nvl(rsTmp!סԺ��λ)
                        .TextMatrix(.Rows - 1, 5) = Format(Nvl(rsTmp!����, 0), gstrDecPrice)
                        .TextMatrix(.Rows - 1, 6) = Format(Nvl(rsTmp!���, 0), gstrDec)
                        .TextMatrix(.Rows - 1, 7) = IIF(Not IsNull(rsTmp!����), FormatEx(Nvl(rsTmp!����, 0), 5) & Nvl(rsTmp!������λ), "")
                        .TextMatrix(.Rows - 1, 8) = Nvl(rsTmp!Ƶ��)
                        .TextMatrix(.Rows - 1, 9) = Nvl(rsTmp!�÷�)
                        .TextMatrix(.Rows - 1, 10) = Format(Nvl(rsTmp!ʱ��), "yyyy-MM-dd HH:mm")
                        .TextMatrix(.Rows - 1, 11) = Nvl(rsTmp!��Ա)
                    Else
                        .TextMatrix(.Rows - 1, 0) = Nvl(rsTmp!��Ч)
                        .TextMatrix(.Rows - 1, 1) = rsTmp!ҩƷ��Ϣ
                        If rsTmp!��� & "" = "7" Then
                            .TextMatrix(.Rows - 1, 2) = rsTmp!���� & ""
                        End If
                        .TextMatrix(.Rows - 1, 3) = FormatEx(rsTmp!����, 5) & Nvl(rsTmp!סԺ��λ)
                        .TextMatrix(.Rows - 1, 4) = Format(Nvl(rsTmp!����, 0), gstrDecPrice)
                        .TextMatrix(.Rows - 1, 5) = Format(Nvl(rsTmp!���, 0), gstrDec)
                        .TextMatrix(.Rows - 1, 6) = IIF(Not IsNull(rsTmp!����), FormatEx(Nvl(rsTmp!����, 0), 5) & Nvl(rsTmp!������λ), "")
                        .TextMatrix(.Rows - 1, 7) = Nvl(rsTmp!Ƶ��)
                        .TextMatrix(.Rows - 1, 8) = Nvl(rsTmp!�÷�)
                        .TextMatrix(.Rows - 1, 9) = Format(Nvl(rsTmp!ʱ��), "yyyy-MM-dd HH:mm")
                        .TextMatrix(.Rows - 1, 10) = Nvl(rsTmp!��Ա)
                    End If
                    
                    If Nvl(rsTmp!״̬, 0) = 0 Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC00000 'δ��ҩ
                    End If
                    curTotal = curTotal + Nvl(rsTmp!���, 0)
                    
                    strKey = rsTmp!NO
                    rsTmp.MoveNext
                Next
                Call vsQuery.AutoSize(2)
            ElseIf tbcQuery.Selected.Index = 1 Or tbcQuery.Selected.Index = 3 Then '��ҩ���ܡ���ҩ����
                For i = 1 To rsTmp.RecordCount
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = Nvl(rsTmp!ҩƷ����)
                    .TextMatrix(.Rows - 1, 1) = rsTmp!ҩƷ��Ϣ
                    .TextMatrix(.Rows - 1, 2) = FormatEx(rsTmp!����, 5) & Nvl(rsTmp!סԺ��λ)
                    .TextMatrix(.Rows - 1, 3) = Format(Nvl(rsTmp!���, 0), gstrDec)
                    
                    curTotal = curTotal + Nvl(rsTmp!���, 0)
                    
                    rsTmp.MoveNext
                Next
                Call vsQuery.AutoSize(1)
            ElseIf tbcQuery.Selected.Index = 4 Then  '���˻���
                For i = 1 To rsTmp.RecordCount
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = Nvl(rsTmp!ҩƷ����)
                    .TextMatrix(.Rows - 1, 1) = rsTmp!ҩƷ��Ϣ
                    If Nvl(rsTmp!Ӧ����, 0) <> 0 Then .TextMatrix(.Rows - 1, 2) = FormatEx(rsTmp!Ӧ����, 5) & Nvl(rsTmp!סԺ��λ)
                    If Nvl(rsTmp!��ҩ��, 0) <> 0 Then .TextMatrix(.Rows - 1, 3) = FormatEx(rsTmp!��ҩ��, 5) & Nvl(rsTmp!סԺ��λ)
                    .TextMatrix(.Rows - 1, 4) = FormatEx(rsTmp!ʵ����, 5) & Nvl(rsTmp!סԺ��λ)
                    .TextMatrix(.Rows - 1, 5) = Format(Nvl(rsTmp!���, 0), gstrDec)
                    
                    If Nvl(rsTmp!��ҩ��, 0) <> 0 Then .Cell(flexcpForeColor, .Rows - 1, 3) = vbRed
                    
                    curTotal = curTotal + Nvl(rsTmp!���, 0)
                    
                    rsTmp.MoveNext
                Next
                Call vsQuery.AutoSize(1)
            End If
            
            '���һ����ͷ
            If InStr(",1,3,4,", tbcQuery.Selected.Index) > 0 Then
                .AddItem "", .FixedRows
                .MergeRow(.FixedRows) = True
                .RowData(.FixedRows) = 1
                .Cell(flexcpText, .FixedRows, 0, .FixedRows, .Cols - 1) = "���ϼ�:" & Format(curTotal, gstrDec)
                .Cell(flexcpBackColor, .FixedRows, 0, .FixedRows, .Cols - 1) = &HEFF0EF    '����ͷ
            ElseIf InStr(",0,2,", tbcQuery.Selected.Index) > 0 Then
                j = .FindRow(CStr(strKey))
                If j <> -1 Then
                    .Cell(flexcpText, j, 0, j, .Cols - 1) = .TextMatrix(j, 0) & Format(curTotal, gstrDec)
                    .Cell(flexcpBackColor, j, 0, j, .Cols - 1) = &HEFF0EF    '����ͷ
                End If
            End If
            
            Call SetMinRowHeight
            .Row = .FixedRows: .Col = 0
            Call vsQuery_AfterRowColChange(-1, -1, .Row, .Col)
            .Redraw = flexRDDirect
        End With
    End If
    Call zlCommFun.StopFlash
    LoadQueryData = True
    Exit Function
errH:
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetMinRowHeight()
'���ܣ�����AutoSize֮��,������ʾ���иߵ���Ϊ��С�߶�
    Dim i As Long
    
    With vsQuery
        .Redraw = flexRDNone
        For i = 0 To .Rows - 1
            If .RowData(i) <> "" Then
                .RowHeight(i) = 300
            ElseIf .RowHeight(i) < .RowHeightMin Then
                .RowHeight(i) = .RowHeightMin
            End If
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim i As Long
    If Not imgColSel.Visible Then Exit Sub
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsColumn
            If .Visible Then
                .Visible = False
                vsQuery.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vsQuery.ColHidden(.RowData(i)) Or vsQuery.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 150
                If vsColumn.Top + vsColumn.Height > Me.ScaleHeight Then
                    vsColumn.Height = Me.ScaleHeight - vsColumn.Top
                    vsColumn.Width = 1750
                Else
                    vsColumn.Width = 1470
                End If
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub mfrmCond_DoQuery(ByVal ҩ��ID As Long, ByVal Mode As Byte, ByVal DateBegin As Date, ByVal DateEnd As Date, ByVal ��ҩDateB As Date, ByVal ��ҩDateE As Date, ByVal NO As String, ByVal ��ҩ�� As String, ByVal ��Ч As Integer, ByVal ״̬ As String, ByVal ����ID As Long, ByVal ����IDs As String, ByVal ��ҩ;�� As String, ByVal ��ҩ����ID As Long)
    mvQuery.ҩ��ID = ҩ��ID
    mvQuery.Mode = Mode
    mvQuery.DateBegin = DateBegin
    mvQuery.DateEnd = DateEnd
    mvQuery.��ҩDateB = ��ҩDateB
    mvQuery.��ҩDateE = ��ҩDateE
    mvQuery.NO = NO
    mvQuery.��ҩ�� = ��ҩ��
    mvQuery.��Ч = ��Ч
    mvQuery.״̬ = ״̬
    mvQuery.����ID = ����ID
    mvQuery.����IDs = ����IDs
    mvQuery.��ҩ;�� = ��ҩ;��
    mvQuery.��ҩ����ID = ��ҩ����ID
    
    '��ѯ�����ѷ�ҩ��ʱ������ʾ��ҩ����
    If Mid(mvQuery.״̬, 2, 1) = "1" Then
        If tbcQuery.Selected.Index >= 2 Then
            tbcQuery(0).Selected = True
        End If
        tbcQuery(2).Visible = False
        tbcQuery(3).Visible = False
        tbcQuery(4).Visible = False
    Else
        tbcQuery(2).Visible = True
        tbcQuery(3).Visible = True
        tbcQuery(4).Visible = True
    End If
    Call tbcQuery_SelectedChanged(tbcQuery.Selected)
    
    vsQuery.SetFocus
End Sub

Private Sub picQuery_Resize()
    With picQuery
        vsQuery.Top = picQuery.ScaleTop
        vsQuery.Left = picQuery.ScaleLeft
        vsQuery.Height = picQuery.ScaleHeight
        vsQuery.Width = picQuery.ScaleWidth
    End With
    fraColSel.Left = vsQuery.Left
    fraColSel.Top = vsQuery.Top

End Sub

Private Sub tbcQuery_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Visible Then
        Call SaveFlexState(vsQuery, App.ProductName & "\" & Me.Name)
    End If
    vsColumn.Visible = False
    
    vsQuery.Tag = Item.Index
    Call InitQueryTable
    
    If Item.Index = 0 Then
        fraColSel.Visible = True
        imgColSel.Visible = True
        Call InitColumnSelect
    Else
        fraColSel.Visible = False
        imgColSel.Visible = False
    End If
    
    If Visible Then
        Call RestoreFlexState(vsQuery, App.ProductName & "\" & Me.Name)
        Call LoadQueryData
    End If
    
    If Visible Then vsQuery.SetFocus
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long, lnPos As Long
    Dim strOldCOLInfo As String, strNewCOLInfo As String
    strOldCOLInfo = Mid(";" & mstrOldHead & ";", InStr(";" & mstrOldHead & ";", ";" & Trim(vsColumn.TextMatrix(Row, 1))))
    strOldCOLInfo = Mid(strOldCOLInfo, 2, InStr(2, strOldCOLInfo, ";") - 2)
    strNewCOLInfo = Mid(";" & mstrNewHead & ";", InStr(";" & mstrNewHead & ";", ";" & Trim(vsColumn.TextMatrix(Row, 1))))
    strNewCOLInfo = Mid(strNewCOLInfo, 2, InStr(2, strNewCOLInfo, ";") - 2)
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            If vsQuery.ColWidth(lngCol) = 0 Then
                mstrNewHead = Replace(mstrNewHead, Trim(vsColumn.TextMatrix(Row, 1)) & ",0," & Split(strNewCOLInfo, ",")(2), strOldCOLInfo)
                vsQuery.ColWidth(lngCol) = Val(Split(strOldCOLInfo, ",")(1))
            Else
                mstrNewHead = Replace(mstrNewHead, strNewCOLInfo, strOldCOLInfo)
                vsQuery.ColWidth(lngCol) = vsQuery.ColData(lngCol)
            End If
            vsQuery.ColHidden(lngCol) = False
        Else
            vsQuery.ColWidth(lngCol) = 0
            vsQuery.ColHidden(lngCol) = True
            mstrNewHead = Replace(mstrNewHead, strNewCOLInfo, Trim(vsColumn.TextMatrix(Row, 1)) & ",0," & Split(strOldCOLInfo, ",")(2))
        End If
    End If
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub

Private Sub vsQuery_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsQuery
        If OldRow <> NewRow Then
            If OldRow >= .FixedRows And OldRow <= .Rows - 1 Then
                If .RowData(OldRow) <> "" Then
                    .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = &HEFF0EF
                Else
                    .Cell(flexcpBackColor, OldRow, 0, OldRow, .Cols - 1) = .BackColor
                End If
            End If
            If NewRow >= .FixedRows And NewRow <= .Rows - 1 Then
                If .RowData(NewRow) <> "" Then
                    .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = &HEFF0EF
                Else
                    .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = &HFFCC99
                End If
            End If
        End If
    End With
End Sub

Private Sub vsQuery_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If vsQuery.TextMatrix(0, Col) = "ҩƷ��Ϣ" Then
        Call vsQuery.AutoSize(Col)
        Call SetMinRowHeight
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsQuery.TextMatrix(vsQuery.FixedRows - 1, Col) & "A")
        If vsQuery.ColWidth(Col) < lngW Then
            vsQuery.ColWidth(Col) = lngW
        ElseIf vsQuery.ColWidth(Col) > vsQuery.Width * 0.5 Then
            vsQuery.ColWidth(Col) = vsQuery.Width * 0.5
        End If
    End If
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    
    '��ͷ
    objOut.Title.Text = tbcQuery.Selected.Caption
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "������" & sys.RowValue("���ű�", mvQuery.����ID, "����") & "/ҩ����" & sys.RowValue("���ű�", mvQuery.ҩ��ID, "����")
    objRow.Add Format(mvQuery.DateBegin, "yyyy-MM-dd HH:mm") & "/" & Format(mvQuery.DateEnd, "yyyy-MM-dd HH:mm")
    objOut.UnderAppRows.Add objRow
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = vsQuery
    
    '���
    vsQuery.Redraw = False
    lngRow = vsQuery.Row: lngCol = vsQuery.Col
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    vsQuery.Row = lngRow: vsQuery.Col = lngCol
    vsQuery.Redraw = True
End Sub

Private Sub InitColumnSelect()
'���ܣ����ݷ�ҩ��ϸ�嵥ԭʼ����ʾ״̬��ʼ����ѡ����
    Dim lngRow As Long, i As Long
    vsColumn.Rows = vsColumn.FixedRows
    With vsQuery
        For i = .FixedCols To .Cols - 1
            If .TextMatrix(0, i) <> "" Then
                vsColumn.Rows = vsColumn.Rows + 1
                lngRow = vsColumn.Rows - 1
                vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                vsColumn.RowData(lngRow) = i
                If vsQuery.ColHidden(i) Then
                    vsColumn.TextMatrix(lngRow, 0) = 0
                End If
                '�̶���ʾ��
                If InStr(",ҩƷ��Ϣ,����,,����,Ƶ��,�÷�,", "," & .TextMatrix(0, i) & ",") > 0 Then
                    vsColumn.TextMatrix(lngRow, 0) = 1
                    vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                End If
            End If
        Next
    End With
    If vsColumn.Rows > 1 Then vsColumn.Row = 1
End Sub




