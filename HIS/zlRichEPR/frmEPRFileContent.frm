VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "ZLRICHEDITOR.OCX"
Begin VB.Form frmEPRFileContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "�����ļ����"
   ClientHeight    =   10440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
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
   ScaleHeight     =   10440
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picWave 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   600
      ScaleHeight     =   3255
      ScaleWidth      =   6150
      TabIndex        =   7
      Top             =   3210
      Visible         =   0   'False
      Width           =   6150
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   1875
         Left            =   90
         TabIndex        =   11
         Top             =   1230
         Width           =   3030
         _cx             =   5345
         _cy             =   3307
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
         BackColorFixed  =   -2147483634
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEPRFileContent.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
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
         OwnerDraw       =   1
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
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   1935
         TabIndex        =   8
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1179648
      End
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   1050
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1179649
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   0
         ScaleHeight     =   900
         ScaleWidth      =   1575
         TabIndex        =   10
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   5430
      ScaleHeight     =   2985
      ScaleWidth      =   3420
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   3420
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   4170
         Left            =   -210
         TabIndex        =   1
         Top             =   780
         Width           =   7500
         _cx             =   13229
         _cy             =   7355
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
         BackColorFixed  =   15790320
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
         Rows            =   4
         Cols            =   6
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
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
      Begin VB.Label lblSubEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע:##"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ�㻤���¼��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2970
         TabIndex        =   3
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:##"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   630
      End
   End
   Begin VB.PictureBox picRich 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   1785
      ScaleHeight     =   3150
      ScaleWidth      =   4830
      TabIndex        =   4
      Top             =   60
      Width           =   4830
      Begin zlRichEditor.Editor edtThis 
         Height          =   2580
         Left            =   150
         TabIndex        =   5
         Top             =   75
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4551
         WithViewButtonas=   0   'False
         ShowRuler       =   0   'False
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRFileContent.frx":0062
      Left            =   120
      Top             =   645
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmEPRFileContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Enum zlEnumCompendParentKind     '��ٸ�����
    cprEmCPKFileDefine = 0              '�ļ���������
    cprEmCPKModelEssay = 1              '��������
End Enum

Private Enum FileType
    conPane_RichEpr = 1
    conPane_TendEpr = 2
    conPane_TablEpr = 3
    conPane_Infection = 4
    conPane_WaveEpr = 5 'ר�����µ�ҳ��
End Enum

Private msinVStep As Single      '�������Ĳ���
Private msinHStep As Single      '�������Ĳ���
'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event DblClick()                                                 '����˫�������¼�
Private mObjTabEprView As cTableEPR
Private mobjInfection As Object

'-----------------------------------------------------
'��ʱ����

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Me.edtThis.Copy
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_RichEpr
        Item.Handle = picRich.hWnd
    Case conPane_TendEpr
        Item.Handle = picTab.hWnd
    Case conPane_TablEpr
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    Case conPane_Infection
        Item.Handle = mobjInfection.zlGetForm.hWnd
    Case conPane_WaveEpr
        Item.Handle = picWave.hWnd
    End Select
End Sub

Private Sub Form_Load()
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    
    Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
        
    Set Pane1 = dkpMan.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_TendEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
    
    Set pane3 = dkpMan.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    Set Pane4 = dkpMan.CreatePane(conPane_Infection, 1200, 200, DockTopOf, Nothing)
    Pane4.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane4.Close
    
    Set Pane5 = dkpMan.CreatePane(conPane_WaveEpr, 1200, 200, DockTopOf, Nothing)
    Pane5.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane5.Close
    
    With dkpMan
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
    
    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "��Ⱦ�����濨", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
    Unload mobjInfection.zlGetForm
    Set mobjInfection.zlGetForm = Nothing
    Set mobjInfection = Nothing
End Sub
Private Sub picRich_Resize()
    edtThis.Top = 0: edtThis.Left = 0
    edtThis.Width = picRich.ScaleWidth: edtThis.Height = picRich.ScaleHeight
End Sub

Private Sub picTab_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    Me.lblTitle.Move Me.picTab.ScaleLeft, Me.picTab.ScaleTop + 120, Me.picTab.ScaleWidth
    Me.lblSubhead.Move Me.picTab.ScaleLeft + 210, Me.lblTitle.Top + Me.lblTitle.Height + 120
    Me.vfgThis.Move Me.picTab.ScaleLeft + 210, Me.lblSubhead.Top + Me.lblSubhead.Height + 45, Me.picTab.ScaleWidth - 210 * 2
    Me.vfgThis.Height = Me.picTab.ScaleHeight - Me.vfgThis.Top - 210 - lblSubEnd.Height - 45
    Me.lblSubEnd.Move lblSubhead.Left, Me.vfgThis.Top + Me.vfgThis.Height + 45
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        Popup.ShowPopup
    End With
End Sub

'-----------------------------------------------------
'���幫������
'-----------------------------------------------------

Public Sub zlRefresh(ByVal lngParentId As Long, Optional bytParentKind As zlEnumCompendParentKind = cprEmCPKFileDefine)
    '���ܣ���ʾָ���ļ�/���ĵ����ݣ�
    Dim strTemp As String, strZipFile As String
    Dim rsTemp As New ADODB.Recordset
    Dim mEPRFileInfo As cEPRFileDefineInfo
    Dim lngCount As Long
    Dim blnCollegeWave As Boolean '�Ƿ���ר�����µ�
    Dim lngTop As Long
    
    dkpMan.FindPane(conPane_TendEpr).Close
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_WaveEpr).Close
    dkpMan.FindPane(conPane_Infection).Close
    dkpMan.ShowPane conPane_RichEpr
    Me.edtThis.ReadOnly = False
    Me.edtThis.NewDoc
    If lngParentId = 0 Then Me.edtThis.ReadOnly = True: Exit Sub
    Me.edtThis.Freeze
        
    If bytParentKind = cprEmCPKFileDefine Then '�����ļ�����
        Err = 0: On Error GoTo errHand
        gstrSQL = "Select b.����, b.����,B.���,B.����, a.��ʽ" & vbNewLine & _
                "From ����ҳ���ʽ a, �����ļ��б� b" & vbNewLine & _
                "Where a.���� = b.���� And a.��� = b.ҳ�� And b.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
    
        '����ҳ���ʽ
        Set mEPRFileInfo = New cEPRFileDefineInfo
        mEPRFileInfo.��ʽ = "" & rsTemp!��ʽ
        mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.��ʽ
        Set mEPRFileInfo = Nothing
        Me.edtThis.ResetWYSIWYG
        
        If Val("" & rsTemp!����) < 0 And rsTemp!���� <> 6 Then
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            
            If rsTemp!���� = 3 Then '���µ�
                dkpMan.FindPane(conPane_TendEpr).Close
                dkpMan.FindPane(conPane_RichEpr).Close
                dkpMan.FindPane(conPane_TablEpr).Close
                dkpMan.FindPane(conPane_Infection).Close
                dkpMan.ShowPane conPane_WaveEpr
                Me.picWave.Visible = True
                VsfData.Visible = False
                msinVStep = 0: msinHStep = 0
                If NVL(rsTemp!����) = "1" Then 'ר�����µ�
                    gstrSQL = _
                        " SELECT Id, �ļ�id, ��id, �������, ��������, ������, ��������, �����д�, �����ı�, �Ƿ���, Ҫ������, Ҫ�ر�ʾ" & vbNewLine & _
                        " FROM �����ļ��ṹ" & vbNewLine & _
                        " WHERE �ļ�id = [1]" & vbNewLine & _
                        " START WITH ��id IS NULL" & vbNewLine & _
                        " CONNECT BY PRIOR Id = ��id"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
                    blnCollegeWave = True
                Else '��׼���µ�չʾ"ʾ����ʽ"
                    blnCollegeWave = False
                    Set rsTemp = GetPipWaveStyle(lngParentId)
                End If
                If rsTemp.RecordCount > 0 Then
                    picDraw.AutoRedraw = True
                    Call DrawWaveStyle(picDraw, rsTemp, Not blnCollegeWave, lngTop)
                    rsTemp.Filter = "Ҫ������ ='Ӥ�����µ�'"
                    If rsTemp.RecordCount > 0 Then
                        If Val(rsTemp!�����ı�) = 1 Then
                            rsTemp.Filter = ""
                            Call ShowTabBaby(rsTemp, lngTop)
                        End If
                    End If
                    Call CalcScrollBarSize
                End If
            Else
                dkpMan.FindPane(conPane_TendEpr).Close
                dkpMan.FindPane(conPane_WaveEpr).Close
                dkpMan.FindPane(conPane_TablEpr).Close
                dkpMan.FindPane(conPane_Infection).Close
                dkpMan.ShowPane conPane_RichEpr
                With Me.edtThis
                    .Text = vbCrLf & Space(4) & "���ļ�Ϊ�����ʽ���������������ʽ..."
                    .SelectAll
                    .ForceEdit = True
                    .Selection.Font.Name = "����": .Selection.Font.Size = 10.5
                    .SelLength = 0
                    .ForceEdit = False
                End With
            End If
        ElseIf NVL(rsTemp!����, 0) = 2 Then
            With Me.edtThis
                .Text = vbCrLf & Space(4) & "���ļ�Ϊ���ʽ���������ڶ�ȡ�ļ���ʽ..."
                .SelectAll
                .ForceEdit = True
                .Selection.Font.Name = "����": .Selection.Font.Size = 10.5
                .SelLength = 0
                .ForceEdit = False
            End With
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_RichEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_TablEpr
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_�����ļ�����, lngParentId, False, 0)
            Call mObjTabEprView.zlRefreshDockfrm 'ˢ����ʾ
        ElseIf NVL(rsTemp!����, 0) = 4 Then
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_RichEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            dkpMan.ShowPane conPane_Infection
            Call mobjInfection.zlRefresh(0, 0, 0, False)
        ElseIf rsTemp!���� = 3 Then
            dkpMan.FindPane(conPane_RichEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_TendEpr
            
            Dim lngCurColor As Long, strCurFont As String, objFont As StdFont
            Me.lblTitle.Caption = "": Me.lblSubhead.Caption = "": Me.lblSubEnd.Caption = ""
            Me.vfgThis.Redraw = flexRDNone
            Me.vfgThis.Clear: Me.vfgThis.MergeCells = flexMergeFixedOnly: vfgThis.MergeCellsFixed = flexMergeRestrictAll
            Me.vfgThis.MergeRow(0) = True
            Me.vfgThis.MergeRow(1) = True
            Me.vfgThis.MergeRow(2) = True
            Me.picTab.Visible = True
'
            gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
                " Order By d.�������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Select Case "" & !Ҫ������
                    Case "��ͷ����"
                        If Val("" & !�����ı�) = 1 Then
                            Me.vfgThis.RowHidden(0) = False
                            Me.vfgThis.RowHidden(1) = True
                            Me.vfgThis.RowHidden(2) = True
                        ElseIf Val("" & !�����ı�) = 2 Then
                            Me.vfgThis.RowHidden(0) = False
                            Me.vfgThis.RowHidden(1) = False
                            Me.vfgThis.RowHidden(2) = True
                        Else
                            Me.vfgThis.RowHidden(0) = False
                            Me.vfgThis.RowHidden(1) = False
                            Me.vfgThis.RowHidden(2) = False
                        End If
                    Case "������":  Me.vfgThis.Cols = Val("" & !�����ı�)
                    Case "��С�и�": Me.vfgThis.RowHeightMin = Val("" & !�����ı�)
                    Case "�ı�����"
                        strCurFont = "" & !�����ı�
                        Set objFont = New StdFont
                        With objFont
                            .Name = Split(strCurFont, ",")(0)
                            .Size = Val(Split(strCurFont, ",")(1))
                            .Bold = False: .Italic = False
                            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                        End With
                        Set Me.vfgThis.Font = objFont
                        Set Me.lblSubhead.Font = Me.vfgThis.Font
                        Set Me.lblSubEnd.Font = Me.vfgThis.Font
                        
                    Case "�ı���ɫ": Me.vfgThis.ForeColor = Val("" & !�����ı�)
                    Case "�����ɫ": Me.vfgThis.GridColor = Val("" & !�����ı�): Me.vfgThis.GridColorFixed = Me.vfgThis.GridColor
                    
                    Case "�����ı�": Me.lblTitle.Caption = "" & !�����ı�
                    Case "��������"
                        strCurFont = "" & !�����ı�
                        Set objFont = New StdFont
                        With objFont
                            .Name = Split(strCurFont, ",")(0)
                            .Size = Val(Split(strCurFont, ",")(1))
                            .Bold = False: .Italic = False
                            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                        End With
                        Set Me.lblTitle.Font = objFont
                        Me.lblTitle.AutoSize = False
                    End Select
                    .MoveNext
                Loop
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
                " Order By d.�������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Me.lblSubhead.Caption = ""
                Do While Not .EOF
                    Me.lblSubhead.Caption = Me.lblSubhead.Caption & " " & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
                    .MoveNext
                Loop
                Me.lblSubhead.Caption = Trim(Me.lblSubhead.Caption)
            End With
            
            '---------------------------------------------------
            gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���±�ǩ'" & _
                " Order By d.�������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Me.lblSubEnd.Caption = ""
                Do While Not .EOF
                    Me.lblSubEnd.Caption = Me.lblSubEnd.Caption & " " & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
                    .MoveNext
                Loop
                Me.lblSubEnd.Caption = Trim(Me.lblSubEnd.Caption)
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.�������, d.�����д�, d.�����ı�" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
                " Order By d.�������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    vfgThis.TextMatrix(!�����д� - 1, !������� - 1) = "" & !�����ı�
                    vfgThis.FixedAlignment(!������� - 1) = flexAlignCenterCenter
                    .MoveNext
                Loop
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
                " Order By d.�������, d.�����д�"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Me.vfgThis.ColWidth(!������� - 1) = Val("" & !��������)
                    .MoveNext
                Loop
            End With
            vfgThis.AutoSizeMode = flexAutoSizeRowHeight
            vfgThis.AutoSize 0, vfgThis.Cols - 1
            Me.vfgThis.Redraw = flexRDDirect
                    
            '---------------------------------------------------
            Call picTab_Resize
        Else
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_RichEpr
            strZipFile = zlBlobRead(1, lngParentId)
            If Len(strZipFile) > 0 Then
                If gobjFSO.FileExists(strZipFile) Then
                    strTemp = zlFileUnzip(strZipFile)
                    If gobjFSO.FileExists(strTemp) Then
                        Me.edtThis.OpenDoc strTemp
                        gobjFSO.DeleteFile strTemp, True
                    End If
                    gobjFSO.DeleteFile strZipFile, True
                End If
            End If
        End If
    Else '������������
        gstrSQL = "Select c.Id, c.����, a.��ʽ" & vbNewLine & _
            "From ����ҳ���ʽ a, �����ļ��б� b, ��������Ŀ¼ c" & vbNewLine & _
            "Where c.�ļ�id = b.Id And b.���� = a.���� And b.ҳ�� = a.��� And c.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
        If rsTemp.RecordCount <= 0 Then Exit Sub
        If NVL(rsTemp!����, 0) = 2 Then
            With Me.edtThis
                .Text = vbCrLf & Space(4) & "���ļ�Ϊ���ʽ�������ݲ�֧�������ʽ..."
                .SelectAll
                .ForceEdit = True
                .Selection.Font.Name = "����": .Selection.Font.Size = 10.5
                .SelLength = 0
                .ForceEdit = False
            End With
            dkpMan.FindPane(conPane_RichEpr).Close
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_TablEpr
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_ȫ��ʾ���༭, lngParentId, False, 0)
            Call mObjTabEprView.zlRefreshDockfrm 'ˢ����ʾ
        Else
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_RichEpr
            Set mEPRFileInfo = New cEPRFileDefineInfo
            mEPRFileInfo.��ʽ = "" & rsTemp!��ʽ
            mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.��ʽ
            Set mEPRFileInfo = Nothing
            Me.edtThis.ResetWYSIWYG
            
            If Val("" & rsTemp!����) = 0 Then
                strZipFile = zlBlobRead(3, lngParentId)
                If Len(strZipFile) > 0 Then
                    If gobjFSO.FileExists(strZipFile) Then
                        strTemp = zlFileUnzip(strZipFile)
                        If gobjFSO.FileExists(strTemp) Then
                            Me.edtThis.OpenDoc strTemp
                            gobjFSO.DeleteFile strTemp, True
                        End If
                        gobjFSO.DeleteFile strZipFile, True
                    End If
                End If
            Else
                Call InsertContent(lngParentId)
            End If
        End If
    End If
    
    '��ͷ��д
    vfgThis.MergeCells = flexMergeFixedOnly
    vfgThis.MergeCellsFixed = flexMergeFree
    For lngCount = 0 To vfgThis.Cols - 1
        vfgThis.MergeCol(lngCount) = True
    Next
    Me.vfgThis.AutoSize 0, Me.vfgThis.Cols - 1
    
    Me.edtThis.UnFreeze
    edtThis.RefreshTargetDC
    Me.edtThis.ReadOnly = True
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetCaption���֤()
    If Not mobjInfection Is Nothing Then Call mobjInfection.SetCaption���֤
End Sub

Private Sub InsertContent(ByVal lngFileID As Long)
    Dim rsTemp As New ADODB.Recordset
    Dim rsText As New ADODB.Recordset, strTSql As String
    Dim Elements As New cEPRElements
    Dim Diagnosises As New cEPRDiagnosises
    Dim aryProp() As String, intWCount As Integer
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lngKey As Long, lngStart As Long, lngLen As Long, strTmp As String
    
    With Me.edtThis
        .Freeze: .ForceEdit = True: .SelStart = 1
        intWCount = (.PaperWidth - .MarginLeft - .MarginRight) / Me.TextWidth("��") - 1
    End With
    
    gstrSQL = "Select Id, �����ı� From ������������ Where �ļ�id = [1] And �������� = 1 Order By �������"
    strTSql = "Select Id, ��������, ��������, �����ı�, �Ƿ���, Ҫ������, ����Ҫ��id, �滻��, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ," & vbNewLine & _
            "       Ҫ�ر�ʾ, Ҫ��ֵ��, ������̬" & vbNewLine & _
            "From ������������" & vbNewLine & _
            "Where �ļ�id = [1] And ��id + 0 = [2]" & vbNewLine & _
            "Order By �������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    Do While Not rsTemp.EOF
        lngStart = Me.edtThis.SelStart
        strTmp = StrConv("" & Trim(rsTemp!�����ı�), vbWide)
        strTmp = vbCrLf & "<" & strTmp & ">" & String(intWCount - Len(strTmp) - 1, "��") & vbCrLf
        lngLen = Len(strTmp)
        Me.edtThis.Range(lngStart, lngStart) = strTmp
        Me.edtThis.Range(lngStart, lngStart + lngLen).Font.Protected = False
        Me.edtThis.Range(lngStart, lngStart + lngLen).Font.Hidden = False
        Me.edtThis.Range(lngStart, lngStart + lngLen).Font.ForeColor = &HFFC0C0
        Me.edtThis.Range(lngStart + lngLen, lngStart + lngLen).Selected
        
        Set rsText = zlDatabase.OpenSQLRecord(strTSql, Me.Caption, lngFileID, CLng(rsTemp!ID))
        Do While Not rsText.EOF
            lngStart = Me.edtThis.SelStart
            Select Case rsText!��������
            Case 2, 3, 5 '�ı�,���,ͼ��
                Select Case rsText!��������
                Case 2: strTmp = "" & rsText!�����ı� & IIf(Val("" & rsText!�Ƿ���) = 1, vbCrLf, "")
                Case 3: strTmp = vbCrLf & "��" & vbCrLf
                Case 5: strTmp = vbCrLf & "��" & vbCrLf
                End Select
                lngLen = Len(strTmp)
                Me.edtThis.Range(lngStart, lngStart) = strTmp
                Me.edtThis.Range(lngStart, lngStart + lngLen).Font.Protected = False
                Me.edtThis.Range(lngStart, lngStart + lngLen).Font.Hidden = False
                Me.edtThis.Range(lngStart + lngLen, lngStart + lngLen).Selected
            Case 4  'Ҫ��
                lngKey = Elements.Add
                With Elements("K" & lngKey)
                    .�����ı� = "" & rsText!�����ı�
                    .Ҫ������ = "" & rsText!Ҫ������
                    .����Ҫ��ID = Val("" & rsText!����Ҫ��ID)
                    .�滻�� = Val("" & rsText!�滻��)
                    .Ҫ������ = Val("" & rsText!Ҫ������)
                    .Ҫ�س��� = Val("" & rsText!Ҫ�س���)
                    .Ҫ��С�� = Val("" & rsText!Ҫ��С��)
                    .Ҫ�ص�λ = "" & rsText!Ҫ�ص�λ
                    .Ҫ�ر�ʾ = Val("" & rsText!Ҫ�ر�ʾ)
                    .Ҫ��ֵ�� = "" & rsText!Ҫ��ֵ��
                    .������̬ = Val("" & rsText!������̬)
                    .�Ƿ��� = Val("" & rsText!�Ƿ���)
                    .InsertIntoEditor Me.edtThis, lngStart, , True
                End With
            Case 7  '���
                lngKey = Diagnosises.Add
                With Diagnosises("K" & lngKey)
                    .���� = "" & rsText!�����ı�
                    aryProp = Split("" & rsText!��������, ";")
                    .���� = Val(aryProp(0))
                    .��ҽ = Val(aryProp(1))
                    .����id = Val(aryProp(2))
                    .���id = Val(aryProp(3))
                    .֤��id = Val(aryProp(4))
                    .���� = Val(aryProp(5))
                    .���� = Format(aryProp(6), "yyyy-mm-dd hh:mm:ss")
                    .InsertIntoEditor Me.edtThis, lngStart, True
                End With
            End Select
            rsText.MoveNext
        Loop
        rsTemp.MoveNext
    Loop
    With Me.edtThis
        .ForceEdit = False: .SelStart = 1: .Modified = False: .UnFreeze
    End With
    Set Elements = Nothing
    Set Diagnosises = Nothing
End Sub


Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ� ���óɹ�����TRUE������FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    vsb.Value = 0: hsb.Value = 0
    picDraw.Top = 0: picDraw.Left = 0
    hsb.Max = picDraw.Width - picWave.Width
    vsb.Max = picDraw.Height - picWave.Height
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    If hsb.Visible Then hsb.ZOrder 0
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    If vsb.Visible Then vsb.ZOrder 0
    
    With vsb
        .Height = picWave.Height
    End With
    
    With hsb
        .Width = picWave.Width - IIf(vsb.Visible = True, vsb.Width, 0)
    End With
    
    'ֻ����û��ʾ�������ǲ��������㲽��
    msinHStep = (picDraw.Width - picWave.Width + IIf(vsb.Visible = True, vsb.Width, 0)) / 10
    msinVStep = (picDraw.Height - picWave.Height + IIf(hsb.Visible = True, hsb.Height, 0)) / 10
    
    '�㶨Ϊ100,ֻ�ǲ��������仯
    If hsb.Enabled Then
        hsb.Max = 10
        hsb.LargeChange = 10 / Int((Round((picDraw.Width - picWave.Width + IIf(vsb.Visible = True, vsb.Width, 0)) / picWave.Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange / 2
    End If
    
    If vsb.Enabled Then
        vsb.Max = 10
        vsb.LargeChange = 10 / Int((Round((picDraw.Height - picWave.Height + IIf(hsb.Visible = True, hsb.Height, 0)) / picWave.Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange / 2
    End If
    
    CalcScrollBarSize = True
End Function

Private Sub picWave_Resize()
    With vsb
        .Left = picWave.Width - .Width
        .Top = 0
        .Height = picWave.Height
    End With
    
    With hsb
        .Left = 0
        .Top = picWave.Height - .Height
        .Width = picWave.Width - vsb.Width
    End With
    
    Call CalcScrollBarSize
End Sub

Private Sub vsb_Change()
    picDraw.Top = -1 * vsb.Value * msinVStep
    VsfData.Top = (picDraw.Height - VsfData.Height) + -1 * vsb.Value * msinVStep
End Sub

Private Sub hsb_Change()
    picDraw.Left = -1 * hsb.Value * msinHStep
    VsfData.Left = -1 * hsb.Value * msinHStep
End Sub

Private Function ShowTabBaby(ByVal rsTmp As ADODB.Recordset, ByVal lngHeight As Long)
    Dim lngCurveRows As Long
    Dim lngMaxValue As Long, lngMinValue As Long
    Dim lngTotal As Long, lngCurveNull As Long
    Dim lngCurveRowHeight As Long
    Dim lngTabBabyRowHeight As Long
    Dim lngRow As Long, lngDay As Long
    Dim lngId  As Long, lngTabBabyTitleID As Long, lngTabBabyNameID As Long
    Dim strSQL  As String
    Dim strBabyTitle As String, strTitleBabyFont As String
    Dim intTitleBabyTitleNum As Integer, i As Integer
    Dim BlnBaby As Boolean
    Dim objFont As StdFont
    
    Dim rsCurve As New ADODB.Recordset
    
    rsTmp.Filter = "��ID=NULL And �������=1 And �����ı�='��ʽ����'"
    If rsTmp.RecordCount > 0 Then
        lngId = rsTmp!ID
        rsTmp.Filter = "��ID=" & lngId
        Do While Not rsTmp.EOF
            Select Case "" & rsTmp!Ҫ������
            Case "����"
                lngDay = Val("" & rsTmp!�����ı�)
            Case "Ӥ�������ı�"
                strBabyTitle = "" & rsTmp!�����ı�
            Case "Ӥ����������"
                strTitleBabyFont = "" & rsTmp!�����ı�
            Case "Ӥ�����߶�"
                lngTabBabyRowHeight = Val("" & rsTmp!�����ı�)
            Case "��ͷ����"
                intTitleBabyTitleNum = Val("" & rsTmp!�����ı�)
            Case "Ӥ�����µ�"
                BlnBaby = Val("" & rsTmp!�����ı�)
            Case "������"
                VsfData.Cols = Val("" & rsTmp!�����ı�)
            End Select
            rsTmp.MoveNext
        Loop
    End If
    If Not BlnBaby Then VsfData.Visible = False: Exit Function
    
    rsTmp.Filter = "��ID=NULL And �������=4 And �����ı�='Ӥ�����µ���ͷ��Ŀ'"
    Do While Not rsTmp.EOF
        lngTabBabyTitleID = Val("" & rsTmp!ID)
        rsTmp.MoveNext
    Loop
    rsTmp.Filter = "��ID=NULL And �������=3 And �����ı�='�����Ŀ����'"
    Do While Not rsTmp.EOF
        lngTabBabyNameID = Val("" & rsTmp!ID)
        rsTmp.MoveNext
    Loop
    
    
    With VsfData
        .Top = Me.ScaleX(lngHeight, vbPixels, vbTwips) + 200
        .Left = 0
        .Rows = .FixedRows + lngDay + 1
        .Width = picDraw.Width
        .Height = lngTabBabyRowHeight * (VsfData.Rows + 2)
        
        Select Case intTitleBabyTitleNum
            Case 1
                .RowHidden(2) = True
                .RowHidden(3) = True
            Case 2
                .RowHidden(3) = True
        End Select
        
        rsTmp.Filter = "��ID= " & lngTabBabyNameID
        rsTmp.Sort = "�������"
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .ColWidth(Val(rsTmp!�������) - 1) = Split(rsTmp!��������, "`")(0)
                rsTmp.MoveNext
            Loop
        End If
        rsTmp.Filter = "��ID= " & lngTabBabyTitleID
        rsTmp.Sort = "�������"
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .TextMatrix((Val(rsTmp!�����д�)), Val(rsTmp!�������) - 1) = NVL(rsTmp!�����ı�)
                rsTmp.MoveNext
            Loop
        End If
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = strBabyTitle
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        .CellBorderRange 1, 0, .Rows - 1, .Cols - 1, vbBlack, 1, 1, 1, 1, 1, 1
        .MergeCellsFixed = flexMergeFree
        .MergeCol(-1) = True
        .MergeRow(-1) = True
        
        Set objFont = New StdFont
        With objFont
            .Name = Split(strTitleBabyFont, ",")(0)
            .Size = Val(Split(strTitleBabyFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strTitleBabyFont, "��") > 0 Then .Bold = True
            If InStr(1, strTitleBabyFont, "б") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 0, .FixedCols, 0, .Cols - 1) = objFont
        .ROWHEIGHT(0) = objFont.Size * 20 + 150
        For i = 4 To .Rows - 1
        .ROWHEIGHT(i) = lngTabBabyRowHeight
        VsfData.Redraw = True
        Next
        
    End With
    picDraw.Height = picDraw.Height + VsfData.Height
    VsfData.Visible = True
    
End Function


