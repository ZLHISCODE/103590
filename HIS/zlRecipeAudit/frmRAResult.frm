VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmRAResult 
   Caption         =   "���������"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmRAResult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picNG 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   480
      ScaleHeight     =   1215
      ScaleWidth      =   1935
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1935
      Begin VSFlex8Ctl.VSFlexGrid vsfNG 
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
         _cx             =   2566
         _cy             =   1296
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
   Begin VB.PictureBox picAuditInfo 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   600
      ScaleHeight     =   1935
      ScaleWidth      =   5625
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   5625
      Begin VB.TextBox txtReason 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label lblAuditInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ۺ����ɣ�"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblAuditInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "�ϸ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Top             =   180
         Width           =   1770
      End
      Begin VB.Label lblAuditInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ�䣺"
         Height          =   180
         Index           =   2
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblAuditInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ˣ�"
         Height          =   180
         Index           =   1
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblAuditInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   900
      End
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   1440
      Top             =   480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmRAResult.frx":6DC2
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   960
      Top             =   480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmRAResult.frx":6DDC
      Left            =   600
      Top             =   480
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmRAResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_VSF As String = "����,,3,1000|���,,3,3000|��Ŀ����,,3,4000|ҩƷ,,3,3000"

Private Enum enuToolsID
    ���� = 10001
    �Ա� = 10002
    ���� = 10003
End Enum

Private mlngMedicalID As Long
Private mblnMemory As Boolean
Private mintResult As Integer
Private mintStatus As Integer
Private mblnLocking As Boolean

Public Sub ShowMe(ByVal lngMedicalID As Long, ByVal frmOwner As Form)
'���ܣ���ʾ����ӿ�
'������
'  lngMedicalID����ҩ;��ҽ��ID
'  frmOwner�������������

    mlngMedicalID = lngMedicalID
    mblnMemory = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1

    Show vbModal, frmOwner

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Call InitCommandbars
    Call InitDockPane
    Call InitVSF
    mdlDefine.SetVSFHead vsfNG, MSTR_VSF
    
    If mblnMemory Then
        Dim strPane As String
        RestoreWinState Me, App.ProductName
        strPane = GetSetting("ZLSOFT", zlStr.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "����")
        dkpMain.LoadStateFromString strPane
    End If
    
    Call SetPatientInfo
    Call SetAuditInfo
    Call SetNGInfo
    
    '�Զ��и�
    vsfNG.AutoSize 0, vsfNG.Cols - 1
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picAuditInfo.hwnd
        Case 2
            Item.Handle = picNG.hwnd
    End Select
End Sub

Private Sub InitDockPane()
    Dim panTop As Pane, panClient As Pane
    
    With dkpMain
        .SetCommandBars cbsMain
        .Options.UseSplitterTracker = False
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .Options.LunaColors = True
        .Options.HideClient = True
        .VisualTheme = ThemeOffice2003
        
        Set panTop = .CreatePane(1, 0, 200, DockTopOf)
        With panTop
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .Title = "�����Ϣ"
            .MaxTrackSize.Height = 250
            .MinTrackSize.Height = 150
        End With
        
        Set panClient = .CreatePane(2, 0, 400, DockBottomOf)
        With panClient
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .Title = "���ϸ������Ŀ"
        End With
    End With
End Sub

Private Sub InitCommandbars()
    Dim cbpTmp As CommandBarPopup
    Dim cbcTmp As CommandBarControl
    Dim cbrTmp As CommandBar
    Dim strName As String, strSex As String, strAge As String
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = False
        .UseDisabledIcons = True
        .LargeIcons = False
        '.SetIconSize True, 24, 24
        '.SetIconSize False, 16, 16
    End With
    With cbsMain
        .ActiveMenuBar.Visible = False
        .EnableCustomization False
    End With
    
    '���幤����
    Set cbrTmp = cbsMain.Add("������", xtpBarTop)
    With cbrTmp
        .ShowTextBelowIcons = True
        .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
        
        Set cbcTmp = .Controls.Add(xtpControlLabel, enuToolsID.����, zlStr.FormatString("���ˣ�[1]  ", strName))
        Set cbcTmp = .Controls.Add(xtpControlLabel, enuToolsID.�Ա�, zlStr.FormatString("�Ա�[1]  ", strSex))
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlLabel, enuToolsID.����, zlStr.FormatString("���䣺[1]  ", strAge))
        cbcTmp.BeginGroup = True
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Height < 5250 Then Height = 5250
    If Width < 7500 Then Width = 7500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strPane As String
    
    SaveWinState Me, App.ProductName
    strPane = dkpMain.SaveStateToString
    SaveSetting "ZLSOFT", zlStr.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "����", strPane
End Sub

Private Sub picAuditInfo_Resize()
    On Error Resume Next
    
    With lblAuditInfo(0)
        .Left = 120
        .Top = 180
    End With
    With lblAuditInfo(3)    '�ϸ�/���ϸ�
        .Left = lblAuditInfo(0).Left + lblAuditInfo(0).Width
        .Top = lblAuditInfo(0).Top - 90
    End With
    With lblAuditInfo(1)
        .Left = (picAuditInfo.Width - 2500) \ 2
        .Top = lblAuditInfo(0).Top
    End With
    With lblAuditInfo(2)
        .Left = picAuditInfo.Width - 120 * 2 - 2500
        .Top = lblAuditInfo(0).Top
    End With
    With lblAuditInfo(4)
        .Left = 120
        .Top = lblAuditInfo(0).Top + lblAuditInfo(0).Height + 120
    End With
    With txtReason
        .Left = 120
        .Top = lblAuditInfo(4).Top + lblAuditInfo(4).Height + 120
        .Width = picAuditInfo.Width - 120 * 2
        .Height = picAuditInfo.Height - .Top - 120
    End With
End Sub

Private Sub picNG_Resize()
    On Error Resume Next
    
    vsfNG.Move 0, 0, picNG.Width, picNG.Height
End Sub

Private Sub SetNGInfo()
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    'If mintResult = 0 and  Then
    
    '��������
    On Error GoTo errHandle
    strSQL = "Select d.Id, d.����, d.���, d.���� ��Ŀ����, f_List2str(Cast(Collect(b.����) As t_Strlist), '��'||Chr(13)) ҩƷ " & vbNewLine & _
             "From ����ҽ����¼ A, ������ĿĿ¼ B, ��������� C, ���������Ŀ D " & vbNewLine & _
             "Where a.Id(+) = c.ҽ��id And a.������Ŀid = b.Id(+) And c.�����Ŀid = d.Id And c.��id = [1] And c.ҩʦ��� = 2 " & vbNewLine & _
             "Group By d.Id, d.����, d.���, d.���� " & vbNewLine & _
             "Order By d.���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�󷽵Ĳ��ϸ���Ϣ", Val(picAuditInfo.Tag))
    mdlDefine.FillVSFData vsfNG, rsTemp
    rsTemp.Close
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetAuditInfo()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    mblnLocking = False
    mintStatus = 0
    mintResult = 0
    
    On Error GoTo errHandle
    
    strSQL = "Select b.Id, b.�����, b.�����, to_char(b.���ʱ��, 'yyyy-mm-dd hh24:mi:ss') ���ʱ��, b.״̬, b.�����û�, b.�ۺ����� " & vbNewLine & _
             "From ���������ϸ A, ��������¼ B, ����ҽ����¼ C " & vbNewLine & _
             "Where a.��id = b.Id And a.ҽ��id = c.ID And c.���ID = [1] And a.����ύ = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ϣ", mlngMedicalID)
    If rsTemp.EOF = False Then
        '���¿ؼ�
        mblnLocking = NVL(rsTemp!�����û�) <> ""
        mintStatus = Val(NVL(rsTemp!״̬))
        mintResult = Val(NVL(rsTemp!�����))
        
        lblAuditInfo(1).Caption = zlStr.FormatString("����ˣ�[1]", NVL(rsTemp!�����))
        lblAuditInfo(2).Caption = zlStr.FormatString("���ʱ�䣺[1]", NVL(rsTemp!���ʱ��))
        If mintResult = 1 Then
            lblAuditInfo(3).Caption = "�ϸ�"
            lblAuditInfo(3).ForeColor = &H8000&
        ElseIf mintResult = 2 Then
            lblAuditInfo(3).Caption = "���ϸ�"
            lblAuditInfo(3).ForeColor = vbRed
        Else
            lblAuditInfo(3).ForeColor = vbBlue
            If mblnLocking Then
                lblAuditInfo(3).Caption = "���������"
            ElseIf mintStatus = 0 Then
                lblAuditInfo(3).Caption = "����"
            ElseIf mintStatus = 2 Or mintStatus = 3 Then
                lblAuditInfo(3).Caption = "����"
                lblAuditInfo(3).ForeColor = &H8000&
            End If
        End If
        txtReason.Text = NVL(rsTemp!�ۺ�����)
        picAuditInfo.Tag = NVL(rsTemp!ID)           '��ID
    Else
        lblAuditInfo(1).Caption = "����ˣ�"
        lblAuditInfo(2).Caption = "���ʱ�䣺"
        lblAuditInfo(3).Caption = ""
        txtReason.Text = ""
    End If
    rsTemp.Close
    
    Exit Sub

errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetPatientInfo()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strName As String, strSex As String, strAge As String
    
    On Error GoTo errHandle
    
    strSQL = "Select ����, �Ա�, ���� From ����ҽ����¼ Where ID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", mlngMedicalID)
    If rsTemp.EOF = False Then
         strName = zlStr.FormatString("���ˣ�[1]  ", rsTemp!����)
         strSex = zlStr.FormatString("�Ա�[1]  ", rsTemp!�Ա�)
         strAge = zlStr.FormatString("���䣺[1]  ", rsTemp!����)
    End If
    rsTemp.Close
    
    cbsMain.FindControl(xtpControlLabel, enuToolsID.����).Caption = strName
    cbsMain.FindControl(xtpControlLabel, enuToolsID.�Ա�).Caption = strSex
    cbsMain.FindControl(xtpControlLabel, enuToolsID.����).Caption = strAge
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txtReason_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub InitVSF()
'���ܣ���ʼ�������VSFlexGrid�ؼ��ķ��

    With vsfNG
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .BackColorBkg = .BackColor
        .AutoResize = True
    End With
End Sub
